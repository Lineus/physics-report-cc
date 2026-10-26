#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 docx 导成 PDF：优先 Word COM，退到 LibreOffice。

用法：
    python3 export_pdf.py "<目录>" "*第5周.docx"

Windows + Word 用 COM（版式与真 Word 一致）；没有 Word 或不在 Windows 时用
LibreOffice（soffice --headless）。成功打印 PDF_OK=<路径>，失败退出码 1。

四个必踩的坑（Word COM 路径）都处理掉了：
  1. ps1 必须落在 Windows 能看见的盘（WSL 的 /home/... 不对应任何 C: 路径）→ 写到 <dir>/_work/
  2. ps1 必须带 UTF-8 BOM，否则 PowerShell 5.1 按 ANSI 读、中文路径乱码 → COMException「文档可能损坏」
  3. 中文文件名交给 Get-ChildItem 解析，不当参数传
  4. `~$xxx.docx` 是 Word 的锁文件、且常比正文新 → 必须排除，否则会挑中它而导不出
"""
import glob
import os
import re
import shutil
import subprocess
import sys

PWSH_CANDIDATES = [
    "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe",
    "powershell.exe",
]
SOFFICE_CANDIDATES = [
    "/mnt/c/Program Files/LibreOffice/program/soffice.exe",
    "/mnt/c/Program Files (x86)/LibreOffice/program/soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "soffice",
    "libreoffice",
]

PS = r'''$ErrorActionPreference = "Continue"
$dir = "{dir}"
$pat = "{pat}"
$f = Get-ChildItem -LiteralPath $dir -Filter $pat |
     Where-Object { $_.Name -notlike '~$*' } |
     Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $f) { Write-Output "NOFILE"; exit 1 }
Write-Output ("SRC=" + $f.Name)
$dst = [System.IO.Path]::ChangeExtension($f.FullName, ".pdf")
if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($f.FullName, [ref]$false, [ref]$true)
  $doc.SaveAs([ref]$dst, [ref]17)
  $doc.Close([ref]0)
  Write-Output ("PDF_OK=" + $dst)
} catch {
  Write-Output ("ERR=" + $_.Exception.Message)
} finally {
  $word.Quit()
}
'''


def decode(blob):
    """PowerShell 5.1 的中文输出是控制台代码页（GBK），不是 UTF-8。"""
    if not blob:
        return ""
    for enc in ("utf-8", "gb18030", "cp936"):
        try:
            return blob.decode(enc).strip()
        except UnicodeDecodeError:
            continue
    return blob.decode("utf-8", errors="replace").strip()


def to_local(path):
    """Windows 路径 → 当前系统路径（WSL: C:\\x -> /mnt/c/x）。"""
    if os.path.exists(path):
        return path
    m = re.match(r"^([A-Za-z]):[\\/](.*)$", path)
    if m and os.path.isdir("/mnt"):
        cand = "/mnt/%s/%s" % (m.group(1).lower(), m.group(2).replace("\\", "/"))
        if os.path.exists(cand):
            return cand
    return path


def to_win(path):
    """当前系统路径 → Windows 路径（WSL: /mnt/c/x -> C:\\x）。已经是 Windows 形式就原样返回。"""
    if re.match(r"^[A-Za-z]:[\\/]", path):
        return path
    m = re.match(r"^/mnt/([a-z])/(.*)$", path)
    if m:
        return "%s:\\%s" % (m.group(1).upper(), m.group(2).replace("/", "\\"))
    return path


def pick_file(directory, pattern):
    """按通配符取最新一份 docx，排除 Word 锁文件。"""
    local = to_local(directory)
    hits = [p for p in glob.glob(os.path.join(local, pattern))
            if not os.path.basename(p).startswith("~$")]
    if not hits:
        print("NOFILE: %s/%s" % (directory, pattern))
        raise SystemExit(1)
    hits.sort(key=os.path.getmtime, reverse=True)
    return hits[0]


def first_existing(cands):
    for c in cands:
        if os.path.sep in c or c.endswith("powershell.exe"):
            if os.path.exists(c):
                return c
        else:
            hit = shutil.which(c)
            if hit:
                return hit
    return None


def word_export(directory, pattern):
    pwsh = first_existing(PWSH_CANDIDATES)
    if not pwsh:
        return None
    work = os.path.join(to_local(directory), "_work")
    os.makedirs(work, exist_ok=True)
    ps_path = os.path.join(work, "_export_pdf.ps1")
    with open(ps_path, "wb") as f:
        body = PS.replace("{dir}", to_win(directory)).replace("{pat}", pattern)
        f.write(b"\xef\xbb\xbf" + body.encode("utf-8"))
    out = subprocess.run([pwsh, "-NoProfile", "-ExecutionPolicy", "Bypass",
                          "-File", to_win(ps_path)], capture_output=True, timeout=300)
    stdout = decode(out.stdout)
    for line in stdout.splitlines():
        if line.strip() and not line.startswith("PDF_OK="):
            print(line)
    if "PDF_OK=" not in stdout:
        return None
    return stdout.split("PDF_OK=", 1)[1].strip().splitlines()[0]


def soffice_export(directory, pattern):
    exe = first_existing(SOFFICE_CANDIDATES)
    if not exe:
        return None
    src = pick_file(directory, pattern)
    win_exe = exe.lower().endswith(".exe")
    src_arg = to_win(src) if win_exe else src
    outdir = to_win(os.path.dirname(src)) if win_exe else os.path.dirname(src)
    cmd = [exe, "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", outdir, src_arg]
    out = subprocess.run(cmd, capture_output=True, timeout=300)
    txt = decode(out.stdout) + decode(out.stderr)
    if txt:
        print(txt.strip())
    pdf = os.path.splitext(src)[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None


def export(directory, pattern):
    order = (word_export, soffice_export)
    force = os.environ.get("PRC_EXPORT", "").lower()
    if force == "soffice":
        order = (soffice_export,)
    elif force == "word":
        order = (word_export,)
    for fn in order:
        try:
            path = fn(directory, pattern)
        except Exception as e:  # noqa: BLE001
            print("%s 失败：%s" % (fn.__name__, e))
            path = None
        if path:
            print("PDF_OK=" + path)
            return path
    print("ERR=两种导出方式都失败（Word COM / LibreOffice）")
    raise SystemExit(1)


def main():
    if len(sys.argv) < 2:
        raise SystemExit("用法: python3 export_pdf.py <目录> [文件名通配符]")
    export(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "*.docx")


if __name__ == "__main__":
    main()
