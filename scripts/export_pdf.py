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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hostenv

SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
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


# WPS 的 COM 名是 KWPS.Application；接口与 Word 兼容（ExportAsFixedFormat 17 = PDF）。
PS_WPS = r'''$ErrorActionPreference = "Continue"
$dir = "{dir}"
$pat = "{pat}"
$f = Get-ChildItem -LiteralPath $dir -Filter $pat |
     Where-Object { $_.Name -notlike '~$*' } |
     Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $f) { Write-Output "NOFILE"; exit 1 }
Write-Output ("SRC=" + $f.Name)
$dst = [System.IO.Path]::ChangeExtension($f.FullName, ".pdf")
if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }
$app = New-Object -ComObject KWPS.Application
$app.Visible = $false
try {
  $doc = $app.Documents.Open($f.FullName, $false, $true)
  $doc.ExportAsFixedFormat($dst, 17)
  $doc.Close(0)
  Write-Output ("PDF_OK=" + $dst)
} catch {
  Write-Output ("ERR=" + $_.Exception.Message)
} finally {
  $app.Quit()
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
    r"""Windows 路径 → 当前系统路径（WSL 下 C:\x -> /mnt/c/x；Windows 上不变）。"""
    return hostenv.to_local(path)


def to_win(path):
    r"""当前系统路径 → Windows 路径（WSL 下 /mnt/c/x -> C:\x；Windows 上不变）。"""
    return hostenv.to_win(path)


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
    """按顺序找第一个存在的可执行文件。带路径的判断存在，裸名字走 PATH。"""
    for c in cands:
        if os.path.isabs(c) or os.path.sep in c or (os.altsep and os.altsep in c):
            if os.path.exists(c):
                return c
        else:
            hit = shutil.which(c)
            if hit:
                return hit
    return None


def word_export(directory, pattern):
    pwsh = hostenv.powershell()
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


def _run_com_export(directory, pattern, template, label):
    """跑一段 COM 导出脚本（Word 或 WPS），返回 PDF 路径或 None。"""
    pwsh = hostenv.powershell()
    if not pwsh:
        return None
    work = os.path.join(to_local(directory), "_work")
    os.makedirs(work, exist_ok=True)
    ps_path = os.path.join(work, "_export_pdf_%s.ps1" % label)
    body = template.replace("{dir}", to_win(directory)).replace("{pat}", pattern)
    with open(ps_path, "wb") as f:
        f.write(b"\xef\xbb\xbf" + body.encode("utf-8"))
    out = subprocess.run([pwsh, "-NoProfile", "-ExecutionPolicy", "Bypass",
                          "-File", to_win(ps_path)], capture_output=True, timeout=300)
    stdout = decode(out.stdout)
    for line in stdout.splitlines():
        if line.strip() and not line.startswith("PDF_OK="):
            print("  [%s] %s" % (label, line))
    if "PDF_OK=" not in stdout:
        return None
    return stdout.split("PDF_OK=", 1)[1].strip().splitlines()[0]


def wps_export(directory, pattern):
    """只装了 WPS 的机器：KWPS.Application 导 PDF。"""
    return _run_com_export(directory, pattern, PS_WPS, "wps")


def soffice_export(directory, pattern):
    exe = first_existing(SOFFICE_CANDIDATES)
    if not exe:
        return None
    src = pick_file(directory, pattern)
    # 只有「在 WSL 里调用 Windows 的 soffice.exe」才需要转路径；
    # Windows 原生运行时路径本来就对，转换反而多余。
    need_conv = exe.lower().endswith(".exe") and hostenv.IS_WSL
    src_arg = to_win(src) if need_conv else src
    outdir = to_win(os.path.dirname(src)) if need_conv else os.path.dirname(src)
    cmd = [exe, "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", outdir, src_arg]
    out = subprocess.run(cmd, capture_output=True, timeout=300)
    txt = decode(out.stdout) + decode(out.stderr)
    if txt:
        print(txt.strip())
    pdf = os.path.splitext(src)[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None


def export(directory, pattern):
    """Word COM → WPS COM → LibreOffice，哪条能跑通用哪条。"""
    order = (word_export, wps_export, soffice_export)
    force = os.environ.get("PRC_EXPORT", "").lower()
    if force == "soffice":
        order = (soffice_export,)
    elif force == "wps":
        order = (wps_export,)
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
    print("ERR=三种导出方式都失败（Word COM / WPS / LibreOffice）")
    raise SystemExit(1)


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print("用法: python3 export_pdf.py <目录> [文件名通配符]")
        print("优先 Word COM，失败自动退 LibreOffice；成功打印 PDF_OK=<路径>。")
        print("环境变量 PRC_EXPORT=word|soffice 可强制指定。")
        raise SystemExit(0 if len(sys.argv) > 1 else 1)
    export(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "*.docx")


if __name__ == "__main__":
    main()
