#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把中文名的 docx 导成 PDF（Windows Word COM）。

用法：
    python3 export_pdf.py "D:\物理实验\3" "*第5周.docx"

为什么不用 soffice：本机没装 LibreOffice。
四个必踩的坑都在这里处理掉了：
  1. ps1 必须落在 Windows 能看见的盘（WSL 的 /home/... 不对应任何 C: 路径）→ 自动写到 <dir>/_work/
  2. ps1 必须带 UTF-8 BOM，否则 PowerShell 5.1 按 ANSI 读、中文路径乱码 → COMException「文档可能损坏」
  3. 中文文件名交给 Get-ChildItem 解析，不当参数传
  4. `~$xxx.docx` 是 Word 的锁文件、且常比正文新 → 必须排除，否则会挑中它而导不出
"""
import os
import subprocess
import sys

PWSH = '/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe'

# 注意：这是普通字符串不是 f-string，花括号要写单个。
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
    """PowerShell 5.1 的中文输出是控制台代码页（GBK），不是 UTF-8。
    text=True 会直接 UnicodeDecodeError: 'utf-8' codec can't decode byte 0xb2。"""
    if not blob:
        return ''
    for enc in ('utf-8', 'gb18030', 'cp936'):
        try:
            return blob.decode(enc).strip()
        except UnicodeDecodeError:
            continue
    return blob.decode('utf-8', errors='replace').strip()


def to_wsl(win_path):
    """C:\\a\\b -> /mnt/c/a/b"""
    drive, rest = win_path.split(':', 1)
    return '/mnt/' + drive.lower() + rest.replace('\\', '/')


def to_win(wsl_path):
    """/mnt/c/a/b -> C:\\a\\b"""
    if not wsl_path.startswith('/mnt/'):
        raise ValueError('not under /mnt: %s' % wsl_path)
    drive, rest = wsl_path[len('/mnt/'):].split('/', 1)
    return drive.upper() + ':\\' + rest.replace('/', '\\')


def export(win_dir, pattern='*.docx', timeout=600):
    wsl_dir = to_wsl(win_dir)
    if not os.path.isdir(wsl_dir):
        raise SystemExit('目录不存在: %s' % wsl_dir)
    work = os.path.join(wsl_dir, '_work')
    os.makedirs(work, exist_ok=True)
    ps1 = os.path.join(work, 'w2p.ps1')
    body = PS.replace('{dir}', win_dir).replace('{pat}', pattern)
    with open(ps1, 'wb') as f:
        f.write(b'\xef\xbb\xbf' + body.encode('utf-8'))   # 坑 2：BOM
    out = subprocess.run(
        [PWSH, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', to_win(ps1)],
        capture_output=True, timeout=timeout)   # 坑 5：不能 text=True，PowerShell 输出是 GBK
    stdout = decode(out.stdout)
    print(stdout)
    if decode(out.stderr):
        print(decode(out.stderr))
    if 'PDF_OK=' not in stdout:
        raise SystemExit(1)
    return stdout.split('PDF_OK=', 1)[1].strip()


def main():
    if len(sys.argv) < 2:
        raise SystemExit('用法: python3 export_pdf.py <Windows下的目录> [文件名通配符]')
    win_dir = sys.argv[1]
    pattern = sys.argv[2] if len(sys.argv) > 2 else '*.docx'
    export(win_dir, pattern)


if __name__ == '__main__':
    main()
