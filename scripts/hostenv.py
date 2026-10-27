# -*- coding: utf-8 -*-
"""宿主差异：Windows / WSL / macOS 上找 PowerShell、换路径。

这个技能最初在 WSL 里写成，早先版本把 `/mnt/c/...` 当成唯一入口，
导致 Windows 原生 Python 找不到 powershell.exe、导出 PDF 直接失败。
所有平台相关的判断都收在这里，别再散在脚本里。
"""
import glob
import os
import re
import shutil
import sys

def _release():
    try:
        with open("/proc/version") as f:
            return f.read().lower()
    except OSError:
        return ""


IS_WIN = os.name == "nt"
IS_WSL = (not IS_WIN) and os.path.isdir("/mnt/c") and "microsoft" in _release()


def powershell():
    """返回可用的 PowerShell 可执行文件；没有则 None。

    顺序：Windows 原生 -> PATH 里的 powershell/pwsh -> WSL 下的 /mnt/c 路径。
    Windows 上必须让 PATH 优先：`C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe`
    在 WSL 里不存在，反过来 `/mnt/c/...` 在 Windows 上也不存在。
    """
    if IS_WIN:
        for name in ("powershell.exe", "pwsh.exe", "powershell", "pwsh"):
            hit = shutil.which(name)
            if hit:
                return hit
        for p in (r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
                  r"C:\Program Files\PowerShell\7\pwsh.exe"):
            if os.path.exists(p):
                return p
        return None

    for name in ("pwsh", "powershell.exe", "powershell"):
        hit = shutil.which(name)
        if hit:
            return hit
    for p in ("/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe",
              "/mnt/c/Program Files/PowerShell/7/pwsh.exe"):
        if os.path.exists(p):
            return p
    return None


def to_win(path):
    """转成 Windows 侧可用的路径。已是 Windows 形式、或本来就在 Windows 上，原样返回。"""
    if not path:
        return path
    if IS_WIN or re.match(r"^[A-Za-z]:[\\/]", path) or path.startswith("\\\\"):
        return path
    m = re.match(r"^/mnt/([a-z])/(.*)$", path)
    if m:
        return "%s:\\%s" % (m.group(1).upper(), m.group(2).replace("/", "\\"))
    return path


def to_local(path):
    """转成当前系统可访问的路径。Windows 路径在 WSL 下换成 /mnt/x/...。"""
    if not path:
        return path
    if os.path.exists(path) or IS_WIN:
        return path
    m = re.match(r"^([A-Za-z]):[\\/](.*)$", path)
    if m and os.path.isdir("/mnt"):
        cand = "/mnt/%s/%s" % (m.group(1).lower(), m.group(2).replace("\\", "/"))
        if os.path.exists(cand):
            return cand
    return path


def office_xsl_candidates():
    """Office 的 MML2OMML.XSL 常见位置，按当前平台给原生路径。"""
    rel = [
        r"Microsoft Office\root\Office16\MML2OMML.XSL",
        r"Microsoft Office\root\Office15\MML2OMML.XSL",
        r"Microsoft Office\Office16\MML2OMML.XSL",
    ]
    out = []
    for base in (r"C:\Program Files", r"C:\Program Files (x86)"):
        out.extend(os.path.join(base, r) for r in rel)
    if not IS_WIN:
        out.extend(p.replace("\\", "/").replace("C:/", "/mnt/c/", 1) for p in list(out))
    out.append("/Applications/Microsoft Word.app/Contents/Resources/MML2OMML.XSL")
    return out


def glob_xsl():
    """兜底：在 Office 安装目录里递归找。"""
    pats = []
    for base in (r"C:\Program Files", r"C:\Program Files (x86)"):
        pats.append(os.path.join(base, "Microsoft Office", "**", "MML2OMML.XSL"))
    if not IS_WIN and os.path.isdir("/mnt/c"):
        pats += ["/mnt/c/Program Files/Microsoft Office/**/MML2OMML.XSL",
                 "/mnt/c/Program Files (x86)/Microsoft Office/**/MML2OMML.XSL"]
    for pat in pats:
        hits = glob.glob(pat, recursive=True)
        if hits:
            return hits[0]
    return None
