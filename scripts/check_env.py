#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查写实验报告所需的环境。用法：python3 scripts/check_env.py"""

import importlib
import os
import shutil
import subprocess
import sys

OK = "OK  "
BAD = "FAIL"


def check_module(name, pip_name=None):
    try:
        m = importlib.import_module(name)
        ver = getattr(m, "__version__", "")
        print(f"{OK} {name} {ver}")
        return True
    except Exception as e:  # noqa: BLE001
        print(f"{BAD} {name}: {e}  -> pip install {pip_name or name}")
        return False


def check_math():
    """公式链路：latex2mathml 必需；Office XSL 可选（没有就走纯 Python 转换）。"""
    ok = check_module("latex2mathml", "latex2mathml")
    try:
        import mathomml
    except Exception as e:  # noqa: BLE001
        print(f"{BAD} mathomml 导入失败：{e}")
        return False
    backend = mathomml.backend()
    if backend == "xsl":
        print(f"{OK} 公式转换：Office XSL（{mathomml.find_xsl()}）")
    elif backend == "python":
        print(f"{OK} 公式转换：纯 Python 回退（未装 Office，WPS / Linux 也适用）")
    else:
        print(f"{BAD} 公式转换不可用：缺 latex2mathml 或 scripts/omml.py")
        ok = False
    return ok


def check_office():
    """导出 PDF 的三条路：Word COM / WPS COM / LibreOffice，有一条就够。"""
    found = []
    ps = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
    if os.path.exists(ps):
        for name, com in (("Word", "Word.Application"), ("WPS", "KWPS.Application")):
            try:
                out = subprocess.run(
                    [ps, "-NoProfile", "-Command",
                     f"try {{ $a = New-Object -ComObject {com}; Write-Output $a.Version; $a.Quit() }} catch {{ }}"],
                    capture_output=True, timeout=40)
                ver = out.stdout.decode("utf-8", "ignore").strip()
            except Exception:  # noqa: BLE001
                ver = ""
            if ver:
                found.append(f"{name} {ver}")
    for exe in ("soffice", "libreoffice"):
        if shutil.which(exe):
            found.append(exe)
            break
    if found:
        print(f"{OK} 导出 PDF：{' / '.join(found)}")
        return True
    print(f"{BAD} 没有可用的 PDF 导出方式（Word / WPS / LibreOffice 都没装）")
    return False


def main():
    ok = True
    for mod, pip in [
        ("docx", "python-docx"),
        ("fitz", "pymupdf"),
        ("PIL", "pillow"),
        ("reportlab", "reportlab"),
    ]:
        ok = check_module(mod, pip) and ok

    ok = check_math() and ok
    ok = check_office() and ok
    print()
    print("环境就绪" if ok else "还有依赖没装好，见上面的提示")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
