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
    """公式链路：latex2mathml + Office 的 MML2OMML.XSL。"""
    ok = check_module("latex2mathml", "latex2mathml")
    try:
        import mathomml
        xsl = mathomml.find_xsl()
    except Exception as e:  # noqa: BLE001
        print(f"{BAD} mathomml 导入失败：{e}")
        return False
    if xsl:
        print(f"{OK} MML2OMML.XSL {xsl}")
    else:
        print(f"{BAD} 未找到 MML2OMML.XSL（Word 2007+ 自带）；公式会退回 _{{}}/^{{}} 写法")
        ok = False
    return ok


def check_word():
    ps = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
    if not os.path.exists(ps):
        print(f"{BAD} 未找到 Windows PowerShell：PDF 导出需要 Windows + Word")
        return False
    try:
        out = subprocess.run(
            [ps, "-NoProfile", "-Command",
             "(New-Object -ComObject Word.Application).Version"],
            capture_output=True, timeout=30,
        )
        ver = out.stdout.decode("utf-8", "ignore").strip()
        print(f"{OK} Word COM {ver}" if ver else f"{BAD} Word COM 不可用")
        return bool(ver)
    except Exception as e:  # noqa: BLE001
        print(f"{BAD} Word COM: {e}")
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
    ok = check_word() and ok
    print()
    print("环境就绪" if ok else "还有依赖没装好，见上面的提示")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
