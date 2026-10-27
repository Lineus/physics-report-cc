# -*- coding: utf-8 -*-
"""LaTeX → Word 原生公式（OMML）。

链路：latex2mathml 把 $...$ 转成 MathML → 用 Office 自带的 MML2OMML.XSL
做 XSLT → 得到 <m:oMath>，直接插进段的 XML。这样 Word 里是可编辑的真公式，
不是图片、也不是 Unicode 上下标拼出来的假货。

依赖：
    pip install latex2mathml
    lxml（python-docx 已依赖）
    Microsoft Office 的 MML2OMML.XSL（Word 2007 起都有）

找不到 XSL 时 add_math() 返回 False，调用方退回旧的 _{}/^{} mini-markup。
"""
import glob
import os

from lxml import etree

_XSL_CANDIDATES = [
    os.environ.get("MML2OMML_XSL", ""),
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
    # WSL 下 Windows 路径的前缀写法
    "/mnt/c/Program Files/Microsoft Office/root/Office16/MML2OMML.XSL",
    "/mnt/c/Program Files (x86)/Microsoft Office/root/Office16/MML2OMML.XSL",
    "/mnt/c/Program Files/Microsoft Office/Office16/MML2OMML.XSL",
    "/mnt/c/Program Files (x86)/Microsoft Office/Office16/MML2OMML.XSL",
    "/Applications/Microsoft Word.app/Contents/Resources/MML2OMML.XSL",
]


def _to_local(path):
    r"""把 Windows 路径转成当前系统可访问的形式（WSL 下 C:\\x -> /mnt/c/x）。"""
    if not path:
        return ""
    if os.path.exists(path):
        return path
    m = __import__("re").match(r"^([A-Za-z]):[\\/](.*)$", path)
    if m and os.path.isdir("/mnt"):
        cand = "/mnt/%s/%s" % (m.group(1).lower(), m.group(2).replace("\\", "/"))
        if os.path.exists(cand):
            return cand
    return path

_cache = {}


def find_xsl():
    """定位 MML2OMML.XSL；找不到返回 None。"""
    if "xsl" in _cache:
        return _cache["xsl"]
    found = None
    for p in _XSL_CANDIDATES:
        p = _to_local(p)
        if p and os.path.exists(p):
            found = p
            break
    if not found:
        for pat in ("/mnt/c/Program Files/Microsoft Office/**/MML2OMML.XSL",
                    "/mnt/c/Program Files (x86)/Microsoft Office/**/MML2OMML.XSL",
                    r"C:\Program Files\Microsoft Office\**\MML2OMML.XSL",
                    r"C:\Program Files (x86)\Microsoft Office\**\MML2OMML.XSL"):
            hits = glob.glob(pat, recursive=True)
            if hits:
                found = hits[0]
                break
    _cache["xsl"] = found
    return found


def available():
    """公式渲染是否可用。

    只要求 latex2mathml：Office 的 XSL 是首选，没有它就走纯 Python 的 omml
    回退（WPS / 纯 Linux / macOS 都是这条）。
    """
    try:
        import latex2mathml.converter  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


def backend():
    """当前会用哪条链路：'xsl'（Word 有 XSL）或 'python'（WPS / 无 Office）。"""
    if find_xsl() is not None:
        return "xsl"
    try:
        import omml  # noqa: F401
    except Exception:  # noqa: BLE001
        return "none"
    return "python"


def latex_to_omml(latex):
    """返回 m:oMath 元素；不可用时返回 None。"""
    try:
        import latex2mathml.converter as conv
    except Exception:  # noqa: BLE001
        return None

    xsl_path = find_xsl()
    if not xsl_path:
        return _python_fallback(conv.convert(latex))
    try:
        mathml = conv.convert(latex)
        dom = etree.fromstring(mathml.encode("utf-8"))
        xslt = _cache.get("xslt")
        if xslt is None:
            xslt = etree.XSLT(etree.parse(xsl_path))
            _cache["xslt"] = xslt
        res = xslt(dom)
        # 保留元素自带的 xmlns:m 声明——去掉后 lxml 无法再解析，
        # 而 Word 接受这种就地声明的 oMath。
        return etree.fromstring(etree.tostring(res, encoding="unicode"))
    except Exception:  # noqa: BLE001
        return _python_fallback(conv.convert(latex))


def _python_fallback(mathml):
    """没有 Office XSL 时用纯 Python 转换器（WPS / Linux / macOS）。"""
    try:
        from omml import mathml_to_omml
    except Exception:  # noqa: BLE001
        return None
    return mathml_to_omml(mathml)


def add_math(paragraph, latex, size=12):
    """把一段 LaTeX 作为行内公式插入 paragraph；成功返回 True。"""
    el = latex_to_omml(latex)
    if el is None:
        return False
    # 公式字号跟正文：给每个 m:r 加 w:rPr/w:sz。
    # 顺序有讲究：m:rPr → w:rPr → m:t，w:rPr 不能追加到 m:t 后面。
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for r in el.iter():
        if not r.tag.endswith("}r"):
            continue
        rpr = etree.Element(W + "rPr")
        sz = etree.SubElement(rpr, W + "sz")
        sz.set(W + "val", str(int(size * 2)))
        pos = 1 if len(r) and r[0].tag.endswith("}rPr") else 0
        r.insert(pos, rpr)
    paragraph._p.append(el)
    return True
