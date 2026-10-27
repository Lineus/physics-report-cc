# -*- coding: utf-8 -*-
"""MathML → OMML（Word 原生公式），纯 Python，不依赖 Office。

mathomml 优先用 Office 的 MML2OMML.XSL；找不到时（只装 WPS、纯 Linux、macOS）
就退到本模块。覆盖 latex2mathml 实际会产出的元素：mrow / mi / mn / mo / mtext /
mfrac / msup / msub / msubsup / msqrt / mroot / mfenced / mover / munder。

用法：
    from omml import mathml_to_omml
    el = mathml_to_omml("<math>…</math>")
"""
import re

from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"m": M, "w": W}

# 这些元素的子节点要按顺序全部展开
CONTAINERS = {"mrow", "mstyle", "mphantom", "mpadded", "menclose",
              "semantics", "math", "mtd"}
SKIP = {"annotation", "annotation-xml", "mspace", "maligngroup", "malignmark"}


def _q(tag):
    return "{%s}%s" % (M, tag)


def _el(tag, parent=None):
    e = etree.Element(_q(tag), nsmap={"m": M})
    if parent is not None:
        parent.append(e)
    return e


def _kids(node):
    """MathML 的子元素是按位置排的（msub 没有 <base> 标签），只能按顺序取。"""
    return [c for c in node if isinstance(c.tag, str)]


def _run(text, parent, italic=False, bold=False):
    """一个 m:r。m:rPr 必须在 w:rPr 之前，m:t 最后。"""
    if not text:
        return None
    r = _el("r")
    if parent is not None:
        parent.append(r)
    if italic or bold:
        rpr = _el("rPr", r)
        if italic:
            sty = _el("sty", rpr)
            sty.set(_q("val"), "i")
        if bold:
            _el("nor", rpr)
    t = _el("t", r)
    t.text = text
    return r


def _local(node):
    return etree.QName(node).localname


def _text(node):
    return "".join(node.itertext())


TOKENS = {"mi", "mn", "mo", "mtext", "ms"}


def _append_children(node, parent):
    """把 node 的内容填进 parent。

    叶子节点（mi/mn/mo…）的文字在 .text 里而不是子节点里，必须直接 build，
    否则 m:e / m:sub 会生成为空。
    """
    if node is None:
        return
    if _local(node) in TOKENS:
        for built in _build(node):
            parent.append(built)
        return
    for child in _kids(node):
        for built in _build(child):
            parent.append(built)


def _wrap(node, tag, parent):
    """把 node 的内容放进 <m:tag>。"""
    holder = _el(tag, parent)
    _append_children(node, holder)
    return holder


def _build(node):
    """递归转换，返回要追加到父节点的元素列表。"""
    tag = _local(node)

    if tag in SKIP:
        return []

    if tag in CONTAINERS:
        holder = _el("e")
        _append_children(node, holder)
        # 透明容器：把内容提升回父级
        return list(holder)

    if tag in TOKENS:
        r = _run(_text(node), None, italic=(tag == "mi"))
        return [r] if r is not None else []

    if tag == "mfrac":
        f = _el("f")
        fpr = _el("fPr", f)
        kind = _el("type", fpr)
        kind.set(_q("val"), "bar")
        ks = _kids(node)
        num = ks[0] if len(ks) > 0 else None
        den = ks[1] if len(ks) > 1 else None
        for src, dst in ((num, "num"), (den, "den")):
            box = _el(dst, f)
            if src is not None:
                _append_children(src, box)
        return [f]

    if tag in ("msup", "msub", "msubsup"):
        ks = _kids(node)
        base = ks[0] if ks else None
        if tag == "msup":
            sup, sub = (ks[1] if len(ks) > 1 else None), None
        elif tag == "msub":
            sub, sup = (ks[1] if len(ks) > 1 else None), None
        else:
            sub, sup = (ks[1] if len(ks) > 1 else None), (ks[2] if len(ks) > 2 else None)
        name = {"msup": "sSup", "msub": "sSub", "msubsup": "sSubSup"}[tag]
        box = _el(name)
        e = _el("e", box)
        if base is not None:
            _append_children(base, e)
        if sub is not None and tag in ("msub", "msubsup"):
            sb = _el("sub", box)
            _append_children(sub, sb)
        if sup is not None and tag in ("msup", "msubsup"):
            sp = _el("sup", box)
            _append_children(sup, sp)
        return [box]

    if tag in ("msqrt", "mroot"):
        rad = _el("rad")
        rpr = _el("radPr", rad)
        _el("degHide", rpr).set(_q("val"), "0" if tag == "mroot" else "1")
        deg = _el("deg", rad)
        if tag == "mroot":
            ks = _kids(node)
            if len(ks) > 1:
                _append_children(ks[1], deg)
        e = _el("e", rad)
        if tag == "msqrt":
            _append_children(node, e)
        else:
            ks = _kids(node)
            if ks:
                _append_children(ks[0], e)
        return [rad]

    if tag in ("mfenced",):
        d = _el("d")
        dpr = _el("dPr", d)
        beg = node.get("open", "(")
        end = node.get("close", ")")
        if beg:
            _el("begChr", dpr).set(_q("val"), beg)
        if end:
            _el("endChr", dpr).set(_q("val"), end)
        e = _el("e", d)
        _append_children(node, e)
        return [d]

    if tag in ("mover", "munder", "munderover"):
        # 极限 / 上划线：用 sSup/sSub 近似，Word 里显示正确、可编辑
        ks = _kids(node)
        base = ks[0] if ks else None
        over = ks[1] if len(ks) > 1 else None
        name = "sSup" if tag == "mover" else "sSub"
        box = _el(name)
        e = _el("e", box)
        if base is not None:
            _append_children(base, e)
        part = _el("sup" if name == "sSup" else "sub", box)
        if over is not None:
            _append_children(over, part)
        return [box]

    if tag == "mtable":
        box = _el("eqArr")
        for row in node:
            if _local(row) != "mtr":
                continue
            e = _el("e", box)
            for cell in row:
                _append_children(cell, e)
        return [box]

    # 其余未知元素：降级成纯文本，至少不丢符号
    text = _text(node).strip()
    if text:
        r = _run(text, None)
        return [r] if r is not None else []
    return []


def mathml_to_omml(mathml, size=None):
    """MathML 字符串 → m:oMath 元素；失败返回 None。"""
    try:
        root = etree.fromstring(mathml.encode("utf-8"))
    except Exception:  # noqa: BLE001
        return None
    omath = _el("oMath")
    _append_children(root, omath)
    if len(omath) == 0:
        return None
    return omath
