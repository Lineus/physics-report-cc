#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""扫描根目录下已有的报告，供新稿参考（不当作底稿）。

用法：
    python3 refs.py "<根目录>" [--n 3] [--kind prep|report] [--full]

输出每份报告的：路径 / 字数 / 图数 / 公式数 / 章节骨架 / 每节首段。
--full 时把正文全打出来（长报告慎用）。
"""
import argparse
import glob
import os
import re
import sys

W_T = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
M_T = "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"
HEAD_RE = re.compile(r"^([一二三四五六七八九十]+、|\d+[.．]|（\d+）|\(\d+\))")


def para_text(p):
    return "".join(n.text for n in p._p.iter()
                   if n.tag in (W_T, M_T) and n.text).strip()


def summarize(path, full=False):
    import docx
    d = docx.Document(path)
    paras = [para_text(p) for p in d.paragraphs]
    paras = [t for t in paras if t]
    chars = sum(len(re.sub(r"\s+", "", t)) for t in paras)
    with_math = sum(1 for p in d.paragraphs if any(
        n.tag == M_T for n in p._p.iter()))
    imgs = len(d.inline_shapes)

    print("=== %s" % os.path.relpath(path))
    print("    %d 字 / %d 段 / %d 图 / %d 段含公式" % (chars, len(paras), imgs, with_math))
    if full:
        for t in paras:
            print("    " + t[:200])
        return
    heads = [(i, t) for i, t in enumerate(paras) if HEAD_RE.match(t) and len(t) < 40]
    for i, t in heads[:20]:
        print("    " + t)
        nxt = next((p for p in paras[i + 1:i + 3] if len(p) > 10), "")
        if nxt:
            print("        · " + nxt[:70])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--n", default="")
    ap.add_argument("--kind", default="", choices=["", "prep", "report"])
    ap.add_argument("--full", action="store_true")
    a = ap.parse_args()

    root = a.root
    files = [p for p in glob.glob(os.path.join(root, "**", "*.docx"), recursive=True)
             if "_work" not in p.replace("\\", "/")
             and "模板" not in p and "template" not in p.lower()]
    if a.kind == "prep":
        files = [p for p in files if "预" in os.path.basename(os.path.dirname(p))]
    elif a.kind == "report":
        files = [p for p in files if "预" not in os.path.basename(os.path.dirname(p))]
    if a.n:
        files = [p for p in files if re.search(r"(^|[\\/])%s(-|预|$)" % re.escape(a.n),
                                               os.path.dirname(p))]
    if not files:
        print("没有可参考的既有报告。")
        return 1
    for p in sorted(files, key=os.path.getmtime):
        summarize(p, a.full)
    print("\n参考它们的：章节顺序、每节篇幅、图注写法、公式密度；不要照抄内容，也不要用它们当底稿。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
