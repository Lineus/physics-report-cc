#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""项目记忆：维护 memory.md，记录文风、版式规则，以及用户每次手改的差异。

子命令：
    init  <根目录>                      用 memory.template.md 生成 memory.md
    show  <根目录>                      打印记忆的结构与条目数
    diff  <根目录> <旧.docx> <新.docx> [--n 5] [--kind report] [--no-write]
    note  <根目录> "<一句话>"            手动追加一条

diff 是核心：用户交回手改版后跑一次，把实质性改动追加到
memory.md 的「改动日志」。下次写稿先读 memory.md，就不再依赖「上一份成品」
当底稿——模板 + 记忆就够复现版式与文风。

只做段落 / 表格 / 图片三个层级，天然免疫 Word 重存噪声（run 拆分、
缩进微调、extent 漂移都不会改变段落文本）。
"""
import argparse
import datetime
import difflib
import hashlib
import os
import re
import shutil
import sys

MEM = "memory.md"
TEMPLATE = "memory.template.md"
LOG_HEAD = "## 改动日志"


# ---------------------------------------------------------------- 读取
W_T = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
M_T = "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"


def para_text(p):
    """段落全文，含公式。

    python-docx 的 p.text 只收 w:t，OMML 公式（m:t）会整段丢失，
    导致「含公式的句子」在 diff 里看起来被删掉。这里两种都收。
    """
    buf = []
    for node in p._p.iter():
        if node.tag in (W_T, M_T) and node.text:
            buf.append(node.text)
    return "".join(buf).strip()


def load(path):
    import docx
    d = docx.Document(path)
    paras = [para_text(p) for p in d.paragraphs]
    paras = [t for t in paras if t]
    cells = []
    for ti, t in enumerate(d.tables):
        for ri, row in enumerate(t.rows):
            for ci, c in enumerate(row.cells):
                cells.append((ti, ri, ci, c.text.strip()))
    hashes = []
    for rel in d.part.rels.values():
        if "image" in rel.reltype:
            hashes.append(hashlib.sha256(rel.target_part.blob).hexdigest()[:12])
    return paras, cells, sorted(hashes)


def clean(s, n=60):
    s = re.sub(r"\s+", " ", s).strip()
    return s if len(s) <= n else s[:n] + "…"


# ---------------------------------------------------------------- 差异
def diff_docx(old, new):
    pa, ca, ha = load(old)
    pb, cb, hb = load(new)
    lines = []

    sm = difflib.SequenceMatcher(None, pa, pb)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if tag == "delete":
            for t in pa[i1:i2]:
                lines.append("- 删除：" + clean(t))
        elif tag == "insert":
            for t in pb[j1:j2]:
                lines.append("+ 新增：" + clean(t))
        else:
            for a, b in zip(pa[i1:i2], pb[j1:j2]):
                lines.append("~ 修改：「%s」→「%s」" % (clean(a, 40), clean(b, 40)))
            if (i2 - i1) != (j2 - j1):
                lines.append("~ 段落数：%d → %d（自第 %d 段）" % (i2 - i1, j2 - j1, i1 + 1))

    if ca != cb:
        changed = sum(1 for x, y in zip(ca, cb) if x != y) + abs(len(ca) - len(cb))
        lines.append("~ 表格：%d 处单元格变更" % changed)

    if ha != hb:
        note = "（有替换）" if len(ha) == len(hb) and ha != hb else ""
        lines.append("~ 图片：%d 张 → %d 张%s" % (len(ha), len(hb), note))
    elif ha:
        lines.append("= 图片未变（%d 张）" % len(ha))

    return lines


# ---------------------------------------------------------------- 写记忆
def append_log(root, title, lines):
    path = os.path.join(root, MEM)
    if not os.path.exists(path):
        sys.exit("没有 %s：先跑 init" % MEM)
    text = open(path, encoding="utf-8").read()
    if LOG_HEAD not in text:
        text = text.rstrip() + "\n\n" + LOG_HEAD + "\n"
    block = "\n### %s\n" % title + "\n".join(lines) + "\n"
    if block.strip() in text:
        return False
    open(path, "w", encoding="utf-8").write(text.rstrip() + "\n" + block)
    return True


def cmd_init(args):
    root = args.root
    dst = os.path.join(root, MEM)
    if os.path.exists(dst):
        print("已存在，不覆盖：" + dst)
        return 0
    src = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", TEMPLATE)
    src = os.path.normpath(src)
    if not os.path.exists(src):
        sys.exit("缺少模板：" + src)
    shutil.copyfile(src, dst)
    print("已生成：" + dst)
    return 0


def cmd_show(args):
    path = os.path.join(args.root, MEM)
    if not os.path.exists(path):
        sys.exit("没有 %s：先跑 init" % MEM)
    for line in open(path, encoding="utf-8"):
        if line.startswith("#"):
            print(line.rstrip())
    return 0


def cmd_diff(args):
    lines = diff_docx(args.old, args.new)
    if not lines:
        print("没有实质差异。")
        return 0
    print("\n".join(lines))
    if args.no_write:
        return 0
    title = "%s 第 %s 次 %s（用户手改）" % (
        datetime.date.today().isoformat(), args.n or "?", "预习" if args.kind == "prep" else "正式")
    if append_log(args.root, title, lines):
        print("\n已写入 memory.md")
    else:
        print("\n相同记录已存在，未重复写入")
    return 0


def cmd_note(args):
    title = "%s 备注" % datetime.date.today().isoformat()
    if append_log(args.root, title, ["- " + args.text]):
        print("已写入 memory.md")
    else:
        print("相同记录已存在")
    return 0


def main():
    ap = argparse.ArgumentParser(description="项目记忆维护")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("init"); p.add_argument("root"); p.set_defaults(fn=cmd_init)
    p = sub.add_parser("show"); p.add_argument("root"); p.set_defaults(fn=cmd_show)

    p = sub.add_parser("diff")
    p.add_argument("root"); p.add_argument("old"); p.add_argument("new")
    p.add_argument("--n", default=""); p.add_argument("--kind", default="report",
                                                      choices=["prep", "report"])
    p.add_argument("--no-write", action="store_true")
    p.set_defaults(fn=cmd_diff)

    p = sub.add_parser("note"); p.add_argument("root"); p.add_argument("text")
    p.set_defaults(fn=cmd_note)

    args = ap.parse_args()
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
