#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""首次配置：建目录、生成 config/memory、登记旧报告、检查环境。

用法（Agent 先问用户，再把答案作为参数传进来）：

    python3 setup.py "<根目录>" \
        --name 张三 --id 1234567890 --class 机械2504 \
        --timetable 课表.png \
        --prep-template "模板/预习报告模板.docx" \
        --report-template "模板/实验报告模板.docx"

只跑机械部分、参数留空也行：它会打印 NEED 列表，告诉 Agent 还要问用户什么。
永不覆盖已有的 config.yaml / memory.md。
"""
import argparse
import glob
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TOP = os.path.dirname(HERE)

NEED = []


def say(tag, msg):
    print("%-5s %s" % (tag, msg))


def to_local(path):
    if not path:
        return path
    if os.path.exists(path):
        return path
    m = re.match(r"^([A-Za-z]):[\\/](.*)$", path)
    if m and os.path.isdir("/mnt"):
        return "/mnt/%s/%s" % (m.group(1).lower(), m.group(2).replace("\\", "/"))
    return path


def read_para_text(p):
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"
    return "".join(n.text for n in p._p.iter() if n.tag in (W, M) and n.text).strip()


def docx_meta(path):
    """从封面抓实验名称与指导教师。"""
    import docx
    name = teacher = ""
    try:
        d = docx.Document(path)
        for p in d.paragraphs[:40]:
            t = read_para_text(p)
            m = re.match(r"^实验名称[：:]\s*(.+)$", t)
            if m and not name:
                name = m.group(1).strip()
            m = re.match(r"^指导教师[：:]\s*(.+)$", t)
            if m and not teacher:
                teacher = m.group(1).strip()
    except Exception:  # noqa: BLE001
        pass
    return name, teacher


def scan_reports(root):
    """扫已有报告 → [(n, week, experiment, teacher, kind)]。"""
    rows = {}
    pat = re.compile(r"^(?P<name>.+?)\s+(?P<sid>\d{6,12})\s+(?P<exp>.+?)\s+第(?P<week>[^周]+)周")
    for path in glob.glob(os.path.join(root, "**", "*.docx"), recursive=True):
        rel = path.replace("\\", "/")
        if "/模板/" in rel or "/_work/" in rel or os.path.basename(path).startswith("~$"):
            continue
        folder = os.path.basename(os.path.dirname(path))
        m = re.match(r"^(\d+)\s*[-\u2010-\u2015]?\s*(预)?$", folder)
        if not m:
            continue
        n = m.group(1)
        kind = "prep" if m.group(2) else "report"
        fm = pat.match(os.path.splitext(os.path.basename(path))[0])
        exp = fm.group("exp") if fm else ""
        week = fm.group("week") if fm else ""
        d_exp, d_teacher = docx_meta(path)
        key = n
        if key not in rows:
            rows[key] = {"n": n, "week": week, "exp": exp or d_exp, "teacher": d_teacher,
                         "kinds": set()}
        rows[key]["kinds"].add(kind)
        if not rows[key]["exp"]:
            rows[key]["exp"] = exp or d_exp
        if not rows[key]["teacher"]:
            rows[key]["teacher"] = d_teacher
        if not rows[key]["week"]:
            rows[key]["week"] = week
    out = []
    for n in sorted(rows, key=lambda x: int(x)):
        r = rows[n]
        label = "预习+正式" if len(r["kinds"]) == 2 else ("预习" if "prep" in r["kinds"] else "正式")
        out.append((r["n"], r["week"], r["exp"], r["teacher"], label))
    return out


def write_table(root, rows):
    """把旧报告登记进 memory.md 的「已确认的实验清单」。"""
    path = os.path.join(root, "memory.md")
    if not os.path.exists(path):
        return False
    s = open(path, encoding="utf-8").read()
    head = "## 已确认的实验清单"
    if head not in s:
        return False
    lines = s.splitlines()
    try:
        i = next(k for k, t in enumerate(lines) if t.strip() == head)
    except StopIteration:
        return False
    j = i + 1
    while j < len(lines) and not lines[j].startswith("|"):
        j += 1
    if j >= len(lines):
        return False
    k = j + 2                                  # 跳过表头与分隔行
    while k < len(lines) and lines[k].startswith("|"):
        k += 1
    table = ["| n | 周次 | 实验名称 | 教师 | 备注 |",
             "| --- | --- | --- | --- | --- |"]
    for n, week, exp, teacher, label in rows:
        table.append("| %s | %s | %s | %s | %s |" % (n, week, exp, teacher, label))
    new = lines[:j] + table + lines[k:]
    open(path, "w", encoding="utf-8").write("\n".join(new) + "\n")
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--name", default="")
    ap.add_argument("--id", default="")
    ap.add_argument("--class", dest="klass", default="")
    ap.add_argument("--timetable", default="课表.png")
    ap.add_argument("--prep-template", default="")
    ap.add_argument("--report-template", default="")
    a = ap.parse_args()

    root = to_local(a.root) or a.root
    os.makedirs(root, exist_ok=True)
    say("OK", "根目录：" + os.path.abspath(root))

    # 1) 目录骨架
    tpl_dir = os.path.join(root, "模板")
    os.makedirs(tpl_dir, exist_ok=True)
    say("OK", "已准备 模板/")

    # 2) config.yaml
    cfg_path = os.path.join(root, "config.yaml")
    if os.path.exists(cfg_path):
        say("SKIP", "config.yaml 已存在，不覆盖")
    else:
        shutil.copyfile(os.path.join(TOP, "config.example.yaml"), cfg_path)
        say("OK", "已生成 config.yaml（模板复制）")

    # 3) 写入用户提供的信息
    if a.name or a.id or a.klass or a.prep_template or a.report_template:
        import yaml
        cfg = yaml.safe_load(open(cfg_path, encoding="utf-8")) or {}
        cfg.setdefault("student", {})
        cfg.setdefault("templates", {})
        if a.name:
            cfg["student"]["name"] = a.name
        if a.id:
            if not (a.id.isdigit() and len(a.id) == 10):
                say("ERR", "学号必须是 10 位数字：%r" % a.id)
                return 1
            cfg["student"]["id"] = a.id
        if a.klass:
            cfg["student"]["class"] = a.klass
        if a.timetable:
            cfg["timetable"] = a.timetable
        if a.prep_template:
            cfg["templates"]["prep"] = a.prep_template
        if a.report_template:
            cfg["templates"]["report"] = a.report_template
        open(cfg_path, "w", encoding="utf-8").write(
            yaml.safe_dump(cfg, allow_unicode=True, sort_keys=False, default_flow_style=False))
        say("OK", "config.yaml 已写入")

    # 4) memory.md
    mem = os.path.join(root, "memory.md")
    if os.path.exists(mem):
        say("SKIP", "memory.md 已存在，不覆盖")
    else:
        shutil.copyfile(os.path.join(TOP, "memory.template.md"), mem)
        say("OK", "已生成 memory.md")

    # 5) 登记旧报告
    rows = scan_reports(root)
    if rows:
        if write_table(root, rows):
            say("OK", "已登记 %d 次旧实验到 memory.md" % len(rows))
        for n, week, exp, teacher, label in rows:
            say("REF", "第%s次 第%s周 %s %s（%s）" % (n, week, exp, teacher, label))
    else:
        say("INFO", "没找到旧报告；第一次从模板起草即可")

    # 6) 模板检查
    for kind, given in (("prep", a.prep_template), ("report", a.report_template)):
        path = given or ""
        if path and not os.path.exists(to_local(path) if os.path.isabs(path)
                                       else os.path.join(root, path)):
            say("ERR", "%s 模板不存在：%s" % (kind, path))
            NEED.append("%s 模板的实际路径" % ("预习报告" if kind == "prep" else "实验报告"))
        elif not path:
            hits = glob.glob(os.path.join(tpl_dir, "*.docx"))
            hit = [h for h in hits if ("预习" in os.path.basename(h)) == (kind == "prep")]
            if hit:
                say("OK", "发现 %s 模板：%s" % (kind, os.path.basename(hit[0])))
            else:
                NEED.append("%s 模板（放进 模板/）" % ("预习报告" if kind == "prep" else "实验报告"))

    # 7) 课表
    tt = os.path.join(root, a.timetable or "课表.png")
    if os.path.exists(tt):
        say("OK", "课表：%s" % os.path.basename(tt))
    else:
        NEED.append("课表截图（默认 课表.png；也可在 config.yaml 里手填 schedule）")

    # 8) 身份
    import yaml
    cfg = yaml.safe_load(open(cfg_path, encoding="utf-8")) or {}
    stu = cfg.get("student") or {}
    if not (stu.get("name") or "").strip():
        NEED.append("姓名")
    if not (stu.get("id") or "").strip():
        NEED.append("10 位学号")

    # 9) 环境
    print()
    print("环境：")
    try:
        import subprocess
        env = subprocess.run([sys.executable, os.path.join(HERE, "check_env.py")],
                             capture_output=True, text=True, timeout=120)
        for line in (env.stdout or "").strip().splitlines():
            print("  " + line)
    except Exception as e:  # noqa: BLE001
        print("  跳过环境检查：%s" % e)

    print()
    if NEED:
        print("NEED=" + "；".join(NEED))
        print("把上面这些问用户，拿到后再跑一次 setup.py（参数同上）。")
        return 1
    print("READY=1")
    print("下一步：对 Agent 说「写第 n 次实验的预习报告」。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
