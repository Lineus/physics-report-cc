#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""解析 config.yaml，算出本次任务的目录 / 文件名 / 模板。

用法：
    python3 plan.py <根目录> <n> <prep|report>

输出 KEY=VALUE，方便直接读；缺项打印 MISSING=... 并以 1 退出。
把「第几次实验 → 周次 / 实验名 / 教师」这件事交给脚本算，别让模型心算。
"""
import os
import sys

import yaml


def die(msg):
    print("ERROR=" + msg)
    raise SystemExit(1)


def main():
    if len(sys.argv) < 4:
        die("用法: python3 plan.py <根目录> <n> <prep|report>")
    root, n, kind = sys.argv[1], sys.argv[2], sys.argv[3]
    if kind not in ("prep", "report"):
        die("分支只能是 prep 或 report")

    cfg_path = os.path.join(root, "config.yaml")
    if not os.path.exists(cfg_path):
        die("找不到 config.yaml：先复制 config.example.yaml")
    cfg = yaml.safe_load(open(cfg_path, encoding="utf-8")) or {}

    stu = cfg.get("student") or {}
    name, sid = str(stu.get("name") or "").strip(), str(stu.get("id") or "").strip()
    missing = []
    if not name:
        missing.append("student.name")
    if not sid:
        missing.append("student.id")
    elif not (sid.isdigit() and len(sid) == 10):
        die("student.id 必须是 10 位数字，当前是 %r" % sid)

    entry = next((e for e in (cfg.get("schedule") or []) if str(e.get("n")) == str(n)), {})
    week = entry.get("week")
    experiment = entry.get("experiment")
    if week is None or not experiment:
        missing.append("schedule[n=%s]（week / experiment）" % n)

    dirs = cfg.get("dirs") or {}
    sub = (dirs.get(kind) or ("{n}-预" if kind == "prep" else "{n}")).format(n=n)
    out_dir = os.path.join(root, sub)

    tpl_key = "prep" if kind == "prep" else "report"
    tpl = ((cfg.get("templates") or {}).get(tpl_key) or "")
    tpl_abs = os.path.join(root, tpl) if tpl else ""
    if not tpl or not os.path.exists(tpl_abs):
        missing.append("templates.%s" % tpl_key)

    fname_tpl = cfg.get("filename") or "{name} {id} {experiment} 第{week}周"
    try:
        fname = fname_tpl.format(name=name, id=sid, experiment=experiment or "",
                                 week=week if week is not None else "")
    except KeyError as e:
        die("filename 模板里有未知字段 %s" % e)

    print("ROOT=" + os.path.abspath(root))
    print("KIND=" + kind)
    print("N=" + str(n))
    print("DIR=" + out_dir)
    print("TEMPLATE=" + (tpl_abs or ""))
    print("FILENAME=" + fname)
    print("OUT=" + os.path.join(out_dir, fname + ".docx"))
    print("WEEK=" + str(week if week is not None else ""))
    print("EXPERIMENT=" + str(experiment or ""))
    print("TEACHER=" + str(entry.get("teacher") or ""))
    print("DATE=" + str(entry.get("date") or ""))
    mem = os.path.join(root, "memory.md")
    print("MEMORY=" + mem)
    print("MEMORY_EXISTS=" + ("1" if os.path.exists(mem) else "0"))
    print("MISSING=" + ",".join(missing))
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
