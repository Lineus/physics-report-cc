---
name: "physics-report-cc"
description: "浙大物理实验报告技能（任何 agent 可用）：首次安装引导配置；之后写预习/正式报告——模板 + memory 出稿、课件只有 .pptx 时抽幻灯片文本并转 PDF 取图、LaTeX 转 Word 原生公式、改 build.py 后重建、Windows 下导 PDF 并逐页目视校对、统计各板块字数与图表编号、把同学报告的对比结论写进 memory.md"
whenToUse: >-
  用户说「写第 N 次实验报告」「写预习报告」「下周实验的预习」；报告出稿后说「导出 PDF」
  「逐页校对」「公式是不是原生公式」「统计各板块字数 / 查图表编号重号」；
  预习目录里只有 `.pptx` 课件、没有 `ppt.pdf`；图注的「图1 / 图2」被换行拆开要修；
  要改 build.py 里的公式或数字再重建；要把同学实验报告的对比结论记进 memory.md；
  或交回手改过的成品要你记住改动。
---

# 浙大物理实验报告（预习 + 正式）

面向**任何能跑 shell、能读写文件的 agent**：OpenClaw、Claude Code、Codex CLI、Gemini CLI、Cursor、Cline、DSH 等。
安装方式见 `README.md`；各宿主差异见 `references/agents.md`。

## 何时用

- **用户刚 clone 下来、说「装好了 / 帮我配置一下」，或什么都没说就把仓库地址给了你** → 先跑 `setup.py`，见下节
- 用户说「写第 N 次实验的预习报告」「下周实验的预习」
- 预习目录里只有 `xxx.pptx`、没有 `ppt.pdf`（课件只给 PPT 原件）→ 见「课件只有 .pptx」
- 用户说「写第 N 次实验报告」「正式报告」，并给了原始记录 / 思考题照片
- 报告已出 docx：要导同名 PDF、**逐页看图校对**、统计各板块字数、查图表编号重号
- 出稿后发现「图1」被排成「图」+ 换行 +「1」，要修图号断行 → 见「图号别被 CJK/数字断行拆开」
- 要改 `build.py` 里的公式或数字后**重建** → 见「改 build.py：先读再补丁」
- 要**对比同学的实验报告**并把结论写进 `memory.md` → 见「对照与写入记忆」
- 用户交回**手改过的成品**，要你记住改动 → 见「用户手改后」一节

## 先判分支

| 触发 | 分支 | 细读 |
| --- | --- | --- |
| 预习报告、`N-预/` | A 预习 | `references/prep.md` + 本页「课件只有 .pptx」 |
| 正式报告、原始记录、`N/` | B 正式 | `references/report.md` |

先读本页公共块，再读对应分支。**两份不要串线**：

- 500 字上限只属预习综述（单实验 500，双实验 800）；正式报告不按这个数压。
- 预习不碰数据处理；正式报告必须 `calc.py` 算、`figs.py` 出图，不用理论值冒充实测。
- 原始记录照片、思考题只出现在正式报告。
- **输入只有 `.pptx` 不算缺材料**：按「课件只有 .pptx」一节先抽文本、转 PDF、取图，再回到预习分支；
  不要因为 `references/prep.md` 只写了 `ppt.pdf` 就停下来问用户。
- **本技能不联网抓课件**：课件由用户自己下好放进来。

## 首次使用：先引导配置，再写稿

用户刚 clone 完、说「帮我装好 / 配置一下」，或根目录没有 `config.yaml` 时，**先做这一步**：

    python3 <技能目录>/scripts/setup.py "<根目录>"

它会：建 `模板/`、**按 12 次实验预建 `1/` `1-预/` … `12/` `12-预/`**、生成 `config.yaml` 与 `memory.md`、
扫旧报告登记进 memory.md 的「实验清单」、检查 Python 依赖与 Word，最后打印 `NEED=…` 或 `READY=1`。
**已有的目录 / config.yaml / memory.md 不会被覆盖或清空**（可重复跑）。

- 次数可改：`--count 16`；不想要空目录：`--no-folders`。目录名跟随 `config.yaml` 的 `dirs`。
- 建完要告诉用户：**旧报告按编号放进去**（第 3 次预习 → `3-预/`，正式 → `3/`），
  工具会扫这些目录学结构；放不放都不影响出稿。

- `NEED` 非空 → 把这几项**一次问完**（姓名、10 位学号、课表、两份模板、周次与实验名），
  别让用户自己去改 YAML；模板文件名字不同就传 `--prep-template` / `--report-template`。
- 拿到答案后带参数重跑，直到 `READY=1`：

        python3 <技能目录>/scripts/setup.py "<根目录>" \
            --name 张三 --id 1234567890 --class 机械2501 \
            --exp 3 --week 5 --experiment "用霍尔法测直流圆线圈与亥姆霍兹线圈磁场" \
            --teacher 王斐然 --date 2026-10-20

  `--exp/--week/--experiment` 写进 `config.yaml` 的 `schedule`，`plan.py` 靠它拼文件名。
  没写这两项时 `plan.py` 会留空 `FILENAME/OUT` 并给 `HINT`，**不要**拿残缺文件名去建文件。

- 用户只放了旧报告、没有模板：先让他在 `模板/` 放两份 docx，这是硬前提，不要自己造模板。
- 配置完成后才开始写稿；此后每次写稿都从「第 0 步」起。

## 第 0 步：plan.py + memory.md

    python3 <技能目录>/scripts/plan.py "<根目录>" <n> <prep|report>

输出 `DIR / TEMPLATE / OUT / WEEK / EXPERIMENT / TEACHER / MEMORY`，缺项打 `MISSING=`。
**目录名、周次、文件名都从它来，不要自己算**；`MISSING` 非空先问用户补齐，不猜。

拿到 `MEMORY` 路径后：

1. 没有 `memory.md` → `python3 <技能目录>/scripts/memory.py init "<根目录>"`，然后问用户关键约定（文风、字数、图注习惯）并填进去。
2. 有就读完。**「文风约定 / 版式约定」优先于本文件默认值**，冲突以记忆为准。
3. **永远从模板起草**，不拿上一份成品当底稿。版式靠模板，风格靠 memory.md，内容靠这次的材料。

写之前可先扫一遍旧稿，看章节顺序与篇幅分布（只看结构，不抄内容、不当底稿）：

    python3 <技能目录>/scripts/refs.py "<根目录>" [--n 3] [--kind prep|report] [--full]

## 课件只有 .pptx：抽文本 → 转 PDF → 取图

预习目录常见的组合是 `实验手册.pdf`（或 `*讲义*.pdf`）+ `xxx.pptx`，不一定有 `ppt.pdf`。
`.pptx` 是 zip 包，**不要因为缺 `ppt.pdf` 就停下来问用户**；按下面四步取出文字和图，
再回到 `references/prep.md` 第 2–6 节走模板 → build.py → export_pdf → proof.py。

### 1 认解释器：python-pptx 与 pymupdf 可能不在同一个 python

- 系统 `python` 往往**没有** `python-pptx`：直接 `from pptx import Presentation` 会报
  `ModuleNotFoundError: No module named 'pptx'`，`ppt_text.txt` 也不会生成。
- DSH 里先调 `load_workspace_dependencies`，用返回 JSON 的 `python` 字段跑抽取
  （本机示例：`C:\Users\Linus\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe`）。
  其他宿主用 `py -0p` / `which -a python` 挨个试 `-c "import pptx"`。
- 渲 PDF 页、裁图常用系统 `python`（带 `pymupdf`、`Pillow`）。**两件事允许用两个解释器**：
  以 `-c "import pptx"`、`-c "import pymupdf"` 当场确认，缺包就换解释器，**不要 pip install**。

### 2 抽逐页文本到 `_work/ppt_text.txt`

用带 `python-pptx` 的解释器跑。先写成 `_work/pptx_extract.py` 再执行，别在命令行里塞长 `python -c`：

```python
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io, sys

p = Presentation(sys.argv[1])
lines = []
for i, s in enumerate(p.slides, 1):
    txt, pics = [], 0
    for sh in s.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            pics += 1
        if sh.has_text_frame and sh.text_frame.text.strip():
            txt.append(sh.text_frame.text.strip().replace("\n", " / "))
    lines.append("=== slide %d pics=%d :: %s" % (i, pics, " || ".join(txt)))
io.open(sys.argv[2], "w", encoding="utf-8").write("\n".join(lines))
```

```bash
"<runtime python>" "N-预/_work/pptx_extract.py" "N-预/xxx.pptx" "N-预/_work/ppt_text.txt"
```

- 落盘 `N-预/_work/ppt_text.txt`；每页一行 `=== slide N pics=K :: …`，shape 间用 ` || `、shape 内换行用 ` / `。
- 写正文时**措辞与图号都能回到具体 slide**（例：slide 15 给「交流电桥的基本电路」与平衡条件）；
  两页说法冲突（例如同一旋钮在两页给不同档位）时不要二选一硬写，正文回避无定论档位。
- 抽不出文字的页（纯图 / SmartArt）记下页号，靠下一步渲图看。

### 3 用 LibreOffice 转 PDF

```bash
cd "<根目录>/N-预/_work" && \
  "C:/Program Files/LibreOffice/program/soffice.exe" --headless --convert-to pdf --outdir . "../xxx.pptx"
```

- 先找 `C:\Program Files\LibreOffice\program\soffice.exe`；`where soffice` 找不到也不影响，直接写全路径。
- 同名 PDF 落在 `_work/`；`ls -la *.pdf` 确认，再 `pymupdf.open(pdf).page_count`，页数应等于 slide 数。
- stderr 的 `Could not find platform independent libraries <prefix>` 是无害噪声，以生成的 PDF 页数为准。

### 4 渲关键页 + 从两条路取图

```python
import os, pymupdf
d = pymupdf.open(pdf)
os.makedirs("slides_png", exist_ok=True)
for i in (2, 4, 5, 10, 12, 13, 14, 15, 18, 22, 26):
    d[i - 1].get_pixmap(dpi=110).save("slides_png/s%02d.png" % i)
```

- 只渲含原理图、电路图、公式的十几页，存 `_work/slides_png/`，逐张 `read_image` 目检：
  图有没有被 SmartArt 拆散、文字有没有出框、整页是不是纯图。
- **原图路**：从 zip 的 `ppt/media/` 取。先用 `ppt/slides/_rels/slideN.xml.rels` 看该页引用了哪些
  `imageN.*`，只取目标页那几张（整包几十张不要全倒）：

```python
import re, zipfile
z = zipfile.ZipFile(pptx)
for n in z.namelist():
    m = re.match(r"ppt/slides/_rels/slide(\d+)\.xml\.rels$", n)
    if m:
        rels = re.findall(r'Target="\.\./media/([^"]+)"', z.read(n).decode("utf-8"))
        print("slide%s" % m.group(1), rels)
```

  取出的图先 `Image.open()` 看实际 `format/size/mode`；PPT 图常是 RGBA，用
  `pdf_extract_figs.to_white()` 铺白底、`getbbox()` 收紧白边，另存 `_work/pmedia/*.png`，
  **不要直接插原始 media 文件**。
- **PDF 裁剪路**：矢量电路图（没有独立位图）用 `scripts/pdf_extract_figs.py` 的 `render_clip()`，
  或 `page.get_pixmap(dpi=300~600, clip=fitz.Rect(...))`；图注自带的「图N」用 `crop_by_caption()` 裁掉，
  免得和报告自编图号打架。
- 裁完必须 `read_image` 复检：四角纯白、非白像素 bbox 贴住图形边界、没有带入原图的「图N」或断行括号。
- 图选好后按 `references/prep.md` 第 2–6 节：先定图号与配图段落，再进 `build.py`；**预习稿至少 1 张图要来自本次课件**（`ppt/media/` 原图或转出 PDF 的裁剪），不能只用手册图。

## 目录与配置

    <实验报告根目录>/
      config.yaml            # 姓名/学号/路径，含个人信息，勿提交
      memory.md              # 项目记忆，含个人信息，勿提交
      课表.png                # 课表截图；读不出来就用 config.schedule 手填
      模板/
        预习报告模板.docx
        实验报告模板.docx
      1-预/  实验手册.pdf  ppt.pdf     # 第 1 次预习，用户自己放（只有 .pptx 也行）
      1/     原始记录.jpg  思考题.jpg   # 第 1 次正式，用户自己放

- 目录名由 `config.yaml` 的 `dirs` 决定，默认 `{n}-预` / `{n}`；也支持 `{n}预`。
- 没有 `config.yaml`：从 `config.example.yaml` 复制，逐项问姓名、10 位学号、模板路径。学号必须 10 位，不沿用旧稿。
- 缺材料：先建目录，说清放什么，**不要编造课件内容**；只有 `.pptx` 不算缺，按「课件只有 .pptx」一节处理。
- 用户没指定根目录时问一次，并记进 memory.md。

## 命名与学号

- 文件名 = `{name} {id} <实验名称> 第{week}周.docx`，由 `config.yaml` 的 `filename` 生成。
- `N` 是第几次实验，`W` 是课表周次，**不是同一个数**（如 `3-预` = 第 5 周）。
- 交付前复核：文件名学号、封面学号、页数三处一致。

## 课表定字段

- 课表是图片时用宿主的多模态能力直接看；看不了就复制到临时目录再让宿主读。**读不到就问用户**，不要猜。
- 取实验名称、周次、日期、星期、上/下午、指导教师；实验名称照课表原文，并与课件首页核对。
- 封面逐项重填，不沿用上一份；读不出来就用 `config.yaml` 的 `schedule`。

## 原始记录：放大读数，拿不准就问

正式报告的全部数字都从原始记录照片来。照片字小、有划改、有格子空着，**读图是一级流程，不是随手一看**。

1. **先落盘**：附件复制到 `N/原始记录.jpg`（一条 `cp` 就够；别在同一条命令里再塞 `python -c`，括号会把 bash 打断）。
2. **整图 → 分区 → 单格**，逐级放大，每一级都 `read_image`：
   - 先看整图，确认表头朝向；手机照片常横躺，四个方向各存一张比一比。
   - 表格行区间用墨迹投影定位，再按 bbox 裁剪，**不要盲裁**：

```python
from PIL import Image
import numpy as np
a = np.asarray(Image.open(src).convert("L"))
ink = (a < 160).sum(axis=1)                        # 每行墨点数
rows = [i for i, n in enumerate(ink) if n > 20]    # 有字/有格线的行
```

   - 仍有歧义的格子单独放大 2–2.6 倍，存 `N/_work/*.png`，逐张看。
3. **只读原始文件**：裁剪、放大、转正都写到 `N/_work/`，`原始记录.jpg` 保持原样（交付时直接用它）。
4. **读数三问**：这个数属于哪个量？单位是什么？和相邻数据、公式能不能对上？（例：同一根金属棒两段 R 差 7 倍 = 对不上。）
5. **答不上来就停笔**：写稿前编号列出疑点——读到什么值、依据哪张图、哪里算不通——等用户确认。用户说过「有不清楚就问我」时这是硬闸门，先问后写。
6. **不按惯例补**：格子空着就空着，单位没写就问；量程、分辨率、仪器允差课程没给的，不要写进「现象 / 操作」。

## 公式：写 LaTeX，不拼上下标

正文用 `$...$` 包 LaTeX，脚本转成 **Word 原生公式**（可编辑、字号跟正文）：

- `$B=\frac{\mu_0 N I R^{2}}{2(R^{2}+x^{2})^{3/2}}$`
- `$u_c=\sqrt{u_A^{2}+u_B^{2}}$`

- 行内公式不必另起段；独立成行时整段只放一个 `$...$`。
- 同一符号前后统一（不要一会儿 `B_0` 一会儿 `B0`）。
- 不用 Unicode 上下标字符，也不把公式贴成图片。
- 没装 Office 时自动走 `scripts/omml.py` 的纯 Python 转换（WPS / Linux / macOS 都有原生公式）；只有两条都不可用才退回 `_{}`/`^{}` 并打 `WARN`，此时把 `WARN` 如实告诉用户。
- 重建后按「公式仍是原生 OMML 的核对」验一遍再交付。

## 调用脚本

`docx_fill` / `mathomml` / `chart` 是**库**，不是命令行工具。写 `build.py` 时先把脚本目录加入 `sys.path`，
否则 `import docx_fill` 会失败（`mathomml` 由 `docx_fill` 内部导入，同目录即可）：

```python
import sys
sys.path.insert(0, "<技能目录>/scripts")
from docx_fill import add_rich, clear_between, figure_block, find, open_base, text_block
from chart import chart
```

命令行脚本（`plan` / `refs` / `proof` / `export_pdf` / `memory` / `setup` / `check_env`）直接运行。
本文件里写的 `python3` 在 Windows 上换成 `py -3` 或 `python`；路径分隔符用 `\`，两者别混。

**不同 python 装不同的包**：`python-pptx` 常只在宿主运行时解释器里，`pymupdf` / `Pillow` 常在系统 python 里。
按包挑解释器（`<解释器> -c "import pptx"` / `-c "import pymupdf"`），缺包就换解释器，**不要 pip install**。

平台差异集中在 `scripts/hostenv.py`（找 PowerShell、Windows↔WSL 路径互转、Office XSL 位置），
脚本里不要再自己判断 `os.name` 或拼 `/mnt/...`。

## 改 build.py：先读再补丁

补丁吃掉一个反斜杠，公式就退成 LaTeX 原文。按下面顺序做：

1. **先读原文**：`read` 目标文件（带 `offset`/`limit` 看附近行），`old_string` 从读到的原文复制，不凭记忆手打。
2. **用 `edit` 精确替换**，`old_string` 必须唯一；补丁脚本里加 `assert s.count(old) == 1`，命中不了先 `grep -n` 看实际文本。
3. **含反斜杠 / 换行的行，改按行号替换**，片段用 `chr(92)`（`\`）和 `chr(10)`（换行）拼：

```python
import io
p = "4/_work/build.py"
lines = io.open(p, encoding="utf-8").read().split(chr(10))
i = next(k for k, ln in enumerate(lines) if "由下式逐对求" in ln)
assert lines[i].startswith("body(")            # 先确认打中的是哪一行
lines[i] = 'body("…$' + chr(92) + 'Delta t' + chr(92) + 'approx5$…")'
io.open(p, "w", encoding="utf-8").write(chr(10).join(lines))
```

4. **不要在 shell heredoc 里塞多行字符串**：真实换行被写进源码 → `SyntaxError: unterminated string literal`；`\a`、`\n` 被吃掉 → 重建打出 `WARN 公式未能转成 OMML，按原文写入：lpha`。复杂补丁写成 `_work/patchN.py` 再跑。
5. **`%` 格式化**：字面百分号写 `%%`（`$0.8\%%$`）。报 `not enough arguments for format string` 先数占位符，别乱加参数；新写的段优先 f-string，不用转义。
6. **一次一类改动**：文字、数字、分页分开跑。每轮 `grep -n` 确认目标行、看构建输出无 `WARN 公式`，再重导 PDF 校对。

## 公式仍是原生 OMML 的核对

构建输出里**不能**有 `WARN 公式未能转成 OMML`（`docx_fill.add_rich` 在 `mathomml.add_math` 失败时打这条，并把 `$...$` 原文写进正文）。再数一遍 `m:oMath` 段：

```python
import docx
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
d = docx.Document(r"4/蔡成 3250103196 直流双臂电桥 第6周.docx")
print("oMath paras", sum(1 for p in d.paragraphs if any(e.tag == M for e in p._p.iter())))
```

- 与 `proof.py` 表头「含公式段：N」对照（同一算法）；N 应等于正文公式段数，明显变少就是公式退化了。
- `p.text` 只收 `w:t`，不含公式文字，别拿它数字数或判断公式在不在。
- 例：一份 50 段公式的稿子，脚本与 `proof.py` 都报 50，可作基线。

## Windows 控制台：别直印中文

- 控制台代码页是 GBK：直接 print `−`（U+2212）、`μ`、`Ω` 会 `UnicodeEncodeError: 'gbk' codec can't encode character`；中文常乱码成 `ʵ�����ơ`。
- 两种写法都行，按场合选：
  - 命令前加 `PYTHONIOENCODING=utf-8`：`PYTHONIOENCODING=utf-8 python 4/_work/dump.py …`；
  - 结果写 UTF-8 文件再用 `read` 看（长文本、含公式符号时首选）：`io.open(out, "w", encoding="utf-8").write(txt)`。
- PowerShell 的输出也是控制台代码页，`export_pdf.py` 内部按 utf-8 → gb18030 依次解码。
- 命令行里别塞长 `python -c "…(…)"`：括号、引号先被 bash 解释，报 `syntax error near unexpected token '('`。写成 `_work/*.py` 再跑。

## 图号别被 CJK/数字断行拆开

Word 允许在中文与数字之间断行：正文「……（图2）」可能排成「图」+ 换行 +「2」。
`proof.py` 只按行首认图注，查不出这种断行，只能从 PDF 文本或放大截图发现。

- **正文引用**在「图 / 表」与数字之间插 U+2060 WORD JOINER：
  `"（图\u20602）"`、`"表\u20603"`。docx / PDF 里不可见，只让两个字黏在一起；
  `proof.py` 的 `re.sub(r"\s+", "", …)` 不会去掉它，**每处字数 +1**，500 字上限要留余量。
- **图注行首保持 `图N` 原样，不要加 joiner**：`proof.py` 用 `^图\s*\d` 认图注，加了会漏数；
  图注短、居中，通常不会断，靠下面的截图目检确认。
- **查**（系统 python，带 pymupdf；结果写文件或 `PYTHONIOENCODING=utf-8`）：

```python
import pymupdf, re
d = pymupdf.open(pdf)
t = "\n".join(p.get_text() for p in d)
bad = re.findall(r"[图表]\s*\n\s*\d", t)
print("BROKEN", len(bad))
```

- 有命中就改 `build.py` 里的字面量（别手改 docx），重建 → 重导 PDF → 再查；同时把图注区域放大截图，
  确认「图1」「图2」都在同一行，`proof.py` 仍 `结论：无异常`。

## 导出与逐页校对（Windows）

**先在宿主的 PowerShell 工具里内联 COM 导出**（DSH 的 `pwsh` 工具就是 Windows PowerShell 5.1，直接粘下面这段，不要包进 .ps1）：

```powershell
$src = 'C:\Users\Linus\Desktop\physical-report-workspace\4\蔡成 3250103196 直流双臂电桥 第6周.docx'
$dst = [IO.Path]::ChangeExtension($src, '.pdf')
if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($src, $false, $true)
  $doc.ExportAsFixedFormat($dst, 17)      # 17 = PDF
  $doc.Close(0)
  "PDF_OK=$dst"
} finally { $word.Quit() }
```

- **不要** `pwsh -File x.ps1`：本机 `pwsh` 不在 PATH，报「无法将“pwsh”项识别为 cmdlet」。**不要** `& x.ps1`：未签名脚本被执行策略拒绝（`无法加载文件…未进行数字签名`）。
- 结束必须 `Quit()`，否则留 WINWORD 进程，下次 `Documents.Open` 会卡住。
- 只装 WPS：`Word.Application` 换成 `KWPS.Application`，同样 `ExportAsFixedFormat($dst, 17)`。
- 宿主能起子进程时也可用脚本：`python <技能目录>/scripts/export_pdf.py "<N 目录>" "*第<W>周.docx"` → `PDF_OK=<路径>`；它自己写带 BOM 的 .ps1 并用 `powershell -ExecutionPolicy Bypass -File` 跑，依次试 Word → WPS → LibreOffice，`PRC_EXPORT=word|wps|soffice` 可强制。

**渲染每页 + 逐张看**：

```bash
python <技能目录>/scripts/proof.py "<N 目录>" "*第<W>周.docx" --dpi 110 --render "<N>/_work/pages"
```

- 它取目录里最新的 docx、找同名 PDF（没有就报「没有对应的 PDF，先跑 export_pdf.py」），用 pymupdf 把每页存成 `p01.png`…，并逐页报字数 / 图数 / OK 或 WARN。也可手工渲：

```python
import pymupdf
d = pymupdf.open(pdf)
for i, pg in enumerate(d, 1):
    pg.get_pixmap(dpi=110).save("4/_work/pages/p%02d.png" % i)   # 110 看版面，170 看细节
```

- **每张 `read_image`，不抽查**：图被分页切断、图注单独跑到次页页首、整页/半页空白、公式变方框、表格压扁，这些脚本查不出来。
- WARN 含义：`疑似空白页（<50 字无图）` / `文字极少（<50 字）` / `图片贴边/溢出` / `图注在本页但本页无图` / `docx 段落在 PDF 里找不到`。文字极少的页先看图，再决定删分页或放大图，**不要为了消 WARN 删内容**。
- 通过判据：表头页数与交付页数一致，逐页 OK，结尾 `结论：无异常`（退出码 0）。

**页数与图宽靠参数调，不改正文凑页数**：

- 页数由板块标题的分页开关控制（`heading(..., new_page=True)` / `body(..., new_page=True)`）。插原始记录照片后常见 11 → 12 页、末页只剩几十字：去掉一个 `new_page=True`，或把照片放大到接近满版心。
- 图宽别写死，用版心宽度算：`w = F.text_width_cm(doc)`（A4 约 17.2 cm），整页原始记录取 `0.65 * w`～`1.0 * w`——低于 65% 字看不清，也容易被判留白过多。
- 每次改脚本都要 重建 docx → 重导 PDF → 重跑 proof.py；docx 比 PDF 新时，校对的是旧 PDF。

## 对照与写入记忆

用于「统计几份报告的字数 / 图表编号」和「把同学报告对比写进 memory」。

1. **原始文件不动**：同学的 PDF 复制到 `_work/ext_compare/`，自己的读 `N/`；文本、页面图、统计全部落 `_work/ext_compare/`。
2. **抽文本**：`pymupdf` 逐页 `get_text()`，每页前插 `===== [PAGE n/m] =====` 方便切分；手写原始记录没有文字层，只能渲图看。
3. **统计脚本落盘再读**（见「Windows 控制台」）。统计前先剥页标记，否则字数被分页头污染：

```python
t = re.sub(r"=+ \[PAGE \d+/\d+\] =+", "", t)
```

4. **四类指标一起报**：
   - 页数；
   - 板块字数：预习综述、二、实验操作与原始数据、三、结果与分析；
   - 图/表编号序列：列出**重号**和**跳号**，并指出各在哪一节（例：预习 图1–3、结果又从 图1 起；表号从表 6 起、表 1–5 全篇不存在）；
   - 图像规格：`page.get_image_info()` 的 `bbox` 换算占页宽/页高比例，加内嵌像素（`width`×`height`），判断原始记录照片够不够大、够不够清。
5. **只做同实验对比**：同一实验、同一位教师的两份成对比（第 1 次示波器 vs 第 1 次示波器），不跨实验直接比分数；写明每次教师不同、单次分数有偶然性。
6. **写进 `memory.md`**：用 `edit` 在 `## 改动日志` 之前插入 `## 外部报告对比（软参考，非严格标准）`，开头写日期、对象、两轮教师，并注明「只作合规自检清单，不作严格评分标准」；`memory.py` 只负责手改 diff，这类分析性小结直接编辑 `memory.md`。
7. **写完回读**该节，确认旧章节没被覆盖；给用户的结论用「客观差异表 + 一句判断」，不堆原始输出。

## 用户手改后：把差异记进 memory

用户说「我改过了」「记住了」时：

    python3 <技能目录>/scripts/memory.py diff "<根目录>" "<旧版.docx>" "<改后.docx>" --n 3 --kind prep

- 按段落 / 表格 / 图片比对，天然忽略 Word 重存噪声，只把实质改动写进 `memory.md` 的「改动日志」。
- 先加 `--no-write` 看一遍，确认是真改动再落盘；只是打开另存过就不用记。
- 记忆里出现新偏好（例如「图注不要写公式」），同时更新「文风约定」或「版式约定」小节，别只躺在日志里。
- 下份稿子按记忆写，同样的退改不犯第二次。

## 交付

给出 docx 与 PDF 路径；宿主支持附件就一并附上。正文：路径 + 页数 / 字数（预习报综述、重点、难点；正式报核心结果与不确定度）/ 数据缺口 / 需要确认的事（签字、缺测）。不超过 300 字。

## 交稿前验收

- 构建输出无 `WARN 公式未能转成 OMML`；`m:oMath` 段数与 `proof.py` 的「含公式段」一致。
- PDF 比 docx 新，页数与交付说明一致；`proof.py` 结尾 `结论：无异常`。
- 输入只有 `.pptx` 时：`N-预/_work/ppt_text.txt` 已按 slide 分段，正文措辞能指回具体 slide；docx 至少 1 张图来自 `ppt/media/` 或课件 PDF 裁剪。
- PDF 文本里没有「图」后直接换行接数字；docx 图号引用含 U+2060，图注行首仍是 `图N`，`proof.py` 图注数没少。
- `pages/` 每张 PNG 都 `read_image` 看过：没有断图、图注落单、整页空白、公式方框。
- 文件名学号、封面学号、`config.yaml` 三处一致；原始记录照片整页可读、未代签。
- 改过 `build.py` 的话，改动已写回脚本（不是手改 docx），重跑一遍仍出同样的稿。

## 注意

- `matplotlib` 在部分环境坏（`_ARRAY_API not found`），用 `reportlab` 画图，别修依赖。
- 不拿理论值或别人的数据补缺测；缺的数据在正文写明。
- 课件 / 讲义没给的量程、分辨率、仪器允差不要写进「现象」或「操作」；课程未规定时按常用值（卡尺 0.02 mm、米尺 0.5 mm、电桥 0.2 级、温度计 0.1 ℃）并在正文注明「按推荐」。
- 不代签、不 P 图；签字栏空着就提醒用户。
- 记忆和对照结论都是软参考：单次分数受教师影响，别写成硬指标。
- **本技能是辅助工具**：整理材料、辅助计算与排版，不替代实验操作与你的最终核对；交付时如实说明这一点，不要承诺「直接可交」。
