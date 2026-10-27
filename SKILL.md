---
name: "physics-report-cc"
description: "浙大物理实验报告技能（任何 agent 可用）：首次安装引导配置，之后写预习/正式报告——模板 + memory 出稿，LaTeX 转 Word 原生公式，导 PDF 逐页校对"
---

# 浙大物理实验报告（预习 + 正式）

面向**任何能跑 shell、能读写文件的 agent**：OpenClaw、Claude Code、Codex CLI、Gemini CLI、Cursor、Cline 等。
安装方式见 `README.md`；各宿主差异见 `references/agents.md`。

## 何时用

- **用户刚 clone 下来、说「装好了 / 帮我配置一下」，或什么都没说就把仓库地址给了你** → 先跑 `setup.py`，见下节
- 用户说「写第 N 次实验的预习报告」「下周实验的预习」
- 用户说「写第 N 次实验报告」「正式报告」，并给了原始记录 / 思考题照片
- 用户交回**手改过的成品**，要你记住改动 → 见「用户手改后」一节

## 先判分支

| 触发 | 分支 | 细读 |
| --- | --- | --- |
| 预习报告、`N-预/` | A 预习 | `references/prep.md` |
| 正式报告、原始记录、`N/` | B 正式 | `references/report.md` |

先读本页公共块，再读对应分支。**两份不要串线**：

- 500 字上限只属预习综述（单实验 500，双实验 800）；正式报告不按这个数压。
- 预习不碰数据处理；正式报告必须 `calc.py` 算、`figs.py` 出图，不用理论值冒充实测。
- 原始记录照片、思考题只出现在正式报告。
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

## 目录与配置

    <实验报告根目录>/
      config.yaml            # 姓名/学号/路径，含个人信息，勿提交
      memory.md              # 项目记忆，含个人信息，勿提交
      课表.png                # 课表截图；读不出来就用 config.schedule 手填
      模板/
        预习报告模板.docx
        实验报告模板.docx
      1-预/  实验手册.pdf  ppt.pdf     # 第 1 次预习，用户自己放
      1/     原始记录.jpg  思考题.jpg   # 第 1 次正式，用户自己放

- 目录名由 `config.yaml` 的 `dirs` 决定，默认 `{n}-预` / `{n}`；也支持 `{n}预`。
- 没有 `config.yaml`：从 `config.example.yaml` 复制，逐项问姓名、10 位学号、模板路径。学号必须 10 位，不沿用旧稿。
- 缺材料：先建目录，说清放什么，**不要编造课件内容**。
- 用户没指定根目录时问一次，并记进 memory.md。

## 命名与学号

- 文件名 = `{name} {id} <实验名称> 第{week}周.docx`，由 `config.yaml` 的 `filename` 生成。
- `N` 是第几次实验，`W` 是课表周次，**不是同一个数**（如 `3-预` = 第 5 周）。
- 交付前复核：文件名学号、封面学号、页数三处一致。

## 课表定字段

- 课表是图片时用宿主的多模态能力直接看；看不了就复制到临时目录再让宿主读。**读不到就问用户**，不要猜。
- 取实验名称、周次、日期、星期、上/下午、指导教师；实验名称照课表原文，并与课件首页核对。
- 封面逐项重填，不沿用上一份；读不出来就用 `config.yaml` 的 `schedule`。

## 公式：写 LaTeX，不拼上下标

正文用 `$...$` 包 LaTeX，脚本转成 **Word 原生公式**（可编辑、字号跟正文）：

- `$B=\frac{\mu_0 N I R^{2}}{2(R^{2}+x^{2})^{3/2}}$`
- `$u_c=\sqrt{u_A^{2}+u_B^{2}}$`

- 行内公式不必另起段；独立成行时整段只放一个 `$...$`。
- 同一符号前后统一（不要一会儿 `B_0` 一会儿 `B0`）。
- 不用 Unicode 上下标字符，也不把公式贴成图片。
- 没装 Office 时自动走 `scripts/omml.py` 的纯 Python 转换（WPS / Linux / macOS 都有原生公式）；只有两条都不可用才退回 `_{}`/`^{}` 并打 `WARN`，此时把 `WARN` 如实告诉用户。

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

平台差异集中在 `scripts/hostenv.py`（找 PowerShell、Windows↔WSL 路径互转、Office XSL 位置），
脚本里不要再自己判断 `os.name` 或拼 `/mnt/...`。

## 排版铁律

- 从**模板**起草：`open_base(模板, OUT)`，不动原件。模板确实缺失才退回上一版成品。
- 按段落**文本**定位锚点，不写死段号；删段先收集元素再删。
- 分页符只加在板块标题段；加在标题下第一段会插出整页空白。
- 图片段设 `keep_with_next=True`，图注段不设；两边都设会把整组图挤到下一页。
- 宋体 12pt、1.25 倍行距、首行缩进 2 字符；图注 10.5pt 居中；中文字体连 `w:eastAsia` 一起设。
- 中文长段里的百分号用 f-string，建完 `count('%%')` 断言为 0。
- 中间产物（脚本、json、png、docx 副本）进 `N/_work/`，成品放 `N/` 或 `N-预/`。

## 用户手改后：把差异记进 memory

用户说「我改过了」「记住了」时：

    python3 <技能目录>/scripts/memory.py diff "<根目录>" "<旧版.docx>" "<改后.docx>" --n 3 --kind prep

- 按段落 / 表格 / 图片比对，天然忽略 Word 重存噪声，只把实质改动写进 `memory.md` 的「改动日志」。
- 先加 `--no-write` 看一遍，确认是真改动再落盘；只是打开另存过就不用记。
- 记忆里出现新偏好（例如「图注不要写公式」），同时更新「文风约定」或「版式约定」小节，别只躺在日志里。
- 下份稿子按记忆写，同样的退改不犯第二次。

## 导出与验版

    python3 <技能目录>/scripts/export_pdf.py "<根目录>/<N 或 N-预>" "*第<W>周.docx"

- 依次尝试 Word COM → WPS COM（`KWPS.Application`）→ LibreOffice；只装 WPS 的机器走第二条。可用 `PRC_EXPORT=word|wps|soffice` 强制。
- 成功打印 `PDF_OK=<路径>`，失败退出码 1。
- 导出后跑逐页校对，它会渲好每页的 PNG：

        python3 <技能目录>/scripts/proof.py "<根目录>/<N 或 N-预>" "*第<W>周.docx" --render <N>/_work/pages

  它报：空白页、图溢出、图注落单、docx 段落在 PDF 里缺失、三处字数、含公式段数。
  退出码 1 表示有 WARN，逐条确认后再交付。
- **有视觉能力必须把 `pages/` 里的图逐页看，不是抽查**（校对脚本查不出断图位置和公式方框）；没有就说明「版面未自检」并把 PDF 交给用户（详见 `references/agents.md`）。
- 每次改文字都重导，别只看 docx。

## 交付

给出 docx 与 PDF 路径；宿主支持附件就一并附上。正文：路径 + 页数 / 字数（预习报综述、重点、难点；正式报核心结果与不确定度）/ 数据缺口 / 需要确认的事（签字、缺测）。不超过 300 字。

## 注意

- `matplotlib` 在部分环境坏（`_ARRAY_API not found`），用 `reportlab` 画图，别修依赖。
- 不拿理论值或别人的数据补缺测；缺的数据在正文写明。
- 不代签、不 P 图；签字栏空着就提醒用户。
- **本技能是辅助工具**：整理材料、辅助计算与排版，不替代实验操作与你的最终核对；交付时如实说明这一点，不要承诺「直接可交」。
