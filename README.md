# physics-report-cc

> 浙大大学物理实验技能：**预习报告 + 正式实验报告**，从课件到可交的 Word/PDF 一条龙。

写一份漂亮的实验报告，麻烦的从来不是做好实验，是这些：把原始数据转成电子表格、写出LaTeX公式、逐项计算、画图、分析。

这个技能把这些固定动作沉淀成流程，Agent 负责读材料、算数据、画图、排版，你负责给出原始数据、最终修改和提交作业。在这个循环中，Agent 会根据已改好的报告变得越来越懂你。

## 它做什么

| 场景 | 输入 | 输出 |
| --- | --- | --- |
| 预习报告 | 实验手册 / 课件 PDF、课表、模板 | `N-预/<姓名> <学号> <实验名> 第W周.docx / .pdf` |
| 正式报告 | 原始记录照片、思考题、预习稿、模板 | `N/<同名>.docx / .pdf` |

- 预习：按实验数量判定 500 / 800 字上限，取 2–3 张原理图，按「现象 → 原理与方法 → ①②③」组织正文。
- 正式：原始记录读数据 → `calc.py` 出 `calc.json` → reportlab 出图 → 填模板 → Word 导 PDF → 逐页验版。
- 两步都做同一件事：**用模板当底稿，前面所有作参考**，只换封面字段和正文，版式不走样。
- 公式一律用 LaTeX（`$...$`）写，脚本转成 **Word 原生公式**：可编辑、可改字号，不是图片也不是拼出来的上下标。
- Agent **有视觉能力时逐页审版**：图有没有超版心、被分页切断、图注跑页首、公式变方框——只看提取的文字发现不了这些。

## 不做什么

- **不自动爬取课件**。学习通的资料请自己下好，放进对应文件夹；技能只消费本地文件。
- 不代签、不 P 图、不拿理论值补缺测。
- **它是辅助工具，不是代写工具**：数据必须来自真实的原始记录，结论与提交前的内容要你自己核对。
- 不打包任何学校模板、课件或学生报告（版权与隐私原因，见 `NOTICE`）。

## 环境要求

- 任何有**文件读写 + shell** 的 Agent（纯聊天宿主不适用，建议用于带有图片能力的Agent）
- Python 3.10+
- 导出 PDF：Windows + Word 最准（走 COM）；没有 Word 就装 LibreOffice，脚本自动退
- Python 包：`python-docx`、`PyMuPDF`、`Pillow`、`reportlab`、`PyYAML`、`latex2mathml`

**公式渲染**需要 Office 自带的 `MML2OMML.XSL`（Word 2007 起都有）。找不到时会退回
`_{}`/`^{}` 写法并打印 `WARN`。

```bash
pip install -r requirements.txt
```

## 安装

克隆到任意位置，再让 Agent 知道它在哪即可。

```bash
git clone https://github.com/Lineus/physics-report-cc.git
```

| 宿主 | 放 / 指向哪里 |
| --- | --- |
| OpenClaw | `git clone ... ~/.openclaw/skills/physics-report-cc`，按 description 自动触发 |
| Claude Code | 复制到 `.claude/skills/physics-report-cc/`（项目级）或 `~/.claude/skills/` |
| Codex CLI / Gemini CLI | 在项目 `AGENTS.md` / `GEMINI.md` 里写一行：写报告前先读 `<路径>/SKILL.md` |
| Cursor / Cline / Windsurf | 同上，或在规则文件里引用 `SKILL.md` |
| 直接调 API | 把 `SKILL.md` + 对应 `references/*.md` 拼进 system prompt |

仓库根目录还有一个 `AGENTS.md`：即使 Agent 不会自动加载技能，只要它读了项目根目录的说明，
也会在第一次被问到时引导你完成配置。不放心的宿主，把这一行加进你自己的 `AGENTS.md`：

```
配置/使用物理实验报告技能时，先读 <克隆路径>/AGENTS.md
```

### 首次配置：让 Agent 引导你做完

把项目路径告诉 Agent，说「帮我配置好」。它会跑：

```bash
python3 scripts/setup.py "<你的实验报告根目录>"
```

然后：

1. 建 `模板/`，**按 12 次实验预建 `1/`、`1-预/` … `12/`、`12-预/`**，生成 `config.yaml` 与 `memory.md`；
2. 扫你已有的报告，把实验清单登记进 `memory.md`；
3. 检查依赖、公式渲染、Word / LibreOffice；
4. 打印 `NEED=…`——把缺的（姓名、10 位学号、课表、两份模板）**一次问完**，
   带参数重跑 `setup.py`，直到 `READY=1`。

已有的 `config.yaml` / `memory.md` 不会被覆盖，重复跑是安全的。

配好后对 Agent 说「写第 n 次实验的预习报告」即可触发。宿主差异、视觉能力降级方案见
[`references/agents.md`](references/agents.md)。

## 目录约定

下面这个目录会成为你的实验报告根目录（路径随意，技能会问一次）：

```
我的物理实验/
  config.yaml              # 从 config.example.yaml 复制
  memory.md                # 项目记忆，Agent 维护（见下）
  课表.png                  # 课程时间表截图
  模板/
    预习报告模板.docx
    实验报告模板.docx
  1-预/
    实验手册.pdf
    ppt.pdf
  1/
    原始记录.jpg
    思考题.jpg
  2-预/
  ...
```

- 目录名默认 `{n}-预` / `{n}`，可在 `config.yaml` 改成 `{n}预`。
- 文件夹已经按 12 次建好：**旧报告按编号放进去**（第 3 次预习 → `3-预/`，正式 → `3/`），
  Agent 会扫这些目录学结构；不放也能从模板起草。
- **永远从模板起草**，前面的报告只作文风参考，不再当底稿。
- `memory.md` 是 Agent 的长期记忆：版式偏好、文风、你每次手改的差异，都记在里面。
- `N` 是第几次实验，`W` 是课表上的周次，不是同一个数。

## 配置

```bash
cp config.example.yaml 我的物理实验/config.yaml
```

至少填 `student.name`、`student.id`（10 位）、两个模板路径、课表路径。课表读不出来时用 `schedule` 手填。

## 用法

**预习报告**

> 写第 n 次实验的预习报告

技能会：读课表定周次 / 教师 → 检查 `3-预/` 里的课件 → 取原理图 → 按 500 字上限写 → 导出 PDF → 报字数。

**正式报告**

> 这是第 n 次实验的原始记录和思考题，帮我写正式报告

照片可以直接发，也可以放进 `3/`。技能会先算数据并打印自检（实测 / 理论 / 相对误差），再出图、填模板、导 PDF、逐页验版。

## Agent 有眼睛，就让它逐页审

导 PDF 之后，有视觉能力的 Agent 会把每页渲成图**逐页看**，而不是只提取文字：

- 图有没有超出页边距、被分页从中间切断
- 图注是否紧跟其图，有没有单独跑到次页页首
- 有没有整页空白、半页空白、封面字段挤压换行
- 公式有没有变方框、问号、半截

这些在 `python-docx` 里都查不出来——Word 的分页逻辑它复现不了。所以「导 PDF + 看图」
是这个技能唯一可靠的验版方式。

先跑自动校对，它会把每页渲成 PNG：

```bash
python3 scripts/proof.py "<根目录>/<N>" "*第<W>周.docx" --render <N>/_work/pages
```

自动报空白页、图溢出、图注落单、docx 段落在 PDF 里缺失、三处字数；退出码 1 表示有 WARN。
剩下「图是不是从中间被切断」「公式有没有变方框」只能靠看 `pages/` 里的图。
没有视觉能力时，把 PDF 交给你过目即可。

## 项目记忆：越用越懂你

第一次写稿时 Agent 会执行 `memory.py init`，在根目录生成 `memory.md`，记下你的文风与版式偏好。

你手改完成品后，直接对 Agent 说「我改过了，记住」：

```bash
python3 scripts/memory.py diff "<根目录>" "<旧版.docx>" "<改后.docx>" --n 3 --kind prep
```

它按**段落 / 表格 / 图片**三个层级比对，Word 重存造成的 run 拆分、缩进微调不会误判成改动；
真正的内容变化会追加进 `memory.md` 的「改动日志」。下次写稿先读记忆，同样的退改不会犯第二次。

想让 Agent 记住别的事，也可以直接说，或：

```bash
python3 scripts/memory.py note "<根目录>" "图注不要带公式"
```

## 脚本

| 脚本 | 用途 |
| --- | --- |
| `scripts/docx_fill.py` | 按锚点填模板、中文字体、LaTeX 公式、插图 |
| `scripts/mathomml.py` | LaTeX → MathML → Word 原生公式（OMML） |
| `scripts/memory.py` | 项目记忆：init / show / diff / note |
| `scripts/pdf_extract_figs.py` | 从课件 PDF 取原图（位图 / 矢量） |
| `scripts/chart.py` | reportlab 画坐标图 + 误差棒（替代坏掉的 matplotlib） |
| `scripts/export_pdf.py` | 导 PDF：Word COM 优先，LibreOffice 兜底 |
| `scripts/plan.py` | 读 config.yaml，算目录 / 文件名 / 周次 / 模板 |
| `scripts/refs.py` | 扫旧稿：字数 / 图数 / 公式数 / 章节骨架，只作结构参考 |
| `scripts/proof.py` | 逐页校对：空白页、图溢出、图注落单、缺段、字数，并渲页面图 |
| `scripts/check_env.py` | 检查依赖与 Word 是否可用 |

## 隐私与学术规范

- `config.yaml` 含姓名学号，已在 `.gitignore` 里；不要提交。
- 课表、课件、原始记录、成品报告都不要提交到公开仓库。
- 本技能是**辅助工具**：整理本人已选课程的材料、辅助计算、画图与排版，不替代实验操作、数据分析判断和你的最终核对。
- 请遵守所在学校的学术规范；提交前自行核对内容、补签字，并对报告负责。

## License

MIT（代码）。课件、教材、模板的版权归各自权利人，见 `NOTICE`。