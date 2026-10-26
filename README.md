# physics-report-cc

> 给 [OpenClaw](https://openclaw.ai) 用的浙大大学物理实验技能：**预习报告 + 正式实验报告**，从课件到可交的 Word/PDF 一条龙。

写一份能交的实验报告，麻烦的从来不是写字，是这些：课表第几次对应第几周、封面字段每年都在变、综述卡 500 字、原始记录是手写照片、不确定度要乘灵敏系数、Word 导 PDF 还会挑中文路径的刺。

这个技能把这些固定动作沉淀成流程，AI 负责读材料、算数据、画图、排版，你负责签字和交作业。

## 它做什么

| 场景 | 输入 | 输出 |
| --- | --- | --- |
| 预习报告 | 实验手册 / 课件 PDF、课表、模板 | `N-预/<姓名> <学号> <实验名> 第W周.docx / .pdf` |
| 正式报告 | 原始记录照片、思考题、预习稿、模板 | `N/<同名>.docx / .pdf` |

- 预习：按实验数量判定 500 / 800 字上限，取 2–3 张原理图，按「现象 → 原理与方法 → ①②③」组织正文。
- 正式：原始记录读数据 → `calc.py` 出 `calc.json` → reportlab 出图 → 填模板 → Word 导 PDF → 逐页验版。
- 两步都做同一件事：**用上一份成品当底稿**，只换封面字段和正文，版式不走样。

## 不做什么

- **不自动爬取课件**。学习通的资料请自己下好，放进对应文件夹；技能只消费本地文件。
- 不代签、不 P 图、不拿理论值补缺测。
- 不打包任何学校模板、课件或学生报告（版权与隐私原因，见 `NOTICE`）。

## 环境要求

- OpenClaw（技能目录 `~/.openclaw/skills/`）
- Python 3.10+
- Windows + Microsoft Word（导出 PDF 走 Word COM；WSL 里调用 Windows 侧 PowerShell）
- Python 包：`python-docx`、`PyMuPDF`、`Pillow`、`reportlab`

```bash
pip install -r requirements.txt
```

## 安装

```bash
git clone https://github.com/Lineus/physics-report-cc.git \
  ~/.openclaw/skills/physics-report-cc
```

装完在 OpenClaw 里直接说「写第 3 次实验的预习报告」即可触发。

## 目录约定

把下面这个目录当成你的实验报告根目录（路径随意，技能会问一次）：

```
我的物理实验/
  config.yaml              # 从 config.example.yaml 复制
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
- 前几次的报告放不放都行；放了会被当作底稿和文风参考，没放就用模板。
- `N` 是第几次实验，`W` 是课表上的周次，**不是同一个数**。

## 配置

```bash
cp config.example.yaml 我的物理实验/config.yaml
```

至少填 `student.name`、`student.id`（10 位）、两个模板路径、课表路径。课表读不出来时用 `schedule` 手填。

## 用法

**预习报告**

> 写第 3 次实验的预习报告，根目录是 `D:\物理实验`

技能会：读课表定周次 / 教师 → 检查 `3-预/` 里的课件 → 取原理图 → 按 500 字上限写 → 导出 PDF → 报字数。

**正式报告**

> 这是第 3 次实验的原始记录和思考题，帮我写正式报告

照片可以直接发，也可以放进 `3/`。技能会先算数据并打印自检（实测 / 理论 / 相对误差），再出图、填模板、导 PDF、逐页验版。

## 脚本

| 脚本 | 用途 |
| --- | --- |
| `scripts/docx_fill.py` | 按锚点填模板、中文字体、真上下标、插图 |
| `scripts/pdf_extract_figs.py` | 从课件 PDF 取原图（位图 / 矢量） |
| `scripts/chart.py` | reportlab 画坐标图 + 误差棒（替代坏掉的 matplotlib） |
| `scripts/export_pdf.py` | Word COM 导 PDF，处理中文路径 / BOM / 锁文件 |
| `scripts/plan.py` | 读 config.yaml，算目录 / 文件名 / 周次 / 模板 |
| `scripts/check_env.py` | 检查依赖与 Word 是否可用 |

## 常见问题

**Q：为什么不用 matplotlib？**
某些环境下系统包的 matplotlib 与 numpy 版本对不上，`import` 就报 `_ARRAY_API not found`。技能直接用 reportlab 矢量绘制，再转 PNG，反而更清晰。

**Q：导 PDF 报「文档可能损坏」？**
PowerShell 5.1 按 ANSI 读脚本，中文路径会乱码。`export_pdf.py` 已经处理：脚本写 BOM、路径在脚本内用 `Get-ChildItem` 解析。别自己另写 ps1 传中文参数。

**Q：生成的 PDF 少了一页 / 多了空白页？**
九成是分页符加错了位置。只加在板块标题段；图片段设 `keep_with_next`，图注段不设。

**Q：学号写错？**
技能只从 `config.yaml` 取学号，并校验 10 位。旧稿里的 9 位号不会被沿用。

## 隐私与学术规范

- `config.yaml` 含姓名学号，已在 `.gitignore` 里；不要提交。
- 课表、课件、原始记录、成品报告都不要提交到公开仓库。
- 本技能用于整理**本人已选课程**的材料、辅助计算与排版；请遵守所在学校的学术规范，提交前自行核对内容并签字。

## License

MIT（代码）。课件、教材、模板的版权归各自权利人，见 `NOTICE`。
