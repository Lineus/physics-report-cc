# 在不同 Agent 里使用

这个技能没有绑定任何一家 Agent。它只是「一堆 Markdown 规则 + 几个 Python 脚本」，
凡是有文件读写和 shell 的宿主都能跑。差别只在**怎么把技能喂给 Agent**、以及
**Agent 能不能看图**。

## 安装位置

| 宿主 | 放哪里 | 触发方式 |
| --- | --- | --- |
| OpenClaw | `~/.openclaw/skills/physics-report-cc/` | 自动按 description 匹配，说「写预习报告」即可 |
| Claude Code | 项目内 `.claude/skills/physics-report-cc/`，或 `~/.claude/skills/` | 技能名可被 `/` 唤起；也可让 Agent 直接读 `SKILL.md` |
| Codex CLI | 项目根 `AGENTS.md` 里加一行「写实验报告前先读 `<路径>/SKILL.md`」 | 会话中直接说需求 |
| Gemini CLI | 同上，写进 `GEMINI.md` | 同上 |
| Cursor / Cline / Windsurf | 仓库内 `.cursor/rules/` 或工作区 `AGENTS.md` 引用 `SKILL.md` | 同上 |
| 任意 API 调用 | 把 `SKILL.md` + 对应 `references/*.md` 拼进 system prompt | 自己控制 |

**没有技能加载机制的宿主，最省事的做法**：在项目根目录放一个 `AGENTS.md`：

```markdown
写物理实验报告/预习报告时，先完整读取 ./physics-report-cc/SKILL.md，
再按其中「先判分支」读取对应的 references 文件。
```

## 能力差异与降级方案

| 能力 | 有 | 没有时怎么办 |
| --- | --- | --- |
| 看图（多模态） | 课表、原始记录、成品 PDF 都可以直接看 | 课表读不出来 → 用 `config.yaml` 的 `schedule` 手填；原始记录 → 让用户口述或贴文本表格；成品验收 → 让用户自己看 PDF |
| shell | 跑 `plan.py` / `export_pdf.py` | 没有 shell 的纯聊天宿主不适合本技能，至少需要文件读写 |
| Word COM | 导 PDF 版式最准 | 装 WPS 就走 `KWPS.Application`；都没有用 LibreOffice，脚本自动退 |
| 联网 | 不需要 | —— 本技能从不联网 |

## 视觉能力：逐页审查

**有视觉能力的 Agent 必须做这一步**，这是本技能质量的主要来源之一：

1. 导出 PDF 后按 110 dpi 渲染成 PNG（`fitz`）。
2. **每一页都看一遍**，不是抽查：
   - 图有没有超出页边距、被分页切断
   - 图注是否紧跟其图，有没有单独跑到次页页首
   - 有没有整页空白、半页空白
   - 公式有没有变成方框 / 问号 / 半截
   - 封面字段有没有换行挤压、下划线残留
3. 发现问题改脚本重跑，不要手动改 docx（下次重跑会覆盖）。
4. 没视觉能力时：用 `fitz` 提取每页文字量做粗筛（`<50` 字的页大概率是空白页），
   并把 PDF 交给用户，明确说「版面没法自检，请你过一眼」。

## 常见宿主坑

- **Claude Code**：默认工作目录是仓库根，脚本里的相对路径要基于 `$PWD`，别假设技能目录。
- **Codex CLI**：沙箱默认只读，跑 `export_pdf.py` 前需要允许写 `_work/` 与目标目录。
- **Windows 原生 Agent**（非 WSL）：`export_pdf.py` 的 `to_win()` 会原样返回 Windows 路径，Word / WPS COM 直接可用；LibreOffice 也走 `.exe`。
- **只装 WPS**：不用装 Office。导出用 `KWPS.Application`，公式走 `scripts/omml.py`；`check_env.py` 会确认这两条链路。
- **macOS**：Word COM 不存在，走 LibreOffice；`MML2OMML.XSL` 在 `Microsoft Word.app/Contents/Resources/`。
- **Windows 上的 Python**：用 `py -3` 或 `python`；`python3` 在不少机器上只是商店占位符（运行会弹安装页）。
- **路径写法**：Windows 用 `E:\物理实验\3`，WSL 用 `/mnt/e/物理实验/3`；`hostenv.to_win()` / `to_local()` 会自动换。
