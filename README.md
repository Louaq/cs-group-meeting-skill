# cs-group-meeting-skill

把一篇 CS 论文 PDF 变成一份中文组会汇报 PPT——按照实验室 `汇报模板.pptx` 的观感，
和 `汇报模板2.pptx` 的版式语汇。

产出的是**一条六段式的论证线**，不是图片集：
文献信息 → 研究背景 → 本文贡献 → 研究方法 → 实验结果 → 结论与不足。
每一页图都必须同时带要点框和结果分析——只有一张图配一句图注的页面，构建会直接报错。

这是一个 [Claude Code skill](https://docs.claude.com/en/docs/claude-code/skills)。
日常用法不是手敲命令行，而是把论文丢给 Claude Code：

> 把这篇论文做成组会 PPT

技能会自动触发，读全文、抽图、写 spec、构建、逐页渲染检查。命令行只是它内部调用的工具，
也可以在调试时自己跑。

## 环境依赖

- Python 3.12+，装 `python-pptx`、`pdfplumber`、`Pillow`
- Poppler 的 `pdftoppm`（渲染 PDF 页面用；不在 PATH 里可以用 `--pdftoppm` 指定）
- PowerPoint（Windows COM）——只在最后渲染 PNG 自检时需要

## 快速上手

```powershell
# 1. 抽图：渲染每页 + 聚类矢量图元定位图片
& python scripts\extract_figures.py extract --pdf paper.pdf --workdir <scratch>

# 2. 看一遍 <scratch>\crops\ 里的每张裁剪图，错的按 PDF point 重裁
& python scripts\extract_figures.py recrop --workdir <scratch> --fig 9 --bbox 80 323 531 416

# 3. 写 <scratch>\deck.json（格式见 SKILL.md 的 Deck spec），然后构建
& python scripts\build_deck.py --spec <scratch>\deck.json `
    --template assets\汇报模板.pptx --out 组会汇报.pptx
```

N 页正文最终会变成 N + 4 + 3 页：四张目录分隔页，加封面 / 论文信息 / 谢谢大家。

## 目录结构

| 路径 | 作用 |
|---|---|
| `SKILL.md` | 技能本体：约束、工作流、deck spec、11 种页型、写作要求 |
| `scripts/extract_figures.py` | 抽图 / 重裁（`extract`、`recrop`、`crop` 三个子命令） |
| `scripts/build_deck.py` | 从 deck.json 构建 PPTX；含 `check_*` 系列构建期校验 |
| `scripts/omml.py` | `$latex$` → PowerPoint 原生公式（OMML） |
| `assets/汇报模板.pptx` | 观感来源，默认模板 |
| `assets/汇报模板2.pptx` | 版式语汇来源（图上红色标注、结果分析框） |
| `assets/content.pptx` | 目录分隔页的样式来源（`--content` 可覆盖） |
| `references/style-profile.md` | 从模板量出来的几何与配色，`build_deck.py` 已编码 |
| `agents/openai.yaml` | — |

## 几个容易踩的点

- **构建期校验会拦人，这是故意的。** `check_variety` 不允许研究方法 / 实验结果
  整段一个页型；`check_emphasis` 不允许实验结果一处红色都没有；`check_heads`
  不允许提示框标签超过约 8 个汉字。改的是内容，不是绕过校验。
- **模板里的身份信息不是你的。** logo、二维码、汇报人、主页、页脚日期默认走占位符。
  不填就留占位框——二维码图片扫出来仍然指向模板作者的链接，跟标题写什么无关。
- **数字不能编。** 页面上的每个指标都要在 PDF 里存在；期刊 / 分区 / 被引数量要现查，
  它们会变。
- **公式写 `$...$`，不要写 `L_rw`。** 后者不是论文里的记号，也打不开公式编辑器。
- **必须逐页渲染看过。** 文字裁切和断行只有渲染出来才看得见。

细节都在 [SKILL.md](SKILL.md)。
