---
name: cs-group-meeting-skill
description: "Build a sectioned Chinese group-meeting (组会/文献汇报) PowerPoint deck from a CS paper PDF, using the bundled 汇报模板.pptx. Use when the user provides a paper PDF and wants a full 文献信息→研究背景→本文贡献→研究方法→实验结果→结论与不足 report deck rather than a figure dump: the skill reads the whole paper, summarises it into Chinese talking points, and lays each slide out with 要点框 + 图 + 结果分析 so no slide is ever just one picture. Triggers on 组会PPT, 文献汇报, 论文汇报, paper reading PPT, 把这篇论文做成组会PPT."
---

# CS Group Meeting Skill

Turns a paper PDF into a Chinese group-meeting deck that follows the lab's
`汇报模板.pptx` for its look and `汇报模板2.pptx` for its layout vocabulary: a
six-section narrative, not a figure-by-figure walkthrough.

## Non-negotiables

1. **Six sections, in order.** 文献信息 → 研究背景 → 本文贡献 → 研究方法 →
   实验结果 → 结论与不足. Every slide declares its `section`.
   The **目录 divider** (`content.pptx` style) that opens each part of the talk
   lists **four** entries, not six — the 目录 is the argument's arc, not a
   section index:

   | 目录 entry | sections |
   |---|---|
   | 01 研究背景与贡献 | 研究背景, 本文贡献 |
   | 02 研究方法 | 研究方法 |
   | 03 实验结果 | 实验结果 |
   | 04 结论与不足 | 结论与不足 |

   文献信息 is front matter and gets **no divider** — it is on screen before the
   talk starts. 研究背景 and 本文贡献 are one beat («here is the gap, here is what
   we did about it») and share one divider; they keep their own slides and
   titles. Edit `TOC_GROUPS` / `TOC_SKIP` in `build_deck.py` to regroup;
   `check_toc_coverage` fails if a section belongs to neither.
   There is **no nav bar** — the top of a content slide is content.
2. **No slide is one figure plus a caption.** Every figure slide carries a
   要点框 (bullets) *and* a 结果分析 callout stating what the figure proves.
   `build_deck.py` refuses to build a figure slide with fewer than two text
   blocks. If you have nothing to say beyond the caption, the figure does not
   deserve a slide.
3. **Vary the layout — especially 研究方法 and 实验结果.** 「介绍文字 + 图 + 一句
   总结」 repeated eight times is the failure mode this skill exists to avoid.
   Pick the layout that fits the *content*: a pipeline wants `process`, a
   comparison table wants `panel`, three findings want `figure_cards`, two
   results that answer different questions want `dual`. `check_variety` fails
   the build if either section is one type throughout, or four alike in a row.
3. **Summarise the paper, not the captions.** Read the abstract, intro, method
   and experiment sections. Talking points come from the text — the problem,
   the insight, the losses, the datasets, the numbers, the ablations, the
   limitations. Translating figure captions is not enough.
4. **Never invent numbers.** Every metric on a slide must exist in the PDF.
   Verify 期刊/分区/被引数量 from a current source; they change over time.
5. **实验结果 marks the best result in red.** `!!DSC 91.19%!!` renders bold red.
   Mark the winning number in every comparison — and only the winning number;
   red on everything is red on nothing. `check_emphasis` fails a 实验结果 section
   with no red at all.
6. **Math is an equation, not characters.** Write inline math as `$...$` LaTeX;
   the builder converts it to a real PowerPoint equation. Never type `L_rw`,
   `X^2` or `A_1` as literal text — that is not what the paper's notation looks
   like and it does not open in the equation editor.
7. **The template's identity is not the presenter's.** 学校 logo、个人主页二维码、
   汇报人、个人主页网址、页脚日期 default to placeholders / today. Fill them from
   `meta` only when the user supplies their own; never ship the template's.
8. **A 提示框 head never wraps.** 结果分析 / 关键设计 / 标注比例 are labels — every
   head shape is sized from its text and has wrapping off, so keep heads to a
   few characters. `check_heads` rejects anything over ~8 CJK chars.
9. **The 左上角标题 is a takeaway, not a label.** Every content slide's title is a
   short, declarative statement of what the slide *says* — what the module does,
   what the result shows — not a bare noun naming the thing on screen. Write
   「CFE：用模态差异跨引导加权本模态特征」, not 「CFE 相关特征探索单元」;
   「消融实验：CVBM 与一致性损失各自有效」, not 「消融实验」. Keep it to one line
   (≈16 CJK chars; the box auto-shrinks a longer one but a title that has to
   shrink is a title trying to be a sentence). The section-opener slides
   (研究背景 / 本文贡献 / 结论与不足) may keep their section name.
10. **Figure and text align on one axis.** In a side-by-side layout the figure
    must line up with the text beside it, not float half a slide above or below
    it. `figure_analysis` centres both on the same vertical axis for you; when
    you place figures yourself, match their vertical centre to the text column's.
11. **图多字少 — the callout is one centred line, not a paragraph.** A 结果分析 /
    分析结果 / 小结 callout states the single takeaway (with its key number) in
    about one line; it renders centred. A multi-sentence callout wraps, balloons
    the box and shrinks the figure — exactly backwards. `check_analysis_terse`
    caps each callout line (~32 字 in a `dual` column, ~34 elsewhere) and the
    number of lines (2 in `dual`, 3 otherwise). Push detail into the figure or a
    要点; the figure should dominate the slide, so give it the width and height
    the terse callout frees up.

## Workflow

1. Confirm the input PDF, the output PPTX path, and the template
   (`assets/汇报模板.pptx` unless the user supplies their own).
2. Make a scratch dir outside the user's folder, e.g.
   `<scratch>/cs-group-meeting/<paper-slug>`.
3. Extract figures:

```powershell
& "<python>" "<skill-dir>\scripts\extract_figures.py" extract `
  --pdf "<paper.pdf>" --workdir "<scratch>"
```

4. **Look at every crop in `<scratch>/crops/`.** The extractor clusters vector
   drawing objects above each caption, which handles the vector figures that
   raster extractors miss — but it cannot separate a figure from a table that
   sits closer to it than the figure's own internal gaps, and it can swallow a
   running header. Measure those off `<scratch>/pages/page-NN.png` and fix:

```powershell
& "<python>" "<skill-dir>\scripts\extract_figures.py" recrop `
  --workdir "<scratch>" --fig 9 --bbox 80 323 531 416      # PDF points
```

5. Read the paper end to end. Then write the deck spec (below), covering the
   whole argument: what problem, why existing methods fail, what is proposed,
   how it works, what the experiments show, what is left open.
6. Build, then **render and inspect every slide** (see Verify).

```powershell
& "<python>" "<skill-dir>\scripts\build_deck.py" `
  --spec "<scratch>\deck.json" `
  --template "<skill-dir>\assets\汇报模板.pptx" `
  --out "<out.pptx>"
```

The 目录 dividers come from `assets\content.pptx` (override with `--content`).
A deck of N content slides ends up with N + 4 + 3 slides: one divider per 目录
entry plus 封面 / 论文信息 / 谢谢大家.

## Deck spec

```json
{
  "meta": {
    "topic_zh": "半监督医学图像分割",
    "title_en": "Paper Title",
    "authors": "A, B, C",
    "venue_line": "IEEE Transactions on ..., 2025",
    "url": "https://github.com/...",
    "cover_crop": "<scratch>/cover_top.png",
    "info": {"发表期刊": "IEEE TIP", "发表时间": "2025",
             "分       区": "中科院一区 / CCF A", "被引数量": "12（谷歌学术）"},

    "presenter": "汇报人姓名",
    "school": "学校 / 学院名称",
    "homepage": "https://...",
    "logo": "<path>/logo.png",
    "qr": "<path>/homepage-qr.png",
    "paper_qr": "<path>/paper-qr.png",
    "date": "2026年7月16日"
  },
  "slides": [
    {"type": "info", "title": "文献信息", "bullets": ["一句话结论", "..."]},
    {"type": "bullets", "section": "研究背景", "title": "半监督分割缺的是一致的伪标签",
     "bullets": ["**问题**：...", "..."], "figures": ["fig01.png"],
     "notes": ["Fig. 1 动机"]},
    {"type": "cards", "section": "本文贡献", "title": "本文贡献",
     "cards": [{"lead": "CVBM", "text": "...", "points": ["..."]}]},
    {"type": "figure_analysis", "section": "研究方法", "title": "CVBM 用体素置信度对齐双分支预测",
     "figures": ["fig07.png"], "figure_side": "right", "figure_width": 7.6,
     "notes": ["Fig. 7 CVBM 框架"],
     "bullets": ["..."], "analysis": ["..."], "analysis_head": "关键设计"},
    {"type": "numbered", "section": "结论与不足", "title": "结论与不足",
     "items": [{"title": "...", "desc": "..."}]}
  ]
}
```

### Slide types — pick by what the content *is*

| type | shape | reach for it when |
|---|---|---|
| `info` | 首页图 + 信息表 + 要点 | 文献信息 only |
| `bullets` | 要点框 + 0–3 图 | a list with no single figure to anchor |
| `cards` | 2–4 并列卡片 | the contributions; parallel claims |
| `figure_analysis` | 图 + 要点框 + 结果分析 | one figure worth three bullets. **The default — so stop and ask whether another type fits better** |
| `process` | 图 + 底部流程条 | a pipeline / 两阶段训练 — name the steps |
| `tree` | 一句论断 → 箭头 → 并列分支 | 现有方法各有什么短板 |
| `figure_cards` | 论断 + 每卡「图 + 小标题 + 说明」 | 2–3 findings, one figure each |
| `panel` | 大图 + 右侧编号分析面板 | a table/big figure with 4–5 readings |
| `dual` | 两图并列，各配各的分析 | two results answering different questions |
| `proscons` | 论断 + 结论/局限 双色条 | 结论与不足 |
| `numbered` | 编号结论框 | a plain numbered list |

Required keys per type are in `REQUIRED` in `build_deck.py`; a missing one fails
the build with a named error.

`figures` entries are a path, or `{"path": ..., "annotations": [...]}` for the
red on-figure notes 汇报模板2 uses. Each annotation is
`{"at": [fx, fy], "text": "...", "w": 2.4}` where `at` is a fraction of the
*placed* figure. **Aim them at the figure's margins and re-render to check** —
a note dropped on a dense panel covers the thing it points at, which is worse
than no note. Prefer few.

Cover / paper banner / 谢谢大家 slides come from the template and are filled
from `meta`.

The last seven `meta` fields are the presenter's own identity, and every one is
optional. Ask the user for them; if they do not say, leave them out and the
slide shows a dashed placeholder box saying what to drop in (`date` defaults to
today). Do **not** fall back to the template's values — its logo, QR codes,
homepage and date belong to whoever built it, and the QR images still resolve
to *their* links no matter what the caption says, so an unsupplied asset is
replaced, not kept.

### Markup inside any string

- `**text**` → bold green. Mark the key term in each bullet, as the template does.
- `!!text!!` → bold red. The best result in a comparison (实验结果), or a red
  note pinned on a figure. Not for general emphasis — that is what `**` is for.
- `$latex$` → a real PowerPoint equation (`omml.py`). Works in bullets, 分析,
  卡片, 图注 and slide titles. Supported: `_` `^`, `\frac`, `\sqrt`,
  `\hat/\bar/\tilde/\vec`, `\mathcal/\mathbb/\mathrm/\text/\mathbf`, greek,
  and the usual operators; anything else fails the build rather than rendering
  wrong. Match the paper's own notation — `$\mathcal{L}_{\mathrm{rw}}$`, not
  `L_rw`. Bold and math cannot nest; bold the words around the equation.

## Writing the talking points

Give every content slide a **declarative title** (non-negotiable 9): 研究方法
slides say what the module does (「ACF 用跨模态差异重标定特征」), 实验结果 slides
say what the numbers show (「BraTS 上 DSC 领先次优方法 1.8%」). A title that just
names the box on screen wastes the one line the audience reads first.

- **研究背景**: the technical problem, why current SOTA falls short, the paper's
  observation. 2–4 bullets, each with one **green key term**.
- **本文贡献**: one card per contribution, `lead` = the named component.
- **研究方法**: one slide per component (总体框架 / 每个模块或损失). Say what goes
  in, what comes out, what the loss optimises. Cite the equation number
  (`Eq.(17)`) next to the claim so the audience can find it. Give the pipeline
  a `process` strip and the loss/framework a `panel` — they are different kinds
  of thing and should not look identical.
- **实验结果**: 数据集与指标 first, then 对比实验, 可视化, 消融. Each 分析 is one
  centred line naming the number that settles it (e.g. `!!DSC 91.19%!! vs BCP
  89.62%`) with the winner in red — not a paragraph restating the table. A table
  wants `panel`, two datasets want `figure_cards`, two visualisations want
  `dual`. Blow the figure up: 图多字少 means the reader looks at the picture, not
  a wall of text beside it.
- **结论与不足**: contributions plus the limitations the paper admits.

Keep each bullet under ~40 Chinese characters; the 要点框 grows with the bullet
count, not with their length. Preserve exact values, units, dataset/method names.
Do not translate labels drawn inside a figure.

## Verify

Rendering is the only way to catch clipped text and bad line breaks:

```powershell
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open("<out.pptx>", $true, $false, $false)
$pres.SaveCopyAs("<scratch>\render\s", 18)   # 18 = ppSaveAsPNG
$pres.Close(); $app.Quit()
```

Then read every PNG: text inside its box, figures neither stretched nor
clipped, 共N页 matching the deck, no CJK punctuation at a line start, and no
slide that is just a picture. A 图注 sits right under its figure (the builder
hangs it off the image's real bottom edge, not the slot's), so a caption
stranded in whitespace means the wrong figure or note pairing. On side-by-side
slides check the figure and the text column share a vertical centre — neither
should float above or below the other. Confirm each title reads as a takeaway sentence, not a bare label, and
sits on one line. Check the red
annotations landed in the figure's whitespace and not across its panels, and
flip through 研究方法 / 实验结果 as a run — if the pages blur together, the
layout is still too uniform whatever `check_variety` says.

## Script notes

- `extract_figures.py` needs `pdfplumber`, `Pillow`, and Poppler's `pdftoppm`.
- `build_deck.py` needs `python-pptx` and `Pillow`. It clones the nav bar,
  footer and title icon from the template per section, rewrites the master's
  hard-coded `共22页`, restamps the footer's hard-coded date (`set_date`, before
  the footer is harvested as furniture — otherwise the template's date ships on
  every slide), and tags every run `zh-CN` so PowerPoint applies CJK
  line-breaking (without that, `。` and `，` strand at line starts).
- `omml.py` turns `$latex$` into the `mc:AlternateContent` / `a14:m` block that
  PowerPoint's equation editor reads, with a plain-text fallback. Bad LaTeX
  fails `check_math` before any slide is built.
- 图注 renders **black** — it is caption text, not an accent colour.
- Template geometry and palette: `references/style-profile.md`.
