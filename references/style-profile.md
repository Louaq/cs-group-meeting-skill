# 汇报模板 Style Profile

Measured from `assets/汇报模板.pptx` (look) and `汇报模板2.pptx` (layout
vocabulary). Use these numbers when placing shapes; `build_deck.py` already
encodes them.

## Deck geometry

- 16:9, 13.333 × 7.5 in.
- Font: **微软雅黑** for all text, CJK and Latin.
- Palette: primary green `007A49`, deep green `005C37`, ink `333333`,
  wash `F2F6F4`, white.
- Sizes: body 20pt, slide title 24pt, card text 18pt, figure note 15pt,
  cover title 48pt, section numeral 54pt.
- Red `FF0000` is 汇报模板2's emphasis: on-figure annotations and the winning
  number in 实验结果. Bold, and sparing — red on every number emphasises nothing.

## Recurring furniture (cloned per slide, never redrawn)

- **The nav bar is not used.** The template carries one at y = 0 across the full
  width; it is deliberately left behind. Sections are announced by a 目录
  divider instead, which frees the top 0.72 in of every slide for content.
- **Footer table**: x −0.36, y 7.00, 2.79 × 0.44 — the date band. The date is
  **literal text** in the table, not a field, so it must be restamped; the
  template's own reads 2026年1月21日. It is on every template slide.
- **Title icon**: small group at (0.22, 0.79), 0.53 × 0.54 — but only on the
  *titled* slides (7–19), not on 3–6, so harvest it and the footer
  independently. It is re-placed at y = 0.28 now the nav bar is gone.

## Vertical grid (nav bar removed)

- Slide title 24pt bold green at (0.88, **0.28**); the icon sits left of it.
- Content band **0.95 → 6.9**; the footer owns everything below 7.0.
- These live in `TITLE_Y` / `BODY_TOP` / `BODY_BOTTOM` — never hard-code them.

## 目录 divider (content.pptx)

- One before each **目录 entry** — four of them, not six: 研究背景与贡献 /
  研究方法 / 实验结果 / 结论与不足. 文献信息 is front matter and gets none; 研究背景
  and 本文贡献 share one. See `TOC_GROUPS` / `TOC_SKIP`.
- Every entry listed every time, numbered 01…0N in `007A49` squares
  0.79 × 0.79 at x 4.88, labels 24pt bold at x 5.92 with `spc="300"` tracking.
- Active item **black**; the rest `D9D9D9` (bg1 lumMod 85%).
- Rows are spaced 1.2035 in, tightened to fit and centred vertically when there
  are more than four sections.
- The 目录 badge (green squares + white 目录) is a picture-free group at
  (0.32, 0.15) 1.77 × 1.57 in `content.pptx`, deep-copied in. The divider takes
  the footer but **not** the title icon — the badge occupies that corner.
- **Date / page number** come from the layout placeholders. The master carries
  a text box reading `第‹#›页/共22页` — the total is **hard-coded**, so rewrite
  it to the real slide count on every build.
- None of the furniture contains a picture, which is why it can be deep-copied
  between slides; a `<p:pic>` would carry an `r:embed` that does not resolve on
  the destination slide.

## Slide patterns

- **Slide title**: x 0.88, y `TITLE_Y`, 24pt bold green, right of the icon.
- **要点框 (bullets)**: full-width wash band at x 0.35, w 12.6, from `BODY_TOP`.
  Height grows 0.46 in per bullet. Each bullet has one **green bold key term**.
- **卡片 (cards)**: 2–4 white rounded rects, green 1.25pt outline, gap 0.35,
  text vertically centred at 18pt. Lead word bold green.
- **结果分析 callout**: white rounded rect with a green pentagon tab carrying
  the head (结果分析 / 关键设计 / 小结); body text 18–20pt to the right of it.
  The tab is **min 1.35 in, widened to fit its head on one line** — the head is
  a label, and 「结果分 / 析」 across two lines reads as a bug. Everything that
  reserves room for a callout must ask `callout_metrics`, never assume 1.35.
- **Figure slides**: figure on one side (default right, w ~7.3), 要点框 plus
  callout on the other. A caption strip under a figure is 15pt centred
  **black** — green is the accent colour, and a caption is not an accent.
- **结论 (numbered)**: one white rounded box, big green numeral at x 0.85,
  bold green-dark title at 1.95, description under it.

### 汇报模板2 patterns (the antidote to 图 + 一句总结 on every page)

- **论断条 (statement)**: full-width white rounded box, green 1pt outline,
  19pt bold deep-green — the claim the slide argues, above everything else.
- **流程条 (process)**: 1.15 in strip on the bottom edge; PENTAGON tab (w 1.5)
  then N steps separated by 0.5 in green RIGHT_ARROWs.
- **分析面板 (panel)**: white rounded box down the right (w ~4.0), a green
  rounded pill as its head, then a numbered list. Auto-shrinks to fit.
- **图卡 (figure_cards)**: 论断条 + N cards, each 图 over a centred 19pt
  deep-green title over its description.
- **双分析 (dual)**: two figures side by side, each with its own chevron
  callout; all callouts share one top edge so the row reads as a row.
- **结论/局限 (proscons)**: two 0.42 in bands 0.75 in apart, green then grey,
  each with a 2.2 × 0.62 pill at the left; lists sit on the bands' outer sides.
- **红色批注 (annotations)**: 13pt bold red, `at` given as a fraction of the
  placed figure. They belong in the figure's margins — over a dense panel they
  hide what they point at.
- **Identity placeholders**: cover logo (0.07, 0.03, 3.04 × 0.95), QR
  (11.10, 0.03, 2.17 × 2.17), 汇报人 table (3.81, 5.24) whose third cell holds
  the name, homepage text box (7.15, 5.24); the 谢谢大家 slide repeats all four,
  and the banner QR sits at (10.55, 0.41). Unreplaced, each becomes a dashed
  grey box (`F2F6F4` fill, `8C8C8C` 1pt dash) naming what to drop in.

## Rules

- Never let a slide be a single figure with a caption — that is what a figure
  dump does, and it is what this template exists to replace.
- Never let a whole section be one layout repeated. 研究方法 and 实验结果 are
  where this goes wrong; match the layout to the content, not to habit.
- Text must fit its box: compute the wrapped height and shrink, never assume
  one line per item. A caption that wraps to two lines will otherwise be drawn
  straight through the footer.
- Content area is `BODY_TOP` → `BODY_BOTTOM`; the footer owns everything below
  7.0. Nothing sits in the reclaimed top band except the title and its icon.
- Scale figures to fit their box (preserve aspect); never stretch.
- Chinese punctuation must not start a line — tag runs `zh-CN`.
- Math is an equation object, never characters: `$\mathcal{L}_{\mathrm{rw}}$`,
  not `L_rw`. Match the paper's notation so a bullet and its figure agree.
- Ship none of the template author's identity — placeholder anything the spec
  does not replace. Their QR images resolve to their links whatever the caption
  says, so a retext is not a fix; the image itself has to go.
