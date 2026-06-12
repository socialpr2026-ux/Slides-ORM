# AI ORM Manager Custom GPT Kit

This archive contains the clean script package for the AI ORM Manager PPTX agent.
It is project-agnostic: do not hard-code client names, file names, months, sheet
names, or one-off corrections. Drive behavior from input files, config, detected
headers, and the user's confirmed rules.

## What Is Included

- `ai_presentation_orm_v0_3_13/` - canonical Python pipeline for slides 01-12.
- `scripts/report_pptx/generate_report_pptx.mjs` - Node wrapper CLI.
- `reference/` - editable PPTX template used as the in-place slide skeleton.
- `reference/regression/` - finished reference decks for regression comparison
  (visual/content targets only; do not treat them as source data).
- `AGENTS.md` and `ai_orm_manager_agent_manifest.json` - operating rules.
- `DEEPSEEK_IMPLEMENTATION_INSTRUCTIONS.md` - detailed overnight implementation
  brief for improving the universal generator.
- `examples/generation_config.example.yaml` - generic config example.

Generated outputs, input workbooks, cache folders, vendored runtimes, screenshots,
and client data are intentionally not included.

## Main Entry Point

Prefer the wrapper when possible:

```bash
node scripts/report_pptx/generate_report_pptx.mjs \
  --engine python-v0.3.13 \
  --template "reference/Бактоблис_ORM_май_2026_upd.pptx" \
  --raw "inputs/analytics_or_raw_export.xlsx" \
  --analytics "inputs/analytics_or_medialogia_export.xlsx" \
  --orm "inputs/project_orm_table.xlsx" \
  --brand "Project brand" \
  --period "April 2026" \
  --campaign-month-sheet "Publication (April)" \
  --ratings-sheet "Ratings" \
  --ratings-current-period "April 2026" \
  --screenshot-backend chatgpt \
  --output "outputs/report.pptx"
```

Direct Python runs are also supported:

```bash
python ai_presentation_orm_v0_3_13/run_pipeline.py --config outputs/config.generated.yaml
```

## Universal Data Rules

- Campaign equals the brand's own ORM-table messages for the month.
- Campaign includes all ORM-month publication rows for reviews and comments,
  including comments in open discussions, top/search comments, SERM comments,
  embedded discussion comments, and native-photo comments.
- Organic equals the remaining brand mentions in all monthly brand mentions.
- Media-plan fact must be overridden by the project ORM month publication count
  when that month table is available. Do not use erroneous media-plan fact cells
  as the authoritative publication fact.
- Before branded SOV calculations, remove non-brand thematic objects/topics.
- Medialogia tags may be named as objects. Treat them as source objects when BA
  exports are absent.
- In BA/Medialogia message sheets, never treat `Object role` / `Роль объекта`
  as a brand/object column. A project-object column must be an object/tag/brand
  column and must actually contain the project brand in data rows.
- Slide 04 thematic blocks must be semantic contexts, not raw single keywords:
  group frequent words/topics into clear 2-3 word labels, remove brand and
  competitor objects, suppress non-medical thematic noise, avoid duplicate raw
  terms already included in a group, and keep chart labels capitalized.
- Slide 06 daily tonality line chart must use the legend/color order:
  `Позитив` green, `Нейтрал` gray, `Негатив` red.
- Slide 06 negative tonality is controlled by the project ORM month sheet:
  always check the `Негатив` section first. Count only actual publication rows
  inside that section until the next section header. If the section is absent or
  empty, show `0` negative on slide 06 and move monitored negative tone into
  neutral for this slide.
- Views may be missing on some platforms. Use the project-table column named
  like `Views received`, `Views`, or the closest localized equivalent.
- For slide 08 view dynamics, use the last two valid monthly ORM sheets in the
  project workbook and label the metric simply as `Просмотры`.
- Count rows with actual data, not worksheet physical rows.
- If ratings data is absent and the user confirms that ratings are unavailable,
  remove or skip slide 11 rather than inventing values.

## Chart Layout Rules

- Slide titles use Arial 28 in the same title slot on every slide.
- Main analytical text uses Arial 11 with consistent line spacing.
- Chart labels, category labels, values, and legends use Arial 10.
- Table text uses Arial 9 with vertical middle alignment.
- Chart numbers use spaces between thousands.
- `Другие` / `Остальные` / `Прочие` / `Other` buckets always stay at the end
  of chart series.
- Prefer editable PowerPoint charts and native labels. If bitmap fallback is
  required, render at final slide size with the same font constants.
- Pie and donut labels must remain outside the chart circle/ring. If Arial 10
  overlaps the chart area, reflow the label layout instead of shrinking below
  Arial 10.
- Pie and donut labels should be editable PowerPoint text blocks when possible:
  color square, label, count, and percent. The color square must match the sector
  from the same data row.
- Stacked bar charts must not place values on tiny segments when the value
  becomes unreadable. Move small non-zero values outside with a clear vertical
  gap and suppress zero-value labels.
- Legends must be separated from axis/category labels with enough bottom margin.
- Slide 11 chart images must use a common vertical chart slot, common baseline,
  and common legend line across all three mini charts.
- Analytical conclusions must not duplicate source notes such as "based on the
  ratings sheet" because sources are shown in the standard slide footnote.
- Source notes stay in the standard bottom slot and must not jump between slides.
- Slide 12 final conclusions must follow the approved summary structure:
  five top KPI blocks (`Упоминания`, `Индекс лояльности`, `SOV без кампании`,
  `SOV с кампанией`, `Прирост инфополя`) and five analytical bullets:
  competitive SOV position, platform/source structure, brand tonality, agency
  campaign activity, and forward recommendation. Do not replace this with the
  older `SOV / Без кампании / С кампанией / Разница` metric set or a ratings
  paragraph.

## Slides 09-10 Screenshot Policy

Slides 09-10 must use real source-page screenshots when the environment allows
browser capture. Do not synthesize, redraw, or generate message/review images.

For Custom GPT, use:

```bash
--screenshot-backend chatgpt
```

or set:

```bash
AI_ORM_SCREENSHOT_BACKEND=chatgpt
```

In this mode the pipeline writes:

```text
slide09_10_screenshots/chatgpt_screenshot_requests.json
```

If ChatGPT/browser tools can capture screenshots, save each image to the
`expected_path` from that manifest and rerun the pipeline. The builder will use
those images as `source_kind: chatgpt_browser_capture`.

If screenshots cannot be created, the pipeline inserts the best available source
links directly into slide 09 and slide 10 slots, in source order. QA status will
be `ready_for_manual_screenshot_links`; this is acceptable for a handoff deck
where a specialist will create screenshots manually later.

## Readiness Check

After each run, inspect:

- `*.generation-log.json` from the wrapper.
- `slide_XX_qa.json` files in the build folder.
- Final PPTX slide count must stay 12 unless the user explicitly approves
  removing a slide.
- Treat QA blockers as stop conditions. Warnings on manual screenshot links are
  expected in restricted Custom GPT environments.
