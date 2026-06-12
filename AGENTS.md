# Agent Notes (AI ORM Manager PPTX)

Last updated: June 2026

Canonical engine is the Python pipeline in `ai_presentation_orm_v0_3_13/` (version `0.3.13`). Slides 01-12 are implemented in the Python pipeline. Do not use the legacy OOXML path unless explicitly requested.

## Entrypoints (Don't Guess)

- Node wrapper CLI: `node scripts/report_pptx/generate_report_pptx.mjs ...`
- Python pipeline: `python ai_presentation_orm_v0_3_13/run_pipeline.py --config <yaml>`
- IMPORTANT: the Node wrapper defaults to a legacy OOXML editor; to use the canonical Python pipeline you MUST pass `--engine python-v0.3.13` (see `ai_orm_manager_agent_manifest.json`). In this mode the wrapper now drives slides 01-12 and promotes per-slide QA blockers to the top-level generation log.

## Fastest Reproducible Run (Repo-Local Inputs)

1. Install deps: `pip install -r ai_presentation_orm_v0_3_13/requirements.txt`
1. Generate slides 01-12 via wrapper (creates a build dir next to `--output`):

```bash
node scripts/report_pptx/generate_report_pptx.mjs \
  --engine python-v0.3.13 \
  --template "reference/Бактоблис_ORM_май_2026_upd.pptx" \
  --raw "inputs/Baktoblis_01.05.2026-21.05.2026_6a1010a43842b7307304257c.xlsx" \
  --analytics "inputs/Baktoblis_01.05.2026-21.05.2026_6a1010a43842b7307304257c.xlsx" \
  --orm "inputs/Бактоблис ORM.xlsx" \
  --brand "Бактоблис" \
  --period "01–21 мая 2026" \
  --campaign-materials-override 35 \
  --campaign-month-sheet "Май" \
  --ratings-sheet "Рейтинги" \
  --ratings-current-period "Апрель 2026" \
  --output "outputs/report.pptx"
```

Notes:
- `--pipeline-root <path>` overrides `ai_presentation_orm_v0_3_13/` (useful if you clone/move the engine).
- The wrapper writes `outputs/report.generation-log.json` plus a build folder `outputs/report_ai_orm_build/` containing the generated YAML + per-slide JSON artifacts.
- Treat `generation-log.json.status=blocked` as a stop even when the PPTX file exists. The wrapper can keep a draft with `--strict false`, but client-ready delivery requires clearing the listed blockers.

## Python Config Gotchas (If Running `run_pipeline.py` Directly)

- `ai_presentation_orm/config_loader.py` requires YAML sections: `project`, `paths`, `processing`, `rules` (missing any will hard-fail).
- For reproducible runs, use `inputs.files` instead of scanning `paths.input_dir` (see `ai_presentation_orm/file_role_detector.py`).
- Always exclude Office lockfiles: `inputs.exclude_patterns: ["~$*"]`.
- Use a fresh `paths.output_dir` for each run if PowerPoint may have locked a generated PPTX.
- Repo-local example for the Stodal May build with agent-made browser screenshots: `python ai_presentation_orm_v0_3_13/run_pipeline.py --config outputs/config.stodal_may_views_browser_screenshots.yaml`.
- Keep code paths project-agnostic. New logic must be driven by config, sheet headers, period labels, and standard slide roles, not by a specific brand/client filename.
- In BA/Medialogia `Сообщения` sheets, do not use `Роль объекта` / `Object role` as the project object or brand column. Project-object filtering may use object/tag/brand columns only after validating that the column contains the project brand in data rows.
- Slide 08 view dynamics must use the last two valid monthly ORM sheets in the workbook and display the metric label as `Просмотры`.

## Slides 09-10 Browser Screenshots

- Slides 09-10 use agent-made browser screenshots from each row's source URL, not ready screenshot links from column `S`.
- Ready screenshot links may remain in extracted data for traceability, but the builder skips them and records `ready_asset_status: skipped`.
- Slides 09-10 must not use generated/rendered cards as a fallback. If a source cannot be browser-captured with the target text, skip it and try another real source row.
- Cropping a real browser screenshot to the slide slot is allowed when needed for readability, but do not synthesize, redraw, or regenerate the message content.
- Do not add picture borders, placeholder rectangles, blue DOM outlines, or other highlight frames around examples.
- Preferred local capture uses `playwright` plus a local Chrome or Edge installation. If Python Playwright is unavailable, the pipeline can fall back to headless Chrome/Edge CLI where the environment permits it.
- Custom GPT / ChatGPT mode: when local browser automation is unavailable, the pipeline writes `slide09_10_screenshots/chatgpt_screenshot_requests.json`. The GPT agent must open each pending URL with its ChatGPT browsing/browser capability, take a real screenshot of the visible source page or message/review, save the PNG/JPG exactly to `expected_path` (or an `accepted_alternative_paths` entry), and rerun the pipeline. On rerun, the builder consumes those images as `source_kind: chatgpt_browser_capture`.
- To force this handoff instead of trying local browser automation, set `AI_ORM_SCREENSHOT_BACKEND=chatgpt` or pass `--screenshot-backend chatgpt` through the Node wrapper.
- If screenshots still cannot be created, the builder inserts the best available source links into slide 09/10 slots in source order and QA returns `ready_for_manual_screenshot_links`. This is acceptable for a handoff deck where a specialist will create screenshots manually later.
- Select examples from browser-visible public pages with compact message text. Avoid sources where the browser cannot find the target message or captures mostly ads.
- In restricted GPT-agent environments, internet access, browser automation, and Windows PowerPoint COM may be unavailable. Treat these as environment limitations, not data failures; for screenshots, use the ChatGPT handoff manifest before accepting a blocked slide 09/10 QA result.
- PNG visual audit export via PowerPoint COM is Windows-only. On non-Windows agents, inspect PPTX artifacts and JSON QA, then ask for a Windows visual export if needed.

## Outputs To Trust (And What To Check)

- The pipeline always writes inventory/diagnostics into `paths.output_dir` (e.g. `template_map.json`, `data_inventory.json`, `file_roles.json`, `missing_data_report.md`).
- Slide builds write JSON next to PPTX drafts (e.g. `slide_05_data_model.json`, `slide_05_build_result.json`, `slide_05_qa.json`). Treat `*_qa.json: status=blocked` as a stop.
- Even when QA passes, visual QA is mandatory (we've had "numbers OK, layout broken" failures: cropped labels, overflow, tiny fonts, wrong proportions).
- For slides 09-10, verify `source_kind: browser_capture` or `chatgpt_browser_capture` when screenshots are expected. In restricted Custom GPT runs, `source_kind: manual_screenshot_link` plus `fallback_used: false` is the approved manual-capture handoff state.

## Non-Negotiable Content/Template Rules (Hard-Learned)

- Template-first: edit the target slide in-place in the reference PPTX and preserve geometry unless explicitly approved.
- Formatting defaults: slide titles are Arial 28, main analytical text is Arial 11, chart labels/values/legends are Arial 10, table text is Arial 9 with vertical middle alignment.
- Keep the title slot and bottom source-note slot consistent across slides. Align left-side slide content to the title left edge unless a reference chart/table slot explicitly requires centered internal content.
- Use spaces between thousands on chart/table values. Keep `Другие` / `Остальные` / `Прочие` / `Other` buckets at the end of chart series.
- Prefer editable PowerPoint charts/native labels; when bitmap chart fallback is unavoidable, render at final slide size using the same font constants.
- Never copy stale template content: old brands, months, dates, numbers, or conclusions (explicit stale checks exist in slide QA).
- Project brand must be **bold** everywhere it appears in client-facing text/labels (combined deck runs `bold_project_brand_in_presentation`).
- Source notes: name source system + period; do NOT include local filenames (`.xlsx`/`.pptx`) and do NOT include hours/minutes.
- No semicolons (`;`) in client-facing copy (explicit QA blockers for slides 04-07).
- Slide 06 negative tone must be reconciled against the project ORM month sheet.
  If the `Негатив` section exists, count only real publication rows inside that
  section until the next section header. If the section is absent or empty,
  slide 06 shows `0` negative and treats monitored negative tone as neutral.
- No causal "campaign impact" claims in client-facing copy. A verified campaign/organic split can be shown as a scenario comparison, but not phrased as proven impact unless the brief explicitly provides impact evidence.
- If you touch OOXML directly (legacy wrapper mode), always open the PPTX and scan for broken Cyrillic (`????`).
- Preserve clickable links in tables/examples.

## Where The Truth Lives

- Slide-by-slide rules + remaining plan: `HANDOFF_CHATGPT_PLUS_CANVAS.md`.
- Canonical engine + how to run: `ai_presentation_orm_v0_3_13/README.md`, `ai_presentation_orm_v0_3_13/run_pipeline.py`.
- Wrapper behavior (engine selection, build dir, strictness): `scripts/report_pptx/generate_report_pptx.mjs`.
