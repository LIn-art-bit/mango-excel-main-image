---
name: mango-excel-main-image
description: Use this skill whenever the user wants to process Mango ERP or AliExpress pending-publish Excel .xlsx files, batch-generate AI-refined ecommerce main images from 产品图片1, preserve 本地ID and 产品标题, handle long-running 100+ row image queues, or return a workbook with a 主图 column.
---

# Mango Excel Main Image

This skill processes Mango/AliExpress pending-publish `.xlsx` workbooks into a compact output workbook for AI-refined product main images.

Use it when the user mentions:
- Mango ERP pending-publish Excel files
- AliExpress/Mango product rows with `本地ID`, `产品标题`, `产品图片1`
- batch AI image generation from product image links
- adding a `主图` column to an Excel file
- naming generated images by local product ID
- long-running AI refinement batches such as 100, 300, or more rows

## Input Contract

- Input must be `.xlsx`.
- Read the first worksheet unless the user names another sheet.
- Treat row 2 as the real header row.
- Required headers:
  - `本地ID`
  - `产品标题`
  - `产品图片1`
- A valid product row must have all three required fields non-empty.
- Skip rows with empty `本地ID`, empty `产品标题`, or empty `产品图片1`; record skips in the log.
- If `本地ID` is duplicated, keep the first occurrence and skip later duplicates to avoid overwriting generated images.

## Output Contract

Create a new `.xlsx` workbook with exactly these columns:

1. `本地ID`
2. `产品标题`
3. `产品图片1`
4. `主图`

The `主图` value is the local absolute path to the verified AI-refined image. Add it as a clickable hyperlink when possible.

The user-facing output must be aggregated by the main agent, even when image generation is done by parallel worker agents. Worker agents may report internal handoff details, but the user should receive one consolidated progress or final summary, not separate worker-by-worker final reports.

By default, save batch outputs under this fixed D drive root:

```text
D:\MangoMainImageBatches
```

If `D:\MangoMainImageBatches` does not exist, create it. If it already exists, reuse it. If the user asks for a different destination drive or folder, pass `--output-root <folder>` and keep all batch outputs there.

For a newly uploaded or newly provided Excel workbook, start a fresh task by default with `prepare --new-task`. This creates a timestamped subfolder and updates the latest-batch pointer, so a new upload does not accidentally continue an older interrupted run for a workbook with the same file name. Resume only when the user explicitly asks to continue a previous run, or when the current run is interrupted and the latest-batch pointer belongs to the same active task.

Save verified AI-refined images under:

```text
D:\MangoMainImageBatches/<input_stem>_main_image_batch/generated_images/
```

Keep non-final outputs outside `generated_images`, for example:

```text
<batch_folder>/staging/
<batch_folder>/review_queue/
<batch_folder>/generated_fast/
```

Every generated main image must be named with the current row's `本地ID`:

```text
<本地ID>.png
```

Never use row numbers as generated-image names. Never overwrite an existing `<本地ID>.png`; skip or ask before replacing unless the user explicitly requests regeneration. Never place an unfinished or unverified image in `generated_images`.

## AI Refinement Requirement

- The default and expected output for this skill is an AI-refined product main image.
- Do not satisfy this skill with deterministic template composition, simple resizing, background padding, thumbnail generation, or scripted poster layout unless the user explicitly asks for a fast non-AI fallback.
- `generated_images/<本地ID>.png` is reserved for verified AI-refined images only.
- If a fast template/composition fallback is explicitly requested, save those files in a clearly separate folder such as `<batch_folder>/generated_fast/` or `<batch_folder>/template_images/`, and do not mark them as `verified` AI-refined outputs.
- Do not use `scripts/compose_batch_images.py` for the final `主图` workflow unless the user explicitly chooses speed over AI refinement for that run.
- `scripts/compose_batch_images.py` may only create `generated_fast/<本地ID>.png` and mark rows `template_generated`; those rows are still unfinished for the final AI-refined workflow.
- Each valid product row requires its own image generation call using the original downloaded product image as reference and the fixed prompt template with that row's title.
- A row may be marked `verified` only after the generated file is an AI-refined result, exists on disk, opens successfully, and is normalized to 800x800 PNG.

## Default Workflow

1. Use the Codex bundled Python executable from `load_workspace_dependencies`; do not rely on bare `python` in this Windows desktop environment.
2. Do not pause for style confirmation by default. For a newly uploaded Excel, prepare the durable full queue immediately with `prepare --limit 0 --run-mode full --new-task`. The manifest and ledger record `run_mode`, `requested_limit`, input fingerprint, valid row count, and `batch_id` so a new task cannot be mistaken for an older run.
3. Use the returned `batch_id` and latest-batch pointer for the rest of the current run. If the user explicitly asks to resume an older batch, pass that older `--batch-id`.
4. Use `status --limit 0 --run-mode full` before generating images.
5. Use 2 concurrent worker agents by default for long full runs. Dispatch both workers before waiting for either one to finish, so worker 1 and worker 2 can each run image generation at the same time when the platform allows it.
6. Each worker reserves its own work with `claim --limit 0 --run-mode full --count 10 --owner <worker-id> --lease-minutes 15`. Do not simulate two workers by having the main agent process worker 1 and then worker 2 serially.
7. For each claimed item, open the downloaded original image with `view_image` so it is visible as the edit/reference image.
8. Use the fixed prompt template in `references/prompt_template.md`; replace `{title}` with the row's `产品标题`.
9. Generate one image per product using the image generation tool.
10. Save intermediate outputs to `<batch_folder>/staging/` or `<batch_folder>/review_queue/`.
11. Only after the generated image is accepted as an AI-refined final, move or copy it into `<batch_folder>/generated_images/<本地ID>.png`.
12. Immediately run `mark --limit 0 --run-mode full --local-id <本地ID> --status verified --owner <worker-id>` after each completed row.
13. If generation fails, save the error with `mark --limit 0 --run-mode full --local-id <本地ID> --status failed --error "<reason>"` and continue to the next item when possible.
14. Use `build --allow-partial` only for interim review workbooks. The final `build` command must be strict and must fail if any row is unfinished or missing a verified image.
15. Run `verify --limit 0 --run-mode full` before final delivery.

## Hard Completion Rule

- The task is not complete until every valid product row has finished the fixed-prompt AI refinement flow.
- Completion requires all of the following at the same time:
  - every valid row is marked `verified` in the durable status ledger
  - every verified row has a final `generated_images/<本地ID>.png` file on disk
  - the final Excel workbook has been rebuilt
  - the final Excel workbook points only to verified images on disk
- If any row remains `pending` or `failed`, the task is still in progress and must not be treated as finished.
- If any row remains `claimed` or `template_generated`, the task is also still in progress; release stale claims or resume AI refinement.
- Do not stop early just because some images already exist on disk or because a partial Excel file was created.
- After each image completes, immediately persist the image to disk, update the durable status ledger, and continue to the next unfinished row.
- If the run is interrupted by a crash, context loss, sandbox issue, network problem, rate limit, or any other external interruption, resume from the next unfinished row instead of restarting completed rows.
- Use heartbeat automation as a recovery mechanism during long full runs. Default to a 15-minute heartbeat while unfinished rows remain.
- On heartbeat resume, run `status --limit 0 --run-mode full`; if unfinished is not 0, continue from the latest batch by claiming the next 10 rows per worker.
- Before final delivery, run `status --limit 0 --run-mode full` and confirm `unfinished` is 0, then run `verify --limit 0 --run-mode full` and confirm `ok` is true.
- For a newly uploaded/provided Excel, create a fresh task folder with `--new-task` by default. Reuse an existing `manifest.json` and `status_ledger.json` only when resuming the current latest batch after an interruption, or when the user explicitly asks to continue a previous batch.

## Long-Run Design For 300+ AI Images

- Treat 100+ or 300+ row jobs as long-running AI production queues, not one-shot chat tasks.
- Initialize the full manifest and status ledger first with `prepare --limit 0 --run-mode full --new-task`. By default this writes to a timestamped subfolder under `D:\MangoMainImageBatches\`.
- Process repeated claimed batches of 10 unfinished rows per worker by default.
- After every AI-refined image, save the PNG, mark that row `verified`, and continue. Never wait until the whole batch is done to update the ledger.
- Rebuild an interim workbook after each completed batch when useful, but do not call the task complete until the final full workbook has been rebuilt and verified.
- Keep a heartbeat automation active during the run with a 15-minute interval. Pause or delete it after `unfinished` is 0.
- If the user explicitly asks for maximum throughput, explain that true parallel AI generation requires an external image-generation API or queue. Do not pretend that local multi-agent orchestration alone guarantees faster built-in image generation.
- If the user asks for a long full run and does not specify worker count, use 2 concurrent worker agents by default. Start both workers before waiting, so two images can be in generation at once when the image tool/platform supports parallel tool calls. Use 3 workers only if the user explicitly asks for maximum throughput. Every worker must call `claim --owner <worker-id>` before generating images. Workers must not scan the whole ledger and choose rows themselves.
- If subagents cannot access the image generation tool or the platform serializes image generation calls, report that limitation clearly and continue with the ledger-based queue rather than pretending the run is truly parallel.
- If rate limits, image-tool failures, or context loss occur, record failed rows in the ledger and resume later from `failed` plus `pending` rows. Expired `claimed` rows can be reclaimed by a later `claim` call.

## Aggregated Reporting

- Parallel workers are execution helpers, not independent user-facing deliverables. The main agent owns all user-facing progress reports and final delivery.
- Each worker must return a concise structured handoff to the main agent with: `worker_id`, claimed local IDs, verified local IDs, failed local IDs with reasons, generated image paths, and whether any claims remain open.
- The main agent must aggregate worker handoffs with a fresh `status --limit 0 --run-mode full` check. For final delivery, the main agent must also run strict `build` and `verify`.
- Interim user updates should summarize the whole run: `batch_id`, batch folder, total valid rows, verified count, pending count, claimed count by worker, failed count, current worker activity, interim workbook path if one exists, and next action.
- Final user output must be one consolidated summary containing: final Excel path, `generated_images` folder, total valid rows, verified rows, failed rows, skipped rows if any, `verify ok=true`, and any limitation encountered with true parallel image generation.
- If unfinished rows remain, label the workbook or status as interim. Do not present worker completion as whole-task completion.
- If one worker finishes early while another is still generating, the main agent should claim another batch for the free worker when unfinished rows remain.

## Batch Defaults

- Default production mode for a newly uploaded Excel: full queue, no style-confirmation pause, `--new-task`, 2 concurrent worker agents, 10 claimed rows per worker, 15-minute heartbeat.
- Sample mode is optional only when the user explicitly asks to preview style first.
- Use a higher `--limit` only when the user asks for a bounded batch instead of full processing.
- For full processing, continue until all valid rows are verified; a partial batch does not count as completion.
- For AI-refinement production runs, use claimed batches of 10 rows per worker by default.

## Script Usage

Prepare a fresh full task for a newly uploaded Excel:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" prepare `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full `
  --new-task
```

Optional: prepare a sample batch only if the user explicitly asks to preview style:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" prepare `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 3 `
  --run-mode sample `
  --new-task
```

Check progress:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" status `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full
```

Claim the next AI-refinement batch:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" claim `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full `
  --count 10 `
  --owner "worker-1" `
  --lease-minutes 15
```

For the default 2-worker long run, claim a second batch for worker 2:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" claim `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full `
  --count 10 `
  --owner "worker-2" `
  --lease-minutes 15
```

Mark one AI-refined row as verified after its final image exists:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" mark `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full `
  --local-id "<本地ID>" `
  --status verified `
  --owner "worker-1"
```

Build an interim workbook during the run:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" build `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full `
  --allow-partial
```

Build the final strict output workbook after every row is verified:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" build `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full
```

Verify final completion:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" verify `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full
```

Release a worker's stale claims if you need to stop or hand off work:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" release `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full `
  --owner "worker-1"
```

The script prints JSON with `manifest_path`, `status_ledger`, `originals_dir`, `generated_dir`, `generated_fast_dir`, `staging_dir`, `review_queue_dir`, `output_xlsx`, status counts, and skip details.

Use a custom output root only when the user explicitly asks:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" prepare `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --run-mode full `
  --new-task `
  --output-root "D:\SomeOtherFolder"
```

## Important Constraints

- This skill does not upload generated images to Mango/CDN/public hosting.
- The `主图` column uses local absolute paths unless the user later provides an upload target.
- `generated_images` is reserved for verified AI-refined outputs only.
- A partial batch is not a successful completion.
- Keep the fixed prompt's safety rules intact: no Chinese text, watermarks, platform marks, protected automotive brand identity, fake certifications, or unsupported promotional claims.
- Use one image generation call per product. Distinct products need distinct prompts and reference images.
- Never silently downgrade an AI-refinement request to scripted composition. If speed and AI quality conflict, tell the user and ask whether they want AI quality, fast template output, or an external API/queue design.
