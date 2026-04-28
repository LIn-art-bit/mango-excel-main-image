---
name: mango-excel-main-image
description: Use this skill whenever the user wants to process Mango ERP or AliExpress pending-publish Excel .xlsx files, preserve 本地ID and 产品标题, use 产品图片1 as the reference image, batch-generate AI-refined ecommerce main images, name each output image by 本地ID, and return a new Excel workbook with a 主图 column containing local generated-image paths.
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

By default, save batch outputs under this fixed D drive root:

```text
D:\MangoMainImageBatches
```

If `D:\MangoMainImageBatches` does not exist, create it. If it already exists, reuse it. If the user asks for a different destination drive or folder, pass `--output-root <folder>` and keep all batch outputs there.

Save verified AI-refined images under:

```text
D:\MangoMainImageBatches/<input_stem>_main_image_batch/generated_images/
```

Keep non-final outputs outside `generated_images`, for example:

```text
<batch_folder>/staging/
<batch_folder>/review_queue/
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
- Each valid product row requires its own image generation call using the original downloaded product image as reference and the fixed prompt template with that row's title.
- A row may be marked `verified` only after the generated file is an AI-refined result, exists on disk, opens successfully, and is normalized to 800x800 PNG.

## Default Workflow

1. Use `scripts/process_mango_excel.py prepare` to inspect the workbook, extract valid rows, download original images, write a manifest, and create/update the durable status ledger.
2. Use `scripts/process_mango_excel.py status` to inspect the durable status ledger before generating images.
3. Use `scripts/process_mango_excel.py next --count <N>` to get the next unfinished AI-refinement batch.
4. For each returned item, open the downloaded original image with `view_image` so it is visible as the edit/reference image.
5. Use the fixed prompt template in `references/prompt_template.md`; replace `{title}` with the row's `产品标题`.
6. Generate one image per product using the image generation tool.
7. Save intermediate outputs to `<batch_folder>/staging/` or `<batch_folder>/review_queue/`.
8. Only after the generated image is accepted as an AI-refined final, move or copy it into `<batch_folder>/generated_images/<本地ID>.png`.
9. Immediately run `scripts/process_mango_excel.py mark --local-id <本地ID> --status verified` after each completed row.
10. If generation fails, save the error with `scripts/process_mango_excel.py mark --local-id <本地ID> --status failed --error "<reason>"` and continue to the next item when possible.
11. Use `scripts/process_mango_excel.py build` to create the output workbook with the `主图` paths after all rows are verified, or after a batch if the user wants an interim workbook.
12. Verify the workbook:
   - output rows match the processed valid item count
   - every `主图` path exists
   - generated image filenames exactly match `本地ID`

## Hard Completion Rule

- The task is not complete until every valid product row has finished the fixed-prompt AI refinement flow.
- Completion requires all of the following at the same time:
  - every valid row is marked `verified` in the durable status ledger
  - every verified row has a final `generated_images/<本地ID>.png` file on disk
  - the final Excel workbook has been rebuilt
  - the final Excel workbook points only to verified images on disk
- If any row remains `pending` or `failed`, the task is still in progress and must not be treated as finished.
- Do not stop early just because some images already exist on disk or because a partial Excel file was created.
- After each image completes, immediately persist the image to disk, update the durable status ledger, and continue to the next unfinished row.
- If the run is interrupted by a crash, context loss, sandbox issue, network problem, rate limit, or any other external interruption, resume from the next unfinished row instead of restarting completed rows.
- Use heartbeat automation only as a recovery mechanism when the main run stops unexpectedly before all rows are verified.
- For long tasks, prefer a 10 to 15 minute heartbeat interval so interrupted runs resume promptly.
- Before final delivery, run `status --limit 0` and confirm `unfinished` is 0.
- Always check the fixed output batch folder before starting a new run. If `manifest.json` and `status_ledger.json` already exist for the same input workbook, reuse them and continue from unfinished rows instead of creating a separate duplicate batch folder.

## Long-Run Design For 300+ AI Images

- Treat 100+ or 300+ row jobs as long-running AI production queues, not one-shot chat tasks.
- Initialize the full manifest and status ledger first with `prepare --limit 0`. By default this writes to `D:\MangoMainImageBatches\<input_stem>_main_image_batch`.
- Process repeated active batches of 20 to 30 unfinished rows by default. This balances progress, reviewability, rate-limit recovery, and interruption recovery.
- After every AI-refined image, save the PNG, mark that row `verified`, and continue. Never wait until the whole batch is done to update the ledger.
- Rebuild an interim workbook after each completed batch when useful, but do not call the task complete until the final full workbook has been rebuilt and verified.
- Keep a heartbeat automation active during the run with a 10 to 15 minute interval. Pause or delete it after `unfinished` is 0.
- If the user explicitly asks for maximum throughput, explain that true parallel AI generation requires an external image-generation API or queue. Do not pretend that local multi-agent orchestration alone guarantees faster built-in image generation.
- If the user explicitly asks to use multiple agents, split only by disjoint row ranges or manifest chunks, and require every worker to use the shared status ledger without overwriting rows owned by another worker.
- If rate limits, image-tool failures, or context loss occur, record failed rows in the ledger and resume later from `failed` plus `pending` rows.

## Batch Defaults

- Default sample limit: `3`.
- Use a higher `--limit` only when the user asks for a larger batch or full processing.
- For full processing, continue until all valid rows are verified; a partial batch does not count as completion.
- For AI-refinement production runs, use active batches of 20 to 30 rows by default.

## Script Usage

Prepare the sample batch:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" prepare `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 3
```

Prepare the full long-run queue:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" prepare `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0
```

Check progress:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" status `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0
```

Get the next AI-refinement batch:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" next `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --count 25
```

Mark one AI-refined row as verified after its final image exists:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" mark `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
  --local-id "<本地ID>" `
  --status verified
```

Build the output workbook after generated AI-refined images exist:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" build `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0
```

The script prints JSON with `manifest_path`, `status_ledger`, `originals_dir`, `generated_dir`, `staging_dir`, `review_queue_dir`, `output_xlsx`, status counts, and skip details.

Use a custom output root only when the user explicitly asks:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" prepare `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 0 `
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
