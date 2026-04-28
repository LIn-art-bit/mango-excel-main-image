---
name: mango-excel-main-image
description: Use this skill whenever the user wants to process Mango ERP or AliExpress pending-publish Excel .xlsx files, preserve 本地ID and 产品标题, use 产品图片1 as the reference image, batch-generate ecommerce main images, name each output image by 本地ID, and return a new Excel workbook with a 主图 column containing local generated-image paths.
---

# Mango Excel Main Image

This skill processes Mango/AliExpress pending-publish `.xlsx` workbooks into a compact output workbook for generated main product images.

Use it when the user mentions:
- Mango ERP pending-publish Excel files
- AliExpress/Mango product rows with `本地ID`, `产品标题`, `产品图片1`
- batch image generation from product image links
- adding a `主图` column to an Excel file
- naming generated images by local product ID

## Input contract

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

## Output contract

Create a new `.xlsx` workbook with exactly these columns:

1. `本地ID`
2. `产品标题`
3. `产品图片1`
4. `主图`

The `主图` value is the local absolute path to the verified AI-refined image. Add it as a clickable hyperlink when possible.

By default, save batch outputs under the input file folder. If the user asks for a destination drive or folder, pass `--output-root <folder>` and keep all batch outputs there.

Save verified AI-refined images under:

```text
<output_root_or_input_folder>/<input_stem>_main_image_batch/generated_images/
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

## Default workflow

1. Use `scripts/process_mango_excel.py prepare` to inspect the workbook, extract valid rows, download original images, and write a manifest.
2. For each manifest item, open the downloaded original image with `view_image` so it is visible as the edit/reference image.
3. Use the fixed prompt template in `references/prompt_template.md`; replace `{title}` with the row's `产品标题`.
4. Generate one image per product using the image generation tool.
5. Save intermediate outputs to `<batch_folder>/staging/` or `<batch_folder>/review_queue/`.
6. Only after the row is marked `verified`, move or copy the final generated image into `<batch_folder>/generated_images/<本地ID>.png`.
7. Use `scripts/process_mango_excel.py build` to create the output workbook with the `主图` paths.
8. Verify the workbook:
   - output rows match the processed valid item count
   - every `主图` path exists
   - generated image filenames exactly match `本地ID`

## Hard completion rule

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
- Use the hourly heartbeat automation only as a recovery mechanism when the main run stops unexpectedly before all rows are verified.

## Batch defaults

- Default sample limit: `3`.
- Use a higher `--limit` only when the user asks for a larger batch or full processing.
- For full processing, continue until all valid rows are verified; a partial batch does not count as completion.

## Script usage

Prepare the sample batch:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" prepare `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 3 `
  --output-root "D:\"
```

Build the output workbook after generated images exist:

```powershell
& "<bundled-python>" "C:\Users\admin\.codex\skills\mango-excel-main-image\scripts\process_mango_excel.py" build `
  --input "C:\Users\admin\Downloads\Mango-ae-0428685372.xlsx" `
  --limit 3 `
  --output-root "D:\"
```

The script prints JSON with `manifest_path`, `originals_dir`, `generated_dir`, `staging_dir`, `review_queue_dir`, `output_xlsx`, and skip details.

## Important constraints

- This skill does not upload generated images to Mango/CDN/public hosting.
- The `主图` column uses local absolute paths unless the user later provides an upload target.
- `generated_images` is reserved for verified AI-refined outputs only.
- A partial batch is not a successful completion.
- Keep the fixed prompt's safety rules intact: no Chinese text, watermarks, platform marks, protected automotive brand identity, fake certifications, or unsupported promotional claims.
- Use one image generation call per product. Distinct products need distinct prompts and reference images.
