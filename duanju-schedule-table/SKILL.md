---
name: duanju-schedule-table
description: Create polished Excel 顺场表 / 顺畅表 / 拍摄顺序表 from short-drama scripts in .doc, .docx, or text-based .pdf files. Use when the user gives one or more 短剧剧本 files and asks to整理顺场表, 按剧本顺序做顺场表, 填场次/场景/拍摄内容/角色/日夜内外/群演/梳化服/道具, 核对后交付 .xlsx.
---

# 短剧顺场表

Use this skill to turn short-drama scripts from Word or PDF into the user's preferred Excel 顺场表 format.

## Output Rules

- Create one `.xlsx` per script unless the user asks otherwise.
- Sheet name must be `顺场表`; do not leave extra sheets.
- Keep one header row only. Do not add separate `演员` / `角色` rows.
- Header order:
  `拍摄顺序`, `实际拍摄场地`, `场次`, `剧本中场景`, `拍摄内容`, `日/夜`, `内/外`, `页数`, role columns, `群演`, `梳化服提示`, `道具提示`, `备注/特殊道具/时间`.
- Leave `实际拍摄场地` blank when no fixed location is provided.
- Keep `页数` in the header but leave all page cells blank unless the user explicitly provides page counts.
- Put role names in individual header columns. Mark scene appearances with a one-character abbreviation in the role's column.
- Use vertical headers only for role columns and the short columns `日/夜`, `内/外`, `页数`. Keep all other headers horizontal.
- `拍摄内容` must be a concise action/event summary in 12-18 Chinese characters. Write it as "who + does what / what happens", not as a short label.
  - Good: `顾宴赶到前台当众维护母亲`
  - Good: `陆延赶到河边当众救下小满`
- Keep `拍摄顺序` in script order unless the user provides a separate shooting order.

## Workflow

1. Use the Spreadsheet, Documents, and PDF skills as needed for `.xlsx`, `.doc/.docx`, and `.pdf` handling.
2. Normalize the input before scene parsing:
   - For `.doc`, convert to `.docx` first with `textutil` or another reliable converter.
   - For text-based `.pdf`, extract text with `pdfplumber`; inspect 1-3 rendered pages when layout may affect scene order.
   - For watermarked PDFs, filter watermark words by PDF attributes when possible, such as oversized gray text or watermark fonts, instead of deleting useful Chinese characters globally.
   - For scanned/image-only PDFs, use OCR if available; otherwise tell the user the PDF is image-only and ask for a text/Word export.
3. Extract script paragraphs/lines and split scenes by episode headings and scene headings.
4. Handle messy headings deliberately:
   - explicit scene headings like `1-1 夜 内 酒店-套房`;
   - generated first-episode headings like `日 内 办公室`;
   - missing spaces like `19-2日 内 ...` or `14-1 日 内公司/...`;
   - missing day in explicit headings like `8-3 内 公司/...`; infer from nearby scenes and record the assumption in your own check notes.
   - PDF-style headings like `1. 村委会会议室 日 内`, `3.路口日外`, and headings split across lines.
   - montage or insert headings like `一组镜头`, `手机屏幕`, `闪回`, or `番外篇`; include them rather than dropping them, and infer or mark missing `日/夜` and `内/外`.
5. Build role columns from `人物` lines. Treat `若干`, `众人`, `公司员工若干`, `应聘者若干`, and similar group labels as `群演`, not actor columns, unless the user wants them split.
6. Fill:
   - `剧本中场景`: source scene location text, lightly normalized only for spacing.
   - `拍摄内容`: summarize the actual scene action, not the episode title. Keep it 12-18 Chinese characters and avoid label-like short phrases.
   - `群演`: groups appearing in `人物` lines or obvious crowd/action text.
   - `梳化服提示`: youth/old-age state, pregnancy, injury, wet clothes, illness, formalwear, uniforms, plain/poverty looks, special styling.
   - `道具提示`: only practical on-screen props needed for the scene. Avoid false positives from dialogue, metaphors, or recalled events.
   - `备注/特殊道具/时间`: subtitles, VO, flashback, intimacy, action/stunt safety, continuity notes like `接15-1` or `和17-1一同拍摄`.
7. Style the workbook cleanly: freeze header row, filters on, light alternating rows, centered short columns, left-aligned longer text columns, landscape page setup.

## Script

Use `scripts/shunchangbiao_builder.py` as a reusable starting point.

Typical use:

```bash
python scripts/shunchangbiao_builder.py script1.docx script2.pdf --out-dir /path/to/output
```

Useful options:

- `--dump-json /path/to/scenes.json`: dump parsed scenes for review before workbook creation.
- `--config /path/to/config.json`: provide role order, role marks, and manually written scene summaries.

The script can generate a formatted draft, but do not rely on the draft summaries blindly. For production delivery, read the parsed scene bodies and revise summaries, props, costume, groups, and notes where the script is too literal.

## Required Checks

Before sending or finalizing, verify all of these:

- Workbook opens and has only the `顺场表` sheet.
- Scene count and scene IDs exactly match parsed script scenes.
- Any malformed headings discovered during parsing are included, not silently dropped.
- For PDFs, watermark/header/footer text is excluded from scene bodies, while real dialogue and headings remain.
- For PDFs with original repeated or skipped scene numbers, keep every scene in script order and record source-number anomalies in check notes or remarks.
- `日/夜`, `内/外`, and `页数` appear immediately after `拍摄内容` and before the role columns.
- `实际拍摄场地` and `页数` cells are blank when expected.
- Every `拍摄内容` cell is filled and 12-18 Chinese characters; rewrite any label-like short phrase.
- Role columns are split individually; no combined actor list cell.
- No extra `演员` / `角色` rows exist.
- `若干` groups are in `群演`, not role columns.
- Visual preview is nonblank and shows only role/日夜/内外/页数 headers vertically.
- Save the checked `.xlsx` output files and deliver them in the current agent conversation when attachments are supported; otherwise report their exact local paths. Do not add a third-party messaging-app delivery step to the default workflow.
