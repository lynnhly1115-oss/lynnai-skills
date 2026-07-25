---
name: duanju-schedule-table
description: Create checked, single-sheet Excel 顺场表 / 顺畅表 / 拍摄顺序表 from short-drama scripts in .doc, .docx, or text-based .pdf files. Use when the user asks to arrange scenes in script order and fill 场次, 场景, 拍摄内容, individual role columns, 日夜内外, 群演, 梳化服, 道具, and production notes in a reviewed .xlsx workbook.
---

# 短剧顺场表

Turn a complete short-drama script into a production-facing schedule table in script order. Keep this skill separate from AI asset breakdown: do not add asset variants, first-appearance deduplication, scene angles, image prompts, or image placeholders.

## Output Contract

- Create one `.xlsx` per script unless the user asks otherwise.
- Keep exactly one sheet named `顺场表`; do not add hidden or helper sheets.
- Keep one header row only. Do not add separate `演员` or `角色` rows.
- Keep this header order unchanged:
  `拍摄顺序`, `实际拍摄场地`, `场次`, `剧本中场景`, `拍摄内容`, `日/夜`, `内/外`, `页数`, role columns, `群演`, `梳化服提示`, `道具提示`, `备注/特殊道具/时间`.
- Leave `实际拍摄场地` blank unless a confirmed practical location is provided.
- Leave `页数` blank unless reliable page counts are provided.
- Create one column per named role. Mark an appearance with a unique one-character abbreviation.
- Put only role columns and `日/夜`, `内/外`, `页数` in vertical headers. Keep all other headers horizontal.
- Keep `拍摄顺序` in script order unless the user supplies a different shooting order.

## Workflow

1. Use the Documents, PDF, and Spreadsheets skills for source handling and final workbook QA.
2. Normalize the complete input:
   - Convert `.doc` to `.docx` automatically with `textutil` when available.
   - Read `.docx` paragraphs and table cells in document order; include text-box text when extractable.
   - Extract text-based PDFs with `pdfplumber`. Filter repeated watermarks conservatively; never discard a font or size globally without checking real text.
   - Use OCR for scanned PDFs when available. Otherwise ask for a Word or text export.
3. Treat every explicit scene-number prefix such as `1-1`, `19-2`, or `80-6` as a hard scene boundary before parsing other fields.
4. Support incomplete and irregular headings, including:
   - `1-1空镜`;
   - `2-2 沈家旧宅 夜 外（闪回 同3-1）`;
   - missing day or inside/outside;
   - full-width dashes, missing spaces, montage, insert, phone-screen, transition, flashback, and extra episodes.
5. Preserve missing values as `需确认` or make a nearby-scene inference with an explicit note. Never drop the scene.
6. Run the parser as a draft and dump complete scene data for review.
7. Read every complete scene body. Prepare a reviewed config using [review-config.md](references/review-config.md). Set `reviewed: true` only after checking the scene against the source.
8. Run the final build with `--strict`. Do not deliver a workbook that fails strict validation.
9. Render the final sheet and visually inspect the populated range. Correct clipping, unreadable vertical headers, excessive row height, or broken printing before delivery.

## Field Rules

### 拍摄内容

- Write the main action or event as `who + does what / what happens`.
- Keep normal scenes at 12-18 Chinese characters.
- Allow 4-18 characters for empty shots, montage, inserts, phone screens, and transitions when a shorter production description is clearer.
- Do not use the first line of dialogue, an environment-only description, or a mechanical source-text truncation as the final summary.

### 角色与群演

- Put named, on-camera roles in individual role columns.
- Put collective labels such as `若干`, `众人`, `队员众`, `舞女`, `劳工`, and similar groups in `群演`.
- Put `VO`, `OS`, phone voices, and broadcast voices in notes, not actor columns.
- For flashback, surveillance, screen playback, or video-call appearances, keep the role when new footage must be shot and state the appearance form in notes.
- Detect duplicate one-character role abbreviations and resolve them before delivery.

### 梳化服提示

- Use the format `角色：服装＋发型＋妆效＋特殊状态/连戏要求`.
- Record only source-backed or production-required information: specific clothing, uniforms, hair changes, injury/blood/illness makeup, dirt, wetness, pregnancy, disguise, age state, damaged clothing, prosthetic state, and continuity.
- Write continuity explicitly, such as `接3-3同服` or `伤口位置保持`.
- Leave the cell blank when the script gives no special requirement and continuity does not require a note. Do not invent costume design.
- Separate multiple roles with `；`.

### 道具提示

- Record props that actors carry, handle, exchange, use, damage, consume, read, or that receive close-up/story attention.
- Include graphic and screen props such as files, photos, certificates, phone screens, surveillance images, and interface content.
- Exclude dialogue-only mentions, metaphors, and ordinary background furniture with no action or story function.
- Repeat a required prop in every scene where it must be prepared. Do not apply asset-style first-appearance deduplication.
- Put weapons, dangerous items, breakaway/damage props, consumable duplicates, vehicles, animals, mechanisms, and special-effect requirements in `备注/特殊道具/时间` as well.

### 备注/特殊道具/时间

- Record subtitles, VO/OS, flashback, surveillance/screen playback, intimacy, action safety, vehicles, animals, effects, breakaway copies, continuity, `接x-x`, and `一同拍摄`.
- Preserve unresolved day/night, inside/outside, location, or source-number issues as clear `需确认` notes.

## Script

Use `scripts/shunchangbiao_builder.py` for parsing, validation, and workbook generation.

Draft parse:

```bash
python scripts/shunchangbiao_builder.py script.docx \
  --out-dir work \
  --dump-json work/scenes.json \
  --report work/validation.json
```

Final checked build:

```bash
python scripts/shunchangbiao_builder.py script.docx \
  --out-dir output \
  --config reviewed.json \
  --report output/validation.json \
  --strict
```

The automatic fields are drafts. Final delivery requires complete source review and a strict build.

## Required Checks

- Source scene count equals workbook data-row count.
- Scene IDs are complete, unique, and in source order; malformed headings are not silently dropped.
- No scene body absorbs text from a following scene.
- Every scene is marked reviewed in strict mode.
- Every `拍摄内容` is meaningful and within the applicable length rule.
- Missing day/night or inside/outside is explicitly marked `需确认`.
- Named roles, groups, and voice-only roles are classified correctly.
- Role abbreviations are unique.
- `梳化服提示` is role-attributed and does not invent styling.
- `道具提示` contains practical scene needs, not background clutter or asset-deduplication logic.
- Workbook contains exactly one `顺场表` sheet and the locked header order.
- Header row and the first eight context columns stay visible while navigating.
- Wrapped text is fully visible; row heights expand as needed.
- Filters, landscape printing, repeated print headers, and visual preview all work.
