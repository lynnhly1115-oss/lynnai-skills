---
name: screen-art-department-breakdown
description: Create checked art-department prep folders from film, TV, or short-drama scripts in .doc, .docx, or text-based .pdf format. Use when the user asks for 美术指导, 美术拆解, 顺场表, 分场表, 场景表, 场景道具汇总, 置景道具执行表, 采买制作表, or art-department prep for a screenplay.
---

# 影视美术拆解

Turn a complete screenplay into a checked art-department prep folder. The workflow must work across genres, periods, countries, episode structures, and production scales. Do not assume that the next project is a modern short drama, a rural story, or similar to a previous script.

## Core Method

Use three layers in this order:

1. **Universal rules**: scene parsing, evidence handling, naming, prop use, continuity, sourcing, and validation.
2. **Project vocabulary**: learn recurring locations, aliases, characters, organizations, objects, and state terms from the complete current script.
3. **Optional project profile**: period, region, genre, production scale, and special technical needs may refine results only when the source or user supports them.

Source evidence always overrides a profile. Never add rural, urban, historical, fantasy, science-fiction, crime, medical, military, school, workplace, or other genre-specific dressing merely because it is common in that genre.

## Output Folder

Create one project folder containing exactly four independent `.xlsx` files. Do not combine them as four sheets in one workbook.

1. `01-顺场表.xlsx`
   - One scene per row in source order.
   - Required columns: `序号`, `场次`, `剧本原场景名`, `归并主场景`, `功能区/次场景`, `日/夜`, `内/外`, `人物/群演`, `美术相关剧情/画面要求`, `场景状态/连续性`, `陈设道具`, `戏用道具`, `特殊/安全项`, `待确认`.

2. `02-分场表.xlsx`
   - Use the same scene records as `顺场表`, sorted by `归并主场景` -> `功能区/次场景` -> original scene order.
   - Required columns: `归并主场景`, `功能区/次场景`, `场次`, `剧本原场景名`, `日/夜`, `内/外`, `人物/群演`, `美术相关剧情/画面要求`, `场景状态/连续性`, `陈设道具`, `戏用道具`, `特殊/安全项`, `待确认`.

3. `03-场景表.xlsx`
   - One row per standardized main location.
   - Required columns: `主场景`, `包含功能区/次场景`, `场次数`, `涉及场次`, `核心置景要求`, `空间状态/版本`, `主要陈设道具`, `主要戏用道具`, `特殊/安全项`, `待确认`.

4. `04-置景道具执行表.xlsx`
   - One row per standardized main-location and functional-space combination.
   - Required columns: `主场景`, `功能区/次场景`, `涉及场次`, `空间处理/改造要求`, `陈设道具`, `戏用道具`, `需采购`, `需制作`, `需租赁/借用`, `特殊/安全确认`, `数量/版本/连续性`, `待确认`, `负责人`, `执行状态`.

Keep one sheet in each file, use filters and frozen panes, and avoid merged cells so production teams can sort, copy, and filter safely.

## Source Calibration

Read the complete script before finalizing the tables.

- For `.docx` and `.docm`, read both ordinary paragraphs and text inside tables in document order.
- For legacy `.doc`, convert or extract text with `textutil` or LibreOffice when available.
- For text-based `.pdf`, extract text with `pdfplumber`; remove only confirmed repeated headers, footers, or watermarks.
- For scanned or image-only files, use OCR when available; otherwise ask for a Word, text, or OCR-ready export.
- Support Chinese and international screenplay headings, numbered or unnumbered scenes, optional episode numbers, Chinese numerals, `INT./EXT.`, day/night markers in different positions, and scripts whose headings are stored in tables.
- Include flashbacks, memories, dreams, montages, inserts, phone or computer screens, split scenes, pre-scenes, extra scenes, and alternate timelines when they function as visible production units.
- If a line looks like a scene heading but cannot be parsed, add a validation diagnostic. Never silently drop it.

For scripts without scene numbers, assign stable document-order scene IDs for the output and state that assumption in delivery notes.

## Evidence Rules

Every important decision belongs to one of three evidence levels:

- **Explicit**: directly stated in the scene heading, action, dialogue, or production note.
- **Strong inference**: required by visible action or by a clearly identified practical location.
- **Unresolved**: plausible but not supported strongly enough.

Use explicit evidence first. Add strong inferences only when necessary for practical preparation. Put unresolved items in `待确认` with a plain-language reason. Do not present a guess as a script fact.

Keep source traceability in the parsed JSON: scene, location, day/night, and interior/exterior should record whether they came from explicit text, inference, or document sequence.

## Standard Naming

All four files must be generated from the same standardized scene and prop records.

- Preserve the source wording only in `剧本原场景名`.
- Model a location as `归并主场景 + 功能区/次场景 + 状态/版本`.
- `归并主场景` is the practical scouting or build unit; `功能区/次场景` is the space within it.
- Learn aliases from the current script. Normalize only when the text supports that two names represent the same practical location.
- If identity is uncertain, keep locations separate and add `需确认是否同景`.
- Do not use vague subspaces such as `整体`, `整体内景`, `整体外景`, or `外景整体` when a clearer functional label is available.

Standardize a prop by `基础物件 + 所属人/用途 + 状态/版本` when those distinctions matter. For example, separate different characters' phones, a phone used for live streaming, and a phone-screen graphic. The same standardized prop name must be reused in every file.

## Art And Prop Rules

Classify an object by what it does in the scene, not by a permanently fixed genre list.

- `陈设道具`: set dressing, furniture, surfaces, signage, atmosphere objects, and background objects that define the visible space.
- `戏用道具`: objects a performer handles, carries, exchanges, reads, damages, activates, or that receive close-up or story attention.
- `图文/屏幕内容`: letters, files, contracts, signs, certificates, photos, maps, newspapers, screens, interfaces, labels, and other readable graphics. Include their required visible content when known.
- `特殊/安全项`: vehicles, live creatures, weapons, breakaway or damage props, fire, smoke, water, blood, stunt interactions, hazardous materials, specialist medical or technical objects, or anything requiring special handling.

The same base object may need different treatment in different scenes. Prefer visible action and concrete object wording; do not trigger a prop from a character name, metaphor, or vague dialogue reference.

Do not use bare placeholders such as `待定` or `待定手机`. If an owner, quantity, content, or version is genuinely unclear, use a readable standard name and put the uncertainty in `待确认`.

Infer basic dressing only from the identified space and visible action, and keep it era-neutral unless the script establishes a period. Do not import the previous project's rural, urban, period, or stylistic assumptions.

## State And Continuity

Track visible changes that affect art preparation, including:

- before/after, clean/dirty, intact/damaged, dry/wet, normal/burned or smoke-affected;
- blood, stains, debris, renovation, construction, aging, weather, seasonal or night-light changes;
- dream, memory, flashback, alternate timeline, historical period, or other time-state versions;
- readable graphic or screen-content versions;
- hero, stunt, breakaway, duplicate, reset, consumable, or continuity copies.

Record these in `场景状态/连续性`, `空间状态/版本`, or `数量/版本/连续性`. Never collapse visually different states into one undifferentiated item.

## Sourcing Rules

Classify preparation by production attributes rather than genre keywords:

- `需采购`: consumables, ordinary replaceable supplies, and low-risk practical dressing.
- `需制作`: custom graphics, signage, breakaway/damage versions, story-specific fabrication, and custom finishes.
- `需租赁/借用`: vehicles, live creatures, specialist equipment, electronics, large or expensive items, and objects requiring licensed or controlled access.

One item may appear in more than one bucket when separate versions or preparation steps are required. State the reason in the execution fields or `待确认`.

## Script Helper

Use `scripts/art_department_breakdown.py` as the reusable builder.

```bash
python3 scripts/art_department_breakdown.py script.docx \
  --out-dir "$HOME/Desktop" \
  --folder-name "项目名-美术拆解" \
  --dump-json /path/to/check/scenes.json
```

The first run installs missing `pdfplumber`, `python-docx`, and `openpyxl` packages into the user's LynnAI cache. It does not modify the host agent, Hermes, or system Python. Internet access is needed only when packages are missing.

## Mandatory Validation

Do not deliver the first generated draft. Validation is a required part of this Skill, not an optional suggestion.

### 1. Source And Parse Check

- Compare scene count, scene IDs, and source order with the complete script.
- Review parser diagnostics and every malformed or nonstandard heading.
- Confirm day/night, interior/exterior, people, and original location text did not shift into the wrong fields.
- Check that flashbacks, montages, inserts, screens, and extra scenes were not silently dropped.
- Inspect the JSON scene dump and note any source-number gap, duplicate, or generated sequence ID.

### 2. Cross-File Data Check

- Confirm the folder contains exactly the four required `.xlsx` files.
- Confirm each file has exactly one correctly named sheet and the exact required headers.
- Confirm `顺场表` and `分场表` contain the same scenes exactly once; only their order may differ.
- Confirm the standardized main-location set is identical across all four files.
- Confirm every summarized or execution prop comes from the standardized prop library used by `顺场表`.
- Confirm source names appear only in `剧本原场景名`; standardized names are reused everywhere else.
- Reject bare `待定`, `待定手机`, unexplained certainty, or silent location merging.

### 3. Semantic Art Check

- Check each scene's visible art requirement against its action, not just keyword matches.
- Check set dressing versus actor-used props, graphic content, special handling, and safety classification.
- Check location and prop state changes, quantities, duplicate versions, resets, consumables, and continuity needs.
- Check purchase, fabrication, rental, and borrowing classifications by actual production attributes.
- Check that no genre-specific assumption was introduced without current-script evidence.
- Read every `待确认` item and make sure it is specific enough for a producer or art director to resolve.

### 4. Workbook And Visual Check

- Reopen all four files after saving.
- Confirm filters, frozen panes, column widths, row heights, wrapping, highlights, dropdowns, and print readability.
- Confirm there are no merged cells, clipped headers, unreadable long text, broken formulas, or overlapping content.
- Highlight special/safety fields in red and unresolved/state/continuity attention fields in yellow.
- Open or render all four final files and visually inspect real rows before reporting success.

The helper's automatic validation must pass, and the semantic and visual checks must also pass. If any check fails, fix the Skill output and rerun validation before delivery.

## Delivery Contract

For every completed run, report:

- source script name;
- final project-folder path and all four workbook paths;
- parsed scene count;
- assumptions, parser diagnostics, and source-number anomalies;
- automatic, semantic, and visual validation status.

Keep JSON dumps, screenshots, and temporary review files outside the final four-file folder. Do not send files to WeChat, Slack, email, or any chat app unless the user explicitly asks in the current task and gives the destination.
