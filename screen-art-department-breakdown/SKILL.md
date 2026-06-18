---
name: screen-art-department-breakdown
description: Create film/TV/short-drama art department workbooks from .doc, .docx, or text-based .pdf scripts. Use when the user asks for 美术指导, 美术拆解, 顺场表, 分场表, 场景表, 场景道具汇总, 置景道具执行表, 采买制作表, or art-department prep for a screenplay.
---

# 影视美术拆解

Use this skill to turn a script into the user's preferred art department prep workbook. Keep the output practical and uncluttered.

## Output Workbook

Create one `.xlsx` per script unless the user asks otherwise. The workbook should contain exactly these four sheets:

1. `顺场表`
   - One scene per row in script order.
   - Required columns: `序号`, `场次`, `主场景`, `次场景`, `剧本中场景`, `日/夜`, `内/外`, `人物/群演`, `拍摄内容`, `陈设道具`, `戏用道具`, `车辆/特殊道具`, `备注`.
   - `拍摄内容` should summarize the scene action for production use.

2. `分场表`
   - Based on the same rows as `顺场表`, but sorted by `主场景` -> `次场景` -> original scene order.
   - Use this to see all scenes under the same location together.
   - Required columns: `主场景`, `次场景`, `场次`, `剧本中场景`, `日/夜`, `内/外`, `人物/群演`, `拍摄内容`, `陈设道具`, `戏用道具`, `车辆/特殊道具`, `备注`.
   - Vertically merge repeated `主场景` cells for consecutive rows in the same main scene. Do not repeat the main scene name on every row.

3. `场景表`
   - This is the clean location summary table. Do not put unrelated risk, style, or department chatter here.
   - One row per `主场景`.
   - Required columns: `主场景`, `包含次场景`, `涉及场次`, `陈设道具`, `戏用道具`, `车辆/特殊道具`, `备注`.
   - Highlight `车辆/特殊道具` in red whenever it contains vehicles, animals/live creatures, dangerous items, breakaway/damage props, fire/water effects, or other special handling.

4. `置景道具执行表`
   - This is the execution/prep list for art, props, purchase, fabrication, rental, and borrowing.
   - Required columns: `主场景`, `次场景`, `需要布置内容`, `陈设道具`, `戏用道具`, `需采购`, `需制作`, `需租赁/借用`, `车辆/特殊道具`, `备注`.
   - Highlight `车辆/特殊道具` in red using the same rule as `场景表`.
   - Sort by `主场景` -> `次场景`, then vertically merge repeated `主场景` cells.

## Scene Grouping

- `顺场表` keeps script order.
- `分场表` is based on `顺场表`, but sorted by scene.
- `主场景` is the scouting/location unit, such as `张家`, `村委会`, `医院`, `福星桥`.
- `次场景` is the functional space inside the main scene, such as `客厅`, `厨房`, `卧室`, `院子`, `门口`, `会议室`.
- Normalize `张家客厅`, `张家厨房`, `张家卧室`, `张家门口` under the same `张家`.
- Keep `主场景` clean. Do not put detailed space words in it when they can be moved to `次场景`: `门口`, `门外`, `院子`, `院内`, `客厅`, `厨房`, `卧室`, `会议室`, `办公室`, `病房`.
- Example: `刘老头家门口`, `刘老头家门外树下`, and `刘老头家外` all use `主场景=刘老头家`; their `次场景` values are `门口`, `门外树下`, and `外景/院外`.
- Example: `村委会会议室` uses `主场景=村委会`, `次场景=会议室`; `纪委办公室` uses `主场景=纪委`, `次场景=办公室`.
- Do not use vague subscene names like `整体`, `整体内景`, `整体外景`, or `外景整体`. If the script does not name a smaller space, write a natural specific label such as `路口外景`, `宿舍内景`, `家中内景`, `桥上/桥面`, or `办公室`.
- If two scene headings may or may not be the same practical location, keep them separate and mark `需确认是否同景`.

## Prop Rules

- Separate `陈设道具` from `戏用道具`.
  - `陈设道具`: furniture, set dressing, wall/floor/window treatment, signs, background objects, atmosphere objects.
  - `戏用道具`: objects actors handle, use, exchange, read, damage, carry, or that receive close-up/story attention.
- Do not merge all repeated generic props. Split by owner/use when the script supports it: `陈星尧手机`, `王泽手机`, `直播拍摄手机`, `手机屏幕画面`, `张峥包袋`.
- Do not output rough placeholders like `待定手机` or `待定`. Infer ownership from action/context whenever possible. If truly unclear, write a human-readable note such as `未明确归属手机（需确认）`.
- Do not leave `陈设道具` as `待定`. If the script does not specify set dressing, infer basic practical dressing from the scene type, such as home furniture/life clutter, village committee meeting setup, village road signs, courtyard tools, hospital/clinic equipment, or use `按实景基础陈设`.
- Treat graphic/screen props as production items: phone screens, live-stream interfaces, contracts, files, signs, certificates, photos, ledgers, notices.
- For live animals/creatures, write the specific animal, such as `鸡`, `鸭子`, `狗`, instead of a vague `活物/动物`.
- Do not infer live animals from character names or metaphors. `李狗娃` and `阿猫阿狗` are not animal props.
- In `场景表` and `置景道具执行表`, mark these in red under `车辆/特殊道具`: vehicles, specific live animals/creatures, dangerous items, breakaway props, damage props, fire/water effects, medical/professional props requiring review, and unusual large props.

## Source Handling

- For `.docx`, read paragraphs directly.
- For `.doc`, convert or extract text with `textutil` when available.
- For text-based `.pdf`, extract text with `pdfplumber`; filter watermark text by PDF attributes when possible.
- For scanned/image-only PDFs, use OCR if available; otherwise ask for a Word/text export.

## Script

Use `scripts/art_department_breakdown.py` as the reusable starting point.

```bash
python scripts/art_department_breakdown.py script.docx script.pdf --out-dir /path/to/output
```

If the default `python` does not have `pdfplumber`, `python-docx`, or `openpyxl`, use the bundled workspace Python from `load_workspace_dependencies`.

Useful options:

- `--dump-json /path/to/scenes.json`: dump parsed scenes for checking.
- `--title 项目名`: override the workbook title/output name when processing one script.

## Required Checks

- Workbook opens and contains exactly `顺场表`, `分场表`, `场景表`, `置景道具执行表`.
- `顺场表` row count equals parsed scene count.
- `分场表` has the same scene count as `顺场表`, but sorted by `主场景` and `次场景`.
- `场景表` has one row per `主场景` and no unrelated extra columns.
- `主场景` values are clean grouping units. Check that they do not unnecessarily include subspace words like `门口`, `院内`, `厨房`, `会议室`, or `办公室`.
- `次场景` values are specific and human-readable; they should not contain `整体`, `整体内景`, `整体外景`, or `外景整体`.
- In `分场表` and `置景道具执行表`, repeated consecutive `主场景` cells are merged vertically so the main scene name appears once per group.
- `车辆/特殊道具` cells are red when they contain vehicles, animals/live creatures, dangerous/damage props, or other special handling.
- No sheet includes `源页` / page columns unless the user explicitly asks.
- Final tables should not contain bare placeholders like `待定` or `待定手机`; use inferred scene dressing or clear `需确认` notes.
- Live animals are specific and source-backed; check that character names/metaphors did not trigger animal props.
- Scene IDs and scene order in `顺场表` match the script.
- Malformed headings such as `一组镜头`, `手机屏幕`, `闪回`, or `番外` are included.
- Save the checked `.xlsx` output path. Do not send files to WeChat or any chat app unless the user explicitly asks in the current task and provides the destination.
