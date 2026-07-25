---
name: ai-live-action-asset-breakdown
description: Read complete AI live-action film, TV, or short-drama scripts and create a single-sheet Excel inventory of first-appearance character, scene, and actor-used prop assets. Use for AI真人剧资产统计, 人物场景道具表, 按集累计去重, 资产变体, or a workbook with reserved image cells from DOCX, DOC, PDF, scanned PDF, or plain-text scripts.
---

# AI真人剧资产统计

Read the complete script supplied in the current request, extract visible production assets, cumulatively deduplicate them across episodes, and deliver one Excel workbook per project. Do not generate images unless separately requested.

## Source Handling

- Treat the complete script provided in the current request as the sole source of truth.
- Start a new workbook for every different project.
- Rebuild from the current complete script; do not append stale data from an older script version.
- Support DOCX, DOC, text-based PDF, scanned PDF, and plain text.
- Use the document or PDF skill appropriate to the source. Apply OCR to scanned PDFs when available; report a blocker if reliable text cannot be recovered.
- Read all episodes before extracting assets.

## Asset Rules

### First appearance and deduplication

- List an asset only in the episode and scene where it first becomes visible.
- Omit identical reuse from all later episodes.
- Deduplicate repeated appearances inside the same episode.
- Exclude characters, scenes, and props that are only mentioned and never shown.
- Count a visible face or design shown through a monitor, hologram, photograph, reflection, or mirror.
- Exclude voice-only appearances and indistinct silhouettes until a usable design becomes visible.

### Variants

- Create a separate variant row for a meaningful visual change in clothing, age, disguise, injury, blood, damage, day/night, power state, destruction state, screen content, function, or activation.
- Attach visible effects to the affected asset: character effects become character variants, prop effects become prop variants, and environment effects become scene variants.
- If a base asset and its variant first appear in the same scene, list both rows.
- Count every base row and every variant row in episode subtotals.
- Do not create variants for expression, ordinary pose/action, camera distance, camera angle, object placement, minor fabric wrinkles, or ordinary dirt.

### Characters

- Include principal cast and every visually distinct person with dialogue, foreground action, close-up, interaction, pursuit, or combat.
- Give unnamed foreground roles clear names such as `治安局汇报队员`, `夜班工人甲`, and `值守稽查员`.
- Exclude undifferentiated background crowds used only to fill the frame.
- Treat age, clothing, disguise, injury, body modification damage, and activated on-body effects as variants of the character.
- Keep prosthetics and permanent body modifications under the character category unless they appear independently as actor-used props.

### Scenes

- Count every visually distinct playable space independently.
- Split functional spaces inside a large location when they need separate design, such as a front gate, security checkpoint, workshop, and inner factory.
- Show hierarchy in `资产名称` without adding a separate column: `大场景·具体空间`.
- Keep different camera angles in the same scene row and use its multiple image cells.
- Treat day/night, intact/destroyed, powered/blackout, or other substantial environmental changes as scene variants.
- Put furniture, pipes, lighting fixtures, fixed machinery, background crowds, and other set dressing in the scene description rather than the prop category.

### Props

- Include only actor-used props: objects that a character handles, carries, exchanges, reads, views, activates, damages, or uses in action.
- Treat an interacted-with screen, terminal, weapon, vehicle, or device as a prop.
- Keep fixed, non-interacted equipment and ordinary dressing inside the scene description.
- Reuse the same prop when appearance and function are identical, even if another character uses it.
- Separate props when owner, appearance, function, or continuity is materially different.

### Descriptions

- Use only visual or functional information supported by the script.
- Do not invent casting, clothing, materials, colors, or design details during the inventory stage.
- Write `剧本未明确` when required information is absent.
- Report source contradictions in the final chat message, not as a workbook column.

## Workbook Contract

Create exactly one worksheet named `资产统计表` with these columns:

`集数｜场次｜类别｜资产名称｜基础/变体｜母资产｜资产说明｜图片1｜图片2｜图片3｜图片4｜图片5｜图片6`

- Do not add asset IDs, continuity columns, or image-prompt columns.
- Order the workbook by episode. Within each episode, group `人物`, then `场景`, then `道具`; sort each group by first scene.
- Add a formula-driven `本集小计` row after each episode.
- Add a formula-driven cumulative total row at the end.
- Reserve image cells 1-3 for characters and props; shade image cells 4-6 gray for those rows.
- Reserve all six image cells for scenes.
- Keep image cells empty and large enough for later insertion.
- Freeze the header and left information columns, enable filters, wrap text, and use the fixed burgundy/gold workbook style.
- Name the file `《剧本名》_AI资产统计表.xlsx`.

## Build the Workbook

Read [references/inventory-schema.md](references/inventory-schema.md), then create the intermediate JSON exactly as specified.

Use the bundled workspace Node.js and `@oai/artifact-tool`. Copy the builder into a writable task directory before running it so Node can resolve the workspace dependency symlink:

```bash
TASK_WORK=/absolute/writable/task-directory
mkdir -p "$TASK_WORK"
cp /absolute/path/to/this-skill/scripts/build_asset_workbook.mjs "$TASK_WORK/build_asset_workbook.mjs"
ln -s /workspace/dependencies/node/node_modules "$TASK_WORK/node_modules"
/workspace/dependencies/node/bin/node "$TASK_WORK/build_asset_workbook.mjs" \
  --input "$TASK_WORK/assets.json" \
  --output "/absolute/output/《剧本名》_AI资产统计表.xlsx" \
  --preview-dir "$TASK_WORK/previews"
```

Use the exact runtime paths returned by `load_workspace_dependencies`; do not use system Node or another spreadsheet library.

## Quality Gate

- Every row has an earliest episode and scene.
- No unchanged cross-episode reuse remains.
- Every variant names a valid mother asset and explains the visible difference.
- No mention-only content or pure background crowd is counted.
- No set dressing is misclassified as a prop.
- Per-episode and cumulative formulas equal the listed asset rows.
- The workbook contains one sheet only and uses the exact headers.
- Inspect values and formulas, scan for formula errors, and visually review every rendered preview segment.
- Deliver only the final XLSX; do not deliver JSON, scripts, or previews.
