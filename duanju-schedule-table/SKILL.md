---
name: duanju-schedule-table
description: Create checked, single-sheet Excel 顺场表 / 顺畅表 / 拍摄顺序表 from short-drama scripts in .doc, .docx, or text-based .pdf files. Use for reviewed script-order schedules with individual-role columns, 特约演员, 群演, VO, OS, 梳化服, 道具, 戏用车, and 连场 marking.
---

# 短剧顺场表

Turn a complete short-drama script into a production-facing schedule table in script order. Keep this skill separate from AI asset breakdown: do not add asset variants, first-appearance deduplication, scene angles, image prompts, or image placeholders.

## Output Contract

- Create one `.xlsx` per script unless the user asks otherwise.
- Keep exactly one sheet named `顺场表`; do not add hidden or helper sheets.
- Keep one header row only. Do not add separate `演员` or `角色` rows.
- Keep this header order unchanged:
  `拍摄顺序`, `实际拍摄场地`, `提示`, `场次`, `剧本中场景`, `拍摄内容`, `日/夜`, `内/外`, `页数`, individual role columns, `特约演员`, `群演`, `VO`, `OS`, `梳化服提示`, `道具提示`, `戏用车`, `备注/特殊道具/时间`.
- Leave `实际拍摄场地` blank unless a confirmed practical location is provided.
- Keep `提示` as a horizontal, wrapped column between `实际拍摄场地` and `场次`. Fill it with concise, source-reviewed production tags such as `雨／烟机／闪回／武戏／AI做`, following the rules below. Leave it blank only when the scene has no applicable hint.
- Calculate `页数` from each scene's reviewed source-layout body line count: every 3 lines count as 0.1, rounding up any remainder. Display one decimal place; unreliable counts show `需确认` with a reason in notes.
- 默认直接按完整剧本整理演员：所有有姓名的剧情角色各占独立列，童年角色另列，其余无姓名的独立小角色放 `特约演员`。**不询问演员分列范围或要求用户确认名单后才继续。** 用户主动指定名单或增删时按其要求执行。每个独立角色用不重复的单字简称标记出镜。
- Put only role columns and `日/夜`, `内/外`, `页数` in vertical headers. Keep all other headers horizontal.
- Keep `拍摄顺序` in script order unless the user supplies a different shooting order.
- Outline each source-reviewed 连场 group in the `场次` column with one thick black outer border per contiguous run. Follow the continuity rules below; do not merge cells or add columns for this marking.

## Workflow

1. Resolve the directory containing this `SKILL.md`. Run all bundled scripts by absolute path so the workflow works from any Hermes workspace.
2. Normalize the complete input with [scripts/shunchangbiao_builder.py](scripts/shunchangbiao_builder.py):
   - Convert `.doc` to `.docx` automatically with `textutil` when available.
   - On Windows or Linux, use LibreOffice for `.doc` conversion when available.
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
7. Read every complete scene body and assemble the full named-character roster without a user-confirmation step; do not limit it to protagonists or a fixed number. Prepare a reviewed config using [review-config.md](references/review-config.md). Review main-location identity and ending-to-opening plot continuity together before recording `continuity_groups` (or `[]` after finding none). Set `reviewed: true` only after checking the source.
8. Run the final build with `--strict`. Do not deliver a workbook that fails strict validation.
9. Render the final sheet and visually inspect the populated range. Correct clipping, unreadable vertical headers, excessive row height, or broken printing before delivery.

## Field Rules

### 页数

- 从当前场次的正文第一行开始，统计到下一场场次标题之前；计入动作、环境描写及对白的正文行。按原剧本实际排版行数统计，跨页继续累计；不计场次标题、人物名单、空行、页眉页脚、页码和水印。
- **每3行计0.1，不足3行向上计：页数 = 向上取整（正文行数 ÷ 3）× 0.1。** 例如1–3行为0.1，4–6行为0.2，28–30行为1.0，31–33行为1.1，60行为2.0。不是按实际纸张页数，也不按四舍五入计算。
- 每场单独统计。已经复核拆成a/b/c的场次按各自正文分别计行，不重复使用母场总行数；连场加框不合并页数。
- 最终复核配置必须明确填写 `body_line_count`（正文行数正整数），Excel按公式生成数值并显示一位小数，如 `0.1`、`1.0`、`1.1`。不得用拍摄内容概括行数、Excel自动换行数或Word段落数替代正文排版行数。
- Word须对照原排版或保持原排版的PDF核行；解析器的段落数量不作为页数依据。PDF自动提取行数也只是草稿，须检查人物名单续行、跨页文字、页眉页脚、水印及场次边界，再填写复核行数。
- 没有正文或无法可靠核行时，`body_line_count` 填 `null`，页数显示“需确认”，备注写明原因及“页数需确认”；不默认填0.1或编造行数。严格校验检查填写及公式，不能代替原页核行。

### 提示

- 用来快速提醒本场的特殊拍摄条件、设备/效果需求、叙事形式或制作方式。写简短标签，不写剧情概括或长段备注；同场多项用 `／` 分隔，如 `雨／烟机／武戏`。这些是常用示例，不是封闭枚举。
- **雨**：正文明确本场在下雨或要拍雨景。不能仅凭“雨后”、湿地、人物名字或台词提到下雨就填写。
- **烟机**：明确需要烟机或人工烟雾效果。不能把普通抽烟、自然雾或蒸汽直接等同于已确定使用烟机；方式不明时在备注写明待确认。
- **闪回**：本场包含实际闪回/回忆画面。只在对白中回忆往事不自动标为闪回；具体闪回时段、接哪场写备注。
- **武戏**：本场存在武打、搏斗或需武行配合的动作内容。不能仅因争吵、威胁台词或提到“打”字就标武戏；动作和配合要求写备注。
- **AI做**：剧本、用户指示或已确认制作方案明确由AI生成/处理的画面。统一写 `AI做`，不因难拍、奇幻或大场面就自行决定用AI；仅部分镜头用AI时，备注标明具体范围。
- 逐场读完整正文并结合已确认制作要求填写 `hint`，相同需求在每个相关场次重复写。该字段在最终复核配置中必须明确提供，没有提示时填 `""`；自动草稿空白不代表已核实无需求。
- “提示”保留短标签，详细要求写“备注/特殊道具/时间”；与梳化服、道具、车辆有关的具体准备仍填写对应栏。不能把通用警告或所有备注整段搬进“提示”。

### 连场判定与场次框线

- **连场黑框必须同时满足：同一主场景，并且剧情时间、人物行动或同一事件紧接发展。** 同一宅院内由客厅追到走廊、门口可归一组；从宅院转到警局，即使紧接同一事件也必须断框。跨集同样逐场核定这两个条件。
- “主场景”是剧本中的同一地点主体，如同一宅院、商号、警局或戏班；其内部房间、走廊及明确附属的门口可以是不同子场景。不能仅因同属“城西”、都是仓库，或共用一个实际拍摄场地就合为同一主场景。先核同名别称和空间归属，在连场各成员的逐场复核项中填写统一的 `main_location`，不为此增加Excel列。
- 逐场对照前场结尾、后场开头以及中间叙事，确认时间和行动衔接。场号相邻、同一地点、同写“日/夜”、同服装，均不能单独作为连场依据；一般的服装/伤妆接戏要求也不等于连场。明确“次日”“几天后”等时间跳跃应断开。
- 复核后在 `continuity_groups` 中明确列出组内场号，`reason` 写明同一主场景以及剧情衔接的正文依据。换到另一个主场景时拆组；拆后不足两场的片段不单独算连场。没有连场时填 `[]`；主场景归属或剧情衔接不清的，在备注写“连场关系需确认”，暂不画确认框并继续制作。
- 同组在表格中连续排列的场次，**只在“场次”列（D列）画一个加粗纯黑外框**：首格上边、末格下边和各格左右边加粗；组内横向分隔线保持普通细线。场号保持原样、每场独立一行，不合并单元格，不加粗场号字体，不把框扩展到其他列。
- 中间插入别条故事线时，只有原线仍在同一主场景、同一连续时段才能保持同组；人物实际转去别处再返回，不因回到原主场景而重新串为一组。不改剧本顺序，不把无关场次圈入。将同组拆成各自相邻的框选段；同组至少两场但某成员单独落在一段时可单格成框。备注追加同组标识和前后关联场号，如“连场1；接1-3；续1-6”，并保留原有备注。用户另要分场表时，复用已核定关系，不能因同景归类后排在一起就新增连场。
- 相邻的不同连场组分别闭合外框。一个场次不重复归入多个组；跨线交叉关系不能可靠归为一组时保留待确认说明。
- 自动解析只出初稿；连场分组必须由完整正文复核确定。严格校验验证分组结构和复核记录，不能代替剧情理解。

### 拍摄内容

- Write the main action or event as `who + does what / what happens`.
- Keep normal scenes at 12-18 Chinese characters.
- Allow 4-18 characters for empty shots, montage, inserts, phone screens, and transitions when a shorter production description is clearer.
- Do not use the first line of dialogue, an environment-only description, or a mechanical source-text truncation as the final summary.

### 独立角色列、特约演员与群演

- 默认由制表者依据全剧正文复核并填写 `roles`：所有有姓名的剧情角色单列，包含具有明确个人身份的固定昵称及外文名；不能按戏份多少筛选，也不能套用上一部剧的名单或人数。童年/幼年角色与成年角色分列；同一身份的别名按正文统一，无法核实的保留原称谓和备注，不要求用户先选演员范围。
- 用户主动指定独立列范围时覆盖上述默认规则；`roles` 的“复核”指制表者核对剧本，不是用户批准。没有需独立列出的角色或用户明确要求不设时填 `[]`。仅职业、关系称谓且无姓名的角色，如班主、门房、某人之妻，按个人台词/对手戏规则放特约，不因角色重要就擅自编造姓名。
- 只在台词或照片中涉及的有姓名剧情角色也可列入名单，但不得因此标记本场实际出镜；备注说明仅提及、照片/影像素材需求。声音出演仍填写VO/OS。独立列名单与每场实际出镜是两件事。
- Apply the user's distinction to other on-camera cast: **一堆人、看不清脸、无单独台词 = 群演；单独角色，有个人台词/单独对手戏，不管几场 = 特约。** Personal dialogue OR an individual scene with another character is sufficient for `特约演员`; both are not required.
- Classify special actors versus crowds from each complete scene body and maintain the same identity across scenes. Never demote an individual to a crowd based on occupation, lack of a personal name, scene count, or a group keyword. A `保安`, `护士`, `服务员`, `村民甲`, or `舞女` with personal dialogue or individual interaction belongs in `特约演员`, unless already in the individual-column roster under the default rule or the user's explicit instructions.
- Use `群演` for background crowds with no individual dialogue or individual scene. If one person in a crowd speaks or has individual interaction, list that person in `特约演员` and only the remaining background crowd in `群演`. Distinguish identities using source labels such as `保安甲`; do not invent character names or headcounts.
- Fill `特约演员` with the full role labels, separated by `、`, in every scene where they appear. Keep the header horizontal and wrap the cell text. Leave the cell blank when none appear. Write performance requirements or unresolved identity questions in notes.
- Keep each on-camera cast identity in only one casting category: individual-role columns, `特约演员`, or `群演`. Voice appearances are a separate dimension and may coexist with an on-camera appearance. Incomplete evidence must be flagged `需确认` in notes rather than resolved by guessing.
- Explicitly fill per-scene `on_camera` with only members of `roles` visibly appearing in that scene. An empty list means none of the independent-column roles is on camera. Sound-only appearances go in `VO`/`OS` without an on-camera mark. Visible special actors still go in `special_actors`.
- For flashback, surveillance, screen playback, or video-call appearances, keep the role when new footage must be shot and state the appearance form in notes.
- Detect duplicate one-character role abbreviations and resolve them before delivery.
- Automated role/group keyword matches are draft candidates only. The reviewer must explicitly fill both `special_actors` and `groups`, including empty strings, after checking dialogue and action. The draft leaves `特约演员` blank until this review; a blank draft cell is not a finding that no special actors are needed.

### VO 与 OS

- **OS（Off-Screen）镜外音：声源就在这个场景空间里，只是镜头没拍到他，属于现场真实发生的对话。** 例如同一场景中未入镜的人在镜外接话。内心独白不归OS。
- **VO（Voice-Over）旁白/内心独白：声源不在当前现场时空，是后期叠加上去的声音，不是现场发生的对白。** 例如叙述性旁白或人物未说出口的心声。
- 分设两个横排表头列，位于“群演”后，顺序为 `VO`、`OS`。填写完整角色名，必要时括注声音形式；多角色以 `；` 分隔，没有则留空。
- 判断依据是剧本中声源所在时空和对白性质，不依据是否现场收音、后期补录，也不因为画面里看不到人物就一律填VO。电话声、广播声等须结合具体语境判断，不能仅凭名称自动分类；信息不足时在备注保留角色、原文标注和“VO/OS类型需确认”。
- 原剧本VO/OS标注与上述定义、正文情境冲突时，不默默互换。按明确情境暂列对应栏，并在备注写明“原标注＋冲突依据＋待导演/编剧确认”；情境也不明确时暂不强行归类。
- 同一角色本场既出镜又有VO/OS时，出镜栏和相应声音栏均记录；若同场确有两种声音形式，可同时填VO和OS。仅声音的特约角色只记声音栏，不作为本场出镜特约重复填写。
- 自动草稿只提取明确的VO/OS标签，保留大小写和括号写法中的类型，不根据关键词替代剧情判断。最终逐场复核 `vo`、`os` 和 `on_camera`；严格检查不能代替声源语境判断。

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
- Put weapons, dangerous items, breakaway/damage props, consumable duplicates, animals, mechanisms, and special-effect requirements in `备注/特殊道具/时间` as well. Put picture vehicles in `戏用车`; retain separately handled items such as keys, documents, or removable plates in props.

### 戏用车

- 独立列置于“道具提示”后、“备注/特殊道具/时间”前，横排表头，内容自动换行。
- 逐场填写需要准备的戏用车辆：车辆类型/明确身份＋原文明示数量＋使用方式或状态，如“轿车1辆（院外停放）”“卡车（行驶；车牌接1-3）”。机动车、摩托车、自行车、人力车、马车等按本场实际需求记录；轮椅、婴儿车等按行动道具处理。
- 记录停放、上下车、驾驶、追逐、撞损、年代外观、车牌、接戏等有依据的要求。不臆造车型、数量或颜色；仅台词提及而无需实物准备的车不列入。明确要拍摄的背景车辆仍须记录。
- 每场需要准备时重复填写，不能只记首次出现。撞车、特技、拖车拍摄、车辆改装或损坏备份等特殊要求同时在备注说明；不使用泛泛安全提示替代具体需求。
- 车辆草稿默认留空，最终必须逐场复核并显式填写 `vehicles`，没有需求时填空字符串。空白草稿不代表已确认无需车辆。

### 备注/特殊道具/时间

- Record subtitles, voice-type conflicts/uncertainty, flashback, surveillance/screen playback, intimacy, action safety, special vehicle requirements, animals, effects, breakaway copies, continuity, `接x-x`, and `一同拍摄`. Routine VO/OS cast and vehicle entries belong in their dedicated columns.
- Preserve unresolved day/night, inside/outside, location, or source-number issues as clear `需确认` notes.

## Script

Use [scripts/shunchangbiao_builder.py](scripts/shunchangbiao_builder.py) for parsing, validation, and workbook generation. It loads [scripts/dependency_bootstrap.py](scripts/dependency_bootstrap.py) and [scripts/requirements.txt](scripts/requirements.txt). On first run, missing packages are installed from PyPI into the user's LynnAI cache; the Hermes environment and system Python are not modified.

Install this complete skill bundle in Hermes with:

```bash
hermes skills install lynnhly1115-oss/lynnai-skills/duanju-schedule-table --yes
```

Use `python3` on macOS/Linux or `python` on Windows. Replace `<skill-dir>` with the installed directory reported by Hermes.

Draft parse:

```bash
python3 "<skill-dir>/scripts/shunchangbiao_builder.py" script.docx \
  --out-dir work \
  --dump-json work/scenes.json \
  --report work/validation.json
```

Final checked build:

```bash
python3 "<skill-dir>/scripts/shunchangbiao_builder.py" script.docx \
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
- Continuity groups are explicitly reviewed, have source-backed reasons for both same main location and continuous plot, contain at least two distinct existing scene IDs in source order, and do not overlap. Every member has a reviewed nonempty `main_location`, identical across the group. `continuity_groups: []` records a checked absence of continuity groups.
- Only scene-number cells receive thick black outer borders; internal separators remain thin, other columns and scene numbering remain intact, and intercut unrelated rows are excluded. Noncontiguous members carry cross-references in notes.
- Every scene is marked reviewed in strict mode.
- Every scene explicitly records a reviewed `body_line_count`; numeric pages follow the ceiling rule and display one decimal. Unknown counts display `需确认` with a reason and `页数需确认` in notes. Verify source-layout body lines rather than Word paragraphs or physical page counts.
- Every reviewed scene explicitly contains `hint`, including an empty string when no tag applies. Tags reflect actual scene/production requirements; `AI做` has an explicit production basis and specific partial-scene scope is preserved in notes.
- Every `拍摄内容` is meaningful and within the applicable length rule.
- Missing day/night or inside/outside is explicitly marked `需确认`.
- The full named-character roster and separate childhood roles are completed from the source without asking the user to approve an actor scope, unless the user has explicitly supplied a different roster. Special actors, crowds, and voice-only appearances are classified from dialogue and action without scene-count thresholds.
- Every reviewed scene explicitly contains `special_actors`; the final sheet contains exactly one horizontal `特约演员` column before `群演`. Cast identities are not duplicated across categories.
- Every reviewed scene explicitly contains `on_camera`, `vo`, `os`, and `vehicles`; the sheet has separate horizontal `VO`, `OS`, and `戏用车` columns in the locked positions. Individual-role appearance marks come from the reviewed `on_camera` list, allowing both visible and voice appearances without marking sound-only actors visible.
- OS is on-site off-screen dialogue; VO is narration/inner voice outside the current scene's spoken dialogue. Source-label conflicts are traceable in notes, and vehicle requirements are source-backed and repeated per scene.
- Role abbreviations are unique.
- `梳化服提示` is role-attributed and does not invent styling.
- `道具提示` contains practical scene needs, not background clutter or asset-deduplication logic.
- Workbook contains exactly one `顺场表` sheet and the locked header order.
- Header row and the first nine context columns through `页数` stay visible while navigating (`J2` freeze panes).
- Wrapped text is fully visible; row heights expand as needed.
- Filters, landscape printing, repeated print headers, and visual preview all work.
