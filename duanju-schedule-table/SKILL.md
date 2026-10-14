---
name: duanju-schedule-table
description: Create checked Excel workbooks with 顺场表 / 顺畅表 / 拍摄顺序表 and a matching location-grouped 分场表 from scripts in .doc, .docx, or text-based .pdf files. Review and split internal location changes into a/b/c scenes, with individual-role columns, 特约演员, 群演, VO, OS, 梳化服, 道具, 戏用车, and 连场 marking.
---

# 剧本顺场表与分场表

Turn a complete screenplay into two production-facing views of one reviewed data set: script order and location grouping. The rules apply across scripts; do not reuse a previous project's cast, locations, or scene groups. Keep this skill separate from AI asset breakdown: do not add asset variants, first-appearance deduplication, scene angles, image prompts, or image placeholders.

## Output Contract

- Create one `.xlsx` per script unless the user asks otherwise.
- 默认每个工作簿包含两张可见表，顺序固定为 `顺场表`、`分场表`。先完成并复核顺场数据，再由同一数据生成分场表；两表表头及全部场次内容一致，只改变排列顺序。不加隐藏表或辅助表。
- Keep one header row on each sheet only. Do not add separate `演员` or `角色` rows.
- Keep this header order unchanged:
  `拍摄顺序`, `实际拍摄场地`, `提示`, `场次`, `剧本中场景`, `拍摄内容`, `日/夜`, `内/外`, `页数`, individual role columns, `特约演员`, `群演`, `VO`, `OS`, `梳化服提示`, `道具提示`, `戏用车`, `备注/特殊道具/时间`.
- Leave `实际拍摄场地` blank unless a confirmed practical location is provided.
- Keep `提示` as a horizontal, wrapped column between `实际拍摄场地` and `场次`. Fill it with concise, source-reviewed production tags such as `雨／烟机／闪回／武戏／AI做`, following the rules below. Leave it blank only when the scene has no applicable hint.
- Calculate `页数` from each scene's reviewed source-layout body line count: every 3 lines count as 0.1, rounding up any remainder. Display one decimal place; unreliable counts show `需确认` with a reason in notes.
- 默认直接按完整剧本整理演员：所有有姓名的剧情角色各占独立列，童年角色另列，其余无姓名的独立小角色放 `特约演员`。**不询问演员分列范围或要求用户确认名单后才继续。** 用户主动指定名单或增删时按其要求执行。每个独立角色用不重复的单字简称标记出镜。
- Put only role columns and `日/夜`, `内/外`, `页数` in vertical headers. Keep all other headers horizontal.
- Keep `拍摄顺序` in script order unless the user supplies a different shooting order.
- 主动逐场检查正文里的实际画面转场。同一原场号包含多个独立拍摄空间时，按正文先后拆成 `1-1a、1-1b、1-1c` 等子场；两表均使用拆分后的同一数据，不等用户逐一指出。
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
7. Read every complete source scene body and locate actual changes of filming space, including mid-paragraph transitions. Record reviewed `scene_splits` using [review-config.md](references/review-config.md); use `{}` only after checking that no splits are needed. Apply these boundaries before reviewing each resulting child's main location, cast, action, departments and page lines. Assemble the full named-character roster without a user-confirmation step; do not limit it to protagonists or a fixed number. Review main-location identity and ending-to-opening plot continuity together before recording `continuity_groups` (or `[]` after finding none). Set `reviewed: true` only after checking the source.
8. Run the final build with `--strict`. It applies reviewed splits before validating and generating `顺场表` and `分场表`. Compare both sheets by child scene ID and verify identical data, cast marks, page totals and coverage. Reconcile original scene count, split mapping and final row count; no parent body may be dropped or repeated. Do not deliver a workbook that fails strict validation.
9. Render both final sheets and visually inspect the populated range. Correct clipping, unreadable vertical headers, excessive row height, or broken printing before delivery.

## Field Rules

### 场内转场与a/b/c拆场

- **同一个原场号内，画面实际转到另一个独立拍摄空间，就单独分行。** 同一主场景的客厅转院子、异地主线交叉、切到另一地点的闪回画面，都逐段核对实际画面；不能因为仍在一个宅院里就放在一行。
- 原场 `1-1` 拆两段为 `1-1a、1-1b`，拆三段为 `1-1a、1-1b、1-1c`；第一段也加a，不额外保留一行母场。下一原场仍为 `1-2`，未拆场号不变。原稿已有后缀时先保留，不能盲目再加后缀或覆盖其他场号。
- **A→B→A要保留三段**，例如客厅→院子→客厅；返回客厅是后一个剧情片段，不能倒回合并到第一行。分场表可以把两段客厅戏排在一起，但仍是不同场号、不同记录。
- 仅换机位、景别、特写、反打、同空间内走动，不拆；台词提到异地、口述回忆、VO/OS声源变化但没有实际切换画面，也不据此拆。转场界限不清时保留原文并备注“场内转场需确认”，不能凭“切”“转”“外”等单字机械拆分。
- 拆分覆盖母场全部正文，按源文顺序不遗漏、不重复，段内转场也须准确定位。每个子场分别复核具体地点、主场景、日夜、内外、拍摄内容、出镜角色、特约、群演、VO/OS、梳化服、道具、车辆和正文行数；不能整场复制到所有子场。
- 每个子场按自己的原排版正文计页数，不继承母场总数。切点落在原排版同一行时，应核对该行两侧片段归属，不能在两个子场重复累计整行；无法可靠计行则明确“页数需确认”。因逐子场向上取整，子场页数之和可能不同于母场直接计算结果。
- **拆场与连场框是两项独立判断。** 客厅→院子可拆为a/b，并在同主场景且剧情连续时画一组框；宅院→警局同样拆场，但必须断框。拆出的子场不自动视为连场。
- 生成程序按制表者读正文后给出的边界执行拆分，检查源文一致性、完整覆盖及场号冲突；它不能凭关键词替代语义判断。最终核验每个子场与原场的对应关系，备注和场号批注保留来源，报告列出原场数、拆分映射及最终行数。

### 顺场表与分场表

- `顺场表`按原剧本顺序排列；`分场表`将相同场景放在一起。每个已复核场次在两表各出现一次，不删场、不合并剧情、不新增汇总或空白分隔行。
- 先按主场景，再按具体子场景及内外分组；组的先后按其在原剧本首次出现的顺序，组内保留原顺场顺序。日夜、跨集、闪回均保留原值，不因归类自行按故事时间重排。同一宅院的客厅、走廊分别成组，同一具体场景的内/外也分组。
- 每场复核 `main_location`（主场景）。同一具体场景存在不同叫法时，可填 `location_group` 作为经正文确认的统一名称；未提供时按原场景名称分组。两个叫法相似不代表同景，不能靠模糊匹配归并，更不能把不同宅院的“客厅”归为一处。
- 两表“剧本中场景”保留原文名称，“拍摄顺序”保留原顺场序号，场号及a/b/c后缀原样保留；实际拍摄场地没有用户提供的确定信息时仍留空。分场表只做地点归类，不代表已安排拍摄日程。
- 未能确认主场景时填 `需确认`，备注说明原因及“主场景需确认”，相关场次单独保留、不互相归并。先完成其余内容，不因此询问整部剧的演员范围或停止制表。
- 两表共用已核定的连场组。归类后，同组相邻成员才共同闭合外框，分散成员分别成框；不能因位置相邻或同景而新增连场。所有组成员在两表相同的备注中保留组号和按原剧本顺序的接/续场号。
- 两表均保持 `J2` 冻结窗格、筛选、自动换行、横向打印及重复表头。分场表用场景单元格底色标出分组起点，不合并单元格或新增表头层级。

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
- “主场景”是剧本中的同一地点主体，如同一宅院、商号、警局或戏班；其内部房间、走廊及明确附属的门口可以是不同子场景。不能仅因同属“城西”、都是仓库，或共用一个实际拍摄场地就合为同一主场景。先核同名别称和空间归属，在每场的逐场复核项中填写 `main_location`，同一主场景使用统一名称，不为此增加Excel列。
- 逐场对照前场结尾、后场开头以及中间叙事，确认时间和行动衔接。场号相邻、同一地点、同写“日/夜”、同服装，均不能单独作为连场依据；一般的服装/伤妆接戏要求也不等于连场。明确“次日”“几天后”等时间跳跃应断开。
- 复核后在 `continuity_groups` 中明确列出组内场号，`reason` 写明同一主场景以及剧情衔接的正文依据。换到另一个主场景时拆组；拆后不足两场的片段不单独算连场。没有连场时填 `[]`；主场景归属或剧情衔接不清的，在备注写“连场关系需确认”，暂不画确认框并继续制作。
- 同组在表格中连续排列的场次，**只在“场次”列（D列）画一个加粗纯黑外框**：首格上边、末格下边和各格左右边加粗；组内横向分隔线保持普通细线。场号保持原样、每场独立一行，不合并单元格，不加粗场号字体，不把框扩展到其他列。
- 中间插入别条故事线时，只有原线仍在同一主场景、同一连续时段才能保持同组；人物实际转去别处再返回，不因回到原主场景而重新串为一组。不改剧本顺序，不把无关场次圈入。将同组拆成各自相邻的框选段；同组至少两场但某成员单独落在一段时可单格成框。所有连场成员在两表备注均追加同组标识和前后关联场号，如“连场1；接1-3；续1-6”，并保留原有备注。分场表复用已核定关系，不能因同景归类后排在一起就新增连场。
- 相邻的不同连场组分别闭合外框。一个场次不重复归入多个组；跨线交叉关系不能可靠归为一组时保留待确认说明。
- 自动解析只出初稿；连场分组必须由完整正文复核确定。严格校验验证分组结构和复核记录，不能代替剧情理解。

### 拍摄内容

- Write the main action or event as `who + does what / what happens`.
- Keep normal scenes at 12-18 Chinese characters.
- Allow 4-18 characters for empty shots, montage, inserts, phone screens, and transitions when a shorter production description is clearer.
- Do not use the first line of dialogue, an environment-only description, or a mechanical source-text truncation as the final summary.

### 独立角色列、特约演员与群演

这里的“独立角色列”是表格的分列方式，不表示演员级别或戏份轻重。按下面顺序判定；先依据全剧确定同一人物身份，再逐场判断其是否实际出镜。不询问演员范围，用户主动给出不同名单时才覆盖默认分列范围。

| 判定顺序 | 判定依据 | 填写位置 |
| --- | --- | --- |
| 1. 有姓名的剧情角色 | 正文有明确姓名，或明确指向一个人的固定昵称、外文名；即使只出现一场也单列 | 各自独立角色列 |
| 2. 其余独立小角色 | 未列入独立角色列，具有个人台词 **或** 与其他人物的单独对手戏；满足任一项即可 | 特约演员 |
| 3. 背景人群 | 只承担背景、围观、陪衬作用，没有个人台词，也没有单独对手戏 | 群演 |

- **特约不设最低场次、台词量或时长。** 只出现一场、只说一句话也可归特约；无台词但单独与主角交涉、扶助等对手戏，同样归特约。不能因为戏少而归群演，也不能因为出现多而自动升为独立角色列。
- 姓名及明确固定昵称决定默认独立分列，不能把所有“有台词的人”都各开一列。`保安甲`、`护士`、`班主`、`门房`、`某人之妻`等职业、编号或关系称谓不自动算姓名；有个人台词或独立对手戏时归特约。固定昵称须有正文身份依据，不能仅因“小/老”字开头就猜作姓名。
- 例如：`林晓`、固定昵称`小周`各自单列；`童年林晓`与成年`林晓`另列；`保安甲`只说一句“请登记”归特约；无台词但单独搀扶主角的`护士`归特约；其余无个人表演的围观群众归群演。例名仅说明规则，不是下一部剧的固定名单。
- 姓名、昵称、职务别称若经正文确认是同一人物，统一为一个身份，不能重复设列或跨类别重复填写。童年/幼年版本与成年版本分列；“童年某人”和“幼年某人”是否同一版本依正文统一。仅服装、伤妆或伪装变化不另开角色列。
- 同一角色在全剧保持一致分类：某个无姓名人物已有个人台词/对手戏，其后无台词的实际出镜场仍放特约；早先列在群演中的同一人物也应回溯拆出。不能把不同场次仅称谓相同的保安、护士擅自认作同一人；身份不明时保留原称谓，在备注标明“人物身份需确认”。
- 人群中某一人有个人台词或单独对手戏，要单独列入对应的独立角色列或特约；群演只记剩余背景人群。不能因看不清脸就将有个人台词的人归群演，也不能仅凭露脸、特写、职业或人数确定为特约。证据不足时保留原角色标签并在备注标“演员类别需确认”，不编造身份或人数。
- 默认 `roles` 覆盖全剧有姓名的剧情角色，按首次出现顺序排列；含剧情中的仅声音、被提及或照片人物，但不收录台词里随口引用的历史名人或比喻对象。同一角色的单字简称须唯一，角色姓名也不能重复。仅提及或静态照片没有现场出镜标记，注明声音/照片素材要求；不得凭姓名出现就勾选。
- `on_camera` 只填本场真实出镜且已列入 `roles` 的角色；童年出镜不能误勾成年。仅声音写VO/OS；本场既出镜又有VO/OS时可同时填写。需要新拍的闪回、监控或屏幕影像按画面人物记录，并在备注标明形式；单纯旧照片或口述不视为现场出镜。
- `special_actors` 在每个实际出镜场填写完整角色标签，以 `、` 分隔，要求/身份疑点写备注；无特约时填空字符串。三种出镜分类互斥，同一角色不得在独立列、特约和群演中重复统计。仅声音的无名小角色写VO/OS，不同时填本场特约。
- 自动关键词只提供草稿线索。制表者必须读完整正文复核 `roles`、`on_camera`、`special_actors`、`groups`；脚本检查重复和分类冲突，不能代替角色身份、个人对白及对手戏判断。

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

- `scene_splits` explicitly records complete source-body review, including `{}` when none are needed. For every split, child body slices concatenate to the entire unchanged parent body in order, with no gaps or duplicates; IDs are unique. Final rows = original scenes + sum(each split's child count minus one). Each child is independently reviewed, and both sheets contain the same unique child scene IDs; no redundant parent rows remain.
- Scene IDs are complete and unique. 顺场表 retains source order; 分场表 groups the same rows by reviewed main/specific location and inside/outside, with source order inside each group. Malformed headings are not silently dropped.
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
- Workbook contains exactly two visible sheets, `顺场表` then `分场表`, with the same locked headers. Compare all row values keyed by scene ID, including source indices, cast, page counts and continuity notes; totals and coverage must match. Every scene has reviewed `main_location`; unknowns stay separate with notes. Aliases require reviewed `location_group`, and the displayed original location is preserved.
- Header row and the first nine context columns through `页数` stay visible while navigating (`J2` freeze panes).
- Wrapped text is fully visible; row heights expand as needed.
- Filters, landscape printing, repeated print headers, and visual preview all work.
