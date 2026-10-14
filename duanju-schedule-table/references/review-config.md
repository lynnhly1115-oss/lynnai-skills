# Reviewed config

Use a JSON config to replace automatic draft fields after reading every complete scene. Keep empty reviewed text fields as empty strings rather than omitting them; lists and body line counts follow their types below. The counts in this example are illustrative and must be replaced with source-reviewed counts.

```json
{
  "roles": [
    {"name": "沈晚晴", "mark": "晴"},
    {"name": "陆锋", "mark": "锋"}
  ],
  "continuity_groups": [],
  "scenes": {
    "1-1": {
      "reviewed": true,
      "summary": "城市空镜交代下城区夜间环境",
      "body_line_count": 3,
      "hint": "",
      "on_camera": [],
      "special_actors": "",
      "groups": "富人、舞女、流民",
      "vo": "",
      "os": "",
      "costume": "舞女：旗袍、盘发；流民：破旧衣服、脏污妆",
      "props": "",
      "vehicles": "",
      "notes": "空镜；日夜、内外需确认"
    },
    "1-2": {
      "reviewed": true,
      "summary": "百乐门与下城区空镜建立环境",
      "body_line_count": 6,
      "hint": "",
      "on_camera": [],
      "special_actors": "",
      "groups": "",
      "vo": "",
      "os": "",
      "costume": "",
      "props": "",
      "vehicles": "",
      "notes": "空镜；日夜、内外需确认"
    },
    "1-3": {
      "reviewed": true,
      "summary": "沈晚晴被保安拦下登记后随护士入内",
      "body_line_count": 31,
      "hint": "",
      "on_camera": ["沈晚晴"],
      "special_actors": "保安甲、护士",
      "groups": "远处围观群众",
      "vo": "沈晚晴（内心独白）",
      "os": "陆锋（同一场景镜外喊话）",
      "costume": "",
      "props": "登记簿、笔",
      "vehicles": "轿车1辆（院外停放）",
      "notes": "保安甲有个人台词；护士与沈晚晴单独对手戏；群演无单独台词"
    }
  }
}
```

Rules:

- Include every scene under `scenes` for a `--strict` build.
- Set `reviewed` to `true` only after comparing the scene fields with the source body.
- Keep `summary` at 12-18 Chinese characters for normal scenes and 4-18 for empty shots, montage, inserts, phone screens, and transitions.
- Explicitly provide `roles` after reviewing the complete script: by default include all named plot characters and list childhood versions separately, without asking the user to choose or approve the roster. An explicitly user-supplied scope overrides this default. `[]` is valid when no individual columns apply. Other unnamed individual roles belong in `special_actors`; raw inferred labels still require source review to separate names, aliases, occupation labels and crowds.
- Every scene must contain all twelve reviewed fields: `summary`, `body_line_count`, `hint`, `on_camera`, `special_actors`, `groups`, `vo`, `os`, `costume`, `props`, `vehicles`, and `notes`. `on_camera` is a list; `body_line_count` is a positive integer or `null`; all other fields are strings. Empty results use `[]` for `on_camera` and `""` for string fields.
- `body_line_count` is the reviewed number of original-layout body text lines in this scene, including action and dialogue and continuing across pages. Exclude scene headings, character rosters, blank lines, headers/footers, page numbers and watermarks. Do not substitute Word paragraph counts or physical page counts. Each already-reviewed a/b/c split is counted independently; continuity groups do not combine counts.
- Pages are `ceil(body_line_count / 3) * 0.1`, displayed with one decimal: 1–3 lines → 0.1, 4–6 → 0.2, 28–30 → 1.0, 31–33 → 1.1. PDF draft counts require source review; DOCX draft counts remain unknown until checked against the original layout. If reliable counting is impossible or no body exists, use `null` and explain the reason with `页数需确认` in `notes`; the cell displays `需确认`. Zero, negative, fractional, string and Boolean counts are invalid.
- `hint` fills the `提示` column between `实际拍摄场地` and `场次` with short production tags, e.g. `雨／烟机／闪回／武戏／AI做`. Use only the tags that apply to that scene, separate multiple tags with `／`, and use `""` after confirming none. This is an open set of tags, not a required list to copy into every scene.
- Review the actual action and confirmed production method: rain must be an actual rain scene, smoke-machine use must have an effect/production basis, flashback must be shown rather than merely discussed, and action tags must reflect performed action. `AI做` requires an explicit script/user/confirmed-plan basis. Put detailed requirements, uncertain methods, and any partial-scene AI scope in `notes` instead of expanding the tag into a paragraph.
- `on_camera` lists only members of top-level `roles` who visibly appear in this scene. Named roles only mentioned, shown in a still photograph, or heard must not be marked physically present; note the image-material requirement or use VO/OS as appropriate. A visible role with narration may be in both `on_camera` and `vo`/`os`. Visible special actors are listed in `special_actors`.
- `os` means actual dialogue from within the current scene space but outside the frame; `vo` means narration/inner voice outside the current scene's spoken dialogue. Do not classify inner monologue as OS. Fill role names, optionally with the sound form in parentheses, and separate entries with `；`.
- Explicit source VO/OS labels are only draft candidates. A source-label/context conflict must be explained in `notes`, preserving the original label and marking confirmation needed. An unclear phone/broadcast voice remains traceable in notes until its spatial/temporal relationship is known.
- `vehicles` lists source-backed picture-vehicle types/identities, stated quantities, action/state and continuity. It is independent of `props`; keys or handheld vehicle documents still belong in props. No requirement means `""`, not an invented vehicle.
- `special_actors` is a string of exact role labels separated by `、`, for example `保安甲、护士`. Keep annotations in `notes`. Repeat the special actors in every scene where they appear; do not duplicate their identities under `roles` or `groups`.
- Individually speaking or interacting visible roles are special actors even if their labels contain `保安`, `村民`, or other group-like words. Only the remaining background crowd belongs in `groups`. Voice-only appearances belong in the appropriate voice column; unresolved voice types stay in notes.
- Existing configs must add all required fields and explicitly populate the source-reviewed `roles` before passing `--strict`; this is the table-maker's review, not a user-approval gate. Do not bulk-fill empty hint/cast/voice/vehicle fields without source review.
- Keep role marks unique.
- Explicitly include `continuity_groups` after reviewing both main-location identity and plot/time/action continuity. Both are required. `[]` means no confirmed groups; omitting it fails `--strict`. Do not bulk-fill `[]` without checking the scenes.
- Every scene used in a continuity group must supply a nonempty string `main_location` in its per-scene reviewed config, such as `沈宅`. This is source-reviewed metadata, not a new Excel column. All members must have the same value; differences in rooms or attached doorways are allowed only when the source establishes the same main location. A shared city district or shooting venue alone is insufficient. A location change breaks the group; a resulting singleton is not a new group.
- Each group must have `reviewed: true`, a source-backed `reason` covering shared main location and continuous plot, and at least two distinct existing scene IDs in script order. Groups must not overlap. The script checks member main-location equality and group structure; it cannot infer the plot or verify the accuracy of reviewed labels.
- Uncertain links belong in the relevant scene's `notes` as `连场关系需确认`, not in a confirmed group. Clothing continuity alone is not a group reason.
- For multiple input scripts, wrap each config under the exact source filename:

```json
{
  "script-a.docx": {"roles": [], "continuity_groups": [], "scenes": {}},
  "script-b.pdf": {"roles": [], "continuity_groups": [], "scenes": {}}
}
```

## Continuity group examples

Add this field to the reviewed config only when the listed scenes exist and the source supports the described relationship:

```json
{
  "continuity_groups": [
    {
      "scenes": ["1-3", "1-4", "1-5"],
      "reviewed": true,
      "reason": "同在沈宅，1-3屋内争吵后夺门而出，1-4立即追到院中继续争执，1-5在沈宅门口上车离开，时间和行动连续"
    },
    {
      "scenes": ["2-1", "2-3", "2-4"],
      "reviewed": true,
      "reason": "同在瑞丰行，2-1追逐未结束，2-2插入另一条故事线，2-3紧接瑞丰行内原追逐并延续至2-4"
    }
  ],
  "scenes": {
    "1-3": {"main_location": "沈宅"},
    "1-4": {"main_location": "沈宅"},
    "1-5": {"main_location": "沈宅"},
    "2-1": {"main_location": "瑞丰行"},
    "2-3": {"main_location": "瑞丰行"},
    "2-4": {"main_location": "瑞丰行"}
  }
}
```

The example shows only continuity fields; merge each `main_location` into the complete per-scene review. A shared main location alone does not establish plot continuity. Moving from 沈宅 to 警局 breaks the group even when the action continues.

The first group gets one outer box over its three adjacent scene-number cells. The second gets a one-cell box for `2-1` and one two-cell box for `2-3`/`2-4`; `2-2` is untouched. All members of the second group receive a common group label and preceding/following member references in their notes. No rows are reordered, no cells are merged, and no extra worksheet column is added.

Group membership is keyed by exact scene IDs, including already-reviewed suffixes such as `1-13a`; this marking feature does not itself split source scenes.
