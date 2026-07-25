# Reviewed config

Use a JSON config to replace automatic draft fields after reading every complete scene. Keep empty reviewed fields as empty strings rather than omitting them.

```json
{
  "roles": [
    {"name": "沈晚晴", "mark": "晴"},
    {"name": "陆锋", "mark": "锋"}
  ],
  "scenes": {
    "1-1": {
      "reviewed": true,
      "summary": "城市空镜交代下城区夜间环境",
      "groups": "富人、舞女、流民",
      "costume": "舞女：旗袍、盘发；流民：破旧衣服、脏污妆",
      "props": "",
      "notes": "空镜；日夜、内外需确认"
    },
    "1-2": {
      "reviewed": true,
      "summary": "百乐门与下城区空镜建立环境",
      "groups": "",
      "costume": "",
      "props": "",
      "notes": "空镜；日夜、内外需确认"
    }
  }
}
```

Rules:

- Include every scene under `scenes` for a `--strict` build.
- Set `reviewed` to `true` only after comparing the scene fields with the source body.
- Keep `summary` at 12-18 Chinese characters for normal scenes and 4-18 for empty shots, montage, inserts, phone screens, and transitions.
- Use `groups`, `costume`, `props`, and `notes` to override the draft. Use an empty string when the reviewed result is intentionally blank.
- Keep role marks unique.
- For multiple input scripts, wrap each config under the exact source filename:

```json
{
  "script-a.docx": {"roles": [], "scenes": {}},
  "script-b.pdf": {"roles": [], "scenes": {}}
}
```
