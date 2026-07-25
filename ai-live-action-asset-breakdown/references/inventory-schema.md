# Intermediate JSON schema

Build one UTF-8 JSON file with this structure:

```json
{
  "projectTitle": "《剧本名》",
  "episodes": [
    {
      "episode": "第1集",
      "assets": [
        {
          "scene": "1-1",
          "category": "人物",
          "name": "人物名称",
          "assetType": "基础",
          "parent": "",
          "description": "只写剧本明确的信息。"
        },
        {
          "scene": "1-2",
          "category": "人物",
          "name": "人物名称—服装变体",
          "assetType": "变体",
          "parent": "人物名称",
          "description": "说明相对母资产发生的视觉变化。"
        }
      ]
    }
  ]
}
```

## Required values

- `projectTitle`: Workbook title, normally wrapped in Chinese book-title marks.
- `episodes`: Complete episodes in script order.
- `episode`: Human-readable episode label such as `第1集`.
- `scene`: Earliest scene identifier where the asset is visible.
- `category`: Exactly one of `人物`, `场景`, or `道具`.
- `name`: Unique workbook-wide asset name. Use `大场景·具体空间` for scene hierarchy.
- `assetType`: Exactly `基础` or `变体`.
- `parent`: Empty for `基础`; exact existing asset name for `变体`.
- `description`: Concise, script-supported production description. Use `剧本未明确` instead of inventing details.

## Ordering and validation

- Put only first appearances in the JSON.
- A reused asset must not appear again in a later episode.
- Names must be unique across the workbook.
- A variant's parent must appear earlier in the JSON or earlier within the same episode.
- Keep episodes in script order. The builder automatically groups each episode by `人物`, `场景`, and `道具`, then sorts each category by scene.
- Do not add subtotal or total objects; the builder inserts formula-driven rows.
