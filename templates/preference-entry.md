# 偏好条目（一条一次交付）

> 复制本文件到临时位置填好，再 `python scripts/pref_memory.py --add <本文件>.json --consented` 写入。
> 规则见 [`../references/preference-memory.md`](../references/preference-memory.md)：只记偏好、不记事实、不记机密。
> 填写前先确认：用户已明确表示"这版满意"，且已同意记录。

```json
{
  "project": "",
  "kind": "deck",
  "satisfiedVersion": "",
  "preferences": [
    {"area": "文案语气", "statement": "", "source": "user-said", "confidence": "high"},
    {"area": "信息密度", "statement": "", "source": "observed", "confidence": "medium"}
  ],
  "avoid": [],
  "evidence": ""
}
```

## 字段怎么填

| 字段 | 取值 | 注意 |
|---|---|---|
| `project` | 项目目录名 | 与 `projects/<name>/` 一致，方便回查 |
| `kind` | `deck` / `doc` / `pdf` / `image` / `sheet` / `other` | 跨形态通用，不同 kind 默认不互用 |
| `satisfiedVersion` | 用户认可的那一版（如 `v3`、`2026-11-15-a`） | 不是最后一次机检，是最后一次**用户点头** |
| `preferences[].area` | 见下方七类 | 一类一条，别把三件事塞进一句 |
| `preferences[].statement` | 一句话，可判定 | 反例："喜欢简洁"（不可判定）→ 正例："每页 3–5 个文本块" |
| `preferences[].source` | `user-said` / `observed` | 用户原话或明确选择记 `user-said`；我从改动推的记 `observed` |
| `preferences[].confidence` | `high` / `medium` / `low` | `low` 默认不进建议；被用户否决两次就降到 `low` 或删掉 |
| `avoid` | 字符串数组 | 明确"不要"的东西：禁用词、不接受的颜色、不要的版式 |
| `evidence` | 一句话 | 用户原话或具体改动（"第 4 页要求删掉一条 bullet"） |

## 七类 area（固定用词，便于统计）

`文案语气` · `信息密度` · `配色倾向` · `字体偏好` · `版式节奏` · `表达禁忌` · `交付方式`

## 三条硬禁止

1. 不写项目事实与数据（数字、结论、公司名、客户名）
2. 不写凭据、内网地址、他人素材原文
3. 不写"他大概喜欢……"这类推断——推断不是记录

## 检查清单

- [ ] `project` 与 `kind` 填了，`satisfiedVersion` 是用户点头的那一版
- [ ] 每条 `statement` 都能判定真假（有数字、有对比、或明确动作）
- [ ] 每条标了 `source` 与 `confidence`
- [ ] `avoid` 写了用户明确说过"不要"的东西
- [ ] 全文没有事实数据、客户名、凭据、他人素材原文
- [ ] 写入前已把本条目原文给用户看过
