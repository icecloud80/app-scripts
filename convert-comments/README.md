# Google Docs Convert Comments

这个子项目提供一个绑定在 Google Docs 上的 Apps Script，用来把当前活动 `DocumentTab` 上的 comments 导出成 plain text 反馈区块。

## 菜单

- 顶层菜单：`Hydra Tools`
- 子菜单项：`Convert Comments To Feedback`

## 目标行为

- 拉取当前文档所有 comments。
- 生成如下结构的反馈内容：

```text
Highlight text
- Person: Alice | Time: 2026-04-22 16:30 UTC | Comment 1
  - Person: Bob | Time: 2026-04-22 16:45 UTC | Reply 1
- Person: Carol | Time: 2026-04-22 17:00 UTC | Comment 2
```

- 优先写入名为 `Feedback` 的现有 tab。
- 如果已有 `Feedback` tab，则在现有内容后面继续追加。
- 如果没有 `Feedback` tab，则回退为在当前 tab 末尾追加一段生成结果。

## 入口

- Apps Script 入口：`Code.gs`
- 纯逻辑模块：`src/convert_comments.js`
- 本地测试：`tests/convert_comments.test.js`

## 限制

- 当前实现会把读到的 comments 全部整理进主列表，不再按当前 tab 做严格过滤。
- 公开 Apps Script 文档当前没有明确提供“创建新的 Google Docs tab”的能力，所以首版不自动新建 `DocumentTab`。

## 测试

```bash
npm test
```
