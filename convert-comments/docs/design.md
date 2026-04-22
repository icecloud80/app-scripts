# Convert Comments 设计文档

## 设计目标

以最小实现成本提供一个可运行的 Google Docs Apps Script，把当前活动 tab 的 comments 转成文档内可编辑的 plain text 反馈列表。

## 方案

### 1. 用当前 active tab 作为输入边界

脚本通过 `DocumentApp.getActiveDocument().getActiveTab()` 读取当前用户所在的 `DocumentTab`，只把匹配到该 tab 正文文本的 comments 纳入导出。

这样做的原因：

- 用户操作模型最直接，不需要再额外选择源 tab。
- Apps Script 对 active tab 有明确支持。

### 2. 用 Drive comments API 读取 comments 与 quoted text

Apps Script 的 `DocumentApp` 不直接暴露 comments 读取能力，因此通过 Drive comments API 拉取：

- comment content
- replies
- `quotedFileContent.value`

其中 `quotedFileContent.value` 作为高亮文本近似值。

这样做的原因：

- 公开 API 能稳定返回 comments 与 quoted text。
- 这是当前最接近“读取 highlight + comments”的可用官方能力。

### 3. 用 quoted text 与 tab 正文做归属判断

脚本把当前 tab 正文标准化后，与每条 comment 的 quoted text 做包含匹配。只有 quoted text 能在当前 tab 正文中找到时，才认为该 comment 属于当前 tab。

这样做的原因：

- 当前公开 API 没有稳定的“comment -> tab”直接映射接口。
- 基于 quoted text 的匹配虽然不是完美锚点，但足以满足首版需求。

### 4. 输出格式保持简单可编辑

输出结构固定为：

- 一段 highlight 文本
- 紧跟其后的顶层 bullet comments
- 每条 comment 前包含作者和时间
- replies 作为该 comment 下的二级 bullet

每个 comment thread 会被整理成一个顶层 bullet，格式为 `Person + Time + comment text`。如果存在 replies，则 replies 会作为二级 bullet 输出，并保留各自的作者和时间。

这样做的原因：

- 用户要求的是 plain text 反馈，不需要保留 comment UI。
- 在 Google Docs 中继续编辑这个结构成本最低。

### 5. 输出位置采用“优先 Feedback tab，否则当前 tab 末尾”的回退策略

脚本先遍历文档 tabs，查找标题精确等于 `Feedback` 的 tab。找到后清空并重写该 tab；找不到时，在当前 tab 末尾追加一段带时间戳的反馈导出区块。

这样做的原因：

- 满足“优先单独反馈区域”的需求。
- 避免因为公开 API 缺少创建 tab 能力而让脚本不可用。
