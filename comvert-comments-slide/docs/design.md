# Comvert Comments Slide 设计文档

## 设计目标

以最小实现成本提供一个可运行的 Google Slides Apps Script，把当前 presentation 的 comments 导出到最后一页新建的 summary slide。

## 方案

### 1. 用 Drive comments API 读取 comments 与 replies

SlidesApp 不直接暴露 comments 读取能力，因此通过 Drive comments REST API 拉取：

- top-level comment content
- replies
- author
- createdTime
- `quotedFileContent.value`

这样做的原因：

- 这是当前公开 API 中最稳定的 comments 读取入口。
- replies、时间和作者信息都能一起取回。

### 2. 用 quoted text 与每页 slide 文本做 best-effort 归页

脚本遍历所有普通 slide，抽取形状和表格单元格里的文本，标准化后作为该页的文本语料。

每条 top-level comment：

- 如果 quoted text 只匹配一页，归到该页
- 如果匹配多页，归到 `Ambiguous Comments`
- 如果没有 quoted text 或匹配不到，归到 `Unmatched Comments`

这样做的原因：

- 公开文档没有稳定的 comment 到 slide 的直接映射。
- 文本匹配虽然不是完美锚点，但在首版里可测试、可解释。

### 3. 输出为一页新的 summary slide

脚本每次运行都在 presentation 末尾追加一页：

- 标题：`Comments Export`
- 副标题：presentation 标题、导出时间、thread 统计
- 正文：按 slide 分组的 comments

每个 slide 分组标题带一个到原 slide 的链接，正文包含 comment metadata、内容和 replies。

这样做的原因：

- 用户明确要求每次都新建一页，而不是覆盖旧结果。
- 单页 summary slide 最容易浏览和复制。
