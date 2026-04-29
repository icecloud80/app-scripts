# Comvert Comments Slide 需求文档

## 背景

用户希望在 Google Slides 中执行一个菜单脚本，遍历 presentation comments，并把所有 comments 导出到最后新建的一页 slide，方便集中审阅和跳转回原页。

## 目标

1. 脚本绑定在 Google Slides 上，通过 `Hydra Tools > Export Comments To Slide` 菜单触发。
2. 读取当前 presentation 的所有 top-level comments 和 replies。
3. comments 按源 slide 分组。
4. 每个 slide 分组内按 top-level comment 时间排序。
5. replies 保留在其 parent comment 下方。
6. 每个 slide 分组标题包含到原 slide 的链接。
7. 每次执行都在 presentation 末尾新建一页新的导出 slide。
8. 导出内容包含评论人、comment 内容、时间戳。

## 非目标

1. 不覆盖或删除历史导出页。
2. 不保证对 duplicated quoted text 的 comments 做 100% 精确归页。
3. 不构建复杂 UI。
4. 不处理 deleted comments 的特殊恢复。
