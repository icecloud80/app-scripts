# Comvert Comments Slide

这个项目是一个绑定在 Google Slides 上的 Apps Script，用来遍历演示文稿 comments，并在最后追加一页新的导出 slide。

## 入口

- Apps Script 入口：[Code.gs](/Users/mo.li/Workspace/appscript/comvert-comments-slide/Code.gs)
- 需求文档：[docs/requirements.md](/Users/mo.li/Workspace/appscript/comvert-comments-slide/docs/requirements.md)
- 设计文档：[docs/design.md](/Users/mo.li/Workspace/appscript/comvert-comments-slide/docs/design.md)

## 菜单

演示文稿打开后会出现 `Hydra Tools` 菜单，里面包含：

- `Export Comments To Slide`

## 功能

脚本会：

1. 读取当前 presentation 的 comments 和 replies。
2. 按源 slide 分组，并在每组内按 comment 时间排序。
3. 对无法稳定归页的 comments 单独放到 `Ambiguous Comments` 或 `Unmatched Comments`。
4. 每次运行都在最后追加一页新的导出 slide。
5. 在每个 slide 分组标题上添加到原 slide 的链接。

## 限制

- comment 到 slide 的映射是 best-effort，不保证 100% 精准。
- 当前只基于 slide 可见文本和 `quotedFileContent.value` 做匹配。
- 当前不删除历史导出页。
