# App Scripts

这个仓库的根目录现在只承担索引作用：每个一级子目录都是一个独立的 Google Apps Script 项目，拥有各自的源码、文档和 README。

## 子项目总览

| 子项目 | 作用 | README |
| --- | --- | --- |
| `convert-comments` | 在 Google Docs 中把当前 tab 的 comments 整理成 plain text feedback 导出区块。 | [convert-comments/README.md](./convert-comments/README.md) |
| `deepcopy-folder` | 递归复制 Google Drive 文件夹，并把复制后 Google Docs 正文里的内部 Google 链接改写到新副本。 | [deepcopy-folder/README.md](./deepcopy-folder/README.md) |
| `translate-slide` | 在 Google Slides 中把当前选中的单页幻灯片在中英文之间互译，保留页面布局，仅替换形状和表格中的文本。 | [translate-slide/README.md](./translate-slide/README.md) |

## 维护约定

- 新增、删除或重命名一级子项目时，更新本文件中的项目说明和 README 链接。
- 每个子项目都应在自己的目录内维护独立 README、需求文档和设计文档。
