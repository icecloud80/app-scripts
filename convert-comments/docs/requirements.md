# Convert Comments 需求文档

## 背景

用户希望在 Google Docs 中执行一个菜单脚本，把当前 tab 上的评论内容整理成 plain text 反馈，方便在文档内继续编辑、汇总或转发。

## 目标

1. 脚本绑定在 Google Docs 上，通过 `Hydra Tools > Convert Comments To Feedback` 菜单触发。
2. 处理当前活动 `DocumentTab`。
3. 读取当前文档 comments，并提取对应的高亮文本。
4. 每条 comment 前增加 Person 和 Time。
5. 如果 comment thread 有 replies，replies 要写成 comment 下方的下层 bullet points。
6. 把结果写成“highlight + bullet comments”的结构。
7. 如果文档中已有名为 `Feedback` 的 tab，则写入该 tab。
8. 如果没有 `Feedback` tab，则回退为写入当前 tab 末尾的反馈导出区块。

## 非目标

1. 不自动创建新的 Google Docs tab。
2. 不保证 100% 精准恢复 Google Docs comment anchor 的真实位置。
3. 不处理 deleted comments。
4. 不提供复杂 HTML 对话框或用户输入表单。
