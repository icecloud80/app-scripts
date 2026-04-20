# 设计文档

## 总体方案

项目采用单一 Apps Script 功能模块：

- `onOpen()` 负责注册 `Copy Folder` 菜单
- `copyFolderWithPrompt()` 负责“复制并改写链接”入口
- `copyFolderOnlyWithPrompt()` 负责“只复制目录树”入口
- 两个入口共用同一套输入收集与目录复制逻辑，再按模式决定是否执行链接回写

## 文件结构

- `Code.gs`：Apps Script 单文件入口，直接部署用
- `src/copy_folder.js`：源码版本
- `tests/copy_folder.test.js`：本地 Node 回归测试
- `docs/user-manual.md`：面向使用者的操作手册

## 菜单设计

- 顶层菜单名：`Copy Folder`
- 动作 1：`Copy Folder and Replace Link`
- 动作 2：`Copy Folder Only`

## 复制算法

1. 弹出 prompt 收集 source folder id
2. source 为空时直接报错中止
3. 弹出 prompt 收集 target folder id
4. 在 target 下创建 `Copy of + 源文件夹名称`
5. 递归遍历源目录
6. 对每个子文件夹创建同名目标文件夹
7. 对每个文件执行 `makeCopy`
8. 记录旧文件夹 ID -> 新文件夹 ID
9. 记录旧文件 ID -> 新文件 ID
10. 对复制出的 Google Docs 记录待改写文档 ID
11. 若当前模式为 `Copy Folder Only`，直接返回复制摘要并标记 `Internal link rewrite: skipped`

## 链接改写算法

1. 仅在 `Copy Folder and Replace Link` 模式下执行
2. 合并文件夹和文件映射表
3. 逐个打开复制后的 Google Docs
4. 优先遍历所有 document tabs 的 body；若运行时不支持 tabs，则回退到 `doc.getBody()`
5. 对富文本 hyperlink，提取目标 URL 中的旧资源 ID 并替换
6. 对裸文本 URL，提取 URL 中的旧资源 ID 并替换
7. 对每条链接累计总链接数、Google 链接数、已替换数、未映射数、未支持格式数
8. 仅对命中复制映射的 Google URL 生效
9. 保存文档并返回替换统计

## URL 识别范围

当前支持以下常见格式：

- `https://docs.google.com/document/d/{id}/...`
- `https://docs.google.com/document/u/0/d/{id}/...`
- `https://docs.google.com/spreadsheets/d/{id}/...`
- `https://docs.google.com/presentation/d/{id}/...`
- `https://drive.google.com/file/d/{id}/...`
- `https://drive.google.com/file/u/0/d/{id}/...`
- `https://drive.google.com/drive/folders/{id}`
- `https://drive.google.com/drive/u/0/folders/{id}`
- `https://drive.google.com/open?id={id}`

## 风险与取舍

- 当前只处理 Google Docs 正文，不处理页眉、页脚、脚注
- 稀有 Google URL 变体如果未命中当前规则，会计入“unsupported format”，后续可以继续补充
- “只复制”和“复制并改写链接”共用同一套复制阶段，减少重复逻辑和行为漂移
- 复制和链接回写是同一次执行，任一步失败都会中断，避免得到半一致状态
