# 需求文档

## 背景

Google Drive 原生复制文件夹时，不会自动把 Google Docs 正文里指向原目录树内部文件的链接改成新副本链接。  
当目标目录权限与源目录不同，这些旧链接会失效或继续指向错误权限域。

## 目标

- 在绑定到 Google Docs 的 Apps Script 项目中提供独立的 `Copy Folder` 菜单
- 用户运行菜单后，可以递归复制 `folder A` 到 `folder B/Copy of + 原文件夹名称`
- 菜单需要同时支持“只复制目录树”和“复制后自动改写内部链接”两种模式
- 在“复制并改写链接”模式下，脚本自动把复制后 Google Docs 正文中的内部 Google 链接改写到新副本
- 项目需要提供一份面向最终使用者的使用手册，说明用途、限制、安装方式、运行步骤和结果解读

## 用户场景

- 用户有一个多层子文件夹组成的资料目录 `folder A`
- 目录内有多份 Google Docs，正文里链接到同目录树中的其它文档或文件夹
- 用户希望把整套资料复制到另一权限域下的 `folder B`
- 用户希望复制后的 Google Docs 继续正确互链，而不是继续指向旧目录

## 输入规则

- source folder A id 必填
- target folder B id 必填
- 目标根目录名称为 `Copy of + 源文件夹名称`
- 两个菜单动作共用相同的 source / target 输入规则

## 非目标

- 不修改源目录中的任何文件
- 不处理复制树之外的外部链接
- 不处理页眉、页脚、脚注中的链接
- 不处理 Google Docs 之外文件类型的正文改写
- `Copy Folder Only` 不负责扫描或改写任何复制后文档内容

## 验收标准

- 打开文档后能看到独立的 `Copy Folder` 菜单
- 菜单中存在 `Copy Folder and Replace Link` 和 `Copy Folder Only` 两个入口
- 两个入口都会先询问 source folder id，再询问 target folder id
- source 输入为空白时，会直接报错并中止执行
- 用户取消任一步输入时，流程安全终止，不执行复制
- 会在目标 `folder B` 下创建 `Copy of + 源文件夹名称`
- 会递归复制 `folder A` 下的所有子文件夹和文件
- 会建立旧资源 ID 到新资源 ID 的完整映射关系
- `Copy Folder and Replace Link` 会遍历所有复制后的 Google Docs 所有可访问 tab，并把正文中命中的旧内部链接替换为新链接
- `Copy Folder Only` 不会打开任何复制后的 Google Docs，也不会执行链接替换
- 对于富文本链接，只改写目标 URL，不改动显示文本
- 对于裸文本 URL，只替换命中的旧资源 ID
- 不在映射中的链接保持原样
- `Copy Folder and Replace Link` 的结果摘要至少包含：新根目录 ID、复制文件数、复制文件夹数、扫描文档数、总链接数、Google 链接数、实际改写文档数、替换链接数、未替换 Google 链接数
- `Copy Folder Only` 的结果摘要至少包含：新根目录 ID、复制文件数、复制文件夹数，以及“Internal link rewrite: skipped” 提示
