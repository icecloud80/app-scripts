# Google Drive Copy Folder

这个项目现在只保留一个功能域：递归复制 Google Drive 中的文件夹，并按菜单模式决定是否把复制后 Google Docs 正文里的内部 Google 链接改写到新副本。

## 入口

- Apps Script 单文件入口：[Code.gs](/Users/mo.li/Workspace/appscript/deepcopy-folder/Code.gs)
- 源码文件：[src/copy_folder.js](/Users/mo.li/Workspace/appscript/deepcopy-folder/src/copy_folder.js)
- 自动测试：[tests/copy_folder.test.js](/Users/mo.li/Workspace/appscript/deepcopy-folder/tests/copy_folder.test.js)
- 用户手册：[docs/user-manual.md](/Users/mo.li/Workspace/appscript/deepcopy-folder/docs/user-manual.md)

## 菜单

文档打开后会出现独立的 `Copy Folder` 菜单，里面包含：

- `Copy Folder and Replace Link`
- `Copy Folder Only`

## 功能

脚本会：

1. 询问 source folder A id
2. 询问目标 folder B id
3. 在目标目录下创建 `Copy of + 源文件夹名称`
4. source 和 target 都要求每次手动输入，不提供默认值
5. 递归复制整个目录树
6. 如果选择 `Copy Folder and Replace Link`，扫描复制后的 Google Docs 所有可访问 tab 正文，把命中的旧内部 Google Drive / Docs 链接改成新副本链接
7. 如果选择 `Copy Folder Only`，只复制目录树，不打开也不修改任何复制后的 Google Docs
8. 执行结束后输出复制结果；在改写模式下额外输出总链接数、Google 链接数、已替换数以及未替换原因计数

## 限制

- 当前只改写复制后 Google Docs 正文中的链接
- 当前不处理页眉、页脚、脚注和 Google Docs 之外文件类型的正文内容
- 不在复制映射中的外部链接保持不变

## 权限

项目需要以下授权范围：

- `https://www.googleapis.com/auth/documents`
- `https://www.googleapis.com/auth/drive`

## 测试

```bash
npm test
```
