# User Manual

## 1. 这个脚本是做什么的

这个脚本用于：

- 递归复制一个 Google Drive 文件夹
- 把复制后的 Google Docs 中，原本指向旧目录树内部文件的 Google 链接，自动改成新副本链接

它适合这样的场景：

- 你有一个资料包文件夹，里面有很多子文件夹和 Google Docs
- 这些 Docs 之间互相有链接
- 你想把整套资料复制到另一个目录，通常是另一套权限空间
- 你希望复制后的文档之间还能继续正确互相跳转
- 或者你只想要一份纯复制结果，不希望脚本额外扫描和改写文档链接

## 2. 这个脚本会做什么

运行后，脚本会按下面顺序执行：

1. 让你输入 source folder id
2. 让你输入 target folder id
3. 在 target folder 下创建一个新目录
4. 新目录名称为 `Copy of + 源文件夹名称`
5. 递归复制 source folder 下的所有子文件夹和文件
6. 如果你选择 `Copy Folder and Replace Link`，脚本会继续找出复制后的 Google Docs
7. 在改写模式下，脚本会扫描这些 Docs 的所有可访问 tabs 正文
8. 在改写模式下，将命中的旧内部 Google Docs / Drive 链接替换成新副本链接
9. 输出执行结果统计

## 3. 不会做什么

当前版本不会处理这些内容：

- 源目录之外的外部链接
- Google Docs 页眉、页脚、脚注中的链接
- Google Docs 之外文件类型的正文内容改写
- Drive 权限设置本身

也就是说，这个脚本负责“复制目录树 + 修正文档内部互链”，不负责“重新配置分享权限”。

## 4. 使用前准备

你需要准备：

- 一个绑定到 Google Docs 的 Apps Script 项目
- 把 [Code.gs](/Users/mo.li/Workspace/appscript/Code.gs) 放进项目
- 把 [appsscript.json](/Users/mo.li/Workspace/appscript/appsscript.json) 的权限配置同步到项目

脚本需要这些权限：

- `https://www.googleapis.com/auth/documents`
- `https://www.googleapis.com/auth/drive`

第一次运行时，Google 会要求你授权。

## 5. 如何找到 folder id

打开 Google Drive 文件夹时，URL 通常像这样：

```text
https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOp
```

其中最后这一段：

```text
1AbCdEfGhIjKlMnOp
```

就是 folder id。

## 6. 如何安装

### 方法 A：直接在 Apps Script 中使用

1. 打开一个 Google Docs 文档
2. 进入 `Extensions > Apps Script`
3. 删除默认示例代码
4. 粘贴 [Code.gs](/Users/mo.li/Workspace/appscript/Code.gs) 的内容
5. 确认 manifest 中包含 [appsscript.json](/Users/mo.li/Workspace/appscript/appsscript.json) 里的权限
6. 保存项目
7. 刷新原来的 Google Docs 页面

刷新后，菜单栏里应该出现：

- `Copy Folder`

菜单下应该有：

- `Copy Folder and Replace Link`
- `Copy Folder Only`

### 方法 B：在本地仓库维护源码

如果你在本地维护这个项目：

- 源码主文件是 [src/copy_folder.js](/Users/mo.li/Workspace/appscript/src/copy_folder.js)
- 发布到 Apps Script 时使用 [Code.gs](/Users/mo.li/Workspace/appscript/Code.gs)

## 7. 如何运行

1. 在绑定文档里点击 `Copy Folder`
2. 选择 `Copy Folder and Replace Link` 或 `Copy Folder Only`
3. 输入 source folder id
4. 输入 target folder id
5. 等待脚本执行完成
6. 查看弹窗结果

## 8. 执行结果怎么看

脚本结束后会看到一段摘要。

重点字段说明：

- `Root folder`
  复制出来的新根目录名称

- `Root folder id`
  新根目录的 Drive id

- `Folders copied`
  一共复制了多少个文件夹

- `Files copied`
  一共复制了多少个文件

- `Internal link rewrite: skipped`
  只会在 `Copy Folder Only` 模式出现，表示这次运行没有扫描也没有改写任何复制后的文档链接

- `Docs scanned`
  一共扫描了多少个复制后的 Google Docs

- `Total links found`
  一共遇到了多少条链接

- `Google links found`
  其中有多少条是 Google Docs / Drive 链接

- `Docs rewritten`
  有多少个文档发生了实际替换

- `Links replaced`
  一共有多少条链接被成功改成新副本链接

- `Google links not replaced (outside copied tree)`
  这类链接是 Google 链接，但它们指向的文件不在本次复制映射里，所以不会改

- `Google links not replaced (unsupported format)`
  这类链接是 Google 链接，但 URL 结构当前版本还没识别出来，所以没法改

## 9. 哪些链接可以被识别

当前版本支持常见的 Google URL，例如：

- `https://docs.google.com/document/d/{id}/...`
- `https://docs.google.com/document/u/0/d/{id}/...`
- `https://docs.google.com/spreadsheets/d/{id}/...`
- `https://docs.google.com/presentation/d/{id}/...`
- `https://drive.google.com/file/d/{id}/...`
- `https://drive.google.com/file/u/0/d/{id}/...`
- `https://drive.google.com/drive/folders/{id}`
- `https://drive.google.com/drive/u/0/folders/{id}`
- `https://drive.google.com/open?id={id}`

同时支持两种正文形式：

- 富文本超链接
- 裸文本 URL

## 10. 为什么有些链接没有被替换

常见原因有三类：

### 1. 链接指向复制树外部

例如文档里链到另一个共享资料库，那个文件不在本次 source folder 目录树中。  
这种链接会计入：

- `Google links not replaced (outside copied tree)`

### 2. 链接格式当前不支持

例如某些更少见的 Google URL 结构还没被解析规则覆盖。  
这种链接会计入：

- `Google links not replaced (unsupported format)`

### 3. 链接不在正文里

当前只处理正文，不处理：

- 页眉
- 页脚
- 脚注

## 11. 常见问题

### 菜单没有出现

检查：

1. 你是否已经保存脚本
2. 你是否刷新了 Google Docs 页面
3. 项目里是否包含 `onOpen()`
4. 是否真的部署的是 [Code.gs](/Users/mo.li/Workspace/appscript/Code.gs)

### 运行时报权限错误

通常是因为还没完成 Google 授权。  
重新运行一次，按提示授权。

### 只替换了一部分链接

如果你运行的是 `Copy Folder Only`，看到 `Internal link rewrite: skipped` 就是预期行为。

如果你运行的是 `Copy Folder and Replace Link`，先看结果弹窗里的：

- `Total links found`
- `Google links found`
- `Links replaced`
- `Google links not replaced (outside copied tree)`
- `Google links not replaced (unsupported format)`

这几个数字能先帮助判断是：

- 有些链接本来就不在复制树里
- 还是有些 Google URL 格式还没覆盖

## 12. 本地验证

如果你在本地维护这个项目，可以运行：

```bash
npm test
```

测试文件在：

- [tests/copy_folder.test.js](/Users/mo.li/Workspace/appscript/tests/copy_folder.test.js)

## 13. 相关文件

- 使用入口：[Code.gs](/Users/mo.li/Workspace/appscript/Code.gs)
- 源码实现：[src/copy_folder.js](/Users/mo.li/Workspace/appscript/src/copy_folder.js)
- 需求文档：[docs/requirements.md](/Users/mo.li/Workspace/appscript/docs/requirements.md)
- 设计文档：[docs/design.md](/Users/mo.li/Workspace/appscript/docs/design.md)
