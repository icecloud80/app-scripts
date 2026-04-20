# Copy Folder Plan

## 目标

在保留单一 Copy Folder 项目边界的前提下，新增“只复制目录树”和“复制后改写链接”两种菜单动作。

## 实施项

### 任务 1：保留唯一功能入口

- 让 `Code.gs` 成为唯一 Apps Script 部署入口
- 保留 `onOpen()`、菜单注册逻辑
- 同时提供 `copyFolderWithPrompt()` 和 `copyFolderOnlyWithPrompt()` 两个公共入口

### 任务 2：保留唯一源码模块

- 保留 `src/copy_folder.js`
- 在同一模块内抽出共享复制流程，分别服务“只复制”和“复制并改写链接”两种模式

### 任务 3：保留唯一测试集

- 保留 `tests/copy_folder.test.js`
- 删除 review 相关测试
- 将 `npm test` 收敛到 copy-folder 测试

### 任务 4：清理文档

- README 只描述 Copy Folder
- requirements/design/change-log 只保留 Copy Folder 内容
- 文档需要同步说明两个菜单动作的差异、输入规则和结果摘要

### 任务 5：清理配置

- 调整 package 名称
- 删除不再需要的 Apps Script Advanced Service 配置
- 保留 Docs 和 Drive scope

### 任务 6：收敛交互与诊断输出

- source folder id 改为必填，不再提供默认值
- 目标根目录名改为 `Copy of + 源文件夹名称`
- 输出总链接数、Google 链接数、已替换数和未替换原因计数
- 扫描所有可访问 document tabs，减少首次运行中“部分链接未替换”的漏扫问题

### 任务 7：新增无链接改写模式

- 新增 `Copy Folder Only` 菜单项
- 该模式只复制目录树，不打开任何复制后的 Google Docs
- 摘要中明确标记 `Internal link rewrite: skipped`
