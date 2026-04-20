# 本次改动说明

## 规则

- 项目现在只保留一个功能：`Copy Folder`
- 单一 Apps Script 部署入口为 `Code.gs`
- 单一源码入口为 `src/copy_folder.js`
- 单一测试入口为 `tests/copy_folder.test.js`

## 玩法逻辑

- 本项目不涉及游戏玩法，`N/A`

## AI 策略

- 本项目不涉及 AI 策略，`N/A`

## AI Heuristic

- 本项目不涉及 AI Heuristic，`N/A`

## 路线图

- 扩展更多 Google URL 变体识别
- 评估是否支持页眉、页脚和脚注中的链接改写
- 评估是否支持 Google Docs 之外更多文件类型的正文改写
- 视使用反馈决定是否给两种复制模式增加不同的结果日志或审计输出

## UI 设计

### Mobile

- 本项目不涉及独立 mobile UI，`N/A`

### PC

- 菜单入口为独立的 `Copy Folder`
- 菜单动作 1 为 `Copy Folder and Replace Link`
- 菜单动作 2 为 `Copy Folder Only`

## 清理结果

- 删除旧功能源码、测试和文档内容，只保留 copy-folder 所需文件
- 删除重复的 copy-folder 脚本副本，只保留单一部署入口和单一源码入口
- `Code.gs` 现在是唯一 Apps Script 入口
- `npm test` 现在只运行 copy-folder 回归测试
- source folder id 不再提供默认值，改为每次手动输入
- 复制出的根目录名称从固定文案改为 `Copy of + 源文件夹名称`
- 链接扫描从单一 body 扩展为所有可访问 document tabs，并在结果里输出总链接数与未替换原因计数
- 新增 `docs/user-manual.md`，用于说明脚本用途、安装方式、运行步骤和结果解读
- 新增 `Copy Folder Only` 菜单项，仅复制目录树，不扫描也不改写复制后的 Google Docs
- 原 `Run Copy Folder` 菜单项更名为 `Copy Folder and Replace Link`
