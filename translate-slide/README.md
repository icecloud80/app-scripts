# Google Slides Translate Current Slide

这个项目是一个单文件 Google Apps Script，用来在 Google Slides 中把当前选中的单页幻灯片在中文和英文之间互译，同时尽量保持原有版式不变。

## 入口

- Apps Script 单文件入口：[google_slides_translate_current_slide.gs](/Users/mo.li/Workspace/appscript/translate-slide/google_slides_translate_current_slide.gs)
- 需求文档：[docs/requirements.md](/Users/mo.li/Workspace/appscript/translate-slide/docs/requirements.md)
- 设计文档：[docs/design.md](/Users/mo.li/Workspace/appscript/translate-slide/docs/design.md)

## 菜单

演示文稿打开后会出现独立的 `Translate` 菜单，里面包含：

- `Translate to Chinese`
- `Translate to English`

## 功能

脚本会：

1. 只处理当前选中的 slide。
2. 遍历 slide 上的文本形状和表格单元格。
3. 使用 `LanguageApp.translate` 在中英文之间互译。
4. 跳过空文本、URL、邮箱地址，以及已经是目标语言的文本。
5. 当没有选中 slide 或没有可翻译文本时弹出提示。

## 限制

- 当前只处理形状和表格中的文本。
- 当前不处理备注、图片、图表、页眉页脚或跨页批量翻译。
- 当前仓库内还没有为这个项目补齐自动化测试。
