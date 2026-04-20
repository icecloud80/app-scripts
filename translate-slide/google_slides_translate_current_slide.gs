const CONFIG = {
  SOURCE_ZH: 'zh',
  SOURCE_EN: 'en',
  TARGET_ZH: 'zh',
  TARGET_EN: 'en',
};

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Translate')
    .addItem('Translate to Chinese', 'translateCurrentSlideToChinese')
    .addItem('Translate to English', 'translateCurrentSlideToEnglish')
    .addToUi();
}

function translateCurrentSlideToChinese() {
  translateCurrentSlide_(CONFIG.SOURCE_EN, CONFIG.TARGET_ZH);
}

function translateCurrentSlideToEnglish() {
  translateCurrentSlide_(CONFIG.SOURCE_ZH, CONFIG.TARGET_EN);
}

function translateCurrentSlide_(sourceLang, targetLang) {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const selection = presentation.getSelection();
  const page = selection ? selection.getCurrentPage() : null;

  if (!page || page.getPageType() !== SlidesApp.PageType.SLIDE) {
    ui.alert('Please select a slide first.');
    return;
  }

  const slide = page.asSlide();
  let translatedCount = 0;

  slide.getPageElements().forEach(function(element) {
    translatedCount += translatePageElement_(element, sourceLang, targetLang);
  });

  if (!translatedCount) {
    ui.alert('No translatable text or table content was found on the current slide.');
  }
}

function translatePageElement_(element, sourceLang, targetLang) {
  const type = element.getPageElementType();
  let count = 0;

  if (type === SlidesApp.PageElementType.SHAPE) {
    count += translateShape_(element.asShape(), sourceLang, targetLang);
  } else if (type === SlidesApp.PageElementType.TABLE) {
    count += translateTable_(element.asTable(), sourceLang, targetLang);
  }

  return count;
}

function translateShape_(shape, sourceLang, targetLang) {
  const textRange = shape.getText();
  if (!textRange) return 0;
  return translateTextRange_(textRange, sourceLang, targetLang);
}

function translateTable_(table, sourceLang, targetLang) {
  let count = 0;
  const numRows = table.getNumRows();
  const numCols = table.getNumColumns();

  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      const cell = getSafeTableCell_(table, row, col);
      if (!cell) continue;
      count += translateTextRange_(cell.getText(), sourceLang, targetLang);
    }
  }

  return count;
}

function getSafeTableCell_(table, row, col) {
  try {
    return table.getCell(row, col);
  } catch (err) {
    const message = String(err && err.message ? err.message : err);
    if (message.indexOf('only allowed on the head') !== -1) {
      return null;
    }
    throw err;
  }
}

function containsChinese_(text) {
  return /[\u3400-\u9FFF]/.test(text);
}

function shouldTranslateText_(text) {
  if (!text) return false;

  const trimmed = text.trim();
  if (!trimmed) return false;
  if (/^(https?:\/\/|www\.)/i.test(trimmed)) return false;
  if (/^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(trimmed)) return false;

  return true;
}

function translateTextRange_(textRange, sourceLang, targetLang) {
  if (!textRange) return 0;

  const originalText = textRange.asString();
  if (!originalText) return 0;

  const normalizedText = originalText.replace(/\n+$/, '');
  if (!shouldTranslateText_(normalizedText)) return 0;

  const paragraphs = normalizedText.split('\n');
  let changed = false;

  const translatedParagraphs = paragraphs.map(function(paragraph) {
    const trimmed = paragraph.trim();

    if (!shouldTranslateText_(trimmed)) {
      return paragraph;
    }

    if (targetLang === 'en' && !containsChinese_(trimmed)) {
      return paragraph;
    }

    if (targetLang === 'zh' && containsChinese_(trimmed)) {
      return paragraph;
    }

    const translated = LanguageApp.translate(trimmed, sourceLang, targetLang);
    if (translated && translated !== trimmed) {
      changed = true;
      return translated;
    }

    return paragraph;
  });

  if (!changed) return 0;

  const finalText = translatedParagraphs
    .join('\n')
    .replace(/\n{2,}/g, '\n');

  textRange.setText(finalText);
  return 1;
}
