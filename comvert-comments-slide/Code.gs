var ConvertCommentsSlideHelpers = (function () {
  function normalizeTextForMatch(text) {
    return String(text || '')
      .replace(/[\u00A0]/g, ' ')
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201C\u201D]/g, '"')
      .replace(/[\u2013\u2014]/g, '-')
      .replace(/\u2026/g, '...')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function formatTimestamp(timestamp) {
    if (!timestamp) {
      return 'Unknown';
    }

    var parsed = new Date(timestamp);
    if (Number.isNaN(parsed.getTime())) {
      return 'Unknown';
    }

    return parsed.toISOString().slice(0, 16).replace('T', ' ') + ' UTC';
  }

  function buildSortValue_(timestamp) {
    var parsed = new Date(timestamp);
    return Number.isNaN(parsed.getTime()) ? Number.MAX_SAFE_INTEGER : parsed.getTime();
  }

  function compareThreads_(left, right) {
    var delta = left.sortValue - right.sortValue;
    if (delta !== 0) {
      return delta;
    }

    return String(left.id || '').localeCompare(String(right.id || ''));
  }

  function buildReply_(reply) {
    var content = normalizeTextForMatch(reply && reply.content);
    if (!content || (reply && reply.deleted)) {
      return null;
    }

    return {
      id: reply && reply.id ? String(reply.id) : '',
      authorName:
        normalizeTextForMatch(reply && reply.author && reply.author.displayName) ||
        'Unknown',
      createdTimeLabel: formatTimestamp(reply && reply.createdTime),
      content: content,
      sortValue: buildSortValue_(reply && reply.createdTime),
    };
  }

  function buildThread_(comment) {
    var content = normalizeTextForMatch(comment && comment.content);
    if (!content || (comment && comment.deleted)) {
      return null;
    }

    var replies = Array.isArray(comment && comment.replies) ? comment.replies : [];
    var builtReplies = [];

    for (var replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
      var builtReply = buildReply_(replies[replyIndex]);
      if (builtReply) {
        builtReplies.push(builtReply);
      }
    }

    builtReplies.sort(compareThreads_);

    return {
      id: comment && comment.id ? String(comment.id) : '',
      authorName:
        normalizeTextForMatch(comment && comment.author && comment.author.displayName) ||
        'Unknown',
      createdTimeLabel: formatTimestamp(comment && comment.createdTime),
      content: content,
      quotedText: normalizeTextForMatch(
        comment && comment.quotedFileContent && comment.quotedFileContent.value
      ),
      replies: builtReplies.map(function (reply) {
        return {
          id: reply.id,
          authorName: reply.authorName,
          createdTimeLabel: reply.createdTimeLabel,
          content: reply.content,
        };
      }),
      sortValue: buildSortValue_(comment && comment.createdTime),
    };
  }

  function normalizeSlideDescriptor_(slide) {
    var normalizedText = normalizeTextForMatch(slide && slide.text);
    return {
      slideIndex: slide && typeof slide.slideIndex === 'number' ? slide.slideIndex : 0,
      objectId: slide && slide.objectId ? String(slide.objectId) : '',
      title: normalizeTextForMatch(slide && slide.title) || 'Untitled Slide',
      text: normalizedText,
    };
  }

  function buildCommentExportModel(comments, slides, options) {
    var sourceComments = Array.isArray(comments) ? comments : [];
    var sourceSlides = Array.isArray(slides) ? slides : [];
    var normalizedSlides = sourceSlides.map(normalizeSlideDescriptor_);
    var slideGroupsById = {};
    var slideGroups = [];
    var ambiguousThreads = [];
    var unmatchedThreads = [];

    for (var commentIndex = 0; commentIndex < sourceComments.length; commentIndex += 1) {
      var thread = buildThread_(sourceComments[commentIndex]);
      if (!thread) {
        continue;
      }

      if (!thread.quotedText) {
        unmatchedThreads.push(thread);
        continue;
      }

      var matches = normalizedSlides.filter(function (slide) {
        return slide.text.indexOf(thread.quotedText) !== -1;
      });

      if (matches.length === 1) {
        var matchedSlide = matches[0];
        var group = slideGroupsById[matchedSlide.objectId];

        if (!group) {
          group = {
            type: 'slide',
            key: matchedSlide.objectId,
            heading:
              'Slide ' +
              String(matchedSlide.slideIndex + 1) +
              ' - ' +
              matchedSlide.title,
            slideIndex: matchedSlide.slideIndex,
            slideObjectId: matchedSlide.objectId,
            threadCount: 0,
            threads: [],
          };
          slideGroupsById[matchedSlide.objectId] = group;
          slideGroups.push(group);
        }

        group.threads.push(thread);
      } else if (matches.length > 1) {
        ambiguousThreads.push(thread);
      } else {
        unmatchedThreads.push(thread);
      }
    }

    slideGroups.sort(function (left, right) {
      return left.slideIndex - right.slideIndex;
    });

    slideGroups.forEach(function (group) {
      group.threads.sort(compareThreads_);
      group.threadCount = group.threads.length;
      group.threads = group.threads.map(function (thread) {
        return {
          id: thread.id,
          authorName: thread.authorName,
          createdTimeLabel: thread.createdTimeLabel,
          content: thread.content,
          replies: thread.replies,
        };
      });
    });

    ambiguousThreads.sort(compareThreads_);
    unmatchedThreads.sort(compareThreads_);

    var groups = slideGroups.slice();

    if (ambiguousThreads.length) {
      groups.push({
        type: 'ambiguous',
        key: 'ambiguous',
        heading: 'Ambiguous Comments',
        threadCount: ambiguousThreads.length,
        threads: ambiguousThreads.map(function (thread) {
          return {
            id: thread.id,
            authorName: thread.authorName,
            createdTimeLabel: thread.createdTimeLabel,
            content: thread.content,
            replies: thread.replies,
          };
        }),
      });
    }

    if (unmatchedThreads.length) {
      groups.push({
        type: 'unmatched',
        key: 'unmatched',
        heading: 'Unmatched Comments',
        threadCount: unmatchedThreads.length,
        threads: unmatchedThreads.map(function (thread) {
          return {
            id: thread.id,
            authorName: thread.authorName,
            createdTimeLabel: thread.createdTimeLabel,
            content: thread.content,
            replies: thread.replies,
          };
        }),
      });
    }

    return {
      presentationTitle:
        normalizeTextForMatch(options && options.presentationTitle) ||
        'Untitled Presentation',
      generatedAtLabel: formatTimestamp(options && options.generatedAt),
      summary: {
        totalThreads:
          slideGroups.reduce(function (sum, group) {
            return sum + group.threadCount;
          }, 0) +
          ambiguousThreads.length +
          unmatchedThreads.length,
        ambiguousThreads: ambiguousThreads.length,
        unmatchedThreads: unmatchedThreads.length,
      },
      groups: groups,
    };
  }

  function renderExportLines(model) {
    var summary = model && model.summary ? model.summary : {};
    var groups = Array.isArray(model && model.groups) ? model.groups : [];
    var lines = [
      'Comments Export',
      'Presentation: ' +
        String((model && model.presentationTitle) || '') +
        ' | Generated: ' +
        String((model && model.generatedAtLabel) || '') +
        ' | Threads: ' +
        String(summary.totalThreads || 0),
    ];

    if (!groups.length) {
      lines.push('', 'No comments found.');
      return lines;
    }

    for (var groupIndex = 0; groupIndex < groups.length; groupIndex += 1) {
      lines.push('');
      lines.push(groups[groupIndex].heading);

      var threads = Array.isArray(groups[groupIndex].threads)
        ? groups[groupIndex].threads
        : [];

      for (var threadIndex = 0; threadIndex < threads.length; threadIndex += 1) {
        lines.push(
          'Person: ' +
            String(threads[threadIndex].authorName || 'Unknown') +
            ' | Time: ' +
            String(threads[threadIndex].createdTimeLabel || 'Unknown')
        );
        lines.push(String(threads[threadIndex].content || ''));

        var replies = Array.isArray(threads[threadIndex].replies)
          ? threads[threadIndex].replies
          : [];

        for (var replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
          lines.push(
            '  Reply - Person: ' +
              String(replies[replyIndex].authorName || 'Unknown') +
              ' | Time: ' +
              String(replies[replyIndex].createdTimeLabel || 'Unknown')
          );
          lines.push('  ' + String(replies[replyIndex].content || ''));
        }
      }
    }

    return lines;
  }

  return {
    buildCommentExportModel: buildCommentExportModel,
    formatTimestamp: formatTimestamp,
    normalizeTextForMatch: normalizeTextForMatch,
    renderExportLines: renderExportLines,
  };
})();

var ComvertCommentsSlideApp = (function () {
  var CONFIG = {
    menuLabel: 'Hydra Tools',
    actionLabel: 'Export Comments To Slide',
    title: 'Comments Export',
  };

  function onOpen() {
    SlidesApp.getUi()
      .createMenu('Hydra Tools')
      .addItem('Export Comments To Slide', 'exportCommentsToSlide')
      .addToUi();
  }

  function exportCommentsToSlide() {
    var ui = SlidesApp.getUi();

    try {
      var presentation = SlidesApp.getActivePresentation();
      if (!presentation) {
        ui.alert('No active presentation is available.');
        return null;
      }

      var slides = presentation.getSlides();
      var slideDescriptors = extractSlideDescriptors_(slides);
      var comments = fetchAllComments_(presentation.getId());
      var model = ConvertCommentsSlideHelpers.buildCommentExportModel(
        comments,
        slideDescriptors,
        {
          presentationTitle: presentation.getName(),
          generatedAt: new Date(),
        }
      );

      appendExportSlide_(presentation, slides, model);

      ui.alert(
        [
          'Comments exported to a new slide.',
          'Threads: ' + model.summary.totalThreads,
          'Ambiguous: ' + model.summary.ambiguousThreads,
          'Unmatched: ' + model.summary.unmatchedThreads,
        ].join('\n')
      );

      return model;
    } catch (error) {
      ui.alert('Export failed: ' + (error && error.message ? error.message : error));
      throw error;
    }
  }

  function fetchAllComments_(fileId) {
    var comments = [];
    var pageToken = '';

    do {
      var response = UrlFetchApp.fetch(buildCommentsListUrl_(fileId, pageToken), {
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true,
      });
      var statusCode = response.getResponseCode();
      var body = response.getContentText();

      if (statusCode < 200 || statusCode >= 300) {
        throw new Error('Failed to load comments: ' + statusCode + ' ' + body);
      }

      var payload = JSON.parse(body);
      comments = comments.concat(payload.comments || []);
      pageToken = payload.nextPageToken || '';
    } while (pageToken);

    return comments;
  }

  function buildCommentsListUrl_(fileId, pageToken) {
    var url =
      'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(fileId) +
      '/comments?pageSize=100&fields=' +
      encodeURIComponent(
        'comments(id,content,createdTime,deleted,quotedFileContent/value,author/displayName,replies(id,content,createdTime,deleted,author/displayName)),nextPageToken'
      );

    if (pageToken) {
      url += '&pageToken=' + encodeURIComponent(pageToken);
    }

    return url;
  }

  function extractSlideDescriptors_(slides) {
    var descriptors = [];

    for (var slideIndex = 0; slideIndex < slides.length; slideIndex += 1) {
      descriptors.push({
        slideIndex: slideIndex,
        objectId: slides[slideIndex].getObjectId(),
        title: deriveSlideTitle_(slides[slideIndex]),
        text: extractSlideText_(slides[slideIndex]),
      });
    }

    return descriptors;
  }

  function deriveSlideTitle_(slide) {
    var titleShape = null;

    try {
      titleShape = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    } catch (error) {}

    if (!titleShape) {
      try {
        titleShape = slide.getPlaceholder(SlidesApp.PlaceholderType.CENTERED_TITLE);
      } catch (error2) {}
    }

    if (titleShape && titleShape.asShape) {
      var titleText = ConvertCommentsSlideHelpers.normalizeTextForMatch(
        titleShape.asShape().getText().asRenderedString()
      );
      if (titleText) {
        return titleText;
      }
    }

    var fallbackText = ConvertCommentsSlideHelpers.normalizeTextForMatch(
      extractSlideText_(slide)
    );
    if (!fallbackText) {
      return 'Untitled Slide';
    }

    return fallbackText.slice(0, 60);
  }

  function extractSlideText_(slide) {
    var fragments = [];
    var elements = slide.getPageElements();

    for (var elementIndex = 0; elementIndex < elements.length; elementIndex += 1) {
      var element = elements[elementIndex];
      var type = element.getPageElementType();

      if (type === SlidesApp.PageElementType.SHAPE) {
        var shapeText = element.asShape().getText().asRenderedString();
        if (shapeText) {
          fragments.push(shapeText);
        }
      } else if (type === SlidesApp.PageElementType.TABLE) {
        fragments.push(extractTableText_(element.asTable()));
      }
    }

    return fragments.join('\n');
  }

  function extractTableText_(table) {
    var fragments = [];

    for (var row = 0; row < table.getNumRows(); row += 1) {
      for (var col = 0; col < table.getNumColumns(); col += 1) {
        var cell = getSafeTableCell_(table, row, col);
        if (!cell) {
          continue;
        }

        var text = cell.getText().asRenderedString();
        if (text) {
          fragments.push(text);
        }
      }
    }

    return fragments.join('\n');
  }

  function getSafeTableCell_(table, row, col) {
    try {
      return table.getCell(row, col);
    } catch (error) {
      var message = String(error && error.message ? error.message : error);
      if (message.indexOf('only allowed on the head') !== -1) {
        return null;
      }
      throw error;
    }
  }

  function appendExportSlide_(presentation, sourceSlides, model) {
    var exportSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    var titleShape = exportSlide.insertTextBox(CONFIG.title, 24, 18, 672, 28);
    titleShape.getText().getTextStyle().setFontSize(20).setBold(true);

    var subtitleShape = exportSlide.insertTextBox(
      buildSubtitleLine_(model),
      24,
      48,
      672,
      24
    );
    subtitleShape.getText().getTextStyle().setFontSize(10);

    var bodyShape = exportSlide.insertTextBox('', 24, 82, 672, 438);
    var bodyText = bodyShape.getText();
    var exportLines = ConvertCommentsSlideHelpers.renderExportLines(model).slice(3);
    var rangeMap = buildHeadingRangeMap_(model, exportLines);
    var bodyValue = exportLines.join('\n');

    bodyText.setText(bodyValue);
    bodyText.getTextStyle().setFontSize(9);

    applyHeadingLinks_(bodyText, model, sourceSlides, rangeMap);
  }

  function buildSubtitleLine_(model) {
    return (
      'Presentation: ' +
      String(model.presentationTitle || '') +
      ' | Generated: ' +
      String(model.generatedAtLabel || '') +
      ' | Threads: ' +
      String(model.summary && model.summary.totalThreads ? model.summary.totalThreads : 0)
    );
  }

  function buildHeadingRangeMap_(model, exportLines) {
    var rangeMap = {};
    var offset = 0;

    for (var lineIndex = 0; lineIndex < exportLines.length; lineIndex += 1) {
      var line = exportLines[lineIndex];

      for (var groupIndex = 0; groupIndex < model.groups.length; groupIndex += 1) {
        var group = model.groups[groupIndex];
        if (group.type === 'slide' && group.heading === line) {
          rangeMap[group.key] = {
            start: offset,
            end: offset + line.length,
          };
          break;
        }
      }

      offset += line.length;
      if (lineIndex < exportLines.length - 1) {
        offset += 1;
      }
    }

    return rangeMap;
  }

  function applyHeadingLinks_(bodyText, model, sourceSlides, rangeMap) {
    for (var groupIndex = 0; groupIndex < model.groups.length; groupIndex += 1) {
      var group = model.groups[groupIndex];
      if (group.type !== 'slide') {
        continue;
      }

      var rangeInfo = rangeMap[group.key];
      if (!rangeInfo || rangeInfo.end <= rangeInfo.start) {
        continue;
      }

      var targetSlide = sourceSlides[group.slideIndex];
      if (!targetSlide) {
        continue;
      }

      bodyText
        .getRange(rangeInfo.start, rangeInfo.end)
        .getTextStyle()
        .setBold(true)
        .setLinkSlide(targetSlide);
    }
  }

  return {
    exportCommentsToSlide: exportCommentsToSlide,
    onOpen: onOpen,
  };
})();

function onOpen() {
  return ComvertCommentsSlideApp.onOpen();
}

function exportCommentsToSlide() {
  return ComvertCommentsSlideApp.exportCommentsToSlide();
}
