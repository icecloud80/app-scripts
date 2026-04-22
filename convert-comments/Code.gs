var ConvertCommentsHelpers = (function () {
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

  function formatCommentLine_(entry) {
    var content = normalizeTextForMatch(entry && entry.content);
    if (!content) {
      return null;
    }

    var author =
      entry &&
      entry.author &&
      normalizeTextForMatch(entry.author.displayName);
    var formattedAuthor = author || 'Unknown';
    var formattedTime = formatTimestamp(entry && entry.createdTime);

    return (
      'Person: ' +
      formattedAuthor +
      ' | Time: ' +
      formattedTime +
      ' | ' +
      content
    );
  }

  function buildThread_(comment) {
    var mainLine = formatCommentLine_(comment);
    if (!mainLine) {
      return null;
    }

    var thread = {
      line: mainLine,
      replies: [],
    };

    var replies = Array.isArray(comment && comment.replies) ? comment.replies : [];

    for (var replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
      if (replies[replyIndex] && replies[replyIndex].deleted) {
        continue;
      }

      var replyLine = formatCommentLine_(replies[replyIndex]);
      if (replyLine) {
        thread.replies.push(replyLine);
      }
    }

    return thread;
  }

  function buildFeedbackSections(comments, activeTabText) {
    var groupedByHighlight = {};
    var sections = [];
    var matchedCommentCount = 0;
    var skippedCommentCount = 0;

    var sourceComments = Array.isArray(comments) ? comments : [];

    for (var commentIndex = 0; commentIndex < sourceComments.length; commentIndex += 1) {
      var comment = sourceComments[commentIndex];

      if (!comment || comment.deleted) {
        skippedCommentCount += 1;
        continue;
      }

      var highlight = normalizeTextForMatch(
        comment.quotedFileContent && comment.quotedFileContent.value
      );

      var thread = buildThread_(comment);
      if (!thread) {
        skippedCommentCount += 1;
        continue;
      }

      var groupHighlight = highlight || '[No Highlight]';

      var section = groupedByHighlight[groupHighlight];
      if (!section) {
        section = {
          highlight: groupHighlight,
          threads: [],
        };
        groupedByHighlight[groupHighlight] = section;
        sections.push(section);
      }

      matchedCommentCount += 1;
      section.threads.push(thread);
    }

    return {
      sections: sections,
      matchedCommentCount: matchedCommentCount,
      unmatchedThreadCount: 0,
      unmatchedThreads: [],
      skippedCommentCount: skippedCommentCount,
    };
  }

  function renderFeedbackPlainText(sections, sourceTabTitle) {
    var lines = [
      'Feedback Export',
      'Source tab: ' + String(sourceTabTitle || ''),
    ];

    var sectionList = Array.isArray(sections) ? sections : [];

    if (!sectionList.length) {
      lines.push('', 'No matching comments found.');
    } else {
      lines.push('');

      for (var sectionIndex = 0; sectionIndex < sectionList.length; sectionIndex += 1) {
        if (sectionIndex > 0) {
          lines.push('');
        }

        lines.push(sectionList[sectionIndex].highlight);

        var threads = Array.isArray(sectionList[sectionIndex].threads)
          ? sectionList[sectionIndex].threads
          : [];

        for (var threadIndex = 0; threadIndex < threads.length; threadIndex += 1) {
          lines.push('- ' + threads[threadIndex].line);

          var replies = Array.isArray(threads[threadIndex].replies)
            ? threads[threadIndex].replies
            : [];

          for (var nestedReplyIndex = 0; nestedReplyIndex < replies.length; nestedReplyIndex += 1) {
            lines.push('  - ' + replies[nestedReplyIndex]);
          }
        }
      }
    }

    return lines.join('\n');
  }

  return {
    buildFeedbackSections: buildFeedbackSections,
    formatTimestamp: formatTimestamp,
    normalizeTextForMatch: normalizeTextForMatch,
    renderFeedbackPlainText: renderFeedbackPlainText,
  };
})();

var ConvertCommentsApp = (function () {
  var CONFIG = {
    menuLabel: 'Hydra Tools',
    menuActionLabel: 'Convert Comments To Feedback',
    feedbackTabTitle: 'Feedback',
    fallbackHeading: 'Feedback Export',
  };

  function onOpen() {
    DocumentApp.getUi()
      .createMenu(CONFIG.menuLabel)
      .addItem(CONFIG.menuActionLabel, 'convertCurrentTabComments')
      .addToUi();
  }

  function convertCurrentTabComments() {
    var ui = DocumentApp.getUi();

    try {
      var doc = DocumentApp.getActiveDocument();
      var activeTab = doc.getActiveTab();

      if (!activeTab) {
        ui.alert('No active tab is available.');
        return null;
      }

      if (activeTab.getType() !== DocumentApp.TabType.DOCUMENT_TAB) {
        ui.alert('The active tab is not a document tab.');
        return null;
      }

      var sourceTab = activeTab.asDocumentTab();
      var sourceTabTitle = activeTab.getTitle() || 'Untitled Tab';
      var comments = fetchAllComments_(doc.getId());
      var report = ConvertCommentsHelpers.buildFeedbackSections(
        comments,
        sourceTab.getBody().getText()
      );
      var outputTarget = resolveOutputTarget_(doc, activeTab);

      writeFeedbackReport_(outputTarget, report.sections, {
        sourceTabTitle: sourceTabTitle,
        generatedAt: new Date(),
      });

      if (outputTarget.tab) {
        doc.setActiveTab(outputTarget.tab.getId());
      }

      ui.alert(
        [
          'Comments converted.',
          'Matched comment threads: ' + report.matchedCommentCount,
          'Skipped comments: ' + report.skippedCommentCount,
          'Output: ' + outputTarget.description,
        ].join('\n')
      );

      return report;
    } catch (error) {
      handleConvertCommentsError_(error, ui);
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
        var error = new Error('Failed to load comments: ' + statusCode + ' ' + body);
        error.httpStatusCode = statusCode;
        error.httpBody = body;
        throw error;
      }

      var payload = JSON.parse(body);
      var pageComments = payload.comments || [];

      for (var index = 0; index < pageComments.length; index += 1) {
        comments.push(pageComments[index]);
      }

      pageToken = payload.nextPageToken || '';
    } while (pageToken);

    return comments;
  }

  function buildCommentsListUrl_(fileId, pageToken) {
    var baseUrl =
      'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(fileId) +
      '/comments';
    var fields =
      'nextPageToken,comments(content,createdTime,deleted,quotedFileContent,author(displayName),replies(content,createdTime,deleted,author(displayName)))';
    var query = [
      'fields=' + encodeURIComponent(fields),
      'pageSize=100',
    ];

    if (pageToken) {
      query.push('pageToken=' + encodeURIComponent(pageToken));
    }

    return baseUrl + '?' + query.join('&');
  }

  function handleConvertCommentsError_(error, ui) {
    var body = error && error.httpBody ? String(error.httpBody) : '';

    if (
      error &&
      error.httpStatusCode === 403 &&
      body.indexOf('ACCESS_TOKEN_SCOPE_INSUFFICIENT') !== -1
    ) {
      ui.alert(
        [
          'This script needs fresh Drive authorization before it can read comments.',
          '',
          'What to do:',
          '1. Save the Apps Script project.',
          '2. Run "convertCurrentTabComments" again and accept the new permission prompt.',
          '3. If Google does not prompt again, revoke this script in Google Account permissions and rerun it.',
          '',
          'Required by Drive comments.list: drive / drive.file / drive.readonly.',
        ].join('\n')
      );
      return;
    }

    ui.alert(String(error && error.message ? error.message : error));
  }

  function resolveOutputTarget_(doc, activeTab) {
    var feedbackTab = findTabByTitle_(doc.getTabs(), CONFIG.feedbackTabTitle);

    if (feedbackTab) {
      return {
        body: feedbackTab.asDocumentTab().getBody(),
        clearFirst: false,
        description: 'Feedback tab (appended)',
        tab: feedbackTab,
      };
    }

    return {
      body: activeTab.asDocumentTab().getBody(),
      clearFirst: false,
      description: 'current tab fallback block',
      tab: activeTab,
    };
  }

  function findTabByTitle_(tabs, title) {
    for (var index = 0; index < tabs.length; index += 1) {
      var tab = tabs[index];
      if (tab.getTitle && tab.getTitle() === title) {
        return tab;
      }

      var childTabs = tab.getChildTabs ? tab.getChildTabs() : [];
      if (childTabs && childTabs.length) {
        var match = findTabByTitle_(childTabs, title);
        if (match) {
          return match;
        }
      }
    }

    return null;
  }

  function writeFeedbackReport_(outputTarget, sections, options) {
    var body = outputTarget.body;

    if (outputTarget.clearFirst) {
      body.clear();
    } else if (body.getNumChildren() > 0) {
      body.appendParagraph('');
      body.appendParagraph('');
    }

    body
      .appendParagraph(CONFIG.fallbackHeading)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Source tab: ' + options.sourceTabTitle);
    body.appendParagraph('Generated at: ' + options.generatedAt);

    if (!sections.length) {
      body.appendParagraph('No confidently matched comments found on the current tab.');
    }

    for (var sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
      body
        .appendParagraph(sections[sectionIndex].highlight)
        .setHeading(DocumentApp.ParagraphHeading.NORMAL)
        .editAsText()
        .setBold(true);

      var threads = sections[sectionIndex].threads || [];
      for (var threadIndex = 0; threadIndex < threads.length; threadIndex += 1) {
        var parentItem = body
          .appendListItem(threads[threadIndex].line)
          .setGlyphType(DocumentApp.GlyphType.BULLET);
        var replies = threads[threadIndex].replies || [];

        for (var replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
          body
            .appendListItem(replies[replyIndex])
            .setListId(parentItem)
            .setNestingLevel(1)
            .setGlyphType(DocumentApp.GlyphType.BULLET);
        }
      }
    }
  }

  return {
    convertCurrentTabComments: convertCurrentTabComments,
    onOpen: onOpen,
  };
})();

function onOpen() {
  ConvertCommentsApp.onOpen();
}

function convertCurrentTabComments() {
  return ConvertCommentsApp.convertCurrentTabComments();
}
