(function (globalScope) {
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

  function formatTimestamp_(timestamp) {
    if (!timestamp) {
      return 'Unknown';
    }

    const parsed = new Date(timestamp);
    if (Number.isNaN(parsed.getTime())) {
      return 'Unknown';
    }

    return parsed.toISOString().slice(0, 16).replace('T', ' ') + ' UTC';
  }

  function formatCommentLine_(entry) {
    const content = normalizeTextForMatch(entry && entry.content);
    if (!content) {
      return null;
    }

    const author =
      entry &&
      entry.author &&
      normalizeTextForMatch(entry.author.displayName);
    const formattedAuthor = author || 'Unknown';
    const formattedTime = formatTimestamp_(entry && entry.createdTime);

    return (
      'Person: ' +
      formattedAuthor +
      ' | Time: ' +
      formattedTime +
      ' | ' +
      content
    );
  }

  function buildThreadLines_(comment) {
    const mainLine = formatCommentLine_(comment);
    if (!mainLine) {
      return null;
    }

    const thread = {
      line: mainLine,
      replies: [],
    };

    const replies = Array.isArray(comment && comment.replies) ? comment.replies : [];

    for (let replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
      if (replies[replyIndex] && replies[replyIndex].deleted) {
        continue;
      }

      const replyLine = formatCommentLine_(replies[replyIndex]);
      if (replyLine) {
        thread.replies.push(replyLine);
      }
    }

    return thread;
  }

  function buildFeedbackSections(comments, activeTabText) {
    const groupedByHighlight = {};
    const sections = [];
    let matchedCommentCount = 0;
    let skippedCommentCount = 0;

    const sourceComments = Array.isArray(comments) ? comments : [];

    for (let commentIndex = 0; commentIndex < sourceComments.length; commentIndex += 1) {
      const comment = sourceComments[commentIndex];

      if (!comment || comment.deleted) {
        skippedCommentCount += 1;
        continue;
      }

      const highlight = normalizeTextForMatch(
        comment.quotedFileContent && comment.quotedFileContent.value
      );

      const threadLines = buildThreadLines_(comment);
      if (!threadLines) {
        skippedCommentCount += 1;
        continue;
      }

      const groupHighlight = highlight || '[No Highlight]';

      let section = groupedByHighlight[groupHighlight];
      if (!section) {
        section = {
          highlight: groupHighlight,
          threads: [],
        };
        groupedByHighlight[groupHighlight] = section;
        sections.push(section);
      }

      matchedCommentCount += 1;
      section.threads.push(threadLines);
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
    const lines = [
      'Feedback Export',
      'Source tab: ' + String(sourceTabTitle || ''),
    ];

    const sectionList = Array.isArray(sections) ? sections : [];

    if (!sectionList.length) {
      lines.push('', 'No matching comments found.');
      return lines.join('\n');
    }

    lines.push('');

    for (let sectionIndex = 0; sectionIndex < sectionList.length; sectionIndex += 1) {
      if (sectionIndex > 0) {
        lines.push('');
      }

      lines.push(sectionList[sectionIndex].highlight);

      const threads = Array.isArray(sectionList[sectionIndex].threads)
        ? sectionList[sectionIndex].threads
        : [];

      for (let threadIndex = 0; threadIndex < threads.length; threadIndex += 1) {
        lines.push('- ' + threads[threadIndex].line);

        const replies = Array.isArray(threads[threadIndex].replies)
          ? threads[threadIndex].replies
          : [];

        for (let replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
          lines.push('  - ' + replies[replyIndex]);
        }
      }
    }

    return lines.join('\n');
  }

  const api = {
    buildFeedbackSections: buildFeedbackSections,
    normalizeTextForMatch: normalizeTextForMatch,
    formatTimestamp: formatTimestamp_,
    renderFeedbackPlainText: renderFeedbackPlainText,
  };

  globalScope.ConvertCommentsHelpers = api;

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  }
})(typeof globalThis !== 'undefined' ? globalThis : this);
