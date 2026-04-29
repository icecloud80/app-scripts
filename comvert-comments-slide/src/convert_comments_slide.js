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

  function formatTimestamp(timestamp) {
    if (!timestamp) {
      return 'Unknown';
    }

    const parsed = new Date(timestamp);
    if (Number.isNaN(parsed.getTime())) {
      return 'Unknown';
    }

    return parsed.toISOString().slice(0, 16).replace('T', ' ') + ' UTC';
  }

  function buildSortValue_(timestamp) {
    const parsed = new Date(timestamp);
    return Number.isNaN(parsed.getTime()) ? Number.MAX_SAFE_INTEGER : parsed.getTime();
  }

  function compareThreads_(left, right) {
    const delta = left.sortValue - right.sortValue;
    if (delta !== 0) {
      return delta;
    }

    return String(left.id || '').localeCompare(String(right.id || ''));
  }

  function buildReply_(reply) {
    const content = normalizeTextForMatch(reply && reply.content);
    if (!content || (reply && reply.deleted)) {
      return null;
    }

    return {
      id: reply && reply.id ? String(reply.id) : '',
      authorName: normalizeTextForMatch(
        reply && reply.author && reply.author.displayName
      ) || 'Unknown',
      createdTime: reply && reply.createdTime ? String(reply.createdTime) : '',
      createdTimeLabel: formatTimestamp(reply && reply.createdTime),
      content: content,
      sortValue: buildSortValue_(reply && reply.createdTime),
    };
  }

  function buildThread_(comment) {
    const content = normalizeTextForMatch(comment && comment.content);
    if (!content || (comment && comment.deleted)) {
      return null;
    }

    const replies = Array.isArray(comment && comment.replies) ? comment.replies : [];
    const builtReplies = [];

    for (let replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
      const builtReply = buildReply_(replies[replyIndex]);
      if (builtReply) {
        builtReplies.push(builtReply);
      }
    }

    builtReplies.sort(compareThreads_);

    return {
      id: comment && comment.id ? String(comment.id) : '',
      authorName:
        normalizeTextForMatch(
          comment && comment.author && comment.author.displayName
        ) || 'Unknown',
      createdTime: comment && comment.createdTime ? String(comment.createdTime) : '',
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
    const normalizedText = normalizeTextForMatch(slide && slide.text);
    return {
      slideIndex: slide && typeof slide.slideIndex === 'number' ? slide.slideIndex : 0,
      objectId: slide && slide.objectId ? String(slide.objectId) : '',
      title: normalizeTextForMatch(slide && slide.title) || 'Untitled Slide',
      text: normalizedText,
    };
  }

  function buildCommentExportModel(comments, slides, options) {
    const sourceComments = Array.isArray(comments) ? comments : [];
    const sourceSlides = Array.isArray(slides) ? slides : [];
    const normalizedSlides = sourceSlides.map(normalizeSlideDescriptor_);
    const slideGroupsById = {};
    const slideGroups = [];
    const ambiguousThreads = [];
    const unmatchedThreads = [];

    for (let commentIndex = 0; commentIndex < sourceComments.length; commentIndex += 1) {
      const thread = buildThread_(sourceComments[commentIndex]);
      if (!thread) {
        continue;
      }

      if (!thread.quotedText) {
        unmatchedThreads.push(thread);
        continue;
      }

      const matches = normalizedSlides.filter(function (slide) {
        return slide.text.indexOf(thread.quotedText) !== -1;
      });

      if (matches.length === 1) {
        const matchedSlide = matches[0];
        let group = slideGroupsById[matchedSlide.objectId];

        if (!group) {
          group = {
            type: 'slide',
            key: matchedSlide.objectId,
            heading: 'Slide ' + String(matchedSlide.slideIndex + 1) + ' - ' + matchedSlide.title,
            slideIndex: matchedSlide.slideIndex,
            slideObjectId: matchedSlide.objectId,
            threadCount: 0,
            threads: [],
          };
          slideGroupsById[matchedSlide.objectId] = group;
          slideGroups.push(group);
        }

        group.threads.push(thread);
        continue;
      }

      if (matches.length > 1) {
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

    const groups = slideGroups.slice();

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
        normalizeTextForMatch(options && options.presentationTitle) || 'Untitled Presentation',
      generatedAtLabel: formatTimestamp(options && options.generatedAt),
      summary: {
        totalThreads: slideGroups.reduce(function (sum, group) {
          return sum + group.threadCount;
        }, 0) + ambiguousThreads.length + unmatchedThreads.length,
        ambiguousThreads: ambiguousThreads.length,
        unmatchedThreads: unmatchedThreads.length,
      },
      groups: groups,
    };
  }

  function renderExportLines(model) {
    const summary = model && model.summary ? model.summary : {};
    const groups = Array.isArray(model && model.groups) ? model.groups : [];
    const lines = [
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

    for (let groupIndex = 0; groupIndex < groups.length; groupIndex += 1) {
      lines.push('');
      lines.push(groups[groupIndex].heading);

      const threads = Array.isArray(groups[groupIndex].threads)
        ? groups[groupIndex].threads
        : [];

      for (let threadIndex = 0; threadIndex < threads.length; threadIndex += 1) {
        lines.push(
          'Person: ' +
            String(threads[threadIndex].authorName || 'Unknown') +
            ' | Time: ' +
            String(threads[threadIndex].createdTimeLabel || 'Unknown')
        );
        lines.push(String(threads[threadIndex].content || ''));

        const replies = Array.isArray(threads[threadIndex].replies)
          ? threads[threadIndex].replies
          : [];

        for (let replyIndex = 0; replyIndex < replies.length; replyIndex += 1) {
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

  const api = {
    buildCommentExportModel: buildCommentExportModel,
    formatTimestamp: formatTimestamp,
    normalizeTextForMatch: normalizeTextForMatch,
    renderExportLines: renderExportLines,
  };

  globalScope.ConvertCommentsSlideHelpers = api;

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  }
})(typeof globalThis !== 'undefined' ? globalThis : this);
