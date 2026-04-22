const assert = require('node:assert/strict');
const fs = require('node:fs');
const convertComments = require('../src/convert_comments.js');

function runTest(title, testFn) {
  try {
    testFn();
    console.log('PASS', title);
  } catch (error) {
    console.error('FAIL', title);
    throw error;
  }
}

runTest('groups comments by quoted highlight text that belongs to active tab', function () {
  const result = convertComments.buildFeedbackSections(
    [
      {
        content: 'Tighten this sentence.',
        createdTime: '2026-04-22T16:30:00Z',
        quotedFileContent: {
          value: 'Alpha beta gamma',
        },
        author: {
          displayName: 'Mo',
        },
        replies: [
          {
            content: 'Agreed.',
            createdTime: '2026-04-22T16:45:00Z',
            author: {
              displayName: 'Alex',
            },
          },
        ],
      },
      {
        content: 'This belongs to another tab.',
        quotedFileContent: {
          value: 'Outside tab text',
        },
      },
      {
        content: 'Clarify wording.',
        createdTime: '2026-04-22T17:00:00Z',
        quotedFileContent: {
          value: 'Alpha beta gamma',
        },
      },
    ],
    'Intro\nAlpha beta gamma\nOutro'
  );

  assert.deepEqual(result.sections, [
    {
      highlight: 'Alpha beta gamma',
      threads: [
        {
          line: 'Person: Mo | Time: 2026-04-22 16:30 UTC | Tighten this sentence.',
          replies: [
            'Person: Alex | Time: 2026-04-22 16:45 UTC | Agreed.',
          ],
        },
        {
          line: 'Person: Unknown | Time: 2026-04-22 17:00 UTC | Clarify wording.',
          replies: [],
        },
      ],
    },
    {
      highlight: 'Outside tab text',
      threads: [
        {
          line: 'Person: Unknown | Time: Unknown | This belongs to another tab.',
          replies: [],
        },
      ],
    },
  ]);
  assert.equal(result.matchedCommentCount, 3);
  assert.equal(result.unmatchedThreadCount, 0);
  assert.equal(result.skippedCommentCount, 0);
  assert.deepEqual(result.unmatchedThreads, []);
});

runTest('formats fallback export as plain text block', function () {
  const text = convertComments.renderFeedbackPlainText(
    [
      {
        highlight: 'Alpha beta gamma',
        threads: [
          {
            line: 'Person: Mo | Time: 2026-04-22 16:30 UTC | First note',
            replies: ['Person: Alex | Time: 2026-04-22 16:45 UTC | Follow-up'],
          },
          {
            line: 'Person: Kim | Time: 2026-04-22 17:00 UTC | Second note',
            replies: [],
          },
        ],
      },
      {
        highlight: 'Delta epsilon',
        threads: [
          {
            line: 'Person: Sam | Time: 2026-04-22 18:00 UTC | Another note',
            replies: [],
          },
        ],
      },
    ],
    'Source Tab'
  );

  assert.equal(
    text,
    [
      'Feedback Export',
      'Source tab: Source Tab',
      '',
      'Alpha beta gamma',
      '- Person: Mo | Time: 2026-04-22 16:30 UTC | First note',
      '  - Person: Alex | Time: 2026-04-22 16:45 UTC | Follow-up',
      '- Person: Kim | Time: 2026-04-22 17:00 UTC | Second note',
      '',
      'Delta epsilon',
      '- Person: Sam | Time: 2026-04-22 18:00 UTC | Another note',
    ].join('\n')
  );
});

runTest('skips empty comment threads instead of creating empty sections', function () {
  const result = convertComments.buildFeedbackSections(
    [
      {
        content: '   ',
        quotedFileContent: {
          value: 'Alpha beta gamma',
        },
      },
    ],
    'Alpha beta gamma'
  );

  assert.deepEqual(result.sections, []);
  assert.equal(result.matchedCommentCount, 0);
  assert.equal(result.skippedCommentCount, 1);
});

runTest('matches highlight text despite smart quote differences', function () {
  const result = convertComments.buildFeedbackSections(
    [
      {
        content: 'Fix wording.',
        createdTime: '2026-04-22T19:00:00Z',
        quotedFileContent: {
          value: "Hydra's launch plan",
        },
      },
    ],
    'Hydra’s launch plan'
  );

  assert.equal(result.matchedCommentCount, 1);
  assert.equal(result.unmatchedThreadCount, 0);
  assert.deepEqual(result.sections, [
    {
      highlight: "Hydra's launch plan",
      threads: [
        {
          line: 'Person: Unknown | Time: 2026-04-22 19:00 UTC | Fix wording.',
          replies: [],
        },
      ],
    },
  ]);
});

runTest('ships a single Apps Script file that defines the global helper object', function () {
  const codeSource = fs.readFileSync(
    '/Users/mo.li/Workspace/appscript/convert-comments/Code.gs',
    'utf8'
  );

  assert.match(codeSource, /var ConvertCommentsHelpers =/);
  assert.match(codeSource, /buildFeedbackSections/);
  assert.match(codeSource, /setGlyphType\(DocumentApp\.GlyphType\.BULLET\)/);
});
