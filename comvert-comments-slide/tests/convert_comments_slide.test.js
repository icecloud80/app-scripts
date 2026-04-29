const assert = require('node:assert/strict');
const fs = require('node:fs');
const convertCommentsSlide = require('../src/convert_comments_slide.js');

function runTest(title, testFn) {
  try {
    testFn();
    console.log('PASS', title);
  } catch (error) {
    console.error('FAIL', title);
    throw error;
  }
}

runTest('groups threads by source slide and sorts them by created time', function () {
  const result = convertCommentsSlide.buildCommentExportModel(
    [
      {
        id: 'c2',
        content: 'Second note',
        createdTime: '2026-04-28T19:00:00Z',
        quotedFileContent: { value: 'Alpha roadmap' },
        author: { displayName: 'Mo' },
        replies: [
          {
            id: 'r2',
            content: 'Later reply',
            createdTime: '2026-04-28T19:10:00Z',
            author: { displayName: 'Alex' },
          },
          {
            id: 'r1',
            content: 'Earlier reply',
            createdTime: '2026-04-28T19:05:00Z',
            author: { displayName: 'Kim' },
          },
        ],
      },
      {
        id: 'c1',
        content: 'First note',
        createdTime: '2026-04-28T18:00:00Z',
        quotedFileContent: { value: 'Alpha roadmap' },
        author: { displayName: 'Sam' },
      },
    ],
    [
      {
        slideIndex: 0,
        objectId: 'slide-1',
        title: 'Roadmap',
        text: 'Alpha roadmap launch sequence',
      },
    ],
    {
      presentationTitle: 'Deck A',
      generatedAt: '2026-04-28T20:00:00Z',
    }
  );

  assert.deepEqual(result.groups, [
    {
      type: 'slide',
      key: 'slide-1',
      heading: 'Slide 1 - Roadmap',
      slideIndex: 0,
      slideObjectId: 'slide-1',
      threadCount: 2,
      threads: [
        {
          id: 'c1',
          authorName: 'Sam',
          createdTimeLabel: '2026-04-28 18:00 UTC',
          content: 'First note',
          replies: [],
        },
        {
          id: 'c2',
          authorName: 'Mo',
          createdTimeLabel: '2026-04-28 19:00 UTC',
          content: 'Second note',
          replies: [
            {
              id: 'r1',
              authorName: 'Kim',
              createdTimeLabel: '2026-04-28 19:05 UTC',
              content: 'Earlier reply',
            },
            {
              id: 'r2',
              authorName: 'Alex',
              createdTimeLabel: '2026-04-28 19:10 UTC',
              content: 'Later reply',
            },
          ],
        },
      ],
    },
  ]);
  assert.equal(result.summary.totalThreads, 2);
  assert.equal(result.summary.ambiguousThreads, 0);
  assert.equal(result.summary.unmatchedThreads, 0);
});

runTest('routes duplicate quoted text matches into ambiguous comments', function () {
  const result = convertCommentsSlide.buildCommentExportModel(
    [
      {
        id: 'c1',
        content: 'Needs clarification',
        createdTime: '2026-04-28T18:00:00Z',
        quotedFileContent: { value: 'Shared phrase' },
      },
    ],
    [
      { slideIndex: 0, objectId: 'slide-1', title: 'One', text: 'Shared phrase alpha' },
      { slideIndex: 1, objectId: 'slide-2', title: 'Two', text: 'beta Shared phrase gamma' },
    ],
    {}
  );

  assert.equal(result.groups.length, 1);
  assert.equal(result.groups[0].type, 'ambiguous');
  assert.equal(result.groups[0].heading, 'Ambiguous Comments');
  assert.equal(result.summary.ambiguousThreads, 1);
});

runTest('routes comments without a unique match into unmatched comments', function () {
  const result = convertCommentsSlide.buildCommentExportModel(
    [
      {
        id: 'c1',
        content: 'No quote',
      },
      {
        id: 'c2',
        content: 'Missing slide',
        quotedFileContent: { value: 'No such text' },
      },
    ],
    [{ slideIndex: 0, objectId: 'slide-1', title: 'One', text: 'Alpha text' }],
    {}
  );

  assert.equal(result.groups.length, 1);
  assert.equal(result.groups[0].type, 'unmatched');
  assert.equal(result.groups[0].threadCount, 2);
  assert.equal(result.summary.unmatchedThreads, 2);
});

runTest('renders export lines with slide headings and nested replies', function () {
  const lines = convertCommentsSlide.renderExportLines({
    presentationTitle: 'Deck A',
    generatedAtLabel: '2026-04-28 20:00 UTC',
    summary: {
      totalThreads: 1,
      ambiguousThreads: 0,
      unmatchedThreads: 0,
    },
    groups: [
      {
        type: 'slide',
        heading: 'Slide 1 - Roadmap',
        threads: [
          {
            authorName: 'Mo',
            createdTimeLabel: '2026-04-28 19:00 UTC',
            content: 'Top level',
            replies: [
              {
                authorName: 'Alex',
                createdTimeLabel: '2026-04-28 19:05 UTC',
                content: 'Nested reply',
              },
            ],
          },
        ],
      },
    ],
  });

  assert.deepEqual(lines, [
    'Comments Export',
    'Presentation: Deck A | Generated: 2026-04-28 20:00 UTC | Threads: 1',
    '',
    'Slide 1 - Roadmap',
    'Person: Mo | Time: 2026-04-28 19:00 UTC',
    'Top level',
    '  Reply - Person: Alex | Time: 2026-04-28 19:05 UTC',
    '  Nested reply',
  ]);
});

runTest('ships a single Apps Script file that defines menu and slide-link writing', function () {
  const codeSource = fs.readFileSync(
    '/Users/mo.li/Workspace/appscript/comvert-comments-slide/Code.gs',
    'utf8'
  );

  assert.match(codeSource, /createMenu\('Hydra Tools'\)/);
  assert.match(codeSource, /Export Comments To Slide/);
  assert.match(codeSource, /setLinkSlide\(/);
});
