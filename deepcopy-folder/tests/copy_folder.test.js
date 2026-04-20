/**
 * 作用：
 * 验证 copy-folder 独立脚本的纯逻辑与菜单结构。
 *
 * 为什么这样写：
 * 该项目现在只保留 copy-folder 功能，需要用回归用例锁定默认值、递归复制和链接改写行为。
 *
 * 输入：
 * @param {void} 无。
 *
 * 输出：
 * @returns {void} 通过抛错让测试进程失败。
 *
 * 注意：
 * - 这里只测试纯逻辑，不测试真实 Google Drive / Docs 服务。
 * - 如果菜单结构或 URL 规则变化，需要同步更新这里的断言。
 */
const assert = require('node:assert/strict');
const copyFolder = require('../src/copy_folder.js');

/**
 * 作用：
 * 封装单个测试用例执行。
 *
 * 为什么这样写：
 * 当前仓库使用轻量 Node 测试方式，不依赖第三方框架也能给出清晰输出。
 *
 * 输入：
 * @param {string} title - 用例标题。
 * @param {Function} testFn - 测试函数。
 *
 * 输出：
 * @returns {void} 输出测试结果。
 *
 * 注意：
 * - 失败时会重新抛错，保证 `npm test` 返回非零状态码。
 * - 不要在这里吞掉异常。
 */
function runTest(title, testFn) {
  try {
    testFn();
    console.log('PASS', title);
  } catch (error) {
    console.error('FAIL', title);
    throw error;
  }
}

/**
 * 作用：
 * 为测试构造最小 Drive iterator mock。
 *
 * 为什么这样写：
 * Apps Script 的 `FolderIterator` / `FileIterator` 在 Node 环境不存在，
 * 这里用最小兼容接口把递归复制逻辑固定在本地测试里。
 *
 * 输入：
 * @param {Array<*>} items - 迭代项列表。
 *
 * 输出：
 * @returns {Object} 带 `hasNext/next` 的最小 iterator。
 *
 * 注意：
 * - 顺序与传入数组保持一致。
 * - 超界访问会抛错，避免测试静默通过。
 */
function createMockIterator_(items) {
  let index = 0;

  return {
    hasNext: function () {
      return index < items.length;
    },
    next: function () {
      if (index >= items.length) {
        throw new Error('Iterator exhausted');
      }

      const item = items[index];
      index += 1;
      return item;
    },
  };
}

/**
 * 作用：
 * 为测试构造最小 Drive file mock。
 *
 * 为什么这样写：
 * 递归复制逻辑依赖 `getId/getName/getMimeType/makeCopy`，
 * 用轻量 mock 就能覆盖主要行为而不需要真实 Drive 访问。
 *
 * 输入：
 * @param {string} id - 源文件 ID。
 * @param {string} name - 文件名。
 * @param {string} mimeType - 文件 MIME 类型。
 * @param {Array<Object>} copySink - 收集复制行为的数组。
 *
 * 输出：
 * @returns {Object} 最小 file mock。
 *
 * 注意：
 * - `makeCopy` 会生成稳定可断言的新文件 ID。
 * - `targetFolder` 会被记录到 `copySink` 便于断言。
 */
function createMockFile_(id, name, mimeType, copySink) {
  return {
    getId: function () {
      return id;
    },
    getName: function () {
      return name;
    },
    getMimeType: function () {
      return mimeType;
    },
    makeCopy: function (copiedName, targetFolder) {
      const copiedFile = {
        getId: function () {
          return 'copy-of-' + id;
        },
        getName: function () {
          return copiedName;
        },
        getMimeType: function () {
          return mimeType;
        },
      };

      copySink.push({
        sourceId: id,
        copiedName: copiedName,
        targetFolder: targetFolder,
      });
      return copiedFile;
    },
  };
}

/**
 * 作用：
 * 为测试构造最小 Drive folder mock。
 *
 * 为什么这样写：
 * copy-folder 功能需要递归走 `getFolders/getFiles/createFolder`，
 * 本地测试里只需要保留最小结构和迭代能力。
 *
 * 输入：
 * @param {string} id - 文件夹 ID。
 * @param {string} name - 文件夹名称。
 * @param {Array<Object>} folders - 子文件夹列表。
 * @param {Array<Object>} files - 文件列表。
 *
 * 输出：
 * @returns {Object} 最小 folder mock。
 *
 * 注意：
 * - `createFolder` 默认创建空文件夹，可在测试里覆盖。
 * - 这里不模拟权限或父目录关系。
 */
function createMockFolder_(id, name, folders, files) {
  return {
    getId: function () {
      return id;
    },
    getName: function () {
      return name;
    },
    getFolders: function () {
      return createMockIterator_(folders);
    },
    getFiles: function () {
      return createMockIterator_(files);
    },
    createFolder: function (childName) {
      return createMockFolder_('created-' + childName, childName, [], []);
    },
  };
}

/**
 * 作用：
 * 为测试构造最小 Text-like mock。
 *
 * 为什么这样写：
 * 链接改写同时依赖 `getTextAttributeIndices/getLinkUrl/setLinkUrl/deleteText/insertText`，
 * 用一个可变字符串 mock 就能在 Node 环境覆盖关键行为。
 *
 * 输入：
 * @param {string} text - 初始文本。
 * @param {Array<Object>=} linkedRanges - 已有 hyperlink 范围配置。
 *
 * 输出：
 * @returns {Object} Text-like mock。
 *
 * 注意：
 * - `linkedRanges` 中每项包含 `start/end/url`。
 * - 为保持测试最小实现，文本替换后不会重建整套 link offset 映射。
 */
function createMockTextElement_(text, linkedRanges) {
  let textValue = text;
  const linkByOffset = {};
  const setLinkCalls = [];
  const ranges = linkedRanges || [];

  for (let rangeIndex = 0; rangeIndex < ranges.length; rangeIndex += 1) {
    for (let offset = ranges[rangeIndex].start; offset <= ranges[rangeIndex].end; offset += 1) {
      linkByOffset[offset] = ranges[rangeIndex].url;
    }
  }

  return {
    getText: function () {
      return textValue;
    },
    getTextAttributeIndices: function () {
      const boundaries = [0];

      for (let rangeIndex = 0; rangeIndex < ranges.length; rangeIndex += 1) {
        boundaries.push(ranges[rangeIndex].start);
        boundaries.push(ranges[rangeIndex].end + 1);
      }

      return boundaries
        .filter(function (value) {
          return value >= 0 && value < textValue.length;
        })
        .sort(function (left, right) {
          return left - right;
        })
        .filter(function (value, index, values) {
          return index === 0 || values[index - 1] !== value;
        });
    },
    getLinkUrl: function (offset) {
      return Object.prototype.hasOwnProperty.call(linkByOffset, offset)
        ? linkByOffset[offset]
        : null;
    },
    setLinkUrl: function (startOffset, endOffset, url) {
      setLinkCalls.push({
        startOffset: startOffset,
        endOffset: endOffset,
        url: url,
      });

      for (let offset = startOffset; offset <= endOffset; offset += 1) {
        linkByOffset[offset] = url;
      }
    },
    deleteText: function (startOffset, endOffset) {
      textValue = textValue.slice(0, startOffset) + textValue.slice(endOffset + 1);
    },
    insertText: function (startOffset, value) {
      textValue = textValue.slice(0, startOffset) + value + textValue.slice(startOffset);
    },
    getSetLinkCalls: function () {
      return setLinkCalls.slice();
    },
  };
}

runTest('getCopyFolderMenuConfig_ returns standalone menu definition', function () {
  assert.deepEqual(copyFolder.getCopyFolderMenuConfig_(), {
    menuLabel: 'Copy Folder',
    items: [
      {
        label: 'Copy Folder and Replace Link',
        functionName: 'copyFolderWithPrompt',
      },
      {
        label: 'Copy Folder Only',
        functionName: 'copyFolderOnlyWithPrompt',
      },
    ],
  });
});

runTest('copy folder defaults expose source id and root folder name', function () {
  assert.equal(
    copyFolder.COPY_FOLDER_DEFAULTS.copiedRootFolderPrefix,
    'Copy of '
  );
  assert.equal(
    copyFolder.COPY_FOLDER_DEFAULTS.menuActionLabel,
    'Copy Folder and Replace Link'
  );
  assert.equal(
    copyFolder.COPY_FOLDER_DEFAULTS.copyOnlyMenuActionLabel,
    'Copy Folder Only'
  );
});

runTest('normalizeFolderIdInput_ trims values and preserves empty input without default', function () {
  assert.equal(
    copyFolder.normalizeFolderIdInput_('  abc123  '),
    'abc123'
  );
  assert.equal(
    copyFolder.normalizeFolderIdInput_('   '),
    ''
  );
});

runTest('promptForCopyFolderRequest_ requires both source and target folder ids', function () {
  const ui = {
    Button: {
      OK: 'ok',
    },
    ButtonSet: {
      OK_CANCEL: 'ok_cancel',
    },
    prompt: function (_title, message) {
      if (message.indexOf('source folder A id') !== -1) {
        return {
          getSelectedButton: function () {
            return 'ok';
          },
          getResponseText: function () {
            return '   ';
          },
        };
      }

      return {
        getSelectedButton: function () {
          return 'ok';
        },
        getResponseText: function () {
          return 'target-folder';
        },
      };
    },
  };

  assert.throws(
    function () {
      copyFolder.promptForCopyFolderRequest_(ui);
    },
    /Source folder A id is required/
  );
});

runTest('extractGoogleDriveResourceIdFromUrl_ handles common google url formats', function () {
  assert.equal(
    copyFolder.extractGoogleDriveResourceIdFromUrl_(
      'https://docs.google.com/document/d/oldDocId123/edit'
    ),
    'oldDocId123'
  );
  assert.equal(
    copyFolder.extractGoogleDriveResourceIdFromUrl_(
      'https://drive.google.com/open?id=spreadsheet456'
    ),
    'spreadsheet456'
  );
  assert.equal(
    copyFolder.extractGoogleDriveResourceIdFromUrl_(
      'https://drive.google.com/drive/folders/folder789'
    ),
    'folder789'
  );
  assert.equal(
    copyFolder.extractGoogleDriveResourceIdFromUrl_(
      'https://docs.google.com/document/u/0/d/docWithUserPath/edit?tab=t.0'
    ),
    'docWithUserPath'
  );
  assert.equal(
    copyFolder.extractGoogleDriveResourceIdFromUrl_('https://example.com/not-google'),
    null
  );
});

runTest('replaceMappedGoogleUrlsInText_ rewrites only mapped google urls', function () {
  const replacement = copyFolder.replaceMappedGoogleUrlsInText_(
    [
      'See https://docs.google.com/document/d/oldDocId123/edit and',
      'https://example.com/leave-me plus https://drive.google.com/drive/folders/oldFolder456.',
    ].join(' '),
    {
      oldDocId123: 'newDocId999',
      oldFolder456: 'newFolder888',
    }
  );

  assert.equal(
    replacement.text,
    [
      'See https://docs.google.com/document/d/newDocId999/edit and',
      'https://example.com/leave-me plus https://drive.google.com/drive/folders/newFolder888.',
    ].join(' ')
  );
  assert.deepEqual(
    {
      totalLinkCount: replacement.totalLinkCount,
      googleLinkCount: replacement.googleLinkCount,
      replacedLinkCount: replacement.replacedLinkCount,
      unmappedGoogleLinkCount: replacement.unmappedGoogleLinkCount,
      unsupportedGoogleLinkCount: replacement.unsupportedGoogleLinkCount,
    },
    {
      totalLinkCount: 3,
      googleLinkCount: 2,
      replacedLinkCount: 2,
      unmappedGoogleLinkCount: 0,
      unsupportedGoogleLinkCount: 0,
    }
  );
});

runTest('copyFolderContentsRecursive_ copies nested folders and records mappings', function () {
  const createdFolders = [];
  const copiedFiles = [];
  const targetFolder = {
    createFolder: function (name) {
      const createdFolder = createMockFolder_('new-folder-' + name, name, [], []);
      createdFolders.push(name);
      return createdFolder;
    },
  };
  const sourceFolder = createMockFolder_(
    'source-root',
    'root',
    [
      createMockFolder_(
        'source-child',
        'child',
        [],
        [
          createMockFile_('old-doc-1', 'child-doc', 'application/vnd.google-apps.document', copiedFiles),
        ]
      ),
    ],
    [
      createMockFile_('old-file-1', 'top-doc', 'application/vnd.google-apps.document', copiedFiles),
      createMockFile_('old-file-2', 'image', 'image/png', copiedFiles),
    ]
  );
  const fileIdMap = {};
  const folderIdMap = {};
  const copyStats = {
    fileCount: 0,
    folderCount: 0,
    copiedDocumentIds: [],
  };

  copyFolder.copyFolderContentsRecursive_(
    sourceFolder,
    targetFolder,
    fileIdMap,
    folderIdMap,
    copyStats,
    {
      GOOGLE_DOCS: 'application/vnd.google-apps.document',
    }
  );

  assert.deepEqual(createdFolders, ['child']);
  assert.deepEqual(Object.keys(folderIdMap), ['source-child']);
  assert.deepEqual(Object.keys(fileIdMap).sort(), ['old-doc-1', 'old-file-1', 'old-file-2']);
  assert.equal(copyStats.folderCount, 1);
  assert.equal(copyStats.fileCount, 3);
  assert.equal(copyStats.copiedDocumentIds.length, 2);
});

runTest('rewriteTextElementLinks_ rewrites hyperlink urls and plain google urls', function () {
  const textElement = createMockTextElement_(
    'See label https://docs.google.com/document/d/oldDocId123/edit.',
    [
      {
        start: 4,
        end: 8,
        url: 'https://drive.google.com/drive/folders/oldFolder456',
      },
    ]
  );
  const replaceCount = copyFolder.rewriteTextElementLinks_(textElement, {
    oldDocId123: 'newDocId999',
    oldFolder456: 'newFolder888',
  });

  assert.deepEqual(replaceCount, {
    totalLinkCount: 2,
    googleLinkCount: 2,
    replacedLinkCount: 2,
    unmappedGoogleLinkCount: 0,
    unsupportedGoogleLinkCount: 0,
  });
  assert.deepEqual(textElement.getSetLinkCalls(), [
    {
      startOffset: 4,
      endOffset: 8,
      url: 'https://drive.google.com/drive/folders/newFolder888',
    },
  ]);
  assert.equal(
    textElement.getText(),
    'See label https://docs.google.com/document/d/newDocId999/edit.'
  );
});

runTest('rewriteDocumentInternalLinks_ scans all document tabs, not only default body', function () {
  const originalDocumentApp = global.DocumentApp;
  const firstTabText = createMockTextElement_(
    'Jump https://docs.google.com/document/d/old-doc-1/edit',
    []
  );
  const secondTabText = createMockTextElement_(
    'Visit https://docs.google.com/document/u/0/d/old-doc-2/edit',
    []
  );

  global.DocumentApp = {
    openById: function () {
      return {
        getTabs: function () {
          return [
            {
              asDocumentTab: function () {
                return {
                  getBody: function () {
                    return {
                      getNumChildren: function () {
                        return 1;
                      },
                      getChild: function () {
                        return firstTabText;
                      },
                    };
                  },
                };
              },
            },
            {
              asDocumentTab: function () {
                return {
                  getBody: function () {
                    return {
                      getNumChildren: function () {
                        return 1;
                      },
                      getChild: function () {
                        return secondTabText;
                      },
                    };
                  },
                };
              },
            },
          ];
        },
        saveAndClose: function () {
          return null;
        },
      };
    },
  };

  assert.deepEqual(
    copyFolder.rewriteDocumentInternalLinks_('copied-doc', {
      'old-doc-1': 'new-doc-1',
      'old-doc-2': 'new-doc-2',
    }),
    {
      documentId: 'copied-doc',
      totalLinkCount: 2,
      googleLinkCount: 2,
      replacedLinkCount: 2,
      unmappedGoogleLinkCount: 0,
      unsupportedGoogleLinkCount: 0,
    }
  );
  assert.equal(
    firstTabText.getText(),
    'Jump https://docs.google.com/document/d/new-doc-1/edit'
  );
  assert.equal(
    secondTabText.getText(),
    'Visit https://docs.google.com/document/u/0/d/new-doc-2/edit'
  );

  global.DocumentApp = originalDocumentApp;
});

runTest('copyFolderAndRewriteInternalLinks_ returns aggregated copy and rewrite counts', function () {
  const originalDriveApp = global.DriveApp;
  const originalDocumentApp = global.DocumentApp;
  const originalMimeType = global.MimeType;
  const copiedFiles = [];
  const sourceFolder = createMockFolder_(
    'source-folder',
    'source',
    [],
    [
      createMockFile_('old-doc-1', 'doc-a', 'application/vnd.google-apps.document', copiedFiles),
      createMockFile_('old-file-2', 'sheet-b', 'application/vnd.google-apps.spreadsheet', copiedFiles),
    ]
  );
  const targetRootFolder = {
    getId: function () {
      return 'target-folder';
    },
    createFolder: function (name) {
      return createMockFolder_('new-root-folder', name, [], []);
    },
  };
  const copiedDocument = {
    getTabs: function () {
      return [
        {
          asDocumentTab: function () {
            return {
              getBody: function () {
                return {
                  getNumChildren: function () {
                    return 1;
                  },
                  getChild: function () {
                    return createMockTextElement_(
                      [
                        'Jump https://docs.google.com/document/d/old-doc-1/edit',
                        'and https://example.com/not-counted-for-google',
                      ].join(' '),
                      []
                    );
                  },
                };
              },
            };
          },
        },
      ];
    },
    saveAndClose: function () {
      return null;
    },
  };

  global.MimeType = {
    GOOGLE_DOCS: 'application/vnd.google-apps.document',
  };
  global.DriveApp = {
    getFolderById: function (folderId) {
      if (folderId === 'source-folder') {
        return sourceFolder;
      }

      if (folderId === 'target-folder') {
        return targetRootFolder;
      }

      throw new Error('Unexpected folder id: ' + folderId);
    },
  };
  global.DocumentApp = {
    openById: function (documentId) {
      if (documentId === 'copy-of-old-doc-1') {
        return copiedDocument;
      }

      throw new Error('Unexpected document id: ' + documentId);
    },
  };

  assert.deepEqual(
    copyFolder.copyFolderAndRewriteInternalLinks_('source-folder', 'target-folder'),
    {
      sourceFolderId: 'source-folder',
      targetFolderId: 'target-folder',
      copiedRootFolderId: 'new-root-folder',
      copiedRootFolderName: 'Copy of source',
      fileCount: 2,
      folderCount: 1,
      scannedDocumentCount: 1,
      rewrittenDocumentCount: 1,
      totalLinkCount: 2,
      googleLinkCount: 1,
      replacedLinkCount: 1,
      unmappedGoogleLinkCount: 0,
      unsupportedGoogleLinkCount: 0,
      linkRewriteEnabled: true,
    }
  );

  global.DriveApp = originalDriveApp;
  global.DocumentApp = originalDocumentApp;
  global.MimeType = originalMimeType;
});

runTest('copyFolderOnly_ returns copy counts without scanning docs or rewriting links', function () {
  const originalDriveApp = global.DriveApp;
  const copiedFiles = [];
  const sourceFolder = createMockFolder_(
    'source-folder',
    'source',
    [],
    [
      createMockFile_('old-doc-1', 'doc-a', 'application/vnd.google-apps.document', copiedFiles),
      createMockFile_('old-file-2', 'sheet-b', 'application/vnd.google-apps.spreadsheet', copiedFiles),
    ]
  );
  const targetRootFolder = {
    getId: function () {
      return 'target-folder';
    },
    createFolder: function (name) {
      return createMockFolder_('new-root-folder', name, [], []);
    },
  };

  global.DriveApp = {
    getFolderById: function (folderId) {
      if (folderId === 'source-folder') {
        return sourceFolder;
      }

      if (folderId === 'target-folder') {
        return targetRootFolder;
      }

      throw new Error('Unexpected folder id: ' + folderId);
    },
  };

  assert.deepEqual(
    copyFolder.copyFolderOnly_('source-folder', 'target-folder'),
    {
      sourceFolderId: 'source-folder',
      targetFolderId: 'target-folder',
      copiedRootFolderId: 'new-root-folder',
      copiedRootFolderName: 'Copy of source',
      fileCount: 2,
      folderCount: 1,
      scannedDocumentCount: 0,
      rewrittenDocumentCount: 0,
      totalLinkCount: 0,
      googleLinkCount: 0,
      replacedLinkCount: 0,
      unmappedGoogleLinkCount: 0,
      unsupportedGoogleLinkCount: 0,
      linkRewriteEnabled: false,
    }
  );

  global.DriveApp = originalDriveApp;
});

runTest('formatCopyFolderSummary_ explains when link rewrite is skipped', function () {
  assert.equal(
    copyFolder.formatCopyFolderSummary_({
      copiedRootFolderName: 'Copy of source',
      copiedRootFolderId: 'new-root-folder',
      folderCount: 1,
      fileCount: 2,
      linkRewriteEnabled: false,
    }),
    [
      'Copy Folder completed.',
      'Root folder: Copy of source',
      'Root folder id: new-root-folder',
      'Folders copied: 1',
      'Files copied: 2',
      'Internal link rewrite: skipped',
    ].join('\n')
  );
});
