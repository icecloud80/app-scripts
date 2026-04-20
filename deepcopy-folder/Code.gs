/**
 * 作用：
 * 独立提供 Google Drive 文件夹递归复制与 Google Docs 内部链接重写能力。
 *
 * 为什么这样写：
 * 当前项目只保留这个单一功能模块，
 * 通过把入口、复制逻辑和链接改写逻辑集中在同一文件里，可以让维护和部署都保持直接。
 *
 * 输入：
 * @param {void} 无。
 *
 * 输出：
 * @returns {void} 通过公开函数对 Google Drive / Docs 生效。
 *
 * 注意：
 * - 该模块默认以菜单形式挂到绑定文档 UI 中。
 * - 当前只改写复制后 Google Docs 正文中的内部 Google URL。
 */
var CopyFolderApp = (function () {
  /**
   * 作用：
   * 定义复制文件夹工作流的固定默认值。
   *
   * 为什么这样写：
   * 源文件夹默认 ID 和目标根目录名称都是用户显式要求的产品行为，
   * 集中放在常量里更容易测试，也避免散落在 UI 和复制逻辑中。
   *
   * 输入：
   * @param {void} 无。
   *
   * 输出：
   * @returns {Object} 复制文件夹默认配置。
   *
   * 注意：
   * - `sourceFolderId` 会在用户留空输入时作为回退值。
   * - `copiedRootFolderName` 保持固定字面值，不跟源文件夹名称联动。
   */
  const COPY_FOLDER_DEFAULTS = {
    copiedRootFolderPrefix: 'Copy of ',
    menuLabel: 'Copy Folder',
    menuActionLabel: 'Copy Folder and Replace Link',
    copyOnlyMenuActionLabel: 'Copy Folder Only',
  };

  /**
   * 作用：
   * 返回 Copy Folder 独立菜单的定义。
   *
   * 为什么这样写：
   * 菜单结构集中定义后，后续如果要调整菜单标题或动作文案，只需要改一个地方。
   *
   * 输入：
   * @param {void} 无。
   *
   * 输出：
   * @returns {Object} 菜单标题和菜单项定义。
   *
   * 注意：
   * - `menuLabel` 是顶层菜单名。
   * - `functionName` 必须和公开入口函数一致。
   */
  function getCopyFolderMenuConfig_() {
    return {
      menuLabel: COPY_FOLDER_DEFAULTS.menuLabel,
      items: [
        {
          label: COPY_FOLDER_DEFAULTS.menuActionLabel,
          functionName: 'copyFolderWithPrompt',
        },
        {
          label: COPY_FOLDER_DEFAULTS.copyOnlyMenuActionLabel,
          functionName: 'copyFolderOnlyWithPrompt',
        },
      ],
    };
  }

  /**
   * 作用：
   * 在当前文档 UI 中注册 Copy Folder 独立菜单。
   *
   * 为什么这样写：
   * 当前项目只保留一个菜单功能，直接注册独立菜单最清晰。
   *
   * 输入：
   * @param {Object=} ui - 可选 Docs UI 对象，便于测试或被外部 onOpen 复用。
   *
   * 输出：
   * @returns {void} 直接修改 Docs UI。
   *
   * 注意：
   * - 如果未传入 `ui`，会自动使用 `DocumentApp.getUi()`。
   * - 当前菜单只有一个动作入口，但后续可继续扩展。
   */
  function registerCopyFolderMenu_(ui) {
    const docsUi = ui || DocumentApp.getUi();
    const menuConfig = getCopyFolderMenuConfig_();
    const menu = docsUi.createMenu(menuConfig.menuLabel);

    for (let itemIndex = 0; itemIndex < menuConfig.items.length; itemIndex += 1) {
      menu.addItem(
        menuConfig.items[itemIndex].label,
        menuConfig.items[itemIndex].functionName
      );
    }

    menu.addToUi();
  }

  /**
   * 作用：
   * 提供 Copy Folder 菜单入口，收集用户输入后执行递归复制与链接改写。
   *
   * 为什么这样写：
   * 用户希望直接从 Google Docs 菜单触发整个工作流；
   * 这里把 UI 交互、复制执行和结果提示组合成一个薄入口，便于 Apps Script 使用。
   *
   * 输入：
   * @param {void} 无。
   *
   * 输出：
   * @returns {Object|null} 复制结果；若用户取消则返回 `null`。
   *
   * 注意：
   * - 用户在任一步点击取消时，不会执行任何复制。
 * - 源文件夹和目标文件夹输入都必须提供。
  */
  function copyFolderWithPrompt() {
    return runCopyFolderFlow_(copyFolderAndRewriteInternalLinks_);
  }

  /**
   * 作用：
   * 提供 Copy Folder Only 菜单入口，收集用户输入后只执行递归复制。
   *
   * 为什么这样写：
   * 新菜单项和原有“复制并改写链接”流程共用同一套 source / target 输入，
   * 但执行阶段只需要复制目录树，不应该触碰任何文档链接。
   *
   * 输入：
   * @param {void} 无。
   *
   * 输出：
   * @returns {Object|null} 复制结果；若用户取消则返回 `null`。
   *
   * 注意：
   * - 该入口会沿用同一套输入校验。
   * - 结果摘要会明确说明跳过了链接改写。
   */
  function copyFolderOnlyWithPrompt() {
    return runCopyFolderFlow_(copyFolderOnly_);
  }

  /**
   * 作用：
   * 执行菜单触发的共享复制流程，包括收集输入、执行复制动作和提示结果。
   *
   * 为什么这样写：
   * 两个菜单项的差异只在“复制完成后是否改写链接”，
   * 把共用的 UI 交互和结果提示抽出来能避免重复逻辑，同时让新增动作保持最小改动。
   *
   * 输入：
   * @param {Function} copyOperation - 接收 source/target folder id 的具体复制函数。
   *
   * 输出：
   * @returns {Object|null} 执行结果；若用户取消则返回 `null`。
   *
   * 注意：
   * - `copyOperation` 必须返回可以被 `formatCopyFolderSummary_` 处理的结果对象。
   * - 该函数本身不关心是否改写链接，只负责串联流程。
   */
  function runCopyFolderFlow_(copyOperation) {
    const ui = DocumentApp.getUi();
    const request = promptForCopyFolderRequest_(ui);

    if (!request) {
      return null;
    }

    const result = copyOperation(
      request.sourceFolderId,
      request.targetFolderId
    );

    ui.alert(formatCopyFolderSummary_(result));
    Logger.log(JSON.stringify(result));
    return result;
  }

  /**
   * 作用：
   * 顺序收集复制文件夹所需的 source / target folder id。
   *
   * 为什么这样写：
   * Docs 原生 prompt 足以满足当前两步输入需求，
   * 不需要为首版功能再引入单独的 HTML 对话框与前后端通信。
   *
   * 输入：
   * @param {Object} ui - Google Docs UI 对象。
   *
   * 输出：
   * @returns {Object|null} 用户输入结果；取消时返回 `null`。
   *
   * 注意：
   * - source / target folder id 都不允许为空。
   * - 当前不再提供默认 source folder id。
   */
  function promptForCopyFolderRequest_(ui) {
    const sourceResponse = ui.prompt(
      COPY_FOLDER_DEFAULTS.menuLabel,
      '请输入 source folder A id。',
      ui.ButtonSet.OK_CANCEL
    );

    if (sourceResponse.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const targetResponse = ui.prompt(
      COPY_FOLDER_DEFAULTS.menuLabel,
      '请输入目标 folder B id。',
      ui.ButtonSet.OK_CANCEL
    );

    if (targetResponse.getSelectedButton() !== ui.Button.OK) {
      return null;
    }

    const sourceFolderId = normalizeFolderIdInput_(sourceResponse.getResponseText());
    const targetFolderId = normalizeFolderIdInput_(
      targetResponse.getResponseText()
    );

    if (!sourceFolderId) {
      throw new Error('Source folder A id is required.');
    }

    if (!targetFolderId) {
      throw new Error('Target folder B id is required.');
    }

    return {
      sourceFolderId: sourceFolderId,
      targetFolderId: targetFolderId,
    };
  }

  /**
   * 作用：
   * 执行目录树复制，并把复制后 Google Docs 中的内部链接改写到新资源。
   *
   * 为什么这样写：
   * 复制和链接回写本质上是同一个用户动作的两个阶段；
   * 先完成整棵树复制，再统一回写链接，可以确保所有新资源 ID 都已经可用。
   *
   * 输入：
   * @param {string} sourceFolderId - 源文件夹 ID。
   * @param {string} targetFolderId - 目标父文件夹 ID。
   *
   * 输出：
   * @returns {Object} 复制与改写结果摘要。
   *
   * 注意：
   * - 结果中的 `folderCount` 包含新建根目录本身。
   * - 链接改写会扫描复制后 Google Docs 的所有可访问 tab 正文。
   */
  function copyFolderAndRewriteInternalLinks_(sourceFolderId, targetFolderId) {
    const copyState = copyFolderTree_(sourceFolderId, targetFolderId);
    const resourceIdMap = buildCombinedResourceIdMap_(
      copyState.folderIdMap,
      copyState.fileIdMap
    );
    let rewrittenDocumentCount = 0;
    const aggregateLinkStats = createEmptyLinkStats_();

    for (
      let documentIndex = 0;
      documentIndex < copyState.copyStats.copiedDocumentIds.length;
      documentIndex += 1
    ) {
      const rewriteResult = rewriteDocumentInternalLinks_(
        copyState.copyStats.copiedDocumentIds[documentIndex],
        resourceIdMap
      );

      mergeLinkStats_(aggregateLinkStats, rewriteResult);

      if (rewriteResult.replacedLinkCount > 0) {
        rewrittenDocumentCount += 1;
      }
    }

    return buildCopyFolderResult_(
      copyState,
      {
        scannedDocumentCount: copyState.copyStats.copiedDocumentIds.length,
        rewrittenDocumentCount: rewrittenDocumentCount,
        totalLinkCount: aggregateLinkStats.totalLinkCount,
        googleLinkCount: aggregateLinkStats.googleLinkCount,
        replacedLinkCount: aggregateLinkStats.replacedLinkCount,
        unmappedGoogleLinkCount: aggregateLinkStats.unmappedGoogleLinkCount,
        unsupportedGoogleLinkCount: aggregateLinkStats.unsupportedGoogleLinkCount,
        linkRewriteEnabled: true,
      }
    );
  }

  /**
   * 作用：
   * 只执行目录树复制，不改写复制后文档中的任何内部链接。
   *
   * 为什么这样写：
   * `Copy Folder Only` 需要复用现有递归复制能力，但必须显式跳过任何文档读写，
   * 这样用户可以在只关心文件复制时避免额外扫描和改写成本。
   *
   * 输入：
   * @param {string} sourceFolderId - 源文件夹 ID。
   * @param {string} targetFolderId - 目标父文件夹 ID。
   *
   * 输出：
   * @returns {Object} 仅复制目录树的结果摘要。
   *
   * 注意：
   * - 返回值中的链接统计字段会全部为 0。
   * - `linkRewriteEnabled` 会显式标记为 `false`，便于 UI 摘要区分两种模式。
   */
  function copyFolderOnly_(sourceFolderId, targetFolderId) {
    const copyState = copyFolderTree_(sourceFolderId, targetFolderId);
    return buildCopyFolderResult_(
      copyState,
      {
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
  }

  /**
   * 作用：
   * 完成目录树复制，并返回后续结果汇总所需的所有中间状态。
   *
   * 为什么这样写：
   * “只复制”和“复制并改写链接”两条流程都依赖同一份复制结果，
   * 把共享的 Drive 复制阶段抽成单独函数后，可以遵守 SRP 并减少重复逻辑。
   *
   * 输入：
   * @param {string} sourceFolderId - 源文件夹 ID。
   * @param {string} targetFolderId - 目标父文件夹 ID。
   *
   * 输出：
   * @returns {Object} 包含复制根目录、映射表和统计信息的中间状态。
   *
   * 注意：
   * - 该函数只负责复制 Drive 资源，不负责改写 Docs 内容。
   * - 会把复制出的 Google Docs ID 记录在 `copyStats.copiedDocumentIds` 里供上层决定是否使用。
   */
  function copyFolderTree_(sourceFolderId, targetFolderId) {
    const sourceFolder = DriveApp.getFolderById(sourceFolderId);
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    const copiedRootFolderName = buildCopiedRootFolderName_(sourceFolder.getName());
    const copiedRootFolder = targetFolder.createFolder(
      copiedRootFolderName
    );
    const fileIdMap = {};
    const folderIdMap = {};
    const copyStats = {
      fileCount: 0,
      folderCount: 1,
      copiedDocumentIds: [],
    };
    let rewrittenDocumentCount = 0;
    const aggregateLinkStats = createEmptyLinkStats_();

    folderIdMap[sourceFolder.getId()] = copiedRootFolder.getId();

    copyFolderContentsRecursive_(
      sourceFolder,
      copiedRootFolder,
      fileIdMap,
      folderIdMap,
      copyStats
    );

    return {
      sourceFolder: sourceFolder,
      targetFolder: targetFolder,
      copiedRootFolder: copiedRootFolder,
      copiedRootFolderName: copiedRootFolderName,
      fileIdMap: fileIdMap,
      folderIdMap: folderIdMap,
      copyStats: copyStats,
    };
  }

  /**
   * 作用：
   * 根据复制阶段状态和执行模式补充字段，生成统一的结果对象。
   *
   * 为什么这样写：
   * 两种菜单模式都需要返回同一结构的结果摘要，
   * 统一组装函数能保证字段稳定，避免在多个执行路径里手写同一批计数键。
   *
   * 输入：
   * @param {Object} copyState - `copyFolderTree_` 返回的复制状态。
   * @param {Object} extraResultFields - 当前执行模式补充的统计字段。
   *
   * 输出：
   * @returns {Object} 统一结构的复制结果。
   *
   * 注意：
   * - Drive 复制计数永远来自 `copyState.copyStats`。
   * - 额外字段会覆盖默认的链接统计值，便于不同模式复用。
   */
  function buildCopyFolderResult_(copyState, extraResultFields) {
    return Object.assign(
      {
        sourceFolderId: copyState.sourceFolder.getId(),
        targetFolderId: copyState.targetFolder.getId(),
        copiedRootFolderId: copyState.copiedRootFolder.getId(),
        copiedRootFolderName: copyState.copiedRootFolderName,
        fileCount: copyState.copyStats.fileCount,
        folderCount: copyState.copyStats.folderCount,
        scannedDocumentCount: 0,
        rewrittenDocumentCount: 0,
        totalLinkCount: 0,
        googleLinkCount: 0,
        replacedLinkCount: 0,
        unmappedGoogleLinkCount: 0,
        unsupportedGoogleLinkCount: 0,
        linkRewriteEnabled: false,
      },
      extraResultFields
    );
  }

  /**
   * 作用：
   * 根据源文件夹名称构造复制后的根目录名称。
   *
   * 为什么这样写：
   * 用户要求目标目录名应为 `Copy of + 原目录名`，
   * 不能再使用示例文案 `folder A` 作为固定名称。
   *
   * 输入：
   * @param {string} sourceFolderName - 源文件夹名称。
   *
   * 输出：
   * @returns {string} 复制后的根目录名称。
   *
   * 注意：
   * - 这里保留源名称原样，不额外清洗。
   * - 前缀由常量统一管理。
   */
  function buildCopiedRootFolderName_(sourceFolderName) {
    return COPY_FOLDER_DEFAULTS.copiedRootFolderPrefix + sourceFolderName;
  }

  /**
   * 作用：
   * 创建空的链接统计对象。
   *
   * 为什么这样写：
   * 复制摘要需要同时展示总链接数、Google 链接数和未替换原因，
   * 统一结构后更容易在文本节点、文档级和整次运行之间累计。
   *
   * 输入：
   * @param {void} 无。
   *
   * 输出：
   * @returns {Object} 空统计对象。
   *
   * 注意：
   * - 所有计数字段都从 0 开始。
   * - 该对象可以被 `mergeLinkStats_` 原地累加。
   */
  function createEmptyLinkStats_() {
    return {
      totalLinkCount: 0,
      googleLinkCount: 0,
      replacedLinkCount: 0,
      unmappedGoogleLinkCount: 0,
      unsupportedGoogleLinkCount: 0,
    };
  }

  /**
   * 作用：
   * 把一份链接统计累加到另一份统计对象里。
   *
   * 为什么这样写：
   * 链接统计会在文本、正文、文档和整次复制四个层级逐步累计，
   * 用统一 merge 可以避免重复写字段累加代码。
   *
   * 输入：
   * @param {Object} targetStats - 目标统计对象。
   * @param {Object} sourceStats - 待累加统计对象。
   *
   * 输出：
   * @returns {Object} 累加后的目标统计对象。
   *
   * 注意：
   * - 会原地修改 `targetStats`。
   * - `sourceStats` 缺失字段时按 0 处理。
   */
  function mergeLinkStats_(targetStats, sourceStats) {
    const statKeys = [
      'totalLinkCount',
      'googleLinkCount',
      'replacedLinkCount',
      'unmappedGoogleLinkCount',
      'unsupportedGoogleLinkCount',
    ];

    for (let statIndex = 0; statIndex < statKeys.length; statIndex += 1) {
      targetStats[statKeys[statIndex]] += sourceStats[statKeys[statIndex]] || 0;
    }

    return targetStats;
  }

  /**
   * 作用：
   * 递归复制源目录下的所有子文件夹和文件，并建立旧 ID 到新 ID 的映射。
   *
   * 为什么这样写：
   * 用户要求保留完整目录层级，而且后续链接改写依赖全量映射；
   * 因此复制阶段必须同步记录每个资源的新旧 ID 对应关系。
   *
   * 输入：
   * @param {Object} sourceFolder - 源文件夹对象。
   * @param {Object} targetFolder - 目标文件夹对象。
   * @param {Object} fileIdMap - 文件映射表。
   * @param {Object} folderIdMap - 文件夹映射表。
   * @param {Object} copyStats - 复制统计对象。
   * @param {Object=} mimeTypes - 可选 MIME 类型常量对象，便于本地测试注入。
   *
   * 输出：
   * @returns {void} 通过传入对象累积结果。
   *
   * 注意：
   * - 这里不负责创建目标根目录，调用方需要先创建。
   * - 复制顺序先子文件夹后文件，便于保持递归层级清晰。
   */
  function copyFolderContentsRecursive_(
    sourceFolder,
    targetFolder,
    fileIdMap,
    folderIdMap,
    copyStats,
    mimeTypes
  ) {
    const mimeTypeConfig = mimeTypes || (typeof MimeType === 'undefined' ? {} : MimeType);
    const childFolders = sourceFolder.getFolders();
    const childFiles = sourceFolder.getFiles();

    while (childFolders.hasNext()) {
      const sourceChildFolder = childFolders.next();
      const copiedChildFolder = targetFolder.createFolder(sourceChildFolder.getName());

      folderIdMap[sourceChildFolder.getId()] = copiedChildFolder.getId();
      copyStats.folderCount += 1;

      copyFolderContentsRecursive_(
        sourceChildFolder,
        copiedChildFolder,
        fileIdMap,
        folderIdMap,
        copyStats,
        mimeTypeConfig
      );
    }

    while (childFiles.hasNext()) {
      const sourceFile = childFiles.next();
      const copiedFile = sourceFile.makeCopy(sourceFile.getName(), targetFolder);

      fileIdMap[sourceFile.getId()] = copiedFile.getId();
      copyStats.fileCount += 1;

      if (
        typeof sourceFile.getMimeType === 'function' &&
        sourceFile.getMimeType() === mimeTypeConfig.GOOGLE_DOCS
      ) {
        copyStats.copiedDocumentIds.push(copiedFile.getId());
      }
    }
  }

  /**
   * 作用：
   * 合并文件夹与文件映射，得到统一的资源 ID 映射表。
   *
   * 为什么这样写：
   * 文档内链接可能同时指向文件或文件夹；
   * 统一成一张映射表后，后续 URL 改写逻辑不需要区分资源类别。
   *
   * 输入：
   * @param {Object} folderIdMap - 文件夹映射表。
   * @param {Object} fileIdMap - 文件映射表。
   *
   * 输出：
   * @returns {Object} 合并后的资源映射表。
   *
   * 注意：
   * - 同名 key 以后写入值为准，但正常情况下文件和文件夹 ID 不应冲突。
   */
  function buildCombinedResourceIdMap_(folderIdMap, fileIdMap) {
    const combinedMap = {};
    const folderIds = Object.keys(folderIdMap);
    const fileIds = Object.keys(fileIdMap);

    for (let folderIndex = 0; folderIndex < folderIds.length; folderIndex += 1) {
      combinedMap[folderIds[folderIndex]] = folderIdMap[folderIds[folderIndex]];
    }

    for (let fileIndex = 0; fileIndex < fileIds.length; fileIndex += 1) {
      combinedMap[fileIds[fileIndex]] = fileIdMap[fileIds[fileIndex]];
    }

    return combinedMap;
  }

  /**
   * 作用：
   * 重写单个复制后 Google Docs 正文中的内部链接。
   *
   * 为什么这样写：
   * 复制阶段只负责生成新文件，真正让新目录树“可用”的关键是把正文里的旧资源链接切换到新资源。
   *
   * 输入：
   * @param {string} documentId - 复制后的 Google Docs 文档 ID。
   * @param {Object} resourceIdMap - 旧资源 ID 到新资源 ID 的映射表。
   *
   * 输出：
   * @returns {Object} 当前文档的改写结果。
   *
   * 注意：
   * - 当前处理所有可访问 tab 的正文；若运行时不支持 tabs，会回退到 `doc.getBody()`。
   * - 返回值会同时包含总链接数、Google 链接数和未替换原因统计。
   */
  function rewriteDocumentInternalLinks_(documentId, resourceIdMap) {
    const document = DocumentApp.openById(documentId);
    const bodies = getDocumentBodies_(document);
    const linkStats = createEmptyLinkStats_();

    for (let bodyIndex = 0; bodyIndex < bodies.length; bodyIndex += 1) {
      mergeLinkStats_(linkStats, rewriteContainerLinks_(bodies[bodyIndex], resourceIdMap));
    }

    document.saveAndClose();

    return Object.assign({ documentId: documentId }, linkStats);
  }

  /**
   * 作用：
   * 读取文档中所有可访问的正文容器。
   *
   * 为什么这样写：
   * Google Docs 现在可能包含多个 tabs；
   * 如果只扫描默认 body，会漏掉其它 tab 中的链接，这正是首次运行可能“部分替换、部分没替换”的高概率原因。
   *
   * 输入：
   * @param {Object} document - 打开的文档对象。
   *
   * 输出：
   * @returns {Array<Object>} 可遍历的正文列表。
   *
   * 注意：
   * - 若运行时不支持 `getTabs()`，会回退到单一 `getBody()`。
   * - 只收集能成功取到 body 的 tab。
   */
  function getDocumentBodies_(document) {
    const documentBodies = [];

    if (typeof document.getTabs === 'function') {
      const tabs = document.getTabs();

      for (let tabIndex = 0; tabIndex < tabs.length; tabIndex += 1) {
        if (
          tabs[tabIndex] &&
          typeof tabs[tabIndex].asDocumentTab === 'function'
        ) {
          const documentTab = tabs[tabIndex].asDocumentTab();

          if (documentTab && typeof documentTab.getBody === 'function') {
            documentBodies.push(documentTab.getBody());
          }
        }
      }
    }

    if (!documentBodies.length && typeof document.getBody === 'function') {
      documentBodies.push(document.getBody());
    }

    return documentBodies;
  }

  /**
   * 作用：
   * 递归遍历正文容器，累计其中所有文本节点的链接改写次数。
   *
   * 为什么这样写：
   * Google Docs 正文可能包含段落、列表、表格等多层嵌套结构，
   * 必须用递归遍历才能稳定覆盖文本节点。
   *
   * 输入：
   * @param {Object} container - 正文或子容器对象。
   * @param {Object} resourceIdMap - 资源映射表。
   *
   * 输出：
   * @returns {Object} 当前容器内的链接统计。
   *
   * 注意：
   * - 不可编辑或无子节点的元素会自动跳过。
   * - 文本节点和容器节点会走不同分支，避免重复改写。
   */
  function rewriteContainerLinks_(container, resourceIdMap) {
    if (!container) {
      return createEmptyLinkStats_();
    }

    return rewriteElementLinks_(container, resourceIdMap);
  }

  /**
   * 作用：
   * 递归处理单个元素，识别文本节点并执行链接改写。
   *
   * 为什么这样写：
   * 通过一个统一入口处理正文树，能同时兼容直接文本节点和多层容器节点。
   *
   * 输入：
   * @param {Object} element - 待处理元素。
   * @param {Object} resourceIdMap - 资源映射表。
   *
   * 输出：
   * @returns {Object} 当前元素及其子元素的链接统计。
   *
   * 注意：
   * - 若元素既可遍历 children，又可 `editAsText()`，优先走 children 路径，避免重复计数。
   * - 仅在没有 children 时才尝试 `editAsText()` 回退。
   */
  function rewriteElementLinks_(element, resourceIdMap) {
    const linkStats = createEmptyLinkStats_();

    if (!element) {
      return linkStats;
    }

    if (isTextElementLike_(element)) {
      return rewriteTextElementLinks_(element, resourceIdMap);
    }

    if (typeof element.getNumChildren === 'function' && typeof element.getChild === 'function') {
      for (let childIndex = 0; childIndex < element.getNumChildren(); childIndex += 1) {
        mergeLinkStats_(
          linkStats,
          rewriteElementLinks_(element.getChild(childIndex), resourceIdMap)
        );
      }

      return linkStats;
    }

    if (typeof element.editAsText === 'function') {
      return rewriteTextElementLinks_(element.editAsText(), resourceIdMap);
    }

    return linkStats;
  }

  /**
   * 作用：
   * 判断一个对象是否具备可直接改写链接的 Text 接口。
   *
   * 为什么这样写：
   * Apps Script 不同元素类型的能力差异很大，先做结构判断可以避免在非文本节点上误调用文本方法。
   *
   * 输入：
   * @param {*} element - 待判断对象。
   *
   * 输出：
   * @returns {boolean} 是否是可处理的 Text-like 对象。
   *
   * 注意：
   * - 这里使用结构判断，而不是强依赖具体 `ElementType.TEXT` 常量。
   * - 这样 Node 本地测试也能复用。
   */
  function isTextElementLike_(element) {
    return Boolean(
      element &&
      typeof element.getText === 'function' &&
      typeof element.getTextAttributeIndices === 'function' &&
      typeof element.getLinkUrl === 'function' &&
      typeof element.setLinkUrl === 'function'
    );
  }

  /**
   * 作用：
   * 改写单个文本节点中的富文本链接和裸文本 URL。
   *
   * 为什么这样写：
   * 一个 Text 节点里既可能存在真正的 hyperlink 属性，也可能存在只是普通字符的裸文本 URL；
   * 这两类都需要覆盖，才能让复制出的文档内部互链完整可用。
   *
   * 输入：
   * @param {Object} textElement - Text-like 对象。
   * @param {Object} resourceIdMap - 资源映射表。
   *
   * 输出：
   * @returns {Object} 当前文本节点的链接统计。
   *
   * 注意：
   * - 富文本链接先改，裸文本 URL 后改。
   * - 裸文本替换会从后往前处理，避免位移破坏后续索引。
   */
  function rewriteTextElementLinks_(textElement, resourceIdMap) {
    const linkStats = createEmptyLinkStats_();
    const attributeIndices = textElement.getTextAttributeIndices();
    const textLength = textElement.getText().length;

    for (let index = 0; index < attributeIndices.length; index += 1) {
      const startOffset = attributeIndices[index];
      const endOffset = index + 1 < attributeIndices.length
        ? attributeIndices[index + 1] - 1
        : textLength - 1;
      const currentLinkUrl = textElement.getLinkUrl(startOffset);

      if (currentLinkUrl) {
        mergeLinkStats_(
          linkStats,
          buildLinkClassificationStats_(currentLinkUrl, resourceIdMap)
        );
      }

      const rewrittenLinkUrl = rewriteMappedGoogleUrl_(currentLinkUrl, resourceIdMap);

      if (currentLinkUrl && rewrittenLinkUrl !== currentLinkUrl) {
        textElement.setLinkUrl(startOffset, endOffset, rewrittenLinkUrl);
      }
    }

    mergeLinkStats_(
      linkStats,
      rewritePlainGoogleUrlsInTextElement_(textElement, resourceIdMap)
    );

    return linkStats;
  }

  /**
   * 作用：
   * 改写文本节点里未绑定 hyperlink 属性的裸文本 Google URL。
   *
   * 为什么这样写：
   * 真实文档中很多链接只是直接粘贴的 URL 文本，不会带富文本链接属性；
   * 如果不额外扫描，复制后仍会指向旧资源。
   *
   * 输入：
   * @param {Object} textElement - Text-like 对象。
   * @param {Object} resourceIdMap - 资源映射表。
   *
   * 输出：
   * @returns {Object} 裸文本 URL 的链接统计。
   *
   * 注意：
   * - 会跳过已经带 hyperlink 属性的字符范围，避免重复改写。
   * - 采用从后往前的顺序，保证前面范围的索引稳定。
   */
  function rewritePlainGoogleUrlsInTextElement_(textElement, resourceIdMap) {
    const matches = findUrlMatches_(textElement.getText());
    const linkStats = createEmptyLinkStats_();

    for (let matchIndex = matches.length - 1; matchIndex >= 0; matchIndex -= 1) {
      const match = matches[matchIndex];
      if (hasLinkUrlInRange_(textElement, match.startOffset, match.endOffset)) {
        continue;
      }

      mergeLinkStats_(linkStats, buildLinkClassificationStats_(match.url, resourceIdMap));
      const rewrittenUrl = rewriteMappedGoogleUrl_(match.url, resourceIdMap);

      if (rewrittenUrl !== match.url) {
        textElement.deleteText(match.startOffset, match.endOffset);
        textElement.insertText(match.startOffset, rewrittenUrl);
      }
    }

    return linkStats;
  }

  /**
   * 作用：
   * 判断一段字符范围内是否已经存在 hyperlink 属性。
   *
   * 为什么这样写：
   * 同一段 URL 既可能显示为原始 URL 文本，又绑定了 hyperlink；
   * 裸文本替换时跳过这类范围，可以避免和 hyperlink 改写重复计数。
   *
   * 输入：
   * @param {Object} textElement - Text-like 对象。
   * @param {number} startOffset - 起始字符索引。
   * @param {number} endOffset - 结束字符索引。
   *
   * 输出：
   * @returns {boolean} 范围内是否存在 hyperlink。
   *
   * 注意：
   * - 这里按字符逐位检查，优先保证正确性。
   * - URL 长度通常不大，可接受线性扫描。
   */
  function hasLinkUrlInRange_(textElement, startOffset, endOffset) {
    for (let offset = startOffset; offset <= endOffset; offset += 1) {
      if (textElement.getLinkUrl(offset)) {
        return true;
      }
    }

    return false;
  }

  /**
   * 作用：
   * 扫描一段字符串中的 URL，并返回可替换的精确字符范围。
   *
   * 为什么这样写：
   * 文本节点里的 URL 需要按字符区间替换；
   * 先把 URL 与尾随标点拆开，能避免把句末标点误当成 URL 内容。
   *
   * 输入：
   * @param {string} text - 待扫描字符串。
   *
   * 输出：
   * @returns {Array<Object>} URL 匹配结果列表。
   *
   * 注意：
   * - 这里只做通用 URL 扫描，不在此阶段判断是否是 Google URL。
   * - 返回的 `endOffset` 指向 URL 字符本身，不包含尾随标点。
   */
  function findUrlMatches_(text) {
    const matches = [];
    const urlRegex = /https?:\/\/[^\s<>"']+/g;
    let match = null;

    while ((match = urlRegex.exec(text)) !== null) {
      const token = splitTrailingUrlPunctuation_(match[0]);

      if (!token.url) {
        continue;
      }

      matches.push({
        url: token.url,
        startOffset: match.index,
        endOffset: match.index + token.url.length - 1,
      });
    }

    return matches;
  }

  /**
   * 作用：
   * 从 URL token 末尾剥离常见句读标点。
   *
   * 为什么这样写：
   * 实际文档中的链接经常紧跟句号、右括号等标点，
   * 如果不先剥离，后续资源 ID 提取和替换都可能失败。
   *
   * 输入：
   * @param {string} token - 正则匹配出来的原始 URL token。
   *
   * 输出：
   * @returns {Object} 拆分后的 URL 主体与尾随标点。
   *
   * 注意：
   * - 这里只移除常见结尾标点，不处理 URL 中间字符。
   * - 返回的 `trailingText` 供纯字符串替换时保留原排版。
   */
  function splitTrailingUrlPunctuation_(token) {
    let trimmedToken = token || '';
    let trailingText = '';

    while (trimmedToken && /[).,!?:;\]]/.test(trimmedToken.charAt(trimmedToken.length - 1))) {
      trailingText = trimmedToken.charAt(trimmedToken.length - 1) + trailingText;
      trimmedToken = trimmedToken.slice(0, -1);
    }

    return {
      url: trimmedToken,
      trailingText: trailingText,
    };
  }

  /**
   * 作用：
   * 对一段纯字符串中的 Google URL 执行映射替换。
   *
   * 为什么这样写：
   * 这部分逻辑与 DocumentApp 无关，单独抽出来后可以直接在 Node 测试里验证 URL 改写规则。
   *
   * 输入：
   * @param {string} text - 原始字符串。
   * @param {Object} resourceIdMap - 资源映射表。
   *
   * 输出：
   * @returns {Object} 替换后的文本和替换次数。
   *
   * 注意：
   * - 只会改写命中映射的 Google URL。
   * - 非 Google URL 或未命中映射的 URL 会保持原样。
   */
  function replaceMappedGoogleUrlsInText_(text, resourceIdMap) {
    const linkStats = createEmptyLinkStats_();
    const replacedText = (text || '').replace(
      /https?:\/\/[^\s<>"']+/g,
      function (token) {
        const tokenParts = splitTrailingUrlPunctuation_(token);
        mergeLinkStats_(linkStats, buildLinkClassificationStats_(tokenParts.url, resourceIdMap));
        const rewrittenUrl = rewriteMappedGoogleUrl_(tokenParts.url, resourceIdMap);
        return rewrittenUrl + tokenParts.trailingText;
      }
    );

    return Object.assign({ text: replacedText }, linkStats);
  }

  /**
   * 作用：
   * 根据一条 URL 生成可累计的链接分类统计。
   *
   * 为什么这样写：
   * 用户现在不仅要知道“替换了多少”，还要知道“一共有多少链接”和“为什么没替换”；
   * 因此每命中一条 URL 都需要先做一次分类。
   *
   * 输入：
   * @param {string|null} url - 待分析链接。
   * @param {Object} resourceIdMap - 资源映射表。
   *
   * 输出：
   * @returns {Object} 单条链接对应的统计结果。
   *
   * 注意：
   * - 所有 URL 都计入 `totalLinkCount`。
   * - 只有 Google URL 才会继续细分为 replaced / unmapped / unsupported。
   */
  function buildLinkClassificationStats_(url, resourceIdMap) {
    const linkStats = createEmptyLinkStats_();
    const resourceId = extractGoogleDriveResourceIdFromUrl_(url);

    if (!url) {
      return linkStats;
    }

    linkStats.totalLinkCount += 1;

    if (!isGoogleUrl_(url)) {
      return linkStats;
    }

    linkStats.googleLinkCount += 1;

    if (!resourceId) {
      linkStats.unsupportedGoogleLinkCount += 1;
      return linkStats;
    }

    if (!resourceIdMap[resourceId]) {
      linkStats.unmappedGoogleLinkCount += 1;
      return linkStats;
    }

    linkStats.replacedLinkCount += 1;
    return linkStats;
  }

  /**
   * 作用：
   * 判断一条 URL 是否属于 Google Docs / Drive 域名。
   *
   * 为什么这样写：
   * 统计阶段需要先区分“普通外链”和“Google 链接但没被替换”，
   * 这样结果才能解释为什么有些链接没改。
   *
   * 输入：
   * @param {string|null} url - 待判断 URL。
   *
   * 输出：
   * @returns {boolean} 是否是 Google Docs / Drive 链接。
   *
   * 注意：
   * - 这里只判断当前需求相关的 docs.google.com / drive.google.com。
   * - 解析失败时返回 `false`。
   */
  function isGoogleUrl_(url) {
    return Boolean(
      url &&
      (
        url.indexOf('https://docs.google.com/') === 0 ||
        url.indexOf('http://docs.google.com/') === 0 ||
        url.indexOf('https://drive.google.com/') === 0 ||
        url.indexOf('http://drive.google.com/') === 0
      )
    );
  }

  /**
   * 作用：
   * 将单个 Google URL 中的旧资源 ID 替换为新资源 ID。
   *
   * 为什么这样写：
   * 只替换 URL 内的资源 ID，而不重建整条 URL，
   * 可以尽量保留原有的参数、锚点和打开模式。
   *
   * 输入：
   * @param {string|null} url - 原始 URL。
   * @param {Object} resourceIdMap - 资源映射表。
   *
   * 输出：
   * @returns {string|null} 改写后的 URL；无变化时返回原值。
   *
   * 注意：
   * - 非 Google URL 会原样返回。
   * - 只有命中映射时才替换。
   */
  function rewriteMappedGoogleUrl_(url, resourceIdMap) {
    const resourceId = extractGoogleDriveResourceIdFromUrl_(url);

    if (!resourceId || !resourceIdMap[resourceId]) {
      return url;
    }

    return url.replace(resourceId, resourceIdMap[resourceId]);
  }

  /**
   * 作用：
   * 从常见 Google Drive / Docs URL 中提取资源 ID。
   *
   * 为什么这样写：
   * 复制后的链接改写建立在“先识别旧资源 ID，再映射到新 ID”的模式上；
   * 统一 URL 解析入口可以减少分支重复。
   *
   * 输入：
   * @param {string|null} url - 待解析 URL。
   *
   * 输出：
   * @returns {string|null} 解析出的资源 ID；无法识别时返回 `null`。
   *
   * 注意：
   * - 当前只覆盖首版明确支持的常见 Google URL 形态。
   * - 新 URL 变体可在这里继续追加规则。
   */
  function extractGoogleDriveResourceIdFromUrl_(url) {
    const patterns = [
      /https?:\/\/docs\.google\.com\/(?:document|spreadsheets|presentation|forms)\/(?:u\/\d+\/)?d\/([a-zA-Z0-9_-]+)/,
      /https?:\/\/drive\.google\.com\/(?:u\/\d+\/)?file\/d\/([a-zA-Z0-9_-]+)/,
      /https?:\/\/drive\.google\.com\/drive\/(?:u\/\d+\/)?folders\/([a-zA-Z0-9_-]+)/,
      /https?:\/\/drive\.google\.com\/open\?[^#\s]*\bid=([a-zA-Z0-9_-]+)/,
    ];

    if (!url) {
      return null;
    }

    for (let patternIndex = 0; patternIndex < patterns.length; patternIndex += 1) {
      const match = url.match(patterns[patternIndex]);

      if (match) {
        return match[1];
      }
    }

    return null;
  }

  /**
   * 作用：
   * 规范化用户输入的 folder id。
   *
   * 为什么这样写：
   * prompt 输入常带前后空白，而且源文件夹支持“空输入回退默认值”；
   * 单独抽成函数后更容易测试和复用。
   *
   * 输入：
   * @param {string} input - 原始输入。
   *
   * 输出：
   * @returns {string} 规范化后的 folder id；空输入返回空字符串。
   *
   * 注意：
   * - 不再承担默认值注入逻辑。
   * - 非字符串输入会被当作空值处理。
   */
  function normalizeFolderIdInput_(input) {
    const normalizedValue = typeof input === 'string' ? input.trim() : '';
    return normalizedValue;
  }

  /**
   * 作用：
   * 生成复制完成后的摘要文案。
   *
   * 为什么这样写：
   * 菜单触发流程需要给用户即时反馈；
   * 用一个集中函数生成摘要，可以让 UI 提示和日志输出保持一致口径。
   *
   * 输入：
   * @param {Object} result - 复制结果对象。
   *
   * 输出：
   * @returns {string} 可直接展示给用户的摘要文本。
   *
   * 注意：
   * - 文案保持简洁，但要把用户关心的“总链接数”和“未替换原因”展示出来。
   * - 返回值同时适合作为 alert 文本和日志辅助信息。
   */
  function formatCopyFolderSummary_(result) {
    const summaryLines = [
      'Copy Folder completed.',
      'Root folder: ' + result.copiedRootFolderName,
      'Root folder id: ' + result.copiedRootFolderId,
      'Folders copied: ' + result.folderCount,
      'Files copied: ' + result.fileCount,
    ];

    if (!result.linkRewriteEnabled) {
      summaryLines.push('Internal link rewrite: skipped');
      return summaryLines.join('\n');
    }

    summaryLines.push(
      'Docs scanned: ' + result.scannedDocumentCount,
      'Total links found: ' + result.totalLinkCount,
      'Google links found: ' + result.googleLinkCount,
      'Docs rewritten: ' + result.rewrittenDocumentCount,
      'Links replaced: ' + result.replacedLinkCount,
      'Google links not replaced (outside copied tree): ' + result.unmappedGoogleLinkCount,
      'Google links not replaced (unsupported format): ' + result.unsupportedGoogleLinkCount
    );

    return summaryLines.join('\n');
  }

  /**
   * 作用：
   * 让 Node 测试环境可以加载同一份 copy-folder 逻辑代码。
   *
   * 为什么这样写：
   * 通过统一的测试导出对象，可以让 Apps Script 版本和本地测试版本保持同一实现。
   *
   * 输入：
   * @param {void} 无。
   *
   * 输出：
   * @returns {Object} 纯逻辑测试导出。
   *
   * 注意：
   * - 只导出当前独立功能真正需要的接口。
   * - 不暴露不必要的全局状态。
   */
  function getTestExports_() {
    return {
      buildCopiedRootFolderName_: buildCopiedRootFolderName_,
      buildLinkClassificationStats_: buildLinkClassificationStats_,
      createEmptyLinkStats_: createEmptyLinkStats_,
      COPY_FOLDER_DEFAULTS: COPY_FOLDER_DEFAULTS,
      getDocumentBodies_: getDocumentBodies_,
      buildCombinedResourceIdMap_: buildCombinedResourceIdMap_,
      copyFolderAndRewriteInternalLinks_: copyFolderAndRewriteInternalLinks_,
      copyFolderOnly_: copyFolderOnly_,
      copyFolderContentsRecursive_: copyFolderContentsRecursive_,
      extractGoogleDriveResourceIdFromUrl_: extractGoogleDriveResourceIdFromUrl_,
      findUrlMatches_: findUrlMatches_,
      formatCopyFolderSummary_: formatCopyFolderSummary_,
      getCopyFolderMenuConfig_: getCopyFolderMenuConfig_,
      hasLinkUrlInRange_: hasLinkUrlInRange_,
      isGoogleUrl_: isGoogleUrl_,
      isTextElementLike_: isTextElementLike_,
      mergeLinkStats_: mergeLinkStats_,
      normalizeFolderIdInput_: normalizeFolderIdInput_,
      promptForCopyFolderRequest_: promptForCopyFolderRequest_,
      replaceMappedGoogleUrlsInText_: replaceMappedGoogleUrlsInText_,
      rewriteContainerLinks_: rewriteContainerLinks_,
      rewriteDocumentInternalLinks_: rewriteDocumentInternalLinks_,
      rewriteMappedGoogleUrl_: rewriteMappedGoogleUrl_,
      rewritePlainGoogleUrlsInTextElement_: rewritePlainGoogleUrlsInTextElement_,
      rewriteTextElementLinks_: rewriteTextElementLinks_,
      splitTrailingUrlPunctuation_: splitTrailingUrlPunctuation_,
    };
  }

  return {
    copyFolderOnlyWithPrompt: copyFolderOnlyWithPrompt,
    copyFolderWithPrompt: copyFolderWithPrompt,
    registerCopyFolderMenu_: registerCopyFolderMenu_,
    __test__: getTestExports_(),
  };
}());

/**
 * 作用：
 * 提供 Apps Script 公共入口，弹出输入框后递归复制文件夹并重写复制后 Docs 的内部链接。
 *
 * 为什么这样写：
 * 用户要求从单独菜单直接发起整个复制流程；
 * 保留顶层薄包装入口最符合 Apps Script 的调用方式，也方便和菜单项绑定。
 *
 * 输入：
 * @param {void} 无。
 *
 * 输出：
 * @returns {Object|null} 复制结果；取消时返回 `null`。
 *
 * 注意：
 * - 源文件夹输入为空时会自动使用内置默认 source folder id。
 * - 该函数会请求 Drive 与 Docs 的读写权限。
 */
function copyFolderWithPrompt() {
  return CopyFolderApp.copyFolderWithPrompt();
}

/**
 * 作用：
 * 提供 Apps Script 公共入口，弹出输入框后只递归复制文件夹，不改写复制后 Docs 的内部链接。
 *
 * 为什么这样写：
 * 新菜单项需要一个清晰的顶层绑定函数，供 Apps Script 菜单直接调用，
 * 同时保持和原有入口一致的调用方式与授权边界。
 *
 * 输入：
 * @param {void} 无。
 *
 * 输出：
 * @returns {Object|null} 复制结果；取消时返回 `null`。
 *
 * 注意：
 * - 该函数仍会复制完整目录树，但不会打开或修改任何复制后的 Google Docs。
 * - 该函数会请求 Drive 的读写权限。
 */
function copyFolderOnlyWithPrompt() {
  return CopyFolderApp.copyFolderOnlyWithPrompt();
}

/**
 * 作用：
 * 提供 Apps Script 菜单注册辅助函数。
 *
 * 为什么这样写：
 * 把菜单注册逻辑单独抽出来，可以让 `onOpen()` 保持很薄，也方便本地测试复用同一实现。
 *
 * 输入：
 * @param {Object=} ui - 可选 Docs UI 对象。
 *
 * 输出：
 * @returns {void} 直接修改 Docs UI。
 *
 * 注意：
 * - 如果 `onOpen()` 未调用这个函数，则 copy-folder 菜单不会出现。
 * - 该函数只注册菜单，不执行复制逻辑。
 */
function registerCopyFolderMenu_(ui) {
  return CopyFolderApp.registerCopyFolderMenu_(ui);
}

/**
 * 作用：
 * 在文档打开时自动注册 Copy Folder 菜单。
 *
 * 为什么这样写：
 * 现在整个项目只保留 copy-folder 功能，因此需要让单一入口脚本自己负责菜单注册，
 * 而不是再依赖其它功能文件里的共享 `onOpen()`。
 *
 * 输入：
 * @param {void} 无。
 *
 * 输出：
 * @returns {void} 直接修改当前文档 UI。
 *
 * 注意：
 * - 菜单会在文档刷新或重新打开后出现。
 * - 该函数不执行复制，只负责把入口挂到菜单中。
 */
function onOpen() {
  registerCopyFolderMenu_(DocumentApp.getUi());
}

/**
 * 作用：
 * 让 Node 测试环境可以加载同一份 copy-folder 逻辑代码。
 *
 * 为什么这样写：
 * 复用一份核心逻辑，避免 Apps Script 文件和测试文件双份维护。
 *
 * 输入：
 * @param {void} 无。
 *
 * 输出：
 * @returns {void} 在 Node 环境下挂载导出对象。
 *
 * 注意：
 * - Apps Script 运行时没有 `module`，因此必须做环境判断。
 * - 只导出 copy-folder 功能相关接口，避免和旧脚本混用。
 */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = CopyFolderApp.__test__;
}
