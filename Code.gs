function doGet(e) {
  const params = parseShortcutParams_(e);
  if (hasShortcutData_(params)) {
    return addShortcutItem_(params);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Study Shorts')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover, user-scalable=no');
}

function doPost(e) {
  return addShortcutItem_(parseShortcutParams_(e));
}

function getSheetNames() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()
    .map((sheet) => sheet.getName());
}

function getBookTreeLayout() {
  const sheetNames = getSheetNames();
  const raw = PropertiesService.getDocumentProperties().getProperty('bookTreeLayout') || '';
  let layout = null;

  try {
    layout = raw ? JSON.parse(raw) : null;
  } catch (error) {
    layout = null;
  }

  return normalizeBookTreeLayout_(layout, sheetNames);
}

function saveBookTreeLayout(layout) {
  const normalized = normalizeBookTreeLayout_(layout, getSheetNames());
  PropertiesService
    .getDocumentProperties()
    .setProperty('bookTreeLayout', JSON.stringify(normalized));
  return normalized;
}

function authorizeDriveAccess() {
  DriveApp.getRootFolder().getName();
  SpreadsheetApp.getActiveSpreadsheet().getName();
  return true;
}

function getStudyItems(sheetName) {
  const sheet = getTargetSheet_(sheetName);
  const resolvedSheetName = sheet.getName();
  const syncResult = syncDriveFolderImagesToSheet_(sheet);
  ensureLetterRanks_(sheet);
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 15);

  if (lastRow < 1) {
    const stats = updateSheetStats_(sheet);
    return {
      items: [],
      deckStats: [{
        sheetName: resolvedSheetName,
        cardCount: 0,
        averageRank: stats.averageRank,
        todayPlayCount: stats.todayPlayCount,
        totalPlayCount: stats.totalPlayCount,
      }],
      backgroundUrls: [],
      backgroundProbability: 1,
      masteredProbability: 0.05,
      fixedBackground: true,
      averageRank: stats.averageRank,
      todayPlayCount: stats.todayPlayCount,
      totalPlayCount: stats.totalPlayCount,
      driveFolders: syncResult.folderCount,
      driveImages: syncResult.imageCount,
      driveAdded: syncResult.added,
      driveRemoved: syncResult.removed,
      driveErrors: syncResult.errors,
    };
  }

  const values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getDisplayValues();
  const questionImageUrls = getQuestionImageUrls_(sheet, lastRow);
  const answerImageUrls = getAnswerImageUrls_(sheet, lastRow);
  const questionDriveUrls = getQuestionDriveUrls_(sheet, lastRow);
  const answerDriveUrls = getAnswerDriveUrls_(sheet, lastRow);
  const backgroundUrls = getBackgroundUrls_(sheet, lastRow);
  const rowBackgroundUrls = getRowBackgroundUrls_(sheet, lastRow);
  const rankNotes = sheet.getRange(1, 4, lastRow, 1).getNotes();
  const rawBackgroundProbability = String(values[1]?.[9] ?? '').trim();
  const backgroundProbability = rawBackgroundProbability === '0' ? 0 : 1;
  const rawMasteredProbability = String(values[1] && values[1][14] || '').trim();
  const masteredProbability = rawMasteredProbability === ''
    ? 0.05
    : normalizeProbability_(rawMasteredProbability);
  const fixedBackground = true;

  const sheetItems = values
    .map((row, index) => {
      const questionImageUrl = questionImageUrls[index] || '';
      const answerImageUrl = answerImageUrls[index] || '';
      return {
        row: index + 1,
        sheetName: resolvedSheetName,
        question: questionImageUrl ? 'Image' : row[0].trim(),
        questionImageUrl,
        questionDriveUrl: questionDriveUrls[index] || '',
        answer: answerImageUrl ? 'Image' : row[1].trim(),
        answerImageUrl,
        answerDriveUrl: answerDriveUrls[index] || '',
        backgroundUrl: rowBackgroundUrls[index] || '',
        comments: row.slice(4, 7).map((comment) => comment.trim()).filter(Boolean),
        skip: Number(row[2]) || 0,
        studyCount: Number(row[2]) || 0,
        rank: normalizeRank_(row[3], rankNotes[index][0]),
      };
    })
    .filter((item) => item.question || item.questionImageUrl || item.answerImageUrl);
  const stats = updateSheetStats_(sheet);

  return {
    items: sheetItems,
    deckStats: [{
      sheetName: resolvedSheetName,
      cardCount: sheetItems.length,
      averageRank: stats.averageRank,
      todayPlayCount: stats.todayPlayCount,
      totalPlayCount: stats.totalPlayCount,
    }],
    backgroundUrls,
    backgroundProbability,
    masteredProbability,
    fixedBackground,
    averageRank: stats.averageRank,
    todayPlayCount: stats.todayPlayCount,
    totalPlayCount: stats.totalPlayCount,
    driveFolders: syncResult.folderCount,
    driveImages: syncResult.imageCount,
    driveAdded: syncResult.added,
    driveRemoved: syncResult.removed,
    driveErrors: syncResult.errors,
  };
}

function getStudyItemsForSheets(sheetNames) {
  const names = Array.isArray(sheetNames)
    ? sheetNames.map((name) => String(name || '').trim()).filter(Boolean)
    : [String(sheetNames || '').trim()].filter(Boolean);
  const uniqueNames = names.filter((name, index) => names.indexOf(name) === index);

  if (uniqueNames.length <= 1) {
    return getStudyItems(uniqueNames[0] || '');
  }

  const payloads = uniqueNames.map((name) => getStudyItems(name));
  const items = [];
  const deckStats = [];
  const backgroundUrls = [];
  const driveErrors = [];
  let driveFolders = 0;
  let driveImages = 0;
  let driveAdded = 0;
  let driveRemoved = 0;
  let todayPlayCount = 0;
  let totalPlayCount = 0;
  let rankSum = 0;
  let rankCount = 0;

  payloads.forEach((payload, payloadIndex) => {
    const sheetName = uniqueNames[payloadIndex];
    const stats = payload.deckStats && payload.deckStats[0]
      ? payload.deckStats[0]
      : {
          sheetName,
          cardCount: (payload.items || []).length,
          averageRank: payload.averageRank,
          todayPlayCount: payload.todayPlayCount,
          totalPlayCount: payload.totalPlayCount,
        };
    deckStats.push(Object.assign({}, stats, { sheetName }));
    (payload.items || []).forEach((item) => {
      const rank = Number(item.rank);
      if (Number.isFinite(rank)) {
        rankSum += rank;
        rankCount += 1;
      }
      items.push(Object.assign({}, item, {
        id: `${sheetName}:${item.row}:${item.question}:${item.answer}`,
        sheetName,
      }));
    });
    (payload.backgroundUrls || []).forEach((url) => {
      if (url && !backgroundUrls.includes(url)) backgroundUrls.push(url);
    });
    driveFolders += Number(payload.driveFolders || 0);
    driveImages += Number(payload.driveImages || 0);
    driveAdded += Number(payload.driveAdded || 0);
    driveRemoved += Number(payload.driveRemoved || 0);
    todayPlayCount += Number(payload.todayPlayCount || 0);
    totalPlayCount += Number(payload.totalPlayCount || 0);
    (payload.driveErrors || []).forEach((error) => {
      driveErrors.push(Object.assign({ sheetName }, error));
    });
  });

  return {
    items,
    sheetNames: uniqueNames,
    deckStats,
    backgroundUrls,
    backgroundProbability: payloads[0] ? payloads[0].backgroundProbability : 1,
    masteredProbability: payloads[0] ? payloads[0].masteredProbability : 0.05,
    fixedBackground: true,
    averageRank: rankCount > 0 ? Math.round((rankSum / rankCount) * 100000000) / 100000000 : '',
    todayPlayCount,
    totalPlayCount,
    driveFolders,
    driveImages,
    driveAdded,
    driveRemoved,
    driveErrors,
  };
}

function setAllStudyRanks(rank, sheetName) {
  const sheet = getTargetSheet_(sheetName);
  const targetRank = normalizeRank_(rank);
  const lastRow = sheet.getLastRow();

  if (lastRow < 1) {
    return { ok: true, count: 0, rank: targetRank };
  }

  const values = sheet.getRange(1, 1, lastRow, 4).getDisplayValues();
  let count = 0;
  const isCard = values.map((row) => Boolean(String(row[0] || row[1] || '').trim()));
  count = isCard.filter(Boolean).length;
  const rankRange = sheet.getRange(1, 4, lastRow, 1);
  const oldNotes = rankRange.getNotes();
  rankRange.setValues(values.map((row, index) => [isCard[index] ? gradeFromRank_(targetRank) : row[3]]));
  rankRange.setNotes(values.map((row, index) => [isCard[index]
    ? rankNote_(oldNotes[index][0], targetRank)
    : oldNotes[index][0]]));
  const stats = updateSheetStats_(sheet);

  return { ok: true, count, rank: targetRank, stats };
}

function setAllStudyRanksForSheets(rank, sheetNames) {
  const names = normalizeSheetNameList_(sheetNames);
  let count = 0;
  let rankValue = normalizeRank_(rank);

  names.forEach((name) => {
    const result = setAllStudyRanks(rankValue, name);
    count += Number(result.count || 0);
    rankValue = result.rank;
  });

  return { ok: true, count, rank: rankValue, sheetCount: names.length };
}

function deleteRankTwoRows(sheetName) {
  const sheet = getTargetSheet_(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow < 1) {
    return { ok: true, deleted: 0 };
  }

  const streak = sheet.getRange('K2').getValue();
  const lastLogin = sheet.getRange('L2').getValue();
  const todayPlayCount = sheet.getRange('N2').getValue();
  const totalPlayCount = sheet.getRange('N4').getValue();
  const values = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 14)).getDisplayValues();
  let deleted = 0;
  let compacted = 0;

  for (let row = lastRow; row >= 1; row -= 1) {
    const rowValues = values[row - 1];
    const rankValue = String(rowValues[3] || '').trim();
    const hasStudyData = rowValues.slice(0, 10).some((value) => String(value || '').trim());
    const hasStatsData = rowValues.slice(10, 14).some((value) => String(value || '').trim());
    const isRankOne = rankValue && normalizeRank_(rankValue) === 1;
    const isBlankSpacer = !hasStudyData && !hasStatsData;

    if (isRankOne || isBlankSpacer) {
      sheet.deleteRow(row);
      if (isRankOne) {
        deleted += 1;
      } else {
        compacted += 1;
      }
    }
  }
  compacted += compactStudyColumns_(sheet);

  sheet.getRange('K2').setValue(streak);
  sheet.getRange('L2').setValue(lastLogin);
  sheet.getRange('N2').setValue(todayPlayCount);
  sheet.getRange('N4').setValue(totalPlayCount);
  const stats = updateSheetStats_(sheet);

  return { ok: true, deleted, compacted, stats };
}

function deleteRankTwoRowsForSheets(sheetNames) {
  const names = normalizeSheetNameList_(sheetNames);
  let deleted = 0;
  let compacted = 0;

  names.forEach((name) => {
    const result = deleteRankTwoRows(name);
    deleted += Number(result.deleted || 0);
    compacted += Number(result.compacted || 0);
  });

  return { ok: true, deleted, compacted, sheetCount: names.length };
}

function deleteStudyCard(row, sheetName) {
  const targetRow = Number(row);

  if (!targetRow || targetRow < 1) {
    return { ok: false, deleted: 0, message: 'invalid row.' };
  }

  const sheet = getTargetSheet_(sheetName);
  const lastRow = sheet.getLastRow();

  if (targetRow > lastRow) {
    return { ok: false, deleted: 0, message: 'row not found.' };
  }

  const controls = getSheetControlValues_(sheet);
  sheet.deleteRow(targetRow);
  restoreSheetControlValues_(sheet, controls);
  const stats = updateSheetStats_(sheet);

  return { ok: true, row: targetRow, deleted: 1, stats };
}

function decrementAllStudyCountsForSheets(sheetNames) {
  const names = normalizeSheetNameList_(sheetNames);
  let updated = 0;

  names.forEach((name) => {
    const result = decrementAllStudyCounts(name);
    updated += Number(result.updated || 0);
  });

  return { ok: true, updated, sheetCount: names.length };
}

function normalizeSheetNameList_(sheetNames) {
  const names = Array.isArray(sheetNames)
    ? sheetNames.map((name) => String(name || '').trim()).filter(Boolean)
    : [String(sheetNames || '').trim()].filter(Boolean);
  return names.filter((name, index) => names.indexOf(name) === index);
}

function normalizeBookTreeLayout_(layout, sheetNames) {
  const validSheetNames = normalizeSheetNameList_(sheetNames);
  const usedSheets = {};

  function cleanNode(node, fallbackIndex) {
    const rawType = String(node && node.type || '').trim();
    const type = rawType === 'folder' ? 'folder' : 'sheet';
    const id = String(node && node.id || `${type}-${fallbackIndex}`).trim();

    if (type === 'folder') {
      return {
        id,
        type: 'folder',
        name: String(node && node.name || 'Folder').trim() || 'Folder',
        children: cleanNodes(node && node.children),
      };
    }

    const sheetName = String(node && node.sheetName || '').trim();
    if (!sheetName || !validSheetNames.includes(sheetName) || usedSheets[sheetName]) {
      return null;
    }

    usedSheets[sheetName] = true;
    return {
      id,
      type: 'sheet',
      name: String(node && node.name || sheetName).trim() || sheetName,
      sheetName,
      children: [],
    };
  }

  function cleanNodes(nodes) {
    return (Array.isArray(nodes) ? nodes : [])
      .map((node, index) => cleanNode(node, index))
      .filter(Boolean);
  }

  const rootChildren = cleanNodes(layout && layout.children);
  validSheetNames.forEach((sheetName) => {
    if (!usedSheets[sheetName]) {
      rootChildren.push({
        id: `sheet-${sheetName}`,
        type: 'sheet',
        name: sheetName,
        sheetName,
        children: [],
      });
    }
  });

  return {
    version: 1,
    children: rootChildren,
  };
}

function getSheetControlValues_(sheet) {
  return {
    backgroundProbability: sheet.getRange('J2').getValue(),
    fixedBackground: sheet.getRange('J4').getValue(),
    streak: sheet.getRange('K2').getValue(),
    lastLogin: sheet.getRange('L2').getValue(),
    todayPlayCount: sheet.getRange('N2').getValue(),
    totalPlayCount: sheet.getRange('N4').getValue(),
    masteredProbability: sheet.getRange('O2').getValue(),
  };
}

function restoreSheetControlValues_(sheet, controls) {
  sheet.getRange('J2').setValue(controls.backgroundProbability);
  sheet.getRange('J4').setValue(controls.fixedBackground);
  sheet.getRange('K2').setValue(controls.streak);
  sheet.getRange('L2').setValue(controls.lastLogin);
  sheet.getRange('N2').setValue(controls.todayPlayCount);
  sheet.getRange('N4').setValue(controls.totalPlayCount);
  sheet.getRange('O2').setValue(controls.masteredProbability);
}

function compactStudyColumns_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  let compacted = 0;
  let writeRow = 1;

  for (let readRow = 1; readRow <= lastRow; readRow += 1) {
    const values = sheet.getRange(readRow, 1, 1, 7).getDisplayValues()[0];
    const hasStudyData = values.some((value) => String(value || '').trim());

    if (!hasStudyData) {
      continue;
    }

    if (readRow !== writeRow) {
      sheet.getRange(readRow, 1, 1, 7).moveTo(sheet.getRange(writeRow, 1, 1, 7));
      compacted += 1;
    }

    writeRow += 1;
  }

  return compacted;
}

function deleteBackgroundUrl(backgroundUrl, sheetName) {
  const targetUrl = normalizeImageUrl_(backgroundUrl);
  if (!targetUrl) {
    return { ok: false, deleted: 0, message: 'missing background url.' };
  }

  const sheet = getTargetSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: true, deleted: 0 };
  }

  const range = sheet.getRange(2, 9, lastRow - 1, 1);
  const displayValues = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const formulas = range.getFormulas();
  let deleted = 0;
  const replacements = [];

  displayValues.forEach((row, index) => {
    const rowNumber = index + 2;
    const textUrl = String(row[0] || '').trim();
    const richText = richTextValues[index] && richTextValues[index][0];
    const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();
    const formulaUrl = extractImageFormulaUrl_(formulas[index] && formulas[index][0]);
    const rawUrl = /^https?:\/\//i.test(textUrl) ? textUrl : linkUrl || formulaUrl;
    const normalizedUrl = normalizeImageUrl_(rawUrl);

    if (normalizedUrl && normalizedUrl === targetUrl) {
      deleted += 1;
      const targetCell = range.getCell(index + 1, 1);
      const replacement = moveNextUnassignedBackgroundToCell_(sheet, targetCell, rowNumber + 1, lastRow, targetUrl);
      if (replacement) {
        replacements.push({ row: rowNumber, backgroundUrl: replacement });
      } else {
        targetCell.clearContent();
      }
      return;
    }
  });

  return { ok: true, deleted, replacements };
}

function setFixedBackgroundMode(enabled, sheetName) {
  const sheet = getTargetSheet_(sheetName);
  const fixed = true;
  sheet.getRange('J4').setValue(fixed ? 1 : '');
  return { ok: true, fixedBackground: fixed };
}

function setBackgroundProbability(probability, sheetName) {
  const sheet = getTargetSheet_(sheetName);
  const rawValue = String(probability === null || probability === undefined ? '' : probability).trim();
  const normalized = rawValue === '0' ? 0 : 1;

  sheet.getRange('J2').setValue(normalized);
  return { ok: true, backgroundProbability: normalized };
}

function moveNextUnassignedBackgroundToCell_(sheet, targetCell, startRow, lastRow, deletedUrl) {
  for (let row = lastRow; row >= startRow; row -= 1) {
    if (rowHasStudyCard_(sheet, row)) {
      continue;
    }

    const sourceCell = sheet.getRange(row, 9);
    const normalizedUrl = getBackgroundCellUrl_(sourceCell);
    if (!normalizedUrl || normalizedUrl === deletedUrl) {
      continue;
    }

    sourceCell.moveTo(targetCell);
    return normalizedUrl;
  }

  return '';
}

function rowHasStudyCard_(sheet, row) {
  const range = sheet.getRange(row, 1, 1, 2);
  const displayValues = range.getDisplayValues()[0] || [];
  const richTextValues = range.getRichTextValues()[0] || [];
  const formulas = range.getFormulas()[0] || [];

  return displayValues.some((value) => String(value || '').trim()) ||
    richTextValues.some((richText) => String((richText && richText.getLinkUrl()) || '').trim()) ||
    formulas.some((formula) => String(formula || '').trim());
}

function getBackgroundCellUrl_(cell) {
  const textUrl = String(cell.getDisplayValue() || '').trim();
  const richText = cell.getRichTextValue();
  const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();
  const formulaUrl = extractImageFormulaUrl_(cell.getFormula());
  const rawUrl = /^https?:\/\//i.test(textUrl) ? textUrl : linkUrl || formulaUrl;
  const normalizedUrl = normalizeImageUrl_(rawUrl);
  return /^https?:\/\//i.test(normalizedUrl) ? normalizedUrl : '';
}

function recordDailyLogin(sheetName) {
  const sheet = getTargetSheet_(sheetName);
  ensureStatsLabels_(sheet);
  const timezone = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  const lastDateCell = sheet.getRange('L2');
  const lastDate = storedDateKey_(lastDateCell, timezone);
  let streak = Number(sheet.getRange('K2').getValue()) || 0;

  if (lastDate === today) {
    lastDateCell.setNumberFormat('@').setValue(today);
    const stats = updateSheetStats_(sheet);
    return { ok: true, streak, today, stats };
  }

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayKey = Utilities.formatDate(yesterday, timezone, 'yyyy-MM-dd');
  streak = lastDate === yesterdayKey ? streak + 1 : 1;

  sheet.getRange('K2').setValue(streak);
  lastDateCell.setNumberFormat('@').setValue(today);
  sheet.getRange('N2').setValue(0);
  const stats = updateSheetStats_(sheet);

  return { ok: true, streak, today, stats };
}

function recordCardPlay(sheetName) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const sheet = getTargetSheet_(sheetName);
    ensureStatsLabels_(sheet);
    const timezone = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
    const lastDateCell = sheet.getRange('L2');
    const lastDate = storedDateKey_(lastDateCell, timezone);

    if (lastDate !== today) {
      recordDailyLogin(sheetName);
    } else {
      lastDateCell.setNumberFormat('@').setValue(today);
    }

    const todayCount = (Number(sheet.getRange('N2').getValue()) || 0) + 1;
    const totalCount = (Number(sheet.getRange('N4').getValue()) || 0) + 1;
    sheet.getRange('N2').setValue(todayCount);
    sheet.getRange('N4').setValue(totalCount);
    const stats = updateSheetStats_(sheet);

    return {
      ok: true,
      todayPlayCount: todayCount,
      totalPlayCount: totalCount,
      averageRank: stats.averageRank,
    };
  } finally {
    lock.releaseLock();
  }
}

function updateStudySkip(row, sheetName) {
  const targetRow = Number(row);

  if (!targetRow || targetRow < 1) {
    return { ok: false, message: 'invalid row.' };
  }

  const sheet = getTargetSheet_(sheetName);
  const cell = sheet.getRange(targetRow, 3);
  const current = Number(cell.getValue()) || 0;
  const next = current + 1;

  cell.setValue(next);

  return { ok: true, row: targetRow, skip: next };
}

function decrementAllStudyCounts(sheetName) {
  const sheet = getTargetSheet_(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow < 1) {
    return { ok: true, updated: 0 };
  }

  const range = sheet.getRange(1, 3, lastRow, 1);
  const values = range.getValues();
  let updated = 0;

  const nextValues = values.map((row) => {
    if (row[0] === '' || row[0] === null) {
      return [row[0]];
    }

    const current = Number(row[0]);
    if (!Number.isFinite(current)) {
      return [row[0]];
    }

    const next = Math.max(0, current - 1);
    if (next !== current) updated += 1;
    return [next];
  });

  range.setValues(nextValues);

  return { ok: true, updated };
}

function updateStudyRank(row, rank, sheetName) {
  const targetRow = Number(row);
  const targetRank = normalizeRank_(rank);

  if (!targetRow || targetRow < 1) {
    return { ok: false, message: 'invalid row.' };
  }

  const sheet = getTargetSheet_(sheetName);
  const cell = sheet.getRange(targetRow, 4);
  cell.setValue(gradeFromRank_(targetRank));
  cell.setNote(rankNote_(cell.getNote(), targetRank));
  const stats = updateSheetStats_(sheet);

  return { ok: true, row: targetRow, rank: targetRank, stats };
}

function parseShortcutParams_(e) {
  const params = Object.assign({}, e && e.parameter ? e.parameter : {});
  const body = String(e && e.postData && e.postData.contents || '').trim();

  if (!body) {
    return params;
  }

  try {
    const parsed = JSON.parse(body);
    if (Array.isArray(parsed)) {
      if (!params.data1) params.data1 = JSON.stringify(parsed);
      return params;
    }
    if (parsed && typeof parsed === 'object') {
      Object.keys(parsed).forEach((key) => {
        if (params[key] === undefined) params[key] = parsed[key];
      });
      return params;
    }
  } catch (error) {
    if (!params.data1 && /^\[[\s\S]*\]$/.test(body)) {
      params.data1 = body;
    }
  }

  return params;
}

function hasShortcutData_(params) {
  return Boolean(params && (params.data1 || params.question));
}

function addShortcutItem_(params) {
  const shortcutList = parseShortcutList_(params.data1);
  if (shortcutList) {
    const sheet = getOrCreateSheet_('進行中');
    appendAfterLastValueInColumnA_(sheet, shortcutList);
    const stats = updateSheetStats_(sheet);

    return jsonResponse_({
      ok: true,
      message: 'added to 進行中.',
      row: shortcutList,
      stats,
    });
  }

  const sheet = getTargetSheet_(params.sheet || params.data0);
  const question = String(params.question || params.data1 || '').trim();
  const answer = String(params.answer || params.data2 || '').trim();
  const rank = normalizeRank_(params.rank || params.level || params.dataRank || params.data5 || 3);
  const comments = collectShortcutComments_(params);

  if (!question) {
    return jsonResponse_({
      ok: false,
      message: 'question or data1 is required.',
    });
  }

  const addedRow = appendAfterLastValueInColumnA_(sheet, [question, answer, '', gradeFromRank_(rank)].concat(comments));
  sheet.getRange(addedRow, 4).setNote(rankNote_('', rank));
  const stats = updateSheetStats_(sheet);

  return jsonResponse_({
    ok: true,
    message: 'added.',
    item: {
      question,
      answer,
      comments,
      rank,
    },
    stats,
  });
}

function parseShortcutList_(value) {
  if (Array.isArray(value)) {
    return normalizeShortcutRow_(value);
  }

  let text = String(value || '').trim();
  if (!text) {
    return null;
  }

  text = text
    .replace(/^```(?:[\w-]+)?\s*/i, '')
    .replace(/```$/i, '')
    .trim();

  const jsonList = parseJsonShortcutList_(text);
  if (jsonList) {
    return jsonList;
  }

  text = text.replace(/^\[/, '').replace(/\]$/, '').trim();

  const labeledList = parseLabeledShortcutList_(text);
  if (labeledList) {
    return labeledList;
  }

  const cleanLine = (line) => String(line || '')
    .replace(/^\s*[-*]\s*/, '')
    .replace(/^\s*\d+[.)]\s*/, '')
    .trim();
  const lines = text
    .split(/\r?\n/)
    .map(cleanLine)
    .filter(Boolean);

  if (lines.length === 7) {
    return normalizeShortcutRow_(lines);
  }

  const dotParts = text
    .split(/\s*[.]\s*/)
    .map((part) => part.trim())
    .filter(Boolean);
  if (dotParts.length >= 7) {
    return normalizeShortcutRow_(dotParts);
  }

  const tabParts = text
    .split(/\t|\s*\|\s*|,/)
    .map((part) => part.trim())
    .filter(Boolean);
  if (tabParts.length >= 7) {
    return normalizeShortcutRow_(tabParts);
  }

  return null;
}

function parseJsonShortcutList_(text) {
  const value = String(text || '').trim();
  const start = value.indexOf('[');
  const end = value.lastIndexOf(']');
  if (start < 0 || end <= start) {
    return null;
  }

  try {
    const parsed = JSON.parse(value.slice(start, end + 1));
    return Array.isArray(parsed) ? normalizeShortcutRow_(parsed) : null;
  } catch (error) {
    return null;
  }
}

function normalizeShortcutRow_(values) {
  if (!Array.isArray(values) || values.length < 7) {
    return null;
  }

  const row = values.slice(0, 6).map((item) => String(item || '').trim());
  row.push(values.slice(6).map((item) => String(item || '').trim()).filter(Boolean).join('\n'));
  return row;
}

function parseLabeledShortcutList_(text) {
  const value = String(text || '').trim();
  if (!value) {
    return null;
  }

  const question = labeledSection_(value, ['\u554f\u984c'], ['\u7b54\u3048']);
  const answerAndRanks = labeledSection_(value, ['\u7b54\u3048'], ['\u89e3\u8aac']);
  const explanation = labeledSection_(value, ['\u89e3\u8aac'], ['\u4f55\u5e74\u306e\u8a71\u304b']);
  const year = labeledSection_(value, ['\u4f55\u5e74\u306e\u8a71\u304b'], ['\u95a2\u9023\u5358\u8a9e']);
  const related = labeledSection_(value, ['\u95a2\u9023\u5358\u8a9e'], []);

  if (!question || !answerAndRanks || !explanation || !year || !related) {
    return null;
  }

  const rankMatch = answerAndRanks.match(/(?:^|\s)(\d+(?:\.\d+)?)\s*$/);
  let answer = answerAndRanks;
  let history = '0';
  let rank = '2';

  if (rankMatch) {
    answer = answerAndRanks.slice(0, rankMatch.index).trim();
    const rankParts = rankMatch[1].split('.');
    history = rankParts[0] || '0';
    rank = rankParts[1] || rankParts[0] || '2';
  }

  return [
    cleanupShortcutField_(question),
    cleanupShortcutField_(answer),
    cleanupShortcutField_(history),
    cleanupShortcutField_(rank),
    cleanupShortcutField_(explanation),
    cleanupShortcutField_(year),
    cleanupShortcutField_(related),
  ];
}

function labeledSection_(text, startLabels, endLabels) {
  const startPattern = startLabels.map(escapeRegExp_).join('|');
  const endPattern = endLabels.map(escapeRegExp_).join('|');
  const pattern = endPattern
    ? new RegExp('(?:^|\\s)(?:' + startPattern + ')\\s*[:\uff1a]\\s*([\\s\\S]*?)(?=\\s*(?:' + endPattern + ')\\s*[:\uff1a])')
    : new RegExp('(?:^|\\s)(?:' + startPattern + ')\\s*[:\uff1a]\\s*([\\s\\S]*)$');
  const match = String(text || '').match(pattern);
  return match ? match[1].trim() : '';
}

function cleanupShortcutField_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function collectShortcutComments_(params) {
  const comments = [];
  const pushComment = (value) => {
    const comment = String(value || '').trim();
    if (comment) {
      comments.push(comment);
    }
  };

  pushComment(params.comment);
  if (params.comments) {
    String(params.comments)
      .split(/\r?\n|\|\|/)
      .forEach(pushComment);
  }

  Object.keys(params)
    .filter((key) => /^comment\d+$/.test(key))
    .sort((a, b) => Number(a.replace('comment', '')) - Number(b.replace('comment', '')))
    .forEach((key) => pushComment(params[key]));

  Object.keys(params)
    .filter((key) => /^data\d+$/.test(key))
    .sort((a, b) => Number(a.replace('data', '')) - Number(b.replace('data', '')))
    .forEach((key) => {
      const index = Number(key.replace('data', ''));
      if (index === 3 || index === 4 || index >= 6) {
        pushComment(params[key]);
      }
    });

  return comments;
}

function getTargetSheet_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName || '') || spreadsheet.getActiveSheet();
}

function getOrCreateSheet_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function appendAfterLastValueInColumnA_(sheet, rowValues) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, lastRow, Math.max(sheet.getLastColumn(), 10)).getDisplayValues();
  let targetRow = 1;

  for (let index = values.length - 1; index >= 0; index -= 1) {
    if (values[index].some((value) => String(value || '').trim())) {
      targetRow = index + 2;
      break;
    }
  }

  sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
  return targetRow;
}

function ensureStatsLabels_(sheet) {
  sheet.getRange('K1').setValue('ログイン日数');
  sheet.getRange('L1').setValue('最終ログイン');
  sheet.getRange('M1').setValue('平均理解度');
  sheet.getRange('N1').setValue('今日のプレイ枚数');
  sheet.getRange('N3').setValue('総プレイ枚数');
  sheet.getRange('O1').setValue('ランク1表示率');
}

function updateSheetStats_(sheet) {
  ensureStatsLabels_(sheet);
  const lastRow = sheet.getLastRow();
  let sum = 0;
  let count = 0;

  if (lastRow >= 1) {
    const values = sheet.getRange(1, 1, lastRow, 4).getDisplayValues();
    const rankNotes = sheet.getRange(1, 4, lastRow, 1).getNotes();
    values.forEach((row, index) => {
      const hasCardContent = String(row[0] || row[1] || '').trim();
      if (hasCardContent) {
        sum += normalizeRank_(row[3], rankNotes[index][0]);
        count += 1;
      }
    });
  }

  const averageRank = count > 0 ? Math.round((sum / count) * 100000000) / 100000000 : '';
  if (averageRank === '') {
    sheet.getRange('M2').clearContent();
  } else {
    sheet.getRange('M2').setValue(averageRank);
  }

  const todayPlayCount = Number(sheet.getRange('N2').getValue()) || 0;
  const totalPlayCount = Number(sheet.getRange('N4').getValue()) || 0;

  return {
    averageRank,
    todayPlayCount,
    totalPlayCount,
  };
}

function storedDateKey_(range, timezone) {
  const value = range.getValue();
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  }

  const text = String(range.getDisplayValue() || value || '').trim();
  if (!text) {
    return '';
  }

  const normalized = text.replace(/[\/.]/g, '-');
  const match = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) {
    return text;
  }

  return [
    match[1],
    String(match[2]).padStart(2, '0'),
    String(match[3]).padStart(2, '0'),
  ].join('-');
}

// D列にはS〜Eを保存します。0.1刻みの詳細値はD列のメモへ保存します。
const RANK_GRADE_VALUES_ = { S: 1, A: 1.1, B: 1.5, C: 2, D: 2.5, E: 3 };
const RANK_NOTE_PREFIX_ = 'StudyShorts numeric rank: ';

function gradeFromRank_(value) {
  const rank = normalizeRank_(value);
  return rank <= 1 ? 'S' : rank < 1.5 ? 'A' : rank < 2 ? 'B'
    : rank < 2.5 ? 'C' : rank < 3 ? 'D' : 'E';
}

function normalizeRank_(value, note) {
  const text = String(value == null ? '' : value).trim().toUpperCase();
  const gradeValue = RANK_GRADE_VALUES_[text];
  if (gradeValue !== undefined) {
    const match = String(note || '').match(/(?:^|\n)StudyShorts numeric rank: (\d(?:\.\d)?)(?:\n|$)/);
    if (match) {
      const detailed = Math.min(3, Math.max(1, Number(match[1])));
      if (gradeFromNumericRank_(detailed) === text) return detailed;
    }
    return gradeValue;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric >= 1
    ? Math.min(3, Math.max(1, Math.round(numeric * 10) / 10)) : 3;
}

function gradeFromNumericRank_(rank) {
  return rank <= 1 ? 'S' : rank < 1.5 ? 'A' : rank < 2 ? 'B'
    : rank < 2.5 ? 'C' : rank < 3 ? 'D' : 'E';
}

function rankNote_(oldNote, rank) {
  const withoutRank = String(oldNote || '')
    .replace(/(?:^|\n)StudyShorts numeric rank: \d(?:\.\d)?(?=\n|$)/g, '')
    .replace(/^\n|\n$/g, '');
  return [withoutRank, RANK_NOTE_PREFIX_ + normalizeRank_(rank)].filter(Boolean).join('\n');
}

function ensureLetterRanks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  const values = sheet.getRange(1, 1, lastRow, 4).getDisplayValues();
  const rankRange = sheet.getRange(1, 4, lastRow, 1);
  const notes = rankRange.getNotes();
  values.forEach((row, index) => {
    if (!String(row[0] || row[1] || '').trim()) return;
    const existing = String(row[3] || '').trim().toUpperCase();
    if (Object.prototype.hasOwnProperty.call(RANK_GRADE_VALUES_, existing)) return;
    const numeric = normalizeRank_(row[3]);
    const cell = sheet.getRange(index + 1, 4);
    cell.setValue(gradeFromRank_(numeric));
    cell.setNote(rankNote_(notes[index][0], numeric));
  });
}

function syncDriveFolderImagesToSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const folderIds = getDriveFolderIds_(sheet, lastRow);
  const errors = [];
  const currentFileIds = {};
  let imageCount = 0;
  let added = 0;
  let removed = 0;

  if (folderIds.length === 0) {
    return { folderCount: 0, imageCount: 0, added: 0, removed: 0, errors };
  }

  const existing = getExistingImageRows_(sheet);

  folderIds.forEach((folderId) => {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        if (!isDriveStudyFile_(file)) {
          continue;
        }

        imageCount += 1;
        const fileId = file.getId();
        currentFileIds[fileId] = true;
        const name = file.getName() || 'Drive image';
        const imageUrl = driveFileImageUrl_(fileId);
        const existingFileRow = existing.byFileId[fileId] || 0;
        const existingNameRow = existing.byName[name] || 0;
        const existingRow = existingFileRow ||
          (existingNameRow && existing.driveRows[existingNameRow] ? existingNameRow : 0);

        if (existingRow) {
          if (existing.driveRows[existingRow]) {
            setLinkedCell_(sheet.getRange(existingRow, 1), name, imageUrl);
            setLinkedCell_(sheet.getRange(existingRow, 2), name, imageUrl);
          }
          existing.byFileId[fileId] = existingRow;
          existing.byName[name] = existingRow;
          continue;
        }

        const row = sheet.getLastRow() + 1;
        sheet.getRange(row, 1, 1, 4).setValues([[name, name, '', 3]]);
        setLinkedCell_(sheet.getRange(row, 1), name, imageUrl);
        setLinkedCell_(sheet.getRange(row, 2), name, imageUrl);
        existing.byFileId[fileId] = row;
        existing.byName[name] = row;
        added += 1;
      }
    } catch (error) {
      errors.push({
        folderId,
        message: String(error && error.message || error),
      });
    }
  });

  if (errors.length === 0) {
    removed = removeStaleDriveRows_(sheet, existing, currentFileIds);
  }

  return {
    folderCount: folderIds.length,
    imageCount,
    added,
    removed,
    errors,
  };
}

function getExistingImageRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const result = { byFileId: {}, byName: {}, driveRows: {}, manualFileIds: {} };

  if (lastRow < 1) {
    return result;
  }

  const range = sheet.getRange(1, 1, lastRow, 2);
  const displayValues = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const formulas = range.getFormulas();

  displayValues.forEach((row, index) => {
    const rowNumber = index + 1;
    const name = String(row[0] || row[1] || '').trim();
    if (name && !result.byName[name]) {
      result.byName[name] = rowNumber;
    }

    const rowFileIds = [];
    for (let column = 0; column < 2; column += 1) {
      const richText = richTextValues[index] && richTextValues[index][column];
      const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();
      const textUrl = String(row[column] || '').trim();
      const formulaUrl = extractImageFormulaUrl_(formulas[index] && formulas[index][column]);
      const fileId = extractDriveFileId_(textUrl) ||
        extractDriveFileId_(linkUrl) ||
        extractDriveFileId_(formulaUrl);
      if (fileId && !result.byFileId[fileId]) {
        result.byFileId[fileId] = rowNumber;
      }
      if (fileId) {
        rowFileIds.push(fileId);
      }
    }

    const isSyncedDriveRow = rowFileIds.length > 0 && rowFileIds.every((fileId) => fileId === rowFileIds[0]);
    if (isSyncedDriveRow) {
      const question = String(row[0] || '').trim();
      const answer = String(row[1] || '').trim();
      if (question && question === answer) {
        result.driveRows[rowNumber] = rowFileIds[0];
      }
    }

    if (!result.driveRows[rowNumber]) {
      rowFileIds.forEach((fileId) => {
        result.manualFileIds[fileId] = rowNumber;
      });
    }
  });

  return result;
}

function removeStaleDriveRows_(sheet, existing, currentFileIds) {
  const rowsToDelete = Object.keys(existing.driveRows)
    .map((row) => Number(row))
    .filter((row) => {
      const fileId = existing.driveRows[row];
      return row > 0 && (!currentFileIds[fileId] || existing.manualFileIds[fileId]);
    })
    .sort((a, b) => b - a);

  rowsToDelete.forEach((row) => sheet.deleteRow(row));
  return rowsToDelete.length;
}

function setLinkedCell_(range, text, url) {
  range.setRichTextValue(
    SpreadsheetApp.newRichTextValue()
      .setText(String(text || 'Drive image'))
      .setLinkUrl(url)
      .build()
  );
}

function getQuestionImageUrls_(sheet, lastRow) {
  return getColumnImageUrls_(sheet, lastRow, 1);
}

function getAnswerImageUrls_(sheet, lastRow) {
  return getColumnImageUrls_(sheet, lastRow, 2);
}

function getQuestionDriveUrls_(sheet, lastRow) {
  return getColumnDriveUrls_(sheet, lastRow, 1);
}

function getAnswerDriveUrls_(sheet, lastRow) {
  return getColumnDriveUrls_(sheet, lastRow, 2);
}

function getColumnImageUrls_(sheet, lastRow, column) {
  if (lastRow < 1) {
    return [];
  }

  const range = sheet.getRange(1, column, lastRow, 1);
  const displayValues = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const formulas = range.getFormulas();

  return displayValues.map((row, index) => {
    const textUrl = String(row[0] || '').trim();
    const richText = richTextValues[index] && richTextValues[index][0];
    const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();
    const formulaUrl = extractImageFormulaUrl_(formulas[index] && formulas[index][0]);
    const fileId = extractDriveFileId_(textUrl) ||
      extractDriveFileId_(linkUrl) ||
      extractDriveFileId_(formulaUrl);
    if (fileId) {
      return driveFileImageUrl_(fileId);
    }

    const url = /^https?:\/\//i.test(textUrl) ? textUrl : linkUrl || formulaUrl;
    return isImageLikeUrl_(url) ? normalizeImageUrl_(url) : '';
  });
}

function getColumnDriveUrls_(sheet, lastRow, column) {
  if (lastRow < 1) {
    return [];
  }

  const range = sheet.getRange(1, column, lastRow, 1);
  const displayValues = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const formulas = range.getFormulas();

  return displayValues.map((row, index) => {
    const textUrl = String(row[0] || '').trim();
    const richText = richTextValues[index] && richTextValues[index][0];
    const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();
    const formulaUrl = extractImageFormulaUrl_(formulas[index] && formulas[index][0]);
    const fileId = extractDriveFileId_(textUrl) ||
      extractDriveFileId_(linkUrl) ||
      extractDriveFileId_(formulaUrl);
    return fileId ? driveFileOpenUrl_(fileId) : '';
  });
}

function getDriveFolderImageItems_(sheet, lastRow) {
  const folderIds = getDriveFolderIds_(sheet, lastRow);
  const items = [];
  const errors = [];

  folderIds.forEach((folderId) => {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        if (!isDriveStudyFile_(file)) {
          continue;
        }

        const imageUrl = driveFileImageUrl_(file.getId());
        const driveUrl = driveFileOpenUrl_(file.getId());
        items.push({
          row: 0,
          id: `drive-${file.getId()}`,
          question: file.getName() || 'Drive image',
          questionImageUrl: imageUrl,
          questionDriveUrl: driveUrl,
          answer: imageUrl,
          answerImageUrl: imageUrl,
          answerDriveUrl: driveUrl,
          comments: [],
          skip: 0,
          rank: 3,
          driveFileId: file.getId(),
        });
      }
    } catch (error) {
      errors.push({
        folderId,
        message: String(error && error.message || error),
      });
    }
  });

  return {
    items,
    folderCount: folderIds.length,
    errors,
  };
}

function isDriveStudyFile_(file) {
  const mimeType = String(file.getMimeType() || '').toLowerCase();
  const name = String(file.getName() || '').toLowerCase();
  return mimeType.startsWith('image/') ||
    mimeType === 'application/pdf' ||
    /\.(png|jpe?g|gif|webp|bmp|svg|heic|heif|pdf)$/i.test(name);
}

function getDriveFolderIds_(sheet, lastRow) {
  if (lastRow < 2) {
    return [];
  }

  const range = sheet.getRange(2, 8, lastRow - 1, 1);
  const displayValues = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const ids = [];

  displayValues.forEach((row, index) => {
    const text = String(row[0] || '').trim();
    const richText = richTextValues[index] && richTextValues[index][0];
    const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();
    const id = extractDriveFolderId_(text) || extractDriveFolderId_(linkUrl);
    if (id && !ids.includes(id)) {
      ids.push(id);
    }
  });

  return ids;
}

function extractDriveFolderId_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const folderMatch = text.match(/drive\.google\.com\/drive\/folders\/([^/?#]+)/);
  if (folderMatch) return folderMatch[1];

  const idMatch = text.match(/[?&]id=([^&]+)/);
  if (idMatch) return idMatch[1];

  return /^[A-Za-z0-9_-]{20,}$/.test(text) ? text : '';
}

function driveFileImageUrl_(fileId) {
  return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1200';
}

function driveFileOpenUrl_(fileId) {
  return 'https://drive.google.com/file/d/' + encodeURIComponent(fileId) + '/view';
}

function extractDriveFileId_(url) {
  const value = String(url || '').trim();
  if (!/(drive|docs)\.google\.com|drive\.usercontent\.google\.com/i.test(value)) {
    return '';
  }

  const driveMatch = value.match(/drive\.google\.com\/file\/d\/([^/?#]+)/i) ||
    value.match(/docs\.google\.com\/(?:document|presentation|spreadsheets|drawings)\/d\/([^/?#]+)/i) ||
    value.match(/[?&]id=([^&#]+)/i);
  return driveMatch ? decodeURIComponent(driveMatch[1]) : '';
}

function extractImageFormulaUrl_(formula) {
  const match = String(formula || '').match(/=IMAGE\(\s*"([^"]+)"/i);
  return match ? match[1] : '';
}

function getBackgroundUrls_(sheet, lastRow) {
  if (lastRow < 2) {
    return [];
  }

  const range = sheet.getRange(2, 9, lastRow - 1, 1);
  const displayValues = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();

  return displayValues
    .map((row, index) => {
      const textUrl = String(row[0] || '').trim();
      const richText = richTextValues[index] && richTextValues[index][0];
      const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();

      const url = /^https?:\/\//i.test(textUrl) ? textUrl : linkUrl;
      return normalizeImageUrl_(url);
    })
    .filter((value) => /^https?:\/\//i.test(value));
}

function getRowBackgroundUrls_(sheet, lastRow) {
  if (lastRow < 1) {
    return [];
  }

  const range = sheet.getRange(1, 9, lastRow, 1);
  const displayValues = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const formulas = range.getFormulas();

  return displayValues.map((row, index) => {
    const textUrl = String(row[0] || '').trim();
    const richText = richTextValues[index] && richTextValues[index][0];
    const linkUrl = String((richText && richText.getLinkUrl()) || '').trim();
    const formulaUrl = extractImageFormulaUrl_(formulas[index] && formulas[index][0]);
    const url = /^https?:\/\//i.test(textUrl) ? textUrl : linkUrl || formulaUrl;
    const normalizedUrl = normalizeImageUrl_(url);
    return /^https?:\/\//i.test(normalizedUrl) ? normalizedUrl : '';
  });
}

function isImageLikeUrl_(url) {
  return /^https?:\/\//i.test(String(url || ''));
}

function normalizeImageUrl_(url) {
  const value = String(url || '').trim();
  const driveFileId = extractDriveFileId_(value);
  if (driveFileId) {
    return driveFileImageUrl_(driveFileId);
  }

  return value;
}

function normalizeProbability_(value) {
  const probability = Number(value);
  if (!Number.isFinite(probability)) {
    return null;
  }
  return Math.min(1, Math.max(0, probability));
}

function parseBooleanSetting_(value) {
  if (value === true) return true;
  if (value === false || value === null || value === undefined) return false;
  const text = String(value).trim().toLowerCase();
  if (!text) return false;
  return ['1', 'true', 'on', 'yes', 'fixed', '固定'].includes(text);
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
