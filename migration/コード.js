// ===================================================
//  工事台帳Excel → GASアプリ 移行ツール
//  株式会社ネクスト・ワン
//
//  【使い方】
//  1. FOLDER_ID にExcel台帳を入れたDriveフォルダのIDを設定
//  2. TARGET_SS_ID にGASアプリのスプレッドシートIDを設定
//  3. migrateAll() を実行
// ===================================================

// ★ 設定 ★
const FOLDER_ID = '';      // Excel台帳が入っているGoogleドライブのフォルダID
const TARGET_SS_ID = '1U27WZTOkEAfNiVmQd02jGswnlZHGxLZn8aFWpPTuMug';  // 本番環境

// ===================================================
//  Webアプリ エントリーポイント
// ===================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('工事台帳 取り込みツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ブラウザからアップロードされたExcelファイルを処理
 * @param {string} base64Data - ファイルのBase64エンコード
 * @param {string} filename - ファイル名
 * @return {Object} {expenses, billings, projectName} or {error}
 */
function processUploadedFile(base64Data, filename) {
  const idMatch = filename.match(/【(\d+)】/);
  if (!idMatch) return { error: '台帳番号が見つかりません' };
  const projectId = idMatch[1];

  let tempFileId = null;
  try {
    // Base64 → Blob → Driveに一時保存 → Sheets変換
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename);
    const tempXlsx = DriveApp.createFile(blob);
    const converted = Drive.Files.copy(
      { title: 'temp_upload_' + projectId, mimeType: 'application/vnd.google-apps.spreadsheet' },
      tempXlsx.getId()
    );
    tempFileId = converted.id;
    tempXlsx.setTrashed(true); // xlsxの一時ファイルも削除

    const ss = SpreadsheetApp.openById(tempFileId);
    const ws = findLedgerSheet_(ss);
    if (!ws) return { error: '工事台帳シートが見つかりません' };

    const data = parseLedger_(ws, projectId, filename);

    // 書き込み
    writeToTarget_([data.project], data.changes, data.expenses, data.billings);

    return {
      projectName: data.project['案件名'],
      expenses: data.expenses.length,
      billings: data.billings.length,
      changes: data.changes.length,
    };

  } catch (e) {
    return { error: e.message };
  } finally {
    if (tempFileId) {
      try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (_) {}
    }
  }
}

// ===================================================
//  メイン処理（フォルダ一括版）
// ===================================================

/**
 * フォルダ内の全Excel台帳を読み込み、GASアプリのスプレッドシートに一括書き込み
 */
function migrateAll() {
  if (!FOLDER_ID) {
    Logger.log('❌ FOLDER_ID を設定してください');
    return;
  }

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  const allProjects = [];
  const allChanges = [];
  const allExpenses = [];
  const allBillings = [];
  const errors = [];

  while (files.hasNext()) {
    const file = files.next();
    const filename = file.getName();

    // ファイル名から台帳番号を抽出
    const idMatch = filename.match(/【(\d+)】/);
    if (!idMatch) {
      Logger.log('⏭ スキップ（台帳番号なし）: ' + filename);
      continue;
    }
    const projectId = idMatch[1];
    Logger.log('📂 処理中: [' + projectId + '] ' + filename.substring(0, 50) + '...');

    let tempFileId = null;
    try {
      // xlsxをGoogle Sheetsに変換（一時ファイル）
      const tempFile = Drive.Files.copy(
        { title: 'temp_migration_' + projectId, mimeType: 'application/vnd.google-apps.spreadsheet' },
        file.getId()
      );
      tempFileId = tempFile.id;
      const ss = SpreadsheetApp.openById(tempFileId);
      const ws = findLedgerSheet_(ss);

      if (!ws) {
        errors.push({ file: filename, error: '工事台帳シートが見つかりません' });
        continue;
      }

      // パース
      const data = parseLedger_(ws, projectId, filename);
      allProjects.push(data.project);
      allChanges.push.apply(allChanges, data.changes);
      allExpenses.push.apply(allExpenses, data.expenses);
      allBillings.push.apply(allBillings, data.billings);

      Logger.log('  ✅ 経費: ' + data.expenses.length + '件, 請求: ' + data.billings.length + '件');

    } catch (e) {
      errors.push({ file: filename, error: e.message });
      Logger.log('  ❌ エラー: ' + e.message);
    } finally {
      // 一時ファイル削除
      if (tempFileId) {
        try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (_) {}
      }
    }
  }

  // スプレッドシートに書き込み
  Logger.log('');
  Logger.log('=== 書き込み開始 ===');
  writeToTarget_(allProjects, allChanges, allExpenses, allBillings);

  // サマリー
  Logger.log('');
  Logger.log('=== 完了 ===');
  Logger.log('  案件: ' + allProjects.length + '件');
  Logger.log('  契約変更: ' + allChanges.length + '件');
  Logger.log('  経費: ' + allExpenses.length + '件');
  Logger.log('  請求入金: ' + allBillings.length + '件');
  if (errors.length > 0) {
    Logger.log('  エラー: ' + errors.length + '件');
    errors.forEach(function(e) { Logger.log('    - ' + e.file.substring(0, 40) + ': ' + e.error); });
  }
}

// ===================================================
//  シート検索
// ===================================================

function findLedgerSheet_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf('工事台帳') >= 0) return sheets[i];
  }
  return sheets[0]; // フォールバック
}

// ===================================================
//  台帳パーサー
// ===================================================

function parseLedger_(ws, projectId, filename) {
  const data = ws.getDataRange().getValues();

  return {
    project: parseHeader_(data, projectId, filename),
    changes: parseContractChanges_(data, projectId),
    expenses: parseExpenses_(data, projectId),
    billings: parseBillings_(data, projectId),
  };
}

// --- ヘルパー ---

function cellVal_(data, row, col) {
  if (row < 0 || row >= data.length) return '';
  if (col < 0 || col >= data[row].length) return '';
  return data[row][col];
}

function strVal_(data, row, col) {
  const v = cellVal_(data, row, col);
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return formatDate_(v);
  return String(v).trim();
}

function numVal_(data, row, col) {
  const v = cellVal_(data, row, col);
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : Math.round(n);
}

function normalize_(s) {
  return String(s || '').replace(/[\s\u3000]+/g, '');
}

function formatDate_(d) {
  if (!d || !(d instanceof Date)) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function findRow_(data, label, col, startRow, endRow) {
  const norm = normalize_(label);
  for (let r = startRow; r <= Math.min(endRow, data.length - 1); r++) {
    if (normalize_(strVal_(data, r, col)).indexOf(norm) >= 0) return r;
  }
  return -1;
}

function statusFromFilename_(filename) {
  if (filename.indexOf('入金完了') >= 0) return '["完工"]';
  if (filename.indexOf('完工') >= 0) return '["完工"]';
  if (filename.indexOf('請求待ち') >= 0) return '["保留金請求待ち"]';
  return '["進行中"]';
}

// ===================================================
//  ヘッダーパース → projects
// ===================================================

function parseHeader_(data, projectId, filename) {
  // 受注金額行を探す
  const juchuRow = findRow_(data, '受注金額', 0, 7, 40);
  let contractBase = 0, contractTax = 0, contractTotal = 0;
  if (juchuRow >= 0) {
    contractBase  = numVal_(data, juchuRow, 6);
    contractTax   = numVal_(data, juchuRow + 1, 6);
    contractTotal = numVal_(data, juchuRow + 2, 6);
  }
  if (contractBase && !contractTax)   contractTax = Math.round(contractBase * 0.1);
  if (contractBase && !contractTotal) contractTotal = contractBase + contractTax;

  // 目標金額行を探す
  const mokuhyoRow = findRow_(data, '目標金額', 0, 7, 40);
  let targetBase = 0, targetTax = 0, targetTotal = 0, targetRate = 78;
  if (mokuhyoRow >= 0) {
    targetBase  = numVal_(data, mokuhyoRow, 6);
    targetTax   = numVal_(data, mokuhyoRow + 1, 6);
    targetTotal = numVal_(data, mokuhyoRow + 2, 6);
    const rateMatch = normalize_(strVal_(data, mokuhyoRow, 0)).match(/(\d+)[%％]/);
    if (rateMatch) targetRate = parseInt(rateMatch[1]);
  }

  // 工期
  const periodStart = parsePeriodDate_(strVal_(data, 6, 1), strVal_(data, 6, 2));
  const periodEnd   = parsePeriodDate_(strVal_(data, 6, 5), strVal_(data, 6, 6));

  return {
    '台帳番号':       projectId,
    '客先名':         strVal_(data, 2, 1),
    '営業担当':       strVal_(data, 3, 2),
    '案件名':         strVal_(data, 4, 1),
    '住所':           strVal_(data, 5, 1),
    '工期開始':       periodStart,
    '工期終了':       periodEnd,
    '契約金額_本体':  contractBase,
    '契約金額_消費税': contractTax,
    '契約金額_税込':  contractTotal,
    '目標粗利率':     targetRate,
    '目標金額_本体':  targetBase,
    '目標金額_消費税': targetTax,
    '目標金額_税込':  targetTotal,
    'ステータス':     statusFromFilename_(filename),
    'アラートメッセージ': '',
    '作成日時':       new Date().toISOString(),
    '更新日時':       new Date().toISOString(),
  };
}

function parsePeriodDate_(reiwaPart, datePart) {
  const rs = String(reiwaPart || '');
  const ds = String(datePart || '');

  // Date型ならそのまま変換
  if (datePart instanceof Date) {
    // 和暦年からwesternYearを得る
    const rm = rs.match(/[RrＲ]?\s*(\d+)/);
    if (rm) {
      const wy = parseInt(rm[1]) + 2018;
      const m = datePart.getMonth() + 1;
      const d = datePart.getDate();
      return wy + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    }
    return formatDate_(datePart);
  }

  const rm = rs.match(/[RrＲ]?\s*(\d+)/);
  if (!rm) return '';
  const wy = parseInt(rm[1]) + 2018;

  const dm = ds.replace(/^[''\/]+/, '').match(/(\d+)\/(\d+)/);
  if (dm) return wy + '-' + String(parseInt(dm[1])).padStart(2, '0') + '-' + String(parseInt(dm[2])).padStart(2, '0');

  const dm2 = ds.match(/(\d+)/);
  if (dm2) return wy + '-' + String(parseInt(dm2[1])).padStart(2, '0') + '-01';

  return '';
}

// ===================================================
//  契約変更パース → contract_changes
// ===================================================

function parseContractChanges_(data, projectId) {
  const changes = [];
  for (let r = 9; r < Math.min(30, data.length); r++) {
    const label = strVal_(data, r, 0);
    if (!label) continue;
    if (label === '計') break;

    const amountBase  = numVal_(data, r, 2);
    const amountTotal = numVal_(data, r, 4);
    if (amountBase === 0 && amountTotal === 0) continue;

    changes.push({
      'ID':           Utilities.getUuid(),
      '台帳番号':     projectId,
      '変更種別':     label,
      '変更日':       strVal_(data, r, 1),
      '変更金額_本体': amountBase,
      '変更金額_税込': amountTotal,
      '備考':         strVal_(data, r, 6),
      '作成日時':     new Date().toISOString(),
    });
  }
  return changes;
}

// ===================================================
//  経費パース → expenses
// ===================================================

function parseExpenses_(data, projectId) {
  const expenses = [];

  // 「業者名」ヘッダー行を探す
  let startRow = -1;
  for (let r = 27; r < Math.min(55, data.length); r++) {
    const v = normalize_(strVal_(data, r, 0));
    if (v.indexOf('業') >= 0 && v.indexOf('名') >= 0) {
      startRow = r + 1;
      break;
    }
  }
  if (startRow < 0) return expenses;

  // 「請求日」行を探す → 経費セクション終端
  let endRow = data.length - 1;
  for (let r = startRow; r < data.length; r++) {
    if (strVal_(data, r, 0) === '請求日') { endRow = r; break; }
  }

  let currentYear = null;
  let currentMonth = null;
  let prevVendor = '';

  for (let r = startRow; r < endRow; r++) {
    const aVal = strVal_(data, r, 0);

    // 年度行: "R7年度", "R 7年度"
    const yearMatch = aVal.match(/[RrＲ]\s*(\d+)\s*年度/);
    if (yearMatch) {
      currentYear = parseInt(yearMatch[1]) + 2018;
      continue;
    }

    // 月ヘッダー: ＜3月分＞, 《4月分》, ＜１月分＞
    const rowText = aVal || strVal_(data, r, 1);
    if (rowText) {
      const monthMatch = String(rowText).match(/[＜《<]?\s*([0-9０-９]+)\s*月分/);
      if (monthMatch) {
        const mStr = monthMatch[1].replace(/[０-９]/g, function(c) { return String(c.charCodeAt(0) - 65296); });
        currentMonth = parseInt(mStr);
        if (currentYear === null) currentYear = 2025;
        continue;
      }
    }

    // 小計・計はスキップ
    if (aVal === '小計' || aVal === '計') continue;

    // 金額チェック
    const amount = numVal_(data, r, 3);
    if (amount === 0 || currentMonth === null) continue;

    // 業者名
    let vendor = aVal;
    if (vendor === '〃') {
      const noteG = strVal_(data, r, 6);
      const noteI = strVal_(data, r, 8);
      const noteCombined = [noteG, noteI].filter(Boolean).join(' ');
      if (noteCombined.indexOf('値引') >= 0) {
        vendor = prevVendor + '(値引)';
      } else if (noteCombined) {
        vendor = prevVendor + '(' + noteCombined.substring(0, 10) + ')';
      } else {
        vendor = prevVendor + '(2)';
      }
    } else if (vendor) {
      prevVendor = vendor;
    } else {
      continue;
    }

    const year = currentYear || 2025;
    const monthDate = year + '-' + String(currentMonth).padStart(2, '0') + '-01';

    const offset = strVal_(data, r, 5);
    const noteG = strVal_(data, r, 6);
    const noteI = strVal_(data, r, 8);
    const note = [noteG, noteI].filter(Boolean).join(' ');

    expenses.push({
      'ID':       Utilities.getUuid(),
      '台帳番号': projectId,
      '月':       monthDate,
      '仕入先':   vendor,
      '金額':     amount,
      '相殺':     (offset && offset !== 'なし') ? offset : '',
      '備考':     note,
      '作成日時': new Date().toISOString(),
    });
  }

  return expenses;
}

// ===================================================
//  請求入金パース → billings
// ===================================================

function parseBillings_(data, projectId) {
  const billings = [];

  let headerRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (strVal_(data, r, 0) === '請求日') { headerRow = r; break; }
  }
  if (headerRow < 0) return billings;

  let currentDate = '';

  for (let r = headerRow + 1; r < data.length; r++) {
    const aVal = strVal_(data, r, 0);
    const bVal = strVal_(data, r, 1);
    const cVal = numVal_(data, r, 2);
    const eVal = strVal_(data, r, 4);
    const fVal = strVal_(data, r, 5);
    const gRaw = cellVal_(data, r, 6);
    const jVal = strVal_(data, r, 9);

    if (aVal === '計') break;
    if (bVal === '契約金額') break;

    // 相殺行スキップ
    if (fVal && fVal.indexOf('相殺') >= 0 && !bVal) continue;

    if (aVal) currentDate = aVal;

    const category = bVal;
    if (!category) continue;

    let confirmed = 0;
    if (gRaw !== null && gRaw !== undefined && gRaw !== '') {
      const n = Number(gRaw);
      confirmed = isNaN(n) ? 0 : Math.round(n);
    }

    if (cVal === 0 && confirmed === 0) continue;

    billings.push({
      'ID':         Utilities.getUuid(),
      '台帳番号':   projectId,
      '請求日':     currentDate,
      '種別':       category,
      '請求金額':   cVal,
      '入金予定日': eVal,
      '入金確認額': confirmed,
      '確認者':     jVal,
      '作成日時':   new Date().toISOString(),
    });
  }

  return billings;
}

// ===================================================
//  スプレッドシートへの書き込み
// ===================================================

function writeToTarget_(projects, changes, expenses, billings) {
  const ss = SpreadsheetApp.openById(TARGET_SS_ID);

  // 今回取り込む台帳番号一覧
  const projectIds = projects.map(function(p) { return String(p['台帳番号']); });

  const configs = [
    {
      sheetName: 'projects',
      headers: ['台帳番号','客先名','営業担当','案件名','住所','工期開始','工期終了',
                '契約金額_本体','契約金額_消費税','契約金額_税込','目標粗利率',
                '目標金額_本体','目標金額_消費税','目標金額_税込',
                'ステータス','アラートメッセージ','作成日時','更新日時'],
      data: projects,
      idCol: 0,  // 台帳番号の列（0始まり）
    },
    {
      sheetName: 'contract_changes',
      headers: ['ID','台帳番号','変更種別','変更日','変更金額_本体','変更金額_税込','備考','作成日時'],
      data: changes,
      idCol: 1,  // 台帳番号の列
    },
    {
      sheetName: 'expenses',
      headers: ['ID','台帳番号','月','仕入先','金額','相殺','備考','作成日時'],
      data: expenses,
      idCol: 1,
    },
    {
      sheetName: 'billings',
      headers: ['ID','台帳番号','請求日','種別','請求金額','入金予定日','入金確認額','確認者','作成日時'],
      data: billings,
      idCol: 1,
    },
  ];

  configs.forEach(function(config) {
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) {
      Logger.log('⚠ シートが見つかりません: ' + config.sheetName);
      return;
    }

    // 同じ台帳番号の既存行を削除（下から順に）
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existingData = sheet.getRange(2, config.idCol + 1, lastRow - 1, 1).getValues();
      for (let i = existingData.length - 1; i >= 0; i--) {
        if (projectIds.indexOf(String(existingData[i][0])) >= 0) {
          sheet.deleteRow(i + 2);
        }
      }
    }

    if (config.data.length === 0) return;

    const rows = config.data.map(function(record) {
      return config.headers.map(function(h) {
        return record[h] !== undefined ? record[h] : '';
      });
    });

    // 新データを追記
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, config.headers.length).setValues(rows);
    Logger.log('  ' + config.sheetName + ': ' + rows.length + '行 書き込み完了');
  });
}

// ===================================================
//  ユーティリティ: 単一ファイルテスト用
// ===================================================

/**
 * 1つのファイルだけテスト実行（ファイルIDを指定）
 */
function migrateOneFile(fileId) {
  const file = DriveApp.getFileById(fileId);
  const filename = file.getName();
  const idMatch = filename.match(/【(\d+)】/);
  if (!idMatch) { Logger.log('台帳番号が見つかりません'); return; }

  const projectId = idMatch[1];
  Logger.log('処理中: [' + projectId + '] ' + filename);

  const tempFile = Drive.Files.copy(
    { title: 'temp_test_' + projectId, mimeType: 'application/vnd.google-apps.spreadsheet' },
    file.getId()
  );

  try {
    const ss = SpreadsheetApp.openById(tempFile.id);
    const ws = findLedgerSheet_(ss);
    const result = parseLedger_(ws, projectId, filename);

    Logger.log('案件名: ' + result.project['案件名']);
    Logger.log('契約金額: ' + result.project['契約金額_税込']);
    Logger.log('経費: ' + result.expenses.length + '件');
    Logger.log('請求入金: ' + result.billings.length + '件');
    Logger.log('契約変更: ' + result.changes.length + '件');

    // 経費合計
    const total = result.expenses.reduce(function(s, e) { return s + e['金額']; }, 0);
    Logger.log('経費合計: ' + total);
  } finally {
    DriveApp.getFileById(tempFile.id).setTrashed(true);
  }
}
