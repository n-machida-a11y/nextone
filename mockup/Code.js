// ===================================================
//  請求工事台帳振分業務 - GAS サーバーサイド
//  株式会社ネクスト・ワン
// ===================================================

// ★ここにスプレッドシートのIDを設定してください
const SPREADSHEET_ID = '1dJ1ZTBWyk4uQVUIXbAC2SslVdSzyVbk29tRQsjYLQTA';

const SHEETS = {
  PROJECTS:         'projects',
  CONTRACT_CHANGES: 'contract_changes',
  EXPENSES:         'expenses',
  BILLINGS:         'billings',
  VENDORS:          'vendors',
  INVOICES:         'invoices'
};

// ===================================================
//  エントリーポイント
// ===================================================

function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('工事台帳管理 | 株式会社ネクスト・ワン')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** GASテンプレートのinclude用ヘルパー */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===================================================
//  共通ヘルパー
// ===================================================

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

function generateId() {
  return Utilities.getUuid();
}

function now() {
  return new Date().toISOString();
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    obj[h] = row[i];
  });
  return obj;
}

/** シート全データを [{ヘッダー名: 値, ...}, ...] 形式で返す
 *  google.script.run はDateオブジェクトを戻り値に含められないため、
 *  Date型セルをISO文字列（YYYY-MM-DD）へ変換して返す。
 */
function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName + '。setupSheets() を先に実行してください。');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tz = Session.getScriptTimeZone();
  return data.slice(1).map(row => {
    const sanitized = row.map(cell =>
      cell instanceof Date
        ? Utilities.formatDate(cell, tz, 'yyyy-MM-dd')
        : cell
    );
    return rowToObject(headers, sanitized);
  });
}

/** IDがUUID列(A列)と一致する行を削除 */
function deleteById(sheetName, id) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { error: 'レコードが見つかりません: ' + id };
  } catch (e) {
    return { error: e.message };
  }
}

/** 汎用upsert: IDなし→append / IDあり→既存行更新 */
function saveRecord(sheetName, recordData) {
  try {
    const sheet = getSheet(sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const isNew = !recordData['ID'] || recordData['ID'] === '';

    if (isNew) {
      recordData['ID'] = generateId();
      recordData['作成日時'] = now();
      const rowValues = headerRow.map(h => (recordData[h] !== undefined ? recordData[h] : ''));
      sheet.appendRow(rowValues);
    } else {
      const data = sheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(recordData['ID'])) {
          const rowValues = headerRow.map(h => (recordData[h] !== undefined ? recordData[h] : ''));
          sheet.getRange(i + 1, 1, 1, headerRow.length).setValues([rowValues]);
          found = true;
          break;
        }
      }
      if (!found) return { error: 'レコードが見つかりません: ' + recordData['ID'] };
    }
    return { success: true, id: recordData['ID'] };
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  ステータス自動算出
// ===================================================

function calcStatus(projectId, contractTotal) {
  const statusList = [];
  const margin = calcGrossMarginValue(projectId, contractTotal);
  if (margin !== null && margin < 0) statusList.push('赤字注意');

  const billings = getBillings(projectId);
  const hasPendingRetention = billings.some(b =>
    b['種別'] === '保留金' &&
    (b['入金確認額'] === '' || b['入金確認額'] === null || b['入金確認額'] === 0)
  );
  if (hasPendingRetention) statusList.push('保留金請求待ち');

  if (billings.length > 0 && !hasPendingRetention &&
      billings.every(b => b['入金確認額'] !== '' && b['入金確認額'] !== null && Number(b['入金確認額']) > 0)) {
    statusList.push('完工');
  } else if (!hasPendingRetention && !statusList.includes('赤字注意')) {
    statusList.push('進行中');
  }

  if (statusList.length === 0) statusList.push('進行中');
  return statusList;
}

// ===================================================
//  粗利率計算
// ===================================================

function calculateGrossMargin(projectId) {
  try {
    const projects = getSheetData(SHEETS.PROJECTS);
    const project = projects.find(p => String(p['台帳番号']) === String(projectId));
    if (!project) return null;
    const contractTotal = Number(project['契約金額_税込']) || 0;
    return calcGrossMarginValue(projectId, contractTotal);
  } catch (e) {
    return null;
  }
}

function calcGrossMarginValue(projectId, contractTotal) {
  if (!contractTotal || contractTotal === 0) return null;
  const expenses = getExpenses(projectId);
  if (expenses.error) return null;
  return _calcMarginFromExpenses(projectId, contractTotal, expenses);
}

/** 読み込み済みの経費配列から粗利率を計算（シート再読み込みなし） */
function _calcMarginFromExpenses(projectId, contractTotal, allExpenses) {
  if (!contractTotal || contractTotal === 0) return null;
  const totalExpenses = allExpenses
    .filter(e => String(e['台帳番号']) === String(projectId))
    .reduce((sum, e) => sum + (Number(e['金額']) || 0), 0);
  const margin = ((contractTotal - totalExpenses) / contractTotal) * 100;
  return Math.round(margin * 10) / 10;
}

// ===================================================
//  案件 (projects) CRUD
// ===================================================

/**
 * 全案件を粗利率付きで返す
 * @param {Array} [preloadedExpenses] - 呼び出し元で読み込み済みの経費データ（省略時は内部で読み込む）
 */
function getProjects(preloadedExpenses) {
  try {
    const projects    = getSheetData(SHEETS.PROJECTS);
    const allExpenses = preloadedExpenses || getSheetData(SHEETS.EXPENSES);
    return projects.map(p => {
      const contractTotal = Number(p['契約金額_税込']) || 0;
      const margin = _calcMarginFromExpenses(String(p['台帳番号']), contractTotal, allExpenses);
      return {
        ...p,
        ステータス: parseStatusArray(p['ステータス']),
        粗利率: margin
      };
    });
  } catch (e) {
    return { error: e.message };
  }
}

/** 1案件を関連データ込みで返す */
function getProjectById(projectId) {
  try {
    const projects = getSheetData(SHEETS.PROJECTS);
    const project = projects.find(p => String(p['台帳番号']) === String(projectId));
    if (!project) return { error: '案件が見つかりません: ' + projectId };

    const contractTotal = Number(project['契約金額_税込']) || 0;
    const margin = calcGrossMarginValue(String(projectId), contractTotal);

    return {
      ...project,
      ステータス: parseStatusArray(project['ステータス']),
      粗利率: margin,
      contractChanges: getContractChanges(projectId),
      expenses: getExpenses(projectId),
      billings: getBillings(projectId)
    };
  } catch (e) {
    return { error: e.message };
  }
}

/** 案件を保存（台帳番号で既存判定、insert or update） */
function saveProject(projectData) {
  try {
    const sheet = getSheet(SHEETS.PROJECTS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const projectId = String(projectData['台帳番号']);

    const base = Number(projectData['契約金額_本体']) || 0;
    const rate = Number(projectData['目標粗利率']) || 0;
    projectData['契約金額_消費税'] = Math.round(base * 0.1);
    projectData['契約金額_税込'] = base + Math.round(base * 0.1);
    projectData['目標金額_本体'] = Math.round(base * (rate / 100));
    projectData['目標金額_消費税'] = Math.round(projectData['目標金額_本体'] * 0.1);
    projectData['目標金額_税込'] = projectData['目標金額_本体'] + projectData['目標金額_消費税'];

    const statusArray = calcStatus(projectId, projectData['契約金額_税込']);
    projectData['ステータス'] = JSON.stringify(statusArray);
    projectData['更新日時'] = now();

    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === projectId) {
        existingRowIndex = i;
        break;
      }
    }

    if (existingRowIndex > 0) {
      const rowValues = headers.map(h => (projectData[h] !== undefined ? projectData[h] : ''));
      sheet.getRange(existingRowIndex + 1, 1, 1, headers.length).setValues([rowValues]);
    } else {
      projectData['作成日時'] = now();
      const rowValues = headers.map(h => (projectData[h] !== undefined ? projectData[h] : ''));
      sheet.appendRow(rowValues);
    }

    return { success: true, id: projectId };
  } catch (e) {
    return { error: e.message };
  }
}

/** 案件削除（関連テーブルもカスケード削除） */
function deleteProject(projectId) {
  try {
    const sheet = getSheet(SHEETS.PROJECTS);
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(projectId)) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    if (!found) return { error: '案件が見つかりません: ' + projectId };

    _deleteRelatedRows(SHEETS.CONTRACT_CHANGES, projectId);
    _deleteRelatedRows(SHEETS.EXPENSES, projectId);
    _deleteRelatedRows(SHEETS.BILLINGS, projectId);

    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

function _deleteRelatedRows(sheetName, projectId) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(projectId)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function parseStatusArray(raw) {
  if (Array.isArray(raw)) return raw;
  if (!raw || raw === '') return [];
  try {
    return JSON.parse(raw);
  } catch {
    return String(raw).split(',').map(s => s.trim()).filter(Boolean);
  }
}

// ===================================================
//  契約変更履歴 (contract_changes) CRUD
// ===================================================

function getContractChanges(projectId) {
  try {
    const all = getSheetData(SHEETS.CONTRACT_CHANGES);
    return all.filter(r => String(r['台帳番号']) === String(projectId));
  } catch (e) {
    return { error: e.message };
  }
}

function saveContractChange(changeData) {
  return saveRecord(SHEETS.CONTRACT_CHANGES, changeData);
}

function deleteContractChange(id) {
  return deleteById(SHEETS.CONTRACT_CHANGES, id);
}

// ===================================================
//  経費 (expenses) CRUD
// ===================================================

function getExpenses(projectId) {
  try {
    const all = getSheetData(SHEETS.EXPENSES);
    return all.filter(r => String(r['台帳番号']) === String(projectId));
  } catch (e) {
    return { error: e.message };
  }
}

function saveExpense(expenseData) {
  return saveRecord(SHEETS.EXPENSES, expenseData);
}

function deleteExpense(id) {
  return deleteById(SHEETS.EXPENSES, id);
}

// ===================================================
//  請求入金 (billings) CRUD
// ===================================================

function getBillings(projectId) {
  try {
    const all = getSheetData(SHEETS.BILLINGS);
    return all.filter(r => String(r['台帳番号']) === String(projectId));
  } catch (e) {
    return { error: e.message };
  }
}

function saveBilling(billingData) {
  return saveRecord(SHEETS.BILLINGS, billingData);
}

function deleteBilling(id) {
  return deleteById(SHEETS.BILLINGS, id);
}

// ===================================================
//  ダッシュボードデータ集計
// ===================================================

/**
 * ホーム画面用のダッシュボードデータを返す
 * - サマリーテーブル（計画/見込/実績）
 * - 案件一覧（粗利率付き）
 * - 支払先ランキング
 * - リスク案件（粗利率 < 10%）
 */
function getDashboardData() {
  try {
    // 各シートを1回ずつだけ読む
    const allExpenses = getSheetData(SHEETS.EXPENSES);
    const allBillings = getSheetData(SHEETS.BILLINGS);
    const projects    = getProjects(allExpenses); // 読み込み済み経費を渡して二重読み込みを防ぐ
    if (projects.error) return projects;

    // --- 計画 ---
    const planContract = projects.reduce((s, p) => s + (Number(p['契約金額_税込']) || 0), 0);
    const planTargetProfit = projects.reduce((s, p) => s + (Number(p['目標金額_税込']) || 0), 0);
    const planExpense = planContract - planTargetProfit;
    const planMargin = planContract > 0 ? Math.round(planTargetProfit / planContract * 1000) / 10 : 0;

    // --- 見込（請求金額合計 vs 経費合計） ---
    const forecastIncome = allBillings.reduce((s, b) => s + (Number(b['請求金額']) || 0), 0);
    const forecastExpense = allExpenses.reduce((s, e) => s + (Number(e['金額']) || 0), 0);
    const forecastProfit = forecastIncome - forecastExpense;
    const forecastMargin = forecastIncome > 0 ? Math.round(forecastProfit / forecastIncome * 1000) / 10 : 0;

    // --- 実績（入金確認額 vs 経費合計） ---
    const actualIncome = allBillings.reduce((s, b) => s + (Number(b['入金確認額']) || 0), 0);
    const actualExpense = forecastExpense;
    const actualProfit = actualIncome - actualExpense;
    const actualMargin = actualIncome > 0 ? Math.round(actualProfit / actualIncome * 1000) / 10 : 0;

    // --- 支払先ランキング ---
    const vendorMap = {};
    allExpenses.forEach(e => {
      const v = e['仕入先'] || '未設定';
      vendorMap[v] = (vendorMap[v] || 0) + (Number(e['金額']) || 0);
    });
    const vendorRanking = Object.entries(vendorMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([name, amount], idx) => ({ rank: idx + 1, name, amount }));

    // --- リスク案件（粗利率 < 10%） ---
    const riskProjects = projects
      .filter(p => p['粗利率'] !== null && p['粗利率'] < 10)
      .map(p => ({ 台帳番号: p['台帳番号'], 案件名: p['案件名'], 粗利率: p['粗利率'] }));

    return {
      summary: {
        plan:     { contract: planContract,    expense: planExpense,    profit: planTargetProfit, margin: planMargin },
        forecast: { contract: forecastIncome,  expense: forecastExpense, profit: forecastProfit,  margin: forecastMargin },
        actual:   { contract: actualIncome,    expense: actualExpense,   profit: actualProfit,    margin: actualMargin }
      },
      projects: projects,
      vendorRanking: vendorRanking,
      riskProjects: riskProjects
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  キャッシュフローデータ
// ===================================================

/**
 * 入出金管理画面用データ
 * @param {string} startYM - 開始年月 "YYYY-MM"
 * @param {string} endYM   - 終了年月 "YYYY-MM"
 * @param {string} projectId - 案件フィルター（空文字=全案件）
 */
function getCashflowData(startYM, endYM, projectId) {
  try {
    const allBillings = getSheetData(SHEETS.BILLINGS);
    const allExpenses = getSheetData(SHEETS.EXPENSES);
    const projects    = getSheetData(SHEETS.PROJECTS);

    // 月リスト生成
    const months = _buildMonthList(startYM, endYM);

    // フィルター
    const filteredBillings = projectId
      ? allBillings.filter(b => String(b['台帳番号']) === String(projectId))
      : allBillings;
    const filteredExpenses = projectId
      ? allExpenses.filter(e => String(e['台帳番号']) === String(projectId))
      : allExpenses;

    // 月別集計マトリクス
    const matrix = {};
    months.forEach(m => {
      matrix[m] = { incomeSchedule: 0, incomeActual: 0, expenseActual: 0 };
    });

    filteredBillings.forEach(b => {
      const schedYM = _dateToYM(b['入金予定日']);
      const actualYM = _dateToYM(b['請求日']);
      if (schedYM && matrix[schedYM] !== undefined) {
        matrix[schedYM].incomeSchedule += Number(b['請求金額']) || 0;
      }
      if (actualYM && matrix[actualYM] !== undefined) {
        matrix[actualYM].incomeActual += Number(b['入金確認額']) || 0;
      }
    });

    filteredExpenses.forEach(e => {
      const ym = _dateToYM(e['月']);
      if (ym && matrix[ym] !== undefined) {
        matrix[ym].expenseActual += Number(e['金額']) || 0;
      }
    });

    // 詳細明細（フィルター済み）
    const details = _buildCashflowDetails(filteredBillings, filteredExpenses, projects, startYM, endYM);

    // 合計
    const totals = months.reduce((acc, m) => {
      acc.incomeSchedule += matrix[m].incomeSchedule;
      acc.incomeActual   += matrix[m].incomeActual;
      acc.expenseActual  += matrix[m].expenseActual;
      return acc;
    }, { incomeSchedule: 0, incomeActual: 0, expenseActual: 0 });

    return { months, matrix, details, totals };
  } catch (e) {
    return { error: e.message };
  }
}

function _buildMonthList(startYM, endYM) {
  const months = [];
  const [sy, sm] = startYM.split('-').map(Number);
  const [ey, em] = endYM.split('-').map(Number);
  let y = sy, m = sm;
  while (y < ey || (y === ey && m <= em)) {
    months.push(`${y}-${String(m).padStart(2, '0')}`);
    m++;
    if (m > 12) { m = 1; y++; }
  }
  return months;
}

function _dateToYM(dateVal) {
  if (!dateVal) return null;
  const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function _buildCashflowDetails(billings, expenses, projects, startYM, endYM) {
  const projectMap = {};
  projects.forEach(p => { projectMap[String(p['台帳番号'])] = p['案件名']; });

  const incomeRows = billings
    .filter(b => {
      const ym = _dateToYM(b['請求日']) || _dateToYM(b['入金予定日']);
      return ym && ym >= startYM && ym <= endYM;
    })
    .map(b => ({
      date: b['入金予定日'] || b['請求日'],
      projectId: b['台帳番号'],
      projectName: projectMap[String(b['台帳番号'])] || '',
      category: b['種別'],
      amount: Number(b['請求金額']) || 0,
      confirmed: Number(b['入金確認額']) || 0,
      type: 'income'
    }));

  const expenseRows = expenses
    .filter(e => {
      const ym = _dateToYM(e['月']);
      return ym && ym >= startYM && ym <= endYM;
    })
    .map(e => ({
      date: e['月'],
      projectId: e['台帳番号'],
      projectName: projectMap[String(e['台帳番号'])] || '',
      category: e['仕入先'],
      amount: Number(e['金額']) || 0,
      confirmed: Number(e['金額']) || 0,
      type: 'expense'
    }));

  return { income: incomeRows, expense: expenseRows };
}

// ===================================================
//  全経費一覧（フィルター付き）
// ===================================================

/**
 * 経費一覧画面用
 * @param {string} month      - "YYYY-MM" or ""
 * @param {string} projectId  - 台帳番号 or ""
 * @param {string} vendor     - 仕入先 or ""
 */
function getAllExpenses(month, projectId, vendor) {
  try {
    const allExpenses = getSheetData(SHEETS.EXPENSES);
    const projects    = getSheetData(SHEETS.PROJECTS);
    const projectMap  = {};
    projects.forEach(p => { projectMap[String(p['台帳番号'])] = p['案件名']; });

    let rows = allExpenses;
    if (month)     rows = rows.filter(e => _dateToYM(e['月']) === month);
    if (projectId) rows = rows.filter(e => String(e['台帳番号']) === String(projectId));
    if (vendor)    rows = rows.filter(e => (e['仕入先'] || '').includes(vendor));

    const result = rows.map(e => ({
      ...e,
      案件名: projectMap[String(e['台帳番号'])] || ''
    }));

    const total = result.reduce((s, e) => s + (Number(e['金額']) || 0), 0);

    return { rows: result, total };
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  ガントチャートデータ
// ===================================================

function getGanttData() {
  try {
    // projects と expenses をそれぞれ1回だけ読む（getProjects()経由は二重読み込みになるため使わない）
    const projects    = getSheetData(SHEETS.PROJECTS);
    const allExpenses = getSheetData(SHEETS.EXPENSES);

    const rows = projects.map(p => {
      const pid = String(p['台帳番号']);
      const contractTotal = Number(p['契約金額_税込']) || 0;
      const expTotal = allExpenses
        .filter(e => String(e['台帳番号']) === pid)
        .reduce((s, e) => s + (Number(e['金額']) || 0), 0);
      const profit = contractTotal - expTotal;
      // expTotal を再利用して粗利率を計算（_calcMarginFromExpenses を呼ぶと同じフィルターを二重実行するため使わない）
      const margin = contractTotal > 0 ? Math.round((contractTotal - expTotal) / contractTotal * 1000) / 10 : null;
      return {
        台帳番号:   p['台帳番号'],
        案件名:     p['案件名'],
        客先名:     p['客先名'],
        契約金額:   contractTotal,
        経費合計:   expTotal,
        粗利益:     profit,
        粗利率:     margin,
        工期開始:   p['工期開始'],
        工期終了:   p['工期終了'],
        ステータス: parseStatusArray(p['ステータス'])
      };
    });

    // ガント表示範囲（全案件の最早開始〜最遅終了）
    const dateTimes = rows
      .flatMap(r => [r['工期開始'], r['工期終了']])
      .map(d => (d ? new Date(d).getTime() : null))
      .filter(t => t !== null && !isNaN(t));
    const minDate = dateTimes.length > 0 ? new Date(Math.min.apply(null, dateTimes)) : null;
    const maxDate = dateTimes.length > 0 ? new Date(Math.max.apply(null, dateTimes)) : null;

    return { rows, minDate, maxDate };
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  経費一括登録
// ===================================================

/**
 * @param {string} projectId
 * @param {Array<{month, vendor, amount, note}>} expenseRows
 */
function saveBulkExpenses(projectId, expenseRows) {
  try {
    const results = expenseRows.map(row => {
      return saveRecord(SHEETS.EXPENSES, {
        台帳番号: projectId,
        月:      row.month,
        仕入先:  row.vendor,
        金額:    Number(row.amount) || 0,
        相殺:    '',
        備考:    row.note || ''
      });
    });
    const errors = results.filter(r => r.error);
    if (errors.length > 0) return { error: errors.map(e => e.error).join(', ') };
    return { success: true, count: results.length };
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  仕入先マスタ (vendors) CRUD
// ===================================================

/** 全仕入先を返す */
function getVendors() {
  try {
    return getSheetData(SHEETS.VENDORS);
  } catch (e) {
    return { error: e.message };
  }
}

/** 仕入先を保存（仕入先名で重複チェック） */
function saveVendor(vendorData) {
  try {
    const sheet = getSheet(SHEETS.VENDORS);
    if (!sheet) throw new Error('vendors シートが見つかりません。setupSheets() を実行してください。');
    const name = String(vendorData['仕入先名'] || '').trim();
    if (!name) return { error: '仕入先名は必須です' };
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      if (existing.some(v => String(v) === name)) return { success: true, existing: true };
    }
    sheet.appendRow([name, String(vendorData['よみがな'] || ''), now()]);
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  請求書履歴 (invoices) CRUD
// ===================================================

/** 処理済み請求書一覧を返す（yearMonth="YYYY-MM" でフィルター、省略で全件） */
function getInvoices(yearMonth) {
  try {
    const all = getSheetData(SHEETS.INVOICES);
    if (!yearMonth) return all;
    return all.filter(r => String(r['請求年月']) === String(yearMonth));
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  請求書振分登録
// ===================================================

/**
 * 台帳番号から案件を検索（クライアントキャッシュが使えない場合のフォールバック）
 * ハイフンなし・先頭ゼロあり/なし等の表記ゆれを吸収する
 */
function lookupProject(projectId) {
  try {
    const id = String(projectId || '').trim();
    if (!id) return { found: false };
    const projects = getSheetData(SHEETS.PROJECTS);
    const p = projects.find(r => {
      const pid = String(r['台帳番号']);
      return pid === id ||
             pid === String(parseInt(id,  10)) ||
             id  === String(parseInt(pid, 10));
    });
    if (p) return { found: true, 台帳番号: String(p['台帳番号']), 案件名: p['案件名'] || '', 客先名: p['客先名'] || '' };
    return { found: false };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * 請求書を保存し、明細を各工事台帳の経費として一括登録する
 * @param {Object} invoiceData  {仕入先, 請求年月, 請求総額}
 * @param {Array}  expenseRows  [{台帳番号, 金額, 備考}, ...]
 * @return {Object} {success, invoiceId, expenseCount, skipped} or {error}
 */
function saveInvoiceWithExpenses(invoiceData, expenseRows) {
  try {
    const vendor = String(invoiceData['仕入先']   || '').trim();
    const month  = String(invoiceData['請求年月'] || '').trim();
    const total  = Number(invoiceData['請求総額']) || 0;
    if (!vendor) return { error: '仕入先は必須です' };
    if (!month)  return { error: '請求年月は必須です' };

    // 有効な台帳番号セットを取得
    const projects = getSheetData(SHEETS.PROJECTS);
    const validIds = new Set(projects.map(p => String(p['台帳番号'])));

    let expenseCount = 0, skipped = 0;
    const errors = [];

    (expenseRows || []).forEach((row, i) => {
      const amount = Number(row['金額']) || 0;
      if (amount <= 0) { skipped++; return; }
      const pid = String(row['台帳番号'] || '').trim();
      if (!pid || !validIds.has(pid)) { skipped++; return; }

      const res = saveRecord(SHEETS.EXPENSES, {
        台帳番号: pid,
        月:      month + '-01',
        仕入先:  vendor,
        金額:    amount,
        相殺:    '',
        備考:    String(row['備考'] || '')
      });
      if (res.error) errors.push('行' + (i + 1) + ': ' + res.error);
      else           expenseCount++;
    });

    if (errors.length > 0) {
      return { error: errors.join('\n'), expenseCount, skipped };
    }

    // 請求書履歴を記録
    const invRes = saveRecord(SHEETS.INVOICES, {
      仕入先:   vendor,
      請求年月: month,
      請求総額: total,
      明細件数: expenseCount
    });

    return { success: true, invoiceId: invRes.id, expenseCount, skipped };
  } catch (e) {
    return { error: e.message };
  }
}

// ===================================================
//  初期セットアップ（一度だけ実行）
// ===================================================

function setupSheets() {
  const ss = getSpreadsheet();

  const configs = [
    {
      name: SHEETS.PROJECTS,
      headers: [
        '台帳番号', '客先名', '営業担当', '案件名', '住所',
        '工期開始', '工期終了',
        '契約金額_本体', '契約金額_消費税', '契約金額_税込',
        '目標粗利率', '目標金額_本体', '目標金額_消費税', '目標金額_税込',
        'ステータス', 'アラートメッセージ', '作成日時', '更新日時'
      ]
    },
    {
      name: SHEETS.CONTRACT_CHANGES,
      headers: ['ID', '台帳番号', '変更種別', '変更日', '変更金額_本体', '変更金額_税込', '備考', '作成日時']
    },
    {
      name: SHEETS.EXPENSES,
      headers: ['ID', '台帳番号', '月', '仕入先', '金額', '相殺', '備考', '作成日時']
    },
    {
      name: SHEETS.BILLINGS,
      headers: ['ID', '台帳番号', '請求日', '種別', '請求金額', '入金予定日', '入金確認額', '確認者', '作成日時']
    },
    {
      name: SHEETS.VENDORS,
      headers: ['仕入先名', 'よみがな', '作成日時']
    },
    {
      name: SHEETS.INVOICES,
      headers: ['ID', '仕入先', '請求年月', '請求総額', '明細件数', '作成日時']
    }
  ];

  configs.forEach(config => {
    let sheet = ss.getSheetByName(config.name);
    if (!sheet) sheet = ss.insertSheet(config.name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
      sheet.getRange(1, 1, 1, config.headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  Logger.log('セットアップ完了: 4シートを作成しました。');
}
