/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
/** シート設計（見出し名で列を探すので、列の並び替えや非表示をしても動く） */
const CONFIG = {
  projects: {
    sheetName: 'Projects',
    idPrefix: 'PRJ',
    headers: { id: 'Id', name: 'Name', priority: 'Priority', status: 'Status', dueDate: 'DueDate', docUrl: 'Document URL' }
  },
  tasks: {
    sheetName: 'Tasks',
    idPrefix: 'TSK',
    headers: { id: 'Id', project: 'Project', name: 'Name', priority: 'Priority', status: 'Status', dueDate: 'DueDate', docUrl: 'Document URL' }
  },
  unassignedFolderName: '_未割当タスク'
};

function getSheet_(def) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(def.sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + def.sheetName);
  return sheet;
}

/** 見出し名 → 列番号（0始まり） */
function getCols_(sheet, headers) {
  const row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const cols = {};
  Object.keys(headers).forEach(key => {
    const idx = row.indexOf(headers[key]);
    if (idx < 0) throw new Error(sheet.getName() + ' に「' + headers[key] + '」列がありません');
    cols[key] = idx;
  });
  return cols;
}

/** 同時実行を防ぐ（ロックが取れなければ何もしない） */
function withLock_(fn, waitMs) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(waitMs || 30000)) return null;
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, 'プロジェクト管理', 8);
}

function fileIdFromUrl_(url) {
  const m = String(url || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

function formatValue_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy/MM/dd');
  }
  return (v === null || v === undefined) ? '' : String(v);
}