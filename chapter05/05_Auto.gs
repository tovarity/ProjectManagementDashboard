/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
/** メニュー：編集時に自動処理するトリガーを登録（初回のみ） */
function enableAutoCreate() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onEditAuto');
  if (!exists) {
    ScriptApp.newTrigger('onEditAuto').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  }
  toast_(exists ? '自動作成はすでにオンです' : '自動作成をオンにしました');
}

/** インストール型トリガー：編集された行の資料作成・名前同期・基本情報同期 */
function onEditAuto(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();
  const def = [CONFIG.projects, CONFIG.tasks].find(d => d.sheetName === name);
  if (!def) return;
  const startRow = Math.max(e.range.getRow(), 2);
  const endRow = e.range.getLastRow();
  if (endRow < startRow) return;

  withLock_(() => {
    assignIds_(def);
    if (def === CONFIG.projects) processProjects_(startRow, endRow);
    else processTasks_(startRow, endRow);
    syncInfo_(def, startRow, endRow);
  });
}

/** メニュー：全行の基本情報テーブルを同期 */
function syncAllInfo() {
  const n = withLock_(() =>
    syncInfo_(CONFIG.projects, 2, getSheet_(CONFIG.projects).getLastRow())
    + syncInfo_(CONFIG.tasks, 2, getSheet_(CONFIG.tasks).getLastRow()));
  toast_('基本情報を ' + (n || 0) + ' 件更新しました');
}

/** 指定行のドキュメントの基本情報テーブルをシートの値に合わせる */
function syncInfo_(def, startRow, endRow) {
  if (endRow < startRow) return 0;
  const sheet = getSheet_(def);
  const cols = getCols_(sheet, def.headers);
  const values = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();
  let updated = 0;

  values.forEach(row => {
    const id = fileIdFromUrl_(row[cols.docUrl]);
    if (!id) return;
    const infoRows = (def === CONFIG.projects) ? projectInfoRows_(row, cols) : taskInfoRows_(row, cols);
    try {
      const doc = DocumentApp.openById(id);
      const table = findInfoTable_(doc);
      if (!table) return;
      let changed = false;
      infoRows.forEach(([label, value]) => {
        if (setTableValue_(table, label, value)) changed = true;
      });
      if (changed) {
        doc.saveAndClose();
        updated++;
      }
    } catch (err) {
      console.warn('基本情報の同期に失敗: ' + id + ' / ' + err);
    }
  });
  return updated;
}

/** ラベル行の値を更新（変化があれば true）。行がなければ追加 */
function setTableValue_(table, label, value) {
  for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 2 || row.getCell(0).getText().trim() !== label) continue;
    if (row.getCell(1).getText() === value) return false;
    row.getCell(1).setText(value);
    return true;
  }
  const newRow = table.appendTableRow();
  newRow.appendTableCell(label);
  newRow.appendTableCell(value);
  return true;
}