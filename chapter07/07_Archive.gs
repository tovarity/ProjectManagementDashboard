/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
/** 完全削除の対象にできる Status の値 */
const ARCHIVED_STATUS = 'Archived';

/** メニュー：選択した Archived の行を削除し、資料を Drive のごみ箱へ移動 */
function deleteArchivedRows() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const def = [CONFIG.projects, CONFIG.tasks].find(d => d.sheetName === sheet.getName());
  if (!def) {
    ui.alert('Projects または Tasks シートで、削除したい行を選択してから実行してください。');
    return;
  }

  const targets = collectSelectedRecords_(sheet, def);
  if (targets.length === 0) {
    ui.alert('削除する行が選択されていません。');
    return;
  }

  const error = validateDeleteTargets_(def, targets);
  if (error) {
    ui.alert('完全削除できません', error, ui.ButtonSet.OK);
    return;
  }

  const label = (def === CONFIG.projects) ? 'プロジェクト' : 'タスク';
  const list = targets.slice(0, 15).map(t => '・' + t.name);
  if (targets.length > 15) list.push('…ほか ' + (targets.length - 15) + ' 件');
  const answer = ui.alert(
    '完全に削除',
    '次の ' + targets.length + ' 件の' + label + 'を完全に削除します。\n'
      + '行を削除し、資料を Drive のごみ箱へ移動します（30日以内ならごみ箱から戻せます）。\n\n'
      + list.join('\n'),
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const n = withLock_(() => executeDelete_(def, new Set(targets.map(t => t.id))));
  toast_((n || 0) + ' 件を完全に削除しました');
}

/** 選択範囲の行を { rowNum, id, name, status, url } の配列で返す（空行は除く） */
function collectSelectedRecords_(sheet, def) {
  const lastRow = sheet.getLastRow();
  const rangeList = sheet.getActiveRangeList();
  if (lastRow < 2 || !rangeList) return [];

  const rowNums = new Set();
  rangeList.getRanges().forEach(r => {
    for (let i = Math.max(r.getRow(), 2); i <= Math.min(r.getLastRow(), lastRow); i++) rowNums.add(i);
  });

  const cols = getCols_(sheet, def.headers);
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const records = [];
  Array.from(rowNums).sort((a, b) => a - b).forEach(rowNum => {
    const row = values[rowNum - 2];
    const id = String(row[cols.id]).trim();
    const name = String(row[cols.name]).trim();
    if (!id && !name) return;
    records.push({
      rowNum: rowNum,
      id: id,
      name: name,
      status: String(row[cols.status]).trim(),
      url: String(row[cols.docUrl]).trim()
    });
  });
  return records;
}

/** 削除してよいかを確認し、問題があればメッセージを返す（問題なければ null） */
function validateDeleteTargets_(def, targets) {
  const notArchived = targets.filter(t => t.status !== ARCHIVED_STATUS);
  if (notArchived.length) {
    return 'Status が「' + ARCHIVED_STATUS + '」ではない行が含まれています。\n'
      + '先に Status を「' + ARCHIVED_STATUS + '」にしてください。\n\n'
      + notArchived.map(t => '・' + t.rowNum + '行目: ' + (t.name || '(名前なし)')).join('\n');
  }
  if (targets.some(t => !t.id)) {
    return 'ID が空の行があります。メニュー「IDを採番」を実行してから、もう一度お試しください。';
  }
  if (def !== CONFIG.projects) return null;

  // タスクが紐付いているプロジェクトは削除しない
  const tSheet = getSheet_(CONFIG.tasks);
  const linked = new Map();
  if (tSheet.getLastRow() >= 2) {
    const tc = getCols_(tSheet, { project: CONFIG.tasks.headers.project });
    tSheet.getRange(2, tc.project + 1, tSheet.getLastRow() - 1, 1).getValues().forEach(r => {
      const p = String(r[0]).trim();
      if (p) linked.set(p, (linked.get(p) || 0) + 1);
    });
  }
  const blocked = targets.filter(t => linked.has(t.name));
  if (blocked.length) {
    return 'タスクが紐付いているプロジェクトがあります。\n'
      + '先にタスクを完全に削除するか、別のプロジェクトへ移してください。\n\n'
      + blocked.map(t => '・' + t.name + '（タスク ' + linked.get(t.name) + ' 件）').join('\n');
  }

  // フォルダにほかのファイル（孤立したタスク資料など）が残っていれば削除しない
  const notEmpty = targets.filter(t => {
    const folder = projectFolderOf_(t.url);
    if (!folder) return false;
    const docId = fileIdFromUrl_(t.url);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getId() !== docId && !f.isTrashed()) return true;
    }
    return false;
  });
  if (notEmpty.length) {
    return 'プロジェクトのフォルダに、ほかのファイルが残っています。\n'
      + '誤って行を消したタスク資料の可能性があるため、先に「Driveから復元」を実行するか、Drive で中身を確認してください。\n\n'
      + notEmpty.map(t => '・' + t.name).join('\n');
  }
  return null;
}

/** プロジェクト資料の親フォルダ（ダッシュボード直下と「_」で始まるフォルダは対象外） */
function projectFolderOf_(url) {
  const id = fileIdFromUrl_(url);
  if (!id) return null;
  try {
    const parents = DriveApp.getFileById(id).getParents();
    if (!parents.hasNext()) return null;
    const folder = parents.next();
    if (folder.getId() === getBaseFolder_().getId() || folder.getName().indexOf('_') === 0) return null;
    return folder;
  } catch (err) {
    return null;
  }
}

/** ID で行を特定し直してから、資料をごみ箱へ移動し、行を下から削除する */
function executeDelete_(def, idSet) {
  const sheet = getSheet_(def);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const cols = getCols_(sheet, def.headers);
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const rowNums = [];
  values.forEach((row, i) => {
    if (!idSet.has(String(row[cols.id]).trim())) return;
    if (String(row[cols.status]).trim() !== ARCHIVED_STATUS) return;
    trashDoc_(def, String(row[cols.docUrl]).trim());
    rowNums.push(i + 2);
  });
  rowNums.reverse().forEach(r => sheet.deleteRow(r));
  return rowNums.length;
}

/** プロジェクトはフォルダごと、タスクはドキュメントだけをごみ箱へ */
function trashDoc_(def, url) {
  const id = fileIdFromUrl_(url);
  if (!id) return;
  try {
    if (def === CONFIG.projects) {
      const folder = projectFolderOf_(url);
      if (folder) {
        folder.setTrashed(true);
        return;
      }
    }
    DriveApp.getFileById(id).setTrashed(true);
  } catch (err) {
    console.warn('ごみ箱への移動に失敗: ' + url + ' / ' + err);
  }
}