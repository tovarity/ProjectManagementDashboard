/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
const TASK_SECTIONS = ['1. 概要', '2. 完了条件', '3. メモ'];

/** メニュー：資料が未作成のタスクをまとめて処理 */
function createMissingTaskDocs() {
  const s = withLock_(() => {
    assignIds_(CONFIG.tasks);
    return processTasks_(2, getSheet_(CONFIG.tasks).getLastRow());
  });
  if (s) {
    toast_('タスク資料：新規 ' + s.created + ' 件 / 移動 ' + s.moved + ' 件'
      + (s.skipped.length ? ' / 対象外の行: ' + s.skipped.join(', ') : ''));
  }
}

/** 指定行：資料がなければ作成、あれば名前と保存先を同期し、Nameにリンクを付ける */
function processTasks_(startRow, endRow) {
  const stats = { created: 0, moved: 0, skipped: [] };
  if (endRow < startRow) return stats;
  const sheet = getSheet_(CONFIG.tasks);
  const cols = getCols_(sheet, CONFIG.tasks.headers);
  const values = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();
  const folders = new Map();

  values.forEach((row, i) => {
    const name = String(row[cols.name]).trim();
    if (!name) return;
    const rowNum = startRow + i;
    const project = String(row[cols.project]).trim();
    if (!folders.has(project)) folders.set(project, getTaskFolder_(project));
    const folder = folders.get(project);
    if (!folder) {
      stats.skipped.push(rowNum);
      return;
    }

    const docName = displayName_(name, row[cols.id]);
    let url = String(row[cols.docUrl]).trim();
    if (!url) {
      url = createDoc_(folder, docName, taskInfoRows_(row, cols), TASK_SECTIONS);
      sheet.getRange(rowNum, cols.docUrl + 1).setValue(url);
      stats.created++;
    } else {
      if (renameDoc_(url, docName, null)) setInfoName_(url, 'タスク名', name);
      if (moveToFolder_(url, folder)) stats.moved++;
    }
    setNameLink_(sheet, rowNum, cols.name, name, url);
  });
  return stats;
}

/** 基本情報テーブルの内容（1行目は名前：印刷・PDFでも名前が分かるように） */
function taskInfoRows_(row, cols) {
  return [
    ['タスク名', formatValue_(row[cols.name])],
    ['タスクID', formatValue_(row[cols.id])],
    ['プロジェクト', formatValue_(row[cols.project])],
    ['優先度', formatValue_(row[cols.priority])],
    ['ステータス', formatValue_(row[cols.status])],
    ['期日', formatValue_(row[cols.dueDate])]
  ];
}

/** Project が空 → 未割当フォルダ / あり → プロジェクト資料のフォルダ / 見つからない → null */
function getTaskFolder_(project) {
  if (!project) return getOrCreateFolder_(getBaseFolder_(), CONFIG.unassignedFolderName);
  const sheet = getSheet_(CONFIG.projects);
  if (sheet.getLastRow() < 2) return null;
  const cols = getCols_(sheet, CONFIG.projects.headers);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const hit = rows.find(r => String(r[cols.name]).trim() === project && fileIdFromUrl_(r[cols.docUrl]));
  if (!hit) return null;
  try {
    const p = DriveApp.getFileById(fileIdFromUrl_(hit[cols.docUrl])).getParents();
    return p.hasNext() ? p.next() : null;
  } catch (err) {
    return null;
  }
}

function moveToFolder_(url, folder) {
  const id = fileIdFromUrl_(url);
  if (!id) return false;
  try {
    const file = DriveApp.getFileById(id);
    const p = file.getParents();
    if (p.hasNext() && p.next().getId() === folder.getId()) return false;
    file.moveTo(folder);
    return true;
  } catch (err) {
    console.warn('移動に失敗: ' + url + ' / ' + err);
    return false;
  }
}