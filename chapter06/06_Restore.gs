/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
/** メニュー：どの行からも参照されていない資料を探し、基本情報テーブルから行を復元 */
function restoreFromDrive() {
  const ui = SpreadsheetApp.getUi();
  const found = withLock_(() => findOrphanDocs_());
  if (!found) return;
  if (found.length === 0) {
    ui.alert('復元できる資料はありません。');
    return;
  }

  const list = found.slice(0, 15).map(d => '・[' + (d.type === 'project' ? 'プロジェクト' : 'タスク') + '] ' + d.name);
  if (found.length > 15) list.push('…ほか ' + (found.length - 15) + ' 件');
  const answer = ui.alert('Driveから復元', '次の ' + found.length + ' 件の行を復元します。\n\n' + list.join('\n'), ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const r = withLock_(() => {
    const docs = findOrphanDocs_();
    const p = appendRestoredRows_(CONFIG.projects, docs.filter(d => d.type === 'project'), null);
    if (p.count) {
      processProjects_(p.startRow, p.startRow + p.count - 1);
      syncInfo_(CONFIG.projects, p.startRow, p.startRow + p.count - 1);
    }
    const t = appendRestoredRows_(CONFIG.tasks, docs.filter(d => d.type === 'task'), p.renamed);
    if (t.count) {
      processTasks_(t.startRow, t.startRow + t.count - 1);
      syncInfo_(CONFIG.tasks, t.startRow, t.startRow + t.count - 1);
    }
    return { total: docs.length, renamed: p.renamed.size };
  });
  if (r) {
    toast_(r.total + ' 件の行を復元しました'
      + (r.renamed ? '（同じ名前があったため「（復元）」を付けたもの: ' + r.renamed + ' 件）' : ''));
  }
}

/** ダッシュボード配下のフォルダから、シートに存在しない資料を探す */
function findOrphanDocs_() {
  const referenced = new Set();
  const existingIds = new Set();
  [CONFIG.projects, CONFIG.tasks].forEach(def => {
    const sheet = getSheet_(def);
    if (sheet.getLastRow() < 2) return;
    const cols = getCols_(sheet, { id: def.headers.id, docUrl: def.headers.docUrl });
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().forEach(r => {
      const fid = fileIdFromUrl_(r[cols.docUrl]);
      if (fid) referenced.add(fid);
      const id = String(r[cols.id]).trim();
      if (id) existingIds.add(id);
    });
  });

  const docs = [];
  const folders = getBaseFolder_().getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const fname = folder.getName();
    if (folder.isTrashed()) continue;
    if (fname.indexOf('_') === 0 && fname !== CONFIG.unassignedFolderName) continue;

    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed() || referenced.has(f.getId())) continue;
      const info = readInfoTable_(f.getId());
      if (!info) continue;
      const isProject = Object.prototype.hasOwnProperty.call(info, 'プロジェクトID');
      const isTask = Object.prototype.hasOwnProperty.call(info, 'タスクID');
      if (!isProject && !isTask) continue;
      const recordId = isProject ? info['プロジェクトID'] : info['タスクID'];
      if (recordId && existingIds.has(recordId)) continue;
      // 名前は表から読む（なければファイル名から短縮IDを外す）
      const nameInTable = isProject ? info['プロジェクト名'] : info['タスク名'];
      const name = nameInTable || f.getName().replace(/\s*\[(PRJ|TSK)-[0-9a-f]{8}\]$/, '');
      docs.push({ type: isProject ? 'project' : 'task', name: name, url: f.getUrl(), info: info });
    }
  }
  return docs;
}

/** 基本情報テーブルを { ラベル: 値 } で返す */
function readInfoTable_(fileId) {
  try {
    const table = findInfoTable_(DocumentApp.openById(fileId));
    if (!table) return null;
    const info = {};
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() >= 2) info[row.getCell(0).getText().trim()] = row.getCell(1).getText().trim();
    }
    return info;
  } catch (err) {
    return null;
  }
}

/** 復元した行をシート末尾に追加する（プロジェクト名が重なる場合は「（復元）」を付ける） */
function appendRestoredRows_(def, docs, projectRenamed) {
  const result = { startRow: 0, count: 0, renamed: new Map() };
  if (docs.length === 0) return result;
  const sheet = getSheet_(def);
  const cols = getCols_(sheet, def.headers);
  const isProject = (def === CONFIG.projects);
  const lastRow = sheet.getLastRow();

  const usedNames = new Set(lastRow < 2 ? [] :
    sheet.getRange(2, cols.name + 1, lastRow - 1, 1).getValues().map(r => normalizeName_(r[0])));

  let projectNames = null;
  if (!isProject) {
    const p = getSheet_(CONFIG.projects);
    const pc = getCols_(p, { name: CONFIG.projects.headers.name });
    projectNames = new Set(p.getLastRow() < 2 ? [] :
      p.getRange(2, pc.name + 1, p.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()));
  }

  const width = sheet.getLastColumn();
  const rows = docs.map(d => {
    let name = d.name;
    if (isProject) {
      let k = 1;
      while (usedNames.has(normalizeName_(name))) {
        name = d.name + '（復元' + (k > 1 ? k : '') + '）';
        k++;
      }
      if (name !== d.name) result.renamed.set(d.name, name);
      usedNames.add(normalizeName_(name));
    }

    const row = new Array(width).fill('');
    row[cols.id] = (isProject ? d.info['プロジェクトID'] : d.info['タスクID']) || '';
    row[cols.name] = name;
    row[cols.priority] = d.info['優先度'] || '';
    row[cols.status] = d.info['ステータス'] || '';
    row[cols.dueDate] = parseDate_(d.info['期日']);
    row[cols.docUrl] = d.url;
    if (!isProject) {
      let p = d.info['プロジェクト'] || '';
      if (projectRenamed && projectRenamed.has(p)) p = projectRenamed.get(p);
      row[cols.project] = projectNames.has(p) ? p : '';
    }
    return row;
  });

  result.startRow = lastRow + 1;
  result.count = rows.length;
  sheet.getRange(result.startRow, 1, rows.length, width).setValues(rows);
  rows.forEach((row, i) => setNameLink_(sheet, result.startRow + i, cols.name, row[cols.name], row[cols.docUrl]));
  return result;
}

function parseDate_(s) {
  const m = String(s || '').match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : (s || '');
}