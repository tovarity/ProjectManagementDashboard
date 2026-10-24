/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
const PROJECT_SECTIONS = ['1. 概要', '2. 目的・ゴール', '3. 詳細', '4. メモ'];

/** Drive上の表示名：名前 [PRJ-xxxxxxxx]（同じ名前でも区別できるよう短縮IDを付ける） */
function displayName_(name, id) {
  const m = String(id || '').match(/^([A-Z]+-[0-9a-f]{8})/);
  return m ? name + ' [' + m[1] + ']' : name;
}

/** 名前の比較用（前後の空白と大文字・小文字の違いを無視） */
function normalizeName_(s) {
  return String(s || '').trim().toLowerCase();
}

/** 重複している名前の一覧 */
function findDuplicateNames_(names) {
  const groups = new Map();
  names.forEach(n => {
    const key = normalizeName_(n);
    if (!key) return;
    groups.set(key, (groups.get(key) || []).concat([String(n).trim()]));
  });
  return Array.from(groups.values()).filter(g => g.length > 1).map(g => g[0]);
}

/** メニュー：資料が未作成のプロジェクトをまとめて処理 */
function createMissingProjectDocs() {
  const s = withLock_(() => {
    assignIds_(CONFIG.projects);
    return processProjects_(2, getSheet_(CONFIG.projects).getLastRow());
  });
  if (s) {
    toast_('プロジェクト資料：新規 ' + s.created + ' 件 / 名前の更新 ' + s.renamed + ' 件'
      + (s.duplicates.length ? ' / ⚠ 同じ名前があります: ' + s.duplicates.join(', ') : ''));
  }
}

/** 指定行：資料がなければ作成、あれば名前を同期し、Nameにリンクを付ける */
function processProjects_(startRow, endRow) {
  const stats = { created: 0, renamed: 0, duplicates: [] };
  if (endRow < startRow) return stats;
  const sheet = getSheet_(CONFIG.projects);
  const cols = getCols_(sheet, CONFIG.projects.headers);
  const values = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();
  const base = getBaseFolder_();

  values.forEach((row, i) => {
    const name = String(row[cols.name]).trim();
    if (!name) return;
    const rowNum = startRow + i;
    const docName = displayName_(name, row[cols.id]);
    let url = String(row[cols.docUrl]).trim();

    if (!url) {
      const folder = getOrCreateFolder_(base, docName);
      url = createDoc_(folder, docName, projectInfoRows_(row, cols), PROJECT_SECTIONS);
      sheet.getRange(rowNum, cols.docUrl + 1).setValue(url);
      stats.created++;
    } else if (renameDoc_(url, docName, base)) {
      setInfoName_(url, 'プロジェクト名', name);
      stats.renamed++;
    }
    setNameLink_(sheet, rowNum, cols.name, name, url);
  });

  // 入力規則をすり抜けた重複（貼り付けなど）を警告する
  const lastRow = sheet.getLastRow();
  stats.duplicates = findDuplicateNames_(lastRow < 2 ? [] :
    sheet.getRange(2, cols.name + 1, lastRow - 1, 1).getValues().map(r => r[0]));
  if (stats.duplicates.length) {
    toast_('⚠ 同じプロジェクト名があります: ' + stats.duplicates.join(', ') + '（どちらかの名前を変えてください）');
  }
  return stats;
}

/** 基本情報テーブルの内容（1行目は名前：印刷・PDFでも名前が分かるように） */
function projectInfoRows_(row, cols) {
  return [
    ['プロジェクト名', formatValue_(row[cols.name])],
    ['プロジェクトID', formatValue_(row[cols.id])],
    ['優先度', formatValue_(row[cols.priority])],
    ['ステータス', formatValue_(row[cols.status])],
    ['期日', formatValue_(row[cols.dueDate])]
  ];
}

/** スプレッドシートが保存されているフォルダ */
function getBaseFolder_() {
  const parents = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
  if (!parents.hasNext()) throw new Error('スプレッドシートの保存先フォルダが取得できません');
  return parents.next();
}

/** 同名フォルダがあれば再利用（ごみ箱内は除く）、なければ作成 */
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return parent.createFolder(name);
}

/** 基本情報テーブルと見出しを持つドキュメントを作成し、URLを返す（本文にタイトル行は作らない） */
function createDoc_(folder, docName, infoRows, sections) {
  const doc = DocumentApp.create(docName);
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);

  const body = doc.getBody();
  const first = body.getParagraphs()[0];
  first.setText('基本情報');
  first.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const table = body.appendTable([['項目', '内容']].concat(infoRows));
  table.getRow(0).editAsText().setBold(true);

  sections.forEach(h => {
    body.appendParagraph(h).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(' ').setHeading(DocumentApp.ParagraphHeading.NORMAL);
  });
  doc.saveAndClose();
  return file.getUrl();
}

/** ファイル名を表示名に合わせる（base を渡すと親フォルダ名も合わせる） */
function renameDoc_(url, docName, base) {
  const id = fileIdFromUrl_(url);
  if (!id) return false;
  try {
    const file = DriveApp.getFileById(id);
    let changed = false;
    if (file.getName() !== docName) {
      file.setName(docName);
      changed = true;
    }
    if (base) {
      const p = file.getParents();
      if (p.hasNext()) {
        const folder = p.next();
        if (folder.getId() !== base.getId() && folder.getName() !== docName) {
          folder.setName(docName);
          changed = true;
        }
      }
    }
    return changed;
  } catch (err) {
    console.warn('名前の同期に失敗: ' + url + ' / ' + err);
    return false;
  }
}

/** 基本情報テーブルの名前の行を更新（行がなければ1行目に追加） */
function setInfoName_(url, label, name) {
  const id = fileIdFromUrl_(url);
  if (!id) return;
  try {
    const doc = DocumentApp.openById(id);
    const table = findInfoTable_(doc);
    if (!table) return;
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 2 || row.getCell(0).getText().trim() !== label) continue;
      if (row.getCell(1).getText() !== name) {
        row.getCell(1).setText(name);
        doc.saveAndClose();
      }
      return;
    }
    const newRow = table.insertTableRow(1);
    newRow.appendTableCell(label);
    newRow.appendTableCell(name);
    doc.saveAndClose();
  } catch (err) {
    console.warn('名前の行の更新に失敗: ' + url + ' / ' + err);
  }
}

/** Name セルにドキュメントへのリンクを付ける（変化がある場合のみ） */
function setNameLink_(sheet, rowNum, colIdx, name, url) {
  const cell = sheet.getRange(rowNum, colIdx + 1);
  const rt = cell.getRichTextValue();
  if (rt && rt.getText() === name && rt.getLinkUrl() === url) return;
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(name).setLinkUrl(url).build());
}

/** 先頭セルが「項目」のテーブルを基本情報テーブルとみなす */
function findInfoTable_(doc) {
  const tables = doc.getBody().getTables();
  for (let i = 0; i < tables.length; i++) {
    const t = tables[i];
    if (t.getNumRows() > 0 && t.getRow(0).getNumCells() >= 2 && t.getCell(0, 0).getText().trim() === '項目') {
      return t;
    }
  }
  return null;
}