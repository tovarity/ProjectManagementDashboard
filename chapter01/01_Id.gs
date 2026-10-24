/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
function newId_(prefix) {
  return prefix + '-' + Utilities.getUuid().toLowerCase();
}

function isValidId_(id, prefix) {
  return new RegExp('^' + prefix + '-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$').test(id);
}

/** 空欄・重複・規則外のIDを採番し直し、変更件数を返す */
function assignIds_(def) {
  const sheet = getSheet_(def);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const idCol = getCols_(sheet, { id: def.headers.id }).id;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const seen = new Set();
  let changed = 0;

  const ids = values.map(row => {
    let id = String(row[idCol]).trim();
    const hasData = row.some((v, c) => c !== idCol && v !== '' && v !== null && v !== false);
    if (!id && !hasData) return [''];
    if (!isValidId_(id, def.idPrefix) || seen.has(id)) {
      do {
        id = newId_(def.idPrefix);
      } while (seen.has(id));
      changed++;
    }
    seen.add(id);
    return [id];
  });

  if (changed > 0) sheet.getRange(2, idCol + 1, ids.length, 1).setValues(ids);
  return changed;
}

/** メニュー：両シートのIDを採番 */
function assignAllIds() {
  const n = withLock_(() => assignIds_(CONFIG.projects) + assignIds_(CONFIG.tasks));
  toast_('IDを ' + (n || 0) + ' 件採番しました');
}

/** シンプルトリガー：編集されたシートのIDを採番 */
function onEdit(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();
  const def = [CONFIG.projects, CONFIG.tasks].find(d => d.sheetName === name);
  if (!def) return;
  withLock_(() => assignIds_(def), 5000);
}