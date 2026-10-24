/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('プロジェクト管理')
    .addItem('IDを採番', 'assignAllIds')
    .addToUi();
}
