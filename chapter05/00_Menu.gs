/**
 * プロジェクト管理ダッシュボード
 * Copyright (c) 2026 Tovarity
 * 個人および所属組織内での利用・改変は可。再配布・販売は禁止（詳細は LICENSE.md）。
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('プロジェクト管理')
    .addItem('IDを採番', 'assignAllIds')
    .addItem('未作成のプロジェクト資料を作成', 'createMissingProjectDocs')
    .addItem('未作成のタスク資料を作成', 'createMissingTaskDocs')
    .addSeparator()
    .addItem('自動作成をオンにする（初回のみ）', 'enableAutoCreate')
    .addItem('基本情報を一括同期', 'syncAllInfo')
    .addToUi();
}
