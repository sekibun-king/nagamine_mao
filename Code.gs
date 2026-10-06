/**
 * Google Sheets 最終更新日時の記録（軽量版・JST固定）
 *
 * 初回:
 * 1. 拡張機能 → Apps Script
 * 2. このコードを Code.gs に貼り付けて保存
 * 3. initializeMeta() を一度だけ実行
 *
 * 以後は通常どおりシートを編集するだけです。
 */

const META_SHEET_NAME = 'meta';
const SITE_TIME_ZONE = 'Asia/Tokyo';

function initializeMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = getOrCreateMetaSheet_(ss);
  const now = new Date();

  ss.getSheets().forEach(sheet => {
    if (sheet.getName() === META_SHEET_NAME) return;
    upsertMeta_(meta, sheet.getName(), sheet.getSheetId(), now);
  });
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() === META_SHEET_NAME) return;

  const meta = getOrCreateMetaSheet_(e.source);
  upsertMeta_(meta, sheet.getName(), sheet.getSheetId(), new Date());
}

function getOrCreateMetaSheet_(ss) {
  let meta = ss.getSheetByName(META_SHEET_NAME);

  if (!meta) {
    meta = ss.insertSheet(META_SHEET_NAME);
  }

  if (meta.getLastRow() === 0) {
    meta.getRange(1, 1, 1, 4).setValues([
      ['sheet_name', 'gid', 'updated', 'updated_ms']
    ]);
    meta.setFrozenRows(1);
  }

  return meta;
}

function upsertMeta_(meta, sheetName, gid, date) {
  const lastRow = meta.getLastRow();
  const rows = lastRow >= 2
    ? meta.getRange(2, 1, lastRow - 1, 2).getValues()
    : [];

  const targetName = String(sheetName).trim().toLowerCase();
  const targetGid = String(gid).trim();
  let row = -1;

  for (let i = 0; i < rows.length; i++) {
    const name = String(rows[i][0]).trim().toLowerCase();
    const id = String(rows[i][1]).trim();

    if (name === targetName || id === targetGid) {
      row = i + 2;
      break;
    }
  }

  if (row === -1) row = meta.getLastRow() + 1;

  const display = Utilities.formatDate(
    date,
    SITE_TIME_ZONE,
    'yyyy-MM-dd HH:mm:ss'
  );

  meta.getRange(row, 1, 1, 4).setValues([
    [sheetName, targetGid, display, date.getTime()]
  ]);
}
