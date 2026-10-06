/**
 * Last update 自動記録
 *
 * 1) Googleスプレッドシートで「拡張機能」→「Apps Script」
 * 2) このコードを Code.gs に貼り付けて保存
 * 3) initializeMeta() を一度だけ実行して権限を許可
 *
 * 以後、任意のシートを編集すると meta シートに
 * sheet_name / gid / updated / updated_ms が自動記録されます。
 */

const META_SHEET_NAME = 'meta';

function initializeMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = getOrCreateMetaSheet_(ss);
  const now = new Date();

  ss.getSheets().forEach(sheet => {
    if (sheet.getName() === META_SHEET_NAME) return;
    upsertMeta_(meta, sheet.getName(), sheet.getSheetId(), now);
  });

  upsertMeta_(meta, 'site', 'site', now);
}

function onEdit(e) {
  if (!e || !e.range) return;

  const ss = e.source;
  const sheet = e.range.getSheet();
  if (sheet.getName() === META_SHEET_NAME) return;

  const meta = getOrCreateMetaSheet_(ss);
  const now = new Date();

  upsertMeta_(meta, sheet.getName(), sheet.getSheetId(), now);
  upsertMeta_(meta, 'site', 'site', now);
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
  const lastRow = Math.max(meta.getLastRow(), 1);
  const values = lastRow >= 2
    ? meta.getRange(2, 1, lastRow - 1, 2).getValues()
    : [];

  const normalizedName = String(sheetName).trim().toLowerCase();
  const normalizedGid = String(gid).trim();
  let row = -1;

  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][0]).trim().toLowerCase();
    const id = String(values[i][1]).trim();

    if (name === normalizedName || id === normalizedGid) {
      row = i + 2;
      break;
    }
  }

  if (row === -1) row = meta.getLastRow() + 1;

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
    || 'Asia/Tokyo';

  const display = Utilities.formatDate(
    date,
    tz,
    "yyyy-MM-dd'T'HH:mm:ss"
  );

  meta.getRange(row, 1, 1, 4).setValues([
    [sheetName, String(gid), display, date.getTime()]
  ]);
}
