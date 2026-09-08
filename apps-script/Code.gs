/**
 * R&D Rotation intake.
 *
 * Receives { key, name, image, aspect, pages } from the GitHub Pages form and
 * appends a row to the spreadsheet, with the photo as a real in-cell image.
 *
 * A PDF is rasterised in the browser before it gets here: every page is drawn
 * side by side into one strip, because a cell holds exactly one image.
 *
 * Uses the spreadsheet service only, so the sole scope requested is
 * .../auth/spreadsheets. Nothing here touches your files or mail.
 *
 * The photo is embedded as an inline data: URL via newCellImage(). Verified
 * 2026-09-08 against this sheet: CellImage(data:) and CellImage(https) both
 * render, while =IMAGE(https) gives #REF! (external fetches need per-sheet
 * permission) and =IMAGE(data:) gives #VALUE!. The data: form therefore needs
 * no file hosting and no publicly readable image URLs.
 */

const SHEET_ID   = '1w-MrnyfrT0Zf2QYaFCdOPifk2wII1ckiWjkH78NO5UM';
const ROW_HEIGHT = 90;
const COL_WIDTH  = 120;   // minimum width of the photo column
const MAX_COL_W  = 480;   // ceiling, so one long PDF cannot swallow the sheet

// Speed bump only, NOT a secret: the calling page is public, so anyone reading
// its source can see this value. It deters drive-by posts, nothing more.
const SHARED_KEY = 'rnd-rotation-2026';

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const byGid = ss.getSheets().filter(function (s) { return s.getSheetId() === 0; });
  return byGid.length ? byGid[0] : ss.getSheets()[0];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Health check - lets us confirm the deployment is reachable. */
function doGet() {
  return json_({ ok: true, service: 'rnd-rotation' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: 'empty request body' });
    }

    const body = JSON.parse(e.postData.contents);
    if (body.key !== SHARED_KEY) return json_({ ok: false, error: 'unauthorized' });

    const name  = String(body.name || '').trim();
    const image = String(body.image || '');

    // Both are mandatory - the form enforces this too, but never trust the client.
    if (!name) return json_({ ok: false, error: 'name is required' });
    if (!/^data:image\/[a-z0-9+.-]+;base64,/i.test(image)) {
      return json_({ ok: false, error: 'image is required' });
    }

    // Presentation hints. Treat them as untrusted and fall back to sane values.
    const aspect = Number(body.aspect) > 0 ? Number(body.aspect) : 1;
    const pages  = Math.max(1, Math.floor(Number(body.pages) || 1));

    const sh  = sheet_();
    const row = sh.getLastRow() + 1;

    sh.getRange(row, 1).setValue(name);
    sh.getRange(row, 2).setValue(
      SpreadsheetApp.newCellImage()
        .setSourceUrl(image)
        .setAltTextTitle(pages > 1 ? name + ' (' + pages + ' pages)' : name)
        .build()
    );
    sh.setRowHeight(row, ROW_HEIGHT);

    // A multi-page strip is far wider than it is tall, and an in-cell image is
    // scaled to fit its cell - so without this the pages become unreadable.
    // Grow the column to suit the widest entry, but never shrink it back and
    // never past the ceiling.
    const wanted = Math.min(MAX_COL_W, Math.max(COL_WIDTH, Math.round(ROW_HEIGHT * aspect)));
    if (wanted > sh.getColumnWidth(2)) sh.setColumnWidth(2, wanted);

    SpreadsheetApp.flush();

    return json_({ ok: true, row: row });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}
