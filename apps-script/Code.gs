/**
 * R&D Rotation intake.
 *
 * Receives { key, name, image } from the GitHub Pages form and appends a row
 * to the spreadsheet, with the photo as a real in-cell image.
 *
 * Uses the spreadsheet service only, so the sole scope requested is
 * .../auth/spreadsheets. Nothing here touches your files or mail.
 */

const SHEET_ID   = '1w-MrnyfrT0Zf2QYaFCdOPifk2wII1ckiWjkH78NO5UM';
const ROW_HEIGHT = 90;
const COL_WIDTH  = 120;

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

    const sh  = sheet_();
    const row = sh.getLastRow() + 1;

    sh.getRange(row, 1).setValue(name);
    sh.getRange(row, 2).setValue(
      SpreadsheetApp.newCellImage()
        .setSourceUrl(image)
        .setAltTextTitle(name)
        .build()
    );
    sh.setRowHeight(row, ROW_HEIGHT);
    sh.setColumnWidth(2, COL_WIDTH);
    SpreadsheetApp.flush();

    return json_({ ok: true, row: row });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ------------------------------------------------------------------ *
 * One-time probe: which image mechanism does Sheets actually render?
 * Run this once to authorize the script, then look at rows 2-5.
 * ------------------------------------------------------------------ */

const HTTPS_URL = 'https://pioneer721.github.io/rnd-rotation-prototype/probe-test.png';
const TEST_B64  = 'iVBORw0KGgoAAAANSUhEUgAAAHgAAAB4CAIAAAC2BqGFAAABxElEQVR42u3QQVEAQBADQZTwRhOa0IkD/riYS7JT1Qr64+v7V4APCwB/P59GE8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8tGQ8vXo7Hl09Hk8t1oePloNL98MfrJ8rnoV8u3oh8uH4p+u3wl+vnyieiE5f3okOXx6Jzl5eio5dnotOXN6MDlwejM5bXo2OWp6OTlnejw5ZHo/OWF6Irl+uiW5e7oouXi6K7l1ui65croxuW+6NLlsuje5abo6uWa6PbljuiB5YLojeX06Jnl6Oil5dzoseXQ6L3lxOjJ5bjo1eWs6OHloOjt5ZTo+eWI6AvL76OPLD+OvrP8MvrU8rPoa8tvog8uP4i+uUxHn11Goy8vc9HHl6Fol4loi4lof4loc4loZ4loW4loT4loQ4loN4loK4loH4loE4loB4lo+4ho74ho44ho14hoy4hov4hos4hop4hom4hoj4hog4hod4hoa4hoX4hoU4hoR4hoO4hoL4hoI4hoFwD/Bd4c+CIWhk4AAAAASUVORK5CYII=';

function probeImageModes() {
  const sh = sheet_();
  const dataUrl = 'data:image/png;base64,' + TEST_B64;
  const out = [];

  function attempt(row, label, fn) {
    try { fn(); sh.getRange(row, 1).setValue(label); out.push(label + ' -> OK'); }
    catch (err) { sh.getRange(row, 1).setValue(label + ' [THREW]'); out.push(label + ' -> THREW: ' + err); }
  }

  attempt(2, 'A: CellImage(data:)', function () {
    sh.getRange(2, 2).setValue(SpreadsheetApp.newCellImage().setSourceUrl(dataUrl).build());
  });
  attempt(3, 'B: CellImage(https)', function () {
    sh.getRange(3, 2).setValue(SpreadsheetApp.newCellImage().setSourceUrl(HTTPS_URL).build());
  });
  attempt(4, 'C: =IMAGE(https)', function () {
    sh.getRange(4, 2).setFormula('=IMAGE("' + HTTPS_URL + '")');
  });
  attempt(5, 'D: =IMAGE(data:)', function () {
    sh.getRange(5, 2).setFormula('=IMAGE("' + dataUrl + '")');
  });

  sh.setRowHeights(2, 4, 70);
  sh.setColumnWidth(2, COL_WIDTH);
  const msg = out.join('\n');
  Logger.log(msg);
  return msg;
}
