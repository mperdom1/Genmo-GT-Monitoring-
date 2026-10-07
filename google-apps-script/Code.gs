const SPREADSHEET_ID = '1cXWQLbecAbnNvF1lp2qegvx7y8R7ktZWBZcFbCI0C7I';
const USER_TIMELINE_SHEET = 'Usertimeline';

function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();

    if (action === 'list_weeks') {
      return json({
        ok: true,
        weeks: getWeekNames_()
      });
    }

    if (action === 'get_sheet') {
      const sheetName = (e.parameter.sheet || '').trim();
      if (!sheetName) throw new Error('Falta el nombre de la hoja.');
      return json({
        ok: true,
        sheet: sheetName,
        ...readSheet_(sheetName)
      });
    }

    throw new Error('Accion GET no valida.');
  } catch (error) {
    return json({ ok: false, error: String(error.message || error) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');

    if (body.action === 'save_week') {
      saveWeek_(body.sheet, body.headers || [], body.rows || []);
      return json({ ok: true, sheet: body.sheet });
    }

    throw new Error('Accion POST no valida.');
  } catch (error) {
    return json({ ok: false, error: String(error.message || error) });
  }
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getWeekNames_() {
  return getSpreadsheet_()
    .getSheets()
    .map(sheet => sheet.getName())
    .filter(name => name !== USER_TIMELINE_SHEET);
}

function readSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('No existe la hoja: ' + sheetName);

  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return { headers: [], rows: [] };

  const headers = values[0];
  const rows = values.slice(1).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index] || '';
    });
    return item;
  });

  return { headers, rows };
}

function saveWeek_(sheetName, headers, rows) {
  if (!sheetName) throw new Error('Falta el nombre de la semana.');
  if (!headers.length) throw new Error('La semana no tiene columnas.');

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  const values = [
    headers,
    ...rows.map(row => headers.map(header => row[header] == null ? '' : String(row[header])))
  ];

  sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
