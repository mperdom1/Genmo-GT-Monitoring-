export type GoogleSheetRow = Record<string, string | number | boolean | null | undefined>;

type GoogleSheetResponse = {
  ok: boolean;
  sheet?: string;
  headers?: string[];
  rows?: GoogleSheetRow[];
  weeks?: string[];
  error?: string;
};

const GOOGLE_SHEETS_API_URL = import.meta.env.VITE_GOOGLE_SHEETS_API_URL || '';

function apiUrl() {
  if (!GOOGLE_SHEETS_API_URL) {
    throw new Error(
      'Falta configurar VITE_GOOGLE_SHEETS_API_URL. Es la URL del Google Apps Script que conecta Genmo con tu Google Sheet.'
    );
  }
  return GOOGLE_SHEETS_API_URL;
}

async function parseResponse(response: Response): Promise<GoogleSheetResponse> {
  const text = await response.text();
  let data: GoogleSheetResponse;
  try {
    data = JSON.parse(text);
  } catch {
    throw new Error('Google Sheets devolvio una respuesta que no es JSON. Revisa el Web App de Apps Script.');
  }
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || `Google Sheets respondio con HTTP ${response.status}`);
  }
  return data;
}

export async function fetchGoogleSheetData(sheetName: string) {
  const response = await fetch(`${apiUrl()}?action=get_sheet&sheet=${encodeURIComponent(sheetName)}`, {
    method: 'GET',
    headers: { Accept: 'application/json' },
  });
  const data = await parseResponse(response);
  return {
    headers: data.headers || [],
    rows: data.rows || [],
  };
}

export async function listGoogleSheetWeeks() {
  const response = await fetch(`${apiUrl()}?action=list_weeks`, {
    method: 'GET',
    headers: { Accept: 'application/json' },
  });
  const data = await parseResponse(response);
  return data.weeks || [];
}

export async function saveGoogleSheetWeek(sheetName: string, rows: GoogleSheetRow[]) {
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  const response = await fetch(apiUrl(), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({
      action: 'save_week',
      sheet: sheetName,
      headers,
      rows,
    }),
  });
  return parseResponse(response);
}
