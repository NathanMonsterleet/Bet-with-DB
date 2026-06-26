const { google } = require('googleapis');

const SHEET_IDS = {
  'Le Gratin': '1Hu9YSqvryZTkldW95dZ3RzQMl4AwUdNCJei19OzTUew',
  'Restaurant Daniel': '1a_z0pCydvC-3HhmhtycZeROhv0W-Zmgad6hLek3Cze4',
  "La Tête D'Or": '1Z1kKY9hA4Gd1XeY7WkRs1GrZg8q_PNuWj_N7LpNg7iQ',
  'Café Boulud at Maison Barnes': '1eI7MnvTm_tEmPc9rDYUhZpPbt0kqyb4spsM5zkBHkUM',
  'Le Pavillon': '1R4WIvTGI6PbFtlLrQymus9maqd2N5UoAOMDxv75TNjA'
};

const SERVICE_ACCOUNT_EMAIL = 'bet-with-daniel-boulud@bamboo-diode-499817-r2.iam.gserviceaccount.com';

const HEADERS = [
  'First Name', 'Last Name', 'Email', 'Restaurant', 'Match', 'Date',
  'Their Pick', "Daniel's Pick", 'Final Result',
  'Real Winner', 'Selected Winner', 'Prize', 'Redeemed'
];

const yn = (b) => (b ? 'Yes' : 'No');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  let body;
  try {
    body = JSON.parse(event.body);
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid JSON' }) };
  }

  const { action, restaurant, firstName, lastName, email, match, date,
          theirPick, danielsPick, finalResult, prize,
          isRealWinner, isSelectedWinner } = body;

  const sheetId = SHEET_IDS[restaurant];
  if (!sheetId) {
    return { statusCode: 400, body: JSON.stringify({ error: `Unknown restaurant: ${restaurant}` }) };
  }

  const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n');
  if (!privateKey) {
    return { statusCode: 500, body: JSON.stringify({ error: 'GOOGLE_PRIVATE_KEY env var not configured' }) };
  }

  const auth = new google.auth.JWT(
    SERVICE_ACCOUNT_EMAIL,
    null,
    privateKey,
    ['https://www.googleapis.com/auth/spreadsheets']
  );
  const sheets = google.sheets({ version: 'v4', auth });

  await ensureHeaders(sheets, sheetId);

  if (action === 'winner') {
    // Locate the first sheet/tab id for any cell updates we may need
    const allRes = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'A:M'
    });

    const rows = allRes.data.values || [];
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      // Match by email (col C = index 2) and match label (col E = index 4)
      if (rows[i][2] === email && rows[i][4] === match) {
        rowIndex = i + 1; // Sheets rows are 1-indexed
        break;
      }
    }

    if (rowIndex !== -1) {
      // Update the Real Winner (col J) and Selected Winner (col K) cells in place.
      // Preserve everything else (notably Redeemed in col M).
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `J${rowIndex}:K${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[yn(isRealWinner), yn(isSelectedWinner)]] }
      });
    } else {
      await sheets.spreadsheets.values.append({
        spreadsheetId: sheetId,
        range: 'A:M',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            firstName, lastName, email, restaurant, match, date || '',
            theirPick, danielsPick || '', finalResult || '',
            yn(isRealWinner), yn(isSelectedWinner), prize || '', 'No'
          ]]
        }
      });
    }

  } else {
    return { statusCode: 400, body: JSON.stringify({ error: `Unknown action: ${action}` }) };
  }

  return {
    statusCode: 200,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ok: true })
  };
};

// Make sure the sheet has the 13-col header. If it's currently in the old 12-col
// layout (single "Winner" column at J), insert a blank column at K so existing
// Prize/Redeemed data shifts right, then write the new header row.
async function ensureHeaders(sheets, spreadsheetId) {
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: 'A1:M1'
  });
  const current = (headerRes.data.values && headerRes.data.values[0]) || [];

  // Fresh / empty sheet — just write the header.
  if (current.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'A1:M1',
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] }
    });
    return;
  }

  // Old layout: 12 cols, J is "Winner". Insert a new K column to make room.
  const isOldLayout = current.length <= 12 &&
                      (current[9] || '').toString().trim().toLowerCase() === 'winner' &&
                      (current[10] || '').toString().trim().toLowerCase() !== 'selected winner';

  if (isOldLayout) {
    // Find the first sheet's gid (usually 0, but resolve it for safety).
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const firstSheet = (meta.data.sheets || [])[0];
    const gid = firstSheet?.properties?.sheetId ?? 0;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{
          insertDimension: {
            range: { sheetId: gid, dimension: 'COLUMNS', startIndex: 10, endIndex: 11 },
            inheritFromBefore: false
          }
        }]
      }
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'A1:M1',
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] }
    });
    return;
  }

  // Already 13-col or some other shape — make sure header text is current.
  const needsRewrite = HEADERS.some((h, i) => (current[i] || '') !== h);
  if (needsRewrite) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: 'A1:M1',
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] }
    });
  }
}
