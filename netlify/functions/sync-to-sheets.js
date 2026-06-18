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
  'Their Pick', "Daniel's Pick", 'Final Result', 'Winner', 'Prize', 'Redeemed'
];

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
          theirPick, danielsPick, finalResult, prize } = body;

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

  // Ensure header row exists
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: 'A1:L1'
  });
  if (!headerRes.data.values || headerRes.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: 'A1:L1',
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] }
    });
  }

  if (action === 'prediction') {
    await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId,
      range: 'A:L',
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          firstName, lastName, email, restaurant, match, date || '',
          theirPick, danielsPick || '', '', 'No', prize || '', 'No'
        ]]
      }
    });

  } else if (action === 'winner') {
    // Find the existing prediction row to update it
    const allRes = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: 'A:L'
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
      // Update Final Result, Winner, Prize columns (I, J, K)
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `I${rowIndex}:K${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[finalResult, 'Yes', prize || '']] }
      });
    } else {
      // Prediction row not found — append a full winner row
      await sheets.spreadsheets.values.append({
        spreadsheetId: sheetId,
        range: 'A:L',
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            firstName, lastName, email, restaurant, match, date || '',
            theirPick, danielsPick || '', finalResult, 'Yes', prize || '', 'No'
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
