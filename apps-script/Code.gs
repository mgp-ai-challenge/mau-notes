/**
 * Business of Apps Berlin 2026 Field Notes: backend
 * Berlin, Estrel Congress Center, 12 November 2026 · BidMachine Supply team
 *
 * Google Apps Script, container-bound to the Google Sheet
 * "Business of Apps Berlin 2026 field notes" (Extensions > Apps Script).
 *
 * Raw notes only. No AI, no API keys, no calls outside Google.
 * Typed, dictated (speech-to-text runs in the rep's browser) and pasted Slack
 * notes are saved exactly as entered. Each note is appended as a row, and a
 * ready-to-paste Slack message is returned to the browser.
 *
 * SETUP
 * 1. Run setupSheet() once to write the header row.
 * 2. Run testPipeline() to check a row is written, then delete that row.
 * 3. Deploy > New deployment > Web app
 *      Execute as: Me      Who has access: Anyone
 *    Paste the /exec URL into WEBHOOK_URL at the top of index.html.
 *
 * NOTE ON CORS: the frontend POSTs with Content-Type text/plain so the browser
 * does not send a CORS preflight (Apps Script cannot answer OPTIONS). The body
 * is still JSON and is parsed as such below. This is what lets the page read
 * the response and show the Slack message.
 */

const SHEET_NAME  = 'Notes';
const EVENT_LABEL = 'Business of Apps';

const HEADERS = ['Timestamp', 'Rep', 'Partner', 'Capture Method', 'Notes'];

// ---------------------------------------------------------------- sheet

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Prefer the 'Notes' tab; otherwise fall back to the first tab rather than
  // silently creating a second, empty one.
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

/** Run once. Renames the first tab if needed and rewrites the header row. */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sh.getName() !== SHEET_NAME) sh.setName(SHEET_NAME);
  sh.getRange(1, 1, 1, sh.getMaxColumns()).clearContent();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(5, 600);
  sh.getRange(1, 5, sh.getMaxRows(), 1).setWrap(true);
  Logger.log('Header row written to \'' + SHEET_NAME + '\': ' + HEADERS.length + ' columns');
}

// ---------------------------------------------------------------- diagnostics

/**
 * End-to-end test without the browser. Writes a real row to the sheet,
 * so delete it afterwards.
 */
function testPipeline() {
  const fake = {
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        event: 'Business of Apps Berlin 2026',
        rep: 'Nathan Boyer',
        company: 'Test Studio',
        capture_method: 'text',
        raw_content: 'Test note. About 400k US DAU on MAX, open to BidMachine as a bidder, ' +
                     'asked for a 30k integration bonus. Next step: send the MSA.'
      })
    }
  };
  Logger.log(doPost(fake).getContent());
}

// ---------------------------------------------------------------- endpoints

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** The note as a Slack message for the team channel. */
function slackMessage_(data, notes) {
  return '*' + EVENT_LABEL + ' · ' + data.company + '* · ' + data.rep + '\n\n' + notes;
}

function doGet() {
  return json_({ ok: true, service: 'Business of Apps Berlin 2026 Field Notes', status: 'up' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    let notes = (data.raw_content || '').trim();

    // Slack notes are pasted text. A bare permalink has no content in it.
    if (data.capture_method === 'slack' &&
        /^https?:\/\/[a-z0-9-]+\.slack\.com\/archives\/\S*$/i.test(notes)) {
      return json_({ ok: false,
        error: 'That is a link, not the message. Copy the message text itself and paste that.' });
    }

    // An old copy of the page may still send a recording. There is no
    // transcription any more, so say so rather than saving an empty note.
    if (data.audio_base64 && !notes) {
      return json_({ ok: false,
        error: 'Voice recordings are no longer transcribed. Reload the page, then dictate or type the note.' });
    }

    if (!notes) {
      return json_({ ok: false, error: 'Nothing to save. The note was empty.' });
    }

    sheet_().appendRow([
      data.timestamp,
      data.rep,
      data.company,
      data.capture_method,
      notes
    ]);

    return json_({ ok: true, notes: notes, slack_internal: slackMessage_(data, notes) });

  } catch (err) {
    try {
      sheet_().appendRow([new Date().toISOString(), 'ERROR', String(err.message)]);
    } catch (e2) {}
    return json_({ ok: false, error: String(err.message) });
  }
}
