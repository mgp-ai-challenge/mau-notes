/**
 * Business of Apps Berlin 2026 Field Notes: backend
 * Berlin, Estrel Congress Center, 12 November 2026 · BidMachine Supply team
 *
 * Google Apps Script, container-bound to the Google Sheet
 * "Business of Apps Berlin 2026 field notes" (Extensions > Apps Script).
 *
 * Pipeline:  voice -> Whisper -> transcript
 *            text/transcript -> Claude -> structured insight + two Slack summaries
 *            -> appended as a row, and returned to the browser as JSON
 *
 * SETUP
 * 1. Project Settings > Script Properties, add:
 *      ANTHROPIC_API_KEY
 *      OPENAI_API_KEY
 * 2. Run setupSheet() once to write the header row.
 * 3. Deploy > New deployment > Web app
 *      Execute as: Me      Who has access: Anyone
 *    Paste the /exec URL into WEBHOOK_URL at the top of index.html.
 *
 * NOTE ON CORS: the frontend POSTs with Content-Type text/plain so the browser
 * does not send a CORS preflight (Apps Script cannot answer OPTIONS). The body
 * is still JSON and is parsed as such below. This is what lets the page read
 * the response and show the Slack summary.
 */

const CLAUDE_MODEL = 'claude-sonnet-5';
const SHEET_NAME   = 'Notes';
const EVENT_LABEL  = 'Business of Apps';

const HEADERS = [
  'Timestamp', 'Rep', 'Partner', 'Capture Method', 'Partner Type',
  'Summary', 'Market Insights', 'Product Feedback', 'Commercial Feedback',
  'Next Steps', 'Slack (Internal)', 'Slack (Customer)', 'Raw Content'
];

// ---------------------------------------------------------------- keys

function props_() { return PropertiesService.getScriptProperties(); }

function anthropicKey_() {
  const k = props_().getProperty('ANTHROPIC_API_KEY');
  if (!k) throw new Error('ANTHROPIC_API_KEY not set in Script Properties');
  return k;
}

function openaiKey_() {
  const k = props_().getProperty('OPENAI_API_KEY');
  if (!k) throw new Error('OPENAI_API_KEY not set in Script Properties');
  return k;
}

// ---------------------------------------------------------------- sheet

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Prefer the 'Notes' tab; otherwise fall back to the first tab rather than
  // silently creating a second, empty one.
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

/** Run once after creating the spreadsheet. Renames the first tab and writes headers. */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sh.getName() !== SHEET_NAME) sh.setName(SHEET_NAME);
  sh.getRange(1, 1, 1, sh.getMaxColumns()).clearContent();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(7, 320);
  Logger.log('Header row written to \'' + SHEET_NAME + '\': ' + HEADERS.length + ' columns');
}

// ---------------------------------------------------------------- diagnostics

function testWhisper() {
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
    headers: { 'Authorization': 'Bearer ' + openaiKey_() },
    muteHttpExceptions: true
  });
  Logger.log('Whisper auth: ' + response.getResponseCode() + ' ' +
             response.getContentText().slice(0, 200));
}

function testClaude() {
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': anthropicKey_(), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 100,
      messages: [{ role: 'user', content: 'Say OK' }]
    }),
    muteHttpExceptions: true
  });
  Logger.log('Claude: ' + response.getContentText());
}

/** Logs the full shape of a Claude response - use this when parsing misbehaves. */
function testClaudeRaw() {
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': anthropicKey_(), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 200,
      messages: [{ role: 'user', content: 'Reply with only this JSON: {"ok": true}' }]
    }),
    muteHttpExceptions: true
  });
  Logger.log('HTTP ' + response.getResponseCode());
  Logger.log(response.getContentText());
}

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
        raw_content: 'Mid-size casual gaming studio, about 400k US DAU, on AppLovin MAX with ' +
                     'AdMob and Unity bidding. Open to adding BidMachine as a bidder in MAX but ' +
                     'want an integration bonus, asked for 30k, and worried about SDK size and ' +
                     'ANRs. Rewarded and interstitial are most of their revenue. Also spending on ' +
                     'UA, CPI target around $1.20 in the US with D7 ROAS 15%, AppsFlyer is their ' +
                     'MMP. Decision maker is their Head of Monetization. Next step: send the MSA ' +
                     'and the bonus terms, UA call next week.'
      })
    }
  };
  Logger.log(doPost(fake).getContent());
}

// ---------------------------------------------------------------- pipeline

function transcribeWithWhisper(audioBase64, mimeType) {
  const ext = (mimeType || '').indexOf('mp4') > -1 ? 'm4a'
            : (mimeType || '').indexOf('ogg') > -1 ? 'ogg' : 'webm';
  const audioBlob = Utilities.newBlob(
    Utilities.base64Decode(audioBase64),
    mimeType || 'audio/webm',
    'recording.' + ext
  );
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + openaiKey_() },
    payload: {
      file: audioBlob,
      model: 'whisper-1',
      language: 'en',
      // Vocabulary hint, so product and partner names transcribe correctly.
      prompt: 'BidMachine, AppGrowth, AppLovin MAX, Unity LevelPlay, AdMob, Google Ad Manager, ' +
              'Mintegral, Moloco, Liftoff, Pangle, AppsFlyer, Adjust, DAU, eCPM, ARPDAU, CPI, ROAS, ' +
              'MSA, rewarded, interstitial, MREC, SDK, adapter, bidder, mediation.'
    },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Transcription failed (' + response.getResponseCode() + '): ' +
                    response.getContentText().slice(0, 300));
  }
  return JSON.parse(response.getContentText()).text || '';
}

const SYSTEM_PROMPT = [
'You process field notes captured by BidMachine\'s Supply team at Business of Apps Berlin (Estrel Congress Center, 12 November 2026).',
'',
'WHO THEY ARE TALKING TO. The reps meet app publishers and developers, gaming and non-gaming, and the companies around them (mediation platforms, agencies, monetization consultants). These are not DSP or demand conversations. Every meeting is about one or both of two kinds of deal:',
'1. SUPPLY deals: the app monetizes its ad inventory through BidMachine. Products:',
'   - Indirect: the BidMachine SDK runs as a bidder inside the publisher\'s existing mediation (AppLovin MAX, Unity LevelPlay, Google AdMob and Google Ad Manager, where BidMachine is a certified SDK bidder). This is the main motion, and the team is measured on the US DAU it brings into the pipeline.',
'   - Mediation: the publisher runs BidMachine\'s own mediation platform.',
'   - BM+ (AnyLook), P2P and Direct deals.',
'2. UA deals: the app developer buys users through BidMachine\'s user acquisition offering (AppGrowth). Typical topics: CPI, CPA or ROAS targets, retention, budgets, geos, the MMP (AppsFlyer, Adjust, Singular), postbacks and unattributed data, creatives and sizes, the learning period, and campaign launch dates.',
'A single app can be a supply deal, a UA deal or both. Note both when both come up.',
'',
'Context you should use to classify accurately:',
'- Scale: the BidMachine exchange reaches over 600M daily active users across apps, with over 180 live demand partners.',
'- Supply deal stages run SQL, proposal, internal review and approval, MSA shared, signed. Say where the deal sits if the notes make it clear.',
'- Integration bonuses: publishers often ask for an upfront bonus to integrate. BidMachine pays them against a 12-month DAU commitment, after 30 days of seeing that DAU live, with a clawback, and sizes them against the expected gross profit. Any bonus ask is important: capture the amount, the apps, the DAU and the conditions discussed.',
'- The numbers that matter on supply: US and Tier 1 DAU, worldwide DAU, impressions, formats (banner, MREC, interstitial, rewarded, native, video), eCPM, ARPDAU, fill, win rate, current mediation and bidders, bundle IDs. On UA: target CPI or ROAS, daily or monthly budget, geos, platforms, KPI windows (D1, D7, D30).',
'- Common friction: SDK size, crashes and ANRs, adapter versions, latency, privacy and consent (ATT, GDPR, TCF), payment terms, reporting granularity, competition for the bidder slot.',
'- Competitors on supply: AppLovin (MAX and its exchange), Unity, Google AdMob, Mintegral, Moloco, Liftoff and Vungle, Pangle, InMobi, DT Exchange, Bigo. On UA: AppLovin, Moloco, Unity, Mintegral, Liftoff and the big self-serve networks.',
'',
'Sort everything the rep captured into exactly three buckets:',
'- market_insights: what is happening in their world - how their monetization or UA mix is shifting, eCPM and CPI trends, which networks and mediation platforms are winning or losing with them, gaming versus non-gaming dynamics, regional moves (US, EMEA, APAC, LATAM), privacy and platform changes, consolidation and studio news.',
'- product_feedback: what they need from BidMachine - the SDK and adapters, mediation features, ad formats, bidding and floors, reporting and analytics, stability and performance, and on UA: targeting, creatives, MMP integration, optimization and reporting. Include blockers and bugs.',
'- commercial_feedback: money, volume and terms - integration bonus asks, revenue share, DAU and impression volumes, eCPM or ARPDAU figures, UA budgets and CPI or ROAS targets, payment terms, contracts and MSAs, deal stage and timing, decision makers and org, and share of wallet against competitors.',
'',
'Keep every number the rep gave (DAU, bonus amounts, eCPMs, CPIs, budgets, dates) exactly as stated. A point belongs in one bucket only - pick the best fit. Leave a bucket as an empty array if the rep said nothing about it. Never invent content to fill a bucket.',
'',
'Return ONLY a JSON object, no markdown fences, with exactly these fields:',
'- summary: 2-3 sentences, plain and specific. Say whether it is a supply deal, a UA deal or both.',
'- partner_type: one of "Publisher (Supply)", "Advertiser (UA)", "Publisher (Supply + UA)", "Mediation/Adtech", "Agency", "Unknown"',
'- market_insights: array of short strings',
'- product_feedback: array of short strings',
'- commercial_feedback: array of short strings',
'- next_steps: array of short action strings, each starting with a verb, with an owner or date if mentioned',
'- slack_internal: a Slack-formatted message for the internal team channel',
'- slack_customer: a Slack-formatted recap the rep can send to the partner',
'',
'PUNCTUATION: never use an em dash or an en dash anywhere, in any field. Use a comma, a colon, a full stop, or a hyphen with a space either side. This matters: em dashes read as machine-written, and the customer summary gets sent to partners.',
'',
'SLACK FORMATTING. Use Slack mrkdwn: *bold* with single asterisks (never **), bullets as "• ", real line breaks. No markdown headers, no tables, no code fences.',
'',
'slack_internal - candid, direct, for colleagues. Shape:',
'*' + EVENT_LABEL + ' · <Partner>* · <Rep> · <Supply, UA or Supply + UA>',
'(blank line)',
'<summary>',
'(blank line)',
'*Market*',
'• ...',
'(blank line)',
'*Product*',
'• ...',
'(blank line)',
'*Commercial*',
'• ...',
'(blank line)',
'*Next steps*',
'• ...',
'Omit any section with nothing in it. Keep internal shorthand, put the key numbers (DAU, bonus ask, CPI or budget) up front, and name the risk plainly.',
'',
'slack_customer - a recap the rep can paste into a Slack Connect channel or an email to the partner. Warm, professional, first person plural ("we"). Shape:',
'*Great to meet at ' + EVENT_LABEL + ': <Partner>*',
'(blank line)',
'One or two sentences thanking them and framing the conversation.',
'(blank line)',
'*What we covered*',
'• ...',
'(blank line)',
'*What we agreed*',
'• ...',
'(blank line)',
'One closing line offering to follow up.',
'Rules for slack_customer: no internal shorthand or codenames, no gross profit, margin, revenue share or bonus figures unless the partner raised them, no candid risk language, nothing about other publishers, advertisers or competitors by name, and nothing that would embarrass the rep if forwarded. Frame their complaints as things "we are picking up internally", never as admissions of failure.'
].join('\n');

/**
 * Pull the text out of a Messages API response. The first content block is not
 * always the answer - a model may emit thinking or tool blocks first - so take
 * every text block rather than assuming content[0].
 */
function extractText_(body) {
  const blocks = (body && body.content) || [];
  const text = blocks
    .filter(function (b) { return b && b.type === 'text' && b.text; })
    .map(function (b) { return b.text; })
    .join('\n')
    .trim();

  if (!text) {
    const types = blocks.map(function (b) { return b && b.type; }).join(', ') || 'none';
    throw new Error('No text in Claude response (blocks: ' + types +
                    ', stop_reason: ' + (body && body.stop_reason) + '). ' +
                    JSON.stringify(body).slice(0, 400));
  }
  return text;
}

/** Parse JSON that may be fenced or wrapped in a sentence. Returns null if it cannot. */
function parseJsonLoose_(text) {
  const cleaned = text.replace(/```json/gi, '').replace(/```/g, '').trim();
  try { return JSON.parse(cleaned); } catch (e) {}
  const start = cleaned.indexOf('{');
  const end = cleaned.lastIndexOf('}');
  if (start > -1 && end > start) {
    try { return JSON.parse(cleaned.slice(start, end + 1)); } catch (e) {}
  }
  return null;
}

/**
 * Belt and braces on the punctuation rule in the prompt: strip any em dash or
 * spaced en dash the model slips through, from every string in the response.
 * Unspaced en dashes are left alone so date ranges survive.
 */
function stripDashes_(value) {
  if (typeof value === 'string') {
    return value
      .replace(/[ \t]*\u2014[ \t]*/g, ' - ')
      .replace(/[ \t]+\u2013[ \t]+/g, ' - ')
      .replace(/[ \t]{2,}/g, ' ');
  }
  if (Array.isArray(value)) return value.map(stripDashes_);
  if (value && typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(function (k) { out[k] = stripDashes_(value[k]); });
    return out;
  }
  return value;
}

function processWithClaude(rawContent, data) {
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': anthropicKey_(), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 2500,
      system: SYSTEM_PROMPT,
      messages: [{
        role: 'user',
        content: 'Rep: ' + data.rep + '\nPartner: ' + data.company +
                 '\nCaptured by: ' + data.capture_method +
                 (data.capture_method === 'slack'
                   ? '\n\nSOURCE NOTE: this came from Slack, not from someone talking or typing it fresh. ' +
                     'It may already be partly written up. Extract what is genuinely there and nothing more. ' +
                     'Do not restate a finished summary as though it were raw notes, and do not invent detail to fill a bucket.'
                   : '') +
                 '\n\nNotes:\n' + rawContent
      }]
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Claude failed (' + response.getResponseCode() + '): ' +
                    response.getContentText().slice(0, 300));
  }

  const body = JSON.parse(response.getContentText());
  const text = extractText_(body);
  const parsed = parseJsonLoose_(text);

  if (parsed) {
    return stripDashes_(parsed);
  } else {
    Logger.log('Claude did not return usable JSON. Raw text:\n' + text);
    return {
      summary: text,
      partner_type: 'Unknown',
      market_insights: [], product_feedback: [], commercial_feedback: [],
      next_steps: [],
      slack_internal: '*' + EVENT_LABEL + ' · ' + data.company + '* · ' + data.rep + '\n\n' + text,
      slack_customer: ''
    };
  }
}

function writeToSheet(parsed, original, rawContent) {
  sheet_().appendRow([
    original.timestamp,
    original.rep,
    original.company,
    original.capture_method,
    parsed.partner_type || 'Unknown',
    parsed.summary || '',
    (parsed.market_insights || []).join(' | '),
    (parsed.product_feedback || []).join(' | '),
    (parsed.commercial_feedback || []).join(' | '),
    (parsed.next_steps || []).join(' | '),
    parsed.slack_internal || '',
    parsed.slack_customer || '',
    rawContent
  ]);
}

// ---------------------------------------------------------------- endpoints

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, service: 'Business of Apps Berlin 2026 Field Notes', status: 'up' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    let rawContent = data.raw_content || '';

    // Slack notes are pasted text. A bare permalink has no content in it, so
    // catch that rather than handing Claude a URL to guess at.
    if (data.capture_method === 'slack' &&
        /^https?:\/\/[a-z0-9-]+\.slack\.com\/archives\/\S*$/i.test(rawContent.trim())) {
      return json_({ ok: false,
        error: 'That is a link, not the message. Copy the message text itself and paste that.' });
    }

    if (data.audio_base64) {
      const transcript = transcribeWithWhisper(data.audio_base64, data.audio_mime_type);
      rawContent = rawContent ? transcript + '\n\n' + rawContent : transcript;
    }

    if (!rawContent.trim()) {
      return json_({ ok: false, error: 'Nothing to process. The note was empty.' });
    }

    const parsed = processWithClaude(rawContent, data);
    writeToSheet(parsed, data, rawContent);

    return json_({
      ok: true,
      partner_type: parsed.partner_type || 'Unknown',
      summary: parsed.summary || '',
      slack_internal: parsed.slack_internal || '',
      slack_customer: parsed.slack_customer || ''
    });

  } catch (err) {
    try {
      sheet_().appendRow([new Date().toISOString(), 'ERROR', String(err.message)]);
    } catch (e2) {}
    return json_({ ok: false, error: String(err.message) });
  }
}
