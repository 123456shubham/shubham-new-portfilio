const SYNC_ENDPOINT = 'https://shubham-new-portfilio.vercel.app/api/sync/google-sheet';
const ENQUIRIES_SHEET = 'Enquiries';
const GMAIL_ENQUIRY_LABEL = 'Portfolio Enquiry';
const GMAIL_PROCESSED_LABEL = 'Portfolio Enquiry/Processed';
const GMAIL_ERROR_LABEL = 'Portfolio Enquiry/Error';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Portfolio Enquiries')
    .addItem('Save selected row to database', 'saveSelectedEnquiry')
    .addItem('Send thank-you email for selected row', 'sendSelectedThankYou')
    .addItem('Send quotation for selected row', 'sendSelectedQuotation')
    .addItem('Send payment receipt for selected row', 'sendSelectedPaymentReceipt')
    .addItem('Send payment reminder for selected row', 'sendSelectedPaymentReminder')
    .addSeparator().addItem('Import Gmail enquiries now', 'importGmailEnquiries')
    .addItem('Install Sheet automatic sync', 'installEnquirySyncTrigger')
    .addItem('Install Gmail enquiry automation', 'installGmailEnquiryTrigger').addToUi();
}

function syncEnquiriesToDatabase() {
  const secret = PropertiesService.getScriptProperties().getProperty('SYNC_SECRET');
  if (!secret) throw new Error('SYNC_SECRET script property is not configured.');
  const response = UrlFetchApp.fetch(SYNC_ENDPOINT, {
    method: 'post', headers: { 'X-Sync-Secret': secret }, muteHttpExceptions: true,
  });
  if (response.getResponseCode() >= 300) {
    throw new Error(`Database sync failed (${response.getResponseCode()}): ${response.getContentText()}`);
  }
  return JSON.parse(response.getContentText());
}

function onEnquiryEdit(event) {
  const range = event && event.range;
  if (!range || range.getSheet().getName() !== ENQUIRIES_SHEET || range.getRow() < 2) return;
  syncEnquiriesToDatabase();
}

function saveSelectedEnquiry() {
  requireSelectedEnquiryRow(); syncEnquiriesToDatabase();
  SpreadsheetApp.getActive().toast('Enquiry saved to the database.', 'Portfolio Enquiries');
}

function sendSelectedQuotation() { sendSelectedEmailAction('Send Quotation', 'Quotation email processed.'); }
function sendSelectedThankYou() { sendSelectedEmailAction('Send Thank You', 'Thank-you email processed.'); }
function sendSelectedPaymentReceipt() { sendSelectedEmailAction('Send Payment Receipt', 'Payment receipt email processed.'); }
function sendSelectedPaymentReminder() { sendSelectedEmailAction('Send Payment Reminder', 'Payment reminder email processed.'); }

function sendSelectedEmailAction(headerName, message) {
  const range = requireSelectedEnquiryRow();
  const column = getHeaderMap(range.getSheet())[headerName];
  if (!column) throw new Error(`${headerName} column was not found.`);
  range.getSheet().getRange(range.getRow(), column).setValue(true);
  SpreadsheetApp.flush(); syncEnquiriesToDatabase();
  SpreadsheetApp.getActive().toast(message, 'Portfolio Enquiries');
}

function installEnquirySyncTrigger() {
  deleteTriggersForFunction('onEnquiryEdit');
  ScriptApp.newTrigger('onEnquiryEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getActive().toast('Automatic Sheet sync installed.', 'Portfolio Enquiries');
}

function installGmailEnquiryTrigger() {
  [GMAIL_ENQUIRY_LABEL, GMAIL_PROCESSED_LABEL, GMAIL_ERROR_LABEL].forEach(getOrCreateGmailLabel);
  deleteTriggersForFunction('importGmailEnquiries');
  ScriptApp.newTrigger('importGmailEnquiries').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getActive().toast('Gmail automation installed. Inbox will be checked every minute.', 'Portfolio Enquiries', 8);
}

function importGmailEnquiries() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { imported: 0, skipped: 0 };
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(ENQUIRIES_SHEET);
    if (!sheet) throw new Error(`Sheet "${ENQUIRIES_SHEET}" was not found.`);
    const headers = getHeaderMap(sheet); validateRequiredHeaders(headers);
    const enquiryLabel = getOrCreateGmailLabel(GMAIL_ENQUIRY_LABEL);
    const processedLabel = getOrCreateGmailLabel(GMAIL_PROCESSED_LABEL);
    const errorLabel = getOrCreateGmailLabel(GMAIL_ERROR_LABEL);
    const properties = PropertiesService.getScriptProperties();
    const ownEmail = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
    // Scan recent inbox mail without requiring specific enquiry keywords. The
    // intent classifier below decides what is a possible business enquiry.
    const enquiryQuery = 'in:inbox newer_than:1d -category:promotions -category:social';
    const threads = uniqueThreads(
      GmailApp.search(enquiryQuery, 0, 30)
        .concat(GmailApp.search(`label:"${GMAIL_ENQUIRY_LABEL}" newer_than:90d`, 0, 30))
    );
    const pending = []; let skipped = 0;
    threads.forEach(thread => {
      const labelled = thread.getLabels().some(label => label.getName() === GMAIL_ENQUIRY_LABEL);
      thread.getMessages().forEach(message => {
        const key = messagePropertyKey(message.getId());
        if (properties.getProperty(key)) { skipped++; return; }
        const sender = parseSender(message.getFrom());
        if (!sender.email || sender.email === ownEmail || isAutomatedSender(sender.email)) {
          skipped++; return;
        }
        const subject = cleanText(message.getSubject(), 250) || 'Project Enquiry';
        const body = cleanEmailBody(message.getPlainBody() || message.getBody());
        if (!body || (!labelled && !isLikelyProjectEnquiry(subject, body))) {
          skipped++; return;
        }
        thread.addLabel(enquiryLabel);
        appendGmailEnquiryRow(sheet, headers, {
          receivedAt: message.getDate(), clientName: sender.name || nameFromEmail(sender.email),
          email: sender.email, subject, details: body, source: 'Gmail Inbox',
          autoSend: labelled,
        });
        pending.push({ key, thread });
      });
    });
    if (!pending.length) return { imported: 0, skipped };
    SpreadsheetApp.flush();
    try {
      syncEnquiriesToDatabase();
      pending.forEach(item => {
        properties.setProperty(item.key, new Date().toISOString());
        item.thread.addLabel(processedLabel); item.thread.removeLabel(errorLabel);
      });
      return { imported: pending.length, skipped };
    } catch (error) {
      pending.forEach(item => item.thread.addLabel(errorLabel)); throw error;
    }
  } finally { lock.releaseLock(); }
}

function isLikelyProjectEnquiry(subject, body) {
  const text = `${subject}\n${body}`.toLowerCase();
  const blocked = [/\botp\b/, /one[- ]time password/, /verify your email/, /password reset/,
    /unsubscribe/, /newsletter/, /job alerts?/, /urgently hiring/, /hiring for/,
    /job opening/, /job vacancy/, /apply now/, /career opportunity/, /recruiter/,
    /order shipped/, /delivery update/,
    /bank statement/, /credit card/, /loan offer/, /promotional offer/];
  if (blocked.some(pattern => pattern.test(text))) return false;
  const projects = [/mobile app/, /android app/, /ios app/, /web app/, /website/, /e-?commerce/,
    /online store/, /landing page/, /admin panel/, /dashboard/, /software/, /saas/, /crm/, /erp/,
    /api integration/, /payment gateway/, /ui\/?ux/, /redesign/, /wordpress/, /shopify/, /react/,
    /flutter/, /full stack/, /automation/, /portfolio/, /booking system/, /management system/,
    /business site/, /company site/, /web portal/, /marketplace/, /delivery app/, /customer app/,
    /vendor app/, /dealer panel/, /custom solution/, /digital product/, /online platform/];
  const intents = [/need (a|an|to|your)/, /looking for/, /want (a|an|to)/, /requirement/, /project/,
    /develop/, /development/, /build/, /create/, /design/, /quotation/, /quote/, /estimate/,
    /cost/, /price/, /budget/, /timeline/, /proposal/, /hire/, /freelanc/, /available/, /discuss/,
    /developer (chahiye|chaiye|required)/, /(app|website|software) (banana|banani|banwana|banwani|chahiye|chaiye)/,
    /(bana|develop|design) (do|karna|karwani|karwana)/, /rate (batao|share)/,
    /kitna (cost|charge|price)/, /kaise bana/, /kab tak/, /charges? kya/];
  const contacts = [/please (share|send|contact|reply|call)/, /let me know/, /can you/, /could you/,
    /whatsapp/, /phone/, /meeting/, /schedule a call/];
  let score = Math.min(projects.filter(p => p.test(text)).length, 3) * 2;
  score += Math.min(intents.filter(p => p.test(text)).length, 3) * 2;
  score += Math.min(contacts.filter(p => p.test(text)).length, 2);
  if (/\b(enquiry|inquiry)\b/.test(text)) score += 3;
  return score >= 5;
}

function appendGmailEnquiryRow(sheet, headers, enquiry) {
  const row = new Array(sheet.getLastColumn()).fill('');
  const values = {
    'Received At': enquiry.receivedAt, 'Client Name': enquiry.clientName,
    'Email Address': enquiry.email, 'Project Subject': enquiry.subject,
    'Project Details': enquiry.details, 'Payment Status': 'Pending', 'Lead Status': 'New',
    'Email Delivery': 'Pending', 'Source': enquiry.source, 'Enquiry Validity': 'Review',
    'Validation Notes': enquiry.autoSend
      ? 'Confirmed Gmail enquiry; automatic response requested'
      : 'Auto-detected from Gmail; review required before sending email',
    'Sync Status': 'Pending', 'Currency': 'INR', 'Delivery Days': 45,
    'Quotation Status': 'Pending', 'Payment Email Status': 'Pending',
    'Send Thank You': enquiry.autoSend, 'Send Quotation': enquiry.autoSend,
  };
  Object.keys(values).forEach(header => { if (headers[header]) row[headers[header] - 1] = values[header]; });
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

function requireSelectedEnquiryRow() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== ENQUIRIES_SHEET || range.getRow() < 2)
    throw new Error('Select an enquiry row in the Enquiries sheet first.');
  return range;
}

function getHeaderMap(sheet) {
  const map = {};
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .forEach((header, index) => { if (String(header).trim()) map[String(header).trim()] = index + 1; });
  return map;
}

function validateRequiredHeaders(map) {
  const required = ['Received At', 'Client Name', 'Email Address', 'Project Subject',
    'Project Details', 'Source', 'Sync Status', 'Send Thank You', 'Send Quotation'];
  const missing = required.filter(header => !map[header]);
  if (missing.length) throw new Error(`Missing Sheet headers: ${missing.join(', ')}`);
}

function parseSender(value) {
  const text = String(value || '').trim();
  const match = text.match(/([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i);
  const email = match ? match[1].toLowerCase() : '';
  return { name: cleanText(text.replace(/<[^>]+>/g, '').replace(email, '').replace(/^["']|["']$/g, ''), 150), email };
}

function isAutomatedSender(email) {
  return /(^|[._-])(no-?reply|noreply|mailer-daemon|notifications?|alerts?|newsletter)([+._-]|@)/i.test(email)
    || /(^|@)(naukri|linkedin|indeed|foundit|monster|glassdoor)\./i.test(email)
    || /naukrialerts/i.test(email);
}

function nameFromEmail(email) {
  return String(email || '').split('@')[0].replace(/[._-]+/g, ' ').trim().split(/\s+/)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ') || 'Prospective Client';
}

function cleanEmailBody(value) {
  return cleanText(String(value || '').replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n').replace(/<[^>]+>/g, ' ')
    .split(/\nOn .+wrote:\s*\n/i)[0].split(/\nFrom:\s.+\nSent:\s/i)[0], 5000);
}

function cleanText(value, maxLength) {
  return String(value || '').replace(/\r/g, '').replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n').trim().slice(0, maxLength);
}

function getOrCreateGmailLabel(name) { return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name); }
function deleteTriggersForFunction(name) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === name).forEach(t => ScriptApp.deleteTrigger(t));
}
function uniqueThreads(threads) {
  const seen = {}; return threads.filter(thread => seen[thread.getId()] ? false : (seen[thread.getId()] = true));
}
function messagePropertyKey(id) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(id));
  return 'GMAIL_IMPORTED_' + digest.map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
}
