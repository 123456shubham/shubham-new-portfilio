const SYNC_ENDPOINT = 'https://shubham-new-portfilio.vercel.app/api/sync/google-sheet';

function syncEnquiriesToDatabase() {
  const secret = PropertiesService.getScriptProperties().getProperty('SYNC_SECRET');
  if (!secret) throw new Error('SYNC_SECRET script property is not configured.');
  const response = UrlFetchApp.fetch(SYNC_ENDPOINT, {
    method: 'post',
    headers: { 'X-Sync-Secret': secret },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() >= 300) {
    throw new Error(`Database sync failed (${response.getResponseCode()}): ${response.getContentText()}`);
  }
  return JSON.parse(response.getContentText());
}

function onEnquiryEdit(event) {
  const range = event && event.range;
  if (!range || range.getSheet().getName() !== 'Enquiries' || range.getRow() < 2) return;
  syncEnquiriesToDatabase();
}

function installEnquirySyncTrigger() {
  const spreadsheet = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'onEnquiryEdit')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('onEnquiryEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
}
