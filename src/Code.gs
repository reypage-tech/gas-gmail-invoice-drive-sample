/**
 * Gmail invoice filing sample (dummy data only).
 * One flow: labeled invoice emails -> Invoice log sheet -> attachments filed in Drive
 * (Root / Vendor / YYYY-MM) -> price check against the Price list tab.
 */
var TABS = { CONFIG: 'Config', PRICES: 'Price list', LOG: 'Invoice log', MISMATCH: 'Mismatch log', RUN: 'Run log' };
var LOG_HEADER = ['Received', 'Vendor', 'Invoice no', 'Item', 'Qty', 'Unit price', 'Expected price', 'Check', 'File link', 'Message id'];
var MISMATCH_HEADER = ['Checked at', 'Vendor', 'Invoice no', 'Item', 'Invoice price', 'Expected price', 'Difference', 'Message id'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invoices').addItem('Process new invoices', 'processInvoices').addToUi();
}

function readConfig_(ss) {
  var rows = ss.getSheetByName(TABS.CONFIG).getDataRange().getValues();
  var cfg = {};
  for (var i = 1; i < rows.length; i++) if (rows[i][0]) cfg[String(rows[i][0]).trim()] = String(rows[i][1]).trim();
  ['search_query', 'processed_label', 'drive_root_folder_id', 'tolerance_pct'].forEach(function (k) {
    if (!(k in cfg)) throw new Error('Config is missing key: ' + k);
  });
  return cfg;
}

function readPrices_(ss) {
  var rows = ss.getSheetByName(TABS.PRICES).getDataRange().getValues();
  var map = {};
  for (var i = 1; i < rows.length; i++) if (rows[i][0]) map[key_(rows[i][0], rows[i][1])] = Number(rows[i][2]);
  return map;
}

function key_(vendor, item) { return String(vendor).trim().toLowerCase() + '|' + String(item).trim().toLowerCase(); }

function vendorFromSender_(from) {
  var m = String(from).match(/@([^.>]+)\./);
  return m ? m[1].charAt(0).toUpperCase() + m[1].slice(1) : 'Unknown';
}

/** Lines like "Item: Copper pipe 15mm | Qty: 10 | Unit price: 12.50" */
function parseLines_(body) {
  var out = [], re = /Item:\s*([^|\n]+?)\s*\|\s*Qty:\s*([\d.]+)\s*\|\s*Unit price:\s*([\d.,]+)/gi, m;
  while ((m = re.exec(body)) !== null) out.push({ item: m[1], qty: Number(m[2]), price: Number(m[3].replace(/,/g, '')) });
  return out;
}

function parseInvoiceNo_(subject, body) {
  var m = (subject + '\n' + body).match(/Invoice\s*(?:no\.?|#)\s*([A-Z0-9-]+)/i);
  return m ? m[1] : '';
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function sheetWithHeader_(ss, name, header) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(header);
  return sh;
}

function processedIds_(logSheet) {
  var rows = logSheet.getDataRange().getValues(), set = {};
  for (var i = 1; i < rows.length; i++) set[rows[i][9]] = true;
  return set;
}

function processInvoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var started = new Date();
  var cfg = readConfig_(ss);
  var prices = readPrices_(ss);
  var tol = Number(cfg.tolerance_pct) / 100;
  var logSh = sheetWithHeader_(ss, TABS.LOG, LOG_HEADER);
  var misSh = sheetWithHeader_(ss, TABS.MISMATCH, MISMATCH_HEADER);
  var runSh = sheetWithHeader_(ss, TABS.RUN, ['Started', 'Emails', 'Rows', 'Files', 'Mismatches', 'Skipped', 'Note']);
  var done = processedIds_(logSh);
  var label = GmailApp.getUserLabelByName(cfg.processed_label) || GmailApp.createLabel(cfg.processed_label);
  var root = DriveApp.getFolderById(cfg.drive_root_folder_id);
  var maxThreads = Number(cfg.max_threads_per_run || 20);
  var threads = GmailApp.search(cfg.search_query + ' -label:"' + cfg.processed_label + '"', 0, maxThreads);
  var n = { emails: 0, rows: 0, files: 0, mismatches: 0, skipped: 0 }, notes = [];

  threads.forEach(function (thread) {
    thread.getMessages().forEach(function (msg) {
      var id = msg.getId();
      if (done[id]) { n.skipped++; return; }
      var vendor = vendorFromSender_(msg.getFrom());
      var body = msg.getPlainBody();
      var invNo = parseInvoiceNo_(msg.getSubject(), body);
      var month = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy-MM');
      var folder = getOrCreateFolder_(getOrCreateFolder_(root, vendor), month);
      var link = '';
      msg.getAttachments().forEach(function (att) {
        var f = folder.createFile(att.copyBlob()).setName(vendor + '_' + (invNo || id) + '_' + att.getName());
        link = link || f.getUrl();
        n.files++;
      });
      var lines = parseLines_(body);
      if (lines.length === 0) notes.push('No item lines found: ' + id);
      lines.forEach(function (ln) {
        var expected = prices[key_(vendor, ln.item)];
        var check;
        if (expected === undefined) check = 'NOT IN PRICE LIST';
        else if (Math.abs(ln.price - expected) > expected * tol) check = 'MISMATCH';
        else check = 'OK';
        logSh.appendRow([msg.getDate(), vendor, invNo, ln.item, ln.qty, ln.price, expected === undefined ? '' : expected, check, link, id]);
        n.rows++;
        if (check === 'MISMATCH') {
          misSh.appendRow([started, vendor, invNo, ln.item, ln.price, expected, Math.round((ln.price - expected) * 100) / 100, id]);
          n.mismatches++;
        }
      });
      n.emails++;
    });
    thread.addLabel(label);
  });
  runSh.appendRow([started, n.emails, n.rows, n.files, n.mismatches, n.skipped, notes.join('; ')]);
  return n;
}
