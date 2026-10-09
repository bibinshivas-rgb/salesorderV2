/**
 * Apps Script additions for "Link existing DEL ticket".
 * Paste into the existing web-app project (the one serving APPS_SCRIPT_URL), then redeploy
 * (Deploy > Manage deployments > edit > New version).
 *
 * 1. In your existing doGet(e), add before the fallthrough:
 *      if (e.parameter.action === 'links') return json_({ ok: true, links: getLinks_() });
 *
 * 2. In your existing doPost(e), where it switches on the parsed body's `action`, add:
 *      if (body.action === 'linkDEL') return json_(linkDel_(body));
 *
 * 3. jiraFetch_ below is a placeholder: reuse whatever helper createDEL already uses to call Jira
 *    (same credentials), or adapt this one.
 *
 * Storage: a "DEL_Links" sheet in the same spreadsheet, columns:
 *   OpportunityId | DEL Key | Deal Name | Linked By | Linked At
 */

var LINK_SHEET = 'DEL_Links';
var JIRA_PO_FIELD = 'customfield_10320';      // PO number
var JIRA_PO_DATE_FIELD = 'customfield_12140'; // PO Date

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function linkSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LINK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LINK_SHEET);
    sh.appendRow(['OpportunityId', 'DEL Key', 'Deal Name', 'Linked By', 'Linked At']);
  }
  return sh;
}

// { oppId: ['DEL-1', 'DEL-2'] }
function getLinks_() {
  var rows = linkSheet_().getDataRange().getValues().slice(1);
  var out = {};
  rows.forEach(function (r) {
    if (!r[0] || !r[1]) return;
    (out[r[0]] = out[r[0]] || []).push(String(r[1]));
  });
  return out;
}

function linkDel_(b) {
  if (!b.key || !/^DEL-\d+$/.test(b.key)) return { ok: false, error: 'Invalid DEL key' };
  if (!b.opportunityId) return { ok: false, error: 'Missing opportunityId' };

  var sh = linkSheet_();
  var rows = sh.getDataRange().getValues().slice(1);
  var exists = rows.some(function (r) { return r[0] === b.opportunityId && r[1] === b.key; });
  if (!exists) {
    sh.appendRow([b.opportunityId, b.key, b.dealName || '', Session.getActiveUser().getEmail(), new Date()]);
  }

  // Back-fill PO on the Jira ticket only when it is blank / "Not Received", so the PO match
  // also works from Jira's side and nothing already entered is overwritten.
  if (b.poNumber) {
    try {
      var issue = jiraFetch_('/rest/api/3/issue/' + b.key + '?fields=' + JIRA_PO_FIELD + ',' + JIRA_PO_DATE_FIELD, 'get');
      var cur = String((issue.fields || {})[JIRA_PO_FIELD] || '').trim().toLowerCase();
      var fields = {};
      if (!cur || cur === 'na' || cur.indexOf('not') === 0) fields[JIRA_PO_FIELD] = b.poNumber;
      if (b.poDate && !(issue.fields || {})[JIRA_PO_DATE_FIELD]) fields[JIRA_PO_DATE_FIELD] = b.poDate;
      if (Object.keys(fields).length) jiraFetch_('/rest/api/3/issue/' + b.key, 'put', { fields: fields });
    } catch (err) {
      // Link is saved even if the Jira back-fill fails.
      return { ok: true, warning: 'Linked, but Jira PO back-fill failed: ' + err.message };
    }
  }
  return { ok: true };
}

// PLACEHOLDER: replace with the helper createDEL already uses.
function jiraFetch_(path, method, payload) {
  var props = PropertiesService.getScriptProperties();
  var auth = Utilities.base64Encode(props.getProperty('JIRA_EMAIL') + ':' + props.getProperty('JIRA_TOKEN'));
  var opts = { method: method, headers: { Authorization: 'Basic ' + auth, 'Content-Type': 'application/json' }, muteHttpExceptions: true };
  if (payload) opts.payload = JSON.stringify(payload);
  var res = UrlFetchApp.fetch('https://ati-motors.atlassian.net' + path, opts);
  if (res.getResponseCode() >= 300) throw new Error('Jira ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  var t = res.getContentText();
  return t ? JSON.parse(t) : {};
}
