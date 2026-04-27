// ═══════════════════════════════════════════════════════════════
//  Nucleous — COMPLETE GOOGLE APPS SCRIPT BACKEND
//  AfterResult Solutions | code.js v3.0
//  All sheets, dropdowns, chat saving, channels, workflows,
//  quotations, tickets, performance tracking
// ═══════════════════════════════════════════════════════════════

const CONFIG = {
  OPENROUTER_API_KEY: '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',
  AI_MODEL: 'meta-llama/llama-3.3-70b-instruct',
  COMPANY: 'AfterResult Solutions',
  APP_NAME: 'Nucleous',
  ADMIN_EMAIL: 'info.afterresult@gmail.com',
  HOT_LEAD_ALERT_EMAIL: 'info.afterresult@gmail.com',
  TZ: Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Kolkata'
};

// ─── SHEET NAMES ────────────────────────────────────────────────
const SN = {
  AUTH:         'AUTHORIZED_USERS',
  LEADS:        'MASTER_LEADS',
  EMPLOYEES:    'EMPLOYEES',
  CHAT:         'CHAT_MESSAGES',
  CHANNELS:     'CHANNELS',
  WORKFLOWS:    'WORKFLOWS',
  QUOTATIONS:   'QUOTATIONS',
  MEETINGS:     'MEETINGS',
  TICKETS:      'TICKETS',
  NOTES:        'NOTES',
  ACTIVITY:     'ACTIVITY_LOG',
  ROUTING:      'ROUTING_RULES',
  COMPANY:      'COMPANY_PROFILE',
  SERVICES:     'SERVICES_CATALOG',
  PROFILES:     'COMPANY_PROFILES',
  // Legacy aliases — keep all existing _getSheet(SN.LOGS/SHIFTS/PERFORMANCE) calls working
  LOGS:         'ACTIVITY_LOG',
  SHIFTS:       'ACTIVITY_LOG',
  PERFORMANCE:  'ACTIVITY_LOG'
};

// ─── DROPDOWN LISTS ─────────────────────────────────────────────
const DD = {
  LEAD_STATUS:       ['New','Contacted','In Progress','Qualified','Proposal Sent','Negotiating','Sold','Closed Lost','On Hold','Archived','Junk'],
  FUNNEL_STAGE:      ['New Lead','Initial Contact','Interested','Demo Scheduled','Proposal Sent','Negotiating','Qualified','Closed Won','Closed Lost'],
  PRIORITY:          ['Low','Medium','High','Critical'],
  INDUSTRY:          ['Technology','Real Estate','Financial','Healthcare','E-Commerce','Education','Manufacturing','SaaS','Agriculture','Sports','FMCG','Hospitality','Logistics','Automobile','Legal','Media','Retail','Other'],
  SOURCE:            ['LinkedIn','WhatsApp','Referral','Cold Call','Meta Ads','Google Ads','Email Campaign','Website','Event','Direct','Instagram','Other'],
  HOT:               ['Yes','No'],
  QUOTATION_STATUS:  ['Draft','Sent','Under Review','Accepted','Rejected','Won','Lost','Expired','Revised'],
  SERVICE:           ['Social Media Management (Organic)','Social Media + Paid Ads','Instagram Growth Plan','LinkedIn Management','Google Ads Management','Meta Ads Management','Amazon Marketplace Management','Flipkart Marketplace Management','SEO – Search Engine Optimization','Website Design & Development','Lead Generation (Leadin)','Brand Identity & Design','Email Marketing Campaigns','Performance Marketing (Full Stack)','E-Commerce Store Setup','OPM – Offline Presence Management','Custom Package'],
  DURATION:          ['1 Month','3 Months','6 Months','12 Months','18 Months','24 Months','Ongoing'],
  MEETING_STATUS:    ['Scheduled','Confirmed','Completed','Cancelled','Rescheduled','No Show'],
  TICKET_STATUS:     ['Open','In Progress','Resolved','Closed','Reopened','Escalated','On Hold'],
  TICKET_PRIORITY:   ['Low','Medium','High','Critical','Urgent'],
  TICKET_CATEGORY:   ['Pricing Query','Technical Issue','Lead Quality','Data Error','Feature Request','Urgent Support','Compliance','Billing','Access Issue','Process Issue','Other'],
  WORKFLOW_STATUS:   ['Pending','In Progress','Completed','Cancelled','On Hold','Rejected'],
  WORKFLOW_TEAM:     ['Meeting Team','Sales Closing Team','Technical Team','Accounts Team','Digital Marketing Team','Support Team','Management'],
  WORKFLOW_TYPE:     ['Forward Lead','Schedule Meeting','Escalate Issue','Quotation Request','Technical Request','Approval Request','Other'],
  EMP_STATUS:        ['Active','Inactive','On Leave','Deactivated','Hold','Suspended','Probation'],
  EMP_ROLE:          ['Junior BDE','BDE','Senior BDE','Lead BDE','BDE Manager','Admin','Intern','Trainer'],
  EMP_EXP:           ['Fresher (0-6 mo)','Junior (0-1 yr)','Mid (1-3 yr)','Senior (3-5 yr)','Expert (5+ yr)'],
  CALL_OUTCOME:      ['Connected','Not Answered','Busy','Call Back Later','Interested','Not Interested','Wrong Number','Voicemail','Follow Up','Converted'],
  MSG_TYPE:          ['text','workflow','system','announcement','file','quotation'],
  CHANNEL_STATUS:    ['Active','Archived','Read Only','Muted'],
  CHANNEL_TYPE:      ['Team Channel','Direct Message','Workflow','Announcement','Support'],
  NOTE_TYPE:         ['General','Call Note','Email Note','Meeting Note','Objection','Follow Up','Escalation','Closure'],
  SHIFT_STATUS:      ['Active','On Break','Ended','Absent','Half Day','Work From Home'],
  AUTH_ROLE:         ['Admin','Manager','Senior BDE','BDE','Junior BDE','Viewer']
};

// ═══════════════════════════════════════════════════════════════
//  ENTRY POINT
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  // ── API CALLS ─────────────────────────────
  if (params.fn) {
    return _handleApiCall(params.fn, params.args);
  }

  // ── ADMIN PANEL ───────────────────────────
  if (params.page === 'admin' || params.admin === 'true') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('Nucleous Admin Control Centre')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── MARKETING PANEL ───────────────────────
  if (params.page === 'marketing') {
    return HtmlService.createHtmlOutputFromFile('marketing-panel')
      .setTitle('Marketing Panel')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── ONBOARDING ────────────────────────────
  if (params.page === 'onboarding') {
    return HtmlService.createHtmlOutputFromFile('Onboarding')
      .setTitle('Onboarding')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── QUOTATION VIEW ────────────────────────
  if (params.quote) {
    const quoteId = params.quote;
    let q = null;

    try {
      const stored = PropertiesService.getScriptProperties().getProperty('qt_' + quoteId);
      if (stored) q = JSON.parse(stored);
    } catch (e) {}

    if (!q) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('QUOTATIONS');
        if (sheet) {
          const data = sheet.getDataRange().getValues();
          const headers = data[0];
          for (let i = 1; i < data.length; i++) {
            const row = {};
            headers.forEach((h, idx) => row[h] = data[i][idx]);
            if (row.quoteNum === quoteId || row.ID === quoteId) {
              q = row;
              break;
            }
          }
        }
      } catch (e) {}
    }

    if (!q) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif;padding:40px;text-align:center"><h2>Quotation not found</h2></div>'
      );
    }

    return HtmlService.createHtmlOutput(buildQuotationHTML(q))
      .setTitle('Quotation ' + quoteId)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── DEFAULT APP ───────────────────────────
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Nucleous by AfterResult')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    let body = {};
    try { body = JSON.parse(e.postData.contents); } catch(x) {}
    const fn   = body.fn   || (e.parameter && e.parameter.fn)   || '';
    const args = body.args || (e.parameter && e.parameter.args ? JSON.parse(e.parameter.args) : []);
    return _routeApiCall(fn, args);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function _handleApiCall(fn, argsRaw) {
  try {
    let args = [];
    if (argsRaw) {
      try { args = JSON.parse(decodeURIComponent(argsRaw)); } catch(x) {
        try { args = JSON.parse(argsRaw); } catch(y) {}
      }
    }
    return _routeApiCall(fn, args);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function _routeApiCall(fn, args) {
  let result;
  try {
    switch (fn) {
      case 'getLeads':              result = getLeads();                          break;
      case 'addLead':               result = addLead(args[0]);                    break;
      case 'updateLead':            result = updateLead(args[0]);                 break;
      case 'archiveLead':           result = archiveLead(args[0]);                break;
      case 'bulkDeleteLeads':       result = bulkDeleteLeads(args[0]);            break;
      case 'tagHotLead':            result = tagHotLead(args[0], args[1]);        break;
      case 'getAllEmployees':        result = getAllEmployees();                    break;
      case 'saveEmployee':          result = saveEmployee(args[0]);               break;
      case 'assignLeadsByName': result = assignLeadsByName(args[0], args[1], args[2], args[3], args[4]); break;
case 'assignLeadsByEmail': result = assignLeadsByEmail(args[0], args[1], args[2], args[3], args[4]); break;
case 'bulkAssignLeadsByEmail': result = bulkAssignLeadsByEmail(args[0], args[1], args[2]); break;
case 'resetEmployeeLeads': result = resetEmployeeLeads(args[0], args[1], args[2], args[3]); break;
case 'savePipelineConfig': result = savePipelineConfig(args[0]); break;
      case 'assignLeads':           result = assignLeads(args[0], args[1], args[2]); break;
      case 'checkUserAuth':         result = checkUserAuth(args[0]);              break;
      case 'getTickets':            result = getTickets(args[0]);                 break;
      case 'raiseTicket':           result = raiseTicket(args[0]);                break;
      case 'resolveTicket':         result = resolveTicket(args[0], args[1]);     break;
      case 'updateTicketStatus':    result = updateTicketStatus(args[0], args[1], args[2]); break;
      case 'getChat':               result = getChat(args[0], args[1], args[2]);  break;
      case 'sendChat':              result = sendChat(args[0], args[1], args[2], args[3], args[4]); break;
      case 'writeChatMessage':      result = writeChatMessage(args[0]);           break;
      case 'markMessagesRead':      result = markMessagesRead(args[0], args[1]);  break;
      case 'getChannels':           result = getChannels();                        break;
      case 'addChannel':            result = addChannel(args[0]);                 break;
      case 'saveWorkflow':          result = saveWorkflow(args[0]);               break;
      case 'getWorkflows':          result = getWorkflows(args[0]);               break;
      case 'updateWorkflowStatus':  result = updateWorkflowStatus(args[0], args[1], args[2]); break;
      case 'saveQuotation':         result = saveQuotation(args[0]);              break;
      case 'updateQuotationStatus': result = updateQuotationStatus(args[0], args[1], args[2]); break;
      case 'getQuotations':         result = getQuotations(args[0]);              break;
      case 'storeQuotationAndGetUrl': result = storeQuotationAndGetUrl(args[0]);  break;
      case 'saveQuotationToSheet':  result = saveQuotationToSheet(args[0]);       break;
      case 'scheduleMeeting':       result = scheduleMeeting(args[0]);            break;
      case 'getMeetings':           result = getMeetings(args[0]);                break;
      case 'updateMeetingStatus':   result = updateMeetingStatus(args[0], args[1], args[2]); break;
      case 'addNote':               result = addNote(args[0], args[1]);           break;
      case 'getNotes':              result = getNotes(args[0]);                   break;
      case 'addCallNote':           result = addCallNote(args[0], args[1], args[2]); break;
      case 'logCall':               result = logCall(args[0], args[1], args[2]);  break;
      case 'getCallLogs':           result = getCallLogs(args[0]);                break;
      case 'startShift':            result = startShift();                         break;
      case 'endShift':              result = endShift();                           break;
      case 'startBreak':            result = startBreak();                         break;
      case 'endBreak':              result = endBreak();                           break;
case 'logShift': result = logShift(args[0], args[1], args[2], args[3]); break;
      case 'pingActivity':          result = pingActivity();                        break;
      case 'getPerformanceSummary': result = getPerformanceSummary(args[0]);      break;
      case 'initializeSheets':      result = initializeSheets();                   break;
      case 'getSheetUrl':           result = getSheetUrl();                        break;
      case 'getDropdownOptions':    result = getDropdownOptions();                 break;
      case 'getFAQs':               result = getFAQs();                            break;
      case 'getZiraSystemPrompt':   result = getZiraSystemPrompt(args[0]);        break;
case 'sendAdminNotificationEmail': result = sendAdminNotificationEmail(args[0], args[1], args[2]); break;
      case 'sendEmailFromPanel':         result = sendEmailFromPanel(args[0]);                            break;
      case 'sendEmailComposer':          result = sendEmailComposer(args[0]);                             break;
      case 'sendCompanyEmail':           result = sendEmailFromCompany(args[0], args[1], args[2]);        break;
      case 'getRecentEmails':            result = getRecentEmails(args[0]);                               break;
      case 'getDeletedLeads':            result = getDeletedLeads();                                      break;
      case 'deleteLead':                 result = deleteLead(args[0]);                                    break;
      case 'setLeadFilter':              result = setLeadFilter(args[0],args[1],args[2],args[3]);         break;
      case 'removeLeadFilter':           result = removeLeadFilter(args[0]);                              break;
      case 'getLeadFilters':             result = getLeadFilters();                                       break;
case 'saveRoutingRule':            result = saveRoutingRule(args[0]);                               break;
case 'getRoutingRules':            result = getRoutingRules();                                      break;
case 'deleteRoutingRule':          result = deleteRoutingRule(args[0]);                             break;
case 'toggleRoutingRule':          result = toggleRoutingRuleServer(args[0], args[1]);              break;
      case 'sendPerformanceReport': result = sendPerformanceReport();              break;
      case 'saveCompany':           result = saveCompany(args[0]);                 break;
      case 'getAllCompanies':        result = getAllCompanies();                    break;
      case 'deleteAdminAccount':    result = deleteAdminAccount(args[0]);          break;
      case 'getAdminConfig':        result = getAdminConfig(args[0]);              break;
      case 'saveAdminConfig':       result = saveAdminConfig(args[0]);             break;
     case 'getLeadActivity':         result = getLeadActivity(args[0]);                   break;
      case 'logUniversalActivity':    result = logUniversalActivity(args[0]);              break;
      case 'saveUserSettings':        result = saveUserSettings(args[0]);                  break;
      case 'getUserSettings':         result = getUserSettings(args[0]);                   break;
      case 'saveUserSettings':        result = saveUserSettings(args[0]);                  break;
      case 'getUserSettings':         result = getUserSettings(args[0]);                   break;
      case 'getShiftEvents':          result = getShiftEvents(args[0]);                    break;
      case 'logShiftEvent':           result = logShiftEvent(args[0]);                     break;
      case 'markEmployeeActive':      result = markEmployeeActive(args[0]);                break;
      case 'updateUserActivity':      result = updateUserActivity(args[0]);                break;
      case 'getUserActivity':         result = getUserActivity();                          break;
      case 'getCompanyProfile':       result = getCompanyProfile();                        break;
      case 'saveCompanyProfile':      result = saveCompanyProfile(args[0]);                break;
      case 'syncCompanyProfileToSheet': result = saveCompanyProfile(args[0]);              break;
      case 'getServicesCatalog':        result = getServicesCatalog(args[0]); break;
case 'saveServiceEntry':          result = saveServiceEntry(args[0]); break;
case 'deleteServiceEntry':        result = deleteServiceEntry(args[0]); break;
case 'saveAllServices':           result = saveAllServices(args[0]); break;
case 'seedDefaultServices':       result = _ensureServicesCatalogSheet() ? { success:true } : { success:false }; break;
case 'getQuotationBuilderConfig': result = getQuotationBuilderConfig(); break;
case 'saveQuotationBuilderRow':   result = saveQuotationBuilderRow(args[0]); break;
case 'deleteQuotationBuilderRow': result = deleteQuotationBuilderRow(args[0]); break;
case 'initQuotationBuilderSheet': result = _initQuotationBuilderSheet(SpreadsheetApp.getActiveSpreadsheet()) ? { success:true } : { success:false }; break;
case 'viewQuotation':
  var qHtml = viewQuotation(args[0]);
  return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
      case 'getCompanyProfilesList':  result = getCompanyProfilesList();                   break;
      case 'saveProfileEntry':        result = saveProfileEntry(args[0]);                  break;
      case 'deleteProfileEntry':      result = deleteProfileEntry(args[0]);                break;
      case 'openSheetTab':            result = openSheetTab(args[0]);                      break;
      default:
        result = { error: 'Unknown function: ' + fn };
    }
  } catch(err) {
    result = { error: err.message, fn: fn };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildQuotationHTML(q) {
  const cp           = getCompanyProfile();
  const companyName  = cp.company_name    || 'Your Company';
  const senderEmail  = cp.sender_email    || cp.admin_email || '';
  const supportPhone = cp.support_phone   || cp.whatsapp_number || '';
  const website      = cp.company_website || '';
  const address      = cp.company_address || '';
  const logoUrl      = cp.company_logo_url|| '';
  const gstNo        = cp.gst_number      || '';
  const footerExtra  = cp.quotation_footer|| '';
  const terms        = cp.quotation_terms || 'This quotation is valid for 30 days from the date of issue. Prices are subject to applicable taxes unless stated as inclusive.';
  const currency     = cp.currency_symbol || '₹';
  const bankInfo     = (cp.bank_name && cp.bank_account)
    ? 'Bank: ' + cp.bank_name + ' | A/C: ' + cp.bank_account + (cp.bank_ifsc ? ' | IFSC: ' + cp.bank_ifsc : '') : '';

  const items    = JSON.parse(q.cartItems || q.items || '[]');
  const discount = parseFloat(q.discount  || 0);
  const subtotal = items.reduce(function(a, i) { return a + (parseFloat(i.price) || 0); }, 0);
  const discAmt  = Math.round(subtotal * discount / 100);
  const total    = subtotal - discAmt;

  function fmt(n) { return currency + Number(n || 0).toLocaleString('en-IN'); }

  const itemRows = items.map(function(item, idx) {
    return '<tr>'
      + '<td style="padding:10px 14px;border-bottom:1px solid #f0f0f0;color:#9ca3af;font-size:12px">' + (idx + 1) + '</td>'
      + '<td style="padding:10px 14px;border-bottom:1px solid #f0f0f0;font-weight:600;font-size:13px">' + (item.name || '') + '</td>'
      + '<td style="padding:10px 14px;border-bottom:1px solid #f0f0f0;color:#6b7280;font-size:11px">'
        + (item.desc || item.coverage || '') + (item.timeline ? '<br><span style="color:#d97706">⏱ ' + item.timeline + '</span>' : '')
        + (item.retainer ? '<br><span style="color:#059669">+ ' + fmt(item.retainer) + ' retainer/mo</span>' : '')
        + (item.additionalCost ? '<br><span style="color:#dc2626">+ ' + item.additionalCost + '</span>' : '')
      + '</td>'
      + '<td style="padding:10px 14px;border-bottom:1px solid #f0f0f0;text-align:right;font-weight:700;font-size:13px;white-space:nowrap">' + fmt(item.price) + '</td>'
      + '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Quotation ' + (q.quoteNum || q.ID || '') + ' — ' + companyName + '</title>'
    + '<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;background:#f4f6f9;color:#1e293b}'
    + '.card{max-width:780px;margin:28px auto;background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 4px 28px rgba(0,0,0,.09)}'
    + '.hdr{background:linear-gradient(135deg,#0f172a,#1e3a5f);color:#fff;padding:28px 36px;display:flex;justify-content:space-between;align-items:flex-start;gap:20px;flex-wrap:wrap}'
    + '.hdr-left{display:flex;flex-direction:column;gap:4px}'
    + '.logo{width:52px;height:52px;border-radius:10px;object-fit:contain;background:#fff;padding:4px;margin-bottom:8px}'
    + '.hdr-left h1{font-size:22px;font-weight:800;margin-bottom:2px}.hdr-left p{font-size:11px;opacity:.55}'
    + '.hdr-right{text-align:right}'
    + '.badge{display:inline-block;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.28);border-radius:20px;padding:3px 14px;font-size:10px;font-weight:800;letter-spacing:1px;margin-bottom:8px}'
    + '.hdr-right .qid{font-size:12px;opacity:.7;font-weight:600}.hdr-right .qdate{font-size:11px;opacity:.5}'
    + '.body{padding:28px 36px}'
    + '.print-btn{display:block;margin:0 0 22px;padding:12px 24px;background:#1e3a5f;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;width:100%;text-align:center}'
    + '@media print{.print-btn{display:none}}'
    + '.meta{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:24px;padding:16px;background:#f8fafc;border-radius:8px;border:1px solid #e2e8f0}'
    + '.meta-item label{font-size:9px;font-weight:800;color:#94a3b8;text-transform:uppercase;letter-spacing:.7px;display:block;margin-bottom:3px}'
    + '.meta-item span{font-size:13px;font-weight:600;color:#1e293b}'
    + 'table{width:100%;border-collapse:collapse}'
    + 'thead tr{background:#f8fafc}'
    + 'thead th{padding:10px 14px;text-align:left;font-size:9px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e2e8f0}'
    + 'thead th:last-child{text-align:right}'
    + '.totals{padding:18px 36px;background:#f8fafc;border-top:1px solid #e2e8f0}'
    + '.trow{display:flex;justify-content:space-between;font-size:12px;color:#64748b;margin-bottom:6px}'
    + '.tfinal{display:flex;justify-content:space-between;font-size:16px;font-weight:800;color:#1e293b;padding:12px 16px;background:#1e3a5f;color:#fff;border-radius:8px;margin-top:8px}'
    + '.terms{padding:14px 36px;background:#fffbeb;border-top:1px solid #fef3c7;font-size:11px;color:#92400e;line-height:1.7}'
    + '.footer{padding:18px 36px;font-size:11px;color:#94a3b8;border-top:1px solid #f0f0f0;line-height:1.9}'
    + '@media(max-width:600px){.card{margin:0;border-radius:0}.hdr,.body,.totals,.footer,.terms{padding-left:18px;padding-right:18px}.meta{grid-template-columns:1fr}.hdr{flex-direction:column}}'
    + '</style></head><body><div class="card">'
    + '<div class="hdr">'
    +   '<div class="hdr-left">'
    +     (logoUrl ? '<img class="logo" src="' + logoUrl + '" alt="Logo">' : '')
    +     '<h1>' + companyName + '</h1>'
    +     (website ? '<p>' + website.replace(/https?:\/\//,'') + '</p>' : '')
    +     (senderEmail ? '<p>' + senderEmail + '</p>' : '')
    +   '</div>'
    +   '<div class="hdr-right">'
    +     '<div class="badge">QUOTATION</div>'
    +     '<div class="qid">' + (q.quoteNum || q.ID || '') + '</div>'
    +     '<div class="qdate">' + new Date().toLocaleDateString('en-IN', { day:'2-digit', month:'long', year:'numeric' }) + '</div>'
    +   '</div>'
    + '</div>'
    + '<div class="body">'
    + '<button class="print-btn" onclick="window.print()">⬇ Download / Print PDF</button>'
    + '<div class="meta">'
    +   '<div class="meta-item"><label>Prepared For</label><span>' + (q.clientName || '—') + '</span></div>'
    +   '<div class="meta-item"><label>Company</label><span>' + (q.clientCompany || '—') + '</span></div>'
    +   '<div class="meta-item"><label>Valid Until</label><span>' + (q.validTill || '—') + '</span></div>'
    +   '<div class="meta-item"><label>Client Email</label><span>' + (q.clientEmail || '—') + '</span></div>'
    +   (gstNo ? '<div class="meta-item"><label>Our GST No.</label><span>' + gstNo + '</span></div>' : '')
    + '</div>'
    + '<table><thead><tr><th>#</th><th>Service / Product</th><th>Details</th><th style="text-align:right">Amount</th></tr></thead>'
    + '<tbody>' + itemRows + '</tbody></table>'
    + '</div>'
    + '<div class="totals">'
    + (discAmt > 0 ? '<div class="trow"><span>Subtotal</span><span>' + fmt(subtotal) + '</span></div>' : '')
    + (discAmt > 0 ? '<div class="trow"><span>Discount (' + discount + '%)</span><span>− ' + fmt(discAmt) + '</span></div>' : '')
    + '<div class="tfinal"><span>Total</span><span>' + fmt(total) + '</span></div>'
    + '</div>'
    + (terms ? '<div class="terms">' + terms + '</div>' : '')
    + '<div class="footer">'
    +   '<strong>' + companyName + '</strong>'
    +   (senderEmail   ? ' · ' + senderEmail   : '')
    +   (supportPhone  ? ' · ' + supportPhone  : '')
    +   (address       ? '<br>' + address       : '')
    +   (bankInfo      ? '<br>' + bankInfo      : '')
    +   (footerExtra   ? '<br>' + footerExtra   : '')
    +   '<br>Quotation valid till ' + (q.validTill || 'N/A') + '.'
    + '</div>'
    + '</div></body></html>';
}

function saveQuotationToSheet(q) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('QUOTATIONS');
    if (!sheet) {
      sheet = ss.insertSheet('QUOTATIONS');
      sheet.appendRow(['quoteNum','clientName','clientCompany','clientEmail','clientPhone','cartItems','discount','total','validTill','savedAt']);
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === q.quoteNum) {
        sheet.getRange(i+1, 1, 1, 10).setValues([[
          q.quoteNum, q.clientName, q.clientCompany, q.clientEmail, q.clientPhone,
          q.cartItems, q.discount, q.total, q.validTill, new Date().toISOString()
        ]]);
        return 'updated';
      }
    }
    sheet.appendRow([
      q.quoteNum, q.clientName, q.clientCompany, q.clientEmail, q.clientPhone,
      q.cartItems, q.discount, q.total, q.validTill, new Date().toISOString()
    ]);
    return 'saved';
  } catch(e) {
    return 'error: ' + e.message;
  }
}

// ═══════════════════════════════════════════════════════════════
//  ACTIVITY LOG — LEAD SHEET COLUMN (single source of truth)
// ═══════════════════════════════════════════════════════════════

function _getActivityIcon(action) {
  const m = {
    'CALL_MADE':'📞','CALL_COMPLETED':'📞','EMAIL_SENT':'✉️','WHATSAPP_SENT':'💬',
    'MEETING_SCHEDULED':'📅','MEETING_COMPLETED':'✅','NOTE_ADDED':'📝',
    'TICKET_RAISED':'🎫','QUOTATION_CREATED':'📄','QUOTATION_SENT':'📤',
    'LEAD_ADDED':'➕','LEAD_UPDATED':'✏️','STATUS_CHANGED':'🔄','FUNNEL_MOVED':'🔀',
    'HOT_LEAD_TAGGED':'🔥','WORKFLOW_CREATED':'⚡','LEAD_ARCHIVED':'📦',
    'LOGIN':'🟢','LOGOUT':'⭕','PROFILE_UPDATED':'🏢'
  };
  return m[action] || '📌';
}

function _getActivityColor(action) {
  const m = {
    'CALL_MADE':'#0D9F6F','CALL_COMPLETED':'#0D9F6F',
    'EMAIL_SENT':'#1A56DB','WHATSAPP_SENT':'#059669',
    'MEETING_SCHEDULED':'#7C3AED','MEETING_COMPLETED':'#059669',
    'NOTE_ADDED':'#6C2BD9','TICKET_RAISED':'#D97706',
    'QUOTATION_CREATED':'#1A56DB','QUOTATION_SENT':'#0891B2',
    'LEAD_ADDED':'#059669','LEAD_UPDATED':'#6B7280',
    'STATUS_CHANGED':'#D97706','FUNNEL_MOVED':'#7C3AED',
    'HOT_LEAD_TAGGED':'#E53E3E','WORKFLOW_CREATED':'#0891B2',
    'LEAD_ARCHIVED':'#9CA3AF'
  };
  return m[action] || '#6B7280';
}

function _getActivityLabel(action) {
  const m = {
    'CALL_MADE':'Call Made','CALL_COMPLETED':'Call Done','EMAIL_SENT':'Email Sent',
    'WHATSAPP_SENT':'WhatsApp Sent','MEETING_SCHEDULED':'Meeting Scheduled',
    'MEETING_COMPLETED':'Meeting Done','NOTE_ADDED':'Note Added',
    'TICKET_RAISED':'Ticket Raised','QUOTATION_CREATED':'Quote Created',
    'QUOTATION_SENT':'Quote Sent','LEAD_ADDED':'Lead Added',
    'LEAD_UPDATED':'Lead Updated','STATUS_CHANGED':'Status Changed',
    'FUNNEL_MOVED':'Stage Moved','HOT_LEAD_TAGGED':'Hot Tagged',
    'WORKFLOW_CREATED':'Workflow Sent','LEAD_ARCHIVED':'Archived'
  };
  return m[action] || action.replace(/_/g,' ');
}

// Find Activity Log column index in MASTER_LEADS (1-indexed)
function _getActivityColIndex(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let idx = headers.findIndex(h => String(h).toLowerCase().replace(/[\s_]/g,'').includes('activitylog'));
  if (idx >= 0) return idx + 1;
  // Auto-add column if missing
  const newCol = sh.getLastColumn() + 1;
  sh.getRange(1, newCol).setValue('Activity Log')
    .setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
  return newCol;
}

// Append one activity entry to the lead's Activity Log column
function _appendLeadActivity(leadId, act) {
  try {
    if (!leadId) return;
    const sh = _getSheet(SN.LEADS);
    if (!sh) return;
    const actCol = _getActivityColIndex(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) !== String(leadId)) continue;
      const cell = sh.getRange(i + 2, actCol);
      let existing = [];
      try { existing = JSON.parse(String(cell.getValue() || '[]')); } catch(e) { existing = []; }
      if (!Array.isArray(existing)) existing = [];
      existing.push({
        t:   act.time    || new Date().toISOString(),
        by:  act.by      || act.empName || _getEmployeeName(act.email || _getEmail()) || 'Agent',
        em:  act.email   || _getEmail(),
        act: act.action  || act.type   || 'ACTION',
        lbl: _getActivityLabel(act.action || act.type || ''),
        det: String(act.details || act.text || '').slice(0, 120),
        ic:  act.icon    || _getActivityIcon(act.action  || act.type),
        col: act.color   || _getActivityColor(act.action || act.type)
      });
      if (existing.length > 100) existing = existing.slice(-100);
      cell.setValue(JSON.stringify(existing));
      try { SpreadsheetApp.flush(); } catch(e) {}
      return;
    }
  } catch(e) { Logger.log('_appendLeadActivity error: ' + e.message); }
}

function logUniversalActivity(payload) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const now = payload.time ? new Date(payload.time).toISOString() : new Date().toISOString();
    const empName  = payload.by    || _getEmployeeName(payload.email || _getEmail()) || 'Agent';
    const empEmail = payload.email || _getEmail() || '';
    const actType  = payload.action || payload.type || 'note';
    const details  = String(payload.text || payload.details || '').slice(0,500);
    const leadId   = payload.leadId || '';

    // 1. Write to LEAD_ACTIVITY sheet (flat row — easy to query per lead)
    if (leadId) {
      const laSh = _ensureLeadActivitySheet(ss);
      laSh.appendRow([
        leadId,
        actType,
        details,
        empName,
        empEmail,
        _getActivityIcon(actType.toUpperCase()),
        _getActivityColor(actType.toUpperCase()),
        payload.extra ? JSON.stringify(payload.extra) : '',
        now
      ]);
    }

    // 2. Also embed in MASTER_LEADS Activity Log column (for fast per-lead read)
    if (leadId) {
      _appendLeadActivity(leadId, {
        email:   empEmail,
        by:      empName,
        action:  actType,
        details: details,
        time:    now
      });
    }

    // 3. Write to global ACTIVITY_LOG sheet (for admin overview / performance)
    _logActivity({
      leadId:  leadId,
      action:  actType.toUpperCase(),
      details: details,
      email:   empEmail,
      empName: empName,
      source:  payload.source || 'ui',
      timestamp: now
    });

    try { SpreadsheetApp.flush(); } catch(e) {}
    return { success: true };
  } catch(e) {
    Logger.log('logUniversalActivity error: ' + e.message);
    return { success: false };
  }
}

function getLeadActivity(leadId) {
  try {
    if (!leadId) return [];
    const sh = _getSheet(SN.LEADS);
    if (!sh || sh.getLastRow() < 2) return [];
    const actCol = _getActivityColIndex(sh);
    const lastRow = sh.getLastRow();
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) !== String(leadId)) continue;
      let acts = [];
      try { acts = JSON.parse(String(sh.getRange(i + 2, actCol).getValue() || '[]')); } catch(e) { acts = []; }
      if (!Array.isArray(acts)) acts = [];
      return [...acts].reverse().map(a => ({
        time:    a.t   || '',
        by:      a.by  || 'Agent',
        email:   a.em  || '',
        action:  a.act || '',
        label:   a.lbl || _getActivityLabel(a.act || ''),
        details: a.det || '',
        icon:    a.ic  || '📌',
        color:   a.col || '#6B7280'
      }));
    }
    return [];
  } catch(e) { Logger.log('getLeadActivity error: ' + e.message); return []; }
}

function storeQuotationAndGetUrl(q) {
  try {
    var sh = _ensureQuotationsSheet();
    if (!sh) return null;
    var quoteNum = q.quoteNum || q.ID || ('Q' + Date.now().toString().slice(-8));
    var itemsJson = JSON.stringify(q.items || []);
    var now = new Date().toISOString();
    // Check if quote already exists and update, else append
    var data = sh.getDataRange().getValues();
    var h = data[0];
    var qnIdx = h.indexOf('QuoteNum');
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][qnIdx]) === String(quoteNum)) {
        sh.getRange(i + 1, 1, 1, 15).setValues([[
          quoteNum, q.clientName || '', q.clientCompany || '', q.clientEmail || '',
          q.validTill || '', q.discount || 0, q.discountAmt || 0,
          q.subtotal || 0, q.total || 0, itemsJson,
          q.notes || '', q.createdBy || '', q.createdAt || now, 'Generated', ''
        ]]);
        found = true;
        break;
      }
    }
    if (!found) {
      sh.appendRow([
        quoteNum, q.clientName || '', q.clientCompany || '', q.clientEmail || '',
        q.validTill || '', q.discount || 0, q.discountAmt || 0,
        q.subtotal || 0, q.total || 0, itemsJson,
        q.notes || '', q.createdBy || '', q.createdAt || now, 'Generated', ''
      ]);
    }
    SpreadsheetApp.flush();
    // Build the quotation HTML and return as a data URL served via the script
    var htmlContent = buildQuotationHTML(q);
    // Store HTML in a separate sheet for retrieval
    var previewSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QUOTATION_PREVIEWS');
    if (!previewSh) {
      previewSh = SpreadsheetApp.getActiveSpreadsheet().insertSheet('QUOTATION_PREVIEWS');
      previewSh.appendRow(['QuoteNum','HTML','CreatedAt']);
    }
    var previewData = previewSh.getDataRange().getValues();
    var previewH = previewData[0];
    var pqIdx = previewH.indexOf('QuoteNum');
    var foundPreview = false;
    for (var pi = 1; pi < previewData.length; pi++) {
      if (String(previewData[pi][pqIdx]) === String(quoteNum)) {
        previewSh.getRange(pi + 1, 2, 1, 2).setValues([[htmlContent, now]]);
        foundPreview = true;
        break;
      }
    }
    if (!foundPreview) previewSh.appendRow([quoteNum, htmlContent, now]);
    SpreadsheetApp.flush();
    // Return the URL to view this quotation
    var scriptUrl = ScriptApp.getService().getUrl();
    return scriptUrl + '?fn=viewQuotation&args=' + encodeURIComponent(JSON.stringify([quoteNum]));
  } catch(e) {
    Logger.log('storeQuotationAndGetUrl error: ' + e.message);
    return null;
  }
}

function viewQuotation(quoteNum) {
  try {
    var previewSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('QUOTATION_PREVIEWS');
    if (!previewSh) return HtmlService.createHtmlOutput('<h2>Quotation not found</h2>');
    var data = previewSh.getDataRange().getValues();
    var h = data[0];
    var qnIdx = h.indexOf('QuoteNum');
    var htmlIdx = h.indexOf('HTML');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][qnIdx]) === String(quoteNum)) {
        return HtmlService.createHtmlOutput(String(data[i][htmlIdx] || '')).setTitle('Quotation ' + quoteNum);
      }
    }
    return HtmlService.createHtmlOutput('<h2>Quotation ' + quoteNum + ' not found</h2>');
  } catch(e) {
    return HtmlService.createHtmlOutput('<h2>Error: ' + e.message + '</h2>');
  }
}

function buildQuotationHTML(q) {
  try {
    var profile = getCompanyProfile();
    var companyName    = profile.companyName    || 'AfterResult Solutions';
    var tagline        = profile.tagline        || 'Digital Marketing Agency';
    var logoUrl        = profile.logoUrl        || '';
    var companyEmail   = profile.email          || 'info.afterresult@gmail.com';
    var companyPhone   = profile.phone          || '+919991283530';
    var companyAddress = profile.address        || '';
    var companyWebsite = profile.website        || '';
    var currency       = profile.currency       || 'Rs.';
    var gstNo          = profile.gst            || '';
    var bankName       = profile.bankName       || '';
    var accountNo      = profile.accountNo      || '';
    var ifsc           = profile.ifsc           || '';
    var upi            = profile.upi            || '';
    var accountHolder  = profile.accountHolder  || '';
    var footerText     = profile.quotationFooter|| 'Thank you for your business.';
    var termsText      = profile.terms          || '';

    var items   = q.items || [];
    var discount = parseFloat(q.discount || 0);
    var subtotal = items.reduce(function(a, i) { return a + (parseFloat(i.price || 0) * parseInt(i.qty || 1)); }, 0);
    var discAmt  = Math.round(subtotal * discount / 100);
    var total    = subtotal - discAmt;

    var itemRows = items.map(function(item, idx) {
      var lineTotal = parseFloat(item.price || 0) * parseInt(item.qty || 1);
      return '<tr>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #f0f0f0;color:#555">' + (idx + 1) + '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #f0f0f0"><strong>' + (item.name || '') + '</strong>' + (item.unit ? '<br><span style="font-size:11px;color:#999">' + item.unit + '</span>' : '') + '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #f0f0f0;text-align:center">' + (item.qty || 1) + '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #f0f0f0;text-align:right">' + currency + Number(item.price || 0).toLocaleString('en-IN') + '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #f0f0f0;text-align:right;font-weight:600">' + currency + Number(lineTotal).toLocaleString('en-IN') + '</td>' +
      '</tr>';
    }).join('');

    var bankHtml = (bankName || accountNo || upi) ? (
      '<div style="margin-top:24px;padding:16px;background:#f8f9fa;border-radius:8px;border-left:3px solid #3B5BDB">' +
      '<div style="font-size:12px;font-weight:700;color:#333;margin-bottom:8px;text-transform:uppercase;letter-spacing:0.5px">Payment Details</div>' +
      (bankName      ? '<div style="font-size:12px;color:#555;margin-bottom:3px"><strong>Bank:</strong> ' + bankName + '</div>' : '') +
      (accountHolder ? '<div style="font-size:12px;color:#555;margin-bottom:3px"><strong>Account Name:</strong> ' + accountHolder + '</div>' : '') +
      (accountNo     ? '<div style="font-size:12px;color:#555;margin-bottom:3px"><strong>Account No:</strong> ' + accountNo + '</div>' : '') +
      (ifsc          ? '<div style="font-size:12px;color:#555;margin-bottom:3px"><strong>IFSC:</strong> ' + ifsc + '</div>' : '') +
      (upi           ? '<div style="font-size:12px;color:#555;margin-bottom:3px"><strong>UPI:</strong> ' + upi + '</div>' : '') +
      '</div>'
    ) : '';

    var termsHtml = termsText ? (
      '<div style="margin-top:16px;padding:12px 16px;background:#fff8e1;border-radius:8px;border-left:3px solid #f59e0b">' +
      '<div style="font-size:11px;font-weight:700;color:#92400e;margin-bottom:6px;text-transform:uppercase">Terms and Conditions</div>' +
      '<div style="font-size:11px;color:#78350f;line-height:1.7">' + termsText.replace(/\n/g, '<br>') + '</div>' +
      '</div>'
    ) : '';

    return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>Quotation ' + (q.quoteNum || '') + ' — ' + companyName + '</title>' +
      '<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Arial,sans-serif;background:#f5f5f5;color:#333}' +
      '.card{max-width:720px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08)}' +
      '.hdr{background:linear-gradient(135deg,#1a1a2e,#16213e);color:#fff;padding:28px 32px}' +
      '.hdr h1{font-size:22px;font-weight:800;margin-bottom:2px}' +
      '.hdr p{font-size:13px;opacity:.65;margin-top:2px}' +
      '.badge{display:inline-block;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:4px 14px;font-size:11px;font-weight:700;margin-top:12px}' +
      '.body{padding:28px 32px}' +
      '.print-btn{display:block;margin:0 0 20px 0;padding:12px 24px;background:#3B5BDB;color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;width:100%}' +
      '@media print{.print-btn{display:none}}' +
      '.meta{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:24px;padding:16px;background:#f8f9fa;border-radius:8px}' +
      '.meta-item label{font-size:10px;font-weight:700;color:#999;text-transform:uppercase;letter-spacing:.5px;display:block;margin-bottom:3px}' +
      '.meta-item span{font-size:13px;font-weight:600;color:#111}' +
      'table{width:100%;border-collapse:collapse}' +
      'thead th{padding:10px 16px;text-align:left;font-size:10px;font-weight:700;color:#999;text-transform:uppercase;background:#f8f9fa}' +
      'thead th:last-child,thead th:nth-child(3),thead th:nth-child(4){text-align:right}' +
      '.totals{padding:16px 32px;background:#f8f9fa;border-top:1px solid #eee}' +
      '.trow{display:flex;justify-content:space-between;font-size:13px;color:#555;margin-bottom:5px}' +
      '.tfinal{display:flex;justify-content:space-between;font-size:18px;font-weight:800;color:#111;padding-top:10px;border-top:2px solid #111;margin-top:6px}' +
      '.footer-section{padding:20px 32px;border-top:1px solid #f0f0f0}' +
      '.company-footer{text-align:center;font-size:11px;color:#aaa;margin-top:16px;padding-top:16px;border-top:1px solid #f0f0f0}' +
      '@media(max-width:600px){.card{margin:0;border-radius:0}.hdr,.body,.totals,.footer-section{padding-left:16px;padding-right:16px}.meta{grid-template-columns:1fr}}' +
      '</style></head><body>' +
      '<div class="card">' +
      '<div class="hdr">' +
      (logoUrl ? '<img src="' + logoUrl + '" style="height:40px;margin-bottom:12px;display:block" onerror="this.style.display=\'none\'">' : '') +
      '<h1>' + companyName + '</h1><p>' + tagline + '</p>' +
      (gstNo ? '<p style="font-size:11px;opacity:.6;margin-top:4px">GST: ' + gstNo + '</p>' : '') +
      '<div class="badge">QUOTATION ' + (q.quoteNum || '') + '</div>' +
      '</div>' +
      '<div class="body">' +
      '<button class="print-btn" onclick="window.print()">Download / Print PDF</button>' +
      '<div class="meta">' +
      '<div class="meta-item"><label>Prepared For</label><span>' + (q.clientName || '—') + '</span></div>' +
      '<div class="meta-item"><label>Company</label><span>' + (q.clientCompany || '—') + '</span></div>' +
      '<div class="meta-item"><label>Email</label><span>' + (q.clientEmail || '—') + '</span></div>' +
      '<div class="meta-item"><label>Valid Until</label><span>' + (q.validTill || '—') + '</span></div>' +
      '</div>' +
      '<table><thead><tr><th>#</th><th>Service / Product</th><th style="text-align:center">Qty</th><th style="text-align:right">Unit Price</th><th style="text-align:right">Amount</th></tr></thead>' +
      '<tbody>' + itemRows + '</tbody></table>' +
      '</div>' +
      '<div class="totals">' +
      (discAmt > 0 ? '<div class="trow"><span>Subtotal</span><span>' + currency + subtotal.toLocaleString('en-IN') + '</span></div>' : '') +
      (discAmt > 0 ? '<div class="trow"><span>Discount (' + discount + '%)</span><span>- ' + currency + discAmt.toLocaleString('en-IN') + '</span></div>' : '') +
      '<div class="tfinal"><span>Total</span><span>' + currency + total.toLocaleString('en-IN') + '</span></div>' +
      '</div>' +
      '<div class="footer-section">' +
      bankHtml +
      termsHtml +
      (q.notes ? '<div style="margin-top:16px;padding:12px;background:#f0f7ff;border-radius:8px;font-size:12px;color:#444;line-height:1.6"><strong>Notes:</strong> ' + q.notes + '</div>' : '') +
      '<div style="margin-top:20px;font-size:12px;color:#666;line-height:1.7">' + footerText + '</div>' +
      '<div class="company-footer">' +
      companyName +
      (companyEmail ? ' · ' + companyEmail : '') +
      (companyPhone ? ' · ' + companyPhone : '') +
      (companyAddress ? '<br>' + companyAddress : '') +
      (companyWebsite ? '<br><a href="' + companyWebsite + '" style="color:#3B5BDB">' + companyWebsite + '</a>' : '') +
      '</div>' +
      '</div>' +
      '</div></body></html>';
  } catch(e) {
    Logger.log('buildQuotationHTML error: ' + e.message);
    return '<html><body><h2>Error building quotation: ' + e.message + '</h2></body></html>';
  }
}

function getQuotations(empEmail) {
  try {
    var sh = _ensureQuotationsSheet();
    if (!sh || sh.getLastRow() < 2) return [];
    var data = sh.getDataRange().getValues();
    var h = data[0];
    function col(k) { return h.indexOf(k); }
    return data.slice(1).filter(function(r) { return r[0]; }).map(function(r) {
      return {
        quoteNum:      String(r[col('QuoteNum')]      || r[0] || ''),
        clientName:    String(r[col('ClientName')]    || r[1] || ''),
        clientCompany: String(r[col('ClientCompany')] || r[2] || ''),
        clientEmail:   String(r[col('ClientEmail')]   || r[3] || ''),
        validTill:     String(r[col('ValidTill')]     || r[4] || ''),
        discount:      parseFloat(r[col('Discount')]  || r[5] || 0),
        subtotal:      parseFloat(r[col('Subtotal')]  || r[7] || 0),
        total:         parseFloat(r[col('Total')]     || r[8] || 0),
        notes:         String(r[col('Notes')]         || r[10] || ''),
        createdBy:     String(r[col('CreatedBy')]     || r[11] || ''),
        createdAt:     String(r[col('CreatedAt')]     || r[12] || ''),
        status:        String(r[col('Status')]        || r[13] || ''),
        viewUrl:       String(r[col('ViewUrl')]       || r[14] || '')
      };
    });
  } catch(e) {
    Logger.log('getQuotations error: ' + e.message);
    return [];
  }
}
// ═══════════════════════════════════════════════════════════════
//  MASTER INITIALIZATION
// ═══════════════════════════════════════════════════════════════
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    _initAuthSheet(ss);
    _initLeadsSheet(ss);
    _initEmployeesSheet(ss);
    _initChatSheet(ss);
    _initChannelsSheet(ss);
    _initWorkflowsSheet(ss);
    _initQuotationsSheet(ss);
    _initMeetingsSheet(ss);
    _initTicketsSheet(ss);
    _initNotesSheet(ss);
    _initActivityLogSheet(ss);
    _initRoutingRulesSheet(ss);
    _initCompanyProfileSheet(ss);
    _initServicesCatalogSheet(ss);
    _initCompanyProfilesSheet(ss);
    _initQuotationBuilderSheet(ss);
    _setupOnOpenTrigger();
    return 'All 16 sheets initialized. Merged into ACTIVITY_LOG: ACTION_LOGS + SHIFT_TRACKER + PERFORMANCE + EMAILS_SENT + LEAD_ACTIVITY + AdminLog. Merged into COMPANY_PROFILE: COMPANIES + ADMIN_CONFIG. Removed: DELETED_LEADS (use archived column), NOTIFICATIONS (use CHAT), LEAD_FILTERS, PIPELINE_CONFIG.';
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

// ─── AUTH SHEET ────────────────────────────────────────────────
function _initAuthSheet(ss) {
  const headers = ['Email','Name','Role','Status','Employee ID','Experience','Categories','Quota','Shift Start','Shift End','Phone','Last Login','Added Date'];
  const sh = _createSheet(ss, SN.AUTH, headers);
  _applyDropdown(sh, 3, 3, 2, 1000, DD.AUTH_ROLE);
  _applyDropdown(sh, 4, 4, 2, 1000, DD.EMP_STATUS);
  _applyDropdown(sh, 6, 6, 2, 1000, DD.EMP_EXP);
  _applyConditionalColors(sh, 4, [
    { value:'Active', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Inactive', bg:'#F3F4F6', fg:'#6B7280' },
    { value:'Deactivated', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'Hold', bg:'#FEF3C7', fg:'#92400E' },
    { value:'On Leave', bg:'#DBEAFE', fg:'#1E40AF' }
  ]);
  // Add default admin
  const existing = sh.getDataRange().getValues();
  if (existing.length < 2 || !existing[1][0]) {
    sh.appendRow([CONFIG.ADMIN_EMAIL,'Admin','Admin','Active','EMP001','Expert (5+ yr)','All','20','09:00','18:00','','',new Date().toISOString()]);
  }
  return sh;
}

// ─── LEADS SHEET ───────────────────────────────────────────────
function _initLeadsSheet(ss) {
  const headers = [
    'Lead ID','Date Added','Full Name','Email','Phone',
    'Company','Domain','Industry','Source','Status','Funnel Stage',
    'Priority','Hot Lead','Assigned To','City','State','Country',
    'Deal Value (₹)','Notes','Last Action','Last Action Date',
    'Call Count','Meeting Count','Quotation Sent','Quotation Status',
    'Tags','Created By','Website','LinkedIn','Activity Log'
  ];
  const sh = _createSheet(ss, SN.LEADS, headers, _dummyLeads());
  _applyDropdown(sh, 8, 8, 2, 5000, DD.INDUSTRY);     // Industry
  _applyDropdown(sh, 9, 9, 2, 5000, DD.SOURCE);           // Source
  _applyDropdown(sh, 10, 10, 2, 5000, DD.LEAD_STATUS);    // Status
  _applyDropdown(sh, 11, 11, 2, 5000, DD.FUNNEL_STAGE);   // Funnel Stage
  _applyDropdown(sh, 12, 12, 2, 5000, DD.PRIORITY);       // Priority
  _applyDropdown(sh, 13, 13, 2, 5000, DD.HOT);            // Hot Lead
  _applyDropdown(sh, 25, 25, 2, 5000, DD.QUOTATION_STATUS); // Quotation Status
  _applyConditionalColors(sh, 10, [
    { value:'New', bg:'#EFF6FF', fg:'#1D4ED8' },
    { value:'Contacted', bg:'#F0FDF4', fg:'#166534' },
    { value:'In Progress', bg:'#FFF7ED', fg:'#9A3412' },
    { value:'Qualified', bg:'#F0FDF4', fg:'#166534' },
    { value:'Proposal Sent', bg:'#EDE9FE', fg:'#5B21B6' },
    { value:'Negotiating', bg:'#FEF3C7', fg:'#92400E' },
    { value:'Sold', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Closed Lost', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'On Hold', bg:'#F3F4F6', fg:'#6B7280' },
    { value:'Archived', bg:'#F9FAFB', fg:'#9CA3AF' },
    { value:'Junk', bg:'#FEE2E2', fg:'#B91C1C' }
  ]);
  _applyConditionalColors(sh, 12, [
    { value:'Critical', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'High', bg:'#FEF3C7', fg:'#92400E' },
    { value:'Medium', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Low', bg:'#F3F4F6', fg:'#6B7280' }
  ]);
  _applyConditionalColors(sh, 13, [
    { value:'Yes', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'No', bg:'#F3F4F6', fg:'#6B7280' }
  ]);
  return sh;
}

// ─── EMPLOYEES SHEET ───────────────────────────────────────────
function _initEmployeesSheet(ss) {
  const headers = [
    'Employee ID','Name','Email','Phone','Role','Experience Level',
    'Categories','Lead Quota','Leads Assigned','Leads Closed',
    'Status','Shift Start','Shift End','Break Duration (min)',
    'Is Active','Last Active','Join Date','Avatar','Daily Score',
    'Monthly Score','Department','Manager Email','Notes'
  ];
  const sh = _createSheet(ss, SN.EMPLOYEES, headers, _dummyEmployees());
  _applyDropdown(sh, 5, 5, 2, 1000, DD.EMP_ROLE);
  _applyDropdown(sh, 6, 6, 2, 1000, DD.EMP_EXP);
  _applyDropdown(sh, 11, 11, 2, 1000, DD.EMP_STATUS);
  _applyConditionalColors(sh, 11, [
    { value:'Active', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Inactive', bg:'#F3F4F6', fg:'#6B7280' },
    { value:'On Leave', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Deactivated', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'Hold', bg:'#FEF3C7', fg:'#92400E' },
    { value:'Suspended', bg:'#FEE2E2', fg:'#7F1D1D' },
    { value:'Probation', bg:'#FEF3C7', fg:'#78350F' }
  ]);
  return sh;
}

// ─── CHAT MESSAGES SHEET ───────────────────────────────────────
function _initChatSheet(ss) {
  const headers = [
    'Message ID','Timestamp','Channel ID','Channel Name',
    'From Email','From Name','To Email','Message Text',
    'Message Type','Is Read','Workflow ID','Lead ID',
    'Quotation ID','Edited','Edit Timestamp','Reactions','Thread ID'
  ];
  const sh = _createSheet(ss, SN.CHAT, headers);
  _applyDropdown(sh, 9, 9, 2, 10000, DD.MSG_TYPE);
  _applyDropdown(sh, 10, 10, 2, 10000, ['Yes','No']);
  _applyConditionalColors(sh, 9, [
    { value:'workflow', bg:'#FEF3C7', fg:'#92400E' },
    { value:'system', bg:'#EDE9FE', fg:'#5B21B6' },
    { value:'announcement', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'quotation', bg:'#D1FAE5', fg:'#065F46' }
  ]);
  return sh;
}

// ─── CHANNELS SHEET ────────────────────────────────────────────
function _initChannelsSheet(ss) {
  const headers = [
    'Channel ID','Channel Name','Channel Type','Description',
    'Status','Created By','Created Date','Member Count','Members (comma)',
    'Message Count','Last Message','Last Message By','Last Message Time',
    'Is Archived','Archive Date','Pinned','Category'
  ];
  const sh = _createSheet(ss, SN.CHANNELS, headers, _defaultChannels());
  _applyDropdown(sh, 3, 3, 2, 500, DD.CHANNEL_TYPE);
  _applyDropdown(sh, 5, 5, 2, 500, DD.CHANNEL_STATUS);
  _applyConditionalColors(sh, 5, [
    { value:'Active', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Archived', bg:'#F3F4F6', fg:'#6B7280' },
    { value:'Read Only', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Muted', bg:'#FEF3C7', fg:'#92400E' }
  ]);
  return sh;
}
// ════ GOOGLE APPS SCRIPT - EMAIL BACKEND FUNCTION ════
// ADD THIS TO YOUR GOOGLE APPS SCRIPT FILE

function sendCompanyEmail(to, subject, body) {
  try {
    const senderName = 'AfterResult Solutions';

    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: body.replace(/\n/g, '<br>'),
      name: senderName,
      replyTo: 'info.afterresult@gmail.com'
    });

    try {
      const ss = SpreadsheetApp.getActive();
      let sheet = ss.getSheetByName('Email_Log');
      if (!sheet) {
        sheet = ss.insertSheet('Email_Log');
        sheet.appendRow(['Timestamp', 'To', 'Subject', 'Body Preview', 'Sender', 'Status']);
      }
      sheet.appendRow([
        new Date(),
        to,
        subject,
        body.substring(0, 100),
        'info.afterresult@gmail.com',
        'Sent Successfully'
      ]);
    } catch (e) {}

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ─── WORKFLOWS SHEET ───────────────────────────────────────────
function _initWorkflowsSheet(ss) {
  const headers = [
    'Workflow ID','Created Date','Type','Lead ID','Lead Name',
    'Lead Company','Industry','Forwarded To Team','Service Required',
    'Client Budget (₹)','Duration','Specific Requirements','Tags',
    'Priority','Status','Submitted By','Assigned To',
    'Target Channel','Chat Message ID','Quotation ID',
    'Completed Date','Notes','Last Updated'
  ];
  const sh = _createSheet(ss, SN.WORKFLOWS, headers);
  _applyDropdown(sh, 3, 3, 2, 2000, DD.WORKFLOW_TYPE);
  _applyDropdown(sh, 8, 8, 2, 2000, DD.WORKFLOW_TEAM);
  _applyDropdown(sh, 9, 9, 2, 2000, DD.SERVICE);
  _applyDropdown(sh, 11, 11, 2, 2000, DD.DURATION);
  _applyDropdown(sh, 14, 14, 2, 2000, DD.PRIORITY);
  _applyDropdown(sh, 15, 15, 2, 2000, DD.WORKFLOW_STATUS);
  _applyConditionalColors(sh, 15, [
    { value:'Pending', bg:'#FEF3C7', fg:'#92400E' },
    { value:'In Progress', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Completed', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Cancelled', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'On Hold', bg:'#F3F4F6', fg:'#6B7280' },
    { value:'Rejected', bg:'#FEE2E2', fg:'#7F1D1D' }
  ]);
  return sh;
}

// ─── QUOTATIONS SHEET ──────────────────────────────────────────
function _initQuotationsSheet(ss) {
  const headers = [
    'Quotation ID','Created Date','Lead ID','Lead Name','Lead Email',
    'Company','Industry','Service Package','Monthly Value (₹)',
    'Duration (Months)','Total Contract Value (₹)',
    'Discount %','Final Value (₹)','GST 18% (₹)','Grand Total (₹)',
    'Status','Version','Sent Date','Valid Until',
    'Accepted Date','Rejected Date','Rejection Reason',
    'Workflow ID','Created By','Notes','Last Updated'
  ];
  const sh = _createSheet(ss, SN.QUOTATIONS, headers);
  _applyDropdown(sh, 8, 8, 2, 2000, DD.SERVICE);
  _applyDropdown(sh, 16, 16, 2, 2000, DD.QUOTATION_STATUS);
  _applyConditionalColors(sh, 16, [
    { value:'Draft', bg:'#F3F4F6', fg:'#6B7280' },
    { value:'Sent', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Under Review', bg:'#FEF3C7', fg:'#92400E' },
    { value:'Accepted', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Rejected', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'Won', bg:'#BBF7D0', fg:'#14532D' },
    { value:'Lost', bg:'#FECACA', fg:'#7F1D1D' },
    { value:'Expired', bg:'#F3F4F6', fg:'#9CA3AF' },
    { value:'Revised', bg:'#EDE9FE', fg:'#5B21B6' }
  ]);
  return sh;
}

// ─── MEETINGS SHEET ────────────────────────────────────────────
function _initMeetingsSheet(ss) {
  const headers = [
    'Meeting ID','Lead ID','Lead Name','Lead Email','Employee Email',
    'Title','Date','Time','Duration (min)','Agenda',
    'Calendar Event ID','Google Meet Link','Recording Link',
    'Status','Outcome Notes','Follow Up Required','Follow Up Date',
    'Created Date','Last Updated'
  ];
  const sh = _createSheet(ss, SN.MEETINGS, headers);
  _applyDropdown(sh, 14, 14, 2, 2000, DD.MEETING_STATUS);
  _applyDropdown(sh, 16, 16, 2, 2000, ['Yes','No']);
  _applyConditionalColors(sh, 14, [
    { value:'Scheduled', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Confirmed', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Completed', bg:'#BBF7D0', fg:'#14532D' },
    { value:'Cancelled', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'Rescheduled', bg:'#FEF3C7', fg:'#92400E' },
    { value:'No Show', bg:'#FECACA', fg:'#7F1D1D' }
  ]);
  return sh;
}

function uploadPublicQuotation(quoteNum, htmlContent) {
  try {
    // Find or create a public folder
    var folderName = 'AfterResult Quotations';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    // Make folder publicly accessible
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Delete old file with same quote number if exists
    var existing = folder.getFilesByName(quoteNum + '.html');
    while (existing.hasNext()) { existing.next().setTrashed(true); }
    
    // Create new file
    var file = folder.createFile(quoteNum + '.html', htmlContent, 'text/html');
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Return direct view URL
    return 'https://drive.google.com/file/d/' + file.getId() + '/view';
  } catch(e) {
    return null;
  }
}

function sendAdminNotificationEmail(toEmail, subject, body) {
  try {
    MailApp.sendEmail({
      to: toEmail,
      bcc: CONFIG.BCC_EMAIL,
      subject: subject,
      body: body,
      name: CONFIG.COMPANY + ' via Nucleous'
    });
    _logEmailSent({ to: toEmail, subject, body, type: 'notification', from: CONFIG.COMPANY_EMAIL });
    return { success: true };
  } catch(e) {
    Logger.log('Admin email error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function sendEmailFromPanel(payload) {
  try {
    if (!payload || !payload.to || !payload.subject || (!payload.body && !payload.html)) {
      return { success: false, message: 'Missing required fields (to / subject / body)' };
    }

    // Sender display name: prefer payload.fromName, then sentBy name portion, then company default
    const senderDisplayName = payload.fromName
      || (payload.sentBy ? payload.sentBy.split('@')[0].replace(/[._]/g, ' ') : null)
      || CONFIG.COMPANY;

    // Plain-text body (fallback strip HTML)
    const plainBody = payload.body
      || (payload.html || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');

    const options = {
      to:      payload.to,
      subject: payload.subject,
      body:    plainBody,
      name:    senderDisplayName,
      replyTo: payload.sentBy || CONFIG.ADMIN_EMAIL
    };

    // Send as HTML when available
    options.htmlBody = payload.html
      || payload.body.replace(/\n/g, '<br>');

    if (payload.cc)  options.cc  = payload.cc;
    if (payload.bcc) options.bcc = payload.bcc;

    MailApp.sendEmail(options);

    const emailRecord = {
      id:       'EMAIL-' + Date.now(),
      to:       payload.to,
      subject:  payload.subject,
      body:     plainBody,
      cc:       payload.cc       || '',
      sentAt:   new Date().toISOString(),
      sentBy:   payload.sentBy   || _getEmail(),
      leadId:   payload.leadId   || '',
      leadName: payload.leadName || '',
      type:     payload.type     || 'manual'
    };
    _logEmailSent(emailRecord);

    // Also log to activity log if a lead is attached
    if (payload.leadId) {
      try {
        _logActivity({
          type:   'EMAIL_SENT',
          leadId: payload.leadId,
          note:   'Email sent via composer — ' + payload.subject,
          by:     senderDisplayName,
          email:  payload.sentBy || _getEmail()
        });
      } catch(ae) {}
    }

    Logger.log('[sendEmailFromPanel] Sent to ' + payload.to + ' — ' + payload.subject);
    return { success: true, emailId: emailRecord.id };

  } catch(e) {
    Logger.log('[sendEmailFromPanel] Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Alias — callable as fn=sendEmailComposer from the email composer "Send" button
function sendEmailComposer(payload) {
  return sendEmailFromPanel(payload);
}

function getRecentEmails(limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('EMAILS_SENT');
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getDataRange().getValues();
    const lim  = parseInt(limit) || 50;
    return data.slice(1).reverse().slice(0, lim).map(r => ({
      id:       r[0], to:r[1], subject:r[2], body:r[3], cc:r[4],
      sentAt:   r[5], sentBy:r[6], leadId:r[7], leadName:r[8], type:r[9]
    }));
  } catch(e) { return []; }
}

function _logEmailSent(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh   = ss.getSheetByName('EMAILS_SENT');
    if (!sh) {
      sh = ss.insertSheet('EMAILS_SENT');
      sh.appendRow(['ID','To','Subject','Body','CC','Sent At','Sent By','Lead ID','Lead Name','Type']);
      sh.getRange(1,1,1,10).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      email.id || 'EMAIL-'+Date.now(),
      email.to || '', email.subject || '', email.body || '', email.cc || '',
      email.sentAt || new Date().toISOString(),
      email.sentBy || _getEmail(),
      email.leadId || '', email.leadName || '', email.type || 'manual'
    ]);
  } catch(e) {}
}

function getDeletedLeads() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('DELETED_LEADS');
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({
      id:r[0], name:r[1], company:r[2], deletedAt:r[3], deletedBy:r[4], reason:r[5]
    })).reverse();
  } catch(e) { return []; }
}

function deleteLead(leadId) {
  try {
    const sh   = _getSheet(SN.LEADS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(leadId)) {
        const lead = data[i];
        // Save to deleted leads sheet
        const ss     = SpreadsheetApp.getActiveSpreadsheet();
        let delSheet = ss.getSheetByName('DELETED_LEADS');
        if (!delSheet) {
          delSheet = ss.insertSheet('DELETED_LEADS');
          delSheet.appendRow(['Lead ID','Name','Company','Deleted At','Deleted By','Reason']);
          delSheet.getRange(1,1,1,6).setBackground('#7F1D1D').setFontColor('#fff').setFontWeight('bold');
        }
        delSheet.appendRow([lead[0],lead[2],lead[5],new Date().toISOString(),_getEmail(),'Deleted by admin']);
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false, message: e.message }; }
}

// ─── TICKETS SHEET ─────────────────────────────────────────────
function _initTicketsSheet(ss) {
  const headers = [
    'Ticket ID','Lead ID','Lead Name','Raised By','Assigned To',
    'Category','Priority','Subject','Description','Status',
    'Created Date','Last Updated','Resolved Date','Closed Date',
    'Resolution Notes','SLA Breach','Reopen Count','Internal Notes','Tags'
  ];
  const sh = _createSheet(ss, SN.TICKETS, headers);
  _applyDropdown(sh, 6, 6, 2, 2000, DD.TICKET_CATEGORY);
  _applyDropdown(sh, 7, 7, 2, 2000, DD.TICKET_PRIORITY);
  _applyDropdown(sh, 10, 10, 2, 2000, DD.TICKET_STATUS);
  _applyDropdown(sh, 16, 16, 2, 2000, ['Yes','No']);
  _applyConditionalColors(sh, 10, [
    { value:'Open', bg:'#FEF3C7', fg:'#92400E' },
    { value:'In Progress', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Resolved', bg:'#D1FAE5', fg:'#065F46' },
    { value:'Closed', bg:'#F3F4F6', fg:'#6B7280' },
    { value:'Reopened', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'Escalated', bg:'#FEE2E2', fg:'#7F1D1D' },
    { value:'On Hold', bg:'#EDE9FE', fg:'#5B21B6' }
  ]);
  _applyConditionalColors(sh, 7, [
    { value:'Urgent', bg:'#7F1D1D', fg:'#FFFFFF' },
    { value:'Critical', bg:'#FEE2E2', fg:'#991B1B' },
    { value:'High', bg:'#FEF3C7', fg:'#92400E' },
    { value:'Medium', bg:'#DBEAFE', fg:'#1E40AF' },
    { value:'Low', bg:'#F3F4F6', fg:'#6B7280' }
  ]);
  return sh;
}

// ─── NOTES SHEET ───────────────────────────────────────────────
function _initNotesSheet(ss) {
  const headers = [
    'Note ID','Lead ID','Employee Email','Note Text','Timestamp',
    'Note Type','Call Log ID','Is Pinned','Visibility'
  ];
  const sh = _createSheet(ss, SN.NOTES, headers);
  _applyDropdown(sh, 6, 6, 2, 5000, DD.NOTE_TYPE);
  _applyDropdown(sh, 8, 8, 2, 5000, ['Yes','No']);
  _applyDropdown(sh, 9, 9, 2, 5000, ['Public','Private','Team Only']);
  return sh;
}

// ─── ACTION LOGS SHEET ─────────────────────────────────────────
// ─── ACTION LOGS / SHIFTS / PERFORMANCE → all merged into ACTIVITY_LOG ──
function _initLogsSheet(ss)       { return _initActivityLogSheet(ss); }
function _initShiftsSheet(ss)     { return _initActivityLogSheet(ss); }
function _initPerformanceSheet(ss){ return _initActivityLogSheet(ss); }

// ─── ACTIVITY LOG SHEET (single merged log for all activity types) ──
function _initActivityLogSheet(ss) {
  const existing = ss.getSheetByName(SN.ACTIVITY);
  if (existing && existing.getLastRow() > 1) return existing;
  const headers = [
    'Log ID','Timestamp','Date','Employee Email','Employee Name',
    'Lead ID','Activity Type','Details','Extra JSON',
    'Duration (sec)','Outcome','Score Delta','Session ID','Source'
  ];
  const sh = _createSheet(ss, SN.ACTIVITY, headers);
  _applyDropdown(sh, 7, 7, 2, 50000, [
    'LEAD_ADDED','LEAD_UPDATED','LEAD_ARCHIVED','LEAD_DELETED',
    'CALL_MADE','CALL_COMPLETED','EMAIL_SENT','WHATSAPP_SENT',
    'MEETING_SCHEDULED','MEETING_COMPLETED','MEETING_CANCELLED',
    'NOTE_ADDED','TICKET_RAISED','TICKET_RESOLVED',
    'QUOTATION_CREATED','QUOTATION_SENT','QUOTATION_WON','QUOTATION_LOST',
    'WORKFLOW_CREATED','WORKFLOW_COMPLETED',
    'SHIFT_START','SHIFT_END','BREAK_START','BREAK_END',
    'HOT_LEAD_TAGGED','STATUS_CHANGED','FUNNEL_MOVED',
    'LOGIN','LOGOUT','ADMIN_ACTION','MENTION_SENT','PROFILE_UPDATED'
  ]);
  return sh;
}

// ─── ROUTING RULES SHEET ─────────────────────────────────────
function _initRoutingRulesSheet(ss) {
  const headers = ['Rule ID','Name','Trigger','Trigger Value','Target Emp ID','Target Emp Name','Target Emp Email','Override','Keep Copy','Notify','Active','Fire Count','Created At','Last Fired'];
  const sh = _createSheet(ss, SN.ROUTING, headers);
  return sh;
}

function _initCompanyProfileSheet(ss) {
  return _ensureCompanyProfileSheet();
}

function _initServicesCatalogSheet(ss) {
  return _ensureServicesCatalogSheet();
}

function _initCompanyProfilesSheet(ss) {
  const headers = ['Profile ID','Name','Category','Description','Pages','Drive URL','WA Message','Active','Sort Order','Created At','Notes'];
  return _createSheet(ss, SN.PROFILES, headers);
}

function _dummyLeads()     { return []; }
function _dummyEmployees() { return []; }

function _setupOnOpenTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const already = triggers.some(t => t.getHandlerFunction() === 'onOpen');
    if (!already) ScriptApp.newTrigger('onOpen').forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  } catch(e) {}
}

// ─── COMPANY PROFILE SHEET ───────────────────────────────────
// Replaces: COMPANIES + ADMIN_CONFIG — one key-value sheet per deployment
function _initCompanyProfileSheet(ss) {
  const existing = ss.getSheetByName(SN.COMPANY);
  if (existing && existing.getLastRow() > 1) return existing;
  const headers = ['Field','Value','Type','Notes'];
  const sh = _createSheet(ss, SN.COMPANY, headers);
  const defaults = [
    ['company_name',      '',               'text',   'Your company or brand name shown on emails and quotations'],
    ['admin_name',        '',               'text',   'Owner / Admin full name'],
    ['admin_email',       '',               'email',  'Admin login email address'],
    ['sender_email',      '',               'email',  'Gmail address used to send automatic emails (must be linked to this Apps Script)'],
    ['sender_name',       '',               'text',   'Display name shown on outgoing emails (e.g. "Team AfterResult")'],
    ['reply_to_email',    '',               'email',  'Reply-to address on outgoing emails (optional)'],
    ['bcc_email',         '',               'email',  'BCC address on every outgoing email (optional — good for monitoring)'],
    ['support_phone',     '',               'phone',  'Support / WhatsApp number in international format (+91XXXXXXXXXX)'],
    ['whatsapp_number',   '',               'phone',  'WhatsApp Business number for auto-notifications'],
    ['google_meet_link',  '',               'url',    'Default Google Meet link inserted into meeting invites'],
    ['meeting_cc_emails', '',               'text',   'Comma-separated CC emails added to all meeting invites'],
    ['company_website',   '',               'url',    'Company website URL shown on quotations and emails'],
    ['company_address',   '',               'text',   'Full company address (used in quotations and email footer)'],
    ['company_logo_url',  '',               'url',    'Direct URL to company logo PNG/JPG (shown on quotations)'],
    ['email_signature',   '',               'text',   'Signature block appended to ALL outgoing emails'],
    ['quotation_footer',  '',               'text',   'Extra footer text on all quotations (e.g. bank details, tagline)'],
    ['quotation_terms',   'This quotation is valid for 30 days. Prices subject to applicable taxes.','text','Terms & conditions text printed at bottom of every quotation'],
    ['quotation_validity','30',             'number', 'Default quotation validity in days'],
    ['gst_number',        '',               'text',   'GST / Tax registration number (shown on quotations)'],
    ['bank_name',         '',               'text',   'Bank name for quotation / invoice payment section'],
    ['bank_account',      '',               'text',   'Bank account number'],
    ['bank_ifsc',         '',               'text',   'IFSC / bank routing code'],
    ['currency_symbol',   '₹',             'text',   'Currency symbol used across quotations and pricing'],
    ['currency_code',     'INR',            'text',   'Currency code (INR, USD, AED, GBP, etc.)'],
    ['industry',          '',               'text',   'Primary industry / sector of your business'],
    ['company_size',      '',               'text',   'Team size range (e.g. 1–10, 11–50, 51–200)'],
    ['plan',              'Trial',          'select', 'Subscription plan (Trial / Starter / Pro / Enterprise)'],
    ['timezone',          'Asia/Kolkata',   'text',   'Timezone for shift tracking and timestamps'],
    ['script_url',        '',               'url',    'Deployed Apps Script web app URL (auto-filled on save)'],
    ['status',            'Active',         'select', 'Account status'],
    ['registered_at',     new Date().toISOString(), 'datetime', 'Account creation date'],
    ['updated_at',        new Date().toISOString(), 'datetime', 'Last profile update']
  ];
  sh.getRange(2, 1, defaults.length, 4).setValues(defaults);
  sh.getRange(2, 1, defaults.length, 1).setFontWeight('bold').setBackground('#F1F5F9');
  sh.getRange(2, 3, defaults.length, 2).setFontColor('#94A3B8');
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 380);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 420);
  return sh;
}

// ─── SERVICES CATALOG SHEET ──────────────────────────────────
// Any company can add their own products/services here — feeds the Quotation Builder
function _initServicesCatalogSheet(ss) {
  const existing = ss.getSheetByName(SN.SERVICES);
  if (existing && existing.getLastRow() > 1) return existing;
  const headers = [
    'Service ID','Category','Plan Name','Description',
    'Price','Price Type','Retainer','Additional Cost',
    'Tax Included','Timeline','Is Addon','Addon For',
    'Active','Sort Order','Created At','Notes'
  ];
  const sh = _createSheet(ss, SN.SERVICES, headers);
  _applyDropdown(sh,  6,  6, 2, 2000, ['monthly','one-time','per-listing','per-unit','custom']);
  _applyDropdown(sh,  9,  9, 2, 2000, ['Yes','No']);
  _applyDropdown(sh, 11, 11, 2, 2000, ['Yes','No']);
  _applyDropdown(sh, 13, 13, 2, 2000, ['Yes','No']);
  const defaults = [
    ['SVC001','Social Media Management','Instagram Organic','15 Creatives + Festivals + Engagement',9999,'monthly',4500,'','Yes','Monthly','No','','Yes',1,new Date().toISOString(),''],
    ['SVC002','Social Media Management','Instagram + LinkedIn Organic','20 Creatives each + Festivals + Engagement',22999,'monthly',7700,'','No','Monthly','No','','Yes',2,new Date().toISOString(),''],
    ['SVC003','Social Media Management','Instagram + LinkedIn + Ads','20 Creatives each + Ads Management',24999,'monthly',9700,'Ads Budget','No','Monthly','No','','Yes',3,new Date().toISOString(),''],
    ['SVC004','Social Media Management','Full Package (All Platforms + Ads)','Instagram + LinkedIn + Facebook + Ads',28999,'monthly',11000,'Ads Budget','No','Monthly','No','','Yes',4,new Date().toISOString(),''],
    ['SVC005','Amazon Marketplace','Basic (35 Listings – 45 Days)','35 product listings optimized',20650,'one-time','','','Yes','45 Days','No','','Yes',10,new Date().toISOString(),''],
    ['SVC006','Amazon Marketplace','Standard (50 Listings – 45 Days)','50 product listings optimized',29500,'one-time','','','Yes','45 Days','No','','Yes',11,new Date().toISOString(),''],
    ['SVC007','Amazon Marketplace','Premium (100 Listings – 50 Days)','100 product listings optimized',59000,'one-time','','','Yes','50 Days','No','','Yes',12,new Date().toISOString(),''],
    ['SVC008','Flipkart Marketplace','Basic (35 Listings – 45 Days)','35 product listings on Flipkart',18585,'one-time','','','Yes','45 Days','No','','Yes',13,new Date().toISOString(),''],
    ['SVC009','SEO','Basic SEO (20 keywords)','Technical SEO + On-Page + Monthly Report',14999,'monthly','','','No','Monthly','No','','Yes',20,new Date().toISOString(),''],
    ['SVC010','Google Ads','Starter Campaign','Campaign setup + monthly management',8999,'monthly','','Ads Budget','No','Monthly','No','','Yes',30,new Date().toISOString(),''],
    ['SVC011','Website Development','Landing Page','Single-page responsive website',14999,'one-time','','','No','7–10 Days','No','','Yes',40,new Date().toISOString(),''],
    ['SVC012','Website Development','Business Website','Multi-page responsive business site',29999,'one-time','','','No','15–20 Days','No','','Yes',41,new Date().toISOString(),''],
    ['SVC013','E-Commerce Development','Shopify Starter (50 products)','Theme setup + payment gateway',24999,'one-time','','','No','15–20 Days','No','','Yes',50,new Date().toISOString(),''],
    ['SVC014','Custom Package','Custom Package','As per client requirements — enter price manually',0,'custom','','','No','TBD','No','','Yes',99,new Date().toISOString(),''],
    ['ADD001','Social Media Management','Instagram Followers (200+)','Active/Inactive followers add-on',510,'one-time','','','Yes','','Yes','Social Media Management',1,new Date().toISOString(),'Addon'],
    ['ADD002','Social Media Management','LinkedIn Connections (200+)','Main account connecting add-on',200,'one-time','','','Yes','','Yes','Social Media Management',2,new Date().toISOString(),'Addon'],
    ['ADD003','Amazon Marketplace','Inventory Management','/month add-on for stock tracking',3000,'monthly','','','No','Monthly','Yes','Amazon Marketplace',1,new Date().toISOString(),'Addon'],
    ['ADD004','Amazon Marketplace','Ads with Plan','Ads management add-on to marketplace plan',5000,'monthly','','Ads Budget','No','Monthly','Yes','Amazon Marketplace',2,new Date().toISOString(),'Addon']
  ];
  sh.getRange(2, 1, defaults.length, 16).setValues(defaults);
  return sh;
}

// ─── COMPANY PROFILES SHEET (PDF / Drive links) ──────────────
function _initCompanyProfilesSheet(ss) {
  const existing = ss.getSheetByName(SN.PROFILES);
  if (existing && existing.getLastRow() > 1) return existing;
  const headers = [
    'Profile ID','Name','Category','Description','Pages',
    'Drive URL','WhatsApp Message','Active','Sort Order','Created At','Notes'
  ];
  const sh = _createSheet(ss, SN.PROFILES, headers);
  _applyDropdown(sh, 3, 3, 2, 500, ['Company Profile','Service Deck','Case Study','Pricing','Portfolio','Other']);
  _applyDropdown(sh, 8, 8, 2, 500, ['Yes','No']);
  return sh;
}

// ═══════════════════════════════════════════════════════════════
//  HELPER UTILITIES
// ═══════════════════════════════════════════════════════════════

function _createSheet(ss, name, headers, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    sh.clearContents();
    sh.clearFormats();
  }
  const hRange = sh.getRange(1, 1, 1, headers.length);
  hRange.setValues([headers]);
  hRange.setBackground('#111827')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(9)
        .setBorder(true, true, true, true, true, true);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, headers.length, 140);
  if (rows && rows.length > 0) {
    sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  return sh;
}

function _applyDropdown(sh, colStart, colEnd, rowStart, rowEnd, values) {
  try {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(false)
      .build();
    sh.getRange(rowStart, colStart, rowEnd - rowStart + 1, colEnd - colStart + 1)
      .setDataValidation(rule);
  } catch(e) {}
}

function _applyConditionalColors(sh, col, rules) {
  try {
    const conditionalRules = rules.map(r => {
      return SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(r.value)
        .setBackground(r.bg)
        .setFontColor(r.fg)
        .setRanges([sh.getRange(2, col, 5000, 1)])
        .build();
    });
    const existing = sh.getConditionalFormatRules();
    sh.setConditionalFormatRules([...existing, ...conditionalRules]);
  } catch(e) {}
}

function _getSheet(name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(name);
    if (!sh) {
      initializeSheets();
      sh = ss.getSheetByName(name);
    }
    return sh || null;
  } catch(e) {
    Logger.log('_getSheet error for ' + name + ': ' + e.message);
    return null;
  }
}

function _getEmail() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch(e) { return ''; }
}

function _now() {
  return new Date().toISOString();
}

function _today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _uid(prefix) {
  return (prefix || 'ID') + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

function _createSheet(ss, name, headers, seedRows) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() < 1) {
    sh.appendRow(headers);
    const hRange = sh.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#111827').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
    sh.setFrozenRows(1);
    if (seedRows && seedRows.length) {
      sh.getRange(2, 1, seedRows.length, seedRows[0].length).setValues(seedRows);
    }
  }
  return sh;
}

function _applyDropdown(sh, startCol, endCol, startRow, endRow, values) {
  try {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build();
    sh.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1).setDataValidation(rule);
  } catch(e) {}
}

function _applyConditionalColors(sh, col, rules) {
  try {
    const cfRules = rules.map(r =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(r.value)
        .setBackground(r.bg).setFontColor(r.fg)
        .setRanges([sh.getRange(2, col, Math.min(sh.getMaxRows(), 5000), 1)])
        .build()
    );
    sh.setConditionalFormatRules(cfRules);
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════
//  AUTH
// ═══════════════════════════════════════════════════════════════

function checkUserAuth(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SN.AUTH);
    if (!sh) {
      initializeSheets();
      sh = ss.getSheetByName(SN.AUTH);
    }
    if (!sh) return { authorized: false, message: 'Setup required. Contact admin.' };
    const data = sh.getDataRange().getValues();
    const norm = (email || '').trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[0]).trim().toLowerCase() === norm) {
        const status = String(row[3] || 'Active');
        if (status === 'Deactivated' || status === 'Suspended') {
          return { authorized: false, message: 'Account ' + status.toLowerCase() + '. Contact admin.' };
        }
        try { sh.getRange(i + 1, 12).setValue(_now()); } catch(e) {}
        return {
          authorized: true,
          email: norm,
          name: String(row[1] || email.split('@')[0]),
          role: String(row[2] || 'BDE'),
          status: status,
          empId: String(row[4] || 'EMP-' + i),
          exp: String(row[5] || 'Mid (1-3 yr)'),
          categories: String(row[6] || '').split(',').map(c => c.trim()).filter(Boolean),
          quota: String(row[7] || '10'),
          shiftStart: String(row[8] || '09:00'),
          shiftEnd: String(row[9] || '18:00'),
          phone: String(row[10] || '')
        };
      }
    }
    // Auto add employee if not found but exists in EMPLOYEES sheet
    const empSh = ss.getSheetByName(SN.EMPLOYEES);
    if (empSh) {
      const empData = empSh.getDataRange().getValues();
      for (let i = 1; i < empData.length; i++) {
        if (String(empData[i][2]).trim().toLowerCase() === norm) {
          // Add to AUTH sheet automatically
          sh.appendRow([
            norm,
            String(empData[i][1] || email.split('@')[0]),
            String(empData[i][4] || 'BDE'),
            'Active',
            String(empData[i][0] || 'EMP-' + Date.now()),
            String(empData[i][5] || 'Mid (1-3 yr)'),
            String(empData[i][6] || ''),
            String(empData[i][7] || '10'),
            String(empData[i][11] || '09:00'),
            String(empData[i][12] || '18:00'),
            String(empData[i][3] || ''),
            _now(),
            _now()
          ]);
          return {
            authorized: true,
            email: norm,
            name: String(empData[i][1] || email.split('@')[0]),
            role: String(empData[i][4] || 'BDE'),
            status: 'Active',
            empId: String(empData[i][0] || 'EMP-' + i),
            exp: String(empData[i][5] || 'Mid (1-3 yr)'),
            categories: String(empData[i][6] || '').split(',').filter(Boolean),
            quota: String(empData[i][7] || '10'),
            shiftStart: String(empData[i][11] || '09:00'),
            shiftEnd: String(empData[i][12] || '18:00'),
            phone: String(empData[i][3] || '')
          };
        }
      }
    }
    return { authorized: false, message: 'Email not authorized. Contact Admin' };
  } catch(e) {
    return { authorized: false, message: 'Server error: ' + e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  CHAT — SAVE ALL MESSAGES TO SHEET
// ═══════════════════════════════════════════════════════════════

function sendChat(fromEmail, toEmail, messageText, channelId, options) {
  try {
    const sh = _getSheet(SN.CHAT);
    if (!sh) return { success: false, message: 'Chat sheet not found' };
    const opts = options || {};
    const msgId = _uid('MSG');
    const channelName = opts.channelName || _getChannelName(channelId) || channelId;
    const fromName = opts.fromName || _getEmployeeName(fromEmail) || fromEmail;
    sh.appendRow([
      msgId,                          // Message ID
      _now(),                         // Timestamp
      channelId || 'general',         // Channel ID
      channelName,                    // Channel Name
      fromEmail || _getEmail(),       // From Email
      fromName,                       // From Name
      toEmail || '',                  // To Email
      messageText || '',              // Message Text
      opts.type || 'text',            // Message Type
      'No',                           // Is Read
      opts.workflowId || '',          // Workflow ID
      opts.leadId || '',              // Lead ID
      opts.quotationId || '',         // Quotation ID
      'No',                           // Edited
      '',                             // Edit Timestamp
      '',                             // Reactions
      opts.threadId || ''             // Thread ID
    ]);
    // Update channel last message info
    _updateChannelLastMessage(channelId, messageText, fromName);
    return { success: true, msgId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  SHEET-BASED CHAT WORKFLOW — Chat from Google Sheets
// ═══════════════════════════════════════════════════════════════

function setupChatWorkflowSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('CHAT_WORKFLOW');
    if (!sh) sh = ss.insertSheet('CHAT_WORKFLOW');
    else { sh.clearContents(); sh.clearFormats(); sh.clearConditionalFormatRules(); }

    // Column layout (no Send Time — auto-captured on send):
    // A: Action | B: Channel | C: From Name | D: From Email (auto)
    // E: Message Type | F: Message | G: Status | H: Reply To ID | I: Message ID (auto)
    const headers = ['Action','Channel','From Name','From Email','Message Type','Message','Status','Reply To ID','Message ID'];
    const hRange = sh.getRange(1, 1, 1, headers.length);
    hRange.setValues([headers]);
    hRange.setBackground('#0F172A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
    sh.setFrozenRows(1);

    const ROWS = 500;

    // Col A: Action
    const actionRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['New Message','Reply','Forward','Read'], true)
      .setAllowInvalid(false).build();
    sh.getRange(2, 1, ROWS, 1).setDataValidation(actionRule);

    // Col B: Channel — pulled from CHANNELS sheet
    const channelNames = _getChatWorkflowChannelList();
    const chRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(channelNames, true)
      .setAllowInvalid(true).build();
    sh.getRange(2, 2, ROWS, 1).setDataValidation(chRule);

    // Col C: From Name — pulled from EMPLOYEES sheet
    const employeeNames = _getChatWorkflowEmployeeNames();
    const nameRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(employeeNames, true)
      .setAllowInvalid(true).build();
    sh.getRange(2, 3, ROWS, 1).setDataValidation(nameRule);

    // Col D: From Email — note for user; auto-filled by onEdit trigger
    sh.getRange(2, 4, ROWS, 1).setFontColor('#6B7280').setFontStyle('italic');
    sh.getRange(2, 4, 1, 1).setValue('(auto-filled from From Name)');

    // Col E: Message Type
    const msgTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Acknowledgement','Not Acknowledged','Forward Required','Needs More Detail','Custom'], true)
      .setAllowInvalid(false).build();
    sh.getRange(2, 5, ROWS, 1).setDataValidation(msgTypeRule);

    // Col G: Status
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending','Sent','Failed'], true)
      .setAllowInvalid(false).build();
    sh.getRange(2, 7, ROWS, 1).setDataValidation(statusRule);

    // Column widths
    sh.setColumnWidth(1, 110);  // Action
    sh.setColumnWidth(2, 130);  // Channel
    sh.setColumnWidth(3, 150);  // From Name
    sh.setColumnWidth(4, 210);  // From Email
    sh.setColumnWidth(5, 155);  // Message Type
    sh.setColumnWidth(6, 380);  // Message
    sh.setColumnWidth(7, 90);   // Status
    sh.setColumnWidth(8, 155);  // Reply To ID
    sh.setColumnWidth(9, 175);  // Message ID

    // Wrap message column
    sh.getRange(2, 6, ROWS, 1).setWrap(true);

    // Conditional formatting on Status
    const pendingRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending').setBackground('#FEF3C7').setFontColor('#92400E')
      .setRanges([sh.getRange(2, 7, ROWS, 1)]).build();
    const sentRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sent').setBackground('#D1FAE5').setFontColor('#065F46')
      .setRanges([sh.getRange(2, 7, ROWS, 1)]).build();
    const failedRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Failed').setBackground('#FEE2E2').setFontColor('#991B1B')
      .setRanges([sh.getRange(2, 7, ROWS, 1)]).build();
    sh.setConditionalFormatRules([pendingRule, sentRule, failedRule]);

    // Instruction row
    sh.getRange(502, 1, 1, 9).merge()
      .setValue('HOW TO USE: Select Action, Channel, From Name (email fills automatically). Pick a Message Type — for Custom, type your message in column F. Status auto-sets to Pending. Run "Send Pending Messages" from the Funnel Chat menu. To reply, set Action=Reply and paste the original Message ID in column H.')
      .setBackground('#1E293B').setFontColor('#94A3B8').setFontSize(9).setWrap(true);
    sh.setRowHeight(502, 70);

    return 'CHAT_WORKFLOW sheet created. Employee and channel dropdowns are loaded. Use "Funnel Chat > Open Compose Box" for the quick-send panel.';
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

function _getChatWorkflowChannelList() {
  try {
    const names = getChannels().map(c => c.name || c.id).filter(Boolean);
    return names.length ? names : ['general','sales','leads','escalations','meetings'];
  } catch(e) {
    return ['general','sales','leads','escalations','meetings'];
  }
}

function _getChatWorkflowEmployeeNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('EMPLOYEES');
    if (!sh) return ['Unknown'];
    const data = sh.getDataRange().getValues();
    // Assumes row 1 is header; find Name column (index 0 or search for it)
    const header = data[0].map(h => String(h).toLowerCase());
    const nameCol = header.indexOf('name') !== -1 ? header.indexOf('name') : 0;
    return data.slice(1).map(r => String(r[nameCol] || '').trim()).filter(Boolean);
  } catch(e) {
    return ['Unknown'];
  }
}

// Process all Pending rows in CHAT_WORKFLOW sheet
function processChatWorkflowRows() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('CHAT_WORKFLOW');
    if (!sh) {
      setupChatWorkflowSheet();
      SpreadsheetApp.getUi().alert('CHAT_WORKFLOW sheet created. Fill in rows and run again.');
      return;
    }

    const data = sh.getDataRange().getValues();
    let processed = 0, failed = 0;

    const PRESETS = {
      'Acknowledgement':  'Acknowledged. Thank you.',
      'Not Acknowledged': 'This has not been acknowledged yet. Please confirm.',
      'Forward Required': 'This message requires forwarding to the relevant team.',
      'Needs More Detail':'Please provide additional detail on this matter.'
    };

    for (let i = 1; i < data.length; i++) {
      const row        = data[i];
      const action     = String(row[0] || '').trim();
      const channelName= String(row[1] || '').trim();
      const fromName   = String(row[2] || '').trim();
      const fromEmail  = String(row[3] || '').trim();
      const msgType    = String(row[4] || '').trim();
      let   message    = String(row[5] || '').trim();
      const status     = String(row[6] || '').trim();
      const replyToId  = String(row[7] || '').trim();

      if (!action || !channelName || status === 'Sent') continue;
      if (action === 'Read') continue;

      if (msgType && msgType !== 'Custom' && !message) {
        message = PRESETS[msgType] || msgType;
      }
      if (!message) continue;

      let msgId = String(row[8] || '').trim();
      if (!msgId) {
        msgId = 'MSG-' + Utilities.getUuid().split('-')[0].toUpperCase();
      }

      try {
        const channelId = _resolveChannelId(channelName);
        const senderEmail = fromEmail || _getEmail();
        const senderName  = fromName  || senderEmail;
        const now = new Date();

        // Write directly to CHAT_MESSAGES (guaranteed delivery)
        _writeToChatMessages({
          id:        msgId,
          channelId: channelId,
          fromEmail: senderEmail,
          fromName:  senderName,
          text:      message,
          type:      'text',
          threadId:  replyToId || '',
          time:      now.toISOString(),
          source:    'workflow'
        });

        // Best-effort call to sendChat for any additional delivery hooks
        try {
          sendChat(senderEmail, null, message, channelId, {
            fromName: senderName,
            type: 'text',
            threadId: replyToId || '',
            msgId: msgId
          });
        } catch(ignored) {}

        sh.getRange(i + 1, 7).setValue('Sent');
        sh.getRange(i + 1, 9).setValue(msgId);
        processed++;

      } catch(rowErr) {
        sh.getRange(i + 1, 7).setValue('Failed');
        failed++;
      }
    }

    const summary = 'Done. Sent: ' + processed + ' | Failed: ' + failed;
    try { SpreadsheetApp.getUi().alert(summary); } catch(e) {}
    return summary;
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

function writeChatMessage(payload) {
  try {
    if (!payload || !payload.text || !payload.channelId) return { success: false, error: 'Missing fields' };
    const msgId = 'MSG-' + Utilities.getUuid().split('-')[0].toUpperCase();
    _writeToChatMessages({
      id:        msgId,
      channelId: payload.channelId,
      fromEmail: payload.fromEmail || _getEmail(),
      fromName:  payload.fromName  || payload.fromEmail || 'Agent',
      text:      payload.text,
      type:      'text',
      threadId:  payload.threadId || '',
      time:      payload.time || new Date().toISOString(),
      source:    payload.source || 'chat-ui'
    });
    return { success: true, msgId: msgId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Load recent messages from a channel INTO the sheet for reading/replying
function loadChannelMessagesIntoSheet(channelId, limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('CHAT_WORKFLOW');
    if (!sh) { setupChatWorkflowSheet(); sh = ss.getSheetByName('CHAT_WORKFLOW'); }

    const msgs = getChat(channelId, null, limit || 50);
    if (!msgs || !msgs.length) {
      try { SpreadsheetApp.getUi().alert('No messages found for: ' + channelId); } catch(e) {}
      return 'No messages found.';
    }

    const existing = sh.getDataRange().getValues();
    let startRow = 2;
    for (let i = 1; i < existing.length; i++) {
      if (!existing[i][0] && !existing[i][5]) { startRow = i + 1; break; }
      startRow = i + 2;
    }

    const rows = msgs.map(m => [
      'Read',
      m.channelName || channelId,
      m.fromName || m.fromEmail || '',
      m.fromEmail || '',
      'Custom',
      m.text || '',
      'Sent',
      m.threadId || '',
      m.id || ''
    ]);

    sh.getRange(startRow, 1, rows.length, 9).setValues(rows);
    sh.getRange(startRow, 1, rows.length, 9).setBackground('#F8FAFC');

    try {
      SpreadsheetApp.getUi().alert(
        'Loaded ' + rows.length + ' messages from #' + channelId + '.\n\n' +
        'To reply: change Action to Reply in any row, edit the Message, clear Status to Pending, then run Send Pending Messages.'
      );
    } catch(e) {}

    return 'Loaded ' + rows.length + ' messages.';
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

// ═══════════════════════════════════════════════════════════════
//  COMPOSE BOX — Full replacement for code.js
//  Fixes: dropdown not loading, adds Reply/Mention with notification
// ═══════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════
//  COMPOSE BOX — Complete replacement (drop into code.js)
//  Fixes: (1) channel/employee dropdown not loading in GAS iframe
//         (2) adds Reply To / @Mention dropdown with notification
// ═══════════════════════════════════════════════════════════════

// ── openComposeBox ───────────────────────────────────────────────
// Opens as a sidebar. Dropdowns are populated by getComposeBoxData()
// via google.script.run. Key fixes vs. old version:
//   • Uses setAttribute('data-*') instead of .dataset on <option>
//     elements — GAS's Caja sandbox strips .dataset assignments on
//     dynamically created nodes but honours setAttribute().
//   • Added explicit disabled=false after population AND a visible
//     "Loading…" placeholder so the user knows to wait.
//   • Mention/Reply section added below the message type picker.
//   • sendFromComposeBox now prepends @Name and writes a
//     NOTIFICATIONS row + sends a GmailApp notification email.
function openComposeBox() {
  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:13px;background:#F8FAFC;color:#1E293B}

/* ── tabs ── */
#tabs{display:flex;border-bottom:2px solid #E2E8F0;background:#fff}
.tab{flex:1;padding:9px 0;text-align:center;cursor:pointer;font-weight:600;font-size:12px;
     color:#64748B;border-bottom:2px solid transparent;margin-bottom:-2px;transition:all .15s}
.tab.active{color:#1D4ED8;border-bottom-color:#1D4ED8}

#paneCompose,#paneHistory{padding:12px;display:none}
#paneCompose.visible,#paneHistory.visible{display:block}

/* ── fields ── */
label{display:block;font-size:11px;font-weight:700;color:#64748B;
      letter-spacing:.3px;text-transform:uppercase;margin:9px 0 4px}
select,input,textarea{width:100%;padding:7px 9px;border:1.5px solid #CBD5E1;
  border-radius:5px;font-size:12px;background:#fff;color:#1E293B;font-family:inherit}
select:focus,input:focus,textarea:focus{outline:none;border-color:#3B82F6}
select:disabled{background:#F1F5F9;color:#94A3B8;cursor:not-allowed}
textarea{resize:vertical;min-height:72px;line-height:1.5}
.row2{display:flex;gap:8px}
.row2 .col{flex:1;min-width:0}

/* ── mention block ── */
#mentionBlock{margin-top:2px;background:#EFF6FF;border:1.5px solid #BFDBFE;
              border-radius:6px;padding:8px 10px}
.mention-note{font-size:10px;color:#1D4ED8;margin-top:4px;line-height:1.4}
.mention-badge{color:#7C3AED;font-weight:700}

/* ── preset hint ── */
#presetHint{font-size:11px;color:#64748B;margin:3px 0 0;padding:5px 8px;
            background:#F1F5F9;border-radius:4px;border-left:3px solid #CBD5E1;display:none}

/* ── buttons ── */
.btn{width:100%;padding:9px;border:none;border-radius:5px;font-size:13px;
     font-weight:700;cursor:pointer;margin-top:8px;transition:all .15s}
#sendBtn{background:#1D4ED8;color:#fff}
#sendBtn:hover:not(:disabled){background:#1E40AF}
#sendBtn:disabled{background:#93C5FD;cursor:not-allowed}
.btn-sm{padding:5px 10px;border:1.5px solid #E2E8F0;border-radius:5px;font-size:11px;
        font-weight:600;background:#fff;color:#475569;cursor:pointer;margin-top:6px}
.btn-sm:hover{border-color:#3B82F6;color:#1D4ED8}
.btn-danger-sm{border-color:#FECACA;color:#EF4444}
.btn-danger-sm:hover{border-color:#EF4444;background:#FFF5F5}

/* ── status ── */
#statusMsg{margin-top:7px;font-size:12px;min-height:16px;padding:5px 8px;
           border-radius:4px;display:none}
#statusMsg.ok {display:block;background:#D1FAE5;color:#065F46;border:1px solid #A7F3D0}
#statusMsg.err{display:block;background:#FEE2E2;color:#991B1B;border:1px solid #FECACA}
#statusMsg.info{display:block;background:#EFF6FF;color:#1E40AF;border:1px solid #BFDBFE}

/* ── spinner ── */
.spinner{display:inline-block;width:11px;height:11px;border:2px solid rgba(255,255,255,.4);
  border-top-color:#fff;border-radius:50%;animation:spin .6s linear infinite;
  margin-right:5px;vertical-align:middle}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── history ── */
#histSearch{margin-bottom:8px}
.hitem{border:1px solid #E2E8F0;border-radius:5px;padding:9px 11px;margin-bottom:7px;
       background:#fff;font-size:12px}
.hitem .hmeta{color:#64748B;font-size:10px;margin-top:4px;line-height:1.4}
.hitem .htxt{margin-top:4px;color:#1E293B;line-height:1.45;word-break:break-word}
.hch{font-weight:700;color:#1D4ED8}
.no-hist{font-size:12px;color:#94A3B8;text-align:center;padding:20px 0}
.hist-actions{display:flex;gap:6px;margin-bottom:8px}

/* loading overlay */
#loadingOverlay{position:fixed;inset:0;background:rgba(248,250,252,.85);
  display:flex;align-items:center;justify-content:center;z-index:99;font-size:13px;color:#64748B}
</style>
</head>
<body>

<!-- Loading overlay shown until data arrives -->
<div id="loadingOverlay">⏳ Loading channels &amp; employees…</div>

<div id="tabs">
  <div class="tab active" onclick="showTab('compose')">Compose</div>
  <div class="tab" onclick="showTab('history')">History</div>
</div>

<!-- ══ COMPOSE PANE ══ -->
<div id="paneCompose" class="visible">

  <label>Action</label>
  <select id="action" onchange="onActionChange()">
    <option value="New Message">New Message</option>
    <option value="Reply">Reply</option>
    <option value="Forward">Forward</option>
  </select>

  <div class="row2">
    <div class="col">
      <label>Channel</label>
      <select id="channel" disabled><option value="">Loading…</option></select>
    </div>
    <div class="col">
      <label>From Name</label>
      <select id="fromName" disabled onchange="onNameChange()"><option value="">Loading…</option></select>
    </div>
  </div>

  <label>From Email</label>
  <select id="fromEmail" disabled><option value="">Loading…</option></select>

  <!-- Reply To Message ID: shown for Reply / Forward -->
  <div id="replyBlock" style="display:none">
    <label>Reply To Message ID</label>
    <input id="replyId" placeholder="Paste original Message ID to thread this reply">
  </div>

  <!-- ── Mention / Reply-To Person ── -->
  <div id="mentionBlock">
    <label style="margin-top:0">
      Mention / Reply To Person
      <span style="font-weight:400;color:#94A3B8;text-transform:none;letter-spacing:0"> — optional</span>
    </label>
    <select id="mentionPerson" onchange="onMentionChange()">
      <option value="">— No mention —</option>
    </select>
    <div class="mention-note" id="mentionNote" style="display:none"></div>
  </div>

  <label>Message Type</label>
  <select id="msgType" onchange="onTypeChange()">
    <option>Acknowledgement</option>
    <option>Not Acknowledged</option>
    <option>Forward Required</option>
    <option>Needs More Detail</option>
    <option>Custom</option>
  </select>
  <div id="presetHint"></div>

  <label>
    Message
    <span id="customNote" style="font-weight:400;color:#94A3B8;text-transform:none;
          letter-spacing:0;display:none"> — type your message below</span>
  </label>
  <textarea id="message">Acknowledged. Thank you.</textarea>

  <button class="btn" id="sendBtn" onclick="doSend()">Send Message</button>
  <div id="statusMsg"></div>
</div>

<!-- ══ HISTORY PANE ══ -->
<div id="paneHistory">
  <input id="histSearch" placeholder="Search history…" oninput="renderHistory()">
  <div class="hist-actions">
    <button class="btn-sm" onclick="refreshHistory()">↺ Refresh</button>
    <button class="btn-sm btn-danger-sm" onclick="clearHistory()">✕ Clear All</button>
  </div>
  <div id="historyList"></div>
</div>

<script>
// ── Preset messages ──────────────────────────────────────────────
var PRESETS = {
  'Acknowledgement':   'Acknowledged. Thank you.',
  'Not Acknowledged':  'This has not been acknowledged yet. Please confirm.',
  'Forward Required':  'This message requires forwarding to the relevant team.',
  'Needs More Detail': 'Please provide additional detail on this matter.',
  'Custom': ''
};

// empData: [{name, email}] — populated after load
var empData       = [];
var pendingMention = '';   // '@Name' string if a mention is active

// ── Tab switcher ─────────────────────────────────────────────────
function showTab(t) {
  ['compose','history'].forEach(function(id, i) {
    document.getElementById('pane' + id.charAt(0).toUpperCase() + id.slice(1))
            .classList.toggle('visible', id === t);
    document.querySelectorAll('.tab')[i].classList.toggle('active', id === t);
  });
  if (t === 'history') renderHistory();
}

// ── Action change — show/hide Reply To ID block ──────────────────
function onActionChange() {
  var a = document.getElementById('action').value;
  document.getElementById('replyBlock').style.display =
    (a === 'Reply' || a === 'Forward') ? 'block' : 'none';
}

// ── From Name → auto-fill From Email ────────────────────────────
// Uses getAttribute('data-emp-email') because .dataset is stripped
// by the GAS Caja sandbox on dynamically created <option> nodes.
function onNameChange() {
  var selectedName = document.getElementById('fromName').value;
  var emailSel = document.getElementById('fromEmail');
  for (var i = 0; i < emailSel.options.length; i++) {
    if (emailSel.options[i].getAttribute('data-emp-name') === selectedName) {
      emailSel.selectedIndex = i;
      return;
    }
  }
}

// ── Mention / Reply-To person ────────────────────────────────────
function onMentionChange() {
  var sel  = document.getElementById('mentionPerson');
  var opt  = sel.options[sel.selectedIndex];
  var note = document.getElementById('mentionNote');
  var ta   = document.getElementById('message');

  if (!opt || !opt.value) {
    // Remove any @mention prefix already in the textarea
    ta.value = ta.value.replace(/^@[^\s]+ ?/, '');
    pendingMention = '';
    note.style.display = 'none';
    return;
  }

  var name  = opt.getAttribute('data-name')  || opt.value;
  var email = opt.getAttribute('data-email') || '';
  pendingMention = '@' + name;

  // Show notification note
  note.style.display = 'block';
  note.innerHTML =
    '<span class="mention-badge">@' + name + '</span> will be mentioned'
    + (email ? ' &amp; notified at <strong>' + email + '</strong>' : '')
    + '.';

  // Inject @mention at the start of the message (replace existing if any)
  var current = ta.value.replace(/^@[^\s]+ ?/, '').trim();
  var preset  = PRESETS[document.getElementById('msgType').value] || '';
  var base    = (!current || current === preset) ? preset : current;
  ta.value    = pendingMention + ' ' + base;
  ta.focus();
  ta.setSelectionRange(ta.value.length, ta.value.length);
}

// ── Message type change — fill preset ───────────────────────────
function onTypeChange() {
  var t    = document.getElementById('msgType').value;
  var ta   = document.getElementById('message');
  var note = document.getElementById('customNote');
  var hint = document.getElementById('presetHint');
  var mentionPrefix = pendingMention ? (pendingMention + ' ') : '';

  if (t === 'Custom') {
    ta.value  = mentionPrefix;
    ta.readOnly = false;
    note.style.display = 'inline';
    hint.style.display = 'none';
    ta.focus();
  } else {
    ta.value  = mentionPrefix + PRESETS[t];
    ta.readOnly = !pendingMention; // editable when a mention is present
    note.style.display = 'none';
    hint.style.display = 'block';
    hint.textContent   = 'Pre-filled: "' + PRESETS[t] + '"';
  }
}

// ── Send ─────────────────────────────────────────────────────────
function doSend() {
  var btn = document.getElementById('sendBtn');
  var st  = document.getElementById('statusMsg');

  var channel = document.getElementById('channel').value;
  var msgText = document.getElementById('message').value.trim();
  var msgType = document.getElementById('msgType').value;

  var mentionSel = document.getElementById('mentionPerson');
  var mentionOpt = mentionSel.options[mentionSel.selectedIndex];

  if (!channel)                        { showStatus('err', 'Please select a channel.'); return; }
  if (!msgText && msgType === 'Custom'){ showStatus('err', 'Message cannot be empty.'); return; }

  var payload = {
    action:       document.getElementById('action').value,
    channel:      channel,
    fromName:     document.getElementById('fromName').value,
    fromEmail:    document.getElementById('fromEmail').value,
    msgType:      msgType,
    message:      msgText || PRESETS[msgType] || '',
    replyId:      (document.getElementById('replyId') || {value:''}).value || '',
    mentionName:  (mentionOpt && mentionOpt.value)
                    ? (mentionOpt.getAttribute('data-name')  || mentionOpt.value) : '',
    mentionEmail: (mentionOpt && mentionOpt.value)
                    ? (mentionOpt.getAttribute('data-email') || '') : ''
  };

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Sending…';
  showStatus('info', 'Sending…');

  google.script.run
    .withSuccessHandler(function(r) {
      btn.disabled = false;
      btn.textContent = 'Send Message';
      if (r && String(r).indexOf('Error') === 0) {
        showStatus('err', r);
      } else {
        showStatus('ok', r || 'Sent successfully.');
        saveToHistory(payload, r);
        // Reset volatile fields
        if (document.getElementById('replyId')) document.getElementById('replyId').value = '';
        document.getElementById('mentionPerson').selectedIndex = 0;
        document.getElementById('mentionNote').style.display = 'none';
        pendingMention = '';
        if (payload.msgType === 'Custom') { document.getElementById('message').value = ''; }
        else { onTypeChange(); }
      }
    })
    .withFailureHandler(function(e) {
      btn.disabled = false;
      btn.textContent = 'Send Message';
      showStatus('err', 'Error: ' + (e.message || e));
    })
    .sendFromComposeBox(payload);
}

function showStatus(cls, msg) {
  var el = document.getElementById('statusMsg');
  el.className  = cls;
  el.textContent = msg;
}

// ── History (localStorage, keyed by sender email) ────────────────
function _histKey() {
  var em = document.getElementById('fromEmail');
  var v  = (em && em.value) ? em.value.replace(/[^a-z0-9]/gi, '_') : 'shared';
  return 'funnelChat_hist_' + v;
}

function saveToHistory(payload, result) {
  try {
    var key  = _histKey();
    var hist = JSON.parse(localStorage.getItem(key) || '[]');
    hist.unshift({
      action:       payload.action,
      channel:      payload.channel,
      fromName:     payload.fromName,
      fromEmail:    payload.fromEmail,
      msgType:      payload.msgType,
      message:      payload.message,
      mentionName:  payload.mentionName  || '',
      mentionEmail: payload.mentionEmail || '',
      replyId:      payload.replyId || '',
      result:       result || '',
      time:         new Date().toLocaleString()
    });
    if (hist.length > 200) hist = hist.slice(0, 200);
    localStorage.setItem(key, JSON.stringify(hist));
  } catch(e) {}
}

function loadHistory() {
  try { return JSON.parse(localStorage.getItem(_histKey()) || '[]'); } catch(e) { return []; }
}

// ─── REAL-TIME CHAT VIA CACHE SERVICE ────────────────────
// Add these functions to your code.js in Apps Script

function writeChatMessage(payload) {
  try {
    if (!payload || !payload.text || !payload.channelId) return { success: false, error: 'Missing fields' };
    
    const msgId = payload.id || ('MSG-' + Utilities.getUuid().split('-')[0].toUpperCase());
    const now = payload.time || new Date().toISOString();
    
    // 1. Write to sheet (permanent storage)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('CHAT_MESSAGES');
    if (!sh) {
      sh = ss.insertSheet('CHAT_MESSAGES');
      sh.appendRow(['id','timestamp','channelId','fromEmail','fromName','text','type','source']);
      sh.getRange(1,1,1,8).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([msgId, now, payload.channelId, payload.fromEmail || '', payload.fromName || '', payload.text, payload.type || 'text', payload.source || 'chat-ui']);
    
    // 2. Also store in Cache for fast retrieval (lasts 6 hours)
    try {
      const cache = CacheService.getScriptCache();
      const cacheKey = 'ch_' + payload.channelId;
      const existing = cache.get(cacheKey);
      const msgs = existing ? JSON.parse(existing) : [];
      msgs.push({ id: msgId, time: now, fromEmail: payload.fromEmail, fromName: payload.fromName, text: payload.text });
      // Keep only last 50 messages in cache
      const trimmed = msgs.slice(-50);
      cache.put(cacheKey, JSON.stringify(trimmed), 21600); // 6 hours
    } catch(cacheErr) {
      Logger.log('Cache write error: ' + cacheErr.message);
    }
    
    // 3. Update channel last message
    try { _updateChannelLastMessage(payload.channelId, payload.text, payload.fromName); } catch(e) {}
    
    return { success: true, msgId: msgId };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getChat(channelId, leadId, limit) {
  try {
    const lim = parseInt(limit) || 50;
    
    // Try cache first (fast path — milliseconds)
    try {
      const cache = CacheService.getScriptCache();
      const cached = cache.get('ch_' + channelId);
      if (cached) {
        const msgs = JSON.parse(cached);
        return msgs.slice(-lim).map(m => ({
          id: m.id, time: m.time, channelId: channelId,
          fromEmail: m.fromEmail, fromName: m.fromName,
          text: m.text, type: 'text'
        }));
      }
    } catch(cacheErr) {}
    
    // Fallback: read from sheet
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CHAT_MESSAGES');
    if (!sh || sh.getLastRow() < 2) return [];
    
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    
    const rows = data.slice(1).filter(r => {
      if (!r[0]) return false;
      if (channelId && String(r[2]) !== String(channelId)) return false;
      return true;
    });
    
    return rows.slice(-lim).map(r => ({
      id: String(r[0]), time: String(r[1]), channelId: String(r[2]),
      fromEmail: String(r[3]), fromName: String(r[4]),
      text: String(r[5]), type: String(r[6] || 'text')
    }));
  } catch(e) {
    Logger.log('getChat error: ' + e.message);
    return [];
  }
}

function clearHistory() {
  if (!confirm('Clear all chat history for this user?')) return;
  try { localStorage.removeItem(_histKey()); } catch(e) {}
  renderHistory();
}

function refreshHistory() {
  renderHistory();
  showStatus('ok', 'History refreshed.');
  setTimeout(function() {
    var el = document.getElementById('statusMsg');
    if (el.textContent === 'History refreshed.') el.className = '';
  }, 1800);
}

function renderHistory() {
  var q    = (document.getElementById('histSearch').value || '').toLowerCase();
  var hist = loadHistory();
  var el   = document.getElementById('historyList');
  if (!hist.length) {
    el.innerHTML = '<div class="no-hist">No history yet.</div>';
    return;
  }
  var filtered = hist.filter(function(h) {
    return !q || (h.message + h.channel + h.fromName + h.msgType + (h.mentionName || ''))
                  .toLowerCase().indexOf(q) >= 0;
  });
  if (!filtered.length) { el.innerHTML = '<div class="no-hist">No results.</div>'; return; }
  el.innerHTML = filtered.map(function(h) {
    var badge = h.mentionName
      ? '<span class="mention-badge">@' + h.mentionName + '</span> &bull; ' : '';
    return '<div class="hitem">'
      + '<span class="hch">#' + h.channel + '</span>&nbsp; ' + h.action
      + (h.replyId ? ' <span style="color:#94A3B8;font-size:10px">↩ ' + h.replyId + '</span>' : '')
      + '<div class="htxt">' + h.message + '</div>'
      + '<div class="hmeta">' + badge + h.fromName + ' &bull; ' + h.msgType + ' &bull; ' + h.time + '</div>'
      + '</div>';
  }).join('');
}

// ── Bootstrap: load channels + employees ─────────────────────────
// CRITICAL FIX: use setAttribute() not .dataset for <option> nodes;
// GAS's Caja sandbox strips .dataset on dynamically created elements.
google.script.run
  .withSuccessHandler(function(d) {
    document.getElementById('loadingOverlay').style.display = 'none';

    // ── Channels ──
    var cs = document.getElementById('channel');
    cs.innerHTML = '<option value="">Select channel…</option>';
    (d.channels || []).forEach(function(c) {
      var o = document.createElement('option');
      o.value = c;
      o.text  = c;
      cs.appendChild(o);
    });
    cs.disabled = false;

    // ── Employees ──
    var ns = document.getElementById('fromName');
    var es = document.getElementById('fromEmail');
    var ms = document.getElementById('mentionPerson');
    ns.innerHTML = '<option value="">Select sender…</option>';
    es.innerHTML = '<option value="">Select email…</option>';
    // mentionPerson already has the "No mention" placeholder in HTML

    empData = d.employees || [];
    empData.forEach(function(e) {
      // From Name option
      var on = document.createElement('option');
      on.value = e.name;
      on.text  = e.name;
      ns.appendChild(on);

      // From Email option — use setAttribute for data attributes (GAS sandbox fix)
      var oe = document.createElement('option');
      oe.value = e.email;
      oe.text  = e.email;
      oe.setAttribute('data-emp-name', e.name);
      es.appendChild(oe);

      // Mention option
      var om = document.createElement('option');
      om.value = e.email || e.name;
      om.text  = e.name + (e.email ? ' (' + e.email + ')' : '');
      om.setAttribute('data-name',  e.name);
      om.setAttribute('data-email', e.email || '');
      ms.appendChild(om);
    });

    ns.disabled = false;
    es.disabled = false;
  })
  .withFailureHandler(function(err) {
    document.getElementById('loadingOverlay').style.display = 'none';
    showStatus('err', 'Could not load data: ' + (err.message || err));
  })
  .getComposeBoxData();
</script>
</body>
</html>`)
    .setTitle('Chat Compose')
    .setWidth(370);

  SpreadsheetApp.getUi().showSidebar(html);
}

// ── getComposeBoxData ────────────────────────────────────────────
// Called by the sidebar bootstrap. Returns channels (names array)
// and employees ({name, email} array).
function getComposeBoxData() {
  return {
    channels:  _getChatWorkflowChannelList(),
    employees: _getChatWorkflowEmployeeDataForCompose()
  };
}

// ── sendFromComposeBox ───────────────────────────────────────────
// Handles preset messages, @mention prepend, direct sheet write,
// mention notification (NOTIFICATIONS sheet + email), and logging.
function sendFromComposeBox(payload) {
  try {
    var PRESETS = {
      'Acknowledgement':   'Acknowledged. Thank you.',
      'Not Acknowledged':  'This has not been acknowledged yet. Please confirm.',
      'Forward Required':  'This message requires forwarding to the relevant team.',
      'Needs More Detail': 'Please provide additional detail on this matter.'
    };

    // Resolve final message text
    var message = (payload.message || '').trim();
    if (!message && payload.msgType && payload.msgType !== 'Custom') {
      message = PRESETS[payload.msgType] || payload.msgType;
    }
    if (!message) return 'Error: No message to send.';
    if (!payload.channel) return 'Error: No channel selected.';

    // Prepend @mention if one was chosen and it isn't already present
    var mentionName  = (payload.mentionName  || '').trim();
    var mentionEmail = (payload.mentionEmail || '').trim();
    if (mentionName && message.indexOf('@' + mentionName) !== 0) {
      message = '@' + mentionName + ' ' + message;
    }

    var channelId   = _resolveChannelId(payload.channel);
    var senderEmail = (payload.fromEmail || _getEmail() || '').trim();
    var senderName  = (payload.fromName  || senderEmail).trim();
    var msgId       = 'MSG-' + Utilities.getUuid().split('-')[0].toUpperCase();
    var now         = new Date();

    // ── 1. Write to CHAT_MESSAGES (primary delivery) ──────────────
    _writeToChatMessages({
      id:        msgId,
      channelId: channelId,
      fromEmail: senderEmail,
      fromName:  senderName,
      text:      message,
      type:      'text',
      threadId:  payload.replyId || '',
      time:      now.toISOString(),
      source:    'compose-box'
    });

    // ── 2. Best-effort sendChat hook ──────────────────────────────
    try {
      sendChat(senderEmail, null, message, channelId, {
        fromName: senderName,
        type:     'text',
        threadId: payload.replyId || '',
        msgId:    msgId
      });
    } catch(ignored) {}

    // ── 3. Mention notification ───────────────────────────────────
    if (mentionName) {
      _sendMentionNotification({
        mentionName:  mentionName,
        mentionEmail: mentionEmail,
        senderName:   senderName,
        senderEmail:  senderEmail,
        channel:      payload.channel,
        messageSnip:  message.slice(0, 120),
        msgId:        msgId,
        time:         now.toISOString()
      });
    }

    // ── 4. Log to CHAT_WORKFLOW sheet ─────────────────────────────
    _logSentMessageToSheet(payload, message, msgId);

    var confirmation = 'Sent. ID: ' + msgId;
    if (mentionName) confirmation += ' | @' + mentionName + ' notified.';
    return confirmation;

  } catch(e) {
    return 'Error: ' + e.message;
  }
}

function getLeadActivity(leadId) {
  try {
    if (!leadId) return [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Primary: read from LEAD_ACTIVITY sheet
    const results = [];
    const laSh = ss.getSheetByName('LEAD_ACTIVITY');
    if (laSh && laSh.getLastRow() >= 2) {
      const numCols = Math.max(laSh.getLastColumn(), 9);
      const data = laSh.getRange(2, 1, laSh.getLastRow() - 1, numCols).getValues();
      data.forEach(function(row) {
        if (String(row[0]).trim() !== String(leadId).trim()) return;
        results.push({
          leadId:    String(row[0]),
          type:      String(row[1] || 'note'),
          text:      String(row[2] || ''),
          by:        String(row[3] || 'Agent'),
          email:     String(row[4] || ''),
          icon:      String(row[5] || ''),
          color:     String(row[6] || ''),
          time:      row[8] ? new Date(row[8]).toISOString() : (row[7] ? new Date(row[7]).toISOString() : new Date().toISOString()),
          source:    'sheet'
        });
      });
    }

    // Secondary: read from ACTIVITY_LOG sheet for this leadId
    const alSh = ss.getSheetByName(SN.ACTIVITY);
    if (alSh && alSh.getLastRow() >= 2) {
      const data = alSh.getDataRange().getValues();
      const h = data[0];
      const lidIdx  = h.indexOf('Lead ID');
      const tsIdx   = h.indexOf('Timestamp');
      const emIdx   = h.indexOf('Employee Email');
      const enIdx   = h.indexOf('Employee Name');
      const actIdx  = h.indexOf('Activity Type');
      const detIdx  = h.indexOf('Details');
      if (lidIdx >= 0) {
        data.slice(1).forEach(function(row) {
          if (String(row[lidIdx]).trim() !== String(leadId).trim()) return;
          const ts = row[tsIdx] ? new Date(row[tsIdx]).toISOString() : new Date().toISOString();
          const key = String(row[actIdx]||'')+'|'+ts.slice(0,16)+'|'+String(row[detIdx]||'').slice(0,20);
          results.push({
            leadId: String(leadId),
            type:   _actTypeFromAction(String(row[actIdx]||'')),
            text:   String(row[detIdx] || ''),
            by:     String(row[enIdx]  || row[emIdx] || 'Agent'),
            email:  String(row[emIdx]  || ''),
            time:   ts,
            source: 'activity_log'
          });
        });
      }
    }

    // Tertiary: read embedded JSON from MASTER_LEADS Activity Log column
    const mSh = _getSheet(SN.LEADS);
    if (mSh && mSh.getLastRow() >= 2) {
      const actCol = _getActivityColIndex(mSh);
      const ids    = mSh.getRange(2, 1, mSh.getLastRow()-1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) !== String(leadId)) continue;
        let acts = [];
        try { acts = JSON.parse(String(mSh.getRange(i+2, actCol).getValue()||'[]')); } catch(e) { acts = []; }
        if (!Array.isArray(acts)) acts = [];
        acts.forEach(function(a) {
          results.push({
            leadId: String(leadId),
            type:   _actTypeFromAction(String(a.act||a.type||'note')),
            text:   String(a.det||a.text||''),
            by:     String(a.by||'Agent'),
            email:  String(a.em||''),
            time:   String(a.t||new Date().toISOString()),
            source: 'master_col'
          });
        });
        break;
      }
    }

    // Deduplicate
    const seen = {};
    const deduped = results.filter(function(a) {
      const k = a.type+'|'+a.time.slice(0,16)+'|'+a.text.slice(0,20);
      if (seen[k]) return false;
      seen[k] = true;
      return true;
    });

    deduped.sort(function(a,b){ return new Date(b.time) - new Date(a.time); });
    return deduped;

  } catch(e) {
    Logger.log('getLeadActivity error: ' + e.message);
    return [];
  }
}

// ══════════════════════════════════════════════════════════════════
//  USER SETTINGS — persist per-user preferences cross-device
// ══════════════════════════════════════════════════════════════════
function saveUserSettings(payload) {
  try {
    var email    = String(payload.email    || '').trim().toLowerCase();
    var settings = String(payload.settings || '{}');
    if (!email) return { success:false, error:'No email' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('USER_SETTINGS');
    if (!sh) {
      sh = ss.insertSheet('USER_SETTINGS');
      sh.getRange(1,1,1,4).setValues([['Email','Settings JSON','Updated At','Version']])
        .setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(1,200); sh.setColumnWidth(2,600); sh.setColumnWidth(3,180);
    }
    var lastRow = sh.getLastRow();
    var found   = false;
    if (lastRow >= 2) {
      var emails = sh.getRange(2,1,lastRow-1,1).getValues();
      for (var i=0; i<emails.length; i++) {
        if (String(emails[i][0]).toLowerCase().trim() === email) {
          sh.getRange(i+2,2,1,3).setValues([[settings, new Date().toISOString(), '1']]);
          found = true; break;
        }
      }
    }
    if (!found) sh.appendRow([email, settings, new Date().toISOString(), '1']);
    try { SpreadsheetApp.flush(); } catch(e) {}
    return { success:true };
  } catch(e) {
    Logger.log('saveUserSettings error: ' + e.message);
    return { success:false, error:e.message };
  }
}

function getUserSettings(email) {
  try {
    email = String(email||'').trim().toLowerCase();
    if (!email) return null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('USER_SETTINGS');
    if (!sh || sh.getLastRow() < 2) return null;
    var lastRow = sh.getLastRow();
    var data    = sh.getRange(2,1,lastRow-1,2).getValues();
    for (var i=0; i<data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === email) {
        var raw = String(data[i][1]||'');
        return raw || null;
      }
    }
    return null;
  } catch(e) {
    Logger.log('getUserSettings error: ' + e.message);
    return null;
  }
}

// ══════════════════════════════════════════════════════════════════
//  USER SETTINGS — persist per-user preferences cross-device
// ══════════════════════════════════════════════════════════════════
function saveUserSettings(payload) {
  try {
    var email    = String(payload.email    || '').trim().toLowerCase();
    var settings = String(payload.settings || '{}');
    if (!email) return { success:false, error:'No email' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('USER_SETTINGS');
    if (!sh) {
      sh = ss.insertSheet('USER_SETTINGS');
      sh.getRange(1,1,1,4).setValues([['Email','Settings JSON','Updated At','Version']])
        .setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(1,200); sh.setColumnWidth(2,600); sh.setColumnWidth(3,180);
    }
    var lastRow = sh.getLastRow();
    var found   = false;
    if (lastRow >= 2) {
      var emails = sh.getRange(2,1,lastRow-1,1).getValues();
      for (var i=0; i<emails.length; i++) {
        if (String(emails[i][0]).toLowerCase().trim() === email) {
          sh.getRange(i+2,2,1,3).setValues([[settings, new Date().toISOString(), '1']]);
          found = true; break;
        }
      }
    }
    if (!found) sh.appendRow([email, settings, new Date().toISOString(), '1']);
    try { SpreadsheetApp.flush(); } catch(e) {}
    return { success:true };
  } catch(e) {
    Logger.log('saveUserSettings error: ' + e.message);
    return { success:false, error:e.message };
  }
}

function getUserSettings(email) {
  try {
    email = String(email||'').trim().toLowerCase();
    if (!email) return null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('USER_SETTINGS');
    if (!sh || sh.getLastRow() < 2) return null;
    var lastRow = sh.getLastRow();
    var data    = sh.getRange(2,1,lastRow-1,2).getValues();
    for (var i=0; i<data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === email) {
        var raw = String(data[i][1]||'');
        return raw || null;
      }
    }
    return null;
  } catch(e) {
    Logger.log('getUserSettings error: ' + e.message);
    return null;
  }
}

function _actTypeFromAction(action) {
  const s = String(action).toLowerCase();
  if (s.includes('call'))     return 'call';
  if (s.includes('email'))    return 'email';
  if (s.includes('whatsapp')) return 'whatsapp';
  if (s.includes('meeting'))  return 'meeting';
  if (s.includes('note'))     return 'note';
  if (s.includes('ticket'))   return 'ticket';
  if (s.includes('status'))   return 'status';
  if (s.includes('funnel') || s.includes('stage')) return 'funnel';
  if (s.includes('workflow')) return 'workflow';
  if (s.includes('forward'))  return 'forward';
  if (s.includes('quotation') || s.includes('quote')) return 'quotation';
  if (s.includes('ai') || s.includes('pitch')) return 'ai_assist';
  if (s.includes('open') || s.includes('view')) return 'open';
  if (s.includes('hot'))      return 'status';
  return action.toLowerCase().replace(/_/g,'');
}

function _ensureLeadActivitySheet(ss) {
  let sh = ss.getSheetByName('LEAD_ACTIVITY');
  if (sh) return sh;
  sh = ss.insertSheet('LEAD_ACTIVITY');
  const headers = ['Lead ID','Activity Type','Text / Details','By (Name)','By (Email)','Icon','Color','Extra JSON','Timestamp'];
  sh.getRange(1,1,1,headers.length).setValues([headers])
    .setBackground('#111827').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1,100); sh.setColumnWidth(2,160); sh.setColumnWidth(3,320);
  sh.setColumnWidth(4,140); sh.setColumnWidth(5,200); sh.setColumnWidth(9,180);
  return sh;
}

// ── _sendMentionNotification ─────────────────────────────────────
// Writes a row to NOTIFICATIONS sheet and sends a GmailApp email
// if the mentioned person has a valid email address.
function _sendMentionNotification(opts) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('NOTIFICATIONS');
    if (!sh) {
      sh = ss.insertSheet('NOTIFICATIONS');
      sh.getRange(1, 1, 1, 7).setValues([[
        'Time', 'To Name', 'To Email', 'From Name', 'Channel', 'Message', 'Msg ID'
      ]]);
      sh.getRange(1, 1, 1, 7)
        .setBackground('#111827').setFontColor('#FFFFFF').setFontWeight('bold');
      sh.setFrozenRows(1);
    }

    var lastRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(lastRow, 1, 1, 7).setValues([[
      opts.time,
      opts.mentionName,
      opts.mentionEmail,
      opts.senderName,
      opts.channel,
      opts.messageSnip,
      opts.msgId
    ]]);

    // Send email notification (best-effort)
    if (opts.mentionEmail && opts.mentionEmail.indexOf('@') > 0) {
      var subject = opts.senderName + ' mentioned you in #' + opts.channel + ' — Nucleous';
      var body    = 'Hi ' + opts.mentionName + ',\n\n'
        + opts.senderName + ' mentioned you in the #' + opts.channel + ' channel:\n\n'
        + '"' + opts.messageSnip + '"\n\n'
        + 'Message ID: ' + opts.msgId + '\n'
        + 'Time: ' + new Date(opts.time).toLocaleString() + '\n\n'
        + '— Nucleous by AfterResult';
      try {
        GmailApp.sendEmail(opts.mentionEmail, subject, body);
      } catch(emailErr) {
        Logger.log('Mention email failed: ' + emailErr.message);
      }
    }
  } catch(e) {
    Logger.log('_sendMentionNotification error: ' + e.message);
  }
}

function _getChatWorkflowEmployeeDataForCompose() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('EMPLOYEES');
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    const header = data[0].map(h => String(h).toLowerCase());
    const nameCol  = header.indexOf('name')  !== -1 ? header.indexOf('name')  : 0;
    const emailCol = header.indexOf('email') !== -1 ? header.indexOf('email') : 1;
    return data.slice(1)
      .map(r => ({ name: String(r[nameCol] || '').trim(), email: String(r[emailCol] || '').trim() }))
      .filter(e => e.name);
  } catch(e) {
    return [];
  }
}

function sendFromComposeBox(payload) {
  try {
    const PRESETS = {
      'Acknowledgement':  'Acknowledged. Thank you.',
      'Not Acknowledged': 'This has not been acknowledged yet. Please confirm.',
      'Forward Required': 'This message requires forwarding to the relevant team.',
      'Needs More Detail':'Please provide additional detail on this matter.'
    };

    const message = (payload.msgType && payload.msgType !== 'Custom' && !payload.message)
      ? (PRESETS[payload.msgType] || payload.msgType)
      : (payload.message || '').trim();

    if (!message) return 'Error: No message to send.';
    if (!payload.channel) return 'Error: No channel selected.';

    const channelId = _resolveChannelId(payload.channel);
    const msgId = 'MSG-' + Utilities.getUuid().split('-')[0].toUpperCase();
    const senderEmail = payload.fromEmail || _getEmail();
    const senderName  = payload.fromName  || senderEmail;
    const now = new Date();

    // Write directly to CHAT_MESSAGES sheet — this is the source of truth the UI reads
    _writeToChatMessages({
      id:          msgId,
      channelId:   channelId,
      fromEmail:   senderEmail,
      fromName:    senderName,
      text:        message,
      type:        'text',
      threadId:    payload.replyId || '',
      time:        now.toISOString(),
      source:      'workflow'
    });

    // Also attempt sendChat if it exists (best-effort; errors are caught)
    try {
      sendChat(senderEmail, null, message, channelId, {
        fromName: senderName,
        type: 'text',
        threadId: payload.replyId || '',
        msgId: msgId
      });
    } catch(ignored) {}

    // Log to CHAT_WORKFLOW sheet for record
    _logSentMessageToSheet(payload, message, msgId);

    return 'Sent. Message ID: ' + msgId;
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

function _writeToChatMessages(msg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('CHAT_MESSAGES');
  if (!sh) return;

  // Detect column order from header row
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                   .map(h => String(h).toLowerCase().trim());

  const col = (name) => {
    const idx = header.indexOf(name);
    return idx >= 0 ? idx : null;
  };

  // Build a row array aligned to actual headers
  const rowSize = header.length || 10;
  const row = new Array(rowSize).fill('');

  const set = (name, val) => { const i = col(name); if (i !== null) row[i] = val; };

  set('id',         msg.id);
  set('messageid',  msg.id);
  set('message_id', msg.id);
  set('channelid',  msg.channelId);
  set('channel_id', msg.channelId);
  set('channel',    msg.channelId);
  set('fromemail',  msg.fromEmail);
  set('from_email', msg.fromEmail);
  set('email',      msg.fromEmail);
  set('fromname',   msg.fromName);
  set('from_name',  msg.fromName);
  set('name',       msg.fromName);
  set('text',       msg.text);
  set('message',    msg.text);
  set('type',       msg.type || 'text');
  set('threadid',   msg.threadId || '');
  set('thread_id',  msg.threadId || '');
  set('time',       msg.time);
  set('timestamp',  msg.time);
  set('source',     msg.source || 'workflow');
  set('status',     'sent');

  sh.appendRow(row);

  // Update last message on CHANNELS sheet
  try { _updateChannelLastMessage(msg.channelId, msg.text, msg.fromName); } catch(e) {}
}

function _logSentMessageToSheet(payload, message, msgId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('CHAT_WORKFLOW');
    if (!sh) return;
    const lastRow = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(lastRow, 1, 1, 9).setValues([[
      payload.action,
      payload.channel,
      payload.fromName,
      payload.fromEmail,
      payload.msgType,
      message,
      'Sent',
      payload.replyId || '',
      msgId
    ]]);
  } catch(e) {}
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== 'CHAT_WORKFLOW') return;

    const col = e.range.getColumn();
    const row = e.range.getRow();
    if (row < 2) return;

    // When From Name (col 3) changes, auto-fill From Email (col 4)
    if (col === 3) {
      const name = String(e.value || '').trim();
      const email = _lookupEmployeeEmail(name);
      if (email) sh.getRange(row, 4).setValue(email);
    }

    // When Message Type (col 5) changes, pre-fill Message (col 6) if blank or was a preset
    if (col === 5) {
      const PRESETS = {
        'Acknowledgement':  'Acknowledged. Thank you.',
        'Not Acknowledged': 'This has not been acknowledged yet. Please confirm.',
        'Forward Required': 'This message needs to be forwarded to the relevant team.',
        'Needs More Detail':'Please provide more detail on this matter.',
        'Custom': ''
      };
      const type = String(e.value || '').trim();
      const existingMsg = String(sh.getRange(row, 6).getValue() || '').trim();
      const isCurrentlyAPreset = Object.values(PRESETS).includes(existingMsg);
      if (!existingMsg || isCurrentlyAPreset) {
        sh.getRange(row, 6).setValue(PRESETS[type] || '');
      }
    }

    // Auto-set Status to Pending when Action or Message changes
    if ((col === 1 || col === 6) && String(sh.getRange(row, 7).getValue()) !== 'Sent') {
      if (String(sh.getRange(row, 1).getValue()).trim()) {
        sh.getRange(row, 7).setValue('Pending');
      }
    }

    // Auto-generate Message ID (col 9) if row has Action set but no ID yet
    if (col === 1 || col === 5 || col === 6) {
      const hasAction = String(sh.getRange(row, 1).getValue()).trim();
      const hasId     = String(sh.getRange(row, 9).getValue()).trim();
      if (hasAction && !hasId) {
        sh.getRange(row, 9).setValue('MSG-' + Utilities.getUuid().split('-')[0].toUpperCase());
      }
    }
  } catch(e) {}
}

function _lookupEmployeeEmail(name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('EMPLOYEES');
    if (!sh) return '';
    const data = sh.getDataRange().getValues();
    const header = data[0].map(h => String(h).toLowerCase());
    const nameCol  = header.indexOf('name')  !== -1 ? header.indexOf('name')  : 0;
    const emailCol = header.indexOf('email') !== -1 ? header.indexOf('email') : 1;
    const match = data.slice(1).find(r => String(r[nameCol] || '').trim() === name);
    return match ? String(match[emailCol] || '').trim() : '';
  } catch(e) {
    return '';
  }
}
// Resolve channel name → channel ID
function _resolveChannelId(nameOrId) {
  try {
    const channels = getChannels();
    const match = channels.find(c =>
      (c.name || '').toLowerCase() === nameOrId.toLowerCase() ||
      (c.id || '').toLowerCase() === nameOrId.toLowerCase()
    );
    return match ? match.id : nameOrId.toLowerCase().replace(/[^a-z0-9\-]/g, '-');
  } catch(e) {
    return nameOrId;
  }
}

// ═══════════════════════════════════════════════════════════════
//  ROUTING RULES — SERVER SIDE ENGINE
// ═══════════════════════════════════════════════════════════════

function saveRoutingRule(rule) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('ROUTING_RULES');
    if (!sh) {
      sh = ss.insertSheet('ROUTING_RULES');
      sh.appendRow([
        'Rule ID','Name','Trigger','Trigger Value',
        'Target Emp ID','Target Emp Name','Target Emp Email',
        'Override','Keep Copy','Notify','Active','Fire Count','Created At','Last Fired'
      ]);
      sh.getRange(1,1,1,14).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(rule.id)) {
        sh.getRange(i+1,1,1,14).setValues([[
          rule.id, rule.name, rule.trigger, rule.triggerValue||'',
          rule.empId, rule.empName, rule.empEmail,
          rule.override ? 'Yes':'No',
          rule.keepCopy ? 'Yes':'No',
          rule.notify   ? 'Yes':'No',
          rule.active   ? 'Yes':'No',
          rule.fireCount||0, rule.createdAt||new Date().toISOString(), ''
        ]]);
        return { success:true };
      }
    }
    sh.appendRow([
      rule.id, rule.name, rule.trigger, rule.triggerValue||'',
      rule.empId, rule.empName, rule.empEmail,
      rule.override ? 'Yes':'No',
      rule.keepCopy ? 'Yes':'No',
      rule.notify   ? 'Yes':'No',
      rule.active   ? 'Yes':'No',
      rule.fireCount||0, rule.createdAt||new Date().toISOString(), ''
    ]);
    return { success:true };
  } catch(e) {
    return { success:false, message:e.message };
  }
}

function getRoutingRules() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('ROUTING_RULES');
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(r => r[0] && r[10] === 'Yes')
      .map(r => ({
        id:           String(r[0]),
        name:         String(r[1]),
        trigger:      String(r[2]),
        triggerValue: String(r[3]),
        empId:        String(r[4]),
        empName:      String(r[5]),
        empEmail:     String(r[6]),
        override:     r[7] === 'Yes',
        keepCopy:     r[8] === 'Yes',
        notify:       r[9] === 'Yes',
        active:       r[10] === 'Yes',
        fireCount:    Number(r[11]||0)
      }));
  } catch(e) { return []; }
}

function deleteRoutingRule(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('ROUTING_RULES');
    if (!sh) return { success:false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sh.deleteRow(i+1);
        return { success:true };
      }
    }
    return { success:false };
  } catch(e) { return { success:false }; }
}

function toggleRoutingRuleServer(id, active) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('ROUTING_RULES');
    if (!sh) return { success:false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sh.getRange(i+1, 11).setValue(active ? 'Yes' : 'No');
        return { success:true };
      }
    }
    return { success:false };
  } catch(e) { return { success:false }; }
}

// ─── CORE ENGINE — fires on every lead sheet edit ─────────────────
// This is called by the onEdit installable trigger below.
// It reads ROUTING_RULES sheet and reassigns leads automatically.
function checkAndApplyRoutingRules(leadRow, headers, sheet) {
  try {
    const rules = getRoutingRules();
    if (!rules.length) return;

    const col = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]).toLowerCase().replace(/[\s_\-]/g,'');
        for (const k of (Array.isArray(keywords)?keywords:[keywords])) {
          if (h.includes(k.toLowerCase().replace(/[\s_\-]/g,''))) return i;
        }
      }
      return -1;
    };

    const hotCol      = col(['hot','ishot','hotlead']);
    const statusCol   = col(['status','leadstatus']);
    const priorityCol = col(['priority']);
    const industryCol = col(['industry','service','category']);
    const assignedToCol    = col(['assignedto','assigned_to','assigned']);
    const assignedEmailCol = col(['assignedemail','assigned_email']);
    const assignedEmpIdCol = col(['assignedempid','assigned_empid']);
    const leadIdCol   = col(['leadid','lead_id','id']);

    if (assignedToCol === -1) return;

    const hotVal      = hotCol      >= 0 ? String(leadRow[hotCol]||'').toLowerCase()      : '';
    const statusVal   = statusCol   >= 0 ? String(leadRow[statusCol]||'')                 : '';
    const priorityVal = priorityCol >= 0 ? String(leadRow[priorityCol]||'')               : '';
    const industryVal = industryCol >= 0 ? String(leadRow[industryCol]||'')               : '';
    const leadId      = leadIdCol   >= 0 ? String(leadRow[leadIdCol]||'')                 : '';

    for (const rule of rules) {
      let matches = false;
      if (rule.trigger === 'hot'      && (hotVal === 'yes' || hotVal === 'true' || hotVal === '1')) matches = true;
      if (rule.trigger === 'status'   && statusVal   === rule.triggerValue) matches = true;
      if (rule.trigger === 'priority' && priorityVal === rule.triggerValue) matches = true;
      if (rule.trigger === 'industry' && industryVal === rule.triggerValue) matches = true;

      if (!matches) continue;

      // Check target employee exists and is active
      const empSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMPLOYEES');
      if (!empSh) continue;
      const empData = empSh.getDataRange().getValues();
      let targetEmp = null;
      for (let ei = 1; ei < empData.length; ei++) {
        if (String(empData[ei][0]) === rule.empId ||
            String(empData[ei][2]).toLowerCase() === rule.empEmail.toLowerCase()) {
          if (String(empData[ei][10]||'Active') === 'Active') {
            targetEmp = { id:String(empData[ei][0]), name:String(empData[ei][1]), email:String(empData[ei][2]) };
          }
          break;
        }
      }
      if (!targetEmp) continue;

      // Don't re-route if already assigned to this target
      const curEmail = assignedEmailCol >= 0 ? String(leadRow[assignedEmailCol]||'') : '';
      const curTo    = String(leadRow[assignedToCol]||'');
      if (curEmail.toLowerCase() === targetEmp.email.toLowerCase()) continue;
      if (curTo.toLowerCase()    === targetEmp.email.toLowerCase()) continue;

      // Find the row number in the sheet — search by lead ID
      const allData = sheet.getDataRange().getValues();
      let sheetRowIndex = -1;
      for (let ri = 1; ri < allData.length; ri++) {
        const rowId = leadIdCol >= 0 ? String(allData[ri][leadIdCol]||'') : '';
        if (rowId === leadId) { sheetRowIndex = ri + 1; break; }
      }
      if (sheetRowIndex === -1) continue;

      // Reassign
      sheet.getRange(sheetRowIndex, assignedToCol + 1).setValue(targetEmp.email);
      if (assignedEmailCol >= 0) sheet.getRange(sheetRowIndex, assignedEmailCol + 1).setValue(targetEmp.email);
      if (assignedEmpIdCol >= 0) sheet.getRange(sheetRowIndex, assignedEmpIdCol + 1).setValue(targetEmp.id);

      // Update fire count
      const rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ROUTING_RULES');
      if (rulesSheet) {
        const rData = rulesSheet.getDataRange().getValues();
        for (let ri = 1; ri < rData.length; ri++) {
          if (String(rData[ri][0]) === rule.id) {
            rulesSheet.getRange(ri+1, 12).setValue((Number(rData[ri][11])||0) + 1);
            rulesSheet.getRange(ri+1, 14).setValue(new Date().toISOString());
            break;
          }
        }
      }

      // Notify escalations channel
      if (rule.notify) {
        const msg = '[AUTO-ROUTE] Lead "' + (leadRow[col(['name','fullname','contactname'])]||leadId)
          + '" triggered rule "' + rule.name + '" ('
          + rule.trigger + (rule.triggerValue ? ': ' + rule.triggerValue : '')
          + '). Reassigned to ' + targetEmp.name + '.';
        try {
          sendChat('routing@system', null, msg, 'escalations', {
            type: 'system', fromName: 'Auto-Router'
          });
        } catch(e) {}
      }

      SpreadsheetApp.flush();
      Logger.log('[ROUTING] Rule "' + rule.name + '" fired → lead ' + leadId + ' reassigned to ' + targetEmp.name);
      break; // only first matching rule fires per edit
    }
  } catch(e) {
    Logger.log('[ROUTING] Error: ' + e.message);
  }
}

// ─── INSTALLABLE onEdit TRIGGER ───────────────────────────────────
// Run installLeadEditTrigger() ONCE from Apps Script editor to set it up.
// After that it fires automatically on every sheet edit.
function onLeadSheetEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName().toLowerCase();

    // Only watch the leads sheet
    if (!sheetName.includes('lead') && !sheetName.includes('master')) return;

    const editedRow = e.range.getRow();
    if (editedRow < 2) return; // skip header

    const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData  = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    checkAndApplyRoutingRules(rowData, headers, sheet);
  } catch(err) {
    Logger.log('[onLeadSheetEdit] Error: ' + err.message);
  }
}

function installLeadEditTrigger() {
  // Delete existing to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onLeadSheetEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onLeadSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  return 'Trigger installed. Routing rules will now fire automatically on every lead edit.';
}

// Convenience: load general channel messages
function loadGeneralChannel() { return loadChannelMessagesIntoSheet('general', 50); }
function loadSalesChannel() { return loadChannelMessagesIntoSheet('sales', 50); }
function loadLeadsChannel() { return loadChannelMessagesIntoSheet('leads', 50); }
function loadEscalationsChannel() { return loadChannelMessagesIntoSheet('escalations', 50); }

// Create Sheet menu for easy access
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Funnel Chat')
      .addItem('Setup Chat Workflow Sheet', 'setupChatWorkflowSheet')
      .addItem('Open Compose Box', 'openComposeBox')
      .addSeparator()
      .addItem('Send Pending Messages', 'processChatWorkflowRows')
      .addSeparator()
      .addSubMenu(
        SpreadsheetApp.getUi().createMenu('Load Channel Messages')
          .addItem('Load general (last 50)',     'loadGeneralChannel')
          .addItem('Load sales (last 50)',        'loadSalesChannel')
          .addItem('Load leads (last 50)',        'loadLeadsChannel')
          .addItem('Load escalations (last 50)', 'loadEscalationsChannel')
      )
      .addSeparator()
      .addItem('Initialize All Sheets', 'initializeSheets')
      .addToUi();
  } catch(e) {}
}

function getChat(channelId, leadId, limit) {
  try {
    const sh = _getSheet(SN.CHAT);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    const lim = parseInt(limit) || 100;
    let rows = data.slice(1).filter(r => {
      if (!r[0]) return false;
      if (channelId) return String(r[2]) === String(channelId);
      if (leadId) return String(r[11]) === String(leadId);
      return false;
    });
    return rows.slice(-lim).map(r => ({
      id: r[0], time: r[1], channelId: r[2], channelName: r[3],
      fromEmail: r[4], fromName: r[5], toEmail: r[6],
      text: r[7], type: r[8], read: r[9],
      workflowId: r[10], leadId: r[11], quotationId: r[12],
      edited: r[13], threadId: r[16]
    }));
  } catch(e) {
    return [];
  }
}

function markMessagesRead(channelId, email) {
  try {
    const sh = _getSheet(SN.CHAT);
    if (!sh) return { success: true };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(channelId) && data[i][9] === 'No') {
        sh.getRange(i + 1, 10).setValue('Yes');
      }
    }
    return { success: true };
  } catch(e) { return { success: true }; }
}

// ═══════════════════════════════════════════════════════════════
//  CHANNELS — ADD / LIST / ARCHIVE
// ═══════════════════════════════════════════════════════════════

function getChannels() {
  try {
    const sh = _getSheet(SN.CHANNELS);
    if (!sh) return _defaultChannelList();
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return _defaultChannelList();
    return data.slice(1).filter(r => r[0]).map(r => ({
      id: r[0], name: r[1], type: r[2], description: r[3],
      status: r[4], createdBy: r[5], createdDate: r[6],
      memberCount: r[7], members: String(r[8] || '').split(',').filter(Boolean),
      messageCount: r[9], lastMessage: r[10], lastMessageBy: r[11],
      lastMessageTime: r[12], isArchived: r[13] === 'Yes',
      pinned: r[15] === 'Yes', category: r[16]
    }));
  } catch(e) {
    return _defaultChannelList();
  }
}

function addChannel(channelData) {
  try {
    const sh = _getSheet(SN.CHANNELS);
    if (!sh) return { success: false };
    const email = _getEmail();
    const channelId = channelData.id || _uid('CH');
    const existing = sh.getDataRange().getValues();
    const nameExists = existing.slice(1).some(r => String(r[1]).toLowerCase() === String(channelData.name).toLowerCase());
    if (nameExists) return { success: false, message: 'Channel name already exists' };
    sh.appendRow([
      channelId,
      channelData.name,
      channelData.type || 'Team Channel',
      channelData.description || '',
      'Active',
      email,
      _now(),
      (channelData.members || []).length || 0,
      (channelData.members || []).join(','),
      0, '', '', '', 'No', '', 'No',
      channelData.category || ''
    ]);
    // Post system message to new channel
    sendChat(email, null, 'Channel #' + channelData.name + ' created by ' + (_getEmployeeName(email) || email), channelId, { type: 'system', channelName: channelData.name });
    return { success: true, channelId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function archiveChannel(channelId) {
  try {
    const sh = _getSheet(SN.CHANNELS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(channelId)) {
        sh.getRange(i + 1, 5).setValue('Archived');
        sh.getRange(i + 1, 14).setValue('Yes');
        sh.getRange(i + 1, 15).setValue(_now());
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function _updateChannelLastMessage(channelId, message, fromName) {
  try {
    const sh = _getSheet(SN.CHANNELS);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(channelId)) {
        const prev = parseInt(data[i][9]) || 0;
        sh.getRange(i + 1, 10).setValue(prev + 1);
        sh.getRange(i + 1, 11).setValue(String(message).substring(0, 100));
        sh.getRange(i + 1, 12).setValue(fromName || '');
        sh.getRange(i + 1, 13).setValue(_now());
        return;
      }
    }
  } catch(e) {}
}

function _getChannelName(channelId) {
  try {
    const sh = _getSheet(SN.CHANNELS);
    if (!sh) return channelId;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(channelId)) return data[i][1];
    }
    return channelId;
  } catch(e) { return channelId; }
}

function _defaultChannelList() {
  return [
    { id:'zira-ai', name:'Zira AI', type:'Team Channel', status:'Active', description:'AI Assistant' },
    { id:'general', name:'general', type:'Team Channel', status:'Active', description:'General team channel' },
    { id:'sales', name:'sales', type:'Team Channel', status:'Active', description:'Sales discussions' },
    { id:'escalations', name:'escalations', type:'Team Channel', status:'Active', description:'Urgent issues' },
    { id:'leads', name:'leads', type:'Workflow', status:'Active', description:'Lead workflows' },
    { id:'meetings-ch', name:'meetings', type:'Workflow', status:'Active', description:'Meeting coordination' }
  ];
}

function _defaultChannels() {
  const now = _now();
  return [
    ['zira-ai','Zira AI','Team Channel','AI Sales Assistant','Active','system',now,0,'',0,'','','','No','','Yes','AI'],
    ['general','general','Team Channel','General team channel','Active','system',now,0,'',0,'','','','No','','No','General'],
    ['sales','sales','Team Channel','Sales discussions','Active','system',now,0,'',0,'','','','No','','No','Sales'],
    ['escalations','escalations','Team Channel','Urgent escalations','Active','system',now,0,'',0,'','','','No','','No','Support'],
    ['leads','leads','Workflow','Lead forwarding channel','Active','system',now,0,'',0,'','','','No','','No','Workflow'],
    ['meetings-ch','meetings','Workflow','Meeting coordination','Active','system',now,0,'',0,'','','','No','','No','Workflow']
  ];
}

// ═══════════════════════════════════════════════════════════════
//  WORKFLOWS — SAVE & TRACK
// ═══════════════════════════════════════════════════════════════

function saveWorkflow(d) {
  try {
    const sh = _getSheet(SN.WORKFLOWS);
    if (!sh) return { success: false };
    const email = _getEmail();
    const wfId = d.id || _uid('WF');
    const tags = Array.isArray(d.tags) ? d.tags.join(', ') : (d.tags || '');
    sh.appendRow([
      wfId,             // Workflow ID
      _now(),           // Created Date
      d.type || 'Forward Lead',
      d.leadId || '',
      d.leadName || '',
      d.leadCompany || '',
      d.leadIndustry || '',
      d.team || '',
      d.service || '',
      d.budget || '',
      d.duration || '',
      d.requirements || '',
      tags,
      d.priority || 'Medium',
      'Pending',
      d.submittedBy || email,
      '',               // Assigned To
      d.targetChannel || 'leads',
      '',               // Chat Message ID
      '',               // Quotation ID
      '',               // Completed Date
      '',               // Notes
      _now()            // Last Updated
    ]);
    // Post workflow message to channel
    const targetCh = d.targetChannel || 'leads';
    const wfMsg = _buildWorkflowMessage(wfId, d);
    const msgResult = sendChat(email, null, wfMsg, targetCh, {
      type: 'workflow',
      workflowId: wfId,
      leadId: d.leadId,
      fromName: d.submittedBy || email
    });
    // Link message ID back to workflow
    if (msgResult.success) {
      _updateWorkflowField(wfId, 19, msgResult.msgId);
    }
    _logAction({ leadId: d.leadId, action: 'WORKFLOW_CREATED', details: 'Team: ' + d.team + ' | Service: ' + d.service });
    return { success: true, wfId, msgId: msgResult.msgId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateWorkflowStatus(wfId, newStatus, notes) {
  try {
    const sh = _getSheet(SN.WORKFLOWS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(wfId)) {
        sh.getRange(i + 1, 15).setValue(newStatus);
        sh.getRange(i + 1, 23).setValue(_now());
        if (notes) sh.getRange(i + 1, 22).setValue(notes);
        if (newStatus === 'Completed') sh.getRange(i + 1, 21).setValue(_now());
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function getWorkflows(filters) {
  try {
    const sh = _getSheet(SN.WORKFLOWS);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    let rows = data.slice(1).filter(r => r[0]);
    if (filters) {
      if (filters.status) rows = rows.filter(r => r[14] === filters.status);
      if (filters.leadId) rows = rows.filter(r => String(r[3]) === String(filters.leadId));
      if (filters.team) rows = rows.filter(r => r[7] === filters.team);
    }
    return rows.map(r => ({
      id: r[0], createdDate: r[1], type: r[2], leadId: r[3],
      leadName: r[4], leadCompany: r[5], industry: r[6],
      team: r[7], service: r[8], budget: r[9], duration: r[10],
      requirements: r[11], tags: r[12], priority: r[13], status: r[14],
      submittedBy: r[15], assignedTo: r[16], channel: r[17],
      chatMsgId: r[18], quotationId: r[19], completedDate: r[20],
      notes: r[21], lastUpdated: r[22]
    })).reverse();
  } catch(e) { return []; }
}

function _buildWorkflowMessage(wfId, d) {
  const lines = [
    '━━━ WORKFLOW ' + wfId + ' ━━━',
    '👤 Lead: ' + (d.leadName || d.leadId || 'N/A') + (d.leadCompany ? ' (' + d.leadCompany + ')' : ''),
    '🏢 Industry: ' + (d.leadIndustry || 'N/A'),
    '📤 Forwarded to: ' + (d.team || 'N/A'),
    '🛠 Service: ' + (d.service || 'N/A'),
    '💰 Budget: ' + (d.budget ? '₹' + d.budget : 'TBD') + ' | ⏱ Duration: ' + (d.duration || 'TBD'),
    d.requirements ? '📋 Notes: ' + d.requirements : null,
    d.tags ? '🏷 Flags: ' + (Array.isArray(d.tags) ? d.tags.join(', ') : d.tags) : null,
    '👤 Submitted by: ' + (d.submittedBy || 'Team')
  ].filter(Boolean);
  return lines.join('\n');
}

function _updateWorkflowField(wfId, colIndex, value) {
  try {
    const sh = _getSheet(SN.WORKFLOWS);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(wfId)) {
        sh.getRange(i + 1, colIndex).setValue(value);
        return;
      }
    }
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════
//  QUOTATIONS — FULL TRACKING
// ═══════════════════════════════════════════════════════════════

function saveQuotation(d) {
  try {
    const sh = _getSheet(SN.QUOTATIONS);
    if (!sh) return { success: false };
    const email = _getEmail();
    const qtId = d.id || _uid('QT');
    const monthly = parseFloat(d.monthlyValue) || 0;
    const months = parseInt(d.durationMonths) || 1;
    const total = monthly * months;
    const discountPct = parseFloat(d.discount) || 0;
    const final = total * (1 - discountPct / 100);
    const gst = final * 0.18;
    const grand = final + gst;
    sh.appendRow([
      qtId,
      _now(),
      d.leadId || '',
      d.leadName || '',
      d.leadEmail || '',
      d.company || '',
      d.industry || '',
      d.service || '',
      monthly,
      months,
      total,
      discountPct,
      Math.round(final),
      Math.round(gst),
      Math.round(grand),
      'Draft',
      d.version || 1,
      '',               // Sent Date
      '',               // Valid Until
      '',               // Accepted Date
      '',               // Rejected Date
      '',               // Rejection Reason
      d.workflowId || '',
      email,
      d.notes || '',
      _now()
    ]);
    // Update lead quotation fields
    _updateLeadQuotationInfo(d.leadId, qtId, 'Draft');
    _logAction({ leadId: d.leadId, action: 'QUOTATION_CREATED', details: d.service + ' | Grand Total: ₹' + Math.round(grand) });
    return { success: true, qtId, grandTotal: Math.round(grand) };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function saveQuotationToSheet(q) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('QUOTATIONS');
    if (!sheet) {
      sheet = ss.insertSheet('QUOTATIONS');
      sheet.appendRow(['quoteNum','clientName','clientCompany','clientEmail','clientPhone','cartItems','discount','total','validTill','savedAt']);
    }
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(q.quoteNum)) {
        sheet.getRange(i+1, 1, 1, 10).setValues([[
          q.quoteNum, q.clientName, q.clientCompany, q.clientEmail, q.clientPhone,
          q.cartItems, q.discount, q.total, q.validTill, new Date().toISOString()
        ]]);
        return 'updated';
      }
    }
    sheet.appendRow([
      q.quoteNum, q.clientName, q.clientCompany, q.clientEmail, q.clientPhone,
      q.cartItems, q.discount, q.total, q.validTill, new Date().toISOString()
    ]);
    return 'saved';
  } catch(e) {
    return 'error: ' + e.message;
  }
}

function updateQuotationStatus(qtId, newStatus, notes) {
  try {
    const sh = _getSheet(SN.QUOTATIONS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(qtId)) {
        sh.getRange(i + 1, 16).setValue(newStatus);
        sh.getRange(i + 1, 26).setValue(_now());
        if (newStatus === 'Sent') sh.getRange(i + 1, 18).setValue(_now());
        if (newStatus === 'Accepted' || newStatus === 'Won') sh.getRange(i + 1, 20).setValue(_now());
        if (newStatus === 'Rejected' || newStatus === 'Lost') {
          sh.getRange(i + 1, 21).setValue(_now());
          if (notes) sh.getRange(i + 1, 22).setValue(notes);
        }
        // Sync to lead
        const leadId = data[i][2];
        if (leadId) _updateLeadQuotationInfo(leadId, qtId, newStatus);
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function getQuotations(filters) {
  try {
    const sh = _getSheet(SN.QUOTATIONS);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    let rows = data.slice(1).filter(r => r[0]);
    if (filters) {
      if (filters.leadId) rows = rows.filter(r => String(r[2]) === String(filters.leadId));
      if (filters.status) rows = rows.filter(r => r[15] === filters.status);
    }
    return rows.map(r => ({
      id: r[0], createdDate: r[1], leadId: r[2], leadName: r[3],
      leadEmail: r[4], company: r[5], industry: r[6], service: r[7],
      monthlyValue: r[8], durationMonths: r[9], totalValue: r[10],
      discount: r[11], finalValue: r[12], gst: r[13], grandTotal: r[14],
      status: r[15], version: r[16], sentDate: r[17], validUntil: r[18],
      acceptedDate: r[19], rejectedDate: r[20], rejectionReason: r[21],
      workflowId: r[22], createdBy: r[23], notes: r[24]
    })).reverse();
  } catch(e) { return []; }
}

function _updateLeadQuotationInfo(leadId, qtId, qtStatus) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(leadId)) {
        sh.getRange(i + 1, 24).setValue('Yes');                  // Quotation Sent
        sh.getRange(i + 1, 25).setValue(qtStatus);               // Quotation Status
        if (qtStatus === 'Sent') sh.getRange(i + 1, 10).setValue('Proposal Sent'); // Update lead status
        if (qtStatus === 'Won' || qtStatus === 'Accepted') sh.getRange(i + 1, 10).setValue('Sold');
        sh.getRange(i + 1, 21).setValue(_now());                 // Last Action Date
        return;
      }
    }
  } catch(e) {}
}

function getLeads(userEmail, isAdmin) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
 
    // Find the leads sheet — try multiple common names
    const sheet = ss.getSheetByName('MASTER_LEADS')
               || ss.getSheetByName('Leads')
               || ss.getSheetByName('leads')
               || ss.getSheetByName('LEADS')
               || ss.getSheets()[0];
 
    if (!sheet) { Logger.log('[getLeads] No sheet found!'); return []; }
 
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) { Logger.log('[getLeads] Sheet has no data rows'); return []; }
 
    const data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0].map(h => String(h || '').toLowerCase().trim());
 
    Logger.log('[getLeads] Sheet: ' + sheet.getName() + ' | Rows: ' + lastRow);
 
    // Flexible column finder
    const findCol = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i].replace(/[\s_\-]/g, '');
        for (const k of (Array.isArray(keywords) ? keywords : [keywords])) {
          const kn = k.toLowerCase().replace(/[\s_\-]/g, '');
          if (h === kn || h.includes(kn)) return i;
        }
      }
      return -1;
    };
 
    const colLeadId        = findCol(['leadid','lead_id','id','srno','sr']);
    const colName          = findCol(['name','fullname','contactname','contact','clientname']);
    const colPhone         = findCol(['phone','mobile','contactnumber','phonenumber']);
    const colEmail         = findCol(['email','emailid','mail','emailaddress']);
    const colCompany       = findCol(['company','companyname','organisation','organization','business']);
    const colDomain        = findCol(['domain','website','url']);
    const colIndustry      = findCol(['industry','service','category','servicerequired','type']);
    const colStatus        = findCol(['status','leadstatus','stage','leadstage']);
    const colFunnel        = findCol(['funnel','funnelstage','pipeline']);
    const colPriority      = findCol(['priority']);
    const colHot           = findCol(['hot','ishot','flag','hotlead']);
    const colSource        = findCol(['source','leadsource','via']);
    const colCity          = findCol(['city','location','region','area']);
    const colDeal          = findCol(['deal','value','dealvalue','amount','budget','revenue']);
    const colAssignedTo    = findCol(['assignedto','assigned_to','assigned','agent','bdm','bde','salesperson']);
    const colAssignedEmail = findCol(['assignedemail','assigned_email','agentemail','bdeemail']);
    const colAssignedEmpId = findCol(['assignedempid','assigned_empid','empid']);
    const colAddedDate     = findCol(['dateadded','date_added','addeddate','added','created','timestamp','createdat']);
    const colArchived      = findCol(['archived','archive','isarchived']);
 
    Logger.log('[getLeads] Key cols — name:' + colName + ' assignedTo:' + colAssignedTo + ' assignedEmail:' + colAssignedEmail);
 
    // Resolve userEmpId from EMPLOYEES sheet (for empId-based matching)
    let userEmpId = '';
    if (userEmail && !isAdmin) {
      try {
        const empSheet = ss.getSheetByName('EMPLOYEES');
        if (empSheet) {
          const empData = empSheet.getDataRange().getValues();
          for (let ei = 1; ei < empData.length; ei++) {
            if (String(empData[ei][2] || '').trim().toLowerCase() === userEmail.trim().toLowerCase()) {
              userEmpId = String(empData[ei][0] || '').trim();
              break;
            }
          }
        }
      } catch(e) {}
    }
 
    Logger.log('[getLeads] User: ' + userEmail + ' | userEmpId: ' + userEmpId + ' | isAdmin: ' + isAdmin);
 
    const leads = [];
 
    for (let i = 1; i < data.length; i++) {
      const row      = data[i];
      const nameVal  = colName  >= 0 ? row[colName]  : '';
      const phoneVal = colPhone >= 0 ? row[colPhone] : '';
      if (!nameVal && !phoneVal) continue; // skip empty rows
 
      // ── ASSIGNMENT FILTER (skip for admins or if no userEmail) ──
      if (!isAdmin && userEmail) {
        const aTo    = String(colAssignedTo    >= 0 ? (row[colAssignedTo]    || '') : '').trim().toLowerCase();
        const aEmail = String(colAssignedEmail >= 0 ? (row[colAssignedEmail] || '') : '').trim().toLowerCase();
        const aEmpId = String(colAssignedEmpId >= 0 ? (row[colAssignedEmpId] || '') : '').trim();
        const uEmail = userEmail.trim().toLowerCase();
        const uEmpId = userEmpId.trim();
 
        let matched = false;
 
        // 1. Best match: dedicated assignedEmail column
        if (!matched && aEmail && aEmail === uEmail) matched = true;
 
        // 2. assignedTo stores email (set by assignLeadsByEmail pre-fix)
        if (!matched && aTo && aTo === uEmail) matched = true;
 
        // 3. assignedEmpId column matches
        if (!matched && uEmpId && aEmpId && aEmpId === uEmpId) matched = true;
 
        // 4. assignedTo stores empId (legacy bulk-assign pre-fix)
        if (!matched && uEmpId && aTo && aTo === uEmpId.toLowerCase()) matched = true;
 
        if (!matched) continue;
      }
 
      // ── Parse date ──────────────────────────────────────────────
      let addedAt = new Date().toISOString();
      try {
        const raw = colAddedDate >= 0 ? row[colAddedDate] : null;
        if (raw) {
          const d = new Date(raw);
          if (!isNaN(d.getTime())) addedAt = d.toISOString();
        }
      } catch(e) {}
 
      const isArchived = colArchived >= 0 &&
        ['yes','true','1','archived'].includes(String(row[colArchived]||'').toLowerCase());
 
      const leadId = (colLeadId >= 0 && row[colLeadId])
        ? String(row[colLeadId]).trim()
        : ('L' + String(i).padStart(4, '0'));
 
      // Get the stored assignedTo value (now NAME for display)
      const assignedToVal    = colAssignedTo    >= 0 ? String(row[colAssignedTo]    || '').trim() : '';
      const assignedEmailVal = colAssignedEmail >= 0 ? String(row[colAssignedEmail] || '').trim() : '';
      const assignedEmpIdVal = colAssignedEmpId >= 0 ? String(row[colAssignedEmpId] || '').trim() : '';
 
      leads.push({
        id:            leadId,
        name:          String(nameVal  || '').trim() || 'Unknown',
        phone:         String(phoneVal || '').trim(),
        email:         colEmail    >= 0 ? String(row[colEmail]    || '').trim() : '',
        company:       colCompany  >= 0 ? String(row[colCompany]  || '').trim() : '',
        domain:        colDomain   >= 0 ? String(row[colDomain]   || '').trim() : '',
        industry:      colIndustry >= 0 ? String(row[colIndustry] || '').trim() : '',
        status:        colStatus   >= 0 ? String(row[colStatus]   || 'New').trim() : 'New',
        funnelStage:   colFunnel   >= 0 ? String(row[colFunnel]   || 'New Lead').trim() : 'New Lead',
        priority:      colPriority >= 0 ? String(row[colPriority] || 'Medium').trim() : 'Medium',
        hot:           colHot      >= 0 && ['yes','true','1','hot'].includes(String(row[colHot]||'').toLowerCase()),
        source:        colSource   >= 0 ? String(row[colSource]   || '').trim() : '',
        city:          colCity     >= 0 ? String(row[colCity]     || '').trim() : '',
        deal:          colDeal     >= 0 ? (parseFloat(String(row[colDeal] || '').replace(/[,₹\s]/g, '')) || 0) : 0,
        assignedTo:    assignedToVal,    // NAME (for display)
        assignedEmail: assignedEmailVal, // EMAIL (for matching)
        assignedEmpId: assignedEmpIdVal, // EMPID (for fallback)
        notes:         [],
        addedAt:       addedAt,
        archived:      isArchived,
        activityLog:   (() => {
          const ac = findCol(['activitylog','activity_log','activity']);
          if (ac < 0) return [];
          try {
            const acts = JSON.parse(String(row[ac] || '[]'));
            if (!Array.isArray(acts)) return [];
            return [...acts].reverse().slice(0, 30).map(a => ({
              time:    a.t   || '',
              by:      a.by  || 'Agent',
              email:   a.em  || '',
              action:  a.act || '',
              label:   a.lbl || (a.act||'').replace(/_/g,' '),
              details: a.det || '',
              icon:    a.ic  || '📌',
              color:   a.col || '#6B7280'
            }));
          } catch(e) { return []; }
        })()
      });
    }
 
    Logger.log('[getLeads] Returning ' + leads.length + ' leads for ' + (userEmail || 'admin'));
    return leads;
 
  } catch(e) {
    Logger.log('[getLeads] FATAL: ' + e.message + '\n' + e.stack);
    return [];
  }
}


// Verify assignment columns exist in sheet
function verifyLeadAssignmentColumns() {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return { success: false, message: 'Leads sheet not found' };
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const findCol = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]).toLowerCase().trim();
        for (const k of (Array.isArray(keywords) ? keywords : [keywords])) {
          if (h.includes(k.toLowerCase())) return { col: i + 1, header: headers[i] };
        }
      }
      return null;
    };

    const assignedTo = findCol(['assigned to', 'assigned_to']);
    const assignedEmail = findCol(['assigned email', 'assigned_email']);
    const assignedId = findCol(['assigned empid', 'assigned_empid']);

    return {
      success: true,
      assignedToCol: assignedTo ? { col: assignedTo.col, name: assignedTo.header } : null,
      assignedEmailCol: assignedEmail ? { col: assignedEmail.col, name: assignedEmail.header } : null,
      assignedIdCol: assignedId ? { col: assignedId.col, name: assignedId.header } : null,
      allHeaders: headers
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
function setLeadFilter(empEmail, hotOnly, industries, priorities, channelStatus, statuses) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh   = ss.getSheetByName('LEAD_FILTERS');
    if (!sh) {
      sh = ss.insertSheet('LEAD_FILTERS');
      sh.appendRow(['Employee Email','Hot Only','Industries','Priorities','Status','Updated','Statuses']);
      sh.getRange(1,1,1,7).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
    }
    const data = sh.getDataRange().getValues();
    const row  = [
      empEmail,
      hotOnly ? 'Yes' : 'No',
      Array.isArray(industries) ? industries.join(',') : (industries||''),
      Array.isArray(priorities) ? priorities.join(',') : (priorities||''),
      'Active',
      new Date().toISOString(),
      statuses || ''
    ];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === empEmail.toLowerCase()) {
        sh.getRange(i+1, 1, 1, 7).setValues([row]);
        return { success: true };
      }
    }
    sh.appendRow(row);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function removeLeadFilter(empEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('LEAD_FILTERS');
    if (!sh) return { success: true };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === empEmail.toLowerCase()) {
        sh.getRange(i+1,5).setValue('Inactive');
        return { success: true };
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false };
  }
}

function getLeadFilters() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('LEAD_FILTERS');
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(r => r[0] && r[4] === 'Active')
      .map(r => ({
        email:      String(r[0]||''),
        hotOnly:    r[1] === 'Yes',
        industries: String(r[2]||''),
        priorities: String(r[3]||''),
        status:     String(r[4]||''),
        statuses:   String(r[6]||'')
      }));
  } catch(e) { return []; }
}

function addLead(d) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return { success: false, message: 'Run Initialize Sheets first.' };
    const email = _getEmail();
    const id = d.id || _uid('LD');
    sh.appendRow([
      id, _now(), d.name||'', d.email||'', d.phone||'',
      d.company||'', d.domain||'', d.industry||'', d.source||'Manual',
      'New', 'New Lead', d.priority||'Medium', 'No',
      email, d.city||'', '', '', d.deal||0,
      d.notes||'', 'Lead Added', _now(),
      0, 0, 'No', '', d.tags||'', email, d.website||'', d.linkedin||''
    ]);
    _logAction({ leadId: id, action: 'LEAD_ADDED', details: 'New lead: ' + d.name });
    return { success: true, id };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


function saveQuotationHTML(quoteNum, htmlContent) {
  try {
    const folder = DriveApp.getRootFolder();
    const fileName = 'quotation_' + quoteNum + '.html';
    
    // Delete old file with same name if exists
    const existing = DriveApp.getFilesByName(fileName);
    while (existing.hasNext()) { existing.next().setTrashed(true); }
    
    // Create new file
    const file = folder.createFile(fileName, htmlContent, MimeType.HTML);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Return direct view URL
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) {
    return null;
  }
}

function updateLead(d) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    const updateHeaders = data[0];
    // Get actual column positions dynamically so they work regardless of sheet layout
    const headers = data[0];
    const findCol = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]).toLowerCase().trim().replace(/[\s_\-]/g,'');
        for (const k of (Array.isArray(keywords)?keywords:[keywords])) {
          if (h.includes(k.toLowerCase().replace(/[\s_\-]/g,''))) return i + 1; // 1-indexed
        }
      }
      return -1;
    };

    const fieldMap = {
      status:        findCol(['status','leadstatus']),
      funnelStage:   findCol(['funnel','funnelstage']),
      priority:      findCol(['priority']),
      hot:           findCol(['hot','ishot','hotlead']),
      assignedTo:    findCol(['assignedto','assigned_to','assigned']),
      assignedEmail: findCol(['assignedemail','assigned_email']),
      assignedEmpId: findCol(['assignedempid','assigned_empid']),
      deal:          findCol(['deal','value','dealvalue','amount']),
      tags:          findCol(['tags']),
      stageId:       findCol(['stageid','stage_id'])
    };
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(d.leadId)) {
        Object.entries(d).forEach(([k, v]) => {
          if (k !== 'leadId' && fieldMap[k]) {
            sh.getRange(i + 1, fieldMap[k]).setValue(v);
          }
        });
        sh.getRange(i + 1, 21).setValue(_now());
        const changedFields = Object.entries(d).filter(([k]) => k !== 'leadId').map(([k,v]) => k + '→' + v).join(', ');
        _appendLeadActivity(d.leadId, {
          email:   _getEmail(),
          action:  d.status ? 'STATUS_CHANGED' : 'LEAD_UPDATED',
          details: changedFields.slice(0, 120)
        });
        _logAction({ leadId: d.leadId, action: 'LEAD_UPDATED', details: JSON.stringify(d) });
        // Re-read the updated row and check routing rules
        try {
          const updatedRow = sh.getRange(i + 1, 1, 1, updateHeaders.length).getValues()[0];
          _applyRoutingRulesToRow(updatedRow, updateHeaders, sh, i + 1);
        } catch(re) { Logger.log('Routing check error: ' + re.message); }
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function tagHotLead(leadId, isHot) {
  try {
    const result = updateLead({ leadId, hot: isHot ? 'Yes' : 'No' });
    if (isHot) {
      try {
        const sh = _getSheet(SN.LEADS);
        const data = sh.getDataRange().getValues();
        const tagHeaders = data[0];
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === String(leadId)) {
            _sendHotLeadAlert({ leadName: data[i][2], company: data[i][5], industry: data[i][7], leadId });
            _applyRoutingRulesToRow(data[i], tagHeaders, sh, i + 1);
            break;
          }
        }
      } catch(e) { Logger.log('tagHotLead routing error: ' + e.message); }
      _logAction({ leadId, action: 'HOT_LEAD_TAGGED', details: 'Tagged as Hot' });
    }
    return result;
  } catch(e) { return { success: false }; }
}

function archiveLead(leadId) {
  try {
    const result = updateLead({ leadId, status: 'Archived' });
    _logAction({ leadId, action: 'LEAD_ARCHIVED', details: 'Lead archived' });
    return result;
  } catch(e) { return { success: false }; }
}

function bulkDeleteLeads(ids) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    const idSet = new Set(ids.map(String));
    for (let i = data.length - 1; i >= 1; i--) {
      if (idSet.has(String(data[i][0]))) sh.deleteRow(i + 1);
    }
    return { success: true };
  } catch(e) { return { success: false }; }
}

function _incrementLeadCounter(leadId, field) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return;
    const colMap = { callCount: 22, meetingCount: 23 };
    const col = colMap[field];
    if (!col) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(leadId)) {
        const prev = parseInt(data[i][col - 1]) || 0;
        sh.getRange(i + 1, col).setValue(prev + 1);
        return;
      }
    }
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════
//  TICKETS
// ═══════════════════════════════════════════════════════════════

function raiseTicket(d) {
  try {
    const sh = _getSheet(SN.TICKETS);
    if (!sh) return { success: false };
    const email = _getEmail();
    const ticketId = d.id || _uid('TKT');
    sh.appendRow([
      ticketId, d.leadId||'', d.leadName||'', email,
      '',                   // Assigned To
      d.category||'General',
      d.priority||'Medium',
      d.subject||'',
      d.description||'',
      'Open',
      _now(), _now(), '', '', '', 'No', 0, '', d.tags||''
    ]);
    _logAction({ leadId: d.leadId, action: 'TICKET_RAISED', details: d.subject });
    return { success: true, ticketId };
  } catch(e) { return { success: false }; }
}

function updateTicketStatus(ticketId, newStatus, notes) {
  try {
    const sh = _getSheet(SN.TICKETS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(ticketId)) {
        const prevStatus = data[i][9];
        sh.getRange(i + 1, 10).setValue(newStatus);
        sh.getRange(i + 1, 12).setValue(_now());
        if (newStatus === 'Resolved') sh.getRange(i + 1, 13).setValue(_now());
        if (newStatus === 'Closed') sh.getRange(i + 1, 14).setValue(_now());
        if (newStatus === 'Reopened') {
          const prev = parseInt(data[i][16]) || 0;
          sh.getRange(i + 1, 17).setValue(prev + 1);
        }
        if (notes) sh.getRange(i + 1, 15).setValue(notes);
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function resolveTicket(ticketId, resolutionNotes) {
  return updateTicketStatus(ticketId, 'Resolved', resolutionNotes);
}

function getTickets(filters) {
  try {
    const sh = _getSheet(SN.TICKETS);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    let rows = data.slice(1).filter(r => r[0]);
    if (filters) {
      if (filters.status) rows = rows.filter(r => r[9] === filters.status);
      if (filters.leadId) rows = rows.filter(r => String(r[1]) === String(filters.leadId));
    }
    return rows.map(r => ({
      id: r[0], leadId: r[1], leadName: r[2], raisedBy: r[3],
      assignedTo: r[4], category: r[5], priority: r[6],
      subject: r[7], description: r[8], status: r[9],
      createdAt: r[10], lastUpdated: r[11], resolvedDate: r[12],
      closedDate: r[13], resolution: r[14], slaBreach: r[15],
      reopenCount: r[16], internalNotes: r[17], tags: r[18]
    })).reverse();
  } catch(e) { return []; }
}

// ═══════════════════════════════════════════════════════════════
//  MEETINGS
// ═══════════════════════════════════════════════════════════════

function scheduleMeeting(d) {
  try {
    const sh = _getSheet(SN.MEETINGS);
    if (!sh) return { success: false };
    const email = _getEmail();
    const mtgId = d.id || _uid('MTG');
    let calEventId = '';
    try {
      const cal = CalendarApp.getDefaultCalendar();
      const startDt = new Date(d.date + 'T' + (d.time || '09:00') + ':00');
      const endDt = new Date(startDt.getTime() + (parseInt(d.duration) || 30) * 60000);
      const event = cal.createEvent(d.title || 'Meeting', startDt, endDt, {
        description: d.summary || '',
        guests: d.leadEmail || '',
        sendInvites: !!(d.leadEmail)
      });
      calEventId = event.getId();
    } catch(e) { calEventId = 'PENDING'; }
    sh.appendRow([
      mtgId, d.leadId||'', d.leadName||'', d.leadEmail||'', email,
      d.title||'Meeting', d.date||'', d.time||'',
      parseInt(d.duration)||30, d.summary||'',
      calEventId, '', '', 'Scheduled', '', 'No', '',
      _now(), _now()
    ]);
    if (d.leadId) _incrementLeadCounter(d.leadId, 'meetingCount');
    _logAction({ leadId: d.leadId, action: 'MEETING_SCHEDULED', details: d.title });
    return { success: true, calEventId, id: mtgId };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateMeetingStatus(meetingId, newStatus, outcomeNotes) {
  try {
    const sh = _getSheet(SN.MEETINGS);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(meetingId)) {
        sh.getRange(i + 1, 14).setValue(newStatus);
        sh.getRange(i + 1, 19).setValue(_now());
        if (outcomeNotes) sh.getRange(i + 1, 15).setValue(outcomeNotes);
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function getMeetings(leadId) {
  try {
    const sh = _getSheet(SN.MEETINGS);
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    let rows = data.slice(1).filter(r => r[0]);
    if (leadId) rows = rows.filter(r => String(r[1]) === String(leadId));
    return rows.map(r => ({
      id: r[0], leadId: r[1], leadName: r[2], leadEmail: r[3],
      employee: r[4], title: r[5], date: r[6], time: r[7],
      duration: r[8], summary: r[9], calEventId: r[10],
      meetLink: r[11], recordingLink: r[12], status: r[13],
      outcomeNotes: r[14], followUp: r[15], followUpDate: r[16],
      createdDate: r[17]
    })).reverse();
  } catch(e) { return []; }
}

// ═══════════════════════════════════════════════════════════════
//  NOTES
// ═══════════════════════════════════════════════════════════════

function addNote(leadId, note) {
  try {
    const sh = _getSheet(SN.NOTES);
    if (!sh) return { success: false };
    const noteId = _uid('N');
    const email = _getEmail();
    sh.appendRow([
      noteId, leadId, email,
      typeof note === 'string' ? note : (note.text || ''),
      _now(),
      typeof note === 'object' ? (note.type || 'General') : 'General',
      typeof note === 'object' ? (note.callId || '') : '',
      'No', 'Public'
    ]);
    _appendLeadActivity(leadId, {
      email:   email,
      action:  'NOTE_ADDED',
      details: (typeof note === 'string' ? note : (note.text || '')).slice(0, 100)
    });
    return { success: true, noteId };
  } catch(e) { return { success: false }; }
}

function getNotes(leadId) {
  try {
    const sh = _getSheet(SN.NOTES);
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(r => r[0] && String(r[1]) === String(leadId))
      .map(r => ({ id: r[0], leadId: r[1], author: r[2], text: r[3], time: r[4], type: r[5], callId: r[6] }))
      .reverse();
  } catch(e) { return []; }
}

function addCallNote(callLogId, leadId, noteText) {
  try {
    addNote(leadId, { text: noteText, type: 'Call Note', callId: callLogId });
    return { success: true };
  } catch(e) { return { success: false }; }
}

// ═══════════════════════════════════════════════════════════════
//  CALL LOGGING
// ═══════════════════════════════════════════════════════════════

function logCall(leadId, duration, outcome) {
  try {
    const callId = _uid('CALL');
    const sh = _getSheet(SN.LOGS);
    if (!sh) return { success: true, callId };
    const email = _getEmail();
    sh.appendRow([callId, _now(), email, leadId, 'CALL_MADE', outcome||'Initiated', duration||'0', outcome||'Initiated', '', '']);
    if (leadId) {
      _incrementLeadCounter(leadId, 'callCount');
      _appendLeadActivity(leadId, {
        email:  email,
        action: 'CALL_MADE',
        details:'Outcome: ' + (outcome || 'Initiated') + (duration && duration !== '—' ? ' | Duration: ' + duration + 's' : '')
      });
    }
    _updatePerf(email, 'calls');
    return { success: true, callId };
  } catch(e) { return { success: true, callId: _uid('CALL') }; }
}

function getCallLogs(leadId) {
  try {
    const sh = _getSheet(SN.LOGS);
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(r => r[0] && r[4] === 'CALL_MADE' && (!leadId || String(r[3]) === String(leadId)))
      .map(r => ({ id: r[0], time: r[1], by: r[2], leadId: r[3], outcome: r[7], duration: r[6], notes: r[5] }))
      .reverse().slice(0, 200);
  } catch(e) { return []; }
}

// ═══════════════════════════════════════════════════════════════
//  SHIFT TRACKING
// ═══════════════════════════════════════════════════════════════

function startShift(emailOverride) {
  try {
    const email = emailOverride || Session.getActiveUser().getEmail();
    if (!email) return _json({ success:false });

    const sh = _getSheet(SN.SHIFTS);
    if (!sh) return _json({ success:false });

    const empName = _getEmployeeName(email) || '';
    const now = new Date();

    sh.appendRow([
      _uid('SHF'),
      email,
      empName,
      Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(now, CONFIG.TZ, 'HH:mm:ss'),
      '',
      'Active',
      '',
      '',
      '',
      '',
      '',
      0,
      0,
      0,
      ''
    ]);
    return _json({ success:true });
  } catch (e) {
    return _json({ success:false, error:e.message });
  }
}

function endShift(emailOverride) {
  try {
    const email = emailOverride || Session.getActiveUser().getEmail();
    const sh = _getSheet(SN.SHIFTS);
    if (!sh) return _json({ success:false });

    const now = Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm:ss');
    const data = sh.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === email && (data[i][6] === 'Active' || data[i][6] === 'On Break')) {
        sh.getRange(i + 1, 6).setValue(now);
        sh.getRange(i + 1, 7).setValue('Ended');
        break;
      }
    }
    return _json({ success:true });
  } catch (e) {
    return _json({ success:false, error:e.message });
  }
}

function startBreak(emailOverride) {
  try {
    const email = emailOverride || Session.getActiveUser().getEmail();
    const sh = _getSheet(SN.SHIFTS);
    if (!sh) return _json({ success:false });

    const now = Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm:ss');
    const data = sh.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === email && data[i][6] === 'Active') {
        sh.getRange(i + 1, 7).setValue('On Break');
        sh.getRange(i + 1, 8).setValue(now);
        break;
      }
    }
    return _json({ success:true });
  } catch (e) {
    return _json({ success:false, error:e.message });
  }
}

function endBreak(emailOverride) {
  try {
    const email = emailOverride || Session.getActiveUser().getEmail();
    const sh = _getSheet(SN.SHIFTS);
    if (!sh) return _json({ success:false });

    const now = Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm:ss');
    const data = sh.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === email && data[i][6] === 'On Break') {
        sh.getRange(i + 1, 7).setValue('Active');
        sh.getRange(i + 1, 9).setValue(now);
        break;
      }
    }
    return _json({ success:true });
  } catch (e) {
    return _json({ success:false, error:e.message });
  }
}

function logShift(empEmail, empId, action, extraData) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const now     = new Date().toISOString();
    const extra   = extraData || {};

    // ── 1. Write to ShiftEvents (read by admin panel's renderShiftsTable) ──
    let shiftEvSh = ss.getSheetByName('ShiftEvents');
    if (!shiftEvSh) {
      shiftEvSh = ss.insertSheet('ShiftEvents');
      shiftEvSh.appendRow(['EmpId','EmpName','EmpEmail','Event','BreakMinutes','TotalBreakSeconds','Time']);
      shiftEvSh.getRange(1,1,1,7).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      shiftEvSh.setFrozenRows(1);
    }

    // Resolve employee name from EMPLOYEES sheet
    let empName = '';
    const empSh = _getSheet(SN.EMPLOYEES);
    if (empSh && empEmail) {
      const empData = empSh.getDataRange().getValues();
      for (let ei = 1; ei < empData.length; ei++) {
        if (String(empData[ei][2]||'').toLowerCase().trim() === empEmail.toLowerCase().trim()) {
          empName = String(empData[ei][1]||'');
          break;
        }
      }
    }

    // Map action string to event name
    const eventMap = {
      'start':       'shift_start',
      'end':         'shift_end',
      'break_start': 'break_start',
      'break_end':   'break_end',
      'shift_start': 'shift_start',
      'shift_end':   'shift_end'
    };
    const event = eventMap[action] || action;

    shiftEvSh.appendRow([
      empId  || '',
      empName,
      empEmail || '',
      event,
      extra.breakMinutes      || '',
      extra.totalBreakSeconds || '',
      now
    ]);

    // ── 2. Also update SHIFT_TRACKER (keeps existing startShift/endShift working) ──
    const sh = _getSheet(SN.SHIFTS);
    if (sh) {
      if (action === 'start' || action === 'shift_start') {
        sh.appendRow([
          _uid('SHF'),
          empEmail,
          empName,
          Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'),
          Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm:ss'),
          '', 'Active', '', '', '', '', '', 0, 0, 0, ''
        ]);
      } else if (action === 'end' || action === 'shift_end') {
        const shData = sh.getDataRange().getValues();
        for (let i = shData.length - 1; i >= 1; i--) {
          if (String(shData[i][1]||'').toLowerCase() === empEmail.toLowerCase() &&
              (shData[i][6] === 'Active' || shData[i][6] === 'On Break')) {
            sh.getRange(i+1, 6).setValue(Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm:ss'));
            sh.getRange(i+1, 7).setValue('Ended');
            break;
          }
        }
      } else if (action === 'break_start') {
        const shData = sh.getDataRange().getValues();
        for (let i = shData.length - 1; i >= 1; i--) {
          if (String(shData[i][1]||'').toLowerCase() === empEmail.toLowerCase() && shData[i][6] === 'Active') {
            sh.getRange(i+1, 7).setValue('On Break');
            sh.getRange(i+1, 8).setValue(Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm:ss'));
            break;
          }
        }
      } else if (action === 'break_end') {
        const shData = sh.getDataRange().getValues();
        for (let i = shData.length - 1; i >= 1; i--) {
          if (String(shData[i][1]||'').toLowerCase() === empEmail.toLowerCase() && shData[i][6] === 'On Break') {
            sh.getRange(i+1, 7).setValue('Active');
            sh.getRange(i+1, 9).setValue(Utilities.formatDate(new Date(), CONFIG.TZ, 'HH:mm:ss'));
            break;
          }
        }
      }
    }

    // ── 3. Update employee status in EMPLOYEES sheet ──
    if (empSh && empEmail) {
      const eData = empSh.getDataRange().getValues();
      const eH    = eData[0];
      const eIdx  = eH.findIndex(h => /email/i.test(h));
      const sIdx  = eH.findIndex(h => /^status$/i.test(String(h).trim()));
      const laIdx = eH.findIndex(h => /last.?active/i.test(h));
  const statusMap = {
        'start':       'Active',
        'shift_start': 'Active',
        'end':         'Inactive',
        'shift_end':   'Inactive',
        'break_start': 'On Break',
        'break_end':   'Active'
      };
      const newStatus = statusMap[action] || 'Active';
      for (let ei = 1; ei < eData.length; ei++) {
        if (String(eData[ei][eIdx]||'').toLowerCase().trim() === empEmail.toLowerCase().trim()) {
          if (sIdx  >= 0) empSh.getRange(ei+1, sIdx+1).setValue(newStatus);
          if (laIdx >= 0) empSh.getRange(ei+1, laIdx+1).setValue(now);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    Logger.log('logShift error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
// ═══════════════════════════════════════════════════════════════
//  EMPLOYEES
// ═══════════════════════════════════════════════════════════════

function saveEmployee(profile) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let empSh = ss.getSheetByName(SN.EMPLOYEES);
    if (!empSh) { initializeSheets(); empSh = ss.getSheetByName(SN.EMPLOYEES); }
    let authSh = ss.getSheetByName(SN.AUTH);
    if (!authSh) { initializeSheets(); authSh = ss.getSheetByName(SN.AUTH); }

    const emailNorm = (profile.email || '').trim().toLowerCase();

    // Save to EMPLOYEES sheet
    const empData = empSh.getDataRange().getValues();
    let empRowFound = false;
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][0]) === String(profile.empId) ||
          String(empData[i][2]).toLowerCase() === emailNorm) {
        empSh.getRange(i+1, 1, 1, 23).setValues([[
          profile.empId || empData[i][0],
          profile.name || empData[i][1],
          emailNorm || empData[i][2],
          profile.phone || empData[i][3],
          profile.role || empData[i][4],
          profile.exp || empData[i][5],
          Array.isArray(profile.categories) ? profile.categories.join(',') : (profile.categories || empData[i][6]),
          profile.quota || empData[i][7],
          0, 0,
          profile.status || empData[i][10],
          profile.shiftStart || empData[i][11],
          profile.shiftEnd || empData[i][12],
          60, 'No', _now(), _now(),
          profile.avatar || (profile.name || 'A')[0].toUpperCase(),
          0, 0, '', '', ''
        ]]);
        empRowFound = true;
        break;
      }
    }
    if (!empRowFound) {
      empSh.appendRow([
        profile.empId || _uid('EMP'),
        profile.name || '',
        emailNorm,
        profile.phone || '',
        profile.role || 'BDE',
        profile.exp || 'Mid (1-3 yr)',
        Array.isArray(profile.categories) ? profile.categories.join(',') : (profile.categories || ''),
        profile.quota || '10',
        0, 0,
        profile.status || 'Active',
        profile.shiftStart || '09:00',
        profile.shiftEnd || '18:00',
        60, 'No', _now(), _now(),
        (profile.name || 'A')[0].toUpperCase(),
        0, 0, '', '', ''
      ]);
    }

    // Save to AUTH sheet (so employee can login immediately)
    const authData = authSh.getDataRange().getValues();
    let authRowFound = false;
    for (let i = 1; i < authData.length; i++) {
      if (String(authData[i][0]).toLowerCase() === emailNorm) {
        authSh.getRange(i+1, 1, 1, 13).setValues([[
          emailNorm,
          profile.name || authData[i][1],
          profile.role || authData[i][2],
          profile.status || authData[i][3],
          profile.empId || authData[i][4],
          profile.exp || authData[i][5],
          Array.isArray(profile.categories) ? profile.categories.join(',') : (profile.categories || authData[i][6]),
          profile.quota || authData[i][7],
          profile.shiftStart || authData[i][8],
          profile.shiftEnd || authData[i][9],
          profile.phone || authData[i][10],
          authData[i][11], _now()
        ]]);
        authRowFound = true;
        break;
      }
    }
    if (!authRowFound) {
      authSh.appendRow([
        emailNorm,
        profile.name || '',
        profile.role || 'BDE',
        profile.status || 'Active',
        profile.empId || _uid('EMP'),
        profile.exp || 'Mid (1-3 yr)',
        Array.isArray(profile.categories) ? profile.categories.join(',') : (profile.categories || ''),
        profile.quota || '10',
        profile.shiftStart || '09:00',
        profile.shiftEnd || '18:00',
        profile.phone || '',
        '', _now()
      ]);
    }

    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// REPLACE getAllEmployees in code.js
function getAllEmployees() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Try EMPLOYEES sheet first
    const empSh = ss.getSheetByName('EMPLOYEES');
    if (empSh && empSh.getLastRow() > 1) {
      const data = empSh.getDataRange().getValues();
      const results = data.slice(1).filter(r => r[0] || r[1]).map(r => ({
        empId:      String(r[0]  || 'EMP-' + Math.random().toString(36).slice(2,7)),
        name:       String(r[1]  || ''),
        email:      String(r[2]  || ''),
        phone:      String(r[3]  || ''),
        role:       String(r[4]  || 'BDE'),
        exp:        String(r[5]  || ''),
        categories: String(r[6]  || '').split(',').map(c=>c.trim()).filter(Boolean),
        quota:      String(r[7]  || '10'),
        leads:      Number(r[8]  || 0),
        closed:     Number(r[9]  || 0),
        status:     String(r[10] || 'Active'),
        shiftStart: String(r[11] || '09:00'),
        shiftEnd:   String(r[12] || '18:00'),
        avatar:     String(r[17] || (r[1]||'A')[0].toUpperCase()),
        score:      Number(r[18] || 0),
        calls:      Number(r[18] || 0),
        meetings:   Number(r[19] || 0)
      }));
      if (results.length > 0) return results;
    }
    
    // Fallback: read from AUTHORIZED_USERS sheet
    const authSh = ss.getSheetByName('AUTHORIZED_USERS');
    if (!authSh || authSh.getLastRow() < 2) return [];
    const authData = authSh.getDataRange().getValues();
    return authData.slice(1).filter(r => r[0]).map((r, i) => ({
      empId:      String(r[4]  || 'EMP-' + (i+1)),
      name:       String(r[1]  || r[0].split('@')[0]),
      email:      String(r[0]  || ''),
      phone:      String(r[10] || ''),
      role:       String(r[2]  || 'BDE'),
      exp:        String(r[5]  || ''),
      categories: String(r[6]  || '').split(',').map(c=>c.trim()).filter(Boolean),
      quota:      String(r[7]  || '10'),
      leads:      0,
      closed:     0,
      status:     String(r[3]  || 'Active'),
      shiftStart: String(r[8]  || '09:00'),
      shiftEnd:   String(r[9]  || '18:00'),
      avatar:     (String(r[1]||r[0]||'A')[0]).toUpperCase(),
      score:      0,
      calls:      0,
      meetings:   0
    }));
  } catch(e) {
    return [];
  }
}

function assignLeadsByName(empId, empName, empEmail, leadIds, category) {
  // Keep for backward compatibility — delegates to assignLeadsByEmail
  return assignLeadsByEmail(empId, empName, empEmail, leadIds, category);
}

function assignLeadsByEmail(empId, empName, empEmail, leadIds, category) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return { success: false, message: 'Leads sheet not found' };
 
    const data    = sh.getDataRange().getValues();
    const headers = data[0];
 
    // Flexible column finder — handles any header naming convention
    const findCol = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]).toLowerCase().trim().replace(/[\s_\-]/g, '');
        for (const k of (Array.isArray(keywords) ? keywords : [keywords])) {
          const kn = k.toLowerCase().replace(/[\s_\-]/g, '');
          if (h === kn || h.includes(kn)) return i;
        }
      }
      return -1;
    };
 
    const leadIdCol        = findCol(['leadid','lead_id','id']);
    const assignedToCol    = findCol(['assignedto','assigned_to','assigned','agent','bde','bdm','salesperson']);
    const assignedEmailCol = findCol(['assignedemail','assigned_email','agentemail','bdeemail']);
    const assignedEmpIdCol = findCol(['assignedempid','assigned_empid','empid']);
    const lastActionCol    = findCol(['lastactiondate','last_action_date','lastupdated','last_updated','lastaction']);
 
    // Log column positions for debugging
    Logger.log('[assignLeadsByEmail] Cols — leadId:' + leadIdCol +
               ' assignedTo:' + assignedToCol +
               ' assignedEmail:' + assignedEmailCol +
               ' assignedEmpId:' + assignedEmpIdCol);
 
    if (leadIdCol === -1) {
      return { success: false, message: 'Lead ID column not found. Headers: ' + JSON.stringify(headers) };
    }
    if (assignedToCol === -1) {
      return { success: false, message: 'Assigned To column not found. Headers: ' + JSON.stringify(headers) };
    }
 
    const idSet = new Set(leadIds.map(String));
    let count = 0;
    const now = new Date().toISOString();
 
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][leadIdCol] || '').trim();
      if (!idSet.has(rowId)) continue;
 
      // ── WRITE: "Assigned To" = NAME (for display in sheet & admin panel)
      sh.getRange(i + 1, assignedToCol + 1).setValue(empName);
 
      // ── WRITE: "Assigned Email" = EMAIL (for reliable getLeads() matching)
      if (assignedEmailCol >= 0) {
        sh.getRange(i + 1, assignedEmailCol + 1).setValue(empEmail);
      }
 
      // ── WRITE: "Assigned EmpId" = empId (fallback matching)
      if (assignedEmpIdCol >= 0) {
        sh.getRange(i + 1, assignedEmpIdCol + 1).setValue(empId);
      }
 
      // ── WRITE: Last Action Date
      if (lastActionCol >= 0) {
        sh.getRange(i + 1, lastActionCol + 1).setValue(now);
      }
 
      count++;
    }
 
    Logger.log('[assignLeadsByEmail] Assigned ' + count + ' leads to ' + empName + ' <' + empEmail + '> (EmpId: ' + empId + ')');
 
    _logAction({
      leadId:  leadIds[0] || '',
      action:  'LEAD_ASSIGNED',
      details: count + ' leads assigned to ' + empName + ' (' + empEmail + ') [EmpId: ' + empId + ']'
    });
 
    SpreadsheetApp.flush(); // Force immediate write
    return { success: true, assigned: count, message: count + ' leads assigned to ' + empName };
 
  } catch(e) {
    Logger.log('assignLeadsByEmail error: ' + e.message);
    return { success: false, message: e.message };
  }
}


// Round-robin bulk assign — used by Auto-Assign button
function bulkAssignLeadsByEmail(employees, leadIds, mode) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return { success: false, message: 'Leads sheet not found' };

    const data    = sh.getDataRange().getValues();
    const headers = data[0];

    const findCol = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]).toLowerCase().trim().replace(/[\s_\-]/g, '');
        for (const k of (Array.isArray(keywords) ? keywords : [keywords])) {
          if (h.includes(k.toLowerCase().replace(/[\s_\-]/g, ''))) return i;
        }
      }
      return -1;
    };

    const leadIdCol        = findCol(['leadid','id']);
    const assignedToCol    = findCol(['assignedto','assigned_to']);
    const assignedEmailCol = findCol(['assignedemail','assigned_email']);
    const assignedEmpIdCol = findCol(['assignedempid','assigned_empid']);
    const lastActionCol    = findCol(['lastactiondate','lastupdated']);

    if (leadIdCol === -1 || assignedToCol === -1) {
      return { success: false, message: 'Required columns not found' };
    }

    const idSet = new Set(leadIds.map(String));
    let count = 0;
    let empIndex = 0;

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][leadIdCol] || '').trim();
      if (!idSet.has(rowId)) continue;

      const emp = employees[empIndex % employees.length];
      empIndex++;

      sh.getRange(i + 1, assignedToCol + 1).setValue(emp.name || emp.empId);
      if (assignedEmailCol >= 0) sh.getRange(i + 1, assignedEmailCol + 1).setValue(emp.email);
      if (assignedEmpIdCol >= 0) sh.getRange(i + 1, assignedEmpIdCol + 1).setValue(emp.empId);
      if (lastActionCol   >= 0) sh.getRange(i + 1, lastActionCol    + 1).setValue(new Date().toISOString());
      count++;
    }

    return { success: true, assigned: count };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  PERFORMANCE
// ═══════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════
//  ACTION LOGGING
// ═══════════════════════════════════════════════════════════════

function _logActivity(d) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sh     = ss.getSheetByName(SN.ACTIVITY);
    if (!sh)   { _initActivityLogSheet(ss); sh = ss.getSheetByName(SN.ACTIVITY); }
    const email = d.email || _getEmail();
    const now   = d.timestamp || new Date().toISOString();
    sh.appendRow([
      d.id         || _uid('LOG'),
      now,
      now.slice(0, 10),
      email,
      d.empName    || _getEmployeeName(email) || '',
      d.leadId     || '',
      d.action     || d.type || 'UNKNOWN',
      d.details    || d.text || '',
      d.extra      ? JSON.stringify(d.extra) : '',
      d.duration   || '',
      d.outcome    || '',
      d.scoreDelta || '',
      d.sessionId  || '',
      d.source     || 'app'
    ]);
  } catch(e) { Logger.log('_logActivity error: ' + e.message); }
}

// Backward-compat wrapper — all existing _logAction calls continue to work
function _logAction(d) {
  _logActivity({
    email:    d.email,
    leadId:   d.leadId,
    action:   d.action,
    details:  d.details,
    duration: d.duration,
    outcome:  d.outcome,
    sessionId:d.sessionId
  });
}

function _updatePerf(email, metric) {
  // Performance data now lives as individual ACTIVITY_LOG rows — no separate sheet
  _logActivity({
    email:  email,
    action: metric === 'loginTime' ? 'LOGIN' : metric === 'logoutTime' ? 'LOGOUT' : (metric || 'ACTION').toUpperCase(),
    details: 'Metric: ' + metric,
    source:  'perf'
  });
}

function getPerformanceSummary(empEmail) {
  try {
    const email = empEmail || _getEmail();
    const sh    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN.ACTIVITY);
    if (!sh || sh.getLastRow() < 2) return { calls:0, emails:0, meetings:0, quotations:0, workflows:0, score:0 };
    const today = _today();
    const data  = sh.getDataRange().getValues();
    const summary = { calls:0, emails:0, meetings:0, quotations:0, workflows:0, hotLeads:0, notes:0, leads:0, loginTime:'', logoutTime:'', score:0 };
    data.slice(1).forEach(r => {
      if (String(r[2]).slice(0,10) !== today) return;
      if (String(r[3]).toLowerCase() !== email.toLowerCase()) return;
      const act = String(r[6] || '');
      if (act === 'CALL_MADE'          || act === 'CALL_COMPLETED')  summary.calls++;
      if (act === 'EMAIL_SENT')                                       summary.emails++;
      if (act === 'MEETING_SCHEDULED'  || act === 'MEETING_COMPLETED') summary.meetings++;
      if (act === 'QUOTATION_CREATED'  || act === 'QUOTATION_SENT')   summary.quotations++;
      if (act === 'WORKFLOW_CREATED'   || act === 'WORKFLOW_COMPLETED') summary.workflows++;
      if (act === 'HOT_LEAD_TAGGED')                                  summary.hotLeads++;
      if (act === 'NOTE_ADDED')                                       summary.notes++;
      if (act === 'LEAD_ADDED')                                       summary.leads++;
      if (act === 'LOGIN'  && !summary.loginTime)   summary.loginTime  = String(r[1]||'');
      if (act === 'LOGOUT')                         summary.logoutTime = String(r[1]||'');
    });
    summary.score = summary.calls * 10 + summary.meetings * 25 + summary.quotations * 50 + summary.hotLeads * 15;
    return summary;
  } catch(e) { return { calls:0, emails:0, meetings:0, quotations:0, workflows:0, score:0 }; }
}
function resetEmployeeLeads(empId, empName, empEmail, affectedIds) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return { success: false, message: 'Leads sheet not found' };

    const data    = sh.getDataRange().getValues();
    const headers = data[0];

    // Flexible column finder — handles any header naming convention
    const findCol = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]).toLowerCase().trim().replace(/[\s_\-]/g, '');
        for (const k of (Array.isArray(keywords) ? keywords : [keywords])) {
          if (h.includes(k.toLowerCase().replace(/[\s_\-]/g, ''))) return i;
        }
      }
      return -1;
    };

    const leadIdCol        = findCol(['leadid','lead_id','id']);
    const assignedToCol    = findCol(['assignedto','assigned_to','assigned','agent','bde']);
    const assignedEmailCol = findCol(['assignedemail','assigned_email','agentemail']);
    const assignedEmpIdCol = findCol(['assignedempid','assigned_empid']);
    const lastActionCol    = findCol(['lastactiondate','last_action_date','lastupdated','last_updated']);

    if (assignedToCol === -1) {
      return { success: false, message: 'Assigned To column not found. Headers found: ' + JSON.stringify(headers) };
    }

    // Build lookup set from affectedIds if provided
    const useIdSet = Array.isArray(affectedIds) && affectedIds.length > 0;
    const idSet    = useIdSet ? new Set(affectedIds.map(String)) : null;

    const normEmail = (empEmail || '').trim().toLowerCase();
    const normName  = (empName  || '').trim().toLowerCase();

    let cleared = 0;

    for (let i = 1; i < data.length; i++) {
      const rowLeadId = leadIdCol >= 0 ? String(data[i][leadIdCol] || '').trim() : '';
      const rowName   = String(data[i][assignedToCol]              || '').trim().toLowerCase();
      const rowEmail  = assignedEmailCol >= 0 ? String(data[i][assignedEmailCol] || '').trim().toLowerCase() : '';
      const rowEmpId  = assignedEmpIdCol >= 0 ? String(data[i][assignedEmpIdCol] || '').trim() : '';

      // Match logic — tries email first (most reliable), then empId, then name
      let isMatch = false;

      if (useIdSet) {
        // If we have specific lead IDs, match only those
        isMatch = idSet.has(rowLeadId);
      } else {
        // Fallback: match by any assignment identifier
        if (normEmail && rowEmail && rowEmail === normEmail) isMatch = true;
        else if (empId && rowEmpId && rowEmpId === empId)   isMatch = true;
        else if (normName && rowName && rowName === normName) isMatch = true;
      }

      if (!isMatch) continue;

      // Clear ALL three assignment fields atomically
      sh.getRange(i + 1, assignedToCol + 1).setValue('');
      if (assignedEmailCol >= 0) sh.getRange(i + 1, assignedEmailCol + 1).setValue('');
      if (assignedEmpIdCol >= 0) sh.getRange(i + 1, assignedEmpIdCol + 1).setValue('');
      if (lastActionCol    >= 0) sh.getRange(i + 1, lastActionCol    + 1).setValue(new Date().toISOString());

      cleared++;
    }

    // Log the reset
    _logAction({
      leadId:  '',
      action:  'LEADS_RESET',
      details: 'Cleared ' + cleared + ' leads from ' + empName + ' (' + empEmail + ') [EmpId: ' + empId + ']',
      email:   empEmail
    });

    // Write to AdminLog sheet for audit trail
    try {
      const ss       = SpreadsheetApp.getActiveSpreadsheet();
      let   logSheet = ss.getSheetByName('AdminLog');
      if (!logSheet) {
        logSheet = ss.insertSheet('AdminLog');
        logSheet.appendRow(['Timestamp','Action','EmpName','EmpEmail','EmpId','LeadsCleared','By']);
        logSheet.getRange(1,1,1,7).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
        logSheet.setFrozenRows(1);
      }
      logSheet.appendRow([
        new Date(), 'RESET_EMPLOYEE_LEADS',
        empName, empEmail, empId,
        cleared + ' leads cleared',
        'Admin Panel'
      ]);
    } catch(e) {}

    SpreadsheetApp.flush(); // Force immediate write to sheet

    return { success: true, cleared: cleared, message: cleared + ' leads unassigned from ' + empName };
  } catch(e) {
    Logger.log('resetEmployeeLeads error: ' + e.message + ' | Stack: ' + e.stack);
    return { success: false, message: e.message };
  }
}

function savePipelineConfig(pipeline) {
  try {
    if (!pipeline || !Array.isArray(pipeline)) return { success: false };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh   = ss.getSheetByName('PIPELINE_CONFIG');
    if (!sh) {
      sh = ss.insertSheet('PIPELINE_CONFIG');
      sh.appendRow(['Stage ID','Order','Name','Entry Status','Trigger Status','Employees (EmpIds)','Allowed Actions','Created At']);
      sh.getRange(1,1,1,8).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
    } else {
      // Clear existing config rows (keep header)
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 8).clearContent();
    }
    pipeline.forEach(stage => {
      sh.appendRow([
        stage.id    || '',
        stage.order || 0,
        stage.name  || '',
        stage.entryStatus    || '',
        stage.triggerStatus  || '',
        (stage.employees     || []).join(','),
        (stage.allowedActions|| []).join(','),
        new Date().toISOString()
      ]);
    });
    SpreadsheetApp.flush();
    return { success: true, saved: pipeline.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  EMAIL ALERTS
// ═══════════════════════════════════════════════════════════════

function _sendHotLeadAlert(d) {
  try {
    MailApp.sendEmail({
      to: CONFIG.HOT_LEAD_ALERT_EMAIL,
      subject: '🔥 Hot Lead Alert: ' + d.leadName + ' | Nucleous',
      htmlBody: '<h2>🔥 Hot Lead Tagged</h2><table><tr><td><b>Name</b></td><td>' + d.leadName + '</td></tr><tr><td><b>Company</b></td><td>' + d.company + '</td></tr><tr><td><b>Industry</b></td><td>' + d.industry + '</td></tr><tr><td><b>Lead ID</b></td><td>' + d.leadId + '</td></tr></table><p><i>Nucleous by AfterResult Solutions</i></p>'
    });
  } catch(e) {}
}

function sendPerformanceReport() {
  try {
    const employees = getAllEmployees();
    const today = _today();
    employees.forEach(emp => {
      if (!emp.email || emp.status === 'Inactive' || emp.status === 'Deactivated') return;
      try {
        const perf = getPerformanceSummary(emp.email);
        MailApp.sendEmail({
          to: emp.email,
          subject: 'Daily Report – ' + today + ' | Nucleous',
          body: 'Daily Summary for ' + emp.name + '\n\nDate: ' + today + '\nCalls: ' + (perf.calls||0) + '\nMeetings: ' + (perf.meetings||0) + '\nQuotations: ' + (perf.quotations||0) + '\nScore: ' + (perf.score||0) + '\n\nNucleous by AfterResult'
        });
      } catch(e) {}
    });
    return { success: true };
  } catch(e) { return { success: true }; }
}

// ═══════════════════════════════════════════════════════════════
//  UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function _getEmployeeName(email) {
  try {
    const sh = _getSheet(SN.EMPLOYEES);
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === String(email||'').toLowerCase()) return String(data[i][1]);
    }
    // Try auth sheet
    const auth = _getSheet(SN.AUTH);
    if (auth) {
      const aData = auth.getDataRange().getValues();
      for (let i = 1; i < aData.length; i++) {
        if (String(aData[i][0]).toLowerCase() === String(email||'').toLowerCase()) return String(aData[i][1]);
      }
    }
    return null;
  } catch(e) { return null; }
}

function _getLeadById(leadId) {
  try {
    const sh = _getSheet(SN.LEADS);
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(leadId)) {
        return { id: data[i][0], name: data[i][2], email: data[i][3], phone: data[i][4], company: data[i][5], industry: data[i][7], status: data[i][9] };
      }
    }
    return null;
  } catch(e) { return null; }
}

function getSheetUrl() {
  try { return SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch(e) { return ''; }
}

// ═══════════════════════════════════════════════════════════════
//  COMPANY PROFILE — READ / WRITE
// ═══════════════════════════════════════════════════════════════

function getCompanyProfile() {
  try {
    var sh = _ensureCompanyProfileSheet();
    if (!sh || sh.getLastRow() < 2) return {};
    var data = sh.getDataRange().getValues();
    var profile = {};
    data.slice(1).forEach(function(r) { if (r[0]) profile[String(r[0])] = String(r[1] || ''); });
    return profile;
  } catch(e) {
    Logger.log('getCompanyProfile error: ' + e.message);
    return {};
  }
}

function saveCompanyProfile(updates) {
  try {
    var sh = _ensureCompanyProfileSheet();
    if (!sh) return { success: false, error: 'Sheet not available' };
    var data = sh.getDataRange().getValues();
    var now = new Date().toISOString();
    var existingRows = {};
    data.slice(1).forEach(function(r, i) { if (r[0]) existingRows[String(r[0])] = i + 2; });
    Object.keys(updates).forEach(function(key) {
      if (key === 'updatedAt') return;
      var val = updates[key] === null || updates[key] === undefined ? '' : updates[key];
      if (existingRows[key]) {
        sh.getRange(existingRows[key], 2, 1, 2).setValues([[String(val), now]]);
      } else {
        sh.appendRow([key, String(val), now]);
      }
    });
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    Logger.log('saveCompanyProfile error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  SERVICES CATALOG — READ / WRITE / DELETE
// ═══════════════════════════════════════════════════════════════

var DEFAULT_SERVICES = [
  {id:'SVC-001',name:'Social Media Management',category:'Social Media',price:15000,unit:'per month',description:'Complete management of Facebook, Instagram, LinkedIn pages including content creation and posting',status:'Active',hsn:'',notes:''},
  {id:'SVC-002',name:'Social Media Marketing - Basic',category:'Social Media',price:8000,unit:'per month',description:'Basic social media marketing with 12 posts per month across 2 platforms',status:'Active',hsn:'',notes:''},
  {id:'SVC-003',name:'Social Media Marketing - Standard',category:'Social Media',price:18000,unit:'per month',description:'Standard plan with 20 posts per month across 3 platforms plus paid ads management',status:'Active',hsn:'',notes:''},
  {id:'SVC-004',name:'Social Media Marketing - Premium',category:'Social Media',price:35000,unit:'per month',description:'Premium plan with unlimited posts, 5 platforms, paid ads, influencer outreach',status:'Active',hsn:'',notes:''},
  {id:'SVC-005',name:'Google Ads Management',category:'Paid Advertising',price:12000,unit:'per month',description:'Complete Google Ads setup and management, keyword research, ad copy, optimization',status:'Active',hsn:'',notes:''},
  {id:'SVC-006',name:'Meta Ads Management',category:'Paid Advertising',price:10000,unit:'per month',description:'Facebook and Instagram paid advertising campaign management',status:'Active',hsn:'',notes:''},
  {id:'SVC-007',name:'SEO - Basic',category:'SEO',price:8000,unit:'per month',description:'On-page SEO, keyword optimization, meta tags, sitemap for up to 10 pages',status:'Active',hsn:'',notes:''},
  {id:'SVC-008',name:'SEO - Standard',category:'SEO',price:18000,unit:'per month',description:'On-page and off-page SEO, link building, monthly reporting for up to 25 pages',status:'Active',hsn:'',notes:''},
  {id:'SVC-009',name:'SEO - Premium',category:'SEO',price:35000,unit:'per month',description:'Full SEO suite including technical SEO, content strategy, link building, competitor analysis',status:'Active',hsn:'',notes:''},
  {id:'SVC-010',name:'Website Development - Basic',category:'Web Development',price:25000,unit:'one time',description:'5-page informational website with contact form, mobile responsive, 1 month support',status:'Active',hsn:'',notes:''},
  {id:'SVC-011',name:'Website Development - Business',category:'Web Development',price:55000,unit:'one time',description:'10-page business website with CMS, blog, contact forms, 3 months support',status:'Active',hsn:'',notes:''},
  {id:'SVC-012',name:'E-Commerce Website',category:'Web Development',price:95000,unit:'one time',description:'Full e-commerce store with payment gateway, product management, order tracking',status:'Active',hsn:'',notes:''},
  {id:'SVC-013',name:'Amazon Marketplace Setup',category:'Marketplace',price:15000,unit:'one time',description:'Amazon seller account setup, product listing optimization, A+ content creation',status:'Active',hsn:'',notes:''},
  {id:'SVC-014',name:'Amazon Marketplace Management',category:'Marketplace',price:12000,unit:'per month',description:'Monthly Amazon store management, listing updates, ads, review management',status:'Active',hsn:'',notes:''},
  {id:'SVC-015',name:'Flipkart Marketplace Setup',category:'Marketplace',price:12000,unit:'one time',description:'Flipkart seller account setup, product listing, catalog creation',status:'Active',hsn:'',notes:''},
  {id:'SVC-016',name:'Flipkart Marketplace Management',category:'Marketplace',price:10000,unit:'per month',description:'Monthly Flipkart store management, listing updates, promotions',status:'Active',hsn:'',notes:''},
  {id:'SVC-017',name:'Content Writing - Blog',category:'Content',price:500,unit:'per article',description:'SEO-optimized blog articles of 800-1200 words',status:'Active',hsn:'',notes:''},
  {id:'SVC-018',name:'Content Writing - Website Copy',category:'Content',price:8000,unit:'one time',description:'Professional website copywriting for up to 5 pages',status:'Active',hsn:'',notes:''},
  {id:'SVC-019',name:'Video Production - Basic',category:'Creative',price:15000,unit:'per video',description:'1-2 minute promotional video with voiceover and background music',status:'Active',hsn:'',notes:''},
  {id:'SVC-020',name:'Graphic Design - Social Media Pack',category:'Creative',price:5000,unit:'per month',description:'Monthly pack of 15 custom social media creatives',status:'Active',hsn:'',notes:''},
  {id:'SVC-021',name:'Email Marketing Setup',category:'Email Marketing',price:8000,unit:'one time',description:'Email marketing account setup, template design, automation workflows',status:'Active',hsn:'',notes:''},
  {id:'SVC-022',name:'Email Marketing Management',category:'Email Marketing',price:6000,unit:'per month',description:'Monthly email campaigns, list management, analytics reporting',status:'Active',hsn:'',notes:''},
  {id:'SVC-023',name:'WhatsApp Marketing Setup',category:'WhatsApp Marketing',price:10000,unit:'one time',description:'WhatsApp Business API setup, broadcast list creation, message templates',status:'Active',hsn:'',notes:''},
  {id:'SVC-024',name:'Lead Generation Campaign',category:'Lead Generation',price:20000,unit:'per month',description:'Multi-channel lead generation including ads, landing pages and follow-up automation',status:'Active',hsn:'',notes:''},
  {id:'SVC-025',name:'Brand Strategy Package',category:'Branding',price:45000,unit:'one time',description:'Complete brand identity including logo, brand guidelines, color palette, typography',status:'Active',hsn:'',notes:''}
];

/* ═══════════════════════════════════════════════════
   DEFAULT SERVICES CATALOG — 25 services seeded on first run
═══════════════════════════════════════════════════ */
var DEFAULT_SERVICES_CATALOG = [
  ['SVC-001','Social Media Management','Social Media',15000,'per month','Complete management of Facebook, Instagram, LinkedIn pages with content creation and posting','Active','',''],
  ['SVC-002','Social Media Marketing - Basic','Social Media',8000,'per month','12 posts per month across 2 platforms with basic engagement','Active','',''],
  ['SVC-003','Social Media Marketing - Standard','Social Media',18000,'per month','20 posts per month across 3 platforms plus paid ads management up to Rs.10,000 ad spend','Active','',''],
  ['SVC-004','Social Media Marketing - Premium','Social Media',35000,'per month','Unlimited posts across 5 platforms, paid ads management, influencer outreach, monthly report','Active','',''],
  ['SVC-005','Google Ads Management','Paid Advertising',12000,'per month','Complete Google Ads setup, keyword research, ad copy creation, bid management and monthly reporting','Active','',''],
  ['SVC-006','Meta Ads Management','Paid Advertising',10000,'per month','Facebook and Instagram paid advertising campaign creation, management and optimization','Active','',''],
  ['SVC-007','SEO - Basic','SEO',8000,'per month','On-page SEO, keyword optimization, meta tags and sitemap for up to 10 pages','Active','',''],
  ['SVC-008','SEO - Standard','SEO',18000,'per month','On-page and off-page SEO, link building, competitor analysis for up to 25 pages','Active','',''],
  ['SVC-009','SEO - Premium','SEO',35000,'per month','Full SEO suite including technical SEO, content strategy, link building, monthly detailed reports','Active','',''],
  ['SVC-010','Website Development - Basic','Web Development',25000,'one time','5-page informational website with contact form, mobile responsive, 1 month support included','Active','',''],
  ['SVC-011','Website Development - Business','Web Development',55000,'one time','10-page business website with CMS, blog section, contact forms, 3 months support included','Active','',''],
  ['SVC-012','E-Commerce Website','Web Development',95000,'one time','Full e-commerce store with payment gateway integration, product management, order tracking system','Active','',''],
  ['SVC-013','Amazon Marketplace Setup','Marketplace',15000,'one time','Amazon seller account setup, product listing optimization, A+ content creation, brand registry','Active','',''],
  ['SVC-014','Amazon Marketplace Management','Marketplace',12000,'per month','Monthly Amazon store management, listing updates, sponsored ads management, review management','Active','',''],
  ['SVC-015','Flipkart Marketplace Setup','Marketplace',12000,'one time','Flipkart seller account setup, product listing creation, catalog management, payment setup','Active','',''],
  ['SVC-016','Flipkart Marketplace Management','Marketplace',10000,'per month','Monthly Flipkart store management, listing updates, promotional campaigns, performance tracking','Active','',''],
  ['SVC-017','Content Writing - Blog Articles','Content',500,'per article','SEO-optimized blog articles of 800 to 1200 words with keyword research and meta description','Active','',''],
  ['SVC-018','Content Writing - Website Copy','Content',8000,'one time','Professional website copywriting for up to 5 pages including home, about, services, contact','Active','',''],
  ['SVC-019','Video Production - Promotional','Creative',15000,'per video','1 to 2 minute promotional video with script, voiceover, background music and branded graphics','Active','',''],
  ['SVC-020','Graphic Design - Social Media Pack','Creative',5000,'per month','Monthly pack of 15 custom social media creatives sized for all major platforms','Active','',''],
  ['SVC-021','Email Marketing Setup','Email Marketing',8000,'one time','Email marketing account setup, template design, list segmentation and automation workflows','Active','',''],
  ['SVC-022','Email Marketing Management','Email Marketing',6000,'per month','Monthly email campaigns, list management, A/B testing, analytics and performance reporting','Active','',''],
  ['SVC-023','WhatsApp Marketing Setup','WhatsApp Marketing',10000,'one time','WhatsApp Business API setup, broadcast list creation, message templates and automation','Active','',''],
  ['SVC-024','Lead Generation Campaign','Lead Generation',20000,'per month','Multi-channel lead generation including ads, landing pages, lead magnets and follow-up automation','Active','',''],
  ['SVC-025','Brand Strategy Package','Branding',45000,'one time','Complete brand identity including logo design, brand guidelines, color palette and typography system','Active','','']
];

function _ensureServicesCatalogSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('SERVICES_CATALOG');
    if (!sh) {
      sh = ss.insertSheet('SERVICES_CATALOG');
      var headers = ['ID','Name','Category','Price','Unit','Description','Status','HSN','Notes','UpdatedAt'];
      sh.appendRow(headers);
      var headerRange = sh.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#0f172a');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setFontSize(11);
      sh.setFrozenRows(1);
      sh.setColumnWidth(1, 100);
      sh.setColumnWidth(2, 220);
      sh.setColumnWidth(3, 140);
      sh.setColumnWidth(4, 90);
      sh.setColumnWidth(5, 100);
      sh.setColumnWidth(6, 320);
      sh.setColumnWidth(7, 80);
      // Seed default services
      var now = new Date().toISOString();
      DEFAULT_SERVICES_CATALOG.forEach(function(row) {
        sh.appendRow(row.concat([now]));
      });
      SpreadsheetApp.flush();
    }
    return sh;
  } catch(e) {
    Logger.log('_ensureServicesCatalogSheet error: ' + e.message);
    return null;
  }
}

function _ensureCompanyProfileSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('COMPANY_PROFILE');
    if (!sh) {
      sh = ss.insertSheet('COMPANY_PROFILE');
      var headers = ['Key','Value','UpdatedAt'];
      sh.appendRow(headers);
      var headerRange = sh.getRange(1, 1, 1, 3);
      headerRange.setBackground('#0f172a');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(1, 200);
      sh.setColumnWidth(2, 400);
      sh.setColumnWidth(3, 160);
      // Seed default company profile keys
      var now = new Date().toISOString();
      var defaults = [
        ['companyName','AfterResult Solutions',now],
        ['tagline','Digital Marketing Agency',now],
        ['logoUrl','',now],
        ['website','https://afterresult.com',now],
        ['industry','Digital Marketing',now],
        ['email','info.afterresult@gmail.com',now],
        ['phone','+919991283530',now],
        ['phone2','',now],
        ['address','',now],
        ['city','',now],
        ['meetLink','',now],
        ['currency','Rs.',now],
        ['gst','',now],
        ['validity','30',now],
        ['quotationFooter','Thank you for your business. We look forward to serving you.',now],
        ['terms','Payment due within 30 days of invoice date. All prices are exclusive of applicable taxes.',now],
        ['bankName','',now],
        ['accountNo','',now],
        ['ifsc','',now],
        ['upi','',now],
        ['accountHolder','',now],
        ['signature','Best regards,\nAfterResult Solutions\ninfo.afterresult@gmail.com | +919991283530',now]
      ];
      defaults.forEach(function(row) { sh.appendRow(row); });
      SpreadsheetApp.flush();
    }
    return sh;
  } catch(e) {
    Logger.log('_ensureCompanyProfileSheet error: ' + e.message);
    return null;
  }
}

function _ensureQuotationsSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('QUOTATIONS');
    if (!sh) {
      sh = ss.insertSheet('QUOTATIONS');
      var headers = ['QuoteNum','ClientName','ClientCompany','ClientEmail','ValidTill','Discount','DiscountAmt','Subtotal','Total','Items','Notes','CreatedBy','CreatedAt','Status','ViewUrl'];
      sh.appendRow(headers);
      var headerRange = sh.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#0f172a');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(1, 110);
      sh.setColumnWidth(2, 160);
      sh.setColumnWidth(3, 160);
      sh.setColumnWidth(4, 200);
      sh.setColumnWidth(5, 100);
      sh.setColumnWidth(10, 400);
      SpreadsheetApp.flush();
    }
    return sh;
  } catch(e) {
    Logger.log('_ensureQuotationsSheet error: ' + e.message);
    return null;
  }
}

function getServicesCatalog(activeOnly) {
  try {
    var sh = _ensureServicesCatalogSheet();
    if (!sh || sh.getLastRow() < 2) return DEFAULT_SERVICES_CATALOG.map(function(row, i) {
      return { id: row[0], name: row[1], serviceName: row[1], category: row[2], price: row[3], unit: row[4], description: row[5], status: row[6], hsn: row[7], notes: row[8] };
    });
    var data = sh.getDataRange().getValues();
    var h = data[0];
    function col(k) { var idx = h.indexOf(k); return idx >= 0 ? idx : -1; }
    var rows = data.slice(1).filter(function(r) { return r[0]; }).map(function(r) {
      return {
        id:          String(r[col('ID')]          || r[0] || ''),
        name:        String(r[col('Name')]         || r[1] || ''),
        serviceName: String(r[col('Name')]         || r[1] || ''),
        category:    String(r[col('Category')]     || r[2] || ''),
        price:       parseFloat(r[col('Price')]    || r[3] || 0),
        unit:        String(r[col('Unit')]         || r[4] || ''),
        description: String(r[col('Description')] || r[5] || ''),
        status:      String(r[col('Status')]       || r[6] || 'Active'),
        hsn:         String(r[col('HSN')]          || r[7] || ''),
        notes:       String(r[col('Notes')]        || r[8] || '')
      };
    });
    if (activeOnly) rows = rows.filter(function(r) { return r.status === 'Active'; });
    return rows;
  } catch(e) {
    Logger.log('getServicesCatalog error: ' + e.message);
    return DEFAULT_SERVICES_CATALOG.map(function(row) {
      return { id: row[0], name: row[1], serviceName: row[1], category: row[2], price: row[3], unit: row[4], description: row[5], status: row[6], hsn: row[7], notes: row[8] };
    });
  }
}

function saveServiceEntry(entry) {
  try {
    var sh = _ensureServicesCatalogSheet();
    if (!sh) return { success: false, error: 'Sheet not found' };
    var data = sh.getDataRange().getValues();
    var h = data[0];
    var idIdx = h.indexOf('ID');
    var now = new Date().toISOString();
    var rowData = [
      entry.id || ('SVC-' + Date.now()),
      entry.name || entry.serviceName || '',
      entry.category || '',
      parseFloat(entry.price) || 0,
      entry.unit || '',
      entry.description || '',
      entry.status || 'Active',
      entry.hsn || '',
      entry.notes || '',
      now
    ];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(entry.id)) {
        sh.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    sh.appendRow(rowData);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    Logger.log('saveServiceEntry error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteServiceEntry(id) {
  try {
    var sh = _ensureServicesCatalogSheet();
    if (!sh) return { success: false };
    var data = sh.getDataRange().getValues();
    var h = data[0];
    var idIdx = h.indexOf('ID');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(id)) {
        sh.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Service not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function saveAllServices(entries) {
  try {
    if (!Array.isArray(entries)) return { success: false };
    entries.forEach(function(e) { saveServiceEntry(e); });
    return { success: true, saved: entries.length };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  COMPANY PROFILES LIST (PDF / Drive) — READ / WRITE / DELETE
// ═══════════════════════════════════════════════════════════════

function getCompanyProfilesList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SN.PROFILES);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(r => r[0] && r[7] !== 'No')
      .sort((a, b) => (parseInt(a[8]) || 99) - (parseInt(b[8]) || 99))
      .map(r => ({
        id:       String(r[0]), name:     String(r[1]),
        category: String(r[2]), desc:     String(r[3]),
        pages:    String(r[4]), driveUrl: String(r[5]),
        waMsg:    String(r[6]), active:   r[7] !== 'No',
        sortOrder:parseInt(r[8]) || 99, notes: String(r[10] || '')
      }));
  } catch(e) { return []; }
}

function saveProfileEntry(entry) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SN.PROFILES);
    if (!sh) { _initCompanyProfilesSheet(ss); sh = ss.getSheetByName(SN.PROFILES); }
    const data = sh.getDataRange().getValues();
    const row  = [
      entry.id       || ('PRF' + Date.now()),
      entry.name     || '',
      entry.category || 'Company Profile',
      entry.desc     || '',
      entry.pages    || '',
      entry.driveUrl || '',
      entry.waMsg    || 'Hi {clientName}, please find our profile attached.',
      entry.active !== false ? 'Yes' : 'No',
      parseInt(entry.sortOrder) || 99,
      new Date().toISOString(),
      entry.notes || ''
    ];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(entry.id || '')) {
        sh.getRange(i + 1, 1, 1, 11).setValues([row]);
        return { success: true, updated: true };
      }
    }
    sh.appendRow(row);
    return { success: true, created: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteProfileEntry(profileId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SN.PROFILES);
    if (!sh) return { success: false };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(profileId)) { sh.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false };
  } catch(e) { return { success: false, message: e.message }; }
}

function openSheetTab(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (sh) ss.setActiveSheet(sh);
    return getSheetUrl();
  } catch(e) { return getSheetUrl(); }
}

function exportLeads() {
  try { return SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch(e) { return ''; }
}

function pingActivity() {
  try {
    const email = _getEmail();
    const sh = _getSheet(SN.EMPLOYEES);
    if (!sh) return { success: true };
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === email.toLowerCase()) {
        sh.getRange(i + 1, 16).setValue(_now());
        return { success: true };
      }
    }
    return { success: true };
  } catch(e) { return { success: true }; }
}

function getDropdownOptions() {
  return DD;
}

function getFAQs() {
  return [
    { cat:'Getting Started', q:'How do I add a new lead?', a:'Click "+ Add Lead" in the top right of Lead Management. Fill contact details and click Add Lead.' },
    { cat:'Getting Started', q:'How do I mark a lead as hot?', a:'Click the flame icon on a lead card, or open the lead detail and click "Mark Hot".' },
    { cat:'Lead Management', q:'When should I tag a lead as Hot?', a:'Tag as Hot when: confirmed budget, decision-maker, timeline within 30 days, responded positively to 2+ touchpoints.' },
    { cat:'Lead Management', q:'What are the funnel stages?', a:'New Lead → Initial Contact → Interested → Demo Scheduled → Proposal Sent → Negotiating → Qualified → Closed Won / Closed Lost.' },
    { cat:'Communication', q:'How does Zira AI help?', a:'Zira drafts emails, handles objections, generates quotations, and suggests closing strategies.' },
    { cat:'Platform Tips', q:'What is the 6-minute AHT timer?', a:'Tracks your interaction duration per lead. Turns red at 2 minutes remaining.' },
    { cat:'AfterResult Services', q:'What services does AfterResult offer?', a:'OPM (Offline Presence Management), DPM (Digital Presence Management: SEO, Social Media, Ads, Website), Leadin (Lead Generation).' },
    { cat:'AfterResult Services', q:'What are the social media plan prices?', a:'Single platform: from ₹9,999/mo. Multi-platform: from ₹22,999/mo. With ads: ₹24,999–28,999/mo + ad budget. All + GST.' }
  ];
}

function getZiraSystemPrompt(lead) {
  const serviceName = (lead && lead.industry) ? lead.industry : 'Digital Marketing';
  const clientName  = (lead && lead.name)     ? lead.name     : '';
  const clientCo    = (lead && lead.company)  ? lead.company  : '';
  const clientCity  = (lead && lead.city)     ? lead.city     : '';

  return `You are Zira, an expert sales assistant for AfterResult Solutions — a full-service Digital Marketing Agency based in India.

CRITICAL CONTEXT — READ THIS CAREFULLY BEFORE GENERATING ANY EMAIL OR MESSAGE:
- The field called "Industry" on a lead does NOT describe what the client's business does.
- "Industry" describes the SERVICE the client is requesting FROM AfterResult Solutions.
- For example: if Industry = "Email Marketing Campaigns", it means the CLIENT WANTS AfterResult to run email marketing FOR THEM.
- If Industry = "SEO", the client wants AfterResult to handle SEO for their business.
- If Industry = "Google Ads Management", the client wants AfterResult to manage their Google Ads.
- ALWAYS frame all communication as: AfterResult is the PROVIDER of this service, and the lead is the BUYER.
- NEVER write sentences like "I came across your company while researching [service]" — that wrongly implies WE are looking for the service.
- ALWAYS write sentences like "We specialize in [service] and help businesses achieve [result]" — we are selling this service TO them.

CURRENT LEAD CONTEXT:
- Client Name: ${clientName || 'the client'}
- Company: ${clientCo || 'their company'}
- City: ${clientCity || 'their city'}
- Service They Are Looking For: ${serviceName}

AFTERRESULT SOLUTIONS — SENDER DETAILS (use these always, no placeholders):
- Company: AfterResult Solutions
- Sender Name: Paras Bhatia
- Email: info.afterresult@gmail.com
- Phone: +91 9991283530
- Website: afterresult.com

SERVICES AFTERRESULT PROVIDES (not limited to this list — use your knowledge to write about any service):
Social Media Management (Organic), Social Media + Paid Ads, Instagram Growth Plan,
LinkedIn Management, Google Ads Management, Meta Ads Management,
Amazon Marketplace Management, Flipkart Marketplace Management,
SEO – Search Engine Optimization, Website Design & Development,
Lead Generation (Leadin), Brand Identity & Design, Email Marketing Campaigns,
Performance Marketing (Full Stack), E-Commerce Store Setup,
OPM – Offline Presence Management, Custom Packages.
If the requested service is not in this list, use your knowledge to write about it confidently as a service AfterResult can deliver.

COLD EMAIL RULES — FOLLOW STRICTLY:
1. NEVER use any placeholder text such as [Your Name], [Company Name], [Your Contact Information], [Insert], etc. The email must be 100% ready to send.
2. Always end every email with this exact sign-off, no variations:
   Best regards,
   Paras Bhatia
   AfterResult Solutions
   📧 info.afterresult@gmail.com | 📞 +91 9991283530
3. The "Industry" field = the service the client wants to buy from us. Write the entire email around delivering that service TO them.
4. Personalize the email using the client's name, company name, and city wherever possible.
5. Keep the email to 3 short paragraphs. Do not write long essays.
6. Write a specific, compelling subject line based on the service requested.
7. Do not repeat the same sentence structure in every email. Make it feel human and natural.

EMAIL STRUCTURE TO FOLLOW:
- Paragraph 1: Brief introduction of AfterResult Solutions as an expert in [requested service] + why you are reaching out to this specific client.
- Paragraph 2: 2 to 3 bullet points showing the specific value/results AfterResult delivers through [requested service].
- Paragraph 3: Simple, low-pressure call to action — invite them for a 15-minute call to explore if we can help.

PRICING (share only if the client or user asks):
- Single platform social media management: from ₹9,999/month
- Multi-platform social media management: from ₹22,999/month
- Social media with paid ads: from ₹24,999 to ₹28,999/month + ad budget
- All prices are exclusive of 18% GST`;
}

function createDailyTrigger() {
  try {
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sendPerformanceReport') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('sendPerformanceReport').timeBased().atHour(19).everyDays(1).create();
    return { success: true, message: 'Daily report trigger set for 7PM.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ═══════════════════════════════════════════════════════════════
//  QUOTATION BUILDER — SHEET-DRIVEN CATALOG
//  Sheet: QUOTATION_BUILDER
//  All services, products, plans, SKUs, and add-ons are managed
//  from this sheet. The employer panel reads it once on load.
//  No edit UI in the panel — all updates via the sheet itself.
// ═══════════════════════════════════════════════════════════════

var QB_SHEET_NAME = 'QUOTATION_BUILDER';
var QB_HEADERS = [
  'RowType','CategoryName','CategoryType','Icon',
  'ItemID','ItemName','Description',
  'Price','PriceType','Retainer','AdditionalCost',
  'TaxIncluded','Timeline','IsAddon','AddonFor','AddonNote',
  'SortOrder','Active','Notes'
];

function _ensureQuotationBuilderSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(QB_SHEET_NAME);
    if (sh) return sh;
    return _initQuotationBuilderSheet(ss);
  } catch(e) {
    Logger.log('_ensureQuotationBuilderSheet error: ' + e.message);
    return null;
  }
}

function _initQuotationBuilderSheet(ss) {
  try {
    var sh = ss.getSheetByName(QB_SHEET_NAME);
    if (sh && sh.getLastRow() > 1) return sh;
    if (!sh) sh = ss.insertSheet(QB_SHEET_NAME);
    sh.clearContents();
    sh.appendRow(QB_HEADERS);
    var hRange = sh.getRange(1, 1, 1, QB_HEADERS.length);
    hRange.setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 110);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 100);
    sh.setColumnWidth(6, 220);
    sh.setColumnWidth(7, 280);
    sh.setColumnWidth(19, 220);

    var now = new Date().toISOString();
    var seed = [
      // ── SETTINGS ROW ──────────────────────────────────────────
      // CategoryName column holds the mode: service | product | both
      ['SETTINGS','service','','','','','','','','','','','','','','','','Yes','Mode: service = services only, product = products only, both = services and products'],

      // ── CATEGORIES ────────────────────────────────────────────
      ['CATEGORY','Social Media Management','service','','','','','','','','','','','','','',1,'Yes',''],
      ['CATEGORY','Amazon Marketplace','service','','','','','','','','','','','','','',2,'Yes',''],
      ['CATEGORY','Flipkart Marketplace','service','','','','','','','','','','','','','',3,'Yes',''],
      ['CATEGORY','Meesho Marketplace','service','','','','','','','','','','','','','',4,'Yes',''],
      ['CATEGORY','Myntra Marketplace','service','','','','','','','','','','','','','',5,'Yes',''],
      ['CATEGORY','TradeIndia / IndiaMart','service','','','','','','','','','','','','','',6,'Yes',''],
      ['CATEGORY','Complete Enablement Bundle','service','','','','','','','','','','','','','',7,'Yes',''],
      ['CATEGORY','E-Commerce Development','service','','','','','','','','','','','','','',8,'Yes',''],
      ['CATEGORY','Custom Package','service','','','','','','','','','','','','','',9,'Yes',''],

      // ── SOCIAL MEDIA MANAGEMENT PLANS ─────────────────────────
      ['PLAN','Social Media Management','','','smm-ig-only','Instagram Enablement + Optimization','15 Creatives + Festivals + Engagement',9999,'monthly',4500,'','Yes','',  'No','','',1,'Yes',''],
      ['PLAN','Social Media Management','','','smm-li-only','LinkedIn Enablement + Optimization',  '15 Creatives + Festivals + Engagement',11999,'monthly',5500,'','Yes','', 'No','','',2,'Yes',''],
      ['PLAN','Social Media Management','','','smm-fb-only','Facebook Enablement + Optimization',  '15 Creatives + Festivals + Engagement',11999,'monthly',5500,'','Yes','', 'No','','',3,'Yes',''],
      ['PLAN','Social Media Management','','','smm-ig-ads','Instagram Enablement + Optimization + Ads','15 Creatives + Festivals + Engagement + Ads',12999,'monthly',6500,'Ads Budget','Yes','','No','','',4,'Yes',''],
      ['PLAN','Social Media Management','','','smm-li-ads','LinkedIn Enablement + Optimization + Ads', '15 Creatives + Festivals + Engagement + Ads',13999,'monthly',7500,'Ads Budget','Yes','','No','','',5,'Yes',''],
      ['PLAN','Social Media Management','','','smm-fb-ads','Facebook Enablement + Optimization + Ads', '15 Creatives + Festivals + Engagement + Ads',13999,'monthly',7500,'Ads Budget','Yes','','No','','',6,'Yes',''],
      ['PLAN','Social Media Management','','','smm-ig-li','Social Media Enablement + Optimization',       'Instagram + LinkedIn + 20 creatives each + Festivals + Engagement',22999,'monthly',7700,'','No','','No','','',7,'Yes',''],
      ['PLAN','Social Media Management','','','smm-ig-li-ads','Social Media Enablement + Optimization + Ads','Instagram + LinkedIn + 20 creatives each + Festivals + Engagement + Ads',24999,'monthly',9700,'Ads Budget','No','','No','','',8,'Yes',''],
      ['PLAN','Social Media Management','','','smm-ig-li-fb','Social Media Enablement + Optimization + Facebook','Instagram + LinkedIn + Facebook + High quality creatives + Engagement',26999,'monthly',9000,'','No','','No','','',9,'Yes',''],
      ['PLAN','Social Media Management','','','smm-ig-li-fb-ads','Social Media Enablement + Optimization + Facebook + Ads','Instagram + LinkedIn + Facebook + High quality creatives + Engagement + Ads',28999,'monthly',11000,'Ads Budget','No','','No','','',10,'Yes',''],

      // ── SMM ADD-ONS ────────────────────────────────────────────
      ['ADDON','Social Media Management','','','addon-ig-followers','Instagram Followers (200+ | 200-245)','Active/Inactive',510,'one-time','','','Yes','','Yes','Social Media Management','Active/Inactive',1,'Yes',''],
      ['ADDON','Social Media Management','','','addon-li-connections','LinkedIn Connections (200+)','Main Account Connecting',200,'one-time','','','Yes','','Yes','Social Media Management','Main Account Connecting',2,'Yes',''],
      ['ADDON','Social Media Management','','','addon-influencer','Influencer Marketing','Starting from',75999,'one-time','','','No','','Yes','Social Media Management','Starting from',3,'Yes',''],
      ['ADDON','Social Media Management','','','addon-ugc','UGC (Scripting + Shoot + Editing + Review)','Starting from',55999,'one-time','','','No','','Yes','Social Media Management','Starting from',4,'Yes',''],

      // ── AMAZON PLANS ──────────────────────────────────────────
      ['PLAN','Amazon Marketplace','','','amz-basic',   'Basic Plan – 45 Days',    '35 Listings · Rs.500/listing',20650,'one-time','','','Yes','45 Days','No','','',1,'Yes',''],
      ['PLAN','Amazon Marketplace','','','amz-standard','Standard Plan – 45 Days', '50 Listings · Rs.500/listing',29500,'one-time','','','Yes','45 Days','No','','',2,'Yes',''],
      ['PLAN','Amazon Marketplace','','','amz-premium', 'Premium Plan – 50 Days',  '100 Listings · Rs.500/listing',59000,'one-time','','','Yes','50 Days','No','','',3,'Yes',''],

      // ── AMAZON ADD-ONS ─────────────────────────────────────────
      ['ADDON','Amazon Marketplace','','','amz-inv',         'Stock & Inventory Management','',3000, 'monthly','','','No','','Yes','Amazon Marketplace','/month',1,'Yes',''],
      ['ADDON','Amazon Marketplace','','','amz-ads-with',    'Ads with Plan',              '',5000, 'monthly','','Ads Budget','No','','Yes','Amazon Marketplace','/month + ads budget',2,'Yes',''],
      ['ADDON','Amazon Marketplace','','','amz-ads-without', 'Ads without Plan',           '',10000,'monthly','','Ads Budget','No','','Yes','Amazon Marketplace','/month + ads budget',3,'Yes',''],

      // ── FLIPKART PLANS ────────────────────────────────────────
      ['PLAN','Flipkart Marketplace','','','fk-basic',   'Basic Plan – 45 Days',    '35 Listings · Rs.450/listing',18585,'one-time','','','Yes','45 Days','No','','',1,'Yes',''],
      ['PLAN','Flipkart Marketplace','','','fk-standard','Standard Plan – 45 Days', '50 Listings · Rs.450/listing',26550,'one-time','','','Yes','45 Days','No','','',2,'Yes',''],
      ['PLAN','Flipkart Marketplace','','','fk-premium', 'Premium Plan – 50 Days',  '100 Listings · Rs.450/listing',53100,'one-time','','','Yes','50 Days','No','','',3,'Yes',''],

      ['ADDON','Flipkart Marketplace','','','fk-inv',         'Stock & Inventory Management','',3000, 'monthly','','','No','','Yes','Flipkart Marketplace','/month',1,'Yes',''],
      ['ADDON','Flipkart Marketplace','','','fk-ads-with',    'Ads with Plan',              '',5000, 'monthly','','Ads Budget','No','','Yes','Flipkart Marketplace','/month + ads budget',2,'Yes',''],
      ['ADDON','Flipkart Marketplace','','','fk-ads-without', 'Ads without Plan',           '',10000,'monthly','','Ads Budget','No','','Yes','Flipkart Marketplace','/month + ads budget',3,'Yes',''],

      // ── MEESHO PLANS ──────────────────────────────────────────
      ['PLAN','Meesho Marketplace','','','ms-basic',   'Basic Plan – 45 Days',    '35 Listings · Rs.390/listing',16107,'one-time','','','Yes','45 Days','No','','',1,'Yes',''],
      ['PLAN','Meesho Marketplace','','','ms-standard','Standard Plan – 45 Days', '50 Listings · Rs.390/listing',23010,'one-time','','','Yes','45 Days','No','','',2,'Yes',''],
      ['PLAN','Meesho Marketplace','','','ms-premium', 'Premium Plan – 50 Days',  '100 Listings · Rs.390/listing',46020,'one-time','','','Yes','50 Days','No','','',3,'Yes',''],

      ['ADDON','Meesho Marketplace','','','ms-inv',         'Stock & Inventory Management','',3000, 'monthly','','','No','','Yes','Meesho Marketplace','/month',1,'Yes',''],
      ['ADDON','Meesho Marketplace','','','ms-ads-with',    'Ads with Plan',              '',5000, 'monthly','','Ads Budget','No','','Yes','Meesho Marketplace','/month + ads budget',2,'Yes',''],
      ['ADDON','Meesho Marketplace','','','ms-ads-without', 'Ads without Plan',           '',10000,'monthly','','Ads Budget','No','','Yes','Meesho Marketplace','/month + ads budget',3,'Yes',''],

      // ── MYNTRA PLANS ──────────────────────────────────────────
      ['PLAN','Myntra Marketplace','','','myn-basic',   'Basic Plan – 45 Days',    '35 Listings · Rs.400/listing',16520,'one-time','','','Yes','45 Days','No','','',1,'Yes',''],
      ['PLAN','Myntra Marketplace','','','myn-standard','Standard Plan – 45 Days', '50 Listings · Rs.400/listing',23600,'one-time','','','Yes','45 Days','No','','',2,'Yes',''],
      ['PLAN','Myntra Marketplace','','','myn-premium', 'Premium Plan – 50 Days',  '100 Listings · Rs.400/listing',47200,'one-time','','','Yes','50 Days','No','','',3,'Yes',''],

      ['ADDON','Myntra Marketplace','','','myn-inv',         'Stock & Inventory Management','',3000, 'monthly','','','No','','Yes','Myntra Marketplace','/month',1,'Yes',''],
      ['ADDON','Myntra Marketplace','','','myn-ads-with',    'Ads with Plan',              '',5000, 'monthly','','Ads Budget','No','','Yes','Myntra Marketplace','/month + ads budget',2,'Yes',''],
      ['ADDON','Myntra Marketplace','','','myn-ads-without', 'Ads without Plan',           '',10000,'monthly','','Ads Budget','No','','Yes','Myntra Marketplace','/month + ads budget',3,'Yes',''],

      // ── TRADEINDIA / INDIAMART PLANS ──────────────────────────
      ['PLAN','TradeIndia / IndiaMart','','','ti-basic',   'Basic Plan – 45 Days',    '35 Listings · Rs.340/listing',14042,'one-time','','','Yes','45 Days','No','','',1,'Yes',''],
      ['PLAN','TradeIndia / IndiaMart','','','ti-standard','Standard Plan – 45 Days', '50 Listings · Rs.340/listing',20060,'one-time','','','Yes','45 Days','No','','',2,'Yes',''],
      ['PLAN','TradeIndia / IndiaMart','','','ti-premium', 'Premium Plan – 50 Days',  '100 Listings · Rs.340/listing',40120,'one-time','','','Yes','50 Days','No','','',3,'Yes',''],

      ['ADDON','TradeIndia / IndiaMart','','','ti-inv',         'Stock & Inventory Management','',3000, 'monthly','','','No','','Yes','TradeIndia / IndiaMart','/month',1,'Yes',''],
      ['ADDON','TradeIndia / IndiaMart','','','ti-ads-with',    'Ads with Plan',              '',5000, 'monthly','','Ads Budget','No','','Yes','TradeIndia / IndiaMart','/month + ads budget',2,'Yes',''],
      ['ADDON','TradeIndia / IndiaMart','','','ti-ads-without', 'Ads without Plan',           '',10000,'monthly','','Ads Budget','No','','Yes','TradeIndia / IndiaMart','/month + ads budget',3,'Yes',''],

      // ── COMPLETE ENABLEMENT BUNDLE PLANS ──────────────────────
      ['PLAN','Complete Enablement Bundle','','','bundle-50', 'Amazon + Flipkart + Client-Choice (50 Products Each)', '3 Platforms · 50 products each',70800,'one-time','','','Yes','60 Days','No','','',1,'Yes',''],
      ['PLAN','Complete Enablement Bundle','','','bundle-100','Amazon + Flipkart + Client-Choice (100 Products Each)','3 Platforms · 100 products each',120360,'one-time','','','Yes','75 Days','No','','',2,'Yes',''],

      // ── E-COMMERCE DEVELOPMENT PLANS ─────────────────────────
      ['PLAN','E-Commerce Development','','','ecom-shopify-basic',    'Shopify Store - Starter',               'Up to 50 products, theme setup, basic pages, payment gateway',      24999,'one-time','','','No','15-20 Days','No','','',1,'Yes',''],
      ['PLAN','E-Commerce Development','','','ecom-shopify-pro',      'Shopify Store - Professional',          'Up to 200 products, custom theme, apps integration, SEO setup',     49999,'one-time','','','No','25-35 Days','No','','',2,'Yes',''],
      ['PLAN','E-Commerce Development','','','ecom-wordpress-basic',  'WordPress WooCommerce - Starter',       'Up to 100 products, premium theme, payment setup, basic customization',19999,'one-time','','Hosting charges extra','No','20-25 Days','No','','',3,'Yes',''],
      ['PLAN','E-Commerce Development','','','ecom-wordpress-pro',    'WordPress WooCommerce - Professional',  'Unlimited products, custom design, advanced filters, speed optimization',44999,'one-time','','Hosting charges extra','No','30-40 Days','No','','',4,'Yes',''],
      ['PLAN','E-Commerce Development','','','ecom-custom',           'Custom E-Commerce Platform',            'Fully custom built, scalable architecture, admin panel, mobile responsive',99999,'one-time','','','No','45-60 Days','No','','',5,'Yes',''],

      // ── E-COMMERCE ADD-ONS ─────────────────────────────────────
      ['ADDON','E-Commerce Development','','','ecom-payment',     'Payment Gateway Integration (Razorpay / PayU / Cashfree)','One-time',4999,'one-time','','','No','','Yes','E-Commerce Development','One-time',1,'Yes',''],
      ['ADDON','E-Commerce Development','','','ecom-shipping',    'Shipping Integration (Shiprocket / Delhivery / Pickrr)',  'One-time',3999,'one-time','','','No','','Yes','E-Commerce Development','One-time',2,'Yes',''],
      ['ADDON','E-Commerce Development','','','ecom-maintenance', 'Monthly Maintenance and Support',                         '/month',  4999,'monthly', '','','No','','Yes','E-Commerce Development','/month',  3,'Yes',''],
      ['ADDON','E-Commerce Development','','','ecom-seo',         'On-Page SEO Setup for Store',                             'One-time',7999,'one-time','','','No','','Yes','E-Commerce Development','One-time',4,'Yes',''],
      ['ADDON','E-Commerce Development','','','ecom-catalog',     'Product Catalog Management (per 50 products)',            'Per batch',2999,'one-time','','','No','','Yes','E-Commerce Development','Per batch',5,'Yes',''],
      ['ADDON','E-Commerce Development','','','ecom-whatsapp',    'WhatsApp Commerce Integration',                           'One-time',5999,'one-time','','','No','','Yes','E-Commerce Development','One-time',6,'Yes',''],

      // ── CUSTOM PACKAGE ────────────────────────────────────────
      ['PLAN','Custom Package','','','custom','Custom Package','As per client requirements',0,'custom','','','No','','No','','',1,'Yes','']
    ];

    if (seed.length > 0) {
      sh.getRange(2, 1, seed.length, QB_HEADERS.length).setValues(seed);
    }

    // Apply dropdowns
    var ruleRowType = SpreadsheetApp.newDataValidation().requireValueInList(['SETTINGS','CATEGORY','PLAN','ADDON'],true).setAllowInvalid(true).build();
    sh.getRange(2, 1, 2000, 1).setDataValidation(ruleRowType);
    var ruleCatType = SpreadsheetApp.newDataValidation().requireValueInList(['service','product'],true).setAllowInvalid(true).build();
    sh.getRange(2, 3, 2000, 1).setDataValidation(ruleCatType);
    var rulePriceType = SpreadsheetApp.newDataValidation().requireValueInList(['monthly','one-time','per-listing','per-unit','custom'],true).setAllowInvalid(true).build();
    sh.getRange(2, 9, 2000, 1).setDataValidation(rulePriceType);
    var ruleYesNo = SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'],true).setAllowInvalid(true).build();
    sh.getRange(2, 12, 2000, 1).setDataValidation(ruleYesNo);
    sh.getRange(2, 14, 2000, 1).setDataValidation(ruleYesNo);
    sh.getRange(2, 18, 2000, 1).setDataValidation(ruleYesNo);

    SpreadsheetApp.flush();
    Logger.log('QUOTATION_BUILDER sheet initialized with ' + seed.length + ' rows.');
    return sh;
  } catch(e) {
    Logger.log('_initQuotationBuilderSheet error: ' + e.message);
    return null;
  }
}

function getQuotationBuilderConfig() {
  try {
    var sh = _ensureQuotationBuilderSheet();
    if (!sh || sh.getLastRow() < 2) return _getDefaultQBConfig();

    var data = sh.getDataRange().getValues();
    var h    = data[0];

    function col(k) { var i = h.indexOf(k); return i >= 0 ? i : -1; }
    function val(row, k) { return String(row[col(k)] !== undefined ? row[col(k)] : '').trim(); }
    function num(row, k) { return parseFloat(String(row[col(k)] || 0)) || 0; }
    function bool(row, k) { return val(row, k).toLowerCase() === 'yes'; }

    var mode    = 'service';
    var catalog = {};
    var catOrder = [];

    for (var i = 1; i < data.length; i++) {
      var row     = data[i];
      var rowType = val(row, 'RowType').toUpperCase();
      if (!rowType) continue;

      if (rowType === 'SETTINGS') {
        var mv = val(row, 'CategoryName').toLowerCase();
        if (mv === 'service' || mv === 'product' || mv === 'both') mode = mv;
        continue;
      }

      if (rowType === 'CATEGORY') {
        var catName = val(row, 'CategoryName');
        if (!catName) continue;
        var catType = val(row, 'CategoryType') || 'service';
        var catIcon = val(row, 'Icon');
        if (!catalog[catName]) {
          catalog[catName] = { type: catType, icon: catIcon, plans: [], addons: [], sortOrder: num(row, 'SortOrder') || 99 };
          catOrder.push(catName);
        }
        continue;
      }

      if (rowType === 'PLAN') {
        var catName = val(row, 'CategoryName');
        var active  = bool(row, 'Active');
        if (!catName || !active) continue;
        if (!catalog[catName]) { catalog[catName] = { type: 'service', icon: '', plans: [], addons: [], sortOrder: 99 }; catOrder.push(catName); }

        var price     = num(row, 'Price');
        var priceType = val(row, 'PriceType') || 'monthly';
        var retainer  = num(row, 'Retainer') || null;
        var addlCost  = val(row, 'AdditionalCost') || null;
        var taxInc    = bool(row, 'TaxIncluded');
        var timeline  = val(row, 'Timeline') || null;
        var desc      = val(row, 'Description');

        var plan = {
          id:             val(row, 'ItemID'),
          name:           val(row, 'ItemName'),
          coverage:       desc,
          platforms:      desc,
          price:          price,
          priceType:      priceType,
          retainer:       retainer,
          additionalCost: addlCost,
          taxIncluded:    taxInc,
          timeline:       timeline,
          sortOrder:      num(row, 'SortOrder') || 99
        };

        if (priceType === 'monthly')      plan.monthly      = price;
        else if (priceType === 'one-time') plan.total        = price;
        else if (priceType === 'per-listing') {
          plan.total = price;
          var lm = desc.match(/^(\d+)\s*[Ll]isting/);
          if (lm) { plan.listings = parseInt(lm[1]); var pp = desc.match(/Rs\.(\d+)\/listing/i); plan.pricePerListing = pp ? parseInt(pp[1]) : 0; }
        }
        if (taxInc) plan.totalWithTax = price;

        catalog[catName].plans.push(plan);
        continue;
      }

      if (rowType === 'ADDON') {
        var catName = val(row, 'CategoryName');
        var active  = bool(row, 'Active');
        if (!catName || !active) continue;
        if (!catalog[catName]) { catalog[catName] = { type: 'service', icon: '', plans: [], addons: [], sortOrder: 99 }; catOrder.push(catName); }

        catalog[catName].addons.push({
          id:          val(row, 'ItemID'),
          name:        val(row, 'ItemName'),
          price:       num(row, 'Price'),
          note:        val(row, 'AddonNote'),
          taxIncluded: bool(row, 'TaxIncluded')
        });
        continue;
      }
    }

    // Sort plans within each category by sortOrder
    Object.keys(catalog).forEach(function(cat) {
      catalog[cat].plans.sort(function(a, b) { return (a.sortOrder || 99) - (b.sortOrder || 99); });
    });

    // Ensure Custom Package exists
    if (!catalog['Custom Package']) {
      catalog['Custom Package'] = { type: 'service', icon: '', plans: [{ id: 'custom', name: 'Custom Package', price: null }], addons: [], sortOrder: 999 };
      catOrder.push('Custom Package');
    }

    // Build ordered catalog
    var orderedCatalog = {};
    catOrder.forEach(function(cat) { if (catalog[cat]) orderedCatalog[cat] = catalog[cat]; });
    // Append any cats not in catOrder
    Object.keys(catalog).forEach(function(cat) { if (!orderedCatalog[cat]) orderedCatalog[cat] = catalog[cat]; });

    return { success: true, mode: mode, catalog: orderedCatalog };
  } catch(e) {
    Logger.log('getQuotationBuilderConfig error: ' + e.message);
    return _getDefaultQBConfig();
  }
}

function saveQuotationBuilderRow(entry) {
  try {
    var sh = _ensureQuotationBuilderSheet();
    if (!sh) return { success: false, error: 'Sheet not found' };
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var idIdx = h.indexOf('ItemID');
    var now   = new Date().toISOString();

    var row = [
      entry.rowType     || 'PLAN',
      entry.categoryName || '',
      entry.categoryType || 'service',
      entry.icon        || '',
      entry.itemId      || ('ITEM-' + Date.now()),
      entry.itemName    || '',
      entry.description || '',
      parseFloat(entry.price)    || 0,
      entry.priceType   || 'monthly',
      parseFloat(entry.retainer) || '',
      entry.additionalCost || '',
      entry.taxIncluded ? 'Yes' : 'No',
      entry.timeline    || '',
      entry.isAddon     ? 'Yes' : 'No',
      entry.addonFor    || '',
      entry.addonNote   || '',
      parseInt(entry.sortOrder) || 99,
      entry.active !== false ? 'Yes' : 'No',
      entry.notes || ''
    ];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(entry.itemId || '')) {
        sh.getRange(i + 1, 1, 1, QB_HEADERS.length).setValues([row]);
        SpreadsheetApp.flush();
        return { success: true, updated: true };
      }
    }
    sh.appendRow(row);
    SpreadsheetApp.flush();
    return { success: true, created: true };
  } catch(e) {
    Logger.log('saveQuotationBuilderRow error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function deleteQuotationBuilderRow(itemId) {
  try {
    var sh = _ensureQuotationBuilderSheet();
    if (!sh) return { success: false };
    var data  = sh.getDataRange().getValues();
    var idIdx = data[0].indexOf('ItemID');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(itemId)) {
        sh.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Row not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function _getDefaultQBConfig() {
  return {
    success: true,
    mode: 'service',
    catalog: {
      'Custom Package': { type: 'service', icon: '', plans: [{ id: 'custom', name: 'Custom Package', price: null }], addons: [] }
    }
  };
}

// ═══════════════════════════════════════════════════════════════
//  DUMMY / FALLBACK DATA
// ═══════════════════════════════════════════════════════════════

function _dummyLeads() {
  const n = _now();
  return [
    ['L001',n,'Tahani Al-Hashemi','tanhaali@gmail.com','+971505843244','Zendesk','zendesk.com','Technology','LinkedIn','In Progress','Proposal Sent','High','No','','Dubai','','UAE',85000,'Initial contact','Called','',0,0,'No','','','System','',''],
    ['L002',n,'Alexe Jordan','alexejordan@gmail.com','+966558441493','Quicken Loans','quickenloans.com','Financial','Cold Call','In Progress','Contacted','Medium','No','','Riyadh','','KSA',120000,'Called once','Call Made','',0,0,'No','','','System','',''],
    ['L003',n,'Shouq Al-Kumaiti','kuwait764@gmail.com','+971503918260','Audi','audi.com','Technology','WhatsApp','New','New Lead','Critical','Yes','','Kuwait City','','Kuwait',200000,'First touch','Added','',0,0,'No','','Urgent','System','',''],
    ['L004',n,'Robert Fox','robertfoxx23@gmail.com','+201000239271','Zendesk','zendesk.com','Technology','Website','Qualified','Qualified','High','Yes','','Bangalore','','India',175000,'Demo done','Demo Done','',0,0,'No','','','System','',''],
    ['L005',n,'Raj Sharma','raj.sharma@builders.co','+919123456789','Skyline Builders','skylinebuilders.co','Real Estate','Referral','Contacted','Interested','High','Yes','','Pune','Maharashtra','India',500000,'Called','Follow Up','',0,0,'No','','','System','','']
  ];
}

function _dummyEmployees() {
  const n = _now();
  return [
    ['EMP001','Arjun Mehta','arjun@afterresult.com','+919911223344','Senior BDE','Senior (3-5 yr)','Technology,E-Commerce','15',0,0,'Active','09:00','18:00',60,'No',n,n,'A',0,0,'Sales','',''],
    ['EMP002','Kavya Nair','kavya@afterresult.com','+919922334455','BDE','Mid (1-3 yr)','Real Estate,Healthcare','12',0,0,'Active','10:00','19:00',60,'No',n,n,'K',0,0,'Sales','','']
  ];
}

function _fallbackLeads() {
  const n = _now();
  return [
    { id:'L001', name:'Tahani Al-Hashemi', phone:'+971505843244', email:'tanhaali@gmail.com', company:'Zendesk', domain:'zendesk.com', industry:'Technology', status:'In Progress', priority:'High', source:'LinkedIn', city:'Dubai', deal:85000, hot:false, addedAt:n, notes:[], funnelStage:'Proposal Sent' },
    { id:'L002', name:'Alexe Jordan', phone:'+966558441493', email:'alexejordan@gmail.com', company:'Quicken Loans', domain:'quickenloans.com', industry:'Financial', status:'In Progress', priority:'Medium', source:'Cold Call', city:'Riyadh', deal:120000, hot:false, addedAt:n, notes:[], funnelStage:'Contacted' },
    { id:'L003', name:'Shouq Al-Kumaiti', phone:'+971503918260', email:'kuwait764@gmail.com', company:'Audi', domain:'audi.com', industry:'Technology', status:'New', priority:'Critical', source:'WhatsApp', city:'Kuwait', deal:200000, hot:true, addedAt:n, notes:[], funnelStage:'New Lead' },
    { id:'L004', name:'Robert Fox', phone:'+201000239271', email:'robertfoxx23@gmail.com', company:'Zendesk', domain:'zendesk.com', industry:'Technology', status:'Qualified', priority:'High', source:'Website', city:'Bangalore', deal:175000, hot:true, addedAt:n, notes:[], funnelStage:'Qualified' },
    { id:'L005', name:'Raj Sharma', phone:'+919123456789', email:'raj.sharma@builders.co', company:'Skyline Builders', domain:'skylinebuilders.co', industry:'Real Estate', status:'Contacted', priority:'High', source:'Referral', city:'Pune', deal:500000, hot:true, addedAt:n, notes:[], funnelStage:'Interested' }
  ];
}

function _setupOnOpenTrigger() {
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'onOpen') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('onOpen').forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════
//  COMPANY REGISTRY — tracks all admins + their employees
// ═══════════════════════════════════════════════════════════════

function _ensureCompaniesSheet(ss) {
  var sh = ss.getSheetByName('COMPANIES');
  if (!sh) {
    sh = ss.insertSheet('COMPANIES');
    sh.appendRow([
      'Company ID','Company Name','Admin Name','Admin Email','Phone',
      'Industry','Website','Company Size','Plan','Emp Count',
      'Script URL','Status','Registered At','Updated At'
    ]);
    sh.getRange(1,1,1,14)
      .setBackground('#111827').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 14, 150);
  }
  return sh;
}

function saveCompany(entry) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = _ensureCompaniesSheet(ss);
    var data = sh.getDataRange().getValues();
    var now  = new Date().toISOString();
    var row  = [
      entry.id            || 'CO-' + Date.now(),
      entry.company       || '',
      entry.adminName     || '',
      entry.email         || '',
      entry.phone         || '',
      entry.industry      || '',
      entry.website       || '',
      entry.size          || '',
      entry.plan          || 'Trial',
      entry.empCount      || 0,
      entry.scriptUrl     || getSheetUrl(),
      entry.status        || 'Active',
      entry.registeredAt  || now,
      now
    ];
    // Update if admin email already exists
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).toLowerCase() === String(entry.email || '').toLowerCase()) {
        sh.getRange(i + 1, 1, 1, 14).setValues([row]);
        return { success: true, updated: true };
      }
    }
    sh.appendRow(row);
    return { success: true, created: true };
  } catch(e) {
    Logger.log('saveCompany error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getAllCompanies() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = _ensureCompaniesSheet(ss);
    if (sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return {
          id:           String(r[0]),
          company:      String(r[1]),
          adminName:    String(r[2]),
          email:        String(r[3]),
          phone:        String(r[4]),
          industry:     String(r[5]),
          website:      String(r[6]),
          size:         String(r[7]),
          plan:         String(r[8]),
          empCount:     Number(r[9] || 0),
          scriptUrl:    String(r[10]),
          status:       String(r[11]),
          registeredAt: String(r[12])
        };
      });
  } catch(e) { return []; }
}

// ── Admin config (stores gsUrl, plan, profile per admin) ──────────
function saveAdminConfig(config) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('ADMIN_CONFIG') || ss.insertSheet('ADMIN_CONFIG');
    var data = sh.getDataRange().getValues();
    if (data.length < 1 || !data[0][0]) {
      sh.appendRow(['Admin Email','Admin Name','Company','Phone','Industry',
                    'Website','Size','Plan','Script URL','Updated At']);
      sh.getRange(1,1,1,10).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
      sh.setFrozenRows(1);
      data = sh.getDataRange().getValues();
    }
    var now = new Date().toISOString();
    var row = [
      config.email || '', config.adminName || '', config.companyName || '',
      config.phone || '', config.industry || '', config.website || '',
      config.size  || '', config.plan     || 'Free',
      config.gsUrl || getSheetUrl(), now
    ];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(config.email || '').toLowerCase()) {
        sh.getRange(i + 1, 1, 1, 10).setValues([row]);
        return { success: true };
      }
    }
    sh.appendRow(row);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function getAdminConfig(email) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('ADMIN_CONFIG');
    if (!sh || sh.getLastRow() < 2) return null;
    var data = sh.getDataRange().getValues();
    var norm = String(email || '').toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === norm) {
        return {
          email:       String(data[i][0]), adminName:   String(data[i][1]),
          companyName: String(data[i][2]), phone:       String(data[i][3]),
          industry:    String(data[i][4]), website:     String(data[i][5]),
          size:        String(data[i][6]), plan:        String(data[i][7]),
          gsUrl:       String(data[i][8])
        };
      }
    }
    return null;
  } catch(e) { return null; }
}

// ── Cascade delete admin + all their employees ────────────────────
// Call this when an admin's account is dissolved from the panel
function deleteAdminAccount(adminEmail) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var norm = String(adminEmail || '').trim().toLowerCase();
    var deleted = { employees: 0, authUsers: 0, companiesRemoved: 0 };

    // 1. Collect all employee emails belonging to this admin from EMPLOYEES sheet
    var empSh = ss.getSheetByName(SN.EMPLOYEES);
    var empEmails = [];
    if (empSh && empSh.getLastRow() > 1) {
      var empData = empSh.getDataRange().getValues();
      for (var i = empData.length - 1; i >= 1; i--) {
        if (empData[i][0]) {
          empEmails.push(String(empData[i][2] || '').toLowerCase());
          empSh.deleteRow(i + 1);
          deleted.employees++;
        }
      }
    }

    // 2. Remove from AUTHORIZED_USERS — admin + all employees
    var authSh = ss.getSheetByName(SN.AUTH);
    if (authSh && authSh.getLastRow() > 1) {
      var authData = authSh.getDataRange().getValues();
      for (var j = authData.length - 1; j >= 1; j--) {
        var rowEmail = String(authData[j][0] || '').toLowerCase();
        if (rowEmail === norm || empEmails.indexOf(rowEmail) >= 0) {
          authSh.deleteRow(j + 1);
          deleted.authUsers++;
        }
      }
    }

    // 3. Mark company as Dissolved in COMPANIES sheet
    var coSh = ss.getSheetByName('COMPANIES');
    if (coSh && coSh.getLastRow() > 1) {
      var coData = coSh.getDataRange().getValues();
      for (var k = 1; k < coData.length; k++) {
        if (String(coData[k][3] || '').toLowerCase() === norm) {
          coSh.getRange(k + 1, 12).setValue('Dissolved');
          coSh.getRange(k + 1, 14).setValue(new Date().toISOString());
          deleted.companiesRemoved++;
        }
      }
    }

    // 4. Log dissolution
    _logAction({
      leadId:  '',
      action:  'ADMIN_ACCOUNT_DISSOLVED',
      details: 'Admin ' + adminEmail + ' dissolved. ' + deleted.employees + ' employees removed.',
      email:   adminEmail
    });

    SpreadsheetApp.flush();
    Logger.log('[deleteAdminAccount] ' + JSON.stringify(deleted));
    return { success: true, ...deleted };
  } catch(e) {
    Logger.log('deleteAdminAccount error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  COMPANY REGISTRY — tracks all admins + their employees
// ═══════════════════════════════════════════════════════════════

function _ensureCompaniesSheet(ss) {
  var sh = ss.getSheetByName('COMPANIES');
  if (!sh) {
    sh = ss.insertSheet('COMPANIES');
    sh.appendRow([
      'Company ID','Company Name','Admin Name','Admin Email','Phone',
      'Industry','Website','Company Size','Plan','Emp Count',
      'Script URL','Status','Registered At','Updated At'
    ]);
    sh.getRange(1,1,1,14)
      .setBackground('#111827').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 14, 150);
  }
  return sh;
}

function saveCompany(entry) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = _ensureCompaniesSheet(ss);
    var data = sh.getDataRange().getValues();
    var now  = new Date().toISOString();
    var row  = [
      entry.id            || 'CO-' + Date.now(),
      entry.company       || '',
      entry.adminName     || '',
      entry.email         || '',
      entry.phone         || '',
      entry.industry      || '',
      entry.website       || '',
      entry.size          || '',
      entry.plan          || 'Trial',
      entry.empCount      || 0,
      entry.scriptUrl     || getSheetUrl(),
      entry.status        || 'Active',
      entry.registeredAt  || now,
      now
    ];
    // Update if admin email already exists
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).toLowerCase() === String(entry.email || '').toLowerCase()) {
        sh.getRange(i + 1, 1, 1, 14).setValues([row]);
        return { success: true, updated: true };
      }
    }
    sh.appendRow(row);
    return { success: true, created: true };
  } catch(e) {
    Logger.log('saveCompany error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getAllCompanies() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = _ensureCompaniesSheet(ss);
    if (sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return {
          id:           String(r[0]),
          company:      String(r[1]),
          adminName:    String(r[2]),
          email:        String(r[3]),
          phone:        String(r[4]),
          industry:     String(r[5]),
          website:      String(r[6]),
          size:         String(r[7]),
          plan:         String(r[8]),
          empCount:     Number(r[9] || 0),
          scriptUrl:    String(r[10]),
          status:       String(r[11]),
          registeredAt: String(r[12])
        };
      });
  } catch(e) { return []; }
}

// ── Admin config (stores gsUrl, plan, profile per admin) ──────────


function _applyRoutingRulesToRow(rowData, headers, sheet, sheetRowNumber) {
  try {
    const rulesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ROUTING_RULES');
    if (!rulesSheet || rulesSheet.getLastRow() < 2) return;

    const rulesData = rulesSheet.getDataRange().getValues();
    const rules = rulesData.slice(1).filter(r => r[0] && r[10] === 'Yes').map(r => ({
      id: String(r[0]), name: String(r[1]), trigger: String(r[2]),
      triggerValue: String(r[3]), empId: String(r[4]),
      empName: String(r[5]), empEmail: String(r[6]),
      override: r[7] === 'Yes', notify: r[9] === 'Yes'
    }));

    if (!rules.length) return;

    const col = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i]).toLowerCase().replace(/[\s_\-]/g, '');
        for (const k of (Array.isArray(keywords) ? keywords : [keywords])) {
          if (h.includes(k.toLowerCase().replace(/[\s_\-]/g, ''))) return i;
        }
      }
      return -1;
    };

    const hotCol           = col(['hot','ishot','hotlead']);
    const statusCol        = col(['status','leadstatus']);
    const priorityCol      = col(['priority']);
    const industryCol      = col(['industry','service','category']);
    const assignedToCol    = col(['assignedto','assigned_to','assigned']);
    const assignedEmailCol = col(['assignedemail','assigned_email']);
    const assignedEmpIdCol = col(['assignedempid','assigned_empid']);
    const nameCol          = col(['name','fullname','contactname']);
    const leadIdCol        = col(['leadid','lead_id','id']);

    if (assignedToCol === -1) return;

    const hotVal      = hotCol      >= 0 ? String(rowData[hotCol]  || '').toLowerCase() : '';
    const statusVal   = statusCol   >= 0 ? String(rowData[statusCol]   || '') : '';
    const priorityVal = priorityCol >= 0 ? String(rowData[priorityCol] || '') : '';
    const industryVal = industryCol >= 0 ? String(rowData[industryCol] || '') : '';
    const leadName    = nameCol     >= 0 ? String(rowData[nameCol]     || '') : '';
    const leadId      = leadIdCol   >= 0 ? String(rowData[leadIdCol]   || '') : '';

    for (const rule of rules) {
      let matches = false;
      if (rule.trigger === 'hot'      && (hotVal === 'yes' || hotVal === 'true' || hotVal === '1')) matches = true;
      if (rule.trigger === 'status'   && statusVal   === rule.triggerValue) matches = true;
      if (rule.trigger === 'priority' && priorityVal === rule.triggerValue) matches = true;
      if (rule.trigger === 'industry' && industryVal === rule.triggerValue) matches = true;

      if (!matches) continue;

      // Find target employee in EMPLOYEES sheet
      const empSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EMPLOYEES');
      if (!empSh) continue;
      const empData = empSh.getDataRange().getValues();
      let targetEmp = null;
      for (let ei = 1; ei < empData.length; ei++) {
        const empEmail = String(empData[ei][2] || '').trim().toLowerCase();
        const empId    = String(empData[ei][0] || '').trim();
        const empStatus= String(empData[ei][10]|| 'Active');
        if ((empId === rule.empId || empEmail === rule.empEmail.toLowerCase()) && empStatus === 'Active') {
          targetEmp = { id: empId, name: String(empData[ei][1]), email: empEmail };
          break;
        }
      }

      if (!targetEmp) {
        Logger.log('[ROUTING] Target employee not found or inactive for rule: ' + rule.name);
        continue;
      }

      // Skip if already assigned to this employee
      const curAssigned = assignedEmailCol >= 0
        ? String(rowData[assignedEmailCol] || '').toLowerCase()
        : String(rowData[assignedToCol]    || '').toLowerCase();

      if (curAssigned === targetEmp.email) {
        Logger.log('[ROUTING] Lead already assigned to ' + targetEmp.name + ', skipping.');
        continue;
      }

      // Write new assignment directly to sheet
      sheet.getRange(sheetRowNumber, assignedToCol    + 1).setValue(targetEmp.email);
      if (assignedEmailCol >= 0) sheet.getRange(sheetRowNumber, assignedEmailCol + 1).setValue(targetEmp.email);
      if (assignedEmpIdCol >= 0) sheet.getRange(sheetRowNumber, assignedEmpIdCol + 1).setValue(targetEmp.id);

      // Update rule fire count and last fired timestamp
      for (let ri = 1; ri < rulesData.length; ri++) {
        if (String(rulesData[ri][0]) === rule.id) {
          rulesSheet.getRange(ri + 1, 12).setValue((Number(rulesData[ri][11]) || 0) + 1);
          rulesSheet.getRange(ri + 1, 14).setValue(new Date().toISOString());
          break;
        }
      }

      // Notify escalations channel
      if (rule.notify) {
        try {
          const msg = '[AUTO-ROUTE] "' + leadName + '" triggered rule "' + rule.name
            + '" (' + rule.trigger + (rule.triggerValue ? ': ' + rule.triggerValue : '')
            + '). Reassigned to ' + targetEmp.name + '.';
          sendChat('routing@system', null, msg, 'escalations', {
            type: 'system', fromName: 'Auto-Router'
          });
        } catch(e) {}
      }

      SpreadsheetApp.flush();
      Logger.log('[ROUTING] Rule "' + rule.name + '" fired → lead "' + leadName + '" reassigned to ' + targetEmp.name + ' (' + targetEmp.email + ')');
      break;
    }
  } catch(e) {
    Logger.log('[_applyRoutingRulesToRow] Error: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════
//  SHIFT EVENTS — logShiftEvent, getShiftEvents, markEmployeeActive
//  These are called by the employer panel (index.html / index__3_.html)
// ═══════════════════════════════════════════════════════════════

// ─── SHIFT EVENTS (called by employer panel) ─────────────────
function logShiftEvent(data) {
  if (!data) return { success: false, error: 'No data' };
  return logShift(
    String(data.empEmail || ''),
    String(data.empId    || ''),
    String(data.event    || ''),
    { totalBreakSeconds: data.totalBreakSeconds || 0,
      totalSeconds:      data.totalSeconds      || 0 }
  );
}

function getShiftEvents(sinceHours) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('ShiftEvents');
    if (!sh || sh.getLastRow() < 2) return [];
    const since   = new Date(Date.now() - ((sinceHours || 24) * 3600000));
    const rawData = sh.getDataRange().getValues();
    const firstCell = String(rawData[0][0] || '').toLowerCase();
    const hasHeader = firstCell.includes('emp') || firstCell.includes('id');
    const headers   = hasHeader
      ? rawData[0].map(x => String(x).toLowerCase().trim().replace(/[\s_\-]/g,''))
      : ['empid','empname','empemail','event','breakminutes','totalbreakseconds','time'];
    const rows = hasHeader ? rawData.slice(1) : rawData;
    const ci = name => { const i = headers.indexOf(name); return i >= 0 ? i : -1; };
    const results = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r || !r[0]) continue;
      const tv = r[ci('time')] || r[6];
      if (!tv) continue;
      const t = new Date(tv);
      if (isNaN(t.getTime()) || t < since) continue;
      results.push({
        empId:             String(r[ci('empid')]             || r[0] || ''),
        empName:           String(r[ci('empname')]           || r[1] || ''),
        empEmail:          String(r[ci('empemail')]          || r[2] || '').toLowerCase().trim(),
        event:             String(r[ci('event')]             || r[3] || '').toLowerCase().trim(),
        breakMinutes:      Number(r[ci('breakminutes')]      || r[4] || 0),
        totalBreakSeconds: Number(r[ci('totalbreakseconds')] || r[5] || 0),
        time:              t.toISOString()
      });
    }
    return results;
  } catch(e) {
    Logger.log('getShiftEvents error: ' + e.message);
    return [];
  }
}

function markEmployeeActive(email) {
  if (!email) return { success: false };
  return logShift(String(email), '', 'shift_start', {});
}

// ─── USER ACTIVITY — one row per user, upserted on each status change ───
function updateUserActivity(data) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const email  = String(data.email || '').toLowerCase().trim();
    const status = data.forceStatus || 'Online';
    if (!email) return { success: false };
    _upsertUserActivity(ss, email, data.empId || '', data.name || '', status, 'employer', new Date().toISOString());
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getUserActivity() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('USER_ACTIVITY');
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getDataRange().getValues();
    const h    = data[0].map(x => String(x).toLowerCase().trim());
    return data.slice(1).filter(r => r[0]).map(r => ({
      email:    String(r[h.indexOf('email')]    || ''),
      empId:    String(r[h.indexOf('empid')]    || ''),
      name:     String(r[h.indexOf('name')]     || ''),
      role:     String(r[h.indexOf('role')]     || ''),
      status:   String(r[h.indexOf('status')]   || ''),
      lastSeen: String(r[h.indexOf('lastseen')] || '')
    }));
  } catch(e) { return []; }
}

function _upsertUserActivity(ss, email, empId, name, status, role, time) {
  try {
    let sh = ss.getSheetByName('USER_ACTIVITY');
    if (!sh) {
      sh = ss.insertSheet('USER_ACTIVITY');
      sh.appendRow(['Email','EmpId','Name','Role','Status','LastSeen']);
      sh.getRange(1,1,1,6).setBackground('#111827').setFontColor('#ffffff').setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(1, 220); sh.setColumnWidth(3, 160); sh.setColumnWidth(5, 110); sh.setColumnWidth(6, 180);
    }
    const data  = sh.getDataRange().getValues();
    const h     = data[0].map(x => String(x).toLowerCase().trim());
    const eIdx  = h.indexOf('email');
    const iIdx  = h.indexOf('empid');
    const nIdx  = h.indexOf('name');
    const rIdx  = h.indexOf('role');
    const sIdx  = h.indexOf('status');
    const lsIdx = h.indexOf('lastseen');
    const norm  = email.toLowerCase().trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][eIdx] || '').toLowerCase().trim() === norm) {
        if (sIdx  >= 0) sh.getRange(i+1, sIdx+1).setValue(status);
        if (lsIdx >= 0) sh.getRange(i+1, lsIdx+1).setValue(time || new Date().toISOString());
        if (empId && iIdx >= 0 && !data[i][iIdx]) sh.getRange(i+1, iIdx+1).setValue(empId);
        SpreadsheetApp.flush();
        return;
      }
    }
    sh.appendRow([norm, empId||'', name||'', role||'employer', status, time||new Date().toISOString()]);
    SpreadsheetApp.flush();
  } catch(e) { Logger.log('_upsertUserActivity error: ' + e.message); }
}
