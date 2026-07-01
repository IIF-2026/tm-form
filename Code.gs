/**
 * Think & Make — PWA form backend
 * Inqui-Lab Foundation
 *
 * One flat row per submission. Routes by formId to a named tab.
 * Form 1 -> "Form1_SchoolOrientation", Form 2 -> "Form2_..." etc.
 * The tab is created with headers on first submission, then rows are appended.
 *
 * SETUP:
 *   1. This script must live INSIDE the target spreadsheet
 *      (Extensions > Apps Script from the sheet), OR set SHEET_ID below.
 *   2. Deploy > New deployment > type "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   3. Copy the /exec URL into the PWA config (APPS_SCRIPT_URL).
 */

// If the script is bound to the sheet, leave this blank.
// Otherwise paste your sheet ID here.
var SHEET_ID = ''; // e.g. '1sVYBBrnSC-FDpnwmwwDLm_nqJbIYaea7iEvEtDd0kdw'

// Optional shared secret. Set the same value in the PWA config.
// Leave '' to disable the check.
var SECRET_TOKEN = 'TM2026SECRET';

// -------- Column definitions per form --------
// Order here IS the column order in the sheet. Add Form 2 later by adding a key.
var FORM_SCHEMAS = {
  form1_school_orientation: {
    tabName: 'Form1_SchoolOrientation',
    columns: [
      'Submission ID', 'Submitted At', 'Form Version',
      // Header
      'Partner', 'School', 'School Code', 'Visited By', 'Visit Date',
      // A — Location & Classification
      'Place (Mandal+Village)', 'Distance from HQ (km)', 'Gender Type',
      'School Type', 'Medium', 'Programme Year',
      // B — Principal & Teacher
      'Principal Name', 'Principal Phone', 'Principal Email',
      'Teacher Name', 'Teacher Phone', 'Teacher Email', 'Teacher Subject',
      // C — Enrollment summary (detail flattened into one cell + totals)
      'Grades', 'Total Sections', 'Est. Students (sum)',
      'Total Girls', 'Total Boys', 'Total SLs', 'Total Teams',
      'Enrollment Detail (JSON)',
      // D — Infrastructure
      'Lab Room', 'Internet', 'Smart Board', 'Kit Storage',
      // E — Schedule
      'Preferred Day', 'Time Slot', 'Blackout Dates',
      // G — IIF Assignment
      'Field Lead',
      // H — Photos & Location
      'School Photo URL', 'Classroom Photo URL', 'Maps Link',
      // I — Notes & Sign-off
      'Observations', 'Next Steps', 'Principal Acknowledged'
    ]
  }
  // form2_xxx: { tabName: 'Form2_...', columns: [...] }  <-- add later
};

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    if (SECRET_TOKEN && payload.token !== SECRET_TOKEN) {
      return json({ status: 'error', message: 'Unauthorised' });
    }

    var schema = FORM_SCHEMAS[payload.formId];
    if (!schema) {
      return json({ status: 'error', message: 'Unknown formId: ' + payload.formId });
    }

    var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID)
                      : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getOrCreateTab(ss, schema);

    var row = buildRow(payload, schema);
    sheet.appendRow(row);

    return json({ status: 'success', submissionId: payload.submissionId || '' });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
}

function doGet() {
  // Simple health check you can open in a browser to confirm deployment.
  return json({ status: 'ok', message: 'TM form backend live' });
}

function getOrCreateTab(ss, schema) {
  var sheet = ss.getSheetByName(schema.tabName);
  if (!sheet) {
    sheet = ss.insertSheet(schema.tabName);
    sheet.appendRow(schema.columns);
    var header = sheet.getRange(1, 1, 1, schema.columns.length);
    header.setFontWeight('bold')
          .setBackground('#0D3B4A')
          .setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function buildRow(p, schema) {
  var h = p.header || {};
  var a = p.sectionA || {};
  var b = p.sectionB || {};
  var c = p.sectionC || {};
  var d = p.sectionD || {};
  var eSec = p.sectionE || {};
  var g = p.sectionG || {};
  var hh = p.sectionH || {};
  var i = p.sectionI || {};

  // Flatten Section C: totals + a JSON dump of the full per-grade/section detail.
  var totals = tallyEnrollment(c);
  var grades = (c.grades || []).join(', ');

  var map = {
    'Submission ID': p.submissionId || '',
    'Submitted At': p.submittedAt || new Date().toISOString(),
    'Form Version': p.formVersion || '',
    'Partner': h.partner || '',
    'School': h.school || '',
    'School Code': h.schoolCode || '',
    'Visited By': h.visitedBy || '',
    'Visit Date': h.visitDate || '',
    'Place (Mandal+Village)': a.a1 || '',
    'Distance from HQ (km)': a.a2 || '',
    'Gender Type': a.a3 || '',
    'School Type': a.a4 || '',
    'Medium': a.a5 || '',
    'Programme Year': a.a6 || '',
    'Principal Name': b.b1 || '',
    'Principal Phone': b.b2 || '',
    'Principal Email': b.b3 || '',
    'Teacher Name': b.b4 || '',
    'Teacher Phone': b.b5 || '',
    'Teacher Email': b.b6 || '',
    'Teacher Subject': b.b7 || '',
    'Grades': grades,
    'Total Sections': totals.sections,
    'Est. Students (sum)': totals.students,
    'Total Girls': totals.girls,
    'Total Boys': totals.boys,
    'Total SLs': totals.sls,
    'Total Teams': totals.teams,
    'Enrollment Detail (JSON)': JSON.stringify(c.perGrade || {}),
    'Lab Room': d.d1 || '',
    'Internet': d.d2 || '',
    'Smart Board': d.d3 || '',
    'Kit Storage': d.d4 || '',
    'Preferred Day': (eSec.e1 || []).join ? (eSec.e1 || []).join(', ') : (eSec.e1 || ''),
    'Time Slot': eSec.e2 || '',
    'Blackout Dates': eSec.e3 || '',
    'Field Lead': g.fieldLead || '',
    'School Photo URL': hh.schoolPhotoUrl || '',
    'Classroom Photo URL': hh.classroomPhotoUrl || '',
    'Maps Link': hh.mapsLink || '',
    'Observations': i.i1 || '',
    'Next Steps': i.i2 || '',
    'Principal Acknowledged': i.i3 || ''
  };

  return schema.columns.map(function (col) {
    return map[col] !== undefined ? map[col] : '';
  });
}

function tallyEnrollment(c) {
  var t = { sections: 0, students: 0, girls: 0, boys: 0, sls: 0, teams: 0 };
  var perGrade = c.perGrade || {};
  Object.keys(perGrade).forEach(function (gr) {
    var arr = (perGrade[gr] && perGrade[gr].perSection) || [];
    arr.forEach(function (s) {
      t.sections += 1;
      t.students += Number(s.total) || 0;
      t.girls += Number(s.girls) || 0;
      t.boys += Number(s.boys) || 0;
      t.sls += Number(s.sls) || 0;
      t.teams += Number(s.teams) || 0;
    });
  });
  return t;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
