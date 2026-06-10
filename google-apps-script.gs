/**
 * TUC Investment — Google Apps Script
 *
 * HOW TO DEPLOY
 * ─────────────
 * 1. Open your Google Sheet (you can leave it empty — tabs are auto-created).
 * 2. Click Extensions → Apps Script.
 * 3. Delete any default code, then paste the entire contents of this file.
 * 4. Click Deploy → New deployment.
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy, approve permissions when prompted.
 * 6. Copy the Web app URL (https://script.google.com/macros/s/AK.../exec).
 * 7. In index.html, replace YOUR_GOOGLE_SCRIPT_URL_HERE with that URL.
 *
 * SHEET TABS
 * ──────────
 * Investors  — one row per investor inquiry
 * Contact    — one row per contact/general message
 * Partners   — one row per partnership inquiry
 *
 * Each tab is auto-created with styled headers on its first submission.
 */

// ─── Column definitions per form type ──────────────────────────────────────

const SCHEMAS = {
  investor: {
    tab: 'Investors',
    headers: [
      'Timestamp', 'Reference',
      'First Name', 'Last Name', 'Email', 'Phone', 'Location', 'Source',
      'Shares', 'Amount (USD)', 'Timeline', 'Experience', 'Message'
    ],
    row: function(d) {
      const sharePrice = 10000;
      const shares = parseInt(d.shares) || 1;
      return [
        new Date(), d.ref,
        d.firstName, d.lastName, d.email, d.phone, d.location, d.source,
        shares, shares * sharePrice,
        d.timeline, d.experience, d.investorMsg
      ];
    }
  },

  contact: {
    tab: 'Contact',
    headers: [
      'Timestamp', 'Reference',
      'First Name', 'Last Name', 'Email', 'Phone', 'Location', 'Source',
      'Topic', 'Preferred Contact', 'Message'
    ],
    row: function(d) {
      return [
        new Date(), d.ref,
        d.firstName, d.lastName, d.email, d.phone, d.location, d.source,
        d.contactTopic, d.contactMethod, d.contactMsg
      ];
    }
  },

  partner: {
    tab: 'Partners',
    headers: [
      'Timestamp', 'Reference',
      'First Name', 'Last Name', 'Email', 'Phone', 'Location', 'Source',
      'Organization', 'Partnership Type', 'Message'
    ],
    row: function(d) {
      return [
        new Date(), d.ref,
        d.firstName, d.lastName, d.email, d.phone, d.location, d.source,
        d.orgName, d.partnerType, d.partnerMsg
      ];
    }
  }
};

// ─── Main handler ───────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const d = e.parameter;
    const schema = SCHEMAS[d.formType];

    if (!schema) {
      return respond({ status: 'error', message: 'Unknown form type: ' + d.formType });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(schema.tab);

    if (!sheet) {
      sheet = ss.insertSheet(schema.tab);
      const headerRange = sheet.getRange(1, 1, 1, schema.headers.length);
      sheet.appendRow(schema.headers);
      sheet.setFrozenRows(1);
      headerRange
        .setBackground('#0f1e35')
        .setFontColor('#c4962a')
        .setFontWeight('bold');
    }

    sheet.appendRow(schema.row(d));

    return respond({ status: 'ok', ref: d.ref });

  } catch (err) {
    return respond({ status: 'error', message: err.message });
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Manual test — run in Apps Script editor to verify each tab ─────────────

function testInvestor() {
  doPost({ parameter: {
    ref: 'TUC-INV001', formType: 'investor',
    firstName: 'Ahmad', lastName: 'Hassan',
    email: 'ahmad@example.com', phone: '+1 860 000 0000',
    location: 'Hartford, CT', source: 'Word of Mouth',
    shares: '2', timeline: 'Within 1–3 months',
    experience: 'Some experience', investorMsg: 'Test investor submission.'
  }});
}

function testContact() {
  doPost({ parameter: {
    ref: 'TUC-CON001', formType: 'contact',
    firstName: 'Sara', lastName: 'Ali',
    email: 'sara@example.com', phone: '', location: 'Boston, MA', source: 'Social Media',
    contactTopic: 'Shariah Compliance', contactMethod: 'Email',
    contactMsg: 'Test contact submission.'
  }});
}

function testPartner() {
  doPost({ parameter: {
    ref: 'TUC-PAR001', formType: 'partner',
    firstName: 'Yusuf', lastName: 'Ibrahim',
    email: 'yusuf@masjid.org', phone: '+1 617 000 0000',
    location: 'Cambridge, MA', source: 'Referral',
    orgName: 'Masjid Al-Noor', partnerType: 'Islamic Institution / Masjid',
    partnerMsg: 'Test partner submission.'
  }});
}
