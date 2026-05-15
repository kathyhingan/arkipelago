// ─────────────────────────────────────────────────────────────────
// Arkipelago Lead Notification Script
// ─────────────────────────────────────────────────────────────────

const NOTIFY_EMAIL = 'kathyhingan@gmail.com';
const CALENDAR_ID  = 'primary';

// ─── FORM HANDLER ────────────────────────────────────────────────
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data  = (e && e.parameter) || {};

    // Auto-create headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp','Name','Email','Journey Type','Dates',
        'Group Size','Message','Quiz — Who','Quiz — Draw',
        'Quiz — Pace','Quiz — Destination','Quiz — Luxury',
        'Quiz — Archetype','Source'
      ]);
    }

    sheet.appendRow([
      data.timestamp    || new Date().toISOString(),
      data.name         || '',
      data.email        || '',
      data.journey_type || '',
      data.dates        || '',
      data.group_size   || '',
      data.message      || '',
      data.quiz_who     || '',
      data.quiz_draw    || '',
      data.quiz_pace    || '',
      data.quiz_dest    || '',
      data.quiz_luxury  || '',
      data.quiz_archetype || '',
      data.source       || ''
    ]);

    // Fire notifications — wrapped separately so a notification failure
    // never breaks the lead save
    try {
      notifyNewLead(data);
    } catch (notifyErr) {
      console.warn('Notification failed (lead was saved):', notifyErr.toString());
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── NOTIFICATION HANDLER ────────────────────────────────────────
function notifyNewLead(data) {
  const name        = data.name           || 'Unknown';
  const email       = data.email          || 'Not provided';
  const journey     = data.journey_type   || 'Not specified';
  const dates       = data.dates          || 'Not specified';
  const groupSize   = data.group_size     || 'Not specified';
  const message     = data.message        || '';
  const archetype   = data.quiz_archetype || 'Not completed';
  const destination = data.quiz_dest      || 'Not specified';
  const pace        = data.quiz_pace      || 'Not specified';
  const luxury      = data.quiz_luxury    || 'Not specified';
  const now         = new Date();
  const phTime      = Utilities.formatDate(now, 'Asia/Manila', "yyyy-MM-dd HH:mm 'PHT'");

  // ── Email ────────────────────────────────────────────────────
  const subject = `🧭 New Lead — ${name} | ${journey} | ${destination}`;

  const body = `New Arkipelago inquiry received — ${phTime} (Manila time)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LEAD DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Name:         ${name}
Email:        ${email}
Journey Type: ${journey}
Travel Dates: ${dates}
Group Size:   ${groupSize}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
QUIZ ANSWERS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Archetype:    ${archetype}
Destination:  ${destination}
Pace:         ${pace}
Luxury means: ${luxury}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MESSAGE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${message || '(none)'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Respond within 24 hours. A calendar block has been created.`;

  MailApp.sendEmail({
    to:      NOTIFY_EMAIL,
    subject: subject,
    body:    body
  });

  // ── Calendar event ───────────────────────────────────────────
  // Creates a 45-min "Review & prepare proposal" block 1 hour from now
  const eventStart = new Date(now.getTime() + 60  * 60 * 1000);
  const eventEnd   = new Date(now.getTime() + 105 * 60 * 1000);

  const eventTitle = `📋 Arkipelago Lead: ${name} — ${journey}`;
  const eventDesc  = `Lead received: ${phTime}

Name:         ${name}
Email:        ${email}
Journey Type: ${journey}
Travel Dates: ${dates}
Group Size:   ${groupSize}
Archetype:    ${archetype}
Destination:  ${destination}
Pace:         ${pace}

Message: ${message || '(none)'}

ACTION ITEMS
→ Reply to ${email} within 24 hours
→ Draft initial proposal using quiz answers above
→ Check seasonal fit for ${destination}`;

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID) || CalendarApp.getDefaultCalendar();
  const event = calendar.createEvent(eventTitle, eventStart, eventEnd, {
    description: eventDesc
  });

  // Reminders must be set after creation; email reminder kept under
  // the event lead time so Calendar doesn't reject it.
  event.removeAllReminders();
  event.addPopupReminder(10);
  event.addEmailReminder(30);
}

// ─── HEALTH CHECK ────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'live', message: 'Arkipelago endpoint running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── TEST FUNCTION (run once to approve permissions) ─────────────
function testNotification() {
  notifyNewLead({
    name:           'Test Lead',
    email:          'test@example.com',
    journey_type:   'Honeymoon',
    dates:          'March 2026',
    group_size:     '2',
    message:        'This is a test submission.',
    quiz_archetype: 'The Romantic',
    quiz_dest:      'Coron',
    quiz_pace:      'Unhurried',
    quiz_luxury:    'Complete privacy'
  });
}
