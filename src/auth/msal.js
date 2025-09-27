/**
 * Bouwt een ICS-bestand met hele-dag events voor de opgegeven data.
 * dates: array van 'YYYY-MM-DD'
 * options: { summary, description, location }
 */
export function buildIcsFromDates(dates, options = {}) {
  const { summary = 'Reservering', description = '', location = '' } = options;
  const fmtDate = (s) => s.replaceAll('-', ''); // YYYYMMDD
  const nextDay = (s) => {
    const [y, m, d] = s.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    dt.setUTCDate(dt.getUTCDate() + 1);
    const y2 = dt.getUTCFullYear();
    const m2 = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const d2 = String(dt.getUTCDate()).padStart(2, '0');
    return `${y2}-${m2}-${d2}`;
  };
  const dtStamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z');

  const events = dates.map((iso) => {
    const start = fmtDate(iso);
    const end = fmtDate(nextDay(iso));
    const uid = `poc-${start}-${Math.random().toString(36).slice(2)}@agenda-web-app`;
    const lines = [
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${dtStamp}`,
      `DTSTART;VALUE=DATE:${start}`,
      `DTEND;VALUE=DATE:${end}`,
      `SUMMARY:${escapeIcs(summary)}`,
      description ? `DESCRIPTION:${escapeIcs(description)}` : null,
      location ? `LOCATION:${escapeIcs(location)}` : null,
      'END:VEVENT'
    ].filter(Boolean);
    return lines.join('\r\n');
  });

  const vcal = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Agenda WebApp POC//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    ...events,
    'END:VCALENDAR'
  ].join('\r\n');

  return vcal;
}

function escapeIcs(text) {
  return String(text)
    .replaceAll('\\', '\\\\')
    .replaceAll('\n', '\\n')
    .replaceAll(',', '\\,')
    .replaceAll(';', '\\;');
}

/** Start een download van een gegeven ICS-string. */
export function downloadIcs(icsContent, filename = 'reserveringen.ics') {
  const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/**
 * Opent Outlook Web compose (gebruiker bevestigt handmatig).
 * base: 'live' voor outlook.live.com (consumers) of 'office' voor outlook.office.com (M365).
 * date: 'YYYY-MM-DD', hele-dag event (end = +1 dag).
 */
export function openOutlookDeeplink(date, { subject = 'Reservering', body = '', base = 'office' } = {}) {
  const addDays = (iso, n) => {
    const [y, m, d] = iso.split('-').map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    dt.setUTCDate(dt.getUTCDate() + n);
    const y2 = dt.getUTCFullYear();
    const m2 = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const d2 = String(dt.getUTCDate()).padStart(2, '0');
    return `${y2}-${m2}-${d2}`;
  };
  const start = date;
  const end = addDays(date, 1);
  const host = base === 'live' ? 'outlook.live.com' : 'outlook.office.com';
  const params = new URLSearchParams({
    path: '/calendar/action/compose',
    rru: 'addevent',
    startdt: start,
    enddt: end,
    allday: 'true',
    subject,
    body
  });
  const url = `https://${host}/calendar/0/deeplink/compose?${params.toString()}`;
  window.open(url, '_blank', 'noopener');
}
