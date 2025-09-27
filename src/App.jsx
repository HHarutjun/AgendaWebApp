import { useState, useMemo } from 'react';
import { buildIcsFromDates, downloadIcs } from './auth/msal';

export default function App() {
  const today = new Date();
  const [view, setView] = useState({ year: today.getFullYear(), month: today.getMonth() }); // month: 0-11
  const [selected, setSelected] = useState(new Set()); // keys: YYYY-MM-DD
  // Nieuw: event-details
  const [summary, setSummary] = useState('Reservering');
  const [location, setLocation] = useState('');
  const [description, setDescription] = useState('');

  const MONTHS = ['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'];
  const WEEKDAYS = ['Ma','Di','Wo','Do','Vr','Za','Zo'];

  const { cells, daysInMonth } = useMemo(() => {
    const startOfMonth = new Date(view.year, view.month, 1);
    const endOfMonth = new Date(view.year, view.month + 1, 0);
    const dim = endOfMonth.getDate();
    const mondayFirst = (startOfMonth.getDay() + 6) % 7; // 0=Ma ... 6=Zo
    const c = [];
    for (let i = 0; i < mondayFirst; i++) c.push(null);
    for (let d = 1; d <= dim; d++) c.push(d);
    return { cells: c, daysInMonth: dim };
  }, [view]);

  const keyFor = (day) =>
    `${view.year}-${String(view.month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

  const isSelected = (day) => selected.has(keyFor(day));

  const toggleDay = (day) => {
    if (!day || day < 1 || day > daysInMonth) return;
    const k = keyFor(day);
    setSelected((prev) => {
      const next = new Set(prev);
      next.has(k) ? next.delete(k) : next.add(k);
      return next;
    });
  };

  const prevMonth = () => {
    setView(({ year, month }) =>
      month === 0 ? { year: year - 1, month: 11 } : { year, month: month - 1 }
    );
  };
  const nextMonth = () => {
    setView(({ year, month }) =>
      month === 11 ? { year: year + 1, month: 0 } : { year, month: month + 1 }
    );
  };

  const handleReserve = async () => {
    const list = Array.from(selected).sort();
    if (list.length === 0) {
      alert('Geen dagen geselecteerd.');
      return;
    }
    const ics = buildIcsFromDates(list, {
      summary: summary?.trim() || 'Reservering',
      description: description?.trim(),
      location: location?.trim()
    });
    downloadIcs(ics, 'reserveringen.ics');

    // Alternatief (handmatige bevestiging in Outlook Web):
    // for (const d of list) openOutlookDeeplink(d);
  };

  // Nieuw: klaar-om-te-reserveren conditie
  const isReady = selected.size > 0 && (summary?.trim().length > 0);

  return (
    <div style={{ fontFamily: 'system-ui, sans-serif', padding: 24 }}>
      <h1>Agenda WebApp</h1>
      {/* Kalender */}
      <div className="calendar">
        <div className="month-nav">
          <button className="nav-btn" onClick={prevMonth} aria-label="Vorige maand">‹</button>
          <div className="title">{MONTHS[view.month]} {view.year}</div>
          <button className="nav-btn" onClick={nextMonth} aria-label="Volgende maand">›</button>
        </div>

        <div className="grid">
          {WEEKDAYS.map((w) => (
            <div key={w} className="weekday">{w}</div>
          ))}

          {cells.map((d, i) => {
            if (d === null) return <div key={`pad-${i}`} className="cell pad" />;
            const selectedDay = isSelected(d);
            return (
              <button
                key={`d-${d}-${i}`}
                className={`cell day${selectedDay ? ' selected' : ''}`}
                onClick={() => toggleDay(d)}
                aria-pressed={selectedDay}
                aria-label={`Dag ${d}`}
              >
                {d}
              </button>
            );
          })}
        </div>
      </div>

      {/* Nieuw: event details invoer */}
      <div className="details">
        <label>
          Titel
          <input
            type="text"
            value={summary}
            onChange={(e) => setSummary(e.target.value)}
            placeholder="Reservering"
          />
        </label>
        <label>
          Locatie
          <input
            type="text"
            value={location}
            onChange={(e) => setLocation(e.target.value)}
            placeholder="Bijv. kantoor"
          />
        </label>
        <label>
          Beschrijving
          <textarea
            rows={3}
            value={description}
            onChange={(e) => setDescription(e.target.value)}
            placeholder="Optioneel: extra info..."
          />
        </label>
      </div>

      <div style={{ marginTop: 16, display: 'flex', gap: 8, alignItems: 'center' }}>
        <button className="primary" onClick={handleReserve} disabled={!isReady}>Reserveren</button>
        <span style={{ opacity: 0.8 }}>
          Geselecteerd: {Array.from(selected).sort().slice(0, 3).join(', ') || '—'}
          {selected.size > 3 ? ` (+${selected.size - 3})` : ''}
        </span>
      </div>
    </div>
  );
}