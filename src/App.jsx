import { useState } from 'react';

export default function App() {
  const [selectedDays] = useState([]); // placeholder; later vervangen door echte dag-selectie

  const handleReserve = () => {
    alert('Reserveren (demo) — later koppelen we Microsoft Graph.');
  };

  return (
    <div style={{ fontFamily: 'system-ui, sans-serif', padding: 24 }}>
      <h1>Agenda WebApp</h1>
      <p>Hallo wereld 👋</p>
      <button onClick={handleReserve}>Reserveren</button>
    </div>
  );
}
