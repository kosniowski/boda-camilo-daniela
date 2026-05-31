import { useState, useEffect } from "react";
import {
  PieChart, Pie, Cell, BarChart, Bar,
  XAxis, YAxis, Tooltip, ResponsiveContainer,
  CartesianGrid, Legend
} from "recharts";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════
//  🌿 CONFIGURACIÓN — Actualiza estos valores
// ═══════════════════════════════════════════════════════
const SCRIPT_URL = ""; // Pega aquí la URL de tu Google Apps Script
const SHEETS_URL = ""; // Pega aquí la URL de tu Google Sheets
const ADMIN_PASSWORD = "CamiloDaniela2025";

const WEDDING = {
  groom: "Camilo",
  bride: "Daniela",
  date: "Sábado, 22 de Agosto de 2026",
  time: "Por confirmar",
  venue: "Por confirmar",
  address: "Cali, Colombia",
  dressCode: "Formal · Verde olivo y tonos tierra",
};

// ═══════════════════════════════════════════════════════
//  PALETA
// ═══════════════════════════════════════════════════════
const C = {
  olive:     "#5C6B2E",
  oliveMid:  "#7A8E3E",
  oliveLight:"#A3B86C",
  olivePale: "#E8EDD8",
  oliveFog:  "#D5DFBA",
  cream:     "#F8F4EE",
  gold:      "#B8933A",
  goldLight: "#D4AE5C",
  brown:     "#2A1A08",
  text:      "#352B1A",
  muted:     "#8A7D6A",
  error:     "#B93A3A",
  success:   "#4A7C4E",
};

// ═══════════════════════════════════════════════════════
//  DEMO DATA — reemplazada por Google Sheets en producción
// ═══════════════════════════════════════════════════════
const DEMO_GROUPS = [
  // Mesa #1
  { id: "1",  name: "Oswaldo",        maxGuests: 4 },
  { id: "2",  name: "Manuel",         maxGuests: 3 },
  { id: "3",  name: "Fernanda",       maxGuests: 1 },
  { id: "4",  name: "Graciela",       maxGuests: 2 },
  // Mesa #2
  { id: "5",  name: "Martha",         maxGuests: 2 },
  { id: "6",  name: "Lorena",         maxGuests: 2 },
  { id: "7",  name: "Ángela",         maxGuests: 2 },
  { id: "8",  name: "Genier",         maxGuests: 2 },
  // Mesa #3
  { id: "9",  name: "Charo",          maxGuests: 5 },
  { id: "10", name: "María Eugenia",  maxGuests: 3 },
  // Mesa #4
  { id: "11", name: "Daniel Urrea",   maxGuests: 2 },
  { id: "12", name: "Ana Gabriel",    maxGuests: 3 },
  { id: "13", name: "Catalina Duque", maxGuests: 2 },
  { id: "14", name: "José Luis",      maxGuests: 1 },
  { id: "15", name: "Juliana",        maxGuests: 1 },
  { id: "16", name: "Hugo",           maxGuests: 1 },
  { id: "17", name: "Yorleiner",      maxGuests: 1 },
  // Mesa #5
  { id: "18", name: "Diana",          maxGuests: 3 },
  { id: "19", name: "Daniela",        maxGuests: 3 },
  { id: "20", name: "Ledy",           maxGuests: 1 },
  { id: "21", name: "Paola",          maxGuests: 1 },
  { id: "22", name: "Yuliana",        maxGuests: 1 },
  { id: "23", name: "Andres",         maxGuests: 1 },
  { id: "24", name: "Olga",           maxGuests: 3 },
  // Mesa #6
  { id: "25", name: "Ruth",           maxGuests: 3 },
  { id: "26", name: "Josecito",       maxGuests: 2 },
  { id: "27", name: "Melchor",        maxGuests: 3 },
  // Mesa #7
  { id: "28", name: "Maricel",        maxGuests: 5 },
  { id: "29", name: "Eduardo",        maxGuests: 2 },
];

// ═══════════════════════════════════════════════════════
//  API — con fallback a localStorage (modo demo)
// ═══════════════════════════════════════════════════════
const api = {
  async getGroups() {
    if (SCRIPT_URL) {
      const r = await fetch(`${SCRIPT_URL}?action=getGroups`);
      return (await r.json()).groups || [];
    }
    const s = localStorage.getItem("wg_groups");
    return s ? JSON.parse(s) : DEMO_GROUPS;
  },
  async checkConfirmed(familyId) {
    if (SCRIPT_URL) {
      const r = await fetch(`${SCRIPT_URL}?action=checkConfirmed&familyId=${familyId}`);
      return (await r.json()).confirmed;
    }
    const c = JSON.parse(localStorage.getItem("wg_confirmations") || "[]");
    return c.some(x => x.familyId === familyId);
  },
  async saveConfirmation(data) {
    if (SCRIPT_URL) {
      const r = await fetch(SCRIPT_URL, {
        method: "POST",
        body: JSON.stringify({ action: "confirm", ...data }),
        headers: { "Content-Type": "application/json" },
      });
      return await r.json();
    }
    const c = JSON.parse(localStorage.getItem("wg_confirmations") || "[]");
    c.push({ ...data, timestamp: new Date().toISOString() });
    localStorage.setItem("wg_confirmations", JSON.stringify(c));
    return { success: true };
  },
  async getConfirmations() {
    if (SCRIPT_URL) {
      const r = await fetch(`${SCRIPT_URL}?action=getConfirmations`);
      return (await r.json()).confirmations || [];
    }
    return JSON.parse(localStorage.getItem("wg_confirmations") || "[]");
  },
  async uploadGroups(groups) {
    if (SCRIPT_URL) {
      const r = await fetch(SCRIPT_URL, {
        method: "POST",
        body: JSON.stringify({ action: "uploadGroups", groups }),
        headers: { "Content-Type": "application/json" },
      });
      return await r.json();
    }
    localStorage.setItem("wg_groups", JSON.stringify(groups));
    return { success: true };
  },
};

// ═══════════════════════════════════════════════════════
//  WHATSAPP HELPER
// ═══════════════════════════════════════════════════════
const buildWAUrl = (phone, familyName, guests) => {
  // Normalizar número para wa.me (necesita código de país completo, sin +)
  let num = phone.replace(/\D/g, "");
  if (num.startsWith("0057"))      num = num.slice(4);       // 005730... → 30...
  else if (num.startsWith("57") && num.length === 12) { /* ya correcto */ }
  else if (num.startsWith("3") && num.length === 10) num = "57" + num; // 300... → 57300...
  // Otros formatos: se usa tal cual
  const msg =
    `🌿 *¡Asistencia Confirmada!* 🌿\n\n` +
    `✅ *${familyName}* confirmó su asistencia a la boda de\n` +
    `*${WEDDING.bride} & ${WEDDING.groom}*\n\n` +
    `👥 Personas: *${guests}*\n` +
    `📅 Fecha: ${WEDDING.date}\n` +
    `⏰ Hora: ${WEDDING.time}\n` +
    `📍 Lugar: ${WEDDING.venue}\n` +
    `🗺️ ${WEDDING.address}\n` +
    `👗 Vestimenta: ${WEDDING.dressCode}\n\n` +
    `¡Los esperamos con mucho amor! 💚🌿`;
  return `https://wa.me/${num}?text=${encodeURIComponent(msg)}`;
};

// ═══════════════════════════════════════════════════════
//  SVG: RAMA DE EUCALIPTO
// ═══════════════════════════════════════════════════════
const EucalyptusBranch = ({ w = 90, h = 340, flip = false }) => {
  const leaves = [
    { cx: 30, cy: 310, rx: 17, ry: 7, a: -22 },
    { cx: 58, cy: 278, rx: 15, ry: 6, a: 24  },
    { cx: 27, cy: 246, rx: 16, ry: 6, a: -26 },
    { cx: 60, cy: 213, rx: 14, ry: 6, a: 21  },
    { cx: 29, cy: 181, rx: 14, ry: 5, a: -19 },
    { cx: 58, cy: 150, rx: 13, ry: 5, a: 23  },
    { cx: 31, cy: 120, rx: 12, ry: 5, a: -20 },
    { cx: 56, cy: 93,  rx: 11, ry: 4, a: 17  },
    { cx: 33, cy: 68,  rx: 10, ry: 4, a: -16 },
    { cx: 53, cy: 44,  rx: 9,  ry: 3, a: 14  },
    { cx: 37, cy: 22,  rx: 7,  ry: 3, a: -10 },
  ];
  const palette = [C.olive, C.oliveMid, C.oliveLight, "#6B7C44", "#8FA44E"];
  return (
    <svg viewBox="0 0 90 340" style={{ width: w, height: h, transform: flip ? "scaleX(-1)" : "none" }} fill="none">
      <path d="M45 340 C43 280 41 220 43 160 C45 100 47 55 45 0"
        stroke={C.olive} strokeWidth="1.8" fill="none" opacity="0.45" />
      {leaves.map(({ cx, cy, rx, ry, a }, i) => (
        <ellipse key={i} cx={cx} cy={cy} rx={rx} ry={ry}
          fill={palette[i % palette.length]}
          opacity={0.52 + i * 0.025}
          transform={`rotate(${a} ${cx} ${cy})`}
        />
      ))}
    </svg>
  );
};

// Small scattered leaves for background
const LeafScatter = () => (
  <svg viewBox="0 0 40 40" style={{ width: 36, height: 36 }} fill="none">
    <ellipse cx="20" cy="20" rx="16" ry="8" fill={C.olive} opacity="0.18"
      transform="rotate(-30 20 20)" />
  </svg>
);

// ═══════════════════════════════════════════════════════
//  GLOBAL STYLES
// ═══════════════════════════════════════════════════════
const GlobalStyles = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,500;0,700;1,400;1,500&family=Cormorant+Garamond:ital,wght@0,300;0,400;0,500;1,300;1,400&family=Jost:wght@300;400;500;600&display=swap');

    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Cormorant Garamond', Georgia, serif;
      background: ${C.cream};
      color: ${C.text};
      min-height: 100vh;
    }

    @keyframes fadeUp {
      from { opacity: 0; transform: translateY(28px); }
      to   { opacity: 1; transform: translateY(0);    }
    }
    @keyframes floatA {
      0%,100% { transform: translateY(0) rotate(0deg);   }
      50%      { transform: translateY(-14px) rotate(2deg); }
    }
    @keyframes floatB {
      0%,100% { transform: translateY(0) rotate(0deg) scaleX(-1);   }
      50%      { transform: translateY(-10px) rotate(-2deg) scaleX(-1); }
    }
    @keyframes ringPulse {
      0%,100% { box-shadow: 0 6px 28px rgba(92,107,46,.45), 0 0 0 0 rgba(92,107,46,.35); }
      50%      { box-shadow: 0 6px 28px rgba(92,107,46,.45), 0 0 0 16px rgba(92,107,46,0);  }
    }
    @keyframes slideCard {
      from { opacity: 0; transform: translateY(24px) scale(.97); }
      to   { opacity: 1; transform: translateY(0)   scale(1);    }
    }
    @keyframes fadeIn {
      from { opacity: 0; } to { opacity: 1; }
    }
    @keyframes checkPop {
      0%   { transform: scale(0); opacity: 0; }
      70%  { transform: scale(1.2); }
      100% { transform: scale(1);   opacity: 1; }
    }
    @keyframes shimmer {
      0%   { background-position: -200% center; }
      100% { background-position: 200% center;  }
    }
    @keyframes gearSpin {
      from { transform: rotate(0deg); }
      to   { transform: rotate(360deg); }
    }

    .d1  { animation: fadeUp .8s ease .1s both; }
    .d2  { animation: fadeUp .8s ease .3s both; }
    .d3  { animation: fadeUp .8s ease .5s both; }
    .d4  { animation: fadeUp .8s ease .7s both; }
    .d5  { animation: fadeUp .8s ease .9s both; }
    .fa  { animation: floatA 6s ease-in-out infinite; }
    .fb  { animation: floatB 7s ease-in-out 1.2s infinite; }
    .rp  { animation: ringPulse 2.6s ease-in-out infinite; }
    .sc  { animation: slideCard .55s cubic-bezier(.22,1,.36,1) both; }
    .fi  { animation: fadeIn .4s ease both; }
    .cp  { animation: checkPop .55s cubic-bezier(.34,1.56,.64,1) .2s both; }

    /* ── Material Expressive buttons ── */
    .btn {
      display: inline-flex; align-items: center; justify-content: center;
      font-family: 'Jost', sans-serif; font-weight: 500; letter-spacing: 2px;
      text-transform: uppercase; font-size: 12px; cursor: pointer;
      border-radius: 50px; border: none; transition: all .3s cubic-bezier(.34,1.56,.64,1);
      padding: 15px 40px; position: relative; overflow: hidden;
    }
    .btn::after {
      content: ''; position: absolute; inset: 0;
      background: rgba(255,255,255,0); transition: background .2s;
    }
    .btn:hover::after { background: rgba(255,255,255,.12); }

    .btn-g {
      background: linear-gradient(135deg, ${C.olive} 0%, ${C.oliveMid} 60%, #8FA44E 100%);
      color: #fff;
      box-shadow: 0 4px 20px rgba(92,107,46,.35), 0 1px 4px rgba(92,107,46,.2);
    }
    .btn-g:hover:not(:disabled) {
      transform: translateY(-3px) scale(1.02);
      box-shadow: 0 10px 32px rgba(92,107,46,.45), 0 2px 8px rgba(92,107,46,.25);
    }
    .btn-g:active:not(:disabled) { transform: translateY(-1px) scale(.99); }
    .btn-g:disabled { background: #ccc; box-shadow: none; cursor: not-allowed; }

    .btn-o {
      background: transparent; color: ${C.olive};
      border: 2px solid ${C.oliveFog};
      padding: 13px 32px;
    }
    .btn-o:hover {
      background: ${C.olivePale};
      border-color: ${C.olive};
      transform: translateY(-2px);
      box-shadow: 0 6px 18px rgba(92,107,46,.15);
    }

    /* ── Inputs ── */
    .inp {
      width: 100%; padding: 14px 20px;
      border: 2px solid ${C.oliveFog}; border-radius: 16px;
      background: rgba(255,255,255,.85); font-size: 17px;
      font-family: 'Cormorant Garamond', serif; color: ${C.text};
      outline: none; transition: all .25s;
      box-shadow: 0 2px 8px rgba(92,107,46,.06);
    }
    .inp:focus {
      border-color: ${C.olive};
      background: white;
      box-shadow: 0 4px 16px rgba(92,107,46,.15);
      transform: translateY(-1px);
    }
    .inp::placeholder { color: ${C.muted}; font-style: italic; }

    /* ── Gold divider ── */
    .gline {
      height: 1.5px;
      background: linear-gradient(90deg, transparent, ${C.gold}, transparent);
      margin: 0 auto;
      border-radius: 2px;
    }

    /* ── Gear button ── */
    .gear-btn {
      width: 40px; height: 40px; border-radius: 50%;
      background: rgba(255,255,255,.7);
      border: 1.5px solid ${C.oliveFog};
      display: flex; align-items: center; justify-content: center;
      cursor: pointer; font-size: 18px;
      transition: all .3s cubic-bezier(.34,1.56,.64,1);
      backdrop-filter: blur(8px);
      box-shadow: 0 2px 12px rgba(92,107,46,.12);
    }
    .gear-btn:hover {
      background: ${C.olivePale};
      border-color: ${C.olive};
      transform: scale(1.12) rotate(30deg);
      box-shadow: 0 6px 20px rgba(92,107,46,.22);
    }

    ::-webkit-scrollbar { width: 5px; }
    ::-webkit-scrollbar-track { background: ${C.cream}; }
    ::-webkit-scrollbar-thumb { background: ${C.oliveLight}; border-radius: 99px; }
  `}</style>
);

// ═══════════════════════════════════════════════════════
//  LANDING PAGE
// ═══════════════════════════════════════════════════════
const LandingPage = ({ onRSVP, onAdmin }) => (
  <div style={{
    minHeight: "100vh",
    background: `
      radial-gradient(ellipse at 15% 85%, rgba(92,107,46,.18) 0%, transparent 50%),
      radial-gradient(ellipse at 85% 10%, rgba(184,147,58,.12) 0%, transparent 45%),
      radial-gradient(ellipse at 50% 50%, rgba(163,184,108,.08) 0%, transparent 70%),
      ${C.cream}`,
    display: "flex", flexDirection: "column",
    alignItems: "center", justifyContent: "center",
    padding: "40px 24px", position: "relative", overflow: "hidden",
  }}>
    {/* Scattered bg leaves */}
    {[
      { top: "8%",  left: "12%",  opacity: .12, rotate: 30 },
      { top: "15%", right: "10%", opacity: .10, rotate: -45 },
      { top: "75%", left: "8%",   opacity: .10, rotate: 60  },
      { top: "80%", right: "14%", opacity: .12, rotate: -20 },
      { top: "45%", left: "4%",   opacity: .08, rotate: 10  },
      { top: "50%", right: "5%",  opacity: .08, rotate: -10 },
    ].map((s, i) => (
      <div key={i} style={{ position: "absolute", ...s, transform: `rotate(${s.rotate}deg)` }}>
        <LeafScatter />
      </div>
    ))}

    {/* Left branch cluster */}
    <div className="fa" style={{ position: "absolute", left: -18, bottom: -10, opacity: .75 }}>
      <EucalyptusBranch w={110} h={400} />
    </div>
    <div className="fa" style={{ position: "absolute", left: 65, top: 10, opacity: .35, animationDelay: "1.5s" }}>
      <EucalyptusBranch w={65} h={240} />
    </div>

    {/* Right branch cluster */}
    <div className="fb" style={{ position: "absolute", right: -18, bottom: -10, opacity: .75 }}>
      <EucalyptusBranch w={110} h={400} flip />
    </div>
    <div className="fb" style={{ position: "absolute", right: 65, top: 20, opacity: .35, animationDelay: "2.5s" }}>
      <EucalyptusBranch w={65} h={240} flip />
    </div>

    {/* ── Main content ── */}
    <div style={{ textAlign: "center", position: "relative", zIndex: 1, maxWidth: 560 }}>

      {/* Tag */}
      <div className="d1">
        <span style={{
          fontFamily: "Jost, sans-serif", fontSize: 11,
          letterSpacing: 4, textTransform: "uppercase",
          background: `linear-gradient(135deg, ${C.olivePale}, #d8e6b0)`,
          color: C.olive, fontWeight: 600,
          padding: "7px 20px", borderRadius: 50,
          display: "inline-block",
          border: `1px solid rgba(92,107,46,.2)`,
          boxShadow: "0 2px 12px rgba(92,107,46,.12)",
        }}>✦ Save the Date ✦</span>
      </div>

      {/* Monogram / symbol */}
      <div className="d2" style={{ margin: "28px 0 0" }}>
        <p style={{
          fontFamily: "Jost, sans-serif", fontSize: 11,
          letterSpacing: 5, color: C.gold,
        }}>✦ ✦ ✦</p>
      </div>

      {/* Names */}
      <div className="d2" style={{ marginTop: 8 }}>
        <h1 style={{
          fontFamily: "Playfair Display, Georgia, serif",
          fontSize: "clamp(50px, 11vw, 84px)",
          fontWeight: 400, fontStyle: "italic",
          color: C.olive, lineHeight: 1.05,
          textShadow: `2px 4px 24px rgba(92,107,46,.12)`,
        }}>
          {WEDDING.bride}
        </h1>
      </div>

      <div className="d3" style={{ margin: "4px 0 6px" }}>
        <p style={{
          fontFamily: "Jost, sans-serif", fontSize: 14,
          letterSpacing: 5, color: C.gold,
        }}>&amp;</p>
      </div>

      <div className="d3">
        <h1 style={{
          fontFamily: "Playfair Display, Georgia, serif",
          fontSize: "clamp(50px, 11vw, 84px)",
          fontWeight: 400, fontStyle: "italic",
          color: C.olive, lineHeight: 1.05,
          textShadow: `2px 4px 24px rgba(92,107,46,.12)`,
        }}>
          {WEDDING.groom}
        </h1>
      </div>

      {/* Divider info */}
      <div className="d4" style={{ margin: "28px 0" }}>
        <div className="gline" style={{ width: 80 }} />
        <p style={{
          fontFamily: "Jost, sans-serif", fontSize: 11,
          letterSpacing: 4, textTransform: "uppercase",
          color: C.muted, marginTop: 20,
        }}>{WEDDING.date}</p>
        <p style={{ fontSize: 19, fontStyle: "italic", color: C.text, marginTop: 5 }}>
          {WEDDING.time} · {WEDDING.venue}
        </p>
        <p style={{
          fontFamily: "Jost, sans-serif", fontSize: 11,
          letterSpacing: 2, color: C.muted, marginTop: 4,
        }}>{WEDDING.address}</p>
        <div className="gline" style={{ width: 80, marginTop: 20 }} />
      </div>

      {/* CTA */}
      <div className="d5">
        <button className="btn btn-g rp" onClick={onRSVP}
          style={{ fontSize: 11, letterSpacing: 3, padding: "17px 52px" }}>
          Confirmar Asistencia
        </button>
      </div>
      <div className="d5" style={{ marginTop: 14 }}>
        <p style={{ fontStyle: "italic", color: C.muted, fontSize: 15 }}>
          Confirma tu asistencia antes del{" "}
          <span style={{ color: C.olive }}>1 de agosto de 2026</span>
        </p>
      </div>
    </div>

    {/* Engranaje admin — esquina superior derecha */}
    <button onClick={onAdmin} className="gear-btn"
      style={{ position: "absolute", top: 18, right: 18, zIndex: 10 }}
      title="Administradores"
    >⚙</button>
  </div>
);

// ═══════════════════════════════════════════════════════
//  RSVP PAGE  (3 steps)
// ═══════════════════════════════════════════════════════
const RSVPPage = ({ onBack }) => {
  const [step,        setStep]        = useState(1);
  const [selectedId,  setSelectedId]  = useState("");
  const [groups,      setGroups]      = useState([]);
  const [family,      setFamily]      = useState(null);
  const [guests,      setGuests]      = useState(1);
  const [phone,       setPhone]       = useState("");
  const [email,       setEmail]       = useState("");
  const [loading,     setLoading]     = useState(false);
  const [error,       setError]       = useState("");
  const [alreadyDone, setAlreadyDone] = useState(false);

  useEffect(() => { api.getGroups().then(setGroups); }, []);

  const handleSearch = async () => {
    setError("");
    if (!selectedId) { setError("Selecciona tu grupo familiar"); return; }
    setLoading(true);
    try {
      const match = groups.find(g => g.id === selectedId);
      if (!match) { setError("Grupo no encontrado."); setLoading(false); return; }
      const done = await api.checkConfirmed(match.id);
      setFamily(match);
      setGuests(match.maxGuests);
      setAlreadyDone(done);
      setStep(2);
    } catch { setError("Error de conexión. Intenta de nuevo."); }
    setLoading(false);
  };

  const handleConfirm = async () => {
    setError("");
    if (!phone && !email) { setError("Ingresa un celular o correo electrónico"); return; }
    if (phone && phone.replace(/\D/g, "").length < 7) { setError("Celular inválido"); return; }
    setLoading(true);
    try {
      await api.saveConfirmation({ familyId: family.id, familyName: family.name, guestCount: guests, phone, email });
      setStep(3);
    } catch { setError("Error al guardar. Intenta de nuevo."); }
    setLoading(false);
  };

  const cardStyle = {
    background: "rgba(255,255,255,.95)",
    borderRadius: 28,
    boxShadow: `0 20px 60px rgba(92,107,46,.14), 0 4px 16px rgba(92,107,46,.08), 0 1px 3px rgba(0,0,0,.06)`,
    padding: "40px 36px",
    width: "100%",
    maxWidth: 480,
    position: "relative",
    zIndex: 1,
    border: `1px solid rgba(213,223,186,.5)`,
    backdropFilter: "blur(12px)",
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: `
        radial-gradient(ellipse at 15% 90%, rgba(92,107,46,.15), transparent 50%),
        radial-gradient(ellipse at 85% 5%, rgba(184,147,58,.10), transparent 45%),
        ${C.cream}`,
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: "24px 16px", position: "relative",
    }}>
      <div style={{ position: "absolute", left: 0, bottom: 0, opacity: .25 }}>
        <EucalyptusBranch w={80} h={300} />
      </div>
      <div style={{ position: "absolute", right: 0, top: 30, opacity: .25 }}>
        <EucalyptusBranch w={72} h={260} flip />
      </div>

      <div className="sc" style={cardStyle}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <span style={{
            fontFamily: "Jost, sans-serif", fontSize: 10, letterSpacing: 4,
            textTransform: "uppercase",
            background: `linear-gradient(135deg, ${C.olivePale}, #d8e6b0)`,
            color: C.olive, fontWeight: 600,
            padding: "5px 16px", borderRadius: 50, display: "inline-block",
            border: `1px solid rgba(92,107,46,.2)`,
            boxShadow: "0 2px 8px rgba(92,107,46,.1)",
          }}>Confirmar Asistencia</span>
          <p style={{
            fontFamily: "Playfair Display, serif", fontSize: 26,
            fontStyle: "italic", color: C.olive, marginTop: 10, fontWeight: 400,
          }}>{WEDDING.bride} &amp; {WEDDING.groom}</p>
          <div className="gline" style={{ width: 60, marginTop: 12 }} />
        </div>

        {/* Step bar */}
        {step < 3 && (
          <div style={{ display: "flex", justifyContent: "center", gap: 8, marginBottom: 28 }}>
            {[1, 2].map(s => (
              <div key={s} style={{
                width: 32, height: 3, borderRadius: 2,
                background: step >= s ? C.olive : C.oliveFog,
                transition: "background .4s",
              }} />
            ))}
          </div>
        )}

        {/* ─── STEP 1: dropdown ─── */}
        {step === 1 && (
          <div className="fi">
            <p style={{ fontSize: 21, textAlign: "center", fontStyle: "italic", marginBottom: 6 }}>
              ¿Cuál es tu nombre?
            </p>
            <p style={{
              fontFamily: "Jost, sans-serif", fontSize: 12, letterSpacing: .5,
              color: C.muted, textAlign: "center", marginBottom: 22,
            }}>
              Selecciona tu nombre de la lista de invitados
            </p>
            <select
              value={selectedId}
              onChange={e => { setSelectedId(e.target.value); setError(""); }}
              style={{
                width: "100%", padding: "14px 20px",
                fontFamily: "Cormorant Garamond, serif", fontSize: 17,
                color: selectedId ? C.text : C.muted,
                background: "rgba(255,255,255,.85)",
                border: `2px solid ${C.oliveFog}`,
                borderRadius: 16, outline: "none",
                cursor: "pointer", marginBottom: 12,
                appearance: "none",
                backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%235C6B2E' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E")`,
                backgroundRepeat: "no-repeat",
                backgroundPosition: "right 18px center",
                paddingRight: 44,
                boxShadow: "0 2px 8px rgba(92,107,46,.06)",
                transition: "all .25s",
              }}
            >
              <option value="">— Selecciona tu nombre —</option>
              {groups.map(g => (
                <option key={g.id} value={g.id}>
                  {g.name} ({g.maxGuests} {g.maxGuests === 1 ? "persona" : "personas"})
                </option>
              ))}
            </select>
            {error && <p style={{
              fontFamily: "Jost, sans-serif", fontSize: 13,
              color: C.error, marginBottom: 10,
            }}>{error}</p>}
            <button className="btn btn-g" onClick={handleSearch} disabled={loading || !selectedId}
              style={{ width: "100%", marginTop: 4, opacity: selectedId ? 1 : .55 }}>
              {loading ? "Verificando..." : "Continuar →"}
            </button>
            <button className="btn btn-o" onClick={onBack}
              style={{ width: "100%", marginTop: 10 }}>
              ← Volver
            </button>
          </div>
        )}

        {/* ─── STEP 2a: ya confirmó ─── */}
        {step === 2 && alreadyDone && (
          <div className="fi" style={{ textAlign: "center" }}>
            <div className="cp" style={{
              width: 68, height: 68, borderRadius: "50%",
              background: C.olivePale, display: "flex",
              alignItems: "center", justifyContent: "center",
              margin: "0 auto 20px", fontSize: 28, color: C.olive,
            }}>✓</div>
            <p style={{
              fontFamily: "Playfair Display, serif", fontSize: 26,
              color: C.olive, fontStyle: "italic",
            }}>¡Ya confirmaste!</p>
            <p style={{ marginTop: 12, fontStyle: "italic", color: C.muted, fontSize: 17 }}>
              <strong style={{ color: C.text }}>{family.name}</strong> ya tiene su
              asistencia registrada. ¡Muchas gracias!
            </p>
            <div style={{
              marginTop: 24, padding: "16px 20px",
              background: `linear-gradient(135deg, ${C.olivePale}, #e5efcc)`,
              borderRadius: 16,
              border: `1px solid rgba(92,107,46,.15)`,
              fontFamily: "Jost, sans-serif", fontSize: 13, letterSpacing: .5,
              lineHeight: 2,
            }}>
              <p>📅 {WEDDING.date}</p>
              <p>⏰ {WEDDING.time}</p>
              <p>📍 {WEDDING.venue} · {WEDDING.address}</p>
            </div>
            <button className="btn btn-g" onClick={onBack} style={{ width: "100%", marginTop: 24 }}>
              Volver al Inicio
            </button>
          </div>
        )}

        {/* ─── STEP 2b: formulario ─── */}
        {step === 2 && !alreadyDone && family && (
          <div className="fi">
            {/* Family banner */}
            <div style={{
              background: `linear-gradient(135deg, ${C.olivePale}, #e5efcc)`,
              borderRadius: 20,
              border: `1px solid rgba(92,107,46,.15)`,
              padding: "16px 20px", marginBottom: 24, textAlign: "center",
              boxShadow: "0 4px 16px rgba(92,107,46,.08)",
            }}>
              <p style={{ fontSize: 18, fontStyle: "italic" }}>
                ¡Hola, <strong>{family.name}</strong>!
              </p>
              <p style={{
                fontFamily: "Jost, sans-serif", fontSize: 12, color: C.muted,
                marginTop: 5, letterSpacing: .5,
              }}>
                Tu grupo puede asistir con hasta{" "}
                <strong style={{ color: C.olive }}>
                  {family.maxGuests} persona{family.maxGuests !== 1 ? "s" : ""}
                </strong>
              </p>
            </div>

            {/* Guest counter */}
            <div style={{ marginBottom: 20 }}>
              <p style={{
                fontFamily: "Jost, sans-serif", fontSize: 11,
                letterSpacing: 2, textTransform: "uppercase",
                color: C.muted, marginBottom: 10,
              }}>¿Cuántos asistirán?</p>
              <div style={{ display: "flex", alignItems: "stretch", height: 52, borderRadius: 16, overflow: "hidden", border: `2px solid ${C.oliveFog}`, boxShadow: "0 2px 8px rgba(92,107,46,.06)" }}>
                <button
                  onClick={() => setGuests(g => Math.max(1, g - 1))}
                  style={{
                    width: 52, background: C.olivePale,
                    border: "none", borderRight: `2px solid ${C.oliveFog}`,
                    fontSize: 22, cursor: "pointer", color: C.olive,
                    transition: "background .2s",
                  }}>−</button>
                <div style={{
                  flex: 1, display: "flex", alignItems: "center", justifyContent: "center",
                  fontFamily: "Playfair Display, serif", fontSize: 24,
                  background: "white",
                }}>{guests}</div>
                <button
                  onClick={() => setGuests(g => Math.min(family.maxGuests, g + 1))}
                  style={{
                    width: 52, background: C.olivePale,
                    border: "none", borderLeft: `2px solid ${C.oliveFog}`,
                    fontSize: 22, cursor: "pointer", color: C.olive,
                    transition: "background .2s",
                  }}>+</button>
              </div>
              <p style={{ fontFamily: "Jost, sans-serif", fontSize: 11, color: C.muted, marginTop: 6 }}>
                Máximo {family.maxGuests} persona{family.maxGuests !== 1 ? "s" : ""}
              </p>
            </div>

            {/* Phone */}
            <div style={{ marginBottom: 14 }}>
              <p style={{
                fontFamily: "Jost, sans-serif", fontSize: 11, letterSpacing: 2,
                textTransform: "uppercase", color: C.muted, marginBottom: 8,
              }}>Celular (para recibir confirmación por WhatsApp)</p>
              <input className="inp" value={phone}
                onChange={e => setPhone(e.target.value)}
                placeholder="Ej: 300 123 4567 o +57 300 123 4567" type="tel"
              />
            </div>

            {/* Email */}
            <div style={{ marginBottom: 8 }}>
              <p style={{
                fontFamily: "Jost, sans-serif", fontSize: 11, letterSpacing: 2,
                textTransform: "uppercase", color: C.muted, marginBottom: 8,
              }}>Correo electrónico (opcional)</p>
              <input className="inp" value={email}
                onChange={e => setEmail(e.target.value)}
                placeholder="tu@correo.com" type="email"
              />
            </div>

            {error && <p style={{
              fontFamily: "Jost, sans-serif", fontSize: 13,
              color: C.error, margin: "8px 0",
            }}>{error}</p>}

            <button className="btn btn-g" onClick={handleConfirm} disabled={loading}
              style={{ width: "100%", marginTop: 18 }}>
              {loading ? "Confirmando..." : "✓  Confirmar Asistencia"}
            </button>
            <button className="btn btn-o"
              onClick={() => { setStep(1); setFamily(null); setSelectedId(""); setError(""); }}
              style={{ width: "100%", marginTop: 10 }}>
              ← Volver
            </button>
          </div>
        )}

        {/* ─── STEP 3: éxito ─── */}
        {step === 3 && (
          <div className="fi" style={{ textAlign: "center" }}>
            <div className="cp" style={{
              width: 76, height: 76, borderRadius: "50%",
              background: `linear-gradient(135deg, ${C.olive}, ${C.oliveMid})`,
              display: "flex", alignItems: "center", justifyContent: "center",
              margin: "0 auto 18px", fontSize: 34,
            }}>🌿</div>
            <p style={{
              fontFamily: "Playfair Display, serif",
              fontSize: 32, fontStyle: "italic", color: C.olive, fontWeight: 400,
            }}>¡Confirmado!</p>
            <p style={{ fontStyle: "italic", color: C.muted, fontSize: 17, marginTop: 8 }}>
              ¡Los esperamos con mucho amor!
            </p>

            {/* Detail card */}
            <div style={{
              margin: "22px 0",
              padding: "20px 22px",
              background: `linear-gradient(135deg, ${C.olivePale}, #e5efcc)`,
              borderRadius: 20,
              border: `1px solid rgba(92,107,46,.15)`,
              boxShadow: "0 4px 16px rgba(92,107,46,.08)",
              textAlign: "left",
            }}>
              <p style={{
                fontFamily: "Jost, sans-serif", fontSize: 10, letterSpacing: 3,
                textTransform: "uppercase", color: C.olive, marginBottom: 14, fontWeight: 600,
              }}>Detalles de la boda</p>
              {[
                ["👰🤵", `${WEDDING.bride} & ${WEDDING.groom}`],
                ["📅",   WEDDING.date],
                ["⏰",   WEDDING.time],
                ["📍",   `${WEDDING.venue} · ${WEDDING.address}`],
                ["👗",   WEDDING.dressCode],
                ["👥",   `${guests} persona${guests !== 1 ? "s" : ""} confirmada${guests !== 1 ? "s" : ""}`],
              ].map(([icon, txt]) => (
                <div key={txt} style={{ display: "flex", gap: 10, marginBottom: 9, alignItems: "flex-start" }}>
                  <span style={{ fontSize: 14, minWidth: 22 }}>{icon}</span>
                  <p style={{ fontSize: 15, lineHeight: 1.4 }}>{txt}</p>
                </div>
              ))}
            </div>

            {/* WhatsApp */}
            {phone && (
              <a href={buildWAUrl(phone, family?.name || "", guests)}
                target="_blank" rel="noopener noreferrer"
                style={{
                  display: "flex", alignItems: "center", justifyContent: "center",
                  gap: 10, background: "linear-gradient(135deg, #25D366, #1db954)", color: "white",
                  padding: "15px 28px", borderRadius: 50,
                  fontFamily: "Jost, sans-serif", fontSize: 11,
                  fontWeight: 500, letterSpacing: 2,
                  textTransform: "uppercase", textDecoration: "none",
                  marginBottom: 10,
                  boxShadow: "0 6px 20px rgba(37,211,102,.35)",
                  transition: "all .3s",
                }}>
                <span style={{ fontSize: 18 }}>📱</span>
                Recibir Tarjeta por WhatsApp
              </a>
            )}

            <button className="btn btn-o" onClick={onBack} style={{ width: "100%", marginTop: 4 }}>
              Volver al Inicio
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

// ═══════════════════════════════════════════════════════
//  ADMIN LOGIN
// ═══════════════════════════════════════════════════════
const AdminLogin = ({ onLogin, onBack }) => {
  const [pass,  setPass]  = useState("");
  const [error, setError] = useState("");

  const tryLogin = () => {
    if (pass === ADMIN_PASSWORD) onLogin();
    else setError("Contraseña incorrecta");
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: C.brown,
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: 20,
    }}>
      <div className="sc" style={{
        background: "white", borderRadius: 28,
        boxShadow: "0 24px 64px rgba(0,0,0,.45)",
        width: "100%", maxWidth: 380, padding: "44px 36px",
        textAlign: "center", border: `1px solid rgba(184,147,58,.2)`,
      }}>
        <p style={{
          fontFamily: "Jost, sans-serif", fontSize: 10,
          letterSpacing: 4, textTransform: "uppercase",
          background: `linear-gradient(135deg, ${C.olivePale}, #d8e6b0)`, color: C.olive,
          fontWeight: 600, padding: "5px 16px", borderRadius: 50,
          display: "inline-block", border: "1px solid rgba(92,107,46,.2)",
        }}>Acceso Privado</p>
        <p style={{
          fontFamily: "Playfair Display, serif",
          fontSize: 28, fontStyle: "italic", color: C.olive, marginTop: 14,
        }}>Administradores</p>
        <div className="gline" style={{ width: 60, marginTop: 12, marginBottom: 28 }} />

        <input className="inp" type="password" value={pass}
          onChange={e => setPass(e.target.value)}
          onKeyDown={e => e.key === "Enter" && tryLogin()}
          placeholder="Contraseña"
          style={{ textAlign: "center", letterSpacing: 6, marginBottom: 8 }}
        />
        {error && <p style={{
          fontFamily: "Jost, sans-serif", fontSize: 13,
          color: C.error, marginBottom: 8,
        }}>{error}</p>}

        <button className="btn btn-g" onClick={tryLogin}
          style={{ width: "100%", marginTop: 10 }}>
          Ingresar
        </button>
        <button className="btn btn-o" onClick={onBack}
          style={{ width: "100%", marginTop: 10 }}>
          Cancelar
        </button>
      </div>
    </div>
  );
};

// ═══════════════════════════════════════════════════════
//  ADMIN DASHBOARD
// ═══════════════════════════════════════════════════════
const AdminDashboard = ({ onLogout }) => {
  const [confirmations, setConfirmations] = useState([]);
  const [groups,        setGroups]        = useState([]);
  const [loading,       setLoading]       = useState(true);
  const [tab,           setTab]           = useState("overview");
  const [uploadMsg,     setUploadMsg]     = useState("");
  const [searchAdmin,   setSearchAdmin]   = useState("");

  const load = async () => {
    setLoading(true);
    const [c, g] = await Promise.all([api.getConfirmations(), api.getGroups()]);
    setConfirmations(c);
    setGroups(g);
    setLoading(false);
  };
  useEffect(() => { load(); }, []);

  const handleExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const parsed = rows.slice(1).filter(r => r[0] && r[1]).map((r, i) => ({
          id: String(Date.now() + i),
          name: String(r[0]).trim(),
          maxGuests: parseInt(r[1]) || 2,
        }));
        await api.uploadGroups(parsed);
        setGroups(parsed);
        setUploadMsg(`✓ ${parsed.length} grupos cargados exitosamente`);
      } catch { setUploadMsg("✗ Error al procesar el archivo"); }
    };
    reader.readAsArrayBuffer(file);
  };

  // Stats
  const totalGroups       = groups.length;
  const confirmed         = confirmations.length;
  const pending           = Math.max(0, totalGroups - confirmed);
  const totalGuests       = confirmations.reduce((s, c) => s + (parseInt(c.guestCount) || 0), 0);
  const maxPossible       = groups.reduce((s, g) => s + (parseInt(g.maxGuests) || 0), 0);

  const pieData   = [
    { name: "Confirmados", value: confirmed },
    { name: "Pendientes",  value: pending   },
  ];
  const barData = confirmations.map(c => ({
    name: c.familyName?.split(" ").slice(-1)[0] || c.familyName,
    Personas: parseInt(c.guestCount) || 0,
  }));

  const filteredConf = confirmations.filter(c =>
    c.familyName?.toLowerCase().includes(searchAdmin.toLowerCase()) ||
    c.phone?.includes(searchAdmin) ||
    c.email?.toLowerCase().includes(searchAdmin.toLowerCase())
  );

  const TH = ({ children }) => (
    <th style={{
      padding: "12px 16px", textAlign: "left",
      fontFamily: "Jost, sans-serif", fontSize: 10,
      letterSpacing: 2.5, textTransform: "uppercase",
      color: "white", fontWeight: 500, whiteSpace: "nowrap",
      background: C.olive,
    }}>{children}</th>
  );
  const TD = ({ children, center }) => (
    <td style={{
      padding: "12px 16px", borderBottom: `1px solid ${C.olivePale}`,
      fontSize: 15, color: C.text, textAlign: center ? "center" : "left",
    }}>{children}</td>
  );

  const Tab = ({ id, label }) => (
    <button onClick={() => setTab(id)} style={{
      padding: "12px 22px", border: "none", cursor: "pointer",
      fontFamily: "Jost, sans-serif", fontSize: 11, letterSpacing: 2.5,
      textTransform: "uppercase",
      background: tab === id ? "white" : "transparent",
      color: tab === id ? C.olive : C.muted,
      borderBottom: tab === id ? `3px solid ${C.gold}` : "3px solid transparent",
      transition: "all .2s",
    }}>{label}</button>
  );

  return (
    <div style={{ minHeight: "100vh", background: "#EEF0EA", fontFamily: "Jost, sans-serif" }}>
      {/* Header */}
      <div style={{
        background: C.brown, padding: "18px 24px",
        display: "flex", alignItems: "center", justifyContent: "space-between",
        flexWrap: "wrap", gap: 12,
      }}>
        <div>
          <p style={{
            fontFamily: "Playfair Display, serif", color: "white",
            fontSize: 19, fontStyle: "italic",
          }}>{WEDDING.bride} &amp; {WEDDING.groom}</p>
          <p style={{ fontFamily: "Jost, sans-serif", color: C.muted, fontSize: 11, letterSpacing: 2 }}>
            {WEDDING.date} · Panel de Administradores
          </p>
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {SHEETS_URL && (
            <a href={SHEETS_URL} target="_blank" rel="noopener noreferrer"
              style={{
                background: "#34A853", color: "white", padding: "8px 16px",
                borderRadius: 2, fontSize: 11, letterSpacing: 2,
                textDecoration: "none", textTransform: "uppercase",
                display: "inline-flex", alignItems: "center", gap: 6,
              }}>📊 Google Sheets</a>
          )}
          <button onClick={load} style={{
            background: C.oliveMid, color: "white", border: "none",
            padding: "8px 16px", borderRadius: 2, cursor: "pointer",
            fontSize: 11, letterSpacing: 2, textTransform: "uppercase",
          }}>↻ Actualizar</button>
          <button onClick={onLogout} style={{
            background: "transparent", color: "#888", border: "1px solid #444",
            padding: "8px 16px", borderRadius: 2, cursor: "pointer",
            fontSize: 11, letterSpacing: 2, textTransform: "uppercase",
          }}>Salir</button>
        </div>
      </div>

      {/* Tabs */}
      <div style={{
        background: "#F5F5EF", borderBottom: `1px solid ${C.oliveFog}`,
        display: "flex", overflowX: "auto",
      }}>
        <Tab id="overview"  label="Resumen"         />
        <Tab id="list"      label="Confirmaciones"  />
        <Tab id="upload"    label="Cargar Grupos"   />
      </div>

      <div style={{ padding: "24px 20px", maxWidth: 1080, margin: "0 auto" }}>
        {loading ? (
          <div style={{ textAlign: "center", padding: 80 }}>
            <p style={{
              fontFamily: "Cormorant Garamond, serif",
              fontStyle: "italic", fontSize: 20, color: C.muted,
            }}>Cargando datos...</p>
          </div>
        ) : (
          <>
            {/* ─── OVERVIEW ─── */}
            {tab === "overview" && (
              <div>
                {/* Stat cards */}
                <div style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit, minmax(152px, 1fr))",
                  gap: 16, marginBottom: 24,
                }}>
                  {[
                    { label: "Total Grupos",    val: totalGroups, icon: "👥", color: C.olive     },
                    { label: "Confirmados",      val: confirmed,   icon: "✅", color: C.success   },
                    { label: "Pendientes",       val: pending,     icon: "⏳", color: C.gold      },
                    { label: "Invitados Conf.", val: totalGuests, icon: "🌿", color: "#4A7CC4"   },
                  ].map(s => (
                    <div key={s.label} style={{
                      background: "white", borderRadius: 4,
                      padding: "18px 16px",
                      boxShadow: "0 2px 12px rgba(0,0,0,.07)",
                      borderLeft: `3px solid ${s.color}`,
                    }}>
                      <div style={{ fontSize: 20, marginBottom: 6 }}>{s.icon}</div>
                      <div style={{
                        fontFamily: "Playfair Display, serif",
                        fontSize: 40, color: s.color, lineHeight: 1,
                      }}>{s.val}</div>
                      <div style={{
                        fontSize: 10, letterSpacing: 2, textTransform: "uppercase",
                        color: C.muted, marginTop: 5,
                      }}>{s.label}</div>
                    </div>
                  ))}
                </div>

                {/* Progress */}
                <div style={{
                  background: "white", borderRadius: 4,
                  padding: "18px 20px",
                  boxShadow: "0 2px 12px rgba(0,0,0,.07)",
                  marginBottom: 20,
                }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                    <p style={{ fontSize: 11, letterSpacing: 2, textTransform: "uppercase", color: C.muted }}>
                      Progreso de confirmaciones
                    </p>
                    <p style={{ fontSize: 13, color: C.olive, fontWeight: 600 }}>
                      {totalGroups > 0 ? Math.round((confirmed / totalGroups) * 100) : 0}%
                    </p>
                  </div>
                  <div style={{ height: 10, background: C.olivePale, borderRadius: 5, overflow: "hidden" }}>
                    <div style={{
                      height: "100%",
                      width: `${totalGroups > 0 ? (confirmed / totalGroups) * 100 : 0}%`,
                      background: `linear-gradient(90deg, ${C.olive}, ${C.oliveMid})`,
                      borderRadius: 5, transition: "width 1s ease",
                    }} />
                  </div>
                  <p style={{ fontSize: 12, color: C.muted, marginTop: 8 }}>
                    {confirmed} de {totalGroups} grupos han confirmado ·{" "}
                    {totalGuests} de {maxPossible} invitados posibles
                  </p>
                </div>

                {/* Charts row */}
                <div style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
                  gap: 20,
                }}>
                  {/* Pie */}
                  <div style={{
                    background: "white", borderRadius: 4,
                    padding: "18px 16px",
                    boxShadow: "0 2px 12px rgba(0,0,0,.07)",
                  }}>
                    <p style={{ fontSize: 11, letterSpacing: 2, textTransform: "uppercase", color: C.muted, marginBottom: 12 }}>
                      Estado de grupos
                    </p>
                    <ResponsiveContainer width="100%" height={200}>
                      <PieChart>
                        <Pie data={pieData} cx="50%" cy="50%" outerRadius={72}
                          dataKey="value" label={({ name, percent }) =>
                            `${name} ${(percent * 100).toFixed(0)}%`
                          }
                          labelLine={false}
                        >
                          {pieData.map((_, i) => (
                            <Cell key={i} fill={[C.olive, C.oliveFog][i]} />
                          ))}
                        </Pie>
                        <Tooltip
                          contentStyle={{ fontFamily: "Jost, sans-serif", fontSize: 12 }}
                        />
                        <Legend
                          formatter={v => (
                            <span style={{ fontFamily: "Jost, sans-serif", fontSize: 11, letterSpacing: 1 }}>{v}</span>
                          )}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>

                  {/* Bar */}
                  {barData.length > 0 && (
                    <div style={{
                      background: "white", borderRadius: 4,
                      padding: "18px 16px",
                      boxShadow: "0 2px 12px rgba(0,0,0,.07)",
                    }}>
                      <p style={{ fontSize: 11, letterSpacing: 2, textTransform: "uppercase", color: C.muted, marginBottom: 12 }}>
                        Personas por familia confirmada
                      </p>
                      <ResponsiveContainer width="100%" height={200}>
                        <BarChart data={barData.slice(0, 12)}
                          margin={{ top: 0, right: 0, bottom: 24, left: -10 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke={C.olivePale} />
                          <XAxis dataKey="name"
                            tick={{ fontSize: 10, fontFamily: "Jost, sans-serif" }}
                            angle={-35} textAnchor="end" />
                          <YAxis tick={{ fontSize: 11, fontFamily: "Jost, sans-serif" }} />
                          <Tooltip contentStyle={{ fontFamily: "Jost, sans-serif", fontSize: 12 }} />
                          <Bar dataKey="Personas" fill={C.olive} radius={[3, 3, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* ─── LIST ─── */}
            {tab === "list" && (
              <div>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 12 }}>
                  <p style={{ fontSize: 11, letterSpacing: 2, textTransform: "uppercase", color: C.muted }}>
                    {confirmations.length} confirmaciones
                  </p>
                  <input
                    placeholder="Buscar familia, teléfono, correo..."
                    value={searchAdmin}
                    onChange={e => setSearchAdmin(e.target.value)}
                    style={{
                      padding: "8px 14px", border: `1.5px solid ${C.oliveFog}`,
                      borderRadius: 2, fontFamily: "Cormorant Garamond, serif",
                      fontSize: 15, outline: "none", width: 280,
                      background: "white",
                    }}
                  />
                </div>
                {filteredConf.length === 0 ? (
                  <div style={{
                    textAlign: "center", padding: 60,
                    background: "white", borderRadius: 4,
                  }}>
                    <p style={{
                      fontFamily: "Cormorant Garamond, serif",
                      fontStyle: "italic", color: C.muted, fontSize: 20,
                    }}>Aún no hay confirmaciones</p>
                  </div>
                ) : (
                  <div style={{ overflowX: "auto", borderRadius: 4, boxShadow: "0 2px 12px rgba(0,0,0,.07)" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", background: "white" }}>
                      <thead>
                        <tr>
                          <TH>#</TH>
                          <TH>Familia</TH>
                          <TH>Personas</TH>
                          <TH>Teléfono</TH>
                          <TH>Correo</TH>
                          <TH>Fecha</TH>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredConf.map((c, i) => (
                          <tr key={i} style={{ transition: "background .15s" }}
                            onMouseEnter={e => e.currentTarget.style.background = C.olivePale}
                            onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                          >
                            <TD><span style={{ color: C.muted, fontSize: 12 }}>{i + 1}</span></TD>
                            <TD><em>{c.familyName}</em></TD>
                            <TD center>
                              <span style={{
                                background: C.olivePale, color: C.olive,
                                padding: "3px 12px", borderRadius: 2,
                                fontSize: 14, fontWeight: 600,
                              }}>{c.guestCount}</span>
                            </TD>
                            <TD><span style={{ color: C.muted, fontSize: 14 }}>{c.phone || "—"}</span></TD>
                            <TD><span style={{ color: C.muted, fontSize: 14 }}>{c.email || "—"}</span></TD>
                            <TD>
                              <span style={{ color: C.muted, fontSize: 12 }}>
                                {c.timestamp
                                  ? new Date(c.timestamp).toLocaleDateString("es-CO", {
                                      day: "2-digit", month: "short", year: "2-digit"
                                    })
                                  : "—"}
                              </span>
                            </TD>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* ─── UPLOAD ─── */}
            {tab === "upload" && (
              <div>
                <div style={{
                  background: "white", borderRadius: 4, padding: 32, maxWidth: 580,
                  boxShadow: "0 2px 12px rgba(0,0,0,.07)",
                }}>
                  <p style={{
                    fontFamily: "Playfair Display, serif",
                    fontSize: 24, fontStyle: "italic", color: C.olive, marginBottom: 8,
                  }}>Cargar Grupos Familiares</p>
                  <p style={{ fontStyle: "italic", color: C.muted, fontSize: 16, marginBottom: 28 }}>
                    Sube un archivo Excel (.xlsx) con la lista de grupos invitados
                  </p>

                  {/* Upload zone */}
                  <label htmlFor="xls" style={{
                    display: "block", border: `2px dashed ${C.oliveFog}`,
                    borderRadius: 4, padding: "36px 28px", textAlign: "center",
                    background: C.olivePale, marginBottom: 18, cursor: "pointer",
                    transition: "border-color .2s",
                  }}>
                    <p style={{ fontSize: 36, marginBottom: 12 }}>📂</p>
                    <p style={{ fontSize: 17, fontStyle: "italic", marginBottom: 14, color: C.text }}>
                      Haz clic para seleccionar tu archivo
                    </p>
                    <span className="btn btn-g" style={{ fontSize: 11, letterSpacing: 2.5, pointerEvents: "none" }}>
                      Elegir Archivo .xlsx
                    </span>
                    <input id="xls" type="file" accept=".xlsx,.xls"
                      onChange={handleExcel} style={{ display: "none" }} />
                  </label>

                  {uploadMsg && (
                    <div style={{
                      padding: "12px 16px", borderRadius: 2, marginBottom: 20,
                      background: uploadMsg.startsWith("✓") ? "#E8F5E9" : "#FFEBEE",
                      color: uploadMsg.startsWith("✓") ? C.success : C.error,
                      fontFamily: "Jost, sans-serif", fontSize: 13,
                    }}>{uploadMsg}</div>
                  )}

                  {/* Format guide */}
                  <div style={{ background: "#F5F5F0", borderRadius: 4, padding: 18 }}>
                    <p style={{
                      fontSize: 10, letterSpacing: 2.5, textTransform: "uppercase",
                      color: C.muted, marginBottom: 12,
                    }}>Formato requerido del Excel</p>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: C.oliveFog }}>
                          <th style={{ padding: "7px 14px", textAlign: "left", fontWeight: 600 }}>Columna A</th>
                          <th style={{ padding: "7px 14px", textAlign: "left", fontWeight: 600 }}>Columna B</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr style={{ fontWeight: 600, borderBottom: `1px solid ${C.oliveFog}` }}>
                          <td style={{ padding: "7px 14px" }}>Nombre del Grupo</td>
                          <td style={{ padding: "7px 14px" }}>Nº Invitados</td>
                        </tr>
                        {[["Familia García", 4], ["Familia López", 2], ["Familia Torres", 6]].map(([n, v]) => (
                          <tr key={n} style={{ borderBottom: `1px solid ${C.oliveFog}`, fontStyle: "italic", color: C.muted }}>
                            <td style={{ padding: "6px 14px" }}>{n}</td>
                            <td style={{ padding: "6px 14px" }}>{v}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <p style={{ fontSize: 11, color: C.muted, marginTop: 10 }}>
                      La primera fila es el encabezado y se omite automáticamente.
                    </p>
                  </div>

                  {groups.length > 0 && (
                    <p style={{ fontSize: 13, color: C.success, marginTop: 16, fontFamily: "Jost, sans-serif" }}>
                      ✓ {groups.length} grupos cargados actualmente en el sistema
                    </p>
                  )}
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
};

// ═══════════════════════════════════════════════════════
//  MAIN APP
// ═══════════════════════════════════════════════════════
export default function WeddingRSVP() {
  const [page, setPage] = useState("landing");

  return (
    <>
      <GlobalStyles />
      {page === "landing"     && <LandingPage onRSVP={() => setPage("rsvp")} onAdmin={() => setPage("admin-login")} />}
      {page === "rsvp"        && <RSVPPage    onBack={() => setPage("landing")} />}
      {page === "admin-login" && <AdminLogin  onLogin={() => setPage("admin")} onBack={() => setPage("landing")} />}
      {page === "admin"       && <AdminDashboard onLogout={() => setPage("landing")} />}
    </>
  );
}
