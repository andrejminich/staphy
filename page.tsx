'use client';

import { useState, useCallback, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  LineChart, Line, CartesianGrid, ReferenceLine, Cell,
  PieChart, Pie, Legend
} from 'recharts';
import { Upload, RefreshCw, AlertCircle, TrendingUp, Activity, Users, Database, ChevronRight, Microscope } from 'lucide-react';
import Image from 'next/image';
import {
  Isolate, computeStats, getTopN, MONTH_NAMES,
  processRawData, RawRow, EARS_NET, HOSP_DATA
} from '@/lib/mdr';

// ── colour tokens ─────────────────────────────────────────────────────────
const SS_BLUE   = '#1a5fa8';
const SS_TEAL   = '#00b896';
const SS_RED    = '#dc2626';
const SS_GREEN  = '#16a34a';
const SS_ORANGE = '#ea580c';
const SS_AMBER  = '#d97706';
const CHART_COLORS = [SS_BLUE, SS_TEAL, '#7c3aed', SS_ORANGE, '#0891b2', SS_GREEN, '#db2777', '#65a30d'];

const pct = (n: number, d: number) => d > 0 ? Math.round((n / d) * 100) : 0;

type FilterPatogen = 'all' | 'aureus' | 'epidermidis';
type FilterMdr = 'all' | 'mdr' | 'nonmdr';

// ── HEATMAP oddelenie matching ────────────────────────────────────────────
const ODDELENIe_MAP: Record<string, string> = {
  'I. interná': 'I. interná kl. SZU',
  'III. interná': 'III. interná kl. LFUK',
  'KIGM': 'KIGM -  dospelí',
  'neurologická': 'II. neurologická klinika',
  'neurochirurgická': 'Neurochirurgická kl.',
  'geriatrie': 'Klinika geriatrie',
  'chirurgická': 'Chirurgická kl.',
  'úrazovej': 'Klinika úrazovej chir.',
  'urologické': 'Urologické oddelenie',
  'anest': 'Kl. anest. a intenz. med.',
  'novorodenecké': 'Novorodenecké odd.',
  'infektologie': 'KIGM -  dospelí',
  'gynekologická': 'Gyn.-pôrod. kl.',
};

function matchOddelenie(od: string): string {
  const lower = od.toLowerCase();
  for (const [key, val] of Object.entries(ODDELENIe_MAP)) {
    if (lower.includes(key.toLowerCase())) return val;
  }
  return '';
}

// ── Custom Tooltip ────────────────────────────────────────────────────────
const CustomTooltip = ({ active, payload, label }: { active?: boolean; payload?: { value: number; name: string; color: string }[]; label?: string }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '10px 14px', fontSize: 12, boxShadow: '0 4px 12px rgba(0,0,0,0.08)' }}>
      <p style={{ fontWeight: 600, color: '#0a1628', marginBottom: 6 }}>{label}</p>
      {payload.map((p, i) => (
        <p key={i} style={{ color: p.color, margin: '2px 0' }}>{p.name}: <b>{p.value}</b></p>
      ))}
    </div>
  );
};

export default function Dashboard() {
  const [isolates, setIsolates] = useState<Isolate[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [progress, setProgress] = useState<string | null>(null);

  const [filterPatogen, setFilterPatogen] = useState<FilterPatogen>('all');
  const [filterMdr, setFilterMdr] = useState<FilterMdr>('all');
  const [filterRok, setFilterRok] = useState<number | null>(null);
  const [filterOddelenie, setFilterOddelenie] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'overview' | 'heatmap' | 'trends' | 'materials'>('overview');

  // localStorage persistence
  useEffect(() => {
    try {
      const saved = localStorage.getItem('ss_isolates');
      const savedName = localStorage.getItem('ss_filename');
      if (saved) {
        setIsolates(JSON.parse(saved));
        setFileName(savedName || 'Načítané z cache');
      }
    } catch { /* ignore */ }
  }, []);

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setError(null);
    setProgress('Načítavam súbor...');
    try {
      const buffer = await file.arrayBuffer();
      setProgress('Spracovávam dáta...');
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<RawRow>(sheet, { defval: '' });
      if (rows.length === 0) throw new Error('Súbor je prázdny.');
      const required = ['CisloProtokoluOKM', 'Patogen', 'NazovATB', 'CitlivostATB'];
      const cols = Object.keys(rows[0]);
      const missing = required.filter(r => !cols.includes(r));
      if (missing.length > 0) throw new Error(`Chýbajú stĺpce: ${missing.join(', ')}`);
      setProgress(`Kalkulujem MDR pre ${rows.length.toLocaleString()} riadkov...`);
      await new Promise(r => setTimeout(r, 50));
      const result = processRawData(rows);
      setIsolates(result);
      setFileName(file.name);
      try {
        localStorage.setItem('ss_isolates', JSON.stringify(result));
        localStorage.setItem('ss_filename', file.name);
      } catch { /* quota exceeded */ }
      setFilterPatogen('all'); setFilterMdr('all'); setFilterRok(null); setFilterOddelenie(null);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Chyba pri spracovaní');
    } finally { setLoading(false); setProgress(null); e.target.value = ''; }
  }, []);

  // ── filtered isolates ────────────────────────────────────────────────
  const filtered = useMemo(() => isolates.filter(iso => {
    if (filterPatogen === 'aureus' && iso.patogen.indexOf('aureus') < 0) return false;
    if (filterPatogen === 'epidermidis' && iso.patogen.indexOf('epidermidis') < 0) return false;
    if (filterMdr === 'mdr' && !iso.isMdr) return false;
    if (filterMdr === 'nonmdr' && iso.isMdr) return false;
    if (filterRok !== null && iso.rok !== filterRok) return false;
    if (filterOddelenie && iso.oddelenie !== filterOddelenie) return false;
    return true;
  }), [isolates, filterPatogen, filterMdr, filterRok, filterOddelenie]);

  const stats = useMemo(() => computeStats(filtered), [filtered]);
  const allStats = useMemo(() => computeStats(isolates), [isolates]);

  const availableRoky = useMemo(() => {
    const s: Record<number, boolean> = {};
    for (const iso of isolates) s[iso.rok] = true;
    return Object.keys(s).map(Number).sort((a, b) => a - b);
  }, [isolates]);

  const availableOddelenia = useMemo(() => {
    const s: Record<string, boolean> = {};
    for (const iso of isolates) s[iso.oddelenie] = true;
    return Object.keys(s).sort();
  }, [isolates]);

  // ── yearly trend data ────────────────────────────────────────────────
  const yearlyData = useMemo(() => {
    const byYear: Record<number, { mdr: number; nonmdr: number; mrsa: number; total: number }> = {};
    for (const iso of isolates) {
      const y = iso.rok;
      if (!byYear[y]) byYear[y] = { mdr: 0, nonmdr: 0, mrsa: 0, total: 0 };
      byYear[y].total++;
      if (iso.isMdr) byYear[y].mdr++; else byYear[y].nonmdr++;
      if (iso.isMrsa) byYear[y].mrsa++;
    }
    return Object.keys(byYear).map(Number).sort((a, b) => a - b).map(y => ({
      rok: y.toString(),
      mdr: byYear[y].mdr,
      nonmdr: byYear[y].nonmdr,
      mrsa: byYear[y].mrsa,
      total: byYear[y].total,
      mdrPct: pct(byYear[y].mdr, byYear[y].total),
      euBenchmark: EARS_NET[y] || null,
    }));
  }, [isolates]);

  // ── heatmap data ─────────────────────────────────────────────────────
  const heatmapData = useMemo(() => {
    const byOdd: Record<string, { mdr: number; total: number; mrsa: number; material: string[] }> = {};
    for (const iso of filtered) {
      const od = iso.oddelenie;
      if (!byOdd[od]) byOdd[od] = { mdr: 0, total: 0, mrsa: 0, material: [] };
      byOdd[od].total++;
      if (iso.isMdr) byOdd[od].mdr++;
      if (iso.isMrsa) byOdd[od].mrsa++;
      byOdd[od].material.push(iso.material);
    }
    return Object.entries(byOdd)
      .map(([od, v]) => ({
        oddelenie: od,
        mdrRate: pct(v.mdr, v.total),
        total: v.total,
        mrsa: v.mrsa,
        mrsaRate: pct(v.mrsa, v.total),
        topMaterial: getTopN(v.material, 1)[0]?.name || 'N/A',
        hospKey: matchOddelenie(od),
      }))
      .sort((a, b) => b.mdrRate - a.mdrRate)
      .slice(0, 14);
  }, [filtered]);

  // ── materials ────────────────────────────────────────────────────────
  const topMaterials = useMemo(() => getTopN(filtered.map(i => i.material), 10), [filtered]);
  const topMaterialsMdr = useMemo(() => {
    const mdrIsos = filtered.filter(i => i.isMdr);
    return getTopN(mdrIsos.map(i => i.material), 8);
  }, [filtered]);

  // ── monthly data ─────────────────────────────────────────────────────
  const monthlyData = useMemo(() => {
    const byMonth: Record<number, { mdr: number; nonmdr: number }> = {};
    for (const iso of filtered) {
      const m = iso.mesiac;
      if (!byMonth[m]) byMonth[m] = { mdr: 0, nonmdr: 0 };
      if (iso.isMdr) byMonth[m].mdr++; else byMonth[m].nonmdr++;
    }
    return Object.keys(byMonth).map(Number).sort((a, b) => a - b)
      .map(m => ({ mesiac: MONTH_NAMES[m], mdr: byMonth[m].mdr, nonmdr: byMonth[m].nonmdr }));
  }, [filtered]);

  // ── heatmap colour ───────────────────────────────────────────────────
  function heatColor(rate: number): string {
    if (rate >= 80) return '#b91c1c';
    if (rate >= 60) return '#dc2626';
    if (rate >= 40) return '#ea580c';
    if (rate >= 20) return '#d97706';
    if (rate >= 10) return '#ca8a04';
    return '#16a34a';
  }

  // ── MRSA per 1000 hospitalizácií ─────────────────────────────────────
  const mrsaPer1000 = useMemo(() => {
    if (!filterRok) return null;
    const totalHosp = Object.values(HOSP_DATA).reduce((sum, d) => sum + (d[filterRok] || 0), 0);
    if (totalHosp === 0) return null;
    const mrsaCount = filtered.filter(i => i.isMrsa).length;
    return ((mrsaCount / totalHosp) * 1000).toFixed(1);
  }, [filtered, filterRok]);

  // ── pie data ─────────────────────────────────────────────────────────
  const pieData = [
    { name: 'S. aureus', value: stats.aureus, color: SS_BLUE },
    { name: 'S. epidermidis', value: stats.epidermidis, color: SS_TEAL },
  ];

  const mdrPieData = [
    { name: 'MDR', value: stats.mdr, color: SS_RED },
    { name: 'Non-MDR', value: stats.nonMdr, color: SS_GREEN },
  ];

  // ── EMPTY STATE ───────────────────────────────────────────────────────
  if (isolates.length === 0) {
    return (
      <div style={{ minHeight: '100vh', background: 'linear-gradient(135deg, #0a1628 0%, #0f3460 60%, #1a5fa8 100%)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ textAlign: 'center', padding: '2rem' }}>
          <Image src="/logo.png" alt="StaphySearch" width={320} height={86} style={{ height: 70, width: 'auto', margin: '0 auto 2rem' }} />
          <div style={{ background: 'rgba(255,255,255,0.06)', borderRadius: 20, border: '1px solid rgba(0,184,150,0.2)', padding: '2.5rem 3rem', backdropFilter: 'blur(20px)', maxWidth: 480 }}>
            <p style={{ color: 'rgba(255,255,255,0.5)', fontSize: 13, marginBottom: '1.5rem', textTransform: 'uppercase', letterSpacing: '0.08em' }}>Virtuálna Nemocnica · Digitálne Dvojča</p>
            <p style={{ color: 'rgba(255,255,255,0.7)', fontSize: 14, marginBottom: '2rem', lineHeight: 1.7 }}>Nahraj export mikrobiologických dát (XLSX alebo CSV) pre aktiváciu dashboardu.</p>
            <label style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 8, background: 'linear-gradient(135deg, #1a5fa8, #00b896)', color: '#fff', padding: '10px 24px', borderRadius: 12, fontSize: 14, fontWeight: 500 }}>
              {loading ? <RefreshCw size={16} style={{ animation: 'spin 1s linear infinite' }} /> : <Upload size={16} />}
              {loading ? (progress || 'Spracovávam...') : 'Nahrať dáta (.xlsx / .csv)'}
              <input type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={handleUpload} disabled={loading} />
            </label>
            {error && <div style={{ marginTop: 16, padding: '10px 14px', background: 'rgba(220,38,38,0.15)', border: '1px solid rgba(220,38,38,0.3)', borderRadius: 10, color: '#fca5a5', fontSize: 13, display: 'flex', alignItems: 'center', gap: 8 }}><AlertCircle size={14} />{error}</div>}
          </div>
          <p style={{ color: 'rgba(255,255,255,0.25)', fontSize: 11, marginTop: '2rem' }}>StaphySearch · Targeting Staphylococci with Digital Precision</p>
        </div>
        <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh', background: '#f0f4f8' }}>

      {/* ── HEADER ── */}
      <header style={{ background: 'linear-gradient(135deg, #0a1628 0%, #0f3460 100%)', borderBottom: '1px solid rgba(0,184,150,0.15)', position: 'sticky', top: 0, zIndex: 100 }}>
        <div style={{ maxWidth: 1400, margin: '0 auto', padding: '0 1.5rem', display: 'flex', alignItems: 'center', justifyContent: 'space-between', height: 60 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 24 }}>
            <Image src="/logo.png" alt="StaphySearch" width={140} height={38} style={{ height: 34, width: 'auto' }} />
            <div style={{ display: 'flex', gap: 4 }}>
              {([['overview','Prehľad'],['heatmap','Heatmapa'],['trends','Trendy'],['materials','Materiály']] as ['overview'|'heatmap'|'trends'|'materials', string][]).map(([tab, label]) => (
                <button key={tab} onClick={() => setActiveTab(tab)}
                  style={{ fontSize: 12, fontWeight: 500, padding: '5px 14px', borderRadius: 8, border: '1px solid transparent', cursor: 'pointer', background: activeTab === tab ? 'rgba(0,184,150,0.2)' : 'transparent', color: activeTab === tab ? '#fff' : 'rgba(255,255,255,0.55)', borderColor: activeTab === tab ? 'rgba(0,184,150,0.4)' : 'transparent', transition: 'all 0.15s' }}>
                  {label}
                </button>
              ))}
              <a href="/chapter2" style={{ fontSize: 12, fontWeight: 500, padding: '5px 14px', borderRadius: 8, border: '1px solid rgba(255,255,255,0.15)', cursor: 'pointer', background: 'transparent', color: 'rgba(255,255,255,0.45)', textDecoration: 'none', display: 'flex', alignItems: 'center', gap: 4, marginLeft: 8 }}>
                Vrstva 2 — Vedec <ChevronRight size={12} />
              </a>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.35)', fontFamily: 'monospace' }}>{fileName}</span>
            <label style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, background: 'rgba(255,255,255,0.08)', color: 'rgba(255,255,255,0.7)', padding: '5px 12px', borderRadius: 8, fontSize: 12, border: '1px solid rgba(255,255,255,0.12)' }}>
              {loading ? <RefreshCw size={12} style={{ animation: 'spin 1s linear infinite' }} /> : <Upload size={12} />}
              Nahrať nové dáta
              <input type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={handleUpload} disabled={loading} />
            </label>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1400, margin: '0 auto', padding: '1.5rem' }}>

        {/* ── FILTERS ── */}
        <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.1)', padding: '1rem 1.25rem', marginBottom: '1.25rem', display: 'flex', flexWrap: 'wrap', gap: 16, alignItems: 'center' }}>
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', marginRight: 4 }}>Patogén</span>
            {(['all','aureus','epidermidis'] as FilterPatogen[]).map(v => (
              <button key={v} onClick={() => setFilterPatogen(v)} className={`ss-chip ${filterPatogen === v ? 'active' : ''}`}>
                {v === 'all' ? 'Všetky' : v === 'aureus' ? 'S. aureus' : 'S. epidermidis'}
              </button>
            ))}
          </div>
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', marginRight: 4 }}>MDR</span>
            {(['all','mdr','nonmdr'] as FilterMdr[]).map(v => (
              <button key={v} onClick={() => setFilterMdr(v)} className={`ss-chip ${filterMdr === v ? 'active' : ''}`}>
                {v === 'all' ? 'Všetky' : v === 'mdr' ? 'MDR' : 'Non-MDR'}
              </button>
            ))}
          </div>
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', marginRight: 4 }}>Rok</span>
            <button onClick={() => setFilterRok(null)} className={`ss-chip ${filterRok === null ? 'active' : ''}`}>Všetky</button>
            {availableRoky.map(r => (
              <button key={r} onClick={() => setFilterRok(filterRok === r ? null : r)} className={`ss-chip ${filterRok === r ? 'active' : ''}`}>{r}</button>
            ))}
          </div>
          <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
            <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', marginRight: 4 }}>Oddelenie</span>
            <select value={filterOddelenie || ''} onChange={e => setFilterOddelenie(e.target.value || null)}
              style={{ fontSize: 12, padding: '5px 10px', borderRadius: 8, border: '1px solid rgba(26,95,168,0.15)', color: '#475569', background: '#fff' }}>
              <option value="">— všetky —</option>
              {availableOddelenia.map(o => <option key={o} value={o}>{o}</option>)}
            </select>
            {filterOddelenie && <button onClick={() => setFilterOddelenie(null)} style={{ fontSize: 11, color: SS_RED, cursor: 'pointer', background: 'none', border: 'none' }}>✕</button>}
          </div>
        </div>

        {/* ── STAT CARDS ── */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, minmax(0,1fr))', gap: 12, marginBottom: '1.25rem' }}>
          {[
            { val: filtered.length.toLocaleString(), lbl: 'Izolátov celkom', sub: `z ${allStats.total} celkovo`, color: SS_BLUE, icon: <Database size={18} /> },
            { val: `${stats.mdrRate}%`, lbl: 'MDR rate', sub: `${stats.mdr} MDR kmeňov`, color: SS_RED, icon: <AlertCircle size={18} /> },
            { val: stats.mrsa.toString(), lbl: 'MRSA / MRCoNS', sub: `${pct(stats.mrsa, stats.total)}% z filtrovaných`, color: SS_ORANGE, icon: <Activity size={18} /> },
            { val: stats.aureus.toString(), lbl: 'S. aureus', sub: `${pct(stats.aureus, stats.total)}% izolátov`, color: SS_BLUE, icon: <Microscope size={18} /> },
            { val: mrsaPer1000 ? `${mrsaPer1000}‰` : '—', lbl: 'MRSA / 1000 hosp.', sub: mrsaPer1000 ? `EU priemer: 4.64/100k` : 'Vyber rok pre výpočet', color: SS_TEAL, icon: <TrendingUp size={18} /> },
          ].map((s, i) => (
            <div key={i} style={{ background: '#fff', borderRadius: 14, padding: '1rem 1.25rem', border: '1px solid rgba(26,95,168,0.08)', boxShadow: '0 1px 3px rgba(10,22,40,0.05)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 8 }}>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: `${s.color}15`, display: 'flex', alignItems: 'center', justifyContent: 'center', color: s.color }}>{s.icon}</div>
              </div>
              <div style={{ fontSize: '1.9rem', fontWeight: 600, lineHeight: 1, color: s.color, fontFamily: 'monospace' }}>{s.val}</div>
              <div style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em', color: '#64748b', marginTop: 5 }}>{s.lbl}</div>
              <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 2 }}>{s.sub}</div>
            </div>
          ))}
        </div>

        {/* ── TAB: OVERVIEW ── */}
        {activeTab === 'overview' && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.25rem' }}>
            {/* Pie charts */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.25rem 1.5rem' }}>
              <p className="ss-section-title">Rozdelenie patogénov</p>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 0 }}>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={pieData} cx="50%" cy="50%" outerRadius={75} dataKey="value" label={({ name, value }) => `${value}`}>
                      {pieData.map((e, i) => <Cell key={i} fill={e.color} />)}
                    </Pie>
                    <Legend iconType="circle" iconSize={8} formatter={(v) => <span style={{ fontSize: 12 }}>{v}</span>} />
                    <Tooltip content={<CustomTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={mdrPieData} cx="50%" cy="50%" outerRadius={75} dataKey="value" label={({ name, value }) => `${value}`}>
                      {mdrPieData.map((e, i) => <Cell key={i} fill={e.color} />)}
                    </Pie>
                    <Legend iconType="circle" iconSize={8} formatter={(v) => <span style={{ fontSize: 12 }}>{v}</span>} />
                    <Tooltip content={<CustomTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Top oddelenia */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.25rem 1.5rem' }}>
              <p className="ss-section-title">Najväčší burden — oddelenia</p>
              {getTopN(filtered.map(i => i.oddelenie), 8).map(({ name, count }, i) => {
                const mdrCount = filtered.filter(iso => iso.oddelenie === name && iso.isMdr).length;
                const mdrPct = pct(mdrCount, count);
                return (
                  <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 }}>
                    <div style={{ fontSize: 11, color: '#94a3b8', width: 16, textAlign: 'right', fontFamily: 'monospace' }}>{i + 1}</div>
                    <div style={{ flex: 1, fontSize: 12, color: '#334155', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={name}>{name}</div>
                    <div style={{ width: 100, background: '#f1f5f9', borderRadius: 4, height: 16, position: 'relative', overflow: 'hidden' }}>
                      <div style={{ position: 'absolute', left: 0, top: 0, height: '100%', width: `${pct(count, filtered.length)}%`, background: `linear-gradient(90deg, ${SS_BLUE}, ${SS_TEAL})`, borderRadius: 4 }} />
                    </div>
                    <div style={{ fontSize: 11, fontFamily: 'monospace', color: '#475569', width: 28, textAlign: 'right' }}>{count}</div>
                    <span style={{ fontSize: 10, padding: '2px 6px', borderRadius: 5, background: mdrPct >= 50 ? '#fef2f2' : '#f0fdf4', color: mdrPct >= 50 ? '#b91c1c' : '#15803d', fontWeight: 600 }}>{mdrPct}%</span>
                  </div>
                );
              })}
            </div>

            {/* Mesačný trend */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.25rem 1.5rem' }}>
              <p className="ss-section-title">Mesačný trend — MDR vs Non-MDR</p>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={monthlyData} margin={{ top: 5, right: 5, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis dataKey="mesiac" tick={{ fontSize: 11 }} />
                  <YAxis tick={{ fontSize: 11 }} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar dataKey="mdr" name="MDR" stackId="a" fill={SS_RED} />
                  <Bar dataKey="nonmdr" name="Non-MDR" stackId="a" fill={SS_TEAL} radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* ATB profil summary */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.25rem 1.5rem' }}>
              <p className="ss-section-title">Najčastejšie materiály</p>
              {topMaterials.slice(0, 8).map(({ name, count }, i) => (
                <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 }}>
                  <div style={{ fontSize: 11, color: '#94a3b8', width: 16, textAlign: 'right' }}>{i + 1}</div>
                  <div style={{ flex: 1, fontSize: 12, color: '#334155' }}>{name}</div>
                  <div style={{ width: 100, background: '#f1f5f9', borderRadius: 4, height: 16, position: 'relative', overflow: 'hidden' }}>
                    <div style={{ position: 'absolute', left: 0, top: 0, height: '100%', width: `${pct(count, filtered.length)}%`, background: `linear-gradient(90deg, ${SS_TEAL}, ${SS_BLUE})`, borderRadius: 4 }} />
                  </div>
                  <div style={{ fontSize: 11, fontFamily: 'monospace', color: '#475569', width: 28, textAlign: 'right' }}>{count}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── TAB: HEATMAPA ── */}
        {activeTab === 'heatmap' && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.25rem' }}>
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem' }}>
              <p className="ss-section-title">Heatmapa oddelení — MDR rate</p>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10, marginBottom: 16 }}>
                {heatmapData.slice(0, 12).map(d => (
                  <div key={d.oddelenie} className="heatmap-cell"
                    style={{ background: `${heatColor(d.mdrRate)}18`, border: `2px solid ${heatColor(d.mdrRate)}40`, padding: '14px 12px', cursor: 'pointer' }}
                    title={`${d.oddelenie}: ${d.mdrRate}% MDR, ${d.total} izolátov`}
                    onClick={() => setFilterOddelenie(d.oddelenie === filterOddelenie ? null : d.oddelenie)}>
                    <div style={{ fontSize: '1.6rem', fontWeight: 700, color: heatColor(d.mdrRate), fontFamily: 'monospace', lineHeight: 1 }}>{d.mdrRate}%</div>
                    <div style={{ fontSize: 10, color: '#64748b', marginTop: 4, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.04em', overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis' }} title={d.oddelenie}>{d.oddelenie.slice(0, 20)}</div>
                    <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>{d.total} izolátov · {d.mrsa} MRSA</div>
                  </div>
                ))}
              </div>
              {/* Legend */}
              <div style={{ display: 'flex', gap: 8, alignItems: 'center', fontSize: 11, color: '#64748b' }}>
                <span>MDR rate:</span>
                {[['≥80%','#b91c1c'],['60–80%','#dc2626'],['40–60%','#ea580c'],['20–40%','#d97706'],['<20%','#16a34a']].map(([label, color]) => (
                  <span key={label} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                    <span style={{ width: 10, height: 10, borderRadius: 2, background: color, display: 'inline-block' }} />{label}
                  </span>
                ))}
              </div>
            </div>

            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem' }}>
              <p className="ss-section-title">Detail oddelení — tabuľka</p>
              <table className="ss-table">
                <thead>
                  <tr>
                    <th>Oddelenie</th>
                    <th>Izolátov</th>
                    <th>MDR%</th>
                    <th>MRSA</th>
                    <th>Top materiál</th>
                  </tr>
                </thead>
                <tbody>
                  {heatmapData.map(d => (
                    <tr key={d.oddelenie} style={{ cursor: 'pointer' }} onClick={() => setFilterOddelenie(d.oddelenie === filterOddelenie ? null : d.oddelenie)}>
                      <td style={{ fontSize: 11, maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={d.oddelenie}>{d.oddelenie}</td>
                      <td style={{ fontFamily: 'monospace', fontWeight: 600 }}>{d.total}</td>
                      <td><span style={{ display: 'inline-flex', padding: '2px 8px', borderRadius: 6, fontSize: 11, fontWeight: 700, background: `${heatColor(d.mdrRate)}18`, color: heatColor(d.mdrRate) }}>{d.mdrRate}%</span></td>
                      <td style={{ fontFamily: 'monospace' }}>{d.mrsa}</td>
                      <td style={{ fontSize: 11, color: '#64748b' }}>{d.topMaterial}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── TAB: TRENDY ── */}
        {activeTab === 'trends' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            {/* Ročný trend + EU benchmark */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
                <p className="ss-section-title" style={{ margin: 0 }}>Ročný trend MDR kmeňov</p>
                <div style={{ display: 'flex', gap: 16, fontSize: 11 }}>
                  <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><span style={{ width: 12, height: 12, borderRadius: 2, background: SS_RED, display: 'inline-block' }} /> MDR</span>
                  <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><span style={{ width: 12, height: 12, borderRadius: 2, background: SS_TEAL, display: 'inline-block' }} /> Non-MDR</span>
                  <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><span style={{ width: 20, height: 2, background: SS_ORANGE, display: 'inline-block' }} /> EU benchmark (EARS-Net)</span>
                </div>
              </div>
              <ResponsiveContainer width="100%" height={280}>
                <BarChart data={yearlyData} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis dataKey="rok" tick={{ fontSize: 12 }} />
                  <YAxis yAxisId="left" tick={{ fontSize: 11 }} />
                  <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11 }} domain={[0, 30]} unit="%" />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar yAxisId="left" dataKey="mdr" name="MDR" fill={SS_RED} stackId="a" />
                  <Bar yAxisId="left" dataKey="nonmdr" name="Non-MDR" fill={SS_TEAL} stackId="a" radius={[4, 4, 0, 0]} />
                  <Line yAxisId="right" type="monotone" dataKey="mdrPct" name="MDR%" stroke={SS_AMBER} strokeWidth={2} dot={{ r: 4 }} />
                  <ReferenceLine yAxisId="left" x="2020" stroke="rgba(234,88,12,0.4)" strokeDasharray="5 3" label={{ value: 'COVID', position: 'top', fontSize: 10, fill: '#ea580c' }} />
                  <ReferenceLine yAxisId="left" x="2023" stroke="rgba(234,88,12,0.25)" strokeDasharray="5 3" />
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* EU Benchmark porovnanie */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem' }}>
              <p className="ss-section-title">Porovnanie s európskym priemerom (EARS-Net)</p>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 12 }}>
                {yearlyData.filter(d => EARS_NET[Number(d.rok)]).map(d => {
                  const eu = EARS_NET[Number(d.rok)];
                  const diff = d.mdrPct - eu;
                  return (
                    <div key={d.rok} style={{ background: '#f8fafc', borderRadius: 12, padding: '14px', border: '1px solid #e2e8f0', textAlign: 'center' }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: '#475569', marginBottom: 8 }}>{d.rok}</div>
                      <div style={{ fontSize: '1.4rem', fontWeight: 700, fontFamily: 'monospace', color: d.mdrPct > eu ? SS_RED : SS_GREEN }}>{d.mdrPct}%</div>
                      <div style={{ fontSize: 10, color: '#94a3b8', margin: '4px 0' }}>Kramáre</div>
                      <div style={{ fontSize: 12, fontFamily: 'monospace', color: '#64748b' }}>{eu}%</div>
                      <div style={{ fontSize: 10, color: '#94a3b8', marginBottom: 6 }}>EU EARS-Net</div>
                      <span style={{ fontSize: 10, padding: '2px 8px', borderRadius: 6, fontWeight: 700, background: diff > 0 ? '#fef2f2' : '#f0fdf4', color: diff > 0 ? '#b91c1c' : '#15803d' }}>
                        {diff > 0 ? '+' : ''}{diff.toFixed(1)}%
                      </span>
                    </div>
                  );
                })}
              </div>
              <p style={{ fontSize: 11, color: '#94a3b8', marginTop: 12 }}>* EARS-Net udáva % MRSA z invazívnych izolátov S. aureus. Hodnoty Kramáre = MDR rate zo všetkých izolátov (odlišná metodika — orientačné porovnanie).</p>
            </div>

            {/* Aureus vs Epidermidis trend */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem' }}>
              <p className="ss-section-title">S. aureus vs S. epidermidis — ročný vývoj</p>
              <ResponsiveContainer width="100%" height={220}>
                <LineChart data={(() => {
                  const byYear: Record<number, { aureus: number; epidermidis: number }> = {};
                  for (const iso of isolates) {
                    const y = iso.rok;
                    if (!byYear[y]) byYear[y] = { aureus: 0, epidermidis: 0 };
                    if (iso.patogen.indexOf('aureus') >= 0) byYear[y].aureus++;
                    else byYear[y].epidermidis++;
                  }
                  return Object.keys(byYear).map(Number).sort((a, b) => a - b)
                    .map(y => ({ rok: y.toString(), aureus: byYear[y].aureus, epidermidis: byYear[y].epidermidis }));
                })()} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis dataKey="rok" tick={{ fontSize: 12 }} />
                  <YAxis tick={{ fontSize: 11 }} />
                  <Tooltip content={<CustomTooltip />} />
                  <Line type="monotone" dataKey="aureus" stroke={SS_BLUE} strokeWidth={2.5} dot={{ r: 5 }} name="S. aureus" />
                  <Line type="monotone" dataKey="epidermidis" stroke={SS_TEAL} strokeWidth={2.5} dot={{ r: 5 }} name="S. epidermidis" />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        {/* ── TAB: MATERIÁLY ── */}
        {activeTab === 'materials' && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.25rem' }}>
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem' }}>
              <p className="ss-section-title">Všetky materiály</p>
              <ResponsiveContainer width="100%" height={320}>
                <BarChart data={topMaterials} layout="vertical" margin={{ left: 10, right: 20 }}>
                  <XAxis type="number" tick={{ fontSize: 11 }} />
                  <YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={150} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar dataKey="count" name="Počet" radius={[0, 6, 6, 0]}>
                    {topMaterials.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem' }}>
              <p className="ss-section-title">Materiály MDR kmeňov</p>
              <ResponsiveContainer width="100%" height={320}>
                <BarChart data={topMaterialsMdr} layout="vertical" margin={{ left: 10, right: 20 }}>
                  <XAxis type="number" tick={{ fontSize: 11 }} />
                  <YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={150} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar dataKey="count" name="MDR" fill={SS_RED} radius={[0, 6, 6, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Materiál × patogén */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.1)', padding: '1.5rem', gridColumn: '1 / -1' }}>
              <p className="ss-section-title">Materiál podľa patogéna</p>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.5rem' }}>
                {(['aureus', 'epidermidis'] as const).map(pathogen => {
                  const subset = filtered.filter(i => i.patogen.indexOf(pathogen) >= 0);
                  const mats = getTopN(subset.map(i => i.material), 8);
                  return (
                    <div key={pathogen}>
                      <p style={{ fontSize: 12, fontWeight: 600, color: pathogen === 'aureus' ? SS_BLUE : SS_TEAL, marginBottom: 12, fontStyle: 'italic' }}>
                        Staphylococcus {pathogen} ({subset.length} izolátov)
                      </p>
                      {mats.map(({ name, count }) => (
                        <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                          <div style={{ fontSize: 12, flex: 1, color: '#334155' }}>{name}</div>
                          <div style={{ width: 80, background: '#f1f5f9', borderRadius: 4, height: 14, position: 'relative', overflow: 'hidden' }}>
                            <div style={{ position: 'absolute', left: 0, top: 0, height: '100%', width: `${pct(count, subset.length)}%`, background: pathogen === 'aureus' ? SS_BLUE : SS_TEAL, borderRadius: 4 }} />
                          </div>
                          <div style={{ fontSize: 11, fontFamily: 'monospace', color: '#64748b', width: 28, textAlign: 'right' }}>{count}</div>
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        <footer style={{ textAlign: 'center', fontSize: 11, color: '#94a3b8', padding: '1.5rem 0', marginTop: '0.5rem' }}>
          StaphySearch · MDR: Magiorakos et al. 2012 (ECDC/CDC) · EARS-Net benchmark: European Antimicrobial Resistance Surveillance Network 2023
        </footer>
      </main>
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}
