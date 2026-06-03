'use client';

import { useState, useCallback, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  LineChart, Line, CartesianGrid, ReferenceLine, Cell,
  PieChart, Pie, Legend
} from 'recharts';
import { Upload, RefreshCw, AlertCircle, ChevronLeft, ChevronRight, TrendingUp, TrendingDown, Minus } from 'lucide-react';
import Image from 'next/image';
import { Isolate, computeStats, getTopN, MONTH_NAMES, processRawData, RawRow, EARS_NET, HOSP_DATA, getPandemicPeriod, PANDEMIC_PERIODS } from '@/lib/mdr';

// ── constants ─────────────────────────────────────────────────────────────
const SS_BLUE = '#1a5fa8';
const SS_TEAL = '#00b896';
const SS_RED = '#dc2626';
const SS_GREEN = '#16a34a';
const SS_ORANGE = '#ea580c';
const SS_AMBER = '#d97706';
const CHART_COLORS = [SS_BLUE, SS_TEAL, '#7c3aed', SS_ORANGE, '#0891b2', SS_GREEN, '#db2777', '#65a30d', '#854d0e', '#374151'];

// ── Ward groups (ECDC/HAI-Net classification) ────────────────────────────
const WARD_GROUPS: { key: string; label: string; keywords: string[] }[] = [
  { key: 'icu', label: 'ICU / Intenzívna', keywords: ['anest', 'intenz', 'jis', 'icu', 'arip'] },
  { key: 'surgical', label: 'Chirurgické', keywords: ['chirurg', 'úrazov', 'urazov', 'neurochirurg', 'urologick'] },
  { key: 'medical', label: 'Interné / Medicínske', keywords: ['interná', 'interna', 'neurolog', 'geriatr', 'nefrol'] },
  { key: 'infectious', label: 'Infektológia', keywords: ['kigm', 'infektol', 'geograf'] },
  { key: 'neonatal', label: 'Novorodenecké / Pôrodné', keywords: ['novorodeneck', 'neonat', 'gyn', 'pôrod', 'porod'] },
  { key: 'longterm', label: 'Dlhodobá starostlivosť', keywords: ['dlhodobo', 'dlhod'] },
];

function getWardGroup(oddelenie: string): string {
  const lower = oddelenie.toLowerCase();
  for (const g of WARD_GROUPS) {
    if (g.keywords.some(k => lower.includes(k))) return g.key;
  }
  return 'other';
}

// Exclude ambulances and Ružinov
function shouldExclude(oddelenie: string): boolean {
  const lower = oddelenie.toLowerCase();
  return lower.includes('amb') || lower.includes(' ru ') || lower.startsWith('ru ') || lower.includes('ružinov') || lower.includes('ruzinov');
}

// Only years >= 2017
function shouldExcludeYear(rok: number): boolean {
  return rok < 2017;
}

const pct = (n: number, d: number) => d > 0 ? Math.round((n / d) * 100) : 0;

// Lôžkové oddelenia (match hospitalization data)
const LÔZKOVÉ_KEYS = Object.keys(HOSP_DATA);
function isLôžkové(oddelenie: string): boolean {
  const lower = oddelenie.toLowerCase();
  return LÔZKOVÉ_KEYS.some(k => lower.includes(k.toLowerCase().trim().slice(0, 8)) || k.toLowerCase().trim().includes(lower.slice(0, 8)));
}

function heatColor(rate: number): string {
  if (rate >= 80) return '#b91c1c';
  if (rate >= 60) return '#dc2626';
  if (rate >= 40) return '#ea580c';
  if (rate >= 20) return '#d97706';
  if (rate >= 10) return '#ca8a04';
  return '#16a34a';
}

const CustomTooltip = ({ active, payload, label }: { active?: boolean; payload?: { value: number; name: string; color: string }[]; label?: string }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '10px 14px', fontSize: 12, boxShadow: '0 4px 12px rgba(0,0,0,0.08)' }}>
      <p style={{ fontWeight: 600, color: '#0a1628', marginBottom: 6 }}>{label}</p>
      {payload.map((p, i) => <p key={i} style={{ color: p.color, margin: '2px 0' }}>{p.name}: <b>{p.value}</b></p>)}
    </div>
  );
};

type ActiveTab = 'prehlad' | 'heatmapa' | 'pandemia';
type FilterPatogen = 'all' | 'aureus' | 'epidermidis';
type FilterMdr = 'all' | 'mdr' | 'nonmdr';
type HeatPeriod = 'tyzdne' | 'mesiac' | 'kvartal' | 'rok';

export default function Dashboard() {
  const [isolates, setIsolates] = useState<Isolate[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [progress, setProgress] = useState<string | null>(null);

  const [activeTab, setActiveTab] = useState<ActiveTab>('prehlad');
  const [filterPatogen, setFilterPatogen] = useState<FilterPatogen>('all');
  const [filterMdr, setFilterMdr] = useState<FilterMdr>('all');
  const [filterRok, setFilterRok] = useState<number | null>(null);
  const [filterOddelenie, setFilterOddelenie] = useState<string | null>(null);
  const [filterGroup, setFilterGroup] = useState<string | null>(null);
  const [heatPeriod, setHeatPeriod] = useState<HeatPeriod>('mesiac');
  const [pandemiaOddelenie, setPandemiaOddelenie] = useState<string | null>(null);
  // Slider for dynamic chart
  const [sliderOffset, setSliderOffset] = useState(0); // quarters back from now
  const SLIDER_WINDOW = 8; // quarters shown

  // localStorage
  useEffect(() => {
    try {
      const saved = localStorage.getItem('ss_isolates_v2');
      const savedName = localStorage.getItem('ss_filename_v2');
      if (saved) { setIsolates(JSON.parse(saved)); setFileName(savedName || 'Cache'); }
    } catch { /* ignore */ }
  }, []);

  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    setLoading(true); setError(null); setProgress('Načítavam...');
    try {
      const buffer = await file.arrayBuffer();
      setProgress('Spracovávam...');
      const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
      const rows = XLSX.utils.sheet_to_json<RawRow>(wb.Sheets[wb.SheetNames[0]], { defval: '' });
      if (!rows.length) throw new Error('Súbor je prázdny.');
      setProgress(`MDR kalkulácia pre ${rows.length.toLocaleString()} riadkov...`);
      await new Promise(r => setTimeout(r, 50));
      const result = processRawData(rows);
      setIsolates(result); setFileName(file.name);
      try { localStorage.setItem('ss_isolates_v2', JSON.stringify(result)); localStorage.setItem('ss_filename_v2', file.name); } catch { /* quota */ }
      setFilterPatogen('all'); setFilterMdr('all'); setFilterRok(null); setFilterOddelenie(null);
    } catch (err: unknown) { setError(err instanceof Error ? err.message : 'Chyba'); }
    finally { setLoading(false); setProgress(null); e.target.value = ''; }
  }, []);

  // ── base filtered (global exclusions) ─────────────────────────────────
  const base = useMemo(() => isolates.filter(iso =>
    !shouldExclude(iso.oddelenie) && !shouldExcludeYear(iso.rok)
  ), [isolates]);

  // ── user filters applied ────────────────────────────────────────────────
  const filtered = useMemo(() => base.filter(iso => {
    if (filterPatogen === 'aureus' && iso.patogen.indexOf('aureus') < 0) return false;
    if (filterPatogen === 'epidermidis' && iso.patogen.indexOf('epidermidis') < 0) return false;
    if (filterMdr === 'mdr' && !iso.isMdr) return false;
    if (filterMdr === 'nonmdr' && iso.isMdr) return false;
    if (filterRok !== null && iso.rok !== filterRok) return false;
    if (filterOddelenie && iso.oddelenie !== filterOddelenie) return false;
    if (filterGroup && getWardGroup(iso.oddelenie) !== filterGroup) return false;
    return true;
  }), [base, filterPatogen, filterMdr, filterRok, filterOddelenie, filterGroup]);

  const stats = useMemo(() => computeStats(filtered), [filtered]);
  const allStats = useMemo(() => computeStats(base), [base]);

  const availableRoky = useMemo(() => {
    const s: Record<number, boolean> = {};
    for (const iso of base) s[iso.rok] = true;
    return Object.keys(s).map(Number).sort((a, b) => a - b);
  }, [base]);

  const availableOddelenia = useMemo(() => {
    const s: Record<string, boolean> = {};
    for (const iso of base) s[iso.oddelenie] = true;
    return Object.keys(s).sort();
  }, [base]);

  // ── PREHĽAD: Last 2 quarters dynamic ──────────────────────────────────
  // Build quarter list from data
  const allQuarters = useMemo(() => {
    const qSet: Record<string, { rok: number; q: number; label: string; from: Date; to: Date }> = {};
    for (const iso of base) {
      const d = new Date(iso.datumOdberu);
      const q = Math.ceil((d.getMonth() + 1) / 3);
      const key = `${iso.rok}-Q${q}`;
      if (!qSet[key]) {
        const from = new Date(iso.rok, (q - 1) * 3, 1);
        const to = new Date(iso.rok, q * 3, 0);
        qSet[key] = { rok: iso.rok, q, label: `${iso.rok} Q${q}`, from, to };
      }
    }
    return Object.values(qSet).sort((a, b) => a.rok !== b.rok ? a.rok - b.rok : a.q - b.q);
  }, [base]);

  const maxOffset = Math.max(0, allQuarters.length - SLIDER_WINDOW);

  const visibleQuarters = useMemo(() => {
    const start = Math.max(0, allQuarters.length - SLIDER_WINDOW - sliderOffset);
    return allQuarters.slice(start, start + SLIDER_WINDOW);
  }, [allQuarters, sliderOffset]);

  const quarterChartData = useMemo(() => {
    return visibleQuarters.map(qInfo => {
      const qIsos = filtered.filter(iso => {
        const d = new Date(iso.datumOdberu);
        return d >= qInfo.from && d <= qInfo.to;
      });
      const s = computeStats(qIsos);
      return { label: qInfo.label, mdr: s.mdr, nonmdr: s.nonMdr, mrsa: s.mrsa, total: qIsos.length, mdrPct: s.mdrRate };
    });
  }, [filtered, visibleQuarters]);

  // Last 2 quarters
  const last2Q = useMemo(() => allQuarters.slice(-2), [allQuarters]);
  const last2QIsos = useMemo(() => filtered.filter(iso => {
    const d = new Date(iso.datumOdberu);
    return last2Q.some(q => d >= q.from && d <= q.to);
  }), [filtered, last2Q]);

  const topOddLast2Q = useMemo(() => getTopN(last2QIsos.map(i => i.oddelenie), 8), [last2QIsos]);
  const topOddMdrLast2Q = useMemo(() => {
    const mdrIsos = last2QIsos.filter(i => i.isMdr);
    return getTopN(mdrIsos.map(i => i.oddelenie), 8);
  }, [last2QIsos]);
  const topMatLast2Q = useMemo(() => getTopN(last2QIsos.map(i => i.material), 8), [last2QIsos]);

  // ── HEATMAPA: weekly/monthly/quarterly/yearly ──────────────────────────
  const heatmapData = useMemo(() => {
    // Group base by oddelenie + period
    type PeriodKey = string;
    const byOddPeriod: Record<string, Record<PeriodKey, number>> = {};

    for (const iso of base) {
      if (!isLôžkové(iso.oddelenie)) continue;
      const d = new Date(iso.datumOdberu);
      let periodKey: string;
      if (heatPeriod === 'tyzdne') {
        const startOfYear = new Date(iso.rok, 0, 1);
        const week = Math.ceil(((d.getTime() - startOfYear.getTime()) / 86400000 + startOfYear.getDay() + 1) / 7);
        periodKey = `${iso.rok}-W${String(week).padStart(2, '0')}`;
      } else if (heatPeriod === 'mesiac') {
        periodKey = `${iso.rok}-${String(iso.mesiac).padStart(2, '0')}`;
      } else if (heatPeriod === 'kvartal') {
        const q = Math.ceil(iso.mesiac / 3);
        periodKey = `${iso.rok}-Q${q}`;
      } else {
        periodKey = `${iso.rok}`;
      }
      if (!byOddPeriod[iso.oddelenie]) byOddPeriod[iso.oddelenie] = {};
      byOddPeriod[iso.oddelenie][periodKey] = (byOddPeriod[iso.oddelenie][periodKey] || 0) + 1;
    }

    // Get all periods sorted
    const allPeriods = Array.from(new Set(Object.values(byOddPeriod).flatMap(p => Object.keys(p)))).sort();
    const recentPeriods = allPeriods.slice(-12); // show last 12 periods

    return {
      oddelenia: Object.keys(byOddPeriod).sort(),
      periods: recentPeriods,
      data: byOddPeriod,
    };
  }, [base, heatPeriod]);

  // ── PANDÉMIA data ──────────────────────────────────────────────────────
  const pandemiaPeriods = ['pred', 'pocas', 'po'] as const;
  const pandemiaData = useMemo(() => {
    return pandemiaPeriods.map(period => {
      const isos = base.filter(iso => {
        const p = getPandemicPeriod(iso.datumOdberu);
        return p === period;
      }).filter(iso => {
        if (filterPatogen === 'aureus' && iso.patogen.indexOf('aureus') < 0) return false;
        if (filterPatogen === 'epidermidis' && iso.patogen.indexOf('epidermidis') < 0) return false;
        if (filterMdr === 'mdr' && !iso.isMdr) return false;
        if (filterMdr === 'nonmdr' && iso.isMdr) return false;
        if (pandemiaOddelenie && iso.oddelenie !== pandemiaOddelenie) return false;
        return true;
      });
      const s = computeStats(isos);
      const topOdd = getTopN(isos.map(i => i.oddelenie), 6);
      const topMat = getTopN(isos.map(i => i.material), 6);
      return { period, label: PANDEMIC_PERIODS[period].label, stats: s, topOdd, topMat, total: isos.length };
    });
  }, [base, filterPatogen, filterMdr, pandemiaOddelenie]);

  const pandemiaTrendData = useMemo(() => {
    const byYear: Record<number, { mdr: number; nonmdr: number; total: number }> = {};
    for (const iso of base) {
      const y = iso.rok;
      if (!byYear[y]) byYear[y] = { mdr: 0, nonmdr: 0, total: 0 };
      byYear[y].total++;
      if (iso.isMdr) byYear[y].mdr++; else byYear[y].nonmdr++;
    }
    return Object.keys(byYear).map(Number).sort((a, b) => a - b).map(y => ({
      rok: y.toString(), mdr: byYear[y].mdr, nonmdr: byYear[y].nonmdr, total: byYear[y].total,
      mdrPct: pct(byYear[y].mdr, byYear[y].total),
    }));
  }, [base]);

  // ── EMPTY STATE ────────────────────────────────────────────────────────
  if (isolates.length === 0) {
    return (
      <div style={{ minHeight: '100vh', background: 'linear-gradient(135deg, #0a1628 0%, #0f3460 60%, #1a5fa8 100%)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ textAlign: 'center', padding: '2rem' }}>
          <Image src="/logo.png" alt="StaphySearch" width={320} height={86} style={{ height: 70, width: 'auto', margin: '0 auto 2rem' }} />
          <div style={{ background: 'rgba(255,255,255,0.06)', borderRadius: 20, border: '1px solid rgba(0,184,150,0.2)', padding: '2.5rem 3rem', backdropFilter: 'blur(20px)', maxWidth: 480 }}>
            <p style={{ color: 'rgba(255,255,255,0.5)', fontSize: 13, marginBottom: '1.5rem', textTransform: 'uppercase', letterSpacing: '0.08em' }}>Virtuálna Nemocnica · Digitálne Dvojča</p>
            <p style={{ color: 'rgba(255,255,255,0.7)', fontSize: 14, marginBottom: '2rem', lineHeight: 1.7 }}>Nahraj export mikrobiologických dát (XLSX alebo CSV).</p>
            <label style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 8, background: 'linear-gradient(135deg, #1a5fa8, #00b896)', color: '#fff', padding: '10px 24px', borderRadius: 12, fontSize: 14, fontWeight: 500 }}>
              {loading ? <RefreshCw size={16} style={{ animation: 'spin 1s linear infinite' }} /> : <Upload size={16} />}
              {loading ? (progress || 'Spracovávam...') : 'Nahrať dáta (.xlsx / .csv)'}
              <input type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={handleUpload} disabled={loading} />
            </label>
            {error && <div style={{ marginTop: 16, padding: '10px 14px', background: 'rgba(220,38,38,0.15)', border: '1px solid rgba(220,38,38,0.3)', borderRadius: 10, color: '#fca5a5', fontSize: 13 }}>{error}</div>}
          </div>
        </div>
        <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }

  const PERIOD_COLORS = { pred: SS_BLUE, pocas: SS_ORANGE, po: SS_TEAL };

  return (
    <div style={{ minHeight: '100vh', background: '#f0f4f8' }}>
      <style>{`
        @keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
        .ss-chip{font-size:12px;font-weight:500;padding:5px 14px;border-radius:99px;border:1px solid rgba(26,95,168,0.15);color:#475569;cursor:pointer;background:#fff;transition:all 0.15s;}
        .ss-chip:hover{border-color:#00b896;color:#1a5fa8;}
        .ss-chip.active{background:linear-gradient(135deg,#1a5fa8,#00b896);color:#fff;border-color:transparent;}
        .heat-cell{border-radius:8px;transition:transform 0.15s,box-shadow 0.15s;cursor:default;}
        .heat-cell:hover{transform:scale(1.05);box-shadow:0 4px 12px rgba(0,0,0,0.12);}
      `}</style>

      {/* HEADER */}
      <header style={{ background: 'linear-gradient(135deg,#0a1628 0%,#0f3460 100%)', borderBottom: '1px solid rgba(0,184,150,0.15)', position: 'sticky', top: 0, zIndex: 100 }}>
        <div style={{ maxWidth: 1400, margin: '0 auto', padding: '0 1.5rem', display: 'flex', alignItems: 'center', justifyContent: 'space-between', height: 60 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 24 }}>
            <Image src="/logo.png" alt="StaphySearch" width={140} height={38} style={{ height: 34, width: 'auto' }} />
            <div style={{ display: 'flex', gap: 4 }}>
              {([['prehlad','Prehľad'],['heatmapa','Heatmapa'],['pandemia','Pandémia']] as [ActiveTab, string][]).map(([tab, label]) => (
                <button key={tab} onClick={() => setActiveTab(tab)} style={{ fontSize: 12, fontWeight: 500, padding: '5px 14px', borderRadius: 8, border: '1px solid transparent', cursor: 'pointer', background: activeTab === tab ? 'rgba(0,184,150,0.2)' : 'transparent', color: activeTab === tab ? '#fff' : 'rgba(255,255,255,0.55)', borderColor: activeTab === tab ? 'rgba(0,184,150,0.4)' : 'transparent', transition: 'all 0.15s' }}>{label}</button>
              ))}
              <a href="/chapter2" style={{ fontSize: 12, fontWeight: 500, padding: '5px 14px', borderRadius: 8, border: '1px solid rgba(255,255,255,0.12)', cursor: 'pointer', background: 'transparent', color: 'rgba(255,255,255,0.4)', textDecoration: 'none', display: 'flex', alignItems: 'center', gap: 4, marginLeft: 8 }}>Vrstva 2 →</a>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.3)', fontFamily: 'monospace' }}>{allStats.total} izolátov · {fileName}</span>
            <label style={{ cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 6, background: 'rgba(255,255,255,0.08)', color: 'rgba(255,255,255,0.7)', padding: '5px 12px', borderRadius: 8, fontSize: 12, border: '1px solid rgba(255,255,255,0.12)' }}>
              {loading ? <RefreshCw size={12} style={{ animation: 'spin 1s linear infinite' }} /> : <Upload size={12} />} Nahrať
              <input type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={handleUpload} disabled={loading} />
            </label>
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1400, margin: '0 auto', padding: '1.25rem 1.5rem' }}>

        {/* ══════════════════════════════════════════════════════════════
            TAB 1 — PREHĽAD
        ══════════════════════════════════════════════════════════════ */}
        {activeTab === 'prehlad' && (
          <div>
            {/* Global filters */}
            <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '0.9rem 1.25rem', marginBottom: '1rem', display: 'flex', flexWrap: 'wrap', gap: 16, alignItems: 'center' }}>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>MDR</span>
                {(['all','mdr','nonmdr'] as FilterMdr[]).map(v => (
                  <button key={v} onClick={() => setFilterMdr(v)} className={`ss-chip ${filterMdr === v ? 'active' : ''}`}>{v === 'all' ? 'Všetky' : v === 'mdr' ? 'MDR' : 'Non-MDR'}</button>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Patogén</span>
                {(['all','aureus','epidermidis'] as FilterPatogen[]).map(v => (
                  <button key={v} onClick={() => setFilterPatogen(v)} className={`ss-chip ${filterPatogen === v ? 'active' : ''}`}>{v === 'all' ? 'Všetky' : `S. ${v}`}</button>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Rok</span>
                <button onClick={() => setFilterRok(null)} className={`ss-chip ${filterRok === null ? 'active' : ''}`}>Všetky</button>
                {availableRoky.map(r => (
                  <button key={r} onClick={() => setFilterRok(filterRok === r ? null : r)} className={`ss-chip ${filterRok === r ? 'active' : ''}`}>{r}</button>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Skupina</span>
                <button onClick={() => setFilterGroup(null)} className={`ss-chip ${filterGroup === null ? 'active' : ''}`}>Všetky</button>
                {WARD_GROUPS.map(g => (
                  <button key={g.key} onClick={() => { setFilterGroup(filterGroup === g.key ? null : g.key); setFilterOddelenie(null); }} className={`ss-chip ${filterGroup === g.key ? 'active' : ''}`}>{g.label}</button>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Odd.</span>
                <select value={filterOddelenie || ''} onChange={e => setFilterOddelenie(e.target.value || null)} style={{ fontSize: 12, padding: '5px 10px', borderRadius: 8, border: '1px solid rgba(26,95,168,0.15)', color: '#475569', background: '#fff' }}>
                  <option value="">— všetky —</option>
                  {availableOddelenia.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
                {filterOddelenie && <button onClick={() => setFilterOddelenie(null)} style={{ fontSize: 11, color: SS_RED, cursor: 'pointer', background: 'none', border: 'none' }}>✕</button>}
              </div>
            </div>

            {/* Stat cards */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, minmax(0,1fr))', gap: 10, marginBottom: '1rem' }}>
              {[
                { val: filtered.length.toLocaleString(), lbl: 'Izolátov', sub: `z ${allStats.total} celkovo`, color: SS_BLUE },
                { val: `${stats.mdrRate}%`, lbl: 'MDR rate', sub: `${stats.mdr} MDR kmeňov`, color: SS_RED },
                { val: stats.mrsa.toString(), lbl: 'MRSA / MRCoNS', sub: `${pct(stats.mrsa, stats.total)}% z filtr.`, color: SS_ORANGE },
                { val: stats.aureus.toString(), lbl: 'S. aureus', sub: `${pct(stats.aureus, stats.total)}% izolátov`, color: SS_BLUE },
                { val: stats.epidermidis.toString(), lbl: 'S. epidermidis', sub: `${pct(stats.epidermidis, stats.total)}% izolátov`, color: SS_TEAL },
              ].map((s, i) => (
                <div key={i} style={{ background: '#fff', borderRadius: 13, padding: '1rem 1.1rem', border: '1px solid rgba(26,95,168,0.08)', boxShadow: '0 1px 3px rgba(10,22,40,0.04)' }}>
                  <div style={{ fontSize: '2rem', fontWeight: 700, lineHeight: 1, color: s.color, fontFamily: 'DM Mono, monospace' }}>{s.val}</div>
                  <div style={{ fontSize: 10.5, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', color: '#64748b', marginTop: 5 }}>{s.lbl}</div>
                  <div style={{ fontSize: 10.5, color: '#94a3b8', marginTop: 2 }}>{s.sub}</div>
                </div>
              ))}
            </div>

            {/* Dynamic quarterly chart + slider */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.08)', padding: '1.25rem 1.5rem', marginBottom: '1rem' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', margin: 0 }}>Kvartálny vývoj — MDR kmeňov</p>
                <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                  <button onClick={() => setSliderOffset(Math.min(maxOffset, sliderOffset + 1))} disabled={sliderOffset >= maxOffset} style={{ padding: '4px 8px', borderRadius: 7, border: '1px solid #e2e8f0', background: '#fff', cursor: sliderOffset >= maxOffset ? 'not-allowed' : 'pointer', color: sliderOffset >= maxOffset ? '#cbd5e1' : '#475569', display: 'flex', alignItems: 'center' }}>
                    <ChevronLeft size={14} />
                  </button>
                  <span style={{ fontSize: 11, color: '#64748b' }}>{visibleQuarters[0]?.label} – {visibleQuarters[visibleQuarters.length - 1]?.label}</span>
                  <button onClick={() => setSliderOffset(Math.max(0, sliderOffset - 1))} disabled={sliderOffset <= 0} style={{ padding: '4px 8px', borderRadius: 7, border: '1px solid #e2e8f0', background: '#fff', cursor: sliderOffset <= 0 ? 'not-allowed' : 'pointer', color: sliderOffset <= 0 ? '#cbd5e1' : '#475569', display: 'flex', alignItems: 'center' }}>
                    <ChevronRight size={14} />
                  </button>
                </div>
              </div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={quarterChartData} margin={{ top: 5, right: 5, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis dataKey="label" tick={{ fontSize: 11 }} />
                  <YAxis tick={{ fontSize: 11 }} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar dataKey="mdr" name="MDR" stackId="a" fill={SS_RED} />
                  <Bar dataKey="nonmdr" name="Non-MDR" stackId="a" fill={SS_TEAL} radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* Last 2Q sections */}
            <div style={{ background: 'linear-gradient(135deg, #0a1628, #0f3460)', borderRadius: 12, padding: '8px 16px', marginBottom: '1rem', display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.5)', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em' }}>Posledné 2 kvartály</span>
              <span style={{ fontSize: 12, color: SS_TEAL, fontWeight: 600 }}>{last2Q.map(q => q.label).join(' · ')}</span>
              <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.4)', marginLeft: 'auto' }}>{last2QIsos.length} izolátov</span>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '1rem', marginBottom: '1rem' }}>
              {/* Top oddelenia - počet */}
              <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '1.1rem 1.25rem' }}>
                <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', margin: '0 0 12px' }}>Oddelenia — počet kmeňov</p>
                {topOddLast2Q.map(({ name, count }, i) => (
                  <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                    <span style={{ fontSize: 10, color: '#cbd5e1', width: 14, textAlign: 'right', fontFamily: 'monospace' }}>{i + 1}</span>
                    <span style={{ flex: 1, fontSize: 11.5, color: '#334155', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={name}>{name}</span>
                    <div style={{ width: 60, background: '#f1f5f9', borderRadius: 3, height: 12, overflow: 'hidden' }}>
                      <div style={{ height: '100%', width: `${pct(count, last2QIsos.length)}%`, background: `linear-gradient(90deg,${SS_BLUE},${SS_TEAL})`, borderRadius: 3 }} />
                    </div>
                    <span style={{ fontSize: 11, fontFamily: 'monospace', color: '#64748b', width: 22, textAlign: 'right' }}>{count}</span>
                  </div>
                ))}
              </div>

              {/* Top oddelenia - MDR */}
              <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '1.1rem 1.25rem' }}>
                <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', margin: '0 0 12px' }}>Oddelenia — MDR kmene</p>
                {topOddMdrLast2Q.map(({ name, count }, i) => (
                  <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                    <span style={{ fontSize: 10, color: '#cbd5e1', width: 14, textAlign: 'right', fontFamily: 'monospace' }}>{i + 1}</span>
                    <span style={{ flex: 1, fontSize: 11.5, color: '#334155', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={name}>{name}</span>
                    <div style={{ width: 60, background: '#f1f5f9', borderRadius: 3, height: 12, overflow: 'hidden' }}>
                      <div style={{ height: '100%', width: `${pct(count, last2QIsos.filter(i => i.isMdr).length)}%`, background: SS_RED, borderRadius: 3 }} />
                    </div>
                    <span style={{ fontSize: 11, fontFamily: 'monospace', color: SS_RED, width: 22, textAlign: 'right' }}>{count}</span>
                  </div>
                ))}
              </div>

              {/* Top materiály */}
              <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '1.1rem 1.25rem' }}>
                <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', margin: '0 0 12px' }}>Najčastejšie materiály</p>
                {topMatLast2Q.map(({ name, count }, i) => (
                  <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 7 }}>
                    <span style={{ fontSize: 10, color: '#cbd5e1', width: 14, textAlign: 'right', fontFamily: 'monospace' }}>{i + 1}</span>
                    <span style={{ flex: 1, fontSize: 11.5, color: '#334155', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{name}</span>
                    <div style={{ width: 60, background: '#f1f5f9', borderRadius: 3, height: 12, overflow: 'hidden' }}>
                      <div style={{ height: '100%', width: `${pct(count, last2QIsos.length)}%`, background: `linear-gradient(90deg,${SS_TEAL},${SS_BLUE})`, borderRadius: 3 }} />
                    </div>
                    <span style={{ fontSize: 11, fontFamily: 'monospace', color: '#64748b', width: 22, textAlign: 'right' }}>{count}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ══════════════════════════════════════════════════════════════
            TAB 2 — HEATMAPA
        ══════════════════════════════════════════════════════════════ */}
        {activeTab === 'heatmapa' && (
          <div>
            {/* Period selector */}
            <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '0.9rem 1.25rem', marginBottom: '1rem', display: 'flex', gap: 12, alignItems: 'center' }}>
              <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Zobrazenie po</span>
              {([['tyzdne','Týždeň'],['mesiac','Mesiac'],['kvartal','Kvartál'],['rok','Rok']] as [HeatPeriod, string][]).map(([v, l]) => (
                <button key={v} onClick={() => setHeatPeriod(v)} className={`ss-chip ${heatPeriod === v ? 'active' : ''}`}>{l}</button>
              ))}
              <span style={{ fontSize: 11, color: '#94a3b8', marginLeft: 'auto' }}>Len lôžkové oddelenia · Farba = počet izolátov</span>
            </div>

            {/* Heatmapa grid */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.08)', padding: '1.5rem', overflowX: 'auto' }}>
              <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', marginBottom: 16 }}>
                Heatmapa izolátov — {heatPeriod === 'tyzdne' ? 'po týždňoch' : heatPeriod === 'mesiac' ? 'po mesiacoch' : heatPeriod === 'kvartal' ? 'po kvartáloch' : 'po rokoch'} · posledných 12 období
              </p>
              {heatmapData.oddelenia.length === 0 ? (
                <p style={{ color: '#94a3b8', fontSize: 13 }}>Žiadne lôžkové oddelenia nenájdené v dátach.</p>
              ) : (
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ borderCollapse: 'separate', borderSpacing: 4, minWidth: 600 }}>
                    <thead>
                      <tr>
                        <th style={{ fontSize: 10, color: '#94a3b8', textAlign: 'left', padding: '4px 8px', fontWeight: 600, minWidth: 160 }}>Oddelenie</th>
                        {heatmapData.periods.map(p => (
                          <th key={p} style={{ fontSize: 9, color: '#94a3b8', textAlign: 'center', padding: '4px 2px', fontWeight: 500, minWidth: 44 }}>{p.slice(-5)}</th>
                        ))}
                        <th style={{ fontSize: 10, color: '#94a3b8', padding: '4px 8px', fontWeight: 600 }}>Celkom</th>
                      </tr>
                    </thead>
                    <tbody>
                      {heatmapData.oddelenia.map(od => {
                        const rowData = heatmapData.data[od];
                        const total = Object.values(rowData).reduce((a, b) => a + b, 0);
                        const maxVal = Math.max(...heatmapData.periods.map(p => rowData[p] || 0), 1);
                        return (
                          <tr key={od}>
                            <td style={{ fontSize: 11, color: '#334155', padding: '3px 8px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 160 }} title={od}>{od}</td>
                            {heatmapData.periods.map(p => {
                              const val = rowData[p] || 0;
                              const ratio = maxVal > 0 ? val / maxVal : 0;
                              const bg = val === 0 ? '#f8fafc' : ratio < 0.15 ? '#dcfce7' : ratio < 0.30 ? '#bbf7d0' : ratio < 0.45 ? '#fef9c3' : ratio < 0.60 ? '#fed7aa' : ratio < 0.75 ? '#fca5a5' : '#ef4444';
                              const color = ratio > 0.65 ? '#7f1d1d' : ratio > 0.35 ? '#78350f' : '#166534';
                              // Compare with previous period
                              const pIdx = heatmapData.periods.indexOf(p);
                              const prevVal = pIdx > 0 ? (rowData[heatmapData.periods[pIdx - 1]] || 0) : null;
                              const diff = prevVal !== null ? val - prevVal : null;
                              return (
                                <td key={p} style={{ padding: 2, textAlign: 'center' }}>
                                  <div className="heat-cell" style={{ background: bg, color, fontSize: 11, fontWeight: 600, fontFamily: 'monospace', padding: '6px 4px', position: 'relative', minWidth: 40 }} title={`${od} · ${p}: ${val} izolátov${diff !== null ? ` (${diff > 0 ? '+' : ''}${diff} vs predch.)` : ''}`}>
                                    {val || '·'}
                                    {diff !== null && val > 0 && (
                                      <span style={{ position: 'absolute', top: 1, right: 2, fontSize: 7, color: diff > 0 ? '#ef4444' : diff < 0 ? '#22c55e' : '#94a3b8' }}>
                                        {diff > 0 ? '↑' : diff < 0 ? '↓' : ''}
                                      </span>
                                    )}
                                  </div>
                                </td>
                              );
                            })}
                            <td style={{ fontSize: 11, fontFamily: 'monospace', fontWeight: 700, color: SS_BLUE, padding: '3px 8px', textAlign: 'right' }}>{total}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}

              {/* Legend */}
              <div style={{ marginTop: 16, display: 'flex', gap: 10, alignItems: 'center', fontSize: 11, color: '#64748b', flexWrap: 'wrap' }}>
                <span style={{ fontWeight: 600 }}>Škála (relatívna):</span>
                {([['#dcfce7','Nízka'],['#fef9c3','Stredná'],['#fed7aa','Zvýšená'],['#fca5a5','Vysoká'],['#ef4444','Kritická']] as [string,string][]).map(([bg, label]) => (
                  <span key={label} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                    <span style={{ width: 20, height: 14, borderRadius: 4, background: bg, display: 'inline-block', border: '1px solid rgba(0,0,0,0.08)' }} />
                    <span style={{ fontSize: 10 }}>{label}</span>
                  </span>
                ))}
                <span style={{ fontSize: 10, color: '#94a3b8', marginLeft: 4 }}>· ↑↓ zmena vs predchádzajúce obdobie</span>
              </div>
            </div>
          </div>
        )}

        {/* ══════════════════════════════════════════════════════════════
            TAB 3 — PANDÉMIA
        ══════════════════════════════════════════════════════════════ */}
        {activeTab === 'pandemia' && (
          <div>
            {/* Filters */}
            <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '0.9rem 1.25rem', marginBottom: '1rem', display: 'flex', flexWrap: 'wrap', gap: 16, alignItems: 'center' }}>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>MDR</span>
                {(['all','mdr','nonmdr'] as FilterMdr[]).map(v => (
                  <button key={v} onClick={() => setFilterMdr(v)} className={`ss-chip ${filterMdr === v ? 'active' : ''}`}>{v === 'all' ? 'Všetky' : v === 'mdr' ? 'MDR' : 'Non-MDR'}</button>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Patogén</span>
                {(['all','aureus','epidermidis'] as FilterPatogen[]).map(v => (
                  <button key={v} onClick={() => setFilterPatogen(v)} className={`ss-chip ${filterPatogen === v ? 'active' : ''}`}>{v === 'all' ? 'Všetky' : `S. ${v}`}</button>
                ))}
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <span style={{ fontSize: 11, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em' }}>Oddelenie</span>
                <select value={pandemiaOddelenie || ''} onChange={e => setPandemiaOddelenie(e.target.value || null)} style={{ fontSize: 12, padding: '5px 10px', borderRadius: 8, border: '1px solid rgba(26,95,168,0.15)', color: '#475569', background: '#fff' }}>
                  <option value="">— všetky —</option>
                  {availableOddelenia.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
                {pandemiaOddelenie && <button onClick={() => setPandemiaOddelenie(null)} style={{ fontSize: 11, color: SS_RED, cursor: 'pointer', background: 'none', border: 'none' }}>✕</button>}
              </div>
            </div>

            {/* Trend graf */}
            <div style={{ background: '#fff', borderRadius: 16, border: '1px solid rgba(26,95,168,0.08)', padding: '1.25rem 1.5rem', marginBottom: '1rem' }}>
              <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', margin: '0 0 12px' }}>Ročný trend MDR kmeňov</p>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={pandemiaTrendData} margin={{ top: 5, right: 5, left: -10, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis dataKey="rok" tick={{ fontSize: 11 }} />
                  <YAxis tick={{ fontSize: 11 }} />
                  <Tooltip content={<CustomTooltip />} />
                  <Bar dataKey="mdr" name="MDR" stackId="a" fill={SS_RED} />
                  <Bar dataKey="nonmdr" name="Non-MDR" stackId="a" fill={SS_TEAL} radius={[4, 4, 0, 0]} />
                  <ReferenceLine x="2020" stroke={SS_ORANGE} strokeDasharray="5 3" label={{ value: 'Pandémia ▶', position: 'insideTopLeft', fontSize: 9, fill: SS_ORANGE }} />
                  <ReferenceLine x="2023" stroke={SS_TEAL} strokeDasharray="5 3" label={{ value: '◀ Koniec', position: 'insideTopRight', fontSize: 9, fill: SS_TEAL }} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            {/* 3 stĺpce: pred / počas / po */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '1rem' }}>
              {pandemiaData.map(pd => (
                <div key={pd.period} style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                  {/* Header card */}
                  <div style={{ background: `linear-gradient(135deg, ${PERIOD_COLORS[pd.period as keyof typeof PERIOD_COLORS]}20, ${PERIOD_COLORS[pd.period as keyof typeof PERIOD_COLORS]}08)`, borderRadius: 14, border: `1px solid ${PERIOD_COLORS[pd.period as keyof typeof PERIOD_COLORS]}30`, padding: '1rem 1.25rem' }}>
                    <div style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.07em', color: PERIOD_COLORS[pd.period as keyof typeof PERIOD_COLORS], marginBottom: 8 }}>{pd.label}</div>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                      <div>
                        <div style={{ fontSize: '1.6rem', fontWeight: 700, fontFamily: 'monospace', color: '#0a1628' }}>{pd.total}</div>
                        <div style={{ fontSize: 10, color: '#64748b' }}>izolátov</div>
                      </div>
                      <div>
                        <div style={{ fontSize: '1.6rem', fontWeight: 700, fontFamily: 'monospace', color: SS_RED }}>{pd.stats.mdrRate}%</div>
                        <div style={{ fontSize: 10, color: '#64748b' }}>MDR rate</div>
                      </div>
                      <div>
                        <div style={{ fontSize: '1.2rem', fontWeight: 700, fontFamily: 'monospace', color: SS_ORANGE }}>{pd.stats.mrsa}</div>
                        <div style={{ fontSize: 10, color: '#64748b' }}>MRSA/MRCoNS</div>
                      </div>
                      <div>
                        <div style={{ fontSize: '1.2rem', fontWeight: 700, fontFamily: 'monospace', color: SS_BLUE }}>{pd.stats.aureus}</div>
                        <div style={{ fontSize: 10, color: '#64748b' }}>S. aureus</div>
                      </div>
                    </div>
                  </div>

                  {/* Oddelenia */}
                  <div style={{ background: '#fff', borderRadius: 13, border: '1px solid rgba(26,95,168,0.08)', padding: '1rem 1.1rem' }}>
                    <p style={{ fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', margin: '0 0 10px' }}>Top oddelenia</p>
                    {pd.topOdd.map(({ name, count }) => (
                      <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
                        <span style={{ flex: 1, fontSize: 11, color: '#334155', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={name}>{name}</span>
                        <div style={{ width: 50, background: '#f1f5f9', borderRadius: 3, height: 10, overflow: 'hidden' }}>
                          <div style={{ height: '100%', width: `${pct(count, pd.total)}%`, background: PERIOD_COLORS[pd.period as keyof typeof PERIOD_COLORS], borderRadius: 3 }} />
                        </div>
                        <span style={{ fontSize: 10, fontFamily: 'monospace', color: '#64748b', width: 20, textAlign: 'right' }}>{count}</span>
                      </div>
                    ))}
                  </div>

                  {/* Materiály */}
                  <div style={{ background: '#fff', borderRadius: 13, border: '1px solid rgba(26,95,168,0.08)', padding: '1rem 1.1rem' }}>
                    <p style={{ fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.07em', color: '#94a3b8', margin: '0 0 10px' }}>Top materiály</p>
                    {pd.topMat.map(({ name, count }) => (
                      <div key={name} style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 }}>
                        <span style={{ flex: 1, fontSize: 11, color: '#334155', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{name}</span>
                        <div style={{ width: 50, background: '#f1f5f9', borderRadius: 3, height: 10, overflow: 'hidden' }}>
                          <div style={{ height: '100%', width: `${pct(count, pd.total)}%`, background: PERIOD_COLORS[pd.period as keyof typeof PERIOD_COLORS], borderRadius: 3 }} />
                        </div>
                        <span style={{ fontSize: 10, fontFamily: 'monospace', color: '#64748b', width: 20, textAlign: 'right' }}>{count}</span>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>

            <p style={{ fontSize: 11, color: '#94a3b8', marginTop: 12, textAlign: 'center' }}>
              Pred pandémiou: do 11.3.2020 · Počas pandémie: 12.3.2020 – 15.9.2023 (mimoriadna situácia SR) · Po pandémii: od 16.9.2023
            </p>
          </div>
        )}

        <footer style={{ textAlign: 'center', fontSize: 10.5, color: '#94a3b8', padding: '1.5rem 0', marginTop: '0.5rem' }}>
          StaphySearch · MDR: Magiorakos et al. 2012 (ECDC/CDC) · Ambulancie a Ružinov sú vylúčené · Rok 2016 vylúčený
        </footer>
      </main>
    </div>
  );
}
