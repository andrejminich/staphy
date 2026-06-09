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
import { Isolate, computeStats, getTopN, MONTH_NAMES, processRawData, RawRow, EARS_NET, HOSP_TOTAL, HOSP_DATA, getHospCount, ratePer1000, getPandemicPeriod, PANDEMIC_PERIODS } from '@/lib/mdr';

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

// Lôžkové oddelenia — match against HOSP_DATA keywords
function isLôžkové(oddelenie: string): boolean {
  const lower = oddelenie.toLowerCase();
  return HOSP_DATA.some(d => d.keywords.some(k => lower.includes(k)));
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

type ActiveTab = 'prehlad' | 'heatmapa' | 'pandemia' | 'trendy';
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

  // Auto-adjust slider when year filter changes
  useEffect(() => {
    if (filterRok !== null) {
      // Find index of last quarter of selected year
      const lastYearQIdx = allQuarters.findLastIndex ? 
        allQuarters.findLastIndex(q => q.rok === filterRok) :
        [...allQuarters].map((q,i) => q.rok === filterRok ? i : -1).filter(i => i >= 0).pop() ?? -1;
      if (lastYearQIdx >= 0) {
        // Calculate offset so the selected year's quarters are centered/visible
        const endIdx = lastYearQIdx + 1;
        const newOffset = Math.max(0, allQuarters.length - endIdx);
        setSliderOffset(Math.min(newOffset, Math.max(0, allQuarters.length - SLIDER_WINDOW)));
      }
    } else {
      setSliderOffset(0); // reset to latest when "Všetky"
    }
  }, [filterRok, allQuarters]);

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
  // When year filter is active, use last 2 quarters of that year; otherwise last 2 of all data
  const last2Q = useMemo(() => {
    if (filterRok !== null) {
      const yearQs = allQuarters.filter(q => q.rok === filterRok);
      return yearQs.slice(-2);
    }
    return allQuarters.slice(-2);
  }, [allQuarters, filterRok]);
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
    type PeriodKey = string;
    // count isolates per oddelenie per period
    const byOddPeriod: Record<string, Record<PeriodKey, number>> = {};
    // track rok per period key for hosp lookup
    const periodRok: Record<PeriodKey, number> = {};

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
      periodRok[periodKey] = iso.rok;
      if (!byOddPeriod[iso.oddelenie]) byOddPeriod[iso.oddelenie] = {};
      byOddPeriod[iso.oddelenie][periodKey] = (byOddPeriod[iso.oddelenie][periodKey] || 0) + 1;
    }

    const allPeriods = Array.from(new Set(Object.values(byOddPeriod).flatMap(p => Object.keys(p)))).sort();
    const recentPeriods = allPeriods.slice(-12);

    // For rok-level heatmap, compute rate per 1000 hosp per oddelenie
    // For sub-year periods, show absolute counts (no good denominator)
    const useRate = heatPeriod === 'rok';

    // Build rate data if useRate
    const rateData: Record<string, Record<PeriodKey, { val: number; isApprox: boolean }>> = {};
    if (useRate) {
      for (const od of Object.keys(byOddPeriod)) {
        rateData[od] = {};
        for (const p of recentPeriods) {
          const rok = periodRok[p] || Number(p);
          const count = byOddPeriod[od][p] || 0;
          const hosp = getHospCount(od, rok);
          rateData[od][p] = { val: ratePer1000(count, hosp.count), isApprox: hosp.isApprox };
        }
      }
    }

    return {
      oddelenia: Object.keys(byOddPeriod).sort(),
      periods: recentPeriods,
      data: byOddPeriod,
      rateData,
      useRate,
      periodRok,
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
              {([['prehlad','Prehľad'],['heatmapa','Heatmapa'],['pandemia','Pandémia'],['trendy','Trendy']] as [ActiveTab, string][]).map(([tab, label]) => (
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
              {(() => {
                const availRoky = Array.from(new Set(base.map(i => i.rok)));
                const lastRokWithHosp = Math.max(...Object.keys(HOSP_TOTAL).map(Number).filter(y => availRoky.includes(y)));
                const hospRok = filterRok && HOSP_TOTAL[filterRok] ? filterRok : lastRokWithHosp;
                const hospCount = HOSP_TOTAL[hospRok] || 0;
                const mrsaForRok = filtered.filter(i => i.isMrsa && (!filterRok || i.rok === filterRok)).length;
                const mrsaRate = hospCount > 0 ? ratePer1000(mrsaForRok, hospCount) : null;
                return [
                  { val: filtered.length.toLocaleString(), lbl: 'Izolátov', sub: `z ${allStats.total} celkovo`, color: SS_BLUE },
                  { val: `${stats.mdrRate}%`, lbl: 'MDR rate', sub: `${stats.mdr} MDR kmeňov`, color: SS_RED },
                  { val: stats.mrsa.toString(), lbl: 'MRSA / MRCoNS', sub: `${pct(stats.mrsa, stats.total)}% z filtr.`, color: SS_ORANGE },
                  { val: mrsaRate !== null ? `${mrsaRate}‰` : '—', lbl: 'MRSA / 1000 hosp.', sub: filterRok ? `Rok ${filterRok}` : `Rok ${hospRok}`, color: SS_TEAL },
                  { val: stats.aureus.toString(), lbl: 'S. aureus', sub: `${pct(stats.aureus, stats.total)}% izolátov`, color: SS_BLUE },
                ];
              })().map((s, i) => (
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
              <span style={{ fontSize: 11, color: '#94a3b8', marginLeft: 'auto' }}>Lôžkové oddelenia · Rok: ‰ per 1000 hosp. · Ostatné: absolútny počet</span>
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
                        // Use rate per 1000 for yearly view, absolute for sub-year
                        const displayVals = heatmapData.periods.map(p => {
                          if (heatmapData.useRate && heatmapData.rateData[od]?.[p]) {
                            return heatmapData.rateData[od][p].val;
                          }
                          return rowData[p] || 0;
                        });
                        const maxVal = Math.max(...displayVals, 1);
                        return (
                          <tr key={od}>
                            <td style={{ fontSize: 11, color: '#334155', padding: '3px 8px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 160 }} title={od}>{od}</td>
                            {heatmapData.periods.map((p, pIdx) => {
                              const absVal = rowData[p] || 0;
                              const displayVal = displayVals[pIdx];
                              const ratio = maxVal > 0 ? displayVal / maxVal : 0;
                              const bg = displayVal === 0 ? '#f8fafc' : ratio < 0.15 ? '#dcfce7' : ratio < 0.30 ? '#bbf7d0' : ratio < 0.45 ? '#fef9c3' : ratio < 0.60 ? '#fed7aa' : ratio < 0.75 ? '#fca5a5' : '#ef4444';
                              const color = ratio > 0.65 ? '#7f1d1d' : ratio > 0.35 ? '#78350f' : '#166534';
                              const prevDisplayVal = pIdx > 0 ? displayVals[pIdx - 1] : null;
                              const diff = prevDisplayVal !== null ? displayVal - prevDisplayVal : null;
                              const isApprox = heatmapData.useRate && heatmapData.rateData[od]?.[p]?.isApprox;
                              const label = heatmapData.useRate ? `${displayVal}‰` : `${absVal}`;
                              return (
                                <td key={p} style={{ padding: 2, textAlign: 'center' }}>
                                  <div className="heat-cell" style={{ background: bg, color, fontSize: 10, fontWeight: 600, fontFamily: 'monospace', padding: '6px 4px', position: 'relative', minWidth: 44 }} title={`${od} · ${p}: ${absVal} izolátov${heatmapData.useRate ? ` · ${displayVal}‰ per 1000 hosp.${isApprox ? ' (aprox.)' : ''}` : ''}${diff !== null ? ` · ${diff > 0 ? '+' : ''}${diff.toFixed(1)} vs predch.` : ''}`}>
                                    {displayVal > 0 ? label : '·'}
                                    {isApprox && displayVal > 0 && <span style={{ position: 'absolute', top: 1, left: 2, fontSize: 7, color: 'inherit', opacity: 0.6 }}>~</span>}
                                    {diff !== null && displayVal > 0 && (
                                      <span style={{ position: 'absolute', top: 1, right: 2, fontSize: 7, color: diff > 0 ? '#ef4444' : diff < 0 ? '#22c55e' : '#94a3b8' }}>
                                        {diff > 0 ? '↑' : diff < 0 ? '↓' : ''}
                                      </span>
                                    )}
                                  </div>
                                </td>
                              );
                            })}
                            <td style={{ fontSize: 11, fontFamily: 'monospace', fontWeight: 700, color: SS_BLUE, padding: '3px 8px', textAlign: 'right' }}>{heatmapData.useRate ? `${Math.round(displayVals.reduce((a,b)=>a+b,0)/Math.max(heatmapData.periods.filter(p=>rowData[p]>0).length,1)*10)/10}‰` : total}</td>
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

        {/* ══ TAB: TRENDY ══ */}
        {activeTab === 'trendy' && (
          <TrendyTab base={base} hospData={HOSP_DATA} hospTotal={HOSP_TOTAL} ratePer1000={ratePer1000} />
        )}

        <footer style={{ textAlign: 'center', fontSize: 10.5, color: '#94a3b8', padding: '1.5rem 0', marginTop: '0.5rem' }}>
          StaphySearch · MDR: Magiorakos et al. 2012 (ECDC/CDC) · Ambulancie a Ružinov sú vylúčené · Rok 2016 vylúčený
        </footer>
      </main>
    </div>
  );
}

// ── Client-side Poisson regression (IRLS) ────────────────────────────────
function fitPoisson(oddelenie: string, data: {rok:number;count:number;hosp:number;rate:number}[], metric: string): TrendResult {
  const n = data.length;
  if (n < 4 || data.every(d => d.count === 0)) {
    return { oddelenie, data, irr:null, ci_low:null, ci_high:null, p_value:null, model:'insufficient', overdispersion:null, significant:false, direction:'ns', fitted:[] };
  }

  // Use count as outcome, log(hosp) as offset, rok as predictor
  const y = data.map(d => d.count);
  const offset = data.map(d => Math.log(Math.max(d.hosp, 1)));
  const x = data.map(d => d.rok);
  const xMean = x.reduce((a,b) => a+b, 0) / n;
  const xc = x.map(v => v - xMean); // center for numerical stability

  // IRLS for Poisson GLM: log(mu) = a + b*xc + offset
  let a = Math.log(y.reduce((s,v) => s+v, 0) / y.length + 0.5);
  let b = 0;

  for (let iter = 0; iter < 50; iter++) {
    const mu = xc.map((_,i) => Math.exp(a + b * xc[i] + offset[i]));
    // Score equations
    let sa = 0, sb = 0, Iaa = 0, Iab = 0, Ibb = 0;
    for (let i = 0; i < n; i++) {
      const r = y[i] - mu[i];
      sa += r;
      sb += r * xc[i];
      Iaa += mu[i];
      Iab += mu[i] * xc[i];
      Ibb += mu[i] * xc[i] * xc[i];
    }
    const det = Iaa * Ibb - Iab * Iab;
    if (Math.abs(det) < 1e-12) break;
    const da = (Ibb * sa - Iab * sb) / det;
    const db = (Iaa * sb - Iab * sa) / det;
    a += da; b += db;
    if (Math.abs(da) < 1e-8 && Math.abs(db) < 1e-8) break;
  }

  // Variance (Fisher information inverse)
  const mu = xc.map((_,i) => Math.exp(a + b * xc[i] + offset[i]));
  let Iaa = 0, Iab = 0, Ibb = 0;
  for (let i = 0; i < n; i++) {
    Iaa += mu[i]; Iab += mu[i] * xc[i]; Ibb += mu[i] * xc[i] * xc[i];
  }
  const det = Iaa * Ibb - Iab * Iab;
  const varB = Math.abs(det) > 1e-12 ? Iaa / det : 1;

  // Deviance and overdispersion
  let deviance = 0;
  for (let i = 0; i < n; i++) {
    if (y[i] > 0) deviance += 2 * (y[i] * Math.log(y[i] / mu[i]) - (y[i] - mu[i]));
    else deviance += 2 * mu[i];
  }
  const df = n - 2;
  const phi = deviance / df; // overdispersion parameter

  const useQP = phi > 1.5;
  const seB = Math.sqrt(varB * (useQP ? phi : 1));
  const irr = Math.exp(b);
  const z = b / seB;

  // Two-tailed p-value (normal approximation)
  const pval = 2 * (1 - normalCDF(Math.abs(z)));
  const z95 = 1.96;
  const ci_low = Math.exp(b - z95 * seB);
  const ci_high = Math.exp(b + z95 * seB);

  // Fitted values with CI for original (uncentered) years
  const fitted = data.map((d,i) => {
    const xci = xc[i];
    const logFit = a + b * xci + offset[i];
    const seFit = Math.sqrt(varB * (useQP ? phi : 1) * xci * xci + 1/n); // approx
    const fRate = Math.exp(logFit) / Math.max(d.hosp, 1) * 1000;
    const fCount = Math.exp(logFit);
    const val = metric === 'rate' ? fRate : fCount;
    const cil = metric === 'rate' ? Math.exp(logFit - z95*seFit)/Math.max(d.hosp,1)*1000 : Math.exp(logFit - z95*seFit);
    const cih = metric === 'rate' ? Math.exp(logFit + z95*seFit)/Math.max(d.hosp,1)*1000 : Math.exp(logFit + z95*seFit);
    return { rok: d.rok, fitted: Math.round(val*100)/100, ci_low: Math.round(cil*100)/100, ci_high: Math.round(cih*100)/100 };
  });

  const sig = pval < 0.05;
  return {
    oddelenie, data,
    irr: Math.round(irr * 1000) / 1000,
    ci_low: Math.round(ci_low * 1000) / 1000,
    ci_high: Math.round(ci_high * 1000) / 1000,
    p_value: Math.round(pval * 10000) / 10000,
    model: useQP ? 'quasi-poisson' : 'poisson',
    overdispersion: Math.round(phi * 100) / 100,
    significant: sig,
    direction: sig ? (b > 0 ? 'up' : 'down') : 'ns',
    fitted,
  };
}

function normalCDF(x: number): number {
  const t = 1 / (1 + 0.2316419 * x);
  const poly = t * (0.319381530 + t * (-0.356563782 + t * (1.781477937 + t * (-1.821255978 + t * 1.330274429))));
  return 1 - 0.3989422803 * Math.exp(-x * x / 2) * poly;
}


// ══════════════════════════════════════════════════════════════════════════
// TRENDY TAB — Poisson regression, CI, export
// ══════════════════════════════════════════════════════════════════════════

interface TrendResult {
  oddelenie: string;
  data: { rok: number; count: number; hosp: number; rate: number }[];
  irr: number | null;
  ci_low: number | null;
  ci_high: number | null;
  p_value: number | null;
  model: 'poisson' | 'quasi-poisson' | 'insufficient';
  overdispersion: number | null;
  significant: boolean;
  direction: 'up' | 'down' | 'ns';
  fitted: { rok: number; fitted: number; ci_low: number; ci_high: number }[];
}

function TrendyTab({ base, hospData, hospTotal, ratePer1000 }: {
  base: Isolate[];
  hospData: { keywords: string[]; data: Record<number, number> }[];
  hospTotal: Record<number, number>;
  ratePer1000: (n: number, d: number) => number;
}) {
  const [selectedOddelenia, setSelectedOddelenia] = useState<string[]>([]);
  const [metric, setMetric] = useState<'rate' | 'count'>('rate');
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState<TrendResult[]>([]);
  const [error, setError] = useState<string | null>(null);
  const chartRefs = useRef<Record<string, HTMLCanvasElement | null>>({});

  const availableOddelenia = useMemo(() => {
    const s = new Set(base.map(i => i.oddelenie));
    return Array.from(s).sort();
  }, [base]);

  const availableRoky = useMemo(() => {
    const s = new Set(base.map(i => i.rok));
    return Array.from(s).map(Number).sort((a, b) => a - b);
  }, [base]);

  // Build yearly data per oddelenie
  const buildOddData = (oddelenie: string) => {
    return availableRoky.map(rok => {
      const isos = base.filter(i => i.oddelenie === oddelenie && i.rok === rok);
      const hospMatch = hospData.find(d => d.keywords.some(k => oddelenie.toLowerCase().includes(k)));
      const hosp = (hospMatch?.data[rok]) || hospTotal[rok] || 0;
      const rate = hosp > 0 ? ratePer1000(isos.length, hosp) : 0;
      return { rok, count: isos.length, hosp, rate };
    }).filter(d => d.hosp > 0);
  };

  const runAnalysis = () => {
    if (!selectedOddelenia.length) return;
    setLoading(true); setError(null); setResults([]);

    const oddData = selectedOddelenia.map(od => ({
      oddelenie: od,
      data: buildOddData(od)
    }));

    try {
      // Client-side Poisson regression via IRLS (no API needed)
      const enriched: TrendResult[] = oddData.map(({ oddelenie, data }) => {
        return fitPoisson(oddelenie, data, metric);
      });
      setResults(enriched);
    } catch (e) {
      setError('Chyba pri analýze: ' + String(e));
    } finally {
      setLoading(false);
    }
  };

  // Draw chart on canvas
  const drawChart = (canvas: HTMLCanvasElement, result: TrendResult) => {
    const ctx = canvas.getContext('2d'); if (!ctx) return;
    const W = canvas.width, H = canvas.height;
    const PAD = { top: 30, right: 20, bottom: 50, left: 55 };
    const plotW = W - PAD.left - PAD.right;
    const plotH = H - PAD.top - PAD.bottom;

    ctx.clearRect(0, 0, W, H);
    ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, W, H);

    const dataToShow = metric === 'rate' ? result.data.map(d => d.rate) : result.data.map(d => d.count);
    const fittedVals = result.fitted || [];
    const allVals = [...dataToShow, ...fittedVals.map(f => f.ci_high), ...fittedVals.map(f => f.ci_low)].filter(v => v != null && !isNaN(v));
    if (!allVals.length) return;

    const minY = Math.max(0, Math.min(...allVals) * 0.9);
    const maxY = Math.max(...allVals) * 1.1;
    const roky = result.data.map(d => d.rok);
    const minX = Math.min(...roky), maxX = Math.max(...roky);

    const xScale = (rok: number) => PAD.left + ((rok - minX) / Math.max(maxX - minX, 1)) * plotW;
    const yScale = (v: number) => PAD.top + plotH - ((v - minY) / (maxY - minY)) * plotH;

    // Grid
    ctx.strokeStyle = '#F1F5F9'; ctx.lineWidth = 1;
    for (let i = 0; i <= 4; i++) {
      const y = PAD.top + (plotH / 4) * i;
      ctx.beginPath(); ctx.moveTo(PAD.left, y); ctx.lineTo(PAD.left + plotW, y); ctx.stroke();
      const val = maxY - ((maxY - minY) / 4) * i;
      ctx.fillStyle = '#94A3B8'; ctx.font = '10px Arial'; ctx.textAlign = 'right';
      ctx.fillText(val.toFixed(1), PAD.left - 5, y + 3);
    }

    // X axis labels
    ctx.fillStyle = '#64748B'; ctx.font = '10px Arial'; ctx.textAlign = 'center';
    roky.forEach(rok => {
      ctx.fillText(String(rok), xScale(rok), H - PAD.bottom + 18);
    });

    // Y axis title
    ctx.save(); ctx.translate(13, H / 2); ctx.rotate(-Math.PI / 2);
    ctx.fillStyle = '#64748B'; ctx.font = '10px Arial'; ctx.textAlign = 'center';
    ctx.fillText(metric === 'rate' ? 'Rate / 1000 hosp.' : 'Počet izolátov', 0, 0);
    ctx.restore();

    // CI shaded area
    if (fittedVals.length) {
      const trendColor = result.direction === 'up' ? '#DC2626' : result.direction === 'down' ? '#16A34A' : '#888780';
      ctx.beginPath();
      fittedVals.forEach((f, i) => {
        const x = xScale(f.rok), y = yScale(f.ci_high);
        i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y);
      });
      [...fittedVals].reverse().forEach(f => {
        ctx.lineTo(xScale(f.rok), yScale(f.ci_low));
      });
      ctx.closePath();
      ctx.fillStyle = trendColor + '22'; ctx.fill();

      // Trend line
      ctx.beginPath();
      fittedVals.forEach((f, i) => {
        const x = xScale(f.rok), y = yScale(f.fitted);
        i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y);
      });
      ctx.strokeStyle = trendColor; ctx.lineWidth = 2.5; ctx.setLineDash([]); ctx.stroke();
    }

    // Observed points
    ctx.fillStyle = '#185FA5';
    result.data.forEach((d, i) => {
      const x = xScale(d.rok);
      const y = yScale(dataToShow[i]);
      ctx.beginPath(); ctx.arc(x, y, 5, 0, Math.PI * 2);
      ctx.fillStyle = '#185FA5'; ctx.fill();
      ctx.strokeStyle = '#fff'; ctx.lineWidth = 1.5; ctx.stroke();
    });

    // Title
    ctx.fillStyle = '#0A1628'; ctx.font = 'bold 11px Arial'; ctx.textAlign = 'left';
    ctx.fillText(result.oddelenie.replace(/^DE\s+/, ''), PAD.left, 18);

    // IRR badge
    if (result.irr !== null) {
      const badge = `IRR ${result.irr.toFixed(3)} (95%CI ${result.ci_low?.toFixed(3)}–${result.ci_high?.toFixed(3)}) p=${result.p_value?.toFixed(4)}`;
      ctx.fillStyle = '#64748B'; ctx.font = '9px Arial'; ctx.textAlign = 'right';
      ctx.fillText(badge, W - PAD.right, 18);
    }
  };

  // After results, draw charts
  useEffect(() => {
    results.forEach(r => {
      const canvas = chartRefs.current[r.oddelenie];
      if (canvas) drawChart(canvas, r);
    });
  }, [results, metric]);

  const downloadChart = (oddelenie: string) => {
    const canvas = chartRefs.current[oddelenie];
    if (!canvas) return;
    const link = document.createElement('a');
    link.download = `trend_${oddelenie.replace(/[^a-zA-Z0-9]/g, '_')}.png`;
    link.href = canvas.toDataURL('image/png', 1.0);
    link.click();
  };

  const downloadAllCharts = () => {
    results.forEach(r => downloadChart(r.oddelenie));
  };

  const SS_BLUE = '#185FA5', SS_TEAL = '#0F6E56', SS_RED = '#A32D2D', SS_GREEN = '#15803D';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>

      {/* Controls */}
      <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '1rem 1.25rem' }}>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16, alignItems: 'flex-start' }}>
          <div>
            <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '.05em', color: '#94a3b8', marginBottom: 8 }}>Metrika</p>
            <div style={{ display: 'flex', gap: 6 }}>
              {([['rate', 'Rate / 1000 hosp.'], ['count', 'Absolútny počet']] as const).map(([v, l]) => (
                <button key={v} onClick={() => setMetric(v)} className={`ss-chip ${metric === v ? 'active' : ''}`}>{l}</button>
              ))}
            </div>
          </div>
          <div style={{ flex: 1, minWidth: 300 }}>
            <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '.05em', color: '#94a3b8', marginBottom: 8 }}>
              Oddelenia (vyber 1–8)
            </p>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 5 }}>
              {availableOddelenia.map(od => {
                const active = selectedOddelenia.includes(od);
                return (
                  <button key={od} onClick={() => setSelectedOddelenia(prev =>
                    active ? prev.filter(x => x !== od) : prev.length < 8 ? [...prev, od] : prev
                  )} className={`ss-chip ${active ? 'active' : ''}`} style={{ fontSize: 11, padding: '3px 10px' }}>
                    {od.replace(/^DE\s+/, '')}
                  </button>
                );
              })}
            </div>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, alignSelf: 'flex-end' }}>
            <button onClick={runAnalysis} disabled={loading || !selectedOddelenia.length}
              style={{ padding: '8px 20px', borderRadius: 9, background: loading || !selectedOddelenia.length ? '#e2e8f0' : 'linear-gradient(135deg,#185FA5,#0F6E56)', color: loading || !selectedOddelenia.length ? '#94a3b8' : '#fff', border: 'none', cursor: loading || !selectedOddelenia.length ? 'not-allowed' : 'pointer', fontSize: 13, fontWeight: 500, display: 'flex', alignItems: 'center', gap: 6 }}>
              {loading ? <span style={{ animation: 'spin 1s linear infinite', display: 'inline-block' }}>⟳</span> : '▶'}
              {loading ? 'Počítam...' : 'Spustiť analýzu'}
            </button>
            {results.length > 0 && (
              <button onClick={downloadAllCharts}
                style={{ padding: '6px 16px', borderRadius: 9, background: '#fff', border: '1px solid rgba(26,95,168,0.2)', color: SS_BLUE, cursor: 'pointer', fontSize: 12, fontWeight: 500 }}>
                ↓ Stiahnuť všetky grafy (PNG)
              </button>
            )}
          </div>
        </div>
        {error && <div style={{ marginTop: 10, padding: '8px 12px', background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 8, color: '#A32D2D', fontSize: 12 }}>{error}</div>}
      </div>

      {/* Info box */}
      <div style={{ background: 'linear-gradient(135deg,rgba(26,95,168,0.05),rgba(15,110,86,0.05))', borderRadius: 12, border: '1px solid rgba(26,95,168,0.1)', padding: '10px 16px', fontSize: 11, color: '#475569', display: 'flex', gap: 16, flexWrap: 'wrap' }}>
        <span>📊 <b>Poisson regression</b> — rok ako prediktor, log(hospitalizácie) ako offset</span>
        <span>📐 <b>Quasi-Poisson</b> — ak deviance/df &gt; 1.5 (overdispersion)</span>
        <span>🎯 <b>95% CI</b> zobrazené ako tienované pásmo okolo trend čiary</span>
        <span>🟢 zelená = štatisticky signifikantný pokles · 🔴 červená = nárast · ⚫ šedá = NS (p≥0.05)</span>
      </div>

      {/* Charts grid */}
      {results.length > 0 && (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(420px, 1fr))', gap: '1rem' }}>
            {results.map(r => (
              <div key={r.oddelenie} style={{ background: '#fff', borderRadius: 14, border: `1px solid ${r.direction === 'up' && r.significant ? '#fca5a5' : r.direction === 'down' && r.significant ? '#bbf7d0' : 'rgba(26,95,168,0.08)'}`, padding: '1rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <span style={{ fontSize: 18 }}>{r.direction === 'up' && r.significant ? '📈' : r.direction === 'down' && r.significant ? '📉' : '➡️'}</span>
                    <div>
                      <p style={{ fontSize: 12, fontWeight: 600, color: '#0A1628', margin: 0 }}>{r.oddelenie.replace(/^DE\s+/, '')}</p>
                      <p style={{ fontSize: 10, color: '#64748b', margin: 0 }}>
                        {r.model === 'insufficient' ? 'Nedostatok dát' : r.model === 'quasi-poisson' ? '⚠️ Quasi-Poisson (overdispersion)' : 'Poisson'}
                        {r.overdispersion !== null && ` · deviance/df: ${r.overdispersion}`}
                      </p>
                    </div>
                  </div>
                  <button onClick={() => downloadChart(r.oddelenie)}
                    style={{ padding: '4px 10px', borderRadius: 7, border: '1px solid rgba(26,95,168,0.15)', background: '#fff', color: SS_BLUE, cursor: 'pointer', fontSize: 11 }}>
                    ↓ PNG
                  </button>
                </div>
                <canvas ref={el => { chartRefs.current[r.oddelenie] = el; }} width={480} height={240}
                  style={{ width: '100%', borderRadius: 8, background: '#fafafa' }} />
              </div>
            ))}
          </div>

          {/* Results table */}
          <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '1.25rem 1.5rem' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
              <p style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: '.07em', color: '#94a3b8', margin: 0 }}>
                Výsledky Poisson regresie — IRR per rok
              </p>
              <span style={{ fontSize: 10, color: '#94a3b8' }}>IRR = Incidence Rate Ratio · CI = Confidence Interval · * p&lt;0.05 · ** p&lt;0.01 · *** p&lt;0.001</span>
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr>
                  {['Oddelenie','Model','IRR (per rok)','95% CI','p-hodnota','Deviance/df','Trend'].map(h => (
                    <th key={h} style={{ textAlign: 'left', padding: '8px 10px', color: '#94a3b8', fontWeight: 600, fontSize: 11, textTransform: 'uppercase', letterSpacing: '.04em', borderBottom: '1px solid #f1f5f9' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {results.map((r, i) => {
                  const sig = r.p_value !== null ? (r.p_value < 0.001 ? '***' : r.p_value < 0.01 ? '**' : r.p_value < 0.05 ? '*' : 'NS') : '—';
                  const trendColor = r.direction === 'up' && r.significant ? SS_RED : r.direction === 'down' && r.significant ? SS_GREEN : '#64748B';
                  return (
                    <tr key={r.oddelenie} style={{ background: i % 2 === 0 ? '#fff' : '#f8fafc' }}>
                      <td style={{ padding: '8px 10px', color: '#334155', borderBottom: '1px solid #f8fafc' }}>{r.oddelenie.replace(/^DE\s+/, '')}</td>
                      <td style={{ padding: '8px 10px', borderBottom: '1px solid #f8fafc' }}>
                        <span style={{ fontSize: 10, padding: '2px 7px', borderRadius: 5, background: r.model === 'quasi-poisson' ? '#fff7ed' : '#eff6ff', color: r.model === 'quasi-poisson' ? '#c2410c' : SS_BLUE, fontWeight: 600 }}>
                          {r.model === 'insufficient' ? 'N/A' : r.model === 'quasi-poisson' ? 'QP' : 'P'}
                        </span>
                      </td>
                      <td style={{ padding: '8px 10px', fontFamily: 'monospace', fontWeight: 700, color: trendColor, borderBottom: '1px solid #f8fafc' }}>
                        {r.irr !== null ? r.irr.toFixed(3) : '—'}
                      </td>
                      <td style={{ padding: '8px 10px', fontFamily: 'monospace', fontSize: 11, color: '#475569', borderBottom: '1px solid #f8fafc' }}>
                        {r.ci_low !== null ? `${r.ci_low.toFixed(3)} – ${r.ci_high?.toFixed(3)}` : '—'}
                      </td>
                      <td style={{ padding: '8px 10px', borderBottom: '1px solid #f8fafc' }}>
                        <span style={{ fontFamily: 'monospace', fontSize: 11 }}>{r.p_value !== null ? r.p_value.toFixed(4) : '—'}</span>
                        <span style={{ marginLeft: 6, fontWeight: 700, color: trendColor }}>{sig}</span>
                      </td>
                      <td style={{ padding: '8px 10px', fontFamily: 'monospace', fontSize: 11, color: r.overdispersion !== null && r.overdispersion > 1.5 ? '#c2410c' : '#64748b', borderBottom: '1px solid #f8fafc' }}>
                        {r.overdispersion !== null ? r.overdispersion.toFixed(2) : '—'}
                        {r.overdispersion !== null && r.overdispersion > 1.5 && <span style={{ marginLeft: 4, fontSize: 9, color: '#c2410c' }}>⚠️ OD</span>}
                      </td>
                      <td style={{ padding: '8px 10px', borderBottom: '1px solid #f8fafc' }}>
                        <span style={{ fontSize: 11, fontWeight: 600, color: trendColor }}>
                          {r.direction === 'up' && r.significant ? '↑ Rastúci' : r.direction === 'down' && r.significant ? '↓ Klesajúci' : '→ Stabilný'}
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <p style={{ fontSize: 10, color: '#94a3b8', marginTop: 10 }}>
              Poisson regression: log(count) ~ rok + offset(log(hosp)) · Quasi-Poisson pri deviance/df &gt; 1.5 · 95% CI zobrazené ako tienované pásmo v grafoch
            </p>
          </div>
        </>
      )}

      {results.length === 0 && !loading && (
        <div style={{ background: '#fff', borderRadius: 14, border: '1px solid rgba(26,95,168,0.08)', padding: '3rem', textAlign: 'center', color: '#94a3b8' }}>
          <p style={{ fontSize: 32, marginBottom: 8 }}>📊</p>
          <p style={{ fontSize: 14 }}>Vyber oddelenia a klikni "Spustiť analýzu"</p>
          <p style={{ fontSize: 12, marginTop: 4 }}>Poisson regression s 95% CI · Export grafov ako PNG</p>
        </div>
      )}
    </div>
  );
}
