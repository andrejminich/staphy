'use client';

import { useState, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend, LineChart, Line, CartesianGrid, ReferenceLine
} from 'recharts';
import { Upload, RefreshCw, AlertCircle, TrendingUp, Activity, Users, Shield } from 'lucide-react';
import Image from 'next/image';
import {
  Isolate, computeStats, getTopN, MONTH_NAMES,
  processRawData, RawRow, PandemicPeriod, PANDEMIC_PERIODS, getPandemicPeriod
} from '@/lib/mdr';

const C_MDR = '#dc2626';
const C_NONMDR = '#16a34a';
const C_AUREUS = '#1e40af';
const C_EPI = '#7c3aed';
const C_PRED = '#0891b2';
const C_POCAS = '#d97706';
const C_PO = '#7c3aed';
const CHART_COLORS = ['#1e40af','#7c3aed','#0891b2','#d97706','#dc2626','#16a34a','#db2777','#65a30d'];

const pct = (n: number, d: number) => d > 0 ? Math.round((n / d) * 100) : 0;

type FilterPatogen = 'all' | 'aureus' | 'epidermidis';
type FilterMdr = 'all' | 'mdr' | 'nonmdr';

export default function Dashboard() {
  const [isolates, setIsolates] = useState<Isolate[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [progress, setProgress] = useState<string | null>(null);

  const [filterPatogen, setFilterPatogen] = useState<FilterPatogen>('all');
  const [filterMdr, setFilterMdr] = useState<FilterMdr>('all');
  const [filterMesiac, setFilterMesiac] = useState<number | null>(null);
  const [filterOddelenie, setFilterOddelenie] = useState<string | null>(null);
  const [filterRok, setFilterRok] = useState<number | null>(null);
  const [filterPandemia, setFilterPandemia] = useState<PandemicPeriod>('all');

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
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<RawRow>(sheet, { defval: '' });
      if (rows.length === 0) throw new Error('Súbor je prázdny alebo má nesprávny formát.');
      const required = ['CisloProtokoluOKM', 'Patogen', 'NazovATB', 'CitlivostATB'];
      const cols = Object.keys(rows[0]);
      const missing = required.filter(r => !cols.includes(r));
      if (missing.length > 0) throw new Error(`Chýbajú stĺpce: ${missing.join(', ')}`);
      setProgress(`Kalkulujem MDR pre ${rows.length.toLocaleString()} riadkov...`);
      await new Promise(r => setTimeout(r, 50));
      const result = processRawData(rows);
      setIsolates(result);
      setFileName(file.name);
      setFilterPatogen('all');
      setFilterMdr('all');
      setFilterMesiac(null);
      setFilterOddelenie(null);
      setFilterRok(null);
      setFilterPandemia('all');
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Chyba pri spracovaní');
    } finally {
      setLoading(false);
      setProgress(null);
      e.target.value = '';
    }
  }, []);

  const filtered = useMemo(() => {
    return isolates.filter(iso => {
      if (filterPatogen === 'aureus' && iso.patogen.indexOf('aureus') < 0) return false;
      if (filterPatogen === 'epidermidis' && iso.patogen.indexOf('epidermidis') < 0) return false;
      if (filterMdr === 'mdr' && !iso.isMdr) return false;
      if (filterMdr === 'nonmdr' && iso.isMdr) return false;
      if (filterMesiac !== null && iso.mesiac !== filterMesiac) return false;
      if (filterOddelenie && iso.oddelenie !== filterOddelenie) return false;
      if (filterRok !== null && new Date(iso.datumOdberu).getFullYear() !== filterRok) return false;
      if (filterPandemia !== 'all' && getPandemicPeriod(iso.datumOdberu) !== filterPandemia) return false;
      return true;
    });
  }, [isolates, filterPatogen, filterMdr, filterMesiac, filterOddelenie, filterRok, filterPandemia]);

  const stats = useMemo(() => computeStats(filtered), [filtered]);
  const allStats = useMemo(() => computeStats(isolates), [isolates]);

  // Pandemic comparison data
  const pandemicComparison = useMemo(() => {
    const periods: PandemicPeriod[] = ['pred', 'pocas', 'po'];
    return periods.map(p => {
      const subset = isolates.filter(i => getPandemicPeriod(i.datumOdberu) === p);
      const s = computeStats(subset);
      return {
        obdobie: PANDEMIC_PERIODS[p].label,
        total: s.total,
        mdr: s.mdr,
        nonmdr: s.nonMdr,
        mdrRate: s.mdrRate,
        mrsa: s.mrsa,
        aureus: s.aureus,
        epidermidis: s.epidermidis,
      };
    }).filter(d => d.total > 0);
  }, [isolates]);

  const mdrPieData = [
    { name: 'MDR', value: stats.mdr, color: C_MDR },
    { name: 'Non-MDR', value: stats.nonMdr, color: C_NONMDR },
  ];

  const patogenPieData = [
    { name: 'S. aureus', value: stats.aureus, color: C_AUREUS },
    { name: 'S. epidermidis', value: stats.epidermidis, color: C_EPI },
  ];

  const monthlyData = useMemo(() => {
    const byMonth: Record<number, { mdr: number; nonmdr: number }> = {};
    for (let i = 0; i < filtered.length; i++) {
      const iso = filtered[i];
      const m = iso.mesiac;
      if (!byMonth[m]) byMonth[m] = { mdr: 0, nonmdr: 0 };
      if (iso.isMdr) byMonth[m].mdr++; else byMonth[m].nonmdr++;
    }
    return Object.keys(byMonth).map(Number).sort((a,b) => a-b)
      .map(m => ({ mesiac: MONTH_NAMES[m], mdr: byMonth[m].mdr, nonmdr: byMonth[m].nonmdr, total: byMonth[m].mdr + byMonth[m].nonmdr }));
  }, [filtered]);

  const yearlyData = useMemo(() => {
    const byYear: Record<number, { mdr: number; nonmdr: number }> = {};
    for (let i = 0; i < isolates.length; i++) {
      const iso = isolates[i];
      const y = new Date(iso.datumOdberu).getFullYear();
      if (!byYear[y]) byYear[y] = { mdr: 0, nonmdr: 0 };
      if (iso.isMdr) byYear[y].mdr++; else byYear[y].nonmdr++;
    }
    return Object.keys(byYear).map(Number).sort((a,b) => a-b)
      .map(y => ({ rok: y.toString(), mdr: byYear[y].mdr, nonmdr: byYear[y].nonmdr, total: byYear[y].mdr + byYear[y].nonmdr }));
  }, [isolates]);

  const topMaterials = useMemo(() => getTopN(filtered.map(i => i.material), 10), [filtered]);
  const topOddelenia = useMemo(() => getTopN(filtered.map(i => i.oddelenie), 10), [filtered]);

  const availableMonths = useMemo(() => {
    const s: Record<number, boolean> = {};
    for (let i = 0; i < isolates.length; i++) s[isolates[i].mesiac] = true;
    return Object.keys(s).map(Number).sort((a,b) => a-b);
  }, [isolates]);

  const availableRoky = useMemo(() => {
    const s: Record<number, boolean> = {};
    for (let i = 0; i < isolates.length; i++) s[new Date(isolates[i].datumOdberu).getFullYear()] = true;
    return Object.keys(s).map(Number).sort((a,b) => a-b);
  }, [isolates]);

  const availableOddelenia = useMemo(() => {
    const s: Record<string, boolean> = {};
    for (let i = 0; i < isolates.length; i++) s[isolates[i].oddelenie] = true;
    return Object.keys(s).sort();
  }, [isolates]);

  const PERIOD_COLORS: Record<string, string> = {
    'Pred pandémiou': C_PRED,
    'Počas pandémie': C_POCAS,
    'Po pandémii': C_PO,
  };

  if (isolates.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="card max-w-lg w-full mx-4 text-center">
          <Image src="/logo.png" alt="StaphySearch" width={280} height={75} className="mx-auto mb-6 w-64 h-auto" />
          <p className="text-slate-500 mb-2">Mikrobiologická analýza · Kramáre</p>
          <p className="text-slate-400 text-sm mb-8">Nahraj export zo systému (XLSX alebo CSV) pre začatie analýzy.</p>
          <label className="btn-primary cursor-pointer inline-flex items-center gap-2">
            {loading ? <RefreshCw size={16} className="animate-spin" /> : <Upload size={16} />}
            {loading ? (progress || 'Spracovávam...') : 'Nahrať súbor (.xlsx / .csv)'}
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleUpload} disabled={loading} />
          </label>
          {error && (
            <div className="mt-4 p-3 bg-red-50 text-red-700 rounded-lg flex items-center gap-2 text-sm">
              <AlertCircle size={16} />{error}
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Image src="/logo.png" alt="StaphySearch" width={160} height={44} className="h-9 w-auto" />
            <p className="text-slate-400 text-xs hidden md:block">{fileName} · {allStats.total} izolátov</p>
          </div>
          <div className="flex items-center gap-3">
            <a href="/chapter2" className="text-sm px-3 py-1.5 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50">Kapitola 2 →</a>
            <label className="btn-outline cursor-pointer inline-flex items-center gap-2">
              {loading ? <RefreshCw size={14} className="animate-spin" /> : <Upload size={14} />}
              {loading ? (progress || 'Spracovávam...') : 'Nahrať nový súbor'}
              <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleUpload} disabled={loading} />
            </label>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6 space-y-6">

        {/* Filters */}
        <div className="card space-y-4">
          <h2 className="font-semibold text-slate-700 text-sm uppercase tracking-wide">Filtre</h2>

          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Patogén</span>
            {(['all','aureus','epidermidis'] as FilterPatogen[]).map(val => (
              <button key={val} onClick={() => setFilterPatogen(val)}
                className={`filter-chip ${filterPatogen === val ? 'filter-chip-active' : 'filter-chip-inactive'}`}>
                {val === 'all' ? 'Všetky' : val === 'aureus' ? 'S. aureus' : 'S. epidermidis'}
              </button>
            ))}
          </div>

          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">MDR status</span>
            {(['all','mdr','nonmdr'] as FilterMdr[]).map(val => (
              <button key={val} onClick={() => setFilterMdr(val)}
                className={`filter-chip ${filterMdr === val ? 'filter-chip-active' : 'filter-chip-inactive'}`}>
                {val === 'all' ? 'Všetky' : val === 'mdr' ? 'MDR' : 'Non-MDR'}
              </button>
            ))}
          </div>

          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Pandémia</span>
            {(['all','pred','pocas','po'] as PandemicPeriod[]).map(val => (
              <button key={val} onClick={() => { setFilterPandemia(val); setFilterRok(null); }}
                className={`filter-chip ${filterPandemia === val ? 'filter-chip-active' : 'filter-chip-inactive'}`}
                style={filterPandemia === val && val !== 'all' ? { backgroundColor: PERIOD_COLORS[PANDEMIC_PERIODS[val].label], borderColor: PERIOD_COLORS[PANDEMIC_PERIODS[val].label] } : {}}>
                {PANDEMIC_PERIODS[val].label}
              </button>
            ))}
          </div>

          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Rok</span>
            <button onClick={() => setFilterRok(null)}
              className={`filter-chip ${filterRok === null ? 'filter-chip-active' : 'filter-chip-inactive'}`}>Všetky</button>
            {availableRoky.map(r => (
              <button key={r} onClick={() => { setFilterRok(filterRok === r ? null : r); setFilterPandemia('all'); }}
                className={`filter-chip ${filterRok === r ? 'filter-chip-active' : 'filter-chip-inactive'}`}>{r}</button>
            ))}
          </div>

          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Mesiac</span>
            <button onClick={() => setFilterMesiac(null)}
              className={`filter-chip ${filterMesiac === null ? 'filter-chip-active' : 'filter-chip-inactive'}`}>Všetky</button>
            {availableMonths.map(m => (
              <button key={m} onClick={() => setFilterMesiac(filterMesiac === m ? null : m)}
                className={`filter-chip ${filterMesiac === m ? 'filter-chip-active' : 'filter-chip-inactive'}`}>{MONTH_NAMES[m]}</button>
            ))}
          </div>

          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Oddelenie</span>
            <select value={filterOddelenie || ''} onChange={e => setFilterOddelenie(e.target.value || null)}
              className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 bg-white">
              <option value="">— všetky oddelenia —</option>
              {availableOddelenia.map(o => <option key={o} value={o}>{o}</option>)}
            </select>
            {filterOddelenie && <button onClick={() => setFilterOddelenie(null)} className="text-xs text-red-500 hover:underline">✕ zrušiť</button>}
          </div>
        </div>

        {/* Stat Cards */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="stat-card border-slate-200">
            <span className="text-3xl font-bold text-slate-800">{stats.total}</span>
            <span className="text-sm text-slate-500">Celkovo izolátov</span>
          </div>
          <div className="stat-card border-red-200">
            <span className="text-3xl font-bold text-red-600">{stats.mdr}</span>
            <span className="text-sm text-slate-500">MDR kmene</span>
            <span className="text-xs text-slate-400">{pct(stats.mdr, stats.total)}% z celku</span>
          </div>
          <div className="stat-card border-green-200">
            <span className="text-3xl font-bold text-green-600">{stats.nonMdr}</span>
            <span className="text-sm text-slate-500">Non-MDR</span>
            <span className="text-xs text-slate-400">{pct(stats.nonMdr, stats.total)}% z celku</span>
          </div>
          <div className="stat-card border-orange-200">
            <span className="text-3xl font-bold text-orange-600">{stats.mrsa}</span>
            <span className="text-sm text-slate-500">MRSA / MRCoNS</span>
            <span className="text-xs text-slate-400">{pct(stats.mrsa, stats.total)}% z celku</span>
          </div>
        </div>

        {/* Pandemic Comparison */}
        {pandemicComparison.length > 1 && (
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-1 flex items-center gap-2">
              <Shield size={16} /> Porovnanie období — pred / počas / po pandémii
            </h3>
            <p className="text-xs text-slate-400 mb-4">Pred: do 11.3.2020 · Počas: 12.3.2020 – 15.9.2023 (mimoriadna situácia SR) · Po: od 16.9.2023</p>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
              {pandemicComparison.map(p => (
                <div key={p.obdobie} className="rounded-xl border p-4 space-y-2"
                  style={{ borderColor: PERIOD_COLORS[p.obdobie] + '60' }}>
                  <div className="font-semibold text-slate-700" style={{ color: PERIOD_COLORS[p.obdobie] }}>{p.obdobie}</div>
                  <div className="text-2xl font-bold text-slate-800">{p.total} <span className="text-sm font-normal text-slate-400">izolátov</span></div>
                  <div className="flex gap-4 text-sm">
                    <div><span className="font-bold text-red-600">{p.mdrRate}%</span> <span className="text-slate-500">MDR</span></div>
                    <div><span className="font-bold text-orange-600">{p.mrsa}</span> <span className="text-slate-500">MRSA</span></div>
                  </div>
                  <div className="flex gap-3 text-xs text-slate-500">
                    <span>S. aureus: <b>{p.aureus}</b></span>
                    <span>S. epi: <b>{p.epidermidis}</b></span>
                  </div>
                  <div className="w-full bg-slate-100 rounded-full h-2 mt-1">
                    <div className="h-2 rounded-full" style={{ width: `${p.mdrRate}%`, backgroundColor: PERIOD_COLORS[p.obdobie] }} />
                  </div>
                </div>
              ))}
            </div>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={pandemicComparison} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                <XAxis dataKey="obdobie" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip />
                <Legend />
                <Bar dataKey="mdr" name="MDR" stackId="a" fill={C_MDR} />
                <Bar dataKey="nonmdr" name="Non-MDR" stackId="a" fill={C_NONMDR} radius={[4,4,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* Pie Charts */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4 flex items-center gap-2"><Activity size={16} /> MDR vs Non-MDR</h3>
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={mdrPieData} cx="50%" cy="50%" outerRadius={80} dataKey="value" label={({ name, value }) => `${name}: ${value}`}>
                  {mdrPieData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4 flex items-center gap-2"><Users size={16} /> Patogény</h3>
            <ResponsiveContainer width="100%" height={220}>
              <PieChart>
                <Pie data={patogenPieData} cx="50%" cy="50%" outerRadius={80} dataKey="value" label={({ name, value }) => `${name}: ${value}`}>
                  {patogenPieData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Yearly Trend */}
        <div className="card">
          <h3 className="font-semibold text-slate-700 mb-4 flex items-center gap-2"><TrendingUp size={16} /> Ročný trend — MDR vs Non-MDR</h3>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={yearlyData} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis dataKey="rok" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip />
              <Legend />
              <Bar dataKey="mdr" name="MDR" fill={C_MDR} stackId="a" />
              <Bar dataKey="nonmdr" name="Non-MDR" fill={C_NONMDR} stackId="a" radius={[4,4,0,0]} />
              <ReferenceLine x="2020" stroke={C_POCAS} strokeDasharray="4 4" label={{ value: '▶ pandémia', position: 'top', fontSize: 10, fill: C_POCAS }} />
              <ReferenceLine x="2023" stroke={C_PO} strokeDasharray="4 4" label={{ value: '▶ po pandémii', position: 'top', fontSize: 10, fill: C_PO }} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Monthly Trend */}
        <div className="card">
          <h3 className="font-semibold text-slate-700 mb-4 flex items-center gap-2"><TrendingUp size={16} /> Mesačný trend</h3>
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={monthlyData} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis dataKey="mesiac" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey="mdr" stroke={C_MDR} strokeWidth={2} dot={{ r: 4 }} name="MDR" />
              <Line type="monotone" dataKey="nonmdr" stroke={C_NONMDR} strokeWidth={2} dot={{ r: 4 }} name="Non-MDR" />
              <Line type="monotone" dataKey="total" stroke="#94a3b8" strokeWidth={1} strokeDasharray="4 4" dot={false} name="Celkom" />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* Materials + Oddelenia */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4">Najčastejšie materiály</h3>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={topMaterials} layout="vertical" margin={{ left: 10, right: 20 }}>
                <XAxis type="number" tick={{ fontSize: 11 }} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={140} />
                <Tooltip />
                <Bar dataKey="count" name="Počet" radius={[0,4,4,0]}>
                  {topMaterials.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4">Top 10 oddelení</h3>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={topOddelenia} layout="vertical" margin={{ left: 10, right: 20 }}>
                <XAxis type="number" tick={{ fontSize: 11 }} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 10 }} width={180} />
                <Tooltip />
                <Bar dataKey="count" name="Počet" radius={[0,4,4,0]}>
                  {topOddelenia.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Patogen per month */}
        <div className="card">
          <h3 className="font-semibold text-slate-700 mb-4">Rozloženie patogénov po mesiacoch</h3>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={(() => {
              const byMonth: Record<number, { aureus: number; epidermidis: number }> = {};
              for (let i = 0; i < filtered.length; i++) {
                const iso = filtered[i];
                const m = iso.mesiac;
                if (!byMonth[m]) byMonth[m] = { aureus: 0, epidermidis: 0 };
                if (iso.patogen.indexOf('aureus') >= 0) byMonth[m].aureus++;
                else byMonth[m].epidermidis++;
              }
              return Object.keys(byMonth).map(Number).sort((a,b) => a-b)
                .map(m => ({ mesiac: MONTH_NAMES[m], ...byMonth[m] }));
            })()} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis dataKey="mesiac" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip />
              <Legend />
              <Bar dataKey="aureus" name="S. aureus" fill={C_AUREUS} stackId="a" />
              <Bar dataKey="epidermidis" name="S. epidermidis" fill={C_EPI} stackId="a" radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Per-pathogen breakdown */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {['aureus', 'epidermidis'].map(pathogen => {
            const subset = filtered.filter(i => i.patogen.indexOf(pathogen) >= 0);
            const mdrCount = subset.filter(i => i.isMdr).length;
            const nonMdrCount = subset.length - mdrCount;
            const topMat = getTopN(subset.map(i => i.material), 6);
            const topDep = getTopN(subset.map(i => i.oddelenie), 5);
            return (
              <div key={pathogen} className="card space-y-4">
                <h3 className="font-semibold text-slate-700">S. {pathogen} <span className="text-slate-400 font-normal">({subset.length} izolátov)</span></h3>
                <div className="flex gap-4">
                  <div className="text-center">
                    <div className="text-2xl font-bold text-red-600">{mdrCount}</div>
                    <div className="text-xs text-slate-500">MDR ({pct(mdrCount, subset.length)}%)</div>
                  </div>
                  <div className="text-center">
                    <div className="text-2xl font-bold text-green-600">{nonMdrCount}</div>
                    <div className="text-xs text-slate-500">Non-MDR ({pct(nonMdrCount, subset.length)}%)</div>
                  </div>
                </div>
                <div>
                  <p className="text-xs text-slate-400 font-medium mb-2 uppercase">Top materiály</p>
                  <div className="space-y-1">
                    {topMat.map(({ name, count }) => (
                      <div key={name} className="flex items-center gap-2">
                        <div className="text-xs text-slate-600 w-36 truncate">{name}</div>
                        <div className="flex-1 bg-slate-100 rounded-full h-2">
                          <div className="h-2 rounded-full bg-blue-500" style={{ width: `${pct(count, subset.length)}%` }} />
                        </div>
                        <div className="text-xs text-slate-500 w-8 text-right">{count}</div>
                      </div>
                    ))}
                  </div>
                </div>
                <div>
                  <p className="text-xs text-slate-400 font-medium mb-2 uppercase">Top oddelenia</p>
                  <div className="space-y-1">
                    {topDep.map(({ name, count }) => (
                      <div key={name} className="flex items-center gap-2">
                        <div className="text-xs text-slate-600 flex-1 truncate">{name}</div>
                        <div className="text-xs font-semibold text-slate-700">{count}</div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        {/* Isolate Table */}
        <div className="card overflow-hidden">
          <h3 className="font-semibold text-slate-700 mb-4">Zoznam izolátov <span className="text-slate-400 font-normal text-sm">({filtered.length})</span></h3>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-200">
                  {['ID izolát','Patogén','Materiál','Oddelenie','Dátum','Obdobie','MDR','MRSA','R kat.'].map(h => (
                    <th key={h} className="text-left py-2 px-3 text-xs font-semibold text-slate-500 uppercase">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.slice(0, 100).map(iso => {
                  const period = getPandemicPeriod(iso.datumOdberu);
                  return (
                    <tr key={iso.id} className="border-b border-slate-100 hover:bg-slate-50">
                      <td className="py-2 px-3 font-mono text-xs text-slate-500">{iso.id}</td>
                      <td className="py-2 px-3 text-xs italic">{iso.patogen}</td>
                      <td className="py-2 px-3 text-xs">{iso.material}</td>
                      <td className="py-2 px-3 text-xs max-w-[140px] truncate" title={iso.oddelenie}>{iso.oddelenie}</td>
                      <td className="py-2 px-3 text-xs">{new Date(iso.datumOdberu).toLocaleDateString('sk-SK')}</td>
                      <td className="py-2 px-3">
                        <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-medium"
                          style={{ backgroundColor: PERIOD_COLORS[PANDEMIC_PERIODS[period].label] + '20', color: PERIOD_COLORS[PANDEMIC_PERIODS[period].label] }}>
                          {period === 'pred' ? 'Pred' : period === 'pocas' ? 'Počas' : 'Po'}
                        </span>
                      </td>
                      <td className="py-2 px-3">
                        <span className={`inline-flex px-2 py-0.5 rounded-full text-xs font-semibold ${iso.isMdr ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>
                          {iso.isMdr ? 'MDR' : 'Non-MDR'}
                        </span>
                      </td>
                      <td className="py-2 px-3">
                        {iso.isMrsa && <span className="inline-flex px-2 py-0.5 rounded-full text-xs font-semibold bg-orange-100 text-orange-700">MRSA</span>}
                      </td>
                      <td className="py-2 px-3 text-xs text-slate-500">{iso.resistantCount}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {filtered.length > 100 && (
              <p className="text-center text-xs text-slate-400 py-3">Zobrazujem prvých 100 z {filtered.length} izolátov</p>
            )}
          </div>
        </div>

        <footer className="text-center text-xs text-slate-400 py-4">
          MDR: Magiorakos et al. 2012 · Non-susceptible = R + SC · MRSA = auto-MDR · Pandémia SR: 12.3.2020 – 15.9.2023
        </footer>
      </main>
    </div>
  );
}
