'use client';

import { useState, useCallback, useMemo } from 'react';
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend, LineChart, Line, CartesianGrid
} from 'recharts';
import { Upload, RefreshCw, AlertCircle, TrendingUp, FlaskConical, Activity, Users } from 'lucide-react';
import { Isolate, computeStats, getTopN, MONTH_NAMES } from '@/lib/mdr';

// ─── colour palette ────────────────────────────────────────────────────────
const C_MDR = '#dc2626';
const C_NONMDR = '#16a34a';
const C_AUREUS = '#1e40af';
const C_EPI = '#7c3aed';
const CHART_COLORS = ['#1e40af','#7c3aed','#0891b2','#d97706','#dc2626','#16a34a','#db2777','#65a30d'];

// ─── helpers ───────────────────────────────────────────────────────────────
const pct = (n: number, d: number) => d > 0 ? Math.round((n / d) * 100) : 0;

type FilterPatogen = 'all' | 'aureus' | 'epidermidis';
type FilterMdr = 'all' | 'mdr' | 'nonmdr';

export default function Dashboard() {
  const [isolates, setIsolates] = useState<Isolate[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);

  // ── filters ──
  const [filterPatogen, setFilterPatogen] = useState<FilterPatogen>('all');
  const [filterMdr, setFilterMdr] = useState<FilterMdr>('all');
  const [filterMesiac, setFilterMesiac] = useState<number | null>(null);
  const [filterOddelenie, setFilterOddelenie] = useState<string | null>(null);

  // ── upload handler ──
  const handleUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setError(null);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch('/api/upload', { method: 'POST', body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Neznáma chyba');
      setIsolates(data.isolates);
      setFileName(file.name);
      setFilterPatogen('all');
      setFilterMdr('all');
      setFilterMesiac(null);
      setFilterOddelenie(null);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : 'Chyba pri nahrávaní');
    } finally {
      setLoading(false);
      e.target.value = '';
    }
  }, []);

  // ── filtered isolates ──
  const filtered = useMemo(() => {
    return isolates.filter(iso => {
      if (filterPatogen === 'aureus' && !iso.patogen.includes('aureus')) return false;
      if (filterPatogen === 'epidermidis' && !iso.patogen.includes('epidermidis')) return false;
      if (filterMdr === 'mdr' && !iso.isMdr) return false;
      if (filterMdr === 'nonmdr' && iso.isMdr) return false;
      if (filterMesiac !== null && iso.mesiac !== filterMesiac) return false;
      if (filterOddelenie && iso.oddelenie !== filterOddelenie) return false;
      return true;
    });
  }, [isolates, filterPatogen, filterMdr, filterMesiac, filterOddelenie]);

  const stats = useMemo(() => computeStats(filtered), [filtered]);
  const allStats = useMemo(() => computeStats(isolates), [isolates]);

  // ── chart data ──
  const mdrPieData = [
    { name: 'MDR', value: stats.mdr, color: C_MDR },
    { name: 'Non-MDR', value: stats.nonMdr, color: C_NONMDR },
  ];

  const patogenPieData = [
    { name: 'S. aureus', value: stats.aureus, color: C_AUREUS },
    { name: 'S. epidermidis', value: stats.epidermidis, color: C_EPI },
  ];

  // Monthly trend — from all filtered data
  const monthlyData = useMemo(() => {
    const byMonth = new Map<number, { mdr: number; nonmdr: number }>();
    for (const iso of filtered) {
      const m = iso.mesiac;
      if (!byMonth.has(m)) byMonth.set(m, { mdr: 0, nonmdr: 0 });
      const entry = byMonth.get(m)!;
      if (iso.isMdr) entry.mdr++; else entry.nonmdr++;
    }
    return Array.from(byMonth.entries())
      .sort((a, b) => a[0] - b[0])
      .map(([m, v]) => ({ mesiac: MONTH_NAMES[m], mdr: v.mdr, nonmdr: v.nonmdr, total: v.mdr + v.nonmdr }));
  }, [filtered]);

  // Top materials
  const topMaterials = useMemo(() =>
    getTopN(filtered.map(i => i.material), 10), [filtered]);

  // Top oddelenia
  const topOddelenia = useMemo(() =>
    getTopN(filtered.map(i => i.oddelenie), 10), [filtered]);

  // All months present in data
  const availableMonths = useMemo(() => {
    const months = new Set(isolates.map(i => i.mesiac));
    return Array.from(months).sort((a, b) => a - b);
  }, [isolates]);

  // All oddelenia
  const availableOddelenia = useMemo(() => {
    const set = new Set(isolates.map(i => i.oddelenie));
    return Array.from(set).sort();
  }, [isolates]);

  // ── empty state ──
  if (isolates.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="card max-w-lg w-full mx-4 text-center">
          <FlaskConical className="mx-auto mb-4 text-blue-700" size={48} />
          <h1 className="text-2xl font-bold text-slate-800 mb-2">Rezistencia Dashboard</h1>
          <p className="text-slate-500 mb-2">Mikrobiologická analýza · Kramáre</p>
          <p className="text-slate-400 text-sm mb-8">
            Nahraj export zo systému (XLSX alebo CSV) pre začatie analýzy.
          </p>
          <label className="btn-primary cursor-pointer inline-flex items-center gap-2">
            {loading ? <RefreshCw size={16} className="animate-spin" /> : <Upload size={16} />}
            {loading ? 'Spracovávam...' : 'Nahrať súbor (.xlsx / .csv)'}
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
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <FlaskConical className="text-blue-700" size={24} />
            <div>
              <h1 className="font-bold text-slate-800 text-lg leading-tight">Rezistencia Dashboard</h1>
              <p className="text-slate-400 text-xs">{fileName} · {allStats.total} izolátov</p>
            </div>
          </div>
          <label className="btn-outline cursor-pointer inline-flex items-center gap-2">
            {loading ? <RefreshCw size={14} className="animate-spin" /> : <Upload size={14} />}
            {loading ? 'Spracovávam...' : 'Nahrať nový súbor'}
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleUpload} disabled={loading} />
          </label>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-6 space-y-6">

        {/* ── Filters ── */}
        <div className="card space-y-4">
          <h2 className="font-semibold text-slate-700 text-sm uppercase tracking-wide">Filtre</h2>
          
          {/* Patogen */}
          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Patogén</span>
            {([
              ['all', 'Všetky'],
              ['aureus', 'S. aureus'],
              ['epidermidis', 'S. epidermidis'],
            ] as [FilterPatogen, string][]).map(([val, label]) => (
              <button
                key={val}
                onClick={() => setFilterPatogen(val)}
                className={`filter-chip ${filterPatogen === val ? 'filter-chip-active' : 'filter-chip-inactive'}`}
              >{label}</button>
            ))}
          </div>

          {/* MDR */}
          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">MDR status</span>
            {([
              ['all', 'Všetky'],
              ['mdr', 'MDR'],
              ['nonmdr', 'Non-MDR'],
            ] as [FilterMdr, string][]).map(([val, label]) => (
              <button
                key={val}
                onClick={() => setFilterMdr(val)}
                className={`filter-chip ${filterMdr === val ? 'filter-chip-active' : 'filter-chip-inactive'}`}
              >{label}</button>
            ))}
          </div>

          {/* Mesiac */}
          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Mesiac</span>
            <button
              onClick={() => setFilterMesiac(null)}
              className={`filter-chip ${filterMesiac === null ? 'filter-chip-active' : 'filter-chip-inactive'}`}
            >Všetky</button>
            {availableMonths.map(m => (
              <button
                key={m}
                onClick={() => setFilterMesiac(filterMesiac === m ? null : m)}
                className={`filter-chip ${filterMesiac === m ? 'filter-chip-active' : 'filter-chip-inactive'}`}
              >{MONTH_NAMES[m]}</button>
            ))}
          </div>

          {/* Oddelenie */}
          <div className="flex flex-wrap gap-2 items-center">
            <span className="text-xs text-slate-400 w-20">Oddelenie</span>
            <button
              onClick={() => setFilterOddelenie(null)}
              className={`filter-chip ${filterOddelenie === null ? 'filter-chip-active' : 'filter-chip-inactive'}`}
            >Všetky</button>
            <select
              value={filterOddelenie || ''}
              onChange={e => setFilterOddelenie(e.target.value || null)}
              className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm text-slate-700 bg-white"
            >
              <option value="">— vyber oddelenie —</option>
              {availableOddelenia.map(o => (
                <option key={o} value={o}>{o}</option>
              ))}
            </select>
            {filterOddelenie && (
              <button onClick={() => setFilterOddelenie(null)} className="text-xs text-red-500 hover:underline">✕ zrušiť</button>
            )}
          </div>
        </div>

        {/* ── Stat Cards ── */}
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

        {/* ── Pie Charts Row ── */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4 flex items-center gap-2">
              <Activity size={16} /> MDR vs Non-MDR
            </h3>
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
            <h3 className="font-semibold text-slate-700 mb-4 flex items-center gap-2">
              <Users size={16} /> Patogény
            </h3>
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

        {/* ── Monthly Trend ── */}
        <div className="card">
          <h3 className="font-semibold text-slate-700 mb-4 flex items-center gap-2">
            <TrendingUp size={16} /> Mesačný trend — MDR vs Non-MDR
          </h3>
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

        {/* ── Materials + Oddelenia ── */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4">Najčastejšie materiály</h3>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={topMaterials} layout="vertical" margin={{ left: 10, right: 20 }}>
                <XAxis type="number" tick={{ fontSize: 11 }} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={140} />
                <Tooltip />
                <Bar dataKey="count" name="Počet" radius={[0, 4, 4, 0]}>
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
                <Bar dataKey="count" name="Počet" radius={[0, 4, 4, 0]}>
                  {topOddelenia.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* ── MDR by Patogen per Month ── */}
        <div className="card">
          <h3 className="font-semibold text-slate-700 mb-4">Rozloženie patogénov po mesiacoch</h3>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={(() => {
              const byMonth = new Map<number, { aureus: number; epidermidis: number }>();
              for (const iso of filtered) {
                const m = iso.mesiac;
                if (!byMonth.has(m)) byMonth.set(m, { aureus: 0, epidermidis: 0 });
                const e = byMonth.get(m)!;
                if (iso.patogen.includes('aureus')) e.aureus++;
                else e.epidermidis++;
              }
              return Array.from(byMonth.entries()).sort((a, b) => a[0] - b[0])
                .map(([m, v]) => ({ mesiac: MONTH_NAMES[m], ...v }));
            })()} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis dataKey="mesiac" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip />
              <Legend />
              <Bar dataKey="aureus" name="S. aureus" fill={C_AUREUS} stackId="a" radius={[0, 0, 0, 0]} />
              <Bar dataKey="epidermidis" name="S. epidermidis" fill={C_EPI} stackId="a" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* ── MDR breakdown: aureus vs epidermidis ── */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          {['aureus', 'epidermidis'].map(pathogen => {
            const subset = filtered.filter(i => i.patogen.includes(pathogen));
            const mdrCount = subset.filter(i => i.isMdr).length;
            const nonMdrCount = subset.length - mdrCount;
            const topMat = getTopN(subset.map(i => i.material), 6);
            const topDep = getTopN(subset.map(i => i.oddelenie), 5);
            return (
              <div key={pathogen} className="card space-y-4">
                <h3 className="font-semibold text-slate-700">
                  S. {pathogen} <span className="text-slate-400 font-normal">({subset.length} izolátov)</span>
                </h3>
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
                          <div
                            className="h-2 rounded-full bg-blue-500"
                            style={{ width: `${pct(count, subset.length)}%` }}
                          />
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

        {/* ── Isolate Table ── */}
        <div className="card overflow-hidden">
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-semibold text-slate-700">Zoznam izolátov <span className="text-slate-400 font-normal text-sm">({filtered.length})</span></h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-slate-200">
                  {['ID izolát', 'Patogén', 'Materiál', 'Oddelenie', 'Dátum', 'MDR', 'MRSA', 'R kategórie'].map(h => (
                    <th key={h} className="text-left py-2 px-3 text-xs font-semibold text-slate-500 uppercase">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.slice(0, 100).map(iso => (
                  <tr key={iso.id} className="border-b border-slate-100 hover:bg-slate-50">
                    <td className="py-2 px-3 font-mono text-xs text-slate-500">{iso.id}</td>
                    <td className="py-2 px-3 text-xs italic">{iso.patogen}</td>
                    <td className="py-2 px-3 text-xs">{iso.material}</td>
                    <td className="py-2 px-3 text-xs max-w-[160px] truncate" title={iso.oddelenie}>{iso.oddelenie}</td>
                    <td className="py-2 px-3 text-xs">{new Date(iso.datumOdberu).toLocaleDateString('sk-SK')}</td>
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
                ))}
              </tbody>
            </table>
            {filtered.length > 100 && (
              <p className="text-center text-xs text-slate-400 py-3">Zobrazujem prvých 100 z {filtered.length} izolátov</p>
            )}
          </div>
        </div>

        <footer className="text-center text-xs text-slate-400 py-4">
          MDR definícia: Magiorakos et al. 2012 (ECDC/CDC) · Non-susceptible = R + SC · MRSA = automaticky MDR
        </footer>
      </main>
    </div>
  );
}
