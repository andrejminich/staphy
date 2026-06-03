'use client';

import { useState, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { Upload, RefreshCw, AlertCircle, FlaskConical } from 'lucide-react';
import Image from 'next/image';

const CC_MAP: Record<string, string> = {
  '1':'CC1','5':'CC5','6':'CC5','8':'CC8','15':'CC15','25':'CC25',
  '45':'CC45','72':'CC72','80':'CC80','225':'CC8','398':'CC398',
  '1570':'CC1','3975':'CC398','7407':'CC398',
};
const CC_COLORS: Record<string,string> = {
  'CC398':'#185FA5','CC8':'#0F6E56','CC15':'#854F0B','CC5':'#993556',
  'CC1':'#534AB7','CC45':'#639922','CC72':'#3B6D11','CC80':'#D85A30',
  'CC25':'#0891b2','Neznámy':'#888780',
};
const COLORS = ['#185FA5','#0F6E56','#854F0B','#993556','#534AB7','#639922','#3B6D11','#D85A30','#0891b2','#888780'];

interface Strain {
  ppid:string; assembly:number; st:string; cc:string; stLabel:string;
  org:string; material:string; klinika:string; rezistencia:string;
  rezistencia2:string; datum:string; isMrsa:boolean;
}

function getCC(st:string):string { return CC_MAP[st] || 'Neznámy'; }
function getTopN(items:string[],n:number):{name:string;count:number}[] {
  const c:Record<string,number>={};
  items.forEach(i=>{c[i]=(c[i]||0)+1;});
  return Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,n).map(([name,count])=>({name,count}));
}
type Filt='all'|'aureus'|'epidermidis'|'mrsa'|'known';

export default function Chapter2() {
  const [strains,setStrains]=useState<Strain[]>([]);
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState<string|null>(null);
  const [progress,setProgress]=useState<string|null>(null);
  const [filt,setFilt]=useState<Filt>('all');
  const [clinLoaded,setClinLoaded]=useState(false);
  const [seqLoaded,setSeqLoaded]=useState(false);
  const [clinData,setClinData]=useState<Record<string,{org:string;material:string;klinika:string;rezistencia:string;rezistencia2:string;datum:string;}>>({});
  const [seqData,setSeqData]=useState<{ppid:string;st:string}[]>([]);

  const merge=useCallback((clin:typeof clinData, seq:typeof seqData)=>{
    const result:Strain[]=seq.map(s=>{
      const c=clin[s.ppid]||{org:'Neznámy',material:'N/A',klinika:'N/A',rezistencia:'',rezistencia2:'',datum:''};
      const isMrsa=c.rezistencia.includes('MRSA')||c.rezistencia.includes('MRCoNS')||c.rezistencia2.includes('MRSA')||c.rezistencia2.includes('MRCoNS');
      return {...c,ppid:s.ppid,assembly:0,st:s.st,cc:getCC(s.st),stLabel:s.st==='-'?'Nový ST':`ST${s.st}`,isMrsa};
    });
    setStrains(result);
  },[]);

  const handleClin=useCallback(async(e:React.ChangeEvent<HTMLInputElement>)=>{
    const file=e.target.files?.[0]; if(!file)return;
    setLoading(true);setError(null);setProgress('Načítavam klinické dáta...');
    try {
      const text=await file.text();
      const lines=text.split('\n');
      const headers=lines[0].split(',').map(h=>h.replace(/^"|"$/g,'').trim());
      const clin:typeof clinData={};
      for(let i=1;i<lines.length;i++){
        if(!lines[i].trim())continue;
        const vals=lines[i].split(',').map(v=>v.replace(/^"|"$/g,'').trim());
        const row:Record<string,string>={};
        headers.forEach((h,j)=>{row[h]=vals[j]||'';});
        const ppid=row['PPID']; if(!ppid)continue;
        clin[ppid]={
          org:(row['Specimen registration- Microbiology#Microorganism (Mikroorganizmus)']||'').replace('Staphylococcus ','S. '),
          material:row['Specimen registration- Microbiology#Type of biological material (Druh biologického materiálu)']||'',
          klinika:row['Specimen registration- Microbiology#Place of origin (Miesto výskytu)']||'',
          rezistencia:row['Specimen registration- Microbiology#Resistance mechanisms (Mechanizmy rezistencie)#1#Resistance mechanism (Mechanizmus rezistencie)']||'',
          rezistencia2:row['Specimen registration- Microbiology#Resistance mechanisms (Mechanizmy rezistencie)#2#Resistance mechanism (Mechanizmus rezistencie)']||'',
          datum:row['Registration Date']||'',
        };
      }
      setClinData(clin);setClinLoaded(true);
      if(seqLoaded)merge(clin,seqData);
    } catch{setError('Chyba pri spracovaní CSV.');}
    finally{setLoading(false);setProgress(null);e.target.value='';}
  },[seqLoaded,seqData,merge]);

  const handleSeq=useCallback(async(e:React.ChangeEvent<HTMLInputElement>)=>{
    const file=e.target.files?.[0]; if(!file)return;
    setLoading(true);setError(null);setProgress('Načítavam sekvenačné dáta...');
    try {
      const buffer=await file.arrayBuffer();
      const wb=XLSX.read(buffer,{type:'array'});
      const sheet=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json<Record<string,unknown>>(sheet,{defval:''});
      const keys=Object.keys(rows[0]);
      const assemblyKey=keys[0]; const stKey=keys[31];
      const seq:{ppid:string;st:string}[]=[];
      rows.slice(1).forEach(row=>{
        const assembly=Number(row[assemblyKey]); if(isNaN(assembly))return;
        const ppid=`STS_${String(assembly+9).padStart(5,'0')}`;
        seq.push({ppid,st:String(row[stKey]||'-').trim()});
      });
      setSeqData(seq);setSeqLoaded(true);
      if(clinLoaded)merge(clinData,seq);
    } catch{setError('Chyba pri spracovaní XLSX.');}
    finally{setLoading(false);setProgress(null);e.target.value='';}
  },[clinLoaded,clinData,merge]);

  const filtered=useMemo(()=>strains.filter(s=>{
    if(filt==='aureus')return s.org.includes('aureus');
    if(filt==='epidermidis')return s.org.includes('epidermidis');
    if(filt==='mrsa')return s.isMrsa;
    if(filt==='known')return s.st!=='-';
    return true;
  }),[strains,filt]);

  const stCounts=useMemo(()=>getTopN(filtered.map(s=>s.stLabel),12),[filtered]);
  const ccCounts=useMemo(()=>getTopN(filtered.map(s=>s.cc),10),[filtered]);
  const matCounts=useMemo(()=>getTopN(filtered.map(s=>s.material),8),[filtered]);
  const rezCounts=useMemo(()=>getTopN(filtered.map(s=>s.rezistencia||'Bez rezistencie'),8),[filtered]);
  const klCounts=useMemo(()=>getTopN(filtered.map(s=>s.klinika),8),[filtered]);
  const pct=(n:number,d:number)=>d>0?Math.round((n/d)*100):0;

  const rezTag=(rez:string)=>{
    if(!rez||rez==='Bez rezistencie')return <span style={{display:'inline-block',fontSize:11,padding:'2px 8px',borderRadius:4,fontWeight:500,background:'#EAF3DE',color:'#3B6D11'}}>Bez rez.</span>;
    if(rez.includes('MRSA')||rez.includes('MRCoNS'))return <span style={{display:'inline-block',fontSize:11,padding:'2px 8px',borderRadius:4,fontWeight:500,background:'#FCEBEB',color:'#A32D2D'}}>{rez}</span>;
    if(rez.includes('penicilins'))return <span style={{display:'inline-block',fontSize:11,padding:'2px 8px',borderRadius:4,fontWeight:500,background:'#FAEEDA',color:'#854F0B'}}>Pen-R</span>;
    if(rez.includes('MLSB')||rez.includes('Constitutive')||rez.includes('Inducible'))return <span style={{display:'inline-block',fontSize:11,padding:'2px 8px',borderRadius:4,fontWeight:500,background:'#EEEDFE',color:'#3C3489'}}>{rez}</span>;
    return <span style={{display:'inline-block',fontSize:11,padding:'2px 8px',borderRadius:4,fontWeight:500,background:'#f1f5f9',color:'#64748b'}}>{rez}</span>;
  };

  if(strains.length===0){
    return (
      <div className="min-h-screen bg-slate-50">
        <header className="bg-white border-b border-slate-200 px-6 py-4">
          <div className="max-w-7xl mx-auto flex items-center gap-4">
            <Image src="/logo.png" alt="StaphySearch" width={160} height={44} className="h-9 w-auto" />
            <div className="flex gap-2">
              <a href="/" className="text-sm px-3 py-1.5 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50">Kapitola 1</a>
              <span className="text-sm px-3 py-1.5 rounded-lg bg-blue-700 text-white">Kapitola 2</span>
            </div>
          </div>
        </header>
        <div className="max-w-2xl mx-auto px-4 py-12 text-center">
          <FlaskConical className="mx-auto mb-4 text-blue-700" size={40} />
          <h2 className="text-xl font-bold text-slate-800 mb-2">Kapitola 2 — Prospektívna štúdia</h2>
          <p className="text-slate-500 text-sm mb-8">Nahraj klinické dáta (CSV) a sekvenačné dáta (XLSX s MLST).</p>
          <div className="grid grid-cols-2 gap-4">
            <div className="card text-center">
              <p className="text-xs text-slate-500 mb-3 font-medium">Klinické dáta (output.csv)</p>
              <label className={`cursor-pointer inline-flex items-center gap-2 text-sm px-4 py-2 rounded-lg text-white ${clinLoaded?'bg-green-600':'bg-blue-700'}`}>
                {clinLoaded?'✓ Načítané':<><Upload size={14}/>Nahrať CSV</>}
                <input type="file" accept=".csv" className="hidden" onChange={handleClin} disabled={loading}/>
              </label>
            </div>
            <div className="card text-center">
              <p className="text-xs text-slate-500 mb-3 font-medium">Sekvenačné dáta (MLST xlsx)</p>
              <label className={`cursor-pointer inline-flex items-center gap-2 text-sm px-4 py-2 rounded-lg text-white ${seqLoaded?'bg-green-600':'bg-blue-700'}`}>
                {seqLoaded?'✓ Načítané':<><Upload size={14}/>Nahrať XLSX</>}
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleSeq} disabled={loading}/>
              </label>
            </div>
          </div>
          {loading&&<p className="text-sm text-slate-400 mt-4 flex items-center justify-center gap-2"><RefreshCw size={14} className="animate-spin"/>{progress}</p>}
          {error&&<div className="mt-4 p-3 bg-red-50 text-red-700 rounded-lg flex items-center gap-2 text-sm"><AlertCircle size={16}/>{error}</div>}
        </div>
      </div>
    );
  }

  const mrsaCount=filtered.filter(s=>s.isMrsa).length;
  const knownST=filtered.filter(s=>s.st!=='-').length;
  const newST=filtered.filter(s=>s.st==='-').length;
  const st398=filtered.filter(s=>s.st==='398').length;

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <Image src="/logo.png" alt="StaphySearch" width={160} height={44} className="h-9 w-auto"/>
            <div className="flex gap-2">
              <a href="/" className="text-sm px-3 py-1.5 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50">Kapitola 1</a>
              <span className="text-sm px-3 py-1.5 rounded-lg bg-blue-700 text-white">Kapitola 2</span>
            </div>
          </div>
          <div className="flex gap-2">
            <label className="cursor-pointer inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50">
              <Upload size={12}/>CSV
              <input type="file" accept=".csv" className="hidden" onChange={handleClin} disabled={loading}/>
            </label>
            <label className="cursor-pointer inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-50">
              <Upload size={12}/>XLSX
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleSeq} disabled={loading}/>
            </label>
          </div>
        </div>
      </header>
      <main className="max-w-7xl mx-auto px-4 py-6 space-y-5">
        <div className="card space-y-3">
          <h2 className="font-semibold text-slate-700 text-sm uppercase tracking-wide">Filtre</h2>
          <div className="flex flex-wrap gap-2">
            {([['all','Všetky'],['aureus','S. aureus'],['epidermidis','S. epidermidis'],['mrsa','MRSA/MRCoNS'],['known','Známy ST']] as [Filt,string][]).map(([v,l])=>(
              <button key={v} onClick={()=>setFilt(v)} className={`filter-chip ${filt===v?'filter-chip-active':'filter-chip-inactive'}`}>{l}</button>
            ))}
          </div>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="stat-card border-slate-200"><span className="text-3xl font-bold text-slate-800">{filtered.length}</span><span className="text-sm text-slate-500">Kmeňov</span></div>
          <div className="stat-card border-red-200"><span className="text-3xl font-bold text-red-600">{mrsaCount}</span><span className="text-sm text-slate-500">MRSA/MRCoNS</span><span className="text-xs text-slate-400">{pct(mrsaCount,filtered.length)}%</span></div>
          <div className="stat-card border-blue-200"><span className="text-3xl font-bold text-blue-700">{st398}</span><span className="text-sm text-slate-500">ST398 (CC398)</span></div>
          <div className="stat-card border-orange-200"><span className="text-3xl font-bold text-orange-600">{newST}</span><span className="text-sm text-slate-500">Nové ST</span><span className="text-xs text-slate-400">{knownST} známych</span></div>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4">Distribúcia ST typov</h3>
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={stCounts} layout="vertical" margin={{left:10,right:20}}>
                <XAxis type="number" tick={{fontSize:11}}/>
                <YAxis type="category" dataKey="name" tick={{fontSize:11}} width={80}/>
                <Tooltip/>
                <Bar dataKey="count" name="Počet" radius={[0,4,4,0]}>
                  {stCounts.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="card">
            <h3 className="font-semibold text-slate-700 mb-4">Klonálne komplexy (CC)</h3>
            <ResponsiveContainer width="100%" height={200}>
              <PieChart>
                <Pie data={ccCounts} cx="50%" cy="50%" outerRadius={80} dataKey="count" nameKey="name" label={({name,count})=>`${name}:${count}`} labelLine={false}>
                  {ccCounts.map((e,i)=><Cell key={i} fill={CC_COLORS[e.name]||COLORS[i%COLORS.length]}/>)}
                </Pie>
                <Tooltip/>
              </PieChart>
            </ResponsiveContainer>
            <div className="flex flex-wrap gap-2 mt-2">
              {ccCounts.map(({name,count})=>(
                <span key={name} className="flex items-center gap-1 text-xs text-slate-500">
                  <span style={{width:10,height:10,borderRadius:2,background:CC_COLORS[name]||'#888',display:'inline-block'}}/>
                  {name} ({count})
                </span>
              ))}
            </div>
          </div>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {[['Rezistenčné mechanizmy',rezCounts,'#dc2626'],['Typ materiálu',matCounts,'#0891b2'],['Top kliniky',klCounts,'#7c3aed']].map(([title,data,color])=>(
            <div key={title as string} className="card">
              <h3 className="font-semibold text-slate-700 mb-4">{title as string}</h3>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={data as {name:string;count:number}[]} layout="vertical" margin={{left:10,right:20}}>
                  <XAxis type="number" tick={{fontSize:10}}/>
                  <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={120}/>
                  <Tooltip/>
                  <Bar dataKey="count" name="Počet" fill={color as string} radius={[0,4,4,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
          ))}
        </div>
        <div className="card overflow-hidden">
          <h3 className="font-semibold text-slate-700 mb-4">Zoznam kmeňov <span className="text-slate-400 font-normal text-sm">({filtered.length})</span></h3>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="border-b border-slate-200">
                {['PPID','Organizmus','ST','CC','Materiál','Klinika','Rezistencia'].map(h=>(
                  <th key={h} className="text-left py-2 px-3 text-xs font-semibold text-slate-500 uppercase">{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {filtered.slice(0,100).map(s=>(
                  <tr key={s.ppid} className="border-b border-slate-100 hover:bg-slate-50">
                    <td className="py-2 px-3 font-mono text-xs text-slate-400">{s.ppid}</td>
                    <td className="py-2 px-3 text-xs italic">{s.org}</td>
                    <td className="py-2 px-3"><span className={`inline-flex px-2 py-0.5 rounded text-xs font-semibold ${s.st==='-'?'bg-slate-100 text-slate-400':'bg-blue-50 text-blue-700'}`}>{s.stLabel}</span></td>
                    <td className="py-2 px-3"><span className="inline-flex px-2 py-0.5 rounded text-xs font-semibold bg-teal-50 text-teal-700">{s.cc}</span></td>
                    <td className="py-2 px-3 text-xs">{s.material}</td>
                    <td className="py-2 px-3 text-xs text-slate-500 max-w-[150px] truncate" title={s.klinika}>{s.klinika}</td>
                    <td className="py-2 px-3">{rezTag(s.rezistencia)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {filtered.length>100&&<p className="text-center text-xs text-slate-400 py-3">Prvých 100 z {filtered.length}</p>}
          </div>
        </div>
        <footer className="text-center text-xs text-slate-400 py-4">
          Kapitola 2 · MLST schéma saureus · CC mapovanie podľa literatúry
        </footer>
      </main>
    </div>
  );
}
