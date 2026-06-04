'use client';

import { useState, useCallback, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, CartesianGrid
} from 'recharts';
import { Upload, RefreshCw, AlertCircle, Dna, FlaskConical, ChevronLeft, ChevronRight } from 'lucide-react';
import Image from 'next/image';

// ── colour palette ─────────────────────────────────────────────────────
const SS_BLUE   = '#185FA5';
const SS_TEAL   = '#0F6E56';
const SS_RED    = '#A32D2D';
const SS_AMBER  = '#854F0B';
const SS_PURPLE = '#534AB7';
const SS_CORAL  = '#993C1D';
const SS_PINK   = '#993556';
const SS_GREEN  = '#3B6D11';

const ST_COLORS: Record<string,string> = {
  'ST398': SS_BLUE, 'ST398-MLSB': SS_TEAL, 'ST398-variant': SS_AMBER,
  'ST8': SS_PURPLE, 'ST15': SS_GREEN, 'ST225': '#0891b2',
  'ST2': SS_BLUE, 'ST2-BF': SS_TEAL, 'ST5': SS_PURPLE, 'ST23': SS_CORAL,
};

const BIOFILM_COLORS: Record<string,string> = {
  'Silný producent': SS_RED, 'Stredný producent': SS_AMBER,
  'Slabý producent': SS_GREEN, 'Non-producent': '#888780',
};

const CHART_COLORS = [SS_BLUE,SS_TEAL,SS_PURPLE,SS_AMBER,SS_CORAL,SS_PINK,SS_GREEN,'#0891b2'];
const pct = (n:number,d:number) => d>0?Math.round((n/d)*100):0;

// ── SA virulence factor categories ──────────────────────────────────────
const SA_VIR_CATS = {
  toxiny: {
    label: 'Toxíny', color: SS_RED,
    genes: ['VIR_pvl','VIR_tsst1','VIR_sea','VIR_seb','VIR_hlα','VIR_hlβ','VIR_lukAB'],
    desc: { VIR_pvl:'PVL — nekrotizujúce infekcie', VIR_tsst1:'TSST-1 — toxic shock', VIR_sea:'Enterotoxín A', VIR_seb:'Enterotoxín B', 'VIR_hlα':'α-hemolyzín', 'VIR_hlβ':'β-hemolyzín', VIR_lukAB:'Leukotoxín AB' }
  },
  adhezia: {
    label: 'Adhéziny / MSCRAMMs', color: SS_BLUE,
    genes: ['VIR_fnbA','VIR_fnbB','VIR_cna','VIR_sdrC'],
    desc: { VIR_fnbA:'Fibronectin-binding A', VIR_fnbB:'Fibronectin-binding B', VIR_cna:'Collagen adhesin', VIR_sdrC:'Sdr protein C' }
  },
  imunoevasion: {
    label: 'Imunoevázia', color: SS_TEAL,
    genes: ['VIR_agr','VIR_scn'],
    desc: { VIR_agr:'Agr — virulence regulator', VIR_scn:'SCIN — complement inhibitor' }
  },
  biofilm: {
    label: 'Biofilm', color: SS_AMBER,
    genes: ['VIR_ica'],
    desc: { VIR_ica:'ICA — polysacharidový biofilm' }
  },
};

// ── SE virulence factor categories ──────────────────────────────────────
const SE_VIR_CATS = {
  pnag: {
    label: 'PNAG/PIA biofilm', color: SS_BLUE,
    genes: ['VIR_icaA','VIR_icaB','VIR_icaC','VIR_icaD'],
    desc: { VIR_icaA:'IcaA — biosyntéza', VIR_icaB:'IcaB — deacetylácia', VIR_icaC:'IcaC — transport', VIR_icaD:'IcaD — biosyntéza' }
  },
  protein_bf: {
    label: 'Proteínový biofilm', color: SS_TEAL,
    genes: ['VIR_aap','VIR_bhp','VIR_embp','VIR_sesC'],
    desc: { VIR_aap:'Aap — akumulačný proteín', VIR_bhp:'Bhp — biofilm homolog', VIR_embp:'Embp — matrix binding', VIR_sesC:'SesC — povrchový proteín' }
  },
  attachment: {
    label: 'Primárna adhézia', color: SS_AMBER,
    genes: ['VIR_atlE','VIR_fbe','VIR_sdrF'],
    desc: { VIR_atlE:'AtlE — autolysin', VIR_fbe:'Fbe — fibrinogen binding', VIR_sdrF:'SdrF — Sdr proteín F' }
  },
  enzymes: {
    label: 'Enzýmy / Proteázy', color: SS_CORAL,
    genes: ['VIR_esp','VIR_gehC'],
    desc: { VIR_esp:'Esp — serinová proteáza', VIR_gehC:'GehC — lipáza' }
  },
  regulation: {
    label: 'Regulácia', color: SS_PURPLE,
    genes: ['VIR_agr'],
    desc: { VIR_agr:'Agr — virulence regulator' }
  },
};

// ── Biofilm type classification ──────────────────────────────────────────
function getBiofilmType(iso: Isolate): string {
  if (iso.patogen.includes('aureus')) {
    const hasIca = (iso as any).VIR_ica;
    const hasFnb = (iso as any).VIR_fnbA || (iso as any).VIR_fnbB;
    if (hasIca && hasFnb) return 'Zmiešaný (PNAG + proteínový)';
    if (hasIca) return 'Polysacharidový (PNAG/PIA)';
    if (hasFnb) return 'Proteínový (FnBP)';
    return 'eDNA / iný';
  } else {
    const icaCount = (['VIR_icaA','VIR_icaB','VIR_icaC','VIR_icaD'] as const)
      .filter(g => (iso as any)[g]).length;
    const hasAap = (iso as any).VIR_aap;
    const hasAtlE = (iso as any).VIR_atlE;
    if (icaCount >= 3 && hasAap) return 'Zmiešaný (PNAG + proteínový)';
    if (icaCount >= 3) return 'Polysacharidový (PNAG/PIA)';
    if (hasAap || (iso as any).VIR_bhp) return 'Proteínový (Aap/Bhp)';
    if (hasAtlE) return 'eDNA-závislý (AtlE)';
    return 'Slabý / neurčený';
  }
}

const BIOFILM_TYPE_COLORS: Record<string,string> = {
  'Zmiešaný (PNAG + proteínový)': SS_RED,
  'Polysacharidový (PNAG/PIA)': SS_BLUE,
  'Proteínový (FnBP)': SS_TEAL,
  'Proteínový (Aap/Bhp)': SS_TEAL,
  'eDNA-závislý (AtlE)': SS_PURPLE,
  'eDNA / iný': SS_PURPLE,
  'Slabý / neurčený': '#888780',
};

// ── Ward map data ──────────────────────────────────────────────────────
const WARDS_MAP: Record<string,{cx:number;cy:number}> = {
  'Kl. anest. a intenz. med.': {cx:135,cy:93},
  'Neurochirurgická kl.':      {cx:255,cy:93},
  'Chirurgická kl.':           {cx:375,cy:93},
  'Klinika úrazovej chir.':    {cx:495,cy:93},
  'I. interná kl. SZU':        {cx:135,cy:193},
  'III. interná kl. LFUK':     {cx:255,cy:193},
  'II. neurologická klinika':  {cx:375,cy:193},
  'Urologické oddelenie':      {cx:495,cy:193},
  'Klinika geriatrie':         {cx:135,cy:293},
  'KIGM -  dospelí':           {cx:255,cy:293},
  'Novorodenecké odd.':        {cx:435,cy:293},
};

// ── Types ──────────────────────────────────────────────────────────────
interface Isolate {
  PPID: string;
  Patogen: string;
  ST: string;
  CC: string;
  Oddelenie: string;
  Material: string;
  DatumOdberu: string;
  Tyzden: number;
  RezistenciaMechanizmus: string;
  Biofilm_OD570: number;
  Biofilm_Kategoria: string;
  Biofilm_Skore: number;
  [key: string]: string | number;
}

function getTopN(items:string[],n:number):{name:string;count:number}[]{
  const c:Record<string,number>={};
  items.forEach(i=>{c[i]=(c[i]||0)+1;});
  return Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,n).map(([name,count])=>({name,count}));
}

const CustomTooltip = ({active,payload,label}:{active?:boolean;payload?:{value:number;name:string;color:string}[];label?:string}) => {
  if(!active||!payload?.length)return null;
  return(
    <div style={{background:'var(--color-background-primary)',border:'0.5px solid var(--color-border-secondary)',borderRadius:8,padding:'8px 12px',fontSize:12,boxShadow:'0 2px 8px rgba(0,0,0,0.08)'}}>
      <p style={{fontWeight:500,color:'var(--color-text-primary)',marginBottom:4}}>{label}</p>
      {payload.map((p,i)=><p key={i} style={{color:p.color,margin:'2px 0'}}>{p.name}: <b>{p.value}</b></p>)}
    </div>
  );
};

export default function Chapter2() {
  const [isolates, setIsolates] = useState<Isolate[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string|null>(null);
  const [fileName, setFileName] = useState<string|null>(null);
  const [progress, setProgress] = useState<string|null>(null);
  const [tab, setTab] = useState<'stats'|'map'>('stats');
  const [filterPat, setFilterPat] = useState<'all'|'aureus'|'epidermidis'>('all');
  const [filterST, setFilterST] = useState<string|null>(null);
  const [mapWeek, setMapWeek] = useState(1);
  const [mapMode, setMapMode] = useState<'patogen'|'st'|'res'|'biofilm'>('patogen');
  const [mapPlaying, setMapPlaying] = useState(false);
  const mapTimer = useRef<ReturnType<typeof setInterval>|null>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);

  useEffect(()=>{
    try{
      const s=localStorage.getItem('ss_ch2_v2');
      const n=localStorage.getItem('ss_ch2_fn');
      if(s){setIsolates(JSON.parse(s));setFileName(n||'Cache');}
    }catch{}
  },[]);

  const handleUpload = useCallback(async(e:React.ChangeEvent<HTMLInputElement>)=>{
    const file=e.target.files?.[0];if(!file)return;
    setLoading(true);setError(null);setProgress('Načítavam...');
    try{
      const buf=await file.arrayBuffer();
      setProgress('Spracovávam...');
      const wb=XLSX.read(buf,{type:'array'});
      const rows=XLSX.utils.sheet_to_json<Isolate>(wb.Sheets[wb.SheetNames[0]],{defval:0});
      if(!rows.length)throw new Error('Prázdny súbor.');
      setIsolates(rows);setFileName(file.name);
      try{localStorage.setItem('ss_ch2_v2',JSON.stringify(rows));localStorage.setItem('ss_ch2_fn',file.name);}catch{}
    }catch(err:unknown){setError(err instanceof Error?err.message:'Chyba');}
    finally{setLoading(false);setProgress(null);e.target.value='';}
  },[]);

  const filtered = useMemo(()=>isolates.filter(iso=>{
    if(filterPat==='aureus'&&!iso.Patogen.includes('aureus'))return false;
    if(filterPat==='epidermidis'&&!iso.Patogen.includes('epidermidis'))return false;
    if(filterST&&iso.ST!==filterST)return false;
    return true;
  }),[isolates,filterPat,filterST]);

  const sa = useMemo(()=>filtered.filter(i=>i.Patogen.includes('aureus')),[filtered]);
  const se = useMemo(()=>filtered.filter(i=>i.Patogen.includes('epidermidis')),[filtered]);

  // ST data
  const saSTData = useMemo(()=>getTopN(sa.map(i=>i.ST),10),[sa]);
  const seSTData = useMemo(()=>getTopN(se.map(i=>i.ST),10),[se]);

  // Biofilm data
  const saBFData = useMemo(()=>getTopN(sa.map(i=>i.Biofilm_Kategoria),10),[sa]);
  const seBFData = useMemo(()=>getTopN(se.map(i=>i.Biofilm_Kategoria),10),[se]);

  // Biofilm types
  const saBFTypes = useMemo(()=>getTopN(sa.map(i=>getBiofilmType(i)),10),[sa]);
  const seBFTypes = useMemo(()=>getTopN(se.map(i=>getBiofilmType(i)),10),[se]);

  // Virulence summary per SA
  const saVirData = useMemo(()=>{
    if(!sa.length)return[];
    const genes=['VIR_pvl','VIR_tsst1','VIR_sea','VIR_hlα','VIR_fnbA','VIR_cna','VIR_agr','VIR_ica','VIR_lukAB','VIR_scn'];
    const labels:Record<string,string>={VIR_pvl:'PVL',VIR_tsst1:'TSST-1',VIR_sea:'SEA','VIR_hlα':'Hlα',VIR_fnbA:'FnbA',VIR_cna:'Cna',VIR_agr:'Agr',VIR_ica:'ICA',VIR_lukAB:'LukAB',VIR_scn:'SCIN'};
    return genes.map(g=>({gene:labels[g]||g,count:sa.filter(i=>i[g]==1).length,pct:pct(sa.filter(i=>i[g]==1).length,sa.length)}));
  },[sa]);

  const seVirData = useMemo(()=>{
    if(!se.length)return[];
    const genes=['VIR_icaA','VIR_icaB','VIR_aap','VIR_bhp','VIR_atlE','VIR_fbe','VIR_esp','VIR_agr','VIR_embp','VIR_sesC'];
    const labels:Record<string,string>={VIR_icaA:'IcaA',VIR_icaB:'IcaB',VIR_aap:'Aap',VIR_bhp:'Bhp',VIR_atlE:'AtlE',VIR_fbe:'Fbe',VIR_esp:'Esp',VIR_agr:'Agr',VIR_embp:'Embp',VIR_sesC:'SesC'};
    return genes.map(g=>({gene:labels[g]||g,count:se.filter(i=>i[g]==1).length,pct:pct(se.filter(i=>i[g]==1).length,se.length)}));
  },[se]);

  // Virulence by ST
  const virBySTData = useMemo(()=>{
    if(!sa.length)return[];
    const sts=['ST398','ST398-MLSB','ST398-variant','ST8','ST15'];
    return sts.map(st=>{
      const sub=sa.filter(i=>i.ST===st);
      if(!sub.length)return null;
      return {
        st, n:sub.length,
        pvl: pct(sub.filter(i=>i.VIR_pvl==1).length,sub.length),
        tsst: pct(sub.filter(i=>i.VIR_tsst1==1).length,sub.length),
        ica: pct(sub.filter(i=>i.VIR_ica==1).length,sub.length),
        agr: pct(sub.filter(i=>i.VIR_agr==1).length,sub.length),
        fnbA: pct(sub.filter(i=>i.VIR_fnbA==1).length,sub.length),
      };
    }).filter(Boolean) as {st:string;n:number;pvl:number;tsst:number;ica:number;agr:number;fnbA:number}[];
  },[sa]);

  // Available STs
  const allSTs = useMemo(()=>{
    const s=new Set(filtered.map(i=>i.ST));return Array.from(s).sort();
  },[filtered]);

  // Weekly trend
  const weeklyData = useMemo(()=>{
    const bw:Record<number,{sa:number;se:number}> = {};
    for(const iso of filtered){
      const w=iso.Tyzden;
      if(!bw[w])bw[w]={sa:0,se:0};
      if(iso.Patogen.includes('aureus'))bw[w].sa++;else bw[w].se++;
    }
    return Object.keys(bw).map(Number).sort((a,b)=>a-b).map(w=>({week:`T${w}`,sa:bw[w].sa,se:bw[w].se}));
  },[filtered]);

  // OD570 scatter per ST
  const odScatterData = useMemo(()=>{
    return filtered.slice(0,200).map(i=>({
      x: i.Biofilm_OD570, y: i.Biofilm_Skore,
      st: i.ST, pat: i.Patogen.includes('aureus')?'SA':'SE',
      ppid: i.PPID,
    }));
  },[filtered]);

  // ── MAP CANVAS ─────────────────────────────────────────────────────────
  const mapIsolates = useMemo(()=>isolates.filter(i=>i.Tyzden<=mapWeek),[isolates,mapWeek]);

  function getMapColor(iso:Isolate):string{
    if(mapMode==='patogen') return iso.Patogen.includes('aureus')?SS_BLUE:SS_CORAL;
    if(mapMode==='st') return ST_COLORS[iso.ST]||'#888780';
    if(mapMode==='res'){
      const r=iso.RezistenciaMechanizmus;
      if(r.includes('MRSA')||r.includes('MRCoNS')) return SS_RED;
      if(r.includes('MLSB')) return SS_AMBER;
      return SS_GREEN;
    }
    if(mapMode==='biofilm') return BIOFILM_COLORS[iso.Biofilm_Kategoria]||'#888780';
    return '#888780';
  }

  useEffect(()=>{
    const canvas=canvasRef.current;if(!canvas)return;
    const ctx=canvas.getContext('2d');if(!ctx)return;
    const W=660,H=420;
    canvas.width=W;canvas.height=H;

    // Background
    ctx.fillStyle='#f8fafc';ctx.fillRect(0,0,W,H);

    // Floor bands
    const FLOORS=[{y:0,h:105,label:'1.p — ICU · Chirurgické'},{y:105,h:105,label:'2.p — Interné · Neurológia'},{y:210,h:105,label:'3.p — Geriatria · KIGM'},{y:315,h:105,label:'4.p — Novorodenecké'}];
    FLOORS.forEach((f,i)=>{
      ctx.fillStyle=i%2===0?'rgba(241,245,249,0.8)':'rgba(248,250,252,0.8)';
      ctx.fillRect(0,f.y,W,f.h);
      ctx.strokeStyle='rgba(203,213,225,0.5)';ctx.lineWidth=0.5;
      ctx.beginPath();ctx.moveTo(0,f.y);ctx.lineTo(W,f.y);ctx.stroke();
      ctx.fillStyle='rgba(148,163,184,0.6)';ctx.font='8px sans-serif';
      ctx.fillText(f.label,4,f.y+11);
    });

    // Corridors
    ctx.strokeStyle='rgba(148,163,184,0.3)';ctx.lineWidth=7;ctx.lineCap='round';
    [52,157,262,367].forEach(y=>{ctx.beginPath();ctx.moveTo(10,y);ctx.lineTo(W-10,y);ctx.stroke();});
    ctx.lineWidth=5;
    [W-20,W/2].forEach(x=>{ctx.beginPath();ctx.moveTo(x,52);ctx.lineTo(x,367);ctx.stroke();});

    // Wards
    const WARD_RECTS = [
      {id:'icu',x:10,y:12,w:110,h:78,label:'ICU/KAIM',fill:'#EFF6FF',stroke:'#B5D4F4',tc:'#0C447C'},
      {id:'neurochir',x:130,y:12,w:110,h:78,label:'Neurochirurgia',fill:'#F0FDF4',stroke:'#C0DD97',tc:'#27500A'},
      {id:'chirurgia',x:250,y:12,w:110,h:78,label:'Chirurgia',fill:'#EEEDFE',stroke:'#CECBF6',tc:'#3C3489'},
      {id:'uraz',x:370,y:12,w:110,h:78,label:'Úrazová chir.',fill:'#EEEDFE',stroke:'#CECBF6',tc:'#3C3489'},
      {id:'urol',x:490,y:12,w:160,h:78,label:'Urológia',fill:'#E1F5EE',stroke:'#9FE1CB',tc:'#085041'},
      {id:'interna1',x:10,y:117,w:110,h:78,label:'I. Interná',fill:'#FAEEDA',stroke:'#FAC775',tc:'#633806'},
      {id:'interna3',x:130,y:117,w:110,h:78,label:'III. Interná',fill:'#FAEEDA',stroke:'#FAC775',tc:'#633806'},
      {id:'neurol',x:250,y:117,w:110,h:78,label:'Neurológia',fill:'#F0FDF4',stroke:'#C0DD97',tc:'#27500A'},
      {id:'geriatria',x:10,y:222,w:110,h:78,label:'Geriatria',fill:'#FBEAF0',stroke:'#F4C0D1',tc:'#72243E'},
      {id:'kigm',x:130,y:222,w:110,h:78,label:'KIGM',fill:'#FBEAF0',stroke:'#F4C0D1',tc:'#72243E'},
      {id:'novor',x:250,y:327,w:400,h:78,label:'Novorodenecké / Gynekológia',fill:'#F1F5F9',stroke:'#CBD5E1',tc:'#475569'},
    ];

    const WARD_NAME_MAP:Record<string,string> = {
      'Kl. anest. a intenz. med.':'icu','Neurochirurgická kl.':'neurochir',
      'Chirurgická kl.':'chirurgia','Klinika úrazovej chir.':'uraz',
      'I. interná kl. SZU':'interna1','III. interná kl. LFUK':'interna3',
      'II. neurologická klinika':'neurol','Urologické oddelenie':'urol',
      'Klinika geriatrie':'geriatria','KIGM -  dospelí':'kigm',
      'Novorodenecké odd.':'novor',
    };

    function rr(x:number,y:number,w:number,h:number,r:number){
      ctx.beginPath();ctx.moveTo(x+r,y);ctx.lineTo(x+w-r,y);ctx.arcTo(x+w,y,x+w,y+r,r);
      ctx.lineTo(x+w,y+h-r);ctx.arcTo(x+w,y+h,x+w-r,y+h,r);
      ctx.lineTo(x+r,y+h);ctx.arcTo(x,y+h,x,y+h-r,r);
      ctx.lineTo(x,y+r);ctx.arcTo(x,y,x+r,y,r);ctx.closePath();
    }

    WARD_RECTS.forEach(w=>{
      ctx.fillStyle=w.fill;ctx.strokeStyle=w.stroke;ctx.lineWidth=1;
      rr(w.x,w.y,w.w,w.h,6);ctx.fill();ctx.stroke();
      ctx.fillStyle=w.tc;ctx.font='500 9px sans-serif';ctx.textAlign='center';
      ctx.fillText(w.label,w.x+w.w/2,w.y+14);
      ctx.textAlign='left';
    });

    // Dots
    const byWard:Record<string,Isolate[]>={};
    mapIsolates.forEach(iso=>{
      const wid=WARD_NAME_MAP[iso.Oddelenie]||iso.Oddelenie;
      if(!byWard[wid])byWard[wid]=[];
      byWard[wid].push(iso);
    });

    WARD_RECTS.forEach(wr=>{
      const isos=byWard[wr.id];if(!isos||!isos.length)return;
      const n=isos.length;
      const cols=Math.min(n,8);const rows=Math.ceil(n/cols);
      const sp=9;
      const sx=wr.x+wr.w/2-(cols-1)*sp/2;
      const sy=wr.y+wr.h/2+4-(rows-1)*sp/2;
      isos.forEach((iso,idx)=>{
        const col=idx%cols,row=Math.floor(idx/cols);
        const x=sx+col*sp,y=sy+row*sp;
        const color=getMapColor(iso);
        const isNew=iso.Tyzden===mapWeek;
        ctx.beginPath();ctx.arc(x,y,isNew?4.5:3,0,Math.PI*2);
        ctx.fillStyle=color;ctx.globalAlpha=isNew?1:0.65;ctx.fill();
        if(isNew){ctx.strokeStyle='#fff';ctx.lineWidth=1.5;ctx.stroke();}
        ctx.globalAlpha=1;
      });
      // count badge
      ctx.fillStyle='rgba(15,23,42,0.6)';ctx.font='500 8px sans-serif';ctx.textAlign='right';
      ctx.fillText(String(n),wr.x+wr.w-3,wr.y+10);ctx.textAlign='left';
    });

  },[mapIsolates,mapWeek,mapMode]);

  // Map play
  useEffect(()=>{
    if(mapPlaying){
      if(mapWeek>=13){setMapWeek(1);}
      mapTimer.current=setInterval(()=>{
        setMapWeek(w=>{
          if(w>=13){clearInterval(mapTimer.current!);setMapPlaying(false);return 13;}
          return w+1;
        });
      },900);
    } else {
      if(mapTimer.current)clearInterval(mapTimer.current);
    }
    return()=>{if(mapTimer.current)clearInterval(mapTimer.current);};
  },[mapPlaying]);

  // ── NOTES about mutations ──────────────────────────────────────────────
  const NOTES = [
    {week:6,text:'Týždeň 6: ST398 získava Constitutive MLSB rezistenciu → ST398-MLSB klon'},
    {week:10,text:'Týždeň 10: Mutácia arcC génu (alel 99) → ST398-variant (nový ST)'},
    {week:11,text:'Týždeň 11: Nezávislý ST8 klon objaví sa na Urológii'},
    {week:4,text:'Týždeň 4: ST2-BF (silný biofilm variant) šíri sa z ICU na Neurochirurgiu'},
    {week:8,text:'Týždeň 8: ST23 epidermidis — nový klon detekovaný na ICU'},
  ];

  // ── EMPTY STATE ────────────────────────────────────────────────────────
  if(!isolates.length){
    return(
      <div style={{minHeight:'100vh',background:'linear-gradient(135deg,#0a1628,#0f3460,#1a5fa8)',display:'flex',alignItems:'center',justifyContent:'center'}}>
        <div style={{textAlign:'center',padding:'2rem'}}>
          <Image src="/logo.png" alt="StaphySearch" width={280} height={75} style={{height:60,width:'auto',margin:'0 auto 2rem'}}/>
          <div style={{background:'rgba(255,255,255,0.06)',borderRadius:20,border:'1px solid rgba(0,184,150,0.2)',padding:'2rem 2.5rem',maxWidth:460}}>
            <Dna size={32} style={{color:'#00b896',margin:'0 auto 1rem',display:'block'}}/>
            <p style={{color:'rgba(255,255,255,0.5)',fontSize:12,marginBottom:'1rem',textTransform:'uppercase',letterSpacing:'0.08em'}}>Kapitola 2 — Prospektívna štúdia</p>
            <p style={{color:'rgba(255,255,255,0.7)',fontSize:13,marginBottom:'1.5rem',lineHeight:1.7}}>Nahraj syntetický dataset (staphy_synthetic_Q1_2025.xlsx) alebo vlastné dáta zo StaphySearch.</p>
            <label style={{cursor:'pointer',display:'inline-flex',alignItems:'center',gap:8,background:'linear-gradient(135deg,#1a5fa8,#00b896)',color:'#fff',padding:'9px 22px',borderRadius:10,fontSize:13,fontWeight:500}}>
              {loading?<RefreshCw size={14} style={{animation:'spin 1s linear infinite'}}/>:<Upload size={14}/>}
              {loading?(progress||'Spracovávam...'):'Nahrať dataset (.xlsx)'}
              <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleUpload} disabled={loading}/>
            </label>
            {error&&<div style={{marginTop:12,padding:'8px 12px',background:'rgba(220,38,38,0.15)',border:'1px solid rgba(220,38,38,0.3)',borderRadius:8,color:'#fca5a5',fontSize:12}}>{error}</div>}
          </div>
        </div>
        <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }

  return(
    <div style={{minHeight:'100vh',background:'#f0f4f8'}}>
      <style>{`
        @keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
        .chip{font-size:11px;padding:4px 12px;border-radius:99px;border:0.5px solid var(--color-border-secondary);color:var(--color-text-secondary);cursor:pointer;background:var(--color-background-primary);transition:all 0.15s}
        .chip:hover{border-color:#00b896}
        .chip.on{background:#185FA5;color:#fff;border-color:#185FA5}
        .card{background:var(--color-background-primary);border-radius:14px;border:0.5px solid var(--color-border-tertiary);padding:1.1rem 1.25rem}
      `}</style>

      {/* HEADER */}
      <header style={{background:'linear-gradient(135deg,#0a1628,#0f3460)',borderBottom:'1px solid rgba(0,184,150,0.15)',position:'sticky',top:0,zIndex:100}}>
        <div style={{maxWidth:1400,margin:'0 auto',padding:'0 1.5rem',display:'flex',alignItems:'center',justifyContent:'space-between',height:56}}>
          <div style={{display:'flex',alignItems:'center',gap:20}}>
            <Image src="/logo.png" alt="StaphySearch" width={130} height={35} style={{height:30,width:'auto'}}/>
            <div style={{display:'flex',gap:3}}>
              <a href="/" style={{fontSize:11,padding:'4px 12px',borderRadius:7,border:'1px solid rgba(255,255,255,0.12)',color:'rgba(255,255,255,0.4)',textDecoration:'none'}}>← Kapitola 1</a>
              <span style={{fontSize:11,padding:'4px 12px',borderRadius:7,background:'rgba(0,184,150,0.2)',color:'#fff',border:'1px solid rgba(0,184,150,0.35)'}}>Kapitola 2</span>
              {(['stats','map'] as const).map(t=>(
                <button key={t} onClick={()=>setTab(t)}
                  style={{fontSize:11,padding:'4px 12px',borderRadius:7,border:'1px solid transparent',cursor:'pointer',background:tab===t?'rgba(255,255,255,0.12)':'transparent',color:tab===t?'#fff':'rgba(255,255,255,0.45)',borderColor:tab===t?'rgba(255,255,255,0.2)':'transparent'}}>
                  {t==='stats'?'Štatistiky':'Mapa šírenia'}
                </button>
              ))}
            </div>
          </div>
          <div style={{display:'flex',alignItems:'center',gap:10}}>
            <span style={{fontSize:10,color:'rgba(255,255,255,0.3)',fontFamily:'monospace'}}>{fileName} · {isolates.length} kmeňov</span>
            <label style={{cursor:'pointer',display:'inline-flex',alignItems:'center',gap:5,background:'rgba(255,255,255,0.08)',color:'rgba(255,255,255,0.7)',padding:'4px 10px',borderRadius:7,fontSize:11,border:'1px solid rgba(255,255,255,0.12)'}}>
              {loading?<RefreshCw size={11} style={{animation:'spin 1s linear infinite'}}/>:<Upload size={11}/>} Nahrať
              <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleUpload} disabled={loading}/>
            </label>
          </div>
        </div>
      </header>

      <main style={{maxWidth:1400,margin:'0 auto',padding:'1.25rem 1.5rem'}}>

        {/* FILTERS */}
        <div className="card" style={{marginBottom:'1rem',display:'flex',flexWrap:'wrap',gap:14,alignItems:'center'}}>
          <div style={{display:'flex',gap:6,alignItems:'center'}}>
            <span style={{fontSize:10,color:'var(--color-text-tertiary)',fontWeight:600,textTransform:'uppercase',letterSpacing:'.05em'}}>Patogén</span>
            {(['all','aureus','epidermidis'] as const).map(v=>(
              <button key={v} onClick={()=>setFilterPat(v)} className={`chip ${filterPat===v?'on':''}`}>
                {v==='all'?'Všetky':v==='aureus'?'S. aureus':'S. epidermidis'}
              </button>
            ))}
          </div>
          <div style={{display:'flex',gap:6,alignItems:'center'}}>
            <span style={{fontSize:10,color:'var(--color-text-tertiary)',fontWeight:600,textTransform:'uppercase',letterSpacing:'.05em'}}>ST typ</span>
            <button onClick={()=>setFilterST(null)} className={`chip ${!filterST?'on':''}`}>Všetky</button>
            {allSTs.map(st=>(
              <button key={st} onClick={()=>setFilterST(filterST===st?null:st)} className={`chip ${filterST===st?'on':''}`}
                style={filterST===st?{background:ST_COLORS[st]||SS_BLUE,borderColor:'transparent'}:{}}>
                {st}
              </button>
            ))}
          </div>
        </div>

        {/* ══ TAB: ŠTATISTIKY ══ */}
        {tab==='stats'&&(
          <div style={{display:'flex',flexDirection:'column',gap:'1rem'}}>

            {/* Stat cards */}
            <div style={{display:'grid',gridTemplateColumns:'repeat(5,minmax(0,1fr))',gap:10}}>
              {[
                {val:filtered.length,lbl:'Kmeňov celkom',color:SS_BLUE},
                {val:sa.length,lbl:'S. aureus',color:SS_BLUE},
                {val:se.length,lbl:'S. epidermidis',color:SS_TEAL},
                {val:filtered.filter(i=>i.RezistenciaMechanizmus.includes('MRSA')||i.RezistenciaMechanizmus.includes('MRCoNS')).length,lbl:'MRSA / MRCoNS',color:SS_RED},
                {val:filtered.filter(i=>i.Biofilm_Kategoria==='Silný producent').length,lbl:'Silný biofilm',color:SS_AMBER},
              ].map((s,i)=>(
                <div key={i} className="card">
                  <div style={{fontSize:'1.8rem',fontWeight:700,fontFamily:'monospace',color:s.color,lineHeight:1}}>{s.val}</div>
                  <div style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.04em',color:'var(--color-text-secondary)',marginTop:4}}>{s.lbl}</div>
                </div>
              ))}
            </div>

            {/* ST distribúcia */}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'1rem'}}>
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>
                  S. aureus — ST typy
                  <span style={{marginLeft:8,fontSize:9,background:'#fef9c3',color:'#78350f',padding:'1px 6px',borderRadius:4}}>⚠ ST398-variant = mutácia arcC génu (týždeň 10)</span>
                </p>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={saSTData} cx="50%" cy="50%" outerRadius={80} dataKey="count" nameKey="name"
                      label={({name,count})=>`${name}: ${count}`} labelLine={false}>
                      {saSTData.map((e,i)=><Cell key={i} fill={ST_COLORS[e.name]||CHART_COLORS[i%CHART_COLORS.length]}/>)}
                    </Pie>
                    <Tooltip content={<CustomTooltip/>}/>
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>
                  S. epidermidis — ST typy
                  <span style={{marginLeft:8,fontSize:9,background:'#fef9c3',color:'#78350f',padding:'1px 6px',borderRadius:4}}>⚠ ST2-BF = silný biofilm variant</span>
                </p>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={seSTData} cx="50%" cy="50%" outerRadius={80} dataKey="count" nameKey="name"
                      label={({name,count})=>`${name}: ${count}`} labelLine={false}>
                      {seSTData.map((e,i)=><Cell key={i} fill={ST_COLORS[e.name]||CHART_COLORS[i%CHART_COLORS.length]}/>)}
                    </Pie>
                    <Tooltip content={<CustomTooltip/>}/>
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Poznámky o mutáciách */}
            <div className="card" style={{background:'linear-gradient(135deg,rgba(26,95,168,0.05),rgba(0,184,150,0.05))'}}>
              <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:10}}>Kľúčové genomické udalosti</p>
              <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(240px,1fr))',gap:8}}>
                {NOTES.map((n,i)=>(
                  <div key={i} style={{padding:'8px 12px',borderRadius:8,background:'var(--color-background-secondary)',borderLeft:`3px solid ${SS_TEAL}`,fontSize:12,color:'var(--color-text-secondary)'}}>
                    {n.text}
                  </div>
                ))}
              </div>
            </div>

            {/* Virulenčné faktory */}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:'1rem'}}>
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>S. aureus — virulenčné faktory (%)</p>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={saVirData} layout="vertical" margin={{left:10,right:20}}>
                    <XAxis type="number" tick={{fontSize:10}} domain={[0,100]} unit="%"/>
                    <YAxis type="category" dataKey="gene" tick={{fontSize:11}} width={50}/>
                    <Tooltip formatter={(v:number)=>`${v}%`}/>
                    <Bar dataKey="pct" name="%" fill={SS_BLUE} radius={[0,4,4,0]}>
                      {saVirData.map((_,i)=><Cell key={i} fill={CHART_COLORS[i%CHART_COLORS.length]}/>)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>S. epidermidis — virulenčné faktory (%)</p>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={seVirData} layout="vertical" margin={{left:10,right:20}}>
                    <XAxis type="number" tick={{fontSize:10}} domain={[0,100]} unit="%"/>
                    <YAxis type="category" dataKey="gene" tick={{fontSize:11}} width={50}/>
                    <Tooltip formatter={(v:number)=>`${v}%`}/>
                    <Bar dataKey="pct" name="%" radius={[0,4,4,0]}>
                      {seVirData.map((_,i)=><Cell key={i} fill={CHART_COLORS[i%CHART_COLORS.length]}/>)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Vir faktory per ST */}
            {virBySTData.length>0&&(
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>S. aureus — virulenčný profil per ST typ (%)</p>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={virBySTData} margin={{top:5,right:10,left:0,bottom:5}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9"/>
                    <XAxis dataKey="st" tick={{fontSize:11}}/>
                    <YAxis tick={{fontSize:11}} unit="%"/>
                    <Tooltip content={<CustomTooltip/>}/>
                    <Bar dataKey="pvl" name="PVL" fill="#A32D2D"/>
                    <Bar dataKey="tsst" name="TSST-1" fill="#D85A30"/>
                    <Bar dataKey="ica" name="ICA" fill="#185FA5"/>
                    <Bar dataKey="agr" name="Agr" fill="#0F6E56"/>
                    <Bar dataKey="fnbA" name="FnbA" fill="#534AB7"/>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}

            {/* Biofilm */}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:'1rem'}}>
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>Biofilm — S. aureus (OD570 kategórie)</p>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={saBFData} cx="50%" cy="50%" outerRadius={65} dataKey="count" nameKey="name" label={({name,count})=>`${count}`}>
                      {saBFData.map((e,i)=><Cell key={i} fill={BIOFILM_COLORS[e.name]||CHART_COLORS[i]}/>)}
                    </Pie>
                    <Tooltip/>
                  </PieChart>
                </ResponsiveContainer>
                <div style={{display:'flex',flexWrap:'wrap',gap:6,marginTop:8}}>
                  {saBFData.map(({name,count})=>(
                    <span key={name} style={{fontSize:10,display:'flex',alignItems:'center',gap:4,color:'var(--color-text-secondary)'}}>
                      <span style={{width:8,height:8,borderRadius:'50%',background:BIOFILM_COLORS[name]||'#888',display:'inline-block'}}/>
                      {name} ({count})
                    </span>
                  ))}
                </div>
              </div>
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>Biofilm — S. epidermidis</p>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={seBFData} cx="50%" cy="50%" outerRadius={65} dataKey="count" nameKey="name" label={({name,count})=>`${count}`}>
                      {seBFData.map((e,i)=><Cell key={i} fill={BIOFILM_COLORS[e.name]||CHART_COLORS[i]}/>)}
                    </Pie>
                    <Tooltip/>
                  </PieChart>
                </ResponsiveContainer>
                <div style={{display:'flex',flexWrap:'wrap',gap:6,marginTop:8}}>
                  {seBFData.map(({name,count})=>(
                    <span key={name} style={{fontSize:10,display:'flex',alignItems:'center',gap:4,color:'var(--color-text-secondary)'}}>
                      <span style={{width:8,height:8,borderRadius:'50%',background:BIOFILM_COLORS[name]||'#888',display:'inline-block'}}/>
                      {name} ({count})
                    </span>
                  ))}
                </div>
              </div>
              <div className="card">
                <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>Typy biofilmu</p>
                {[...saBFTypes,...seBFTypes.slice(0,3)].slice(0,8).map(({name,count})=>(
                  <div key={name} style={{display:'flex',alignItems:'center',gap:8,marginBottom:6}}>
                    <span style={{width:8,height:8,borderRadius:'50%',background:BIOFILM_TYPE_COLORS[name]||'#888',display:'inline-block',flexShrink:0}}/>
                    <span style={{fontSize:11,flex:1,color:'var(--color-text-secondary)',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}} title={name}>{name}</span>
                    <span style={{fontSize:11,fontFamily:'monospace',color:'var(--color-text-primary)'}}>{count}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Weekly trend */}
            <div className="card">
              <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:12}}>Týždenný vývoj — počet nových izolátov</p>
              <ResponsiveContainer width="100%" height={180}>
                <BarChart data={weeklyData} margin={{top:5,right:5,left:-10,bottom:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9"/>
                  <XAxis dataKey="week" tick={{fontSize:10}}/>
                  <YAxis tick={{fontSize:10}}/>
                  <Tooltip content={<CustomTooltip/>}/>
                  <Bar dataKey="sa" name="S. aureus" fill={SS_BLUE} stackId="a"/>
                  <Bar dataKey="se" name="S. epidermidis" fill={SS_TEAL} stackId="a" radius={[4,4,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>

          </div>
        )}

        {/* ══ TAB: MAPA ŠÍRENIA ══ */}
        {tab==='map'&&(
          <div>
            {/* Map controls */}
            <div className="card" style={{marginBottom:'1rem',display:'flex',flexWrap:'wrap',gap:12,alignItems:'center'}}>
              <div style={{display:'flex',gap:5,alignItems:'center'}}>
                <span style={{fontSize:10,color:'var(--color-text-tertiary)',fontWeight:600,textTransform:'uppercase',letterSpacing:'.05em'}}>Farba</span>
                {([['patogen','Patogén'],['st','ST typ'],['res','Rezistencia'],['biofilm','Biofilm']] as const).map(([m,l])=>(
                  <button key={m} onClick={()=>setMapMode(m)} className={`chip ${mapMode===m?'on':''}`}>{l}</button>
                ))}
              </div>
              <div style={{marginLeft:'auto',display:'flex',alignItems:'center',gap:8}}>
                <button onClick={()=>setMapWeek(w=>Math.max(1,w-1))} className="chip" aria-label="Predchádzajúci týždeň"><ChevronLeft size={13}/></button>
                <span style={{fontSize:13,fontWeight:500,color:'var(--color-text-primary)',minWidth:70,textAlign:'center'}}>Týždeň {mapWeek}</span>
                <button onClick={()=>setMapWeek(w=>Math.min(13,w+1))} className="chip" aria-label="Nasledujúci týždeň"><ChevronRight size={13}/></button>
                <input type="range" min="1" max="13" value={mapWeek} step="1" onChange={e=>setMapWeek(+e.target.value)} style={{width:100}}/>
                <button onClick={()=>setMapPlaying(p=>!p)} className={`chip ${mapPlaying?'on':''}`} style={{display:'flex',alignItems:'center',gap:4}}>
                  {mapPlaying?<><span>⏸</span>Pause</>:<><span>▶</span>Play</>}
                </button>
              </div>
            </div>

            {/* Canvas map */}
            <div style={{position:'relative',marginBottom:'1rem'}}>
              <canvas ref={canvasRef} style={{width:'100%',borderRadius:12,border:'0.5px solid var(--color-border-tertiary)',display:'block'}}/>
            </div>

            {/* Stats row */}
            <div style={{display:'flex',gap:16,flexWrap:'wrap',marginBottom:'1rem'}}>
              {[
                {lbl:'Izolátov viditeľných',val:mapIsolates.length,color:SS_BLUE},
                {lbl:'S. aureus',val:mapIsolates.filter(i=>i.Patogen.includes('aureus')).length,color:SS_BLUE},
                {lbl:'S. epidermidis',val:mapIsolates.filter(i=>i.Patogen.includes('epidermidis')).length,color:SS_TEAL},
                {lbl:'MRSA/MRCoNS',val:mapIsolates.filter(i=>i.RezistenciaMechanizmus.includes('MRSA')||i.RezistenciaMechanizmus.includes('MRCoNS')).length,color:SS_RED},
                {lbl:'Nové tento týždeň',val:isolates.filter(i=>i.Tyzden===mapWeek).length,color:SS_AMBER},
              ].map((s,i)=>(
                <div key={i} style={{background:'var(--color-background-primary)',borderRadius:10,padding:'8px 14px',border:'0.5px solid var(--color-border-tertiary)'}}>
                  <div style={{fontSize:'1.4rem',fontWeight:700,fontFamily:'monospace',color:s.color,lineHeight:1}}>{s.val}</div>
                  <div style={{fontSize:10,color:'var(--color-text-secondary)',marginTop:2}}>{s.lbl}</div>
                </div>
              ))}
            </div>

            {/* Legend */}
            <div className="card">
              <div style={{display:'flex',flexWrap:'wrap',gap:10}}>
                {mapMode==='patogen'&&[['S. aureus',SS_BLUE],['S. epidermidis',SS_CORAL]].map(([l,c])=>(
                  <span key={l} style={{display:'flex',alignItems:'center',gap:5,fontSize:12,color:'var(--color-text-secondary)'}}>
                    <span style={{width:10,height:10,borderRadius:'50%',background:c as string,display:'inline-block'}}/>
                    {l}
                  </span>
                ))}
                {mapMode==='st'&&Object.entries(ST_COLORS).map(([st,c])=>(
                  <span key={st} style={{display:'flex',alignItems:'center',gap:5,fontSize:11,color:'var(--color-text-secondary)'}}>
                    <span style={{width:9,height:9,borderRadius:'50%',background:c,display:'inline-block'}}/>
                    {st}
                  </span>
                ))}
                {mapMode==='res'&&[['MRSA/MRCoNS',SS_RED],['MLSB',SS_AMBER],['Citlivý',SS_GREEN]].map(([l,c])=>(
                  <span key={l} style={{display:'flex',alignItems:'center',gap:5,fontSize:12,color:'var(--color-text-secondary)'}}>
                    <span style={{width:10,height:10,borderRadius:'50%',background:c as string,display:'inline-block'}}/>
                    {l}
                  </span>
                ))}
                {mapMode==='biofilm'&&Object.entries(BIOFILM_COLORS).map(([l,c])=>(
                  <span key={l} style={{display:'flex',alignItems:'center',gap:5,fontSize:11,color:'var(--color-text-secondary)'}}>
                    <span style={{width:9,height:9,borderRadius:'50%',background:c,display:'inline-block'}}/>
                    {l}
                  </span>
                ))}
                <span style={{fontSize:11,color:'var(--color-text-tertiary)',marginLeft:'auto'}}>● veľká bodka = nový tento týždeň</span>
              </div>
            </div>

            {/* Genomické udalosti timeline */}
            <div className="card" style={{marginTop:'1rem'}}>
              <p style={{fontSize:10,fontWeight:600,textTransform:'uppercase',letterSpacing:'.07em',color:'var(--color-text-tertiary)',marginBottom:10}}>Genomické udalosti — Q1 2025</p>
              <div style={{display:'flex',gap:0,overflowX:'auto',paddingBottom:4}}>
                {Array.from({length:13},(_,i)=>i+1).map(w=>{
                  const note=NOTES.find(n=>n.week===w);
                  return(
                    <div key={w} style={{minWidth:80,padding:'6px 8px',borderLeft:'1px solid var(--color-border-tertiary)',background:w===mapWeek?'rgba(26,95,168,0.08)':note?'rgba(234,179,8,0.06)':'transparent'}}>
                      <div style={{fontSize:10,fontWeight:600,color:w===mapWeek?SS_BLUE:'var(--color-text-tertiary)'}}>T{w}</div>
                      {note&&<div style={{fontSize:9,color:'#78350f',marginTop:3,lineHeight:1.3}}>{note.text.replace(`Týždeň ${w}: `,'')}</div>}
                    </div>
                  );
                })}
              </div>
            </div>

          </div>
        )}

        <footer style={{textAlign:'center',fontSize:10,color:'var(--color-text-tertiary)',padding:'1.5rem 0',marginTop:'0.5rem'}}>
          StaphySearch · Kapitola 2 · Syntetické dáta Q1 2025 · MLST: PubMLST saureus / sepidermidis schéma · Biofilm: Crystal violet assay (OD570)
        </footer>
      </main>
    </div>
  );
}

