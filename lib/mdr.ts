// MDR definitions based on Magiorakos et al. 2012 (ECDC/CDC standard)

export const ATB_CATEGORIES: Record<string, string> = {
  'Ampicilín': 'Penicilíny',
  'Ampicilín-sulbaktám': 'Penicilíny',
  'Piperacilín-tazobaktám': 'Penicilíny',
  'Oxacilín': 'Anti-SAST beta-laktámy',
  'Cefoxitín': 'Anti-SAST beta-laktámy',
  'Cefalotín': 'Anti-SAST beta-laktámy',
  'Vankomycín': 'Glykopeptidy',
  'Teikoplanín': 'Glykopeptidy',
  'Linezolid': 'Oxazolidinóny',
  'Erytromycín': 'Makrolidy',
  'Klindamycín': 'Linkosamidy',
  'Gentamicín': 'Aminoglykozidy',
  'Tobramycín': 'Aminoglykozidy',
  'Ciprofloxacín': 'Fluorochinolóny',
  'Moxifloxacín': 'Fluorochinolóny',
  'Ofloxacín': 'Fluorochinolóny',
  'Tetracyklín': 'Tetracyklíny',
  'Tigecyklín': 'Tetracyklíny (glycylcyklíny)',
  'Chloramfenikol': 'Fenikoly',
  'Rifampicín': 'Rifamycíny',
  'Trimetoprim': 'Folátové inhibítory',
  'Trimetoprim-Sulfametoxazol': 'Folátové inhibítory',
  'Fusidová kyselina': 'Fusidany',
  'Mupirocín': 'Pseudomonické kyseliny',
  'Nitrofurantoín': 'Nitrofurány',
  'Bacitracín': 'Polypeptidy',
  'Neomycín-Bacitracín': 'Polypeptidy',
  'Kolistín': 'Polymyxíny',
};

function isNonSusceptible(val: string): boolean {
  return val === 'R' || val === 'SC' || val === 'Sc';
}
function isMrsaScreen(atb: string): boolean {
  return atb === 'Cefoxitín' || atb === 'Oxacilín';
}

export interface RawRow {
  rok: number;
  tyzden: number;
  CisloProtokoluOKM: string;
  Oddelenie: string;
  'Oddelenie Kod': string;
  Patogen: string;
  DruhMaterialu: string;
  RezistenciaKomentar: string;
  DatumOdberu: string | Date | number;
  DatumVysetrenia: string | Date | number;
  Diagnoza: string;
  NazovATB: string;
  CitlivostATB: string;
}

export interface Isolate {
  id: string;
  patogen: string;
  oddelenie: string;
  oddeleniekod: string;
  material: string;
  datumOdberu: string;
  rok: number;
  mesiac: number;
  tyzden: number;
  diagnoza: string;
  isMdr: boolean;
  isMrsa: boolean;
  resistantCategories: string[];
  resistantCount: number;
  atbProfile: Record<string, string>;
}

function parseDate(val: string | Date | number): Date {
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date((val - 25569) * 86400 * 1000);
  return new Date(val);
}

export function processRawData(rows: RawRow[]): Isolate[] {
  const groups: Record<string, RawRow[]> = {};
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const id = String(row.CisloProtokoluOKM).split(' ')[0].trim();
    if (!groups[id]) groups[id] = [];
    groups[id].push(row);
  }

  const ids = Object.keys(groups);
  const isolates: Isolate[] = [];

  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    const group = groups[id];
    const first = group[0];
    const date = parseDate(first.DatumOdberu);

    const atbProfile: Record<string, string> = {};
    for (let j = 0; j < group.length; j++) {
      atbProfile[group[j].NazovATB] = group[j].CitlivostATB;
    }

    let isMrsa = false;
    for (let j = 0; j < group.length; j++) {
      const r = group[j];
      if (isMrsaScreen(r.NazovATB) && (r.CitlivostATB === '!' || isNonSusceptible(r.CitlivostATB))) {
        isMrsa = true;
        break;
      }
    }

    const resistantCatsObj: Record<string, boolean> = {};
    for (let j = 0; j < group.length; j++) {
      const r = group[j];
      if (isNonSusceptible(r.CitlivostATB) || r.CitlivostATB === '!') {
        const cat = ATB_CATEGORIES[r.NazovATB];
        if (cat) resistantCatsObj[cat] = true;
      }
    }
    const resistantCategories = Object.keys(resistantCatsObj);
    const resistantCount = resistantCategories.length;
    const isMdr = isMrsa || resistantCount >= 3;

    isolates.push({
      id,
      patogen: first.Patogen ? first.Patogen.trim() : '',
      oddelenie: first.Oddelenie ? first.Oddelenie.trim() : '',
      oddeleniekod: first['Oddelenie Kod'] ? first['Oddelenie Kod'].trim() : '',
      material: first.DruhMaterialu ? first.DruhMaterialu.trim() : '',
      datumOdberu: date.toISOString(),
      rok: date.getFullYear(),
      mesiac: date.getMonth() + 1,
      tyzden: first.tyzden || 0,
      diagnoza: first.Diagnoza ? first.Diagnoza.trim() : '',
      isMdr,
      isMrsa,
      resistantCategories,
      resistantCount,
      atbProfile,
    });
  }

  return isolates;
}

export interface DashboardStats {
  total: number;
  mdr: number;
  nonMdr: number;
  mrsa: number;
  aureus: number;
  epidermidis: number;
  mdrRate: number;
}

export function computeStats(isolates: Isolate[]): DashboardStats {
  let mdr = 0, mrsa = 0, aureus = 0, epidermidis = 0;
  for (let i = 0; i < isolates.length; i++) {
    const iso = isolates[i];
    if (iso.isMdr) mdr++;
    if (iso.isMrsa) mrsa++;
    if (iso.patogen.indexOf('aureus') >= 0) aureus++;
    if (iso.patogen.indexOf('epidermidis') >= 0) epidermidis++;
  }
  return {
    total: isolates.length,
    mdr,
    nonMdr: isolates.length - mdr,
    mrsa,
    aureus,
    epidermidis,
    mdrRate: isolates.length > 0 ? Math.round((mdr / isolates.length) * 100) : 0,
  };
}

export function getTopN<T extends string>(items: T[], n: number): { name: T; count: number }[] {
  const counts: Record<string, number> = {};
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    counts[item] = (counts[item] || 0) + 1;
  }
  const keys = Object.keys(counts) as T[];
  keys.sort((a, b) => counts[b] - counts[a]);
  const result: { name: T; count: number }[] = [];
  for (let i = 0; i < Math.min(n, keys.length); i++) {
    result.push({ name: keys[i], count: counts[keys[i]] });
  }
  return result;
}

export const MONTH_NAMES: Record<number, string> = {
  1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Máj', 6: 'Jún',
  7: 'Júl', 8: 'Aug', 9: 'Sep', 10: 'Okt', 11: 'Nov', 12: 'Dec',
};

// Pandemic periods (Slovak mimoriadna situácia)
export type PandemicPeriod = 'pred' | 'pocas' | 'po' | 'all';
export const PANDEMIC_PERIODS: Record<PandemicPeriod, { label: string; from?: string; to?: string }> = {
  all:   { label: 'Všetky' },
  pred:  { label: 'Pred pandémiou', to: '2020-03-11' },
  pocas: { label: 'Počas pandémie', from: '2020-03-12', to: '2023-09-15' },
  po:    { label: 'Po pandémii', from: '2023-09-16' },
};
export function getPandemicPeriod(isoDate: string): PandemicPeriod {
  const d = isoDate.slice(0, 10);
  if (d <= '2020-03-11') return 'pred';
  if (d <= '2023-09-15') return 'pocas';
  return 'po';
}

// Hospitalization data 2019-2024 (Kramáre — ukončené hospitalizácie, lôžkové oddelenia)
export const HOSP_TOTAL: Record<number, number> = {
  2019: 19975, 2020: 16398, 2021: 15934, 2022: 16694, 2023: 15923, 2024: 14718,
};

export const HOSP_DATA: { keywords: string[]; data: Record<number, number> }[] = [
  // Generic codes (anonymized) + original names (backward compatible)
  { keywords: ['int-1', 'i. intern', 'intern kl. szu', 'i.intern'], data: {2019:1347,2020:1144,2021:1230,2022:1258,2023:1139,2024:1066} },
  { keywords: ['int-2', 'iii. intern', 'intern kl. lfuk', 'iii.intern'], data: {2019:1129,2020:1010,2021:1058,2022:895,2023:829,2024:956} },
  { keywords: ['infekt-1', 'infekt-2', 'kigm', 'infektol', 'geograf'], data: {2019:2039,2020:1465,2021:1726,2022:1999,2023:1862,2024:1865} },
  { keywords: ['neur-1', 'ii. neurolog', 'neurolog klinika', 'neurologická kl'], data: {2019:1410,2020:1117,2021:873,2022:895,2023:1046,2024:1034} },
  { keywords: ['gyn-1', 'gyn', 'pôrod', 'porod'], data: {2019:4012,2020:3266,2021:3285,2022:3414,2023:3009,2024:2283} },
  { keywords: ['chir-1', 'chirurgická kl', 'chirurgická k'], data: {2019:1711,2020:1348,2021:1292,2022:1366,2023:1333,2024:1301} },
  { keywords: ['urol-1', 'urologick'], data: {2019:1345,2020:1258,2021:1040,2022:1231,2023:1248,2024:1044} },
  { keywords: ['uraz-1', 'úrazovej', 'urazovej', 'úraz.chir', 'uraz.chir'], data: {2019:1771,2020:1517,2021:1518,2022:1684,2023:1719,2024:1670} },
  { keywords: ['icu-1', 'anest', 'intenz. med', 'kl. anest'], data: {2019:168,2020:186,2021:199,2022:114,2023:110,2024:125} },
  { keywords: ['neurochir-1', 'neurochirurg'], data: {2019:1193,2020:806,2021:699,2022:723,2023:743,2024:781} },
  { keywords: ['neonat-1', 'novorodeneck', 'neonat'], data: {2019:2074,2020:1787,2021:1824,2022:1860,2023:1725,2024:1245} },
  { keywords: ['ger-1', 'geriatri'], data: {2019:1149,2020:961,2021:600,2022:621,2023:624,2024:736} },
  { keywords: ['nefr-1', 'nefrol', 'transpl'], data: {2024:362} },
  { keywords: ['odch-1', 'dlhodobo'], data: {2019:233,2020:259,2021:294,2022:252,2023:256,2024:250} },
];

export function getHospCount(oddelenie: string, rok: number): { count: number; isApprox: boolean } {
  const lower = oddelenie.toLowerCase();
  const match = HOSP_DATA.find(d => d.keywords.some(k => lower.includes(k)));
  if (match && match.data[rok]) return { count: match.data[rok], isApprox: false };
  const total = HOSP_TOTAL[rok];
  return { count: total || 0, isApprox: true };
}

export function ratePer1000(isolates: number, hosp: number): number {
  if (hosp === 0) return 0;
  return Math.round((isolates / hosp) * 1000 * 10) / 10;
}

// EARS-Net EU benchmarks (MRSA % of S. aureus invasive isolates)
export const EARS_NET: Record<number, number> = {
  2016: 16.8, 2017: 15.5, 2018: 14.4, 2019: 13.1,
  2020: 12.8, 2021: 12.5, 2022: 11.8, 2023: 10.9,
};
