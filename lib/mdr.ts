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

export const NON_SUSCEPTIBLE_SET = new Set(['R', 'SC', 'Sc']);
export const MRSA_SCREEN_ATB_SET = new Set(['Cefoxitín', 'Oxacilín']);

// Use plain object instead of Set for compatibility
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
  if (typeof val === 'number') {
    return new Date((val - 25569) * 86400 * 1000);
  }
  return new Date(val);
}

export function processRawData(rows: RawRow[]): Isolate[] {
  // Group by isolate ID using plain object (no Map iteration needed)
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

    // Build ATB profile
    const atbProfile: Record<string, string> = {};
    for (let j = 0; j < group.length; j++) {
      atbProfile[group[j].NazovATB] = group[j].CitlivostATB;
    }

    // Check MRSA
    let isMrsa = false;
    for (let j = 0; j < group.length; j++) {
      const r = group[j];
      if (isMrsaScreen(r.NazovATB) && (r.CitlivostATB === '!' || isNonSusceptible(r.CitlivostATB))) {
        isMrsa = true;
        break;
      }
    }

    // Count resistant categories using plain object
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

export function getTopN<T extends string>(
  items: T[],
  n: number
): { name: T; count: number }[] {
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
