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

export const NON_SUSCEPTIBLE = new Set(['R', 'SC', 'Sc']);
export const MRSA_SCREEN_ATB = new Set(['Cefoxitín', 'Oxacilín']);

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
  datumOdberu: string; // ISO string
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
    // Excel serial date
    return new Date((val - 25569) * 86400 * 1000);
  }
  return new Date(val);
}

export function processRawData(rows: RawRow[]): Isolate[] {
  // Group by isolate ID (first part of CisloProtokoluOKM before space)
  const groups = new Map<string, RawRow[]>();
  
  for (const row of rows) {
    const id = String(row.CisloProtokoluOKM).split(' ')[0].trim();
    if (!groups.has(id)) groups.set(id, []);
    groups.get(id)!.push(row);
  }

  const isolates: Isolate[] = [];

  for (const [id, group] of Array.from(groups.entries())) {
    const first = group[0];
    const date = parseDate(first.DatumOdberu);

    // Build ATB profile
    const atbProfile: Record<string, string> = {};
    for (const row of group) {
      atbProfile[row.NazovATB] = row.CitlivostATB;
    }

    // Check MRSA: cefoxitin/oxacillin screen (!) or resistance
    const isMrsa = group.some(
      r => MRSA_SCREEN_ATB.has(r.NazovATB) && 
           (r.CitlivostATB === '!' || NON_SUSCEPTIBLE.has(r.CitlivostATB))
    );

    // Count resistant categories
    const resistantCats = new Set<string>();
    for (const row of group) {
      if (NON_SUSCEPTIBLE.has(row.CitlivostATB) || row.CitlivostATB === '!') {
        const cat = ATB_CATEGORIES[row.NazovATB];
        if (cat) resistantCats.add(cat);
      }
    }

    const isMdr = isMrsa || resistantCats.size >= 3;

    isolates.push({
      id,
      patogen: first.Patogen?.trim() || '',
      oddelenie: first.Oddelenie?.trim() || '',
      oddeleniekod: first['Oddelenie Kod']?.trim() || '',
      material: first.DruhMaterialu?.trim() || '',
      datumOdberu: date.toISOString(),
      mesiac: date.getMonth() + 1,
      tyzden: first.tyzden,
      diagnoza: first.Diagnoza?.trim() || '',
      isMdr,
      isMrsa,
      resistantCategories: Array.from(resistantCats),
      resistantCount: resistantCats.size,
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
  const mdr = isolates.filter(i => i.isMdr).length;
  const mrsa = isolates.filter(i => i.isMrsa).length;
  const aureus = isolates.filter(i => i.patogen.includes('aureus')).length;
  const epidermidis = isolates.filter(i => i.patogen.includes('epidermidis')).length;
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
  const counts = new Map<T, number>();
  for (const item of items) {
    counts.set(item, (counts.get(item) || 0) + 1);
  }
  return Array.from(counts.entries())
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, n);
}

export const MONTH_NAMES: Record<number, string> = {
  1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Máj', 6: 'Jún',
  7: 'Júl', 8: 'Aug', 9: 'Sep', 10: 'Okt', 11: 'Nov', 12: 'Dec',
};
