# Rezistencia Dashboard · Kramáre

Interaktívna webová aplikácia na analýzu MDR kmeňov *S. aureus* a *S. epidermidis*.

## Čo aplikácia robí

- **Nahrávanie dát** — upload .xlsx alebo .csv exportu z nemocničného systému
- **MDR kalkulácia** — podľa Magiorakos et al. 2012 (ECDC/CDC štandard)
  - Non-susceptible = R (rezistentný) + SC (stredne citlivý)
  - MRSA (cefoxitín/oxacilín screen `!`) = automaticky MDR
  - MDR = rezistencia v ≥3 ATB kategóriách
- **Interaktívne filtre** — patogén, MDR status, mesiac, oddelenie
- **Grafy** — MDR/non-MDR trend, patogény po mesiacoch, materiály, oddelenia
- **Tabuľka izolátov** — prehľad s MDR/MRSA statusom

## Lokálne spustenie

```bash
npm install
npm run dev
```

Otvorí sa na http://localhost:3000

## Deploy na Vercel

1. Pushni repozitár na GitHub
2. Na [vercel.com](https://vercel.com) klikni **New Project** → vyber repozitár
3. Framework: **Next.js** (automaticky detekovaný)
4. Klikni **Deploy**

Hotovo — aplikácia beží na `https://tvoj-projekt.vercel.app`

## Napojenie na živú databázu (neskôr)

V súbore `app/api/upload/route.ts` nahraď logiku spracovania súboru priamym query z DB:

```typescript
// Namiesto XLSX spracovania:
import { db } from '@/lib/db'; // tvoj DB klient (Supabase, Neon, Prisma...)
const rows = await db.query('SELECT * FROM mikrobiologia WHERE ...');
const isolates = processRawData(rows);
```

Funkcia `processRawData()` v `lib/mdr.ts` ostáva rovnaká.

## ATB kategórie (Magiorakos 2012 pre Staphylococcus)

| Kategória | ATB |
|---|---|
| Penicilíny | Ampicilín, Ampicilín-sulbaktám, Piperacilín-tazobaktám |
| Anti-SAST beta-laktámy | Oxacilín, Cefoxitín, Cefalotín |
| Glykopeptidy | Vankomycín, Teikoplanín |
| Oxazolidinóny | Linezolid |
| Makrolidy | Erytromycín |
| Linkosamidy | Klindamycín |
| Aminoglykozidy | Gentamicín, Tobramycín |
| Fluorochinolóny | Ciprofloxacín, Moxifloxacín, Ofloxacín |
| Tetracyklíny | Tetracyklín, Tigecyklín |
| Fenikoly | Chloramfenikol |
| Rifamycíny | Rifampicín |
| Folátové inhibítory | Trimetoprim, Trimetoprim-Sulfametoxazol |
| Fusidany | Fusidová kyselina |

## Štruktúra projektu

```
├── app/
│   ├── api/upload/route.ts   # API endpoint pre upload súboru
│   ├── page.tsx              # Hlavný dashboard
│   ├── layout.tsx
│   └── globals.css
├── lib/
│   └── mdr.ts               # MDR logika + dátové typy
└── package.json
```
