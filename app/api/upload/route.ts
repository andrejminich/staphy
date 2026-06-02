import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { processRawData, RawRow } from '@/lib/mdr';

export const runtime = 'nodejs';
export const maxDuration = 30;

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File | null;

    if (!file) {
      return NextResponse.json({ error: 'Žiadny súbor nebol nahraný.' }, { status: 400 });
    }

    const ext = file.name.split('.').pop()?.toLowerCase();
    const buffer = Buffer.from(await file.arrayBuffer());

    let rows: RawRow[] = [];

    if (ext === 'xlsx' || ext === 'xls') {
      const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      rows = XLSX.utils.sheet_to_json<RawRow>(sheet, { defval: '' });
    } else if (ext === 'csv') {
      const workbook = XLSX.read(buffer, { type: 'buffer', raw: false });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      rows = XLSX.utils.sheet_to_json<RawRow>(sheet, { defval: '' });
    } else {
      return NextResponse.json({ error: 'Nepodporovaný formát. Použite .xlsx alebo .csv' }, { status: 400 });
    }

    if (rows.length === 0) {
      return NextResponse.json({ error: 'Súbor je prázdny alebo má nesprávny formát.' }, { status: 400 });
    }

    // Validate required columns
    const required = ['CisloProtokoluOKM', 'Patogen', 'NazovATB', 'CitlivostATB'];
    const cols = Object.keys(rows[0]);
    const missing = required.filter(r => !cols.includes(r));
    if (missing.length > 0) {
      return NextResponse.json({ 
        error: `Chýbajú stĺpce: ${missing.join(', ')}` 
      }, { status: 400 });
    }

    const isolates = processRawData(rows);

    return NextResponse.json({ 
      isolates,
      rowCount: rows.length,
      isolateCount: isolates.length,
    });

  } catch (err) {
    console.error('Upload error:', err);
    return NextResponse.json({ error: 'Chyba pri spracovaní súboru.' }, { status: 500 });
  }
}
