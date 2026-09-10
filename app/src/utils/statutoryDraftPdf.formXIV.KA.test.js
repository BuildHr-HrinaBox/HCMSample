import {
    looksLikeFormXIVKarnatakaPdfContext,
    normalizeFormXIVKarnatakaPdfMatrix,
  } from './statutoryDraftPdf.formXIV.KA';
  
  describe('Form XIV Karnataka PDF normalization', () => {
    test('detects Karnataka Form XIV sheet context', () => {
      expect(
        looksLikeFormXIVKarnatakaPdfContext(
          ['FORM XIV', '[See rule 76]', 'Employment Card'],
          [['1', 'Name of the Workman', '', '', 'Ameerkhan']],
          'Form_XIV_KA'
        )
      ).toBe(true);
    });
  
    test('extracts inline Karnataka header values from the same cell', () => {
      const rows = [
        ['FORM XIV'],
        ['Employment Card'],
        ['Name and address of contractor VAYONA ENERGY PRIVATE LIMITED'],
        [
          'Name and address of Establishment in/under which contract is carried on Aditya Birla solar power limited, 132/33KV Pooli',
        ],
        ['Nature and location of work MP-Aditya Birla Dhar'],
        [
          'Name and address of Principal Employer Aditya Birla solar power limited, 132/33KV Pooling Sub Station',
        ],
        ['1 Name of the Workman Rajesh Kola'],
      ];
      const normalized = normalizeFormXIVKarnatakaPdfMatrix(rows, 1, 0);
      expect(normalized.colCount).toBe(2);
      expect(normalized.rows[0][1]).toBe('VAYONA ENERGY PRIVATE LIMITED');
      expect(normalized.rows[1][1]).toMatch(/Aditya Birla solar power limited/i);
      expect(normalized.rows[2][1]).toBe('MP-Aditya Birla Dhar');
      expect(normalized.rows[3][1]).toMatch(/Pooling Sub Station/i);
    });
  
    test('keeps remarks and signature footer, dropping trailing noise', () => {
      const rows = [
        ['1 Name of the Workman Rajesh Kola'],
        ['7 Remarks ................. On deputation'],
        ['............................ ............................'],
        ['Signature of Contractor'],
        ['Some trailing garbage'],
      ];
      const normalized = normalizeFormXIVKarnatakaPdfMatrix(rows, 1, 0);
      const labels = normalized.rows.map((r) => String(r?.[0] || '').toLowerCase());
      const remarksRow = normalized.rows.find((r) => /^7\.\s*remarks/i.test(String(r?.[0] || '')));
  
      expect(remarksRow).toBeTruthy();
      expect(String(remarksRow[1])).toBe('On deputation');
      expect(labels.some((t) => /signature of contractor/.test(t))).toBe(false);
      expect(normalized.formXIVFooterLines?.[0]).toMatch(/signature of contractor/i);
      expect(labels.some((t) => /trailing garbage/.test(t))).toBe(false);
    });
  
    test('puts continuation serial value into the right column', () => {
      const rows = [
        ['2. Serial No. in the Register of workmen employed ........................................'],
        ['VE0712'],
      ];
      const normalized = normalizeFormXIVKarnatakaPdfMatrix(rows, 1, 0);
      const serialRow = normalized.rows.find((r) => /^2\.\s*serial\s+no/i.test(String(r?.[0] || '')));
      expect(serialRow).toBeTruthy();
      expect(String(serialRow[1])).toBe('VE0712');
    });
  
    test('moves inline workman value to right column', () => {
      const rows = [
        ['1. Name of the Workman ........................................ Ameerkhan Naregal'],
      ];
      const normalized = normalizeFormXIVKarnatakaPdfMatrix(rows, 1, 0);
      const nameRow = normalized.rows.find((r) => /^1\.\s*name\s+of\s+the\s+workman/i.test(String(r?.[0] || '')));
      expect(nameRow).toBeTruthy();
      expect(String(nameRow[1])).toBe('Ameerkhan Naregal');
    });
  
    test('moves inline wage-rate value to right column', () => {
      const rows = [
        ['4. Wage rate (with particulars of unit in case of piece-work) ........................................ 74992'],
      ];
      const normalized = normalizeFormXIVKarnatakaPdfMatrix(rows, 1, 0);
      const wageRateRow = normalized.rows.find((r) => /^4\.\s*wage\s+rate/i.test(String(r?.[0] || '')));
      expect(wageRateRow).toBeTruthy();
      expect(String(wageRateRow[1])).toBe('74992');
    });
  
    test('detects official dotted Karnataka fill-in even on Sheet1', () => {
      expect(
        looksLikeFormXIVKarnatakaPdfContext(
          ['FORM XIV', '[See rule 76]', 'Employment Card'],
          [['1 Name of the Workman ........................................................................ Ashok Jangamashetti']],
          'Sheet1'
        )
      ).toBe(true);
    });
  
    test('splits official dotted KA card into 2-column label|value rows', () => {
      const rows = [
        ['FORM XIV'],
        ['[See rule 76]'],
        ['Employment Card'],
        ['Name and address if contractor: VAYONA ENERGY PRIVATE LIMITED'],
        [
          'Name and address of Establishment in/under which contract is carried on: M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation',
        ],
        ['Nature and location of work: KA-Bableshwar'],
        [
          'Name and address of Principal Employer: M/s Clean Wind Power Bableshwar Pvt Ltd, 33/220KV Substation',
        ],
        ['1 Name of the Workman ........................................................................ Ameerkhan Naregal'],
        ['2 Serial No. in the Register of workmen employed ........................................................ VE0712'],
        ['3 Nature of employment/Designation ........................................................ Engineer'],
        ['4 Wage rate (with particulars of unit in case of piece-work) ........................................................ 74992'],
        ['5 Wage period........................................................ September 2026'],
        ['6 Tenure of employment........................................................ From 01 Dec 2025'],
        ['7 Remarks........................................................'],
        ['Signature of Contractor'],
      ];
      const normalized = normalizeFormXIVKarnatakaPdfMatrix(rows, 1, 0);
      expect(normalized.formXIVKALayout).toBe(true);
      expect(normalized.colCount).toBe(2);
      expect(normalized.metaLines?.[0]).toMatch(/^FORM XIV$/i);
      expect(normalized.rows[0][0]).toMatch(/name and address if contractor/i);
      expect(normalized.rows[0][1]).toBe('VAYONA ENERGY PRIVATE LIMITED');
      expect(normalized.rows[2][1]).toBe('KA-Bableshwar');
      const nameRow = normalized.rows.find((r) => /^1\.\s*name of the workman/i.test(String(r?.[0] || '')));
      expect(nameRow[1]).toBe('Ameerkhan Naregal');
      expect(String(nameRow[0])).not.toMatch(/\.{4,}/);
      const wageRateRow = normalized.rows.find((r) => /wage\s+rate/i.test(String(r?.[0] || '')));
      expect(wageRateRow[1]).toBe('74992');
      const remarksRow = normalized.rows.find((r) => /remarks/i.test(String(r?.[0] || '')));
      expect(String(remarksRow[0])).toMatch(/\.{4,}/);
      expect(String(remarksRow[1])).toBe('');
      expect(normalized.formXIVFooterLines?.[0]).toBe('Signature of Contractor');
      expect(normalized.rows.every((r) => !/signature of contractor/i.test(String(r?.[0] || '')))).toBe(true);
    });
  });
  
  