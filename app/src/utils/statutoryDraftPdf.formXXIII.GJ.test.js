import {
  FORM_XXIII_GJ_OT_PDF_HEADERS,
  formXXIIIGJOtPdfColumnWeight,
  getFormXXIIIGJOtHeaderTitles,
  isFormXXIIIGJOtAdminRow,
  isFormXXIIIGJOtCenterValueHeader,
  isFormXXIIIGJOtPreambleRow,
  isFormXXIIIGJOtTableHeaderRow,
  isFormXXIIIGJOtSpuriousPdfDataRow,
  isFormXXIIIGJOtTitleRow,
  looksLikeFormXXIIIGJOtPdfContext,
  normalizeFormXXIIIGJOtPdfMatrix,
  scrubFormXXIIIGJOtPdfDataLine,
  trimFormXXIIIGJOtPdfToTwelveColumns,
} from './statutoryDraftPdf.formXXIII.GJ';

describe('Form XXIII GJ OT PDF helpers', () => {
  it('detects Gujarat Register of Overtime context', () => {
    expect(
      looksLikeFormXXIIIGJOtPdfContext(
        ['FORM XXIII', 'Register of Overtime'],
        [['Serial No.', 'Name and surname of workmen']],
        'Sheet1',
        'Form_XXIII_GJ_-_Gujarat.xlsx'
      )
    ).toBe(true);
    expect(
      looksLikeFormXXIIIGJOtPdfContext(
        ['FORM XXIII', 'Register of Overtime'],
        [],
        'Sheet1',
        'Form_XXIII_MP.xlsx'
      )
    ).toBe(false);
  });

  it('trims to twelve OT columns', () => {
    const rows = [
      ['1', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'extra', 'more'],
    ];
    const trimmed = trimFormXXIIIGJOtPdfToTwelveColumns(rows, 14);
    expect(trimmed.colCount).toBe(12);
    expect(trimmed.rows[0]).toHaveLength(12);
    expect(trimmed.rows[0][11]).toBe('k');
  });

  it('weights name columns wider than serial', () => {
    expect(formXXIIIGJOtPdfColumnWeight('Serial No.')).toBeLessThan(
      formXXIIIGJOtPdfColumnWeight('Name and surname of workmen', 20)
    );
  });

  it('rejects Form XIV preamble as OT table header', () => {
    expect(
      isFormXXIIIGJOtPreambleRow([
        '2 Serial No. in the Register of workmen employed ..............................................',
      ])
    ).toBe(true);
    expect(
      isFormXXIIIGJOtTableHeaderRow([
        '2 Serial No. in the Register of workmen employed ..............................................',
      ])
    ).toBe(false);
    expect(
      isFormXXIIIGJOtTableHeaderRow([
        'Serial No.',
        'Name and surname of workmen',
        "Father's/Husband's name",
        'Sex',
        'Designation/ Nature of employment',
        'Dates on which overtime worked',
        'Total overtime worked or production in case of piece rates',
        'Normal rate of wages',
        'Overtime rate of wages',
        'Overtime earnings',
        'Date on which overtime wages paid',
        'Remarks',
      ])
    ).toBe(true);
  });

  it('normalizes PDF matrix to August model (meta + 12-col OT table)', () => {
    const rows = [
      ['FORM XXIII', '', '', '', '', '', '', '', '', '', '', ''],
      ['1[See rule 78 (1) (a) (iii)]', '', '', '', '', '', '', '', '', '', '', ''],
      ['Register of Overtime', '', '', '', '', '', '', '', '', '', '', ''],
      ['Name and address of Contractor: VAYONA ENERGY PRIVATE LIMITED, vvd/14'],
      ['Nature and location of work: GJ-Alfanar'],
      [
        'Name and address of establishment in/under which contract is carried on: Alfanar Site',
      ],
      ['Name and address of Principal Employer: Amreli Renewable Energy Pvt Ltd'],
      ['1 Name of the Workman ..............................................'],
      [
        'Serial No.',
        'Name and surname of workmen',
        "Father's/Husband's name",
        'Sex',
        'Designation/ Nature of employment',
        'Dates on which overtime worked',
        'Total overtime worked or production in case of piece rates',
        'Normal rate of wages',
        'Overtime rate of wages',
        'Overtime earnings',
        'Date on which overtime wages paid',
        'Remarks',
        'Date of Payment',
      ],
      ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13'],
      [
        '1',
        'Makvana Sahilbhai',
        'Dineshbhai Makvana',
        'Male',
        'Senior Engineer',
        'NIL',
        'NIL',
        '42596',
        'NIL',
        'NIL',
        'NIL',
        '',
        'leak',
      ],
      ['Signature of Contractor'],
    ];
    const normalized = normalizeFormXXIIIGJOtPdfMatrix(rows, 13, 0, []);
    expect(normalized.colCount).toBe(12);
    expect(normalized.tableStartRow).toBe(0);
    expect(normalized.metaLines[0]).toMatch(/form\s*xxiii/i);
    expect(normalized.metaLines.some((l) => /register of overtime/i.test(l))).toBe(true);
    expect(normalized.metaLines.some((l) => /contractor.*vayona/i.test(l))).toBe(true);
    expect(normalized.metaLines.some((l) => /name of the workman/i.test(l))).toBe(false);
    expect(normalized.rows[0]).toEqual(FORM_XXIII_GJ_OT_PDF_HEADERS);
    expect(normalized.rows[1]).toEqual(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '']);
    expect(normalized.rows[2][1]).toBe('Makvana Sahilbhai');
    expect(normalized.rows[2]).toHaveLength(12);
    expect(normalized.rows.some((r) => /signature of contractor/i.test(String(r[0] || '')))).toBe(
      false
    );
  });

  it('drops trailing Remarks / 12 bleed rows from PDF matrix', () => {
    const rows = [
      ['Serial No.', 'Name and surname of workmen', '', '', '', '', '', '', '', '', '', 'Remarks'],
      ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'],
      ['1', 'Worker One', 'Father', 'Male', 'Engineer', 'NIL', 'NIL', '100', 'NIL', 'NIL', 'NIL', ''],
      ['', '', '', '', '', '', '', '', '', '', '', 'Remarks'],
      ['', '', '', '', '', '', '', '', '', '', '', '12'],
    ];
    const normalized = normalizeFormXXIIIGJOtPdfMatrix(rows, 12, 0, []);
    const dataRows = normalized.rows.slice(2);
    expect(dataRows).toHaveLength(1);
    expect(dataRows[0][1]).toBe('Worker One');
    expect(isFormXXIIIGJOtSpuriousPdfDataRow(scrubFormXXIIIGJOtPdfDataLine(['', '', '', '', '', '', '', '', '', '', '', '12']))).toBe(
      true
    );
  });

  it('centers serial / sex / NIL / rate headers', () => {
    expect(isFormXXIIIGJOtCenterValueHeader('Serial No.')).toBe(true);
    expect(isFormXXIIIGJOtCenterValueHeader('Sex')).toBe(true);
    expect(isFormXXIIIGJOtCenterValueHeader('Normal rate of wages')).toBe(true);
    expect(isFormXXIIIGJOtCenterValueHeader('Overtime earnings')).toBe(true);
    expect(isFormXXIIIGJOtCenterValueHeader('Name and surname of workmen')).toBe(false);
  });

  it('builds ordered title band', () => {
    const titles = getFormXXIIIGJOtHeaderTitles(
      ['Register of Overtime', 'FORM XXIII', '1[See rule 78 (1) (a) (iii)]'],
      [],
      0
    );
    expect(titles[0]).toMatch(/form\s*xxiii/i);
    expect(titles[1]).toMatch(/see\s+rule\s*78/i);
    expect(titles[2]).toMatch(/register of overtime/i);
  });

  it('detects admin rows', () => {
    expect(
      isFormXXIIIGJOtAdminRow(['Name and address of Contractor: VAYONA ENERGY PRIVATE LIMITED'])
    ).toBe(true);
    expect(isFormXXIIIGJOtTitleRow(['FORM XXIII'])).toBe(true);
  });
});
