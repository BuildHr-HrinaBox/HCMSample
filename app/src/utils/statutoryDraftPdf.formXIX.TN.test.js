import {
  looksLikeFormXIXTamilNaduPdfContext,
  resolveFormXIXTamilNaduNetAmountPdfSpan,
} from './statutoryDraftPdf.formXIX.TN';
import { looksLikeFormXIXGJWageSlipPdfContext } from './statutoryDraftPdf.formXIX.GJ';
import { looksLikeFormXIXKarnatakaPdfContext } from './statutoryDraftPdf.formXIX.KA';
import { statutoryDraftPdfTestUtils } from './statutoryDraftPdf';

const { looksLikeFormXIXTamilNaduPdfContext: looksLikeTnFromUtils } = statutoryDraftPdfTestUtils;

describe('Form XIX Tamil Nadu Wage Slip PDF layout', () => {
  const tnRows = [
    ['Form XIX', '', '', ''],
    ['1[See rule 78 (1) (b)]', '', '', ''],
    ['Wage Slip', '', '', ''],
    ['Workman Code', 'VE0054', 'Date of Joining', '01 Dec 2025'],
    ['Workman Name', 'Ajikumar', 'No. of Working Days', '30'],
    ["Father's Name", 'Sivaramapillai', 'No. of Overtime', ''],
    ['UAN', '100098058247', 'Rate of Daily Wages/Piece Rate', '2235.47'],
    ['ESIC IP Number', '', 'Nature of Work', 'Assistant Manager'],
    ['Wages Computation', '', '', ''],
    ['Basic', '67064', 'Employee Provident Fund', ''],
    ['Dearness Allowance', '', 'ESIC', ''],
    ['House Rent Allowance', '32345', 'Advance/Loan', 'NIL'],
    ['Leave with Wages Including Cash in Lieu of Kinds', 'NIL', 'Labour Welfare Fund', ''],
    ['Other Allowances', '52719', 'Professional Tax', ''],
    ['Gross Wages', '152128', 'Total Wage Deductions', '0'],
    ['Net Amount of Wages Paid', '', '', '130787'],
  ];

  test('detects Tamil Nadu Form XIX from layout', () => {
    expect(looksLikeFormXIXTamilNaduPdfContext([], tnRows, 'Form_XIX_-_TamilNadu')).toBe(true);
    expect(looksLikeTnFromUtils([], tnRows, 'Sheet1', 'Form_XIX_-_TamilNadu.xlsx')).toBe(true);
  });

  test('does not treat Gujarat / Karnataka Form XIX as Tamil Nadu', () => {
    const gjRows = [
      ['FORM XIX'],
      ['[See rule 78 (2)(b)]'],
      ['Wage Slip'],
      ['Name and address of contractor', 'ACME'],
      ['1. Number of days worked', '26'],
      ['7. Net amount of wages paid', '1000'],
    ];
    expect(looksLikeFormXIXGJWageSlipPdfContext([], gjRows, 'Form_XIX_GJ_-_Gujarat')).toBe(true);
    expect(
      looksLikeFormXIXTamilNaduPdfContext([], gjRows, 'Form_XIX_GJ_-_Gujarat.xlsx')
    ).toBe(false);

    const kaRows = [
      ['FORM XIX'],
      ['[See rule 78 (1)(b)]'],
      ['Wage Slip'],
      ['No. of days worked', 'Rate of daily wages/piece - rate', 'Overtime hours'],
    ];
    expect(looksLikeFormXIXKarnatakaPdfContext([], kaRows, 'Form_XIX_KA')).toBe(true);
    expect(looksLikeFormXIXTamilNaduPdfContext([], kaRows, 'Form_XIX_KA.xlsx')).toBe(false);
  });

  test('groups Net Amount of Wages Paid label across A–C with amount in D', () => {
    const span = resolveFormXIXTamilNaduNetAmountPdfSpan(
      ['Net Amount of Wages Paid', '', '', '130787'],
      4
    );
    expect(span).toEqual({
      labelStart: 0,
      labelEnd: 2,
      amountCol: 3,
      label: 'Net Amount of Wages Paid',
      amount: '130787',
    });
  });

  test('returns null when row is not the net-amount footer', () => {
    expect(
      resolveFormXIXTamilNaduNetAmountPdfSpan(['Gross Wages', '152128', 'Total Wage Deductions', '0'], 4)
    ).toBeNull();
  });
});
