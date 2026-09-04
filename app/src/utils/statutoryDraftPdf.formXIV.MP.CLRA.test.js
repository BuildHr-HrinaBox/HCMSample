import {
    formatFormXIVMPCLRAMoneyPdfValue,
    isFormXIVMPCLRAMoneyLabel,
    looksLikeFormXIVMPCLRAPdfContext
  } from './statutoryDraftPdf.formXIV.MP.CLRA.js';
  
  describe('Form XIV MP CLRA PDF helpers', () => {
    test('detects Form XIV MP employment-card context only', () => {
      expect(
        looksLikeFormXIVMPCLRAPdfContext(
          ['FORM XIV', '[See rule 76]', 'Employment Card', 'Madhya Pradesh'],
          [['1', 'Wage rate (with particulars of unit in case of piece-work)', '', '', '71392']],
          'Form_XIV_MP'
        )
      ).toBe(true);
      expect(
        looksLikeFormXIVMPCLRAPdfContext(
          ['FORM XIV', '[See rule 76]', 'Employment Card', 'Karnataka'],
          [['1', 'Wage rate (with particulars of unit in case of piece-work)', '', '', '71392']],
          'Form_XIV_KA'
        )
      ).toBe(false);
    });
  
    test('formats Indian commas only for money-like Form XIV MP labels', () => {
      expect(isFormXIVMPCLRAMoneyLabel('4. Wage rate (with particulars of unit in case of piece-work)')).toBe(true);
      expect(isFormXIVMPCLRAMoneyLabel("4. Wage' rate with particulars or unit, in case of piece of work")).toBe(true);
      expect(isFormXIVMPCLRAMoneyLabel('Serial No. in the Register of workmen employed')).toBe(false);
      expect(formatFormXIVMPCLRAMoneyPdfValue('12345')).toBe('12,345');
      expect(formatFormXIVMPCLRAMoneyPdfValue('125000')).toBe('1,25,000');
      expect(formatFormXIVMPCLRAMoneyPdfValue('1250000')).toBe('12,50,000');
      expect(formatFormXIVMPCLRAMoneyPdfValue('NIL')).toBe('NIL');
    });
  });
  