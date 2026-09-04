import {
    formatFormXXIIIMPCLRAMoneyPdfValue,
    isFormXXIIIMPCLRACenterValueHeader,
    isFormXXIIIMPCLRAMoneyHeader,
    isFormXXIIIMPCLRANumberRow,
    isFormXXIIIMPCLRATableHeaderRow,
    looksLikeFormXXIIIMPCLRAPdfContext
  } from './statutoryDraftPdf.formXXIII.MP.CLRA.js';
  
  describe('Form XXIII MP CLRA PDF helpers', () => {
    test('detects MP register of overtime context only', () => {
      expect(
        looksLikeFormXXIIIMPCLRAPdfContext(
          ['FORM XXIII', '[See rule 78 (1) (a) (iii)]', 'Register of Overtime'],
          [['Name and Address of the Establishment : Dhar'], ['S. No.', 'Name of workman']],
          'Form XXIII_MP'
        )
      ).toBe(true);
      expect(
        looksLikeFormXXIIIMPCLRAPdfContext(
          ['Form XXIII', 'Register of Overtime'],
          [['S. No.', 'Name of workman']],
          'Form XXIII Tamil Nadu'
        )
      ).toBe(false);
    });
  
    test('recognizes Form XXIII MP overtime table header and numbering row', () => {
      expect(
        isFormXXIIIMPCLRATableHeaderRow([
          'S. No.',
          'Name of workman',
          "Father's/Husband's name",
          'Sex',
          'Designation/ Nature of employment',
          'Dates on which overtime worked',
          'Total overtime worked or production in case of piece rates',
          'Normal rate of wages',
          'Overtime rate of wages',
          'Overtime earnings',
          'Date on which overtime wages paid',
          'Remarks'
        ])
      ).toBe(true);
      expect(isFormXXIIIMPCLRANumberRow(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'])).toBe(
        true
      );
    });
  
    test('centers only S. No. values and formats overtime money columns', () => {
      expect(isFormXXIIIMPCLRACenterValueHeader('S. No.')).toBe(true);
      expect(isFormXXIIIMPCLRACenterValueHeader('Name of workman')).toBe(false);
      expect(isFormXXIIIMPCLRAMoneyHeader('Normal rate of wages')).toBe(true);
      expect(isFormXXIIIMPCLRAMoneyHeader('Overtime rate of wages')).toBe(true);
      expect(isFormXXIIIMPCLRAMoneyHeader('Overtime earnings')).toBe(true);
      expect(isFormXXIIIMPCLRAMoneyHeader('Date on which overtime wages paid')).toBe(false);
      expect(formatFormXXIIIMPCLRAMoneyPdfValue('12345')).toBe('12,345');
      expect(formatFormXXIIIMPCLRAMoneyPdfValue('125000')).toBe('1,25,000');
      expect(formatFormXXIIIMPCLRAMoneyPdfValue('1250000')).toBe('12,50,000');
      expect(formatFormXXIIIMPCLRAMoneyPdfValue('NIL')).toBe('NIL');
    });
  });
  
  