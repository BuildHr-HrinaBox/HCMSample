import ExcelJS from 'exceljs';
import {
  FORM_XV_RJ_HEADER_SPECS,
  applyFormXVRJAutofillFromSite,
  isFormXVRJWageSlipContext,
  writeFormXVRJHeaderFieldsToWorksheet,
} from './formXVRJWageSlip';
import {
  applySiteCompanyHeaderAutofill,
  isEstablishmentContractCarriedHeaderLabel,
  resolveHeaderFieldExportValue,
} from '../../utils/statutorySiteCompanyHeaders';

describe('Form XV RJ wage slip establishment fetch', () => {
  const establishmentSpec = FORM_XV_RJ_HEADER_SPECS.find((s) => s.key === 'form_xv_rj_establishment');

  it('detects Form XV_MH wage slip context', () => {
    expect(
      isFormXVRJWageSlipContext(
        { title: 'FORM XV', subtitle: 'Wage Slip' },
        { formFileName: 'Form_XV_MH_-_Maharashtra.xlsx', formName: 'Wage slip', state: 'Maharashtra' },
        'Form_XV_MH_-_Maharashtra.xlsx',
        ''
      )
    ).toBe(true);
  });

  it('matches template typo establiishment with dotted underline', () => {
    const label =
      'Name and address of establiishment in/under which contract is carried on……………………………………………………………………….';
    expect(establishmentSpec.match.test(label)).toBe(true);
    expect(isEstablishmentContractCarriedHeaderLabel(label)).toBe(true);
  });

  it('autofill overwrites sample establishment with site name and address', () => {
    const next = applyFormXVRJAutofillFromSite(
      { form_xv_rj_establishment: 'Lalbabu Singh' },
      {
        establishmentText: 'RJ-Fatehgarh-2, Jaisalmer, Rajasthan',
        contractorText: 'VAYONA ENERGY PRIVATE LIMITED',
        natureLocationText: 'RJ-Fatehgarh-2',
        principalEmployerText: 'RSEPL HYBRID POWER ONER ONE LIMITED',
      }
    );
    expect(next.form_xv_rj_establishment).toBe('RJ-Fatehgarh-2, Jaisalmer, Rajasthan');
    expect(next.form_xv_rj_contractor).toBe('VAYONA ENERGY PRIVATE LIMITED');
    expect(next.form_xv_rj_principal_employer).toBe('RSEPL HYBRID POWER ONER ONE LIMITED');
  });

  it('shared site headers fill form_xv_rj_establishment', () => {
    const next = applySiteCompanyHeaderAutofill(
      {},
      {
        site: {
          siteName: 'RJ-Fatehgarh-2',
          siteAddress: 'Sandhua, Fatehgarh',
          siteCity: 'Jaisalmer',
          siteState: 'Rajasthan',
        },
        formHeaderFields: [
          {
            key: 'form_xv_rj_establishment',
            label:
              'Name and address of establiishment in/under which contract is carried on…………………………………………',
          },
        ],
      }
    );
    expect(next.form_xv_rj_establishment).toMatch(/RJ-Fatehgarh-2/);
    expect(next.form_xv_rj_establishment).toMatch(/Rajasthan/);
  });

  it('resolveHeaderFieldExportValue aliases establishment for Form XV RJ', () => {
    const val = resolveHeaderFieldExportValue(
      { form_xv_establishment_contract_carried: 'Site Establishment, Jaipur' },
      {
        key: 'form_xv_rj_establishment',
        label: 'Name and address of establiishment in/under which contract is carried on',
      }
    );
    expect(val).toBe('Site Establishment, Jaipur');
  });

  it('writes establishment beside the dotted template label in Excel order', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell(2, 2).value = 'FORM XV';
    ws.getCell(3, 2).value = '[See Rule 77(2)(b)]';
    ws.getCell(4, 2).value = 'Wages Slip';
    ws.getCell(6, 2).value = 'Name and address of contractor';
    ws.getCell(8, 2).value = 'Name and location of work';
    ws.getCell(10, 2).value =
      'Name and address of establiishment in/under which contract is carried on……………………………………………………………………….';
    ws.getCell(12, 2).value = 'Name and address of principal employer';
    ws.getCell(14, 2).value = "Name and Father's name of the workman";
    ws.getCell(16, 2).value = 'Sex and identification token/ticket No.';
    ws.getCell(18, 2).value = 'For the week/fortnight/month';

    writeFormXVRJHeaderFieldsToWorksheet(ws, {
      form_xv_rj_contractor: 'VAYONA ENERGY PRIVATE LIMITED',
      form_xv_rj_nature_location: 'RJ-Fatehgarh-2',
      form_xv_rj_establishment: 'RJ-Fatehgarh-2, Sandhua, Fatehgarh, Jaisalmer, Rajasthan',
      form_xv_rj_principal_employer: 'RSEPL HYBRID POWER ONER ONE LIMITED',
    });

    expect(String(ws.getCell(6, 3).value || '')).toBe('VAYONA ENERGY PRIVATE LIMITED');
    expect(String(ws.getCell(8, 3).value || '')).toBe('RJ-Fatehgarh-2');
    expect(String(ws.getCell(10, 3).value || '')).toBe(
      'RJ-Fatehgarh-2, Sandhua, Fatehgarh, Jaisalmer, Rajasthan'
    );
    expect(String(ws.getCell(10, 2).value || '')).not.toMatch(/[.…]{4,}/u);
    expect(String(ws.getCell(12, 3).value || '')).toBe('RSEPL HYBRID POWER ONER ONE LIMITED');
  });
});
