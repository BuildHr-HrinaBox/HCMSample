import {
  applyFormMGJGujaratEmployeeToRow,
  buildFormMGJGujaratHeaderFormDataForRow,
  resolveFormMGJGujaratDisplayRows,
  resolveFormMGJGujaratTableHeaders,
} from './formMGJGujarat';

const headers = resolveFormMGJGujaratTableHeaders();
const employee = {
  FirstName: 'Shailesh',
  LastName: 'Babaria',
  Mobile: '7999401694',
};

describe('Form M GJ Contact No.', () => {
  it('does not inject employee Mobile into a blank imported row', () => {
    const row = applyFormMGJGujaratEmployeeToRow({}, employee, headers);

    expect(row['Contact No.'] || '').toBe('');
  });

  it('keeps Contact No. blank while remapping rows with employee data', () => {
    const [row] = resolveFormMGJGujaratDisplayRows(
      [{ 'The full name and address of the worker': 'Shailesh Babaria' }],
      headers,
      headers,
      [employee]
    );

    expect(row['Contact No.']).toBe('');
  });

  it('exports only a contact explicitly present in the row', () => {
    const blank = buildFormMGJGujaratHeaderFormDataForRow(
      {},
      { 'The full name and address of the worker': 'Shailesh Babaria' },
      headers,
      { empItem: employee }
    );
    const imported = buildFormMGJGujaratHeaderFormDataForRow(
      {},
      { 'Contact No.': '9876543210' },
      headers,
      { empItem: employee }
    );

    expect(blank.form_m_gj_contact_no).toBe('');
    expect(imported.form_m_gj_contact_no).toBe('9876543210');
  });
});
