import { toFormLGJPersonNameDisplay } from './formLGJGujarat';

describe('Form L GJ worker name display', () => {
  test('title-cases first and last name parts', () => {
    expect(toFormLGJPersonNameDisplay('vikash gupta')).toBe('Vikash Gupta');
    expect(toFormLGJPersonNameDisplay('harshad dk')).toBe('Harshad Dk');
    expect(toFormLGJPersonNameDisplay('shailesh babaria')).toBe('Shailesh Babaria');
    expect(toFormLGJPersonNameDisplay('patel chandrakant virabhai')).toBe(
      'Patel Chandrakant Virabhai'
    );
  });

  test('normalizes whitespace and empty values', () => {
    expect(toFormLGJPersonNameDisplay('  vikash   gupta  ')).toBe('Vikash Gupta');
    expect(toFormLGJPersonNameDisplay('')).toBe('');
    expect(toFormLGJPersonNameDisplay(null)).toBe('');
  });
});
