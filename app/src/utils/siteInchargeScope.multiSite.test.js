import {
  buildLoginInchargeSiteRecords,
  siteInchargeEmails,
  normalizeEmail
} from './siteInchargeScope';

describe('buildLoginInchargeSiteRecords multi-site', () => {
  const login = 'ashwath021104@gmail.com';

  it('returns both Rajasthan and Gujarat sites for the same login Mail Id', () => {
    const details = [
      {
        siteName: 'Fatehgarh Site',
        location: 'RJ-Fatehgarh-2',
        siteState: 'Rajasthan',
        industry: 'CLRA',
        inchargeEmail: 'ashwath021104@gmail.com'
      },
      {
        siteName: 'Fatehgarh Site',
        location: 'GJ-Maliya',
        siteState: 'Gujarat',
        industry: 'CLRA',
        inchargeEmail: '2217002@nec.edu.in, ashwath021104@gmail.com'
      }
    ];
    const recs = buildLoginInchargeSiteRecords(details, login);
    expect(recs.length).toBe(2);
    const names = recs.map((r) => r.siteName).sort();
    expect(names).toContain('GJ-Maliya');
    expect(names.some((n) => /fatehgarh|rj-fatehgarh/i.test(n))).toBe(true);
    expect(recs.map((r) => r.siteState).sort()).toEqual(['Gujarat', 'Rajasthan']);
  });

  it('splits multi-email Mail Id for matching', () => {
    expect(
      siteInchargeEmails({
        inchargeEmail: '2217002@nec.edu.in, ashwath021104@gmail.com'
      })
    ).toEqual(['2217002@nec.edu.in', 'ashwath021104@gmail.com'].map(normalizeEmail));
  });
});
