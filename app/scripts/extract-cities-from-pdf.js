/**
 * One-off: extract city names from app/src/City-List.pdf → app/src/utils/indianCities.js
 * Run: node scripts/extract-cities-from-pdf.js
 */
const fs = require('fs');
const path = require('path');
const pdf = require('pdf-parse');

const PDF_PATH = path.join(__dirname, '../src/City-List.pdf');
const OUT_PATH = path.join(__dirname, '../src/utils/indianCities.js');

const STATE_AND_SECTION_HEADERS = new Set(
  [
    'Andaman & Nicobar',
    'Islands',
    'Andhra Pradesh',
    'Arunachal Pradesh',
    'Assam',
    'Bihar',
    'Chandigarh',
    'Chandigarh Metro',
    'Chhattisgarh',
    'Dadra & Nagar Haveli',
    'Delhi',
    'Daman & Diu',
    'Goa',
    'New Delhi',
    'Gujarat',
    'Haryana',
    'Himachal Pradesh',
    'Jammu & Kashmir',
    'Jharkhand',
    'Karnataka',
    'Kerala',
    'Lakshadweep',
    'Madhya Pradesh',
    'Maharashtra',
    'Manipur',
    'Meghalaya',
    'Mizoram',
    'Nagaland',
    'Orissa',
    'Odisha',
    'Pondicherry',
    'Puducherry',
    'Punjab',
    'Rajasthan',
    'Sikkim',
    'Tamil Nadu',
    'Telangana',
    'Tripura',
    'Uttar Pradesh',
    'Uttarakhand',
    'West Bengal',
    'Travel Destinations',
    'Cities with House',
    'Addresses*',
    'CITY LIST',
    'INDIA FINDS ITS WAY WITH',
  ].map((s) => s.toLowerCase())
);

const FRAGMENT_ONLY = new Set(
  [
    'district',
    'factory',
    'industrial',
    'area',
    'township',
    'cement',
    'limited',
    'colony',
    'refinery',
    'plant',
    'sez',
    'midc',
    'metro',
    'complex',
    'centre',
    'center',
    'park',
    'growth',
    'project',
    'site',
    'camp',
    'power',
    'station',
    'thermal',
    'riico',
    'sidco',
    'gidc',
    'iidc',
    'epip',
    'apiic',
    'ina',
    'sir',
    'additional',
    'near',
    'alias',
    'the',
    'of',
    'and',
    'in',
    'at',
    'no',
    'ii',
    'iii',
    'iv',
    'v',
    'vi',
  ]
);

function shouldSkipLine(line) {
  const t = line.trim();
  if (!t) return true;
  if (/^4+$/.test(t)) return true;
  if (/^--\s*\d+/.test(t)) return true;
  if (/^Contact us/i.test(t)) return true;
  if (/^(Ph|Email|Web):/i.test(t)) return true;
  if (/mapmyindia/i.test(t)) return true;
  if (/Okhla Industrial/i.test(t)) return true;
  if (STATE_AND_SECTION_HEADERS.has(t.toLowerCase())) return true;
  if (FRAGMENT_ONLY.has(t.toLowerCase())) return true;
  if (/^\d+[\s-]*$/.test(t)) return true;
  if (t.length < 2 || t.length > 100) return true;
  return false;
}

function looksLikeCityName(name) {
  const t = name.trim();
  if (shouldSkipLine(t)) return false;
  if (!/^[A-Za-z0-9]/.test(t)) return false;
  if (!/^[A-Za-z0-9\s.,'()/&-]+$/.test(t)) return false;
  const lower = t.toLowerCase();
  if (lower.endsWith(' district') && t.length < 40) return false;
  return true;
}

function cleanCityName(name) {
  let t = String(name || '').trim();
  if (!t) return null;
  // PDF page column marker merged when tab is lost
  if (/^4[A-Za-z]/.test(t)) t = t.slice(1).trim();
  if (/^4\d/.test(t)) return null;
  if (/^\d+SGM$/i.test(t) || /^\d+STR$/i.test(t)) return null;
  return t;
}

function parsePrefixedCity(line) {
  const m = line.match(/^4[\t\s]+(.+)$/);
  if (m) return cleanCityName(m[1]);
  const glued = line.match(/^4([A-Za-z].+)$/);
  if (glued) return cleanCityName(glued[1]);
  return null;
}

async function main() {
  const buffer = fs.readFileSync(PDF_PATH);
  const data = await pdf(buffer);
  const lines = data.text.split(/\r?\n/);
  const cities = new Set();
  let inCityList = false;

  for (const raw of lines) {
    const line = raw.replace(/\f/g, '').trimEnd();
    if (/^Andaman\s*&\s*Nicobar/i.test(line) || /^CITY LIST/i.test(line)) {
      inCityList = true;
    }
    if (!inCityList) continue;

    const prefixed = parsePrefixedCity(line);
    if (prefixed && looksLikeCityName(prefixed)) {
      cities.add(prefixed);
      continue;
    }

    const trimmed = cleanCityName(line);
    if (!trimmed || shouldSkipLine(trimmed)) continue;

    if (looksLikeCityName(trimmed)) {
      cities.add(trimmed);
    }
  }

  const sorted = [...cities].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

  const out = `/** Indian cities from MapmyIndia City-List.pdf (auto-generated — run scripts/extract-cities-from-pdf.js to refresh). */
export const INDIAN_CITIES = ${JSON.stringify(sorted, null, 2)};

export default INDIAN_CITIES;
`;

  fs.writeFileSync(OUT_PATH, out, 'utf8');
  console.log(`Wrote ${sorted.length} cities to ${OUT_PATH}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
