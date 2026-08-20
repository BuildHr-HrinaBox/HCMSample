'use strict';

const fs = require('fs');
const path = require('path');

const pngPath = path.join(__dirname, '..', 'assets', 'vayona-logo-email.png');
const outPath = path.join(__dirname, '..', 'emailAssets.js');
const b64 = fs.readFileSync(pngPath).toString('base64');
const content = `'use strict';

/** White VAYONA ENERGY logo for dark email header (embedded for Gmail). */
module.exports = {
  VAYONA_LOGO_DATA_URI: 'data:image/png;base64,${b64}'
};
`;
fs.writeFileSync(outPath, content);
console.log('Wrote', outPath, content.length, 'chars');
