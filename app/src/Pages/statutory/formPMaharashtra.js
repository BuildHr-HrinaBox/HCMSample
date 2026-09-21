/**
 * Maharashtra Form P — Notice of Maximum Leave Accumulated (See rule 20).
 * Same per-employee notice letter + leave table as Karnataka Form P → ZIP of employee files.
 */

import { buildFormPKarnatakaPerEmployeeDownload } from './formPKarnataka';

export { isFormPMaharashtraAccumulatedLeaveContext } from './formOGJGujarat';

export async function buildFormPMaharashtraPerEmployeeDownload(opts = {}) {
  return buildFormPKarnatakaPerEmployeeDownload({
    ...opts,
    formFileName: opts.formFileName || 'Form_P_Maharashtra.xlsx',
    fileNamePrefix: 'Form_P_Maharashtra',
  });
}
