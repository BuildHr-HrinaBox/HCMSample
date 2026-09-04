const normalizeFormXXIIIMpText = (text) =>
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  
  const lowerFormXXIIIMpText = (text) => normalizeFormXXIIIMpText(text).toLowerCase();
  
  export const looksLikeFormXXIIIMPCLRAPdfContext = (metaLines = [], rows = [], sheetName = '') => {
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 28).flat(), sheetName || '']
      .map(lowerFormXXIIIMpText)
      .join(' ');
    if (!blob) return false;
    const hasForm = /form[\s._-]*xxiii(?![a-z])/i.test(blob);
    const hasOvertime = /register[\s._-]*of[\s._-]*overtime/i.test(blob);
    const hasMpContext =
      /madhya\s+pradesh/.test(blob) || /form[\s._-]*xxiii[\s._-]*mp/.test(blob) || /\bxxiii[_\s.-]*mp\b/.test(blob);
    return hasForm && hasOvertime && hasMpContext;
  };
  
  export const isFormXXIIIMPCLRATitleRow = (row = []) => {
    const blob = (Array.isArray(row) ? row : []).map(normalizeFormXXIIIMpText).filter(Boolean).join(' ');
    return (
      /form[\s._-]*xxiii/i.test(blob) ||
      /register[\s._-]*of[\s._-]*overtime/i.test(blob) ||
      /see\s+rule\s*78/i.test(blob)
    );
  };
  
  export const isFormXXIIIMPCLRAAdminRow = (row = []) => {
    const blob = lowerFormXXIIIMpText((Array.isArray(row) ? row : []).join(' '));
    return (
      /name and address of contractor/.test(blob) ||
      /nature and location of work/.test(blob) ||
      /name and address of the establishment/.test(blob) ||
      /name and address of principal employer/.test(blob)
    );
  };
  
  export const isFormXXIIIMPCLRATableHeaderRow = (row = []) => {
    const blob = lowerFormXXIIIMpText((Array.isArray(row) ? row : []).join(' '));
    return (
      /(?:^|\s)(?:sr|s)\.?\s*no\b/.test(blob) &&
      /name of workman/.test(blob) &&
      /normal\s+rate\s+of\s+wages/.test(blob) &&
      /overtime\s+earnings/.test(blob)
    );
  };
  
  export const isFormXXIIIMPCLRANumberRow = (row = []) => {
    const filled = (Array.isArray(row) ? row : [])
      .map((cell) => String(cell || '').trim())
      .filter(Boolean);
    if (filled.length < 8) return false;
    const nums = filled.filter((cell) => /^\d{1,2}$/.test(cell)).map((cell) => Number(cell));
    if (nums.length < 8) return false;
    if (nums[0] !== 1) return false;
    for (let i = 1; i < nums.length; i += 1) {
      if (nums[i] !== nums[i - 1] + 1) return false;
    }
    return nums[nums.length - 1] >= 10;
  };
  
  export const isFormXXIIIMPCLRACenterValueHeader = (headerText = '') => {
    const lower = lowerFormXXIIIMpText(headerText);
    return /(?:^|\s)(?:sr|s)\.?\s*no\b/.test(lower);
  };
  
  export const isFormXXIIIMPCLRAMoneyHeader = (headerText = '') => {
    const lower = lowerFormXXIIIMpText(headerText);
    return (
      /normal\s+rate\s+of\s+wages/.test(lower) ||
      /overtime\s+rate\s+of\s+wages/.test(lower) ||
      /overtime\s+earnings/.test(lower)
    );
  };
  
  export const formatFormXXIIIMPCLRAMoneyPdfValue = (raw) => {
    const text = String(raw ?? '').trim();
    if (!text || /^nill?$/i.test(text)) return text;
    const normalized = text.replace(/,/g, '');
    if (!/^-?\d+(\.\d+)?$/.test(normalized)) return text;
    const amount = Number(normalized);
    if (!Number.isFinite(amount)) return text;
    const decimals = normalized.includes('.') ? Math.min(normalized.split('.')[1].length, 2) : 0;
    return amount.toLocaleString('en-IN', {
      maximumFractionDigits: decimals,
      minimumFractionDigits: decimals
    });
  };
  
  