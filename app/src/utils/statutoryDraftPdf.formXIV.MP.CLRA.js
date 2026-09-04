const normalizeFormXIVMpText = (text) =>
    String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  
  const lowerFormXIVMpText = (text) => normalizeFormXIVMpText(text).toLowerCase();
  
  export const looksLikeFormXIVMPCLRAPdfContext = (metaLines = [], rows = [], sheetName = '') => {
    const blob = [...(metaLines || []), ...(rows || []).slice(0, 24).flat(), sheetName || '']
      .map(lowerFormXIVMpText)
      .join(' ');
    if (!blob) return false;
    const hasFormXIV =
      /form\s*xiv\b/.test(blob) ||
      /form[\s._-]*xiv[\s._-]/.test(blob) ||
      (/employment\s+card/.test(blob) && /see\s+rule\s*76/.test(blob));
    if (!hasFormXIV) return false;
    return (
      /madhya\s+pradesh/.test(blob) ||
      /form[\s._-]*xiv[\s._-]*mp/.test(blob) ||
      /form_xiv_mp/.test(blob)
    );
  };
  
  export const isFormXIVMPCLRAMoneyLabel = (labelText = '') => {
    const lower = lowerFormXIVMpText(labelText).replace(/^\d+[.)]\s*/, '');
    const compact = lower.replace(/[^a-z0-9]+/g, ' ').trim();
    if (!lower) return false;
    return (
      /wage\s*['’]?\s*rate/.test(lower) ||
      /wage\s+rate/.test(compact) ||
      /rate\s+of\s+wages?/.test(compact) ||
      /gross\s+wages?/.test(compact) ||
      /overtime\s+wages?/.test(compact) ||
      /net\s+amount/.test(compact) ||
      /deductions?/.test(compact)
    );
  };
  
  export const formatFormXIVMPCLRAMoneyPdfValue = (raw) => {
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
  