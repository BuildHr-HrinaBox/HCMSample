const TITLE_CONCAT_SPLIT_RE =
  /\s+(?=(?:REGISTER\s+OF|OVERTIME\s+MUSTER\s+ROLL|(?<!OVERTIME\s)MUSTER\s+ROLL|SERVICE\s+CERTIFICATE|WAGE\s+SLIP|LIST\s+OF|LETTER\s+OF|NOTICE\s+OF|COMBINED\s+|\[?\(?Vide\s+Rule|\[?\(?See\b))/i;

const normalizeApHeaderLine = (value) =>
  String(value || '')
    .replace(/\s+/g, ' ')
    .trim();

export const expandApHeaderTitleParts = (raw) => {
  const text = normalizeApHeaderLine(raw);
  if (!text) return [];
  const blocks = String(raw || '')
    .split(/\n+/)
    .map((part) => normalizeApHeaderLine(part))
    .filter(Boolean);
  const source = blocks.length ? blocks : [text];
  const parts = [];
  source.forEach((blob) => {
    if (
      /^form\b/i.test(blob) &&
      blob.length > 36 &&
      (TITLE_CONCAT_SPLIT_RE.test(blob) || /register of|muster\s+roll|vide\s+rule/i.test(blob))
    ) {
      const split = blob
        .split(TITLE_CONCAT_SPLIT_RE)
        .map((part) => normalizeApHeaderLine(part))
        .filter(Boolean);
      if (split.length > 1) {
        split.forEach((part) => parts.push(part));
        return;
      }
    }
    parts.push(blob);
  });
  return parts;
};

export const collectUniqueApHeaderLines = (metaLines = [], rows = [], rowLimit = 16) => {
  const seen = new Set();
  const lines = [];
  const push = (raw) => {
    expandApHeaderTitleParts(raw).forEach((line) => {
      const key = line.toLowerCase();
      if (!line || seen.has(key)) return;
      seen.add(key);
      lines.push(line);
    });
  };
  (metaLines || []).forEach(push);
  (rows || []).slice(0, rowLimit).forEach((row) => {
    if (Array.isArray(row)) row.forEach(push);
    else push(row);
  });
  return lines;
};

export const pickExclusiveApHeaderTitles = (lines = [], specs = {}) => {
  const formRe = specs.form;
  const registerRe = specs.register;
  const ruleRe = specs.rule;
  const formLine =
    lines.find(
      (line) =>
        formRe.test(line) &&
        !registerRe.test(line) &&
        !(ruleRe && ruleRe.test(line))
    ) || specs.formFallback;
  const registerLine =
    lines.find(
      (line) =>
        registerRe.test(line) &&
        !formRe.test(line) &&
        !(ruleRe && ruleRe.test(line))
    ) || specs.registerFallback;
  let ruleLine = ruleRe
    ? lines.find(
        (line) => ruleRe.test(line) && !formRe.test(line) && !registerRe.test(line)
      )
    : '';
  if (!ruleLine && typeof specs.ruleExtract === 'function') {
    ruleLine = specs.ruleExtract(lines.join(' ')) || '';
  }
  if (!ruleLine) ruleLine = specs.ruleFallback || '';
  return [formLine, registerLine, ruleLine]
    .map((line) => normalizeApHeaderLine(line))
    .filter(Boolean)
    .filter((line, index, all) => all.findIndex((item) => item.toLowerCase() === line.toLowerCase()) === index);
};

export const isApPdfHeadingRow = (row = []) => {
  const filled = (Array.isArray(row) ? row : [])
    .map((cell) => normalizeApHeaderLine(cell))
    .filter(Boolean);
  if (!filled.length) return false;
  const unique = [...new Set(filled.map((cell) => cell.toLowerCase()))];
  if (unique.length > 2) return false;
  const blob = unique.join(' ');
  return (
    /^form\s*[-._ ]*[ivxlcdm0-9]+\b/i.test(blob) ||
    /register of |muster roll|service certificate|^wage slip\b/i.test(blob) ||
    /\[?\s*vide\s+rule/i.test(blob)
  );
};
