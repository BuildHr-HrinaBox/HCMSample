/**
 * Fetches checklist and statutory items for notifications.
 * - Checklist: all items with due dates (existing behavior).
 * - Statutory: all forms with due dates; for "Monthly Basis" (15th), show notification only when within 5 days before the 15th.
 * Returns { count, items } for use in the NewHome notification dropdown.
 */

const parseDueDate = (value) => {
  if (!value || typeof value !== 'string') return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  const d = new Date(trimmed);
  return Number.isNaN(d.getTime()) ? null : d;
};

const daysFromToday = (date) => {
  if (!date) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const other = new Date(date);
  other.setHours(0, 0, 0, 0);
  return Math.ceil((other - today) / (1000 * 60 * 60 * 24));
};

/** Next 15th from today (current month if today <= 15, else next month). */
const getNext15th = () => {
  const today = new Date();
  const day = today.getDate();
  const year = today.getFullYear();
  const month = today.getMonth();
  if (day <= 15) {
    return new Date(year, month, 15);
  }
  const next = new Date(year, month + 1, 15);
  return next;
};

/** True if dueDate string indicates monthly (15th) deadline. */
const isMonthlyBasis = (dueDate) => {
  if (!dueDate || typeof dueDate !== 'string') return false;
  return String(dueDate).toLowerCase().trim().includes('monthly basis');
};

/**
 * Build notification item.
 * @param {Object} opts - { id, formName, message, description, daysRemaining, source, sector?, act? }
 */
function buildItem(opts) {
  const { id, formName, message, description, daysRemaining, source, sector, act } = opts;
  return {
    id: id ?? `stat-${formName}-${Date.now()}`,
    formName: formName || (source === 'statutory' ? 'Statutory form' : 'Checklist'),
    message,
    description: description || null,
    daysRemaining: daysRemaining ?? null,
    source: source || 'checklist',
    sector: sector ?? null,
    act: act ?? null
  };
}

/** User-email constants: same as Statutory page – which user sees which sector only. */
const SHOPS_ESTABLISHMENT_ONLY_EMAIL = 'afrindinusha.j@buildhr.co.in';
const CLRA_ONLY_EMAIL = 'afrindinusha@gmail.com';
const FACTORIES_ACT_ONLY_EMAIL = 'afrinatlin@gmail.com';

/** Get act category from item (sector/act) – same logic as Statutory page. */
function getActCategory(item) {
  const normalizeString = (str) => {
    if (!str || typeof str !== 'string') return '';
    return String(str).trim().toLowerCase();
  };
  const act = normalizeString(item.act || item.Act || '');
  const sector = normalizeString(item.sector || item.Sector || '');

  if (
    act.includes('factories act') || act.includes('factory act') || act.includes('factories') || act.includes('factory') ||
    sector === 'factories act' || sector.includes('factories act') || sector.includes('factory act') || sector.includes('factories') || sector.includes('factory')
  ) return 'factories';
  if (
    act.includes('shops and establishments') || act.includes('shops and establishment') || act.includes('shop and establishment') ||
    sector === 'shops and establishment' || sector === 'shops and establishments' || sector.includes('shops and establishment') || sector.includes('shop and establishment')
  ) return 'shops_and_establishment';
  if (
    act.includes('clra') || act.includes('contract labour') || act.includes('contract labor') ||
    sector === 'clra' || sector.includes('clra') || sector.includes('contract labour') || sector.includes('contract labor')
  ) return 'clra';
  return 'other';
}

/** Filter notification items by logged-in user: only show forms for that user's sector (same as Statutory page). */
function filterByUserEmail(items, userEmail) {
  if (!items || items.length === 0) return items;
  const email = (userEmail || '').trim().toLowerCase();
  if (!email) return items;

  if (email === SHOPS_ESTABLISHMENT_ONLY_EMAIL) {
    return items.filter((item) => getActCategory(item) === 'shops_and_establishment');
  }
  if (email === CLRA_ONLY_EMAIL) {
    return items.filter((item) => getActCategory(item) === 'clra');
  }
  if (email === FACTORIES_ACT_ONLY_EMAIL) {
    return items.filter((item) => getActCategory(item) === 'factories');
  }
  return items;
}

/**
 * Fetch statutory data and build notification items.
 * All statutory forms are displayed. For "Monthly Basis" (15th): message shows "Due on 15th (in X days)";
 * notification is shown 5 days before the 15th (when 0 <= days to 15th <= 5).
 */
async function fetchStatutoryNotifications() {
  try {
    const response = await fetch('/server/statutoryreg_function/statutory', {
      method: 'GET',
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) return [];

    const data = await response.json();
    const list = Array.isArray(data?.data?.statutoryData) ? data.data.statutoryData : [];
    if (!list.length) return [];

    const next15 = getNext15th();
    const daysTo15 = daysFromToday(next15);
    const NOTIFY_DAYS_BEFORE_15 = 5; // Show as upcoming when 5 days or less before 15th

    const items = [];
    for (const row of list) {
      const id = row.id ?? row.ROWID;
      const formName = row.formName ?? row.FormName ?? '';
      const dueDate = row.dueDate ?? row.DueDate ?? '';
      const description = row.description ?? row.Description ?? null;
      const sector = row.sector ?? row.Sector ?? null;
      const act = row.act ?? row.Act ?? null;

      if (!formName) continue;

      if (isMonthlyBasis(dueDate)) {
        const daysRemaining = daysTo15 !== null && daysTo15 >= 0 ? daysTo15 : null;
        const message =
          daysRemaining === null
            ? 'Due on 15th (monthly)'
            : daysRemaining === 0
              ? 'Due today (15th)'
              : daysRemaining === 1
                ? 'Due tomorrow (15th)'
                : daysRemaining <= NOTIFY_DAYS_BEFORE_15
                  ? `Due on 15th (in ${daysRemaining} days)`
                  : `Due on 15th (monthly)`;
        items.push(
          buildItem({
            id,
            formName,
            message,
            description,
            daysRemaining: daysRemaining !== null && daysRemaining <= NOTIFY_DAYS_BEFORE_15 ? daysRemaining : null,
            source: 'statutory',
            sector,
            act
          })
        );
        continue;
      }

      const date = parseDueDate(dueDate);
      const daysRemaining = date !== null ? daysFromToday(date) : null;
      const message =
        daysRemaining === null
          ? 'Statutory form'
          : daysRemaining < 0
            ? `Overdue by ${Math.abs(daysRemaining)} day(s)`
            : daysRemaining === 0
              ? 'Due today'
              : daysRemaining === 1
                ? 'Due tomorrow'
                : `Due in ${daysRemaining} days`;

      items.push(
        buildItem({
          id,
          formName,
          message,
          description,
          daysRemaining,
          source: 'statutory',
          sector,
          act
        })
      );
    }
    return items;
  } catch (err) {
    console.error('fetchStatutoryNotifications error:', err);
    return [];
  }
}

/**
 * Build one notification item from a row (statutory or bulk shape: formName, dueDate, description, id).
 * Same due-date logic: monthly 15th (5 days before) and specific date.
 */
function rowToNotificationItem(row, source) {
  const id = row.id ?? row.ROWID;
  const formName = row.formName ?? row.FormName ?? '';
  const dueDate = row.dueDate ?? row.DueDate ?? '';
  const description = row.description ?? row.Description ?? null;
  const sector = row.sector ?? row.Sector ?? null;
  const act = row.act ?? row.Act ?? null;
  if (!formName) return null;

  const next15 = getNext15th();
  const daysTo15 = daysFromToday(next15);
  const NOTIFY_DAYS_BEFORE_15 = 5;

  if (isMonthlyBasis(dueDate)) {
    const daysRemaining = daysTo15 !== null && daysTo15 >= 0 ? daysTo15 : null;
    const message =
      daysRemaining === null
        ? 'Due on 15th (monthly)'
        : daysRemaining === 0
          ? 'Due today (15th)'
          : daysRemaining === 1
            ? 'Due tomorrow (15th)'
            : daysRemaining <= NOTIFY_DAYS_BEFORE_15
              ? `Due on 15th (in ${daysRemaining} days)`
              : `Due on 15th (monthly)`;
    return buildItem({
      id,
      formName,
      message,
      description,
      daysRemaining: daysRemaining !== null && daysRemaining <= NOTIFY_DAYS_BEFORE_15 ? daysRemaining : null,
      source,
      sector,
      act
    });
  }

  const date = parseDueDate(dueDate);
  const daysRemaining = date !== null ? daysFromToday(date) : null;
  const message =
    daysRemaining === null
      ? 'Statutory form'
      : daysRemaining < 0
        ? `Overdue by ${Math.abs(daysRemaining)} day(s)`
        : daysRemaining === 0
          ? 'Due today'
          : daysRemaining === 1
            ? 'Due tomorrow'
            : `Due in ${daysRemaining} days`;

  return buildItem({ id, formName, message, description, daysRemaining, source, sector, act });
}

/**
 * Fetch ChecklistBulk (master list of all forms) and build notification items.
 * Used so notifications show ALL forms (Form U, B, C, I, D, etc.), not only submitted ones.
 */
async function fetchBulkNotifications() {
  try {
    const response = await fetch('/server/checklistbulk_function/checklistbulk?action=getAll', {
      method: 'GET',
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) return [];

    const result = await response.json();
    const list = Array.isArray(result?.data) ? result.data : [];
    if (!list.length) return [];

    const items = [];
    for (const row of list) {
      const item = rowToNotificationItem(row, 'statutory');
      if (item) items.push(item);
    }
    return items;
  } catch (err) {
    console.error('fetchBulkNotifications error:', err);
    return [];
  }
}

/**
 * Fetch checklist notifications (items with due dates).
 */
async function fetchChecklistOnly() {
  try {
    const response = await fetch('/server/checklist_function/checklist', {
      method: 'GET',
      headers: { Accept: 'application/json' }
    });
    if (!response.ok) return [];

    const data = await response.json();
    const list = Array.isArray(data?.data?.checklistData)
      ? data.data.checklistData
      : Array.isArray(data?.checklistData)
        ? data.checklistData
        : Array.isArray(data?.data)
          ? data.data
          : [];
    if (!list.length) return [];

    return list
      .map((row) => {
        const id = row.id ?? row.ROWID;
        const formName = row.formName ?? row.FormName ?? '';
        const dueDate = row.dueDate ?? row.DueDate ?? '';
        const description = row.description ?? row.Description ?? null;
        const sector = row.sector ?? row.Sector ?? null;
        const act = row.act ?? row.Act ?? null;
        const date = parseDueDate(dueDate);
        const daysRemaining = date !== null ? daysFromToday(date) : null;

        const message =
          daysRemaining !== null
            ? daysRemaining < 0
              ? `Overdue by ${Math.abs(daysRemaining)} day(s)`
              : daysRemaining === 0
                ? 'Due today'
                : daysRemaining === 1
                  ? 'Due tomorrow'
                  : `Due in ${daysRemaining} days`
            : 'Checklist item';

        return buildItem({
          id,
          formName: formName || 'Checklist',
          message,
          description,
          daysRemaining,
          source: 'checklist',
          sector,
          act
        });
      })
      .filter((item) => item.formName);
  } catch (err) {
    console.error('fetchChecklistOnly error:', err);
    return [];
  }
}

/** Normalize form name for deduplication (trim, lowercase). */
function normFormName(name) {
  return (name || '').trim().toLowerCase();
}

/** Base form name for matching (e.g. "Form U - Register" and "Form U" both -> "form u"). */
function baseFormName(name) {
  const n = normFormName(name);
  const m = n.match(/^form\s+[a-z0-9]+/);
  return m ? m[0] : n;
}

/**
 * Fetch all notifications: show statutory forms corresponding to the logged-in user only (login-based filter).
 * - userEmail: current user's email (e.g. afrindinusha.j@buildhr.co.in) – only forms for that user's sector are shown.
 * - Statutory table + ChecklistBulk: merged; then filtered by user so e.g. afrindinusha.j@buildhr.co.in sees only Shops and Establishment forms.
 * - Checklist: added if form name not already present; also filtered by user sector when applicable.
 */
export async function fetchChecklistNotifications(userEmail) {
  try {
    const [statutoryItems, bulkItems, checklistItems] = await Promise.all([
      fetchStatutoryNotifications(),
      fetchBulkNotifications(),
      fetchChecklistOnly()
    ]);

    // Build sector map from bulk by base form name (so "Form U" and "Form U - Register" match).
    const bulkByBaseForm = new Map();
    for (const item of bulkItems) {
      const key = baseFormName(item.formName);
      if (!bulkByBaseForm.has(key) || getActCategory(item) !== 'other')
        bulkByBaseForm.set(key, { sector: item.sector ?? null, act: item.act ?? null });
    }

    // Like Statutory page: use BULK as base so every form has sector; then overlay statutory for due date/message.
    const byBaseForm = new Map();
    for (const item of bulkItems) {
      const key = baseFormName(item.formName);
      if (!byBaseForm.has(key)) byBaseForm.set(key, { ...item });
    }
    for (const item of statutoryItems) {
      const key = baseFormName(item.formName);
      const existing = byBaseForm.get(key);
      if (existing) {
        byBaseForm.set(key, {
          ...existing,
          formName: item.formName || existing.formName,
          message: item.message ?? existing.message,
          daysRemaining: item.daysRemaining ?? existing.daysRemaining,
          id: item.id ?? existing.id,
          description: item.description || existing.description,
          source: 'statutory'
        });
      } else {
        byBaseForm.set(key, { ...item });
      }
    }
    for (const item of checklistItems) {
      const key = baseFormName(item.formName);
      if (byBaseForm.has(key)) continue;
      byBaseForm.set(key, { ...item });
    }
    let combined = Array.from(byBaseForm.values());

    // Enrich any item still missing sector/act from bulk (e.g. statutory-only forms).
    combined = combined.map((item) => {
      if (getActCategory(item) !== 'other') return item;
      const meta = bulkByBaseForm.get(baseFormName(item.formName))
        || bulkByBaseForm.get(normFormName(item.formName));
      if (!meta || (!meta.sector && !meta.act)) return item;
      return {
        ...item,
        sector: item.sector || meta.sector,
        act: item.act || meta.act
      };
    });

    combined = filterByUserEmail(combined, userEmail);

    combined.sort((a, b) => {
      const da = a.daysRemaining ?? 9999;
      const db = b.daysRemaining ?? 9999;
      return da - db;
    });

    return {
      count: combined.length,
      items: combined
    };
  } catch (err) {
    console.error('fetchChecklistNotifications error:', err);
    return { count: 0, items: [] };
  }
}
