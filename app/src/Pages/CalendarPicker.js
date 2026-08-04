import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import './CalendarPicker.css';
import {
  fetchInchargeDisplayScopeFromSites,
  getActCategoryFromActSector,
  industryLabelToActCategory,
  sectorMatchesInchargeSiteIndustries,
  statesFieldMatchesInchargeSiteStates,
} from '../utils/siteInchargeScope';

const DISPLAY_STATUS = {
  YET: 'Yet to Complete',
  PENDING: 'Pending',
  APPROVED: 'Approved',
  RETURNED: 'Returned',
};

function buildPaginationItems(currentPage, totalPages) {
  if (totalPages <= 0) return [];
  if (totalPages <= 9) {
    return Array.from({ length: totalPages }, (_, i) => i + 1);
  }
  const set = new Set([1, totalPages, currentPage]);
  for (let i = currentPage - 2; i <= currentPage + 2; i += 1) {
    if (i >= 1 && i <= totalPages) set.add(i);
  }
  const sorted = [...set].sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i < sorted.length; i += 1) {
    if (i > 0 && sorted[i] - sorted[i - 1] > 1) {
      out.push('ellipsis');
    }
    out.push(sorted[i]);
  }
  return out;
}

const STATUS_LABEL = {
  [DISPLAY_STATUS.YET]: 'Yet to Submit',
  [DISPLAY_STATUS.PENDING]: 'Pending',
  [DISPLAY_STATUS.APPROVED]: 'Approved',
  [DISPLAY_STATUS.RETURNED]: 'Returned',
};

const CalendarPicker = ({ userEmail, userRole }) => {
  const navigate = useNavigate();
  // Initialize with current date and year
  const today = new Date();
  const [currentDate, setCurrentDate] = useState(new Date(today.getFullYear(), today.getMonth(), 1)); // Current month and year
  const [selectedDate, setSelectedDate] = useState(new Date(today.getFullYear(), today.getMonth(), today.getDate())); // Current date
  const [showYearPicker, setShowYearPicker] = useState(false);
  const [showTooltip, setShowTooltip] = useState(false);
  const [tooltipData, setTooltipData] = useState(null);
  const [selectedNotificationId, setSelectedNotificationId] = useState(null);
  const [activeStatusFilter, setActiveStatusFilter] = useState(null); // null = All
  const [formsPage, setFormsPage] = useState(1);
  const [error, setError] = useState(null);
  const [submittingStatutoryRowId, setSubmittingStatutoryRowId] = useState('');
  
  // State for site management data and industry filtering
  const [siteManagementData, setSiteManagementData] = useState([]);
  const [loadingSiteData, setLoadingSiteData] = useState(false);
  const [availableIndustries, setAvailableIndustries] = useState([]);
  const [selectedIndustry, setSelectedIndustry] = useState('');
  
  // State for checklist bulk data (changed from calendar bulk data)
  const [checklistBulkData, setChecklistBulkData] = useState([]);
  const [loadingCalendarData, setLoadingCalendarData] = useState(false);
  const [firstNotificationId, setFirstNotificationId] = useState(null);
  const [statutoryData, setStatutoryData] = useState([]);
  // State for statutory master data (from statutorymaster_function)
  const [statutoryMasterData, setStatutoryMasterData] = useState([]);
  const [loadingStatutoryMaster, setLoadingStatutoryMaster] = useState(false);
  const [inchargeDisplayScope, setInchargeDisplayScope] = useState({
    ready: false,
    actCategories: null,
    industryLabels: null,
    stateLabels: null,
  });

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const FORM_ICON_COLORS = ['#7c3aed', '#db2777', '#0891b2', '#16a34a', '#ea580c', '#2563eb', '#ca8a04'];

  const allowedActCategoryList = inchargeDisplayScope.actCategories;
  const hasSiteBasedScope = Array.isArray(allowedActCategoryList) && allowedActCategoryList.length > 0;
  
  // Check if user is admin (App Administrator role)
  const isAdminUser = userRole === 'App Administrator';

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setInchargeDisplayScope({
        ready: false,
        actCategories: null,
        industryLabels: null,
        stateLabels: null,
      });
      try {
        const scope = await fetchInchargeDisplayScopeFromSites(userEmail);
        if (!cancelled) {
          setInchargeDisplayScope({
            ready: true,
            actCategories: scope.actCategories,
            industryLabels: scope.industryLabels,
            stateLabels: scope.stateLabels,
          });
        }
      } catch {
        if (!cancelled) {
          setInchargeDisplayScope({
            ready: true,
            actCategories: null,
            industryLabels: null,
            stateLabels: null,
          });
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [userEmail]);

  /**
   * Site-incharge filter (aligned with Statutory / Act Description):
   * - act category from Site Management Industry (factories / shops / clra)
   * - reject sectors that clearly belong to another industry
   * - when row has a state, it must match the incharge site state(s)
   */
  const rowMatchesInchargeSiteScope = (row) => {
    if (!hasSiteBasedScope) return true;
    const act = row?.act || row?.Act || '';
    const sector = row?.sector || row?.Sector || '';
    const cat = getActCategoryFromActSector(act, sector);
    if (!allowedActCategoryList.includes(cat)) return false;

    // Sector clearly maps to a different industry (e.g. Shops row on a Factories site)
    if (String(sector || '').trim()) {
      const sectorCat =
        industryLabelToActCategory(sector) || getActCategoryFromActSector('', sector);
      if (sectorCat && sectorCat !== 'other' && !allowedActCategoryList.includes(sectorCat)) {
        return false;
      }
      const inds = inchargeDisplayScope.industryLabels;
      if (
        inds &&
        inds.length > 0 &&
        sectorCat &&
        sectorCat !== 'other' &&
        !sectorMatchesInchargeSiteIndustries(sector, inds)
      ) {
        return false;
      }
    }

    const sts = inchargeDisplayScope.stateLabels;
    if (sts && sts.length > 0) {
      const stateField = row?.states || row?.state || row?.State || '';
      // Blank state kept (checklist rows often omit it); explicit other states excluded
      if (String(stateField || '').trim() && !statesFieldMatchesInchargeSiteStates(stateField, sts)) {
        return false;
      }
    }
    return true;
  };

  // Fetch site management data to get available industries
  const fetchSiteManagementData = async () => {
    if (!isAdminUser) return;
    
    setLoadingSiteData(true);
    try {
      console.log('Fetching site management data for industry filtering...');
      const response = await fetch('/server/sitemanagement_function/sitemanagement');
      
      if (response.ok) {
        const data = await response.json();
        console.log('Site management response:', data);
        if (data.status === 'success' && data.data && data.data.siteDetails) {
          setSiteManagementData(data.data.siteDetails);
          // Extract unique industries from site details
          const siteIndustries = [...new Set(data.data.siteDetails.map(site => site.industry).filter(Boolean))];
          console.log('Available industries from site management:', siteIndustries);
          
          // Also get industries from checklist bulk data
          const calendarIndustries = checklistBulkData.length > 0 
            ? [...new Set(checklistBulkData.map(item => item.sector).filter(Boolean))]
            : [];
          console.log('Available industries from checklist bulk data:', calendarIndustries);
          
          // Combine and deduplicate industries (case-insensitive)
          const combinedIndustries = [...siteIndustries, ...calendarIndustries];
          const uniqueIndustries = [];
          const seenIndustries = new Set();
          
          combinedIndustries.forEach(industry => {
            const lowerIndustry = industry.toLowerCase();
            if (!seenIndustries.has(lowerIndustry)) {
              seenIndustries.add(lowerIndustry);
              uniqueIndustries.push(industry); // Keep original case from first occurrence
            }
          });
          
          console.log('Combined available industries:', uniqueIndustries);
          setAvailableIndustries(uniqueIndustries);
          
          // Set first industry as default if available
          if (uniqueIndustries.length > 0) setSelectedIndustry(uniqueIndustries[0]);
        }
      } else {
        console.error('Failed to fetch site management data:', response.status);
      }
    } catch (err) {
      console.error('Error fetching site management data:', err);
    } finally {
      setLoadingSiteData(false);
    }
  };

  // Fetch site management data on component mount for admin users
  useEffect(() => {
    if (isAdminUser) {
      fetchSiteManagementData();
    }
  }, [isAdminUser, checklistBulkData]);

  // Fetch checklist bulk data from backend and localStorage (for Schedule Of Submission + Calendar for all users)
  const fetchChecklistBulkData = async () => {
    setLoadingCalendarData(true);
    try {
      console.log('Fetching checklist bulk data for admin/specific user...');
      
      // First try to fetch from backend
      try {
        const response = await fetch('/server/checklistbulk_function/checklistbulk?action=getAll');
        if (response.ok) {
          const result = await response.json();
          if (result.status === 'success' && result.data && result.data.length > 0) {
            console.log('Checklist bulk data loaded from backend:', result.data.length, 'records');
            setChecklistBulkData(result.data);
            // Also update localStorage
            localStorage.setItem('checklistBulkData', JSON.stringify(result.data));
            setLoadingCalendarData(false);
            return;
          }
        }
      } catch (backendError) {
        console.warn('Failed to fetch from backend, trying localStorage:', backendError);
      }
      
      // Fallback to localStorage
      const savedData = localStorage.getItem('checklistBulkData');
      if (savedData) {
        const parsedData = JSON.parse(savedData);
        console.log('Checklist bulk data loaded from localStorage:', parsedData.length, 'records');
        setChecklistBulkData(parsedData);
      } else {
        console.log('No checklist bulk data found in localStorage');
        setChecklistBulkData([]);
      }
    } catch (err) {
      console.error('Error loading checklist bulk data:', err);
      setChecklistBulkData([]);
    } finally {
      setLoadingCalendarData(false);
    }
  };

  // Fetch checklist bulk data on component mount (Schedule Of Submission + Calendar use checklistbulk)
  useEffect(() => {
    fetchChecklistBulkData();
  }, []);

  // Fetch statutory master data from statutorymaster_function
  const fetchStatutoryMasterData = async () => {
    setLoadingStatutoryMaster(true);
    try {
      const res = await fetch('/server/statutorymaster_function/statutorymaster');
      const data = await res.json();
      if (data.status === 'success' && data.data && Array.isArray(data.data.statutoryMasterData)) {
        setStatutoryMasterData(data.data.statutoryMasterData);
      } else {
        setStatutoryMasterData([]);
      }
    } catch (err) {
      console.error('Error fetching statutory master:', err);
      setStatutoryMasterData([]);
    } finally {
      setLoadingStatutoryMaster(false);
    }
  };

  useEffect(() => {
    fetchStatutoryMasterData();
  }, []);

  // Fetch statutory transaction data to mark completed forms by draft upload
  useEffect(() => {
    const fetchStatutoryData = async () => {
      try {
        const response = await fetch('/server/statutoryreg_function/statutory', { cache: 'no-store' });
        if (response.ok) {
          const payload = await response.json();
          if (payload.status === 'success' && payload.data && Array.isArray(payload.data.statutoryData)) {
            setStatutoryData(payload.data.statutoryData);
            return;
          }
        }
      } catch (err) {
        console.error('Error fetching statutory data for calendar completion:', err);
      }

      try {
        const saved = localStorage.getItem('statutoryData');
        if (saved) {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed)) {
            setStatutoryData(parsed);
            return;
          }
        }
      } catch (parseErr) {
        console.error('Error reading statutoryData from localStorage:', parseErr);
      }

      setStatutoryData([]);
    };

    fetchStatutoryData();

    const handleStatutoryDataUpdated = () => fetchStatutoryData();
    window.addEventListener('statutoryDataUpdated', handleStatutoryDataUpdated);
    return () => {
      window.removeEventListener('statutoryDataUpdated', handleStatutoryDataUpdated);
    };
  }, []);

  // Dates from statutory master for calendar highlights (DateofBirth, DateofEntryintoService, DateofExit, etc.)
  const getStatutoryMasterDatesSet = () => {
    const set = new Set();
    statutoryMasterData.forEach((record) => {
      const dateFields = ['DateofBirth', 'DateofEntryintoService', 'DateonwhichCompletion', 'DateonwhichMadePayment', 'DateofExit'];
      dateFields.forEach((field) => {
        const val = record[field];
        if (!val) return;
        const d = new Date(val);
        if (!isNaN(d.getTime())) set.add(d.toDateString());
      });
    });
    return set;
  };

  const statutoryMasterDates = getStatutoryMasterDatesSet();

  // Canonical month keys used in calendar deadline maps (must match getDaysInMonth lookup)
  const CALENDAR_MONTH_KEYS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // Parse a single due date string to { monthKey, day } for calendar (e.g. "15-Feb" -> { monthKey: "Feb", day: "15" })
  const parseDueDateToMonthDay = (dueDate) => {
    const s = String(dueDate || '').trim();
    if (!s) return null;
    const monthAbbr = CALENDAR_MONTH_KEYS;
    // ISO: 2026-07-07 or 2026/07/07
    const iso = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (iso) {
      const monthNum = parseInt(iso[2], 10);
      const day = parseInt(iso[3], 10);
      if (day >= 1 && day <= 31 && monthNum >= 1 && monthNum <= 12)
        return { monthKey: monthAbbr[monthNum - 1], day: String(day) };
    }
    // Match "15-Feb", "15-Feb-2026", "15 Feb", "Feb 15", "15/02", "15-02", "31-Dec"
    const ddm = s.match(/^(\d{1,2})[-/](\d{1,2})(?:[-/]\d{2,4})?$/); // 15-02 or 15/02
    if (ddm) {
      const day = parseInt(ddm[1], 10);
      const monthNum = parseInt(ddm[2], 10);
      if (day >= 1 && day <= 31 && monthNum >= 1 && monthNum <= 12)
        return { monthKey: monthAbbr[monthNum - 1], day: String(day) };
    }
    const dm = s.match(/^(\d{1,2})(?:st|nd|rd|th)?[-/\s]([A-Za-z]{3,})/i); // 15-Feb, 7th Jul
    if (dm) {
      const day = parseInt(dm[1], 10);
      const mon = dm[2].toLowerCase();
      const mi = monthAbbr.findIndex(m => m.toLowerCase().startsWith(mon.substring(0, 3)));
      if (day >= 1 && day <= 31 && mi >= 0)
        return { monthKey: monthAbbr[mi], day: String(day) };
    }
    const md = s.match(/([A-Za-z]{3,})[-/\s](\d{1,2})(?:st|nd|rd|th)?/i); // Feb 15 or Jul-7
    if (md) {
      const mon = md[1].toLowerCase();
      const mi = monthAbbr.findIndex(m => m.toLowerCase().startsWith(mon.substring(0, 3)));
      const day = parseInt(md[2], 10);
      if (day >= 1 && day <= 31 && mi >= 0)
        return { monthKey: monthAbbr[mi], day: String(day) };
    }
    return null;
  };

  const extractChecklistDayOfMonth = (dueDate) => {
    if (dueDate == null || dueDate === '') return null;
    if (typeof dueDate === 'number' && Number.isInteger(dueDate) && dueDate >= 1 && dueDate <= 31) {
      return dueDate;
    }
    const s = String(dueDate).trim();
    if (/^\d{1,2}$/.test(s)) {
      const day = parseInt(s, 10);
      if (day >= 1 && day <= 31) return day;
    }
    // Excel serial day-of-month quirk: 1900-01-N
    const legacy = s.match(/^1900-01-(\d{1,2})$/);
    if (legacy) {
      const recovered = parseInt(legacy[1], 10) + 1;
      if (recovered >= 1 && recovered <= 31) return recovered;
    }
    return null;
  };

  // Convert checklist bulk data to calendar format (monthly basis + specific dates)
  const convertChecklistBulkToCalendarFormat = (bulkData) => {
    const calendarData = [];
    
    try {
      if (!Array.isArray(bulkData)) return calendarData;
      
      bulkData.forEach(item => {
        if (!item || typeof item !== 'object') return;
        
        const dueDate = item.dueDate || item.DueDate || '';
        const frequency = String(item.frequency || item.Frequency || '').toLowerCase();
        const formName = item.formName || item.FormName || 'N/A';
        const description = item.description || item.Description || 'No description available';
        const sector = item.sector || item.Sector || 'General';
        const dueDateStr = String(dueDate).toLowerCase().trim();
        const bareDay = extractChecklistDayOfMonth(dueDate);
        const isMonthlyBasis =
          frequency.includes('monthly') ||
          dueDateStr.includes('monthly') ||
          bareDay != null;
        
        if (isMonthlyBasis) {
          let dayNumber = bareDay != null ? bareDay : 15;
          if (bareDay == null) {
            const dueDateDayMatch = String(dueDate).match(/\b(\d{1,2})\b/);
            if (dueDateDayMatch) {
              const extractedDay = parseInt(dueDateDayMatch[1], 10);
              if (extractedDay >= 1 && extractedDay <= 31) dayNumber = extractedDay;
            }
          }
          CALENDAR_MONTH_KEYS.forEach(month => {
            calendarData.push({
              rule: formName,
              schedule: description,
              deadlines: { [month]: String(dayNumber) },
              industry: sector
            });
          });
        } else {
          const parsed = parseDueDateToMonthDay(dueDate);
          if (parsed) {
            calendarData.push({
              rule: formName,
              schedule: description,
              deadlines: { [parsed.monthKey]: parsed.day },
              industry: sector
            });
          }
        }
      });
    } catch (error) {
      console.error('Error converting checklist bulk data:', error);
    }
    
    return calendarData;
  };

  const checklistItemMatchesCalendarDate = (item, date) => {
    if (!item?.deadlines || !date) return false;
    const targetMonth = normalizeMonthKey(CALENDAR_MONTH_KEYS[date.getMonth()]);
    const dayStr = String(date.getDate());
    return Object.entries(item.deadlines).some(
      ([monthKey, day]) => normalizeMonthKey(monthKey) === targetMonth && String(day) === dayStr
    );
  };

  const displayChecklistBulkData = (() => {
    if (!checklistBulkData || checklistBulkData.length === 0) return [];
    if (!inchargeDisplayScope.ready) return checklistBulkData;
    if (hasSiteBasedScope) {
      return checklistBulkData.filter((item) => rowMatchesInchargeSiteScope(item));
    }
    return checklistBulkData;
  })();

  // Get base calendar data from checklist bulk only (Schedule Of Submission + Calendar)
  const getBaseCalendarData = () => {
    if (displayChecklistBulkData.length > 0) {
      const convertedData = convertChecklistBulkToCalendarFormat(displayChecklistBulkData);
      return convertedData;
    }
    return [];
    
    // Removed fallback sample data - only checklistbulk data is shown
    /*
    // Old fallback sample data - removed to only show checklistbulk data
    return [
      {
        rule: "3",
        schedule: "Permission to be obtained with relevant documents before construction, reconstruction or extension.",
        deadlines: {
          "Aug": "31"
        },
        industry: "Manufacturing"
      },
      {
        rule: "4",
        schedule: "Registration/license to be obtained before commencement of manufacturing process for the first time and application to be submitted on or before October 31st of every year for renewal.",
        deadlines: {
          "Oct": "31"
        },
        industry: "Manufacturing"
      },
      {
        rule: "12-A",
        schedule: "The notice of change of manager shall be in Form 3-A.",
        deadlines: {
          "Oct": "31"
        },
        industry: "Healthcare"
      },
      {
        rule: "4(6)",
        schedule: "The registration and license shall be obtained from the concerned authority.",
        deadlines: {
          "Dec": "15"
        },
        industry: "Finance"
      }
    ];
    */
  };

  const getFilteredCalendarData = () => getBaseCalendarData();

  const calendarData = getFilteredCalendarData();

  // Schedule Of Submission list from checklistbulk (unique forms for left panel when no date selected; filtered by sector for Shops-and-Establishment-only user)
  const scheduleItemsFromChecklistBulk = (() => {
    if (!displayChecklistBulkData || displayChecklistBulkData.length === 0) return [];
    const seen = new Set();
    const list = [];
    displayChecklistBulkData.forEach(item => {
      const formName = (item.formName || item.FormName || '').trim() || 'N/A';
      const key = formName.toLowerCase();
      if (key && !seen.has(key)) {
        seen.add(key);
        list.push({
          rule: formName,
          schedule: (item.description || item.Description || 'No description available').trim()
        });
      }
    });
    return list;
  })();

  const normalizeMonthKey = (value) => {
    const raw = String(value || '').trim().toLowerCase();
    if (!raw) return '';
    if (raw.startsWith('jan')) return 'jan';
    if (raw.startsWith('feb')) return 'feb';
    if (raw.startsWith('mar')) return 'mar';
    if (raw.startsWith('apr')) return 'apr';
    if (raw.startsWith('may')) return 'may';
    if (raw.startsWith('jun')) return 'jun';
    if (raw.startsWith('jul')) return 'jul';
    if (raw.startsWith('aug')) return 'aug';
    if (raw.startsWith('sep')) return 'sep';
    if (raw.startsWith('oct')) return 'oct';
    if (raw.startsWith('nov')) return 'nov';
    if (raw.startsWith('dec')) return 'dec';
    return raw.substring(0, 3);
  };

  const filteredStatutoryData = (() => {
    if (!Array.isArray(statutoryData) || statutoryData.length === 0) return [];
    if (!inchargeDisplayScope.ready) return statutoryData;
    if (hasSiteBasedScope) {
      // Site login: only Yet to Complete / status rows for this site's industry + state
      return statutoryData.filter((item) => rowMatchesInchargeSiteScope(item));
    }
    return statutoryData;
  })();

  const normLabel = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();

  const scheduleMatchesStatutoryDescription = (scheduleFromCard, statDescription) => {
    const a = normLabel(scheduleFromCard);
    const b = normLabel(statDescription);
    if (a && b) {
      if (a === b) return true;
      const minLen = Math.min(a.length, b.length);
      if (minLen >= 10) return a.includes(b) || b.includes(a);
      return false;
    }
    return false;
  };

  const statutoryMonthMatches = (row, targetMonth) => {
    const itemMonthFilter = row.monthFilter || row.MonthFilter || row.monthfilter || '';
    const parsedDueMonth = parseDueDateToMonthDay(row.dueDate || row.DueDate || '')?.monthKey || '';
    const itemMonth = normalizeMonthKey(itemMonthFilter || parsedDueMonth);
    if (!targetMonth) return true;
    return !itemMonth || itemMonth === targetMonth;
  };

  /** Same month matching as Statutory Transaction (incl. nomonth / monthly / bare day). */
  const statutoryRowMatchesTransactionMonth = (row, calendarDate) => {
    if (!calendarDate) return false;
    const calMonthName = monthNames[calendarDate.getMonth()];
    const selectedMonthNorm = String(calMonthName).trim().toLowerCase().substring(0, 3);
    const storedMonth = row.monthFilter || row.MonthFilter || row.monthfilter || '';
    const storedMonthNorm = String(storedMonth).trim().toLowerCase().substring(0, 3);
    if (storedMonthNorm) return storedMonthNorm === selectedMonthNorm;
    const dueRaw = row.dueDate || row.DueDate;
    const dueDateLower = String(dueRaw || '').toLowerCase().trim();
    // nomonth rows: empty, monthly, or bare day-of-month match any selected month
    if (!dueDateLower) return true;
    if (dueDateLower.includes('monthly')) return true;
    if (/^\d{1,2}$/.test(dueDateLower)) return true;
    const monthNamesFull = [
      'january', 'february', 'march', 'april', 'may', 'june',
      'july', 'august', 'september', 'october', 'november', 'december',
    ];
    const monthAbbr = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    const monthIndex = monthNamesFull.findIndex((m) => m.startsWith(calMonthName.toLowerCase()));
    if (monthIndex >= 0) {
      return (
        dueDateLower.includes(monthNamesFull[monthIndex]) ||
        dueDateLower.includes(monthAbbr[monthIndex])
      );
    }
    return false;
  };

  const statutoryRowMatchesCalendarDate = (row, date) => {
    if (!date) return false;
    if (!statutoryRowMatchesTransactionMonth(row, date)) return false;
    const monthKey = normalizeMonthKey(monthNames[date.getMonth()].substring(0, 3));
    const dayStr = String(date.getDate());
    const parsed = parseDueDateToMonthDay(row.dueDate || row.DueDate);
    if (parsed) {
      return normalizeMonthKey(parsed.monthKey) === monthKey && String(parsed.day) === dayStr;
    }
    const ud = String(row.dueDate || row.DueDate || '').toLowerCase();
    if (ud.includes('monthly')) {
      const dayMatch = String(row.dueDate || row.DueDate || '').match(/\b(\d{1,2})\b/);
      const monthlyDay = dayMatch ? String(parseInt(dayMatch[1], 10)) : '15';
      return monthlyDay === dayStr;
    }
    // Bare numeric due date (e.g. "7", "07", Excel-recovered day) should match that exact day.
    const recoveredDay = extractChecklistDayOfMonth(row.dueDate || row.DueDate);
    if (recoveredDay != null) {
      return String(recoveredDay) === dayStr;
    }
    // If no reliable day can be inferred, do not force a fake 15th-day due marker.
    return false;
  };

  const statutoryRowToCalendarItem = (row) => {
    const rule = String(row.formName || row.FormName || 'N/A').trim() || 'N/A';
    const schedule = String(
      row.description || row.Description || row.act || row.Act || 'No description available'
    ).trim();
    return { rule, schedule, _statutoryRow: row };
  };

  /** All Statutory forms for the viewed month (same logic every month). */
  const buildCalendarItemsFromStatutoryForMonth = (date) => {
    if (!date) return [];
    const seen = new Set();
    const out = [];
    filteredStatutoryData.forEach((row) => {
      if (!statutoryRowMatchesTransactionMonth(row, date)) return;
      const rule = String(row.formName || row.FormName || 'N/A').trim() || 'N/A';
      const schedule = String(
        row.description || row.Description || row.act || row.Act || 'No description available'
      ).trim();
      const rowId = row.id ?? row.ROWID ?? row.StatutoryId;
      const key = rowId != null && String(rowId).trim() !== ''
        ? `id:${String(rowId)}`
        : `${normLabel(rule)}|${normLabel(schedule)}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push(statutoryRowToCalendarItem(row));
    });
    return out;
  };

  const buildCalendarItemsFromStatutoryForDate = (date) => {
    if (!date) return [];
    const seen = new Set();
    const out = [];
    filteredStatutoryData.forEach((row) => {
      if (!statutoryRowMatchesCalendarDate(row, date)) return;
      const rule = String(row.formName || row.FormName || 'N/A').trim() || 'N/A';
      const schedule = String(
        row.description || row.Description || row.act || row.Act || 'No description available'
      ).trim();
      const rowId = row.id ?? row.ROWID ?? row.StatutoryId;
      const key = rowId != null && String(rowId).trim() !== ''
        ? `id:${String(rowId)}`
        : `${normLabel(rule)}|${normLabel(schedule)}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push(statutoryRowToCalendarItem(row));
    });
    return out;
  };

  const statutorySendForApprovalIsSent = (row) =>
    /^sent$/i.test(String(row?.sendForApproval ?? row?.SendForApproval ?? '').trim());

  const isStatutoryRowReturned = (row) => {
    const rawA = row?.approval ?? row?.Approval ?? '';
    const aNorm = String(rawA).trim().toLowerCase();
    if (aNorm === 'rejected' || aNorm === 'reject') return true;
    const rawS = row?.status != null && row.status !== '' ? row.status : row?.Status;
    const sNorm = String(rawS || '').trim().toLowerCase();
    return sNorm === 'rejected' || sNorm === 'reject';
  };

  /** Same display status rules as Statutory Transaction Status column. */
  const getStatutoryTransactionDisplayStatus = (row) => {
    const raw = row?.status != null && row.status !== '' ? row.status : row?.Status;
    const norm = String(raw || '').trim();
    const normalized = norm.toLowerCase();
    const rawApproval = row?.approval != null && row.approval !== '' ? row.approval : row?.Approval;
    const approvalNorm = String(rawApproval || '').trim().toLowerCase();
    if (
      approvalNorm === 'approved' ||
      approvalNorm === 'approve' ||
      normalized === 'approved' ||
      normalized === 'approve'
    ) {
      return 'Approved';
    }
    if (
      approvalNorm === 'rejected' ||
      approvalNorm === 'reject' ||
      normalized === 'rejected' ||
      normalized === 'reject'
    ) {
      return 'Returned';
    }
    if (statutorySendForApprovalIsSent(row)) return 'Pending';
    const draftVal = row?.draftFile ?? row?.DraftFile ?? null;
    const hasDraftFile =
      draftVal != null &&
      String(draftVal).trim() !== '' &&
      String(draftVal).trim() !== 'null' &&
      String(draftVal).trim() !== 'undefined';
    if (!hasDraftFile) return 'Yet to Complete';
    if (norm === '' || norm === '-' || norm === '—') return 'Pending';
    if (normalized === 'pending') return 'Pending';
    if (normalized === 'yet to complete' || normalized === 'yet to comply') {
      return statutorySendForApprovalIsSent(row) ? 'Pending' : 'Yet to Complete';
    }
    return norm || 'Pending';
  };

  const getCalendarItemDisplayStatus = (item) => {
    const row = item?._statutoryRow;
    if (!row) return null;
    if (isStatutoryRowReturned(row)) return 'Returned';
    return getStatutoryTransactionDisplayStatus(row);
  };

  const filterApprovedCalendarItems = (items) =>
    (items || []).filter((item) => getCalendarItemDisplayStatus(item) === 'Approved');

  const filterReturnedCalendarItems = (items) =>
    (items || []).filter((item) => getCalendarItemDisplayStatus(item) === 'Returned');

  const filterPendingCalendarItems = (items) =>
    (items || []).filter((item) => getCalendarItemDisplayStatus(item) === 'Pending');

  const filterYetToCompleteCalendarItems = (items) =>
    (items || []).filter((item) => getCalendarItemDisplayStatus(item) === 'Yet to Complete');

  /** Left panel: checklistbulk forms due on the selected calendar day. */
  const buildCalendarItemsFromChecklistForDate = (date) => {
    if (!date) return [];
    const seen = new Set();
    const out = [];
    calendarData.forEach((item) => {
      if (!checklistItemMatchesCalendarDate(item, date)) return;
      const rule = String(item.rule || 'N/A').trim() || 'N/A';
      const schedule = String(item.schedule || 'No description available').trim();
      const key = `${normLabel(rule)}|${normLabel(schedule)}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push({ rule, schedule, industry: item.industry });
    });
    return out;
  };

  /** Prefer Statutory Transaction rows for the day; fall back to checklist. */
  const getCalendarItemsForDate = (date) => {
    if (!date) return [];
    const itemKey = (item) => `${normLabel(item?.rule)}|${normLabel(item?.schedule)}`;
    const checklistItems = buildCalendarItemsFromChecklistForDate(date);
    const statutoryItems = buildCalendarItemsFromStatutoryForDate(date);
    if (checklistItems.length === 0) return statutoryItems;
    if (statutoryItems.length === 0) return checklistItems;

    // Keep checklist as base for "All" due forms, then enrich with statutory status rows.
    const statutoryByKey = new Map(statutoryItems.map((item) => [itemKey(item), item]));
    const mergedChecklist = checklistItems.map((item) => {
      const matched = statutoryByKey.get(itemKey(item));
      return matched?._statutoryRow ? { ...item, _statutoryRow: matched._statutoryRow } : item;
    });

    // Include statutory-only rows not present in checklist, so nothing is hidden.
    const mergedKeys = new Set(mergedChecklist.map((item) => itemKey(item)));
    const statutoryOnly = statutoryItems.filter((item) => !mergedKeys.has(itemKey(item)));
    return [...mergedChecklist, ...statutoryOnly];
  };

  const getChecklistDueDaysInMonth = (year, month) => {
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const dueDays = [];
    for (let day = 1; day <= daysInMonth; day += 1) {
      if (buildCalendarItemsFromChecklistForDate(new Date(year, month, day)).length > 0) {
        dueDays.push(day);
      }
    }
    return dueDays;
  };

  const pickNearestChecklistDueDay = (dueDays, preferredDay) => {
    const days = [...dueDays].filter((day) => day != null).sort((a, b) => a - b);
    if (!days.length) return null;
    const upcoming = days.filter((day) => day >= preferredDay);
    return upcoming.length ? upcoming[0] : days[0];
  };

  
  // Debug logging for admin users
  useEffect(() => {
    if (isAdminUser) {
      console.log('Admin user calendar data:', {
        checklistBulkData: checklistBulkData,
        finalCalendarData: calendarData,
        dataCount: calendarData.length
      });
    }
  }, [isAdminUser, calendarData, checklistBulkData]);

  const getDaysInMonth = (date) => {
    const year = date.getFullYear();
    const month = date.getMonth();
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const daysInMonth = lastDay.getDate();
    const startingDayOfWeek = firstDay.getDay();
    const today = new Date();

    const days = [];
    
    // Add empty cells for days before the first day of the month
    for (let i = 0; i < startingDayOfWeek; i++) {
      days.push({ date: null, isCurrentMonth: false, isEmpty: true });
    }

    // Add days from current month
    for (let day = 1; day <= daysInMonth; day++) {
      const currentDate = new Date(year, month, day);
      
      // Pink highlight: checklistbulk_function due dates for this day
      const hasChecklistDue = calendarData.some((item) =>
        checklistItemMatchesCalendarDate(item, currentDate)
      );
      const hasStatutoryMasterData = statutoryMasterDates.has(currentDate.toDateString());
      const hasData = hasChecklistDue || hasStatutoryMasterData;
      const isToday = currentDate.toDateString() === today.toDateString();
      
      days.push({
        date: currentDate,
        isCurrentMonth: true,
        isSelected: currentDate.toDateString() === selectedDate.toDateString(),
        isEmpty: false,
        hasData: hasData,
        hasChecklistDue,
        isToday: isToday
      });
    }

    // Fill remaining cells to complete the grid (6 weeks = 42 cells)
    const totalCells = 42;
    const remainingCells = totalCells - days.length;
    for (let i = 0; i < remainingCells; i++) {
      days.push({ date: null, isCurrentMonth: false, isEmpty: true });
    }

    return days;
  };

  const navigateMonth = (direction) => {
    setCurrentDate(prevDate => {
      const newDate = new Date(prevDate);
      newDate.setMonth(prevDate.getMonth() + direction);
      return newDate;
    });
  };

  const handleDateSelect = (date, event) => {
    setSelectedDate(date);
    setSelectedNotificationId(null);
    setActiveStatusFilter(null); // show all checklistbulk forms for the selected due date
    setFormsPage(1);
    const items = getCalendarItemsForDate(date);
    setTooltipData(items);
    setShowTooltip(true);
  };


  const handleYearSelect = (year) => {
    setCurrentDate(new Date(year, currentDate.getMonth(), 1));
    setShowYearPicker(false);
  };

  const getYearRange = () => {
    const currentYear = new Date().getFullYear();
    const years = [];
    for (let year = currentYear - 10; year <= currentYear + 10; year++) {
      years.push(year);
    }
    return years;
  };

  const days = getDaysInMonth(currentDate);

  // Convert calendar data to notification forms
  const generateNotificationForms = () => {
    try {
      const today = new Date();
      const currentYear = today.getFullYear();
      const monthNamesShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const monthMap = {
        'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'June': 5,
        'July': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
      };
      
      // Return empty array if no calendar data
      if (!calendarData || calendarData.length === 0) {
        return [];
      }
    
    // Collect all upcoming dates from calendar data
    const dateMap = new Map();
    
    calendarData.forEach((item, index) => {
      if (item.deadlines) {
        Object.keys(item.deadlines).forEach(monthKey => {
          const day = parseInt(item.deadlines[monthKey], 10);
          const monthIndex = monthMap[monthKey];
          
          if (monthIndex !== undefined && day >= 1 && day <= 31) {
            // Create date for this year
            let notificationDate = new Date(currentYear, monthIndex, day);
            
            // If date has passed this year, use next year
            if (notificationDate < today) {
              notificationDate = new Date(currentYear + 1, monthIndex, day);
            }
            
            const dateKey = `${notificationDate.getTime()}`;
            
            if (!dateMap.has(dateKey)) {
              dateMap.set(dateKey, {
                date: notificationDate,
                items: []
              });
            }
            
            dateMap.get(dateKey).items.push(item);
          }
        });
      }
    });
    
    // Convert to array and sort by date
    const notifications = Array.from(dateMap.values())
      .sort((a, b) => a.date - b.date)
      .slice(0, 16) // Limit to 16 notifications
      .map((dateGroup, index) => {
        const date = dateGroup.date;
        const items = dateGroup.items;
        // Use the first item for the card display, but store all items for tooltip
        const firstItem = items[0];
        
        // Determine button type based on form name or description
        const formName = firstItem.rule || '';
        const formNameLower = formName.toLowerCase();
        let buttonType = 'PRESENTAR'; // default
        
        if (formNameLower.includes('exam') || formNameLower.includes('test') || formNameLower.includes('examen')) {
          buttonType = 'EXAMEN';
        } else if (formNameLower.includes('submit') || formNameLower.includes('entregar') || formNameLower.includes('deliver')) {
          buttonType = 'ENTREGAR';
        } else if (formNameLower.includes('present') || formNameLower.includes('presentar')) {
          buttonType = 'PRESENTAR';
        } else {
          // Alternate based on index for variety
          buttonType = index % 3 === 0 ? 'EXAMEN' : index % 2 === 0 ? 'ENTREGAR' : 'PRESENTAR';
        }
        
        const day = date.getDate();
        const month = monthNamesShort[date.getMonth()];
        
        return {
          id: index + 1,
          date: `${day} ${month}`,
          dateObj: date,
          title: firstItem.rule || 'Form',
          description: firstItem.schedule || 'No description available',
          type: buttonType,
          priority: index % 4 === 0 ? 'high' : 'normal',
          calendarItems: items // Store all items for this date for tooltip
        };
      });
    
      return notifications;
    } catch (err) {
      console.error('Error generating notification forms:', err);
      setError(err);
      return [];
    }
  };

  const notificationForms = generateNotificationForms();

  // Close tooltip when clicking outside (but keep it open by default)
  useEffect(() => {
    const handleClickOutside = (event) => {
      // Don't close if clicking on notification cards or calendar days - allow switching
      if (showTooltip && 
          !event.target.closest('.schedule-tooltip') && 
          !event.target.closest('.calendar-day') &&
          !event.target.closest('.notification-card') &&
          !event.target.closest('.notification-action-btn') &&
          !event.target.closest('.calendar-tooltip-overlay')) {
        // Only close if clicking on empty space, but keep first notification selected
        if (selectedNotificationId !== firstNotificationId) {
          setShowTooltip(false);
          setSelectedNotificationId(null);
        }
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [showTooltip, selectedNotificationId, firstNotificationId]);

  // On load / month change: select nearest checklistbulk due day and list those forms
  useEffect(() => {
    if (loadingCalendarData) return;
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();
    const todayRef = new Date();
    const preferredDay =
      selectedDate &&
      selectedDate.getFullYear() === year &&
      selectedDate.getMonth() === month
        ? selectedDate.getDate()
        : todayRef.getFullYear() === year && todayRef.getMonth() === month
          ? todayRef.getDate()
          : 1;

    const dueDays = getChecklistDueDaysInMonth(year, month);
    const dueDay = pickNearestChecklistDueDay(dueDays, preferredDay);
    const selectedInViewedMonth =
      dueDay != null
        ? new Date(year, month, dueDay)
        : new Date(year, month, preferredDay);

    const itemsForDay = getCalendarItemsForDate(selectedInViewedMonth);
    setSelectedDate(selectedInViewedMonth);
    setSelectedNotificationId(null);
    setActiveStatusFilter(null); // show all checklistbulk forms for the due date
    setTooltipData(itemsForDay);
    setShowTooltip(true);
  }, [loadingCalendarData, checklistBulkData, inchargeDisplayScope, currentDate]);

  // Handle notification card click - update left panel with rule forms
  const handleNotificationClick = (notification, event) => {
    event.stopPropagation();
    
    // Set selected notification
    setSelectedNotificationId(notification.id);
    
    // Use the calendar items stored in the notification to update left panel
    if (notification.calendarItems && notification.calendarItems.length > 0) {
      setTooltipData(notification.calendarItems);
      setShowTooltip(true);
    } else {
      // Fallback: find calendar data for this notification's date
      const monthName = monthNames[notification.dateObj.getMonth()].substring(0, 3);
      const dayStr = notification.dateObj.getDate().toString();
      
      const relevantData = calendarData.filter(item => 
        item.deadlines && item.deadlines[monthName] === dayStr
      );
      
      if (relevantData.length > 0) {
        setTooltipData(relevantData);
        setShowTooltip(true);
      } else {
        // Create tooltip data from notification if no calendar data matches
        setTooltipData([{
          rule: notification.title,
          schedule: notification.description,
          deadlines: {
            [monthName]: dayStr
          }
        }]);
        setShowTooltip(true);
      }
    }
  };

  // Error boundary - show error message if there's an error
  if (error) {
    return (
      <div className="calendar-picker-container" style={{ padding: '20px', textAlign: 'center' }}>
        <h2>Error loading calendar</h2>
        <p>Please refresh the page or contact support if the issue persists.</p>
        <p style={{ color: 'red', fontSize: '12px' }}>{error.toString()}</p>
      </div>
    );
  }

  // Get tooltip data — selected calendar day uses Statutory Transaction rows for that due date
  const displayTooltipData = () => {
    let data = [];
    let dateInfo = null;
    const monthNamesShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

    if (selectedNotificationId && tooltipData) {
      data = tooltipData;
      const selectedNotification = notificationForms.find((n) => n.id === selectedNotificationId);
      if (selectedNotification) {
        dateInfo = { date: selectedNotification.date, dateObj: selectedNotification.dateObj };
      }
    } else if (selectedDate) {
      data = getCalendarItemsForDate(selectedDate);
      dateInfo = {
        date: `${selectedDate.getDate()} ${monthNamesShort[selectedDate.getMonth()]}`,
        dateObj: selectedDate,
      };
    } else if (tooltipData && tooltipData.length > 0) {
      data = tooltipData;
    } else {
      data = [];
    }

    return { data, dateInfo };
  };

  const { data: currentTooltipData, dateInfo } = displayTooltipData();
  const statusMonthKey = dateInfo?.dateObj
    ? monthNames[dateInfo.dateObj.getMonth()].substring(0, 3)
    : monthNames[currentDate.getMonth()].substring(0, 3);

  /** Match checklist calendar card to a Statutory row (form + description), same as Transaction page pairing. */
  const calendarRowMatchesStatutoryRow = (calendarItem, stat) => {
    const formRule = normLabel(calendarItem.rule);
    const itemForm = normLabel(stat.formName || stat.FormName);
    if (!formRule || !itemForm || itemForm !== formRule) return false;
    const statDesc = stat.description || stat.Description || '';
    return scheduleMatchesStatutoryDescription(calendarItem.schedule, statDesc);
  };

  const STATUS_PRIORITY = ['Returned', 'Approved', 'Pending', 'Yet to Complete'];

  const getStatusPriorityIndex = (status) => {
    const index = STATUS_PRIORITY.indexOf(status);
    return index === -1 ? STATUS_PRIORITY.length : index;
  };

  const getBestMatchingStatutoryRowForCalendarItem = (calendarItem, targetMonth) => {
    if (calendarItem?._statutoryRow) return calendarItem._statutoryRow;
    const matches = filteredStatutoryData.filter(
      (stat) => statutoryMonthMatches(stat, targetMonth) && calendarRowMatchesStatutoryRow(calendarItem, stat)
    );
    if (matches.length === 0) return null;
    const ranked = matches
      .map((row) => ({ row, status: getStatutoryTransactionDisplayStatus(row) }))
      .sort((a, b) => getStatusPriorityIndex(a.status) - getStatusPriorityIndex(b.status));
    return ranked[0]?.row || null;
  };

  const getDueDateFromStatutoryRow = (row, anchorDate) => {
    if (!row || !anchorDate) return null;
    const rawDueDate = row.dueDate ?? row.DueDate ?? '';
    const parsed = parseDueDateToMonthDay(rawDueDate);
    const anchorYear = anchorDate.getFullYear();
    const anchorMonth = anchorDate.getMonth();
    if (parsed?.day) {
      const parsedDay = parseInt(parsed.day, 10);
      if (!Number.isNaN(parsedDay) && parsedDay >= 1 && parsedDay <= 31) {
        const parsedMonthIndex = CALENDAR_MONTH_KEYS.findIndex(
          (key) => normalizeMonthKey(key) === normalizeMonthKey(parsed.monthKey)
        );
        const monthIndex = parsedMonthIndex >= 0 ? parsedMonthIndex : anchorMonth;
        return new Date(anchorYear, monthIndex, parsedDay);
      }
    }
    const recoveredDay = extractChecklistDayOfMonth(rawDueDate);
    if (recoveredDay != null) {
      return new Date(anchorYear, anchorMonth, recoveredDay);
    }
    return null;
  };

  const getDueDateFromChecklistItem = (item, anchorDate) => {
    if (!item?.deadlines || !anchorDate) return null;
    const selectedMonthNorm = normalizeMonthKey(CALENDAR_MONTH_KEYS[anchorDate.getMonth()]);
    const entry = Object.entries(item.deadlines).find(
      ([monthKey, day]) => normalizeMonthKey(monthKey) === selectedMonthNorm && String(day || '').trim() !== ''
    );
    if (!entry) return null;
    const dayNumber = parseInt(String(entry[1]), 10);
    if (Number.isNaN(dayNumber) || dayNumber < 1 || dayNumber > 31) return null;
    return new Date(anchorDate.getFullYear(), anchorDate.getMonth(), dayNumber);
  };

  const getCalendarItemDueDate = (calendarItem, anchorDate, targetMonth) => {
    const row = getBestMatchingStatutoryRowForCalendarItem(calendarItem, targetMonth);
    const statutoryDueDate = getDueDateFromStatutoryRow(row, anchorDate);
    if (statutoryDueDate) return statutoryDueDate;
    const checklistDueDate = getDueDateFromChecklistItem(calendarItem, anchorDate);
    if (checklistDueDate) return checklistDueDate;
    return anchorDate;
  };

  const getCalendarItemTransactionStatus = (calendarItem, targetMonth) => {
    if (calendarItem._statutoryRow) {
      return getCalendarItemDisplayStatus(calendarItem);
    }
    const matches = filteredStatutoryData.filter(
      (stat) => statutoryMonthMatches(stat, targetMonth) && calendarRowMatchesStatutoryRow(calendarItem, stat)
    );
    if (matches.length === 0) return 'Yet to Complete';
    const statuses = matches.map(getStatutoryTransactionDisplayStatus);
    for (let i = 0; i < STATUS_PRIORITY.length; i += 1) {
      if (statuses.includes(STATUS_PRIORITY[i])) return STATUS_PRIORITY[i];
    }
    return 'Yet to Complete';
  };

  const targetMonthNorm = normalizeMonthKey(statusMonthKey);
  const transactionStatusByIndex = (currentTooltipData || []).map((item) =>
    getCalendarItemTransactionStatus(item, targetMonthNorm)
  );

  const getTaskStatusByIndex = (index) => transactionStatusByIndex[index] || null;

  const countByStatus = (statuses) => ({
    yet: statuses.filter((s) => s === DISPLAY_STATUS.YET).length,
    pending: statuses.filter((s) => s === DISPLAY_STATUS.PENDING).length,
    approved: statuses.filter((s) => s === DISPLAY_STATUS.APPROVED).length,
    returned: statuses.filter((s) => s === DISPLAY_STATUS.RETURNED).length,
  });

  /** Primary: unique checklist forms due in the viewed month. */
  const getUniqueChecklistFormsDueInMonth = (year, month) => {
    const monthKey = normalizeMonthKey(CALENDAR_MONTH_KEYS[month]);
    const seen = new Set();
    const items = [];
    calendarData.forEach((item) => {
      if (!item.deadlines) return;
      const hasDeadlineInMonth = Object.entries(item.deadlines).some(
        ([mk, day]) => normalizeMonthKey(mk) === monthKey && day
      );
      if (!hasDeadlineInMonth) return;
      const key = `${normLabel(item.rule)}|${normLabel(item.schedule)}`;
      if (seen.has(key)) return;
      seen.add(key);
      items.push(item);
    });
    return items;
  };

  const viewedMonthNorm = normalizeMonthKey(CALENDAR_MONTH_KEYS[currentDate.getMonth()]);
  const monthChecklistItems = getUniqueChecklistFormsDueInMonth(
    currentDate.getFullYear(),
    currentDate.getMonth()
  );
  const primaryStatuses = monthChecklistItems.map((item) =>
    getCalendarItemTransactionStatus(item, viewedMonthNorm)
  );
  const primaryCounts = countByStatus(primaryStatuses);
  const primaryTotal =
    primaryCounts.yet + primaryCounts.pending + primaryCounts.approved + primaryCounts.returned;

  const dayFallbackCounts = countByStatus(transactionStatusByIndex);

  const monthAnchorForCounts = currentDate;
  const statutoryFallbackCounts = (() => {
    let yet = 0;
    let pending = 0;
    let approved = 0;
    let returned = 0;
    filteredStatutoryData.forEach((row) => {
      if (!statutoryRowMatchesTransactionMonth(row, monthAnchorForCounts)) return;
      const status = isStatutoryRowReturned(row)
        ? DISPLAY_STATUS.RETURNED
        : getStatutoryTransactionDisplayStatus(row);
      if (status === DISPLAY_STATUS.RETURNED) returned += 1;
      else if (status === DISPLAY_STATUS.APPROVED) approved += 1;
      else if (status === DISPLAY_STATUS.PENDING) pending += 1;
      else yet += 1;
    });
    return { yet, pending, approved, returned };
  })();

  const kpiCounts =
    primaryTotal > 0
      ? primaryCounts
      : dayFallbackCounts.yet +
          dayFallbackCounts.pending +
          dayFallbackCounts.approved +
          dayFallbackCounts.returned >
        0
        ? dayFallbackCounts
        : statutoryFallbackCounts;

  const yetToCompleteCount = kpiCounts.yet;
  const pendingCount = kpiCounts.pending;
  const approvedCount = kpiCounts.approved;
  const returnedCount = kpiCounts.returned;

  const getDayStatusSummary = (date) => {
    if (!date) {
      return { total: 0, pending: 0, approved: 0, returned: 0, yet: 0 };
    }
    const items = getCalendarItemsForDate(date);
    const monthKey = normalizeMonthKey(monthNames[date.getMonth()].substring(0, 3));
    let pending = 0;
    let approved = 0;
    let returned = 0;
    let yet = 0;
    items.forEach((item) => {
      const status = getCalendarItemTransactionStatus(item, monthKey);
      if (status === DISPLAY_STATUS.RETURNED) returned += 1;
      else if (status === DISPLAY_STATUS.APPROVED) approved += 1;
      else if (status === DISPLAY_STATUS.PENDING) pending += 1;
      else yet += 1;
    });
    return { total: items.length, pending, approved, returned, yet };
  };

  const daysWithStatus = days.map((day) => {
    if (day.isEmpty || !day.date) return day;
    const summary = getDayStatusSummary(day.date);
    const yetDueCount = summary.yet || 0;
    return { ...day, statusSummary: summary, dueCount: yetDueCount };
  });

  const visibleTooltipData = (currentTooltipData || [])
    .map((item, index) => ({ item, index, status: getTaskStatusByIndex(index) }))
    .filter(({ status }) => {
      if (!activeStatusFilter) return true;
      return status === activeStatusFilter;
    });

  const handleStatusTagClick = (status) => {
    setActiveStatusFilter((prev) => (prev === status ? null : status));
    setFormsPage(1);
  };

  const emptyStatusLabel = () => {
    if (activeStatusFilter === DISPLAY_STATUS.RETURNED) return 'returned';
    if (activeStatusFilter === DISPLAY_STATUS.APPROVED) return 'approved';
    if (activeStatusFilter === DISPLAY_STATUS.PENDING) return 'pending';
    if (activeStatusFilter === DISPLAY_STATUS.YET) return 'yet to submit';
    return '';
  };

  const formatLongDate = (date) => {
    if (!date) return '';
    const dd = String(date.getDate()).padStart(2, '0');
    return `${dd} ${monthNames[date.getMonth()]} ${date.getFullYear()}`;
  };

  const formatShortDate = (date) => {
    if (!date) return '';
    const dd = String(date.getDate()).padStart(2, '0');
    const mon = monthNames[date.getMonth()].substring(0, 3);
    return `${dd} ${mon} ${date.getFullYear()}`;
  };

  const openStatutoryForForm = (item) => {
    const selectedMonthName =
      selectedDate && selectedDate instanceof Date
        ? monthNames[selectedDate.getMonth()]
        : monthNames[currentDate.getMonth()];
    navigate('/rule-book/statutory', {
      state: {
        fromCalendar: true,
        formName: item?.rule || '',
        description: item?.schedule || '',
        monthFilter: selectedMonthName || '',
        dueDate: selectedDate ? selectedDate.toISOString() : null,
      },
    });
  };

  const resolveStatutoryRowId = (row) => {
    const raw = row?.id ?? row?.ROWID ?? row?.StatutoryId ?? '';
    const id = String(raw || '').trim();
    if (!id || id === '0' || id.toLowerCase() === 'null' || id.toLowerCase() === 'undefined') {
      return '';
    }
    return id;
  };

  const todayIsoLocal = () => {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };

  const markCalendarRowSentLocally = (targetId, submitDateIso) => {
    if (!targetId) return;
    setStatutoryData((prev) =>
      Array.isArray(prev)
        ? prev.map((row) => {
            if (String(row?.id ?? row?.ROWID ?? row?.StatutoryId ?? '').trim() !== String(targetId)) {
              return row;
            }
            return {
              ...row,
              sendForApproval: 'Sent',
              SendForApproval: 'Sent',
              status: 'Pending',
              Status: 'Pending',
              approval: '',
              Approval: '',
              remarks: '',
              Remarks: '',
              submittedDate: submitDateIso,
              SubmittedDate: submitDateIso,
            };
          })
        : prev
    );
  };

  const handleCalendarSubmit = async (item) => {
    const row = item?._statutoryRow || null;
    if (!row) {
      openStatutoryForForm(item);
      return;
    }
    if (statutorySendForApprovalIsSent(row)) {
      openStatutoryForForm(item);
      return;
    }
    const targetId = resolveStatutoryRowId(row);
    if (!targetId) {
      openStatutoryForForm(item);
      return;
    }

    const submitDateIso = todayIsoLocal();
    setSubmittingStatutoryRowId(targetId);
    markCalendarRowSentLocally(targetId, submitDateIso);

    try {
      const resp = await fetch(`/server/statutoryreg_function/statutory/${targetId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ...row,
          sendForApproval: 'Sent',
          SendForApproval: 'Sent',
          status: 'Pending',
          Status: 'Pending',
          approval: null,
          Approval: null,
          remarks: null,
          Remarks: null,
          submittedDate: submitDateIso,
          SubmittedDate: submitDateIso,
        }),
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok || data?.status !== 'success') {
        throw new Error(data?.message || 'Failed to send for approval');
      }
      if (typeof window !== 'undefined') {
        window.dispatchEvent(
          new CustomEvent('statutoryDataUpdated', { detail: { id: targetId, source: 'calendarSubmit' } })
        );
      }
      openStatutoryForForm(item);
    } catch (err) {
      console.error('Calendar submit failed:', err);
      setError(err?.message || 'Failed to send for approval');
      setTimeout(() => setError(null), 4000);
      openStatutoryForForm(item);
    } finally {
      setSubmittingStatutoryRowId('');
    }
  };

  const scopeParts = [];
  if (inchargeDisplayScope.industryLabels?.length) {
    scopeParts.push(inchargeDisplayScope.industryLabels.join(', '));
  } else if (hasSiteBasedScope) {
    scopeParts.push(allowedActCategoryList.join(', '));
  } else {
    scopeParts.push('Factories Act');
  }
  if (inchargeDisplayScope.stateLabels?.length) {
    scopeParts.push(inchargeDisplayScope.stateLabels.join(', '));
  } else {
    scopeParts.push('TamilNadu');
  }
  const scopeSubtitle = scopeParts.join(' • ');

  const todayAnchor = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const todayDueCount = getDayStatusSummary(todayAnchor).yet;
  const approvedPct =
    yetToCompleteCount + pendingCount + approvedCount + returnedCount > 0
      ? Math.round(
          (approvedCount / (yetToCompleteCount + pendingCount + approvedCount + returnedCount)) * 100
        )
      : 0;

  const pageSize = 5;
  const totalVisible = visibleTooltipData.length;
  const totalPages = Math.max(1, Math.ceil(totalVisible / pageSize));
  const safeFormsPage = Math.min(formsPage, totalPages);
  const pageStart = (safeFormsPage - 1) * pageSize;
  const displayedRows = visibleTooltipData.slice(pageStart, pageStart + pageSize);
  const showingFrom = totalVisible === 0 ? 0 : pageStart + 1;
  const showingTo = Math.min(pageStart + displayedRows.length, totalVisible);

  return (
    <div className="calendar-picker-container hcm-cal-flow">
      <div className="hcm-cal-card">
        <div className="hcm-cal-toolbar">
          <div className="hcm-cal-month-nav">
            <button type="button" className="hcm-cal-nav-btn" onClick={() => navigateMonth(-1)} aria-label="Previous Month">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                <path d="M15 18L9 12L15 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
            <button
              type="button"
              className={`hcm-cal-month-label ${showYearPicker ? 'open' : ''}`}
              onClick={() => setShowYearPicker(!showYearPicker)}
            >
              {monthNames[currentDate.getMonth()]} {currentDate.getFullYear()}
            </button>
            <button type="button" className="hcm-cal-nav-btn" onClick={() => navigateMonth(1)} aria-label="Next Month">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                <path d="M9 18L15 12L9 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
          </div>

        </div>

        <div className="hcm-cal-grid-wrap">
          <div className="hcm-cal-day-headers">
            {dayNames.map((day) => (
              <div key={day} className="hcm-cal-day-header">
                {day}
              </div>
            ))}
          </div>
          <div className="hcm-cal-grid">
            {daysWithStatus.map((day, index) => {
              if (day.isEmpty || !day.date) {
                return <div key={index} className="hcm-cal-day empty" />;
              }
              const summary = day.statusSummary || {};
              const dueCount = day.dueCount || 0;
              return (
                <button
                  key={index}
                  type="button"
                  className={`hcm-cal-day${day.isToday ? ' today' : ''}${day.isSelected ? ' selected' : ''}${dueCount > 0 ? ' has-due' : ''}`}
                  onClick={(e) => {
                    handleDateSelect(day.date, e);
                  }}
                  title={dueCount > 0 ? `${dueCount} yet-to-submit form(s)` : undefined}
                >
                  <span className="hcm-cal-day-top">
                    <span className="hcm-cal-day-number">{day.date.getDate()}</span>
                    {dueCount > 0 ? <span className="hcm-cal-day-badge">{dueCount}</span> : null}
                  </span>
                  <span className="hcm-cal-day-dots" aria-hidden>
                    {summary.pending > 0 ? <span className="dot pending" /> : null}
                    {summary.approved > 0 ? <span className="dot approved" /> : null}
                    {summary.returned > 0 ? <span className="dot returned" /> : null}
                    {summary.yet > 0 && !summary.pending && !summary.approved && !summary.returned ? (
                      <span className="dot yet" />
                    ) : null}
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        <div className="hcm-cal-summary-row">
          <button
            type="button"
            className={`hcm-cal-summary-card yet ${activeStatusFilter === DISPLAY_STATUS.YET ? 'active' : ''}`}
            onClick={() => handleStatusTagClick(DISPLAY_STATUS.YET)}
          >
            <div className="hcm-cal-summary-icon" aria-hidden>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="#fff" strokeWidth="2" />
                <path d="M14 2v6h6" stroke="#fff" strokeWidth="2" />
                <path d="M9 15h6M9 11h2" stroke="#fff" strokeWidth="2" strokeLinecap="round" />
              </svg>
            </div>
            <div>
              <div className="hcm-cal-summary-label">Yet to Submit</div>
              <div className="hcm-cal-summary-value">{yetToCompleteCount}</div>
              <div className="hcm-cal-summary-sub">↑ {todayDueCount} Due Today</div>
            </div>
          </button>
          <button
            type="button"
            className={`hcm-cal-summary-card pending ${activeStatusFilter === DISPLAY_STATUS.PENDING ? 'active' : ''}`}
            onClick={() => handleStatusTagClick(DISPLAY_STATUS.PENDING)}
          >
            <div className="hcm-cal-summary-icon" aria-hidden>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
                <circle cx="12" cy="12" r="9" stroke="#fff" strokeWidth="2" />
                <path d="M12 7v5l3 2" stroke="#fff" strokeWidth="2" strokeLinecap="round" />
              </svg>
            </div>
            <div>
              <div className="hcm-cal-summary-label">Pending</div>
              <div className="hcm-cal-summary-value">{pendingCount}</div>
              <div className="hcm-cal-summary-sub">Awaiting Approval</div>
            </div>
          </button>
          <button
            type="button"
            className={`hcm-cal-summary-card approved ${activeStatusFilter === DISPLAY_STATUS.APPROVED ? 'active' : ''}`}
            onClick={() => handleStatusTagClick(DISPLAY_STATUS.APPROVED)}
          >
            <div className="hcm-cal-summary-icon" aria-hidden>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
                <circle cx="12" cy="12" r="9" stroke="#fff" strokeWidth="2" />
                <path d="M8.5 12.5l2.5 2.5 4.5-5" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </div>
            <div>
              <div className="hcm-cal-summary-label">Approved</div>
              <div className="hcm-cal-summary-value">{approvedCount}</div>
              <div className="hcm-cal-summary-sub">{approvedPct}% Completed</div>
            </div>
          </button>
          <button
            type="button"
            className={`hcm-cal-summary-card returned ${activeStatusFilter === DISPLAY_STATUS.RETURNED ? 'active' : ''}`}
            onClick={() => handleStatusTagClick(DISPLAY_STATUS.RETURNED)}
          >
            <div className="hcm-cal-summary-icon" aria-hidden>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
                <circle cx="12" cy="12" r="9" stroke="#fff" strokeWidth="2" />
                <circle cx="12" cy="12" r="3" stroke="#fff" strokeWidth="2" />
                <path d="M16.5 7.5l2-2" stroke="#fff" strokeWidth="2" strokeLinecap="round" />
              </svg>
            </div>
            <div>
              <div className="hcm-cal-summary-label">Returned</div>
              <div className="hcm-cal-summary-value">{returnedCount}</div>
              <div className="hcm-cal-summary-sub">Need Correction</div>
            </div>
          </button>
        </div>

        <div className="hcm-cal-forms-section">
          <div className="hcm-cal-forms-header">
            <div className="hcm-cal-forms-header-main">
              <div className="hcm-cal-forms-header-row">
                <h2 className="hcm-cal-forms-title">
                  Forms Due on {formatLongDate(selectedDate)}
                  <span className="hcm-cal-forms-count">{currentTooltipData?.length || 0}</span>
                </h2>
                <div className="hcm-cal-forms-tabs">
                  <button
                    type="button"
                    className={`hcm-cal-tab ${!activeStatusFilter ? 'active' : ''}`}
                    onClick={() => {
                      setActiveStatusFilter(null);
                      setFormsPage(1);
                    }}
                  >
                    All
                  </button>
                  <button
                    type="button"
                    className={`hcm-cal-tab ${activeStatusFilter === DISPLAY_STATUS.YET ? 'active' : ''}`}
                    onClick={() => handleStatusTagClick(DISPLAY_STATUS.YET)}
                  >
                    Yet to Submit
                  </button>
                  <button
                    type="button"
                    className={`hcm-cal-tab ${activeStatusFilter === DISPLAY_STATUS.PENDING ? 'active' : ''}`}
                    onClick={() => handleStatusTagClick(DISPLAY_STATUS.PENDING)}
                  >
                    Pending
                  </button>
                  <button
                    type="button"
                    className={`hcm-cal-tab ${activeStatusFilter === DISPLAY_STATUS.APPROVED ? 'active' : ''}`}
                    onClick={() => handleStatusTagClick(DISPLAY_STATUS.APPROVED)}
                  >
                    Approved
                  </button>
                  <button
                    type="button"
                    className={`hcm-cal-tab ${activeStatusFilter === DISPLAY_STATUS.RETURNED ? 'active' : ''}`}
                    onClick={() => handleStatusTagClick(DISPLAY_STATUS.RETURNED)}
                  >
                    Returned
                  </button>
                </div>
              </div>
              <p className="hcm-cal-forms-scope">{scopeSubtitle}</p>
            </div>
          </div>

          {displayedRows.length > 0 ? (
            <div className="hcm-cal-forms-list">
              {displayedRows.map(({ item, index: origIndex, status }, index) => {
                const iconColor = FORM_ICON_COLORS[index % FORM_ICON_COLORS.length];
                const displayStatus = STATUS_LABEL[status] || status || 'Yet to Submit';
                const statusClass =
                  status === DISPLAY_STATUS.APPROVED
                    ? 'approved'
                    : status === DISPLAY_STATUS.PENDING
                      ? 'pending'
                      : status === DISPLAY_STATUS.RETURNED
                        ? 'returned'
                        : 'yet';
                const dueDateForRow = getCalendarItemDueDate(item, selectedDate, targetMonthNorm);
                const statutoryRowId = resolveStatutoryRowId(item?._statutoryRow);
                const isSubmittingThisRow =
                  !!submittingStatutoryRowId && statutoryRowId && submittingStatutoryRowId === statutoryRowId;
                const canSubmit =
                  status === DISPLAY_STATUS.YET ||
                  status === DISPLAY_STATUS.PENDING ||
                  status === DISPLAY_STATUS.RETURNED ||
                  !status;
                return (
                  <div key={`${origIndex}-${normLabel(item.rule)}`} className="hcm-cal-form-row">
                    <div className="hcm-cal-form-icon" style={{ background: iconColor }}>
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="#fff" strokeWidth="2" />
                        <path d="M14 2v6h6" stroke="#fff" strokeWidth="2" />
                      </svg>
                    </div>
                    <div className="hcm-cal-form-info">
                      <div className="hcm-cal-form-name">{item.rule}</div>
                      <div className="hcm-cal-form-desc">{item.schedule}</div>
                    </div>
                    <div className={`hcm-cal-form-status ${statusClass}`}>{displayStatus}</div>
                    <div className="hcm-cal-form-due">Due Date {formatShortDate(dueDateForRow)}</div>
                    <div className="hcm-cal-form-actions">
                      <button type="button" className="hcm-cal-btn-view" onClick={() => openStatutoryForForm(item)}>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden>
                          <path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z" stroke="currentColor" strokeWidth="2" />
                          <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2" />
                        </svg>
                        View
                      </button>
                      {canSubmit ? (
                        <button
                          type="button"
                          className="hcm-cal-btn-submit"
                          style={{ background: iconColor }}
                          onClick={() => handleCalendarSubmit(item)}
                          disabled={isSubmittingThisRow}
                        >
                          {isSubmittingThisRow ? 'Submitting...' : 'Submit'}
                        </button>
                      ) : (
                        <button
                          type="button"
                          className="hcm-cal-btn-submit muted"
                          onClick={() => openStatutoryForForm(item)}
                        >
                          Open
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="hcm-cal-forms-empty">
              <p>No{emptyStatusLabel() ? ` ${emptyStatusLabel()}` : ''} forms for this date</p>
              <p className="hcm-cal-forms-empty-sub">
                {activeStatusFilter
                    ? 'Try another status tab or date'
                    : 'Click a calendar date to view due forms'}
              </p>
            </div>
          )}

          {totalVisible > 0 ? (
            <div className="hcm-cal-forms-footer hcm-cal-forms-footer--company-model">
              <span>
                Showing {showingFrom} to {showingTo} of {totalVisible} results
              </span>
              <nav className="company-details-pagination" aria-label="Forms pagination">
                <button
                  type="button"
                  className="company-details-pagination-nav"
                  disabled={safeFormsPage <= 1}
                  onClick={() => setFormsPage(1)}
                  title="First page"
                  aria-label="First page"
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                    <polyline points="11 17 6 12 11 7" />
                    <polyline points="18 17 13 12 18 7" />
                  </svg>
                </button>
                <button
                  type="button"
                  className="company-details-pagination-nav"
                  disabled={safeFormsPage <= 1}
                  onClick={() => setFormsPage((p) => Math.max(1, Math.min(p, totalPages) - 1))}
                  title="Previous page"
                  aria-label="Previous page"
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                    <polyline points="15 18 9 12 15 6" />
                  </svg>
                </button>
                <div className="company-details-pagination-pages">
                  {buildPaginationItems(safeFormsPage, totalPages).map((item, i) =>
                    item === 'ellipsis' ? (
                      <span key={`e-${i}`} className="company-details-pagination-ellipsis" aria-hidden>
                        ...
                      </span>
                    ) : (
                      <button
                        key={item}
                        type="button"
                        className={`company-details-pagination-page${item === safeFormsPage ? ' company-details-pagination-page--active' : ''}`}
                        onClick={() => setFormsPage(item)}
                        aria-label={`Page ${item}`}
                        aria-current={item === safeFormsPage ? 'page' : undefined}
                      >
                        {item}
                      </button>
                    )
                  )}
                </div>
                <button
                  type="button"
                  className="company-details-pagination-nav"
                  disabled={safeFormsPage >= totalPages}
                  onClick={() => setFormsPage((p) => Math.min(totalPages, Math.min(p, totalPages) + 1))}
                  title="Next page"
                  aria-label="Next page"
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                    <polyline points="9 18 15 12 9 6" />
                  </svg>
                </button>
                <button
                  type="button"
                  className="company-details-pagination-nav"
                  disabled={safeFormsPage >= totalPages}
                  onClick={() => setFormsPage(totalPages)}
                  title="Last page"
                  aria-label="Last page"
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
                    <polyline points="13 17 18 12 13 7" />
                    <polyline points="6 17 11 12 6 7" />
                  </svg>
                </button>
              </nav>
            </div>
          ) : null}
        </div>
      </div>

      {showYearPicker && (
        <div className="year-picker">
          <div className="year-picker-content">
            {getYearRange().map((year) => (
              <button
                key={year}
                type="button"
                className={`year-option ${currentDate.getFullYear() === year ? 'selected' : ''}`}
                onClick={() => handleYearSelect(year)}
              >
                {year}
              </button>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

export default CalendarPicker;