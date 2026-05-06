import React, { useState, useEffect, useMemo } from 'react';
import './CalendarPicker.css';
import { fetchAllowedActCategoriesFromSites, getActCategoryFromActSector } from '../utils/siteInchargeScope';

const CalendarPicker = ({ userEmail, userRole }) => {
  // Initialize with current date and year
  const today = new Date();
  const [currentDate, setCurrentDate] = useState(new Date(today.getFullYear(), today.getMonth(), 1)); // Current month and year
  const [selectedDate, setSelectedDate] = useState(new Date(today.getFullYear(), today.getMonth(), today.getDate())); // Current date
  const [showYearPicker, setShowYearPicker] = useState(false);
  const [showTooltip, setShowTooltip] = useState(false);
  const [tooltipData, setTooltipData] = useState(null);
  const [selectedNotificationId, setSelectedNotificationId] = useState(null);
  const [activeStatusFilter, setActiveStatusFilter] = useState(null);
  const [error, setError] = useState(null);
  const [calendarSearch, setCalendarSearch] = useState('');
  
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
  const [allowedActCategoryList, setAllowedActCategoryList] = useState(null);

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  const dayNames = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

  const hasSiteBasedScope = Array.isArray(allowedActCategoryList) && allowedActCategoryList.length > 0;
  
  // Check if user is admin (App Administrator role)
  const isAdminUser = userRole === 'App Administrator';

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const cats = await fetchAllowedActCategoriesFromSites(userEmail);
      if (!cancelled) setAllowedActCategoryList(cats);
    })();
    return () => {
      cancelled = true;
    };
  }, [userEmail]);

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

  // Parse a single due date string to { monthKey, day } for calendar (e.g. "15-Feb" -> { monthKey: "Feb", day: "15" })
  const parseDueDateToMonthDay = (dueDate) => {
    const s = String(dueDate || '').trim();
    if (!s) return null;
    const monthAbbr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    // Match "15-Feb", "15-Feb-2026", "15 Feb", "Feb 15", "15/02", "15-02", "31-Dec"
    const ddm = s.match(/^(\d{1,2})[-/](\d{1,2})(?:[-/]\d{2,4})?$/); // 15-02 or 15/02
    if (ddm) {
      const day = parseInt(ddm[1], 10);
      const monthNum = parseInt(ddm[2], 10);
      if (day >= 1 && day <= 31 && monthNum >= 1 && monthNum <= 12)
        return { monthKey: monthAbbr[monthNum - 1], day: String(day) };
    }
    const dm = s.match(/^(\d{1,2})[-/\s]([A-Za-z]{3,})/); // 15-Feb or 15 Feb
    if (dm) {
      const day = parseInt(dm[1], 10);
      const mon = dm[2].toLowerCase();
      const mi = monthAbbr.findIndex(m => m.toLowerCase().startsWith(mon));
      if (day >= 1 && day <= 31 && mi >= 0)
        return { monthKey: monthAbbr[mi], day: String(day) };
    }
    const md = s.match(/([A-Za-z]{3,})[-/\s](\d{1,2})/); // Feb 15 or Feb-15
    if (md) {
      const mon = md[1].toLowerCase();
      const mi = monthAbbr.findIndex(m => m.toLowerCase().startsWith(mon));
      const day = parseInt(md[2], 10);
      if (day >= 1 && day <= 31 && mi >= 0)
        return { monthKey: monthAbbr[mi], day: String(day) };
    }
    return null;
  };

  // Convert checklist bulk data to calendar format (monthly basis + specific dates)
  const convertChecklistBulkToCalendarFormat = (bulkData) => {
    const calendarData = [];
    
    try {
      if (!Array.isArray(bulkData)) return calendarData;
      
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      
      bulkData.forEach(item => {
        if (!item || typeof item !== 'object') return;
        
        const dueDate = item.dueDate || '';
        const formName = item.formName || 'N/A';
        const description = item.description || 'No description available';
        const sector = item.sector || 'General';
        const dueDateStr = String(dueDate).toLowerCase().trim();
        const isMonthlyBasis = dueDateStr.includes('monthly') || dueDateStr === 'monthly basis';
        
        if (isMonthlyBasis) {
          let dayNumber = 16;
          const dueDateDayMatch = String(dueDate).match(/\b(\d{1,2})\b/);
          if (dueDateDayMatch) {
            const extractedDay = parseInt(dueDateDayMatch[1], 10);
            if (extractedDay >= 1 && extractedDay <= 31) dayNumber = extractedDay;
          }
          monthNames.forEach(month => {
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

  const displayChecklistBulkData = (() => {
    if (!checklistBulkData || checklistBulkData.length === 0) return [];
    if (hasSiteBasedScope) {
      const allow = new Set(allowedActCategoryList);
      return checklistBulkData.filter((item) =>
        allow.has(getActCategoryFromActSector(item.act || item.Act || '', item.sector || item.Sector || ''))
      );
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

  // Filter calendar data based on selected industry for specific user
  const getFilteredCalendarData = () => {
    const baseData = getBaseCalendarData();
    
    // For admin users, show all data without filtering
    if (isAdminUser) {
      return baseData;
    }
    
    if (hasSiteBasedScope) {
      return baseData;
    }
    if (!selectedIndustry) return baseData;
    return baseData.filter((item) => item.industry && item.industry.toLowerCase() === selectedIndustry.toLowerCase());
  };

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

  const calendarSearchLower = calendarSearch.trim().toLowerCase();
  const calendarDataForSearch = useMemo(() => {
    if (!calendarSearchLower) return calendarData;
    return calendarData.filter((item) => {
      const rule = String(item.rule || '').toLowerCase();
      const schedule = String(item.schedule || '').toLowerCase();
      return rule.includes(calendarSearchLower) || schedule.includes(calendarSearchLower);
    });
  }, [calendarData, calendarSearchLower]);
  
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
      const monthName = monthNames[month].substring(0, 3);
      const dayStr = day.toString();
      
      // Check if this date has any scheduled data (checklist bulk)
      const hasScheduleData = calendarDataForSearch.some(item => 
        item.deadlines && item.deadlines[monthName] === dayStr
      );
      // Check if this date has statutory master data
      const hasStatutoryData = statutoryMasterDates.has(currentDate.toDateString());
      const hasData = hasScheduleData || hasStatutoryData;
      const isToday = currentDate.toDateString() === today.toDateString();
      
      days.push({
        date: currentDate,
        isCurrentMonth: true,
        isSelected: currentDate.toDateString() === selectedDate.toDateString(),
        isEmpty: false,
        hasData: hasData,
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
    setSelectedNotificationId(null); // Clear notification selection when calendar date is clicked
    
    // Find data for this date from checklistbulk data only
    const monthName = monthNames[date.getMonth()].substring(0, 3);
    const dayStr = date.getDate().toString();
    
    console.log('Date selected:', date.getDate(), 'Month:', monthName, 'Day:', dayStr);
    console.log('Calendar data from checklistbulk:', calendarData);
    
    // Filter calendar data to find entries matching this date
    const relevantData = calendarDataForSearch.filter(item => 
      item.deadlines && item.deadlines[monthName] === dayStr
    );
    
    console.log('Relevant checklistbulk data found:', relevantData);
    
    // Update tooltip data to show in left panel
    if (relevantData.length > 0) {
      setTooltipData(relevantData);
      setShowTooltip(true);
      console.log('Tooltip data updated for left panel:', relevantData);
    } else {
      setTooltipData([]);
      setShowTooltip(false);
      console.log('No relevant data found for this date');
    }
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
      if (!calendarDataForSearch || calendarDataForSearch.length === 0) {
        return [];
      }
    
    // Collect all upcoming dates from calendar data
    const dateMap = new Map();
    
    calendarDataForSearch.forEach((item, index) => {
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

  // Auto-select first notification and show tooltip on page load
  useEffect(() => {
    // Wait for data to load and ensure we haven't already selected a notification
    if (!loadingCalendarData && notificationForms.length > 0 && !selectedNotificationId && !showTooltip) {
      const firstNotification = notificationForms[0];
      setFirstNotificationId(firstNotification.id);
      setSelectedNotificationId(firstNotification.id);
      
      // Set tooltip data for first notification
      if (firstNotification.calendarItems && firstNotification.calendarItems.length > 0) {
        setTooltipData(firstNotification.calendarItems);
        setShowTooltip(true);
      } else {
        // Fallback: find calendar data for this notification's date
        const monthName = monthNames[firstNotification.dateObj.getMonth()].substring(0, 3);
        const dayStr = firstNotification.dateObj.getDate().toString();
        
        const relevantData = calendarDataForSearch.filter(item => 
          item.deadlines && item.deadlines[monthName] === dayStr
        );
        
        if (relevantData.length > 0) {
          setTooltipData(relevantData);
          setShowTooltip(true);
        } else {
          setTooltipData([{
            rule: firstNotification.title,
            schedule: firstNotification.description,
            deadlines: {
              [monthName]: dayStr
            }
          }]);
          setShowTooltip(true);
        }
      }
    }
  }, [loadingCalendarData, notificationForms.length, calendarData.length, calendarDataForSearch.length, calendarSearchLower]); // Run when data is loaded

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
      
      const relevantData = calendarDataForSearch.filter(item => 
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

  // Get tooltip data to display - use selected date/notification, else show all Schedule Of Submission from checklistbulk
  const displayTooltipData = () => {
    let data = [];
    let dateInfo = null;
    
    if (selectedNotificationId && tooltipData) {
      data = tooltipData;
      const selectedNotification = notificationForms.find(n => n.id === selectedNotificationId);
      if (selectedNotification) {
        dateInfo = { date: selectedNotification.date, dateObj: selectedNotification.dateObj };
      }
    } else if (notificationForms.length > 0 && notificationForms[0].calendarItems && tooltipData && tooltipData.length > 0) {
      data = tooltipData;
      dateInfo = { date: notificationForms[0].date, dateObj: notificationForms[0].dateObj };
    } else if (tooltipData && tooltipData.length > 0) {
      data = tooltipData;
      if (selectedDate) {
        const monthNamesShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        dateInfo = { date: `${selectedDate.getDate()} ${monthNamesShort[selectedDate.getMonth()]}`, dateObj: selectedDate };
      }
    } else {
      const q = calendarSearch.trim().toLowerCase();
      if (!q) {
        data = scheduleItemsFromChecklistBulk;
      } else {
        data = scheduleItemsFromChecklistBulk.filter((item) => {
          const rule = String(item.rule || '').toLowerCase();
          const schedule = String(item.schedule || '').toLowerCase();
          return rule.includes(q) || schedule.includes(q);
        });
      }
    }
    
    return { data, dateInfo };
  };

  const { data: currentTooltipData, dateInfo } = displayTooltipData();
  const statusMonthKey = dateInfo?.dateObj
    ? monthNames[dateInfo.dateObj.getMonth()].substring(0, 3)
    : monthNames[currentDate.getMonth()].substring(0, 3);

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
    if (hasSiteBasedScope) {
      const allow = new Set(allowedActCategoryList);
      return statutoryData.filter((item) =>
        allow.has(getActCategoryFromActSector(item.act || item.Act || '', item.sector || item.Sector || ''))
      );
    }
    if (!isAdminUser && selectedIndustry) {
      return statutoryData.filter((item) =>
        String(item.sector || item.Sector || '').trim().toLowerCase() === selectedIndustry.trim().toLowerCase()
      );
    }
    return statutoryData;
  })();

  const normLabel = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();

  /** Statutory row has an uploaded draft file */
  const statutoryRowHasDraft = (row) => {
    const draftFile = row.draftFile || row.DraftFile || null;
    const draftFileName = row.draftFileName || row.DraftFileName || null;
    const hasId = draftFile != null && String(draftFile).trim() !== '' && String(draftFile).trim() !== 'null';
    const hasName = draftFileName != null && String(draftFileName).trim() !== '' && String(draftFileName).trim() !== 'null';
    return hasId || hasName;
  };

  /** Align checklist/bulk schedule text with Statutory description (same row, not same form name only). */
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

  /** Unique key so one Statutory draft cannot complete multiple duplicate bulk cards. */
  const statutoryDraftConsumeKey = (stat) => {
    const id = stat.id ?? stat.StatutoryId ?? stat.statutoryId;
    if (id != null && String(id).trim() !== '') return `id:${String(id)}`;
    const draftBit = String(stat.draftFile || stat.DraftFile || stat.draftFileName || stat.DraftFileName || '').slice(0, 64);
    return `h:${normLabel(stat.formName || stat.FormName)}|${normLabel(stat.description || stat.Description)}|${draftBit}`;
  };

  const calendarRowMatchesStatutory = (calendarItem, stat) => {
    const formRule = normLabel(calendarItem.rule);
    const itemForm = normLabel(stat.formName || stat.FormName);
    if (!formRule || !itemForm || itemForm !== formRule) return false;
    if (!statutoryRowHasDraft(stat)) return false;
    const statDesc = stat.description || stat.Description || '';
    return scheduleMatchesStatutoryDescription(calendarItem.schedule, statDesc);
  };

  /**
   * One Statutory draft row completes at most one calendar card (fixes duplicate bulk rows
   * e.g. two "Form I" cards sharing one draft submission).
   */
  const draftCompletionByIndex = (() => {
    const items = currentTooltipData || [];
    const targetMonth = normalizeMonthKey(statusMonthKey);
    const candidates = filteredStatutoryData.filter(
      (stat) => statutoryRowHasDraft(stat) && statutoryMonthMatches(stat, targetMonth)
    );
    const usedStatutoryKeys = new Set();
    return items.map((item) => {
      for (let i = 0; i < candidates.length; i += 1) {
        const stat = candidates[i];
        const key = statutoryDraftConsumeKey(stat);
        if (usedStatutoryKeys.has(key)) continue;
        if (!calendarRowMatchesStatutory(item, stat)) continue;
        usedStatutoryKeys.add(key);
        return true;
      }
      return false;
    });
  })();

  const getTaskStatusByIndex = (index) => (draftCompletionByIndex[index] ? 'Completed' : 'Pending');

  const pending = (currentTooltipData || []).length - draftCompletionByIndex.filter(Boolean).length;
  const completed = draftCompletionByIndex.filter(Boolean).length;

  const visibleTooltipData = (currentTooltipData || [])
    .map((item, index) => ({ item, index }))
    .filter(({ index }) => {
      if (!activeStatusFilter) return true;
      return getTaskStatusByIndex(index) === activeStatusFilter;
    });

  const handleStatusTagClick = (status) => {
    setActiveStatusFilter((prev) => (prev === status ? null : status));
  };

  return (
    <div className="calendar-picker-container">
      <div className="calendar-combined-card">
        <div className="calendar-combined-toolbar">
          <div className="calendar-toolbar-sync-row">
            <div className="calendar-toolbar-left">
              <input
                type="text"
                value={calendarSearch}
                onChange={(e) => setCalendarSearch(e.target.value)}
                className="calendar-statutory-search-input"
                placeholder="Search by Form Name or Description..."
                title="Search"
                aria-label="Search by Form Name or Description"
              />
            </div>
            <div className="calendar-toolbar-spacer" aria-hidden="true" />
          </div>
        </div>
        <div className="calendar-combined-body">
      {/* Left Side - Rule Form Boxes */}
      <div className="notifications-panel">
        <div className="notifications-header">
          <div className="notifications-header-left">
            <div className="notifications-header-icon">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M14 2H6C5.46957 2 4.96086 2.21071 4.58579 2.58579C4.21071 2.96086 4 3.46957 4 4V20C4 20.5304 4.21071 21.0391 4.58579 21.4142C4.96086 21.7893 5.46957 22 6 22H18C18.5304 22 19.0391 21.7893 19.4142 21.4142C19.7893 21.0391 20 20.5304 20 20V8L14 2Z" stroke="#1976d2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M14 2V8H20" stroke="#1976d2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M16 13H8" stroke="#1976d2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M16 17H8" stroke="#1976d2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M10 9H9H8" stroke="#1976d2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            </div>
            <h2 className="notifications-title">Schedule Of Submission/ Maintenance</h2>
            {hasSiteBasedScope && (
              <span className="notifications-subtitle" style={{ fontSize: '12px', color: '#6b7280', marginLeft: '8px', fontWeight: 'normal' }}>
                (filtered by Site Management industry scope: {allowedActCategoryList.join(', ')})
              </span>
            )}
          </div>
          <div className="notification-status-tags">
            <button
              type="button"
              className={`status-tag pending${activeStatusFilter === 'Pending' ? ' is-active' : ''}${activeStatusFilter === 'Completed' ? ' is-inactive' : ''}`}
              onClick={() => handleStatusTagClick('Pending')}
              aria-pressed={activeStatusFilter === 'Pending'}
              style={{ cursor: 'pointer' }}
              title="Show pending forms"
            >
              PENDING ({pending})
            </button>
            <button
              type="button"
              className={`status-tag completed${activeStatusFilter === 'Completed' ? ' is-active' : ''}${activeStatusFilter === 'Pending' ? ' is-inactive' : ''}`}
              onClick={() => handleStatusTagClick('Completed')}
              aria-pressed={activeStatusFilter === 'Completed'}
              style={{ cursor: 'pointer' }}
              title="Show completed forms"
            >
              COMPLETED ({completed})
            </button>
          </div>
        </div>
        <div className="notifications-scroll-container rule-forms-container">
          {visibleTooltipData && visibleTooltipData.length > 0 ? (
            <div className="rule-forms-list">
              {visibleTooltipData.map(({ item, index: origIndex }, index) => {
                const colors = [
                  { bg: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)' },
                  { bg: 'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)' },
                  { bg: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)' },
                  { bg: 'linear-gradient(135deg, #43e97b 0%, #38f9d7 100%)' },
                  { bg: 'linear-gradient(135deg, #fa709a 0%, #fee140 100%)' },
                  { bg: 'linear-gradient(135deg, #30cfd0 0%, #330867 100%)' },
                  { bg: 'linear-gradient(135deg, #a8edea 0%, #fed6e3 100%)' },
                  { bg: 'linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%)' }
                ];
                const colorScheme = colors[index % colors.length];
                const taskStatus = getTaskStatusByIndex(origIndex);
                return (
                  <div key={`${origIndex}-${normLabel(item.rule)}`} className="rule-form-card">
                    <div
                      className="rule-form-card-header"
                      style={{ background: colorScheme.bg }}
                    >
                      <div className="doc-icon">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <path d="M14 2H6C5.46957 2 4.96086 2.21071 4.58579 2.58579C4.21071 2.96086 4 3.46957 4 4V20C4 20.5304 4.21071 21.0391 4.58579 21.4142C4.96086 21.7893 5.46957 22 6 22H18C18.5304 22 19.0391 21.7893 19.4142 21.4142C19.7893 21.0391 20 20.5304 20 20V8L14 2Z" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <path d="M14 2V8H20" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <path d="M16 13H8" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <path d="M16 17H8" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                          <path d="M10 9H9H8" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                      </div>
                      <h3 className="rule-form-card-title">{item.rule}</h3>
                    </div>
                    <div className="rule-form-card-body">
                      {item.schedule}
                      {taskStatus === 'Completed' && (
                        <div style={{ marginTop: '6px', fontSize: '11px', color: '#16a34a', fontWeight: 600 }}>
                          Completed (Draft Uploaded)
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="no-notifications">
              <p>No {activeStatusFilter ? activeStatusFilter.toLowerCase() : ''} rule forms available</p>
              <p className="no-notifications-subtitle">
                {activeStatusFilter
                  ? 'Try another status tab or date to view forms'
                  : 'Click on a calendar date or notification to view forms'}
              </p>
            </div>
          )}
        </div>
      </div>

      {/* Right Side - Calendar View */}
      <div className="calendar-picker-card">
        {/* Calendar Heading with Icon */}
        <div className="calendar-main-heading">
          <div className="calendar-heading-icon">
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2" stroke="#1976d2" strokeWidth="2"/>
              <line x1="16" y1="2" x2="16" y2="6" stroke="#1976d2" strokeWidth="2" strokeLinecap="round"/>
              <line x1="8" y1="2" x2="8" y2="6" stroke="#1976d2" strokeWidth="2" strokeLinecap="round"/>
              <line x1="3" y1="10" x2="21" y2="10" stroke="#1976d2" strokeWidth="2"/>
              <circle cx="8" cy="14" r="1" fill="#1976d2"/>
              <circle cx="12" cy="14" r="1" fill="#1976d2"/>
              <circle cx="16" cy="14" r="1" fill="#1976d2"/>
              <circle cx="8" cy="18" r="1" fill="#1976d2"/>
              <circle cx="12" cy="18" r="1" fill="#1976d2"/>
            </svg>
          </div>
          <h1 className="calendar-heading-text">Calendar</h1>
        </div>

        {/* Optional industry selector for non-admin users without Site scope */}
        {!hasSiteBasedScope && !isAdminUser && (
          <div className="industry-selector">
            <div>
              <label>Select Industry:</label>
              <select
                value={selectedIndustry}
                onChange={(e) => setSelectedIndustry(e.target.value)}
                disabled={loadingSiteData || loadingCalendarData}
              >
                {(loadingSiteData || loadingCalendarData) ? (
                  <option value="">Loading industries...</option>
                ) : availableIndustries.length === 0 ? (
                  <option value="">No industries available</option>
                ) : (
                  availableIndustries.map((industry, index) => (
                    <option key={index} value={industry}>
                      {industry}
                    </option>
                  ))
                )}
              </select>
              {selectedIndustry && (
                <div className="industry-status">
                  Showing calendar for: <strong>{selectedIndustry}</strong>
                  {checklistBulkData.length > 0 && (
                    <div className="industry-count">
                      Based on {checklistBulkData.length} checklist bulk entries
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        )}
        {/* Month navigation: prev | month + year | next (reference layout) */}
        <div className="calendar-picker-header">
          <button
            type="button"
            className="nav-button prev-button"
            onClick={() => navigateMonth(-1)}
            aria-label="Previous Month"
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
              <path d="M15 18L9 12L15 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
          </button>

          <div className="calendar-picker-header-center">
            <h2 className="month-title">
              <span className="month-title-name">{monthNames[currentDate.getMonth()]}</span>
              <button
                type="button"
                className={`year-selector ${showYearPicker ? 'open' : ''}`}
                onClick={() => setShowYearPicker(!showYearPicker)}
                aria-expanded={showYearPicker}
                aria-haspopup="dialog"
              >
                {currentDate.getFullYear()}
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" aria-hidden>
                  <path d="M6 9L12 15L18 9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </button>
            </h2>
          </div>

          <button
            type="button"
            className="nav-button next-button"
            onClick={() => navigateMonth(1)}
            aria-label="Next Month"
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
              <path d="M9 18L15 12L9 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
          </button>
        </div>



        {/* Day of week headers */}
        <div className="day-headers">
          {dayNames.map(day => (
            <div key={day} className="day-header">{day}</div>
          ))}
        </div>

        {/* Calendar grid */}
        <div className="calendar-grid">
          {days.map((day, index) => {
            if (day.isEmpty || !day.date) {
              return <div key={index} className="calendar-day empty"></div>;
            }
            
            return (
              <button
                key={index}
                type="button"
                className={`calendar-day ${day.isToday ? 'today' : ''} ${day.isSelected ? 'selected' : ''} ${day.hasData ? 'has-data' : ''}`}
                onClick={(e) => handleDateSelect(day.date, e)}
              >
                <span className="calendar-day-number">{day.date.getDate()}</span>
              </button>
            );
          })}
        </div>



      </div>
        </div>
      </div>

      {/* Year picker as separate card */}
      {showYearPicker && (
        <div className="year-picker">
          <div className="year-picker-content">
            {getYearRange().map(year => (
              <button
                key={year}
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