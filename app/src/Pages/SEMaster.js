import React, { useState, useRef, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import './SEMaster.css';

const SEMaster = () => {
  const navigate = useNavigate();
  const [importFile, setImportFile] = useState(null);
  const [isImporting, setIsImporting] = useState(false);
  const [importStatus, setImportStatus] = useState('');
  const [importedData, setImportedData] = useState([]);
  const [formFields, setFormFields] = useState([]);
  const [formData, setFormData] = useState({});
  const [showForm, setShowForm] = useState(false);
  const [currentRecordIndex, setCurrentRecordIndex] = useState(0);
  const [formHeader, setFormHeader] = useState(null); // Store form header info (FORM - U, EMPLOYEE REGISTER, etc.)
  const [headerFormData, setHeaderFormData] = useState({}); // Store header field values
  const [subColumns, setSubColumns] = useState(null); // Store sub-column structure { mainColumn: [subColumns] }
  const [tableHeaders, setTableHeaders] = useState([]); // Store main column headers for rendering
  const [isSavingToChecklist, setIsSavingToChecklist] = useState(false);
  const [showFormNameModal, setShowFormNameModal] = useState(false);
  const [formNameInput, setFormNameInput] = useState('');
  const [savedForms, setSavedForms] = useState(() => {
    // Initialize saved forms from localStorage immediately
    try {
      const saved = localStorage.getItem('semaster_saved_forms');
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      console.error('Error loading saved forms:', e);
      return [];
    }
  });
  const [showSavedForms, setShowSavedForms] = useState(false); // Don't auto-show, let user choose
  const [isFormFromSaved, setIsFormFromSaved] = useState(false); // Track if form was loaded from saved forms
  const fileInputRef = useRef(null);

  const normalizeHeaderLabel = (label) =>
    String(label || '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');

  const isApprovedFestivalHolidaysField = (label) =>
    normalizeHeaderLabel(label).includes('approvedfestivalholidays');

  const parseFiveValues = (raw) => {
    if (Array.isArray(raw)) {
      const arr = raw.map((v) => String(v ?? '').trim()).filter(Boolean);
      return [...arr.slice(0, 5), '', '', '', '', ''].slice(0, 5);
    }
    const str = String(raw ?? '').trim();
    if (!str) return ['', '', '', '', ''];
    const tokens = str
      .split('|')
      .flatMap((part) => part.split(','))
      .flatMap((part) => part.split(/\s+/))
      .map((t) => t.trim())
      .filter(Boolean);
    return [...tokens.slice(0, 5), '', '', '', '', ''].slice(0, 5);
  };

  const packFiveValues = (arr) => parseFiveValues(arr).join('|');

  // Load saved forms from localStorage on mount and when component becomes visible
  useEffect(() => {
    const loadSavedForms = () => {
      try {
        const saved = localStorage.getItem('semaster_saved_forms');
        if (saved) {
          const parsed = JSON.parse(saved);
          setSavedForms(parsed);
          // Only auto-show saved forms on initial mount, not when user explicitly navigates
          // Don't override user's choice to view import section
        }
      } catch (e) {
        console.error('Error loading saved forms:', e);
      }
    };

    // Load on mount only (not on every state change)
    loadSavedForms();

    // Reload when component becomes visible (user navigates back)
    const handleVisibilityChange = () => {
      if (!document.hidden) {
        loadSavedForms();
      }
    };

    // Reload when window regains focus
    const handleFocus = () => {
      loadSavedForms();
    };

    // Listen for storage changes (in case saved forms are updated in another tab/window)
    const handleStorageChange = (e) => {
      if (e.key === 'semaster_saved_forms') {
        loadSavedForms();
      }
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);
    window.addEventListener('focus', handleFocus);
    window.addEventListener('storage', handleStorageChange);

    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      window.removeEventListener('focus', handleFocus);
      window.removeEventListener('storage', handleStorageChange);
    };
  }, []); // Only run on mount, not on state changes

  // Sync formData with current record when currentRecordIndex or importedData changes
  useEffect(() => {
    if (importedData.length > 0 && currentRecordIndex >= 0 && currentRecordIndex < importedData.length) {
      setFormData(importedData[currentRecordIndex] || {});
    }
  }, [currentRecordIndex, importedData]);

  const handleFileUpload = (file) => {
    if (file) {
      setImportFile(file);
    }
  };

  const handleImport = async () => {
    if (!importFile) return;
   
    setIsImporting(true);
    setImportStatus('Importing Excel file...');
   
    try {
      const reader = new FileReader();
     
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
         
          // Get the first worksheet
          const worksheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[worksheetName];
         
          // Convert worksheet to JSON
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
         
          if (jsonData.length === 0) {
            setImportStatus('Excel file is empty or has no data.');
            setIsImporting(false);
            return;
          }

          // Find the header row by looking for common table header patterns
          // Headers typically contain keywords like S.No, Name, Employee, etc.
          let headerRowIndex = -1;
          let headers = [];
          let subColumnsData = null; // Store sub-column structure
          let subColumnRowIndex = -1; // Track the row index for sub columns
         
          // Search through first 20 rows to find the header row (some forms have headers at row 10-11)
          for (let rowIdx = 0; rowIdx < Math.min(20, jsonData.length); rowIdx++) {
            const row = jsonData[rowIdx] || [];
            if (row.length === 0) continue;
           
            // Check if this row looks like table headers (not header fields)
            // Table headers typically have multiple columns with specific patterns
            const rowText = row.map(cell => String(cell || '').toLowerCase().trim()).join(' ');
           
            // Count how many cells look like table headers (not single-field labels)
            const tableHeaderKeywords = [
              's.no', 'serial', 'serial number',
              'name of the employee', 'employee name',
              'employee identification', 'identification no', 'identification number',
              'gender',
              'date of birth', 'date of joining', 'birth', 'joining',
              'designation',
              'father', 'spouse', 'father / spouse',
              // Form W specific keywords
              'basic wage', 'dearness allowance', 'house rent allowance',
              'overtime wages', 'gross wages', 'net wages', 'deductions',
              'number of days', 'number of days worked', 'provident fund',
              'labour welfare fund', 'advances', 'advance paid', 'advance recovery'
            ];
           
            // Count how many cells in this row match table header keywords
            let matchingHeaderCells = 0;
            for (let cellIdx = 0; cellIdx < row.length; cellIdx++) {
              const cellText = String(row[cellIdx] || '').toLowerCase().trim();
              if (cellText && tableHeaderKeywords.some(keyword => cellText.includes(keyword))) {
                matchingHeaderCells++;
              }
            }
           
            const hasTableHeaderKeywords = matchingHeaderCells > 0;
           
            // Also check if row has multiple text cells (likely table headers, not single field)
            const textCellCount = row.filter(cell => {
              const str = String(cell || '').trim();
              return str && str.length > 0 && !str.match(/^\d+$/); // Not just numbers
            }).length;
           
            // A row is a table header if:
            // 1. It has "S.No" or "S.No." (strong indicator of header row), OR
            // 2. It has 2+ cells matching table header keywords AND multiple columns (at least 3), OR
            // 3. It has 4+ text cells (likely a table header row)
            // Exclude rows that look like single header fields (e.g., "Registration Certificate No:")
            const isSingleFieldRow = textCellCount === 1 && row.length <= 2;
            const hasSNo = rowText.includes('s.no') || rowText.includes('s.no.');
            // Require at least 2 matching header cells OR 4+ text cells OR has S.No to be confident it's a table header
            const isTableHeaderRow = hasSNo ||
                                     (matchingHeaderCells >= 2 && textCellCount >= 3) ||
                                     (textCellCount >= 4 && !isSingleFieldRow);
           
            if (isTableHeaderRow) {
              // Special handling: If this row has "S.No" but the next row also has "S.No" or other main headers,
              // prefer the row with "S.No" as it's more likely the actual header row
              if (hasSNo && rowIdx + 1 < jsonData.length) {
                const nextRow = jsonData[rowIdx + 1] || [];
                const nextRowText = nextRow.map(cell => String(cell || '').toLowerCase().trim()).join(' ');
                const nextRowHasSNo = nextRowText.includes('s.no') || nextRowText.includes('s.no.');
                // If next row also has S.No, check which one has more header-like content
                if (nextRowHasSNo) {
                  const nextRowTextCount = nextRow.filter(cell => {
                    const str = String(cell || '').trim();
                    return str && str.length > 0 && !str.match(/^\d+$/);
                  }).length;
                  // Prefer the row with more text cells (more complete headers)
                  if (nextRowTextCount > textCellCount) {
                    // Skip this row, will process next row instead
                    continue;
                  }
                }
              }
             
              headerRowIndex = rowIdx;
              // Extract headers from this row - use exact Excel headers, preserve all columns
              // Find the maximum column count to preserve all columns
              const maxCols = Math.max(
                row.length,
                ...jsonData.slice(rowIdx + 1, Math.min(rowIdx + 10, jsonData.length))
                  .map(r => r ? r.length : 0)
              );
             
              headers = [];
              // First, check which columns have data in subsequent rows
              const columnsWithData = new Set();
              for (let checkRow = rowIdx + 1; checkRow < Math.min(rowIdx + 20, jsonData.length); checkRow++) {
                const dataRow = jsonData[checkRow] || [];
                for (let colIdx = 0; colIdx < dataRow.length; colIdx++) {
                  const cell = dataRow[colIdx];
                  if (cell !== undefined && cell !== null && String(cell).trim()) {
                    columnsWithData.add(colIdx);
                  }
                }
              }
             
              // Extract all headers from the table header row
              // Only extract from the detected table header row, not from header field rows
              for (let colIdx = 0; colIdx < maxCols; colIdx++) {
                const header = row[colIdx] ? String(row[colIdx]).trim() : '';
               
                // Skip headers that look like header fields (labels ending with colons that are not table headers)
                // Table headers typically don't end with colons, but header fields do
                // However, some table headers might have colons, so we need to be careful
                // Only skip if it's a single-field-like pattern (long text with colon, not a short table header)
                if (header && header.endsWith(':') && header.length > 20) {
                  // This looks like a header field label (e.g., "Registration Certificate No:"), skip it
                  // But only if this row doesn't have multiple table header-like columns
                  // Check if other columns have table header-like text
                  let hasOtherTableHeaders = false;
                  for (let checkCol = 0; checkCol < row.length; checkCol++) {
                    if (checkCol !== colIdx) {
                      const otherHeader = row[checkCol] ? String(row[checkCol]).trim() : '';
                      if (otherHeader &&
                          (otherHeader.toLowerCase().includes('s.no') ||
                           otherHeader.toLowerCase().includes('name') ||
                           otherHeader.toLowerCase().includes('employee') ||
                           otherHeader.toLowerCase().includes('gender') ||
                           otherHeader.toLowerCase().includes('designation'))) {
                        hasOtherTableHeaders = true;
                        break;
                      }
                    }
                  }
                  // If this is the only column or there are no other table headers, skip it
                  if (!hasOtherTableHeaders) {
                    continue;
                  }
                }
               
                // Keep the header if it has text OR if this column has data in subsequent rows
                if (header) {
                  // Always keep headers with text (that aren't header fields)
                  headers.push(header);
                } else if (columnsWithData.has(colIdx)) {
                  // Keep empty headers if column has data
                  headers.push(`Column ${String.fromCharCode(65 + colIdx)}`);
                } else {
                  // For empty headers with no data, still add them to preserve column positions
                  // but we'll trim trailing ones later
                  headers.push('');
                }
              }
             
              // Remove only trailing completely empty headers (no text, no data)
              while (headers.length > 0) {
                const lastIdx = headers.length - 1;
                const lastHeader = headers[lastIdx];
                if (!lastHeader || (!lastHeader.trim() && !columnsWithData.has(lastIdx))) {
                  headers.pop();
                } else {
                  break;
                }
              }
             
              // Before checking for sub-columns, verify we have the right header row
              // If the next row has "S.No", it might be the actual header row
              if (headerRowIndex >= 0 && headerRowIndex + 1 < jsonData.length) {
                const nextRow = jsonData[headerRowIndex + 1] || [];
                const nextRowText = nextRow.map(cell => String(cell || '').toLowerCase().trim()).join(' ');
                const nextRowHasSNo = nextRowText.includes('s.no') || nextRowText.includes('s.no.');
                const currentRowText = jsonData[headerRowIndex].map(cell => String(cell || '').toLowerCase().trim()).join(' ');
                const currentRowHasSNo = currentRowText.includes('s.no') || currentRowText.includes('s.no.');
               
                // If next row has S.No but current doesn't, use next row as header
                if (nextRowHasSNo && !currentRowHasSNo) {
                  // Recalculate columnsWithData for the new header row
                  const newColumnsWithData = new Set();
                  for (let checkRow = headerRowIndex + 2; checkRow < Math.min(headerRowIndex + 22, jsonData.length); checkRow++) {
                    const dataRow = jsonData[checkRow] || [];
                    for (let colIdx = 0; colIdx < dataRow.length; colIdx++) {
                      const cell = dataRow[colIdx];
                      if (cell !== undefined && cell !== null && String(cell).trim()) {
                        newColumnsWithData.add(colIdx);
                      }
                    }
                  }
                 
                  // Re-extract headers from next row
                  const nextRowMaxCols = Math.max(
                    nextRow.length,
                    ...jsonData.slice(headerRowIndex + 1, Math.min(headerRowIndex + 10, jsonData.length))
                      .map(r => r ? r.length : 0)
                  );
                 
                  headers = [];
                  for (let colIdx = 0; colIdx < nextRowMaxCols; colIdx++) {
                    const header = nextRow[colIdx] ? String(nextRow[colIdx]).trim() : '';
                    if (header) {
                      headers.push(header);
                    } else if (newColumnsWithData.has(colIdx)) {
                      headers.push(`Column ${String.fromCharCode(65 + colIdx)}`);
                    } else {
                      headers.push('');
                    }
                  }
                 
                  // Clean up trailing empty headers
                  while (headers.length > 0) {
                    const lastIdx = headers.length - 1;
                    const lastHeader = headers[lastIdx];
                    if (!lastHeader || (!lastHeader.trim() && !newColumnsWithData.has(lastIdx))) {
                      headers.pop();
                    } else {
                      break;
                    }
                  }
                 
                  headerRowIndex = headerRowIndex + 1;
                  columnsWithData = newColumnsWithData; // Update the variable for later use
                }
              }
             
              // Check for sub-columns in the next row
              subColumnsData = null;
              subColumnRowIndex = -1;
              if (headerRowIndex >= 0 && headerRowIndex + 1 < jsonData.length) {
                const subColumnRow = jsonData[headerRowIndex + 1] || [];
                // Check if this row looks like sub-columns (contains numbers or short text)
                const subColumnCells = subColumnRow.map(cell => String(cell || '').trim());
               
                // Count how many cells look like sub-columns
                // Sub-columns can be: numbers, short text, or longer descriptive text that's not a main header
                let subColumnCount = 0;
                let totalCells = 0;
                let hasMainHeaderPattern = false;
               
                subColumnCells.forEach((cell, idx) => {
                  if (cell) {
                    totalCells++;
                    const cellLower = cell.toLowerCase();
                    // Check if this looks like a main header (has common header keywords)
                    const isMainHeaderPattern = cellLower.includes('s.no') ||
                                                (cellLower.includes('name') && cellLower.includes('employee')) ||
                                                (cellLower.includes('employee') && cellLower.includes('identification')) ||
                                                cellLower.includes('gross wages') ||
                                                cellLower.includes('net wages') ||
                                                cellLower.includes('deductions');
                   
                    if (isMainHeaderPattern) {
                      hasMainHeaderPattern = true;
                    }
                   
                    // Sub-column indicators:
                    // 1. Numbers (1-31 for days)
                    const isNumber = /^\d+$/.test(cell);
                    // 2. Short text (up to 10 chars, not "Column X")
                    const isShortText = cell.length > 0 && cell.length <= 10 &&
                                       !cell.includes(':') &&
                                       !cellLower.includes('column') &&
                                       !isMainHeaderPattern;
                    // 3. Longer descriptive text that's not a main header pattern
                    const isDescriptiveText = cell.length > 10 &&
                                            !cell.includes(':') &&
                                            !cellLower.includes('column') &&
                                            !isMainHeaderPattern &&
                                            !cellLower.includes('total') &&
                                            !cellLower.includes('gross') &&
                                            !cellLower.includes('net');
                   
                    if (isNumber || isShortText || isDescriptiveText) {
                      subColumnCount++;
                    }
                  }
                });
               
                // If this row has sub-columns and doesn't look like a main header row, treat it as sub-column row
                // Also check if it aligns with the main headers (has cells in similar positions)
                const hasSubColumns = totalCells > 0 &&
                                     !hasMainHeaderPattern &&
                                     (subColumnCount / totalCells) > 0.3 &&
                                     subColumnCount >= 2;
               
                if (hasSubColumns) {
                  // Build sub-column structure
                  subColumnsData = {};
                  let currentMainColumn = null;
                 
                  // Map sub-columns to main columns based on column position
                  for (let colIdx = 0; colIdx < Math.max(headers.length, subColumnRow.length); colIdx++) {
                    const mainHeader = headers[colIdx] || '';
                    const subCell = subColumnCells[colIdx] || '';
                    const subCellLower = subCell.toLowerCase();
                   
                    // Check if sub-cell is a valid sub-column (not a main header pattern)
                    const isValidSubColumn = subCell &&
                                            !subCell.includes(':') &&
                                            !subCellLower.includes('column') &&
                                            !subCellLower.includes('total') &&
                                            !(subCellLower.includes('name') && subCellLower.includes('employee') && subCellLower.includes('identification')) &&
                                            !subCellLower.includes('gross wages') &&
                                            !subCellLower.includes('net wages');
                   
                    if (mainHeader && mainHeader.trim()) {
                      // New main column found
                      currentMainColumn = mainHeader;
                      if (!subColumnsData[currentMainColumn]) {
                        subColumnsData[currentMainColumn] = [];
                      }
                     
                      // Check if there's a sub-column in this position
                      if (isValidSubColumn) {
                        subColumnsData[currentMainColumn].push(subCell);
                      }
                    } else if (currentMainColumn && isValidSubColumn) {
                      // Continue adding to current main column's sub-columns
                      if (!subColumnsData[currentMainColumn]) {
                        subColumnsData[currentMainColumn] = [];
                      }
                      subColumnsData[currentMainColumn].push(subCell);
                    } else if (subCell && subCell.trim() && !mainHeader) {
                      // Empty main header but has sub-column - might be continuation of previous main column
                      // Only add if we have a current main column
                      if (currentMainColumn && isValidSubColumn) {
                        if (!subColumnsData[currentMainColumn]) {
                          subColumnsData[currentMainColumn] = [];
                        }
                        subColumnsData[currentMainColumn].push(subCell);
                      }
                    }
                  }
                 
                  // Validate: each main column should have at least one sub-column
                  const validMainColumns = Object.keys(subColumnsData).filter(mainCol => subColumnsData[mainCol].length > 0);
                 
                  // If we found valid sub-columns, mark this row and adjust header row index
                  if (validMainColumns.length > 0) {
                    subColumnRowIndex = headerRowIndex + 1;
                    // Update subColumnsData to only include valid main columns
                    const validSubColumnsData = {};
                    validMainColumns.forEach(mainCol => {
                      validSubColumnsData[mainCol] = subColumnsData[mainCol];
                    });
                    subColumnsData = validSubColumnsData;
                  } else {
                    subColumnsData = null; // Not valid sub-columns
                    subColumnRowIndex = -1;
                  }
                }
              }
             
              // Store sub-columns if found
              if (subColumnsData && Object.keys(subColumnsData).length > 0) {
                console.log('Sub-columns detected:', subColumnsData);
                console.log('Header row index:', headerRowIndex);
                console.log('Sub-column row index:', subColumnRowIndex);
                setSubColumns(subColumnsData);
              } else {
                console.log('No sub-columns detected');
                setSubColumns(null);
              }
             
              // If we found headers, break
              if (headers.length > 0) break;
            }
          }
         
          // If no header row found, try using first row if it has text
          if (headerRowIndex === -1 && jsonData.length > 0) {
            const firstRow = jsonData[0] || [];
            const firstRowTextCount = firstRow.filter(cell => {
              const str = String(cell || '').trim();
              return str && str.length > 0;
            }).length;
           
            if (firstRowTextCount >= 2) {
              headerRowIndex = 0;
              headers = firstRow.map((h) => String(h || '').trim()).filter(h => h);
            }
          }
         
          // If still no headers, use first non-empty row
          if (headerRowIndex === -1) {
            for (let rowIdx = 0; rowIdx < jsonData.length; rowIdx++) {
              const row = jsonData[rowIdx] || [];
              const nonEmptyCount = row.filter(cell => String(cell || '').trim()).length;
              if (nonEmptyCount >= 2) {
                headerRowIndex = rowIdx;
                headers = row.map((h) => String(h || '').trim());
                break;
              }
            }
          }
         
          // Extract ALL form header information (rows before the table header)
          // Dynamically read all header rows and extract all information
          let formHeaderInfo = null;
          let headerFields = []; // Store header fields as key-value pairs
          let headerTextRows = []; // Store text-only header rows (titles, subtitles, etc.)
         
          if (headerRowIndex > 0) {
            const headerRows = [];
            for (let i = 0; i < headerRowIndex; i++) {
              const row = jsonData[i] || [];
              const rowText = row.map(cell => String(cell || '').trim()).filter(cell => cell).join(' ');
              if (rowText) {
                headerRows.push(row);
              }
            }
           
            if (headerRows.length > 0) {
              // Extract form title, subtitle, reference, and act line
              let formTitle = '';
              let formSubtitle = '';
              let formReference = '';
              let formActLine = '';
             
              // Process each header row to extract all information
              for (let rowIdx = 0; rowIdx < headerRows.length; rowIdx++) {
                const row = headerRows[rowIdx];
                const rowText = row.map(cell => String(cell || '').trim()).join(' ').toUpperCase();
                const fullRowText = row.map(cell => String(cell || '').trim()).join(' ');
                const rowCells = row.map(cell => String(cell || '').trim());
               
                let rowProcessedAsSpecial = false;
               
                // Check for form title pattern (FORM - X) - usually in right-aligned cells
                if (rowText.includes('FORM') && rowText.includes('-') && !rowText.includes(':')) {
                  // Find the cell containing FORM
                  for (let colIdx = 0; colIdx < row.length; colIdx++) {
                    const cell = rowCells[colIdx];
                    if (cell && cell.toUpperCase().includes('FORM') && cell.includes('-')) {
                      formTitle = cell;
                      rowProcessedAsSpecial = true;
                      break;
                    }
                  }
                }
                // Check for register/subtitle pattern
                // Only catch actual subtitles, not field labels (field labels usually have colons)
                if (!rowProcessedAsSpecial &&
                    (rowText.includes('REGISTER') || rowText.includes('REGISTRATION') ||
                     rowText.includes('FORM') || rowText.includes('REPORT')) &&
                    !rowText.includes(':')) { // Exclude rows with colons (likely field labels)
                  // Find the cell containing the subtitle
                  for (let colIdx = 0; colIdx < row.length; colIdx++) {
                    const cell = rowCells[colIdx];
                    if (cell && (cell.toUpperCase().includes('REGISTER') ||
                                 cell.toUpperCase().includes('REGISTRATION') ||
                                 cell.toUpperCase().includes('FORM') ||
                                 cell.toUpperCase().includes('REPORT'))) {
                      // Make sure it's not a field label (field labels usually end with colon)
                      if (!cell.includes(':') && !cell.toUpperCase().includes('CERTIFICATE NO')) {
                        if (!formSubtitle) {
                          formSubtitle = cell;
                          rowProcessedAsSpecial = true;
                        }
                        break;
                      }
                    }
                  }
                }
                // Check for reference pattern (brackets)
                if (!rowProcessedAsSpecial && fullRowText.includes('[') && fullRowText.includes(']')) {
                  // Find the cell with brackets
                  for (let colIdx = 0; colIdx < row.length; colIdx++) {
                    const cell = rowCells[colIdx];
                    if (cell && cell.includes('[') && cell.includes(']')) {
                      formReference = cell;
                      rowProcessedAsSpecial = true;
                      break;
                    }
                  }
                }
                // Check for ACT & RULES line (4th line)
                if (!rowProcessedAsSpecial &&
                    (rowText.includes('ACT') || rowText.includes('RULES') || rowText.includes('TAMIL NADU') || rowText.includes('INDUSTRIAL ESTABLISHMENTS'))) {
                  // Find the cell containing ACT or RULES
                  if (fullRowText.length > 20 && !fullRowText.includes(':')) {
                    formActLine = fullRowText;
                    rowProcessedAsSpecial = true;
                  }
                }
               
                // Extract header fields from rows, but exclude cells that are already captured as title/subtitle/reference
                // This ensures we don't miss fields like "Registration Certificate No:" but don't duplicate title/subtitle/reference
                for (let colIdx = 0; colIdx < row.length; colIdx++) {
                  const cell = rowCells[colIdx];
                  if (cell && cell.length > 0) {
                    // Skip cells that are already captured as title/subtitle/reference/actLine
                    if (formTitle && cell === formTitle) continue;
                    if (formSubtitle && cell === formSubtitle) continue;
                    if (formReference && cell === formReference) continue;
                    if (formActLine && fullRowText === formActLine) continue;
                   
                    const cellUpper = cell.toUpperCase();
                   
                    // Check if this cell looks like a label
                    // More lenient criteria to catch all fields, but exclude title/subtitle/reference/actLine patterns
                    const isTitleSubtitleRef = (cellUpper.includes('FORM') && cell.includes('-')) ||
                                              (cellUpper.includes('REGISTER') && !cell.includes(':')) ||
                                              (cell.includes('[') && cell.includes(']')) ||
                                              (cellUpper.includes('ACT') || cellUpper.includes('RULES') || cellUpper.includes('TAMIL NADU') || cellUpper.includes('INDUSTRIAL ESTABLISHMENTS'));
                   
                    if (isTitleSubtitleRef && rowProcessedAsSpecial) {
                      // This is already captured as title/subtitle/reference/actLine, skip it
                      continue;
                    }
                   
                    let isLabel = cell.includes(':') ||
                        cellUpper.includes('NAME') ||
                        cellUpper.includes('ADDRESS') ||
                        cellUpper.includes('DATE') ||
                        cellUpper.includes('NO') ||
                        cellUpper.includes('NUMBER') ||
                        cellUpper.includes('MANAGER') ||
                        cellUpper.includes('EMPLOYER') ||
                        cellUpper.includes('ESTABLISHMENT') ||
                        cellUpper.includes('CERTIFICATE') ||
                        cellUpper.includes('REGISTRATION') ||
                        cellUpper.includes('INCHARGE') ||
                        cellUpper.includes('IN-CHARGE');
                   
                    // Only use length check if it's not a title/subtitle/reference pattern
                    if (!isLabel && !isTitleSubtitleRef && cell.length > 8 && !cell.match(/^\d+$/)) {
                      // This might be a label, but be more careful
                      // Only treat as label if it has field-like characteristics
                      if (cell.includes(':') || cellUpper.includes('NO') || cellUpper.includes('NUMBER')) {
                        isLabel = true;
                      }
                    }
                   
                    if (isLabel && !isTitleSubtitleRef) {
                      let label = cell;
                      let value = '';
                     
                      // Look for value in subsequent cells (check up to 3 cells ahead)
                      for (let nextCol = colIdx + 1; nextCol < Math.min(colIdx + 4, row.length); nextCol++) {
                        const nextCell = rowCells[nextCol];
                        if (nextCell && nextCell.length > 0) {
                          const nextCellUpper = nextCell.toUpperCase();
                          // Value should not be another label
                          if (!nextCell.includes(':') &&
                              !nextCellUpper.includes('NAME') &&
                              !nextCellUpper.includes('ADDRESS') &&
                              !nextCellUpper.includes('FORM') &&
                              !nextCellUpper.includes('REGISTER') &&
                              !nextCellUpper.includes('EMPLOYER') &&
                              !nextCellUpper.includes('ESTABLISHMENT') &&
                              !nextCellUpper.includes('MANAGER') &&
                              !nextCellUpper.includes('CERTIFICATE')) {
                            value = nextCell;
                            break;
                          }
                        }
                      }

                      // Special handling for Form V: "Approved Festival Holidays" is typically 5 small boxes
                      // and often appears as values spread across adjacent cells (sometimes on the next row).
                      if (isApprovedFestivalHolidaysField(label)) {
                        const collected = [];
                        const collectFromRow = (r) => {
                          if (!r) return;
                          for (let scanCol = colIdx; scanCol < Math.min(colIdx + 12, r.length); scanCol++) {
                            const v = String(r[scanCol] ?? '').trim();
                            if (!v) continue;
                            // Skip the label cell itself
                            if (scanCol === colIdx && v.toLowerCase().includes('approved')) continue;
                            // Skip the printed 1..5 box labels from the template
                            if (/^[1-5]$/.test(v)) continue;
                            collected.push(v);
                            if (collected.length >= 5) break;
                          }
                        };

                        // Try same row first (sometimes values are to the right)
                        collectFromRow(rowCells);

                        // If still not enough, try the next header row (common in templates)
                        if (collected.length < 5 && rowIdx + 1 < headerRows.length) {
                          const nextRowCells = headerRows[rowIdx + 1].map((c) => String(c || '').trim());
                          collectFromRow(nextRowCells);
                        }

                        value = packFiveValues(collected);
                      }
                     
                      // Make sure we don't duplicate fields
                      const existingField = headerFields.find(f => {
                        const fLabel = f.label.toLowerCase().replace(/[^a-zA-Z0-9]/g, '');
                        const newLabel = label.toLowerCase().replace(/[^a-zA-Z0-9]/g, '');
                        return fLabel === newLabel ||
                               fLabel.includes(newLabel) ||
                               newLabel.includes(fLabel);
                      });
                     
                      if (!existingField && label.length > 2) {
                        headerFields.push({
                          label: label,
                          value: value,
                          key: `header_${rowIdx}_${colIdx}_${label.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase()}`
                        });
                      }
                    }
                  }
                }
               
                // If no fields were found but row has meaningful text, check if it's a text row
                // (only if we didn't already process it as title/subtitle/reference/actLine)
                if (headerFields.length === 0 && fullRowText.length > 3 && !rowProcessedAsSpecial) {
                  // Check if it's not already captured as title/subtitle/reference/actLine
                  if (!formTitle && !formSubtitle && !formReference && !formActLine) {
                    // Store as additional text row (could be subtitle or other info)
                    if (!formSubtitle && fullRowText.length > 5) {
                      formSubtitle = fullRowText;
                    } else {
                      headerTextRows.push(fullRowText);
                    }
                  }
                }
              }
             
              // If we found any header info, store it
              if (formTitle || formSubtitle || formReference || formActLine || headerFields.length > 0 || headerTextRows.length > 0) {
                formHeaderInfo = {
                  title: formTitle,
                  subtitle: formSubtitle,
                  reference: formReference,
                  actLine: formActLine, // Store the ACT & RULES line (4th line)
                  fields: headerFields, // Store header fields
                  textRows: headerTextRows, // Store additional text rows
                  allRows: headerRows // Store all header rows for reference
                };
              }
            }
          }
         
          // Determine start index for data rows
          // If sub-columns were detected, skip both the header row and sub-column row
          let startIndex = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
          if (subColumnRowIndex >= 0) {
            startIndex = subColumnRowIndex + 1;
          } else if (subColumnsData && Object.keys(subColumnsData).length > 0) {
            // Fallback: sub-columns detected but index not captured (shouldn't happen)
            startIndex = headerRowIndex + 2;
          }
         
          // If we still don't have headers, create them from the maximum column count
          if (headers.length === 0 && jsonData.length > 0) {
            const maxCols = Math.max(...jsonData.map(row => row ? row.length : 0));
            // Only create generic headers as last resort
            headers = Array.from({ length: maxCols }, (_, i) => {
              // Try to infer from first data row
              if (startIndex < jsonData.length) {
                const firstDataRow = jsonData[startIndex];
                const cellValue = firstDataRow && firstDataRow[i];
                if (cellValue && String(cellValue).trim()) {
                  // If first cell is a number, might be S.No
                  if (i === 0 && !isNaN(cellValue)) {
                    return 'S.No';
                  }
                }
              }
              return `Column ${String.fromCharCode(65 + i)}`;
            });
          }
         
          // Clean up headers - remove empty ones at the end, but keep all non-empty ones
          while (headers.length > 0 && !headers[headers.length - 1]) {
            headers.pop();
          }
         
          // Process data rows - preserve exact Excel structure
          const processedData = [];
          for (let i = startIndex; i < jsonData.length; i++) {
            const row = jsonData[i] || [];
           
            // Skip completely empty rows
            if (row.length === 0 || row.every(cell => !cell || (typeof cell === 'string' && !cell.trim()))) {
              continue;
            }
           
            // Create record with exact header names from Excel
            const record = { id: processedData.length + 1 };
           
            // Map each header to its corresponding column value
            // If sub-columns exist, map data to sub-column field names
            if (subColumnsData && Object.keys(subColumnsData).length > 0) {
              // When sub-columns exist, the Excel data has one cell per sub-column
              // We need to map based on the actual column positions in Excel
              let excelColIndex = 0;
             
              headers.forEach((header, headerIndex) => {
                const mainColumnSubCols = subColumnsData[header] || [];
                if (mainColumnSubCols.length > 0) {
                  // This main column has sub-columns
                  // Map each sub-column from the Excel row
                  mainColumnSubCols.forEach((subCol, subIdx) => {
                    const subColName = `${header}_${subCol}`;
                    // Use the current Excel column index
                    const cellValue = row[excelColIndex];
                    if (cellValue !== undefined && cellValue !== null) {
                      record[subColName] = String(cellValue).trim();
                    } else {
                      record[subColName] = '';
                    }
                    excelColIndex++;
                  });
                } else {
                  // Regular column without sub-columns
                  const cellValue = row[excelColIndex];
                  if (cellValue !== undefined && cellValue !== null) {
                    record[header] = String(cellValue).trim();
                  } else {
                    record[header] = '';
                  }
                  excelColIndex++;
                }
              });
             
              // Debug: Log first few records to verify mapping
              if (processedData.length < 3) {
                console.log(`Record ${processedData.length + 1} mapping:`, {
                  rowLength: row.length,
                  recordKeys: Object.keys(record),
                  sampleValues: Object.entries(record).slice(0, 5)
                });
              }
            } else {
              // No sub-columns, map normally - one header per column
              headers.forEach((header, index) => {
                const cellValue = row[index];
                if (cellValue !== undefined && cellValue !== null) {
                  record[header] = String(cellValue).trim();
                } else {
                  record[header] = '';
                }
              });
            }
           
            // Only add record if it has at least one non-empty value
            const hasData = Object.values(record).some((value, idx) => {
              // Skip the 'id' field
              if (idx === 0) return false;
              return value && String(value).trim().length > 0;
            });
           
            if (hasData) {
              processedData.push(record);
            }
          }

          if (processedData.length === 0) {
            setImportStatus('No valid data found in the Excel file.');
            setIsImporting(false);
            return;
          }

          // Create form fields from headers - use exact Excel headers
          // If sub-columns exist, create fields for main columns with sub-column info
          const fields = [];
          if (subColumnsData && Object.keys(subColumnsData).length > 0) {
            // Create fields for main columns with sub-columns
            headers.forEach((header, index) => {
              if (header && header.trim()) {
                const mainColumnSubCols = subColumnsData[header] || [];
                if (mainColumnSubCols.length > 0) {
                  // Main column with sub-columns
                  fields.push({
                    name: header,
                    label: header,
                    type: 'text',
                    placeholder: `Enter ${header}`,
                    isMainColumn: true,
                    subColumns: mainColumnSubCols
                  });
                 
                  // Create individual fields for each sub-column
                  mainColumnSubCols.forEach((subCol, subIdx) => {
                    const subColName = `${header}_${subCol}`;
                    fields.push({
                      name: subColName,
                      label: subCol,
                      type: 'text',
                      placeholder: `Enter ${subCol}`,
                      isSubColumn: true,
                      mainColumn: header,
                      subColumnIndex: subIdx
                    });
                  });
                } else {
                  // Regular column without sub-columns
                  fields.push({
                    name: header,
                    label: header,
                    type: 'text',
                    placeholder: `Enter ${header}`
                  });
                }
              }
            });
          } else {
            // No sub-columns, create regular fields
            fields.push(...headers
              .filter(header => header && header.trim())
              .map(header => ({
                name: header,
                label: header,
                type: 'text',
                placeholder: `Enter ${header}`
              }))
            );
          }

          // Set the first record as initial form data
          let initialFormData = processedData[0] || {};

          // Initialize header form data
          const initialHeaderData = {};
          if (formHeaderInfo && formHeaderInfo.fields) {
            formHeaderInfo.fields.forEach(field => {
              initialHeaderData[field.key] = field.value || '';
            });
          }

          // Fetch employee data and populate matching fields
          setImportStatus('Fetching employee data to populate matching fields...');
          let employees = null; // Declare at higher scope for localStorage fallback
          try {
            // Fetch all employee data (no pagination parameters to get all records)
            let employeeResponse = await fetch('/server/employeedata_function/employee');
           
            // If response is not ok, try with explicit parameters to get all records
            if (!employeeResponse.ok) {
              console.log('First fetch failed, trying with explicit parameters...');
              employeeResponse = await fetch('/server/employeedata_function/employee?page=1&perPage=10000');
            }
           
            if (employeeResponse.ok) {
              const employeeData = await employeeResponse.json();
              console.log('Employee data response:', {
                status: employeeData.status,
                hasData: !!employeeData.data,
                hasEmployees: !!(employeeData.data && employeeData.data.employees),
                employeeCount: employeeData.data?.employees?.length || 0,
                total: employeeData.data?.total || 0
              });
             
              if (employeeData.status === 'success' && employeeData.data && employeeData.data.employees) {
                employees = employeeData.data.employees;
                const total = employeeData.data.total || employees.length;
               
                // If we got paginated results and there are more, fetch all pages
                if (employeeData.data.hasMore && total > employees.length) {
                  console.log(`Fetching all employees (${total} total, have ${employees.length})...`);
                  const allEmployees = [...employees];
                  let page = 2;
                  while (allEmployees.length < total) {
                    const pageResponse = await fetch(`/server/employeedata_function/employee?page=${page}&perPage=10000`);
                    if (pageResponse.ok) {
                      const pageData = await pageResponse.json();
                      if (pageData.status === 'success' && pageData.data && pageData.data.employees) {
                        allEmployees.push(...pageData.data.employees);
                        page++;
                        if (!pageData.data.hasMore) break;
                      } else {
                        break;
                      }
                    } else {
                      break;
                    }
                  }
                  employees = allEmployees;
                }
               
                console.log(`Fetched ${employees.length} employees from server (total available: ${total})`);
                console.log(`Processing ${processedData.length} records from Excel file`);
                console.log(`Will attempt to match and populate employee data for all ${processedData.length} records`);
               
                if (employees.length === 0) {
                  console.log('No employees found in database');
                  setImportStatus(`Successfully imported ${processedData.length} records. No employee data found to populate.`);
                } else {
               
                // Create a mapping function to match form field names with employee data fields
                const normalizeFieldName = (fieldName) => {
                  if (!fieldName) return '';
                  return String(fieldName).toLowerCase()
                    .replace(/[^a-z0-9]/g, '')
                    .trim();
                };

                // Field mapping: form field variations -> employee data field names
                const fieldMapping = {
                  'nameoftheemployee': ['name', 'employee', 'nameoftheemployee', 'employeename', 'nameofthe'],
                  'sex': ['sex', 'gender'],
                  'fathersname': ['father', 'fathersname', 'fathername', 'fathers', 'spouse'],
                  'designation': ['designation'],
                  'employeeidentificationno': ['employeenumber', 'employeeno', 'empno', 'empnumber', 'employeeid', 'employeeidno', 'identificationno', 'identificationnumber', 'employeeidentificationno'],
                  'dateofbirth': ['dateofbirth', 'dob', 'birthdate', 'dateofbirth'],
                  'dateofjoining': ['dateofjoining', 'doj', 'joiningdate', 'dateofjoining'],
                  'presentaddress': ['presentaddress', 'present', 'currentaddress'],
                  'permanentaddress': ['permanentaddress', 'permanent'],
                  'aadhaarno': ['aadhaar', 'aadhaarno', 'aadhar', 'aadharno', 'aadhaarnumber']
                };

                // Function to find matching employee data field for a form field
                const findMatchingEmployeeField = (formFieldName) => {
                  if (!formFieldName) return null;
                 
                  const normalizedFormField = normalizeFieldName(formFieldName);
                 
                  // PRIORITY: Check for "Name of the employee" first (exact match or normalized)
                  if (formFieldName === 'Name of the employee' ||
                      formFieldName.toLowerCase() === 'name of the employee' ||
                      normalizedFormField === 'nameoftheemployee') {
                    return 'nameoftheemployee';
                  }
                 
                  // Check for Employee Identification No
                  if (normalizedFormField.includes('employeeidentificationno') ||
                      normalizedFormField.includes('employeeidentification') ||
                      (normalizedFormField.includes('employee') && normalizedFormField.includes('identification'))) {
                    return 'employeeidentificationno';
                  }
                 
                  // Direct field name mappings (exact matches first)
                  const directMappings = {
                    'nameoftheemployee': 'nameoftheemployee',
                    'nameofthe': 'nameoftheemployee',
                    'employeename': 'nameoftheemployee',
                    'gender': 'sex',
                    'sex': 'sex',
                    'fathersname': 'fathersname',
                    'fathername': 'fathersname',
                    'father': 'fathersname',
                    'spouse': 'fathersname',
                    'employeeidentificationno': 'employeeidentificationno',
                    'employeenumber': 'employeeidentificationno',
                    'employeeno': 'employeeidentificationno',
                    'empno': 'employeeidentificationno',
                    'employeeid': 'employeeidentificationno',
                    'identificationno': 'employeeidentificationno',
                    'identificationnumber': 'employeeidentificationno',
                    'dateofbirth': 'dateofbirth',
                    'dob': 'dateofbirth',
                    'birthdate': 'dateofbirth',
                    'dateofjoining': 'dateofjoining',
                    'doj': 'dateofjoining',
                    'joiningdate': 'dateofjoining',
                    'dateofexit': 'dateofexit',
                    'exitdate': 'dateofexit',
                    'presentaddress': 'presentaddress',
                    'present': 'presentaddress',
                    'currentaddress': 'presentaddress',
                    'permanentaddress': 'permanentaddress',
                    'permanent': 'permanentaddress',
                    'aadhaarno': 'aadhaarno',
                    'aadhaar': 'aadhaarno',
                    'aadhar': 'aadhaarno',
                    'aadharno': 'aadhaarno',
                    'designation': 'designation'
                  };
                 
                  // Check direct mappings first
                  if (directMappings[normalizedFormField]) {
                    return directMappings[normalizedFormField];
                  }
                 
                  // Check field mapping with variations
                  for (const [employeeField, variations] of Object.entries(fieldMapping)) {
                    for (const variation of variations) {
                      if (normalizedFormField.includes(variation) || variation.includes(normalizedFormField)) {
                        return employeeField;
                      }
                      if (normalizedFormField === variation) {
                        return employeeField;
                      }
                    }
                  }
                 
                  return null;
                };

                // Populate matching fields for all records
                const enrichedData = processedData.map((record, recordIdx) => {
                  const enrichedRecord = { ...record };
                  let hasMatch = false;
                  let populatedFields = [];
                 
                  // Try to find a matching employee by name or employee identification number
                  let matchingEmployee = null;
                 
                  // Find name field - check multiple variations including sub-columns
                  // First try main column, then try sub-columns
                  let nameField = fields.find(f => {
                    const fieldNameLower = f.name.toLowerCase();
                    const normalized = normalizeFieldName(f.name);
                    return f.name === 'Name of the employee' ||
                           fieldNameLower === 'name of the employee' ||
                           (fieldNameLower.includes('name of the employee') && !f.name.includes('_')) ||
                           (normalized.includes('name') && normalized.includes('employee') && !normalized.includes('identification') && !normalized.includes('number') && !f.name.includes('_'));
                  });
                 
                  // If main column not found or empty, try sub-columns
                  if (!nameField || !record[nameField.name] || !String(record[nameField.name]).trim()) {
                    nameField = fields.find(f => {
                      const fieldNameLower = f.name.toLowerCase();
                      const normalized = normalizeFieldName(f.name);
                      return (fieldNameLower.includes('name of the employee') ||
                              (normalized.includes('name') && normalized.includes('employee') && !normalized.includes('identification') && !normalized.includes('number'))) &&
                             f.name.includes('_'); // Sub-column has underscore
                    });
                  }
                 
                  // Find employee identification number field - check multiple variations including sub-columns
                  let empIdField = fields.find(f => {
                    const fieldNameLower = f.name.toLowerCase();
                    const normalized = normalizeFieldName(f.name);
                    return (fieldNameLower.includes('employee identification no') ||
                           fieldNameLower.includes('employeeidentificationno') ||
                           fieldNameLower.includes('employee identification') ||
                           normalized.includes('employeeidentificationno') ||
                           (normalized.includes('employee') && normalized.includes('identification'))) &&
                           !f.name.includes('_'); // Main column doesn't have underscore
                  });
                 
                  // If main column not found or empty, try sub-columns
                  if (!empIdField || !record[empIdField.name] || !String(record[empIdField.name]).trim()) {
                    empIdField = fields.find(f => {
                      const fieldNameLower = f.name.toLowerCase();
                      const normalized = normalizeFieldName(f.name);
                      return (fieldNameLower.includes('employee identification no') ||
                             fieldNameLower.includes('employeeidentificationno') ||
                             fieldNameLower.includes('employee identification') ||
                             normalized.includes('employeeidentificationno') ||
                             (normalized.includes('employee') && normalized.includes('identification'))) &&
                             f.name.includes('_'); // Sub-column has underscore
                    });
                  }
                 
                  // Debug: Log found fields (only for first record to avoid spam)
                  if (recordIdx === 0) {
                    console.log('\n=== FIELD MATCHING DEBUG ===');
                    console.log('All form fields:', fields.map(f => `"${f.name}"`).join(', '));
                    console.log(`Name field found: ${nameField ? `"${nameField.name}"` : 'NOT FOUND'}`);
                    if (nameField) {
                      console.log(`  Name field value: "${record[nameField.name] || ''}"`);
                    }
                    console.log(`Employee ID field found: ${empIdField ? `"${empIdField.name}"` : 'NOT FOUND'}`);
                    if (empIdField) {
                      console.log(`  Employee ID field value: "${record[empIdField.name] || ''}"`);
                    }
                    if (employees.length > 0) {
                      console.log('Available employee fields:', Object.keys(employees[0] || {}).join(', '));
                      console.log('Sample employee:', employees[0]);
                    }
                    console.log('============================\n');
                  }
                 
                  // Try to match by employee name - check both main column and sub-columns
                  if (nameField) {
                    const recordName = String(record[nameField.name] || '').trim();
                    if (recordName) {
                      const recordNameLower = recordName.toLowerCase();
                      console.log(`  Trying to match by name: "${recordName}"`);
                      matchingEmployee = employees.find(emp => {
                        if (!emp.nameoftheemployee) return false;
                        const empName = String(emp.nameoftheemployee).trim().toLowerCase();
                        // Try exact match first
                        if (empName === recordNameLower) return true;
                        // Try partial match
                        if (empName.includes(recordNameLower) || recordNameLower.includes(empName)) {
                          // Make sure it's a meaningful match (at least 3 characters)
                          if (recordNameLower.length >= 3 || empName.length >= 3) {
                            return true;
                          }
                        }
                        return false;
                      });
                      if (matchingEmployee) {
                        console.log(`  ✓ Matched by name: "${recordName}" -> "${matchingEmployee.nameoftheemployee}"`);
                      } else {
                        console.log(`  ✗ No match found for name: "${recordName}"`);
                      }
                    }
                  }
                 
                  // If no match by name, try by employee identification number - check both main column and sub-columns
                  if (!matchingEmployee && empIdField) {
                    const recordEmpId = String(record[empIdField.name] || '').trim();
                    if (recordEmpId) {
                      console.log(`  Trying to match by Employee ID: "${recordEmpId}"`);
                      matchingEmployee = employees.find(emp => {
                        if (!emp.employeeidentificationno) return false;
                        const empId = String(emp.employeeidentificationno).trim();
                        return empId === recordEmpId;
                      });
                      if (matchingEmployee) {
                        console.log(`  ✓ Matched by Employee ID: "${recordEmpId}" -> "${matchingEmployee.nameoftheemployee}"`);
                      } else {
                        console.log(`  ✗ No match found for Employee ID: "${recordEmpId}"`);
                      }
                    }
                  }
                 
                  // If no match found and both name and ID fields are empty, try to populate with employee data in order
                  // This helps when Excel has empty rows that need to be filled
                  if (!matchingEmployee && employees.length > 0) {
                    // If both name and ID are empty, assign employees in order (record 1 = employee 1, etc.)
                    if ((!nameField || !record[nameField.name] || !String(record[nameField.name]).trim()) &&
                        (!empIdField || !record[empIdField.name] || !String(record[empIdField.name]).trim())) {
                      // Assign employee by record index (if we have enough employees)
                      if (recordIdx < employees.length) {
                        matchingEmployee = employees[recordIdx];
                        console.log(`  ℹ No match found, assigning employee by index: "${matchingEmployee.nameoftheemployee}" (Record ${recordIdx + 1} → Employee ${recordIdx + 1})`);
                      } else {
                        console.log(`  ⚠ No match found and no employee available for index ${recordIdx + 1} (only ${employees.length} employees available)`);
                      }
                    }
                  }
                 
                  // If we found a matching employee, populate ALL matching fields (including sub-columns)
                  if (matchingEmployee) {
                    hasMatch = true;
                    console.log(`\nRecord ${recordIdx + 1}: Matched employee "${matchingEmployee.nameoftheemployee}"`);
                    console.log(`  Employee data available:`, Object.keys(matchingEmployee).filter(k => matchingEmployee[k]).join(', '));
                   
                    // Populate ALL matching fields from employee data (both main columns and sub-columns)
                    fields.forEach(field => {
                      const employeeField = findMatchingEmployeeField(field.name);
                      if (employeeField) {
                        const currentValue = String(enrichedRecord[field.name] || '').trim();
                        const employeeValue = matchingEmployee[employeeField];
                       
                        // Always populate from employee data if available (overwrite existing values)
                        if (employeeValue !== null && employeeValue !== undefined && employeeValue !== '') {
                          const employeeValueStr = String(employeeValue).trim();
                          if (employeeValueStr) {
                            enrichedRecord[field.name] = employeeValueStr;
                            populatedFields.push(`${field.name} → ${employeeField}: "${employeeValueStr}"`);
                            console.log(`  ✓ Populated "${field.name}" with "${employeeValueStr}" from ${employeeField}${currentValue ? ` (was: "${currentValue}")` : ''}`);
                          }
                        }
                      }
                    });
                   
                    // Also populate sub-column fields that correspond to main columns
                    // For example, if "Name of the employee" is populated, also populate "Name of the employee_2"
                    if (matchingEmployee.nameoftheemployee) {
                      const nameValue = String(matchingEmployee.nameoftheemployee).trim();
                      fields.forEach(field => {
                        if (field.name.includes('Name of the employee') && field.name.includes('_')) {
                          const currentValue = String(enrichedRecord[field.name] || '').trim();
                          if (!currentValue && nameValue) {
                            enrichedRecord[field.name] = nameValue;
                            populatedFields.push(`${field.name} → nameoftheemployee: "${nameValue}"`);
                            console.log(`  ✓ Populated sub-column "${field.name}" with "${nameValue}"`);
                          }
                        }
                      });
                    }
                   
                    if (matchingEmployee.employeeidentificationno) {
                      const empIdValue = String(matchingEmployee.employeeidentificationno).trim();
                      fields.forEach(field => {
                        if ((field.name.includes('Employee Identification No') || field.name.includes('EmployeeIdentificationNo')) &&
                            field.name.includes('_')) {
                          const currentValue = String(enrichedRecord[field.name] || '').trim();
                          if (!currentValue && empIdValue) {
                            enrichedRecord[field.name] = empIdValue;
                            populatedFields.push(`${field.name} → employeeidentificationno: "${empIdValue}"`);
                            console.log(`  ✓ Populated sub-column "${field.name}" with "${empIdValue}"`);
                          }
                        }
                      });
                    }
                   
                    // Populate other matching sub-columns
                    fields.forEach(field => {
                      if (field.name.includes('_')) {
                        const employeeField = findMatchingEmployeeField(field.name);
                        if (employeeField) {
                          const currentValue = String(enrichedRecord[field.name] || '').trim();
                          const employeeValue = matchingEmployee[employeeField];
                          if (!currentValue && employeeValue !== null && employeeValue !== undefined && employeeValue !== '') {
                            const employeeValueStr = String(employeeValue).trim();
                            if (employeeValueStr) {
                              enrichedRecord[field.name] = employeeValueStr;
                              populatedFields.push(`${field.name} → ${employeeField}: "${employeeValueStr}"`);
                              console.log(`  ✓ Populated sub-column "${field.name}" with "${employeeValueStr}"`);
                            }
                          }
                        }
                      }
                    });
                   
                    if (populatedFields.length > 0) {
                      console.log(`✓ Record ${recordIdx + 1}: Successfully populated ${populatedFields.length} field(s)`);
                    } else {
                      console.log(`⚠ Record ${recordIdx + 1}: Matched employee but no fields were populated`);
                    }
                  } else {
                    console.log(`\n✗ Record ${recordIdx + 1}: No matching employee found`);
                    if (nameField) {
                      const nameValue = String(record[nameField.name] || '').trim();
                      console.log(`  Name field: "${nameField.name}" = "${nameValue}"`);
                      if (!nameValue) {
                        // Check sub-columns
                        const nameSubCols = fields.filter(f => f.name.includes('Name of the employee') && f.name.includes('_'));
                        nameSubCols.forEach(subCol => {
                          const subValue = String(record[subCol.name] || '').trim();
                          if (subValue) {
                            console.log(`    Sub-column "${subCol.name}" = "${subValue}"`);
                          }
                        });
                      }
                    }
                    if (empIdField) {
                      const idValue = String(record[empIdField.name] || '').trim();
                      console.log(`  Employee ID field: "${empIdField.name}" = "${idValue}"`);
                      if (!idValue) {
                        // Check sub-columns
                        const idSubCols = fields.filter(f => (f.name.includes('Employee Identification No') || f.name.includes('EmployeeIdentificationNo')) && f.name.includes('_'));
                        idSubCols.forEach(subCol => {
                          const subValue = String(record[subCol.name] || '').trim();
                          if (subValue) {
                            console.log(`    Sub-column "${subCol.name}" = "${subValue}"`);
                          }
                        });
                      }
                    }
                  }
                 
                  return enrichedRecord;
                });

                // Update processedData with enriched data
                processedData.splice(0, processedData.length, ...enrichedData);
               
                // Update initial form data with enriched first record
                initialFormData = enrichedData[0] || {};
               
                // Count how many records were matched and how many fields were populated
                let totalFieldsPopulated = 0;
                const matchedCount = enrichedData.filter((record, idx) => {
                  const nameField = fields.find(f => {
                    const normalized = normalizeFieldName(f.name);
                    return normalized.includes('name') && normalized.includes('employee');
                  });
                  const empIdField = fields.find(f => {
                    const normalized = normalizeFieldName(f.name);
                    return normalized.includes('employeeidentificationno') || normalized.includes('employeenumber');
                  });
                 
                  let wasMatched = false;
                  if (nameField && record[nameField.name]) {
                    const recordName = String(record[nameField.name]).trim().toLowerCase();
                    wasMatched = employees.some(emp => {
                      if (!emp.nameoftheemployee) return false;
                      return String(emp.nameoftheemployee).trim().toLowerCase() === recordName;
                    });
                  }
                  if (!wasMatched && empIdField && record[empIdField.name]) {
                    const recordEmpId = String(record[empIdField.name]).trim();
                    wasMatched = employees.some(emp => {
                      if (!emp.employeeidentificationno) return false;
                      return String(emp.employeeidentificationno).trim() === recordEmpId;
                    });
                  }
                 
                  // Count populated fields for this record
                  if (wasMatched) {
                    fields.forEach(field => {
                      const employeeField = findMatchingEmployeeField(field.name);
                      if (employeeField && record[field.name]) {
                        const value = String(record[field.name] || '').trim();
                        if (value) {
                          totalFieldsPopulated++;
                        }
                      }
                    });
                  }
                 
                  return wasMatched;
                }).length;
               
                // Log comprehensive summary
                console.log('\n=== EMPLOYEE DATA POPULATION SUMMARY ===');
                console.log(`Total records imported: ${processedData.length}`);
                console.log(`Total employees available: ${employees.length}`);
                console.log(`Records matched: ${matchedCount} out of ${processedData.length}`);
                console.log(`Total fields populated: ${totalFieldsPopulated}`);
                console.log('==========================================\n');
               
                if (matchedCount > 0) {
                  setImportStatus(`Successfully imported ${processedData.length} record(s). Employee data populated for ${matchedCount} record(s) (${totalFieldsPopulated} fields). Showing record 1 of ${processedData.length}.`);
                } else {
                  setImportStatus(`Successfully imported ${processedData.length} record(s). No employee matches found. Check browser console for details. Showing record 1 of ${processedData.length}.`);
                }
                }
              } else {
                console.log('Employee data response structure:', employeeData);
                setImportStatus(`Successfully imported ${processedData.length} records. Employee data format unexpected. Showing record 1 of ${processedData.length}.`);
              }
            } else {
              const errorText = await employeeResponse.text();
              console.error('Employee data fetch failed:', employeeResponse.status, errorText);
             
              // Try to load from localStorage as fallback
              try {
                const localData = localStorage.getItem('employeesData');
                if (localData) {
                  const parsedData = JSON.parse(localData);
                  if (Array.isArray(parsedData) && parsedData.length > 0) {
                    console.log(`Loaded ${parsedData.length} employees from localStorage (server unavailable)`);
                    employees = parsedData;
                   
                    // Process with localStorage data (same logic as above)
                    const normalizeFieldName = (fieldName) => {
                      if (!fieldName) return '';
                      return String(fieldName).toLowerCase().replace(/[^a-z0-9]/g, '').trim();
                    };
                   
                    const findMatchingEmployeeField = (formFieldName) => {
                      if (!formFieldName) return null;
                      const normalizedFormField = normalizeFieldName(formFieldName);
                      if (formFieldName === 'Name of the employee' || normalizedFormField === 'nameoftheemployee') {
                        return 'nameoftheemployee';
                      }
                      if (normalizedFormField.includes('employeeidentificationno') || normalizedFormField.includes('employeeidentification')) {
                        return 'employeeidentificationno';
                      }
                      const directMappings = {
                        'nameoftheemployee': 'nameoftheemployee',
                        'gender': 'sex',
                        'sex': 'sex',
                        'fathersname': 'fathersname',
                        'spouse': 'fathersname',
                        'employeeidentificationno': 'employeeidentificationno',
                        'dateofbirth': 'dateofbirth',
                        'dateofjoining': 'dateofjoining',
                        'presentaddress': 'presentaddress',
                        'permanentaddress': 'permanentaddress',
                        'aadhaarno': 'aadhaarno',
                        'designation': 'designation'
                      };
                      return directMappings[normalizedFormField] || null;
                    };
                   
                    const enrichedData = processedData.map((record) => {
                      const enrichedRecord = { ...record };
                      const nameField = fields.find(f => normalizeFieldName(f.name).includes('name') && normalizeFieldName(f.name).includes('employee'));
                      const empIdField = fields.find(f => normalizeFieldName(f.name).includes('employeeidentificationno') || normalizeFieldName(f.name).includes('employeenumber'));
                     
                      let matchingEmployee = null;
                      if (nameField && record[nameField.name]) {
                        const recordName = String(record[nameField.name]).trim().toLowerCase();
                        matchingEmployee = employees.find(emp => {
                          if (!emp.nameoftheemployee) return false;
                          return String(emp.nameoftheemployee).trim().toLowerCase() === recordName;
                        });
                      }
                      if (!matchingEmployee && empIdField && record[empIdField.name]) {
                        const recordEmpId = String(record[empIdField.name]).trim();
                        matchingEmployee = employees.find(emp => {
                          if (!emp.employeeidentificationno) return false;
                          return String(emp.employeeidentificationno).trim() === recordEmpId;
                        });
                      }
                     
                      if (matchingEmployee) {
                        fields.forEach(field => {
                          const employeeField = findMatchingEmployeeField(field.name);
                          if (employeeField && (!enrichedRecord[field.name] || !String(enrichedRecord[field.name]).trim())) {
                            const employeeValue = matchingEmployee[employeeField];
                            if (employeeValue !== null && employeeValue !== undefined && employeeValue !== '') {
                              enrichedRecord[field.name] = String(employeeValue).trim();
                            }
                          }
                        });
                      }
                      return enrichedRecord;
                    });
                   
                    processedData.splice(0, processedData.length, ...enrichedData);
                    initialFormData = enrichedData[0] || {};
                    setImportStatus(`Successfully imported ${processedData.length} records. Employee data populated from local storage. Showing record 1 of ${processedData.length}.`);
                  } else {
                    setImportStatus(`Successfully imported ${processedData.length} records. Could not fetch employee data (${employeeResponse.status}). Showing record 1 of ${processedData.length}.`);
                  }
                } else {
                  setImportStatus(`Successfully imported ${processedData.length} records. Could not fetch employee data (${employeeResponse.status}). Showing record 1 of ${processedData.length}.`);
                }
              } catch (localError) {
                console.error('Error loading from localStorage:', localError);
                setImportStatus(`Successfully imported ${processedData.length} records. Could not fetch employee data (${employeeResponse.status}). Showing record 1 of ${processedData.length}.`);
              }
            }
          } catch (employeeError) {
            console.error('Error fetching employee data:', employeeError);
            setImportStatus(`Successfully imported ${processedData.length} records. Error fetching employee data: ${employeeError.message}. Showing record 1 of ${processedData.length}.`);
          }
         
          setImportedData(processedData);
          setFormFields(fields);
          setTableHeaders(headers); // Store main column headers for rendering
          setFormData(initialFormData);
          setCurrentRecordIndex(0);
          setFormHeader(formHeaderInfo); // Store form header info
          setHeaderFormData(initialHeaderData); // Store header field values
          setShowForm(true);
          setIsFormFromSaved(false); // Reset flag when importing new file
         
          // Reset file input
          setTimeout(() => {
            setImportFile(null);
            setImportStatus('');
          }, 3000);
         
        } catch (parseError) {
          console.error('Error parsing Excel file:', parseError);
          setImportStatus('Error parsing Excel file. Please check the file format.');
        } finally {
          setIsImporting(false);
        }
      };
     
      reader.onerror = () => {
        setImportStatus('Error reading file. Please try again.');
        setIsImporting(false);
      };
     
      reader.readAsArrayBuffer(importFile);
     
    } catch (error) {
      console.error('Import error:', error);
      setImportStatus('Import failed. Please try again.');
      setIsImporting(false);
    }
  };

  const handleFormFieldChange = (fieldName, value) => {
    setFormData(prev => ({
      ...prev,
      [fieldName]: value
    }));
  };

  // Handle field change for a specific row in the multi-row table
  const handleRowFieldChange = (rowIndex, fieldName, value) => {
    setImportedData(prev => {
      const updated = [...prev];
      if (updated[rowIndex]) {
        updated[rowIndex] = {
          ...updated[rowIndex],
          [fieldName]: value
        };
      }
      return updated;
    });
    // Also update formData if this is the current record
    if (rowIndex === currentRecordIndex) {
      setFormData(prev => ({
        ...prev,
        [fieldName]: value
      }));
    }
  };

  const handleHeaderFieldChange = (fieldKey, value) => {
    setHeaderFormData(prev => ({
      ...prev,
      [fieldKey]: value
    }));
  };

  const handleFestivalHolidayBoxChange = (fieldKey, index, value) => {
    setHeaderFormData((prev) => {
      const current = parseFiveValues(prev[fieldKey] || '');
      const updated = [...current];
      updated[index] = String(value ?? '').trim();
      return {
        ...prev,
        [fieldKey]: packFiveValues(updated)
      };
    });
  };

  const handlePrevious = () => {
    if (currentRecordIndex > 0) {
      const newIndex = currentRecordIndex - 1;
      setCurrentRecordIndex(newIndex);
      setFormData(importedData[newIndex]);
    }
  };

  const handleNext = () => {
    if (currentRecordIndex < importedData.length - 1) {
      const newIndex = currentRecordIndex + 1;
      setCurrentRecordIndex(newIndex);
      setFormData(importedData[newIndex]);
    }
  };

  const handleSave = () => {
    // Update the current record in importedData
    const updatedData = [...importedData];
    updatedData[currentRecordIndex] = { ...formData };
    setImportedData(updatedData);
    setImportStatus(`Record ${currentRecordIndex + 1} saved successfully!`);
    setTimeout(() => setImportStatus(''), 3000);
  };

  // Add a new row with all columns from imported Excel
  const handleAddRow = () => {
    if (formFields.length === 0 || tableHeaders.length === 0) {
      setImportStatus('No columns available. Please import an Excel file first.');
      setTimeout(() => setImportStatus(''), 3000);
      return;
    }

    // Create a new record with all fields initialized to empty
    // Use the same structure as records created during import
    const newRecord = { id: importedData.length + 1 };

    // Initialize fields based on the same logic used during import
    if (subColumns && Object.keys(subColumns).length > 0) {
      // Handle sub-columns structure
      tableHeaders.forEach((header) => {
        const mainColumnSubCols = subColumns[header] || [];
        if (mainColumnSubCols.length > 0) {
          // This main column has sub-columns - initialize each sub-column
          mainColumnSubCols.forEach((subCol) => {
            const subColName = `${header}_${subCol}`;
            newRecord[subColName] = '';
          });
        } else {
          // Regular column without sub-columns
          newRecord[header] = '';
        }
      });
    } else {
      // No sub-columns, initialize all headers as regular fields
      tableHeaders.forEach((header) => {
        newRecord[header] = '';
      });
    }

    // Add the new record to importedData
    const updatedData = [...importedData, newRecord];
    setImportedData(updatedData);
   
    // Set the new row as the current record
    const newIndex = updatedData.length - 1;
    setCurrentRecordIndex(newIndex);
    setFormData(newRecord);
   
    setImportStatus(`New row added successfully! (Row ${newIndex + 1})`);
    setTimeout(() => setImportStatus(''), 3000);
  };

  const handleReset = () => {
    setImportedData([]);
    setFormFields([]);
    setTableHeaders([]);
    setSubColumns(null);
    setFormData({});
    setShowForm(false);
    setCurrentRecordIndex(0);
    setFormHeader(null);
    setHeaderFormData({});
    setImportFile(null);
    setImportStatus('');
    setShowSavedForms(false);
    setIsFormFromSaved(false);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  // Load a saved form
  const handleLoadSavedForm = (savedForm) => {
    setFormHeader(savedForm.formHeader);
    setHeaderFormData(savedForm.headerFormData || {});
    setFormFields(savedForm.formFields || []);
    setImportedData(savedForm.importedData || []);
    setCurrentRecordIndex(savedForm.currentRecordIndex || 0);
    setFormData(savedForm.formData || {});
    setShowForm(true);
    setShowSavedForms(false);
    setIsFormFromSaved(true); // Mark that form was loaded from saved forms
    setImportStatus(`Loaded saved form: ${savedForm.formName}`);
    setTimeout(() => setImportStatus(''), 3000);
  };

  // Delete a saved form
  const handleDeleteSavedForm = async (formId) => {
    try {
      // Find the form being deleted
      const formToDelete = savedForms.find(f => f.id === formId);
     
      if (!formToDelete) {
        setImportStatus('Form not found');
        setTimeout(() => setImportStatus(''), 2000);
        return;
      }

      const formName = formToDelete.formName;

      // Remove form file from checklist entry (keep the entry, just remove the file)
      // If duplicate exists with no file, delete the entry we're updating to avoid duplicates
      try {
        const checklistResp = await fetch('/server/checklist_function/checklist');
        if (checklistResp.ok) {
          const checklistData = await checklistResp.json();
          if (checklistData.status === 'success' && checklistData.data && checklistData.data.checklistData) {
            // Find ALL matching checklist entries by form name
            const matchingChecklists = checklistData.data.checklistData.filter(
              item => item.formName && item.formName.toLowerCase().trim() === formName.toLowerCase().trim()
            );

            if (matchingChecklists.length > 0) {
              // Find the entry that has a form file (the one we want to update)
              const entryWithFile = matchingChecklists.find(item =>
                item.formFile &&
                item.formFile !== null &&
                item.formFile !== 'null' &&
                String(item.formFile).trim() !== ''
              );

              // Find entries without files
              const entriesWithoutFile = matchingChecklists.filter(item =>
                !item.formFile ||
                item.formFile === null ||
                item.formFile === 'null' ||
                String(item.formFile).trim() === ''
              );

              if (entryWithFile && entryWithFile.id) {
                // If there's already an entry without a file, delete the one we're updating to avoid duplicates
                if (entriesWithoutFile.length > 0) {
                  // First, delete proof submission file if it exists before deleting the entry
                  if (entryWithFile.proofSubmissionFile &&
                      entryWithFile.proofSubmissionFile !== null &&
                      entryWithFile.proofSubmissionFile !== 'null' &&
                      String(entryWithFile.proofSubmissionFile).trim() !== '') {
                    try {
                      const deleteProofResp = await fetch(`/server/checklist_function/checklist/${entryWithFile.id}/file/ProofSubmission`, {
                        method: 'DELETE'
                      });
                     
                      if (deleteProofResp.ok) {
                        const deleteProofData = await deleteProofResp.json();
                        if (deleteProofData.status === 'success') {
                          console.log('Successfully deleted proof submission file before deleting checklist entry:', entryWithFile.id);
                        }
                      }
                    } catch (proofDeleteError) {
                      console.error('Error deleting proof submission file:', proofDeleteError);
                      // Continue with entry deletion even if proof submission deletion fails
                    }
                  }
                 
                  // Delete the entry with file since there's already one without file
                  const deleteResp = await fetch(`/server/checklist_function/checklist/${entryWithFile.id}`, {
                method: 'DELETE'
              });

              if (deleteResp.ok) {
                const deleteData = await deleteResp.json();
                if (deleteData.status === 'success') {
                      console.log('Deleted checklist entry with file to avoid duplicate:', entryWithFile.id);
                    }
                  }
                } else {
                  // No duplicate without file, just remove the file from this entry
                 
                  // First, delete proof submission file if it exists
                  if (entryWithFile.proofSubmissionFile &&
                      entryWithFile.proofSubmissionFile !== null &&
                      entryWithFile.proofSubmissionFile !== 'null' &&
                      String(entryWithFile.proofSubmissionFile).trim() !== '') {
                    try {
                      const deleteProofResp = await fetch(`/server/checklist_function/checklist/${entryWithFile.id}/file/ProofSubmission`, {
                        method: 'DELETE'
                      });
                     
                      if (deleteProofResp.ok) {
                        const deleteProofData = await deleteProofResp.json();
                        if (deleteProofData.status === 'success') {
                          console.log('Successfully deleted proof submission file from checklist entry:', entryWithFile.id);
                        }
                      }
                    } catch (proofDeleteError) {
                      console.error('Error deleting proof submission file:', proofDeleteError);
                      // Continue with form file removal even if proof submission deletion fails
                    }
                  }
                 
                  // Remove form file from this entry
                  const updateResp = await fetch(`/server/checklist_function/checklist/${entryWithFile.id}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                      formName: entryWithFile.formName, // Keep existing form name
                      act: entryWithFile.act || null,
                      consultgovtdep: entryWithFile.consultgovtdep || null,
                      dueDate: entryWithFile.dueDate || null,
                      description: entryWithFile.description || null,
                      sector: entryWithFile.sector || null,
                      state: entryWithFile.state || null,
                      formFile: null, // Remove form file
                      formFileName: null, // Remove form file name
                      proofSubmissionFile: null, // Remove proof submission file
                      proofSubmissionFileName: null, // Remove proof submission file name
                      complianceStatus: entryWithFile.complianceStatus || null,
                      approvalStatus: entryWithFile.approvalStatus || null,
                      marks: entryWithFile.marks || null
                    })
                  });

                  if (updateResp.ok) {
                    const updateData = await updateResp.json();
                    if (updateData.status === 'success') {
                      console.log('Successfully removed form file and proof submission from checklist entry:', entryWithFile.id);
                    }
                  }
                }
              }
            }
          }
        }
      } catch (checklistError) {
        console.error('Error removing form file from checklist:', checklistError);
        // Continue with SEMaster deletion even if checklist update fails
      }

      // Delete from SEMaster saved forms
      const updated = savedForms.filter(f => f.id !== formId);
      setSavedForms(updated);
      localStorage.setItem('semaster_saved_forms', JSON.stringify(updated));
      setImportStatus('Saved form deleted, form file and proof submission removed from checklist');
      setTimeout(() => setImportStatus(''), 3000);
    } catch (error) {
      console.error('Error deleting saved form:', error);
      setImportStatus('Error deleting form: ' + error.message);
      setTimeout(() => setImportStatus(''), 3000);
    }
  };

  // Convert current form data to Excel file
  const convertFormDataToExcel = () => {
    try {
      // Create workbook
      const wb = XLSX.utils.book_new();
     
      // Prepare data array with header rows
      const dataArray = [];
     
      // Add form header information if available
      if (formHeader) {
        if (formHeader.title) {
          dataArray.push([formHeader.title]);
        }
        if (formHeader.subtitle) {
          dataArray.push([formHeader.subtitle]);
        }
        if (formHeader.reference) {
          dataArray.push([formHeader.reference]);
        }
        if (formHeader.textRows && formHeader.textRows.length > 0) {
          formHeader.textRows.forEach(text => {
            dataArray.push([text]);
          });
        }
       
        // Add header fields
        if (formHeader.fields && formHeader.fields.length > 0) {
          formHeader.fields.forEach(field => {
            const value = headerFormData[field.key] || field.value || '';
            dataArray.push([field.label, value]);
          });
        }
       
        // Add empty row before table
        dataArray.push([]);
      }
     
      // Add table headers
      const headers = formFields.map(field => field.label);
      dataArray.push(headers);
     
      // Add current form data row
      const rowData = formFields.map(field => formData[field.name] || '');
      dataArray.push(rowData);
     
      // Create worksheet
      const ws = XLSX.utils.aoa_to_sheet(dataArray);
     
      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
     
      // Generate Excel file as blob
      const excelBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
     
      // Create file name
      const formName = formHeader?.title || formHeader?.subtitle || 'SEMaster_Form';
      const fileName = `${formName.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}.xlsx`;
     
      return new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    } catch (error) {
      console.error('Error converting form data to Excel:', error);
      throw error;
    }
  };

  // Handle click on "Save to Checklist" button - show form name modal
  const handleSaveToChecklistClick = () => {
    // Set default form name from header
    const defaultFormName = formHeader?.title || formHeader?.subtitle || 'S&E Master Form';
    setFormNameInput(defaultFormName);
    setShowFormNameModal(true);
  };

  // Save form to checklist
  const handleSaveToChecklist = async (userFormName) => {
    if (isSavingToChecklist) return;
   
    setIsSavingToChecklist(true);
    setShowFormNameModal(false);
    setImportStatus('Saving to checklist...');
   
    try {
      // Use user-provided form name or fallback to header/default
      const formName = userFormName?.trim() || formHeader?.title || formHeader?.subtitle || 'S&E Master Form';
     
      // Convert form data to Excel file
      let formFile = null;
      let formFileName = null;
     
      try {
        formFile = convertFormDataToExcel();
        formFileName = formFile.name;
      } catch (excelError) {
        console.error('Error creating Excel file:', excelError);
        // Continue without file if Excel conversion fails
      }
     
      // Upload form file if available
      let formFileId = null;
      if (formFile) {
        try {
          const formDataUpload = new FormData();
          formDataUpload.append('file', formFile);
         
          const uploadResp = await fetch('/server/checklist_function/checklist/upload/Form', {
            method: 'POST',
            body: formDataUpload
          });
         
          if (uploadResp.ok) {
            const uploadData = await uploadResp.json();
            if (uploadData.status === 'success') {
              formFileId = uploadData.fileId;
            }
          }
        } catch (uploadError) {
          console.error('Error uploading form file:', uploadError);
          // Continue without file if upload fails
        }
      }
     
      // Helper function to extract field value
      const getFieldValue = (searchKeywords) => {
        if (!formHeader?.fields) return null;
        const field = formHeader.fields.find(f => {
          const labelLower = f.label.toLowerCase();
          return searchKeywords.some(keyword => labelLower.includes(keyword));
        });
        if (field) {
          return headerFormData[field.key] || field.value || null;
        }
        return null;
      };
     
      // Check if a checklist entry with the same form name already exists
      // First try exact match (case-insensitive), then try normalized match
      // Use the user-entered form name exactly as they typed it
      let existingChecklist = null;
      let checklistData = null;
     
      // Normalize form name for comparison (remove extra spaces, dashes, case-insensitive)
      // This will match "FORM - U" with "FORM U" and "Form U"
      const normalizeFormName = (name) => {
        if (!name) return '';
        return name.toLowerCase()
          .replace(/\s*-\s*/g, ' ') // Replace " - " or "-" with single space first
          .replace(/\s+/g, ' ') // Replace multiple spaces with single space
          .trim();
      };
     
      const userFormNameTrimmed = formName.trim();
      const normalizedFormName = normalizeFormName(userFormNameTrimmed);
     
      try {
        const checklistResp = await fetch('/server/checklist_function/checklist');
        if (checklistResp.ok) {
          checklistData = await checklistResp.json();
          if (checklistData.status === 'success' && checklistData.data && checklistData.data.checklistData) {
           
            // Helper function to check if entry has no file
            const hasNoFile = (item) => {
              const noFile = !item.formFile ||
                  item.formFile === null ||
                  item.formFile === 'null' ||
                  String(item.formFile).trim() === '' ||
                  item.formFileName === null ||
                  item.formFileName === 'No File' ||
                  !item.formFileName ||
                  String(item.formFileName).trim() === '';
              return noFile;
            };
             
            // Collect all possible matches (exact, normalized, and partial)
            let allMatches = [];
           
            // First, try exact match (case-insensitive)
            const exactMatches = checklistData.data.checklistData.filter(item => {
              if (!item.formName) return false;
              const itemFormName = item.formName.trim().toLowerCase();
              const searchFormName = userFormNameTrimmed.toLowerCase();
              return itemFormName === searchFormName;
            });
           
            if (exactMatches.length > 0) {
              allMatches = exactMatches;
              console.log('Found exact matches:', exactMatches.length);
            } else {
              // If no exact match, try normalized match
              const normalizedMatches = checklistData.data.checklistData.filter(item => {
                if (!item.formName) return false;
                const normalizedItemName = normalizeFormName(item.formName);
                return normalizedItemName === normalizedFormName;
              });
             
              if (normalizedMatches.length > 0) {
                allMatches = normalizedMatches;
                console.log('Found normalized matches:', normalizedMatches.length);
              } else {
                // Last attempt: try partial match (e.g., "Form U" matches "FORM - U")
                const partialMatches = checklistData.data.checklistData.filter(item => {
                  if (!item.formName) return false;
                  const itemFormName = item.formName.trim().toLowerCase();
                  const searchFormName = userFormNameTrimmed.toLowerCase();
                 
                  // Extract just the form identifier (e.g., "u" from "form u" or "form - u")
                  const extractFormId = (name) => {
                    return name.replace(/^form\s*[-]?\s*/i, '').trim();
                  };
                 
                  const itemId = extractFormId(itemFormName);
                  const searchId = extractFormId(searchFormName);
                 
                  return itemId && searchId && itemId === searchId;
                });
               
                if (partialMatches.length > 0) {
                  allMatches = partialMatches;
                  console.log('Found partial matches:', partialMatches.length);
                }
              }
            }
           
            // If we found any matches, select the best one to update
            if (allMatches.length > 0) {
              // Prioritize entries with "No File" - these should be updated first
              const noFileEntries = allMatches.filter(item => hasNoFile(item));
             
              if (noFileEntries.length > 0) {
                // Use the first "No File" entry
                existingChecklist = noFileEntries[0];
                console.log('Selected "No File" entry to update:', {
                  id: existingChecklist.id,
                  formName: existingChecklist.formName,
                  formFileName: existingChecklist.formFileName
                });
              } else {
                // If no "No File" entries, use the first match anyway (to avoid duplicates)
                existingChecklist = allMatches[0];
                console.log('No "No File" entry found, using first match to update:', {
                  id: existingChecklist.id,
                  formName: existingChecklist.formName,
                  formFileName: existingChecklist.formFileName,
                  totalMatches: allMatches.length
                });
              }
             
              console.log('Match summary:', {
                searchingFor: userFormNameTrimmed,
                found: existingChecklist.formName,
                id: existingChecklist.id,
                hasNoFile: hasNoFile(existingChecklist),
                formFileName: existingChecklist.formFileName,
                totalMatches: allMatches.length,
                noFileMatches: noFileEntries.length
              });
            } else {
              console.log('No matching checklist entry found:', {
                searchingFor: userFormNameTrimmed,
                normalizedSearch: normalizedFormName,
                availableFormNames: checklistData.data.checklistData.slice(0, 20).map(item => ({
                  original: item.formName,
                  normalized: normalizeFormName(item.formName),
                  hasNoFile: hasNoFile(item),
                  formFileName: item.formFileName
                }))
              });
            }
          }
        }
      } catch (checklistError) {
        console.error('Error checking existing checklist:', checklistError);
        // Continue with creating new entry if check fails
      }
     
      // If no existing checklist found, try to fetch default values from ChecklistBulk (master data)
      let defaultValues = null;
      if (!existingChecklist) {
        try {
          const bulkResp = await fetch('/server/checklistbulk_function/checklistbulk?action=getAll');
          if (bulkResp.ok) {
            const bulkData = await bulkResp.json();
            if (bulkData.status === 'success' && bulkData.data && Array.isArray(bulkData.data)) {
              // Find matching form name in bulk data (case-insensitive)
              const bulkMatch = bulkData.data.find(item => {
                if (!item.formName) return false;
                return item.formName.trim().toLowerCase() === userFormNameTrimmed.toLowerCase();
              });
             
              if (bulkMatch) {
                defaultValues = {
                  act: bulkMatch.act || null,
                  dueDate: bulkMatch.dueDate || null,
                  description: bulkMatch.description || null,
                  sector: bulkMatch.sector || null,
                  state: bulkMatch.state || null,
                  consultgovtdep: bulkMatch.concernedGovtDepartment || null
                };
                console.log('Found default values from ChecklistBulk:', defaultValues);
              }
            }
          }
        } catch (bulkError) {
          console.error('Error fetching ChecklistBulk data:', bulkError);
        }
       
        // Also try to find similar form names in existing checklist data
        if (!defaultValues && checklistData && checklistData.data && checklistData.data.checklistData) {
          // Find any entry with similar form name (normalized match)
          const similarEntry = checklistData.data.checklistData.find(item => {
            if (!item.formName) return false;
            const normalizedItemName = normalizeFormName(item.formName);
            return normalizedItemName === normalizedFormName;
          });
         
          if (similarEntry) {
            defaultValues = {
              act: similarEntry.act || null,
              dueDate: similarEntry.dueDate || null,
              description: similarEntry.description || null,
              sector: similarEntry.sector || null,
              state: similarEntry.state || null,
              consultgovtdep: similarEntry.consultgovtdep || null
            };
            console.log('Found default values from similar checklist entry:', defaultValues);
          }
        }
      }
     
      // Prepare checklist payload
      // Always use the user-entered form name exactly as they typed it
      const checklistPayload = {
        formName: userFormNameTrimmed, // Always use the exact form name the user entered
        act: getFieldValue(['act']) || existingChecklist?.act || defaultValues?.act || null,
        consultgovtdep: getFieldValue(['department', 'govt', 'government', 'consult']) || existingChecklist?.consultgovtdep || defaultValues?.consultgovtdep || null,
        dueDate: existingChecklist?.dueDate || defaultValues?.dueDate || null,
        description: existingChecklist?.description || defaultValues?.description || `Form imported from S&E Master. Contains ${importedData.length} record(s). Currently viewing record ${currentRecordIndex + 1} of ${importedData.length}.`,
        sector: getFieldValue(['sector']) || existingChecklist?.sector || defaultValues?.sector || null,
        state: getFieldValue(['state']) || existingChecklist?.state || defaultValues?.state || null,
        // Always use the new file if available (from SEMaster), otherwise keep existing (for updates)
        // When updating a "No File" entry, we want to replace it with the new file
        formFile: formFileId ? formFileId : (existingChecklist?.formFile || null),
        formFileName: formFileName ? formFileName : (existingChecklist?.formFileName || null),
        proofSubmissionFile: existingChecklist?.proofSubmissionFile || null,
        proofSubmissionFileName: existingChecklist?.proofSubmissionFileName || null,
        complianceStatus: existingChecklist?.complianceStatus || null,
        approvalStatus: existingChecklist?.approvalStatus || null,
        marks: existingChecklist?.marks || null
      };
     
      // Log for debugging
      console.log('Saving to checklist:', {
        userEnteredFormName: userFormNameTrimmed,
        existingChecklistFormName: existingChecklist?.formName,
        finalFormName: checklistPayload.formName,
        formFile: checklistPayload.formFile,
        formFileName: checklistPayload.formFileName,
        isUpdate: !!existingChecklist,
        existingChecklistId: existingChecklist?.id
      });
     
      // Save to checklist - update if exists, create if new
      let resp, data;
      if (existingChecklist && existingChecklist.id) {
        // Update existing entry - replace "No File" with the new file
        console.log('Updating existing checklist entry:', {
          id: existingChecklist.id,
          currentFormName: existingChecklist.formName,
          currentFormFile: existingChecklist.formFile,
          currentFormFileName: existingChecklist.formFileName,
          newFormFile: formFileId,
          newFormFileName: formFileName,
          willUpdate: true
        });
       
        resp = await fetch(`/server/checklist_function/checklist/${existingChecklist.id}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(checklistPayload)
        });
        data = await resp.json();
       
        if (!resp.ok || data.status !== 'success') {
          console.error('Failed to update checklist:', data);
          throw new Error(data.message || 'Failed to update checklist entry');
        }
       
        console.log('Successfully updated checklist entry:', {
          id: existingChecklist.id,
          formName: checklistPayload.formName,
          formFileName: checklistPayload.formFileName
        });
      } else {
        // Triple-check before creating: search with all possible matching strategies
        if (checklistData && checklistData.data && checklistData.data.checklistData) {
            // Try exact match first
          let finalCheck = checklistData.data.checklistData.find(item => {
            if (!item.formName) return false;
            const itemFormName = item.formName.trim().toLowerCase();
            const searchFormName = userFormNameTrimmed.toLowerCase();
            return itemFormName === searchFormName;
          });
         
          // If no exact match, try normalized match
          if (!finalCheck) {
            finalCheck = checklistData.data.checklistData.find(item => {
              if (!item.formName) return false;
              const normalizedItemName = normalizeFormName(item.formName);
              return normalizedItemName === normalizedFormName;
            });
          }
         
          // If still no match, try partial match (extract form ID)
          if (!finalCheck) {
            const extractFormId = (name) => {
              return name.toLowerCase().replace(/^form\s*[-]?\s*/i, '').trim();
            };
            const searchId = extractFormId(userFormNameTrimmed);
           
            finalCheck = checklistData.data.checklistData.find(item => {
              if (!item.formName) return false;
              const itemId = extractFormId(item.formName);
              return itemId && searchId && itemId === searchId;
          });
          }
         
          if (finalCheck && finalCheck.id) {
            console.log('Found match on final check, updating instead of creating:', {
              id: finalCheck.id,
              formName: finalCheck.formName,
              searchFor: userFormNameTrimmed
            });
            existingChecklist = finalCheck;
           
            // Rebuild payload with the found checklist's data
            const finalPayload = {
              ...checklistPayload,
              act: checklistPayload.act || finalCheck.act || null,
              consultgovtdep: checklistPayload.consultgovtdep || finalCheck.consultgovtdep || null,
              dueDate: checklistPayload.dueDate || finalCheck.dueDate || null,
              description: checklistPayload.description || finalCheck.description || null,
              sector: checklistPayload.sector || finalCheck.sector || null,
              state: checklistPayload.state || finalCheck.state || null
            };
           
            resp = await fetch(`/server/checklist_function/checklist/${existingChecklist.id}`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(finalPayload)
            });
            data = await resp.json();
           
            if (!resp.ok || data.status !== 'success') {
              console.error('Failed to update checklist on final check:', data);
              throw new Error(data.message || 'Failed to update checklist entry');
            }
           
            console.log('Successfully updated checklist entry on final check:', existingChecklist.id);
          } else {
            // Only create new entry if NO matching entry found after all checks
            console.log('No existing checklist entry found after all checks, creating new entry');
            console.log('Search was for:', userFormNameTrimmed, 'normalized:', normalizedFormName);
           
            // Create new entry with the user-entered form name
            console.log('Creating new checklist entry with form name:', userFormNameTrimmed);
            resp = await fetch('/server/checklist_function/checklist', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(checklistPayload)
            });
            data = await resp.json();
           
            if (!resp.ok || data.status !== 'success') {
              throw new Error(data.message || 'Failed to create checklist entry');
            }
          }
        } else {
          // Fallback: create new if we can't check
          console.log('Cannot verify duplicates, creating new entry');
          resp = await fetch('/server/checklist_function/checklist', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(checklistPayload)
          });
          data = await resp.json();
         
          if (!resp.ok || data.status !== 'success') {
            throw new Error(data.message || 'Failed to create checklist entry');
          }
        }
      }
     
      // Sync form file to Statutory (Data Store Statutory table) so it displays in Statutory page, not only in Checklist
      if (formFile) {
        try {
          const statutoryFormData = new FormData();
          statutoryFormData.append('file', formFile);
          const statutoryUploadResp = await fetch('/server/statutoryreg_function/statutory/upload/Form', {
            method: 'POST',
            body: statutoryFormData
          });
          let statutoryUploadData = null;
          if (statutoryUploadResp.headers.get('content-type')?.includes('application/json')) {
            statutoryUploadData = await statutoryUploadResp.json();
          }
          if (statutoryUploadResp.ok && statutoryUploadData?.status === 'success' && statutoryUploadData?.fileId) {
            const statutoryFileId = statutoryUploadData.fileId;
            const statutoryFileName = statutoryUploadData.fileName || formFileName;
            const statutoryResp = await fetch('/server/statutoryreg_function/statutory');
            if (statutoryResp.ok) {
              const statutoryResult = await statutoryResp.json();
              const statutoryList = statutoryResult?.data?.statutoryData || [];
              const extractFormIdForStat = (name) => {
                return String(name || '').toLowerCase().replace(/\s*-\s*/g, ' ').replace(/\s+/g, ' ').trim().replace(/^form\s*/i, '').trim();
              };
              const searchFormId = extractFormIdForStat(userFormNameTrimmed);
              const allMatches = statutoryList.filter(item => {
                const itemFormId = extractFormIdForStat(item.formName);
                return itemFormId && searchFormId && itemFormId === searchFormId;
              });
              // Update form file on existing records only (all months). Do not create new Statutory record.
              if (allMatches.length > 0) {
                for (const match of allMatches) {
                  const updatePayload = {
                    formName: match.formName || userFormNameTrimmed,
                    formFile: statutoryFileId,
                    formFileName: statutoryFileName,
                    act: match.act ?? null,
                    description: match.description ?? null,
                    dueDate: match.dueDate ?? null,
                    sector: match.sector ?? null,
                    state: match.state ?? null,
                    monthfilter: match.monthFilter ?? match.MonthFilter ?? match.monthfilter ?? null,
                    autofill: match.autofill ?? null,
                    draft: match.draft ?? null,
                    proofSubmissionFile: match.proofSubmissionFile ?? null,
                    proofSubmissionFileName: match.proofSubmissionFileName ?? null,
                    draftFile: match.draftFile ?? null,
                    draftFileName: match.draftFileName ?? null
                  };
                  const updateResp = await fetch(`/server/statutoryreg_function/statutory/${match.id}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(updatePayload)
                  });
                  if (updateResp.ok) console.log('Updated Statutory record with form file (all months):', match.id);
                }
              }
              // If no existing Statutory record matches this form name, do not create one; only Checklist gets the form file.
            }
          }
        } catch (statutoryErr) {
          console.warn('Could not sync form file to Statutory (file still saved to Checklist):', statutoryErr);
        }
      }
     
      // Also save the form to SEMaster's saved forms list
      // Check if form with same name already exists and update it, otherwise add new
      const existingFormIndex = savedForms.findIndex(
        f => f.formName && f.formName.toLowerCase().trim() === userFormNameTrimmed.toLowerCase()
      );

      const savedFormEntry = {
        id: existingFormIndex >= 0 ? savedForms[existingFormIndex].id : Date.now().toString(),
        formName: userFormNameTrimmed,
        formHeader: formHeader,
        headerFormData: { ...headerFormData },
        formFields: formFields,
        importedData: importedData,
        currentRecordIndex: currentRecordIndex,
        formData: { ...formData },
        savedAt: new Date().toISOString(),
        description: checklistPayload.description
      };

      let updatedSavedForms;
      if (existingFormIndex >= 0) {
        // Update existing form
        updatedSavedForms = [...savedForms];
        updatedSavedForms[existingFormIndex] = savedFormEntry;
      } else {
        // Add new form
        updatedSavedForms = [...savedForms, savedFormEntry];
      }

      setSavedForms(updatedSavedForms);
      localStorage.setItem('semaster_saved_forms', JSON.stringify(updatedSavedForms));

      setImportStatus('Successfully saved to Checklist, Statutory, and SEMaster!');
      setTimeout(() => {
        setImportStatus('');
        // Optionally navigate to checklist page
        // navigate('/rule-book/checklist');
      }, 3000);
     
    } catch (error) {
      console.error('Error saving to checklist:', error);
      setImportStatus(`Error: ${error.message || 'Failed to save to checklist'}`);
      setTimeout(() => setImportStatus(''), 5000);
    } finally {
      setIsSavingToChecklist(false);
    }
  };

  return (
    <div className="semaster-container">
      {/* Header */}
      <div className="semaster-header">
        <div className="header-content">
          <div className="header-icon">
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="#3b82f6" strokeWidth="2"/>
              <polyline points="14,2 14,8 20,8" stroke="#3b82f6" strokeWidth="2"/>
            </svg>
          </div>
          <div className="header-text">
            <h1 className="semaster-title">S&E Master</h1>
            <p className="semaster-subtitle">Import Excel and view data in form format</p>
          </div>
        </div>
      </div>

      <div className="semaster-content">
        {!showForm && !showSavedForms ? (
          /* Import Section */
          <div className="import-section-wrapper">
            {/* Show saved forms count if forms exist */}
            {savedForms.length > 0 && (
              <div className="saved-forms-banner">
                <div className="banner-content">
                  <div className="banner-info">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14,2 14,8 20,8"/>
                    </svg>
                    <span>You have {savedForms.length} saved form(s)</span>
                  </div>
                  <button
                    className="btn-view-saved"
                    onClick={() => setShowSavedForms(true)}
                  >
                    View Saved Forms
                  </button>
                </div>
              </div>
            )}
            <div className="import-grid">
            <div className="upload-card">
              <div className="card-header">
                <div className="card-icon">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="#10b981" strokeWidth="2"/>
                    <polyline points="7,10 12,15 17,10" stroke="#10b981" strokeWidth="2"/>
                    <line x1="12" y1="15" x2="12" y2="3" stroke="#10b981" strokeWidth="2"/>
                  </svg>
                </div>
                <h2 className="card-title">Import Excel File</h2>
              </div>
             
              <div className="card-content">
                <div className="file-upload-zone">
                  <input
                    ref={fileInputRef}
                    type="file"
                    id="excel-upload"
                    accept=".xlsx,.xls"
                    onChange={(e) => handleFileUpload(e.target.files[0])}
                    className="file-input"
                  />
                  <label htmlFor="excel-upload" className="upload-zone">
                    {importFile ? (
                      <div className="file-selected-state">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="#10b981" strokeWidth="2"/>
                            <polyline points="14,2 14,8 20,8" stroke="#10b981" strokeWidth="2"/>
                          </svg>
                        </div>
                        <div className="file-info">
                          <span className="file-name">{importFile.name}</span>
                          <span className="file-size">{(importFile.size / 1024).toFixed(1)} KB</span>
                        </div>
                        <div className="file-status">
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                            <path d="M20 6L9 17l-5-5" stroke="#10b981" strokeWidth="2"/>
                          </svg>
                        </div>
                      </div>
                    ) : (
                      <div className="upload-placeholder">
                        <div className="upload-icon">
                          <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="#6b7280" strokeWidth="1.5"/>
                            <polyline points="7,10 12,15 17,10" stroke="#6b7280" strokeWidth="1.5"/>
                            <line x1="12" y1="15" x2="12" y2="3" stroke="#6b7280" strokeWidth="1.5"/>
                          </svg>
                        </div>
                        <div className="upload-text">
                          <h3>Drop your Excel file here</h3>
                          <p>or click to browse files</p>
                          <span className="file-types">Supports .xlsx, .xls files</span>
                        </div>
                      </div>
                    )}
                  </label>
                </div>

                <button
                  className={`import-btn ${importFile ? 'btn-primary' : 'btn-disabled'}`}
                  onClick={handleImport}
                  disabled={!importFile || isImporting}
                >
                  {isImporting ? (
                    <>
                      <div className="spinner-small"></div>
                      Processing...
                    </>
                  ) : (
                    <>
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2"/>
                        <polyline points="7,10 12,15 17,10" stroke="currentColor" strokeWidth="2"/>
                        <line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="2"/>
                      </svg>
                      Import Data
                    </>
                  )}
                </button>
              </div>

              {importStatus && (
                <div className={`status-message ${importStatus.includes('Successfully') ? 'success' : 'error'}`}>
                  <div className="status-icon">
                    {importStatus.includes('Successfully') ? (
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                        <path d="M20 6L9 17l-5-5" stroke="currentColor" strokeWidth="2"/>
                      </svg>
                    ) : (
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                        <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2"/>
                      </svg>
                    )}
                  </div>
                  <span>{importStatus}</span>
                </div>
              )}
            </div>

            <div className="info-card">
              <div className="card-header">
                <div className="card-icon">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none">
                    <circle cx="12" cy="12" r="10" stroke="#3b82f6" strokeWidth="2"/>
                    <path d="M12 16v-4M12 8h.01" stroke="#3b82f6" strokeWidth="2"/>
                  </svg>
                </div>
                <h2 className="card-title">Import Guidelines</h2>
              </div>
             
              <div className="card-content">
                <div className="info-list">
                  <div className="info-item">
                    <div className="info-icon">
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="#10b981" strokeWidth="2"/>
                      </svg>
                    </div>
                    <div className="info-text">
                      <strong>File Format:</strong> Excel (.xlsx, .xls)
                    </div>
                  </div>
                  <div className="info-item">
                    <div className="info-icon">
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                        <path d="M9 12l2 2 4-4" stroke="#10b981" strokeWidth="2"/>
                      </svg>
                    </div>
                    <div className="info-text">
                      <strong>First Row:</strong> Can be headers or data
                    </div>
                  </div>
                  <div className="info-item">
                    <div className="info-icon">
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                        <path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z" stroke="#10b981" strokeWidth="2"/>
                      </svg>
                    </div>
                    <div className="info-text">
                      <strong>Display:</strong> Data will be shown in form format
                    </div>
                  </div>
                </div>
              </div>
            </div>
            </div>
          </div>
        ) : showSavedForms ? (
          /* Saved Forms Section */
          <div className="saved-forms-section">
            <div className="form-card">
              <div className="form-header">
                <div className="form-title-section">
                  <h2>Saved Forms</h2>
                  <p>{savedForms.length} form(s) saved</p>
                </div>
                <div className="form-actions">
                  <button
                    className="btn-nav"
                    onClick={() => {
                      setShowSavedForms(false);
                      setShowForm(false);
                      // Reset form state to ensure clean import
                      setImportFile(null);
                      setImportedData([]);
                      setFormFields([]);
                      setFormData({});
                      setFormHeader(null);
                      setHeaderFormData({});
                      setCurrentRecordIndex(0);
                      if (fileInputRef.current) {
                        fileInputRef.current.value = '';
                      }
                    }}
                    title="Go to Import Section"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                      <polyline points="7,10 12,15 17,10"/>
                      <line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                    Import New File
                  </button>
                </div>
              </div>

              {savedForms.length === 0 ? (
                <div className="empty-state">
                  <div className="empty-icon">
                    <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14,2 14,8 20,8"/>
                    </svg>
                  </div>
                  <h3>No saved forms yet</h3>
                  <p>Forms saved to checklist will appear here</p>
                </div>
              ) : (
                <div className="saved-forms-list">
                  {savedForms.map((savedForm) => (
                    <div key={savedForm.id} className="saved-form-item">
                      <div className="saved-form-info">
                        <div className="saved-form-header">
                          <h3>{savedForm.formName}</h3>
                          <span className="saved-form-date">
                            {new Date(savedForm.savedAt).toLocaleDateString()} {new Date(savedForm.savedAt).toLocaleTimeString()}
                          </span>
                        </div>
                        <p className="saved-form-description">{savedForm.description}</p>
                        <div className="saved-form-meta">
                          <span>{savedForm.importedData?.length || 0} record(s)</span>
                          {savedForm.formHeader?.title && (
                            <span className="form-badge">{savedForm.formHeader.title}</span>
                          )}
                        </div>
                      </div>
                      <div className="saved-form-actions">
                        <button
                          className="btn-load"
                          onClick={() => handleLoadSavedForm(savedForm)}
                          title="Load this form"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7,10 12,15 17,10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                          </svg>
                          Load
                        </button>
                        <button
                          className="btn-delete"
                          onClick={() => handleDeleteSavedForm(savedForm.id)}
                          title="Delete this form"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <polyline points="3,6 5,6 21,6"/>
                            <path d="M19,6v14a2,2 0 0,1 -2,2H7a2,2 0 0,1 -2,-2V6m3,0V4a2,2 0 0,1 2,-2h4a2,2 0 0,1 2,2v2"/>
                          </svg>
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        ) : (
          /* Form Display Section */
          <div className="form-display-section">
            <div className="form-card">
              <div className="form-header">
                <div className="form-title-section">
                  <h2>Form View</h2>
                  <p>Record {currentRecordIndex + 1} of {importedData.length}</p>
                </div>
                <div className="form-actions">
                  {isFormFromSaved && (
                    <button
                      className="btn-nav"
                      onClick={() => {
                        setShowForm(false);
                        setShowSavedForms(true);
                        setIsFormFromSaved(false);
                      }}
                      title="Back to Saved Forms"
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <polyline points="15,18 9,12 15,6"/>
                      </svg>
                      Back
                    </button>
                  )}
                  <button
                    className="btn-nav"
                    onClick={() => setShowSavedForms(true)}
                    title="View Saved Forms"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14,2 14,8 20,8"/>
                    </svg>
                    Saved Forms ({savedForms.length})
                  </button>
                  <button
                    className="btn-reset"
                    onClick={handleReset}
                    title="Reset and Import New File"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <polyline points="1,4 1,10 7,10"/>
                      <path d="m3.51,15a9,9 0 1,0 2.13,-9.36L1,10"/>
                    </svg>
                    Reset
                  </button>
                </div>
              </div>

              {importStatus && (
                <div className={`status-message ${importStatus.includes('saved') ? 'success' : 'info'}`}>
                  <span>{importStatus}</span>
                </div>
              )}

              {/* Form Header Section */}
              {formHeader && (
                <div className="form-header-section">
                  {/* Line 1: FORM - I */}
                  {formHeader.title && (
                    <div className="form-header-title">{formHeader.title}</div>
                  )}
                  {/* Line 2: REGISTER OF WORKMEN */}
                  {formHeader.subtitle && (
                    <div className="form-header-subtitle">{formHeader.subtitle}</div>
                  )}
                  {/* Line 3: [See sub-rule (1) under rule 6] */}
                  {formHeader.reference && (
                    <div className="form-header-reference">{formHeader.reference}</div>
                  )}
                  {/* Line 4: THE TAMIL NADU INDUSTRIAL ESTABLISHMENTS (CONFERMENT OF PERMANENT STATUS TO WORKMEN) ACT & RULES, 1981 */}
                  {formHeader.actLine && (
                    <div className="form-header-act-line">{formHeader.actLine}</div>
                  )}
                 
                  {/* Additional text rows from header */}
                  {formHeader.textRows && formHeader.textRows.length > 0 && (
                    <div className="form-header-text-rows">
                      {formHeader.textRows.map((text, index) => (
                        <div key={index} className="form-header-text-row">{text}</div>
                      ))}
                    </div>
                  )}
                 
                  {/* Header Fields (Name and Address, etc.) */}
                  {formHeader.fields && formHeader.fields.length > 0 && (
                    <div className="form-header-fields">
                      {formHeader.fields.map((field, index) => (
                        <div key={index} className="form-header-field">
                          <label className="form-header-field-label">
                            {field.label}
                          </label>
                          {isApprovedFestivalHolidaysField(field.label) ? (
                            <div className="festival-holidays-group" role="group" aria-label="Approved Festival Holidays">
                              <div className="festival-holidays-labels" aria-hidden="true">
                                {[1, 2, 3, 4, 5].map((n) => (
                                  <span key={n} className="festival-holiday-label">{n}</span>
                                ))}
                              </div>
                              <div className="festival-holidays-inputs">
                                {parseFiveValues(headerFormData[field.key] || '').map((v, i) => (
                                  <input
                                    key={i}
                                    type="text"
                                    inputMode="numeric"
                                    className="festival-holiday-box"
                                    value={v}
                                    onChange={(e) => handleFestivalHolidayBoxChange(field.key, i, e.target.value)}
                                    placeholder=""
                                    aria-label={`Approved Festival Holiday ${i + 1}`}
                                  />
                                ))}
                              </div>
                            </div>
                          ) : (
                            <input
                              type="text"
                              className="form-header-field-input"
                              value={headerFormData[field.key] || ''}
                              onChange={(e) => handleHeaderFieldChange(field.key, e.target.value)}
                              placeholder={`Enter ${field.label.replace(':', '')}`}
                            />
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              <div className="form-content">
                {importedData.length > 0 && (
                  <div style={{
                    marginBottom: '16px',
                    display: 'flex',
                    justifyContent: 'flex-end',
                    gap: '8px'
                  }}>
                    <button
                      className="btn-add-row"
                      onClick={handleAddRow}
                      title="Add a new row with all columns from imported Excel"
                      style={{
                        padding: '10px 20px',
                        backgroundColor: '#10b981',
                        color: 'white',
                        border: 'none',
                        borderRadius: '6px',
                        fontSize: '14px',
                        fontWeight: '500',
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        transition: 'all 0.2s ease'
                      }}
                      onMouseOver={(e) => {
                        e.currentTarget.style.backgroundColor = '#059669';
                      }}
                      onMouseOut={(e) => {
                        e.currentTarget.style.backgroundColor = '#10b981';
                      }}
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <line x1="12" y1="5" x2="12" y2="19"/>
                        <line x1="5" y1="12" x2="19" y2="12"/>
                      </svg>
                      Add Row
                    </button>
                  </div>
                )}
                <div className="form-table-wrapper">
                  <table className="semaster-form-table">
                    <thead>
                      {/* Main column headers row */}
                      {subColumns && Object.keys(subColumns).length > 0 ? (
                        <>
                          <tr>
                            {(() => {
                              const mainColumns = [];
                              let colSpan = 0;
                              tableHeaders.forEach((header, index) => {
                                const mainColumnSubCols = subColumns[header] || [];
                                if (mainColumnSubCols.length > 0) {
                                  mainColumns.push(
                                    <th
                                      key={index}
                                      className="form-table-header form-table-main-header"
                                      colSpan={mainColumnSubCols.length}
                                    >
                                      {header}
                                    </th>
                                  );
                                } else {
                                  mainColumns.push(
                                    <th
                                      key={index}
                                      className="form-table-header form-table-main-header"
                                      rowSpan={2}
                                    >
                                      {header}
                                    </th>
                                  );
                                }
                              });
                              return mainColumns;
                            })()}
                          </tr>
                          {/* Sub-column headers row */}
                          <tr>
                            {(() => {
                              const subColHeaders = [];
                              tableHeaders.forEach((header, index) => {
                                const mainColumnSubCols = subColumns[header] || [];
                                if (mainColumnSubCols.length > 0) {
                                  mainColumnSubCols.forEach((subCol, subIdx) => {
                                    subColHeaders.push(
                                      <th key={`${index}-${subIdx}`} className="form-table-header form-table-sub-header">
                                        {subCol}
                                      </th>
                                    );
                                  });
                                }
                                // Regular columns without sub-columns are already rendered with rowSpan=2
                              });
                              return subColHeaders;
                            })()}
                          </tr>
                        </>
                      ) : (
                        <tr>
                          {formFields.map((field, index) => (
                            <th key={index} className="form-table-header">
                              {field.label}
                            </th>
                          ))}
                        </tr>
                      )}
                    </thead>
                    <tbody>
                      {(() => {
                        // Calculate start index to show rows around currentRecordIndex
                        // Show 10 rows, centered around currentRecordIndex when possible
                        const totalRows = importedData.length;
                        const rowsToShow = 10;
                        let startIndex = 0;
                       
                        if (totalRows > rowsToShow) {
                          // Center the view around currentRecordIndex
                          startIndex = Math.max(0, Math.min(
                            currentRecordIndex - Math.floor(rowsToShow / 2),
                            totalRows - rowsToShow
                          ));
                        }
                       
                        const endIndex = Math.min(startIndex + rowsToShow, totalRows);
                        const rowsToDisplay = importedData.slice(startIndex, endIndex);
                       
                        return rowsToDisplay.map((record, displayIndex) => {
                          const actualRowIndex = startIndex + displayIndex;
                          return (
                            <tr
                              key={actualRowIndex}
                              className={actualRowIndex === currentRecordIndex ? 'active-row' : ''}
                              ref={actualRowIndex === currentRecordIndex ? (el) => {
                                if (el) {
                                  setTimeout(() => el.scrollIntoView({ behavior: 'smooth', block: 'nearest' }), 100);
                                }
                              } : null}
                            >
                              {subColumns && Object.keys(subColumns).length > 0 ? (
                                // Render with sub-columns
                                (() => {
                                  const cells = [];
                                  tableHeaders.forEach((header, headerIndex) => {
                                    const mainColumnSubCols = subColumns[header] || [];
                                    if (mainColumnSubCols.length > 0) {
                                      // Render sub-column cells
                                      mainColumnSubCols.forEach((subCol, subIdx) => {
                                        const subColName = `${header}_${subCol}`;
                                        cells.push(
                                          <td key={`${headerIndex}-${subIdx}`} className="form-table-cell">
                                            <input
                                              type="text"
                                              id={`${subColName}-${actualRowIndex}`}
                                              name={subColName}
                                              value={record[subColName] || ''}
                                              onChange={(e) => handleRowFieldChange(actualRowIndex, subColName, e.target.value)}
                                              placeholder={`Enter ${subCol}`}
                                              className="form-table-input"
                                              onClick={() => setCurrentRecordIndex(actualRowIndex)}
                                            />
                                          </td>
                                        );
                                      });
                                    } else {
                                      // Regular column without sub-columns
                                      cells.push(
                                        <td key={headerIndex} className="form-table-cell">
                                          <input
                                            type="text"
                                            id={`${header}-${actualRowIndex}`}
                                            name={header}
                                            value={record[header] || ''}
                                            onChange={(e) => handleRowFieldChange(actualRowIndex, header, e.target.value)}
                                            placeholder={`Enter ${header}`}
                                            className="form-table-input"
                                            onClick={() => setCurrentRecordIndex(actualRowIndex)}
                                          />
                                        </td>
                                      );
                                    }
                                  });
                                  return cells;
                                })()
                              ) : (
                                // Render without sub-columns (regular table)
                                formFields.map((field, fieldIndex) => (
                                  <td key={fieldIndex} className="form-table-cell">
                                    <input
                                      type={field.type}
                                      id={`${field.name}-${actualRowIndex}`}
                                      name={field.name}
                                      value={record[field.name] || ''}
                                      onChange={(e) => handleRowFieldChange(actualRowIndex, field.name, e.target.value)}
                                      placeholder={field.placeholder}
                                      className="form-table-input"
                                      onClick={() => setCurrentRecordIndex(actualRowIndex)}
                                    />
                                  </td>
                                ))
                              )}
                            </tr>
                          );
                        });
                      })()}
                      {importedData.length === 0 && (
                        <tr>
                          <td colSpan={formFields.length} style={{ textAlign: 'center', padding: '20px' }}>
                            No data available
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
                {importedData.length > 0 && (
                  <div style={{
                    padding: '10px 0',
                    textAlign: 'center',
                    color: '#6b7280',
                    fontSize: '14px'
                  }}>
                    {(() => {
                      const totalRows = importedData.length;
                      const rowsToShow = 10;
                      let startIndex = 0;
                     
                      if (totalRows > rowsToShow) {
                        startIndex = Math.max(0, Math.min(
                          currentRecordIndex - Math.floor(rowsToShow / 2),
                          totalRows - rowsToShow
                        ));
                      }
                     
                      const endIndex = Math.min(startIndex + rowsToShow, totalRows);
                      return `Showing rows ${startIndex + 1}-${endIndex} of ${totalRows} (Current: ${currentRecordIndex + 1})`;
                    })()}
                  </div>
                )}
              </div>

              <div className="form-footer">
                <div className="navigation-buttons">
                  <button
                    className="btn-nav"
                    onClick={handlePrevious}
                    disabled={currentRecordIndex === 0}
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <polyline points="15,18 9,12 15,6"/>
                    </svg>
                    Previous
                  </button>
                  <button
                    className="btn-save"
                    onClick={handleSave}
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                      <polyline points="17,21 17,13 7,13 7,21"/>
                      <polyline points="7,3 7,8 15,8"/>
                    </svg>
                    Save
                  </button>
                  <button
                    className="btn-nav"
                    onClick={handleNext}
                    disabled={currentRecordIndex === importedData.length - 1}
                  >
                    Next
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <polyline points="9,18 15,12 9,6"/>
                    </svg>
                  </button>
                </div>
                <div className="checklist-button-container">
                  <button
                    className="btn-checklist"
                    onClick={handleSaveToChecklistClick}
                    disabled={isSavingToChecklist}
                    title="Save current form to Checklist"
                  >
                    {isSavingToChecklist ? (
                      <>
                        <div className="spinner-small"></div>
                        Saving...
                      </>
                    ) : (
                      <>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M9 11l3 3L22 4"/>
                          <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
                        </svg>
                        Save to Checklist
                      </>
                    )}
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Form Name Modal */}
      {showFormNameModal && (
        <div className="form-name-modal-overlay" onClick={(e) => {
          if (e.target === e.currentTarget) {
            setShowFormNameModal(false);
            setFormNameInput('');
          }
        }}>
          <div className="form-name-modal">
            <div className="form-name-modal-header">
              <h3>Enter Form Name</h3>
              <button
                className="form-name-modal-close"
                onClick={() => {
                  setShowFormNameModal(false);
                  setFormNameInput('');
                }}
                title="Close"
              >
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <line x1="18" y1="6" x2="6" y2="18"/>
                  <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>
            <div className="form-name-modal-body">
              <label htmlFor="form-name-input" className="form-name-label">
                Form Name <span className="required">*</span>
              </label>
              <input
                id="form-name-input"
                type="text"
                className="form-name-input"
                value={formNameInput}
                onChange={(e) => setFormNameInput(e.target.value)}
                placeholder="Enter form name"
                autoFocus
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && formNameInput.trim()) {
                    handleSaveToChecklist(formNameInput);
                  } else if (e.key === 'Escape') {
                    setShowFormNameModal(false);
                    setFormNameInput('');
                  }
                }}
              />
              <div className="form-name-modal-footer">
                <button
                  className="btn-cancel"
                  onClick={() => {
                    setShowFormNameModal(false);
                    setFormNameInput('');
                  }}
                >
                  Cancel
                </button>
                <button
                  className="btn-save-checklist"
                  onClick={() => handleSaveToChecklist(formNameInput)}
                  disabled={!formNameInput.trim()}
                >
                  Save to Checklist
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default SEMaster;