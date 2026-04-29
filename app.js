// Register Chart.js datalabels plugin
if (window.ChartDataLabels) Chart.register(ChartDataLabels);

// ============================================================
// State
// ============================================================
let rawEntries = [];
let employeeGroups = {};
let caseGroups = {};
let valueMode = 'billable';       // 'billable' | 'work'
let ungroupedMode = 'individual';  // 'individual' | 'other'
let colMode = 'months';            // 'months' | 'employees'
let groupEmployees = false;
let caseGroupMode = 'groups';       // 'none' | 'client' | 'groups'
let selectedCaseGroups = new Set(); // which case groups/clients to show in pivot
let subChartMode = null;            // 'case' | 'casegroup' | 'client' — null = use smart default
let subChartModeManual = false;     // true once user clicks the sub-chart toggle
let chartDistMode = 'employees';    // 'employees' | 'cases' — pie chart distribution mode
let sortCol = null;                 // null = default order, '__total__' = sort by row total, else col key
let sortDir = 'desc';              // 'asc' | 'desc'
let selectedEmployees = new Set();       // individual employee filter
let selectedEmployeeGroups = new Set();  // employee group filter (when groupEmployees=true)
let caseFilterMode = 'case';             // 'case'|'client'|'groups' — filter mode when caseGroupMode='none'
let phantomEmployees = [];               // ungrouped employees from file not in current data
let phantomCases = [];                   // ungrouped case keys from file not in current data
let selectedSubCharts = new Set();  // empty = all sub-charts shown
let _lastSubChartArgs = null;       // cached for re-render from sub-chart filter
let _caseFilterAllItems = [];      // current item list for case filter
let _empFilterAllItems = [];       // current item list for employee filter
let _empFilterSelectedSet = null;  // reference to active employee selected set
let _subChartFilterAllItems = [];  // current item list for sub-chart filter
// Derived-data cache — invalidated whenever rawEntries changes
let _cache = { valid: false, employees: null, cases: null, clients: null, months: null };
function invalidateCache() { _cache.valid = false; }

let pivotChart = null;              // Chart.js instance
let totalsChart = null;             // Totals bar chart instance
let subCharts = [];                 // Sub chart instances

// ============================================================
// DOM refs
// ============================================================
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// ============================================================
// File Upload
// ============================================================
const uploadArea = $('#upload-area');
const fileInput = $('#file-input');

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('drag-over'));
uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFiles(e.target.files);
});

let fileQueue = Promise.resolve();
function handleFiles(fileList) {
    for (const file of fileList) {
        // Serialize file processing to avoid race conditions on rawEntries
        fileQueue = fileQueue.then(() => handleFile(file));
    }
}

function handleFile(file) {
    return new Promise((resolve) => {
        const name = file.name;

        // Check 1: File extension
        const ext = name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls', 'xlsm'].includes(ext)) {
            alert(`סוג הקובץ "${ext}" אינו נתמך.\nיש להעלות קבצי Excel בלבד (xlsx, xls, xlsm).`);
            resolve(); return;
        }

        // Check 2: File size
        const sizeMB = file.size / (1024 * 1024);
        if (sizeMB > 10) {
            if (!confirm(`הקובץ "${name}" גדול (${sizeMB.toFixed(1)} MB).\nעיבוד קובץ גדול עלול להיות איטי.\n\nלהמשיך?`)) { resolve(); return; }
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const wb = XLSX.read(data, { type: 'array', cellDates: true });

                // Check 3: Empty file
                if (!wb.SheetNames.length) {
                    alert(`הקובץ "${name}" ריק — לא נמצאו גליונות.`);
                    resolve(); return;
                }
                const sheet = wb.Sheets[wb.SheetNames[0]];
                const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
                if (range.e.r === 0 && range.e.c === 0 && !sheet['A1']) {
                    alert(`הקובץ "${name}" ריק — הגליון הראשון אינו מכיל נתונים.`);
                    resolve(); return;
                }

                // Use header row (not first data row) for file-type detection — first data row may have empty cells
                const headerRow = getSheetHeaderRow(wb.Sheets[wb.SheetNames[0]]);
                if (name.includes('עובדים')) {
                    if (headerRow.includes('קבוצה') && headerRow.includes('עובד')) {
                        importEmployeeGroups(wb);
                        showFileStatus(name, 'קבוצות עובדים נטענו');
                    } else {
                        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
                        parseReport(rows);
                        showFileStatus(name, 'דוח שעות נטען');
                    }
                } else if (name.includes('תיקים')) {
                    if (headerRow.includes('קבוצה')) {
                        importCaseGroups(wb);
                        showFileStatus(name, 'קבוצות תיקים נטענו');
                    } else {
                        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
                        parseReport(rows);
                        showFileStatus(name, 'דוח שעות נטען');
                    }
                } else {
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
                    parseReport(rows);
                    showFileStatus(name, 'דוח שעות נטען');
                }
            } catch (err) {
                console.error('Error reading file:', err);
                alert(`שגיאה בקריאת הקובץ "${name}": ${err.message}`);
            }
            resolve();
        };
        reader.readAsArrayBuffer(file);
    });
}

const _fileStatusHistory = [];
function showFileStatus(name, msg) {
    const el = $('#file-name');
    _fileStatusHistory.push(`${msg} (${name})`);
    el.textContent = _fileStatusHistory.join(' | ');
    clearTimeout(el._timer);
    el._timer = setTimeout(() => { _fileStatusHistory.length = 0; el.textContent = ''; }, 6000);
}

// ============================================================
// Validation & Import: Employee Groups
// ============================================================
function importEmployeeGroups(wb) {
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    // Read actual header row to avoid false negatives when first data row has empty cells
    const headerRow = getSheetHeaderRow(sheet);

    // Check 15: Required columns
    if (!rows.length && !headerRow.length) { alert('קובץ קבוצות עובדים ריק'); return; }
    if (!headerRow.includes('קבוצה') || !headerRow.includes('עובד')) {
        alert('קובץ קבוצות עובדים לא תקין.\nנדרשות עמודות: "קבוצה", "עובד"');
        return;
    }

    const errors = [];
    const newGroups = {};
    const memberToGroups = {}; // track which groups each member belongs to
    const newPhantomEmps = []; // ungrouped employees from file not yet in current data

    rows.forEach((r, i) => {
        const group = r['קבוצה'];
        const member = r['עובד'];
        if (!group && !member) return; // skip empty rows
        if (!group) {
            // Ungrouped employee — keep as phantom if not in current data
            if (member) {
                const m = String(member).trim();
                const inData = getAllEmployees().includes(m);
                if (!inData && !newPhantomEmps.includes(m)) newPhantomEmps.push(m);
            }
            return;
        }
        // Check 17: Empty member names
        if (!member) { errors.push(`שורה ${i + 2}: חסר שם עובד`); return; }
        const g = String(group).trim();
        const m = String(member).trim();
        if (!newGroups[g]) newGroups[g] = [];
        if (!newGroups[g].includes(m)) newGroups[g].push(m);
        // Check 18: Track duplicate assignments
        if (!memberToGroups[m]) memberToGroups[m] = [];
        if (!memberToGroups[m].includes(g)) memberToGroups[m].push(g);
    });

    if (errors.length > 0 && Object.keys(newGroups).length === 0) {
        alert('קובץ קבוצות עובדים לא תקין:\n' + errors.slice(0, 5).join('\n'));
        return;
    }

    // Reserved name check — must be first, before any other validation
    if (Object.keys(newGroups).some(g => g.trim() === 'אחר')) {
        alert('קובץ קבוצות עובדים לא תקין.\nלא ניתן להשתמש בשם "אחר" לקבוצה.\nשם זה שמור לפריטים שאינם משויכים לקבוצה.');
        return;
    }

    if (errors.length > 0) {
        alert('אזהרות בטעינת קבוצות עובדים:\n' + errors.slice(0, 5).join('\n') +
            (errors.length > 5 ? `\n...ועוד ${errors.length - 5} אזהרות` : ''));
    }

    // Check 18: Duplicate assignments
    const dups = Object.entries(memberToGroups).filter(([_, groups]) => groups.length > 1);
    if (dups.length > 0) {
        const dupList = dups.slice(0, 5).map(([m, gs]) => `"${m}" → ${gs.join(', ')}`).join('\n');
        alert(`אזהרה: ${dups.length} עובדים משויכים ליותר מקבוצה אחת:\n${dupList}` +
            (dups.length > 5 ? `\n...ועוד ${dups.length - 5}` : '') +
            '\n\nהעובד ישויך לקבוצה האחרונה בלבד.');
    }

    // Check 20: Empty groups
    const emptyGroups = Object.entries(newGroups).filter(([_, members]) => members.length === 0).map(([g]) => g);
    if (emptyGroups.length > 0) {
        alert(`אזהרה: ${emptyGroups.length} קבוצות ריקות (ללא חברים):\n${emptyGroups.join(', ')}`);
    }

    if (Object.keys(employeeGroups).length > 0) {
        if (!confirm(`כבר קיימות ${Object.keys(employeeGroups).length} קבוצות עובדים.\nטעינת הקובץ תחליף את הקבוצות הקיימות.\n\nלהמשיך?`)) return;
    }

    employeeGroups = newGroups;
    phantomEmployees = newPhantomEmps;
    renderEmployeeGroups();
    renderPivot();
}

// ============================================================
// Validation & Import: Case Groups
// ============================================================
function importCaseGroups(wb) {
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);
    // Read actual header row (row 1) to avoid false negatives when first data row has empty cells
    const headerRow = getSheetHeaderRow(sheet);

    if (!rows.length && !headerRow.length) { alert('קובץ קבוצות תיקים ריק'); return; }
    if (!headerRow.includes('קבוצה')) {
        alert('קובץ קבוצות תיקים לא תקין.\nנדרשת עמודה: "קבוצה"\nאופציונלי: "לקוח", "תיק"');
        return;
    }
    // Must have at least one of לקוח or תיק
    if (!headerRow.includes('לקוח') && !headerRow.includes('תיק')) {
        alert('קובץ קבוצות תיקים לא תקין.\nנדרשת לפחות אחת מהעמודות: "לקוח", "תיק"');
        return;
    }

    const errors = [];
    const newGroups = {};
    const newPhantomCases = []; // ungrouped case keys from file not yet in current data

    rows.forEach((r, i) => {
        const group = r['קבוצה'];
        const client = r['לקוח'] || '';
        const cas = r['תיק'] || '';
        if (!group && !client && !cas) return; // skip empty rows
        if (!group) {
            // Ungrouped case — keep as phantom if not in current data
            const key = String(client).trim() + '|' + String(cas).trim();
            if (key !== '|') {
                const inData = getAllCases().some(c => c.key === key);
                if (!inData && !newPhantomCases.includes(key)) newPhantomCases.push(key);
            }
            return;
        }
        const g = String(group).trim();
        const key = String(client).trim() + '|' + String(cas).trim();
        if (key === '|') { errors.push(`שורה ${i + 2}: חסרים לקוח ותיק`); return; }
        if (!newGroups[g]) newGroups[g] = [];
        if (!newGroups[g].includes(key)) newGroups[g].push(key);
    });

    if (errors.length > 0 && Object.keys(newGroups).length === 0) {
        alert('קובץ קבוצות תיקים לא תקין:\n' + errors.slice(0, 5).join('\n'));
        return;
    }

    // Reserved name check — must be first, before any other validation
    if (Object.keys(newGroups).some(g => g.trim() === 'אחר')) {
        alert('קובץ קבוצות תיקים לא תקין.\nלא ניתן להשתמש בשם "אחר" לקבוצה.\nשם זה שמור לפריטים שאינם משויכים לקבוצה.');
        return;
    }

    if (errors.length > 0) {
        alert('אזהרות בטעינת קבוצות תיקים:\n' + errors.slice(0, 5).join('\n') +
            (errors.length > 5 ? `\n...ועוד ${errors.length - 5} אזהרות` : ''));
    }

    // Check 24: Duplicate assignments (same case in multiple groups)
    const caseToGroups = {};
    Object.entries(newGroups).forEach(([g, keys]) => {
        keys.forEach(k => {
            if (!caseToGroups[k]) caseToGroups[k] = [];
            if (!caseToGroups[k].includes(g)) caseToGroups[k].push(g);
        });
    });
    const caseDups = Object.entries(caseToGroups).filter(([_, groups]) => groups.length > 1);
    if (caseDups.length > 0) {
        const dupList = caseDups.slice(0, 5).map(([k, gs]) => `"${k.replace('|', ' / ')}" → ${gs.join(', ')}`).join('\n');
        alert(`אזהרה: ${caseDups.length} תיקים משויכים ליותר מקבוצה אחת:\n${dupList}` +
            (caseDups.length > 5 ? `\n...ועוד ${caseDups.length - 5}` : '') +
            '\n\nהתיק ישויך לקבוצה האחרונה בלבד.');
    }

    // Check 26: Empty groups
    const emptyCaseGroups = Object.entries(newGroups).filter(([_, members]) => members.length === 0).map(([g]) => g);
    if (emptyCaseGroups.length > 0) {
        alert(`אזהרה: ${emptyCaseGroups.length} קבוצות תיקים ריקות (ללא חברים):\n${emptyCaseGroups.join(', ')}`);
    }

    if (Object.keys(caseGroups).length > 0) {
        if (!confirm(`כבר קיימות ${Object.keys(caseGroups).length} קבוצות תיקים.\nטעינת הקובץ תחליף את הקבוצות הקיימות.\n\nלהמשיך?`)) return;
    }

    caseGroups = newGroups;
    phantomCases = newPhantomCases;
    // Select all new groups (+ אחר) in the filter
    Object.keys(newGroups).forEach(g => selectedCaseGroups.add(g));
    selectedCaseGroups.add('אחר');
    renderCaseGroups();
    rebuildCaseFilter();
    renderPivot();
}

// ============================================================
// Validation: Time Report
// ============================================================
function validateReportHeaders(rows) {
    // Scan for header row with "עובד"
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const row = rows[i];
        if (!row) continue;
        for (let j = 0; j < row.length; j++) {
            if (row[j] === 'עובד') return true;
        }
    }
    return false;
}

// ============================================================
// Parse Report (auto-detect format)
// ============================================================
function parseReport(rows) {
    // Check 4: Header row not found
    if (!validateReportHeaders(rows)) {
        // Collect any column headers found in first 30 rows for diagnostic
        const foundHeaders = [];
        for (let i = 0; i < Math.min(rows.length, 30); i++) {
            if (!rows[i]) continue;
            rows[i].forEach(c => { if (c && typeof c === 'string' && c.trim()) foundHeaders.push(c.trim()); });
        }
        const sample = foundHeaders.slice(0, 10).join(', ');
        alert('הקובץ אינו דוח שעות תקין.\nלא נמצאה עמודת "עובד" ב-30 השורות הראשונות.\n\n' +
            (sample ? `תוכן שנמצא: ${sample}\n\n` : '') +
            'אם זהו קובץ קבוצות, העלו אותו בלשונית המתאימה:\n• קבוצות עובדים\n• קבוצות תיקים');
        return;
    }

    // Find header row
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const row = rows[i];
        if (!row) continue;
        for (let j = 0; j < row.length; j++) {
            if (row[j] === 'עובד') { headerRowIdx = i; break; }
        }
        if (headerRowIdx >= 0) break;
    }
    if (headerRowIdx < 0) return;

    // Detect format
    const headerRow = rows[headerRowIdx];
    const headerTexts = headerRow.map(c => c ? String(c).trim() : '');
    const isFormat2 = headerTexts.includes('חשבון') && !headerTexts.includes('לקוח');

    // Check 5: Missing critical columns
    const requiredCols = ['עובד', 'תאריך'];
    const hoursCols = ['שעות חיוב', 'שעות עבודה', 'שעות דיווח'];
    const missingRequired = requiredCols.filter(c => !headerTexts.includes(c));
    const hasHoursCol = hoursCols.some(c => headerTexts.includes(c));
    if (missingRequired.length > 0 || !hasHoursCol) {
        const missing = [...missingRequired];
        if (!hasHoursCol) missing.push('שעות חיוב / שעות עבודה');
        alert(`חסרות עמודות נדרשות בדוח:\n${missing.join(', ')}\n\nעמודות שנמצאו: ${headerTexts.filter(h => h).join(', ')}`);
        return;
    }

    // Parse into a new array (parsers return entries instead of mutating global)
    let newEntries;
    if (isFormat2) {
        newEntries = parseReportFormat2(rows, headerRowIdx, headerRow);
    } else {
        newEntries = parseReportFormat1(rows, headerRowIdx, headerRow);
    }

    // Check 6: No data rows
    if (newEntries.length === 0) { alert('לא נמצאו רשומות תקינות בקובץ.\nוודאו שהקובץ מכיל שורות נתונים מתחת לשורת הכותרת.'); return; }

    // Reject file if entries are missing case, client or date — wrong format
    const formatErrorMsg = 'קובץ שגוי. המערכת מקבלת דוחות משני סוגים: דוח שעות מפורט לפי עורך דין ותאריך ממערכת הדוחות לשותף, או פרוט דיווחי שעות לפי לקוח/תיק שהתקבל מהנהלת חשבונות.';
    const noClientCase = newEntries.filter(e => (!e.client || !e.client.trim()) && (!e.caseName || !e.caseName.trim()));
    if (noClientCase.length > 0) { alert(formatErrorMsg); return; }
    const noDate = newEntries.filter(e => !(e.date instanceof Date) || isNaN(e.date.getTime()));
    if (noDate.length > 0) { alert(formatErrorMsg); return; }

    // Collect all warnings into a single summary
    const warnings = [];

    // Check 8: Negative hours
    const negativeHours = newEntries.filter(e => e.billableHours < 0 || e.workHours < 0);
    if (negativeHours.length > 0) {
        warnings.push(`${negativeHours.length} רשומות עם שעות שליליות.`);
    }

    // Check 9: Unreasonable hours (>24 per entry)
    const unreasonableHours = newEntries.filter(e => e.billableHours > 24 || e.workHours > 24);
    if (unreasonableHours.length > 0) {
        const examples = unreasonableHours.slice(0, 3).map(e =>
            `  ${e.employee}: ${Math.max(e.billableHours, e.workHours)} שעות (${e.caseName})`
        ).join('\n');
        warnings.push(`${unreasonableHours.length} רשומות עם יותר מ-24 שעות:\n${examples}` +
            (unreasonableHours.length > 3 ? `\n  ...ועוד ${unreasonableHours.length - 3}` : ''));
    }

    // Check 12: Duplicate rows
    const keyCount = {};
    newEntries.forEach(e => {
        const k = entryKey(e);
        keyCount[k] = (keyCount[k] || 0) + 1;
    });
    const internalDups = Object.values(keyCount).filter(c => c > 1).reduce((s, c) => s + (c - 1), 0);
    if (internalDups > 0) {
        warnings.push(`${internalDups} שורות כפולות בתוך הקובץ (אותו לקוח, תיק, עובד, תאריך, תיאור ושעות חיוב).`);
    }

    // Show all warnings in a single alert
    if (warnings.length > 0) {
        alert('אזהרות בטעינת הדוח:\n\n• ' + warnings.join('\n• ') + '\n\nהנתונים ייטענו כפי שהם.');
    }

    // If existing data, ask user what to do
    if (rawEntries.length > 0) {
        const action = confirm(
            `כבר טעונות ${rawEntries.length} רשומות.\n` +
            `הקובץ החדש מכיל ${newEntries.length} רשומות.\n\n` +
            `לחץ "אישור" כדי לצרף את הנתונים החדשים לנתונים הקיימים.\n` +
            `לחץ "ביטול" כדי להחליף את הנתונים הקיימים.`
        );

        if (action) {
            // Append mode — check for duplicates
            appendEntries(newEntries);
        } else {
            // Replace mode
            rawEntries = newEntries;
            const totalBillable = newEntries.reduce((s, e) => s + (e.billableHours || 0), 0);
            alert(`הנתונים הוחלפו.\n${newEntries.length} רשומות נטענו.\nסה"כ שעות חיוב: ${totalBillable.toFixed(2)}`);
        }
    } else {
        // First load — just set
        rawEntries = newEntries;
    }

    // Invalidate derived-data cache since rawEntries changed
    invalidateCache();

    // Prune phantoms that are now present in the new data
    if (phantomEmployees.length) {
        const empSet = new Set(getAllEmployees());
        phantomEmployees = phantomEmployees.filter(e => !empSet.has(e));
    }
    if (phantomCases.length) {
        const caseKeySet = new Set(getAllCases().map(c => c.key));
        phantomCases = phantomCases.filter(k => !caseKeySet.has(k));
    }

    // Update date filters based on all current entries
    updateDateFilters();
    rebuildCaseFilter();
    rebuildEmployeeFilter();

    try {
        $('#controls-section').classList.remove('hidden');
        $('#tabs-section').classList.remove('hidden');
        showTab('pivot');
        renderCleanTable();
        renderEmployeeGroups();
        renderCaseGroups();
        renderPivot();
    } catch (err) {
        console.error('Error rendering UI:', err);
        alert('שגיאה בהצגת הנתונים: ' + err.message);
    }
}

function entryKey(e) {
    const dateStr = (e.date instanceof Date && !isNaN(e.date.getTime()))
        ? e.date.toISOString().slice(0, 10) : '';
    return `${e.client}|${e.caseName}|${e.employee}|${dateStr}|${e.billableHours}|${e.description || ''}`;
}

function appendEntries(newEntries) {
    // Find duplicates
    const existingKeys = new Set(rawEntries.map(e => entryKey(e)));
    const duplicates = newEntries.filter(e => existingKeys.has(entryKey(e)));
    const unique = newEntries.filter(e => !existingKeys.has(entryKey(e)));

    let addedEntries;

    if (duplicates.length > 0) {
        const ignoreDups = confirm(
            `נמצאו ${duplicates.length} רשומות כפולות (מתוך ${newEntries.length}).\n` +
            `רשומה כפולה = אותו לקוח, תיק, עובד, תאריך, תיאור ושעות חיוב.\n\n` +
            `לחץ "אישור" כדי להתעלם מהכפולות ולצרף רק ${unique.length} רשומות חדשות.\n` +
            `לחץ "ביטול" כדי לצרף את כל ${newEntries.length} הרשומות כולל הכפולות.`
        );

        if (ignoreDups) {
            addedEntries = unique;
        } else {
            addedEntries = newEntries;
        }
    } else {
        addedEntries = newEntries;
    }

    rawEntries = rawEntries.concat(addedEntries);
    invalidateCache();
    const totalBillable = addedEntries.reduce((s, e) => s + (e.billableHours || 0), 0);
    alert(`${addedEntries.length} רשומות צורפו.\nסה"כ שעות חיוב ברשומות החדשות: ${totalBillable.toFixed(2)}`);
}

function updateDateFilters() {
    const validDates = rawEntries
        .filter(e => e.date instanceof Date && !isNaN(e.date.getTime()))
        .map(e => e.date.getTime());
    if (validDates.length) {
        // Use reduce instead of Math.min/max spread to avoid stack overflow on large arrays
        const minDate = validDates.reduce((a, b) => a < b ? a : b);
        const maxDate = validDates.reduce((a, b) => a > b ? a : b);
        $('#date-from').value = formatDateISO(new Date(minDate));
        $('#date-to').value = formatDateISO(new Date(maxDate));
    } else {
        $('#date-from').value = '';
        $('#date-to').value = '';
    }
}

// ============================================================
// Format 1: Rpt_TD_ByLawyerDate (לקוח + תיק columns in each row)
// ============================================================
function parseReportFormat1(rows, headerRowIdx, headerRow) {
    const entries = [];
    const colMap = {};
    const headerNames = {
        'תיאור': 'description', 'סה"כ': 'total', 'סה״כ': 'total',
        'תעריף': 'rate', 'שעות חיוב': 'billableHours', 'שעות עבודה': 'workHours',
        'סטטוס': 'status', 'תיק': 'caseName', 'לקוח': 'clientName',
        'תאריך': 'date', 'עובד': 'employee'
    };
    for (let j = 0; j < headerRow.length; j++) {
        const cell = headerRow[j];
        if (cell && headerNames[cell]) colMap[headerNames[cell]] = j;
    }

    let currentEmployee = null;
    let currentDate = null;

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const dateCell = row[colMap.date];
        if (typeof dateCell === 'string' && dateCell.includes('סה')) continue;

        const desc = row[colMap.description];
        const empCell = row[colMap.employee];
        const caseCell = row[colMap.caseName];
        const clientCell = row[colMap.clientName];

        if (!desc && !empCell && !caseCell) continue;

        if (empCell) currentEmployee = String(empCell).trim();

        if (dateCell != null && typeof dateCell !== 'string') {
            if (dateCell instanceof Date) currentDate = dateCell;
            else if (typeof dateCell === 'number') currentDate = excelDateToJS(dateCell);
        } else if (typeof dateCell === 'string' && dateCell.trim() && !dateCell.includes('סה')) {
            const parsed = parseDate(dateCell);
            if (parsed) currentDate = parsed;
        }

        if (!desc) continue;

        const billableHours = toNum(row[colMap.billableHours]);
        const workHours = toNum(row[colMap.workHours]);
        if (billableHours === null && workHours === null) continue;

        entries.push({
            employee: currentEmployee || '',
            date: currentDate,
            description: String(desc),
            client: clientCell ? String(clientCell) : '',
            caseName: caseCell ? String(caseCell) : '',
            caseKey: (clientCell ? String(clientCell) : '') + '|' + (caseCell ? String(caseCell) : ''),
            status: row[colMap.status] ? String(row[colMap.status]) : '',
            rate: toNum(row[colMap.rate]) || 0,
            billableHours: billableHours || 0,
            workHours: workHours || 0,
            total: toNum(row[colMap.total]) || 0
        });
    }
    return entries;
}

// ============================================================
// Format 2: פרוט דיווחי שעות לפי לקוח/תיק (section headers for client/case)
// ============================================================
function parseReportFormat2(rows, headerRowIdx, headerRow) {
    const entries = [];
    const colMap = {};
    const headerNames = {
        'תיאור': 'description', 'סה"כ': 'total', 'סה״כ': 'total',
        'תעריף': 'rate', 'שעות חיוב': 'billableHours', 'שעות דיווח': 'workHours',
        'מחיקות': 'deletions', 'סטטוס': 'status', 'חשבון': 'account',
        'תאריך': 'date', 'עובד': 'employee'
    };
    for (let j = 0; j < headerRow.length; j++) {
        const cell = headerRow[j];
        if (cell && headerNames[String(cell).trim()]) colMap[headerNames[String(cell).trim()]] = j;
    }

    let currentClient = '';
    let currentCase = '';
    let orphanRows = 0; // Check 13: data rows before any section header
    let lastSectionHeader = null; // Check 14: track section headers without data
    let emptySections = []; // Check 14

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const firstCell = row[0] != null ? String(row[0]).trim() : '';

        // Skip subtotal rows
        if (firstCell.startsWith('סה"כ') || firstCell.startsWith('סה״כ')) continue;

        // Detect client section header: "לקוח: NNN - Name"
        // Keep full "NNN - Name" and normalize to match Format 1:
        //   - replace " (double quote) with '' (two single quotes)
        //   - replace hyphen between Hebrew words with space
        if (firstCell.startsWith('לקוח:')) {
            // Check 14: previous section had no data
            if (lastSectionHeader && lastSectionHeader.dataCount === 0) {
                emptySections.push(lastSectionHeader.label);
            }
            currentClient = firstCell.replace(/^לקוח:\s*/, '').trim()
                .replace(/"/g, "''")
                .replace(/([\u0590-\u05FF])-(?=[\u0590-\u05FF])/g, '$1 ');
            lastSectionHeader = { label: 'לקוח: ' + currentClient, dataCount: 0 };
            continue;
        }

        // Detect case section header: "תיק: N - Name"
        // Keep full "N - Name" to match Format 1 output
        if (firstCell.match(/^\s*תיק:/)) {
            // Check 14: previous section had no data
            if (lastSectionHeader && lastSectionHeader.dataCount === 0) {
                emptySections.push(lastSectionHeader.label);
            }
            currentCase = firstCell.replace(/^\s*תיק:\s*/, '').trim();
            lastSectionHeader = { label: 'תיק: ' + currentCase, dataCount: 0 };
            continue;
        }

        // Data row: must have employee and description
        const empCell = row[colMap.employee];
        const desc = row[colMap.description];
        if (!empCell && !desc) continue;

        // Check 13: data row before any section header
        if (!currentClient && !currentCase) {
            orphanRows++;
        }

        const dateCell = row[colMap.date];
        let entryDate = null;
        if (dateCell != null && typeof dateCell !== 'string') {
            if (dateCell instanceof Date) entryDate = dateCell;
            else if (typeof dateCell === 'number') entryDate = excelDateToJS(dateCell);
        } else if (typeof dateCell === 'string' && dateCell.trim()) {
            entryDate = parseDate(dateCell);
        }

        if (!desc) continue;

        const billableHours = toNum(row[colMap.billableHours]);
        const workHours = toNum(row[colMap.workHours]);
        if (billableHours === null && workHours === null) continue;

        // Track data count for Check 14
        if (lastSectionHeader) lastSectionHeader.dataCount++;

        entries.push({
            employee: empCell ? String(empCell).trim() : '',
            date: entryDate,
            description: String(desc),
            client: currentClient,
            caseName: currentCase,
            caseKey: currentClient + '|' + currentCase,
            status: row[colMap.status] ? String(row[colMap.status]) : '',
            rate: toNum(row[colMap.rate]) || 0,
            billableHours: billableHours || 0,
            workHours: workHours || 0,
            total: toNum(row[colMap.total]) || 0
        });
    }

    // Check 14: last section header without data
    if (lastSectionHeader && lastSectionHeader.dataCount === 0) {
        emptySections.push(lastSectionHeader.label);
    }

    return entries;
}

// ============================================================
// Case filter checkboxes
// ============================================================
function rebuildCaseFilter() {
    const filterGroup = $('#case-filter-group');
    const filterLabel = $('#case-filter-label');
    const modeToggle = $('#case-filter-mode-toggle');

    if (filterGroup) filterGroup.style.display = '';

    let items, labelFn = null;

    if (caseGroupMode === 'none') {
        // Show the sub-mode toggle
        if (modeToggle) {
            modeToggle.classList.remove('hidden');
            modeToggle.querySelectorAll('.toggle-btn').forEach(b => b.classList.toggle('active', b.dataset.value === caseFilterMode));
            // Disable 'groups' button if no case groups defined
            const groupsBtn = modeToggle.querySelector('[data-value="groups"]');
            if (groupsBtn) {
                const hasGroups = Object.keys(caseGroups).length > 0;
                groupsBtn.disabled = !hasGroups;
                groupsBtn.style.opacity = hasGroups ? '' : '0.4';
            }
        }
        if (caseFilterMode === 'case') {
            items = getAllCases().map(c => c.key);
            labelFn = (k) => caseLabel(k);
            if (filterLabel) filterLabel.textContent = 'תיקים להצגה:';
        } else if (caseFilterMode === 'client') {
            items = getAllClients();
            if (filterLabel) filterLabel.textContent = 'לקוחות להצגה:';
        } else { // 'groups'
            items = [...Object.keys(caseGroups), 'אחר'];
            if (filterLabel) filterLabel.textContent = 'קבוצות תיקים להצגה:';
        }
    } else {
        if (modeToggle) modeToggle.classList.add('hidden');
        if (caseGroupMode === 'client') {
            items = getAllClients();
            if (filterLabel) filterLabel.textContent = 'לקוחות להצגה:';
        } else {
            items = [...Object.keys(caseGroups), 'אחר'];
            if (filterLabel) filterLabel.textContent = 'קבוצות תיקים להצגה:';
        }
    }

    _caseFilterAllItems = items;
    const validSet = new Set(items);
    selectedCaseGroups = new Set([...selectedCaseGroups].filter(i => validSet.has(i)));
    if (selectedCaseGroups.size === 0) items.forEach(i => selectedCaseGroups.add(i));

    renderCaseFilterList(items, '', labelFn);
    updateCaseFilterTrigger(items);
}

function renderCaseFilterList(allItems, searchTerm, labelFn = null) {
    const list = $('#case-filter-list');
    if (!list) return;
    const term = searchTerm.toLowerCase();
    const displayLabel = labelFn || ((i) => i);
    const filtered = term ? allItems.filter(i => displayLabel(i).toLowerCase().includes(term)) : allItems;

    const allChecked = allItems.length > 0 && allItems.every(i => selectedCaseGroups.has(i));
    const someChecked = !allChecked && allItems.some(i => selectedCaseGroups.has(i));

    list.innerHTML = '';

    const allLi = document.createElement('li');
    allLi.className = 'xf-item xf-select-all';
    allLi.innerHTML = `<label><input type="checkbox" /><span>בחר הכל</span></label>`;
    const allCb = allLi.querySelector('input');
    allCb.checked = allChecked;
    allCb.indeterminate = someChecked;
    allCb.addEventListener('change', () => {
        if (allCb.checked) allItems.forEach(i => selectedCaseGroups.add(i));
        else allItems.forEach(i => selectedCaseGroups.delete(i));
        renderCaseFilterList(allItems, $('#case-filter-search').value, labelFn);
        updateCaseFilterTrigger(allItems);
        debouncedRenderPivot();
    });
    list.appendChild(allLi);

    filtered.forEach(item => {
        const li = document.createElement('li');
        li.className = 'xf-item';
        li.innerHTML = `<label><input type="checkbox" /><span>${esc(displayLabel(item))}</span></label>`;
        const cb = li.querySelector('input');
        cb.checked = selectedCaseGroups.has(item);
        cb.addEventListener('change', () => {
            if (cb.checked) selectedCaseGroups.add(item);
            else selectedCaseGroups.delete(item);
            renderCaseFilterList(allItems, $('#case-filter-search').value, labelFn);
            updateCaseFilterTrigger(allItems);
            debouncedRenderPivot();
        });
        list.appendChild(li);
    });
}

function updateCaseFilterTrigger(allItems) {
    const trigger = $('#case-filter-trigger');
    if (!trigger) return;
    const selected = allItems.filter(i => selectedCaseGroups.has(i)).length;
    trigger.textContent = selected === allItems.length ? 'הכל ▾' : `${selected} מתוך ${allItems.length} ▾`;
}

function rebuildEmployeeFilter() {
    const filterLabel = $('#employee-filter-group label');
    if (groupEmployees && Object.keys(employeeGroups).length > 0) {
        // Build group list
        const allGroups = [...Object.keys(employeeGroups)];
        const groupedSet = new Set(Object.values(employeeGroups).flat());
        const hasUngrouped = getAllEmployees().some(e => !groupedSet.has(e));
        if (hasUngrouped) allGroups.push('אחר');
        if (allGroups.length === 0) return;

        const validSet = new Set(allGroups);
        selectedEmployeeGroups = new Set([...selectedEmployeeGroups].filter(g => validSet.has(g)));
        if (selectedEmployeeGroups.size === 0) allGroups.forEach(g => selectedEmployeeGroups.add(g));

        if (filterLabel) filterLabel.textContent = 'קבוצות עובדים להצגה:';
        _empFilterAllItems = allGroups;
        _empFilterSelectedSet = selectedEmployeeGroups;
        renderEmpFilterList(allGroups, '', selectedEmployeeGroups);
        updateEmpFilterTrigger(allGroups, selectedEmployeeGroups);
    } else {
        const allEmps = getAllEmployees();
        if (allEmps.length === 0) return;

        const validEmpSet = new Set(allEmps);
        selectedEmployees = new Set([...selectedEmployees].filter(e => validEmpSet.has(e)));
        if (selectedEmployees.size === 0) allEmps.forEach(e => selectedEmployees.add(e));

        if (filterLabel) filterLabel.textContent = 'עובדים להצגה:';
        _empFilterAllItems = allEmps;
        _empFilterSelectedSet = selectedEmployees;
        renderEmpFilterList(allEmps, '', selectedEmployees);
        updateEmpFilterTrigger(allEmps, selectedEmployees);
    }
}

function renderEmpFilterList(allItems, searchTerm, selectedSet) {
    const list = $('#emp-filter-list');
    if (!list) return;
    const term = searchTerm.toLowerCase();
    const filtered = term ? allItems.filter(e => e.toLowerCase().includes(term)) : allItems;

    // "Select all" state reflects ALL items, not just filtered
    const allChecked = allItems.length > 0 && allItems.every(e => selectedSet.has(e));
    const someChecked = !allChecked && allItems.some(e => selectedSet.has(e));

    list.innerHTML = '';

    const allLi = document.createElement('li');
    allLi.className = 'xf-item xf-select-all';
    allLi.innerHTML = `<label><input type="checkbox" /><span>בחר הכל</span></label>`;
    const allCb = allLi.querySelector('input');
    allCb.checked = allChecked;
    allCb.indeterminate = someChecked;
    allCb.addEventListener('change', () => {
        if (allCb.checked) allItems.forEach(e => selectedSet.add(e));
        else allItems.forEach(e => selectedSet.delete(e));
        renderEmpFilterList(allItems, $('#emp-filter-search').value, selectedSet);
        updateEmpFilterTrigger(allItems, selectedSet);
        debouncedRenderPivot();
    });
    list.appendChild(allLi);

    filtered.forEach(item => {
        const li = document.createElement('li');
        li.className = 'xf-item';
        li.innerHTML = `<label><input type="checkbox" /><span>${esc(item)}</span></label>`;
        const cb = li.querySelector('input');
        cb.value = item;
        cb.checked = selectedSet.has(item);
        cb.addEventListener('change', () => {
            if (cb.checked) selectedSet.add(item);
            else selectedSet.delete(item);
            renderEmpFilterList(allItems, $('#emp-filter-search').value, selectedSet);
            updateEmpFilterTrigger(allItems, selectedSet);
            debouncedRenderPivot();
        });
        list.appendChild(li);
    });
}

function updateEmpFilterTrigger(allItems, selectedSet) {
    const trigger = $('#emp-filter-trigger');
    if (!trigger) return;
    const selected = allItems.filter(e => selectedSet.has(e)).length;
    trigger.textContent = selected === allItems.length ? 'הכל ▾' : `${selected} מתוך ${allItems.length} ▾`;
}

// ============================================================
// Utility functions
// ============================================================
function toNum(val) {
    if (val == null) return null;
    if (typeof val === 'number') return val;
    const n = parseFloat(val);
    return isNaN(n) ? null : n;
}

function excelDateToJS(serial) {
    // Excel incorrectly treats 1900 as a leap year; serial 60 = Feb 29, 1900 (invalid)
    // For serials > 60, subtract 1 day to correct for the phantom leap day
    const adjusted = serial > 60 ? serial - 1 : serial;
    return new Date(Date.UTC(1899, 11, 30 + Math.floor(adjusted)));
}

function parseDate(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'string') {
        const trimmed = val.trim();
        // Try DD/MM/YYYY first (Hebrew date format)
        const parts = trimmed.split('/');
        if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
                const d = new Date(year, month - 1, day);
                // Validate the components match to catch rollover (e.g., day 35 → next month)
                if (!isNaN(d.getTime()) && d.getDate() === day && d.getMonth() === month - 1) return d;
            }
        }
        // Fallback to native parsing (ISO format, etc.)
        const d = new Date(trimmed);
        if (!isNaN(d.getTime())) return d;
    }
    return null;
}

function formatDateISO(d) {
    if (!d) return '';
    const dd = d instanceof Date ? d : new Date(d);
    if (isNaN(dd.getTime())) return '';
    return `${dd.getFullYear()}-${String(dd.getMonth() + 1).padStart(2, '0')}-${String(dd.getDate()).padStart(2, '0')}`;
}

function formatDateHebrew(d) {
    if (!d) return '';
    const dd = d instanceof Date ? d : new Date(d);
    if (isNaN(dd.getTime())) return '';
    return `${String(dd.getDate()).padStart(2, '0')}/${String(dd.getMonth() + 1).padStart(2, '0')}/${dd.getFullYear()}`;
}

function formatMonth(d) {
    if (!d) return '';
    const dd = d instanceof Date ? d : new Date(d);
    if (isNaN(dd.getTime())) return '';
    return `${String(dd.getMonth() + 1).padStart(2, '0')}/${dd.getFullYear()}`;
}

function getFilteredEntries() {
    let entries = rawEntries;
    const from = $('#date-from').value;
    const to = $('#date-to').value;
    const hasDateFilter = from || to;
    if (hasDateFilter) {
        // When a date filter is active, exclude entries with no valid date
        entries = entries.filter(e => e.date instanceof Date && !isNaN(e.date.getTime()));
    }
    if (from) {
        const fd = new Date(from + 'T00:00:00');
        if (!isNaN(fd.getTime())) entries = entries.filter(e => e.date >= fd);
    }
    if (to) {
        const td = new Date(to + 'T23:59:59');
        if (!isNaN(td.getTime())) entries = entries.filter(e => e.date <= td);
    }
    if (groupEmployees && Object.keys(employeeGroups).length > 0) {
        const empGroupMap = {};
        Object.entries(employeeGroups).forEach(([g, members]) => members.forEach(m => { empGroupMap[m] = g; }));
        entries = entries.filter(e => {
            const group = empGroupMap[e.employee] || (ungroupedMode === 'individual' ? e.employee : 'אחר');
            return selectedEmployeeGroups.has(group);
        });
    } else {
        entries = entries.filter(e => selectedEmployees.has(e.employee));
    }
    return entries;
}

function _buildCache() {
    if (_cache.valid) return;
    const employees = new Set();
    const cases = new Map();
    const clients = new Set();
    const months = new Set();
    for (const e of rawEntries) {
        if (e.employee) employees.add(e.employee);
        if (!cases.has(e.caseKey)) cases.set(e.caseKey, { client: e.client, caseName: e.caseName, key: e.caseKey });
        if (e.client && e.client.trim()) clients.add(e.client);
        const m = formatMonth(e.date);
        if (m) months.add(m);
    }
    _cache.employees = [...employees].sort();
    _cache.cases = [...cases.values()].sort((a, b) => a.key.localeCompare(b.key));
    _cache.clients = [...clients].sort();
    _cache.months = [...months].sort((a, b) => {
        const [mA, yA] = a.split('/').map(Number);
        const [mB, yB] = b.split('/').map(Number);
        return yA !== yB ? yA - yB : mA - mB;
    });
    _cache.valid = true;
}

function getAllEmployees() { _buildCache(); return _cache.employees; }
function getAllCases()     { _buildCache(); return _cache.cases; }
function getAllClients()   { _buildCache(); return _cache.clients; }
function getAllMonths()    { _buildCache(); return _cache.months; }

function getAssignedEmployees() {
    const assigned = new Set();
    Object.values(employeeGroups).forEach(members => members.forEach(m => assigned.add(m)));
    return assigned;
}

function getAssignedCases() {
    const assigned = new Set();
    Object.values(caseGroups).forEach(members => members.forEach(m => assigned.add(m)));
    return assigned;
}

function caseLabel(key) {
    const parts = key.split('|');
    const client = parts[0] || '';
    const cas = parts[1] || '';
    const clientShort = client.length > 30 ? client.substring(0, 30) + '...' : client;
    return `${clientShort} / ${cas}`;
}


// ============================================================
// Pivot table sort — delegated, one-time
(function initPivotSort() {
    const table = $('#pivot-table');
    if (!table) return;
    table.addEventListener('click', (e) => {
        const th = e.target.closest('th.pivot-sortable');
        if (!th) return;
        const col = th.dataset.col;
        if (sortCol === col) {
            sortDir = sortDir === 'desc' ? 'asc' : 'desc';
        } else {
            sortCol = col;
            sortDir = 'desc';
        }
        debouncedRenderPivot();
    });
})();

// Debounced render for performance on rapid control changes
// ============================================================
let _renderPivotTimer;
function debouncedRenderPivot() {
    clearTimeout(_renderPivotTimer);
    _renderPivotTimer = setTimeout(renderPivot, 120);
}

// ============================================================
// Tabs
// ============================================================
$$('.tab').forEach(tab => {
    tab.addEventListener('click', () => showTab(tab.dataset.tab));
});

function showTab(tabId) {
    $$('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tabId));
    $$('.tab-content').forEach(tc => tc.classList.toggle('hidden', tc.id !== tabId));
    if (tabId === 'pivot') renderPivot();
    if (tabId === 'clean-table') renderCleanTable();
}

// ============================================================
// Controls
// ============================================================
// Value toggle
$('#value-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $('#value-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        valueMode = btn.dataset.value;
        debouncedRenderPivot();
    });
});

// Case filter mode toggle (only shown when caseGroupMode === 'none')
$('#case-filter-mode-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        if (btn.disabled) return;
        $('#case-filter-mode-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        caseFilterMode = btn.dataset.value;
        selectedCaseGroups.clear(); // reset selection for new filter type
        rebuildCaseFilter();
        debouncedRenderPivot();
    });
});

// Ungrouped toggle
$('#ungrouped-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $('#ungrouped-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        ungroupedMode = btn.dataset.value;
        debouncedRenderPivot();
    });
});

// Column mode toggle (months/employees)
$('#col-mode-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $('#col-mode-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        colMode = btn.dataset.value;
        // Enable/disable group employees checkbox
        const cb = $('#group-employees-cb');
        cb.disabled = (colMode !== 'employees');
        if (colMode !== 'employees') {
            cb.checked = false;
            groupEmployees = false;
        }
        debouncedRenderPivot();
    });
});

// Group employees checkbox
$('#group-employees-cb').addEventListener('change', (e) => {
    groupEmployees = e.target.checked;
    rebuildEmployeeFilter();
    debouncedRenderPivot();
});

// Case group mode toggle (none/client/groups)
$('#case-group-mode-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $('#case-group-mode-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        caseGroupMode = btn.dataset.value;
        // Show/hide ungrouped toggle — only relevant in 'groups' mode
        const ungroupedToggle = $('#ungrouped-toggle').closest('.control-group');
        if (ungroupedToggle) ungroupedToggle.style.display = caseGroupMode === 'groups' ? '' : 'none';
        if (caseGroupMode === 'none') caseFilterMode = 'case'; // reset sub-filter mode
        selectedCaseGroups.clear(); // reset filter for new mode
        rebuildCaseFilter();
        debouncedRenderPivot();
    });
});

// Sub-chart mode toggle (below main chart, employees mode only)
$('#sub-chart-mode-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $('#sub-chart-mode-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        subChartMode = btn.dataset.value;
        subChartModeManual = true;
        debouncedRenderPivot();
    });
});

// Chart distribution mode toggle (employees vs cases)
$('#chart-dist-mode-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $('#chart-dist-mode-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        chartDistMode = btn.dataset.value;
        debouncedRenderPivot();
    });
});

// Date controls
$('#date-from').addEventListener('change', () => { renderCleanTable(); debouncedRenderPivot(); });
$('#date-to').addEventListener('change', () => { renderCleanTable(); debouncedRenderPivot(); });
$('#clear-dates').addEventListener('click', () => {
    $('#date-from').value = '';
    $('#date-to').value = '';
    renderCleanTable();
    debouncedRenderPivot();
});

// ============================================================
// Clean Table
// ============================================================
function renderCleanTable() {
    const entries = getFilteredEntries();
    const thead = $('#clean-table-data thead');
    const tbody = $('#clean-table-data tbody');

    thead.innerHTML = '<tr><th>עובד</th><th>תאריך</th><th>חודש</th><th>לקוח</th><th>תיק</th><th>תיאור</th><th>סטטוס</th><th>תעריף</th><th>שעות עבודה</th><th>שעות חיוב</th><th>סה"כ</th></tr>';
    tbody.innerHTML = entries.map(e => `<tr>
        <td>${esc(e.employee)}</td>
        <td>${formatDateHebrew(e.date)}</td>
        <td>${formatMonth(e.date)}</td>
        <td>${esc(e.client)}</td>
        <td>${esc(e.caseName)}</td>
        <td>${esc(e.description)}</td>
        <td>${esc(e.status)}</td>
        <td>${esc(String(e.rate))}</td>
        <td>${esc(String(e.workHours))}</td>
        <td>${esc(String(e.billableHours))}</td>
        <td>${esc(String(e.total))}</td>
    </tr>`).join('');
}

$('#download-clean').addEventListener('click', () => {
    const entries = getFilteredEntries();
    const data = entries.map(e => ({
        'עובד': e.employee, 'תאריך': formatDateHebrew(e.date), 'חודש': formatMonth(e.date),
        'לקוח': e.client, 'תיק': e.caseName, 'תיאור': e.description,
        'סטטוס': e.status, 'תעריף': e.rate, 'שעות עבודה': e.workHours,
        'שעות חיוב': e.billableHours, 'סה"כ': e.total
    }));
    downloadExcel(data, 'נתונים_מלאים.xlsx');
});

// ============================================================
// Employee Groups
// ============================================================
$('#add-emp-group').addEventListener('click', () => {
    const name = $('#new-emp-group-name').value.trim();
    if (!name || employeeGroups[name]) return;
    if (name === 'אחר') { alert('לא ניתן ליצור קבוצה בשם "אחר".\nשם זה שמור לפריטים שאינם משויכים לקבוצה.'); return; }
    employeeGroups[name] = [];
    $('#new-emp-group-name').value = '';
    renderEmployeeGroups();
    renderPivot();
});

$('#new-emp-group-name').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') $('#add-emp-group').click();
});

function renderEmployeeGroups() {
    const assigned = getAssignedEmployees();
    const allEmps = getAllEmployees();
    const unassigned = allEmps.filter(e => !assigned.has(e));

    const groupsList = $('#emp-groups-list');
    groupsList.innerHTML = Object.entries(employeeGroups).map(([name, members]) =>
        `<div class="group-card">
            <div class="group-header">
                <strong>${esc(name)}</strong>
                <button class="rename-btn" data-action="rename-emp-group" data-name="${escData(name)}" title="שנה שם">✎</button>
                <button data-action="remove-emp-group" data-name="${escData(name)}" title="מחק קבוצה">&times;</button>
            </div>
            <div class="group-members" data-group="${escData(name)}" data-type="emp">
                ${members.map(m => `<span class="member-tag" draggable="true" data-member="${escData(m)}" data-type="emp">
                    ${esc(m)} <span class="remove-member" data-action="remove-emp-member" data-group="${escData(name)}" data-member="${escData(m)}">&times;</span>
                </span>`).join('')}
            </div>
        </div>`
    ).join('');

    $('#unassigned-employees').innerHTML = unassigned.map(e =>
        `<span class="unassigned-tag" draggable="true" data-member="${escData(e)}" data-type="emp">${esc(e)}</span>`
    ).join('');

    setupDragDrop('emp');
}

// ============================================================
// Case Groups
// ============================================================
$('#add-case-group').addEventListener('click', () => {
    const name = $('#new-case-group-name').value.trim();
    if (!name || caseGroups[name]) return;
    if (name === 'אחר') { alert('לא ניתן ליצור קבוצה בשם "אחר".\nשם זה שמור לפריטים שאינם משויכים לקבוצה.'); return; }
    caseGroups[name] = [];
    $('#new-case-group-name').value = '';
    renderCaseGroups();
    rebuildCaseFilter();
    renderPivot();
});

$('#new-case-group-name').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') $('#add-case-group').click();
});

function renderCaseGroups() {
    const assigned = getAssignedCases();
    const allCases = getAllCases();
    const unassigned = allCases.filter(c => !assigned.has(c.key));

    const groupsList = $('#case-groups-list');
    groupsList.innerHTML = Object.entries(caseGroups).map(([name, members]) =>
        `<div class="group-card">
            <div class="group-header">
                <strong>${esc(name)}</strong>
                <button class="rename-btn" data-action="rename-case-group" data-name="${escData(name)}" title="שנה שם">✎</button>
                <button data-action="remove-case-group" data-name="${escData(name)}" title="מחק קבוצה">&times;</button>
            </div>
            <div class="group-members" data-group="${escData(name)}" data-type="case">
                ${members.map(m => `<span class="member-tag" draggable="true" data-member="${escData(m)}" data-type="case">
                    ${esc(caseLabel(m))} <span class="remove-member" data-action="remove-case-member" data-group="${escData(name)}" data-member="${escData(m)}">&times;</span>
                </span>`).join('')}
            </div>
        </div>`
    ).join('');

    $('#unassigned-cases').innerHTML = unassigned.map(c =>
        `<span class="unassigned-tag" draggable="true" data-member="${escData(c.key)}" data-type="case">${esc(caseLabel(c.key))}</span>`
    ).join('');

    setupDragDrop('case');
}

// ============================================================
// Group Clicks — one-time delegated event listeners (avoid accumulation)
// ============================================================
(function initGroupClicks() {
    function handleGroupClick(e) {
        const btn = e.target.closest('[data-action]');
        if (!btn) return;
        const action = btn.dataset.action;
        if (action === 'rename-emp-group') {
            const oldName = btn.dataset.name;
            const newName = prompt('שם חדש לקבוצה:', oldName)?.trim();
            if (!newName || newName === oldName) return;
            if (newName === 'אחר') { alert('לא ניתן להשתמש בשם "אחר"'); return; }
            if (employeeGroups[newName]) { alert(`קבוצה בשם "${newName}" כבר קיימת`); return; }
            employeeGroups[newName] = employeeGroups[oldName];
            delete employeeGroups[oldName];
            if (selectedEmployeeGroups.has(oldName)) { selectedEmployeeGroups.delete(oldName); selectedEmployeeGroups.add(newName); }
            rebuildEmployeeFilter(); renderEmployeeGroups(); renderPivot();
        } else if (action === 'rename-case-group') {
            const oldName = btn.dataset.name;
            const newName = prompt('שם חדש לקבוצה:', oldName)?.trim();
            if (!newName || newName === oldName) return;
            if (newName === 'אחר') { alert('לא ניתן להשתמש בשם "אחר"'); return; }
            if (caseGroups[newName]) { alert(`קבוצה בשם "${newName}" כבר קיימת`); return; }
            caseGroups[newName] = caseGroups[oldName];
            delete caseGroups[oldName];
            if (selectedCaseGroups.has(oldName)) { selectedCaseGroups.delete(oldName); selectedCaseGroups.add(newName); }
            renderCaseGroups(); rebuildCaseFilter(); renderPivot();
        } else if (action === 'remove-emp-group') { delete employeeGroups[btn.dataset.name]; rebuildEmployeeFilter(); renderEmployeeGroups(); renderPivot(); }
        else if (action === 'remove-emp-member') { employeeGroups[btn.dataset.group] = employeeGroups[btn.dataset.group].filter(m => m !== btn.dataset.member); renderEmployeeGroups(); renderPivot(); }
        else if (action === 'remove-case-group') { delete caseGroups[btn.dataset.name]; renderCaseGroups(); rebuildCaseFilter(); renderPivot(); }
        else if (action === 'remove-case-member') { caseGroups[btn.dataset.group] = caseGroups[btn.dataset.group].filter(m => m !== btn.dataset.member); renderCaseGroups(); rebuildCaseFilter(); renderPivot(); }
    }
    const empContainer = $('#employee-groups');
    const caseContainer = $('#case-groups');
    if (empContainer) empContainer.addEventListener('click', handleGroupClick);
    if (caseContainer) caseContainer.addEventListener('click', handleGroupClick);
})();

// ============================================================
// Drag & Drop — uses delegation on stable parent containers
// ============================================================
function setupDragDrop(type) {
    const container = type === 'emp' ? $('#employee-groups') : $('#case-groups');
    if (!container) return;

    // Only attach dragstart to newly rendered draggable elements (these are recreated each render)
    container.querySelectorAll('[draggable="true"]').forEach(tag => {
        tag.addEventListener('dragstart', (e) => {
            e.dataTransfer.setData('text/plain', JSON.stringify({
                member: tag.dataset.member, type: tag.dataset.type,
                fromGroup: tag.closest('.group-members')?.dataset.group || null
            }));
        });
    });
}

// One-time delegated drag-over/drop listeners on stable parent containers
(function initDragDropDelegation() {
    function handleDrop(e, type) {
        const zone = e.target.closest('.group-members');
        if (!zone || zone.dataset.type !== type) return;
        e.preventDefault(); zone.classList.remove('drag-over');
        let payload;
        try { payload = JSON.parse(e.dataTransfer.getData('text/plain')); } catch { return; }
        if (!payload || typeof payload.member !== 'string' || payload.type !== type) return;
        const groups = type === 'emp' ? employeeGroups : caseGroups;
        if (payload.fromGroup && groups[payload.fromGroup]) groups[payload.fromGroup] = groups[payload.fromGroup].filter(m => m !== payload.member);
        if (groups[zone.dataset.group] && !groups[zone.dataset.group].includes(payload.member)) groups[zone.dataset.group].push(payload.member);
        if (type === 'emp') renderEmployeeGroups(); else { renderCaseGroups(); rebuildCaseFilter(); }
        renderPivot();
    }

    function handleUnassignedDrop(e, type) {
        e.preventDefault();
        e.currentTarget.classList.remove('drag-over');
        let payload;
        try { payload = JSON.parse(e.dataTransfer.getData('text/plain')); } catch { return; }
        if (!payload || typeof payload.member !== 'string' || payload.type !== type) return;
        const groups = type === 'emp' ? employeeGroups : caseGroups;
        if (payload.fromGroup && groups[payload.fromGroup]) groups[payload.fromGroup] = groups[payload.fromGroup].filter(m => m !== payload.member);
        if (type === 'emp') renderEmployeeGroups(); else { renderCaseGroups(); rebuildCaseFilter(); }
        renderPivot();
    }

    ['emp', 'case'].forEach(type => {
        const container = type === 'emp' ? $('#employee-groups') : $('#case-groups');
        if (!container) return;

        // Delegated dragover/dragleave/drop on the container
        container.addEventListener('dragover', (e) => {
            const zone = e.target.closest('.group-members');
            if (zone && zone.dataset.type === type) { e.preventDefault(); zone.classList.add('drag-over'); }
        });
        container.addEventListener('dragleave', (e) => {
            const zone = e.target.closest('.group-members');
            if (zone) zone.classList.remove('drag-over');
        });
        container.addEventListener('drop', (e) => handleDrop(e, type));

        // Unassigned zone — stable element
        const unassignedZone = type === 'emp' ? $('#unassigned-employees') : $('#unassigned-cases');
        if (!unassignedZone) return;
        unassignedZone.addEventListener('dragover', (e) => { e.preventDefault(); unassignedZone.classList.add('drag-over'); });
        unassignedZone.addEventListener('dragleave', () => unassignedZone.classList.remove('drag-over'));
        unassignedZone.addEventListener('drop', (e) => handleUnassignedDrop(e, type));
    });
})();

// ============================================================
// Pivot Table
// ============================================================
function renderPivot() {
    const entries = getFilteredEntries();
    if (!entries.length) {
        $('#pivot-table thead').innerHTML = '';
        $('#pivot-table tbody').innerHTML = '<tr><td>אין נתונים להצגה</td></tr>';
        return;
    }

    const hourKey = valueMode === 'billable' ? 'billableHours' : 'workHours';

    // --- Build COLUMNS ---
    let cols = [];
    let getCol; // function: entry -> column label

    if (colMode === 'months') {
        cols = getAllMonths();
        getCol = (e) => formatMonth(e.date) || 'ללא תאריך';
        if (!cols.includes('ללא תאריך') && entries.some(e => !formatMonth(e.date))) cols.push('ללא תאריך');
    } else {
        // employees mode
        if (groupEmployees && Object.keys(employeeGroups).length > 0) {
            const empGroupMap = {};
            Object.entries(employeeGroups).forEach(([gName, members]) => {
                members.forEach(m => { empGroupMap[m] = gName; });
            });
            const colSet = new Set();
            getAllEmployees().forEach(emp => {
                if (empGroupMap[emp]) colSet.add(empGroupMap[emp]);
                else if (ungroupedMode === 'individual') colSet.add(emp);
                else colSet.add('אחר');
            });
            cols = [...colSet].sort();
            getCol = (e) => empGroupMap[e.employee] || (ungroupedMode === 'individual' ? e.employee : 'אחר');
        } else {
            cols = getAllEmployees().filter(emp => selectedEmployees.has(emp));
            getCol = (e) => e.employee || 'ללא עובד';
            if (!cols.includes('ללא עובד') && entries.some(e => !e.employee)) cols.push('ללא עובד');
        }
    }

    // --- Build ROWS ---
    let rowKeys = [];
    let getRow; // function: entry -> row label
    let rowLabel; // function: row key -> display label

    if (caseGroupMode === 'client') {
        // Group by client name
        const allClients = getAllClients();
        // Filter by selected clients
        rowKeys = allClients.filter(c => selectedCaseGroups.has(c));
        getRow = (e) => e.client || 'ללא לקוח';
        rowLabel = (r) => r;
    } else if (caseGroupMode === 'groups') {
        const caseGroupMap = {};
        Object.entries(caseGroups).forEach(([gName, members]) => {
            members.forEach(m => { caseGroupMap[m] = gName; });
        });
        const rowSet = new Set();
        getAllCases().forEach(c => {
            if (caseGroupMap[c.key]) rowSet.add(caseGroupMap[c.key]);
            else if (ungroupedMode === 'individual') rowSet.add(c.key);
            else rowSet.add('אחר');
        });
        // Filter by selected case groups; individual ungrouped cases always pass through
        const groupNames = new Set(Object.keys(caseGroups));
        rowKeys = [...rowSet].filter(r => {
            if (groupNames.has(r) || r === 'אחר') return selectedCaseGroups.has(r);
            return true; // individual ungrouped case — always include
        }).sort();
        getRow = (e) => caseGroupMap[e.caseKey] || (ungroupedMode === 'individual' ? e.caseKey : 'אחר');
        rowLabel = (r) => {
            if (groupNames.has(r) || r === 'אחר') return r;
            return caseLabel(r);
        };
    } else {
        // caseGroupMode === 'none' — individual cases, filtered by caseFilterMode
        const allCases = getAllCases();
        let filteredKeys;
        if (caseFilterMode === 'case') {
            filteredKeys = allCases.map(c => c.key).filter(k => selectedCaseGroups.has(k));
        } else if (caseFilterMode === 'client') {
            const caseClientMap = {};
            allCases.forEach(c => { caseClientMap[c.key] = c.client; });
            filteredKeys = allCases.map(c => c.key).filter(k => selectedCaseGroups.has(caseClientMap[k]));
        } else { // 'groups'
            const cgMap = {};
            Object.entries(caseGroups).forEach(([g, members]) => members.forEach(m => { cgMap[m] = g; }));
            filteredKeys = allCases.map(c => c.key).filter(k => selectedCaseGroups.has(cgMap[k] || 'אחר'));
        }
        rowKeys = filteredKeys;
        getRow = (e) => e.caseKey;
        rowLabel = (r) => caseLabel(r);
    }

    // --- Build pivot data ---
    const pivotData = {};
    const rowTotals = {};
    const colTotals = {};
    let grandTotal = 0;

    rowKeys.forEach(r => { pivotData[r] = {}; rowTotals[r] = 0; });
    cols.forEach(c => { colTotals[c] = 0; });

    // --- Build employee group map for sub bar charts (months mode) ---
    let empGMap = null;
    let empGroupLabels = [];
    let subBarData = {};
    const rowKeySet = new Set(rowKeys);
    if (colMode === 'months' && Object.keys(employeeGroups).length > 0) {
        empGMap = {};
        Object.entries(employeeGroups).forEach(([gName, members]) => {
            members.forEach(m => { empGMap[m] = gName; });
        });
        const egSet = new Set();
        getAllEmployees().forEach(emp => {
            if (empGMap[emp]) egSet.add(empGMap[emp]);
            else egSet.add('אחר');
        });
        empGroupLabels = [...egSet].sort();
        rowKeys.forEach(r => {
            subBarData[r] = {};
            empGroupLabels.forEach(eg => { subBarData[r][eg] = {}; });
        });
    } else if (colMode === 'months') {
        // No employee groups defined — treat each individual employee as their own group
        empGMap = {};
        getAllEmployees().forEach(emp => { empGMap[emp] = emp; });
        empGroupLabels = getAllEmployees().slice().sort();
        rowKeys.forEach(r => {
            subBarData[r] = {};
            empGroupLabels.forEach(eg => { subBarData[r][eg] = {}; });
        });
    }

    // --- Single pass: build pivot data + subBarData ---
    for (const e of entries) {
        const col = getCol(e);
        const row = getRow(e);
        if (!rowKeySet.has(row)) continue;
        const val = e[hourKey];
        pivotData[row][col] = (pivotData[row][col] || 0) + val;
        rowTotals[row] = (rowTotals[row] || 0) + val;
        colTotals[col] = (colTotals[col] || 0) + val;
        grandTotal += val;
        if (colMode === 'months' && empGMap) {
            const month = formatMonth(e.date) || 'ללא תאריך';
            const eg = empGMap[e.employee] || 'אחר';
            if (!subBarData[row][eg]) subBarData[row][eg] = {};
            subBarData[row][eg][month] = (subBarData[row][eg][month] || 0) + val;
        }
    }

    // --- Drop empty rows and columns ---
    rowKeys = rowKeys.filter(r => (rowTotals[r] || 0) > 0);
    cols = cols.filter(c => (colTotals[c] || 0) > 0);

    // --- Render table (synchronous — user sees this first) ---
    const rowTypeLabel = caseGroupMode === 'client' ? 'לקוח' : (caseGroupMode === 'groups' && Object.keys(caseGroups).length > 0 ? 'קבוצה' : 'תיק');
    const cornerLabel = colMode === 'months' ? `${rowTypeLabel} / חודש` : `${rowTypeLabel} / עובד`;
    const thead = $('#pivot-table thead');
    const tbody = $('#pivot-table tbody');

    // --- Sort rowKeys ---
    if (sortCol !== null) {
        rowKeys.sort((a, b) => {
            let va, vb;
            if (sortCol === '__total__') {
                va = rowTotals[a] || 0;
                vb = rowTotals[b] || 0;
            } else {
                va = pivotData[a]?.[sortCol] || 0;
                vb = pivotData[b]?.[sortCol] || 0;
            }
            return sortDir === 'desc' ? vb - va : va - vb;
        });
    }

    const sortIndicator = (key) => {
        if (sortCol === key) return sortDir === 'desc' ? ' ▼' : ' ▲';
        return '<span class="sort-idle">⇅</span>';
    };

    thead.innerHTML = `<tr>
        <th class="pivot-corner">${cornerLabel}</th>
        ${cols.map(c => `<th class="pivot-sortable" data-col="${escData(c)}">${esc(c)}${sortIndicator(c)}</th>`).join('')}
        <th class="pivot-sortable pivot-total-header" data-col="__total__">סה"כ${sortIndicator('__total__')}</th>
    </tr>`;

    const tbodyParts = [];
    for (const r of rowKeys) {
        const label = rowLabel(r);
        tbodyParts.push(`<tr><td><strong>${esc(label)}</strong></td>`);
        for (const c of cols) {
            const v = pivotData[r]?.[c] || 0;
            tbodyParts.push(`<td class="pivot-value">${v ? v.toFixed(2) : '-'}</td>`);
        }
        tbodyParts.push(`<td class="pivot-value pivot-total">${(rowTotals[r] || 0).toFixed(2)}</td></tr>`);
    }
    tbodyParts.push(`<tr class="pivot-total-row"><td><strong>סה"כ</strong></td>`);
    for (const c of cols) {
        tbodyParts.push(`<td class="pivot-value">${(colTotals[c] || 0).toFixed(2)}</td>`);
    }
    tbodyParts.push(`<td class="pivot-value pivot-total">${grandTotal.toFixed(2)}</td></tr>`);
    tbody.innerHTML = tbodyParts.join('');

    // --- Defer chart rendering until after table paints ---
    requestAnimationFrame(() => {
        renderPivotChart(cols, rowKeys, rowLabel, pivotData, colTotals, empGroupLabels, subBarData, getCol);
    });
}

// ============================================================
// Pivot Chart
// ============================================================
const CHART_COLORS = [
    '#1a2a4a', '#b8964e', '#2c4066', '#d4b06a', '#4a7c9b', '#8b6914',
    '#3d6b8e', '#c4a25c', '#1f3d5e', '#9ba3b0', '#5a8aad', '#7a6322',
    '#6b9dc2', '#a0834a', '#345878', '#c8ccd4', '#4e8fb5', '#b39240',
    '#2a5a7d', '#dcc07a', '#1e4d6e', '#e6c878', '#3a7099', '#8c7530'
];

// Distinct palette for the cases/case-groups bar chart (greens/teals/warm tones)
const CASE_COLORS = [
    '#2e7d5f', '#c26a3d', '#3a8a8c', '#a85c90', '#6a9e3b', '#d4883e',
    '#4c8eaf', '#9b6b4a', '#5aad7a', '#b5555a', '#3d7a6e', '#c49240',
    '#7a6aad', '#5e9960', '#d47a6a', '#488080', '#b08a30', '#6a5e8a',
    '#80b050', '#c07070', '#3a9070', '#d0a050', '#5a7ab0', '#a06a50'
];

function renderPivotChart(cols, rowKeys, rowLabel, pivotData, colTotals, empGroupLabels, subBarData, getCol) {
    const canvas = $('#pivot-chart');
    if (!canvas) return;

    // Destroy existing charts
    if (pivotChart) { pivotChart.destroy(); pivotChart = null; }
    if (totalsChart) { totalsChart.destroy(); totalsChart = null; }
    subCharts.forEach(c => c.destroy());
    subCharts = [];
    $('#sub-charts-area').innerHTML = '';
    $('#totals-chart-area').classList.add('hidden');
    // Reset main canvas visibility (may have been hidden by fallback table)
    canvas.style.display = '';
    // Remove any previous fallback tables from chart area
    canvas.parentElement.querySelectorAll('.pie-fallback-table, .pie-others-expand').forEach(el => el.remove());

    // Build consistent color map for columns (employees)
    const colColorMap = {};
    cols.forEach((c, i) => { colColorMap[c] = CHART_COLORS[i % CHART_COLORS.length]; });

    // Build consistent color map for rows (cases/case groups)
    const rowColorMap = {};
    rowKeys.forEach((r, i) => { rowColorMap[r] = CASE_COLORS[i % CASE_COLORS.length]; });

    // Show/hide toggles
    const subChartModeArea = $('#sub-chart-mode-area');
    const chartDistModeArea = $('#chart-dist-mode-area');

    if (colMode === 'employees') {
        // Show chart distribution toggle
        if (chartDistModeArea) {
            chartDistModeArea.classList.remove('hidden');
            chartDistModeArea.querySelectorAll('.toggle-btn').forEach(b => {
                b.classList.toggle('active', b.dataset.value === chartDistMode);
            });
        }

        const entries = getFilteredEntries();
        const hourKey = valueMode === 'billable' ? 'billableHours' : 'workHours';
        const clients = getAllClients();
        const hasCaseGroups = Object.keys(caseGroups).length > 0;

        if (chartDistMode === 'cases') {
            // ============ Cases distribution mode ============
            // Hide sub-chart mode toggle (not relevant here)
            if (subChartModeArea) subChartModeArea.classList.add('hidden');

            // Main pie: share of each case/case group (matches pivot rows)
            const rowTotals = {};
            rowKeys.forEach(r => {
                rowTotals[r] = cols.reduce((sum, c) => sum + (pivotData[r]?.[c] || 0), 0);
            });
            const mainLabels = rowKeys.map(r => rowLabel(r));
            const mainData = rowKeys.map(r => rowTotals[r] || 0);
            const mainColors = rowKeys.map(r => rowColorMap[r]);

            pivotChart = renderSmartPie(canvas, mainLabels, mainData, mainColors, 'right', 12, canvas.parentElement);

            // Sub pies: one per column (employee/employee group), slices = cases/case groups
            const subPivotData = {};
            cols.forEach(c => {
                subPivotData[c] = {};
                rowKeys.forEach(r => {
                    subPivotData[c][r] = pivotData[r]?.[c] || 0;
                });
            });

            // Render sub charts with row colors for consistency
            renderSubCharts(rowKeys, cols, (c) => c, subPivotData, rowColorMap, 'pie', null, (r) => rowLabel(r));
        } else {
            // ============ Employees distribution mode (original) ============
            if (!subChartModeManual) {
                if (clients.length > 1) subChartMode = 'client';
                else if (hasCaseGroups) subChartMode = 'casegroup';
                else subChartMode = 'case';
            }
            if (subChartModeArea) {
                subChartModeArea.classList.remove('hidden');
                const cgBtn = subChartModeArea.querySelector('[data-value="casegroup"]');
                if (cgBtn) {
                    cgBtn.disabled = !hasCaseGroups;
                    cgBtn.style.opacity = hasCaseGroups ? '' : '0.4';
                }
                subChartModeArea.querySelectorAll('.toggle-btn').forEach(b => {
                    b.classList.toggle('active', b.dataset.value === subChartMode);
                });
            }

            // Main pie: share of each employee
            const allLabels = cols;
            const allColors = cols.map(c => colColorMap[c]);
            const mainData = cols.map(c => colTotals[c] || 0);

            pivotChart = renderSmartPie(canvas, allLabels, mainData, allColors, 'right', 12, canvas.parentElement);

            // Build sub-chart data based on subChartMode
            let subRowKeys, subGetRow, subRowLabel2, subPivotData;

            if (subChartMode === 'client') {
                subRowKeys = clients;
                subGetRow = (e) => e.client || 'ללא לקוח';
                subRowLabel2 = (r) => r;
            } else if (subChartMode === 'casegroup' && hasCaseGroups) {
                const caseGroupMap = {};
                Object.entries(caseGroups).forEach(([gName, members]) => {
                    members.forEach(m => { caseGroupMap[m] = gName; });
                });
                const rowSet = new Set();
                getAllCases().forEach(c => {
                    if (caseGroupMap[c.key]) rowSet.add(caseGroupMap[c.key]);
                    else rowSet.add('אחר');
                });
                subRowKeys = [...rowSet].sort();
                subGetRow = (e) => caseGroupMap[e.caseKey] || 'אחר';
                subRowLabel2 = (r) => r;
            } else {
                const allCases = getAllCases();
                subRowKeys = allCases.map(c => c.key);
                subGetRow = (e) => e.caseKey;
                subRowLabel2 = (r) => caseLabel(r);
            }

            subPivotData = {};
            subRowKeys.forEach(r => { subPivotData[r] = {}; });
            entries.forEach(e => {
                const row = subGetRow(e);
                if (!subPivotData[row]) return;
                const col = getCol(e);
                const val = e[hourKey];
                subPivotData[row][col] = (subPivotData[row][col] || 0) + val;
            });

            renderSubCharts(cols, subRowKeys, subRowLabel2, subPivotData, colColorMap, 'pie');
        }
    } else {
        // Hide toggles in months mode
        if (subChartModeArea) subChartModeArea.classList.add('hidden');
        if (chartDistModeArea) chartDistModeArea.classList.add('hidden');

        // Main bar chart (uses CASE_COLORS to distinguish from employee-group charts below)
        // Apply top-N grouping + smart stacked/grouped threshold
        const TOP_N = 8;
        const M = cols.length;
        let chartRowKeys = rowKeys;
        let othersDataset = null;

        if (rowKeys.length > TOP_N) {
            // Rank rows by total hours descending, keep top N
            const rowTotals = {};
            rowKeys.forEach(r => { rowTotals[r] = cols.reduce((s, c) => s + (pivotData[r]?.[c] || 0), 0); });
            const sorted = [...rowKeys].sort((a, b) => (rowTotals[b] || 0) - (rowTotals[a] || 0));
            chartRowKeys = sorted.slice(0, TOP_N);
            const othersRows = sorted.slice(TOP_N);
            // Build "אחרים" dataset
            const othersData = cols.map(c => othersRows.reduce((s, r) => s + (pivotData[r]?.[c] || 0), 0));
            othersDataset = { label: 'אחרים', data: othersData, backgroundColor: OTHERS_COLOR, borderColor: OTHERS_COLOR, borderWidth: 1, borderRadius: 2 };
        }

        const S_eff = chartRowKeys.length + (othersDataset ? 1 : 0);
        const useStacked = S_eff * M > 72;

        const datasets = chartRowKeys.map((r, i) => {
            const color = CASE_COLORS[i % CASE_COLORS.length];
            return { label: rowLabel(r), data: cols.map(c => pivotData[r]?.[c] || 0), backgroundColor: color, borderColor: color, borderWidth: 1, borderRadius: 2 };
        });
        if (othersDataset) datasets.push(othersDataset);

        pivotChart = new Chart(canvas, {
            type: 'bar',
            data: { labels: cols, datasets },
            options: barOptions(useStacked)
        });

        // Totals stacked bar chart — stacked by employee group across all cases
        const totalsCanvas = $('#totals-chart');
        $('#totals-chart-area').classList.remove('hidden');
        const egColorMap = {};
        empGroupLabels.forEach((eg, i) => { egColorMap[eg] = CHART_COLORS[i % CHART_COLORS.length]; });

        // Aggregate subBarData across all rows: { empGroup: { month: hours } }
        const totalsByEg = {};
        empGroupLabels.forEach(eg => { totalsByEg[eg] = {}; });
        Object.values(subBarData).forEach(egMap => {
            Object.entries(egMap).forEach(([eg, months]) => {
                Object.entries(months).forEach(([month, hrs]) => {
                    if (!totalsByEg[eg]) totalsByEg[eg] = {};
                    totalsByEg[eg][month] = (totalsByEg[eg][month] || 0) + hrs;
                });
            });
        });

        const totalsDatasets = empGroupLabels.map(eg => ({
            label: eg,
            data: cols.map(c => totalsByEg[eg]?.[c] || 0),
            backgroundColor: egColorMap[eg],
            borderColor: egColorMap[eg],
            borderWidth: 1,
            borderRadius: 2
        }));
        totalsChart = new Chart(totalsCanvas, {
            type: 'bar',
            data: { labels: cols, datasets: totalsDatasets },
            options: barOptions(true)
        });
        renderSubCharts(cols, rowKeys, rowLabel, subBarData, egColorMap, 'bar', empGroupLabels);
    }
}

const OTHERS_COLOR = '#9ba3b0';

function renderSmartPie(canvasEl, labels, data, colors, legendPos, fontSize, containerEl) {
    const total = data.reduce((a, b) => a + b, 0);
    if (total === 0) return null;

    // Pair up and remove zeros
    const items = labels.map((l, i) => ({ label: l, value: data[i], color: colors[i] }))
        .filter(it => it.value > 0)
        .sort((a, b) => b.value - a.value);

    const hasLarge = items.some(it => (it.value / total) * 100 >= 5);

    if (!hasLarge) {
        // Scenario 1: all small — draw ranked bar table instead of pie
        canvasEl.style.display = 'none';
        const tableDiv = document.createElement('div');
        tableDiv.className = 'pie-fallback-table';
        const maxVal = items[0]?.value || 1;
        tableDiv.innerHTML = items.map(it => {
            const pct = ((it.value / total) * 100).toFixed(1);
            const barW = Math.round((it.value / maxVal) * 100);
            return `<div class="pft-row">
                <div class="pft-color" style="background:${it.color}"></div>
                <div class="pft-label">${esc(it.label)}</div>
                <div class="pft-bar-wrap"><div class="pft-bar" style="width:${barW}%;background:${it.color}"></div></div>
                <div class="pft-pct">${pct}%</div>
                <div class="pft-val">${it.value.toFixed(2)}</div>
            </div>`;
        }).join('');
        containerEl.appendChild(tableDiv);
        return null;
    }

    // Scenario 2: some items ≥5% — group items <5% into "אחר"
    const large = items.filter(it => (it.value / total) * 100 >= 5);
    const small = items.filter(it => (it.value / total) * 100 < 5);

    let chartLabels, chartData, chartColors, smallGroup = null;

    if (small.length > 0) {
        const othersTotal = small.reduce((s, it) => s + it.value, 0);
        chartLabels = large.map(it => it.label).concat(['אחרים']);
        chartData = large.map(it => it.value).concat([othersTotal]);
        chartColors = large.map(it => it.color).concat([OTHERS_COLOR]);
        smallGroup = { items: small, total: othersTotal };
    } else {
        chartLabels = large.map(it => it.label);
        chartData = large.map(it => it.value);
        chartColors = large.map(it => it.color);
    }

    canvasEl.style.display = '';
    const chart = new Chart(canvasEl, {
        type: 'pie',
        data: {
            labels: chartLabels,
            datasets: [{ data: chartData, backgroundColor: chartColors, borderColor: '#fff', borderWidth: 1.5 }]
        },
        options: pieOptions(legendPos, fontSize)
    });

    // Add expandable "אחרים" section if needed
    if (smallGroup) {
        const expandDiv = document.createElement('div');
        expandDiv.className = 'pie-others-expand';
        const othPct = ((smallGroup.total / total) * 100).toFixed(1);
        expandDiv.innerHTML = `<button class="pie-others-btn">▶ אחרים (${othPct}%) — לחץ להרחבה</button><div class="pie-others-list hidden"></div>`;

        const btn = expandDiv.querySelector('.pie-others-btn');
        const listDiv = expandDiv.querySelector('.pie-others-list');
        const allOthers = smallGroup.items;
        const showItems = allOthers.slice(0, 15);
        const overflow = allOthers.length - showItems.length;
        listDiv.innerHTML = showItems.map(it => {
            const pct = ((it.value / total) * 100).toFixed(1);
            return `<div class="poi-row"><span class="poi-dot" style="background:${it.color}"></span><span class="poi-label">${esc(it.label)}</span><span class="poi-pct">${pct}%</span><span class="poi-val">${it.value.toFixed(2)}</span></div>`;
        }).join('') + (overflow > 0 ? `<div class="poi-overflow">ועוד ${overflow} פריטים נוספים</div>` : '');

        btn.addEventListener('click', () => {
            const open = !listDiv.classList.contains('hidden');
            listDiv.classList.toggle('hidden', open);
            btn.textContent = (open ? '▶' : '▼') + btn.textContent.slice(1);
        });

        containerEl.appendChild(expandDiv);
    }

    return chart;
}

function pieOptions(legendPos, fontSize) {
    return {
        responsive: true, maintainAspectRatio: false,
        plugins: {
            legend: {
                position: legendPos, rtl: true, textDirection: 'rtl',
                labels: {
                    font: { family: "'Assistant', sans-serif", size: fontSize },
                    padding: legendPos === 'right' ? 12 : 6,
                    usePointStyle: true, pointStyleWidth: legendPos === 'right' ? 16 : 10,
                    boxWidth: legendPos === 'right' ? 16 : 10,
                    filter: (item, chart) => {
                        const val = chart.datasets[0].data[item.index];
                        return val > 0;
                    },
                    generateLabels: (chart) => {
                        const data = chart.data;
                        const total = data.datasets[0].data.reduce((a, b) => a + b, 0);
                        return data.labels.map((label, i) => {
                            const val = data.datasets[0].data[i];
                            const pct = total > 0 ? ((val / total) * 100).toFixed(1) : '0.0';
                            return {
                                text: `${label} (${pct}%)`,
                                fillStyle: data.datasets[0].backgroundColor[i],
                                strokeStyle: data.datasets[0].borderColor || '#fff',
                                lineWidth: data.datasets[0].borderWidth || 1,
                                hidden: false,
                                index: i,
                                pointStyle: 'circle'
                            };
                        });
                    }
                }
            },
            tooltip: {
                rtl: true, textDirection: 'rtl',
                callbacks: {
                    label: function(ctx) {
                        const val = ctx.parsed;
                        const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                        const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                        return `${ctx.label}: ${val.toFixed(2)} (${pct}%)`;
                    }
                }
            },
            datalabels: {
                color: '#fff',
                font: { family: "'Assistant', sans-serif", size: fontSize, weight: 'bold' },
                formatter: (value, ctx) => {
                    const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                    const pct = total > 0 ? (value / total) * 100 : 0;
                    return pct >= 5 ? pct.toFixed(0) + '%' : '';
                },
                display: (ctx) => {
                    const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                    const pct = total > 0 ? (ctx.dataset.data[ctx.dataIndex] / total) * 100 : 0;
                    return pct >= 5;
                }
            }
        }
    };
}

function barOptions(stacked) {
    return {
        responsive: true, maintainAspectRatio: false,
        scales: {
            x: {
                stacked: stacked,
                ticks: { font: { family: "'Assistant', sans-serif", size: 11 } },
                grid: { display: false }
            },
            y: {
                stacked: stacked,
                beginAtZero: true,
                ticks: { font: { family: "'Assistant', sans-serif", size: 11 } },
                title: { display: true, text: valueMode === 'billable' ? 'שעות חיוב' : 'שעות עבודה', font: { family: "'Assistant', sans-serif", size: 13, weight: '600' } }
            }
        },
        plugins: {
            legend: {
                position: 'bottom', rtl: true, textDirection: 'rtl',
                labels: { font: { family: "'Assistant', sans-serif", size: 11 }, padding: 10, usePointStyle: true, pointStyleWidth: 12, boxWidth: 12 }
            },
            tooltip: {
                rtl: true, textDirection: 'rtl',
                callbacks: { label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toFixed(2)}` }
            },
            datalabels: { display: false }
        }
    };
}

function reRenderSubCharts() {
    if (!_lastSubChartArgs) return;
    subCharts.forEach(c => c.destroy());
    subCharts = [];
    $('#sub-charts-area').innerHTML = '';
    renderSubCharts(..._lastSubChartArgs);
}

function rebuildSubChartFilter(validRows, rowLabelFn) {
    const area = $('#sub-chart-filter-area');
    if (!area) return;

    if (!validRows || validRows.length < 2) {
        area.classList.add('hidden');
        return;
    }
    area.classList.remove('hidden');

    const allKeys = validRows.map(r => String(r));
    const newKeySet = new Set(allKeys);

    // Detect whether the available items changed (new data/mode)
    const prevKeys = _subChartFilterAllItems.map(i => i.key);
    const keysChanged = allKeys.length !== prevKeys.length || allKeys.some((k, i) => k !== prevKeys[i]);

    _subChartFilterAllItems = allKeys.map(k => ({ key: k, label: rowLabelFn(validRows.find(r => String(r) === k) ?? k) }));

    // Remove keys no longer valid
    selectedSubCharts = new Set([...selectedSubCharts].filter(k => newKeySet.has(k)));
    // Only reset to "all selected" when items change — preserve user's deliberate deselection
    if (keysChanged) selectedSubCharts = new Set(allKeys);

    renderSubChartFilterList(_subChartFilterAllItems, '');
    updateSubChartFilterTrigger(_subChartFilterAllItems);
}

function renderSubChartFilterList(allItems, searchTerm) {
    const list = $('#sub-chart-filter-list');
    if (!list) return;
    const term = searchTerm.toLowerCase();
    const filtered = term ? allItems.filter(i => i.label.toLowerCase().includes(term)) : allItems;

    const allChecked = allItems.length > 0 && allItems.every(i => selectedSubCharts.has(i.key));
    const someChecked = !allChecked && allItems.some(i => selectedSubCharts.has(i.key));

    list.innerHTML = '';

    const allLi = document.createElement('li');
    allLi.className = 'xf-item xf-select-all';
    allLi.innerHTML = `<label><input type="checkbox" /><span>בחר הכל</span></label>`;
    const allCb = allLi.querySelector('input');
    allCb.checked = allChecked;
    allCb.indeterminate = someChecked;
    allCb.addEventListener('change', () => {
        if (allCb.checked) allItems.forEach(i => selectedSubCharts.add(i.key));
        else allItems.forEach(i => selectedSubCharts.delete(i.key));
        renderSubChartFilterList(allItems, $('#sub-chart-filter-search').value);
        updateSubChartFilterTrigger(allItems);
        reRenderSubCharts();
    });
    list.appendChild(allLi);

    filtered.forEach(item => {
        const li = document.createElement('li');
        li.className = 'xf-item';
        li.innerHTML = `<label><input type="checkbox" /><span>${esc(item.label)}</span></label>`;
        const cb = li.querySelector('input');
        cb.checked = selectedSubCharts.has(item.key);
        cb.addEventListener('change', () => {
            if (cb.checked) selectedSubCharts.add(item.key);
            else selectedSubCharts.delete(item.key);
            renderSubChartFilterList(allItems, $('#sub-chart-filter-search').value);
            updateSubChartFilterTrigger(allItems);
            reRenderSubCharts();
        });
        list.appendChild(li);
    });
}

function updateSubChartFilterTrigger(allItems) {
    const trigger = $('#sub-chart-filter-trigger');
    if (!trigger) return;
    const selected = allItems.filter(i => selectedSubCharts.has(i.key)).length;
    trigger.textContent = selected === allItems.length ? 'הכל ▾' : `${selected} מתוך ${allItems.length} ▾`;
}

function renderSubCharts(cols, rowKeys, rowLabel, dataMap, colorMap, chartType, empGroupLabels, colLabelFn) {
    _lastSubChartArgs = [cols, rowKeys, rowLabel, dataMap, colorMap, chartType, empGroupLabels, colLabelFn];
    const container = $('#sub-charts-area');
    if (!container || rowKeys.length === 0) return;
    const getColLabel = colLabelFn || ((c) => c);

    // All rows with data
    const allValidRows = rowKeys.filter(r => {
        if (chartType === 'pie') {
            return cols.some(c => (dataMap[r]?.[c] || 0) > 0);
        } else {
            return empGroupLabels.some(eg => cols.some(c => (dataMap[r]?.[eg]?.[c] || 0) > 0));
        }
    });

    rebuildSubChartFilter(allValidRows, rowLabel);

    // Apply sub-chart filter
    const validRows = allValidRows.filter(r => selectedSubCharts.has(String(r)));

    const BATCH = 4;
    let shown = 0;

    function showNextBatch() {
        const batch = validRows.slice(shown, shown + BATCH);
        batch.forEach(r => {
            const label = rowLabel(r);
            const card = document.createElement('div');
            card.className = 'sub-chart-card';
            card.dataset.rowKey = String(r);
            card.innerHTML = `<h4>${esc(label)}</h4><canvas></canvas>`;
            // Insert before the "show more" button if it exists
            const moreBtn = container.querySelector('.show-more-btn');
            if (moreBtn) container.insertBefore(card, moreBtn);
            else container.appendChild(card);

            const subCanvas = card.querySelector('canvas');

            if (chartType === 'pie') {
                const data = cols.map(c => dataMap[r]?.[c] || 0);
                const colors = cols.map(c => colorMap[c]);
                const labels = cols.map(c => getColLabel(c));
                const chart = renderSmartPie(subCanvas, labels, data, colors, 'bottom', 10, card);
                if (chart) subCharts.push(chart);
            } else {
                // Stacked bar: months on X, one dataset per employee group
                const datasets = empGroupLabels.map(eg => ({
                    label: eg,
                    data: cols.map(c => dataMap[r]?.[eg]?.[c] || 0),
                    backgroundColor: colorMap[eg],
                    borderColor: colorMap[eg],
                    borderWidth: 1
                })).filter(ds => ds.data.some(v => v > 0));

                const chart = new Chart(subCanvas, {
                    type: 'bar',
                    data: { labels: cols, datasets },
                    options: barOptions(true)
                });
                subCharts.push(chart);
            }
        });

        shown += batch.length;

        // Remove old button
        const oldBtn = container.querySelector('.show-more-btn');
        if (oldBtn) oldBtn.remove();

        // Add "show more" button if there are more
        if (shown < validRows.length) {
            const btn = document.createElement('button');
            btn.className = 'btn show-more-btn';
            btn.textContent = 'הצג גרפים נוספים...';
            btn.addEventListener('click', showNextBatch);
            container.appendChild(btn);
        }
    }

    showNextBatch();
}

$('#download-pivot').addEventListener('click', () => {
    const table = $('#pivot-table');
    const wb = XLSX.utils.table_to_book(table, { sheet: 'טבלת ציר' });
    XLSX.writeFile(wb, 'טבלת_ציר.xlsx');
});

// ============================================================
// PDF Report Generation
// ============================================================

$('#download-pdf').addEventListener('click', async () => {
    const entries = getFilteredEntries();
    if (!entries.length) { alert('אין נתונים להצגה'); return; }

    // Build suggested title and let user edit it
    const fromVal = $('#date-from').value;
    const toVal = $('#date-to').value;
    let suggestedDateRange = '';
    if (fromVal && toVal) suggestedDateRange = `${formatDateHebrew(new Date(fromVal + 'T00:00:00'))} - ${formatDateHebrew(new Date(toVal + 'T00:00:00'))}`;
    else if (fromVal) suggestedDateRange = `${formatDateHebrew(new Date(fromVal + 'T00:00:00'))} ואילך`;
    else if (toVal) suggestedDateRange = `עד ${formatDateHebrew(new Date(toVal + 'T00:00:00'))}`;
    else {
        const dates = entries.filter(e => e.date instanceof Date && !isNaN(e.date.getTime())).map(e => e.date);
        if (dates.length) {
            const ts = dates.map(d => d.getTime());
            const minD = new Date(ts.reduce((a, b) => a < b ? a : b));
            const maxD = new Date(ts.reduce((a, b) => a > b ? a : b));
            suggestedDateRange = `${formatDateHebrew(minD)} - ${formatDateHebrew(maxD)}`;
        }
    }
    const suggestedClients = [...new Set(entries.map(e => e.client).filter(Boolean))].sort();
    let suggestedClientStr = suggestedClients.slice(0, 3).join(', ');
    if (suggestedClients.length > 3) suggestedClientStr += ' ...';
    const suggestedParts = ['סיכום דוח שעות'];
    if (suggestedDateRange) suggestedParts.push(suggestedDateRange);
    if (suggestedClientStr) suggestedParts.push(suggestedClientStr);
    const suggestedTitle = suggestedParts.join(' || ');

    const userTitle = prompt('כותרת הדוח:', suggestedTitle);
    if (userTitle === null) return; // user cancelled

    const btn = $('#download-pdf');
    btn.disabled = true;
    btn.textContent = '...מייצר דוח';

    try {
        await generatePdfReport(entries, userTitle.trim() || suggestedTitle);
    } catch (e) {
        console.error('PDF generation error:', e);
        alert('שגיאה ביצירת הדוח: ' + e.message);
    } finally {
        btn.disabled = false;
        btn.textContent = 'דוח מודפס';
    }
});

async function captureElement(el, scale) {
    // Use html2canvas to capture a DOM element as an image
    const canvas = await html2canvas(el, {
        scale: scale || 2,
        useCORS: true,
        backgroundColor: '#ffffff',
    });
    return canvas;
}

async function generatePdfReport(entries, reportTitle) {
    const { jsPDF } = window.jspdf;
    // Track temp DOM elements for cleanup on error
    const _tempElements = [];
    function _addTemp(el) { document.body.appendChild(el); _tempElements.push(el); return el; }
    function _removeTemp(el) { if (el.parentNode) el.parentNode.removeChild(el); const idx = _tempElements.indexOf(el); if (idx >= 0) _tempElements.splice(idx, 1); }
    function _cleanupAllTemp() { _tempElements.forEach(el => { if (el.parentNode) el.parentNode.removeChild(el); }); _tempElements.length = 0; }

    try {

    // --- Determine orientation based on pivot table width ---
    const pivotTable = $('#pivot-table');
    const colCount = pivotTable.querySelectorAll('thead th').length;
    const orientation = colCount > 7 ? 'landscape' : 'portrait';

    const doc = new jsPDF({ orientation, unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 12;
    const contentW = pageW - margin * 2;
    let curY = margin;

    function ensureSpace(needed) {
        if (curY + needed > pageH - margin - 10) {
            doc.addPage();
            curY = margin;
        }
    }

    // --- Helper: add an image from a canvas, fitting within given width/maxHeight ---
    function addImage(imgCanvas, widthMm, maxHeightMm) {
        const ratio = imgCanvas.height / imgCanvas.width;
        let imgW = widthMm;
        let imgH = imgW * ratio;
        if (imgH > maxHeightMm) {
            imgH = maxHeightMm;
            imgW = imgH / ratio;
        }
        const imgData = imgCanvas.toDataURL('image/png');
        const xOffset = margin + (contentW - imgW) / 2;
        doc.addImage(imgData, 'PNG', xOffset, curY, imgW, imgH);
        curY += imgH + 3;
    }

    // --- Title (rendered as a temporary DOM element, captured as image) ---
    const titleText = reportTitle;

    const titleEl = document.createElement('div');
    titleEl.style.cssText = `
        font-family: 'Assistant', sans-serif; font-size: 18px; font-weight: 700;
        color: #1a2a4a; text-align: center; padding: 8px 16px;
        direction: rtl; background: white; width: ${Math.round(contentW * 3.78)}px;
    `;
    titleEl.textContent = titleText;
    _addTemp(titleEl);
    const titleCanvas = await captureElement(titleEl, 2);
    _removeTemp(titleEl);
    addImage(titleCanvas, contentW, 20);

    curY += 2;

    // --- Pivot Table (rebuild as flat table to avoid thead/tbody rendering issues) ---
    const availableH = pageH - curY - margin - 10;
    const targetPxW = Math.round(contentW * 3.78);

    // Extract all rows: header row first, then body rows
    const headerCells = [...pivotTable.querySelectorAll('thead th')].map(th => th.textContent.trim());
    const bodyRowsData = [...pivotTable.querySelectorAll('tbody tr')].map(tr =>
        [...tr.querySelectorAll('td')].map(td => td.textContent.trim())
    );

    // Build a plain table (no thead/tbody) to avoid rendering order issues
    function buildPdfTable(fontSize) {
        const tbl = document.createElement('table');
        tbl.style.cssText = `
            width: 100%; border-collapse: collapse;
            font-family: 'Assistant', sans-serif; font-size: ${fontSize}px;
            direction: rtl;
        `;
        const cellStyle = `padding: 3px 5px; border: 1px solid #ccc; text-align: center; white-space: nowrap; font-size: ${fontSize}px;`;

        // Header row
        const hRow = document.createElement('tr');
        headerCells.forEach(text => {
            const td = document.createElement('td');
            td.textContent = text;
            td.style.cssText = cellStyle + 'background-color: #1a2a4a; color: #fff; font-weight: bold;';
            hRow.appendChild(td);
        });
        tbl.appendChild(hRow);

        // Body rows
        bodyRowsData.forEach((cells, rowIdx) => {
            const tr = document.createElement('tr');
            const isLastRow = rowIdx === bodyRowsData.length - 1;
            cells.forEach(text => {
                const td = document.createElement('td');
                td.textContent = text;
                td.style.cssText = cellStyle;
                if (isLastRow) {
                    td.style.fontWeight = 'bold';
                    td.style.backgroundColor = '#f0f0f5';
                }
                tr.appendChild(td);
            });
            tbl.appendChild(tr);
        });

        return tbl;
    }

    const tableContainer = document.createElement('div');
    tableContainer.style.cssText = `
        position: absolute; left: -9999px; top: 0;
        background: white; width: ${targetPxW}px; overflow: visible;
    `;
    _addTemp(tableContainer);

    // Shrink font until the table fits within the page
    let tableFontSize = 11;
    const mmPerPx = contentW / targetPxW;
    for (let attempt = 0; attempt < 12; attempt++) {
        tableContainer.innerHTML = '';
        const tbl = buildPdfTable(tableFontSize);
        tableContainer.appendChild(tbl);
        const tblW = tbl.scrollWidth;
        const tblH = tbl.scrollHeight;
        const fitsWidth = tblW <= targetPxW + 2;
        const renderedH = (targetPxW / Math.max(tblW, 1)) * tblH * mmPerPx;
        const fitsHeight = renderedH <= availableH;
        if (fitsWidth && fitsHeight) break;
        tableFontSize -= 0.5;
        if (tableFontSize < 5) break;
    }

    const tableCanvas = await captureElement(tableContainer, 2);
    _removeTemp(tableContainer);
    addImage(tableCanvas, contentW, availableH);

    // --- Charts ---

    async function addChartWithTitle(chartCanvas, title, widthMm, maxHeightMm) {
        if (!chartCanvas || chartCanvas.width === 0) return;
        const ratio = chartCanvas.height / chartCanvas.width;
        let imgH = widthMm * ratio;
        if (imgH > maxHeightMm) imgH = maxHeightMm;

        // Build a composite element with title + chart image
        const wrapper = document.createElement('div');
        wrapper.style.cssText = `background: white; padding: 4px; direction: rtl; width: ${Math.round(widthMm * 3.78)}px;`;
        if (title) {
            const h = document.createElement('div');
            h.style.cssText = 'font-family: "Assistant", sans-serif; font-size: 13px; font-weight: 700; color: #1a2a4a; text-align: center; margin-bottom: 4px;';
            h.textContent = title;
            wrapper.appendChild(h);
        }
        const img = document.createElement('img');
        img.src = chartCanvas.toDataURL('image/png');
        img.style.cssText = 'width: 100%; display: block;';
        wrapper.appendChild(img);
        _addTemp(wrapper);
        // Wait for image to load
        await new Promise(r => { if (img.complete) r(); else img.onload = r; });
        const compositeCanvas = await captureElement(wrapper, 2);
        _removeTemp(wrapper);

        const cRatio = compositeCanvas.height / compositeCanvas.width;
        let cH = widthMm * cRatio;
        if (cH > maxHeightMm + 8) cH = maxHeightMm + 8;
        ensureSpace(cH + 2);
        addImage(compositeCanvas, widthMm, maxHeightMm + 8);
    }

    // Main chart
    const mainCanvas = $('#pivot-chart');
    if (pivotChart && mainCanvas) {
        const mainTitle = colMode === 'employees' ? 'התפלגות שעות לפי עובדים' : 'שעות לפי חודשים ותיקים';
        await addChartWithTitle(mainCanvas, mainTitle, contentW, pageH * 0.35);
    }

    // Totals chart (months mode only)
    const totalsArea = $('#totals-chart-area');
    const totalsCanvasEl = $('#totals-chart');
    if (totalsChart && totalsCanvasEl && !totalsArea.classList.contains('hidden')) {
        await addChartWithTitle(totalsCanvasEl, 'מצטבר בכל התיקים', contentW, pageH * 0.35);
    }

    // Sub charts - rendered 2 per row
    const subChartCards = [...document.querySelectorAll('#sub-charts-area .sub-chart-card')];
    const halfW = (contentW - 4) / 2;

    for (let i = 0; i < subChartCards.length; i += 2) {
        const card1 = subChartCards[i];
        const card2 = subChartCards[i + 1];

        // Build a side-by-side wrapper
        const pairWrapper = document.createElement('div');
        pairWrapper.style.cssText = `display: flex; gap: 8px; direction: rtl; background: white; width: ${Math.round(contentW * 3.78)}px;`;

        // Clone cards for capture
        async function buildCardClone(card) {
            const clone = document.createElement('div');
            clone.style.cssText = 'flex: 1; text-align: center;';
            const title = card.querySelector('h4')?.textContent || '';
            if (title) {
                const h = document.createElement('div');
                h.style.cssText = 'font-family: "Assistant", sans-serif; font-size: 11px; font-weight: 700; color: #1a2a4a; text-align: center; margin-bottom: 2px;';
                h.textContent = title;
                clone.appendChild(h);
            }
            const cvs = card.querySelector('canvas');
            if (cvs) {
                const img = document.createElement('img');
                img.src = cvs.toDataURL('image/png');
                img.style.cssText = 'width: 100%; display: block;';
                clone.appendChild(img);
                await new Promise(r => { if (img.complete) r(); else img.onload = r; });
            }
            return clone;
        }

        const clone1 = await buildCardClone(card1);
        pairWrapper.appendChild(clone1);

        if (card2) {
            const clone2 = await buildCardClone(card2);
            pairWrapper.appendChild(clone2);
        } else {
            // Empty placeholder for odd last chart
            const empty = document.createElement('div');
            empty.style.cssText = 'flex: 1;';
            pairWrapper.appendChild(empty);
        }

        _addTemp(pairWrapper);
        const pairCanvas = await captureElement(pairWrapper, 2);
        _removeTemp(pairWrapper);

        const pRatio = pairCanvas.height / pairCanvas.width;
        let pH = contentW * pRatio;
        ensureSpace(pH + 4);
        addImage(pairCanvas, contentW, 90);
    }

    // --- Footer on all pages ---
    const totalPages = doc.getNumberOfPages();
    const now = new Date();
    const dateStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;
    const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(150, 150, 150);
        // Footer text — use simple LTR format to avoid font issues
        doc.text(`${dateStr} ${timeStr}`, pageW - margin, pageH - 6, { align: 'right' });
        doc.text(`${i} / ${totalPages}`, margin, pageH - 6, { align: 'left' });
        doc.setTextColor(0, 0, 0);
    }

    doc.save('דוח_שעות.pdf');

    } finally {
        // Clean up any orphaned temp DOM elements
        _cleanupAllTemp();
    }
}

// ============================================================
// Group Import/Export
// ============================================================
$('#download-emp-groups').addEventListener('click', () => {
    const data = [];
    Object.entries(employeeGroups).forEach(([group, members]) => {
        members.forEach(m => data.push({ 'קבוצה': group, 'עובד': m }));
    });
    // Append ungrouped employees from current data (not already in a group)
    const groupedSet = new Set(Object.values(employeeGroups).flat());
    getAllEmployees().forEach(emp => {
        if (!groupedSet.has(emp)) data.push({ 'קבוצה': '', 'עובד': emp });
    });
    // Append phantom employees (not in current data, not already included)
    const includedEmps = new Set(data.map(r => r['עובד']));
    phantomEmployees.forEach(emp => {
        if (!includedEmps.has(emp)) data.push({ 'קבוצה': '', 'עובד': emp });
    });
    if (!data.length) { alert('אין עובדים לייצוא'); return; }
    downloadExcel(data, 'קבוצות_עובדים.xlsx');
});

$('#upload-emp-groups').addEventListener('click', () => $('#emp-groups-file').click());
$('#emp-groups-file').addEventListener('change', (e) => {
    if (!e.target.files.length) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
        try {
            const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
            importEmployeeGroups(wb);
        } catch (err) {
            alert('שגיאה בקריאת קובץ קבוצות עובדים: ' + err.message);
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
    e.target.value = '';
});

$('#download-case-groups').addEventListener('click', () => {
    const data = [];
    Object.entries(caseGroups).forEach(([group, members]) => {
        members.forEach(m => {
            const parts = m.split('|');
            data.push({ 'קבוצה': group, 'לקוח': parts[0] || '', 'תיק': parts[1] || '' });
        });
    });
    // Append ungrouped cases from current data (not already in a group)
    const groupedKeys = new Set(Object.values(caseGroups).flat());
    getAllCases().forEach(c => {
        if (!groupedKeys.has(c.key)) {
            const parts = c.key.split('|');
            data.push({ 'קבוצה': '', 'לקוח': parts[0] || '', 'תיק': parts[1] || '' });
        }
    });
    // Append phantom cases (not in current data, not already included)
    const includedKeys = new Set(data.map(r => `${r['לקוח']}|${r['תיק']}`));
    phantomCases.forEach(key => {
        if (!includedKeys.has(key)) {
            const parts = key.split('|');
            data.push({ 'קבוצה': '', 'לקוח': parts[0] || '', 'תיק': parts[1] || '' });
        }
    });
    if (!data.length) { alert('אין תיקים לייצוא'); return; }
    downloadExcel(data, 'קבוצות_תיקים.xlsx');
});

$('#upload-case-groups').addEventListener('click', () => $('#case-groups-file').click());
$('#case-groups-file').addEventListener('change', (e) => {
    if (!e.target.files.length) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
        try {
            const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
            importCaseGroups(wb);
        } catch (err) {
            alert('שגיאה בקריאת קובץ קבוצות תיקים: ' + err.message);
        }
    };
    reader.readAsArrayBuffer(e.target.files[0]);
    e.target.value = '';
});

// ============================================================
// Utilities
// ============================================================
function downloadExcel(data, filename) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filename);
}

function getSheetHeaderRow(sheet) {
    return XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] || [];
}

function esc(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function escData(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ============================================================
// Visitor Counter
// ============================================================
(async function() {
    const el = $('#visitor-count');
    if (!el) return;
    try {
        // Only increment once per session to avoid inflating count on refresh
        if (!sessionStorage.getItem('visitor_counted')) {
            const resp = await fetch('https://api.counterapi.dev/v1/amirv01-time-report/visits/up');
            if (resp.ok) {
                const data = await resp.json();
                el.textContent = data.count != null ? data.count.toLocaleString() : '—';
                sessionStorage.setItem('visitor_counted', '1');
            } else { el.textContent = '—'; }
        } else {
            // Already counted this session — just fetch current count without incrementing
            const resp = await fetch('https://api.counterapi.dev/v1/amirv01-time-report/visits');
            if (resp.ok) {
                const data = await resp.json();
                el.textContent = data.count != null ? data.count.toLocaleString() : '—';
            } else { el.textContent = '—'; }
        }
    } catch (e) {
        el.textContent = '—';
    }
})();

// ============================================================
// Excel-style filter dropdown initialization
// ============================================================
function initXfDropdown(triggerId, dropdownId, searchId, onSearch, onClose) {
    const trigger = $(triggerId);
    const dropdown = $(dropdownId);
    const search = $(searchId);
    if (!trigger || !dropdown) return;

    trigger.addEventListener('click', (e) => {
        e.stopPropagation();
        const isOpen = !dropdown.classList.contains('hidden');
        // Close all other open dropdowns first
        $$('.xf-dropdown').forEach(d => d.classList.add('hidden'));
        if (!isOpen) {
            dropdown.classList.remove('hidden');
            setTimeout(() => search && search.focus(), 0);
        }
    });

    if (search) {
        search.addEventListener('input', () => onSearch(search.value));
    }
}

// Shared click-outside handler to close all XF dropdowns
document.addEventListener('click', (e) => {
    if (!e.target.closest('.xf-wrapper')) {
        $$('.xf-dropdown').forEach(d => d.classList.add('hidden'));
    }
});

initXfDropdown('#case-filter-trigger', '#case-filter-dropdown', '#case-filter-search',
    (term) => renderCaseFilterList(_caseFilterAllItems, term));

initXfDropdown('#emp-filter-trigger', '#emp-filter-dropdown', '#emp-filter-search',
    (term) => renderEmpFilterList(_empFilterAllItems, term, _empFilterSelectedSet));

initXfDropdown('#sub-chart-filter-trigger', '#sub-chart-filter-dropdown', '#sub-chart-filter-search',
    (term) => renderSubChartFilterList(_subChartFilterAllItems, term));
