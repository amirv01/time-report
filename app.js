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
let groupCases = true;
let selectedCaseGroups = new Set(); // which case groups to show in pivot
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

                if (name.includes('עובדים')) {
                    // Validate structure before treating as employee groups
                    const testRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { range: 0 });
                    if (testRows.length > 0 && 'קבוצה' in testRows[0] && 'עובד' in testRows[0]) {
                        importEmployeeGroups(wb);
                        showFileStatus(name, 'קבוצות עובדים נטענו');
                    } else {
                        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
                        parseReport(rows);
                        showFileStatus(name, 'דוח שעות נטען');
                    }
                } else if (name.includes('תיקים')) {
                    // Validate structure before treating as case groups
                    const testRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { range: 0 });
                    if (testRows.length > 0 && 'קבוצה' in testRows[0]) {
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
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

    // Check 15: Required columns
    if (!rows.length) { alert('קובץ קבוצות עובדים ריק'); return; }
    const first = rows[0];
    if (!('קבוצה' in first) || !('עובד' in first)) {
        alert('קובץ קבוצות עובדים לא תקין.\nנדרשות עמודות: "קבוצה", "עובד"');
        return;
    }

    const errors = [];
    const newGroups = {};
    const memberToGroups = {}; // track which groups each member belongs to

    rows.forEach((r, i) => {
        const group = r['קבוצה'];
        const member = r['עובד'];
        if (!group && !member) return; // skip empty rows
        // Check 16: Empty group names
        if (!group) { errors.push(`שורה ${i + 2}: חסר שם קבוצה`); return; }
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
    renderEmployeeGroups();
    renderPivot();
    console.log(`Imported ${Object.keys(newGroups).length} employee groups`);
}

// ============================================================
// Validation & Import: Case Groups
// ============================================================
function importCaseGroups(wb) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

    if (!rows.length) { alert('קובץ קבוצות תיקים ריק'); return; }
    const first = rows[0];
    if (!('קבוצה' in first)) {
        alert('קובץ קבוצות תיקים לא תקין.\nנדרשת עמודה: "קבוצה"\nאופציונלי: "לקוח", "תיק"');
        return;
    }
    // Must have at least one of לקוח or תיק
    if (!('לקוח' in first) && !('תיק' in first)) {
        alert('קובץ קבוצות תיקים לא תקין.\nנדרשת לפחות אחת מהעמודות: "לקוח", "תיק"');
        return;
    }

    const errors = [];
    const newGroups = {};
    rows.forEach((r, i) => {
        const group = r['קבוצה'];
        const client = r['לקוח'] || '';
        const cas = r['תיק'] || '';
        if (!group && !client && !cas) return; // skip empty rows
        if (!group) { errors.push(`שורה ${i + 2}: חסר שם קבוצה`); return; }
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
    // Select all new groups (+ אחר) in the filter
    Object.keys(newGroups).forEach(g => selectedCaseGroups.add(g));
    selectedCaseGroups.add('אחר');
    renderCaseGroups();
    rebuildCaseFilter();
    renderPivot();
    console.log(`Imported ${Object.keys(newGroups).length} case groups`);
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
    console.log('Detected format:', isFormat2 ? 'Format 2 (לפי לקוח/תיק)' : 'Format 1 (ByLawyerDate)');

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

    console.log('Parsed new entries:', newEntries.length);
    // Check 6: No data rows
    if (newEntries.length === 0) { alert('לא נמצאו רשומות תקינות בקובץ.\nוודאו שהקובץ מכיל שורות נתונים מתחת לשורת הכותרת.'); return; }

    // Collect all warnings into a single summary (Issue 13: avoid cascading alerts)
    const warnings = [];

    // Check 7: All dates invalid
    const validDateCount = newEntries.filter(e => e.date instanceof Date && !isNaN(e.date.getTime())).length;
    if (validDateCount === 0) {
        warnings.push(`כל ${newEntries.length} הרשומות ללא תאריך תקין — ייתכן שפורמט התאריך אינו מזוהה.`);
    } else if (validDateCount < newEntries.length) {
        const invalidCount = newEntries.length - validDateCount;
        console.warn(`${invalidCount} רשומות עם תאריך לא תקין מתוך ${newEntries.length}`);
    }

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

    // Check 10: Missing employee names
    const noEmployee = newEntries.filter(e => !e.employee || !e.employee.trim());
    if (noEmployee.length > 0) {
        warnings.push(`${noEmployee.length} רשומות ללא שם עובד.`);
    }

    // Check 11: Missing client/case
    const noClientCase = newEntries.filter(e => (!e.client || !e.client.trim()) && (!e.caseName || !e.caseName.trim()));
    if (noClientCase.length > 0) {
        warnings.push(`${noClientCase.length} רשומות ללא לקוח ותיק.`);
    }

    // Check 12: Duplicate rows
    const keyCount = {};
    newEntries.forEach(e => {
        const k = entryKey(e);
        keyCount[k] = (keyCount[k] || 0) + 1;
    });
    const internalDups = Object.values(keyCount).filter(c => c > 1).reduce((s, c) => s + (c - 1), 0);
    if (internalDups > 0) {
        warnings.push(`${internalDups} שורות כפולות בתוך הקובץ (אותו לקוח, תיק, עובד, תאריך ושעות חיוב).`);
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

    // Update date filters based on all current entries
    updateDateFilters();
    rebuildCaseFilter();

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
    return `${e.client}|${e.caseName}|${e.employee}|${dateStr}|${e.billableHours}`;
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
            `רשומה כפולה = אותו לקוח, תיק, עובד, תאריך ושעות חיוב.\n\n` +
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
    console.log('Format 1 - Header row:', headerRowIdx, 'Column map:', colMap);

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
    console.log('Format 2 - Header row:', headerRowIdx, 'Column map:', colMap);

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

    // Check 13: Warn about orphan rows
    if (orphanRows > 0) {
        console.warn(`Format 2: ${orphanRows} שורות נתונים לפני כותרת לקוח/תיק ראשונה`);
    }

    // Check 14: Warn about empty sections
    if (emptySections.length > 0) {
        console.warn('Format 2: Empty sections:', emptySections);
    }

    return entries;
}

// ============================================================
// Case filter checkboxes
// ============================================================
function rebuildCaseFilter() {
    const container = $('#case-filter-list');
    const groupNames = Object.keys(caseGroups);
    const items = [...groupNames, 'אחר'];

    // If selectedCaseGroups is empty, select all
    if (selectedCaseGroups.size === 0) {
        items.forEach(item => selectedCaseGroups.add(item));
    }

    container.innerHTML = items.map(name => {
        const checked = selectedCaseGroups.has(name) ? 'checked' : '';
        return `<label><input type="checkbox" value="${escData(name)}" ${checked} />${esc(name)}</label>`;
    }).join('');

    // Add event listeners
    container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.addEventListener('change', () => {
            if (cb.checked) selectedCaseGroups.add(cb.value);
            else selectedCaseGroups.delete(cb.value);
            renderPivot();
        });
    });
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
    return entries;
}

function getAllEmployees() {
    return [...new Set(rawEntries.map(e => e.employee))].filter(Boolean).sort();
}

function getAssignedEmployees() {
    const assigned = new Set();
    Object.values(employeeGroups).forEach(members => members.forEach(m => assigned.add(m)));
    return assigned;
}

function getAllCases() {
    const cases = new Map();
    rawEntries.forEach(e => {
        if (!cases.has(e.caseKey)) cases.set(e.caseKey, { client: e.client, caseName: e.caseName, key: e.caseKey });
    });
    return [...cases.values()].sort((a, b) => a.key.localeCompare(b.key));
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

function getAllMonths() {
    const months = new Set();
    rawEntries.forEach(e => {
        const m = formatMonth(e.date);
        if (m) months.add(m);
    });
    return [...months].sort((a, b) => {
        const [mA, yA] = a.split('/').map(Number);
        const [mB, yB] = b.split('/').map(Number);
        return yA !== yB ? yA - yB : mA - mB;
    });
}

// ============================================================
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
    debouncedRenderPivot();
});

// Group cases checkbox
$('#group-cases-cb').addEventListener('change', (e) => {
    groupCases = e.target.checked;
    rebuildCaseFilter();
    debouncedRenderPivot();
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
        if (action === 'remove-emp-group') { delete employeeGroups[btn.dataset.name]; renderEmployeeGroups(); renderPivot(); }
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
            cols = getAllEmployees();
            getCol = (e) => e.employee || 'ללא עובד';
            if (!cols.includes('ללא עובד') && entries.some(e => !e.employee)) cols.push('ללא עובד');
        }
    }

    // --- Build ROWS ---
    let rowKeys = [];
    let getRow; // function: entry -> row label
    let rowLabel; // function: row key -> display label

    if (groupCases && Object.keys(caseGroups).length > 0) {
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
            // If it's a group name, return as-is; if it's a caseKey, format it
            if (groupNames.has(r) || r === 'אחר') return r;
            return caseLabel(r);
        };
    } else {
        // Individual cases
        const allCases = getAllCases();
        rowKeys = allCases.map(c => c.key);
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

    entries.forEach(e => {
        const col = getCol(e);
        const row = getRow(e);
        if (!pivotData[row]) return; // skip rows not in selected groups
        const val = e[hourKey];
        pivotData[row][col] = (pivotData[row][col] || 0) + val;
        rowTotals[row] = (rowTotals[row] || 0) + val;
        colTotals[col] = (colTotals[col] || 0) + val;
        grandTotal += val;
    });

    // --- Render ---
    const cornerLabel = colMode === 'months' ? 'תיק / חודש' : 'תיק / עובד';
    const thead = $('#pivot-table thead');
    const tbody = $('#pivot-table tbody');

    thead.innerHTML = `<tr>
        <th class="pivot-corner">${cornerLabel}</th>
        ${cols.map(c => `<th>${esc(c)}</th>`).join('')}
        <th class="pivot-total-header">סה"כ</th>
    </tr>`;

    tbody.innerHTML = rowKeys.map(r => {
        const label = rowLabel(r);
        return `<tr>
            <td><strong>${esc(label)}</strong></td>
            ${cols.map(c => {
                const v = pivotData[r]?.[c] || 0;
                return `<td class="pivot-value">${v ? v.toFixed(2) : '-'}</td>`;
            }).join('')}
            <td class="pivot-value pivot-total">${(rowTotals[r] || 0).toFixed(2)}</td>
        </tr>`;
    }).join('') + `<tr class="pivot-total-row">
        <td><strong>סה"כ</strong></td>
        ${cols.map(c => `<td class="pivot-value">${(colTotals[c] || 0).toFixed(2)}</td>`).join('')}
        <td class="pivot-value pivot-total">${grandTotal.toFixed(2)}</td>
    </tr>`;

    // --- Build per-row-col-empgroup data for sub bar charts ---
    let empGroupLabels = [];
    let subBarData = {};  // { rowKey: { empGroup: { month: hours } } }
    if (colMode === 'months') {
        // Build employee group mapping
        const empGMap = {};
        if (Object.keys(employeeGroups).length > 0) {
            Object.entries(employeeGroups).forEach(([gName, members]) => {
                members.forEach(m => { empGMap[m] = gName; });
            });
        }
        const egSet = new Set();
        getAllEmployees().forEach(emp => {
            const g = empGMap[emp] || (Object.keys(employeeGroups).length > 0 ? 'אחר' : emp);
            // Only use groups, not individual employees for stacked bars
            if (empGMap[emp]) egSet.add(empGMap[emp]);
            else egSet.add('אחר');
        });
        empGroupLabels = [...egSet].sort();

        rowKeys.forEach(r => {
            subBarData[r] = {};
            empGroupLabels.forEach(eg => { subBarData[r][eg] = {}; });
        });

        entries.forEach(e => {
            const row = getRow(e);
            if (!subBarData[row]) return;
            const month = formatMonth(e.date) || 'ללא תאריך';
            const eg = empGMap[e.employee] || 'אחר';
            if (!subBarData[row][eg]) subBarData[row][eg] = {};
            subBarData[row][eg][month] = (subBarData[row][eg][month] || 0) + e[hourKey];
        });
    }

    // --- Render chart ---
    renderPivotChart(cols, rowKeys, rowLabel, pivotData, colTotals, empGroupLabels, subBarData);
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

function renderPivotChart(cols, rowKeys, rowLabel, pivotData, colTotals, empGroupLabels, subBarData) {
    const canvas = $('#pivot-chart');
    if (!canvas) return;

    // Destroy existing charts
    if (pivotChart) { pivotChart.destroy(); pivotChart = null; }
    if (totalsChart) { totalsChart.destroy(); totalsChart = null; }
    subCharts.forEach(c => c.destroy());
    subCharts = [];
    $('#sub-charts-area').innerHTML = '';
    $('#totals-chart-area').classList.add('hidden');

    // Build consistent color map for columns
    const colColorMap = {};
    cols.forEach((c, i) => { colColorMap[c] = CHART_COLORS[i % CHART_COLORS.length]; });

    if (colMode === 'employees') {
        // Include ALL employees in every pie chart (zeros hidden by legend filter)
        // so that colors are always consistent across main + sub charts
        const allLabels = cols;
        const allColors = cols.map(c => colColorMap[c]);
        const mainData = cols.map(c => colTotals[c] || 0);

        pivotChart = new Chart(canvas, {
            type: 'pie',
            data: {
                labels: allLabels,
                datasets: [{ data: mainData, backgroundColor: allColors, borderColor: '#fff', borderWidth: 1.5 }]
            },
            options: pieOptions('right', 12)
        });

        // Sub pie charts
        renderSubCharts(cols, rowKeys, rowLabel, pivotData, colColorMap, 'pie');
    } else {
        // Main bar chart (uses CASE_COLORS to distinguish from employee-group charts below)
        const datasets = rowKeys.map((r, i) => {
            const color = CASE_COLORS[i % CASE_COLORS.length];
            return { label: rowLabel(r), data: cols.map(c => pivotData[r]?.[c] || 0), backgroundColor: color, borderColor: color, borderWidth: 1, borderRadius: 2 };
        });
        pivotChart = new Chart(canvas, {
            type: 'bar',
            data: { labels: cols, datasets },
            options: barOptions(false)
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

function renderSubCharts(cols, rowKeys, rowLabel, dataMap, colorMap, chartType, empGroupLabels) {
    const container = $('#sub-charts-area');
    if (!container || rowKeys.length === 0) return;

    // Filter to rows with data
    const validRows = rowKeys.filter(r => {
        if (chartType === 'pie') {
            return cols.some(c => (dataMap[r]?.[c] || 0) > 0);
        } else {
            return empGroupLabels.some(eg => cols.some(c => (dataMap[r]?.[eg]?.[c] || 0) > 0));
        }
    });

    const BATCH = 4;
    let shown = 0;

    function showNextBatch() {
        const batch = validRows.slice(shown, shown + BATCH);
        batch.forEach(r => {
            const label = rowLabel(r);
            const card = document.createElement('div');
            card.className = 'sub-chart-card';
            card.innerHTML = `<h4>${esc(label)}</h4><canvas></canvas>`;
            // Insert before the "show more" button if it exists
            const moreBtn = container.querySelector('.show-more-btn');
            if (moreBtn) container.insertBefore(card, moreBtn);
            else container.appendChild(card);

            const subCanvas = card.querySelector('canvas');

            if (chartType === 'pie') {
                // Include ALL cols so colors stay consistent; zeros hidden by legend filter
                const data = cols.map(c => dataMap[r]?.[c] || 0);
                const colors = cols.map(c => colorMap[c]);
                const chart = new Chart(subCanvas, {
                    type: 'pie',
                    data: {
                        labels: cols,
                        datasets: [{ data: data, backgroundColor: colors, borderColor: '#fff', borderWidth: 1 }]
                    },
                    options: pieOptions('bottom', 10)
                });
                subCharts.push(chart);
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

    const btn = $('#download-pdf');
    btn.disabled = true;
    btn.textContent = '...מייצר דוח';

    try {
        await generatePdfReport(entries);
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

async function generatePdfReport(entries) {
    const { jsPDF } = window.jspdf;
    // Track temp DOM elements for cleanup on error
    const _tempElements = [];
    function _addTemp(el) { document.body.appendChild(el); _tempElements.push(el); return el; }
    function _removeTemp(el) { if (el.parentNode) el.parentNode.removeChild(el); const idx = _tempElements.indexOf(el); if (idx >= 0) _tempElements.splice(idx, 1); }
    function _cleanupAllTemp() { _tempElements.forEach(el => { if (el.parentNode) el.parentNode.removeChild(el); }); _tempElements.length = 0; }

    try {

    // --- Gather metadata ---
    const fromVal = $('#date-from').value;
    const toVal = $('#date-to').value;
    let dateRange = '';
    if (fromVal && toVal) dateRange = `${formatDateHebrew(new Date(fromVal + 'T00:00:00'))} - ${formatDateHebrew(new Date(toVal + 'T00:00:00'))}`;
    else if (fromVal) dateRange = `${formatDateHebrew(new Date(fromVal + 'T00:00:00'))} ואילך`;
    else if (toVal) dateRange = `עד ${formatDateHebrew(new Date(toVal + 'T00:00:00'))}`;
    else {
        const dates = entries.filter(e => e.date instanceof Date && !isNaN(e.date.getTime())).map(e => e.date);
        if (dates.length) {
            const timestamps = dates.map(d => d.getTime());
            const minD = new Date(timestamps.reduce((a, b) => a < b ? a : b));
            const maxD = new Date(timestamps.reduce((a, b) => a > b ? a : b));
            dateRange = `${formatDateHebrew(minD)} - ${formatDateHebrew(maxD)}`;
        }
    }

    const clients = [...new Set(entries.map(e => e.client).filter(Boolean))].sort();
    let clientStr = clients.slice(0, 3).join(', ');
    if (clients.length > 3) clientStr += ' ...';

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
    const titleParts = ['סיכום דוח שעות'];
    if (dateRange) titleParts.push(dateRange);
    if (clientStr) titleParts.push(clientStr);
    const titleText = titleParts.join(' || ');

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
    if (!data.length) { alert('אין קבוצות עובדים לייצוא'); return; }
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
    if (!data.length) { alert('אין קבוצות תיקים לייצוא'); return; }
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
