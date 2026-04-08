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
let groupCases = false;
let selectedCaseGroups = new Set(); // which case groups to show in pivot
let pivotChart = null;              // Chart.js instance

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

function handleFiles(fileList) {
    for (const file of fileList) {
        handleFile(file);
    }
}

function handleFile(file) {
    const name = file.name;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array', cellDates: true });
            const sheet = wb.Sheets[wb.SheetNames[0]];

            if (name.includes('עובדים')) {
                importEmployeeGroups(wb);
                showFileStatus(name, 'קבוצות עובדים נטענו');
            } else if (name.includes('תיקים')) {
                importCaseGroups(wb);
                showFileStatus(name, 'קבוצות תיקים נטענו');
            } else {
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
                parseReport(rows);
                showFileStatus(name, 'דוח שעות נטען');
            }
        } catch (err) {
            console.error('Error reading file:', err);
            alert(`שגיאה בקריאת הקובץ "${name}": ${err.message}`);
        }
    };
    reader.readAsArrayBuffer(file);
}

function showFileStatus(name, msg) {
    const el = $('#file-name');
    el.textContent = `${msg} (${name})`;
    clearTimeout(el._timer);
    el._timer = setTimeout(() => { el.textContent = ''; }, 4000);
}

// ============================================================
// Validation & Import: Employee Groups
// ============================================================
function importEmployeeGroups(wb) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

    // Validate structure
    if (!rows.length) { alert('קובץ קבוצות עובדים ריק'); return; }
    const first = rows[0];
    if (!('קבוצה' in first) || !('עובד' in first)) {
        alert('קובץ קבוצות עובדים לא תקין.\nנדרשות עמודות: "קבוצה", "עובד"');
        return;
    }

    const errors = [];
    const newGroups = {};
    rows.forEach((r, i) => {
        const group = r['קבוצה'];
        const member = r['עובד'];
        if (!group && !member) return; // skip empty rows
        if (!group) { errors.push(`שורה ${i + 2}: חסר שם קבוצה`); return; }
        if (!member) { errors.push(`שורה ${i + 2}: חסר שם עובד`); return; }
        const g = String(group).trim();
        const m = String(member).trim();
        if (!newGroups[g]) newGroups[g] = [];
        if (!newGroups[g].includes(m)) newGroups[g].push(m);
    });

    if (errors.length > 0 && Object.keys(newGroups).length === 0) {
        alert('קובץ קבוצות עובדים לא תקין:\n' + errors.slice(0, 5).join('\n'));
        return;
    }
    if (errors.length > 0) {
        console.warn('Employee groups import warnings:', errors);
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
    if (errors.length > 0) {
        console.warn('Case groups import warnings:', errors);
    }

    caseGroups = newGroups;
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
// Parse Report
// ============================================================
function parseReport(rows) {
    // Validate this looks like a time report
    if (!validateReportHeaders(rows)) {
        alert('הקובץ אינו דוח שעות תקין.\nלא נמצאו עמודות נדרשות (עובד, תאריך, שעות חיוב וכו\').\n\nאם זהו קובץ קבוצות, שנה את שמו ל:\n• קבוצות_עובדים.xlsx\n• קבוצות_תיקים.xlsx');
        return;
    }

    rawEntries = [];
    let currentEmployee = null;
    let currentDate = null;

    let headerRowIdx = -1;
    let colMap = {};
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const row = rows[i];
        if (!row) continue;
        for (let j = 0; j < row.length; j++) {
            if (row[j] === 'עובד') { headerRowIdx = i; break; }
        }
        if (headerRowIdx >= 0) break;
    }

    if (headerRowIdx < 0) return; // already validated above

    const headerRow = rows[headerRowIdx];
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

    console.log('Header row:', headerRowIdx, 'Column map:', colMap);

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

        rawEntries.push({
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

    console.log('Parsed entries:', rawEntries.length);

    if (rawEntries.length === 0) { alert('לא נמצאו רשומות בקובץ'); return; }

    const validDates = rawEntries
        .filter(e => e.date instanceof Date && !isNaN(e.date.getTime()))
        .map(e => e.date.getTime());
    if (validDates.length) {
        $('#date-from').value = formatDateISO(new Date(Math.min(...validDates)));
        $('#date-to').value = formatDateISO(new Date(Math.max(...validDates)));
    } else {
        $('#date-from').value = '';
        $('#date-to').value = '';
    }

    // Initialize case filter with all case groups + "אחר"
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
    return new Date((Math.floor(serial - 25569)) * 86400 * 1000);
}

function parseDate(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'string') {
        const d = new Date(val);
        if (!isNaN(d.getTime())) return d;
        const parts = val.split('/');
        if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
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
    if (from) {
        const fd = new Date(from + 'T00:00:00');
        if (!isNaN(fd.getTime())) entries = entries.filter(e => !e.date || !(e.date instanceof Date) || e.date >= fd);
    }
    if (to) {
        const td = new Date(to + 'T23:59:59');
        if (!isNaN(td.getTime())) entries = entries.filter(e => !e.date || !(e.date instanceof Date) || e.date <= td);
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
    return [...months].sort();
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
        renderPivot();
    });
});

// Ungrouped toggle
$('#ungrouped-toggle').querySelectorAll('.toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        $('#ungrouped-toggle').querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        ungroupedMode = btn.dataset.value;
        renderPivot();
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
        renderPivot();
    });
});

// Group employees checkbox
$('#group-employees-cb').addEventListener('change', (e) => {
    groupEmployees = e.target.checked;
    renderPivot();
});

// Group cases checkbox
$('#group-cases-cb').addEventListener('change', (e) => {
    groupCases = e.target.checked;
    rebuildCaseFilter();
    renderPivot();
});

// Date controls
$('#date-from').addEventListener('change', () => { renderCleanTable(); renderPivot(); });
$('#date-to').addEventListener('change', () => { renderCleanTable(); renderPivot(); });
$('#clear-dates').addEventListener('click', () => {
    $('#date-from').value = '';
    $('#date-to').value = '';
    renderCleanTable();
    renderPivot();
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
        <td>${e.rate}</td>
        <td>${e.workHours}</td>
        <td>${e.billableHours}</td>
        <td>${e.total}</td>
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
    setupGroupClicks('emp');
}

function setupGroupClicks(type) {
    const container = type === 'emp' ? $('#employee-groups') : $('#case-groups');
    if (!container) return;
    container.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-action]');
        if (!btn) return;
        const action = btn.dataset.action;
        if (action === 'remove-emp-group') { delete employeeGroups[btn.dataset.name]; renderEmployeeGroups(); renderPivot(); }
        else if (action === 'remove-emp-member') { employeeGroups[btn.dataset.group] = employeeGroups[btn.dataset.group].filter(m => m !== btn.dataset.member); renderEmployeeGroups(); renderPivot(); }
        else if (action === 'remove-case-group') { delete caseGroups[btn.dataset.name]; renderCaseGroups(); rebuildCaseFilter(); renderPivot(); }
        else if (action === 'remove-case-member') { caseGroups[btn.dataset.group] = caseGroups[btn.dataset.group].filter(m => m !== btn.dataset.member); renderCaseGroups(); rebuildCaseFilter(); renderPivot(); }
    });
}

// ============================================================
// Case Groups
// ============================================================
$('#add-case-group').addEventListener('click', () => {
    const name = $('#new-case-group-name').value.trim();
    if (!name || caseGroups[name]) return;
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
    setupGroupClicks('case');
}

// ============================================================
// Drag & Drop
// ============================================================
function setupDragDrop(type) {
    const container = type === 'emp' ? $('#employee-groups') : $('#case-groups');
    if (!container) return;

    container.querySelectorAll('[draggable="true"]').forEach(tag => {
        tag.addEventListener('dragstart', (e) => {
            e.dataTransfer.setData('text/plain', JSON.stringify({
                member: tag.dataset.member, type: tag.dataset.type,
                fromGroup: tag.closest('.group-members')?.dataset.group || null
            }));
        });
    });

    container.querySelectorAll('.group-members').forEach(zone => {
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault(); zone.classList.remove('drag-over');
            const payload = JSON.parse(e.dataTransfer.getData('text/plain'));
            if (payload.type !== type) return;
            const groups = type === 'emp' ? employeeGroups : caseGroups;
            if (payload.fromGroup && groups[payload.fromGroup]) groups[payload.fromGroup] = groups[payload.fromGroup].filter(m => m !== payload.member);
            if (!groups[zone.dataset.group].includes(payload.member)) groups[zone.dataset.group].push(payload.member);
            if (type === 'emp') renderEmployeeGroups(); else { renderCaseGroups(); rebuildCaseFilter(); }
            renderPivot();
        });
    });

    const unassignedZone = type === 'emp' ? $('#unassigned-employees') : $('#unassigned-cases');
    unassignedZone.addEventListener('dragover', (e) => { e.preventDefault(); unassignedZone.classList.add('drag-over'); });
    unassignedZone.addEventListener('dragleave', () => unassignedZone.classList.remove('drag-over'));
    unassignedZone.addEventListener('drop', (e) => {
        e.preventDefault(); unassignedZone.classList.remove('drag-over');
        const payload = JSON.parse(e.dataTransfer.getData('text/plain'));
        if (payload.type !== type) return;
        const groups = type === 'emp' ? employeeGroups : caseGroups;
        if (payload.fromGroup && groups[payload.fromGroup]) groups[payload.fromGroup] = groups[payload.fromGroup].filter(m => m !== payload.member);
        if (type === 'emp') renderEmployeeGroups(); else { renderCaseGroups(); rebuildCaseFilter(); }
        renderPivot();
    });
}

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
            else rowSet.add('אחר');
        });
        // Filter by selected case groups
        rowKeys = [...rowSet].filter(r => selectedCaseGroups.has(r)).sort();
        getRow = (e) => caseGroupMap[e.caseKey] || 'אחר';
        rowLabel = (r) => r;
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

    // --- Render chart ---
    renderPivotChart(cols, rowKeys, rowLabel, pivotData, colTotals);
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

function renderPivotChart(cols, rowKeys, rowLabel, pivotData, colTotals) {
    const canvas = $('#pivot-chart');
    if (!canvas) return;

    // Destroy existing chart
    if (pivotChart) {
        pivotChart.destroy();
        pivotChart = null;
    }

    if (colMode === 'employees') {
        // Pie chart: hours split by employee/employee group
        const labels = cols;
        const data = cols.map(c => colTotals[c] || 0);
        const colors = cols.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]);

        pivotChart = new Chart(canvas, {
            type: 'pie',
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    backgroundColor: colors,
                    borderColor: '#fff',
                    borderWidth: 1.5
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                        rtl: true,
                        textDirection: 'rtl',
                        labels: {
                            font: { family: "'Assistant', sans-serif", size: 12 },
                            padding: 12,
                            usePointStyle: true,
                            pointStyleWidth: 16
                        }
                    },
                    tooltip: {
                        rtl: true,
                        textDirection: 'rtl',
                        callbacks: {
                            label: function(ctx) {
                                const val = ctx.parsed;
                                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                                const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                                return `${ctx.label}: ${val.toFixed(2)} (${pct}%)`;
                            }
                        }
                    }
                }
            }
        });
    } else {
        // Bar chart: months on X, bars for each case/case group
        const labels = cols; // months
        const datasets = rowKeys.map((r, i) => {
            const label = rowLabel(r);
            const color = CHART_COLORS[i % CHART_COLORS.length];
            return {
                label: label,
                data: cols.map(c => pivotData[r]?.[c] || 0),
                backgroundColor: color,
                borderColor: color,
                borderWidth: 1,
                borderRadius: 2
            };
        });

        pivotChart = new Chart(canvas, {
            type: 'bar',
            data: { labels, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: {
                        ticks: {
                            font: { family: "'Assistant', sans-serif", size: 11 }
                        },
                        grid: { display: false }
                    },
                    y: {
                        beginAtZero: true,
                        ticks: {
                            font: { family: "'Assistant', sans-serif", size: 11 }
                        },
                        title: {
                            display: true,
                            text: valueMode === 'billable' ? 'שעות חיוב' : 'שעות עבודה',
                            font: { family: "'Assistant', sans-serif", size: 13, weight: '600' }
                        }
                    }
                },
                plugins: {
                    legend: {
                        position: 'bottom',
                        rtl: true,
                        textDirection: 'rtl',
                        labels: {
                            font: { family: "'Assistant', sans-serif", size: 11 },
                            padding: 10,
                            usePointStyle: true,
                            pointStyleWidth: 12,
                            boxWidth: 12
                        }
                    },
                    tooltip: {
                        rtl: true,
                        textDirection: 'rtl',
                        callbacks: {
                            label: function(ctx) {
                                return `${ctx.dataset.label}: ${ctx.parsed.y.toFixed(2)}`;
                            }
                        }
                    }
                }
            }
        });
    }
}

$('#download-pivot').addEventListener('click', () => {
    const table = $('#pivot-table');
    const wb = XLSX.utils.table_to_book(table, { sheet: 'טבלת ציר' });
    XLSX.writeFile(wb, 'טבלת_ציר.xlsx');
});

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
