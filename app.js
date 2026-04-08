// ============================================================
// State
// ============================================================
let rawEntries = [];
let employeeGroups = {};
let caseGroups = {};
let valueMode = 'billable';
let ungroupedMode = 'individual';

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
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            // Use raw:true so numbers stay as numbers and dates stay as Date objects
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
            parseReport(rows);
        } catch (err) {
            console.error('Error reading file:', err);
            alert('שגיאה בקריאת הקובץ: ' + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ============================================================
// Parse Report
// ============================================================
function parseReport(rows) {
    rawEntries = [];
    let currentEmployee = null;
    let currentDate = null;

    // Find header row by scanning for "עובד" in any cell
    let headerRowIdx = -1;
    let colMap = {};
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
        const row = rows[i];
        if (!row) continue;
        for (let j = 0; j < row.length; j++) {
            if (row[j] === 'עובד') {
                headerRowIdx = i;
                break;
            }
        }
        if (headerRowIdx >= 0) break;
    }

    if (headerRowIdx < 0) {
        alert('לא נמצאה שורת כותרת בקובץ');
        return;
    }

    // Build column map from header row
    const headerRow = rows[headerRowIdx];
    const headerNames = {
        'תיאור': 'description',
        'סה"כ': 'total',
        'סה״כ': 'total',
        'תעריף': 'rate',
        'שעות חיוב': 'billableHours',
        'שעות עבודה': 'workHours',
        'סטטוס': 'status',
        'תיק': 'caseName',
        'לקוח': 'clientName',
        'תאריך': 'date',
        'עובד': 'employee'
    };

    for (let j = 0; j < headerRow.length; j++) {
        const cell = headerRow[j];
        if (cell && headerNames[cell]) {
            colMap[headerNames[cell]] = j;
        }
    }

    console.log('Header row:', headerRowIdx, 'Column map:', colMap);

    const dataStartIdx = headerRowIdx + 1;

    for (let i = dataStartIdx; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        // Check for subtotal rows
        const dateCell = row[colMap.date];
        if (typeof dateCell === 'string' && dateCell.includes('סה')) continue;

        const desc = row[colMap.description];
        const empCell = row[colMap.employee];
        const caseCell = row[colMap.caseName];
        const clientCell = row[colMap.clientName];

        // Skip rows that have no description, no employee, and no case (subtotal/blank rows)
        if (!desc && !empCell && !caseCell) continue;

        // Update current employee if present
        if (empCell) {
            currentEmployee = String(empCell).trim();
        }

        // Update current date if present
        if (dateCell != null && typeof dateCell !== 'string') {
            // Date object from SheetJS
            if (dateCell instanceof Date) {
                currentDate = dateCell;
            } else if (typeof dateCell === 'number') {
                // Excel serial date
                currentDate = excelDateToJS(dateCell);
            }
        } else if (typeof dateCell === 'string' && dateCell.trim() && !dateCell.includes('סה')) {
            const parsed = parseDate(dateCell);
            if (parsed) currentDate = parsed;
        }

        // Must have a description to be a real entry
        if (!desc) continue;

        // Must have hours
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
    if (rawEntries.length > 0) {
        console.log('Sample entry:', JSON.stringify(rawEntries[0], (k, v) => v instanceof Date ? v.toISOString() : v));
    }

    if (rawEntries.length === 0) {
        alert('לא נמצאו רשומות בקובץ');
        return;
    }

    // Set date range bounds
    const validDates = rawEntries
        .filter(e => e.date instanceof Date && !isNaN(e.date.getTime()))
        .map(e => e.date.getTime());
    if (validDates.length) {
        $('#date-from').value = formatDateISO(new Date(Math.min(...validDates)));
        $('#date-to').value = formatDateISO(new Date(Math.max(...validDates)));
    } else {
        // No valid dates — clear filters so nothing gets filtered out
        $('#date-from').value = '';
        $('#date-to').value = '';
    }

    // Show UI
    try {
        $('#controls-section').classList.remove('hidden');
        $('#tabs-section').classList.remove('hidden');
        showTab('clean-table');
        renderCleanTable();
        renderEmployeeGroups();
        renderCaseGroups();
        renderPivot();
    } catch (err) {
        console.error('Error rendering UI:', err);
        alert('שגיאה בהצגת הנתונים: ' + err.message);
    }
}

function toNum(val) {
    if (val == null) return null;
    if (typeof val === 'number') return val;
    const n = parseFloat(val);
    return isNaN(n) ? null : n;
}

function excelDateToJS(serial) {
    const utcDays = Math.floor(serial - 25569);
    return new Date(utcDays * 86400 * 1000);
}

function parseDate(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'string') {
        const d = new Date(val);
        if (!isNaN(d.getTime())) return d;
        const parts = val.split('/');
        if (parts.length === 3) {
            return new Date(parts[2], parts[1] - 1, parts[0]);
        }
    }
    return null;
}

function formatDateISO(d) {
    if (!d) return '';
    const dd = d instanceof Date ? d : new Date(d);
    if (isNaN(dd.getTime())) return '';
    const year = dd.getFullYear();
    const month = String(dd.getMonth() + 1).padStart(2, '0');
    const day = String(dd.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatDateHebrew(d) {
    if (!d) return '';
    const dd = d instanceof Date ? d : new Date(d);
    if (isNaN(dd.getTime())) return '';
    const day = String(dd.getDate()).padStart(2, '0');
    const month = String(dd.getMonth() + 1).padStart(2, '0');
    const year = dd.getFullYear();
    return `${day}/${month}/${year}`;
}

// ============================================================
// Filtered entries
// ============================================================
function getFilteredEntries() {
    let entries = rawEntries;
    const from = $('#date-from').value;
    const to = $('#date-to').value;
    if (from) {
        const fd = new Date(from + 'T00:00:00');
        if (!isNaN(fd.getTime())) {
            entries = entries.filter(e => !e.date || !(e.date instanceof Date) || e.date >= fd);
        }
    }
    if (to) {
        const td = new Date(to + 'T23:59:59');
        if (!isNaN(td.getTime())) {
            entries = entries.filter(e => !e.date || !(e.date instanceof Date) || e.date <= td);
        }
    }
    return entries;
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
}

// ============================================================
// Clean Table
// ============================================================
function renderCleanTable() {
    const entries = getFilteredEntries();
    const thead = $('#clean-table-data thead');
    const tbody = $('#clean-table-data tbody');

    thead.innerHTML = '<tr><th>עובד</th><th>תאריך</th><th>לקוח</th><th>תיק</th><th>תיאור</th><th>סטטוס</th><th>תעריף</th><th>שעות עבודה</th><th>שעות חיוב</th><th>סה"כ</th></tr>';
    tbody.innerHTML = entries.map(e => `<tr>
        <td>${esc(e.employee)}</td>
        <td>${formatDateHebrew(e.date)}</td>
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
        'עובד': e.employee,
        'תאריך': formatDateHebrew(e.date),
        'לקוח': e.client,
        'תיק': e.caseName,
        'תיאור': e.description,
        'סטטוס': e.status,
        'תעריף': e.rate,
        'שעות עבודה': e.workHours,
        'שעות חיוב': e.billableHours,
        'סה"כ': e.total
    }));
    downloadExcel(data, 'טבלה_שטוחה.xlsx');
});

// ============================================================
// Toggle controls
// ============================================================
$$('.toggle-group').forEach(group => {
    group.querySelectorAll('.toggle-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            group.querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const val = btn.dataset.value;
            if (val === 'billable' || val === 'work') valueMode = val;
            if (val === 'individual' || val === 'other') ungroupedMode = val;
            renderPivot();
        });
    });
});

$('#date-from').addEventListener('change', () => { renderCleanTable(); renderPivot(); });
$('#date-to').addEventListener('change', () => { renderCleanTable(); renderPivot(); });
$('#clear-dates').addEventListener('click', () => {
    $('#date-from').value = '';
    $('#date-to').value = '';
    renderCleanTable();
    renderPivot();
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

function getAllEmployees() {
    return [...new Set(rawEntries.map(e => e.employee))].filter(Boolean).sort();
}

function getAssignedEmployees() {
    const assigned = new Set();
    Object.values(employeeGroups).forEach(members => members.forEach(m => assigned.add(m)));
    return assigned;
}

function renderEmployeeGroups() {
    const assigned = getAssignedEmployees();
    const allEmps = getAllEmployees();
    const unassigned = allEmps.filter(e => !assigned.has(e));

    const groupsList = $('#emp-groups-list');
    groupsList.innerHTML = Object.entries(employeeGroups).map(([name, members]) => {
        const groupEl = document.createElement('div');
        return `<div class="group-card">
            <div class="group-header">
                <strong>${esc(name)}</strong>
                <button data-action="remove-emp-group" data-name="${escData(name)}" title="מחק קבוצה">&times;</button>
            </div>
            <div class="group-members" data-group="${escData(name)}" data-type="emp">
                ${members.map(m => `<span class="member-tag" draggable="true" data-member="${escData(m)}" data-type="emp">
                    ${esc(m)} <span class="remove-member" data-action="remove-emp-member" data-group="${escData(name)}" data-member="${escData(m)}">&times;</span>
                </span>`).join('')}
            </div>
        </div>`;
    }).join('');

    const unassignedEl = $('#unassigned-employees');
    unassignedEl.innerHTML = unassigned.map(e =>
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
        const name = btn.dataset.name;
        const group = btn.dataset.group;
        const member = btn.dataset.member;

        if (action === 'remove-emp-group') {
            delete employeeGroups[name];
            renderEmployeeGroups();
            renderPivot();
        } else if (action === 'remove-emp-member') {
            employeeGroups[group] = employeeGroups[group].filter(m => m !== member);
            renderEmployeeGroups();
            renderPivot();
        } else if (action === 'remove-case-group') {
            delete caseGroups[name];
            renderCaseGroups();
            renderPivot();
        } else if (action === 'remove-case-member') {
            caseGroups[group] = caseGroups[group].filter(m => m !== member);
            renderCaseGroups();
            renderPivot();
        }
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
    renderPivot();
});

$('#new-case-group-name').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') $('#add-case-group').click();
});

function getAllCases() {
    const cases = new Map();
    rawEntries.forEach(e => {
        if (!cases.has(e.caseKey)) {
            cases.set(e.caseKey, { client: e.client, caseName: e.caseName, key: e.caseKey });
        }
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

function renderCaseGroups() {
    const assigned = getAssignedCases();
    const allCases = getAllCases();
    const unassigned = allCases.filter(c => !assigned.has(c.key));

    const groupsList = $('#case-groups-list');
    groupsList.innerHTML = Object.entries(caseGroups).map(([name, members]) => `
        <div class="group-card">
            <div class="group-header">
                <strong>${esc(name)}</strong>
                <button data-action="remove-case-group" data-name="${escData(name)}" title="מחק קבוצה">&times;</button>
            </div>
            <div class="group-members" data-group="${escData(name)}" data-type="case">
                ${members.map(m => `<span class="member-tag" draggable="true" data-member="${escData(m)}" data-type="case">
                    ${esc(caseLabel(m))} <span class="remove-member" data-action="remove-case-member" data-group="${escData(name)}" data-member="${escData(m)}">&times;</span>
                </span>`).join('')}
            </div>
        </div>
    `).join('');

    const unassignedEl = $('#unassigned-cases');
    unassignedEl.innerHTML = unassigned.map(c =>
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
                member: tag.dataset.member,
                type: tag.dataset.type,
                fromGroup: tag.closest('.group-members')?.dataset.group || null
            }));
        });
    });

    container.querySelectorAll('.group-members').forEach(zone => {
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            const payload = JSON.parse(e.dataTransfer.getData('text/plain'));
            if (payload.type !== type) return;
            const targetGroup = zone.dataset.group;
            const groups = type === 'emp' ? employeeGroups : caseGroups;
            if (payload.fromGroup && groups[payload.fromGroup]) {
                groups[payload.fromGroup] = groups[payload.fromGroup].filter(m => m !== payload.member);
            }
            if (!groups[targetGroup].includes(payload.member)) {
                groups[targetGroup].push(payload.member);
            }
            if (type === 'emp') renderEmployeeGroups(); else renderCaseGroups();
            renderPivot();
        });
    });

    const unassignedZone = type === 'emp' ? $('#unassigned-employees') : $('#unassigned-cases');
    unassignedZone.addEventListener('dragover', (e) => { e.preventDefault(); unassignedZone.classList.add('drag-over'); });
    unassignedZone.addEventListener('dragleave', () => unassignedZone.classList.remove('drag-over'));
    unassignedZone.addEventListener('drop', (e) => {
        e.preventDefault();
        unassignedZone.classList.remove('drag-over');
        const payload = JSON.parse(e.dataTransfer.getData('text/plain'));
        if (payload.type !== type) return;
        const groups = type === 'emp' ? employeeGroups : caseGroups;
        if (payload.fromGroup && groups[payload.fromGroup]) {
            groups[payload.fromGroup] = groups[payload.fromGroup].filter(m => m !== payload.member);
        }
        if (type === 'emp') renderEmployeeGroups(); else renderCaseGroups();
        renderPivot();
    });
}

// ============================================================
// Pivot Table
// ============================================================
function renderPivot() {
    const entries = getFilteredEntries();
    if (!entries.length) return;

    const hourKey = valueMode === 'billable' ? 'billableHours' : 'workHours';

    const empGroupMap = {};
    Object.entries(employeeGroups).forEach(([gName, members]) => {
        members.forEach(m => { empGroupMap[m] = gName; });
    });

    const colLabels = new Set();
    getAllEmployees().forEach(emp => {
        if (empGroupMap[emp]) colLabels.add(empGroupMap[emp]);
        else if (ungroupedMode === 'individual') colLabels.add(emp);
        else colLabels.add('אחר');
    });
    const cols = [...colLabels].sort();

    const caseGroupMap = {};
    Object.entries(caseGroups).forEach(([gName, members]) => {
        members.forEach(m => { caseGroupMap[m] = gName; });
    });

    const rowLabels = new Set();
    getAllCases().forEach(c => {
        if (caseGroupMap[c.key]) rowLabels.add(caseGroupMap[c.key]);
        else if (ungroupedMode === 'individual') rowLabels.add(c.key);
        else rowLabels.add('אחר');
    });
    const rows = [...rowLabels].sort();

    const pivotData = {};
    const rowTotals = {};
    const colTotals = {};
    let grandTotal = 0;

    rows.forEach(r => { pivotData[r] = {}; rowTotals[r] = 0; });
    cols.forEach(c => { colTotals[c] = 0; });

    entries.forEach(e => {
        const col = empGroupMap[e.employee] || (ungroupedMode === 'individual' ? e.employee : 'אחר');
        const row = caseGroupMap[e.caseKey] || (ungroupedMode === 'individual' ? e.caseKey : 'אחר');
        const val = e[hourKey];
        if (!pivotData[row]) pivotData[row] = {};
        pivotData[row][col] = (pivotData[row][col] || 0) + val;
        rowTotals[row] = (rowTotals[row] || 0) + val;
        colTotals[col] = (colTotals[col] || 0) + val;
        grandTotal += val;
    });

    const thead = $('#pivot-table thead');
    const tbody = $('#pivot-table tbody');

    thead.innerHTML = `<tr>
        <th class="pivot-corner">תיק / עובד</th>
        ${cols.map(c => `<th>${esc(c)}</th>`).join('')}
        <th class="pivot-total-header">סה"כ</th>
    </tr>`;

    tbody.innerHTML = rows.map(r => {
        const label = r.includes('|') ? caseLabel(r) : r;
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
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        employeeGroups = {};
        rows.forEach(r => {
            const group = r['קבוצה'];
            const member = r['עובד'];
            if (group && member) {
                if (!employeeGroups[group]) employeeGroups[group] = [];
                if (!employeeGroups[group].includes(member)) employeeGroups[group].push(member);
            }
        });
        renderEmployeeGroups();
        renderPivot();
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
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        caseGroups = {};
        rows.forEach(r => {
            const group = r['קבוצה'];
            const client = r['לקוח'] || '';
            const cas = r['תיק'] || '';
            const key = client + '|' + cas;
            if (group) {
                if (!caseGroups[group]) caseGroups[group] = [];
                if (!caseGroups[group].includes(key)) caseGroups[group].push(key);
            }
        });
        renderCaseGroups();
        renderPivot();
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
