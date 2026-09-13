let datasetData = null;
let rigNumbers = [];
let parsedRigs = {};
let rigColumnIndex = -1;
let componentNames = [];
let catalogueData = [];
const selectedOrderParts = new Set();

const datasetPaths = ['Rig 6(Dataset).csv', 'Rig(Dataset).csv'];
const cataloguePaths = ['Parts catalogue.xlsx', 'Parts catalouge.xlsx'];

window.addEventListener('DOMContentLoaded', autoLoadFiles);

async function autoLoadFiles() {
    const statusDiv = document.getElementById('statusMessage');

    try {
        let datasetResponse;
        let datasetFilename;
        for (const path of datasetPaths) {
            try {
                const response = await fetch(encodeURI(path));
                if (response.ok) {
                    datasetResponse = response;
                    datasetFilename = path;
                    break;
                }
            } catch {
            }
        }

        if (!datasetResponse) {
            throw new Error('No dataset CSV was found');
        }

        await loadCsvText(await datasetResponse.text(), statusDiv, getRigNumberFromFilename(datasetFilename));
        try {
            for (const path of cataloguePaths) {
                const catalogueResponse = await fetch(encodeURI(path));
                if (catalogueResponse.ok) {
                    await loadCatalogueBuffer(await catalogueResponse.arrayBuffer());
                    break;
                }
            }
        } catch {
        }
    } catch (error) {
        statusDiv.innerHTML = '<div class="status info">ℹ️ Choose the RIG folder to load its dataset and parts catalogue.</div>';
    }
}

async function chooseRigFolder() {
    try {
        if (window.showDirectoryPicker) {
            const directoryHandle = await window.showDirectoryPicker({ mode: 'read' });
            let datasetHandle;
            let datasetFilename;
            for (const path of datasetPaths) {
                try {
                    datasetHandle = await directoryHandle.getFileHandle(path);
                    datasetFilename = path;
                    break;
                } catch (error) {
                    if (error.name !== 'NotFoundError') throw error;
                }
            }
            if (!datasetHandle) throw new Error('No matching dataset file was found');
            await loadCsvText(
                await (await datasetHandle.getFile()).text(),
                document.getElementById('statusMessage'),
                getRigNumberFromFilename(datasetFilename)
            );
            for (const path of cataloguePaths) {
                try {
                    const catalogueHandle = await directoryHandle.getFileHandle(path);
                    await loadCatalogueFile(await catalogueHandle.getFile());
                    break;
                } catch (error) {
                    if (error.name !== 'NotFoundError') throw error;
                }
            }
            return;
        }

        document.getElementById('rigFolderFile').click();
    } catch (error) {
        if (error.name !== 'AbortError') {
            document.getElementById('statusMessage').innerHTML =
                `<div class="status error">❌ Could not load the RIG folder: ${escapeHTML(error.message)}</div>`;
        }
    }
}

async function loadFromFolder(files) {
    const selectedFiles = Array.from(files);
    const datasetFile = selectedFiles.find(file => datasetPaths.includes(file.name));
    const catalogueFile = selectedFiles.find(file => cataloguePaths.includes(file.name));
    const statusDiv = document.getElementById('statusMessage');

    if (!datasetFile) {
        statusDiv.innerHTML = '<div class="status error">❌ The selected folder must contain a matching dataset CSV</div>';
        return;
    }

    try {
        await loadCsvText(
            await datasetFile.text(),
            statusDiv,
            getRigNumberFromFilename(datasetFile.name)
        );
        if (catalogueFile) await loadCatalogueFile(catalogueFile);
    } catch (error) {
        statusDiv.innerHTML = `<div class="status error">❌ Error loading files: ${escapeHTML(error.message)}</div>`;
    }
}

async function loadCatalogueFile(file) {
    if (!file) return;
    await loadCatalogueBuffer(await file.arrayBuffer());
}

async function loadCatalogueBuffer(buffer) {
    if (!window.XLSX) throw new Error('Excel reader is unavailable');

    const workbook = XLSX.read(buffer, { type: 'array' });
    catalogueData = workbook.SheetNames.map(sheetName => ({
        sheetName,
        rows: XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            header: 1,
            defval: '',
            blankrows: false
        }).filter(row => row.some(value => String(value).trim() !== ''))
    })).filter(sheet => sheet.rows.length);

    const status = document.getElementById('statusMessage');
    status.innerHTML = `<div class="status success">✅ Dataset and parts catalogue loaded (${catalogueData.length} catalogue sheets)</div>`;
    fillForm();
}

function getRigNumberFromFilename(filename) {
    const match = String(filename || '').match(/rig[\s_-]*(\d+)/i);
    return match ? match[1] : '6';
}

async function loadCsvText(datasetText, statusDiv, rigNumber = '6') {
    datasetData = parseCSV(datasetText);

    finalizeDataset(statusDiv, rigNumber);
}

function parseCSV(text) {
    const delimiter = text.includes('\t') ? '\t' : ',';
    const rows = [];
    let row = [];
    let cell = '';
    let quoted = false;

    for (let index = 0; index < text.length; index++) {
        const character = text[index];
        const nextCharacter = text[index + 1];

        if (character === '"' && quoted && nextCharacter === '"') {
            cell += '"';
            index++;
        } else if (character === '"') {
            quoted = !quoted;
        } else if (character === delimiter && !quoted) {
            row.push(cell.trim());
            cell = '';
        } else if ((character === '\n' || character === '\r') && !quoted) {
            if (character === '\r' && nextCharacter === '\n') index++;
            row.push(cell.trim());
            if (row.some(value => value !== '')) rows.push(row);
            row = [];
            cell = '';
        } else {
            cell += character;
        }
    }

    if (cell !== '' || row.length > 0) {
        row.push(cell.trim());
        if (row.some(value => value !== '')) rows.push(row);
    }

    const width = rows.reduce((maximum, currentRow) => Math.max(maximum, currentRow.length), 0);
    return {
        headers: Array.from({ length: width }, (_, index) => `Column ${index + 1}`),
        data: rows.map(currentRow => currentRow.concat(Array(width - currentRow.length).fill('')))
    };
}

function normalize(value) {
    return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function escapeHTML(value) {
    return String(value || '').replace(/[&<>"']/g, character => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[character]));
}

function getDatasetParameter(row, rigColumnIndex) {
    for (let index = rigColumnIndex - 1; index >= 0; index--) {
        if (row[index] !== '') return row[index];
    }
    return '';
}

function getDatasetComponents() {
    const components = [];
    const seen = new Set();
    datasetData.data.slice(1).forEach(row => {
        const component = (row[0] || '').trim();
        if (component === 'No.') return;
        const key = normalize(component);
        if (component && !seen.has(key)) {
            seen.add(key);
            components.push(component);
        }
    });
    return components;
}

function extractRigsFromDataset(targetRigNumber) {
    if (!datasetData || !datasetData.data.length) return {};

    const rigHeaderRow = datasetData.data[0];
    const rigColumns = rigHeaderRow
        .map((value, index) => ({ value: String(value || '').trim(), index }))
        .filter(({ value }) => value === String(targetRigNumber));

    const rigs = {};
    rigColumns.forEach(({ value: rigNumber, index: columnIndex }) => {
        const fields = new Map();
        const entries = [];
        let currentComponent = 'General';
        let currentCentrifugeNumber = '';

        datasetData.data.slice(1).forEach(row => {
            const rowComponent = (row[0] || '').trim();
            if (rowComponent) currentComponent = rowComponent;

            const parameter = getDatasetParameter(row, columnIndex);
            const rigValue = row[columnIndex] || '';
            if (!parameter || !rigValue.trim()) return;

            if (rowComponent === 'No.') currentCentrifugeNumber = rigValue.trim();

            const entry = {
                component: currentComponent,
                name: parameter,
                value: rigValue,
                centrifugeNumber: currentCentrifugeNumber
            };
            entries.push(entry);
            if (!fields.has(normalize(parameter))) fields.set(normalize(parameter), entry);
        });

        if (entries.length) rigs[`RIG ${rigNumber}`] = { fields, entries };
    });

    rigColumnIndex = rigColumns.length ? rigColumns[0].index : -1;
    return rigs;
}

function finalizeDataset(statusDiv, rigNumber) {
    parsedRigs = extractRigsFromDataset(rigNumber);
    rigNumbers = Object.keys(parsedRigs);

    if (rigNumbers.length === 0) {
        statusDiv.innerHTML = '<div class="status error">❌ No RIGs found in dataset.</div>';
        return;
    }

    // Populate RIG selector
    const rigSelector = document.getElementById('rigSelector');
    rigSelector.innerHTML = '<option value="">-- Select a RIG --</option>';
    rigSelector.disabled = false;
    
    rigNumbers.forEach(rig => {
        const option = document.createElement('option');
        option.value = rig;
        option.textContent = rig;
        rigSelector.appendChild(option);
    });

    componentNames = getDatasetComponents();
    const componentSelector = document.getElementById('componentSelector');
    componentSelector.innerHTML = '<option value="">-- All components --</option>';
    if (componentNames.includes('Centrifuge')) {
        componentNames = ['Centrifuge', ...componentNames.filter(component => component !== 'Centrifuge')];
    }
    componentNames.forEach(component => {
        const option = document.createElement('option');
        option.value = component;
        option.textContent = component;
        componentSelector.appendChild(option);
    });
    componentSelector.disabled = false;
    populateCentrifugeSelector(parsedRigs[rigNumbers[0]]?.entries || []);

    statusDiv.innerHTML = `<div class="status success">✅ Dataset loaded! Found ${rigNumbers.length} RIG</div>`;
}

function populateCentrifugeSelector(entries) {
    const centrifugeNumbers = [...new Set(entries
        .map(entry => entry.centrifugeNumber)
        .filter(Boolean))];
    const centrifugeSelector = document.getElementById('centrifugeSelector');
    centrifugeSelector.innerHTML = '<option value="">-- Select No. --</option>';
    centrifugeNumbers.forEach(number => {
        const option = document.createElement('option');
        option.value = number;
        option.textContent = `No. ${number}`;
        centrifugeSelector.appendChild(option);
    });
    centrifugeSelector.disabled = centrifugeNumbers.length === 0;
}

function normalizePartText(value) {
    return normalize(value).replace(/[^a-z0-9]+/g, '');
}

function getComponentMatches(componentEntries) {
    const makes = componentEntries
        .filter(entry => normalize(entry.name) === 'make')
        .map(entry => normalizePartText(entry.value))
        .filter(Boolean);
    const models = componentEntries
        .filter(entry => normalize(entry.name) === 'model')
        .map(entry => normalizePartText(entry.value))
        .filter(Boolean);
    if (!makes.length && !models.length) return [];

    return catalogueData
        .map(sheet => {
            const sheetKey = normalizePartText(sheet.sheetName);
            const makeMatch = makes.some(make => sheetKey.includes(make));
            const modelMatch = models.some(model => sheetKey.includes(model));
            return { sheet, score: (makeMatch ? 2 : 0) + (modelMatch ? 2 : 0) };
        })
        .filter(({ score }) => score > 0)
        .sort((left, right) => right.score - left.score)
        .map(({ sheet }) => sheet);
}

function renderCatalogue(componentEntries) {
    const cataloguePreview = document.getElementById('cataloguePreview');
    if (!catalogueData.length) {
        cataloguePreview.innerHTML = '<div class="empty-state"><p>Parts catalogue not loaded</p></div>';
        return;
    }

    const matches = getComponentMatches(componentEntries);
    if (!matches.length) {
        cataloguePreview.innerHTML = '<div class="empty-state"><p>No matching catalogue part found for this component</p></div>';
        return;
    }

    cataloguePreview.innerHTML = matches.map(match => {
        const headerRow = match.rows[0];
        const hasHeader = headerRow.some(cell => /part|description|qty|quantity|page|stl/i.test(String(cell)));
        const headers = hasHeader
            ? headerRow.map((cell, index) => String(cell).trim() || `Field ${index + 1}`)
            : headerRow.map((_, index) => `Field ${index + 1}`);
        const dataRows = hasHeader ? match.rows.slice(1) : match.rows;
        const headerHTML = `<thead><tr><th class="order-column">Order</th>${headers.map((header, index) => `<th class="${getCatalogueColumnClass(header, index)}">${escapeHTML(header)}</th>`).join('')}</tr></thead>`;
        const rows = dataRows.map((row, index) => {
            const selectionKey = getCatalogueSelectionKey(match.sheetName, index, row);
            const checked = selectedOrderParts.has(selectionKey) ? ' checked' : '';
            const cells = headers.map((header, cellIndex) => `<td class="${getCatalogueColumnClass(header, cellIndex)}">${escapeHTML(row[cellIndex] || '')}</td>`).join('');
            return `<tr><td class="order-column"><input type="checkbox" aria-label="Select ${escapeHTML(match.sheetName)} part" data-selection-key="${escapeHTML(selectionKey)}"${checked} onchange="updateOrderSelection(this)"></td>${cells}</tr>`;
        }).join('');
        return `<h3>${escapeHTML(match.sheetName)}</h3><div class="catalogue-table-wrap"><table>${headerHTML}<tbody>${rows}</tbody></table></div>`;
    }).join('');
}

function getCatalogueColumnClass(header, index) {
    const normalizedHeader = normalize(header);
    if (normalizedHeader === 'page') return 'page-column';
    if (normalizedHeader === 'qty' || normalizedHeader === 'quantity') return 'quantity-column';
    return `catalogue-column-${index}`;
}

function getCatalogueSelectionKey(sheetName, rowIndex, row) {
    return `${sheetName}::${rowIndex}::${row.join('|')}`;
}

function updateOrderSelection(checkbox) {
    const selectionKey = checkbox.dataset.selectionKey;
    if (checkbox.checked) selectedOrderParts.add(selectionKey);
    else selectedOrderParts.delete(selectionKey);

    const orderCount = document.getElementById('orderCount');
    if (orderCount) orderCount.textContent = `${selectedOrderParts.size} selected for order`;
}

function fillForm() {
    const rigSelector = document.getElementById('rigSelector');
    const selectedRig = rigSelector.value;
    const datasetPreview = document.getElementById('datasetPreview');
    const statusDiv = document.getElementById('statusMessage');

    statusDiv.innerHTML = '';

    if (!selectedRig) {
        statusDiv.innerHTML = '<div class="status error">❌ Please select a RIG</div>';
        return;
    }

    if (!parsedRigs[selectedRig]) {
        statusDiv.innerHTML = '<div class="status error">❌ RIG data not found</div>';
        return;
    }

    const rigData = parsedRigs[selectedRig];
    const selectedComponent = document.getElementById('componentSelector').value;
    const centrifugeSelectorGroup = document.getElementById('centrifugeSelectorGroup');
    const centrifugeSelector = document.getElementById('centrifugeSelector');
    const selectedCentrifuge = centrifugeSelector.value;
    centrifugeSelectorGroup.style.display = selectedComponent === 'Centrifuge' ? 'flex' : 'none';
    const visibleEntries = selectedComponent === 'Centrifuge'
        ? (selectedCentrifuge
            ? rigData.entries.filter(entry => entry.centrifugeNumber === selectedCentrifuge)
            : [])
        : selectedComponent
            ? rigData.entries.filter(entry => entry.component === selectedComponent)
            : rigData.entries;
    let previewHTML = '';
    let filledCount = 0;

    visibleEntries.forEach(({ component, name, value }) => {
        if (value) {
            previewHTML += `
                <div class="preview-row">
                    <strong>${escapeHTML(component)} / ${escapeHTML(name)}:</strong> <span>${escapeHTML(value)}</span>
                </div>
            `;
            filledCount++;
        }
    });

    datasetPreview.innerHTML = previewHTML || '<div class="empty-state"><p>No populated values for this RIG</p></div>';
    renderCatalogue(selectedComponent ? visibleEntries : []);

    statusDiv.innerHTML = `<div class="status success">✅ Showing ${selectedRig} component data (${filledCount} values)</div>`;
}

