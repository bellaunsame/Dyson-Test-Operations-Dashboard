document.addEventListener('DOMContentLoaded', () => {
    // State
    let selectedFilePath = '';
    let excelFileName = '';
    let sheetsData = {}; // sheetName -> { columns: [...], rows: [...] }
    let selectedSheets = new Set(); // Set of sheet names checked
    let activeSheet = ''; // Currently previewed/edited sheet
    let removedColumns = {}; // sheetName -> list of removed column names
    let appliedSteps = {}; // sheetName -> list of step objects
    let rightClickedColumn = ''; // Track column for context menu

    // DOM Elements - Step 1 (Browse)
    const fileListTbody = document.getElementById('file-list-tbody');
    const btnFileNext   = document.getElementById('btn-file-next');

    // DOM Elements - Step 2 (Navigator)
    const panelFileSelect = document.getElementById('panel-file-select');
    const panelNavigator  = document.getElementById('panel-navigator');
    const navFilenameLabel = document.getElementById('nav-filename-label');
    const navSheetsContainer = document.getElementById('nav-sheets-container');
    const navSearchInput   = document.getElementById('nav-search-input');
    const navPreviewTitle  = document.getElementById('nav-preview-title');
    const navTableThead    = document.getElementById('nav-table-thead');
    const navTableTbody    = document.getElementById('nav-table-tbody');
    const btnNavCancel     = document.getElementById('btn-nav-cancel');
    const btnNavClose      = document.getElementById('btn-nav-close');
    const btnNavTransform  = document.getElementById('btn-nav-transform');
    const btnNavLoad       = document.getElementById('btn-nav-load');

    // DOM Elements - Step 3 (Query Editor)
    const panelQueryEditor = document.getElementById('panel-query-editor');
    const formulaBarInput  = document.getElementById('formula-bar-input');
    const editorQueriesContainer = document.getElementById('editor-queries-container');
    const editorGridThead  = document.getElementById('editor-grid-thead');
    const editorGridTbody  = document.getElementById('editor-grid-tbody');
    const queryNameInput   = document.getElementById('query-name-input');
    const appliedStepsContainer = document.getElementById('applied-steps-container');
    const btnEditorBack    = document.getElementById('btn-editor-back');
    const btnEditorApply   = document.getElementById('btn-editor-apply');
    const colContextMenu   = document.getElementById('column-context-menu');

    // Init
    initImportFlow();

    function initImportFlow() {
        loadSharePointFiles();

        // Step 1 -> Open Navigator
        btnFileNext.addEventListener('click', () => {
            openNavigator();
        });

        // Navigator Buttons
        btnNavCancel.addEventListener('click', closeNavigator);
        btnNavClose.addEventListener('click', closeNavigator);
        btnNavTransform.addEventListener('click', openQueryEditor);
        btnNavLoad.addEventListener('click', loadDirectly);

        // Navigator Search
        navSearchInput.addEventListener('input', (e) => {
            filterNavigatorTree(e.target.value.toLowerCase());
        });

        // Query Editor Buttons
        btnEditorBack.addEventListener('click', () => {
            panelQueryEditor.classList.remove('active');
            panelNavigator.classList.add('active');
        });
        btnEditorApply.addEventListener('click', submitTransformedData);

        // Close context menu on click elsewhere
        document.addEventListener('click', () => {
            colContextMenu.style.display = 'none';
        });

        // Context Menu Actions
        document.getElementById('menu-remove-col').addEventListener('click', () => {
            if (rightClickedColumn) {
                removeColumn(activeSheet, rightClickedColumn);
            }
        });
        document.getElementById('menu-remove-other').addEventListener('click', () => {
            if (rightClickedColumn) {
                removeOtherColumns(activeSheet, rightClickedColumn);
            }
        });
    }

    // ── Load SharePoint File List ──
    function loadSharePointFiles() {
        fetch('/api/sharepoint/browse')
            .then(res => res.json())
            .then(files => {
                fileListTbody.innerHTML = '';
                if (files.length === 0) {
                    fileListTbody.innerHTML = `<tr><td colspan="4" class="text-center">No Excel files found in SharePoint.</td></tr>`;
                    return;
                }

                files.forEach(f => {
                    const tr = document.createElement('tr');
                    tr.setAttribute('data-path', f.path);
                    tr.setAttribute('data-name', f.name);
                    tr.innerHTML = `
                        <td style="font-weight: 600;"><i class="fa-regular fa-file-excel" style="color: #27ae60; margin-right: 8px;"></i> ${f.name}</td>
                        <td style="color: var(--color-text-muted); font-family: monospace; font-size: 11px;">${f.path}</td>
                        <td>${f.size}</td>
                        <td style="color: var(--color-text-muted);">${f.modified}</td>
                    `;

                    tr.addEventListener('click', () => {
                        document.querySelectorAll('.file-select-table tr').forEach(row => row.classList.remove('selected'));
                        tr.classList.add('selected');
                        selectedFilePath = f.path;
                        excelFileName = f.name;
                        btnFileNext.disabled = false;
                    });

                    fileListTbody.appendChild(tr);
                });
            })
            .catch(() => {
                fileListTbody.innerHTML = `<tr><td colspan="4" class="text-center" style="color: #e74c3c;">Failed to connect to SharePoint. Please check your credentials.</td></tr>`;
            });
    }

    // ── Step 1 -> Step 2: Open Navigator ──
    function openNavigator() {
        btnFileNext.disabled = true;
        btnFileNext.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Loading File...';

        fetch('/api/sharepoint/preview-all', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ file_path: selectedFilePath })
        })
        .then(res => res.json())
        .then(data => {
            if (data.error) throw new Error(data.error);

            sheetsData = data.sheets;
            selectedSheets.clear();
            removedColumns = {};
            appliedSteps = {};

            // Set up tree filename label
            navFilenameLabel.textContent = `${excelFileName} [${Object.keys(sheetsData).length}]`;

            // Render Navigator Left Tree
            renderNavigatorTree();

            // Clear preview
            navPreviewTitle.textContent = "Select an item to preview";
            navTableThead.innerHTML = '';
            navTableTbody.innerHTML = '<tr><td class="text-center" style="padding: 40px; color: var(--color-text-muted);">Select a sheet on the left to display its data preview.</td></tr>';

            // Show Navigator Panel
            panelFileSelect.classList.remove('active');
            panelNavigator.classList.add('active');
        })
        .catch(err => showToast(err.message || 'Failed to load Excel sheets', 'error'))
        .finally(() => {
            btnFileNext.disabled = false;
            btnFileNext.innerHTML = 'Open Navigator <i class="fa-solid fa-arrow-right"></i>';
        });
    }

    function closeNavigator() {
        panelNavigator.classList.remove('active');
        panelFileSelect.classList.add('active');
    }

    // ── Render Navigator Left Tree ──
    function renderNavigatorTree() {
        navSheetsContainer.innerHTML = '';
        Object.keys(sheetsData).forEach(sheetName => {
            const div = document.createElement('div');
            div.className = 'tree-item';
            div.setAttribute('data-sheet', sheetName.toLowerCase());
            div.innerHTML = `
                <input type="checkbox" class="sheet-checkbox" data-sheet="${sheetName}">
                <i class="fa-solid fa-table" style="color: #6b7280;"></i>
                <span>${sheetName}</span>
            `;

            // Checkbox change
            const chk = div.querySelector('.sheet-checkbox');
            chk.addEventListener('change', (e) => {
                if (e.target.checked) {
                    selectedSheets.add(sheetName);
                } else {
                    selectedSheets.delete(sheetName);
                }
                updateNavigatorButtons();
            });

            // Row click to preview
            div.addEventListener('click', (e) => {
                if (e.target.type === 'checkbox') return; // Don't trigger twice if clicking checkbox
                document.querySelectorAll('.tree-item').forEach(el => el.classList.remove('active'));
                div.classList.add('active');
                activeSheet = sheetName;
                renderNavigatorPreview(sheetName);
            });

            navSheetsContainer.appendChild(div);
        });
    }

    function filterNavigatorTree(query) {
        document.querySelectorAll('.tree-item').forEach(el => {
            const sheet = el.getAttribute('data-sheet');
            if (sheet.includes(query)) {
                el.style.display = 'flex';
            } else {
                el.style.display = 'none';
            }
        });
    }

    function updateNavigatorButtons() {
        const hasSelection = selectedSheets.size > 0;
        btnNavLoad.disabled = !hasSelection;
        btnNavTransform.disabled = !hasSelection;
    }

    // ── Render Preview in Navigator ──
    function renderNavigatorPreview(sheetName) {
        const data = sheetsData[sheetName];
        navPreviewTitle.textContent = sheetName;
        navPreviewSubtitle.textContent = `Preview downloaded today • ${data.rows.length} rows loaded`;

        // Render Head
        navTableThead.innerHTML = `
            <tr style="background: rgba(255,255,255,0.02); text-align: left; border-bottom: 2px solid var(--color-border);">
                ${data.columns.map(col => `<th style="padding: 10px 12px; font-weight: 600;">${col}</th>`).join('')}
            </tr>
        `;

        // Render Body (first 15 rows for preview speed)
        navTableTbody.innerHTML = data.rows.slice(0, 15).map(row => `
            <tr style="border-bottom: 1px solid var(--color-border);">
                ${row.map(val => `<td style="padding: 8px 12px;">${val}</td>`).join('')}
            </tr>
        `).join('');

        if (data.rows.length === 0) {
            navTableTbody.innerHTML = `<tr><td colspan="${data.columns.length}" class="text-center" style="padding: 30px; color: var(--color-text-muted);">This worksheet is empty.</td></tr>`;
        }
    }

    // ── Step 2 -> Step 3: Open Query Editor ──
    function openQueryEditor() {
        // Initialize removed columns & applied steps for selected sheets
        selectedSheets.forEach(sheet => {
            if (!removedColumns[sheet]) removedColumns[sheet] = [];
            if (!appliedSteps[sheet]) {
                appliedSteps[sheet] = [{ id: 'source', name: 'Source' }];
            }
        });

        // Set active sheet in editor to the first selected sheet
        activeSheet = Array.from(selectedSheets)[0];

        // Render Left Queries list
        renderEditorQueriesList();

        // Render Grid & Settings
        renderEditorActiveQuery();

        // Switch panel
        panelNavigator.classList.remove('active');
        panelQueryEditor.add = panelQueryEditor.classList.add('active');
    }

    // ── Render Left Queries in Editor ──
    function renderEditorQueriesList() {
        editorQueriesContainer.innerHTML = '';
        selectedSheets.forEach(sheetName => {
            const div = document.createElement('div');
            div.className = `query-list-item ${sheetName === activeSheet ? 'active' : ''}`;
            div.innerHTML = `
                <i class="fa-solid fa-table" style="color: var(--color-primary);"></i>
                <span>${sheetName}</span>
            `;
            div.addEventListener('click', () => {
                activeSheet = sheetName;
                document.querySelectorAll('.query-list-item').forEach(el => el.classList.remove('active'));
                div.classList.add('active');
                renderEditorActiveQuery();
            });
            editorQueriesContainer.appendChild(div);
        });
    }

    // ── Render Active Query in Grid ──
    function renderEditorActiveQuery() {
        queryNameInput.value = activeSheet;
        const data = sheetsData[activeSheet];
        const removed = removedColumns[activeSheet] || [];

        // Filter out removed columns
        const activeColumns = data.columns.filter(col => !removed.includes(col));

        // Update Formula Bar
        updateFormulaBar(activeSheet);

        // Render Head
        editorGridThead.innerHTML = `
            <tr>
                ${activeColumns.map(col => {
                    // Try to guess type icon
                    let typeIcon = 'ABC';
                    const colLower = col.toLowerCase();
                    if (colLower.includes('date') || colLower.includes('start')) typeIcon = '📅';
                    else if (colLower.includes('qty') || colLower.includes('weeks') || colLower.includes('days') || colLower.includes('number')) typeIcon = '123';
                    
                    return `
                        <th>
                            <div class="th-content">
                                <span class="th-type-icon">${typeIcon}</span>
                                <span class="col-name" style="flex: 1; margin-left: 6px;">${col}</span>
                                <button class="th-menu-btn" data-col="${col}">
                                    <i class="fa-solid fa-chevron-down"></i>
                                </button>
                            </div>
                        </th>
                    `;
                }).join('')}
            </tr>
        `;

        // Wire Column Header Context Menus (right-click and button click)
        const thElements = editorGridThead.querySelectorAll('th');
        thElements.forEach(th => {
            const colName = th.querySelector('.col-name').textContent;
            
            // Right click
            th.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                showColumnContextMenu(e.clientX, e.clientY, colName);
            });

            // Chevron click
            const btn = th.querySelector('.th-menu-btn');
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const rect = btn.getBoundingClientRect();
                showColumnContextMenu(rect.left, rect.bottom + window.scrollY, colName);
            });
        });

        // Render Body (all loaded preview rows)
        // Find indices of active columns to render rows correctly
        const colIndices = activeColumns.map(col => data.columns.indexOf(col));

        editorGridTbody.innerHTML = data.rows.map(row => `
            <tr>
                ${colIndices.map(idx => `<td>${row[idx]}</td>`).join('')}
            </tr>
        `).join('');

        // Render Right Applied Steps
        renderAppliedSteps();
    }

    function showColumnContextMenu(x, y, columnName) {
        rightClickedColumn = columnName;
        colContextMenu.style.left = `${x}px`;
        colContextMenu.style.top = `${y}px`;
        colContextMenu.style.display = 'block';
    }

    // ── Update Formula Bar ──
    function updateFormulaBar(sheetName) {
        const removed = removedColumns[sheetName] || [];
        if (removed.length > 0) {
            const colList = removed.map(c => `"${c}"`).join(', ');
            formulaBarInput.value = `= Table.RemoveColumns(Source, {${colList}})`;
        } else {
            formulaBarInput.value = `= Table.TransformColumnTypes(#"Promoted Headers")`;
        }
    }

    // ── Render Right Applied Steps Pane ──
    function renderAppliedSteps() {
        appliedStepsContainer.innerHTML = '';
        const steps = appliedSteps[activeSheet] || [];

        steps.forEach((step, idx) => {
            const div = document.createElement('div');
            div.className = 'step-list-item';
            div.innerHTML = `
                <span>${step.name}</span>
                ${step.id !== 'source' ? `<button class="step-delete-btn" data-index="${idx}">&times;</button>` : ''}
            `;

            // Wire delete (undo) button
            if (step.id !== 'source') {
                div.querySelector('.step-delete-btn').addEventListener('click', (e) => {
                    e.stopPropagation();
                    undoStep(activeSheet, idx);
                });
            }

            appliedStepsContainer.appendChild(div);
        });
    }

    // ── Remove Column (Power Query Operation) ──
    function removeColumn(sheetName, columnName) {
        if (!removedColumns[sheetName].includes(columnName)) {
            removedColumns[sheetName].push(columnName);
            appliedSteps[sheetName].push({
                id: `remove_${columnName}`,
                name: `Removed Column: ${columnName}`,
                columnName: columnName
            });
            renderEditorActiveQuery();
        }
    }

    function removeOtherColumns(sheetName, columnName) {
        const data = sheetsData[sheetName];
        const colsToRemove = data.columns.filter(c => c !== columnName);
        
        colsToRemove.forEach(col => {
            if (!removedColumns[sheetName].includes(col)) {
                removedColumns[sheetName].push(col);
            }
        });

        appliedSteps[sheetName].push({
            id: 'remove_other',
            name: `Removed Other Columns`,
            restoredColumns: colsToRemove
        });

        renderEditorActiveQuery();
    }

    // ── Undo Step (Power Query Undo) ──
    function undoStep(sheetName, stepIndex) {
        const steps = appliedSteps[sheetName];
        const step = steps[stepIndex];

        if (step.id.startsWith('remove_')) {
            // Restore specific column
            if (step.id === 'remove_other') {
                const restored = step.restoredColumns || [];
                removedColumns[sheetName] = removedColumns[sheetName].filter(c => !restored.includes(c));
            } else {
                removedColumns[sheetName] = removedColumns[sheetName].filter(c => c !== step.columnName);
            }
        }

        // Remove this step and all subsequent steps (as they depend on it in a real pipeline)
        appliedSteps[sheetName] = steps.slice(0, stepIndex);

        renderEditorActiveQuery();
    }

    // ── Direct Load (No transformation) ──
    function loadDirectly() {
        btnNavLoad.disabled = true;
        btnNavLoad.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Loading...';

        const payload = {
            file_path: selectedFilePath,
            selected_sheets: Array.from(selectedSheets),
            removed_columns: {}
        };

        fetch('/api/sharepoint/load-transformed', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        })
        .then(res => res.json())
        .then(data => {
            if (data.error) throw new Error(data.error);
            showToast(data.message, 'success');
            setTimeout(() => {
                window.location.href = '/tables';
            }, 1500);
        })
        .catch(err => {
            showToast(err.message || 'Load failed', 'error');
            btnNavLoad.disabled = false;
            btnNavLoad.innerHTML = 'Load';
        });
    }

    // ── Submit Transformed Data (Close & Apply) ──
    function submitTransformedData() {
        btnEditorApply.disabled = true;
        btnEditorApply.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Applying Changes...';

        const payload = {
            file_path: selectedFilePath,
            selected_sheets: Array.from(selectedSheets),
            removed_columns: removedColumns
        };

        fetch('/api/sharepoint/load-transformed', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        })
        .then(res => res.json())
        .then(data => {
            if (data.error) throw new Error(data.error);
            showToast(data.message, 'success');
            setTimeout(() => {
                window.location.href = '/tables';
            }, 1500);
        })
        .catch(err => {
            showToast(err.message || 'Failed to apply transformations', 'error');
            btnEditorApply.disabled = false;
            btnEditorApply.innerHTML = 'Close & Apply';
        });
    }
});
