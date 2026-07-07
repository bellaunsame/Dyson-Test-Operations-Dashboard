// ==========================================
// DASHBOARD - State, init, data, table render
// ==========================================

// Shared state exposed on window so gantt.js, pdf-export.js, email-chat.js can read it
window.OPS = window.OPS || {};

document.addEventListener('DOMContentLoaded', () => {
    // State
    window.OPS.currentProject  = '893';
    window.OPS.currentScale    = 'day';
    window.OPS.currentCategory = 'all';
    window.OPS.projectRows     = [];
    window.OPS.ganttInstance   = null;

    // DOM refs
    const totalTasksEl       = document.getElementById('stat-total-tasks');
    const inProgressEl       = document.getElementById('stat-in-progress');
    const completedEl        = document.getElementById('stat-completed');
    const avgProgressEl      = document.getElementById('stat-avg-progress');
    const projectSelect      = document.getElementById('project-select');
    const categorySelect     = document.getElementById('category-select');
    const selectedProjectTitle = document.getElementById('selected-project-title');
    const lnkManageProject   = document.getElementById('lnk-manage-project');
    const btnRefreshGantt    = document.getElementById('btn-refresh-gantt');
    const btnScaleButtons    = document.querySelectorAll('.btn-scale');
    const tasksTableBody     = document.querySelector('#tasks-table tbody');
    const inventorySearch    = document.getElementById('inventory-search');
    const inventoryFilter    = document.getElementById('inventory-filter');

    // Expose refs needed by other modules
    window.OPS.projectSelect  = projectSelect;
    window.OPS.tasksTableBody = tasksTableBody;

    initDashboard();

    function initDashboard() {
        loadProjectsDropdown();

        // Scale buttons
        btnScaleButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                btnScaleButtons.forEach(b => b.classList.remove('active'));
                e.currentTarget.classList.add('active');
                window.OPS.currentScale = e.currentTarget.getAttribute('data-scale');
                filterAndRenderGantt();
            });
        });

        // Project select
        if (projectSelect) {
            projectSelect.addEventListener('change', (e) => {
                window.OPS.currentProject = e.target.value;
                onProjectChanged();
            });
        }

        // Category select
        if (categorySelect) {
            categorySelect.addEventListener('change', (e) => {
                window.OPS.currentCategory = e.target.value;
                filterAndRenderGantt();
            });
        }

        // Refresh button
        if (btnRefreshGantt) {
            btnRefreshGantt.addEventListener('click', onProjectChanged);
        }

        // Inventory filters
        if (inventorySearch) inventorySearch.addEventListener('input', applyFilters);
        if (inventoryFilter) inventoryFilter.addEventListener('change', applyFilters);
    }

    function loadProjectsDropdown() {
        fetch('/api/projects')
            .then(res => res.json())
            .then(projects => {
                if (!projectSelect) return;
                projectSelect.innerHTML = '';
                projects.forEach(p => {
                    const opt = document.createElement('option');
                    opt.value = p.name;
                    opt.textContent = p.name;
                    if (p.name === window.OPS.currentProject) opt.selected = true;
                    projectSelect.appendChild(opt);
                });
                if (projectSelect.options.length > 0) {
                    window.OPS.currentProject = projectSelect.value;
                }
                onProjectChanged();
            })
            .catch(err => console.error('Error loading projects:', err));
    }

    function onProjectChanged() {
        if (!window.OPS.currentProject) return;
        if (inventorySearch) inventorySearch.value = '';
        if (inventoryFilter) inventoryFilter.value = 'all';
        if (selectedProjectTitle) {
            selectedProjectTitle.textContent = `Project Name: ${window.OPS.currentProject}`;
        }
        if (lnkManageProject) {
            lnkManageProject.href = `/tables/${window.OPS.currentProject}`;
        }
        fetchProjectRows();
    }

    function fetchProjectRows() {
        if (!window.OPS.currentProject) return;
        fetch(`/api/projects/${window.OPS.currentProject}/rows`)
            .then(res => res.json())
            .then(rows => {
                window.OPS.projectRows = rows;
                updateStats(rows);
                applyFilters();
                populateCategoryDropdown(rows);
                filterAndRenderGantt();
            })
            .catch(err => {
                console.error('Error fetching project rows:', err);
                showToast('Error loading project data', 'error');
            });
    }

    function populateCategoryDropdown(rows) {
        if (!categorySelect) return;
        const prev = categorySelect.value;
        const cats = [...new Set(rows.map(r => r['Category']).filter(Boolean))];
        categorySelect.innerHTML = '<option value="all">All Categories</option>';
        cats.forEach(cat => {
            const opt = document.createElement('option');
            opt.value = cat;
            opt.textContent = cat;
            categorySelect.appendChild(opt);
        });
        if (cats.includes(prev)) {
            categorySelect.value = prev;
            window.OPS.currentCategory = prev;
        } else {
            categorySelect.value = 'all';
            window.OPS.currentCategory = 'all';
        }
    }

    function filterAndRenderGantt() {
        let filtered = window.OPS.projectRows;
        if (window.OPS.currentCategory !== 'all') {
            filtered = filtered.filter(r => r['Category'] === window.OPS.currentCategory);
        }
        if (typeof renderGanttChart === 'function') {
            renderGanttChart(filtered, window.OPS.currentScale);
        }
    }
    // Expose so other modules can call it
    window.OPS.filterAndRenderGantt = filterAndRenderGantt;
    window.OPS.loadProjectsDropdown = loadProjectsDropdown;

    function applyFilters() {
        const searchVal = inventorySearch ? inventorySearch.value.toLowerCase().trim() : '';
        const filterVal = inventoryFilter ? inventoryFilter.value : 'all';
        let filtered = window.OPS.projectRows;

        // Filter by search term
        if (searchVal) {
            filtered = filtered.filter(r =>
                (r['Category']    && r['Category'].toLowerCase().includes(searchVal)) ||
                (r['Test Method'] && r['Test Method'].toLowerCase().includes(searchVal)) ||
                (r['Test Number'] && r['Test Number'].toLowerCase().includes(searchVal))
            );
        }

        // Filter by rejections
        if (filterVal === 'rejections') {
            filtered = filtered.filter(r => r['Comments'] && r['Comments'].trim() !== '');
        }

        renderTable(filtered);
    }

    function updateStats(rows) {
        if (!totalTasksEl || !inProgressEl || !completedEl || !avgProgressEl) return;
        const totalMethods = rows.length;
        let totalQty = 0;
        let totalRejects = 0;

        rows.forEach(r => {
            totalQty += parseInt(r['Proto Qty'] || 0) +
                        parseInt(r['DVT Qty']   || 0) +
                        parseInt(r['EVT Qty']   || 0) +
                        parseInt(r['PVT Qty']   || 0);
            if (r['Comments'] && r['Comments'].trim() !== '') totalRejects++;
        });

        const avgRejects = totalMethods > 0 ? (totalRejects / totalMethods).toFixed(1) : '0';
        totalTasksEl.textContent = totalMethods;
        inProgressEl.textContent = totalQty;
        completedEl.textContent  = totalRejects;
        avgProgressEl.textContent = avgRejects;
    }

    function renderTable(rows) {
        if (!tasksTableBody) return;
        tasksTableBody.innerHTML = '';

        if (rows.length === 0) {
            tasksTableBody.innerHTML = `
                <tr>
                    <td colspan="9" class="text-center" style="padding: 20px;">No test methods match the filters.</td>
                </tr>`;
            return;
        }

        rows.forEach(r => {
            const tr = document.createElement('tr');

            const formatCell = (w, d, q) => {
                const qtyStr = parseInt(q) > 0
                    ? `<strong style="color: #2ecc71;">(${q})</strong>`
                    : `(${q})`;
                return `${w}w ${d}d ${qtyStr}`;
            };

            const defectQty = parseInt(r['Defect Qty'] || 0);
            const defectCell = defectQty > 0
                ? `<span class="defect-badge" style="display:inline-flex;align-items:center;gap:4px;background:rgba(231,76,60,0.12);color:#e74c3c;border:1px solid rgba(231,76,60,0.25);border-radius:4px;padding:2px 7px;font-weight:700;font-size:11px;"><i class="fa-solid fa-bug" style="font-size:9px;"></i> ${defectQty}</span>`
                : `<span style="color:var(--color-text-muted);">0</span>`;

            tr.innerHTML = `
                <td style="font-weight:600;">${r['Category']}</td>
                <td>${r['Test Method']}</td>
                <td style="font-family:monospace;">${r['Test Number']}</td>
                <td>${formatCell(r['Proto Weeks'], r['Proto Days'], r['Proto Qty'])}</td>
                <td>${formatCell(r['DVT Weeks'],   r['DVT Days'],   r['DVT Qty'])}</td>
                <td>${formatCell(r['EVT Weeks'],   r['EVT Days'],   r['EVT Qty'])}</td>
                <td>${formatCell(r['PVT Weeks'],   r['PVT Days'],   r['PVT Qty'])}</td>
                <td style="text-align:center;">${defectCell}</td>
                <td style="text-align:center;">
                    <a href="/tables/${window.OPS.currentProject}" class="btn btn-secondary" style="padding:4px 8px;font-size:11px;">
                        <i class="fa-solid fa-pencil"></i> Edit
                    </a>
                </td>`;
            tasksTableBody.appendChild(tr);
        });
    }

    // Toast notification helper (used by all modules via window.showToast)
    window.showToast = function showToast(message, type = 'info') {
        const container = document.getElementById('toast-container');
        if (!container) return;

        const icons = { success: 'fa-circle-check', error: 'fa-circle-xmark', info: 'fa-circle-info' };
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.innerHTML = `
            <i class="fa-solid ${icons[type] || icons.info}"></i>
            <div class="toast-content">${message}</div>`;

        container.appendChild(toast);
        requestAnimationFrame(() => toast.classList.add('show'));
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 400);
        }, 4000);
    };

    // Also expose as local alias for backward compatibility within this file
    function showToast(msg, type) { window.showToast(msg, type); }
});
