// ==========================================
// OPS DASHBOARD - FRONTEND LOGIC
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
    // Global State
    let currentProject = '893';
    let currentScale = 'day';
    let currentCategory = 'all';
    let projectRows = [];
    let ganttChartInstance = null;

    // DOM Elements
    const totalTasksEl = document.getElementById('stat-total-tasks'); // Total Projects/Test Methods
    const inProgressEl = document.getElementById('stat-in-progress'); // Total Items (Qty)
    const completedEl = document.getElementById('stat-completed');   // Total Rejects
    const avgProgressEl = document.getElementById('stat-avg-progress'); // Avg Rejects / Product
    
    const projectSelect = document.getElementById('project-select');
    const categorySelect = document.getElementById('category-select');
    const selectedProjectTitle = document.getElementById('selected-project-title');
    const lnkManageProject = document.getElementById('lnk-manage-project');
    const btnRefreshGantt = document.getElementById('btn-refresh-gantt');
    const btnScaleButtons = document.querySelectorAll('.btn-scale');
    const tasksTableBody = document.querySelector('#tasks-table tbody');
    const btnPdfDownload = document.getElementById('btn-pdf-download');
    
    // Inventory Filter Elements
    const inventorySearch = document.getElementById('inventory-search');
    const inventoryFilter = document.getElementById('inventory-filter');

    // Chatbot Elements
    const chatbotPanel = document.getElementById('chatbot-panel');
    const btnMinimizeChat = document.getElementById('btn-minimize-chat');
    const btnChatbotBubble = document.getElementById('btn-chatbot-bubble');
    const chatMessages = document.getElementById('chat-messages');
    const chatForm = document.getElementById('chat-form');
    const chatInput = document.getElementById('chat-input');
    
    // Email Elements
    const emailModal = document.getElementById('email-modal');
    const btnEmailModalTrigger = document.getElementById('btn-email-modal-trigger');
    const btnCloseModal = document.getElementById('btn-close-modal');
    const btnCancelEmail = document.getElementById('btn-cancel-email');
    const emailForm = document.getElementById('email-form');
    const btnSyncSharePoint = document.getElementById('btn-sync-sharepoint');

    // SharePoint sync state
    const syncStatusBadge = document.getElementById('sync-status-badge');
    const syncLastTime = document.getElementById('sync-last-time');
    const syncIcon = document.getElementById('sync-icon');

    // Initial Load
    initDashboard();

    function initDashboard() {
        loadProjectsDropdown();
        
        // Wire scale buttons
        btnScaleButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                btnScaleButtons.forEach(b => b.classList.remove('active'));
                e.currentTarget.classList.add('active');
                currentScale = e.currentTarget.getAttribute('data-scale');
                filterAndRenderGantt();
            });
        });

        // Wire project select change
        if (projectSelect) {
            projectSelect.addEventListener('change', (e) => {
                currentProject = e.target.value;
                onProjectChanged();
            });
        }

        // Wire category select change
        if (categorySelect) {
            categorySelect.addEventListener('change', (e) => {
                currentCategory = e.target.value;
                filterAndRenderGantt();
            });
        }

        // Wire refresh button
        if (btnRefreshGantt) {
            btnRefreshGantt.addEventListener('click', () => {
                onProjectChanged();
            });
        }

        // Wire local filters for Project Inventory
        if (inventorySearch) {
            inventorySearch.addEventListener('input', () => applyFilters());
        }
        if (inventoryFilter) {
            inventoryFilter.addEventListener('change', () => applyFilters());
        }
    }

    function loadProjectsDropdown() {
        fetch('/api/projects')
            .then(res => res.json())
            .then(projects => {
                if (projectSelect) {
                    projectSelect.innerHTML = '';
                    projects.forEach(p => {
                        const opt = document.createElement('option');
                        opt.value = p.name;
                        opt.textContent = p.name;
                        if (p.name === currentProject) {
                            opt.selected = true;
                        }
                        projectSelect.appendChild(opt);
                    });
                    // Set currentProject to whatever is selected
                    if (projectSelect.options.length > 0) {
                        currentProject = projectSelect.value;
                    }
                    onProjectChanged();
                }
            })
            .catch(err => console.error('Error loading projects:', err));
    }

    function onProjectChanged() {
        if (!currentProject) return;

        // Reset local filter inputs on project change
        if (inventorySearch) inventorySearch.value = '';
        if (inventoryFilter) inventoryFilter.value = 'all';

        // Update UI Titles and links
        if (selectedProjectTitle) {
            selectedProjectTitle.textContent = `Project Name: ${currentProject}`;
        }
        if (lnkManageProject) {
            lnkManageProject.href = `/tables/${currentProject}`;
        }

        fetchProjectRows();
    }

    function fetchProjectRows() {
        if (!currentProject) return;

        fetch(`/api/projects/${currentProject}/rows`)
            .then(res => res.json())
            .then(rows => {
                projectRows = rows;
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
        
        const previousSelection = categorySelect.value;
        const categories = [...new Set(rows.map(r => r["Category"]).filter(Boolean))];
        
        categorySelect.innerHTML = '<option value="all">All Categories</option>';
        categories.forEach(cat => {
            const opt = document.createElement('option');
            opt.value = cat;
            opt.textContent = cat;
            categorySelect.appendChild(opt);
        });
        
        if (categories.includes(previousSelection)) {
            categorySelect.value = previousSelection;
            currentCategory = previousSelection;
        } else {
            categorySelect.value = 'all';
            currentCategory = 'all';
        }
    }

    function filterAndRenderGantt() {
        let filtered = projectRows;
        if (currentCategory !== 'all') {
            filtered = projectRows.filter(r => r["Category"] === currentCategory);
        }
        renderGanttChart(filtered, currentScale);
    }

    // Apply local search and rejection filters
    function applyFilters() {
        const searchVal = inventorySearch ? inventorySearch.value.toLowerCase().trim() : '';
        const filterVal = inventoryFilter ? inventoryFilter.value : 'all';

        let filtered = projectRows;

        // 1. Filter by search term (Category, Method, or Number)
        if (searchVal) {
            filtered = filtered.filter(r => 
                (r["Category"] && r["Category"].toLowerCase().includes(searchVal)) ||
                (r["Test Method"] && r["Test Method"].toLowerCase().includes(searchVal)) ||
                (r["Test Number"] && r["Test Number"].toLowerCase().includes(searchVal))
            );
        }

        // 2. Filter by rejections status
        if (filterVal === 'rejections') {
            filtered = filtered.filter(r => r["Comments"] && r["Comments"].trim() !== "");
        }

        renderTable(filtered);
    }

    // Update Dashboard Statistics
    function updateStats(rows) {
        if (!totalTasksEl || !inProgressEl || !completedEl || !avgProgressEl) return;
        const totalMethods = rows.length;
        
        // Sum Qty across all phases
        let totalQty = 0;
        let totalRejects = 0; // Calculated from rows with comments
        
        rows.forEach(r => {
            totalQty += parseInt(r["Proto Qty"] || 0) + 
                       parseInt(r["DVT Qty"] || 0) + 
                       parseInt(r["EVT Qty"] || 0) + 
                       parseInt(r["PVT Qty"] || 0);
            if (r["Comments"] && r["Comments"].trim() !== "") {
                totalRejects += 1;
            }
        });
        
        const avgRejects = totalMethods > 0 ? (totalRejects / totalMethods).toFixed(1) : '0';

        totalTasksEl.textContent = totalMethods;
        inProgressEl.textContent = totalQty;
        completedEl.textContent = totalRejects;
        avgProgressEl.textContent = avgRejects;
    }

    // Render Table
    function renderTable(rows) {
        if (!tasksTableBody) return;
        tasksTableBody.innerHTML = '';

        if (rows.length === 0) {
            tasksTableBody.innerHTML = `
                <tr>
                    <td colspan="9" class="text-center" style="padding: 20px;">No test methods match the filters.</td>
                </tr>
            `;
            return;
        }

        rows.forEach(r => {
            const tr = document.createElement('tr');
            
            const formatCell = (w, d, q) => {
                const qtyStr = parseInt(q) > 0 ? `<strong style="color: #2ecc71;">(${q})</strong>` : `(${q})`;
                return `${w}w ${d}d ${qtyStr}`;
            };

            const defectQty = parseInt(r["Defect Qty"] || 0);
            const defectCell = defectQty > 0
                ? `<span class="defect-badge" style="display: inline-flex; align-items: center; gap: 4px; background: rgba(231,76,60,0.12); color: #e74c3c; border: 1px solid rgba(231,76,60,0.25); border-radius: 4px; padding: 2px 7px; font-weight: 700; font-size: 11px;"><i class="fa-solid fa-bug" style="font-size: 9px;"></i> ${defectQty}</span>`
                : `<span style="color: var(--color-text-muted);">0</span>`;

            tr.innerHTML = `
                <td style="font-weight: 600;">${r["Category"]}</td>
                <td>${r["Test Method"]}</td>
                <td style="font-family: monospace;">${r["Test Number"]}</td>
                <td>${formatCell(r["Proto Weeks"], r["Proto Days"], r["Proto Qty"])}</td>
                <td>${formatCell(r["DVT Weeks"], r["DVT Days"], r["DVT Qty"])}</td>
                <td>${formatCell(r["EVT Weeks"], r["EVT Days"], r["EVT Qty"])}</td>
                <td>${formatCell(r["PVT Weeks"], r["PVT Days"], r["PVT Qty"])}</td>
                <td style="text-align: center;">${defectCell}</td>
                <td style="text-align: center;">
                    <a href="/tables/${currentProject}" class="btn btn-secondary" style="padding: 4px 8px; font-size: 11px;">
                        <i class="fa-solid fa-pencil"></i> Edit
                    </a>
                </td>
            `;
            tasksTableBody.appendChild(tr);
        });
    }

    // Render Gantt Chart — Detailed Version
    function renderGanttChart(rows, scale = 'day') {
        const canvas = document.getElementById('ganttChart');
        if (!canvas) return;

        const ctx = canvas.getContext('2d');
        if (ganttChartInstance) {
            ganttChartInstance.destroy();
            ganttChartInstance = null;
        }

        const labels      = [];
        const chartData   = [];
        const bgColors    = [];
        const tooltipInfo = [];

        const phaseColors = {
            Proto: '#8E44AD',
            DVT:   '#2980B9',
            EVT:   '#27AE60',
            PVT:   '#D35400'
        };

        // Build bar data — one bar per phase per row
        rows.forEach(r => {
            let cursor = new Date(r['Start Date']);
            if (isNaN(cursor.getTime())) cursor = new Date();

            const phases = [
                { name: 'Proto', weeks: r['Proto Weeks'], days: r['Proto Days'], qty: r['Proto Qty'] },
                { name: 'DVT',   weeks: r['DVT Weeks'],   days: r['DVT Days'],   qty: r['DVT Qty']   },
                { name: 'EVT',   weeks: r['EVT Weeks'],   days: r['EVT Days'],   qty: r['EVT Qty']   },
                { name: 'PVT',   weeks: r['PVT Weeks'],   days: r['PVT Days'],   qty: r['PVT Qty']   }
            ];

            phases.forEach(p => {
                const totalDays = parseInt(p.weeks || 0) * 7 + parseInt(p.days || 0);
                if (totalDays > 0) {
                    const start = new Date(cursor);
                    cursor.setDate(cursor.getDate() + totalDays);
                    const end = new Date(cursor);

                    // Short label shown on Y-axis: "893 · Stress Imaging · Proto"
                    const shortLabel = `${r['Test Number']} — ${r['Test Method']} [${p.name}]`;
                    labels.push(shortLabel);

                    chartData.push({ x: [start, end], y: shortLabel });
                    bgColors.push(phaseColors[p.name]);
                    tooltipInfo.push({
                        phase:    p.name,
                        category: r['Category'],
                        method:   r['Test Method'],
                        number:   r['Test Number'],
                        start,
                        end,
                        totalDays,
                        weeks:    parseInt(p.weeks || 0),
                        days:     parseInt(p.days  || 0),
                        qty:      parseInt(p.qty   || 0),
                        comments: r['Comments'] || ''
                    });
                }
            });
        });

        // ── Empty state ──────────────────────────────────────────────────────
        if (chartData.length === 0) {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.font = '14px Inter, sans-serif';
            ctx.fillStyle = '#8e8e93';
            ctx.textAlign = 'center';
            ctx.fillText('No active test milestones to display.', canvas.width / 2, canvas.height / 2);
            return;
        }

        // ── Date range ───────────────────────────────────────────────────────
        const allDates = chartData.flatMap(d => d.x);
        const minDate  = new Date(Math.min(...allDates));
        const maxDate  = new Date(Math.max(...allDates));
        const paddingDays = scale === 'month' ? 14 : scale === 'week' ? 7 : 4;
        minDate.setDate(minDate.getDate() - paddingDays);
        maxDate.setDate(maxDate.getDate() + paddingDays);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        // ── Update summary stats ─────────────────────────────────────────────
        const statBars  = document.getElementById('gantt-stat-bars');
        const statSpan  = document.getElementById('gantt-stat-span');
        const statToday = document.getElementById('gantt-stat-today');

        const fmtDate = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        if (statBars)  statBars.textContent  = `${chartData.length} phase bar${chartData.length !== 1 ? 's' : ''} · ${rows.length} test record${rows.length !== 1 ? 's' : ''}`;
        if (statSpan)  statSpan.textContent  = `Timeline: ${fmtDate(new Date(Math.min(...allDates)))} → ${fmtDate(new Date(Math.max(...allDates)))}`;
        if (statToday) {
            const allStarts  = chartData.map(d => d.x[0].getTime());
            const allEnds    = chartData.map(d => d.x[1].getTime());
            const projectEnd = new Date(Math.max(...allEnds));
            const daysLeft   = Math.ceil((projectEnd - today) / 86400000);
            statToday.textContent = daysLeft > 0
                ? `⏱ ${daysLeft} day${daysLeft !== 1 ? 's' : ''} until project completion`
                : daysLeft === 0
                    ? '✅ Final phase ends today'
                    : `⚠️ Project ended ${Math.abs(daysLeft)} day${Math.abs(daysLeft) !== 1 ? 's' : ''} ago`;
        }

        // ── Dynamic chart height: ~32px per bar, min 300px ───────────────────
        const barHeight    = 32;
        const chartHeight  = Math.max(300, chartData.length * barHeight + 60);
        const wrapper      = document.getElementById('gantt-chart-wrapper');
        const scrollBox    = document.getElementById('gantt-chart-scroll');
        if (wrapper)   wrapper.style.height = `${chartHeight}px`;
        if (scrollBox) scrollBox.style.maxHeight = '520px';
        canvas.style.height = `${chartHeight}px`;

        // ── Custom plugins ────────────────────────────────────────────────────
        const todayLinePlugin = {
            id: 'todayLine',
            afterDatasetsDraw(chart) {
                const { ctx: c, chartArea, scales } = chart;
                const xScale = scales.x;
                const todayX = xScale.getPixelForValue(today);
                if (todayX < chartArea.left || todayX > chartArea.right) return;

                c.save();
                c.beginPath();
                c.strokeStyle = '#ff3b30';
                c.lineWidth   = 2;
                c.setLineDash([5, 4]);
                c.moveTo(todayX, chartArea.top);
                c.lineTo(todayX, chartArea.bottom);
                c.stroke();

                // "TODAY" label at top
                c.font = 'bold 10px Inter, sans-serif';
                c.fillStyle = '#ff3b30';
                c.textAlign = 'center';
                c.fillText('TODAY', todayX, chartArea.top - 4);
                c.restore();
            }
        };

        const barLabelPlugin = {
            id: 'barLabels',
            afterDatasetsDraw(chart) {
                const { ctx: c, data, chartArea } = chart;
                const meta = chart.getDatasetMeta(0);

                c.save();
                c.font      = 'bold 10px Inter, sans-serif';
                c.fillStyle = 'rgba(255,255,255,0.9)';
                c.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const info  = tooltipInfo[i];
                    const barW  = Math.abs(bar.width);
                    if (barW < 30) return; // skip too-narrow bars

                    const barX  = Math.min(bar.x, bar.base);
                    const barY  = bar.y;

                    // Clip text to the bar boundaries
                    c.save();
                    c.beginPath();
                    c.rect(Math.max(barX, chartArea.left), chartArea.top, Math.min(barW, chartArea.right - barX), chartArea.bottom - chartArea.top);
                    c.clip();

                    const text = barW > 100
                        ? `${info.phase} · ${info.totalDays}d${info.qty > 0 ? ` · Qty ${info.qty}` : ''}`
                        : info.phase;
                    c.fillText(text, barX + 8, barY);
                    c.restore();
                });
                c.restore();
            }
        };

        // ── Build Chart.js config ─────────────────────────────────────────────
        ganttChartInstance = new Chart(ctx, {
            type: 'bar',
            plugins: [todayLinePlugin, barLabelPlugin],   // ← inline plugins go HERE (not inside options)
            data: {
                labels,
                datasets: [{
                    data:            chartData,
                    backgroundColor: bgColors.map(c => c + 'dd'),    // slight transparency
                    borderColor:     bgColors,
                    borderWidth:     1.5,
                    borderRadius:    5,
                    borderSkipped:   false,
                    barPercentage:   0.72,
                    categoryPercentage: 0.88,
                    grouped: false
                }]
            },
            options: {
                indexAxis:          'y',
                responsive:         true,
                maintainAspectRatio: false,
                animation: {
                    duration: 600,
                    easing:   'easeOutQuart'
                },
                layout: {
                    padding: { top: 18, right: 12, bottom: 4, left: 0 }
                },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            title(ctx) {
                                const i = tooltipInfo[ctx[0].dataIndex];
                                return `${i.category}`;
                            },
                            label(ctx) {
                                const i = tooltipInfo[ctx.dataIndex];
                                const s = i.start.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
                                const e = i.end.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
                                const lines = [
                                    `  Phase    : ${i.phase}`,
                                    `  Method   : ${i.method}`,
                                    `  Ref #    : ${i.number}`,
                                    `  Start    : ${s}`,
                                    `  End      : ${e}`,
                                    `  Duration : ${i.weeks}w ${i.days}d (${i.totalDays} days total)`,
                                    `  Qty      : ${i.qty}`
                                ];
                                if (i.comments) lines.push(`  ⚠ Note   : ${i.comments}`);
                                return lines;
                            }
                        },
                        backgroundColor: 'rgba(18,18,20,0.96)',
                        titleColor:      '#f5f5f7',
                        bodyColor:       '#c0c0c6',
                        titleFont:       { family: 'Inter, sans-serif', size: 12, weight: 'bold' },
                        bodyFont:        { family: 'Inter, sans-serif', size: 11 },
                        borderColor:     '#3a3a3c',
                        borderWidth:     1,
                        padding:         12,
                        boxPadding:      4,
                        cornerRadius:    8,
                        displayColors:   true
                    }
                },
                scales: {
                    x: {
                        type: 'time',
                        time: {
                            unit: scale,
                            displayFormats: {
                                day:   'MMM d',
                                week:  'MMM d',
                                month: 'MMM yy'
                            },
                            tooltipFormat: 'MMM d, yyyy'
                        },
                        min: minDate,
                        max: maxDate,
                        position: 'top',   // date axis at top — like a real Gantt
                        grid: {
                            color:       'rgba(255,255,255,0.05)',
                            drawBorder:  false,
                            lineWidth:   1
                        },
                        ticks: {
                            color:      '#8e8e93',
                            font:       { family: 'Inter, sans-serif', size: 11 },
                            maxRotation: 0,
                            autoSkip:   true,
                            maxTicksLimit: scale === 'day' ? 14 : scale === 'week' ? 12 : 8
                        }
                    },
                    y: {
                        grid: {
                            color:      'rgba(255,255,255,0.03)',
                            drawBorder: false
                        },
                        ticks: {
                            color:     '#c0c0c6',
                            font:      { family: 'Inter, sans-serif', size: 11 },
                            padding:   6,
                            // Truncate long labels
                            callback(val, index) {
                                const lbl = this.getLabelForValue(val);
                                return lbl && lbl.length > 42 ? lbl.slice(0, 40) + '…' : lbl;
                            }
                        },
                        afterFit(scale) {
                            scale.width = 230; // fixed Y-axis width so bars align neatly
                        }
                    }
                }
            }
        });
    }


    // ==========================================
    // PDF EXPORT — Fixed with project context
    // ==========================================

    if (btnPdfDownload) {
        btnPdfDownload.addEventListener('click', () => {
            const projectParam = currentProject ? `?project=${encodeURIComponent(currentProject)}` : '';
            const url = `/generate-report${projectParam}`;
            
            showToast(`Generating PDF report for Project ${currentProject || 'All'}...`, 'info');
            window.location.href = url;
        });
    }

    // ==========================================
    // SHAREPOINT CLOUD SYNC — with status indicator
    // ==========================================

    function setSyncStatus(status, timeText) {
        if (!syncStatusBadge) return;
        syncStatusBadge.className = 'sync-status-badge';
        if (status === 'syncing') {
            syncStatusBadge.classList.add('syncing');
            if (syncIcon) {
                syncIcon.className = 'fa-solid fa-arrows-rotate fa-spin';
            }
        } else if (status === 'synced') {
            syncStatusBadge.classList.add('synced');
            if (syncIcon) syncIcon.className = 'fa-solid fa-arrows-rotate';
        } else if (status === 'error') {
            syncStatusBadge.classList.add('error');
            if (syncIcon) syncIcon.className = 'fa-solid fa-arrows-rotate';
        } else {
            // idle/default
            if (syncIcon) syncIcon.className = 'fa-solid fa-arrows-rotate';
        }

        if (syncLastTime && timeText) {
            syncLastTime.textContent = timeText;
            syncLastTime.classList.add('visible');
        }
    }

    if (btnSyncSharePoint) {
        btnSyncSharePoint.addEventListener('click', (e) => {
            e.preventDefault();
            showToast('Connecting to SharePoint cloud...', 'info');
            
            setSyncStatus('syncing', 'Syncing...');
            btnSyncSharePoint.style.pointerEvents = 'none';

            fetch('/api/sync-sharepoint', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            })
            .then(res => res.json())
            .then(data => {
                btnSyncSharePoint.style.pointerEvents = 'auto';
                if (data.success) {
                    const now = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
                    setSyncStatus('synced', `Last synced: ${now}`);
                    showToast(data.message, 'success');
                    // Reload all projects and current project data
                    loadProjectsDropdown();
                } else {
                    setSyncStatus('error', 'Sync failed');
                    showToast(data.error || 'Failed to sync with SharePoint.', 'error');
                }
            })
            .catch(err => {
                btnSyncSharePoint.style.pointerEvents = 'auto';
                setSyncStatus('error', 'Connection error');
                showToast('SharePoint sync failed. Connection error.', 'error');
                console.error('SharePoint Sync Error:', err);
            });
        });
    }

    // ==========================================
    // EMAIL MODAL & AUTOMATION
    // ==========================================

    // Helper to get active project from current page context
    function getActiveProjectContext() {
        const path = window.location.pathname;
        if (path.startsWith('/tables/')) {
            // Project Detail page: extract from path
            return decodeURIComponent(path.split('/').pop());
        }
        const chartSelect = document.getElementById('gantt-project-select');
        if (chartSelect && chartSelect.value) {
            // Chart Module page: get from dropdown
            return chartSelect.value;
        }
        const dashboardSelect = document.getElementById('project-select');
        if (dashboardSelect && dashboardSelect.value) {
            // Old dashboard page (if any)
            return dashboardSelect.value;
        }
        return '';
    }

    // Sidebar PDF Export
    const btnPdfDownloadSidebar = document.getElementById('btn-pdf-download-sidebar');
    if (btnPdfDownloadSidebar) {
        btnPdfDownloadSidebar.addEventListener('click', () => {
            const activeProj = getActiveProjectContext();
            if (!activeProj) {
                // Fallback: export consolidated report for all projects
                showToast('No project selected — exporting full consolidated PDF report...', 'info');
                triggerConsolidatedPdfDownload();
                return;
            }
            triggerPdfDownload(activeProj);
        });
    }

    // ==========================================
    // EXCEL FULL REPORT EXPORT
    // ==========================================
    const btnExcelExport = document.getElementById('btn-excel-export');
    if (btnExcelExport) {
        btnExcelExport.addEventListener('click', () => {
            const icon = btnExcelExport.querySelector('i');
            const origClass = icon ? icon.className : 'fa-solid fa-file-excel';
            if (icon) icon.className = 'fa-solid fa-spinner fa-spin';
            btnExcelExport.disabled = true;

            showToast('Preparing Excel report for download...', 'info');

            // Use a hidden link to trigger the download
            fetch('/api/export-excel')
                .then(res => {
                    if (!res.ok) return res.json().then(d => { throw new Error(d.error || 'Export failed'); });
                    return res.blob();
                })
                .then(blob => {
                    const today = new Date().toISOString().slice(0,10);
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `Full_Operations_Report_${today}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
                    showToast('Excel report downloaded successfully!', 'success');
                })
                .catch(err => {
                    showToast(`Excel export failed: ${err.message}`, 'error');
                    console.error('Excel Export Error:', err);
                })
                .finally(() => {
                    if (icon) icon.className = origClass;
                    btnExcelExport.disabled = false;
                });
        });
    }

    // ==========================================
    // CONSOLIDATED PDF EXPORT (internal helper)
    // ==========================================
    function triggerConsolidatedPdfDownload() {
        const exportProgressModal = document.getElementById('export-progress-modal');
        const exportProgressBar = document.getElementById('export-progress-bar');
        const exportProgressMessage = document.getElementById('export-progress-message');
        const exportProgressPercent = document.getElementById('export-progress-percent');

        const progressTitle = document.getElementById('export-progress-title');
        const progressDesc = document.getElementById('export-progress-desc');
        if (progressTitle) progressTitle.textContent = 'Compiling Full PDF Report';
        if (progressDesc) progressDesc.textContent = 'Generating timelines and defect summaries for all projects.';

        if (exportProgressBar) exportProgressBar.style.width = '0%';
        if (exportProgressPercent) exportProgressPercent.textContent = '0%';
        if (exportProgressMessage) exportProgressMessage.textContent = 'Initializing...';
        if (exportProgressModal) exportProgressModal.classList.add('active');

        let pollInterval = null;
        const cleanupExportUI = () => {
            if (exportProgressModal) exportProgressModal.classList.remove('active');
        };

        const pollExportProgress = (taskId) => {
            pollInterval = setInterval(() => {
                fetch(`/api/export-consolidated/progress/${taskId}`)
                    .then(res => { if (!res.ok) throw new Error('Task state unavailable'); return res.json(); })
                    .then(state => {
                        const progressVal = state.progress || 0;
                        const msg = state.message || 'Processing...';
                        if (exportProgressBar) exportProgressBar.style.width = `${progressVal}%`;
                        if (exportProgressPercent) exportProgressPercent.textContent = `${progressVal}%`;
                        if (exportProgressMessage) exportProgressMessage.textContent = msg;
                        if (state.status === 'completed') {
                            clearInterval(pollInterval);
                            window.location.href = `/api/export-consolidated/download/${taskId}/Consolidated_Project_Report.pdf`;
                            showToast('Consolidated PDF downloaded!', 'success');
                            setTimeout(cleanupExportUI, 800);
                        } else if (state.status === 'failed') {
                            clearInterval(pollInterval);
                            showToast(state.message || 'PDF generation failed', 'error');
                            cleanupExportUI();
                        }
                    })
                    .catch(err => {
                        clearInterval(pollInterval);
                        showToast(`Export tracking failed: ${err.message}`, 'error');
                        cleanupExportUI();
                    });
            }, 600);
        };

        fetch('/api/export-consolidated/start', { method: 'POST' })
            .then(r => { if (!r.ok) throw new Error('Failed to start export'); return r.json(); })
            .then(data => pollExportProgress(data.task_id))
            .catch(err => {
                showToast(`PDF export failed: ${err.message}`, 'error');
                cleanupExportUI();
            });
    }


    // ==========================================
    // EMAIL MODAL & AUTOMATION
    // ==========================================

    if (btnEmailModalTrigger) {
        btnEmailModalTrigger.addEventListener('click', () => {
            const activeProj = getActiveProjectContext();
            const projectInput = document.getElementById('email-project');
            if (projectInput) {
                projectInput.value = activeProj;
            }

            emailModal.classList.add('active');
            const formattedDate = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
            let subjectText = `Daily Test Operations Report - ${formattedDate}`;
            if (activeProj) {
                subjectText = `Project ${activeProj} Test Operations Report - ${formattedDate}`;
                document.querySelector('.attachment-details span').textContent = `Gantt_Report_${activeProj}.pdf`;
            } else {
                document.querySelector('.attachment-details span').textContent = `Operations_Report.pdf`;
            }
            document.getElementById('email-subject').value = subjectText;
        });
    }

    function closeEmailModal() {
        emailModal.classList.remove('active');
        emailForm.reset();
    }

    if (btnCloseModal) btnCloseModal.addEventListener('click', closeEmailModal);
    if (btnCancelEmail) btnCancelEmail.addEventListener('click', closeEmailModal);

    if (emailForm) {
        emailForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const recipient = document.getElementById('email-to').value.trim();
            const subject = document.getElementById('email-subject').value.trim();
            const body = document.getElementById('email-body').value.trim();
            const projectVal = document.getElementById('email-project').value;

            const btnSubmit = emailForm.querySelector('button[type="submit"]');
            btnSubmit.disabled = true;
            btnSubmit.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Dispatched...';

            const formData = new FormData();
            formData.append('recipient', recipient);
            formData.append('subject', subject);
            formData.append('body', body);
            if (projectVal) {
                formData.append('project', projectVal);
            }

            fetch('/send-email', {
                method: 'POST',
                body: formData
            })
            .then(res => res.json())
            .then(data => {
                if (data.success) {
                    closeEmailModal();
                    showToast(data.message, 'success');
                } else {
                    showToast(data.message || 'Unable to send email.', 'error');
                }

                const logModal = document.getElementById('email-log-modal');
                const logContent = document.getElementById('email-log-content');
                if (data.mode === 'simulation' && logModal && logContent) {
                    const timestamp = new Date().toISOString();
                    const pdfName = projectVal ? `Gantt_Report_${projectVal}.pdf` : 'Operations_Report.pdf';
                    logContent.textContent = `[SMTP SIMULATOR - ${timestamp}]
Connecting to mail server... SIMULATED
FROM: ${document.getElementById('email-to').value}
TO: ${recipient}
SUBJECT: ${subject}
ATTACHMENTS: ${pdfName} (Generated)
BODY:
${body}
--------------------------------------------------
Mail Status: Queued and Sent successfully (Simulated)`;
                    logModal.style.display = 'flex';
                } else if (logModal) {
                    logModal.style.display = 'none';
                }
            })
            .catch(err => {
                showToast('Failed to dispatch report email.', 'error');
                console.error(err);
            })
            .finally(() => {
                btnSubmit.disabled = false;
                btnSubmit.innerHTML = '<i class="fa-solid fa-paper-plane"></i> Send Email';
            });
        });
    }

    // Close Log Modal
    const btnCloseLogModal = document.getElementById('btn-close-log-modal');
    const btnCloseLogOk = document.getElementById('btn-close-log-ok');
    if (btnCloseLogModal) btnCloseLogModal.addEventListener('click', () => document.getElementById('email-log-modal').style.display = 'none');
    if (btnCloseLogOk) btnCloseLogOk.addEventListener('click', () => document.getElementById('email-log-modal').style.display = 'none');

    // ==========================================
    // CHATBOT INTERACTION
    // ==========================================

    if (btnChatbotBubble) {
        btnChatbotBubble.addEventListener('click', () => {
            chatbotPanel.classList.toggle('active');
            btnChatbotBubble.classList.toggle('active');
        });
    }

    if (btnMinimizeChat) {
        btnMinimizeChat.addEventListener('click', () => {
            chatbotPanel.classList.remove('active');
            btnChatbotBubble.classList.remove('active');
        });
    }

    if (chatForm) {
        chatForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const query = chatInput.value.trim();
            if (!query) return;

            appendMessage(query, 'user');
            chatInput.value = '';

            // Check for local report generation commands before hitting API
            const lowerQuery = query.toLowerCase();
            const generateMatch = lowerQuery.match(/generate\s+(?:a\s+)?(?:pdf\s+)?report\s+(?:for\s+)?(?:project\s+)?(\S+)/i)
                || lowerQuery.match(/create\s+(?:a\s+)?(?:pdf\s+)?report\s+(?:for\s+)?(?:project\s+)?(\S+)/i)
                || lowerQuery.match(/export\s+(?:pdf\s+)?(?:for\s+)?(?:project\s+)?(\S+)/i);

            if (lowerQuery.includes('generate') || lowerQuery.includes('create report') || lowerQuery.includes('export pdf')) {
                // Try to extract project name
                let targetProject = currentProject;
                if (generateMatch && generateMatch[1]) {
                    targetProject = generateMatch[1].toUpperCase();
                }

                // Simulate typing then trigger download
                const typingId = appendTypingIndicator();
                setTimeout(() => {
                    removeTypingIndicator(typingId);
                    appendMessage(`📄 Generating PDF report for **Project ${targetProject}**... Downloading now!`, 'bot');
                    triggerPdfDownload(targetProject);
                }, 1200);
                return;
            }

            // Simulate Bot typing
            const typingId = appendTypingIndicator();

            fetch('/api/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ query: query, project: currentProject })
            })
            .then(res => res.json())
            .then(data => {
                removeTypingIndicator(typingId);
                appendMessage(data.response, 'bot');

                // If AI response includes a report generation trigger
                if (data.generate_report && data.report_project) {
                    setTimeout(() => triggerPdfDownload(data.report_project), 800);
                }
            })
            .catch(err => {
                removeTypingIndicator(typingId);
                appendMessage("Sorry, I'm having trouble connecting right now. You can still use the **Export PDF Report** button in the sidebar to generate reports.", 'bot');
                console.error('Chat Error:', err);
            });
        });
    }

    // Trigger PDF download programmatically
    function triggerPdfDownload(projectName) {
        const comment = sessionStorage.getItem('pdf_comment_' + projectName) || '';
        const exportProgressModal = document.getElementById('export-progress-modal');
        const exportProgressBar = document.getElementById('export-progress-bar');
        const exportProgressMessage = document.getElementById('export-progress-message');
        const exportProgressPercent = document.getElementById('export-progress-percent');
        
        // Update Modal Title and Description for Project Export
        const progressTitle = document.getElementById('export-progress-title');
        const progressDesc = document.getElementById('export-progress-desc');
        if (progressTitle) progressTitle.textContent = `Compiling Report for Project ${projectName}`;
        if (progressDesc) progressDesc.textContent = `Generating timeline, dashboard, and tables for Project ${projectName}.`;
        
        // Reset and Show Progress Modal
        if (exportProgressBar) exportProgressBar.style.width = '0%';
        if (exportProgressPercent) exportProgressPercent.textContent = '0%';
        if (exportProgressMessage) exportProgressMessage.textContent = 'Initializing...';
        if (exportProgressModal) exportProgressModal.classList.add('active');
        
        let pollInterval = null;
        
        const cleanupExportUI = () => {
            if (exportProgressModal) exportProgressModal.classList.remove('active');
        };
        
        const triggerDownload = (taskId) => {
            const filename = `Gantt_Report_${projectName}.pdf`;
            window.location.href = `/api/export-project/download/${taskId}/${encodeURIComponent(filename)}`;
            showToast(`PDF report for Project ${projectName} downloaded successfully!`, 'success');
            setTimeout(() => {
                cleanupExportUI();
            }, 800);
        };
        
        const pollExportProgress = (taskId) => {
            pollInterval = setInterval(() => {
                fetch(`/api/export-project/progress/${taskId}`)
                    .then(res => {
                        if (!res.ok) throw new Error('Task state unavailable');
                        return res.json();
                    })
                    .then(state => {
                        const progressVal = state.progress || 0;
                        const msg = state.message || 'Processing...';
                        
                        if (exportProgressBar) exportProgressBar.style.width = `${progressVal}%`;
                        if (exportProgressPercent) exportProgressPercent.textContent = `${progressVal}%`;
                        if (exportProgressMessage) exportProgressMessage.textContent = msg;
                        
                        if (state.status === 'completed') {
                            clearInterval(pollInterval);
                            triggerDownload(taskId);
                        } else if (state.status === 'failed') {
                            clearInterval(pollInterval);
                            throw new Error(state.message || 'Server task failed');
                        }
                    })
                    .catch(err => {
                        clearInterval(pollInterval);
                        showToast(`Export tracking failed: ${err.message}`, 'error');
                        cleanupExportUI();
                    });
            }, 600);
        };
        
        // Start export task
        fetch('/api/export-project/start', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ project: projectName, comment: comment })
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('Failed to start PDF export task');
            }
            return response.json();
        })
        .then(data => {
            pollExportProgress(data.task_id);
        })
        .catch(err => {
            showToast(`PDF export failed: ${err.message}`, 'error');
            console.error('PDF Export Error:', err);
            cleanupExportUI();
        });
    }

    window.triggerPdfDownload = triggerPdfDownload;

    function appendMessage(text, sender) {
        const msg = document.createElement('div');
        msg.className = `chat-message ${sender}`;
        // Support basic markdown-like bold with **
        const formatted = text
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\n/g, '<br>');
        msg.innerHTML = `
            <div class="message-bubble">
                <p>${formatted}</p>
            </div>
        `;
        chatMessages.appendChild(msg);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    function appendTypingIndicator() {
        const id = 'typing-' + Date.now();
        const msg = document.createElement('div');
        msg.className = 'chat-message bot typing';
        msg.id = id;
        msg.innerHTML = `
            <div class="message-bubble">
                <div class="typing-dots">
                    <span></span><span></span><span></span>
                </div>
            </div>
        `;
        chatMessages.appendChild(msg);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        return id;
    }

    function removeTypingIndicator(id) {
        const el = document.getElementById(id);
        if (el) el.remove();
    }
});
