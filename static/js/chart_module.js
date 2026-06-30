document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const projectSelect = document.getElementById('gantt-project-select');
    const categorySelect = document.getElementById('gantt-category-select');
    const scaleButtons = document.querySelectorAll('.scale-btn');
    const btnExportPdf = document.getElementById('btn-export-pdf');
    const emptyState = document.getElementById('gantt-empty-state');
    const loadingState = document.getElementById('gantt-loading');
    const canvasWrap = document.getElementById('gantt-canvas-wrap');
    const ctx = document.getElementById('ganttChart').getContext('2d');

    const statRows = document.getElementById('stat-rows');
    const statCats = document.getElementById('stat-cats');
    const statSpan = document.getElementById('stat-span');
    const statBars = document.getElementById('stat-bars');
    const statDefects = document.getElementById('stat-defects');

    // Chart Instance
    let ganttChartInstance = null;

    // State variables
    let allProjects = [];
    let currentProjectName = '';
    let currentCategory = 'all';
    let currentScale = 'week'; // 'day' | 'week' | 'month' | 'year'
    let rawRecords = []; // All records for the selected project

    // Initialize Page
    initGanttPage();

    function initGanttPage() {
        // Fetch projects list
        fetchProjects();

        // Project change handler
        projectSelect.addEventListener('change', (e) => {
            currentProjectName = e.target.value;
            if (currentProjectName) {
                // Enable PDF export link
                btnExportPdf.href = `/generate-report?project=${encodeURIComponent(currentProjectName)}`;
                btnExportPdf.style.pointerEvents = 'auto';
                btnExportPdf.style.opacity = '1';
                loadProjectData(currentProjectName);
            } else {
                // Disable PDF export
                btnExportPdf.href = '#';
                btnExportPdf.style.pointerEvents = 'none';
                btnExportPdf.style.opacity = '0.4';
                resetToEmptyState();
            }
        });

        // Category change handler
        categorySelect.addEventListener('change', (e) => {
            currentCategory = e.target.value;
            if (rawRecords.length > 0) {
                renderGanttChart();
            }
        });

        // Scale buttons handler
        scaleButtons.forEach(btn => {
            btn.addEventListener('click', (e) => {
                scaleButtons.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                currentScale = btn.getAttribute('data-scale');
                if (rawRecords.length > 0) {
                    renderGanttChart();
                }
            });
        });
    }

    // Fetch active projects to populate the select dropdown
    function fetchProjects() {
        fetch('/api/projects?status=Active')
            .then(res => res.json())
            .then(projects => {
                allProjects = projects;
                projectSelect.innerHTML = '<option value="">— Select a project —</option>';
                if (projects.length === 0) {
                    projectSelect.innerHTML = '<option value="">No active projects found</option>';
                    return;
                }
                projects.forEach(p => {
                    const option = document.createElement('option');
                    option.value = p.name;
                    option.textContent = `Project ${p.name}`;
                    projectSelect.appendChild(option);
                });
            })
            .catch(err => {
                console.error("Failed to load projects:", err);
                projectSelect.innerHTML = '<option value="">Failed to load projects</option>';
            });
    }

    // Load data for selected project
    function loadProjectData(projectName) {
        emptyState.style.display = 'none';
        canvasWrap.style.display = 'none';
        loadingState.style.display = 'block';

        fetch(`/api/projects/${encodeURIComponent(projectName)}/rows`)
            .then(res => res.json())
            .then(data => {
                loadingState.style.display = 'none';
                rawRecords = Array.isArray(data) ? data : (data.rows || []);

                if (rawRecords.length === 0) {
                    resetToEmptyState("No Data Available", `Project ${projectName} has no test method records.`);
                    return;
                }

                // Populate categories dropdown dynamically
                populateCategories(rawRecords);

                // Show canvas and render chart
                canvasWrap.style.display = 'block';
                renderGanttChart();
            })
            .catch(err => {
                console.error("Error loading project data:", err);
                loadingState.style.display = 'none';
                resetToEmptyState("Failed to Load Data", "There was an error fetching records for this project.");
            });
    }

    function populateCategories(records) {
        const categories = new Set();
        records.forEach(r => {
            if (r.Category) categories.add(r.Category);
        });

        // Store selected value if possible
        const prevVal = categorySelect.value;
        categorySelect.innerHTML = '<option value="all">All Categories</option>';
        categories.forEach(cat => {
            const option = document.createElement('option');
            option.value = cat;
            option.textContent = cat;
            categorySelect.appendChild(option);
        });

        // Restore value if still exists
        if (categories.has(prevVal)) {
            categorySelect.value = prevVal;
            currentCategory = prevVal;
        } else {
            categorySelect.value = 'all';
            currentCategory = 'all';
        }
    }

    function resetToEmptyState(title = "No Project Selected", text = "Select a project above to render its AI-calculated Gantt chart.") {
        emptyState.style.display = 'block';
        canvasWrap.style.display = 'none';
        loadingState.style.display = 'none';

        emptyState.querySelector('h3').textContent = title;
        emptyState.querySelector('p').textContent = text;

        statRows.textContent = '—';
        statCats.textContent = '—';
        statSpan.textContent = '—';
        statBars.textContent = '—';
        statDefects.textContent = '—';

        if (ganttChartInstance) {
            ganttChartInstance.destroy();
            ganttChartInstance = null;
        }
    }

    // AI auto-calculation and formatting of Gantt data
    let renderTimeout = null;
    function renderGanttChart() {
        if (renderTimeout) {
            clearTimeout(renderTimeout);
        }

        // 1. Filter records by Category
        let filteredRecords = rawRecords;
        if (currentCategory !== 'all') {
            filteredRecords = rawRecords.filter(r => r.Category === currentCategory);
        }

        if (filteredRecords.length === 0) {
            canvasWrap.style.display = 'none';
            emptyState.style.display = 'block';
            emptyState.querySelector('h3').textContent = "No Matches";
            emptyState.querySelector('p').textContent = `No records match the category: "${currentCategory}".`;
            if (ganttChartInstance) {
                ganttChartInstance.destroy();
                ganttChartInstance = null;
            }
            return;
        } else {
            canvasWrap.style.display = 'block';
            emptyState.style.display = 'none';
        }

        // 2. Auto-calculate start/end dates for each phase
        // Chained phases: Proto -> DVT -> EVT -> PVT
        const chartData = [];
        let minDate = null;
        let maxDate = null;
        let totalBarsCount = 0;
        let totalDefectsCount = 0;

        filteredRecords.forEach((r, idx) => {
            const label = `[${r.Category}] ${r['Test Method']} (${r['Test Number']})`;
            let baseStart = new Date(r['Start Date'] || new Date());
            if (isNaN(baseStart.getTime())) {
                baseStart = new Date();
            }

            totalDefectsCount += parseInt(r['Defect Qty'] || 0);

            // Phase definitions
            const phases = [
                { name: 'Proto', weeks: parseInt(r['Proto Weeks'] || 0), days: parseInt(r['Proto Days'] || 0), color: '#8E44AD' },
                { name: 'DVT', weeks: parseInt(r['DVT Weeks'] || 0), days: parseInt(r['DVT Days'] || 0), color: '#2980B9' },
                { name: 'EVT', weeks: parseInt(r['EVT Weeks'] || 0), days: parseInt(r['EVT Days'] || 0), color: '#27AE60' },
                { name: 'PVT', weeks: parseInt(r['PVT Weeks'] || 0), days: parseInt(r['PVT Days'] || 0), color: '#D35400' }
            ];

            phases.forEach(phase => {
                const durationDays = phase.weeks * 7 + phase.days;
                if (durationDays > 0) {
                    const start = new Date(baseStart.getTime());
                    const end = new Date(baseStart.getTime() + durationDays * 24 * 60 * 60 * 1000);

                    chartData.push({
                        y: label,
                        x: [start, end],
                        phase: phase.name,
                        color: phase.color,
                        defects: r['Defect Qty'] || 0,
                        comment: r['Comments'] || ''
                    });

                    // Update min/max dates
                    if (!minDate || start < minDate) minDate = start;
                    if (!maxDate || end > maxDate) maxDate = end;
                    totalBarsCount++;

                    // Move baseStart forward for the next phase (chaining)
                    baseStart = end;
                }
            });
        });

        // 3. Update Statistics Panel
        statRows.textContent = filteredRecords.length;
        const uniqueCats = new Set(filteredRecords.map(r => r.Category));
        statCats.textContent = uniqueCats.size;
        statBars.textContent = totalBarsCount;
        statDefects.textContent = totalDefectsCount;

        if (minDate && maxDate) {
            const diffTime = Math.abs(maxDate - minDate);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            statSpan.textContent = `${diffDays} Days (~${Math.ceil(diffDays / 7)} Wks)`;
        } else {
            statSpan.textContent = '—';
        }

        // Adjust Canvas parent container height dynamically based on number of bars
        const uniqueYLabels = Array.from(new Set(chartData.map(d => d.y)));
        const dynamicHeight = Math.max(350, uniqueYLabels.length * 65 + 100);
        canvasWrap.style.height = `${dynamicHeight}px`;

        // 4. Configure Chart.js Options & Scale
        let timeUnit = 'week';
        let stepSize = 1;
        if (currentScale === 'day') {
            timeUnit = 'day';
            stepSize = 1;
        } else if (currentScale === 'week') {
            timeUnit = 'week';
            stepSize = 1;
        } else if (currentScale === 'month') {
            timeUnit = 'month';
            stepSize = 1;
        } else if (currentScale === 'year') {
            timeUnit = 'year';
            stepSize = 1;
        }

        // Build dataset for Chart.js
        const datasetData = chartData.map(item => {
            return {
                x: item.x,
                y: item.y,
                phase: item.phase,
                color: item.color,
                defects: item.defects,
                comment: item.comment
            };
        });

        // Delay Chart.js instantiation slightly to allow browser layout/reflow
        renderTimeout = setTimeout(() => {
            if (ganttChartInstance) {
                ganttChartInstance.destroy();
            }

            // 5. Create Chart Instance
            ganttChartInstance = new Chart(ctx, {
                type: 'bar',
                data: {
                    datasets: [{
                        label: 'Timeline',
                        data: datasetData,
                        backgroundColor: (context) => {
                            const raw = context.raw;
                            return raw ? raw.color : '#999999';
                        },
                        borderRadius: 4,
                        borderSkipped: false,
                        barPercentage: 0.7,
                        categoryPercentage: 0.8
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                title: (context) => {
                                    return context[0].raw.y;
                                },
                                label: (context) => {
                                    const raw = context.raw;
                                    const startStr = raw.x[0].toISOString().split('T')[0];
                                    const endStr = raw.x[1].toISOString().split('T')[0];
                                    let lines = [
                                        `Phase: ${raw.phase}`,
                                        `Duration: ${startStr} to ${endStr}`,
                                        `Defects: ${raw.defects}`
                                    ];
                                    if (raw.comment) {
                                        lines.push(`Note: ${raw.comment}`);
                                    }
                                    return lines;
                                }
                            }
                        }
                    },
                    scales: {
                        x: {
                            type: 'time',
                            time: {
                                unit: timeUnit,
                                stepSize: stepSize,
                                displayFormats: {
                                    day: 'MMM dd',
                                    week: "'Wk' w (MMM dd)",
                                    month: 'MMM yyyy',
                                    year: 'yyyy'
                                }
                            },
                            grid: {
                                color: 'rgba(255, 255, 255, 0.08)'
                            },
                            ticks: {
                                color: '#a0a0a0',
                                font: { size: 11 }
                            },
                            title: {
                                display: true,
                                text: 'Project Timeline Progression',
                                color: '#808080',
                                font: { size: 12, weight: 'bold' }
                            }
                        },
                        y: {
                            grid: { display: false },
                            ticks: {
                                color: '#e0e0e0',
                                font: { size: 11, weight: '600' }
                            }
                        }
                    }
                },
                plugins: [{
                    id: 'todayLine',
                    afterDatasetsDraw(chart) {
                        const { ctx, scales: { x, y } } = chart;
                        const today = new Date();
                        if (today >= x.min && today <= x.max) {
                            const todayX = x.getPixelForValue(today);
                            ctx.save();
                            ctx.beginPath();
                            ctx.strokeStyle = '#ff3b30';
                            ctx.lineWidth = 2;
                            ctx.setLineDash([5, 5]);
                            ctx.moveTo(todayX, y.top);
                            ctx.lineTo(todayX, y.bottom);
                            ctx.stroke();

                            // Add "Today" label
                            ctx.fillStyle = '#ff3b30';
                            ctx.font = '10px sans-serif';
                            ctx.fillText('Today', todayX + 5, y.top + 15);
                            ctx.restore();
                        }
                    }
                }]
            });
        }, 50);
    }
});
