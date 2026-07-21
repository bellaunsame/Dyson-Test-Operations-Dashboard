document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const projectSelect = document.getElementById('gantt-project-select');
    const laboratorySelect = document.getElementById('gantt-laboratory-select');
    const categorySelect = document.getElementById('gantt-category-select');
    const testMethodSelect = document.getElementById('gantt-testmethod-select');
    const scaleButtons = document.querySelectorAll('.scale-btn');
    const btnSubmitComment = document.getElementById('btn-submit-comment');
    const pdfComment = document.getElementById('pdf-comment');
    const emptyState = document.getElementById('gantt-empty-state');
    const loadingState = document.getElementById('gantt-loading');
    const canvasWrap = document.getElementById('gantt-canvas-wrap');
    const ctx = document.getElementById('ganttChart').getContext('2d');

    const statRows = document.getElementById('stat-rows');
    const statCats = document.getElementById('stat-cats');
    const statDefects = document.getElementById('stat-defects');
    const testMethodList = document.getElementById('test-method-list');
    const statsRow = document.querySelector('.gantt-stats-row');

    // Chart Instance
    let ganttChartInstance = null;

    // State variables
    let allProjects = [];
    let currentProjectName = '';
    let currentLaboratory = 'all';
    let currentCategory = 'all';
    let currentTestMethod = 'all';
    let selectedTestMethods = new Set();
    let customStartDate = '';
    let currentScale = 'week'; // 'day' | 'week' | 'month' | 'year'
    let spanStart = ''; // user-chosen timeline span start date (YYYY-MM-DD)
    let spanEnd = '';   // user-chosen timeline span end date (YYYY-MM-DD)
    let rawRecords = []; // All records for the selected project
    let masterDataLabs = [];
    let labTestNumbersMap = {};
    let labTestMethodsMap = {};

    // Initialize Page
    initGanttPage();

    function initGanttPage() {
        // Fetch master data laboratories and projects list
        fetchMasterDataLaboratories();
        fetchProjects();
        loadLaboratories();

        // Project change handler
        projectSelect.addEventListener('change', (e) => {
            currentProjectName = e.target.value;
            if (currentProjectName) {
                if (btnSubmitComment) {
                    btnSubmitComment.style.pointerEvents = 'auto';
                    btnSubmitComment.style.opacity = '1';
                }
                const savedComment = sessionStorage.getItem('pdf_comment_' + currentProjectName) || '';
                if (pdfComment) pdfComment.value = savedComment;
                loadProjectData(currentProjectName);
            } else {
                if (pdfComment) pdfComment.value = '';
                if (btnSubmitComment) {
                    btnSubmitComment.style.pointerEvents = 'none';
                    btnSubmitComment.style.opacity = '0.4';
                }
                resetToEmptyState();
            }
        });

        // Export PDF handler
        if (btnSubmitComment) {
            btnSubmitComment.addEventListener('click', () => {
                if (!currentProjectName) {
                    if (window.showToast) {
                        window.showToast('Select a project before exporting the PDF.', 'info');
                    }
                    return;
                }

                const commentValue = pdfComment ? (pdfComment.value || '').trim() : '';
                sessionStorage.setItem('pdf_comment_' + currentProjectName, commentValue);

                const methodParam = selectedTestMethods.size > 0 ? Array.from(selectedTestMethods).join(',') : currentTestMethod;
                const params = new URLSearchParams({ 
                    project: currentProjectName,
                    laboratory: currentLaboratory,
                    category: currentCategory,
                    test_method: methodParam,
                    scale: currentScale
                 });
                if (commentValue) {
                    params.set('comment', commentValue);
                }
                if (customStartDate) {
                    params.set('custom_start_date', customStartDate);
                }
                if (spanStart) {
                    params.set('span_start', spanStart);
                }
                if (spanEnd) {
                    params.set('span_end', spanEnd);
                }

                if (window.showToast) {
                    window.showToast(`Exporting PDF for ${currentProjectName}...`, 'info');
                }
                if (window.triggerPdfDownload) {
                    window.triggerPdfDownload(currentProjectName, currentCategory, currentScale, methodParam, customStartDate, currentLaboratory, spanStart, spanEnd);
                } else {
                    const exportUrl = `/generate-report?${params.toString()}`;
                    window.location.assign(exportUrl);
                }
            });
        }

        // Laboratory change handler
        if (laboratorySelect) {
            laboratorySelect.addEventListener('change', (e) => {
                currentLaboratory = e.target.value;
                currentTestMethod = 'all';
                selectedTestMethods.clear();
                if (testMethodSelect) testMethodSelect.value = 'all';

                applyFiltersAndRender();
            });
        }

        // Category change handler
        if (categorySelect) {
            categorySelect.addEventListener('change', (e) => {
                currentCategory = e.target.value;
                currentTestMethod = 'all';
                selectedTestMethods.clear();
                if (testMethodSelect) testMethodSelect.value = 'all';

                applyFiltersAndRender();
            });
        }



        if (testMethodSelect) {
            testMethodSelect.addEventListener('change', (e) => {
                currentTestMethod = e.target.value;
                selectedTestMethods.clear();
                if (currentTestMethod !== 'all' && currentTestMethod !== 'multiple') {
                    selectedTestMethods.add(currentTestMethod);
                }

                updateChipActiveStates();
                renderGanttChart();
            });
        }

        // Scale buttons handler
        if (scaleButtons) {
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

        // Custom start date change handler
        const ganttStartDateInput = document.getElementById('gantt-start-date');
        if (ganttStartDateInput) {
            ganttStartDateInput.addEventListener('change', (e) => {
                customStartDate = e.target.value;
                if (rawRecords.length > 0) {
                    renderGanttChart();
                }
            });
        }

        // Span date range inputs
        const spanStartInput = document.getElementById('gantt-span-start');
        const spanEndInput = document.getElementById('gantt-span-end');

        if (spanStartInput) {
            spanStartInput.addEventListener('change', (e) => {
                spanStart = e.target.value;
                if (rawRecords.length > 0) {
                    renderGanttChart();
                }
            });
        }
        if (spanEndInput) {
            spanEndInput.addEventListener('change', (e) => {
                spanEnd = e.target.value;
                if (rawRecords.length > 0) {
                    renderGanttChart();
                }
            });
        }

        // Scroll to Gantt Chart handler
        const btnScrollToGantt = document.getElementById('btn-scroll-to-gantt');
        if (btnScrollToGantt) {
            btnScrollToGantt.addEventListener('click', () => {
                const chartCanvas = document.getElementById('ganttChart');
                if (chartCanvas) {
                    chartCanvas.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            });
        }
    }

    function loadLaboratories() {
        fetch('/api/laboratories')
            .then(res => res.json())
            .then(labs => {
                if (!laboratorySelect) return;

                laboratorySelect.innerHTML =
                    '<option value="all">All Laboratories</option>';

                labs.forEach(lab => {
                    const option = document.createElement('option');
                    option.value = lab;
                    option.textContent = lab;
                    laboratorySelect.appendChild(option);
                });
            })
            .catch(err => {
                console.error(
                    'Failed to load laboratories:',
                    err
                );
            });
    }

    // Fetch active projects to populate the select dropdown
    function fetchProjects() {
        fetch('/api/projects?status=Active')
            .then(res => res.json())
            .then(projects => {
                allProjects = projects;
                projectSelect.innerHTML = '<option value="">-- Select a project --</option>';
                if (projects.length === 0) {
                    projectSelect.innerHTML = '<option value="">No active projects found</option>';
                    return;
                }
                projects.forEach(p => {
                    const option = document.createElement('option');
                    option.value = p.name;
                    option.textContent = p.name;
                    projectSelect.appendChild(option);
                });
                // Auto-select based on URL parameter or fallback to first project
                if (!currentProjectName && projects.length > 0) {
                    const urlParams = new URLSearchParams(window.location.search);
                    const urlProj = urlParams.get('project');
                    if (urlProj && projects.some(p => p.name === urlProj)) {
                        projectSelect.value = urlProj;
                        currentProjectName = urlProj;
                    } else {
                        projectSelect.value = projects[0].name;
                        currentProjectName = projects[0].name;
                    }
                    projectSelect.dispatchEvent(new Event('change'));
                }
            })
            .catch(err => {
                console.error("Failed to load projects:", err);
                projectSelect.innerHTML = '<option value="">Failed to load projects</option>';
            });
    }

    function fetchMasterDataLaboratories() {
        fetch('/api/master-data/laboratories')
            .then(res => res.json())
            .then(data => {
                masterDataLabs = data.laboratories || [];
                labTestNumbersMap = data.lab_test_numbers || {};
                labTestMethodsMap = data.lab_test_methods || {};

                if (laboratorySelect) {
                    laboratorySelect.innerHTML = '<option value="all">All Laboratories</option>';
                    masterDataLabs.forEach(lab => {
                        const option = document.createElement('option');
                        option.value = lab;
                        option.textContent = lab;
                        laboratorySelect.appendChild(option);
                    });
                }
            })
            .catch(err => {
                console.error("Failed to load master data laboratories:", err);
            });
    }

    function getFilteredRecordsForMethods() {
        let records = rawRecords;

        if (currentLaboratory !== 'all') {
            const allowedNums = new Set((labTestNumbersMap[currentLaboratory] || []).map(x => String(x).trim()));
            const allowedMethods = new Set((labTestMethodsMap[currentLaboratory] || []).map(x => String(x).trim()));

            records = records.filter(r => {
                const numMatch = r['Test Number'] && allowedNums.has(String(r['Test Number']).trim());
                const methodMatch = r['Test Method'] && allowedMethods.has(String(r['Test Method']).trim());
                const catMatch = r['Category'] && String(r['Category']).trim().toLowerCase() === currentLaboratory.toLowerCase();
                const labMatch = r['Laboratory'] && String(r['Laboratory']).trim().toLowerCase() === currentLaboratory.toLowerCase();
                return numMatch || methodMatch || catMatch || labMatch;
            });
        }

        if (currentCategory !== 'all') {
            records = records.filter(r => r.Category === currentCategory);
        }

        return records;
    }

    function applyFiltersAndRender() {
        const filteredForMethods = getFilteredRecordsForMethods();
        populateTestMethods(filteredForMethods);
        renderTestMethods(filteredForMethods);
        renderGanttChart();
    }

    // Load data for selected project
    function loadProjectData(projectName) {
        emptyState.style.display = 'none';
        canvasWrap.style.display = 'none';
        if (statsRow) statsRow.style.display = 'none';
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

                currentLaboratory = 'all';
                if (laboratorySelect) laboratorySelect.value = 'all';
                currentCategory = 'all';
                currentTestMethod = 'all';
                selectedTestMethods.clear();
                customStartDate = '';
                const ganttStartDateInput = document.getElementById('gantt-start-date');
                if (ganttStartDateInput) ganttStartDateInput.value = '';

                populateCategories(rawRecords);
                applyFiltersAndRender();
            })
            .catch(err => {
                console.error("Error loading project data:", err);
                loadingState.style.display = 'none';
                resetToEmptyState("Failed to Load Data", "There was an error fetching records for this project.");
            });
    }

    function renderTestMethods(records) {
        if (!testMethodList) return;

        testMethodList.innerHTML = '';
        testMethodList.style.display = 'block';

        const uniqueMethods = new Set();
        records.forEach(r => {
            if (r['Test Method']) {
                uniqueMethods.add(r['Test Method']);
            }
        });

        const sortedMethods = Array.from(uniqueMethods).sort();

        if (sortedMethods.length === 0) {
            testMethodList.innerHTML = '<div style="color: #9ca3af; font-style: italic;">No active test methods</div>';
            return;
        }

        const chipsWrap = document.createElement('div');
        chipsWrap.className = 'flat-chips-wrap';
        chipsWrap.style.display = 'flex';
        chipsWrap.style.flexWrap = 'wrap';
        chipsWrap.style.gap = '8px 12px';

        sortedMethods.forEach(method => {
            const item = document.createElement('div');
            item.textContent = method;
            item.className = 'test-method-chip';
            if (selectedTestMethods.has(method)) {
                item.classList.add('active');
            }

            item.addEventListener('click', () => {
                if (selectedTestMethods.has(method)) {
                    selectedTestMethods.delete(method);
                } else {
                    selectedTestMethods.add(method);
                }

                if (selectedTestMethods.size === 0) {
                    currentTestMethod = 'all';
                    if (testMethodSelect) testMethodSelect.value = 'all';
                } else if (selectedTestMethods.size === 1) {
                    currentTestMethod = Array.from(selectedTestMethods)[0];
                    if (testMethodSelect) testMethodSelect.value = currentTestMethod;
                } else {
                    currentTestMethod = 'multiple';
                    if (testMethodSelect) {
                        let multOpt = testMethodSelect.querySelector('option[value="multiple"]');
                        if (!multOpt) {
                            multOpt = document.createElement('option');
                            multOpt.value = 'multiple';
                            multOpt.textContent = 'Multiple Selected';
                            testMethodSelect.appendChild(multOpt);
                        }
                        testMethodSelect.value = 'multiple';
                    }
                }

                updateChipActiveStates();
                renderGanttChart();
            });

            chipsWrap.appendChild(item);
        });

        testMethodList.appendChild(chipsWrap);
    }

    function updateChipActiveStates() {
        if (!testMethodList) return;
        const chips = testMethodList.querySelectorAll('.test-method-chip');
        chips.forEach(chip => {
            const name = chip.textContent;
            if (selectedTestMethods.has(name)) {
                chip.classList.add('active');
            } else {
                chip.classList.remove('active');
            }
        });
    }

    function populateTestMethods(records) {
        if (!testMethodSelect) return;

        const methods = new Set();
        records.forEach(r => {
            if (r['Test Method']) {
                methods.add(r['Test Method']);
            }
        });

        testMethodSelect.innerHTML =
            '<option value="all">All Test Methods</option>';

        Array.from(methods)
            .sort()
            .forEach(method => {
                const option = document.createElement('option');
                option.value = method;
                option.textContent = method;
                testMethodSelect.appendChild(option);
            });
    }

    function populateCategories(records) {
        const categories = new Set();
        records.forEach(r => {
            if (r.Category) categories.add(r.Category);
        });

        const prevVal = categorySelect ? categorySelect.value : 'all';
        if (categorySelect) {
            categorySelect.innerHTML = '<option value="all">All Categories</option>';
            categories.forEach(cat => {
                const option = document.createElement('option');
                option.value = cat;
                option.textContent = cat;
                categorySelect.appendChild(option);
            });

            if (categories.has(prevVal)) {
                categorySelect.value = prevVal;
                currentCategory = prevVal;
            } else {
                categorySelect.value = 'all';
                currentCategory = 'all';
            }
        }
    }

    function parseDateValue(value) {
        if (!value) return null;
        const date = new Date(value);
        if (!isNaN(date.getTime())) return date;

        const normalized = String(value).trim().replace(/\//g, '-').replace(/\s+/g, ' ');
        const altDate = new Date(normalized);
        return isNaN(altDate.getTime()) ? null : altDate;
    }

    function resetToEmptyState(title = "No Project Selected", text = "Select a project above to render its timeline from actual project data.") {
        emptyState.style.display = 'block';
        canvasWrap.style.display = 'none';
        loadingState.style.display = 'none';

        emptyState.querySelector('h3').textContent = title;
        emptyState.querySelector('p').textContent = text;

        statRows.textContent = '--';
        statCats.textContent = '--';
        statDefects.textContent = '--';

        if (ganttChartInstance) {
            ganttChartInstance.destroy();
            ganttChartInstance = null;
        }
    }

    // Build timeline bars from actual row data
    let renderTimeout = null;
    function renderGanttChart() {
        if (renderTimeout) {
            clearTimeout(renderTimeout);
        }

        // 1. Filter records by Laboratory, Category and selected Test Methods
        let filteredRecords = rawRecords;

        if (currentLaboratory !== 'all') {
            const allowedNums = new Set((labTestNumbersMap[currentLaboratory] || []).map(x => String(x).trim()));
            const allowedMethods = new Set((labTestMethodsMap[currentLaboratory] || []).map(x => String(x).trim()));

            filteredRecords = filteredRecords.filter(r => {
                const numMatch = r['Test Number'] && allowedNums.has(String(r['Test Number']).trim());
                const methodMatch = r['Test Method'] && allowedMethods.has(String(r['Test Method']).trim());
                const catMatch = r['Category'] && String(r['Category']).trim().toLowerCase() === currentLaboratory.toLowerCase();
                const labMatch = r['Laboratory'] && String(r['Laboratory']).trim().toLowerCase() === currentLaboratory.toLowerCase();
                return numMatch || methodMatch || catMatch || labMatch;
            });
        }

        if (currentCategory !== 'all') {
            filteredRecords = rawRecords.filter(r => r.Category === currentCategory);
        }
        if (selectedTestMethods.size > 0) {
            filteredRecords = filteredRecords.filter(r => selectedTestMethods.has(r['Test Method']));
        }

        if (filteredRecords.length === 0) {
            canvasWrap.style.display = 'none';
            if (statsRow) statsRow.style.display = 'none';
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
            if (statsRow) statsRow.style.display = 'flex';
        }

        // 2. Build chart bars from real project records
        // Chained phases: Proto -> DVT -> EVT -> PVT
        const chartData = [];
        let minDate = null;
        let maxDate = null;
        let totalBarsCount = 0;
        let totalDefectsCount = 0;

        filteredRecords.forEach((r, idx) => {
            const label = `[${r.Category}] ${r['Test Method']} (${r['Test Number']}) #${r.id || idx}`;
            let baseStart = customStartDate ? new Date(customStartDate + 'T00:00:00') : parseDateValue(r['Start Date']);
            if (!baseStart) {
                console.warn('Skipping record with invalid start date:', r);
                return;
            }

            totalDefectsCount += parseInt(r['Defect Qty'] || 0);

            // Phase definitions
            const phases = [{name: 'Proto', weeks: parseInt(r['Proto Weeks'] || 0),days: parseInt(r['Proto Days'] || 0),color: '#8E44AD'},
                            {name: 'DVT', weeks: parseInt(r['DVT Weeks'] || 0), days: parseInt(r['DVT Days'] || 0), color: '#2980B9'},
                            {name: 'EVT', weeks: parseInt(r['EVT Weeks'] || 0), days: parseInt(r['EVT Days'] || 0), color: '#27AE60'},
                            { name: 'PVT',weeks: parseInt(r['PVT Weeks'] || 0),days: parseInt(r['PVT Days'] || 0),color: '#D35400'}];

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
        statDefects.textContent = totalDefectsCount;

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

        function wrapLabel(text, maxLength = 32) {
            if (!text || text.length <= maxLength) return text;
            const words = text.split(' ');
            const lines = [];
            let current = '';
            words.forEach(word => {
                if ((current + ' ' + word).trim().length <= maxLength) {
                    current = current ? `${current} ${word}` : word;
                } else {
                    if (current) lines.push(current);
                    current = word;
                }
            });
            if (current) lines.push(current);
            return lines;
        }

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
                        borderColor: (context) => {
                            const raw = context.raw;
                            return raw ? shadeColor(raw.color, -15) : '#777777';
                        },
                        borderWidth: 1,
                        borderRadius: 6,
                        borderSkipped: false,
                        barThickness: 28,
                        maxBarThickness: 28,
                        minBarLength: 12,
                        categoryPercentage: 0.75,
                        barPercentage: 0.8
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {
                        padding: { top: 10, right: 16, left: 8, bottom: 10 }
                    },
                    interaction: {
                        mode: 'nearest',
                        axis: 'x',
                        intersect: true
                    },
                    plugins: {
                        legend: { display: false },
                        title: {
                            display: true,
                            text: `Project ${currentProjectName} - Gantt Timeline`,
                            color: '#ffffff',
                            font: { size: 16, weight: '700' }
                        },
                        subtitle: {
                            display: true,
                            text: 'Proto | DVT | EVT | PVT',
                            color: '#d1d5db',
                            font: { size: 12, weight: '600' }
                        },
                        tooltip: {
                            backgroundColor: 'rgba(30, 30, 30, 0.95)',
                            titleColor: '#ffffff',
                            bodyColor: '#f1f1f1',
                            borderColor: 'rgba(255,255,255,0.08)',
                            borderWidth: 1,
                            padding: 12,
                            callbacks: {
                                title: (context) => {
                                    return context[0].raw.y.split(' #')[0];
                                },
                                label: (context) => {
                                    const raw = context.raw;
                                    const startStr = raw.x[0].toISOString().split('T')[0];
                                    const endStr = raw.x[1].toISOString().split('T')[0];
                                    const durationDays = Math.max(1, Math.round((new Date(raw.x[1]) - new Date(raw.x[0])) / (1000 * 60 * 60 * 24)));

                                    return [
                                        `Start: ${startStr}`,
                                        `End: ${endStr}`,
                                        `Duration: ${durationDays} days`,
                                    ];
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
                            min: spanStart ? new Date(spanStart + 'T00:00:00') : new Date(minDate.getTime() - 7 * 24 * 60 * 60 * 1000),
                            max: spanEnd ? new Date(spanEnd + 'T00:00:00') : new Date(maxDate.getTime() + 7 * 24 * 60 * 60 * 1000),
                            grid: {
                                color: 'rgba(255, 255, 255, 0.08)'
                            },
                            ticks: {
                                color: '#a0a0a0',
                                font: { size: 11 }
                            },
                            title: {
                                display: true,
                                text: 'Project timeline',
                                color: '#9aa0ac',
                                font: { size: 12, weight: '600' }
                            }
                        },
                        y: {
                            grid: { color: 'rgba(255, 255, 255, 0.04)' },
                            ticks: {
                                color: '#e0e0e0',
                                font: { size: 11, weight: '600' },
                                callback: function(value) {
                                    const label = this.getLabelForValue(value);
                                    const cleanLabel = label.split(' #')[0];
                                    return wrapLabel(cleanLabel, 30);
                                }
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

                            ctx.fillStyle = '#ff3b30';
                            ctx.font = '11px sans-serif';
                            ctx.fillText('Today', todayX + 6, y.top + 16);
                            ctx.restore();
                        }
                    }
                }]
            });
        }, 50);
    }

    function shadeColor(color, percent) {
        let R = parseInt(color.substring(1,3),16);
        let G = parseInt(color.substring(3,5),16);
        let B = parseInt(color.substring(5,7),16);
        R = Math.min(255, Math.max(0, R + Math.round(2.55 * percent)));
        G = Math.min(255, Math.max(0, G + Math.round(2.55 * percent)));
        B = Math.min(255, Math.max(0, B + Math.round(2.55 * percent)));
        return `#${((1<<24) + (R<<16) + (G<<8) + B).toString(16).slice(1)}`;
    }
});
