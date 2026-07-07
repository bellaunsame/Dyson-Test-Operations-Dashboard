// ==========================================
// GANTT - Chart rendering (Chart.js)
// ==========================================

// Called by dashboard.js via filterAndRenderGantt() or window.OPS references
function renderGanttChart(rows, scale = 'day') {
    const canvas = document.getElementById('ganttChart');
    if (!canvas) return;

    const ctx = canvas.getContext('2d');

    // Destroy previous instance
    if (window.OPS && window.OPS.ganttInstance) {
        window.OPS.ganttInstance.destroy();
        window.OPS.ganttInstance = null;
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

    // Build one bar per phase per row
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

                const shortLabel = `${r['Test Number']} — ${r['Test Method']} [${p.name}]`;
                labels.push(shortLabel);
                chartData.push({ x: [start, end], y: shortLabel });
                bgColors.push(phaseColors[p.name]);
                tooltipInfo.push({
                    phase: p.name, category: r['Category'],
                    method: r['Test Method'], number: r['Test Number'],
                    start, end, totalDays,
                    weeks: parseInt(p.weeks || 0),
                    days:  parseInt(p.days  || 0),
                    qty:   parseInt(p.qty   || 0),
                    comments: r['Comments'] || ''
                });
            }
        });
    });

    // Empty state
    if (chartData.length === 0) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.font = '14px Inter, sans-serif';
        ctx.fillStyle = '#8e8e93';
        ctx.textAlign = 'center';
        ctx.fillText('No active test milestones to display.', canvas.width / 2, canvas.height / 2);
        return;
    }

    // Date range
    const allDates    = chartData.flatMap(d => d.x);
    const minDate     = new Date(Math.min(...allDates));
    const maxDate     = new Date(Math.max(...allDates));
    const paddingDays = scale === 'month' ? 14 : scale === 'week' ? 7 : 4;
    minDate.setDate(minDate.getDate() - paddingDays);
    maxDate.setDate(maxDate.getDate() + paddingDays);

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Update summary stats
    const statToday = document.getElementById('gantt-stat-today');
    const fmtDate   = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    if (statToday) {
        const projectEnd = new Date(Math.max(...chartData.map(d => d.x[1].getTime())));
        const daysLeft   = Math.ceil((projectEnd - today) / 86400000);
        statToday.textContent = daysLeft > 0
            ? `⏱ ${daysLeft} day${daysLeft !== 1 ? 's' : ''} until project completion`
            : daysLeft === 0
                ? '✅ Final phase ends today'
                : `⚠️ Project ended ${Math.abs(daysLeft)} day${Math.abs(daysLeft) !== 1 ? 's' : ''} ago`;
    }

    // Dynamic chart height
    const chartHeight = Math.max(300, chartData.length * 32 + 60);
    const wrapper     = document.getElementById('gantt-chart-wrapper');
    const scrollBox   = document.getElementById('gantt-chart-scroll');
    if (wrapper)   wrapper.style.height = `${chartHeight}px`;
    if (scrollBox) scrollBox.style.maxHeight = '520px';
    canvas.style.height = `${chartHeight}px`;

    // TODAY line plugin
    const todayLinePlugin = {
        id: 'todayLine',
        afterDatasetsDraw(chart) {
            const { ctx: c, chartArea, scales } = chart;
            const todayX = scales.x.getPixelForValue(today);
            if (todayX < chartArea.left || todayX > chartArea.right) return;

            c.save();
            c.beginPath();
            c.strokeStyle = '#ff3b30';
            c.lineWidth   = 2;
            c.setLineDash([5, 4]);
            c.moveTo(todayX, chartArea.top);
            c.lineTo(todayX, chartArea.bottom);
            c.stroke();
            c.font = 'bold 10px Inter, sans-serif';
            c.fillStyle = '#ff3b30';
            c.textAlign = 'center';
            c.fillText('TODAY', todayX, chartArea.top - 4);
            c.restore();
        }
    };

    // Bar label plugin
    const barLabelPlugin = {
        id: 'barLabels',
        afterDatasetsDraw(chart) {
            const { ctx: c, chartArea } = chart;
            const meta = chart.getDatasetMeta(0);
            c.save();
            c.font = 'bold 10px Inter, sans-serif';
            c.fillStyle = 'rgba(255,255,255,0.9)';
            c.textBaseline = 'middle';

            meta.data.forEach((bar, i) => {
                const info = tooltipInfo[i];
                const barW = Math.abs(bar.width);
                if (barW < 30) return;
                const barX = Math.min(bar.x, bar.base);
                c.save();
                c.beginPath();
                c.rect(Math.max(barX, chartArea.left), chartArea.top, Math.min(barW, chartArea.right - barX), chartArea.bottom - chartArea.top);
                c.clip();
                const text = barW > 100
                    ? `${info.phase} · ${info.totalDays}d${info.qty > 0 ? ` · Qty ${info.qty}` : ''}`
                    : info.phase;
                c.fillText(text, barX + 8, bar.y);
                c.restore();
            });
            c.restore();
        }
    };

    // Build chart
    window.OPS.ganttInstance = new Chart(ctx, {
        type: 'bar',
        plugins: [todayLinePlugin, barLabelPlugin],
        data: {
            labels,
            datasets: [{
                data:            chartData,
                backgroundColor: bgColors.map(c => c + 'dd'),
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
            indexAxis:           'y',
            responsive:          true,
            maintainAspectRatio: false,
            animation: { duration: 600, easing: 'easeOutQuart' },
            layout: { padding: { top: 18, right: 12, bottom: 4, left: 0 } },
            plugins: {
                legend: { display: false },
                tooltip: {
                    callbacks: {
                        title(ctx) { return tooltipInfo[ctx[0].dataIndex].category; },
                        label(ctx) {
                            const i = tooltipInfo[ctx.dataIndex];
                            const s = i.start.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
                            const e = i.end.toLocaleDateString('en-US',   { month: 'short', day: 'numeric', year: 'numeric' });
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
                        displayFormats: { day: 'MMM d', week: 'MMM d', month: 'MMM yy' },
                        tooltipFormat: 'MMM d, yyyy'
                    },
                    min: minDate,
                    max: maxDate,
                    position: 'top',
                    grid: { color: 'rgba(255,255,255,0.05)', drawBorder: false, lineWidth: 1 },
                    ticks: {
                        color: '#8e8e93',
                        font:  { family: 'Inter, sans-serif', size: 11 },
                        maxRotation: 0,
                        autoSkip: true,
                        maxTicksLimit: scale === 'day' ? 14 : scale === 'week' ? 12 : 8
                    }
                },
                y: {
                    grid: { color: 'rgba(255,255,255,0.03)', drawBorder: false },
                    ticks: {
                        color:   '#c0c0c6',
                        font:    { family: 'Inter, sans-serif', size: 11 },
                        padding: 6,
                        callback(val) {
                            const lbl = this.getLabelForValue(val);
                            return lbl && lbl.length > 42 ? lbl.slice(0, 40) + '…' : lbl;
                        }
                    },
                    afterFit(s) { s.width = 230; }
                }
            }
        }
    });
}
