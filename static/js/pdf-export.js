// ==========================================
// PDF EXPORT - Project and consolidated PDF
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
    // Simple quick-export button (nav link in header area)
    const btnPdfDownload = document.getElementById('btn-pdf-download');
    if (btnPdfDownload) {
        btnPdfDownload.addEventListener('click', () => {
            const proj  = window.OPS && window.OPS.currentProject;
            const cat   = window.OPS && window.OPS.currentCategory;
            const scale = window.OPS && window.OPS.currentScale;
            const params = new URLSearchParams();
            if (proj)  params.set('project', proj);
            if (cat && cat !== 'all') params.set('category', cat);
            if (scale) params.set('scale', scale);
            const label = `Project ${proj || 'All'} (${cat !== 'all' ? cat + ' / ' : ''}${scale})`;
            window.showToast(`Generating PDF report for ${label}...`, 'info');
            window.location.href = `/generate-report?${params.toString()}`;
        });
    }

    // Sidebar PDF export — uses active project context
    const btnPdfSidebar = document.getElementById('btn-pdf-download-sidebar');
    if (btnPdfSidebar) {
        btnPdfSidebar.addEventListener('click', () => {
            const activeProj = getActiveProjectContext();
            if (!activeProj) {
                window.showToast('No project selected — exporting full consolidated PDF report...', 'info');
                const catSel = document.getElementById('gantt-category-select') || document.getElementById('category-select');
                const activeCategory = catSel ? catSel.value : 'all';
                const activeScaleBtn = document.querySelector('.scale-btn.active');
                const activeScale = activeScaleBtn ? activeScaleBtn.getAttribute('data-scale') : 'week';
                triggerConsolidatedPdfDownload(activeCategory, activeScale);
                return;
            }
            const catSel = document.getElementById('gantt-category-select') || document.getElementById('category-select');
            const activeCategory = catSel ? catSel.value : 'all';
            const activeScaleBtn = document.querySelector('.scale-btn.active');
            const activeScale = activeScaleBtn ? activeScaleBtn.getAttribute('data-scale') : 'week';
            triggerPdfDownload(activeProj, activeCategory, activeScale);
        });
    }

    // Excel export
    const btnExcelExport = document.getElementById('btn-excel-export');
    if (btnExcelExport) {
        btnExcelExport.addEventListener('click', () => {
            const icon = btnExcelExport.querySelector('i');
            const origClass = icon ? icon.className : 'fa-solid fa-file-excel';
            if (icon) icon.className = 'fa-solid fa-spinner fa-spin';
            btnExcelExport.disabled = true;
            window.showToast('Preparing Excel report for download...', 'info');

            fetch('/api/export-excel')
                .then(res => {
                    if (!res.ok) return res.json().then(d => { throw new Error(d.error || 'Export failed'); });
                    return res.blob();
                })
                .then(blob => {
                    const today = new Date().toISOString().slice(0, 10);
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `Full_Operations_Report_${today}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
                    window.showToast('Excel report downloaded successfully!', 'success');
                })
                .catch(err => {
                    window.showToast(`Excel export failed: ${err.message}`, 'error');
                    console.error('Excel Export Error:', err);
                })
                .finally(() => {
                    if (icon) icon.className = origClass;
                    btnExcelExport.disabled = false;
                });
        });
    }
});

// Get current project from URL path or active selects
function getActiveProjectContext() {
    const path = window.location.pathname;
    if (path.startsWith('/tables/')) return decodeURIComponent(path.split('/').pop());
    const chartSel = document.getElementById('gantt-project-select');
    if (chartSel && chartSel.value) return chartSel.value;
    const dashSel  = document.getElementById('project-select');
    if (dashSel && dashSel.value) return dashSel.value;
    return '';
}

// Project PDF download with progress modal
function triggerPdfDownload(projectName, categoryFilter, scaleFilter) {
    const modal   = document.getElementById('export-progress-modal');
    const bar     = document.getElementById('export-progress-bar');
    const msg     = document.getElementById('export-progress-message');
    const pct     = document.getElementById('export-progress-percent');
    const title   = document.getElementById('export-progress-title');
    const desc    = document.getElementById('export-progress-desc');

    if (title) title.textContent = `Compiling Report for Project ${projectName}`;
    if (desc)  desc.textContent  = `Generating timeline, dashboard, and tables for Project ${projectName}.`;
    if (bar)   bar.style.width   = '0%';
    if (pct)   pct.textContent   = '0%';
    if (msg)   msg.textContent   = 'Initializing...';
    if (modal) modal.classList.add('active');

    let pollInterval = null;
    const cleanup = () => { if (modal) modal.classList.remove('active'); };

    const triggerDownload = (taskId) => {
        const filename = `Gantt_Report_${projectName}.pdf`;
        window.location.href = `/api/export-project/download/${taskId}/${encodeURIComponent(filename)}`;
        window.showToast(`PDF report for Project ${projectName} downloaded successfully!`, 'success');
        setTimeout(cleanup, 800);
    };

    const poll = (taskId) => {
        pollInterval = setInterval(() => {
            fetch(`/api/export-project/progress/${taskId}`)
                .then(res => { if (!res.ok) throw new Error('Task state unavailable'); return res.json(); })
                .then(state => {
                    const prog = state.progress || 0;
                    if (bar) bar.style.width = `${prog}%`;
                    if (pct) pct.textContent = `${prog}%`;
                    if (msg) msg.textContent = state.message || 'Processing...';
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
                    window.showToast(`Export tracking failed: ${err.message}`, 'error');
                    cleanup();
                });
        }, 600);
    };

    fetch('/api/export-project/start', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ project: projectName, comment: sessionStorage.getItem('pdf_comment_' + projectName) || '', category: categoryFilter, scale: scaleFilter })
    })
    .then(r => { if (!r.ok) throw new Error('Failed to start PDF export task'); return r.json(); })
    .then(data => poll(data.task_id))
    .catch(err => {
        window.showToast(`PDF export failed: ${err.message}`, 'error');
        console.error('PDF Export Error:', err);
        cleanup();
    });
}

// Consolidated PDF (all projects)
function triggerConsolidatedPdfDownload(categoryFilter, scaleFilter) {
    const modal = document.getElementById('export-progress-modal');
    const bar     = document.getElementById('export-progress-bar');
    const msg     = document.getElementById('export-progress-message');
    const pct     = document.getElementById('export-progress-percent');
    const title   = document.getElementById('export-progress-title');
    const desc    = document.getElementById('export-progress-desc');

    if (title) title.textContent = 'Compiling Full PDF Report';
    if (desc)  desc.textContent  = 'Generating timelines and defect summaries for all projects.';
    if (bar)   bar.style.width   = '0%';
    if (pct)   pct.textContent   = '0%';
    if (msg)   msg.textContent   = 'Initializing...';
    if (modal) modal.classList.add('active');

    let pollInterval = null;
    const cleanup = () => { if (modal) modal.classList.remove('active'); };

    const poll = (taskId) => {
        pollInterval = setInterval(() => {
            fetch(`/api/export-consolidated/progress/${taskId}`)
                .then(res => { if (!res.ok) throw new Error('Task state unavailable'); return res.json(); })
                .then(state => {
                    if (bar) bar.style.width = `${state.progress || 0}%`;
                    if (pct) pct.textContent = `${state.progress || 0}%`;
                    if (msg) msg.textContent = state.message || 'Processing...';
                    if (state.status === 'completed') {
                        clearInterval(pollInterval);
                        window.location.href = `/api/export-consolidated/download/${taskId}/Consolidated_Project_Report.pdf`;
                        window.showToast('Consolidated PDF downloaded!', 'success');
                        setTimeout(cleanup, 800);
                    } else if (state.status === 'failed') {
                        clearInterval(pollInterval);
                        window.showToast(state.message || 'PDF generation failed', 'error');
                        cleanup();
                    }
                })
                .catch(err => {
                    clearInterval(pollInterval);
                    window.showToast(`Export tracking failed: ${err.message}`, 'error');
                    cleanup();
                });
        }, 600);
    };

    fetch('/api/export-consolidated/start', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ category: categoryFilter, scale: scaleFilter })
    })
        .then(r => { if (!r.ok) throw new Error('Failed to start export'); return r.json(); })
        .then(data => poll(data.task_id))
        .catch(err => {
            window.showToast(`PDF export failed: ${err.message}`, 'error');
            cleanup();
        });
}

// Expose to window for chatbot module
window.triggerPdfDownload = triggerPdfDownload;
