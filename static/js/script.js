// ==========================================
// DYSON DASHBOARD - FRONTEND LOGIC
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
    // Global State
    let tasksData = [];
    let editingTaskId = null;
    let ganttChartInstance = null;

    // DOM Elements
    const totalTasksEl = document.getElementById('stat-total-tasks');
    const inProgressEl = document.getElementById('stat-in-progress');
    const completedEl = document.getElementById('stat-completed');
    const avgProgressEl = document.getElementById('stat-avg-progress');
    const tasksTableBody = document.querySelector('#tasks-table tbody');
    const addTaskForm = document.getElementById('add-task-form');
    const chatbotPanel = document.getElementById('chatbot-panel');
    const btnMinimizeChat = document.getElementById('btn-minimize-chat');
    const btnChatbotBubble = document.getElementById('btn-chatbot-bubble');
    const chatMessages = document.getElementById('chat-messages');
    const chatForm = document.getElementById('chat-form');
    const chatInput = document.getElementById('chat-input');
    const btnSyncSharePoint = document.getElementById('btn-sync-sharepoint');
    
    // Modal Elements
    const emailModal = document.getElementById('email-modal');
    const btnEmailModalTrigger = document.getElementById('btn-email-modal-trigger');
    const btnCloseModal = document.getElementById('btn-close-modal');
    const btnCancelEmail = document.getElementById('btn-cancel-email');
    const emailForm = document.getElementById('email-form');
    
    // Edit Form Controls
    const btnCancelEdit = document.getElementById('btn-cancel-edit');
    const btnAddTask = document.getElementById('btn-add-task');
    
    // Log Modal Elements
    const emailLogModal = document.getElementById('email-log-modal');
    const btnCloseLogModal = document.getElementById('btn-close-log-modal');
    const btnCloseLogOk = document.getElementById('btn-close-log-ok');
    const emailLogContent = document.getElementById('email-log-content');

    // Set Default Start Date in Form to Today
    const todayStr = new Date().toISOString().split('T')[0];
    document.getElementById('start-date').value = todayStr;

    // --- INITIAL DATA FETCH ---
    fetchTasks();

    // ==========================================
    // DATA FETCHING & RENDERING
    // ==========================================

    function fetchTasks() {
        fetch('/api/tasks')
            .then(res => res.json())
            .then(data => {
                tasksData = data;
                updateStats(data);
                renderTable(data);
                renderGanttChart(data);
            })
            .catch(err => {
                console.error('Error fetching tasks:', err);
                showToast('Error loading tasks from database', 'error');
            });
    }

    // Update Top Row Statistics
    function updateStats(tasks) {
        const total = tasks.length;
        const completed = tasks.filter(t => t.Progress === 100).length;
        const inProgress = tasks.filter(t => t.Progress > 0 && t.Progress < 100).length;
        const totalProgress = tasks.reduce((sum, t) => sum + t.Progress, 0);
        const avg = total > 0 ? Math.round(totalProgress / total) : 0;

        totalTasksEl.textContent = total;
        inProgressEl.textContent = inProgress;
        completedEl.textContent = completed;
        avgProgressEl.textContent = `${avg}%`;
    }

    // Render Tasks Table
    function renderTable(tasks) {
        tasksTableBody.innerHTML = '';

        if (tasks.length === 0) {
            tasksTableBody.innerHTML = `
                <tr>
                    <td colspan="9" class="text-center">No tasks found. Add a task above to begin.</td>
                </tr>
            `;
            return;
        }

        tasks.forEach(task => {
            const tr = document.createElement('tr');
            
            let statusBadge = '';
            if (task.Progress === 100) {
                statusBadge = '<span class="badge badge-completed">Completed</span>';
            } else if (task.Progress > 0) {
                statusBadge = '<span class="badge badge-in-progress">In Progress</span>';
            } else {
                statusBadge = '<span class="badge badge-not-started">Not Started</span>';
            }

            tr.innerHTML = `
                <td><strong>${task['Task ID']}</strong></td>
                <td>${task['Task Name']}</td>
                <td>${task['Start Date']}</td>
                <td>${task['End Date']}</td>
                <td>${task['Duration']} days</td>
                <td>
                    <div class="task-progress-bar">
                        <div class="task-progress-fill ${task.Progress === 100 ? 'completed' : ''}" style="width: ${task.Progress}%"></div>
                    </div>
                    <span class="progress-text">${task.Progress}%</span>
                </td>
                <td>${task['Owner']}</td>
                <td>${statusBadge}</td>
                <td>
                    <button class="edit-btn" data-id="${task['Task ID']}" title="Edit Task">
                        <i class="fa-solid fa-pencil"></i>
                    </button>
                    <button class="delete-btn" data-id="${task['Task ID']}" title="Delete Task">
                        <i class="fa-solid fa-trash-can"></i>
                    </button>
                </td>
            `;

            // Attach Edit Event Listener
            tr.querySelector('.edit-btn').addEventListener('click', (e) => {
                const taskId = e.currentTarget.getAttribute('data-id');
                startEditTask(taskId);
            });

            // Attach Delete Event Listener
            tr.querySelector('.delete-btn').addEventListener('click', (e) => {
                const taskId = e.currentTarget.getAttribute('data-id');
                deleteTask(taskId);
            });

            tasksTableBody.appendChild(tr);
        });
    }

    // Render Chart.js Gantt Chart
    function renderGanttChart(tasks) {
        const ctx = document.getElementById('ganttChart').getContext('2d');
        
        if (ganttChartInstance) {
            ganttChartInstance.destroy();
        }

        if (tasks.length === 0) {
            ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
            return;
        }

        // Sort tasks by start date so they display in chronological order
        const sortedTasks = [...tasks].sort((a, b) => new Date(a['Start Date']) - new Date(b['Start Date']));

        const labels = sortedTasks.map(t => t['Task Name']);
        
        // Map tasks to floating bar data: [start_date, end_date]
        const chartData = sortedTasks.map(t => [t['Start Date'], t['End Date']]);

        // Dynamically color bars: Completed = green, In Progress = blue/fuchsia, Not Started = gray
        const backgroundColors = sortedTasks.map(t => {
            if (t.Progress === 100) return '#34c759'; // Success Green
            if (t.Progress > 0) return '#00539C';      // Prussian Blue
            return '#555555';                          // Muted Gray
        });

        const borderColors = sortedTasks.map(t => {
            if (t.Progress === 100) return '#34c759';
            if (t.Progress > 0) return '#00539C';
            return '#2c2c2e';
        });

        // Determine axis min/max padding
        const startDates = sortedTasks.map(t => new Date(t['Start Date']));
        const endDates = sortedTasks.map(t => new Date(t['End Date']));
        const absoluteMin = new Date(Math.min(...startDates));
        const absoluteMax = new Date(Math.max(...endDates));
        
        // Pad min/max by 3 days for aesthetics
        absoluteMin.setDate(absoluteMin.getDate() - 3);
        absoluteMax.setDate(absoluteMax.getDate() + 3);

        const config = {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Timeline',
                    data: chartData,
                    backgroundColor: backgroundColors,
                    borderColor: borderColors,
                    borderWidth: 1,
                    borderRadius: 4,
                    borderSkipped: false,
                    barPercentage: 0.6
                }]
            },
            options: {
                indexAxis: 'y', // Makes it a horizontal bar chart
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false // No legend needed for Gantt
                    },
                    tooltip: {
                        callbacks: {
                            title: (context) => {
                                return sortedTasks[context[0].dataIndex]['Task Name'];
                            },
                            label: (context) => {
                                const task = sortedTasks[context.dataIndex];
                                return [
                                    `Duration: ${task['Duration']} days`,
                                    `Timeline: ${task['Start Date']} to ${task['End Date']}`,
                                    `Progress: ${task['Progress']}%`,
                                    `Owner: ${task['Owner']}`
                                ];
                            }
                        },
                        backgroundColor: '#1c1c1e',
                        titleFont: { family: 'Inter', weight: 'bold' },
                        bodyFont: { family: 'Inter' },
                        borderColor: '#2c2c2e',
                        borderWidth: 1
                    }
                },
                scales: {
                    x: {
                        type: 'time',
                        time: {
                            unit: 'day',
                            displayFormats: {
                                day: 'MMM dd'
                            }
                        },
                        min: absoluteMin.toISOString().split('T')[0],
                        max: absoluteMax.toISOString().split('T')[0],
                        grid: {
                            color: '#2c2c2e',
                            drawBorder: false
                        },
                        ticks: {
                            color: '#8e8e93',
                            font: { family: 'Inter', size: 11 }
                        }
                    },
                    y: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            color: '#f5f5f7',
                            font: { family: 'Inter', size: 12, weight: 500 }
                        }
                    }
                }
            }
        };

        ganttChartInstance = new Chart(ctx, config);
    }

    // ==========================================
    // TASK CRUD OPERATIONS
    // ==========================================

    // Add / Edit Task Form Submission
    addTaskForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const taskName = document.getElementById('task-name').value.trim();
        const startDate = document.getElementById('start-date').value;
        const duration = parseInt(document.getElementById('duration').value);
        const progress = parseInt(document.getElementById('progress').value);
        const owner = document.getElementById('owner').value;

        const payload = {
            "Task Name": taskName,
            "Start Date": startDate,
            "Duration": duration,
            "Progress": progress,
            "Owner": owner
        };

        const btn = document.getElementById('btn-add-task');
        const originalText = btn.innerHTML;
        btn.disabled = true;
        btn.innerHTML = editingTaskId 
            ? '<i class="fa-solid fa-spinner fa-spin"></i> Updating...' 
            : '<i class="fa-solid fa-spinner fa-spin"></i> Adding...';

        const url = editingTaskId ? `/api/tasks/${editingTaskId}` : '/api/tasks';
        const method = editingTaskId ? 'PUT' : 'POST';

        fetch(url, {
            method: method,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        })
        .then(res => {
            if (!res.ok) {
                return res.json().then(data => { throw new Error(data.error || 'Server error'); });
            }
            return res.json();
        })
        .then(data => {
            if (editingTaskId) {
                showToast(`Task "${data['Task Name']}" updated successfully!`, 'success');
            } else {
                showToast(`Task "${data['Task Name']}" created successfully!`, 'success');
            }
            cancelEditTask();
            fetchTasks();
        })
        .catch(err => {
            showToast(err.message, 'error');
        })
        .finally(() => {
            btn.disabled = false;
            if (btn.innerHTML.includes('fa-spinner')) {
                if (editingTaskId) {
                    btn.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Update Task';
                } else {
                    btn.innerHTML = '<i class="fa-solid fa-plus"></i> Add Task';
                }
            }
        });
    });

    // Edit Task Mode Toggle
    function startEditTask(taskId) {
        const task = tasksData.find(t => t['Task ID'] === taskId);
        if (!task) return;

        editingTaskId = taskId;
        
        // Populate form fields
        document.getElementById('task-name').value = task['Task Name'];
        document.getElementById('start-date').value = task['Start Date'];
        document.getElementById('duration').value = task['Duration'];
        document.getElementById('progress').value = task['Progress'];
        document.getElementById('owner').value = task['Owner'];

        // Update UI buttons
        btnAddTask.innerHTML = '<i class="fa-solid fa-floppy-disk"></i> Update Task';
        btnAddTask.className = 'btn btn-accent'; // Dyson Fuchsia accent for editing
        btnCancelEdit.style.display = 'inline-block';
        
        // Scroll to form for better UX
        document.getElementById('add-task-form').scrollIntoView({ behavior: 'smooth' });
    }

    function cancelEditTask() {
        editingTaskId = null;
        addTaskForm.reset();
        document.getElementById('start-date').value = todayStr;
        
        btnAddTask.innerHTML = '<i class="fa-solid fa-plus"></i> Add Task';
        btnAddTask.className = 'btn btn-primary';
        btnCancelEdit.style.display = 'none';
    }

    btnCancelEdit.addEventListener('click', cancelEditTask);

    // Delete Task Operation
    function deleteTask(taskId) {
        if (!confirm(`Are you sure you want to delete task ${taskId}?`)) return;

        showToast(`Deleting task ${taskId}...`, 'info');

        fetch(`/api/tasks/${taskId}`, {
            method: 'DELETE'
        })
        .then(res => {
            if (!res.ok) {
                return res.json().then(data => { throw new Error(data.error || 'Server error'); });
            }
            return res.json();
        })
        .then(data => {
            showToast(data.message, 'success');
            fetchTasks();
        })
        .catch(err => {
            showToast(err.message, 'error');
        });
    }

    // ==========================================
    // SHAREPOINT CLOUD SYNC SIMULATION
    // ==========================================

    btnSyncSharePoint.addEventListener('click', (e) => {
        e.preventDefault();
        showToast('Connecting to SharePoint cloud...', 'info');
        
        btnSyncSharePoint.classList.add('active');
        const originalContent = btnSyncSharePoint.innerHTML;
        btnSyncSharePoint.innerHTML = '<i class="fa-solid fa-arrows-rotate fa-spin"></i> Syncing Cloud...';
        btnSyncSharePoint.style.pointerEvents = 'none';

        fetch('/api/sync-sharepoint', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' }
        })
        .then(res => res.json())
        .then(data => {
            btnSyncSharePoint.classList.remove('active');
            btnSyncSharePoint.innerHTML = originalContent;
            btnSyncSharePoint.style.pointerEvents = 'auto';
            if (data.success) {
                showToast(data.message, 'success');
                fetchTasks(); // Reload from Excel to show it's synced
            } else {
                showToast(data.error || 'Failed to sync with SharePoint.', 'error');
            }
        })
        .catch(err => {
            btnSyncSharePoint.classList.remove('active');
            btnSyncSharePoint.innerHTML = originalContent;
            btnSyncSharePoint.style.pointerEvents = 'auto';
            showToast('SharePoint sync failed. Connection error.', 'error');
            console.error('SharePoint Sync Error:', err);
        });
    });

    // ==========================================
    // EMAIL MODAL & AUTOMATION
    // ==========================================

    // Open Modal
    btnEmailModalTrigger.addEventListener('click', () => {
        emailModal.classList.add('active');
        // Pre-fill subject with current date
        const formattedDate = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
        document.getElementById('email-subject').value = `Dyson Operations Daily Report - ${formattedDate}`;
    });

    // Close Modal Helpers
    function closeEmailModal() {
        emailModal.classList.remove('active');
    }
    
    btnCloseModal.addEventListener('click', closeEmailModal);
    btnCancelEmail.addEventListener('click', closeEmailModal);

    // Send Email Form Submission
    emailForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const recipient = document.getElementById('email-recipient').value;
        const subject = document.getElementById('email-subject').value;

        const submitBtn = document.getElementById('btn-send-email-submit');
        const originalText = submitBtn.innerHTML;
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Sending...';

        const formData = new FormData();
        formData.append('recipient', recipient);
        formData.append('subject', subject);

        fetch('/send-email', {
            method: 'POST',
            body: formData
        })
        .then(res => res.json())
        .then(data => {
            closeEmailModal();
            if (data.success) {
                showToast(data.message, 'success');
                
                // If it was simulated, show the simulated SMTP log modal
                if (data.simulated) {
                    setTimeout(() => {
                        emailLogContent.textContent = JSON.stringify(data.log.details, null, 2);
                        emailLogModal.classList.add('active');
                    }, 500);
                }
            } else {
                showToast(data.error || 'Failed to send email.', 'error');
            }
        })
        .catch(err => {
            console.error('Error sending email:', err);
            showToast('SMTP Connection failed. Check server logs.', 'error');
        })
        .finally(() => {
            submitBtn.disabled = false;
            submitBtn.innerHTML = originalText;
        });
    });

    // Close Log Modal
    btnCloseLogModal.addEventListener('click', () => emailLogModal.classList.remove('active'));
    btnCloseLogOk.addEventListener('click', () => emailLogModal.classList.remove('active'));

    // Trigger PDF Download toast
    document.getElementById('btn-pdf-download').addEventListener('click', () => {
        showToast('Generating PDF report. Downloading...', 'info');
    });

    // ==========================================
    // AI CHATBOT SYSTEM
    // ==========================================

    // Toggle Chatbot Open/Close from bubble and close button
    btnMinimizeChat.addEventListener('click', () => {
        chatbotPanel.classList.remove('active');
        btnChatbotBubble.classList.remove('active');
    });

    btnChatbotBubble.addEventListener('click', () => {
        const isActive = chatbotPanel.classList.toggle('active');
        btnChatbotBubble.classList.toggle('active');
        
        // When opening the chat, scroll to bottom of messages and focus input
        if (isActive) {
            setTimeout(() => {
                chatMessages.scrollTop = chatMessages.scrollHeight;
                chatInput.focus();
            }, 100);
        }
    });

    // Handle suggestion chips
    document.addEventListener('click', (e) => {
        if (e.target.classList.contains('chat-chip')) {
            const query = e.target.textContent;
            chatInput.value = query;
            chatForm.dispatchEvent(new Event('submit'));
        }
    });

    // Chat Form Submission
    chatForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const message = chatInput.value.trim();
        if (!message) return;

        // Append user message
        appendMessage(message, 'user');
        chatInput.value = '';

        // Append typing indicator
        const typingId = appendTypingIndicator();
        chatMessages.scrollTop = chatMessages.scrollHeight;

        // Send to backend
        fetch('/chat', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ message: message })
        })
        .then(res => res.json())
        .then(data => {
            removeTypingIndicator(typingId);
            appendMessage(data.response, 'bot');
        })
        .catch(err => {
            console.error('Chat error:', err);
            removeTypingIndicator(typingId);
            appendMessage("Sorry, I'm having trouble connecting to my cognitive system. Please try again.", 'bot');
        })
        .finally(() => {
            chatMessages.scrollTop = chatMessages.scrollHeight;
        });
    });

    // Append Message helper
    function appendMessage(text, sender) {
        const msgDiv = document.createElement('div');
        msgDiv.classList.add('message', `${sender}-message`);
        
        // Basic Markdown parsing for bold text and bullet points
        let formattedText = text
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\*(.*?)\*/g, '<em>$1</em>')
            .replace(/`(.*?)`/g, '<code>$1</code>');
            
        // Check if message has bullet points
        if (formattedText.includes('\n- ')) {
            const lines = formattedText.split('\n');
            let parsedHtml = '';
            let inList = false;
            
            lines.forEach(line => {
                if (line.trim().startsWith('- ')) {
                    if (!inList) {
                        parsedHtml += '<ul>';
                        inList = true;
                    }
                    parsedHtml += `<li>${line.trim().substring(2)}</li>`;
                } else {
                    if (inList) {
                        parsedHtml += '</ul>';
                        inList = false;
                    }
                    if (line.trim() !== '') {
                        parsedHtml += `<p>${line}</p>`;
                    }
                }
            });
            
            if (inList) parsedHtml += '</ul>';
            msgDiv.innerHTML = parsedHtml;
        } else {
            // Replace newlines with <br> if no lists
            formattedText = formattedText.replace(/\n/g, '<br>');
            msgDiv.innerHTML = `<p>${formattedText}</p>`;
        }

        chatMessages.appendChild(msgDiv);
    }

    // Typing Indicator helpers
    function appendTypingIndicator() {
        const id = 'typing-' + Date.now();
        const indicator = document.createElement('div');
        indicator.classList.add('message', 'bot-message', 'typing-container');
        indicator.id = id;
        indicator.innerHTML = `
            <div class="typing-indicator">
                <span></span>
                <span></span>
                <span></span>
            </div>
        `;
        chatMessages.appendChild(indicator);
        return id;
    }

    function removeTypingIndicator(id) {
        const indicator = document.getElementById(id);
        if (indicator) {
            indicator.remove();
        }
    }

    // ==========================================
    // TOAST NOTIFICATIONS SYSTEM
    // ==========================================

    function showToast(message, type = 'info') {
        const container = document.getElementById('toast-container');
        const toast = document.createElement('div');
        toast.classList.add('toast', `toast-${type}`);
        
        let icon = 'fa-circle-info';
        if (type === 'success') icon = 'fa-circle-check';
        if (type === 'error') icon = 'fa-triangle-exclamation';
        
        toast.innerHTML = `
            <i class="fa-solid ${icon}"></i>
            <div class="toast-content">${message}</div>
        `;

        container.appendChild(toast);
        
        // Trigger reflow for transition
        toast.offsetHeight;
        
        toast.classList.add('show');

        // Auto remove
        setTimeout(() => {
            toast.classList.remove('show');
            toast.addEventListener('transitionend', () => {
                toast.remove();
            });
        }, 4000);
    }
});
