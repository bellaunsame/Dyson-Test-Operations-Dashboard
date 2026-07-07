# Test Operations Dashboard

A premium, full-stack web application designed for internal engineering operations. The dashboard allows project leads to manage product development milestones, visualize sprint timelines with automated Gantt charts, generate daily status reports, and query the project database using an integrated AI assistant.

---

## 🧠 Key Features

1. **AI Operations Assistant**
   - An integrated floating AI chatbot (`AI Assistant`) that queries the active task database in real-time.
   - Answers questions regarding task counts, average progress, overdue deadlines, and owner teams.

2. **Interactive Gantt Chart**
   - Renders horizontal timeline bars using **Chart.js**.
   - Automatically calculates task **End Dates** and **Status** (Not Started, In Progress, Completed) based on the inputted *Start Date*, *Duration*, and *Progress*.

3. **SharePoint & Excel Integration**
   - Persists all project data locally in a structured Excel spreadsheet (`uploads/tasks.xlsx`) using **pandas** and **openpyxl**.
   - Includes a **SharePoint Sync** feature to simulate cloud-to-local database synchronization.

4. **Automated Daily PDF Report**
   - Generates a styled, branded PDF status report containing a system summary, completion statistics, and a detailed task inventory table.

5. **One-Click Email Broadcast**
   - Dispatches the generated PDF report directly to operations managers via SMTP using **Flask-Mail**.
   - Automatically falls back to a simulated console-log mode if SMTP credentials are not configured.

6. **Secure Access Controls**
   - Session-based login system ensuring only authorized engineers can access internal operations room data.

---

## ⚙️ Tech Stack

- **Backend:** Python (Flask), Gunicorn, pandas, openpyxl, ReportLab, Flask-Mail, python-dotenv
- **Frontend:** HTML5, CSS3 (Vanilla), JavaScript (ES6+)
- **Visualization:** Chart.js (with Date Adapter)

---

## 📁 Project Structure

```text
test-ops-dashboard/
│
├── app.py                  # Flask application server & API routes
├── test_app.py             # Backend unit tests
├── requirements.txt        # Python package dependencies
├── .env.example            # Template for environment variables
│
├── static/
│   ├── css/
│   │   ├── base.css        # Reset, variables, global elements
│   │   ├── layout.css      # Sidebar, main layout, header, stats grid
│   │   ├── components.css  # Buttons, tables, chatbot, modals, toasts
│   │   └── login.css       # Stale login styles (retained)
│   └── js/
│       ├── dashboard.js    # State management, data loading, filters
│       ├── gantt.js        # Timeline Chart.js visualization
│       ├── pdf-export.js   # PDF and Excel report download triggers
│       └── email-chat.js   # Email dispatch forms & AI chatbot interactions
│
├── templates/
│   ├── base.html           # Core HTML skeleton (Google Fonts, Chart.js)
│   ├── home.html           # Landing page
│   ├── login.html          # Authentication page
│   └── dashboard.html      # Main operational interface & AI panel
│
└── uploads/
    └── tasks.xlsx          # Excel database (auto-initialized)
```

---

## 🚀 Local Setup & Installation

### 1. Clone & Navigate to Project
```bash
git clone <repository-url>
cd test-ops-dashboard
```

### 2. Set Up Virtual Environment
Create and activate a virtual environment to isolate project dependencies:

**On Windows:**
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate
```

**On macOS/Linux:**
```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate
```

### 3. Install Dependencies
Ensure you install all required packages listed in the pinned requirements file:
```bash
pip install -r requirements.txt
```

### 4. Environment Configuration
Copy the template `.env.example` file to `.env`:
```bash
cp .env.example .env
```
Open `.env` and configure your Flask secret key and optional SharePoint or SMTP credentials if you want to test real cloud synchronization and email reporting.

### 5. Run the Application
Start the Flask development server:

**On Windows:**
```bash
python app.py
```

**On macOS/Linux:**
```bash
python3 app.py
```

Open your browser and navigate to `http://localhost:5002/`.

### 6. Run Unit Tests
To verify backend functionality, you can run tests using `pytest` or Python's built-in `unittest` module:
```bash
# Run tests with pytest (recommended)
pytest

# Run tests with unittest
python -m unittest test_app.py
```

---

## 🔑 Demo Credentials

To access the dashboard, use the following internal credentials:
- **Username:** `admin`
- **Password:** `admin2026`

---

## 🔒 Confidentiality
*This software is proprietary and confidential. For internal Technology Operations use only.*
