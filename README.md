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
│   │   └── style.css       # Core design system and layout styles
│   └── js/
│       └── script.js       # Chart rendering, AJAX requests, and UI logic
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

### 1. Navigate to Project
```bash
cd test-ops-dashboard
```

### 2. Set Up Virtual Environment
```bash
# Create virtual environment
python -m venv venv

# Activate on Windows
venv\Scripts\activate

# Activate on macOS/Linux
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Environment Configuration
Copy the template `.env.example` file to `.env`:
```bash
cp .env.example .env
```
Open `.env` and configure your Flask secret key and optional SMTP credentials if you want to test real email sending.

### 5. Run the Application
Start the Flask development server:
```bash
python app.py
```
Open your browser and navigate to `http://localhost:5002/`.

### 6. Run Unit Tests
To verify backend functionality:
```bash
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
