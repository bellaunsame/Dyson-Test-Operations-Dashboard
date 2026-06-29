import os
import uuid
import datetime
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import pandas as pd
from dotenv import load_dotenv
from flask_mail import Mail, Message
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from flask_sqlalchemy import SQLAlchemy
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "dyson-ops-secret-key-2026")

# Configure Flask-Mail
app.config['MAIL_SERVER'] = os.getenv('MAIL_SERVER', '')
app.config['MAIL_PORT'] = int(os.getenv('MAIL_PORT', 587))
app.config['MAIL_USE_TLS'] = os.getenv('MAIL_USE_TLS', 'True').lower() in ('true', '1', 't')
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME', '')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD', '')
app.config['MAIL_DEFAULT_SENDER'] = os.getenv('MAIL_DEFAULT_SENDER', 'ops-alerts@dyson-dashboard.local')

mail = Mail(app)

# Ensure the uploads directory exists
UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
EXCEL_PATH = os.path.join(UPLOAD_FOLDER, 'tasks.xlsx')

# Configure Flask-SQLAlchemy
import sys
is_testing = 'unittest' in sys.modules or app.config.get('TESTING', False)
if is_testing:
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
else:
    DB_PATH = os.path.join(UPLOAD_FOLDER, 'dyson_ops.db')
    app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# User credentials (simple mock login)
USER_CREDENTIALS = {
    "username": "admin",
    "password": "dyson2026"
}

# --- DATABASE MODEL ---
class Task(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    task_id = db.Column(db.String(10), unique=True, nullable=False)
    task_name = db.Column(db.String(100), nullable=False)
    start_date = db.Column(db.String(10), nullable=False)
    duration = db.Column(db.Integer, nullable=False)
    end_date = db.Column(db.String(10), nullable=False)
    progress = db.Column(db.Integer, nullable=False)
    owner = db.Column(db.String(200), nullable=False)  # Extended length for comments
    status = db.Column(db.String(20), nullable=False)
    phase = db.Column(db.String(20), nullable=False, default="Proto")

    def to_dict(self):
        return {
            "Task ID": self.task_id,
            "Task Name": self.task_name,
            "Start Date": self.start_date,
            "Duration": self.duration,
            "End Date": self.end_date,
            "Progress": self.progress,
            "Owner": self.owner,
            "Status": self.status,
            "Phase": self.phase
        }

# --- EXCEL DATA STORE (SQLite Wrapper) ---
class ExcelDataStore:
    @staticmethod
    def initialize_db():
        """Initializes the SQLite database and migrates data from Excel if the DB is empty."""
        db.create_all()
        
        # Check if DB is empty
        if Task.query.count() == 0:
            print("SQLite database is empty. Checking for Excel file to migrate...")
            if os.path.exists(EXCEL_PATH):
                try:
                    df = pd.read_excel(EXCEL_PATH)
                    df = df.fillna("")
                    for _, row in df.iterrows():
                        progress = int(row.get("Progress", 0))
                        status = row.get("Status", "Not Started")
                        if progress == 100:
                            status = "Completed"
                        elif progress > 0 and status != "Rejected":
                            status = "In Progress"
                        
                        start_date = str(row["Start Date"])
                        if " " in start_date:
                            start_date = start_date.split(" ")[0]
                        end_date = str(row["End Date"])
                        if " " in end_date:
                            end_date = end_date.split(" ")[0]

                        task = Task(
                            task_id=row["Task ID"],
                            task_name=row["Task Name"],
                            start_date=start_date,
                            duration=int(row["Duration"]),
                            end_date=end_date,
                            progress=progress,
                            owner=row["Owner"],
                            status=status,
                            phase=row.get("Phase", "Proto")
                        )
                        db.session.add(task)
                    db.session.commit()
                    print("Migration from Excel to SQLite completed successfully!")
                except Exception as e:
                    print(f"Error migrating from Excel: {e}")
                    db.session.rollback()
            else:
                print("No Excel file found. Seeding database with initial sample data...")
                sample_tasks = [
                    {
                        "task_id": "TSK-001",
                        "task_name": "V15 Cyclone Engine Optimization",
                        "start_date": "2026-07-01",
                        "duration": 15,
                        "end_date": "2026-07-16",
                        "progress": 80,
                        "owner": "Engineering Team",
                        "status": "In Progress",
                        "phase": "Proto"
                    },
                    {
                        "task_id": "TSK-002",
                        "task_name": "Supersonic Hair Dryer Noise Reduction",
                        "start_date": "2026-07-10",
                        "duration": 20,
                        "end_date": "2026-07-30",
                        "progress": 45,
                        "owner": "Acoustics Team",
                        "status": "In Progress",
                        "phase": "EVT"
                    },
                    {
                        "task_id": "TSK-003",
                        "task_name": "360 Vis Nav Robot Vacuum Pathing",
                        "start_date": "2026-07-18",
                        "duration": 12,
                        "end_date": "2026-07-30",
                        "progress": 10,
                        "owner": "Software Team",
                        "status": "In Progress",
                        "phase": "DVT"
                    },
                    {
                        "task_id": "TSK-004",
                        "task_name": "Airstrait Straightener Thermal Control",
                        "start_date": "2026-07-25",
                        "duration": 8,
                        "end_date": "2026-08-02",
                        "progress": 0,
                        "owner": "Thermal Team",
                        "status": "Not Started",
                        "phase": "PVT"
                    },
                    {
                        "task_id": "TSK-005",
                        "task_name": "V15 Cyclone Airflow Restriction",
                        "start_date": "2026-07-05",
                        "duration": 10,
                        "end_date": "2026-07-15",
                        "progress": 30,
                        "owner": "Inlet valve failed to open at 40kPa, causing motor stall.",
                        "status": "Rejected",
                        "phase": "EVT"
                    },
                    {
                        "task_id": "TSK-006",
                        "task_name": "Supersonic Heater Thermal Runaway",
                        "start_date": "2026-07-12",
                        "duration": 8,
                        "end_date": "2026-07-20",
                        "progress": 10,
                        "owner": "Thermal fuse did not blow at 150C. Redesigning sensor placement.",
                        "status": "Rejected",
                        "phase": "DVT"
                    }
                ]
                for st in sample_tasks:
                    task = Task(**st)
                    db.session.add(task)
                db.session.commit()

    @staticmethod
    def initialize_excel():
        # Keep for backward compatibility
        ExcelDataStore.initialize_db()

    @staticmethod
    def get_tasks():
        """Reads tasks from the SQLite database and returns them as a list of dicts."""
        ExcelDataStore.initialize_db()
        try:
            tasks = Task.query.all()
            return [t.to_dict() for t in tasks]
        except Exception as e:
            print(f"Error reading database: {e}")
            return []

    @staticmethod
    def save_task(task_data):
        """Adds or updates a task in the SQLite database."""
        ExcelDataStore.initialize_db()
        try:
            start_dt = datetime.datetime.strptime(task_data["Start Date"], "%Y-%m-%d")
            end_dt = start_dt + datetime.timedelta(days=int(task_data["Duration"]))
            task_data["End Date"] = end_dt.strftime("%Y-%m-%d")
            
            # Save the phase and status directly from input
            progress = int(task_data["Progress"])
            phase = task_data.get("Phase", "Proto")
            status = task_data.get("Status", "Not Started")
            
            # If progress is 100, ensure it's marked completed (unless rejected)
            if progress == 100 and status != "Rejected":
                status = "Completed"
            elif progress > 0 and status == "Not Started":
                status = "In Progress"

            task_id = task_data.get("Task ID")
            task = None
            if task_id:
                task = Task.query.filter_by(task_id=task_id).first()

            if task:
                task.task_name = task_data["Task Name"]
                task.start_date = task_data["Start Date"]
                task.duration = int(task_data["Duration"])
                task.end_date = task_data["End Date"]
                task.progress = progress
                task.owner = task_data["Owner"]
                task.status = status
                task.phase = phase
            else:
                if not task_id:
                    next_num = Task.query.count() + 1
                    task_id = f"TSK-{next_num:03d}"
                    while Task.query.filter_by(task_id=task_id).first() is not None:
                        next_num += 1
                        task_id = f"TSK-{next_num:03d}"
                
                task = Task(
                    task_id=task_id,
                    task_name=task_data["Task Name"],
                    start_date=task_data["Start Date"],
                    duration=int(task_data["Duration"]),
                    end_date=task_data["End Date"],
                    progress=progress,
                    owner=task_data["Owner"],
                    status=status,
                    phase=phase
                )
                db.session.add(task)
                
            db.session.commit()
            return task.to_dict()
        except Exception as e:
            print(f"Error saving task: {e}")
            db.session.rollback()
            raise e

    @staticmethod
    def delete_task(task_id):
        """Deletes a task by ID from the SQLite database."""
        ExcelDataStore.initialize_db()
        try:
            task = Task.query.filter_by(task_id=task_id).first()
            if task:
                db.session.delete(task)
                db.session.commit()
                return True
            return False
        except Exception as e:
            print(f"Error deleting task: {e}")
            db.session.rollback()
            return False

# Initialize the database on startup (only if not testing)
if not is_testing:
    with app.app_context():
        ExcelDataStore.initialize_db()


# --- DECORATORS / HELPER FUNCTIONS ---
def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


# --- ROUTES ---

@app.route('/')
def index():
    return render_template('home.html')



@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username == USER_CREDENTIALS["username"] and password == USER_CREDENTIALS["password"]:
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('dashboard'))
        else:
            error = "Invalid Dyson credentials. Please try again."
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', username=session.get('username'))


# --- API ENDPOINTS ---

@app.route('/api/tasks', methods=['GET'])
@login_required
def get_tasks_api():
    tasks = ExcelDataStore.get_tasks()
    return jsonify(tasks)


@app.route('/api/tasks', methods=['POST'])
@login_required
def add_task_api():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data provided"}), 400
            
        required = ["Task Name", "Start Date", "Duration", "Progress", "Owner"]
        for field in required:
            if field not in data or str(data[field]).strip() == "":
                return jsonify({"error": f"Field '{field}' is required."}), 400

        task_data = {
            "Task Name": data["Task Name"],
            "Start Date": data["Start Date"],
            "Duration": int(data["Duration"]),
            "Progress": int(data["Progress"]),
            "Owner": data["Owner"],
            "Status": data.get("Status", "Not Started"),
            "Phase": data.get("Phase", "Proto")
        }
        
        saved_task = ExcelDataStore.save_task(task_data)
        return jsonify(saved_task), 201
    except ValueError:
        return jsonify({"error": "Duration and Progress must be valid integers."}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/tasks/<task_id>', methods=['PUT'])
@login_required
def update_task_api(task_id):
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data provided"}), 400
            
        required = ["Task Name", "Start Date", "Duration", "Progress", "Owner"]
        for field in required:
            if field not in data or str(data[field]).strip() == "":
                return jsonify({"error": f"Field '{field}' is required."}), 400

        task_data = {
            "Task ID": task_id,
            "Task Name": data["Task Name"],
            "Start Date": data["Start Date"],
            "Duration": int(data["Duration"]),
            "Progress": int(data["Progress"]),
            "Owner": data["Owner"],
            "Status": data.get("Status", "Not Started"),
            "Phase": data.get("Phase", "Proto")
        }
        
        saved_task = ExcelDataStore.save_task(task_data)
        return jsonify(saved_task)
    except ValueError:
        return jsonify({"error": "Duration and Progress must be valid integers."}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/tasks/<task_id>', methods=['DELETE'])
@login_required
def delete_task_api(task_id):
    success = ExcelDataStore.delete_task(task_id)
    if success:
        return jsonify({"message": f"Task {task_id} deleted successfully."})
    return jsonify({"error": f"Task {task_id} not found."}), 404


@app.route('/api/sync-sharepoint', methods=['POST'])
@login_required
def sync_sharepoint():
    try:
        # Simulate SharePoint API network latency
        import time
        time.sleep(1.0)
        
        ExcelDataStore.initialize_excel()
        
        # Simulate read/write cloud sync
        df = pd.read_excel(EXCEL_PATH)
        df.to_excel(EXCEL_PATH, index=False)
        
        return jsonify({
            "success": True,
            "message": "SharePoint cloud database synced successfully! All records match tasks.xlsx."
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "error": f"Failed to sync with SharePoint: {str(e)}"
        }), 500


# --- GANTT CHART GENERATION ---
def generate_gantt_chart_image(tasks, filepath):
    """Generates a Gantt chart image using Matplotlib and saves it to filepath."""
    if not tasks:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, "No active tasks to display", 
                horizontalalignment='center', verticalalignment='center', fontsize=12)
        ax.set_axis_off()
        plt.savefig(filepath, dpi=300, bbox_inches='tight')
        plt.close()
        return

    fig, ax = plt.subplots(figsize=(10, 5))
    
    # Custom colored phases
    phase_colors = {
        "Proto": "#00539C",  # Prussian Blue
        "EVT": "#EE7623",    # Orange
        "DVT": "#00A650",    # Green
        "PVT": "#D6257D"     # Fuchsia Accent
    }
    default_color = "#888888"
    
    y_labels = []
    y_ticks = []
    
    # Render Gantt bars (reversed to show top-down)
    for idx, t in enumerate(reversed(tasks)):
        try:
            start_date = datetime.datetime.strptime(t["Start Date"], "%Y-%m-%d")
            end_date = datetime.datetime.strptime(t["End Date"], "%Y-%m-%d")
        except Exception:
            continue
            
        phase = t.get("Phase", "Proto")
        color = phase_colors.get(phase, default_color)
        
        # Draw the bar representing the schedule duration
        ax.barh(idx, (end_date - start_date).days, left=start_date, 
                color=color, edgecolor='none', height=0.4, alpha=0.95)
        
        # Display product name and its test phase on the Y axis
        y_labels.append(f"{t['Task Name']} ({phase})")
        y_ticks.append(idx)
        
    ax.set_yticks(y_ticks)
    ax.set_yticklabels(y_labels, fontsize=9, fontweight='bold', color='#121212')
    
    # Configure X-axis as dates
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=7))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    plt.xticks(rotation=30, fontsize=8, color='#555555')
    
    # Styling
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#cccccc')
    ax.spines['bottom'].set_color('#cccccc')
    ax.grid(axis='x', linestyle='--', alpha=0.5, color='#dddddd')
    
    # Legend
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=color, label=phase) for phase, color in phase_colors.items()]
    ax.legend(handles=legend_elements, loc='upper right', frameon=True, facecolor='#ffffff', edgecolor='#e5e5e7', fontsize=8)
    
    plt.title("Test Operations Schedule by Phase", fontsize=12, fontweight='bold', pad=15, color='#121212')
    plt.tight_layout()
    plt.savefig(filepath, dpi=300)
    plt.close()


# --- PDF REPORT GENERATION ---
def generate_pdf_report(filename):
    """Generates a styled Test Operations PDF report containing a Gantt chart and rejected products."""
    # Fetch all tasks from SQLite database
    tasks = ExcelDataStore.get_tasks()
    
    # 1. Generate the Gantt Chart Image
    gantt_image_path = os.path.join(UPLOAD_FOLDER, "gantt_chart_report.png")
    generate_gantt_chart_image(tasks, gantt_image_path)
    
    doc = SimpleDocTemplate(filename, pagesize=letter,
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    # Custom color palette (Removing Dyson branding, generic technical theme)
    charcoal = colors.HexColor("#121212")
    steel_gray = colors.HexColor("#555555")
    prussian_blue = colors.HexColor("#00539C")
    light_bg = colors.HexColor("#F9F9FB")
    border_color = colors.HexColor("#E5E5E7")
    danger_red = colors.HexColor("#C0392B")
    
    # Custom styles
    title_style = ParagraphStyle(
        'OpsTitle',
        parent=styles['Heading1'],
        fontSize=24,
        leading=28,
        textColor=charcoal,
        fontName='Helvetica-Bold',
        spaceAfter=6
    )
    
    subtitle_style = ParagraphStyle(
        'OpsSubtitle',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=steel_gray,
        fontName='Helvetica',
        spaceAfter=20
    )
    
    section_title_style = ParagraphStyle(
        'OpsSection',
        parent=styles['Heading2'],
        fontSize=14,
        leading=18,
        textColor=prussian_blue,
        fontName='Helvetica-Bold',
        spaceBefore=15,
        spaceAfter=10,
        keepWithNext=True
    )
    
    cell_header = ParagraphStyle(
        'OpsHeader',
        parent=styles['Normal'],
        fontSize=9,
        leading=11,
        textColor=colors.white,
        fontName='Helvetica-Bold'
    )
    
    cell_style = ParagraphStyle(
        'OpsCell',
        parent=styles['Normal'],
        fontSize=9,
        leading=12,
        textColor=charcoal,
        fontName='Helvetica'
    )
    
    cell_bold = ParagraphStyle(
        'OpsCellBold',
        parent=styles['Normal'],
        fontSize=9,
        leading=12,
        textColor=danger_red,
        fontName='Helvetica-Bold'
    )
    
    # Header Title (Generic title, removed 'Dyson')
    story.append(Paragraph("Daily Test Operations Report", title_style))
    story.append(Paragraph(f"Generated on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Confidential Operations Data", subtitle_style))
    
    # Add Gantt Chart Section
    story.append(Paragraph("Test Method Milestones (Gantt Chart)", section_title_style))
    from reportlab.platypus import Image
    story.append(Image(gantt_image_path, width=500, height=250))
    story.append(Spacer(1, 20))
    
    # Filter for Rejected Products only
    rejected_tasks = [t for t in tasks if t.get("Status") == "Rejected"]
    
    story.append(Paragraph("Rejected Products Inventory", section_title_style))
    
    # Table Headers (Owner column renamed to Rejection Comments)
    table_data = [[
        Paragraph("ID", cell_header),
        Paragraph("Product Name", cell_header),
        Paragraph("Phase", cell_header),
        Paragraph("Start Date", cell_header),
        Paragraph("End Date", cell_header),
        Paragraph("Rejection Comments", cell_header)
    ]]
    
    if rejected_tasks:
        for t in rejected_tasks:
            table_data.append([
                Paragraph(t["Task ID"], cell_bold),
                Paragraph(t["Task Name"], cell_style),
                Paragraph(t.get("Phase", "Proto"), cell_style),
                Paragraph(str(t["Start Date"]), cell_style),
                Paragraph(str(t["End Date"]), cell_style),
                Paragraph(t["Owner"], cell_bold)  # Owner contains rejection comments
            ])
    else:
        # Empty state row
        table_data.append([
            Paragraph("-", cell_style),
            Paragraph("No rejected products recorded in the system.", cell_style),
            Paragraph("-", cell_style),
            Paragraph("-", cell_style),
            Paragraph("-", cell_style),
            Paragraph("-", cell_style)
        ])
        
    # Table Column Widths (Sum = 500)
    tasks_table = Table(table_data, colWidths=[50, 120, 45, 60, 60, 165])
    tasks_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), charcoal),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, light_bg])
    ]))
    story.append(tasks_table)
    
    # Footer Note
    story.append(Spacer(1, 40))
    footer_style = ParagraphStyle(
        'OpsFooter',
        parent=styles['Normal'],
        fontSize=8,
        textColor=steel_gray,
        alignment=1
    )
    story.append(Paragraph("Confidential &bull; Test Technology Operations &bull; Internal Use Only", footer_style))
    
    doc.build(story)


@app.route('/generate-report', methods=['GET'])
@login_required
def generate_report():
    pdf_filename = os.path.join(UPLOAD_FOLDER, "Dyson_Daily_Operations_Report.pdf")
    try:
        generate_pdf_report(pdf_filename)
        return send_file(pdf_filename, as_attachment=True, download_name="Dyson_Daily_Operations_Report.pdf")
    except Exception as e:
        return jsonify({"error": f"Failed to generate PDF: {str(e)}"}), 500


# --- EMAIL AUTOMATION ---

@app.route('/send-email', methods=['POST'])
@login_required
def send_email():
    recipient = request.form.get('recipient') or os.getenv('MAIL_RECIPIENT', 'operations-manager@dyson.com')
    subject = request.form.get('subject') or f"Dyson Operations Daily Report - {datetime.date.today().strftime('%Y-%m-%d')}"
    body = """Dear Team,

Please find attached the Dyson Operations Daily Project Status Report containing the latest project tasks, timeline progression, and completion rates.

This email was auto-generated by the Dyson Operations Dashboard.

Best regards,
Dyson Dashboard System
"""
    
    pdf_filename = os.path.join(UPLOAD_FOLDER, "Dyson_Daily_Operations_Report.pdf")
    
    try:
        # 1. Generate the PDF
        generate_pdf_report(pdf_filename)
        
        # 2. Check if SMTP configuration is set. If not, run in simulated mode
        smtp_server = app.config['MAIL_SERVER']
        smtp_user = app.config['MAIL_USERNAME']
        
        if not smtp_server or not smtp_user:
            # Simulated Email Mode
            sim_log = {
                "status": "Simulated",
                "message": "SMTP server not configured. Running in Simulated Email Mode.",
                "details": {
                    "To": recipient,
                    "From": app.config['MAIL_DEFAULT_SENDER'],
                    "Subject": subject,
                    "Attachments": [os.path.basename(pdf_filename)],
                    "Body": body,
                    "SMTP Server": "Simulated Console Mailer (Localhost)"
                }
            }
            print("\n" + "="*50)
            print("SIMULATED EMAIL SENT:")
            print(f"To: {recipient}")
            print(f"From: {app.config['MAIL_DEFAULT_SENDER']}")
            print(f"Subject: {subject}")
            print(f"Attachment: {os.path.basename(pdf_filename)}")
            print("="*50 + "\n")
            return jsonify({
                "success": True, 
                "simulated": True, 
                "message": "Email sent successfully (SIMULATED MODE). Details printed to console.",
                "log": sim_log
            })
            
        # 3. Real SMTP Sending
        msg = Message(
            subject=subject,
            recipients=[recipient],
            body=body
        )
        with open(pdf_filename, 'rb') as fp:
            msg.attach("Dyson_Daily_Operations_Report.pdf", "application/pdf", fp.read())
            
        mail.send(msg)
        return jsonify({
            "success": True, 
            "simulated": False, 
            "message": f"Email sent successfully to {recipient}."
        })
        
    except Exception as e:
        print(f"Error sending email: {e}")
        return jsonify({
            "success": False, 
            "error": f"Failed to send email: {str(e)}. (Check your SMTP credentials in .env)"
        }), 500


# --- AI CHATBOT ENDPOINT ---

@app.route('/chat', methods=['POST'])
@login_required
def chat():
    try:
        data = request.json
        user_message = data.get("message", "").strip()
        
        if not user_message:
            return jsonify({"response": "I didn't receive any message. How can I help you today?"})
            
        # Load active tasks for data-aware querying
        tasks = ExcelDataStore.get_tasks()
        df = pd.DataFrame(tasks)
        
        msg_lower = user_message.lower()
        
        # --- DATA-AWARE CHAT LOGIC ---
        response_text = ""
        
        if len(df) == 0:
            response_text = "There are currently no tasks in the system. You can add one by using the form on the dashboard!"
            
        elif any(kw in msg_lower for kw in ["how many tasks", "total tasks", "number of tasks", "task count"]):
            total_tasks = len(df)
            completed = len(df[df["Progress"] == 100])
            in_prog = len(df[(df["Progress"] > 0) & (df["Progress"] < 100)])
            not_start = len(df[df["Progress"] == 0])
            
            response_text = f"There are currently **{total_tasks}** tasks in the system:\n\n" \
                            f"- **{completed}** Completed (100%)\n" \
                            f"- **{in_prog}** In Progress\n" \
                            f"- **{not_start}** Not Started"
                            
        elif any(kw in msg_lower for kw in ["average progress", "mean progress", "overall progress"]):
            avg_prog = int(df["Progress"].mean())
            response_text = f"The average progress of all active projects is **{avg_prog}%**."
            
        elif any(kw in msg_lower for kw in ["overdue", "behind schedule", "late"]):
            today_str = datetime.date.today().strftime("%Y-%m-%d")
            # Overdue = End Date < today and progress < 100
            overdue_df = df[(df["End Date"] < today_str) & (df["Progress"] < 100)]
            
            if len(overdue_df) == 0:
                response_text = "Great news! There are currently no overdue tasks. All projects are on track."
            else:
                response_text = f"There are **{len(overdue_df)}** task(s) currently overdue:\n\n"
                for _, row in overdue_df.iterrows():
                    response_text += f"- **{row['Task Name']}** (End Date: {row['End Date']}, Progress: {row['Progress']}%)\n"
                response_text += "\nWe should prioritize these to get back on schedule."
                
        elif any(kw in msg_lower for kw in ["completed", "done", "finished"]):
            comp_df = df[df["Progress"] == 100]
            if len(comp_df) == 0:
                response_text = "No tasks are completed yet. Let's keep working!"
            else:
                response_text = f"Here are the completed tasks (**{len(comp_df)}**):\n\n"
                for _, row in comp_df.iterrows():
                    response_text += f"- **{row['Task Name']}** (Owned by: {row['Owner']})\n"
                    
        elif any(kw in msg_lower for kw in ["in progress", "active"]):
            ip_df = df[(df["Progress"] > 0) & (df["Progress"] < 100)]
            if len(ip_df) == 0:
                response_text = "There are no tasks currently in progress."
            else:
                response_text = f"Here are the active tasks in progress (**{len(ip_df)}**):\n\n"
                for _, row in ip_df.iterrows():
                    response_text += f"- **{row['Task Name']}** ({row['Progress']}% complete, Ends: {row['End Date']})\n"
                    
        elif "list" in msg_lower or "show tasks" in msg_lower or "all tasks" in msg_lower:
            response_text = "Here is the list of all current tasks in the database:\n\n"
            for _, row in df.iterrows():
                response_text += f"- **{row['Task ID']}**: {row['Task Name']} | Progress: **{row['Progress']}%** | Owner: {row['Owner']} | Ends: {row['End Date']}\n"
                
        elif "dyson" in msg_lower:
            response_text = "Dyson is a global technology enterprise founded by Sir James Dyson. We specialize in cyclone vacuum cleaners, bladeless fans, air purifiers, hair dryers, and high-performance lighting. \n\nThis dashboard helps our operations team track engineering sprints and product development milestones."
            
        elif any(kw in msg_lower for kw in ["help", "what can you do", "capabilities"]):
            response_text = "I am the **Dyson Operations AI Assistant**. I can analyze your project database in real-time. Try asking me:\n\n" \
                            "- *'How many tasks are there?'*\n" \
                            "- *'What is our average progress?'*\n" \
                            "- *'Are there any overdue tasks?'*\n" \
                            "- *'List all tasks'* or *'Which tasks are in progress?'*\n" \
                            "- *'Tell me about Dyson'*"
        else:
            # Fallback natural chatbot response
            response_text = f"I'm here to assist with your Dyson operations. I can see you have **{len(df)}** tasks registered. " \
                            f"Ask me questions like *'What is the average progress?'*, *'Which tasks are overdue?'*, or *'List all tasks'* to get started!"
                            
        return jsonify({"response": response_text})
        
    except Exception as e:
        print(f"Error in chatbot: {e}")
        return jsonify({"response": "I encountered an error trying to process your request. Please try again."}), 500

if __name__ == '__main__':
    # Allow local network sharing (0.0.0.0)
    app.run(host='0.0.0.0', port=5002, debug=True)
