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

# User credentials (simple mock login)
USER_CREDENTIALS = {
    "username": "admin",
    "password": "dyson2026"
}

# --- EXCEL DATA STORE ---
class ExcelDataStore:
    @staticmethod
    def initialize_excel():
        """Creates the Excel file with initial Dyson-themed sample data if it doesn't exist."""
        if not os.path.exists(EXCEL_PATH):
            sample_tasks = [
                {
                    "Task ID": "TSK-001",
                    "Task Name": "Dyson V15 Cyclone Engine Optimization",
                    "Start Date": "2026-07-01",
                    "Duration": 15,
                    "End Date": "2026-07-16",
                    "Progress": 80,
                    "Owner": "Engineering Team",
                    "Status": "In Progress"
                },
                {
                    "Task ID": "TSK-002",
                    "Task Name": "Supersonic Hair Dryer Noise Reduction",
                    "Start Date": "2026-07-10",
                    "Duration": 20,
                    "End Date": "2026-07-30",
                    "Progress": 45,
                    "Owner": "Acoustics Team",
                    "Status": "In Progress"
                },
                {
                    "Task ID": "TSK-003",
                    "Task Name": "360 Vis Nav Robot Vacuum Pathing",
                    "Start Date": "2026-07-18",
                    "Duration": 12,
                    "End Date": "2026-07-30",
                    "Progress": 10,
                    "Owner": "Software Team",
                    "Status": "In Progress"
                },
                {
                    "Task ID": "TSK-004",
                    "Task Name": "Airstrait Straightener Thermal Control",
                    "Start Date": "2026-07-25",
                    "Duration": 8,
                    "End Date": "2026-08-02",
                    "Progress": 0,
                    "Owner": "Thermal Team",
                    "Status": "Not Started"
                }
            ]
            df = pd.DataFrame(sample_tasks)
            df.to_excel(EXCEL_PATH, index=False)

    @staticmethod
    def get_tasks():
        """Reads tasks from the Excel file and returns them as a list of dicts."""
        ExcelDataStore.initialize_excel()
        try:
            df = pd.read_excel(EXCEL_PATH)
            # Fill NaNs to avoid JSON serialization issues
            df = df.fillna("")
            # Ensure proper types
            df["Duration"] = df["Duration"].astype(int)
            df["Progress"] = df["Progress"].astype(int)
            return df.to_dict(orient="records")
        except Exception as e:
            print(f"Error reading Excel: {e}")
            return []

    @staticmethod
    def save_task(task_data):
        """Adds or updates a task in the Excel sheet."""
        ExcelDataStore.initialize_excel()
        try:
            df = pd.read_excel(EXCEL_PATH)
            
            # Calculate End Date based on Start Date + Duration
            start_dt = datetime.datetime.strptime(task_data["Start Date"], "%Y-%m-%d")
            end_dt = start_dt + datetime.timedelta(days=int(task_data["Duration"]))
            task_data["End Date"] = end_dt.strftime("%Y-%m-%d")
            
            # Generate Task ID if not provided
            if not task_data.get("Task ID"):
                # Clean prefix and number
                next_num = len(df) + 1
                task_data["Task ID"] = f"TSK-{next_num:03d}"
            
            # Determine Status based on Progress
            progress = int(task_data["Progress"])
            if progress == 100:
                task_data["Status"] = "Completed"
            elif progress > 0:
                task_data["Status"] = "In Progress"
            else:
                task_data["Status"] = "Not Started"

            # Check if task already exists, update it, else append
            if task_data["Task ID"] in df["Task ID"].values:
                idx = df[df["Task ID"] == task_data["Task ID"]].index[0]
                for col, val in task_data.items():
                    df.at[idx, col] = val
            else:
                # Add new row
                new_row = pd.DataFrame([task_data])
                df = pd.concat([df, new_row], ignore_index=True)
                
            df.to_excel(EXCEL_PATH, index=False)
            return task_data
        except Exception as e:
            print(f"Error saving task to Excel: {e}")
            raise e

    @staticmethod
    def delete_task(task_id):
        """Deletes a task by ID from the Excel sheet."""
        ExcelDataStore.initialize_excel()
        try:
            df = pd.read_excel(EXCEL_PATH)
            if task_id in df["Task ID"].values:
                df = df[df["Task ID"] != task_id]
                df.to_excel(EXCEL_PATH, index=False)
                return True
            return False
        except Exception as e:
            print(f"Error deleting task: {e}")
            return False

# Initialize the Excel file on startup
ExcelDataStore.initialize_excel()


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
            "Owner": data["Owner"]
        }
        
        saved_task = ExcelDataStore.save_task(task_data)
        return jsonify(saved_task), 201
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


# --- PDF REPORT GENERATION ---

def generate_pdf_report(filename):
    """Generates a styled Dyson Operations PDF report using ReportLab."""
    doc = SimpleDocTemplate(filename, pagesize=letter,
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    # Custom Dyson color palette
    charcoal = colors.HexColor("#121212")
    steel_gray = colors.HexColor("#555555")
    prussian_blue = colors.HexColor("#00539C")
    iron_pink = colors.HexColor("#D6257D")
    light_bg = colors.HexColor("#F9F9FB")
    border_color = colors.HexColor("#E5E5E7")
    
    # Custom styles
    title_style = ParagraphStyle(
        'DysonTitle',
        parent=styles['Heading1'],
        fontSize=24,
        leading=28,
        textColor=charcoal,
        fontName='Helvetica-Bold',
        spaceAfter=6
    )
    
    subtitle_style = ParagraphStyle(
        'DysonSubtitle',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=steel_gray,
        fontName='Helvetica',
        spaceAfter=20
    )
    
    section_title_style = ParagraphStyle(
        'DysonSection',
        parent=styles['Heading2'],
        fontSize=14,
        leading=18,
        textColor=prussian_blue,
        fontName='Helvetica-Bold',
        spaceBefore=12,
        spaceAfter=8
    )
    
    cell_style = ParagraphStyle(
        'DysonCell',
        parent=styles['Normal'],
        fontSize=9,
        leading=12,
        textColor=charcoal,
        fontName='Helvetica'
    )
    
    cell_bold = ParagraphStyle(
        'DysonCellBold',
        parent=cell_style,
        fontName='Helvetica-Bold'
    )

    cell_header = ParagraphStyle(
        'DysonHeader',
        parent=cell_style,
        fontName='Helvetica-Bold',
        textColor=colors.white
    )

    # Document Header
    story.append(Paragraph("DYSON OPERATIONS DASHBOARD", title_style))
    current_date = datetime.date.today().strftime("%B %d, %Y")
    story.append(Paragraph(f"DAILY PROJECT STATUS REPORT &bull; GENERATED ON: {current_date}", subtitle_style))
    
    # Decorative colored bar (Dyson signature fuchsia line)
    d = Table([[""]], colWidths=[532], rowHeights=[3])
    d.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), iron_pink),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(d)
    story.append(Spacer(1, 15))
    
    # Load data
    tasks = ExcelDataStore.get_tasks()
    
    # Calculate statistics
    total_tasks = len(tasks)
    avg_progress = int(sum(t["Progress"] for t in tasks) / total_tasks) if total_tasks > 0 else 0
    completed_tasks = sum(1 for t in tasks if t["Progress"] == 100)
    in_progress_tasks = sum(1 for t in tasks if 0 < t["Progress"] < 100)
    
    # Summary Table
    summary_data = [
        [
            Paragraph("Total Tasks", cell_bold),
            Paragraph(str(total_tasks), cell_style),
            Paragraph("Average Progress", cell_bold),
            Paragraph(f"{avg_progress}%", cell_style)
        ],
        [
            Paragraph("Completed Tasks", cell_bold),
            Paragraph(str(completed_tasks), cell_style),
            Paragraph("In Progress Tasks", cell_bold),
            Paragraph(str(in_progress_tasks), cell_style)
        ]
    ]
    summary_table = Table(summary_data, colWidths=[130, 136, 130, 136])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), light_bg),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('INNERGRID', (0,0), (-1,-1), 0.5, border_color),
        ('BOX', (0,0), (-1,-1), 1, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    
    story.append(Paragraph("System Summary", section_title_style))
    story.append(summary_table)
    story.append(Spacer(1, 20))
    
    # Detailed Tasks Table
    story.append(Paragraph("Active Task Details", section_title_style))
    
    # Table Headers
    table_data = [[
        Paragraph("ID", cell_header),
        Paragraph("Task Name", cell_header),
        Paragraph("Start Date", cell_header),
        Paragraph("End Date", cell_header),
        Paragraph("Progress", cell_header),
        Paragraph("Owner", cell_header)
    ]]
    
    for t in tasks:
        table_data.append([
            Paragraph(t["Task ID"], cell_style),
            Paragraph(t["Task Name"], cell_style),
            Paragraph(str(t["Start Date"]), cell_style),
            Paragraph(str(t["End Date"]), cell_style),
            Paragraph(f"{t['Progress']}%", cell_bold if t['Progress'] == 100 else cell_style),
            Paragraph(t["Owner"], cell_style)
        ])
        
    tasks_table = Table(table_data, colWidths=[55, 170, 75, 75, 57, 100])
    tasks_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), charcoal),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        # Row styling alternates
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, light_bg])
    ]))
    story.append(tasks_table)
    
    # Footer Note
    story.append(Spacer(1, 40))
    footer_style = ParagraphStyle(
        'DysonFooter',
        parent=styles['Normal'],
        fontSize=8,
        textColor=steel_gray,
        alignment=1 # Centered
    )
    story.append(Paragraph("Confidential &bull; Dyson Technology Operations &bull; Internal Use Only", footer_style))
    
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
