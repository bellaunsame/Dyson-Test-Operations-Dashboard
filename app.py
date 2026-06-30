import os
import datetime
import sys
from functools import wraps
from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Headless mode for web servers
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# ReportLab imports for PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__)
load_dotenv()
app.secret_key = "super_secret_key_for_ops_operations"

# Configure App Data and Upload Folders
UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
EXCEL_PATH = os.path.join(UPLOAD_FOLDER, 'tasks.xlsx')

# Configure Flask-SQLAlchemy
is_testing = any('pytest' in arg or 'test' in arg for arg in sys.argv) or app.config.get('TESTING', False)
if is_testing:
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
else:
    DB_PATH = os.path.join(UPLOAD_FOLDER, 'ops_dashboard.db')
    app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
with app.app_context():
    db.create_all()


# User authentication credentials
USER_CREDENTIALS = {
    "username": "admin",
    "password": "admin2026"
}


def clean_dataframe_headers(df):
    """
    Cleans a pandas DataFrame by checking if the first row contains sub-header-like values
    (e.g., 'wk', 'day', 'qty') and merging them with the main header (columns).
    Handles merged/multi-row headers elegantly and returns the cleaned DataFrame.
    """
    if df.empty:
        return df

    first_row = df.iloc[0].tolist()
    has_sub_header = False
    sub_header_keywords = {'wk', 'week', 'day', 'qty', 'quantity', 'comment', 'method', 'number', 'start'}
    
    for val in first_row:
        if isinstance(val, str) and val.lower().strip() in sub_header_keywords:
            has_sub_header = True
            break
            
    if not has_sub_header:
        return df
        
    new_columns = []
    current_main_header = ""
    for col_name, sub_name in zip(df.columns, first_row):
        col_str = str(col_name).strip()
        sub_str = str(sub_name).strip() if not pd.isna(sub_name) else ""
        
        if not col_str.startswith("Unnamed:"):
            current_main_header = col_str
            
        if current_main_header and sub_str:
            combined = f"{current_main_header}_{sub_str}"
        elif current_main_header:
            combined = current_main_header
        elif sub_str:
            combined = sub_str
        else:
            combined = col_str
            
        new_columns.append(combined)
        
    df.columns = new_columns
    df = df.iloc[1:].reset_index(drop=True)
    return df

# --- DATABASE MODELS ---

class Project(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    description = db.Column(db.String(200), nullable=True)
    status = db.Column(db.String(20), default="Active")  # Active, Archived
    created_at = db.Column(db.String(20), nullable=False)

    def to_dict(self, row_count=None):
        if row_count is None:
            row_count = TestRecord.query.filter_by(project_name=self.name).count()
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description or "",
            "status": self.status,
            "created_at": self.created_at,
            "row_count": row_count
        }

class TestRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(50), nullable=False)
    category = db.Column(db.String(100), nullable=False)
    test_method = db.Column(db.String(100), nullable=False)
    test_number = db.Column(db.String(50), nullable=False)
    start_date = db.Column(db.String(10), nullable=False)

    # Proto Phase
    proto_weeks = db.Column(db.Integer, default=0)
    proto_days = db.Column(db.Integer, default=0)
    proto_qty = db.Column(db.Integer, default=0)

    # DVT Phase
    dvt_weeks = db.Column(db.Integer, default=0)
    dvt_days = db.Column(db.Integer, default=0)
    dvt_qty = db.Column(db.Integer, default=0)

    # EVT Phase
    evt_weeks = db.Column(db.Integer, default=0)
    evt_days = db.Column(db.Integer, default=0)
    evt_qty = db.Column(db.Integer, default=0)

    # PVT Phase
    pvt_weeks = db.Column(db.Integer, default=0)
    pvt_days = db.Column(db.Integer, default=0)
    pvt_qty = db.Column(db.Integer, default=0)

    comments = db.Column(db.String(500), nullable=True)
    defect_qty = db.Column(db.Integer, default=0)

    def to_dict(self):
        return {
            "id": self.id,
            "Project Name": self.project_name,
            "Category": self.category,
            "Test Method": self.test_method,
            "Test Number": self.test_number,
            "Start Date": self.start_date,
            "Proto Weeks": self.proto_weeks,
            "Proto Days": self.proto_days,
            "Proto Qty": self.proto_qty,
            "DVT Weeks": self.dvt_weeks,
            "DVT Days": self.dvt_days,
            "DVT Qty": self.dvt_qty,
            "EVT Weeks": self.evt_weeks,
            "EVT Days": self.evt_days,
            "EVT Qty": self.evt_qty,
            "PVT Weeks": self.pvt_weeks,
            "PVT Days": self.pvt_days,
            "PVT Qty": self.pvt_qty,
            "Comments": self.comments or "",
            "Defect Qty": self.defect_qty or 0
        }

class ExcelMapping(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(50), unique=True, nullable=False)
    file_path = db.Column(db.String(255), nullable=False)
    sheet_name = db.Column(db.String(100), nullable=False)
    mapping_json = db.Column(db.Text, nullable=False)

    def to_dict(self):
        import json
        return {
            "id": self.id,
            "project_name": self.project_name,
            "file_path": self.file_path,
            "sheet_name": self.sheet_name,
            "mapping": json.loads(self.mapping_json) if self.mapping_json else {}
        }

# --- SHAREPOINT SERVICE (Microsoft Graph API) ---
try:
    from office365.graph_client import GraphClient
    HAS_GRAPH_LIB = True
except ImportError:
    HAS_GRAPH_LIB = False

class TokenObject:
    def __init__(self, access_token):
        self.tokenType = "Bearer"
        self.accessToken = access_token

class SharePointService:
    @staticmethod
    def get_graph_client():
        tenant_id = os.getenv('TENANT_ID')
        client_id = os.getenv('SHAREPOINT_CLIENT_ID')
        client_secret = os.getenv('SHAREPOINT_CLIENT_SECRET')
        
        if not (tenant_id and client_id and client_secret) or not HAS_GRAPH_LIB:
            return None
            
        try:
            client = GraphClient(tenant=tenant_id).with_client_secret(client_id, client_secret)
            return client
        except Exception as e:
            print(f"Graph API Connection Error: {e}")
            return None

    @staticmethod
    def list_files():
        client = SharePointService.get_graph_client()
        site_url = os.getenv('SHAREPOINT_SITE_URL')
        doc_lib = os.getenv('SHAREPOINT_DOC_LIB', 'Test Data')
        
        # Fallback Mock Files if SharePoint/Graph is not configured or fails
        mock_files = [
            {"name": "Phase_Test_Method_Master_Data.xlsx", "path": "Phase_Test_Method_Master_Data.xlsx", "size": "11.3 KB", "modified": "2026-06-29"},
            {"name": "Ops_Milestones_893.xlsx", "path": "Ops_Milestones_893.xlsx", "size": "45 KB", "modified": "2026-06-29"},
            {"name": "Vacuum_990_Schedule.xlsx", "path": "Vacuum_990_Schedule.xlsx", "size": "58 KB", "modified": "2026-06-29"}
        ]
        
        if not client or not site_url:
            return mock_files
            
        try:
            site = client.sites.get_by_url(site_url)
            drives = site.drives
            client.load(drives)
            client.execute_query()
            
            target_drive = None
            for d in drives:
                if d.name.lower() == doc_lib.lower():
                    target_drive = d
                    break
                    
            if not target_drive:
                return mock_files
                
            items = target_drive.root.children
            client.load(items)
            client.execute_query()
            
            result = []
            for item in items:
                if item.file is not None:
                    # Parse modified date
                    modified_str = 'N/A'
                    if hasattr(item, 'last_modified_datetime') and item.last_modified_datetime:
                        modified_str = item.last_modified_datetime.strftime('%Y-%m-%d')
                    elif hasattr(item, 'lastModifiedDateTime') and item.lastModifiedDateTime:
                        modified_str = item.lastModifiedDateTime.strftime('%Y-%m-%d')
                        
                    result.append({
                        "name": item.name,
                        "path": item.name,
                        "size": f"{round(item.size / 1024, 1)} KB" if hasattr(item, 'size') and item.size else "Unknown",
                        "modified": modified_str
                    })
            return result if result else mock_files
        except Exception as e:
            print(f"Error listing Graph files: {e}")
            return mock_files

    @staticmethod
    def download_file(filename):
        client = SharePointService.get_graph_client()
        site_url = os.getenv('SHAREPOINT_SITE_URL')
        doc_lib = os.getenv('SHAREPOINT_DOC_LIB', 'Test Data')
        
        if not client or not site_url:
            # Fallback: check if we have a local copy in the uploads folder
            local_mock_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.exists(local_mock_path):
                with open(local_mock_path, 'rb') as f:
                    return f.read()
            # Otherwise generate a mock Excel file in memory with pandas
            import io
            output = io.BytesIO()
            df = pd.DataFrame([
                {
                    "Test Category": "Device: Hair Dryer",
                    "Method Name": "Stress Imaging",
                    "Ref Number": "TM-035128",
                    "Start": "2026-06-28",
                    "Proto_Wk": 3, "Proto_Day": 2, "Proto_Quantity": 15,
                    "DVT_Wk": 11, "DVT_Day": 2, "DVT_Quantity": 12,
                    "EVT_Wk": 3, "EVT_Day": 0, "EVT_Quantity": 10,
                    "PVT_Wk": 2, "PVT_Day": 0, "PVT_Quantity": 8,
                    "Defective_Units": 3,
                    "Rejections_Notes": "unit failure at 72hr soak"
                }
            ])
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            return output.getvalue()
            
        try:
            from urllib.parse import urlparse
            parsed = urlparse(site_url)
            hostname = parsed.netloc
            
            # Build the absolute URL to the file
            absolute_file_url = f"https://{hostname}/sites/TestOperationsDashboard/{doc_lib}/{filename}"
            
            file_item = client.shares.by_url(absolute_file_url).drive_item
            import io
            file_buffer = io.BytesIO()
            file_item.download(file_buffer)
            client.execute_query()
            return file_buffer.getvalue()
        except Exception as e:
            print(f"Error downloading Graph file: {e}")
            import io
            output = io.BytesIO()
            df = pd.DataFrame([
                {
                    "Test Category": "Device: Hair Dryer",
                    "Method Name": "Stress Imaging",
                    "Ref Number": "TM-035128",
                    "Start": "2026-06-28",
                    "Proto_Wk": 3, "Proto_Day": 2, "Proto_Quantity": 15,
                    "DVT_Wk": 11, "DVT_Day": 2, "DVT_Quantity": 12,
                    "EVT_Wk": 3, "EVT_Day": 0, "EVT_Quantity": 10,
                    "PVT_Wk": 2, "PVT_Day": 0, "PVT_Quantity": 8,
                    "Defective_Units": 3,
                    "Rejections_Notes": "unit failure at 72hr soak"
                }
            ])
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            return output.getvalue()

# --- EXCEL DATA STORE (SQLite Wrapper) ---
class ExcelDataStore:
    @staticmethod
    def initialize_db():
        """Initializes the SQLite database and migrates schema."""
        db.create_all()

        # ── Schema migration: add defect_qty column if it doesn't exist ──
        try:
            db.session.execute(db.text('ALTER TABLE test_record ADD COLUMN defect_qty INTEGER DEFAULT 0'))
            db.session.commit()
        except Exception:
            db.session.rollback()  # Column already exists — safe to ignore

    @staticmethod
    def initialize_excel():
        ExcelDataStore.initialize_db()

    @staticmethod
    def get_tasks():
        """Compat wrapper for legacy test calls."""
        ExcelDataStore.initialize_db()
        return [r.to_dict() for r in TestRecord.query.all()]

    @staticmethod
    def save_task(data):
        """Compat wrapper for legacy test calls."""
        ExcelDataStore.initialize_db()
        # Convert legacy/flat keys to TestRecord structure
        record_data = {
            "project_name": data.get("Project Name", "893"),
            "category": data.get("Category", data.get("Product Name", "Device")),
            "test_method": data.get("Test Method", "Testing"),
            "test_number": data.get("Test Number", "TM-001"),
            "start_date": data.get("Start Date", "2026-07-01"),
            "proto_weeks": int(data.get("Proto Weeks", 0)),
            "proto_days": int(data.get("Proto Days", 0)),
            "proto_qty": int(data.get("Proto Qty", data.get("Qty", 0))),
            "dvt_weeks": int(data.get("DVT Weeks", 0)),
            "dvt_days": int(data.get("DVT Days", 0)),
            "dvt_qty": int(data.get("DVT Qty", 0)),
            "evt_weeks": int(data.get("EVT Weeks", 0)),
            "evt_days": int(data.get("EVT Days", 0)),
            "evt_qty": int(data.get("EVT Qty", 0)),
            "pvt_weeks": int(data.get("PVT Weeks", data.get("pvt_weeks", 0))),
            "pvt_days": int(data.get("PVT Days", 0)),
            "pvt_qty": int(data.get("PVT Qty", 0)),
            "comments": data.get("Comments", "")
        }
        
        record_id = data.get("Product ID") # Compat ID
        record = None
        if record_id:
            record = TestRecord.query.get(record_id)
        
        if not record:
            record = TestRecord(**record_data)
            db.session.add(record)
        else:
            for k, v in record_data.items():
                setattr(record, k, v)
        
        db.session.commit()
        return record.to_dict()

    @staticmethod
    def delete_task(record_id):
        """Compat wrapper for legacy test calls."""
        ExcelDataStore.initialize_db()
        record = TestRecord.query.get(record_id)
        if record:
            db.session.delete(record)
            db.session.commit()
            return True
        return False

# Initialize the database on startup (only if not testing)
if not is_testing:
    with app.app_context():
        ExcelDataStore.initialize_db()

# Ensure DB is always ready even when Flask reloader spawns a child process
_db_initialized = False

@app.before_request
def ensure_db():
    global _db_initialized
    if not _db_initialized:
        ExcelDataStore.initialize_db()
        _db_initialized = True

# --- DECORATORS / HELPER FUNCTIONS ---
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- VIEWS ---

@app.route('/')
@login_required
def index():
    return redirect(url_for('tables'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username == USER_CREDENTIALS["username"] and password == USER_CREDENTIALS["password"]:
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('tables'))
        else:
            error = "Invalid credentials. Please try again."
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    return redirect(url_for('chart_page'))

@app.route('/chart')
@login_required
def chart_page():
    return render_template('gantt.html', username=session.get('username'), active_page='chart')

@app.route('/tables')
@login_required
def tables():
    return render_template('tables.html', username=session.get('username'), active_page='tables')

@app.route('/import')
@login_required
def import_page():
    return render_template('import.html', username=session.get('username'), active_page='import')

@app.route('/tables/<project_name>')
@login_required
def project_detail(project_name):
    project = Project.query.filter_by(name=project_name).first()
    if not project:
        return redirect(url_for('tables'))
    return render_template('project_detail.html', project=project, username=session.get('username'), active_page='tables')

# --- API ENDPOINTS ---

@app.route('/api/projects', methods=['GET', 'POST'])
@login_required
def api_projects():
    if request.method == 'GET':
        status_filter = request.args.get('status', 'Active')
        projects = Project.query.filter_by(status=status_filter).all()
        
        # Single query to fetch row counts for all projects (avoys N+1 query issue)
        counts = db.session.query(
            TestRecord.project_name, 
            db.func.count(TestRecord.id)
        ).group_by(TestRecord.project_name).all()
        counts_map = {project_name: count for project_name, count in counts}
        
        return jsonify([p.to_dict(row_count=counts_map.get(p.name, 0)) for p in projects])
    
    # Create project
    data = request.json
    if not data or 'name' not in data or not str(data['name']).strip():
        return jsonify({"error": "Project name is required"}), 400
        
    name = str(data['name']).strip()
    if Project.query.filter_by(name=name).first():
        return jsonify({"error": f"Project '{name}' already exists"}), 400
        
    today_str = datetime.date.today().strftime('%m/%d/%Y')
    proj = Project(
        name=name,
        description=data.get('description', ''),
        status='Active',
        created_at=today_str
    )
    db.session.add(proj)
    db.session.commit()
    return jsonify(proj.to_dict(0)), 201

@app.route('/api/projects/<project_name>', methods=['PUT', 'DELETE'])
@login_required
def api_project_detail(project_name):
    proj = Project.query.filter_by(name=project_name).first()
    if not proj:
        return jsonify({"error": "Project not found"}), 404
        
    if request.method == 'DELETE':
        # Delete project and its records
        TestRecord.query.filter_by(project_name=project_name).delete()
        db.session.delete(proj)
        db.session.commit()
        return jsonify({"message": f"Project {project_name} deleted successfully"})
        
    # Update project (Rename or Archive)
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
        
    if 'name' in data and data['name'] != proj.name:
        new_name = str(data['name']).strip()
        if Project.query.filter_by(name=new_name).first():
            return jsonify({"error": f"Project '{new_name}' already exists"}), 400
        # Cascade update rows
        TestRecord.query.filter_by(project_name=proj.name).update({"project_name": new_name})
        proj.name = new_name
        
    if 'description' in data:
        proj.description = data['description']
    if 'status' in data:
        proj.status = data['status']
        
    db.session.commit()
    return jsonify(proj.to_dict())

@app.route('/api/projects/<project_name>/rows', methods=['GET', 'POST'])
@login_required
def api_project_rows(project_name):
    if request.method == 'GET':
        records = TestRecord.query.filter_by(project_name=project_name).all()
        return jsonify([r.to_dict() for r in records])
        
    # Add Row
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
        
    required = ["Category", "Test Method", "Test Number", "Start Date"]
    for f in required:
        if f not in data or not str(data[f]).strip():
            return jsonify({"error": f"Field '{f}' is required"}), 400
            
    rec = TestRecord(
        project_name=project_name,
        category=data["Category"],
        test_method=data["Test Method"],
        test_number=data["Test Number"],
        start_date=data["Start Date"],
        proto_weeks=int(data.get("Proto Weeks", 0)),
        proto_days=int(data.get("Proto Days", 0)),
        proto_qty=int(data.get("Proto Qty", 0)),
        dvt_weeks=int(data.get("DVT Weeks", 0)),
        dvt_days=int(data.get("DVT Days", 0)),
        dvt_qty=int(data.get("DVT Qty", 0)),
        evt_weeks=int(data.get("EVT Weeks", 0)),
        evt_days=int(data.get("EVT Days", 0)),
        evt_qty=int(data.get("EVT Qty", 0)),
        pvt_weeks=int(data.get("PVT Weeks", 0)),
        pvt_days=int(data.get("PVT Days", 0)),
        pvt_qty=int(data.get("PVT Qty", 0)),
        comments=data.get("Comments", ""),
        defect_qty=int(data.get("Defect Qty", 0))
    )
    db.session.add(rec)
    db.session.commit()
    return jsonify(rec.to_dict()), 201

@app.route('/api/projects/<project_name>/rows/<int:row_id>', methods=['PUT', 'DELETE'])
@login_required
def api_project_row_detail(project_name, row_id):
    rec = TestRecord.query.filter_by(project_name=project_name, id=row_id).first()
    if not rec:
        return jsonify({"error": "Row not found"}), 404
        
    if request.method == 'DELETE':
        db.session.delete(rec)
        db.session.commit()
        return jsonify({"message": "Row deleted successfully"})
        
    # Update Row
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
        
    rec.category = data.get("Category", rec.category)
    rec.test_method = data.get("Test Method", rec.test_method)
    rec.test_number = data.get("Test Number", rec.test_number)
    rec.start_date = data.get("Start Date", rec.start_date)
    
    rec.proto_weeks = int(data.get("Proto Weeks", rec.proto_weeks))
    rec.proto_days = int(data.get("Proto Days", rec.proto_days))
    rec.proto_qty = int(data.get("Proto Qty", rec.proto_qty))
    
    rec.dvt_weeks = int(data.get("DVT Weeks", rec.dvt_weeks))
    rec.dvt_days = int(data.get("DVT Days", rec.dvt_days))
    rec.dvt_qty = int(data.get("DVT Qty", rec.dvt_qty))
    
    rec.evt_weeks = int(data.get("EVT Weeks", rec.evt_weeks))
    rec.evt_days = int(data.get("EVT Days", rec.evt_days))
    rec.evt_qty = int(data.get("EVT Qty", rec.evt_qty))
    
    rec.pvt_weeks = int(data.get("PVT Weeks", rec.pvt_weeks))
    rec.pvt_days = int(data.get("PVT Days", rec.pvt_days))
    rec.pvt_qty = int(data.get("PVT Qty", rec.pvt_qty))
    
    rec.comments = data.get("Comments", rec.comments)
    rec.defect_qty = int(data.get("Defect Qty", rec.defect_qty or 0))
    
    db.session.commit()
    return jsonify(rec.to_dict())

# --- NEW SHAREPOINT & EXCEL COLUMN MAPPING ENDPOINTS ---

@app.route('/api/sharepoint/browse', methods=['GET'])
@login_required
def api_sharepoint_browse():
    """Lists Excel files available in the SharePoint Document Library."""
    try:
        files = SharePointService.list_files()
        return jsonify(files)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/sharepoint/headers', methods=['POST'])
@login_required
def api_sharepoint_headers():
    """Downloads the selected Excel file from SharePoint and extracts its worksheets and headers from the first sheet."""
    data = request.json
    if not data or 'file_path' not in data:
        return jsonify({"error": "No file_path provided"}), 400
        
    file_path = data['file_path']
    try:
        file_content = SharePointService.download_file(file_path)
        
        import io
        xl = pd.ExcelFile(io.BytesIO(file_content))
        sheets = xl.sheet_names
        
        if not sheets:
            return jsonify({"error": "No worksheets found in the Excel file"}), 400
            
        # Parse headers from the first sheet
        df = xl.parse(sheets[0], nrows=5)
        df = clean_dataframe_headers(df)
        headers = [str(col) for col in df.columns.tolist()]
        
        return jsonify({
            "success": True,
            "sheets": sheets,
            "headers": headers,
            "file_path": file_path
        })
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel file: {str(e)}"}), 500

@app.route('/api/sharepoint/sync', methods=['POST'])
@login_required
def api_sharepoint_sync():
    """Saves the global column mapping and imports the Excel data from EACH sheet as a separate project."""
    import json
    data = request.json
    if not data or 'file_path' not in data or 'mapping' not in data:
        return jsonify({"error": "Missing required mapping configuration"}), 400
        
    file_path = data['file_path']
    mapping = data['mapping']
    
    try:
        # Save or update global ExcelMapping in database
        existing_mapping = ExcelMapping.query.filter_by(project_name='GLOBAL_SYNC').first()
        if existing_mapping:
            existing_mapping.file_path = file_path
            existing_mapping.sheet_name = 'ALL_SHEETS'
            existing_mapping.mapping_json = json.dumps(mapping)
        else:
            new_mapping = ExcelMapping(
                project_name='GLOBAL_SYNC',
                file_path=file_path,
                sheet_name='ALL_SHEETS',
                mapping_json=json.dumps(mapping)
            )
            db.session.add(new_mapping)
        db.session.commit()
            
        # Download and parse Excel file
        file_content = SharePointService.download_file(file_path)
        
        import io
        xl = pd.ExcelFile(io.BytesIO(file_content))
        all_sheets = xl.sheet_names
        
        synced_projects = []
        total_records_imported = 0
        
        for sheet_name in all_sheets:
            # Skip "Master Data" sheet (case-insensitive)
            if sheet_name.lower().strip() == 'master data':
                continue
                
            # Create or find Project
            proj = Project.query.filter_by(name=sheet_name).first()
            if not proj:
                proj = Project(
                    name=sheet_name,
                    description=f"Imported from SharePoint sheet {sheet_name}",
                    status="Active",
                    created_at=datetime.date.today().strftime('%m/%d/%Y')
                )
                db.session.add(proj)
                db.session.commit()
                
            # Clear existing test records for this project
            TestRecord.query.filter_by(project_name=sheet_name).delete()
            
            # Parse worksheet
            df = xl.parse(sheet_name)
            df = clean_dataframe_headers(df)
            
            imported_count = 0
            for index, row in df.iterrows():
                def get_val(field, default=None, is_int=False):
                    col = mapping.get(field)
                    if not col or col not in df.columns:
                         return default
                    val = row[col]
                    if pd.isna(val):
                        return default
                    if is_int:
                        try:
                            return int(float(val))
                        except ValueError:
                            return 0
                    return str(val)
                    
                # Parse start date
                start_date_val = get_val("Start Date", "")
                if start_date_val:
                    try:
                        col_mapping = mapping.get("Start Date")
                        raw_date = row[col_mapping]
                        if hasattr(raw_date, 'strftime'):
                            start_date_val = raw_date.strftime('%Y-%m-%d')
                        else:
                            start_date_val = pd.to_datetime(start_date_val).strftime('%Y-%m-%d')
                    except Exception:
                        start_date_val = datetime.date.today().strftime('%Y-%m-%d')
                else:
                    start_date_val = datetime.date.today().strftime('%Y-%m-%d')
    
                rec = TestRecord(
                    project_name=sheet_name,
                    category=get_val("Category", "General"),
                    test_method=get_val("Test Method", "Testing"),
                    test_number=get_val("Test Number", "TM-000"),
                    start_date=start_date_val,
                    
                    proto_weeks=get_val("Proto Weeks", 0, is_int=True),
                    proto_days=get_val("Proto Days", 0, is_int=True),
                    proto_qty=get_val("Proto Qty", 0, is_int=True),
                    
                    dvt_weeks=get_val("DVT Weeks", 0, is_int=True),
                    dvt_days=get_val("DVT Days", 0, is_int=True),
                    dvt_qty=get_val("DVT Qty", 0, is_int=True),
                    
                    evt_weeks=get_val("EVT Weeks", 0, is_int=True),
                    evt_days=get_val("EVT Days", 0, is_int=True),
                    evt_qty=get_val("EVT Qty", 0, is_int=True),
                    
                    pvt_weeks=get_val("PVT Weeks", 0, is_int=True),
                    pvt_days=get_val("PVT Days", 0, is_int=True),
                    pvt_qty=get_val("PVT Qty", 0, is_int=True),
                    
                    defect_qty=get_val("Defect Qty", 0, is_int=True),
                    comments=get_val("Comments", "")
                )
                db.session.add(rec)
                imported_count += 1
                
            total_records_imported += imported_count
            synced_projects.append(sheet_name)
            
        db.session.commit()
        return jsonify({
            "success": True,
            "message": f"Successfully synced {len(synced_projects)} projects with {total_records_imported} records.",
            "projects": synced_projects,
            "count": total_records_imported
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Failed to sync SharePoint Excel sheets: {str(e)}"}), 500

@app.route('/api/sync-sharepoint', methods=['POST'])
@login_required
def sync_sharepoint():
    """Triggers a real sync of all Excel sheets from SharePoint using the saved global configuration."""
    import json
    # Get saved global mapping
    mapping_record = ExcelMapping.query.filter_by(project_name='GLOBAL_SYNC').first()
    if not mapping_record:
        return jsonify({
            "success": False,
            "error": "No saved column mapping. Please configure the columns first using the SharePoint Extractor in the sidebar."
        }), 400
        
    file_path = mapping_record.file_path
    config = json.loads(mapping_record.mapping_json) if mapping_record.mapping_json else {}
    selected_sheets = config.get("selected_sheets", [])
    removed_cols_map = config.get("removed_columns", {})
    
    try:
        file_content = SharePointService.download_file(file_path)
        
        import io
        xl = pd.ExcelFile(io.BytesIO(file_content))
        
        synced_projects = []
        total_records_imported = 0
        
        for sheet_name in selected_sheets:
            if sheet_name not in xl.sheet_names:
                continue
                
            # Parse sheet
            df = xl.parse(sheet_name)
            df = clean_dataframe_headers(df)
            
            # Filter out removed columns
            removed_cols = removed_cols_map.get(sheet_name, [])
            active_cols = [col for col in df.columns if str(col) not in removed_cols]
            
            # Perform automatic fuzzy mapping on remaining columns
            mapping = {}
            fuzzy_rules = {
                "category": ['category', 'test category', 'group', 'device', 'product'],
                "test_method": ['method', 'test method', 'method name', 'name'],
                "test_number": ['number', 'test number', 'test #', 'ref', 'ref number', 'code'],
                "start_date": ['start', 'start date', 'date', 'timeline start'],
                "defect_qty": ['defect', 'defect qty', 'defective', 'defects', 'defect quantity', 'defective units'],
                "comments": ['comments', 'rejections', 'notes', 'rejection comment/s', 'comment'],
                
                "proto_weeks": ['proto wk', 'proto week', 'proto_wk', 'proto_weeks', 'proto weeks', 'week1'],
                "proto_days": ['proto day', 'proto_day', 'proto_days', 'proto days', 'day1'],
                "proto_qty": ['proto qty', 'proto_qty', 'proto_quantity', 'proto quantity', 'proto_qty', 'proto1'],
                
                "dvt_weeks": ['dvt wk', 'dvt week', 'dvt_wk', 'dvt_weeks', 'dvt weeks', 'week2'],
                "dvt_days": ['dvt day', 'dvt_day', 'dvt_days', 'dvt days', 'day3'],
                "dvt_qty": ['dvt qty', 'dvt_qty', 'dvt_quantity', 'dvt quantity', 'dvt_qty', 'dvt1'],
                
                "evt_weeks": ['evt wk', 'evt week', 'evt_wk', 'evt_weeks', 'evt weeks', 'week4'],
                "evt_days": ['evt day', 'evt_day', 'evt_days', 'evt days', 'day5'],
                "evt_qty": ['evt qty', 'evt_qty', 'evt_quantity', 'evt quantity', 'evt_qty', 'evt1'],
                
                "pvt_weeks": ['pvt wk', 'pvt week', 'pvt_wk', 'pvt_weeks', 'pvt weeks', 'week6'],
                "pvt_days": ['pvt day', 'pvt_day', 'pvt_days', 'pvt days', 'day7'],
                "pvt_qty": ['pvt qty', 'pvt_qty', 'pvt_quantity', 'pvt quantity', 'pvt_qty', 'pvt1'],
            }
            
            for field, rules in fuzzy_rules.items():
                for col in active_cols:
                    col_lower = str(col).lower().strip()
                    if any(r == col_lower or r in col_lower for r in rules):
                        mapping[field] = str(col)
                        break
                        
            # Create or find Project
            proj = Project.query.filter_by(name=sheet_name).first()
            if not proj:
                proj = Project(
                    name=sheet_name,
                    description=f"Imported from SharePoint sheet {sheet_name}",
                    status="Active",
                    created_at=datetime.date.today().strftime('%m/%d/%Y')
                )
                db.session.add(proj)
                db.session.commit()
                
            # Clear existing test records for this project
            TestRecord.query.filter_by(project_name=sheet_name).delete()
            
            imported_count = 0
            for index, row in df.iterrows():
                def get_val(field, default=None, is_int=False):
                    col = mapping.get(field)
                    if not col or col not in df.columns:
                         return default
                    val = row[col]
                    if pd.isna(val):
                        return default
                    if is_int:
                        try:
                            return int(float(val))
                        except ValueError:
                            return 0
                    return str(val)
                    
                # Parse start date
                start_date_val = get_val("start_date", "")
                if start_date_val:
                    try:
                        col_mapping = mapping.get("start_date")
                        raw_date = row[col_mapping]
                        if hasattr(raw_date, 'strftime'):
                            start_date_val = raw_date.strftime('%Y-%m-%d')
                        else:
                            start_date_val = pd.to_datetime(start_date_val).strftime('%Y-%m-%d')
                    except Exception:
                        start_date_val = datetime.date.today().strftime('%Y-%m-%d')
                else:
                    start_date_val = datetime.date.today().strftime('%Y-%m-%d')
    
                rec = TestRecord(
                    project_name=sheet_name,
                    category=get_val("category", "General"),
                    test_method=get_val("test_method", "Testing"),
                    test_number=get_val("test_number", "TM-000"),
                    start_date=start_date_val,
                    
                    proto_weeks=get_val("proto_weeks", 0, is_int=True),
                    proto_days=get_val("proto_days", 0, is_int=True),
                    proto_qty=get_val("proto_qty", 0, is_int=True),
                    
                    dvt_weeks=get_val("dvt_weeks", 0, is_int=True),
                    dvt_days=get_val("dvt_days", 0, is_int=True),
                    dvt_qty=get_val("dvt_qty", 0, is_int=True),
                    
                    evt_weeks=get_val("evt_weeks", 0, is_int=True),
                    evt_days=get_val("evt_days", 0, is_int=True),
                    evt_qty=get_val("evt_qty", 0, is_int=True),
                    
                    pvt_weeks=get_val("pvt_weeks", 0, is_int=True),
                    pvt_days=get_val("pvt_days", 0, is_int=True),
                    pvt_qty=get_val("pvt_qty", 0, is_int=True),
                    
                    defect_qty=get_val("defect_qty", 0, is_int=True),
                    comments=get_val("comments", "")
                )
                db.session.add(rec)
                imported_count += 1
                
            total_records_imported += imported_count
            synced_projects.append(sheet_name)
            
        db.session.commit()
        return jsonify({
            "success": True,
            "message": f"SharePoint sync complete — synced {len(synced_projects)} projects with {total_records_imported} records.",
            "projects": synced_projects,
            "count": total_records_imported
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({"success": False, "error": f"Failed to sync SharePoint Excel: {str(e)}"}), 500

# --- LOCAL FILE UPLOAD ENDPOINT ---

@app.route('/api/local/preview-all', methods=['POST'])
@login_required
def api_local_preview_all():
    """Accepts a locally uploaded Excel file and returns columns and first 100 rows of all sheets."""
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files['file']
    if uploaded_file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    filename = uploaded_file.filename
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({"error": "Only Excel files (.xlsx, .xls) are supported"}), 400

    try:
        import io
        file_bytes = uploaded_file.read()
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        sheets = xl.sheet_names

        result = {}
        for sheet in sheets:
            if sheet.lower().strip() == 'master data':
                continue

            df = xl.parse(sheet, nrows=100)
            df = clean_dataframe_headers(df)
            columns = [str(col) for col in df.columns.tolist()]

            rows = []
            for _, row in df.iterrows():
                row_vals = []
                for val in row.tolist():
                    if pd.isna(val):
                        row_vals.append("")
                    else:
                        if isinstance(val, (datetime.date, datetime.datetime)):
                            row_vals.append(val.strftime('%Y-%m-%d'))
                        else:
                            row_vals.append(val)
                rows.append(row_vals)

            result[sheet] = {"columns": columns, "rows": rows}

        return jsonify({
            "success": True,
            "file_name": filename,
            "file_path": f"local://{filename}",
            "sheets": result
        })
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel file: {str(e)}"}), 500


# --- SHARED TRANSFORMATION HELPER ---

def _import_sheets_to_db(xl, selected_sheets, removed_cols_map, file_path):
    """
    Core transformation helper. Given a pd.ExcelFile, list of selected sheet names,
    a dict of removed columns per sheet, and the original file_path string,
    this function creates/updates Project records and imports TestRecord rows.
    Returns (synced_projects, total_records_imported).
    """
    fuzzy_rules = {
        "category": ['category', 'test category', 'group', 'device', 'product'],
        "test_method": ['method', 'test method', 'method name', 'name'],
        "test_number": ['number', 'test number', 'test #', 'ref', 'ref number', 'code'],
        "start_date": ['start', 'start date', 'date', 'timeline start'],
        "defect_qty": ['defect', 'defect qty', 'defective', 'defects', 'defect quantity', 'defective units'],
        "comments": ['comments', 'rejections', 'notes', 'rejection comment/s', 'comment'],
        "proto_weeks": ['proto wk', 'proto week', 'proto_wk', 'proto_weeks', 'proto weeks', 'week1'],
        "proto_days": ['proto day', 'proto_day', 'proto_days', 'proto days', 'day1'],
        "proto_qty": ['proto qty', 'proto_qty', 'proto_quantity', 'proto quantity', 'proto1'],
        "dvt_weeks": ['dvt wk', 'dvt week', 'dvt_wk', 'dvt_weeks', 'dvt weeks', 'week2'],
        "dvt_days": ['dvt day', 'dvt_day', 'dvt_days', 'dvt days', 'day3'],
        "dvt_qty": ['dvt qty', 'dvt_qty', 'dvt_quantity', 'dvt quantity', 'dvt1'],
        "evt_weeks": ['evt wk', 'evt week', 'evt_wk', 'evt_weeks', 'evt weeks', 'week4'],
        "evt_days": ['evt day', 'evt_day', 'evt_days', 'evt days', 'day5'],
        "evt_qty": ['evt qty', 'evt_qty', 'evt_quantity', 'evt quantity', 'evt1'],
        "pvt_weeks": ['pvt wk', 'pvt week', 'pvt_wk', 'pvt_weeks', 'pvt weeks', 'week6'],
        "pvt_days": ['pvt day', 'pvt_day', 'pvt_days', 'pvt days', 'day7'],
        "pvt_qty": ['pvt qty', 'pvt_qty', 'pvt_quantity', 'pvt quantity', 'pvt1'],
    }

    synced_projects = []
    total_records_imported = 0

    for sheet_name in selected_sheets:
        if sheet_name not in xl.sheet_names:
            continue

        df = xl.parse(sheet_name)
        df = clean_dataframe_headers(df)
        removed_cols = removed_cols_map.get(sheet_name, [])
        active_cols = [col for col in df.columns if str(col) not in removed_cols]

        # Auto fuzzy mapping
        mapping = {}
        for field, rules in fuzzy_rules.items():
            for col in active_cols:
                col_lower = str(col).lower().strip()
                if any(r == col_lower or r in col_lower for r in rules):
                    mapping[field] = str(col)
                    break

        # Upsert Project
        proj = Project.query.filter_by(name=sheet_name).first()
        if not proj:
            proj = Project(
                name=sheet_name,
                description=f"Imported from {file_path}",
                status="Active",
                created_at=datetime.date.today().strftime('%m/%d/%Y')
            )
            db.session.add(proj)
        else:
            proj.description = f"Imported from {file_path}"
        db.session.commit()

        # Clear and reimport rows
        TestRecord.query.filter_by(project_name=sheet_name).delete()

        imported_count = 0
        for index, row in df.iterrows():
            def get_val(field, default=None, is_int=False):
                col = mapping.get(field)
                if not col or col not in df.columns:
                    return default
                val = row[col]
                if pd.isna(val):
                    return default
                if is_int:
                    try:
                        return int(float(val))
                    except (ValueError, TypeError):
                        return 0
                return str(val)

            start_date_val = get_val("start_date", "")
            if start_date_val:
                try:
                    col_mapping = mapping.get("start_date")
                    raw_date = row[col_mapping]
                    if hasattr(raw_date, 'strftime'):
                        start_date_val = raw_date.strftime('%Y-%m-%d')
                    else:
                        start_date_val = pd.to_datetime(start_date_val).strftime('%Y-%m-%d')
                except Exception:
                    start_date_val = datetime.date.today().strftime('%Y-%m-%d')
            else:
                start_date_val = datetime.date.today().strftime('%Y-%m-%d')

            rec = TestRecord(
                project_name=sheet_name,
                category=get_val("category", "General"),
                test_method=get_val("test_method", "Testing"),
                test_number=get_val("test_number", "TM-000"),
                start_date=start_date_val,
                proto_weeks=get_val("proto_weeks", 0, is_int=True),
                proto_days=get_val("proto_days", 0, is_int=True),
                proto_qty=get_val("proto_qty", 0, is_int=True),
                dvt_weeks=get_val("dvt_weeks", 0, is_int=True),
                dvt_days=get_val("dvt_days", 0, is_int=True),
                dvt_qty=get_val("dvt_qty", 0, is_int=True),
                evt_weeks=get_val("evt_weeks", 0, is_int=True),
                evt_days=get_val("evt_days", 0, is_int=True),
                evt_qty=get_val("evt_qty", 0, is_int=True),
                pvt_weeks=get_val("pvt_weeks", 0, is_int=True),
                pvt_days=get_val("pvt_days", 0, is_int=True),
                pvt_qty=get_val("pvt_qty", 0, is_int=True),
                defect_qty=get_val("defect_qty", 0, is_int=True),
                comments=get_val("comments", "")
            )
            db.session.add(rec)
            imported_count += 1

        total_records_imported += imported_count
        synced_projects.append(sheet_name)

    # Save global sync config
    import json as _json
    mapping_record = ExcelMapping.query.filter_by(project_name='GLOBAL_SYNC').first()
    if not mapping_record:
        mapping_record = ExcelMapping(project_name='GLOBAL_SYNC')
        db.session.add(mapping_record)
    mapping_record.file_path = file_path
    mapping_record.sheet_name = 'ALL'
    mapping_record.mapping_json = _json.dumps({
        "selected_sheets": selected_sheets,
        "removed_columns": removed_cols_map
    })
    db.session.commit()

    return synced_projects, total_records_imported


@app.route('/api/local/load-transformed', methods=['POST'])
@login_required
def api_local_load_transformed():
    """Imports selected sheets from a locally uploaded Excel file into the Project Table."""
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files['file']
    if uploaded_file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    filename = uploaded_file.filename
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({"error": "Only Excel files (.xlsx, .xls) are supported"}), 400

    import json as _json
    import io
    config_raw = request.form.get('config', '{}')
    try:
        config = _json.loads(config_raw)
    except Exception:
        config = {}

    selected_sheets = config.get('selected_sheets', [])
    removed_cols_map = config.get('removed_columns', {})

    try:
        file_bytes = uploaded_file.read()
        xl = pd.ExcelFile(io.BytesIO(file_bytes))

        if not selected_sheets:
            # Default to all non-master sheets
            selected_sheets = [s for s in xl.sheet_names if s.lower().strip() != 'master data']

        synced_projects, total_records = _import_sheets_to_db(
            xl, selected_sheets, removed_cols_map, f"local://{filename}"
        )

        return jsonify({
            "success": True,
            "message": f"Successfully imported {len(synced_projects)} project(s) with {total_records} records from {filename}.",
            "projects": synced_projects,
            "count": total_records
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Failed to import local file: {str(e)}"}), 500


# --- POWER BI NAVIGATOR & QUERY EDITOR VIEWS ---

@app.route('/api/sharepoint/preview-all', methods=['POST'])
@login_required
def api_sharepoint_preview_all():
    """Downloads the selected Excel file and returns the columns and first 100 rows of all sheets."""
    data = request.json
    if not data or 'file_path' not in data:
        return jsonify({"error": "No file_path provided"}), 400
        
    file_path = data['file_path']
    try:
        file_content = SharePointService.download_file(file_path)
        
        import io
        xl = pd.ExcelFile(io.BytesIO(file_content))
        sheets = xl.sheet_names
        
        result = {}
        for sheet in sheets:
            # Skip "Master Data" sheet (case-insensitive)
            if sheet.lower().strip() == 'master data':
                continue
                
            df = xl.parse(sheet, nrows=100)
            df = clean_dataframe_headers(df)
            
            # Convert all column names to string
            columns = [str(col) for col in df.columns.tolist()]
            
            # Convert rows to list of lists (handling NaN)
            rows = []
            for _, row in df.iterrows():
                row_vals = []
                for val in row.tolist():
                    if pd.isna(val):
                        row_vals.append("")
                    else:
                        if isinstance(val, (datetime.date, datetime.datetime)):
                            row_vals.append(val.strftime('%Y-%m-%d'))
                        else:
                            row_vals.append(val)
                rows.append(row_vals)
                
            result[sheet] = {
                "columns": columns,
                "rows": rows
            }
            
        return jsonify({
            "success": True,
            "file_path": file_path,
            "sheets": result
        })
    except Exception as e:
        return jsonify({"error": f"Failed to preview Excel file: {str(e)}"}), 500

@app.route('/api/sharepoint/load-transformed', methods=['POST'])
@login_required
def api_sharepoint_load_transformed():
    """Imports selected sheets from Excel, applying column removals and automatic fuzzy mapping."""
    data = request.json
    if not data or 'file_path' not in data or 'selected_sheets' not in data:
        return jsonify({"error": "Missing required configuration"}), 400
        
    file_path = data['file_path']
    selected_sheets = data['selected_sheets']
    removed_cols_map = data.get('removed_columns', {}) # dict of sheet_name -> list of removed cols
    
    try:
        file_content = SharePointService.download_file(file_path)
        
        import io
        xl = pd.ExcelFile(io.BytesIO(file_content))
        
        synced_projects = []
        total_records_imported = 0
        
        for sheet_name in selected_sheets:
            if sheet_name not in xl.sheet_names:
                continue
                
            # Parse sheet
            df = xl.parse(sheet_name)
            df = clean_dataframe_headers(df)
            
            # Filter out removed columns
            removed_cols = removed_cols_map.get(sheet_name, [])
            active_cols = [col for col in df.columns if str(col) not in removed_cols]
            
            # Perform automatic fuzzy mapping on remaining columns
            mapping = {}
            fuzzy_rules = {
                "category": ['category', 'test category', 'group', 'device', 'product'],
                "test_method": ['method', 'test method', 'method name', 'name'],
                "test_number": ['number', 'test number', 'test #', 'ref', 'ref number', 'code'],
                "start_date": ['start', 'start date', 'date', 'timeline start'],
                "defect_qty": ['defect', 'defect qty', 'defective', 'defects', 'defect quantity', 'defective units'],
                "comments": ['comments', 'rejections', 'notes', 'rejection comment/s', 'comment'],
                
                "proto_weeks": ['proto wk', 'proto week', 'proto_wk', 'proto_weeks', 'proto weeks', 'week1'],
                "proto_days": ['proto day', 'proto_day', 'proto_days', 'proto days', 'day1'],
                "proto_qty": ['proto qty', 'proto_qty', 'proto_quantity', 'proto quantity', 'proto_qty', 'proto1'],
                
                "dvt_weeks": ['dvt wk', 'dvt week', 'dvt_wk', 'dvt_weeks', 'dvt weeks', 'week2'],
                "dvt_days": ['dvt day', 'dvt_day', 'dvt_days', 'dvt days', 'day3'],
                "dvt_qty": ['dvt qty', 'dvt_qty', 'dvt_quantity', 'dvt quantity', 'dvt_qty', 'dvt1'],
                
                "evt_weeks": ['evt wk', 'evt week', 'evt_wk', 'evt_weeks', 'evt weeks', 'week4'],
                "evt_days": ['evt day', 'evt_day', 'evt_days', 'evt days', 'day5'],
                "evt_qty": ['evt qty', 'evt_qty', 'evt_quantity', 'evt quantity', 'evt_qty', 'evt1'],
                
                "pvt_weeks": ['pvt wk', 'pvt week', 'pvt_wk', 'pvt_weeks', 'pvt weeks', 'week6'],
                "pvt_days": ['pvt day', 'pvt_day', 'pvt_days', 'pvt days', 'day7'],
                "pvt_qty": ['pvt qty', 'pvt_qty', 'pvt_quantity', 'pvt quantity', 'pvt_qty', 'pvt1'],
            }
            
            for field, rules in fuzzy_rules.items():
                for col in active_cols:
                    col_lower = str(col).lower().strip()
                    if any(r == col_lower or r in col_lower for r in rules):
                        mapping[field] = str(col)
                        break
                        
            # Create or find Project
            proj = Project.query.filter_by(name=sheet_name).first()
            if not proj:
                proj = Project(
                    name=sheet_name,
                    description=f"Imported from SharePoint sheet {sheet_name}",
                    status="Active",
                    created_at=datetime.date.today().strftime('%m/%d/%Y')
                )
                db.session.add(proj)
                db.session.commit()
                
            # Clear existing test records for this project
            TestRecord.query.filter_by(project_name=sheet_name).delete()
            
            imported_count = 0
            for index, row in df.iterrows():
                def get_val(field, default=None, is_int=False):
                    col = mapping.get(field)
                    if not col or col not in df.columns:
                         return default
                    val = row[col]
                    if pd.isna(val):
                        return default
                    if is_int:
                        try:
                            return int(float(val))
                        except ValueError:
                            return 0
                    return str(val)
                    
                # Parse start date
                start_date_val = get_val("start_date", "")
                if start_date_val:
                    try:
                        col_mapping = mapping.get("start_date")
                        raw_date = row[col_mapping]
                        if hasattr(raw_date, 'strftime'):
                            start_date_val = raw_date.strftime('%Y-%m-%d')
                        else:
                            start_date_val = pd.to_datetime(start_date_val).strftime('%Y-%m-%d')
                    except Exception:
                        start_date_val = datetime.date.today().strftime('%Y-%m-%d')
                else:
                    start_date_val = datetime.date.today().strftime('%Y-%m-%d')
    
                rec = TestRecord(
                    project_name=sheet_name,
                    category=get_val("category", "General"),
                    test_method=get_val("test_method", "Testing"),
                    test_number=get_val("test_number", "TM-000"),
                    start_date=start_date_val,
                    
                    proto_weeks=get_val("proto_weeks", 0, is_int=True),
                    proto_days=get_val("proto_days", 0, is_int=True),
                    proto_qty=get_val("proto_qty", 0, is_int=True),
                    
                    dvt_weeks=get_val("dvt_weeks", 0, is_int=True),
                    dvt_days=get_val("dvt_days", 0, is_int=True),
                    dvt_qty=get_val("dvt_qty", 0, is_int=True),
                    
                    evt_weeks=get_val("evt_weeks", 0, is_int=True),
                    evt_days=get_val("evt_days", 0, is_int=True),
                    evt_qty=get_val("evt_qty", 0, is_int=True),
                    
                    pvt_weeks=get_val("pvt_weeks", 0, is_int=True),
                    pvt_days=get_val("pvt_days", 0, is_int=True),
                    pvt_qty=get_val("pvt_qty", 0, is_int=True),
                    
                    defect_qty=get_val("defect_qty", 0, is_int=True),
                    comments=get_val("comments", "")
                )
                db.session.add(rec)
                imported_count += 1
                
            total_records_imported += imported_count
            synced_projects.append(sheet_name)
            
        # Save global sync configuration
        import json
        mapping_record = ExcelMapping.query.filter_by(project_name='GLOBAL_SYNC').first()
        if not mapping_record:
            mapping_record = ExcelMapping(project_name='GLOBAL_SYNC')
            db.session.add(mapping_record)
        mapping_record.file_path = file_path
        mapping_record.sheet_name = 'ALL'
        mapping_record.mapping_json = json.dumps({
            "selected_sheets": selected_sheets,
            "removed_columns": removed_cols_map
        })
        
        db.session.commit()
        return jsonify({
            "success": True,
            "message": f"Successfully loaded and transformed {len(synced_projects)} projects with {total_records_imported} records.",
            "projects": synced_projects,
            "count": total_records_imported
        })
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": f"Failed to load transformed data: {str(e)}"}), 500


# --- AI CHAT ENDPOINT ---
@app.route('/api/chat', methods=['POST'])
@login_required
def api_chat():
    """Intelligent AI assistant for operations data analysis and report generation."""
    data = request.json
    if not data or 'query' not in data:
        return jsonify({"response": "Please send a query."}), 400

    query = data.get('query', '').strip().lower()
    context_project = data.get('project', None)

    # Gather all projects for context
    all_projects = Project.query.filter_by(status='Active').all()
    all_project_names = [p.name for p in all_projects]

    # --- Intent: List all projects ---
    if any(kw in query for kw in ['list projects', 'all projects', 'show projects', 'what projects']):
        names = ', '.join(all_project_names) if all_project_names else 'No active projects'
        return jsonify({"response": f"📋 There are {len(all_project_names)} active projects:\n{names}"})

    # --- Intent: Count / summary of a specific project ---
    target_project = None
    for pname in all_project_names:
        if pname.lower() in query:
            target_project = pname
            break
    if not target_project and context_project:
        target_project = context_project

    # --- Intent: Generate / create report ---
    report_keywords = ['generate report', 'create report', 'make report', 'export report', 'generate pdf', 'create pdf', 'export pdf', 'download report']
    if any(kw in query for kw in report_keywords):
        proj = target_project or context_project or (all_project_names[0] if all_project_names else None)
        if proj:
            return jsonify({
                "response": f"📄 Generating PDF report for **Project {proj}**... The download will start shortly!",
                "generate_report": True,
                "report_project": proj
            })
        else:
            return jsonify({"response": "Please specify a project name for the report, e.g. 'Generate report for 893'."})

    # --- Intent: Summarize a project ---
    summary_keywords = ['summary', 'summarize', 'overview', 'status', 'how many', 'records', 'rows', 'test methods', 'details']
    if any(kw in query for kw in summary_keywords) and target_project:
        records = TestRecord.query.filter_by(project_name=target_project).all()
        total_qty = sum(
            (r.proto_qty or 0) + (r.dvt_qty or 0) + (r.evt_qty or 0) + (r.pvt_qty or 0)
            for r in records
        )
        rejections = [r for r in records if r.comments and r.comments.strip()]
        proj_obj = Project.query.filter_by(name=target_project).first()
        desc = proj_obj.description if proj_obj and proj_obj.description else 'No description'

        response = (
            f"📊 **Project {target_project} Summary**\n"
            f"Description: {desc}\n"
            f"Test Records: {len(records)}\n"
            f"Total Qty across all phases: {total_qty}\n"
            f"Records with rejections: {len(rejections)}\n"
        )
        if rejections:
            response += "\nRejection notes:\n"
            for r in rejections[:3]:  # Show max 3
                response += f"  • [{r.category}] {r.test_method}: {r.comments}\n"
            if len(rejections) > 3:
                response += f"  ... and {len(rejections) - 3} more."
        return jsonify({"response": response})

    # --- Intent: Summarize all projects ---
    if any(kw in query for kw in ['summarize all', 'overview all', 'all projects status', 'total projects']):
        total_records = TestRecord.query.count()
        total_rejections = TestRecord.query.filter(TestRecord.comments != None, TestRecord.comments != '').count()
        response = (
            f"📊 **Operations Overview**\n"
            f"Active Projects: {len(all_project_names)}\n"
            f"Total Test Records: {total_records}\n"
            f"Records with Rejections: {total_rejections}\n\n"
            f"Projects: {', '.join(all_project_names[:10])}" +
            (f" ... and {len(all_project_names) - 10} more" if len(all_project_names) > 10 else "")
        )
        return jsonify({"response": response})

    # --- Intent: Rejection analysis ---
    if any(kw in query for kw in ['rejection', 'rejected', 'fail', 'not working', 'issues', 'problems']):
        if target_project:
            records = TestRecord.query.filter_by(project_name=target_project).filter(
                TestRecord.comments != None, TestRecord.comments != ''
            ).all()
            if records:
                resp = f"⚠️ **Rejections in Project {target_project}** ({len(records)} found):\n"
                for r in records:
                    resp += f"• [{r.category}] {r.test_method} — {r.comments}\n"
                return jsonify({"response": resp})
            else:
                return jsonify({"response": f"✅ No rejections recorded for Project {target_project}."})
        else:
            total_rej = TestRecord.query.filter(
                TestRecord.comments != None, TestRecord.comments != ''
            ).count()
            return jsonify({"response": f"⚠️ There are **{total_rej}** test records with rejection notes across all projects. Specify a project name for details."})

    # --- Intent: SharePoint sync help ---
    if any(kw in query for kw in ['sync', 'sharepoint', 'refresh data', 'update data']):
        return jsonify({"response": "🔄 To sync data from SharePoint, click **SharePoint Sync** in the left sidebar. This will pull in any new projects and update existing ones."})

    # --- Intent: PDF/export help ---
    if any(kw in query for kw in ['pdf', 'export', 'download', 'report']):
        proj = target_project or context_project or 'current project'
        return jsonify({
            "response": f"📄 I can generate a PDF report for you! Just say:\n'Generate report for {proj}'\nOr use the **Export PDF Report** button in the left sidebar."
        })

    # --- Default / help response ---
    help_text = (
        "🤖 **AI Assistant** — I can help you with:\n"
        "• **'List all projects'** — see all active projects\n"
        "• **'Summary of project 893'** — get stats for a specific project\n"
        "• **'Generate report for 990'** — create and download a PDF\n"
        "• **'Rejections in project 893'** — view rejection analysis\n"
        "• **'Summarize all projects'** — full operations overview\n"
        f"\nCurrently viewing: **Project {context_project or 'N/A'}**"
    )
    return jsonify({"response": help_text})

# --- GANTT CHART GENERATION ---
def generate_gantt_chart_image(records, filepath, project_name=None):
    """Generates a multi-row project Gantt timeline image using Matplotlib and saves it."""
    def parse_date(value):
        if not value:
            return None
        if isinstance(value, datetime.datetime):
            return value
        if isinstance(value, datetime.date):
            return datetime.datetime.combine(value, datetime.time.min)
        try:
            return datetime.datetime.strptime(value, "%Y-%m-%d")
        except Exception:
            try:
                return datetime.datetime.fromisoformat(value)
            except Exception:
                return None

    def wrap_label(text, width=44):
        if not text:
            return ""
        words = str(text).split()
        lines = []
        current = []
        for word in words:
            if len(" ".join(current + [word])) <= width:
                current.append(word)
            else:
                lines.append(" ".join(current))
                current = [word]
        if current:
            lines.append(" ".join(current))
        return "\n".join(lines)

    if project_name:
        records = [r for r in records if r.get("Project Name") == project_name]

    if not records:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, f"No active test methods for Project {project_name}" if project_name else "No active test methods to display",
                horizontalalignment='center', verticalalignment='center', fontsize=12)
        ax.set_axis_off()
        plt.savefig(filepath, dpi=300, bbox_inches='tight')
        plt.close()
        return

    phase_colors = {
        "Proto": "#8E44AD",
        "DVT": "#2980B9",
        "EVT": "#27AE60",
        "PVT": "#D35400"
    }

    rows = []
    min_date = None
    max_date = None

    for record in records:
        start_date = parse_date(record.get("Start Date"))
        if not start_date:
            continue

        if min_date is None or start_date < min_date:
            min_date = start_date
        if max_date is None or start_date > max_date:
            max_date = start_date

        phases = [
            ("Proto", int(record.get("Proto Weeks", 0) or 0), int(record.get("Proto Days", 0) or 0)),
            ("DVT", int(record.get("DVT Weeks", 0) or 0), int(record.get("DVT Days", 0) or 0)),
            ("EVT", int(record.get("EVT Weeks", 0) or 0), int(record.get("EVT Days", 0) or 0)),
            ("PVT", int(record.get("PVT Weeks", 0) or 0), int(record.get("PVT Days", 0) or 0))
        ]

        label = f"{record.get('Test Number', 'Unknown')} — {record.get('Test Method', 'Untitled')}"
        label = wrap_label(label, width=44)

        bars = []
        phase_start = start_date
        for phase_name, weeks, days in phases:
            duration = weeks * 7 + days
            if duration <= 0:
                continue
            phase_end = phase_start + datetime.timedelta(days=duration)
            bar = {
                "phase": phase_name,
                "start": phase_start,
                "duration": duration,
                "color": phase_colors.get(phase_name, "#999999")
            }
            bars.append(bar)
            if min_date is None or bar["start"] < min_date:
                min_date = bar["start"]
            if max_date is None or phase_end > max_date:
                max_date = phase_end
            phase_start = phase_end

        if bars:
            rows.append({"label": label, "bars": bars})

    if not rows:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, "No valid timeline data to display.",
                horizontalalignment='center', verticalalignment='center', fontsize=12)
        ax.set_axis_off()
        plt.savefig(filepath, dpi=300, bbox_inches='tight')
        plt.close()
        return

    # Visual style tuned to the provided design: dark background, rounded bars, weekly grid
    fig_height = max(4, len(rows) * 0.6 + 2)
    bg_color = '#06293a'  # deep navy
    fg_text = '#eaf6ff'
    grid_color = '#10475a'

    # Use a two-column layout: left sidebar for labels, right for the gantt
    from matplotlib.patches import FancyBboxPatch, Patch, Rectangle
    fig = plt.figure(figsize=(12, fig_height), facecolor=bg_color)
    gs = fig.add_gridspec(1, 2, width_ratios=[0.28, 0.72], wspace=0.02)
    ax_left = fig.add_subplot(gs[0, 0])
    ax = fig.add_subplot(gs[0, 1])

    ax.set_facecolor(bg_color)
    ax_left.set_facecolor('#053242')

    # Configure left axis (labels sidebar)
    ax_left.set_ylim(-0.5, len(rows) - 0.5)
    ax_left.invert_yaxis()
    ax_left.xaxis.set_visible(False)
    ax_left.yaxis.set_visible(False)

    # Draw category groups on the left sidebar
    categories = [r.get('category', '') for r in rows]
    group_positions = {}
    for idx, cat in enumerate(categories):
        group_positions.setdefault(cat, []).append(idx)

    for cat, indices in group_positions.items():
        mid = (indices[0] + indices[-1]) / 2
        ax_left.text(0.05, mid, str(cat), va='center', ha='left', fontsize=11, color=fg_text, weight='bold')

    # Also print each row's short method label under the category area
    for y, row in enumerate(rows):
        short = row['label'].split('\n')[0]
        ax_left.text(0.05, y + 0.28, short, va='center', ha='left', fontsize=8, color='#cdeefb')

    # Right axis: Gantt chart area
    start_num = mdates.date2num(min_date - datetime.timedelta(days=1))
    end_num = mdates.date2num(max_date + datetime.timedelta(days=3))
    ax.set_xlim(start_num, end_num)
    ax.set_ylim(-0.5, len(rows) - 0.5)
    ax.invert_yaxis()

    # Weekly vertical lines and subtle grid
    cur_date = min_date
    while cur_date <= max_date:
        x = mdates.date2num(cur_date)
        ax.axvline(x, color=grid_color, linewidth=0.7, alpha=0.6)
        cur_date += datetime.timedelta(days=7)

    # Month labels above the gantt area
    month_iter = datetime.date(min_date.year, min_date.month, 1)
    top_y = -0.8
    while month_iter <= max_date.date():
        month_start = datetime.datetime(month_iter.year, month_iter.month, 1)
        next_month = (month_start.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
        month_mid = month_start + (next_month - month_start) / 2
        ax.text(mdates.date2num(month_mid), top_y, month_start.strftime('%B'), ha='center', va='bottom', color=fg_text, fontsize=11, fontweight='bold')
        month_iter = next_month.date()

    # Draw rounded bars
    for y, row in enumerate(rows):
        for bar in row['bars']:
            x = mdates.date2num(bar['start'])
            width = bar['duration']
            height = 0.6
            y_pos = y - height / 2
            box = FancyBboxPatch((x, y_pos), width, height,
                                 boxstyle='round,pad=0.02,rounding_size=6',
                                 linewidth=0, facecolor=bar['color'], alpha=0.98)
            ax.add_patch(box)
            # small outline to make bars pop
            edge = Rectangle((x, y_pos), width, height, linewidth=0.6, edgecolor='#083a4a', facecolor='none')
            ax.add_patch(edge)
            # label inside if enough space
            try:
                label_text = row['label'].split('\n')[0]
                if width >= 6:
                    ax.text(x + 0.4, y, label_text, va='center', ha='left', color='white', fontsize=8, fontweight='bold')
            except Exception:
                pass

    # X-axis formatting
    ax.xaxis_date()
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %d'))
    plt.setp(ax.get_xticklabels(), rotation=30, ha='right', color=fg_text, fontsize=8)

    # Clean spines
    for spine in ['top', 'right', 'left', 'bottom']:
        ax.spines[spine].set_visible(False)

    # Legend and title
    legend_elements = [Patch(facecolor=color, edgecolor='none', label=phase) for phase, color in phase_colors.items()]
    ax.legend(handles=legend_elements, loc='upper right', frameon=False, fontsize=9)

    today = datetime.datetime.now()
    if min_date <= today <= max_date:
        ax.axvline(mdates.date2num(today), color='#ffcc00', linewidth=1.2, linestyle='--', alpha=0.9)

    ax.set_title(f"Project Gantt Timeline — Project {project_name}" if project_name else "Project Gantt Timeline", fontsize=14, fontweight='bold', pad=14, color=fg_text)

    plt.tight_layout()
    plt.savefig(filepath, dpi=300, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)

# --- PDF REPORT GENERATION ---
def generate_pdf_report(filename, project_name=None, comment=None):
    """Generates a styled Test Operations PDF report containing a Gantt chart and rejected products."""
    ExcelDataStore.initialize_db()
    
    # Fetch records filtered by project
    if project_name:
        records = TestRecord.query.filter_by(project_name=project_name).all()
        records_dict = [r.to_dict() for r in records]
    else:
        records_dict = ExcelDataStore.get_tasks()
        
    # 1. Generate the Gantt Chart Image
    gantt_image_path = os.path.join(UPLOAD_FOLDER, "gantt_chart_report.png")
    generate_gantt_chart_image(records_dict, gantt_image_path, project_name)
    
    doc = SimpleDocTemplate(filename, pagesize=letter,
                            rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    styles = getSampleStyleSheet()
    
    # Custom color palette
    charcoal = colors.HexColor("#121212")
    steel_gray = colors.HexColor("#555555")
    prussian_blue = colors.HexColor("#00539C")
    light_bg = colors.HexColor("#F9F9FB")
    border_color = colors.HexColor("#E5E5E7")
    danger_red = colors.HexColor("#C0392B")
    
    title_style = ParagraphStyle(
        'OpsTitle',
        parent=styles['Heading1'],
        fontSize=22,
        leading=26,
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
        fontSize=13,
        leading=17,
        textColor=prussian_blue,
        fontName='Helvetica-Bold',
        spaceBefore=15,
        spaceAfter=10,
        keepWithNext=True
    )
    
    cell_header = ParagraphStyle(
        'OpsHeader',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.white,
        fontName='Helvetica-Bold'
    )
    
    cell_style = ParagraphStyle(
        'OpsCell',
        parent=styles['Normal'],
        fontSize=8,
        leading=11,
        textColor=charcoal,
        fontName='Helvetica'
    )
    
    cell_bold = ParagraphStyle(
        'OpsCellBold',
        parent=styles['Normal'],
        fontSize=8,
        leading=11,
        textColor=danger_red,
        fontName='Helvetica-Bold'
    )
    
    # Header Title
    doc_title = f"Project Gantt Timeline - Project {project_name}" if project_name else "Daily Test Operations Report"
    story.append(Paragraph(doc_title, title_style))
    story.append(Paragraph(f"Generated on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Confidential Operations Data", subtitle_style))

    summary_style = ParagraphStyle(
        'OpsSummary',
        parent=styles['Normal'],
        fontSize=9,
        leading=12,
        textColor=steel_gray,
        spaceAfter=6
    )

    total_records = len(records_dict)
    total_defects = sum(int(r.get('Defect Qty', 0) or 0) for r in records_dict)
    total_comments = sum(1 for r in records_dict if str(r.get('Comments', '')).strip())

    summary_texts = [
        f"Project: {project_name}" if project_name else "Project: All Active Projects",
        f"Total task records: {total_records}",
        f"Total defects logged: {total_defects}",
        f"Records with comments or rejections: {total_comments}"
    ]
    for text in summary_texts:
        story.append(Paragraph(text, summary_style))
    story.append(Spacer(1, 12))

    if comment and str(comment).strip():
        story.append(Paragraph("Report Comment", section_title_style))
        story.append(Paragraph(str(comment).strip().replace('\n', '<br/>'), summary_style))
        story.append(Spacer(1, 12))

    story.append(Paragraph("Timeline Summary", section_title_style))
    story.append(Paragraph("The Gantt chart below represents scheduled phase progression for each test method, based on actual start dates and phase durations. Entries are ordered by earliest start date to provide a clean timeline overview.", summary_style))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Test Method Milestones (Gantt Chart)", section_title_style))
    story.append(Image(gantt_image_path, width=500, height=270))
    story.append(Spacer(1, 20))
    story.append(Paragraph("Test Records Summary", section_title_style))
    all_table_data = [[
        Paragraph("#", cell_header),
        Paragraph("Category", cell_header),
        Paragraph("Test Method", cell_header),
        Paragraph("Test #", cell_header),
        Paragraph("Proto\n(W/D/Qty)", cell_header),
        Paragraph("DVT\n(W/D/Qty)", cell_header),
        Paragraph("EVT\n(W/D/Qty)", cell_header),
        Paragraph("PVT\n(W/D/Qty)", cell_header),
        Paragraph("Defect\nQty", cell_header),
        Paragraph("Rejection Comment/s", cell_header)
    ]]
    for idx, r in enumerate(records_dict, 1):
        defect_qty = r.get("Defect Qty", 0)
        defect_style = cell_bold if defect_qty and int(defect_qty) > 0 else cell_style
        all_table_data.append([
            Paragraph(str(idx), cell_style),
            Paragraph(r.get("Category", ""), cell_style),
            Paragraph(r.get("Test Method", ""), cell_style),
            Paragraph(r.get("Test Number", ""), cell_style),
            Paragraph(f"{r.get('Proto Weeks',0)}w {r.get('Proto Days',0)}d / {r.get('Proto Qty',0)}", cell_style),
            Paragraph(f"{r.get('DVT Weeks',0)}w {r.get('DVT Days',0)}d / {r.get('DVT Qty',0)}", cell_style),
            Paragraph(f"{r.get('EVT Weeks',0)}w {r.get('EVT Days',0)}d / {r.get('EVT Qty',0)}", cell_style),
            Paragraph(f"{r.get('PVT Weeks',0)}w {r.get('PVT Days',0)}d / {r.get('PVT Qty',0)}", cell_style),
            Paragraph(str(defect_qty), defect_style),
            Paragraph(r.get("Comments", "-") or "-", cell_bold if r.get("Comments") else cell_style)
        ])
    if not records_dict:
        all_table_data.append([Paragraph("No records found.", cell_style)] + [Paragraph("-", cell_style)] * 9)

    all_table = Table(all_table_data, colWidths=[18, 80, 90, 55, 62, 62, 62, 62, 35, 100])
    all_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), charcoal),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('ALIGN', (8,0), (8,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.4, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, light_bg])
    ]))
    story.append(all_table)
    story.append(Spacer(1, 14))

    # Rejections-only sub-table
    rejected_records = [r for r in records_dict if str(r.get("Comments", "")).strip() != "" or int(r.get("Defect Qty", 0)) > 0]
    story.append(Paragraph("Rejected Products Inventory", section_title_style))
    table_data = [[
        Paragraph("Category", cell_header),
        Paragraph("Test Method", cell_header),
        Paragraph("Test Number", cell_header),
        Paragraph("Defect Qty", cell_header),
        Paragraph("Rejection Comment/s", cell_header)
    ]]
    if rejected_records:
        for r in rejected_records:
            table_data.append([
                Paragraph(r.get("Category", ""), cell_style),
                Paragraph(r.get("Test Method", ""), cell_style),
                Paragraph(r.get("Test Number", ""), cell_style),
                Paragraph(str(r.get("Defect Qty", 0)), cell_bold),
                Paragraph(r.get("Comments", "") or "-", cell_bold)
            ])
    else:
        table_data.append([
            Paragraph("No rejected products recorded for this project.", cell_style),
            Paragraph("-", cell_style), Paragraph("-", cell_style),
            Paragraph("-", cell_style), Paragraph("-", cell_style)
        ])

    tasks_table = Table(table_data, colWidths=[110, 110, 70, 55, 185])
    tasks_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), charcoal),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, light_bg])
    ]))
    story.append(tasks_table)
    
    # Footer Note
    story.append(Spacer(1, 30))
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
    project_name = request.args.get('project')
    comment = request.args.get('comment')
    pdf_filename = os.path.join(UPLOAD_FOLDER, "Daily_Operations_Report.pdf")
    try:
        generate_pdf_report(pdf_filename, project_name, comment)
        download_name = f"Gantt_Report_{project_name}.pdf" if project_name else "Operations_Report.pdf"
        return send_file(pdf_filename, as_attachment=True, download_name=download_name, mimetype='application/pdf')
    except Exception as e:
        return jsonify({"error": f"Failed to generate PDF: {str(e)}"}), 500

# --- EMAIL AUTOMATION ---

@app.route('/send-email', methods=['POST'])
@login_required
def send_email():
    recipient = request.form.get('recipient') or os.getenv('MAIL_RECIPIENT', 'operations-manager@ops.com')
    subject = request.form.get('subject') or f"Daily Test Operations Report - {datetime.date.today().strftime('%Y-%m-%d')}"
    body = request.form.get('body') or """Dear Team,

Please find attached the Daily Test Operations Project Status Report containing the latest project tasks, timeline progression, and completion rates.

This email was auto-generated by the Operations Dashboard."""

    project_name = request.form.get('project')

    # Step 1: Always generate the PDF
    pdf_filename = os.path.join(UPLOAD_FOLDER, "email_report.pdf")
    try:
        generate_pdf_report(pdf_filename, project_name)
    except Exception as e:
        return jsonify({"error": f"Failed to generate PDF for email: {str(e)}"}), 500

    # Step 2: Attempt real SMTP send if credentials are configured
    mail_server = os.getenv('MAIL_SERVER')
    mail_username = os.getenv('MAIL_USERNAME')
    mail_password = os.getenv('MAIL_PASSWORD')
    mail_sender = os.getenv('MAIL_DEFAULT_SENDER') or os.getenv('MAIL_SENDER') or mail_username
    mail_port = int(os.getenv('MAIL_PORT', 587))
    mail_tls = os.getenv('MAIL_USE_TLS', 'True').lower() in ('true', '1', 'yes')
    mail_ssl = os.getenv('MAIL_USE_SSL', 'False').lower() in ('true', '1', 'yes')
    mail_timeout = int(os.getenv('MAIL_TIMEOUT', 15))

    if mail_server:
        try:
            import smtplib
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            from email.mime.application import MIMEApplication

            if not mail_sender:
                raise ValueError('MAIL_DEFAULT_SENDER or MAIL_USERNAME must be configured to send email.')

            msg = MIMEMultipart()
            msg['From'] = mail_sender
            msg['To'] = recipient
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            with open(pdf_filename, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype='pdf')
                pdf_download_name = f"Gantt_Report_{project_name}.pdf" if project_name else "Operations_Report.pdf"
                attachment.add_header('Content-Disposition', 'attachment', filename=pdf_download_name)
                msg.attach(attachment)

            if mail_ssl:
                smtp_class = smtplib.SMTP_SSL
            else:
                smtp_class = smtplib.SMTP

            with smtp_class(mail_server, mail_port, timeout=mail_timeout) as smtp:
                if not mail_ssl and mail_tls:
                    smtp.starttls()
                if mail_username and mail_password:
                    smtp.login(mail_username, mail_password)
                smtp.sendmail(mail_sender, recipient, msg.as_string())

            return jsonify({
                "success": True,
                "mode": "smtp",
                "message": f"Email with PDF report successfully sent to {recipient}!"
            })
        except Exception as e:
            return jsonify({
                "success": False,
                "mode": "smtp",
                "message": f"Failed to send email via SMTP: {str(e)}",
                "error": str(e)
            }), 500

    # Step 3: Simulation mode (no SMTP configured)
    return jsonify({
        "success": True,
        "mode": "simulation",
        "message": f"Email simulated (no SMTP configured). PDF report generated and would be sent to {recipient}.",
        "pdf_generated": True,
        "recipient": recipient,
        "subject": subject
    })

    # Step 3: Simulation mode (no SMTP configured)
    return jsonify({
        "success": True,
        "mode": "simulation",
        "message": f"Email simulated (no SMTP configured). PDF report generated and would be sent to {recipient}.",
        "pdf_generated": True,
        "recipient": recipient,
        "subject": subject
    })


@app.errorhandler(Exception)
def handle_exception(e):
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e
    import traceback
    tb = traceback.format_exc()
    log_path = os.path.join(app.config['UPLOAD_FOLDER'], 'error_log.txt')
    try:
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(f"\n--- ERROR AT {datetime.datetime.now()} ---\n")
            f.write(tb)
            f.write("\n------------------------------------\n")
    except Exception:
        pass
    return jsonify({"error": str(e), "traceback": tb}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True)

