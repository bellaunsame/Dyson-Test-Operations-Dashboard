import os
import datetime
import sys
import json
import uuid
import threading
from functools import wraps
from dotenv import load_dotenv
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
import pandas as pd
from PIL import Image as PILImage
PILImage.MAX_IMAGE_PIXELS = None
import matplotlib
matplotlib.use('Agg')  # Headless mode for web servers
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# ReportLab imports for PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

app = Flask(__name__)
load_dotenv()
app.secret_key = "super_secret_key_for_ops_operations"

# Configure App Data and Upload Folders
UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
EXCEL_PATH = os.path.join(UPLOAD_FOLDER, 'tasks.xlsx')
SYNC_CONFIG_PATH = os.path.join(UPLOAD_FOLDER, 'sync_config.json')

# Initialize Excel file by copying template if it doesn't exist
TEMPLATE_PATH = os.path.join(app.root_path, 'templates', 'SystemBoard.xlsx')
if not os.path.exists(EXCEL_PATH) and os.path.exists(TEMPLATE_PATH):
    import shutil
    shutil.copy(TEMPLATE_PATH, EXCEL_PATH)

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
        
        if not client or not site_url:
            return []
            
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
                return []
                
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
            return result
        except Exception as e:
            print(f"Error listing Graph files: {e}")
            return []

    @staticmethod
    def download_file(filename):
        client = SharePointService.get_graph_client()
        site_url = os.getenv('SHAREPOINT_SITE_URL')
        doc_lib = os.getenv('SHAREPOINT_DOC_LIB', 'Test Data')
        
        if not client or not site_url:
            # Fallback: check if we have a local copy in the uploads folder
            local_copy_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.exists(local_copy_path):
                with open(local_copy_path, 'rb') as f:
                    return f.read()
            raise FileNotFoundError(f"SharePoint client not configured and file '{filename}' not found locally.")
            
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
            # Fallback: check if we have a local copy in the uploads folder
            local_copy_path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.exists(local_copy_path):
                with open(local_copy_path, 'rb') as f:
                    return f.read()
            raise e


# --- EXCEL DATA STORE (Pure Excel Implementation) ---

FUZZY_RULES = {
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

class ExcelDataStore:
    @staticmethod
    def get_projects():
        """Returns list of active projects (sheets in the Excel file)."""
        if not os.path.exists(EXCEL_PATH):
            return []
        try:
            with pd.ExcelFile(EXCEL_PATH) as xl:
                projects = []
                for sheet in xl.sheet_names:
                    if sheet.lower().strip() == 'master data':
                        continue
                    try:
                        df = xl.parse(sheet)
                        row_count = len(df)
                    except Exception:
                        row_count = 0
                    projects.append({
                        "id": sheet,
                        "name": sheet,
                        "description": f"Excel Sheet: {sheet}",
                        "status": "Active",
                        "created_at": "N/A",
                        "row_count": row_count
                    })
                return projects
        except Exception as e:
            print(f"Error getting projects: {e}")
            return []

    @staticmethod
    def get_project_rows(project_name):
        """Returns all mapped test records for a specific project/sheet."""
        if not os.path.exists(EXCEL_PATH):
            return []
        try:
            with pd.ExcelFile(EXCEL_PATH) as xl:
                if project_name not in xl.sheet_names:
                    return []
                df = xl.parse(project_name)
            df = clean_dataframe_headers(df)
            
            # Map columns using fuzzy rules
            mapping = {}
            for field, rules in FUZZY_RULES.items():
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    if any(r == col_lower or r in col_lower for r in rules):
                        mapping[field] = col
                        break
            
            records = []
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
                
                records.append({
                    "id": index,  # Use row index as unique ID
                    "Project Name": project_name,
                    "Category": get_val("category", "General"),
                    "Test Method": get_val("test_method", "Testing"),
                    "Test Number": get_val("test_number", "TM-000"),
                    "Start Date": start_date_val,
                    "Proto Weeks": get_val("proto_weeks", 0, is_int=True),
                    "Proto Days": get_val("proto_days", 0, is_int=True),
                    "Proto Qty": get_val("proto_qty", 0, is_int=True),
                    "DVT Weeks": get_val("dvt_weeks", 0, is_int=True),
                    "DVT Days": get_val("dvt_days", 0, is_int=True),
                    "DVT Qty": get_val("dvt_qty", 0, is_int=True),
                    "EVT Weeks": get_val("evt_weeks", 0, is_int=True),
                    "EVT Days": get_val("evt_days", 0, is_int=True),
                    "EVT Qty": get_val("evt_qty", 0, is_int=True),
                    "PVT Weeks": get_val("pvt_weeks", 0, is_int=True),
                    "PVT Days": get_val("pvt_days", 0, is_int=True),
                    "PVT Qty": get_val("pvt_qty", 0, is_int=True),
                    "Defect Qty": get_val("defect_qty", 0, is_int=True),
                    "Comments": get_val("comments", "")
                })
            return records
        except Exception as e:
            print(f"Error getting project rows: {e}")
            return []

    @staticmethod
    def save_project_row(project_name, row_id, data):
        """Saves (inserts or updates) a row in the specific sheet."""
        if not os.path.exists(EXCEL_PATH):
            return False
        try:
            # Load all sheets
            with pd.ExcelFile(EXCEL_PATH) as xl:
                sheets_data = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
            
            # If sheet doesn't exist, create it
            if project_name not in sheets_data:
                sheets_data[project_name] = pd.DataFrame(columns=[
                    "Category", "Test Method", "Test Number", "Start Date",
                    "Proto Weeks", "Proto Days", "Proto Qty",
                    "DVT Weeks", "DVT Days", "DVT Qty",
                    "EVT Weeks", "EVT Days", "EVT Qty",
                    "PVT Weeks", "PVT Days", "PVT Qty",
                    "Defect Qty", "Comments"
                ])
                
            df = sheets_data[project_name]
            
            # Determine mapping to write back to correct columns
            mapping = {}
            for field, rules in FUZZY_RULES.items():
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    if any(r == col_lower or r in col_lower for r in rules):
                        mapping[field] = col
                        break
                # If column not found, use a default name
                if field not in mapping:
                    mapping[field] = field.replace('_', ' ').title()
            
            # Prepare row data
            row_dict = {}
            field_mappings = {
                "category": data.get("Category", "General"),
                "test_method": data.get("Test Method", "Testing"),
                "test_number": data.get("Test Number", "TM-000"),
                "start_date": data.get("Start Date", datetime.date.today().strftime('%Y-%m-%d')),
                "proto_weeks": int(data.get("Proto Weeks", 0)),
                "proto_days": int(data.get("Proto Days", 0)),
                "proto_qty": int(data.get("Proto Qty", 0)),
                "dvt_weeks": int(data.get("DVT Weeks", 0)),
                "dvt_days": int(data.get("DVT Days", 0)),
                "dvt_qty": int(data.get("DVT Qty", 0)),
                "evt_weeks": int(data.get("EVT Weeks", 0)),
                "evt_days": int(data.get("EVT Days", 0)),
                "evt_qty": int(data.get("EVT Qty", 0)),
                "pvt_weeks": int(data.get("PVT Weeks", 0)),
                "pvt_days": int(data.get("PVT Days", 0)),
                "pvt_qty": int(data.get("PVT Qty", 0)),
                "defect_qty": int(data.get("Defect Qty", 0)),
                "comments": data.get("Comments", "")
            }
            
            for field, val in field_mappings.items():
                col_name = mapping[field]
                row_dict[col_name] = val
                
            if row_id is None or row_id < 0 or row_id >= len(df):
                # Append new row
                new_row = pd.DataFrame([row_dict])
                for col in new_row.columns:
                    if col in df.columns and isinstance(row_dict[col], str):
                        df[col] = df[col].astype(object)
                df = pd.concat([df, new_row], ignore_index=True)
            else:
                # Update existing row
                for col_name, val in row_dict.items():
                    if col_name not in df.columns:
                        df[col_name] = None
                    if isinstance(val, str) and not pd.api.types.is_object_dtype(df[col_name]):
                        df[col_name] = df[col_name].astype(object)
                    df.at[row_id, col_name] = val
                    
            sheets_data[project_name] = df
            
            # Write all sheets back
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
                for sheet, sheet_df in sheets_data.items():
                    sheet_df.to_excel(writer, index=False, sheet_name=sheet)
                    
            return True
        except Exception as e:
            print(f"Error saving project row: {e}")
            return False

    @staticmethod
    def delete_project_row(project_name, row_id):
        """Deletes a row from the specific sheet."""
        if not os.path.exists(EXCEL_PATH):
            return False
        try:
            with pd.ExcelFile(EXCEL_PATH) as xl:
                sheets_data = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
            
            if project_name not in sheets_data:
                return False
                
            df = sheets_data[project_name]
            if row_id < 0 or row_id >= len(df):
                return False
                
            # Drop the row
            df = df.drop(df.index[row_id]).reset_index(drop=True)
            sheets_data[project_name] = df
            
            # Write back
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
                for sheet, sheet_df in sheets_data.items():
                    sheet_df.to_excel(writer, index=False, sheet_name=sheet)
            return True
        except Exception as e:
            print(f"Error deleting project row: {e}")
            return False

    @staticmethod
    def create_project(project_name):
        """Creates a new sheet in the Excel file."""
        if not os.path.exists(EXCEL_PATH):
            try:
                with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
                    df = pd.DataFrame(columns=["Category"])
                    df.to_excel(writer, index=False, sheet_name=project_name)
                return True
            except Exception as e:
                print(f"Error creating Excel: {e}")
                return False
        try:
            with pd.ExcelFile(EXCEL_PATH) as xl:
                sheets_data = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
            
            if project_name in sheets_data:
                return False  # Already exists
                
            sheets_data[project_name] = pd.DataFrame(columns=[
                "Category", "Test Method", "Test Number", "Start Date",
                "Proto Weeks", "Proto Days", "Proto Qty",
                "DVT Weeks", "DVT Days", "DVT Qty",
                "EVT Weeks", "EVT Days", "EVT Qty",
                "PVT Weeks", "PVT Days", "PVT Qty",
                "Defect Qty", "Comments"
            ])
            
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
                for sheet, sheet_df in sheets_data.items():
                    sheet_df.to_excel(writer, index=False, sheet_name=sheet)
            return True
        except Exception as e:
            print(f"Error creating project: {e}")
            return False

    @staticmethod
    def delete_project(project_name):
        """Deletes a sheet from the Excel file."""
        if not os.path.exists(EXCEL_PATH):
            return False
        try:
            with pd.ExcelFile(EXCEL_PATH) as xl:
                sheets_data = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
            
            if project_name not in sheets_data:
                return False
                
            del sheets_data[project_name]
            
            if not sheets_data:
                sheets_data["Sheet1"] = pd.DataFrame(columns=["Category"])
                
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
                for sheet, sheet_df in sheets_data.items():
                    sheet_df.to_excel(writer, index=False, sheet_name=sheet)
            return True
        except Exception as e:
            print(f"Error deleting project: {e}")
            return False


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
    projects = ExcelDataStore.get_projects()
    if not any(p["name"] == project_name for p in projects):
        return redirect(url_for('tables'))
    return render_template('project_detail.html', project={"name": project_name}, username=session.get('username'), active_page='tables')

# --- API ENDPOINTS ---

@app.route('/api/projects', methods=['GET', 'POST'])
@login_required
def api_projects():
    if request.method == 'GET':
        projects = ExcelDataStore.get_projects()
        return jsonify(projects)
    
    # Create project
    data = request.json
    if not data or 'name' not in data or not str(data['name']).strip():
        return jsonify({"error": "Project name is required"}), 400
        
    name = str(data['name']).strip()
    if ExcelDataStore.create_project(name):
        return jsonify({
            "id": name,
            "name": name,
            "description": f"Excel Sheet: {name}",
            "status": "Active",
            "created_at": "N/A",
            "row_count": 0
        }), 201
    else:
        return jsonify({"error": f"Project '{name}' already exists"}), 400

@app.route('/api/projects/<project_name>', methods=['PUT', 'DELETE'])
@login_required
def api_project_detail(project_name):
    if request.method == 'DELETE':
        if ExcelDataStore.delete_project(project_name):
            return jsonify({"message": f"Project {project_name} deleted successfully"})
        return jsonify({"error": "Project not found"}), 404
        
    # Update project (Rename)
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
        
    if 'name' in data and data['name'] != project_name:
        new_name = str(data['name']).strip()
        
        # Rename sheet in Excel
        if not os.path.exists(EXCEL_PATH):
            return jsonify({"error": "Excel file not found"}), 404
            
        try:
            with pd.ExcelFile(EXCEL_PATH) as xl:
                sheets_data = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
            if new_name in sheets_data:
                return jsonify({"error": f"Project '{new_name}' already exists"}), 400
            if project_name not in sheets_data:
                return jsonify({"error": "Project not found"}), 404
                
            sheets_data[new_name] = sheets_data.pop(project_name)
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
                for sheet, sheet_df in sheets_data.items():
                    sheet_df.to_excel(writer, index=False, sheet_name=sheet)
            return jsonify({
                "id": new_name,
                "name": new_name,
                "description": f"Excel Sheet: {new_name}",
                "status": "Active",
                "created_at": "N/A",
                "row_count": len(sheets_data[new_name])
            })
        except Exception as e:
            return jsonify({"error": f"Failed to rename sheet: {str(e)}"}), 500
            
    return jsonify({
        "id": project_name,
        "name": project_name,
        "description": f"Excel Sheet: {project_name}",
        "status": "Active",
        "created_at": "N/A",
        "row_count": 0
    })

@app.route('/api/projects/<project_name>/rows', methods=['GET', 'POST'])
@login_required
def api_project_rows(project_name):
    if request.method == 'GET':
        records = ExcelDataStore.get_project_rows(project_name)
        return jsonify(records)
        
    # Add Row
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
        
    required = ["Category", "Test Method", "Test Number", "Start Date"]
    for f in required:
        if f not in data or not str(data[f]).strip():
            return jsonify({"error": f"Field '{f}' is required"}), 400
            
    if ExcelDataStore.save_project_row(project_name, None, data):
        rows = ExcelDataStore.get_project_rows(project_name)
        return jsonify(rows[-1] if rows else {}), 201
    return jsonify({"error": "Failed to add row"}), 500

@app.route('/api/projects/<project_name>/rows/<int:row_id>', methods=['PUT', 'DELETE'])
@login_required
def api_project_row_detail(project_name, row_id):
    if request.method == 'DELETE':
        if ExcelDataStore.delete_project_row(project_name, row_id):
            return jsonify({"message": "Row deleted successfully"})
        return jsonify({"error": "Row not found"}), 404
        
    # Update Row
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400
        
    if ExcelDataStore.save_project_row(project_name, row_id, data):
        rows = ExcelDataStore.get_project_rows(project_name)
        return jsonify(rows[row_id] if row_id < len(rows) else {})
    return jsonify({"error": "Failed to update row"}), 500

# --- SHAREPOINT & EXCEL ENDPOINTS ---

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
        with pd.ExcelFile(io.BytesIO(file_content)) as xl:
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
    """Downloads the Excel file from SharePoint and saves it as the active Excel data file."""
    data = request.json
    if not data or 'file_path' not in data:
        return jsonify({"error": "Missing file_path"}), 400
        
    file_path = data['file_path']
    try:
        file_content = SharePointService.download_file(file_path)
        with open(EXCEL_PATH, 'wb') as f:
            f.write(file_content)
            
        # Save sync configuration
        with open(SYNC_CONFIG_PATH, 'w') as f:
            json.dump({"file_path": file_path}, f)
            
        projects = ExcelDataStore.get_projects()
        total_records = sum(p["row_count"] for p in projects)
        return jsonify({
            "success": True,
            "message": f"Successfully synced with SharePoint Excel file. Loaded {len(projects)} sheets and {total_records} records.",
            "projects": [p["name"] for p in projects],
            "count": total_records
        })
    except Exception as e:
        return jsonify({"error": f"Failed to sync SharePoint Excel: {str(e)}"}), 500

@app.route('/api/sync-sharepoint', methods=['POST'])
@login_required
def sync_sharepoint():
    """Triggers a sync of the Excel file from SharePoint using the saved configuration."""
    if not os.path.exists(SYNC_CONFIG_PATH):
        return jsonify({
            "success": False,
            "error": "No saved sync configuration. Please configure the file first using the SharePoint Extractor in the sidebar."
        }), 400
        
    try:
        with open(SYNC_CONFIG_PATH, 'r') as f:
            config = json.load(f)
        file_path = config.get("file_path")
        if not file_path:
            return jsonify({"success": False, "error": "Invalid sync configuration."}), 400
            
        file_content = SharePointService.download_file(file_path)
        with open(EXCEL_PATH, 'wb') as f:
            f.write(file_content)
            
        projects = ExcelDataStore.get_projects()
        total_records = sum(p["row_count"] for p in projects)
        return jsonify({
            "success": True,
            "message": f"SharePoint sync complete — synced {len(projects)} projects with {total_records} records.",
            "projects": [p["name"] for p in projects],
            "count": total_records
        })
    except Exception as e:
        return jsonify({"success": False, "error": f"Failed to sync SharePoint Excel: {str(e)}"}), 500

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
        
        result = {}
        with pd.ExcelFile(io.BytesIO(file_bytes)) as xl:
            sheets = xl.sheet_names

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

@app.route('/api/local/load-transformed', methods=['POST'])
@login_required
def api_local_load_transformed():
    """Imports a locally uploaded Excel file as the active Excel data store."""
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files['file']
    if uploaded_file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    filename = uploaded_file.filename
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({"error": "Only Excel files (.xlsx, .xls) are supported"}), 400

    try:
        uploaded_file.save(EXCEL_PATH)
        
        # Save local configuration
        with open(SYNC_CONFIG_PATH, 'w') as f:
            json.dump({"file_path": f"local://{filename}"}, f)
            
        projects = ExcelDataStore.get_projects()
        total_records = sum(p["row_count"] for p in projects)
        return jsonify({
            "success": True,
            "message": f"Successfully imported {len(projects)} project(s) with {total_records} records from {filename}.",
            "projects": [p["name"] for p in projects],
            "count": total_records
        })
    except Exception as e:
        return jsonify({"error": f"Failed to import local file: {str(e)}"}), 500

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
        result = {}
        with pd.ExcelFile(io.BytesIO(file_content)) as xl:
            sheets = xl.sheet_names
            
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
    """Imports selected sheets from SharePoint Excel file as the active Excel data store."""
    data = request.json
    if not data or 'file_path' not in data:
        return jsonify({"error": "Missing required configuration"}), 400
        
    file_path = data['file_path']
    try:
        file_content = SharePointService.download_file(file_path)
        with open(EXCEL_PATH, 'wb') as f:
            f.write(file_content)
            
        # Save sync configuration
        with open(SYNC_CONFIG_PATH, 'w') as f:
            json.dump({"file_path": file_path}, f)
            
        projects = ExcelDataStore.get_projects()
        total_records = sum(p["row_count"] for p in projects)
        return jsonify({
            "success": True,
            "message": f"Successfully loaded and transformed {len(projects)} projects with {total_records} records.",
            "projects": [p["name"] for p in projects],
            "count": total_records
        })
    except Exception as e:
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
    all_projects = ExcelDataStore.get_projects()
    all_project_names = [p["name"] for p in all_projects]

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
        records = ExcelDataStore.get_project_rows(target_project)
        total_qty = sum(
            (r.get("Proto Qty") or 0) + (r.get("DVT Qty") or 0) + (r.get("EVT Qty") or 0) + (r.get("PVT Qty") or 0)
            for r in records
        )
        rejections = [r for r in records if r.get("Comments") and str(r.get("Comments")).strip()]

        response = (
            f"📊 **Project {target_project} Summary**\n"
            f"Test Records: {len(records)}\n"
            f"Total Qty across all phases: {total_qty}\n"
            f"Records with rejections: {len(rejections)}\n"
        )
        if rejections:
            response += "\nRejection notes:\n"
            for r in rejections[:3]:  # Show max 3
                response += f"  • [{r.get('Category')}] {r.get('Test Method')}: {r.get('Comments')}\n"
            if len(rejections) > 3:
                response += f"  ... and {len(rejections) - 3} more."
        return jsonify({"response": response})

    # --- Intent: Summarize all projects ---
    if any(kw in query for kw in ['summarize all', 'overview all', 'all projects status', 'total projects']):
        total_records = 0
        total_rejections = 0
        for pname in all_project_names:
            rows = ExcelDataStore.get_project_rows(pname)
            total_records += len(rows)
            total_rejections += sum(1 for r in rows if r.get("Comments") and str(r.get("Comments")).strip())
            
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
            records = ExcelDataStore.get_project_rows(target_project)
            rejection_records = [r for r in records if r.get("Comments") and str(r.get("Comments")).strip()]
            if rejection_records:
                resp = f"⚠️ **Rejections in Project {target_project}** ({len(rejection_records)} found):\n"
                for r in rejection_records:
                    resp += f"• [{r.get('Category')}] {r.get('Test Method')} — {r.get('Comments')}\n"
                return jsonify({"response": resp})
            else:
                return jsonify({"response": f"✅ No rejections recorded for Project {target_project}."})
        else:
            total_rej = 0
            for pname in all_project_names:
                rows = ExcelDataStore.get_project_rows(pname)
                total_rej += sum(1 for r in rows if r.get("Comments") and str(r.get("Comments")).strip())
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
def generate_gantt_chart_image(records, filepath, project_name=None, theme='dark', timeline_type='weeks', dpi=300):
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
        plt.savefig(filepath, dpi=dpi, bbox_inches='tight')
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
        plt.savefig(filepath, dpi=dpi, bbox_inches='tight')
        plt.close()
        return

    # Visual style tuned to the theme (dark or light)
    if theme == 'light':
        bg_color = '#ffffff'
        fg_text = '#0f172a'
        grid_color = '#f1f5f9'
        sidebar_bg = '#f8fafc'
        fg_sub = '#64748b'
        today_line_color = '#e11d48'
    else:
        bg_color = '#06293a'  # deep navy
        fg_text = '#eaf6ff'
        grid_color = '#10475a'
        sidebar_bg = '#053242'
        fg_sub = '#cdeefb'
        today_line_color = '#ffcc00'

    fig_height = max(4, len(rows) * 0.6 + 2)

    # Use a two-column layout: left sidebar for labels, right for the gantt
    from matplotlib.patches import FancyBboxPatch, Patch, Rectangle
    fig = plt.figure(figsize=(12, fig_height), facecolor=bg_color)
    gs = fig.add_gridspec(1, 2, width_ratios=[0.28, 0.72], wspace=0.02)
    ax_left = fig.add_subplot(gs[0, 0])
    ax = fig.add_subplot(gs[0, 1])

    ax.set_facecolor(bg_color)
    ax_left.set_facecolor(sidebar_bg)

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
        ax_left.text(0.05, y + 0.28, short, va='center', ha='left', fontsize=8, color=fg_sub)

    # Right axis: Gantt chart area
    start_num = mdates.date2num(min_date - datetime.timedelta(days=1))
    end_num = mdates.date2num(max_date + datetime.timedelta(days=3))
    ax.set_xlim(start_num, end_num)
    ax.set_ylim(-0.5, len(rows) - 0.5)
    ax.invert_yaxis()

    # Vertical grid lines based on timeline_type
    cur_date = min_date
    if timeline_type == 'days':
        while cur_date <= max_date:
            x = mdates.date2num(cur_date)
            ax.axvline(x, color=grid_color, linewidth=0.5, alpha=0.4)
            cur_date += datetime.timedelta(days=1)
    elif timeline_type == 'weeks':
        while cur_date <= max_date:
            x = mdates.date2num(cur_date)
            ax.axvline(x, color=grid_color, linewidth=0.7, alpha=0.6)
            cur_date += datetime.timedelta(days=7)
    elif timeline_type == 'months':
        month_iter = datetime.date(min_date.year, min_date.month, 1)
        while month_iter <= max_date.date():
            x = mdates.date2num(datetime.datetime(month_iter.year, month_iter.month, 1))
            ax.axvline(x, color=grid_color, linewidth=0.8, alpha=0.6)
            if month_iter.month == 12:
                month_iter = datetime.date(month_iter.year + 1, 1, 1)
            else:
                month_iter = datetime.date(month_iter.year, month_iter.month + 1, 1)
    elif timeline_type == 'years':
        year_iter = min_date.year
        while year_iter <= max_date.year:
            x = mdates.date2num(datetime.datetime(year_iter, 1, 1))
            ax.axvline(x, color=grid_color, linewidth=1.0, alpha=0.7)
            year_iter += 1

    # Month labels above the gantt area (only for days and weeks views to avoid clutter)
    if timeline_type in ['days', 'weeks']:
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
            edge = Rectangle((x, y_pos), width, height, linewidth=0.6, edgecolor='#083a4a' if theme == 'dark' else '#cbd5e1', facecolor='none')
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
    if timeline_type == 'days':
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, int((end_num - start_num)/15))))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %d'))
    elif timeline_type == 'weeks':
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %d'))
    elif timeline_type == 'months':
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
    elif timeline_type == 'years':
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))

    plt.setp(ax.get_xticklabels(), rotation=30, ha='right', color=fg_text, fontsize=8)

    # Clean spines
    for spine in ['top', 'right', 'left', 'bottom']:
        ax.spines[spine].set_visible(False)

    # Legend and title
    legend_elements = [Patch(facecolor=color, edgecolor='none', label=phase) for phase, color in phase_colors.items()]
    ax.legend(handles=legend_elements, loc='upper right', frameon=False, fontsize=9, labelcolor=fg_text)

    today = datetime.datetime.now()
    if min_date <= today <= max_date:
        ax.axvline(mdates.date2num(today), color=today_line_color, linewidth=1.2, linestyle='--', alpha=0.9)

    timeline_labels = {
        'days': 'Daily Timeline',
        'weeks': 'Weekly Timeline',
        'months': 'Monthly Timeline',
        'years': 'Yearly Timeline'
    }
    scale_label = timeline_labels.get(timeline_type, 'Timeline')
    title_text = f"Project Gantt Timeline — Project {project_name} ({scale_label})" if project_name else f"Project Gantt Timeline ({scale_label})"
    ax.set_title(title_text, fontsize=14, fontweight='bold', pad=14, color=fg_text)

    plt.tight_layout()
    plt.savefig(filepath, dpi=dpi, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig)


# --- PDF REPORT DYNAMIC PAGE NUMBERING ---
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748b"))
        
        # Top Accent Double Stripe Bar
        self.setFillColor(colors.HexColor("#0f172a")) # Primary Slate 900
        self.rect(0, 782, 612, 10, fill=True, stroke=False)
        self.setFillColor(colors.HexColor("#0284c7")) # Accent Sky 600
        self.rect(0, 778, 612, 4, fill=True, stroke=False)
        
        # Header line
        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(0.75)
        self.line(40, 745, 572, 745)
        
        self.setFillColor(colors.HexColor("#0f172a"))
        self.setFont("Helvetica-Bold", 8)
        
        # Draw Excel report name if available, otherwise default
        excel_name = getattr(self, 'excel_filename', '')
        header_left = f"EXCEL REPORT: {excel_name.upper()}" if excel_name else "TEST OPERATIONS DASHBOARD"
        self.drawString(40, 750, header_left)
        
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748b"))
        self.drawRightString(572, 750, "PROJECT STATUS REPORT")
            
        # Footer
        self.line(40, 45, 572, 45) # Footer line
        page_text = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(572, 32, page_text)
        self.drawString(40, 32, "CONFIDENTIAL - INTERNAL USE ONLY")
        self.drawCentredString(306, 32, datetime.datetime.now().strftime("%Y-%m-%d"))
        
        self.restoreState()


# --- PDF REPORT GENERATION ---
def generate_pdf_report(filename, project_name=None, comment=None):
    """Generates a styled Test Operations PDF report containing a Gantt chart and rejected products."""
    # Fetch records filtered by project
    if project_name:
        records_dict = ExcelDataStore.get_project_rows(project_name)
    else:
        records_dict = []
        for p in ExcelDataStore.get_projects():
            records_dict.extend(ExcelDataStore.get_project_rows(p["name"]))
        
    # 1. Generate the Gantt Chart Image
    gantt_image_path = os.path.join(UPLOAD_FOLDER, "gantt_chart_report.png")
    generate_gantt_chart_image(records_dict, gantt_image_path, project_name, theme='light')
    
    # Extract Excel filename
    excel_filename = ""
    if os.path.exists(SYNC_CONFIG_PATH):
        try:
            with open(SYNC_CONFIG_PATH, 'r') as f:
                config = json.load(f)
            path = config.get("file_path", "")
            if "://" in path:
                path = path.split("://", 1)[1]
            excel_filename = os.path.basename(path)
        except Exception:
            pass
    if not excel_filename:
        excel_filename = "tasks.xlsx"
        
    doc = SimpleDocTemplate(filename, pagesize=letter,
                            rightMargin=40, leftMargin=40, topMargin=60, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    # Custom color palette
    primary_color = colors.HexColor("#0f172a")
    secondary_color = colors.HexColor("#475569")
    accent_color = colors.HexColor("#0284c7")
    bg_light = colors.HexColor("#f8fafc")
    border_color = colors.HexColor("#e2e8f0")
    danger_red = colors.HexColor("#e11d48")
    
    title_style = ParagraphStyle(
        'CampTitle',
        parent=styles['Heading1'],
        fontSize=24,
        leading=28,
        textColor=primary_color,
        fontName='Helvetica-Bold',
        spaceAfter=4
    )
    
    subtitle_style = ParagraphStyle(
        'CampSubtitle',
        parent=styles['Normal'],
        fontSize=9,
        leading=13,
        textColor=secondary_color,
        fontName='Helvetica',
        spaceAfter=15
    )
    
    section_title_style = ParagraphStyle(
        'CampSection',
        parent=styles['Heading2'],
        fontSize=12,
        leading=16,
        textColor=primary_color,
        fontName='Helvetica-Bold',
        spaceBefore=14,
        spaceAfter=8,
        keepWithNext=True
    )
    
    cell_header = ParagraphStyle(
        'CampHeader',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.white,
        fontName='Helvetica-Bold'
    )
    
    cell_style = ParagraphStyle(
        'CampCell',
        parent=styles['Normal'],
        fontSize=8,
        leading=11,
        textColor=primary_color,
        fontName='Helvetica'
    )
    
    cell_bold = ParagraphStyle(
        'CampCellBold',
        parent=styles['Normal'],
        fontSize=8,
        leading=11,
        textColor=danger_red,
        fontName='Helvetica-Bold'
    )
    
    def create_section_header(title_text):
        p = Paragraph(title_text, section_title_style)
        t = Table([[p]], colWidths=[532])
        t.setStyle(TableStyle([
            ('LINEBELOW', (0,0), (-1,-1), 1.5, accent_color),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        return t

    # Header Title
    doc_title = f"Project Status Report - Project {project_name}" if project_name else "Daily Test Operations Report"
    story.append(Paragraph(doc_title, title_style))
    story.append(Paragraph(f"Generated on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Confidential Operations Data", subtitle_style))

    # KPI summary cards
    total_records = len(records_dict)
    total_defects = sum(int(r.get('Defect Qty', 0) or 0) for r in records_dict)
    total_comments = sum(1 for r in records_dict if str(r.get('Comments', '')).strip())

    kpi_table_data = [
        [
            Paragraph(f"<font size=7.5 color='#64748b'>PROJECT NAME</font><br/><font size=11 color='#0f172a'><b>{project_name or 'ALL PROJECTS'}</b></font>", cell_style),
            Paragraph(f"<font size=7.5 color='#64748b'>TOTAL TASKS</font><br/><font size=11 color='#0f172a'><b>{total_records}</b></font>", cell_style),
            Paragraph(f"<font size=7.5 color='#64748b'>TOTAL DEFECTS</font><br/><font size=11 color='#e11d48'><b>{total_defects}</b></font>", cell_style),
            Paragraph(f"<font size=7.5 color='#64748b'>REJECTIONS</font><br/><font size=11 color='#d97706'><b>{total_comments}</b></font>", cell_style),
        ]
    ]
    kpi_table = Table(kpi_table_data, colWidths=[133, 133, 133, 133])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg_light),
        ('BOX', (0,0), (-1,-1), 1, border_color),
        ('INNERGRID', (0,0), (-1,-1), 0.5, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 10))

    if comment and str(comment).strip():
        story.append(create_section_header("Report Comment"))
        story.append(Spacer(1, 4))
        story.append(Paragraph(str(comment).strip().replace('\n', '<br/>'), subtitle_style))
        story.append(Spacer(1, 6))

    story.append(create_section_header("Project Gantt Timeline"))
    story.append(Spacer(1, 4))
    
    # Wrap Gantt chart image in a table with a subtle border
    gantt_table = Table([[Image(gantt_image_path, width=524, height=275)]], colWidths=[532])
    gantt_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(gantt_table)
    story.append(Spacer(1, 12))

    story.append(create_section_header("Test Records Summary"))
    story.append(Spacer(1, 6))
    
    all_table_data = [[
        Paragraph("#", cell_header),
        Paragraph("Category", cell_header),
        Paragraph("Test Method", cell_header),
        Paragraph("Test #", cell_header),
        Paragraph("Proto", cell_header),
        Paragraph("DVT", cell_header),
        Paragraph("EVT", cell_header),
        Paragraph("PVT", cell_header),
        Paragraph("Defect", cell_header),
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
            Paragraph(f"{r.get('Proto Weeks',0)}w {r.get('Proto Days',0)}d<br/><font color='#64748b'>Qty: {r.get('Proto Qty',0)}</font>", cell_style),
            Paragraph(f"{r.get('DVT Weeks',0)}w {r.get('DVT Days',0)}d<br/><font color='#64748b'>Qty: {r.get('DVT Qty',0)}</font>", cell_style),
            Paragraph(f"{r.get('EVT Weeks',0)}w {r.get('EVT Days',0)}d<br/><font color='#64748b'>Qty: {r.get('EVT Qty',0)}</font>", cell_style),
            Paragraph(f"{r.get('PVT Weeks',0)}w {r.get('PVT Days',0)}d<br/><font color='#64748b'>Qty: {r.get('PVT Qty',0)}</font>", cell_style),
            Paragraph(str(defect_qty), defect_style),
            Paragraph(r.get("Comments", "-") or "-", cell_bold if r.get("Comments") else cell_style)
        ])
    if not records_dict:
        all_table_data.append([Paragraph("No records found.", cell_style)] + [Paragraph("-", cell_style)] * 9)

    all_table = Table(all_table_data, colWidths=[15, 70, 80, 45, 52, 52, 52, 52, 30, 84])
    all_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_color),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('ALIGN', (8,0), (8,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LINEBELOW', (0,0), (-1,0), 1, primary_color),
        ('LINEBELOW', (0,1), (-1,-1), 0.5, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, bg_light])
    ]))
    story.append(all_table)
    story.append(Spacer(1, 14))

    # Rejections-only sub-table
    rejected_records = [r for r in records_dict if str(r.get("Comments", "")).strip() != "" or int(r.get("Defect Qty", 0)) > 0]
    story.append(create_section_header("Rejected Products Inventory"))
    story.append(Spacer(1, 6))
    
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

    tasks_table = Table(table_data, colWidths=[100, 110, 65, 55, 202])
    tasks_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_color),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LINEBELOW', (0,0), (-1,0), 1, primary_color),
        ('LINEBELOW', (0,1), (-1,-1), 0.5, border_color),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, bg_light])
    ]))
    story.append(tasks_table)
    
    # Set the excel_filename on NumberedCanvas class so it can be accessed during build
    NumberedCanvas.excel_filename = excel_filename
    doc.build(story, canvasmaker=NumberedCanvas)


# --- CONSOLIDATED PDF REPORT GENERATION ---
def render_gantt_task_module(records_dict, img_path, project_name, timeline_type, dpi):
    try:
        generate_gantt_chart_image(records_dict, img_path, project_name=project_name, theme='light', timeline_type=timeline_type, dpi=dpi)
        return True
    except Exception as e:
        print(f"Error rendering Gantt chart {project_name} ({timeline_type}): {e}")
        return False

def generate_consolidated_pdf_report(filename, progress_callback=None):
    """Generates a consolidated PDF report for ALL active projects, displaying multiple timelines and defect tables."""
    # Fetch all active projects
    projects = ExcelDataStore.get_projects()
    
    # Extract Excel filename
    excel_filename = ""
    if os.path.exists(SYNC_CONFIG_PATH):
        try:
            with open(SYNC_CONFIG_PATH, 'r') as f:
                config = json.load(f)
            path = config.get("file_path", "")
            if "://" in path:
                path = path.split("://", 1)[1]
            excel_filename = os.path.basename(path)
        except Exception:
            pass
    if not excel_filename:
        excel_filename = "tasks.xlsx"
        
    doc = SimpleDocTemplate(filename, pagesize=letter,
                            rightMargin=40, leftMargin=40, topMargin=60, bottomMargin=60)
    story = []
    styles = getSampleStyleSheet()
    
    # Custom color palette
    primary_color = colors.HexColor("#0f172a")
    secondary_color = colors.HexColor("#475569")
    accent_color = colors.HexColor("#0284c7")
    bg_light = colors.HexColor("#f8fafc")
    border_color = colors.HexColor("#e2e8f0")
    danger_red = colors.HexColor("#e11d48")
    
    title_style = ParagraphStyle(
        'OpsTitle',
        parent=styles['Heading1'],
        fontSize=24,
        leading=28,
        textColor=primary_color,
        fontName='Helvetica-Bold',
        spaceAfter=4
    )
    
    subtitle_style = ParagraphStyle(
        'OpsSubtitle',
        parent=styles['Normal'],
        fontSize=9,
        leading=13,
        textColor=secondary_color,
        fontName='Helvetica',
        spaceAfter=15
    )
    
    section_title_style = ParagraphStyle(
        'ConsolidatedSection',
        parent=styles['Heading2'],
        fontSize=14,
        leading=18,
        textColor=primary_color,
        fontName='Helvetica-Bold',
        spaceBefore=18,
        spaceAfter=10,
        keepWithNext=True
    )
    
    subsection_title_style = ParagraphStyle(
        'ConsolidatedSubSection',
        parent=styles['Heading3'],
        fontSize=11,
        leading=15,
        textColor=secondary_color,
        fontName='Helvetica-Bold',
        spaceBefore=10,
        spaceAfter=6,
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
        textColor=primary_color,
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
    
    callout_style = ParagraphStyle(
        'CalloutText',
        parent=styles['Normal'],
        fontSize=9,
        leading=13,
        textColor=secondary_color,
        fontName='Helvetica-Oblique'
    )

    def create_section_header(title_text):
        p = Paragraph(title_text, section_title_style)
        t = Table([[p]], colWidths=[532])
        t.setStyle(TableStyle([
            ('LINEBELOW', (0,0), (-1,-1), 1.5, accent_color),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        return t

    # Document Header
    story.append(Paragraph("Consolidated Project Status Report", title_style))
    story.append(Paragraph(f"Excel Source: {excel_filename} | Compiled on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subtitle_style))
    story.append(Spacer(1, 10))
    
    temp_files = []
    
    # 1. Collect all rendering tasks
    render_tasks = []
    project_records = {}
    
    for proj in projects:
        pname = proj["name"]
        records_dict = ExcelDataStore.get_project_rows(pname)
        if not records_dict:
            continue
        project_records[pname] = records_dict
        
        timelines = ['days', 'weeks', 'months', 'years']
        for t_type in timelines:
            img_path = os.path.join(UPLOAD_FOLDER, f"temp_gantt_{pname}_{t_type}.jpg")
            render_tasks.append((records_dict, img_path, pname, t_type, 80))
            temp_files.append(img_path)

    # 2. Run rendering tasks in parallel using ProcessPoolExecutor
    total_tasks = len(render_tasks)
    completed_tasks = 0
    
    if progress_callback:
        progress_callback(5, f"Spawning worker processes for {total_tasks} charts...")
        
    try:
        from concurrent.futures import ProcessPoolExecutor, as_completed
        with ProcessPoolExecutor() as executor:
            futures = {
                executor.submit(render_gantt_task_module, t[0], t[1], t[2], t[3], t[4]): t
                for t in render_tasks
            }
            
            for future in as_completed(futures):
                completed_tasks += 1
                if progress_callback:
                    pct = int((completed_tasks / max(1, total_tasks)) * 85) + 5
                    progress_callback(pct, f"Rendered chart {completed_tasks}/{total_tasks}...")
    except Exception as e:
        print(f"Error during parallel rendering: {e}")
        
    try:
        from reportlab.platypus import PageBreak
        
        valid_project_count = 0
        for pname, records_dict in project_records.items():
            if not records_dict:
                continue
                
            # If not the first project, start on a new page
            if valid_project_count > 0:
                story.append(PageBreak())
            valid_project_count += 1
                
            story.append(create_section_header(f"Project Sheet: {pname}"))
            story.append(Spacer(1, 10))
            
            img_days = os.path.join(UPLOAD_FOLDER, f"temp_gantt_{pname}_days.jpg")
            img_weeks = os.path.join(UPLOAD_FOLDER, f"temp_gantt_{pname}_weeks.jpg")
            img_months = os.path.join(UPLOAD_FOLDER, f"temp_gantt_{pname}_months.jpg")
            img_years = os.path.join(UPLOAD_FOLDER, f"temp_gantt_{pname}_years.jpg")
            
            # Add Gantt charts in pairs (2 per page)
            # Pair 1: Days and Weeks
            story.append(Paragraph("Daily & Weekly Timelines", subsection_title_style))
            t_table_1_data = []
            row_images = []
            if os.path.exists(img_days):
                row_images.append(Image(img_days, width=260, height=140))
            else:
                row_images.append(Paragraph("Failed to render daily timeline.", cell_style))
            if os.path.exists(img_weeks):
                row_images.append(Image(img_weeks, width=260, height=140))
            else:
                row_images.append(Paragraph("Failed to render weekly timeline.", cell_style))
            t_table_1_data.append(row_images)
            
            t_table_1 = Table(t_table_1_data, colWidths=[266, 266])
            t_table_1.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 2),
                ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ]))
            story.append(t_table_1)
            story.append(Spacer(1, 15))
            
            # Pair 2: Months and Years
            story.append(Paragraph("Monthly & Yearly Timelines", subsection_title_style))
            t_table_2_data = []
            row_images_2 = []
            if os.path.exists(img_months):
                row_images_2.append(Image(img_months, width=260, height=140))
            else:
                row_images_2.append(Paragraph("Failed to render monthly timeline.", cell_style))
            if os.path.exists(img_years):
                row_images_2.append(Image(img_years, width=260, height=140))
            else:
                row_images_2.append(Paragraph("Failed to render yearly timeline.", cell_style))
            t_table_2_data.append(row_images_2)
            
            t_table_2 = Table(t_table_2_data, colWidths=[266, 266])
            t_table_2.setStyle(TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 2),
                ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ]))
            story.append(t_table_2)
            story.append(Spacer(1, 15))
            
            # Defect & Rejection Analysis Table
            story.append(Paragraph("Defect & Rejection Analysis", subsection_title_style))
            
            # Filter records with defects
            defective_records = [r for r in records_dict if int(r.get("Defect Qty", 0) or 0) > 0]
            
            if defective_records:
                # Render defect table
                table_data = [[
                    Paragraph("Category (Product)", cell_header),
                    Paragraph("Test Method", cell_header),
                    Paragraph("Test Number", cell_header),
                    Paragraph("Defect Qty", cell_header)
                ]]
                for r in defective_records:
                    table_data.append([
                        Paragraph(r.get("Category", ""), cell_style),
                        Paragraph(r.get("Test Method", ""), cell_style),
                        Paragraph(r.get("Test Number", ""), cell_style),
                        Paragraph(str(r.get("Defect Qty", 0)), cell_bold)
                    ])
                
                defect_table = Table(table_data, colWidths=[132, 180, 100, 120])
                defect_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), primary_color),
                    ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('LINEBELOW', (0,0), (-1,0), 1, primary_color),
                    ('LINEBELOW', (0,1), (-1,-1), 0.5, border_color),
                    ('TOPPADDING', (0,0), (-1,-1), 5),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                    ('LEFTPADDING', (0,0), (-1,-1), 6),
                    ('RIGHTPADDING', (0,0), (-1,-1), 6),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, bg_light])
                ]))
                story.append(defect_table)
                story.append(Spacer(1, 10))
                
                # Render rejection comments below the table
                story.append(Paragraph("<b>Rejection Comments & Defect Reasons:</b>", cell_style))
                story.append(Spacer(1, 4))
                
                for r in defective_records:
                    comment_text = r.get("Comments", "").strip()
                    if not comment_text:
                        comment_text = "No comments/reasons recorded for this defect."
                    
                    prod_info = f"<b>{r.get('Category')}</b> ({r.get('Test Method')} - {r.get('Test Number')}):"
                    story.append(Paragraph(f"• {prod_info} {comment_text}", cell_style))
                    story.append(Spacer(1, 2))
            else:
                # Render "no issues" message
                no_issue_table = Table([[
                    Paragraph("No issues occurred during the test methods processes.", callout_style)
                ]], colWidths=[532])
                no_issue_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,-1), bg_light),
                    ('BOX', (0,0), (-1,-1), 1, border_color),
                    ('TOPPADDING', (0,0), (-1,-1), 10),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 10),
                    ('LEFTPADDING', (0,0), (-1,-1), 12),
                    ('RIGHTPADDING', (0,0), (-1,-1), 12),
                ]))
                story.append(no_issue_table)
                
            story.append(Spacer(1, 15))
            
        # Build document
        # Set the excel_filename on NumberedCanvas class so it can be accessed during build
        NumberedCanvas.excel_filename = excel_filename
        
        if progress_callback:
            progress_callback(95, "Compiling PDF document...")
            
        doc.build(story, canvasmaker=NumberedCanvas)
        
        if progress_callback:
            progress_callback(100, "Completed!")
        
    finally:
        # Clean up temporary files
        for f in temp_files:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except Exception as ex:
                print(f"Error removing temp file {f}: {ex}")


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

@app.route('/generate-consolidated-report', methods=['GET'])
@login_required
def generate_consolidated_report():
    pdf_filename = os.path.join(UPLOAD_FOLDER, "Consolidated_Project_Report.pdf")
    try:
        generate_consolidated_pdf_report(pdf_filename)
        return send_file(pdf_filename, as_attachment=True, download_name="Consolidated_Project_Report.pdf", mimetype='application/pdf')
    except Exception as e:
        return jsonify({"error": f"Failed to generate consolidated PDF: {str(e)}"}), 500

# --- ASYNCHRONOUS EXPORT PROGRESS ENDPOINTS ---
export_tasks = {}

@app.route('/api/export-consolidated/start', methods=['POST'])
@login_required
def start_consolidated_export():
    task_id = str(uuid.uuid4())
    export_tasks[task_id] = {
        "status": "processing",
        "progress": 0,
        "message": "Initializing consolidated PDF generation..."
    }
    
    def run_compilation():
        temp_pdf = os.path.join(UPLOAD_FOLDER, f"consolidated_{task_id}.pdf")
        
        def update_progress(percent, msg):
            if task_id in export_tasks:
                export_tasks[task_id]["progress"] = percent
                export_tasks[task_id]["message"] = msg
            
        try:
            generate_consolidated_pdf_report(temp_pdf, progress_callback=update_progress)
            if task_id in export_tasks:
                export_tasks[task_id]["status"] = "completed"
                export_tasks[task_id]["progress"] = 100
                export_tasks[task_id]["message"] = "Completed!"
        except Exception as e:
            if task_id in export_tasks:
                export_tasks[task_id]["status"] = "failed"
                export_tasks[task_id]["progress"] = 0
                export_tasks[task_id]["message"] = f"Failed: {str(e)}"
            import traceback
            traceback.print_exc()

    thread = threading.Thread(target=run_compilation)
    thread.daemon = True
    thread.start()
    
    return jsonify({"task_id": task_id})

@app.route('/api/export-consolidated/progress/<task_id>', methods=['GET'])
@login_required
def consolidated_export_progress(task_id):
    state = export_tasks.get(task_id)
    if not state:
        return jsonify({"status": "failed", "progress": 0, "message": "Task not found"}), 404
    return jsonify(state)

@app.route('/api/export-consolidated/download/<task_id>/Consolidated_Project_Report.pdf', methods=['GET'])
@login_required
def consolidated_export_download(task_id):
    temp_pdf = os.path.join(UPLOAD_FOLDER, f"consolidated_{task_id}.pdf")
    if not os.path.exists(temp_pdf):
        return jsonify({"error": "Compiled report file not found or expired"}), 404
        
    try:
        def generate_pdf_stream():
            try:
                with open(temp_pdf, 'rb') as f:
                    while True:
                        chunk = f.read(8192)
                        if not chunk:
                            break
                        yield chunk
            finally:
                if task_id in export_tasks:
                    del export_tasks[task_id]
                try:
                    if os.path.exists(temp_pdf):
                        os.remove(temp_pdf)
                except Exception as ex:
                    print(f"Error removing temp export PDF file: {ex}")

        from flask import Response
        response = Response(generate_pdf_stream(), mimetype='application/pdf')
        response.headers['Content-Disposition'] = 'attachment; filename="Consolidated_Project_Report.pdf"'
        return response
    except Exception as e:
        return jsonify({"error": f"Failed to stream PDF: {str(e)}"}), 500

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

@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True)
