import os
import datetime
import sys
import json
import uuid
import threading
from functools import wraps
from dotenv import load_dotenv, find_dotenv
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
load_dotenv(find_dotenv(filename='.env', raise_error_if_not_found=False), override=True)
app.secret_key = "super_secret_key_for_ops_operations"

# Configure App Data and Upload Folders
UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
EXCEL_PATH = os.path.join(UPLOAD_FOLDER, 'tasks.xlsx')
SYNC_CONFIG_PATH = os.path.join(UPLOAD_FOLDER, 'sync_config.json')
EMAIL_ATTACHMENT_MAX_MB = int(os.getenv('EMAIL_ATTACHMENT_MAX_MB', 18))

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
    
    # Check if headers are unnamed (e.g. starting with "Unnamed:")
    is_unnamed = any(str(col).startswith("Unnamed:") for col in df.columns)
    
    if is_unnamed:
        has_sub_header = True
    else:
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
        else:
            sub_lower = sub_str.lower()
            if any(p in sub_lower for p in ['proto', 'dvt', 'evt', 'pvt', 'other']):
                current_main_header = sub_str
            
        if current_main_header and sub_str:
            if sub_str == current_main_header:
                combined = current_main_header
            elif sub_str.lower().startswith(current_main_header.lower()):
                combined = sub_str
            else:
                combined = f"{current_main_header}_{sub_str}"
        elif current_main_header:
            combined = current_main_header
        elif sub_str:
            combined = sub_str
        else:
            combined = col_str
            
        new_columns.append(combined)
        
    # Deduplicate column names to ensure uniqueness
    seen = {}
    deduped_columns = []
    for col in new_columns:
        col_str = str(col)
        if col_str in seen:
            seen[col_str] += 1
            deduped_columns.append(f"{col_str}.{seen[col_str]}")
        else:
            seen[col_str] = 0
            deduped_columns.append(col_str)
            
    df.columns = deduped_columns
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
    
    "proto_weeks": ['proto wk', 'proto week', 'proto_wk', 'proto_weeks', 'proto weeks', 'week1', 'proto1_date (week)', 'proto_date (week)', 'proto1_week', 'proto_week'],
    "proto_days": ['proto day', 'proto_day', 'proto_days', 'proto days', 'day1', 'proto1_date (day)', 'proto_date (day)', 'proto1_day', 'proto_day'],
    "proto_qty": ['proto qty', 'proto_qty', 'proto_quantity', 'proto quantity', 'proto1', 'proto1_qty', 'proto_qty'],
    
    "dvt_weeks": ['dvt wk', 'dvt week', 'dvt_wk', 'dvt_weeks', 'dvt weeks', 'week2', 'dvt1_date (week)', 'dvt_date (week)', 'dvt1_week', 'dvt_week'],
    "dvt_days": ['dvt day', 'dvt_day', 'dvt_days', 'dvt days', 'day3', 'dvt1_date (day)', 'dvt_date (day)', 'dvt1_day', 'dvt_day'],
    "dvt_qty": ['dvt qty', 'dvt_qty', 'dvt_quantity', 'dvt quantity', 'dvt1', 'dvt1_qty', 'dvt_qty'],
    
    "evt_weeks": ['evt wk', 'evt week', 'evt_wk', 'evt_weeks', 'evt weeks', 'week4', 'evt1_date (week)', 'evt_date (week)', 'evt1_week', 'evt_week'],
    "evt_days": ['evt day', 'evt_day', 'evt_days', 'evt days', 'day5', 'evt1_date (day)', 'evt_date (day)', 'evt1_day', 'evt_day'],
    "evt_qty": ['evt qty', 'evt_qty', 'evt_quantity', 'evt quantity', 'evt1', 'evt1_qty', 'evt_qty'],
    
    "pvt_weeks": ['pvt wk', 'pvt week', 'pvt_wk', 'pvt_weeks', 'pvt weeks', 'week6', 'pvt1_date (week)', 'pvt_date (week)', 'pvt1_week', 'pvt_week'],
    "pvt_days": ['pvt day', 'pvt_day', 'pvt_days', 'pvt days', 'day7', 'pvt1_date (day)', 'pvt_date (day)', 'pvt1_day', 'pvt_day'],
    "pvt_qty": ['pvt qty', 'pvt_qty', 'pvt_quantity', 'pvt quantity', 'pvt1', 'pvt1_qty', 'pvt_qty'],
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
                    # Skip matching week/day/qty columns as start_date
                    if field == 'start_date' and any(x in col_lower for x in ['week', 'day', 'qty']):
                        if 'start date' not in col_lower and col_lower != 'start' and col_lower != 'date':
                            continue
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
            df = clean_dataframe_headers(df)
            
            # Determine mapping to write back to correct columns
            mapping = {}
            for field, rules in FUZZY_RULES.items():
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    # Skip matching week/day/qty columns as start_date
                    if field == 'start_date' and any(x in col_lower for x in ['week', 'day', 'qty']):
                        if 'start date' not in col_lower and col_lower != 'start' and col_lower != 'date':
                            continue
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
            df = clean_dataframe_headers(df)
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
    selected_sheets = data.get('selected_sheets', [])
    removed_columns = data.get('removed_columns', {})
    
    # Try to load existing configuration if same file_path and no sheets provided
    if not selected_sheets and os.path.exists(SYNC_CONFIG_PATH):
        try:
            with open(SYNC_CONFIG_PATH, 'r') as f:
                old_config = json.load(f)
            if old_config.get("file_path") == file_path:
                selected_sheets = old_config.get("selected_sheets", [])
                removed_columns = old_config.get("removed_columns", {})
        except Exception:
            pass
            
    try:
        file_content = SharePointService.download_file(file_path)
        
        import io
        sheets_data = {}
        with pd.ExcelFile(io.BytesIO(file_content)) as xl:
            all_sheets = xl.sheet_names
            if not selected_sheets:
                selected_sheets = [s for s in all_sheets if s.lower().strip() != 'master data']
                
            for sheet in selected_sheets:
                if sheet in all_sheets:
                    df = xl.parse(sheet)
                    df = clean_dataframe_headers(df)
                    removed = removed_columns.get(sheet, [])
                    if removed:
                        df = df.drop(columns=[col for col in removed if col in df.columns], errors='ignore')
                    sheets_data[sheet] = df
                    
        if not sheets_data:
            sheets_data["Sheet1"] = pd.DataFrame(columns=["Category"])
            
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            for sheet, df in sheets_data.items():
                df.to_excel(writer, index=False, sheet_name=sheet)
                
        # Save sync configuration
        with open(SYNC_CONFIG_PATH, 'w') as f:
            json.dump({
                "file_path": file_path,
                "selected_sheets": selected_sheets,
                "removed_columns": removed_columns
            }, f)
            
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
        selected_sheets = config.get("selected_sheets", [])
        removed_columns = config.get("removed_columns", {})
        if not file_path:
            return jsonify({"success": False, "error": "Invalid sync configuration."}), 400
            
        file_content = SharePointService.download_file(file_path)
        
        import io
        sheets_data = {}
        with pd.ExcelFile(io.BytesIO(file_content)) as xl:
            all_sheets = xl.sheet_names
            if not selected_sheets:
                selected_sheets = [s for s in all_sheets if s.lower().strip() != 'master data']
                
            for sheet in selected_sheets:
                if sheet in all_sheets:
                    df = xl.parse(sheet)
                    df = clean_dataframe_headers(df)
                    removed = removed_columns.get(sheet, [])
                    if removed:
                        df = df.drop(columns=[col for col in removed if col in df.columns], errors='ignore')
                    sheets_data[sheet] = df
                    
        if not sheets_data:
            sheets_data["Sheet1"] = pd.DataFrame(columns=["Category"])
            
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            for sheet, df in sheets_data.items():
                df.to_excel(writer, index=False, sheet_name=sheet)
                
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
        import io
        file_bytes = uploaded_file.read()
        
        config_str = request.form.get('config')
        selected_sheets = []
        removed_columns = {}
        if config_str:
            try:
                config = json.loads(config_str)
                selected_sheets = config.get('selected_sheets', [])
                removed_columns = config.get('removed_columns', {})
            except Exception as e:
                print(f"Error parsing config: {e}")
                
        sheets_data = {}
        with pd.ExcelFile(io.BytesIO(file_bytes)) as xl:
            all_sheets = xl.sheet_names
            if not selected_sheets:
                selected_sheets = [s for s in all_sheets if s.lower().strip() != 'master data']
                
            for sheet in selected_sheets:
                if sheet in all_sheets:
                    df = xl.parse(sheet)
                    df = clean_dataframe_headers(df)
                    removed = removed_columns.get(sheet, [])
                    if removed:
                        df = df.drop(columns=[col for col in removed if col in df.columns], errors='ignore')
                    sheets_data[sheet] = df
                    
        if not sheets_data:
            sheets_data["Sheet1"] = pd.DataFrame(columns=["Category"])
            
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            for sheet, df in sheets_data.items():
                df.to_excel(writer, index=False, sheet_name=sheet)
                
        # Save local configuration with selected sheets and transformations
        with open(SYNC_CONFIG_PATH, 'w') as f:
            json.dump({
                "file_path": f"local://{filename}",
                "selected_sheets": selected_sheets,
                "removed_columns": removed_columns
            }, f)
            
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
    selected_sheets = data.get('selected_sheets', [])
    removed_columns = data.get('removed_columns', {})
    
    try:
        file_content = SharePointService.download_file(file_path)
        
        import io
        sheets_data = {}
        with pd.ExcelFile(io.BytesIO(file_content)) as xl:
            all_sheets = xl.sheet_names
            if not selected_sheets:
                selected_sheets = [s for s in all_sheets if s.lower().strip() != 'master data']
                
            for sheet in selected_sheets:
                if sheet in all_sheets:
                    df = xl.parse(sheet)
                    df = clean_dataframe_headers(df)
                    removed = removed_columns.get(sheet, [])
                    if removed:
                        df = df.drop(columns=[col for col in removed if col in df.columns], errors='ignore')
                    sheets_data[sheet] = df
                    
        if not sheets_data:
            sheets_data["Sheet1"] = pd.DataFrame(columns=["Category"])
            
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            for sheet, df in sheets_data.items():
                df.to_excel(writer, index=False, sheet_name=sheet)
                
        # Save sync configuration with selected sheets and transformations
        with open(SYNC_CONFIG_PATH, 'w') as f:
            json.dump({
                "file_path": file_path,
                "selected_sheets": selected_sheets,
                "removed_columns": removed_columns
            }, f)
            
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
def parse_date(value):
    if not value:
        return None
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, datetime.date):
        return datetime.datetime.combine(value, datetime.time.min)
    try:
        return datetime.datetime.strptime(str(value).strip(), "%Y-%m-%d")
    except Exception:
        try:
            return datetime.datetime.fromisoformat(str(value).strip())
        except Exception:
            return None

# --- GANTT CHART GENERATION ---
def generate_gantt_chart_image(records, filepath, project_name=None, theme='dark', timeline_type='weeks', dpi=300, start_date_bound=None, end_date_bound=None):
    """Generates a multi-row project Gantt timeline image using Matplotlib and saves it."""
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
    min_date = start_date_bound
    max_date = end_date_bound

    for record in records:
        start_date = parse_date(record.get("Start Date"))
        if not start_date:
            continue

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
            phase_start = phase_end

        # Calculate duration summary text column
        duration_parts = []
        for phase_name, weeks, days in phases:
            qty = int(record.get(f"{phase_name} Qty", 0) or 0)
            if weeks > 0 or days > 0:
                dur_str = f"{weeks}w" if weeks > 0 else ""
                if days > 0:
                    dur_str += f"{days}d"
                duration_parts.append(f"{phase_name[0]}:{dur_str}(Q:{qty})")
        duration_summary = " | ".join(duration_parts)

        # Filter and clamp if bounds are specified
        if start_date_bound and end_date_bound:
            overlapping_bars = []
            for bar in bars:
                bar_start = bar["start"]
                bar_end = bar_start + datetime.timedelta(days=bar["duration"])
                if bar_start < end_date_bound and bar_end > start_date_bound:
                    clamped_start = max(bar_start, start_date_bound)
                    clamped_end = min(bar_end, end_date_bound)
                    clamped_duration = (clamped_end - clamped_start).days
                    if clamped_duration > 0:
                        overlapping_bars.append({
                            "phase": bar["phase"],
                            "start": clamped_start,
                            "duration": clamped_duration,
                            "color": bar["color"]
                        })
            if overlapping_bars:
                rows.append({
                    "label": label,
                    "bars": overlapping_bars,
                    "category": record.get("Category", "General"),
                    "duration_summary": duration_summary
                })
        else:
            if bars:
                rows.append({
                    "label": label,
                    "bars": bars,
                    "category": record.get("Category", "General"),
                    "duration_summary": duration_summary
                })
                # Keep track of min/max if bounds are not restricted
                for bar in bars:
                    bar_end = bar["start"] + datetime.timedelta(days=bar["duration"])
                    if min_date is None or bar["start"] < min_date:
                        min_date = bar["start"]
                    if max_date is None or bar_end > max_date:
                        max_date = bar_end

    if not rows:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, f"No active timeline data in this range for Project {project_name}" if project_name else "No active timeline data to display",
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
    gs = fig.add_gridspec(1, 2, width_ratios=[0.38, 0.62], wspace=0.02)
    ax_left = fig.add_subplot(gs[0, 0])
    ax = fig.add_subplot(gs[0, 1])

    ax.set_facecolor(bg_color)
    ax_left.set_facecolor(sidebar_bg)

    # Configure left axis (labels sidebar)
    ax_left.set_ylim(-0.5, len(rows) - 0.5)
    ax_left.invert_yaxis()
    ax_left.xaxis.set_visible(False)
    ax_left.yaxis.set_visible(False)
    ax_left.set_xlim(0, 1)

    # Draw category groups on the left sidebar
    categories = [r.get('category', '') for r in rows]
    group_positions = {}
    for idx, cat in enumerate(categories):
        group_positions.setdefault(cat, []).append(idx)

    for cat, indices in group_positions.items():
        mid = (indices[0] + indices[-1]) / 2
        ax_left.text(0.02, mid, str(cat), va='center', ha='left', fontsize=11, color=fg_text, weight='bold')

    # Print each row's short method label and duration text
    for y, row in enumerate(rows):
        short = row['label'].split('\n')[0]
        ax_left.text(0.02, y + 0.28, short, va='center', ha='left', fontsize=8, color=fg_sub)
        
        dur_summary = row.get("duration_summary", "")
        ax_left.text(0.52, y, dur_summary, va='center', ha='left', fontsize=8, color=fg_text, weight='semibold')

    # Right axis: Gantt chart area
    if start_date_bound and end_date_bound:
        start_num = mdates.date2num(start_date_bound)
        end_num = mdates.date2num(end_date_bound)
    else:
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

    # Month labels above the gantt area
    if timeline_type in ['days', 'weeks']:
        month_iter = datetime.date(min_date.year, min_date.month, 1)
        top_y = -0.8
        while month_iter <= max_date.date():
            month_start = datetime.datetime(month_iter.year, month_iter.month, 1)
            next_month = (month_start.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
            month_mid = month_start + (next_month - month_start) / 2
            ax.text(mdates.date2num(month_mid), top_y, month_start.strftime('%B'), ha='center', va='bottom', color=fg_text, fontsize=11, fontweight='bold')
            month_iter = next_month.date()

    # Draw rounded bars (without labels inside)
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
            edge = Rectangle((x, y_pos), width, height, linewidth=0.6, edgecolor='#083a4a' if theme == 'dark' else '#cbd5e1', facecolor='none')
            ax.add_patch(edge)

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
# --- PDF CONVERSION HELPER ---
def convert_html_to_pdf(html_path, pdf_path):
    """Converts local HTML file to PDF by calling an installed Chromium-based browser (Edge/Chrome/Chromium) headlessly."""
    import subprocess
    import shutil
    import os
    import platform
    
    browser_path = None
    system = platform.system()
    
    if system == "Windows":
        # Check standard installation locations for Edge and Chrome
        paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ]
        for p in paths:
            if os.path.exists(p):
                browser_path = p
                break
    elif system == "Darwin": # macOS
        paths = [
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
        ]
        for p in paths:
            if os.path.exists(p):
                browser_path = p
                break
    else: # Linux
        # Search PATH
        for name in ["google-chrome", "chrome", "chromium", "chromium-browser", "microsoft-edge", "edge"]:
            path = shutil.which(name)
            if path:
                browser_path = path
                break
                
    if not browser_path:
        raise RuntimeError("No Chromium-based browser (Microsoft Edge, Google Chrome, or Chromium) was detected on this system. Please make sure one is installed.")

    # Create an isolated temp user-data-dir so headless Edge doesn't try to
    # reuse or lock-conflict with an existing running browser session on Windows.
    import tempfile
    user_data_dir = tempfile.mkdtemp(prefix="edge_headless_")

    file_url = "file:///" + os.path.abspath(html_path).replace("\\", "/")

    args = [
        browser_path,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--no-pdf-header-footer",
        "--disable-extensions",
        "--disable-background-networking",
        "--disable-default-apps",
        "--disable-sync",
        "--disable-translate",
        "--hide-scrollbars",
        "--metrics-recording-only",
        "--mute-audio",
        "--no-first-run",
        f"--user-data-dir={user_data_dir}",
        f"--print-to-pdf={pdf_path}",
        file_url
    ]
    
    startupinfo = None
    if system == "Windows":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
    try:
        result = subprocess.run(
            args, startupinfo=startupinfo,
            capture_output=True, text=True,
            timeout=120
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("Headless browser timed out after 120s. This may be caused by a locked browser profile or missing font/resource. Try closing all Edge/Chrome windows and retrying.")
    finally:
        # Clean up the isolated temp profile directory
        try:
            import shutil as _shutil
            _shutil.rmtree(user_data_dir, ignore_errors=True)
        except Exception:
            pass

    if result.returncode != 0:
        raise RuntimeError(f"Headless browser PDF conversion failed (exit {result.returncode}): {result.stderr[:600]}")
    if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
        raise RuntimeError("Headless browser ran but produced no PDF output. Check that the HTML template is valid.")
    return True


def generate_pdf_fallback(filename, project_name=None, comment=None, records_dict=None, timeline_type='weeks'):
    """Create a simple ReportLab PDF when the browser-based renderer is unavailable."""
    if records_dict is None:
        records_dict = ExcelDataStore.get_project_rows(project_name) if project_name else []

    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Title'], fontSize=16, leading=20, spaceAfter=10)
    heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontSize=12, leading=14, spaceAfter=6, textColor=colors.HexColor('#0f172a'))
    body_style = ParagraphStyle('Body', parent=styles['BodyText'], fontSize=9, leading=11)
    small_style = ParagraphStyle('Small', parent=styles['BodyText'], fontSize=7.5, leading=9.5, textColor=colors.HexColor('#475569'))

    story = []
    story.append(Paragraph(f"Project Status Report - {project_name or 'All Projects'}", title_style))
    story.append(Paragraph(f"Generated on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", small_style))
    story.append(Spacer(1, 10))

    if comment:
        story.append(Paragraph('Report Comment', heading_style))
        story.append(Paragraph(comment, body_style))
        story.append(Spacer(1, 8))

    if records_dict:
        # Render a clean Gantt image using the Matplotlib helper and embed it in the PDF
        story.append(Paragraph('Timeline Summary', heading_style))
        import tempfile as _tmp
        png_tmp = _tmp.NamedTemporaryFile(delete=False, suffix='.png')
        png_tmp.close()
        # generate image with weekly scale and light theme to match PDF
        try:
            generate_gantt_chart_image(records_dict, png_tmp.name, project_name=project_name, theme='light', timeline_type='weeks')
        except Exception:
            # fallback to a small text placeholder image
            try:
                generate_gantt_chart_image([], png_tmp.name, project_name=project_name, theme='light', timeline_type='weeks')
            except Exception:
                pass
        # embed image full-width (keep file until doc.build completes)
        img = Image(png_tmp.name, width=doc.width, height=doc.width * 0.45)
        story.append(Spacer(1, 8))
        story.append(img)
        story.append(Spacer(1, 8))
    else:
        story.append(Paragraph('No project data available.', body_style))

    doc.build(story)
    # cleanup image temp file now that PDF has been composed
    try:
        if os.path.exists(png_tmp.name):
            os.remove(png_tmp.name)
    except Exception:
        pass
    return True


def normalize_timeline_type(scale):
    if not scale:
        return 'weeks'
    scale = str(scale).strip().lower()
    if scale in ('day', 'days'):
        return 'days'
    if scale in ('week', 'weeks'):
        return 'weeks'
    if scale in ('month', 'months'):
        return 'months'
    if scale in ('year', 'years'):
        return 'years'
    return 'weeks'


def filter_records_by_category(records, category_filter):
    if not category_filter:
        return records
    category_filter = str(category_filter).strip().lower()
    if category_filter in ('all', '', 'none'):
        return records
    return [r for r in records if str(r.get('Category', '')).strip().lower() == category_filter]


def generate_email_summary_pdf(filename, project_name=None, comment=None):
    """Generate a compact summary PDF suitable for email attachment fallback."""
    records = ExcelDataStore.get_project_rows(project_name) if project_name else []
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Title'], fontSize=16, leading=20, spaceAfter=10)
    heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontSize=12, leading=14, spaceAfter=6)
    body_style = ParagraphStyle('Body', parent=styles['BodyText'], fontSize=9, leading=12)
    small_style = ParagraphStyle('Small', parent=styles['BodyText'], fontSize=8, leading=10, textColor=colors.HexColor('#475569'))

    total_records = len(records)
    total_defects = sum(int(r.get('Defect Qty', 0) or 0) for r in records)
    total_comments = sum(1 for r in records if str(r.get('Comments', '')).strip())

    story = []
    story.append(Paragraph(f"Project Summary Report - {project_name or 'All Projects'}", title_style))
    story.append(Paragraph(f"Generated on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", small_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph('This attached summary report has been generated as a compact email attachment because the full project report exceeded the allowed attachment size.', body_style))
    story.append(Spacer(1, 10))

    summary_items = [
        ('Project', project_name or 'All Projects'),
        ('Total Task Records', str(total_records)),
        ('Total Defects', str(total_defects)),
        ('Total Comment Entries', str(total_comments)),
    ]

    for label, value in summary_items:
        story.append(Paragraph(f'<b>{label}:</b> {value}', body_style))
    story.append(Spacer(1, 10))

    if comment:
        story.append(Paragraph('Report Comment:', heading_style))
        story.append(Paragraph(comment, body_style))
        story.append(Spacer(1, 10))

    if records:
        table_data = [['Test Number', 'Status', 'Defect Qty', 'Comments']]
        for record in records[:10]:
            test_number = record.get('Test Number', '')
            status = record.get('Status', '') or record.get('Current Status', '') or 'N/A'
            defect_qty = str(record.get('Defect Qty', '') or '')
            comments = str(record.get('Comments', '') or '')[:80]
            table_data.append([test_number, status, defect_qty, comments])

        table = Table(table_data, colWidths=[100, 110, 70, 240])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f2937')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('ALIGN', (2, 0), (2, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#d1d5db')),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(table)
    else:
        story.append(Paragraph('No task records were found for the selected project.', body_style))

    doc.build(story)
    return True


# --- PDF REPORT GENERATION ---
def generate_pdf_report(filename, project_name=None, comment=None, progress_callback=None, category_filter=None, timeline_type='weeks'):
    """Generates a styled Test Operations PDF report by rendering a Jinja HTML template and using Edge/Chrome headless to print to PDF."""
    if progress_callback:
        progress_callback(10, "Analyzing project data...")

    timeline_type = normalize_timeline_type(timeline_type)
    if project_name:
        records_dict = ExcelDataStore.get_project_rows(project_name)
    else:
        records_dict = []
        for p in ExcelDataStore.get_projects():
            records_dict.extend(ExcelDataStore.get_project_rows(p["name"]))

    records_dict = filter_records_by_category(records_dict, category_filter)

    # Calculate global date boundaries
    min_date = None
    max_date = None
    for r in records_dict:
        start_date = parse_date(r.get("Start Date"))
        if not start_date:
            continue
        
        phases = [
            ("Proto", int(r.get("Proto Weeks", 0) or 0), int(r.get("Proto Days", 0) or 0)),
            ("DVT", int(r.get("DVT Weeks", 0) or 0), int(r.get("DVT Days", 0) or 0)),
            ("EVT", int(r.get("EVT Weeks", 0) or 0), int(r.get("EVT Days", 0) or 0)),
            ("PVT", int(r.get("PVT Weeks", 0) or 0), int(r.get("PVT Days", 0) or 0))
        ]
        
        if min_date is None or start_date < min_date:
            min_date = start_date
            
        phase_start = start_date
        for phase_name, w, d in phases:
            dur = w * 7 + d
            if dur > 0:
                phase_end = phase_start + datetime.timedelta(days=dur)
                if max_date is None or phase_end > max_date:
                    max_date = phase_end
                phase_start = phase_end
            else:
                if max_date is None or phase_start > max_date:
                    max_date = phase_start

    # Mathematical timeline slicing into 4-week chunks (28 days)
    timeline_chunks = []
    phase_colors = {
        "Proto": "#8E44AD",
        "DVT": "#2980B9",
        "EVT": "#27AE60",
        "PVT": "#D35400"
    }
    
    if progress_callback:
        progress_callback(30, "Calculating visual timeline chunks...")

    chunk_length_days = 28
    if timeline_type == 'days':
        chunk_length_days = 7
    elif timeline_type == 'weeks':
        chunk_length_days = 28
    elif timeline_type == 'months':
        chunk_length_days = 90
    elif timeline_type == 'years':
        chunk_length_days = 365

    scale_labels = {
        'days': 'Daily Timeline',
        'weeks': 'Weekly Timeline',
        'months': 'Monthly Timeline',
        'years': 'Yearly Timeline'
    }
    timeline_label = scale_labels.get(timeline_type, 'Timeline')
    slice_label = {
        'days': '7-Day Slice',
        'weeks': '4-Week Slice',
        'months': 'Quarter Slice',
        'years': 'Annual Slice'
    }.get(timeline_type, 'Timeline Slice')

    # Calculate dynamic scale timeline headers
    timeline_headers = []

    if min_date and max_date:
        # Convert to date objects for clean boundary adjustments
        min_d = min_date.date() if isinstance(min_date, datetime.datetime) else min_date
        max_d = max_date.date() if isinstance(max_date, datetime.datetime) else max_date

        if timeline_type == 'days':
            pass
        elif timeline_type == 'months':
            min_d = datetime.date(min_d.year, min_d.month, 1)
            if max_d.month == 12:
                max_d = datetime.date(max_d.year, 12, 31)
            else:
                next_month = datetime.date(max_d.year, max_d.month + 1, 1)
                max_d = next_month - datetime.timedelta(days=1)
        elif timeline_type == 'years':
            min_d = datetime.date(min_d.year, 1, 1)
            max_d = datetime.date(max_d.year, 12, 31)
        else: # 'weeks'
            span = (max_d - min_d).days + 1
            rem = span % 7
            if rem > 0:
                max_d = max_d + datetime.timedelta(days=(7 - rem))

        # Re-assign back to min_date and max_date as datetime objects combined with time.min
        min_date = datetime.datetime.combine(min_d, datetime.time.min)
        max_date = datetime.datetime.combine(max_d, datetime.time.min)

        total_span_days = max((max_date - min_date).days, 1)
        # full timeline in weeks for header numbering
        total_weeks = int((total_span_days + 6) // 7)

        # Build timeline headers list with exact width percentages based on duration of each unit in total_span_days
        min_date_val = min_date.date()
        max_date_val = max_date.date()
        
        if timeline_type == 'days':
            width_pct = (1.0 / float(total_span_days)) * 100.0
            for i in range(1, total_span_days + 1):
                label = ""
                if total_span_days <= 35:
                    label = str(i)
                else:
                    if i == 1 or i % 5 == 0:
                        label = str(i)
                timeline_headers.append({
                    "label": label,
                    "width_pct": round(width_pct, 4)
                })
        elif timeline_type == 'months':
            curr = min_date_val
            while curr <= max_date_val:
                month_start = max(curr, min_date_val)
                if curr.month == 12:
                    next_month_start = datetime.date(curr.year + 1, 1, 1)
                else:
                    next_month_start = datetime.date(curr.year, curr.month + 1, 1)
                month_end = min(next_month_start, max_date_val + datetime.timedelta(days=1))
                
                days_in_month = (month_end - month_start).days
                if days_in_month > 0:
                    width_pct = (days_in_month / float(total_span_days)) * 100.0
                    timeline_headers.append({
                        "label": curr.strftime('%b'),
                        "width_pct": round(width_pct, 4)
                    })
                curr = next_month_start
        elif timeline_type == 'years':
            curr_year = min_date_val.year
            while curr_year <= max_date_val.year:
                year_start = max(datetime.date(curr_year, 1, 1), min_date_val)
                year_end = min(datetime.date(curr_year + 1, 1, 1), max_date_val + datetime.timedelta(days=1))
                
                days_in_year = (year_end - year_start).days
                if days_in_year > 0:
                    width_pct = (days_in_year / float(total_span_days)) * 100.0
                    timeline_headers.append({
                        "label": str(curr_year),
                        "width_pct": round(width_pct, 4)
                    })
                curr_year += 1
        else: # 'weeks'
            width_pct = (7.0 / float(total_span_days)) * 100.0
            total_weeks_val = total_span_days // 7
            for i in range(1, total_weeks_val + 1):
                timeline_headers.append({
                    "label": str(i),
                    "width_pct": round(width_pct, 4)
                })

        chunk_rows = []
        for r in records_dict:
            start_date = parse_date(r.get("Start Date"))
            if not start_date:
                continue
            
            phases = [
                ("Proto", int(r.get("Proto Weeks", 0) or 0), int(r.get("Proto Days", 0) or 0), int(r.get("Proto Qty", 0) or 0)),
                ("DVT", int(r.get("DVT Weeks", 0) or 0), int(r.get("DVT Days", 0) or 0), int(r.get("DVT Qty", 0) or 0)),
                ("EVT", int(r.get("EVT Weeks", 0) or 0), int(r.get("EVT Days", 0) or 0), int(r.get("EVT Qty", 0) or 0)),
                ("PVT", int(r.get("PVT Weeks", 0) or 0), int(r.get("PVT Days", 0) or 0), int(r.get("PVT Qty", 0) or 0))
            ]
            
            row_bars = []
            phase_start = start_date
            for phase_name, w, d, qty in phases:
                dur = w * 7 + d
                if dur <= 0:
                    continue
                phase_end = phase_start + datetime.timedelta(days=dur)

                clamped_start = max(phase_start, min_date)
                clamped_end = min(phase_end, max_date)
                clamped_dur = (clamped_end - clamped_start).days

                if clamped_dur > 0:
                    offset_days = (clamped_start - min_date).days
                    full_left = ((clamped_start - min_date).days / float(total_span_days)) * 100.0
                    full_width = (clamped_dur / float(total_span_days)) * 100.0

                    weeks_val = clamped_dur // 7
                    days_val = clamped_dur % 7
                    dur_text = f"{weeks_val}w" if weeks_val > 0 else ""
                    if days_val > 0:
                        dur_text += f"{days_val}d"
                    if not dur_text:
                        dur_text = "0d"

                    row_bars.append({
                        "phase": phase_name,
                        "color": phase_colors.get(phase_name, "#64748b"),
                        "left_percent_full": round(full_left, 2),
                        "width_percent_full": round(full_width, 2),
                        "duration_text": dur_text,
                        "qty": qty,
                        "start_date": clamped_start.strftime("%Y-%m-%d"),
                        "end_date": clamped_end.strftime("%Y-%m-%d")
                    })

                phase_start = phase_end
                
            if row_bars:
                chunk_rows.append({
                    "method": r.get("Test Method", "Untitled"),
                    "number": r.get("Test Number", "Unknown"),
                    "category": r.get("Category", "General"),
                    "bars": row_bars
                })

        if chunk_rows:
            timeline_chunks.append({
                "start_str": min_date.strftime("%Y-%m-%d"),
                "end_str": max_date.strftime("%Y-%m-%d"),
                "rows": chunk_rows
            })

    if progress_callback:
        progress_callback(50, "Preparing executive summary...")

    category_summary = {}
    for r in records_dict:
        cat = r.get("Category", "General")
        category_summary.setdefault(cat, {
            "tasks_count": 0,
            "defects_count": 0,
            "total_qty": 0
        })
        category_summary[cat]["tasks_count"] += 1
        category_summary[cat]["defects_count"] += int(r.get("Defect Qty", 0) or 0)
        category_summary[cat]["total_qty"] += sum(
            int(r.get(f"{p} Qty", 0) or 0) for p in ["Proto", "DVT", "EVT", "PVT"]
        )

    records_by_cat = {}
    for r in records_dict:
        cat = r.get("Category", "General")
        records_by_cat.setdefault(cat, []).append(r)

    rejected_records = [
        r for r in records_dict if str(r.get("Comments", "")).strip() != "" or int(r.get("Defect Qty", 0)) > 0
    ]
    rej_by_cat = {}
    for r in rejected_records:
        cat = r.get("Category", "General")
        rej_by_cat.setdefault(cat, []).append(r)

    excel_filename = "tasks.xlsx"
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

    context = {
        "title": f"Project Status Report - Project {project_name}" if project_name else "Daily Test Operations Report",
        "project_name": project_name or "All Active Projects",
        "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "excel_filename": excel_filename,
        "total_records": len(records_dict),
        "total_defects": sum(int(r.get('Defect Qty', 0) or 0) for r in records_dict),
        "total_comments": sum(1 for r in records_dict if str(r.get('Comments', '')).strip()),
        "category_summary": category_summary,
        "comment": comment,
        "timeline_chunks": timeline_chunks,
        "records_by_cat": records_by_cat,
        "rejected_records": rejected_records,
        "rej_by_cat": rej_by_cat,
        "selected_category": category_filter if category_filter and str(category_filter).strip().lower() != 'all' else 'All Categories',
        "timeline_type": timeline_type,
        "timeline_label": timeline_label,
        "timeline_slice_label": slice_label,
        "timeline_headers": timeline_headers
    }
    # Add full timeline meta for template rendering
    try:
        if min_date and max_date:
            # inclusive day count
            total_span_days = max((max_date - min_date).days + 1, 1)
            total_weeks = int(((total_span_days) + 6) // 7)
        else:
            total_span_days = 1
            total_weeks = 1
    except Exception:
        total_span_days = 1
        total_weeks = 1

    # Build per-day labels for the timeline footer (e.g. 'Jun 28', 'Jun 29')
    day_labels = []
    try:
        if min_date and max_date:
            day_count = (max_date - min_date).days + 1
            for i in range(day_count):
                d = (min_date + datetime.timedelta(days=i))
                # Short format 'Mon dd' or 'Jun 28'
                day_labels.append(d.strftime('%b %d'))
    except Exception:
        day_labels = []

    context.update({
        "full_timeline_weeks": total_weeks,
        "full_timeline_start": min_date.strftime("%Y-%m-%d") if min_date else None,
        "full_timeline_end": max_date.strftime("%Y-%m-%d") if max_date else None,
        "full_timeline_days": total_span_days,
        "full_timeline_day_labels": day_labels
    })

    if progress_callback:
        progress_callback(75, "Rendering HTML report...")

    with app.app_context():
        html_content = render_template("pdf_report.html", **context)
    temp_html_path = filename.replace(".pdf", ".html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    if progress_callback:
        progress_callback(90, "Compiling PDF document via headless browser...")

    try:
        try:
            convert_html_to_pdf(temp_html_path, filename)
        except Exception as browser_error:
            if progress_callback:
                progress_callback(90, "Browser PDF renderer unavailable; using fallback export...")
            generate_pdf_fallback(filename, project_name=project_name, comment=comment, records_dict=records_dict, timeline_type=timeline_type)
        if progress_callback:
            progress_callback(100, "Completed!")
    finally:
        try:
            if os.path.exists(temp_html_path):
                os.remove(temp_html_path)
        except Exception:
            pass


# --- CONSOLIDATED PDF REPORT GENERATION ---
def generate_consolidated_pdf_report(filename, progress_callback=None, category_filter=None, timeline_type='weeks'):
    """Generates a consolidated PDF for ALL active projects by rendering each project
    separately and merging the resulting PDFs with PyPDF2."""
    projects = ExcelDataStore.get_projects()
    if not projects:
        raise RuntimeError("No projects found. Please import data first.")

    import tempfile as _tmp_mod
    per_project_pdfs = []
    n = len(projects)

    for i, proj in enumerate(projects):
        pct_start = 10 + int(i / n * 70)
        pct_end   = 10 + int((i + 1) / n * 70)
        pname = proj["name"]
        if progress_callback:
            progress_callback(pct_start, f"Generating report for project {pname}...")

        tmp_pdf = _tmp_mod.NamedTemporaryFile(delete=False, suffix=".pdf", prefix=f"proj_{pname}_")
        tmp_pdf.close()

        def _sub_progress(pct, msg, ps=pct_start, pe=pct_end):
            if progress_callback:
                mapped = ps + int((pct / 100) * (pe - ps))
                progress_callback(mapped, msg)

        try:
            generate_pdf_report(tmp_pdf.name, project_name=pname, progress_callback=_sub_progress, category_filter=category_filter, timeline_type=timeline_type)
            per_project_pdfs.append(tmp_pdf.name)
        except Exception as e:
            # Skip failed projects but log the error
            print(f"[WARN] Skipped project {pname} during consolidated export: {e}")
            try:
                os.remove(tmp_pdf.name)
            except Exception:
                pass

    if not per_project_pdfs:
        raise RuntimeError("All project PDF generations failed. Check the server logs for details.")

    if progress_callback:
        progress_callback(82, "Merging project reports...")

    # Merge all per-project PDFs into the final output
    try:
        from PyPDF2 import PdfMerger
        merger = PdfMerger()
        for pdf_path in per_project_pdfs:
            merger.append(pdf_path)
        with open(filename, "wb") as out_f:
            merger.write(out_f)
        merger.close()
    except ImportError:
        # PyPDF2 not installed — fall back to concatenating raw bytes
        # (works for simple PDFs but may produce a malformed cross-reference table)
        import shutil as _sh2
        if len(per_project_pdfs) == 1:
            _sh2.copy2(per_project_pdfs[0], filename)
        else:
            with open(filename, "wb") as out_f:
                for pdf_path in per_project_pdfs:
                    with open(pdf_path, "rb") as in_f:
                        out_f.write(in_f.read())
    finally:
        for pdf_path in per_project_pdfs:
            try:
                os.remove(pdf_path)
            except Exception:
                pass

    if progress_callback:
        progress_callback(100, "Completed!")



@app.route('/generate-report', methods=['GET'])
@login_required
def generate_report():
    project_name = request.args.get('project')
    comment = request.args.get('comment')
    category_filter = request.args.get('category')
    timeline_type = request.args.get('scale', request.args.get('timeline', 'weeks'))
    pdf_filename = os.path.join(UPLOAD_FOLDER, "Daily_Operations_Report.pdf")
    try:
        generate_pdf_report(pdf_filename, project_name, comment, category_filter=category_filter, timeline_type=timeline_type)
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
    data = request.json or {}
    category_filter = data.get('category')
    timeline_type = data.get('scale') or data.get('timeline')

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
            generate_consolidated_pdf_report(temp_pdf, progress_callback=update_progress, category_filter=category_filter, timeline_type=timeline_type)
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


@app.route('/api/export-project/start', methods=['POST'])
@login_required
def start_project_export():
    data = request.json or {}
    project_name = data.get('project')
    comment = data.get('comment')
    category_filter = data.get('category')
    timeline_type = data.get('scale') or data.get('timeline')
    
    task_id = str(uuid.uuid4())
    export_tasks[task_id] = {
        "status": "processing",
        "progress": 0,
        "message": "Initializing project PDF generation..."
    }
    
    def run_compilation():
        temp_pdf = os.path.join(UPLOAD_FOLDER, f"project_{task_id}.pdf")
        
        def update_progress(percent, msg):
            if task_id in export_tasks:
                export_tasks[task_id]["progress"] = percent
                export_tasks[task_id]["message"] = msg
            
        try:
            generate_pdf_report(temp_pdf, project_name, comment, progress_callback=update_progress, category_filter=category_filter, timeline_type=timeline_type)
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

@app.route('/api/export-project/progress/<task_id>', methods=['GET'])
@login_required
def project_export_progress(task_id):
    state = export_tasks.get(task_id)
    if not state:
        return jsonify({"status": "failed", "progress": 0, "message": "Task not found"}), 404
    return jsonify(state)

@app.route('/api/export-project/download/<task_id>/<filename>', methods=['GET'])
@login_required
def project_export_download(task_id, filename):
    temp_pdf = os.path.join(UPLOAD_FOLDER, f"project_{task_id}.pdf")
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
        response.headers['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    except Exception as e:
        return jsonify({"error": f"Failed to stream PDF: {str(e)}"}), 500

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

# --- EXCEL FULL EXPORT ---
@app.route('/api/export-excel', methods=['GET'])
@login_required
def export_excel():
    """Downloads the current tasks.xlsx data file as a full Excel report."""
    if not os.path.exists(EXCEL_PATH):
        return jsonify({"error": "No Excel data file found. Please import data first."}), 404
    try:
        import shutil as _sh
        import tempfile as _tmp
        # Copy to a temp file so we can stream it safely even if it's being written
        tmp = _tmp.NamedTemporaryFile(delete=False, suffix='.xlsx')
        tmp.close()
        _sh.copy2(EXCEL_PATH, tmp.name)
        today = datetime.date.today().strftime('%Y-%m-%d')
        return send_file(
            tmp.name,
            as_attachment=True,
            download_name=f"Full_Operations_Report_{today}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({"error": f"Failed to export Excel file: {str(e)}"}), 500

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

    # Step 1: Generate the PDF report
    pdf_filename = os.path.join(UPLOAD_FOLDER, "email_report.pdf")
    try:
        generate_pdf_report(pdf_filename, project_name)
    except Exception as e:
        return jsonify({"error": f"Failed to generate PDF for email: {str(e)}"}), 500

    # Step 2: Resolve attachment — compact summary if PDF exceeds size limit
    attachment_size_limit_mb = int(os.getenv('EMAIL_ATTACHMENT_MAX_MB', EMAIL_ATTACHMENT_MAX_MB))
    attachment_size_limit_bytes = attachment_size_limit_mb * 1024 * 1024
    email_attachment_path = pdf_filename
    attachment_warning = None

    if os.path.getsize(pdf_filename) > attachment_size_limit_bytes:
        compact_pdf_filename = os.path.join(UPLOAD_FOLDER, "email_summary_report.pdf")
        generate_email_summary_pdf(compact_pdf_filename, project_name)
        email_attachment_path = compact_pdf_filename
        attachment_warning = (
            f"The full report exceeded the {attachment_size_limit_mb} MB attachment limit. "
            "A compact summary is attached instead."
        )
        body = f"{body}\n\n{attachment_warning}"

    pdf_download_name = f"Gantt_Report_{project_name}.pdf" if project_name else "Operations_Report.pdf"
    if email_attachment_path != pdf_filename:
        pdf_download_name = f"Summary_Report_{project_name or 'Operations'}.pdf"

    # Step 3: Send via Microsoft Graph API (Office 365)
    tenant_id     = os.getenv('TENANT_ID')
    client_id     = os.getenv('SHAREPOINT_CLIENT_ID')
    client_secret = os.getenv('SHAREPOINT_CLIENT_SECRET')
    sender_email  = os.getenv('MAIL_DEFAULT_SENDER') or os.getenv('MAIL_USERNAME')

    if tenant_id and client_id and client_secret and sender_email:
        try:
            import msal
            import requests as _requests
            import base64

            # Acquire OAuth2 access token using client credentials (same pattern as SharePoint)
            authority = f"https://login.microsoftonline.com/{tenant_id}"
            msal_app = msal.ConfidentialClientApplication(
                client_id,
                authority=authority,
                client_credential=client_secret,
            )
            token_result = msal_app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )

            if "access_token" not in token_result:
                error_desc = token_result.get(
                    "error_description", token_result.get("error", "Unknown MSAL error")
                )
                raise RuntimeError(f"Failed to acquire Microsoft Graph token: {error_desc}")

            access_token = token_result["access_token"]

            # Build the sendMail JSON payload with the PDF as a base64 attachment
            with open(email_attachment_path, "rb") as f:
                pdf_b64 = base64.b64encode(f.read()).decode("utf-8")

            mail_payload = {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "Text",
                        "content": body
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": recipient}}
                    ],
                    "attachments": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": pdf_download_name,
                            "contentType": "application/pdf",
                            "contentBytes": pdf_b64
                        }
                    ]
                },
                "saveToSentItems": "true"
            }

            # POST to the Graph API sendMail endpoint
            graph_url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            }
            resp = _requests.post(graph_url, headers=headers, json=mail_payload, timeout=30)

            if resp.status_code == 202:
                success_message = f"Email with PDF report successfully sent to {recipient} via Microsoft 365!"
                if attachment_warning:
                    success_message = (
                        f"Email sent to {recipient} via Microsoft 365 with compact summary attachment "
                        f"because the full report exceeded the {attachment_size_limit_mb} MB limit."
                    )
                return jsonify({
                    "success": True,
                    "mode": "microsoft_graph",
                    "message": success_message,
                    "attachment_mode": "compact_summary" if attachment_warning else "full_pdf",
                    "attachment_warning": attachment_warning
                })
            else:
                # Graph returned a non-202 response — surface the error detail
                try:
                    err_detail = resp.json().get("error", {}).get("message", resp.text[:400])
                except Exception:
                    err_detail = resp.text[:400]
                raise RuntimeError(f"Graph API returned HTTP {resp.status_code}: {err_detail}")

        except Exception as e:
            print(f"Microsoft Graph Email Error: {type(e).__name__}: {e}")
            return jsonify({
                "success": False,
                "mode": "microsoft_graph",
                "message": f"Failed to send email via Microsoft Graph API: {type(e).__name__}: {str(e)}",
                "error": str(e)
            }), 500

    # Step 4: Simulation mode — Graph credentials not fully configured in .env
    return jsonify({
        "success": True,
        "mode": "simulation",
        "message": (
            f"Email simulated (Microsoft Graph credentials not fully configured). "
            f"PDF report generated and would be sent to {recipient}."
        ),
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
    
