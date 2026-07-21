import os
import unittest
import datetime
import shutil
from unittest.mock import patch
import pandas as pd
import app as app_module
from app import app, ExcelDataStore, generate_pdf_report, UPLOAD_FOLDER, clean_dataframe_headers

class TestOpsDashboardBackend(unittest.TestCase):
    
    def setUp(self):
        # Point the app to a temporary test Excel file
        self.test_excel_path = os.path.join(UPLOAD_FOLDER, 'test_tasks.xlsx')
        app_module.EXCEL_PATH = self.test_excel_path
        
        # Point the app to a temporary sync config file
        self.test_sync_config_path = os.path.join(UPLOAD_FOLDER, 'test_sync_config.json')
        app_module.SYNC_CONFIG_PATH = self.test_sync_config_path
        
        # Configure app for testing
        app.config['TESTING'] = True
        self.app = app.test_client()
        
        # Push application context
        self.app_context = app.app_context()
        self.app_context.push()
        
        # Start with a clean test Excel file containing a sample sheet
        if os.path.exists(self.test_excel_path):
            os.remove(self.test_excel_path)
            
        # Seed test-specific data in the Excel file
        df = pd.DataFrame([
            {
                "Category": "Device: Hair Dryer",
                "Test Method": "Stress Imaging",
                "Test Number": "TM-035128",
                "Start Date": "2026-06-28",
                "Proto Weeks": 3,
                "Proto Days": 2,
                "Proto Qty": 15,
                "DVT Weeks": 0,
                "DVT Days": 0,
                "DVT Qty": 0,
                "EVT Weeks": 0,
                "EVT Days": 0,
                "EVT Qty": 0,
                "PVT Weeks": 0,
                "PVT Days": 0,
                "PVT Qty": 0,
                "Defect Qty": 0,
                "Comments": ""
            }
        ])
        with pd.ExcelWriter(self.test_excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='893')

    def tearDown(self):
        # Clean up test files and pop context
        if os.path.exists(self.test_excel_path):
            try:
                os.remove(self.test_excel_path)
            except Exception:
                pass
        if os.path.exists(self.test_sync_config_path):
            try:
                os.remove(self.test_sync_config_path)
            except Exception:
                pass
        self.app_context.pop()
  
    def test_excel_initialization(self):
        """Test that the Excel file is correctly initialized and contains our seed sheet."""
        projects = ExcelDataStore.get_projects()
        self.assertEqual(len(projects), 1)
        self.assertEqual(projects[0]["name"], "893")
        
        records = ExcelDataStore.get_project_rows("893")
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Project Name"], "893")
        self.assertEqual(records[0]["Category"], "Device: Hair Dryer")
        self.assertEqual(records[0]["Test Method"], "Stress Imaging")
        self.assertEqual(records[0]["Proto Qty"], 15)
  
    def test_create_project(self):
        """Test adding a new project (sheet) to the Excel file."""
        success = ExcelDataStore.create_project("TestProj")
        self.assertTrue(success)
        
        projects = ExcelDataStore.get_projects()
        self.assertEqual(len(projects), 2)
        self.assertTrue(any(p["name"] == "TestProj" for p in projects))
  
    def test_add_and_get_record(self):
        """Test adding a new test record to a sheet and verifying its fields."""
        data = {
            "Category": "Device: Fan",
            "Test Method": "Airflow Speed",
            "Test Number": "TM-999",
            "Start Date": "2026-07-15",
            "Proto Weeks": 1,
            "Proto Days": 2,
            "Proto Qty": 10,
            "Comments": "Loud motor noise"
        }
        success = ExcelDataStore.save_project_row("893", None, data)
        self.assertTrue(success)
        
        rows = ExcelDataStore.get_project_rows("893")
        self.assertEqual(len(rows), 2)
        
        saved = rows[1]
        self.assertEqual(saved["Category"], "Device: Fan")
        self.assertEqual(saved["Proto Qty"], 10)
        self.assertEqual(saved["Comments"], "Loud motor noise")
  
    def test_update_record(self):
        """Test updating an existing test record in the Excel file."""
        rows = ExcelDataStore.get_project_rows("893")
        self.assertEqual(len(rows), 1)
        
        # Update fields for the row at index 0
        data = rows[0]
        data["Proto Qty"] = 20
        data["Comments"] = "Updated comments"
        
        success = ExcelDataStore.save_project_row("893", 0, data)
        self.assertTrue(success)
        
        updated_rows = ExcelDataStore.get_project_rows("893")
        self.assertEqual(updated_rows[0]["Proto Qty"], 20)
        self.assertEqual(updated_rows[0]["Comments"], "Updated comments")

    def test_clean_dataframe_headers_promotes_repeated_date_headers(self):
        """Test that repeated generic Date headers are promoted and flattened correctly."""
        raw = pd.DataFrame(
            [
                ['Category', 'Test Method', 'Test Number', 'Proto1', 'Date (Week)', 'Date (Day)', 'DVT1', 'Date (Week)', 'Date (Day)'],
                ['Device (Hair Dryer)', 'Tress Imaging', 'TM-003128', 15, 11, 2, 15, 13, 10],
                ['Device (Hair Dryer)', 'Drop Test', 'TM-003440', 15, 11, 2, 15, 13, 10],
            ],
            columns=['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
        )

        cleaned = clean_dataframe_headers(raw)
        expected_columns = [
            'Category', 'Test Method', 'Test Number', 'Proto1',
            'Proto1_Date (Week)', 'Proto1_Date (Day)',
            'DVT1', 'DVT1_Date (Week)', 'DVT1_Date (Day)'
        ]
        self.assertEqual(cleaned.columns.tolist(), expected_columns)
        self.assertEqual(int(cleaned.iloc[0]['Proto1_Date (Week)']), 11)
        self.assertEqual(int(cleaned.iloc[0]['DVT1_Date (Day)']), 10)
   
    def test_delete_record(self):
        """Test deleting a record from the Excel sheet."""
        rows = ExcelDataStore.get_project_rows("893")
        self.assertEqual(len(rows), 1)
        
        success = ExcelDataStore.delete_project_row("893", 0)
        self.assertTrue(success)
        
        deleted_rows = ExcelDataStore.get_project_rows("893")
        self.assertEqual(len(deleted_rows), 0)
  
    def test_pdf_generation(self):
        """Test that a PDF report is successfully generated from the Excel data."""
        pdf_test_path = os.path.join(UPLOAD_FOLDER, "Test_Project_Report.pdf")
        if os.path.exists(pdf_test_path):
            os.remove(pdf_test_path)
            
        try:
            generate_pdf_report(pdf_test_path, project_name="893")
            self.assertTrue(os.path.exists(pdf_test_path))
            self.assertGreater(os.path.getsize(pdf_test_path), 0)
        finally:
            if os.path.exists(pdf_test_path):
                try:
                    os.remove(pdf_test_path)
                except Exception:
                    pass

    def test_generate_report_route_returns_pdf_for_selected_project(self):
        """The report export route should return a downloadable PDF for the selected project."""
        with self.app.session_transaction() as sess:
            sess['logged_in'] = True
            sess['username'] = 'admin'

        response = self.app.get('/generate-report?project=893&comment=Test+comment')

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.content_type, 'application/pdf')
        self.assertGreater(len(response.data), 0)

    def test_pdf_generation_falls_back_when_browser_conversion_fails(self):
        """PDF generation should still succeed when the browser renderer times out."""
        pdf_test_path = os.path.join(UPLOAD_FOLDER, "Fallback_Project_Report.pdf")
        if os.path.exists(pdf_test_path):
            os.remove(pdf_test_path)

        with patch.object(app_module, 'convert_html_to_pdf', side_effect=RuntimeError('Headless browser timed out')):
            generate_pdf_report(pdf_test_path, project_name="893", comment="Fallback test")

        self.assertTrue(os.path.exists(pdf_test_path))
        self.assertGreater(os.path.getsize(pdf_test_path), 0)

    def test_load_transformed_filters_sheets_and_columns(self):
        """Test that api_local_load_transformed only saves the selected sheets and removes specified columns."""
        import json
        mock_file_path = os.path.join(UPLOAD_FOLDER, 'mock_upload.xlsx')
        
        df1 = pd.DataFrame([{"Category": "Cat A", "Test Method": "Method A", "RemoveMe": "Trash"}])
        df2 = pd.DataFrame([{"Category": "Cat B", "Test Method": "Method B"}])
        
        with pd.ExcelWriter(mock_file_path, engine='openpyxl') as writer:
            df1.to_excel(writer, index=False, sheet_name='SheetA')
            df2.to_excel(writer, index=False, sheet_name='SheetB')
            
        # Call the endpoint
        with open(mock_file_path, 'rb') as f:
            data = {
                'file': (f, 'mock_upload.xlsx'),
                'config': json.dumps({
                    'selected_sheets': ['SheetA'],
                    'removed_columns': {
                        'SheetA': ['RemoveMe']
                    }
                })
            }
            # Log in as admin
            with self.app.session_transaction() as sess:
                sess['logged_in'] = True
                sess['username'] = 'admin'
                
            response = self.app.post('/api/local/load-transformed', data=data, content_type='multipart/form-data')
            
        # Cleanup mock upload file
        if os.path.exists(mock_file_path):
            os.remove(mock_file_path)
            
        self.assertEqual(response.status_code, 200)
        res_data = response.get_json()
        self.assertTrue(res_data['success'])
        
        # Verify saved Excel file
        self.assertTrue(os.path.exists(self.test_excel_path))
        with pd.ExcelFile(self.test_excel_path) as xl:
            self.assertEqual(xl.sheet_names, ['SheetA'])
            df_saved = xl.parse('SheetA')
            
        # Verify columns
        self.assertIn('Category', df_saved.columns)
        self.assertNotIn('RemoveMe', df_saved.columns)

    def test_api_projects_scan(self):
        # Tests that the validation scan endpoint validates sheets, mapping, and integrity correctly.
        with self.app.session_transaction() as sess:
            sess['logged_in'] = True
            sess['username'] = 'admin'
            sess['_permanent'] = True

        import io
        excel_data = io.BytesIO()
        df = pd.DataFrame([
            {
                "Category": "Attachment",
                "Test Method": "Drop Test",
                "Test Number": "TM-003"
            }
        ])
        with pd.ExcelWriter(excel_data, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='SheetScan')
        excel_data.seek(0)

        response = self.app.post(
            '/api/projects/scan',
            data={'file': (excel_data, 'test_scan.xlsx')},
            content_type='multipart/form-data'
        )

        self.assertEqual(response.status_code, 200)
        res_data = response.get_json()
        self.assertTrue(res_data['success'])
        self.assertEqual(res_data['filename'], 'test_scan.xlsx')
        self.assertIn('SheetScan', res_data['sheets'])
        self.assertEqual(res_data['sheets']['SheetScan']['status'], 'warning')

    def test_laboratory_resolution(self):
        """Test resolve_record_laboratory resolution logic."""
        lab_test_numbers = {"Electronics": {"TM-100", "TM-101"}, "Mechanical": {"TM-200"}}
        lab_test_methods = {"Electronics": {"Drop Test"}, "Mechanical": {"Vibration Test"}}

        # Explicit Laboratory field present
        r1 = {"Laboratory": "Custom Lab", "Test Number": "TM-100", "Test Method": "Vibration Test"}
        self.assertEqual(app_module.resolve_record_laboratory(r1, lab_test_numbers, lab_test_methods), "Custom Lab")

        # Match by Test Number
        r2 = {"Test Number": "TM-101", "Test Method": "Unknown Method", "Category": "General"}
        self.assertEqual(app_module.resolve_record_laboratory(r2, lab_test_numbers, lab_test_methods), "Electronics")

        # Match by Test Method
        r3 = {"Test Number": "TM-999", "Test Method": "Vibration Test", "Category": "General"}
        self.assertEqual(app_module.resolve_record_laboratory(r3, lab_test_numbers, lab_test_methods), "Mechanical")

        # Fallback to Category
        r4 = {"Test Number": "TM-999", "Test Method": "Unknown", "Category": "Sensors"}
        self.assertEqual(app_module.resolve_record_laboratory(r4, lab_test_numbers, lab_test_methods), "Sensors")

    def test_api_master_data_laboratories(self):
        """Test /api/master-data/laboratories endpoint returns valid structure."""
        with self.app.session_transaction() as sess:
            sess['logged_in'] = True
            sess['username'] = 'admin'

        response = self.app.get('/api/master-data/laboratories')
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertIn("laboratories", data)
        self.assertIn("lab_test_numbers", data)
        self.assertIn("lab_test_methods", data)
        self.assertIsInstance(data["laboratories"], list)

    def test_pdf_generation_with_laboratory_filter(self):
        """Test PDF generation when a laboratory filter is applied."""
        pdf_test_path = os.path.join(UPLOAD_FOLDER, "Lab_Filtered_Report.pdf")
        if os.path.exists(pdf_test_path):
            os.remove(pdf_test_path)

        try:
            generate_pdf_report(pdf_test_path, project_name="893", laboratory_filter="Electronics")
            self.assertTrue(os.path.exists(pdf_test_path))
            self.assertGreater(os.path.getsize(pdf_test_path), 0)
        finally:
            if os.path.exists(pdf_test_path):
                try:
                    os.remove(pdf_test_path)
                except Exception:
                    pass

if __name__ == "__main__":
    unittest.main()



