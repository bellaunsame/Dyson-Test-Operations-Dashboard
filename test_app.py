import os
import unittest
import datetime
import shutil
import pandas as pd
import app as app_module
from app import app, ExcelDataStore, generate_pdf_report, UPLOAD_FOLDER

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

if __name__ == "__main__":
    unittest.main()
