import os
import unittest
import pandas as pd
import datetime
from app import app, db, ExcelDataStore, Task, generate_pdf_report, UPLOAD_FOLDER, EXCEL_PATH

class TestDysonDashboardBackend(unittest.TestCase):
    
    def setUp(self):
        # Configure app for testing with an in-memory SQLite database
        app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
        app.config['TESTING'] = True
        self.app = app.test_client()
        
        # Backup existing Excel if it exists to avoid triggering migration in tests
        self.backup_path = EXCEL_PATH + ".bak"
        if os.path.exists(EXCEL_PATH):
            if os.path.exists(self.backup_path):
                os.remove(self.backup_path)
            os.rename(EXCEL_PATH, self.backup_path)
            
        # Push application context
        self.app_context = app.app_context()
        self.app_context.push()
        
        # Create tables and seed data
        db.create_all()
        ExcelDataStore.initialize_db()

    def tearDown(self):
        # Clean up database and pop context
        db.session.remove()
        db.drop_all()
        self.app_context.pop()
        
        # Restore Excel backup
        if os.path.exists(self.backup_path):
            if os.path.exists(EXCEL_PATH):
                os.remove(EXCEL_PATH)
            os.rename(self.backup_path, EXCEL_PATH)

    def test_excel_initialization(self):
        """Test that the database is correctly initialized with sample tasks."""
        tasks = ExcelDataStore.get_tasks()
        self.assertEqual(len(tasks), 6)  # 4 active tasks + 2 rejected tasks
        self.assertEqual(tasks[0]["Task ID"], "TSK-001")
        self.assertEqual(tasks[0]["Task Name"], "V15 Cyclone Engine Optimization")
        self.assertEqual(tasks[0]["Phase"], "Proto")

    def test_add_and_get_task(self):
        """Test adding a task and calculating the end date and status."""
        new_task = {
            "Task Name": "Airstrait Speed Run Testing",
            "Start Date": "2026-08-10",
            "Duration": 5,
            "Progress": 50,
            "Owner": "Engineering Team",
            "Phase": "EVT",
            "Status": "In Progress"
        }
        
        saved = ExcelDataStore.save_task(new_task)
        
        # Verify End Date calculation: 2026-08-10 + 5 days = 2026-08-15
        self.assertEqual(saved["End Date"], "2026-08-15")
        self.assertEqual(saved["Status"], "In Progress")
        self.assertEqual(saved["Phase"], "EVT")
        self.assertTrue(saved["Task ID"].startswith("TSK-"))
        
        # Verify it was written to database
        tasks = ExcelDataStore.get_tasks()
        self.assertEqual(len(tasks), 7)
        task_names = [t["Task Name"] for t in tasks]
        self.assertIn("Airstrait Speed Run Testing", task_names)

    def test_update_task(self):
        """Test updating an existing task."""
        task_data = {
            "Task ID": "TSK-001",
            "Task Name": "V15 Cyclone Engine Optimization - Updated",
            "Start Date": "2026-07-01",
            "Duration": 18,  # Changed from 15 to 18
            "Progress": 90,  # Changed from 80 to 90
            "Owner": "Engineering Team",
            "Phase": "Proto",
            "Status": "In Progress"
        }
        
        updated = ExcelDataStore.save_task(task_data)
        
        # Verify changes
        self.assertEqual(updated["Task Name"], "V15 Cyclone Engine Optimization - Updated")
        self.assertEqual(updated["Duration"], 18)
        self.assertEqual(updated["Progress"], 90)
        self.assertEqual(updated["End Date"], "2026-07-19")
        self.assertEqual(updated["Status"], "In Progress")
        self.assertEqual(updated["Phase"], "Proto")
        
        # Verify it is updated in database
        tasks = ExcelDataStore.get_tasks()
        self.assertEqual(len(tasks), 6)
        task = [t for t in tasks if t["Task ID"] == "TSK-001"][0]
        self.assertEqual(task["Task Name"], "V15 Cyclone Engine Optimization - Updated")

    def test_delete_task(self):
        """Test deleting a task from the database."""
        tasks_before = ExcelDataStore.get_tasks()
        self.assertEqual(len(tasks_before), 6)
        
        # Delete TSK-003
        success = ExcelDataStore.delete_task("TSK-003")
        self.assertTrue(success)
        
        tasks_after = ExcelDataStore.get_tasks()
        self.assertEqual(len(tasks_after), 5)
        task_ids = [t["Task ID"] for t in tasks_after]
        self.assertNotIn("TSK-003", task_ids)

    def test_pdf_generation(self):
        """Test that a PDF report is successfully generated and is non-empty."""
        pdf_test_path = os.path.join(UPLOAD_FOLDER, "Test_Report.pdf")
        if os.path.exists(pdf_test_path):
            os.remove(pdf_test_path)
            
        try:
            generate_pdf_report(pdf_test_path)
            self.assertTrue(os.path.exists(pdf_test_path))
            self.assertGreater(os.path.getsize(pdf_test_path), 0)
        finally:
            if os.path.exists(pdf_test_path):
                os.remove(pdf_test_path)

if __name__ == "__main__":
    unittest.main()
