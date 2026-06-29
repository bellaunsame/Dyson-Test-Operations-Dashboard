import os
import unittest
import datetime
from app import app, db, ExcelDataStore, Project, TestRecord, ExcelMapping, generate_pdf_report, UPLOAD_FOLDER

class TestOpsDashboardBackend(unittest.TestCase):
    
    def setUp(self):
        # Configure app for testing with an in-memory SQLite database
        app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
        app.config['TESTING'] = True
        self.app = app.test_client()
        
        # Push application context
        self.app_context = app.app_context()
        self.app_context.push()
        
        # Create tables and initialize
        db.create_all()
        ExcelDataStore.initialize_db()
        
        # Explicitly seed test-specific data so tests are independent of production seeding
        proj = Project(name="893", description="Test Project", status="Active", created_at="06/28/2026")
        db.session.add(proj)
        
        rec = TestRecord(
            project_name="893",
            category="Device: Hair Dryer",
            test_method="Stress Imaging",
            test_number="TM-035128",
            start_date="2026-06-28",
            proto_weeks=3,
            proto_days=2,
            proto_qty=15,
            comments=""
        )
        db.session.add(rec)
        db.session.commit()

    def tearDown(self):
        # Clean up database and pop context
        db.session.remove()
        db.drop_all()
        self.app_context.pop()
 
    def test_database_initialization(self):
        """Test that the database is correctly initialized and contains our test seed data."""
        projects = Project.query.all()
        self.assertEqual(len(projects), 1)
        self.assertEqual(projects[0].name, "893")
        
        records = TestRecord.query.all()
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0].project_name, "893")
        self.assertEqual(records[0].category, "Device: Hair Dryer")
        self.assertEqual(records[0].test_method, "Stress Imaging")
        self.assertEqual(records[0].proto_qty, 15)
 
    def test_add_and_get_project(self):
        """Test adding a new project."""
        new_proj = Project(name="TestProj", description="Test desc", status="Active", created_at="06/29/2026")
        db.session.add(new_proj)
        db.session.commit()
        
        proj = Project.query.filter_by(name="TestProj").first()
        self.assertIsNotNone(proj)
        self.assertEqual(proj.description, "Test desc")
 
    def test_add_and_get_record(self):
        """Test adding a test record and verifying its fields."""
        rec = TestRecord(
            project_name="893",
            category="Device: Fan",
            test_method="Airflow Speed",
            test_number="TM-999",
            start_date="2026-07-15",
            proto_weeks=1,
            proto_days=2,
            proto_qty=10,
            comments="Loud motor noise"
        )
        db.session.add(rec)
        db.session.commit()
        
        saved = TestRecord.query.filter_by(test_number="TM-999").first()
        self.assertIsNotNone(saved)
        self.assertEqual(saved.category, "Device: Fan")
        self.assertEqual(saved.proto_qty, 10)
        self.assertEqual(saved.comments, "Loud motor noise")
 
    def test_update_record(self):
        """Test updating an existing test record."""
        rec = TestRecord.query.filter_by(test_number="TM-035128").first()
        self.assertIsNotNone(rec)
        
        # Update fields
        rec.proto_qty = 20
        rec.comments = "Updated comments"
        db.session.commit()
        
        updated = TestRecord.query.filter_by(test_number="TM-035128").first()
        self.assertEqual(updated.proto_qty, 20)
        self.assertEqual(updated.comments, "Updated comments")
 
    def test_delete_record(self):
        """Test deleting a record from the database."""
        rec = TestRecord.query.filter_by(test_number="TM-035128").first()
        self.assertIsNotNone(rec)
        
        db.session.delete(rec)
        db.session.commit()
        
        deleted = TestRecord.query.filter_by(test_number="TM-035128").first()
        self.assertIsNone(deleted)
 
    def test_pdf_generation(self):
        """Test that a PDF report is successfully generated for a project."""
        pdf_test_path = os.path.join(UPLOAD_FOLDER, "Test_Project_Report.pdf")
        if os.path.exists(pdf_test_path):
            os.remove(pdf_test_path)
            
        try:
            generate_pdf_report(pdf_test_path, project_name="893")
            self.assertTrue(os.path.exists(pdf_test_path))
            self.assertGreater(os.path.getsize(pdf_test_path), 0)
        finally:
            if os.path.exists(pdf_test_path):
                os.remove(pdf_test_path)
 
    def test_excel_mapping_model(self):
        """Test that the ExcelMapping model can save and retrieve mapping configurations."""
        mapping = ExcelMapping(
            project_name="893",
            file_path="/Shared Documents/Ops_Milestones_893.xlsx",
            sheet_name="Sheet1",
            mapping_json='{"Category": "Test Category", "Test Method": "Method Name"}'
        )
        db.session.add(mapping)
        db.session.commit()
 
        saved = ExcelMapping.query.filter_by(project_name="893").first()
        self.assertIsNotNone(saved)
        self.assertEqual(saved.sheet_name, "Sheet1")
        self.assertEqual(saved.to_dict()["mapping"]["Category"], "Test Category")
 
    def test_import_excel_endpoint(self):
        """Test that the global SharePoint load-transformed endpoint imports sheets as projects successfully."""
        # Set up logged_in session
        with self.app.session_transaction() as sess:
            sess['logged_in'] = True
            sess['username'] = 'admin'
 
        payload = {
            "file_path": "/Shared Documents/Ops_Milestones_893.xlsx",
            "selected_sheets": ["Sheet1"],
            "removed_columns": {
                "Sheet1": ["Rejections_Notes"]
            }
        }
        
        # Call the global load-transformed endpoint
        response = self.app.post(
            '/api/sharepoint/load-transformed',
            json=payload,
            follow_redirects=True
        )
        
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertIsNotNone(data)
        self.assertTrue(data["success"])
        self.assertGreater(data["count"], 0)
        
        # Verify that records were saved under project "Sheet1" (created from mock Excel sheet name)
        records = TestRecord.query.filter_by(project_name="Sheet1").all()
        self.assertGreater(len(records), 0)
        self.assertEqual(records[0].category, "Device: Hair Dryer")
        # In this test, we removed 'Rejections_Notes' which was mapped to 'comments'.
        # Let's verify that the comment is empty because it was removed from the query!
        self.assertEqual(records[0].comments, "")

if __name__ == "__main__":
    unittest.main()
