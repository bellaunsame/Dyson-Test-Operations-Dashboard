import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from app import app, db, ExcelDataStore, Project, TestRecord, generate_pdf_report, UPLOAD_FOLDER
import os

with app.app_context():
    ExcelDataStore.initialize_db()
    # Ensure project and record exist
    proj = Project.query.filter_by(name='SAMPLE').first()
    if not proj:
        proj = Project(name='SAMPLE', description='Sample Project', status='Active', created_at='06/30/2026')
        db.session.add(proj)
    # Clear old sample records
    TestRecord.query.filter_by(project_name='SAMPLE').delete()
    rec = TestRecord(
        project_name='SAMPLE', category='Packaging', test_method='Sample Method', test_number='TM-SAMPLE', start_date='2026-07-01',
        proto_weeks=1, proto_days=3, proto_qty=5,
        dvt_weeks=2, dvt_days=0, dvt_qty=3,
        evt_weeks=1, evt_days=0, evt_qty=2,
        pvt_weeks=0, pvt_days=5, pvt_qty=1,
        comments='Sample comment', defect_qty=1
    )
    db.session.add(rec)
    db.session.commit()
    out = os.path.join(UPLOAD_FOLDER, 'Sample_Gantt.pdf')
    try:
        generate_pdf_report(out, project_name='SAMPLE')
        print('WROTE', out, os.path.exists(out) and os.path.getsize(out))
    except Exception as e:
        print('ERROR', e)
