"""
Standalone CLI to generate a clean, traditional Gantt chart PNG from the project's Excel data.
Usage:
    python gantt_chart.py --project 893 --out mygantt.png

If no project specified, all project rows are plotted.
"""
import argparse
import os
from app import ExcelDataStore, generate_gantt_chart_image

parser = argparse.ArgumentParser(description='Generate a Gantt chart image from Excel project data')
parser.add_argument('--project', '-p', help='Project (sheet) name to render', default=None)
parser.add_argument('--out', '-o', help='Output image path', default='gantt_output.png')
parser.add_argument('--dpi', type=int, default=200)

args = parser.parse_args()

records = []
if args.project:
    records = ExcelDataStore.get_project_rows(args.project)
else:
    projects = ExcelDataStore.get_projects()
    for p in projects:
        records.extend(ExcelDataStore.get_project_rows(p['name']))

if not records:
    print('No records found for the requested project(s).')
else:
    generate_gantt_chart_image(records, args.out, project_name=args.project, theme='light', timeline_type='weeks', dpi=args.dpi)
    print(f'Gantt image written to {args.out}')
