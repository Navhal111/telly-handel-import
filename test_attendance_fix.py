#!/usr/bin/env python3
"""
Enhanced attendance Excel processing that properly handles the Excel format with overtime columns
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor
import pandas as pd

def test_attendance_excel_processing():
    """Test attendance processing with the actual Excel format from the image."""
    
    # Create sample data matching the Excel format in the image
    test_data = {
        'A': [None, None, None, 'S.NO', 1, 2, 3, None, None, None],
        'B': [None, None, None, 'EMPL NO', 20150701009, 20150701014, 20150702029, None, None, None],
        'C': [None, None, None, 'EMPLOYEE NAME', 'Said Ahmed Ibrahim', 'Shaaban Said Khamis', 'Silima Makame Haji', None, None, None],
        'D': [None, None, None, 'Attendance', 18, 19, 12, 49, None, None],
        'E': [None, None, None, 'BASIC', 4620.000, 5940.000, 4620.000, 15180.000, None, None],
        'F': [None, None, None, 'HRA', 1400.000, 1800.000, 1400.000, 4600.000, None, None],
        'G': [None, None, None, 'MEDICAL', 98.000, 126.000, 98.000, 322.000, None, None],
        'H': [None, 'PAYROLL STATEMENT FOR THE MONTH OF DECEMBER 2025', None, 'RESPONSIBLE ALLOWANCE', 1000.000, None, None, 1000.000, None, None],
        'I': [None, None, None, 'FUEL', None, None, 14000, 14000, None, None],
        'J': [None, None, None, 'FOOD ALLOWANCE', None, None, None, 0, None, None],
        'K': [None, None, None, 'OVERTIME @ 1.25', 2.25, 5.50, 7.75, 15.50, None, None],
        'L': [None, None, None, 'OVERTIME NORMAL DAYS', 8.663, 27.225, 29.838, 65.726, None, None],
        'M': [None, None, None, 'OVERTIME @ 1.50', 9.00, None, 2.50, 7.50, None, None],
        'N': [None, None, None, 'OVERTIME WEEKENDS', 19250.00, None, None, 28875, None, None],
        'O': [None, None, None, 'OVERTIME @ 2.00', 7.50, 6.25, 4.25, 18.00, None, None]
    }
    
    # Create DataFrame
    df = pd.DataFrame(test_data)
    
    print("🧪 Testing Enhanced Attendance Excel Processing")
    print("=" * 60)
    
    processor = ExcelProcessor()
    
    # Find header row
    employee_start_row = 3  # Row with headers (0-indexed)
    
    print(f"📊 Header row (row {employee_start_row + 1}):")
    header_row = df.iloc[employee_start_row]
    for idx, header in enumerate(header_row):
        if not pd.isna(header):
            print(f"  Column {chr(65+idx)}: {header}")
    
    # Test the enhanced column mapping
    column_mapping = {}
    attendance_details_mapping = {}
    
    for idx, header in enumerate(header_row):
        if not pd.isna(header):
            header_clean = str(header).strip().upper()
            print(f"🔍 Processing header: '{header_clean}'")
            
            if 'EMPL' in header_clean and 'NO' in header_clean:
                column_mapping['employee_no'] = idx
                print(f"  → Employee No: Column {idx}")
            elif 'EMPLOYEE' in header_clean and 'NAME' in header_clean:
                column_mapping['employee_name'] = idx
                print(f"  → Employee Name: Column {idx}")
            elif header_clean == 'ATTENDANCE':
                column_mapping['attendance_days'] = idx
                print(f"  → Attendance Days: Column {idx}")
            elif 'OVERTIME' in header_clean and '@' in header_clean:
                if '1.25' in header_clean:
                    attendance_details_mapping['overtime_125'] = idx
                    print(f"  → Overtime @ 1.25: Column {idx}")
                elif '1.50' in header_clean or '1.5' in header_clean:
                    attendance_details_mapping['overtime_150'] = idx
                    print(f"  → Overtime @ 1.50: Column {idx}")
                elif '2.00' in header_clean or '2.0' in header_clean:
                    attendance_details_mapping['overtime_200'] = idx
                    print(f"  → Overtime @ 2.00: Column {idx}")
            elif 'OVERTIME' in header_clean and 'WEEKEND' in header_clean:
                attendance_details_mapping['overtime_weekends'] = idx
                print(f"  → Overtime Weekends: Column {idx}")
            elif 'OVERTIME' in header_clean and 'NORMAL' in header_clean:
                attendance_details_mapping['overtime_normal'] = idx
                print(f"  → Overtime Normal Days: Column {idx}")
    
    print(f"\n📋 Column Mapping Results:")
    print(f"  Basic columns: {column_mapping}")
    print(f"  Attendance details: {attendance_details_mapping}")
    
    # Extract employee data
    print(f"\n👥 Extracting Employee Data:")
    data_rows = df.iloc[employee_start_row + 1:employee_start_row + 4]  # Rows 5-7 (employees)
    
    employee_data = []
    for idx, row in data_rows.iterrows():
        if row.isna().all():
            continue
            
        employee_record = {}
        
        # Basic data
        if 'employee_no' in column_mapping:
            emp_no = row.iloc[column_mapping['employee_no']]
            employee_record['employee_no'] = str(int(emp_no)) if not pd.isna(emp_no) else ''
        
        if 'employee_name' in column_mapping:
            emp_name = row.iloc[column_mapping['employee_name']]
            employee_record['employee_name'] = str(emp_name) if not pd.isna(emp_name) else ''
        
        if 'attendance_days' in column_mapping:
            att_days = row.iloc[column_mapping['attendance_days']]
            employee_record['attendance_days'] = float(att_days) if not pd.isna(att_days) else 0
        
        # Attendance details
        attendance_details = {}
        for detail_key, col_idx in attendance_details_mapping.items():
            cell_value = row.iloc[col_idx]
            if not pd.isna(cell_value):
                attendance_details[detail_key] = float(cell_value)
        
        if attendance_details:
            employee_record['attendance_details'] = attendance_details
        
        employee_data.append(employee_record)
        
        print(f"  Employee: {employee_record.get('employee_name', 'Unknown')}")
        print(f"    - Employee No: {employee_record.get('employee_no', 'N/A')}")
        print(f"    - Attendance Days: {employee_record.get('attendance_days', 0)}")
        print(f"    - Details: {attendance_details}")
        print()
    
    # Test XML generation
    print("🔄 Generating XML with extracted data...")
    
    result = {
        'success': True,
        'date': '2025-12-30',
        'company_name': 'Transworld Aviation Limited',
        'narration': 'Attendance for December 2025',
        'employee_data': employee_data,
        'total_employees': len(employee_data)
    }
    
    xml_content = processor.generate_attendance_xml(result, 'Transworld Aviation Limited')
    
    if xml_content:
        output_file = processor.save_xml_file(xml_content, 'attendance_fixed_extraction')
        print(f"✅ Generated XML: {output_file}")
        
        # Show attendance entries
        lines = xml_content.split('\n')
        print(f"\n📄 ATTENDANCEENTRIES.LIST sections:")
        for i, line in enumerate(lines):
            if 'ATTENDANCEENTRIES.LIST>' in line:
                for j in range(5):
                    if i+j < len(lines):
                        print(f"{i+j+1:3d}: {lines[i+j]}")
                print()
    else:
        print("❌ Failed to generate XML")

if __name__ == "__main__":
    test_attendance_excel_processing()