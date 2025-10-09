#!/usr/bin/env python3
"""
Debug script to test the Excel processing directly
"""

import sys
import os

# Add the current directory to Python path
sys.path.append(os.path.dirname(__file__))

from src.excel_processor import ExcelProcessor

def test_excel_processing():
    print("🔍 DEBUG: Testing Excel processing...")
    
    processor = ExcelProcessor()
    
    # Test with the sample file
    sample_file = "sample_attendance.xlsx"
    
    if os.path.exists(sample_file):
        print(f"📁 Testing with file: {sample_file}")
        result = processor.process_attendance_sheet(sample_file)
        
        print(f"\n📊 PROCESSING RESULT:")
        print(f"Success: {result.get('success', False)}")
        print(f"Keys: {list(result.keys()) if result else 'None'}")
        
        if result.get('success', False):
            print(f"\n📅 HEADER INFO:")
            print(f"Date: {result.get('date', 'Not found')}")
            print(f"Company: {result.get('company_name', 'Not found')}")
            print(f"Narration: {result.get('narration', 'Not found')}")
            
            employee_data = result.get('employee_data', [])
            print(f"\n👥 EMPLOYEE DATA ({len(employee_data)} employees):")
            for i, emp in enumerate(employee_data[:3], 1):  # Show first 3
                print(f"  {i}. ID: {emp.get('employee_no', 'N/A')}, Name: {emp.get('employee_name', 'N/A')}, Type: {emp.get('attendance_type', 'N/A')}, Days: {emp.get('attendance_days', 'N/A')}")
        else:
            print(f"❌ Error: {result.get('error', 'Unknown error')}")
            if 'validation_errors' in result:
                print(f"Validation errors: {result['validation_errors']}")
    else:
        print(f"❌ Sample file not found: {sample_file}")
        print("Create it first with: python test_create_sample.py")

if __name__ == "__main__":
    test_excel_processing()