#!/usr/bin/env python3
"""
Test script to validate XML export functionality
"""

import sys
import os

# Add the src directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_xml_export():
    """Test XML export with sample attendance data."""
    
    processor = ExcelProcessor()
    
    # Create sample result data (simulating processed attendance)
    sample_result = {
        "success": True,
        "file_name": "test_attendance.xlsx",
        "sheet_type": "Attendance",
        "date": "09-10-2025",
        "company_name": "LIGHT",
        "narration": "Test attendance import",
        "employee_data": [
            {
                "employee_no": "1",
                "employee_name": "Ritesh",
                "attendance_type": "Present",
                "attendance_days": "23"
            },
            {
                "employee_no": "2", 
                "employee_name": "Milan",
                "attendance_type": "Absent",
                "attendance_days": "2"
            },
            {
                "employee_no": "3",
                "employee_name": "Anil", 
                "attendance_type": "Overtime @ 1.25",
                "attendance_days": "23"
            },
            {
                "employee_no": "4",
                "employee_name": "Utkarsh",
                "attendance_type": "Overtime @ 1.50", 
                "attendance_days": "12"
            }
        ],
        "total_employees": 4,
        "total_rows": 10,
        "total_columns": 4
    }
    
    print("🧪 Testing XML Export Functionality...")
    print("=" * 50)
    
    # Test XML export
    output_file = "test_attendance_export.xml"
    result = processor.export_attendance_xml(sample_result, output_file)
    
    if result:
        print(f"✅ XML Export successful: {result}")
        
        # Read and display the generated XML
        if os.path.exists(result):
            print("\n📄 Generated XML Content:")
            print("-" * 30)
            with open(result, 'r', encoding='utf-8') as f:
                content = f.read()
                print(content)
            
            print(f"\n📁 XML file saved to: {os.path.abspath(result)}")
            
            # Validate XML structure
            print("\n🔍 Validating XML Structure:")
            if "<ENVELOPE>" in content and "</ENVELOPE>" in content:
                print("✅ Valid XML envelope structure")
            if f"<SVCURRENTCOMPANY>{sample_result['company_name']}</SVCURRENTCOMPANY>" in content:
                print("✅ Company name correctly set")
            if "<DATE>20251009</DATE>" in content:
                print("✅ Date formatted correctly for Tally")
            if sample_result['narration'] in content:
                print("✅ Narration included")
            
            employee_count = content.count("<ATTENDANCEENTRIES.LIST>")
            if employee_count == len(sample_result['employee_data']):
                print(f"✅ All {employee_count} employees included in XML")
            else:
                print(f"❌ Employee count mismatch: expected {len(sample_result['employee_data'])}, found {employee_count}")
        
        else:
            print("❌ XML file was not created")
    else:
        print("❌ XML Export failed")
    
    print("\n" + "=" * 50)
    print("🏁 XML Export Test Complete")

if __name__ == "__main__":
    test_xml_export()