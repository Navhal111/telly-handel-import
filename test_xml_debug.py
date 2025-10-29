#!/usr/bin/env python3
"""
Quick test script to check XML generation functionality
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from excel_processor import ExcelProcessor

def test_xml_generation():
    """Test basic XML generation"""
    processor = ExcelProcessor()
    
    # Test data
    test_result = {
        "success": True,
        "date": "2025-01-29",
        "company_name": "Test Company",
        "employee_data": [
            {
                "EMPL.NO": "001",
                "EMPLOYEE NAME": "John Doe",
                "DAYS PRESENT": 25,
                "TOTAL SALARY": 50000
            }
        ]
    }
    
    company_name = "Test Company"
    narration = "Test narration"
    
    print("Testing attendance XML generation...")
    try:
        xml_content = processor.generate_attendance_xml(test_result, company_name, narration)
        print("✅ Attendance XML generated successfully")
        print(f"XML length: {len(xml_content)} characters")
        
        # Check for None values in XML
        if "None" in xml_content:
            print("❌ Found 'None' values in XML!")
            # Find where None appears
            lines = xml_content.split('\n')
            for i, line in enumerate(lines):
                if "None" in line:
                    print(f"Line {i+1}: {line}")
        else:
            print("✅ No 'None' values found in XML")
            
    except Exception as e:
        print(f"❌ Error generating attendance XML: {e}")
        import traceback
        traceback.print_exc()
    
    print("\nTesting payroll XML generation...")
    try:
        xml_content = processor.generate_payroll_xml(test_result, company_name, "Cash", narration)
        print("✅ Payroll XML generated successfully")
        print(f"XML length: {len(xml_content)} characters")
        
        # Check for None values in XML
        if "None" in xml_content:
            print("❌ Found 'None' values in XML!")
            # Find where None appears
            lines = xml_content.split('\n')
            for i, line in enumerate(lines):
                if "None" in line:
                    print(f"Line {i+1}: {line}")
        else:
            print("✅ No 'None' values found in XML")
            
    except Exception as e:
        print(f"❌ Error generating payroll XML: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_xml_generation()