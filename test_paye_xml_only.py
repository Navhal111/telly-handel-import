#!/usr/bin/env python3
"""
Simple test script to generate PAYE XML without any Tally API calls.
This will help isolate the string concatenation error.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_paye_xml_only():
    """Test PAYE XML generation only (no API calls)."""
    
    print("🧪 Testing PAYE XML Generation Only (No API Calls)")
    print("=" * 60)
    
    # Initialize processor
    processor = ExcelProcessor()
    
    # Create test data with various None scenarios
    test_scenarios = [
        {
            "name": "All valid data",
            "data": {
                "success": True,
                "file_name": "test_paye.xlsx",
                "sheet_type": "PAYE",
                "date": "30/12/2025",
                "company_name": "My Company",
                "narration": "PAYE and SDL for December 2025",
                "employee_data": [
                    {
                        "employee_name": "John Doe",
                        "paye": 150000.00,
                        "sdl": 25000.00
                    },
                    {
                        "employee_name": "Jane Smith", 
                        "paye": 120000.00,
                        "sdl": 20000.00
                    }
                ],
                "total_employees": 2,
                "total_paye": 270000.00,
                "total_sdl": 45000.00,
                "total_amount": 315000.00
            },
            "company": "Test Company",
            "account": "Test Bank Account",
            "narration": "Test Narration"
        },
        {
            "name": "Some None values",
            "data": {
                "success": True,
                "file_name": "test_paye.xlsx",
                "sheet_type": "PAYE",
                "date": "30/12/2025",
                "company_name": None,
                "narration": None,
                "employee_data": [
                    {
                        "employee_name": "John Doe",
                        "paye": 150000.00,
                        "sdl": 25000.00
                    }
                ],
                "total_employees": 1,
                "total_paye": 150000.00,
                "total_sdl": 25000.00,
                "total_amount": 175000.00
            },
            "company": None,
            "account": None,
            "narration": None
        },
        {
            "name": "Employee with None name",
            "data": {
                "success": True,
                "file_name": "test_paye.xlsx",
                "sheet_type": "PAYE",
                "date": "30/12/2025",
                "company_name": "Test Company",
                "narration": "Test Narration",
                "employee_data": [
                    {
                        "employee_name": None,
                        "paye": 100000.00,
                        "sdl": 15000.00
                    },
                    {
                        "employee_name": "",
                        "paye": 80000.00,
                        "sdl": 12000.00
                    }
                ],
                "total_employees": 2,
                "total_paye": 180000.00,
                "total_sdl": 27000.00,
                "total_amount": 207000.00
            },
            "company": "Test Company",
            "account": "Test Bank",
            "narration": "Test Narration"
        }
    ]
    
    for i, scenario in enumerate(test_scenarios, 1):
        print(f"\n🔍 Scenario {i}: {scenario['name']}")
        print("-" * 40)
        
        try:
            xml_content = processor.generate_paye_xml(
                scenario["data"],
                company_name=scenario["company"],
                account_name=scenario["account"],
                narration=scenario["narration"]
            )
            
            if xml_content:
                print(f"✅ XML generation successful!")
                print(f"📄 XML length: {len(xml_content)} characters")
                
                # Check for "None" strings (should not exist)
                if "None" in xml_content:
                    print("❌ Found 'None' strings in XML:")
                    lines = xml_content.split('\n')
                    for line_no, line in enumerate(lines, 1):
                        if "None" in line:
                            print(f"   Line {line_no}: {line.strip()}")
                else:
                    print("✅ No 'None' strings found in XML")
                
                # Save XML file
                output_path = processor.save_xml_file(xml_content, "paye")
                if output_path:
                    print(f"💾 XML saved: {output_path}")
                    
                    # Check file naming
                    if "payroll_paye_import_" in output_path:
                        print("✅ Correct file naming pattern")
                
            else:
                print("❌ XML generation returned empty content")
                
        except Exception as e:
            print(f"❌ Error in scenario {i}: {str(e)}")
            import traceback
            print("Full traceback:")
            traceback.print_exc()
    
    print(f"\n🎯 PAYE XML Only Test Complete")
    print("=" * 60)

if __name__ == "__main__":
    test_paye_xml_only()