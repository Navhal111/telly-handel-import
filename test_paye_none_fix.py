#!/usr/bin/env python3
"""
Test script to verify PAYE implementation is fixed for None string concatenation.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_paye_with_none_values():
    """Test PAYE processing with potential None values."""
    
    print("🧪 Testing PAYE with None Values")
    print("=" * 50)
    
    # Initialize processor
    processor = ExcelProcessor()
    
    # Test with problematic data that might have None values
    mock_paye_result = {
        "success": True,
        "file_name": "test_paye.xlsx",
        "sheet_type": "PAYE",
        "date": "30/12/2025",
        "company_name": None,  # This could be None
        "narration": None,      # This could be None
        "employee_data": [
            {
                "employee_name": None,  # This could be None
                "paye": 150000.00,
                "sdl": 25000.00
            },
            {
                "employee_name": "Jane Smith", 
                "paye": 120000.00,
                "sdl": 20000.00
            },
            {
                "employee_name": "",    # Empty string
                "paye": 80000.00,
                "sdl": 15000.00
            }
        ],
        "total_employees": 3,
        "total_paye": 350000.00,
        "total_sdl": 60000.00,
        "total_amount": 410000.00
    }
    
    print("🔍 Testing PAYE XML generation with None values...")
    
    try:
        xml_content = processor.generate_paye_xml(
            mock_paye_result,
            company_name=None,  # Also testing None company name
            account_name=None,  # Also testing None account name
            narration=None      # Also testing None narration
        )
        
        if xml_content:
            print("✅ PAYE XML generation successful!")
            print(f"📄 XML length: {len(xml_content)} characters")
            
            # Check that we don't have any "None" strings in the XML
            if "None" in xml_content:
                print("❌ Found 'None' strings in XML - this should not happen")
                # Find and show where None appears
                lines = xml_content.split('\n')
                for i, line in enumerate(lines):
                    if "None" in line:
                        print(f"   Line {i+1}: {line.strip()}")
            else:
                print("✅ No 'None' strings found in XML")
            
            # Check for proper default values
            checks = [
                ("TEST COMPANY", "Default company name used"),
                ("PAYE and SDL for December 2025", "Default narration used"),  
                ("THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS", "Default bank account used"),
                ("Unknown Employee", "Default employee name used for None values"),
                ("<AMOUNT>410000.00</AMOUNT>", "Correct total amount"),
                ("<AMOUNT>-350000.00</AMOUNT>", "Correct PAYE amount"),
                ("<AMOUNT>-60000.00</AMOUNT>", "Correct SDL amount")
            ]
            
            print("\n🔍 Checking XML content:")
            for check_text, description in checks:
                if check_text in xml_content:
                    print(f"   ✅ {description}")
                else:
                    print(f"   ⚠️  {description} - not found")
            
            # Save test XML
            test_output = processor.save_xml_file(xml_content, "paye")
            if test_output:
                print(f"\n💾 Test XML saved: {test_output}")
                
                # Verify file naming
                if "payroll_paye_import_" in test_output:
                    print("✅ Correct PAYE file naming pattern")
                else:
                    print("❌ Incorrect file naming pattern")
            
        else:
            print("❌ PAYE XML generation failed - no content returned")
            
    except Exception as e:
        print(f"❌ PAYE XML generation error: {str(e)}")
        import traceback
        print("Full traceback:")
        traceback.print_exc()
    
    print("\n🎯 PAYE None Values Test Complete")
    print("=" * 50)

if __name__ == "__main__":
    test_paye_with_none_values()