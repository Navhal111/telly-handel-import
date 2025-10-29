#!/usr/bin/env python3
"""
Test script to verify the improved XML generation and file naming.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_xml_file_naming():
    """Test the improved XML file naming."""
    
    print("🧪 Testing XML File Naming and Generation")
    print("=" * 60)
    
    # Initialize processor
    processor = ExcelProcessor()
    
    # Test different file types
    test_content = '''<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE>
     <VOUCHER VCHTYPE="Test">
      <DATE>20251029</DATE>
      <NARRATION>Test XML</NARRATION>
     </VOUCHER>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>'''
    
    # Test different file types
    file_types = ["payroll", "paye", "attendance"]
    
    for file_type in file_types:
        print(f"\n🔍 Testing {file_type} XML generation...")
        
        try:
            output_path = processor.save_xml_file(test_content, file_type)
            if output_path:
                print(f"   ✅ Success: {output_path}")
                
                # Check if file was actually created
                if os.path.exists(output_path):
                    print(f"   📁 File exists: {os.path.getsize(output_path)} bytes")
                    
                    # Check filename pattern
                    expected_patterns = {
                        "payroll": "payroll_tally_import_",
                        "paye": "payroll_paye_import_",
                        "attendance": "attendance_tally_import_"
                    }
                    
                    if expected_patterns[file_type] in output_path:
                        print(f"   ✅ Correct naming pattern: {expected_patterns[file_type]}")
                    else:
                        print(f"   ❌ Wrong naming pattern. Expected: {expected_patterns[file_type]}")
                        
                else:
                    print(f"   ❌ File was not created: {output_path}")
            else:
                print(f"   ❌ Failed to save {file_type} XML")
                
        except Exception as e:
            print(f"   ❌ Error testing {file_type}: {str(e)}")
    
    print(f"\n🎯 XML File Naming Test Complete")
    print("=" * 60)
    
    # Test PAYE XML generation with mock data
    print(f"\n🧪 Testing PAYE XML Content Generation")
    print("-" * 40)
    
    mock_paye_result = {
        "success": True,
        "file_name": "test_paye.xlsx",
        "sheet_type": "PAYE",
        "date": "30/12/2025",
        "company_name": "TEST COMPANY",
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
    }
    
    try:
        paye_xml = processor.generate_paye_xml(
            mock_paye_result,
            company_name="TEST COMPANY",
            account_name="THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS",
            narration="Test PAYE and SDL payment"
        )
        
        if paye_xml:
            # Save the PAYE XML with proper naming
            paye_output_path = processor.save_xml_file(paye_xml, "paye")
            print(f"✅ PAYE XML generated and saved: {paye_output_path}")
            
            # Verify content
            if "payroll_paye_import_" in paye_output_path:
                print("✅ Correct PAYE file naming: payroll_paye_import_YYYYMMDD_HHMMSS.xml")
            
            if "VCHTYPE=\"Payment\"" in paye_xml:
                print("✅ Correct voucher type: Payment")
            
            if "LEDGERNAME>PAYE</LEDGERNAME" in paye_xml:
                print("✅ PAYE ledger found")
                
            if "LEDGERNAME>Skill &amp; Development Levy</LEDGERNAME" in paye_xml:
                print("✅ SDL ledger found")
            
            if "315000.00" in paye_xml:
                print("✅ Correct total amount in XML")
                
        else:
            print("❌ Failed to generate PAYE XML")
            
    except Exception as e:
        print(f"❌ Error generating PAYE XML: {str(e)}")

if __name__ == "__main__":
    test_xml_file_naming()