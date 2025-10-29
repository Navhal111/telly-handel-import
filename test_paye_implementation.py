#!/usr/bin/env python3
"""
Test script to verify PAYE implementation functionality.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_paye_implementation():
    """Test the PAYE processing functionality."""
    
    print("🧪 Testing PAYE Implementation")
    print("=" * 50)
    
    # Initialize processor
    processor = ExcelProcessor()
    
    # Test the methods exist
    methods_to_check = [
        'process_paye_sheet',
        'find_paye_employee_data_start', 
        'extract_paye_employee_data',
        'generate_paye_xml'
    ]
    
    print("\n✅ Checking if PAYE methods exist:")
    for method_name in methods_to_check:
        if hasattr(processor, method_name):
            print(f"   ✓ {method_name} - Found")
        else:
            print(f"   ❌ {method_name} - Missing")
    
    # Test with a mock PAYE result
    print("\n🧪 Testing PAYE XML generation with mock data:")
    
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
                "sdl": 25000.00,
                "gross_salary": 500000.00
            },
            {
                "employee_name": "Jane Smith", 
                "paye": 120000.00,
                "sdl": 20000.00,
                "gross_salary": 400000.00
            }
        ],
        "total_employees": 2,
        "total_paye": 270000.00,
        "total_sdl": 45000.00,
        "total_amount": 315000.00
    }
    
    try:
        xml_content = processor.generate_paye_xml(
            mock_paye_result,
            company_name="TEST COMPANY",
            account_name="THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS",
            narration="Test PAYE and SDL payment"
        )
        
        if xml_content:
            print("   ✅ PAYE XML generation successful")
            print(f"   📄 XML length: {len(xml_content)} characters")
            
            # Check key elements in XML
            key_elements = [
                'VCHTYPE="Payment"',
                '<LEDGERNAME>PAYE</LEDGERNAME>',
                '<LEDGERNAME>Skill &amp; Development Levy</LEDGERNAME>',
                '<LEDGERNAME>THE PEOLPE\'S BANK OF ZANZIBAR LIMITED - TZS</LEDGERNAME>',
                '<NAME>John Doe</NAME>',
                '<NAME>Jane Smith</NAME>',
                '<AMOUNT>-270000.00</AMOUNT>',  # Total PAYE
                '<AMOUNT>-45000.00</AMOUNT>',   # Total SDL
                '<AMOUNT>315000.00</AMOUNT>'    # Bank credit
            ]
            
            print("   🔍 Checking XML content:")
            for element in key_elements:
                if element in xml_content:
                    print(f"      ✓ Found: {element}")
                else:
                    print(f"      ❌ Missing: {element}")
            
            # Save test XML file
            test_xml_path = "/Users/goku/Documents/excel_processor/test_paye_output.xml"
            with open(test_xml_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            print(f"   💾 Test XML saved: {test_xml_path}")
            
        else:
            print("   ❌ PAYE XML generation failed - no content returned")
            
    except Exception as e:
        print(f"   ❌ PAYE XML generation error: {str(e)}")
    
    print("\n🎯 PAYE Implementation Test Complete")
    print("=" * 50)

if __name__ == "__main__":
    test_paye_implementation()