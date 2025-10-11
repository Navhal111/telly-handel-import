#!/usr/bin/env python3
"""Test script for payroll upload functionality."""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor
from src.tally_api_service import TallyApiService


def test_payroll_processing():
    """Test payroll Excel processing and XML generation."""
    print("🧪 Testing Payroll Upload Functionality")
    print("=" * 50)
    
    # Initialize processor
    processor = ExcelProcessor()
    api_service = TallyApiService()
    
    # Test file path
    test_file = "sample_payroll.xlsx"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return
    
    print(f"📁 Testing with file: {test_file}")
    
    # Step 1: Process payroll Excel
    print("\n1️⃣ Processing payroll Excel file...")
    result = processor.process_payroll_sheet(test_file)
    
    if not result.get('success', False):
        print(f"❌ Failed to process Excel: {result.get('error', 'Unknown error')}")
        return
    
    print(f"✅ Excel processing successful!")
    print(f"   📊 Total employees: {result.get('total_employees', 0)}")
    print(f"   💰 Total gross salary: {result.get('total_gross_salary', 0):,.2f}")
    print(f"   📅 Date: {result.get('date', 'N/A')}")
    print(f"   🏢 Company: {result.get('company_name', 'N/A')}")
    
    # Step 2: Show some employee data
    employee_data = result.get('employee_data', [])
    if employee_data:
        print(f"\n👥 Sample employee data (first 2 employees):")
        for i, emp in enumerate(employee_data[:2]):
            print(f"   {i+1}. {emp.get('employee_name', 'Unknown')}")
            print(f"      💰 Gross: {emp.get('total_gross_salary', 0):,.2f}")
            salary_components = emp.get('salary_components', {})
            print(f"      📋 Components: {list(salary_components.keys())[:3]}...")
    
    # Step 3: Generate XML
    print("\n2️⃣ Generating XML for Tally...")
    try:
        company_name = "LIGHT"  # Test company
        account_name = "PEAPULE BANK OF UNITE"  # From Excel sample
        
        xml_content = api_service._generate_payroll_xml(result, company_name, account_name)
        
        print("✅ XML generation successful!")
        print(f"   📄 XML length: {len(xml_content)} characters")
        
        # Show XML preview
        lines = xml_content.split('\n')
        print(f"\n📄 XML Preview (first 15 lines):")
        for i, line in enumerate(lines[:15]):
            print(f"   {line}")
        if len(lines) > 15:
            print(f"   ... ({len(lines) - 15} more lines)")
        
        # Save XML for inspection
        xml_file = "test_payroll_output.xml"
        with open(xml_file, 'w', encoding='utf-8') as f:
            f.write(xml_content)
        print(f"\n💾 Full XML saved to: {xml_file}")
        
    except Exception as e:
        print(f"❌ XML generation failed: {str(e)}")
        return
    
    # Step 4: Test XML structure validation
    print("\n3️⃣ Validating XML structure...")
    try:
        import xml.etree.ElementTree as ET
        root = ET.fromstring(xml_content)
        
        # Check basic structure
        voucher = root.find('.//VOUCHER')
        if voucher is not None:
            vchtype = voucher.get('VCHTYPE')
            action = voucher.get('ACTION')
            print(f"✅ XML structure valid!")
            print(f"   📋 Voucher Type: {vchtype}")
            print(f"   🔧 Action: {action}")
            
            # Count employees
            employees = root.findall('.//EMPLOYEEENTRIES.LIST')
            print(f"   👥 Employees in XML: {len(employees)}")
            
            # Count payhead allocations
            payheads = root.findall('.//PAYHEADALLOCATIONS.LIST')
            print(f"   💰 Payhead allocations: {len(payheads)}")
            
        else:
            print("⚠️ VOUCHER element not found in XML")
            
    except ET.ParseError as e:
        print(f"❌ XML parsing error: {str(e)}")
    
    print("\n🎉 Payroll upload test completed!")
    print("\nNOTE: To complete the test, you would need:")
    print("   1. Tally running on localhost:9000")
    print("   2. Call api_service.upload_payroll_data(result, company_name, account_name)")


if __name__ == "__main__":
    test_payroll_processing()