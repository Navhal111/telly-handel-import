#!/usr/bin/env python3
"""Test fully dynamic column detection."""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_fully_dynamic():
    """Test completely dynamic column detection."""
    print("🧪 Testing FULLY DYNAMIC Column Detection")
    print("=" * 50)
    
    processor = ExcelProcessor()
    test_file = "test_dynamic_payroll.xlsx"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return
    
    print(f"📁 Processing: {test_file}")
    result = processor.process_payroll_sheet(test_file)
    
    if not result.get('success', False):
        print(f"❌ Failed: {result.get('error', 'Unknown error')}")
        return
    
    print("✅ Processing successful!")
    
    employee_data = result.get('employee_data', [])
    if employee_data:
        first_emp = employee_data[0]
        
        print(f"\n🎯 DYNAMIC ANALYSIS:")
        print(f"   📋 Column Headers Found: {first_emp.get('column_headers', [])}")
        print(f"   📊 All Data Keys: {list(first_emp.get('all_data', {}).keys())}")
        print(f"   💰 Salary Components: {list(first_emp.get('salary_components', {}).keys())}")
        
        print(f"\n👤 Employee Data:")
        print(f"   📍 Employee No: {first_emp.get('employee_no', 'N/A')}")
        print(f"   👤 Employee Name: {first_emp.get('employee_name', 'N/A')}")
        print(f"   💰 Total Gross: {first_emp.get('total_gross_salary', 0)}")
        
        print(f"\n📊 UI TABLE PREVIEW:")
        column_headers = first_emp.get('column_headers', [])
        print(f"   Columns: {' | '.join(column_headers + ['Total Gross'])}")
        
        # Show actual data row
        all_data = first_emp.get('all_data', {})
        row_values = []
        for col_name in column_headers:
            value = all_data.get(col_name, '')
            if isinstance(value, (int, float)):
                row_values.append(f"{value:,.0f}")
            else:
                row_values.append(str(value))
        row_values.append(f"{first_emp.get('total_gross_salary', 0):,.0f}")
        
        print(f"   Values:  {' | '.join(row_values)}")
        
        print(f"\n🎉 SUCCESS! The UI will show exactly these column names from your Excel:")
        print(f"   {column_headers}")
    else:
        print("❌ No employee data found")

if __name__ == "__main__":
    test_fully_dynamic()