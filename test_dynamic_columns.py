#!/usr/bin/env python3
"""Test dynamic column detection for payroll."""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_dynamic_columns():
    """Test if dynamic columns are properly detected."""
    print("🧪 Testing Dynamic Column Detection")
    print("=" * 40)
    
    # Initialize processor
    processor = ExcelProcessor()
    
    # Test file path - using new test file with dynamic columns
    test_file = "test_dynamic_payroll.xlsx"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return
    
    print(f"📁 Testing with file: {test_file}")
    
    # Process payroll Excel
    result = processor.process_payroll_sheet(test_file)
    
    if not result.get('success', False):
        print(f"❌ Failed to process Excel: {result.get('error', 'Unknown error')}")
        return
    
    print(f"✅ Excel processing successful!")
    print(f"   📊 Total employees: {result.get('total_employees', 0)}")
    
    # Show employee data with dynamic columns
    employee_data = result.get('employee_data', [])
    if employee_data:
        print(f"\n👥 Employee salary components (showing ALL dynamic columns):")
        for i, emp in enumerate(employee_data):
            emp_name = emp.get('employee_name', 'Unknown')
            salary_components = emp.get('salary_components', {})
            
            print(f"\n   {i+1}. Employee: {emp_name}")
            print(f"      💰 Total Gross: {emp.get('total_gross_salary', 0):,.2f}")
            print(f"      📋 Salary Components Found:")
            
            for comp_name, amount in salary_components.items():
                print(f"         • {comp_name}: {amount:,.2f}")
    else:
        print("❌ No employee data found")

if __name__ == "__main__":
    test_dynamic_columns()