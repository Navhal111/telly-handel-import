#!/usr/bin/env python3
"""Test complete payroll workflow with dynamic columns."""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_complete_dynamic_workflow():
    """Test complete workflow with dynamic columns."""
    print("🧪 Testing Complete Dynamic Column Workflow")
    print("=" * 50)
    
    # Initialize processor
    processor = ExcelProcessor()
    
    # Test with dynamic columns file
    test_file = "test_dynamic_payroll.xlsx"
    
    if not os.path.exists(test_file):
        print(f"❌ Test file not found: {test_file}")
        return
    
    print(f"📁 Processing: {test_file}")
    
    # Process payroll Excel
    result = processor.process_payroll_sheet(test_file)
    
    if not result.get('success', False):
        print(f"❌ Failed to process Excel: {result.get('error', 'Unknown error')}")
        return
    
    print(f"✅ Excel processing successful!")
    
    # Show detailed employee data
    employee_data = result.get('employee_data', [])
    if employee_data:
        print(f"\n👥 Processing {len(employee_data)} employees with DYNAMIC columns:")
        
        for i, emp in enumerate(employee_data):
            emp_name = emp.get('employee_name', 'Unknown')
            emp_no = emp.get('employee_no', 'N/A')
            salary_components = emp.get('salary_components', {})
            total_gross = emp.get('total_gross_salary', 0)
            
            print(f"\n   Employee {i+1}:")
            print(f"      📋 ID: {emp_no}")
            print(f"      👤 Name: {emp_name}")
            print(f"      💰 Total Gross: {total_gross:,.2f}")
            print(f"      🏷️  Dynamic Salary Components:")
            
            for comp_name, amount in salary_components.items():
                print(f"         • {comp_name}: {amount:,.2f}")
        
        # Simulate what UI will display
        print(f"\n📊 UI DISPLAY SIMULATION:")
        print(f"   🔹 Columns that will appear in UI table:")
        
        first_emp = employee_data[0]
        salary_component_names = list(first_emp.get('salary_components', {}).keys())
        ui_columns = ['Employee ID', 'Employee Name'] + salary_component_names + ['Total Gross']
        
        print(f"      {' | '.join(ui_columns)}")
        
        # Show sample row
        print(f"   🔹 Sample data row:")
        sample_values = [
            first_emp.get('employee_no', 'N/A'),
            first_emp.get('employee_name', 'N/A')
        ]
        
        for comp_name in salary_component_names:
            sample_values.append(f"{first_emp.get('salary_components', {}).get(comp_name, 0):,.0f}")
        
        sample_values.append(f"{first_emp.get('total_gross_salary', 0):,.0f}")
        
        print(f"      {' | '.join(sample_values)}")
    
    print(f"\n🎉 Test completed!")
    print(f"✅ The UI will now show YOUR Excel column names: {salary_component_names}")
    print(f"✅ No more hardcoded 'Basic Salary', 'HRA', 'Medical' columns!")

if __name__ == "__main__":
    test_complete_dynamic_workflow()