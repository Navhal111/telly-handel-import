#!/usr/bin/env python3
"""
Debug script to test the payroll processing
"""

import sys
import os

# Add the current directory to Python path
sys.path.append(os.path.dirname(__file__))

from src.excel_processor import ExcelProcessor

def test_payroll_processing():
    print("🔍 DEBUG: Testing Payroll processing...")
    
    processor = ExcelProcessor()
    
    # Test with the sample payroll file
    sample_file = "sample_payroll.xlsx"
    
    if os.path.exists(sample_file):
        print(f"📁 Testing with file: {sample_file}")
        result = processor.process_payroll_sheet(sample_file)
        
        print(f"\n📊 PAYROLL PROCESSING RESULT:")
        print(f"Success: {result.get('success', False)}")
        print(f"Keys: {list(result.keys()) if result else 'None'}")
        
        if result.get('success', False):
            print(f"\n📅 HEADER INFO:")
            print(f"Date: {result.get('date', 'Not found')}")
            print(f"Company: {result.get('company_name', 'Not found')}")
            print(f"Account: {result.get('account', 'Not found')}")
            print(f"Narration: {result.get('narration', 'Not found')}")
            
            print(f"\n💰 PAYROLL SUMMARY:")
            print(f"Total Employees: {result.get('total_employees', 0)}")
            print(f"Total Gross Salary: {result.get('total_gross_salary', 0):,.2f}")
            print(f"Average Salary: {result.get('average_salary', 0):,.2f}")
            
            employee_data = result.get('employee_data', [])
            print(f"\n👥 EMPLOYEE PAYROLL DATA ({len(employee_data)} employees):")
            for i, emp in enumerate(employee_data[:3], 1):  # Show first 3
                salary_components = emp.get('salary_components', {})
                print(f"  {i}. ID: {emp.get('employee_no', 'N/A')}, Name: {emp.get('employee_name', 'N/A')}")
                print(f"     Basic: {salary_components.get('BASIC', 0):,.0f}, HRA: {salary_components.get('HRA', 0):,.0f}")
                print(f"     Total Gross: {emp.get('total_gross_salary', 0):,.0f}")
        else:
            print(f"❌ Error: {result.get('error', 'Unknown error')}")
            if 'validation_errors' in result:
                print(f"Validation errors: {result['validation_errors']}")
    else:
        print(f"❌ Sample file not found: {sample_file}")
        print("Create it first with: python create_sample_payroll.py")

if __name__ == "__main__":
    test_payroll_processing()