#!/usr/bin/env python3
"""
Quick test for the updated payroll XML generation
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from excel_processor import ExcelProcessor

def test_payroll_processing():
    processor = ExcelProcessor()
    
    # Look for any Excel files in the current directory
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') or f.endswith('.xls')]
    
    if not excel_files:
        print("❌ No Excel files found in current directory")
        return
        
    test_file = excel_files[0]
    print(f"🧪 Testing with file: {test_file}")
    
    # Process the payroll data
    result = processor.process_payroll_sheet(test_file)
    
    if result.get('success'):
        print(f"✅ Success! Generated XML: {result.get('xml_file')}")
        print(f"📊 Processed {len(result.get('employee_data', []))} employees")
        
        # Check if overtime columns are mapped correctly
        for employee in result.get('employee_data', []):
            print(f"\n👤 Employee: {employee.get('employee_name')}")
            salary_components = employee.get('salary_components', {})
            
            # Check for overtime entries
            overtime_keys = [key for key in salary_components.keys() if 'Overtime @' in key]
            if overtime_keys:
                print(f"⏰ Overtime entries found: {overtime_keys}")
                for ot_key in overtime_keys:
                    print(f"   {ot_key}: {salary_components[ot_key]}")
            else:
                print("⏰ No overtime entries found")
    else:
        print(f"❌ Error: {result.get('error')}")

if __name__ == "__main__":
    test_payroll_processing()