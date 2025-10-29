#!/usr/bin/env python3
"""
Test script for new Excel processing and XML generation functionality
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor

def test_payroll_xml_generation():
    """Test the new payroll XML generation functionality."""
    processor = ExcelProcessor()
    
    # Create sample payroll result data (simulating processed Excel data)
    sample_payroll_result = {
        "success": True,
        "file_name": "sample_payroll.xlsx",
        "sheet_type": "Payroll",
        "date": "2025-12-01",
        "company_name": "LIGHT COMPANY",
        "narration": "December 2025 Payroll",
        "employee_data": [
            {
                "employee_no": "07001009",
                "employee_name": "Said Ahmed Ibrahim", 
                "total_gross_salary": 8764.657,
                "all_salary_components": {
                    "EMPL NO": "07001009",
                    "EMPLOYEE NAME": "Said Ahmed Ibrahim",
                    "BASIC": 4627.00,
                    "HRA": 1400.00,
                    "MEDICAL": 98.00,
                    "FUEL": 1000.00,
                    "FOOD": 2.25,
                    "OVERTIME": 8.663,
                    "NIGHT": 19.250,
                    "GROSS SALARY": 8764.657,
                    "PAYE": 816.147,
                    "ADVANCE": 82.037,
                    "TOTAL DEDUCTION": 1300.50
                }
            },
            {
                "employee_no": "07001010", 
                "employee_name": "Ahmed Hassan Ali",
                "total_gross_salary": 7500.50,
                "all_salary_components": {
                    "EMPL NO": "07001010",
                    "EMPLOYEE NAME": "Ahmed Hassan Ali",
                    "BASIC": 4000.00,
                    "HRA": 1200.00,
                    "MEDICAL": 75.00,
                    "FUEL": 800.00,
                    "FOOD": 1.50,
                    "OVERTIME": 5.00,
                    "NIGHT": 15.00,
                    "GROSS SALARY": 7500.50,
                    "PAYE": 650.00,
                    "ADVANCE": 50.00,
                    "TOTAL DEDUCTION": 1100.00
                }
            }
        ],
        "total_employees": 2,
        "total_gross_salary": 16265.157,
        "average_salary": 8132.58
    }
    
    print("🧪 Testing Payroll XML Generation...")
    print("=" * 50)
    
    # Test XML generation
    xml_content = processor.generate_payroll_xml(
        sample_payroll_result, 
        company_name="LIGHT COMPANY",
        account_name="Cash"
    )
    
    if xml_content:
        print("✅ XML generation successful!")
        print(f"📊 Generated XML length: {len(xml_content)} characters")
        
        # Save to test file
        output_file = processor.save_xml_file(xml_content, "payroll_test")
        if output_file:
            print(f"📁 Test XML saved to: {output_file}")
            
            # Show first few lines of XML
            lines = xml_content.split('\n')
            print("\n📄 First 10 lines of generated XML:")
            print("-" * 40)
            for i, line in enumerate(lines[:10]):
                print(f"{i+1:2d}: {line}")
            print("... (truncated)")
            
            return True
    else:
        print("❌ XML generation failed!")
        return False

def test_attendance_xml_generation():
    """Test the attendance XML generation functionality."""
    processor = ExcelProcessor()
    
    # Create sample attendance result data
    sample_attendance_result = {
        "success": True,
        "file_name": "sample_attendance.xlsx", 
        "sheet_type": "Attendance",
        "date": "2025-12-01",
        "company_name": "LIGHT COMPANY",
        "narration": "December 2025 Attendance",
        "employee_data": [
            {
                "employee_no": "001",
                "employee_name": "John Doe",
                "attendance_type": "Regular",
                "attendance_days": 22
            },
            {
                "employee_no": "002", 
                "employee_name": "Jane Smith",
                "attendance_type": "Regular",
                "attendance_days": 20
            },
            {
                "employee_no": "003",
                "employee_name": "Bob Johnson", 
                "attendance_type": "Part-time",
                "attendance_days": 15
            }
        ],
        "total_employees": 3
    }
    
    print("\n🧪 Testing Attendance XML Generation...")
    print("=" * 50)
    
    # Test XML generation
    xml_content = processor.generate_attendance_xml(
        sample_attendance_result,
        company_name="LIGHT COMPANY"
    )
    
    if xml_content:
        print("✅ Attendance XML generation successful!")
        print(f"📊 Generated XML length: {len(xml_content)} characters")
        
        # Save to test file
        output_file = processor.save_xml_file(xml_content, "attendance_test")
        if output_file:
            print(f"📁 Test XML saved to: {output_file}")
            
            # Show first few lines of XML
            lines = xml_content.split('\n')
            print("\n📄 First 10 lines of generated XML:")
            print("-" * 40)
            for i, line in enumerate(lines[:10]):
                print(f"{i+1:2d}: {line}")
            print("... (truncated)")
            
            return True
    else:
        print("❌ Attendance XML generation failed!")
        return False

def main():
    """Run all tests."""
    print("🚀 Starting Excel Processor & XML Generator Tests")
    print("=" * 60)
    
    try:
        # Test payroll XML generation
        payroll_success = test_payroll_xml_generation()
        
        # Test attendance XML generation  
        attendance_success = test_attendance_xml_generation()
        
        # Summary
        print("\n📋 Test Summary:")
        print("=" * 30)
        print(f"Payroll XML Generation: {'✅ PASS' if payroll_success else '❌ FAIL'}")
        print(f"Attendance XML Generation: {'✅ PASS' if attendance_success else '❌ FAIL'}")
        
        if payroll_success and attendance_success:
            print("\n🎉 All tests passed! The new functionality is working correctly.")
            print("\n📝 Next steps:")
            print("   1. Run the main application: python main.py")
            print("   2. Select a company from Tally")
            print("   3. Upload Excel files and generate XML")
            print("   4. Import the generated XML files into Tally")
        else:
            print("\n⚠️ Some tests failed. Please check the error messages above.")
            
    except Exception as e:
        print(f"\n❌ Test execution failed: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()