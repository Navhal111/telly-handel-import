#!/usr/bin/env python3
"""
Test both attendance and payroll export functionality
"""

import sys
import os

# Add the current directory to Python path
sys.path.append(os.path.dirname(__file__))

def test_export_functionality():
    """Test the export functionality without GUI"""
    from src.excel_processor import ExcelProcessor
    
    print("🧪 TESTING EXPORT FUNCTIONALITY")
    print("=" * 40)
    
    processor = ExcelProcessor()
    
    # Test attendance export
    if os.path.exists("sample_attendance.xlsx"):
        print("📊 Testing Attendance Export...")
        result = processor.process_attendance_sheet("sample_attendance.xlsx")
        
        if result.get('success', False):
            export_path = processor.export_attendance_data(result, "test_attendance_export.json")
            if export_path and os.path.exists(export_path):
                print(f"   ✅ Attendance export successful: {export_path}")
            else:
                print("   ❌ Attendance export failed")
        else:
            print("   ❌ Attendance processing failed")
    
    # Test payroll export
    if os.path.exists("sample_payroll.xlsx"):
        print("💰 Testing Payroll Export...")
        result = processor.process_payroll_sheet("sample_payroll.xlsx")
        
        if result.get('success', False):
            export_path = processor.export_payroll_data(result, "test_payroll_export.json")
            if export_path and os.path.exists(export_path):
                print(f"   ✅ Payroll export successful: {export_path}")
            else:
                print("   ❌ Payroll export failed")
        else:
            print("   ❌ Payroll processing failed")
    
    print("\n🎯 GUI Test:")
    print("   Now test the GUI export buttons - they should work without errors!")

if __name__ == "__main__":
    test_export_functionality()
    
    # Also start the GUI for manual testing
    print("\n🚀 Starting GUI for manual export testing...")
    from main import ModernExcelProcessor
    app = ModernExcelProcessor()
    app.run()