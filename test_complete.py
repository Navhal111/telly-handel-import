#!/usr/bin/env python3
"""
Complete test for both Attendance and Payroll processing
"""

import sys
import os

# Add the current directory to Python path
sys.path.append(os.path.dirname(__file__))

from main import ModernExcelProcessor

def main():
    print("🚀 Excel Processor - Full Feature Test")
    print("=" * 50)
    print("✅ Attendance Processing:")
    print("   📊 Extracts Date, Company, Narration")
    print("   👥 Employee attendance data with types and days")
    print("   📋 Table display with ID, Name, Type, Days")
    print()
    print("✅ Payroll Processing:")
    print("   📊 Extracts Date, Company, Account, Narration")
    print("   💰 Employee salary components (Basic, HRA, Medical, etc.)")
    print("   📈 Payroll summary with totals and averages")
    print("   📋 Table display with salary breakdown")
    print()
    print("✅ Both Support:")
    print("   🔄 Same window display with Back/Cancel buttons")
    print("   📤 Export to JSON functionality")
    print("   ⚠️ Validation with helpful error messages")
    print()
    print("📁 Available Test Files:")
    print("   • sample_attendance.xlsx - Test attendance processing")
    print("   • sample_payroll.xlsx - Test payroll processing")
    print()
    print("🎯 Test Instructions:")
    print("   1. Upload attendance file → Process → See employee attendance")
    print("   2. Upload payroll file → Process → See salary details")
    print("   3. Try invalid files → See validation errors")
    print("   4. Export data → Save as JSON")
    print()
    print("🚀 Starting application...")
    
    app = ModernExcelProcessor()
    app.run()

if __name__ == "__main__":
    main()