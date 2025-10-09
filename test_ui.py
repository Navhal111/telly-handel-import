#!/usr/bin/env python3
"""
Quick UI test to check if the interface improvements work
"""

import sys
import os

# Add the current directory to Python path
sys.path.append(os.path.dirname(__file__))

from main import ModernExcelProcessor

def main():
    print("🎨 Testing UI Improvements...")
    print("✅ Window size: 900x700 (increased from 800x600)")
    print("✅ Table columns: All 4 columns should be visible")  
    print("✅ Date format: Should show 09-10-2025 (not timestamp)")
    print("✅ Button layout: Better spacing and organization")
    print("✅ Section sizing: Reduced padding, better proportions")
    print()
    print("🚀 Starting application...")
    print("📝 Test steps:")
    print("   1. Click 'Browse File' and select sample_attendance.xlsx")
    print("   2. Click 'Process Attendance Sheet'")
    print("   3. Check if all table columns are visible")
    print("   4. Verify date shows as '09-10-2025'")
    print("   5. Test Back/Cancel buttons")
    print()
    
    app = ModernExcelProcessor()
    app.run()

if __name__ == "__main__":
    main()