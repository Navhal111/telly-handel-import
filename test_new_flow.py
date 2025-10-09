#!/usr/bin/env python3
"""
Test the new flow:
1. Load application
2. Show loading on button
3. Process data → new screen if valid OR error popup if invalid
"""

import sys
import os

# Add the current directory to Python path
sys.path.append(os.path.dirname(__file__))

from main import ModernExcelProcessor

def main():
    print("🚀 Starting Excel Processor with new flow...")
    print("📋 New Features:")
    print("   ✅ Loading button during processing")
    print("   ✅ New data screen for valid Excel files")  
    print("   ✅ Error popup for invalid Excel files")
    print("   ✅ No processing results in main screen")
    print("\n🎯 How to test:")
    print("   1. Select an attendance Excel file")
    print("   2. Click 'Process Attendance Sheet'")
    print("   3. Watch button show 'Processing...'")
    print("   4. If valid → Data screen opens")
    print("   5. If invalid → Error popup shows")
    print()
    
    app = ModernExcelProcessor()
    app.run()

if __name__ == "__main__":
    main()