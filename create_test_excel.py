#!/usr/bin/env python3
"""Create a test Excel file matching the user's image structure."""

import pandas as pd
import os

def create_test_excel_with_dynamic_columns():
    """Create Excel file with dynamic columns like in user's image."""
    
    # Create data structure matching the user's image
    data = {
        'A': ['Date', 'Company Name', 'Account', 'Narration', None, 'EMPL NO', 1],
        'B': ['01-04-2025', 'LIGHT', 'Cash', 'Monthly Payroll', None, 'EMPLOYEE NAME', 'Ritesh'], 
        'C': [None, None, None, None, None, 'Check', 150],
        'D': [None, None, None, None, None, 'ON Hand', 200],
        'E': [None, None, None, None, None, None, None],
        'F': [None, None, None, None, None, None, None]
    }
    
    # Create DataFrame
    df = pd.DataFrame.from_dict(data)
    
    # Save to Excel
    output_file = "test_dynamic_payroll.xlsx"
    df.to_excel(output_file, index=False, header=False)
    
    print(f"✅ Created test Excel file: {output_file}")
    print("📋 Structure:")
    print("   Row 1: Date = 01-04-2025")
    print("   Row 2: Company Name = LIGHT") 
    print("   Row 3: Account = Cash")
    print("   Row 4: Narration = Monthly Payroll")
    print("   Row 6: Headers = EMPL NO, EMPLOYEE NAME, Check, ON Hand")
    print("   Row 7: Data = 1, Ritesh, 150, 200")
    
    return output_file

if __name__ == "__main__":
    create_test_excel_with_dynamic_columns()