#!/usr/bin/env python3
"""Simple test of Excel reading."""

import pandas as pd
import os

def simple_test():
    print("=== Simple Excel Test ===")
    
    try:
        # Test file exists
        test_file = "test_dynamic_payroll.xlsx"
        if os.path.exists(test_file):
            print(f"✅ File exists: {test_file}")
        else:
            print(f"❌ File not found: {test_file}")
            return
        
        # Read Excel
        df = pd.read_excel(test_file, header=None)
        print(f"✅ Excel read successfully: {df.shape}")
        
        # Show structure
        print("\n📊 Excel Structure:")
        for i in range(min(8, len(df))):
            row_data = [str(x) for x in df.iloc[i].tolist()]
            print(f"   Row {i}: {row_data}")
        
        # Find header row (row 5 in our test file)
        header_row_idx = 5
        if len(df) > header_row_idx:
            header_row = df.iloc[header_row_idx]
            print(f"\n📋 Header Row {header_row_idx}:")
            headers = []
            for idx, val in enumerate(header_row):
                if not pd.isna(val):
                    header_name = str(val).strip()
                    if header_name:
                        headers.append(header_name)
                        print(f"   Column {idx}: {header_name}")
            
            print(f"\n🎯 Dynamic Headers Found: {headers}")
            
            # Test data row
            if len(df) > header_row_idx + 1:
                data_row = df.iloc[header_row_idx + 1]
                print(f"\n📊 Sample Data Row:")
                for idx, val in enumerate(data_row):
                    if idx < len(headers):
                        print(f"   {headers[idx]}: {val}")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    simple_test()