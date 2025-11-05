import pandas as pd
import os
from typing import Optional, Dict, Any, List
import re
from datetime import datetime


class ExcelProcessor:
    """Class to handle Excel file processing for Attendance and Payroll sheets."""
    
    def __init__(self):
        self.supported_extensions = ['.xlsx', '.xls']
        
        # Expected structure for attendance sheet validation
        self.attendance_required_headers = [
            'EMPL NO', 'EMPLOYEE NAME', 'Attendance/Production Types', 'Attendance Days'
        ]
    
    def validate_file(self, file_path: str) -> bool:
        """Validate if the file exists and has a supported extension."""
        if not os.path.exists(file_path):
            return False
        
        _, ext = os.path.splitext(file_path)
        return ext.lower() in self.supported_extensions
    
    def read_excel_file(self, file_path: str, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
        """Read Excel file and return DataFrame."""
        try:
            if not self.validate_file(file_path):
                raise ValueError("Invalid file or unsupported format")
            
            # First, get the sheet names to understand the file structure
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            # If no specific sheet is requested, use the first sheet
            if sheet_name is None:
                sheet_name = sheet_names[0] if sheet_names else 0
            
            # Read the specific sheet without header detection
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # Handle case where multiple sheets are returned as dict
            if isinstance(df, dict):
                # Take the first sheet from the dictionary
                first_key = list(df.keys())[0]
                df = df[first_key]
            
            print(f"📊 Successfully read Excel file: {df.shape[0]} rows × {df.shape[1]} columns")
            return df
        
        except Exception as e:
            print(f"❌ Failed to read Excel file: {str(e)}")
            return None
    
    def extract_header_info(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Extract Date, Company Name, and Narration from the first few rows."""
        header_info = {
            "date": None,
            "company_name": None,
            "narration": None
        }
        
        try:
            # Check if we have enough rows for header information
            if len(df) < 3:
                return header_info
            
            # Try to extract from first 3 rows, column B (index 1)
            if len(df.columns) >= 2:
                # Row 1: Date (09-10-2025)
                date_value = df.iloc[0, 1] if not pd.isna(df.iloc[0, 1]) else None
                if date_value:
                    # Handle different date formats
                    date_str = str(date_value).strip()
                    if 'Timestamp' in str(type(date_value)) or '00:00:00' in date_str:
                        # Convert pandas timestamp to readable format
                        try:
                            if hasattr(date_value, 'strftime'):
                                date_str = date_value.strftime('%d-%m-%Y')
                            else:
                                # Extract just the date part if it's a string with time
                                date_str = date_str.split(' ')[0]
                        except:
                            pass
                    header_info["date"] = date_str
                
                # Row 2: Company Name (LIGHT)
                company_value = df.iloc[1, 1] if not pd.isna(df.iloc[1, 1]) else None
                if company_value:
                    header_info["company_name"] = str(company_value).strip()
                
                # Row 3: Narration (Test attendance)
                narration_value = df.iloc[2, 1] if not pd.isna(df.iloc[2, 1]) else None
                if narration_value:
                    header_info["narration"] = str(narration_value).strip()
            
        except Exception as e:
            print(f"⚠️ Warning: Could not extract header info: {str(e)}")
        
        return header_info
    
    def find_employee_data_start(self, df: pd.DataFrame) -> int:
        """Find the row where employee data starts by looking for 'EMPL NO' pattern."""
        try:
            for idx, row in df.iterrows():
                # Check if any cell in this row contains 'EMPL NO'
                row_values = [str(val).strip().upper() for val in row if not pd.isna(val)]
                if any('EMPL' in val and 'NO' in val for val in row_values):
                    return idx
            
            # Fallback: look for row with multiple non-null values that could be headers
            for idx, row in df.iterrows():
                non_null_count = row.notna().sum()
                if non_null_count >= 3:  # At least 3 columns should have data
                    row_str = ' '.join([str(val) for val in row if not pd.isna(val)]).upper()
                    if 'EMPLOYEE' in row_str or 'NAME' in row_str:
                        return idx
            
            return 4  # Default fallback to row 5 (index 4)
            
        except Exception as e:
            print(f"⚠️ Warning: Could not find employee data start: {str(e)}")
            return 4
    
    def validate_attendance_structure(self, df: pd.DataFrame, employee_start_row: int) -> Dict[str, Any]:
        """Validate if the Excel file has the expected attendance structure."""
        validation_result = {
            "is_valid": False,
            "errors": [],
            "warnings": []
        }
        
        try:
            # Check if we have enough rows
            if len(df) <= employee_start_row:
                validation_result["errors"].append("File does not have enough data rows")
                return validation_result
            
            # Extract header row
            header_row = df.iloc[employee_start_row]
            headers = [str(val).strip() for val in header_row if not pd.isna(val)]
            
            # Check for required columns
            header_text = ' '.join(headers).upper()
            required_patterns = ['EMPL', 'EMPLOYEE', 'NAME', 'ATTENDANCE']
            
            missing_patterns = []
            for pattern in required_patterns:
                if pattern not in header_text:
                    missing_patterns.append(pattern)
            
            if missing_patterns:
                validation_result["errors"].append(f"Missing expected headers: {', '.join(missing_patterns)}")
            
            # Check if we have employee data
            data_rows = df.iloc[employee_start_row + 1:]
            non_empty_data_rows = data_rows.dropna(how='all')
            
            if len(non_empty_data_rows) == 0:
                validation_result["errors"].append("No employee data found")
            elif len(non_empty_data_rows) < 2:
                validation_result["warnings"].append("Very few employee records found")
            
            # If no critical errors, mark as valid
            if not validation_result["errors"]:
                validation_result["is_valid"] = True
            
        except Exception as e:
            validation_result["errors"].append(f"Validation error: {str(e)}")
        
        return validation_result
    
    def extract_employee_data(self, df: pd.DataFrame, employee_start_row: int) -> List[Dict[str, Any]]:
        """Extract employee data from the DataFrame."""
        employee_data = []
        
        try:
            # Get header row
            header_row = df.iloc[employee_start_row]
            
            # Create comprehensive column mapping for attendance
            column_mapping = {}
            
            # Find basic columns
            for idx, header in enumerate(header_row):
                if not pd.isna(header):
                    header_clean = str(header).strip().upper()
                    
                    if 'EMPL' in header_clean and 'NO' in header_clean:
                        column_mapping['employee_no'] = idx
                    elif 'EMPLOYEE' in header_clean and 'NAME' in header_clean:
                        column_mapping['employee_name'] = idx
                    elif 'ATTENDANCE' in header_clean and 'PRODUCTION' in header_clean:
                        column_mapping['attendance_type'] = idx
                    elif ('ATTENDANCE' in header_clean and 'DAYS' in header_clean) or header_clean == 'ATTENDANCE':
                        column_mapping['attendance_days'] = idx
            
            print(f"Using column mapping: {column_mapping}")
            print(f"Fixed overtime column positions: 10=Overtime@1.25, 12=Overtime@1.50, 14=Overtime@2.00, 16=NightHours")
            
            # Extract data rows
            data_rows = df.iloc[employee_start_row + 1:]
            
            for idx, row in data_rows.iterrows():
                # Skip completely empty rows
                if row.isna().all():
                    continue
                
                employee_record = {}
                
                # Extract basic data
                if 'employee_no' in column_mapping:
                    emp_no = row.iloc[column_mapping['employee_no']]
                    employee_record['employee_no'] = str(int(emp_no)) if not pd.isna(emp_no) and emp_no != '' else ''
                
                if 'employee_name' in column_mapping:
                    emp_name = row.iloc[column_mapping['employee_name']]
                    employee_record['employee_name'] = str(emp_name) if not pd.isna(emp_name) else ''
                
                if 'attendance_type' in column_mapping:
                    att_type = row.iloc[column_mapping['attendance_type']]
                    employee_record['attendance_type'] = str(att_type) if not pd.isna(att_type) else 'Regular'
                else:
                    employee_record['attendance_type'] = 'Regular'
                
                if 'attendance_days' in column_mapping:
                    att_days = row.iloc[column_mapping['attendance_days']]
                    employee_record['attendance_days'] = float(att_days) if not pd.isna(att_days) else 0
                
                # Extract attendance details from fixed columns
                attendance_details = {}
                
                # Add present days if available
                if employee_record.get('attendance_days', 0) > 0:
                    attendance_details['Present'] = int(employee_record['attendance_days'])
                
                # Extract overtime from specific columns with proper key mapping (adjusted -1)
                overtime_key_mapping = {
                    11: 'overtime_125',    # Column K - OVERTIME @ 1.25
                    13: 'overtime_150',    # Column M - OVERTIME @ 1.50  
                    15: 'overtime_200',    # Column O - OVERTIME @ 2.00
                    17: 'night_hours',     # Column Q - Night Hours
                }
                
                for col_idx, xml_key in overtime_key_mapping.items():
                    if col_idx < len(row):
                        cell_value = row.iloc[col_idx]
                        if not pd.isna(cell_value) and cell_value != 0 and cell_value != '':
                            try:
                                # Store the raw decimal value for XML processing
                                decimal_hours = float(cell_value)
                                if decimal_hours > 0:
                                    attendance_details[xml_key] = decimal_hours
                            except (ValueError, TypeError):
                                # If conversion fails, store as string
                                attendance_details[xml_key] = str(cell_value)
                
                if attendance_details:
                    employee_record['attendance_details'] = attendance_details
                
                # Only add record if it has at least employee name or number
                if employee_record.get('employee_name') or employee_record.get('employee_no'):
                    employee_data.append(employee_record)
                    print(f"Added employee: {employee_record.get('employee_name', 'Unknown')} with details: {attendance_details}")
            
        except Exception as e:
            print(f"❌ Error extracting employee data: {str(e)}")
        
        return employee_data
    
    def process_attendance_sheet(self, file_path: str) -> Dict[str, Any]:
        """Process attendance sheet and extract relevant data."""
        df = self.read_excel_file(file_path)
        
        if df is None:
            return {"error": "Could not read the Excel file", "success": False}
        
        try:
            # Extract header information (Date, Company Name, Narration)
            header_info = self.extract_header_info(df)
            
            # Find where employee data starts
            employee_start_row = self.find_employee_data_start(df)
            
            # Validate the attendance structure
            validation = self.validate_attendance_structure(df, employee_start_row)
            
            if not validation["is_valid"]:
                error_msg = "Invalid attendance sheet format. " + "; ".join(validation["errors"])
                return {"error": error_msg, "success": False, "validation_errors": validation["errors"]}
            
            # Extract employee data
            employee_data = self.extract_employee_data(df, employee_start_row)
            
            # Prepare result
            result = {
                "file_name": os.path.basename(file_path),
                "sheet_type": "Attendance",
                "success": True,
                
                # Header information
                "date": header_info["date"],
                "company_name": header_info["company_name"],
                "narration": header_info["narration"],
                
                # Employee data
                "employee_data": employee_data,
                "total_employees": len(employee_data),
                
                # File statistics
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "employee_data_start_row": employee_start_row + 1,  # 1-indexed for user display
                
                # Validation info
                "validation_warnings": validation.get("warnings", [])
            }
            
            print(f"✅ Attendance processing completed: {len(employee_data)} employees found")
            return result
            
        except Exception as e:
            print(f"❌ Error in attendance processing: {str(e)}")
            return {"error": f"Error processing attendance sheet: {str(e)}", "success": False}
    
    def find_payroll_employee_data_start(self, df: pd.DataFrame) -> int:
        """Find the row where payroll employee data starts by looking for 'EMPL NO' pattern."""
        try:
            for idx, row in df.iterrows():
                # Check if any cell in this row contains 'EMPL NO'
                row_values = [str(val).strip().upper() for val in row if not pd.isna(val)]
                if any('EMPL' in val and 'NO' in val for val in row_values):
                    return idx
            
            # Fallback: look for row with 'EMPLOYEE NAME'
            for idx, row in df.iterrows():
                row_str = ' '.join([str(val) for val in row if not pd.isna(val)]).upper()
                if 'EMPLOYEE' in row_str and 'NAME' in row_str:
                    return idx
            
            return 5  # Default fallback
            
        except Exception as e:
            print(f"⚠️ Warning: Could not find payroll employee data start: {str(e)}")
            return 5
    
    def validate_payroll_structure(self, df: pd.DataFrame, employee_start_row: int) -> Dict[str, Any]:
        """Validate if the Excel file has the expected payroll structure."""
        validation_result = {
            "is_valid": False,
            "errors": [],
            "warnings": []
        }
        
        try:
            # Check if we have enough rows
            if len(df) <= employee_start_row:
                validation_result["errors"].append("File does not have enough data rows")
                return validation_result
            
            # Extract header row
            header_row = df.iloc[employee_start_row]
            headers = [str(val).strip() for val in header_row if not pd.isna(val)]
            
            # Check for required columns
            header_text = ' '.join(headers).upper()
            required_patterns = ['EMPL', 'EMPLOYEE', 'NAME']
            
            missing_patterns = []
            for pattern in required_patterns:
                if pattern not in header_text:
                    missing_patterns.append(pattern)
            
            if missing_patterns:
                validation_result["errors"].append(f"Missing expected headers: {', '.join(missing_patterns)}")
            
            # Check if we have employee data
            data_rows = df.iloc[employee_start_row + 1:]
            non_empty_data_rows = data_rows.dropna(how='all')
            
            if len(non_empty_data_rows) == 0:
                validation_result["errors"].append("No employee data found")
            elif len(non_empty_data_rows) < 2:
                validation_result["warnings"].append("Very few employee records found")
            
            # If no critical errors, mark as valid
            if not validation_result["errors"]:
                validation_result["is_valid"] = True
            
        except Exception as e:
            validation_result["errors"].append(f"Validation error: {str(e)}")
        
        return validation_result
    
    def extract_payroll_employee_data(self, df: pd.DataFrame, employee_start_row: int) -> List[Dict[str, Any]]:
        """Extract payroll employee data from the DataFrame - FULLY DYNAMIC."""
        employee_data = []
        
        try:
            # Get header row
            header_row = df.iloc[employee_start_row]
            
            # Create COMPLETELY DYNAMIC column mapping - capture ALL non-empty headers
            all_columns = []
            column_headers = []
            
            for idx, header in enumerate(header_row):
                if not pd.isna(header):
                    header_name = str(header).strip()
                    if header_name:  # Only add if header is not empty
                        all_columns.append((idx, header_name))
                        column_headers.append(header_name)
            
            print(f"🎯 DYNAMIC DEBUG: Found {len(all_columns)} columns: {column_headers}")
            
            # If we have no columns at all, this might not be a valid header row
            if not all_columns:
                print("⚠️ WARNING: No valid column headers found")
                return employee_data
            
            # Extract data rows
            data_rows = df.iloc[employee_start_row + 1:]
            
            for idx, row in data_rows.iterrows():
                # Skip completely empty rows
                if row.isna().all():
                    continue
                
                employee_record = {}
                all_data = {}
                total_salary = 0
                
                # Extract ALL column data dynamically
                for col_idx, col_name in all_columns:
                    try:
                        cell_value = row.iloc[col_idx]
                        
                        if not pd.isna(cell_value):
                            # For first 4 columns (S.NO, EMPL.NO, EMPLOYEE NAME, Attendance), keep as string
                            if col_idx < 4:
                                all_data[col_name] = str(cell_value).strip()
                            else:
                                # For salary columns, try to convert to number
                                try:
                                    numeric_value = float(str(cell_value).replace(',', ''))
                                    all_data[col_name] = numeric_value
                                except (ValueError, TypeError):
                                    # Keep as string if not numeric
                                    all_data[col_name] = str(cell_value).strip()
                        else:
                            all_data[col_name] = 0 if col_idx >= 4 else ''  # Numeric 0 for salary columns, empty string for text columns
                    except Exception as e:
                        print(f"⚠️ Warning processing column {col_name}: {str(e)}")
                        all_data[col_name] = 0 if col_idx >= 3 else ''
                
                # Set dynamic fields
                if all_columns:
                    # Use first column as employee_no and third column as employee_name (S.NO, EMPL.NO, EMPLOYEE NAME)
                    if len(all_columns) >= 2:
                        second_col_name = all_columns[1][1]  # EMPL.NO
                        employee_record['employee_no'] = str(all_data.get(second_col_name, ''))
                    
                    if len(all_columns) >= 3:
                        third_col_name = all_columns[2][1]   # EMPLOYEE NAME
                        employee_record['employee_name'] = str(all_data.get(third_col_name, ''))
                    
                    # Map Excel column names to proper XML payhead names
                    excel_to_xml_mapping = {
                        'BASIC': 'Basic Salary',
                        'HRA': 'House Rent Allowance', 
                        'MEDICAL': 'Medical Allowance',
                        'RESPONSIBILITY': 'Responsibility Allowance',
                        'FUEL': 'Fuel Allowance',
                        'FOOD': 'Food Allowance',
                        'NIGHT': 'Night Allowance',  # Map NIGHT (amount) to Night Allowance
                        'ZSSF @ 7%': 'ZSSF @ 7%',
                        'PAYE': 'PAYE',
                        'ZHSF @ 3.5%': 'ZHSF Employee @ 3.5%',
                        'ADDVANCE': 'Salary Advance',
                        'COTWU': 'COTWU'
                    }
                    
                    # Columns to exclude from payhead allocations (calculated fields)
                    excluded_columns = {
                        'GROSS', 'TAXABLE INCOME', 'TOTAL', 'NET', 'ATTENDANCE', 'OT HOURS', 'NIGHT DUTY'
                    }
                    
                    # All columns after the first four become salary components (skip S.NO, EMPL.NO, EMPLOYEE NAME, Attendance)
                    salary_components = {}
                    if len(all_columns) > 4:
                        # Direct position mapping for overtime columns - read from specific column indices
                        overtime_position_mapping = {
                            12: 'Overtime @ 1.25',  # Column index 12 (0-indexed)
                            14: 'Overtime @ 1.50',  # Column index 14 (0-indexed)
                            16: 'Overtime @ 2.00',  # Column index 16 (0-indexed)
                        }
                        
                        for col_idx, col_name in all_columns[4:]:  # Skip first four columns (S.NO, EMPL.NO, EMPLOYEE NAME, Attendance)
                            # Skip excluded columns (calculated fields)
                            if col_name.upper().strip() in excluded_columns:
                                continue
                            
                            # Special handling for OVERTIME ON columns - read directly from cell using row.iloc[col_idx]
                            if col_name.upper().strip() == 'OVERTIME ON' and col_idx in overtime_position_mapping:
                                xml_payhead_name = overtime_position_mapping[col_idx]
                                
                                # READ DIRECTLY FROM THE CELL AT THIS COLUMN INDEX
                                try:
                                    col_value = row.iloc[col_idx]
                                    if pd.isna(col_value):
                                        col_value = 0
                                    else:
                                        # Convert to float
                                        try:
                                            col_value = float(str(col_value).replace(',', ''))
                                        except:
                                            col_value = 0
                                    
                                    # 🐛 DEBUG: Print overtime values for specific employee
                                    if employee_record.get('employee_name') == 'Shaaban Said Khamis':
                                        print(f"🐛 DEBUG - {employee_record.get('employee_name')} - {xml_payhead_name}: col_idx={col_idx}, value={col_value}")
                                except Exception as e:
                                    print(f"⚠️ Error reading overtime column {col_idx}: {str(e)}")
                                    col_value = 0
                            else:
                                # Regular mapping for other columns - use all_data
                                col_value = all_data.get(col_name, 0)
                                xml_payhead_name = excel_to_xml_mapping.get(col_name.upper().strip(), col_name)
                            
                            # Only include numeric values that are not zero
                            if isinstance(col_value, (int, float)) and col_value != 0:
                                salary_components[xml_payhead_name] = col_value
                            elif isinstance(col_value, str) and col_value.strip():
                                # Try to convert string to number
                                try:
                                    numeric_value = float(col_value.replace(',', ''))
                                    if numeric_value != 0:
                                        salary_components[xml_payhead_name] = numeric_value
                                except (ValueError, TypeError):
                                    # Skip non-numeric values
                                    continue
                    
                    employee_record['salary_components'] = salary_components
                    employee_record['all_data'] = all_data  # Keep all data for reference
                    employee_record['column_headers'] = column_headers  # Store headers for UI
                
                employee_record['total_gross_salary'] = total_salary
                
                # Only add record if it has at least employee name or number
                if employee_record.get('employee_name') or employee_record.get('employee_no'):
                    employee_data.append(employee_record)
            
        except Exception as e:
            print(f"❌ Error extracting payroll employee data: {str(e)}")
        
        return employee_data
    
    def process_payroll_sheet(self, file_path: str) -> Dict[str, Any]:
        """Process payroll sheet and extract relevant data."""
        df = self.read_excel_file(file_path)
        
        if df is None:
            return {"error": "Could not read the Excel file", "success": False}
        
        try:
            # Extract header information (Date, Company Name, Account, Narration)
            header_info = self.extract_header_info(df)
            
            # Find where employee data starts
            employee_start_row = self.find_payroll_employee_data_start(df)
            
            # Validate the payroll structure
            validation = self.validate_payroll_structure(df, employee_start_row)
            
            if not validation["is_valid"]:
                error_msg = "Invalid payroll sheet format. " + "; ".join(validation["errors"])
                return {"error": error_msg, "success": False, "validation_errors": validation["errors"]}
            
            # Extract employee payroll data
            employee_data = self.extract_payroll_employee_data(df, employee_start_row)
            
            # Calculate summary statistics
            total_gross = sum(emp.get('total_gross_salary', 0) for emp in employee_data)
            avg_salary = total_gross / len(employee_data) if employee_data else 0
            
            # Prepare result
            result = {
                "file_name": os.path.basename(file_path),
                "sheet_type": "Payroll",
                "success": True,
                
                # Header information
                "date": header_info["date"],
                "company_name": header_info["company_name"],
                "account": header_info.get("narration", None),  # In payroll, 3rd row might be account
                "narration": header_info.get("narration", None),
                
                # Employee data
                "employee_data": employee_data,
                "total_employees": len(employee_data),
                
                # Payroll summary
                "total_gross_salary": total_gross,
                "average_salary": round(avg_salary, 2),
                
                # File statistics
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "employee_data_start_row": employee_start_row + 1,  # 1-indexed for user display
                
                # Validation info
                "validation_warnings": validation.get("warnings", [])
            }
            
            print(f"✅ Payroll processing completed: {len(employee_data)} employees found")
            return result
            
        except Exception as e:
            print(f"❌ Error in payroll processing: {str(e)}")
            return {"error": f"Error processing payroll sheet: {str(e)}", "success": False}
    
    def export_attendance_data(self, result: Dict[str, Any], output_path: str = None) -> str:
        """Export processed attendance data to JSON file."""
        try:
            if not result.get("success", False):
                raise ValueError("Cannot export unsuccessful processing result")
            
            # Prepare data for export
            export_data = {
                "file_info": {
                    "file_name": result["file_name"],
                    "processing_date": datetime.now().isoformat(),
                    "total_rows": result["total_rows"],
                    "total_columns": result["total_columns"]
                },
                "header_info": {
                    "date": result.get("date"),
                    "company_name": result.get("company_name"),
                    "narration": result.get("narration")
                },
                "employee_data": result.get("employee_data", []),
                "summary": {
                    "total_employees": result.get("total_employees", 0),
                    "employee_data_start_row": result.get("employee_data_start_row"),
                    "validation_warnings": result.get("validation_warnings", [])
                }
            }
            
            # Generate output filename if not provided
            if output_path is None:
                base_name = os.path.splitext(result["file_name"])[0]
                output_path = f"{base_name}_processed_attendance.json"
            
            # Write to JSON file
            import json
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)
            
            print(f"✅ Attendance data exported to: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"❌ Error exporting attendance data: {str(e)}")
            return ""
    
    def export_payroll_data(self, result: Dict[str, Any], output_path: str = None) -> str:
        """Export processed payroll data to JSON file."""
        try:
            if not result.get("success", False):
                raise ValueError("Cannot export unsuccessful processing result")
            
            # Prepare data for export
            export_data = {
                "file_info": {
                    "file_name": result["file_name"],
                    "processing_date": datetime.now().isoformat(),
                    "total_rows": result["total_rows"],
                    "total_columns": result["total_columns"]
                },
                "header_info": {
                    "date": result.get("date"),
                    "company_name": result.get("company_name"),
                    "account": result.get("account"),
                    "narration": result.get("narration")
                },
                "employee_payroll_data": result.get("employee_data", []),
                "summary": {
                    "total_employees": result.get("total_employees", 0),
                    "total_gross_salary": result.get("total_gross_salary", 0),
                    "average_salary": result.get("average_salary", 0),
                    "employee_data_start_row": result.get("employee_data_start_row"),
                    "validation_warnings": result.get("validation_warnings", [])
                }
            }
            
            # Generate output filename if not provided
            if output_path is None:
                base_name = os.path.splitext(result["file_name"])[0]
                output_path = f"{base_name}_processed_payroll.json"
            
            # Write to JSON file
            import json
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)
            
            print(f"✅ Payroll data exported to: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"❌ Error exporting payroll data: {str(e)}")
            return ""
    
    def get_sheet_names(self, file_path: str) -> list:
        """Get all sheet names from an Excel file."""
        try:
            if not self.validate_file(file_path):
                return []
            
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
            
        except Exception as e:
            print(f"Error getting sheet names: {str(e)}")
            return []
    
    def generate_payroll_xml(self, result: Dict[str, Any], company_name: str = None, account_name: str = "Cash", narration: str = None, voucher_date: str = None) -> str:
        """Generate Tally XML for payroll voucher from processed payroll data - Proper Payroll Voucher Format."""
        try:
            if not result.get("success", False):
                raise ValueError("Invalid result data")
            
            employee_data = result.get("employee_data", [])
            if not employee_data:
                raise ValueError("No employee data found")
            
            # Extract header info - use provided voucher_date if available, otherwise fallback to result date
            voucher_date = voucher_date or result.get("date", "2025-12-01")
            comp_name = company_name or result.get("company_name", "")
            narration_text = narration or result.get("narration", "Payroll for December 2025")
            
            # Format date for Tally (YYYYMMDD)
            if isinstance(voucher_date, str):
                try:
                    from datetime import datetime
                    if '-' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%Y-%m-%d")
                    else:
                        dt = datetime.strptime(voucher_date, "%d-%m-%Y")
                    formatted_date = dt.strftime("%Y%m%d")
                except:
                    formatted_date = "20251230"
            else:
                formatted_date = "20251230"
            
            # Calculate total net payable amount
            total_net_amount = 0
            for emp in employee_data:
                # Calculate net salary for each employee (total gross - deductions)
                salary_components = emp.get('salary_components', {})
                gross_total = 0
                deductions_total = 0
                
                for component_name, amount in salary_components.items():
                    amount_val = float(amount) if amount else 0
                    # Deductions are typically ZSSF, ZHSF, PAYE, COTWU, Advances
                    if any(x in component_name.upper() for x in ['ZSSF', 'ZHSF', 'PAYE', 'COTWU', 'ADVANCE']):
                        deductions_total += abs(amount_val)
                    else:
                        gross_total += abs(amount_val)
                
                net_salary = gross_total - deductions_total
                total_net_amount += net_salary
            
            # Start building XML for proper Payroll voucher
            xml_lines = [
                '<?xml version="1.0" encoding="UTF-8"?>',
                '<ENVELOPE>',
                ' <HEADER>',
                '  <TALLYREQUEST>Import Data</TALLYREQUEST>',
                ' </HEADER>',
                ' <BODY>',
                '  <IMPORTDATA>',
                '   <REQUESTDESC>',
                '    <REPORTNAME>Vouchers</REPORTNAME>',
                '    <STATICVARIABLES>',
                f'     <SVCURRENTCOMPANY>{comp_name}</SVCURRENTCOMPANY>',
                '    </STATICVARIABLES>',
                '   </REQUESTDESC>',
                '   <REQUESTDATA>',
                '    <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                '     <VOUCHER REMOTEID="" VCHKEY="" VCHTYPE="Payroll" ACTION="Create" OBJVIEW="PaySlip Voucher View">',
                '      <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '      </OLDAUDITENTRYIDS.LIST>',
                f'      <DATE>{formatted_date}</DATE>',
                f'      <NARRATION>{narration_text}</NARRATION>',
                f'      <PARTYLEDGERNAME>{account_name}</PARTYLEDGERNAME>',
                '      <VOUCHERTYPENAME>Payroll</VOUCHERTYPENAME>',
                '      <VOUCHERNUMBER>1</VOUCHERNUMBER>',
                '      <CSTFORMISSUETYPE/>',
                '      <CSTFORMRECVTYPE/>',
                '      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>',
                '      <PERSISTEDVIEW>PaySlip Voucher View</PERSISTEDVIEW>',
                '      <VCHGSTCLASS/>',
                '      <ENTEREDBY>Administrator</ENTEREDBY>',
                '      <DIFFACTUALQTY>No</DIFFACTUALQTY>',
                '      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>',
                '      <ASORIGINAL>No</ASORIGINAL>',
                '      <AUDITED>No</AUDITED>',
                '      <FORJOBCOSTING>No</FORJOBCOSTING>',
                '      <ISOPTIONAL>No</ISOPTIONAL>',
                f'      <EFFECTIVEDATE>{formatted_date}</EFFECTIVEDATE>',
                '      <USEFOREXCISE>No</USEFOREXCISE>',
                '      <ISFORJOBWORKIN>No</ISFORJOBWORKIN>',
                '      <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>',
                '      <USEFORINTEREST>No</USEFORINTEREST>',
                '      <USEFORGAINLOSS>No</USEFORGAINLOSS>',
                '      <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>',
                '      <USEFORCOMPOUND>No</USEFORCOMPOUND>',
                '      <USEFORSERVICETAX>No</USEFORSERVICETAX>',
                '      <ISDELETED>No</ISDELETED>',
                '      <ISONHOLD>No</ISONHOLD>',
                '      <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>',
                '      <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>',
                '      <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>',
                '      <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>',
                '      <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>',
                '      <EXCISEOPENING>No</EXCISEOPENING>',
                '      <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>',
                '      <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>',
                '      <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>',
                '      <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>',
                '      <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>',
                '      <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>',
                '      <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>',
                '      <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>',
                '      <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>',
                '      <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>',
                '      <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>',
                '      <ISISDVOUCHER>No</ISISDVOUCHER>',
                '      <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>',
                '      <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>',
                '      <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>',
                '      <GSTNOTEXPORTED>No</GSTNOTEXPORTED>',
                '      <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>',
                '      <ISGSTREFUND>No</ISGSTREFUND>',
                '      <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>',
                '      <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>',
                '      <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>',
                '      <ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>',
                '      <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>',
                '      <ISCANCELLED>No</ISCANCELLED>',
                '      <HASCASHFLOW>Yes</HASCASHFLOW>',
                '      <ISPOSTDATED>No</ISPOSTDATED>',
                '      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>',
                '      <ISINVOICE>No</ISINVOICE>',
                '      <MFGJOURNAL>No</MFGJOURNAL>',
                '      <HASDISCOUNTS>No</HASDISCOUNTS>',
                '      <ASPAYSLIP>Yes</ASPAYSLIP>',
                '      <ISCOSTCENTRE>No</ISCOSTCENTRE>',
                '      <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>',
                '      <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>',
                '      <ISBLANKCHEQUE>No</ISBLANKCHEQUE>',
                '      <ISVOID>No</ISVOID>',
                '      <ORDERLINESTATUS>No</ORDERLINESTATUS>',
                '      <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>',
                '      <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>',
                '      <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>',
                '      <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>',
                '      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>',
                '      <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>',
                '      <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>',
                '      <CHANGEVCHMODE>No</CHANGEVCHMODE>',
                '      <EWAYBILLDETAILS.LIST>      </EWAYBILLDETAILS.LIST>',
                '      <EXCLUDEDTAXATIONS.LIST>      </EXCLUDEDTAXATIONS.LIST>',
                '      <OLDAUDITENTRIES.LIST>      </OLDAUDITENTRIES.LIST>',
                '      <ACCOUNTAUDITENTRIES.LIST>      </ACCOUNTAUDITENTRIES.LIST>',
                '      <AUDITENTRIES.LIST>      </AUDITENTRIES.LIST>',
                '      <DUTYHEADDETAILS.LIST>      </DUTYHEADDETAILS.LIST>',
                '      <SUPPLEMENTARYDUTYHEADDETAILS.LIST>      </SUPPLEMENTARYDUTYHEADDETAILS.LIST>',
                '      <INVOICEDELNOTES.LIST>      </INVOICEDELNOTES.LIST>',
                '      <INVOICEORDERLIST.LIST>      </INVOICEORDERLIST.LIST>',
                '      <INVOICEINDENTLIST.LIST>      </INVOICEINDENTLIST.LIST>',
                '      <ATTENDANCEENTRIES.LIST>      </ATTENDANCEENTRIES.LIST>',
                '      <ORIGINVOICEDETAILS.LIST>      </ORIGINVOICEDETAILS.LIST>',
                '      <INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>'
            ]
            
            # Add LEDGERENTRIES.LIST with bank account entry
            xml_lines.extend([
                '      <LEDGERENTRIES.LIST>',
                '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '       </OLDAUDITENTRYIDS.LIST>',
                f'       <LEDGERNAME>{account_name}</LEDGERNAME>',
                '       <GSTCLASS/>',
                '       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                '       <LEDGERFROMITEM>No</LEDGERFROMITEM>',
                '       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>',
                '       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>',
                '       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>',
                '       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>',
                '       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>',
                f'       <AMOUNT>{total_net_amount:.2f}</AMOUNT>',
                '       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>',
                '       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>',
                '       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>',
                '       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>',
                '       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>',
                '       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>',
                '       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>',
                '       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>',
                '       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>',
                '       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>',
                '       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>',
                '       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>',
                '       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>',
                '       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>',
                '       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>',
                '       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>',
                '       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>',
                '       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>',
                '       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>',
                '       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>',
                '       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>',
                '       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>',
                '      </LEDGERENTRIES.LIST>'
            ])
            
            # Add CATEGORYENTRY.LIST with employee entries and payhead allocations
            xml_lines.extend([
                '      <CATEGORYENTRY.LIST>',
                '       <CATEGORY>Primary Cost Category</CATEGORY>'
            ])
            
            # Add each employee with their salary components
            employee_sort_order = 1
            for emp in employee_data:
                emp_name = emp.get('employee_name', 'Unknown Employee')
                salary_components = emp.get('salary_components', {})
                
                # Calculate employee net amount
                gross_total = 0
                deductions_total = 0
                for component_name, amount in salary_components.items():
                    amount_val = float(amount) if amount else 0
                    if any(x in component_name.upper() for x in ['ZSSF', 'ZHSF', 'PAYE', 'COTWU', 'ADVANCE']):
                        deductions_total += abs(amount_val)
                    else:
                        gross_total += abs(amount_val)
                
                employee_net = -(gross_total - deductions_total)  # Negative for payroll
                
                xml_lines.extend([
                    '       <EMPLOYEEENTRIES.LIST>',
                    f'        <EMPLOYEENAME>{emp_name}</EMPLOYEENAME>',
                    f'        <EMPLOYEESORTORDER> {employee_sort_order}</EMPLOYEESORTORDER>',
                    f'        <AMOUNT>{employee_net:.2f}</AMOUNT>'
                ])
                
                # Add payhead allocations for each salary component
                payhead_sort_order = 1
                for component_name, amount in salary_components.items():
                    if amount and float(amount) != 0:
                        amount_val = float(amount)
                        
                        # Determine if it's a deduction or earning
                        is_deduction = any(x in component_name.upper() for x in ['ZSSF', 'ZHSF', 'PAYE', 'COTWU', 'ADVANCE'])
                        is_deemed_positive = "No" if is_deduction else "Yes"
                        formatted_amount = f"{amount_val:.2f}" if is_deduction else f"-{amount_val:.2f}"
                        
                        xml_lines.extend([
                            '        <PAYHEADALLOCATIONS.LIST>',
                            f'         <PAYHEADNAME>{component_name}</PAYHEADNAME>',
                            f'         <ISDEEMEDPOSITIVE>{is_deemed_positive}</ISDEEMEDPOSITIVE>',
                            f'         <PAYHEADSORTORDER> {payhead_sort_order}</PAYHEADSORTORDER>',
                            f'         <AMOUNT>{formatted_amount}</AMOUNT>',
                            '        </PAYHEADALLOCATIONS.LIST>'
                        ])
                        payhead_sort_order += 1
                
                xml_lines.append('       </EMPLOYEEENTRIES.LIST>')
                employee_sort_order += 1
            
            # Close CATEGORYENTRY.LIST and finish voucher
            xml_lines.extend([
                '      </CATEGORYENTRY.LIST>',
                '      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>',
                '      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>',
                '      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>',
                '      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>',
                '      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>',
                '     </VOUCHER>',
                '    </TALLYMESSAGE>',
                '   </REQUESTDATA>',
                '  </IMPORTDATA>',
                ' </BODY>',
                '</ENVELOPE>'
            ])
            
            xml_content = '\n'.join(xml_lines)
            print(f"✅ Generated Payroll XML for {len(employee_data)} employees, Net Payable: ₹{total_net_amount:.2f}")
            return xml_content
            
        except Exception as e:
            print(f"❌ Error generating payroll XML: {str(e)}")
            return ""
    
    def generate_attendance_xml(self, result: Dict[str, Any], company_name: str = None, narration: str = None, voucher_date: str = None) -> str:
        """Generate Tally XML for attendance voucher from processed attendance data - Proper Attendance Voucher Format."""
        try:
            if not result.get("success", False):
                raise ValueError("Invalid result data")
            
            employee_data = result.get("employee_data", [])
            if not employee_data:
                raise ValueError("No employee data found")
            
            # Extract header info - use provided voucher_date if available, otherwise fallback to result date
            voucher_date = voucher_date or result.get("date", "2025-12-01")
            comp_name = company_name or result.get("company_name", "")
            narration_text = narration or result.get("narration", "Attendance for December 2025")
            
            # Format date for Tally (YYYYMMDD)
            if isinstance(voucher_date, str):
                try:
                    from datetime import datetime
                    if '-' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%Y-%m-%d")
                    else:
                        dt = datetime.strptime(voucher_date, "%d-%m-%Y")
                    formatted_date = dt.strftime("%Y%m%d")
                except:
                    formatted_date = "20251201"
            else:
                formatted_date = "20251201"
            
            # Start building XML for proper Attendance voucher
            xml_lines = [
                '<?xml version="1.0" encoding="UTF-8"?>',
                '<ENVELOPE>',
                ' <HEADER>',
                '  <TALLYREQUEST>Import Data</TALLYREQUEST>',
                ' </HEADER>',
                ' <BODY>',
                '  <IMPORTDATA>',
                '   <REQUESTDESC>',
                '    <REPORTNAME>Vouchers</REPORTNAME>',
                '    <STATICVARIABLES>',
                f'     <SVCURRENTCOMPANY>{comp_name}</SVCURRENTCOMPANY>',
                '    </STATICVARIABLES>',
                '   </REQUESTDESC>',
                '   <REQUESTDATA>',
                '    <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                '     <VOUCHER REMOTEID="" VCHKEY="" VCHTYPE="Attendance" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                '      <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '      </OLDAUDITENTRYIDS.LIST>',
                f'      <DATE>{formatted_date}</DATE>',
                f'      <NARRATION>{narration_text}</NARRATION>',
                '      <VOUCHERTYPENAME>Attendance</VOUCHERTYPENAME>',
                '      <VOUCHERNUMBER>1</VOUCHERNUMBER>',
                '      <CSTFORMISSUETYPE/>',
                '      <CSTFORMRECVTYPE/>',
                '      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>',
                '      <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>',
                '      <VCHGSTCLASS/>',
                '      <ENTEREDBY>Administrator</ENTEREDBY>',
                '      <DIFFACTUALQTY>No</DIFFACTUALQTY>',
                '      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>',
                '      <ASORIGINAL>No</ASORIGINAL>',
                '      <AUDITED>No</AUDITED>',
                '      <FORJOBCOSTING>No</FORJOBCOSTING>',
                '      <ISOPTIONAL>No</ISOPTIONAL>',
                f'      <EFFECTIVEDATE>{formatted_date}</EFFECTIVEDATE>',
                '      <USEFOREXCISE>No</USEFOREXCISE>',
                '      <ISFORJOBWORKIN>No</ISFORJOBWORKIN>',
                '      <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>',
                '      <USEFORINTEREST>No</USEFORINTEREST>',
                '      <USEFORGAINLOSS>No</USEFORGAINLOSS>',
                '      <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>',
                '      <USEFORCOMPOUND>No</USEFORCOMPOUND>',
                '      <USEFORSERVICETAX>No</USEFORSERVICETAX>',
                '      <ISDELETED>No</ISDELETED>',
                '      <ISONHOLD>No</ISONHOLD>',
                '      <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>',
                '      <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>',
                '      <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>',
                '      <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>',
                '      <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>',
                '      <EXCISEOPENING>No</EXCISEOPENING>',
                '      <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>',
                '      <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>',
                '      <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>',
                '      <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>',
                '      <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>',
                '      <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>',
                '      <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>',
                '      <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>',
                '      <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>',
                '      <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>',
                '      <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>',
                '      <ISISDVOUCHER>No</ISISDVOUCHER>',
                '      <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>',
                '      <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>',
                '      <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>',
                '      <GSTNOTEXPORTED>No</GSTNOTEXPORTED>',
                '      <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>',
                '      <ISGSTREFUND>No</ISGSTREFUND>',
                '      <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>',
                '      <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>',
                '      <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>',
                '      <ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>',
                '      <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>',
                '      <ISCANCELLED>No</ISCANCELLED>',
                '      <HASCASHFLOW>No</HASCASHFLOW>',
                '      <ISPOSTDATED>No</ISPOSTDATED>',
                '      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>',
                '      <ISINVOICE>No</ISINVOICE>',
                '      <MFGJOURNAL>No</MFGJOURNAL>',
                '      <HASDISCOUNTS>No</HASDISCOUNTS>',
                '      <ASPAYSLIP>No</ASPAYSLIP>',
                '      <ISCOSTCENTRE>No</ISCOSTCENTRE>',
                '      <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>',
                '      <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>',
                '      <ISBLANKCHEQUE>No</ISBLANKCHEQUE>',
                '      <ISVOID>No</ISVOID>',
                '      <ORDERLINESTATUS>No</ORDERLINESTATUS>',
                '      <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>',
                '      <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>',
                '      <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>',
                '      <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>',
                '      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>',
                '      <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>',
                '      <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>',
                '      <CHANGEVCHMODE>No</CHANGEVCHMODE>',
                '      <EWAYBILLDETAILS.LIST>      </EWAYBILLDETAILS.LIST>',
                '      <EXCLUDEDTAXATIONS.LIST>      </EXCLUDEDTAXATIONS.LIST>',
                '      <OLDAUDITENTRIES.LIST>      </OLDAUDITENTRIES.LIST>',
                '      <ACCOUNTAUDITENTRIES.LIST>      </ACCOUNTAUDITENTRIES.LIST>',
                '      <AUDITENTRIES.LIST>      </AUDITENTRIES.LIST>',
                '      <DUTYHEADDETAILS.LIST>      </DUTYHEADDETAILS.LIST>',
                '      <SUPPLEMENTARYDUTYHEADDETAILS.LIST>      </SUPPLEMENTARYDUTYHEADDETAILS.LIST>',
                '      <INVOICEDELNOTES.LIST>      </INVOICEDELNOTES.LIST>',
                '      <INVOICEORDERLIST.LIST>      </INVOICEORDERLIST.LIST>',
                '      <INVOICEINDENTLIST.LIST>      </INVOICEINDENTLIST.LIST>'
            ]
            
            # Add attendance entries for each employee
            for emp in employee_data:
                emp_name = emp.get('employee_name', 'Unknown Employee')
                attendance_days = emp.get('attendance_days', 0)
                attendance_details = emp.get('attendance_details', {})
                
                # Add Present attendance entry (main attendance days)
                present_days = attendance_details.get('present_days', attendance_days)
                if present_days > 0:
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        '       <ATTENDANCETYPE>Present</ATTENDANCETYPE>',
                        f'       <ATTDTYPETIMEVALUE> {int(present_days)}</ATTDTYPETIMEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
                
                # Add overtime entries if available
                if 'overtime_125' in attendance_details and attendance_details['overtime_125']:
                    overtime_val = attendance_details['overtime_125']
                    # Format as "X Hrs Y Mins" if it's a number, otherwise use as is
                    if isinstance(overtime_val, (int, float)):
                        hours = int(overtime_val)
                        mins = int((overtime_val - hours) * 60)
                        overtime_formatted = f"{hours} Hrs {mins} Mins"
                    else:
                        overtime_formatted = str(overtime_val)
                    
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        '       <ATTENDANCETYPE>Overtime @ 1.25</ATTENDANCETYPE>',
                        f'       <ATTDTYPEVALUE> {overtime_formatted}</ATTDTYPEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
                
                if 'overtime_150' in attendance_details and attendance_details['overtime_150']:
                    overtime_val = attendance_details['overtime_150']
                    if isinstance(overtime_val, (int, float)):
                        hours = int(overtime_val)
                        mins = int((overtime_val - hours) * 60)
                        overtime_formatted = f"{hours} Hrs {mins} Mins"
                    else:
                        overtime_formatted = str(overtime_val)
                    
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        '       <ATTENDANCETYPE>Overtime @ 1.50</ATTENDANCETYPE>',
                        f'       <ATTDTYPEVALUE> {overtime_formatted}</ATTDTYPEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
                
                if 'overtime_200' in attendance_details and attendance_details['overtime_200']:
                    overtime_val = attendance_details['overtime_200']
                    if isinstance(overtime_val, (int, float)):
                        hours = int(overtime_val)
                        mins = int((overtime_val - hours) * 60)
                        overtime_formatted = f"{hours} Hrs {mins} Mins"
                    else:
                        overtime_formatted = str(overtime_val)
                    
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        '       <ATTENDANCETYPE>Overtime @ 2.00</ATTENDANCETYPE>',
                        f'       <ATTDTYPEVALUE> {overtime_formatted}</ATTDTYPEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
                
                # Add overtime normal days if available
                if 'overtime_normal' in attendance_details and attendance_details['overtime_normal']:
                    normal_val = attendance_details['overtime_normal']
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        '       <ATTENDANCETYPE>Overtime Normal Days</ATTENDANCETYPE>',
                        f'       <ATTDTYPEVALUE> {normal_val}</ATTDTYPEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
                
                # Add overtime weekends if available
                if 'overtime_weekends' in attendance_details and attendance_details['overtime_weekends']:
                    weekend_val = attendance_details['overtime_weekends']
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        '       <ATTENDANCETYPE>Overtime Weekends</ATTENDANCETYPE>',
                        f'       <ATTDTYPEVALUE> {weekend_val}</ATTDTYPEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
                
                # Add night hours if available
                if 'night_hours' in attendance_details and attendance_details['night_hours']:
                    night_val = attendance_details['night_hours']
                    if isinstance(night_val, (int, float)):
                        hours = int(night_val)
                        mins = int((night_val - hours) * 60)
                        night_formatted = f"{hours} Hrs {mins} Mins"
                    else:
                        night_formatted = str(night_val)
                    
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        '       <ATTENDANCETYPE>Night Hours</ATTENDANCETYPE>',
                        f'       <ATTDTYPEVALUE> {night_formatted}</ATTDTYPEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
            
            # Close the XML structure
            xml_lines.extend([
                '      <ORIGINVOICEDETAILS.LIST>      </ORIGINVOICEDETAILS.LIST>',
                '      <INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>',
                '      <ALLLEDGERENTRIES.LIST>      </ALLLEDGERENTRIES.LIST>',
                '      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>',
                '      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>',
                '      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>',
                '      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>',
                '      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>',
                '     </VOUCHER>',
                '    </TALLYMESSAGE>',
                '   </REQUESTDATA>',
                '  </IMPORTDATA>',
                ' </BODY>',
                '</ENVELOPE>'
            ])
            
            xml_content = '\n'.join(xml_lines)
            total_days = sum(float(emp.get('attendance_days', 0)) for emp in employee_data)
            print(f"✅ Generated Proper Attendance XML for {len(employee_data)} employees, Total Days: {total_days}")
            return xml_content
            
        except Exception as e:
            print(f"❌ Error generating attendance XML: {str(e)}")
            return ""
    
    def process_paye_sheet(self, file_path: str) -> Dict[str, Any]:
        """Process PAYE sheet (Sheet 2) and extract PAYE/SDL data."""
        # Read the second sheet (index 1)
        df = self.read_excel_file(file_path, sheet_name=1)
        
        if df is None:
            return {"error": "Could not read the Excel file", "success": False}
        
        try:
            print(f"📊 Processing PAYE sheet: {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Extract header information from first few rows
            header_info = self.extract_header_info(df)
            
            # Find where employee data starts - look for row with EMPLOYEE NAME
            employee_start_row = self.find_paye_employee_data_start(df)
            
            # Extract PAYE employee data
            employee_data = self.extract_paye_employee_data(df, employee_start_row)
            
            # Calculate totals
            total_paye = sum(emp.get('paye', 0) for emp in employee_data)
            total_sdl = sum(emp.get('sdl', 0) for emp in employee_data)
            total_amount = total_paye + total_sdl
            
            # Prepare result
            result = {
                "file_name": os.path.basename(file_path),
                "sheet_type": "PAYE",
                "success": True,
                "date": header_info["date"],
                "company_name": header_info["company_name"],
                "narration": header_info.get("narration", "PAYE and SDL for current period"),
                "employee_data": employee_data,
                "total_employees": len(employee_data),
                "total_paye": total_paye,
                "total_sdl": total_sdl,
                "total_amount": total_amount,
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "employee_data_start_row": employee_start_row + 1
            }
            
            print(f"✅ PAYE processing completed: {len(employee_data)} employees, PAYE: ₹{total_paye:,.2f}, SDL: ₹{total_sdl:,.2f}")
            return result
            
        except Exception as e:
            print(f"❌ Error in PAYE processing: {str(e)}")
            return {"error": f"Error processing PAYE sheet: {str(e)}", "success": False}

    def find_paye_employee_data_start(self, df: pd.DataFrame) -> int:
        """Find the row where PAYE employee data starts."""
        try:
            for idx, row in df.iterrows():
                row_str = ' '.join([str(val).upper() for val in row if not pd.isna(val)])
                if 'EMPLOYEE' in row_str and 'NAME' in row_str:
                    return idx
            
            return 5  # Default fallback
            
        except Exception as e:
            print(f"⚠️ Warning: Could not find PAYE employee data start: {str(e)}")
            return 5

    def extract_paye_employee_data(self, df: pd.DataFrame, employee_start_row: int) -> List[Dict[str, Any]]:
        """Extract PAYE employee data from the DataFrame."""
        employee_data = []
        
        try:
            # Get header row
            header_row = df.iloc[employee_start_row]
            
            # Create column mapping for PAYE sheet
            column_mapping = {}
            for idx, header in enumerate(header_row):
                if not pd.isna(header):
                    header_str = str(header).strip().upper()
                    if 'EMPLOYEE' in header_str and 'NAME' in header_str:
                        column_mapping['employee_name'] = idx
                    elif header_str == 'PAYE':
                        column_mapping['paye'] = idx
                    elif header_str == 'SDL':
                        column_mapping['sdl'] = idx
                    elif 'GROSS' in header_str and 'SALARY' in header_str:
                        column_mapping['gross_salary'] = idx
                    elif header_str == 'BASIC':
                        column_mapping['basic'] = idx
                    elif 'ALLOW' in header_str:
                        column_mapping['allowance'] = idx
                    elif 'DEDUCTION' in header_str:
                        column_mapping['deduction'] = idx
                    elif 'TIN' in header_str:
                        column_mapping['tin'] = idx
                    elif 'ZSSF' in header_str and '#' in header_str:
                        column_mapping['zssf_no'] = idx
                    elif header_str == 'NAME':
                        column_mapping['name'] = idx
                    elif header_str == 'ZSSF':
                        column_mapping['zssf'] = idx
                    elif header_str == 'ZHSF':
                        column_mapping['zhsf'] = idx
            
            print(f"🎯 PAYE Column mapping: {column_mapping}")
            
            # Extract data rows
            data_rows = df.iloc[employee_start_row + 1:]
            
            for idx, row in data_rows.iterrows():
                # Skip empty rows
                if row.isna().all():
                    continue
                
                employee_record = {}
                
                # Extract employee data using column mapping
                for field, col_idx in column_mapping.items():
                    if col_idx < len(row):
                        value = row.iloc[col_idx]
                        if not pd.isna(value):
                            if field in ['paye', 'sdl', 'gross_salary', 'basic', 'allowance', 'deduction', 'zssf', 'zhsf']:
                                # Convert to numeric
                                try:
                                    employee_record[field] = float(value)
                                except (ValueError, TypeError):
                                    employee_record[field] = 0
                            else:
                                # Ensure string fields are properly converted - safe handling
                                str_value = str(value).strip() if value is not None and str(value).strip() != 'nan' else ""
                                employee_record[field] = str_value
                        else:
                            employee_record[field] = 0 if field in ['paye', 'sdl', 'gross_salary', 'basic', 'allowance', 'deduction', 'zssf', 'zhsf'] else ""
                
                # Ensure employee_name is always a valid string
                if 'employee_name' in employee_record and not employee_record['employee_name']:
                    employee_record['employee_name'] = f"Employee_{idx}"
                
                # Only add if we have essential data
                if employee_record.get('employee_name') and (employee_record.get('paye', 0) > 0 or employee_record.get('sdl', 0) > 0):
                    employee_data.append(employee_record)
                    emp_name = employee_record.get('employee_name', 'Unknown')
                    print(f"Added PAYE employee: {emp_name}, PAYE: {employee_record.get('paye', 0)}, SDL: {employee_record.get('sdl', 0)}")
            
        except Exception as e:
            print(f"❌ Error extracting PAYE employee data: {str(e)}")
        
        return employee_data

    def generate_paye_xml(self, result: Dict[str, Any], company_name: str = None, account_name: str = "THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS", narration: str = None, voucher_date: str = None) -> str:
        """Generate Tally XML for PAYE voucher from processed PAYE data - Payment Voucher Format."""
        try:
            if not result.get("success", False):
                raise ValueError("Invalid result data")
            
            employee_data = result.get("employee_data", [])
            if not employee_data:
                raise ValueError("No employee data found")
            
            # Extract header info with safe string handling - use provided voucher_date if available
            voucher_date = voucher_date or result.get("date", "2025-12-01")
            comp_name = company_name or result.get("company_name", "TEST COMPANY")
            narration_text = narration or result.get("narration", "PAYE and SDL for December 2025")
            
            # Ensure all string values are not None
            if comp_name is None or comp_name == "":
                comp_name = "TEST COMPANY"
            if narration_text is None or narration_text == "":
                narration_text = "PAYE and SDL for December 2025"
            if account_name is None or account_name == "":
                account_name = "THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS"
            
            # Format date for Tally (YYYYMMDD)
            if isinstance(voucher_date, str):
                try:
                    from datetime import datetime
                    if '-' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%Y-%m-%d")
                    elif '/' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%d/%m/%Y")
                    else:
                        dt = datetime.strptime("30/12/2025", "%d/%m/%Y")
                    formatted_date = dt.strftime("%Y%m%d")
                except:
                    formatted_date = "20251230"
            else:
                formatted_date = "20251230"
            
            # Calculate totals
            total_paye = sum(emp.get('paye', 0) for emp in employee_data)
            total_sdl = sum(emp.get('sdl', 0) for emp in employee_data)
            total_amount = total_paye + total_sdl
            
            # Start building XML for PAYE Payment voucher
            xml_lines = [
                '<?xml version="1.0" encoding="UTF-8"?>',
                '<ENVELOPE>',
                ' <HEADER>',
                '  <TALLYREQUEST>Import Data</TALLYREQUEST>',
                ' </HEADER>',
                ' <BODY>',
                '  <IMPORTDATA>',
                '   <REQUESTDESC>',
                '    <REPORTNAME>Vouchers</REPORTNAME>',
                '    <STATICVARIABLES>',
                f'     <SVCURRENTCOMPANY>{comp_name}</SVCURRENTCOMPANY>',
                '    </STATICVARIABLES>',
                '   </REQUESTDESC>',
                '   <REQUESTDATA>',
                '    <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                '     <VOUCHER REMOTEID="" VCHKEY="" VCHTYPE="Payment" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                '      <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '      </OLDAUDITENTRYIDS.LIST>',
                f'      <DATE>{formatted_date}</DATE>',
                f'      <NARRATION>{narration_text}</NARRATION>',
                f'      <PARTYLEDGERNAME>{account_name}</PARTYLEDGERNAME>',
                '      <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>',
                '      <VOUCHERNUMBER>1</VOUCHERNUMBER>',
                '      <CSTFORMISSUETYPE/>',
                '      <CSTFORMRECVTYPE/>',
                '      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>',
                '      <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>',
                '      <VCHGSTCLASS/>',
                '      <ENTEREDBY>Administrator</ENTEREDBY>',
                '      <DIFFACTUALQTY>No</DIFFACTUALQTY>',
                '      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>',
                '      <ASORIGINAL>No</ASORIGINAL>',
                '      <AUDITED>No</AUDITED>',
                '      <FORJOBCOSTING>No</FORJOBCOSTING>',
                '      <ISOPTIONAL>No</ISOPTIONAL>',
                f'      <EFFECTIVEDATE>{formatted_date}</EFFECTIVEDATE>',
                '      <USEFOREXCISE>No</USEFOREXCISE>',
                '      <ISFORJOBWORKIN>No</ISFORJOBWORKIN>',
                '      <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>',
                '      <USEFORINTEREST>No</USEFORINTEREST>',
                '      <USEFORGAINLOSS>No</USEFORGAINLOSS>',
                '      <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>',
                '      <USEFORCOMPOUND>No</USEFORCOMPOUND>',
                '      <USEFORSERVICETAX>No</USEFORSERVICETAX>',
                '      <ISDELETED>No</ISDELETED>',
                '      <ISONHOLD>No</ISONHOLD>',
                '      <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>',
                '      <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>',
                '      <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>',
                '      <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>',
                '      <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>',
                '      <EXCISEOPENING>No</EXCISEOPENING>',
                '      <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>',
                '      <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>',
                '      <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>',
                '      <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>',
                '      <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>',
                '      <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>',
                '      <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>',
                '      <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>',
                '      <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>',
                '      <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>',
                '      <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>',
                '      <ISISDVOUCHER>No</ISISDVOUCHER>',
                '      <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>',
                '      <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>',
                '      <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>',
                '      <GSTNOTEXPORTED>No</GSTNOTEXPORTED>',
                '      <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>',
                '      <ISGSTREFUND>No</ISGSTREFUND>',
                '      <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>',
                '      <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>',
                '      <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>',
                '      <ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>',
                '      <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>',
                '      <ISCANCELLED>No</ISCANCELLED>',
                '      <HASCASHFLOW>Yes</HASCASHFLOW>',
                '      <ISPOSTDATED>No</ISPOSTDATED>',
                '      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>',
                '      <ISINVOICE>No</ISINVOICE>',
                '      <MFGJOURNAL>No</MFGJOURNAL>',
                '      <HASDISCOUNTS>No</HASDISCOUNTS>',
                '      <ASPAYSLIP>No</ASPAYSLIP>',
                '      <ISCOSTCENTRE>No</ISCOSTCENTRE>',
                '      <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>',
                '      <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>',
                '      <ISBLANKCHEQUE>No</ISBLANKCHEQUE>',
                '      <ISVOID>No</ISVOID>',
                '      <ORDERLINESTATUS>No</ORDERLINESTATUS>',
                '      <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>',
                '      <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>',
                '      <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>',
                '      <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>',
                '      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>',
                '      <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>',
                '      <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>',
                '      <CHANGEVCHMODE>No</CHANGEVCHMODE>',
                '      <EWAYBILLDETAILS.LIST>      </EWAYBILLDETAILS.LIST>',
                '      <EXCLUDEDTAXATIONS.LIST>      </EXCLUDEDTAXATIONS.LIST>',
                '      <OLDAUDITENTRIES.LIST>      </OLDAUDITENTRIES.LIST>',
                '      <ACCOUNTAUDITENTRIES.LIST>      </ACCOUNTAUDITENTRIES.LIST>',
                '      <AUDITENTRIES.LIST>      </AUDITENTRIES.LIST>',
                '      <DUTYHEADDETAILS.LIST>      </DUTYHEADDETAILS.LIST>',
                '      <SUPPLEMENTARYDUTYHEADDETAILS.LIST>      </SUPPLEMENTARYDUTYHEADDETAILS.LIST>',
                '      <INVOICEDELNOTES.LIST>      </INVOICEDELNOTES.LIST>',
                '      <INVOICEORDERLIST.LIST>      </INVOICEORDERLIST.LIST>',
                '      <INVOICEINDENTLIST.LIST>      </INVOICEINDENTLIST.LIST>',
                '      <ATTENDANCEENTRIES.LIST>      </ATTENDANCEENTRIES.LIST>',
                '      <ORIGINVOICEDETAILS.LIST>      </ORIGINVOICEDETAILS.LIST>',
                '      <INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>'
            ]
            
            # Add PAYE ledger entry
            if total_paye > 0:
                xml_lines.extend([
                    '      <ALLLEDGERENTRIES.LIST>',
                    '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                    '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                    '       </OLDAUDITENTRYIDS.LIST>',
                    '       <LEDGERNAME>PAYE</LEDGERNAME>',
                    '       <GSTCLASS/>',
                    '       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    '       <LEDGERFROMITEM>No</LEDGERFROMITEM>',
                    '       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>',
                    '       <ISPARTYLEDGER>No</ISPARTYLEDGER>',
                    '       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>',
                    '       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>',
                    '       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>',
                    f'       <AMOUNT>-{total_paye:.2f}</AMOUNT>',
                    '       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>',
                    '       <CATEGORYALLOCATIONS.LIST>',
                    '        <CATEGORY>Primary Cost Category</CATEGORY>',
                    '        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
                ])
                
                # Add cost center allocations for each employee's PAYE
                for emp in employee_data:
                    if emp.get('paye', 0) > 0:
                        # Ensure employee name is always a string
                        emp_name = emp.get("employee_name") or "Unknown Employee"
                        if emp_name is None:
                            emp_name = "Unknown Employee"
                        xml_lines.extend([
                            '        <COSTCENTREALLOCATIONS.LIST>',
                            f'         <NAME>{emp_name}</NAME>',
                            f'         <AMOUNT>-{emp.get("paye", 0):.2f}</AMOUNT>',
                            '        </COSTCENTREALLOCATIONS.LIST>'
                        ])
                
                xml_lines.extend([
                    '       </CATEGORYALLOCATIONS.LIST>',
                    '       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>',
                    '       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>',
                    '       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>',
                    '       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>',
                    '       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>',
                    '       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>',
                    '       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>',
                    '       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>',
                    '       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>',
                    '       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>',
                    '       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>',
                    '       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>',
                    '       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>',
                    '       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>',
                    '       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>',
                    '       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>',
                    '       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>',
                    '       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>',
                    '       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>',
                    '       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>',
                    '       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>',
                    '       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>',
                    '      </ALLLEDGERENTRIES.LIST>'
                ])
            
            # Add SDL ledger entry
            if total_sdl > 0:
                xml_lines.extend([
                    '      <ALLLEDGERENTRIES.LIST>',
                    '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                    '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                    '       </OLDAUDITENTRYIDS.LIST>',
                    '       <LEDGERNAME>Skill &amp; Development Levy</LEDGERNAME>',
                    '       <GSTCLASS/>',
                    '       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    '       <LEDGERFROMITEM>No</LEDGERFROMITEM>',
                    '       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>',
                    '       <ISPARTYLEDGER>No</ISPARTYLEDGER>',
                    '       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>',
                    '       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>',
                    '       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>',
                    f'       <AMOUNT>-{total_sdl:.2f}</AMOUNT>',
                    '       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>',
                    '       <CATEGORYALLOCATIONS.LIST>',
                    '        <CATEGORY>Primary Cost Category</CATEGORY>',
                    '        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
                ])
                
                # Add cost center allocations for each employee's SDL
                for emp in employee_data:
                    if emp.get('sdl', 0) > 0:
                        # Ensure employee name is always a string
                        emp_name = emp.get("employee_name") or "Unknown Employee"
                        if emp_name is None:
                            emp_name = "Unknown Employee"
                        xml_lines.extend([
                            '        <COSTCENTREALLOCATIONS.LIST>',
                            f'         <NAME>{emp_name}</NAME>',
                            f'         <AMOUNT>-{emp.get("sdl", 0):.2f}</AMOUNT>',
                            '        </COSTCENTREALLOCATIONS.LIST>'
                        ])
                
                xml_lines.extend([
                    '       </CATEGORYALLOCATIONS.LIST>',
                    '       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>',
                    '       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>',
                    '       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>',
                    '       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>',
                    '       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>',
                    '       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>',
                    '       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>',
                    '       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>',
                    '       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>',
                    '       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>',
                    '       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>',
                    '       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>',
                    '       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>',
                    '       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>',
                    '       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>',
                    '       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>',
                    '       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>',
                    '       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>',
                    '       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>',
                    '       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>',
                    '       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>',
                    '       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>',
                    '      </ALLLEDGERENTRIES.LIST>'
                ])
            
            # Add Bank ledger entry (credit side)
            xml_lines.extend([
                '      <ALLLEDGERENTRIES.LIST>',
                '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '       </OLDAUDITENTRYIDS.LIST>',
                f'       <LEDGERNAME>{account_name}</LEDGERNAME>',
                '       <GSTCLASS/>',
                '       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                '       <LEDGERFROMITEM>No</LEDGERFROMITEM>',
                '       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>',
                '       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>',
                '       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>',
                '       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>',
                '       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>',
                f'       <AMOUNT>{total_amount:.2f}</AMOUNT>',
                '       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>',
                '       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>',
                '       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>',
                '       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>',
                '       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>',
                '       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>',
                '       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>',
                '       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>',
                '       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>',
                '       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>',
                '       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>',
                '       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>',
                '       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>',
                '       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>',
                '       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>',
                '       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>',
                '       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>',
                '       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>',
                '       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>',
                '       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>',
                '       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>',
                '       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>',
                '       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>',
                '      </ALLLEDGERENTRIES.LIST>',
                '      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>',
                '      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>',
                '      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>',
                '      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>',
                '      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>',
                '     </VOUCHER>',
                '    </TALLYMESSAGE>',
                '   </REQUESTDATA>',
                '  </IMPORTDATA>',
                ' </BODY>',
                '</ENVELOPE>'
            ])
            
            xml_content = '\n'.join(xml_lines)
            print(f"✅ Generated PAYE XML for {len(employee_data)} employees, Total Amount: ₹{total_amount:,.2f}")
            return xml_content
            
        except Exception as e:
            print(f"❌ Error generating PAYE XML: {str(e)}")
            return ""

    def save_xml_file(self, xml_content: str, file_type: str = "payroll", output_path: str = None) -> str:
        """Save XML content to file."""
        try:
            if not xml_content:
                raise ValueError("No XML content to save")
            
            if output_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                # Create appropriate filename based on file type
                if file_type == "paye":
                    filename = f"payroll_paye_import_{timestamp}.xml"
                elif file_type == "payroll":
                    filename = f"payroll_tally_import_{timestamp}.xml"
                elif file_type == "attendance":
                    filename = f"attendance_tally_import_{timestamp}.xml"
                elif file_type == "zssf":
                    filename = f"zssf_tally_import_{timestamp}.xml"
                elif file_type == "zhsf":
                    filename = f"zhsf_tally_import_{timestamp}.xml"
                else:
                    filename = f"{file_type}_tally_import_{timestamp}.xml"
                
                # Try to save to user's Documents folder first, fallback to Downloads, then temp
                possible_locations = [
                    # User Documents folder
                    os.path.join(os.path.expanduser("~"), "Documents", "Tally_XML_Files"),
                    # User Downloads folder
                    os.path.join(os.path.expanduser("~"), "Downloads"),
                    # User home directory
                    os.path.expanduser("~"),
                    # Current directory (if writable)
                    os.getcwd(),
                    # Temporary directory as last resort
                    os.path.join(os.path.expanduser("~"), "temp"),
                ]
                
                output_path = None
                for location in possible_locations:
                    try:
                        # Create directory if it doesn't exist (for Documents/Tally_XML_Files)
                        if not os.path.exists(location):
                            os.makedirs(location, exist_ok=True)
                        
                        # Test if we can write to this location
                        test_path = os.path.join(location, filename)
                        
                        # Try to write the file
                        with open(test_path, 'w', encoding='utf-8') as f:
                            f.write(xml_content)
                        
                        output_path = test_path
                        break
                        
                    except (OSError, PermissionError, IOError) as e:
                        print(f"⚠️ Cannot write to {location}: {str(e)}")
                        continue
                
                if not output_path:
                    raise ValueError("Could not find a writable location to save XML file")
            else:
                # Output path was provided, use it directly
                # But ensure the directory exists
                directory = os.path.dirname(output_path)
                if directory and not os.path.exists(directory):
                    os.makedirs(directory, exist_ok=True)
                
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(xml_content)
            
            print(f"✅ XML file saved: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"❌ Error saving XML file: {str(e)}")
            import traceback
            traceback.print_exc()
            return ""

    def generate_zhsf_xml(self, result: Dict[str, Any], company_name: str = None, account_name: str = "THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS", narration: str = None, voucher_date: str = None) -> str:
        """Generate Tally XML for ZHSF voucher from processed ZHSF data - Payment Voucher Format."""
        try:
            if not result.get("success", False):
                raise ValueError("Invalid result data")
            
            employee_data = result.get("employee_data", [])
            if not employee_data:
                raise ValueError("No employee data found")
            
            # Extract header info with safe string handling - use provided voucher_date if available
            voucher_date = voucher_date or result.get("date", "2025-12-01")
            comp_name = company_name or result.get("company_name", "TEST COMPANY")
            narration_text = narration or result.get("narration", "ZHSF for December 2025")
            
            # Ensure all string values are not None
            if comp_name is None or comp_name == "":
                comp_name = "TEST COMPANY"
            if narration_text is None or narration_text == "":
                narration_text = "ZHSF for December 2025"
            if account_name is None or account_name == "":
                account_name = "THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS"
            
            # Format date for Tally (YYYYMMDD)
            if isinstance(voucher_date, str):
                try:
                    from datetime import datetime
                    if '-' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%Y-%m-%d")
                    elif '/' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%d/%m/%Y")
                    else:
                        dt = datetime.strptime("30/12/2025", "%d/%m/%Y")
                    formatted_date = dt.strftime("%Y%m%d")
                except:
                    formatted_date = "20251230"
            else:
                formatted_date = "20251230"
            
            # Calculate totals for ZHSF - Employee 3.5% and TWA 3.5%
            total_employee_zhsf = sum(emp.get('employee_35', 0) for emp in employee_data)
            total_twa_zhsf = sum(emp.get('twa_35', 0) for emp in employee_data)
            total_amount = total_employee_zhsf + total_twa_zhsf
            
            # Build XML content
            xml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>{comp_name}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <VOUCHER REMOTEID="" VCHKEY="" VCHTYPE="Payment" ACTION="Create" OBJVIEW="Accounting Voucher View">
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>{formatted_date}</DATE>
      <NARRATION>{narration_text}</NARRATION>
      <PARTYLEDGERNAME>{account_name}</PARTYLEDGERNAME>
      <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>
      <VOUCHERNUMBER>1</VOUCHERNUMBER>
      <CSTFORMISSUETYPE/>
      <CSTFORMRECVTYPE/>
      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
      <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
      <VCHGSTCLASS/>
      <ENTEREDBY>Administrator</ENTEREDBY>
      <DIFFACTUALQTY>No</DIFFACTUALQTY>
      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>
      <ASORIGINAL>No</ASORIGINAL>
      <AUDITED>No</AUDITED>
      <FORJOBCOSTING>No</FORJOBCOSTING>
      <ISOPTIONAL>No</ISOPTIONAL>
      <EFFECTIVEDATE>{formatted_date}</EFFECTIVEDATE>
      <USEFOREXCISE>No</USEFOREXCISE>
      <ISFORJOBWORKIN>No</ISFORJOBWORKIN>
      <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>
      <USEFORINTEREST>No</USEFORINTEREST>
      <USEFORGAINLOSS>No</USEFORGAINLOSS>
      <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>
      <USEFORCOMPOUND>No</USEFORCOMPOUND>
      <USEFORSERVICETAX>No</USEFORSERVICETAX>
      <ISDELETED>No</ISDELETED>
      <ISONHOLD>No</ISONHOLD>
      <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>
      <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>
      <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>
      <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>
      <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>
      <EXCISEOPENING>No</EXCISEOPENING>
      <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>
      <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>
      <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>
      <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>
      <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>
      <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>
      <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>
      <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>
      <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>
      <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>
      <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>
      <ISISDVOUCHER>No</ISISDVOUCHER>
      <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>
      <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>
      <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>
      <GSTNOTEXPORTED>No</GSTNOTEXPORTED>
      <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>
      <ISGSTREFUND>No</ISGSTREFUND>
      <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>
      <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>
      <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>
      <ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>
      <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>
      <ISCANCELLED>No</ISCANCELLED>
      <HASCASHFLOW>Yes</HASCASHFLOW>
      <ISPOSTDATED>No</ISPOSTDATED>
      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>
      <ISINVOICE>No</ISINVOICE>
      <MFGJOURNAL>No</MFGJOURNAL>
      <HASDISCOUNTS>No</HASDISCOUNTS>
      <ASPAYSLIP>No</ASPAYSLIP>
      <ISCOSTCENTRE>No</ISCOSTCENTRE>
      <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>
      <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>
      <ISBLANKCHEQUE>No</ISBLANKCHEQUE>
      <ISVOID>No</ISVOID>
      <ORDERLINESTATUS>No</ORDERLINESTATUS>
      <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>
      <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>
      <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>
      <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>
      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>
      <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>
      <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>
      <CHANGEVCHMODE>No</CHANGEVCHMODE>
      <ALLLEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>ZHSF Employee @ 3.5%</LEDGERNAME>
       <GSTCLASS/>
       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>-{total_employee_zhsf:.2f}</AMOUNT>
       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>
       <CATEGORYALLOCATIONS.LIST>
        <CATEGORY>Primary Cost Category</CATEGORY>
        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'''

            # Add cost center allocations for each employee (Employee 3.5%)
            for emp in employee_data:
                emp_name = emp.get('employee_name', 'Unknown')
                emp_employee_zhsf = emp.get('employee_35', 0)
                if emp_employee_zhsf > 0:
                    xml_content += f'''
        <COSTCENTREALLOCATIONS.LIST>
         <NAME>{emp_name}</NAME>
         <AMOUNT>-{emp_employee_zhsf:.2f}</AMOUNT>
        </COSTCENTREALLOCATIONS.LIST>'''

            xml_content += f'''
       </CATEGORYALLOCATIONS.LIST>
       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>
       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>
       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>
       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>
       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>
       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>
       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>
       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>
       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>
       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>
       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>
       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>
       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>
       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>
       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>
       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>
       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>
      </ALLLEDGERENTRIES.LIST>
      <ALLLEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>ZHSF TWA @ 3.5%</LEDGERNAME>
       <GSTCLASS/>
       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>No</ISPARTYLEDGER>
       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>-{total_twa_zhsf:.2f}</AMOUNT>
       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>
       <CATEGORYALLOCATIONS.LIST>
        <CATEGORY>Primary Cost Category</CATEGORY>
        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'''

            # Add cost center allocations for each employee (TWA 3.5%)
            for emp in employee_data:
                emp_name = emp.get('employee_name', 'Unknown')
                emp_twa_zhsf = emp.get('twa_35', 0)
                if emp_twa_zhsf > 0:
                    xml_content += f'''
        <COSTCENTREALLOCATIONS.LIST>
         <NAME>{emp_name}</NAME>
         <AMOUNT>-{emp_twa_zhsf:.2f}</AMOUNT>
        </COSTCENTREALLOCATIONS.LIST>'''

            xml_content += f'''
       </CATEGORYALLOCATIONS.LIST>
       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>
       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>
       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>
       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>
       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>
       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>
       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>
       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>
       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>
       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>
       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>
       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>
       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>
       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>
       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>
       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>
       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>
      </ALLLEDGERENTRIES.LIST>
      <ALLLEDGERENTRIES.LIST>
       <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
       </OLDAUDITENTRYIDS.LIST>
       <LEDGERNAME>{account_name}</LEDGERNAME>
       <GSTCLASS/>
       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
       <LEDGERFROMITEM>No</LEDGERFROMITEM>
       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>
       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>
       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
       <AMOUNT>{total_amount:.2f}</AMOUNT>
       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>
       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>
       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>
       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>
       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>
       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>
       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>
       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>
       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>
       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>
       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>
       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>
       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>
       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>
       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>
       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>
       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>
       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>
      </ALLLEDGERENTRIES.LIST>
      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>
      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>
      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>
      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>
      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>
     </VOUCHER>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>'''
            
            return xml_content
            
        except Exception as e:
            print(f"❌ Error generating ZHSF XML: {str(e)}")
            return ""

    def generate_zssf_xml(self, result: Dict[str, Any], company_name: str = None, account_name: str = "THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS", narration: str = None, voucher_date: str = None) -> str:
        """Generate Tally XML for ZSSF voucher from processed ZSSF data - Payment Voucher Format."""
        try:
            if not result.get("success", False):
                raise ValueError("Invalid result data")
            
            employee_data = result.get("employee_data", [])
            if not employee_data:
                raise ValueError("No employee data found")
            
            # Extract header info with safe string handling - use provided voucher_date if available
            voucher_date = voucher_date or result.get("date", "2025-12-01")
            comp_name = company_name or result.get("company_name", "TEST COMPANY")
            narration_text = narration or result.get("narration", "ZSSF for December 2025")
            
            # Ensure all string values are not None
            if comp_name is None or comp_name == "":
                comp_name = "TEST COMPANY"
            if narration_text is None or narration_text == "":
                narration_text = "ZSSF for December 2025"
            if account_name is None or account_name == "":
                account_name = "THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS"
            
            # Format date for Tally (YYYYMMDD)
            if isinstance(voucher_date, str):
                try:
                    from datetime import datetime
                    if '-' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%Y-%m-%d")
                    elif '/' in voucher_date:
                        dt = datetime.strptime(voucher_date, "%d/%m/%Y")
                    else:
                        dt = datetime.strptime("30/12/2025", "%d/%m/%Y")
                    formatted_date = dt.strftime("%Y%m%d")
                except:
                    formatted_date = "20251230"
            else:
                formatted_date = "20251230"
            
            # Calculate totals for ZSSF
            total_employee_zssf = sum(emp.get('zssf_7', 0) + emp.get('zssf_14', 0) + emp.get('zssf_21', 0) for emp in employee_data)
            total_employer_zssf = total_employee_zssf  # Employer matches employee contribution
            total_amount = total_employee_zssf + total_employer_zssf
            
            # Start building XML for ZSSF Payment voucher
            xml_lines = [
                '<?xml version="1.0" encoding="UTF-8"?>',
                '<ENVELOPE>',
                ' <HEADER>',
                '  <TALLYREQUEST>Import Data</TALLYREQUEST>',
                ' </HEADER>',
                ' <BODY>',
                '  <IMPORTDATA>',
                '   <REQUESTDESC>',
                '    <REPORTNAME>Vouchers</REPORTNAME>',
                '    <STATICVARIABLES>',
                f'     <SVCURRENTCOMPANY>{comp_name}</SVCURRENTCOMPANY>',
                '    </STATICVARIABLES>',
                '   </REQUESTDESC>',
                '   <REQUESTDATA>',
                '    <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                '     <VOUCHER REMOTEID="" VCHKEY="" VCHTYPE="Payment" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                '      <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '      </OLDAUDITENTRYIDS.LIST>',
                f'      <DATE>{formatted_date}</DATE>',
                f'      <NARRATION>{narration_text}</NARRATION>',
                f'      <PARTYLEDGERNAME>{account_name}</PARTYLEDGERNAME>',
                '      <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>',
                '      <VOUCHERNUMBER>1</VOUCHERNUMBER>',
                '      <CSTFORMISSUETYPE/>',
                '      <CSTFORMRECVTYPE/>',
                '      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>',
                '      <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>',
                '      <VCHGSTCLASS/>',
                '      <ENTEREDBY>Administrator</ENTEREDBY>',
                '      <DIFFACTUALQTY>No</DIFFACTUALQTY>',
                '      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>',
                '      <ASORIGINAL>No</ASORIGINAL>',
                '      <AUDITED>No</AUDITED>',
                '      <FORJOBCOSTING>No</FORJOBCOSTING>',
                '      <ISOPTIONAL>No</ISOPTIONAL>',
                f'      <EFFECTIVEDATE>{formatted_date}</EFFECTIVEDATE>',
                '      <USEFOREXCISE>No</USEFOREXCISE>',
                '      <ISFORJOBWORKIN>No</ISFORJOBWORKIN>',
                '      <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>',
                '      <USEFORINTEREST>No</USEFORINTEREST>',
                '      <USEFORGAINLOSS>No</USEFORGAINLOSS>',
                '      <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>',
                '      <USEFORCOMPOUND>No</USEFORCOMPOUND>',
                '      <USEFORSERVICETAX>No</USEFORSERVICETAX>',
                '      <ISDELETED>No</ISDELETED>',
                '      <ISONHOLD>No</ISONHOLD>',
                '      <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>',
                '      <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>',
                '      <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>',
                '      <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>',
                '      <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>',
                '      <EXCISEOPENING>No</EXCISEOPENING>',
                '      <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>',
                '      <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>',
                '      <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>',
                '      <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>',
                '      <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>',
                '      <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>',
                '      <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>',
                '      <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>',
                '      <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>',
                '      <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>',
                '      <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>',
                '      <ISISDVOUCHER>No</ISISDVOUCHER>',
                '      <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>',
                '      <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>',
                '      <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>',
                '      <GSTNOTEXPORTED>No</GSTNOTEXPORTED>',
                '      <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>',
                '      <ISGSTREFUND>No</ISGSTREFUND>',
                '      <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>',
                '      <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>',
                '      <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>',
                '      <ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>',
                '      <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>',
                '      <ISCANCELLED>No</ISCANCELLED>',
                '      <HASCASHFLOW>Yes</HASCASHFLOW>',
                '      <ISPOSTDATED>No</ISPOSTDATED>',
                '      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>',
                '      <ISINVOICE>No</ISINVOICE>',
                '      <MFGJOURNAL>No</MFGJOURNAL>',
                '      <HASDISCOUNTS>No</HASDISCOUNTS>',
                '      <ASPAYSLIP>No</ASPAYSLIP>',
                '      <ISCOSTCENTRE>No</ISCOSTCENTRE>',
                '      <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>',
                '      <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>',
                '      <ISBLANKCHEQUE>No</ISBLANKCHEQUE>',
                '      <ISVOID>No</ISVOID>',
                '      <ORDERLINESTATUS>No</ORDERLINESTATUS>',
                '      <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>',
                '      <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>',
                '      <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>',
                '      <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>',
                '      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>',
                '      <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>',
                '      <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>',
                '      <CHANGEVCHMODE>No</CHANGEVCHMODE>',
            ]
            
            # Add ALLLEDGERENTRIES for Employee ZSSF
            xml_lines.extend([
                '      <ALLLEDGERENTRIES.LIST>',
                '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '       </OLDAUDITENTRYIDS.LIST>',
                '       <LEDGERNAME>ZSSF Employee</LEDGERNAME>',
                '       <GSTCLASS/>',
                '       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                '       <LEDGERFROMITEM>No</LEDGERFROMITEM>',
                '       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>',
                '       <ISPARTYLEDGER>No</ISPARTYLEDGER>',
                '       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>',
                '       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>',
                '       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>',
                f'       <AMOUNT>-{total_employee_zssf:.2f}</AMOUNT>',
                '       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>',
                '       <CATEGORYALLOCATIONS.LIST>',
                '        <CATEGORY>Primary Cost Category</CATEGORY>',
                '        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
            ])
            
            # Add cost center allocations for each employee
            for emp in employee_data:
                emp_name = emp.get('employee_name', 'Unknown')
                emp_total_zssf = emp.get('zssf_7', 0) + emp.get('zssf_14', 0) + emp.get('zssf_21', 0)
                if emp_total_zssf > 0:
                    xml_lines.extend([
                        '        <COSTCENTREALLOCATIONS.LIST>',
                        f'         <NAME>{emp_name}</NAME>',
                        f'         <AMOUNT>-{emp_total_zssf:.2f}</AMOUNT>',
                        '        </COSTCENTREALLOCATIONS.LIST>',
                    ])
            
            xml_lines.extend([
                '       </CATEGORYALLOCATIONS.LIST>',
                '       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>',
                '       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>',
                '       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>',
                '       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>',
                '       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>',
                '       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>',
                '       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>',
                '       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>',
                '       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>',
                '       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>',
                '       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>',
                '       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>',
                '       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>',
                '       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>',
                '       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>',
                '       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>',
                '       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>',
                '       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>',
                '       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>',
                '       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>',
                '       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>',
                '       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>',
                '      </ALLLEDGERENTRIES.LIST>',
            ])
            
            # Add ALLLEDGERENTRIES for Employer ZSSF
            xml_lines.extend([
                '      <ALLLEDGERENTRIES.LIST>',
                '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '       </OLDAUDITENTRYIDS.LIST>',
                '       <LEDGERNAME>ZSSF Employer</LEDGERNAME>',
                '       <GSTCLASS/>',
                '       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                '       <LEDGERFROMITEM>No</LEDGERFROMITEM>',
                '       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>',
                '       <ISPARTYLEDGER>No</ISPARTYLEDGER>',
                '       <ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>',
                '       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>',
                '       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>',
                f'       <AMOUNT>-{total_employer_zssf:.2f}</AMOUNT>',
                '       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>',
                '       <CATEGORYALLOCATIONS.LIST>',
                '        <CATEGORY>Primary Cost Category</CATEGORY>',
                '        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
            ])
            
            # Add cost center allocations for employer portion
            for emp in employee_data:
                emp_name = emp.get('employee_name', 'Unknown')
                emp_total_zssf = emp.get('zssf_7', 0) + emp.get('zssf_14', 0) + emp.get('zssf_21', 0)
                if emp_total_zssf > 0:
                    xml_lines.extend([
                        '        <COSTCENTREALLOCATIONS.LIST>',
                        f'         <NAME>{emp_name}</NAME>',
                        f'         <AMOUNT>-{emp_total_zssf:.2f}</AMOUNT>',
                        '        </COSTCENTREALLOCATIONS.LIST>',
                    ])
            
            xml_lines.extend([
                '       </CATEGORYALLOCATIONS.LIST>',
                '       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>',
                '       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>',
                '       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>',
                '       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>',
                '       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>',
                '       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>',
                '       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>',
                '       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>',
                '       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>',
                '       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>',
                '       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>',
                '       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>',
                '       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>',
                '       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>',
                '       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>',
                '       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>',
                '       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>',
                '       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>',
                '       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>',
                '       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>',
                '       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>',
                '       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>',
                '      </ALLLEDGERENTRIES.LIST>',
            ])
            
            # Add Bank/Payment ALLLEDGERENTRIES
            xml_lines.extend([
                '      <ALLLEDGERENTRIES.LIST>',
                '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '       </OLDAUDITENTRYIDS.LIST>',
                f'       <LEDGERNAME>{account_name}</LEDGERNAME>',
                '       <GSTCLASS/>',
                '       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                '       <LEDGERFROMITEM>No</LEDGERFROMITEM>',
                '       <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>',
                '       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>',
                '       <ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>',
                '       <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>',
                '       <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>',
                f'       <AMOUNT>{total_amount:.2f}</AMOUNT>',
                '       <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>',
                '       <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>',
                '       <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>',
                '       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>',
                '       <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>',
                '       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>',
                '       <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>',
                '       <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>',
                '       <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>',
                '       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>',
                '       <RATEDETAILS.LIST>       </RATEDETAILS.LIST>',
                '       <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>',
                '       <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>',
                '       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>',
                '       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>',
                '       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>',
                '       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>',
                '       <VATSTATUTORYDETAILS.LIST>       </VATSTATUTORYDETAILS.LIST>',
                '       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>',
                '       <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>',
                '       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>',
                '       <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>',
                '       <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>',
                '      </ALLLEDGERENTRIES.LIST>',
            ])
            
            # Close the XML structure
            xml_lines.extend([
                '      <PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>',
                '      <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>',
                '      <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>',
                '      <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>',
                '      <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>',
                '     </VOUCHER>',
                '    </TALLYMESSAGE>',
                '   </REQUESTDATA>',
                '  </IMPORTDATA>',
                ' </BODY>',
                '</ENVELOPE>'
            ])
            
            return '\n'.join(xml_lines)
            
        except Exception as e:
            print(f"❌ Error generating ZSSF XML: {str(e)}")
            return ""

    def process_zssf_sheet(self, file_path: str) -> Dict[str, Any]:
        """Process ZSSF sheet (Sheet 3) and extract data."""
        # Read the third sheet (index 2)
        df = self.read_excel_file(file_path, sheet_name=2)
        
        if df is None:
            return {"error": "Could not read the Excel file", "success": False}
        
        try:
            print(f"📊 Processing ZSSF sheet: {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Extract header information from first few rows
            header_info = self.extract_header_info(df)
            
            # Find where employee data starts - look for row with EMPLOYEE NAME
            employee_start_row = self.find_employee_data_start(df)
            
            # Extract ZSSF employee data
            employee_data = self.extract_zssf_employee_data(df, employee_start_row)
            
            # Calculate totals
            total_zssf = sum(emp.get('zssf_7', 0) + emp.get('zssf_14', 0) + emp.get('zssf_21', 0) for emp in employee_data)
            
            # Prepare result
            result = {
                "file_name": os.path.basename(file_path),
                "sheet_type": "ZSSF",
                "success": True,
                "date": header_info["date"],
                "company_name": header_info["company_name"],
                "narration": header_info.get("narration", "ZSSF for current period"),
                "employee_data": employee_data,
                "total_employees": len(employee_data),
                "total_zssf": total_zssf,
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "employee_data_start_row": employee_start_row + 1
            }
            
            print(f"✅ ZSSF processing completed: {len(employee_data)} employees, Total ZSSF: ₹{total_zssf:,.2f}")
            return result
            
        except Exception as e:
            print(f"❌ Error in ZSSF processing: {str(e)}")
            return {"error": f"Error processing ZSSF sheet: {str(e)}", "success": False}

    def process_zhsf_sheet(self, file_path: str) -> Dict[str, Any]:
        """Process ZHSF sheet (Sheet 4) and extract data."""
        # Read the fourth sheet (index 3)
        df = self.read_excel_file(file_path, sheet_name=3)
        
        if df is None:
            return {"error": "Could not read the Excel file", "success": False}
        
        try:
            print(f"📊 Processing ZHSF sheet: {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Extract header information from first few rows
            header_info = self.extract_header_info(df)
            
            # Find where employee data starts - look for row with EMPLOYEE NAME
            employee_start_row = self.find_employee_data_start(df)
            
            # Extract ZHSF employee data
            employee_data = self.extract_zhsf_employee_data(df, employee_start_row)
            
            # Calculate totals
            total_zhsf = sum(emp.get('employee_35', 0) + emp.get('twa_35', 0) for emp in employee_data)
            
            # Prepare result
            result = {
                "file_name": os.path.basename(file_path),
                "sheet_type": "ZHSF",
                "success": True,
                "date": header_info["date"],
                "company_name": header_info["company_name"],
                "narration": header_info.get("narration", "ZHSF for current period"),
                "employee_data": employee_data,
                "total_employees": len(employee_data),
                "total_zhsf": total_zhsf,
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "employee_data_start_row": employee_start_row + 1
            }
            
            print(f"✅ ZHSF processing completed: {len(employee_data)} employees, Total ZHSF: ₹{total_zhsf:,.2f}")
            return result
            
        except Exception as e:
            print(f"❌ Error in ZHSF processing: {str(e)}")
            return {"error": f"Error processing ZHSF sheet: {str(e)}", "success": False}

    def extract_zssf_employee_data(self, df: pd.DataFrame, employee_start_row: int) -> List[Dict[str, Any]]:
        """Extract ZSSF employee data from the DataFrame."""
        employee_data = []
        
        try:
            # Get header row
            header_row = df.iloc[employee_start_row]
            
            # Create column mapping for ZSSF sheet
            column_mapping = {}
            for idx, header in enumerate(header_row):
                if not pd.isna(header):
                    header_str = str(header).strip().upper()
                    if 'EMPLOYEE' in header_str and 'NAME' in header_str:
                        column_mapping['employee_name'] = idx
                    elif 'ZSSF' in header_str and '7%' in header_str:
                        column_mapping['zssf_7'] = idx
                    elif 'ZSSF' in header_str and '14%' in header_str:
                        column_mapping['zssf_14'] = idx
                    elif '@' in header_str and '21%' in header_str:
                        column_mapping['zssf_21'] = idx
                    elif header_str == 'REMARKS':
                        column_mapping['remarks'] = idx
                    elif 'ZSSF' in header_str and 'NO' in header_str:
                        column_mapping['zssf_no'] = idx
                    elif header_str == 'SALARY':
                        column_mapping['salary'] = idx
                    elif header_str == 'NAME':
                        column_mapping['name'] = idx
            
            print(f"🎯 ZSSF Column mapping: {column_mapping}")
            
            # Extract data rows
            data_rows = df.iloc[employee_start_row + 1:]
            
            for idx, row in data_rows.iterrows():
                # Skip empty rows
                if row.isna().all():
                    continue
                
                employee_record = {}
                
                # Extract employee data using column mapping
                for field, col_idx in column_mapping.items():
                    if col_idx < len(row):
                        value = row.iloc[col_idx]
                        if not pd.isna(value):
                            if field in ['zssf_7', 'zssf_14', 'zssf_21', 'salary']:
                                # Convert to numeric
                                try:
                                    employee_record[field] = float(value)
                                except (ValueError, TypeError):
                                    employee_record[field] = 0
                            else:
                                # Ensure string fields are properly converted - safe handling
                                str_value = str(value).strip() if value is not None and str(value).strip() != 'nan' else ""
                                employee_record[field] = str_value
                        else:
                            employee_record[field] = 0 if field in ['zssf_7', 'zssf_14', 'zssf_21', 'salary'] else ""
                
                # Ensure employee_name is always a valid string
                if 'employee_name' in employee_record and not employee_record['employee_name']:
                    employee_record['employee_name'] = f"Employee_{idx}"
                
                # Only add if we have essential data
                total_zssf = employee_record.get('zssf_7', 0) + employee_record.get('zssf_14', 0) + employee_record.get('zssf_21', 0)
                if employee_record.get('employee_name') and total_zssf > 0:
                    employee_data.append(employee_record)
                    emp_name = employee_record.get('employee_name', 'Unknown')
                    print(f"Added ZSSF employee: {emp_name}, ZSSF 7%: {employee_record.get('zssf_7', 0)}, ZSSF 14%: {employee_record.get('zssf_14', 0)}, ZSSF 21%: {employee_record.get('zssf_21', 0)}")
            
        except Exception as e:
            print(f"❌ Error extracting ZSSF employee data: {str(e)}")
        
        return employee_data

    def extract_zhsf_employee_data(self, df: pd.DataFrame, employee_start_row: int) -> List[Dict[str, Any]]:
        """Extract ZHSF employee data from the DataFrame."""
        employee_data = []
        
        try:
            # Get header row
            header_row = df.iloc[employee_start_row]
            
            # Create column mapping for ZHSF sheet
            column_mapping = {}
            for idx, header in enumerate(header_row):
                if not pd.isna(header):
                    header_str = str(header).strip().upper()
                    if 'EMPLOYEE' in header_str and 'NAME' in header_str:
                        column_mapping['employee_name'] = idx
                    elif 'EMPLOYEE' in header_str and '3.5%' in header_str:
                        column_mapping['employee_35'] = idx
                    elif 'TWA' in header_str and '3.5%' in header_str:
                        column_mapping['twa_35'] = idx
                    elif header_str == 'SALARY':
                        column_mapping['salary'] = idx
                    elif header_str == 'TOTAL':
                        column_mapping['total'] = idx
                    elif 'EMPL' in header_str and 'NO' in header_str:
                        column_mapping['empl_no'] = idx
            
            print(f"🎯 ZHSF Column mapping: {column_mapping}")
            
            # Extract data rows
            data_rows = df.iloc[employee_start_row + 1:]
            
            for idx, row in data_rows.iterrows():
                # Skip empty rows
                if row.isna().all():
                    continue
                
                employee_record = {}
                
                # Extract employee data using column mapping
                for field, col_idx in column_mapping.items():
                    if col_idx < len(row):
                        value = row.iloc[col_idx]
                        if not pd.isna(value):
                            if field in ['employee_35', 'twa_35', 'salary', 'total']:
                                # Convert to numeric
                                try:
                                    employee_record[field] = float(value)
                                except (ValueError, TypeError):
                                    employee_record[field] = 0
                            else:
                                # Ensure string fields are properly converted - safe handling
                                str_value = str(value).strip() if value is not None and str(value).strip() != 'nan' else ""
                                employee_record[field] = str_value
                        else:
                            employee_record[field] = 0 if field in ['employee_35', 'twa_35', 'salary', 'total'] else ""
                
                # Ensure employee_name is always a valid string
                if 'employee_name' in employee_record and not employee_record['employee_name']:
                    employee_record['employee_name'] = f"Employee_{idx}"
                
                # Only add if we have essential data
                total_zhsf = employee_record.get('employee_35', 0) + employee_record.get('twa_35', 0)
                if employee_record.get('employee_name') and total_zhsf > 0:
                    employee_data.append(employee_record)
                    emp_name = employee_record.get('employee_name', 'Unknown')
                    print(f"Added ZHSF employee: {emp_name}, Employee 3.5%: {employee_record.get('employee_35', 0)}, TWA 3.5%: {employee_record.get('twa_35', 0)}")
            
        except Exception as e:
            print(f"❌ Error extracting ZHSF employee data: {str(e)}")
        
        return employee_data