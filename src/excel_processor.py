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
            
            # Create column mapping
            column_mapping = {}
            for idx, header in enumerate(header_row):
                if not pd.isna(header):
                    header_clean = str(header).strip().upper()
                    if 'EMPL' in header_clean and 'NO' in header_clean:
                        column_mapping['employee_no'] = idx
                    elif 'EMPLOYEE' in header_clean and 'NAME' in header_clean:
                        column_mapping['employee_name'] = idx
                    elif 'ATTENDANCE' in header_clean and 'PRODUCTION' in header_clean:
                        column_mapping['attendance_type'] = idx
                    elif 'ATTENDANCE' in header_clean and 'DAYS' in header_clean:
                        column_mapping['attendance_days'] = idx
            
            # Extract data rows
            data_rows = df.iloc[employee_start_row + 1:]
            
            for idx, row in data_rows.iterrows():
                # Skip completely empty rows
                if row.isna().all():
                    continue
                
                employee_record = {}
                
                # Extract data based on column mapping
                if 'employee_no' in column_mapping:
                    emp_no = row.iloc[column_mapping['employee_no']]
                    employee_record['employee_no'] = str(emp_no) if not pd.isna(emp_no) else ''
                
                if 'employee_name' in column_mapping:
                    emp_name = row.iloc[column_mapping['employee_name']]
                    employee_record['employee_name'] = str(emp_name) if not pd.isna(emp_name) else ''
                
                if 'attendance_type' in column_mapping:
                    att_type = row.iloc[column_mapping['attendance_type']]
                    employee_record['attendance_type'] = str(att_type) if not pd.isna(att_type) else ''
                
                if 'attendance_days' in column_mapping:
                    att_days = row.iloc[column_mapping['attendance_days']]
                    employee_record['attendance_days'] = str(att_days) if not pd.isna(att_days) else ''
                
                # Only add record if it has at least employee name or number
                if employee_record.get('employee_name') or employee_record.get('employee_no'):
                    employee_data.append(employee_record)
            
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
                            # Try to convert to number if possible, otherwise keep as string
                            try:
                                numeric_value = float(str(cell_value).replace(',', ''))
                                all_data[col_name] = numeric_value
                                
                                # Add to total if it's a numeric value and not the first two columns 
                                # (assuming first two are usually ID and Name type columns)
                                if col_idx >= 2:  # Skip first two columns from total calculation
                                    total_salary += numeric_value
                            except (ValueError, TypeError):
                                # Keep as string if not numeric
                                all_data[col_name] = str(cell_value).strip()
                        else:
                            all_data[col_name] = 0 if col_idx >= 2 else ''  # Numeric 0 for data columns, empty string for text columns
                    except Exception as e:
                        print(f"⚠️ Warning processing column {col_name}: {str(e)}")
                        all_data[col_name] = 0 if col_idx >= 2 else ''
                
                # Set dynamic fields
                if all_columns:
                    # Use first column as employee_no and second as employee_name (if they exist)
                    if len(all_columns) >= 1:
                        first_col_name = all_columns[0][1]
                        employee_record['employee_no'] = str(all_data.get(first_col_name, ''))
                    
                    if len(all_columns) >= 2:
                        second_col_name = all_columns[1][1]
                        employee_record['employee_name'] = str(all_data.get(second_col_name, ''))
                    
                    # All columns after the first two become salary components
                    salary_components = {}
                    if len(all_columns) > 2:
                        for col_idx, col_name in all_columns[2:]:  # Skip first two columns
                            salary_components[col_name] = all_data.get(col_name, 0)
                    
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