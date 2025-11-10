import requests
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional
import logging


class TallyApiService:
    """Service class to handle Tally API communication."""
    
    def __init__(self, base_url: str = "http://localhost:9000/"):
        """Initialize the Tally API service with base URL."""
        self.base_url = base_url.rstrip('/')
        self.session = requests.Session()
        self.session.headers.update({
            'Content-Type': 'application/xml',
            'Accept': 'application/xml'
        })
        
        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def check_tally_version(self) -> Dict[str, any]:
        """
        Check if Tally Prime or Tally ERP is running.
        
        Returns:
            Dict containing version info: {"is_prime": bool, "version": str, "success": bool}
        """
        try:
            self.logger.info("Checking Tally version...")
            
            # Make simple GET request to check response
            response = requests.get(self.base_url, timeout=10)
            
            if response.status_code == 200:
                response_text = response.text.strip()
                
                # Check if response contains "TallyPrime"
                if "TallyPrime" in response_text or "Tally Prime" in response_text:
                    self.logger.info("✅ Detected: Tally Prime")
                    return {
                        "success": True,
                        "is_prime": True,
                        "version": "Tally Prime",
                        "message": "TallyPrime Server is Running"
                    }
                else:
                    self.logger.info("✅ Detected: Tally ERP")
                    return {
                        "success": True,
                        "is_prime": False,
                        "version": "Tally ERP 9",
                        "message": "Tally Server is Running"
                    }
            else:
                return {
                    "success": False,
                    "is_prime": False,
                    "version": "Unknown",
                    "error": f"HTTP {response.status_code}"
                }
                
        except requests.exceptions.ConnectionError:
            return {
                "success": False,
                "is_prime": False,
                "version": "Unknown",
                "error": "Cannot connect to Tally. Please ensure Tally is running."
            }
        except Exception as e:
            return {
                "success": False,
                "is_prime": False,
                "version": "Unknown",
                "error": str(e)
            }
    
    def get_server_info(self) -> Dict[str, any]:
        """
        Get basic information about the Tally instance.
        
        Returns:
            Dict containing Tally server information
        """
        # This could be expanded to get Tally version, license info, etc.
        # For now, just return connection status
        return self.test_connection()
    
    def upload_payroll_data(self, payroll_result: Dict[str, any], company_name: str, account_name: str = "Cash") -> Dict[str, any]:
        """
        Upload payroll data to Tally.
        
        Args:
            payroll_result: Processed payroll data from Excel processor
            company_name: Selected company name
            account_name: Account name to use (PARTYLEDGERNAME), defaults to "Cash"
            
        Returns:
            Dict containing success status and upload results
        """
        try:
            self.logger.info("Starting payroll data upload to Tally...")
            
            # Generate XML for payroll voucher
            xml_request = self._generate_payroll_xml(payroll_result, company_name, account_name)
            
            self.logger.info("Generated payroll XML for Tally upload")
            self.logger.debug(f"Payroll XML content: {xml_request[:1000]}...")
            
            # Make POST request to Tally
            response = self.session.post(
                self.base_url,
                data=xml_request,
                timeout=30
            )
            
            # Check if request was successful
            if response.status_code == 200:
                self.logger.info("✅ Successfully sent payroll request to Tally")
                
                # Parse Tally response to check for errors
                tally_result = self._parse_tally_response(response.text)
                return tally_result
            else:
                error_msg = f"HTTP {response.status_code}: {response.reason}"
                self.logger.error(f"❌ Payroll upload failed: {error_msg}")
                return {
                    "success": False,
                    "error": error_msg,
                    "response": response.text if response.text else None,
                    "errors_count": 0,
                    "created_count": 0
                }
                
        except Exception as e:
            error_msg = f"Payroll upload error: {str(e)}"
            self.logger.error(f"❌ {error_msg}")
            return {
                "success": False,
                "error": error_msg
            }
    
    def _generate_payroll_xml(self, payroll_result: Dict[str, any], company_name: str, account_name: str = "Cash") -> str:
        """
        Generate XML envelope for payroll data upload.
        
        Args:
            payroll_result: Processed payroll data
            company_name: Selected company name
            account_name: Account name to use as PARTYLEDGERNAME
            
        Returns:
            XML string formatted for Tally import
        """
        try:
            # Extract data from payroll result
            date = payroll_result.get('date', '20250401')  # Default date if not found
            narration = payroll_result.get('narration', 'Payroll Import')
            employee_data = payroll_result.get('employee_data', [])
            
            # Convert date format if needed (from DD-MM-YYYY to YYYYMMDD)
            formatted_date = self._format_date_for_tally(date)
            
            # Start building XML
            xml_lines = [
                '<ENVELOPE>',
                ' <HEADER>',
                '  <TALLYREQUEST>Import Data</TALLYREQUEST>',
                ' </HEADER>',
                ' <BODY>',
                '  <IMPORTDATA>',
                '   <REQUESTDESC>',
                '    <REPORTNAME>Vouchers</REPORTNAME>',
                '    <STATICVARIABLES>',
                f'     <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>',
                '    </STATICVARIABLES>',
                '   </REQUESTDESC>',
                '   <REQUESTDATA>',
                '    <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                '     <VOUCHER VCHTYPE="Payroll" ACTION="Create" OBJVIEW="PaySlip Voucher View">',
                f'      <DATE>{formatted_date}</DATE>',
                f'      <NARRATION>{narration}</NARRATION>',
                f'      <PARTYLEDGERNAME>{account_name}</PARTYLEDGERNAME>',
                '      <VOUCHERTYPENAME>Payroll</VOUCHERTYPENAME>',
                '      <CATEGORYENTRY.LIST>',
                '       <CATEGORY>Primary Cost Category</CATEGORY>'
            ]
            
            # Add employee entries
            employee_sort_order = 1
            for employee in employee_data:
                emp_name = employee.get('employee_name', '').strip()
                salary_components = employee.get('salary_components', {})
                total_gross = employee.get('total_gross_salary', 0)
                
                if emp_name and salary_components:  # Only add if we have name and salary data
                    xml_lines.extend([
                        '       <EMPLOYEEENTRIES.LIST>',
                        f'        <EMPLOYEENAME>{emp_name}</EMPLOYEENAME>',
                        f'        <EMPLOYEESORTORDER> {employee_sort_order}</EMPLOYEESORTORDER>',
                        f'        <AMOUNT>-{total_gross:.2f}</AMOUNT>'
                    ])
                    
                    # Add payhead allocations for each salary component
                    payhead_sort_order = 1
                    for payhead_name, amount in salary_components.items():
                        if amount != 0:  # Only include non-zero amounts
                            xml_lines.extend([
                                '        <PAYHEADALLOCATIONS.LIST>',
                                f'         <PAYHEADNAME>{payhead_name}</PAYHEADNAME>',
                                '         <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                                f'         <PAYHEADSORTORDER> {payhead_sort_order}</PAYHEADSORTORDER>',
                                f'         <AMOUNT>-{amount:.2f}</AMOUNT>',
                                '        </PAYHEADALLOCATIONS.LIST>'
                            ])
                            payhead_sort_order += 1
                    
                    xml_lines.append('       </EMPLOYEEENTRIES.LIST>')
                    employee_sort_order += 1
            
            # Close XML structure
            xml_lines.extend([
                '      </CATEGORYENTRY.LIST>',
                '     </VOUCHER>',
                '    </TALLYMESSAGE>',
                '   </REQUESTDATA>',
                '  </IMPORTDATA>',
                ' </BODY>',
                '</ENVELOPE>'
            ])
            
            xml_content = '\n'.join(xml_lines)
            self.logger.debug(f"Generated payroll XML with {len(employee_data)} employees")
            return xml_content
            
        except Exception as e:
            self.logger.error(f"Error generating payroll XML: {str(e)}")
            raiseession.headers.update({
            'Content-Type': 'application/xml',
            'Accept': 'application/xml'
        })
        
        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def get_companies(self) -> Dict[str, any]:
        """
        Fetch list of companies from Tally.
        
        Returns:
            Dict containing success status, companies list, or error message
        """
        try:
            # XML envelope for getting company list
            xml_request = """<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Companies</ID>
  </HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>List of Companies</REPORTNAME>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>"""
            
            self.logger.info("Sending request to Tally for company list...")
            
            # Make POST request to Tally
            response = self.session.post(
                self.base_url,
                data=xml_request,
                timeout=30
            )
            
            # Check if request was successful
            if response.status_code == 200:
                self.logger.info("✅ Successfully received response from Tally")
                
                # Parse the XML response
                companies = self._parse_companies_xml(response.text)
                
                if companies:
                    return {
                        "success": True,
                        "companies": companies,
                        "message": f"Found {len(companies)} companies"
                    }
                else:
                    return {
                        "success": False,
                        "companies": [],
                        "error": "No companies found in response"
                    }
            else:
                error_msg = f"HTTP {response.status_code}: {response.reason}"
                self.logger.error(f"❌ Request failed: {error_msg}")
                return {
                    "success": False,
                    "companies": [],
                    "error": error_msg
                }
                
        except requests.exceptions.ConnectionError:
            error_msg = "Cannot connect to Tally. Please ensure Tally is running and accessible."
            self.logger.error(f"❌ Connection error: {error_msg}")
            return {
                "success": False,
                "companies": [],
                "error": error_msg
            }
        except requests.exceptions.Timeout:
            error_msg = "Request timed out. Tally may be busy or unresponsive."
            self.logger.error(f"❌ Timeout error: {error_msg}")
            return {
                "success": False,
                "companies": [],
                "error": error_msg
            }
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"
            self.logger.error(f"❌ Unexpected error: {error_msg}")
            return {
                "success": False,
                "companies": [],
                "error": error_msg
            }
    
    def _parse_companies_xml(self, xml_response: str) -> List[str]:
        """
        Parse XML response to extract company names.
        
        Args:
            xml_response: XML response string from Tally
            
        Returns:
            List of company names
        """
        try:
            self.logger.info("Parsing XML response for company names...")
            
            # Parse XML
            root = ET.fromstring(xml_response)
            
            # Find COMPANYNAME.LIST element
            company_list = root.find('.//COMPANYNAME.LIST')
            
            if company_list is None:
                self.logger.warning("⚠️ COMPANYNAME.LIST not found in XML response")
                return []
            
            # Extract all COMPANYNAME elements
            companies = []
            for company_elem in company_list.findall('COMPANYNAME'):
                company_name = company_elem.text
                if company_name and company_name.strip():
                    companies.append(company_name.strip())
            
            self.logger.info(f"✅ Parsed {len(companies)} companies from XML")
            return companies
            
        except ET.ParseError as e:
            self.logger.error(f"❌ XML parsing error: {str(e)}")
            self.logger.debug(f"XML content: {xml_response[:500]}...")
            return []
        except Exception as e:
            self.logger.error(f"❌ Error parsing companies XML: {str(e)}")
            return []
    
    def test_connection(self) -> Dict[str, any]:
        """
        Test connection to Tally server.
        
        Returns:
            Dict containing connection test results
        """
        try:
            self.logger.info("Testing connection to Tally...")
            
            # Simple ping request to test connectivity
            response = self.session.get(self.base_url, timeout=10)
            
            return {
                "success": True,
                "status_code": response.status_code,
                "message": "Connection successful"
            }
            
        except requests.exceptions.ConnectionError:
            return {
                "success": False,
                "error": "Cannot connect to Tally server"
            }
        except Exception as e:
            return {
                "success": False,
                "error": f"Connection test failed: {str(e)}"
            }
    
    def upload_xml_to_tally(self, xml_content: str, voucher_type: str = "Generic") -> Dict[str, any]:
        """
        Upload pre-generated XML to Tally.
        
        Args:
            xml_content: Pre-generated XML string to send to Tally
            voucher_type: Type of voucher (for logging)
            
        Returns:
            Dict containing upload status and response
        """
        try:
            # Log the XML being sent
            print("="*80)
            print(f"📤 SENDING {voucher_type.upper()} XML TO TALLY:")
            print("="*80)
            print(xml_content)
            print("="*80)
            
            self.logger.info(f"Uploading {voucher_type} data to Tally...")
            
            # Make POST request to Tally
            response = self.session.post(
                self.base_url,
                data=xml_content.encode('utf-8'),
                timeout=30
            )
            
            self.logger.info(f"📥 Received response from Tally (Status: {response.status_code})")
            
            # Check if request was successful
            if response.status_code == 200:
                self.logger.info("✅ Successfully sent request to Tally")
                
                # Parse Tally response to check for errors
                tally_result = self._parse_tally_response(response.text)
                return tally_result
            else:
                error_msg = f"HTTP {response.status_code}: {response.reason}"
                self.logger.error(f"❌ Upload failed: {error_msg}")
                return {
                    "success": False,
                    "error": error_msg,
                    "response": response.text if response.text else None,
                    "errors_count": 0,
                    "created_count": 0
                }
                
        except Exception as e:
            error_msg = f"Upload error: {str(e)}"
            self.logger.error(f"❌ {error_msg}")
            import traceback
            traceback.print_exc()
            return {
                "success": False,
                "error": error_msg
            }
    
    def upload_attendance_data(self, attendance_result: Dict[str, any], company_name: str) -> Dict[str, any]:
        """
        Upload attendance data to Tally.
        
        Args:
            attendance_result: Processed attendance data from Excel
            company_name: Selected company name
            
        Returns:
            Dict containing upload status and response
        """
        # This method is deprecated - XML should be generated by excel_processor
        # and then uploaded using upload_xml_to_tally
        self.logger.warning("⚠️ upload_attendance_data is deprecated, use upload_xml_to_tally instead")
        
        try:
            # Generate XML from attendance data
            xml_request = self._generate_attendance_xml(attendance_result, company_name)
            
            # Use the new upload method
            return self.upload_xml_to_tally(xml_request, "Attendance")
                
        except Exception as e:
            error_msg = f"Upload error: {str(e)}"
            self.logger.error(f"❌ {error_msg}")
            return {
                "success": False,
                "error": error_msg
            }
    
    def _generate_attendance_xml(self, attendance_result: Dict[str, any], company_name: str) -> str:
        """
        Generate XML envelope for attendance data upload.
        
        Args:
            attendance_result: Processed attendance data
            company_name: Selected company name
            
        Returns:
            XML string formatted for Tally import
        """
        try:
            # Extract data from attendance result
            date = attendance_result.get('date', '20250401')  # Default date if not found
            narration = attendance_result.get('narration', 'Attendance Import')
            employee_data = attendance_result.get('employee_data', [])
            
            # Convert date format if needed (from DD-MM-YYYY to YYYYMMDD)
            formatted_date = self._format_date_for_tally(date)
            
            # Generate voucher number (simple increment, you might want to make this more sophisticated)
            voucher_number = "1"
            
            # Start building XML
            xml_lines = [
                '<ENVELOPE>',
                ' <HEADER>',
                '  <TALLYREQUEST>Import Data</TALLYREQUEST>',
                ' </HEADER>',
                ' <BODY>',
                '  <IMPORTDATA>',
                '   <REQUESTDESC>',
                '    <REPORTNAME>Vouchers</REPORTNAME>',
                '    <STATICVARIABLES>',
                f'     <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>',
                '    </STATICVARIABLES>',
                '   </REQUESTDESC>',
                '   <REQUESTDATA>',
                '    <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                '     <VOUCHER VCHTYPE="Attendance" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                f'      <DATE>{formatted_date}</DATE>',
                '      <VOUCHERTYPENAME>Attendance</VOUCHERTYPENAME>',
                f'      <NARRATION>{narration}</NARRATION>',
                '      <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>',
                '      <VCHGSTCLASS/>'
            ]
            
            # Add attendance entries for each employee
            for employee in employee_data:
                emp_name = employee.get('employee_name', '').strip()
                att_type = employee.get('attendance_type', '').strip()
                att_days = employee.get('attendance_days', '0').strip()
                
                if emp_name and att_type:  # Only add if we have name and type
                    xml_lines.extend([
                        '      <ATTENDANCEENTRIES.LIST>',
                        f'       <NAME>{emp_name}</NAME>',
                        f'       <ATTENDANCETYPE>{att_type}</ATTENDANCETYPE>',
                        f'       <ATTDTYPETIMEVALUE>{att_days}</ATTDTYPETIMEVALUE>',
                        '      </ATTENDANCEENTRIES.LIST>'
                    ])
            
            # Close XML structure
            xml_lines.extend([
                '     </VOUCHER>',
                '    </TALLYMESSAGE>',
                '   </REQUESTDATA>',
                '  </IMPORTDATA>',
                ' </BODY>',
                '</ENVELOPE>'
            ])
            
            return '\n'.join(xml_lines)
            
        except Exception as e:
            self.logger.error(f"❌ Error generating attendance XML: {str(e)}")
            return ""
    
    def _parse_tally_response(self, xml_response: str) -> Dict[str, any]:
        """
        Parse Tally XML response to check for errors and success.
        
        Args:
            xml_response: XML response from Tally
            
        Returns:
            Dict with success status and parsed information
        """
        try:
            self.logger.info("Parsing Tally response...")
            
            # Parse XML
            root = ET.fromstring(xml_response)
            
            # Extract key fields
            errors_elem = root.find('ERRORS')
            created_elem = root.find('CREATED')
            line_error_elem = root.find('LINEERROR')
            
            # Get counts
            errors_count = int(errors_elem.text) if errors_elem is not None and errors_elem.text else 0
            created_count = int(created_elem.text) if created_elem is not None and created_elem.text else 0
            
            # Get error message if any
            error_message = line_error_elem.text if line_error_elem is not None and line_error_elem.text else ""
            
            self.logger.info(f"Tally Response - Errors: {errors_count}, Created: {created_count}")
            
            if errors_count == 0 and created_count > 0:
                # Success - no errors and something was created
                return {
                    "success": True,
                    "message": f"Successfully uploaded {created_count} attendance record(s)",
                    "response": xml_response,
                    "errors_count": errors_count,
                    "created_count": created_count,
                    "error_message": ""
                }
            else:
                # Error - either has errors or nothing was created
                error_msg = "Upload failed"
                if error_message:
                    error_msg = f"Tally Error: {error_message}"
                elif errors_count > 0:
                    error_msg = f"Tally reported {errors_count} error(s)"
                
                return {
                    "success": False,
                    "error": error_msg,
                    "response": xml_response,
                    "errors_count": errors_count,
                    "created_count": created_count,
                    "error_message": error_message
                }
                
        except ET.ParseError as e:
            self.logger.error(f"❌ XML parsing error: {str(e)}")
            return {
                "success": False,
                "error": f"Invalid XML response from Tally: {str(e)}",
                "response": xml_response,
                "errors_count": 0,
                "created_count": 0,
                "error_message": ""
            }
        except Exception as e:
            self.logger.error(f"❌ Error parsing Tally response: {str(e)}")
            return {
                "success": False,
                "error": f"Error parsing Tally response: {str(e)}",
                "response": xml_response,
                "errors_count": 0,
                "created_count": 0,
                "error_message": ""
            }
    
    def _format_date_for_tally(self, date_str: str) -> str:
        """
        Format date string for Tally (YYYYMMDD format).
        
        Args:
            date_str: Date string in various formats
            
        Returns:
            Formatted date string for Tally
        """
        try:
            if not date_str:
                return "20250401"  # Default date
            
            # Handle different date formats
            date_str = str(date_str).strip()
            
            # If it's already in YYYYMMDD format
            if len(date_str) == 8 and date_str.isdigit():
                return date_str
            
            # Handle DD-MM-YYYY format
            if '-' in date_str and len(date_str.split('-')) == 3:
                parts = date_str.split('-')
                if len(parts[2]) == 4:  # DD-MM-YYYY
                    day, month, year = parts
                    return f"{year}{month.zfill(2)}{day.zfill(2)}"
                elif len(parts[0]) == 4:  # YYYY-MM-DD
                    year, month, day = parts
                    return f"{year}{month.zfill(2)}{day.zfill(2)}"
            
            # Default fallback
            return "20250401"
            
        except Exception as e:
            self.logger.error(f"❌ Error formatting date: {str(e)}")
            return "20250401"
    
    def upload_paye_data(self, paye_result: Dict[str, any], company_name: str, account_name: str = "THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS") -> Dict[str, any]:
        """
        Upload PAYE data to Tally.
        
        Args:
            paye_result: Processed PAYE data from Excel processor
            company_name: Selected company name
            account_name: Bank account name to use for payment, defaults to Zanzibar bank
            
        Returns:
            Dict containing success status and upload results
        """
        try:
            self.logger.info("Starting PAYE data upload to Tally...")
            
            # Generate XML for PAYE payment voucher
            xml_request = self._generate_paye_xml(paye_result, company_name, account_name)
            
            # Send XML to Tally
            self.logger.info("Sending PAYE XML to Tally...")
            response = self.session.post(self.base_url, data=xml_request)
            
            if response.status_code == 200:
                # Parse response
                try:
                    root = ET.fromstring(response.text)
                    
                    # Check for errors in response
                    error_element = root.find('.//ERROR')
                    if error_element is not None:
                        error_msg = error_element.text or "Unknown Tally error"
                        self.logger.error(f"❌ Tally error: {error_msg}")
                        return {
                            "success": False,
                            "error": f"Tally error: {error_msg}",
                            "response": response.text
                        }
                    
                    # Check for success indicators
                    created_element = root.find('.//CREATED')
                    if created_element is not None and created_element.text == "1":
                        self.logger.info("✅ PAYE data uploaded successfully to Tally")
                        return {
                            "success": True,
                            "message": "PAYE data uploaded successfully",
                            "employees_count": paye_result.get("total_employees", 0),
                            "total_amount": paye_result.get("total_amount", 0),
                            "response": response.text
                        }
                    else:
                        self.logger.warning("⚠️ Tally response unclear, check manually")
                        return {
                            "success": False,
                            "error": "Unclear response from Tally - please check manually",
                            "response": response.text
                        }
                        
                except ET.ParseError as e:
                    self.logger.error(f"❌ Error parsing Tally response: {str(e)}")
                    return {
                        "success": False,
                        "error": f"Error parsing Tally response: {str(e)}",
                        "response": response.text
                    }
            else:
                self.logger.error(f"❌ HTTP error {response.status_code}: {response.text}")
                return {
                    "success": False,
                    "error": f"HTTP error {response.status_code}",
                    "response": response.text
                }
                
        except requests.exceptions.ConnectionError:
            self.logger.error("❌ Could not connect to Tally. Is Tally running?")
            return {
                "success": False,
                "error": "Could not connect to Tally. Please ensure Tally is running and configured properly."
            }
        except requests.exceptions.RequestException as e:
            self.logger.error(f"❌ Request error: {str(e)}")
            return {
                "success": False,
                "error": f"Network error: {str(e)}"
            }
        except Exception as e:
            self.logger.error(f"❌ Unexpected error during PAYE upload: {str(e)}")
            return {
                "success": False,
                "error": f"Unexpected error: {str(e)}"
            }

    def _generate_paye_xml(self, paye_result: Dict[str, any], company_name: str, account_name: str) -> str:
        """
        Generate Tally XML for PAYE payment voucher.
        
        Args:
            paye_result: Processed PAYE data
            company_name: Company name for the voucher
            account_name: Bank account name for payment
            
        Returns:
            XML string for Tally import
        """
        try:
            # Get employee data and totals
            employee_data = paye_result.get("employee_data", [])
            if not employee_data:
                raise ValueError("No employee data found in PAYE result")
            
            total_paye = paye_result.get("total_paye", 0)
            total_sdl = paye_result.get("total_sdl", 0)
            total_amount = paye_result.get("total_amount", 0)
            
            # Format date
            voucher_date = self._format_date_for_tally(paye_result.get("date", ""))
            
            # Build XML
            xml_parts = [
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
                f'     <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>',
                '    </STATICVARIABLES>',
                '   </REQUESTDESC>',
                '   <REQUESTDATA>',
                '    <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                '     <VOUCHER REMOTEID="" VCHKEY="" VCHTYPE="Payment" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                '      <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '       <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '      </OLDAUDITENTRYIDS.LIST>',
                f'      <DATE>{voucher_date}</DATE>',
                f'      <NARRATION>{paye_result.get("narration", "PAYE and SDL payment")}</NARRATION>',
                f'      <PARTYLEDGERNAME>{account_name}</PARTYLEDGERNAME>',
                '      <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>',
                '      <VOUCHERNUMBER>1</VOUCHERNUMBER>',
                '      <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>',
                '      <VCHGSTCLASS/>',
                '      <ENTEREDBY>Administrator</ENTEREDBY>'
            ]
            
            # Add PAYE ledger entry if there's PAYE amount
            if total_paye > 0:
                xml_parts.extend([
                    '      <ALLLEDGERENTRIES.LIST>',
                    '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                    '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                    '       </OLDAUDITENTRYIDS.LIST>',
                    '       <LEDGERNAME>PAYE</LEDGERNAME>',
                    '       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'       <AMOUNT>-{total_paye:.2f}</AMOUNT>',
                    '       <CATEGORYALLOCATIONS.LIST>',
                    '        <CATEGORY>Primary Cost Category</CATEGORY>',
                    '        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
                ])
                
                # Add cost center allocations for each employee's PAYE
                for emp in employee_data:
                    if emp.get('paye', 0) > 0:
                        xml_parts.extend([
                            '        <COSTCENTREALLOCATIONS.LIST>',
                            f'         <NAME>{emp.get("employee_name", "Unknown")}</NAME>',
                            f'         <AMOUNT>-{emp.get("paye", 0):.2f}</AMOUNT>',
                            '        </COSTCENTREALLOCATIONS.LIST>'
                        ])
                
                xml_parts.append('       </CATEGORYALLOCATIONS.LIST>')
                xml_parts.append('      </ALLLEDGERENTRIES.LIST>')
            
            # Add SDL ledger entry if there's SDL amount
            if total_sdl > 0:
                xml_parts.extend([
                    '      <ALLLEDGERENTRIES.LIST>',
                    '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                    '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                    '       </OLDAUDITENTRYIDS.LIST>',
                    '       <LEDGERNAME>Skill &amp; Development Levy</LEDGERNAME>',
                    '       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'       <AMOUNT>-{total_sdl:.2f}</AMOUNT>',
                    '       <CATEGORYALLOCATIONS.LIST>',
                    '        <CATEGORY>Primary Cost Category</CATEGORY>',
                    '        <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>'
                ])
                
                # Add cost center allocations for each employee's SDL
                for emp in employee_data:
                    if emp.get('sdl', 0) > 0:
                        xml_parts.extend([
                            '        <COSTCENTREALLOCATIONS.LIST>',
                            f'         <NAME>{emp.get("employee_name", "Unknown")}</NAME>',
                            f'         <AMOUNT>-{emp.get("sdl", 0):.2f}</AMOUNT>',
                            '        </COSTCENTREALLOCATIONS.LIST>'
                        ])
                
                xml_parts.append('       </CATEGORYALLOCATIONS.LIST>')
                xml_parts.append('      </ALLLEDGERENTRIES.LIST>')
            
            # Add Bank ledger entry (credit side)
            xml_parts.extend([
                '      <ALLLEDGERENTRIES.LIST>',
                '       <OLDAUDITENTRYIDS.LIST TYPE="Number">',
                '        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>',
                '       </OLDAUDITENTRYIDS.LIST>',
                f'       <LEDGERNAME>{account_name}</LEDGERNAME>',
                '       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                '       <ISPARTYLEDGER>Yes</ISPARTYLEDGER>',
                f'       <AMOUNT>{total_amount:.2f}</AMOUNT>',
                '      </ALLLEDGERENTRIES.LIST>',
                '     </VOUCHER>',
                '    </TALLYMESSAGE>',
                '   </REQUESTDATA>',
                '  </IMPORTDATA>',
                ' </BODY>',
                '</ENVELOPE>'
            ])
            
            xml_content = '\n'.join(xml_parts)
            self.logger.info(f"✅ Generated PAYE XML for {len(employee_data)} employees, Total: ₹{total_amount:,.2f}")
            return xml_content
            
        except Exception as e:
            self.logger.error(f"❌ Error generating PAYE XML: {str(e)}")
            raise

    def _parse_tally_response(self, xml_response: str) -> Dict[str, any]:
        """
        Parse Tally XML response to check for errors and extract LINEERROR messages.
        
        Args:
            xml_response: XML response string from Tally
            
        Returns:
            Dict with success status, error messages, and parsed information
        """
        try:
            # Print full response for debugging
            print("="*80)
            print("📥 TALLY RESPONSE XML:")
            print("="*80)
            print(xml_response)
            print("="*80)
            
            # Parse XML
            root = ET.fromstring(xml_response)
            
            # Check for LINEERROR
            lineerror = root.find('.//LINEERROR')
            if lineerror is not None and lineerror.text:
                lineerror_msg = lineerror.text.strip()
                self.logger.error(f"❌ LINEERROR from Tally: {lineerror_msg}")
                return {
                    "success": False,
                    "error": lineerror_msg,
                    "lineerror": lineerror_msg,
                    "response": xml_response,
                    "errors_count": 1,
                    "created_count": 0
                }
            
            # Check for CREATED count (success indicator)
            created = root.find('.//CREATED')
            if created is not None:
                created_count = int(created.text) if created.text else 0
                
                # If CREATED > 0, it's a success
                if created_count > 0:
                    self.logger.info(f"✅ Tally created {created_count} records")
                    return {
                        "success": True,
                        "created_count": created_count,
                        "response": xml_response
                    }
            
            # Check for other error indicators
            errors = root.find('.//ERRORS')
            if errors is not None:
                errors_count = int(errors.text) if errors.text else 0
                if errors_count > 0:
                    self.logger.error(f"❌ Tally reported {errors_count} errors")
                    return {
                        "success": False,
                        "error": f"Tally reported {errors_count} errors. Check the response for details.",
                        "errors_count": errors_count,
                        "response": xml_response,
                        "created_count": 0
                    }
            
            # If no CREATED and no ERRORS, check response content
            if 'does not exist' in xml_response.lower():
                return {
                    "success": False,
                    "error": "One or more ledgers/employees do not exist in Tally",
                    "response": xml_response,
                    "errors_count": 1,
                    "created_count": 0
                }
            
            # Default: assume success if no errors found
            return {
                "success": True,
                "message": "Request processed",
                "response": xml_response,
                "created_count": 0
            }
                
        except ET.ParseError as e:
            self.logger.error(f"❌ XML parsing error: {str(e)}")
            self.logger.debug(f"XML content: {xml_response[:500]}...")
            return {
                "success": False,
                "error": f"Invalid XML response from Tally: {str(e)}",
                "response": xml_response
            }
        except Exception as e:
            self.logger.error(f"❌ Error parsing Tally response: {str(e)}")
            return {
                "success": False,
                "error": f"Error parsing response: {str(e)}",
                "response": xml_response
            }

    def get_tally_info(self) -> Dict[str, any]:
        """
        Get basic information about the Tally instance.
        
        Returns:
            Dict containing Tally server information
        """
        # This could be expanded to get Tally version, license info, etc.
        # For now, just return connection status
        return self.test_connection()


# Example usage and testing
if __name__ == "__main__":
    # Test the API service
    api_service = TallyApiService()
    
    print("Testing Tally API Service...")
    print("="*50)
    
    # Test connection
    connection_result = api_service.test_connection()
    print(f"Connection Test: {connection_result}")
    
    # Get companies
    companies_result = api_service.get_companies()
    print(f"Companies Result: {companies_result}")
    
    if companies_result.get("success"):
        print("\nFound Companies:")
        for i, company in enumerate(companies_result.get("companies", []), 1):
            print(f"  {i}. {company}")
    else:
        print(f"Error: {companies_result.get('error')}")