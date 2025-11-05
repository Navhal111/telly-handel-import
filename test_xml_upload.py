"""
Quick test to verify XML is being sent to Tally correctly.
"""
from src.tally_api_service import TallyApiService

# Sample attendance XML
test_xml = """<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
 <HEADER>
  <TALLYREQUEST>Import Data</TALLYREQUEST>
 </HEADER>
 <BODY>
  <IMPORTDATA>
   <REQUESTDESC>
    <REPORTNAME>Vouchers</REPORTNAME>
    <STATICVARIABLES>
     <SVCURRENTCOMPANY>LIGHT</SVCURRENTCOMPANY>
    </STATICVARIABLES>
   </REQUESTDESC>
   <REQUESTDATA>
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
     <VOUCHER VCHTYPE="Attendance" ACTION="Create" OBJVIEW="Accounting Voucher View">
      <DATE>20251102</DATE>
      <NARRATION>Test attendance</NARRATION>
      <VOUCHERTYPENAME>Attendance</VOUCHERTYPENAME>
      <ATTENDANCEENTRIES.LIST>
       <NAME>Test Employee</NAME>
       <ATTENDANCETYPE>Present</ATTENDANCETYPE>
       <ATTDTYPETIMEVALUE>20</ATTDTYPETIMEVALUE>
      </ATTENDANCEENTRIES.LIST>
     </VOUCHER>
    </TALLYMESSAGE>
   </REQUESTDATA>
  </IMPORTDATA>
 </BODY>
</ENVELOPE>"""

def test_upload():
    """Test uploading XML to Tally."""
    print("Testing XML upload to Tally...")
    print("="*80)
    
    # Create API service
    api = TallyApiService(base_url="http://localhost:9000/")
    
    # Test connection first
    print("1. Testing connection...")
    conn_result = api.test_connection()
    print(f"   Connection: {conn_result}")
    print()
    
    # Upload XML
    print("2. Uploading test XML...")
    result = api.upload_xml_to_tally(test_xml, "Test Attendance")
    print()
    
    print("="*80)
    print("RESULT:")
    print(f"Success: {result.get('success')}")
    print(f"Error: {result.get('error', 'None')}")
    print(f"Line Error: {result.get('lineerror', 'None')}")
    print(f"Created: {result.get('created_count', 0)}")
    print(f"Errors: {result.get('errors_count', 0)}")
    print("="*80)

if __name__ == "__main__":
    test_upload()
