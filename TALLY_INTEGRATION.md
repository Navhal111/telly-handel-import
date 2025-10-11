# Tally Integration Documentation

## Overview

The Excel File Processor now includes integration with Tally ERP to fetch company information via API before processing Excel files.

## New Features

### 1. Tally API Service (`src/tally_api_service.py`)

- **Base URL**: `http://localhost:9000/`
- **Function**: `get_companies()` - Fetches list of companies from Tally
- **XML Request Format**:

```xml
<ENVELOPE>
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
</ENVELOPE>
```

- **XML Response Format**:

```xml
<ENVELOPE>
    <COMPANYNAME.LIST>
        <COMPANYNAME>BKD</COMPANYNAME>
        <COMPANYNAME>LIGHT</COMPANYNAME>
    </COMPANYNAME.LIST>
</ENVELOPE>
```

### 2. New Application Flow

#### Startup Flow:

1. **Company Selection Screen**: Opens first with dropdown of available companies
2. **Company Loading**: Automatically fetches companies from Tally API
3. **Company Selection**: User selects a company from dropdown
4. **Main Screen**: Shows main Excel processing interface with selected company displayed at top

#### Main Screen Changes:

- **Company Display**: Shows selected company at top in green header
- **Change Company Button**: Allows user to return to company selection
- **Same Excel Processing**: All existing Excel processing functionality remains unchanged

### 3. API Service Features

- **Connection Testing**: Tests connectivity to Tally server
- **Error Handling**: Graceful handling of connection failures, timeouts, and parsing errors
- **XML Parsing**: Robust parsing of Tally XML response format
- **Logging**: Detailed logging of API operations

### 4. Files Modified/Created

#### New Files:

- `src/tally_api_service.py` - Main API service class
- `test_tally_api.py` - API testing script
- `test_complete_flow.py` - Complete integration testing
- `mock_tally_api_service.py` - Mock service for testing without Tally

#### Modified Files:

- `main.py` - Added company selection screen and updated UI flow
- `requirements.txt` - Added requests library dependency

## Usage Instructions

### Prerequisites

1. Tally ERP should be running and accessible at `http://localhost:9000/`
2. Install required dependencies: `pip install -r requirements.txt`

### Running the Application

1. Start Tally ERP and ensure it's accessible
2. Run: `python main.py`
3. Select a company from the dropdown
4. Proceed to Excel processing as before

### Testing

- **API Only**: `python test_tally_api.py`
- **Complete Flow**: `python test_complete_flow.py`

## Error Handling

- **Tally Not Available**: Shows error message with retry option
- **No Companies Found**: Displays appropriate message
- **Connection Timeout**: Graceful timeout handling with user feedback
- **XML Parsing Errors**: Logs errors and shows user-friendly messages

## Configuration

The API service can be configured with different base URLs by modifying the `TallyApiService` initialization in `main.py`.

## Future Enhancements

This integration provides the foundation for:

- Company-specific Excel processing rules
- Direct data export to selected Tally company
- Additional Tally data integration (ledgers, vouchers, etc.)
- Multi-company batch processing
