"""
Test script for the Tally API Service
"""
import sys
import os

# Add the src directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.tally_api_service import TallyApiService


def test_api_service():
    """Test the Tally API service functionality."""
    print("="*60)
    print("🧪 TESTING TALLY API SERVICE")
    print("="*60)
    
    # Initialize API service
    print("\n1. Initializing API Service...")
    api_service = TallyApiService()
    print(f"   ✅ Base URL: {api_service.base_url}")
    
    # Test connection
    print("\n2. Testing Connection...")
    connection_result = api_service.test_connection()
    print(f"   Result: {connection_result}")
    
    if connection_result.get("success"):
        print("   ✅ Connection successful!")
    else:
        print("   ❌ Connection failed!")
        print("   💡 Make sure Tally is running and accessible at http://localhost:9000/")
    
    # Test getting companies
    print("\n3. Getting Company List...")
    companies_result = api_service.get_companies()
    
    print(f"   Success: {companies_result.get('success', False)}")
    
    if companies_result.get("success"):
        companies = companies_result.get("companies", [])
        print(f"   ✅ Found {len(companies)} companies:")
        for i, company in enumerate(companies, 1):
            print(f"      {i}. {company}")
    else:
        error = companies_result.get("error", "Unknown error")
        print(f"   ❌ Error: {error}")
    
    print("\n" + "="*60)
    print("🏁 TEST COMPLETED")
    print("="*60)
    
    return companies_result


if __name__ == "__main__":
    test_api_service()