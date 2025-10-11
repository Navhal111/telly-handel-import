"""
Mock Tally API Service for testing when actual Tally is not available
"""
from typing import List, Dict
import time
import random


class MockTallyApiService:
    """Mock version of TallyApiService for testing purposes."""
    
    def __init__(self, base_url: str = "http://localhost:9000/"):
        """Initialize the mock Tally API service."""
        self.base_url = base_url
        self.mock_companies = [
            "LIGHT",
            "BKD", 
            "ACME Corporation",
            "Tech Solutions Ltd",
            "Global Trading Co"
        ]
    
    def get_companies(self) -> Dict[str, any]:
        """
        Mock implementation to get list of companies.
        
        Returns:
            Dict containing success status and mock companies list
        """
        # Simulate network delay
        time.sleep(0.5)
        
        # Simulate occasional failures for testing
        if random.random() < 0.1:  # 10% chance of failure
            return {
                "success": False,
                "companies": [],
                "error": "Mock connection timeout"
            }
        
        return {
            "success": True,
            "companies": self.mock_companies,
            "message": f"Found {len(self.mock_companies)} companies"
        }
    
    def test_connection(self) -> Dict[str, any]:
        """
        Mock connection test.
        
        Returns:
            Dict containing connection test results
        """
        # Simulate network delay
        time.sleep(0.2)
        
        return {
            "success": True,
            "status_code": 200,
            "message": "Mock connection successful"
        }


# You can use this for testing by replacing the import in main.py
# from src.tally_api_service import TallyApiService
# with:
# from mock_tally_api_service import MockTallyApiService as TallyApiService