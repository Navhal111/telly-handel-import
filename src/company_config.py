import json
import os
import sys
from typing import List, Dict


class CompanyConfigManager:
    """Manage local company names for Tally Prime (CRUD operations)."""
    
    def __init__(self, config_file: str = "tally_prime_companies.json"):
        """Initialize config manager with config file path in persistent location."""
        # Get the directory where the executable/script is located
        if getattr(sys, 'frozen', False):
            # Running as compiled executable (PyInstaller)
            # Save in the same directory as the EXE
            app_dir = os.path.dirname(sys.executable)
        else:
            # Running as script
            app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        self.config_file = os.path.join(app_dir, config_file)
        print(f"📁 Company config file location: {self.config_file}")
        self._ensure_config_exists()
    
    def _ensure_config_exists(self):
        """Create config file if it doesn't exist."""
        if not os.path.exists(self.config_file):
            self._save_config({"companies": []})
    
    def _load_config(self) -> Dict:
        """Load config from JSON file."""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
            return {"companies": []}
    
    def _save_config(self, config: Dict):
        """Save config to JSON file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving config: {e}")
    
    def get_companies(self) -> List[str]:
        """Get list of saved company names."""
        config = self._load_config()
        return config.get("companies", [])
    
    def add_company(self, company_name: str) -> bool:
        """
        Add a new company name to the list.
        
        Args:
            company_name: Name of the company to add
            
        Returns:
            True if added successfully, False if already exists
        """
        company_name = company_name.strip()
        if not company_name:
            return False
        
        config = self._load_config()
        companies = config.get("companies", [])
        
        # Check if company already exists (case-insensitive)
        if any(c.lower() == company_name.lower() for c in companies):
            return False
        
        companies.append(company_name)
        config["companies"] = sorted(companies)  # Keep sorted
        self._save_config(config)
        return True
    
    def delete_company(self, company_name: str) -> bool:
        """
        Delete a company name from the list.
        
        Args:
            company_name: Name of the company to delete
            
        Returns:
            True if deleted successfully, False if not found
        """
        config = self._load_config()
        companies = config.get("companies", [])
        
        # Find and remove company (case-insensitive)
        original_count = len(companies)
        companies = [c for c in companies if c.lower() != company_name.lower()]
        
        if len(companies) < original_count:
            config["companies"] = companies
            self._save_config(config)
            return True
        
        return False
    
    def update_company(self, old_name: str, new_name: str) -> bool:
        """
        Update/rename a company name.
        
        Args:
            old_name: Current name of the company
            new_name: New name for the company
            
        Returns:
            True if updated successfully, False if failed
        """
        new_name = new_name.strip()
        if not new_name:
            return False
        
        config = self._load_config()
        companies = config.get("companies", [])
        
        # Find and update company
        for i, c in enumerate(companies):
            if c.lower() == old_name.lower():
                companies[i] = new_name
                config["companies"] = sorted(companies)
                self._save_config(config)
                return True
        
        return False
    
    def clear_all(self) -> bool:
        """Clear all saved companies."""
        try:
            self._save_config({"companies": []})
            return True
        except:
            return False
