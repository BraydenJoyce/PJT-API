"""
SEC EDGAR Financial Data Extractor - Company Selection Menu
Allows user to select from multiple companies and extracts their financial data.
"""

import os
import sys
import subprocess
from pathlib import Path

# Company configurations
COMPANIES = {
    '1': {
        'name': 'Caterpillar Inc.',
        'ticker': 'CAT',
        'cik': '0000018230',
        'output_file': 'caterpillar_financials.xlsx',
        'script': 'caterpillar_extractor.py'
    },
    '2': {
        'name': 'Deere & Company',
        'ticker': 'DE',
        'cik': '0000315189',
        'output_file': 'deere_financials.xlsx',
        'script': 'deere_extractor.py'
    },
    '3': {
        'name': 'Toro Company',
        'ticker': 'TTC',
        'cik': '0000737758',
        'output_file': 'toro_financials.xlsx',
        'script': 'toro_extractor.py'
    },
    '4': {
        'name': 'Polaris Inc.',
        'ticker': 'PII',
        'cik': '0000931015',
        'output_file': 'polaris_financials.xlsx',
        'script': 'polaris_extractor.py'
    },
    '5': {
        'name': 'Case New Holland Industrial',
        'ticker': 'CNH',
        'cik': '0001567094',
        'output_file': 'cnh_financials.xlsx',
        'script': 'cnh_extractor.py'
    },
    '6': {
        'name': 'AGCO Corporation',
        'ticker': 'AGCO',
        'cik': '0000880266',
        'output_file': 'agco_financials.xlsx',
        'script': 'agco_extractor.py'
    },
    '7': {
        'name': 'Hyster-Yale Inc.',
        'ticker': 'HY',
        'cik': '0001173514',
        'output_file': 'hysteryale_financials.xlsx',
        'script': 'hysteryale_extractor.py'
    },
    '8': {
        'name': 'Ingersoll Rand Inc.',
        'ticker': 'IR',
        'cik': '0001699150',
        'output_file': 'ingersollrand_financials.xlsx',
        'script': 'ingersollrand_extractor.py'
    },
    '9': {
        'name': 'Generac Holdings Inc.',
        'ticker': 'GNRC',
        'cik': '0001474735',
        'output_file': 'generac_financials.xlsx',
        'script': 'generac_extractor.py'
    },
}

def display_menu():
    """Display the company selection menu"""
    print("\n" + "="*60)
    print("SEC EDGAR Financial Data Extractor")
    print("="*60)
    print("\nSelect a company to extract financial data:\n")
    
    for key, company in COMPANIES.items():
        print(f"  {key}. {company['name']} ({company['ticker']})")
    
    print("\n  Q. Quit")
    print("="*60)

def get_user_choice():
    """Get and validate user input"""
    while True:
        choice = input("\nEnter your choice (1-9 or Q): ").strip().upper()
        
        if choice == 'Q':
            print("\nExiting program. Goodbye!")
            sys.exit(0)
        
        if choice in COMPANIES:
            return choice
        
        print("Invalid choice. Please enter 1, 2, 3, 4, 5, 6, 7, 8, 9 or Q.")

def check_extractor_exists(company_info):
    """
    Check if the extractor script exists for the selected company
    
    Args:
        company_info: Dictionary containing company details
        
    Returns:
        bool: True if script exists, False otherwise
    """
    return os.path.exists(company_info['script'])

def run_extractor(company_info, email):
    """
    Run the SEC extractor for the selected company
    
    Args:
        company_info: Dictionary containing company details
        email: User's email for SEC API
    """
    print(f"\n{'='*60}")
    print(f"Extracting financial data for {company_info['name']} ({company_info['ticker']})")
    print(f"{'='*60}\n")
    
    # Check if extractor script exists
    if not check_extractor_exists(company_info):
        print(f"\n{'!'*60}")
        print(f"ERROR: Extractor script '{company_info['script']}' not found!")
        print(f"Please ensure the file is in the same directory as this script.")
        print(f"{'!'*60}\n")
        return
    
    # Run the extractor script
    result = subprocess.run(
        [sys.executable, company_info['script']],
        capture_output=False,
        text=True
    )
    
    if result.returncode == 0:
        print(f"\n{'='*60}")
        print(f"✓ SUCCESS!")
        print(f"{'='*60}")
        print(f"\nFinancial data for {company_info['name']} has been exported to:")
        print(f"  → {company_info['output_file']}")
        print(f"\nThe Excel file contains:")
        print(f"  • Income Statement (Raw data + Quarterly pivot)")
        print(f"  • Balance Sheet (Raw data + Quarterly pivot)")
        print(f"  • Cash Flow Statement (Raw data + Quarterly pivot)")
    else:
        print(f"\n✗ Error occurred during extraction. Please check the output above.")

def main():
    """Main execution function"""
    
    # Get user's email
    print("\n" + "="*60)
    print("SEC API Configuration")
    print("="*60)
    email = input("\nEnter your email address (required by SEC): ").strip()
    
    if not email or '@' not in email:
        print("\n✗ Invalid email address. Exiting.")
        sys.exit(1)
    
    # Main loop
    while True:
        display_menu()
        choice = get_user_choice()
        
        company = COMPANIES[choice]
        run_extractor(company, email)
        
        # Ask if user wants to extract another company
        print("\n" + "="*60)
        continue_choice = input("\nWould you like to extract data for another company? (Y/N): ").strip().upper()
        
        if continue_choice != 'Y':
            print("\nThank you for using the SEC EDGAR Financial Data Extractor!")
            print("Goodbye!\n")
            break

if __name__ == "__main__":
    main()
