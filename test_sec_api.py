"""
Test script to verify SEC EDGAR extraction works
This fetches a small sample of data to validate functionality
"""

import requests
import pandas as pd
from datetime import datetime
import time
import json

def test_sec_api():
    """Test basic SEC API connectivity and data retrieval"""
    
    # Caterpillar CIK
    cik = "0000018230"
    base_url = "https://data.sec.gov"
    headers = {
        'User-Agent': 'test@example.com',  # SEC requires this
        'Accept-Encoding': 'gzip, deflate',
        'Host': 'data.sec.gov'
    }
    
    print("Testing SEC EDGAR API connection...")
    print("="*60)
    
    # Test 1: Fetch company facts
    url = f"{base_url}/api/xbrl/companyfacts/CIK{cik}.json"
    
    try:
        print(f"\n1. Fetching company facts for Caterpillar (CIK: {cik})")
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()
        
        print(f"   ✓ Successfully retrieved data")
        print(f"   Company: {data.get('entityName', 'N/A')}")
        print(f"   CIK: {data.get('cik', 'N/A')}")
        
        # Test 2: Check available facts
        facts = data.get('facts', {})
        us_gaap = facts.get('us-gaap', {})
        
        print(f"\n2. Available US-GAAP tags: {len(us_gaap)}")
        
        # Test 3: Check for key financial metrics
        print(f"\n3. Checking for key financial statement items:")
        
        key_items = {
            'Revenues': 'Revenue',
            'NetIncomeLoss': 'Net Income',
            'Assets': 'Total Assets',
            'Liabilities': 'Total Liabilities',
            'StockholdersEquity': 'Stockholders Equity',
            'NetCashProvidedByUsedInOperatingActivities': 'Operating Cash Flow',
            'EarningsPerShareBasic': 'EPS (Basic)'
        }
        
        found_items = 0
        for tag, label in key_items.items():
            if tag in us_gaap:
                found_items += 1
                # Get the number of data points
                units = us_gaap[tag].get('units', {})
                usd_data = units.get('USD', [])
                print(f"   ✓ {label}: {len(usd_data)} data points")
            else:
                print(f"   ✗ {label}: Not found")
        
        print(f"\n4. Summary:")
        print(f"   Found {found_items}/{len(key_items)} key financial items")
        
        # Test 4: Sample data extraction
        if 'Revenues' in us_gaap:
            print(f"\n5. Sample Revenue data (last 5 entries):")
            revenue_data = us_gaap['Revenues']['units']['USD']
            
            # Get last 5 annual filings (10-K)
            annual_revenue = [r for r in revenue_data if r.get('form') == '10-K'][-5:]
            
            for record in annual_revenue:
                value = record.get('val', 0)
                end_date = record.get('end', 'N/A')
                fy = record.get('fy', 'N/A')
                
                # Format value in billions
                value_b = value / 1_000_000_000 if value else 0
                
                print(f"   FY{fy} (ending {end_date}): ${value_b:,.2f}B")
        
        print("\n" + "="*60)
        print("✓ Test completed successfully!")
        print("The SEC EDGAR API is working and data is available.")
        print("="*60)
        
        return True
        
    except requests.exceptions.RequestException as e:
        print(f"\n✗ Error: {e}")
        print("\nPossible issues:")
        print("  - No internet connection")
        print("  - SEC servers are down")
        print("  - Rate limiting (too many requests)")
        return False
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        return False


if __name__ == "__main__":
    success = test_sec_api()
    
    if success:
        print("\n✓ You can now run the full extraction script!")
        print("  Remember to update YOUR_EMAIL in sec_edgar_extractor.py")
    else:
        print("\n✗ Please check your internet connection and try again")
