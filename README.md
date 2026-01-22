# PJT-API (Caterpillar)
SEC EDGAR Financial Data Extractor for Caterpillar Inc.

Overview

This Python script extracts Caterpillar's complete financial statement history from the SEC EDGAR API and exports it to organized Excel files.
Features

✅ Complete historical data from SEC EDGAR
✅ All major financial statements:

Income Statement
Balance Sheet
Cash Flow Statement
Statement of Stockholders' Equity


✅ Two views per statement:

Raw Data: All filings with complete metadata
Annual Pivot: Clean annual view (10-K only)


✅ Professional Excel formatting with headers and column widths
✅ Comprehensive logging to track progress
✅ Respects SEC API rate limits

Requirements

Python 3.7 or higher
Internet connection
Valid email address (required by SEC)

Installation
Step 1: Install Dependencies
bashpip install -r requirements.txt
Or install individually:
bashpip install requests pandas openpyxl
Step 2: Update Email Address
Open sec_edgar_extractor.py and update line 451:
pythonYOUR_EMAIL = "your.email@example.com"  # Replace with your actual email
Important: The SEC requires a valid email address in the User-Agent header. This helps them contact you if there are issues with your requests.
Usage
Basic Usage
bashpython sec_edgar_extractor.py
This will create caterpillar_financials.xlsx in the current directory.
Custom Output Filename
Modify the main() function:
pythonoutput_file = extractor.export_to_excel('my_custom_filename.xlsx')
Output Structure
The Excel file contains 8 sheets:
1. Income Statement - Raw
Complete income statement data with columns:

Line_Item: Human-readable label
XBRL_Tag: Technical XBRL tag name
Value: The actual dollar value
End_Date: Period end date
Start_Date: Period start date
Filed_Date: When filed with SEC
Form: Filing type (10-K, 10-Q, etc.)
Fiscal_Year: Fiscal year
Fiscal_Period: Fiscal period (FY, Q1, Q2, etc.)
Accession_Number: SEC accession number
Frame: Standardized reporting frame

2. Income Statement - Annual
Pivot table showing annual values (10-K filings only):

Rows: Line items (Revenue, Net Income, etc.)
Columns: Fiscal year end dates
Values: Dollar amounts

3-4. Balance Sheet (Raw + Annual)
Complete balance sheet data including:

Current and non-current assets
Current and non-current liabilities
Stockholders' equity components

5-6. Cash Flow Statement (Raw + Annual)
Cash flow data including:

Operating activities
Investing activities
Financing activities

7-8. Statement of Equity (Raw + Annual)
Equity statement data including:

Beginning and ending equity balances
Changes in common stock, retained earnings, AOCI, etc.

Data Notes
What's Included

All historical data: From Caterpillar's first SEC electronic filing to present
Multiple filing types: 10-K (annual), 10-Q (quarterly), 8-K (current reports)
Standardized XBRL tags: Uses US-GAAP taxonomy for consistency

Data Quality

Values are in USD
Dates are properly formatted
Missing values appear as NaN (you can filter these in Excel)
Some line items may not appear in older filings (accounting standards change)

Annual vs Raw Data

Annual sheets: Only 10-K filings, cleaner view for year-over-year analysis
Raw sheets: All filings, useful for quarterly trends and detailed analysis

Customization
Adding More Companies
Change the CIK (Central Index Key) in the __init__ method:
pythonself.cik = "0000018230"  # Current: Caterpillar
self.company_name = "Caterpillar Inc."
Find CIK numbers at: https://www.sec.gov/edgar/searchedgar/companysearch
Adding More Line Items
Edit the _get_statement_items() method to add XBRL tags:
python'NewXBRLTag': 'Display Name',
Find XBRL tags at: https://xbrlview.fasb.org/yeti/
Filtering to Specific Years
Add filtering in the create_pivot_table() method:
pythonannual_df = annual_df[annual_df['Fiscal_Year'] >= 2010]
Troubleshooting
"No data found" errors

Some line items may not exist for all companies
Older filings may use different XBRL tags
Check the raw data sheets to see what's available

Rate limiting errors

The script includes 0.1 second delays between requests
SEC allows up to 10 requests per second
If you get rate limited, the script will log an error

Missing quarters in raw data

Some companies don't file 10-Q for Q4 (they file 10-K instead)
This is normal and expected

Excel file errors

Make sure the file isn't open in Excel when running the script
Close the file before re-running

Advanced Usage
Programmatic Access
pythonfrom sec_edgar_extractor import SECEdgarExtractor

# Initialize
extractor = SECEdgarExtractor(email="your@email.com")

# Get raw data
facts = extractor.get_company_facts()

# Extract specific statement
income_df = extractor.extract_financial_statement_data(facts, 'income')

# Create pivot
income_pivot = extractor.create_pivot_table(income_df, 'income')

# Export
income_df.to_csv('income_statement.csv', index=False)
Processing Multiple Companies
pythoncompanies = {
    '0000018230': 'Caterpillar',
    '0000012927': 'Deere',
    '0001773910': 'Doosan Bobcat'
}

for cik, name in companies.items():
    extractor = SECEdgarExtractor(email="your@email.com")
    extractor.cik = cik
    extractor.company_name = name
    extractor.export_to_excel(f'{name.lower()}_financials.xlsx')
SEC EDGAR API Documentation

Official API docs: https://www.sec.gov/edgar/sec-api-documentation
Company facts endpoint: https://data.sec.gov/api/xbrl/companyfacts/
Rate limits: 10 requests per second (enforced by SEC)

License
This script is provided as-is for educational and research purposes.
Support
For issues related to:

Script functionality: Check the logs for detailed error messages
SEC data: Visit https://www.sec.gov/os/accessing-edgar-data
XBRL tags: Visit https://xbrlview.fasb.org/yeti/

Changelog

v1.0 (2026-01): Initial release with all four major financial statements
