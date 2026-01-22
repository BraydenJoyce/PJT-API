# QUICK START GUIDE
## SEC EDGAR Financial Data Extractor for Caterpillar

### 🚀 Get Started in 3 Steps

#### Step 1: Install Dependencies
```bash
pip install requests pandas openpyxl
```

#### Step 2: Update Your Email
Open `sec_edgar_extractor.py` and change line 389:
```python
YOUR_EMAIL = "your.actual.email@example.com"
```

#### Step 3: Run the Script
```bash
python sec_edgar_extractor.py
```

That's it! You'll get `caterpillar_financials.xlsx` with complete historical data.

---

## 📊 What You'll Get

### Excel File with 8 Sheets:

**Income Statement**
- Raw: All filings with metadata (Form, dates, fiscal periods)
- Annual: Clean pivot table (rows = line items, columns = years)

**Balance Sheet**
- Raw: Complete asset/liability/equity data
- Annual: Pivot view of annual balances

**Cash Flow Statement**
- Raw: Operating, investing, financing activities
- Annual: Annual cash flow trends

**Statement of Equity**
- Raw: Changes in equity components
- Annual: Year-over-year equity changes

---

## 🎯 Key Features

✅ **Complete History**: From first electronic filing to present
✅ **All Periods**: Annual (10-K) and Quarterly (10-Q) data
✅ **Organized**: Both raw data and clean pivot tables
✅ **Professional**: Excel formatting with headers and frozen panes
✅ **Metadata Rich**: Includes filing dates, accession numbers, fiscal periods

---

## 💡 Usage Examples

### Example 1: Basic Revenue Analysis
1. Open `caterpillar_financials.xlsx`
2. Go to "Income Statement - Annual" sheet
3. Find "Total Revenues" row
4. See revenues across all years

### Example 2: Trend Analysis
1. Open "Income Statement - Raw" sheet
2. Filter "Form" column to "10-Q" for quarterly data
3. Sort by "End_Date" 
4. Analyze quarterly revenue trends

### Example 3: Calculate Ratios
The Annual pivot sheets are perfect for ratio calculations:
- **Profit Margin**: Net Income / Total Revenues
- **ROE**: Net Income / Stockholders' Equity
- **Current Ratio**: Current Assets / Current Liabilities
- **Debt-to-Equity**: Total Liabilities / Stockholders' Equity

---

## 🔧 Customization

### Change Company
```python
# In sec_edgar_extractor.py, __init__ method:
self.cik = "0000012927"  # Deere & Company
self.company_name = "Deere & Company"
```

Find CIK at: https://www.sec.gov/edgar/searchedgar/companysearch

### Filter Years
```python
# In create_pivot_table method, add:
annual_df = annual_df[annual_df['Fiscal_Year'] >= 2015]
```

### Add Line Items
```python
# In _get_statement_items method, add to the dict:
'RevenueFromContractWithCustomerExcludingAssessedTax': 'Contract Revenue',
```

Find tags at: https://xbrlview.fasb.org/yeti/

---

## ⚡ Pro Tips

1. **Raw Data First**: Check the raw sheets to see what's available before pivoting
2. **Form Types**: 
   - 10-K = Annual reports (most reliable)
   - 10-Q = Quarterly reports (more frequent)
   - 8-K = Current events (less consistent)
3. **Missing Data**: Not all line items exist in all periods (accounting standards change)
4. **Date Formats**: End_Date is when the period ended, Filed_Date is when SEC received it

---

## 🏢 Perfect for Bobcat Analysis

Since you're doing competitive intelligence at Bobcat:

### Compare Caterpillar Data With:
- **Deere (John Deere)**: CIK 0000012927
- **Doosan Bobcat**: CIK 0001773910
- **Komatsu**: CIK 0001121357

### Key Metrics to Track:
- Revenue growth rates
- Operating margins
- R&D spending (% of revenue)
- Capital expenditures
- Geographic segment data
- Product line breakdowns

### Quarterly Tracking:
Filter to 10-Q forms in the raw data sheets to see:
- Seasonal patterns
- Quarter-over-quarter changes
- Market response to product launches

---

## 📞 Need Help?

**Script Issues**: Check the console output - logging shows exactly what's happening

**Missing Data**: Some line items don't exist for all companies - this is normal

**Rate Limits**: Script includes delays to respect SEC limits (10 req/sec)

**Excel Errors**: Close the file before re-running the script

---

## 🎓 Learning Resources

- SEC EDGAR API: https://www.sec.gov/edgar/sec-api-documentation
- XBRL Taxonomy: https://xbrlview.fasb.org/yeti/
- Company Search: https://www.sec.gov/edgar/searchedgar/companysearch
- Financial Statement Guide: https://www.sec.gov/oiea/investor-alerts-and-bulletins/ib_beginnerguide

---

**Ready to analyze Caterpillar's complete financial history? Run the script and start exploring! 🚀**
