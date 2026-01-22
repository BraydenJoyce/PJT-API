"""
SEC EDGAR Financial Data Extractor for Caterpillar Inc.
Extracts complete financial statement history and exports to Excel files.
"""

import requests
import pandas as pd
from datetime import datetime
import time
import json
import logging
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class SECEdgarExtractor:
    """Extract financial data from SEC EDGAR API"""
    
    def __init__(self, email):
        """
        Initialize the extractor
        
        Args:
            email: Your email for SEC API user agent (required by SEC)
        """
        self.base_url = "https://data.sec.gov"
        self.headers = {
            'User-Agent': f'{email}',
            'Accept-Encoding': 'gzip, deflate',
            'Host': 'data.sec.gov'
        }
        self.cik = "0000018230"  # Caterpillar Inc. CIK
        self.company_name = "Caterpillar Inc."
        
    def get_company_facts(self):
        """
        Retrieve all company facts from SEC EDGAR API
        
        Returns:
            dict: Complete company facts data
        """
        url = f"{self.base_url}/api/xbrl/companyfacts/CIK{self.cik}.json"
        
        logger.info(f"Fetching company facts for {self.company_name} (CIK: {self.cik})")
        
        try:
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            time.sleep(0.1)  # Rate limiting - be respectful to SEC servers
            
            data = response.json()
            logger.info("Successfully retrieved company facts")
            return data
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error fetching data: {e}")
            raise
    
    def extract_financial_statement_data(self, facts_data, statement_type):
        """
        Extract specific financial statement data
        
        Args:
            facts_data: Complete company facts from API
            statement_type: Type of statement ('income', 'balance', 'cashflow', 'equity')
            
        Returns:
            pd.DataFrame: Organized financial data
        """
        logger.info(f"Extracting {statement_type} statement data...")
        
        # Define line items for each statement type
        statement_items = self._get_statement_items(statement_type)
        
        # Extract US-GAAP facts
        us_gaap = facts_data.get('facts', {}).get('us-gaap', {})
        dei = facts_data.get('facts', {}).get('dei', {})
        
        all_records = []
        
        for item_name, item_label in statement_items.items():
            if item_name in us_gaap:
                item_data = us_gaap[item_name]
                units = item_data.get('units', {})
                
                # Process USD data
                if 'USD' in units:
                    for record in units['USD']:
                        all_records.append({
                            'Line_Item': item_label,
                            'XBRL_Tag': item_name,
                            'Value': record.get('val'),
                            'End_Date': record.get('end'),
                            'Start_Date': record.get('start'),
                            'Filed_Date': record.get('filed'),
                            'Form': record.get('form'),
                            'Fiscal_Year': record.get('fy'),
                            'Fiscal_Period': record.get('fp'),
                            'Accession_Number': record.get('accn'),
                            'Frame': record.get('frame')
                        })
        
        if not all_records:
            logger.warning(f"No data found for {statement_type} statement")
            return pd.DataFrame()
        
        df = pd.DataFrame(all_records)
        
        # Convert dates
        df['End_Date'] = pd.to_datetime(df['End_Date'])
        df['Filed_Date'] = pd.to_datetime(df['Filed_Date'])
        df['Start_Date'] = pd.to_datetime(df['Start_Date'])
        
        # Sort by end date and line item
        df = df.sort_values(['End_Date', 'Line_Item'], ascending=[False, True])
        
        logger.info(f"Extracted {len(df)} records for {statement_type} statement")
        return df
    
    def _get_statement_items(self, statement_type):
        """
        Get relevant line items for each statement type
        
        Args:
            statement_type: Type of financial statement
            
        Returns:
            dict: Mapping of XBRL tags to readable labels
        """
        if statement_type == 'income':
            return {
                # Revenue
                'Revenues': 'Total Revenues',
                'RevenueFromContractWithCustomerExcludingAssessedTax': 'Revenue from Contracts',
                'SalesRevenueNet': 'Sales Revenue (Net)',
                
                # Costs and expenses
                'CostOfRevenue': 'Cost of Revenue',
                'CostOfGoodsAndServicesSold': 'Cost of Goods Sold',
                'SellingGeneralAndAdministrativeExpense': 'SG&A Expense',
                'ResearchAndDevelopmentExpense': 'R&D Expense',
                'OperatingExpenses': 'Operating Expenses',
                
                # Operating income
                'OperatingIncomeLoss': 'Operating Income',
                'GrossProfit': 'Gross Profit',
                
                # Other income/expenses
                'InterestExpense': 'Interest Expense',
                'InterestIncomeExpenseNet': 'Interest Income (Expense), Net',
                'OtherNonoperatingIncomeExpense': 'Other Income (Expense)',
                'NonoperatingIncomeExpense': 'Non-operating Income',
                
                # Income before taxes
                'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest': 'Income Before Taxes',
                
                # Taxes
                'IncomeTaxExpenseBenefit': 'Income Tax Expense',
                
                # Net income
                'NetIncomeLoss': 'Net Income',
                'ProfitLoss': 'Profit (Loss)',
                'NetIncomeLossAvailableToCommonStockholdersBasic': 'Net Income Available to Common',
                
                # EPS
                'EarningsPerShareBasic': 'EPS - Basic',
                'EarningsPerShareDiluted': 'EPS - Diluted',
                'WeightedAverageNumberOfSharesOutstandingBasic': 'Shares Outstanding - Basic',
                'WeightedAverageNumberOfDilutedSharesOutstanding': 'Shares Outstanding - Diluted',
            }
        
        elif statement_type == 'balance':
            return {
                # Assets - Current
                'AssetsCurrent': 'Total Current Assets',
                'CashAndCashEquivalentsAtCarryingValue': 'Cash and Cash Equivalents',
                'ShortTermInvestments': 'Short-term Investments',
                'AccountsReceivableNetCurrent': 'Accounts Receivable (Net)',
                'InventoryNet': 'Inventory',
                'PrepaidExpenseAndOtherAssetsCurrent': 'Prepaid and Other Current Assets',
                
                # Assets - Non-current
                'PropertyPlantAndEquipmentNet': 'Property, Plant & Equipment (Net)',
                'Goodwill': 'Goodwill',
                'IntangibleAssetsNetExcludingGoodwill': 'Intangible Assets (Net)',
                'LongTermInvestments': 'Long-term Investments',
                'OtherAssetsNoncurrent': 'Other Non-current Assets',
                
                # Total Assets
                'Assets': 'Total Assets',
                
                # Liabilities - Current
                'LiabilitiesCurrent': 'Total Current Liabilities',
                'AccountsPayableCurrent': 'Accounts Payable',
                'ShortTermBorrowings': 'Short-term Debt',
                'LongTermDebtCurrent': 'Current Portion of Long-term Debt',
                'AccruedLiabilitiesCurrent': 'Accrued Liabilities',
                
                # Liabilities - Non-current
                'LongTermDebtNoncurrent': 'Long-term Debt',
                'DeferredTaxLiabilitiesNoncurrent': 'Deferred Tax Liabilities',
                'OtherLiabilitiesNoncurrent': 'Other Non-current Liabilities',
                
                # Total Liabilities
                'Liabilities': 'Total Liabilities',
                
                # Equity
                'StockholdersEquity': "Stockholders' Equity",
                'CommonStockValue': 'Common Stock',
                'AdditionalPaidInCapital': 'Additional Paid-in Capital',
                'RetainedEarningsAccumulatedDeficit': 'Retained Earnings',
                'TreasuryStockValue': 'Treasury Stock',
                'AccumulatedOtherComprehensiveIncomeLossNetOfTax': 'Accumulated Other Comprehensive Income',
                
                # Total Liabilities and Equity
                'LiabilitiesAndStockholdersEquity': 'Total Liabilities and Equity',
            }
        
        elif statement_type == 'cashflow':
            return {
                # Operating Activities
                'NetCashProvidedByUsedInOperatingActivities': 'Net Cash from Operating Activities',
                'DepreciationDepletionAndAmortization': 'Depreciation and Amortization',
                'DeferredIncomeTaxExpenseBenefit': 'Deferred Income Taxes',
                'IncreaseDecreaseInAccountsReceivable': 'Change in Accounts Receivable',
                'IncreaseDecreaseInInventories': 'Change in Inventories',
                'IncreaseDecreaseInAccountsPayable': 'Change in Accounts Payable',
                'IncreaseDecreaseInAccruedLiabilities': 'Change in Accrued Liabilities',
                
                # Investing Activities
                'NetCashProvidedByUsedInInvestingActivities': 'Net Cash from Investing Activities',
                'PaymentsToAcquirePropertyPlantAndEquipment': 'Capital Expenditures',
                'PaymentsToAcquireBusinessesNetOfCashAcquired': 'Acquisitions',
                'ProceedsFromSaleOfPropertyPlantAndEquipment': 'Proceeds from Asset Sales',
                'PaymentsToAcquireInvestments': 'Purchase of Investments',
                'ProceedsFromSaleOfInvestments': 'Proceeds from Sale of Investments',
                
                # Financing Activities
                'NetCashProvidedByUsedInFinancingActivities': 'Net Cash from Financing Activities',
                'ProceedsFromIssuanceOfLongTermDebt': 'Proceeds from Debt Issuance',
                'RepaymentsOfLongTermDebt': 'Repayment of Long-term Debt',
                'ProceedsFromIssuanceOfCommonStock': 'Proceeds from Stock Issuance',
                'PaymentsForRepurchaseOfCommonStock': 'Stock Repurchases',
                'PaymentsOfDividends': 'Dividends Paid',
                
                # Net change
                'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect': 'Net Change in Cash',
                'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents': 'Cash at End of Period',
            }
        
        elif statement_type == 'equity':
            return {
                # Beginning balance
                'StockholdersEquity': "Stockholders' Equity",
                
                # Components
                'CommonStockValue': 'Common Stock',
                'AdditionalPaidInCapital': 'Additional Paid-in Capital',
                'RetainedEarningsAccumulatedDeficit': 'Retained Earnings',
                'TreasuryStockValue': 'Treasury Stock',
                'AccumulatedOtherComprehensiveIncomeLossNetOfTax': 'AOCI',
                
                # Changes
                'NetIncomeLoss': 'Net Income',
                'OtherComprehensiveIncomeLossNetOfTax': 'Other Comprehensive Income',
                'StockIssuedDuringPeriodValueNewIssues': 'Stock Issued',
                'StockRepurchasedDuringPeriodValue': 'Stock Repurchased',
                'Dividends': 'Dividends Declared',
                'DividendsCommonStock': 'Common Stock Dividends',
            }
        
        else:
            logger.warning(f"Unknown statement type: {statement_type}")
            return {}
    
    def create_pivot_table(self, df, statement_type):
        """
        Create a pivot table view of the financial data
        
        Args:
            df: DataFrame with financial data
            statement_type: Type of statement
            
        Returns:
            pd.DataFrame: Pivoted data showing values across periods
        """
        if df.empty:
            return pd.DataFrame()
        
        # Filter to annual reports only (10-K)
        annual_df = df[df['Form'] == '10-K'].copy()
        
        if annual_df.empty:
            logger.warning(f"No 10-K data found for {statement_type}")
            return df
        
        # Create pivot table
        pivot = annual_df.pivot_table(
            index='Line_Item',
            columns='End_Date',
            values='Value',
            aggfunc='first'
        )
        
        # Sort columns by date (most recent first)
        pivot = pivot[sorted(pivot.columns, reverse=True)]
        
        # Format column names
        pivot.columns = [col.strftime('%Y-%m-%d') if isinstance(col, datetime) else str(col) 
                        for col in pivot.columns]
        
        return pivot
    
    def format_excel_sheet(self, writer, sheet_name, df):
        """
        Apply formatting to Excel sheet
        
        Args:
            writer: ExcelWriter object
            sheet_name: Name of the sheet
            df: DataFrame that was written
        """
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Header formatting
        header_format = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        
        for col_num, value in enumerate(df.columns.values, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_format
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Freeze header row
        worksheet.freeze_panes = 'A2'
    
    def export_to_excel(self, output_filename='caterpillar_financials.xlsx'):
        """
        Main function to extract all data and export to Excel
        
        Args:
            output_filename: Name of output Excel file
        """
        logger.info("="*60)
        logger.info("Starting SEC EDGAR data extraction for Caterpillar Inc.")
        logger.info("="*60)
        
        # Get all company facts
        facts_data = self.get_company_facts()
        
        # Create Excel writer
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            
            # 1. Income Statement
            logger.info("\n" + "="*60)
            logger.info("PROCESSING INCOME STATEMENT")
            logger.info("="*60)
            income_df = self.extract_financial_statement_data(facts_data, 'income')
            if not income_df.empty:
                # Raw data
                income_df.to_excel(writer, sheet_name='Income Statement - Raw', index=False)
                self.format_excel_sheet(writer, 'Income Statement - Raw', income_df)
                
                # Pivot view
                income_pivot = self.create_pivot_table(income_df, 'income')
                if not income_pivot.empty:
                    income_pivot.to_excel(writer, sheet_name='Income Statement - Annual')
                    self.format_excel_sheet(writer, 'Income Statement - Annual', income_pivot)
            
            # 2. Balance Sheet
            logger.info("\n" + "="*60)
            logger.info("PROCESSING BALANCE SHEET")
            logger.info("="*60)
            balance_df = self.extract_financial_statement_data(facts_data, 'balance')
            if not balance_df.empty:
                # Raw data
                balance_df.to_excel(writer, sheet_name='Balance Sheet - Raw', index=False)
                self.format_excel_sheet(writer, 'Balance Sheet - Raw', balance_df)
                
                # Pivot view
                balance_pivot = self.create_pivot_table(balance_df, 'balance')
                if not balance_pivot.empty:
                    balance_pivot.to_excel(writer, sheet_name='Balance Sheet - Annual')
                    self.format_excel_sheet(writer, 'Balance Sheet - Annual', balance_pivot)
            
            # 3. Cash Flow Statement
            logger.info("\n" + "="*60)
            logger.info("PROCESSING CASH FLOW STATEMENT")
            logger.info("="*60)
            cashflow_df = self.extract_financial_statement_data(facts_data, 'cashflow')
            if not cashflow_df.empty:
                # Raw data
                cashflow_df.to_excel(writer, sheet_name='Cash Flow - Raw', index=False)
                self.format_excel_sheet(writer, 'Cash Flow - Raw', cashflow_df)
                
                # Pivot view
                cashflow_pivot = self.create_pivot_table(cashflow_df, 'cashflow')
                if not cashflow_pivot.empty:
                    cashflow_pivot.to_excel(writer, sheet_name='Cash Flow - Annual')
                    self.format_excel_sheet(writer, 'Cash Flow - Annual', cashflow_pivot)
            
            # 4. Statement of Equity
            logger.info("\n" + "="*60)
            logger.info("PROCESSING STATEMENT OF EQUITY")
            logger.info("="*60)
            equity_df = self.extract_financial_statement_data(facts_data, 'equity')
            if not equity_df.empty:
                # Raw data
                equity_df.to_excel(writer, sheet_name='Equity - Raw', index=False)
                self.format_excel_sheet(writer, 'Equity - Raw', equity_df)
                
                # Pivot view
                equity_pivot = self.create_pivot_table(equity_df, 'equity')
                if not equity_pivot.empty:
                    equity_pivot.to_excel(writer, sheet_name='Equity - Annual')
                    self.format_excel_sheet(writer, 'Equity - Annual', equity_pivot)
        
        logger.info("\n" + "="*60)
        logger.info(f"✓ Export complete! File saved: {output_filename}")
        logger.info("="*60)
        
        return output_filename


def main():
    """Main execution function"""
    
    # IMPORTANT: Replace with your email address
    YOUR_EMAIL = "your.email@example.com"
    
    if YOUR_EMAIL == "your.email@example.com":
        print("\n" + "!"*60)
        print("IMPORTANT: Please update YOUR_EMAIL in the script")
        print("The SEC requires a valid email in the User-Agent header")
        print("!"*60 + "\n")
        return
    
    # Create extractor instance
    extractor = SECEdgarExtractor(email=YOUR_EMAIL)
    
    # Extract and export data
    output_file = extractor.export_to_excel('caterpillar_financials.xlsx')
    
    print(f"\n✓ Success! Financial data exported to: {output_file}")
    print("\nThe Excel file contains:")
    print("  • Income Statement (Raw data + Annual pivot)")
    print("  • Balance Sheet (Raw data + Annual pivot)")
    print("  • Cash Flow Statement (Raw data + Annual pivot)")
    print("  • Statement of Equity (Raw data + Annual pivot)")
    print("\nEach statement has complete historical data from SEC EDGAR!")


if __name__ == "__main__":
    main()
