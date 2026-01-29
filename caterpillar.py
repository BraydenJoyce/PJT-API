"""
SEC EDGAR XBRL Parser - Complete Financial Statement Extractor
Extracts both consolidated and segment-level data with Q4 calculations
"""

import requests
import pandas as pd
from datetime import datetime
import time
import logging
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from collections import OrderedDict

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class ComprehensiveXBRLExtractor:
    """Extract complete financial statements with segment breakdowns"""
    
    def __init__(self, email, cik, company_name, ticker):
        """
        Initialize the extractor
        
        Args:
            email: Your email for SEC API user agent
            cik: Company CIK number (with leading zeros)
            company_name: Company name for logging
            ticker: Company ticker symbol (lowercase, for URL construction)
        """
        self.base_url = "https://data.sec.gov"
        self.sec_archives = "https://www.sec.gov/Archives/edgar/data"
        self.headers = {
            'User-Agent': f'{email}',
            'Accept-Encoding': 'gzip, deflate'
        }
        self.cik = cik
        self.cik_int = str(int(cik))
        self.company_name = company_name
        self.ticker = ticker.lower()
        
        # XBRL namespaces
        self.namespaces = {
            'xbrli': 'http://www.xbrl.org/2003/instance',
            'xbrldi': 'http://xbrl.org/2006/xbrldi',
        }
    
    def _get_statement_items(self, statement_type):
        """Get line items in proper order for each statement"""
        if statement_type == 'income':
            return OrderedDict([
                # Sales and Revenues
                ('Revenues_MET', '     Sales of Machinery, Energy & Transportation'),
                ('Revenues_FinancialProducts', '     Revenues of Financial Products'),
                ('Revenues_Total', '     Total sales and revenues'),
                
                # Operating Costs
                ('CostOfRevenue', '     Cost of goods sold'),
                ('SellingGeneralAndAdministrativeExpense', '     SG&A Expenses'),
                ('ResearchAndDevelopmentExpense', '     R&D Expenses'),
                ('FinancingInterestExpense_FinancialProducts', '     Interest expense of Financial Products'),
                ('OtherOperatingIncomeExpenseNet', '     Other operating (income) expenses'),
                ('CostsAndExpenses', '     Total operating costs'),
                
                ('OperatingIncomeLoss', 'Operating Profit'),
                
                ('InterestExpenseNonoperating_EXFP', '     Interest expense excluding Financial Products'),
                ('OtherNonoperatingIncomeExpense', '     Other income (expense)'),
                
                ('IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments', 'Consolidated profit before taxes'),
                
                ('IncomeTaxExpenseBenefit', '     Provision (benefit) for income taxes'),
                ('ProfitOfConsolidatedCompanies', '     Profit of consolidated companies'),
                
                ('IncomeLossFromEquityMethodInvestments', '     Equity in profit (loss) of unconsolidated affiliated companies'),
                
                ('ProfitLoss', 'Profit of consolidated and affiliated companies'),
                
                ('NetIncomeLossAttributableToNoncontrollingInterest', 'Less: Profit (loss) attributable to noncontrolling interests'),
                
                ('NetIncomeLossAvailableToCommonStockholdersBasic', 'Profit (Attributable to Common Stockholders)'),
                
                # EPS
                ('EarningsPerShareBasic', 'Profit per common share'),
                
                ('EarningsPerShareDiluted', 'Profit per common share - diluted'),
                
                # Shares Outstanding
                ('WeightedAverageNumberOfSharesOutstandingBasic', 'Shares Outstanding - Basic'),
                ('WeightedAverageNumberOfDilutedSharesOutstanding', 'Shares Outstanding - Diluted'),
            ])
        
        elif statement_type == 'balance':
            return OrderedDict([
                # Assets
                # Current Assets
                ('CashAndCashEquivalentsAtCarryingValue', '          Cash & Cash Equivalents'),
                ('AccountsReceivableNetCurrent', '          Receivables - trade and other'),
                ('NotesAndLoansReceivableNetCurrent', '          Receivables - finance'),
                ('PrepaidExpenseAndOtherAssetsCurrent', '          Prepaid Expenses And Other Assets Current'),
                ('InventoryNet', '          Inventories'),
                ('AssetsCurrent', '     Total Current Assets'),
                
                ('PropertyPlantAndEquipmentNet', '     Property, Plant, & Equipment - net'),
                ('AccountsReceivableNetNoncurrent', '     Long-term receivables - trade and other'),
                ('NotesAndLoansReceivableNetNoncurrent', '     Long-term receivables - finance'),
                ('NoncurrentDeferredAndRefundableIncomeTaxes', '     Noncurrent deferred and refundable income taxes'),
                ('IntangibleAssetsNetExcludingGoodwill', '     Intangible Assets'),
                ('Goodwill', '     Goodwill'),
                ('OtherAssetsNoncurrent', '     Other assets'),
                ('Assets', 'Total assets'),
                
                # Liabilities
                # Current liabilities
                ('ShortTermBorrowings_FinancialProducts', '               Financial Products'),
                ('AccountsPayableCurrent', '          Accounts payable'),
                ('AccruedLiabilitiesCurrent', '          Accrued expenses'),
                ('EmployeeRelatedLiabilitiesCurrent', '          Accrued wages: salaries and employee benefits'),
                ('ContractWithCustomerLiabilityCurrent', '          Customer advances'),
                ('DividendsPayableCurrent', '          Dividends payable'),
                ('OtherLiabilitiesCurrent', '          Other current liabilities'),
                
                # Long-term debt
                ('LongTermDebtAndCapitalLeaseObligationsCurrent_MET', '               Machinery: Energy & Transportation'),
                ('LongTermDebtAndCapitalLeaseObligationsCurrent_FP', '               Financial Products'),
                ('LiabilitiesCurrent', '     Total current liabilities'),
                
                ('LongTermDebtAndCapitalLeaseObligations_MET', '               Machinery: Energy & Transportation'),
                ('LongTermDebtAndCapitalLeaseObligations_FP', '               Financial Products'),
                ('PensionAndOtherPostretirementAndPostemploymentBenefitPlansLiabilitiesNoncurrent', '     Liability for postemployment benefits'),
                ('OtherLiabilitiesNoncurrent', '     Other liabilities'),
                ('Liabilities', 'Total Liabilities'),
                
                # Shareholders' Equity
                ('CommonStocksIncludingAdditionalPaidInCapital', '     Issued shares at paid-in amount'),
                ('TreasuryStockValue', '     Treasury stock at cost'),
                ('RetainedEarningsAccumulatedDeficit', '     Profit employed in the business'),
                ('AccumulatedOtherComprehensiveIncomeLossNetOfTax', '     Accumulated other comprehensive income (loss)'),
                ('MinorityInterest', '     Noncontrolling interests'),
                ('StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest', "Total shareholders' equity"),
                ('LiabilitiesAndStockholdersEquity', "Total liabilities and shareholders' equity"),
            ])
        
        elif statement_type == 'cashflow':
            return OrderedDict([
                # Operating Activities
                ('ProfitLoss', '     Profit of consolidated and affiliated companies'),
                # Adjustments
                ('DepreciationDepletionAndAmortization', '          Depreciation and Amortization'),
                ('DeferredIncomeTaxExpenseBenefit', '          Provision (benefit) for deferred income taxes'),
                ('NonCashGainLossOnDivestiture', '          (Gain) loss on divestiture'),
                ('OtherNoncashIncomeExpense', '          Other'),
                # Changes in assets and liabilities
                ('IncreaseDecreaseInReceivables', '          Receivables – trade and other'),
                ('IncreaseDecreaseInInventories', '          Inventories'),
                ('IncreaseDecreaseInAccountsPayable', '          Accounts payable'),
                ('IncreaseDecreaseInAccruedLiabilities', '           Accrued expenses'),
                ('IncreaseDecreaseInEmployeeRelatedLiabilities', '          Accrued wages, salaries and employee benefits'),
                ('IncreaseDecreaseInContractWithCustomerLiability', '          Customer advances'),
                ('IncreaseDecreaseInOtherOperatingAssets', '          Other assets - net'),
                ('IncreaseDecreaseInOtherOperatingLiabilities', '          Other liabilities - net'),
                ('NetCashProvidedByUsedInOperatingActivities', 'Net cash provided by (used for) operating activities'),
                
                # Investing Activities
                ('PaymentsToAcquirePropertyPlantAndEquipment', '     Capital expenditures – excluding equipment leased to others'),
                ('PaymentsToAcquireEquipmentOnLease', '     Expenditures for equipment leased to others'),
                ('ProceedsFromSaleOfPropertyPlantAndEquipment', '     Proceeds from disposals of leased assets and property, plant and equipment'),
                ('PaymentsToAcquireFinanceReceivables', '     Additions to finance receivables'),
                ('ProceedsFromCollectionOfFinanceReceivables', '     Collections of finance receivables'),
                ('ProceedsFromSaleOfFinanceReceivables', '     Proceeds from sale of finance receivables'),
                ('PaymentsToAcquireBusinessesNetOfCashAcquired', '     Investments and acquisitions (net of cash acquired)'),
                ('ProceedsFromDivestitureOfBusinessesNetOfCashDivested', '     Proceeds from sale of businesses and investments (net of cash sold)'),
                ('ProceedsFromSaleAndMaturityOfMarketableSecurities', '      Proceeds from maturities and sale of securities'),
                ('PaymentsToAcquireMarketableSecurities', '      Investments in securities'),
                ('PaymentsForProceedsFromOtherInvestingActivities', '     Other – net'),
                ('NetCashProvidedByUsedInInvestingActivities', 'Net cash provided by (used for) investing activities'),
                
                # Financing Activities
                ('PaymentsOfDividendsCommonStock', '     Dividends paid'),
                ('ProceedsFromIssuanceOrSaleOfEquity', '     Common stock issued, and other stock compensation transactions, net'),
                ('PaymentsForRepurchaseOfCommonStock', '     Payments to purchase common stock'),
                ('PaymentsForExciseTaxOnPurchaseOfCommonStock', '     Excise tax paid on purchases of common stock'),
                ('ProceedsFromDebtMaturingInMoreThanThreeMonths_MET', '          - Machinery, Energy & Transportation'),
                ('ProceedsFromDebtMaturingInMoreThanThreeMonths_FP', '          - Financial Products'),
                ('RepaymentsOfDebtMaturingInMoreThanThreeMonths_MET', '          - Machinery, Energy & Transportation'),
                ('RepaymentsOfDebtMaturingInMoreThanThreeMonths_FP', '          - Financial Products'),
                ('ProceedsFromRepaymentsOfShortTermDebtMaturingInThreeMonthsOrLess', '     Short-term borrowings – net (original maturities three months or less)'),
                ('NetCashProvidedByUsedInFinancingActivities', 'Net cash provided by (used for) financing activities'),
                ('EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents', 'Effect of exchange rate changes on cash'),
                ('CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect', 'Increase (decrease) in cash, cash equivalents and restricted cash'),
                ('CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents_Beginning', 'Cash, cash equivalents and restricted cash at beginning of period'),
                ('CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents_End', 'Cash, cash equivalents and restricted cash at end of period'),
            ])
        
        else:
            logger.warning(f"Unknown statement type: {statement_type}")
            return OrderedDict()
    
    def get_all_filings(self, start_year=2010):
        """Get all filings from start_year to present"""
        url = f"{self.base_url}/submissions/CIK{self.cik}.json"
        
        logger.info(f"Fetching all filings since {start_year} for {self.company_name}")
        
        try:
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            time.sleep(0.1)
            
            data = response.json()
            recent = data['filings']['recent']
            
            filings = []
            for i in range(len(recent['form'])):
                filing_date = recent['filingDate'][i]
                filing_year = int(filing_date.split('-')[0])
                
                if filing_year >= start_year:
                    form = recent['form'][i]
                    if form in ['10-Q', '10-K']:
                        filings.append({
                            'accession': recent['accessionNumber'][i],
                            'filing_date': filing_date,
                            'report_date': recent['reportDate'][i],
                            'form': form,
                            'primary_document': recent['primaryDocument'][i]
                        })
            
            filings.sort(key=lambda x: x['report_date'])
            
            logger.info(f"Found {len(filings)} filings since {start_year}")
            return filings
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error fetching submissions: {e}")
            raise
    
    def construct_instance_url(self, accession, report_date):
        """Construct URL for XBRL instance document"""
        accession_no_dash = accession.replace('-', '')
        date_obj = datetime.strptime(report_date, '%Y-%m-%d')
        date_str = date_obj.strftime('%Y%m%d')
        url = f"{self.sec_archives}/{self.cik_int}/{accession_no_dash}/{self.ticker}-{date_str}_htm.xml"
        return url
    
    def download_instance_document(self, url):
        """Download XBRL instance document"""
        try:
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            time.sleep(0.15)
            return response.content
        except requests.exceptions.RequestException as e:
            logger.error(f"Error downloading {url}: {e}")
            raise
    
    def parse_context_elements(self, root):
        """Parse context elements to understand segments"""
        contexts = {}
        
        for context in root.findall('.//xbrli:context', self.namespaces):
            context_id = context.get('id')
            
            period = context.find('xbrli:period', self.namespaces)
            instant = period.find('xbrli:instant', self.namespaces)
            start = period.find('xbrli:startDate', self.namespaces)
            end = period.find('xbrli:endDate', self.namespaces)
            
            context_info = {
                'id': context_id,
                'segments': {},
                'instant': instant.text if instant is not None else None,
                'start': start.text if start is not None else None,
                'end': end.text if end is not None else None,
            }
            
            entity = context.find('xbrli:entity', self.namespaces)
            if entity is not None:
                segment = entity.find('xbrli:segment', self.namespaces)
                if segment is not None:
                    for member in segment.findall('.//xbrldi:explicitMember', self.namespaces):
                        dimension = member.get('dimension')
                        member_value = member.text
                        
                        if ':' in member_value:
                            member_value = member_value.split(':')[1]
                        
                        context_info['segments'][dimension] = member_value
            
            contexts[context_id] = context_info
        
        return contexts
    
    def extract_facts_from_xbrl(self, xml_content):
        """Extract all facts from XBRL instance"""
        root = ET.fromstring(xml_content)
        
        # Update namespaces from document
        for prefix, uri in root.attrib.items():
            if prefix.startswith('{http://www.w3.org/2000/xmlns/}'):
                ns_prefix = prefix.split('}')[1]
                self.namespaces[ns_prefix] = uri
        
        contexts = self.parse_context_elements(root)
        
        facts = []
        
        for elem in root.iter():
            tag_name = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            context_ref = elem.get('contextRef')
            unit_ref = elem.get('unitRef')
            decimals = elem.get('decimals')
            
            if context_ref in contexts and elem.text:
                context = contexts[context_ref]
                
                segment_name = "Consolidated"
                segment_dimension = None
                
                if context['segments']:
                    for dim, member in context['segments'].items():
                        segment_name = member
                        segment_dimension = dim
                        break
                
                try:
                    value = float(elem.text)
                except (ValueError, TypeError):
                    continue
                
                fact = {
                    'tag': tag_name,
                    'value': value,
                    'context_id': context_ref,
                    'segment': segment_name,
                    'dimension': segment_dimension,
                    'start_date': context['start'],
                    'end_date': context['end'],
                    'instant_date': context['instant'],
                    'decimals': decimals,
                    'unit': unit_ref
                }
                
                facts.append(fact)
        
        return facts
    
    def process_filing(self, filing):
        """Process a single filing"""
        try:
            url = self.construct_instance_url(filing['accession'], filing['report_date'])
            logger.info(f"Processing {filing['form']} from {filing['report_date']}")
            
            xml_content = self.download_instance_document(url)
            facts = self.extract_facts_from_xbrl(xml_content)
            
            for fact in facts:
                fact['accession'] = filing['accession']
                fact['filing_date'] = filing['filing_date']
                fact['report_date'] = filing['report_date']
                fact['form'] = filing['form']
            
            logger.info(f"  Extracted {len(facts)} facts")
            return facts
            
        except Exception as e:
            logger.error(f"Error processing filing: {e}")
            return []
    
    def extract_all_data(self, start_year=2010):
        """Extract all financial data"""
        filings = self.get_all_filings(start_year=start_year)
        
        all_facts = []
        for i, filing in enumerate(filings, 1):
            logger.info(f"\n[{i}/{len(filings)}] " + "="*50)
            facts = self.process_filing(filing)
            all_facts.extend(facts)
        
        df = pd.DataFrame(all_facts)
        
        if not df.empty:
            for date_col in ['start_date', 'end_date', 'instant_date', 'filing_date', 'report_date']:
                if date_col in df.columns:
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            
            sort_col = 'end_date' if 'end_date' in df.columns else 'instant_date'
            df = df.sort_values([sort_col, 'tag', 'segment'], ascending=[True, True, True])
        
        logger.info(f"\nTotal facts extracted: {len(df)}")
        return df
    
    def calculate_q4_data(self, df):
        """Calculate Q4 by subtracting Q1+Q2+Q3 from annual"""
        logger.info("\n" + "="*60)
        logger.info("Calculating Q4 data")
        logger.info("="*60)
        
        if df.empty:
            return df
        
        quarterly_df = df[df['form'] == '10-Q'].copy()
        annual_df = df[df['form'] == '10-K'].copy()
        
        q4_records = []
        
        for tag in df['tag'].unique():
            for segment in df['segment'].unique():
                annual_subset = annual_df[
                    (annual_df['tag'] == tag) & 
                    (annual_df['segment'] == segment)
                ].copy()
                
                for _, annual_row in annual_subset.iterrows():
                    fiscal_year_end = annual_row['end_date']
                    
                    if pd.isna(fiscal_year_end):
                        continue
                    
                    fiscal_year = fiscal_year_end.year
                    
                    quarterly_subset = quarterly_df[
                        (quarterly_df['tag'] == tag) &
                        (quarterly_df['segment'] == segment) &
                        (quarterly_df['end_date'] > pd.Timestamp(year=fiscal_year-1, month=12, day=31)) &
                        (quarterly_df['end_date'] <= fiscal_year_end)
                    ]
                    
                    if len(quarterly_subset) == 3:
                        q1q2q3_total = quarterly_subset['value'].sum()
                        annual_total = annual_row['value']
                        q4_value = annual_total - q1q2q3_total
                        
                        q3_end = quarterly_subset['end_date'].max()
                        
                        q4_record = {
                            'tag': tag,
                            'value': q4_value,
                            'segment': segment,
                            'start_date': q3_end + pd.Timedelta(days=1),
                            'end_date': fiscal_year_end,
                            'instant_date': None,
                            'report_date': fiscal_year_end,
                            'form': '10-Q (Q4 Calculated)',
                            'accession': annual_row['accession'],
                            'filing_date': annual_row['filing_date'],
                            'context_id': f"Q4_{fiscal_year}_{segment}",
                            'dimension': annual_row.get('dimension'),
                            'decimals': annual_row.get('decimals'),
                            'unit': annual_row.get('unit')
                        }
                        
                        q4_records.append(q4_record)
        
        if q4_records:
            q4_df = pd.DataFrame(q4_records)
            combined_df = pd.concat([df, q4_df], ignore_index=True)
            combined_df = combined_df.sort_values(['end_date', 'tag', 'segment'])
            logger.info(f"Added {len(q4_records)} Q4 records")
            return combined_df
        
        return df
    
    def create_statement_pivot(self, df, statement_type):
        """Create pivot table for a financial statement in proper order"""
        if df.empty:
            return pd.DataFrame()
        
        statement_items = self._get_statement_items(statement_type)
        
        # Determine date column
        if statement_type == 'balance':
            date_col = 'instant_date'
            df_filtered = df[df['instant_date'].notna()].copy()
        else:
            date_col = 'end_date'
            df_filtered = df[df['end_date'].notna()].copy()
        
        if df_filtered.empty:
            return pd.DataFrame()
        
        # Filter to quarterly data only
        df_filtered = df_filtered[df_filtered['form'].str.contains('10-Q', na=False)]
        
        # Build pivot data in proper order
        pivot_data = []
        
        for tag_key, label in statement_items.items():
            if tag_key == '':
                # Blank row
                pivot_data.append({'Line_Item': label})
                continue
            
            # Extract base tag (remove segment suffix)
            base_tag = tag_key.split('_')[0]
            
            # Check if this line item needs segment breakdown
            if '_' in tag_key:
                # This is a segment-specific line
                segment_suffix = tag_key.split('_', 1)[1]
                
                # Map suffix to actual segment name
                segment_map = {
                    'FinancialProducts': 'FinancialProductsMember',
                    'FP': 'FinancialProductsMember',
                    'MET': 'MachineryEnergyTransportationMember',
                    'EXFP': 'AllOtherExcludingFinancialProductsMember',
                    'Total': 'Consolidated'
                }
                
                target_segment = segment_map.get(segment_suffix, segment_suffix)
                
                subset = df_filtered[
                    (df_filtered['tag'] == base_tag) &
                    (df_filtered['segment'].str.contains(target_segment, case=False, na=False))
                ]
            else:
                # Get consolidated data (no segment or "Consolidated" segment)
                subset = df_filtered[
                    (df_filtered['tag'] == tag_key) &
                    ((df_filtered['segment'] == 'Consolidated') | (df_filtered['segment'].isna()))
                ]
            
            if not subset.empty:
                pivot_row = subset.pivot_table(
                    index='tag',
                    columns=date_col,
                    values='value',
                    aggfunc='first'
                )
                
                if not pivot_row.empty:
                    row_dict = {'Line_Item': label}
                    row_dict.update(pivot_row.iloc[0].to_dict())
                    pivot_data.append(row_dict)
        
        if not pivot_data:
            return pd.DataFrame()
        
        result_df = pd.DataFrame(pivot_data)
        
        # Sort columns by date (most recent first), keeping Line_Item first
        date_columns = [col for col in result_df.columns if col != 'Line_Item']
        date_columns_sorted = sorted(date_columns, reverse=True)
        result_df = result_df[['Line_Item'] + date_columns_sorted]
        
        # Format column names
        for col in date_columns_sorted:
            if isinstance(result_df[col].dtype, pd.DatetimeTZDtype) or pd.api.types.is_datetime64_any_dtype(result_df[col]):
                continue
        
        result_df.columns = [
            col.strftime('%Y-%m-%d') if isinstance(col, pd.Timestamp) else col
            for col in result_df.columns
        ]
        
        return result_df
    
    def format_excel_sheet(self, writer, sheet_name, df):
        """Apply formatting to Excel sheet"""
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        header_format = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        
        for cell in next(worksheet.iter_rows(min_row=1, max_row=1)):
            cell.fill = header_format
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    text = str(cell.value) if cell.value is not None else ""
                    if len(text) > max_length:
                        max_length = len(text)
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        accounting_format = '#,##0'
        for row in worksheet.iter_rows(min_row=2):
            for cell in row[1:]:
                if cell.value is not None and isinstance(cell.value, (int, float)):
                    cell.number_format = accounting_format
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        
        worksheet.freeze_panes = 'B2'
    
    def export_to_excel(self, output_filename, start_year=2010):
        """Extract and export all financial data"""
        logger.info("="*60)
        logger.info(f"Starting comprehensive extraction for {self.company_name}")
        logger.info(f"Data range: {start_year} - Present")
        logger.info("="*60)
        
        df = self.extract_all_data(start_year)
        
        if df.empty:
            logger.warning("No data extracted!")
            return None
        
        df = self.calculate_q4_data(df)
        
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            
            # Raw data
            df.to_excel(writer, sheet_name='All Data - Raw', index=False)
            self.format_excel_sheet(writer, 'All Data - Raw', df)
            
            # Income Statement
            logger.info("\nCreating Income Statement")
            income_pivot = self.create_statement_pivot(df, 'income')
            if not income_pivot.empty:
                income_pivot.to_excel(writer, sheet_name='Income Statement - Quarterly', index=False)
                self.format_excel_sheet(writer, 'Income Statement - Quarterly', income_pivot)
            
            # Balance Sheet
            logger.info("Creating Balance Sheet")
            balance_pivot = self.create_statement_pivot(df, 'balance')
            if not balance_pivot.empty:
                balance_pivot.to_excel(writer, sheet_name='Balance Sheet - Quarterly', index=False)
                self.format_excel_sheet(writer, 'Balance Sheet - Quarterly', balance_pivot)
            
            # Cash Flow
            logger.info("Creating Cash Flow Statement")
            cashflow_pivot = self.create_statement_pivot(df, 'cashflow')
            if not cashflow_pivot.empty:
                cashflow_pivot.to_excel(writer, sheet_name='Cash Flow - Quarterly', index=False)
                self.format_excel_sheet(writer, 'Cash Flow - Quarterly', cashflow_pivot)
        
        logger.info("\n" + "="*60)
        logger.info(f"✓ Export complete! File saved: {output_filename}")
        logger.info("="*60)
        
        return output_filename


def main():
    """Main execution"""
    
    YOUR_EMAIL = "brayden.joyce@doosan.com"
    CIK = "0000018230"
    COMPANY_NAME = "Caterpillar Inc."
    TICKER = "cat"
    START_YEAR = 2010
    
    extractor = ComprehensiveXBRLExtractor(
        email=YOUR_EMAIL,
        cik=CIK,
        company_name=COMPANY_NAME,
        ticker=TICKER
    )
    
    output_file = extractor.export_to_excel(
        output_filename='caterpillar_financials.xlsx',
        start_year=START_YEAR
    )
    
    if output_file:
        print(f"\nSuccess! Complete financial data exported to: {output_file}")
        print(f"\nData range: {START_YEAR} - Present")
        print("\nThe Excel file contains:")
        print("  Income Statement - Quarterly")
        print("  Balance Sheet - Quarterly")
        print("  Cash Flow Statement - Quarterly")
        print("  All raw data")


if __name__ == "__main__":
    main()
