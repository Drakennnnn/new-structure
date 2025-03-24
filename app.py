import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import base64
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import tempfile
import zipfile
from docxtpl import DocxTemplate

# Set page configuration
st.set_page_config(
    page_title="Real Estate Cost Sheet Generator",
    page_icon="🏙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Define CSS styles
st.markdown("""
<style>
    .main-header {
        font-size: 30px;
        font-weight: bold;
        color: #1E3A8A;
        margin-bottom: 20px;
    }
    .section-header {
        font-size: 24px;
        font-weight: bold;
        color: #2563EB;
        margin-top: 30px;
        margin-bottom: 15px;
        padding-bottom: 5px;
        border-bottom: 2px solid #BFDBFE;
    }
    .info-text {
        font-size: 16px;
        color: #1F2937;
        margin-bottom: 20px;
    }
    .success-box {
        padding: 15px;
        border-radius: 5px;
        background-color: #ECFDF5;
        border-left: 5px solid #10B981;
        margin-bottom: 20px;
    }
    .warning-box {
        padding: 15px;
        border-radius: 5px;
        background-color: #FFFBEB;
        border-left: 5px solid #F59E0B;
        margin-bottom: 20px;
    }
    .error-box {
        padding: 15px;
        border-radius: 5px;
        background-color: #FEF2F2;
        border-left: 5px solid #EF4444;
        margin-bottom: 20px;
    }
    .status-badge {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 14px;
        font-weight: bold;
    }
    .status-verified {
        background-color: #D1FAE5;
        color: #047857;
    }
    .status-warning {
        background-color: #FEF3C7;
        color: #B45309;
    }
    .status-error {
        background-color: #FEE2E2;
        color: #B91C1C;
    }
    .footer {
        margin-top: 50px;
        text-align: center;
        color: #6B7280;
        font-size: 14px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #F3F4F6;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #DBEAFE;
        border-radius: 4px 4px 0px 0px;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">Real Estate Cost Sheet Generator</div>', unsafe_allow_html=True)

# Initialize session state for storing data
if 'sales_master_df' not in st.session_state:
    st.session_state.sales_master_df = None
if 'collection_df' not in st.session_state:
    st.session_state.collection_df = None
if 'accounts_info' not in st.session_state:
    st.session_state.accounts_info = None
if 'verification_results' not in st.session_state:
    st.session_state.verification_results = {}
if 'selected_customers' not in st.session_state:
    st.session_state.selected_customers = []
if 'noc_template' not in st.session_state:
    st.session_state.noc_template = None
if 'preview_data' not in st.session_state:
    st.session_state.preview_data = {}

# Helper functions
def extract_excel_date(excel_date):
    """Convert Excel date number to Python datetime"""
    try:
        # Excel date starts from 1900-01-01
        if isinstance(excel_date, (int, float)):
            return pd.to_datetime('1899-12-30') + pd.Timedelta(days=int(excel_date))
        elif isinstance(excel_date, str):
            return pd.to_datetime(excel_date, errors='coerce')
        else:
            return excel_date
    except:
        return None

def identify_sales_master_sheet(workbook):
    """Return the Annex - Sales Master sheet"""
    # Check if the exact sheet name exists
    if "Annex - Sales Master" in workbook.sheetnames:
        return "Annex - Sales Master"
    
    # Fallback: try to find a sheet with a similar name
    for sheet_name in workbook.sheetnames:
        if "Annex" in sheet_name and "Sales" in sheet_name:
            return sheet_name
    
    # If still not found, return first sheet as last resort
    return workbook.sheetnames[0] if workbook.sheetnames else None

def identify_collection_sheet(workbook):
    """Return the Main Collection AC P1_P2_P3 sheet"""
    # Check if the exact sheet name exists
    if "Main Collection AC P1_P2_P3" in workbook.sheetnames:
        return "Main Collection AC P1_P2_P3"
    
    # Fallback: try to find a sheet with a similar name
    for sheet_name in workbook.sheetnames:
        if "Main Collection" in sheet_name:
            return sheet_name
    
    # If still not found, return None
    return None

def parse_sales_master(sheet):
    """Parse the Annex - Sales Master sheet and return a DataFrame"""
    # Find header row - look for key column names
    header_row = None
    header_cols = {}
    
    for r_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=True)):
        col_matches = 0
        for c_idx, cell in enumerate(row):
            if cell and isinstance(cell, str):
                cell_lower = cell.lower()
                if any(key in cell_lower for key in ['sr no', 'customer', 'unit', 'tower', 'booking']):
                    col_matches += 1
        
        if col_matches >= 3:  # Found at least 3 key columns
            header_row = r_idx + 1
            break
    
    if header_row is None:
        return None
    
    # Map column indices to expected column names
    header_values = list(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    
    # Map common column variations to standardized names
    column_mapping = {}
    for idx, header in enumerate(header_values):
        if header:
            header_str = str(header).lower()
            
            # Map variations to standard names
            if any(term in header_str for term in ['sr', 'serial']):
                column_mapping[idx] = 'Sr No'
            elif 'customer' in header_str or 'name' in header_str:
                column_mapping[idx] = 'Name of Customer'
            elif 'unit' in header_str and 'number' in header_str:
                column_mapping[idx] = 'Unit Number'
            elif 'tower' in header_str:
                column_mapping[idx] = 'Tower No'
            elif 'booking' in header_str and 'date' in header_str:
                column_mapping[idx] = 'Booking date'
            elif 'status' in header_str:
                column_mapping[idx] = 'Booking Status'
            elif 'payment' in header_str and 'plan' in header_str:
                column_mapping[idx] = 'Payment Plan'
            elif 'area' in header_str and not 'carpet' in header_str:
                column_mapping[idx] = 'Area(sqft)'
            elif 'carpet' in header_str:
                column_mapping[idx] = 'Carpet Area(sqft)'
            elif 'bsp' in header_str and 'sqft' in header_str:
                column_mapping[idx] = 'BSP/SqFt'
            elif 'basic' in header_str and 'price' in header_str:
                column_mapping[idx] = 'Basic Price ( Exl Taxes)'
            elif 'received' in header_str and ('amount' in header_str or 'amt' in header_str) and not 'tax' in header_str:
                column_mapping[idx] = 'Amount received ( Exl Taxes)'
            elif 'tax' in header_str and 'received' in header_str:
                column_mapping[idx] = 'Taxes Received'
            elif 'received' in header_str and ('inc' in header_str or 'with' in header_str) and 'tax' in header_str:
                column_mapping[idx] = 'Amount received (Inc Taxes)'
            elif 'balance' in header_str:
                column_mapping[idx] = 'Balance receivables (Total Sale Consideration )'
            elif 'broker' in header_str and 'name' in header_str:
                column_mapping[idx] = 'Broker Name'
            else:
                # Keep original header for other columns
                column_mapping[idx] = header
    
    # Read data into a DataFrame
    data = []
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        # Skip completely empty rows
        if all(cell is None or cell == '' for cell in row):
            continue
        
        row_data = {}
        for idx, cell in enumerate(row):
            if idx in column_mapping:
                row_data[column_mapping[idx]] = cell
        
        # Only add rows with unit number or customer name
        if ('Unit Number' in row_data and row_data['Unit Number']) or \
           ('Name of Customer' in row_data and row_data['Name of Customer']):
            data.append(row_data)
    
    df = pd.DataFrame(data)
    
    # Process date columns
    date_columns = [col for col in df.columns if 'date' in col.lower()]
    for col in date_columns:
        df[col] = df[col].apply(extract_excel_date)
    
    return df

def identify_account_sections(sheet):
    """Identify different account sections in the collection sheet"""
    accounts = []
    current_account = None
    
    # Get sheet dimensions
    range_ref = sheet['!ref']
    if not range_ref:
        return accounts
    
    range_data = XLSX.utils.decode_range(range_ref)
    max_row = range_data.e.r
    
    # First, find the header row in the first section
    header_row = -1
    header_indices = {}
    
    # Look for the header row with Txn Date, Description, etc.
    for r in range(min(10, max_row)):
        row_data = []
        for c in range(15):  # Check the first 15 columns
            cell = sheet[XLSX.utils.encode_cell({r: r, c: c})]
            if cell:
                row_data.append(cell.v)
            else:
                row_data.append(None)
        
        # Convert to string for easier checking
        row_text = ' '.join([str(val) for val in row_data if val is not None])
        
        # Check if this looks like the header row
        if 'Txn Date' in row_text and 'Dr/Cr' in row_text and 'Amount' in row_text:
            header_row = r
            
            # Map headers to column indices
            for c, val in enumerate(row_data):
                if val:
                    val_str = str(val).lower()
                    if 'txn date' in val_str:
                        header_indices['date'] = c
                    elif 'description' in val_str:
                        header_indices['description'] = c
                    elif 'amount' in val_str and not 'running' in val_str:
                        header_indices['amount'] = c
                    elif ('dr' in val_str and 'cr' in val_str) or val_str == 'dr/cr':
                        header_indices['type'] = c
                    elif 'sales' in val_str and 'tag' in val_str:
                        header_indices['sales_tag'] = c
                    elif 'running' in val_str and 'total' in val_str:
                        header_indices['running_total'] = c
            
            break
    
    # Now find all phase transitions
    for r in range(max_row + 1):
        # Look for rows with "Main Collection Escrow A/c Phase-X"
        phase_name = None
        account_number = None
        
        for c in range(5):  # Check first 5 columns
            cell = sheet[XLSX.utils.encode_cell({r: r, c: c})]
            if cell and cell.v:
                # Check for phase name
                if isinstance(cell.v, str) and "Main Collection Escrow A/c Phase-" in cell.v:
                    phase_name = cell.v
                # Check for account number (long numeric value)
                elif (isinstance(cell.v, (int, float)) and 
                      len(str(int(cell.v))) > 10):
                    account_number = str(int(cell.v))
        
        if phase_name and account_number:
            # Close previous account if any
            if current_account:
                current_account['end_row'] = r - 1
                accounts.append(current_account)
            
            # Start new account section
            current_account = {
                'name': phase_name,
                'number': account_number,
                'start_row': r + 1,  # Start after the header row
                'end_row': None,
                'header_indices': header_indices,
                'header_row': header_row
            }
    
    # Add the last account
    if current_account:
        current_account['end_row'] = max_row
        accounts.append(current_account)
    
    return accounts

def parse_collection_transactions(sheet, accounts_info):
    """Parse transactions from collection sheet based on identified account sections"""
    all_transactions = []
    
    for account in accounts_info:
        # Get header indices for this account
        header_indices = account['header_indices']
        
        # Read transactions for this account
        for r_idx in range(account['start_row'], account.get('end_row', sheet.max_row) + 1):
            # Skip the account header or column header rows
            if r_idx <= account['header_row']:
                continue
                
            # Get row data
            row_data = {}
            row_has_data = False
            
            for c in range(20):  # Check first 20 columns
                cell = sheet[XLSX.utils.encode_cell({r: r_idx, c: c})]
                if cell and cell.v !== undefined:
                    row_data[c] = cell.v
                    row_has_data = true
            
            # Skip empty rows
            if !row_has_data:
                continue
            
            transaction = {
                'account_name': account['name'],
                'account_number': account['number'],
                'row': r_idx
            }
            
            // Extract data based on header indices
            for (const [field, idx] of Object.entries(header_indices)) {
                if (row_data[idx] !== undefined) {
                    // Special handling for sales_tag and type fields
                    if (field === 'sales_tag') {
                        transaction[field] = row_data[idx] !== null ? String(row_data[idx]) : null;
                    } 
                    else if (field === 'type') {
                        transaction[field] = row_data[idx] !== null ? String(row_data[idx]) : null;
                    }
                    else {
                        transaction[field] = row_data[idx];
                    }
                }
            }
            
            // Only include rows with amount and date (valid transactions)
            if (transaction.amount !== undefined && transaction.amount !== null && 
                transaction.date !== undefined && transaction.date !== null) {
                // Convert Excel date to Python datetime
                transaction.date = extract_excel_date(transaction.date);
                all_transactions.push(transaction);
            }
        }
    }
    
    return pd.DataFrame(all_transactions);
}

def verify_transactions(sales_master_df, collection_df):
    """Verify transactions against customer data"""
    verification_results = {}
    
    # Process each customer
    for _, customer in sales_master_df.iterrows():
        unit_number = customer.get('Unit Number')
        customer_name = customer.get('Name of Customer')
        
        if not unit_number:
            continue
        
        # Format unit number consistently
        if isinstance(unit_number, str):
            unit_number = unit_number.strip().upper()
        else:
            unit_number = str(unit_number).strip().upper()
        
        # Find related transactions for this unit
        try:
            # Check if sales_tag column exists
            if 'sales_tag' in collection_df.columns:
                # Normalize unit number for matching: remove spaces and dashes
                search_unit = unit_number.replace(' ', '').replace('-', '').upper()
                
                # Try different matching patterns to handle variations
                patterns = [
                    search_unit,  # Exact match 
                    f"CA{search_unit.replace('CA', '')}"  # Handle CA prefix variations
                ]
                
                # Create a mask for matching any pattern
                mask = collection_df['sales_tag'].notna()
                for pattern in patterns:
                    mask = mask & collection_df['sales_tag'].astype(str).str.upper().str.replace(' ', '').str.replace('-', '').str.contains(pattern, regex=False)
                
                unit_transactions = collection_df[mask]
                
                # If nothing found, try looser matching
                if unit_transactions.empty:
                    # Just match the numeric part (like 402 from CA 04-402)
                    if '-' in unit_number:
                        unit_num = unit_number.split('-')[-1]
                        unit_transactions = collection_df[
                            collection_df['sales_tag'].notna() &
                            collection_df['sales_tag'].astype(str).str.contains(unit_num, regex=False)
                        ]
            else:
                unit_transactions = pd.DataFrame()  # Empty DataFrame if no sales_tag column
        except Exception as e:
            print(f"Error matching transactions for unit {unit_number}: {str(e)}")
            unit_transactions = pd.DataFrame()
        
        # Calculate expected amounts
        expected_amount = customer.get('Amount received (Inc Taxes)', 0)
        if pd.isna(expected_amount):
            expected_amount = 0
        
        expected_base_amount = customer.get('Amount received ( Exl Taxes)', 0) 
        if pd.isna(expected_base_amount):
            expected_base_amount = 0
        
        expected_tax_amount = customer.get('Taxes Received', 0)
        if pd.isna(expected_tax_amount):
            expected_tax_amount = 0
        
        # Calculate actual received from transactions
        try:
            if 'type' in unit_transactions.columns and not unit_transactions.empty:
                credit_transactions = unit_transactions[
                    unit_transactions['type'].notna() & 
                    unit_transactions['type'].astype(str).str.upper().str.contains('C', regex=False)
                ]
                
                debit_transactions = unit_transactions[
                    unit_transactions['type'].notna() & 
                    unit_transactions['type'].astype(str).str.upper().str.contains('D', regex=False)
                ]
            else:
                # If no type column, assume all transactions are credits
                credit_transactions = unit_transactions
                debit_transactions = pd.DataFrame()
        except Exception as e:
            print(f"Error processing transaction types for unit {unit_number}: {str(e)}")
            credit_transactions = pd.DataFrame()
            debit_transactions = pd.DataFrame()
        
        total_credits = credit_transactions['amount'].sum() if not credit_transactions.empty else 0
        total_debits = debit_transactions['amount'].sum() if not debit_transactions.empty else 0
        
        # Calculate net amount (credits - debits)
        actual_amount = total_credits - total_debits
        
        # Check for bounced transactions
        bounced_transactions = []
        for _, cr_txn in credit_transactions.iterrows():
            if 'date' in cr_txn and 'amount' in cr_txn:
                cr_date = cr_txn['date']
                cr_amount = cr_txn['amount']
                
                # Look for debits with same amount within 7 days (potential bounce)
                try:
                    if not debit_transactions.empty and 'date' in debit_transactions.columns:
                        potential_bounces = debit_transactions[
                            (debit_transactions['date'] >= cr_date) & 
                            (debit_transactions['date'] <= cr_date + pd.Timedelta(days=7)) &
                            (debit_transactions['amount'] == cr_amount)
                        ]
                        
                        if not potential_bounces.empty:
                            for _, bounce in potential_bounces.iterrows():
                                bounced_transactions.append({
                                    'credit_date': cr_date,
                                    'credit_amount': cr_amount,
                                    'debit_date': bounce.get('date'),
                                    'debit_amount': bounce.get('amount'),
                                    'description': bounce.get('description', '')
                                })
                except Exception as e:
                    print(f"Error checking for bounced transactions: {str(e)}")
        
        # Determine verification status
        # Tolerance of 1 rupee for floating point discrepancies
        amount_match = abs(actual_amount - expected_amount) <= 1
        
        status = "verified" if amount_match and not bounced_transactions else "warning"
        
        if not unit_transactions.empty and not amount_match:
            status = "error"
        
        verification_results[unit_number] = {
            'unit_number': unit_number,
            'customer_name': customer_name,
            'expected_amount': expected_amount,
            'expected_base_amount': expected_base_amount,
            'expected_tax_amount': expected_tax_amount,
            'actual_amount': actual_amount,
            'amount_match': amount_match,
            'total_credits': total_credits,
            'total_debits': total_debits,
            'transaction_count': len(unit_transactions),
            'bounced_transactions': bounced_transactions,
            'has_bounced': len(bounced_transactions) > 0,
            'status': status,
            'transactions': unit_transactions.to_dict('records') if not unit_transactions.empty else []
        }
    
    return verification_results

def generate_cost_sheet_data(customer_info, verification_info):
    """Generate data for cost sheet based on customer info and verification results"""
    unit_number = customer_info.get('Unit Number')
    if not unit_number:
        return None
    
    # Format unit number consistently
    if isinstance(unit_number, str):
        unit_number = unit_number.strip().upper()
    else:
        unit_number = str(unit_number).strip().upper()
    
    tower_no = customer_info.get('Tower No', '')
    if tower_no and not isinstance(tower_no, str):
        tower_no = str(tower_no)
    
    # Format as "CA 04-402" format
    formatted_unit = f"{tower_no}-{unit_number.split('-')[-1] if '-' in unit_number else unit_number}"
    formatted_unit = formatted_unit.replace(" ", "-").upper()
    
    # Get values with proper null/None checking
    super_area = customer_info.get('Area(sqft)', 0)
    if super_area is None or pd.isna(super_area):
        super_area = 0
        
    carpet_area = customer_info.get('Carpet Area(sqft)', 0)
    if carpet_area is None or pd.isna(carpet_area):
        carpet_area = 0
        
    bsp_rate = customer_info.get('BSP/SqFt', 0)
    if bsp_rate is None or pd.isna(bsp_rate):
        bsp_rate = 0
        
    bsp_amount = customer_info.get('Basic Price ( Exl Taxes)', 0)
    if bsp_amount is None or pd.isna(bsp_amount):
        bsp_amount = 0
    
    # Basic customer and unit info
    cost_sheet_data = {
        'unit_number': unit_number,
        'formatted_unit': formatted_unit,
        'tower': tower_no,
        'customer_name': customer_info.get('Name of Customer', ''),
        'booking_date': customer_info.get('Booking date'),
        'payment_plan': customer_info.get('Payment Plan', ''),
        'super_area': super_area,
        'carpet_area': carpet_area,
    }
    
    # Cost calculations
    ifms_rate = 25  # Standard rates based on your examples
    ifms_amount = ifms_rate * super_area
    amc_rate = 3
    amc_amount = amc_rate * super_area * 12  # Annual (12 months)
    gst_rate = 5  # 5% on BSP
    gst_amount = 0.05 * bsp_amount
    amc_gst_rate = 18  # 18% on AMC
    amc_gst_amount = 0.18 * amc_amount
    
    cost_sheet_data.update({
        'bsp_rate': bsp_rate,
        'bsp_amount': bsp_amount,
        'ifms_rate': ifms_rate,
        'ifms_amount': ifms_amount,
        'amc_rate': amc_rate,
        'amc_amount': amc_amount,
        'gst_rate': gst_rate,
        'gst_amount': gst_amount,
        'amc_gst_rate': amc_gst_rate,
        'amc_gst_amount': amc_gst_amount,
    })
    
    # Payment information
    verification = verification_info.get(unit_number, {})
    
    # Handle null/None values
    expected_base_amount = verification.get('expected_base_amount', 0)
    if expected_base_amount is None or pd.isna(expected_base_amount):
        expected_base_amount = 0
        
    expected_tax_amount = verification.get('expected_tax_amount', 0) 
    if expected_tax_amount is None or pd.isna(expected_tax_amount):
        expected_tax_amount = 0
        
    expected_amount = verification.get('expected_amount', 0)
    if expected_amount is None or pd.isna(expected_amount):
        expected_amount = 0
        
    balance_receivable = customer_info.get('Balance receivables (Total Sale Consideration )', 0)
    if balance_receivable is None or pd.isna(balance_receivable):
        balance_receivable = 0
    
    cost_sheet_data.update({
        'amount_received': expected_base_amount,
        'gst_received': expected_tax_amount,
        'total_received': expected_amount,
        'balance_receivable': balance_receivable,
    })
    
    # Bank transaction details
    transactions = verification.get('transactions', [])
    cost_sheet_data['transactions'] = sorted(transactions, key=lambda x: x.get('date', pd.NaT) or pd.NaT)
    
    # Calculate total consideration and balance receivable
    total_consideration = bsp_amount + ifms_amount + amc_amount
    cost_sheet_data['total_consideration'] = total_consideration
    cost_sheet_data['balance_receivable'] = total_consideration - expected_base_amount
    
    # Broker information
    broker_name = customer_info.get('Broker Name', '')
    if not broker_name or pd.isna(broker_name):
        broker_name = 'DIRECT'
        
    broker_rate = 100  # Standard rate from examples
    broker_amount = broker_rate * super_area
    
    cost_sheet_data.update({
        'broker_name': broker_name,
        'broker_rate': broker_rate,
        'broker_amount': broker_amount,
    })
    
    # Calculate floor number from unit number
    floor_number = "Ground"
    try:
        unit_num_part = formatted_unit.split('-')[-1]
        if unit_num_part.isdigit():
            floor = int(unit_num_part) // 100
            floor_number = f"{floor}th" if floor != 0 else "Ground"
    except:
        pass
    
    # Additional info
    home_loan = customer_info.get('Self-funded or loan availed', '')
    if not home_loan or pd.isna(home_loan):
        home_loan = 'N/A'
        
    co_applicant = customer_info.get('CO-APPLICANT NAME', '')
    if not co_applicant or pd.isna(co_applicant):
        co_applicant = 'N/A'
    
    cost_sheet_data.update({
        'floor_number': floor_number,
        'home_loan': home_loan,
        'co_applicant': co_applicant
    })
    
    return cost_sheet_data

def generate_cost_sheet_excel(cost_sheet_data):
    """Generate Excel file for cost sheet"""
    if not cost_sheet_data:
        return None
    
    # Create a new workbook
    wb = openpyxl.Workbook()
    
    # Remove default sheet and create our three sheets
    wb.remove(wb.active)
    data_entry_sheet = wb.create_sheet(f"{cost_sheet_data['formatted_unit']} - Data Entry")
    bank_credit_sheet = wb.create_sheet("Bank Credit Details")
    noc_sheet = wb.create_sheet("Sales NOC SWAMIH")
    
    # ----- Data Entry Sheet -----
    
    # Title row
    data_entry_sheet['A1'] = "COST SHEET"
    data_entry_sheet['A1'].font = Font(bold=True, size=14)
    
    # Header row
    data_entry_sheet['A2'] = "Particulars"
    data_entry_sheet['B2'] = "Details"
    data_entry_sheet['C2'] = "Remarks"
    
    # Apply styling to header
    for col in ['A2', 'B2', 'C2']:
        data_entry_sheet[col].font = Font(bold=True)
        data_entry_sheet[col].alignment = Alignment(horizontal='center')
    
    # Customer info
    data_entry_sheet['A3'] = "DATE OF BOOKING             "
    data_entry_sheet['B3'] = cost_sheet_data['booking_date']
    
    data_entry_sheet['A4'] = "APPLICANT NAME                 "
    data_entry_sheet['B4'] = cost_sheet_data['customer_name']
    
    data_entry_sheet['A5'] = "CO-APPLICANT NAME          "
    data_entry_sheet['B5'] = cost_sheet_data.get('co_applicant', 'N/A')
    
    data_entry_sheet['A6'] = "PAYMENT PLAN"
    data_entry_sheet['B6'] = cost_sheet_data['payment_plan']
    
    data_entry_sheet['A7'] = "TOWER:"
    data_entry_sheet['B7'] = cost_sheet_data['tower']
    
    data_entry_sheet['A8'] = "UNIT NO:"
    data_entry_sheet['B8'] = cost_sheet_data['unit_number'].split('-')[-1] if '-' in cost_sheet_data['unit_number'] else cost_sheet_data['unit_number']
    
    data_entry_sheet['A9'] = "FLOOR NUMBER"
    data_entry_sheet['B9'] = cost_sheet_data['floor_number']
    
    data_entry_sheet['A10'] = "SUPER AREA(SQ. FT.)"
    data_entry_sheet['B10'] = cost_sheet_data['super_area']
    
    data_entry_sheet['A11'] = "CARPET AREA(SQ. FT.)"
    data_entry_sheet['B11'] = cost_sheet_data['carpet_area']
    
    # Cost breakdown header
    data_entry_sheet['A13'] = "Particulars"
    data_entry_sheet['B13'] = "Rate"
    data_entry_sheet['C13'] = "Amount"
    data_entry_sheet['D13'] = "Psft"
    
    for col in ['A13', 'B13', 'C13', 'D13']:
        data_entry_sheet[col].font = Font(bold=True)
        data_entry_sheet[col].alignment = Alignment(horizontal='center')
    
    # Cost breakdown data
    data_entry_sheet['A15'] = "BASIC SALE PRICE            "
    data_entry_sheet['B15'] = cost_sheet_data['bsp_rate']
    data_entry_sheet['C15'] = cost_sheet_data['bsp_amount']
    data_entry_sheet['D15'] = data_entry_sheet['C15'].value / data_entry_sheet['B10'].value
    
    data_entry_sheet['A16'] = "Less: Discount (if any)"
    data_entry_sheet['B16'] = "-"
    data_entry_sheet['C16'] = 0
    
    data_entry_sheet['A17'] = "NET Price"
    data_entry_sheet['C17'] = "=C15-C16"
    
    data_entry_sheet['A18'] = "ADD:-"
    
    data_entry_sheet['A19'] = "IFMS"
    data_entry_sheet['B19'] = cost_sheet_data['ifms_rate']
    data_entry_sheet['C19'] = "=B19*B10"
    data_entry_sheet['D19'] = "=C19/B10"
    
    data_entry_sheet['A20'] = "1 Year Annual Maintenance Charge"
    data_entry_sheet['B20'] = cost_sheet_data['amc_rate']
    data_entry_sheet['C20'] = "=(B20*B10)*12"
    data_entry_sheet['D20'] = "=C20/B10"
    
    data_entry_sheet['A21'] = "Lease Rent"
    data_entry_sheet['B21'] = 0
    data_entry_sheet['C21'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A22'] = "EEC"
    data_entry_sheet['B22'] = 0
    data_entry_sheet['C22'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A23'] = "FFC"
    data_entry_sheet['B23'] = 0
    data_entry_sheet['C23'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A24'] = "IDC"
    data_entry_sheet['B24'] = 0
    data_entry_sheet['C24'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A26'] = "Covered Car Parking       "
    data_entry_sheet['B26'] = 475000
    data_entry_sheet['C26'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A27'] = "Additional Car Parking"
    data_entry_sheet['B27'] = 475000
    data_entry_sheet['C27'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A28'] = "Club Membership           "
    data_entry_sheet['B28'] = 100000
    data_entry_sheet['C28'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A29'] = "Power Backup( 1.K.V.A)                  "
    data_entry_sheet['B29'] = 20000
    data_entry_sheet['C29'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A30'] = "Add.Power Backup( 3.K.V.A)                  "
    data_entry_sheet['B30'] = 60000
    data_entry_sheet['C30'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A31'] = "ADD:               "
    
    data_entry_sheet['A32'] = "PLC'S:-"
    
    data_entry_sheet['A33'] = "PARK"
    data_entry_sheet['B33'] = 150
    data_entry_sheet['C33'] = "INCLUSIVE IN BSP"
    
    data_entry_sheet['A34'] = "CLUB"
    data_entry_sheet['B34'] = 100
    data_entry_sheet['C34'] = "NA"
    
    data_entry_sheet['A35'] = "CORNER"
    data_entry_sheet['B35'] = 100
    data_entry_sheet['C35'] = "NA"
    
    data_entry_sheet['A36'] = "ROAD"
    data_entry_sheet['B36'] = 100
    data_entry_sheet['C36'] = "NA"
    
    # Total considerations
    data_entry_sheet['A38'] = "Total Sale Consideration "
    data_entry_sheet['C38'] = "=SUM(C17:C37)"
    
    data_entry_sheet['A39'] = "GST AMOUNT "
    data_entry_sheet['C39'] = "=(C15)*5%"
    data_entry_sheet['D39'] = "5% GST"
    
    data_entry_sheet['A40'] = "AMC GST"
    data_entry_sheet['C40'] = "=C20*18%"
    data_entry_sheet['D40'] = "18% GST"
    
    data_entry_sheet['A41'] = "Total GST"
    data_entry_sheet['C41'] = "=C40+C39"
    
    data_entry_sheet['A42'] = "GRAND TOTAL"
    data_entry_sheet['C42'] = "=C41+C38"
    
    # Receipt details
    data_entry_sheet['A44'] = "Receipt details"
    
    data_entry_sheet['A45'] = "Amount Received"
    data_entry_sheet['C45'] = cost_sheet_data['amount_received']
    
    data_entry_sheet['A46'] = "BALANCE RECEIVABLE"
    data_entry_sheet['C46'] = "=C38-C45"
    
    data_entry_sheet['A47'] = "GST received"
    data_entry_sheet['C47'] = cost_sheet_data['gst_received']
    
    data_entry_sheet['A48'] = "BALANCE GST RECEIVABLE"
    data_entry_sheet['C48'] = "=C41-C47"
    
    # Broker info
    data_entry_sheet['A50'] = "Brokerage"
    data_entry_sheet['B50'] = cost_sheet_data['broker_rate']
    data_entry_sheet['C50'] = "=B50*B10"
    
    data_entry_sheet['A51'] = "Broker Name "
    data_entry_sheet['C51'] = cost_sheet_data['broker_name']
    
    data_entry_sheet['A52'] = "Amount to be collected Excluding Brokerage"
    data_entry_sheet['C52'] = "=C42-C50"
    
    # Additional information
    data_entry_sheet['A54'] = "Home Loan Taken from"
    data_entry_sheet['C54'] = cost_sheet_data.get('home_loan', 'N/A')
    
    data_entry_sheet['A55'] = "Contact Number of Purchaser"
    data_entry_sheet['C55'] = ""  # Would need to add this to the sales master parsing
    
    data_entry_sheet['A56'] = "Address of Purchaser"
    data_entry_sheet['C56'] = ""  # Would need to add this to the sales master parsing
    
    data_entry_sheet['A58'] = "Sale Status as per DTD"
    data_entry_sheet['C58'] = "Unsold"
    
    data_entry_sheet['A59'] = "Sale Status as per Actual "
    data_entry_sheet['C59'] = "Sold"

    # Adjust column widths
    for col, width in [('A', 40), ('B', 20), ('C', 25), ('D', 15)]:
        data_entry_sheet.column_dimensions[col].width = width
    
    # ----- Bank Credit Details Sheet -----
    bank_credit_sheet['A1'] = "Date"
    bank_credit_sheet['B1'] = "Amount received"
    bank_credit_sheet['C1'] = "Bank A/c Name"
    bank_credit_sheet['D1'] = "Bank A/c No"
    bank_credit_sheet['E1'] = "Bank statement verified"
    
    # Apply styling to header
    for col in ['A1', 'B1', 'C1', 'D1', 'E1']:
        bank_credit_sheet[col].font = Font(bold=True)
        bank_credit_sheet[col].alignment = Alignment(horizontal='center')
    
    # Add transaction data
    transactions = cost_sheet_data.get('transactions', [])
    credit_transactions = [t for t in transactions if t.get('type') == 'C']
    
    for i, txn in enumerate(credit_transactions, start=2):
        bank_credit_sheet[f'A{i}'] = txn.get('date')
        bank_credit_sheet[f'B{i}'] = txn.get('amount')
        bank_credit_sheet[f'C{i}'] = txn.get('account_name')
        bank_credit_sheet[f'D{i}'] = txn.get('account_number')
        bank_credit_sheet[f'E{i}'] = "Yes"
    
    # Add total row
    total_row = len(credit_transactions) + 2
    bank_credit_sheet[f'A{total_row}'] = "Total Sum received"
    bank_credit_sheet[f'B{total_row}'] = f"=SUM(B2:B{total_row-1})"
    
    # Adjust column widths
    for col, width in [('A', 15), ('B', 15), ('C', 35), ('D', 20), ('E', 20)]:
        bank_credit_sheet.column_dimensions[col].width = width
    
    # ----- Sales NOC SWAMIH Sheet -----
    noc_sheet['B1'] = "Particulars"
    noc_sheet['C1'] = "Details"
    
    noc_sheet['B2'] = "Flat/Unit No."
    noc_sheet['C2'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!B8"
    
    noc_sheet['B3'] = "Floor No."
    noc_sheet['C3'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!B9"
    
    noc_sheet['B4'] = "Building Name"
    noc_sheet['C4'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!B7"
    
    noc_sheet['B5'] = "Carpet Area of the Flat / Unit (Sq.Ft.)"
    noc_sheet['C5'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!B11"
    
    noc_sheet['B6'] = "Name of the Applicant (Purchaser)"
    noc_sheet['C6'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!B4"
    
    noc_sheet['B7'] = "Name of the Co-Applicant (Co- Purchaser)"
    noc_sheet['C7'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!B5"
    
    noc_sheet['B8'] = "Total Sale Consideration (including parking charges) (Excl. GST)"
    noc_sheet['C8'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!C38"
    
    noc_sheet['B9'] = "Total GST"
    noc_sheet['C9'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!C41"
    
    noc_sheet['B10'] = "Sale Consideration Amount Received as on date (Excl. GST)"
    noc_sheet['C10'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!C45"
    
    noc_sheet['B11'] = "GST Amount received as on date"
    noc_sheet['C11'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!C47"
    
    noc_sheet['B12'] = "Balance Amount yet to be received (excluding taxes)"
    noc_sheet['C12'] = "=C8-C10"
    
    noc_sheet['B13'] = "Booking Date"
    noc_sheet['C13'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!B3"
    
    noc_sheet['B14'] = "Home loan taken from"
    noc_sheet['C14'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!C54"
    
    noc_sheet['B15'] = "Contact number of Purchaser"
    noc_sheet['C15'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!C55"
    
    noc_sheet['B16'] = "Address of Homebuyer of Purchaser"
    noc_sheet['C16'] = f"='{cost_sheet_data['formatted_unit']} - Data Entry'!C56"
    
    # Adjust column widths
    for col, width in [('B', 60), ('C', 30)]:
        noc_sheet.column_dimensions[col].width = width
    
    # Save to BytesIO object
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

def generate_noc_document(customer_info, template_file):
    """Generate NOC document for a customer"""
    if not template_file or not customer_info:
        return None
    
    try:
        # Load the template
        doc = DocxTemplate(template_file)
        
        # Format unit number consistently
        unit_number = customer_info.get('Unit Number')
        if isinstance(unit_number, str):
            unit_number = unit_number.strip().upper()
        else:
            unit_number = str(unit_number).strip().upper()
        
        tower_no = customer_info.get('Tower No', '')
        if tower_no and not isinstance(tower_no, str):
            tower_no = str(tower_no)
        
        # Prepare context data for template
        context = {
            'unit_no': unit_number,
            'floor_no': f"{int(unit_number.split('-')[-1]) // 100}th" if '-' in unit_number else "Ground",
            'building_name': tower_no,
            'booking_date': customer_info.get('Booking date').strftime('%d-%m-%Y'),
            'saleable_area': customer_info.get('Area(sqft)'),
            'carpet_area': customer_info.get('Carpet Area(sqft)'),
            'applicant_name': customer_info.get('Name of Customer'),
            'co_applicant_name': customer_info.get('Co-Applicant Name', 'N/A'),
            'today_date': datetime.now().strftime('%d-%m-%Y')
        }
        
        # Render the template with the data
        doc.render(context)
        
        # Save to a byte stream
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        return output
    
    except Exception as e:
        st.error(f"Error generating NOC document: {str(e)}")
        return None

def create_download_link(file_obj, file_name):
    """Create a download link for a file object"""
    b64 = base64.b64encode(file_obj.read()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{file_name}">Download {file_name}</a>'
    return href

# Sidebar for uploading files
with st.sidebar:
    st.markdown('<div class="section-header">File Upload</div>', unsafe_allow_html=True)
    
    uploaded_sales_mis = st.file_uploader("Upload Sales MIS Template Excel", type=["xlsx", "xls"])
    
    st.markdown('<div class="section-header">Optional</div>', unsafe_allow_html=True)
    uploaded_noc_template = st.file_uploader("Upload NOC Document Template (Optional)", type=["docx"])
    
    if uploaded_noc_template:
        st.session_state.noc_template = uploaded_noc_template

# Main area
st.markdown('<div class="section-header">Data Processing & Verification</div>', unsafe_allow_html=True)

if uploaded_sales_mis:
    # Process the uploaded file
    with st.spinner('Processing Sales MIS Template...'):
        try:
            # Load workbook
            workbook = openpyxl.load_workbook(uploaded_sales_mis, data_only=True)
            
            # Identify the relevant sheets
            sales_master_sheet_name = identify_sales_master_sheet(workbook)
            collection_sheet_name = identify_collection_sheet(workbook)
            
            if not sales_master_sheet_name:
                st.error("Could not identify Annex - Sales Master sheet in the uploaded file.")
            elif not collection_sheet_name:
                st.error("Could not identify Main Collection sheet in the uploaded file.")
            else:
                st.success(f"Successfully identified sheets: '{sales_master_sheet_name}' and '{collection_sheet_name}'")
                
                # Parse Sales Master sheet
                sales_master_df = parse_sales_master(workbook[sales_master_sheet_name])
                st.session_state.sales_master_df = sales_master_df
                
                # Identify account sections in collection sheet
                accounts_info = identify_account_sections(workbook[collection_sheet_name])
                st.session_state.accounts_info = accounts_info
                
                # Parse collection transactions
                collection_df = parse_collection_transactions(workbook[collection_sheet_name], accounts_info)
                st.session_state.collection_df = collection_df
                
                # Verify transactions against customer data
                verification_results = verify_transactions(sales_master_df, collection_df)
                st.session_state.verification_results = verification_results
                
                # Show summary
                st.markdown('<div class="section-header">Verification Summary</div>', unsafe_allow_html=True)
                
                verified_count = sum(1 for v in verification_results.values() if v['status'] == 'verified')
                warning_count = sum(1 for v in verification_results.values() if v['status'] == 'warning')
                error_count = sum(1 for v in verification_results.values() if v['status'] == 'error')
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Verified Units", verified_count, f"{verified_count/len(verification_results)*100:.1f}%")
                with col2:
                    st.metric("Units with Warnings", warning_count, f"{warning_count/len(verification_results)*100:.1f}%")
                with col3:
                    st.metric("Units with Errors", error_count, f"{error_count/len(verification_results)*100:.1f}%")
                
                # Show accounts found
                st.markdown('<div class="section-header">Bank Accounts Identified</div>', unsafe_allow_html=True)
                
                accounts_df = pd.DataFrame([
                    {
                        'Account Name': account['name'],
                        'Account Number': account['number'],
                        'Transaction Count': sum(1 for _, txn in collection_df.iterrows() 
                                            if txn.get('account_number') == account['number'])
                    }
                    for account in accounts_info
                ])
                
                st.dataframe(accounts_df, use_container_width=True)
        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    
    # Show customer selection and verification details if data is processed
    if st.session_state.sales_master_df is not None and st.session_state.verification_results:
        st.markdown('<div class="section-header">Customer Selection</div>', unsafe_allow_html=True)
        
        # Create a DataFrame for display
        customers_df = pd.DataFrame([
            {
                'Select': False,
                'Unit Number': v['unit_number'],
                'Customer Name': v['customer_name'],
                'Expected Amount': v['expected_amount'],
                'Actual Amount': v['actual_amount'],
                'Difference': v['expected_amount'] - v['actual_amount'],
                'Transaction Count': v['transaction_count'],
                'Bounced Transactions': len(v['bounced_transactions']),
                'Status': v['status']
            }
            for k, v in st.session_state.verification_results.items()
        ])
        
        # Sort by unit number
        customers_df = customers_df.sort_values('Unit Number')
        
        # Use st.data_editor to make it selectable
        edited_df = st.data_editor(
            customers_df,
            column_config={
                "Select": st.column_config.CheckboxColumn(
                    "Select",
                    help="Select customer for cost sheet generation",
                    default=False,
                ),
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    help="Verification status",
                    options=["verified", "warning", "error"],
                    required=True,
                )
            },
            disabled=["Unit Number", "Customer Name", "Expected Amount", "Actual Amount", 
                     "Difference", "Transaction Count", "Bounced Transactions", "Status"],
            use_container_width=True,
            hide_index=True
        )
        
        # Store selected customers
        selected_customers = edited_df[edited_df['Select']]['Unit Number'].tolist()
        st.session_state.selected_customers = selected_customers
        
        # Show detailed verification for selected customers
        if selected_customers:
            st.markdown('<div class="section-header">Verification Details</div>', unsafe_allow_html=True)
            
            # Create tabs for each selected customer
            tabs = st.tabs([f"{unit_no}" for unit_no in selected_customers])
            
            for i, tab in enumerate(tabs):
                unit_no = selected_customers[i]
                verification = st.session_state.verification_results.get(unit_no, {})
                
                with tab:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("Customer Information")
                        st.write(f"**Customer:** {verification.get('customer_name', 'N/A')}")
                        st.write(f"**Unit:** {unit_no}")
                        
                        customer_row = st.session_state.sales_master_df[
                            st.session_state.sales_master_df['Unit Number'] == unit_no
                        ]
                        
                        if not customer_row.empty:
                            st.write(f"**Tower:** {customer_row.iloc[0].get('Tower No', 'N/A')}")
                            st.write(f"**Booking Date:** {customer_row.iloc[0].get('Booking date', 'N/A')}")
                            st.write(f"**Payment Plan:** {customer_row.iloc[0].get('Payment Plan', 'N/A')}")
                    
                    with col2:
                        st.subheader("Payment Verification")
                        st.write(f"**Expected Amount:** ₹{verification.get('expected_amount', 0):,.2f}")
                        st.write(f"**Actual Amount:** ₹{verification.get('actual_amount', 0):,.2f}")
                        
                        difference = verification.get('expected_amount', 0) - verification.get('actual_amount', 0)
                        st.write(f"**Difference:** ₹{difference:,.2f}")
                        
                        status = verification.get('status', 'unknown')
                        if status == 'verified':
                            st.success("✅ Payments Verified")
                        elif status == 'warning':
                            st.warning("⚠️ Verification Warning")
                        else:
                            st.error("❌ Verification Failed")
                    
                    # Show transactions
                    st.subheader("Transactions")
                    transactions = verification.get('transactions', [])
                    
                    if transactions:
                        transactions_df = pd.DataFrame(transactions)
                        
                        # Format date column
                        if 'date' in transactions_df.columns:
                            transactions_df['date'] = pd.to_datetime(transactions_df['date']).dt.strftime('%Y-%m-%d')
                        
                        # Select relevant columns
                        display_cols = ['date', 'description', 'type', 'amount', 'account_name', 'sales_tag']
                        display_cols = [col for col in display_cols if col in transactions_df.columns]
                        
                        st.dataframe(transactions_df[display_cols], use_container_width=True)
                    else:
                        st.info("No transactions found for this unit.")
                    
                    # Show bounced transactions if any
                    bounced = verification.get('bounced_transactions', [])
                    if bounced:
                        st.subheader("Potential Bounced Transactions")
                        bounced_df = pd.DataFrame(bounced)
                        st.dataframe(bounced_df, use_container_width=True)
                    
                    # Generate preview data for this customer
                    customer_info = st.session_state.sales_master_df[
                        st.session_state.sales_master_df['Unit Number'] == unit_no
                    ].iloc[0]
                    
                    cost_sheet_data = generate_cost_sheet_data(customer_info, verification)
                    st.session_state.preview_data[unit_no] = cost_sheet_data
                    
                    # Preview button
                    if st.button(f"Preview Cost Sheet for {unit_no}", key=f"preview_{unit_no}"):
                        preview_data = st.session_state.preview_data.get(unit_no)
                        
                        if preview_data:
                            st.subheader("Cost Sheet Preview")
                            
                            preview_tabs = st.tabs(["Data Entry", "Bank Credit Details", "Sales NOC SWAMIH"])
                            
                            with preview_tabs[0]:
                                st.write("**COST SHEET**")
                                
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.write("**Customer Information**")
                                    st.write(f"**APPLICANT NAME:** {preview_data.get('customer_name')}")
                                    st.write(f"**CO-APPLICANT NAME:** {preview_data.get('co_applicant', 'N/A')}")
                                    st.write(f"**BOOKING DATE:** {preview_data.get('booking_date')}")
                                    st.write(f"**PAYMENT PLAN:** {preview_data.get('payment_plan')}")
                                
                                with col2:
                                    st.write("**Unit Information**")
                                    st.write(f"**TOWER:** {preview_data.get('tower')}")
                                    st.write(f"**UNIT NO:** {preview_data.get('unit_number')}")
                                    st.write(f"**FLOOR NUMBER:** {preview_data.get('floor_number')}")
                                    st.write(f"**SUPER AREA:** {preview_data.get('super_area')} sq.ft.")
                                    st.write(f"**CARPET AREA:** {preview_data.get('carpet_area')} sq.ft.")
                                
                                st.write("**Cost Breakdown**")
                                cost_df = pd.DataFrame([
                                    {"Particulars": "BASIC SALE PRICE", "Rate": preview_data.get('bsp_rate'), 
                                     "Amount": preview_data.get('bsp_amount')},
                                    {"Particulars": "IFMS", "Rate": preview_data.get('ifms_rate'), 
                                     "Amount": preview_data.get('ifms_amount')},
                                    {"Particulars": "1 Year Annual Maintenance Charge", "Rate": preview_data.get('amc_rate'), 
                                     "Amount": preview_data.get('amc_amount')},
                                    {"Particulars": "Total Sale Consideration", "Rate": "", 
                                     "Amount": preview_data.get('total_consideration')},
                                    {"Particulars": "GST AMOUNT (5%)", "Rate": "", 
                                     "Amount": preview_data.get('gst_amount')},
                                    {"Particulars": "AMC GST (18%)", "Rate": "", 
                                     "Amount": preview_data.get('amc_gst_amount')},
                                    {"Particulars": "GRAND TOTAL", "Rate": "", 
                                     "Amount": preview_data.get('total_consideration') + 
                                             preview_data.get('gst_amount') + 
                                             preview_data.get('amc_gst_amount')}
                                ])
                                st.dataframe(cost_df, use_container_width=True)
                                
                                st.write("**Receipt Details**")
                                receipt_df = pd.DataFrame([
                                    {"Detail": "Amount Received", "Value": preview_data.get('amount_received')},
                                    {"Detail": "BALANCE RECEIVABLE", "Value": preview_data.get('balance_receivable')},
                                    {"Detail": "GST received", "Value": preview_data.get('gst_received')},
                                    {"Detail": "BALANCE GST RECEIVABLE", "Value": preview_data.get('gst_amount') - 
                                                                                preview_data.get('gst_received')}
                                ])
                                st.dataframe(receipt_df, use_container_width=True)
                            
                            with preview_tabs[1]:
                                st.write("**Bank Credit Details**")
                                
                                transactions = preview_data.get('transactions', [])
                                credit_transactions = [t for t in transactions if t.get('type') == 'C']
                                
                                if credit_transactions:
                                    bank_df = pd.DataFrame([
                                        {
                                            "Date": txn.get('date'),
                                            "Amount received": txn.get('amount'),
                                            "Bank A/c Name": txn.get('account_name'),
                                            "Bank A/c No": txn.get('account_number'),
                                            "Bank statement verified": "Yes"
                                        }
                                        for txn in credit_transactions
                                    ])
                                    
                                    # Add total row
                                    bank_df.loc[len(bank_df)] = {
                                        "Date": "Total Sum received",
                                        "Amount received": bank_df["Amount received"].sum(),
                                        "Bank A/c Name": "",
                                        "Bank A/c No": "",
                                        "Bank statement verified": ""
                                    }
                                    
                                    st.dataframe(bank_df, use_container_width=True)
                                else:
                                    st.info("No credit transactions found for this unit.")
                            
                            with preview_tabs[2]:
                                st.write("**Sales NOC SWAMIH**")
                                
                                noc_df = pd.DataFrame([
                                    {"Particulars": "Flat/Unit No.", "Details": preview_data.get('unit_number')},
                                    {"Particulars": "Floor No.", "Details": preview_data.get('floor_number')},
                                    {"Particulars": "Building Name", "Details": preview_data.get('tower')},
                                    {"Particulars": "Carpet Area of the Flat / Unit (Sq.Ft.)", 
                                     "Details": preview_data.get('carpet_area')},
                                    {"Particulars": "Name of the Applicant (Purchaser)", 
                                     "Details": preview_data.get('customer_name')},
                                    {"Particulars": "Name of the Co-Applicant (Co- Purchaser)", 
                                     "Details": preview_data.get('co_applicant', 'N/A')},
                                    {"Particulars": "Total Sale Consideration (including parking charges) (Excl. GST)", 
                                     "Details": preview_data.get('total_consideration')},
                                    {"Particulars": "Total GST", 
                                     "Details": preview_data.get('gst_amount') + preview_data.get('amc_gst_amount')},
                                    {"Particulars": "Sale Consideration Amount Received as on date (Excl. GST)", 
                                     "Details": preview_data.get('amount_received')},
                                    {"Particulars": "GST Amount received as on date", 
                                     "Details": preview_data.get('gst_received')},
                                    {"Particulars": "Balance Amount yet to be received (excluding taxes)", 
                                     "Details": preview_data.get('balance_receivable')},
                                    {"Particulars": "Booking Date", "Details": preview_data.get('booking_date')},
                                    {"Particulars": "Home loan taken from", "Details": preview_data.get('home_loan')},
                                ])
                                
                                st.dataframe(noc_df, use_container_width=True)
            
            # Generate cost sheets section
            st.markdown('<div class="section-header">Generate Cost Sheets</div>', unsafe_allow_html=True)
            
            if st.button("Generate Cost Sheets for Selected Customers", key="generate_cost_sheets"):
                if not selected_customers:
                    st.warning("Please select at least one customer.")
                else:
                    with st.spinner(f"Generating cost sheets for {len(selected_customers)} customers..."):
                        # Create a temporary directory to store the files
                        with tempfile.TemporaryDirectory() as temp_dir:
                            cost_sheet_files = []
                            noc_files = []
                            
                            # Generate cost sheets for each selected customer
                            for unit_no in selected_customers:
                                # Get customer info and verification data
                                customer_row = st.session_state.sales_master_df[
                                    st.session_state.sales_master_df['Unit Number'] == unit_no
                                ]
                                
                                if customer_row.empty:
                                    continue
                                
                                customer_info = customer_row.iloc[0]
                                verification = st.session_state.verification_results.get(unit_no, {})
                                cost_sheet_data = generate_cost_sheet_data(customer_info, verification)
                                
                                # Generate Excel file
                                excel_file = generate_cost_sheet_excel(cost_sheet_data)
                                
                                if excel_file:
                                    # Save Excel file to temp directory
                                    excel_path = os.path.join(temp_dir, f"COST SHEET-{unit_no}.xlsx")
                                    with open(excel_path, 'wb') as f:
                                        f.write(excel_file.getvalue())
                                    
                                    cost_sheet_files.append((excel_path, f"COST SHEET-{unit_no}.xlsx"))
                                
                                # Generate NOC document if template is available
                                if st.session_state.noc_template:
                                    noc_doc = generate_noc_document(customer_info, st.session_state.noc_template)
                                    
                                    if noc_doc:
                                        # Save NOC file to temp directory
                                        noc_path = os.path.join(temp_dir, f"NOC-{unit_no}.docx")
                                        with open(noc_path, 'wb') as f:
                                            f.write(noc_doc.getvalue())
                                        
                                        noc_files.append((noc_path, f"NOC-{unit_no}.docx"))
                            
                            # Create a zip file if multiple files
                            if len(cost_sheet_files) > 1:
                                zip_path = os.path.join(temp_dir, "Cost_Sheets.zip")
                                with zipfile.ZipFile(zip_path, 'w') as zipf:
                                    for file_path, file_name in cost_sheet_files:
                                        zipf.write(file_path, file_name)
                                    
                                    if noc_files:
                                        for file_path, file_name in noc_files:
                                            zipf.write(file_path, file_name)
                                
                                # Create download link for zip
                                with open(zip_path, 'rb') as f:
                                    zip_data = f.read()
                                
                                st.success(f"Successfully generated {len(cost_sheet_files)} cost sheets and {len(noc_files)} NOC documents.")
                                st.download_button(
                                    label="Download All Files (ZIP)",
                                    data=zip_data,
                                    file_name="Cost_Sheets.zip",
                                    mime="application/zip"
                                )
                            else:
                                # Create individual download links
                                st.success("Cost sheet generated successfully!")
                                
                                for file_path, file_name in cost_sheet_files:
                                    with open(file_path, 'rb') as f:
                                        file_data = f.read()
                                    
                                    st.download_button(
                                        label=f"Download {file_name}",
                                        data=file_data,
                                        file_name=file_name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"download_{file_name}"
                                    )
                                
                                for file_path, file_name in noc_files:
                                    with open(file_path, 'rb') as f:
                                        file_data = f.read()
                                    
                                    st.download_button(
                                        label=f"Download {file_name}",
                                        data=file_data,
                                        file_name=file_name,
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"download_{file_name}"
                                    )

else:
    st.info("Please upload the Sales MIS Template Excel file to start.")

# Footer
st.markdown('<div class="footer">Real Estate Cost Sheet Generator © 2025</div>', unsafe_allow_html=True)
