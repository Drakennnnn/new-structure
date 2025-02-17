import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
import re
import logging
import warnings
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
import json

# Configure logging and suppress warnings
logging.getLogger("streamlit_health_check").setLevel(logging.ERROR)
warnings.filterwarnings('ignore')

# Custom color schemes
COLORS = {
    'primary': ['#1f77b4', '#2ca02c', '#ff7f0e', '#d62728', '#9467bd', '#8c564b'],
    'sequential': px.colors.sequential.Blues,
    'background': 'rgba(0,0,0,0)',
    'text': '#ffffff'
}

@dataclass
class PhaseInfo:
    """Store phase information"""
    phase_number: int
    account_number: str
    start_row: int
    end_row: Optional[int] = None
    total_credits: float = 0.0
    total_debits: float = 0.0
    running_balance: float = 0.0

class PhaseProcessor:
    """Enhanced phase-wise data processing"""
    
    def __init__(self):
        self.phase_pattern = r'Main Collection Escrow A/c Phase-(\d+)'
        self.account_pattern = r'\d{14}'

    def extract_phase_info(self, row: pd.Series) -> Optional[Tuple[int, str]]:
        """Extract phase number and account number from row"""
        for val in row.values:
            val_str = str(val)
            phase_match = re.search(self.phase_pattern, val_str)
            if phase_match:
                phase_num = int(phase_match.group(1))
                # Find account number in this row
                account_num = None
                for cell in row.values:
                    account_match = re.search(self.account_pattern, str(cell))
                    if account_match:
                        account_num = account_match.group(0)
                        break
                return phase_num, account_num
        return None

    def find_phase_sections(self, df: pd.DataFrame) -> Dict[str, PhaseInfo]:
        """Find phase sections with improved detection"""
        phases = {}
        current_phase = None
        
        # Check first row for Phase-1
        first_info = self.extract_phase_info(df.iloc[0])
        if first_info:
            phase_num, account_num = first_info
            current_phase = f"Phase-{phase_num}"
            phases[current_phase] = PhaseInfo(
                phase_number=phase_num,
                account_number=account_num,
                start_row=1  # Skip header row
            )
        
        # Find other phases
        for idx, row in df.iterrows():
            phase_info = self.extract_phase_info(row)
            if phase_info:
                phase_num, account_num = phase_info
                
                # Set end for previous phase
                if current_phase:
                    phases[current_phase].end_row = idx - 1
                
                current_phase = f"Phase-{phase_num}"
                phases[current_phase] = PhaseInfo(
                    phase_number=phase_num,
                    account_number=account_num,
                    start_row=idx + 1
                )
        
        # Set end for last phase
        if current_phase:
            phases[current_phase].end_row = len(df)
        
        return phases

    def process_phase_data(self, df: pd.DataFrame, phase_info: PhaseInfo) -> pd.DataFrame:
        """Process data for a specific phase"""
        if phase_info.end_row is None:
            phase_info.end_row = len(df)
            
        phase_data = df.iloc[phase_info.start_row:phase_info.end_row].copy()
        
        # Clean amount data
        phase_data['Amount'] = pd.to_numeric(
            phase_data['Amount'].astype(str).str.replace(',', ''),
            errors='coerce'
        ).fillna(0)
        
        # Calculate running balance
        phase_data['Running Balance'] = phase_data.apply(
            lambda x: x['Amount'] if x['Dr/Cr'] == 'C' else -x['Amount'],
            axis=1
        ).cumsum()
        
        # Update phase totals
        phase_info.total_credits = phase_data[phase_data['Dr/Cr'] == 'C']['Amount'].sum()
        phase_info.total_debits = phase_data[phase_data['Dr/Cr'] == 'D']['Amount'].sum()
        phase_info.running_balance = phase_data['Running Balance'].iloc[-1] if not phase_data.empty else 0
        
        # Add flags
        phase_data['Is_Bounce'] = phase_data.apply(
            lambda x: 'Bounce' in str(x.get('Sales Tag', '')) or 
                     'RET-' in str(x.get('Description', '')),
            axis=1
        )
        
        phase_data['Receipt_Status'] = phase_data['Sales Tag'].apply(
            lambda x: 'Pending' if 'RECEIPT NOT GENERATED' in str(x) else 'Generated'
        )
        
        return phase_data

class TransactionProcessor:
    """Enhanced transaction processing and categorization"""
    
    def __init__(self):
        self.transaction_patterns = {
            'Cheque': ['CHQ', 'CHEQUE', 'MICR'],
            'UPI': ['UPI'],
            'NEFT': ['NEFT'],
            'RTGS': ['RTGS'],
            'IMPS': ['IMPS'],
            'Transfer': ['TRANSFER', 'TRF'],
            'Cash': ['CASH'],
            'DD': ['DD', 'DEMAND DRAFT']
        }

    def categorize_transaction(self, description: str) -> str:
        """Categorize transaction based on description"""
        description = str(description).upper()
        
        for category, patterns in self.transaction_patterns.items():
            if any(pattern in description for pattern in patterns):
                return category
        
        return 'Other'

    def extract_unit_info(self, tag: str) -> Optional[str]:
        """Extract unit number from sales tag"""
        if pd.isna(tag):
            return None
            
        tag = str(tag).strip()
        unit_match = re.search(r'CA \d{2}-\d+', tag)
        if unit_match:
            return unit_match.group(0)
        return None

    def process_transaction(self, row: pd.Series) -> Dict[str, Any]:
        """Process a single transaction with enhanced details"""
        return {
            'date': pd.to_datetime(row['Value Date'], errors='coerce'),
            'amount': pd.to_numeric(
                str(row['Amount']).replace(',', ''), 
                errors='coerce'
            ),
            'type': 'CREDIT' if row['Dr/Cr'] == 'C' else 'DEBIT',
            'is_bounce': bool(
                row.get('Is_Bounce', False) or 
                'Bounce' in str(row.get('Sales Tag', ''))
            ),
            'receipt_status': row.get('Receipt_Status', 'Generated'),
            'transaction_mode': self.categorize_transaction(row['Description']),
            'unit_no': self.extract_unit_info(row.get('Sales Tag', '')),
            'phase': row.get('Phase', 'Unknown'),
            'account': row.get('Account', 'Unknown')
        }

@dataclass
class ProcessedFile:
    """Store processed data from a single file"""
    filename: str
    collection_df: pd.DataFrame
    sales_master_df: pd.DataFrame
    unsold_df: pd.DataFrame
    phases: Dict[str, PhaseInfo]
    date: datetime

class DataProcessor:
    """Enhanced data processing with multiple file support"""
    def __init__(self):
        self.phase_processor = PhaseProcessor()
        self.transaction_processor = TransactionProcessor()
        
    def process_dates(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process date columns with enhanced format handling"""
        date_formats = [
            '%d/%m/%Y %H:%M:%S',
            '%d/%m/%Y',
            '%d-%m-%Y',
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d',
            '%d-%b-%y',
            '%d-%B-%y',
            '%d/%b/%y'
        ]
        
        for col in df.columns:
            if any(text in col.lower() for text in ['date', 'dt']):
                for fmt in date_formats:
                    try:
                        df[col] = pd.to_datetime(df[col], format=fmt, errors='coerce')
                        if not df[col].isna().all():
                            break
                    except:
                        continue
                        
                # If all formats failed, try a final generic parse
                if df[col].isna().all():
                    df[col] = pd.to_datetime(df[col], errors='coerce')
        
        return df

    def process_amounts(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process amount columns with enhanced validation"""
        amount_patterns = [
            'amount', 'balance', 'consideration', 'demanded', 
            'received', 'receivables', 'total', 'bsp'
        ]
        
        for col in df.columns:
            if any(pattern in col.lower() for pattern in amount_patterns):
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.replace('₹', '')
                    .str.replace(',', '')
                    .str.replace(' ', '')
                    .str.strip()
                )
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        return df

    def process_collection_account(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, PhaseInfo]]:
        """Process collection account data with enhanced phase handling"""
        df = df.copy()
        
        # Find phase sections
        phases = self.phase_processor.find_phase_sections(df)
        
        # Process dates and amounts
        df = self.process_dates(df)
        df = self.process_amounts(df)
        
        # Add payment mode
        df['Payment Mode'] = df['Description'].apply(
            self.transaction_processor.categorize_transaction
        )
        
        # Process each phase
        processed_phases = {}
        for phase_name, phase_info in phases.items():
            processed_df = self.phase_processor.process_phase_data(df, phase_info)
            processed_df['Phase'] = phase_name
            processed_df['Account'] = phase_info.account_number
            processed_phases[phase_name] = processed_df
        
        # Combine processed data
        processed_df = pd.concat(processed_phases.values(), ignore_index=True)
        
        return processed_df, phases

    def process_sales_master(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process sales master data with enhanced column detection"""
        df = df.copy()
        
        # Skip summary rows
        header_row = None
        for idx, row in df.iterrows():
            if row.notna().any() and any('Sr. No.' in str(val) for val in row):
                header_row = idx
                break
        
        if header_row is not None:
            df = df.iloc[header_row:].reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
        
        # Clean column names
        df.columns = df.columns.str.strip()
        
        # Process dates and amounts
        df = self.process_dates(df)
        df = self.process_amounts(df)
        
        # Calculate collection percentage
        amount_cols = [col for col in df.columns if 'amount' in col.lower()]
        received_col = next((col for col in amount_cols if 'received' in col.lower()), None)
        demanded_col = next((col for col in amount_cols if 'demanded' in col.lower()), None)
        
        if received_col and demanded_col:
            df['Collection Percentage'] = (
                df[received_col] / df[demanded_col] * 100
            ).clip(0, 100)
        
        return df

    def process_unsold_inventory(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process unsold inventory with enhanced validation"""
        df = df.copy()
        
        # Clean area column
        area_cols = [col for col in df.columns if 'area' in col.lower()]
        if area_cols:
            for col in area_cols:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(',', ''),
                    errors='coerce'
                ).fillna(0)
        
        # Add status if missing
        if 'Status' not in df.columns:
            df['Status'] = 'Unsold'
        
        return df

    def process_file(self, file) -> ProcessedFile:
        """Process a single Excel file with enhanced error handling"""
        try:
            excel_file = pd.ExcelFile(file)
            filename = getattr(file, 'name', 'Unknown')
            
            # Process Collection Account
            collection_df = pd.read_excel(
                excel_file, 
                'Main Collection AC P1_P2_P3',
                skiprows=1
            )
            collection_df, phases = self.process_collection_account(collection_df)
            
            # Process Sales Master
            sales_master_df = pd.read_excel(
                excel_file,
                'Annex - Sales Master',
                skiprows=3
            )
            sales_master_df = self.process_sales_master(sales_master_df)
            
            # Process Unsold Inventory
            unsold_df = pd.read_excel(
                excel_file,
                'Un Sold Inventory'
            )
            unsold_df = self.process_unsold_inventory(unsold_df)
            
            # Extract date from filename or use current date
            try:
                date_str = re.search(r'\d{8}', filename)
                if date_str:
                    file_date = datetime.strptime(date_str.group(0), '%Y%m%d')
                else:
                    file_date = datetime.now()
            except:
                file_date = datetime.now()
            
            return ProcessedFile(
                filename=filename,
                collection_df=collection_df,
                sales_master_df=sales_master_df,
                unsold_df=unsold_df,
                phases=phases,
                date=file_date
            )
        except Exception as e:
            raise Exception(f"Error processing file {getattr(file, 'name', 'Unknown')}: {str(e)}")

def format_currency(value: float) -> str:
    """Format currency values with enhanced validation"""
    try:
        return f"₹{float(value):,.2f}"
    except:
        return "₹0.00"

def format_percentage(value: float) -> str:
    """Format percentage values with enhanced validation"""
    try:
        return f"{float(value):.1f}%"
    except:
        return "0.0%"

def create_phase_summary(collection_df: pd.DataFrame, 
                        phases: Dict[str, PhaseInfo]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Create phase-wise summary with enhanced aggregation"""
    summaries = []
    
    try:
        for phase_name, phase_info in phases.items():
            try:
                phase_data = collection_df[
                    collection_df['Phase'] == phase_name
                ]
                
                credits = phase_data[
                    phase_data['Dr/Cr'].str.strip().str.upper() == 'C'
                ]['Amount'].sum()
                
                debits = phase_data[
                    phase_data['Dr/Cr'].str.strip().str.upper() == 'D'
                ]['Amount'].sum()
                
                last_balance = phase_data['Running Balance'].iloc[-1] if len(phase_data) > 0 else 0
                
                summaries.append({
                    'Phase': phase_name,
                    'Credits': float(credits),
                    'Debits': float(debits),
                    'Net': float(credits - debits),
                    'Last_Balance': float(last_balance),
                    'Account': phase_info.account_number
                })
            except Exception as e:
                st.error(f"Error processing phase {phase_name}: {str(e)}")
                continue
        
        # Create summary DataFrame
        summary_df = pd.DataFrame(summaries)
        
        # Create plot DataFrame
        plot_df = pd.melt(
            summary_df,
            id_vars=['Phase'],
            value_vars=['Credits', 'Debits'],
            var_name='Type',
            value_name='Amount'
        )
        
        return plot_df, summary_df
    except Exception as e:
        st.error(f"Error creating summary: {str(e)}")
        return pd.DataFrame(columns=['Phase', 'Type', 'Amount']), pd.DataFrame()

def create_payment_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Create payment mode summary with enhanced aggregation"""
    if df.empty:
        return pd.DataFrame(columns=['Mode', 'Amount', 'Count', 'Percentage'])
    
    try:    
        payment_summary = df[
            df['Dr/Cr'] == 'C'
        ].groupby('Payment Mode').agg({
            'Amount': ['sum', 'count']
        }).reset_index()
        
        payment_summary.columns = ['Mode', 'Amount', 'Count']
        total_amount = payment_summary['Amount'].sum()
        
        if total_amount > 0:
            payment_summary['Percentage'] = (
                payment_summary['Amount'] / total_amount * 100
            )
        else:
            payment_summary['Percentage'] = 0
        
        return payment_summary
    except Exception as e:
        st.error(f"Error creating payment summary: {str(e)}")
        return pd.DataFrame(columns=['Mode', 'Amount', 'Count', 'Percentage'])

def compare_periods(df1: ProcessedFile, df2: ProcessedFile) -> Dict[str, pd.DataFrame]:
    """Compare data between two periods"""
    try:
        comparisons = {}
        
        # Collection comparison
        collection_comparison = pd.DataFrame({
            'Metric': ['Total Collections', 'Total Units', 'Collection Percentage'],
            df1.filename: [
                df1.collection_df[df1.collection_df['Dr/Cr'] == 'C']['Amount'].sum(),
                len(df1.sales_master_df),
                df1.sales_master_df['Collection Percentage'].mean() if 'Collection Percentage' in df1.sales_master_df else 0
            ],
            df2.filename: [
                df2.collection_df[df2.collection_df['Dr/Cr'] == 'C']['Amount'].sum(),
                len(df2.sales_master_df),
                df2.sales_master_df['Collection Percentage'].mean() if 'Collection Percentage' in df2.sales_master_df else 0
            ]
        })
        
        collection_comparison['Change'] = (
            collection_comparison[df2.filename] - 
            collection_comparison[df1.filename]
        )
        
        comparisons['collection'] = collection_comparison
        
        # Payment mode comparison
        payment_summary1 = create_payment_summary(df1.collection_df)
        payment_summary2 = create_payment_summary(df2.collection_df)
        
        payment_comparison = pd.merge(
            payment_summary1, 
            payment_summary2,
            on='Mode',
            suffixes=('_1', '_2'),
            how='outer'
        ).fillna(0)
        
        comparisons['payment_modes'] = payment_comparison
        
        return comparisons
    except Exception as e:
        st.error(f"Error comparing periods: {str(e)}")
        return {}

def process_collection_trends(df: pd.DataFrame, 
                            start_date: Optional[datetime] = None,
                            end_date: Optional[datetime] = None) -> pd.DataFrame:
    """Process collection trends with enhanced date handling"""
    try:
        # Ensure Value Date is datetime
        df['Value Date'] = pd.to_datetime(df['Value Date'], errors='coerce')
        df = df[df['Value Date'].notna()]
        
        if start_date and end_date:
            df = df[
                (df['Value Date'].dt.date >= start_date) &
                (df['Value Date'].dt.date <= end_date)
            ]
        
        # Calculate monthly trends
        monthly = df[df['Dr/Cr'] == 'C'].copy()
        monthly['Month'] = monthly['Value Date'].dt.to_period('M')
        monthly_summary = monthly.groupby('Month')['Amount'].sum().reset_index()
        monthly_summary['Month'] = monthly_summary['Month'].dt.to_timestamp()
        
        return monthly_summary
    except Exception as e:
        st.error(f"Error processing collection trends: {str(e)}")
        return pd.DataFrame(columns=['Month', 'Amount'])

def filter_collection_data(df: pd.DataFrame, 
                         phase_name: Optional[str] = None,
                         start_date: Optional[datetime] = None,
                         end_date: Optional[datetime] = None,
                         payment_modes: Optional[List[str]] = None) -> pd.DataFrame:
    """Filter collection data with enhanced validation"""
    try:
        filtered_df = df.copy()
        
        if phase_name and phase_name != "All Phases":
            filtered_df = filtered_df[filtered_df['Phase'] == phase_name]
        
        if start_date and end_date:
            filtered_df['Value Date'] = pd.to_datetime(
                filtered_df['Value Date'], 
                errors='coerce'
            )
            filtered_df = filtered_df[
                (filtered_df['Value Date'].dt.date >= start_date) &
                (filtered_df['Value Date'].dt.date <= end_date)
            ]
        
        if payment_modes:
            filtered_df = filtered_df[
                filtered_df['Payment Mode'].isin(payment_modes)
            ]
        
        return filtered_df
    except Exception as e:
        st.error(f"Error filtering data: {str(e)}")
        return df

# Set page configuration
st.set_page_config(
    page_title="Real Estate Analytics Dashboard",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main { 
        padding: 0rem 1rem;
        color: white;
    }
    .stSelectbox { margin-bottom: 1rem; }
    .plot-container {
        margin-bottom: 2rem;
        background-color: rgba(0,0,0,0) !important;
        padding: 1rem;
        border-radius: 0.5rem;
    }
    .custom-metric {
        background-color: rgba(255,255,255,0.1);
        padding: 1rem;
        border-radius: 0.5rem;
    }
    .section-title {
        font-size: 1.5rem;
        font-weight: bold;
        margin-bottom: 1rem;
        color: #ffffff;
    }
    .dataframe {
        font-size: 12px;
        color: white !important;
    }
    .stDataFrame {
        background-color: rgba(0,0,0,0) !important;
        border-radius: 0.5rem;
        padding: 1rem;
    }
    div[data-testid="stMetricValue"] {
        font-size: 2rem;
        color: white !important;
    }
    .stAlert {
        border-radius: 0.5rem;
    }
    .plot-container > div {
        border-radius: 0.5rem;
    }
    div.css-12w0qpk.e1tzin5v1 {
        background-color: rgba(255,255,255,0.1);
        border-radius: 0.5rem;
        padding: 1rem;
    }
    div.css-1r6slb0.e1tzin5v2 {
        background-color: rgba(255,255,255,0.1);
        border-radius: 0.5rem;
        padding: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    st.title("Real Estate Analytics Dashboard")
    st.markdown("---")

    # Initialize data processor
    processor = DataProcessor()

    # Initialize session state
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = {}

    # File upload
    uploaded_files = st.file_uploader(
        "Upload MIS Excel File(s)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )

    if uploaded_files:
        for file in uploaded_files:
            try:
                if file.name not in st.session_state.processed_files:
                    with st.spinner(f'Processing {file.name}...'):
                        processed_file = processor.process_file(file)
                        st.session_state.processed_files[file.name] = processed_file
                        st.success(f"Processed {file.name} successfully!")
            except Exception as e:
                st.error(f"Error processing {file.name}: {str(e)}")

    if st.session_state.processed_files:
        # File selection
        selected_file = st.selectbox(
            "Select File to Analyze",
            options=list(st.session_state.processed_files.keys()),
            key="file_selector"
        )
        
        current_file = st.session_state.processed_files[selected_file]

        # Compare with another file
        if len(st.session_state.processed_files) > 1:
            comparison_file = st.selectbox(
                "Compare with",
                options=[f for f in st.session_state.processed_files.keys() if f != selected_file],
                key="comparison_selector"
            )
            
            if comparison_file:
                comparisons = compare_periods(
                    current_file,
                    st.session_state.processed_files[comparison_file]
                )
                
                st.markdown("### Period Comparison")
                st.dataframe(comparisons['collection'])

        # Main tabs
        tab1, tab2, tab3, tab4 = st.tabs([
            "Project Overview",
            "Collection Analysis",
            "Sales Status",
            "Unit Details"
        ])

        with tab1:
            st.markdown("### Project Overview")
            
            # Phase selection
            selected_phase = st.selectbox(
                "Select Phase",
                ["All Phases"] + list(current_file.phases.keys())
            )
            
            if selected_phase != "All Phases":
                phase_info = current_file.phases[selected_phase]
                filtered_df = current_file.collection_df[
                    current_file.collection_df['Phase'] == selected_phase
                ]
                st.info(f"Showing data for {selected_phase} (Account: {phase_info.account_number})")
            else:
                filtered_df = current_file.collection_df

            # Summary metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                total_units = len(current_file.sales_master_df)
                unsold_units = len(current_file.unsold_df)
                st.metric(
                    "Total Units",
                    f"{total_units}",
                    delta=f"{unsold_units} Unsold",
                    delta_color="inverse"
                )

            with col2:
                # Find consideration column
                consideration_cols = [
                    col for col in current_file.sales_master_df.columns 
                    if 'consideration' in col.lower()
                ]
                if consideration_cols:
                    total_consideration = current_file.sales_master_df[
                        consideration_cols[0]
                    ].sum()
                    st.metric(
                        "Total Consideration",
                        format_currency(total_consideration),
                        delta="View Details"
                    )

            with col3:
                # Find collection columns
                collection_cols = {
                    'received': next((col for col in current_file.sales_master_df.columns 
                                    if 'received' in col.lower()), None),
                    'demanded': next((col for col in current_file.sales_master_df.columns 
                                    if 'demanded' in col.lower()), None)
                }
                
                if all(collection_cols.values()):
                    total_collection = current_file.sales_master_df[
                        collection_cols['received']
                    ].sum()
                    total_demand = current_file.sales_master_df[
                        collection_cols['demanded']
                    ].sum()
                    
                    collection_percentage = (total_collection / total_demand * 100)
                    pending_collection = total_demand - total_collection
                    
                    st.metric(
                        "Collection Achievement",
                        format_percentage(collection_percentage),
                        delta=format_currency(pending_collection),
                        delta_color="inverse"
                    )

            with col4:
                # Find area column
                area_cols = [
                    col for col in current_file.sales_master_df.columns 
                    if 'area' in col.lower()
                ]
                if area_cols:
                    area_col = area_cols[0]
                    total_area = current_file.sales_master_df[area_col].sum()
                    unsold_area = current_file.unsold_df[area_col].sum() if area_col in current_file.unsold_df else 0
                    
                    st.metric(
                        "Total Area",
                        f"{total_area:,.0f} sq.ft",
                        delta=f"{unsold_area:,.0f} sq.ft unsold",
                        delta_color="inverse"
                    )

            # Phase-wise collection summary
            st.markdown("### Phase-wise Collection Summary")
            
            plot_df, summary_df = create_phase_summary(
                current_file.collection_df,
                current_file.phases
            )
            
            if not plot_df.empty:
                col1, col2 = st.columns(2)
                
                with col1:
                    fig_phase = px.bar(
                        plot_df,
                        x='Phase',
                        y='Amount',
                        color='Type',
                        title="Collections by Phase",
                        barmode='group',
                        color_discrete_sequence=COLORS['primary']
                    )
                    fig_phase.update_layout(
                        plot_bgcolor=COLORS['background'],
                        paper_bgcolor=COLORS['background'],
                        font_color=COLORS['text']
                    )
                    st.plotly_chart(fig_phase, use_container_width=True)
                
                with col2:
                    display_summary = summary_df.copy()
                    for col in ['Credits', 'Debits', 'Net', 'Last_Balance']:
                        if col in display_summary.columns:
                            display_summary[col] = display_summary[col].apply(format_currency)
                    st.dataframe(display_summary, use_container_width=True)
            else:
                st.warning("No phase data available for visualization")

            # Payment mode analysis
            st.markdown("### Payment Mode Analysis")
            
            payment_modes = st.multiselect(
                "Select Payment Modes",
                options=filtered_df['Payment Mode'].unique(),
                default=filtered_df['Payment Mode'].unique()
            )
            
            filtered_df = filtered_df[filtered_df['Payment Mode'].isin(payment_modes)]
            payment_summary = create_payment_summary(filtered_df)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if not payment_summary.empty:
                    fig_payment = px.pie(
                        payment_summary,
                        values='Amount',
                        names='Mode',
                        title="Collection by Payment Mode",
                        color_discrete_sequence=COLORS['primary']
                    )
                    fig_payment.update_layout(
                        plot_bgcolor=COLORS['background'],
                        paper_bgcolor=COLORS['background'],
                        font_color=COLORS['text']
                    )
                    st.plotly_chart(fig_payment, use_container_width=True)
                else:
                    st.warning("No payment data available")
            
            with col2:
                if not payment_summary.empty:
                    display_payment = payment_summary.copy()
                    display_payment['Amount'] = display_payment['Amount'].apply(format_currency)
                    display_payment['Percentage'] = display_payment['Percentage'].apply(format_percentage)
                    st.dataframe(display_payment, use_container_width=True)

            # Collection trends
            st.markdown("### Collection Trends")
            
            try:
                # Date range filter
                col1, col2 = st.columns(2)
                with col1:
                    start_date = st.date_input(
                        "Start Date",
                        value=filtered_df['Value Date'].min()
                    )
                with col2:
                    end_date = st.date_input(
                        "End Date",
                        value=filtered_df['Value Date'].max()
                    )
                
                filtered_df = filter_collection_data(
                    filtered_df,
                    phase_name=selected_phase if selected_phase != "All Phases" else None,
                    start_date=start_date,
                    end_date=end_date,
                    payment_modes=payment_modes
                )
                
                if not filtered_df.empty:
                    monthly_collection = process_collection_trends(
                        filtered_df,
                        start_date,
                        end_date
                    )
                    
                    fig_monthly = px.line(
                        monthly_collection,
                        x='Month',
                        y='Amount',
                        title="Monthly Collection Trend",
                        markers=True
                    )
                    fig_monthly.update_layout(
                        plot_bgcolor=COLORS['background'],
                        paper_bgcolor=COLORS['background'],
                        font_color=COLORS['text']
                    )
                    st.plotly_chart(fig_monthly, use_container_width=True)
                    
                    # Show detailed transactions
                    st.markdown("### Transaction Details")
                    
                    # Column selection
                    default_columns = [
                        'Value Date', 'Description', 'Dr/Cr', 
                        'Amount', 'Payment Mode', 'Sales Tag'
                    ]
                    
                    display_columns = st.multiselect(
                        "Select columns to display",
                        options=filtered_df.columns,
                        default=[col for col in default_columns if col in filtered_df.columns]
                    )
                    
                    display_df = filtered_df[display_columns].copy()
                    if 'Amount' in display_columns:
                        display_df['Amount'] = display_df['Amount'].apply(format_currency)
                    if 'Running Balance' in display_columns:
                        display_df['Running Balance'] = display_df['Running Balance'].apply(format_currency)
                    
                    st.dataframe(
                        display_df.sort_values('Value Date', ascending=False),
                        use_container_width=True
                    )
                else:
                    st.warning("No data available for selected filters")
                    
            except Exception as e:
                st.error(f"Error processing collection trends: {str(e)}")

        with tab2:
            st.markdown("### Collection Analysis")
            
            # Phase filter for collection analysis
            selected_phase = st.selectbox(
                "Select Phase",
                ["All Phases"] + list(current_file.phases.keys()),
                key="collection_phase"
            )
            
            if selected_phase != "All Phases":
                collection_data = current_file.collection_df[
                    current_file.collection_df['Phase'] == selected_phase
                ]
                phase_info = current_file.phases[selected_phase]
                st.info(f"Analyzing {selected_phase} (Account: {phase_info.account_number})")
            else:
                collection_data = current_file.collection_df.copy()
            
            # Collection status
            st.markdown("### Collection Status")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Collection by payment mode
                mode_summary = collection_data[
                    collection_data['Dr/Cr'] == 'C'
                ].groupby('Payment Mode')['Amount'].agg(['sum', 'count']).reset_index()
                
                mode_summary.columns = ['Mode', 'Amount', 'Count']
                mode_summary['Percentage'] = (
                    mode_summary['Amount'] / mode_summary['Amount'].sum() * 100
                )
                
                fig_mode = px.pie(
                    mode_summary,
                    values='Amount',
                    names='Mode',
                    title="Collections by Payment Mode"
                )
                fig_mode.update_layout(
                    plot_bgcolor=COLORS['background'],
                    paper_bgcolor=COLORS['background'],
                    font_color=COLORS['text']
                )
                st.plotly_chart(fig_mode, use_container_width=True)
            
            with col2:
                st.dataframe(
                    mode_summary.assign(
                        Amount=mode_summary['Amount'].apply(format_currency),
                        Percentage=mode_summary['Percentage'].apply(format_percentage)
                    ),
                    use_container_width=True
                )

            # Collection trends for selected phase
            st.markdown("### Collection Trends")
            
            try:
                monthly_collection = process_collection_trends(collection_data)
                
                fig_trend = px.line(
                    monthly_collection,
                    x='Month',
                    y='Amount',
                    title="Monthly Collection Trend",
                    markers=True
                )
                fig_trend.update_layout(
                    plot_bgcolor=COLORS['background'],
                    paper_bgcolor=COLORS['background'],
                    font_color=COLORS['text']
                )
                st.plotly_chart(fig_trend, use_container_width=True)
            except Exception as e:
                st.error(f"Error creating collection trend: {str(e)}")

        with tab3:
            st.markdown("### Sales Status")
            
            # Filters
            col1, col2, col3 = st.columns(3)
            
            with col1:
                tower_filter = st.multiselect(
                    "Select Towers",
                    options=list(current_file.sales_master_df['Tower'].unique()),
                    default=list(current_file.sales_master_df['Tower'].unique())
                )
            
            with col2:
                unit_type_cols = [col for col in current_file.sales_master_df.columns 
                                if 'type' in col.lower() and 'unit' in col.lower()]
                if unit_type_cols:
                    unit_type_col = unit_type_cols[0]
                    unit_type_filter = st.multiselect(
                        "Select Unit Types",
                        options=list(current_file.sales_master_df[unit_type_col].unique()),
                        default=list(current_file.sales_master_df[unit_type_col].unique())
                    )
                else:
                    unit_type_filter = []
            
            with col3:
                collection_range = st.slider(
                    "Collection Percentage Range",
                    min_value=0,
                    max_value=100,
                    value=(0, 100),
                    step=5
                )

            # Filter sales data
            filtered_sales = current_file.sales_master_df[
                current_file.sales_master_df['Tower'].isin(tower_filter)
            ]
            
            if unit_type_cols and unit_type_filter:
                filtered_sales = filtered_sales[
                    filtered_sales[unit_type_col].isin(unit_type_filter)
                ]
            
            if 'Collection Percentage' in filtered_sales.columns:
                filtered_sales = filtered_sales[
                    (filtered_sales['Collection Percentage'] >= collection_range[0]) &
                    (filtered_sales['Collection Percentage'] <= collection_range[1])
                ]

            # Sales Analysis visualizations
            col1, col2 = st.columns(2)
            
            with col1:
                tower_sales = filtered_sales.groupby('Tower').size().reset_index()
                tower_sales.columns = ['Tower', 'Count']
                
                fig_tower = px.bar(
                    tower_sales,
                    x='Tower',
                    y='Count',
                    title="Units Sold by Tower",
                    color='Count',
                    color_continuous_scale=COLORS['sequential']
                )
                fig_tower.update_layout(
                    plot_bgcolor=COLORS['background'],
                    paper_bgcolor=COLORS['background'],
                    font_color=COLORS['text']
                )
                st.plotly_chart(fig_tower, use_container_width=True)
            
            with col2:
                if unit_type_cols:
                    unit_type_dist = filtered_sales[unit_type_col].value_counts()
                    
                    fig_unit_type = px.pie(
                        values=unit_type_dist.values,
                        names=unit_type_dist.index,
                        title="Unit Type Distribution",
                        color_discrete_sequence=COLORS['primary']
                    )
                    fig_unit_type.update_layout(
                        plot_bgcolor=COLORS['background'],
                        paper_bgcolor=COLORS['background'],
                        font_color=COLORS['text']
                    )
                    st.plotly_chart(fig_unit_type, use_container_width=True)

            # Collection percentage distribution
            st.markdown("### Collection Analysis")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if 'Collection Percentage' in filtered_sales.columns:
                    collection_dist = pd.cut(
                        filtered_sales['Collection Percentage'],
                        bins=[0, 25, 50, 75, 90, 100],
                        labels=['0-25%', '25-50%', '50-75%', '75-90%', '90-100%']
                    ).value_counts()

                    fig_collection = px.pie(
                        values=collection_dist.values,
                        names=collection_dist.index,
                        title="Collection Percentage Distribution",
                        color_discrete_sequence=COLORS['primary']
                    )
                    fig_collection.update_layout(
                        plot_bgcolor=COLORS['background'],
                        paper_bgcolor=COLORS['background'],
                        font_color=COLORS['text']
                    )
                    st.plotly_chart(fig_collection, use_container_width=True)
            
            with col2:
                # BSP Analysis
                bsp_cols = [col for col in filtered_sales.columns if 'bsp' in col.lower()]
                if bsp_cols and unit_type_cols:
                    fig_bsp = px.box(
                        filtered_sales,
                        x='Tower',
                        y=bsp_cols[0],
                        color=unit_type_col,
                        title="BSP Distribution by Tower and Unit Type"
                    )
                    fig_bsp.update_layout(
                        plot_bgcolor=COLORS['background'],
                        paper_bgcolor=COLORS['background'],
                        font_color=COLORS['text']
                    )
                    st.plotly_chart(fig_bsp, use_container_width=True)

        with tab4:
            st.markdown("### Unit Details")
            
            # Column selection
            all_columns = current_file.sales_master_df.columns.tolist()
            unit_cols = [col for col in all_columns if any(text in col.lower() 
                        for text in ['apartment', 'unit', 'tower', 'type', 'area', 'bsp', 
                                   'consideration', 'received', 'collection'])]
            
            selected_columns = st.multiselect(
                "Select columns to display",
                options=all_columns,
                default=unit_cols
            )

            # Format and display data
            display_sales = filtered_sales[selected_columns].copy()
            
            # Format numeric columns based on patterns
            numeric_formats = {
                'area': lambda x: f"{x:,.0f}",
                'bsp': lambda x: f"₹{x:,.2f}",
                'consideration': lambda x: f"₹{x:,.2f}",
                'amount': lambda x: f"₹{x:,.2f}",
                'collection': lambda x: f"{x:.1f}%",
                'balance': lambda x: f"₹{x:,.2f}"
            }
            
            for col in selected_columns:
                for key, formatter in numeric_formats.items():
                    if key in col.lower():
                        display_sales[col] = display_sales[col].apply(formatter)

            st.dataframe(
                display_sales.sort_values('Apartment No' if 'Apartment No' in selected_columns else selected_columns[0]),
                use_container_width=True
            )

            # Export options
            col1, col2 = st.columns(2)
            
            with col1:
                csv = filtered_sales[selected_columns].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Filtered Data as CSV",
                    data=csv,
                    file_name=f"unit_details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with col2:
                # Summary statistics
                summary_cols = {
                    'area': next((col for col in selected_columns if 'area' in col.lower()), None),
                    'consideration': next((col for col in selected_columns if 'consideration' in col.lower()), None),
                    'received': next((col for col in selected_columns if 'received' in col.lower()), None),
                    'bsp': next((col for col in selected_columns if 'bsp' in col.lower()), None),
                    'collection': 'Collection Percentage' if 'Collection Percentage' in selected_columns else None
                }
                
                summary_stats = pd.DataFrame({
                    'Metric': ['Total Units', 'Total Area', 'Total Consideration', 
                             'Total Collection', 'Average BSP', 'Average Collection %'],
                    'Value': [
                        len(filtered_sales),
                        f"{filtered_sales[summary_cols['area']].sum():,.0f} sq.ft" if summary_cols['area'] else "N/A",
                        format_currency(filtered_sales[summary_cols['consideration']].sum()) if summary_cols['consideration'] else "N/A",
                        format_currency(filtered_sales[summary_cols['received']].sum()) if summary_cols['received'] else "N/A",
                        format_currency(filtered_sales[summary_cols['bsp']].mean()) if summary_cols['bsp'] else "N/A",
                        format_percentage(filtered_sales[summary_cols['collection']].mean()) if summary_cols['collection'] else "N/A"
                    ]
                })
                
                csv_summary = summary_stats.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Summary Statistics",
                    data=csv_summary,
                    file_name=f"summary_stats_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )

    else:
        st.markdown("""
            <div style="text-align: center; padding: 2rem;">
                <h2>Welcome to the Real Estate Analytics Dashboard</h2>
                <p>Please upload Excel file(s) to begin analysis.</p>
                <p>The file(s) should contain the following sheets:</p>
                <ul style="list-style-type: none;">
                    <li>Main Collection AC P1_P2_P3</li>
                    <li>Annex - Sales Master</li>
                    <li>Un Sold Inventory</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
