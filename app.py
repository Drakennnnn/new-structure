import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
import re
import logging
import warnings
from typing import Dict, List, Optional, Tuple

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

class PhaseProcessor:
    """Handle phase-wise data processing"""
    
    def __init__(self):
        self.phase_pattern = r'Main Collection Escrow A/c Phase-\d+'
        self.account_pattern = r'\d{14}'
        self.default_columns = ['Description', 'Dr/Cr', 'Amount', 'Value Date', 'Sales Tag']

    def find_phase_markers(self, df: pd.DataFrame) -> Dict:
        """Find phase boundaries and account numbers"""
        phases = {}
        current_phase = None
        
        for idx, row in df.iterrows():
            # Check for phase header in entire row
            for col in df.columns:
                cell_value = str(row[col])
                phase_match = re.search(r'Phase-(\d+)', cell_value)
                
                if phase_match and 'Main Collection Escrow A/c' in cell_value:
                    # Found a phase marker
                    if current_phase:
                        phases[current_phase]['end'] = idx - 1
                    
                    phase_num = phase_match.group(1)
                    current_phase = f"Phase-{phase_num}"
                    
                    # Look for account number in this row
                    account_no = None
                    for col2 in df.columns:
                        val = str(row[col2])
                        if re.match(r'^\d{14}$', val.strip()):
                            account_no = val.strip()
                            break
                    
                    phases[current_phase] = {
                        'start': idx + 1,
                        'account': account_no,
                        'phase_num': int(phase_num)
                    }
                    break
        
        # Set end for last phase
        if current_phase:
            phases[current_phase]['end'] = len(df)
        
        return phases

    def process_phase_data(self, df: pd.DataFrame, phase_info: Dict) -> pd.DataFrame:
        """Process data for a specific phase"""
        phase_data = df.iloc[phase_info['start']:phase_info['end']].copy()
        
        # Calculate running balance within phase
        phase_data['Amount'] = pd.to_numeric(
            phase_data['Amount'].astype(str).str.replace(',', ''), 
            errors='coerce'
        ).fillna(0)
        
        phase_data['Running Balance'] = phase_data.apply(
            lambda x: x['Amount'] if x['Dr/Cr'] == 'C' else -x['Amount'],
            axis=1
        ).cumsum()
        
        # Add transaction flags
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
    """Handle transaction processing and categorization"""
    
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

    def process_transaction(self, row: pd.Series) -> Dict:
        """Process a single transaction row"""
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
            'unit_no': self.extract_unit_info(row.get('Sales Tag', ''))
        }

class DataProcessor:
    """Main data processing class"""
    
    def __init__(self):
        self.phase_processor = PhaseProcessor()
        self.transaction_processor = TransactionProcessor()
        
    def process_dates(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process date columns with multiple format handling"""
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
                        break
                    except:
                        continue
        
        return df

    def process_amounts(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process amount columns"""
        amount_columns = [
            'Amount', 'Running Balance', 'Total Consideration',
            'Amount Demanded', 'Amount received', 'Balance receivables'
        ]
        
        for col in df.columns:
            if any(amount_text in col for amount_text in amount_columns):
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(',', ''),
                    errors='coerce'
                ).fillna(0)
        
        return df

    def process_collection_account(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict]:
        """Process collection account data"""
        df = df.copy()
        
        # Find phase information
        phases = self.phase_processor.find_phase_markers(df)
        
        # Process dates and amounts
        df = self.process_dates(df)
        df = self.process_amounts(df)
        
        # Add payment mode
        df['Payment Mode'] = df['Description'].apply(
            self.transaction_processor.categorize_transaction
        )
        
        # Process each phase
        processed_phases = {}
        for phase, info in phases.items():
            processed_phases[phase] = self.phase_processor.process_phase_data(df, info)
        
        # Combine processed data
        processed_df = pd.concat(processed_phases.values())
        
        return processed_df, phases

    def process_sales_master(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process sales master data"""
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
        """Process unsold inventory data"""
        df = df.copy()
        
        # Clean area column
        area_cols = [col for col in df.columns if 'area' in col.lower()]
        if area_cols:
            df[area_cols[0]] = pd.to_numeric(
                df[area_cols[0]].astype(str).str.replace(',', ''),
                errors='coerce'
            ).fillna(0)
        
        # Add status if missing
        if 'Status' not in df.columns:
            df['Status'] = 'Unsold'
        
        return df

def format_currency(value: float) -> str:
    """Format currency values"""
    try:
        return f"₹{float(value):,.2f}"
    except:
        return "₹0.00"

def format_percentage(value: float) -> str:
    """Format percentage values"""
    try:
        return f"{float(value):.1f}%"
    except:
        return "0.0%"

def create_phase_summary(df: pd.DataFrame, phases: Dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Create phase-wise summary with proper error handling"""
    summaries = []
    
    try:
        for phase, info in phases.items():
            try:
                phase_data = df.iloc[info['start']:info['end']]
                
                credits = phase_data[
                    phase_data['Dr/Cr'].str.strip().str.upper() == 'C'
                ]['Amount'].sum()
                
                debits = phase_data[
                    phase_data['Dr/Cr'].str.strip().str.upper() == 'D'
                ]['Amount'].sum()
                
                last_balance = phase_data['Running Balance'].iloc[-1] if len(phase_data) > 0 else 0
                
                summaries.append({
                    'Phase': phase,
                    'Credits': float(credits),
                    'Debits': float(debits),
                    'Net': float(credits - debits),
                    'Last_Balance': float(last_balance),
                    'Account': info.get('account', 'N/A')
                })
            except Exception as e:
                st.error(f"Error processing phase {phase}: {str(e)}")
                continue
        
        # Create DataFrame for plotting
        summary_df = pd.DataFrame(summaries)
        
        # Melt for visualization
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
    """Create payment mode summary"""
    if df.empty:
        return pd.DataFrame(columns=['Mode', 'Amount', 'Count', 'Percentage'])
        
    payment_summary = df[
        df['Dr/Cr'] == 'C'
    ].groupby('Payment Mode').agg({
        'Amount': ['sum', 'count']
    }).reset_index()
    
    payment_summary.columns = ['Mode', 'Amount', 'Count']
    payment_summary['Percentage'] = (
        payment_summary['Amount'] / payment_summary['Amount'].sum() * 100
    )
    
    return payment_summary

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

def filter_collection_data(df: pd.DataFrame, phase_info: Optional[Dict] = None,
                         start_date: Optional[datetime] = None,
                         end_date: Optional[datetime] = None,
                         payment_modes: Optional[List] = None) -> pd.DataFrame:
    """Filter collection data based on criteria"""
    filtered_df = df.copy()
    
    if phase_info:
        filtered_df = filtered_df.iloc[phase_info['start']:phase_info['end']]
    
    if start_date and end_date:
        filtered_df = filtered_df[
            (filtered_df['Value Date'].dt.date >= start_date) &
            (filtered_df['Value Date'].dt.date <= end_date)
        ]
    
    if payment_modes:
        filtered_df = filtered_df[filtered_df['Payment Mode'].isin(payment_modes)]
    
    return filtered_df

def main():
    st.title("Real Estate Analytics Dashboard")
    st.markdown("---")

    # Initialize data processor
    processor = DataProcessor()

    # Initialize session state
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
        st.session_state.collection_df = None
        st.session_state.sales_master_df = None
        st.session_state.unsold_df = None
        st.session_state.phases = None

    # File upload
    uploaded_file = st.file_uploader("Upload MIS Excel File", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            with st.spinner('Processing data...'):
                # Read Excel file
                excel_file = pd.ExcelFile(uploaded_file)
                
                # Verify required sheets
                required_sheets = [
                    'Main Collection AC P1_P2_P3',
                    'Annex - Sales Master',
                    'Un Sold Inventory'
                ]
                missing_sheets = [
                    sheet for sheet in required_sheets 
                    if sheet not in excel_file.sheet_names
                ]
                if missing_sheets:
                    st.error(f"Missing required sheets: {', '.join(missing_sheets)}")
                    st.stop()

                # Process Collection Account
                collection_df = pd.read_excel(
                    excel_file, 
                    'Main Collection AC P1_P2_P3',
                    skiprows=1
                )
                collection_df, phases = processor.process_collection_account(collection_df)

                # Process Sales Master
                try:
                    sales_master_df = pd.read_excel(
                        excel_file,
                        'Annex - Sales Master',
                        skiprows=3
                    )
                    sales_master_df = processor.process_sales_master(sales_master_df)
                except Exception as e:
                    st.error(f"Error processing Sales Master sheet: {str(e)}")
                    sales_master_df = pd.DataFrame()

                # Process Unsold Inventory
                try:
                    unsold_df = pd.read_excel(
                        excel_file,
                        'Un Sold Inventory'
                    )
                    unsold_df = processor.process_unsold_inventory(unsold_df)
                except Exception as e:
                    st.error(f"Error processing Unsold Inventory sheet: {str(e)}")
                    unsold_df = pd.DataFrame()

                # Store in session state
                st.session_state.data_loaded = True
                st.session_state.collection_df = collection_df
                st.session_state.sales_master_df = sales_master_df
                st.session_state.unsold_df = unsold_df
                st.session_state.phases = phases

                st.success("Data loaded successfully!")

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.stop()

    if st.session_state.data_loaded:
        # Main tabs
        tab1, tab2, tab3, tab4 = st.tabs([
            "Project Overview",
            "Collection Analysis",
            "Sales Status",
            "Unit Details"
        ])

        with tab1:
            st.markdown("### Project Overview")
            
            # Phase selection for filtering
            selected_phase = st.selectbox(
                "Select Phase",
                ["All Phases"] + list(st.session_state.phases.keys())
            )
            
            # Filter data based on phase
            if selected_phase != "All Phases":
                phase_info = st.session_state.phases[selected_phase]
                filtered_df = st.session_state.collection_df.iloc[
                    phase_info['start']:phase_info['end']
                ]
                st.info(f"Showing data for {selected_phase} (Account: {phase_info['account']})")
            else:
                filtered_df = st.session_state.collection_df
                phase_info = None

            # Summary metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                total_units = len(st.session_state.sales_master_df)
                unsold_units = len(st.session_state.unsold_df)
                st.metric(
                    "Total Units",
                    f"{total_units}",
                    delta=f"{unsold_units} Unsold",
                    delta_color="inverse"
                )

            with col2:
                consideration_cols = [col for col in st.session_state.sales_master_df.columns 
                                   if 'consideration' in col.lower()]
                if consideration_cols:
                    total_consideration = st.session_state.sales_master_df[consideration_cols[0]].sum()
                    st.metric(
                        "Total Consideration",
                        format_currency(total_consideration),
                        delta="View Details"
                    )

            with col3:
                if not st.session_state.sales_master_df.empty:
                    received_cols = [col for col in st.session_state.sales_master_df.columns 
                                   if 'received' in col.lower()]
                    demanded_cols = [col for col in st.session_state.sales_master_df.columns 
                                   if 'demanded' in col.lower()]
                    
                    if received_cols and demanded_cols:
                        total_collection = st.session_state.sales_master_df[received_cols[0]].sum()
                        total_demand = st.session_state.sales_master_df[demanded_cols[0]].sum()
                        collection_percentage = (total_collection / total_demand * 100)
                        st.metric(
                            "Collection Achievement",
                            format_percentage(collection_percentage),
                            delta=format_currency(total_demand - total_collection),
                            delta_color="inverse"
                        )

            with col4:
                area_cols = [col for col in st.session_state.sales_master_df.columns 
                           if 'area' in col.lower()]
                if area_cols:
                    area_col = area_cols[0]
                    total_area = st.session_state.sales_master_df[area_col].sum()
                    unsold_area = st.session_state.unsold_df[area_col].sum() if area_col in st.session_state.unsold_df else 0
                    st.metric(
                        "Total Area",
                        f"{total_area:,.0f} sq.ft",
                        delta=f"{unsold_area:,.0f} sq.ft unsold",
                        delta_color="inverse"
                    )

            # Phase-wise collection summary
            st.markdown("### Phase-wise Collection Summary")
            
            plot_df, summary_df = create_phase_summary(
                st.session_state.collection_df,
                st.session_state.phases
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
                    # Display summary table
                    display_summary = summary_df.copy()
                    for col in ['Credits', 'Debits', 'Net', 'Last_Balance']:
                        if col in display_summary.columns:
                            display_summary[col] = display_summary[col].apply(format_currency)
                    st.dataframe(display_summary, use_container_width=True)
            else:
                st.warning("No phase data available for visualization")

            # Payment mode analysis
            st.markdown("### Payment Mode Analysis")
            
            # Add payment mode filter
            payment_modes = st.multiselect(
                "Select Payment Modes",
                options=filtered_df['Payment Mode'].unique(),
                default=filtered_df['Payment Mode'].unique()
            )
            
            # Filter by payment modes
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
                
                # Filter by date range
                filtered_df = filter_collection_data(
                    filtered_df,
                    phase_info=phase_info if selected_phase != "All Phases" else None,
                    start_date=start_date,
                    end_date=end_date,
                    payment_modes=payment_modes
                )
                
                if not filtered_df.empty:
                    # Monthly trend
                    monthly_collection = filtered_df[
                        filtered_df['Dr/Cr'] == 'C'
                    ].groupby(
                        filtered_df['Value Date'].dt.to_period('M')
                    )['Amount'].sum().reset_index()
                    
                    monthly_collection['Value Date'] = monthly_collection['Value Date'].dt.to_timestamp()
                    
                    fig_monthly = px.line(
                        monthly_collection,
                        x='Value Date',
                        y='Amount',
                        title="Monthly Collection Trend"
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
                    
                    # Format display data
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
            
            # Phase filter
            selected_phase = st.selectbox(
                "Select Phase",
                ["All Phases"] + list(st.session_state.phases.keys()),
                key="collection_phase"
            )
            
            # Filter data
            if selected_phase != "All Phases":
                phase_info = st.session_state.phases[selected_phase]
                collection_data = st.session_state.collection_df.iloc[
                    phase_info['start']:phase_info['end']
                ]
                st.info(f"Analyzing {selected_phase} (Account: {phase_info['account']})")
            else:
                collection_data = st.session_state.collection_df.copy()
            
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

            # Collection trends
            st.markdown("### Collection Trends")
            
            try:
                monthly_trend = collection_data[
                    collection_data['Dr/Cr'] == 'C'
                ].groupby(
                    collection_data['Value Date'].dt.to_period('M')
                )['Amount'].sum().reset_index()
                
                monthly_trend['Value Date'] = monthly_trend['Value Date'].dt.to_timestamp()
                
                fig_trend = px.line(
                    monthly_trend,
                    x='Value Date',
                    y='Amount',
                    title="Monthly Collection Trend"
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
                    options=list(st.session_state.sales_master_df['Tower'].unique()),
                    default=list(st.session_state.sales_master_df['Tower'].unique())
                )
            
            with col2:
                unit_type_cols = [col for col in st.session_state.sales_master_df.columns 
                                if 'type' in col.lower() and 'unit' in col.lower()]
                if unit_type_cols:
                    unit_type_col = unit_type_cols[0]
                    unit_type_filter = st.multiselect(
                        "Select Unit Types",
                        options=list(st.session_state.sales_master_df[unit_type_col].unique()),
                        default=list(st.session_state.sales_master_df[unit_type_col].unique())
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
            filtered_sales = st.session_state.sales_master_df[
                (st.session_state.sales_master_df['Tower'].isin(tower_filter))
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
            all_columns = st.session_state.sales_master_df.columns.tolist()
            default_columns = [
                'Apartment No', 'Tower', unit_type_col if unit_type_cols else None,
                area_cols[0] if area_cols else None,
                bsp_cols[0] if bsp_cols else None,
                consideration_cols[0] if consideration_cols else None,
                received_cols[0] if received_cols else None,
                'Collection Percentage',
                demanded_cols[0] if demanded_cols else None
            ]
            
            default_columns = [col for col in default_columns if col is not None]
            
            selected_columns = st.multiselect(
                "Select columns to display",
                options=all_columns,
                default=[col for col in default_columns if col in all_columns]
            )

            # Format and display data
            display_sales = filtered_sales[selected_columns].copy()
            
            # Format numeric columns
            numeric_formats = {
                'Area': lambda x: f"{x:,.0f}",
                'BSP': lambda x: f"₹{x:,.2f}",
                'Consideration': lambda x: f"₹{x:,.2f}",
                'Amount': lambda x: f"₹{x:,.2f}",
                'Collection': lambda x: f"{x:.1f}%",
                'Balance': lambda x: f"₹{x:,.2f}"
            }
            
            for col in selected_columns:
                for key, formatter in numeric_formats.items():
                    if key.lower() in col.lower():
                        display_sales[col] = display_sales[col].apply(formatter)

            st.dataframe(
                display_sales.sort_values('Apartment No'),
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
                summary_stats = pd.DataFrame({
                    'Metric': ['Total Units', 'Total Area', 'Total Consideration', 
                             'Total Collection', 'Average BSP', 'Average Collection %'],
                    'Value': [
                        len(filtered_sales),
                        f"{filtered_sales[area_cols[0]].sum():,.0f} sq.ft" if area_cols else "N/A",
                        format_currency(filtered_sales[consideration_cols[0]].sum()) if consideration_cols else "N/A",
                        format_currency(filtered_sales[received_cols[0]].sum()) if received_cols else "N/A",
                        format_currency(filtered_sales[bsp_cols[0]].mean()) if bsp_cols else "N/A",
                        format_percentage(filtered_sales['Collection Percentage'].mean()) if 'Collection Percentage' in filtered_sales else "N/A"
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
                <p>Please upload an Excel file to begin analysis.</p>
                <p>The file should contain the following sheets:</p>
                <ul style="list-style-type: none;">
                    <li>Main Collection AC P1_P2_P3</li>
                    <li>Annex - Sales Master</li>
                    <li>Un Sold Inventory</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
