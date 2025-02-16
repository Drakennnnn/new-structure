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

# Column patterns for matching
COLUMN_PATTERNS = {
    'unit_no': {
        'patterns': ['apartment no', 'unit no', 'flat no'],
        'transforms': ['trim', 'lowercase', 'remove_linebreaks']
    },
    'tower': {
        'patterns': ['tower', 'building', 'block'],
        'transforms': ['trim', 'lowercase']
    },
    'unit_type': {
        'patterns': ['type of unit', 'bhk', 'unit type'],
        'transforms': ['trim', 'lowercase', 'remove_linebreaks']
    },
    'area_sqft': {
        'patterns': ['area.*sqft', 'sqft.*area', 'square feet', 'saleable area'],
        'transforms': ['trim', 'lowercase', 'remove_parentheses']
    },
    'total_consideration': {
        'patterns': ['total consideration', 'consideration f', 'total amount'],
        'transforms': ['trim', 'lowercase', 'remove_formula']
    },
    'amount_demanded': {
        'patterns': ['amount demanded', 'demanded amount', 'demand raised'],
        'transforms': ['trim', 'lowercase', 'remove_date']
    },
    'amount_received': {
        'patterns': ['amount received', 'received amount', 'collection amount'],
        'transforms': ['trim', 'lowercase', 'remove_date']
    },
    'balance_receivable': {
        'patterns': ['balance receivable', 'receivable balance', 'pending amount'],
        'transforms': ['trim', 'lowercase', 'remove_date']
    }
}

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

class ColumnMatcher:
    """Handle column name matching across different Excel formats"""
    
    def __init__(self, patterns: Dict):
        self.patterns = patterns
        self.transforms = {
            'trim': lambda s: s.strip(),
            'lowercase': lambda s: s.lower(),
            'remove_linebreaks': lambda s: re.sub(r'[\r\n]+', ' ', s),
            'remove_parentheses': lambda s: re.sub(r'\([^)]*\)', '', s),
            'remove_date': lambda s: re.sub(r'\d{2}[-.\/]\d{2}[-.\/]\d{4}', '', s),
            'remove_formula': lambda s: re.sub(r'\([^)]*=[^)]*\)', '', s),
            'remove_spaces': lambda s: re.sub(r'\s+', '', s)
        }

    def normalize_header(self, header: str, transforms: List[str]) -> str:
        """Apply transformations to normalize header string"""
        result = str(header)
        for transform in transforms:
            result = self.transforms[transform](result)
        return result

    def find_matching_columns(self, df: pd.DataFrame) -> Dict[str, Optional[str]]:
        """Find matching columns in dataframe"""
        matches = {}
        headers = df.columns.tolist()
        
        for std_col, matcher in self.patterns.items():
            best_match = None
            best_score = 0
            
            for header in headers:
                if header in matches.values():
                    continue
                
                norm_header = self.normalize_header(header, matcher['transforms'])
                
                for pattern in matcher['patterns']:
                    if re.search(pattern, norm_header):
                        score = len(pattern)
                        if score > best_score:
                            best_score = score
                            best_match = header
            
            matches[std_col] = best_match
        
        return matches

class PhaseTracker:
    """Track and manage phase information in collection accounts"""
    
    def __init__(self):
        self.phase_pattern = r'Main Collection Escrow A/c Phase-\d+'
        self.account_pattern = r'\d{14}'

    def extract_phase_info(self, row: pd.Series) -> Optional[Dict]:
        """Extract phase and account information from row"""
        desc = str(row.get('Description', '')).strip()
        if re.search(self.phase_pattern, desc):
            account_match = re.search(self.account_pattern, str(row.get('Account No', '')))
            account_num = None
            # Look for account number in different columns if not found
            if not account_match:
                for col in row.index:
                    val = str(row.get(col, ''))
                    account_match = re.search(self.account_pattern, val)
                    if account_match:
                        account_num = account_match.group(0)
                        break
            else:
                account_num = account_match.group(0)
            return {
                'phase': desc,
                'account': account_num
            }
        return None

    def get_phase_sections(self, df: pd.DataFrame) -> Dict:
        """Identify phase sections in collection account data"""
        phases = {}
        current_phase = None
        
        for idx, row in df.iterrows():
            phase_info = self.extract_phase_info(row)
            if phase_info:
                if current_phase:
                    phases[current_phase['phase']]['end'] = idx - 1
                current_phase = phase_info
                phases[current_phase['phase']] = {
                    'start': idx + 1,
                    'account': current_phase['account']
                }
        
        if current_phase:
            phases[current_phase['phase']]['end'] = len(df)
        
        return phases

def process_collection_account(df: pd.DataFrame) -> pd.DataFrame:
    """Process collection account data"""
    df = df.copy()
    
    # Handle empty cells
    df = df.fillna({
        'Description': '',
        'Amount': 0,
        'Dr/Cr': '',
        'Sales Tag': ''
    })
    
    # Convert date columns with flexible formats
    date_formats = ['%d/%m/%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d']
    for date_col in ['Txn Date', 'Value Date']:
        if date_col in df.columns:
            for fmt in date_formats:
                try:
                    df[date_col] = pd.to_datetime(df[date_col], format=fmt, errors='coerce')
                    break
                except:
                    continue
    
    # Process amounts - handle different number formats
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].astype(str).replace('[\₹,]', '', regex=True)
        df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
    
    # Enhanced transaction categorization
    df['Payment Mode'] = df['Description'].apply(categorize_transaction)
    
    # Extract and clean unit info
    df['Unit'] = df['Sales Tag'].fillna('').str.strip()
    
    # Calculate running balance with phase reset
    df['Running Balance'] = df.apply(
        lambda x: x['Amount'] if x['Dr/Cr'] == 'C' else -x['Amount'], 
        axis=1
    ).fillna(0).cumsum()
    
    return df

def process_sales_master(df: pd.DataFrame, matcher: ColumnMatcher) -> Tuple[pd.DataFrame, Dict]:
    """Process sales master data with enhanced column matching"""
    df = df.copy()
    
    # Get column matches
    column_matches = matcher.find_matching_columns(df)
    
    # Convert dates
    date_patterns = ['date', 'booking']
    for col in df.columns:
        if any(pattern in col.lower() for pattern in date_patterns):
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Process numeric columns
    numeric_columns = [
        'area_sqft', 'total_consideration', 'amount_demanded',
        'amount_received', 'balance_receivable'
    ]
    for col_type in numeric_columns:
        if column_matches.get(col_type):
            col = column_matches[col_type]
            df[col] = pd.to_numeric(
                df[col].astype(str).replace('[\₹,]', '', regex=True), 
                errors='coerce'
            ).fillna(0)
    
    # Calculate collection percentage
    received_col = column_matches.get('amount_received')
    demanded_col = column_matches.get('amount_demanded')
    if received_col and demanded_col:
        df['Collection Percentage'] = (
            df[received_col] / df[demanded_col] * 100
        ).clip(0, 100)
    
    return df, column_matches

def process_unsold_inventory(df: pd.DataFrame, matcher: ColumnMatcher) -> pd.DataFrame:
    """Process unsold inventory with enhanced features"""
    df = df.copy()
    
    # Get column matches
    column_matches = matcher.find_matching_columns(df)
    
    # Clean area column
    area_col = column_matches.get('area_sqft')
    if area_col:
        df[area_col] = pd.to_numeric(
            df[area_col].astype(str).replace('[\₹,]', '', regex=True), 
            errors='coerce'
        ).fillna(0)
    
    # Add status if missing
    if 'Status' not in df.columns:
        df['Status'] = 'Unsold'
    
    return df

def categorize_transaction(description: str) -> str:
    """Enhanced transaction categorization"""
    description = str(description).upper()
    categories = {
        'Cheque': ['CHQ', 'CHEQUE', 'MICR'],
        'UPI': ['UPI'],
        'NEFT': ['NEFT'],
        'RTGS': ['RTGS'],
        'IMPS': ['IMPS'],
        'Transfer': ['TRANSFER', 'TRF'],
        'Cash': ['CASH'],
        'DD': ['DD', 'DEMAND DRAFT']
    }
    
    for category, keywords in categories.items():
        if any(keyword in description for keyword in keywords):
            return category
    return 'Other'

def format_currency(value: float) -> str:
    """Format currency values with INR symbol"""
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

def create_collection_summary(df: pd.DataFrame, phase_sections: Dict) -> pd.DataFrame:
    """Create summary of collections by phase"""
    summaries = []
    
    for phase, indices in phase_sections.items():
        phase_data = df.iloc[indices['start']:indices['end']]
        
        credits = phase_data[phase_data['Dr/Cr'] == 'C']['Amount'].sum()
        debits = phase_data[phase_data['Dr/Cr'] == 'D']['Amount'].sum()
        last_balance = phase_data['Running Balance'].iloc[-1] if len(phase_data) > 0 else 0
        
        summaries.append({
            'Phase': phase,
            'Credits': credits,
            'Debits': debits,
            'Net': credits - debits,
            'Last Balance': last_balance,
            'Account': phase_sections[phase].get('account', 'N/A')
        })
    
    return pd.DataFrame(summaries)

def main():
    """Main application function"""
    st.title("Real Estate Analytics Dashboard")
    st.markdown("---")

    # Initialize components
    column_matcher = ColumnMatcher(COLUMN_PATTERNS)
    phase_tracker = PhaseTracker()

    # Initialize session state
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
        st.session_state.collection_df = None
        st.session_state.sales_master_df = None
        st.session_state.unsold_df = None

    # File upload
    uploaded_file = st.file_uploader("Upload MIS Excel File", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            with st.spinner('Processing data...'):
                # Read Excel file
                excel_file = pd.ExcelFile(uploaded_file)
                
                # Verify required sheets
                required_sheets = ['Main Collection AC P1_P2_P3', 'Annex - Sales Master', 'Un Sold Inventory']
                missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
                if missing_sheets:
                    st.error(f"Missing required sheets: {', '.join(missing_sheets)}")
                    st.stop()

                # Process Collection Account
                collection_df = pd.read_excel(
                    excel_file, 
                    'Main Collection AC P1_P2_P3',
                    skiprows=1
                )
                collection_df = process_collection_account(collection_df)
                
                # Process Sales Master
                try:
                    sales_master_df = pd.read_excel(
                        excel_file,
                        'Annex - Sales Master',
                        skiprows=3
                    )
                    sales_master_df, column_matches = process_sales_master(sales_master_df, column_matcher)
                    st.session_state.column_matches = column_matches
                except Exception as e:
                    st.error(f"Error processing Sales Master sheet: {str(e)}")
                    sales_master_df = pd.DataFrame()
                    column_matches = {}

                # Process Unsold Inventory
                try:
                    unsold_df = pd.read_excel(
                        excel_file,
                        'Un Sold Inventory'
                    )
                    unsold_df = process_unsold_inventory(unsold_df, column_matcher)
                except Exception as e:
                    st.error(f"Error processing Unsold Inventory sheet: {str(e)}")
                    unsold_df = pd.DataFrame()

                # Get phase information using PhaseTracker
                phase_sections = phase_tracker.get_phase_sections(collection_df)

                # Store in session state
                st.session_state.data_loaded = True
                st.session_state.collection_df = collection_df
                st.session_state.sales_master_df = sales_master_df
                st.session_state.unsold_df = unsold_df
                st.session_state.phase_sections = phase_sections

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
                consideration_col = st.session_state.column_matches.get('total_consideration')
                if consideration_col:
                    total_consideration = st.session_state.sales_master_df[consideration_col].sum()
                    st.metric(
                        "Total Consideration",
                        format_currency(total_consideration),
                        delta="View Details"
                    )

            with col3:
                received_col = st.session_state.column_matches.get('amount_received')
                if received_col and consideration_col:
                    total_collection = st.session_state.sales_master_df[received_col].sum()
                    collection_percentage = (total_collection / total_consideration * 100)
                    st.metric(
                        "Collection Achievement",
                        format_percentage(collection_percentage),
                        delta=format_currency(total_consideration - total_collection),
                        delta_color="inverse"
                    )

            with col4:
                area_col = st.session_state.column_matches.get('area_sqft')
                if area_col:
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
            
            # Create phase summary using the summary function
            phase_summary_df = create_collection_summary(
                st.session_state.collection_df,
                st.session_state.phase_sections
            )
            
            # Show phase summary
            col1, col2 = st.columns(2)
            
            with col1:
                fig_phase = px.bar(
                    phase_summary_df,
                    x='Phase',
                    y=['Credits', 'Debits'],
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
                # Format phase summary for display
                display_summary = phase_summary_df.copy()
                for col in ['Credits', 'Debits', 'Net', 'Last Balance']:
                    display_summary[col] = display_summary[col].apply(format_currency)
                st.dataframe(display_summary, use_container_width=True)

            # Payment mode analysis
            st.markdown("### Payment Mode Analysis")
            
            col1, col2 = st.columns(2)
            
            with col1:
                payment_summary = st.session_state.collection_df[
                    st.session_state.collection_df['Dr/Cr'] == 'C'
                ].groupby('Payment Mode')['Amount'].agg(['sum', 'count']).reset_index()
                
                payment_summary.columns = ['Mode', 'Amount', 'Count']
                
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
            
            with col2:
                # Format payment summary
                display_payment = payment_summary.copy()
                display_payment['Amount'] = display_payment['Amount'].apply(format_currency)
                display_payment['Percentage'] = (
                    payment_summary['Amount'] / payment_summary['Amount'].sum() * 100
                ).apply(format_percentage)
                st.dataframe(display_payment, use_container_width=True)

            # Monthly collection trend
            st.markdown("### Collection Trends")
            
            monthly_collection = st.session_state.collection_df[
                st.session_state.collection_df['Dr/Cr'] == 'C'
            ].groupby(pd.Grouper(key='Value Date', freq='M'))['Amount'].sum().reset_index()
            
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

        with tab2:
            st.markdown("### Collection Analysis")
            
            # Filters
            col1, col2, col3 = st.columns(3)
            
            with col1:
                selected_phase = st.selectbox(
                    "Select Phase",
                    ["All Phases"] + list(st.session_state.phase_sections.keys())
                )
            
            with col2:
                date_range = st.date_input(
                    "Select Date Range",
                    value=(
                        st.session_state.collection_df['Value Date'].min(),
                        st.session_state.collection_df['Value Date'].max()
                    )
                )
            
            with col3:
                payment_modes = st.multiselect(
                    "Payment Modes",
                    options=st.session_state.collection_df['Payment Mode'].unique(),
                    default=st.session_state.collection_df['Payment Mode'].unique()
                )

            # Filter collection data
            filtered_collection = st.session_state.collection_df.copy()
            
            if selected_phase != "All Phases":
                phase_indices = st.session_state.phase_sections[selected_phase]
                filtered_collection = filtered_collection.iloc[
                    phase_indices['start']:phase_indices['end']
                ]
            
            filtered_collection = filtered_collection[
                (filtered_collection['Value Date'].dt.date >= date_range[0]) &
                (filtered_collection['Value Date'].dt.date <= date_range[1]) &
                (filtered_collection['Payment Mode'].isin(payment_modes))
            ]

            # Show filtered transactions
            st.markdown("### Transaction Details")
            
            # Column selection
            default_columns = ['Txn Date', 'Value Date', 'Description', 'Dr/Cr', 
                             'Amount', 'Payment Mode', 'Unit', 'Running Balance']
            
            selected_columns = st.multiselect(
                "Select columns to display",
                options=filtered_collection.columns,
                default=default_columns
            )

            # Format and display filtered data
            display_collection = filtered_collection[selected_columns].copy()
            if 'Amount' in selected_columns:
                display_collection['Amount'] = display_collection['Amount'].apply(format_currency)
            if 'Running Balance' in selected_columns:
                display_collection['Running Balance'] = display_collection['Running Balance'].apply(format_currency)
            
            st.dataframe(
                display_collection.sort_values('Txn Date', ascending=False),
                use_container_width=True
            )

            # Collection summary
            col1, col2 = st.columns(2)
            
            with col1:
                # Daily collection trend
                daily_collection = filtered_collection[
                    filtered_collection['Dr/Cr'] == 'C'
                ].groupby('Value Date')['Amount'].sum().reset_index()
                
                fig_daily = px.line(
                    daily_collection,
                    x='Value Date',
                    y='Amount',
                    title="Daily Collection Trend"
                )
                fig_daily.update_layout(
                    plot_bgcolor=COLORS['background'],
                    paper_bgcolor=COLORS['background'],
                    font_color=COLORS['text']
                )
                st.plotly_chart(fig_daily, use_container_width=True)
            
            with col2:
                # Running balance trend
                fig_balance = px.line(
                    filtered_collection,
                    x='Value Date',
                    y='Running Balance',
                    title="Running Balance Trend"
                )
                fig_balance.update_layout(
                    plot_bgcolor=COLORS['background'],
                    paper_bgcolor=COLORS['background'],
                    font_color=COLORS['text']
                )
                st.plotly_chart(fig_balance, use_container_width=True)

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
                unit_type_col = st.session_state.column_matches.get('unit_type', 'Type of Unit \n(2BHK, 3BHK etc.)')
                unit_type_filter = st.multiselect(
                    "Select Unit Types",
                    options=list(st.session_state.sales_master_df[unit_type_col].unique()),
                    default=list(st.session_state.sales_master_df[unit_type_col].unique())
                )
            
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
                (st.session_state.sales_master_df['Tower'].isin(tower_filter)) &
                (st.session_state.sales_master_df[unit_type_col].isin(unit_type_filter)) &
                (st.session_state.sales_master_df['Collection Percentage'] >= collection_range[0]) &
                (st.session_state.sales_master_df['Collection Percentage'] <= collection_range[1])
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
                bsp_col = st.session_state.column_matches.get('bsp_sqft', 'BSP/SqFt')
                if bsp_col in filtered_sales.columns:
                    fig_bsp = px.box(
                        filtered_sales,
                        x='Tower',
                        y=bsp_col,
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
                'Apartment No', 'Tower', unit_type_col,
                st.session_state.column_matches.get('area_sqft', 'Area(sqft) (A)'),
                bsp_col,
                st.session_state.column_matches.get('total_consideration', 'Total Consideration (F = C+D+E)'),
                st.session_state.column_matches.get('amount_received', 'Amount received'),
                'Collection Percentage',
                st.session_state.column_matches.get('balance_receivable', 'Balance receivables')
            ]
            
            selected_columns = st.multiselect(
                "Select columns to display",
                options=all_columns,
                default=[col for col in default_columns if col in all_columns]
            )

            # Format and display data
            display_sales = filtered_sales[selected_columns].copy()
            
            # Format numeric columns
            numeric_formats = {
                st.session_state.column_matches.get('area_sqft', 'Area(sqft) (A)'): lambda x: f"{x:,.0f}",
                bsp_col: lambda x: f"₹{x:,.2f}",
                st.session_state.column_matches.get('total_consideration', 'Total Consideration'): format_currency,
                st.session_state.column_matches.get('amount_received', 'Amount received'): format_currency,
                'Collection Percentage': format_percentage,
                st.session_state.column_matches.get('balance_receivable', 'Balance receivables'): format_currency
            }
            
            for col in selected_columns:
                if col in numeric_formats:
                    display_sales[col] = display_sales[col].apply(numeric_formats[col])

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
                area_col = st.session_state.column_matches.get('area_sqft', 'Area(sqft) (A)')
                consideration_col = st.session_state.column_matches.get('total_consideration', 'Total Consideration')
                received_col = st.session_state.column_matches.get('amount_received', 'Amount received')
                
                summary_stats = pd.DataFrame({
                    'Metric': ['Total Units', 'Total Area', 'Total Consideration', 
                             'Total Collection', 'Average BSP', 'Average Collection %'],
                    'Value': [
                        len(filtered_sales),
                        f"{filtered_sales[area_col].sum():,.0f} sq.ft",
                        format_currency(filtered_sales[consideration_col].sum()),
                        format_currency(filtered_sales[received_col].sum()),
                        format_currency(filtered_sales[bsp_col].mean()),
                        format_percentage(filtered_sales['Collection Percentage'].mean())
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
