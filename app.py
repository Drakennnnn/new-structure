import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
import re
import logging
import warnings

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

# Set page configuration
st.set_page_config(
    page_title="Real Estate Analytics Dashboard",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS with dark theme
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

def get_phase_sections(df):
    """Identify phase sections in collection account data"""
    phase_markers = []
    current_phase = None
    phases = {}
    
    for idx, row in df.iterrows():
        desc = str(row['Description']).strip()
        if 'Main Collection Escrow A/c Phase-' in desc:
            if current_phase:
                phases[current_phase]['end'] = idx - 1
            current_phase = desc
            phases[current_phase] = {'start': idx + 1}
            phase_markers.append(idx)
    
    # Set end for last phase
    if current_phase:
        phases[current_phase]['end'] = len(df)
    
    return phases

def process_collection_account(df):
    """Process collection account data"""
    df = df.copy()
    
    # Convert date columns
    df['Txn Date'] = pd.to_datetime(df['Txn Date'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce')
    
    # Process amounts
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    
    # Categorize transactions
    df['Payment Mode'] = df['Description'].apply(categorize_transaction)
    
    # Extract unit info from Sales Tag
    df['Unit'] = df['Sales Tag'].fillna('').str.strip()
    
    # Calculate running balance
    df['Running Balance'] = df.apply(lambda x: 
        x['Amount'] if x['Dr/Cr'] == 'C' else -x['Amount'], axis=1).cumsum()
    
    return df

def categorize_transaction(description):
    """Categorize payment modes"""
    description = str(description).upper()
    if any(word in description for word in ['CHQ', 'CHEQUE', 'MICR']):
        return 'Cheque'
    elif 'UPI' in description:
        return 'UPI'
    elif 'NEFT' in description:
        return 'NEFT'
    elif 'RTGS' in description:
        return 'RTGS'
    elif 'IMPS' in description:
        return 'IMPS'
    elif 'TRANSFER' in description:
        return 'Transfer'
    return 'Other'

def process_sales_master(df):
    """Process sales master data"""
    df = df.copy()
    
    # Convert dates
    date_columns = ['Booking date', 'Agreement date']
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Clean numeric columns - shorter column names
    numeric_columns = [
        'Area(sqft)', 'Basic Price', 'BSP/SqFt', 
        'Total Consideration', 'Amount Demanded',
        'Amount received', 'Balance receivables'
    ]
    
    for col in df.columns:
        for num_col in numeric_columns:
            if num_col in col:  # Match partial names
                df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Calculate collection percentage
    collection_col = [col for col in df.columns if 'Amount received' in col][0]
    demanded_col = [col for col in df.columns if 'Amount Demanded' in col][0]
    
    if collection_col and demanded_col:
        df['Collection Percentage'] = (
            df[collection_col] / df[demanded_col] * 100
        ).clip(0, 100)
    
    return df

def process_unsold_inventory(df):
    """Process unsold inventory data"""
    df = df.copy()
    
    # Clean area column
    if 'Saleable area (sq. ft.)' in df.columns:
        df['Saleable area (sq. ft.)'] = pd.to_numeric(df['Saleable area (sq. ft.)'], errors='coerce')
    
    return df

def format_currency(value):
    """Format currency values"""
    return f"₹{value:,.2f}"

def main():
    st.title("Real Estate Analytics Dashboard")
    st.markdown("---")

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
                sales_master_df = pd.read_excel(
                    excel_file,
                    'Annex - Sales Master',
                    skiprows=3
                )
                sales_master_df = process_sales_master(sales_master_df)

                # Process Unsold Inventory
                unsold_df = pd.read_excel(
                    excel_file,
                    'Un Sold Inventory'
                )
                unsold_df = process_unsold_inventory(unsold_df)

                # Store in session state
                st.session_state.data_loaded = True
                st.session_state.collection_df = collection_df
                st.session_state.sales_master_df = sales_master_df
                st.session_state.unsold_df = unsold_df
                
                # Get phase information
                st.session_state.phase_sections = get_phase_sections(collection_df)

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
                    delta=f"{unsold_units} Unsold"
                )

            with col2:
                consideration_col = [col for col in st.session_state.sales_master_df.columns if 'Total Consideration' in col][0]
                total_consideration = st.session_state.sales_master_df[consideration_col].sum()
                st.metric(
                    "Total Consideration",
                    format_currency(total_consideration),
                    delta="View Details"
                )

            with col3:
                total_collection = st.session_state.sales_master_df['Amount received as on 31.01.2025'].sum()
                collection_percentage = (total_collection / total_consideration * 100)
                st.metric(
                    "Collection Achievement",
                    f"{collection_percentage:.1f}%",
                    delta=format_currency(total_consideration - total_collection)
                )

            with col4:
                total_area = st.session_state.sales_master_df['Area(sqft) (A)'].sum()
                st.metric(
                    "Total Area",
                    f"{total_area:,.0f} sq.ft",
                    delta=f"{unsold_units/total_units*100:.1f}% unsold"
                )

            # Phase-wise collection summary
            st.markdown("### Phase-wise Collection Summary")
            
            # Get phase-wise totals
            phase_summary = []
            for phase, indices in st.session_state.phase_sections.items():
                phase_data = st.session_state.collection_df.iloc[
                    indices['start']:indices['end']
                ]
                credits = phase_data[phase_data['Dr/Cr'] == 'C']['Amount'].sum()
                debits = phase_data[phase_data['Dr/Cr'] == 'D']['Amount'].sum()
                
                phase_summary.append({
                    'Phase': phase,
                    'Credits': credits,
                    'Debits': debits,
                    'Net': credits - debits,
                    'Last Balance': phase_data['Running Balance'].iloc[-1]
                    if len(phase_data) > 0 else 0
                })
            
            phase_summary_df = pd.DataFrame(phase_summary)
            
            # Show phase summary
            col1, col2 = st.columns(2)
            
            with col1:
                fig_phase = px.bar(
                    phase_summary_df,
                    x='Phase',
                    y=['Credits', 'Debits'],
                    title="Collections by Phase",
                    barmode='group'
                )
                fig_phase.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
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
                    title="Collection by Payment Mode"
                )
                fig_payment.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
                )
                st.plotly_chart(fig_payment, use_container_width=True)
            
            with col2:
                # Format payment summary
                display_payment = payment_summary.copy()
                display_payment['Amount'] = display_payment['Amount'].apply(format_currency)
                display_payment['Percentage'] = (
                    payment_summary['Amount'] / payment_summary['Amount'].sum() * 100
                ).apply(lambda x: f"{x:.1f}%")
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
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='#ffffff'
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
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
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
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
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
                unit_type_filter = st.multiselect(
                    "Select Unit Types",
                    options=list(st.session_state.sales_master_df['Type of Unit \n(2BHK, 3BHK etc.)'].unique()),
                    default=list(st.session_state.sales_master_df['Type of Unit \n(2BHK, 3BHK etc.)'].unique())
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
                (st.session_state.sales_master_df['Type of Unit \n(2BHK, 3BHK etc.)'].isin(unit_type_filter)) &
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
                    color_continuous_scale='Blues'
                )
                fig_tower.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
                )
                st.plotly_chart(fig_tower, use_container_width=True)
            
            with col2:
                unit_type_dist = filtered_sales['Type of Unit \n(2BHK, 3BHK etc.)'].value_counts()
                
                fig_unit_type = px.pie(
                    values=unit_type_dist.values,
                    names=unit_type_dist.index,
                    title="Unit Type Distribution"
                )
                fig_unit_type.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
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
                    title="Collection Percentage Distribution"
                )
                fig_collection.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
                )
                st.plotly_chart(fig_collection, use_container_width=True)
            
            with col2:
                # BSP Analysis
                fig_bsp = px.box(
                    filtered_sales,
                    x='Tower',
                    y='BSP/SqFt',
                    color='Type of Unit \n(2BHK, 3BHK etc.)',
                    title="BSP Distribution by Tower and Unit Type"
                )
                fig_bsp.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
                )
                st.plotly_chart(fig_bsp, use_container_width=True)

        with tab4:
            st.markdown("### Unit Details")
            
            # Column selection
            all_columns = st.session_state.sales_master_df.columns.tolist()
            default_columns = [
                'Apartment No', 'Tower', 'Type of Unit \n(2BHK, 3BHK etc.)',
                'Area(sqft) (A)', 'BSP/SqFt', 'Total Consideration (F = C+D+E)',
                'Amount received as on 31.01.2025', 'Collection Percentage',
                'Balance receivables (on total sales)'
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
                'Area(sqft) (A)': lambda x: f"{x:,.0f}",
                'BSP/SqFt': lambda x: f"₹{x:,.2f}",
                'Total Consideration (F = C+D+E)': lambda x: f"₹{x:,.2f}",
                'Amount received as on 31.01.2025': lambda x: f"₹{x:,.2f}",
                'Collection Percentage': lambda x: f"{x:.1f}%",
                'Balance receivables (on total sales)': lambda x: f"₹{x:,.2f}"
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
                summary_stats = pd.DataFrame({
                    'Metric': ['Total Units', 'Total Area', 'Total Consideration', 
                             'Total Collection', 'Average BSP', 'Average Collection %'],
                    'Value': [
                        len(filtered_sales),
                        f"{filtered_sales['Area(sqft) (A)'].sum():,.0f} sq.ft",
                        f"₹{filtered_sales['Total Consideration (F = C+D+E)'].sum():,.2f}",
                        f"₹{filtered_sales['Amount received as on 31.01.2025'].sum():,.2f}",
                        f"₹{filtered_sales['BSP/SqFt'].mean():,.2f}",
                        f"{filtered_sales['Collection Percentage'].mean():.1f}%"
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
