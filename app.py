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
    page_title="Real Estate Project Management Dashboard",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS with dark theme (keeping the same styling)
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

def process_collection_account(df, phase_row_indices):
    """Process collection account data with phase information"""
    df = df.copy()
    
    # Convert date columns
    df['Txn Date'] = pd.to_datetime(df['Txn Date'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce')
    
    # Add phase information
    df['Phase'] = 'Unknown'
    current_phase = None
    for idx, row in df.iterrows():
        if idx in phase_row_indices:
            current_phase = row['Description']
        df.at[idx, 'Phase'] = current_phase
    
    # Process amounts
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    df['Running Total'] = pd.to_numeric(df['Running Total'], errors='coerce')
    
    # Create transaction type
    df['Transaction Type'] = df.apply(lambda x: categorize_transaction(x['Description']), axis=1)
    
    # Extract unit info from Sales Tag
    df['Unit_Info'] = df['Sales Tag'].fillna('')
    
    return df

def categorize_transaction(description):
    """Categorize transactions based on description"""
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
    else:
        return 'Other'

def process_sales_master(df):
    """Process sales master data"""
    df = df.copy()
    
    # Convert dates
    df['Booking date'] = pd.to_datetime(df['Booking date'], errors='coerce')
    df['Agreement date'] = pd.to_datetime(df['Agreement date'], errors='coerce')
    
    # Clean numeric columns
    numeric_columns = ['Area(sqft) (A)', 'Basic Price', 'BSP/SqFt', 'Total Consideration (F = C+D+E)',
                      'Amount Demanded', 'Amount received as on 31.01.2025']
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Calculate collection percentage
    df['Collection Percentage'] = (df['Amount received as on 31.01.2025'] / 
                                 df['Amount Demanded'] * 100).clip(0, 100)
    
    return df

def process_unsold_inventory(df):
    """Process unsold inventory data"""
    df = df.copy()
    
    # Clean area column
    df['Saleable area (sq. ft.)'] = pd.to_numeric(df['Saleable area (sq. ft.)'], errors='coerce')
    
    return df

def get_phase_row_indices(df):
    """Get row indices where phase information starts"""
    phase_rows = []
    for idx, row in df.iterrows():
        if isinstance(row['Description'], str) and 'Phase' in row['Description']:
            phase_rows.append(idx)
    return phase_rows

def main():
    st.title("Real Estate Project Management Dashboard")
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
                collection_df = pd.read_excel(excel_file, 'Main Collection AC P1_P2_P3', skiprows=1)
                phase_rows = get_phase_row_indices(collection_df)
                collection_df = process_collection_account(collection_df, phase_rows)

                # Process Sales Master
                sales_master_df = pd.read_excel(excel_file, 'Annex - Sales Master', skiprows=3)
                sales_master_df = process_sales_master(sales_master_df)

                # Process Unsold Inventory
                unsold_df = pd.read_excel(excel_file, 'Un Sold Inventory')
                unsold_df = process_unsold_inventory(unsold_df)

                # Store in session state
                st.session_state.data_loaded = True
                st.session_state.collection_df = collection_df
                st.session_state.sales_master_df = sales_master_df
                st.session_state.unsold_df = unsold_df

                st.success("Data loaded successfully!")

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.stop()

    if st.session_state.data_loaded:
        # Top level tabs
        tab1, tab2, tab3 = st.tabs(["Project Overview", "Collection Analysis", "Sales Status"])

        with tab1:
            st.markdown("### Project Overview")
            
            # Summary metrics
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                total_collection = st.session_state.collection_df[
                    st.session_state.collection_df['Dr/Cr'] == 'C'
                ]['Amount'].sum()
                st.metric(
                    "Total Collections",
                    f"₹{total_collection:,.2f}",
                    delta="View Details"
                )

            with col2:
                total_units = len(st.session_state.sales_master_df)
                unsold_units = len(st.session_state.unsold_df)
                st.metric(
                    "Total Units Sold",
                    f"{total_units}",
                    delta=f"{unsold_units} Unsold"
                )

            with col3:
                avg_collection = st.session_state.sales_master_df['Collection Percentage'].mean()
                st.metric(
                    "Average Collection %",
                    f"{avg_collection:.1f}%",
                    delta="View Details"
                )

            with col4:
                recent_transactions = len(st.session_state.collection_df[
                    st.session_state.collection_df['Txn Date'] > 
                    (datetime.now() - timedelta(days=30))
                ])
                st.metric(
                    "Recent Transactions",
                    f"{recent_transactions}",
                    delta="Last 30 days"
                )

            # Phase-wise collection summary
            st.markdown("### Phase-wise Collection Summary")
            phase_summary = st.session_state.collection_df[
                st.session_state.collection_df['Dr/Cr'] == 'C'
            ].groupby('Phase')['Amount'].sum().reset_index()
            
            fig_phase = px.bar(
                phase_summary,
                x='Phase',
                y='Amount',
                title="Collections by Phase",
                color='Amount',
                color_continuous_scale='Blues'
            )
            fig_phase.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font_color='#ffffff'
            )
            st.plotly_chart(fig_phase, use_container_width=True)

            # Transaction Type Analysis
            col1, col2 = st.columns(2)
            
            with col1:
                transaction_summary = st.session_state.collection_df[
                    st.session_state.collection_df['Dr/Cr'] == 'C'
                ].groupby('Transaction Type')['Amount'].sum().reset_index()
                
                fig_trans = px.pie(
                    transaction_summary,
                    values='Amount',
                    names='Transaction Type',
                    title="Collection by Payment Mode"
                )
                fig_trans.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font_color='#ffffff'
                )
                st.plotly_chart(fig_trans, use_container_width=True)

            with col2:
                daily_collection = st.session_state.collection_df[
                    st.session_state.collection_df['Dr/Cr'] == 'C'
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

        with tab2:
            st.markdown("### Collection Analysis")
            
            # Filters
            col1, col2, col3 = st.columns(3)
            
            with col1:
                selected_phase = st.selectbox(
                    "Select Phase",
                    ["All Phases"] + list(st.session_state.collection_df['Phase'].unique())
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
                transaction_types = st.multiselect(
                    "Transaction Types",
                    options=st.session_state.collection_df['Transaction Type'].unique(),
                    default=st.session_state.collection_df['Transaction Type'].unique()
                )

            # Filter data
            filtered_collection = st.session_state.collection_df.copy()
            if selected_phase != "All Phases":
                filtered_collection = filtered_collection[
                    filtered_collection['Phase'] == selected_phase
                ]
            
            filtered_collection = filtered_collection[
                (filtered_collection['Value Date'].dt.date >= date_range[0]) &
                (filtered_collection['Value Date'].dt.date <= date_range[1]) &
                (filtered_collection['Transaction Type'].isin(transaction_types))
            ]

            # Show filtered transactions
            st.markdown("### Transaction Details")
            
            # Column selection
            default_columns = ['Txn Date', 'Value Date', 'Description', 'Dr/Cr', 
                             'Amount', 'Transaction Type', 'Sales Tag', 'Phase']
            
            selected_columns = st.multiselect(
                "Select columns to display",
                options=filtered_collection.columns,
                default=default_columns
            )

            # Display filtered data
            st.dataframe(
                filtered_collection[selected_columns].sort_values('Txn Date', ascending=False),
                use_container_width=True
            )

            # Export option
            col1, col2 = st.columns(2)
            with col1:
                csv = filtered_collection[selected_columns].to_csv(index=False , encoding='utf-8')
                st.download_button(
                    label="Download Filtered Transactions",
                    data=csv,
                    file_name=f"collection_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with col2:
                # Summary statistics
                summary_stats = pd.DataFrame({
                    'Metric': ['Total Transactions', 'Total Amount', 'Average Amount', 
                             'Credit Transactions', 'Debit Transactions'],
                    'Value': [
                        len(filtered_collection),
                        f"₹{filtered_collection['Amount'].sum():,.2f}",
                        f"₹{filtered_collection['Amount'].mean():,.2f}",
                        len(filtered_collection[filtered_collection['Dr/Cr'] == 'C']),
                        len(filtered_collection[filtered_collection['Dr/Cr'] == 'D'])
                    ]
                })
                st.write(summary_stats)

        with tab3:
            st.markdown("### Sales Status")

            # Filters
            col1, col2 = st.columns(2)
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

            # Filter sales data
            filtered_sales = st.session_state.sales_master_df[
                (st.session_state.sales_master_df['Tower'].isin(tower_filter)) &
                (st.session_state.sales_master_df['Type of Unit \n(2BHK, 3BHK etc.)'].isin(unit_type_filter))
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

            # Detailed unit information
            st.markdown("### Unit Details")
            
            # Column selection for unit details
            default_unit_columns = ['Apartment No', 'Tower', 'Type of Unit \n(2BHK, 3BHK etc.)', 
                                  'Area(sqft) (A)', 'Total Consideration (F = C+D+E)', 
                                  'Amount received as on 31.01.2025', 'Collection Percentage']
            
            selected_unit_columns = st.multiselect(
                "Select columns to display",
                options=filtered_sales.columns,
                default=default_unit_columns
            )

            # Display filtered unit data
            formatted_sales = filtered_sales[selected_unit_columns].copy()
            
            # Format numeric columns
            numeric_cols = ['Area(sqft) (A)', 'Total Consideration (F = C+D+E)', 
                          'Amount received as on 31.01.2025']
            for col in numeric_cols:
                if col in selected_unit_columns:
                    formatted_sales[col] = formatted_sales[col].apply(lambda x: f"₹{x:,.2f}" if 'Consideration' in col or 'Amount' in col else f"{x:,.0f}")
            
            if 'Collection Percentage' in selected_unit_columns:
                formatted_sales['Collection Percentage'] = formatted_sales['Collection Percentage'].apply(lambda x: f"{x:.1f}%")

            st.dataframe(
                formatted_sales.sort_values('Apartment No'),
                use_container_width=True
            )

            # Export unit details
            col1, col2 = st.columns(2)
            with col1:
                csv_units = filtered_sales[selected_unit_columns].to_csv(index=False, encoding='utf-8')
                st.download_button(
                    label="Download Unit Details",
                    data=csv_units,
                    file_name=f"unit_details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )

            with col2:
                # Unit summary statistics
                unit_summary = pd.DataFrame({
                    'Metric': ['Total Units', 'Total Area', 'Total Value', 
                             'Total Collection', 'Average Collection %'],
                    'Value': [
                        len(filtered_sales),
                        f"{filtered_sales['Area(sqft) (A)'].sum():,.0f} sq.ft",
                        f"₹{filtered_sales['Total Consideration (F = C+D+E)'].sum():,.2f}",
                        f"₹{filtered_sales['Amount received as on 31.01.2025'].sum():,.2f}",
                        f"{filtered_sales['Collection Percentage'].mean():.1f}%"
                    ]
                })
                st.write(unit_summary)

            # Show unsold inventory
            st.markdown("### Unsold Inventory")
            st.dataframe(
                st.session_state.unsold_df,
                use_container_width=True
            )

    else:
        st.markdown("""
            <div style="text-align: center; padding: 2rem;">
                <h2>Welcome to the Real Estate Project Management Dashboard</h2>
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
