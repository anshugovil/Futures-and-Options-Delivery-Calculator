"""
Streamlit Futures Delivery Calculator
Web application for calculating physical delivery from futures/options positions
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import tempfile
import os
import logging
from typing import Dict, List, Optional

# These might cause issues - add them one by one
try:
    import yfinance as yf
except ImportError:
    st.error("Please install yfinance: pip install yfinance")
    st.stop()

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    st.error("Please install openpyxl: pip install openpyxl")
    st.stop()
#import msoffcrypto
import sys
sys.path.append('.')

# Then add the imports:
from input_parser import InputParser, Position
from price_fetcher import PriceFetcher
from excel_writer import ExcelWriter


# Import your existing modules
from input_parser import InputParser, Position
from price_fetcher import PriceFetcher
from excel_writer import ExcelWriter

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page config
st.set_page_config(
    page_title="Futures Delivery Calculator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 16px;
    }
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        padding: 1rem;
        border-bottom: 3px solid #1f77b4;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #333;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }
    .info-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border: 1px solid #c3e6cb;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border: 1px solid #ffeeba;
    }
    .error-box {
        background-color: #f8d7da;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
        border: 1px solid #f5c6cb;
    }
</style>
""", unsafe_allow_html=True)


class StreamlitDeliveryApp:
    """Main Streamlit application class"""
    
    def __init__(self):
        self.initialize_session_state()
    
    def initialize_session_state(self):
        """Initialize session state variables"""
        if 'positions' not in st.session_state:
            st.session_state.positions = []
        if 'prices' not in st.session_state:
            st.session_state.prices = {}
        if 'unmapped_symbols' not in st.session_state:
            st.session_state.unmapped_symbols = []
        if 'report_generated' not in st.session_state:
            st.session_state.report_generated = False
        if 'output_file' not in st.session_state:
            st.session_state.output_file = None
    
    def run(self):
        """Main application entry point"""
        # Header
        st.markdown('<h1 class="main-header">📊 Futures & Options Delivery Calculator</h1>', 
                   unsafe_allow_html=True)
        
        # Sidebar for configuration
        with st.sidebar:
            st.header("⚙️ Configuration")
            
            # USDINR Rate
            usdinr_rate = st.number_input(
                "USD/INR Exchange Rate",
                min_value=50.0,
                max_value=150.0,
                value=88.0,
                step=0.1,
                help="Current USD to INR exchange rate for IV calculations"
            )
            
            # Mapping file upload
            st.subheader("📁 Symbol Mapping File")
            mapping_file = st.file_uploader(
                "Upload futures mapping CSV",
                type=['csv'],
                help="CSV file with symbol to ticker mappings"
            )
            
            if not mapping_file:
                st.info("ℹ️ Using default 'futures mapping.csv' if available")
                if os.path.exists('futures mapping.csv'):
                    mapping_file_path = 'futures mapping.csv'
                else:
                    st.error("⚠️ No mapping file found. Please upload one.")
                    mapping_file_path = None
            else:
                # Save uploaded mapping file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp_file:
                    tmp_file.write(mapping_file.getvalue())
                    mapping_file_path = tmp_file.name
            
            st.divider()
            
            # Price fetching options
            st.subheader("💹 Price Options")
            fetch_prices = st.checkbox("Fetch prices from Yahoo Finance", value=True)
            
            if fetch_prices:
                st.info("Prices will be fetched automatically when processing")
        
        # Main content area with tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📤 Upload & Process", "📊 Positions Review", 
                                          "💰 Deliverables Preview", "📥 Download Report"])
        
        with tab1:
            self.upload_and_process_tab(mapping_file_path, usdinr_rate, fetch_prices)
        
        with tab2:
            self.positions_review_tab()
        
        with tab3:
            self.deliverables_preview_tab()
        
        with tab4:
            self.download_report_tab()
    
    def upload_and_process_tab(self, mapping_file_path, usdinr_rate, fetch_prices):
        """Handle file upload and processing"""
        st.markdown('<h2 class="sub-header">Upload Position File</h2>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            uploaded_file = st.file_uploader(
                "Choose your position file",
                type=['xlsx', 'xls', 'csv'],
                help="Upload BOD, CONTRACT, or MS format position file"
            )
        
        with col2:
            if uploaded_file:
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.write("**File Details:**")
                st.write(f"📁 Name: {uploaded_file.name}")
                st.write(f"📏 Size: {uploaded_file.size:,} bytes")
                st.write(f"📋 Type: {uploaded_file.type}")
                st.markdown('</div>', unsafe_allow_html=True)
        
        if uploaded_file and mapping_file_path:
            # Password input for Excel files
            password = None
            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                with st.expander("🔐 Password Protected File?"):
                    password = st.text_input("Enter password (leave blank if not protected)", 
                                           type="password")
            
            # Process button
            if st.button("🚀 Process File", type="primary", use_container_width=True):
                with st.spinner("Processing position file..."):
                    success, message = self.process_file(
                        uploaded_file, mapping_file_path, password, 
                        usdinr_rate, fetch_prices
                    )
                    
                    if success:
                        st.success(f"✅ {message}")
                        st.balloons()
                    else:
                        st.error(f"❌ {message}")
    
    def process_file(self, uploaded_file, mapping_file_path, password, usdinr_rate, fetch_prices):
        """Process the uploaded file"""
        try:
            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                input_file_path = tmp_file.name
            
            # Handle password-protected Excel files
            if password and uploaded_file.name.endswith(('.xlsx', '.xls')):
                try:
                    decrypted = io.BytesIO()
                    with open(input_file_path, 'rb') as f:
                        file = msoffcrypto.OfficeFile(f)
                        file.load_key(password=password)
                        file.decrypt(decrypted)
                    
                    # Save decrypted file
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                        tmp_file.write(decrypted.getvalue())
                        input_file_path = tmp_file.name
                except Exception as e:
                    return False, f"Failed to decrypt file: {str(e)}"
            
            # Parse positions
            parser = InputParser(mapping_file_path)
            positions = parser.parse_file(input_file_path)
            
            if not positions:
                return False, "No valid positions found in the file"
            
            st.session_state.positions = positions
            st.session_state.unmapped_symbols = parser.unmapped_symbols
            
            # Fetch prices if enabled
            if fetch_prices:
                with st.spinner("Fetching prices from Yahoo Finance..."):
                    price_fetcher = PriceFetcher()
                    symbols_to_fetch = list(set(p.symbol for p in positions))
                    symbol_prices = price_fetcher.fetch_prices_for_symbols(symbols_to_fetch)
                    
                    # Map to underlying tickers
                    symbol_map = {}
                    for p in positions:
                        symbol_map[p.underlying_ticker] = p.symbol
                    
                    prices = {}
                    for underlying, symbol in symbol_map.items():
                        if symbol in symbol_prices:
                            prices[underlying] = symbol_prices[symbol]
                    
                    st.session_state.prices = prices
            
            # Generate Excel report
            with st.spinner("Generating Excel report..."):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                format_type = getattr(parser, 'format_type', 'UNKNOWN')
                
                if format_type in ['BOD', 'CONTRACT']:
                    prefix = "GS_AURIGIN_DELIVERY"
                elif format_type == 'MS':
                    prefix = "MS_WAFRA_DELIVERY"
                else:
                    prefix = "DELIVERY_REPORT"
                
                output_file = f"{prefix}_{timestamp}.xlsx"
                
                writer = ExcelWriter(output_file, usdinr_rate)
                writer.create_report(positions, st.session_state.prices, parser.unmapped_symbols)
                
                st.session_state.output_file = output_file
                st.session_state.report_generated = True
            
            return True, f"Successfully processed {len(positions)} positions"
            
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            return False, f"Error processing file: {str(e)}"
    
    def positions_review_tab(self):
        """Display parsed positions for review"""
        st.markdown('<h2 class="sub-header">Position Summary</h2>', unsafe_allow_html=True)
        
        if not st.session_state.positions:
            st.info("📤 Please upload and process a position file first")
            return
        
        positions = st.session_state.positions
        
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Positions", len(positions))
        
        with col2:
            unique_underlyings = len(set(p.underlying_ticker for p in positions))
            st.metric("Unique Underlyings", unique_underlyings)
        
        with col3:
            unique_expiries = len(set(p.expiry_date for p in positions))
            st.metric("Unique Expiries", unique_expiries)
        
        with col4:
            futures_count = sum(1 for p in positions if p.is_future)
            options_count = len(positions) - futures_count
            st.metric("Futures/Options", f"{futures_count}/{options_count}")
        
        # Detailed positions table
        st.subheader("📋 Position Details")
        
        # Convert positions to dataframe
        df_data = []
        for p in positions:
            df_data.append({
                'Underlying': p.underlying_ticker,
                'Symbol': p.symbol,
                'Bloomberg Ticker': p.bloomberg_ticker,
                'Expiry': p.expiry_date.strftime('%Y-%m-%d'),
                'Type': p.security_type,
                'Strike': p.strike_price if p.strike_price > 0 else '',
                'Position (Lots)': p.position_lots,
                'Lot Size': p.lot_size
            })
        
        df = pd.DataFrame(df_data)
        
        # Filters
        with st.expander("🔍 Filters"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                selected_underlyings = st.multiselect(
                    "Filter by Underlying",
                    options=sorted(df['Underlying'].unique()),
                    default=[]
                )
            
            with col2:
                selected_types = st.multiselect(
                    "Filter by Type",
                    options=sorted(df['Type'].unique()),
                    default=[]
                )
            
            with col3:
                selected_expiries = st.multiselect(
                    "Filter by Expiry",
                    options=sorted(df['Expiry'].unique()),
                    default=[]
                )
        
        # Apply filters
        filtered_df = df.copy()
        if selected_underlyings:
            filtered_df = filtered_df[filtered_df['Underlying'].isin(selected_underlyings)]
        if selected_types:
            filtered_df = filtered_df[filtered_df['Type'].isin(selected_types)]
        if selected_expiries:
            filtered_df = filtered_df[filtered_df['Expiry'].isin(selected_expiries)]
        
        # Display table
        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Strike': st.column_config.NumberColumn(format="%.2f"),
                'Position (Lots)': st.column_config.NumberColumn(format="%.2f"),
            }
        )
        
        # Unmapped symbols warning
        if st.session_state.unmapped_symbols:
            st.warning(f"⚠️ {len(st.session_state.unmapped_symbols)} unmapped symbols found")
            with st.expander("View Unmapped Symbols"):
                unmapped_df = pd.DataFrame(st.session_state.unmapped_symbols)
                st.dataframe(unmapped_df, use_container_width=True, hide_index=True)
    
    def deliverables_preview_tab(self):
        """Preview deliverables calculation"""
        st.markdown('<h2 class="sub-header">Deliverables Analysis</h2>', unsafe_allow_html=True)
        
        if not st.session_state.positions:
            st.info("📤 Please upload and process a position file first")
            return
        
        positions = st.session_state.positions
        prices = st.session_state.prices
        
        # Group by underlying
        grouped = {}
        for p in positions:
            if p.underlying_ticker not in grouped:
                grouped[p.underlying_ticker] = []
            grouped[p.underlying_ticker].append(p)
        
        # Sensitivity analysis
        st.subheader("📈 Sensitivity Analysis")
        sensitivity_pct = st.slider(
            "Price Change %",
            min_value=-20.0,
            max_value=20.0,
            value=0.0,
            step=1.0,
            help="Analyze deliverables at different price levels"
        )
        
        # Calculate deliverables for each underlying
        deliverables_data = []
        
        for underlying in sorted(grouped.keys()):
            underlying_positions = grouped[underlying]
            spot_price = prices.get(underlying, 0)
            
            if spot_price:
                adjusted_price = spot_price * (1 + sensitivity_pct / 100)
            else:
                adjusted_price = 0
            
            total_deliverable = 0
            
            for pos in underlying_positions:
                if pos.security_type == 'Futures':
                    deliverable = pos.position_lots
                elif pos.security_type == 'Call':
                    if adjusted_price > pos.strike_price:
                        deliverable = pos.position_lots
                    else:
                        deliverable = 0
                elif pos.security_type == 'Put':
                    if adjusted_price < pos.strike_price:
                        deliverable = -pos.position_lots
                    else:
                        deliverable = 0
                else:
                    deliverable = 0
                
                total_deliverable += deliverable
            
            deliverables_data.append({
                'Underlying': underlying,
                'Current Price': spot_price,
                'Adjusted Price': adjusted_price if spot_price else 'N/A',
                'Total Positions': len(underlying_positions),
                'Net Deliverable (Lots)': total_deliverable
            })
        
        # Display deliverables table
        deliverables_df = pd.DataFrame(deliverables_data)
        
        st.dataframe(
            deliverables_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Current Price': st.column_config.NumberColumn(format="%.2f"),
                'Adjusted Price': st.column_config.NumberColumn(format="%.2f"),
                'Net Deliverable (Lots)': st.column_config.NumberColumn(format="%.0f"),
            }
        )
        
        # Summary metrics
        st.subheader("📊 Summary Statistics")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            total_long = sum(row['Net Deliverable (Lots)'] for row in deliverables_data 
                           if row['Net Deliverable (Lots)'] > 0)
            st.metric("Total Long Deliverable", f"{total_long:.0f} lots")
        
        with col2:
            total_short = sum(row['Net Deliverable (Lots)'] for row in deliverables_data 
                            if row['Net Deliverable (Lots)'] < 0)
            st.metric("Total Short Deliverable", f"{total_short:.0f} lots")
        
        with col3:
            net_deliverable = sum(row['Net Deliverable (Lots)'] for row in deliverables_data)
            st.metric("Net Deliverable", f"{net_deliverable:.0f} lots")
        
        # Price warnings
        missing_prices = [underlying for underlying in grouped.keys() 
                         if underlying not in prices or prices[underlying] == 0]
        
        if missing_prices:
            st.warning(f"⚠️ Missing prices for: {', '.join(missing_prices[:5])}" + 
                      (f" and {len(missing_prices)-5} more" if len(missing_prices) > 5 else ""))
    
    def download_report_tab(self):
        """Download generated Excel report"""
        st.markdown('<h2 class="sub-header">Download Report</h2>', unsafe_allow_html=True)
        
        if not st.session_state.report_generated or not st.session_state.output_file:
            st.info("📤 Please process a position file first to generate the report")
            return
        
        # Report ready message
        st.markdown('<div class="success-box">', unsafe_allow_html=True)
        st.success("✅ **Report Generated Successfully!**")
        st.write(f"**Filename:** {st.session_state.output_file}")
        st.write(f"**Generated at:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Download button
        try:
            with open(st.session_state.output_file, 'rb') as f:
                excel_data = f.read()
            
            st.download_button(
                label="📥 Download Excel Report",
                data=excel_data,
                file_name=st.session_state.output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )
            
            # Report contents description
            st.subheader("📋 Report Contents")
            
            st.info("""
            The Excel report contains the following sheets:
            
            **Deliverable Sheets:**
            - **Master_All_Expiries**: Consolidated view of all positions with deliverable calculations
            - **Expiry_YYYY_MM_DD**: Individual sheets for each expiry date
            - Includes sensitivity analysis columns (±% price changes)
            
            **IV (Intrinsic Value) Sheets:**
            - **IV_All_Expiries**: Consolidated IV calculations in INR and USD
            - **IV_Expiry_YYYY_MM_DD**: IV calculations by expiry date
            
            **Additional Sheets:**
            - **All_Positions**: Complete list of all positions
            - **Unmapped_Symbols**: Symbols that couldn't be mapped (if any)
            
            **Features:**
            - Grouped by underlying with collapsible rows
            - System prices from Yahoo Finance
            - Override columns for manual price adjustments
            - Bloomberg price formulas (=BDP)
            - Dynamic deliverable calculations based on option type
            """)
            
        except Exception as e:
            st.error(f"Error reading report file: {str(e)}")


def main():
    """Main entry point"""
    app = StreamlitDeliveryApp()
    app.run()


if __name__ == "__main__":
    main()
