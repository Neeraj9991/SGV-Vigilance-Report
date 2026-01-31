"""
SGV Vigilance Report Generator
A Streamlit application to generate professional PDF reports from Google Forms inspection data.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import sys
import os

# Add utils to path
sys.path.append(os.path.dirname(__file__))

from utils.sheets_reader import SheetsReader
from utils.image_handler import ImageHandler
from utils.pdf_generator import PDFGenerator

# Page configuration
st.set_page_config(
    page_title="SGV Vigilance Report Generator",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #1a5490 0%, #2563eb 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
    }
    .main-header p {
        margin: 0.5rem 0 0 0;
        opacity: 0.9;
    }
    .stButton>button {
        background-color: #2563eb;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 2rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #1a5490;
    }
    .info-box {
        background: #f0f9ff;
        border-left: 4px solid #2563eb;
        padding: 1rem;
        border-radius: 4px;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="main-header">
        <h1>SGV Vigilance Report Generator</h1>
        <p>Generate Professional Inspection Reports from Google Forms Data</p>
    </div>
""", unsafe_allow_html=True)

# Initialize session state
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'df' not in st.session_state:
    st.session_state.df = None

# Sidebar - Configuration
with st.sidebar:
    st.header("📋 Configuration")
    
    # Check if environment is configured
    try:
        reader = SheetsReader()
        st.success("✅ Connected to Google Sheet")
        
        # Load data button
        if st.button("🔄 Load/Refresh Data", use_container_width=True):
            with st.spinner("Loading data from Google Sheet..."):
                try:
                    st.session_state.df = reader.read_sheet_data()
                    st.session_state.data_loaded = True
                    st.success(f"✅ Loaded {len(st.session_state.df)} inspection records")
                except Exception as e:
                    st.error(f"Error loading data: {str(e)}")
    
    except Exception as e:
        st.error(f"⚠️ Configuration Error: {str(e)}")
        st.info("Please ensure your .env file is configured with GOOGLE_SHEET_ID")
    
    st.divider()
    
    # Help section
    with st.expander("ℹ️ How to Use"):
        st.markdown("""
        1. **Load Data**: Click 'Load/Refresh Data' to fetch latest data
        2. **Apply Filters**: Select date range, shift, and site
        3. **Preview**: Review filtered data in the table
        4. **Generate PDF**: Click to create your report
        5. **Download**: Save the PDF to your device
        """)

# Main content
if not st.session_state.data_loaded:
    st.markdown("""
        <div class="info-box">
            <h3>Welcome!</h3>
            <p>Click <strong>"Load/Refresh Data"</strong> in the sidebar to get started.</p>
            <p>This will fetch the latest inspection data from your Google Sheet.</p>
        </div>
    """, unsafe_allow_html=True)
else:
    df = st.session_state.df
    
    # Filters section
    st.header("Filters")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Date range filter
        st.subheader("Date Range")
        
        # Get min and max dates from data
        try:
            df['Date_parsed'] = pd.to_datetime(df['Date'], errors='coerce')
            min_date = df['Date_parsed'].min().date() if not df['Date_parsed'].isna().all() else datetime.now().date() - timedelta(days=30)
            max_date = df['Date_parsed'].max().date() if not df['Date_parsed'].isna().all() else datetime.now().date()
        except:
            min_date = datetime.now().date() - timedelta(days=30)
            max_date = datetime.now().date()
        
        start_date = st.date_input(
            "Start Date",
            value=max_date - timedelta(days=7),
            min_value=min_date,
            max_value=max_date
        )
        
        end_date = st.date_input(
            "End Date",
            value=max_date,
            min_value=min_date,
            max_value=max_date
        )
    
    with col2:
        # Shift filter
        st.subheader("Shift")
        shifts = ["All"] + reader.get_unique_shifts(df)
        selected_shift = st.selectbox("Select Shift", shifts)
    
    with col3:
        # Site filter
        st.subheader("Site")
        sites = ["All"] + reader.get_unique_sites(df)
        selected_site = st.selectbox("Select Site", sites)
    
    # Apply filters
    filtered_df = df.copy()
    filtered_df = reader.filter_by_date_range(filtered_df, start_date, end_date)
    filtered_df = reader.filter_by_shift(filtered_df, selected_shift)
    filtered_df = reader.filter_by_site(filtered_df, selected_site)
    
    st.divider()
    
    # Display filtered data
    st.header("📊 Filtered Data Preview")
    
    if len(filtered_df) == 0:
        st.warning("⚠️ No inspection records match the selected filters.")
    else:
        st.success(f"✅ Found **{len(filtered_df)}** inspection record(s)")
        
        # Display data table
        display_columns = [
            'Date', 'Time', 'Site Name', 'Shift', 
            'Performance Check [Overall Rating]', 'Inspected By'
        ]
        
        # Only show columns that exist
        available_columns = [col for col in display_columns if col in filtered_df.columns]
        
        st.dataframe(
            filtered_df[available_columns],
            use_container_width=True,
            hide_index=True
        )
        
        st.divider()
        
        # Generate PDF section
        st.header("📄 Generate Report")
        if st.button("Generate PDF Report", use_container_width=True, type="primary"):
            with st.spinner("Generating PDF report... This may take a moment to download images."):
                    try:
                        # Initialize handlers
                        image_handler = ImageHandler()
                        pdf_generator = PDFGenerator()
                        
                        # Generate PDF
                        pdf_bytes = pdf_generator.generate_pdf(filtered_df, image_handler)
                        
                        # Generate filename
                        filename = pdf_generator.generate_report_filename(
                            start_date=start_date,
                            end_date=end_date,
                            shift=selected_shift if selected_shift != "All" else None
                        )
                        
                        # Download button
                        st.success("✅ Report generated successfully!")
                        
                        st.download_button(
                            label="⬇️ Download PDF Report",
                            data=pdf_bytes,
                            file_name=filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                        
                        # Cleanup
                        image_handler.cleanup_temp_dir()
                    
                    except Exception as e:
                        st.error(f"❌ Error generating PDF: {str(e)}")
                        st.exception(e)

# Footer
st.divider()
st.markdown("""
    <div style="text-align: center; color: #666; font-size: 0.9rem; padding: 1rem;">
        <strong>SGV SUPER SECURITY SERVICE PVT. LTD</strong><br>
        Vigilance Report Generator v1.0 | For support, contact your IT department
    </div>
""", unsafe_allow_html=True)
