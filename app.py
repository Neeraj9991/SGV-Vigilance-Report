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
    page_title="SGV ReportGen",
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
        <h1>SGV ReportGen</h1>
        <p>Generate Professional Reports in minutes</p>
    </div>
""", unsafe_allow_html=True)

# Initialize session state
if 'monitoring_data_loaded' not in st.session_state:
    st.session_state.monitoring_data_loaded = False
if 'incident_data_loaded' not in st.session_state:
    st.session_state.incident_data_loaded = False
if 'monitoring_df' not in st.session_state:
    st.session_state.monitoring_df = None
if 'incident_df' not in st.session_state:
    st.session_state.incident_df = None
if 'active_tab' not in st.session_state:
    st.session_state.active_tab = 'monitoring'

# Sidebar - Navigation Tabs
with st.sidebar:
    st.header("📋 Navigation")
    
    # Tab selection
    tab = st.radio(
        "Select Form Type:",
        ["Monitoring Visit Record", "Report Incident"],
        key="nav_tabs"
    )
    
    # Set active tab
    if tab == "Monitoring Visit Record":
        st.session_state.active_tab = 'monitoring'
    else:
        st.session_state.active_tab = 'incident'
    
    st.divider()
    
    # Help section
    with st.expander("ℹ️ How to Use"):
        st.markdown("""
        1. **Select Form Type**: Choose Monitoring or Incident
        2. **Load Data**: Click 'Load/Refresh Data' to fetch latest data
        3. **Apply Filters**: Select date range, shift, and site
        4. **Preview**: Review filtered data in the table
        5. **Generate PDF**: Click to create your report
        6. **Download**: Save the PDF to your device
        """)

# Main content
st.header(f"{'Monitoring & Evaluation' if st.session_state.active_tab == 'monitoring' else 'Incident Reporting'}")

# Load/Refresh Data Section
col1, col2 = st.columns([2, 2])
    
with col1:
    # Check if environment is configured
    try:
        reader = SheetsReader(sheet_type=st.session_state.active_tab)
        
        # Load data button
        if st.button("🔄 Load/Refresh Data", width="stretch", type="primary"):
            with st.spinner(f"Loading {st.session_state.active_tab} data from Google Sheet..."):
                try:
                    data = reader.read_sheet_data()
                    
                    if st.session_state.active_tab == 'monitoring':
                        st.session_state.monitoring_df = data
                        st.session_state.monitoring_data_loaded = True
                        st.success(f"✅ Loaded {len(data)} monitoring records")
                    else:
                        st.session_state.incident_df = data
                        st.session_state.incident_data_loaded = True
                        st.success(f"✅ Loaded {len(data)} incident records")
                except Exception as e:
                    st.error(f"Error loading data: {str(e)}")
    
    except Exception as e:
        st.error(f"⚠️ Configuration Error: {str(e)}")
        st.info("Please ensure your .env file is configured correctly")

st.divider()

# Get current data based on active tab
current_data_loaded = (st.session_state.monitoring_data_loaded if st.session_state.active_tab == 'monitoring' 
                       else st.session_state.incident_data_loaded)
current_df = (st.session_state.monitoring_df if st.session_state.active_tab == 'monitoring' 
              else st.session_state.incident_df)

# Display content based on data load status
if not current_data_loaded:
    st.markdown("""
        <div class="info-box">
            <h3>Welcome!</h3>
            <p>Click <strong>"Load/Refresh Data"</strong> above to get started.</p>
            <p>This will fetch the latest data from your Google Sheet.</p>
        </div>
    """, unsafe_allow_html=True)
else:
    df = current_df
    
    # Filters section
    st.header("🔍 Filters")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Date range filter
        st.subheader("Date Range")
        
        # Get min and max dates from data
        try:
            # Use correct date column based on sheet type
            date_column = 'Date' if st.session_state.active_tab == 'monitoring' else 'Date of Incident'
            df['Date_parsed'] = pd.to_datetime(df[date_column], errors='coerce')
            min_date = df['Date_parsed'].min().date() if not df['Date_parsed'].isna().all() else datetime.now().date() - timedelta(days=30)
            max_date = df['Date_parsed'].max().date() if not df['Date_parsed'].isna().all() else datetime.now().date()
        except:
            min_date = datetime.now().date() - timedelta(days=30)
            max_date = datetime.now().date()
        
        start_date = st.date_input(
            "Start Date",
            value=max(min_date, max_date - timedelta(days=7)),  # Ensure value is within range
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
        st.warning("⚠️ No records match the selected filters.")
    else:
        st.success(f"✅ Found **{len(filtered_df)}** record(s)")
        
        # Display data table with different columns based on form type
        if st.session_state.active_tab == 'monitoring':
            display_columns = [
                'Date', 'Time', 'Site Name', 'Shift', 
                'Performance Check [Grooming]', 
                'Performance Check [Alertness]',
                'Performance Check [Post Discipline]',
                'Inspected By'
            ]
        else:
            display_columns = [
                'Date of Incident', 'Time of Incident', 'Site Name', 
                'Category of Incident', 'Status', 'Full Name of Reporting Person'
            ]
        
        # Only show columns that exist
        available_columns = [col for col in display_columns if col in filtered_df.columns]
        
        st.dataframe(
            filtered_df[available_columns],
            width="stretch",
            hide_index=True
        )
        
        
        # Generate PDF section
        st.header("📄 Generate Report")
        
        if st.session_state.active_tab == 'monitoring':
            # Report Type Selection
            st.subheader("Select Report Type")
            report_type = st.radio(
                "Choose report format:",
                ["Internal Report", "Client Report"],
                help="Internal reports include all 25 fields. Client reports use your uploaded DOCX template."
            )
            
            st.divider()
            
            # Internal Report
            if report_type == "Internal Report":
                st.info("✨ **Internal Report** includes all monitoring fields: documentation checks, performance metrics, employee cases, incidents, and observations.")
                
                if st.button("Generate Internal PDF Report", width="stretch", type="primary"):
                    with st.spinner("Generating comprehensive PDF report... This may take a moment to download images."):
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
                            st.success("✅ Internal report generated successfully!")
                            
                            st.download_button(
                                label="⬇️ Download Internal PDF Report",
                                data=pdf_bytes,
                                file_name=filename,
                                mime="application/pdf",
                                width="stretch"
                            )
                            
                            # Cleanup
                            image_handler.cleanup_temp_dir()
                        
                        except Exception as e:
                            st.error(f"❌ Error generating PDF: {str(e)}")
                            st.exception(e)
            
            # Client Report
            else:
                from utils.docx_handler import DocxHandler
                
                st.info("📄 **Client Report** uses your custom DOCX template. Upload a template with placeholders like `{Date}`, `{Site Name}`, etc.")
                
                # File uploader
                uploaded_template = st.file_uploader(
                    "Upload DOCX Template",
                    type=['docx'],
                    help="Template should contain placeholders matching your Excel column names, e.g., {Date}, {Shift}, {Site Name}"
                )
                
                if uploaded_template:
                    try:
                        # Initialize handlers
                        image_handler = ImageHandler()
                        docx_handler = DocxHandler(image_handler=image_handler)
                        
                        # Extract placeholders
                        placeholders = docx_handler.extract_placeholders(uploaded_template)
                        uploaded_template.seek(0)  # Reset file pointer
                        
                        # Match with dataframe columns
                        matched, unmatched = docx_handler.match_placeholders_to_columns(
                            placeholders,
                            filtered_df.columns.tolist()
                        )
                        
                        # Display placeholder information
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.success(f"✅ **Matched Placeholders:** {len(matched)}")
                            if matched:
                                with st.expander("View Matched Placeholders"):
                                    for placeholder in sorted(matched.keys()):
                                        st.write(f"• `{{{placeholder}}}`")
                        
                        with col2:
                            if unmatched:
                                st.warning(f"⚠️ **Unmatched Placeholders:** {len(unmatched)}")
                                with st.expander("View Unmatched Placeholders"):
                                    for placeholder in sorted(unmatched):
                                        st.write(f"• `{{{placeholder}}}`")
                            else:
                                st.info("✨ All placeholders matched!")
                        
                        
                        st.divider()
                        
                        # Generate button
                        record_count = len(filtered_df)
                        if record_count > 1:
                            st.info(f"You have **{record_count} records**. A separate PDF will be generated for each record, named by site name.")
                        else:
                            st.info(f"You have **1 record**.")
                        
                        if st.button("Generate Client PDF Report", width="stretch", type="primary"):
                            with st.spinner(f"Generating {record_count} client report(s)... Downloading and embedding images..."):
                                try:
                                    uploaded_template.seek(0)  # Reset file pointer
                                    
                                    if record_count == 1:
                                        # Single record - direct PDF download
                                        pdf_bytes = docx_handler.generate_client_report(
                                            uploaded_template,
                                            filtered_df,
                                            row_index=0
                                        )
                                        
                                        # Get site name for filename
                                        site_name = filtered_df.iloc[0].get('Site Name', 'Site')
                                        import re
                                        clean_site_name = re.sub(r'[<>:"/\\|?*]', '_', str(site_name)).strip()
                                        date_str = filtered_df.iloc[0].get('Date', datetime.now().strftime('%Y-%m-%d'))
                                        if pd.notna(date_str):
                                            date_str = str(date_str).replace('/', '-')
                                        else:
                                            date_str = datetime.now().strftime('%Y-%m-%d')
                                        
                                        filename = f"{clean_site_name}_{date_str}_Report.pdf"
                                        
                                        st.success("✅ Client report with embedded images generated successfully!")
                                        
                                        st.download_button(
                                            label="⬇️ Download Client PDF Report",
                                            data=pdf_bytes,
                                            file_name=filename,
                                            mime="application/pdf",
                                            width="stretch"
                                        )
                                    
                                    else:
                                        # Multiple records - generate ZIP
                                        # Validate DataFrame before generating
                                        if filtered_df is None or len(filtered_df) == 0:
                                            st.error("❌ No data available to generate reports. Please check your filters.")
                                        else:
                                            zip_bytes, report_count, file_list = docx_handler.generate_multiple_client_reports(
                                                uploaded_template,
                                                filtered_df
                                            )
                                            
                                            # Generate ZIP filename
                                            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                                            zip_filename = f"SGV_Client_Reports_{timestamp}.zip"
                                            
                                            st.success(f"✅ Generated {report_count} client reports with embedded images!")
                                            
                                            # Show list of generated files
                                            with st.expander(f"📁 View generated files ({report_count} PDFs)"):
                                                for i, filename in enumerate(file_list, 1):
                                                    st.write(f"{i}. `{filename}`")
                                            
                                            st.download_button(
                                                label=f"⬇️ Download ZIP with {report_count} Reports",
                                                data=zip_bytes,
                                                file_name=zip_filename,
                                                mime="application/zip",
                                                width="stretch"
                                            )
                                    
                                    # Cleanup
                                    image_handler.cleanup_temp_dir()
                                
                                except Exception as e:
                                    st.error(f"❌ Error generating client PDF: {str(e)}")
                                    st.exception(e)
                    
                    except Exception as e:
                        st.error(f"❌ Error processing template: {str(e)}")
                else:
                    st.warning("👆 Please upload a DOCX template to continue")
        
        else:
            # Incident Reporting
            # Report Type Selection
            st.subheader("Select Report Type")
            report_type = st.radio(
                "Choose report format:",
                ["Internal Report", "Client Report"],
                help="Internal reports include all incident fields. Client reports use your uploaded DOCX template."
            )
            
            st.divider()
            
            # Internal Report
            if report_type == "Internal Report":
                st.info("✨ **Internal Report** includes all incident fields: reported by, incident details, injury information, actions taken, authorities informed, and evidence.")
                
                if st.button("Generate Internal PDF Report", key="incident_internal_pdf", type="primary"):
                    with st.spinner("Generating comprehensive incident PDF report... This may take a moment to download images."):
                        try:
                            # Initialize handlers
                            image_handler = ImageHandler()
                            pdf_generator = PDFGenerator()
                            
                            # Generate PDF
                            pdf_bytes = pdf_generator.generate_incident_pdf(filtered_df, image_handler)
                            
                            # Generate filename
                            filename = pdf_generator.generate_report_filename(
                                start_date=start_date,
                                end_date=end_date,
                                shift=None
                            ).replace('Vigilance_Report', 'Incident_Report')
                            
                            # Download button
                            st.success("✅ Internal incident report generated successfully!")
                            
                            st.download_button(
                                label="⬇️ Download Internal PDF Report",
                                data=pdf_bytes,
                                file_name=filename,
                                mime="application/pdf",
                                key="download_incident_internal_pdf"
                            )
                            
                            # Cleanup
                            image_handler.cleanup_temp_dir()
                        
                        except Exception as e:
                            st.error(f"❌ Error generating PDF: {str(e)}")
                            st.exception(e)
            
            # Client Report
            else:
                from utils.docx_handler import DocxHandler
                
                st.info("📄 **Client Report** uses your custom DOCX template. Upload a template with placeholders like `{Date of Incident}`, `{Site Name}`, `{Category of Incident}`, `{EVIDENCE & ATTACHMENTS - Photos}`, etc.")
                
                # File uploader
                uploaded_template = st.file_uploader(
                    "Upload DOCX Template",
                    type=['docx'],
                    help="Template should contain placeholders matching your incident form column names, e.g., {Date of Incident}, {Site Name}, {EVIDENCE & ATTACHMENTS - Photos}",
                    key="incident_template_upload"
                )
                
                if uploaded_template:
                    try:
                        # Initialize handlers
                        image_handler = ImageHandler()
                        docx_handler = DocxHandler(image_handler=image_handler)
                        
                        # Extract placeholders
                        placeholders = docx_handler.extract_placeholders(uploaded_template)
                        uploaded_template.seek(0)  # Reset file pointer
                        
                        # Match with dataframe columns
                        matched, unmatched = docx_handler.match_placeholders_to_columns(
                            placeholders,
                            filtered_df.columns.tolist()
                        )
                        
                        # Display placeholder information
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.success(f"✅ **Matched Placeholders:** {len(matched)}")
                            if matched:
                                with st.expander("View Matched Placeholders"):
                                    for placeholder in sorted(matched.keys()):
                                        st.write(f"• `{{{placeholder}}}`")
                        
                        with col2:
                            if unmatched:
                                st.warning(f"⚠️ **Unmatched Placeholders:** {len(unmatched)}")
                                with st.expander("View Unmatched Placeholders"):
                                    for placeholder in sorted(unmatched):
                                        st.write(f"• `{{{placeholder}}}`")
                            else:
                                st.info("✨ All placeholders matched!")
                        
                        
                        st.divider()
                        
                        # Generate button
                        record_count = len(filtered_df)
                        if record_count > 1:
                            st.info(f"You have **{record_count} records**. A separate PDF will be generated for each record, named by site name.")
                        else:
                            st.info(f"You have **1 record**.")
                        
                        if st.button("Generate Client PDF Report", key="incident_client_pdf", type="primary"):
                            with st.spinner(f"Generating {record_count} client report(s)... Downloading and embedding images..."):
                                try:
                                    uploaded_template.seek(0)  # Reset file pointer
                                    
                                    if record_count == 1:
                                        # Single record - direct PDF download
                                        pdf_bytes = docx_handler.generate_client_report(
                                            uploaded_template,
                                            filtered_df,
                                            row_index=0
                                        )
                                        
                                        # Get site name for filename
                                        site_name = filtered_df.iloc[0].get('Site Name', 'Site')
                                        import re
                                        clean_site_name = re.sub(r'[<>:"/\\|?*]', '_', str(site_name)).strip()
                                        date_str = filtered_df.iloc[0].get('Date of Incident', datetime.now().strftime('%Y-%m-%d'))
                                        if pd.notna(date_str):
                                            date_str = str(date_str).replace('/', '-')
                                        else:
                                            date_str = datetime.now().strftime('%Y-%m-%d')
                                        
                                        filename = f"{clean_site_name}_{date_str}_Incident_Report.pdf"
                                        
                                        st.success("✅ Client incident report with embedded images generated successfully!")
                                        
                                        st.download_button(
                                            label="⬇️ Download Client PDF Report",
                                            data=pdf_bytes,
                                            file_name=filename,
                                            mime="application/pdf",
                                            key="download_incident_client_pdf"
                                        )
                                    
                                    else:
                                        # Multiple records - generate ZIP
                                        # Validate DataFrame before generating
                                        if filtered_df is None or len(filtered_df) == 0:
                                            st.error("❌ No data available to generate reports. Please check your filters.")
                                        else:
                                            zip_bytes, report_count, file_list = docx_handler.generate_multiple_client_reports(
                                                uploaded_template,
                                                filtered_df
                                            )
                                            
                                            # Generate ZIP filename
                                            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                                            zip_filename = f"SGV_Incident_Reports_{timestamp}.zip"
                                            
                                            st.success(f"✅ Generated {report_count} incident reports with embedded images!")
                                            
                                            # Show list of generated files
                                            with st.expander(f"📁 View generated files ({report_count} PDFs)"):
                                                for i, filename in enumerate(file_list, 1):
                                                    st.write(f"{i}. `{filename}`")
                                            
                                            st.download_button(
                                                label=f"⬇️ Download ZIP with {report_count} Reports",
                                                data=zip_bytes,
                                                file_name=zip_filename,
                                                mime="application/zip",
                                                key="download_incident_client_zip"
                                            )
                                    
                                    # Cleanup
                                    image_handler.cleanup_temp_dir()
                                
                                except Exception as e:
                                    st.error(f"❌ Error generating client PDF: {str(e)}")
                                    st.exception(e)
                    
                    except Exception as e:
                        st.error(f"❌ Error processing template: {str(e)}")
                else:
                    st.warning("👆 Please upload a DOCX template to continue")

# Footer
st.divider()
st.markdown("""
    <div style="text-align: center; color: #666; font-size: 0.9rem; padding: 1rem;">
        <strong>SGV SUPER SECURITY SERVICE PVT. LTD</strong><br>
        Vigilance Report Generator v2.0 | For support, contact your IT department
    </div>
""", unsafe_allow_html=True)
