# Dual Report UI Section for app.py
# This code should replace lines 255-294 in app.py

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
        
        if st.button("Generate Internal PDF Report", use_container_width=True, type="primary"):
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
                        use_container_width=True
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
                docx_handler = DocxHandler()
                
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
                if len(filtered_df) > 1:
                    st.info(f"You have {len(filtered_df)} records. The first record will be used for the report.")
                
                if st.button("Generate Client PDF Report", use_container_width=True, type="primary"):
                    with st.spinner("Generating client PDF report from template..."):
                        try:
                            uploaded_template.seek(0)  # Reset file pointer
                            
                            # Generate PDF from template
                            pdf_bytes = docx_handler.generate_client_report(
                                uploaded_template,
                                filtered_df,
                                row_index=0
                            )
                            
                            # Generate filename
                            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                            filename = f"SGV_Client_Report_{timestamp}.pdf"
                            
                            # Download button
                            st.success("✅ Client report generated successfully!")
                            
                            st.download_button(
                                label="⬇️ Download Client PDF Report",
                                data=pdf_bytes,
                                file_name=filename,
                                mime="application/pdf",
                                use_container_width=True
                            )
                        
                        except Exception as e:
                            st.error(f"❌ Error generating client PDF: {str(e)}")
                            st.exception(e)
            
            except Exception as e:
                st.error(f"❌ Error processing template: {str(e)}")
        else:
            st.warning("👆 Please upload a DOCX template to continue")

else:
    st.info("Incident report PDF generation will be available soon.")
