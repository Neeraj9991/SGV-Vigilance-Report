"""
PDF Generator Module
Generates professional PDF reports from inspection data using HTML templates.
"""

from jinja2 import Environment, FileSystemLoader
from xhtml2pdf import pisa
from datetime import datetime
from io import BytesIO
import os

class PDFGenerator:
    def __init__(self, template_dir="templates"):
        """
        Initialize the PDF generator.
        
        Args:
            template_dir (str): Directory containing HTML templates
        """
        self.template_dir = template_dir
        self.env = Environment(loader=FileSystemLoader(template_dir))
    
    def parse_site_name(self, site_name):
        """
        Parse site name in format: 4-311-DLF SCO-84
        Returns: (zone, unit_code, actual_site_name)
        """
        if not site_name or '-' not in str(site_name):
            return ('', '', str(site_name or ''))
        
        try:
            parts = str(site_name).split('-', 2)  # Split into max 3 parts
            if len(parts) >= 3:
                return (parts[0].strip(), parts[1].strip(), parts[2].strip())
            elif len(parts) == 2:
                return (parts[0].strip(), '', parts[1].strip())
            else:
                return ('', '', site_name)
        except:
            return ('', '', str(site_name or ''))
    
    def load_logo_base64(self):
        """
        Load logo from assets/logo.png and convert to base64.
        Returns base64 string or None if not found.
        """
        import base64
        from pathlib import Path
        
        logo_path = Path('assets/logo.png')
        
        try:
            if logo_path.exists():
                with open(logo_path, 'rb') as f:
                    logo_data = f.read()
                    return base64.b64encode(logo_data).decode('utf-8')
        except Exception as e:
            print(f"Warning: Could not load logo: {str(e)}")
        
        return None
    
    def format_data_for_template(self, df, image_handler):
        """
        Format dataframe rows for template rendering.
        
        Args:
            df (pd.DataFrame): Filtered inspection data
            image_handler (ImageHandler): Image handler instance
            
        Returns:
            list: List of dictionaries with inspection data and images
        """
        import pandas as pd
        
        inspections = []
        
        for _, row in df.iterrows():
            # Download and encode images
            images = []
            if 'Images' in row and pd.notna(row['Images']) and row['Images']:
                images = image_handler.download_and_encode_images(row['Images'])
            
            # Helper function to safely get string value
            def safe_get(key, default=''):
                val = row.get(key, default)
                if pd.isna(val):
                    return default
                return str(val)
            
            # Parse site name
            site_name = safe_get('Site Name')
            zone, unit_code, actual_site = self.parse_site_name(site_name)
            
            inspection = {
                'timestamp': safe_get('Timestamp'),
                'date': safe_get('Date'),
                'time': safe_get('Time'),
                'site_name': site_name,
                'zone': zone,
                'unit_code': unit_code,
                'actual_site_name': actual_site,
                'shift': safe_get('Shift'),
                'inspected_by': safe_get('Inspected By'),
                'email': safe_get('Email Address'),
                
                # Documentation Checks
                'attendance_register': safe_get('Documentation Check [Attendance Register]'),
                'handover_register': safe_get('Documentation Check [Handling / Taking Over Register]'),
                'visitor_log': safe_get('Documentation Check [Visitor Log Register]'),
                'incident_log': safe_get('Documentation Check [Incident Log]'),
                
                # Performance Checks
                'grooming': safe_get('Performance Check [Grooming]'),
                'alertness': safe_get('Performance Check [Alertness]'),
                'post_discipline': safe_get('Performance Check [Post Discipline]'),
                'job_awareness': safe_get('Performance Check [Job Awareness]'),
                
                # Incident & Actions
                'incident_reported': safe_get('Incident Reported, if any during the QC/ Night Check (Provide Details)'),
                
                # Employee Cases
                'sleeping_cases': safe_get('Sleeping cases Found (Provide Name, Emp Id / Father Name)'),
                'not_on_duty_cases': safe_get('Not on Duty/Post Cases Found (Provide Name, Emp Id / Father Name)'),
                'misbehaving_drunk_cases': safe_get('Found Misbehaving / Drunk (Provide Name, Emp Id / Father Name)'),
                'cases_shared': safe_get('Were the identified cases shared with Supervisor, FO and Manager Operations for necessary action'),
                
                # Observations
                'security_observations': safe_get('Any other security related observations during Quality Check'),
                
                # Images
                'images': images
            }
            
            inspections.append(inspection)
        
        return inspections
    
    def format_incident_data_for_template(self, df, image_handler):
        """
        Format dataframe rows for incident report template rendering.
        
        Args:
            df (pd.DataFrame): Filtered incident data
            image_handler (ImageHandler): Image handler instance
            
        Returns:
            list: List of dictionaries with incident data and images
        """
        import pandas as pd
        
        incidents = []
        
        for _, row in df.iterrows():
            # Download and encode images/evidence (Photos only, not Videos)
            images = []
            if 'EVIDENCE & ATTACHMENTS - Photos' in row and pd.notna(row['EVIDENCE & ATTACHMENTS - Photos']) and row['EVIDENCE & ATTACHMENTS - Photos']:
                images = image_handler.download_and_encode_images(row['EVIDENCE & ATTACHMENTS - Photos'])
            
            # Helper function to safely get string value
            def safe_get(key, default=''):
                val = row.get(key, default)
                if pd.isna(val):
                    return default
                return str(val)
            
            # Parse site name
            site_name = safe_get('Site Name')
            zone, unit_code, actual_site = self.parse_site_name(site_name)
            
            # Get video evidence URLs (if present)
            video_urls = []
            if 'EVIDENCE & ATTACHMENTS - Videos' in row and pd.notna(row['EVIDENCE & ATTACHMENTS - Videos']) and row['EVIDENCE & ATTACHMENTS - Videos']:
                # Parse video URLs (comma-separated)
                video_string = str(row['EVIDENCE & ATTACHMENTS - Videos'])
                video_urls = [url.strip() for url in video_string.split(',') if url.strip()]
            
            # Create brief summary from description (first 100 chars)
            description = safe_get('Detailed Description of the Incident')
            brief_summary = (description[:100] + '...') if len(description) > 100 else description
            
            incident = {
                'timestamp': safe_get('Timestamp'),
                'email': safe_get('Email Address'),
                'reporting_person': safe_get('Full Name of Reporting Person'),
                'designation': safe_get('Designation / Role'),
                'date': safe_get('Date of Incident'),
                'time': safe_get('Time of Incident'),
                'site_name': site_name,
                'zone': zone,
                'unit_code': unit_code,
                'actual_site_name': actual_site,
                'location': safe_get('Exact Location of Incident within Site'),
                'category': safe_get('Category of Incident'),
                'description': description,
                'brief_summary': brief_summary,
                'injury_reported': safe_get('Was any injury or harm reported?'),
                'injury_details': safe_get('If YES, please provide details of the injury / medical condition.'),
                'immediate_action': safe_get('Immediate Action Taken'),
                'authorities_informed': safe_get('Any Internal / External Authorities Informed?'),
                'cctv_available': safe_get('Availability of CCTV Footage'),
                'support_required': safe_get('Is further support or intervention required from Head Office / Management?'),
                'support_details': safe_get('If YES, mention required support'),
                'status': safe_get('Status'),
                'images': images,
                'video_urls': video_urls
            }
            
            incidents.append(incident)
        
        return incidents

    
    def generate_pdf(self, df, image_handler, output_path=None):
        """
        Generate PDF report from inspection data.
        
        Args:
            df (pd.DataFrame): Filtered inspection data
            image_handler (ImageHandler): Image handler instance
            output_path (str): Optional path to save PDF file
            
        Returns:
            bytes: PDF file as bytes
        """
        try:
            # Format data for template
            inspections = self.format_data_for_template(df, image_handler)
            
            # Load template
            template = self.env.get_template('report_template.html')
            
            # Prepare context data
            context = {
                'inspections': inspections,
                'report_date': datetime.now().strftime('%B %d, %Y'),
                'report_time': datetime.now().strftime('%I:%M %p'),
                'total_inspections': len(inspections),
                'logo_base64': self.load_logo_base64()
            }
            
            # Render HTML
            html_content = template.render(context)
            
            # Generate PDF using xhtml2pdf
            pdf_buffer = BytesIO()
            pisa_status = pisa.CreatePDF(
                html_content,
                dest=pdf_buffer
            )
            
            if pisa_status.err:
                raise Exception(f"PDF generation failed with error code: {pisa_status.err}")
            
            pdf = pdf_buffer.getvalue()
            
            # Save to file if output path provided
            if output_path:
                with open(output_path, 'wb') as f:
                    f.write(pdf)
            
            return pdf
        
        except Exception as e:
            raise Exception(f"Error generating PDF: {str(e)}")
    
    def generate_incident_pdf(self, df, image_handler, output_path=None):
        """
        Generate incident PDF report from incident data.
        
        Args:
            df (pd.DataFrame): Filtered incident data
            image_handler (ImageHandler): Image handler instance
            output_path (str): Optional path to save PDF file
            
        Returns:
            bytes: PDF file as bytes
        """
        try:
            # Format data for template
            incidents = self.format_incident_data_for_template(df, image_handler)
            
            # Load incident template
            template = self.env.get_template('incident_report_template.html')
            
            # Prepare context data
            context = {
                'incidents': incidents,
                'report_date': datetime.now().strftime('%B %d, %Y'),
                'report_time': datetime.now().strftime('%I:%M %p'),
                'total_incidents': len(incidents),
                'logo_base64': self.load_logo_base64()
            }
            
            # Render HTML
            html_content = template.render(context)
            
            # Generate PDF using xhtml2pdf
            pdf_buffer = BytesIO()
            pisa_status = pisa.CreatePDF(
                html_content,
                dest=pdf_buffer
            )
            
            if pisa_status.err:
                raise Exception(f"PDF generation failed with error code: {pisa_status.err}")
            
            pdf = pdf_buffer.getvalue()
            
            # Save to file if output path provided
            if output_path:
                with open(output_path, 'wb') as f:
                    f.write(pdf)
            
            return pdf
        
        except Exception as e:
            raise Exception(f"Error generating incident PDF: {str(e)}")

    
    def generate_report_filename(self, start_date=None, end_date=None, shift=None):
        """
        Generate a filename for the report.
        
        Args:
            start_date: Start date for the report
            end_date: End date for the report
            shift (str): Shift filter
            
        Returns:
            str: Generated filename
        """
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        if start_date and end_date:
            date_str = f"{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}"
        else:
            date_str = "all_dates"
        
        shift_str = f"_{shift}" if shift and shift != "All" else ""
        
        return f"SGV_Vigilance_Report_{date_str}{shift_str}_{timestamp}.pdf"
