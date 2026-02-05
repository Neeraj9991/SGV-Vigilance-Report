"""
Google Sheets Reader Module
Reads data from a public Google Sheet and filters it for report generation.
"""

import pandas as pd
import os
from dotenv import load_dotenv
from datetime import datetime

# Load environment variables
load_dotenv()

class SheetsReader:
    def __init__(self, sheet_type='monitoring'):
        """
        Initialize the sheets reader with configuration from environment variables.
        
        Args:
            sheet_type (str): Type of sheet to read - 'monitoring' or 'incident'
        """
        self.sheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.sheet_type = sheet_type
        
        # Get the appropriate GID based on sheet type
        if sheet_type == 'monitoring':
            self.sheet_gid = os.getenv('MONITORING_SHEET_GID', '0')
        elif sheet_type == 'incident':
            self.sheet_gid = os.getenv('INCIDENT_SHEET_GID', '0')
        else:
            raise ValueError(f"Invalid sheet_type: {sheet_type}. Must be 'monitoring' or 'incident'")
        
        if not self.sheet_id:
            raise ValueError("GOOGLE_SHEET_ID not found in environment variables")
    
    def get_csv_url(self):
        """Construct the public CSV export URL for the Google Sheet."""
        return f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/export?format=csv&gid={self.sheet_gid}"
    
    def read_sheet_data(self):
        """
        Read data from the public Google Sheet.
        
        Returns:
            pd.DataFrame: DataFrame containing all sheet data
        """
        try:
            csv_url = self.get_csv_url()
            df = pd.read_csv(csv_url)
            
            # Expected columns based on sheet type
            if self.sheet_type == 'monitoring':
                expected_columns = [
                    'Timestamp', 'Date', 'Time', 'Site Name', 'Shift',
                    'Documentation Check [Attendance Register]',
                    'Documentation Check [Handling / Taking Over Register]',
                    'Documentation Check [Visitor Log Register]',
                    'Documentation Check [Incident Log]',
                    'Performance Check [Grooming]',
                    'Performance Check [Alertness]',
                    'Performance Check [Post Discipline]',
                    'Performance Check [Job Awareness]',
                    'Any other security related observations during Quality Check',
                    'Inspected By',
                    'Images',
                    'Email Address',
                    'Incident Reported, if any during the QC/ Night Check (Provide Details)',
                    'Sleeping cases Found (Provide Name, Emp Id / Father Name)',
                    'Not on Duty/Post Cases Found (Provide Name, Emp Id / Father Name)',
                    'Found Misbehaving / Drunk (Provide Name, Emp Id / Father Name)',
                    'Were the identified cases shared with Supervisor, FO and Manager Operations for necessary action'
                ]
            else:
                # Incident form columns - updated with separate photo/video fields
                expected_columns = [
                    'Timestamp',
                    'Email Address',
                    'Full Name of Reporting Person',
                    'Designation / Role',
                    'Date of Incident',
                    'Time of Incident',
                    'Site Name',
                    'Exact Location of Incident within Site',
                    'Category of Incident',
                    'Detailed Description of the Incident',
                    'Was any injury or harm reported?',
                    'If YES, please provide details of the injury / medical condition.',
                    'Immediate Action Taken',
                    'Any Internal / External Authorities Informed?',
                    'Availability of CCTV Footage',
                    'EVIDENCE & ATTACHMENTS - Photos',
                    'EVIDENCE & ATTACHMENTS - Videos',
                    'Is further support or intervention required from Head Office / Management?',
                    'If YES, mention required support',
                    'Status'
                ]
            
            # Check if all expected columns are present
            missing_cols = set(expected_columns) - set(df.columns)
            if missing_cols:
                print(f"Warning: Missing columns: {missing_cols}")
            
            return df
        
        except Exception as e:
            raise Exception(f"Error reading Google Sheet: {str(e)}")
    
    def filter_by_date_range(self, df, start_date, end_date):
        """
        Filter dataframe by date range.
        
        Args:
            df (pd.DataFrame): The dataframe to filter
            start_date (datetime.date): Start date
            end_date (datetime.date): End date
            
        Returns:
            pd.DataFrame: Filtered dataframe
        """
        try:
            # Use correct date column based on sheet type
            date_column = 'Date' if self.sheet_type == 'monitoring' else 'Date of Incident'
            
            # Convert Date column to datetime
            df['Date_parsed'] = pd.to_datetime(df[date_column], errors='coerce')
            
            # Filter by date range
            mask = (df['Date_parsed'] >= pd.Timestamp(start_date)) & \
                   (df['Date_parsed'] <= pd.Timestamp(end_date))
            
            return df[mask]
        
        except Exception as e:
            print(f"Error filtering by date: {str(e)}")
            return df
    
    def filter_by_shift(self, df, shift):
        """
        Filter dataframe by shift.
        
        Args:
            df (pd.DataFrame): The dataframe to filter
            shift (str): Shift value to filter by
            
        Returns:
            pd.DataFrame: Filtered dataframe
        """
        if not shift or shift == "All":
            return df
        
        try:
            # Check if Shift column exists
            if 'Shift' in df.columns:
                return df[df['Shift'] == shift]
            else:
                print("Warning: 'Shift' column not found")
                return df
        
        except Exception as e:
            print(f"Error filtering by shift: {str(e)}")
            return df
    
    def filter_by_site(self, df, site):
        """
        Filter dataframe by site name.
        
        Args:
            df (pd.DataFrame): The dataframe to filter
            site (str): Site name to filter by
            
        Returns:
            pd.DataFrame: Filtered dataframe
        """
        if not site or site == "All":
            return df
        
        try:
            if 'Site Name' in df.columns:
                return df[df['Site Name'] == site]
            else:
                print("Warning: 'Site Name' column not found")
                return df
        
        except Exception as e:
            print(f"Error filtering by site: {str(e)}")
            return df
    
    def get_unique_sites(self, df):
        """
        Get list of unique site names from the dataframe.
        
        Args:
            df (pd.DataFrame): The dataframe
            
        Returns:
            list: List of unique site names
        """
        if 'Site Name' in df.columns:
            return sorted(df['Site Name'].dropna().unique().tolist())
        return []
    
    def get_unique_shifts(self, df):
        """
        Get list of unique shifts from the dataframe.
        
        Args:
            df (pd.DataFrame): The dataframe
            
        Returns:
            list: List of unique shifts
        """
        if 'Shift' in df.columns:
            return sorted(df['Shift'].dropna().unique().tolist())
        return []
