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
    def __init__(self):
        """Initialize the sheets reader with configuration from environment variables."""
        self.sheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.sheet_gid = os.getenv('SHEET_GID', '0')
        
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
            
            # Ensure expected columns exist
            expected_columns = [
                'Timestamp', 'Date', 'Time', 'Site Name',
                'Documentation Check [Attendance Register]',
                'Documentation Check [Handling / Taking Over Register]',
                'Documentation Check [Visitor Log Register]',
                'Performance Check [Grooming]',
                'Performance Check [Alertness]',
                'Performance Check [Post Discipline]',
                'Performance Check [Overall Rating]',
                'Observation', 'Inspected By', 'Images', 'Email Address',
                'Shift', 'Incident Report', 'Action Taken'
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
            # Convert Date column to datetime
            df['Date_parsed'] = pd.to_datetime(df['Date'], errors='coerce')
            
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
