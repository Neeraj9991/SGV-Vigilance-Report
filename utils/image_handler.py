"""
Image Handler Module
Downloads images from Google Drive URLs and prepares them for embedding in reports.
"""

import requests
import re
import base64
from io import BytesIO
from PIL import Image
import os

class ImageHandler:
    def __init__(self):
        """Initialize the image handler."""
        self.temp_dir = "temp_images"
        os.makedirs(self.temp_dir, exist_ok=True)
    
    def extract_drive_file_id(self, url):
        """
        Extract file ID from Google Drive URL.
        
        Args:
            url (str): Google Drive URL
            
        Returns:
            str: File ID or None if not found
        """
        # Pattern for various Google Drive URL formats
        patterns = [
            r'/d/([a-zA-Z0-9_-]+)',  # /d/FILE_ID
            r'id=([a-zA-Z0-9_-]+)',   # id=FILE_ID
            r'/file/d/([a-zA-Z0-9_-]+)',  # /file/d/FILE_ID
        ]
        
        for pattern in patterns:
            match = re.search(pattern, url)
            if match:
                return match.group(1)
        
        return None
    
    def get_direct_download_url(self, drive_url):
        """
        Convert Google Drive sharing URL to direct download URL.
        
        Args:
            drive_url (str): Google Drive sharing URL
            
        Returns:
            str: Direct download URL
        """
        file_id = self.extract_drive_file_id(drive_url)
        if file_id:
            # Use Google Drive direct download format
            return f"https://drive.google.com/uc?export=download&id={file_id}"
        return None
    
    def download_image(self, url):
        """
        Download image from URL.
        
        Args:
            url (str): Image URL (can be Google Drive URL)
            
        Returns:
            bytes: Image data or None if failed
        """
        try:
            # If it's a Google Drive URL, convert it
            if 'drive.google.com' in url:
                download_url = self.get_direct_download_url(url)
                if not download_url:
                    print(f"Failed to extract file ID from: {url}")
                    return None
            else:
                download_url = url
            
            # Download the image
            response = requests.get(download_url, timeout=10)
            response.raise_for_status()
            
            # Verify it's an image
            try:
                Image.open(BytesIO(response.content))
                return response.content
            except Exception as e:
                print(f"Downloaded content is not a valid image: {str(e)}")
                return None
        
        except Exception as e:
            print(f"Error downloading image from {url}: {str(e)}")
            return None
    
    def image_to_base64(self, image_data):
        """
        Convert image data to base64 string for HTML embedding.
        
        Args:
            image_data (bytes): Image data
            
        Returns:
            str: Base64 encoded string
        """
        try:
            # Open image and convert to JPEG if needed
            img = Image.open(BytesIO(image_data))
            
            # Convert to RGB if necessary (for transparency)
            if img.mode in ('RGBA', 'LA', 'P'):
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                img = background
            
            # Resize if too large (max width 800px)
            max_width = 800
            if img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
            
            # Save to bytes
            buffered = BytesIO()
            img.save(buffered, format="JPEG", quality=85)
            img_bytes = buffered.getvalue()
            
            # Encode to base64
            return base64.b64encode(img_bytes).decode('utf-8')
        
        except Exception as e:
            print(f"Error converting image to base64: {str(e)}")
            return None
    
    def parse_image_urls(self, images_string):
        """
        Parse comma-separated image URLs from the Images column.
        
        Args:
            images_string (str): Comma-separated list of image URLs
            
        Returns:
            list: List of image URLs
        """
        if pd.isna(images_string) or not images_string:
            return []
        
        # Split by comma and clean up
        urls = [url.strip() for url in str(images_string).split(',')]
        return [url for url in urls if url]
    
    def download_and_encode_images(self, images_string):
        """
        Download images from URLs and encode them as base64.
        
        Args:
            images_string (str): Comma-separated list of image URLs
            
        Returns:
            list: List of dictionaries with 'url' and 'base64' keys
        """
        urls = self.parse_image_urls(images_string)
        encoded_images = []
        
        for url in urls:
            image_data = self.download_image(url)
            if image_data:
                base64_str = self.image_to_base64(image_data)
                if base64_str:
                    encoded_images.append({
                        'url': url,
                        'base64': base64_str
                    })
        
        return encoded_images
    
    def cleanup_temp_dir(self):
        """Clean up temporary image directory."""
        try:
            if os.path.exists(self.temp_dir):
                for file in os.listdir(self.temp_dir):
                    os.remove(os.path.join(self.temp_dir, file))
        except Exception as e:
            print(f"Error cleaning up temp directory: {str(e)}")

# Import pandas for isna check
import pandas as pd
