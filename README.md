# SGV Vigilance Report Generator

A professional Streamlit web application for generating PDF inspection reports from Google Forms data.

![Status](https://img.shields.io/badge/status-active-success.svg)
![Version](https://img.shields.io/badge/version-1.0-blue.svg)

## 📋 Features

- ✅ **Read from Public Google Sheets** - No API credentials needed
- 📅 **Flexible Filtering** - Filter by date range, shift, and site
- 📊 **Data Preview** - Review filtered data before generating reports
- 📄 **Professional PDF Reports** - Beautiful, branded inspection reports
- 🖼️ **Image Embedding** - Automatically downloads and embeds images from Google Drive
- 🎨 **Color-Coded Status** - Visual indicators for compliance and performance ratings
- ⬇️ **Easy Download** - Generate and download PDF reports instantly
- 💻 **Windows Compatible** - Uses xhtml2pdf for seamless Windows operation

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- Public Google Sheet with inspection data
- Google Drive images with public access

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/Neeraj9991/SGV-Vigilance-Report.git
   cd SGV-Vigilance-Report
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment**
   ```bash
   cp .env.example .env
   ```
   
   Edit `.env` and add your Google Sheet ID:
   ```env
   GOOGLE_SHEET_ID=your_sheet_id_here
   SHEET_GID=0
   ```

4. **Run the application**
   ```bash
   streamlit run app.py
   ```

The app will open in your browser at `http://localhost:8501`

## ⚙️ Configuration

### Getting Your Google Sheet ID

Your Google Sheet URL looks like this:
```
https://docs.google.com/spreadsheets/d/1a2b3c4d5e6f7g8h9i0j/edit
```

The Sheet ID is the part between `/d/` and `/edit`:
```
1a2b3c4d5e6f7g8h9i0j
```

### Making Your Sheet Public

1. Open your Google Sheet
2. Click **Share** (top right)
3. Click **Change to anyone with the link**
4. Set permission to **Viewer**
5. Click **Copy link** and then **Done**

### Making Drive Images Accessible

For images to appear in reports, they must be publicly accessible:

1. Open Google Drive
2. Right-click on the image
3. Select **Share**
4. Click **Change to anyone with the link**
5. Ensure permission is set to **Viewer**

## 📊 Expected Sheet Format

Your Google Sheet should have these columns (as created by your Google Form):

| Column Name | Description |
|-------------|-------------|
| Timestamp | Form submission timestamp |
| Date | Inspection date |
| Time | Inspection time |
| Site Name | Name of the inspected site |
| Documentation Check [Attendance Register] | Compliant/Non-Compliant |
| Documentation Check [Handling / Taking Over Register] | Compliant/Non-Compliant |
| Documentation Check [Visitor Log Register] | Compliant/Non-Compliant |
| Performance Check [Grooming] | Good/Average/Poor |
| Performance Check [Alertness] | Good/Average/Poor |
| Performance Check [Post Discipline] | Good/Average/Poor |
| Performance Check [Overall Rating] | Good/Average/Poor |
| Observation | Text observations |
| Inspected By | Inspector name |
| Images | Comma-separated Google Drive URLs |
| Email Address | Inspector email |
| Shift | Shift identifier |
| Incident Report | Incident details (if any) |
| Action Taken | Actions taken (if any) |

## 🎯 Usage

1. **Load Data**: Click "Load/Refresh Data" in the sidebar
2. **Apply Filters**: 
   - Select date range
   - Choose shift (or "All")
   - Choose site (or "All")
3. **Preview Data**: Review filtered inspection records
4. **Generate PDF**: Click "Generate PDF Report"
5. **Download**: Save the PDF to your device

## 🛠️ Troubleshooting

### Error: "GOOGLE_SHEET_ID not found"

- Make sure you've created a `.env` file (copy from `.env.example`)
- Verify the Sheet ID is correctly copied
- Don't include quotes around the Sheet ID

### Error: "Error reading Google Sheet"

- Verify your sheet is set to public (anyone with link can view)
- Check that the Sheet ID is correct
- Ensure your internet connection is active

### Images Not Appearing in PDF

- Make sure Google Drive images are publicly accessible
- Verify the Images column contains valid Google Drive URLs
- URLs should be separated by commas if multiple images

### PDF Generation is Slow

- This is normal when downloading multiple images
- Large images take longer to process
- Consider reducing image sizes in Google Drive

## 📁 Project Structure

```
SGV-Vigilance-Report/
├── app.py                      # Main Streamlit application
├── requirements.txt            # Python dependencies
├── .env.example               # Environment template
├── .gitignore                 # Git ignore rules
├── README.md                  # This file
├── templates/
│   └── report_template.html  # PDF report HTML template
├── utils/
│   ├── __init__.py
│   ├── sheets_reader.py       # Google Sheets data reader
│   ├── image_handler.py       # Image download & encoding
│   └── pdf_generator.py       # PDF generation logic
└── assets/
    └── sgv_logo.png           # Company logo
```

## 🔒 Security Notes

- Never commit your `.env` file to version control
- Keep your Google Sheet ID private if it contains sensitive data
- Regularly review who has access to your Google Sheet
- Consider using a service account for production deployments

## 🤝 Support

For technical support or questions:
- Contact your IT department
- Email: support@sgvsecurity.com
- Raise an issue on GitHub

## 📝 License

Copyright © 2026 SGV SUPER SECURITY SERVICE PVT. LTD. All rights reserved.

---

**Made with ❤️ by SGV IT Department**
