## OTF Attendance System
A Google Apps Script project to automate attendance tracking for Orange Tree Foundation (OTF) Zoom classes, specifically for college boys. Processes Zoom CSV reports to calculate student attendance based on join/leave times, matches against a master Google Sheet, and updates scores (1.0 for >25 min, 0.5 for 5-25 min, 0 for <5 min).
Features
⦁	Zoom CSV Processing: Aggregates multiple join/leave entries per student.
⦁	Matching Logic: Matches by OTF ID (normalized) first, then name, with case-insensitive and space-cleaned comparisons.
⦁	Google Sheets Integration: Updates attendance in a master sheet (college boys attendance 2025-2026) under specific month/date columns.
⦁	Sidebar UI: User-friendly interface for selecting Zoom CSV, month, and date via Google Picker API.
⦁	Free Tools Only: Built with Google Apps Script, Sheets, and Drive (no paid services).

Project Structure
⦁	code.gs: Main script with logic for reading Zoom CSV, normalizing identifiers, calculating attendance, and updating the master sheet.
⦁	Sidebar.html: HTML interface for selecting files and parameters.
⦁	appsscript.json (optional): Apps Script project manifest.

Setup Instructions
1.	git clone https://github.com/yourusername/OTF-Attendance-System.git
2.	Google Apps Script:
⦁	Create a new Apps Script project in Google Sheets (Extensions > Apps Script).
⦁	Copy code.gs and Sidebar.html into the project.
⦁	Update CONFIG.MASTER_SHEET_NAME in code.gs to match your Google Sheet name.
3.	Google Picker API:
⦁	Enable Google Picker and Drive APIs in Google Cloud Console.
⦁	Add your CLIENT_ID, DEVELOPER_KEY, and API_KEY in code.gs (see getFilePicker).
4.	Run:
⦁	Reload the Google Sheet to see the "OTF Attendance Helper" menu.
⦁	Use the sidebar to select a Zoom CSV, month, and date, then process.

Requirements
⦁	Google Account with access to Sheets and Apps Script.
⦁	Google Cloud Project with Picker and Drive APIs enabled (free tier).
⦁	Zoom CSV reports with Name and Duration columns.

Usage
⦁	Open the Google Sheet (college boys attendance 2025-2026).
⦁	Click "OTF Attendance Helper > Process Zoom Report".
⦁	Select a Zoom CSV from Google Drive, choose a month (e.g., "Oct") and date (e.g., "02"), and click "Process Attendance".
⦁	Attendance scores update in the corresponding date column.
