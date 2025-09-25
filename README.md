# Intune Hardware Hash CSV Consolidator

A PowerShell GUI tool that combines multiple hardware hash CSV files into one consolidated file for easy Intune/Autopilot deployment.

## Features
- Dark-themed, user-friendly interface
- Combines multiple CSV files without modifying originals
- Company name integration in output filename
- Progress tracking and error handling
- Duplicate hardware hash detection
- Timestamped output files

## Requirements
- Windows PC with PowerShell
- Hardware hash CSV files (typical Intune/Autopilot format)

## How to Use
1. Save the script as `IntuneCSVConsolidator.ps1`
2. Right-click and "Run with PowerShell"
3. Enter your company name
4. Browse to folder containing your CSV files
5. Choose output location (defaults to source folder)
6. Click "Process Files"

## Output
Creates a consolidated CSV file named:
`CompanyName_ConsolidatedHardwareHashes_YYYY-MM-DD_HHMM.csv`

Perfect for bulk uploading to Microsoft Intune!

---
*Streamline your Intune deployments - consolidate once, deploy everywhere.*
