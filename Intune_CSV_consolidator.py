#!/usr/bin/env python3
"""
Intune Hardware Hash CSV Consolidator - Python Edition
A streamlined tool for consolidating multiple hardware hash CSV files
for Intune/Autopilot deployment
"""

import os
import sys
import csv
import glob
import time
import shutil
from pathlib import Path
from datetime import datetime
from collections import Counter
from typing import List, Dict, Tuple, Optional

# ============================================
#  UTILITY FUNCTIONS
# ============================================

def isRunningInIDLE():
    """Check if script is running in IDLE"""
    import sys
    return 'idlelib.run' in sys.modules

def clearScreen():
    """Clear the terminal screen"""
    os.system('cls' if os.name == 'nt' else 'clear')

# ============================================
#  COLOR CONFIGURATION CLASS
# ============================================

class Colors:
    """
    ANSI color codes for terminal output styling.
    
    This class provides a centralized collection of ANSI escape codes
    for colorizing terminal output. These codes work across most modern
    terminals on Windows, macOS, and Linux systems.
    
    Attributes:
        CYAN: Bright cyan color for highlights and prompts
        GREEN: Success messages and completion indicators
        YELLOW: Warning messages and important notices
        RED: Error messages and failure indicators
        WHITE: Primary text and important labels
        GRAY: Secondary text and subtle information
        BOLD: Bold text formatting
        RESET: Reset all formatting to terminal defaults
    
    Usage:
        print(f"{Colors.GREEN}Success!{Colors.RESET}")
    """
    # Disable colors if running in IDLE
    if isRunningInIDLE():
        CYAN = ''
        GREEN = ''
        YELLOW = ''
        RED = ''
        WHITE = ''
        GRAY = ''
        BOLD = ''
        RESET = ''
    else:
        CYAN = '\033[96m'
        GREEN = '\033[92m'
        YELLOW = '\033[93m'
        RED = '\033[91m'
        WHITE = '\033[97m'
        GRAY = '\033[90m'
        BOLD = '\033[1m'
        RESET = '\033[0m'

# ============================================
#  CSV CONSOLIDATOR CLASS
# ============================================

class CsvConsolidator:
    """
    A class to handle the consolidation of multiple CSV files into a single output file.
    
    This class provides all the functionality needed to find, read, process, and merge
    CSV files from a specified directory. It handles various CSV formats, validates data,
    detects duplicates, and creates a properly formatted output file.
    
    Attributes:
        companyName: The company name used in the output filename
        sourceDirectory: Path to the directory containing source CSV files
        includeSubdirectories: Whether to search subdirectories for CSV files
        consolidatedData: List of all consolidated CSV data rows
        allHeaders: Set of all unique headers found across CSV files
        csvFiles: List of Path objects for found CSV files
        
    Methods:
        run(): Main execution method that orchestrates the entire consolidation process
        showBanner(): Displays the application welcome banner
        getUserInputs(): Collects and validates user inputs
        findCsvFiles(): Locates all CSV files in the specified directory
        consolidateFiles(): Reads and merges all CSV files
        checkDuplicates(): Identifies duplicate entries in the consolidated data
        exportResults(): Saves the consolidated data to a new CSV file
        
    Usage:
        consolidator = CsvConsolidator()
        consolidator.run()
    """
    
    def __init__(self):
        """Initialize the CSV Consolidator with default values"""
        self.companyName = ""
        self.sourceDirectory = ""
        self.includeSubdirectories = False
        self.consolidatedData = []
        self.allHeaders = set()
        self.csvFiles = []
        self.processedCount = 0
        self.errorCount = 0
        self.errorFiles = []
        
    def showBanner(self):
        """Display the welcome banner with appropriate formatting for the environment"""
        print()
        if isRunningInIDLE():
            # Simple ASCII for IDLE
            print(f"  +===========================================+")
            print(f"  |   INTUNE CSV CONSOLIDATOR - PYTHON       |")
            print(f"  |   Merge Hardware Hash Files with Ease    |")
            print(f"  +===========================================+")
        else:
            # Unicode box drawing for terminals
            print(f"{Colors.CYAN}  ╔═══════════════════════════════════════════╗{Colors.RESET}")
            print(f"{Colors.CYAN}  ║   INTUNE CSV CONSOLIDATOR - PYTHON        ║{Colors.RESET}")
            print(f"{Colors.CYAN}  ║   Merge Hardware Hash Files with Ease     ║{Colors.RESET}")
            print(f"{Colors.CYAN}  ╚═══════════════════════════════════════════╝{Colors.RESET}")
        print()
    
    def writeStatus(self, message: str, statusType: str = "Info"):
        """
        Write a formatted status message with timestamp and icon
        
        Args:
            message: The status message to display
            statusType: Type of status (Success, Error, Warning, Process, Info)
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if isRunningInIDLE():
            # Simple ASCII icons for IDLE
            statusConfig = {
                "Success": (Colors.GREEN, "[OK]"),
                "Error": (Colors.RED, "[ERROR]"),
                "Warning": (Colors.YELLOW, "[WARN]"),
                "Process": (Colors.CYAN, "[PROC]"),
                "Info": (Colors.GRAY, "[INFO]")
            }
        else:
            # Unicode icons for terminals
            statusConfig = {
                "Success": (Colors.GREEN, "✓"),
                "Error": (Colors.RED, "✗"),
                "Warning": (Colors.YELLOW, "⚠"),
                "Process": (Colors.CYAN, "◆"),
                "Info": (Colors.GRAY, "•")
            }
        
        color, icon = statusConfig.get(statusType, (Colors.GRAY, "•"))
        print(f"  [{timestamp}] {color}{icon} {message}{Colors.RESET}")
    
    def showProgressBar(self, current: int, total: int, activity: str = ""):
        """
        Display a progress bar for file processing
        
        Args:
            current: Current item number being processed
            total: Total number of items to process
            activity: Current activity description
        """
        percentComplete = int((current / total) * 100) if total > 0 else 0
        barLength = 50
        filledLength = int(barLength * current // total)
        
        bar = "=" * filledLength + " " * (barLength - filledLength)
        
        if isRunningInIDLE():
            # IDLE doesn't handle \r well, so just print dots
            if current == 1:
                print("  Processing: ", end="")
            print(".", end="", flush=True)
            if current == total:
                print(f" Complete! ({total} files)")
        else:
            # Use \r to overwrite the same line in terminal
            print(f"\r  Processing: [{Colors.CYAN}{bar}{Colors.RESET}] {percentComplete}% - {activity}", end="")
            
            if current == total:
                print()  # New line when complete
    
    def getValidatedInput(self, prompt: str, validationFunc, errorMessage: str) -> str:
        """
        Get validated input from user with error handling
        
        Args:
            prompt: The prompt to display to the user
            validationFunc: Function to validate the input
            errorMessage: Error message to show if validation fails
            
        Returns:
            Validated user input string
        """
        while True:
            print()
            print(f"  {Colors.WHITE}{prompt}{Colors.RESET}")
            if isRunningInIDLE():
                userInput = input("  > ").strip()
            else:
                print(f"  {Colors.CYAN}> {Colors.RESET}", end="")
                userInput = input().strip()
            
            if validationFunc(userInput):
                return userInput
            else:
                self.writeStatus(errorMessage, "Error")
    
    def getYesNoInput(self, prompt: str) -> bool:
        """
        Get a yes/no response from user
        
        Args:
            prompt: The yes/no question to ask
            
        Returns:
            Boolean True for yes, False for no
        """
        print()
        print(f"  {Colors.WHITE}{prompt} (y/n){Colors.RESET}")
        if isRunningInIDLE():
            response = input("  > ").strip().lower()
        else:
            print(f"  {Colors.CYAN}> {Colors.RESET}", end="")
            response = input().strip().lower()
        return response == 'y' or response == 'yes'
    
    def validateCompanyName(self, name: str) -> bool:
        """Validate that company name is not empty"""
        return bool(name and name.strip())
    
    def validateDirectory(self, path: str) -> bool:
        """Validate that directory exists"""
        return os.path.exists(path) and os.path.isdir(path)
    
    def getUserInputs(self):
        """Collect all necessary user inputs with validation"""
        # Get company name
        self.companyName = self.getValidatedInput(
            "Enter Company Name:",
            self.validateCompanyName,
            "Company name cannot be empty"
        )
        
        # Get source directory
        self.sourceDirectory = self.getValidatedInput(
            "Enter directory path containing CSV files:",
            self.validateDirectory,
            "Directory does not exist or is invalid"
        )
        
        # Ask about subdirectories
        self.includeSubdirectories = self.getYesNoInput("Include subdirectories?")
    
    def findCsvFiles(self):
        """
        Find all CSV files in the specified directory
        
        Returns:
            Boolean indicating if any CSV files were found
        """
        print()
        self.writeStatus("Scanning for CSV files...", "Process")
        
        if self.includeSubdirectories:
            # Recursive search
            pattern = os.path.join(self.sourceDirectory, "**", "*.csv")
            self.csvFiles = [Path(f) for f in glob.glob(pattern, recursive=True)]
        else:
            # Top-level only
            pattern = os.path.join(self.sourceDirectory, "*.csv")
            self.csvFiles = [Path(f) for f in glob.glob(pattern)]
        
        self.csvFiles = sorted(self.csvFiles)
        
        if not self.csvFiles:
            self.writeStatus("No CSV files found in the specified directory", "Error")
            return False
        
        self.writeStatus(f"Found {len(self.csvFiles)} CSV file(s)", "Success")
        
        # Show preview of files
        if len(self.csvFiles) <= 5:
            for csvFile in self.csvFiles:
                print(f"    • {csvFile.name}")
        else:
            for csvFile in self.csvFiles[:3]:
                print(f"    • {csvFile.name}")
            print(f"    ... and {len(self.csvFiles) - 3} more")
        
        return True
    
    def readCsvFile(self, filePath: Path) -> Tuple[List[Dict], List[str], bool]:
        """
        Read a CSV file and return its data, headers, and success status
        
        Args:
            filePath: Path to the CSV file to read
            
        Returns:
            Tuple of (data_rows, headers, success_boolean)
        """
        try:
            with open(filePath, 'r', encoding='utf-8-sig', newline='') as csvFile:
                # Read the entire file content
                content = csvFile.read()
                
                # Check if file is empty
                if not content.strip():
                    return [], [], False
                
                # Reset to beginning
                csvFile.seek(0)
                
                # Try to detect the dialect
                try:
                    dialect = csv.Sniffer().sniff(content[:1024])
                except:
                    dialect = csv.excel  # Default to Excel dialect
                
                # First, try to read as a regular CSV with headers
                csvFile.seek(0)
                reader = csv.DictReader(csvFile, dialect=dialect)
                
                # Check if we have valid headers
                if reader.fieldnames:
                    data = list(reader)
                    
                    # If no data rows (only header), try reading as headerless CSV
                    if not data:
                        csvFile.seek(0)
                        reader = csv.reader(csvFile, dialect=dialect)
                        rows = list(reader)
                        
                        if rows:
                            # Generate generic headers based on number of columns
                            numColumns = len(rows[0])
                            headers = [f"Column{i+1}" for i in range(numColumns)]
                            
                            # Convert rows to dictionaries
                            data = []
                            for row in rows:
                                if row:  # Skip empty rows
                                    rowDict = {}
                                    for i, value in enumerate(row[:numColumns]):
                                        rowDict[headers[i]] = value
                                    data.append(rowDict)
                            
                            return data, headers, True
                    else:
                        return data, reader.fieldnames, True
                
                # If no fieldnames, treat as headerless CSV
                csvFile.seek(0)
                reader = csv.reader(csvFile, dialect=dialect)
                rows = list(reader)
                
                if rows:
                    # Generate generic headers
                    numColumns = len(rows[0])
                    headers = [f"Column{i+1}" for i in range(numColumns)]
                    
                    # Convert to dictionaries
                    data = []
                    for row in rows:
                        if row:  # Skip empty rows
                            rowDict = {}
                            for i, value in enumerate(row[:numColumns]):
                                rowDict[headers[i]] = value
                            data.append(rowDict)
                    
                    return data, headers, True
                
                return [], [], False
                
        except Exception as e:
            # For debugging - you can uncomment this line to see specific errors
            # print(f"\n    Debug: Error reading {filePath.name}: {str(e)}")
            return [], [], False
    
    def consolidateFiles(self):
        """
        Consolidate all CSV files into a single dataset
        
        Returns:
            Boolean indicating if consolidation was successful
        """
        totalFiles = len(self.csvFiles)
        
        print()
        self.writeStatus("Starting consolidation process...", "Process")
        print()
        
        for index, csvFile in enumerate(self.csvFiles, 1):
            # Update progress
            fileName = csvFile.name
            self.showProgressBar(index, totalFiles, fileName[:30] + "..." if len(fileName) > 30 else fileName)
            
            # Read CSV file
            data, headers, success = self.readCsvFile(csvFile)
            
            if success and data:
                self.consolidatedData.extend(data)
                self.allHeaders.update(headers)
                self.processedCount += 1
            else:
                self.errorCount += 1
                self.errorFiles.append(csvFile.name)
        
        print()  # New line after progress bar
        print()
        
        # Show error files if any
        if self.errorFiles:
            self.writeStatus(f"Failed to process {len(self.errorFiles)} file(s):", "Warning")
            for errorFile in self.errorFiles[:5]:  # Show first 5 errors
                print(f"    - {errorFile}")
            if len(self.errorFiles) > 5:
                print(f"    ... and {len(self.errorFiles) - 5} more")
        
        return len(self.consolidatedData) > 0
    
    def checkDuplicates(self) -> int:
        """
        Check for duplicate serial numbers in the consolidated data
        
        Returns:
            Number of duplicate entries found
        """
        serialNumbers = []
        
        for row in self.consolidatedData:
            # Try different possible column names for serial number
            serialNumber = row.get('Device Serial Number') or \
                          row.get('Serial Number') or \
                          row.get('DeviceSerialNumber')
            if serialNumber:
                serialNumbers.append(serialNumber)
        
        if serialNumbers:
            counter = Counter(serialNumbers)
            duplicates = {k: v for k, v in counter.items() if v > 1}
            duplicateCount = sum(duplicates.values()) - len(duplicates)
            return duplicateCount
        
        return 0
    
    def exportResults(self) -> Optional[str]:
        """
        Export consolidated data to a new CSV file
        
        Returns:
            Path to the output file if successful, None otherwise
        """
        try:
            # Sanitize company name for filename
            sanitizedCompanyName = "".join(c for c in self.companyName if c.isalnum() or c in (' ', '-', '_'))
            sanitizedCompanyName = sanitizedCompanyName.replace(' ', '_')
            
            # Generate timestamp
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
            
            # Build output filename
            outputFileName = f"{sanitizedCompanyName}_ConsolidatedHardwareHashes_{timestamp}.csv"
            outputPath = os.path.join(self.sourceDirectory, outputFileName)
            
            # Write consolidated CSV
            with open(outputPath, 'w', encoding='utf-8-sig', newline='') as csvFile:
                writer = csv.DictWriter(csvFile, fieldnames=sorted(self.allHeaders))
                writer.writeheader()
                writer.writerows(self.consolidatedData)
            
            return outputPath
            
        except Exception as e:
            self.writeStatus(f"Failed to save file: {str(e)}", "Error")
            return None
    
    def displaySummary(self, outputPath: str):
        """
        Display the final summary of the consolidation process
        
        Args:
            outputPath: Path where the consolidated file was saved
        """
        print()
        if isRunningInIDLE():
            print(f"  +===========================================+")
            print(f"  |          CONSOLIDATION COMPLETE          |")
            print(f"  +===========================================+")
        else:
            print(f"{Colors.GREEN}  ╔═══════════════════════════════════════════╗{Colors.RESET}")
            print(f"{Colors.GREEN}  ║           CONSOLIDATION COMPLETE          ║{Colors.RESET}")
            print(f"{Colors.GREEN}  ╚═══════════════════════════════════════════╝{Colors.RESET}")
        print()
        
        self.writeStatus(f"Files processed: {self.processedCount}", "Success")
        self.writeStatus(f"Total entries: {len(self.consolidatedData)}", "Success")
        
        if self.errorCount > 0:
            self.writeStatus(f"Files with errors: {self.errorCount}", "Warning")
        
        print()
        print(f"  {Colors.WHITE}Output saved to:{Colors.RESET}")
        print(f"  {Colors.CYAN}{outputPath}{Colors.RESET}")
    
    def run(self):
        """
        Main execution method that orchestrates the entire consolidation process
        
        This method runs through the complete workflow:
        1. Shows banner
        2. Collects user inputs
        3. Finds CSV files
        4. Confirms with user
        5. Consolidates files
        6. Checks for duplicates
        7. Exports results
        8. Displays summary
        """
        try:
            # Clear screen and show banner
            clearScreen()
            self.showBanner()
            
            # Get user inputs
            self.getUserInputs()
            
            # Find CSV files
            if not self.findCsvFiles():
                input("\n  Press Enter to exit...")
                return
            
            # Confirm processing
            print()
            print(f"  {Colors.YELLOW}Ready to consolidate {len(self.csvFiles)} CSV files for {self.companyName}{Colors.RESET}")
            
            if not self.getYesNoInput("Continue?"):
                self.writeStatus("Operation cancelled by user", "Warning")
                input("\n  Press Enter to exit...")
                return
            
            # Consolidate files
            if not self.consolidateFiles():
                self.writeStatus("No valid data found in CSV files", "Error")
                input("\n  Press Enter to exit...")
                return
            
            # Check for duplicates
            duplicateCount = self.checkDuplicates()
            if duplicateCount > 0:
                self.writeStatus(f"Found {duplicateCount} duplicate serial number(s)", "Warning")
            
            # Export consolidated file
            self.writeStatus("Saving consolidated file...", "Process")
            outputPath = self.exportResults()
            
            if outputPath:
                self.displaySummary(outputPath)
            else:
                self.writeStatus("Failed to save consolidated file", "Error")
            
            print()
            print(f"  {Colors.GRAY}Press Enter to exit...{Colors.RESET}", end="")
            input()
            
        except KeyboardInterrupt:
            print()
            print()
            self.writeStatus("Operation cancelled by user", "Warning")
            sys.exit(0)
            
        except Exception as e:
            print()
            self.writeStatus(f"Unexpected error occurred: {str(e)}", "Error")
            print(f"\n  {Colors.GRAY}Press Enter to exit...{Colors.RESET}", end="")
            input()
            sys.exit(1)

# ============================================
#  MAIN
# ============================================

def main():
    """Main entry point - creates and runs the CSV consolidator"""
    consolidator = CsvConsolidator()
    consolidator.run()

# ============================================
#  SCRIPT ENTRY POINT
# ============================================

if __name__ == "__main__":
    main()
  
