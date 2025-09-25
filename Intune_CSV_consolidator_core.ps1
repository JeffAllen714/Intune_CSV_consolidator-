<#
.SYNOPSIS
    Intune Hardware Hash CSV Consolidator - Terminal Edition
.DESCRIPTION
    A streamlined PowerShell script that consolidates multiple hardware hash CSV files
    into a single file for Intune/Autopilot deployment
.AUTHOR
    Jeffrey Allen
#>

# ============================================
#  CONFIGURATION & INITIALIZATION
# ============================================

Clear-Host
$host.UI.RawUI.WindowTitle = "Intune CSV Consolidator"

# Display welcome banner
function Show-Banner {
    Write-Host ""
    Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║   INTUNE CSV CONSOLIDATOR - TERMINAL      ║" -ForegroundColor Cyan
    Write-Host "  ║   Merge Hardware Hash Files with Ease     ║" -ForegroundColor Cyan
    Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

# ============================================
#  HELPER FUNCTIONS
# ============================================

function Write-Status {
    param(
        [string]$message,
        [string]$type = "Info"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    
    switch ($type) {
        "Success" { 
            Write-Host "  [$timestamp] ✓ " -NoNewline -ForegroundColor Green
            Write-Host $message -ForegroundColor Green
        }
        "Error" { 
            Write-Host "  [$timestamp] ✗ " -NoNewline -ForegroundColor Red
            Write-Host $message -ForegroundColor Red
        }
        "Warning" { 
            Write-Host "  [$timestamp] ⚠ " -NoNewline -ForegroundColor Yellow
            Write-Host $message -ForegroundColor Yellow
        }
        "Process" {
            Write-Host "  [$timestamp] ◆ " -NoNewline -ForegroundColor Cyan
            Write-Host $message
        }
        default { 
            Write-Host "  [$timestamp] • " -NoNewline -ForegroundColor Gray
            Write-Host $message
        }
    }
}

function Get-ValidatedInput {
    param(
        [string]$prompt,
        [scriptblock]$validation,
        [string]$errorMessage
    )
    
    while ($true) {
        Write-Host ""
        Write-Host "  $prompt" -ForegroundColor White
        Write-Host "  > " -NoNewline -ForegroundColor Cyan
        $userInput = Read-Host
        
        if ($validation.Invoke($userInput)) {
            return $userInput
        } else {
            Write-Status $errorMessage "Error"
        }
    }
}

function Show-ProgressBar {
    param(
        [int]$current,
        [int]$total,
        [string]$activity
    )
    
    $percentComplete = [math]::Round(($current / $total) * 100)
    $progressBar = "[" + ("=" * [math]::Floor($percentComplete / 2)) + (" " * (50 - [math]::Floor($percentComplete / 2))) + "]"
    
    Write-Host "`r  Processing: $progressBar $percentComplete% - $activity" -NoNewline -ForegroundColor Cyan
}

# ============================================
#  MAIN PROCESSING FUNCTIONS
# ============================================

function Get-CsvFiles {
    param(
        [string]$directoryPath,
        [bool]$includeSubdirectories
    )
    
    $searchParams = @{
        Path = $directoryPath
        Filter = "*.csv"
    }
    
    if ($includeSubdirectories) {
        $searchParams.Add("Recurse", $true)
    }
    
    return Get-ChildItem @searchParams
}

function Process-CsvConsolidation {
    param(
        [string]$companyName,
        [string]$sourceDirectory,
        [array]$csvFiles
    )
    
    $consolidatedData = @()
    $processedCount = 0
    $errorCount = 0
    $totalFiles = $csvFiles.Count
    
    Write-Host ""
    Write-Status "Starting consolidation process..." "Process"
    Write-Host ""
    
    foreach ($file in $csvFiles) {
        try {
            # Update progress
            $processedCount++
            Show-ProgressBar -current $processedCount -total $totalFiles -activity $file.Name
            
            # Import CSV data
            $csvData = Import-Csv -Path $file.FullName -ErrorAction Stop
            
            if ($csvData) {
                $consolidatedData += $csvData
            }
            
        } catch {
            $errorCount++
            Write-Host "" # New line after progress bar
            Write-Status "Failed to process: $($file.Name)" "Error"
        }
    }
    
    Write-Host "" # Clear progress line
    Write-Host ""
    
    return @{
        Data = $consolidatedData
        ProcessedCount = ($processedCount - $errorCount)
        ErrorCount = $errorCount
        TotalEntries = $consolidatedData.Count
    }
}

function Export-ConsolidatedCsv {
    param(
        [array]$data,
        [string]$companyName,
        [string]$outputDirectory
    )
    
    # Sanitize company name for filename
    $sanitizedCompanyName = $companyName -replace '[^\w\s-]', '' -replace '\s+', '_'
    
    # Generate timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
    
    # Build output filename
    $outputFileName = "${sanitizedCompanyName}_ConsolidatedHardwareHashes_${timestamp}.csv"
    $outputPath = Join-Path $outputDirectory $outputFileName
    
    # Export the consolidated data
    $data | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
    
    return $outputPath
}

function Check-Duplicates {
    param(
        [array]$data
    )
    
    $serialNumbers = $data | Where-Object { $_."Device Serial Number" } | Group-Object "Device Serial Number"
    $duplicates = $serialNumbers | Where-Object { $_.Count -gt 1 }
    
    if ($duplicates) {
        $duplicateCount = ($duplicates | Measure-Object -Property Count -Sum).Sum - $duplicates.Count
        return $duplicateCount
    }
    
    return 0
}

# ============================================
#  MAIN SCRIPT EXECUTION
# ============================================

function Start-CsvConsolidation {
    
    Show-Banner
    
    # Get company name
    $companyName = Get-ValidatedInput `
        -prompt "Enter Company Name:" `
        -validation { -not [string]::IsNullOrWhiteSpace($args[0]) } `
        -errorMessage "Company name cannot be empty"
    
    # Get source directory
    $sourceDirectory = Get-ValidatedInput `
        -prompt "Enter directory path containing CSV files:" `
        -validation { Test-Path $args[0] -PathType Container } `
        -errorMessage "Directory does not exist or is invalid"
    
    # Ask about subdirectories
    Write-Host ""
    Write-Host "  Include subdirectories? (y/n)" -ForegroundColor White
    Write-Host "  > " -NoNewline -ForegroundColor Cyan
    $includeSubdirs = (Read-Host).ToLower() -eq 'y'
    
    Write-Host ""
    Write-Status "Scanning for CSV files..." "Process"
    
    # Get CSV files
    $csvFiles = Get-CsvFiles -directoryPath $sourceDirectory -includeSubdirectories $includeSubdirs
    
    if ($csvFiles.Count -eq 0) {
        Write-Status "No CSV files found in the specified directory" "Error"
        return
    }
    
    Write-Status "Found $($csvFiles.Count) CSV file(s)" "Success"
    
    # Confirm processing
    Write-Host ""
    Write-Host "  Ready to consolidate $($csvFiles.Count) CSV files for $companyName" -ForegroundColor Yellow
    Write-Host "  Continue? (y/n)" -ForegroundColor White
    Write-Host "  > " -NoNewline -ForegroundColor Cyan
    $confirm = (Read-Host).ToLower()
    
    if ($confirm -ne 'y') {
        Write-Status "Operation cancelled by user" "Warning"
        return
    }
    
    # Process files
    $result = Process-CsvConsolidation `
        -companyName $companyName `
        -sourceDirectory $sourceDirectory `
        -csvFiles $csvFiles
    
    if ($result.TotalEntries -eq 0) {
        Write-Status "No valid data found in CSV files" "Error"
        return
    }
    
    # Check for duplicates
    $duplicateCount = Check-Duplicates -data $result.Data
    if ($duplicateCount -gt 0) {
        Write-Status "Found $duplicateCount duplicate serial number(s)" "Warning"
    }
    
    # Export consolidated file
    Write-Status "Saving consolidated file..." "Process"
    
    try {
        $outputPath = Export-ConsolidatedCsv `
            -data $result.Data `
            -companyName $companyName `
            -outputDirectory $sourceDirectory
        
        Write-Host ""
        Write-Host "  ╔═══════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "  ║           CONSOLIDATION COMPLETE          ║" -ForegroundColor Green
        Write-Host "  ╚═══════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Status "Files processed: $($result.ProcessedCount)" "Success"
        Write-Status "Total entries: $($result.TotalEntries)" "Success"
        if ($result.ErrorCount -gt 0) {
            Write-Status "Files with errors: $($result.ErrorCount)" "Warning"
        }
        Write-Host ""
        Write-Host "  Output saved to:" -ForegroundColor White
        Write-Host "  $outputPath" -ForegroundColor Cyan
        
    } catch {
        Write-Status "Failed to save consolidated file: $_" "Error"
        return
    }
}

# ============================================
#  SCRIPT ENTRY POINT
# ============================================

try {
    Start-CsvConsolidation
    
    Write-Host ""
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
} catch {
    Write-Host ""
    Write-Status "Unexpected error occurred: $_" "Error"
    Write-Host ""
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
