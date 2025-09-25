# Intune Hardware Hash CSV Consolidator
# A GUI tool to combine multiple hardware hash CSV files into one consolidated file

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intune Hardware Hash CSV Consolidator"
$form.Size = New-Object System.Drawing.Size(650, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
$form.ForeColor = [System.Drawing.Color]::White

# Company Name Section
$companyLabel = New-Object System.Windows.Forms.Label
$companyLabel.Location = New-Object System.Drawing.Point(20, 20)
$companyLabel.Size = New-Object System.Drawing.Size(150, 20)
$companyLabel.Text = "Company Name:"
$companyLabel.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($companyLabel)

$companyTextBox = New-Object System.Windows.Forms.TextBox
$companyTextBox.Location = New-Object System.Drawing.Point(20, 45)
$companyTextBox.Size = New-Object System.Drawing.Size(300, 25)
$companyTextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$companyTextBox.ForeColor = [System.Drawing.Color]::White
$companyTextBox.BorderStyle = "FixedSingle"
$form.Controls.Add($companyTextBox)

# Source Directory Section
$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Location = New-Object System.Drawing.Point(20, 85)
$sourceLabel.Size = New-Object System.Drawing.Size(200, 20)
$sourceLabel.Text = "Source Directory (CSV Files):"
$sourceLabel.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($sourceLabel)

$sourceTextBox = New-Object System.Windows.Forms.TextBox
$sourceTextBox.Location = New-Object System.Drawing.Point(20, 110)
$sourceTextBox.Size = New-Object System.Drawing.Size(450, 25)
$sourceTextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$sourceTextBox.ForeColor = [System.Drawing.Color]::White
$sourceTextBox.BorderStyle = "FixedSingle"
$form.Controls.Add($sourceTextBox)

$sourceBrowseButton = New-Object System.Windows.Forms.Button
$sourceBrowseButton.Location = New-Object System.Drawing.Point(480, 108)
$sourceBrowseButton.Size = New-Object System.Drawing.Size(80, 28)
$sourceBrowseButton.Text = "Browse..."
$sourceBrowseButton.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$sourceBrowseButton.ForeColor = [System.Drawing.Color]::White
$sourceBrowseButton.FlatStyle = "Flat"
$form.Controls.Add($sourceBrowseButton)

# Include Subdirectories Checkbox
$subdirsCheckbox = New-Object System.Windows.Forms.CheckBox
$subdirsCheckbox.Location = New-Object System.Drawing.Point(20, 145)
$subdirsCheckbox.Size = New-Object System.Drawing.Size(200, 20)
$subdirsCheckbox.Text = "Include subdirectories"
$subdirsCheckbox.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($subdirsCheckbox)

# Output Directory Section
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(20, 175)
$outputLabel.Size = New-Object System.Drawing.Size(200, 20)
$outputLabel.Text = "Output Directory:"
$outputLabel.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($outputLabel)

$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Location = New-Object System.Drawing.Point(20, 200)
$outputTextBox.Size = New-Object System.Drawing.Size(450, 25)
$outputTextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$outputTextBox.ForeColor = [System.Drawing.Color]::White
$outputTextBox.BorderStyle = "FixedSingle"
$form.Controls.Add($outputTextBox)

$outputBrowseButton = New-Object System.Windows.Forms.Button
$outputBrowseButton.Location = New-Object System.Drawing.Point(480, 198)
$outputBrowseButton.Size = New-Object System.Drawing.Size(80, 28)
$outputBrowseButton.Text = "Browse..."
$outputBrowseButton.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$outputBrowseButton.ForeColor = [System.Drawing.Color]::White
$outputBrowseButton.FlatStyle = "Flat"
$form.Controls.Add($outputBrowseButton)

# CSV Files Preview Section
$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Location = New-Object System.Drawing.Point(20, 240)
$previewLabel.Size = New-Object System.Drawing.Size(200, 20)
$previewLabel.Text = "CSV Files Found (0):"
$previewLabel.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($previewLabel)

$previewListBox = New-Object System.Windows.Forms.ListBox
$previewListBox.Location = New-Object System.Drawing.Point(20, 265)
$previewListBox.Size = New-Object System.Drawing.Size(540, 120)
$previewListBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$previewListBox.ForeColor = [System.Drawing.Color]::White
$previewListBox.BorderStyle = "FixedSingle"
$form.Controls.Add($previewListBox)

# Refresh Files Button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(570, 265)
$refreshButton.Size = New-Object System.Drawing.Size(60, 30)
$refreshButton.Text = "Refresh"
$refreshButton.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$refreshButton.ForeColor = [System.Drawing.Color]::White
$refreshButton.FlatStyle = "Flat"
$form.Controls.Add($refreshButton)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 395)
$progressBar.Size = New-Object System.Drawing.Size(440, 20)
$progressBar.Style = "Continuous"
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 425)
$statusLabel.Size = New-Object System.Drawing.Size(500, 40)
$statusLabel.Text = "Ready to consolidate CSV files..."
$statusLabel.ForeColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($statusLabel)

# Process Button
$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(480, 395)
$processButton.Size = New-Object System.Drawing.Size(100, 35)
$processButton.Text = "Process Files"
$processButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$processButton.ForeColor = [System.Drawing.Color]::White
$processButton.FlatStyle = "Flat"
$processButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($processButton)

# Open Output Button
$openOutputButton = New-Object System.Windows.Forms.Button
$openOutputButton.Location = New-Object System.Drawing.Point(480, 440)
$openOutputButton.Size = New-Object System.Drawing.Size(100, 30)
$openOutputButton.Text = "Open Output"
$openOutputButton.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$openOutputButton.ForeColor = [System.Drawing.Color]::White
$openOutputButton.FlatStyle = "Flat"
$openOutputButton.Enabled = $false
$form.Controls.Add($openOutputButton)

# Global variables
$script:lastOutputFile = ""

# Functions
function Update-FilePreview {
    $previewListBox.Items.Clear()
    
    if (-not [string]::IsNullOrWhiteSpace($sourceTextBox.Text) -and (Test-Path $sourceTextBox.Text)) {
        $searchOption = if ($subdirsCheckbox.Checked) { "AllDirectories" } else { "TopDirectoryOnly" }
        $csvFiles = Get-ChildItem -Path $sourceTextBox.Text -Filter "*.csv" -Recurse:$subdirsCheckbox.Checked
        
        foreach ($file in $csvFiles) {
            $relativePath = if ($subdirsCheckbox.Checked) { 
                $file.FullName.Replace($sourceTextBox.Text, "").TrimStart("\")
            } else { 
                $file.Name 
            }
            $previewListBox.Items.Add($relativePath)
        }
        
        $previewLabel.Text = "CSV Files Found ($($csvFiles.Count)):"
        
        # Auto-set output directory if empty
        if ([string]::IsNullOrWhiteSpace($outputTextBox.Text)) {
            $outputTextBox.Text = $sourceTextBox.Text
        }
    } else {
        $previewLabel.Text = "CSV Files Found (0):"
    }
}

function Show-FolderDialog {
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select folder"
    $folderDialog.ShowNewFolderButton = $true
    
    if ($folderDialog.ShowDialog() -eq "OK") {
        return $folderDialog.SelectedPath
    }
    return $null
}

function Validate-Inputs {
    $errors = @()
    
    if ([string]::IsNullOrWhiteSpace($companyTextBox.Text)) {
        $errors += "Company name is required"
    }
    
    if ([string]::IsNullOrWhiteSpace($sourceTextBox.Text)) {
        $errors += "Source directory is required"
    } elseif (-not (Test-Path $sourceTextBox.Text)) {
        $errors += "Source directory does not exist"
    }
    
    if ([string]::IsNullOrWhiteSpace($outputTextBox.Text)) {
        $errors += "Output directory is required"
    } elseif (-not (Test-Path $outputTextBox.Text)) {
        $errors += "Output directory does not exist"
    }
    
    if ($previewListBox.Items.Count -eq 0) {
        $errors += "No CSV files found in source directory"
    }
    
    return $errors
}

function Process-CSVFiles {
    try {
        $progressBar.Visible = $true
        $progressBar.Value = 0
        $processButton.Enabled = $false
        
        # Get all CSV files
        $csvFiles = Get-ChildItem -Path $sourceTextBox.Text -Filter "*.csv" -Recurse:$subdirsCheckbox.Checked
        $totalFiles = $csvFiles.Count
        $consolidatedData = @()
        $processedCount = 0
        
        $statusLabel.Text = "Processing CSV files..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Yellow
        
        foreach ($file in $csvFiles) {
            try {
                $statusLabel.Text = "Processing: $($file.Name)"
                $form.Refresh()
                
                # Import CSV data
                $csvData = Import-Csv -Path $file.FullName
                
                if ($csvData) {
                    $consolidatedData += $csvData
                    $processedCount++
                }
                
                # Update progress
                $progress = [math]::Round(($processedCount / $totalFiles) * 100)
                $progressBar.Value = $progress
                
            } catch {
                $statusLabel.Text = "Error processing $($file.Name): $($_.Exception.Message)"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
                Start-Sleep 2
            }
        }
        
        if ($consolidatedData.Count -gt 0) {
            # Create output filename
            $companyName = $companyTextBox.Text -replace '[^\w\s-]', '' -replace '\s+', '_'
            $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
            $outputFileName = "${companyName}_ConsolidatedHardwareHashes_${timestamp}.csv"
            $outputPath = Join-Path $outputTextBox.Text $outputFileName
            
            # Check for duplicates
            $duplicateCheck = $consolidatedData | Group-Object "Hardware Hash" | Where-Object { $_.Count -gt 1 }
            if ($duplicateCheck) {
                $duplicateCount = ($duplicateCheck | Measure-Object -Property Count -Sum).Sum - $duplicateCheck.Count
                [System.Windows.Forms.MessageBox]::Show(
                    "Warning: Found $duplicateCount duplicate hardware hash(es). All entries will be included in the output file.",
                    "Duplicates Found",
                    "OK",
                    "Warning"
                )
            }
            
            # Export consolidated data
            $statusLabel.Text = "Saving consolidated file..."
            $consolidatedData | Export-Csv -Path $outputPath -NoTypeInformation
            
            $script:lastOutputFile = $outputPath
            $openOutputButton.Enabled = $true
            
            $statusLabel.Text = "Success! Consolidated $processedCount files with $($consolidatedData.Count) total entries."
            $statusLabel.ForeColor = [System.Drawing.Color]::LightGreen
            
            [System.Windows.Forms.MessageBox]::Show(
                "Successfully consolidated $processedCount CSV files into:`n$outputPath`n`nTotal entries: $($consolidatedData.Count)",
                "Processing Complete",
                "OK",
                "Information"
            )
        } else {
            $statusLabel.Text = "No valid data found in CSV files."
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
        }
        
    } catch {
        $statusLabel.Text = "Error: $($_.Exception.Message)"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred during processing:`n$($_.Exception.Message)",
            "Processing Error",
            "OK",
            "Error"
        )
    } finally {
        $progressBar.Visible = $false
        $processButton.Enabled = $true
    }
}

# Event Handlers
$sourceBrowseButton.Add_Click({
    $folder = Show-FolderDialog
    if ($folder) {
        $sourceTextBox.Text = $folder
        Update-FilePreview
    }
})

$outputBrowseButton.Add_Click({
    $folder = Show-FolderDialog
    if ($folder) {
        $outputTextBox.Text = $folder
    }
})

$sourceTextBox.Add_TextChanged({
    Update-FilePreview
})

$subdirsCheckbox.Add_CheckedChanged({
    Update-FilePreview
})

$refreshButton.Add_Click({
    Update-FilePreview
})

$processButton.Add_Click({
    $errors = Validate-Inputs
    
    if ($errors.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please fix the following issues:`n`n" + ($errors -join "`n"),
            "Validation Error",
            "OK",
            "Warning"
        )
        return
    }
    
    # Confirm before processing
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Ready to consolidate $($previewListBox.Items.Count) CSV files for $($companyTextBox.Text).`n`nProceed?",
        "Confirm Processing",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        Process-CSVFiles
    }
})

$openOutputButton.Add_Click({
    if ($script:lastOutputFile -and (Test-Path $script:lastOutputFile)) {
        Start-Process "explorer.exe" -ArgumentList "/select,`"$($script:lastOutputFile)`""
    }
})

# Initialize form
$statusLabel.Text = "Ready to consolidate CSV files... Select source directory to begin."

# Show the form
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.ShowDialog()
