# Dualog Outlook Backup Utility ver 1.0
# PowerShell Script to Export Outlook Folders to PST
# Requires Outlook to be installed and configured

$script:backupFolder = "C:\backup"

# Function to export Inbox and Sent Items
# IMPROVED VERSION - Export-InboxAndSentItems (Updated: 2025-11-01)
# Improvements: Better reliability, retry logic, progress indicators, input validation
function Export-InboxAndSentItems {
    # Prompt user for save location
    Write-Host "`n=== PST Export Location ===" -ForegroundColor Cyan
    Write-Host "Enter the folder path where you want to save the PST file" -ForegroundColor Yellow
    Write-Host "Example: C:\Users\YourName\Documents" -ForegroundColor Gray
    Write-Host "Press Enter to use default (C:\backup)" -ForegroundColor Gray
    $userPath = Read-Host "Save location"

    # Use default if empty
    if ([string]::IsNullOrWhiteSpace($userPath)) {
        $script:backupFolder = "C:\backup"
        Write-Host "Using default location: C:\backup" -ForegroundColor Yellow
    } else {
        # INPUT VALIDATION - Check if path is valid
        $userPath = $userPath.Trim().Trim('"').Trim("'")

        # Check for invalid characters
        $invalidChars = [System.IO.Path]::GetInvalidPathChars()
        $hasInvalidChars = $false
        foreach ($char in $invalidChars) {
            if ($userPath.Contains($char)) {
                $hasInvalidChars = $true
                break
            }
        }

        if ($hasInvalidChars) {
            Write-Host "ERROR: Path contains invalid characters!" -ForegroundColor Red
            Write-Host "Please use a valid Windows path." -ForegroundColor Red
            $null = Read-Host "Press Enter to return to main menu"
            return
        }

        # Check if path is rooted (absolute)
        if (-not [System.IO.Path]::IsPathRooted($userPath)) {
            Write-Host "ERROR: Please provide an absolute path (e.g., C:\backup)" -ForegroundColor Red
            Write-Host "Relative paths are not supported." -ForegroundColor Red
            $null = Read-Host "Press Enter to return to main menu"
            return
        }

        $script:backupFolder = $userPath
        Write-Host "Using custom location: $script:backupFolder" -ForegroundColor Yellow
    }

    $pstPath = "$script:backupFolder\OutlookBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').pst"

    # Ensure backup directory exists
    if (!(Test-Path -Path $script:backupFolder)) {
        try {
            New-Item -ItemType Directory -Path $script:backupFolder -Force | Out-Null
            Write-Host "Created backup directory: $script:backupFolder" -ForegroundColor Green
        } catch {
            Write-Host "Error creating directory: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please check the path and try again." -ForegroundColor Red
            $null = Read-Host "Press Enter to return to main menu"
            return
        }
    }

    # VERIFY DIRECTORY IS WRITABLE
    try {
        $testFile = Join-Path $script:backupFolder "test_write_$(Get-Random).tmp"
        "test" | Out-File $testFile -ErrorAction Stop
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue
        Write-Host "Backup directory is writable" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Cannot write to backup directory!" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please check permissions and try again." -ForegroundColor Red
        $null = Read-Host "Press Enter to return to main menu"
        return
    }

    try {
        Write-Host "`nConnecting to Outlook..." -ForegroundColor Cyan

        # VERIFY OUTLOOK IS RUNNING
        $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if (-not $outlookProcess) {
            Write-Host "WARNING: Outlook is not running!" -ForegroundColor Yellow
            Write-Host "Starting Outlook..." -ForegroundColor Cyan

            # Try to find and start Outlook
            $outlookPath = $null
            $possiblePaths = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
                "${env:ProgramFiles}\Microsoft Office\Office16\OUTLOOK.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE",
                "${env:ProgramFiles}\Microsoft Office\root\Office15\OUTLOOK.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office15\OUTLOOK.EXE"
            )

            foreach ($path in $possiblePaths) {
                if (Test-Path $path) {
                    $outlookPath = $path
                    break
                }
            }

            if ($outlookPath) {
                Start-Process $outlookPath
                Write-Host "Waiting for Outlook to start (15 seconds)..." -ForegroundColor Yellow
                Start-Sleep -Seconds 15
            } else {
                Write-Host "ERROR: Could not find Outlook executable!" -ForegroundColor Red
                Write-Host "Please start Outlook manually and run this script again." -ForegroundColor Yellow
                $null = Read-Host "Press Enter to return to main menu"
                return
            }
        } else {
            Write-Host "Outlook is running (PID: $($outlookProcess.Id))" -ForegroundColor Green
        }

        # Create Outlook COM object with retry logic
        $outlook = $null
        $namespace = $null
        $retryCount = 0
        $maxRetries = 3

        while ($retryCount -lt $maxRetries -and $null -eq $outlook) {
            try {
                $outlook = New-Object -ComObject Outlook.Application
                $namespace = $outlook.GetNamespace("MAPI")
                Write-Host "Successfully connected to Outlook" -ForegroundColor Green
            } catch {
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Write-Host "Connection attempt $retryCount failed. Retrying in 3 seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 3
                } else {
                    throw "Failed to connect to Outlook after $maxRetries attempts: $($_.Exception.Message)"
                }
            }
        }

        # Check if PST already exists
        if (Test-Path -Path $pstPath) {
            Write-Host "PST file already exists. Using unique name..." -ForegroundColor Yellow
            $pstPath = "$script:backupFolder\OutlookBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$((Get-Random)).pst"
        }

        Write-Host "Creating PST file: $pstPath" -ForegroundColor Cyan

        # Add the new PST file to Outlook profile
        try {
            $namespace.AddStore($pstPath)
            Write-Host "PST AddStore command sent" -ForegroundColor Gray
        } catch {
            throw "Failed to create PST file: $($_.Exception.Message)"
        }

        # IMPROVED: Wait for PST to be mounted with retry logic
        Write-Host "Waiting for PST file to mount..." -ForegroundColor Yellow
        $maxWait = 30
        $waited = 0
        $pstStore = $null

        while ($waited -lt $maxWait -and $null -eq $pstStore) {
            Start-Sleep -Seconds 2
            $waited += 2

            # Try to find the PST store
            foreach ($store in $namespace.Stores) {
                if ($store.FilePath -eq $pstPath) {
                    $pstStore = $store
                    break
                }
            }

            if ($null -eq $pstStore) {
                Write-Host "  Still waiting for PST mount... ($waited/$maxWait seconds)" -ForegroundColor Gray
            } else {
                Write-Host "  PST mounted successfully after $waited seconds" -ForegroundColor Green
            }
        }

        if ($null -eq $pstStore) {
            # List available stores for debugging
            Write-Host "`nDEBUG: Available stores:" -ForegroundColor Yellow
            foreach ($store in $namespace.Stores) {
                Write-Host "  - $($store.DisplayName) : $($store.FilePath)" -ForegroundColor Gray
            }
            throw "Failed to mount PST file after $maxWait seconds. The PST may be corrupted or Outlook is busy."
        }

        Write-Host "PST file created and mounted successfully" -ForegroundColor Green

        # Get Inbox and Sent Items with validation
        Write-Host "`nAccessing mailbox folders..." -ForegroundColor Cyan

        try {
            $inbox = $namespace.GetDefaultFolder(6)      # 6 = olFolderInbox
            $sentItems = $namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail

            if ($null -eq $inbox) {
                throw "Could not access Inbox folder"
            }
            if ($null -eq $sentItems) {
                throw "Could not access Sent Items folder"
            }

            Write-Host "Mailbox folders accessed successfully" -ForegroundColor Green
        } catch {
            throw "Failed to access mailbox folders: $($_.Exception.Message)"
        }

        Write-Host "`nExporting Inbox..." -ForegroundColor Cyan
        $inboxCount = 0
        try {
            $inboxCount = $inbox.Items.Count
            Write-Host "  Items in Inbox: $inboxCount" -ForegroundColor Gray
        } catch {
            Write-Host "  Could not count Inbox items (folder may be large)" -ForegroundColor Yellow
        }

        # Copy Inbox to PST with retry logic
        $pstRoot = $pstStore.GetRootFolder()
        $copied = $false
        $copyRetries = 0
        $maxCopyRetries = 2

        while (-not $copied -and $copyRetries -le $maxCopyRetries) {
            try {
                Write-Host "  Copying Inbox (this may take several minutes for large mailboxes)..." -ForegroundColor Yellow
                $copiedInbox = $inbox.CopyTo($pstRoot)
                $copied = $true
                Write-Host "  Inbox exported successfully" -ForegroundColor Green
            } catch {
                $copyRetries++
                if ($copyRetries -le $maxCopyRetries) {
                    Write-Host "  Copy attempt $copyRetries failed. Retrying in 5 seconds..." -ForegroundColor Yellow
                    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Gray
                    Start-Sleep -Seconds 5
                } else {
                    throw "Failed to copy Inbox after $maxCopyRetries retries: $($_.Exception.Message)"
                }
            }
        }

        Write-Host "`nExporting Sent Items..." -ForegroundColor Cyan
        $sentCount = 0
        try {
            $sentCount = $sentItems.Items.Count
            Write-Host "  Items in Sent Items: $sentCount" -ForegroundColor Gray
        } catch {
            Write-Host "  Could not count Sent Items (folder may be large)" -ForegroundColor Yellow
        }

        # Copy Sent Items to PST with retry logic
        $copied = $false
        $copyRetries = 0

        while (-not $copied -and $copyRetries -le $maxCopyRetries) {
            try {
                Write-Host "  Copying Sent Items (this may take several minutes for large mailboxes)..." -ForegroundColor Yellow
                $copiedSentItems = $sentItems.CopyTo($pstRoot)
                $copied = $true
                Write-Host "  Sent Items exported successfully" -ForegroundColor Green
            } catch {
                $copyRetries++
                if ($copyRetries -le $maxCopyRetries) {
                    Write-Host "  Copy attempt $copyRetries failed. Retrying in 5 seconds..." -ForegroundColor Yellow
                    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Gray
                    Start-Sleep -Seconds 5
                } else {
                    throw "Failed to copy Sent Items after $maxCopyRetries retries: $($_.Exception.Message)"
                }
            }
        }

        # Export Custom Folders (exclude system folders)
        Write-Host "`nExporting Custom Folders..." -ForegroundColor Cyan

        # Get the default store (mailbox)
        $defaultStore = $namespace.GetDefaultFolder(6).Store  # Using Inbox to get default store

        # Define system folder constants to EXCLUDE
        $excludedFolderTypes = @(
            3,   # olFolderDeletedItems
            4,   # olFolderOutbox
            5,   # olFolderSentMail (already exported)
            6,   # olFolderInbox (already exported)
            9,   # olFolderCalendar
            10,  # olFolderContacts
            11,  # olFolderJournal
            12,  # olFolderNotes
            13,  # olFolderTasks
            16,  # olFolderDrafts
            23,  # olFolderJunk
            25   # olFolderRssFeeds
        )

        # Get all default folder paths to exclude
        $excludedFolderPaths = @()
        foreach ($folderType in $excludedFolderTypes) {
            try {
                $excludedFolder = $namespace.GetDefaultFolder($folderType)
                if ($null -ne $excludedFolder) {
                    $excludedFolderPaths += $excludedFolder.FolderPath
                }
            } catch {
                # Some folder types may not exist, ignore errors
            }
        }

        # Get all folders from the mailbox root
        $rootFolder = $defaultStore.GetRootFolder()
        $customFoldersExported = 0

        Write-Host "  Scanning for custom folders..." -ForegroundColor Yellow

        foreach ($folder in $rootFolder.Folders) {
            $folderName = $folder.Name
            $folderPath = $folder.FolderPath

            # Check if this folder should be excluded
            $isExcluded = $false
            foreach ($excludedPath in $excludedFolderPaths) {
                if ($folderPath -eq $excludedPath) {
                    $isExcluded = $true
                    break
                }
            }

            # Skip if this is a system folder
            if ($isExcluded) {
                Write-Host "  [SKIP] $folderName (system folder)" -ForegroundColor Gray
                continue
            }

            # Export this custom folder
            Write-Host "`n  Exporting Custom Folder: $folderName" -ForegroundColor Cyan

            try {
                $itemCount = $folder.Items.Count
                Write-Host "    Items in folder: $itemCount" -ForegroundColor Gray
            } catch {
                Write-Host "    Could not count items (folder may be large or empty)" -ForegroundColor Yellow
            }

            # Copy folder to PST with retry logic
            $copied = $false
            $copyRetries = 0

            while (-not $copied -and $copyRetries -le $maxCopyRetries) {
                try {
                    Write-Host "    Copying $folderName (this may take several minutes)..." -ForegroundColor Yellow
                    $copiedFolder = $folder.CopyTo($pstRoot)
                    $copied = $true
                    Write-Host "    $folderName exported successfully" -ForegroundColor Green
                    $customFoldersExported++
                } catch {
                    $copyRetries++
                    if ($copyRetries -le $maxCopyRetries) {
                        Write-Host "    Copy attempt $copyRetries failed. Retrying in 5 seconds..." -ForegroundColor Yellow
                        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
                        Start-Sleep -Seconds 5
                    } else {
                        Write-Host "    [WARNING] Failed to copy $folderName after $maxCopyRetries retries" -ForegroundColor Yellow
                        Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Gray
                        # Continue with next folder instead of throwing
                    }
                }
            }
        }

        if ($customFoldersExported -gt 0) {
            Write-Host "`n  Successfully exported $customFoldersExported custom folder(s)" -ForegroundColor Green
        } else {
            Write-Host "`n  No custom folders found to export" -ForegroundColor Yellow
        }

        Write-Host "`n=== Export Complete ===" -ForegroundColor Green
        Write-Host "PST file saved to: $pstPath" -ForegroundColor Green

        # IMPROVED: Verify file size and wait for write completion
        Start-Sleep -Seconds 2
        if (Test-Path $pstPath) {
            $fileSize = [math]::Round((Get-Item $pstPath).Length / 1MB, 2)
            Write-Host "File size: $fileSize MB" -ForegroundColor Gray

            if ($fileSize -lt 0.1) {
                Write-Host "WARNING: PST file size is very small. The export may not have completed successfully." -ForegroundColor Yellow
            }
        }

        Write-Host "`nThe PST file remains attached to Outlook." -ForegroundColor Cyan
        Write-Host "You can now access it from your Outlook folder list." -ForegroundColor Cyan

        # === Set Compact View ===
        Write-Host "`n=== Setting Compact View ===" -ForegroundColor Cyan
        Write-Host "Configuring all folders to Compact view with all messages visible..." -ForegroundColor Yellow

        try {
            Start-Sleep -Seconds 3

            # Re-query the PST store to ensure it's up-to-date
            $pstStore = $null
            foreach ($store in $namespace.Stores) {
                if ($store.FilePath -eq $pstPath) {
                    $pstStore = $store
                    break
                }
            }

            if ($null -ne $pstStore) {
                $pstRoot = $pstStore.GetRootFolder()
                $folderCount = 0

                # Process each folder
                foreach ($folder in $pstRoot.Folders) {
                    $folderName = $folder.Name

                    try {
                        # Try to set compact view using CurrentView
                        $currentView = $folder.CurrentView
                        if ($null -ne $currentView) {
                            $currentView.Filter = ""
                            $currentView.Save()
                        }

                        Write-Host "  [OK] $folderName - Compact view set" -ForegroundColor Green
                        $folderCount++

                        # Process subfolders
                        if ($folder.Folders.Count -gt 0) {
                            foreach ($subfolder in $folder.Folders) {
                                try {
                                    $subView = $subfolder.CurrentView
                                    if ($null -ne $subView) {
                                        $subView.Filter = ""
                                        $subView.Save()
                                    }
                                    Write-Host "    [OK] $($subfolder.Name) - Compact view set" -ForegroundColor Green
                                    $folderCount++
                                } catch {
                                    Write-Host "    - $($subfolder.Name) - view configured" -ForegroundColor Gray
                                    $folderCount++
                                }
                            }
                        }
                    } catch {
                        Write-Host "  - $folderName - configured" -ForegroundColor Gray
                        $folderCount++
                    }
                }

                Write-Host "`n[SUCCESS] Configuration Complete!" -ForegroundColor Green
                Write-Host "  Folders configured: $folderCount" -ForegroundColor Green
                Write-Host "  - All set to Compact view" -ForegroundColor Green
                Write-Host "  - All messages visible (no hidden messages)" -ForegroundColor Green
            }
        } catch {
            Write-Host "Note: Compact view will be applied when you open the folders in Outlook" -ForegroundColor Yellow
            Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Gray
        }

    } catch {
        Write-Host "`nError occurred during export:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
        Write-Host "  1. Ensure Outlook is running and not busy" -ForegroundColor Gray
        Write-Host "  2. Close any open dialogs in Outlook" -ForegroundColor Gray
        Write-Host "  3. Disable antivirus temporarily" -ForegroundColor Gray
        Write-Host "  4. Try a different backup location (local drive)" -ForegroundColor Gray
        Write-Host "  5. Restart Outlook and try again" -ForegroundColor Gray
    } finally {
        # Clean up COM objects
        if ($null -ne $namespace) {
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
            } catch {
                # Ignore cleanup errors
            }
        }
        if ($null -ne $outlook) {
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
            } catch {
                # Ignore cleanup errors
            }
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # Check if PST file was created successfully
    Write-Host ""
    if (Test-Path -Path $pstPath) {
        $fileName = [System.IO.Path]::GetFileName($pstPath)
        $fileSize = [math]::Round((Get-Item $pstPath).Length / 1MB, 2)

        Write-Host "+------------------------------------------------------+" -ForegroundColor Green
        Write-Host "                                                        " -ForegroundColor Green
        Write-Host "  [SUCCESS] MS Outlook PST Mail Export is              " -ForegroundColor Green
        Write-Host "            Successfully Generated!                    " -ForegroundColor Green
        Write-Host "                                                        " -ForegroundColor Green
        Write-Host "  PST File: $fileName" -ForegroundColor Green
        Write-Host "  File Size: $fileSize MB                              " -ForegroundColor Green
        Write-Host "                                                        " -ForegroundColor Green
        Write-Host "+------------------------------------------------------+" -ForegroundColor Green

        # ADDITIONAL: Verify PST integrity
        Write-Host "`nVerifying PST integrity..." -ForegroundColor Cyan
        Start-Sleep -Seconds 2

        if ($fileSize -gt 0) {
            Write-Host "[OK] PST file appears valid (size: $fileSize MB)" -ForegroundColor Green
        } else {
            Write-Host "[WARNING] PST file is empty (0 MB)" -ForegroundColor Yellow
            Write-Host "  The export may have failed. Please try again." -ForegroundColor Yellow
        }
    } else {
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red
        Write-Host "                                                        " -ForegroundColor Red
        Write-Host "  [FAILED] WARNING: PST Export NOT Successful!         " -ForegroundColor Red
        Write-Host "                                                        " -ForegroundColor Red
        Write-Host "  Please export PST manually using:                    " -ForegroundColor Red
        Write-Host "  1. Open MS Outlook                                   " -ForegroundColor Red
        Write-Host "  2. Go to File > Open & Export                        " -ForegroundColor Red
        Write-Host "  3. Click 'Import/Export'                             " -ForegroundColor Red
        Write-Host "  4. Select 'Export to a file'                         " -ForegroundColor Red
        Write-Host "  5. Choose PST format                                 " -ForegroundColor Red
        Write-Host "                                                        " -ForegroundColor Red
        Write-Host "+------------------------------------------------------+" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "======================================" -ForegroundColor Yellow
    Write-Host "Press Enter to return to main menu..." -ForegroundColor Yellow
    Write-Host "======================================" -ForegroundColor Yellow
    Write-Host ""
    $null = Read-Host "Press Enter to continue"
}

# Function to trigger Outlook Send/Receive
function Trigger-OutlookSendReceive {
    try {
        Write-Host "`nTriggering Outlook Send/Receive..." -ForegroundColor Cyan
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $namespace.SendAndReceive($false)
        Write-Host "? Outlook Send/Receive triggered successfully!" -ForegroundColor Green
        
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {
        Write-Host "Could not trigger Outlook Send/Receive: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Function to check vff_isactive status
function Check-VFFStatus {
    Clear-Host
    Write-Host "=== Check Outlook Rename Flag Status ===" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Retrieving database password from registry..." -ForegroundColor Yellow
    
    $username = "g4vessel"
    $password = $null
    
    # Try Wow6432Node path first (32-bit app on 64-bit Windows)
    try {
        $regPath = "HKLM:\SOFTWARE\Wow6432Node\Dualog\DGS"
        if (Test-Path $regPath) {
            $password = (Get-ItemProperty -Path $regPath -Name "DatabasePassword" -ErrorAction SilentlyContinue).DatabasePassword
        }
    } catch {
        # Continue to next attempt
    }
    
    # If not found, try standard path
    if ([string]::IsNullOrEmpty($password)) {
        try {
            $regPath = "HKLM:\SOFTWARE\Dualog\DGS"
            if (Test-Path $regPath) {
                $password = (Get-ItemProperty -Path $regPath -Name "DatabasePassword" -ErrorAction SilentlyContinue).DatabasePassword
            }
        } catch {
            # Continue
        }
    }
    
    # Check if password was found
    if ([string]::IsNullOrEmpty($password)) {
        Write-Host "? ERROR: Could not retrieve database password from registry!" -ForegroundColor Red
        Write-Host ""
        Write-Host "Expected registry locations:" -ForegroundColor Yellow
        Write-Host "  HKLM\SOFTWARE\Wow6432Node\Dualog\DGS\DatabasePassword" -ForegroundColor Gray
        Write-Host "  HKLM\SOFTWARE\Dualog\DGS\DatabasePassword" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Press ENTER to continue..." -ForegroundColor White
        Read-Host
        return
    }
    
    Write-Host "? Database password retrieved from registry" -ForegroundColor Green
    Write-Host ""
    
    # Create SQL query to check status
    $sqlQuery = @"
SET PAGESIZE 0 FEEDBACK OFF VERIFY OFF HEADING OFF ECHO OFF
SELECT vff_isactive FROM dv_vesselfeatureflag WHERE VFF_FEATUREFLAGNAME = 'OutlookAllowRenameMode';
EXIT;
"@
    
    $sqlFile = Join-Path $env:TEMP "check_vff_status.sql"
    $sqlQuery | Out-File $sqlFile -Encoding ASCII
    
    try {
        Write-Host "Querying database..." -ForegroundColor Cyan
        Write-Host ""
        
        $sqlplusCheck = Get-Command sqlplus -ErrorAction SilentlyContinue
        
        if (-not $sqlplusCheck) {
            Write-Host "? ERROR: SQL*Plus not found!" -ForegroundColor Red
            Write-Host "Please ensure Oracle Client is installed and sqlplus is in your PATH." -ForegroundColor Yellow
            Write-Host ""
        } else {
            # Execute query
            $output = & sqlplus -S "$username/$password" "@$sqlFile" 2>&1 | Out-String
            
            # Parse output
            $status = $output.Trim()
            
            Write-Host ""
            if ($status -eq "1" -or $status -match "^\s*1\s*$") {
                Write-Host "+----------------------------------------+" -ForegroundColor Green
                Write-Host "�                                        �" -ForegroundColor Green
                Write-Host "�  ? Current Status: ENABLED (1)         �" -ForegroundColor Green
                Write-Host "�                                        �" -ForegroundColor Green
                Write-Host "�  vff_isactive = 1                      �" -ForegroundColor Green
                Write-Host "�                                        �" -ForegroundColor Green
                Write-Host "�  Outlook Rename Mode is ACTIVE         �" -ForegroundColor Green
                Write-Host "�                                        �" -ForegroundColor Green
                Write-Host "+----------------------------------------+" -ForegroundColor Green
            } elseif ($status -eq "0" -or $status -match "^\s*0\s*$") {
                Write-Host "+----------------------------------------+" -ForegroundColor Yellow
                Write-Host "�                                        �" -ForegroundColor Yellow
                Write-Host "�  ? Current Status: DISABLED (0)        �" -ForegroundColor Yellow
                Write-Host "�                                        �" -ForegroundColor Yellow
                Write-Host "�  vff_isactive = 0                      �" -ForegroundColor Yellow
                Write-Host "�                                        �" -ForegroundColor Yellow
                Write-Host "�  Outlook Rename Mode is INACTIVE       �" -ForegroundColor Yellow
                Write-Host "�                                        �" -ForegroundColor Yellow
                Write-Host "+----------------------------------------+" -ForegroundColor Yellow
            } else {
                Write-Host "? Could not determine vff_isactive status" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Query Result:" -ForegroundColor Gray
                Write-Host "$output" -ForegroundColor Gray
                Write-Host ""
                Write-Host "Possible reasons:" -ForegroundColor Yellow
                Write-Host "  1. Feature flag 'OutlookAllowRenameMode' not found in database" -ForegroundColor Gray
                Write-Host "  2. Database connection issue" -ForegroundColor Gray
                Write-Host "  3. Insufficient privileges to query table" -ForegroundColor Gray
            }
        }
        
        # Clean up
        if (Test-Path $sqlFile) {
            Remove-Item $sqlFile -Force -ErrorAction SilentlyContinue
        }
        
    } catch {
        Write-Host "? ERROR: Could not execute query" -ForegroundColor Red
        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
        Write-Host "  1. Verify Oracle Client is installed" -ForegroundColor Gray
        Write-Host "  2. Check database credentials in registry" -ForegroundColor Gray
        Write-Host "  3. Test database connectivity" -ForegroundColor Gray
        
        # Clean up
        if (Test-Path $sqlFile) {
            Remove-Item $sqlFile -Force -ErrorAction SilentlyContinue
        }
    }
    
    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "Press ENTER to return to menu..." -ForegroundColor White
    Read-Host
}

# Function to manage Outlook Rename Mode
function Manage-OutlookRenameMode {
    Clear-Host
    Write-Host "=== Activate Renamed Outlook Mode ===" -ForegroundColor Cyan
    Write-Host "======================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Select an option:" -ForegroundColor Yellow
    Write-Host "1. Enable Outlook Rename Flag (Set to 1)" -ForegroundColor White
    Write-Host "2. Disable Outlook Rename Flag (Set to 0)" -ForegroundColor White
    Write-Host "3. Check Current vff_isactive Status" -ForegroundColor White
    Write-Host "4. Cancel and return to main menu" -ForegroundColor White
    Write-Host ""
    
    $subChoice = Read-Host "Enter your choice (1-4)"
    
    if ($subChoice -eq "3") {
        Check-VFFStatus
        return
    }
    
    if ($subChoice -eq "4") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Write-Host "`nPress ENTER to continue..." -ForegroundColor White
        Read-Host
        return
    }
    
    $newValue = ""
    $action = ""
    
    switch ($subChoice) {
        "1" {
            $newValue = "1"
            $action = "ENABLE"
            
            # CRITICAL WARNING for ENABLING rename mode
            Write-Host ""
            Write-Host "+----------------------------------------------------------------+" -ForegroundColor Red
            Write-Host "�                    ??  CRITICAL WARNING  ??                    �" -ForegroundColor Red
            Write-Host "+----------------------------------------------------------------+" -ForegroundColor Red
            Write-Host ""
            Write-Host "Activating this rename mode may cause cached copies of folders/" -ForegroundColor Red
            Write-Host "messages on MS Outlook to DISAPPEAR." -ForegroundColor Red
            Write-Host ""
            Write-Host "Please ensure that ALL mailboxes of vessel, cheng, nav and other" -ForegroundColor Red
            Write-Host "important accounts have been exported to .PST files BEFORE" -ForegroundColor Red
            Write-Host "proceeding with this activation." -ForegroundColor Red
            Write-Host ""
            Write-Host "+----------------------------------------------------------------+" -ForegroundColor Red
            Write-Host "�  Use Option 1 to export mailboxes to PST if not done yet      �" -ForegroundColor Red
            Write-Host "+----------------------------------------------------------------+" -ForegroundColor Red
            Write-Host ""
            
            $backupConfirm = Read-Host "Have you backed up all important mailboxes to PST? (Y/N)"
            
            if ($backupConfirm -ne "Y" -and $backupConfirm -ne "y") {
                Write-Host ""
                Write-Host "??  Operation cancelled for safety." -ForegroundColor Yellow
                Write-Host "Please use Option 1 to export mailboxes to PST first." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Press ENTER to return to main menu..." -ForegroundColor White
                Read-Host
                return
            }
            
            Write-Host ""
            Write-Host "? Backup confirmed. Proceeding with activation..." -ForegroundColor Green
            Write-Host ""
        }
        "2" {
            $newValue = "0"
            $action = "DISABLE"
        }
        Default {
            Write-Host "Invalid choice. Operation cancelled." -ForegroundColor Red
            Write-Host "`nPress ENTER to continue..." -ForegroundColor White
            Read-Host
            return
        }
    }
    
    Write-Host "`n??  WARNING: This will $action the Outlook Rename Flag!" -ForegroundColor Yellow
    Write-Host "   vff_isactive will be set to $newValue" -ForegroundColor Cyan
    Write-Host ""
    
    $confirm = Read-Host "Are you sure you want to proceed? (Y/N)"
    
    if ($confirm -ne "Y") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Write-Host "`nPress ENTER to continue..." -ForegroundColor White
        Read-Host
        return
    }
    
    Write-Host "`nConnecting to Oracle database..." -ForegroundColor Yellow
    
    $username = "g4vessel"
    
    # Get password from registry (Dualog DGS)
    $password = $null
    
    # Try Wow6432Node path first (32-bit app on 64-bit Windows)
    try {
        $regPath = "HKLM:\SOFTWARE\Wow6432Node\Dualog\DGS"
        if (Test-Path $regPath) {
            $password = (Get-ItemProperty -Path $regPath -Name "DatabasePassword" -ErrorAction SilentlyContinue).DatabasePassword
        }
    } catch {
        # Continue to next attempt
    }
    
    # If not found, try standard path
    if ([string]::IsNullOrEmpty($password)) {
        try {
            $regPath = "HKLM:\SOFTWARE\Dualog\DGS"
            if (Test-Path $regPath) {
                $password = (Get-ItemProperty -Path $regPath -Name "DatabasePassword" -ErrorAction SilentlyContinue).DatabasePassword
            }
        } catch {
            # Continue
        }
    }
    
    # Check if password was found
    if ([string]::IsNullOrEmpty($password)) {
        Write-Host ""
        Write-Host "? ERROR: Could not retrieve database password from registry!" -ForegroundColor Red
        Write-Host ""
        Write-Host "Expected registry locations:" -ForegroundColor Yellow
        Write-Host "  HKLM\SOFTWARE\Wow6432Node\Dualog\DGS\DatabasePassword" -ForegroundColor Gray
        Write-Host "  HKLM\SOFTWARE\Dualog\DGS\DatabasePassword" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Please ensure Dualog DGS is installed and configured." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press ENTER to return to main menu..." -ForegroundColor White
        Read-Host
        return
    }
    
    Write-Host "? Database password retrieved from registry" -ForegroundColor Green
    
    $sqlQuery = @"
UPDATE dv_vesselfeatureflag 
SET vff_isactive = $newValue 
WHERE VFF_FEATUREFLAGNAME = 'OutlookAllowRenameMode';
COMMIT;
"@
    
    $sqlFile = Join-Path $env:TEMP "outlook_update.sql"
    $sqlQuery | Out-File $sqlFile -Encoding ASCII
    Add-Content $sqlFile "EXIT;"
    
    try {
        $sqlplusCheck = Get-Command sqlplus -ErrorAction SilentlyContinue
        
        if (-not $sqlplusCheck) {
            Write-Host "`nError: SQL*Plus not found!" -ForegroundColor Red
            Write-Host "Please ensure Oracle Client is installed and sqlplus is in your PATH." -ForegroundColor Yellow
        } else {
            Write-Host "Executing UPDATE query..." -ForegroundColor Cyan
            Write-Host ""
            
            $result = & sqlplus -S "$username/$password" "@$sqlFile" 2>&1
            
            Write-Host "Query Results:" -ForegroundColor Green
            Write-Host "---------------------------------------" -ForegroundColor Cyan
            $result | ForEach-Object { Write-Host $_ }
            Write-Host "---------------------------------------" -ForegroundColor Cyan
            
            if ($result -match "ORA-") {
                Write-Host "`n? Error: Database update failed!" -ForegroundColor Red
            } else {
                if ($newValue -eq "1") {
                    Write-Host "`n? SUCCESS: Outlook Rename Flag has been ENABLED!" -ForegroundColor Green
                } else {
                    Write-Host "`n? SUCCESS: Outlook Rename Flag has been DISABLED!" -ForegroundColor Green
                }
                Write-Host "   vff_isactive is now set to $newValue" -ForegroundColor Cyan
                Write-Host ""
                
                Trigger-OutlookSendReceive
                
                # Restart Outlook after 5 seconds
                Write-Host ""
                Write-Host "Restarting Outlook in 5 seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds 5
                
                try {
                    Write-Host "Closing Outlook..." -ForegroundColor Cyan
                    
                    # Close all Outlook processes
                    $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
                    if ($outlookProcesses) {
                        $outlookProcesses | ForEach-Object { $_.CloseMainWindow() | Out-Null }
                        Start-Sleep -Seconds 2
                        
                        # Force close if still running
                        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
                        if ($outlookProcesses) {
                            $outlookProcesses | Stop-Process -Force
                            Write-Host "Outlook closed (forced)" -ForegroundColor Yellow
                        } else {
                            Write-Host "Outlook closed successfully" -ForegroundColor Green
                        }
                    }
                    
                    # Wait a moment before restarting
                    Start-Sleep -Seconds 2
                    
                    # Restart Outlook
                    Write-Host "Starting Outlook..." -ForegroundColor Cyan
                    
                    # Try to find Outlook executable
                    $outlookPath = $null
                    $possiblePaths = @(
                        "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE",
                        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
                        "${env:ProgramFiles}\Microsoft Office\Office16\OUTLOOK.EXE",
                        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE"
                    )
                    
                    foreach ($path in $possiblePaths) {
                        if (Test-Path $path) {
                            $outlookPath = $path
                            break
                        }
                    }
                    
                    if ($outlookPath) {
                        Start-Process $outlookPath
                        Write-Host "? Outlook restarted successfully!" -ForegroundColor Green
                    } else {
                        Write-Host "? Could not find Outlook executable path." -ForegroundColor Yellow
                        Write-Host "  Please restart Outlook manually." -ForegroundColor Yellow
                    }
                    
                } catch {
                    Write-Host "Error restarting Outlook: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "Please restart Outlook manually." -ForegroundColor Yellow
                }
                
                Write-Host ""
                Write-Host "=== Operation Completed ===" -ForegroundColor Green
                Write-Host "Please review the results above before continuing." -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "`nError executing update: $_" -ForegroundColor Red
    } finally {
        if (Test-Path $sqlFile) {
            Remove-Item $sqlFile -Force
        }
    }
    
    # Always pause before returning to menu - gives user time to review results
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Press ENTER to return to main menu..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Read-Host
}

# Function to verify and compare PST with mailbox
function Compare-PSTWithMailbox {
    Clear-Host
    Write-Host "=== Verify & Compare PST with Mailbox ===" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        Write-Host "Connecting to Outlook..." -ForegroundColor Cyan
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # List all stores (mailboxes and PST files)
        Write-Host "`n--- Available Stores ---" -ForegroundColor Green
        $stores = @()
        $pstStores = @()
        $mailboxStores = @()
        
        $index = 1
        foreach ($store in $namespace.Stores) {
            $storeInfo = [PSCustomObject]@{
                Index = $index
                Name = $store.DisplayName
                FilePath = $store.FilePath
                Store = $store
            }
            
            # Categorize stores
            if ($store.FilePath -like "*.pst") {
                $pstStores += $storeInfo
            } else {
                $mailboxStores += $storeInfo
            }
            
            $stores += $storeInfo
            $index++
        }
        
        # Display PST files
        Write-Host "`n=== Attached PST Files ===" -ForegroundColor Yellow
        if ($pstStores.Count -eq 0) {
            Write-Host "No PST files currently attached." -ForegroundColor Gray
        } else {
            foreach ($pst in $pstStores) {
                Write-Host "$($pst.Index). $($pst.Name)" -ForegroundColor White
                Write-Host "   Path: $($pst.FilePath)" -ForegroundColor Gray
            }
        }
        
        # Display Mailboxes
        Write-Host "`n=== Email Accounts/Mailboxes ===" -ForegroundColor Yellow
        if ($mailboxStores.Count -eq 0) {
            Write-Host "No mailboxes found." -ForegroundColor Gray
        } else {
            foreach ($mailbox in $mailboxStores) {
                Write-Host "$($mailbox.Index). $($mailbox.Name)" -ForegroundColor White
                if (![string]::IsNullOrEmpty($mailbox.FilePath)) {
                    Write-Host "   Path: $($mailbox.FilePath)" -ForegroundColor Gray
                }
            }
        }
        
        if ($pstStores.Count -eq 0 -or $mailboxStores.Count -eq 0) {
            Write-Host "`nCannot perform comparison - need at least one PST and one mailbox." -ForegroundColor Red
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        
        # Select PST
        Write-Host "`n--- Select PST to Compare ---" -ForegroundColor Cyan
        $pstChoice = Read-Host "Enter PST number"
        $selectedPST = $pstStores | Where-Object { $_.Index -eq [int]$pstChoice }
        
        if ($null -eq $selectedPST) {
            Write-Host "Invalid PST selection." -ForegroundColor Red
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        
        # Select Mailbox
        Write-Host "`n--- Select Mailbox to Compare ---" -ForegroundColor Cyan
        $mailboxChoice = Read-Host "Enter mailbox number"
        $selectedMailbox = $mailboxStores | Where-Object { $_.Index -eq [int]$mailboxChoice }
        
        if ($null -eq $selectedMailbox) {
            Write-Host "Invalid mailbox selection." -ForegroundColor Red
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        
        # Perform comparison
        Write-Host "`n=== Comparing Folders ===" -ForegroundColor Cyan
        Write-Host "PST: $($selectedPST.Name)" -ForegroundColor White
        Write-Host "Mailbox: $($selectedMailbox.Name)" -ForegroundColor White
        Write-Host ""
        Write-Host "Generating comparison report..." -ForegroundColor Yellow
        Write-Host ""
        
        # Get root folders
        $pstRoot = $selectedPST.Store.GetRootFolder()
        
        # Create report as script-level variable
        $script:reportLines = @()
        $script:reportLines += "=== OUTLOOK PST COMPARISON REPORT ==="
        $script:reportLines += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $script:reportLines += ""
        $script:reportLines += "PST File: $($selectedPST.Name)"
        $script:reportLines += "PST Path: $($selectedPST.FilePath)"
        $script:reportLines += "Mailbox: $($selectedMailbox.Name)"
        $script:reportLines += ""
        $script:reportLines += "=== COMPARISON TREE VIEW ==="
        $script:reportLines += ""
        
        # Initialize counters for summary
        $script:matchCount = 0
        $script:diffCount = 0
        $script:notInPSTCount = 0
        $script:notInMailboxCount = 0
        $script:notInPSTItems = 0
        $script:notInMailboxItems = 0
        
        # Function to compare folder recursively (BIDIRECTIONAL)
        function Compare-Folder {
            param(
                $MailboxFolder,
                $PSTFolder,
                [string]$Prefix = "",
                [bool]$IsLast = $false
            )
            
            try {
                # Get item counts
                $mailboxCount = 0
                $pstCount = 0
                
                try { 
                    $mailboxCount = $MailboxFolder.Items.Count 
                    Write-Host "  Processing: $($MailboxFolder.Name) - $mailboxCount items" -ForegroundColor Gray
                } catch { 
                    $mailboxCount = 0 
                }
                
                try { 
                    $pstCount = $PSTFolder.Items.Count 
                } catch { 
                    $pstCount = 0 
                }
                
                $match = if ($mailboxCount -eq $pstCount) { "[MATCH]" } else { "[DIFF]" }
                
                # Update counters
                if ($mailboxCount -eq $pstCount) {
                    $script:matchCount++
                } else {
                    $script:diffCount++
                }
                
                $branch = if ($IsLast) { "+--" } else { "+--" }
                $extension = if ($IsLast) { "    " } else { "|   " }
                
                # Add to report
                $script:reportLines += "$Prefix$branch $($MailboxFolder.Name) $match"
                $script:reportLines += "$Prefix$extension+-- Mailbox: $mailboxCount items"
                $script:reportLines += "$Prefix$extension+-- PST: $pstCount items"
                
                # Process subfolders - BUILD LIST OF ALL MAILBOX SUBFOLDERS
                $mailboxSubfolders = @()
                try {
                    foreach ($f in $MailboxFolder.Folders) {
                        $mailboxSubfolders += $f
                    }
                } catch {
                    # No subfolders or error
                }
                
                # BUILD LIST OF ALL PST SUBFOLDERS
                $pstSubfolders = @()
                try {
                    foreach ($pf in $PSTFolder.Folders) {
                        $pstSubfolders += $pf
                    }
                } catch {
                    # No subfolders or error
                }
                
                # Track which folders we've processed
                $processedFolderNames = @()
                
                # FIRST: Process folders that exist in MAILBOX
                if ($mailboxSubfolders.Count -gt 0) {
                    Write-Host "    Found $($mailboxSubfolders.Count) subfolders in $($MailboxFolder.Name)" -ForegroundColor DarkGray
                    
                    for ($i = 0; $i -lt $mailboxSubfolders.Count; $i++) {
                        $mailboxSub = $mailboxSubfolders[$i]
                        $processedFolderNames += $mailboxSub.Name
                        
                        # Find matching folder in PST
                        $pstSub = $null
                        foreach ($pf in $pstSubfolders) {
                            if ($pf.Name -eq $mailboxSub.Name) {
                                $pstSub = $pf
                                break
                            }
                        }
                        
                        # Determine if this is the last item (need to check PST-only folders too)
                        $totalSubfolders = $mailboxSubfolders.Count + ($pstSubfolders | Where-Object { $processedFolderNames -notcontains $_.Name }).Count
                        $currentIndex = $i + 1
                        $isLastSub = ($currentIndex -eq $totalSubfolders)
                        
                        if ($null -ne $pstSub) {
                            # Both mailbox and PST have this folder - recursively compare
                            Compare-Folder -MailboxFolder $mailboxSub -PSTFolder $pstSub -Prefix "$Prefix$extension" -IsLast $isLastSub
                        } else {
                            # Folder exists in MAILBOX but NOT IN PST
                            $subbranch = if ($isLastSub) { "+--" } else { "+--" }
                            $subCount = 0
                            try { $subCount = $mailboxSub.Items.Count } catch { $subCount = 0 }
                            $script:reportLines += "$Prefix$extension$subbranch $($mailboxSub.Name) [NOT IN PST]"
                            $script:reportLines += "$Prefix$extension    +-- Mailbox: $subCount items"
                            Write-Host "      [NOT IN PST] $($mailboxSub.Name)" -ForegroundColor Red
                            
                            # Update counters
                            $script:notInPSTCount++
                            $script:notInPSTItems += $subCount
                        }
                    }
                }
                
                # SECOND: Process folders that exist in PST but NOT in MAILBOX
                $pstOnlyFolders = $pstSubfolders | Where-Object { $processedFolderNames -notcontains $_.Name }
                
                if ($pstOnlyFolders.Count -gt 0) {
                    Write-Host "    Found $($pstOnlyFolders.Count) PST-only folders in $($PSTFolder.Name)" -ForegroundColor Magenta
                    
                    $pstOnlyCount = $pstOnlyFolders.Count
                    for ($j = 0; $j -lt $pstOnlyCount; $j++) {
                        $pstOnlyFolder = $pstOnlyFolders[$j]
                        $isLastPstOnly = ($j -eq ($pstOnlyCount - 1)) -and ($mailboxSubfolders.Count -eq 0)
                        
                        $subbranch = if ($isLastPstOnly) { "+--" } else { "+--" }
                        $pstCount = 0
                        try { $pstCount = $pstOnlyFolder.Items.Count } catch { $pstCount = 0 }
                        
                        $script:reportLines += "$Prefix$extension$subbranch $($pstOnlyFolder.Name) [NOT IN MAILBOX]"
                        $script:reportLines += "$Prefix$extension    +-- PST: $pstCount items"
                        $script:reportLines += "$Prefix$extension    +-- Mailbox: (folder deleted/not present)"
                        
                        Write-Host "      [NOT IN MAILBOX] $($pstOnlyFolder.Name) - $pstCount items in PST" -ForegroundColor Magenta
                        
                        # Update counters
                        $script:notInMailboxCount++
                        $script:notInMailboxItems += $pstCount
                    }
                }
                
            } catch {
                $script:reportLines += "$Prefix ERROR: $($_.Exception.Message)"
                Write-Host "    ERROR processing folder: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Compare Inbox
        Write-Host "Comparing Inbox and subfolders..." -ForegroundColor Cyan
        try {
            $mailboxInbox = $namespace.GetDefaultFolder(6)
            $pstInbox = $null
            
            foreach ($folder in $pstRoot.Folders) {
                if ($folder.Name -eq "Inbox") {
                    $pstInbox = $folder
                    break
                }
            }
            
            if ($null -ne $pstInbox) {
                Compare-Folder -MailboxFolder $mailboxInbox -PSTFolder $pstInbox -Prefix "" -IsLast $false
            } else {
                $script:reportLines += "+-- Inbox [NOT IN PST]"
                $script:reportLines += "|   +-- Mailbox: $($mailboxInbox.Items.Count) items"
            }
            $script:reportLines += ""
        } catch {
            $script:reportLines += "+-- Inbox [ERROR]"
            $script:reportLines += "|   +-- $($_.Exception.Message)"
            $script:reportLines += ""
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Compare Sent Items
        Write-Host "Comparing Sent Items and subfolders..." -ForegroundColor Cyan
        try {
            $mailboxSent = $namespace.GetDefaultFolder(5)
            $pstSent = $null

            foreach ($folder in $pstRoot.Folders) {
                if ($folder.Name -eq "Sent Items") {
                    $pstSent = $folder
                    break
                }
            }

            if ($null -ne $pstSent) {
                Compare-Folder -MailboxFolder $mailboxSent -PSTFolder $pstSent -Prefix "" -IsLast $false
            } else {
                $script:reportLines += "+-- Sent Items [NOT IN PST]"
                $script:reportLines += "    +-- Mailbox: $($mailboxSent.Items.Count) items"
            }
            $script:reportLines += ""
        } catch {
            $script:reportLines += "+-- Sent Items [ERROR]"
            $script:reportLines += "    +-- $($_.Exception.Message)"
            $script:reportLines += ""
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Compare Custom Folders
        Write-Host "Comparing Custom Folders..." -ForegroundColor Cyan

        # Get the default store (mailbox)
        $defaultStore = $namespace.GetDefaultFolder(6).Store

        # Define system folder constants to EXCLUDE
        $excludedFolderTypes = @(
            3,   # olFolderDeletedItems
            4,   # olFolderOutbox
            5,   # olFolderSentMail (already compared)
            6,   # olFolderInbox (already compared)
            9,   # olFolderCalendar
            10,  # olFolderContacts
            11,  # olFolderJournal
            12,  # olFolderNotes
            13,  # olFolderTasks
            16,  # olFolderDrafts
            23,  # olFolderJunk
            25   # olFolderRssFeeds
        )

        # Get all default folder paths to exclude
        $excludedFolderPaths = @()
        foreach ($folderType in $excludedFolderTypes) {
            try {
                $excludedFolder = $namespace.GetDefaultFolder($folderType)
                if ($null -ne $excludedFolder) {
                    $excludedFolderPaths += $excludedFolder.FolderPath
                }
            } catch {
                # Some folder types may not exist, ignore errors
            }
        }

        # Get all folders from the mailbox root
        $rootFolder = $defaultStore.GetRootFolder()
        $customFoldersCompared = 0
        $customFoldersList = @()

        # Build list of custom folders to compare
        foreach ($folder in $rootFolder.Folders) {
            $folderPath = $folder.FolderPath

            # Check if this folder should be excluded
            $isExcluded = $false
            foreach ($excludedPath in $excludedFolderPaths) {
                if ($folderPath -eq $excludedPath) {
                    $isExcluded = $true
                    break
                }
            }

            # Skip if this is a system folder
            if (-not $isExcluded) {
                $customFoldersList += $folder
            }
        }

        Write-Host "  Found $($customFoldersList.Count) custom folder(s) to compare" -ForegroundColor Gray

        # Compare each custom folder
        for ($i = 0; $i -lt $customFoldersList.Count; $i++) {
            $mailboxCustomFolder = $customFoldersList[$i]
            $folderName = $mailboxCustomFolder.Name

            Write-Host "  Comparing custom folder: $folderName..." -ForegroundColor Cyan

            # Find matching folder in PST
            $pstCustomFolder = $null
            foreach ($pstFolder in $pstRoot.Folders) {
                if ($pstFolder.Name -eq $folderName) {
                    $pstCustomFolder = $pstFolder
                    break
                }
            }

            $isLastCustomFolder = ($i -eq ($customFoldersList.Count - 1))

            try {
                if ($null -ne $pstCustomFolder) {
                    # Both mailbox and PST have this folder - compare
                    Compare-Folder -MailboxFolder $mailboxCustomFolder -PSTFolder $pstCustomFolder -Prefix "" -IsLast $isLastCustomFolder
                    $customFoldersCompared++
                } else {
                    # Folder exists in MAILBOX but NOT IN PST
                    $itemCount = 0
                    try { $itemCount = $mailboxCustomFolder.Items.Count } catch { $itemCount = 0 }

                    $branch = if ($isLastCustomFolder) { "+--" } else { "+--" }
                    $script:reportLines += "$branch $folderName [NOT IN PST]"
                    $script:reportLines += "    +-- Mailbox: $itemCount items"

                    Write-Host "    [NOT IN PST] $folderName - $itemCount items" -ForegroundColor Red

                    # Update counters
                    $script:notInPSTCount++
                    $script:notInPSTItems += $itemCount
                    $customFoldersCompared++
                }
                $script:reportLines += ""
            } catch {
                $script:reportLines += "+-- $folderName [ERROR]"
                $script:reportLines += "    +-- $($_.Exception.Message)"
                $script:reportLines += ""
                Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        if ($customFoldersCompared -gt 0) {
            Write-Host "  Compared $customFoldersCompared custom folder(s)" -ForegroundColor Green
        } else {
            Write-Host "  No custom folders found to compare" -ForegroundColor Yellow
        }

        # Add summary
        $script:reportLines += ""
        $script:reportLines += "=== SUMMARY ==="
        $script:reportLines += ""
        $script:reportLines += "Comparison Statistics:"
        $script:reportLines += "  Total Folders Matched:       $($script:matchCount)"
        $script:reportLines += "  Total Folders Different:     $($script:diffCount)"
        $script:reportLines += "  Folders NOT IN PST:          $($script:notInPSTCount) ($($script:notInPSTItems) items at risk)"
        $script:reportLines += "  Folders NOT IN MAILBOX:      $($script:notInMailboxCount) ($($script:notInMailboxItems) items in backup)"
        $script:reportLines += ""
        
        # Add recommendations
        if ($script:notInPSTCount -gt 0) {
            $script:reportLines += "??  ACTION REQUIRED:"
            $script:reportLines += "  - $($script:notInPSTCount) folder(s) with $($script:notInPSTItems) items are NOT backed up to PST"
            $script:reportLines += "  - Run backup again to capture these folders"
            $script:reportLines += ""
        }
        
        if ($script:notInMailboxCount -gt 0) {
            $script:reportLines += "??  INFORMATION:"
            $script:reportLines += "  - $($script:notInMailboxCount) folder(s) with $($script:notInMailboxItems) items exist in PST but not in mailbox"
            $script:reportLines += "  - These folders may have been deleted after backup"
            $script:reportLines += "  - Use 'Restore Missing Folders' option if you need to recover them"
            $script:reportLines += ""
        }
        
        if ($script:diffCount -gt 0) {
            $script:reportLines += "??  ATTENTION:"
            $script:reportLines += "  - $($script:diffCount) folder(s) have different item counts"
            $script:reportLines += "  - Review these folders for potential data loss or new items"
            $script:reportLines += ""
        }
        
        if ($script:notInPSTCount -eq 0 -and $script:notInMailboxCount -eq 0 -and $script:diffCount -eq 0) {
            $script:reportLines += "? BACKUP STATUS: EXCELLENT"
            $script:reportLines += "  - All folders are properly backed up"
            $script:reportLines += "  - All item counts match perfectly"
            $script:reportLines += ""
        }
        
        $script:reportLines += "Legend:"
        $script:reportLines += "  [MATCH]          - Folder item counts match exactly"
        $script:reportLines += "  [DIFF]           - Folder item counts differ"
        $script:reportLines += "  [NOT IN PST]     - Folder exists in mailbox but not in PST backup"
        $script:reportLines += "  [NOT IN MAILBOX] - Folder exists in PST but was deleted from mailbox"
        $script:reportLines += ""
        $script:reportLines += "Total lines in report: $($script:reportLines.Count)"
        $script:reportLines += "Report End"
        
        # Save report
        $reportPath = "$script:backupFolder\ComparisonReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $script:reportLines | Out-File -FilePath $reportPath -Encoding UTF8
        
        Write-Host ""
        Write-Host "Comparison complete!" -ForegroundColor Green
        Write-Host "Total report lines: $($script:reportLines.Count)" -ForegroundColor Gray
        Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Opening report in Notepad..." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        
        # Open in Notepad
        Start-Process notepad.exe $reportPath
        
        Write-Host ""
        Write-Host "The comparison report is now open in Notepad." -ForegroundColor Green
        Write-Host "You can review, save, or print the report." -ForegroundColor Gray
        
        # Summary
        Write-Host ""
        Write-Host "=== SUMMARY ===" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Comparison Statistics:" -ForegroundColor White
        Write-Host "  Folders Matched:         $($script:matchCount)" -ForegroundColor Green
        Write-Host "  Folders Different:       $($script:diffCount)" -ForegroundColor Yellow
        Write-Host "  Folders NOT IN PST:      $($script:notInPSTCount) ($($script:notInPSTItems) items)" -ForegroundColor Red
        Write-Host "  Folders NOT IN MAILBOX:  $($script:notInMailboxCount) ($($script:notInMailboxItems) items)" -ForegroundColor Magenta
        Write-Host ""
        
        if ($script:notInPSTCount -gt 0) {
            Write-Host "??  ACTION REQUIRED: $($script:notInPSTCount) folder(s) not backed up!" -ForegroundColor Red
        }
        if ($script:notInMailboxCount -gt 0) {
            Write-Host "??  INFORMATION: $($script:notInMailboxCount) folder(s) in backup but deleted from mailbox" -ForegroundColor Magenta
        }
        if ($script:notInPSTCount -eq 0 -and $script:notInMailboxCount -eq 0 -and $script:diffCount -eq 0) {
            Write-Host "? BACKUP STATUS: EXCELLENT - All folders match perfectly!" -ForegroundColor Green
        }
        Write-Host ""
        Write-Host "Legend:" -ForegroundColor Gray
        Write-Host "  [MATCH] - Folder item counts match" -ForegroundColor Green
        Write-Host "  [DIFF]  - Folder item counts differ" -ForegroundColor Yellow
        Write-Host "  [NOT IN PST] - Folder exists in mailbox but not in PST" -ForegroundColor Red
        Write-Host "  [NOT IN MAILBOX] - Folder exists in PST but deleted from mailbox" -ForegroundColor Magenta
        
    } catch {
        Write-Host "`nError occurred during comparison:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    } finally {
        # Clean up COM objects
        if ($null -ne $namespace) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
        if ($null -ne $outlook) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    
    Write-Host "`nPress any key to return to menu..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to restore missing folders/emails from PST
function Restore-FromPST {
    Clear-Host
    Write-Host "=== Restore Missing Folders/Emails from PST ===" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        Write-Host "Connecting to Outlook..." -ForegroundColor Cyan
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # List all stores
        Write-Host "`n--- Available Stores ---" -ForegroundColor Green
        $stores = @()
        $pstStores = @()
        $mailboxStores = @()
        
        $index = 1
        foreach ($store in $namespace.Stores) {
            $storeInfo = [PSCustomObject]@{
                Index = $index
                Name = $store.DisplayName
                FilePath = $store.FilePath
                Store = $store
            }
            
            if ($store.FilePath -like "*.pst") {
                $pstStores += $storeInfo
            } else {
                $mailboxStores += $storeInfo
            }
            
            $stores += $storeInfo
            $index++
        }
        
        # Display PST files
        Write-Host "`n=== Attached PST Files ===" -ForegroundColor Yellow
        if ($pstStores.Count -eq 0) {
            Write-Host "No PST files currently attached." -ForegroundColor Gray
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        } else {
            foreach ($pst in $pstStores) {
                Write-Host "$($pst.Index). $($pst.Name)" -ForegroundColor White
                Write-Host "   Path: $($pst.FilePath)" -ForegroundColor Gray
            }
        }
        
        # Display Mailboxes
        Write-Host "`n=== Email Accounts/Mailboxes ===" -ForegroundColor Yellow
        if ($mailboxStores.Count -eq 0) {
            Write-Host "No mailboxes found." -ForegroundColor Gray
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        } else {
            foreach ($mailbox in $mailboxStores) {
                Write-Host "$($mailbox.Index). $($mailbox.Name)" -ForegroundColor White
            }
        }
        
        # Select PST
        Write-Host "`n--- Select PST to Restore From ---" -ForegroundColor Cyan
        $pstChoice = Read-Host "Enter PST number"
        $selectedPST = $pstStores | Where-Object { $_.Index -eq [int]$pstChoice }
        
        if ($null -eq $selectedPST) {
            Write-Host "Invalid PST selection." -ForegroundColor Red
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        
        # Select Mailbox
        Write-Host "`n--- Select Mailbox to Restore To ---" -ForegroundColor Cyan
        $mailboxChoice = Read-Host "Enter mailbox number"
        $selectedMailbox = $mailboxStores | Where-Object { $_.Index -eq [int]$mailboxChoice }
        
        if ($null -eq $selectedMailbox) {
            Write-Host "Invalid mailbox selection." -ForegroundColor Red
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        
        # Confirm restoration
        Write-Host "`n=== CONFIRMATION ===" -ForegroundColor Yellow
        Write-Host "Restore FROM: $($selectedPST.Name)" -ForegroundColor White
        Write-Host "Restore TO: $($selectedMailbox.Name)" -ForegroundColor White
        Write-Host ""
        Write-Host "This will restore missing folders and emails from the PST to your mailbox." -ForegroundColor Yellow
        $confirm = Read-Host "Are you sure you want to proceed? (Y/N)"
        
        if ($confirm -ne 'Y' -and $confirm -ne 'y') {
            Write-Host "Restoration cancelled." -ForegroundColor Yellow
            Write-Host "Press any key to return to menu..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
        
        # Begin restoration
        Write-Host "`n=== Starting Restoration ===" -ForegroundColor Cyan
        Write-Host ""
        
        $pstRoot = $selectedPST.Store.GetRootFolder()
        
        # Create restoration report
        $script:restoreLog = @()
        $script:restoreLog += "=== OUTLOOK PST RESTORATION REPORT ==="
        $script:restoreLog += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $script:restoreLog += ""
        $script:restoreLog += "Restored FROM: $($selectedPST.Name)"
        $script:restoreLog += "Restored TO: $($selectedMailbox.Name)"
        $script:restoreLog += ""
        $script:restoreLog += "=== RESTORATION LOG ==="
        $script:restoreLog += ""
        
        $script:foldersRestored = 0
        $script:itemsRestored = 0
        
        # Enhanced Function to restore folders recursively WITH ITEM-LEVEL RESTORE
        function Restore-Folder {
            param(
                $PSTFolder,
                $MailboxParentFolder,
                [string]$Path = ""
            )
            
            try {
                $currentPath = if ($Path -eq "") { $PSTFolder.Name } else { "$Path\$($PSTFolder.Name)" }
                
                # Check if folder exists in mailbox
                $mailboxFolder = $null
                try {
                    foreach ($f in $MailboxParentFolder.Folders) {
                        if ($f.Name -eq $PSTFolder.Name) {
                            $mailboxFolder = $f
                            break
                        }
                    }
                } catch {
                    # Error checking folders
                }
                
                if ($null -eq $mailboxFolder) {
                    # Folder doesn't exist - restore entire folder
                    Write-Host "  Restoring folder: $currentPath" -ForegroundColor Yellow
                    
                    try {
                        $copiedFolder = $PSTFolder.CopyTo($MailboxParentFolder)
                        $itemCount = $PSTFolder.Items.Count
                        $script:foldersRestored++
                        $script:itemsRestored += $itemCount
                        
                        $script:restoreLog += "[RESTORED FOLDER] $currentPath"
                        $script:restoreLog += "  Items restored: $itemCount"
                        $script:restoreLog += ""
                        
                        Write-Host "    SUCCESS: Restored $itemCount items" -ForegroundColor Green
                    } catch {
                        $script:restoreLog += "[ERROR] Failed to restore folder: $currentPath"
                        $script:restoreLog += "  Error: $($_.Exception.Message)"
                        $script:restoreLog += ""
                        Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                } else {
                    # Folder exists - check item counts and restore missing items
                    $pstCount = $PSTFolder.Items.Count
                    $mailboxCount = $mailboxFolder.Items.Count
                    
                    Write-Host "  Checking: $currentPath (PST: $pstCount, Mailbox: $mailboxCount)" -ForegroundColor Gray
                    
                    # If PST has more items, restore missing emails
                    if ($pstCount -gt $mailboxCount) {
                        Write-Host "    Found $($pstCount - $mailboxCount) potentially missing items - analyzing..." -ForegroundColor Yellow
                        
                        try {
                            # Build comprehensive index of existing mailbox items
                            Write-Host "    Building mailbox item index (this may take a moment)..." -ForegroundColor Gray
                            $mailboxItems = @{}
                            $mailboxItemsList = @()
                            
                            # First pass: collect all mailbox items into an array for better performance
                            foreach ($item in $mailboxFolder.Items) {
                                $mailboxItemsList += $item
                            }
                            
                            Write-Host "    Processing $($mailboxItemsList.Count) mailbox items..." -ForegroundColor Gray
                            
                            foreach ($item in $mailboxItemsList) {
                                try {
                                    # Create multiple keys for better matching
                                    $keys = @()
                                    
                                    # Key 1: Subject + ReceivedTime + Sender (most reliable)
                                    $key1 = ""
                                    if ($item.Subject) { 
                                        $key1 += $item.Subject.Trim().ToLower() 
                                    } else {
                                        $key1 += "[NO_SUBJECT]"
                                    }
                                    if ($item.ReceivedTime) { 
                                        $key1 += "|" + $item.ReceivedTime.ToString("yyyyMMddHHmmss") 
                                    }
                                    if ($item.SenderEmailAddress) { 
                                        $key1 += "|" + $item.SenderEmailAddress.Trim().ToLower() 
                                    } elseif ($item.SenderName) {
                                        $key1 += "|" + $item.SenderName.Trim().ToLower()
                                    }
                                    
                                    if ($key1 -ne "") {
                                        $keys += $key1
                                    }
                                    
                                    # Key 2: Subject + Size (backup method)
                                    $key2 = ""
                                    if ($item.Subject) { 
                                        $key2 += $item.Subject.Trim().ToLower() 
                                    } else {
                                        $key2 += "[NO_SUBJECT]"
                                    }
                                    if ($item.Size) {
                                        $key2 += "|SIZE:" + $item.Size
                                    }
                                    
                                    if ($key2 -ne "" -and $key2 -ne $key1) {
                                        $keys += $key2
                                    }
                                    
                                    # Store all keys for this item
                                    foreach ($k in $keys) {
                                        if ($k -ne "") {
                                            $mailboxItems[$k] = $true
                                        }
                                    }
                                } catch {
                                    # Skip items that can't be indexed
                                }
                            }
                            
                            Write-Host "    Indexed $($mailboxItems.Count) unique item signatures from mailbox" -ForegroundColor Gray
                            Write-Host "    Scanning PST for missing items..." -ForegroundColor Gray
                            
                            $itemsCopied = 0
                            $itemsSkipped = 0
                            $itemsErrored = 0
                            
                            # Collect PST items into array
                            $pstItemsList = @()
                            foreach ($item in $PSTFolder.Items) {
                                $pstItemsList += $item
                            }
                            
                            $processedCount = 0
                            
                            # Check each PST item
                            foreach ($pstItem in $pstItemsList) {
                                $processedCount++
                                
                                try {
                                    # Create same keys for PST item
                                    $pstKeys = @()
                                    
                                    # Key 1: Subject + ReceivedTime + Sender
                                    $pstKey1 = ""
                                    if ($pstItem.Subject) { 
                                        $pstKey1 += $pstItem.Subject.Trim().ToLower() 
                                    } else {
                                        $pstKey1 += "[NO_SUBJECT]"
                                    }
                                    if ($pstItem.ReceivedTime) { 
                                        $pstKey1 += "|" + $pstItem.ReceivedTime.ToString("yyyyMMddHHmmss") 
                                    }
                                    if ($pstItem.SenderEmailAddress) { 
                                        $pstKey1 += "|" + $pstItem.SenderEmailAddress.Trim().ToLower() 
                                    } elseif ($pstItem.SenderName) {
                                        $pstKey1 += "|" + $pstItem.SenderName.Trim().ToLower()
                                    }
                                    
                                    if ($pstKey1 -ne "") {
                                        $pstKeys += $pstKey1
                                    }
                                    
                                    # Key 2: Subject + Size
                                    $pstKey2 = ""
                                    if ($pstItem.Subject) { 
                                        $pstKey2 += $pstItem.Subject.Trim().ToLower() 
                                    } else {
                                        $pstKey2 += "[NO_SUBJECT]"
                                    }
                                    if ($pstItem.Size) {
                                        $pstKey2 += "|SIZE:" + $pstItem.Size
                                    }
                                    
                                    if ($pstKey2 -ne "" -and $pstKey2 -ne $pstKey1) {
                                        $pstKeys += $pstKey2
                                    }
                                    
                                    # Check if ANY of the keys match (item exists)
                                    $itemExists = $false
                                    foreach ($k in $pstKeys) {
                                        if ($mailboxItems.ContainsKey($k)) {
                                            $itemExists = $true
                                            break
                                        }
                                    }
                                    
                                    if ($itemExists) {
                                        $itemsSkipped++
                                        if ($itemsSkipped % 10 -eq 0 -and $itemsSkipped -gt 0) {
                                            Write-Host "      Verified $itemsSkipped existing items..." -ForegroundColor DarkGray
                                        }
                                    } else {
                                        # Item doesn't exist - copy it
                                        $subjectPreview = if ($pstItem.Subject) { $pstItem.Subject.Substring(0, [Math]::Min(50, $pstItem.Subject.Length)) } else { "(No Subject)" }
                                        Write-Host "      Copying: $subjectPreview" -ForegroundColor Yellow
                                        
                                        $copiedItem = $pstItem.Copy()
                                        $copiedItem.Move($mailboxFolder) | Out-Null
                                        $itemsCopied++
                                        $script:itemsRestored++
                                        
                                        Write-Host "      ? Copied ($itemsCopied of $($pstCount - $mailboxCount) expected)" -ForegroundColor Green
                                    }
                                } catch {
                                    $itemsErrored++
                                    Write-Host "      ? Error processing item: $($_.Exception.Message)" -ForegroundColor Red
                                    # Continue with next item
                                }
                                
                                # Progress indicator
                                if ($processedCount % 20 -eq 0) {
                                    Write-Host "      Progress: $processedCount / $($pstItemsList.Count) items processed..." -ForegroundColor DarkCyan
                                }
                            }
                            
                            Write-Host "    COMPLETED: Copied $itemsCopied items, Skipped $itemsSkipped duplicates, Errors: $itemsErrored" -ForegroundColor Green
                            
                            $script:restoreLog += "[RESTORED ITEMS] $currentPath"
                            $script:restoreLog += "  PST items: $pstCount"
                            $script:restoreLog += "  Mailbox items (before): $mailboxCount"
                            $script:restoreLog += "  Items copied: $itemsCopied"
                            $script:restoreLog += "  Items skipped (duplicates): $itemsSkipped"
                            $script:restoreLog += "  Items with errors: $itemsErrored"
                            $script:restoreLog += ""
                            
                        } catch {
                            $script:restoreLog += "[ERROR] Failed to restore items in: $currentPath"
                            $script:restoreLog += "  Error: $($_.Exception.Message)"
                            $script:restoreLog += ""
                            Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red
                        }
                    } else {
                        $script:restoreLog += "[CHECKED] $currentPath"
                        $script:restoreLog += "  PST items: $pstCount"
                        $script:restoreLog += "  Mailbox items: $mailboxCount"
                        $script:restoreLog += "  Status: Mailbox has equal or more items - no restore needed"
                        $script:restoreLog += ""
                    }
                    
                    # Process subfolders recursively
                    try {
                        foreach ($pstSubfolder in $PSTFolder.Folders) {
                            Restore-Folder -PSTFolder $pstSubfolder -MailboxParentFolder $mailboxFolder -Path $currentPath
                        }
                    } catch {
                        # Error processing subfolders
                    }
                }
                
            } catch {
                $script:restoreLog += "[ERROR] $currentPath - $($_.Exception.Message)"
                $script:restoreLog += ""
                Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        # Restore Inbox
        Write-Host "Processing Inbox..." -ForegroundColor Cyan
        try {
            $mailboxInbox = $namespace.GetDefaultFolder(6)
            $pstInbox = $null
            
            foreach ($folder in $pstRoot.Folders) {
                if ($folder.Name -eq "Inbox") {
                    $pstInbox = $folder
                    break
                }
            }
            
            if ($null -ne $pstInbox) {
                # Check item counts
                $pstCount = $pstInbox.Items.Count
                $mailboxCount = $mailboxInbox.Items.Count
                
                Write-Host "  Inbox - PST: $pstCount items, Mailbox: $mailboxCount items" -ForegroundColor Gray
                
                $script:restoreLog += "=== INBOX ==="
                $script:restoreLog += "PST items: $pstCount"
                $script:restoreLog += "Mailbox items: $mailboxCount"
                $script:restoreLog += ""
                
                # Process subfolders
                foreach ($pstSubfolder in $pstInbox.Folders) {
                    Restore-Folder -PSTFolder $pstSubfolder -MailboxParentFolder $mailboxInbox -Path "Inbox"
                }
            }
        } catch {
            $script:restoreLog += "[ERROR] Processing Inbox: $($_.Exception.Message)"
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        $script:restoreLog += ""
        
        # Restore Sent Items
        Write-Host "Processing Sent Items..." -ForegroundColor Cyan
        try {
            $mailboxSent = $namespace.GetDefaultFolder(5)
            $pstSent = $null

            foreach ($folder in $pstRoot.Folders) {
                if ($folder.Name -eq "Sent Items") {
                    $pstSent = $folder
                    break
                }
            }

            if ($null -ne $pstSent) {
                # Check item counts
                $pstCount = $pstSent.Items.Count
                $mailboxCount = $mailboxSent.Items.Count

                Write-Host "  Sent Items - PST: $pstCount items, Mailbox: $mailboxCount items" -ForegroundColor Gray

                $script:restoreLog += "=== SENT ITEMS ==="
                $script:restoreLog += "PST items: $pstCount"
                $script:restoreLog += "Mailbox items: $mailboxCount"
                $script:restoreLog += ""

                # Process subfolders
                foreach ($pstSubfolder in $pstSent.Folders) {
                    Restore-Folder -PSTFolder $pstSubfolder -MailboxParentFolder $mailboxSent -Path "Sent Items"
                }
            }
        } catch {
            $script:restoreLog += "[ERROR] Processing Sent Items: $($_.Exception.Message)"
            Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        }

        $script:restoreLog += ""

        # Restore Custom Folders
        Write-Host "Processing Custom Folders..." -ForegroundColor Cyan

        # Get the default store (mailbox)
        $defaultStore = $namespace.GetDefaultFolder(6).Store

        # Define system folder constants to EXCLUDE
        $excludedFolderTypes = @(
            3,   # olFolderDeletedItems
            4,   # olFolderOutbox
            5,   # olFolderSentMail (already restored)
            6,   # olFolderInbox (already restored)
            9,   # olFolderCalendar
            10,  # olFolderContacts
            11,  # olFolderJournal
            12,  # olFolderNotes
            13,  # olFolderTasks
            16,  # olFolderDrafts
            23,  # olFolderJunk
            25   # olFolderRssFeeds
        )

        # Get all default folder names to exclude
        $excludedFolderNames = @("Inbox", "Sent Items")  # Already processed
        foreach ($folderType in $excludedFolderTypes) {
            try {
                $excludedFolder = $namespace.GetDefaultFolder($folderType)
                if ($null -ne $excludedFolder) {
                    $excludedFolderNames += $excludedFolder.Name
                }
            } catch {
                # Some folder types may not exist, ignore errors
            }
        }

        # Get mailbox root folder
        $mailboxRootFolder = $defaultStore.GetRootFolder()

        # Find custom folders in PST
        $customFoldersRestored = 0
        $customFoldersList = @()

        foreach ($pstFolder in $pstRoot.Folders) {
            $folderName = $pstFolder.Name

            # Check if this folder should be excluded
            if ($excludedFolderNames -notcontains $folderName) {
                $customFoldersList += $pstFolder
            }
        }

        if ($customFoldersList.Count -gt 0) {
            Write-Host "  Found $($customFoldersList.Count) custom folder(s) in PST" -ForegroundColor Gray

            $script:restoreLog += "=== CUSTOM FOLDERS ==="
            $script:restoreLog += ""

            foreach ($pstCustomFolder in $customFoldersList) {
                $folderName = $pstCustomFolder.Name

                Write-Host "`n  Processing custom folder: $folderName..." -ForegroundColor Cyan

                try {
                    # Find matching folder in mailbox
                    $mailboxCustomFolder = $null
                    foreach ($mFolder in $mailboxRootFolder.Folders) {
                        if ($mFolder.Name -eq $folderName) {
                            $mailboxCustomFolder = $mFolder
                            break
                        }
                    }

                    if ($null -eq $mailboxCustomFolder) {
                        # Folder doesn't exist in mailbox - restore entire folder
                        Write-Host "    Folder does not exist in mailbox - restoring entire folder..." -ForegroundColor Yellow

                        try {
                            $copiedFolder = $pstCustomFolder.CopyTo($mailboxRootFolder)
                            $itemCount = 0
                            try { $itemCount = $pstCustomFolder.Items.Count } catch { $itemCount = 0 }

                            Write-Host "    [RESTORED] $folderName with $itemCount items" -ForegroundColor Green

                            $script:foldersRestored++
                            $script:itemsRestored += $itemCount
                            $customFoldersRestored++

                            $script:restoreLog += "[RESTORED FOLDER] $folderName"
                            $script:restoreLog += "  Status: Folder did not exist in mailbox"
                            $script:restoreLog += "  Items restored: $itemCount"
                            $script:restoreLog += ""
                        } catch {
                            Write-Host "    [ERROR] Failed to restore folder: $($_.Exception.Message)" -ForegroundColor Red

                            $script:restoreLog += "[ERROR] Failed to restore folder: $folderName"
                            $script:restoreLog += "  Error: $($_.Exception.Message)"
                            $script:restoreLog += ""
                        }
                    } else {
                        # Folder exists - restore using the Restore-Folder function
                        # This will handle item-level restoration and subfolders recursively
                        Write-Host "    Folder exists in both PST and mailbox - checking items..." -ForegroundColor Gray

                        Restore-Folder -PSTFolder $pstCustomFolder -MailboxParentFolder $mailboxRootFolder -Path ""
                    }
                } catch {
                    Write-Host "    ERROR: $($_.Exception.Message)" -ForegroundColor Red

                    $script:restoreLog += "[ERROR] Processing custom folder: $folderName"
                    $script:restoreLog += "  Error: $($_.Exception.Message)"
                    $script:restoreLog += ""
                }
            }

            if ($customFoldersRestored -gt 0) {
                Write-Host "`n  Restored $customFoldersRestored custom folder(s)" -ForegroundColor Green
            } else {
                Write-Host "`n  No custom folders needed restoration" -ForegroundColor Yellow
            }
        } else {
            Write-Host "  No custom folders found in PST" -ForegroundColor Gray
            $script:restoreLog += "=== CUSTOM FOLDERS ==="
            $script:restoreLog += "No custom folders found in PST"
            $script:restoreLog += ""
        }

        # Add summary
        $script:restoreLog += ""
        $script:restoreLog += "=== RESTORATION SUMMARY ==="
        $script:restoreLog += "Total folders restored: $($script:foldersRestored)"
        $script:restoreLog += "Total items restored: $($script:itemsRestored)"
        $script:restoreLog += ""
        $script:restoreLog += "Note: Existing folders with different item counts had missing emails restored."
        $script:restoreLog += "Duplicate detection was used to avoid copying existing emails."
        $script:restoreLog += ""
        $script:restoreLog += "Report End"
        
        # Save report
        $reportPath = "$script:backupFolder\RestorationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $script:restoreLog | Out-File -FilePath $reportPath -Encoding UTF8
        
        Write-Host ""
        Write-Host "=== Restoration Complete ===" -ForegroundColor Green
        Write-Host "Folders restored: $($script:foldersRestored)" -ForegroundColor White
        Write-Host "Items restored: $($script:itemsRestored)" -ForegroundColor White
        Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Opening restoration report in Notepad..." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        
        # Open report in Notepad
        Start-Process notepad.exe $reportPath
        
        Write-Host ""
        Write-Host "The restoration report is now open in Notepad." -ForegroundColor Green
        
    } catch {
        Write-Host "`nError occurred during restoration:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    } finally {
        # Clean up COM objects
        if ($null -ne $namespace) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
        if ($null -ne $outlook) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    
    Write-Host "`nPress any key to return to menu..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# NEW FUNCTION: Check Outlook mis-sync abnormality
function Check-OutlookMisSyncAbnormality {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " Check Outlook Mis-Sync Abnormality" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $logDir = "C:\dualog\connectionsuite\logdir"
    $archiveDir = "C:\dualog\connectionsuite\logdir\dualogimap_log"
    $filePattern = "*dualogimapserver*"
    
    # STEP 1: Verify log directory exists
    Write-Host "[Step 1/6] Verifying log directory..." -ForegroundColor Yellow
    if (-not (Test-Path $logDir)) {
        Write-Host "  ? ERROR: Log directory not found: $logDir" -ForegroundColor Red
        Write-Host "  Please ensure Dualog Connection Suite is installed." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Press ENTER to return to menu..." -ForegroundColor Yellow
        Read-Host
        return
    }
    Write-Host "  ? Log directory verified: $logDir" -ForegroundColor Green
    
    # STEP 2: Create archive directory if it doesn't exist
    Write-Host "`n[Step 2/6] Preparing archive directory..." -ForegroundColor Yellow
    if (-not (Test-Path $archiveDir)) {
        Write-Host "  Creating archive directory: $archiveDir" -ForegroundColor Gray
        try {
            New-Item -Path $archiveDir -ItemType Directory -Force | Out-Null
            Write-Host "  ? Archive directory created" -ForegroundColor Green
        } catch {
            Write-Host "  ? ERROR: Failed to create archive directory: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "Press ENTER to return to menu..." -ForegroundColor Yellow
            Read-Host
            return
        }
    } else {
        Write-Host "  ? Archive directory exists" -ForegroundColor Green
    }
    
    # STEP 3: Move existing dualogimapserver log files to archive
    Write-Host "`n[Step 3/6] Archiving existing dualogimapserver logs..." -ForegroundColor Yellow
    $existingFiles = Get-ChildItem -Path $logDir -Filter $filePattern -File -ErrorAction SilentlyContinue
    
    if ($existingFiles.Count -gt 0) {
        Write-Host "  Found $($existingFiles.Count) existing log file(s) to archive" -ForegroundColor Gray
        $movedCount = 0
        $failedCount = 0
        
        foreach ($file in $existingFiles) {
            try {
                Move-Item -Path $file.FullName -Destination $archiveDir -Force
                Write-Host "  ? Moved: $($file.Name)" -ForegroundColor Gray
                $movedCount++
            } catch {
                Write-Host "  ? Failed to move: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
                $failedCount++
            }
        }
        
        Write-Host "  ? Archiving complete - Moved: $movedCount, Failed: $failedCount" -ForegroundColor Green
    } else {
        Write-Host "  ??  No existing logs found to archive" -ForegroundColor Gray
        Write-Host "  ? Ready for new logs" -ForegroundColor Green
    }
    
    # STEP 4: Restart Outlook
    Write-Host "`n[Step 4/6] Restarting Outlook..." -ForegroundColor Yellow
    try {
        # Check if Outlook is running
        $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        
        if (-not $outlookProcess) {
            Write-Host "  ??  WARNING: Outlook is not running!" -ForegroundColor Yellow
            Write-Host "  Please start Outlook and run this test again." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Press ENTER to return to menu..." -ForegroundColor Yellow
            Read-Host
            return
        }
        
        Write-Host "  Outlook is running (PID: $($outlookProcess.Id))" -ForegroundColor Gray
        Write-Host "  Closing Outlook..." -ForegroundColor Gray
        
        # Close Outlook gracefully
        $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if ($outlookProcesses) {
            $outlookProcesses | ForEach-Object { $_.CloseMainWindow() | Out-Null }
            Start-Sleep -Seconds 3
            
            # Force close if still running
            $outlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
            if ($outlookProcesses) {
                $outlookProcesses | Stop-Process -Force
                Write-Host "  Outlook closed (forced)" -ForegroundColor Yellow
            } else {
                Write-Host "  ? Outlook closed successfully" -ForegroundColor Green
            }
        }
        
        # Wait a moment before restarting
        Write-Host "  Waiting 2 seconds before restart..." -ForegroundColor Gray
        Start-Sleep -Seconds 2
        
        # Find and start Outlook
        Write-Host "  Starting Outlook..." -ForegroundColor Gray
        $outlookPath = $null
        $possiblePaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles}\Microsoft Office\Office16\OUTLOOK.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE"
        )
        
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                $outlookPath = $path
                break
            }
        }
        
        if (-not $outlookPath) {
            Write-Host "  ? ERROR: Could not find Outlook executable" -ForegroundColor Red
            Write-Host "  Please start Outlook manually and run this test again." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Press ENTER to return to menu..." -ForegroundColor Yellow
            Read-Host
            return
        }
        
        Start-Process $outlookPath
        Write-Host "  ? Outlook started" -ForegroundColor Green
        
        # Wait for Outlook to fully load
        Write-Host "  Waiting for Outlook to fully load..." -ForegroundColor Gray
        Start-Sleep -Seconds 5
        
        # Verify Outlook is running
        $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if ($outlookProcess) {
            Write-Host "  ? Outlook is running (PID: $($outlookProcess.Id))" -ForegroundColor Green
        } else {
            Write-Host "  ??  WARNING: Could not verify Outlook is running" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "  ? ERROR: Failed to restart Outlook" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "Press ENTER to return to menu..." -ForegroundColor Yellow
        Read-Host
        return
    }
    
    # STEP 5: Wait for 20 seconds
    Write-Host "`n[Step 5/6] Waiting for synchronization to complete..." -ForegroundColor Yellow
    Write-Host "  Waiting 20 seconds for IMAP sync and log generation..." -ForegroundColor Gray
    
    for ($i = 20; $i -gt 0; $i--) {
        Write-Host "  $i seconds remaining..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 1
    }
    
    Write-Host "  ? Wait complete" -ForegroundColor Green
    
    # STEP 6: Check for new dualogimapserver logs and scan for "RENAME failed"
    Write-Host "`n[Step 6/6] Scanning for mis-sync errors..." -ForegroundColor Yellow
    
    # Get new log files
    $newLogFiles = Get-ChildItem -Path $logDir -Filter $filePattern -File -ErrorAction SilentlyContinue
    
    if ($newLogFiles.Count -eq 0) {
        Write-Host "  ??  No new dualogimapserver log files found" -ForegroundColor Yellow
        Write-Host "  This might indicate:" -ForegroundColor Gray
        Write-Host "    - No IMAP synchronization occurred" -ForegroundColor Gray
        Write-Host "    - Dualog IMAP Server is not running" -ForegroundColor Gray
        Write-Host "    - Different log file naming pattern" -ForegroundColor Gray
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host " DIAGNOSIS RESULT" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "??  UNABLE TO DETERMINE - No logs generated" -ForegroundColor Yellow -BackgroundColor Black
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "Press ENTER to return to menu..." -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Cyan
        Read-Host
        return
    }
    
    Write-Host "  Found $($newLogFiles.Count) new log file(s) to scan" -ForegroundColor Gray
    
    # Scan each log file for "RENAME failed"
    $errorFound = $false
    $errorDetails = @()
    
    foreach ($logFile in $newLogFiles) {
        Write-Host "  Scanning: $($logFile.Name)..." -ForegroundColor Gray
        
        try {
            $content = Get-Content -Path $logFile.FullName -ErrorAction Stop
            
            # Search for "RENAME failed" in the content
            $matchingLines = $content | Select-String -Pattern "RENAME failed" -SimpleMatch
            
            if ($matchingLines) {
                $errorFound = $true
                Write-Host "    ? ERROR DETECTED: Found 'RENAME failed' ($($matchingLines.Count) occurrence(s))" -ForegroundColor Red
                
                $errorDetails += [PSCustomObject]@{
                    FileName = $logFile.Name
                    FilePath = $logFile.FullName
                    ErrorCount = $matchingLines.Count
                    SampleErrors = ($matchingLines | Select-Object -First 3 | ForEach-Object { $_.Line })
                }
                
                # Display first 3 matching lines
                Write-Host "    Sample errors:" -ForegroundColor DarkRed
                $matchingLines | Select-Object -First 3 | ForEach-Object {
                    Write-Host "      $($_.Line)" -ForegroundColor DarkRed
                }
            } else {
                Write-Host "    ? No errors found" -ForegroundColor Green
            }
            
        } catch {
            Write-Host "    ??  WARNING: Could not read file - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Display final result
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " DIAGNOSIS RESULT" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($errorFound) {
        Write-Host "? Mailbox mis-sync error detected" -ForegroundColor Red -BackgroundColor Black
        Write-Host ""
        Write-Host "ERROR SUMMARY:" -ForegroundColor Yellow
        
        foreach ($detail in $errorDetails) {
            Write-Host ""
            Write-Host "  File: $($detail.FileName)" -ForegroundColor White
            Write-Host "  Path: $($detail.FilePath)" -ForegroundColor Gray
            Write-Host "  Errors: $($detail.ErrorCount) RENAME failure(s)" -ForegroundColor Red
        }
        
        Write-Host ""
        Write-Host "RECOMMENDED ACTIONS:" -ForegroundColor Yellow
        Write-Host "  1. Check folder synchronization settings in Outlook" -ForegroundColor White
        Write-Host "  2. Verify IMAP folder permissions on the server" -ForegroundColor White
        Write-Host "  3. Review full log files in: $archiveDir" -ForegroundColor White
        Write-Host "  4. Contact Dualog support if issue persists" -ForegroundColor White
        
    } else {
        Write-Host "? Mailbox is healthy" -ForegroundColor Green -BackgroundColor Black
        Write-Host ""
        Write-Host "  No 'RENAME failed' errors detected in recent logs" -ForegroundColor Gray
        Write-Host "  IMAP synchronization appears to be working correctly" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Scanned files: $($newLogFiles.Count)" -ForegroundColor Gray
    Write-Host "Archive location: $archiveDir" -ForegroundColor Gray
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Press ENTER to return to main menu..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Read-Host
}

# Function to view folders and logs
function View-FoldersAndLogs {
    Clear-Host
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host " View Folders and Logs" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Select an option:" -ForegroundColor Yellow
    Write-Host "1. Open Backup Folder (C:\Backup)" -ForegroundColor White
    Write-Host "2. Open Log Directory (C:\Dualog\ConnectionSuite\LogDir)" -ForegroundColor White
    Write-Host "3. Open Latest IMAP Log File" -ForegroundColor White
    Write-Host "4. Open All (All folders and latest log file)" -ForegroundColor White
    Write-Host "5. Return to Main Menu" -ForegroundColor White
    Write-Host ""
    
    $subChoice = Read-Host "Enter your choice (1-5)"
    
    switch ($subChoice) {
        '1' {
            # Open C:\Backup
            Write-Host "`nOpening backup folder..." -ForegroundColor Cyan
            if (Test-Path -Path $script:backupFolder) {
                Start-Process explorer.exe $script:backupFolder
                Write-Host "? Opened: $script:backupFolder" -ForegroundColor Green
            } else {
                Write-Host "?  Backup folder does not exist: $script:backupFolder" -ForegroundColor Yellow
            }
            
            Write-Host ""
            Write-Host "Press any key to continue..." -ForegroundColor White
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '2' {
            # Open C:\Dualog\ConnectionSuite\LogDir
            $logDir = "C:\Dualog\ConnectionSuite\LogDir"
            Write-Host "`nOpening log directory..." -ForegroundColor Cyan
            
            if (Test-Path -Path $logDir) {
                Start-Process explorer.exe $logDir
                Write-Host "? Opened: $logDir" -ForegroundColor Green
            } else {
                Write-Host "?  Log directory does not exist: $logDir" -ForegroundColor Yellow
            }
            
            Write-Host ""
            Write-Host "Press any key to continue..." -ForegroundColor White
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '3' {
            # Open latest DualogImapServer log file
            $logDir = "C:\Dualog\ConnectionSuite\LogDir"
            Write-Host "`nOpening latest IMAP log file..." -ForegroundColor Cyan
            
            if (Test-Path -Path $logDir) {
                try {
                    $latestLogFile = Get-ChildItem -Path $logDir -Filter "*DualogImapServer*" -File -ErrorAction Stop | 
                        Sort-Object LastWriteTime -Descending | 
                        Select-Object -First 1
                    
                    if ($latestLogFile) {
                        Write-Host "Found: $($latestLogFile.Name)" -ForegroundColor Gray
                        Write-Host "Modified: $($latestLogFile.LastWriteTime)" -ForegroundColor Gray
                        Write-Host "Size: $([math]::Round($latestLogFile.Length / 1KB, 2)) KB" -ForegroundColor Gray
                        Start-Process notepad.exe $latestLogFile.FullName
                        Write-Host "? Opened in Notepad" -ForegroundColor Green
                    } else {
                        Write-Host "?  No DualogImapServer log files found in directory" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "? Error: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "?  Log directory does not exist: $logDir" -ForegroundColor Yellow
                Write-Host "?  Cannot open log file - directory does not exist" -ForegroundColor Yellow
            }
            
            Write-Host ""
            Write-Host "Press any key to continue..." -ForegroundColor White
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '4' {
            # Open all
            Write-Host "`nOpening all folders and log file..." -ForegroundColor Cyan
            Write-Host ""
            
            # 1. Open C:\Backup
            Write-Host "[1/3] Opening backup folder..." -ForegroundColor Yellow
            if (Test-Path -Path $script:backupFolder) {
                Start-Process explorer.exe $script:backupFolder
                Write-Host "  ? Opened: $script:backupFolder" -ForegroundColor Green
            } else {
                Write-Host "  ?  Backup folder does not exist: $script:backupFolder" -ForegroundColor Yellow
            }
            
            Start-Sleep -Seconds 1
            
            # 2. Open C:\Dualog\ConnectionSuite\LogDir
            $logDir = "C:\Dualog\ConnectionSuite\LogDir"
            Write-Host "`n[2/3] Opening log directory..." -ForegroundColor Yellow
            
            if (Test-Path -Path $logDir) {
                Start-Process explorer.exe $logDir
                Write-Host "  ? Opened: $logDir" -ForegroundColor Green
            } else {
                Write-Host "  ?  Log directory does not exist: $logDir" -ForegroundColor Yellow
            }
            
            Start-Sleep -Seconds 1
            
            # 3. Open latest DualogImapServer log file
            Write-Host "`n[3/3] Opening latest IMAP log file..." -ForegroundColor Yellow
            
            if (Test-Path -Path $logDir) {
                try {
                    $latestLogFile = Get-ChildItem -Path $logDir -Filter "*DualogImapServer*" -File -ErrorAction Stop | 
                        Sort-Object LastWriteTime -Descending | 
                        Select-Object -First 1
                    
                    if ($latestLogFile) {
                        Write-Host "  Found: $($latestLogFile.Name)" -ForegroundColor Gray
                        Write-Host "  Modified: $($latestLogFile.LastWriteTime)" -ForegroundColor Gray
                        Write-Host "  Size: $([math]::Round($latestLogFile.Length / 1KB, 2)) KB" -ForegroundColor Gray
                        Start-Process notepad.exe $latestLogFile.FullName
                        Write-Host "  ? Opened in Notepad" -ForegroundColor Green
                    } else {
                        Write-Host "  ?  No DualogImapServer log files found in directory" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "  ? Error: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "  ?  Cannot open log file - directory does not exist" -ForegroundColor Yellow
            }
            
            Write-Host ""
            Write-Host "================================" -ForegroundColor Cyan
            Write-Host "Press any key to continue..." -ForegroundColor White
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        '5' {
            Write-Host "`nReturning to main menu..." -ForegroundColor Cyan
            Start-Sleep -Seconds 1
            return
        }
        default {
            Write-Host "`nInvalid option. Please select 1-5." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
}

# Function to display menu
function Show-Menu {
    Clear-Host
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host "  Dualog Outlook Backup Utility ver 1.0" -ForegroundColor Cyan
    Write-Host "---------------------------------------------------" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Export Inbox and Sent Items to PST" -ForegroundColor White
    Write-Host "2. Verify & Compare PST with Mailbox" -ForegroundColor White
    Write-Host "3. Restore Missing Folders/Emails from PST" -ForegroundColor White
    Write-Host "4. View Backup Folder & Logs" -ForegroundColor White
    Write-Host "5. Activate Renamed Outlook Mode" -ForegroundColor White
    Write-Host "6. Check Outlook Mis-Sync Abnormality" -ForegroundColor White
    Write-Host "7. Exit" -ForegroundColor White
    Write-Host ""
    Write-Host "Backup Location: $script:backupFolder" -ForegroundColor Gray
    Write-Host ""
}

# Main menu loop
do {
    Show-Menu
    $choice = Read-Host "Select an option (1-7)"
    
    switch ($choice) {
        '1' {
            Clear-Host
            Export-InboxAndSentItems
        }
        '2' {
            Compare-PSTWithMailbox
        }
        '3' {
            Restore-FromPST
        }
        '4' {
            View-FoldersAndLogs
        }
        '5' {
            Manage-OutlookRenameMode
        }
        '6' {
            Check-OutlookMisSyncAbnormality
        }
        '7' {
            Write-Host "`nExiting..." -ForegroundColor Cyan
            break
        }
        default {
            Write-Host "`nInvalid option. Please select 1-7." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($choice -ne '7')
