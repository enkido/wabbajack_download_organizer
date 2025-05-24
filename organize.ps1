Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = $MyInvocation.MyCommand.Path
$defaultFolder = Split-Path -Parent $scriptPath
$organizerLogPath = Join-Path $defaultFolder "organize_activity.log"

$form = New-Object System.Windows.Forms.Form
$form.Text = "File Organizer"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(400, 300)

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source Folder:"
$lblSource.Location = New-Object System.Drawing.Point(10,20)
$lblSource.Size = New-Object System.Drawing.Size(100,20)
$lblSource.Anchor = "Top, Left"
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(120,20)
$txtSource.Size = New-Object System.Drawing.Size(350,20)
$txtSource.Anchor = "Top, Left, Right"
$txtSource.Text = $defaultFolder
$form.Controls.Add($txtSource)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Text = "..."
$btnBrowseSource.Location = New-Object System.Drawing.Point(480,20)
$btnBrowseSource.Size = New-Object System.Drawing.Size(30,20)
$btnBrowseSource.Anchor = "Top, Right"
$btnBrowseSource.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.SelectedPath = $txtSource.Text
    if ($fbd.ShowDialog() -eq "OK") {
        $txtSource.Text = $fbd.SelectedPath
    }
})
$form.Controls.Add($btnBrowseSource)

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry Run (nur anzeigen)"
$chkDryRun.Location = New-Object System.Drawing.Point(10,50)
$chkDryRun.Anchor = "Top, Left"
$form.Controls.Add($chkDryRun)

$lstLog = New-Object System.Windows.Forms.ListBox
$lstLog.Location = New-Object System.Drawing.Point(10, 80)
$lstLog.Size = New-Object System.Drawing.Size(560, 330)
$lstLog.Anchor = "Top, Bottom, Left, Right"
$form.Controls.Add($lstLog)

$logBuffer = New-Object System.Collections.ArrayList
$logTimer = New-Object System.Windows.Forms.Timer
$logTimer.Interval = 200
$logTimer.Add_Tick({
    if ($logBuffer.Count -gt 0) {
        foreach ($msg in $logBuffer) {
            $lstLog.Items.Add($msg) | Out-Null
        }
        $lstLog.TopIndex = $lstLog.Items.Count - 1
        $lstLog.Refresh()
        $logBuffer.Clear()
    }
})
$logTimer.Start()

function Add-Log {
    param([string]$message)

    # Add to UI log buffer (existing functionality)
    [void]$logBuffer.Add($message)

    # Add to file log
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp - $message"
		Write-Host $logEntry
        Add-Content -Path $organizerLogPath -Value $logEntry -Encoding UTF8 -ErrorAction Stop
    } catch {
        # If file logging fails, output error to console/error stream.
        # This won't go into the UI log via Add-Log to prevent recursion.
        $uiErrorMessage = "[ERROR] Failed to write to log file: $($_.Exception.Message)"
        [void]$logBuffer.Add($uiErrorMessage) # Try to add error to UI buffer as a fallback
        Write-Error "FILE_LOG_ERROR: Failed to write to log file $organizerLogPath. Error: $($_.Exception.Message)"
    }
}

function Parse-MetaFile {
    param([string]$metaPath)
    $result = @{}
    try {
        $lines = Get-Content $metaPath -ErrorAction Stop
        foreach ($line in $lines) {
            if ($line -match '^([a-zA-Z0-9_]+)=(.+)$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $result[$key] = $value
            }
        }
    } catch {
        Add-Log "[ERROR] Fehler beim Lesen der Meta-Datei: $metaPath - $_"
    }
    return $result
}

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Start"
$btnRun.Location = New-Object System.Drawing.Point(10, 420)
$btnRun.Size = New-Object System.Drawing.Size(100, 30)
$btnRun.Anchor = "Bottom, Left"
$form.Controls.Add($btnRun)

$btnRun.Add_Click({
    $source = $txtSource.Text.Trim()
    $dryRun = $chkDryRun.Checked

    if ([string]::IsNullOrWhiteSpace($source)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte Quellordner angeben.", "Fehlende Eingabe", 'OK', 'Error')
        return
    }

    Add-Log "BTN_CLICK: Attempting to start job. Source: [$source], DryRun: [$dryRun]"
    $job = $null # Initialize $job to null
    try {
        $job = Start-Job -ScriptBlock {
            param($source, $dryRun)

            $installed = Join-Path $source "installed"
            $unused = Join-Path $source "unused"
			$other = Join-Path $source "other"

			Write-Output "Start"
			if (-not $dryRun) {
				foreach ($dir in @($installed, $unused, $other)) {                
					if (!(Test-Path $dir)) {
						New-Item -ItemType Directory -Force -Path $dir | Out-Null
					}
				}
			}

			Get-ChildItem -Path $source -File | Where-Object { $_.Extension -ne ".meta" } | ForEach-Object {
				$file = $_
				$metaPath = $file.FullName + ".meta"

				if (Test-Path $metaPath) {
					$meta = @{}
					$lines = Get-Content $metaPath -ErrorAction SilentlyContinue
					foreach ($line in $lines) {
						if ($line -match '^([a-zA-Z0-9_]+)=(.+)$') {
							$meta[$matches[1]] = $matches[2]
						}
					}

					if ($meta['installed'] -eq 'true') {
						$dest = $installed
						$cat = "[INSTALLED]"
					} elseif ($meta['removed'] -eq 'true') {
						$dest = $unused
						$cat = "[UNUSED]"
					} else {
						$dest = $other
						$cat = "[OTHER]"
					}

					Write-Output "$cat $($file.Name) -> $dest"
					if (-not $dryRun) {
						Move-Item $file.FullName -Destination $dest -Force
						Move-Item $metaPath -Destination $dest -Force                    
					}
				} else {
					Write-Output "[NOMETA] $($file.Name) -> $other"
					if (-not $dryRun) {
						Move-Item $file.FullName -Destination $other -Force
					}
				}
			}
			Write-Output "JOB_EVENT: Processing complete."
		} -ArgumentList $source, $dryRun -ErrorAction Stop # Added -ErrorAction Stop here

        if ($null -ne $job) {
            Add-Log "BTN_CLICK: Job object created. ID: $($job.Id), Initial State: $($job.State)"
        } else {
            Add-Log "BTN_CLICK: Job object was NULL after Start-Job call."
        }
    } catch {
        Add-Log "BTN_CLICK: ERROR during Start-Job call: $($_.Exception.ToString())"
    }

    if ($null -ne $job) {
        Add-Log "BTN_CLICK: Starting to poll job ID $($job.Id)."
        $jobStillRunning = $true
        while ($jobStillRunning) {
            $jobState = $job.JobStateInfo.State
            Add-Log "BTN_CLICK_POLL: Job ID $($job.Id) state: $jobState"

            # Try to get output without blocking or changing job state
            # Receive-Job by default uses -Keep for running jobs
            $currentOutput = Receive-Job -Job $job
            if ($null -ne $currentOutput) {
                foreach ($line in $currentOutput) {
                    # Using Add-Log here relies on Add-Log writing to the file log.
                    Add-Log "JOB_POLLED_OUTPUT: $line"
                }
            }

            if ($jobState -in ([System.Management.Automation.JobState]::Completed, [System.Management.Automation.JobState]::Failed, [System.Management.Automation.JobState]::Stopped, [System.Management.Automation.JobState]::Suspended)) {
                $jobStillRunning = $false
                Add-Log "BTN_CLICK_POLL: Job ID $($job.Id) is in a terminal state: $jobState."
            } else {
                Start-Sleep -Milliseconds 100 # Wait before next poll
            }
        }

        # Final attempt to get any remaining output
        Add-Log "BTN_CLICK: Final Receive-Job for job ID $($job.Id)."
        $finalOutput = Receive-Job -Job $job # This will also get remaining output and clear it for a completed job
        if ($null -ne $finalOutput) {
            foreach ($line in $finalOutput) {
                Add-Log "JOB_FINAL_OUTPUT: $line"
            }
        }
        
        Add-Log "BTN_CLICK: Removing job ID $($job.Id)."
        Remove-Job -Job $job
        Add-Log "BTN_CLICK: Polling finished for job ID $($job.Id)."

    } else {
        Add-Log "BTN_CLICK: Skipping polling loop because job object was null."
    }
})


Write-Host $organizerLogPath
Add-Log $organizerLogPath
[void]$form.ShowDialog()