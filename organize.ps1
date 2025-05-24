Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptPath = $MyInvocation.MyCommand.Path
$defaultFolder = Split-Path -Parent $scriptPath

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
$logTimer.Interval = 1000
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
    [void]$logBuffer.Add($message)
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

    $job = Start-Job -ScriptBlock {
        param($source, $dryRun)

        $installed = Join-Path $source "installed"
        $unused = Join-Path $source "unused"
        $other = Join-Path $source "other"

        if (-not $dryRun) {
            foreach ($dir in @($installed, $unused, $other)) {
                if (!(Test-Path $dir)) {
                    try {
                        New-Item -ItemType Directory -Force -Path $dir -ErrorAction Stop | Out-Null
                    } catch {
                        Write-Output "[JOB_ERROR] Failed to create directory ${dir}: $($_.Exception.Message)"
                    }
                }
            }
        }

        Get-ChildItem -Path $source -File | Where-Object { $_.Extension -ne ".meta" } | ForEach-Object {
            $file = $_
            $metaPath = $file.FullName + ".meta"

            if (Test-Path $metaPath) {
                $meta = @{}
                $lines = Get-Content $metaPath -ErrorAction Stop
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
                    try {
                        Move-Item $file.FullName -Destination $dest -Force -ErrorAction Stop
                    } catch {
                        Write-Output "[JOB_ERROR] Failed to move $($file.Name) to ${dest}: $($_.Exception.Message)"
                    }
                    try {
                        Move-Item $metaPath -Destination $dest -Force -ErrorAction Stop
                    } catch {
                        Write-Output "[JOB_ERROR] Failed to move $metaPath to ${dest}: $($_.Exception.Message)"
                    }
                }
            } else {
                Write-Output "[NOMETA] $($file.Name) -> $other"
                if (-not $dryRun) {
                    try {
                        Move-Item $file.FullName -Destination $other -Force -ErrorAction Stop
                    } catch {
                        Write-Output "[JOB_ERROR] Failed to move $($file.Name) to ${other}: $($_.Exception.Message)"
                    }
                }
            }
        }
    } -ArgumentList $source, $dryRun

    Register-ObjectEvent -InputObject $job -EventName StateChanged -Action {
        if ($Event.SourceEventArgs.JobStateInfo.State -eq 'Completed') {
            $output = Receive-Job -Job $Event.Sender
            foreach ($line in $output) {
                Add-Log $line
            }
            Remove-Job -Job $Event.Sender
        }
    } | Out-Null
})

[void]$form.ShowDialog()
