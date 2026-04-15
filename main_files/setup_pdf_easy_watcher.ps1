Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptRoot   = Split-Path -Parent $MyInvocation.MyCommand.Path
$settingsPath = Join-Path $scriptRoot "easy_watcher_settings.json"

function Get-DocumentsFolder {
    return [Environment]::GetFolderPath("MyDocuments")
}

function Save-Settings {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WatchFolder,

        [Parameter(Mandatory = $true)]
        [string]$ExportFolder
    )

    $settings = [PSCustomObject]@{
        watchFolder  = $WatchFolder
        exportFolder = $ExportFolder
    }

    $settings | ConvertTo-Json -Depth 3 | Set-Content -Path $settingsPath -Encoding UTF8
}

function Pick-Folder {
    param(
        [string]$InitialPath
    )

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($InitialPath -and (Test-Path $InitialPath)) {
        $dialog.SelectedPath = $InitialPath
    }

    $result = $dialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }

    return $null
}

$documents = Get-DocumentsFolder

$bg        = [System.Drawing.Color]::FromArgb(245, 247, 250)
$card      = [System.Drawing.Color]::White
$textMain  = [System.Drawing.Color]::FromArgb(30, 41, 59)
$textMuted = [System.Drawing.Color]::FromArgb(71, 85, 105)
$accent    = [System.Drawing.Color]::FromArgb(25, 118, 210)

$form = New-Object System.Windows.Forms.Form
$form.Text = "Easy PDF Watcher Setup"
$form.Size = New-Object System.Drawing.Size(760, 340)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.TopMost = $true
$form.BackColor = $bg
$form.Opacity = 0.98

$title = New-Object System.Windows.Forms.Label
$title.Location = New-Object System.Drawing.Point(24, 18)
$title.Size = New-Object System.Drawing.Size(500, 30)
$title.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = $textMain
$title.Text = "Easy PDF Watcher Setup"
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Location = New-Object System.Drawing.Point(26, 50)
$subtitle.Size = New-Object System.Drawing.Size(650, 20)
$subtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$subtitle.ForeColor = $textMuted
$subtitle.Text = "Choose your watch folder and export folder. Cancel will keep both set to Documents."
$form.Controls.Add($subtitle)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(18, 86)
$panel.Size = New-Object System.Drawing.Size(708, 165)
$panel.BackColor = $card
$panel.BorderStyle = "FixedSingle"
$form.Controls.Add($panel)

$watchLabel = New-Object System.Windows.Forms.Label
$watchLabel.Location = New-Object System.Drawing.Point(18, 18)
$watchLabel.Size = New-Object System.Drawing.Size(280, 20)
$watchLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
$watchLabel.ForeColor = $textMuted
$watchLabel.Text = "Choose what folder to watch (your PDF)"
$panel.Controls.Add($watchLabel)

$watchBox = New-Object System.Windows.Forms.TextBox
$watchBox.Location = New-Object System.Drawing.Point(18, 42)
$watchBox.Size = New-Object System.Drawing.Size(540, 26)
$watchBox.Text = $documents
$watchBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$panel.Controls.Add($watchBox)

$watchBrowse = New-Object System.Windows.Forms.Button
$watchBrowse.Location = New-Object System.Drawing.Point(575, 40)
$watchBrowse.Size = New-Object System.Drawing.Size(95, 30)
$watchBrowse.Text = "Browse"
$watchBrowse.BackColor = $accent
$watchBrowse.ForeColor = [System.Drawing.Color]::White
$watchBrowse.FlatStyle = "Flat"
$watchBrowse.FlatAppearance.BorderSize = 0
$watchBrowse.Add_Click({
    $picked = Pick-Folder -InitialPath $watchBox.Text
    if ($picked) {
        $watchBox.Text = $picked
    }
})
$panel.Controls.Add($watchBrowse)

$exportLabel = New-Object System.Windows.Forms.Label
$exportLabel.Location = New-Object System.Drawing.Point(18, 88)
$exportLabel.Size = New-Object System.Drawing.Size(250, 20)
$exportLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
$exportLabel.ForeColor = $textMuted
$exportLabel.Text = "Choose where to export"
$panel.Controls.Add($exportLabel)

$exportBox = New-Object System.Windows.Forms.TextBox
$exportBox.Location = New-Object System.Drawing.Point(18, 112)
$exportBox.Size = New-Object System.Drawing.Size(540, 26)
$exportBox.Text = $documents
$exportBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$panel.Controls.Add($exportBox)

$exportBrowse = New-Object System.Windows.Forms.Button
$exportBrowse.Location = New-Object System.Drawing.Point(575, 110)
$exportBrowse.Size = New-Object System.Drawing.Size(95, 30)
$exportBrowse.Text = "Browse"
$exportBrowse.BackColor = $accent
$exportBrowse.ForeColor = [System.Drawing.Color]::White
$exportBrowse.FlatStyle = "Flat"
$exportBrowse.FlatAppearance.BorderSize = 0
$exportBrowse.Add_Click({
    $picked = Pick-Folder -InitialPath $exportBox.Text
    if ($picked) {
        $exportBox.Text = $picked
    }
})
$panel.Controls.Add($exportBrowse)

$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Location = New-Object System.Drawing.Point(250, 265)
$applyButton.Size = New-Object System.Drawing.Size(110, 36)
$applyButton.Text = "Apply"
$applyButton.BackColor = $accent
$applyButton.ForeColor = [System.Drawing.Color]::White
$applyButton.FlatStyle = "Flat"
$applyButton.FlatAppearance.BorderSize = 0
$applyButton.Add_Click({
    $watchFolder = $watchBox.Text.Trim()
    $exportFolder = $exportBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($watchFolder)) {
        $watchFolder = $documents
    }
    if ([string]::IsNullOrWhiteSpace($exportFolder)) {
        $exportFolder = $documents
    }

    Save-Settings -WatchFolder $watchFolder -ExportFolder $exportFolder

    [System.Windows.Forms.MessageBox]::Show(
        "Settings saved successfully.",
        "Easy PDF Watcher Setup",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $form.Close()
})
$form.Controls.Add($applyButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(380, 265)
$cancelButton.Size = New-Object System.Drawing.Size(110, 36)
$cancelButton.Text = "Cancel"
$cancelButton.BackColor = [System.Drawing.Color]::FromArgb(100, 116, 139)
$cancelButton.ForeColor = [System.Drawing.Color]::White
$cancelButton.FlatStyle = "Flat"
$cancelButton.FlatAppearance.BorderSize = 0
$cancelButton.Add_Click({
    Save-Settings -WatchFolder $documents -ExportFolder $documents

    [System.Windows.Forms.MessageBox]::Show(
        "Default settings saved to Documents.",
        "Easy PDF Watcher Setup",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    $form.Close()
})
$form.Controls.Add($cancelButton)

[void]$form.ShowDialog()