Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptRoot   = Split-Path -Parent $MyInvocation.MyCommand.Path
$settingsPath = Join-Path $scriptRoot "easy_watcher_settings.json"
$marcaPath    = Join-Path $scriptRoot "marca_sola.png"
$logoPath     = Join-Path $scriptRoot "physical_fac_logo.png"

$script:lastFinalPath  = $null
$script:lastFolderPath = $null
$script:processedFiles = @{}
$script:recentEvents   = @{}
$script:isDialogOpen   = $false

function Get-DocumentsFolder {
    return [Environment]::GetFolderPath("MyDocuments")
}

function New-DefaultSettingsObject {
    $documents = Get-DocumentsFolder
    return [PSCustomObject]@{
        watchFolder  = $documents
        exportFolder = $documents
    }
}

function Save-Settings {
    param(
        [Parameter(Mandatory = $true)]
        $Settings
    )

    $Settings | ConvertTo-Json -Depth 3 | Set-Content -Path $settingsPath -Encoding UTF8
}

function Load-Settings {
    if (-not (Test-Path $settingsPath)) {
        $defaultSettings = New-DefaultSettingsObject
        Save-Settings -Settings $defaultSettings
        return $defaultSettings
    }

    try {
        $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json

        if (-not $settings.watchFolder -or [string]::IsNullOrWhiteSpace($settings.watchFolder)) {
            $settings.watchFolder = Get-DocumentsFolder
        }

        if (-not $settings.exportFolder -or [string]::IsNullOrWhiteSpace($settings.exportFolder)) {
            $settings.exportFolder = Get-DocumentsFolder
        }

        return $settings
    }
    catch {
        $defaultSettings = New-DefaultSettingsObject
        Save-Settings -Settings $defaultSettings
        return $defaultSettings
    }
}

$settings = Load-Settings
$watchFolder  = $settings.watchFolder
$exportFolder = $settings.exportFolder

if (-not (Test-Path $watchFolder)) {
    $watchFolder = Get-DocumentsFolder
}
if (-not (Test-Path $exportFolder)) {
    $exportFolder = Get-DocumentsFolder
}

function Get-DatePart {
    return Get-Date -Format "yyMMdd"
}

function Get-TimePart {
    return Get-Date -Format "HHmm"
}

function Get-WOFromName {
    param(
        [string]$OriginalName
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalName)

    if ($baseName -match '(?i)\b([a-z]\d{4,})\b') {
        return $matches[1].ToUpper()
    }

    if ($baseName -match '(?i)([a-z]\d{4,})') {
        return $matches[1].ToUpper()
    }

    return "EDIT_ME"
}

function Build-FileName {
    param(
        [string]$DatePart,
        [string]$TimePart,
        [string]$WO
    )

    return "${DatePart}_${TimePart}_${WO}.pdf"
}

function Wait-ForFileReady {
    param(
        [string]$Path
    )

    for ($i = 0; $i -lt 30; $i++) {
        try {
            $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
            $stream.Close()
            return $true
        }
        catch {
            Start-Sleep -Milliseconds 500
        }
    }

    return $false
}

function Open-FileAndFolder {
    param(
        [string]$FolderPath,
        [string]$FinalPath
    )

    try {
        if ($FinalPath -and (Test-Path $FinalPath)) {
            Start-Process explorer.exe "/select,`"$FinalPath`""
        }
        elseif ($FolderPath -and (Test-Path $FolderPath)) {
            Start-Process explorer.exe $FolderPath
        }
    }
    catch {
    }
}

function Show-NotificationBalloon {
    param(
        [string]$Title,
        [string]$Body,
        [string]$FolderPath,
        [string]$FinalPath
    )

    try {
        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Information
        $notify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $notify.BalloonTipTitle = $Title
        $notify.BalloonTipText = $Body
        $notify.Visible = $true

        $script:lastFolderPath = $FolderPath
        $script:lastFinalPath  = $FinalPath

        Register-ObjectEvent -InputObject $notify -EventName BalloonTipClicked -Action {
            Open-FileAndFolder -FolderPath $script:lastFolderPath -FinalPath $script:lastFinalPath
        } | Out-Null

        $notify.ShowBalloonTip(7000)

        Start-Sleep -Seconds 8
        $notify.Dispose()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $Body,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}

function Get-AvailableTargetPath {
    param(
        [string]$DesiredPath
    )

    if (-not (Test-Path $DesiredPath)) {
        return $DesiredPath
    }

    $overwriteResult = [System.Windows.Forms.MessageBox]::Show(
        "A file with this name already exists.`n`nYes = Replace`nNo = Create v2`nCancel = Cancel",
        "Existing file",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($overwriteResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return $null
    }

    if ($overwriteResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Remove-Item $DesiredPath -Force
        return $DesiredPath
    }

    $folder = [System.IO.Path]::GetDirectoryName($DesiredPath)
    $base   = [System.IO.Path]::GetFileNameWithoutExtension($DesiredPath)
    $ext    = [System.IO.Path]::GetExtension($DesiredPath)

    $candidate = Join-Path $folder ($base + "_v2" + $ext)
    if (-not (Test-Path $candidate)) {
        return $candidate
    }

    $i = 3
    while ($true) {
        $candidate = Join-Path $folder ($base + "_v$i" + $ext)
        if (-not (Test-Path $candidate)) {
            return $candidate
        }
        $i++
    }
}

function New-ModernButton {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 120,
        [int]$Height = 38,
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::FromArgb(32, 99, 155),
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::White
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.Text = $Text
    $button.BackColor = $BackColor
    $button.ForeColor = $ForeColor
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 0
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $button
}

function Show-RenameDialog {
    param(
        [string]$OriginalPath
    )

    $originalName = [System.IO.Path]::GetFileName($OriginalPath)
    $detectedWO   = Get-WOFromName -OriginalName $originalName
    $datePart     = Get-DatePart
    $timePart     = Get-TimePart

    $bg         = [System.Drawing.Color]::FromArgb(245, 247, 250)
    $card       = [System.Drawing.Color]::White
    $textMain   = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $textMuted  = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $accent     = [System.Drawing.Color]::FromArgb(25, 118, 210)
    $accentSoft = [System.Drawing.Color]::FromArgb(232, 240, 254)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PDF detected"
    $form.Size = New-Object System.Drawing.Size(930, 500)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $bg
    $form.Opacity = 0.97

    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(930, 92)
    $headerPanel.BackColor = $card
    $form.Controls.Add($headerPanel)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(24, 16)
    $titleLabel.Size = New-Object System.Drawing.Size(420, 30)
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = $textMain
    $titleLabel.Text = "PDF file detected"
    $headerPanel.Controls.Add($titleLabel)

    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Location = New-Object System.Drawing.Point(26, 48)
    $subtitleLabel.Size = New-Object System.Drawing.Size(450, 20)
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subtitleLabel.ForeColor = $textMuted
    $subtitleLabel.Text = "Review the WO and confirm the final file name"
    $headerPanel.Controls.Add($subtitleLabel)

    if (Test-Path $marcaPath) {
        try {
            $marcaBox = New-Object System.Windows.Forms.PictureBox
            $marcaBox.Location = New-Object System.Drawing.Point(590, 14)
            $marcaBox.Size = New-Object System.Drawing.Size(120, 56)
            $marcaBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $marcaBox.Image = [System.Drawing.Image]::FromFile($marcaPath)
            $headerPanel.Controls.Add($marcaBox)
        }
        catch {
        }
    }

    if (Test-Path $logoPath) {
        try {
            $logoBox = New-Object System.Windows.Forms.PictureBox
            $logoBox.Location = New-Object System.Drawing.Point(720, 12)
            $logoBox.Size = New-Object System.Drawing.Size(175, 60)
            $logoBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
            $logoBox.Image = [System.Drawing.Image]::FromFile($logoPath)
            $headerPanel.Controls.Add($logoBox)
        }
        catch {
        }
    }

    $mainPanel = New-Object System.Windows.Forms.Panel
    $mainPanel.Location = New-Object System.Drawing.Point(18, 102)
    $mainPanel.Size = New-Object System.Drawing.Size(878, 300)
    $mainPanel.BackColor = $card
    $mainPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($mainPanel)

    $labelOriginal = New-Object System.Windows.Forms.Label
    $labelOriginal.Location = New-Object System.Drawing.Point(20, 18)
    $labelOriginal.Size = New-Object System.Drawing.Size(180, 20)
    $labelOriginal.Text = "Original file"
    $labelOriginal.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
    $labelOriginal.ForeColor = $textMuted
    $mainPanel.Controls.Add($labelOriginal)

    $textboxOriginal = New-Object System.Windows.Forms.TextBox
    $textboxOriginal.Location = New-Object System.Drawing.Point(20, 42)
    $textboxOriginal.Size = New-Object System.Drawing.Size(838, 28)
    $textboxOriginal.ReadOnly = $true
    $textboxOriginal.Text = $originalName
    $textboxOriginal.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $textboxOriginal.BackColor = [System.Drawing.Color]::White
    $textboxOriginal.ForeColor = $textMain
    $textboxOriginal.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainPanel.Controls.Add($textboxOriginal)

    $labelWatch = New-Object System.Windows.Forms.Label
    $labelWatch.Location = New-Object System.Drawing.Point(20, 84)
    $labelWatch.Size = New-Object System.Drawing.Size(140, 20)
    $labelWatch.Text = "Watch folder"
    $labelWatch.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
    $labelWatch.ForeColor = $textMuted
    $mainPanel.Controls.Add($labelWatch)

    $textboxWatch = New-Object System.Windows.Forms.TextBox
    $textboxWatch.Location = New-Object System.Drawing.Point(20, 108)
    $textboxWatch.Size = New-Object System.Drawing.Size(400, 28)
    $textboxWatch.ReadOnly = $true
    $textboxWatch.Text = $watchFolder
    $textboxWatch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $textboxWatch.BackColor = $accentSoft
    $textboxWatch.ForeColor = $textMain
    $textboxWatch.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainPanel.Controls.Add($textboxWatch)

    $labelExport = New-Object System.Windows.Forms.Label
    $labelExport.Location = New-Object System.Drawing.Point(458, 84)
    $labelExport.Size = New-Object System.Drawing.Size(140, 20)
    $labelExport.Text = "Export folder"
    $labelExport.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
    $labelExport.ForeColor = $textMuted
    $mainPanel.Controls.Add($labelExport)

    $textboxExport = New-Object System.Windows.Forms.TextBox
    $textboxExport.Location = New-Object System.Drawing.Point(458, 108)
    $textboxExport.Size = New-Object System.Drawing.Size(400, 28)
    $textboxExport.ReadOnly = $true
    $textboxExport.Text = $exportFolder
    $textboxExport.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $textboxExport.BackColor = $accentSoft
    $textboxExport.ForeColor = $textMain
    $textboxExport.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainPanel.Controls.Add($textboxExport)

    $labelWO = New-Object System.Windows.Forms.Label
    $labelWO.Location = New-Object System.Drawing.Point(20, 152)
    $labelWO.Size = New-Object System.Drawing.Size(140, 20)
    $labelWO.Text = "Detected WO"
    $labelWO.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
    $labelWO.ForeColor = $textMuted
    $mainPanel.Controls.Add($labelWO)

    $textboxWO = New-Object System.Windows.Forms.TextBox
    $textboxWO.Location = New-Object System.Drawing.Point(20, 176)
    $textboxWO.Size = New-Object System.Drawing.Size(190, 28)
    $textboxWO.Text = $detectedWO
    $textboxWO.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $textboxWO.BackColor = [System.Drawing.Color]::White
    $textboxWO.ForeColor = $textMain
    $textboxWO.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainPanel.Controls.Add($textboxWO)

    $labelDate = New-Object System.Windows.Forms.Label
    $labelDate.Location = New-Object System.Drawing.Point(250, 152)
    $labelDate.Size = New-Object System.Drawing.Size(100, 20)
    $labelDate.Text = "Date"
    $labelDate.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
    $labelDate.ForeColor = $textMuted
    $mainPanel.Controls.Add($labelDate)

    $textboxDate = New-Object System.Windows.Forms.TextBox
    $textboxDate.Location = New-Object System.Drawing.Point(250, 176)
    $textboxDate.Size = New-Object System.Drawing.Size(130, 28)
    $textboxDate.ReadOnly = $true
    $textboxDate.Text = $datePart
    $textboxDate.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $textboxDate.BackColor = $accentSoft
    $textboxDate.ForeColor = $textMain
    $textboxDate.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainPanel.Controls.Add($textboxDate)

    $labelTime = New-Object System.Windows.Forms.Label
    $labelTime.Location = New-Object System.Drawing.Point(420, 152)
    $labelTime.Size = New-Object System.Drawing.Size(100, 20)
    $labelTime.Text = "24h time"
    $labelTime.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
    $labelTime.ForeColor = $textMuted
    $mainPanel.Controls.Add($labelTime)

    $textboxTime = New-Object System.Windows.Forms.TextBox
    $textboxTime.Location = New-Object System.Drawing.Point(420, 176)
    $textboxTime.Size = New-Object System.Drawing.Size(130, 28)
    $textboxTime.ReadOnly = $true
    $textboxTime.Text = $timePart
    $textboxTime.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $textboxTime.BackColor = $accentSoft
    $textboxTime.ForeColor = $textMain
    $textboxTime.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainPanel.Controls.Add($textboxTime)

    $previewCard = New-Object System.Windows.Forms.Panel
    $previewCard.Location = New-Object System.Drawing.Point(20, 224)
    $previewCard.Size = New-Object System.Drawing.Size(838, 58)
    $previewCard.BackColor = $accentSoft
    $previewCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $mainPanel.Controls.Add($previewCard)

    $labelPreview = New-Object System.Windows.Forms.Label
    $labelPreview.Location = New-Object System.Drawing.Point(12, 6)
    $labelPreview.Size = New-Object System.Drawing.Size(220, 18)
    $labelPreview.Text = "Final file name preview"
    $labelPreview.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Bold)
    $labelPreview.ForeColor = $textMuted
    $previewCard.Controls.Add($labelPreview)

    $textboxPreview = New-Object System.Windows.Forms.TextBox
    $textboxPreview.Location = New-Object System.Drawing.Point(12, 26)
    $textboxPreview.Size = New-Object System.Drawing.Size(810, 24)
    $textboxPreview.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 11, [System.Drawing.FontStyle]::Bold)
    $textboxPreview.ReadOnly = $true
    $textboxPreview.BackColor = [System.Drawing.Color]::White
    $textboxPreview.ForeColor = $accent
    $textboxPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $previewCard.Controls.Add($textboxPreview)

    $updatePreview = {
        $woValue = $textboxWO.Text.Trim().ToUpper()
        if ([string]::IsNullOrWhiteSpace($woValue)) {
            $woValue = "EDIT_ME"
        }
        $textboxPreview.Text = Build-FileName -DatePart $textboxDate.Text -TimePart $textboxTime.Text -WO $woValue
    }

    $textboxWO.Add_TextChanged($updatePreview)
    & $updatePreview

    $footerPanel = New-Object System.Windows.Forms.Panel
    $footerPanel.Location = New-Object System.Drawing.Point(18, 412)
    $footerPanel.Size = New-Object System.Drawing.Size(878, 56)
    $footerPanel.BackColor = $bg
    $form.Controls.Add($footerPanel)

    $acceptButton = New-ModernButton -Text "Apply" -X 240 -Y 8 -BackColor $accent
    $acceptButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $footerPanel.Controls.Add($acceptButton)

    $openFolderButton = New-ModernButton -Text "Open watch folder" -X 384 -Y 8 -Width 150 -BackColor ([System.Drawing.Color]::FromArgb(15, 118, 110))
    $openFolderButton.Add_Click({
        Start-Process explorer.exe $watchFolder
    })
    $footerPanel.Controls.Add($openFolderButton)

    $cancelButton = New-ModernButton -Text "Cancel" -X 548 -Y 8 -BackColor ([System.Drawing.Color]::FromArgb(100, 116, 139))
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $footerPanel.Controls.Add($cancelButton)

    $form.AcceptButton = $acceptButton
    $form.CancelButton = $cancelButton

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textboxPreview.Text
    }

    return $null
}

function Should-IgnoreEvent {
    param(
        [string]$Path
    )

    $now = Get-Date

    $keysToRemove = @()
    foreach ($key in $script:recentEvents.Keys) {
        if (($now - $script:recentEvents[$key]).TotalSeconds -gt 10) {
            $keysToRemove += $key
        }
    }

    foreach ($key in $keysToRemove) {
        $script:recentEvents.Remove($key)
    }

    if ($script:recentEvents.ContainsKey($Path)) {
        return $true
    }

    $script:recentEvents[$Path] = $now
    return $false
}

function Process-PdfFile {
    param(
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if ($script:processedFiles.ContainsKey($fullPath)) {
        return
    }

    if (Should-IgnoreEvent -Path $fullPath) {
        return
    }

    if ($script:isDialogOpen) {
        return
    }

    $script:isDialogOpen = $true

    try {
        Start-Sleep -Seconds 2

        if (-not (Wait-ForFileReady -Path $fullPath)) {
            return
        }

        $newName = Show-RenameDialog -OriginalPath $fullPath

        if ([string]::IsNullOrWhiteSpace($newName)) {
            return
        }

        if (-not $newName.ToLower().EndsWith(".pdf")) {
            $newName += ".pdf"
        }

        if (-not (Test-Path $exportFolder)) {
            New-Item -Path $exportFolder -ItemType Directory -Force | Out-Null
        }

        $targetPath = Join-Path $exportFolder $newName
        $targetPath = Get-AvailableTargetPath -DesiredPath $targetPath

        if ([string]::IsNullOrWhiteSpace($targetPath)) {
            return
        }

        Move-Item -Path $fullPath -Destination $targetPath -Force

        $script:processedFiles[$fullPath]  = $true
        $script:processedFiles[$targetPath] = $true

        Show-NotificationBalloon `
            -Title "File ready to push to Bluebeam" `
            -Body ([System.IO.Path]::GetFileName($targetPath)) `
            -FolderPath $exportFolder `
            -FinalPath $targetPath
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "The file could not be processed.`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $script:isDialogOpen = $false
    }
}

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchFolder
$watcher.Filter = "*.pdf"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true
$watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, CreationTime'

Write-Host "Watching folder: $watchFolder"
Write-Host "Export folder: $exportFolder"
Write-Host "Drop a PDF into the watch folder to test..."

Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier "PdfCreated" -Action {
    $path = $Event.SourceEventArgs.FullPath
    Process-PdfFile -Path $path
} | Out-Null

while ($true) {
    Start-Sleep -Seconds 1
}