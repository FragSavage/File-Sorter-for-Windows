Set-StrictMode -Version Latest

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process -FilePath 'powershell.exe' -Verb RunAs -WorkingDirectory $PSScriptRoot -ArgumentList @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $scriptPath)
    ) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$modulePath = Join-Path -Path $PSScriptRoot -ChildPath 'Sorter.psm1'
Import-Module $modulePath -Force

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'File Sorter'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(620, 360)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'File Sorter'
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(20, 18)
$form.Controls.Add($titleLabel)

$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = 'Choose a folder and pick a grouping mode. The app pulls files up from subfolders into root-level folders, then sends the original subfolders to the Recycle Bin.'
$descriptionLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$descriptionLabel.Size = New-Object System.Drawing.Size(560, 34)
$descriptionLabel.Location = New-Object System.Drawing.Point(22, 58)
$form.Controls.Add($descriptionLabel)

$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Text = 'Grouping mode'
$modeLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$modeLabel.AutoSize = $true
$modeLabel.Location = New-Object System.Drawing.Point(24, 98)
$form.Controls.Add($modeLabel)

$modeComboBox = New-Object System.Windows.Forms.ComboBox
$modeComboBox.DropDownStyle = 'DropDownList'
$modeComboBox.Size = New-Object System.Drawing.Size(220, 24)
$modeComboBox.Location = New-Object System.Drawing.Point(24, 120)
$modeComboBox.Font = New-Object System.Drawing.Font('Segoe UI', 9)
[void]$modeComboBox.Items.Add('By Extension')
[void]$modeComboBox.Items.Add('By Category')
$modeComboBox.SelectedIndex = 0
$form.Controls.Add($modeComboBox)

$modeHelpLabel = New-Object System.Windows.Forms.Label
$modeHelpLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$modeHelpLabel.Size = New-Object System.Drawing.Size(556, 34)
$modeHelpLabel.Location = New-Object System.Drawing.Point(24, 150)
$form.Controls.Add($modeHelpLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Size = New-Object System.Drawing.Size(430, 24)
$pathTextBox.Location = New-Object System.Drawing.Point(24, 194)
$pathTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.Controls.Add($pathTextBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = 'Browse'
$browseButton.Size = New-Object System.Drawing.Size(110, 28)
$browseButton.Location = New-Object System.Drawing.Point(470, 192)
$browseButton.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Controls.Add($browseButton)

$sortButton = New-Object System.Windows.Forms.Button
$sortButton.Text = 'Sort Folder'
$sortButton.Size = New-Object System.Drawing.Size(140, 38)
$sortButton.Location = New-Object System.Drawing.Point(24, 238)
$sortButton.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($sortButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Pick a folder to begin.'
$statusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$statusLabel.Size = New-Object System.Drawing.Size(556, 48)
$statusLabel.Location = New-Object System.Drawing.Point(24, 286)
$form.Controls.Add($statusLabel)

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = 'Select the folder to sort'
if ($null -ne ($folderDialog | Get-Member -Name 'UseDescriptionForTitle' -MemberType Property)) {
    $folderDialog.UseDescriptionForTitle = $true
}

$activeJob = $null
$pollTimer = New-Object System.Windows.Forms.Timer
$pollTimer.Interval = 500

function Update-ModeHelpText {
    if ($modeComboBox.SelectedItem -eq 'By Category') {
        $modeHelpLabel.Text = 'Category mode groups into Images, Video, Project Files, Documents, and Misc.'
        return
    }

    $modeHelpLabel.Text = 'Extension mode creates one folder per extension, such as JPG, PDF, TXT, or No Extension.'
}

function Set-UiBusyState {
    param(
        [Parameter(Mandatory)]
        [bool]$IsBusy
    )

    $sortButton.Enabled = $true
    $browseButton.Enabled = $true
    $form.UseWaitCursor = $false
    if ($IsBusy) {
        $sortButton.Enabled = $false
        $browseButton.Enabled = $false
        $modeComboBox.Enabled = $false
        $form.UseWaitCursor = $true
        return
    }

    $modeComboBox.Enabled = $true
}

$pollTimer.Add_Tick({
    if ($null -eq $activeJob) {
        return
    }

    if ($activeJob.State -eq [System.Management.Automation.JobState]::Running -or
        $activeJob.State -eq [System.Management.Automation.JobState]::NotStarted) {
        return
    }

    $pollTimer.Stop()
    Set-UiBusyState -IsBusy $false

    try {
        $jobOutput = Receive-Job -Job $activeJob -ErrorAction Stop
    }
    catch {
        $statusLabel.Text = 'Sorting failed.'
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'Sorting Failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }
    finally {
        Remove-Job -Job $activeJob -Force -ErrorAction SilentlyContinue
        $script:activeJob = $null
    }

    if ($null -eq $jobOutput) {
        $statusLabel.Text = 'Sorting failed.'
        [System.Windows.Forms.MessageBox]::Show(
            'The sort job finished without returning a result.',
            'Sorting Failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $finalResult = $jobOutput | Select-Object -Last 1
    if (-not $finalResult.Success) {
        $statusLabel.Text = 'Sorting failed.'
        [System.Windows.Forms.MessageBox]::Show(
            $finalResult.ErrorMessage,
            'Sorting Failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $result = $finalResult.Result
    $folderTypeLabel = if ($result.GroupingMode -eq 'Category') { 'category' } else { 'extension' }
    $summary = "Moved $($result.MovedFiles) file(s), created $($result.CreatedFolders) $folderTypeLabel folder(s), and sent $($result.DeletedFolders) original folder(s) to the Recycle Bin."
    if ($result.SkippedDirectories -gt 0) {
        $summary += " Skipped $($result.SkippedDirectories) protected folder(s)."
    }
    if ($result.SkippedFiles -gt 0) {
        $summary += " Skipped $($result.SkippedFiles) protected file(s)."
    }

    $statusLabel.Text = $summary
    [System.Windows.Forms.MessageBox]::Show(
        $summary,
        'Sorting Complete',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
})

$form.Add_FormClosing({
    if ($null -ne $activeJob) {
        $pollTimer.Stop()
        Stop-Job -Job $activeJob -ErrorAction SilentlyContinue
        Remove-Job -Job $activeJob -Force -ErrorAction SilentlyContinue
        $script:activeJob = $null
    }
})

$browseButton.Add_Click({
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathTextBox.Text = $folderDialog.SelectedPath
        $statusLabel.Text = "Selected folder: $($folderDialog.SelectedPath)"
    }
})

$modeComboBox.Add_SelectedIndexChanged({
    Update-ModeHelpText
})

$sortButton.Add_Click({
    $selectedPath = $pathTextBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($selectedPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Choose a folder first.',
            'No Folder Selected',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $selectedPath -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Choose a valid folder path.',
            'Invalid Folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $resolvedSelectedPath = (Resolve-Path -LiteralPath $selectedPath).Path
    if ([System.String]::Equals($resolvedSelectedPath, $PSScriptRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Choose another folder. Sorting the app folder itself would move the sorter files.',
            'Protected Folder',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $groupingMode = if ($modeComboBox.SelectedItem -eq 'By Category') { 'Category' } else { 'Extension' }

    $script:activeJob = Start-Job -ScriptBlock {
        param($jobModulePath, $jobFolderPath, $jobGroupingMode)

        try {
            Import-Module $jobModulePath -Force -ErrorAction Stop
            $jobResult = Move-FolderFiles -FolderPath $jobFolderPath -GroupingMode $jobGroupingMode
            [pscustomobject]@{
                Success = $true
                Result = $jobResult
            }
        }
        catch {
            [pscustomobject]@{
                Success = $false
                ErrorMessage = $_.Exception.Message
            }
        }
    } -ArgumentList $modulePath, $resolvedSelectedPath, $groupingMode

    Set-UiBusyState -IsBusy $true
    $statusLabel.Text = "Sorting files by $groupingMode. This can take a while for large folders."
    $pollTimer.Start()
})

Update-ModeHelpText
[void]$form.ShowDialog()
