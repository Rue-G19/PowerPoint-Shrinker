#requires -Version 5.1
<#
    PowerPoint Bulk Shrinker
    PowerShell 5.1 / WinForms
    Portable - no PowerPoint COM automation.
#>
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (-not ('ModernFolderPicker' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ModernFolderPicker
{[ComImport]
[Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
private class FileOpenDialog { }
[ComImport]
[Guid("d57c7288-d4ad-4768-be02-9d969532d960")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
private interface IFileOpenDialog
{[PreserveSig] int Show(IntPtr parent);
void SetFileTypes(uint count, IntPtr filters);
void SetFileTypeIndex(uint index);
void GetFileTypeIndex(out uint index);
void Advise(IntPtr events, out uint cookie);
void Unadvise(uint cookie);
void SetOptions(uint options);
void GetOptions(out uint options);
void SetDefaultFolder(IShellItem folder);
void SetFolder(IShellItem folder);
void GetFolder(out IShellItem folder);
void GetCurrentSelection(out IShellItem item);
void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string name);
void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string name);
void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string text);
void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
void GetResult(out IShellItem item);
void AddPlace(IShellItem item, int placement);
void RemovePlace(IShellItem item);
void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string extension);
void Close(int hr);
void SetClientGuid(ref Guid guid);
void ClearClientData();
void SetFilter(IntPtr filter);}
[ComImport]
[Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
private interface IShellItem{
void BindToHandler(IntPtr bindCtx, ref Guid handler, ref Guid riid, out IntPtr ppv);
void GetParent(out IShellItem parent);
void GetDisplayName(uint sigdnName, out IntPtr name);
void GetAttributes(uint attributes, out uint mask);
void Compare(IShellItem item, uint hint, out int order);}
private const uint FOS_PICKFOLDERS = 0x00000020;
private const uint SIGDN_FILESYSPATH = 0x80058000;
public static string Pick(IntPtr owner, string title){
var dialog = (IFileOpenDialog)new FileOpenDialog();
try{
uint options;
dialog.GetOptions(out options);
dialog.SetOptions(options | FOS_PICKFOLDERS);
dialog.SetTitle(title);
if (dialog.Show(owner) != 0)
return null;
IShellItem item;
dialog.GetResult(out item);
IntPtr name;
item.GetDisplayName(SIGDN_FILESYSPATH, out name);
try { return Marshal.PtrToStringUni(name); }
finally { Marshal.FreeCoTaskMem(name); }}
finally{
Marshal.ReleaseComObject(dialog);}}}
'@}
$script:IsAdmin = $false
$script:ScriptPath = $PSCommandPath
$ScriptDirectory = Split-Path -Parent $script:ScriptPath
$ImageMagick = Join-Path $ScriptDirectory 'ImageMagick\magick.exe'
function Get-ImageMagickExecutable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedExePath)
    if (Test-Path -LiteralPath $ExpectedExePath -PathType Leaf) {
        return $ExpectedExePath}
    $expectedDir = Split-Path -Parent $ExpectedExePath
    $downloadUrl = 'https://download.imagemagick.org/archive/binaries/ImageMagick-portable-Q16-HDRI-x64.zip'
    $downloadPath = Join-Path $expectedDir 'ImageMagick-portable-Q16-HDRI-x64.zip'
    try {if (-not (Test-Path -LiteralPath $expectedDir -PathType Container)) {$null = New-Item -ItemType Directory -Path $expectedDir -Force}
        [System.Windows.Forms.MessageBox]::Show(
            "ImageMagick was not found. The app will now try to download the portable version automatically.`r`n`r`nExpected path:`r`n$ExpectedExePath",
            "PowerPoint Bulk Shrinker - Downloading ImageMagick",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -UseBasicParsing | Out-Null
        if (Test-Path -LiteralPath $downloadPath -PathType Leaf) {
            Expand-Archive -LiteralPath $downloadPath -DestinationPath $expectedDir -Force}
        $foundExe = Get-ChildItem -Path $expectedDir -Filter 'magick.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($foundExe) {$resolvedExe = $foundExe.FullName
            if (-not (Test-Path -LiteralPath $ExpectedExePath -PathType Leaf)) {
                $null = New-Item -ItemType Directory -Path $expectedDir -Force
                Copy-Item -LiteralPath $resolvedExe -Destination $ExpectedExePath -Force}
            if (Test-Path -LiteralPath $ExpectedExePath -PathType Leaf) {
                return $ExpectedExePath}}}
    catch {$downloadError = $_.Exception.Message}
    $manualMessage = @"
ImageMagick could not be found and the automatic download failed.
Expected location:
$ExpectedExePath

Download the portable Windows build from:
https://imagemagick.org/script/download.php#windows

Then extract the files into the folder beside this script so it ends up as:
$ExpectedExePath

If the download page is blocked, install ImageMagick from the official website and keep the portable 'magick.exe' in the ImageMagick folder next to this script.
"@
    if ($downloadError) {$manualMessage += "`r`n`r`nAuto-download error:`r`n$downloadError"}
    [System.Windows.Forms.MessageBox]::Show(
        $manualMessage,
        "PowerPoint Bulk Shrinker - Dependency Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    $browserUrl = 'https://imagemagick.org/script/download.php#windows'
    try {Start-Process $browserUrl}
    catch {# ignore browser launch failures; the message already tells the user where to go.
        }
    return $null}
$ImageMagick = Get-ImageMagickExecutable -ExpectedExePath $ImageMagick
if (-not $ImageMagick) {exit}
# Winforms fuckery
$form = New-Object System.Windows.Forms.Form
$form.Text = 'PowerPoint Bulk Shrinker v2'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.MinimumSize = New-Object System.Drawing.Size(850, 600)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = 'Folder:'
$lblFolder.Location = New-Object System.Drawing.Point(20, 22)
$lblFolder.AutoSize = $true
$form.Controls.Add($lblFolder)
$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(75, 18)
$txtFolder.Size = New-Object System.Drawing.Size(790, 25)
$txtFolder.Anchor = 'Top,Left,Right'
$form.Controls.Add($txtFolder)
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(875, 17)
$btnBrowse.Size = New-Object System.Drawing.Size(90, 27)
$btnBrowse.Anchor = 'Top,Right'
$form.Controls.Add($btnBrowse)
$groupCompression = New-Object System.Windows.Forms.GroupBox
$groupCompression.Text = 'Compression Level (affects image resizing and JPEG quality, High = smallest file size but will be slow, Low = best image quality, high file size)'
$groupCompression.Location = New-Object System.Drawing.Point(20, 58)
$groupCompression.Size = New-Object System.Drawing.Size(945, 95)
$groupCompression.Anchor = 'Top,Left,Right'
$form.Controls.Add($groupCompression)
$rbHigh = New-Object System.Windows.Forms.RadioButton
$rbHigh.Text = 'High Compression (JPEG 90%)'
$rbHigh.Location = New-Object System.Drawing.Point(25, 38)
$rbHigh.AutoSize = $true
$groupCompression.Controls.Add($rbHigh)
$rbMedium = New-Object System.Windows.Forms.RadioButton
$rbMedium.Text = 'Medium Compression - Recommended (JPEG 75%)'
$rbMedium.Location = New-Object System.Drawing.Point(345, 38)
$rbMedium.AutoSize = $true
$rbMedium.Checked = $true
$groupCompression.Controls.Add($rbMedium)
$rbLow = New-Object System.Windows.Forms.RadioButton
$rbLow.Text = 'Low Compression (JPEG 60%)'
$rbLow.Location = New-Object System.Drawing.Point(675, 38)
$rbLow.AutoSize = $true
$groupCompression.Controls.Add($rbLow)
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Compress Checked'
$btnStart.Location = New-Object System.Drawing.Point(150, 165)
$btnStart.Size = New-Object System.Drawing.Size(160, 38)
$form.Controls.Add($btnStart)
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = 'Scan for PPTX Files'
$btnScan.Location = New-Object System.Drawing.Point(20, 165)
$btnScan.Size = New-Object System.Drawing.Size(120, 38)
$form.Controls.Add($btnScan)
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(315, 165)
$btnCancel.Size = New-Object System.Drawing.Size(120, 38)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)
$btnOpenTemp = New-Object System.Windows.Forms.Button
$btnOpenTemp.Text = 'Open Temp Folder'
$btnOpenTemp.Location = New-Object System.Drawing.Point(445, 165)
$btnOpenTemp.Size = New-Object System.Drawing.Size(145, 38)
$btnOpenTemp.Visible = $false
$form.Controls.Add($btnOpenTemp)
$btnDeleteBaks = New-Object System.Windows.Forms.Button
$btnDeleteBaks.Text = 'Delete .bak Files'
$btnDeleteBaks.Location = New-Object System.Drawing.Point(575, 165)
$btnDeleteBaks.Size = New-Object System.Drawing.Size(120, 38)
$btnDeleteBaks.Visible = $false
$form.Controls.Add($btnDeleteBaks)
$btnCleanup = New-Object System.Windows.Forms.Button
$btnCleanup.Text = 'Cleanup Outputs'
$btnCleanup.Location = New-Object System.Drawing.Point(700, 165)
$btnCleanup.Size = New-Object System.Drawing.Size(120, 38)
$btnCleanup.Visible = $false
$form.Controls.Add($btnCleanup)
$btnOpenLogs = New-Object System.Windows.Forms.Button
$btnOpenLogs.Text = 'Open Logs & Backups'
$btnOpenLogs.Location = New-Object System.Drawing.Point(825, 165)
$btnOpenLogs.Size = New-Object System.Drawing.Size(145, 38)
$form.Controls.Add($btnOpenLogs)
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(300, 241)
$lblStatus.Size = New-Object System.Drawing.Size(650, 22)
$lblStatus.Anchor = 'Top,Left,Right'
$form.Controls.Add($lblStatus)
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 215)
$progress.Size = New-Object System.Drawing.Size(945, 22)
$progress.Anchor = 'Top,Left,Right'
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progress)
$lblProgressPercent = New-Object System.Windows.Forms.Label
$lblProgressPercent.Text = '0% / 100%'
$lblProgressPercent.Location = New-Object System.Drawing.Point(20, 241)
$lblProgressPercent.AutoSize = $true
$form.Controls.Add($lblProgressPercent)
$lblLog = New-Object System.Windows.Forms.Label
$lblFiles = New-Object System.Windows.Forms.Label
$lblFiles.Text = 'Eligible files (15 MB or larger), sorted largest first:'
$lblFiles.Location = New-Object System.Drawing.Point(20, 268)
$lblFiles.AutoSize = $true
$form.Controls.Add($lblFiles)
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = 'Select All'
$btnSelectAll.Location = New-Object System.Drawing.Point(760, 263)
$btnSelectAll.Size = New-Object System.Drawing.Size(90, 24)
$form.Controls.Add($btnSelectAll)
$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Text = 'Select None'
$btnSelectNone.Location = New-Object System.Drawing.Point(855, 263)
$btnSelectNone.Size = New-Object System.Drawing.Size(110, 24)
$form.Controls.Add($btnSelectNone)
$fileList = New-Object System.Windows.Forms.ListView
$fileList.Location = New-Object System.Drawing.Point(20, 290)
$fileList.Size = New-Object System.Drawing.Size(945, 155)
$fileList.Anchor = 'Top,Left,Right'
$fileList.View = [System.Windows.Forms.View]::Details
$fileList.CheckBoxes = $true
$fileList.FullRowSelect = $true
$fileList.GridLines = $true
$fileList.HideSelection = $false
$null = $fileList.Columns.Add('PowerPoint file', 290)
$null = $fileList.Columns.Add('Size', 90)
$null = $fileList.Columns.Add('Path', 535)
$form.Controls.Add($fileList)
$lblLog.Text = 'Verbose Debug Log:'
$lblLog.Location = New-Object System.Drawing.Point(20, 455)
$lblLog.AutoSize = $true
$form.Controls.Add($lblLog)
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 481)
$txtLog.Size = New-Object System.Drawing.Size(945, 152)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)
$lblCredit = New-Object System.Windows.Forms.Label
$lblCredit.Text = 'Designed and created by Rueben Gill'
$lblCredit.Location = New-Object System.Drawing.Point(20, 640)
$lblCredit.AutoSize = $true
$lblCredit.Anchor = 'Bottom,Left'
$form.Controls.Add($lblCredit)
$chkAdminMode = New-Object System.Windows.Forms.CheckBox
$chkAdminMode.Text = 'Enable admin mode'
$chkAdminMode.Location = New-Object System.Drawing.Point(820, 640)
$chkAdminMode.Size = New-Object System.Drawing.Size(145, 24)
$chkAdminMode.Anchor = 'Bottom,Right'
$form.Controls.Add($chkAdminMode)
$script:Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
$script:CancelRequested = $false
$script:WorkerPowerShell = $null
$script:WorkerAsync = $null
$script:WorkerControl = $null
$script:RunState = 'Idle'
$script:RunMode = $null
$script:RequestedMode = $null
$script:EligibleFiles = @()
$script:DebugLogPath = $null
$script:LogWriter = $null
$script:StartTime = $null
$script:TempDirectory = [IO.Path]::GetTempPath()
$script:BackupRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PPTX shrinker backups'
function Add-UILog {
    param([string]$Message)
    $stamp = (Get-Date).ToString('HH:mm:ss.fff')
    $line = "[$stamp] $Message"
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
    if ($script:LogWriter) {try {$script:LogWriter.WriteLine($line)
        $script:LogWriter.Flush()} catch {}}}
function Read-QueueIntoUi {
    $item = $null
    while ($script:Queue.TryDequeue([ref]$item)) {
        if ($item -match '^FOUND\|(\d+)\|(\d+)\|(\d+)$') {
            $found = $matches[1]
            $eligible = $matches[2]
            $ignored = $matches[3]
            $lblStatus.Text = "Found $found PowerPoint files; $eligible to process"
            Add-UILog "SCAN SUMMARY: found $found PowerPoint file(s); $eligible eligible; $ignored ignored under 15 MB"}
        elseif ($item -match '^PROGRESS\|(\d+)\|(\d+)\|(\d+)\|(.*)$') {
            $pct = [Math]::Max(0, [Math]::Min(100, [int]$matches[1]))
            $current = $matches[2]
            $total = $matches[3]
            $name = $matches[4]
            $progress.Value = $pct
            $lblProgressPercent.Text = "$pct% / 100%"
            $lblStatus.Text = "Processing $current / $total - $name"
            Add-UILog "PROGRESS: $pct% | $current / $total | $name"}
        else {Add-UILog $item}
        $item = $null}}
function Get-CompressionSettings {
    if ($rbHigh.Checked) {
        return [pscustomobject]@{
            Name = 'High'
            MaxWidth = 2200
            Quality = 90}}
    if ($rbLow.Checked) {
        return [pscustomobject]@{
            Name = 'Low'
            MaxWidth = 1280
            Quality = 60}}
    return [pscustomobject]@{
        Name = 'Medium'
        MaxWidth = 1600
        Quality = 75}}
$workerCode = {param(
        [string]$Folder,
        [string]$BackupRoot,
        [string]$MagickPath,
        [string]$ProfileName,
        [int]$MaxWidth,
        [int]$Quality,
        [System.Collections.Concurrent.ConcurrentQueue[string]]$Queue,
        [hashtable]$Control,
        [string]$Mode,
        [string[]]$SelectedFiles)
    Set-StrictMode -Version 2.0
    $ErrorActionPreference = 'Stop'
    function Log {param([string]$Message)
        [void]$Queue.Enqueue("WORKER | $Message")}
    function Update-Progress {param(
            [int]$FileIndex,
            [int]$FileTotal,
            [string]$FileName,
            [int]$PhasePercent,
            [string]$Phase)
        $boundedPhase = [Math]::Max(0, [Math]::Min(100, $PhasePercent))
        $overallPercent = [int](((($FileIndex - 1) + ($boundedPhase / 100.0)) / [double]$FileTotal) * 100)
        [void]$Queue.Enqueue("PROGRESS|$overallPercent|$FileIndex|$FileTotal|$FileName - $Phase")}
    function Format-MB {param([Int64]$Bytes)
        return [Math]::Round(($Bytes / 1MB), 2)}
    function Invoke-Magick {param(
            [string[]]$Arguments,
            [string]$WorkingDirectory)
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $MagickPath
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.WorkingDirectory = $WorkingDirectory
        $quotedArgs = foreach ($arg in $Arguments) {
            if ($null -eq $arg) { '""' }
            else { '"' + ($arg.ToString() -replace '(\\*)"', '$1$1\"' -replace '(\\+)$', '$1$1') + '"' }}
        $psi.Arguments = ($quotedArgs -join ' ')
        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        Log ("ImageMagick START: " + $MagickPath + " " + (($Arguments | ForEach-Object { '"' + $_ + '"' }) -join ' '))
        if (-not $p.Start()) {
            throw "Unable to start ImageMagick. Permission issue maybe..?"}
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        $p.WaitForExit()
        if ($stdout.Trim().Length -gt 0) {
            Log ("ImageMagick STDOUT: " + $stdout.Trim().Replace([Environment]::NewLine, ' | '))}
        if ($stderr.Trim().Length -gt 0) {
            Log ("ImageMagick STDERR: " + $stderr.Trim().Replace([Environment]::NewLine, ' | '))}
        Log ("ImageMagick EXIT CODE: " + $p.ExitCode)
        $result = [pscustomobject]@{
            ExitCode = $p.ExitCode
            StdOut = $stdout
            StdErr = $stderr}
        $p.Dispose()
        return $result}
    function Get-ImageDimensions {param([string]$ImagePath)
        $result = Invoke-Magick @(
            'identify',
            '-format',
            '%w %h',
            $ImagePath) (Split-Path -Parent $ImagePath)
        if ($result.ExitCode -ne 0) {throw "ImageMagick identify failed for '$ImagePath'. $($result.StdErr)"} $parts = $result.StdOut.Trim() -split '\s+'
        if ($parts.Count -lt 2) {throw "Could not read image dimensions from ImageMagick for '$ImagePath'. Output: $($result.StdOut)"}
        return [pscustomobject]@{
            Width = [int]$parts[0]
            Height = [int]$parts[1]}}
    function Optimize-Image {
        param(
            [string]$ImagePath,
            [string]$Extension)
        if ($Control.Cancelled) {throw 'Cancellation requested.'}
        $before = (Get-Item -LiteralPath $ImagePath).Length
        Log ("IMAGE: $ImagePath | $(Format-MB $before) MB")
        $isPng = ($Extension -eq '.png')
        $isJpeg = (($Extension -eq '.jpg') -or ($Extension -eq '.jpeg'))
        if (-not ($isPng -or $isJpeg)) {
            Log "IMAGE SKIP: unsupported extension '$Extension'"
            return [pscustomobject]@{
                Processed = $false
                Changed = $false
                Before = $before
                After = $before
                Saved = 0
                Reason = 'Unsupported'}}
$pngColors = 256
if ($ProfileName -eq 'Medium') {
$pngColors = 128} elseif ($ProfileName -eq 'Low') {
$pngColors = 64}
        $tempName = ([IO.Path]::GetFileNameWithoutExtension($ImagePath)) + '.__pptopt_' + [Guid]::NewGuid().ToString('N') + $Extension
        $tempPath = Join-Path (Split-Path -Parent $ImagePath) $tempName
        try {$magickArguments = New-Object System.Collections.Generic.List[string]
            [void]$magickArguments.Add($ImagePath)
            [void]$magickArguments.Add('-resize')
            [void]$magickArguments.Add("${MaxWidth}x>")
            Log "IMAGE RESIZE: capped at max width $MaxWidth; smaller images are left unchanged"
            [void]$magickArguments.Add('-strip')
            if ($isPng) {
                [void]$magickArguments.Add('-colors')
                [void]$magickArguments.Add([string]$pngColors)
                [void]$magickArguments.Add('-define')
                [void]$magickArguments.Add('png:compression-level=9')
                Log "PNG COLOR LIMIT: $pngColors colors"}
            if ($isJpeg) {
                [void]$magickArguments.Add('-quality')
                [void]$magickArguments.Add([string]$Quality)
                Log "JPEG QUALITY: $Quality"}
            [void]$magickArguments.Add($tempPath)
            $result = Invoke-Magick $magickArguments.ToArray() (Split-Path -Parent $ImagePath)
            if ($result.ExitCode -ne 0) {
                throw "ImageMagick failed for '$ImagePath' with exit code $($result.ExitCode). $($result.StdErr)"}
            if (-not (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
                throw "ImageMagick returned success but did not create '$tempPath'."}
            $after = (Get-Item -LiteralPath $tempPath).Length
            Log "IMAGE OUTPUT: $(Format-MB $after) MB"
            if ($after -ge $before) {Log "IMAGE UNCHANGED: shrunk file is not smaller; keeping original ($(Format-MB $before) MB -> $(Format-MB $after) MB)"
                return [pscustomobject]@{
                    Processed = $true
                    Changed = $false
                    Before = $before
                    After = $before
                    Saved = 0
                    Reason = 'Shrunk file was not shrunked :<'}}
            Move-Item -LiteralPath $tempPath -Destination $ImagePath -Force
            Log "IMAGE REPLACED: $ImagePath | saved $(Format-MB ($before - $after)) MB"
            return [pscustomobject]@{
                Processed = $true
                Changed = $true
                Before = $before
                After = $after
                Saved = ($before - $after)
                Reason = 'Shrunk'}}
        finally {if (Test-Path -LiteralPath $tempPath) {Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue}}}
    function Invoke-PresentationOptimization {
        param(
            [System.IO.FileInfo]$File,
            [int]$FileIndex,
            [int]$FileTotal)
        if ($Control.Cancelled) {
            throw 'Cancellation requested.'}
        $originalSize = $File.Length
        $fullPath = $File.FullName
        $outputPath = Join-Path $File.DirectoryName ('Small_' + $File.Name)
        $relativeDirectory = $File.DirectoryName.Substring($Folder.Length).TrimStart('\', '/')
        $backupDirectory = if ([string]::IsNullOrWhiteSpace($relativeDirectory)) {$BackupRoot} else {Join-Path $BackupRoot $relativeDirectory}
        $backupPath = Join-Path $backupDirectory ($File.Name + '.bak')
        $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('PPTShrinker_' + [Guid]::NewGuid().ToString('N'))
        $extractDir = Join-Path $tempRoot 'extracted'
        $rebuiltPath = Join-Path $tempRoot ($File.Name + '.new')
        Log ('------------------------')
        Log ("PROCESSING: $fullPath")
        Log ("ORIGINAL SIZE: $(Format-MB $originalSize) MB")
        Update-Progress $FileIndex $FileTotal $File.Name 0 'starting'
        Log "BACKUP PATH: $backupPath"
        Log "OUTPUT PATH: $outputPath"
        Log "TEMP ROOT: $tempRoot"
        try {New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
            New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
            Log "BACKUP: saving original before processing"
            Copy-Item -LiteralPath $fullPath -Destination $backupPath -Force
            Log "BACKUP: complete"
            Update-Progress $FileIndex $FileTotal $File.Name 10 'backup complete'
            Log "ZIP EXTRACT: starting"
            [IO.Compression.ZipFile]::ExtractToDirectory($fullPath, $extractDir)
            Log "ZIP EXTRACT: complete"
            Update-Progress $FileIndex $FileTotal $File.Name 20 'extraction complete'
            $mediaDir = Join-Path $extractDir 'ppt\media'
            if (-not (Test-Path -LiteralPath $mediaDir -PathType Container)) {Log "MEDIA: ppt\media does not exist. No images to process."} 
            else {$mediaFiles = @(Get-ChildItem -LiteralPath $mediaDir -File -ErrorAction Stop)
                Log "MEDIA: found $($mediaFiles.Count) file(s) in ppt\media"
                $imageIndex = 0
                foreach ($media in $mediaFiles) {if ($Control.Cancelled) {throw 'Cancellation requested.'}
                    $imageIndex++
                    $ext = $media.Extension.ToLowerInvariant()
                    if (($ext -eq '.png') -or ($ext -eq '.jpg') -or ($ext -eq '.jpeg')) {try {
                            $mediaStartPercent = if ($mediaFiles.Count -gt 0) {[int](($imageIndex - 1) / [double]$mediaFiles.Count * 100)} else {0}
                            $mediaStartPhase = 20 + [int]($mediaStartPercent * 0.45)
                            Update-Progress $FileIndex $FileTotal $File.Name $mediaStartPhase "image $imageIndex / $($mediaFiles.Count) starting"
                            Optimize-Image -ImagePath $media.FullName -Extension $ext | Out-Null} catch {
                            Log ("IMAGE ERROR: " + $media.FullName + " | " + $_.Exception.Message)throw}} else {
                        Log "MEDIA SKIP: $($media.Name) | extension '$ext' is not supported"}
                    $mediaPercent = if ($mediaFiles.Count -gt 0) {[int](($imageIndex / [double]$mediaFiles.Count) * 100)} else {100}
                    $phasePercent = 20 + [int]($mediaPercent * 0.45)
                        Update-Progress $FileIndex $FileTotal $File.Name $phasePercent "image $imageIndex / $($mediaFiles.Count)"}}
                    Update-Progress $FileIndex $FileTotal $File.Name 65 'media complete'
            Log "ZIP REBUILD: creating $rebuiltPath"
            if (Test-Path -LiteralPath $rebuiltPath) {
                Remove-Item -LiteralPath $rebuiltPath -Force}
            [IO.Compression.ZipFile]::CreateFromDirectory(
                $extractDir,
                $rebuiltPath,
                [IO.Compression.CompressionLevel]::Optimal,
                $false)
        if (-not (Test-Path -LiteralPath $rebuiltPath -PathType Leaf)) {
        throw "ZIP rebuild did not produce an output file."}
$newSize = (Get-Item -LiteralPath $rebuiltPath).Length
            Log "ZIP REBUILD: complete | $(Format-MB $newSize) MB"
            Update-Progress $FileIndex $FileTotal $File.Name 78 'rebuild complete'
            Log "VALIDATION: opening rebuilt ZIP"
            $zip = [IO.Compression.ZipFile]::OpenRead($rebuiltPath)
            try {$required = @('ppt/presentation.xml', '[Content_Types].xml')
                $actualEntries = @($zip.Entries | ForEach-Object {$_.FullName.TrimStart('./').Replace('\', '/')})
                foreach ($entryName in $required) {if ($actualEntries -notcontains $entryName) {throw "Rebuilt PowerPoint file is missing required ZIP entry '$entryName'."}}
                Log "VALIDATION: required PowerPoint entries present"}
            finally {$zip.Dispose()}
            Update-Progress $FileIndex $FileTotal $File.Name 88 'validation complete'
            if ($newSize -ge $originalSize) {
                Log "RESULT: $($File.Name) skipped because rebuilt file is not smaller ($(Format-MB $originalSize) MB -> $(Format-MB $newSize) MB)"
                return [pscustomobject]@{
                    File = $fullPath
                    OutputFile = $null
                    OriginalMB = [Math]::Round(($originalSize / 1MB), 2)
                    NewMB = [Math]::Round(($originalSize / 1MB), 2)
                    SavedMB = 0
                    PercentReduction = 0
                    Status = 'Skipped: rebuilt file was not smaller'}}
            Log "OUTPUT: creating $outputPath"
            Copy-Item -LiteralPath $rebuiltPath -Destination $outputPath -Force
            Log "OUTPUT: complete"
            Update-Progress $FileIndex $FileTotal $File.Name 94 'output created'
            Log "DELETE: removing original PowerPoint file after validated output creation"
            Remove-Item -LiteralPath $fullPath -Force -ErrorAction Stop
            Log "DELETE: original removed"
            $finalSize = (Get-Item -LiteralPath $outputPath).Length
            $saved = $originalSize - $finalSize
            $pct = if ($originalSize -gt 0) { ($saved / [double]$originalSize) * 100 } else { 0 }
            Log "RESULT: $($File.Name) -> $(Split-Path -Leaf $outputPath) | $(Format-MB $originalSize) MB -> $(Format-MB $finalSize) MB | saved $(Format-MB $saved) MB | $([Math]::Round($pct, 2))%"
            Update-Progress $FileIndex $FileTotal $File.Name 100 'complete'
            return [pscustomobject]@{
                File = $fullPath
                OutputFile = $outputPath
                OriginalMB = [Math]::Round(($originalSize / 1MB), 2)
                NewMB = [Math]::Round(($finalSize / 1MB), 2)
                SavedMB = [Math]::Round(($saved / 1MB), 2)
                PercentReduction = [Math]::Round($pct, 2)
                Status = 'Success'}}
        catch {Log ("PRESENTATION ERROR: $fullPath | " + $_.Exception.GetType().FullName + " | " + $_.Exception.Message)
            Log ("STACK: " + $_.ScriptStackTrace)
            if (Test-Path -LiteralPath $outputPath -PathType Leaf) {try {Remove-Item -LiteralPath $outputPath -Force -ErrorAction Stop
                    Log "OUTPUT CLEANUP: removed incomplete output"} catch {Log ("OUTPUT CLEANUP ERROR: " + $_.Exception.Message)}}
            return [pscustomobject]@{
                File = $fullPath
                OutputFile = $outputPath
                OriginalMB = [Math]::Round(($originalSize / 1MB), 2)
                NewMB = [Math]::Round(($originalSize / 1MB), 2)
                SavedMB = 0
                PercentReduction = 0
                Status = 'Error: ' + $_.Exception.Message}}
        finally {try {if (Test-Path -LiteralPath $tempRoot) {Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue} } catch {
                Log ("TEMP CLEANUP ERROR: " + $_.Exception.Message)}}}
    $results = New-Object System.Collections.Generic.List[object]
    try {Log '------------------------------------------'
        Log 'RG-WORKER STARTED'
        Log "PowerShell: $($PSVersionTable.PSVersion)"
        Log "Folder: $Folder"
        Log "Backup root: $BackupRoot"
        Log "ImageMagick: $MagickPath"
        Log "Profile: $ProfileName"
        Log "Max width: $MaxWidth"
        Log "JPEG quality: $Quality"
        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {throw "Selected folder does not exist or is not accessible: $Folder"}
        if (-not (Test-Path -LiteralPath $MagickPath -PathType Leaf)) {throw "ImageMagick executable does not exist: $MagickPath"}
        New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
        Log 'SCAN: starting recursive *.pptx and *.pptm search'
        $discoveredFiles = @(Get-ChildItem -LiteralPath $Folder -File -Recurse -ErrorAction Stop |
            Where-Object { $_.Name -notlike 'Small_*' -and (($_.Extension -ieq '.pptx') -or ($_.Extension -ieq '.pptm')) })
        # Blanket rule: ignore PowerPoint files smaller than 15 MB.
        $minimumFileSize = 15MB
        $files = @($discoveredFiles | Where-Object { $_.Length -ge $minimumFileSize })
        $ignoredSmallFiles = @($discoveredFiles | Where-Object { $_.Length -lt $minimumFileSize })
        if ($SelectedFiles -and $SelectedFiles.Count -gt 0) {
            $selectedLookup = @{}
            foreach ($selectedPath in $SelectedFiles) {$selectedLookup[$selectedPath.ToLowerInvariant()] = $true}
            $files = @($files | Where-Object { $selectedLookup.ContainsKey($_.FullName.ToLowerInvariant()) })}
        Log "SCAN: COMPLETE. Found $($discoveredFiles.Count) PowerPoint file(s)."
        Log "SCAN: eligible files at least 15 MB: $($files.Count)"
        Log "SCAN: ignored files under 15 MB: $($ignoredSmallFiles.Count)"
        [void]$Queue.Enqueue("FOUND|$($discoveredFiles.Count)|$($files.Count)|$($ignoredSmallFiles.Count)")
        foreach ($smallFile in $ignoredSmallFiles) {
            Log "IGNORED (< 15 MB): $($smallFile.FullName) | $(Format-MB $smallFile.Length) MB"}
        if ($Mode -eq 'Scan') {
            [pscustomobject]@{
                Success = $true
                ScanOnly = $true
                Folder = $Folder
                FilesFound = $discoveredFiles.Count
                FilesEligible = $files.Count
                FilesIgnored = $ignoredSmallFiles.Count
                EligibleFiles = @($files | Select-Object FullName, Name, Length)
                Results = @()}
            return}
        if ($files.Count -eq 0) {Log "SCAN WARNING: no .pptx or .pptm files found under '$Folder'" } else {
            for ($i = 0; $i -lt $files.Count; $i++) {if ($Control.Cancelled) {Log 'CANCELLATION: requested before next presentation.'break}
                $file = $files[$i]
                Log "FILE [$($i + 1)/$($files.Count)]: $($file.FullName)"
                Log "FILE SIZE: $(Format-MB $file.Length) MB"
                $result = Invoke-PresentationOptimization -File $file -FileIndex ($i + 1) -FileTotal $files.Count
                $results.Add($result) | Out-Null
                Update-Progress ($i + 1) $files.Count $file.Name 100 'complete'}}
        $successful = @($results | Where-Object { $_.Status -eq 'Success' })
        $failed = @($results | Where-Object { $_.Status -like 'Error:*' })
        $skipped = @($results | Where-Object { $_.Status -like 'Skipped:*' })
        $originalTotal = 0.0
        $newTotal = 0.0
        foreach ($r in $results) {
            $originalTotal += [double]$r.OriginalMB
            $newTotal += [double]$r.NewMB}
        $savedTotal = $originalTotal - $newTotal
        $pctTotal = if ($originalTotal -gt 0) { ($savedTotal / $originalTotal) * 100 } else { 0 }
        Log '----------------------------------'
        Log 'RG-WORKER COMPLETE'
        Log "Files discovered: $($discoveredFiles.Count)"
        Log "Files eligible (15 MB or larger): $($files.Count)"
        Log "Files ignored (under 15 MB): $($ignoredSmallFiles.Count)"
        Log "Files processed: $($results.Count)"
        Log "Successful: $($successful.Count)"
        Log "Failed: $($failed.Count)"
        Log "Skipped (not smaller): $($skipped.Count)"
        Log "Original total: $([Math]::Round($originalTotal, 2)) MB"
        Log "New total: $([Math]::Round($newTotal, 2)) MB"
        Log "Saved total: $([Math]::Round($savedTotal, 2)) MB"
        Log "Percentage reduction: $([Math]::Round($pctTotal, 2))%"
        $resultItems = @($results.ToArray())
        [pscustomobject]@{
            Success = $true
            Cancelled = [bool]$Control.Cancelled
            Folder = $Folder
            FilesFound = $discoveredFiles.Count
            FilesEligible = $files.Count
            FilesIgnored = $ignoredSmallFiles.Count
            FilesProcessed = $results.Count
            Successful = $successful.Count
            Failed = $failed.Count
            Skipped = $skipped.Count
            OriginalMB = [Math]::Round($originalTotal, 2)
            NewMB = [Math]::Round($newTotal, 2)
            SavedMB = [Math]::Round($savedTotal, 2)
            PercentReduction = [Math]::Round($pctTotal, 2)
            Results = $resultItems}}
    catch {
        Log ("*** FATAL WORKER ERROR ***")
        Log ("TYPE: " + $_.Exception.GetType().FullName)
        Log ("MESSAGE: " + $_.Exception.Message)
        Log ("SCRIPT STACK: " + $_.ScriptStackTrace)
        [pscustomobject]@{
            Success = $false
            Cancelled = [bool]$Control.Cancelled
            Error = $_.Exception.Message
            FilesFound = 0
            FilesProcessed = $results.Count
            Successful = 0
            Failed = $results.Count
            OriginalMB = 0
            NewMB = 0
            SavedMB = 0
            PercentReduction = 0
            Results = @($results.ToArray())}}}
$btnBrowse.Add_Click({$selectedFolder = [ModernFolderPicker]::Pick($form.Handle, 'Select the folder containing the PowerPoint files')
    if (-not [string]::IsNullOrWhiteSpace($selectedFolder)) {$txtFolder.Text = $selectedFolder}})
$btnSelectAll.Add_Click({foreach ($item in $fileList.Items) {$item.Checked = $true}})
$btnSelectNone.Add_Click({foreach ($item in $fileList.Items) {$item.Checked = $false}})
$chkAdminMode.Add_CheckedChanged({
    if ($chkAdminMode.Checked -and -not $script:IsAdmin) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Admin mode exposes maintenance tools that can open temporary working files and permanently delete .bak backup files.`r`n`r`nOnly enable this mode if you understand these risks and are responsible for the selected files. Continue?",
            'Enable Admin Mode - Risk Warning',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $chkAdminMode.Checked = $false
            return}
        $script:IsAdmin = $true
        $btnOpenTemp.Visible = $true
        $btnDeleteBaks.Visible = $true
        $btnCleanup.Visible = $true
        $lblStatus.Text = 'Admin mode enabled'}
    elseif (-not $chkAdminMode.Checked) {
        $script:IsAdmin = $false
        $btnOpenTemp.Visible = $false
        $btnDeleteBaks.Visible = $false
        $btnCleanup.Visible = $false
        if ($script:RunState -ne 'Running') {$lblStatus.Text = 'Admin mode disabled'}}})
$btnOpenTemp.Add_Click({if (-not $script:IsAdmin) {return}
    try {Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:TempDirectory.TrimEnd('\')) } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not open the temp folder:`r`n$($script:TempDirectory)`r`n`r`n$($_.Exception.Message)",
            'PowerPoint Bulk Shrinker - Temp Folder Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}})
$btnOpenLogs.Add_Click({try {if (-not (Test-Path -LiteralPath $script:BackupRoot -PathType Container)) {$null = New-Item -ItemType Directory -Path $script:BackupRoot -Force}
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:BackupRoot.TrimEnd('\')) } catch {[System.Windows.Forms.MessageBox]::Show(
            "Could not open the logs and backups folder:`r`n$($script:BackupRoot)`r`n`r`n$($_.Exception.Message)",
            'PowerPoint Bulk Shrinker - Folder Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}})
$btnDeleteBaks.Add_Click({if (-not $script:IsAdmin -or $script:RunState -eq 'Running') {return}
    $folder = $txtFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder) -or
        (-not (Test-Path -LiteralPath $folder -PathType Container))) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please select an existing folder first.',
            'PowerPoint Bulk Shrinker',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return}
    try {$searchRoots = @($folder)
        if (Test-Path -LiteralPath $script:BackupRoot -PathType Container) {
            $searchRoots += $script:BackupRoot}
        $bakFiles = @($searchRoots | ForEach-Object {
            Get-ChildItem -LiteralPath $_ -Filter '*.bak' -File -Recurse -ErrorAction Stop})
        if ($bakFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'No .bak files were found in the selected folder, its subfolders, or the RG PPTX shrinker backup folder.',
                'PowerPoint Bulk Shrinker',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return}
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Delete $($bakFiles.Count) .bak file(s) from the selected folder and the RG PPTX shrinker backup folder?`r`n`r`n$folder`r`n$script:BackupRoot",
            'Confirm .bak File Deletion',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            return}
        $deleted = 0
        foreach ($bakFile in $bakFiles) {
            Remove-Item -LiteralPath $bakFile.FullName -Force -ErrorAction Stop
            $deleted++}
        Add-UILog "BAK CLEANUP: deleted $deleted file(s) from $folder"
        [System.Windows.Forms.MessageBox]::Show(
            "Deleted $deleted .bak file(s).",
            'PowerPoint Bulk Shrinker - Cleanup Complete',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } catch {
        Add-UILog ("BAK CLEANUP ERROR: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            "Could not delete all .bak files:`r`n$($_.Exception.Message)",
            'PowerPoint Bulk Shrinker - Cleanup Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}})
$btnCleanup.Add_Click({if (-not $script:IsAdmin -or $script:RunState -eq 'Running') {return}
    $folder = $txtFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder) -or
        (-not (Test-Path -LiteralPath $folder -PathType Container))) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please select an existing folder first.',
            'PowerPoint Bulk Shrinker',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return}
    try {$outputFiles = @(Get-ChildItem -LiteralPath $folder -Filter 'Small_*.pptx' -File -Recurse -ErrorAction Stop)
        $reportPath = Join-Path $folder 'PPT_Shrinking_Report.csv'
        if ($outputFiles.Count -eq 0 -and -not (Test-Path -LiteralPath $reportPath -PathType Leaf)) {
            [System.Windows.Forms.MessageBox]::Show(
                'No Small_ PowerPoint files or shrinking report were found in the selected folder or its subfolders.',
                'PowerPoint Bulk Shrinker',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return}
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Delete the shrinking report and rename $($outputFiles.Count) Small_ PowerPoint file(s) in:`r`n$folder",
            'Confirm Output Cleanup',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {return}
        $renamed = 0
        $renameFailures = @()
        foreach ($outputFile in $outputFiles) {
            $destinationName = $outputFile.Name.Substring(6)
            try {Rename-Item -LiteralPath $outputFile.FullName -NewName $destinationName -ErrorAction Stop
                $renamed++} catch {$renameFailures += "$($outputFile.FullName): $($_.Exception.Message)"}}
        $reportDeleted = $false
        if (Test-Path -LiteralPath $reportPath -PathType Leaf) {
            Remove-Item -LiteralPath $reportPath -Force -ErrorAction Stop
            $reportDeleted = $true}
        Add-UILog "OUTPUT CLEANUP: renamed $renamed file(s), report deleted: $reportDeleted"
        if ($renameFailures.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Cleanup finished with $($renameFailures.Count) rename failure(s).`r`n`r`n$($renameFailures -join "`r`n")",
                'PowerPoint Bulk Shrinker - Cleanup Warning',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null}
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Cleanup complete. Renamed $renamed file(s) and deleted the report: $reportDeleted.",
                'PowerPoint Bulk Shrinker - Cleanup Complete',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null}} catch {
        Add-UILog ("OUTPUT CLEANUP ERROR: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            "Could not complete output cleanup:`r`n$($_.Exception.Message)",
            'PowerPoint Bulk Shrinker - Cleanup Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}})
$btnCancel.Add_Click({
    if ($script:RunState -eq 'Running') {
        $script:CancelRequested = $true
        $btnCancel.Enabled = $false
        $lblStatus.Text = 'Cancellation requested...'
        Add-UILog 'CANCEL: cancellation requested by user.'
        if ($script:WorkerPowerShell) {try {} catch {}}}})
$btnScan.Add_Click({
    if ($script:RunState -eq 'Running') {return}
    $script:RequestedMode = 'Scan'
    $btnStart.PerformClick()})
$btnStart.Add_Click({
    if ($script:RunState -eq 'Running') {return}
    $mode = if ($script:RequestedMode) {$script:RequestedMode} else {'Compress'}
    $script:RequestedMode = $null
    $folder = $txtFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) {
        [System.Windows.Forms.MessageBox]::Show(
            'pick a folder pls',
            'What folder?',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null return}
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show(
            "The selected folder does not exist or cannot be accessed:`r`n$folder",
            'What folder?',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null return}
    try {
        $folder = (Resolve-Path -LiteralPath $folder -ErrorAction Stop).Path} catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not resolve the selected folder:`r`n$($_.Exception.Message)",
            'What folder?',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null return}
    if ($mode -eq 'Compress') {
        $selectedFiles = @($fileList.CheckedItems | ForEach-Object {$_.Tag})
        if ($selectedFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'Scan the folder and check at least one PowerPoint file first.',
                'No files selected',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            return}
        [System.Windows.Forms.MessageBox]::Show(
            "WARNING: Each checked original PowerPoint file will be backed up to your Documents folder, then deleted after its Small_ replacement is created and validated.`r`n`r`nFor best performance, please close other demanding applications before continuing.",
            'PowerPoint Bulk Shrinker - Compression Starting',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } else {$selectedFiles = @()}
    $settings = Get-CompressionSettings
    $backupRoot = $script:BackupRoot
    try {if (-not (Test-Path -LiteralPath $backupRoot -PathType Container)) {$null = New-Item -ItemType Directory -Path $backupRoot -Force}} catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not create the logs and backups folder:`r`n$backupRoot`r`n`r`n$($_.Exception.Message)",
            'PowerPoint Bulk Shrinker - Folder Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null return}
    $script:DebugLogPath = Join-Path $backupRoot 'PPT_Optimization_Debug.log'
    # close old files before opening a new log file
    if ($script:LogWriter) {try { $script:LogWriter.Dispose() } catch {}
        $script:LogWriter = $null}
    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $script:LogWriter = New-Object System.IO.StreamWriter(
            $script:DebugLogPath,
            $true,
            $utf8NoBom)} catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not open the debug log for writing:`r`n$($script:DebugLogPath)`r`n`r`n$($_.Exception.Message)",
            'PowerPoint Bulk Shrinker - Log Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null return}
    $script:CancelRequested = $false
    $script:RunState = 'Running'
    $script:RunMode = $mode
    $script:StartTime = Get-Date
    $btnStart.Enabled = $false
    $btnScan.Enabled = $false
    $btnBrowse.Enabled = $false
    $btnDeleteBaks.Enabled = $false
    $btnCleanup.Enabled = $false
    $chkAdminMode.Enabled = $false
    $btnCancel.Enabled = $true
    $rbHigh.Enabled = $false
    $rbMedium.Enabled = $false
    $rbLow.Enabled = $false
    $progress.Value = 0
    $lblProgressPercent.Text = '0% / 100%'
    $lblStatus.Text = 'Starting...'
    Add-UILog '-----------------------------------------'
    Add-UILog 'PowerPoint Bulk Shrinker START'
    Add-UILog "Start time: $($script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Add-UILog "Compression profile: $($settings.Name)"
    Add-UILog "Maximum image width: $($settings.MaxWidth)"
    Add-UILog "I refuse to take responsibility for any data loss. By continuing, you agree to this risk.`r`nBackups will be saved to: $backupRoot (please verify this folder exists and is writable)"
    Add-UILog "PowerShell version: $($PSVersionTable.PSVersion)"
    Add-UILog "Script path: $script:ScriptPath"
    Add-UILog "Script directory: $ScriptDirectory"
    Add-UILog "Selected folder: $folder"
    Add-UILog "Backup root: $backupRoot"
    Add-UILog "Folder exists: $(Test-Path -LiteralPath $folder -PathType Container)"
    Add-UILog "ImageMagick: $ImageMagick"
    Add-UILog "ImageMagick exists: $(Test-Path -LiteralPath $ImageMagick -PathType Leaf)"
    Add-UILog "Compression profile: $($settings.Name)"
    Add-UILog "Maximum image width: $($settings.MaxWidth)"
    Add-UILog "JPEG quality: $($settings.Quality)"
    Add-UILog "Verbose debug log: $script:DebugLogPath"
    Add-UILog 'Creating dedicated PowerShell runspace...'
    $script:Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    # Hashtable shared between the UI and worker runspace..
    $control = [hashtable]::Synchronized(@{
        Cancelled = $false})
    $script:WorkerControl = $control
    try {
        $script:WorkerPowerShell = [PowerShell]::Create()
        $script:WorkerPowerShell.Runspace = [RunspaceFactory]::CreateRunspace()
        $script:WorkerPowerShell.Runspace.Open()
        Add-UILog 'Dedicated runspace opened successfully.'
        Add-UILog "Runspace state: $($script:WorkerPowerShell.Runspace.RunspaceStateInfo.State)"
        [void]$script:WorkerPowerShell.AddScript($workerCode)
        [void]$script:WorkerPowerShell.AddArgument($folder)
        [void]$script:WorkerPowerShell.AddArgument($backupRoot)
        [void]$script:WorkerPowerShell.AddArgument($ImageMagick)
        [void]$script:WorkerPowerShell.AddArgument($settings.Name)
        [void]$script:WorkerPowerShell.AddArgument($settings.MaxWidth)
        [void]$script:WorkerPowerShell.AddArgument($settings.Quality)
        [void]$script:WorkerPowerShell.AddArgument($script:Queue)
        [void]$script:WorkerPowerShell.AddArgument($control)
        [void]$script:WorkerPowerShell.AddArgument($mode)
        [void]$script:WorkerPowerShell.AddArgument([string[]]$selectedFiles)
        # We love Async pipelines! woohoo! 
        Add-UILog 'Starting asynchronous PowerShell pipeline...'
        $script:WorkerAsync = $script:WorkerPowerShell.BeginInvoke()
        Add-UILog "BeginInvoke completed. IsCompleted: $($script:WorkerAsync.IsCompleted)"
        $timer.Start()
        Add-UILog 'UI timer started; GUI remains responsive while runspace processes.'
        $lblStatus.Text = if ($mode -eq 'Scan') {'Scanning...'} else {'Compressing...'}}
    catch {
        Add-UILog '*** FAILED TO START WORKER ***'
        Add-UILog ("TYPE: " + $_.Exception.GetType().FullName)
        Add-UILog ("MESSAGE: " + $_.Exception.Message)
        Add-UILog ("STACK: " + $_.ScriptStackTrace)
        $script:RunState = 'Error'
        $btnStart.Enabled = $true
        $btnScan.Enabled = $true
        $btnBrowse.Enabled = $true
        $btnDeleteBaks.Enabled = $true
        $btnCleanup.Enabled = $true
        $chkAdminMode.Enabled = $true
        $btnCancel.Enabled = $false
        $rbHigh.Enabled = $true
        $rbMedium.Enabled = $true
        $rbLow.Enabled = $true
        $lblStatus.Text = 'Error' #fuck you powershell
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'PowerPoint Bulk Shrinker - Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}})
# timer polls the worker and drains the queue on the UI thread
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 150
$timer.Add_Tick({
    Read-QueueIntoUi
    if ($script:RunState -ne 'Running') {return}
    # shared cancellation flag synchro.
    if ($script:CancelRequested -and $script:WorkerControl) {
        $script:WorkerControl.Cancelled = $true}
    if ($script:WorkerAsync -and $script:WorkerAsync.IsCompleted) {
        $timer.Stop()
        Read-QueueIntoUi
        $lblStatus.Text = 'Finalising...'
        Add-UILog 'Worker pipeline reports IsCompleted = True.'
        Add-UILog "Async result IsCompleted: $($script:WorkerAsync.IsCompleted) Wow :D"
        $result = $null
        $endInvokeError = $null
        try {$result = $script:WorkerPowerShell.EndInvoke($script:WorkerAsync)}
        catch {$endInvokeError = $_
            Add-UILog '*** EndInvoke ERROR ***'
            Add-UILog ("TYPE: " + $_.Exception.GetType().FullName)
            Add-UILog ("MESSAGE: " + $_.Exception.Message)
            Add-UILog ("STACK: " + $_.ScriptStackTrace)}
        Read-QueueIntoUi
        if ($endInvokeError) {
            $script:RunState = 'Error'
            $lblStatus.Text = 'Error'
            [System.Windows.Forms.MessageBox]::Show(
            $endInvokeError.Exception.Message,
            'PowerPoint Bulk Shrinker - Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}
        elseif (-not $result -or $result.Count -eq 0) {
            Add-UILog '*** FATAL: Worker returned no result object. ***'
            $script:RunState = 'Error'
            $lblStatus.Text = 'Error - no result returned'
            [System.Windows.Forms.MessageBox]::Show(
            'The worker returned no result. I have no idea how you managed to trigger this but good job! Check the verbose debug log ig.',
            'I am Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}
        else {Add-UILog "Worker returned $($result.Count) pipeline object(s)."
            $summary = $result[-1]
            Add-UILog "Final worker result type: $($summary.GetType().FullName)"
            if ($summary.Success -and $summary.PSObject.Properties['ScanOnly'] -and [bool]$summary.ScanOnly) {
                $fileList.Items.Clear()
                $script:EligibleFiles = @($summary.EligibleFiles | Sort-Object Length -Descending)
                foreach ($eligibleFile in $script:EligibleFiles) {
                    $item = New-Object -TypeName System.Windows.Forms.ListViewItem -ArgumentList $eligibleFile.Name
                    $null = $item.SubItems.Add(('{0:N2} MB' -f ($eligibleFile.Length / 1MB)))
                    $null = $item.SubItems.Add($eligibleFile.FullName)
                    $item.Tag = $eligibleFile.FullName
                    $item.Checked = $true
                    $null = $fileList.Items.Add($item)}
                $script:RunState = 'Complete'
                $progress.Value = 100
                $lblProgressPercent.Text = '100% / 100%'
                $lblStatus.Text = "Found $($summary.FilesEligible) eligible file(s)"
                Add-UILog "SCAN RESULTS: $($summary.FilesEligible) eligible file(s) loaded; all are checked by default."
            } elseif ($summary.Success) {$script:RunState = 'Complete'
                $progress.Value = 100
                $lblProgressPercent.Text = '100% / 100%'
                $lblStatus.Text = if ($summary.Cancelled) { 'Cancelled' } else { 'Finished' }
                Add-UILog '-------------------------------------------'
                Add-UILog 'COMPLETION SUMMARY'
                Add-UILog "Finished time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Add-UILog ("Time taken: {0:hh\:mm\:ss}" -f ((Get-Date) - $script:StartTime))
                Add-UILog "Files Found: $($summary.FilesFound)"
                Add-UILog "Files Eligible (15 MB or larger): $($summary.FilesEligible)"
                Add-UILog "Files Ignored (under 15 MB): $($summary.FilesIgnored)"
                Add-UILog "Files Processed: $($summary.FilesProcessed)"
                Add-UILog "Files Skipped (not smaller): $($summary.Skipped)"
                Add-UILog "Original Size: $($summary.OriginalMB) MB"
                Add-UILog "New Size: $($summary.NewMB) MB"
                Add-UILog "Saved Size: $($summary.SavedMB) MB"
                Add-UILog "Percentage Reduction: $($summary.PercentReduction)%"
                Add-UILog "CSV Report Location (in case you wanted it...?): $(Join-Path $summary.Folder 'PPT_Shrinking_Report.csv')"
                Add-UILog "Pretty neat, eh?" 
                # CSV report IS OVER HERE 
                $csvPath = Join-Path $summary.Folder 'PPT_Shrinking_Report.csv'
                try {$summary.Results |
                Select-Object File, OutputFile, OriginalMB, NewMB, SavedMB, PercentReduction |
                Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
                Add-UILog "CSV: created successfully: $csvPath"}
                catch {Add-UILog ("CSV ERROR: " + $_.Exception.Message)}
                $elapsed = (Get-Date) - $script:StartTime
                Add-UILog ("Elapsed time: {0:hh\:mm\:ss}" -f $elapsed)
                Add-UILog 'PowerPoint Bulk Shrinker FINISHED'
                $savedMessage = if ([double]$summary.SavedMB -gt 0) {
                "Saved space: $($summary.SavedMB) MB ($($summary.PercentReduction)%)"} else {
                'Saved space: 0 MB (the rebuilt files were not smaller than the originals) What a waste of time :<'}
                $message = @"
Files Processed: $($summary.FilesProcessed)
Original Size: $($summary.OriginalMB) MB
New Size: $($summary.NewMB) MB
$savedMessage
CSV Report:
$csvPath
Debug Log:
$script:DebugLogPath
"@
[System.Windows.Forms.MessageBox]::Show(
    $message,
    'PowerPoint Bulk Shrinker - Complete',
        [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null}
            else {$script:RunState = 'Error'
                $lblStatus.Text = 'Error'
                Add-UILog '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=--'
                Add-UILog 'WORKER FAILED. HE IS ON STRIKE. HE REFUSES TO WORK. HE IS MAD. HE IS ANGRY. HE IS SAD. HE IS TIRED. HE IS DONE. HE IS DRAMATIC. HE IS ERROR.'
                Add-UILog "Error: $($summary.Error)"
                [System.Windows.Forms.MessageBox]::Show(
                    "$($summary.Error)`r`n`r`nSee the verbose debug log for full details:`r`n$script:DebugLogPath",
                    'PowerPoint Bulk Shrinker - Error',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null}}
        # bin the pipeline and runspace, all done now
        try {if ($script:WorkerPowerShell) {
        $script:WorkerPowerShell.Dispose()}} catch {}
        $script:WorkerPowerShell = $null
        $script:WorkerAsync = $null
        $script:WorkerControl = $null
        $btnStart.Enabled = $true
        $btnScan.Enabled = $true
        $btnBrowse.Enabled = $true
        $btnDeleteBaks.Enabled = $true
        $chkAdminMode.Enabled = $true
        $btnCancel.Enabled = $false
        $rbHigh.Enabled = $true
        $rbMedium.Enabled = $true
        $rbLow.Enabled = $true
        Read-QueueIntoUi}})
$form.Add_Shown({Add-UILog 'GUI initialised successfully.'
    Add-UILog "PowerShell version: $($PSVersionTable.PSVersion)"
    Add-UILog "ImageMagick path: $ImageMagick"
    Add-UILog "ImageMagick exists: $(Test-Path -LiteralPath $ImageMagick -PathType Leaf)"
    Add-UILog 'Ready. Select a folder and click Scan for PPTX Files.'})
$form.Add_FormClosing({
    if ($script:RunState -eq 'Running') {
        $answer = [System.Windows.Forms.MessageBox]::Show('Compression is still running. You might corrupt your files if you do this. You have been warned.',
            'ALERT ALERT ALERT!!!',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -eq [System.Windows.Forms.DialogResult]::No) {
        $_.Cancel = $true 
        return }
        if ($script:WorkerControl) {$script:WorkerControl.Cancelled = $true}
        if ($script:WorkerPowerShell) {try { $script:WorkerPowerShell.Stop() } catch {}
        try { $script:WorkerPowerShell.Dispose() } catch {}}}
    try { $timer.Stop() } catch {}
    try { if ($script:LogWriter) { $script:LogWriter.Dispose() } } catch {}})
[void]$form.ShowDialog()
try {if ($script:LogWriter) {$script:LogWriter.Dispose()}} catch {}
