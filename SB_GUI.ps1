[System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsBase') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.ComponentModel') | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Data')           | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')        | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null

[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')      | out-null
[System.Reflection.Assembly]::LoadFrom($PSScriptRoot + "\assembly\MahApps.Metro.dll")       | out-null
[System.Reflection.Assembly]::LoadFrom($PSScriptRoot + "\assembly\System.Windows.Interactivity.dll") | out-null

Add-Type -AssemblyName "System.Windows.Forms"
Add-Type -AssemblyName "System.Drawing"




<#
#Live-compile the Switchboard.ps1 to get around stupid windows bullshit saying it doesnt like to run files not created on the device due to execution policy
$sourceFilePath = "SB.txt"
$destinationFilePath = "Switchboard.ps1"

# Check if Switchboard.ps1 already exists
if (-not (Test-Path $destinationFilePath)) {
    # Check if the source file exists
    if (Test-Path $sourceFilePath) {
        # Read the contents of SB.txt
        $fileContent = Get-Content -Path $sourceFilePath
        
        # Write the contents to Switchboard.ps1
        Set-Content -Path $destinationFilePath -Value $fileContent
        
        Write-Output "Switchboard.ps1 has been created and filled with the contents of SB.txt."
    } else {
        Write-Error "The file SB.txt does not exist. Please check the file path and try again."
    }
}
#>
<#
if (-not (Get-PackageProvider -ListAvailable -Name NuGet -Force)) {
    Write-Host "NuGet package provider not found. Installing..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
} else {
    Write-Host "NuGet package provider is already installed. Skipping the installation process"
}
#>

$PSWriteOfficePath = Join-Path -Path $PSScriptRoot ".\PSWriteOffice\0.2.0\PSWriteOffice.psd1"
$ImportExcelPath = Join-Path -Path $PSScriptRoot ".\ImportExcel\ImportExcel.psd1"

Import-Module $PSWriteOfficePath -Force
Import-Module $ImportExcelPath -Force


# Load a xml file
function LoadXml ($filename)
{
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

# Load MainWindow
$XamlMainWindow=LoadXml($PSScriptRoot + ".\mahapps.xaml")
$Reader=(New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form=[Windows.Markup.XamlReader]::Load($Reader)

#$CompName = $env:COMPUTERNAME



####################################
##### Buttons initialization
####################################
#$Computer_name = $Form.Findname("Computer_name")

$Run = $Form.Findname("Run")

$Image_Handler = $Form.Findname("Image_Control")

##### Buttons and TextBox




##### Metro Circle Button in the Flyout content
#$Open_Installed_Soft = $Form.Findname("Open_Installed_Soft")


####################################
##### Buttons Actions
####################################

# $Computer_name.Text = $CompName

# Function to invoke the external PowerShell script and capture its output
# Function to invoke the external PowerShell script asynchronously
# Function to invoke the external PowerShell script asynchronously

# Function to prompt user for input
function PromptUserInput {
    param (
        [string]$PromptText
    )
    
    $InputForm = New-Object System.Windows.Forms.Form
    $InputForm.Text = $PromptText
    $InputForm.Size = New-Object System.Drawing.Size(300,200)
    $InputForm.StartPosition = "CenterScreen"

    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Point(10,20)
    $Label.Size = New-Object System.Drawing.Size(280,40)
    $Label.Text = $PromptText
    $InputForm.Controls.Add($Label)

    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point(10,60)
    $TextBox.Size = New-Object System.Drawing.Size(260,20)
    $InputForm.Controls.Add($TextBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(180,110)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $InputForm.Controls.Add($OKButton)

    $InputForm.AcceptButton = $OKButton

    $Result = $InputForm.ShowDialog()

    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $TextBox.Text
    } else {
        return $null
    }
}


function PromptValidDriveInput {
    param (
        [string]$PromptText
    )
    
    $InputForm = New-Object System.Windows.Forms.Form
    $InputForm.Text = $PromptText
    $InputForm.Size = New-Object System.Drawing.Size(300,200)
    $InputForm.StartPosition = "CenterScreen"

    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Point(10,20)
    $Label.Size = New-Object System.Drawing.Size(280,50)
    $Label.Text = $PromptText
    $InputForm.Controls.Add($Label)

    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point(10,70)
    $TextBox.Size = New-Object System.Drawing.Size(260,20)
    $InputForm.Controls.Add($TextBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(180,110)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $InputForm.Controls.Add($OKButton)

    $InputForm.AcceptButton = $OKButton

    $Result = $InputForm.ShowDialog()

    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $TextBox.Text
    } else {
        return $null
    }
}
function Get-Chart {
    $documentsFolder = [System.Environment]::GetFolderPath('MyDocuments')
    $switchboardChartsFolder = Join-Path -Path $documentsFolder -ChildPath 'SwitchboardCharts'

    # Check if the folder exists, and create it if not
#if (-not (Test-Path $switchboardChartsFolder -PathType Container)) {
#    New-Item -Path $switchboardChartsFolder -ItemType Directory
#}

    # Get all image files matching the pattern $newname-Chart.png
    $imageFiles = Get-ChildItem -Path $switchboardChartsFolder -Filter "$newname-Chart.png" -File

    if ($imageFiles) {
        # Sort the image files by LastWriteTime in descending order to get the latest one
        $latestImageFile = $imageFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        return $latestImageFile.FullName
    } else {
        Write-Host "No image file found for $newname-Chart.png"
        return $null
    }
}

$Run.Add_Click({
    $Run.IsEnabled = $false

    $Image_Handler.Source = $null
    

    $path = $null
    
    function IsDriveLetterValid($driveLetter) {
        return ($driveLetter.Length -eq 1 -and $driveLetter -match '^[a-zA-Z]$' -and $driveLetter -ne 'C')
    }

    function IsNameValid($name) {
        return ($name.Length -ge 3 -and $name.Length -le 11 -and $name -match '^[a-zA-Z0-9]+$')
    }

    $newname = PromptUserInput -PromptText "Please enter the new name:"
    if ($null -eq $newname) {
        Write-Host "New name not provided. Operation canceled."
        $Run.IsEnabled = $true
        return
    }

    while (-not (IsNameValid $newname)) {
        Write-Host "Invalid name provided. Please enter a name that is between 3 and 11 characters long and contains only letters and numbers."
        $newname = PromptUserInput -PromptText "Please enter a valid name (3-11 characters, letters and numbers only):"
        if ($null -eq $newname) {
            Write-Host "New name not provided. Operation canceled."
            $Run.IsEnabled = $true
            return
        }
    }

    <#
    $drive = ((Get-CimInstance -Class Win32_DiskDrive -Filter 'InterfaceType = "USB"' -KeyOnly | Get-CimAssociatedInstance -ResultClassName Win32_DiskPartition -KeyOnly | Get-CimAssociatedInstance -ResultClassName Win32_LogicalDisk).DeviceID -replace ':', '')


    #$drive = PromptValidDriveInput -PromptText "Please enter the drive letter:"    
    if ($null -eq $drive) {
        Write-Host "Drive letter not provided. Operation canceled."
        $Run.IsEnabled = $true
        return
    }

    while (-not (IsDriveLetterValid $drive)) {
        Write-Host "Invalid drive letter provided. Please enter a single letter that is not 'C'."
        $drive = PromptValidDriveInput -PromptText "Please enter a valid drive letter (a single letter that is not 'C'):"
        if ($null -eq $drive) {
            Write-Host "Drive letter not provided. Operation canceled."
            $Run.IsEnabled = $true
            return
        }
    }
#>


# Retry loop to get SD card drive letter

 $drive = $null
 $mrFreeze = gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"} | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}
$drive = $mrFreeze -replace ':', ''

 if ($null -or '' -eq $drive) {
    Write-Host "No drive letter provided. Please enter a single letter that is not 'C'."
    $drive = PromptValidDriveInput -PromptText "Automatic drive detection failed. Please insert your drive, then enter the corresponding drive letter:"
    if ($null -or ''-eq $drive) {
        Write-Host "Drive letter not provided. Operation canceled."
        $Run.IsEnabled = $true
        return
 }

 while (-not (IsDriveLetterValid $drive)) {
     Write-Host "Invalid drive letter provided. Please enter a single letter that is not 'C'."
     $drive = PromptValidDriveInput -PromptText "Please enter a valid drive letter (a single letter that is not 'C'):"
     if ($null -eq $drive) {
         Write-Host "Drive letter not provided. Operation canceled."
         $Run.IsEnabled = $true
         return
     }
 }
 }
  

    # Create "Job running" form
    $jobRunningForm = New-Object System.Windows.Forms.Form
    $jobRunningForm.StartPosition = 'CenterScreen'
    $jobRunningForm.Width = 400
    $jobRunningForm.Height = 150
    $jobRunningForm.Text = ""
    $jobRunningForm.ControlBox = $false

    $jobRunningLabel = New-Object System.Windows.Forms.Label
    $jobRunningLabel.AutoSize = $true
    $jobRunningLabel.Text = "Benchmark running, please wait..."
    $jobRunningLabel.TextAlign = 'MiddleCenter'
    $jobRunningLabel.Dock = 'Fill'
    $jobRunningLabel.Font = New-Object System.Drawing.Font("Arial", 14)
    $jobRunningForm.Controls.Add($jobRunningLabel)

    $null = $jobRunningForm.Show()

    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Switchboard.ps1"
    $job = Start-Job -ScriptBlock {
        param($newname, $drive)
        & $using:scriptPath -newname $newname -drive $drive
    } -ArgumentList $newname, $drive

    Write-Host "DriveTestUltra job started with -newname $newname -drive $drive"

    Wait-Job -Job $job

    # Retrieve and display the job results
    $jobResult = Receive-Job -Job $job
    Write-Host "Benchmark complete. Results:"
    Write-Host $jobResult
    
    # Clean up the job
    Remove-Job -Job $job

    # Close the "Job running" pop-up window
    if ($jobRunningForm) {
        $jobRunningForm.Close()
    }

    $TempFileLocation = $PSScriptRoot + ".\temp.txt"
    $PassOrFail = Get-Content -Path $TempFileLocation

    $fail = -not ($PassOrFail -eq "Pass")

    # Retrieve and display the job results
    # $jobResult = Receive-Job -Job $job
    # Write-Host "Benchmark complete. Results:"
    # Write-Host $jobResult

    #UPDATE IMAGE
    $path = Get-Chart
    $Image_Handler.Source = [System.Windows.Media.Imaging.BitmapImage]::new()
    $bi = $Image_Handler.Source
    # Ensure that $path is not null and is a valid URI before setting the Image Source
if ($null -ne $path -and [System.IO.File]::Exists($path)) {
    try {
        $bi.BeginInit()
        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bi.CreateOptions = [System.Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache
        $bi.UriSource = [uri]::new($path, [urikind]::RelativeOrAbsolute)
        $bi.EndInit()
        
        
        
        #$Image_Handler.Source = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$path, "None", "OnLoad")
    } catch {
        Write-Host "Error setting image source: $_"
    }
} else {
    Write-Host "Image path is null or file does not exist: $path"
}


    $completionText = ""
    if ($fail){ $completionText = "Benchmark Failed!`r`nCheck log file for more details."}
    else {$completionText = "Benchmark Passed!"}
    # Show the "Job complete" message and auto-close after 3 seconds
    $completionForm = New-Object System.Windows.Forms.Form
    $completionForm.StartPosition = 'CenterScreen'
    $completionForm.Width = 400
    $completionForm.Height = 150
    $completionForm.Text = ""
    $completionForm.ControlBox = $false

    $completionLabel = New-Object System.Windows.Forms.Label
    $completionLabel.AutoSize = $true
    $completionLabel.Text = $completionText
    $completionLabel.TextAlign = 'MiddleCenter'
    $completionLabel.Dock = 'Fill'
    $completionLabel.Font = New-Object System.Drawing.Font("Arial", 14)
    $completionForm.Controls.Add($completionLabel)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 3000

    $timer.Add_Tick({
        #close form once status changes
        $completionForm.Close()
    })

    $timer.Start()

    $completionForm.ShowDialog()
    
    $timer.Dispose()

    $completionForm.Dispose()
    #Re enable run
    $Run.IsEnabled = $true
    


    
  

    
})


		

$Form.MaxHeight = 250
$Form.MaxWidth = 575
$Form.MinHeight = 250
$Form.MinWidth = 575
$Form.ShowDialog() | Out-Null