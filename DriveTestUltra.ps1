param (
    [string]$batchId=(Get-Date -format "yyyy-MM-dd_hh-mm-ss"), # 'u' and 's' will have colons, which is bad for filenames
    [string]$testSize='1M',
    [int]$durationSec=3, # changed from 5 seconds - 3 works fine on modern hardware.
    [int]$warmupSec=0,
    [int]$cooldownSec=0,
    #[int]$restSec=5,
    [string]$diskspd= $PSScriptRoot + ".\diskspd.exe"
    # Used to be: 'C:\Users\holter\Downloads\DiskSpdAuto-master\DiskSpdAuto-master\DiskSpd\amd64\diskspd.exe'
)

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..."
    Install-Module -Name ImportExcel -Force -ErrorAction Stop
} else {
    Write-Host "ImportExcel module is already installed. Skipping the installation process"
}

# Check if NuGet package provider is installed, and install if not
if (-not (Get-PackageProvider -ListAvailable -Name NuGet -Force)) {
    Write-Host "NuGet package provider not found. Installing..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
} else {
    Write-Host "NuGet package provider is already installed. Skipping the installation process"
}

# Check if the module is already installed
<#if (-not (Get-Module -Name PSWriteOffice -ListAvailable -ErrorAction SilentlyContinue)) {
    # Module not installed, proceed with the installation
    Install-Module -Name PSWriteOffice -Force -Scope CurrentUser -AllowClobber
    Write-Host "Module 'PSWriteOffice' has been installed."
} else {
    Write-Host "Module 'PSWriteOffice' is already installed. Skipping the installation process."
}
#>

# set folder paths for charts
$mainfolder = ([System.Environment]::GetFolderPath('MyDocuments'))
$chartsFolder = Join-Path -Path $mainfolder -ChildPath "DTUCharts"


if (-not (Get-Module -Name PSWriteWord -ListAvailable -ErrorAction SilentlyContinue)) {
    # Module not installed, proceed with the installation
    Install-Module -Name PSWriteWord -Force -Scope CurrentUser -AllowClobber

    #replace file inside module forder with custom one added in (if available)
    $pswritewordModule = Join-Path -Path $mainfolder -ChildPath "WindowsPowerShell\Modules\PSWriteWord\1.1.14\PSWriteWord.psm1"
    $customModule = Join-Path -Path $PSScriptRoot ".\PSWriteWord.psm1"
    if ([System.IO.File]::Exists($pswritewordModule) -and [System.IO.File]::Exists($customModule))
    {
        Set-Content -Path $pswritewordModule -Value $customModule
    }

    #proper install
    Write-Host "Module 'PSWriteWord' has been installed."
} else {
    Write-Host "Module 'PSWriteWord' is already installed. Skipping the installation process."
}



#all examples of this program added this in, so to make sure the program works, we will have it in here too
Import-Module PSWriteWord #-Force

# validity of drive letter
$validDrive = $false

while (-not $validDrive) {
    $drive = Read-Host -Prompt "Drive Letter"

    #make sure drive is not C drive since I think people want to keep system32
    if ($drive -match '^[A-Za-z]$' -and $drive -ne 'C') {
        $validDrive = $true
    } else {
        Write-Host "Invalid drive letter. Please enter a valid drive letter other than 'C'."
    }
}

# validity of drive name
$validname = $false

while (-not $validname)
{
    $newname = Read-Host -Prompt "Barcode"

    if ($newname.Length -le 11) # 11 is max FAT character limit, will *maybe* add NTFS support later not sure
    {
        $validname = $true
    }
    else {
        Write-Host "Invalid name of drive. Please enter a drive name less than or equal to 11 characters."
    }
}

# Spire allows the program to generate the graphs as .pngs
Add-Type -Path (Join-Path $PSScriptRoot ".\Spire.XLS.dll")




# Check if the folder exists, and create it if not
if (-not (Test-Path $chartsFolder -PathType Container)) {
    New-Item -Path $chartsFolder -ItemType Directory
}

#install excel spreadsheet tool (cannot be run without admin permissions however, when compiled with ps2exe, admin perms will be required)

#ImportExcel can be found here https://github.com/dfinke/ImportExcel
# Check if ImportExcel module is installed, and install if not



#get new name of the disk (should be ctrl v from barcode scanner)
$BDStandardName = "Default"

# renames drive to new name now because, if left until after tests, the new name will not appear on csvs, spreadsheets, and images
Set-Volume -DriveLetter $drive -NewFileSystemLabel $newname


# get test summary object
# assume one target and one timespan
function OneTargetRead {
    param ( $test, $xmlFilePath, $driveObj )
    $x = [xml](Get-Content $xmlFilePath)
    $o = New-Object psobject
    # test meta data
    Add-Member -InputObject $o -MemberType noteproperty -Name 'ComputerName' -Value $x.Results.System.ComputerName
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Drive' -Value $drive
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Drive VolumeLabel' -Value $newname
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Batch' -Value $batchId
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Test Time' -Value (Get-Date)
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Test Name' -Value $test.name
    # io meta data
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Test File Size' -Value $testSize
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Duration [s]' -Value $durationSec
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Warmup [s]' -Value $warmupSec
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Cooldown [s]' -Value $cooldownSec
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Test Params' -Value $test.params
    # io metrics
    Add-Member -InputObject $o -MemberType noteproperty -Name 'TestTimeSeconds' -Value $x.Results.TimeSpan.TestTimeSeconds
    Add-Member -InputObject $o -MemberType noteproperty -Name 'WriteRatio' -Value ($x.Results.Profile.TimeSpans.TimeSpan.Targets.Target.WriteRatio | Select-Object -first 1)
    Add-Member -InputObject $o -MemberType noteproperty -Name 'ThreadCount' -Value $x.Results.TimeSpan.ThreadCount
    Add-Member -InputObject $o -MemberType noteproperty -Name 'RequestCount' -Value ($x.Results.Profile.TimeSpans.TimeSpan.Targets.Target.RequestCount | Select-Object -first 1)
    Add-Member -InputObject $o -MemberType noteproperty -Name 'BlockSize' -Value ($x.Results.Profile.TimeSpans.TimeSpan.Targets.Target.BlockSize | Select-Object -first 1)

    # sum read and write iops across all threads and targets
    $ri = ($x.Results.TimeSpan.Thread.Target |
            Measure-Object -sum -Property ReadCount).Sum
    $wi = ($x.Results.TimeSpan.Thread.Target |
            Measure-Object -sum -Property WriteCount).Sum
    $rb = ($x.Results.TimeSpan.Thread.Target |
            Measure-Object -sum -Property ReadBytes).Sum
    $wb = ($x.Results.TimeSpan.Thread.Target |
            Measure-Object -sum -Property WriteBytes).Sum
    Add-Member -InputObject $o -MemberType noteproperty -Name 'ReadCount' -Value $ri
    Add-Member -InputObject $o -MemberType noteproperty -Name 'WriteCount' -Value $wi
    Add-Member -InputObject $o -MemberType noteproperty -Name 'ReadBytes' -Value $rb
    Add-Member -InputObject $o -MemberType noteproperty -Name 'WriteBytes' -Value $wb

    # latency
    $l = @(); foreach ($i in 25,50,75,90,95,99,99.9,100) { $l += ,[string]$i }
    $h = @{}; $x.Results.TimeSpan.Latency.Bucket |ForEach-Object { $h[$_.Percentile] = $_ } # AY, hash all percentiles in $h
    $l |ForEach-Object {
        $b = $h[$_];
        Add-Member -InputObject $o -MemberType noteproperty -Name ('{0}% r' -f $_) -Value $b.ReadMilliseconds
        Add-Member -InputObject $o -MemberType noteproperty -Name ('{0}% w' -f $_) -Value $b.WriteMilliseconds
    }

    return $o
}

function TotalTargetRead {
    param ( $tests )

    #create document for .csv
    $o = New-Object psobject

    # drive meta data
    Add-Member -InputObject $o -MemberType noteproperty -Name 'ComputerName' -Value $tests[0].ComputerName
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Drive' -Value $tests[0].Drive
    Add-Member -InputObject $o -MemberType noteproperty -Name 'DriveVolumeLabel' -Value $tests[0].'Drive VolumeLabel'
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Batch' -Value $tests[0].Batch
    Add-Member -InputObject $o -MemberType noteproperty -Name 'TestTime' -Value $tests[0].'Test Time'
    # io meta data
    Add-Member -InputObject $o -MemberType noteproperty -Name 'TestFileSize' -Value $tests[0].'Test File Size'
    Add-Member -InputObject $o -MemberType noteproperty -Name 'TestDuration' -Value $tests[0].'Duration [s]'

    # io
    $t_sr=$tests |Where-Object {$_.'Test Name' -eq 'Sequential read'}
    $v=$t_sr.ReadBytes/$t_sr.TestTimeSeconds/1024/1024
    Add-Member -InputObject $o -MemberType noteproperty -Name 'SequentialRead' -Value $v

    $t_sw=$tests |Where-Object {$_.'Test Name' -eq 'Sequential write'}
    $v=$t_sw.WriteBytes/$t_sw.TestTimeSeconds/1024/1024
    Add-Member -InputObject $o -MemberType noteproperty -Name 'SequentialWrite' -Value $v

    # $t_rr=$tests |Where-Object {$_.'Test Name' -eq 'Random read'}
    # $v=$t_rr.ReadBytes/$t_rr.TestTimeSeconds/1024/1024
    #  Add-Member -InputObject $o -MemberType noteproperty -Name 'Random Read 4KB (QD=1) [MB/s]' -Value $v

    #  $t_rw=$tests |Where-Object {$_.'Test Name' -eq 'Random write'}
    # $v=$t_rw.WriteBytes/$t_rw.TestTimeSeconds/1024/1024
    # Add-Member -InputObject $o -MemberType noteproperty -Name 'Random Write 4KB (QD=1) [MB/s]' -Value $v

    # $t_r2r=$tests |Where-Object {$_.'Test Name' -eq 'Random QD32 read'}
    # $v=$t_r2r.ReadBytes/$t_r2r.TestTimeSeconds/1024/1024
    # Add-Member -InputObject $o -MemberType noteproperty -Name 'Random Read 4KB (QD=32) [MB/s]' -Value $v

    # $t_r2w=$tests |Where-Object {$_.'Test Name' -eq 'Random QD32 write'}
    #  $v=$t_r2w.WriteBytes/$t_r2w.TestTimeSeconds/1024/1024
    # Add-Member -InputObject $o -MemberType noteproperty -Name 'Random Write 4KB (QD=32) [MB/s]' -Value $v

    return $o
}

#.tmp file used by diskspd.exe in order to run tests
$benchmarkContent = Get-Content -Raw ($PSScriptRoot + "\benchmark.tmp")
# initialize test file
$testFileParams = '{0}:\benchmark.tmp' -f $drive

# Check if the file already exists, and if not, create it
if ((-not (Test-Path $testFileParams)) -and ($null -ne $benchmarkContent)) {
    New-Item -Path $testFileParams -Value $benchmarkContent -Force 
} else {
    # cannot use new item 
    $params = ('{0} -d1 -S -Z1M -c{1}' -f $testFileParams, $testSize)
    & $diskspd ($params -split ' ') > $xmlFile
}

# $testFileParams='{0}:\benchmark.tmp' -f $drive --> already declared above
$xmlFile=('{0}-Generation.xml' -f $batchId);
$params=( ('-Rxml -d1 -S -Z1M -c{0}' -f $testSize) ,$testFileParams) -join ' '; # make sure to write with cache disabled, or else on slow systems this will exit with data still writing from cache to disk.
# Write-Host $params
# Write-Host $xmlFile
& $diskspd ($params -split ' ') > $xmlFile

# fixed params for tests
$fixedParams='-L -S -Rxml'

# batch auto params
$batchAutoParam='-d{0} -W{1} -C{2}' -f $durationSec, $warmupSec, $cooldownSec


# iterate over tests
$tests=@()
foreach ($test in @{name='Sequential read'; params='-b1M -o1 -t1 -w0 -Z1M'},
    @{name='Sequential write'; params='-b1M -o1 -t1 -w100 -Z1M'}<#,
    @{name='Random read'; params='-b4K -o1 -t1 -r -w0 -Z1M'},
    @{name='Random write'; params='-b4K -o1 -t1 -r -w100 -Z1M'},
    @{name='Random QD32 read'; params='-b4K -o32 -t1 -r -w0 -Z1M'},
    @{name='Random QD32 write'; params='-b4K -o32 -t1 -r -w100 -Z1M'},
    @{name='Random T32 read'; params='-b4k -o1 -t32 -r -w0 -Z1M'},
    @{name='Random T32 write'; params='-b4k -o1 -t32 -r -w100 -Z1M'}#>) {
        # run test
        $params=($fixedParams,$batchAutoParam,$test.params,$testFileParams) -join ' ';
        # set xml file
        $xmlFile=('{0}-{1}.xml' -f $batchId, $test.name);

        # Write-Host $params

        # Write-Host $xmlFile

        #highly doubt this is necessary with modern hardware
        #Start-Sleep $restSec # sleep a sec to calm down IO
        
        # run diskspd tests and output to xml
        & $diskspd ($params -split ' ') > $xmlFile

        # read result and write to batch file
        $driveObj=[System.IO.DriveInfo]::GetDrives() | Where-Object {$_.Name -eq $drive }
        $testResult = OneTargetRead $test $xmlFile $driveObj
         # Store the values in new variables
    <#if ($test.name -eq 'Sequential read') {
        $q_sr = $testResult.SequentialRead
    } elseif ($test.name -eq 'Sequential write') {
        $q_sw = $testResult.SequentialWrite
    } #>

    <#
    # Output the values to the terminal for debugging
Write-Host "Sequential Read Values: $q_sr"
Write-Host "Sequential Write Values: $q_sw"
#>

#export to csv
        $testResult | Export-Csv ('{0}.csv' -f $batchId) -NoTypeInformation -Append
        $tests+=$testResult
}

   

# sum drive tests to a single row
$testsSum = TotalTargetRead $tests

# $testsSum 
$date = Get-Date -Format "yyyy-MM-dd"
$csvoutputPath = Join-Path -Path $mainfolder -ChildPath ('BD-{0}.csv' -f $date)
$csv2outputPath = Join-Path -Path $mainfolder -ChildPath ('BD-{0}.csv' -f $newname)
$exceloutputPath = Join-Path -Path $mainfolder -ChildPath ('BD.xlsx')


#set chart definition
$chart = New-ExcelChartDefinition -Title $newname -YMaxValue 100 -XAxisTitleText "Read & Write" -YAxisTitleText "Transfer Rate [MB/s]" -YRange 'SequentialRead','SequentialWrite' -ChartType "BarClustered" -LegendBold -SeriesHeader "Read", "Write" 

$testsSum | Export-Csv -Path $csvoutputPath -NoTypeInformation -Append -Force
$testsSum | Export-Csv -Path $csv2outputPath -NoTypeInformation -Force

# more of matts fucky wucky code
# Get the latest .csv file
# Import the latest .csv file
$latestData = Import-Csv $csv2outputPath

# Extract the latest values for SequentialRead and SequentialWrite
$y_sr = @($latestData.SequentialRead)
$y_sw = @($latestData.SequentialWrite)


# Output the latest values for SequentialRead and SequentialWrite
Write-Host "Latest Sequential Read Value: $y_sr"
Write-Host "Latest Sequential Write Value: $y_sw"


# Export data to an excel graph
Import-Csv -Path $csv2outputPath | Export-Excel $exceloutputpath -AutoNameRange -ExcelChartDefinition $chart -WorkSheetname $BDStandardName -ReturnRange
##### format drive afterwards ######
Format-Volume -DriveLetter $drive -NewFileSystemLabel $newname

## Matts fucky wucky code

# Add-WordBarChart needs to receive the values as individual elements, not arrays
# Convert the arrays to individual elements if they are arrays

# y values for charts
$y_sr_values = $y_sr -join ','
$y_sw_values = $y_sw -join ','


# Use spire to export the images
$workbook = New-Object Spire.Xls.Workbook
$workbook.LoadFromFile($exceloutputPath)
<# THIS IS FOR THE FUTURE WHERE WE EXPORT FROM THE MAIN EXCEL SPREADSHEET

for ($p = 0; $p -lt $workbook.Worksheets.Length; $p++)
{
    $sheet = $workbook.Worksheets[$p]
    $imgs = $workbook.SaveChartAsImage($sheet)


    # Save the charts to png files

    for ($i = 0; $i -lt $imgs.Length; $i++) {
        $outputchartsPath = Join-Path -Path $chartsFolder -ChildPath ('{0}-Chart.png' -f $newname)
        $fileStream = New-Object System.IO.FileStream($outputchartsPath, [System.IO.FileMode]::Create)
        $imgs[$i].Save($fileStream, [System.Drawing.Imaging.ImageFormat]::Png)
        $fileStream.Close()
    }
}   




#>



$sheet = $workbook.Worksheets[0] #gets the first sheet
$imgs = $workbook.SaveChartAsImage($sheet) # exports as an array


# Save the charts to png files
for ($i = 0; $i -lt $imgs.Length; $i++) {
    $outputchartsPath = Join-Path -Path $chartsFolder -ChildPath ('{0}-Chart.png' -f $newname) #need to create new path every time
    $fileStream = New-Object System.IO.FileStream($outputchartsPath, [System.IO.FileMode]::Create) #open new filestream
    $imgs[$i].Save($fileStream, [System.Drawing.Imaging.ImageFormat]::Png) #export as png
    $fileStream.Close() #can't keep open bc memory leaks
}

$ChartImgDoc = Join-Path -Path $mainfolder -ChildPath ('ChartDoc-{0}.docx' -f $date) #chart doc yippee!

if (-not (Test-Path $ChartImgDoc)) {
    $WordDocument2 = New-WordDocument $ChartImgDoc # new doc
    $PlaceToAddPicture = Add-WordText -WordDocument $WordDocument2 -Text 'Adding a picture...' -Supress $false #blahblahblah 
    Add-WordPicture -WordDocument $WordDocument2 -ImagePath $outputchartsPath -ImageWidth 200 -ImageHeight 150 -Alignment left -Verbose #adds picture from DTUCharts folder
} else {
    $WordDocument2 = Get-WordDocument $ChartImgDoc #import doc info
    $PlaceToAddPicture = Add-WordText -WordDocument $WordDocument2 -Text 'Adding a picture...' -Supress $false #blahblahblah
    #adds picture from DTUCharts folder
    Add-WordPicture -WordDocument $WordDocument2 -ImagePath $outputchartsPath -ImageWidth 200 -ImageHeight 150 -Alignment both -Verbose -Paragraph $p.AppendPicture
    # tell user
    Write-Host "The file already exists, writing to file."
}

#make sure word doc is saved
Save-WordDocument $WordDocument2 -Language 'en-US' -Supress $true -OpenDocument
<#


#TESTING ONLY
$wordDocLocation = Join-Path -Path $chartsFolder -ChildPath ('Charts.docx')


$WordDocument = New-WordDocument $wordDocLocation
Add-WordText -WordDocument $WordDocument -Text 'Bar Chart test #1' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True
#    Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances' -Names 'Today', 'Yesterday' -Values 1050.50, 2000 -ChartLegendPosition Bottom -ChartLegendOverlay $false
Add-WordBarChart -WordDocument $WordDocument -ChartX 200 -ChartY 200 -ChartName "$newname" -Names "Sequential Read: $y_sr_values MB/s", "Sequential Write: $y_sw_values MB/s" -Values $y_sr_values, $y_sw_values -ChartLegendPosition Bottom -ChartLegendOverlay $false 

Save-WordDocument $WordDocument -Supress $True

### Start Word with file
#Invoke-Item $wordDocLocation
#>


## REMOVE TEMP 
Remove-Item -Path $csv2outputPath


#"Press enter to continue..." :3
pause