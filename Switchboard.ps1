param (
    [string]$batchId=(Get-Date -format "yyyy-MM-dd_hh-mm-ss"), # 'u' and 's' will have colons, which is bad for filenames
    [string]$testSize='8M',
    $newname,
    $drive,
    [int]$durationSec=3, # changed from 5 seconds - 3 works fine on modern hardware.
    [int]$warmupSec=0,
    [int]$cooldownSec=0,
    #[int]$restSec=5,
    [string]$diskspd= $PSScriptRoot + ".\diskspd.exe"
    # Used to be: 'C:\Users\holter\Downloads\DiskSpdAuto-master\DiskSpdAuto-master\DiskSpd\amd64\diskspd.exe'
)

$PSWriteOfficePath = Join-Path -Path $PSScriptRoot ".\PSWriteOffice\0.2.0\PSWriteOffice.psd1"
$ImportExcelPath = Join-Path -Path $PSScriptRoot ".\ImportExcel\ImportExcel.psd1"

Import-Module $PSWriteOfficePath -Force
Import-Module $ImportExcelPath -Force


#Write-host $drive
#Write-host $newname

# set folder paths for charts
$mainfolder = ([System.Environment]::GetFolderPath('MyDocuments'))
$chartsFolder = Join-Path -Path $mainfolder -ChildPath "SwitchboardCharts"

if (-not (Test-Path $chartsFolder -PathType Container)) {
    New-Item -Path $chartsFolder -ItemType Directory
}

# Spire allows the program to generate the graphs as .pngs
Add-Type -Path (Join-Path $PSScriptRoot ".\Spire.XLS.dll")




# Check if the folder exists, and create it if not
#if (-not (Test-Path $chartsFolder -PathType Container)) {
#    New-Item -Path $chartsFolder -ItemType Directory
#}

#install excel spreadsheet tool (cannot be run without admin permissions however, when compiled with ps2exe, admin perms will be required)

#ImportExcel can be found here https://github.com/dfinke/ImportExcel
# Check if ImportExcel module is installed, and install if not



#get new name of the disk (should be ctrl v from barcode scanner)
$BDStandardName = "Default"

# renames drive to new name now because, if left until after tests, the new name will not appear on csvs, spreadsheets, and images
Set-Volume -DriveLetter $drive.ToUpper() -NewFileSystemLabel $newname


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
    #Add-Member -InputObject $o -MemberType noteproperty -Name 'Batch' -Value $batchId
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
    # Matt Metrics
    Add-Member -InputObject $o -MemberType noteproperty -Name 'PassOrFail' -Value $PassOrFail 
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





#Zach please see if I have to define this here or if i can define it down below with the rest of testsum plz and ty habibi
$csv3outputPath = Join-Path -Path $mainfolder -ChildPath ('{0}.csv' -f $batchId)





function TotalTargetRead {
    param ( $tests )

    #create document for .csv
    $o = New-Object psobject

    # drive meta data
    Add-Member -InputObject $o -MemberType noteproperty -Name 'ComputerName' -Value $tests[0].ComputerName
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Drive' -Value $tests[0].Drive
    Add-Member -InputObject $o -MemberType noteproperty -Name 'DriveVolumeLabel' -Value $tests[0].'Drive VolumeLabel'
    #Add-Member -InputObject $o -MemberType noteproperty -Name 'Batch' -Value $tests[0].Batch
    Add-Member -InputObject $o -MemberType noteproperty -Name 'TestTime' -Value $tests[0].'Test Time'
    # io meta data
    Add-Member -InputObject $o -MemberType noteproperty -Name 'TestFileSize' -Value $tests[0].'Test File Size'
    Add-Member -InputObject $o -MemberType noteproperty -Name 'TestDuration' -Value $tests[0].'Duration [s]'
    

    # io
   <# $t_sr=$tests |Where-Object {$_.'Test Name' -eq 'Sequential read'}
    $v=$t_sr.ReadBytes/$t_sr.TestTimeSeconds/1024/1024
    $v2=$t_sr.ReadBytes/$t_sr.TestTimeSeconds/1024/1024
    #Note: This used to say "Sequential Read", but it was changed to "ReadValue" for GISDCARDTOOL. For public Switchboard release, the word "Sequential" should be reinstated
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Read' -Value $v

#>

# io
$t_sr = $tests | Where-Object { $_.'Test Name' -eq 'Random QD4 read' }
#$v = $t_sr.ReadBytes / $t_sr.TestTimeSeconds / 1024 / 1024
# Note: This used to say "Sequential Read", but it was changed to "ReadValue" for GISDCARDTOOL. For public Switchboard release, the word "Sequential" should be reinstated
$v2 = $t_sr.ReadBytes / $t_sr.TestTimeSeconds / 1024 / 1024
$vRounded = [math]::Round($v2)
Add-Member -InputObject $o -MemberType noteproperty -Name 'Read' -Value $vRounded

<#
    $t_sw=$tests |Where-Object {$_.'Test Name' -eq 'Sequential write'}
    $v=$t_sw.WriteBytes/$t_sw.TestTimeSeconds/1024/1024
    $v3=$t_sw.WriteBytes/$t_sw.TestTimeSeconds/1024/1024
    #Note: This used to say "Sequential Write", but it was changed to "WriteValue" for GISDCARDTOOL. For public Switchboard release, the word "Sequential" should be reinstated
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Write' -Value $v
#>
$t_sw = $tests | Where-Object { $_.'Test Name' -eq 'Random QD4 write' }
#$v = $t_sw.WriteBytes / $t_sw.TestTimeSeconds / 1024 / 1024
# Note: This used to say "Sequential Read", but it was changed to "ReadValue" for GISDCARDTOOL. For public Switchboard release, the word "Sequential" should be reinstated
$v3 = $t_sw.WriteBytes / $t_sw.TestTimeSeconds / 1024 / 1024
$vRounded2 = [math]::Round($v3)
Add-Member -InputObject $o -MemberType noteproperty -Name 'Write' -Value $vRounded2


    $PassOrFail = "Null"
    $Reason = ""
    if ($v2 -lt 70.00000 -and $v2 -gt 10.00000) {
        $PassOrFail = "Fail"
        $Reason = "Slow Read"
    }
    elseif ($v3 -lt 20.00000 -and $v3 -gt 5.00000 ) {
        $PassOrFail = "Fail"
        $Reason = "Slow Write"
    }
    elseif ($v2 -lt 1.00000 -or $v3 -lt 1.00000) {
        $PassOrFail = "Fail"
        $Reason = "Untestable"
    }
    else {
        $PassOrFail = "Pass"
    }
    
    #Write-Host "Result: $PassOrFail"

    #Matt
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Pass?' -Value $PassOrFail
    Add-Member -InputObject $o -MemberType noteproperty -Name 'Reason' -Value $Reason

    $resultFile = $PSScriptRoot + ".\temp.txt"

        # Passing results from Switchboard to SB_GUI through file reparsing
        Out-File -FilePath $resultFile -Force -InputObject $PassOrFail, $vRounded, $vRounded2

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
$junkFile=('{0}-Generation.xml' -f $batchId)
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


foreach ($test in <#@{name='Sequential read'; params='-b1M -o1 -t1 -w0 -Z1M'},
    @{name='Sequential write'; params='-b1M -o1 -t1 -w100 -Z1M'},#>
    @{name='Random QD4 write'; params='-b1M -o4 -t1 -r -w100 -Z1M'},
    @{name='Random QD4 read'; params='-b1M -o4 -t1 -r -w0 -Z1M'}
    <#,
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
        $testResult | Export-Csv $csv3outputPath -NoTypeInformation -Append
        $tests+=$testResult
}

   

# sum drive tests to a single row
$testsSum = TotalTargetRead $tests

# $testsSum 
$xmlJunkFile=('{0}-Random QD4 write.xml' -f $batchId)
$xmlJunkFile2=('{0}-Random QD4 read.xml' -f $batchId)
$date = Get-Date -Format "yyyy-MM-dd"
$csvoutputPath = Join-Path -Path $mainfolder -ChildPath ('LogFile.csv' <#the following command makes it so different dates will have differnet log files: "-f $date". Make sure to edit the filename to have "-{0} in it if you want to enable this!" #>)
$csv2outputPath = Join-Path -Path $mainfolder -ChildPath ('BD-{0}.csv' -f $newname)
$exceloutputPath = Join-Path -Path $mainfolder -ChildPath ('BD.xlsx')
#$excel2outputPath = Join-Path -Path $mainfolder -ChildPath ('BD-{0}.xlsx' -f $date)

$testsSum | Export-Csv -Path $csvoutputPath -NoTypeInformation -Append -Force
$testsSum | Export-Csv -Path $csv2outputPath -NoTypeInformation -Force



#FOR ZACH - This should work, but it doesnt and I have no idea why. 
$latestData = Import-Csv $csv2outputPath
$y_sr = @($latestData.Read)

$y_sw = @($latestData.Write)



#set chart definition
#See previous message about "Sequential Read" and "Sequential Write" having their "Sequential" name removed. They should be added back below        (here)           (and here) for public release!
$chart = New-ExcelChartDefinition -Title $newname -YMaxValue 100 <#-XAxisTitleText "Read & Write"#> -YAxisTitleText "Transfer Rate [MB/s]" -YRange 'Read','Write' -ChartType "BarClustered" -LegendBold -SeriesHeader "Read", "Write" -LegendSize 20 -TitleBold -TitleSize 20 -XAxisTitleSize 20 -YAxisTitleSize 20 -



# Get the latest .csv file
# Import the latest .csv file


# Extract the latest values for SequentialRead and SequentialWrite



# Output the latest values for SequentialRead and SequentialWrite
#Write-Host "Latest Sequential Read Value: $y_sr"
#Write-Host "Latest Sequential Write Value: $y_sw"


# Export data to an excel graph
Import-Csv -Path $csv2outputPath | Export-Excel $exceloutputPath -AutoNameRange -ExcelChartDefinition $chart -WorkSheetname $BDStandardName

##### format drive afterwards ######
Format-Volume -DriveLetter $drive -NewFileSystemLabel $newname

## Matts fucky wucky code

# Add-WordBarChart needs to receive the values as individual elements, not arrays
# Convert the arrays to individual elements if they are arrays

# Use spire to export the images
$workbook = New-Object Spire.Xls.Workbook
$workbook.LoadFromFile($exceloutputPath)

$sheet = $workbook.Worksheets[0] #gets the first sheet
$imgs = $workbook.SaveChartAsImage($sheet) # exports as an array

[string[]]$imageFileRelativePaths = @("") * $imgs.Count
# Save the charts to png files
for ($i = 0; $i -lt $imgs.Length; $i++) {
    $imageFileRelativePaths[$i] = ('{0}-Chart.png' -f $newname);
    $outputchartsPath = Join-Path -Path $chartsFolder -ChildPath $imageFileRelativePaths[$i] #need to create new path every time
    $fileStream = New-Object System.IO.FileStream($outputchartsPath, [System.IO.FileMode]::Create) #open new filestream
    $imgs[$i].Save($fileStream, [System.Drawing.Imaging.ImageFormat]::Png) #export as png
    $fileStream.Close() #can't keep open bc memory leaks
}

# For private fork version, the following should be COMMENTED OUT:
 $ChartImgDoc = Join-Path -Path $mainfolder -ChildPath ('ChartDoc-{0}.docx' -f $date) #chart doc yippee!

if ([System.IO.File]::Exists($ChartImgDoc)) {$Document = Get-OfficeWord $ChartImgDoc}
else {$Document = New-OfficeWord $ChartImgDoc}

New-OfficeWordImage -Document $Document -folderPath $mainfolder -imageFolder $chartsFolder -imageFile $imageFileRelativePaths[0] -openWord $true
# END OF COMMENTED OUT PART

# Check if any files need to be cleaned up
if (Test-Path $csv2outputPath, $exceloutputPath, $csv3outputPath, $junkFile, $xmlJunkFile, $xmlJunkFile2) {
    # Remove all files specified for cleanup
    Remove-Item -Path $csv2outputPath, $exceloutputPath, $csv3outputPath, $junkFile, $xmlJunkFile, $xmlJunkFile2 -Force
} else {
    Write-Host "No files created by the script found to delete."
}
