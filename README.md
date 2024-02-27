# Switchboard: A Benchmark Tool
 Powershell program that tests specified storage device, formats said storage device, and exports test data to CSV, Excel charts, and [soon] organizes photos in a .docx file.
 
## How it works
 This program initially asks for the drive letter and the name you wish to set the drive to. It will then run preset testing (currently 1MB read/write for 3 seconds each) - custom testing functionality will be added at a later date. After this, the target drive is formatted. 

 WARNING: Do NOT use this program on a drive you wish to keep data on!

 The program will then output to a .csv, conveniently located in your 'Documents' folder, as well as an .XLSX file with charts located in the same folder (separated by 'Worksheet' name which equals the new name). The program also outputs image files (.PNGs) of the graphs (charts) to "DTUCharts", a folder this program creates in your 'Documents' folder. In the future, there will be functionality to automatically organize graph (chart) images into a .docx document.
 
### Credits for [DriveTestUltra](https://github.com/binarie0/DriveTestUltra)        
 https://github.com/dfinke/ImportExcel [Dependency in order to develop graphs (charts) without having Excel natively installed]    
 https://github.com/ayavilevich/DiskSpdAuto [Modified version of this is used to heavily automate testing with DiskSPD]        
 https://github.com/binarie0/PSWriteShmurd [Custom version of PSWriteWord with tiny changes, required due to PSWriteWord being deprecated. Required for native chart output to .docx (Feature TBA)]        
 https://github.com/EvotecIT/PSWriteWord [Origin of PSWriteSmurd]        
 https://github.com/microsoft/diskspd [Foundation of benchmark structure]        
 https://www.nuget.org/packages/FreeSpire.XLS [Excel graph (chart) output]        
 https://www.nuget.org/downloads [One of the above required NuGet I forget which lol]        

 binarie0 -> Majority of code compilation, graphics, output to .xlsx and .csv      
 EarthToFatt -> Initial conceptualization, code restructuring, graphics, graph (chart) output, .docx output (pending)      


#### NOTE:
This program currently requires its sister program, IFI Enabler, to function at full capacity. You can download IFI Enabler here: https://github.com/binarie0/IFI-Enabler

 
#### Changelog
    28 Jan 2024 - Initial Commit to Github (applying licenses and attaching actual code) (binarie0)
    01 Feb 2024 - Added ability to export to an excel graph + archived sister program IFI-Enabler (binarie0)
    08 Feb 2024 - General code restructuring, general README.md improvements (EarthToFatt)
    11 Feb 2024 - Unarchived IFI-Enabler as solution seems to need improvement, minor code restructuring and change as to how a file is created / edited (binarie0)
    24 Feb 2024 - Major testing on outputting charts to .docx natively, as well as testing outputting / organizing chart .png files to separate .docx. General code cleanup performed, and false drive name detection implemented. Program no longer uses PSWriteOffice, instead opting for PSWriteShmurd. (binarie0/EarthToFatt)
