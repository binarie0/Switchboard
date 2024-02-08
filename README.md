# DriveTestUltra
 Tests specified storage device and exports to CSV, Excel charts, and [soon] organizes photos in a .docx file.
 
## How it works
 This program initially asks for the drive letter and the name you wish to set the drive to. It will then run preset testing (currently 1MB read/write for 3 seconds each) - custom testing functionality will be added at a later date. The program will then output to a .csv, conveniently located in your 'Documents' folder, as well as an .XLSX file with charts located in the same folder (separated by 'Worksheet' name which equals the new name). The program also outputs image files (.PNGs) of the graphs (charts) to "DTUCharts", a folder this program creates in your 'Documents' folder. In the future, there will be functionality to automatically organize graph (chart) images into a .docx document.
 
### Credits for [DriveTestUltra](https://github.com/binarie0/DriveTestUltra)        
 https://github.com/dfinke/ImportExcel [Dependency in order to develop graphs without having Excel natively installed]    
 https://github.com/ayavilevich/DiskSpdAuto [Used to heavily automate testing with DiskSPD]
 https://github.com/EvotecIT/PSWriteOffice [TBD]
 https://github.com/microsoft/diskspd [What this entire program is built off of]
 https://www.nuget.org/packages/FreeSpire.XLS [Excel graph (chart) output]
 https://www.nuget.org/downloads [One of the above required NuGet I forget which]

 binarie0 -> Majority of code compilation, graphics, output to .xlsx and .csv      
 EarthToFatt -> Initial conceptualization, code restructuring, graphics, graph (chart) output, .docx output (pending)      


#### NOTE:
This program currently requires its sister program, IFI Enabler, to function at full capacity. You can download IFI Enabler here: https://github.com/binarie0/IFI-Enabler

 
#### Changelog
    28 Jan 2024 - Initial Commit to Github (applying licenses and attaching actual code) (binarie0)
    01 Feb 2024 - Added ability to export to an excel graph + archived sister program IFI-Enabler (binarie0)
    08 Feb 2024 - General code restructuring, general README.md improvements (EarthToFatt)
