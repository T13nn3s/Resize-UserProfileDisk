> **Warning**
> This repository has been archived and is no longer actively maintained.

# Shrink UPD or other VHDX with Resize-UserProfileDisk
This Powershell Script is reclaiming white space from User Profile Disk without the need for the installation of the Hyper-V role. This script is scanning recursively for VHDX-files from the selected path. When the User Profile Disk is being used by another process this script will skip this UPD from compacting. This script is compacting the VHDX-files with the DISKPART utility. The script can perform VHDX file defragmentation if this parameter is used. This script has also the possibility to empty white space if this parameter is being used.

# How to use
First, you need to load this function into your Powershell session.
```powershell
Import-Module .\Resize-UserProfileDisk.ps1
```
Now you can use this function. If you only want to shrink the VHDX-files in an specified location. The script is recursive checking the location for VHDX-files. 
```powershell
Resize-UserProfileDisk -path D:\UPD
```
It is recommended to defragment the VHDX files before shrinking them. **(Recommended)**
```powershell
Resize-UserProfileDisk -path D:\UPD -defrag
```
By adding the -ZeroFreeSpace parameter all the free space on this disk will be written and removed to make sure that the free space is dfree on the VHDX-file.
```powershell
Resize-UserProfileDisk -path D:\UPD -defrag -ZeroFreeSpace
```
If you only want to shrink a single VHDX-file you need to specify the -SingleVhdxFile
```powershell
Resize-UserProfileDisk -SingleVhdxFile "D:\UPD\UVHD-S-1-5-21-1813453066-1828016147-3244441213-1125.vhdx" -defrag
```
It is also possible to use SDelete (Secure Delete) to clean free space. The -ZeroFreeSpace parameter will zero all the deleted disk on the VHDX-file.
```powershell
Resize-UserProfileDisk -SingleVhdxFile "D:\UPD\UVHD-S-1-5-21-1813453066-1828016147-3244441213-1125.vhdx" -defrag -ZeroFreeSpace
```
### NOTE
Please note that when using the parameter -ZeroFreeSpace that no active writing operations are performed on the VHDX-files, otherwise data loss may occur.

## Compact VHDX-files with Powershell
It is a huge challenge to compact User Profile Disks (and VHDX files in common), without the use of the Hyper-V role installed. If you have any recommendations or other improvements to this script, please let me know!

## MFT (Master File Table)
On all NFTS formatted file systems, information about a file, its creation date, size, permissions, and similar information is stored in the MFT (Master File Table). This file is placed in the middle of the disk and therefore a compact process may stop when it hits the MFT. It is not possible to compact along with the MFT. If you have a solution for moving the MFT to the end of the used space, please let me know or add it as your contribution to this script

# Changelog

## [3.2] 26 march 2020

### Fixed
- Sometimes path is not working
- Problem with VHDX-files not all dismounting

### Change
- Code cleanup

## [3.1] 15 october 2019

## Fixed 
- SingleVhdxFile not functioning fixed
- Fixed the creation of double logfiles
- Various small bugfixes

## [3.0] 14 october 2019
Script has been rewritten.

### Added 
- Parameter 'SingleVhdxFile'. With this parameter it is now possible to compact a single VHDX file.
- Parameter 'IncludeTemplate'. User Profile Disk template can be included (UHVHD-Template.vhdx).
- Parameter 'Defrag'. VHDX-files can be defragmented prior to compacting (defrag.exe)
- Parameter 'ZeroFreeSpace'. For optimal space saving, a zero free space can be executed for removed bits. (sdelete.exe is needed!)
- Write-Log function added
- Test-FileIsLocked function added

### Change
- Logging options are extended
- VHDX in use control enhanced
- VHDX file handling has been improved
- Script must now be opened with Administrator rights.

### Fixed
- Date and time notation.

## [1.2] 22 june 2019

### Added
- Adding sizing calculation with before and after compacting

### Changed
- Calculation will be exported to savings.csv file

## [1.0] 15 june 2019
Initial release of the script.
