#Requires -RunAsAdministrator

$ErrorActionPreference = 'SilentlyContinue'

if ($global:timestamp_logfile -gt 0) {
    # Cool, do nothing
}
Else {
    $global:timestamp_logfile = (Get-Date -Format 'd-M-yyyy_HH-mm')
}

#startregion function Write-Log
function Write-Log {
    <#
    .SYNOPSIS      
        This function is for writing output to a logfile with several severities
    .PARAMETER Severity
        Specifiy the severity of the logging entry
        - Information
        - Warning
        - Error
    .PARAMETER Message
        The message entry       
    .EXAMPLE
        Write-Log -Severity Information -Message "This is a test"
    #>
    [CmdletBinding()]
    param (
        [parameter(
            Mandatory = $true,
            HelpMessage = "Specify log category ('Information', 'Warning' or 'Error'.",
            Position = 2
        )][alias('info')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information', 'Warning', 'Error')]
        [string]
        $Severity,      

        [parameter(
            Mandatory = $true,
            HelpMessage = "Specify log messaging.",
            Position = 1
        )][alias("msg")]
        [ValidateNotNullOrEmpty()]
        [string]
        $Message        
    ) # End param
    
    $time = Get-Date -Format "d-M-yyyy HH:mm:ss"
    Push-Location $PSScriptRoot

    Add-content -path $PSScriptRoot\Logging\LogFile_$timestamp_logfile.log -value "$severity $time $message"

} # End function
#Endregion function Write-Log

function Test-FileIsLocked {
    <#
    .SYNOPSIS
        This function is checking if a file is being used by another process.
    .EXAMPLE
        Write-Log -Severity 
    .PARAMETER Path
        Specify the file with path which needs to be checked.
    #>
    param (
        [parameter(Mandatory = $true)]
        [string]$Path
    )
  
    $oFile = New-Object System.IO.FileInfo $Path

    $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
  
    if ($oStream) {
        $false    
        $oStream.Close()
    }
    Else {
        $true
    }
}
function Resize-UserProfileDisk {
    <#
    .SYNOPSIS
        This script performs a compact action on the selected VHDX files. Without the need to have the Hyper-V role installed.
    .DESCRIPTION
        This script performs a compact action on the selected VHDX files.
        Without the need to have the Hyper-V role installed. 
        The script uses the defrag.exe utillity for defragmenting and the diskpart utillity for compacting the vhdx files. 
        To remove empty space on the vhdx file, sdelete.exe is used. 
        During the execution of the script, a check is made to see if the vhdx file is in use. 
        If the vhdx file is used by another process, the script will try to close the file, if this is not possible because the user is still using the User Profile Disk, for example, the vhdx file will be excluded from further actions.
        The script keeps a complete log of the actions that are performed and a calculation (in GB) is performed before the compacting takes place and at the end of the compacting process. 
        In this way, a total picture is made of the saving space. The results are written to a CSV file.
    .EXAMPLE
        Resize-UserProfileDisk -path D:\UPD -Defrag
        
        Runs a compact operation on all .vhdx files in path D:\UPD (recursive) because the -Defrag parameter is called, a defragmentation is performed on the .vhdx files first.
    .EXAMPLE
        Resize-UserProfileDisk -path D:\UPD -Defrag -ZeroFreeSpace

        Runs a compact operation on all .vhdx files in path D:\UPD (recursive) because the -Defrag parameter is called, a defragmentation is performed on the .vhdx files first. 
        Then there is a zero unused space operation with a clean free space action because the -ZeroFreeSpace parameter becomes groups.

    .EXAMPLE
        Resize-UserProfileDisk -SingleVhdxFile "D:\UPD\UVHD-S-1-5-21-1813453066-1828016147-3244441213-1125.vhdx" -Defrag -sDelete

        A compact version of the file D:\UPD\UVHD-S-1-5-21-1813453066-1828016147-3244441213-1125.vhdx is produced. Before the compact operation starts, the virtual hard disk is defragmented. 
        Since -sDelete is an alias on the ZeroFreeSpace a zero free space and a clean free space are executed before the virtual hard disk is compacted.
    .PARAMETER path
        Specify folder path which contains the User Profile Disks (vhdx-files). You cannot use this parameter along with the SingleVhdxFile parameter.
    .PARAMETER SingleVhdxFile
        Use this parameter if you want to run a compact output on a single VHDX file. You cannot use this parameter along with the path parameter. 
    .PARAMETER IncludeTemplate
        After the specification of this parameter the UVHD-Template.vhdx will also be compacted. You cannot use this parameter along with the SingleVhdxFile parameter.
    .PARAMETER Defrag
        After specifying this parameter the script will perform defragmentation of the specified vhdx-files. 
        It optimizes the allocation of space used by virtual hard disk files, except for fixed virtual hard disks. 
        When using this parameter, the files are allocated as optimally as possible so that the best possible disk space can be recovered.
    .PARAMETER ZeroFreeSpace
        When using this parameter, the software sdelete.exe is called with the parameters -z (Zero free space (good for virtual disk optimization) and -c (Clean free space). 
        Please note that when using this parameter, no active writing operations are performed on the vhdx files, otherwise data loss may occur.
    .NOTES
        Created by  : T13nn3s
        Version     : 3.2 (26 March 2020)
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')] 

    param (
        # Specify path where the UPDs are stored 
        [parameter(
            ParameterSetName = 'Object',
            Mandatory = $false,
            HelpMessage = "Specify path where the UPDs are stored",
            Position = 0
        )][string]$path,

        # This parameter cannot be used in combination with the -path parameter.
        [parameter(
            ParameterSetName = 'Object',
            Mandatory = $false,
            HelpMessage = "Specify the location of the .vhdx-file",
            Position = 1
        )][string]$SingleVhdxFile,

        # Includes the UHVHD-template.vhdx in the defragmentation and shrink proces
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Use this switch if you want to include the UVHD-template.vhdx",
            Position = 2
        )]
        [switch]
        $IncludeTemplate,

        # Defrag the VHDX-files
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Use this switch if you want to defragment the .vhdx file before a shrink is performed. The defrag.exe will be used with the /h /x switch.",
            Position = 3
        )]
        [switch]
        $Defrag,

        # Parameter help description
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Use this parameter to remove unused blocks in the .vhdx-file. For this parameter the script wil use the software Sdelete",
            Position = 4
        )][alias("sDelete")]
        [switch]
        $ZeroFreeSpace
    ) # End param

    Begin {
              
        #startregion General settings
        $timestamp = (Get-Date -Format 'd-M-yyyy_HH-mm')
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $timer.Start()
        $date = (Get-Date -Format 'd-M-yyyy')
        [System.Collections.ArrayList]$vhdxArray = @()

        try {
            New-Item -Path $PSScriptRoot\Logging -name Diskpart_log_$timestamp.log -ItemType file -Force | Out-Null   
            New-Item -Path $PSScriptRoot\diskpart_script.txt -ItemType file -Force | out-null

            if ((Test-Path $PSScriptRoot\Logging\Diskpart_log_$timestamp.log) -eq $true) {
                Write-Log -severity Information -message "Creating log Diskpart_log_$timestamp.log... Done"
            }
            
            $Calc = New-Object System.Collections.Generic.List[System.Object]
            Push-Location $PSScriptRoot
        }
        Catch {
            $error.clear()
            $ErrorMessage = $_.Exception.Message
            Write-Log -severity Error -message $ErrorMessage
            Write-Verbose
            return
        }
        #endregionregion General settings

        if (!$PSBoundParameters.ContainsKey('path') -and (!$PSBoundParameters.ContainsKey('SingleVhdxFile'))) {
            Write-host "WARNING: Please specify the -path or -SingleVhdxFile parameters to use this script." -ForegroundColor Yellow
            Write-Log -severity Warning -message "Please specify the -path or -SingleVhdxFile parameters to use this script."
            return
        }

        if ($path.Length -gt 0 -and $SingleVhdxFile.Length -gt 0) {
            Write-host "WARNING: Please do not use the -path and -SingleVhdxFile parameters at the same time." -ForegroundColor Yellow
            #Write-Log -severity Warning -message "Please do not use the -path and -SingleVhdxFile parameters at the same time."
            #Write-Verbose "Check 1"
            return
        }

        if ($PSBoundParameters.ContainsKey('IncludeTemplate') -and ($PSBoundParameters.ContainsKey('SingleVhdxFile'))) {
            Write-host "WARNING: Please do not use the -IncludeTemplate and -SingleVhdxFile parameters at the same time." -ForegroundColor Yellow
            Write-Log -severity Warning -message "Please do not use the -IncludeTemplate and -SingleVhdxFile parameters at the same time."
            return
        }

        if ($path) {
            # Check if the destination folder is existing
            if ((Test-Path $path) -eq $false) {
                Write-Host "Cannot find path $path." -ForegroundColor Red
                Write-Log -severity Error -message "Cannot find path $path"
                return
            }
        }

        if ($SingleVhdxFile -gt 0) {
            $path = $SingleVhdxFile

            #startregion Calculate size before compact
            Write-Log -severity Information -message "Start calculation .vhdx file size..."
            $measure_before = (($path | Measure-Object Length -s).Sum) / 1GB
            $measure_before = [math]::Round($measure_before / 1GB, 2)
            Write-Log -severity Information -message "Start calculation .vhdx file size... Done"
            #endregion Calculate size before compact
        }

        #startregion Calculate size before compact
        Write-Log -severity Information -message "---------- STARTING --------"
        Write-Log -severity Information -message "Start calculation .vhdx file sizes..."
        $measure_before = ((Get-ChildItem $path -Recurse | Where-Object { $_.name -notlike "UVHD-template.vhdx" -and $_.Extension -like "*.vhdx" } | Measure-Object Length -s).Sum) / 1GB
        $measure_before = [math]::Round($measure_before, 2)
        Write-Log -severity Information -message "Start calculation .vhdx file sizes... Done"
        #endregion Calculate size before compact
            
    } # Begin
 
    Process {

        if ($defrag) {
            Write-Log -severity Information -message "Defrag parameter detected. Script will defrag the vhdx-files prior to shrinking."
        }

        if ($SingleVhdxFile -gt 0) {
            $vhdxFiles = $SingleVhdxFile
        }
        Else {
            if (!$IncludeTemplate) {
                # Search the folder from top to down on only .vhdx-filesx
                Write-Log -severity Information -message "No IncludeTemplate parameter detected. UHVHD-Template.vhdx will be excluded"
                Write-Log -severity Information -message "Searching for .vhdx-files..."
                $vhdxFiles = Get-ChildItem $path -Recurse | Where-Object { $_.name -notlike "UVHD-template.vhdx" -and $_.Extension -like "*.vhdx" } | Select-Object -ExpandProperty Fullname
                Write-Log -severity Information -message "Searching for .vhdx-files... Done. Found $($vhdxfiles.count) vhdx-files."

                if ($vhdxfiles -eq 0) {
                    Write-Log -severity Information -message "No .vhdx-files found. Script is returning"
                    return
                }
            }
            Else {
                # Search the folder from top to down on only .vhdx-filesx include UVHD-template.vhdx
                Write-Log -severity Information -message "IncludeTemplate parameter detected. UHVHD-Template.vhdx will be excluded"
                Write-Log -severity Information -message "Searching for .vhdx-files..."
                $vhdxFiles = Get-ChildItem $path -Recurse | Where-Object { $_.Extension -like "*.vhdx" } | Select-Object -ExpandProperty Fullname
                Write-Log -severity Information -message "Searching for .vhdx-files... Done"

                if ($vhdxfiles -eq 0) {
                    Write-Log -severity Information -message "No .vhdx-files found."
                    return       
                }
            }   
        }

        #startregion foreach loop through vhdx files
        foreach ($vhdx in $vhdxFiles) {

            #startregion Check if file is locked
            if ((Test-FileIsLocked -Path $vhdx -ErrorAction SilentlyContinue) -eq $false) {
                $fileString = 'select vdisk file="' + $vhdx + '"' 
                Add-Content -Path diskpart_script.txt $fileString
                Add-Content diskpart_script.txt "attach vdisk readonly"
                Add-Content diskpart_script.txt "compact vdisk"
                Write-Log -severity Information -message "$vhdx is accessible."
                $vhdxArray.add($vhdx) | Out-Null
            }
            Else {
                Write-Log -severity Error -message "Access denied on $vhdx."                
            }          
        }
        #endtregion Check if file is locked
        
        #startregion defrag
        if ($defrag) {
            Write-Log -severity Information -message "Defrag parameter detected. Script wil defragging the VHDX-files. Script wil use the defrag.exe with the /x switches."

            foreach ($vhdx in $vhdxArray) {             
                try {
                    Write-Log -severity Information -message " "
                    Write-Log -severity Information -message "Defrag: Try to mount VHDX-file $vhdx..."
                    $MountResult = Mount-DiskImage $vhdx | Out-Null
                    Start-Sleep -Seconds 7
                    Write-Log -severity Information -message "Defrag: Try to mount VHDX-file $vhdx... Done"
                }
                Catch {
                    $error.clear()
                    $errormessage = $_.exception.message
                    Write-Log -severity Information -message "Defrag: Try to mount VHDX-file $vhdx... Failed. Error: $errormessage" 
                    return
                }

                try {
                    Write-Log -severity Information -message "Defrag: Checking for driveletter mounted $vhdx..." 
                    $DriveLetter = ($MountResult | Get-Volume).DriveLetter
                    Write-Log -severity Information -message "Defrag: Checking for driveletter mounted $vhdx... Done. Detected driveletter: $Driveletter"
                } 
                Catch {
                    $error.clear()
                    $errormessage = $_.exception.message
                    Write-Log -severity Error -message "Defrag: Checking for driveletter mounted disk... failed. Error: $errormessage"     
                    return               
                }                    
                Write-Log -severity Information -message "Defrag: Starting Defrag of $vhdx... Moutend with driveletter $DriveLetter"
                defrag.exe $Driveletter.Split("\")[0] -x

                Do {
                    $defragging = Get-Process defrag.exe -ErrorAction SilentlyContinue
                    Start-sleep -Seconds 7
                } while (Get-Process -Id $defragging.Id -ErrorAction SilentlyContinue)
                                                
                Write-Log -severity Information -message "Defrag: Defrag of $vhdx done..."
                    
                try {
                    Start-sleep -Seconds 10
                    Write-Log -severity Information -message "Defrag: Try to dismount VHDX-file $vhdx..."
                    try {
                        Dismount-DiskImage $vhdx | Out-Null
                        Write-Log -severity Information -message "Defrag: Try to dismount VHDX-file $vhdx... Done"
                    }
                    Catch {
                        $error.clear()
                        $ErrorMessage = $_.exception.Message
                        Write-Log -severity Error -message $errormessage
                    }
                }
                Catch {
                    $error.clear()
                    $errormessage = $_.exception.message
                    Write-Log -severity Information -message "Defrag: Try to dismount VHDX-file $vhdx... Failed. Error: $errormessage" 
                    return
                }                    
            } 
            #endregion defrag             
        }

        #startregion ZeroFreeSpace
        if ($ZeroFreeSpace) {
            Write-Log -severity Information -message "ZeroFreeSpace parameter detected. Script wil use SDelete to zero free space."

            foreach ($vhdx in $vhdxArray) {   

                # Checks if sDelete.exe exists
                Write-Log -severity Information -message "SDelete: Check if file SDelete.exe exists in location $PSScriptRoot"
                if (Test-Path $PSScriptRoot\sdelete.exe) {
                    Write-Log -severity Information -message "SDelete: SDelete.exe found..."
                }
                Else {
                    Write-Log -severity Error -message "SDelete: File SDelete.exe not found in location $PSScriptRoot."
                    Write-Log -severity Information -message "SDelete: This issue can be solved by placing SDelete.exe in folder $PSScriptRoot."
                    Write-Log -severity Information -message "SDelete: Download link: https://docs.microsoft.com/en-us/sysinternals/downloads/sdelete"
                    Write-Log -severity Error -message "SDelete: Script stopped."
                    return
                }
                
                try {
                    Write-Log -severity Information -message " "
                    Write-Log -severity Information -message "SDelete: Try to mount VHDX-file $vhdx..."
                    $MountResult = Mount-DiskImage $vhdx | Out-Null
                    Start-Sleep -Seconds 7
                    Write-Log -severity Information -message "SDelete: Try to mount VHDX-file $vhdx... Done"
                }
                Catch {
                    $error.clear()
                    $errormessage = $_.exception.message
                    Write-Log -severity Information -message "SDelete: Try to mount VHDX-file $vhdx... Failed. Error: $errormessage" 
                    return
                }

                try {
                    $DriveLetter = ($MountResult | Get-Volume).DriveLetter
                    Write-Log -severity Information -message "SDelete: Checking for driveletter mounted $vhdx..." 
                    Write-Log -severity Information -message "SDelete: Checking for driveletter mounted $vhdx... Done. Detected driveletter: $Driveletter"
                } 
                Catch {
                    $error.clear()
                    $errormessage = $_.exception.message
                    Write-Log -severity Error -message "SDelete: Checking for driveletter mounted disk... failed. Error: $errormessage"     
                    return               
                }                    
                Write-Log -severity Information -message "SDelete: Starting zero out of unused space from $vhdx... Moutend with driveletter $DriveLetter"
                .\sdelete.exe -z -c $Driveletter.Split("\")[0] /AcceptEULA

                Do {
                    $sdelete = Get-Process sdelete.exe -ErrorAction SilentlyContinue
                    Start-sleep -Seconds 7
                } while (Get-Process -Id $sdelete.Id -ErrorAction SilentlyContinue)
                                                
                Write-Log -severity Information -message "SDelete: Zero out unused space of $vhdx done..."
                    
                Write-Log -severity Information -message "SDelete: Try to dismount VHDX-file $vhdx..."
                try {
                    Dismount-DiskImage $vhdx | Out-Null
                }
                Catch {
                    $error.clear()
                    $ErrorMessage = $_.exception.Message
                    Write-Log -severity Error -message $errormessage
                }
                 
                #endregion Check mounted UPDs 
                Write-Log -severity Information -message "SDelete: Try to dismount VHDX-file $vhdx... Done"
            } 
        }

        #startregion Diskpart is compacting the .vhdx-file(s)
        try {
            Write-Log -severity Information -message "Starting shrinking of the .vhdx-files. DISKPART /S is using the $PSScriptRoot\diskpart_script.txt script file."
            DISKPART /s $PSScriptRoot\diskpart_script.txt > .\Logging\Diskpart_log_$timestamp.log
            Write-Log -severity Information -message "Starting shrinking of the .vhdx-files... Done"
        }
        Catch {
            $error.clear()
            $ErrorMessage = $_.exception.Message
            Write-Log -severity Error -message $errormessage
        }
        #endregion Diskpart is compacting the .vhdx-file(s)

    } # End Proces
    
    End {

        # Verify that all User Profile Disks have been dismounted to avoid temporary profiles on users.
        #startregion Check mounted UPDs
        Write-Log -severity Information -message "Check if there are still mounted UPDs"

        try {
            $mountedupd = Get-Volume | Where-Object filesystemlabel -eq "User Disk"
        }
        Catch {
            $error.clear()
            $ErrorMessage = $_.exception.Message
            Write-Log -severity Error -message $errormessage
        }
        Write-Log -severity Information -message "$($mountedupd.Count) mounted UPD(s) found. Try to dismounting."
        if ($mountedupd.Count -gt 0) {
            try {
                $mountedupd | ForEach-Object {
                    Get-DiskImage -DevicePath  $_.Path.trimend('\') -EA SilentlyContinue
                } | Where-Object ImagePath -notlike *$(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value)* | Dismount-DiskImage
            }
            Catch {
                $error.clear()
                $ErrorMessage = $_.exception.Message
                Write-Log -severity Error -message $errormessage
            }
        }
        #endregion Check mounted UPDs

        # Calculate file 
        $measure_after = ((Get-ChildItem $path -Recurse | Where-Object { $_.name -notlike "UVHD-template.vhdx" -and $_.Extension -like "*.vhdx" } | Measure-Object Length -s).Sum) / 1GB
        $measure_after = [math]::Round($measure_after, 2)

        # Total calculation and parse it to an array
        $savings = $measure_before - $measure_after
        $savings = [math]::Round($savings , 2)
              
        $timer.Stop()
        $ElapsedTime = "$($timer.Elapsed.Hours) hour, $($timer.Elapsed.Minutes) minutes and $($timer.Elapsed.Seconds) Seconds"
        $FilesCount = $vhdxFiles.count

        $ObjCalc = New-Object PSObject
        $ObjCalc | add-member Noteproperty "Date" $Date
        $ObjCalc | add-member Noteproperty "Path" $path
        $ObjCalc | add-member NoteProperty "Size Before (GB)" $measure_before
        $ObjCalc | Add-Member NoteProperty "Size After (GB)"  $measure_after
        $ObjCalc | add-member Noteproperty "Savings (GB)"  $savings
        $ObjCalc | add-member NoteProperty "Processed Files" $vhdxArray.count
        $ObjCalc | add-member NoteProperty "Script Runtime" $ElapsedTime
        $Calc.Add($ObjCalc)
        $ObjCalc | Export-Csv savings.csv -Append -NoTypeInformation -Delimiter ";"

        # Clear script files
        try {
            Clear-Content $PSScriptRoot\diskpart_script.txt -Force # Clearing diskpart script
            Write-log -severity Information -message "diskpart_script.txt cleared..."
        }
        Catch {
            $error.clear()
            $ErrorMessage = $_.exception.Message
            Write-Log -severity Error -message "clearing diskpart_script.txt failed. Error: $errormessage"
        }

        Write-Log -severity Information -message "Script has an elapsed time of $ElapsedTime"
        Write-Log -severity Information -message "Script has processed $filescount files"
        Write-Log -severity Information -message "Resizing-UserProfileDisk script ending..."
        Write-Log -severity Information -message "---------- ENDING --------"
    } # End 
} # End Resize-UserProfileDisk function 
