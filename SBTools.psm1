$SBLogFile = "c:\support\SBLog.txt"

function Log {
<# 
 .Synopsis
  Function to log input string to file and display it to screen

 .Description
  Function to log input string to file and display it to screen. Log entries in the log file are time stamped. Function allows for displaying text to screen in different colors.

 .Parameter String
  The string to be displayed to the screen and saved to the log file

 .Parameter Color
  The color in which to display the input string on the screen
  Default is White
  Valid options are
    Black
    Blue
    Cyan
    DarkBlue
    DarkCyan
    DarkGray
    DarkGreen
    DarkMagenta
    DarkRed
    DarkYellow
    Gray
    Green
    Magenta
    Red
    White
    Yellow

 .Parameter LogFile
  Path to the file where the input string should be saved.
  Example: c:\log.txt
  If absent, the input string will be displayed to the screen only and not saved to log file

 .Example
  Log -String "Hello World" -Color Yellow -LogFile c:\log.txt
  This example displays the "Hello World" string to the console in yellow, and adds it as a new line to the file c:\log.txt
  If c:\log.txt does not exist it will be created.
  Log entries in the log file are time stamped. Sample output:
    2014.08.06 06:52:17 AM: Hello World

 .Example
  Log "$((Get-Location).Path)" Cyan
  This example displays current path in Cyan, and does not log the displayed text to log file.

 .Example 
  "Java process ID is $((Get-Process -Name java).id )" | log -color Yellow
  Sample output of this example:
    "Java process ID is 4492" in yellow

 .Example
  "Drive 'd' on VM 'CM01' is on VHDX file '$((Get-SBVHD CM01 d).VHDPath)'" | log -color Green -LogFile D:\Sandbox\Serverlog.txt
  Sample output of this example:
    Drive 'd' on VM 'CM01' is on VHDX file 'D:\VMs\Virtual Hard Disks\CM01_D1.VHDX'
  and the same is logged to file D:\Sandbox\Serverlog.txt as in:
    2014.08.06 07:28:59 AM: Drive 'd' on VM 'CM01' is on VHDX file 'D:\VMs\Virtual Hard Disks\CM01_D1.VHDX'

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/06/2014

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [String]$String, 
        [Parameter(Mandatory=$false,
                   Position=1)]
            [ValidateSet("Black","Blue","Cyan","DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","Red","White","Yellow")]
            [String]$Color = "White", 
        [Parameter(Mandatory=$false,
                   Position=2)]
            [String]$LogFile
    )

    write-host $String -foregroundcolor $Color 
    if ($LogFile.Length -gt 2) {
        ((Get-Date -format "yyyy.MM.dd hh:mm:ss tt") + ": " + $String) | out-file -Filepath $Logfile -append
    } else {
        Write-Verbose "Log: Missing -LogFile parameter. Will not save input string to log file.."
    }
}

function Get-SBVHD {
<# 
 .Synopsis
  Function to retrieve VHD file path for a Hyper-V Virtual Machine drive

 .Description
  Function to retrieve VHD file path for a VM drive. Errors related to Hyper-V hosts are saved in .\HVErrors.txt. Errors related to VMs are savved to .\VMErrors.txt
  Function ouputs an object 

 .Parameter VMName
  The virtual machine that we need to lookup one or more of its disks' VHD path
  Physical machines are ignored. Machines that cannot be accessed are ignored.
  Required. 4-15 characters long. Alias: VM. Positional parameter 0.
  Example: V-2012R2-Lab01

 .Parameter Drive
  The drive letter of the virtual machine drive that we need to identify its VHD path
  Optional. Gets all fixed drives that have a drive letter in the VM if omitted.
  Drive letters that don't exist on target VM are ignored.

 .Example
  Get-SBVHD OM01
  This example gets the VHD file names for all drives on VM OM01. Sample output:
    Drive HVName                                        VMName                                        VHDPath                                      
    ----- ------                                        ------                                        -------                                      
        F XHOST11                                       OM01                                          D:\VMs\Virtual Hard Disks\OM01_f1.VHDX       
        D XHOST11                                       OM01                                          D:\VMs\Virtual Hard Disks\OM01_D1.VHDX       
        C XHOST11                                       OM01                                          D:\VMs\Virtual Hard Disks\OM01_C1.vhdx  

 .Example
  $myVHD = Get-SBVHD -VMName V-2012R2-WEB5
  $myHVD | FL
  This example lists all volumes on VM V-2012R2-WEB5. Sample output:
    Drive   : E
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : D:\vms\v-2012r2-web5\SCSI2.vhdx

    Drive   : D
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : d:\VMs\V-2012R2-WEB5\Virtual Hard Disks\V-2012R2-WEB5-D.vhdx

    Drive   : F
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : D:\vms\v-2012r2-web5\IDE2.vhdx

    Drive   : H
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : D:\vms\v-2012r2-web5\IDE3.vhdx

    Drive   : I
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : D:\vms\v-2012r2-web5\SCSI3.vhdx

    Drive   : J
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : D:\vms\v-2012r2-web5\SCSI3.vhdx

    Drive   : K
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : D:\vms\v-2012r2-web5\SCSI3.vhdx

    Drive   : C
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : d:\VMs\V-2012R2-WEB5\Virtual Hard Disks\V-2012R2-WEB5-C.vhdx

    Drive   : L
    HVName  : XHOST12
    VMName  : V-2012R2-WEB5
    VHDPath : D:\vms\v-2012r2-web5\scsi4.vhdx

 .Example
  Get-SBVHD -VMName CM01 -Drive d -Verbose 
  This examples gets the VHD file name for drive d: of VM V-2012R2-WEB5.. Sample output:
    VERBOSE: Logging is disabled
    VERBOSE: VMError file name: VMErrors-0.txt
    VERBOSE: HVError file name: HVErrors-0.txt
    VERBOSE: Checking Drive input
    VERBOSE: Input drive(s): d
    VERBOSE: Performing the operation "Get-SBVHD" on target "CM01".
    VERBOSE: Connection to VM CM01 successfull
    VERBOSE: HVName: XHOST11
    VERBOSE: Openeing PS session to Hyper-V host: XHOST11
    VERBOSE: Connection to HV XHOST11 successfull
    VERBOSE: DiskNumber: 2, Controller: scsi, PartitionNumber: 1, Volume_DriveLetter: E
    VERBOSE: DiskNumber: 4, Controller: scsi, PartitionNumber: 1, Volume_DriveLetter: G
    VERBOSE: DiskNumber: 1, Controller: scsi, PartitionNumber: 1, Volume_DriveLetter: D
    VERBOSE: DiskNumber: 0, Controller: ide, PartitionNumber: 1, Volume_DriveLetter:  
    VERBOSE: DiskNumber: 0, Controller: ide, PartitionNumber: 2, Volume_DriveLetter: C
    VERBOSE: DiskNumber: 3, Controller: scsi, PartitionNumber: 1, Volume_DriveLetter: F
    VERBOSE: DiskNumber: 5, Controller: scsi, PartitionNumber: 1, Volume_DriveLetter: H
    VERBOSE: VHD Path: D:\VMs\Virtual Hard Disks\CM01_D1.VHDX
    Drive HVName                                        VMName                                        VHDPath                                      
    ----- ------                                        ------                                        -------                                      
        d XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_D1.VHDX

 .Example
  Get-SBVHD OM01,V-2012R2-WEB5 c
  This example gets the VHD file names for c: drives on VMs OM01,V-2012R2-WEB5, even if they're on different Hyper-V hosts. Sample output:
    Drive HVName                                        VMName                                        VHDPath                                      
    ----- ------                                        ------                                        -------                                      
        c XHOST11                                       OM01                                          D:\VMs\Virtual Hard Disks\OM01_C1.vhdx       
        c XHOST12                                       V-2012R2-WEB5                                 d:\VMs\V-2012R2-WEB5\Virtual Hard Disks\V-...

 .Example
  Get-Content VMs.txt | Get-SBVHD
  This example runs Get-SBVHD against each drive in each of the computers listed in Vms.txt file. Sample output:
    Drive HVName                                        VMName                                        VHDPath                                      
    ----- ------                                        ------                                        -------                                      
        E XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_E1.VHDX       
        G XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_G1.VHDX       
        D XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_D1.VHDX       
        C XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_C1.vhdx       
        F XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_F1.VHDX       
        H XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_H1.VHDX       
        F XHOST11                                       VMM01                                         D:\VMs\Virtual Hard Disks\VMM01_F1.VHDX      
        D XHOST11                                       VMM01                                         D:\VMs\Virtual Hard Disks\VMM01_D1.VHDX      
        C XHOST11                                       VMM01                                         D:\VMs\Virtual Hard Disks\VMM01_C1.vhdx  

 .Example
  Get-ADComputer -filter * | Select-Object @{n='VMName';e={$_.name}} | Get-SBVHD
  This example will return VHD path for each drive in each computer in the current domain. Physical machines will be ignored. Sample Output: 

    Drive HVName                                        VMName                                        VHDPath                                      
    ----- ------                                        ------                                        -------                                      
        C XHOST15                                       DC02                                          d:\VMs\DC02\Virtual Hard Disks\DC02_C1.vhdx  
        E XHOST11                                       DC01                                          D:\VMs\Virtual Hard Disks\DC01_E1.VHDX       
        D XHOST11                                       DC01                                          D:\VMs\Virtual Hard Disks\DC01_D1.VHDX       
        C XHOST11                                       DC01                                          D:\VMs\Virtual Hard Disks\DC01_C1.vhdx       
        F XHOST11                                       DC01                                          D:\VMs\Virtual Hard Disks\PDT2.vhdx          
        F XHOST11                                       RD01                                          D:\VMs\Virtual Hard Disks\RD01_F1.VHDX       
        D XHOST11                                       RD01                                          D:\VMs\Virtual Hard Disks\RD01_D1.VHDX       
        C XHOST11                                       RD01                                          D:\VMs\Virtual Hard Disks\RD01_C1.vhdx       
        E XHOST11                                       DB01                                          D:\VMs\Virtual Hard Disks\DB01_E1.VHDX       
        D XHOST11                                       DB01                                          D:\VMs\Virtual Hard Disks\DB01_D1.VHDX       
        C XHOST11                                       DB01                                          D:\VMs\Virtual Hard Disks\DB01_C1.vhd        
        F XHOST11                                       DB01                                          D:\VMs\Virtual Hard Disks\DB01_F1.VHDX       
        F XHOST11                                       VMM01                                         D:\VMs\Virtual Hard Disks\VMM01_F1.VHDX      
        D XHOST11                                       VMM01                                         D:\VMs\Virtual Hard Disks\VMM01_D1.VHDX      
        C XHOST11                                       VMM01                                         D:\VMs\Virtual Hard Disks\VMM01_C1.vhdx      
        F XHOST11                                       OR01                                          D:\VMs\Virtual Hard Disks\OR01_F1.VHDX       
        D XHOST11                                       OR01                                          D:\VMs\Virtual Hard Disks\OR01_D1.VHDX       
        C XHOST11                                       OR01                                          D:\VMs\Virtual Hard Disks\OR01_C1.vhdx       
        E XHOST11                                       DB02                                          D:\VMs\Virtual Hard Disks\DB02_E1.VHDX       
        D XHOST11                                       DB02                                          D:\VMs\Virtual Hard Disks\DB02_D1.VHDX       
        C XHOST11                                       DB02                                          D:\VMs\Virtual Hard Disks\DB02_C1.vhdx       
        F XHOST11                                       DB02                                          D:\VMs\Virtual Hard Disks\DB02_F1.VHDX       
        F XHOST11                                       OM01                                          D:\VMs\Virtual Hard Disks\OM01_f1.VHDX       
        D XHOST11                                       OM01                                          D:\VMs\Virtual Hard Disks\OM01_D1.VHDX       
        C XHOST11                                       OM01                                          D:\VMs\Virtual Hard Disks\OM01_C1.vhdx       
        E XHOST11                                       DB04                                          D:\VMs\Virtual Hard Disks\DB04_E1.VHDX       
        D XHOST11                                       DB04                                          D:\VMs\Virtual Hard Disks\DB04_D1.VHDX       
        C XHOST11                                       DB04                                          D:\VMs\Virtual Hard Disks\DB04_C1.vhdx       
        F XHOST11                                       DB04                                          D:\VMs\Virtual Hard Disks\DB04_F1.VHDX       
        E XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_E1.VHDX       
        G XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_G1.VHDX       
        D XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_D1.VHDX       
        C XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_C1.vhdx       
        F XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_F1.VHDX       
        H XHOST11                                       CM01                                          D:\VMs\Virtual Hard Disks\CM01_H1.VHDX       
        E XHOST11                                       DB05                                          D:\VMs\Virtual Hard Disks\DB05_E1.VHDX       
        G XHOST11                                       DB05                                          D:\VMs\Virtual Hard Disks\DB05_G1.VHDX       
        D XHOST11                                       DB05                                          D:\VMs\Virtual Hard Disks\DB05_D1.VHDX       
        I XHOST11                                       DB05                                          D:\VMs\Virtual Hard Disks\DB05_I1.VHDX       
        C XHOST11                                       DB05                                          D:\VMs\Virtual Hard Disks\DB05_C1.vhdx       
        F XHOST11                                       DB05                                          D:\VMs\Virtual Hard Disks\DB05_F1.VHDX       
        H XHOST11                                       DB05                                          D:\VMs\Virtual Hard Disks\DB05_H1.VHDX       
        E XHOST11                                       DPM01                                         D:\VMs\Virtual Hard Disks\DPM01_E1.VHDX      
        G XHOST11                                       DPM01                                         D:\VMs\Virtual Hard Disks\DPM01_G1.VHDX      
        D XHOST11                                       DPM01                                         D:\VMs\Virtual Hard Disks\DPM01_D1.VHDX      
        C XHOST11                                       DPM01                                         D:\VMs\Virtual Hard Disks\DPM01_C1.vhdx      
        F XHOST11                                       DPM01                                         D:\VMs\Virtual Hard Disks\DPM01_F1.VHDX      
        H XHOST11                                       DPM01                                         D:\VMs\Virtual Hard Disks\DPM01_H1.VHDX      
        F XHOST11                                       SM02                                          D:\VMs\Virtual Hard Disks\SM02_F1.VHDX       
        D XHOST11                                       SM02                                          D:\VMs\Virtual Hard Disks\SM02_D1.VHDX       
        C XHOST11                                       SM02                                          D:\VMs\Virtual Hard Disks\SM02_C1.vhdx       
        F XHOST11                                       SM01                                          D:\VMs\Virtual Hard Disks\SM01_F1.VHDX       
        D XHOST11                                       SM01                                          D:\VMs\Virtual Hard Disks\SM01_D1.VHDX       
        C XHOST11                                       SM01                                          D:\VMs\Virtual Hard Disks\SM01_C1.vhdx       
        E XHOST11                                       DB03                                          D:\VMs\Virtual Hard Disks\DB03_E1.VHDX       
        G XHOST11                                       DB03                                          D:\VMs\Virtual Hard Disks\DB03_G1.VHDX       
        D XHOST11                                       DB03                                          D:\VMs\Virtual Hard Disks\DB03_D1.VHDX       
        I XHOST11                                       DB03                                          D:\VMs\Virtual Hard Disks\DB03_I1.VHDX       
        C XHOST11                                       DB03                                          D:\VMs\Virtual Hard Disks\DB03_C1.vhdx       
        F XHOST11                                       DB03                                          D:\VMs\Virtual Hard Disks\DB03_F1.VHDX       
        H XHOST11                                       DB03                                          D:\VMs\Virtual Hard Disks\DB03_H1.VHDX       
        E XHOST12                                       V-2012R2-WEB5                                 D:\vms\v-2012r2-web5\SCSI2.vhdx              
        D XHOST12                                       V-2012R2-WEB5                                 d:\VMs\V-2012R2-WEB5\Virtual Hard Disks\V-...
        F XHOST12                                       V-2012R2-WEB5                                 D:\vms\v-2012r2-web5\IDE2.vhdx               
        H XHOST12                                       V-2012R2-WEB5                                 D:\vms\v-2012r2-web5\IDE3.vhdx               
        I XHOST12                                       V-2012R2-WEB5                                 D:\vms\v-2012r2-web5\SCSI3.vhdx              
        J XHOST12                                       V-2012R2-WEB5                                 D:\vms\v-2012r2-web5\SCSI3.vhdx              
        K XHOST12                                       V-2012R2-WEB5                                 D:\vms\v-2012r2-web5\SCSI3.vhdx              
        C XHOST12                                       V-2012R2-WEB5                                 d:\VMs\V-2012R2-WEB5\Virtual Hard Disks\V-...
        L XHOST12                                       V-2012R2-WEB5                                 D:\vms\v-2012r2-web5\scsi4.vhdx  

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/03/2014
  Next update: add logging feature, tweak error logs

#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [Alias('VM')]
            [ValidateLength(3,15)] # VM name cannot be lt 3 or gt 15
            [String[]]$VMName, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=1)]
            [Alias('Disk')]
            [Char[]]$Drive, 
        [switch]$Log
    )

    BEGIN { 
        if ($Log) {
            Write-Verbose "Logging is enabled"
            $i = 0
            do {
                $LogFile = "Get-VHD-$i.txt"
                $i++
            } while (Test-Path -Path $LogFile)
            Write-Verbose "LogFile name: $LogFile"
        } else {
            Write-Verbose "Logging is disabled"
        }

        $i = 0
        do {
            $VMError = "VMErrors-$i.txt"
            $i++
        } while (Test-Path -Path $VMError)
        Write-Verbose "VMError file name: $VMError"

        $i = 0
        do {
            $HVError = "HVErrors-$i.txt"
            $i++
        } while (Test-Path -Path $HVError)
        Write-Verbose "HVError file name: $HVError"

        Write-Verbose "Checking Drive input"
        if ($Drive -eq $null) {
            Write-Verbose "No drive(s) in input, getting all drives."
            $GetAllDrives = $true
        } else {
            Write-Verbose "Input drive(s): $Drive"
        }
    }

    PROCESS { 
        if ($PSCmdlet.ShouldProcess($VMName)) { 
            Write-Debug "Starting the process"

            foreach ($VM in $VMName) {
                try {
                    $contrinue = $true
                    $VMSession = New-PSSession -ComputerName $VM -ErrorVariable sbErr -ErrorAction 'stop'
                } catch { 
                    $contrinue = $false
                    Write-Verbose "Connection to VM $VM failed. For more details see $VMError file."
                    $sbErr | Out-File $VMError -Append
                }
                if ($contrinue) {
                    Write-Verbose "Connection to VM $VM successfull"
                    $IsVirtual = Invoke-Command -ScriptBlock { 
                        Test-Path -Path "HKLM:SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters\" } -Session $VMSession 
                    if ($IsVirtual) {
                        $HVName = Invoke-Command -ScriptBlock { 
                            (Get-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters\" -Name "PhysicalHostName").PhysicalHostName } -Session $VMSession 
                        Write-Verbose "HVName: $HVName"

                        if ((Get-PSSession -ComputerName $HVName -State Opened) -eq $null) {
                            Write-Verbose "Openeing PS session to Hyper-V host: $HVName"
                            try {
                                $HVSession = New-PSSession -ComputerName $HVName -ErrorVariable sbErr -ErrorAction 'stop'
                            } catch { 
                                $contrinue = $false
                                Write-Verbose "Connection to HV $HVName failed"
                                $sbErr | Out-File $HVError -Append
                            }
                        } else {
                            Write-Verbose "PS session already open to Hyper-V host: $HVName"
                        }
                
                        if ($contrinue) {
                            Write-Verbose "Connection to HV $HVName successfull"

                            # Build an object array to store each volume's partition and disk information
                            $properties = @{'DiskNumber' = $null;
                                            'Controller' = $null;
                                            'PartitionNumber' = $null;
                                            'DriveLetter' = $null}
                            [array]$sbVolumes = New-Object -TypeName psobject -Property $properties 
                            $AllDrives = @()
                            $Disks = Invoke-Command -ScriptBlock { Get-Disk } -Session $VMSession 
                            foreach ($Disk in $Disks) {
                                $Partitions = Invoke-Command -ScriptBlock { param($Disk)
                                    Get-Partition -Disk $Disk } -Session $VMSession -ArgumentList $Disk
                                foreach ($Partition in $Partitions) { 
                                    $Volumes = Invoke-Command -ScriptBlock { param($Partition)
                                        Get-Volume -Partition $Partition } -Session $VMSession -ArgumentList $Partition
                                    foreach ($Volume in $Volumes) {
                                        $Controller = ($Partition.DiskId.Split('#')[0]).Replace('\\?\','')
                                        Write-Verbose "DiskNumber: $($Disk.Number), Controller: $Controller, PartitionNumber: $($Partition.PartitionNumber), Volume_DriveLetter: $($Volume.DriveLetter)"
                                        $VolumeMap += $Disk.Number,$Controller,$Partition.PartitionNumber,$Volume.DriveLetter
                                        $properties = @{'DiskNumber' = $disk.Number;
                                                        'Controller' = $Controller;
                                                        'PartitionNumber' = $Partition.PartitionNumber;
                                                        'DriveLetter' = $Volume.DriveLetter}
                                        $sbVolume = New-Object -TypeName psobject -Property $properties 
                                        $sbVolumes += $sbVolume
                                        if ($sbVolume.DriveLetter -match "^[a-zA-Z]+$") { $AllDrives += $sbVolume.DriveLetter }
                                    }
                                }
                            }                        
                       
                            if ($GetAllDrives) { 
                                $Drive = $AllDrives 
                                Write-Verbose "Drives: $Drive"
                            }

                            $VMDisks = Invoke-Command -ScriptBlock { param($VM)
                                Get-VMHardDiskDrive -VMName $VM } -Session $HVSession -ArgumentList $VM
                            $SortedDisks = $VMDisks | Select -Property ControllerType,ControllerNumber,ControllerLocation,Path       

                            foreach ($DriveLetter in $Drive) { 
                                if ($DriveLetter -in $AllDrives) {
                                    $CurrentsbVolume = $sbVolumes | Where-Object { $_.DriveLetter -eq $DriveLetter }
                                    $VHD = $SortedDisks[$CurrentsbVolume.DiskNumber]
                                    Write-Verbose "VHD Path: $($VHD.Path)"

                                    Write-Debug "Finished querying, ready to produce output"
                                    $properties = @{'VMName' = $VM;
                                                    'HVName' = $HVName;
                                                    'Drive' = $DriveLetter;
                                                    'VHDPath' = $VHD.Path}
                                    $obj = New-Object -TypeName psobject -Property $properties
                                    $obj.PSObject.TypeNames.Insert(0,'SB.VHDPath') 

                                    Write-Debug "created output, ready to write to pipeline"
                                    Write-Output $obj
                                } else {
                                    Write-Verbose "Drive $DriveLetter not found on VM $VMName"
                                }
                            }
                        }
                        Remove-PSSession -Session $HVSession
                    } else {
                        Write-Verbose "Computer $VMName is not a vitual machine, ignoring.."
                    }
                    Remove-PSSession -Session $VMSession
                }
            }
        }
    }
}

function New-SBSeed {
<# 
 .Synopsis
  Function to create files for disk performance testing. Files will have random numbers as content. 
  File name will have 'Seed' prefix and .txt extension. 

 .Description
  Function will start with the smallest seed file of 10KB, and end with the Seed file specified in the -SeedSize parameter. 
  Function will create seed files in order of magnitude starting with 10KB and ending with 'SeedSize'. 
  Files will be created in the current folder.

 .Parameter SeedSize
  Size of the largest seed file generated. Accepted values are:
      10KB
      100KB
      1MB
      10MB
      100MB
      1GB
      10GB
      100GB
      1TB

 .Example
  New-SBSeed -SeedSize 10MB -Verbose
  
  This example creates seed files starting from the smallest seed 10KB to the seed size specified in the -SeedSize parameter 10MB.
  To see the output you can type in:

    Get-ChildItem -Path .\ -Filter *Seed*

  Sample output:
 
    Mode                LastWriteTime     Length Name                                                                                                                                      
    ----                -------------     ------ ----                                                                                                                                      
    -a---          8/6/2014   8:26 AM     102544 Seed100KB.txt
    -a---          8/6/2014   8:26 AM      10254 Seed10KB.txt
    -a---          8/6/2014   8:39 AM   10254444 Seed10MB.txt
    -a---          8/6/2014   8:26 AM    1025444 Seed1MB.txt

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/01/2014

#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [Alias('Seed')]
            [ValidateSet(10KB,100KB,1MB,10MB,100MB,1GB,10GB,100GB,1TB)] 
            [Int64]$SeedSize
    )

    $Acceptable = @(10KB,100KB,1MB,10MB,100MB,1GB,10GB,100GB,1TB)
    $Strings = @("10KB","100KB","1MB","10MB","100MB","1GB","10GB","100GB","1TB")
    for ($i=0; $i -lt $Acceptable.Count; $i++) {
        if ($SeedSize -eq $Acceptable[$i]) { $Seed = $i } 
    }
    $SeedName = "Seed" + $Strings[$Seed] + ".txt"
    if ($Acceptable[$Seed] -eq 10KB) { # Smallest seed starts from scratch
        $Duration = Measure-Command {
            do {Get-Random -Minimum 100000000 -Maximum 999999999 | out-file -Filepath $SeedName -append} while ((Get-Item $SeedName).length -lt $Acceptable[$Seed])
        }
    } else { # Each subsequent seed depends on the prior one
        $PriorSeed = "Seed" + $Strings[$Seed-1] + ".txt"
        if ( -not (Test-Path $PriorSeed)) { New-SBSeed $Acceptable[$Seed-1] } # Recursive function :)
        $Duration = Measure-Command {
            $command = @'
            cmd.exe /C copy $PriorSeed+$PriorSeed+$PriorSeed+$PriorSeed+$PriorSeed+$PriorSeed+$PriorSeed+$PriorSeed+$PriorSeed+$PriorSeed $SeedName /y
'@
            Invoke-Expression -Command:$command
            Get-Random -Minimum 100000000 -Maximum 999999999 | out-file -Filepath $SeedName -append
        }
    }
    Write-Verbose ("Created " + $Strings[$Seed] + " seed $SeedName file in " + $Duration.TotalSeconds + " seconds")
}

function Test-SBDisk {
<# 
 .Synopsis
  Function to test disk IO performance

 .Description
  This function tests disk IO performance by creating random files on the target disk and measuring IO performance
  It leaves 2 files in the WorkFolder:
    A log file that lists script progress, and
    a CSV file that has a record for each testing cycle

 .Parameter WorkFolder
  This is where the test and log files will be created. 
  Must be on a local drive. UNC paths are not supported.
  The function will create the folder if it does not exist
  The function will fail if it's run in a security context that cannot create files and folders in the WorkFolder
  Example: c:\support 
  
 .Parameter MaxSpaceToUseOnDisk
  Maximum Space To Use On Disk (in Bytes)
  Example: 10GB, or 115MB or 1234567890

 .Parameter Threads
  This is the maximum number of concurrent copy processes the script will spawn. 
  Maximum is 16. Default is 1. 

 .Parameter Cycles
  The function generates random files in a subfolder under the WorkFolder. 
  When the total WorkSubFolder size reaches 90% of MaxSpaceToUseOnDisk, the script deletes all test files and starts over.
  This is a cycle. Each cycle stats are recorded in the CVS and log files
  Default value is 3.

 .Parameter SmallestFile
  Order of magnitude of the smallest file.
  The function uses the following 9 orders of magnitude: (10KB,100KB,1MB,10MB,100MB,1GB,10GB,100GB,1TB) referred to as 0..8
  For example, SmallestFile value of 4 tells the script to use smallest file size of 100MB 
  The script uses a variable: LargestFile, it's selected to be one order of magnitude below MaxSpaceToUseOnDisk
  To see higher IOPS select a high SmallestFile value
  Default value is 4 (100MB). If the SmallestFile is too high, the script adjusts it to be one order of magnitude below LargestFile

 .Example
   Test-SBDisk "i:\support" 3GB
   This example tests the i: drive, generates files under i:\support, uses a maximum of 3GB disk space on i: drive.
   It runs a single thread, runs for 3 cycles, uses largest file = 100MB (1 order of magnitude below 3GB entered), smallest file = 10MB (1 order of magnitude below largest file)
   Sample screen output:
        Computer: XHOST16, WorkFolder = i:\support, MaxSpaceToUseOnDisk = 3GB, Threads = 1, Cycles = 3, SmallestFile = 10MB, LargestFile = 100MB
        Cycle #1 stats:
              Duration          10.02 seconds
              Files copied      2.70 GB
              Number of files   25
              Average file size 110.70 MB
              Throughput        276.33 MB/s
              IOPS              8.84k (32KB block size)
        Cycle #2 stats:
              Duration          9.83 seconds
              Files copied      2.70 GB
              Number of files   25
              Average file size 110.70 MB
              Throughput        281.60 MB/s
              IOPS              9.01k (32KB block size)
        Cycle #3 stats:
              Duration          12.12 seconds
              Files copied      2.70 GB
              Number of files   33
              Average file size 83.87 MB
              Throughput        228.25 MB/s
              IOPS              7.30k (32KB block size)
        Testing completed successfully.
   Sample log file content:
        2014.08.06 09:20:30 AM: Computer: XHOST16, WorkFolder = i:\support, MaxSpaceToUseOnDisk = 3GB, Threads = 1, Cycles = 3, SmallestFile = 10MB, LargestFile = 100MB
        2014.08.06 09:20:46 AM: Cycle #1 stats:
        2014.08.06 09:20:46 AM:       Duration          10.02 seconds
        2014.08.06 09:20:46 AM:       Files copied      2.70 GB
        2014.08.06 09:20:46 AM:       Number of files   25
        2014.08.06 09:20:46 AM:       Average file size 110.70 MB
        2014.08.06 09:20:46 AM:       Throughput        276.33 MB/s
        2014.08.06 09:20:46 AM:       IOPS              8.84k (32KB block size)
        2014.08.06 09:20:59 AM: Cycle #2 stats:
        2014.08.06 09:20:59 AM:       Duration          9.83 seconds
        2014.08.06 09:20:59 AM:       Files copied      2.70 GB
        2014.08.06 09:20:59 AM:       Number of files   25
        2014.08.06 09:20:59 AM:       Average file size 110.70 MB
        2014.08.06 09:20:59 AM:       Throughput        281.60 MB/s
        2014.08.06 09:20:59 AM:       IOPS              9.01k (32KB block size)
        2014.08.06 09:21:15 AM: Cycle #3 stats:
        2014.08.06 09:21:15 AM:       Duration          12.12 seconds
        2014.08.06 09:21:15 AM:       Files copied      2.70 GB
        2014.08.06 09:21:15 AM:       Number of files   33
        2014.08.06 09:21:15 AM:       Average file size 83.87 MB
        2014.08.06 09:21:15 AM:       Throughput        228.25 MB/s
        2014.08.06 09:21:15 AM:       IOPS              7.30k (32KB block size)
        2014.08.06 09:21:15 AM: Testing completed successfully.
   Sample CSV file content:
        Cycle #	Duration (sec)	Files (GB)	# of Files	Avg. File (MB)	Throughput (MB/s)	IOPS (K) (32KB blocks)	Machine Name	Start Time	End Time
        1	    10.02	        2.7	        25	        110.7	        276.33	            8.84	                XHOST16	        8/6/2014 9:20	8/6/2014 9:20
        2	    9.83	        2.7	        25	        110.7	        281.6	            9.01	                XHOST16	        8/6/2014 9:20	8/6/2014 9:20
        3	    12.12	        2.7	        33	        83.87	        228.25	            7.3	                    XHOST16	        8/6/2014 9:20	8/6/2014 9:21
        
 .Example
   Test-SBDisk "i:\support" 11GB -Threads 8 -Cycles 5 -SmallestFile 4
   This example tests the i: drive, generates files under i:\support, uses a maximum of 11GB disk space on i: drive, uses a maximum of 8 threads, runs for 5 cycles, and uses SamllestFile 100MB.

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 07/23/2014 - Script leaves log file and CSV file in the WorkFolder

#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Medium')] 
    Param(
        [Parameter (Mandatory=$true,
                    ValueFromPipeLine=$true,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=0)]
            [String]$WorkFolder,
        [Parameter (Mandatory=$true,
                    Position=1)]
            [Int64]$MaxSpaceToUseOnDisk,
        [Parameter (Mandatory=$false,
                    Position=2)]
            [ValidateRange(1,16)]
            [Int32]$Threads = 1,
        [Parameter (Mandatory=$false,
                    Position=3)]
            [ValidateRange(1,100)]
            [Int32]$Cycles = 3,
        [Parameter (Mandatory=$false,
                    Position=4)]
            [ValidateRange(0,8)]
            [Int32]$SmallestFile = 4
    )

    $Acceptable = @(10KB,100KB,1MB,10MB,100MB,1GB,10GB,100GB,1TB)
    $Strings = @("10KB","100KB","1MB","10MB","100MB","1GB","10GB","100GB","1TB")
    $BlockSize = (Get-WmiObject -Class Win32_Volume | Where-Object {$_.DriveLetter -eq ($WorkFolder[0]+":")}).BlockSize
    $logfile = (Get-Location).path + "\Busy_" + $env:COMPUTERNAME + (Get-Date -format yyyyMMdd_hhmmsstt) + ".txt"
    if ( -not (Test-Path $WorkFolder)) {
        try {
            New-Item -ItemType directory -Path $WorkFolder | Out-Null
        } catch {
            log "Error: WorkFolder $WorkFolder does not exist and unable to create it.. stopping" Magenta $logfile
            break
        }
    }
    Set-Location $WorkFolder 
    $logfile = $WorkFolder + "\Test-SBDisk_" + $env:COMPUTERNAME + (Get-Date -format yyyyMMdd_hhmmsstt) + ".txt"
    $WorkSubFolder = $WorkFolder + "\" + [string](Get-Random -Minimum 100000000 -Maximum 999999999) # Random name for work session subfolder
    $CSV = "$WorkFolder \Test-SBDisk_" + $env:COMPUTERNAME + (Get-Date -format yyyyMMdd_hhmmsstt) + ".csv"
    if ( -not (Test-Path $CSV)) {
        write-output ("Cycle #,Duration (sec),Files (GB),# of Files,Avg. File (MB),Throughput (MB/s),IOPS (K) (" + '{0:N0}' -f ($BlockSize/1KB) + "KB blocks),Machine Name,Start Time,End Time") | 
            out-file -Filepath $CSV -append -encoding ASCII
    }
    #
    # $LargestFile should be just below $MaxSpaceToUseOnDisk 
    if ($MaxSpaceToUseOnDisk -lt $Acceptable[0]) {
        log "Error: MaxSpaceToUseOnDisk $MaxSpaceToUseOnDisk is less than the minimum seed size of 10KB. MaxSpaceToUseOnDisk must be more than 10KB" Yellow $logfile; break
    } else {
        $LargestFile = '{0:N0}' -f ([Math]::Log10($MaxSpaceToUseOnDisk/10KB) - 1) # 1 order of magnitude below $MaxSpaceToUseOnDisk
    }
    if ($SmallestFile -gt $LargestFile-1) { $SmallestFile = $LargestFile-1 }
    log ("Computer: $env:COMPUTERNAME, WorkFolder = $WorkFolder, MaxSpaceToUseOnDisk = " + '{0:N0}' -f ($MaxSpaceToUseOnDisk/1GB) + "GB, Threads = $Threads, Cycles = $Cycles, SmallestFile = " + $Strings[$SmallestFile] + ", LargestFile = " + $Strings[$LargestFile]) Green $logfile
    New-SBSeed $Acceptable[$LargestFile] # Make seed files 
    #
    Get-Job | Remove-Job -Force # Remove any old jobs
    Get-ChildItem -Path $WorkFolder -Directory | Remove-Item -Force -Recurse -Confirm:$false # Delete any old subfolders
    New-Item -ItemType directory -Path $WorkSubFolder | Out-Null
    $StartTime = Get-Date
    $Cycle=0 # Cycle number
    do {
        # Delete all test files when you reach 90% capacity in $WorkSubFolder
        $WorkFolderData = Get-ChildItem $WorkSubFolder | Measure-Object -property length -sum
        $FolderFiles = "{0:N0}" -f $WorkFolderData.Count
        Write-Verbose ("WorkSubfolder $WorkSubFolder size " + '{0:N0}' -f ($WorkFolderData.Sum/1GB) + " GB, number of files = $FolderFiles")
        if ($WorkFolderData.Sum -gt $MaxSpaceToUseOnDisk*0.9) {
            $EndTime = Get-Date 
            do { # Wait for all jobs to finish before attempting to delete test files
                Get-Job -State Completed | Remove-Job # Remove completed jobs
                Start-Sleep 1
            } while ((Get-Job).Count -gt 0)
            Remove-Item $WorkSubFolder\* -Force 
            $CycleDuration = ($EndTime - $StartTime).TotalSeconds
            $Cycle++
            $CycleThru = ($WorkFolderData.Sum/$CycleDuration)/1MB # MB/s
            $IOPS = ($WorkFolderData.Sum/$CycleDuration)/$BlockSize
            log "Cycle #$Cycle stats:" Green $logfile
            log ("      Duration          " + "{0:N2}" -f $CycleDuration + " seconds") Green $logfile
            log ("      Files copied      " + "{0:N2}" -f ($WorkFolderData.Sum/1GB) + " GB") Green $Logfile
            log ("      Number of files   $FolderFiles") Green $Logfile
            log ("      Average file size " + "{0:N2}" -f (($WorkFolderData.Sum/1MB)/$FolderFiles) + " MB") Green $Logfile
            log ("      Throughput        " + "{0:N2}" -f $CycleThru + " MB/s") Yellow $Logfile
            log ("      IOPS              " + "{0:N2}" -f ($IOPS/1000) + , "k (" + "{0:N0}" -f ($BlockSize/1KB) + "KB block size)") Yellow $Logfile
            $CSVString = "$Cycle," + ("{0:N2}" -f $CycleDuration).replace(',','')  + "," + ("{0:N2}" -f ($WorkFolderData.Sum/1GB)).replace(',','')
            $CSVString += "," + $FolderFiles.replace(',','') + "," + ("{0:N2}" -f (($WorkFolderData.Sum/1MB)/$FolderFiles)).replace(',','')  + ","
            $CSVString += ("{0:N2}" -f $CycleThru).replace(',','') + "," + ("{0:N2}" -f ($IOPS/1000)).replace(',','')  + ","
            $CSVString += $env:COMPUTERNAME + "," + $StartTime + "," + $EndTime
            Write-Output $CSVString | out-file -Filepath $CSV -append -encoding ASCII
            $StartTime = Get-Date # Resetting $StartTime for next cycle
        } 
        if ((Get-Job).Count -lt $Threads+1) {
            Start-Job -ScriptBlock { param ($LargestFile,$WorkSubFolder,$Strings,$WorkFolder,$SmallestFile)
                $Seed2Copy = "Seed" + $Strings[(Get-Random -Minimum $SmallestFile -Maximum $LargestFile)] + ".txt" # Get a random seed
                $File2Copy =  $WorkSubFolder + "\" + [string](Get-Random -Minimum 100000000 -Maximum 999999999) + ".txt" # Get a random file name
                $Repeat = $Seed2Copy
                Set-Location $WorkFolder # Scriptblock runs at %HomeDrive%\%HomePath\Documents folder by default (like c:\users\samb\documents)
                for ($i=0; $i -lt (Get-Random -Minimum 0 -Maximum 9); $i++) {$Repeat += "+$Repeat"}
                $command = @'
                cmd.exe /C copy $Repeat $File2Copy /y
'@ 
                Invoke-Expression -Command:$command | Out-Null
                Get-Random -Minimum 100000000 -Maximum 999999999 | out-file -Filepath $File2Copy -append # Make all the files slightly different than each other
            } -ArgumentList $LargestFile,$WorkSubFolder,$Strings,$WorkFolder,$SmallestFile | Out-Null 
            Get-Job -State Completed | Remove-Job # Remove completed jobs
            Write-Verbose ("Current job count = " + (Get-Job).count) 
        }
        Get-Job -State Completed | Remove-Job # Remove completed jobs
        Write-Verbose ("Current Cycle = $Cycle, Cycles = $Cycles" )
    } while ($Cycle -lt $Cycles)
    log ("Testing completed successfully.") Green $logfile
}

function Get-SBRDPSession {
<# 
 .Synopsis
  Function to get RDP sessions on one or more computers

 .Description
  Function to get RDP sessions on one or more computers. Returns object collection, each corresponding to a session. 
  object properties: ComputerName, UserName, SessionName, ID, State
  ID refers to RDP session ID. State refers to RDP session State

 .Parameter ComputerName
  If absent, function assumes localhost.

 .Parameter State
  Filters result by one or more States
  Valid options are:
    Disc
    Conn
    Active
    Listen

 .EXAMPLE
  Get-SBRDPSession -ComputerName MyPC -State Disc | FT

  This example lists disconnected RDP sessions on the computer MyPC in table format.

  Sample output:

    UserName                State                   SessionName             ID                      ComputerName
    --------                -----                   -----------             --                      ------------
                            Disc                    services                0                       MyPC

 .EXAMPLE
  Get-SBRDPSession -state Active,Disc | FT

  This example lists RDP sessions on the local machine, and returns those with State Active or Disc in table format.

  Sample output:

    UserName                State                   SessionName             ID                      ComputerName
    --------                -----                   -----------             --                      ------------
                            Disc                    services                0                       PC1
    samb                    Active                  rdp-tcp#73              2                       PC1

 .Example
  Get-SBRDPSession xhost11,xhost12 | FT

  This example lists RDP sesions on the computers xHost11 and xHost12 and outputs the result in table format.

  Sample output:

    UserName                State                   SessionName             ID                      ComputerName
    --------                -----                   -----------             --                      ------------
                            Disc                    services                0                       xhost11
                            Conn                    console                 1                       xhost11
    samb                    Active                  rdp-tcp#10              2                       xhost11
                            Listen                  rdp-tcp                 65536                   xhost11
                            Disc                    services                0                       xhost12
                            Conn                    console                 1                       xhost12
                            Listen                  rdp-tcp                 65536                   xhost12

 .EXAMPLE
  Get-SBRDPSession (Get-Content .\computers.txt) Disc -Verbose |  FT

  This example reads a computer list from the file .\computers.txt and displays disconnected sessions
  In this example .\computers.txt contains:
    PC1
    PC2
    PC3

  Sample output:

    VERBOSE: Computer name(s): "xhost13" "xhost14" "xhost16"
    VERBOSE: Filtering on state(s): Disc
    VERBOSE: Reading RDP sessions on computer "xhost13"
    VERBOSE: Reading RDP session: ==> services                                    0  Disc
    VERBOSE: Reading RDP session: ==> console                                     1  Conn
    VERBOSE: Reading RDP session: ==> rdp-tcp                                 65536  Listen
    VERBOSE: Reading RDP sessions on computer "xhost14"
    VERBOSE: Reading RDP session: ==> services                                    0  Disc
    VERBOSE: Reading RDP session: ==> console                                     1  Conn
    VERBOSE: Reading RDP session: ==>                   samb                      2  Disc
    VERBOSE: Reading RDP session: ==> rdp-tcp                                 65536  Listen
    VERBOSE: Reading RDP sessions on computer "xhost16"
    VERBOSE: Reading RDP session: ==> services                                    0  Disc
    VERBOSE: Reading RDP session: ==> console                                     1  Conn
    VERBOSE: Reading RDP session: ==> rdp-tcp#73        samb                      2  Active
    VERBOSE: Reading RDP session: ==> rdp-tcp                                 65536  Listen

    UserName                State                   SessionName             ID                      ComputerNam
    --------                -----                   -----------             --                      -----------
                            Disc                    services                0                       PC1
                            Disc                    services                0                       PC2
    samb                    Disc                                            2                       PC2
                            Disc                    services                0                       PC3
 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/07/2014

#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$false,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [Alias('Host')]
            [String[]]$ComputerName,
        [Parameter(Mandatory=$false,
                   Position=1)]
            [ValidateSet("Disc","Conn","Active","Listen")] 
            [String[]]$State
    )

    if ( -not ($ComputerName)) { 
        $ComputerName = $env:COMPUTERNAME 
        Write-Verbose "No Computer Name(s) provided, assuming localhost.."
    }
    Write-Verbose "Computer name(s): $ComputerName"
    if ($State) { Write-Verbose "Filtering on state(s): $State" }
    $Fields = @(1,19,41,48,56)
    $ProprtyNames = @('SessionName','UserName','ID','State','ComputerName')
    $Properties = @{'SessionName'  = $null;
                    'ComputerName' = $null;
                    'UserName'     = $null;
                    'ID'           = $null;
                    'State'        = $null}

    foreach ($Computer in $ComputerName) {
        Write-Verbose "Reading RDP sessions on computer $Computer"
        $strRDPSession = query session /Server:$Computer 
        if ($strRDPSession -ne $null) {
            for ($LineNo=1; $LineNo -lt $strRDPSession.Count; $LineNo++) { 
                Write-Verbose "Reading RDP session: ==>$($strRDPSession[$LineNo])<=="
                $objRDPSession = New-Object -TypeName psobject -Property $Properties
                $objRDPSession.ComputerName = $Computer
                for ($i=0; $i -lt $Fields.Count-1; $i++) { 
                    $objRDPSession.$($ProprtyNames[$i]) = $strRDPSession[$LineNo].Substring($Fields[$i],$Fields[$i+1]-$Fields[$i]).trim()
                }
                if ($State) {
                    if ($objRDPSession.State -in $State) {
                        [array]$objRDPSessions += $objRDPSession
                    }
                } else {
                    [array]$objRDPSessions += $objRDPSession
                }
            }
        } else {
            Write-Verbose "Unable to connect to computer $Computer"
        }
    }    
    Write-Debug "Created output, ready to write to pipeline"
    Write-Output $objRDPSessions
}

Function Test-SBVHDIntegrity {
<#
 .Synopsis
  Test VHD files for integrity
 
 .Description
  Test the drive files for a Hyper-V virtual machine and verify the specified
  file exists and that there are no problems with it.

 .Parameter VM
  VM object. Function will check integrity of input VM disk files.
 
 .Example
  Get-VM | Test-SBVHDIntegrity | Out-Gridview
 
 .Example
  Get-VM -ComputerName xHost11,xHost12 | Test-SBVHDIntegrity | Out-Gridview
  This example displays disk integrity results for each VM on the Hyper-V hosts xHost11 and xHost12
 
 .Example
  Get-VM -ComputerName (Get-Content .\computers.txt) | Test-SBVHDIntegrity | Out-Gridview
  This example displays disk integrity results for each VM on the Hyper-V hosts listed n the .\computers.txt file
 
 .Example  
  Get-VM | Test-SBVHDIntegrity | where {(-NOT $_.TestPath) -OR (-NOT $_.TestVHD)}
  This examples displays disks that fail either the TestPath or TestVHD checks
 
 .Link
  Test-Path
  Test-VHD
 
 .Notes
  Function by Jeff Hicks
  http://www.altaro.com/hyper-v/testing-and-validating-hyper-v-vhdvhdx-files/
  Modified by Sam Boutros 08/09/2014, added functionality to work on VMs from other than localhost
#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param (
        [Parameter( Position=0,
                    Mandatory=$true,
                    ValueFromPipelineByPropertyName,
                    ValueFromPipeline)]
            [alias("vmname","name")]
            [ValidateNotNullorEmpty()]
            [object[]]$VM
    )
    Process {
        foreach ($item in $VM) {
            Write-Verbose "Analyzing disks for VM: $($item.VMName) on Hyper-V host: $($item.ComputerName)"
            Get-VMHardDiskDrive -VMName $item.VMName -ComputerName $item.ComputerName |
                Select VMName,@{ Name="Host" ;Expression={ $item.ComputerName } },Controller*,Path,
                @{ Name="TestPath";Expression={ Test-Path -Path "\\$($item.ComputerName)\$($_.Path.Replace(":","$"))" } },
                @{ Name="TestVHD" ;Expression={ $_ | Test-VHD -ComputerName $item.ComputerName } }
        } 
    }
}

Function Get-SBIPInfo {
<#
 .Synopsis
  Function to get computer IP information.
 
 .Description
  Function to get computer IP information.
  Function returns an object that has the following properties:
  ComputerName, AdapterDescription, IPAddress, IPVersion, SubnetMask, CIDR, MAC, DefaultGateway, DNSServer

 .Parameter ComputerName
  Name(s) of the computer(s) to be used to get their IP information.
 
 .Example
  Get-SBIPInfo
  This example retuns the current computer IP information

 .Example
  Get-SBIPInfo xhost11 | Out-GridView
  This example displays IP information of computer xHost11
 
 .Example
  Get-Content .\computers.txt | Get-SBIPInfo 
  This example returns IP information of every computer listed in the computers.txt file

 .Example
  Get-SBIPInfo | Where-Object { $_.IPversion -eq 4 } | Select-Object { $_.IPAddress }
  This example lists the IPv4 address(es) of the local computer
 
 .Outputs
  System.Object

 .Inputs
  System.String[]

 .Link
  http://superwidgets.wordpress.com/category/powershell/
 
 .Notes
  Function by Sam Boutros
  v1.0 - 08/12/2014
#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param (
        [Parameter( Position=0,
                    Mandatory=$false,
                    ValueFromPipelineByPropertyName,
                    ValueFromPipeline)]
            [alias("PC","Host")]
            [String[]]$ComputerName = $env:COMPUTERNAME
    )
    Begin {
        $Error.Clear()
        $Properties = @{'ComputerName'       = $null;
                        'AdapterDescription' = $null;
                        'IPAddress'          = $null;
                        'IPVersion'          = $null;
                        'SubnetMask'         = $null;
                        'CIDR'               = $null;
                        'MAC'                = $null;
                        'DefaultGateway'     = $null;
                        'DNSServer'          = $null}
    }
    Process {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Checking computer $Computer IP information.."
            try {
                $Adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer -ErrorAction Stop | 
                    Where-Object {$_.IPEnabled}
                Write-Verbose "Got computer $Computer IP information, processing.."
                foreach ($Adapter in $Adapters) {
                    for ($i=0; $i -lt $Adapter.IPAddress.Count; $i++) {
                        $objIPInfo = New-Object -TypeName psobject -Property $Properties
                        if ($Adapter.IPAddress[$i].Contains(".")) { 
                            $IPv = 4 
                            $CIDR = [Convert]::ToString(([Net.IPAddress]$Adapter.IPSubnet[$i]).Address,2).Replace('0','').Length
                            Write-Verbose "Adapter $($Adapter.Description) has IPv4 address $($Adapter.IPAddress[$i])"
                        } else { 
                            $IPv = 6 
                            $CIDR = $Adapter.IPSubnet[$i]
                            Write-Verbose "Adapter $($Adapter.Description) has IPv6 address $($Adapter.IPAddress[$i])"
                        }
                        $objSBIPInfo.ComputerName = $Computer
                        $objSBIPInfo.AdapterDescription = $Adapter.Description
                        $objSBIPInfo.IPAddress = $Adapter.IPAddress[$i]
                        $objSBIPInfo.IPVersion = $IPv
                        $objSBIPInfo.SubnetMask = $Adapter.IPSubnet[$i]
                        $objSBIPInfo.CIDR = "$($Adapter.IPAddress[$i])\$CIDR"
                        $objSBIPInfo.MAC = $Adapter.MACAddress
                        if ($Adapter.DefaultIPGateway) { $objSBIPInfo.DefaultGateway = $Adapter.DefaultIPGateway[$i] }
                        if ($Adapter.DNSServerSearchOrder) { $objSBIPInfo.DNSServer = $Adapter.DNSServerSearchOrder[$i] }
                        $objSBIPInfo
                    }
                }
                [array]$objSBIPInfos += $objSBIPInfo
            } catch {
                Write-Verbose "Unable to connect to computer $Computer"
                Write-Verbose "Error detail: $Error"
                Write-Output "Computer $Computer is offline or does not exist"
            }
        } 
    }
}

Function SBBytes {
<#
 .Synopsis
  Function to display Bytes in easy to read format.
 
 .Description
  While scripting we often have values of disk or memory sizes that are like 536870912 or 2147483648
  which are hard to read, compared to 512 Mb or 2 GB
  This function takes an inut like 2147483648 and returns a readable representation of it like 2 GB

 .Parameter Bytes
  Number of bytes
 
 .Example
  SBBytes 1234567890
  This example retuns 1.15 GB

 .Example
  hrbytes 123456789012345678901234567890123456789012345678901234567890
  This example returns 102,121,062,359,062,000,000,000,000,000,000,000.00 YB [YottaByte - 24 zeros]
 
 .Example
  hrbytes (Get-Volume)[0].SizeRemaining 
  This example returns the amount of available disk space on the first volume in easy to read format

 .Link
  http://superwidgets.wordpress.com/category/powershell/
 
 .Notes
  Function by Sam Boutros
  v1.0 - 08/14/2014
#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param (
        [Parameter( Position=0,
                    Mandatory=$true,
                    ValueFromPipelineByPropertyName,
                    ValueFromPipeline)]
            [ValidateNotNullOrEmpty()]
            [String]$Bytes 
    )
    Begin {
        $Units = @(" Bytes"," KB"," MB"," GB"," TB"," PB [PetaByte - 15 zeros]"," EB [ExaByte - 18 zeros]"," ZB [ZettaByte - 21 zeros]"," YB [YottaByte - 24 zeros]")
        if (([string]$Bytes).Contains("E+")) {
            $Digits = [int]([string]$Bytes).substring(([string]$Bytes).length-2,2)+1
        } else {$Digits = ([string]$Bytes).Length}
        $cUnit = 0
        if ($Digits -gt 3) { 
            do { $cUnit++; $Digits -=3 } 
                while (($Digits -gt 3) -and ($cUnit -lt $Units.Count-1))
        }
        $Value = '{0:N2}' -f ($Bytes/([Math]::Pow(2,$cUnit*10)))
        return $Value + $Units[$cUnit]  
    } 
}

function New-SBVM {
<# 
 .Synopsis
  Function to create Hyper-V virtual machines

 .Description
  This function automates the process of creating Hyper-V Virtual Machines. 

 .Parameter VMName
  VM Name: Mandatory. This is positional parameter 1
  2-15 characters long. Observe NetBIOS naming limitations.
  Must contain alphanumeric or these ! @ # $ % ^ & ( ) - _ ' { } . ~ characters only (http://support.microsoft.com/kb/188997)
  Example: V-2012R2-Lab01
  
 .Parameter VMFolder
  VM Folder: Mandatory. This is positional parameter 2
  This is the path to put all VM files under
  Example: g:\VMs\V-2012R2-Lab01 or \\SOFS\VMs\VM05

 .Parameter vSwitch
  vSwitch name: Mandatory. This is positional parameter 3
  It must exist on current Hyper-V host, example "My_vSwitch"

 .Parameter GoldenImageDiskPath
  Path to VHDX file: Mandatory. This is positional parameter  4
  This can be local or UNC path for a sys-prepped guest OS image.

 .Parameter VMG
  VM Generation: Optional. This is positional parameter 5
  Can be either 1 or 2. Assumed to be 1 if not provided.
  
 .Parameter VMMemoryType
  VM memory type: Optional. This is positional parameter 6
  Can be either "Static" or "Dynamic". Assumed to be Dynamic if not provided.

 .Parameter VMStartupRAM
  VM startup memory: Optional. This is positional parameter 7
  Assumed to be 1GB if not provided. 
  Examples: 536870912 or 512MB or 2GB

 .Parameter VMminRAM
  VM minimum memory: Optional. This is positional parameter 8
  Relevent if VMMemoryType is "Dynamic". Ignored if VMMemoryType is "Static". 
  Assumed to be 512MB if not provided. Minimal accepted value is 512MB.
  Examples: 536870912 or 512MB or 2GB

 .Parameter VMmaxRAM
  VM maximum memory: Optional. This is positional parameter 9
  Relevent if VMMemoryType is "Dynamic". Ignored if VMMemoryType is "Static". 
  Assumed to be 4GB if not provided. Minimal accepted value is 512MB.
  Examples: 1024MB or 2GB or 1TB

 .Parameter VMCores
  Number of CPU cores to assign to the VM: Optional. This is positional parameter 10
  Example: 4

 .Parameter VLAN
  VLAN ID: Optional. This is positional parameter 11
  Example: 193

 .Parameter AdditionalDisksTotal
  Number of additional disks to create and attach to VM. 
  Optional. This is positional parameter 12
  Valid range: 1-256. They will be dynamic disks.

 .Parameter AdditionalDisksSize
  Size of each of the additional disks.
  Optional. This is positional parameter 13
  If AdditinalDisksTotal is provided and AdditionalDisksSize is not, size is assumed to be 1GB.
  Examples: 536870912 or 512MB or 2GB or 10TB
  
 .Parameter CSV
  Path to CSV file used to store disk copy statistics.
  Optional. This is positional parameter  14
  Script appends to existing CSV file

 .Example
   New-SBVM "VM01" "k:\VMs\VM01" "My_vSwitch" "E:\Golden\V-2012R2-3-C.VHDX"

 .Example
   New-SBVM "VM02" "i:\VMs\VM02" "My_vSwitch" "E:\Golden\V-2012R2-3-C.VHDX" -VMMemoryType Static

 .Example
   $VMName = "VM01"
   $VMFolder = "k:\VMs\VM01"
   $VMG = 2 
   $VMStartupRAM = 1GB 
   $VMminRAM = 512MB
   $VMmaxRAM = 1GB
   $vSwitch = "My_vSwitch"
   $VMCores = 2 
   $GoldenImageDiskPath = "E:\Golden\V-2012R2-3-C.VHDX" 
   New-SBVM -VMName $VMName -VMFolder $VMFolder -vSwitch $vSwitch -GoldenImageDiskPath $GoldenImageDiskPath -VMG $VMG -VMStartupRAM $VMStartupRAM -VMminRAM $VMminRAM -VMmaxRAM $VMmaxRAM -VMCores $VMCores -verbose
   
   This example creates a single VM

.Example
   $NumberofVMs = 5
   $VMrootPath = "G:\VMs"
   $VMPrefix = "V-2012R2-LAB"
   For ($j=1; $j -lt $NumberofVMs+1; $j++) {
       $Params = @{ VMName = $VMPrefix + $j;
                    VMFolder = $VMRootPath + "\" + $VMPrefix + $j;
                    VMG = 2;
                    VMMemoryType = "Dynamic";
                    VMStartupRAM = 1GB; 
                    VMminRAM = 512MB;
                    VMmaxRAM = 2GB;
                    vSwitch = "My_vSwitch";
                    VMCores = 2; 
                    VLAN = 19; 
                    AdditionalDisksTotal = 3; 
                    AdditionalDisksSize = 2TB;
                    GoldenImageDiskPath = "E:\Golden\V-2012R2-3-C.VHDX"
       }
       New-SBVM @Params
       Start-VM -Name ($VMPrefix + $j)
   }

   This example creates 5 VMs and starts them.
   VM names will be V-2012R2-Lab1, V-2012R2-Lab2,..

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/14/2014

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Medium')] 
    Param(
        [Parameter (Mandatory=$true,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=0)]
            [Alias('VM','ComputerName','PC')]
            [ValidateLength(2,15)] # VM name cannot be lt 2 or gt 15
            [ValidatePattern("^[!@#$%^&()-_'{}.~a-zA-Z0-9]+$")] # NetBIOS naming restriction http://support.microsoft.com/kb/188997
            [String]$VMName,
        [Parameter (Mandatory=$true,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=1)]
            [ValidateLength(4,240)] 
            [String]$VMFolder,
        [Parameter (Mandatory=$true,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=2)]
            [String]$vSwitch,
        [Parameter (Mandatory=$true,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=3)]
            [Alias('vmTemplate')]
            [ValidateScript({ Test-Path $_ })]
            [String]$GoldenImageDiskPath,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=4)]
            [ValidateRange(1,2)] 
            [Alias('Generation')]
            [Int32]$VMG = 1,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=5)]
            [ValidateSet("Dynamic","Static")] 
            [Alias('RAMType','MemoryType')]
            [String]$VMMemoryType = "Dynamic",
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=6)]
            [ValidateRange(512MB,1TB)] # Hyper-V 3 max http://technet.microsoft.com/en-us/library/jj680093.aspx
            [Alias('StartupRAM')]
            [Int64]$VMStartupRAM = 1GB,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=7)]
            [ValidateRange(512MB,1TB)] # Hyper-V 3 max http://technet.microsoft.com/en-us/library/jj680093.aspx
            [Alias('MinRAM')]
            [Int64]$VMMinRam = 512MB,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=8)]
            [ValidateRange(512MB,1TB)] # Hyper-V 3 max http://technet.microsoft.com/en-us/library/jj680093.aspx
            [Alias('MaxRAM')]
            [Int64]$VMMaxRAM = 4GB,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=9)]
            [ValidateRange(1,64)] # Hyper-V 3 max http://technet.microsoft.com/en-us/library/jj680093.aspx
            [Alias('Cores')]
            [Int32]$VMCores = 1,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=10)]
            [ValidateRange(0,4095)] 
            [Int32]$VLAN = 0,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=11)]
            [ValidateRange(0,256)] 
            [Alias('AdditionalDisks')]
            [Int32]$AdditionalDisksTotal = 0,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=12)]
            [ValidateRange(1GB,64TB)] 
            [Int64]$AdditionalDisksSize = 1GB,
        [Parameter (Mandatory=$false,
                    ValueFromPipeLineByPropertyName=$true,
                    Position=13)]
            [String]$CSV = “.\IOPS_$($VMName)_$(Get-Date -format yyyyMMdd_hhmmsstt).csv”
    )
    #
    $logfile = ".\New-SBVM_$($VMName)_$(Get-Date -format yyyyMMdd_hhmmsstt).txt"
    if ( -not (Test-Path $CSV)) {
        Write-Output "Source, Destination, File size (GB), Copy start time, Copy end time, VM Name, Copy duration (sec), Throughput (MB/s), Throughput (Mbps)" | 
            Out-File -Filepath $CSV -Append -Encoding ASCII 
    }
    Write-Verbose "Input parameters:" 
    Write-Verbose "VMName: $VMName"
    Write-Verbose "VMFolder: $VMFolder"
    Write-Verbose "vSwitch: $vSwitch"
    Write-Verbose "GoldenImageDiskPath: $GoldenImageDiskPath" 
    Write-Verbose "VMG: $VMG"
    Write-Verbose "VMMemoryType: $VMMemoryType"
    Write-Verbose "VMStartupRAM: $(SBBytes $VMStartupRAM)"
    Write-Verbose "VMminRAM: $(SBBytes $VMMinRAM)"
    Write-Verbose "VMmaxRAM: $(SBBytes $VMMaxRAM)"
    Write-Verbose "VMCores: $VMCores" 
    Write-Verbose "VLAN: $VLAN" 
    Write-Verbose "AdditionalDisksTotal: $AdditionalDisksTotal"
    Write-Verbose "AdditionalDisksSize: $(SBBytes $AdditionalDisksSize)"
    Write-Verbose "===================================================="
    #
    # Some error checking:
    $ErrorList = @()
    $WarningList = @()
    If ( -not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        $ErrorList += "Error: This script must be run under elevated credentials. Please run this script 'as Administrator'" 
    }
    if ( -not (Test-Path $VMFolder)) { 
        New-Item -type directory -path $VMFolder -ErrorAction SilentlyContinue | Out-Null
    }
    if ( -not (Test-Path $VMFolder)) { 
        $ErrorList += "Error: Failed to create VM folder '$VMFolder'" 
    }
    $VMSwitch = Get-VMSwitch
    $vSwitchExists = $false
    for ($i=0; $i -lt $VMSwitch.Count; $i++) { 
        if ($vSwitch -match $VMSwitch.name[$i]) { $vSwitchExists=$true } 
    }
    if ( -not $vSwitchExists) { $ErrorList += "Error: vSwitch '$vSwitch' does not exist on host '$env:COMPUTERNAME'" }
    if ($VMMemoryType -match "static") { $Dynamic=$false } else { $Dynamic=$true } 
    $HVHost = Get-WmiObject Win32_ComputerSystem -ComputerName $env:COMPUTERNAME
    [long]$HostRAM = $HVHost.TotalPhysicalMemory
    If ($VMStartupRAM -lt 512MB) { 
        $VMStartupRAM=1GB 
        Write-Verbose "VMStartupRAM adjusted to: $(SBBytes $VMStartupRAM)"
    }
    If ($VMMinRAM -lt 512MB) { 
        $VMMinRAM=512MB 
        Write-Verbose "VMMinRAM adjusted to: $(SBBytes $VMMinRAM)"
    }
    If ($VMMaxRAM -lt 512MB) { 
        $VMMaxRAM=4GB 
        Write-Verbose "VMMaxRAM adjusted to: $(SBBytes $VMMaxRAM)"
    }
    Write-Verbose "HostRAM: $(SBBytes $HostRAM)"
    if ($VMStartupRAM -gt $HostRAM) {
        $ErrorList += "Error: cannot assign $(SBBytes $VMStartupRAM) for VM '$VMName' startup RAM. The host '$env:COMPUTERNAME' has a total of $(SBBytes $HostRAM) RAM" 
    }
    if ($VMMinRAM -gt $HostRAM) {
        $ErrorList += "Error: cannot assign $(SBBytes $VMMinRAM) for VM '$VMName' minimum RAM. The host '$env:COMPUTERNAME' has a total of $(SBBytes $HostRAM) RAM"
    }
    if ($VMMaxRAM -gt $HostRAM) {
        $ErrorList += "Error: cannot assign $(SBBytes $VMMaxRAM) for VM '$VMName' maximum RAM. The host '$env:COMPUTERNAME' has a total of $(SBBytes $HostRAM) RAM"
    }
    if ($VMMinRAM -gt $VMMaxRAM) {
        $WarningList += "Warning: VM maximum memory $(SBBytes $VMMaxRAM) cannot be less than VM minimum memory $(SBBytes $VMMinRAM). Setting VM maximum memory to $(SBBytes $VMMinRAM)"
        $VMMaxRAM = $VMMinRAM
        Write-Verbose "VMMaxRAM adjusted to: $(SBBytes $VMMaxRAM)"
    }
    if ($VMMinRAM -gt $VMStartupRAM) {
        $WarningList += "Warning: VM Startup memory $(SBBytes $VMStartupRAM) cannot be less than VM minimum memory $(SBBytes $VMMinRAM). Setting VM Startup memory to $(SBBytes $VMMinRAM)"
        $VMStartupRAM = $VMMinRAM
        Write-Verbose "VMStartupRAM adjusted to: $(SBBytes $VMStartupRAM)"
    }
    if ($VMStartupRAM -gt $VMMaxRAM) {
        $WarningList += "Warning: VM Startup memory $(SBBytes $VMStartupRAM) cannot be less than VM maximum memory $(SBBytes $VMMaxRAM). Setting VM Startup memory to $(SBBytes $VMMaxRAM)"
        $VMStartupRAM = $VMMaxRAM
        Write-Verbose "VMStartupRAM adjusted to: $(SBBytes $VMStartupRAM)"
    }
    [int]$HVHostCores = $HVHost.NumberOfLogicalProcessors
    if ($VMCores -gt $HVHostCores) { 
        $WarningList += "Warning: Cannot assign '$VMCores' CPU cores to VM '$VMName' - the Hyper-V host '$env:COMPUTERNAME' has '$HVHostCores' cores. Assinging 1 CPU Core"
        $VMCores=1 
        Write-Verbose "VMCores adjusted to: $VMCores"
    }
    If (-not $HVHost.HypervisorPresent) { 
        $ErrorList += "Error: Hyper-V role is not installed on this host" 
    }
    if ($ErrorList) { 
        log "Errors found:" Magenta $logfile
        For ($i=0; $i -lt $ErrorList.Count; $i++) { log $ErrorList[$i] Magenta $logfile } 
    }
    if ($WarningList) { 
        log "Warnings found:" Yellow $logfile
        For ($i=0; $i -lt $WarningList.Count; $i++) {log $WarningList[$i] Yellow $logfile } 
    }
    if ($ErrorList) { 
        log "Stopping.." Magenta $logfile
        log " " yellow $logfile
        log "VMName: $VMName" yellow $logfile
        log "VMFolder: $VMFolder" yellow $logfile
        log "vSwitch: $vSwitch" yellow $logfile
        log "GoldenImageDiskPath: $GoldenImageDiskPath"  yellow $logfile
        log "VMG: $VMG" yellow $logfile
        log "VMMemoryType: $VMMemoryType" yellow $logfile
        log "VMStartupRAM: $(SBBytes $VMStartupRAM)" yellow $logfile
        log "VMminRAM: $(SBBytes $VMMinRAM)" yellow $logfile
        log "VMmaxRAM: $(SBBytes $VMMaxRAM)" yellow $logfile
        log "VMCores: $VMCores"  yellow $logfile
        log "VLAN: $VLAN"  yellow $logfile
        log "AdditionalDisksTotal: $AdditionalDisksTotal" yellow $logfile
        log "AdditionalDisksSize: $(SBBytes $AdditionalDisksSize)" yellow $logfile
        log "Dynamic memory: $Dynamic" Yellow $logfile
        break 
    }
    # Create VM:
    log " " -LogFile $logfile
    log "Creating VM with the following parameters:" green $logfile
    log "VM Name: '$VMName'" green $logfile
    log "VM Folder: '$VMFolder'" green $logfile
    log "VM Generation: '$VMG'" green $logfile
    if ($Dynamic) {
        log "VM Memory: Dynamic, Startup memory: $(SBBytes $VMStartupRAM), Minimum memory: $(SBBytes $VMMinRAM), Maximum memory: $(SBBytes $VMMaxRAM)" green  $logfile
    } else { 
        log "VM Memory: Static: $(SBBytes $VMStartupRAM)" green $logfile
    }
    log "vSwitch Name: '$vSwitch'" green $logfile
    log "VM CPU Cores: '$VMCores'" green $logfile
    $CDrive = $VMFolder + "\" + $VMName + "-C.VHDX"
    log "Copying disk image $CDrive from golden image $GoldenImageDiskPath" green $logfile
    $CopyStart = Get-Date
    Copy-Item -Path $GoldenImageDiskPath -Destination $CDrive -Force
    $CopyEnd = Get-Date
    if (-not (Test-Path $CDrive)) { log "Error: Failed to copy file '$CDrive'" Magenta $logfile; break }
    $CopyTime = $CopyEnd - $CopyStart
    $Filesize = (Get-Item $CDrive).length
    $Thru = ($filesize*1000)/($CopyTime.TotalMilliseconds*1MB)
    log "Copy time: $('{0:c}' -f $CopyTime) (hh:mm:ss.fffffff)" Cyan
    log "File $CDrive, size: $(SBBytes $Filesize)" cyan $logfile
    log "Throughput: $('{0:N2}' -f $Thru) MB/s = $('{0:N2}' -f ($Thru*8)) Mbps" cyan $logfile
    write-output "$GoldenImageDiskPath, $CDrive, $('{0:N2}' -f ($Filesize/1GB)), '$CopyStart', '$CopyEnd', $VMName, $($CopyTime.TotalSeconds), $Thru, $($Thru*8)" | 
        out-file -Filepath $CSV -append -encoding ASCII
    $error.Clear()
    try {
        log ((New-VM -VHDPath $CDrive -MemoryStartupBytes $VMStartupRAM -Generation $VMG  -Name $VMName -Path $VMFolder -SwitchName $vSwitch) | Out-String) -LogFile $logfile
    } catch {
        log "Error: Failed to create VM '$VMName'.. stopping" Magenta $logfile
        Log "Error details: $error" Magenta $logfile
        break
    }
    # Configure VM:
    log "Configuring memory for VM '$VMName'" green $logfile
    try {
        if ($Dynamic) { 
            log ((Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $True -MinimumBytes $VMMinRAM -MaximumBytes $VMMaxRAM -Passthru) | Out-String) -LogFile $logfile
        } else {
            log ((Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $false -Passthru) | Out-String) -LogFile $logfile
        }
    
    } catch {
        log "Error configuring memory for VM '$VMName'.. stopping" Magenta
        Log "Error details: $error" Magenta $logfile
        break
    }
    log "Configuring CPU cores for VM '$VMName'" green $logfile
    try {
        log ((Set-VMProcessor -VMName $VMName -Count $VMCores -Passthru) | Out-String) -LogFile $logfile
    } catch {
        log "Error configuring CPU cores for VM '$VMName'.. stopping" Magenta $logfile
        Log "Error details: $error" Magenta $logfile
        break
    }
    log "Configuring vNIC for VM '$VMName'" green $logfile
    try {
        log ((Set-VMNetworkAdapterVlan -VMName $VMName -VlanId $VLAN -VMNetworkAdapterName "Network Adapter" -Access -Passthr) | Out-String) -LogFile $logfile
    } catch {
        log "Error configuring vNIC for VM '$VMName'.. stopping" Magenta $logfile
        Log "Error details: $error" Magenta $logfile
        break
    }
    # Create and attach additional dynamic disk drives:
    For ($i=1; $i -lt $AdditionalDisksTotal+1; $i++) {
        $AdditionalDrive = $VMFolder + "\" + $VMName + "-Additional-" + $i + ".VHDX"
        log "Creating new dynamic disk '$AdditionalDrive'" green $logfile
        try {
            log ((New-VHD –Path $AdditionalDrive –BlockSizeBytes 128MB –LogicalSectorSize 4KB –SizeBytes $AdditionalDisksSize) | Out-String) -LogFile $logfile
        } catch {
            log "Error creating new dynamic disk '$AdditionalDrive'..stopping" Magenta $logfile
            Log "Error details: $error" Magenta $logfile
            break
        }
        log "Attaching new disk '$AdditionalDrive' to VM '$VMName'" green
        try {
            log ((Add-VMHardDiskDrive -VMName $VMName -ControllerType SCSI -Path $AdditionalDrive -Passthru) | Out-String) -LogFile $logfile
        } catch {
            log "Error attaching new disk '$AdditionalDrive' to VM '$VMName'.. stopping" Magenta $logfile
            Log "Error details: $error" Magenta $logfile
            break
        }
    }
}

function Del-EmptyColumns {
<# 
 .Synopsis
  Function to delete empty columns from CSV files

 .Description
  Function makes a backup of the CSV file in a folder ".\Del-EmptyColumns" before updating it.
  Function leaves a log file of its progress in the current folder

 .Parameter FileName
  Name of the CSV file

 .Parameter Delimiter
  comma (,) or semicolon (;) are accepted

 .Parameter BackupFolder
  Name of folder to backup files before updating them. Function will create it if it does not exist

 .Example
  Del-EmptyColumns test1.csv
  This example deletes empty columns in test1.csv file in the current folder

 .Example
  Del-EmptyColumns (Get-ChildItem D:\Sandbox\csv\* -include *.csv)
  This example deletes empty columns in the CSV files in the d:\Sandbox\csv folder

 .Example
    $CSVFiles = Get-ChildItem C:\scripts\Results\csv\* -include *.csv
    for ($i = 0; $i -lt $CSVFiles.Count; $i++) {
        $Columns = (import-csv $CSVFiles[$i][0] | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue).Count
        if ($Columns -gt 2) {
            Del-EmptyColumns $CSVFiles[$i]
        } else {
            log "File $($CSVFiles[$i]) has only $Columns columns, skipping.." Yellow
        }
    } 
  This example deletes empty columns in the CSV files in the C:\scripts\Results\csv
  folder, skipping those with 2 or fewer columns
 
 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/14/2014

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [ValidateScript({ Test-Path $_ })]
            [String[]]$FileName,
        [Parameter(Mandatory=$false,
                   Position=1)]
            [ValidateSet(",",";")]
            [String]$Delimiter = ",",
        [Parameter(Mandatory=$false,
                   Position=2)]
            [String]$BackupFolder = ".\Del-EmptyColumns"
    )

    $Logfile = ".\Del-EmptyColumns_$(Get-Date -format yyyyMMdd_hhmmsstt).txt"
    if (-not (Test-Path -Path $BackupFolder)) { New-Item -Path $BackupFolder -ItemType Directory -Force }
    foreach ($File in $FileName) {
        [array]$DeleteMe = $null # Columns to be deleted
        $Error.Clear()
        try {
            $csvFile = Import-Csv -Path $File -Delimiter $Delimiter 
            if ($csvFile) { 
                $ColumnCount = ($csvFile | Get-Member | Where-Object { $_.MemberType -match "Property" }).Count
                $Columns = ($csvFile | Get-Member | Where-Object { $_.MemberType -match "Property" } | Select-Object "Name")
                log "Filename: $File, Delimiter: ($Delimiter), Columns: $ColumnCount, Rows: $($csvFile.Count)" Green $Logfile
                # 
                # Identify empty columns 
                for ($Column=0; $Column -lt $ColumnCount; $Column++) {
                    Write-Verbose "Column # $Column, Header: $($Columns.name[$Column])"
                    $IsEmpty = $true
                    $Row = 0
                    Do { # using Do loop instead of a For loop to be able to break out easily on the first non-empty cell in this column
                        $Cell = $csvFile[$Row]."$($Columns.name[$Column])"
                        Write-Verbose "          - Row # $Row - Content: ($Cell)"
                        if ($Cell.Length -gt 0) { $IsEmpty = $false }                
                        $Row++
                    } while ($Row -lt $csvFile.Count -and $IsEmpty)
                    if ($IsEmpty) { $DeleteMe += $Columns.name[$Column] }
                }
                Write-Verbose "Empty columns: $DeleteMe"
                #
                # Backup original file
                $BackupFile = "$BackupFolder\$(Get-Date -format yyyyMMdd_hhmmsstt)_$($file.split("\")[$file.split("\").count-1])"
                Copy-Item -Path $File -Destination $BackupFile -Force
                if (-not (Test-Path -Path $BackupFile)) {
                    Write-Output "Unable to backup the file $File to $BackupFile, stopping.."
                    break
                }
                #
                # Delete empty columns
                foreach ($DeleteColumn in $DeleteMe) {
                    Write-Verbose "Deleting column: $DeleteColumn"
                    $csvFile = $csvFile | select * -exclude $DeleteColumn # $csvFile.PSObject.Properties.Remove($DeleteColumn) # does not work!!
                }
                $csvFile | Export-Csv -Path $File -Force -NoTypeInformation
                log "Deleted $($DeleteMe.Count) empty columns from file $File" Green $Logfile
            } else {
                log "File $File is empty" yellow $Logfile
            }
        } catch {
            Write-Output "Errors encountered"
            Write-Output "Error Details: $Error"
        }
    }
}

function Get-FilesContainingText {
<# 
 .Synopsis
  Function to get a list of iles containg certain text in a given set of 
  folders and their subfolders

 .Description
  Function returns a list of file names, each containing the search text

 .Parameter SearchString
  The text string to search for

 .Parameter FolderName
  Name of the folder(s) to search in.
  Function searches in each folder and its subfolders

 .Example
  Get-FilesContainingText "import"
  This example lists all the files in the current folder and 
  its subfolders that contain the string "import"

 .Example
  Get-FilesContainingText -SearchString "cheese" -FolderName "d:\Sandbox","\\MyServer1\install" -Verbose
  This example lists all the files in the folders"d:\Sandbox" and 
  "\\MyServer1\install" and their subfolders that contain the string "cheese"

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/16/2014

#>
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [Alias ("Text")]
            [String]$SearchString,
        [Parameter(Mandatory=$false,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=1)]
            [ValidateScript({ Test-Path $_ })]
            [Alias ("Directory")]
            [String[]]$FolderName = ".\" 
    )
    $FoundCount = 0
    $FileCount = 0
    $FileList = @()
    $Time = Measure-Command { 
        foreach ($Folder in $FolderName) {
            $Files = Get-ChildItem -Path $Folder -File -Recurse
            foreach ($File in $Files) {
                $FileCount++
                $FilePath = Join-Path $File.DirectoryName -ChildPath $File.Name
                $FileContent = Get-Content -Path $FilePath -Raw
                if ($FileContent.Contains($SearchString)) { $FileList += $FilePath; $FoundCount++ }
            }
        }
    }
    Write-Verbose "Searched $FileCount files in $($Time.Seconds) second(s)"
    for ($i=0; $i -lt $FileList.Count; $i++) { Write-Output "$($FileList[$i])" }
    Write-Verbose "Found $FoundCount files containing the search string '$SearchString' in the folder(s) '$FolderName'"
}

function Import-SessionCommands {
<# 
 .Synopsis
  Function to import commands from another computer

 .Description
  Function will import commands from remote computer from the module(s) listed.

 .Parameter ModuleName
  Name(s) of the module(s) that we want to import their commands into the current
  PS console. 
  Note that session commands will not be available in other PS instances.

 .Parameter ComputerName
  Computer name that has the module(s) that we need to import their commands.

 .Parameter Keep
  This is a switch. When selected, the function will export the imported module(s)
  locally under "C:\Program Files\WindowsPowerShell\Modules" if it's in the PSModulePath,
  otherwise, it will export it to the default path "$home\Documents\WindowsPowerShell\Modules"
  - Note 1: Exported modules and their commands can be used directly from any PS instance 
            after a module has been exported with the -keep switch
  - Note 2: Even though a module has been exported locally, everytime you try to use one of 
            its commands, PS will start an implicit remoting session to the server where the 
            module was imported from.

 .Example
  Import-SessionCommands -ModuleName ActiveDirectory -ComputerName DC01
  This example imports all the commands from the ActiveDirectory module from the DC01 server
  So, in this PS console instance we can use AD commands like Get-ADComputer without the need
  to install AD features, tools, or PS modules on this computer!

 .Example
  Import-SessionCommands SQLPS,Storage V-2012R2-SQL1 -Verbose
  This example imports all the commands from the PSSQL and Storage modules from the MySQLServer
  server into the current PS instance

 .Example 
  Import-SessionCommands WebAdministration,BestPractices,MMAgent CM01 -keep
  This example imports all the commands from the WebAdministration, BestPractices, and MMAgent 
  modules from the CM01 server into the current PS instance, and exports them locally.

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  Requires PS 3.0
  v1.0 - 08/17/2014
    Although we need to eventually run:
    Remove-PSSession -Session $Session
    We cannot do that in the function as we'll lose the imported session commands
    Two things to consider:
    1. The session will be automatically removed when the PS console is closed
    2. If in the parent script that's using this function a blanket Remove-PSSession 
    command is run, like:
    Get-PSSession | Remove-PSSession
    We'll lose this session and its commands, which could cripple the parent script
#>
    
    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [String[]]$ModuleName, 
        [Parameter(Mandatory=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=1)]
            [String]$ComputerName,
        [Parameter(Mandatory=$false,
                   Position=2)]
            [Switch]$Keep
    )
    
    # Get a random session name:
    Do { $SessionName = "Import" + (Get-Random -Minimum 10000000 -Maximum 99999999) }
        While (Get-PSSession -Name $SessionName -ErrorAction SilentlyContinue) 
    Write-Verbose "New PSSession name: $SessionName"
    if ($Env:PSModulePath.Split(";") -contains "C:\Program Files\WindowsPowerShell\Modules") {
        $ExportTo = "C:\Program Files\WindowsPowerShell\Modules"
    } else {
        $ExportTo = "$home\Documents\WindowsPowerShell\Modules"
    }
    try { 
        Write-Output "Connecting to computer $ComputerName.."
        $CurrentSessions = Get-PSSession -ErrorAction SilentlyContinue -ComputerName $ComputerName
        if ($CurrentSessions.ComputerName -Contains $ComputerName) {
            $Session = $CurrentSessions[0]
        } else {
            $Session = New-PSSession -ComputerName $ComputerName -Name $SessionName -ErrorAction Stop
        }
        Write-Verbose "Current PSSessions: $(Get-PSSession)"
        $RemoteModules = Invoke-Command -ScriptBlock { 
            Get-Module -ListAvailable  | Select-Object "Name" } -Session $Session 
        $LocalModules = Get-Module -ListAvailable  | Select-Object "Name"
        foreach ($Module in $ModuleName) {
            if ($LocalModules.Name -Contains $Module -or $LocalModules.Name -Contains "Imported-$Module") {
                Write-Output "Module $Module exists locally, not importing.."
            } else {
                if ($RemoteModules.Name -Contains $Module) {
                    Write-Output "Found module $Module on server $ComputerName, importing its commands.."
                    Invoke-Command -ScriptBlock { Param($Module)
                        Import-Module $Module } -Session $Session -ArgumentList $Module
                    try { 
                        $ImportedModule = Import-PSSession -Session $Session -Module $Module -DisableNameChecking -ErrorAction Stop 
                        if ($Keep) {
                            Write-Output "Keeping module $Module locally.."
                            Remove-Module -Name $ImportedModule.Name
                            Export-PSSession -Module $Module -OutputModule "$ExportTo\Imported-$Module" -Session $Session -Force
                            Import-Module -Name "Imported-$Module"
                        }
                    } catch { 
                        Write-Output "Module $Module already imported, skipping.." 
                    }
                } else {
                    Write-Output "Error: module $Module not found on server $ComputerName"
                }
            }
        }
    } catch {
        Write-Output "Error: unable to connect to server $ComputerName"
        Write-Output "       Check if $ComputerName exists, is online, "
        Write-Output "       has WinRM enabled and configured, and "
        Write-Output "       you have sufficient permissions on it"
    }
}

function Remove-ArrayElement {
<# 
 .Synopsis
  Function to remove element(s) from array

 .Description
  This function will remove a given element from the input array

 .Parameter Array
  The array we want to remove element(s) from

 .Parameter Element
  The element(s) we want to remove from the input array

 .Example
  $a = @(1,2,33,44,55,66,77,88,99)
  $a = Remove-ArrayElement -Array $a -Element 44
  Removes element "44" from the $a array

 .Example
  $a = @(1,2,33,44,55,66,77,88,99)
  $a = Remove-ArrayElement $a 44,77
  Removes elements "44" and "77" from the $a array

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Function by Sam Boutros
  v1.0 - 08/23/2014

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   Position=0)]
            [array]$Array, 
        [Parameter(Mandatory=$true,
                   Position=1)]
            [String[]]$Element 
    )

    foreach ($sElement in $Element) {
        $indexToRemove = $null
        for ($i=0; $i -lt $Array.Count;$i++) { if ($Array[$i] -match $sElement) { $indexToRemove = $i } }
        if ($indexToRemove) {
            $Array = $Array[0..($indexToRemove-1) + ($indexToRemove+1)..($Array.length - 1)] 
        } else {
            Write-Host "Element $sElement not found in array $Array" -ForegroundColor Yellow
        }
    }
    return $Array
}

function ConvertTo-EnhancedHTML {
<#
.SYNOPSIS
Provides an enhanced version of the ConvertTo-HTML command that includes
inserting an embedded CSS style sheet, JQuery, and JQuery Data Tables for
interactivity. Intended to be used with HTML fragments that are produced
by ConvertTo-EnhancedHTMLFragment. This command does not accept pipeline
input.

.PARAMETER jQueryURI
A Uniform Resource Indicator (URI) pointing to the location of the 
jQuery script file. You can download jQuery from www.jquery.com; you should
host the script file on a local intranet Web server and provide a URI
that starts with http:// or https://. Alternately, you can also provide
a file system path to the script file, although this may create security
issues for the Web browser in some configurations.

Tested with v1.8.2.

Defaults to http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.2.min.js, which
will pull the file from Microsoft's ASP.NET Content Delivery Network.

.PARAMETER jQueryDataTableURI
A Uniform Resource Indicator (URI) pointing to the location of the 
jQuery Data Table script file. You can download this from www.datatables.net;
you should host the script file on a local intranet Web server and provide a URI
that starts with http:// or https://. Alternately, you can also provide
a file system path to the script file, although this may create security
issues for the Web browser in some configurations.

Tested with jQuery DataTable v1.9.4

Defaults to http://ajax.aspnetcdn.com/ajax/jquery.dataTables/1.9.3/jquery.dataTables.min.js,
which will pull the file from Microsoft's ASP.NET Content Delivery Network.

.PARAMETER CssStyleSheet
The CSS style sheet content - not a file name. If you have a CSS file,
you can load it into this parameter as follows:

    -CSSStyleSheet (Get-Content MyCSSFile.css)

Alternately, you may link to a Web server-hosted CSS file by using the
-CssUri parameter.

.PARAMETER CssUri
A Uniform Resource Indicator (URI) to a Web server-hosted CSS file.
Must start with either http:// or https://. If you omit this, you
can still provide an embedded style sheet, which makes the resulting
HTML page more standalone. To provide an embedded style sheet, use
the -CSSStyleSheet parameter.

.PARAMETER Title
A plain-text title that will be displayed in the Web browser's window
title bar. Note that not all browsers will display this.

.PARAMETER PreContent
Raw HTML to insert before all HTML fragments. Use this to specify a main
title for the report:

    -PreContent "<H1>My HTML Report</H1>"

.PARAMETER PostContent
Raw HTML to insert after all HTML fragments. Use this to specify a 
report footer:

    -PostContent "Created on $(Get-Date)"

.PARAMETER HTMLFragments
One or more HTML fragments, as produced by ConvertTo-EnhancedHTMLFragment.

    -HTMLFragments $part1,$part2,$part3
.EXAMPLE
The following is a complete example script showing how to use
ConvertTo-EnhancedHTMLFragment and ConvertTo-EnhancedHTML. The
example queries 6 pieces of information from the local computer
and produces a report in C:\. This example uses most of the
avaiable options. It relies on Internet connectivity to retrieve
JavaScript from Microsoft's Content Delivery Network. This 
example uses an embedded stylesheet, which is defined as a here-string
at the top of the script.

$computername = 'localhost'
$path = 'c:\'
$style = @"
<style>
body {
    color:#333333;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
}
h1 {
    text-align:center;
}
h2 {
    border-top:1px solid #666666;
}


th {
    font-weight:bold;
    color:#eeeeee;
    background-color:#333333;
    cursor:pointer;
}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
.paginate_enabled_next, .paginate_enabled_previous {
    cursor:pointer; 
    border:1px solid #222222; 
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.paginate_disabled_previous, .paginate_disabled_next {
    color:#666666; 
    cursor:pointer;
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.dataTables_info { margin-bottom:4px; }
.sectionheader { cursor:pointer; }
.sectionheader:hover { color:red; }
.grid { width:100% }
.red {
    color:red;
    font-weight:bold;
} 
</style>
"@

function Get-InfoOS {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $os = Get-WmiObject -class Win32_OperatingSystem -ComputerName $ComputerName
    $props = @{'OSVersion'=$os.version;
               'SPVersion'=$os.servicepackmajorversion;
               'OSBuild'=$os.buildnumber}
    New-Object -TypeName PSObject -Property $props
}

function Get-InfoCompSystem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $cs = Get-WmiObject -class Win32_ComputerSystem -ComputerName $ComputerName
    $props = @{'Model'=$cs.model;
               'Manufacturer'=$cs.manufacturer;
               'RAM (GB)'="{0:N2}" -f ($cs.totalphysicalmemory / 1GB);
               'Sockets'=$cs.numberofprocessors;
               'Cores'=$cs.numberoflogicalprocessors}
    New-Object -TypeName PSObject -Property $props
}

function Get-InfoBadService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $svcs = Get-WmiObject -class Win32_Service -ComputerName $ComputerName `
           -Filter "StartMode='Auto' AND State<>'Running'"
    foreach ($svc in $svcs) {
        $props = @{'ServiceName'=$svc.name;
                   'LogonAccount'=$svc.startname;
                   'DisplayName'=$svc.displayname}
        New-Object -TypeName PSObject -Property $props
    }
}

function Get-InfoProc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $procs = Get-WmiObject -class Win32_Process -ComputerName $ComputerName
    foreach ($proc in $procs) { 
        $props = @{'ProcName'=$proc.name;
                   'Executable'=$proc.ExecutablePath}
        New-Object -TypeName PSObject -Property $props
    }
}

function Get-InfoNIC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $nics = Get-WmiObject -class Win32_NetworkAdapter -ComputerName $ComputerName `
           -Filter "PhysicalAdapter=True"
    foreach ($nic in $nics) {      
        $props = @{'NICName'=$nic.servicename;
                   'Speed'=$nic.speed / 1MB -as [int];
                   'Manufacturer'=$nic.manufacturer;
                   'MACAddress'=$nic.macaddress}
        New-Object -TypeName PSObject -Property $props
    }
}

function Get-InfoDisk {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $drives = Get-WmiObject -class Win32_LogicalDisk -ComputerName $ComputerName `
           -Filter "DriveType=3"
    foreach ($drive in $drives) {      
        $props = @{'Drive'=$drive.DeviceID;
                   'Size'=$drive.size / 1GB -as [int];
                   'Free'="{0:N2}" -f ($drive.freespace / 1GB);
                   'FreePct'=$drive.freespace / $drive.size * 100 -as [int]}
        New-Object -TypeName PSObject -Property $props 
    }
}

foreach ($computer in $computername) {
    try {
        $everything_ok = $true
        Write-Verbose "Checking connectivity to $computer"
        Get-WmiObject -class Win32_BIOS -ComputerName $Computer -EA Stop | Out-Null
    } catch {
        Write-Warning "$computer failed"
        $everything_ok = $false
    }

    if ($everything_ok) {
        $filepath = Join-Path -Path $Path -ChildPath "$computer.html"

        $params = @{'As'='List';
                    'PreContent'='<h2>OS</h2>'}
        $html_os = Get-InfoOS -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 

        $params = @{'As'='List';
                    'PreContent'='<h2>Computer System</h2>'}
        $html_cs = Get-InfoCompSystem -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 

        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; Local Disks</h2>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';
                    'Properties'='Drive',
                                 @{n='Size(GB)';e={$_.Size}},
                                 @{n='Free(GB)';e={$_.Free};css={if ($_.FreePct -lt 80) { 'red' }}},
                                 @{n='Free(%)';e={$_.FreePct};css={if ($_.FreeePct -lt 80) { 'red' }}}}
        $html_dr = Get-InfoDisk -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params

        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; Processes</h2>';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid'}
        $html_pr = Get-InfoProc -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 

        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; Services to Check</h2>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeHiddenSection'=$true;
                    'TableCssClass'='grid'}
        $html_sv = Get-InfoBadService -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 

        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; NICs</h2>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeHiddenSection'=$true;
                    'TableCssClass'='grid'}
        $html_na = Get-InfoNIC -ComputerName $Computer |
                   ConvertTo-EnhancedHTMLFragment @params

        $params = @{'CssStyleSheet'=$style;
                    'Title'="System Report for $computer";
                    'PreContent'="<h1>System Report for $computer</h1>";
                    'HTMLFragments'=@($html_os,$html_cs,$html_dr,$html_pr,$html_sv,$html_na)}
        ConvertTo-EnhancedHTML @params |
        Out-File -FilePath $filepath
    }
}

 .Notes
  Function by Don Jones
  Generated on: 9/10/2013
  For more information see Powershell.org
  included in SBTools module with permission by Don Jones
  
#>
    [CmdletBinding()]
    param(
        [string]$jQueryURI = 'http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.2.min.js',
        [string]$jQueryDataTableURI = 'http://ajax.aspnetcdn.com/ajax/jquery.dataTables/1.9.3/jquery.dataTables.min.js',
        [Parameter(ParameterSetName='CSSContent')][string[]]$CssStyleSheet,
        [Parameter(ParameterSetName='CSSURI')][string[]]$CssUri,
        [string]$Title = 'Report',
        [string]$PreContent,
        [string]$PostContent,
        [Parameter(Mandatory=$True)][string[]]$HTMLFragments
    )


    <#
        Add CSS style sheet. If provided in -CssUri, add a <link> element.
        If provided in -CssStyleSheet, embed in the <head> section.
        Note that BOTH may be supplied - this is legitimate in HTML.
    #>
    Write-Verbose "Making CSS style sheet"
    $stylesheet = ""
    if ($PSBoundParameters.ContainsKey('CssUri')) {
        $stylesheet = "<link rel=`"stylesheet`" href=`"$CssUri`" type=`"text/css`" />"
    }
    if ($PSBoundParameters.ContainsKey('CssStyleSheet')) {
        $stylesheet = "<style>$CssStyleSheet</style>" | Out-String
    }


    <#
        Create the HTML tags for the page title, and for
        our main javascripts.
    #>
    Write-Verbose "Creating <TITLE> and <SCRIPT> tags"
    $titletag = ""
    if ($PSBoundParameters.ContainsKey('title')) {
        $titletag = "<title>$title</title>"
    }
    $script += "<script type=`"text/javascript`" src=`"$jQueryURI`"></script>`n<script type=`"text/javascript`" src=`"$jQueryDataTableURI`"></script>"


    <#
        Render supplied HTML fragments as one giant string
    #>
    Write-Verbose "Combining HTML fragments"
    $body = $HTMLFragments | Out-String


    <#
        If supplied, add pre- and post-content strings
    #>
    Write-Verbose "Adding Pre and Post content"
    if ($PSBoundParameters.ContainsKey('precontent')) {
        $body = "$PreContent`n$body"
    }
    if ($PSBoundParameters.ContainsKey('postcontent')) {
        $body = "$body`n$PostContent"
    }


    <#
        Add a final script that calls the datatable code
        We dynamic-ize all tables with the .enhancedhtml-dynamic-table
        class, which is added by ConvertTo-EnhancedHTMLFragment.
    #>
    Write-Verbose "Adding interactivity calls"
    $datatable = ""
    $datatable = "<script type=`"text/javascript`">"
    $datatable += '$(document).ready(function () {'
    $datatable += "`$('.enhancedhtml-dynamic-table').dataTable();"
    $datatable += '} );'
    $datatable += "</script>"


    <#
        Datatables expect a <thead> section containing the
        table header row; ConvertTo-HTML doesn't produce that
        so we have to fix it.
    #>
    Write-Verbose "Fixing table HTML"
    $body = $body -replace '<tr><th>','<thead><tr><th>'
    $body = $body -replace '</th></tr>','</th></tr></thead>'


    <#
        Produce the final HTML. We've more or less hand-made
        the <head> amd <body> sections, but we let ConvertTo-HTML
        produce the other bits of the page.
    #>
    Write-Verbose "Producing final HTML"
    ConvertTo-HTML -Head "$stylesheet`n$titletag`n$script`n$datatable" -Body $body  
    Write-Debug "Finished producing final HTML"


}

function ConvertTo-EnhancedHTMLFragment {
<#
.SYNOPSIS
Creates an HTML fragment (much like ConvertTo-HTML with the -Fragment switch
that includes CSS class names for table rows, CSS class and ID names for the
table, and wraps the table in a <DIV> tag that has a CSS class and ID name.

.PARAMETER InputObject
The object to be converted to HTML. You cannot select properties using this
command; precede this command with Select-Object if you need a subset of
the objects' properties.

.PARAMETER EvenRowCssClass
The CSS class name applied to even-numbered <TR> tags. Optional, but if you
use it you must also include -OddRowCssClass.

.PARAMETER OddRowCssClass
The CSS class name applied to odd-numbered <TR> tags. Optional, but if you 
use it you must also include -EvenRowCssClass.

.PARAMETER TableCssID
Optional. The CSS ID name applied to the <TABLE> tag.

.PARAMETER DivCssID
Optional. The CSS ID name applied to the <DIV> tag which is wrapped around the table.

.PARAMETER TableCssClass
Optional. The CSS class name to apply to the <TABLE> tag.

.PARAMETER DivCssClass
Optional. The CSS class name to apply to the wrapping <DIV> tag.

.PARAMETER As
Must be 'List' or 'Table.' Defaults to Table. Actually produces an HTML
table either way; with Table the output is a grid-like display. With
List the output is a two-column table with properties in the left column
and values in the right column.

.PARAMETER Properties
A comma-separated list of properties to include in the HTML fragment.
This can be * (which is the default) to include all properties of the
piped-in object(s). In addition to property names, you can also use a
hashtable similar to that used with Select-Object. For example:

 Get-Process | ConvertTo-EnhancedHTMLFragment -As Table `
               -Properties Name,ID,@{n='VM';
                                     e={$_.VM};
                                     css={if ($_.VM -gt 100) { 'red' }
                                          else { 'green' }}}

This will create table cell rows with the calculated CSS class names.
E.g., for a process with a VM greater than 100, you'd get:

  <TD class="red">475858</TD>
  
You can use this feature to specify a CSS class for each table cell
based upon the contents of that cell. Valid keys in the hashtable are:

  n, name, l, or label: The table column header
  e or expression: The table cell contents
  css or csslcass: The CSS class name to apply to the <TD> tag 
  
Another example:

  @{n='Free(MB)';
    e={$_.FreeSpace / 1MB -as [int]};
    css={ if ($_.FreeSpace -lt 100) { 'red' } else { 'blue' }}
    
This example creates a column titled "Free(MB)". It will contain
the input object's FreeSpace property, divided by 1MB and cast
as a whole number (integer). If the value is less than 100, the
table cell will be given the CSS class "red." If not, the table
cell will be given the CSS class "blue." The supplied cascading
style sheet must define ".red" and ".blue" for those to have any
effect.  

.PARAMETER PreContent
Raw HTML content to be placed before the wrapping <DIV> tag. 
For example:

    -PreContent "<h2>Section A</h2>"

.PARAMETER PostContent
Raw HTML content to be placed after the wrapping <DIV> tag.
For example:

    -PostContent "<hr />"

.PARAMETER MakeHiddenSection
Used in conjunction with -PreContent. Adding this switch, which
needs no value, turns your -PreContent into  clickable report
section header. The section will be hidden by default, and clicking
the header will toggle its visibility.

When using this parameter, consider adding a symbol to your -PreContent
that helps indicate this is an expandable section. For example:

    -PreContent '<h2>&diams; My Section</h2>'

If you use -MakeHiddenSection, you MUST provide -PreContent also, or
the hidden section will not have a section header and will not be
visible.

.PARAMETER MakeTableDynamic
When using "-As Table", makes the table dynamic. Will be ignored
if you use "-As List". Dynamic tables are sortable, searchable, and
are paginated.

You should not use even/odd styling with tables that are made
dynamic. Dynamic tables automatically have their own even/odd
styling. You can apply CSS classes named ".odd" and ".even" in 
your CSS to style the even/odd in a dynamic table.

.EXAMPLE
 $fragment = Get-WmiObject -Class Win32_LogicalDisk |
             Select-Object -Property PSComputerName,DeviceID,FreeSpace,Size |
             ConvertTo-HTMLFragment -EvenRowClass 'even' `
                                    -OddRowClass 'odd' `
                                    -PreContent '<h2>Disk Report</h2>' `
                                    -MakeHiddenSection `
                                    -MakeTableDynamic

 You will usually save fragments to a variable, so that multiple fragments
 (each in its own variable) can be passed to ConvertTo-EnhancedHTML.
.NOTES
Consider adding the following to your CSS when using dynamic tables:

    .paginate_enabled_next, .paginate_enabled_previous {
        cursor:pointer; 
        border:1px solid #222222; 
        background-color:#dddddd; 
        padding:2px; 
        margin:4px;
        border-radius:2px;
    }
    .paginate_disabled_previous, .paginate_disabled_next {
        color:#666666; 
        cursor:pointer;
        background-color:#dddddd; 
        padding:2px; 
        margin:4px;
        border-radius:2px;
    }
    .dataTables_info { margin-bottom:4px; }

This applies appropriate coloring to the next/previous buttons,
and applies a small amount of space after the dynamic table.

If you choose to make sections hidden (meaning they can be shown
and hidden by clicking on the section header), consider adding
the following to your CSS:

    .sectionheader { cursor:pointer; }
    .sectionheader:hover { color:red; }

This will apply a hover-over color, and change the cursor icon,
to help visually indicate that the section can be toggled.

 .Notes
  Function by Don Jones
  Generated on: 9/10/2013
  For more information see Powershell.org
  included in SBTools module with permission by Don Jones

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [object[]]$InputObject,
        [string]$EvenRowCssClass,
        [string]$OddRowCssClass,
        [string]$TableCssID,
        [string]$DivCssID,
        [string]$DivCssClass,
        [string]$TableCssClass,
        [ValidateSet('List','Table')]
        [string]$As = 'Table',
        [object[]]$Properties = '*',
        [string]$PreContent,
        [switch]$MakeHiddenSection,
        [switch]$MakeTableDynamic,
        [string]$PostContent
    )
    BEGIN {
        <#
            Accumulate output in a variable so that we don't
            produce an array of strings to the pipeline, but
            instead produce a single string.
        #>
        $out = ''
        <#
            Add the section header (pre-content). If asked to
            make this section of the report hidden, set the
            appropriate code on the section header to toggle
            the underlying table. Note that we generate a GUID
            to use as an additional ID on the <div>, so that
            we can uniquely refer to it without relying on the
            user supplying us with a unique ID.
        #>
        Write-Verbose "Precontent"
        if ($PSBoundParameters.ContainsKey('PreContent')) {
            if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
               [string]$tempid = [System.Guid]::NewGuid()
               $out += "<span class=`"sectionheader`" onclick=`"`$('#$tempid').toggle(500);`">$PreContent</span>`n"
            } else {
                $out += $PreContent
                $tempid = ''
            }
        }
        <#
            The table will be wrapped in a <div> tag for styling
            purposes. Note that THIS, not the table per se, is what
            we hide for -MakeHiddenSection. So we will hide the section
            if asked to do so.
        #>
        Write-Verbose "DIV"
        if ($PSBoundParameters.ContainsKey('DivCSSClass')) {
            $temp = " class=`"$DivCSSClass`""
        } else {
            $temp = ""
        }
        if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
            $temp += " id=`"$tempid`" style=`"display:none;`""
        } else {
            $tempid = ''
        }
        if ($PSBoundParameters.ContainsKey('DivCSSID')) {
            $temp += " id=`"$DivCSSID`""
        }
        $out += "<div $temp>"
        <#
            Create the table header. If asked to make the table dynamic,
            we add the CSS style that ConvertTo-EnhancedHTML will look for
            to dynamic-ize tables.
        #>
        Write-Verbose "TABLE"
        $_TableCssClass = ''
        if ($PSBoundParameters.ContainsKey('MakeTableDynamic') -and $As -eq 'Table') {
            $_TableCssClass += 'enhancedhtml-dynamic-table '
        }
        if ($PSBoundParameters.ContainsKey('TableCssClass')) {
            $_TableCssClass += $TableCssClass
        }
        if ($_TableCssClass -ne '') {
            $css = "class=`"$_TableCSSClass`""
        } else {
            $css = ""
        }
        if ($PSBoundParameters.ContainsKey('TableCSSID')) {
            $css += "id=`"$TableCSSID`""
        } else {
            if ($tempid -ne '') {
                $css += "id=`"$tempid`""
            }
        }
        $out += "<table $css>"
        <#
            We're now setting up to run through our input objects
            and create the table rows
        #>
        $fragment = ''
        $wrote_first_line = $false
        $even_row = $false

        if ($properties -eq '*') {
            $all_properties = $true
        } else {
            $all_properties = $false
        }

    }
    PROCESS {

        foreach ($object in $inputobject) {
            Write-Verbose "Processing object"
            $datarow = ''
            $headerrow = ''

            <#
                Apply even/odd row class. Note that this will mess up the output
                if the table is made dynamic. That's noted in the help.
            #>
            if ($PSBoundParameters.ContainsKey('EvenRowCSSClass') -and $PSBoundParameters.ContainsKey('OddRowCssClass')) {
                if ($even_row) {
                    $row_css = $OddRowCSSClass
                    $even_row = $false
                    Write-Verbose "Even row"
                } else {
                    $row_css = $EvenRowCSSClass
                    $even_row = $true
                    Write-Verbose "Odd row"
                }
            } else {
                $row_css = ''
                Write-Verbose "No row CSS class"
            }


            <#
                If asked to include all object properties, get them.
            #>
            if ($all_properties) {
                $properties = $object | Get-Member -MemberType Properties | Select -ExpandProperty Name
            }


            <#
                We either have a list of all properties, or a hashtable of
                properties to play with. Process the list.
            #>
            foreach ($prop in $properties) {
                Write-Verbose "Processing property"
                $name = $null
                $value = $null
                $cell_css = ''


                <#
                    $prop is a simple string if we are doing "all properties,"
                    otherwise it is a hashtable. If it's a string, then we
                    can easily get the name (it's the string) and the value.
                #>
                if ($prop -is [string]) {
                    Write-Verbose "Property $prop"
                    $name = $Prop
                    $value = $object.($prop)
                } elseif ($prop -is [hashtable]) {
                    Write-Verbose "Property hashtable"
                    <#
                        For key "css" or "cssclass," execute the supplied script block.
                        It's expected to output a class name; we embed that in the "class"
                        attribute later.
                    #>
                    if ($prop.ContainsKey('cssclass')) { $cell_css = $Object | ForEach $prop['cssclass'] }
                    if ($prop.ContainsKey('css')) { $cell_css = $Object | ForEach $prop['css'] }


                    <#
                        Get the current property name.
                    #>
                    if ($prop.ContainsKey('n')) { $name = $prop['n'] }
                    if ($prop.ContainsKey('name')) { $name = $prop['name'] }
                    if ($prop.ContainsKey('label')) { $name = $prop['label'] }
                    if ($prop.ContainsKey('l')) { $name = $prop['l'] }


                    <#
                        Execute the "expression" or "e" key to get the value of the property.
                    #>
                    if ($prop.ContainsKey('e')) { $value = $Object | ForEach $prop['e'] }
                    if ($prop.ContainsKey('expression')) { $value = $tObject | ForEach $prop['expression'] }


                    <#
                        Make sure we have a name and a value at this point.
                    #>
                    if ($name -eq $null -or $value -eq $null) {
                        Write-Error "Hashtable missing Name and/or Expression key"
                    }
                } else {
                    <#
                        We got a property list that wasn't strings and
                        wasn't hashtables. Bad input.
                    #>
                    Write-Warning "Unhandled property $prop"
                }


                <#
                    When constructing a table, we have to remember the
                    property names so that we can build the table header.
                    In a list, it's easier - we output the property name
                    and the value at the same time, since they both live
                    on the same row of the output.
                #>
                if ($As -eq 'table') {
                    Write-Verbose "Adding $name to header and $value to row"
                    $headerrow += "<th>$name</th>"
                    $datarow += "<td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$value</td>"
                } else {
                    $wrote_first_line = $true
                    $headerrow = ""
                    $datarow = "<td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$name :</td><td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$value</td>"
                    $out += "<tr$(if ($row_css -ne '') { ' class="'+$row_css+'"' })>$datarow</tr>"
                }
            }


            <#
                Write the table header, if we're doing a table.
            #>
            if (-not $wrote_first_line -and $as -eq 'Table') {
                Write-Verbose "Writing header row"
                $out += "<tr>$headerrow</tr><tbody>"
                $wrote_first_line = $true
            }


            <#
                In table mode, write the data row.
            #>
            if ($as -eq 'table') {
                Write-Verbose "Writing data row"
                $out += "<tr$(if ($row_css -ne '') { ' class="'+$row_css+'"' })>$datarow</tr>"
            }
        }
    }
    END {
        <#
            Finally, post-content code, the end of the table,
            the end of the <div>, and write the final string.
        #>
        Write-Verbose "PostContent"
        if ($PSBoundParameters.ContainsKey('PostContent')) {
            $out += "`n$PostContent"
        }
        Write-Verbose "Done"
        $out += "</tbody></table></div>"
        Write-Output $out
    }
}

function Get-PII {
<# 
 .Synopsis
  Function to report on files containing Personally Identifiable Information (PII)

 .Description
  Function produces an HTML report of files of certain extensions in certain folders,
  that contain PII such as social security numbers and credit card card numbers.

 .Parameter FileType
  File extensions of the files to search in. Wildcards are allowed.
  Examples: "txt" or "txt","doc?","xls"

 .Parameter TargetFolder
  Folder(s) to search in. Wildcards are allowed.
  Examples: "c:\support" or "d:\sandbox","\\server11\share\sales*"

 .Parameter FileProperties
  These will make up the columns in the tabular HTML report. If not provided, the default is:
  "Name","DirectoryName","CreationTime","LastAccessTime","LastWriteTime"
  Acceptable values are:
  "Name","DirectoryName","Extension","Length","CreationTime","LastAccessTime","LastWriteTime","Attributes"

 .Parameter LogFile
  File name where log information is saved.
  Default is "Get-PII_20140822_081426AM.txt" (Date/time script was run)
  Examples: ".\log.txt" or "c:\scripts\pii.log"

 .Parameter HTMLFile
  HTML file name where report information is saved.
  Default is "Get-PII_20140822_081426AM.htm" (Date/time script was run)
  Examples: ".\Report.htm" or "c:\scripts\pii.html"

 .Example
  Get-PII -FileType "txt" -TargetFolder "D:\Sandbox"
  This example produces a report on files with "txt" extension in the folder
  "d:\sandbox" that contain PII

 .Example
  Get-PII "txt","csv","doc?" "D:\Sandbox","\\Myserver\Install\Script?"
  This example produces a report on files with the extensions "txt","csv","doc?"
  in the folders "D:\Sandbox","\\Myserver\Install\Script?" that contain PII

 .Example
  $FileExtensions = "txt","csv","xls?","ppt*"
  $SearchFolders = "D:\Sandbox","\\Server2\Share\User*"
  Get-PII $FileExtensions $SearchFolders

 .Example
  Get-PII -FileType "txt" -TargetFolder "D:\Sandbox" -FileProperties "Name","DirectoryName","Attributes"
  This example displays the listed file properties int he HTML report

 .Link
  https://superwidgets.wordpress.com/category/powershell/

 .Notes
  Script to report on files containing Personally Identifiable Information
  By Sam Boutros
  v1.0 - 8/20/2014
  Requires:
  Powershell 3.0
  EnhancedHTML2 from https://onedrive.live.com/?cid=7F868AA697B937FE&id=7F868AA697B937FE%21113
  SBToolsmodule from http://gallery.technet.microsoft.com/scriptcenter/SBTools-module-adds-38992422

#>

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')] 
    Param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeLine=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=0)]
            [ValidateLength(1,5)]
            [String[]]$FileType, 
        [Parameter(Mandatory=$true,
                   ValueFromPipeLineByPropertyName=$true,
                   Position=1)]
            [ValidateScript({ Test-Path $_ })]
            [String[]]$TargetFolder, 
        [Parameter(Mandatory=$false,
                   Position=2)]
            [ValidateSet("Name","DirectoryName","Extension","Length","CreationTime","LastAccessTime","LastWriteTime","Attributes")]
            [String[]]$FileProperties = ("Name","DirectoryName","CreationTime","LastAccessTime","LastWriteTime"), 
#            [ValidateSet("Name","DirectoryName","Extension","Length","CreationTime","LastAccessTime","LastWriteTime","Attributes","File")]
#            [String[]]$FileProperties = ("Name","DirectoryName","CreationTime","LastAccessTime","LastWriteTime","File"), 
        [Parameter(Mandatory=$false,
                   Position=3)]
            [String]$LogFile = ".\Get-PII_" + (Get-Date -format yyyyMMdd_hhmmsstt) + ".txt",
        [Parameter(Mandatory=$false,
                   Position=4)]
            [String]$HTMLFile = ".\Get-PII_" + (Get-Date -format yyyyMMdd_hhmmsstt) + ".htm"
    )

    $Patterns = @()
    $Pattern0 = 'Credit Card - Visa','(4\d{3}[-| ]\d{4}[-| ]\d{4}[-| ]\d{4})|(4\d{15})','Start with a 4 and have 16 digits, may be split as xxxx-xxxx-xxxx-xxxx by dashes or spaces'
    $Pattern1 = 'Credit Card - MasterCard','(5[1-5]\d{14})|(5[1-5]\d{2}[-| ]\d{4}[-| ]\d{4}[-| ]\d{4})','Starts with 51-55 and have 16 digits, may be split as xxxx-xxxx-xxxx-xxxx by dashes or spaces'
    $Pattern2 = 'Credit Card - Amex','(3[47]\d{13})|(3[47]\d{2}[-| ]\d{6}[-| ]\d{5})','Starts with 34 or 37 and have 15 digits, may be split as xxxx-xxxxxx-xxxxx by dashes or spaces'
    $Pattern3 = 'Credit Card - DinersClub','(3(?:0[0-5]|[68]\d)\d{11})|(3(?:0[0-5]|[68]\d)\d[-| ]\d{6}[-| ]\d{4})','Starts with 300-305, or 36-38 and have 14 digits, may be split as xxxx-xxxxxx-xxxx by dashes or spaces'
    $Pattern4 = 'Credit Card - Discover','(6(?:011|5\d{2})\d{12})|(6(?:011|5\d{2})[-| ]\d{4}[-| ]\d{4}[-| ]\d{4})','Start with 6011 or 65 and have 16 digits, may be split as xxxx-xxxx-xxxx-xxxx by dashes or spaces'
    $Pattern5 = 'Credit Card - JCB','((?:2131|1800|35\d{3})\d{11})|((?:2131|1800|35\d{2})[-| ]\d{4}[-| ]\d{4}[-| ]\d{3}[\d| ])','Start with 2131 or 1800 and have 15 digits) or (Start with 35 and have 16 digits'
    $Pattern6 = 'Social Security Number','(\d{3}[-| ]\d{2}[-| ]\d{4})|(\d{9})','9 digits, may be split as xxx-xx-xxxx by dashes or spaces'
    $Patterns = $Pattern0,$Pattern1,$Pattern2,$Pattern3,$Pattern4,$Pattern5,$Pattern6
    #    
    $Style = @"
    <style>
    body {
        color:#333333;
        font-family:Calibri,Tahoma;
        font-size: 10pt;
    }
    h1 {
        text-align:center;
    }
    h2 {
        border-top:1px solid #666666;
    }
    th {
        font-weight:bold;
        color:#eeeeee;
        background-color:#333333;
        cursor:pointer;
    }
    .odd  { background-color:#ffffff; }
    .even { background-color:#dddddd; }
    .paginate_enabled_next, .paginate_enabled_previous {
        cursor:pointer; 
        border:1px solid #222222; 
        background-color:#dddddd; 
        padding:2px; 
        margin:4px;
        border-radius:2px;
    }
    .paginate_disabled_previous, .paginate_disabled_next {
        color:#666666; 
        cursor:pointer;
        background-color:#dddddd; 
        padding:2px; 
        margin:4px;
        border-radius:2px;
    }
    .dataTables_info { margin-bottom:4px; }
    .sectionheader { cursor:pointer; }
    .sectionheader:hover { color:red; }
    .grid { width:100% }
    .red {
        color:red;
        font-weight:bold;
    } 
    </style>
"@
    #
    if (-not ("Name" -in $FileProperties)) { $FileProperties += "Name" }
    if (-not ("DirectoryName" -in $FileProperties)) { $FileProperties += "DirectoryName" }
    $Properties = $FileProperties
    For ($i=0; $i -lt $Patterns.Count; $i++) { $Properties += "Pattern$($i)" }
    $Tables = @()
    $TargetFolderString = $null
    foreach ($Folder in $TargetFolder) {
        $TargetFolderString += $Folder + ", "
        $FileTypeString = $null
        foreach ($Type in $FileType) {
            log "Searching for files with $Type extension on folder $Folder and its subfolders" Green $Logfile
            $FileTypeString += $Type + ", "
            $Files = Get-ChildItem -Path $Folder -Include *.$Type -Force -Recurse -ErrorAction SilentlyContinue
            $FileList = @()
            ForEach ($File in $Files) {
                $Props = @{}
                foreach ($Prop in $FileProperties) { $Props.Add($Prop, $($File.$Prop)) } 
                For ($i=0; $i -lt $Patterns.Count; $i++) { $Props.Add("Pattern$($i)", 0) }
                $objFile = New-Object -TypeName PSObject -Property $Props
#                $objFile.File = "<a href='$($objFile.DirectoryName)\$($objFile.Name)' target= '_blank'>$($objFile.Name)</a>"
                try { 
                    $FileContent = Get-Content -Path "$($objFile.DirectoryName)\$($objFile.Name)" -ErrorAction Stop
                    for ($j=0; $j -lt $FileContent.Count; $j++) { # each line in a file
                        For ($i=0; $i -lt $Patterns.Count; $i++) { 
                            if ($FileContent[$j] -match $Patterns[$i][1]) { 
                                $objFile.$("Pattern$i") += 1 
                                Write-Verbose "Found '$($Patterns[$i][0])' match in line '$($FileContent[$j])' in file '$($objFile.DirectoryName)\$($objFile.Name)'"
                            } 
                        }
                        Write-Debug "Finished checking line '$($FileContent[$j])' in file '$($objFile.DirectoryName)\$($objFile.Name)'"
                    }
                    $PatternMatch = $false
                    For ($k=0; $k -lt $Patterns.Count; $k++) {
                        if ($objFile.$("Pattern$k") -gt 0 ) { $PatternMatch = $true }
                    }
                    if ($PatternMatch) { $FileList += $objFile }
#                    if ($PatternMatch) { $FileList += ($objFile | Select-Object -Property * -ExcludeProperty Name) }
                } catch {
                    log "Unable to read file $($objFile.DirectoryName)\$($objFile.Name)" Yellow $LogFile
                }
            } 
            if ($FileList) { # Skip empty tables
                $Params = @{'As'='Table';
                        'PreContent'='<h2><font color=blue>''' + $FileList.Count + ''' </font>file(s) with <font color=blue>''' + $Type + '''</font> extension under folder <font color=blue>''' + $Folder + ''' </font>contain PII:</h2>';
                        'EvenRowCssClass'='even';
                        'OddRowCssClass'='odd';
                        'MakeTableDynamic'=$true;
                        'TableCssClass'='grid';
                        'Properties'= $Properties }
#                        'Properties'= ($Properties | Where-Object { $_ -ne "Name" }) }
                $Table = $FileList | ConvertTo-EnhancedHTMLFragment @params
                $Tables += $Table
                Write-Debug "Added HTML Fragment"
            }
        } 
    } 
    Write-Debug "Done all folders"
    $PreContent =  "<h1> Report of files containing PII </h1>"
    $PreContent += "<br> <u>Search crieteria:</u>"
    $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;Files with extension(s): <font color=blue>$($FileTypeString.Substring(0,$FileTypeString.Length-2))</font>"
    $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;In folder(s) - including their sub-folders: <font color=blue>$($TargetFolderString.Substring(0,$TargetFolderString.Length-2))</font>"
    $PreContent += "<br> <u>Patterns searched for:</u>"
    For ($i=0; $i -lt $Patterns.Count; $i++) { $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;Pattern$($i): $($Patterns[$i][0]) ($($Patterns[$i][2]))" }
    $PreContent += "<br> <u>Notes:</u>"
    $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;1. Social Security number pattern macthes may include false positives"
    $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;For example a 4234567890123456 number will match as a Visa card number <i><b>AND</b></i> a Social Security number"
    $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;2. Files with full name (including path) longer than 260 characters are not checked"
    $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;3. Files that cannot be accessed due to NTFS permissions are not checked"
    $PreContent += "<br> &nbsp;&nbsp;&nbsp;&nbsp;4. A pattern match is reported once per line per file (even if it occurs multiple times on the same line)"
    if ($Tables) {
        $Params = @{'CssStyleSheet'=$Style;
                'Title'="Report of files containing PII";
                'PreContent'=$PreContent;
                'HTMLFragments'=$Tables;
                'jQueryDataTableUri'='http://ajax.aspnetcdn.com/ajax/jquery.dataTables/1.9.3/jquery.dataTables.min.js';
                'jQueryUri'='http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.2.min.js'} 
        ConvertTo-EnhancedHTML @params | Out-File -FilePath $HTMLFile
        log "Done. Report saved to file $HTMLFile" Green $Logfile
        Invoke-Expression $HTMLFile
    } else {
        log "No files found with PII in files with extension(s): '$($FileTypeString.Substring(0,$FileTypeString.Length-2))' in folder(s) '$($TargetFolderString.Substring(0,$TargetFolderString.Length-2))'" green $LogFile
    }
    Write-Debug "Done"
}

Export-ModuleMember -Function Get-SBVHD, Log, New-SBSeed, Test-SBDisk, Get-SBRDPSession, 
    Test-SBVHDIntegrity, Get-SBIPInfo, SBBytes, New-SBVM, Del-EmptyColumns, 
    Get-FilesContainingText, Import-SessionCommands, Remove-ArrayElement,
    ConvertTo-EnhancedHTML, ConvertTo-EnhancedHTMLFragment, Get-PII
Export-ModuleMember -Variable SBLogFile