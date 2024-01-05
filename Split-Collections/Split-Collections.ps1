<#
.SYNOPSIS
This script splits a collection into multiple child collections in Microsoft Endpoint Configuration Manager (MEMCM).

.DESCRIPTION
The Split-Collections.ps1 script takes a collection name, the desired number of child collections, MEMCM site code, SMS provider, and the base name for the child collections. It then splits the specified collection into the specified number of child collections, evenly distributing the devices from the original collection to the child collections.

.PARAMETER CollectionName
The name of the collection that needs to be split into child collections.

.PARAMETER NummberOfCollections
The number of child collections desired.

.PARAMETER SiteCode
The MEMCM site code.

.PARAMETER SiteServer
The SMS provider.

.PARAMETER NewCollectionName
The base name for the child collections. It is recommended to add a space to the end of the name.

.PARAMETER ResetLogFile
If specified, the log file will be reset.

.EXAMPLE
.\Split-Collections.ps1 -CollectionName "MyCollection" -NummberOfCollections 3 -SiteCode "ABC" -SiteServer "SMSProvider" -NewCollectionName "ChildCollection " -ResetLogFile

This example splits the collection named "MyCollection" into 3 child collections. The MEMCM site code is "ABC" and the SMS provider is "SMSProvider". The child collections will be named "ChildCollection 1", "ChildCollection 2", and "ChildCollection 3". The log file will be reset before running the script.

.NOTES
- The function Split-Collection was originally written by Peter van der Woude and edited for use in this script. (https://www.petervanderwoude.nl/post/divide-a-collection-into-multiple-smaller-collections-in-configmgr-2012-via-powershell/)
- The function Add-ResourceToCollection was written by Keith Garner. (https://keithga.wordpress.com/2013/08/26/adding-devices-to-a-collection-with-powershell/)
- The function Write-Log was written by Ben Whitmore and edited for use in this script. (https://github.com/byteben/Win32App-Migration-Tool/blob/138fad5cb5206394dee2ff90f93f112bf44af1ae/Private/Write-Log.ps1#L4)
- The default log location is C:\Windows\Logs and the log file is named SplitCollections.log.
- This script requires the ConfigurationManager module to be imported.
- The script uses the Write-Log function to write log entries to a log file.
- The script uses the ConnectToSCCM function to connect to the MEMCM site.
- The script uses the Split-Collection function to split the collection into child collections.
- The Add-ResourceToCollection function is used by the Split-Collection function to add devices to child collections.

.LINK
[Microsoft Endpoint Configuration Manager (MEMCM)](https://docs.microsoft.com/en-us/mem/configmgr/core/understand/introduction)

#>

param (
[Parameter(Mandatory = $false, ValuefromPipeline = $false, HelpMessage = "The component (script name) passed as LogID to the 'Write-Log' function")]
[string]$LogId = $($MyInvocation.MyCommand).Name,
[Parameter(Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'Collection that needs to be split into child collections.')]
$CollectionName,
[Parameter(Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'Number of child collections desired.')]
$NummberOfCollections,
[Parameter(Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'MEMCM Site Code.')]
$SiteCode,
[Parameter(Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'SMS Provider.')]
$SiteServer,
[Parameter(Mandatory = $true, ValueFromPipeline = $false, HelpMessage = 'Name child collections are based on. Recommend adding a space to end.')]
[string]$NewCollectionName,
[Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 2, HelpMessage = 'Name of the log file to write to. SplitCollections.log is the default log file')]
[String]$Log = 'SplitCollections.log',
[Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 4, HelpMessage = 'If specified, the log file will be reset.')]
[Switch]$ResetLogFile
)

$StartPath = Get-Location

#region Functions
function Split-Collection {
    param (
    [string]$CollectionName,
    [string]$NummberOfCollections,
    [string]$NewCollectionName
    )
    Write-Log -Message "Collection that will be split: $($CollectionName)." -Severity 1
    $Devices = Get-CMDevice -CollectionName $CollectionName -Fast
    $NumberOfDevices = $Devices.Count
    Write-Log -Message "$($CollectionName) contains $($NumberOfDevices) Devices." -Severity 1
    $NumberOfDevicesPerCollection = [math]::ceiling($NumberOfDevices / $NummberOfCollections)
    Write-Log -Message "Each child collection will contain approximately $($NumberOfDevicesPerCollection) devices." -Severity 1

    for ($i=1; $i -le $NummberOfCollections; $i++){
        $NewCollName = $NewCollectionName+$i
        Write-Log -Message "The $($i) child collection's name will be $($NewCollName)." -Severity 1
        $NewCollName

        Write-Log -Message "Attempting to create new child collection $($NewCollName)." -Severity 1
        Write-Progress -Activity "Creating New Collections" -Status "$PBPercentageColl% Complete:" -PercentComplete $PBPercentageColl

        Try {
            New-CMDeviceCollection -Name $NewCollName -LimitingCollectionName $CollectionName -ErrorAction Stop | Out-Null
        }
        Catch {
            Write-Host 'Unable to create new collection.' -ForegroundColor Red
            Write-Log -Message "Failed to create child collection $($NewCollName)." -Severity 3
            Write-Log -Message "Does this collection already exist?" -Severity 2
            Set-Location $StartPath
            break
        }

        Write-Log -Message "Gathering child collection $($NewCollName) members." -Severity 1
        $NewDevices = Get-Random -InputObject $Devices -Count $NumberOfDevicesPerCollection
        
        $PBTotal = ($NewDevices).Count
        $PBCurrent = 0
        $PBPercentage = 0

        $PBTotalColl = $NummberOfCollections
        $PBCurrentColl = 0
        $PBPercentageColl = 0

        Write-Log -Message "Beginning foeach loop to add devices to child collection $($NewCollName)." -Severity 1
        foreach ($NewDevice in $NewDevices) {
            Write-Progress -Activity "Adding Devices to $NewCollName" -Status "$PBPercentage% Complete:" -PercentComplete $PBPercentage
            Add-ResourceToCollection -SiteCode $SiteCode -SiteServer $SiteServer -CollectionName $NewCollName -System $NewDevice
            Write-Host $NewDevice.Name added to $NewCollName
            Write-Log -Message "$($NewDevice.Name) added to $($NewCollName)." -Severity 1

            $PBCurrent++
            $PBPercentage = [int](($PBCurrent / $PBTotal) * 100)
        }

        $Devices = $Devices | Where-Object { $NewDevices -notcontains $_ }
        $NumberOfDevicesLeft = $Devices.Count
        $NummberOfCollectionsLeft = $NummberOfCollections-$i
        Write-Log -Message "There are $($NumberOfDevicesLeft) number of devices left to be sorted into $($NummberOfCollectionsLeft) child collections." -Severity 1
        if ($NummberOfCollectionsLeft -gt 0) {
            $NumberOfDevicesPerCollection = [math]::ceiling($NumberOfDevicesLeft / $NummberOfCollectionsLeft)
        }
    Write-Host ""
    Write-Log -Message "" -Severity 1
    }
}

Function Add-ResourceToCollection {
    [CmdLetBinding()]
    Param(
        [string]   $SiteCode,
        [string]   $SiteServer,
        [string]   $CollectionName,
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $System
    )

    begin {
        $WmiArgs = @{ NameSpace = "Root\SMS\Site_$SiteCode"; ComputerName = $SiteServer }
        $CollectionQuery = Get-WmiObject @WMIArgs -Class SMS_Collection -Filter "Name = '$CollectionName' and CollectionType='2'"
        $InParams = $CollectionQuery.PSBase.GetMethodParameters('AddMembershipRules')
        $Cls = [WMIClass]"Root\SMS\Site_$($SiteCode):SMS_CollectionRuleDirect"
        $Rules = @()
    }
    process {
        foreach ( $sys in $System ) {
            $NewRule = $cls.CreateInstance()
            $NewRule.ResourceClassName = "SMS_R_System"
            $NewRule.ResourceID = $sys.ResourceID
            $NewRule.Rulename = $sys.Name
            $Rules += $NewRule.psobject.BaseObject 
        }
    }
    end {
        $InParams.CollectionRules += $Rules.psobject.BaseOBject
        $CollectionQuery.PSBase.InvokeMethod('AddMembershipRules',$InParams,$null) | Out-null
        $CollectionQuery.RequestRefresh() | out-null
    }             
}

Function ConnectToSCCM {
    [CmdLetBinding()]
    Param(
        [string]   $SiteCode,
        [string]   $SiteServer
    )
    $initParams = @{}
    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
    }

    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0, HelpMessage = 'Message to write to the log file')]
        [AllowEmptyString()]
        [String]$Message,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 1, HelpMessage = 'Location of the log file to write to')]
        [String]$LogFolder = "c:\windows\Logs", #$workingFolder is defined as a Global parameter in the main script
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 2, HelpMessage = 'Name of the log file to write to. Main is the default log file')]
        [String]$Log = 'SplitCollections.log',
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'LogId name of the script of the calling function')]
        [String]$LogId = $($MyInvocation.MyCommand).Name,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 3, HelpMessage = 'Severity of the log entry 1-3')]
        [ValidateSet(1, 2, 3)]
        [string]$Severity = 1,
        [Parameter(Mandatory = $false, ValueFromPipeline = $false, HelpMessage = 'The component (script name) passed as LogID to the Write-Log function including line number of invociation')]
        [string]$Component = [string]::Format('{0}:{1}', $logID, $($MyInvocation.ScriptLineNumber)),
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, Position = 4, HelpMessage = 'If specified, the log file will be reset')]
        [Switch]$ResetLogFile
    )

    Begin {
        $dateTime = Get-Date
        $date = $dateTime.ToString("MM-dd-yyyy", [Globalization.CultureInfo]::InvariantCulture)
        $time = $dateTime.ToString("HH:mm:ss.ffffff", [Globalization.CultureInfo]::InvariantCulture)
        $logToWrite = Join-Path -Path $LogFolder -ChildPath $Log
    }

    Process {
        if ($PSBoundParameters.ContainsKey('ResetLogFile')) {
            try {

                # Check if the logfile exists. We only need to reset it if it already exists
                if (Test-Path -Path $logToWrite) {

                    # Create a StreamWriter instance and open the file for writing
                    $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList $logToWrite
        
                    # Write an empty string to the file without the append parameter
                    $streamWriter.Write("")
        
                    # Close the StreamWriter, which also flushes the content to the file
                    $streamWriter.Close()
                    Write-Host ("Log file '{0}' wiped" -f $logToWrite) -ForegroundColor Yellow
                }
                else {
                    Write-Host ("Log file not found at '{0}'. Not restting log file" -f $logToWrite) -ForegroundColor Yellow
                }
            }
            catch {
                Write-Error -Message ("Unable to wipe log file. Error message: {0}" -f $_.Exception.Message)
                throw
            }
        }
            
        try {

            # Extract log object and construct format for log line entry
            foreach ($messageLine in $Message) {
                $logDetail = [string]::Format('<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="{4}" type="{5}" thread="{6}" file="">', $messageLine, $time, $date, $Component, $Context, $Severity, $PID)

                # Attempt log write
                try {
                    $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList $logToWrite, 'Append'
                    $streamWriter.WriteLine($logDetail)
                    $streamWriter.Close()
                }
                catch {
                    Write-Error -Message ("Unable to append log entry to '{0}' file. Error message: {1}" -f $logToWrite, $_.Exception.Message)
                    throw
                }
            }
        }
        catch [System.Exception] {
            Write-Warning -Message ("Unable to append log entry to '{0}' file" -f $logToWrite)
            throw
        }
    }
}
#endregion

if ($ResetLogFile) {
    Write-Log -Message $null -ResetLogFile
}

Write-Log -Message 'Beginning Script' -Severity 1
Write-Log -Message 'Calling Function ConnectToSCCM' -Severity 1
ConnectToSCCM -SiteCode $SiteCode -SiteServer $SiteServer
Write-Log -Message 'Calling Function SplitCollection' -Severity 1
Split-Collection -CollectionName $CollectionName -NummberOfCollections $NummberOfCollections -NewCollectionName $NewCollectionName

Write-Log -Message "Reverting Path to $($StartPath) and ending script." -Severity 1
Set-Location $StartPath



