#Powershell collector script for Dell Open Manage
#v1.0 vMAN.ch, 02.09.2017 - Initial Version - Requires Module --> https://github.com/dfinke/ImportExcel if $Format is XLS

<#

    .SYNOPSIS

    Collecting Alarm, state, inventory and other data depending on the 

    Script requires Powershell v3 and above.

    Run the command below to store user and pass in secure credential XML for each environment

        $cred = Get-Credential
        $cred | Export-Clixml -Path "E:\DellOM\Config\OM.xml"
		
		.\DellOpenManageReporter.ps1 -OMServer 'DellOM.vMan.ch' -CollectionType 'Inventory' -creds 'OM' -FileName 'Inventory' -OutputLocation 'D:\DellOM\Report\' -Format 'XLS'

#>

param
(
    [String]$OMServer,
    [String]$CollectionType,
    [String]$creds,
    [String]$FileName,
    [String]$OutputLocation,
    [String]$Format
)

#Vars
$ScriptPath = (Get-Item -Path ".\" -Verbose).FullName
$LogFileLoc = $ScriptPath + '\Log\Logfile.log'
$RunDateTime = (Get-date)
$RunDateTime = $RunDateTime.tostring("yyyyMMddHHmmss")

#Logging Function
Function Log([String]$message, [String]$LogType, [String]$LogFile){
    $date = Get-Date -UFormat '%m-%d-%Y %H:%M:%S'
    $message = $date + "`t" + $LogType + "`t" + $message
    $message >> $LogFile
}

#Take all certs.
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls

#Get Stored Credentials

if($creds -gt ""){

    $cred = Import-Clixml -Path "$ScriptPath\config\$creds.xml"
    }
    else
    {
    echo "Credentials not specified, stop hammer time!"
    Exit
    }



#Script Starts here baby!

switch($CollectionType)
    {

Alerts {

        Echo "Running Alert Report"
        Log -Message "Running Alert Report" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 


$AlertFilterList = @()



$AlertFilterURL = 'https://' + $OMServer + ":2607/api/OME.svc/AlertFilters"

#Get List of Event filters

Echo "Connecting to $AlertFilterURL to get a list of AlertFilters"
Log -Message "Connecting to $AlertFilterURL to get a list of AlertFilters" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

[xml]$AlertFilter = Invoke-RestMethod -Method Get -Uri $AlertFilterURL -Credential $cred

ForEach ($AlertGroup in $AlertFilter.GetAllAlertFiltersResponse.GetAllAlertFiltersResult.AlertFilter){

                $AlertFilterList += New-Object PSObject -Property @{

                Name       = $AlertGroup.Name
                Id         = $AlertGroup.Id
                Type       = $AlertGroup.Type
                IsEnabled  = $AlertGroup.IsEnabled 
                IsReadOnly = $AlertGroup.IsReadOnly

}
}

#Start Collecting all Alerts for Filters which are enabled.

$AlertReport = @()

[Array]$AlertTypes = $AlertFilterList | Where-Object IsEnabled -EQ 'True' | select Name,Id,Type 

            Foreach ($AlertType in $AlertTypes.Id){

                $AlertURL = 'https://' + $OMServer + ":2607/api/OME.svc/AlertFilters/$AlertType/Alerts"

                Echo "Connecting to $AlertURL and Dumping alerts"
                Log -Message "Connecting to $AlertURL and Dumping alerts" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

                [xml]$Alerts = Invoke-RestMethod -Method Get -Uri $AlertUrl -Credential $cred

                If ($Alerts.GetAlertsForFilterResponse.GetAlertsForFilterResult.Alert){

                ForEach ($Alert in $Alerts.GetAlertsForFilterResponse.GetAlertsForFilterResult.Alert){

                    $AlertReport += New-Object PSObject -Property @{

                    AlertTypeID           = $AlertType
                    DeviceIdentifier	  = $Alert.DeviceIdentifier
                    DeviceName            = $Alert.DeviceName
                    DeviceNodeId	      = $Alert.DeviceNodeId
                    DeviceServiceTag	  = $Alert.DeviceServiceTag
                    DeviceSystemModelType = $Alert.DeviceSystemModelType
                    DeviceType            = $Alert.DeviceType
                    DeviceTypeName        = $Alert.DeviceTypeName
                    EventCategory         = $Alert.EventCategory
                    EventSource           = $Alert.EventSource
                    Id                    = $Alert.Id
                    IsIdrac               = $Alert.IsIdrac
                    IsInband              = $Alert.IsInband
                    Message               = $Alert.Message
                    OSName                = $Alert.OSName
                    Package               = $Alert.Package
                    SNMPEnterpriseOID     = $Alert.SNMPEnterpriseOID
                    SNMPGenericTrapID     = $Alert.SNMPGenericTrapID
                    SNMPSpecificTrapID    = $Alert.SNMPSpecificTrapID
                    Severity              = $Alert.Severity
                    SourceName            = $Alert.SourceName
                    Status                = $Alert.Status
                    Time                  = $Alert.Time
                }
               }
              }

}

#Sort Report by date / time, Save output to file

Echo "Sorting Report By Date / Time"
Log -Message "Sorting Report By Date / Time" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

$AlertReport = $AlertReport | Sort-Object { $_.Time -as [datetime] }

    If ($Format -eq 'XLS'){ 

        $OutputExcelFile = $OutputLocation + '\' + $FileName + '.xlsx'

        Echo "Exporting XLSX Report to $OutputExcelFile"
        Log -Message "Exporting XLSX Report to $OutputExcelFile" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

        $AlertFilterList  | select Name,Id,Type,IsEnabled,IsReadOnly | Export-Excel $OutputExcelFile -WorkSheetname AlertTypes   
        $AlertReport | select AlertTypeID,DeviceIdentifier,DeviceName,DeviceNodeId,DeviceServiceTag,DeviceSystemModelType,DeviceType,DeviceTypeName,EventCategory,EventSource,Id,IsIdrac,IsInband,Message,OSName,Package,SNMPEnterpriseOID,SNMPGenericTrapID,SNMPSpecificTrapID,Severity,SourceName,Status,Time | Export-Excel $OutputExcelFile -WorkSheetname Alerts 

    }

    If ($Format -eq 'CSV'){ 

        $OutputCSVFile = $OutputLocation + '\' + $FileName + '.csv'

        Echo "Exporting CSV Report to $OutputCSVFile"
        Log -Message "Exporting CSV Report to $OutputCSVFile" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

        $AlertReport | select AlertTypeID,DeviceIdentifier,DeviceName,DeviceNodeId,DeviceServiceTag,DeviceSystemModelType,DeviceType,DeviceTypeName,EventCategory,EventSource,Id,IsIdrac,IsInband,Message,OSName,Package,SNMPEnterpriseOID,SNMPGenericTrapID,SNMPSpecificTrapID,Severity,SourceName,Status,Time | Export-csv $OutputCSVFile 

    }

    echo 'All done, Terminating Script'
    Log -Message "All done, Terminating Script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    
    Remove-Variable *  -Force -ErrorAction SilentlyContinue

}


Inventory {

        Echo "Running Group Inventory Report"
        Log -Message "Running Group Inventory Report" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 


$GroupInventoryList = @()


$GroupInventoryFilterURL = 'https://' + $OMServer + ":2607/api/OME.svc/DeviceGroups"
#Get List of Groups filters

Echo "Connecting to $GroupInventoryFilterURL to get a list of Inventory Groups"
Log -Message "Connecting to $GroupInventoryFilterURL to get a list of Inventory Groups" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

[xml]$GroupInventoryFilter = Invoke-RestMethod -Method Get -Uri $GroupInventoryFilterURL -Credential $cred 

$GroupInventoryFilteredList = $GroupInventoryFilter.GetDeviceGroupsResponse.GetDeviceGroupsResult.DeviceGroup | Where { $_.DeviceCount -gt 0 }

ForEach ($Group in $GroupInventoryFilteredList){

                $GroupInventoryList += New-Object PSObject -Property @{

                Name          = $Group.Name
                Id            = $Group.Id
                Type          = $Group.Type
                Description   = $Group.Description
                DeviceCount   = $Group.DeviceCount
                RollupHealth  = $Group.RollupHealth

}
}

#Start Collecting all Devices in each of the Device Groups.


$DeviceReport = @()
$DeviceSoftwareReport = @()
$DeviceNICReport = @()
$DeviceFirmwareReport = @()
$DeviceMemoryReport = @()
$DeviceProcessorReport = @()


[Array]$DeviceGroup = $GroupInventoryList | select Name,Id

            Foreach ($DeviceGroupId in $DeviceGroup.Id){

                    $DeviceGroupURL = 'https://' + $OMServer + ":2607/api/OME.svc/DeviceGroups/$DeviceGroupId/Devices"

                    Echo "Connecting to $DeviceGroupURL and Dumping Device Groups"
                    Log -Message "Connecting to $DeviceGroupURL and Dumping Device Groups" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

                        [xml]$Devices = Invoke-RestMethod -Method Get -Uri $DeviceGroupURL -Credential $cred

                            ForEach ($Device in $Devices.GetDevicesResponse.GetDevicesResult.Device){

                                $DeviceReport += New-Object PSObject -Property @{

                                    GroupId            = $DeviceGroupId
                                    AssetTag           = $Device.AssetTag
                                    DNSName            = $Device.DNSName
                                    DiscoveryTime      = $Device.DiscoveryTime
                                    ExpressServiceCode = $Device.ExpressServiceCode
                                    GlobalStatus       = $Device.GlobalStatus
                                    Id                 = $Device.Id
                                    InventoryTime      = $Device.InventoryTime
                                    IsIdrac            = $Device.IsIdrac
                                    IsInband           = $Device.IsInband
                                    LaunchURL          = $Device.LaunchURL
                                    Name               = $Device.Name
                                    NodeId             = $Device.NodeId
                                    OSName             = $Device.OSName
                                    OSRevision         = $Device.OSRevision
                                    PowerStatus        = $Device.PowerStatus
                                    ServiceTag         = $Device.ServiceTag
                                    StatusTime         = $Device.StatusTime
                                    SystemId           = $Device.SystemId
                                    SystemModel        = $Device.SystemModel
                                    Type               = $Device.Type

                            }
                           }

}

#Extract Inventory information from each Device.


$DeviceUnique = $DeviceReport | Select Name,Id,AssetTag,NodeId,DNSName,DiscoveryTime,ExpressServiceCode,GlobalStatus,InventoryTime,IsIdrac,IsInband,LaunchURL,OSName,OSRevision,PowerStatus,ServiceTag,StatusTime,SystemId,SystemModel,Type -Unique

            Foreach ($DeviceId in $DeviceUnique.id){

                    $DeviceURL = 'https://' + $OMServer + ":2607/api/OME.svc/Devices/$DeviceId/Inventory"

                        If ($DeviceReport.IsIdrac -eq 'true'){

                        Echo "Connecting to $DeviceURL to get Device Inventory"
                        Log -Message "Connecting to $DeviceURL to get Device Inventory" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

                        [xml]$Dev = Invoke-RestMethod -Method Get -Uri $DeviceURL -Credential $cred

                        $DevSoftware = $Dev.DeviceInventory2Response.DeviceInventory2Result.Software.Software

                            ForEach ($Software in $DevSoftware){

                                $DeviceSoftwareReport += New-Object PSObject -Property @{

                                DeviceID      = $DeviceId
                                Description   = $Software.Description
                                Type          = $Software.Type
                                Version       = $Software.Version

                                }
                            }

                        $DevNIC = $Dev.DeviceInventory2Response.DeviceInventory2Result.NIC.NIC

                            ForEach ($Nic in $DevNIC){

                                $DeviceNICReport += New-Object PSObject -Property @{

                                DeviceID      = $DeviceId
                                Description   = $Nic.Description
                                IPAddress     = $Nic.IPAddress
                                MACAddress    = $Nic.MACAddress
                                Pingable      = $Nic.Pingable
                                Vendor        = $Nic.Vendor

                                }
                            }

                        $DevFirmware = $Dev.DeviceInventory2Response.DeviceInventory2Result.Firmware.Firmware

                            ForEach ($Firmware in $DevFirmware){

                                $DeviceFirmwareReport += New-Object PSObject -Property @{

                                DeviceID      = $DeviceId
                                Name          = $Firmware.Name
                                Type          = $Firmware.Type
                                Version       = $Firmware.Version

                                }
                            }

                        $DevMemory = $Dev.DeviceInventory2Response.DeviceInventory2Result.Memory.MemoryEntries.Memory

                            ForEach ($Memory in $DevMemory){

                                $DeviceMemoryReport += New-Object PSObject -Property @{

                                DeviceID      = $DeviceId
                                Name          = $Memory.Name
                                Type          = $Memory.Type
                                Manufacturer  = $Memory.Manufacturer
                                Size          = $Memory.Size
                                PartNumber    = $Memory.PartNumber

                                }
                            }


                        $DevProcessor = $Dev.DeviceInventory2Response.DeviceInventory2Result.Processor.Processor

                            ForEach ($Processor in $DevProcessor){

                                $DeviceProcessorReport += New-Object PSObject -Property @{

                                DeviceID      = $DeviceId
                                Brand         = $Processor.Brand
                                Cores         = $Processor.Cores
                                CurSpeed      = $Processor.CurSpeed
                                MaxSpeed      = $Processor.MaxSpeed
                                Model         = $Processor.Model

                                }
                            }
                        }
}





#Output to file

    If ($Format -eq 'XLS'){ 

        $OutputExcelFile = $OutputLocation + '\' + $FileName + '.xlsx'

        Echo "Exporting XLSX Report to $OutputExcelFile"
        Log -Message "Exporting XLSX Report to $OutputExcelFile" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

        $GroupInventoryList  | select Name,Id | Export-Excel $OutputExcelFile -WorkSheetname GroupInventory  
        $DeviceReport | select GroupId,Name,Id,AssetTag,NodeId,DNSName,DiscoveryTime,ExpressServiceCode,GlobalStatus,InventoryTime,IsIdrac,IsInband,LaunchURL,OSName,OSRevision,PowerStatus,ServiceTag,StatusTime,SystemId,SystemModel,Type | Export-Excel $OutputExcelFile -WorkSheetname Devices
        $DeviceSoftwareReport | select DeviceID,Description,Type,Version | Export-Excel $OutputExcelFile -WorkSheetname Software
        $DeviceFirmwareReport | select DeviceID,Name,Type,Version | Export-Excel $OutputExcelFile -WorkSheetname Firmware
        $DeviceMemoryReport | select DeviceID,Name,Type,Manufacturer,Size,PartNumber | Export-Excel $OutputExcelFile -WorkSheetname Memory
        $DeviceProcessorReport | select DeviceID,Brand,Cores,CurSpeed,MaxSpeed,Model | Export-Excel $OutputExcelFile -WorkSheetname Processor
        $DeviceNICReport | select DeviceID,Description,IPAddress,MACAddress,Pingable,Vendor | Export-Excel $OutputExcelFile -WorkSheetname NIC
    }

    If ($Format -eq 'CSV'){ 

        $OutputCSVFile = $OutputLocation + '\' + $FileName + '.csv'

        Echo "Exporting CSV Report to $OutputCSVFile"
        Log -Message "Exporting CSV Report to $OutputCSVFile" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc 

        $DeviceReport | select GroupId,Name,Id,AssetTag,NodeId,DNSName,DiscoveryTime,ExpressServiceCode,GlobalStatus,InventoryTime,IsIdrac,IsInband,LaunchURL,OSName,OSRevision,PowerStatus,ServiceTag,StatusTime,SystemId,SystemModel,Type | Export-Excel $OutputExcelFile -WorkSheetname Devices

    }

    echo 'All done, Terminating Script'
    Log -Message "All done, Terminating Script" -LogType "JOB-$RunDateTime" -LogFile $LogFileLoc
    
    #Remove-Variable *  -Force -ErrorAction SilentlyContinue

}

Default {"

Script does something!!


"}

}