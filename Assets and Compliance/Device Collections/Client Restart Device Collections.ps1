#Load Configuration Manager PowerShell Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5)+ '\ConfigurationManager.psd1')

#Get SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

#Error Handling and output
Clear-Host
$ErrorActionPreference= 'SilentlyContinue'

#Create Default Folder 
$CollectionFolder = @{Name ="Operational"; ObjectType =5000; ParentContainerNodeId =0}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolder -ComputerName $SiteCode.Root
$FolderPath =($SiteCode.Name +":\DeviceCollection\" + $CollectionFolder.Name)

#Set Default limiting collections
$LimitingCollection = "All Desktop and Server Clients"

$Schedule =New-CMSchedule –RecurInterval Days –RecurCount 7


#Find Existing Collections
$ExistingCollections = Get-CMDeviceCollection -Name "* | *" | Select-Object CollectionID, Name

#List of Collections Query
$DummyObject = New-Object -TypeName PSObject 
$Collections = @()

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 1"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr, File Rename"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 2"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of a File Rename operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | File Rename"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 3"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager or a File Rename operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Windows Update"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 4"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of a Windows Update."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr, Windows Update"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 5"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager or Windows Updates."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | File Rename, Windows Update"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 6"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of a File Rename operation or Windows Updates."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr, File Rename, Windows Update"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 7"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager, A file Rename operation or Windows Updates."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 8"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of an Add\Remove Feature operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr, Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 9"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager or a Add\Remove Feature operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | File Rename, Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 10"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of a File Rename or Add\Remove Feature operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr, File Rename, Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 11"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager, a File Rename or Add\Remove Feature operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Windows Update, Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 12"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of a Windows Update or Add\Remove Feature operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr, Windows Update, Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 13"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager, a Windows Update or Add\Remove Feature operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | File Rename, Windows Update, Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 14"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of a File Rename, Windows Update or Add\Remove Feature operation."}}

$Collections +=
$DummyObject |
Select-Object @{L="Name"
; E={"Client Restart | Config Mgr, File Rename, Windows Update, Add Remove Feature"}},@{L="Query"
; E={"select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System join sms_combineddeviceresources on  sms_combineddeviceresources.resourceid = sms_r_system.resourceid  where sms_combineddeviceresources.clientstate = 15"}},@{L="LimitingCollection"
; E={$LimitingCollection}},@{L="Comment"
; E={"All devices detected needing a restart because of Config Manager, a File Rename, Windows Update or Add\Remove Feature operation."}}

#Handle the Client Restart | All Collection removal first as Inclusive collections reside in here and cause issues with Collection Removals
If ($ExistingCollections.Name -contains "Client Restart | All") 
    {
        Remove-CMDeviceCollection -Name "Client Restart | All" -Force
        Write-host "*** Collection Client Restart | All removed and will be recreated ***"
    }

New-CMDeviceCollection -Name "Client Restart | All" -Comment "All Devices requiring a restart." -LimitingCollectionName $Collection.LimitingCollection -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
Write-host *** Collection "Client Restart | All" created ***
Move-CMObject -FolderPath $FolderPath -InputObject $(Get-CMDeviceCollection -Name "Client Restart | All")
Write-host *** Collection "Client Restart | All" moved to $CollectionFolder.Name folder***

#Check Existing Collections
$Overwrite = 1
$ErrorCount = 0
$ErrorHeader = "The script has already been run. The following collections already exist in your environment:`n`r"
$ErrorCollections = @()
$ErrorFooter = "Would you like to delete and recreate the collections above? (Default : No) "
$ExistingCollections | Sort-Object Name | ForEach-Object {If($Collections.Name -Contains $_.Name) {$ErrorCount +=1 ; $ErrorCollections += $_.Name}}

#Error
If ($ErrorCount -ge1) 
    {
    Write-Host $ErrorHeader $($ErrorCollections | ForEach-Object {(" " + $_ + "`n`r")}) $ErrorFooter -ForegroundColor Yellow -NoNewline
    $ConfirmOverwrite = Read-Host "[Y/N]"
    If ($ConfirmOverwrite -ne "Y") {$Overwrite =0}
    }

#Create Collection And Move the collection to the right folder
If ($Overwrite -eq1) {
$ErrorCount =0

ForEach ($Collection
In $($Collections | Sort-Object LimitingCollection -Descending))

{
If ($ErrorCollections -Contains $Collection.Name)
    {
    Remove-CMDeviceCollection -Name $Collection.Name -Force
    Write-host *** Collection $Collection.Name removed and will be recreated ***
    }
}

ForEach ($Collection In $($Collections | Sort-Object LimitingCollection))
{

Try 
    {
    New-CMDeviceCollection -Name $Collection.Name -Comment $Collection.Comment -LimitingCollectionName $Collection.LimitingCollection -RefreshSchedule $Schedule -RefreshType 2 | Out-Null
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Collection.Name -QueryExpression $Collection.Query -RuleName $Collection.Name
    Write-host *** Collection $Collection.Name created ***
    If ($Collection.Name -like "Client Restart |*") {Add-CMDeviceCollectionIncludeMembershipRule -CollectionName "Client Restart | All" -IncludeCollectionName $Collection.Name | Out-Null ; Write-host "*** Collection" $Collection.Name "Included in Client Restart | All ***"}
    }

Catch {
        Write-host "-----------------"
        Write-host -ForegroundColor Red ("There was an error creating the: " + $Collection.Name + " collection.")
        Write-host "-----------------"
        $ErrorCount += 1
        Pause
}

Try {
        Move-CMObject -FolderPath $FolderPath -InputObject $(Get-CMDeviceCollection -Name $Collection.Name)
        Write-host *** Collection $Collection.Name moved to $CollectionFolder.Name folder***
    }

Catch {
        Write-host "-----------------"
        Write-host -ForegroundColor Red ("There was an error moving the: " + $Collection.Name +" collection to " + $CollectionFolder.Name +".")
        Write-host "-----------------"
        $ErrorCount += 1
        Pause
      }

}

If ($ErrorCount -ge1) {

        Write-host "-----------------"
        Write-Host -ForegroundColor Red "The script execution completed, but with errors."
        Write-host "-----------------"
        Pause
}

Else{
        Write-host "-----------------"
        Write-Host -ForegroundColor Green "Script execution completed without error. Operational Collections created sucessfully."
        Write-host "-----------------"
        Pause
    }
}

Else {
        Write-host "-----------------"
        Write-host -ForegroundColor Red ("The following collections already exist in your environment:`n`r" + $($ErrorCollections | ForEach-Object {(" " +$_ + "`n`r")}) + "Please delete all collections manually or rename them before re-executing the script! You can also select Y to do it automaticaly")
        Write-host "-----------------"
        Pause
}