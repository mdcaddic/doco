$VerbosePreference = 'Continue'
Function ConvertTo-MDTest {

    [cmdletbinding()]
        [outputtype([string[]])]
        [alias('ctm')]
    
        Param(
            [Parameter(Position = 0, ValueFromPipeline)]
            [object]$Inputobject,
            [Parameter()]
            [string]$Title,
            [string[]]$PreContent,
            [string[]]$PostContent,
            [ValidateScript( {$_ -ge 10})]
            [int]$Width = 80,
            #display results as a markdown table
            [switch]$AsTable
        )
    
        Begin {
            Write-Verbose "[BEGIN  ] Starting $($myinvocation.MyCommand)"
            #initialize an array to hold incoming data
            $data = @()
    
            #initialize an empty here string for markdown text
            $Text = @"

"@
            If ($title) {
                Write-Verbose "[BEGIN  ] Adding Title: $Title"
                $Text += "# $Title`n`n"
            }
            If ($precontent) {
                Write-Verbose "[BEGIN  ] Adding Precontent"
                $Text += $precontent
                $text += "`n`n"
            }
    
        } #begin
        Process {
            #add incoming objects to data array
            Write-Verbose "[PROCESS] Adding processed object"
            $data += $Inputobject
    
        } #process
        End {
            #add the data to the text
            if ($data) {
                if ($AsTable) {
                    Write-Verbose "[END    ] Formatting as a table"
                    $names = $data[0].psobject.Properties.name
                    $head = "| $($names -join " | ") |"
                    $text += $head
                    $text += "`n"
    
                    $bars = "| $(($names -replace '.','-') -join " | ") |"
    
                    $text += $bars
                    $text += "`n"
    
                    foreach ($item in $data) {
                        $line = "| "
                        $values = @()
                        for ($i = 0; $i -lt $names.count; $i++) {
                            
                            #if an item value contains return and new line replace them with <br> Issue #97
                            if ($item.($names[$i]) -match "`n") {
                                Write-Verbose "[END    ] Replacing line returns for property $($names[$i])"
                                [string]$val = $($item.($names[$i])).replace("`r`n","<br>") -join ""
                                Write-Verbose $val
                            }
                            else {
                                [string]$val = $item.($names[$i])
                            }
                            
                            $values += $val
                        }
                        $line += $values -join " | "
                        $line += " |"
                        $text += $line
                        $text += "`r"
                    }
                }
                else {
                    #convert data to strings and trim each line
                    Write-Verbose "[END    ] Converting data to strings"
                    [string]$trimmed = (($data | Out-String -Width $width).split("`n")).ForEach({ "$($_.trim())`n" })
                    Write-Verbose "[END    ] Adding to markdown"
                    $clean = $($trimmed.trimend())
                    $text += @"
    ``````text
    $clean
"@
            } #else as text
        } #if $data
        If ($postcontent) {
            Write-Verbose "[END    ] Adding postcontent"
            $text += "`n"
            $text += $postcontent
        }
    
        #write the markdown to the pipeline
        $text.TrimEnd()
        Write-Verbose "[END    ] Ending $($myinvocation.MyCommand)"
    } #end
    }

    
##########WARNING##########WARNING##########WARNING##########WARNING##########WARNING##########WARNING##########WARNING##
# This script and supporting materials may not be reproduced or distributed, in whole or in part, without the prior     #
# written permission of either Nathaniel O'Reilly or oobe PL. Any reproduction or distribution, in whatever form and    #
# by whatever media, is expressly prohibited without the prior written consent of either Nathaniel O'Reilly or oobe PL. #
##########WARNING##########WARNING##########WARNING##########WARNING##########WARNING##########WARNING##########WARNING##

Write-Host "     ______  _______  ______  __   ____  ______ "
Write-Host "    /  __  \ |  ____||  ____||  | / ___||  ____|"
Write-Host "   |  /  \  ||  |__  |  ___  |  || /    | |____ "
Write-Host "   |  |  |  ||  ___| |  ___| |  || |    |  ____|"
Write-Host "   |  \__/  ||  |    |  |    |  || \___ | |____ "
Write-Host "    \______/ |__|    |__|    |__| \____||______|"
Write-host "             _      ______      _      ____     "
Write-Host "            / \    |  __  \    / \    /  __|    "
Write-Host "           / _ \   |    __/   / _ \  |  /       "
Write-Host "          / ___ \  |  __  \  / ___ \ |  \__     "
Write-Host "         /_/   \_\ |______/ /_/   \_\ \____|    "
Write-Host "                                                     V1.0 "


###########################
#region Arrays
###########################
$InboundMailConnectorsArray                  =@()
$OutboundMailConnectorsArray                 =@()
$MXrecordsArray                              =@()
$SPFRecordsArray                             =@()
$DKIMRecordsArray                            =@()
$DMARCRecordsArray                           =@()
$EOSPFrecordslist                            =@()
$RemoteDomainsArray                          =@()
$UserMailboxConfigArray                      =@()
$AuthenticationPolicyArray                   =@()
$AssociatedDocumentationArray                =@()
$EOMXrecordArray                             =@()
$CASMailboxPlanArray                         =@()
$EOAuthenticationPolicyArray                 =@()
$OWAMailboxPolicyArray                       =@()
$EOAddressListsArray                         =@()
$EOPConnectionFilterArray                    =@()
$EOMXrecordslist                             =@()
$EOMXrecordMXlist                            =@()
$EOPMalwareFilterArray                       =@()
$EOPPolicyFilterArray                        =@()
$EOPPolicyFilteringExceptArray               =@()
$EOPContentFilterArray                       =@()
$EOPPolicyFilteringEactionArray              =@()
$EOPPolicyFilteringConditionsArray           =@()
$EOPContentFilterIncreaseScoreArray          =@()
$EOPContentFilterMarkAsSpameArray            =@()
$EOPContentFilterEndUserSpamNotificationArray=@()
$EOPMalwareFilterCustomNotificationArray     =@()
#endregion Arrays

###########################
#Variables
###########################
Set-Location $PSScriptroot
$domains                        = @("precisionservices.biz")
$Logging                        = "$pwd\Logs\ExchangeOnlineProtectionABAC_$($(Get-Date).ToString(`"yyyy-MM-dd hhmmss`")).log"
$ReportFileName                 = "$pwd\Reports\EOP_MDTest.md" # "$pwd\Reports\Office365ABAC_$($(Get-Date).ToString(`"yyyy-MM-dd hhmmss`")).md"
# $Tenant                         = "precisionservicesptyltd"
# $admindomain                    = "precisionservicesptyltd.onmicrosoft.com"
$global:UserPrincipalName       = "martin@precisionservices.biz"

#Enable Full Logging
Start-Transcript -Path $Logging
Write-Host ""
# Overview
Write-Host " - Overview" -foregroundcolor Gray
Set-Content -Path $ReportFileName -value "# Overview"
#############################################
#Exchange Online Protection
#############################################
Write-Host "Querying Exchange Online Protection configuration..." -foregroundcolor Yellow

#############################################
#Exchange Online Protection - Connection Filtering
#############################################

Write-Host " - Connection Filtering" -foregroundcolor Gray

$EOPConnectionFilter = Get-HostedConnectionFilterPolicy 

If ($null -eq $EOPConnectionFilter) {
     $EOPConnectionFilterDetail = [ordered]@{
        "Name"                             = "Not Configured"
        "IP Allow List"                    = "N/A"
        "IP Block List"                    = "N/A"
        "Enable Safe List"                 = "N/A"
        "Directory Based Edge Block Mode"  = "N/A"
    }
    $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPConnectionFilterDetail
    $EOPConnectionFilterArray += $EOConfigurationObject
}
else {
    If ($EOPConnectionFilter -isnot [array]) {
        $EOPConnectionFilterName    = $EOPConnectionFilter.name      
        $EOPConnectionFilterIPAllow  = $EOPConnectionFilter.IPAllowlist |out-string
        $EOPConnectionFilterIPAllow  = [string]$EOPConnectionFilterIPAllow.trim()
        $EOPConnectionFilterEnableSafeList    = $EOPConnectionFilter.EnableSafeList      
        $EOPConnectionFilterIPBlock  = $EOPConnectionFilter.IPBlocklist |out-string
        $EOPConnectionFilterIPBlock  = [string]$EOPConnectionFilterIPBlock.trim()
        $EOPConnectionFilterDirectorybasededgeblock  = $EOPConnectionFilter.DirectorybasedEdgeBlockMode
            
        $EOPConnectionFilterDetail = [ordered]@{
            "Name"                             = $EOPConnectionFilterName 
            "IP Allow List"                    = $EOPConnectionFilterIPAllow
            "IP Block List"                    = $EOPConnectionFilterIPBlock 
            "Enable Safe List"                 = $EOPConnectionFilterEnableSafeList 
            "Directory Based Edge Block Mode"  = $EOPConnectionFilterDirectorybasededgeblock
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPConnectionFilterDetail
        $EOPConnectionFilterArray += $EOConfigurationObject
    }
    else {
        $EOPConnectionFilter | Foreach-Object {
            if ($_ -eq $EOPConnectionFilter[-1]) {
                if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                    $EOPConnectionFilterName    = $_.name      
                    $EOPConnectionFilterIPAllow  = $_.IPAllowlist  |out-string
                    $EOPConnectionFilterIPAllow  = [string]$EOPConnectionFilterIPAllow.trim()
                    $EOPConnectionFilterEnableSafeList    = $_.EnableSafeList      
                    $EOPConnectionFilterIPBlock  = $_.IPBlocklist  |out-string
                    $EOPConnectionFilterIPBlock  = [string]$EOPConnectionFilterIPBlock.trim()
                    $EOPConnectionFilterDirectorybasededgeblock  = $_.DirectorybasedEdgeBlockMode
                }
            }
            else {
                if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                    $EOPConnectionFilterName    = $_.name      
                    $EOPConnectionFilterIPAllow  = $_.IPAllowlist  |out-string
                    $EOPConnectionFilterIPAllow  = [string]$EOPConnectionFilterIPAllow.trim()
                    $EOPConnectionFilterEnableSafeList    = $_.EnableSafeList      
                    $EOPConnectionFilterIPBlock  = $_.IPBlocklist  |out-string
                    $EOPConnectionFilterIPBlock  = [string]$EOPConnectionFilterIPBlock.trim()
                    $EOPConnectionFilterDirectorybasededgeblock  = $_.DirectorybasedEdgeBlockMode
                }
            }
            $EOPConnectionFilterDetail = [ordered]@{
                "Name"                             = $EOPConnectionFilterName 
                "IP Allow List"                    = $EOPConnectionFilterIPAllow
                "IP Block List"                    = $EOPConnectionFilterIPBlock 
                "Enable Safe List"                 = $EOPConnectionFilterEnableSafeList 
                "Directory Based Edge Block Mode"  = $EOPConnectionFilterDirectorybasededgeblock
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPConnectionFilterDetail
            $EOPConnectionFilterArray += $EOConfigurationObject
        }
    }
}


#############################################
#Exchange Online Protection - Anti-Malware
#############################################
Write-Host " - Anti-Malware" -foregroundcolor Gray

$EOPMalwareFilter = Get-MalwareFilterPolicy

If ($null -eq $EOPMalwareFilter) {
     $EOPMalwareFilterDetail = [ordered]@{
        "Name"                             = "Not Configured"
        "Action"                           = "N/A"
        "Custom Notifications"            = "N/A"
    }
    $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
    $EOPMalwareFilterArray += $EOConfigurationObject
}
else {
    If ($EOPMalwareFilter -isnot [array]) {
        If (([string]::isnullorempty($EOPMalwareFilter.name)) -EQ $FALSE) {
            $EOPMalwareFiltername = $EOPMalwareFilter.name
        }
        else {
            $EOPMalwareFiltername = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Name"
            "Value"                 = $EOPMalwareFiltername
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($EOPMalwareFilter.CustomNotifications)) -EQ $FALSE) {
            $EOPMalwareFilterCustomNotifications = $EOPMalwareFilter.CustomNotifications
        }
        else {
            $EOPMalwareFilterCustomNotifications = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Custom Notifications"
            "Value"                 = $EOPMalwareFilterCustomNotifications
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If ($EOPMalwareFilter.customalerttext.length -gt 1) {
            $EOPMalwareFilterCustomNotificationArray += "Custom alert text: " +($EOPMalwareFilter.customalerttext|out-string)
        }
        If ($EOPMalwareFilter.custominternalsubject.length -gt 1) {
            $EOPMalwareFilterCustomNotificationArray += "Custom internal subject: " +($EOPMalwareFilter.custominternalsubject|out-string)
        }
        If ($EOPMalwareFilter.custominternalbody.length -gt 1) {
            $EOPMalwareFilterCustomNotificationArray +=  "Custom internal body: " +($EOPMalwareFilter.custominternalbody|out-string)
        }
        If ($EOPMalwareFilter.customexternalsubject.length -gt 1) {
            $EOPMalwareFilterCustomNotificationArray +=  "Custom external subject: " +($EOPMalwareFilter.customexternalsubject|out-string)
        }
        If ($EOPMalwareFilter.customExternalbody.length -gt 1) {
            $EOPMalwareFilterCustomNotificationArray +=  "Custom external body: " +($EOPMalwareFilter.customExternalbody|out-string)
        }
        If ($EOPMalwareFilter.customFromName.length -gt 1) {
            $EOPMalwareFilterCustomNotificationArray +=  "Custom from name: " +($EOPMalwareFilter.customFromName|out-string)
        }    
        If ($EOPMalwareFilter.customFromaddress.length -gt 1) {
            $EOPMalwareFilterCustomNotificationArray +=  "Custom from address: " +($EOPMalwareFilter.customFromaddress|out-string)
        }
        If ($EOPMalwareFilterCustomNotificationArray.length -lt 14) {
            $EOPMalwareFilterCustomNotificationArray = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Custom notification details"
            "Value"                 = [string]$EOPMalwareFilterCustomNotificationArray
        }
        $EOPMalwareFilterCustomNotificationArray =$null
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject
 
        If ($EOPMalwareFilter.InternalSenderAdminAddress.length -gt 1) {
            $EOPMalwareFilterInternalSenderAdminAddress = $EOPMalwareFilter.InternalSenderAdminAddress
        }
        else {
            $EOPMalwareFilterInternalSenderAdminAddress = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Internal Sender Admin Address"
            "Value"                 = $EOPMalwareFilterInternalSenderAdminAddress
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If ($EOPMalwareFilter.ExternalSenderAdminAddress.length -gt 1) {
            $EOPMalwareFilterExternalSenderAdminAddress = $EOPMalwareFilter.ExternalSenderAdminAddress
        }
        else {
            $EOPMalwareFilterExternalSenderAdminAddress = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "External Sender Admin Address"
            "Value"                 = $EOPMalwareFilterExternalSenderAdminAddress
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPMalwareFilter.Action)) -EQ $FALSE) {
            $EOPMalwareFilterAction = $EOPMalwareFilter.Action
        }
        else {
            $EOPMalwareFilterAction = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Action"
            "Value"                 = $EOPMalwareFilterAction
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPMalwareFilter.EnableInternalSenderNotifications)) -EQ $FALSE) {
            $EOPMalwareFilterEnableInternalSenderNotifications = $EOPMalwareFilter.EnableInternalSenderNotifications
        }
        else {
            $EOPMalwareFilterEnableInternalSenderNotifications = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Enable Internal Sender Notifications"
            "Value"                 = $EOPMalwareFilterEnableInternalSenderNotifications
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPMalwareFilter.EnableExternalSenderNotifications)) -EQ $FALSE) {
            $EOPMalwareFilterEnableExternalSenderNotifications = $EOPMalwareFilter.EnableExternalSenderNotifications
        }
        else {
            $EOPMalwareFilterEnableExternalSenderNotifications = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Enable External Sender Notifications"
            "Value"                 = $EOPMalwareFilterEnableExternalSenderNotifications
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPMalwareFilter.EnableInternalSenderAdminNotifications)) -EQ $FALSE) {
            $EOPMalwareFilterEnableInternalSenderAdminNotifications = $EOPMalwareFilter.EnableInternalSenderAdminNotifications
        }
        else {
            $EOPMalwareFilterEnableInternalSenderAdminNotifications = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Enable Internal Sender Admin Notifications"
            "Value"                 = $EOPMalwareFilterEnableInternalSenderAdminNotifications
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPMalwareFilter.EnableExternalSenderAdminNotifications)) -EQ $FALSE) {
            $EOPMalwareFilterEnableExternalSenderAdminNotifications = $EOPMalwareFilter.EnableExternalSenderAdminNotifications
        }
        else {
            $EOPMalwareFilterEnableExternalSenderAdminNotifications = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Enable External Sender Admin Notifications"
            "Value"                 = $EOPMalwareFilterEnableExternalSenderAdminNotifications
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPMalwareFilter.EnableFileFilter)) -EQ $FALSE) {
            $EOPMalwareFilterEnableFileFilter = $EOPMalwareFilter.EnableFileFilter
        }
        else {
            $EOPMalwareFilterEnableFileFilter = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Enable File Filter"
            "Value"                 = $EOPMalwareFilterEnableFileFilter
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPMalwareFilter.FileTypes)) -EQ $FALSE) {
            $EOPMalwareFilterFileTypes = $EOPMalwareFilter.FileTypes |Out-String
            $EOPMalwareFilterFileTypes = $EOPMalwareFilterFileTypes.trim()
        }
        else {
            $EOPMalwareFilterFileTypes = "Not Configured"
        }
        $EOPMalwareFilterDetail = [ordered]@{
            "Configuration Item"    = "Filter file types"
            "Value"                 = $EOPMalwareFilterFileTypes
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
        $EOPMalwareFilterArray += $EOConfigurationObject
    }
    else {
        $EOPMalwareFilter | Foreach-Object {
            if ($_ -eq $EOPMalwareFilter[-1]) {
                if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                    $EOPMalwareFiltername = $_.name
                }
                else {
                    $EOPMalwareFiltername = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Name [TBA]"
                    "Value"                 = $EOPMalwareFiltername
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.CustomNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterCustomNotifications = $_.CustomNotifications
                }
                else {
                    $EOPMalwareFilterCustomNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Custom Notifications"
                    "Value"                 = $EOPMalwareFilterCustomNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject

                If ($_.customalerttext.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray += "Custom alert text: " +($_.customalerttext|out-string)
                }
                If ($_.custominternalsubject.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray += "Custom internal subject: " +($_.custominternalsubject|out-string)
                }
                If ($_.custominternalbody.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom internal body: " +($_.custominternalbody|out-string)
                }
                If ($_.customexternalsubject.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom external subject: " +($_.customexternalsubject|out-string)
                }
                If ($_.customExternalbody.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom external body: " +($_.customExternalbody|out-string)
                }
                If ($_.customFromName.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom from name: " +($_.customFromName|out-string)
                }    
                If ($_.customFromaddress.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom from address: " +($_.customFromaddress|out-string)
                }
                If ($EOPMalwareFilterCustomNotificationArray.Length -lt 1) {
                    $EOPMalwareFilterCustomNotificationArray = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Custom notification details"
                    "Value"                 = [string]$EOPMalwareFilterCustomNotificationArray
                }
                $EOPMalwareFilterCustomNotificationArray =$null
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
         
                If ($_.InternalSenderAdminAddress.length -gt 1) {
                    $EOPMalwareFilterInternalSenderAdminAddress = $_.InternalSenderAdminAddress
                }
                else {
                    $EOPMalwareFilterInternalSenderAdminAddress = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Internal Sender Admin Address"
                    "Value"                 = $EOPMalwareFilterInternalSenderAdminAddress
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If ($_.ExternalSenderAdminAddress.length -gt 1) {
                    $EOPMalwareFilterExternalSenderAdminAddress = $_.ExternalSenderAdminAddress
                }
                else {
                    $EOPMalwareFilterExternalSenderAdminAddress = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "External Sender Admin Address"
                    "Value"                 = $EOPMalwareFilterExternalSenderAdminAddress
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
                
                If (([string]::isnullorempty($_.Action)) -EQ $FALSE) {
                    $EOPMalwareFilterAction = $_.Action
                }
                else {
                    $EOPMalwareFilterAction = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Action"
                    "Value"                 = $EOPMalwareFilterAction
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableInternalSenderNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableInternalSenderNotifications = $_.EnableInternalSenderNotifications
                }
                else {
                    $EOPMalwareFilterEnableInternalSenderNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable Internal Sender Notifications"
                    "Value"                 = $EOPMalwareFilterEnableInternalSenderNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableExternalSenderNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableExternalSenderNotifications = $_.EnableExternalSenderNotifications
                }
                else {
                    $EOPMalwareFilterEnableExternalSenderNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable External Sender Notifications"
                    "Value"                 = $EOPMalwareFilterEnableExternalSenderNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableInternalSenderAdminNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableInternalSenderAdminNotifications = $_.EnableInternalSenderAdminNotifications
                }
                else {
                    $EOPMalwareFilterEnableInternalSenderAdminNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable Internal Sender Admin Notifications"
                    "Value"                 = $EOPMalwareFilterEnableInternalSenderAdminNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableExternalSenderAdminNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableExternalSenderAdminNotifications = $_.EnableExternalSenderAdminNotifications
                }
                else {
                    $EOPMalwareFilterEnableExternalSenderAdminNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable External Sender Admin Notifications"
                    "Value"                 = $EOPMalwareFilterEnableExternalSenderAdminNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableFileFilter)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableFileFilter = $_.EnableFileFilter
                }
                else {
                    $EOPMalwareFilterEnableFileFilter = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable File Filter"
                    "Value"                 = $EOPMalwareFilterEnableFileFilter
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.FileTypes)) -EQ $FALSE) {
                    $EOPMalwareFilterFileTypes = $_.FileTypes |Out-String
                    $EOPMalwareFilterFileTypes = $EOPMalwareFilterFileTypes.trim()
                }
                else {
                    $EOPMalwareFilterFileTypes = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Filter file types"
                    "Value"                 = $EOPMalwareFilterFileTypes
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
            }
            else {
                if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                    $EOPMalwareFiltername = $_.name
                }
                else {
                    $EOPMalwareFiltername = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Name [TBA]"
                    "Value"                 = $EOPMalwareFiltername
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.CustomNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterCustomNotifications = $_.CustomNotifications
                }
                else {
                    $EOPMalwareFilterCustomNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Custom Notifications"
                    "Value"                 = $EOPMalwareFilterCustomNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject

                If ($_.customalerttext.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray += "Custom alert text: " +($_.customalerttext|out-string)
                }
                If ($_.custominternalsubject.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray += "Custom internal subject: " +($_.custominternalsubject|out-string)
                }
                If ($_.custominternalbody.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom internal body: " +($_.custominternalbody|out-string)
                }
                If ($_.customexternalsubject.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom external subject: " +($_.customexternalsubject|out-string)
                }
                If ($_.customExternalbody.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom external body: " +($_.customExternalbody|out-string)
                }
                If ($_.customFromName.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom from name: " +($_.customFromName|out-string)
                }    
                If ($_.customFromaddress.length -gt 1) {
                    $EOPMalwareFilterCustomNotificationArray +=  "Custom from address: " +($_.customFromaddress|out-string)
                }
                If ($EOPMalwareFilterCustomNotificationArray.length -lt 1) {
                    $EOPMalwareFilterCustomNotificationArray = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Custom notification details"
                    "Value"                 = [string]$EOPMalwareFilterCustomNotificationArray
                }
                $EOPMalwareFilterCustomNotificationArray =$null
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
         
                If ($_.InternalSenderAdminAddress.length -gt 1) {
                    $EOPMalwareFilterInternalSenderAdminAddress = $_.InternalSenderAdminAddress
                }
                else {
                    $EOPMalwareFilterInternalSenderAdminAddress = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Internal Sender Admin Address"
                    "Value"                 = $EOPMalwareFilterInternalSenderAdminAddress
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If ($_.ExternalSenderAdminAddress.length -gt 1) {
                    $EOPMalwareFilterExternalSenderAdminAddress = $_.ExternalSenderAdminAddress
                }
                else {
                    $EOPMalwareFilterExternalSenderAdminAddress = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "External Sender Admin Address"
                    "Value"                 = $EOPMalwareFilterExternalSenderAdminAddress
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
                If (([string]::isnullorempty($_.Action)) -EQ $FALSE) {
                    $EOPMalwareFilterAction = $_.Action
                }
                else {
                    $EOPMalwareFilterAction = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Action"
                    "Value"                 = $EOPMalwareFilterAction
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableInternalSenderNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableInternalSenderNotifications = $_.EnableInternalSenderNotifications
                }
                else {
                    $EOPMalwareFilterEnableInternalSenderNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable Internal Sender Notifications"
                    "Value"                 = $EOPMalwareFilterEnableInternalSenderNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableExternalSenderNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableExternalSenderNotifications = $_.EnableExternalSenderNotifications
                }
                else {
                    $EOPMalwareFilterEnableExternalSenderNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable External Sender Notifications"
                    "Value"                 = $EOPMalwareFilterEnableExternalSenderNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableInternalSenderAdminNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableInternalSenderAdminNotifications = $_.EnableInternalSenderAdminNotifications
                }
                else {
                    $EOPMalwareFilterEnableInternalSenderAdminNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable Internal Sender Admin Notifications"
                    "Value"                 = $EOPMalwareFilterEnableInternalSenderAdminNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableExternalSenderAdminNotifications)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableExternalSenderAdminNotifications = $_.EnableExternalSenderAdminNotifications
                }
                else {
                    $EOPMalwareFilterEnableExternalSenderAdminNotifications = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable External Sender Admin Notifications"
                    "Value"                 = $EOPMalwareFilterEnableExternalSenderAdminNotifications
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.EnableFileFilter)) -EQ $FALSE) {
                    $EOPMalwareFilterEnableFileFilter = $_.EnableFileFilter
                }
                else {
                    $EOPMalwareFilterEnableFileFilter = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Enable File Filter"
                    "Value"                 = $EOPMalwareFilterEnableFileFilter
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
        
                If (([string]::isnullorempty($_.FileTypes)) -EQ $FALSE) {
                    $EOPMalwareFilterFileTypes = $_.FileTypes |Out-String
                    $EOPMalwareFilterFileTypes = $EOPMalwareFilterFileTypes.trim()
                }
                else {
                    $EOPMalwareFilterFileTypes = "Not Configured"
                }
                $EOPMalwareFilterDetail = [ordered]@{
                    "Configuration Item"    = "Filter file types"
                    "Value"                 = $EOPMalwareFilterFileTypes
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPMalwareFilterDetail
                $EOPMalwareFilterArray += $EOConfigurationObject
            }
        }
    }
}

#############################################
#Exchange Online Protection - Policy Filtering
#############################################
Write-Host " - Policy Filtering" -foregroundcolor Gray

$EOPPolicyFilter = Get-TransportRule
If ($null -eq $EOPPolicyFilter) {
     $EOPPolicyFilterDetail = [ordered]@{
        "Name"                       = "Not Configured"
        "State"                      = "N/A"
        "Mode"                       = "N/A"
        "Priority"                   = "N/A"
        "Comments"                   = "N/A"
    }
    $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
    $EOPPolicyFilterArray += $EOConfigurationObject    
}
Else {
    If($EOPPolicyFilter -isnot [array]) {
        If(([string]::isnullorempty($EOPPolicyFilter.identity)) -EQ $FALSE) {
            $EOPPolicyfilterName = $($EOPPolicyFilter.identity |Out-String).trim()
        }
        else {
            $EOPPolicyfilterName = "Not Configured"
        }
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "Name [TBA]"
            "Value"                      = $EOPPolicyfilterName
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject
        
        If(([string]::isnullorempty($EOPPolicyFilter.Priority)) -EQ $FALSE) {
            $EOPPolicyfilterPriority = $($EOPPolicyFilter.Priority |Out-String).trim()
        }
        else {
            $EOPPolicyfilterPriority = "Not Configured"
        }
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "Priority"
            "Value"                      = $EOPPolicyfilterPriority
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject

        If(([string]::isnullorempty($EOPPolicyFilter.Description)) -EQ $FALSE) {
            $EOPPolicyfilterDescription = $($EOPPolicyFilter.Description |Out-String).trim()
        }
        else {
            $EOPPolicyfilterDescription = "Not Configured"
        }
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "Description"
            "Value"                      = $EOPPolicyfilterDescription
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject

        If(([string]::isnullorempty($EOPPolicyFilter.State)) -EQ $FALSE) {
            $EOPPolicyfilterState = $($EOPPolicyFilter.State |Out-String).trim()
        }
        else {
            $EOPPolicyfilterState = "Not Configured"
        }
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "State"
            "Value"                      = $EOPPolicyfilterState
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject

        If(([string]::isnullorempty($EOPPolicyFilter.mode)) -EQ $FALSE) {
            $EOPPolicyfiltermode = $($EOPPolicyFilter.mode |Out-String).trim()
        }
        else {
            $EOPPolicyfiltermode = "Not Configured"
        }
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "Mode"
            "Value"                      = $EOPPolicyfiltermode
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject
        If (([string]::isnullorempty($EOPPolicyFilter.From)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "From: " + $EOPPolicyFilter.From
        }
        If (([string]::isnullorempty($EOPPolicyFilter.FromMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "From member of: " + $EOPPolicyFilter.FromMemberOf
        }
        If (([string]::isnullorempty($EOPPolicyFilter.FromScope)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "From scope: " + $EOPPolicyFilter.FromScope
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SentTo)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sent to: " + $EOPPolicyFilter.SentTo
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SentToMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sent to member of: " + $EOPPolicyFilter.SentToMemberOf
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SentToScope)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sent to scope: " + $EOPPolicyFilter.SentToScope
        }
        If (([string]::isnullorempty($EOPPolicyFilter.BetweenMemberOf1)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Between member of 1: " + $EOPPolicyFilter.BetweenMemberOf1
        }
        If (([string]::isnullorempty($EOPPolicyFilter.BetweenMemberOf2)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Between member of 2: " + $EOPPolicyFilter.BetweenMemberOf2
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ManagerAddresses)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Manager addresses: " + $EOPPolicyFilter.ManagerAddresses
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ManagerForEvaluatedUser)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Manager for evaluated user: " + $EOPPolicyFilter.ManagerForEvaluatedUser
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SenderManagementRelationship )) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sender management relationship: " + $EOPPolicyFilter.SenderManagementRelationship 
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ADComparisonAttribute)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "AD comparison attribute: " + $EOPPolicyFilter.ADComparisonAttribute
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ADComparisonOperator)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "AD comparison operator: " + $EOPPolicyFilter.ADComparisonOperator
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SenderADAttributeContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sender AD attribute contains words: " + $EOPPolicyFilter.SenderADAttributeContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SenderADAttributeMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sender AD attribute matches patterns: " + $EOPPolicyFilter.SenderADAttributeMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RecipientADAttributeContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Recipient AD attribute contains words: " + $EOPPolicyFilter.RecipientADAttributeContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RecipientADAttributeMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Recipient AD attribute matches patterns: " + $EOPPolicyFilter.RecipientADAttributeMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfToHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of To header: " + $EOPPolicyFilter.AnyOfToHeader
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfToHeaderMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of To header member of: " + $EOPPolicyFilter.AnyOfToHeaderMemberOf
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfCcHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of CC header: " + $EOPPolicyFilter.AnyOfCcHeader
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfCcHeaderMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of CC header member of: " + $EOPPolicyFilter.AnyOfCcHeaderMemberOf
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfToCcHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of To/CC header: " + $EOPPolicyFilter.AnyOfToCcHeader
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfToCcHeaderMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of To/CC header member of: " + $EOPPolicyFilter.AnyOfToCcHeaderMemberOf
        }
        If (([string]::isnullorempty($EOPPolicyFilter.HasClassification)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Has classification: " + $EOPPolicyFilter.HasClassification
        }
        If (([string]::isnullorempty($EOPPolicyFilter.HasNoClassification)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Has no classification: " + $EOPPolicyFilter.HasNoClassification
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SubjectContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Subject contains words: " + $EOPPolicyFilter.SubjectContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SubjectOrBodyContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Subject or body contains words: " + $EOPPolicyFilter.SubjectOrBodyContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.HeaderContainsMessageHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Header contains message header: " + $EOPPolicyFilter.HeaderContainsMessageHeader
        }
        If (([string]::isnullorempty($EOPPolicyFilter.HeaderContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Header contains words: " + $EOPPolicyFilter.HeaderContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.FromAddressContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "From address contains words: " + $EOPPolicyFilter.FromAddressContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SenderDomainIs)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sender domain is: " + $EOPPolicyFilter.SenderDomainIs
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RecipientDomainIs)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Recipient domain is: " + $EOPPolicyFilter.RecipientDomainIs
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SubjectMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Subject matches patterns: " + $EOPPolicyFilter.SubjectMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SubjectOrBodyMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Subject or body matches patterns: " + $EOPPolicyFilter.SubjectOrBodyMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.HeaderMatchesMessageHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Header matches message header: " + $EOPPolicyFilter.HeaderMatchesMessageHeader
        }
        If (([string]::isnullorempty($EOPPolicyFilter.HeaderMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Header matches patterns: " + $EOPPolicyFilter.HeaderMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.FromAddressMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "From address matches patterns: " + $EOPPolicyFilter.FromAddressMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentNameMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment name matches patterns: " + $EOPPolicyFilter.AttachmentNameMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentExtensionMatchesWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment extension matches words: " + $EOPPolicyFilter.AttachmentExtensionMatchesWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentPropertyContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment property contains words: " + $EOPPolicyFilter.AttachmentPropertyContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ContentCharacterSetContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Content characters set contains words: " + $EOPPolicyFilter.ContentCharacterSetContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.HasSenderOverride)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Has sender override: " + $EOPPolicyFilter.HasSenderOverride
        }
        If (([string]::isnullorempty($EOPPolicyFilter.MessageContainsDataClassifications)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Message contains data classifications: " + $EOPPolicyFilter.MessageContainsDataClassifications
        }
        If (([string]::isnullorempty($EOPPolicyFilter.messageContainsAllDataClassifications)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Message contains all data classifications: " + $EOPPolicyFilter.messageContainsAllDataClassifications
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SenderIpRanges)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sender IP ranges: " + $EOPPolicyFilter.SenderIpRanges
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SCLOver)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "SCL over: " + $EOPPolicyFilter.SCLOver
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentSizeOver)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment size over: " + $EOPPolicyFilter.AttachmentSizeOver
        }
        If (([string]::isnullorempty($EOPPolicyFilter.MessageSizeOver)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Message size over: " + $EOPPolicyFilter.MessageSizeOver
        }
        If (([string]::isnullorempty($EOPPolicyFilter.WithImportance)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "With importance: " + $EOPPolicyFilter.WithImportance
        }
        If (([string]::isnullorempty($EOPPolicyFilter.MessageTypeMatches)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Message type matches: " + $EOPPolicyFilter.MessageTypeMatches
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RecipientAddressContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Recipient address contains words: " + $EOPPolicyFilter.RecipientAddressContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RecipientAddressMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Recipient address matches patterns: " + $EOPPolicyFilter.RecipientAddressMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SenderInRecipientList)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Sender in recipient list: " + $EOPPolicyFilter.SenderInRecipientList
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RecipientInSenderList)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Recipient in sender list: " + $EOPPolicyFilter.RecipientInSenderList
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment contains words: " + $EOPPolicyFilter.AttachmentContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment matches patterns: " + $EOPPolicyFilter.AttachmentMatchesPatterns
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentIsUnsupported)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment is unsupported: " + $EOPPolicyFilter.AttachmentIsUnsupported
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentProcessingLimitExceeded)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment processing limit exceeded: " + $EOPPolicyFilter.AttachmentProcessingLimitExceeded
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentHasExecutableContent)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment has executable content: " + $EOPPolicyFilter.AttachmentHasExecutableContent
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AttachmentIsPasswordProtected)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Attachment is password protected: " + $EOPPolicyFilter.AttachmentIsPasswordProtected
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfRecipientAddressContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of recipient address contains words: " + $EOPPolicyFilter.AnyOfRecipientAddressContainsWords
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AnyOfRecipientAddressMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringConditionsArray += "Any of recipipent address matches patterns: " + $EOPPolicyFilter.AnyOfRecipientAddressMatchesPatterns
        }
        $EOPPolicyFilteringConditionsArray =[string]$($EOPPolicyFilteringConditionsArray|Out-String).trim()
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "Condition"
            "Value"                      = $EOPPolicyFilteringConditionsArray
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject

        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfFrom)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if From: " + $EOPPolicyFilter.ExceptIfFrom 
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfFromMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if From Member Of: " + $EOPPolicyFilter.ExceptIfFromMemberOf
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfFromScope)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if From Scope: " + $EOPPolicyFilter.ExceptIfFromScope  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSentTo)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sent to: " + $EOPPolicyFilter.ExceptIfSentTo  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSentToMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sent to Member Of: " + $EOPPolicyFilter.ExceptIfSentToMemberOf  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSentToScope)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sent to scope: " + $EOPPolicyFilter.ExceptIfSentToScope  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSentToMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sent to Member Of: " + $EOPPolicyFilter.ExceptIfSentToMemberOf  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfBetweenMemberOf1)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if between member Of 1: " + $EOPPolicyFilter.ExceptIfBetweenMemberOf1  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfBetweenMemberOf2)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if between member Of 2: " + $EOPPolicyFilter.ExceptIfBetweenMemberOf2  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfManagerAddresses)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if manager addresses: " + $EOPPolicyFilter.ExceptIfManagerAddresses  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfManagerForEvaluatedUser )) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if manager for evaluated user: " + $EOPPolicyFilter.ExceptIfManagerForEvaluatedUser   
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSenderManagementRelationship)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sender management relationship: " + $EOPPolicyFilter.ExceptIfSenderManagementRelationship  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfADComparisonAttribute)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if AD Comparison Attribute: " + $EOPPolicyFilter.ExceptIfADComparisonAttribute  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfADComparisonOperator)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if AD Comparison Operator: " + $EOPPolicyFilter.ExceptIfADComparisonOperator  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSenderADAttributeContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sender AD attribute contains words: " + $EOPPolicyFilter.ExceptIfSenderADAttributeContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfRecipientADAttributeMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if recipient AD attribute matches patterns: " + $EOPPolicyFilter.ExceptIfRecipientADAttributeMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfToHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of to header: " + $EOPPolicyFilter.ExceptIfAnyOfToHeader  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfToHeaderMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of to header member of: " + $EOPPolicyFilter.ExceptIfAnyOfToHeaderMemberOf  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfCcHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of CC header: " + $EOPPolicyFilter.ExceptIfAnyOfCcHeader  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfCcHeaderMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of CC header Member of: " + $EOPPolicyFilter.ExceptIfAnyOfCcHeaderMemberOf  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfToCcHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of CC/To header: " + $EOPPolicyFilter.ExceptIfAnyOfToCcHeader  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfToCcHeaderMemberOf)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of CC/TO header Member of: " + $EOPPolicyFilter.ExceptIfAnyOfToCcHeaderMemberOf  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfHasClassification)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if has classification: " + $EOPPolicyFilter.ExceptIfHasClassification  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfHasNoClassification)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if has no classification: " + $EOPPolicyFilter.ExceptIfHasNoClassification  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSubjectContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if subject contains words: " + $EOPPolicyFilter.ExceptIfSubjectContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSubjectOrBodyContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if subject or body contains words: " + $EOPPolicyFilter.ExceptIfSubjectOrBodyContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfHeaderContainsMessageHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if header contains message header: " + $EOPPolicyFilter.ExceptIfHeaderContainsMessageHeader  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfHeaderContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if header contains words: " + $EOPPolicyFilter.ExceptIfHeaderContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfFromAddressContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if from address contains words: " + $EOPPolicyFilter.ExceptIfFromAddressContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSenderDomainIs)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sender domain is: " + $EOPPolicyFilter.ExceptIfSenderDomainIs  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfRecipientDomainIs)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if recipient domain is: " + $EOPPolicyFilter.ExceptIfRecipientDomainIs  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSubjectMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if subject matches patterns: " + $EOPPolicyFilter.ExceptIfSubjectMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSubjectOrBodyMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if subject or body matches patterns: " + $EOPPolicyFilter.ExceptIfSubjectOrBodyMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfHeaderMatchesMessageHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if header matches message header: " + $EOPPolicyFilter.ExceptIfHeaderMatchesMessageHeader  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfHeaderMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if header matches patterns: " + $EOPPolicyFilter.ExceptIfHeaderMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfFromAddressMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if from address matches patterns: " + $EOPPolicyFilter.ExceptIfFromAddressMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentNameMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment name matches patterns: " + $EOPPolicyFilter.ExceptIfAttachmentNameMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentExtensionMatchesWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment extension matches words: " + $EOPPolicyFilter.ExceptIfAttachmentExtensionMatchesWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentPropertyContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment property contains words: " + $EOPPolicyFilter.ExceptIfAttachmentPropertyContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfContentCharacterSetContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if content character set contains words: " + $EOPPolicyFilter.ExceptIfContentCharacterSetContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSCLOver)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if SCL over: " + $EOPPolicyFilter.ExceptIfSCLOver  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentSizeOver)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment size over: " + $EOPPolicyFilter.ExceptIfAttachmentSizeOver  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfMessageSizeOver)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if message size over: " + $EOPPolicyFilter.ExceptIfMessageSizeOver  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfWithImportance)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if with importance: " + $EOPPolicyFilter.ExceptIfWithImportance  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfMessageTypeMatches)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if message type matches: " + $EOPPolicyFilter.ExceptIfMessageTypeMatches  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfRecipientAddressContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if recipient address contains words: " + $EOPPolicyFilter.ExceptIfRecipientAddressContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfRecipientAddressMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if recipient address matches patterns: " + $EOPPolicyFilter.ExceptIfRecipientAddressMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSenderInRecipientList)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sender in recipient list: " + $EOPPolicyFilter.ExceptIfSenderInRecipientList  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfRecipientInSenderList)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if recipient in sender list: " + $EOPPolicyFilter.ExceptIfRecipientInSenderList  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment contains words: " + $EOPPolicyFilter.ExceptIfAttachmentContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment matches patterns: " + $EOPPolicyFilter.ExceptIfAttachmentMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentIsUnsupported)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment is unsupported: " + $EOPPolicyFilter.ExceptIfAttachmentIsUnsupported  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentProcessingLimitExceeded)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment processing limit exceeded: " + $EOPPolicyFilter.ExceptIfAttachmentProcessingLimitExceeded  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentHasExecutableContent)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment has executable content: " + $EOPPolicyFilter.ExceptIfAttachmentHasExecutableContent  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAttachmentIsPasswordProtected)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if attachment is password protected: " + $EOPPolicyFilter.ExceptIfAttachmentIsPasswordProtected  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfRecipientAddressContainsWords)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of recipient address contains words: " + $EOPPolicyFilter.ExceptIfAnyOfRecipientAddressContainsWords  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfAnyOfRecipientAddressMatchesPatterns)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if any of recipient addresses matches patterns: " + $EOPPolicyFilter.ExceptIfAnyOfRecipientAddressMatchesPatterns  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfHasSenderOverride)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if has sender override: " + $EOPPolicyFilter.ExceptIfHasSenderOverride  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfMessageContainsDataClassifications)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if Message contains data classifications: " + $EOPPolicyFilter.ExceptIfMessageContainsDataClassifications  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfMessageContainsAllDataClassifications)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if message contains all data classifications: " + $EOPPolicyFilter.ExceptIfMessageContainsAllDataClassifications  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ExceptIfSenderIpRanges)) -EQ $FALSE) {
            $EOPPolicyFilteringExceptArray += "Except if sender IP ranges: " + $EOPPolicyFilter.ExceptIfSenderIpRanges  
        }
        $EOPPolicyFilteringExceptArray =[string]$($EOPPolicyFilteringExceptArray|Out-String).trim()
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "Exception"
            "Value"                      = $EOPPolicyFilteringExceptArray
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject
        If (([string]::isnullorempty($EOPPolicyFilter.PrependSubject )) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Prepend Subject: " + $EOPPolicyFilter.PrependSubject   
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SetAuditSeverity)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Set audit severity: " + $EOPPolicyFilter.SetAuditSeverity  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ApplyClassification)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Apply classification: " + $EOPPolicyFilter.ApplyClassification  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ApplyHtmlDisclaimerLocation)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Apply HTML disclaimer location: " + $EOPPolicyFilter.ApplyHtmlDisclaimerLocation  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ApplyHtmlDisclaimerText)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Apply HTML disclaimer text: " + $EOPPolicyFilter.ApplyHtmlDisclaimerText  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ApplyHtmlDisclaimerFallbackAction)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Apply HTML disclaimer fall back action: " + $EOPPolicyFilter.ApplyHtmlDisclaimerFallbackAction  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ApplyRightsProtectionTemplate)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Apply rights protection template: " + $EOPPolicyFilter.ApplyRightsProtectionTemplate  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ApplyRightsProtectionCustomizationTemplate)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Apply rights protection customization template: " + $EOPPolicyFilter.ApplyRightsProtectionCustomizationTemplate  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SetSCL)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Set SCL: " + $EOPPolicyFilter.SetSCL  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SetHeaderName)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Set header name: " + $EOPPolicyFilter.SetHeaderName  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SetHeaderValue)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Set header value: " + $EOPPolicyFilter.SetHeaderValue   
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RemoveHeader)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Remove header: " + $EOPPolicyFilter.RemoveHeader  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AddToRecipients)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Add to recipients: " + $EOPPolicyFilter.AddToRecipients  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.CopyTo)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Copy to: " + $EOPPolicyFilter.CopyTo  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.BlindCopyTo)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Blind copy to: " + $EOPPolicyFilter.BlindCopyTo  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.AddManagerAsRecipientType)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Add manager as recipient type: " + $EOPPolicyFilter.AddManagerAsRecipientType  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ModerateMessageByUser)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Moderate message by user: " + $EOPPolicyFilter.ModerateMessageByUser  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ModerateMessageByManager)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Moderate message by manager: " + $EOPPolicyFilter.ModerateMessageByManager  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RedirectMessageTo)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Redirect message to: " + $EOPPolicyFilter.RedirectMessageTo  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RejectMessageEnhancedStatusCode)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Reject message enhanced status code: " + $EOPPolicyFilter.RejectMessageEnhancedStatusCode  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RejectMessageReasonText)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Reject message reason text: " + $EOPPolicyFilter.RejectMessageReasonText  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.Disconnect)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "disconnect: " + $EOPPolicyFilter.Disconnect  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.DeleteMessage)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Delete message: " + $EOPPolicyFilter.DeleteMessage  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.Quarantine)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Quarantine: " + $EOPPolicyFilter.Quarantine  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SmtpRejectMessageRejectText)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "SMTP reject message reject text: " + $EOPPolicyFilter.SmtpRejectMessageRejectText  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SmtpRejectMessageRejectStatusCode)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "SMTP reject message reject status code: " + $EOPPolicyFilter.SmtpRejectMessageRejectStatusCode  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.LogEventText)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Log event text: " + $EOPPolicyFilter.LogEventText  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.StopRuleProcessing)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Stop rule processing: " + $EOPPolicyFilter.StopRuleProcessing  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.SenderNotificationType)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Sender notification type: " + $EOPPolicyFilter.SenderNotificationType  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.GenerateIncidentReport)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Generate incident report: " + $EOPPolicyFilter.GenerateIncidentReport  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.IncidentReportContent)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Incident report content: " + $EOPPolicyFilter.IncidentReportContent  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RouteMessageOutboundConnector)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Route message outbound connector : " + $EOPPolicyFilter.RouteMessageOutboundConnector  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RouteMessageOutboundRequireTls)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Route message outbound require TLS: " + $EOPPolicyFilter.RouteMessageOutboundRequireTls  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.ApplyOME)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Apply OME: " + $EOPPolicyFilter.ApplyOME  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RemoveOME)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Remove OME: " + $EOPPolicyFilter.RemoveOME  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.RemoveOMEv2)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Remove OMEv2: " + $EOPPolicyFilter.RemoveOMEv2  
        }
        If (([string]::isnullorempty($EOPPolicyFilter.GenerateNotification)) -EQ $FALSE) {
            $EOPPolicyFilteringEactionArray += "Generate notification: " + $EOPPolicyFilter.GenerateNotification  
        }

        $EOPPolicyFilteringEactionArray =[string]$($EOPPolicyFilteringEactionArray|Out-String).trim()
        $EOPPolicyFilterDetail = [ordered]@{
            "Configuration Item"         = "Action"
            "Value"                      = $EOPPolicyFilteringEactionArray
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
        $EOPPolicyFilterArray += $EOConfigurationObject
    }
    Else {
        $EOPPolicyFilter |ForEach-Object {
            If(([string]::isnullorempty($_.identity)) -EQ $FALSE) {
                $EOPPolicyfilterName =  $($_.identity |Out-String).trim()
            }
            else {
                $EOPPolicyfilterName = "Not Configured"
            }
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "Name [TBA]"
                "Value"                      = $EOPPolicyfilterName
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject
            
            If(([string]::isnullorempty( $_.Priority)) -EQ $FALSE) {
                $EOPPolicyfilterPriority =  $($_.Priority |Out-String).trim()
            }
            else {
                $EOPPolicyfilterPriority = "Not Configured"
            }
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "Priority"
                "Value"                      = $EOPPolicyfilterPriority
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject
    
            If(([string]::isnullorempty( $_.Description)) -EQ $FALSE) {
                $EOPPolicyfilterDescription =  $($_.Description |Out-String).trim()
            }
            else {
                $EOPPolicyfilterDescription = "Not Configured"
            }
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "Description"
                "Value"                      = $EOPPolicyfilterDescription
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject
    
            If(([string]::isnullorempty( $_.State)) -EQ $FALSE) {
                $EOPPolicyfilterState =  $($_.State |Out-String).trim()
            }
            else {
                $EOPPolicyfilterState = "Not Configured"
            }
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "State"
                "Value"                      = $EOPPolicyfilterState
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject
    
            If(([string]::isnullorempty( $_.mode)) -EQ $FALSE) {
                $EOPPolicyfiltermode =  $($_.mode |Out-String).trim()
            }
            else {
                $EOPPolicyfiltermode = "Not Configured"
            }
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "Mode"
                "Value"                      = $EOPPolicyfiltermode
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject
            If (([string]::isnullorempty( $_.From)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "From: " +  $_.From
            }
            If (([string]::isnullorempty( $_.FromMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "From member of: " +  $_.FromMemberOf
            }
            If (([string]::isnullorempty( $_.FromScope)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "From scope: " +  $_.FromScope
            }
            If (([string]::isnullorempty( $_.SentTo)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sent to: " +  $_.SentTo
            }
            If (([string]::isnullorempty( $_.SentToMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sent to member of: " +  $_.SentToMemberOf
            }
            If (([string]::isnullorempty( $_.SentToScope)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sent to scope: " +  $_.SentToScope
            }
            If (([string]::isnullorempty( $_.BetweenMemberOf1)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Between member of 1: " +  $_.BetweenMemberOf1
            }
            If (([string]::isnullorempty( $_.BetweenMemberOf2)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Between member of 2: " +  $_.BetweenMemberOf2
            }
            If (([string]::isnullorempty( $_.ManagerAddresses)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Manager addresses: " +  $_.ManagerAddresses
            }
            If (([string]::isnullorempty( $_.ManagerForEvaluatedUser)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Manager for evaluated user: " +  $_.ManagerForEvaluatedUser
            }
            If (([string]::isnullorempty( $_.SenderManagementRelationship )) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sender management relationship: " +  $_.SenderManagementRelationship 
            }
            If (([string]::isnullorempty( $_.ADComparisonAttribute)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "AD comparison attribute: " +  $_.ADComparisonAttribute
            }
            If (([string]::isnullorempty( $_.ADComparisonOperator)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "AD comparison operator: " +  $_.ADComparisonOperator
            }
            If (([string]::isnullorempty( $_.SenderADAttributeContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sender AD attribute contains words: " +  $_.SenderADAttributeContainsWords
            }
            If (([string]::isnullorempty( $_.SenderADAttributeMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sender AD attribute matches patterns: " +  $_.SenderADAttributeMatchesPatterns
            }
            If (([string]::isnullorempty( $_.RecipientADAttributeContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Recipient AD attribute contains words: " +  $_.RecipientADAttributeContainsWords
            }
            If (([string]::isnullorempty( $_.RecipientADAttributeMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Recipient AD attribute matches patterns: " +  $_.RecipientADAttributeMatchesPatterns
            }
            If (([string]::isnullorempty( $_.AnyOfToHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of To header: " +  $_.AnyOfToHeader
            }
            If (([string]::isnullorempty( $_.AnyOfToHeaderMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of To header member of: " +  $_.AnyOfToHeaderMemberOf
            }
            If (([string]::isnullorempty( $_.AnyOfCcHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of CC header: " +  $_.AnyOfCcHeader
            }
            If (([string]::isnullorempty( $_.AnyOfCcHeaderMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of CC header member of: " +  $_.AnyOfCcHeaderMemberOf
            }
            If (([string]::isnullorempty( $_.AnyOfToCcHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of To/CC header: " +  $_.AnyOfToCcHeader
            }
            If (([string]::isnullorempty( $_.AnyOfToCcHeaderMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of To/CC header member of: " +  $_.AnyOfToCcHeaderMemberOf
            }
            If (([string]::isnullorempty( $_.HasClassification)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Has classification: " +  $_.HasClassification
            }
            If (([string]::isnullorempty( $_.HasNoClassification)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Has no classification: " +  $_.HasNoClassification
            }
            If (([string]::isnullorempty( $_.SubjectContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Subject contains words: " +  $_.SubjectContainsWords
            }
            If (([string]::isnullorempty( $_.SubjectOrBodyContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Subject or body contains words: " +  $_.SubjectOrBodyContainsWords
            }
            If (([string]::isnullorempty( $_.HeaderContainsMessageHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Header contains message header: " +  $_.HeaderContainsMessageHeader
            }
            If (([string]::isnullorempty( $_.HeaderContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Header contains words: " +  $_.HeaderContainsWords
            }
            If (([string]::isnullorempty( $_.FromAddressContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "From address contains words: " +  $_.FromAddressContainsWords
            }
            If (([string]::isnullorempty( $_.SenderDomainIs)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sender domain is: " +  $_.SenderDomainIs
            }
            If (([string]::isnullorempty( $_.RecipientDomainIs)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Recipient domain is: " +  $_.RecipientDomainIs
            }
            If (([string]::isnullorempty( $_.SubjectMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Subject matches patterns: " +  $_.SubjectMatchesPatterns
            }
            If (([string]::isnullorempty( $_.SubjectOrBodyMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Subject or body matches patterns: " +  $_.SubjectOrBodyMatchesPatterns
            }
            If (([string]::isnullorempty( $_.HeaderMatchesMessageHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Header matches message header: " +  $_.HeaderMatchesMessageHeader
            }
            If (([string]::isnullorempty( $_.HeaderMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Header matches patterns: " +  $_.HeaderMatchesPatterns
            }
            If (([string]::isnullorempty( $_.FromAddressMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "From address matches patterns: " +  $_.FromAddressMatchesPatterns
            }
            If (([string]::isnullorempty( $_.AttachmentNameMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment name matches patterns: " +  $_.AttachmentNameMatchesPatterns
            }
            If (([string]::isnullorempty( $_.AttachmentExtensionMatchesWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment extension matches words: " +  $_.AttachmentExtensionMatchesWords
            }
            If (([string]::isnullorempty( $_.AttachmentPropertyContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment property contains words: " +  $_.AttachmentPropertyContainsWords
            }
            If (([string]::isnullorempty( $_.ContentCharacterSetContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Content characters set contains words: " +  $_.ContentCharacterSetContainsWords
            }
            If (([string]::isnullorempty( $_.HasSenderOverride)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Has sender override: " +  $_.HasSenderOverride
            }
            If (([string]::isnullorempty( $_.MessageContainsDataClassifications)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Message contains data classifications: " +  $_.MessageContainsDataClassifications
            }
            If (([string]::isnullorempty( $_.messageContainsAllDataClassifications)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Message contains all data classifications: " +  $_.messageContainsAllDataClassifications
            }
            If (([string]::isnullorempty( $_.SenderIpRanges)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sender IP ranges: " +  $_.SenderIpRanges
            }
            If (([string]::isnullorempty( $_.SCLOver)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "SCL over: " +  $_.SCLOver
            }
            If (([string]::isnullorempty( $_.AttachmentSizeOver)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment size over: " +  $_.AttachmentSizeOver
            }
            If (([string]::isnullorempty( $_.MessageSizeOver)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Message size over: " +  $_.MessageSizeOver
            }
            If (([string]::isnullorempty( $_.WithImportance)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "With importance: " +  $_.WithImportance
            }
            If (([string]::isnullorempty( $_.MessageTypeMatches)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Message type matches: " +  $_.MessageTypeMatches
            }
            If (([string]::isnullorempty( $_.RecipientAddressContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Recipient address contains words: " +  $_.RecipientAddressContainsWords
            }
            If (([string]::isnullorempty( $_.RecipientAddressMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Recipient address matches patterns: " +  $_.RecipientAddressMatchesPatterns
            }
            If (([string]::isnullorempty( $_.SenderInRecipientList)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Sender in recipient list: " +  $_.SenderInRecipientList
            }
            If (([string]::isnullorempty( $_.RecipientInSenderList)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Recipient in sender list: " +  $_.RecipientInSenderList
            }
            If (([string]::isnullorempty( $_.AttachmentContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment contains words: " +  $_.AttachmentContainsWords
            }
            If (([string]::isnullorempty( $_.AttachmentMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment matches patterns: " +  $_.AttachmentMatchesPatterns
            }
            If (([string]::isnullorempty( $_.AttachmentIsUnsupported)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment is unsupported: " +  $_.AttachmentIsUnsupported
            }
            If (([string]::isnullorempty( $_.AttachmentProcessingLimitExceeded)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment processing limit exceeded: " +  $_.AttachmentProcessingLimitExceeded
            }
            If (([string]::isnullorempty( $_.AttachmentHasExecutableContent)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment has executable content: " +  $_.AttachmentHasExecutableContent
            }
            If (([string]::isnullorempty( $_.AttachmentIsPasswordProtected)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Attachment is password protected: " +  $_.AttachmentIsPasswordProtected
            }
            If (([string]::isnullorempty( $_.AnyOfRecipientAddressContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of recipient address contains words: " +  $_.AnyOfRecipientAddressContainsWords
            }
            If (([string]::isnullorempty( $_.AnyOfRecipientAddressMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringConditionsArray += "Any of recipipent address matches patterns: " +  $_.AnyOfRecipientAddressMatchesPatterns
            }
            $EOPPolicyFilteringConditionsArray =[string]$($EOPPolicyFilteringConditionsArray|Out-String).trim()
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "Condition"
                "Value"                      = $EOPPolicyFilteringConditionsArray
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject
    
            If (([string]::isnullorempty( $_.ExceptIfFrom)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if From: " +  $_.ExceptIfFrom 
            }
            If (([string]::isnullorempty( $_.ExceptIfFromMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if From Member Of: " +  $_.ExceptIfFromMemberOf
            }
            If (([string]::isnullorempty( $_.ExceptIfFromScope)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if From Scope: " +  $_.ExceptIfFromScope  
            }
            If (([string]::isnullorempty( $_.ExceptIfSentTo)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sent to: " +  $_.ExceptIfSentTo  
            }
            If (([string]::isnullorempty( $_.ExceptIfSentToMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sent to Member Of: " +  $_.ExceptIfSentToMemberOf  
            }
            If (([string]::isnullorempty( $_.ExceptIfSentToScope)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sent to scope: " +  $_.ExceptIfSentToScope  
            }
            If (([string]::isnullorempty( $_.ExceptIfSentToMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sent to Member Of: " +  $_.ExceptIfSentToMemberOf  
            }
            If (([string]::isnullorempty( $_.ExceptIfBetweenMemberOf1)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if between member Of 1: " +  $_.ExceptIfBetweenMemberOf1  
            }
            If (([string]::isnullorempty( $_.ExceptIfBetweenMemberOf2)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if between member Of 2: " +  $_.ExceptIfBetweenMemberOf2  
            }
            If (([string]::isnullorempty( $_.ExceptIfManagerAddresses)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if manager addresses: " +  $_.ExceptIfManagerAddresses  
            }
            If (([string]::isnullorempty( $_.ExceptIfManagerForEvaluatedUser )) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if manager for evaluated user: " +  $_.ExceptIfManagerForEvaluatedUser   
            }
            If (([string]::isnullorempty( $_.ExceptIfSenderManagementRelationship)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sender management relationship: " +  $_.ExceptIfSenderManagementRelationship  
            }
            If (([string]::isnullorempty( $_.ExceptIfADComparisonAttribute)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if AD Comparison Attribute: " +  $_.ExceptIfADComparisonAttribute  
            }
            If (([string]::isnullorempty( $_.ExceptIfADComparisonOperator)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if AD Comparison Operator: " +  $_.ExceptIfADComparisonOperator  
            }
            If (([string]::isnullorempty( $_.ExceptIfSenderADAttributeContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sender AD attribute contains words: " +  $_.ExceptIfSenderADAttributeContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfRecipientADAttributeMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if recipient AD attribute matches patterns: " +  $_.ExceptIfRecipientADAttributeMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfToHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of to header: " +  $_.ExceptIfAnyOfToHeader  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfToHeaderMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of to header member of: " +  $_.ExceptIfAnyOfToHeaderMemberOf  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfCcHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of CC header: " +  $_.ExceptIfAnyOfCcHeader  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfCcHeaderMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of CC header Member of: " +  $_.ExceptIfAnyOfCcHeaderMemberOf  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfToCcHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of CC/To header: " +  $_.ExceptIfAnyOfToCcHeader  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfToCcHeaderMemberOf)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of CC/TO header Member of: " +  $_.ExceptIfAnyOfToCcHeaderMemberOf  
            }
            If (([string]::isnullorempty( $_.ExceptIfHasClassification)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if has classification: " +  $_.ExceptIfHasClassification  
            }
            If (([string]::isnullorempty( $_.ExceptIfHasNoClassification)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if has no classification: " +  $_.ExceptIfHasNoClassification  
            }
            If (([string]::isnullorempty( $_.ExceptIfSubjectContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if subject contains words: " +  $_.ExceptIfSubjectContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfSubjectOrBodyContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if subject or body contains words: " +  $_.ExceptIfSubjectOrBodyContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfHeaderContainsMessageHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if header contains message header: " +  $_.ExceptIfHeaderContainsMessageHeader  
            }
            If (([string]::isnullorempty( $_.ExceptIfHeaderContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if header contains words: " +  $_.ExceptIfHeaderContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfFromAddressContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if from address contains words: " +  $_.ExceptIfFromAddressContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfSenderDomainIs)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sender domain is: " +  $_.ExceptIfSenderDomainIs  
            }
            If (([string]::isnullorempty( $_.ExceptIfRecipientDomainIs)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if recipient domain is: " +  $_.ExceptIfRecipientDomainIs  
            }
            If (([string]::isnullorempty( $_.ExceptIfSubjectMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if subject matches patterns: " +  $_.ExceptIfSubjectMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfSubjectOrBodyMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if subject or body matches patterns: " +  $_.ExceptIfSubjectOrBodyMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfHeaderMatchesMessageHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if header matches message header: " +  $_.ExceptIfHeaderMatchesMessageHeader  
            }
            If (([string]::isnullorempty( $_.ExceptIfHeaderMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if header matches patterns: " +  $_.ExceptIfHeaderMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfFromAddressMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if from address matches patterns: " +  $_.ExceptIfFromAddressMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentNameMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment name matches patterns: " +  $_.ExceptIfAttachmentNameMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentExtensionMatchesWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment extension matches words: " +  $_.ExceptIfAttachmentExtensionMatchesWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentPropertyContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment property contains words: " +  $_.ExceptIfAttachmentPropertyContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfContentCharacterSetContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if content character set contains words: " +  $_.ExceptIfContentCharacterSetContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfSCLOver)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if SCL over: " +  $_.ExceptIfSCLOver  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentSizeOver)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment size over: " +  $_.ExceptIfAttachmentSizeOver  
            }
            If (([string]::isnullorempty( $_.ExceptIfMessageSizeOver)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if message size over: " +  $_.ExceptIfMessageSizeOver  
            }
            If (([string]::isnullorempty( $_.ExceptIfWithImportance)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if with importance: " +  $_.ExceptIfWithImportance  
            }
            If (([string]::isnullorempty( $_.ExceptIfMessageTypeMatches)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if message type matches: " +  $_.ExceptIfMessageTypeMatches  
            }
            If (([string]::isnullorempty( $_.ExceptIfRecipientAddressContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if recipient address contains words: " +  $_.ExceptIfRecipientAddressContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfRecipientAddressMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if recipient address matches patterns: " +  $_.ExceptIfRecipientAddressMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfSenderInRecipientList)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sender in recipient list: " +  $_.ExceptIfSenderInRecipientList  
            }
            If (([string]::isnullorempty( $_.ExceptIfRecipientInSenderList)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if recipient in sender list: " +  $_.ExceptIfRecipientInSenderList  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment contains words: " +  $_.ExceptIfAttachmentContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment matches patterns: " +  $_.ExceptIfAttachmentMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentIsUnsupported)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment is unsupported: " +  $_.ExceptIfAttachmentIsUnsupported  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentProcessingLimitExceeded)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment processing limit exceeded: " +  $_.ExceptIfAttachmentProcessingLimitExceeded  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentHasExecutableContent)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment has executable content: " +  $_.ExceptIfAttachmentHasExecutableContent  
            }
            If (([string]::isnullorempty( $_.ExceptIfAttachmentIsPasswordProtected)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if attachment is password protected: " +  $_.ExceptIfAttachmentIsPasswordProtected  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfRecipientAddressContainsWords)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of recipient address contains words: " +  $_.ExceptIfAnyOfRecipientAddressContainsWords  
            }
            If (([string]::isnullorempty( $_.ExceptIfAnyOfRecipientAddressMatchesPatterns)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if any of recipient addresses matches patterns: " +  $_.ExceptIfAnyOfRecipientAddressMatchesPatterns  
            }
            If (([string]::isnullorempty( $_.ExceptIfHasSenderOverride)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if has sender override: " +  $_.ExceptIfHasSenderOverride  
            }
            If (([string]::isnullorempty( $_.ExceptIfMessageContainsDataClassifications)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if Message contains data classifications: " +  $_.ExceptIfMessageContainsDataClassifications  
            }
            If (([string]::isnullorempty( $_.ExceptIfMessageContainsAllDataClassifications)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if message contains all data classifications: " +  $_.ExceptIfMessageContainsAllDataClassifications  
            }
            If (([string]::isnullorempty( $_.ExceptIfSenderIpRanges)) -EQ $FALSE) {
                $EOPPolicyFilteringExceptArray += "Except if sender IP ranges: " +  $_.ExceptIfSenderIpRanges  
            }
            $EOPPolicyFilteringExceptArray =[string]$($EOPPolicyFilteringExceptArray|Out-String).trim()
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "Exception"
                "Value"                      = $EOPPolicyFilteringExceptArray
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject
            If (([string]::isnullorempty( $_.PrependSubject )) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Prepend Subject: " +  $_.PrependSubject   
            }
            If (([string]::isnullorempty( $_.SetAuditSeverity)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Set audit severity: " +  $_.SetAuditSeverity  
            }
            If (([string]::isnullorempty( $_.ApplyClassification)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Apply classification: " +  $_.ApplyClassification  
            }
            If (([string]::isnullorempty( $_.ApplyHtmlDisclaimerLocation)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Apply HTML disclaimer location: " +  $_.ApplyHtmlDisclaimerLocation  
            }
            If (([string]::isnullorempty( $_.ApplyHtmlDisclaimerText)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Apply HTML disclaimer text: " +  $_.ApplyHtmlDisclaimerText  
            }
            If (([string]::isnullorempty( $_.ApplyHtmlDisclaimerFallbackAction)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Apply HTML disclaimer fall back action: " +  $_.ApplyHtmlDisclaimerFallbackAction  
            }
            If (([string]::isnullorempty( $_.ApplyRightsProtectionTemplate)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Apply rights protection template: " +  $_.ApplyRightsProtectionTemplate  
            }
            If (([string]::isnullorempty( $_.ApplyRightsProtectionCustomizationTemplate)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Apply rights protection customization template: " +  $_.ApplyRightsProtectionCustomizationTemplate  
            }
            If (([string]::isnullorempty( $_.SetSCL)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Set SCL: " +  $_.SetSCL  
            }
            If (([string]::isnullorempty( $_.SetHeaderName)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Set header name: " +  $_.SetHeaderName  
            }
            If (([string]::isnullorempty( $_.SetHeaderValue)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Set header value: " +  $_.SetHeaderValue   
            }
            If (([string]::isnullorempty( $_.RemoveHeader)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Remove header: " +  $_.RemoveHeader  
            }
            If (([string]::isnullorempty( $_.AddToRecipients)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Add to recipients: " +  $_.AddToRecipients  
            }
            If (([string]::isnullorempty( $_.CopyTo)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Copy to: " +  $_.CopyTo  
            }
            If (([string]::isnullorempty( $_.BlindCopyTo)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Blind copy to: " +  $_.BlindCopyTo  
            }
            If (([string]::isnullorempty( $_.AddManagerAsRecipientType)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Add manager as recipient type: " +  $_.AddManagerAsRecipientType  
            }
            If (([string]::isnullorempty( $_.ModerateMessageByUser)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Moderate message by user: " +  $_.ModerateMessageByUser  
            }
            If (([string]::isnullorempty( $_.ModerateMessageByManager)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Moderate message by manager: " +  $_.ModerateMessageByManager  
            }
            If (([string]::isnullorempty( $_.RedirectMessageTo)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Redirect message to: " +  $_.RedirectMessageTo  
            }
            If (([string]::isnullorempty( $_.RejectMessageEnhancedStatusCode)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Reject message enhanced status code: " +  $_.RejectMessageEnhancedStatusCode  
            }
            If (([string]::isnullorempty( $_.RejectMessageReasonText)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Reject message reason text: " +  $_.RejectMessageReasonText  
            }
            If (([string]::isnullorempty( $_.Disconnect)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "disconnect: " +  $_.Disconnect  
            }
            If (([string]::isnullorempty( $_.DeleteMessage)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Delete message: " +  $_.DeleteMessage  
            }
            If (([string]::isnullorempty( $_.Quarantine)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Quarantine: " +  $_.Quarantine  
            }
            If (([string]::isnullorempty( $_.SmtpRejectMessageRejectText)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "SMTP reject message reject text: " +  $_.SmtpRejectMessageRejectText  
            }
            If (([string]::isnullorempty( $_.SmtpRejectMessageRejectStatusCode)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "SMTP reject message reject status code: " +  $_.SmtpRejectMessageRejectStatusCode  
            }
            If (([string]::isnullorempty( $_.LogEventText)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Log event text: " +  $_.LogEventText  
            }
            If (([string]::isnullorempty( $_.StopRuleProcessing)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Stop rule processing: " +  $_.StopRuleProcessing  
            }
            If (([string]::isnullorempty( $_.SenderNotificationType)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Sender notification type: " +  $_.SenderNotificationType  
            }
            If (([string]::isnullorempty( $_.GenerateIncidentReport)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Generate incident report: " +  $_.GenerateIncidentReport  
            }
            If (([string]::isnullorempty( $_.IncidentReportContent)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Incident report content: " +  $_.IncidentReportContent  
            }
            If (([string]::isnullorempty( $_.RouteMessageOutboundConnector)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Route message outbound connector : " +  $_.RouteMessageOutboundConnector  
            }
            If (([string]::isnullorempty( $_.RouteMessageOutboundRequireTls)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Route message outbound require TLS: " +  $_.RouteMessageOutboundRequireTls  
            }
            If (([string]::isnullorempty( $_.ApplyOME)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Apply OME: " +  $_.ApplyOME  
            }
            If (([string]::isnullorempty( $_.RemoveOME)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Remove OME: " +  $_.RemoveOME  
            }
            If (([string]::isnullorempty( $_.RemoveOMEv2)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Remove OMEv2: " +  $_.RemoveOMEv2  
            }
            If (([string]::isnullorempty( $_.GenerateNotification)) -EQ $FALSE) {
                $EOPPolicyFilteringEactionArray += "Generate notification: " +  $_.GenerateNotification  
            }
    
            $EOPPolicyFilteringEactionArray =[string]$($EOPPolicyFilteringEactionArray|Out-String).trim()
            $EOPPolicyFilterDetail = [ordered]@{
                "Configuration Item"         = "Action"
                "Value"                      = $EOPPolicyFilteringEactionArray
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPPolicyFilterDetail
            $EOPPolicyFilterArray += $EOConfigurationObject    
        }
    }
}

#############################################
#Exchange Online Protection - Content Filtering
#############################################
Write-Host " - Content Filtering" -foregroundcolor Gray

$EOPContentFilter = Get-HostedContentFilterPolicy
If ($null -eq  $EOPContentFilter) {
     $EOPContentFilterDetail = [ordered]@{
        "Name"                             = "Not Configured"
        "Spam Action"                      = "N/A"
        "High Confidence Spam Action"      = "N/A"
    }
    $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
    $EOPContentFilterArray += $EOConfigurationObject
}
Else {
    If($EOPContentFilter -isnot [array]) {
        If (([string]::isnullorempty( $EOPContentFilter.Name)) -EQ $FALSE) {
            $EOPContentFilterName = $EOPContentFilter.Name
        }
        else {

            $EOPContentFilterName = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Name"
            "Value"                      = $EOPContentFilterName
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject

        If ($EOPContentFilter.AddxHeaderValue.Length  -gt 1 ) {
            $EOPContentFilterAddXHeaderValue = $EOPContentFilter.AddxHeaderValue
        }
        else {

            $EOPContentFilterAddXHeaderValue = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Add X Header Value"
            "Value"                      = $EOPContentFilterAddXHeaderValue
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject
        
        If  ($EOPContentFilter.ModifySubjectValue.Length  -gt 1 ) {
            $EOPContentFilterModifySubjectValue = $EOPContentFilter.ModifySubjectValue
        }
        else {

            $EOPContentFilterModifySubjectValue = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Modify Subject value"
            "Value"                      = $EOPContentFilterModifySubjectValue
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject
        
        If ($EOPContentFilter.RedirectToRecipients.Length  -gt 1 ) {
            $EOPContentFilterRedirectToRecipients = $EOPContentFilter.RedirectToRecipients |out-string
            $EOPContentFilterRedirectToRecipients = $EOPContentFilterRedirectToRecipients.trim()
        }
        else {

            $EOPContentFilterRedirectToRecipients = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Redirect to recipients"
            "Value"                      = $EOPContentFilterRedirectToRecipients
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If ($EOPContentFilter.FalsePositiveAdditionalRecipients.Length  -gt 1 ) {
            $EOPContentFilterFalsePositiveAdditionalRecipients = $EOPContentFilter.FalsePositiveAdditionalRecipients |out-string
            $EOPContentFilterFalsePositiveAdditionalRecipients =$EOPContentFilterFalsePositiveAdditionalRecipients.Trim()
        }
        else {

            $EOPContentFilterFalsePositiveAdditionalRecipients = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "False positive additional recipients"
            "Value"                      = $EOPContentFilterFalsePositiveAdditionalRecipients
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.QuarantineRetentionPeriod)) -EQ $FALSE) {
            $EOPContentFilterQuarantineRetentionPeriod = $EOPContentFilter.QuarantineRetentionPeriod
        }
        else {

            $EOPContentFilterQuarantineRetentionPeriod = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Quarantine retention period"
            "Value"                      = $EOPContentFilterQuarantineRetentionPeriod
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.EndUserSpamNotificationFrequency)) -EQ $FALSE) {
            $EOPContentFilterEndUserSpamNotificationFrequency = $EOPContentFilter.EndUserSpamNotificationFrequency
        }
        else {

            $EOPContentFilterEndUserSpamNotificationFrequency = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "End user spam notification frequency"
            "Value"                      = $EOPContentFilterEndUserSpamNotificationFrequency
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If(([string]::isnullorempty($EOPContentFilter.IncreaseScoreWithImageLinks)) -EQ $FALSE) {
            $EOPContentFilterIncreaseScoreArray += "Increase score with image links: " + ($EOPContentFilter.IncreaseScoreWithImageLinks)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.IncreaseScoreWithNumericIps)) -EQ $FALSE) {
            $EOPContentFilterIncreaseScoreArray += "Increase score with numeric IPs: " + ($EOPContentFilter.IncreaseScoreWithNumericIps)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.IncreaseScoreWithRedirectToOtherPort)) -EQ $FALSE) {
            $EOPContentFilterIncreaseScoreArray += "Increase score with redirect to other port: " + ($EOPContentFilter.IncreaseScoreWithRedirectToOtherPort)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.IncreaseScoreWithBizOrInfoUrls)) -EQ $FALSE) {
            $EOPContentFilterIncreaseScoreArray += "Increase score with Biz or info URLs: " + ($EOPContentFilter.IncreaseScoreWithBizOrInfoUrls)+"`n"
        }
        $EOPContentFilterIncreaseScoreArray = $([string]$EOPContentFilterIncreaseScoreArray|out-string).trimend()
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Increase Score"
            "Value"                      = $EOPContentFilterIncreaseScoreArray
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject

        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamEmptyMessages)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam empty messages: " + ($EOPContentFilter.MarkAsSpamEmptyMessages)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamJavaScriptInHtml)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam javascript in html: " + ($EOPContentFilter.MarkAsSpamJavaScriptInHtml)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamFramesInHtml)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam frames in HTML: " + ($EOPContentFilter.MarkAsSpamFramesInHtml)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamObjectTagsInHtml)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam object tags in HTML: " + ($EOPContentFilter.MarkAsSpamObjectTagsInHtml)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamEmbedTagsInHtml)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam embed tags in HTML: " + ($EOPContentFilter.MarkAsSpamEmbedTagsInHtml) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamFormTagsInHtml)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam form tags in HTML: " + ($EOPContentFilter.MarkAsSpamFormTagsInHtml) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamWebBugsInHtml)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam web bugs in HTML: " + ($EOPContentFilter.MarkAsSpamWebBugsInHtml) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamSensitiveWordList)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam sensitive word list: " + ($EOPContentFilter.MarkAsSpamSensitiveWordList) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamSpfRecordHardFail)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam SPF record hard fail: " + ($EOPContentFilter.MarkAsSpamSpfRecordHardFail) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamFromAddressAuthFail)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam from address auth fail: " + ($EOPContentFilter.MarkAsSpamFromAddressAuthFail) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamBulkMail)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam bulk mail: " + ($EOPContentFilter.MarkAsSpamBulkMail) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.MarkAsSpamNdrBackscatter)) -EQ $FALSE) {
            $EOPContentFilterMarkAsSpameArray += "Mark as spam NDR backscatter: " + ($EOPContentFilter.MarkAsSpamNdrBackscatter) +"`n"
        }
        $EOPContentFilterMarkAsSpameArray = $([string]$EOPContentFilterMarkAsSpameArray|out-string).trimend()
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Mark as spam"
            "Value"                      = $EOPContentFilterMarkAsSpameArray
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.HighConfidenceSpamAction)) -EQ $FALSE) {
            $EOPContentFilterHighConfidenceSpamAction = $EOPContentFilter.HighConfidenceSpamAction
        }
        else {

            $EOPContentFilterHighConfidenceSpamAction = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "High confidence spam action"
            "Value"                      = $EOPContentFilterHighConfidenceSpamAction
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.SpamAction)) -EQ $FALSE) {
            $EOPContentFilterSpamAction = $EOPContentFilter.SpamAction
        }
        else {

            $EOPContentFilterSpamAction = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Spam action"
            "Value"                      = $EOPContentFilterSpamAction
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.BulkSpamAction)) -EQ $FALSE) {
            $EOPContentBulkSpamAction = $EOPContentFilter.BulkSpamAction
        }
        else {

            $EOPContentBulkSpamAction = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Bulk spam action"
            "Value"                      = $EOPContentBulkSpamAction
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.PhishSpamAction)) -EQ $FALSE) {
            $EOPContentFilterPhishSpamAction = $EOPContentFilter.PhishSpamAction
        }
        else {

            $EOPContentFilterPhishSpamAction = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Phish Spam action"
            "Value"                      = $EOPContentFilterPhishSpamAction
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.EnableEndUserSpamNotifications)) -EQ $FALSE) {
            $EOPContentFilterEnableEndUserSpamNotifications = $EOPContentFilter.EnableEndUserSpamNotifications
        }
        else {

            $EOPContentFilterEnableEndUserSpamNotifications = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Enable end user spam notifications"
            "Value"                      = $EOPContentFilterEnableEndUserSpamNotifications
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If(([string]::isnullorempty($EOPContentFilter.EndUserSpamNotificationCustomFromAddress)) -EQ $FALSE) {
            $EOPContentFilterEndUserSpamNotificationArray += "Notification custom from address: " + ($EOPContentFilter.EndUserSpamNotificationCustomFromAddress) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.EndUserSpamNotificationCustomFromName)) -EQ $FALSE) {
            $EOPContentFilterEndUserSpamNotificationArray += "Notification custom from name: " + ($EOPContentFilter.EndUserSpamNotificationCustomFromName)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.EndUserSpamNotificationCustomSubject)) -EQ $FALSE) {
            $EOPContentFilterEndUserSpamNotificationArray += "Notification custom from subject: " + ($EOPContentFilter.EndUserSpamNotificationCustomSubject) +"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.EndUserSpamNotificationLanguage)) -EQ $FALSE) {
            $EOPContentFilterEndUserSpamNotificationArray += "Notification language: " + ($EOPContentFilter.EndUserSpamNotificationLanguage)+"`n"
        }
        If(([string]::isnullorempty($EOPContentFilter.EndUserSpamNotificationLimit)) -EQ $FALSE) {
            $EOPContentFilterEndUserSpamNotificationArray += "Notification limit: " + ($EOPContentFilter.EndUserSpamNotificationLimit) +"`n"
        }
        $EOPContentFilterEndUserSpamNotificationArray =$([string]$EOPContentFilterEndUserSpamNotificationArray |out-string).trimend()
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "End user spam notification"
            "Value"                      = $EOPContentFilterEndUserSpamNotificationArray
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.DownloadLink)) -EQ $FALSE) {
            $EOPContentFilterDownloadLink = $EOPContentFilter.DownloadLink
        }
        else {

            $EOPContentFilterDownloadLink = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Download link"
            "Value"                      = $EOPContentFilterDownloadLink
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.EnableRegionBlockList)) -EQ $FALSE) {
            $EOPContentFilterEnableRegionBlockList = $EOPContentFilter.EnableRegionBlockList
        }
        else {

            $EOPContentFilterEnableRegionBlockList = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Enable region block list"
            "Value"                      = $EOPContentFilterEnableRegionBlockList
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If ($EOPContentFilter.RegionBlockList.length -gt 1) {
            $EOPContentFilterRegionBlockList = $EOPContentFilter.RegionBlockList |Out-String
        }
        else {

            $EOPContentFilterRegionBlockList = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Region block list"
            "Value"                      = $EOPContentFilterRegionBlockList
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  
        
        If ($EOPContentFilter.EnableLanguageBlockList -gt 1) {
            $EOPContentFilterEnableLanguageBlockList = $EOPContentFilter.EnableLanguageBlockList
        }
        else {

            $EOPContentFilterEnableLanguageBlockList = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Enable language block list"
            "Value"                      = $EOPContentFilterEnableLanguageBlockList
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If ($EOPContentFilter.LanguageBlockList -gt 1) {
            $EOPContentFilterLanguageBlockList = $EOPContentFilter.LanguageBlockList |Out-String
        }
        else {

            $EOPContentFilterLanguageBlockList = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Language block list"
            "Value"                      = $EOPContentFilterLanguageBlockList
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If (([string]::isnullorempty( $EOPContentFilter.BulkThreshold)) -EQ $FALSE) {
            $EOPContentFilterBulkThreshold = $EOPContentFilter.BulkThreshold |Out-String
            $EOPContentFilterBulkThreshold =$EOPContentFilterBulkThreshold.trim()
        }
        else {

            $EOPContentFilterBulkThreshold = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Bulk Threshold"
            "Value"                      = $EOPContentFilterBulkThreshold
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If ($EOPContentFilter.AllowedSenders  -gt 1) {
            $EOPContentFilterAllowedSenders = $EOPContentFilter.AllowedSenders |Out-String
        }
        else {

            $EOPContentFilterAllowedSenders = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Allowed Senders"
            "Value"                      = $EOPContentFilterAllowedSenders
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If ($EOPContentFilter.AllowedSenderDomains -gt 1) {
            $EOPContentFilterAllowedSenderDomains = $EOPContentFilter.AllowedSenderDomains |Out-String
        }
        else {

            $EOPContentFilterAllowedSenderDomains = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Allowed sender domains"
            "Value"                      = $EOPContentFilterAllowedSenderDomains
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If ($EOPContentFilter.BlockedSenders -gt 1) {
            $EOPContentFilterBlockedSenders = $EOPContentFilter.BlockedSenders |Out-String
        }
        else {

            $EOPContentFilterBlockedSenders = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Blocked Senders"
            "Value"                      = $EOPContentFilterBlockedSenders
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  

        If ($EOPContentFilter.BlockedSenderDomains -gt 1) {
            $EOPContentFilterBlockedSenderDomains = $EOPContentFilter.BlockedSenderDomains |Out-String
        }
        else {

            $EOPContentFilterBlockedSenderDomains = "Not Configured"
        }
        $EOPContentFilterDetail = [ordered]@{
            "Configuration Item"         = "Blocked Sender Domains"
            "Value"                      = $EOPContentFilterBlockedSenderDomains
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
        $EOPContentFilterArray += $EOConfigurationObject  
    }
    Else {
        $EOPContentFilter |ForEach-Object {
            If (([string]::isnullorempty($_.Name)) -EQ $FALSE) {
                $EOPContentFilterName = $_.Name
            }
            else {
    
                $EOPContentFilterName = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Name"
                "Value"                      = $EOPContentFilterName
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject
    
            If ($EOPContentFilter.AddxHeaderValue.Length  -gt 1 ) {
                $EOPContentFilterAddXHeaderValue =  $_.AddxHeaderValue
            }
            else {
    
                $EOPContentFilterAddXHeaderValue = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Add X Header Value"
                "Value"                      = $EOPContentFilterAddXHeaderValue
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject
            
            If (([string]::isnullorempty(  $_.ModifySubjectValue)) -EQ $FALSE) {
                $EOPContentFilterModifySubjectValue =  $_.ModifySubjectValue
            }
            else {
    
                $EOPContentFilterModifySubjectValue = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Modify Subject value"
                "Value"                      = $EOPContentFilterModifySubjectValue
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject
            
            If ($_.RedirectToRecipients -gt 1) {
                $EOPContentFilterRedirectToRecipients =  $_.RedirectToRecipients |out-string
                $EOPContentFilterRedirectToRecipients = $EOPContentFilterRedirectToRecipients.trim()
            }
            else {
    
                $EOPContentFilterRedirectToRecipients = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Redirect to recipients"
                "Value"                      = $EOPContentFilterRedirectToRecipients
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If ($_.FalsePositiveAdditionalRecipients -gt 1) {
                $EOPContentFilterFalsePositiveAdditionalRecipients =  $_.FalsePositiveAdditionalRecipients |out-string
                $EOPContentFilterFalsePositiveAdditionalRecipients =$EOPContentFilterFalsePositiveAdditionalRecipients.Trim()
            }
            else {
    
                $EOPContentFilterFalsePositiveAdditionalRecipients = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "False positive additional recipients"
                "Value"                      = $EOPContentFilterFalsePositiveAdditionalRecipients
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.QuarantineRetentionPeriod)) -EQ $FALSE) {
                $EOPContentFilterQuarantineRetentionPeriod =  $_.QuarantineRetentionPeriod
            }
            else {
    
                $EOPContentFilterQuarantineRetentionPeriod = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Quarantine retention period"
                "Value"                      = $EOPContentFilterQuarantineRetentionPeriod
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.EndUserSpamNotificationFrequency)) -EQ $FALSE) {
                $EOPContentFilterEndUserSpamNotificationFrequency =  $_.EndUserSpamNotificationFrequency
            }
            else {
    
                $EOPContentFilterEndUserSpamNotificationFrequency = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "End user spam notification frequency"
                "Value"                      = $EOPContentFilterEndUserSpamNotificationFrequency
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If(([string]::isnullorempty( $_.IncreaseScoreWithImageLinks)) -EQ $FALSE) {
                $EOPContentFilterIncreaseScoreArray += "Increase score with image links: " + ($_.IncreaseScoreWithImageLinks) +"`n"
            }
            If(([string]::isnullorempty( $_.IncreaseScoreWithNumericIps)) -EQ $FALSE) {
                $EOPContentFilterIncreaseScoreArray += "Increase score with numeric IPs: " + ($_.IncreaseScoreWithNumericIps)+"`n" 
            }
            If(([string]::isnullorempty( $_.IncreaseScoreWithRedirectToOtherPort)) -EQ $FALSE) {
                $EOPContentFilterIncreaseScoreArray += "Increase score with redirect to other port: " + ($_.IncreaseScoreWithRedirectToOtherPort) +"`n"
            }
            If(([string]::isnullorempty( $_.IncreaseScoreWithBizOrInfoUrls)) -EQ $FALSE) {
                $EOPContentFilterIncreaseScoreArray += "Increase score with Biz or info URLs: " + ($_.IncreaseScoreWithBizOrInfoUrls) +"`n"
            }
            $EOPContentFilterIncreaseScoreArray = $($EOPContentFilterIncreaseScoreArray|out-string).trimend()
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Increase Score"
                "Value"                      = $EOPContentFilterIncreaseScoreArray
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject
    
            If(([string]::isnullorempty( $_.MarkAsSpamEmptyMessages)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam empty messages: " + ($_.MarkAsSpamEmptyMessages) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamJavaScriptInHtml)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam javascript in html: " + ($_.MarkAsSpamJavaScriptInHtml) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamFramesInHtml)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam frames in HTML: " + ($_.MarkAsSpamFramesInHtml) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamObjectTagsInHtml)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam object tags in HTML: " + ($_.MarkAsSpamObjectTagsInHtml) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamEmbedTagsInHtml)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam embed tags in HTML: " + ($_.MarkAsSpamEmbedTagsInHtml) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamFormTagsInHtml)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam form tags in HTML: " + ($_.MarkAsSpamFormTagsInHtml) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamWebBugsInHtml)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam web bugs in HTML: " + ($_.MarkAsSpamWebBugsInHtml) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamSensitiveWordList)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam sensitive word list: " + ($_.MarkAsSpamSensitiveWordList) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamSpfRecordHardFail)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam SPF record hard fail: " + ($_.MarkAsSpamSpfRecordHardFail) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamFromAddressAuthFail)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam from address auth fail: " + ($_.MarkAsSpamFromAddressAuthFail) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamBulkMail)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam bulk mail: " + ($_.MarkAsSpamBulkMail) +"`n"
            }
            If(([string]::isnullorempty( $_.MarkAsSpamNdrBackscatter)) -EQ $FALSE) {
                $EOPContentFilterMarkAsSpameArray += "Mark as spam NDR backscatter: " + ($_.MarkAsSpamNdrBackscatter) +"`n"
            }
            $EOPContentFilterMarkAsSpameArray = $([string]$EOPContentFilterMarkAsSpameArray|out-string).trimend()
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Increase Score"
                "Value"                      = $EOPContentFilterMarkAsSpameArray
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.HighConfidenceSpamAction)) -EQ $FALSE) {
                $EOPContentFilterHighConfidenceSpamAction =  $_.HighConfidenceSpamAction
            }
            else {
    
                $EOPContentFilterHighConfidenceSpamAction = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "High confidence spam action"
                "Value"                      = $EOPContentFilterHighConfidenceSpamAction
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.SpamAction)) -EQ $FALSE) {
                $EOPContentFilterSpamAction =  $_.SpamAction
            }
            else {
    
                $EOPContentFilterSpamAction = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Spam action"
                "Value"                      = $EOPContentFilterSpamAction
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.BulkSpamAction)) -EQ $FALSE) {
                $EOPContentBulkSpamAction =  $_.BulkSpamAction
            }
            else {
    
                $EOPContentBulkSpamAction = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Bulk spam action"
                "Value"                      = $EOPContentBulkSpamAction
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.PhishSpamAction)) -EQ $FALSE) {
                $EOPContentFilterPhishSpamAction =  $_.PhishSpamAction
            }
            else {
    
                $EOPContentFilterPhishSpamAction = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Phish Spam action"
                "Value"                      = $EOPContentFilterPhishSpamAction
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.EnableEndUserSpamNotifications)) -EQ $FALSE) {
                $EOPContentFilterEnableEndUserSpamNotifications =  $_.EnableEndUserSpamNotifications
            }
            else {
    
                $EOPContentFilterEnableEndUserSpamNotifications = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Enable end user spam notifications"
                "Value"                      = $EOPContentFilterEnableEndUserSpamNotifications
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If(([string]::isnullorempty( $_.EndUserSpamNotificationCustomFromAddress)) -EQ $FALSE) {
                $EOPContentFilterEndUserSpamNotificationArray += "Notification custom from address: " + ($_.EndUserSpamNotificationCustomFromAddress) +"`n"
            }
            If(([string]::isnullorempty( $_.EndUserSpamNotificationCustomFromName)) -EQ $FALSE) {
                $EOPContentFilterEndUserSpamNotificationArray += "Notification custom from name: " + ($_.EndUserSpamNotificationCustomFromName) +"`n"
            }
            If(([string]::isnullorempty( $_.EndUserSpamNotificationCustomSubject)) -EQ $FALSE) {
                $EOPContentFilterEndUserSpamNotificationArray += "Notification custom from subject: " + ($_.EndUserSpamNotificationCustomSubject) +"`n"
            }
            If(([string]::isnullorempty( $_.EndUserSpamNotificationLanguage)) -EQ $FALSE) {
                $EOPContentFilterEndUserSpamNotificationArray += "Notification language: " + ($_.EndUserSpamNotificationLanguage) +"`n"
            }
            If(([string]::isnullorempty( $_.EndUserSpamNotificationLimit)) -EQ $FALSE) {
                $EOPContentFilterEndUserSpamNotificationArray += "Notification limit: " + ($_.EndUserSpamNotificationLimit)+"`n" 
            }
            $EOPContentFilterEndUserSpamNotificationArray = $([string]$EOPContentFilterEndUserSpamNotificationArray|out-string).trimend()
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "End user spam notification"
                "Value"                      = $EOPContentFilterEndUserSpamNotificationArray
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.DownloadLink)) -EQ $FALSE) {
                $EOPContentFilterDownloadLink =  $_.DownloadLink
            }
            else {
    
                $EOPContentFilterDownloadLink = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "DownLoad link"
                "Value"                      = $EOPContentFilterDownloadLink
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.EnableRegionBlockList)) -EQ $FALSE) {
                $EOPContentFilterEnableRegionBlockList =  $_.EnableRegionBlockList
            }
            else {
    
                $EOPContentFilterEnableRegionBlockList = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Enable region block list"
                "Value"                      = $EOPContentFilterEnableRegionBlockList
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.RegionBlockList)) -EQ $FALSE) {
                $EOPContentFilterRegionBlockList =  $_.RegionBlockList |Out-String
            }
            else {
    
                $EOPContentFilterRegionBlockList = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Region block list"
                "Value"                      = $EOPContentFilterRegionBlockList
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
            
            If (([string]::isnullorempty(  $_.EnableLanguageBlockList)) -EQ $FALSE) {
                $EOPContentFilterEnableLanguageBlockList =  $_.EnableLanguageBlockList
            }
            else {
    
                $EOPContentFilterEnableLanguageBlockList = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Enable language block list"
                "Value"                      = $EOPContentFilterEnableLanguageBlockList
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If ($_.LanguageBlockList -gt 1) {
                $EOPContentFilterLanguageBlockList =  $_.LanguageBlockList |Out-String
            }
            else {
    
                $EOPContentFilterLanguageBlockList = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Language block list"
                "Value"                      = $EOPContentFilterLanguageBlockList
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If (([string]::isnullorempty(  $_.BulkThreshold)) -EQ $FALSE) {
                $EOPContentFilterBulkThreshold =  $_.BulkThreshold |Out-String
            }
            else {
    
                $EOPContentFilterBulkThreshold = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Bulk Threshold"
                "Value"                      = $EOPContentFilterBulkThreshold
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If ($_.AllowedSenders -gt 1) {
                $EOPContentFilterAllowedSenders =  $_.AllowedSenders |Out-String
            }
            else {
    
                $EOPContentFilterAllowedSenders = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Allowed Senders"
                "Value"                      = $EOPContentFilterAllowedSenders
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If ($_.AllowedSenderDomains -gt 1) {
                $EOPContentFilterAllowedSenderDomains =  $_.AllowedSenderDomains |Out-String
            }
            else {
    
                $EOPContentFilterAllowedSenderDomains = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Allowed sender domains"
                "Value"                      = $EOPContentFilterAllowedSenderDomains
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If ($_.BlockedSenders -gt 1) {
                $EOPContentFilterBlockedSenders =  $_.BlockedSenders |Out-String
            }
            else {
    
                $EOPContentFilterBlockedSenders = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Blocked Senders"
                "Value"                      = $EOPContentFilterBlockedSenders
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  
    
            If ($_.BlockedSenderDomains -gt 1) {
                $EOPContentFilterBlockedSenderDomains =  $_.BlockedSenderDomains |Out-String
            }
            else {
    
                $EOPContentFilterBlockedSenderDomains = "Not Configured"
            }
            $EOPContentFilterDetail = [ordered]@{
                "Configuration Item"         = "Blocked Sender Domains"
                "Value"                      = $EOPContentFilterBlockedSenderDomains
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOPContentFilterDetail
            $EOPContentFilterArray += $EOConfigurationObject  

        }       
    }
}

#############################################
#Document - Configuration - Exchange Online Protection
#############################################
Add-Content -path $ReportFileName -value "## Exchange Online Protection"
Add-Content -path $ReportFileName -value "The ABAC settings for the Department's Exchange Online Protection instance can be found below. This includes the Connection Filtering, Anti-Malware, Policy Filtering, and Content Filtering Configuration." 
Add-Content -path $ReportFileName -value "Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting."

#############################################
#Document - Exchange Online Protection - Connection Filtering
#############################################
Write-Host " - Connection Filtering" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Connection Filtering"

$GUCount = $($EOPConnectionFilterArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Connection Filters have been configured."
    $EOPConnectionFilterArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Connection Filters configured"
}

#############################################
#Document - Exchange Online Protection - Anti-Malware
#############################################
Write-Host " - Anti-Malware" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Anti-Malware"

$GUCount = $($EOPMalwareFilterArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Malware Filters have been configured."
    $EOPMalwareFilterArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Malware Filters configured"
}

#############################################
#Document - Exchange Online Protection - Policy Filtering
#############################################
Write-Host " - Policy Filtering" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Policy Filtering"

$GUCount = $($EOPPolicyFilterArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Policy Filters have been configured."
    $EOPPolicyFilterArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Policy Filters configured"
}

#############################################
#Document - Exchange Online Protection - Content Filtering
#############################################
Write-Host " - Content Filtering" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Content Filtering"

$GUCount = $($EOPContentFilterArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Content Filters have been configured."
    $EOPContentFilterArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Content Filters configured"
}
