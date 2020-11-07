Function ConvertTo-MDTest {

    [cmdletbinding()]
        [outputtype([string[]])]
        [alias('ctm')]
    
        Param(
            [Parameter(Position = 0, ValueFromPipeline)]
            [object]$Inputobject,
            [Parameter()]
            [string]$Title,
            [string]$TableTitle,
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
            If ($TableTitle) {
                Write-Verbose "[BEGIN  ] Adding Table Title: $TableTitle"
                $Text += $TableTitle
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
    

#Office 365 As Built As Configured Reporting. Nathaniel O'Reilly 2019.
#Version 1.0 - Nathaniel O'Reilly - 2019-07-16

#This script leverages a word script developed by Greg Roll of oobe PTY LTD. 

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
#Arrays
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
$SharepointArray                             =@()
$TeamsClientConfigArray                      =@()
$TeamsChannelPolicyArray                     =@()
$TeamsCallingPolicyArray                     =@()
$TeamsMeetingConfigArray                     =@()
$TeamsMeetingPolicyArray                     =@()
$TeamsMessagingPolicyArray                   =@()
$TeamsMeetingBroadcastPolicyArray            =@()
$RetentionPolicyArray                        =@()
$RetentionlabelArray                         =@()
$SensitivitylabelpolicyArray                 =@()
$SensitivityLabelsArray                      =@()
$DlpCompliancePolicyArray                    =@()
###########################
#Variables
###########################
$domains                        = @("precisionservices.biz")
$Logging                        = "$pwd\Logs\Office365ABAC_$($(Get-Date).ToString(`"yyyy-MM-dd hhmmss`")).log"
$ReportFileName                 = "$pwd\Reports\Office365ABAC_$($(Get-Date).ToString(`"yyyy-MM-dd hhmmss`")).md"
$Tenant                         = "precisionservicesptyltd"
$admindomain                    = "precisionservicesptyltd.onmicrosoft.com"
$global:UserPrincipalName       = "martin@precisionservices.biz"

#Enable Full Logging
Start-Transcript -Path $Logging
Write-Host ""

###########################
#Connect to Exchange Online
###########################
Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
Connect-EXOPSSession -UserPrincipalName $UserPrincipalName
Write-Host " - Done" -foregroundcolor Green
Write-Host ""

#############################################
#Exchange Online
#############################################
Write-Host "Querying Exchange Online configuration..." -foregroundcolor Yellow

#############################################
#Exchange Online - Inbound Mail Connector
#############################################
Write-Host " - Inbound Mail Connector Details" -foregroundcolor Gray

$EOInboundConnector = Get-InboundConnector
If ($null -eq $EOInboundConnector) {
    $EOInboundConnectorDetail = [ordered]@{
        "Item"              = "Exchange Online Inbound Connector"
        "Name"              = "Not Configured"
        "Status"            = "N/A"
        "TLS"               = "N/A"
        "Certificate"       = "N/A"
        "Comments"          = "The Inbound connector should be configured if a hybrid configuration is in user or an external mail gateway is in use."
    }
    $EOConfigurationObject = New-Object -TypeName psobject -Property $EOInboundConnectorDetail
    $InboundMailConnectorsArray += $EOConfigurationObject
}
else {
    If ($EOInboundConnector -isnot [array]) {
        $EOInboundConnectorName    = $EOInboundConnector.name      
        $EOInboundConnectorStatus  = $EOInboundConnector.enabled
        $EOInboundConnectorTLS     = $EOInboundConnector.requireTLs
        $EOInboundConnectorTLSCert = $EOInboundConnector.TlsSenderCertificateName
            
        $EOInboundConnectorDetail = [ordered]@{
            "Name"              = $EOInboundConnectorName 
            "Status"            = $EOInboundConnectorStatus
            "TLS"               = $EOInboundConnectorTLS  
            "Certificate"       = $EOInboundConnectorTLSCert
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOInboundConnectorDetail
        $InboundMailConnectorsArray += $EOConfigurationObject
    }
    else {
        $EOInboundConnector | Foreach-Object {
        if ($_ -eq $EOInboundConnector[-1]) {
            if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                $EOInboundConnectorName    = $_.name         
                If ($_.enabled -eq $True) {
                    $EOInboundConnectorStatus  = "True" 
                }
                else {
                    $EOInboundConnectorStatus  = "False"   
                }
                If ($_.requireTLs -eq $True) {
                    $EOInboundConnectorTLS  = "True"
                }
                else {
                    $EOInboundConnectorTLS  = "False"   
                }
                If ($_.TlsSenderCertificateName -eq $null) {
                    $EOInboundConnectorTLSCert  = "Not Configured"
                }
                else {
                    $EOInboundConnectorTLSCert  = $_.TlsSenderCertificateName
                }
            }
        }
        else {
            if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                $EOInboundConnectorName    = $_.name 
                If ($_.enabled -eq $True) {
                    $EOInboundConnectorStatus  = "True"  
                }
                else {
                    $EOInboundConnectorStatus  = "False"    
                }
                If ($_.requireTLs -eq $True) {
                    $EOInboundConnectorTLS  = "True" 
                }
                else {
                    $EOInboundConnectorTLS  = "False"   
                }
                If ($_.TlsSenderCertificateName -eq $null) {
                    $EOInboundConnectorTLSCert  = "Not Configured" 
                }
                else {
                    $EOInboundConnectorTLSCert  = $_.TlsSenderCertificateName 
                }
            }
        }
        $EOInboundConnectorDetail = [ordered]@{
            "Name"              = $EOInboundConnectorName 
            "Status"            = $EOInboundConnectorStatus
            "TLS"               = $EOInboundConnectorTLS  
            "Certificate"       = $EOInboundConnectorTLSCert
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOInboundConnectorDetail
        $InboundMailConnectorsArray += $EOConfigurationObject
        }
    }
 }

#############################################
#Exchange Online - Outbound Mail Connector
#############################################
Write-Host " - Outbound Mail Connector Details" -foregroundcolor Gray

$EOOutboundConnector = Get-OutboundConnector

If ($null -eq $EOOutboundConnector) {
    $EOOutboundConnectorDetail = [ordered]@{
        "Item"              = "Exchange Online Outbound Connector"
        "Name"              = "Not Configured"
        "Status"            = "N/A"
        "TLS"               = "N/A"
        "Certificate"       = "N/A"
        "Comments"          = "The Outbound connector should be configured if a hybrid configuration is in user or an external mail gateway is in use."
    }
    $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOutboundConnectorDetail
    $OutboundMailConnectorsArray += $EOConfigurationObject
}
else {
    If ($EOOutboundConnector -isnot [array]) {
        $EOOutboundConnectorName    = $EOOutboundConnector.name      
        $EOOutboundConnectorStatus  = $EOOutboundConnector.enabled
        $EOOutboundConnectorTLS     = $EOOutboundConnector.requireTLs
        $EOOutboundConnectorTLSCert = $EOOutboundConnector.TlsSenderCertificateName
        
        $EOOutboundConnectorDetail = [ordered]@{
            "Name"              = $EOOutboundConnectorName 
            "Status"            = $EOOutboundConnectorStatus
            "TLS"               = $EOOutboundConnectorTLS  
            "Certificate"       = $EOOutboundConnectorTLSCert
        }
       $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOutboundConnectorDetail
       $OutboundMailConnectorsArray += $EOConfigurationObject
    }
    else {
        $EOOutboundConnector | Foreach-Object {
            if ($_ -eq $EOOutboundConnector[-1]) {
                if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                    $EOOutboundConnectorName    = $_.name         
                    If ($_.enabled -eq $True) {
                        $EOOutboundConnectorStatus  = "True" 
                    }
                    else {
                        $EOOutboundConnectorStatus  = "False"   
                    }
                    If ($_.requireTLs -eq $True) {
                        $EOOutboundConnectorTLS  = "True"
                    }
                    else {
                        $EOOutboundConnectorTLS  = "False"   
                    }
                    If ($_.TlsSenderCertificateName -eq $null) {
                        $EOOutboundConnectorTLSCert  = "Not Configured"
                    }
                    else {
                        $EOOutboundConnectorTLSCert  = $_.TlsSenderCertificateName
                    }
                }
            }
            else {
                if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                    $EOOutboundConnectorName    = $_.name     
                    If ($_.enabled -eq $True) {
                        $EOOutboundConnectorStatus  = "True"
                    }
                    else {
                        $EOOutboundConnectorStatus  = "False"  
                    }
                    If ($_.requireTLs -eq $True) {
                        $EOOutboundConnectorTLS  = "True"
                    }
                    else {
                        $EOOutboundConnectorTLS  = "False" 
                    }
                    If ($_.TlsSenderCertificateName -eq $null) {
                        $EOOutboundConnectorTLSCert  = "Not Configured"
                    }
                    else {
                        $EOOutboundConnectorTLSCert  = $_.TlsSenderCertificateName
                    }
                }
            }
        $EOOutboundConnectorDetail = [ordered]@{
            "Name"              = $EOOutboundConnectorName 
            "Status"            = $EOOutboundConnectorStatus
            "TLS"               = $EOOutboundConnectorTLS  
            "Certificate"       = $EOOutboundConnectorTLSCert
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOutboundConnectorDetail
        $OutboundMailConnectorsArray += $EOConfigurationObject
        }
    }
}

#############################################
#Exchange Online - MX Records
#############################################
Write-Host " - Mail Exchange Records" -foregroundcolor Gray

$EOMXrecordPreferencelist =@()
If ($domains -ne $null) {
    $domains |Foreach-Object {
        $EOMXrecords= nslookup.exe -type=MX $_ 2>$null |select-string "MX"
        If ($null -eq $EOMXrecords) {
            $EOMXrecordsDetail = [ordered]@{
                "Domain"              = $_ 
                "MX Preference"       = "Not Configured"
                "Mail Exchanger"      = "Not Configured"
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOMXrecordsDetail
            $MXrecordsArray += $EOConfigurationObject
        }
        Else{
            If ($EOMXrecords -isnot [array]) {
                $EOMXrecords = $EOMXrecords|Out-String
                $EOMXrecords = [string]$EOMXrecords.Substring([string]$EOMXrecords.IndexOf("MX"))
                $EOMXrecordPreference = [string]$EOMXrecords.Substring($EOMXrecords.IndexOf("MX"),$EOMXrecords.IndexOf(","))
                $EOMXrecordMX = [string]$EOMXrecords.Substring($EOMXrecords.IndexOf("mail"))            
                $EOMXrecordsDetail = [ordered]@{
                    "Domain"              = $_ 
                    "MX Preference"       = $EOMXrecordPreference
                    "Mail Exchanger"      = $EOMXrecordMX
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOMXrecordsDetail
                $MXrecordsArray += $EOConfigurationObject  
            }
            Else {
                Foreach($EOMXrecord in  $EOMXrecords) {
                    if ($EOMXrecord -eq $EOMXrecords[0]) {
                        $EOMXrecord = $EOMXrecord|Out-String
                        $EOMXrecord = [string]$EOMXrecord.Substring([string]$EOMXrecord.IndexOf("MX"))
                        $EOMXrecordPreference = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("MX"),$EOMXrecord.IndexOf(","))
                        $EOMXrecordMX = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("mail"))  
                        $EOMXrecordMX = $EOMXrecordMX.TrimEnd()      
                        [string]$EOMXrecordPreferencelist +=  [string]$EOMXrecordPreference
                        [string]$EOMXrecordMXlist +=  [string]$EOMXrecordMX
                    }
                    else {
                        $EOMXrecord = $EOMXrecord|Out-String
                        $EOMXrecord = [string]$EOMXrecord.Substring([string]$EOMXrecord.IndexOf("MX"))
                        $EOMXrecordPreference = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("MX"),$EOMXrecord.IndexOf(","))
                        $EOMXrecordMX = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("mail"))    
                        $EOMXrecordMX = $EOMXrecordMX.TrimEnd()     
                        [string]$EOMXrecordPreferencelist += "`n" + [string]$EOMXrecordPreference
                        [string]$EOMXrecordMXlist += "`n" + [string]$EOMXrecordMX                    
                    }
                }
                $EOMXrecordsDetail = [ordered]@{
                    "Domain"              = $_ 
                    "MX Preference"       = $EOMXrecordPreferencelist
                    "Mail Exchanger"      = $EOMXrecordMXlist
                }
                $EOMXrecordPreferencelist =$null
                $EOMXrecordMXlist =$null
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOMXrecordsDetail
                $MXrecordsArray += $EOConfigurationObject
            }
        }
    }

}

#############################################
#Exchange Online - SPF Records
#############################################
Write-Host " - SPF Records and DMARC" -foregroundcolor Gray

If ($domains -ne $null) {
    $domains |Foreach-Object {
        $EOSPFrecords= nslookup.exe -type=txt $_ 2>$null|select-string "v=SPF"
        $DMARCDomain = "_dmarc." + [string]$_
        $EODMARCrecords= nslookup.exe -type=txt $DMARCDomain 2>$null|select-string "v="
        If (($null -eq $EOSPFrecords) -AND ($null -eq $EODMARCrecords)) {
            $EOSPFrecordsDetail = [ordered]@{
                "Domain"              = $_ 
                "SPF Record"          = "Not Configured"
                "DMARC Policy"        = "Not Configured"
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOSPFrecordsDetail
            $SPFrecordsArray += $EOConfigurationObject  
        }
        else {
            If ($EOSPFrecords -isnot [object]) {
                if(([string]::isnullorempty($EODMARCrecords)) -EQ $FALSE) {
                    $EODMARCrecords = ($EODMARCrecords|Out-String)
                    $EODMARCrecords = [string]$EODMARCrecords.Substring([string]$EODMARCrecords.IndexOf('"'))
                    $EODMARCrecords = $EODMARCrecords.Trim()
                    $EODMARCrecords = $EODMARCrecords.TrimEnd()
                }
                else {
                    $EODMARCrecords = "Not configured"                
                }
                
                if(([string]::isnullorempty($EOSPFrecords)) -EQ $FALSE) {
                    $EOSPFrecords = ($EOSPFrecords|Out-String)
                    $EOSPFrecords = [string]$EOSPFrecords.Substring($EOSPFrecords.IndexOf('"'),$EOSPFrecords.LastIndexOf('"'))
                    $EOSPFrecords = $EOSPFrecords.TrimEnd()
                    $EOSPFrecords = $EOSPFrecords.TrimStart()
                }
                else {
                    $EOSPFrecords = "Not configured"
                }
                
                $EOSPFrecordsDetail = [ordered]@{
                    "Domain"              = $_ 
                    "SPF Record"          = $EOSPFrecords
                    "DMARC Policy"        = $EODMARCrecords
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOSPFrecordsDetail
                $SPFrecordsArray += $EOConfigurationObject  
            }
            Else {
                if(([string]::isnullorempty($EODMARCrecords)) -EQ $FALSE) {
                    $EODMARCrecords = ($EODMARCrecords|Out-String)
                    $EODMARCrecords = [string]$EODMARCrecords.Substring([string]$EODMARCrecords.IndexOf('"'))
                    $EODMARCrecords = $EODMARCrecords.Trim()
                    $EODMARCrecords = $EODMARCrecords.Trimend()
                }
                else {
                    $EODMARCrecords = "Not configured"
                }
                
                if(([string]::isnullorempty($EOSPFrecords)) -EQ $FALSE) {
                    $EOSPFrecords = ($EOSPFrecords|Out-String)
                    $EOSPFrecords = [string]$EOSPFrecords.Substring($EOSPFrecords.IndexOf('"'),$EOSPFrecords.LastIndexOf('"'))
                    $EOSPFrecords = $EOSPFrecords.TrimEnd()
                    $EOSPFrecords = $EOSPFrecords.TrimStart()
                }
                else {
                    $EOSPFrecords = "Not configured"
                }
                $EOSPFrecordsDetail = [ordered]@{
                    "Domain"              = $_ 
                    "SPF Record"           = $EOSPFrecords
                    "DMARC Policy"        = $EODMARCrecords
                }
                $EOConfigurationObject = New-Object -TypeName psobject -Property $EOSPFrecordsDetail
                $SPFrecordsArray += $EOConfigurationObject
            }
        }
    }
}

#############################################
#Exchange Online - Remote Domains
#############################################
Write-Host " - Remote Domains" -foregroundcolor Gray

$EORemoteDomains =  get-remotedomain 
If ($EORemoteDomains -ne $null) {
    $EORemoteDomains |Foreach-Object {
        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Name [TBD]"
            "Value"                 = $_.identity
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        
        $RemoteDomainsArray += $EOConfigurationObject
        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Remote Domain"
            "Value"                 = $_.domainname
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject
        
        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Allowed Out Of Office Type"
            "Value"                 = $_.allowedOOFType
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject
        
        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Automatic Reply"
            "Value"                 = $_.AutoReplyEnabled
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject
    
        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Automatic Forward"
            "Value"                 = $_.AutoForwardEnabled
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject
            
        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Delivery Reports"
            "Value"                 = $_.DeliveryReportEnabled
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject
    
        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Non Delivery Reports"
            "Value"                 = $_.NDREnabled
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject

        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Meeting Forward Notifications"
            "Value"                 = $_.MeetingForwardNotificationEnabled
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject

        $EORemoteDomainsDetail = [ordered]@{
            "Configuration Item"    = "Content Type"
            "Value"                 = $_.ContentType
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EORemoteDomainsDetail
        $RemoteDomainsArray += $EOConfigurationObject
    }
}

#############################################
#Exchange Online - CAS Mailbox Plan
#############################################
Write-Host " - CAS Mailbox Plans" -foregroundcolor Gray

$EOCASMailboxPlan =  Get-CASMailboxPlan
If ($EOCASMailboxPlan -ne $null) {
    $EOCASMailboxPlan |Foreach-Object {
        If (([string]::isnullorempty($_.displayname)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Name [TBD]"
                "Value"                 = $_.Displayname
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Name [TBD]"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.ActiveSyncEnabled)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "ActiveSync"
                "Value"                 = $_.ActiveSyncEnabled
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "ActiveSync"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.activesyncmailboxpolicy)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "ActiveSync Mailbox Policy"
                "Value"                 = $_.activesyncmailboxpolicy
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "ActiveSync Mailbox Policy"
                "Value"                 = "Not Configured"
            }
        }   
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.IMAPEnabled)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "IMAP"
                "Value"                 = $_.IMAPEnabled
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "IMAP"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.MAPIEnabled)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "MAPI"
                "Value"                 = $_.MAPIEnabled
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "MAPI"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.OWAEnabled)) -EQ $FALSE) { 
        $EOCASMailboxPlanDetail = [ordered]@{
            "Configuration Item"    = "Outlook Web Access"
            "Value"                 = $_.OWAEnabled
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Outlook Web Access"
                "Value"                 = "Not Configured"
        }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.OWAMailboxPolicy)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Outlook Web Access Policy"
                "Value"                 = $_.OWAMailboxPolicy
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Outlook Web Access Policy"
                "Value"                 = "Not Configured"
        }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.POPEnabled)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "POP"
                "Value"                 = $_.POPEnabled
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "POP"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
        
        If (([string]::isnullorempty($_.EWSEnabled)) -EQ $FALSE) {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "EWS"
                "Value"                 = $_.EWSEnabled
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "EWS"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOCASMailboxPlanDetail
        $CASMailboxPlanArray += $EOConfigurationObject
    }
}

#############################################
#Exchange Online - Authentication Policy
#############################################
Write-Host " - Authentication Policy" -foregroundcolor Gray

$EOAuthenticationPolicy =  Get-AuthenticationPolicy
If ($EOAuthenticationPolicy -ne $null) {
    $EOAuthenticationPolicy |Foreach-Object {
        If (([string]::isnullorempty($_.identity)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Name [TBD]"
                "Value"                 = $_.identity
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Name [TBD]"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthActiveSync)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication ActiveSync"
                "Value"                 = $_.AllowBasicAuthActiveSync
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication ActiveSync"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthAutodiscover)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Autodiscover"
                "Value"                 = $_.AllowBasicAuthAutodiscover
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Autodiscover"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthIMAP)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication IMAP"
                "Value"                 = $_.AllowBasicAuthIMAP
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication IMAP"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthMapi)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication MAPI"
                "Value"                 = $_.AllowBasicAuthMapi
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication MAPI"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthOfflineAddressBook)) -EQ $FALSE) {   
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Offline AddressBook"
                "Value"                 = $_.AllowBasicAuthOfflineAddressBook
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Offline AddressBook"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthOutlookService)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Outlook Service"
                "Value"                 = $_.AllowBasicAuthOutlookService
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Outlook Service"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthPop)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication POP"
                "Value"                 = $_.AllowBasicAuthPop
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication POP"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthReportingWebServices)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Reporting Web Services"
                "Value"                 = $_.AllowBasicAuthReportingWebServices
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Reporting Web Services"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthRest)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Rest"
                "Value"                 = $_.AllowBasicAuthRest
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Rest"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthRPC)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication RPC"
                "Value"                 = $_.AllowBasicAuthRPC
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication RPC"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthSMTP)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication SMTP"
                "Value"                 = $_.AllowBasicAuthSMTP
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication SMTP"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthWebServices)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Web Services"
                "Value"                 = $_.AllowBasicAuthWebServices
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication Web Services"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
        If (([string]::isnullorempty($_.AllowBasicAuthPowershell)) -EQ $FALSE) {
            $EOAuthenticationPolicyDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication PowerShell"
                "Value"                 = $_.AllowBasicAuthPowershell
            }
        }
        else {
            $EOCASMailboxPlanDetail = [ordered]@{
                "Configuration Item"    = "Allow Basic Authentication PowerShell"
                "Value"                 = "Not Configured"
            }
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAuthenticationPolicyDetail
        $EOAuthenticationPolicyArray += $EOConfigurationObject
    }
}
#############################################
#Exchange Online - Outlook Web Access Policy
#############################################
Write-Host " - Outlook Web Access Policy" -foregroundcolor Gray

$EOOWAMailboxPolicy =  Get-OwaMailboxPolicy
If ($EOOWAMailboxPolicy -ne $null) {
    $EOOWAMailboxPolicy |Foreach-Object {
        $EOOWAMailboxPolicyDetail = [ordered]@{
        "Configuration Item"    = "Name [TBD]"
        "Value"                 = $_.identity
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Wac Editing Enabled"
            "Value"                 = $_.WacEditingEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Print Without Download Enabled"
            "Value"                 = $_.PrintWithoutDownloadEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "OneDrive Attachments Enabled"
            "Value"                 = $_.OneDriveAttachmentsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Third Party File Providers Enabled"
            "Value"                 = $_.ThirdPartyFileProvidersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Classic Attachments Enabled"
            "Value"                 = $_.ClassicAttachmentsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Reference Attachments Enabled"
            "Value"                 = $_.ReferenceAttachmentsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Save Attachments To Cloud Enabled"
            "Value"                 = $_.SaveAttachmentsToCloudEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Message Previews Disabled"
            "Value"                 = $_.MessagePreviewsDisabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Direct File Access On Public Computers Enabled"
            "Value"                 = $_.DirectFileAccessOnPublicComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Direct File Access On Private Computers Enabled"
            "Value"                 = $_.DirectFileAccessOnPrivateComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Web Ready Document Viewing On Public Computers Enabled"
            "Value"                 = $_.WebReadyDocumentViewingOnPublicComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Web Ready Document Viewing On Private Computers Enabled"
            "Value"                 = $_.WebReadyDocumentViewingOnPrivateComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Force Web Ready Document Viewing First On Public Computers"
            "Value"                 = $_.ForceWebReadyDocumentViewingFirstOnPublicComputers
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Force Web Ready Document Viewing First On Private Computers"
            "Value"                 = $_.ForceWebReadyDocumentViewingFirstOnPrivateComputers
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Wac Viewing On Public Computers Enabled"
            "Value"                 = $_.WacViewingOnPublicComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Wac Viewing On Private Computers Enabled"
            "Value"                 = $_.WacViewingOnPrivateComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Force Wac Viewing First On Public Computers"
            "Value"                 = $_.ForceWacViewingFirstOnPublicComputers
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Force Wac Viewing First On Private Computers"
            "Value"                 = $_.ForceWacViewingFirstOnPrivateComputers
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Action For Unknown File And MIME Types"
            "Value"                 = $_.WacEditingEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Phonetic Support Enabled"
            "Value"                 = $_.PhoneticSupportEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Default Client Language"
            "Value"                 = $_.DefaultClientLanguage
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Use GB18030"
            "Value"                 = $_.UseGB18030
        }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Use ISO885915"
            "Value"                 = $_.UseISO885915
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Outbound Charset"
            "Value"                 = $_.OutboundCharset
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Global Address List Enabled"
            "Value"                 = $_.GlobalAddressListEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Organization Enabled"
            "Value"                 = $_.OrganizationEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Explicit Logon Enabled"
            "Value"                 = $_.ExplicitLogonEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "OWA Light Enabled"
            "Value"                 = $_.OWALightEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Delegate Access Enabled"
            "Value"                 = $_.DelegateAccessEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "IRM Enabled"
            "Value"                 = $_.IRMEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Calendar Enabled"
            "Value"                 = $_.CalendarEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Contacts Enabled"
            "Value"                 = $_.ContactsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Tasks Enabled"
            "Value"                 = $_.TasksEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Journal Enabled"
            "Value"                 = $_.JournalEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Notes Enabled"
            "Value"                 = $_.NotesEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "On Send Addins Enabled"
            "Value"                 = $_.OnSendAddinsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Reminders And Notifications Enabled"
            "Value"                 = $_.RemindersAndNotificationsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Premium Client Enabled"
            "Value"                 = $_.WacEditingEnabled
            }   
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Spell Checker Enabled"
            "Value"                 = $_.ThirdPartyFileProvidersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Classic Attachments Enabled"
            "Value"                 = $_.ClassicAttachmentsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Search Folders Enabled"
            "Value"                 = $_.SearchFoldersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Signatures Enabled"
            "Value"                 = $_.SignaturesEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Theme Selection Enabled"
            "Value"                 = $_.ThemeSelectionEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Junk Email Enabled"
            "Value"                 = $_.JunkEmailEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "UM Integration Enabled"
            "Value"                 = $_.UMIntegrationEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "WSS Access On Public Computers Enabled"
            "Value"                 = $_.WSSAccessOnPublicComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "WSS Access On Private Computers Enabled"
            "Value"                 = $_.WSSAccessOnPrivateComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Change Password Enabled"
            "Value"                 = $_.ChangePasswordEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "UNC Access On Public Computers Enabled"
            "Value"                 = $_.UNCAccessOnPublicComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "UNC Access On Private Computers Enabled"
            "Value"                 = $_.UNCAccessOnPrivateComputersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "ActiveSync Integration Enabled"
            "Value"                 = $_.ActiveSyncIntegrationEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "All Address Lists Enabled"
            "Value"                 = $_.AllAddressListsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Rules Enabled"
            "Value"                 = $_.RulesEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Public Folders Enabled"
            "Value"                 = $_.PublicFoldersEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "SMime Enabled"
            "Value"                 = $_.SMimeEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Recover Deleted Items Enabled"
            "Value"                 = $_.RecoverDeletedItemsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Instant Messaging Enabled"
            "Value"                 = $_.InstantMessagingEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Text Messaging Enabled"
            "Value"                 = $_.TextMessagingEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Force Save Attachment Filtering Enabled"
            "Value"                 = $_.ForceSaveAttachmentFilteringEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Silverlight Enabled"
            "Value"                 = $_.SilverlightEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Instant Messaging Type"
            "Value"                 = $_.InstantMessagingType
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Display Photos Enabled"
            "Value"                 = $_.DisplayPhotosEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Set Photo Enabled"
            "Value"                 = $_.SetPhotoEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Allow Offline On"
            "Value"                 = $_.AllowOfflineOn
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Set Photo URL"
            "Value"                 = $_.SetPhotoURL
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Places Enabled"
            "Value"                 = $_.PlacesEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Weather Enabled"
            "Value"                 = $_.WeatherEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Local Events Enabled"
            "Value"                 = $_.LocalEventsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Interesting Calendars Enabled"
            "Value"                 = $_.InterestingCalendarsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Allow Copy Contacts To Device Address Book"
            "Value"                 = $_.AllowCopyContactsToDeviceAddressBook
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Predicted Actions Enabled"
            "Value"                 = $_.PredictedActionsEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject

        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "User Diagnostic Enabled"
            "Value"                 = $_.UserDiagnosticEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Facebook Enabled"
            "Value"                 = $_.FacebookEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "LinkedIn Enabled"
            "Value"                 = $_.LinkedInEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Wac External Services Enabled"
            "Value"                 = $_.WacExternalServicesEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Wac OMEX Enabled"
            "Value"                 = $_.WacOMEXEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Report Junk Email Enabled"
            "Value"                 = $_.ReportJunkEmailEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Group Creation Enabled"
            "Value"                 = $_.GroupCreationEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
        
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Skip Create Unified Group Custom Sharepoint Classification"
            "Value"                 = $_.SkipCreateUnifiedGroupCustomSharepointClassification
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "User Voice Enabled"
            "Value"                 = $_.UserVoiceEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
                    
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Satisfaction Enabled"
            "Value"                 = $_.SatisfactionEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
            
        $EOOWAMailboxPolicyDetail = [ordered]@{
            "Configuration Item"    = "Outlook Beta Toggle Enabled"
            "Value"                 = $_.OutlookBetaToggleEnabled
            }
        $EOConfigurationObject = New-Object -TypeName psobject -Property $EOOWAMailboxPolicyDetail
        $OWAMailboxPolicyArray += $EOConfigurationObject
    }
}

#############################################
#Exchange Online - Address Lists
#############################################
Write-Host " - Address List Details" -foregroundcolor Gray

$EOAddressLists = Get-AddressList
If ($null -eq $EOAddressLists) {
    $EOAddressListsDetail = [ordered]@{
        "Name"              = "Not Configured"
        "Recipient Filter"  = "N/A"
    }
    $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAddressListsDetail
    $EOAddressListsArray += $EOConfigurationObject
    }
    else {
        If ($EOAddressLists -isnot [array]) {
            $EOAddressListsName    = $EOAddressLists.name      
            $EOAddressListsRecipientFilter  = $EOAddressLists.recipientfilter
            
            $EOAddressListsDetail = [ordered]@{
                "Name"              = $EOAddressListsName 
                "Recipient Filter"  = $EOAddressListsRecipientFilter
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAddressListsDetail
            $EOAddressListsArray += $EOConfigurationObject
        }
        else {
            $EOAddressLists | Foreach-Object {
                if ($_ -eq $EOAddressLists[-1]) {
                    if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                        $EOAddressListsName    = $_.name         
                        $EOAddressListsRecipientFilter  = $_.recipientfilter
                    }
                }
                else {
                    if (($_.Name -ne $null) -and ($_.Name -ne "")) {
                        $EOAddressListsName    = $_.name 
                        $EOAddressListsRecipientFilter  = $_.recipientfilter
                    }
                }
            $EOAddressListsDetail = [ordered]@{
                "Name"              = $EOAddressListsName 
                "Recipient Filter"  = $EOAddressListsRecipientFilter
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOAddressListsDetail
            $EOAddressListsArray += $EOConfigurationObject
            }
        }
    }
######################################################################################################################################################################################################################################################################################################
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
######################################################################################################################################################################################################################################################################################################
#############################################
#SharePoint
#############################################
Write-Host "Connecting to SharePoint Online" -foregroundcolor Yellow
Connect-SPOService -url "https://$tenant-admin.sharepoint.com"
Write-Host "Querying SharePoint configuration..." -foregroundcolor Yellow

#############################################
#SharePoint - Tenant Configuration
#############################################
Write-Host " - Tenant Configuration" -foregroundcolor Gray
$SharepointTenant = Get-SPOTenant

If ($null -eq  $SharepointTenant) {
     $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject
}
Else {
    If (([string]::isnullorempty( $SharepointTenant.StorageQuota)) -EQ $FALSE) {
        $SharepointTenantStoragequota = $($SharepointTenant.StorageQuota|out-string).trim()
    }
    else{
        $SharepointTenantStoragequota = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Storage quota"
        "Value"                    = $SharepointTenantStoragequota
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.StorageQuotaAllocated)) -EQ $FALSE) {
        $SharepointTenantStorageQuotaAllocated = $($SharepointTenant.StorageQuotaAllocated|out-string).trim()
    }
    else{
        $SharepointTenantStorageQuotaAllocated = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Storage quota allocated"
        "Value"                    = $SharepointTenantStorageQuotaAllocated
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.ResourceQuota)) -EQ $FALSE) {
        $SharepointTenantResourceQuota = $($SharepointTenant.ResourceQuota|out-string).trim()
    }
    else{
        $SharepointTenantResourceQuota = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Resource quota"
        "Value"                    = $SharepointTenantResourceQuota
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.ResourceQuotaAllocated)) -EQ $FALSE) {
        $SharepointTenantResourceQuotaAllocated = $($SharepointTenant.ResourceQuotaAllocated|out-string).trim()
    }
    else{
        $SharepointTenantResourceQuotaAllocated = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Resource quota allocated"
        "Value"                    = $SharepointTenantResourceQuotaAllocated
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.ExternalServicesEnabled)) -EQ $FALSE) {
        $SharepointTenantExternalServicesEnabled = $($SharepointTenant.ExternalServicesEnabled|out-string).trim()
    }
    else{
        $SharepointTenantExternalServicesEnabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "External services enabled"
        "Value"                    = $SharepointTenantExternalServicesEnabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.NoAccessRedirectUrl)) -EQ $FALSE) {
        $SharepointTenantNoAccessRedirectUrl = $($SharepointTenant.NoAccessRedirectUrl|out-string).trim()
    }
    else{
        $SharepointTenantNoAccessRedirectUrl = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "No access redirect URL"
        "Value"                    = $SharepointTenantNoAccessRedirectUrl
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.SharingCapability)) -EQ $FALSE) {
        $SharepointTenantSharingCapability = $($SharepointTenant.SharingCapability|out-string).trim()
    }
    else{
        $SharepointTenantSharingCapability = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Sharing Capability"
        "Value"                    = $SharepointTenantSharingCapability
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.DisplayStartASiteOption)) -EQ $FALSE) {
        $SharepointTenantDisplayStartASiteOption = $($SharepointTenant.DisplayStartASiteOption|out-string).trim()
    }
    else{
        $SharepointTenantDisplayStartASiteOption = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Display start a site option"
        "Value"                    = $SharepointTenantDisplayStartASiteOption
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.StartASiteFormUrl)) -EQ $FALSE) {
        $SharepointTenantStartASiteFormUrl = $($SharepointTenant.StartASiteFormUrl|out-string).trim()
    }
    else{
        $SharepointTenantStartASiteFormUrl = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Start a site form URL"
        "Value"                    = $SharepointTenantStartASiteFormUrl
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.OfficeClientADALDisabled)) -EQ $FALSE) {
        $SharepointTenantOfficeClientADALDisabled = $($SharepointTenant.OfficeClientADALDisabled|out-string).trim()
    }
    else{
        $SharepointTenantOfficeClientADALDisabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Office client ADAL Disabled"
        "Value"                    = $SharepointTenantOfficeClientADALDisabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.LegacyAuthProtocolsEnabled)) -EQ $FALSE) {
        $SharepointTenantLegacyAuthProtocolsEnabled = $($SharepointTenant.LegacyAuthProtocolsEnabled|out-string).trim()
    }
    else{
        $SharepointTenantLegacyAuthProtocolsEnabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Legacy authentication protocols enabled"
        "Value"                    = $SharepointTenantLegacyAuthProtocolsEnabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.SearchResolveExactEmailOrUPN)) -EQ $FALSE) {
        $SharepointTenantSearchResolveExactEmailOrUPN = $($SharepointTenant.SearchResolveExactEmailOrUPN|out-string).trim()
    }
    else{
        $SharepointTenantSearchResolveExactEmailOrUPN = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Search resolve exact email or UPN"
        "Value"                    = $SharepointTenantSearchResolveExactEmailOrUPN
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject
    
    If (([string]::isnullorempty( $SharepointTenant.RequireAcceptingAccountMatchInvitedAccount)) -EQ $FALSE) {
        $SharepointTenantRequireAcceptingAccountMatchInvitedAccount = $($SharepointTenant.RequireAcceptingAccountMatchInvitedAccount|out-string).trim()
    }
    else{
        $SharepointTenantRequireAcceptingAccountMatchInvitedAccount = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Require accepting account match invited account"
        "Value"                    = $SharepointTenantRequireAcceptingAccountMatchInvitedAccount
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.ProvisionSharedWithEveryoneFolder)) -EQ $FALSE) {
        $SharepointTenantProvisionSharedWithEveryoneFolder = $($SharepointTenant.ProvisionSharedWithEveryoneFolder|out-string).trim()
    }
    else{
        $SharepointTenantProvisionSharedWithEveryoneFolder = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Provision shared with everyone folder"
        "Value"                    = $SharepointTenantProvisionSharedWithEveryoneFolder
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.BccExternalSharingInvitations)) -EQ $FALSE) {
        $SharepointTenantBccExternalSharingInvitations = $($SharepointTenant.BccExternalSharingInvitations|out-string).trim()
    }
    else{
        $SharepointTenantBccExternalSharingInvitations = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Bcc external sharing invitations"
        "Value"                    = $SharepointTenantBccExternalSharingInvitations
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.BccExternalSharingInvitationsList)) -EQ $FALSE) {
        $SharepointTenantBccExternalSharingInvitationsList = $($SharepointTenant.BccExternalSharingInvitationsList|out-string).trim()
    }
    else{
        $SharepointTenantProvisionSharedWithEveryoneFolder = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Bcc external sharing invitation list"
        "Value"                    = $SharepointTenantBccExternalSharingInvitationsList
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.UserVoiceForFeedbackEnabled)) -EQ $FALSE) {
        $SharepointTenantUserVoiceForFeedbackEnabled = $($SharepointTenant.UserVoiceForFeedbackEnabled|out-string).trim()
    }
    else{
        $SharepointTenantUserVoiceForFeedbackEnabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "User Voice for feedback enabled"
        "Value"                    = $SharepointTenantUserVoiceForFeedbackEnabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.PublicCdnEnabled)) -EQ $FALSE) {
        $SharepointTenantPublicCdnEnabled = $($SharepointTenant.PublicCdnEnabled|out-string).trim()
    }
    else{
        $SharepointTenantPublicCdnEnabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Public CDN enabled"
        "Value"                    = $SharepointTenantPublicCdnEnabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.PublicCdnAllowedFileTypes)) -EQ $FALSE) {
        $SharepointTenantPublicCdnAllowedFileTypes = $($SharepointTenant.PublicCdnAllowedFileTypes|out-string).trim()
    }
    else{
        $SharepointTenantPublicCdnAllowedFileTypes = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Public CDN allowed file types"
        "Value"                    = $SharepointTenantPublicCdnAllowedFileTypes
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.PublicCdnOrigins)) -EQ $FALSE) {
        $SharepointTenantPublicCdnOrigins = $($SharepointTenant.PublicCdnOrigins|out-string).trim()
    }
    else{
        $SharepointTenantPublicCdnOrigins = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Public CDN origins"
        "Value"                    = $SharepointTenantPublicCdnOrigins
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.RequireAnonymousLinksExpireInDays)) -EQ $FALSE) {
        $SharepointTenantRequireAnonymousLinksExpireInDays = $($SharepointTenant.RequireAnonymousLinksExpireInDays|out-string).trim()
    }
    else{
        $SharepointTenantRequireAnonymousLinksExpireInDays = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Require anonymous links expiration in days"
        "Value"                    = $SharepointTenantRequireAnonymousLinksExpireInDays
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.SharingAllowedDomainList)) -EQ $FALSE) {
        $SharepointTenantSharingAllowedDomainList = $($SharepointTenant.SharingAllowedDomainList|out-string).trim()
    }
    else{
        $SharepointTenantSharingAllowedDomainList = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Sharing allowed domain list"
        "Value"                    = $SharepointTenantSharingAllowedDomainList
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.SharingBlockedDomainList)) -EQ $FALSE) {
        $SharepointTenantSharingBlockedDomainList = $($SharepointTenant.SharingBlockedDomainList|out-string).trim()
    }
    else{
        $SharepointTenantSharingBlockedDomainList = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Sharing blocked domain list"
        "Value"                    = $SharepointTenantSharingBlockedDomainList
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject


    If (([string]::isnullorempty( $SharepointTenant.SharingDomainRestrictionMode)) -EQ $FALSE) {
        $SharepointTenantSharingDomainRestrictionMode = $($SharepointTenant.SharingDomainRestrictionMode|out-string).trim()
    }
    else{
        $SharepointTenantSharingDomainRestrictionMode = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Sharing domain restriction mode"
        "Value"                    = $SharepointTenantSharingDomainRestrictionMode
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.IPAddressEnforcement)) -EQ $FALSE) {
        $SharepointTenantIPAddressEnforcement = $($SharepointTenant.IPAddressEnforcement|out-string).trim()
    }
    else{
        $SharepointTenantIPAddressEnforcement = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "IP address enforcement"
        "Value"                    = $SharepointTenantIPAddressEnforcement
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.IPAddressAllowList)) -EQ $FALSE) {
        $SharepointTenantIPAddressAllowList = $($SharepointTenant.IPAddressAllowList|out-string).trim()
    }
    else{
        $SharepointTenantIPAddressAllowList = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "IP address allow list"
        "Value"                    = $SharepointTenantIPAddressAllowList
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.IPAddressWACTokenLifetime)) -EQ $FALSE) {
        $SharepointTenantIPAddressWACTokenLifetime = $($SharepointTenant.IPAddressWACTokenLifetime|out-string).trim()
    }
    else{
        $SharepointTenantIPAddressWACTokenLifetime = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "IP address WAC token lifetime"
        "Value"                    = $SharepointTenantIPAddressWACTokenLifetime
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.UseFindPeopleInPeoplePicker)) -EQ $FALSE) {
        $SharepointTenantUseFindPeopleInPeoplePicker = $($SharepointTenant.UseFindPeopleInPeoplePicker|out-string).trim()
    }
    else{
        $SharepointTenantUseFindPeopleInPeoplePicker = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Use find people in people picker"
        "Value"                    = $SharepointTenantUseFindPeopleInPeoplePicker
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.DefaultSharingLinkType)) -EQ $FALSE) {
        $SharepointTenantDefaultSharingLinkType = $($SharepointTenant.DefaultSharingLinkType|out-string).trim()
    }
    else{
        $SharepointTenantDefaultSharingLinkType = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Default sharing link type"
        "Value"                    = $SharepointTenantDefaultSharingLinkType
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.ODBMembersCanShare)) -EQ $FALSE) {
        $SharepointTenantODBMembersCanShare = $($SharepointTenant.ODBMembersCanShare|out-string).trim()
    }
    else{
        $SharepointTenantODBMembersCanShare = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "ODB members can share"
        "Value"                    = $SharepointTenantODBMembersCanShare
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.ODBAccessRequests)) -EQ $FALSE) {
        $SharepointTenantODBAccessRequests = $($SharepointTenant.ODBAccessRequests|out-string).trim()
    }
    else{
        $SharepointTenantODBAccessRequests = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "ODB access requests"
        "Value"                    = $SharepointTenantODBAccessRequests
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.PreventExternalUsersFromResharing)) -EQ $FALSE) {
        $SharepointTenantPreventExternalUsersFromResharing = $($SharepointTenant.PreventExternalUsersFromResharing|out-string).trim()
    }
    else{
        $SharepointTenantPreventExternalUsersFromResharing = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Prevent external users from resharing"
        "Value"                    = $SharepointTenantPreventExternalUsersFromResharing
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.FileAnonymousLinkType)) -EQ $FALSE) {
        $SharepointTenantFileAnonymousLinkType = $($SharepointTenant.FileAnonymousLinkType|out-string).trim()
    }
    else{
        $SharepointTenantFileAnonymousLinkType = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "File anonymous link type"
        "Value"                    = $SharepointTenantFileAnonymousLinkType
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.FolderAnonymousLinkType)) -EQ $FALSE) {
        $SharepointTenantFolderAnonymousLinkType = $($SharepointTenant.FolderAnonymousLinkType|out-string).trim()
    }
    else{
        $SharepointTenantFolderAnonymousLinkType = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Folder anonymous link type"
        "Value"                    = $SharepointTenantFolderAnonymousLinkType
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.NotifyOwnersWhenItemsReshared)) -EQ $FALSE) {
        $SharepointTenantNotifyOwnersWhenItemsReshared = $($SharepointTenant.NotifyOwnersWhenItemsReshared|out-string).trim()
    }
    else{
        $SharepointTenantNotifyOwnersWhenItemsReshared = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Notify owners when items are reshared"
        "Value"                    = $SharepointTenantNotifyOwnersWhenItemsReshared
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.NotifyOwnersWhenInvitationsAccepted)) -EQ $FALSE) {
        $SharepointTenantNotifyOwnersWhenInvitationsAccepted = $($SharepointTenant.NotifyOwnersWhenInvitationsAccepted|out-string).trim()
    }
    else{
        $SharepointTenantNotifyOwnersWhenInvitationsAccepted = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Notify owners when invitations accepted"
        "Value"                    = $SharepointTenantNotifyOwnersWhenInvitationsAccepted
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.NotificationsInOneDriveForBusinessEnabled)) -EQ $FALSE) {
        $SharepointTenantNotificationsInOneDriveForBusinessEnabled = $($SharepointTenant.NotificationsInOneDriveForBusinessEnabled|out-string).trim()
    }
    else{
        $SharepointTenantNotificationsInOneDriveForBusinessEnabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Notifications in OneDrive for business enabled"
        "Value"                    = $SharepointTenantNotificationsInOneDriveForBusinessEnabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.NotificationsInSharePointEnabled)) -EQ $FALSE) {
        $SharepointTenantNotificationsInSharePointEnabled = $($SharepointTenant.NotificationsInSharePointEnabled|out-string).trim()
    }
    else{
        $SharepointTenantNotificationsInSharePointEnabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Notifications in SharePoint enabled"
        "Value"                    = $SharepointTenantNotificationsInSharePointEnabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.SpecialCharactersStateInFileFolderNames)) -EQ $FALSE) {
        $SharepointTenantSpecialCharactersStateInFileFolderNames = $($SharepointTenant.SpecialCharactersStateInFileFolderNames|out-string).trim()
    }
    else{
        $SharepointTenantSpecialCharactersStateInFileFolderNames = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Special characters state in file folder names"
        "Value"                    = $SharepointTenantSpecialCharactersStateInFileFolderNames
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.OwnerAnonymousNotification)) -EQ $FALSE) {
        $SharepointTenantOwnerAnonymousNotification = $($SharepointTenant.OwnerAnonymousNotification|out-string).trim()
    }
    else{
        $SharepointTenantOwnerAnonymousNotification = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Owner anonymous notification"
        "Value"                    = $SharepointTenantOwnerAnonymousNotification
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.CommentsOnSitePagesDisabled)) -EQ $FALSE) {
        $SharepointTenantCommentsOnSitePagesDisabled = $($SharepointTenant.CommentsOnSitePagesDisabled|out-string).trim()
    }
    else{
        $SharepointTenantCommentsOnSitePagesDisabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Comments on site pages disabled"
        "Value"                    = $SharepointTenantCommentsOnSitePagesDisabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.CommentsOnFilesDisabled)) -EQ $FALSE) {
        $SharepointTenantCommentsOnFilesDisabled = $($SharepointTenant.CommentsOnFilesDisabled|out-string).trim()
    }
    else{
        $SharepointTenantCommentsOnFilesDisabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Comments on files disabled"
        "Value"                    = $SharepointTenantCommentsOnFilesDisabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.SocialBarOnSitePagesDisabled)) -EQ $FALSE) {
        $SharepointTenantSocialBarOnSitePagesDisabled = $($SharepointTenant.SocialBarOnSitePagesDisabled|out-string).trim()
    }
    else{
        $SharepointTenantSocialBarOnSitePagesDisabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Social bar on site pages disabled"
        "Value"                    = $SharepointTenantSocialBarOnSitePagesDisabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.orphanedPersonalSitesRetentionPeriod)) -EQ $FALSE) {
        $SharepointTenantorphanedPersonalSitesRetentionPeriod = $($SharepointTenant.orphanedPersonalSitesRetentionPeriod|out-string).trim()
    }
    else{
        $SharepointTenantorphanedPersonalSitesRetentionPeriod = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Orphaned personal site retention period"
        "Value"                    = $SharepointTenantorphanedPersonalSitesRetentionPeriod
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.DisallowInfectedFileDownload)) -EQ $FALSE) {
        $SharepointTenantDisallowInfectedFileDownload= $($SharepointTenant.DisallowInfectedFileDownload|out-string).trim()
    }
    else{
        $SharepointTenantDisallowInfectedFileDownload = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Disallow infected file download"
        "Value"                    = $SharepointTenantDisallowInfectedFileDownload
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.DefaultLinkPermission)) -EQ $FALSE) {
        $SharepointTenantDefaultLinkPermission = $($SharepointTenant.DefaultLinkPermission|out-string).trim()
    }
    else{
        $SharepointTenantDefaultLinkPermission = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Default link permission"
        "Value"                    = $SharepointTenantDefaultLinkPermission
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.AllowDownloadingNonWebViewableFiles)) -EQ $FALSE) {
        $SharepointTenantAllowDownloadingNonWebViewableFiles = $($SharepointTenant.AllowDownloadingNonWebViewableFiles|out-string).trim()
    }
    else{
        $SharepointTenantAllowDownloadingNonWebViewableFiles = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Allow downloading nonweb viewable files"
        "Value"                    = $SharepointTenantAllowDownloadingNonWebViewableFiles
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.LimitedAccessFileType)) -EQ $FALSE) {
        $SharepointTenantLimitedAccessFileType = $($SharepointTenant.LimitedAccessFileType|out-string).trim()
    }
    else{
        $SharepointTenantLimitedAccessFileType = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Limited access file type"
        "Value"                    = $SharepointTenantLimitedAccessFileType
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.ApplyAppEnforcedRestrictionsToAdHocRecipients)) -EQ $FALSE) {
        $SharepointTenantApplyAppEnforcedRestrictionsToAdHocRecipients = $($SharepointTenant.ApplyAppEnforcedRestrictionsToAdHocRecipients|out-string).trim()
    }
    else{
        $SharepointTenantApplyAppEnforcedRestrictionsToAdHocRecipients = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Apply app enforced restrictions to adhoc recipients"
        "Value"                    = $SharepointTenantApplyAppEnforcedRestrictionsToAdHocRecipients
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.FilePickerExternalImageSearchEnabled)) -EQ $FALSE) {
        $SharepointTenantFilePickerExternalImageSearchEnabled = $($SharepointTenant.FilePickerExternalImageSearchEnabled|out-string).trim()
    }
    else{
        $SharepointTenantFilePickerExternalImageSearchEnabled = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "File picker external image search enabled"
        "Value"                    = $SharepointTenantFilePickerExternalImageSearchEnabled
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.EmailAttestationRequired)) -EQ $FALSE) {
        $SharepointTenantEmailAttestationRequired = $($SharepointTenant.EmailAttestationRequired|out-string).trim()
    }
    else{
        $SharepointTenantEmailAttestationRequired = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Email attestation required"
        "Value"                    = $SharepointTenantEmailAttestationRequired
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.EmailAttestationReAuthDays)) -EQ $FALSE) {
        $SharepointTenantEmailAttestationReAuthDays = $($SharepointTenant.EmailAttestationReAuthDays|out-string).trim()
    }
    else{
        $SharepointTenantEmailAttestationReAuthDays = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Email attestation reauth days"
        "Value"                    = $SharepointTenantEmailAttestationReAuthDays
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.SyncPrivacyProfileProperties)) -EQ $FALSE) {
        $SharepointTenantSyncPrivacyProfileProperties = $($SharepointTenant.SyncPrivacyProfileProperties|out-string).trim()
    }
    else{
        $SharepointTenantSyncPrivacyProfileProperties = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Sync privacy profile properties"
        "Value"                    = $SharepointTenantSyncPrivacyProfileProperties
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.MarkNewFilesSensitiveByDefault)) -EQ $FALSE) {
        $SharepointTenantMarkNewFilesSensitiveByDefault = $($SharepointTenant.MarkNewFilesSensitiveByDefault|out-string).trim()
    }
    else{
        $SharepointTenantMarkNewFilesSensitiveByDefault = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Mark new files sensitive by default"
        "Value"                    = $SharepointTenantMarkNewFilesSensitiveByDefault
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject

    If (([string]::isnullorempty( $SharepointTenant.EnableAIPIntegration)) -EQ $FALSE) {
        $SharepointTenantEnableAIPIntegration = $($SharepointTenant.EnableAIPIntegration|out-string).trim()
    }
    else{
        $SharepointTenantEnableAIPIntegration = "Not Configured"
    }
    $SharepointTenantDetail = [ordered]@{
        "Configuration Item"       = "Enable AIP integration"
        "Value"                    = $SharepointTenantEnableAIPIntegration
    }
    $SharepointConfigurationObject = New-Object -TypeName psobject -Property $SharepointTenantDetail
    $SharepointArray += $SharepointConfigurationObject
}

######################################################################################################################################################################################################################################################################################################
#############################################
#Teams
#############################################
Write-Host "Connecting to Teams" -foregroundcolor Yellow
Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1"  
$sfbSession = New-CsOnlineSession -UserName $UserPrincipalName -OverrideAdminDomain $admindomain
Import-PSSession $sfbSession -allowclobber
Write-Host "Querying Teams configuration..." -foregroundcolor Yellow

#############################################
#Teams - Client configuration
#############################################
Write-Host " - Tenant Configuration" -foregroundcolor Gray
$TeamsClientConfiguration = Get-CsTeamsClientConfiguration

If ($null -eq  $TeamsClientConfiguration) {
     $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject
}
Else {
    If (([string]::isnullorempty( $TeamsClientConfiguration.Identity)) -EQ $FALSE) {
        $TeamsClientConfigurationName = $($TeamsClientConfiguration.Identity|out-string).trim()
    }
    else {
        $TeamsClientConfigurationName = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Name"
        "Value"                    = $TeamsClientConfigurationName
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowEmailIntoChannel)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowEmailIntoChannel = $($TeamsClientConfiguration.AllowEmailIntoChannel|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowEmailIntoChannel = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow email into the channel"
        "Value"                    = $TeamsClientConfigurationAllowEmailIntoChannel
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.RestrictedSenderList)) -EQ $FALSE) {
        $TeamsClientConfigurationRestrictedSenderList = $($TeamsClientConfiguration.RestrictedSenderList|out-string).trim()
    }
    else {
        $TeamsClientConfigurationRestrictedSenderList = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Restricted sender list"
        "Value"                    = $TeamsClientConfigurationRestrictedSenderList
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowDropBox)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowDropBox = $($TeamsClientConfiguration.AllowDropBox|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowDropBox = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow DropBox"
        "Value"                    = $TeamsClientConfigurationAllowDropBox
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowBox)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowBox = $($TeamsClientConfiguration.AllowBox|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowBox = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow box"
        "Value"                    = $TeamsClientConfigurationAllowBox
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowGoogleDrive)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowGoogleDrive = $($TeamsClientConfiguration.AllowGoogleDrive|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowGoogleDrive = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow Google drive"
        "Value"                    = $TeamsClientConfigurationAllowGoogleDrive
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowShareFile)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowShareFile = $($TeamsClientConfiguration.AllowShareFile|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowShareFile = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow Sharefile"
        "Value"                    = $TeamsClientConfigurationAllowShareFile
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowOrganizationTab)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowOrganizationTab = $($TeamsClientConfiguration.AllowOrganizationTab|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowOrganizationTab = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow organisation tab"
        "Value"                    = $TeamsClientConfigurationAllowOrganizationTab
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowSkypeBusinessInterop)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowSkypeBusinessInterop = $($TeamsClientConfiguration.AllowSkypeBusinessInterop|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowSkypeBusinessInterop = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow Skype for business interop"
        "Value"                    = $TeamsClientConfigurationAllowSkypeBusinessInterop
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.ContentPin)) -EQ $FALSE) {
        $TeamsClientConfigurationContentPin = $($TeamsClientConfiguration.ContentPin|out-string).trim()
    }
    else {
        $TeamsClientConfigurationContentPin = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Content pin"
        "Value"                    = $TeamsClientConfigurationContentPin
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowResourceAccountSendMessage)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowResourceAccountSendMessage = $($TeamsClientConfiguration.AllowResourceAccountSendMessage|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowResourceAccountSendMessage = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow resource account send message"
        "Value"                    = $TeamsClientConfigurationAllowResourceAccountSendMessage
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.ResourceAccountContentAccess)) -EQ $FALSE) {
        $TeamsClientConfigurationResourceAccountContentAccess = $($TeamsClientConfiguration.ResourceAccountContentAccess|out-string).trim()
    }
    else {
        $TeamsClientConfigurationResourceAccountContentAccess = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Resource account content access"
        "Value"                    = $TeamsClientConfigurationResourceAccountContentAccess
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowGuestUser)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowGuestUser = $($TeamsClientConfiguration.AllowGuestUser|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowGuestUser = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow guest user"
        "Value"                    = $TeamsClientConfigurationAllowGuestUser
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject

    If (([string]::isnullorempty( $TeamsClientConfiguration.AllowScopedPeopleSearchandAccess)) -EQ $FALSE) {
        $TeamsClientConfigurationAllowScopedPeopleSearchandAccess = $($TeamsClientConfiguration.AllowScopedPeopleSearchandAccess|out-string).trim()
    }
    else {
        $TeamsClientConfigurationAllowScopedPeopleSearchandAccess = "Not Configured"
    }
    $TeamsClientConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Allow scoped people search and access"
        "Value"                    = $TeamsClientConfigurationAllowScopedPeopleSearchandAccess
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsClientConfigurationdetail
    $TeamsClientConfigArray += $TeamsConfigurationObject
}

#############################################
#Teams - Channels Policy
#############################################
Write-Host " - Channels Policy" -foregroundcolor Gray
$TeamsChannelConfiguration = Get-CsTeamsChannelsPolicy

If ($null -eq  $TeamsChannelConfiguration) {
     $TeamsChannelConfigurationdetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
    $TeamsChannelPolicyArray += $TeamsConfigurationObject
}
Else {
    If($TeamsChannelConfiguration -isnot [array]) {
        If (([string]::isnullorempty( $TeamsChannelConfiguration.identity)) -EQ $FALSE) {
            $TeamsChannelConfigurationidentity = $($TeamsChannelConfiguration.identity|out-string).trim()
        }
        else {
            $TeamsChannelConfigurationidentity = "Not Configured"
        }
        $TeamsChannelConfigurationdetail = [ordered]@{
            "Configuration Item"       = "Name [TBA]"
            "Value"                    = $TeamsChannelConfigurationidentity
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
        $TeamsChannelPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsChannelConfiguration.Description)) -EQ $FALSE) {
            $TeamsChannelConfigurationDescription = $($TeamsChannelConfiguration.Description|out-string).trim()
        }
        else {
            $TeamsChannelConfigurationDescription = "Not Configured"
        }
        $TeamsChannelConfigurationdetail = [ordered]@{
            "Configuration Item"       = "Description"
            "Value"                    = $TeamsChannelConfigurationDescription
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
        $TeamsChannelPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsChannelConfiguration.AllowOrgWideTeamCreation)) -EQ $FALSE) {
            $TeamsChannelConfigurationAllowOrgWideTeamCreation = $($TeamsChannelConfiguration.AllowOrgWideTeamCreation|out-string).trim()
        }
        else {
            $TeamsChannelConfigurationAllowOrgWideTeamCreation = "Not Configured"
        }
        $TeamsChannelConfigurationdetail = [ordered]@{
            "Configuration Item"       = "Org wide Team creation"
            "Value"                    = $TeamsChannelConfigurationAllowOrgWideTeamCreation
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
        $TeamsChannelPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsChannelConfiguration.AllowPrivateTeamDiscovery)) -EQ $FALSE) {
            $TeamsChannelConfigurationAllowPrivateTeamDiscovery = $($TeamsChannelConfiguration.AllowPrivateTeamDiscovery|out-string).trim()
        }
        else {
            $TeamsChannelConfigurationAllowPrivateTeamDiscovery = "Not Configured"
        }
        $TeamsChannelConfigurationdetail = [ordered]@{
            "Configuration Item"       = "Private Team discovery"
            "Value"                    = $TeamsChannelConfigurationAllowPrivateTeamDiscovery
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
        $TeamsChannelPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsChannelConfiguration.AllowPrivateChannelCreation)) -EQ $FALSE) {
            $TeamsChannelConfigurationAllowPrivateChannelCreation = $($TeamsChannelConfiguration.AllowPrivateChannelCreation|out-string).trim()
        }
        else {
            $TeamsChannelConfigurationAllowPrivateChannelCreation = "Not Configured"
        }
        $TeamsChannelConfigurationdetail = [ordered]@{
            "Configuration Item"       = "Allow private channel creation"
            "Value"                    = $TeamsChannelConfigurationAllowPrivateChannelCreation
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
        $TeamsChannelPolicyArray += $TeamsConfigurationObject
    }
    else {
        $TeamsChannelConfiguration |ForEach-Object {
            If (([string]::isnullorempty( $_.identity)) -EQ $FALSE) {
                $TeamsChannelConfigurationidentity = $($_.identity|out-string).trim()
            }
            else {
                $TeamsChannelConfigurationidentity = "Not Configured"
            }
            $TeamsChannelConfigurationdetail = [ordered]@{
                "Configuration Item"       = "Name [TBA]"
                "Value"                    = $TeamsChannelConfigurationidentity
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
            $TeamsChannelPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty( $_.Description)) -EQ $FALSE) {
                $TeamsChannelConfigurationDescription = $($_.Description|out-string).trim()
            }
            else {
                $TeamsChannelConfigurationDescription = "Not Configured"
            }
            $TeamsChannelConfigurationdetail = [ordered]@{
                "Configuration Item"       = "Description"
                "Value"                    = $TeamsChannelConfigurationDescription
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
            $TeamsChannelPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty( $_.AllowOrgWideTeamCreation)) -EQ $FALSE) {
                $TeamsChannelConfigurationAllowOrgWideTeamCreation = $($_.AllowOrgWideTeamCreation|out-string).trim()
            }
            else {
                $TeamsChannelConfigurationAllowOrgWideTeamCreation = "Not Configured"
            }
            $TeamsChannelConfigurationdetail = [ordered]@{
                "Configuration Item"       = "Org wide Team creation"
                "Value"                    = $TeamsChannelConfigurationAllowOrgWideTeamCreation
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
            $TeamsChannelPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty( $_.AllowPrivateTeamDiscovery)) -EQ $FALSE) {
                $TeamsChannelConfigurationAllowPrivateTeamDiscovery = $($_.AllowPrivateTeamDiscovery|out-string).trim()
            }
            else {
                $TeamsChannelConfigurationAllowPrivateTeamDiscovery = "Not Configured"
            }
            $TeamsChannelConfigurationdetail = [ordered]@{
                "Configuration Item"       = "Private Team discovery"
                "Value"                    = $TeamsChannelConfigurationAllowPrivateTeamDiscovery
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
            $TeamsChannelPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty( $_.AllowPrivateChannelCreation)) -EQ $FALSE) {
                $TeamsChannelConfigurationAllowPrivateChannelCreation = $($_.AllowPrivateChannelCreation|out-string).trim()
            }
            else {
                $TeamsChannelConfigurationAllowPrivateChannelCreation = "Not Configured"
            }
            $TeamsChannelConfigurationdetail = [ordered]@{
                "Configuration Item"       = "Allow private channel creation"
                "Value"                    = $TeamsChannelConfigurationAllowPrivateChannelCreation
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsChannelConfigurationdetail
            $TeamsChannelPolicyArray += $TeamsConfigurationObject
        }
    }
}


#############################################
#Teams - Calling Policy
#############################################
Write-Host " - Calling Policy" -foregroundcolor Gray
$TeamsCallingPolicy = Get-CsTeamsCallingPolicy

If ($null -eq  $TeamsCallingPolicy) {
     $TeamsCallingPolicydetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
    $TeamsCallingPolicyArray += $TeamsConfigurationObject
}
Else {
    If($TeamsCallingPolicy -isnot [array]) {
        If (([string]::isnullorempty( $TeamsCallingPolicy.Identity)) -EQ $FALSE) {
            $TeamsCallingPolicyIdentity = $($TeamsCallingPolicy.Identity|out-string).trim()
        }
        else {
            $TeamsCallingPolicyIdentity = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Name [TBA]"
            "Value"                    = $TeamsCallingPolicyIdentity
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.Description)) -EQ $FALSE) {
            $TeamsCallingPolicyDescription = $($TeamsCallingPolicy.Description|out-string).trim()
        }
        else {
            $TeamsCallingPolicyDescription = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Description"
            "Value"                    = $TeamsCallingPolicyDescription
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.AllowPrivateCalling)) -EQ $FALSE) {
            $TeamsCallingPolicyAllowPrivateCalling = $($TeamsCallingPolicy.AllowPrivateCalling|out-string).trim()
        }
        else {
            $TeamsCallingPolicyAllowPrivateCalling = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Allow private calling"
            "Value"                    = $TeamsCallingPolicyAllowPrivateCalling
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.AllowVoicemail)) -EQ $FALSE) {
            $TeamsCallingPolicyAllowVoicemail = $($TeamsCallingPolicy.AllowVoicemail|out-string).trim()
        }
        else {
            $TeamsCallingPolicyAllowVoicemail = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Allow voicemail"
            "Value"                    = $TeamsCallingPolicyAllowVoicemail
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.AllowCallGroups)) -EQ $FALSE) {
            $TeamsCallingPolicyAllowCallGroups = $($TeamsCallingPolicy.AllowCallGroups|out-string).trim()
        }
        else {
            $TeamsCallingPolicyAllowCallGroups = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Allow call groups"
            "Value"                    = $TeamsCallingPolicyAllowCallGroups
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.AllowDelegation)) -EQ $FALSE) {
            $TeamsCallingPolicyAllowDelegation = $($TeamsCallingPolicy.AllowDelegation|out-string).trim()
        }
        else {
            $TeamsCallingPolicyAllowDelegation = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Allow delegation"
            "Value"                    = $TeamsCallingPolicyAllowDelegation
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.AllowCallForwardingToUser)) -EQ $FALSE) {
            $TeamsCallingPolicyAllowCallForwardingToUser = $($TeamsCallingPolicy.AllowCallForwardingToUser|out-string).trim()
        }
        else {
            $TeamsCallingPolicyAllowCallForwardingToUser = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Allow call forwarding to user"
            "Value"                    = $TeamsCallingPolicyAllowCallForwardingToUser
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.AllowCallForwardingToPhone)) -EQ $FALSE) {
            $TeamsCallingPolicyAllowCallForwardingToPhone = $($TeamsCallingPolicy.AllowCallForwardingToPhone|out-string).trim()
        }
        else {
            $TeamsCallingPolicyAllowCallForwardingToPhone = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Allow call forwarding to phone"
            "Value"                    = $TeamsCallingPolicyAllowCallForwardingToPhone
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.PreventTollBypass)) -EQ $FALSE) {
            $TeamsCallingPolicyPreventTollBypass = $($TeamsCallingPolicy.PreventTollBypass|out-string).trim()
        }
        else {
            $TeamsCallingPolicyIdentity = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Prevent toll bypass"
            "Value"                    = $TeamsCallingPolicyPreventTollBypass
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty( $TeamsCallingPolicy.BusyOnBusyEnabledType)) -EQ $FALSE) {
            $TeamsCallingPolicyBusyOnBusyEnabledType = $($TeamsCallingPolicy.BusyOnBusyEnabledType|out-string).trim()
        }
        else {
            $TeamsCallingPolicyBusyOnBusyEnabledType = "Not Configured"
        }
        $TeamsCallingPolicydetail = [ordered]@{
            "Configuration Item"       = "Busy on busy enabled type"
            "Value"                    = $TeamsCallingPolicyBusyOnBusyEnabledType
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
        $TeamsCallingPolicyArray += $TeamsConfigurationObject
    }
    else {
        $TeamsCallingPolicy |Foreach-Object {
            If (([string]::isnullorempty( $_.Identity)) -EQ $FALSE) {
                $TeamsCallingPolicyIdentity = $($_.Identity|out-string).trim()
            }
            else {
                $TeamsCallingPolicyIdentity = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Name [TBA]"
                "Value"                    = $TeamsCallingPolicyIdentity
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.Description)) -EQ $FALSE) {
                $TeamsCallingPolicyDescription = $($_.Description|out-string).trim()
            }
            else {
                $TeamsCallingPolicyDescription = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Description"
                "Value"                    = $TeamsCallingPolicyDescription
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.AllowPrivateCalling)) -EQ $FALSE) {
                $TeamsCallingPolicyAllowPrivateCalling = $($_.AllowPrivateCalling|out-string).trim()
            }
            else {
                $TeamsCallingPolicyAllowPrivateCalling = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow private calling"
                "Value"                    = $TeamsCallingPolicyAllowPrivateCalling
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.AllowVoicemail)) -EQ $FALSE) {
                $TeamsCallingPolicyAllowVoicemail = $($_.AllowVoicemail|out-string).trim()
            }
            else {
                $TeamsCallingPolicyAllowVoicemail = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow voicemail"
                "Value"                    = $TeamsCallingPolicyAllowVoicemail
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.AllowCallGroups)) -EQ $FALSE) {
                $TeamsCallingPolicyAllowCallGroups = $($_.AllowCallGroups|out-string).trim()
            }
            else {
                $TeamsCallingPolicyAllowCallGroups = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow call groups"
                "Value"                    = $TeamsCallingPolicyAllowCallGroups
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.AllowDelegation)) -EQ $FALSE) {
                $TeamsCallingPolicyAllowDelegation = $($_.AllowDelegation|out-string).trim()
            }
            else {
                $TeamsCallingPolicyAllowDelegation = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow delegation"
                "Value"                    = $TeamsCallingPolicyAllowDelegation
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.AllowCallForwardingToUser)) -EQ $FALSE) {
                $TeamsCallingPolicyAllowCallForwardingToUser = $($_.AllowCallForwardingToUser|out-string).trim()
            }
            else {
                $TeamsCallingPolicyAllowCallForwardingToUser = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow call forwarding to user"
                "Value"                    = $TeamsCallingPolicyAllowCallForwardingToUser
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.AllowCallForwardingToPhone)) -EQ $FALSE) {
                $TeamsCallingPolicyAllowCallForwardingToPhone = $($_.AllowCallForwardingToPhone|out-string).trim()
            }
            else {
                $TeamsCallingPolicyAllowCallForwardingToPhone = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow call forwarding to phone"
                "Value"                    = $TeamsCallingPolicyAllowCallForwardingToPhone
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.PreventTollBypass)) -EQ $FALSE) {
                $TeamsCallingPolicyPreventTollBypass = $($_.PreventTollBypass|out-string).trim()
            }
            else {
                $TeamsCallingPolicyIdentity = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Prevent toll bypass"
                "Value"                    = $TeamsCallingPolicyPreventTollBypass
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject

            If (([string]::isnullorempty( $_.BusyOnBusyEnabledType)) -EQ $FALSE) {
                $TeamsCallingPolicyBusyOnBusyEnabledType = $($_.BusyOnBusyEnabledType|out-string).trim()
            }
            else {
                $TeamsCallingPolicyBusyOnBusyEnabledType = "Not Configured"
            }
            $TeamsCallingPolicydetail = [ordered]@{
                "Configuration Item"       = "Busy on busy enabled type"
                "Value"                    = $TeamsCallingPolicyBusyOnBusyEnabledType
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsCallingPolicydetail
            $TeamsCallingPolicyArray += $TeamsConfigurationObject
        
        }
    }
}

#############################################
#Teams - Meeting Configuration & Meeting Policy
#############################################
Write-Host " - Meeting Configuration & Meeting Policy" -foregroundcolor Gray
$TeamsMeetingConfig = Get-CsTeamsMeetingConfiguration

If ($null -eq  $TeamsMeetingConfig) {
     $TeamsMeetingConfigdetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
    $TeamsMeetingConfigArray += $TeamsConfigurationObject
}
Else {
    If($TeamsMeetingConfig -isnot [array]) {
        If (([string]::isnullorempty($TeamsMeetingConfig.Identity)) -EQ $FALSE) {
            $TeamsMeetingConfigIdentity = $($TeamsMeetingConfig.Identity|out-string).trim()
        }
        else {
            $TeamsMeetingConfigIdentity = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Name [TBA]"
               "Value"                    = $TeamsMeetingConfigIdentity
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
        
        If (([string]::isnullorempty($TeamsMeetingConfig.LogoURL)) -EQ $FALSE) {
            $TeamsMeetingConfigLogoURL = $($TeamsMeetingConfig.LogoURL|out-string).trim()
        }
        else {
            $TeamsMeetingConfigLogoURL = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Logo URL"
               "Value"                    = $TeamsMeetingConfigLogoURL
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
            
        If (([string]::isnullorempty($TeamsMeetingConfig.LegalURL)) -EQ $FALSE) {
            $TeamsMeetingConfigLegalURL = $($TeamsMeetingConfig.LegalURL|out-string).trim()
        }
        else {
            $TeamsMeetingConfigLegalURL = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Legal URL"
               "Value"                    = $TeamsMeetingConfigLegalURL
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.HelpURL)) -EQ $FALSE) {
            $TeamsMeetingConfigHelpURL = $($TeamsMeetingConfig.HelpURL|out-string).trim()
        }
        else {
            $TeamsMeetingConfigHelpURL = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Help URL"
               "Value"                    = $TeamsMeetingConfigHelpURL
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.CustomFooterText)) -EQ $FALSE) {
            $TeamsMeetingConfigCustomFooterText = $($TeamsMeetingConfig.CustomFooterText|out-string).trim()
        }
        else {
            $TeamsMeetingConfigCustomFooterText = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Custom Footer Text"
               "Value"                    = $TeamsMeetingConfigCustomFooterText
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.DisableAnonymousJoin)) -EQ $FALSE) {
            $TeamsMeetingConfigDisableAnonymousJoin = $($TeamsMeetingConfig.DisableAnonymousJoin|out-string).trim()
        }
        else {
            $TeamsMeetingConfigDisableAnonymousJoin = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Disable Anonymous Join"
               "Value"                    = $TeamsMeetingConfigDisableAnonymousJoin
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.EnableQoS)) -EQ $FALSE) {
            $TeamsMeetingConfigEnableQoS = $($TeamsMeetingConfig.EnableQoS|out-string).trim()
        }
        else {
            $TeamsMeetingConfigEnableQoS = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Enable QoS"
               "Value"                    = $TeamsMeetingConfigEnableQoS
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.ClientAudioPort)) -EQ $FALSE) {
            $TeamsMeetingConfigClientAudioPort = $($TeamsMeetingConfig.ClientAudioPort|out-string).trim()
        }
        else {
            $TeamsMeetingConfigClientAudioPort = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Client Audio Port"
               "Value"                    = $TeamsMeetingConfigClientAudioPort
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.ClientAudioPortRange)) -EQ $FALSE) {
            $TeamsMeetingConfigClientAudioPortRange = $($TeamsMeetingConfig.ClientAudioPortRange|out-string).trim()
        }
        else {
            $TeamsMeetingConfigIdentity = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Client Audio Port Range"
               "Value"                    = $TeamsMeetingConfigClientAudioPortRange
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.ClientVideoPort)) -EQ $FALSE) {
            $TeamsMeetingConfigClientVideoPort = $($TeamsMeetingConfig.ClientVideoPort|out-string).trim()
        }
        else {
            $TeamsMeetingConfigClientVideoPort = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Client Video Port"
               "Value"                    = $TeamsMeetingConfigClientVideoPort
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.ClientVideoPortRange)) -EQ $FALSE) {
            $TeamsMeetingConfigClientVideoPortRange = $($TeamsMeetingConfig.ClientVideoPortRange|out-string).trim()
        }
        else {
            $TeamsMeetingConfigClientVideoPortRange = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Client Video Port Range"
               "Value"                    = $TeamsMeetingConfigClientVideoPortRange
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.ClientAppSharingPort)) -EQ $FALSE) {
            $TeamsMeetingConfigClientAppSharingPort = $($TeamsMeetingConfig.ClientAppSharingPort|out-string).trim()
        }
        else {
            $TeamsMeetingConfigClientAppSharingPort = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Client App Sharing Port"
               "Value"                    = $TeamsMeetingConfigClientAppSharingPort
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.ClientAppSharingPortRange)) -EQ $FALSE) {
            $TeamsMeetingConfigClientAppSharingPortRange = $($TeamsMeetingConfig.ClientAppSharingPortRange|out-string).trim()
        }
        else {
            $TeamsMeetingConfigClientAppSharingPortRange = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Client App Sharing Port Range"
               "Value"                    = $TeamsMeetingConfigClientAppSharingPortRange
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
        If (([string]::isnullorempty($TeamsMeetingConfig.ClientMediaPortRangeEnabled)) -EQ $FALSE) {
            $TeamsMeetingConfigClientMediaPortRangeEnabled = $($TeamsMeetingConfig.ClientMediaPortRangeEnabled|out-string).trim()
        }
        else {
            $TeamsMeetingConfigClientMediaPortRangeEnabled = "Not Configured"
        }
            $TeamsMeetingConfigdetail = [ordered]@{
               "Configuration Item"       = "Client Media Port Range Enabled"
               "Value"                    = $TeamsMeetingConfigClientMediaPortRangeEnabled
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
        $TeamsMeetingConfigArray += $TeamsConfigurationObject
    }
    else {
        $TeamsMeetingConfig |foreach-object {
            If (([string]::isnullorempty($_.Identity)) -EQ $FALSE) {
                $TeamsMeetingConfigIdentity = $($_.Identity|out-string).trim()
            }
            else {
                $TeamsMeetingConfigIdentity = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Name [TBA]"
                   "Value"                    = $TeamsMeetingConfigIdentity
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
            
            If (([string]::isnullorempty($_.LogoURL)) -EQ $FALSE) {
                $TeamsMeetingConfigLogoURL = $($_.LogoURL|out-string).trim()
            }
            else {
                $TeamsMeetingConfigLogoURL = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Logo URL"
                   "Value"                    = $TeamsMeetingConfigLogoURL
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                
            If (([string]::isnullorempty($_.LegalURL)) -EQ $FALSE) {
                $TeamsMeetingConfigLegalURL = $($_.LegalURL|out-string).trim()
            }
            else {
                $TeamsMeetingConfigLegalURL = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Legal URL"
                   "Value"                    = $TeamsMeetingConfigLegalURL
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.HelpURL)) -EQ $FALSE) {
                $TeamsMeetingConfigHelpURL = $($_.HelpURL|out-string).trim()
            }
            else {
                $TeamsMeetingConfigHelpURL = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Help URL"
                   "Value"                    = $TeamsMeetingConfigHelpURL
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.CustomFooterText)) -EQ $FALSE) {
                $TeamsMeetingConfigCustomFooterText = $($_.CustomFooterText|out-string).trim()
            }
            else {
                $TeamsMeetingConfigCustomFooterText = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Custom Footer Text"
                   "Value"                    = $TeamsMeetingConfigCustomFooterText
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.DisableAnonymousJoin)) -EQ $FALSE) {
                $TeamsMeetingConfigDisableAnonymousJoin = $($_.DisableAnonymousJoin|out-string).trim()
            }
            else {
                $TeamsMeetingConfigDisableAnonymousJoin = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Disable Anonymous Join"
                   "Value"                    = $TeamsMeetingConfigDisableAnonymousJoin
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.EnableQoS)) -EQ $FALSE) {
                $TeamsMeetingConfigEnableQoS = $($_.EnableQoS|out-string).trim()
            }
            else {
                $TeamsMeetingConfigEnableQoS = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Enable QoS"
                   "Value"                    = $TeamsMeetingConfigEnableQoS
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.ClientAudioPort)) -EQ $FALSE) {
                $TeamsMeetingConfigClientAudioPort = $($_.ClientAudioPort|out-string).trim()
            }
            else {
                $TeamsMeetingConfigClientAudioPort = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Client Audio Port"
                   "Value"                    = $TeamsMeetingConfigClientAudioPort
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.ClientAudioPortRange)) -EQ $FALSE) {
                $TeamsMeetingConfigClientAudioPortRange = $($_.ClientAudioPortRange|out-string).trim()
            }
            else {
                $TeamsMeetingConfigIdentity = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Client Audio Port Range"
                   "Value"                    = $TeamsMeetingConfigClientAudioPortRange
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.ClientVideoPort)) -EQ $FALSE) {
                $TeamsMeetingConfigClientVideoPort = $($_.ClientVideoPort|out-string).trim()
            }
            else {
                $TeamsMeetingConfigClientVideoPort = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Client Video Port"
                   "Value"                    = $TeamsMeetingConfigClientVideoPort
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.ClientVideoPortRange)) -EQ $FALSE) {
                $TeamsMeetingConfigClientVideoPortRange = $($_.ClientVideoPortRange|out-string).trim()
            }
            else {
                $TeamsMeetingConfigClientVideoPortRange = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Client Video Port Range"
                   "Value"                    = $TeamsMeetingConfigClientVideoPortRange
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.ClientAppSharingPort)) -EQ $FALSE) {
                $TeamsMeetingConfigClientAppSharingPort = $($_.ClientAppSharingPort|out-string).trim()
            }
            else {
                $TeamsMeetingConfigClientAppSharingPort = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Client App Sharing Port"
                   "Value"                    = $TeamsMeetingConfigClientAppSharingPort
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.ClientAppSharingPortRange)) -EQ $FALSE) {
                $TeamsMeetingConfigClientAppSharingPortRange = $($_.ClientAppSharingPortRange|out-string).trim()
            }
            else {
                $TeamsMeetingConfigClientAppSharingPortRange = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Client App Sharing Port Range"
                   "Value"                    = $TeamsMeetingConfigClientAppSharingPortRange
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
                    
            If (([string]::isnullorempty($_.ClientMediaPortRangeEnabled)) -EQ $FALSE) {
                $TeamsMeetingConfigClientMediaPortRangeEnabled = $($_.ClientMediaPortRangeEnabled|out-string).trim()
            }
            else {
                $TeamsMeetingConfigClientMediaPortRangeEnabled = "Not Configured"
            }
                $TeamsMeetingConfigdetail = [ordered]@{
                   "Configuration Item"       = "Client Media Port Range Enabled"
                   "Value"                    = $TeamsMeetingConfigClientMediaPortRangeEnabled
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingConfigdetail
            $TeamsMeetingConfigArray += $TeamsConfigurationObject
        }
    }
}

$TeamsMeetingPolicy = Get-CsTeamsMeetingPolicy

If ($null -eq  $TeamsMeetingPolicy) {
     $TeamsMeetingPolicydetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
    $TeamsMeetingPolicyArray += $TeamsConfigurationObject
}
Else {
    If($TeamsMeetingPolicy -isnot [array]) {
        If (([string]::isnullorempty($TeamsMeetingPolicy.Identity)) -EQ $FALSE) {
            $TeamsMeetingPolicyIdentity = $($TeamsMeetingPolicy.Identity|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyIdentity = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Name [TBA]"
               "Value"                    = $TeamsMeetingPolicyIdentity
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.Description)) -EQ $FALSE) {
            $TeamsMeetingPolicyDescription = $($TeamsMeetingPolicy.Description|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyDescription = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Description"
               "Value"                    = $TeamsMeetingPolicyDescription
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowChannelMeetingScheduling)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowChannelMeetingScheduling = $($TeamsMeetingPolicy.AllowChannelMeetingScheduling|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowChannelMeetingScheduling = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Channel Meeting Scheduling"
               "Value"                    = $TeamsMeetingPolicyAllowChannelMeetingScheduling
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowMeetNow)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowMeetNow = $($TeamsMeetingPolicy.AllowMeetNow|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowMeetNow = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Meet Now"
               "Value"                    = $TeamsMeetingPolicyAllowMeetNow
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowPrivateMeetNow)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowPrivateMeetNow = $($TeamsMeetingPolicy.AllowPrivateMeetNow|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowPrivateMeetNow = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Private Meet Now"
               "Value"                    = $TeamsMeetingPolicyAllowPrivateMeetNow
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.MeetingChatEnabledType)) -EQ $FALSE) {
            $TeamsMeetingPolicyMeetingChatEnabledType = $($TeamsMeetingPolicy.MeetingChatEnabledType|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyMeetingChatEnabledType = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Meeting Chat Enabled Type"
               "Value"                    = $TeamsMeetingPolicyMeetingChatEnabledType
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.LiveCaptionsEnabledType)) -EQ $FALSE) {
            $TeamsMeetingPolicyLiveCaptionsEnabledType = $($TeamsMeetingPolicy.LiveCaptionsEnabledType|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyLiveCaptionsEnabledType = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Live Captions Enabled Type"
               "Value"                    = $TeamsMeetingPolicyLiveCaptionsEnabledType
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowIPVideo)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowIPVideo = $($TeamsMeetingPolicy.AllowIPVideo|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowIPVideo = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow IP Video"
               "Value"                    = $TeamsMeetingPolicyAllowIPVideo
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowAnonymousUsersToDialOut)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowAnonymousUsersToDialOut = $($TeamsMeetingPolicy.AllowAnonymousUsersToDialOut|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowAnonymousUsersToDialOut = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Anonymous Users to Dial Out"
               "Value"                    = $TeamsMeetingPolicyAllowAnonymousUsersToDialOut
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowAnonymousUsersToStartMeeting)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowAnonymousUsersToStartMeeting = $($TeamsMeetingPolicy.AllowAnonymousUsersToStartMeeting|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowAnonymousUsersToStartMeeting = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Anonymous Users to Start Meeting"
               "Value"                    = $TeamsMeetingPolicyAllowAnonymousUsersToStartMeeting
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowPrivateMeetingScheduling)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowPrivateMeetingScheduling = $($TeamsMeetingPolicy.AllowPrivateMeetingScheduling|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowPrivateMeetingScheduling = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Private Meeting Scheduling"
               "Value"                    = $TeamsMeetingPolicyAllowPrivateMeetingScheduling
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AutoAdmittedUsers)) -EQ $FALSE) {
            $TeamsMeetingPolicyAutoAdmittedUsers = $($TeamsMeetingPolicy.AutoAdmittedUsers|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAutoAdmittedUsers = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Auto Admitted Users"
               "Value"                    = $TeamsMeetingPolicyAutoAdmittedUsers
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowCloudRecording)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowCloudRecording = $($TeamsMeetingPolicy.AllowCloudRecording|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowCloudRecording = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Cloud Recording"
               "Value"                    = $TeamsMeetingPolicyAllowCloudRecording
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowOutlookAddIn)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowOutlookAddIn = $($TeamsMeetingPolicy.AllowOutlookAddIn|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowOutlookAddIn = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Outlook Add-in"
               "Value"                    = $TeamsMeetingPolicyAllowOutlookAddIn
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowPowerPointSharing)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowPowerPointSharing = $($TeamsMeetingPolicy.AllowPowerPointSharing|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowPowerPointSharing = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow PowerPoint Sharing"
               "Value"                    = $TeamsMeetingPolicyAllowPowerPointSharing
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowParticipantGiveRequestControl)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowParticipantGiveRequestControl = $($TeamsMeetingPolicy.AllowParticipantGiveRequestControl|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowParticipantGiveRequestControl = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Participant Give/Request Control"
               "Value"                    = $TeamsMeetingPolicyAllowParticipantGiveRequestControl
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowExternalParticipantGiveRequestControl)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowExternalParticipantGiveRequestControl = $($TeamsMeetingPolicy.AllowExternalParticipantGiveRequestControl|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowExternalParticipantGiveRequestControl = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow External Participant Give/Request Control"
               "Value"                    = $TeamsMeetingPolicyAllowExternalParticipantGiveRequestControl
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowSharedNotes)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowSharedNotes = $($TeamsMeetingPolicy.AllowSharedNotes|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowSharedNotes = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Shared Notes"
               "Value"                    = $TeamsMeetingPolicyAllowSharedNotes
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowWhiteboard)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowWhiteboard = $($TeamsMeetingPolicy.AllowWhiteboard|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowWhiteboard = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Whiteboard"
               "Value"                    = $TeamsMeetingPolicyAllowWhiteboard
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowTranscription)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowTranscription = $($TeamsMeetingPolicy.AllowTranscription|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowTranscription = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Transcription"
               "Value"                    = $TeamsMeetingPolicyAllowTranscription
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.ScreenSharingMode)) -EQ $FALSE) {
            $TeamsMeetingPolicyScreenSharingMode = $($TeamsMeetingPolicy.ScreenSharingMode|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyScreenSharingMode = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Screen Sharing Mode"
               "Value"                    = $TeamsMeetingPolicyScreenSharingMode
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowPSTNUsersToBypassLobby)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowPSTNUsersToBypassLobby = $($TeamsMeetingPolicy.AllowPSTNUsersToBypassLobby|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowPSTNUsersToBypassLobby = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow PSTN Users to Bypass Lobby"
               "Value"                    = $TeamsMeetingPolicyAllowPSTNUsersToBypassLobby
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingPolicy.AllowOrganizersToOverrideLobbySettings)) -EQ $FALSE) {
            $TeamsMeetingPolicyAllowOrganizersToOverrideLobbySettings = $($TeamsMeetingPolicy.AllowOrganizersToOverrideLobbySettings|out-string).trim()
        }
        else {
            $TeamsMeetingPolicyAllowOrganizersToOverrideLobbySettings = "Not Configured"
        }
            $TeamsMeetingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Organizers to Override Lobby Settings"
               "Value"                    = $TeamsMeetingPolicyAllowOrganizersToOverrideLobbySettings
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
        $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    }
    else {
        $TeamsMeetingPolicy |foreach-object {
            If (([string]::isnullorempty($_.Identity)) -EQ $FALSE) {
                $TeamsMeetingPolicyIdentity = $($_.Identity|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyIdentity = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Name [TBA]"
                   "Value"                    = $TeamsMeetingPolicyIdentity
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.Description)) -EQ $FALSE) {
                $TeamsMeetingPolicyDescription = $($_.Description|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyDescription = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Description"
                   "Value"                    = $TeamsMeetingPolicyDescription
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowChannelMeetingScheduling)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowChannelMeetingScheduling = $($_.AllowChannelMeetingScheduling|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowChannelMeetingScheduling = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Channel Meeting Scheduling"
                   "Value"                    = $TeamsMeetingPolicyAllowChannelMeetingScheduling
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowMeetNow)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowMeetNow = $($_.AllowMeetNow|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowMeetNow = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Meet Now"
                   "Value"                    = $TeamsMeetingPolicyAllowMeetNow
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowPrivateMeetNow)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowPrivateMeetNow = $($_.AllowPrivateMeetNow|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowPrivateMeetNow = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Private Meet Now"
                   "Value"                    = $TeamsMeetingPolicyAllowPrivateMeetNow
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.MeetingChatEnabledType)) -EQ $FALSE) {
                $TeamsMeetingPolicyMeetingChatEnabledType = $($_.MeetingChatEnabledType|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyMeetingChatEnabledType = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Meeting Chat Enabled Type"
                   "Value"                    = $TeamsMeetingPolicyMeetingChatEnabledType
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.LiveCaptionsEnabledType)) -EQ $FALSE) {
                $TeamsMeetingPolicyLiveCaptionsEnabledType = $($_.LiveCaptionsEnabledType|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyLiveCaptionsEnabledType = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Live Captions Enabled Type"
                   "Value"                    = $TeamsMeetingPolicyLiveCaptionsEnabledType
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowIPVideo)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowIPVideo = $($_.AllowIPVideo|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowIPVideo = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow IP Video"
                   "Value"                    = $TeamsMeetingPolicyAllowIPVideo
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowAnonymousUsersToDialOut)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowAnonymousUsersToDialOut = $($_.AllowAnonymousUsersToDialOut|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowAnonymousUsersToDialOut = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Anonymous Users to Dial Out"
                   "Value"                    = $TeamsMeetingPolicyAllowAnonymousUsersToDialOut
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowAnonymousUsersToStartMeeting)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowAnonymousUsersToStartMeeting = $($_.AllowAnonymousUsersToStartMeeting|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowAnonymousUsersToStartMeeting = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Anonymous Users to Start Meeting"
                   "Value"                    = $TeamsMeetingPolicyAllowAnonymousUsersToStartMeeting
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowPrivateMeetingScheduling)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowPrivateMeetingScheduling = $($_.AllowPrivateMeetingScheduling|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowPrivateMeetingScheduling = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Private Meeting Scheduling"
                   "Value"                    = $TeamsMeetingPolicyAllowPrivateMeetingScheduling
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AutoAdmittedUsers)) -EQ $FALSE) {
                $TeamsMeetingPolicyAutoAdmittedUsers = $($_.AutoAdmittedUsers|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAutoAdmittedUsers = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Auto Admitted Users"
                   "Value"                    = $TeamsMeetingPolicyAutoAdmittedUsers
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowCloudRecording)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowCloudRecording = $($_.AllowCloudRecording|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowCloudRecording = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Cloud Recording"
                   "Value"                    = $TeamsMeetingPolicyAllowCloudRecording
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowOutlookAddIn)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowOutlookAddIn = $($_.AllowOutlookAddIn|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowOutlookAddIn = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Outlook Add-in"
                   "Value"                    = $TeamsMeetingPolicyAllowOutlookAddIn
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowPowerPointSharing)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowPowerPointSharing = $($_.AllowPowerPointSharing|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowPowerPointSharing = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow PowerPoint Sharing"
                   "Value"                    = $TeamsMeetingPolicyAllowPowerPointSharing
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowParticipantGiveRequestControl)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowParticipantGiveRequestControl = $($_.AllowParticipantGiveRequestControl|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowParticipantGiveRequestControl = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Participant Give/Request Control"
                   "Value"                    = $TeamsMeetingPolicyAllowParticipantGiveRequestControl
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowExternalParticipantGiveRequestControl)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowExternalParticipantGiveRequestControl = $($_.AllowExternalParticipantGiveRequestControl|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowExternalParticipantGiveRequestControl = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow External Participant Give/Request Control"
                   "Value"                    = $TeamsMeetingPolicyAllowExternalParticipantGiveRequestControl
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowSharedNotes)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowSharedNotes = $($_.AllowSharedNotes|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowSharedNotes = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Shared Notes"
                   "Value"                    = $TeamsMeetingPolicyAllowSharedNotes
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowWhiteboard)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowWhiteboard = $($_.AllowWhiteboard|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowWhiteboard = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Whiteboard"
                   "Value"                    = $TeamsMeetingPolicyAllowWhiteboard
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowTranscription)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowTranscription = $($_.AllowTranscription|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowTranscription = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Transcription"
                   "Value"                    = $TeamsMeetingPolicyAllowTranscription
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.ScreenSharingMode)) -EQ $FALSE) {
                $TeamsMeetingPolicyScreenSharingMode = $($_.ScreenSharingMode|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyScreenSharingMode = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Screen Sharing Mode"
                   "Value"                    = $TeamsMeetingPolicyScreenSharingMode
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowPSTNUsersToBypassLobby)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowPSTNUsersToBypassLobby = $($_.AllowPSTNUsersToBypassLobby|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowPSTNUsersToBypassLobby = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow PSTN Users to Bypass Lobby"
                   "Value"                    = $TeamsMeetingPolicyAllowPSTNUsersToBypassLobby
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowOrganizersToOverrideLobbySettings)) -EQ $FALSE) {
                $TeamsMeetingPolicyAllowOrganizersToOverrideLobbySettings = $($_.AllowOrganizersToOverrideLobbySettings|out-string).trim()
            }
            else {
                $TeamsMeetingPolicyAllowOrganizersToOverrideLobbySettings = "Not Configured"
            }
                $TeamsMeetingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Organizers to Override Lobby Settings"
                   "Value"                    = $TeamsMeetingPolicyAllowOrganizersToOverrideLobbySettings
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingPolicydetail
            $TeamsMeetingPolicyArray += $TeamsConfigurationObject
        }
    }
}

#ATP for teams



#############################################
#Teams - Messaging  Policy
#############################################
Write-Host " - Messaging  Policy" -foregroundcolor Gray
$TeamsMessagingPolicy = Get-CsTeamsMessagingPolicy

If ($null -eq  $TeamsMessagingPolicy) {
     $TeamsTeamsMessagingPolicydetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
    $TeamsMessagingPolicyArray += $TeamsConfigurationObject
}
Else {
    If($TeamsMessagingPolicy -isnot [array]) {
        If (([string]::isnullorempty($TeamsMessagingPolicy.Identity)) -EQ $FALSE) {
            $TeamsMessagingPolicyIdentity = $($TeamsMessagingPolicy.Identity|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyIdentity = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Name [TBA]"
               "Value"                    = $TeamsMessagingPolicyIdentity
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.Description)) -EQ $FALSE) {
            $TeamsMessagingPolicyDescription = $($TeamsMessagingPolicy.Description|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyDescription = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Description"
               "Value"                    = $TeamsMessagingPolicyDescription
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowUrlPreviews)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowUrlPreviews = $($TeamsMessagingPolicy.AllowUrlPreviews|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowUrlPreviews = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow URL previews"
               "Value"                    = $TeamsMessagingPolicyAllowUrlPreviews
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowOwnerDeleteMessage)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowOwnerDeleteMessage = $($TeamsMessagingPolicy.AllowOwnerDeleteMessage|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowOwnerDeleteMessage = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Owner Delete Message"
               "Value"                    = $TeamsMessagingPolicyAllowOwnerDeleteMessage
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowUserEditMessage)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowUserEditMessage = $($TeamsMessagingPolicy.AllowUserEditMessage|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowUserEditMessage = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow User Edit Message"
               "Value"                    = $TeamsMessagingPolicyAllowUserEditMessage
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowUserDeleteMessage)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowUserDeleteMessage = $($TeamsMessagingPolicy.AllowUserDeleteMessage|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowUserDeleteMessage = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow User Delete Message"
               "Value"                    = $TeamsMessagingPolicyAllowUserDeleteMessage
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowUserChat)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowUserChat = $($TeamsMessagingPolicy.AllowUserChat|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowUserChat = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow User Chat"
               "Value"                    = $TeamsMessagingPolicyAllowUserChat
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowRemoveUser)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowRemoveUser = $($TeamsMessagingPolicy.AllowRemoveUser|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowRemoveUser = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Remove User"
               "Value"                    = $TeamsMessagingPolicyAllowRemoveUser
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowGiphy)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowGiphy = $($TeamsMessagingPolicy.AllowGiphy|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowGiphy = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Giphy"
               "Value"                    = $TeamsMessagingPolicyAllowGiphy
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.GiphyRatingType)) -EQ $FALSE) {
            $TeamsMessagingPolicyGiphyRatingType = $($TeamsMessagingPolicy.GiphyRatingType|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyGiphyRatingType = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Giphy Rating Type"
               "Value"                    = $TeamsMessagingPolicyGiphyRatingType
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowMemes)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowMemes = $($TeamsMessagingPolicy.AllowMemes|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowMemes = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Memes"
               "Value"                    = $TeamsMessagingPolicyAllowMemes
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowImmersiveReader)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowImmersiveReader = $($TeamsMessagingPolicy.AllowImmersiveReader|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowImmersiveReader = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Immersive Reader"
               "Value"                    = $TeamsMessagingPolicyAllowImmersiveReader
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowStickers)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowStickers = $($TeamsMessagingPolicy.AllowStickers|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowStickers = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Stickers"
               "Value"                    = $TeamsMessagingPolicyAllowStickers
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowUserTranslation)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowUserTranslation = $($TeamsMessagingPolicy.AllowUserTranslation|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowUserTranslation = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow User Translation"
               "Value"                    = $TeamsMessagingPolicyAllowUserTranslation
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.ReadReceiptsEnabledType)) -EQ $FALSE) {
            $TeamsMessagingPolicyReadReceiptsEnabledType = $($TeamsMessagingPolicy.ReadReceiptsEnabledType|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyIdentity = "Not Configured"
        }
            $TeamsMessagingPolicyReadReceiptsEnabledType = [ordered]@{
               "Configuration Item"       = "Read Recipts Enabled Type"
               "Value"                    = $TeamsMessagingPolicyReadReceiptsEnabledType
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.AllowPriorityMessages)) -EQ $FALSE) {
            $TeamsMessagingPolicyAllowPriorityMessages = $($TeamsMessagingPolicy.AllowPriorityMessages|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAllowPriorityMessages = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Allow Priority Messages"
               "Value"                    = $TeamsMessagingPolicyAllowPriorityMessages
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMessagingPolicy.ChannelsInChatListEnabledType)) -EQ $FALSE) {
            $TeamsMessagingPolicyChannelsInChatListEnabledType = $($TeamsMessagingPolicy.ChannelsInChatListEnabledType|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyChannelsInChatListEnabledType = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Channels in Chat List Enabled Type"
               "Value"                    = $TeamsMessagingPolicyChannelsInChatListEnabledType
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject
        
        If (([string]::isnullorempty($TeamsMessagingPolicy.AudioMessageEnabledType)) -EQ $FALSE) {
            $TeamsMessagingPolicyAudioMessageEnabledType = $($TeamsMessagingPolicy.AudioMessageEnabledType|out-string).trim()
        }
        else {
            $TeamsMessagingPolicyAudioMessageEnabledType = "Not Configured"
        }
            $TeamsTeamsMessagingPolicydetail = [ordered]@{
               "Configuration Item"       = "Audio Message Enabled Type"
               "Value"                    = $TeamsMessagingPolicyAudioMessageEnabledType
           }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
        $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    }
    else {
        $TeamsMessagingPolicy |foreach-object {
            If (([string]::isnullorempty($_.Identity)) -EQ $FALSE) {
                $TeamsMessagingPolicyIdentity = $($_.Identity|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyIdentity = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Name [TBA]"
                   "Value"                    = $TeamsMessagingPolicyIdentity
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.Description)) -EQ $FALSE) {
                $TeamsMessagingPolicyDescription = $($_.Description|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyDescription = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Description"
                   "Value"                    = $TeamsMessagingPolicyDescription
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowUrlPreviews)) -EQ $FALSE) {
                $TeamsMessagingPolicyAllowUrlPreviews = $($_.AllowUrlPreviews|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowUrlPreviews = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow URL previews"
                   "Value"                    = $TeamsMessagingPolicyAllowUrlPreviews
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowOwnerDeleteMessage)) -eq $False) {
                $TeamsMessagingPolicyAllowOwnerDeleteMessage = $($_.AllowOwnerDeleteMessage|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowOwnerDeleteMessage = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Owner Delete Message"
                   "Value"                    = $TeamsMessagingPolicyAllowOwnerDeleteMessage
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowUserEditMessage)) -eq $False) {
                $TeamsMessagingPolicyAllowUserEditMessage = $($_.AllowUserEditMessage|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowUserEditMessage = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow User Edit Message"
                   "Value"                    = $TeamsMessagingPolicyAllowUserEditMessage
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowUserDeleteMessage)) -eq $False) {
                $TeamsMessagingPolicyAllowUserDeleteMessage = $($_.AllowUserDeleteMessage|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowUserDeleteMessage = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow User Delete Message"
                   "Value"                    = $TeamsMessagingPolicyAllowUserDeleteMessage
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowUserChat)) -eq $False) {
                $TeamsMessagingPolicyAllowUserChat = $($_.AllowUserChat|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowUserChat = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow User Chat"
                   "Value"                    = $TeamsMessagingPolicyAllowUserChat
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowRemoveUser)) -eq $False) {
                $TeamsMessagingPolicyAllowRemoveUser = $($_.AllowRemoveUser|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowRemoveUser = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Remove User"
                   "Value"                    = $TeamsMessagingPolicyAllowRemoveUser
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowGiphy)) -eq $False) {
                $TeamsMessagingPolicyAllowGiphy = $($_.AllowGiphy|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowGiphy = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Giphy"
                   "Value"                    = $TeamsMessagingPolicyAllowGiphy
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.GiphyRatingType)) -eq $False) {
                $TeamsMessagingPolicyGiphyRatingType = $($_.GiphyRatingType|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyGiphyRatingType = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Giphy Rating Type"
                   "Value"                    = $TeamsMessagingPolicyGiphyRatingType
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowMemes)) -eq $False) {
                $TeamsMessagingPolicyAllowMemes = $($_.AllowMemes|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowMemes = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Memes"
                   "Value"                    = $TeamsMessagingPolicyAllowMemes
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowImmersiveReader)) -eq $False) {
                $TeamsMessagingPolicyAllowImmersiveReader = $($_.AllowImmersiveReader|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowImmersiveReader = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Immersive Reader"
                   "Value"                    = $TeamsMessagingPolicyAllowImmersiveReader
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowStickers)) -eq $False) {
                $TeamsMessagingPolicyAllowStickers = $($_.AllowStickers|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowStickers = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Stickers"
                   "Value"                    = $TeamsMessagingPolicyAllowStickers
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowUserTranslation)) -eq $False) {
                $TeamsMessagingPolicyAllowUserTranslation = $($_.AllowUserTranslation|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowUserTranslation = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow User Translation"
                   "Value"                    = $TeamsMessagingPolicyAllowUserTranslation
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.ReadReceiptsEnabledType)) -eq $False) {
                $TeamsMessagingPolicyReadReceiptsEnabledType = $($_.ReadReceiptsEnabledType|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyIdentity = "Not Configured"
            }
                $TeamsMessagingPolicyReadReceiptsEnabledType = [ordered]@{
                   "Configuration Item"       = "Read Recipts Enabled Type"
                   "Value"                    = $TeamsMessagingPolicyReadReceiptsEnabledType
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowPriorityMessages)) -eq $False) {
                $TeamsMessagingPolicyAllowPriorityMessages = $($_.AllowPriorityMessages|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAllowPriorityMessages = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Allow Priority Messages"
                   "Value"                    = $TeamsMessagingPolicyAllowPriorityMessages
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject 
    
            If (([string]::isnullorempty($_.ChannelsInChatListEnabledType)) -eq $False) {
                $TeamsMessagingPolicyChannelsInChatListEnabledType = $($_.ChannelsInChatListEnabledType|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyChannelsInChatListEnabledType = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Channels in Chat List Enabled Type"
                   "Value"                    = $TeamsMessagingPolicyChannelsInChatListEnabledType
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
            
            If (([string]::isnullorempty($_.AudioMessageEnabledType)) -eq $False) {
                $TeamsMessagingPolicyAudioMessageEnabledType = $($_.AudioMessageEnabledType|out-string).trim()
            }
            else {
                $TeamsMessagingPolicyAudioMessageEnabledType = "Not Configured"
            }
                $TeamsTeamsMessagingPolicydetail = [ordered]@{
                   "Configuration Item"       = "Audio Message Enabled Type"
                   "Value"                    = $TeamsMessagingPolicyAudioMessageEnabledType
               }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsTeamsMessagingPolicydetail
            $TeamsMessagingPolicyArray += $TeamsConfigurationObject
        }
    }
}

#############################################
#Teams - Meeting Broadcast Policy
#############################################
Write-Host " - Meeting Broadcast Policy" -foregroundcolor Gray
$TeamsMeetingBroadcastPolicy = Get-CsTeamsMeetingBroadcastPolicy

If ($null -eq  $TeamsMeetingBroadcastPolicy) {
     $TeamsMeetingBroadcastPolicydetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
    $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
}
Else {
    If($TeamsMeetingBroadcastPolicy -isnot [array]) {
        If (([string]::isnullorempty($TeamsMeetingBroadcastPolicy.Identity)) -eq $False) {
            $TeamsMeetingBroadcastPolicyIdentity = $($TeamsMeetingBroadcastPolicy.Identity|out-string).trim()
        }
        else {
            $TeamsMeetingBroadcastPolicyIdentity = "Not Configured"
        }
        $TeamsMeetingBroadcastPolicydetail = [ordered]@{
           "Configuration Item"       = "Name [TBA]"
           "Value"                    = $TeamsMeetingBroadcastPolicyIdentity
       }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
        $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingBroadcastPolicy.Description)) -eq $False) {
            $TeamsMeetingBroadcastPolicyDescription = $($TeamsMeetingBroadcastPolicy.Description|out-string).trim()
        }
        else {
            $TeamsMeetingBroadcastPolicyDescription = "Not Configured"
        }
        $TeamsMeetingBroadcastPolicydetail = [ordered]@{
           "Configuration Item"       = "Description"
           "Value"                    = $TeamsMeetingBroadcastPolicyDescription
        }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
        $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingBroadcastPolicy.AllowBroadcastScheduling)) -eq $False) {
            $TeamsMeetingBroadcastPolicyAllowBroadcastScheduling = $($TeamsMeetingBroadcastPolicy.AllowBroadcastScheduling|out-string).trim()
        }
        else {
            $TeamsMeetingBroadcastPolicyAllowBroadcastScheduling = "Not Configured"
        }
        $TeamsMeetingBroadcastPolicydetail = [ordered]@{
           "Configuration Item"       = "Allow Broadcast Scheduling"
           "Value"                    = $TeamsMeetingBroadcastPolicyAllowBroadcastScheduling
       }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
        $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingBroadcastPolicy.AllowBroadcastTranscription)) -eq $False) {
            $TeamsMeetingBroadcastPolicyAllowBroadcastTranscription = $($TeamsMeetingBroadcastPolicy.AllowBroadcastTranscription|out-string).trim()
        }
        else {
            $TeamsMeetingBroadcastPolicyAllowBroadcastTranscription = "Not Configured"
        }
        $TeamsMeetingBroadcastPolicydetail = [ordered]@{
           "Configuration Item"       = "Allow Broadcast Transcription"
           "Value"                    = $TeamsMeetingBroadcastPolicyAllowBroadcastTranscription
       }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
        $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingBroadcastPolicy.BroadcastAttendeeVisibilityMode)) -eq $False) {
            $TeamsMeetingBroadcastPolicyBroadcastAttendeeVisibilityMode = $($TeamsMeetingBroadcastPolicy.BroadcastAttendeeVisibilityMode|out-string).trim()
        }
        else {
            $TeamsMeetingBroadcastPolicyBroadcastAttendeeVisibilityMode = "Not Configured"
        }
        $TeamsMeetingBroadcastPolicydetail = [ordered]@{
           "Configuration Item"       = "Broadcast Attendee Visibility Mode"
           "Value"                    = $TeamsMeetingBroadcastPolicyBroadcastAttendeeVisibilityMode
       }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
        $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject

        If (([string]::isnullorempty($TeamsMeetingBroadcastPolicy.BroadcastRecordingMode)) -eq $False) {
            $TeamsMeetingBroadcastPolicyBroadcastRecordingMode = $($TeamsMeetingBroadcastPolicy.BroadcastRecordingMode|out-string).trim()
        }
        else {
            $TeamsMeetingBroadcastPolicyBroadcastRecordingMode = "Not Configured"
        }
        $TeamsMeetingBroadcastPolicydetail = [ordered]@{
           "Configuration Item"       = "Broadcast Record Mode"
           "Value"                    = $TeamsMeetingBroadcastPolicyBroadcastRecordingMode
       }
        $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
        $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
    }
    else {
        $TeamsMeetingBroadcastPolicy |foreach-object {
            If (([string]::isnullorempty($_.Identity)) -eq $False) {
                $TeamsMeetingBroadcastPolicyIdentity = $($_.Identity|out-string).trim()
            }
            else {
                $TeamsMeetingBroadcastPolicyIdentity = "Not Configured"
            }
            $TeamsMeetingBroadcastPolicydetail = [ordered]@{
                "Configuration Item"       = "Name [TBA]"
                "Value"                    = $TeamsMeetingBroadcastPolicyIdentity
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
            $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.Description)) -eq $False) {
                $TeamsMeetingBroadcastPolicyDescription = $($_.Description|out-string).trim()
            }
            else {
                $TeamsMeetingBroadcastPolicyDescription = "Not Configured"
            }
            $TeamsMeetingBroadcastPolicydetail = [ordered]@{
                "Configuration Item"       = "Description"
                "Value"                    = $TeamsMeetingBroadcastPolicyDescription
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
            $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowBroadcastScheduling)) -eq $False) {
                $TeamsMeetingBroadcastPolicyAllowBroadcastScheduling = $($_.AllowBroadcastScheduling|out-string).trim()
            }
            else {
                $TeamsMeetingBroadcastPolicyAllowBroadcastScheduling = "Not Configured"
            }
            $TeamsMeetingBroadcastPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow Broadcast Scheduling"
                "Value"                    = $TeamsMeetingBroadcastPolicyAllowBroadcastScheduling
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
            $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.AllowBroadcastTranscription)) -eq $False) {
                $TeamsMeetingBroadcastPolicyAllowBroadcastTranscription = $($_.AllowBroadcastTranscription|out-string).trim()
            }
            else {
                $TeamsMeetingBroadcastPolicyAllowBroadcastTranscription = "Not Configured"
            }
            $TeamsMeetingBroadcastPolicydetail = [ordered]@{
                "Configuration Item"       = "Allow Broadcast Transcription"
                "Value"                    = $TeamsMeetingBroadcastPolicyAllowBroadcastTranscription
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
            $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.BroadcastAttendeeVisibilityMode)) -eq $False) {
                $TeamsMeetingBroadcastPolicyBroadcastAttendeeVisibilityMode = $($_.BroadcastAttendeeVisibilityMode|out-string).trim()
            }
            else {
                $TeamsMeetingBroadcastPolicyBroadcastAttendeeVisibilityMode = "Not Configured"
            }
            $TeamsMeetingBroadcastPolicydetail = [ordered]@{
                "Configuration Item"       = "Broadcast Attendee Visibility Mode"
                "Value"                    = $TeamsMeetingBroadcastPolicyBroadcastAttendeeVisibilityMode
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
            $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
    
            If (([string]::isnullorempty($_.BroadcastRecordingMode)) -eq $False) {
                $TeamsMeetingBroadcastPolicyBroadcastRecordingMode = $($_.BroadcastRecordingMode|out-string).trim()
            }
            else {
                $TeamsMeetingBroadcastPolicyBroadcastRecordingMode = "Not Configured"
            }
            $TeamsMeetingBroadcastPolicydetail = [ordered]@{
                "Configuration Item"       = "Broadcast Record Mode"
                "Value"                    = $TeamsMeetingBroadcastPolicyBroadcastRecordingMode
            }
            $TeamsConfigurationObject = New-Object -TypeName psobject -Property $TeamsMeetingBroadcastPolicydetail
            $TeamsMeetingBroadcastPolicyArray += $TeamsConfigurationObject
        }
    }
}

#Get-CsTeamsMessagingPolicy
######################################################################################################################################################################################################################################################################################################
#############################################
#Security and Compliance
#############################################
Write-Host "Connecting to Office 365 Security and Compliance Center" -foregroundcolor Yellow
Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
Connect-IPPSSession -UserPrincipalName $UserPrincipalName
Write-Host "Querying Office 365 Security and Compliance configuration..." -foregroundcolor Yellow

#############################################
##Security and Compliance - Retention Labels
#############################################
Write-Host " - Retention Labels" -foregroundcolor Gray
$RetentionLabel = Get-ComplianceTag

If ($null -eq  $RetentionLabel) {
     $RetentionLabelDetail = [ordered]@{
        "Configuration Item"       = "Not Configured"
        "Value"                    = "N/A"
    }
    $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
    $RetentionLabelArray += $RetentionlabelConfigurationObject
}
Else {
    If($RetentionLabel -isnot [array]) {
        If (([string]::isnullorempty($RetentionLabel.name)) -eq $False) {
            $Retentionlabelname = $($RetentionLabel.name|out-string).trim()
        }
        else {
            $Retentionlabelname = "Not Configured"
        }
        $RetentionLabelDetail = [ordered]@{
           "Configuration Item"       = "Name [TBA]"
           "Value"                    = $Retentionlabelname
        }
        $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
        $RetentionLabelArray += $RetentionlabelConfigurationObject
        
        If (([string]::isnullorempty($RetentionLabel.Disabled)) -eq $False) {
            $RetentionlabelDisabled = $($RetentionLabel.Disabled|out-string).trim()
        }
        else {
            $RetentionlabelDisabled = "Not Configured"
        }
        $RetentionLabelDetail = [ordered]@{
           "Configuration Item"       = "Disabled"
           "Value"                    = $RetentionlabelDisabled
        }
        $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
        $RetentionLabelArray += $RetentionlabelConfigurationObject
        
        If (([string]::isnullorempty($RetentionLabel.mode)) -eq $False) {
            $Retentionlabelmode = $($RetentionLabel.mode|out-string).trim()
        }
        else {
            $Retentionlabelmode = "Not Configured"
        }
        $RetentionLabelDetail = [ordered]@{
           "Configuration Item"       = "Mode"
           "Value"                    = $Retentionlabelmode
        }
        $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
        $RetentionLabelArray += $RetentionlabelConfigurationObject

        If (([string]::isnullorempty($RetentionLabel.Workload)) -eq $False) {
            $RetentionlabelWorkload = $($RetentionLabel.Workload|out-string).trim()
        }
        else {
            $RetentionlabelWorkload = "Not Configured"
        }
        $RetentionLabelDetail = [ordered]@{
           "Configuration Item"       = "Workload"
           "Value"                    = $RetentionlabelWorkload
        }
        $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
        $RetentionLabelArray += $RetentionlabelConfigurationObject

        If (([string]::isnullorempty($RetentionLabel.RetentionAction)) -eq $False) {
            $RetentionlabelRetentionAction = $($RetentionLabel.RetentionAction|out-string).trim()
        }
        else {
            $RetentionlabelRetentionAction = "Not Configured"
        }
        $RetentionLabelDetail = [ordered]@{
           "Configuration Item"       = "Retention Action"
           "Value"                    = $RetentionlabelRetentionAction
        }
        $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
        $RetentionLabelArray += $RetentionlabelConfigurationObject

        If (([string]::isnullorempty($RetentionLabel.RetentionType)) -eq $False) {
            $RetentionlabelRetentionType = $($RetentionLabel.RetentionType|out-string).trim()
        }
        else {
            $RetentionlabelRetentionType = "Not Configured"
        }
        $RetentionLabelDetail = [ordered]@{
           "Configuration Item"       = "Retention Type"
           "Value"                    = $RetentionlabelRetentionType
        }
        $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
        $RetentionLabelArray += $RetentionlabelConfigurationObject

        If (([string]::isnullorempty($RetentionLabel.RetentionDuration)) -eq $False) {
            $RetentionlabelRetentionDuration = $($RetentionLabel.RetentionDuration|out-string).trim()
        }
        else {
            $RetentionlabelRetentionDuration = "Not Configured"
        }
        $RetentionLabelDetail = [ordered]@{
           "Configuration Item"       = "Retention Duration"
           "Value"                    = $RetentionlabelRetentionDuration
        }
        $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
        $RetentionLabelArray += $RetentionlabelConfigurationObject
    }
    else {
        $retentionlabel |foreach-object {
            If (([string]::isnullorempty($_.name)) -eq $False) {
                $Retentionlabelname = $($_.name|out-string).trim()
            }
            else {
                $Retentionlabelname = "Not Configured"
            }
            $RetentionLabelDetail = [ordered]@{
               "Configuration Item"       = "Name [TBA]"
               "Value"                    = $Retentionlabelname
            }
            $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
            $RetentionLabelArray += $RetentionlabelConfigurationObject
            
            If (([string]::isnullorempty($_.Disabled)) -eq $False) {
                $RetentionlabelDisabled = $($_.Disabled|out-string).trim()
            }
            else {
                $RetentionlabelDisabled = "Not Configured"
            }
            $RetentionLabelDetail = [ordered]@{
               "Configuration Item"       = "Disabled"
               "Value"                    = $RetentionlabelDisabled
            }
            $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
            $RetentionLabelArray += $RetentionlabelConfigurationObject
            
            If (([string]::isnullorempty($_.mode)) -eq $False) {
                $Retentionlabelmode = $($_.mode|out-string).trim()
            }
            else {
                $Retentionlabelmode = "Not Configured"
            }
            $RetentionLabelDetail = [ordered]@{
               "Configuration Item"       = "Mode"
               "Value"                    = $Retentionlabelmode
            }
            $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
            $RetentionLabelArray += $RetentionlabelConfigurationObject
    
            If (([string]::isnullorempty($_.Workload)) -eq $False) {
                $RetentionlabelWorkload = $($_.Workload|out-string).trim()
            }
            else {
                $RetentionlabelWorkload = "Not Configured"
            }
            $RetentionLabelDetail = [ordered]@{
               "Configuration Item"       = "Workload"
               "Value"                    = $RetentionlabelWorkload
            }
            $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
            $RetentionLabelArray += $RetentionlabelConfigurationObject
    
            If (([string]::isnullorempty($_.RetentionAction)) -eq $False) {
                $RetentionlabelRetentionAction = $($_.RetentionAction|out-string).trim()
            }
            else {
                $RetentionlabelRetentionAction = "Not Configured"
            }
            $RetentionLabelDetail = [ordered]@{
               "Configuration Item"       = "Retention Action"
               "Value"                    = $RetentionlabelRetentionAction
            }
            $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
            $RetentionLabelArray += $RetentionlabelConfigurationObject
    
            If (([string]::isnullorempty($_.RetentionType)) -eq $False) {
                $RetentionlabelRetentionType = $($_.RetentionType|out-string).trim()
            }
            else {
                $RetentionlabelRetentionType = "Not Configured"
            }
            $RetentionLabelDetail = [ordered]@{
               "Configuration Item"       = "Retention Type"
               "Value"                    = $RetentionlabelRetentionType
            }
            $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
            $RetentionLabelArray += $RetentionlabelConfigurationObject
    
            If (([string]::isnullorempty($_.RetentionDuration)) -eq $False) {
                $RetentionlabelRetentionDuration = $($_.RetentionDuration|out-string).trim()
            }
            else {
                $RetentionlabelRetentionDuration = "Not Configured"
            }
            $RetentionLabelDetail = [ordered]@{
               "Configuration Item"       = "Retention Duration"
               "Value"                    = $RetentionlabelRetentionDuration
            }
            $RetentionlabelConfigurationObject = New-Object -TypeName psobject -Property $RetentionLabelDetail
            $RetentionLabelArray += $RetentionlabelConfigurationObject
        }

    }
}

#############################################
##Security and Compliance - Retention Policy
#############################################
Write-Host " - Retention Policy" -foregroundcolor Gray
$RetentionPolicy = Get-RetentionCompliancePolicy

If ($null -eq  $RetentionPolicy) {
    $RetentionPolicyDetail = [ordered]@{
       "Configuration Item"       = "Not Configured"
       "Value"                    = "N/A"
   }
   $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
   $RetentionPolicyArray += $RetentionPolicyConfigurationObject
}
Else {
   If($RetentionPolicy -isnot [array]) {
        If (([string]::isnullorempty($RetentionPolicy.name)) -eq $False) {
           $RetentionPolicyname = $($RetentionPolicy.name|out-string).trim()
        }
        else {
           $RetentionPolicyname = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
          "Configuration Item"       = "Name [TBA]"
          "Value"                    = $RetentionPolicyname
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.enabled)) -eq $False) {
            $RetentionPolicyEnabled = $($RetentionPolicy.Enabled|out-string).trim()
        }
        else {
            $RetentionPolicyEnabled = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Enabled"
        "Value"                    = $RetentionPolicyEnabled
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.Mode)) -eq $False) {
            $RetentionPolicymode = $($RetentionPolicy.mode|out-string).trim()
        }
        else {
            $RetentionPolicymode = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Mode"
        "Value"                    = $RetentionPolicymode
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.Workload)) -eq $False) {
            $RetentionPolicyWorkload = $($RetentionPolicy.Workload|out-string).trim()
        }
        else {
            $RetentionPolicyWorkload = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Workload"
        "Value"                    = $RetentionPolicyWorkload
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.type)) -eq $False) {
            $RetentionPolicytype = $($RetentionPolicy.type|out-string).trim()
        }
        else {
            $RetentionPolicytype = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Type"
        "Value"                    = $RetentionPolicytype
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.TeamsPolicy)) -eq $False) {
            $RetentionPolicyTeamsPolicy = $($RetentionPolicy.TeamsPolicy|out-string).trim()
        }
        else {
            $RetentionPolicyTeamsPolicy = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Teams Policy"
        "Value"                    = $RetentionPolicyTeamsPolicy
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.SharePointLocation)) -eq $False) {
            $RetentionPolicySharePointLocation = $($RetentionPolicy.SharePointLocation|out-string).trim()
        }
        else {
            $RetentionPolicySharePointLocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "SharePoint Location"
        "Value"                    = $RetentionPolicySharePointLocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.SharePointLocationException)) -eq $False) {
            $RetentionPolicySharePointLocationException = $($RetentionPolicy.SharePointLocationException|out-string).trim()
        }
        else {
            $RetentionPolicySharePointLocationException = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "SharePoint Location Exception"
        "Value"                    = $RetentionPolicySharePointLocationException
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.Retentionruletypes)) -eq $False) {
            $RetentionPolicyRetentionruletypes = $($RetentionPolicy.Retentionruletypes|out-string).trim()
        }
        else {
            $RetentionPolicyRetentionruletypes = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Retention Rule Types"
        "Value"                    = $RetentionPolicyRetentionruletypes
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.ExchangeLocation)) -eq $False) {
            $RetentionPolicyExchangeLocation = $($RetentionPolicy.ExchangeLocation|out-string).trim()
        }
        else {
            $RetentionPolicyExchangeLocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Exchange Location"
        "Value"                    = $RetentionPolicyExchangeLocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.ExchangeLocationException)) -eq $False) {
            $RetentionPolicyExchangeLocationException = $($RetentionPolicy.ExchangeLocationException|out-string).trim()
        }
        else {
            $RetentionPolicyExchangeLocationException = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Exchange Location Exception"
        "Value"                    = $RetentionPolicyExchangeLocationException
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.PublicFolderlocation)) -eq $False) {
            $RetentionPolicyPublicFolderlocation = $($RetentionPolicy.PublicFolderlocation|out-string).trim()
        }
        else {
            $RetentionPolicyPublicFolderlocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Public Folder Location"
        "Value"                    = $RetentionPolicyPublicFolderlocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.SkypeLocation)) -eq $False) {
            $RetentionPolicySkypeLocation = $($RetentionPolicy.SkypeLocation|out-string).trim()
        }
        else {
            $RetentionPolicySkypeLocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Skype Location"
        "Value"                    = $RetentionPolicySkypeLocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.SkypeLocationexception)) -eq $False) {
            $RetentionPolicySkypeLocationexception = $($RetentionPolicy.SkypeLocationexception|out-string).trim()
        }
        else {
            $RetentionPolicySkypeLocationexception = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Skype Location Exception"
        "Value"                    = $RetentionPolicySkypeLocationexception
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.moderngrouplocation)) -eq $False) {
            $RetentionPolicymoderngrouplocation = $($RetentionPolicy.moderngrouplocation|out-string).trim()
        }
        else {
            $RetentionPolicymoderngrouplocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Modern Group Location"
        "Value"                    = $RetentionPolicymoderngrouplocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.moderngrouplocationexception)) -eq $False) {
            $RetentionPolicymoderngrouplocationexception = $($RetentionPolicy.moderngrouplocationexception|out-string).trim()
        }
        else {
            $RetentionPolicymoderngrouplocationexception = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Modern Group Location Exception"
        "Value"                    = $RetentionPolicymoderngrouplocationexception
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.onedrivelocation)) -eq $False) {
            $RetentionPolicyonedrivelocation = $($RetentionPolicy.onedrivelocation|out-string).trim()
        }
        else {
            $RetentionPolicyonedrivelocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "OneDrive Location"
        "Value"                    = $RetentionPolicyonedrivelocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.onedrivelocationexception)) -eq $False) {
            $RetentionPolicyonedrivelocationexception = $($RetentionPolicy.onedrivelocationexception|out-string).trim()
        }
        else {
            $RetentionPolicyonedrivelocationexception = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "OneDrive Location Exception"
        "Value"                    = $RetentionPolicyonedrivelocationexception
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.teamschatlocation)) -eq $False) {
            $RetentionPolicyteamschatlocation = $($RetentionPolicy.teamschatlocation|out-string).trim()
        }
        else {
            $RetentionPolicyteamschatlocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Teams Chat Location"
        "Value"                    = $RetentionPolicyteamschatlocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.teamschatlocationexception)) -eq $False) {
            $RetentionPolicyteamschatlocationexception = $($RetentionPolicy.teamschatlocationexception|out-string).trim()
        }
        else {
            $RetentionPolicyteamschatlocationexception = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Teams Chat Location Exception"
        "Value"                    = $RetentionPolicyteamschatlocationexception
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.teamschannellocation)) -eq $False) {
            $RetentionPolicyteamschannellocation = $($RetentionPolicy.teamschannellocation|out-string).trim()
        }
        else {
            $RetentionPolicyteamschannellocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Teams Channel Location"
        "Value"                    = $RetentionPolicyteamschannellocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.teamschannellocationexception)) -eq $False) {
            $RetentionPolicyteamschannellocationexception = $($RetentionPolicy.teamschannellocationexception|out-string).trim()
        }
        else {
            $RetentionPolicyteamschannellocationexception = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Teams Channel Location Exception"
        "Value"                    = $RetentionPolicyteamschannellocationexception
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

        If (([string]::isnullorempty($RetentionPolicy.dynamicscopelocation)) -eq $False) {
            $RetentionPolicydynamicscopelocation = $($RetentionPolicy.dynamicscopelocation|out-string).trim()
        }
        else {
            $RetentionPolicydynamicscopelocation = "Not Configured"
        }
        $RetentionPolicyDetail = [ordered]@{
        "Configuration Item"       = "Dynamic Scope Location"
        "Value"                    = $RetentionPolicydynamicscopelocation
        }
        $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
        $RetentionPolicyArray += $RetentionPolicyConfigurationObject

   }
   else {
       $RetentionPolicy |ForEach-Object {
            If (([string]::isnullorempty($_.name)) -eq $False) {
                $RetentionPolicyname = $($_.name|out-string).trim()
            }
            else {
                $RetentionPolicyname = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Name [TBA]"
            "Value"                    = $RetentionPolicyname
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.enabled)) -eq $False) {
                $RetentionPolicyEnabled = $($_.Enabled|out-string).trim()
            }
            else {
                $RetentionPolicyEnabled = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Enabled"
            "Value"                    = $RetentionPolicyEnabled
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.Mode)) -eq $False) {
                $RetentionPolicymode = $($_.mode|out-string).trim()
            }
            else {
                $RetentionPolicymode = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Mode"
            "Value"                    = $RetentionPolicymode
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.Workload)) -eq $False) {
                $RetentionPolicyWorkload = $($_.Workload|out-string).trim()
            }
            else {
                $RetentionPolicyWorkload = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Workload"
            "Value"                    = $RetentionPolicyWorkload
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.type)) -eq $False) {
                $RetentionPolicytype = $($_.type|out-string).trim()
            }
            else {
                $RetentionPolicytype = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Type"
            "Value"                    = $RetentionPolicytype
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.TeamsPolicy)) -eq $False) {
                $RetentionPolicyTeamsPolicy = $($_.TeamsPolicy|out-string).trim()
            }
            else {
                $RetentionPolicyTeamsPolicy = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Teams Policy"
            "Value"                    = $RetentionPolicyTeamsPolicy
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.SharePointLocation)) -eq $False) {
                $RetentionPolicySharePointLocation = $($_.SharePointLocation|out-string).trim()
            }
            else {
                $RetentionPolicySharePointLocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "SharePoint Location"
            "Value"                    = $RetentionPolicySharePointLocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.SharePointLocationException)) -eq $False) {
                $RetentionPolicySharePointLocationException = $($_.SharePointLocationException|out-string).trim()
            }
            else {
                $RetentionPolicySharePointLocationException = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "SharePoint Location Exception"
            "Value"                    = $RetentionPolicySharePointLocationException
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.Retentionruletypes)) -eq $False) {
                $RetentionPolicyRetentionruletypes = $($_.Retentionruletypes|out-string).trim()
            }
            else {
                $RetentionPolicyRetentionruletypes = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Retention Rule Types"
            "Value"                    = $RetentionPolicyRetentionruletypes
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.ExchangeLocation)) -eq $False) {
                $RetentionPolicyExchangeLocation = $($_.ExchangeLocation|out-string).trim()
            }
            else {
                $RetentionPolicyExchangeLocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Exchange Location"
            "Value"                    = $RetentionPolicyExchangeLocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.ExchangeLocationException)) -eq $False) {
                $RetentionPolicyExchangeLocationException = $($_.ExchangeLocationException|out-string).trim()
            }
            else {
                $RetentionPolicyExchangeLocationException = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Exchange Location Exception"
            "Value"                    = $RetentionPolicyExchangeLocationException
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.PublicFolderlocation)) -eq $False) {
                $RetentionPolicyPublicFolderlocation = $($_.PublicFolderlocation|out-string).trim()
            }
            else {
                $RetentionPolicyPublicFolderlocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Public Folder Location"
            "Value"                    = $RetentionPolicyPublicFolderlocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.SkypeLocation)) -eq $False) {
                $RetentionPolicySkypeLocation = $($_.SkypeLocation|out-string).trim()
            }
            else {
                $RetentionPolicySkypeLocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Skype Location"
            "Value"                    = $RetentionPolicySkypeLocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.SkypeLocationexception)) -eq $False) {
                $RetentionPolicySkypeLocationexception = $($_.SkypeLocationexception|out-string).trim()
            }
            else {
                $RetentionPolicySkypeLocationexception = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Skype Location Exception"
            "Value"                    = $RetentionPolicySkypeLocationexception
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.moderngrouplocation)) -eq $False) {
                $RetentionPolicymoderngrouplocation = $($_.moderngrouplocation|out-string).trim()
            }
            else {
                $RetentionPolicymoderngrouplocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Modern Group Location"
            "Value"                    = $RetentionPolicymoderngrouplocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.moderngrouplocationexception)) -eq $False) {
                $RetentionPolicymoderngrouplocationexception = $($_.moderngrouplocationexception|out-string).trim()
            }
            else {
                $RetentionPolicymoderngrouplocationexception = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Modern Group Location Exception"
            "Value"                    = $RetentionPolicymoderngrouplocationexception
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.onedrivelocation)) -eq $False) {
                $RetentionPolicyonedrivelocation = $($_.onedrivelocation|out-string).trim()
            }
            else {
                $RetentionPolicyonedrivelocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "OneDrive Location"
            "Value"                    = $RetentionPolicyonedrivelocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.onedrivelocationexception)) -eq $False) {
                $RetentionPolicyonedrivelocationexception = $($_.onedrivelocationexception|out-string).trim()
            }
            else {
                $RetentionPolicyonedrivelocationexception = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "OneDrive Location Exception"
            "Value"                    = $RetentionPolicyonedrivelocationexception
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.teamschatlocation)) -eq $False) {
                $RetentionPolicyteamschatlocation = $($_.teamschatlocation|out-string).trim()
            }
            else {
                $RetentionPolicyteamschatlocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Teams Chat Location"
            "Value"                    = $RetentionPolicyteamschatlocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.teamschatlocationexception)) -eq $False) {
                $RetentionPolicyteamschatlocationexception = $($_.teamschatlocationexception|out-string).trim()
            }
            else {
                $RetentionPolicyteamschatlocationexception = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Teams Chat Location Exception"
            "Value"                    = $RetentionPolicyteamschatlocationexception
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.teamschannellocation)) -eq $False) {
                $RetentionPolicyteamschannellocation = $($_.teamschannellocation|out-string).trim()
            }
            else {
                $RetentionPolicyteamschannellocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Teams Channel Location"
            "Value"                    = $RetentionPolicyteamschannellocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.teamschannellocationexception)) -eq $False) {
                $RetentionPolicyteamschannellocationexception = $($_.teamschannellocationexception|out-string).trim()
            }
            else {
                $RetentionPolicyteamschannellocationexception = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Teams Channel Location Exception"
            "Value"                    = $RetentionPolicyteamschannellocationexception
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
    
            If (([string]::isnullorempty($_.dynamicscopelocation)) -eq $False) {
                $RetentionPolicydynamicscopelocation = $($_.dynamicscopelocation|out-string).trim()
            }
            else {
                $RetentionPolicydynamicscopelocation = "Not Configured"
            }
            $RetentionPolicyDetail = [ordered]@{
            "Configuration Item"       = "Dynamic Scope Location"
            "Value"                    = $RetentionPolicydynamicscopelocation
            }
            $RetentionPolicyConfigurationObject = New-Object -TypeName psobject -Property $RetentionPolicyDetail
            $RetentionPolicyArray += $RetentionPolicyConfigurationObject
       }
   }
}


#############################################
##Security and Compliance - Sensitivity Labels
#############################################
Write-Host " - Sensitivity Labels" -foregroundcolor Gray
$SensitivityLabels = Get-Label

If ($null -eq  $SensitivityLabels) {
    $SensitivityLabelsDetail = [ordered]@{
       "Configuration Item"       = "Not Configured"
       "Value"                    = "N/A"
   }
   $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
   $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
}
Else {
   If($SensitivityLabels -isnot [array]) {
        If (([string]::isnullorempty($SensitivityLabels.name)) -eq $False) {
           $SensitivityLabelsname = $($SensitivityLabels.name|out-string).trim()
        }
        else {
           $SensitivityLabelsname = "Not Configured"
        }
        $SensitivityLabelsDetail = [ordered]@{
          "Configuration Item"       = "Name [TBA]"
          "Value"                    = $SensitivityLabelsname
        }
        $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
        $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject

        If (([string]::isnullorempty($SensitivityLabels.Workload)) -eq $False) {
            $SensitivityLabelsWorkload = $($SensitivityLabels.Workload|out-string).trim()
         }
         else {
            $SensitivityLabelsWorkload = "Not Configured"
         }
         $SensitivityLabelsDetail = [ordered]@{
           "Configuration Item"       = "Workload"
           "Value"                    = $SensitivityLabelsWorkload
         }
         $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
         $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject

         If (([string]::isnullorempty($SensitivityLabels.settings)) -eq $False) {
            $SensitivityLabelssettings = $($SensitivityLabels.settings|out-string).trim()
         }
         else {
            $SensitivityLabelssettings = "Not Configured"
         }
         $SensitivityLabelsDetail = [ordered]@{
           "Configuration Item"       = "Settings"
           "Value"                    = $SensitivityLabelssettings
         }
         $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
         $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject

         If (([string]::isnullorempty($SensitivityLabels.labelactions)) -eq $False) {
            $SensitivityLabelslabelactions = $($SensitivityLabels.labelactions|out-string).trim()
         }
         else {
            $SensitivityLabelslabelactions = "Not Configured"
         }
         $SensitivityLabelsDetail = [ordered]@{
           "Configuration Item"       = "Label Actions"
           "Value"                    = $SensitivityLabelslabelactions
         }
         $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
         $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject

         If (([string]::isnullorempty($SensitivityLabels.conditions)) -eq $False) {
            $SensitivityLabelsconditions = $($SensitivityLabels.conditions|out-string).trim()
         }
         else {
            $SensitivityLabelsconditions = "Not Configured"
         }
         $SensitivityLabelsDetail = [ordered]@{
           "Configuration Item"       = "Conditions"
           "Value"                    = $SensitivityLabelsconditions
         }
         $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
         $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject

         If (([string]::isnullorempty($SensitivityLabels.localsettings)) -eq $False) {
            $SensitivityLabelslocalsettings = $($SensitivityLabels.localsettings|out-string).trim()
         }
         else {
            $SensitivityLabelslocalsettings = "Not Configured"
         }
         $SensitivityLabelsDetail = [ordered]@{
           "Configuration Item"       = "Local Settings"
           "Value"                    = $SensitivityLabelslocalsettings
         }
         $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
         $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject

         If (([string]::isnullorempty($SensitivityLabels.tooltip)) -eq $False) {
            $SensitivityLabelstooltip = $($SensitivityLabels.tooltip|out-string).trim()
         }
         else {
            $SensitivityLabelstooltip = "Not Configured"
         }
         $SensitivityLabelsDetail = [ordered]@{
           "Configuration Item"       = "Tooltip"
           "Value"                    = $SensitivityLabelstooltip
         }
         $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
         $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
    }
    else {
        $SensitivityLabels |foreach-object {
            If (([string]::isnullorempty($_.name)) -eq $False) {
                $SensitivityLabelsname = $($_.name|out-string).trim()
             }
             else {
                $SensitivityLabelsname = "Not Configured"
             }
             $SensitivityLabelsDetail = [ordered]@{
               "Configuration Item"       = "Name [TBA]"
               "Value"                    = $SensitivityLabelsname
             }
             $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
             $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
     
             If (([string]::isnullorempty($_.Workload)) -eq $False) {
                 $SensitivityLabelsWorkload = $($_.Workload|out-string).trim()
              }
              else {
                 $SensitivityLabelsWorkload = "Not Configured"
              }
              $SensitivityLabelsDetail = [ordered]@{
                "Configuration Item"       = "Workload"
                "Value"                    = $SensitivityLabelsWorkload
              }
              $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
              $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
     
              If (([string]::isnullorempty($_.settings)) -eq $False) {
                 $SensitivityLabelssettings = $($_.settings|out-string).trim()
              }
              else {
                 $SensitivityLabelssettings = "Not Configured"
              }
              $SensitivityLabelsDetail = [ordered]@{
                "Configuration Item"       = "Settings"
                "Value"                    = $SensitivityLabelssettings
              }
              $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
              $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
     
              If (([string]::isnullorempty($_.labelactions)) -eq $False) {
                 $SensitivityLabelslabelactions = $($_.labelactions|out-string).trim()
              }
              else {
                 $SensitivityLabelslabelactions = "Not Configured"
              }
              $SensitivityLabelsDetail = [ordered]@{
                "Configuration Item"       = "Label Actions"
                "Value"                    = $SensitivityLabelslabelactions
              }
              $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
              $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
     
              If (([string]::isnullorempty($_.conditions)) -eq $False) {
                 $SensitivityLabelsconditions = $($_.conditions|out-string).trim()
              }
              else {
                 $SensitivityLabelsconditions = "Not Configured"
              }
              $SensitivityLabelsDetail = [ordered]@{
                "Configuration Item"       = "Conditions"
                "Value"                    = $SensitivityLabelsconditions
              }
              $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
              $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
     
              If (([string]::isnullorempty($_.localsettings)) -eq $False) {
                 $SensitivityLabelslocalsettings = $($_.localsettings|out-string).trim()
              }
              else {
                 $SensitivityLabelslocalsettings = "Not Configured"
              }
              $SensitivityLabelsDetail = [ordered]@{
                "Configuration Item"       = "Local Settings"
                "Value"                    = $SensitivityLabelslocalsettings
              }
              $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
              $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
     
              If (([string]::isnullorempty($_.tooltip)) -eq $False) {
                 $SensitivityLabelstooltip = $($_.tooltip|out-string).trim()
              }
              else {
                 $SensitivityLabelstooltip = "Not Configured"
              }
              $SensitivityLabelsDetail = [ordered]@{
                "Configuration Item"       = "Tooltip"
                "Value"                    = $SensitivityLabelstooltip
              }
              $SensitivityLabelsConfigurationObject = New-Object -TypeName psobject -Property $SensitivityLabelsDetail
              $SensitivityLabelsArray += $SensitivityLabelsConfigurationObject
        }
    }
}
        
#############################################
##Security and Compliance - Sensitivity label policy
#############################################
Write-Host " - Sensitivity label policy" -foregroundcolor Gray
$Sensitivitylabelpolicy = get-labelpolicy

If ($null -eq  $Sensitivitylabelpolicy) {
    $SensitivitylabelpolicyDetail = [ordered]@{
       "Configuration Item"       = "Not Configured"
       "Value"                    = "N/A"
   }
   $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
   $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
}
Else {
   If($Sensitivitylabelpolicy -isnot [array]) {
        If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.name)) -eq $False) {
           $Sensitivitylabelpolicyname = $($Sensitivitylabelpolicy.name|out-string).trim()
        }
        else {
           $Sensitivitylabelpolicyname = "Not Configured"
        }
        $SensitivitylabelpolicyDetail = [ordered]@{
          "Configuration Item"       = "Name [TBA]"
          "Value"                    = $Sensitivitylabelpolicyname
        }
        $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
        $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

        If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.mode)) -eq $False) {
            $Sensitivitylabelpolicymode = $($Sensitivitylabelpolicy.mode|out-string).trim()
         }
         else {
            $Sensitivitylabelpolicyname = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Mode"
           "Value"                    = $Sensitivitylabelpolicymode
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.enabled)) -eq $False) {
            $Sensitivitylabelpolicyenabled = $($Sensitivitylabelpolicy.enabled|out-string).trim()
         }
         else {
            $Sensitivitylabelpolicyenabled = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Enabled"
           "Value"                    = $Sensitivitylabelpolicyenabled
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.Workload)) -eq $False) {
            $SensitivitylabelpolicyWorkload = $($Sensitivitylabelpolicy.Workload|out-string).trim()
         }
         else {
            $SensitivitylabelpolicyWorkload = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Workload"
           "Value"                    = $SensitivitylabelpolicyWorkload
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.TYPE)) -eq $False) {
            $SensitivitylabelpolicyTYPE = $($Sensitivitylabelpolicy.TYPE|out-string).trim()
         }
         else {
            $SensitivitylabelpolicyTYPE = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Type"
           "Value"                    = $SensitivitylabelpolicyTYPE
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.Settings)) -eq $False) {
            $SensitivitylabelpolicySettings = $($Sensitivitylabelpolicy.Settings|out-string).trim()
         }
         else {
            $SensitivitylabelpolicySettings = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Settings"
           "Value"                    = $SensitivitylabelpolicySettings
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.Labels)) -eq $False) {
            $SensitivitylabelpolicyLabels = $($Sensitivitylabelpolicy.Labels|out-string).trim()
         }
         else {
            $SensitivitylabelpolicyLabels = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Labels"
           "Value"                    = $SensitivitylabelpolicyLabels
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.SharePointLocation)) -eq $False) {
            $SensitivitylabelpolicySharePointLocation = $($Sensitivitylabelpolicy.SharePointLocation|out-string).trim()
         }
         else {
            $SensitivitylabelpolicySharePointLocation = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "SharePoint Location"
           "Value"                    = $SensitivitylabelpolicySharePointLocation
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.SharePointLocationException)) -eq $False) {
            $SensitivitylabelpolicySharePointLocationException = $($Sensitivitylabelpolicy.SharePointLocationException|out-string).trim()
         }
         else {
            $SensitivitylabelpolicySharePointLocationException = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "SharePoint Location Exception"
           "Value"                    = $SensitivitylabelpolicySharePointLocationException
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.ExchangeLocation)) -eq $False) {
            $SensitivitylabelpolicyExchangeLocation = $($Sensitivitylabelpolicy.ExchangeLocation|out-string).trim()
         }
         else {
            $SensitivitylabelpolicyExchangeLocation = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange Location"
           "Value"                    = $SensitivitylabelpolicyExchangeLocation
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.ExchangeLocationException)) -eq $False) {
            $SensitivitylabelpolicyExchangeLocationException = $($Sensitivitylabelpolicy.ExchangeLocationException|out-string).trim()
         }
         else {
            $SensitivitylabelpolicyExchangeLocationException = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange Location Exception"
           "Value"                    = $SensitivitylabelpolicyExchangeLocationException
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.PublicFolderLocation)) -eq $False) {
            $SensitivitylabelpolicyPublicFolderLocation = $($Sensitivitylabelpolicy.PublicFolderLocation|out-string).trim()
         }
         else {
            $SensitivitylabelpolicyPublicFolderLocation = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Public Folder Location"
           "Value"                    = $SensitivitylabelpolicyPublicFolderLocation
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.SkypeLocation)) -eq $False) {
            $SensitivitylabelpolicySkypeLocation = $($Sensitivitylabelpolicy.SkypeLocation|out-string).trim()
         }
         else {
            $SensitivitylabelpolicySkypeLocation = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Skype Location"
           "Value"                    = $SensitivitylabelpolicySkypeLocation
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.SkypeLocationexception)) -eq $False) {
            $SensitivitylabelpolicySkypeLocationexception = $($Sensitivitylabelpolicy.SkypeLocationexception|out-string).trim()
         }
         else {
            $SensitivitylabelpolicySkypeLocationexception = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Skype Location Exception"
           "Value"                    = $SensitivitylabelpolicySkypeLocationexception
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.moderngrouplocation)) -eq $False) {
            $Sensitivitylabelpolicymoderngrouplocation = $($Sensitivitylabelpolicy.moderngrouplocation|out-string).trim()
         }
         else {
            $Sensitivitylabelpolicymoderngrouplocation = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Modern Group Location"
           "Value"                    = $Sensitivitylabelpolicymoderngrouplocation
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.moderngrouplocationexception)) -eq $False) {
            $Sensitivitylabelpolicymoderngrouplocationexception = $($Sensitivitylabelpolicy.moderngrouplocationexception|out-string).trim()
         }
         else {
            $Sensitivitylabelpolicymoderngrouplocationexception = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "Modern Group Location Exception"
           "Value"                    = $Sensitivitylabelpolicymoderngrouplocationexception
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.onedrivelocation)) -eq $False) {
            $Sensitivitylabelpolicyonedrivelocation = $($Sensitivitylabelpolicy.onedrivelocation|out-string).trim()
         }
         else {
            $Sensitivitylabelpolicyonedrivelocation = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "OneDrive Location"
           "Value"                    = $Sensitivitylabelpolicyonedrivelocation
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject

         If (([string]::IsNullOrEmpty($Sensitivitylabelpolicy.onedrivelocationexception)) -eq $False) {
            $Sensitivitylabelpolicyonedrivelocationexception = $($Sensitivitylabelpolicy.onedrivelocationexception|out-string).trim()
         }
         else {
            $Sensitivitylabelpolicyonedrivelocationexception = "Not Configured"
         }
         $SensitivitylabelpolicyDetail = [ordered]@{
           "Configuration Item"       = "OneDrive Location Exception"
           "Value"                    = $Sensitivitylabelpolicyonedrivelocationexception
         }
         $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
         $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
    }
    else {
        $Sensitivitylabelpolicy |foreach-object{
            If (([string]::IsNullOrEmpty($_.name)) -eq $False) {
                $Sensitivitylabelpolicyname = $($_.name|out-string).trim()
             }
             else {
                $Sensitivitylabelpolicyname = "Not Configured"
             }
             $SensitivitylabelpolicyDetail = [ordered]@{
               "Configuration Item"       = "Name [TBA]"
               "Value"                    = $Sensitivitylabelpolicyname
             }
             $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
             $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
             If (([string]::IsNullOrEmpty($_.mode)) -eq $False) {
                 $Sensitivitylabelpolicymode = $($_.mode|out-string).trim()
              }
              else {
                 $Sensitivitylabelpolicyname = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Mode"
                "Value"                    = $Sensitivitylabelpolicymode
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.enabled)) -eq $False) {
                 $Sensitivitylabelpolicyenabled = $($_.enabled|out-string).trim()
              }
              else {
                 $Sensitivitylabelpolicyenabled = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Enabled"
                "Value"                    = $Sensitivitylabelpolicyenabled
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.Workload)) -eq $False) {
                 $SensitivitylabelpolicyWorkload = $($_.Workload|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicyWorkload = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Workload"
                "Value"                    = $SensitivitylabelpolicyWorkload
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.TYPE)) -eq $False) {
                 $SensitivitylabelpolicyTYPE = $($_.TYPE|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicyTYPE = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Type"
                "Value"                    = $SensitivitylabelpolicyTYPE
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.Settings)) -eq $False) {
                 $SensitivitylabelpolicySettings = $($_.Settings|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicySettings = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Settings"
                "Value"                    = $SensitivitylabelpolicySettings
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.Labels)) -eq $False) {
                 $SensitivitylabelpolicyLabels = $($_.Labels|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicyLabels = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Labels"
                "Value"                    = $SensitivitylabelpolicyLabels
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.SharePointLocation)) -eq $False) {
                 $SensitivitylabelpolicySharePointLocation = $($_.SharePointLocation|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicySharePointLocation = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "SharePoint Location"
                "Value"                    = $SensitivitylabelpolicySharePointLocation
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.SharePointLocationException)) -eq $False) {
                 $SensitivitylabelpolicySharePointLocationException = $($_.SharePointLocationException|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicySharePointLocationException = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "SharePoint Location Exception"
                "Value"                    = $SensitivitylabelpolicySharePointLocationException
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.ExchangeLocation)) -eq $False) {
                 $SensitivitylabelpolicyExchangeLocation = $($_.ExchangeLocation|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicyExchangeLocation = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange Location"
                "Value"                    = $SensitivitylabelpolicyExchangeLocation
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.ExchangeLocationException)) -eq $False) {
                 $SensitivitylabelpolicyExchangeLocationException = $($_.ExchangeLocationException|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicyExchangeLocationException = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange Location Exception"
                "Value"                    = $SensitivitylabelpolicyExchangeLocationException
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.PublicFolderLocation)) -eq $False) {
                 $SensitivitylabelpolicyPublicFolderLocation = $($_.PublicFolderLocation|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicyPublicFolderLocation = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Public Folder Location"
                "Value"                    = $SensitivitylabelpolicyPublicFolderLocation
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.SkypeLocation)) -eq $False) {
                 $SensitivitylabelpolicySkypeLocation = $($_.SkypeLocation|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicySkypeLocation = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Skype Location"
                "Value"                    = $SensitivitylabelpolicySkypeLocation
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.SkypeLocationexception)) -eq $False) {
                 $SensitivitylabelpolicySkypeLocationexception = $($_.SkypeLocationexception|out-string).trim()
              }
              else {
                 $SensitivitylabelpolicySkypeLocationexception = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Skype Location Exception"
                "Value"                    = $SensitivitylabelpolicySkypeLocationexception
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.moderngrouplocation)) -eq $False) {
                 $Sensitivitylabelpolicymoderngrouplocation = $($_.moderngrouplocation|out-string).trim()
              }
              else {
                 $Sensitivitylabelpolicymoderngrouplocation = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Modern Group Location"
                "Value"                    = $Sensitivitylabelpolicymoderngrouplocation
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.moderngrouplocationexception)) -eq $False) {
                 $Sensitivitylabelpolicymoderngrouplocationexception = $($_.moderngrouplocationexception|out-string).trim()
              }
              else {
                 $Sensitivitylabelpolicymoderngrouplocationexception = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "Modern Group Location Exception"
                "Value"                    = $Sensitivitylabelpolicymoderngrouplocationexception
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.onedrivelocation)) -eq $False) {
                 $Sensitivitylabelpolicyonedrivelocation = $($_.onedrivelocation|out-string).trim()
              }
              else {
                 $Sensitivitylabelpolicyonedrivelocation = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "OneDrive Location"
                "Value"                    = $Sensitivitylabelpolicyonedrivelocation
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
     
              If (([string]::IsNullOrEmpty($_.onedrivelocationexception)) -eq $False) {
                 $Sensitivitylabelpolicyonedrivelocationexception = $($_.onedrivelocationexception|out-string).trim()
              }
              else {
                 $Sensitivitylabelpolicyonedrivelocationexception = "Not Configured"
              }
              $SensitivitylabelpolicyDetail = [ordered]@{
                "Configuration Item"       = "OneDrive Location Exception"
                "Value"                    = $Sensitivitylabelpolicyonedrivelocationexception
              }
              $SensitivitylabelpolicyConfigurationObject = New-Object -TypeName psobject -Property $SensitivitylabelpolicyDetail
              $SensitivitylabelpolicyArray += $SensitivitylabelpolicyConfigurationObject
        }
    }
}

        
#############################################
##Security and Compliance - Dlp Compliance Policy
#############################################
Write-Host " - Dlp Compliance Policy" -foregroundcolor Gray
$DlpCompliancePolicy = Get-DlpCompliancePolicy

If ($null -eq  $DlpCompliancePolicy) {
    $DlpCompliancePolicyDetail = [ordered]@{
       "Configuration Item"       = "Not Configured"
       "Value"                    = "N/A"
   }
   $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
   $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
}
Else {
   If($DlpCompliancePolicy -isnot [array]) {
        If (([string]::IsNullOrEmpty($DlpCompliancePolicy.name)) -eq $False) {
           $DlpCompliancePolicyname = $($DlpCompliancePolicy.name|out-string).trim()
        }
        else {
           $RetentionPolicyname = "Not Configured"
        }
        $DlpCompliancePolicyDetail = [ordered]@{
          "Configuration Item"       = "Name [TBA]"
          "Value"                    = $DlpCompliancePolicyname
        }
        $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
        $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

        If (([string]::IsNullOrEmpty($DlpCompliancePolicy.enabled)) -eq $False) {
            $DlpCompliancePolicyenabled = $($DlpCompliancePolicy.enabled|out-string).trim()
         }
         else {
            $DlpCompliancePolicyenabled = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Enabled"
           "Value"                    = $DlpCompliancePolicyenabled
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.workload)) -eq $False) {
            $DlpCompliancePolicyworkload = $($DlpCompliancePolicy.workload|out-string).trim()
         }
         else {
            $DlpCompliancePolicyworkload = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Workload"
           "Value"                    = $DlpCompliancePolicyworkload
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.mode)) -eq $False) {
            $DlpCompliancePolicymode = $($DlpCompliancePolicy.mode|out-string).trim()
         }
         else {
            $DlpCompliancePolicymode = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Mode"
           "Value"                    = $DlpCompliancePolicymode
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.type)) -eq $False) {
            $DlpCompliancePolicytype = $($DlpCompliancePolicy.type|out-string).trim()
         }
         else {
            $DlpCompliancePolicytype = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Type"
           "Value"                    = $DlpCompliancePolicytype
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.ExchangeLocation)) -eq $False) {
            $DlpCompliancePolicyExchangeLocation = $($DlpCompliancePolicy.ExchangeLocation|out-string).trim()
         }
         else {
            $DlpCompliancePolicyExchangeLocation = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange Location"
           "Value"                    = $DlpCompliancePolicyExchangeLocation
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.SharePointLocation)) -eq $False) {
            $DlpCompliancePolicySharePointLocation = $($DlpCompliancePolicy.SharePointLocation|out-string).trim()
         }
         else {
            $DlpCompliancePolicySharePointLocation = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "SharePoint Location"
           "Value"                    = $DlpCompliancePolicySharePointLocation
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.SharePointLocationException)) -eq $False) {
            $DlpCompliancePolicySharePointLocationException = $($DlpCompliancePolicy.SharePointLocationException|out-string).trim()
         }
         else {
            $DlpCompliancePolicySharePointLocationException = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "SharePoint Location Exception"
           "Value"                    = $DlpCompliancePolicySharePointLocationException
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.OneDriveLocation)) -eq $False) {
            $DlpCompliancePolicyOneDriveLocation = $($DlpCompliancePolicy.OneDriveLocation|out-string).trim()
         }
         else {
            $DlpCompliancePolicyOneDriveLocation = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "OneDrive Location"
           "Value"                    = $DlpCompliancePolicyOneDriveLocation
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.OneDriveLocationException)) -eq $False) {
            $DlpCompliancePolicyOneDriveLocationException = $($DlpCompliancePolicy.OneDriveLocationException|out-string).trim()
         }
         else {
            $DlpCompliancePolicyOneDriveLocationException = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "OneDrive Location Exception"
           "Value"                    = $DlpCompliancePolicyOneDriveLocationException
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.ExchangeOnPremisesLocation)) -eq $False) {
            $DlpCompliancePolicyExchangeOnPremisesLocation = $($DlpCompliancePolicy.ExchangeOnPremisesLocation|out-string).trim()
         }
         else {
            $DlpCompliancePolicyExchangeOnPremisesLocation = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange On-Premises Location"
           "Value"                    = $DlpCompliancePolicyExchangeOnPremisesLocation
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::IsNullOrEmpty($DlpCompliancePolicy.SharePointOnPremisesLocation)) -eq $False) {
            $DlpCompliancePolicySharePointOnPremisesLocation = $($DlpCompliancePolicy.SharePointOnPremisesLocation|out-string).trim()
         }
         else {
            $DlpCompliancePolicySharePointOnPremisesLocation = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "SharePoint On-Premises Location"
           "Value"                    = $DlpCompliancePolicySharePointOnPremisesLocation
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::isnullorempty($DlpCompliancePolicy.SharePointOnPremisesLocationException)) -eq $False) {
            $DlpCompliancePolicySharePointOnPremisesLocationException = $($DlpCompliancePolicy.SharePointOnPremisesLocationException|out-string).trim()
         }
         else {
            $DlpCompliancePolicySharePointOnPremisesLocationException = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "SharePoint On-Premises Location Exception"
           "Value"                    = $DlpCompliancePolicySharePointOnPremisesLocationException
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::isnullorempty($DlpCompliancePolicy.teamsLocation)) -eq $False) {
            $DlpCompliancePolicyteamsLocation = $($DlpCompliancePolicy.teamsLocation|out-string).trim()
         }
         else {
            $DlpCompliancePolicyteamsLocation = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Teams Location"
           "Value"                    = $DlpCompliancePolicyteamsLocation
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::isnullorempty($DlpCompliancePolicy.teamsLocationexception)) -eq $False) {
            $DlpCompliancePolicyteamsLocationexception = $($DlpCompliancePolicy.teamsLocationexception|out-string).trim()
         }
         else {
            $DlpCompliancePolicyteamsLocationexception = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Teams Location Exception"
           "Value"                    = $DlpCompliancePolicyteamsLocationexception
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::isnullorempty($DlpCompliancePolicy.ExchangeSender)) -eq $False) {
            $DlpCompliancePolicyExchangeSender = $($DlpCompliancePolicy.ExchangeSender|out-string).trim()
         }
         else {
            $DlpCompliancePolicyExchangeSender = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange Sender"
           "Value"                    = $DlpCompliancePolicyExchangeSender
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::isnullorempty($DlpCompliancePolicy.ExchangeSenderMemberOf)) -eq $False) {
            $DlpCompliancePolicyExchangeSenderMemberOf = $($DlpCompliancePolicy.ExchangeSenderMemberOf|out-string).trim()
         }
         else {
            $DlpCompliancePolicyExchangeSenderMemberOf = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange Sender Member Of"
           "Value"                    = $DlpCompliancePolicyExchangeSenderMemberOf
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::isnullorempty($DlpCompliancePolicy.ExchangeSenderException)) -eq $False) {
            $DlpCompliancePolicyExchangeSenderException = $($DlpCompliancePolicy.ExchangeSenderException|out-string).trim()
         }
         else {
            $DlpCompliancePolicyExchangeSenderException = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange Sender Exception"
           "Value"                    = $DlpCompliancePolicyExchangeSenderException
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject

         If (([string]::isnullorempty($DlpCompliancePolicy.ExchangeSenderMemberOFException)) -eq $False) {
            $DlpCompliancePolicyExchangeSenderMemberOFException = $($DlpCompliancePolicy.ExchangeSenderMemberOFException|out-string).trim()
         }
         else {
            $DlpCompliancePolicyExchangeSenderMemberOFException = "Not Configured"
         }
         $DlpCompliancePolicyDetail = [ordered]@{
           "Configuration Item"       = "Exchange Sender Member Of Exception"
           "Value"                    = $DlpCompliancePolicyExchangeSenderMemberOFException
         }
         $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
         $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
    }
    else{
        $DlpCompliancePolicy |foreach-object {
            If (([string]::isnullorempty($_.name)) -eq $False) {
                $DlpCompliancePolicyname = $($_.name|out-string).trim()
             }
             else {
                $RetentionPolicyname = "Not Configured"
             }
             $DlpCompliancePolicyDetail = [ordered]@{
               "Configuration Item"       = "Name [TBA]"
               "Value"                    = $DlpCompliancePolicyname
             }
             $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
             $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
             If (([string]::isnullorempty($_.enabled)) -eq $False) {
                 $DlpCompliancePolicyenabled = $($_.enabled|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyenabled = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Enabled"
                "Value"                    = $DlpCompliancePolicyenabled
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.workload)) -eq $False) {
                 $DlpCompliancePolicyworkload = $($_.workload|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyworkload = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Workload"
                "Value"                    = $DlpCompliancePolicyworkload
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.mode)) -eq $False) {
                 $DlpCompliancePolicymode = $($_.mode|out-string).trim()
              }
              else {
                 $DlpCompliancePolicymode = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Mode"
                "Value"                    = $DlpCompliancePolicymode
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.type)) -eq $False) {
                 $DlpCompliancePolicytype = $($_.type|out-string).trim()
              }
              else {
                 $DlpCompliancePolicytype = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Type"
                "Value"                    = $DlpCompliancePolicytype
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.ExchangeLocation)) -eq $False) {
                 $DlpCompliancePolicyExchangeLocation = $($_.ExchangeLocation|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyExchangeLocation = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange Location"
                "Value"                    = $DlpCompliancePolicyExchangeLocation
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.SharePointLocation)) -eq $False) {
                 $DlpCompliancePolicySharePointLocation = $($_.SharePointLocation|out-string).trim()
              }
              else {
                 $DlpCompliancePolicySharePointLocation = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "SharePoint Location"
                "Value"                    = $DlpCompliancePolicySharePointLocation
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.SharePointLocationException)) -eq $False) {
                 $DlpCompliancePolicySharePointLocationException = $($_.SharePointLocationException|out-string).trim()
              }
              else {
                 $DlpCompliancePolicySharePointLocationException = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "SharePoint Location Exception"
                "Value"                    = $DlpCompliancePolicySharePointLocationException
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.OneDriveLocation)) -eq $False) {
                 $DlpCompliancePolicyOneDriveLocation = $($_.OneDriveLocation|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyOneDriveLocation = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "OneDrive Location"
                "Value"                    = $DlpCompliancePolicyOneDriveLocation
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.OneDriveLocationException)) -eq $False) {
                 $DlpCompliancePolicyOneDriveLocationException = $($_.OneDriveLocationException|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyOneDriveLocationException = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "OneDrive Location Exception"
                "Value"                    = $DlpCompliancePolicyOneDriveLocationException
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.ExchangeOnPremisesLocation)) -eq $False) {
                 $DlpCompliancePolicyExchangeOnPremisesLocation = $($_.ExchangeOnPremisesLocation|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyExchangeOnPremisesLocation = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange On-Premises Location"
                "Value"                    = $DlpCompliancePolicyExchangeOnPremisesLocation
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.SharePointOnPremisesLocation)) -eq $False) {
                 $DlpCompliancePolicySharePointOnPremisesLocation = $($_.SharePointOnPremisesLocation|out-string).trim()
              }
              else {
                 $DlpCompliancePolicySharePointOnPremisesLocation = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "SharePoint On-Premises Location"
                "Value"                    = $DlpCompliancePolicySharePointOnPremisesLocation
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.SharePointOnPremisesLocationException)) -eq $False) {
                 $DlpCompliancePolicySharePointOnPremisesLocationException = $($_.SharePointOnPremisesLocationException|out-string).trim()
              }
              else {
                 $DlpCompliancePolicySharePointOnPremisesLocationException = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "SharePoint On-Premises Location Exception"
                "Value"                    = $DlpCompliancePolicySharePointOnPremisesLocationException
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.teamsLocation)) -eq $False) {
                 $DlpCompliancePolicyteamsLocation = $($_.teamsLocation|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyteamsLocation = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Teams Location"
                "Value"                    = $DlpCompliancePolicyteamsLocation
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.teamsLocationexception)) -eq $False) {
                 $DlpCompliancePolicyteamsLocationexception = $($_.teamsLocationexception|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyteamsLocationexception = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Teams Location Exception"
                "Value"                    = $DlpCompliancePolicyteamsLocationexception
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.ExchangeSender)) -eq $False) {
                 $DlpCompliancePolicyExchangeSender = $($_.ExchangeSender|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyExchangeSender = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange Sender"
                "Value"                    = $DlpCompliancePolicyExchangeSender
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.ExchangeSenderMemberOf)) -eq $False) {
                 $DlpCompliancePolicyExchangeSenderMemberOf = $($_.ExchangeSenderMemberOf|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyExchangeSenderMemberOf = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange Sender Member Of"
                "Value"                    = $DlpCompliancePolicyExchangeSenderMemberOf
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.ExchangeSenderException)) -eq $False) {
                 $DlpCompliancePolicyExchangeSenderException = $($_.ExchangeSenderException|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyExchangeSenderException = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange Sender Exception"
                "Value"                    = $DlpCompliancePolicyExchangeSenderException
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
     
              If (([string]::isnullorempty($_.ExchangeSenderMemberOFException)) -eq $False) {
                 $DlpCompliancePolicyExchangeSenderMemberOFException = $($_.ExchangeSenderMemberOFException|out-string).trim()
              }
              else {
                 $DlpCompliancePolicyExchangeSenderMemberOFException = "Not Configured"
              }
              $DlpCompliancePolicyDetail = [ordered]@{
                "Configuration Item"       = "Exchange Sender Member Of Exception"
                "Value"                    = $DlpCompliancePolicyExchangeSenderMemberOFException
              }
              $DlpCompliancePolicyConfigurationObject = New-Object -TypeName psobject -Property $DlpCompliancePolicyDetail
              $DlpCompliancePolicyArray += $DlpCompliancePolicyConfigurationObject
        }
    }
}


######################################################################################################################################################################################################################################################################################################

#############################################
#Document - Report Overview
#############################################
#Insert Heading
Write-Host " - Overview" -foregroundcolor Gray
Set-Content -Path $ReportFileName -value "# Overview"

#############################################
#Document - Purpose
#############################################
Write-Host " - Purpose" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "## Purpose"
Add-Content -path $ReportFileName -value  "The purpose of this as-built-as-configured (ABAC) document is to detail each configuration item (CI) applied to the Department's Exchange Online instance. These CI's align to the design decisions captured within the Department's Office 365 Detailed Design."

#############################################
#Document - ASSOCIATED DOCUMENTATION
#############################################
Write-Host " - Associated Documentation" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n## Associated Documentation"
Add-Content -path $ReportFileName -value  "`nThe following table lists the documents that were referenced during the creation of this ABAC."

$AssociatedDocumentationDetail = [ordered]@{
    "Name" = "GovDesk - Office 365 Design"
    "Version" = "1.0"
    "Date" = "06/2019"
}

#Add new entry
$AssociatedDocumentationDetailObject = New-Object -TypeName psobject -Property $AssociatedDocumentationDetail
$AssociatedDocumentationArray += $AssociatedDocumentationDetailObject

$AssociatedDocumentationDetail = [ordered]@{
    "Name"      = "GovDesk - Platform Design"
    "Version"   = "1.0"
    "Date"      = "06/2019"
}

#Add new entry
$AssociatedDocumentationDetailObject = New-Object -TypeName psobject -Property $AssociatedDocumentationDetail
$AssociatedDocumentationArray += $AssociatedDocumentationDetailObject

$AssociatedDocumentationDetail = [ordered]@{
    "Name"      = "GovDesk - Workstation Design"
    "Version"   = "1.0"
    "Date"      = "06/2019"
}

#Add new entry
$AssociatedDocumentationDetailObject = New-Object -TypeName psobject -Property $AssociatedDocumentationDetail
$AssociatedDocumentationArray += $AssociatedDocumentationDetailObject

$AssociatedDocumentationArray | ConvertTo-MDTest -tabletitle "Associated Documentation" -AsTable | Add-Content -path $ReportFileName

#############################################
#Document - Configuration
#############################################
Write-Host " - Configuration" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "# Configuration"
#############################################
#Document - Exchange Online
#############################################
Add-Content -path $ReportFileName -value "## Exchange Online"
Add-Content -path $ReportFileName -value "The ABAC settings for the Department's Exchange Online instance can be found below. This includes connectors, Mail Exchange (MX) records, SPF, DMARC, DKIM, Remote Domains, User mailbox configurations, Authentication Policies, Outlook on the Web policies, Mailbox Archiving, and Address lists." 
Add-Content -path $ReportFileName -value "Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting."

#############################################
#Document - Exchange Online - Connectors
#############################################
Write-Host " - Inbound Exchange Connectors" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### Connectors"

$GUCount = $($InboundMailConnectorsArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "Exchange Online contains the following inbound mail connectors."
    $MailConnectorsArray = $InboundMailConnectorsArray |Select-Object "Name","Status","TLS","Certificate"
    $MailConnectorsArray | ConvertTo-MDTest -tabletitle "Inbound Mail Connector Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no inbound mail connectors configured in Exchange Online."
}

Write-Host " - Outbound Exchange Connectors" -foregroundcolor Gray

$GUCount = $($OutboundMailConnectorsArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "Exchange Online contains the following outbound mail connectors."
    $MailConnectorsArray = $OutboundMailConnectorsArray |Select-Object "Name","Status","TLS","Certificate"
    $MailConnectorsArray | ConvertTo-MDTest -tabletitle "Outbound Mail Connectors Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no outbound mail connectors configured in Exchange Online."
}

#############################################
#Document - Exchange Online - MX Records
#############################################
Write-Host " - MX Records" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### MX Records"

$GUCount = $($MXrecordsArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following MX records have been configured."
    $MXrecordsArray | ConvertTo-MDTest -tabletitle "MX Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no MX records configured"
}

#############################################
#Document - Exchange Online - SPF Records
#############################################
Write-Host " - SPF Records" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### SPF Records"

$GUCount = $($SPFrecordsArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following SPF records have been configured."
    $SPFrecordsArray | ConvertTo-MDTest -tabletitle "SPF Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no SPF records configured"
}

#############################################
#Document - Exchange Online - Remote Domains
#############################################
Write-Host " - Remote Domains" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Remote Domains"

$GUCount = $($RemoteDomainsArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Remote Domains have been configured."
    $RemoteDomainsArray | ConvertTo-MDTest -tabletitle "Remote Domain Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Remote Domains records configured"
}

#############################################
#Document - Exchange Online - CAS Mailbox Plan
#############################################
Write-Host " - CAS Mailbox Plan" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### CAS Mailbox Plan"

$GUCount = $($CASMailboxPlanArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following CAS Mailbox Plans have been configured."
    $CASMailboxPlanArray | ConvertTo-MDTest -tabletitle "CAS Mailbox Plan Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Remote Domains records configured"
}


#############################################
#Document - Exchange Online - Authentication Policy
#############################################
Write-Host " - Authentication Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Authentication Policy"

$GUCount = $($EOAuthenticationPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Authentication Policies have been configured."
    $EOAuthenticationPolicyArray | ConvertTo-MDTest -tabletitle "Authentication Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Authentication Policies configured"
}

#############################################
#Document - Exchange Online - OWA Policy
#############################################
Write-Host " - OWA Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Outlook Web Access Policy"

$GUCount = $($OWAMailboxPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Outlook Web Access Policies have been configured."
    $OWAMailboxPolicyArray | ConvertTo-MDTest -tabletitle "Outlook Web Access Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Outlook Web Access Policies configured"
}

#############################################
#Document - Exchange Online - Address Lists
#############################################
Write-Host " - Address Lists" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Address Lists"

$GUCount = $($EOAddressListsArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following Address Lists have been configured."
    $EOAddressListsArray | ConvertTo-MDTest -tabletitle "Address List Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Address Lists configured"
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
    $EOPConnectionFilterArray | ConvertTo-MDTest -tabletitle "Connection Filters Configuration" -AsTable | Add-Content -path $ReportFileName
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
    $EOPMalwareFilterArray | ConvertTo-MDTest -tabletitle "Malware Filter Configuration" -AsTable | Add-Content -path $ReportFileName
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
    $EOPPolicyFilterArray | ConvertTo-MDTest -tabletitle "Policy Filter Configuration" -AsTable | Add-Content -path $ReportFileName
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
    $EOPContentFilterArray | ConvertTo-MDTest -tabletitle "Content Filter Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Content Filters configured"
}

#############################################
#Document - Configuration - Teams
#############################################
Add-Content -path $ReportFileName -value "## Teams"
Add-Content -path $ReportFileName -value "The ABAC settings for the Department's Teams instance can be found below. This includes the Client Configuration, Channels Policy, Calling Policy, Meetings Policy, Messaging Policy, and Guest Meeting/Calling/Messaging configuration." 
Add-Content -path $ReportFileName -value "Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting."

#############################################
#Document - Teams - Client Configuration
#############################################
Write-Host " - Client Configuration" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Client Configuration"

$GUCount = $($TeamsClientConfigArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following is the Teams client configuration."
    $TeamsClientConfigArray | ConvertTo-MDTest -tabletitle "Client Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Teams client configurations have been made."
}

#############################################
#Document - Teams - Channel Policy
#############################################
Write-Host " - Channel Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Channel Policy"

$GUCount = $($TeamsChannelPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following is the Teams Channel Policy configuration."
    $TeamsChannelPolicyArray | ConvertTo-MDTest -tabletitle "Channel Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Teams Channel Policy configurations."
}

#############################################
#Document - Teams - Calling Policy
#############################################
Write-Host " - Calling Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Calling Policy"

$GUCount = $($TeamsCallingPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following is the Teams Calling Policy configuration."
    $TeamsCallingPolicyArray | ConvertTo-MDTest -tabletitle "Calling Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Teams Calling Policy configurations."
}

#############################################
#Document - Teams - Meeting Policy
#############################################
Write-Host " - Meeting Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Meeting Policy"

$GUCount = $($TeamsMeetingPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following is the Teams Meeting Policy configuration."
    $TeamsMeetingPolicyArray | ConvertTo-MDTest -tabletitle "Meeting Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Teams Meeting Policy configurations."
}

#############################################
#Document - Teams - Messaging Policy
#############################################
Write-Host " - Messaging Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Messaging Policy"

$GUCount = $($TeamsMessagingPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following is the Teams Messaging Policy configuration."
    $TeamsMessagingPolicyArray | ConvertTo-MDTest -tabletitle "Messaging Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Teams Messaging Policy configurations."
}

#############################################
#Document - Teams - Meeting Broadcast Policy
#############################################
Write-Host " - Meeting Broadcast Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Meeting Broadcast Policy"

$GUCount = $($TeamsMeetingBroadcastPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following is the Teams Meeting Broadcast Policy configuration."
    $TeamsMeetingBroadcastPolicyArray | ConvertTo-MDTest -tabletitle "Meeting Broadcast Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Teams Meeting Broadcast Policy configurations."
}

#############################################
#Document - Configuration - SharePoint
#############################################
Add-Content -path $ReportFileName -value "## SharePoint Online & OneDrive"
Add-Content -path $ReportFileName -value "The ABAC settings for the Department's SharePoint Online and OneDrive instances can be found below. This includes the Site and Sharing configuration." 
Add-Content -path $ReportFileName -value "Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting."

#############################################
#Document - Sharepoint - Tenant configuration
#############################################
Write-Host " - Tenant configuration" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Tenant Configuration"

$GUCount = $($SharepointArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following table lists the SharePoint Online Tenant configuration."
    $SharepointArray | ConvertTo-MDTest -tabletitle "Tenant Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Content Filters configured"
}

#############################################
#Document - Configuration - Security and Compliance
#############################################
Add-Content -path $ReportFileName -value "## Security and Compliance"
Add-Content -path $ReportFileName -value "The ABAC settings for the Department's Office 365 Security and Compliance instance can be found below. This includes the Alerts, Labels, Data Loss Prevention, Retention Policies, Audit Logging, and Customer Key configuration." 
Add-Content -path $ReportFileName -value "Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting."

#############################################
#Document - Security and Compliance - Retention Labels
#############################################
Write-Host " - Retention Labels" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Retention Labels"

$GUCount = $($RetentionLabelArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following table lists the Retention Labels configuration."
    $RetentionLabelArray | ConvertTo-MDTest -tabletitle "Retention Labels Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There are no Retention Labels configured"
}

#############################################
#Document - Security and Compliance - Retention Policies
#############################################
Write-Host " - Retention Label Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Retention Label Policy"

$GUCount = $($RetentionpolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following table lists the Retention Label Policy configuration."
    $RetentionpolicyArray | ConvertTo-MDTest -tabletitle "Retention Label Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There is no Retention Label Policy configured"
}

#############################################
#Document - Security and Compliance - Sensitivity labels
#############################################
Write-Host " - Sensitivity labels" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Sensitivity labels"

$GUCount = $($SensitivityLabelsArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following table lists the Sensitivity Label configuration."
    $SensitivityLabelsArray | ConvertTo-MDTest -tabletitle "Sensitivity Label Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There is no Sensitivity Label configured"
}

#############################################
#Document - Security and Compliance - Sensitivity label Policy
#############################################
Write-Host " - Sensitivity label Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Sensitivity label Policy"

$GUCount = $($SensitivitylabelpolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following table lists the Sensitivity label Policy configuration."
    $SensitivitylabelpolicyArray | ConvertTo-MDTest -tabletitle "Sensitivity label Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There is no Sensitivity label Policy configured"
}
#############################################
#Document - Security and Compliance - Dlp Compliance Policy 
#############################################
Write-Host " - Dlp Compliance Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "### Dlp Compliance Policy"

$GUCount = $($DlpCompliancePolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
    Add-Content -path $ReportFileName -value "The following table lists the Dlp Compliance Policy configuration."
    $DlpCompliancePolicyArray | ConvertTo-MDTest -tabletitle "Dlp Compliance Policy Configuration" -AsTable | Add-Content -path $ReportFileName
}
else {
    Add-Content -path $ReportFileName -value "There is no Dlp Compliance Policy configured"
}



Write-Host " - Done" -foregroundcolor Green
Write-Host ""

Write-Host "Automated Assessment Complete" -foregroundcolor Green
Write-Host "-------------------------------------------------------------------------"
Write-Host ""
Write-Host "File Saved To: $ReportFileName"
Write-Host "Log File Saved To: $Logging"
Write-Host ""
Write-Host "-------------------------------------------------------------------------"
Write-Host "!!! IMPORTANT !!! IMPORTANT !!! IMPORTANT !!! IMPORTANT !!! IMPORTANT !!!"
Write-Host "-------------------------------------------------------------------------"
Write-Host ""
Write-Host "Complete the following tasks: "
Write-Host " - Check for any errors above in the output"
Write-Host " - Save a copy of the log file (detailed above)"
Write-Host " - Import Customer Logo"
Write-Host " - Update the date field and client name on the main page"
Write-Host " - Update the client name in the document properties"
Write-Host " - Complete any sections with the text '!!!! TODO !!!!' or 'TBD'"
Write-Host " - Review all recommendations and analysis"
Write-Host " - Perform full read through"
Write-Host " - Have this document reviewed by a senior resource"
Write-Host ""
Write-Host "-------------------------------------------------------------------------"
Write-Host ""

#Stop Logging
Stop-Transcript

#FUTURE AUTOMATION
########################################################
#Secure Score - MS Graph
#Risk Events - MS Graph
#PIM enabled - https://ittechnews.net/microsoft/official-microsoft-news/powershell-sample-for-privileged-identity-management-pim-for-azure-ad-roles/
#AAD Audit = https://blogs.technet.microsoft.com/motiba/2018/02/12/list-of-azure-active-directory-audit-activities/
# $url  = 'https://graph.windows.net' + $tenantdomain + '/activities/signinEvents?api-version=beta&`$filter=signinDateTime ge ' + $7daysago

# UsersPermissionToCreateGroupsEnabled    : True
# UsersPermissionToCreateLOBAppsEnabled   : False
# UsersPermissionToReadOtherUsersEnabled  : True
# UsersPermissionToUserConsentToAppEnabled: True

########################################################

#Re-enable TODO
#TEST - Validate!