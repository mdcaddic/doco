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
#endregion Arrays

###########################
#Variables
###########################
Set-Location $PSScriptroot
$domains                        = @("precisionservices.biz")
$Logging                        = "$pwd\Logs\ExchangeOnlineABAC_$($(Get-Date).ToString(`"yyyy-MM-dd hhmmss`")).log"
$ReportFileName                 = "$pwd\Reports\EOL_MDTest.md" # "$pwd\Reports\Office365ABAC_$($(Get-Date).ToString(`"yyyy-MM-dd hhmmss`")).md"
# $Tenant                         = "precisionservicesptyltd"
# $admindomain                    = "precisionservicesptyltd.onmicrosoft.com"
$global:UserPrincipalName       = "martin@precisionservices.biz"

#Enable Full Logging
Start-Transcript -Path $Logging
Write-Host ""
####################
# Start writing document
# Overview
Write-Host " - Overview" -foregroundcolor Gray
Set-Content -Path $ReportFileName -value "# Overview"
# Purpose
Write-Host " - Purpose" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n"
Add-Content -path $ReportFileName -value "## Purpose"
Add-Content -path $ReportFileName -value  "The purpose of this as-built-as-configured (ABAC) document is to detail each configuration item (CI) applied to the Department's Exchange Online instance. These CI's align to the design decisions captured within the Department's Office 365 Detailed Design."
# Associated Documentation
Write-Host " - Associated Documentation" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n"
Add-Content -path $ReportFileName -value "## Associated Documentation"
Add-Content -path $ReportFileName -value  "The following table lists the documents that were referenced during the creation of this ABAC."
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
# Write the associated documentation
Add-Content -path $ReportFileName -value "`n"
$AssociatedDocumentationArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
# Start writing Configuration
Write-Host " - Configuration" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n"
Add-Content -path $ReportFileName -value "# Configuration"

#region Exchange Online
###########################
#Connect to Exchange Online
###########################
# Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
# Connect-EXOPSSession -UserPrincipalName $UserPrincipalName
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
Write-Host " - Done" -foregroundcolor Green
Write-Host ""
# Start writing
Add-Content -path $ReportFileName -value "`n"
Add-Content -path $ReportFileName -value "## Exchange Online"
Add-Content -path $ReportFileName -value "The ABAC settings for the Department's Exchange Online instance can be found below. This includes connectors, Mail Exchange (MX) records, SPF, DMARC, DKIM, Remote Domains, User mailbox configurations, Authentication Policies, Outlook on the Web policies, Mailbox Archiving, and Address lists." 
Add-Content -path $ReportFileName -value "Please note, if a setting is not mentioned in the below, it should be assumed to have been left at its default setting."

#############################################
#Exchange Online
#############################################
Write-Host "Querying Exchange Online configuration..." -foregroundcolor Yellow
#region Mail Connectors
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
# Write connector information
Write-Host " - Inbound Exchange Connectors" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n"
Add-Content -path $ReportFileName -value "### Connectors"

$GUCount = $($InboundMailConnectorsArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "Exchange Online contains the following inbound mail connectors."
$MailConnectorsArray = $InboundMailConnectorsArray |Select-Object "Name","Status","TLS","Certificate"
Add-Content -path $ReportFileName -value "`n"
$MailConnectorsArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no inbound mail connectors configured in Exchange Online."
}

Write-Host " - Outbound Exchange Connectors" -foregroundcolor Gray

$GUCount = $($OutboundMailConnectorsArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "Exchange Online contains the following outbound mail connectors."
$MailConnectorsArray = $OutboundMailConnectorsArray |Select-Object "Name","Status","TLS","Certificate"
Add-Content -path $ReportFileName -value "`n"
$MailConnectorsArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no outbound mail connectors configured in Exchange Online."
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
#endregion  Mail Connectors

#region MX Records
#############################################
#Exchange Online - MX Records
#############################################
Write-Host " - Mail Exchange Records" -foregroundcolor Gray

# $EOMXrecordPreferencelist =@()
If ($null -ne $domains) {
$domains |Foreach-Object {
    # $EOMXrecords = nslookup.exe -type=MX $_ 2>$null |select-string "MX"
    $EOMXrecords = Resolve-DnsName -Type MX -Name $_ 
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
            # $EOMXrecords = $EOMXrecords|Out-String
            # $EOMXrecords = [string]$EOMXrecords.Substring([string]$EOMXrecords.IndexOf("MX"))
            $EOMXrecordPreference = $EOMXrecords.Preference #[string]$EOMXrecords.Substring($EOMXrecords.IndexOf("MX"),$EOMXrecords.IndexOf(","))
            $EOMXrecordMX = $EOMXrecords.NameExchange # [string]$EOMXrecords.Substring($EOMXrecords.IndexOf("mail"))            
            $EOMXrecordsDetail = [ordered]@{
                "Domain"              = $_ 
                "MX Preference"       = $EOMXrecordPreference
                "Mail Exchanger"      = $EOMXrecordMX
            }
            $EOConfigurationObject = New-Object -TypeName psobject -Property $EOMXrecordsDetail
            $MXrecordsArray += $EOConfigurationObject  
        }
        Else {
            ## Needs to be reworked for multiple MX records

            # Foreach($EOMXrecord in $EOMXrecords) {
            #     if ($EOMXrecord -eq $EOMXrecords[0]) {
            #         # $EOMXrecord = $EOMXrecord|Out-String
            #         # $EOMXrecord = [string]$EOMXrecord.Substring([string]$EOMXrecord.IndexOf("MX"))
            #         $EOMXrecordPreference = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("MX"),$EOMXrecord.IndexOf(","))
            #         $EOMXrecordMX = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("mail"))  
            #         $EOMXrecordMX = $EOMXrecordMX.TrimEnd()      
            #         [string]$EOMXrecordPreferencelist +=  [string]$EOMXrecordPreference
            #         [string]$EOMXrecordMXlist +=  [string]$EOMXrecordMX
            #     }
            #     else {
            #         $EOMXrecord = $EOMXrecord|Out-String
            #         $EOMXrecord = [string]$EOMXrecord.Substring([string]$EOMXrecord.IndexOf("MX"))
            #         $EOMXrecordPreference = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("MX"),$EOMXrecord.IndexOf(","))
            #         $EOMXrecordMX = [string]$EOMXrecord.Substring($EOMXrecord.IndexOf("mail"))    
            #         $EOMXrecordMX = $EOMXrecordMX.TrimEnd()     
            #         [string]$EOMXrecordPreferencelist += "`n" + [string]$EOMXrecordPreference
            #         [string]$EOMXrecordMXlist += "`n" + [string]$EOMXrecordMX                    
            #     }
            # }
            # $EOMXrecordsDetail = [ordered]@{
            #     "Domain"              = $_ 
            #     "MX Preference"       = $EOMXrecordPreferencelist
            #     "Mail Exchanger"      = $EOMXrecordMXlist
            # }
            # $EOMXrecordPreferencelist =$null
            # $EOMXrecordMXlist =$null
            # $EOConfigurationObject = New-Object -TypeName psobject -Property $EOMXrecordsDetail
            # $MXrecordsArray += $EOConfigurationObject
        }
    }
}

}
#############################################
#Document - Exchange Online - MX Records
#############################################
Write-Host " - MX Records" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### MX Records"

$GUCount = $($MXrecordsArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "The following MX records have been configured."
Add-Content -path $ReportFileName -value "`n"
$MXrecordsArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no MX records configured"
}
#endregion MX Records

#region SPF Records
#############################################
#Exchange Online - SPF Records
#############################################
Write-Host " - SPF Records and DMARC" -foregroundcolor Gray

If ($null -ne $domains) {
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
#Document - Exchange Online - SPF Records
#############################################
Write-Host " - SPF Records" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### SPF Records"

$GUCount = $($SPFrecordsArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "The following SPF records have been configured."
Add-Content -path $ReportFileName -value "`n"
$SPFrecordsArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no SPF records configured"
}
#endregion SPF Records

#region Remote Domains
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
#Document - Exchange Online - Remote Domains
#############################################
Write-Host " - Remote Domains" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### Remote Domains"

$GUCount = $($RemoteDomainsArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "The following Remote Domains have been configured."
Add-Content -path $ReportFileName -value "`n"
$RemoteDomainsArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no Remote Domains records configured"
}
#endregion Remote Domains

#region CAS Mailbox Plan
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
#Document - Exchange Online - CAS Mailbox Plan
#############################################
Write-Host " - CAS Mailbox Plan" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### CAS Mailbox Plan"

$GUCount = $($CASMailboxPlanArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "The following CAS Mailbox Plans have been configured."
Add-Content -path $ReportFileName -value "`n"
$CASMailboxPlanArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no Remote Domains records configured"
}
#endregion CAS Mailbox Plan

#region Authentication Policy
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
#Document - Exchange Online - Authentication Policy
#############################################
Write-Host " - Authentication Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### Authentication Policy"

$GUCount = $($EOAuthenticationPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "The following Authentication Policies have been configured."
Add-Content -path $ReportFileName -value "`n"
$EOAuthenticationPolicyArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no Authentication Policies configured"
}
#endregion Authentication Policy

#region OWA Policy
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
#Document - Exchange Online - OWA Policy
#############################################
Write-Host " - OWA Policy" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### Outlook Web Access Policy"

$GUCount = $($OWAMailboxPolicyArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "The following Outlook Web Access Policies have been configured."
Add-Content -path $ReportFileName -value "`n"
$OWAMailboxPolicyArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no Outlook Web Access Policies configured"
}
#endregion OWA Policy

#region Address Lists
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
#############################################
#Document - Exchange Online - Address Lists
#############################################
Write-Host " - Address Lists" -foregroundcolor Gray
Add-Content -path $ReportFileName -value "`n### Address Lists"

$GUCount = $($EOAddressListsArray | Measure-Object).Count
If ($GUCount -gt 0) {
Add-Content -path $ReportFileName -value "The following Address Lists have been configured."
Add-Content -path $ReportFileName -value "`n"
$EOAddressListsArray | ConvertTo-MDTest -AsTable | Add-Content -path $ReportFileName
}
else {
Add-Content -path $ReportFileName -value "There are no Address Lists configured"
}
#endregion Address Lists

#endregion

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

