install-module az.accounts # -force -allowclobber
install-module az.automation #-force -allowclobber
 
#Update the values below specific to your tenant!
$tenantID = "dc536b4a-8910-4ef0-8816-90b496a9ae29"
$subscriptionID = "8407beb3-507f-4d85-9311-7fe21b12cd60"
$automationAccount = "m365dsc-assess"
$resourceGroup = "azure-automation-m365dsc"
 
$moduleName = "Microsoft365DSC"
Connect-AzAccount -SubscriptionId $subscriptionID -Tenant $tenantID
 
Function Get-Dependency {
#Function modifed from: https://4bes.nl/2019/09/05/script-update-all-powershell-modules-in-your-automation-account/
    param(
        [Parameter(Mandatory = $true)]
        [string] $ModuleName   
    )
 
    $OrderedModules = [System.Collections.ArrayList]@()
     
    # Getting dependencies from the gallery
    # Write-Verbose "Checking dependencies for $ModuleName"
    # $ModuleUri = "https://www.powershellgallery.com/api/v2/Search()?`$filter={1}" # &amp;searchTerm=%27{0}%27&amp;targetFramework=%27%27&amp;includePrerelease=false&amp;`$skip=0&amp;`$top=40"
    # $CurrentModuleUrl = $ModuleUri -f $ModuleName, 'IsLatestVersion'
    # $SearchResult = Invoke-RestMethod -Method Get -Uri $CurrentModuleUrl -UseBasicParsing | Where-Object { $_.title.InnerText -eq $ModuleName }
 
    # if ($null -eq $SearchResult) {
    #     Write-Output "Could not find module $ModuleName in PowerShell Gallery."
    #     Continue
    # }
    # $ModuleInformation = (Invoke-RestMethod -Method Get -UseBasicParsing -Uri $SearchResult.id)
    Write-Verbose "Checking dependencies for $ModuleName"
    $ModuleUri = "https://www.powershellgallery.com/api/v2/Search()?`$filter={1}&searchTerm=%27{0}%27&targetFramework=%27%27&includePrerelease=false&`$skip=0&`$top=40"
    $CurrentModuleUrl = $ModuleUri -f $ModuleName, 'IsLatestVersion'
    $SearchResult = Invoke-RestMethod -Method Get -Uri $CurrentModuleUrl -UseBasicParsing | Where-Object { $_.title.InnerText -eq $ModuleName }

    if ($null -eq $SearchResult) {
        Write-Output "Could not find module $ModuleName in PowerShell Gallery."
        Continue
    }
    $ModuleInformation = (Invoke-RestMethod -Method Get -UseBasicParsing -Uri $SearchResult.id)

    #Creating Variables to get an object
    $ModuleVersion = $ModuleInformation.entry.properties.version
    $Dependencies = $ModuleInformation.entry.properties.dependencies
    $DependencyReadable = $Dependencies -split ":\|"
 
    $ModuleObject = [PSCustomObject]@{
        ModuleName    = $ModuleName
        ModuleVersion = $ModuleVersion
    }
     
    # If no dependencies are found, the module is added to the list
    if (![string]::IsNullOrEmpty($Dependencies) ) {
        foreach ($dependency in $DependencyReadable){
            $DepenencyObject = [PSCustomObject]@{
                ModuleName    = $($dependency.split(':')[0])
                ModuleVersion = $($dependency.split(':')[1].substring(1).split(',')[0])
            }
            $OrderedModules.Add($DepenencyObject) | Out-Null
        }
    }
 
    $OrderedModules.Add($ModuleObject) | Out-Null
 
    return $OrderedModules
}
 
$ModulesAndDependencies = Get-Dependency -moduleName $moduleName
#$ModulesAndDependencies
 
write-output "Installing $($ModulesAndDependencies | ConvertTo-Json)"
 
#Install Module and Dependencies into Automation Account
foreach($module in $ModulesAndDependencies){
    $CheckInsalled = get-AzAutomationModule -AutomationAccountName $automationAccount -ResourceGroupName $resourceGroup -Name $($module.modulename) -ErrorAction SilentlyContinue
    if($CheckInsalled.ProvisioningState -eq "Succeeded" -and $CheckInsalled.Version -ge $module.ModuleVersion){
        write-output "$($module.modulename) existing: v$($CheckInsalled.Version), required: v$($module.moduleVersion)"
    }
    else{
        New-AzAutomationModule -AutomationAccountName $automationAccount -ResourceGroupName $resourceGroup -Name $($module.modulename) -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/$($module.modulename)/$($module.moduleVersion)" -Verbose    
        While($(get-AzAutomationModule -AutomationAccountName $automationAccount -ResourceGroupName $resourceGroup -Name $($module.modulename)).ProvisioningState -eq 'Creating'){
            Write-output 'Importing $($module.modulename)...'
            start-sleep -Seconds 10
        }
    }
}