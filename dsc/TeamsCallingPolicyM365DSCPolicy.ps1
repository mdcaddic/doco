Configuration M365
{
    Import-DscResource -ModuleName Microsoft365DSC
    $password = ConvertTo-SecureString "Hol48620" -AsPlainText -Force
    $GlobalAdmin = New-Object System.Management.Automation.PSCredential ('ga@oobedta.onmicrosoft.com',$password)

    
    Node localhost
    {
        #region Teams
        TeamsTeam DevOPS
        {
            DisplayName        = "DevOPS Demo"
            Description        = "This is me demoing the Teams Resource"
            GlobalAdminAccount = $GlobalAdmin
            Ensure             = "Present"
        }
        TeamsChannel DSCChannel
        {
            TeamName           = "DevOPS Demo"
            DisplayName        = "DSC Discussions"
            GlobalAdminAccount = $GlobalAdmin
            Ensure             = "Present"
        }
        #endregion
    }
}
$ConfigData = @{
    AllNodes = @(
        @{
        NodeName = "localhost"
        PsDSCAllowPlainTextPassword = $true
        PSDscAllowDomainUser = $true
        }
    )
}
M365 -ConfigurationData $ConfigData
Start-DscConfiguration M365 -Wait -Force -Verbose
