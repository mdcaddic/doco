TeamsCallingPolicy M365DSCPolicy
{
    AllowCallForwardingToPhone = $True; ### L2|We <strong>recommend</strong> allowing call forwarding to phone lines because...
    AllowCallForwardingToUser  = $True; ### L3|Information about <a href="https://docs.microsoft.com/en-us/MicrosoftTeams/teams-calling-policy">call forwarding</a>
    AllowPrivateCalling        = $false; ### L1|<img src='warning.png' />We don't recommend allowing people to allow private calls due to....
    AllowVoicemail             = "UserOverride";
    BusyOnBusyEnabledType      = "Disabled";
    Identity                   = "Microsoft365DSC Policy";
    PreventTollBypass          = $True;
    Ensure                     = "Present";
    GlobalAdminAccount         = $GlobalAdminAccount
}