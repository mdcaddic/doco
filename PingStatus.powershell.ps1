## Note: This script uses Test-Connection to see if the device is alive.
# It is also possible to use Test-NetConnection and then pass a port number (ie. 80 or 443)
# to see if the service is alive


$date = Get-Date -Format ddMMMyyyy

$a = @'
<style>
    body{
        background-color:antiquewhite;
    }
    table, td, th{
        border-width: 1px;
        border-style: solid;
        border-color: black;
        border-collapse: collapse;
    }
    th {
        background-color:thistle
    }
    td {
        background-color:PaleGoldenrod
    }
    tr.red { color: red; }
    tr.green { color: green; }
</style>
'@
$html = ""
$htmlFile = ((Get-Location).path + "\" + 'connection.html')
$pre = '<h2>Ping results for servers</h2><br />'
$pre = $pre + $date
$post = '<h3>PING RESULTS</h3>'
$pingservers = @('localhost','4125host3','thisdoesntresolve')
# $pingtarget = @{}
# $pingtarget.Add("4125host3","3389,80")
# # $pingtarget.Add("4125host2","3389,443")

$Ports  = "135","389","636","3268","53","88","445","3269", "80", "443"
$AllDCs = Get-ADDomainController -Filter * | Select-Object Hostname,Ipv4address,isGlobalCatalog,Site,Forest,OperatingSystem
ForEach($DC in $AllDCs)
{
Foreach ($P in $Ports){
$check=Test-NetConnection $DC.Hostname -Port $P -WarningAction SilentlyContinue
If ($check.tcpTestSucceeded -eq $true)
    {Write-Host $DC.Hostname $P -ForegroundColor Green -Separator " => "}
else
    {Write-Host $DC.Hostname $P -Separator " => " -ForegroundColor Red}
}
}

# $pingtarget.GetEnumerator() | ForEach-Object {
#     if (Resolve-DnsName -Name $_.Key -Type A){
#         # If port (value) is an array split it
#         if ($_.Value -match ","){
#             $ports = $_.value.split(",")
#             ForEach ($port in $ports){
#                 $return = Test-NetConnection -ComputerName $_.Key -Port $port
#                 $html += $_.Key + $return.RemoteAddress + $return.TcpTestSucceeded
#             }
#         }else{
#             # Only monitoring one port
#             $return = Test-NetConnection -ComputerName $_.Key -Port $_.Value
#             $html += $return.TcpTestSucceeded           
#         }

#     }else{
#         Write-Verbose "Could not resolve " $_.Key
#     }
# # $html = $return
# }


$html = $pingservers |
    ForEach-Object{
        # Resolve name
        Resolve-DNSName -Name $_ -type A | Select-Object Name, IPAddress
        # If resolves, check each port listed

               @{n = 'Responding'; e = { (Test-NetConnection -ComputerName $_.Name -port ).PingSucceeded } }
    } | 
    ConvertTo-HTML -head $a -PreContent $pre -PostContent $post

[xml]$xml = $html
$rowCount = $xml.html.body.table.tr.Count - 1
$xml.html.body.table.tr[1.. $rowCount] | 
    ForEach-Object{
        $class = if($_.td[2] -eq 'True'){'green' }else{ 'red' }
        $_.SetAttribute('class', $class)
    }
$xml.Save($htmlFile)

. $htmlFile 