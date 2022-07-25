Import-Csv C:ip-Addresses.csv | ForEach-Object{
$hostname = ([System.Net.Dns]::GetHostByAddress($_.IPAddress)).Hostname
if($? -eq $False){
$hostname="Cannot resolve hostname"
}
New-Object -TypeName PSObject -Property @{
      IPAddress = $_.IPAddress
      HostName = $hostname
}} | Export-Csv C:machinenames.csv -NoTypeInformation -Encoding UTF8