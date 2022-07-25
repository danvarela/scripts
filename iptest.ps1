 Get-Content ".\ip-Addresses.csv" |
     ForEach-Object{
         $result = [ordered]@{
                     DNSName     = $_.Trim()
                     HostName    = ""
                     DNSIPv4     = ""
                     Up          = ""
                     TestedIPv4  = ""
                     }
         Try{
             $entry = [System.Net.Dns]::GetHostEntry($_.Trim())
             $result.HostName = $entry.HostName
             $entry.AddressList |
                 Where-Object {$_.AddressFamily -eq 'InterNetwork'} |
                     ForEach-Object{
                         $result.DNSIPv4 = $_.IPAddressToString
                         [array]$x = Test-Connection -Delay 15 -ComputerName $result.DNSName -Count 1 -ErrorAction SilentlyContinue
                         if ($x){
                             $result.Up = "Yes"
                             $result.TestedIPv4 = $x[0].IPV4Address
                         }
                         else{
                             $result.Up = "No"
                             $result.TestedIPv4 = "N/A"
                         }
                     }
         }
         Catch{
             $result.HostName = "Host Unknown"
             $result.Up = "Unknown"
             $result.TestedIPv4 = "N/A"
         }
         [PSCustomObject]$result
     } | Export-Csv C:output.csv -NoTypeInformation -Encoding UTF8