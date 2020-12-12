
$PC = get-adcomputer -Filter * -searchbase 'OU=Computers,DC=Domain,DC=local' -Properties name | select -ExcludeProperty name | sort name

$AllComputerNames = $PC.name

$AllComputerNames

$i = 0


#$OnlineTest = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet

$results = @()

    foreach ($ComputerName in $AllComputerNames) {

            $PingTest = Test-Connection -ComputerName $ComputerName -count 1 -quiet

 if (($PingTest -eq $true)) {
    
    $ErrorActionPreference = 'silentlycontinue'

    $Service = Get-Service winrm -ComputerName $ComputerName

        $Service | start-service
                  
$i++
write-progress -Activity "Pulling Data for $ComputerName" -Status: "$i of $($AllComputerNames.count)"                  

                $ComputerOS = Get-ADComputer $ComputerName -Properties OperatingSystem,OperatingSystemServicePack, operatingsystemversion, DistinguishedName
                $ComputerHW = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName | select Manufacturer,Model,NumberOfProcessors,@{Expression={$_.TotalPhysicalMemory / 1GB};Label="TotalPhysicalMemoryGB"}
                $ComputerCPU = Get-WmiObject win32_processor -ComputerName $ComputerName | select DeviceID,Name,Manufacturer,NumberOfCores,NumberOfLogicalProcessors
                $ComputerAddress = get-wmiobject win32_networkadapterconfiguration -computername $ComputerName -filter "IPEnabled='True'" |
                    Select DHCPEnabled, MACAddress,@{Name="IP";Expression={$_.IPAddress[0]}},
                    @{Name="DefaultGateway";Expression={$_.DefaultIPGateway[0]}},
                    DNSHostname,@{Name="DNSServer";Expression={$_.DNSServerSearchOrder[0]}} | Select IP, DHCPEnabled, MACAddress
                #$ComputerLogon = Get-WinEvent -Computer -filterhashtable $ComputerName  @{ Logname = ‘Security’; ID = 4672 } -MaxEvents 1 | Select -ExpandProperty @{ N = ‘User’; E = { $_.Properties[1].Value } }
                $FindUser = (Get-WmiObject -Class Win32_Process -ComputerName $Computername -Filter 'Name="explorer.exe"').
                    GetOwner().
                    User
                $ComputerDisksALT = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $ComputerName | select DeviceID,VolumeName,@{Expression={$_.Size / 1GB};Label="SizeGB"} 
                $ComputerDisks = get-ciminstance win32_diskdrive -ComputerName $Computername | select DeviceID,model, caption,@{Expression={$_.Size / 1GB};Label="SizeGB"} 
                $ComputerAVStatus = get-service EPSecurityService -ComputerName $Computername | select -ExpandProperty status
                $ComputerBios = Get-CimInstance win32_Bios -ComputerName $ComputerName | Select smbiosbiosversion, serialnumber, releasedate
                #$ComputerSerial = Get-CimInstance win32_bios -ComputerName $ComputerName | select serialnumber
                $ComputerMonitor = Get-CimInstance -namespace root\wmi -ClassName wmimonitorid -ComputerName $ComputerName | foreach {New-Object -TypeName psobject -Property @{Manufacturer = ($_.ManufacturerName -notmatch '^0$' | foreach {[char]$_}) -join "" }}
                $ComputerMonitorNames = Get-CimInstance -Namespace root\wmi -ClassName wmimonitorid -ComputerName $ComputerName | Foreach {New-Object -TypeName psobject -Property @{Names = ($_.UserFriendlyName -notmatch '^0$' | foreach {[char]$_}) -join ""}}
                $ComputerMonitorSerial = Get-CimInstance -Namespace root\wmi -classname wmimonitorid -computername $ComputerName | Foreach {New-Object -TypeName psobject -Property @{Serial = ($_.SerialNumberID -notmatch '^0$' | foreach {[char]$_}) -join ""}
                $VPNService = get-service -ComputerName $ComputerName | where{$_.DisplayName -eq 'VPN Agent'} | select displayname, status
                $AppVersion = Invoke-Command -ComputerName $ComputerName -scriptblock {
                    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*  | where {$_.displayname -eq 'Agent'} | select DisplayVersion
                        }
                $AppVersionAV = Invoke-Command -ComputerName $ComputerName -scriptblock {
                    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*  | where {$_.displayname -eq 'AV'} | select DisplayVersion
                        }
                #Gets Information on RAM
                $RAMINFO = Get-Ciminstance Win32_PhysicalMemory -ComputerName $ComputerName | select devicelocator, configuredclockspeed, serialnumber, manufacturer, @{Name="Ram Capacity"; Expression={[math]::round($_.Capacity/1GB, 3)}}
                $RAMLOC = $RAMINFO.devicelocator
                $RAMClockSpeed = $RAMINFO.configuredclockspeed
                $RAMSN = $RAMINFO.SerialNumber
                $RAMMANU = $RAMINFO.manufacturer
                $RAMCAP = $RAMINFO.'Ram Capacity'

$SN = $ComputerBios.serialnumber
######################################################
#Start Function Retrieves Warranty Status for HP Only
Function Get-HPAssetInfo {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory,ValueFromPipeline)]
        [String]$SerialNumber
    )

    Begin {
        Function Invoke-HPIncSOAPRequest {
            Param (
                [Parameter(Mandatory)]
                [Xml]$SOAPRequest,
                [String]$Url = 'https://api-uns-sgw.external.hp.com/gw/hpit/egit/obligation.sa/1.1'
            )

            $soapWebRequest = [System.Net.WebRequest]::Create($URL) 
            $soapWebRequest.Headers.Add('X-HP-SBS-ApplicationId','hpi-obligation-hpsa')
            $soapWebRequest.Headers.Add('X-HP-SBS-ApplicationKey','ft2VGa2hx9j$')
            $soapWebRequest.ContentType = 'text/xml; charset=utf-8'
            $soapWebRequest.Accept = 'text/xml'
            $soapWebRequest.Method = 'POST'

            try {
                $SOAPRequest.Save(($requestStream = $soapWebRequest.GetRequestStream()))
                $requestStream.Close() 
                $responseStream = ($soapWebRequest.GetResponse()).GetResponseStream()
                [XML]([System.IO.StreamReader]($responseStream)).ReadToEnd()
                $responseStream.Close() 
            }
            catch {
                throw $_
            }
        }
    }

    Process {
        foreach ($S in $SerialNumber) {
            $request = @"
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:int="http://interfaces.obligation.sbs.it.hp.com/">
  <soapenv:Header />
  <soapenv:Body>
    <int:retrieveServiceObligationResponsesByServiceObligationRequests>
      <context>
        <appContextName>HPSF</appContextName>
        <userLocale>en-US</userLocale>
      </context>
      <obligationRequests>
        <lnkServiceObligationDepthFilter>
          <includeProductObjectOfServiceInstance>true</includeProductObjectOfServiceInstance>
          <includeServiceObligation>true</includeServiceObligation>
          <includeServiceObligationHeaderOffer>true</includeServiceObligationHeaderOffer>
          <includeServiceObligationMessage>true</includeServiceObligationMessage>
          <maxNumberOfProductObjectOfServiceInstance>100</maxNumberOfProductObjectOfServiceInstance>
        </lnkServiceObligationDepthFilter>
        <lnkServiceObligationEnrichment>
          <iso2CountryCode>US</iso2CountryCode>
        </lnkServiceObligationEnrichment>
        <lnkServiceObligationProductObjectOfServiceIdentifier>
          <hpSerialNumber>$S</hpSerialNumber>
        </lnkServiceObligationProductObjectOfServiceIdentifier>
      </obligationRequests>
    </int:retrieveServiceObligationResponsesByServiceObligationRequests>
  </soapenv:Body>
</soapenv:Envelope>
"@

            Try {
                [XML]$entitlement = Invoke-HPIncSoapRequest -SOAPRequest $request -ErrorAction Stop
            }
            Catch {
                $P = $_
                $Global:Error.RemoveAt(0)
                throw "Failed to invoke SOAP request: $P"
            }

            Try {
                if ($entitlement) {
                    $HPAsset = $entitlement.Envelope.Body.retrieveServiceObligationResponsesByServiceObligationRequestsResponse.return
                    
                    [PSCustomObject][Ordered]@{
                        SerialNumber           = $S
                        ProductNumber          = $HPAsset.lnkProductObjectOfServiceInstance.ProductNumber
                        SalesOrderNumber       = $HPAsset.lnkServiceObligations.salesOrderNumber | where {$_}
                        ProductDescription     = $HPAsset.lnkProductObjectOfServiceInstance.productDescription
                        ProductLineDescription = $HPAsset.lnkProductObjectOfServiceInstance.productLineDescription
                        ActiveEntitlement      = $HPAsset.lnkServiceObligations.serviceObligationActiveIndicator
                        OfferDescription       = $HPAsset.lnkServiceObligationHeaderOffer | where serviceQuantity -GE 1 | Select-Object -ExpandProperty offerDescription
                        StartDate              = $HPAsset.lnkServiceObligations.serviceObligationStartDate | ForEach-Object {[DateTime]$_}
                        EndDate                = $HPAsset.lnkServiceObligations.serviceObligationEndDate | ForEach-Object {[DateTime]$_}
                    }

                    Write-Verbose "HP asset '$($HPAsset.lnkProductObjectOfServiceInstance.productDescription)' with serial number '$S'"
                }
                else {
                    Write-Warning "No HP asset information found for serial number '$S'"
                    continue
                }
            }
            Catch {
                $P = $_
                $Global:Error.RemoveAt(0)
                throw "Failed to invoke SOAP request: $P"
            }
        }
    }
}
#End Function
###############################################################

$HPWarranty = Get-HPAssetInfo -SerialNumber $SN | select -ExpandProperty EndDate








#Start of Chassis Finder

                                $Chassis = get-ciminstance -ComputerName $ComputerName -ClassName win32_systemenclosure -Namespace 'root\CIMV2' -Property ChassisTypes | select -ExpandProperty ChassisTypes

#https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-systemenclosure?redirectedfrom=MSDN
     Switch ($Chassis) {

                        '1' {
                        $Chassis = 'Other'
                        }'2' {
                        $Chassis = 'Unknown'
                        }'6' {
                        $Chassis = 'Mini Tower'
                        }'4' {
                        $Chassis = 'Low Profile Desktop'
                        }'7' {
                        $Chassis = 'Tower'
                        }'8' {
                        $Chassis = 'Portable'
                        }'9' {
                        $Chassis = 'Laptop'
                        }'10' {
                        $Chassis = 'Notebook'
                        }'13' {
                        $Chassis = 'All in One'
                        }'30' {
                        $Chassis = 'Tablet'
                        }'3' {
                        $Chassis = 'Desktop'
                        }'23' {
                        $Chassis = 'Rack Mount Chassis'
                        }'11' {
                        $Chassis = 'Hand Held'
                        }'12' {
                        $Chassis = 'Docking Station'
                        }'14' {
                        $Chassis = 'Sub Notebook'
                        }'15' { 
                        $Chassis = 'Space-Saving'
                        }'17' {
                        $Chassis = 'Main System Chassis'
                        }'18' {
                        $Chassis = 'Expansion Chassis'
                        }'19' {
                        $Chassis = 'SubChassis'
                        }'20' {
                        $Chassis = 'Bus Expansion Chassis'
                        }'21' {
                        $Chassis = 'Peripheral Chassis'
                        }'22' {
                        $Chassis = 'Storage Chassis'
                        }'24' {
                        $Chassis = 'Sealed-Case PC'
                        }'31' {
                        $Chassis = 'Convertible'
                        }'32' {
                        $Chassis = 'Detachable'
                        }

                    }
#End of Chassis Finder

                #https://en.wikipedia.org/wiki/Windows_10_version_history
                $OSVersion = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -Property version | select -ExpandProperty version 
                    Switch ($OSVersion) {
                       
                        '10.0.15063'  {
                        $OSVersion = 'Build 1703 (Fall Creators Update)'
                        }'10.0.14393' {
                        $OSVersion =  'Build 1607 (Anniversary Update)'
                        }'10.0.18362' {
                        $OSVersion = 'Build 1903'
                        }'10.0.18363' {
                        $OSVersion = 'Build 1909'
                        }'10.0.19041' {
                        $OSVersion = 'Build 2004'
                        }'10.0.19042' {
                        $OSVersion = 'Build 20H2'
                        }
                      }

                   } 

                   } else {

                   write-host "$Computername Offline" -ForegroundColor white -BackgroundColor Red
                   $computername | out-file "\\Domain\users\Username\InventoryTest_$((Get-Date).ToString("yyyyMMdd")).txt" -Append

                   }



                               $results+= [pscustomobject]@{
            
            PCName =           $Computername
            DeviceType =       $Chassis
            OU =               $ComputerOS.DistinguishedName
            OS =               $ComputerOS.OperatingSystem
            OSBuild =          $OSVersion
            AVStatus =         $ComputerAVStatus
            AVVersion =        $AppVersionAV.DisplayVersion
            AgentVerion =      $AppVersion.DisplayVersion
            VPNApplication =   $VPNService.displayname
            VPNStatus =        $VPNService.status
            PCOnline =         $PingTest
            IPAddress =        $ComputerAddress.IP
            MAC =              $ComputerAddress.MACAddress
            DHCP =             $ComputerAddress.DHCPENabled
            LoggedOnUser =     (@($FindUser) | out-string).Trim()
            Make =             $ComputerHW.Manufacturer
            Model =            $ComputerHW.Model
            CPU =              $ComputerCPU.Name
            ProcessorID =      $ComputerCPU.DeviceID
            CPUName =          $ComputerCPU.Manufacturer
            Cores =            $ComputerCPU.NumberOfCores
            CPUCount =         $ComputerHW.NumberOfProcessors
            Memory =           $ComputerHW.TotalPhysicalMemoryGB
            RAMCapacity =      $RAMCAP
            RAMManufacturer =  $RAMMANU
            RAMSpeed =         $RAMClockSpeed
            RAMSerialNumber =  $RAMSN
            HDDModel =         $ComputerDisks.model
            HDDPartitions =    $ComputerDisksALT.Count
            HDDSize =          $ComputerDisks.SizeGB
            BiosVersion =      $ComputerBios.smbiosbiosversion
            BiosReleaseDate =  (@($computerbios.releasedate) | out-string).Trim()
            SerialNumber =     $ComputerBios.serialnumber
            WarrantyEndDate =  (@($HPWarranty.Date) | out-string).Trim()
            MonitorName =      (@($ComputerMonitorNames.Names) | out-string).Trim()
            MoniotrCount =     $ComputerMonitor.Count
            Monitor =          (@($ComputerMonitorSerial.Serial) | out-string).Trim()
            MonitorBrand =     (@($ComputerMonitor.Manufacturer) | out-string).Trim()
            }
            #Fixed System.Object[] using below
            #https://stackoverflow.com/questions/41672713/exporting-system-object-to-string-in-powershell
            $Service | stop-service


                   }

            $results | export-csv "\\Domain\users\Username\InventoryTest_$((Get-Date).ToString("yyyyMMdd")).csv" -NoTypeInformation