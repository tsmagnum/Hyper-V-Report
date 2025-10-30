#DO NOT MODIFY
function Get-VmHostInfo
{
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vmHosts
        )

        $vmHostsList = [System.Collections.ArrayList]@()

        foreach ($vmHost in $vmHosts)
        {

            $hostInfos =  Get-CimInstance Win32_OperatingSystem -ComputerName $vmHost.ComputerName
            $vhdPathDrive = (Get-VMHost).VirtualHardDiskPath.Split(":")[0]
            $hypervInfo = [PSCustomObject]@{     
                    Name = $vmHost.ComputerName
                    LogicalCPU = $vmHost.LogicalProcessorCount
                    RAM_GB = [System.Math]::round((($vmHost.MemoryCapacity)/1GB),2)
                    Free_RAM_GB = [System.Math]::round((($hostInfos).FreePhysicalMemory/1MB),2)
                    VHD_Volume = $vhdPathDrive+":"
                    Free_VHD_Vol_GB = [System.Math]::round(((Get-Volume $vhdPathDrive -CimSession $vmHost.ComputerName).SizeRemaining/1GB),2)
                    LiveMig = $vmHost.VirtualMachineMigrationEnabled
                    Last_Boot = $hostInfos.LastBootUpTime.ToString('dd/MM/yy HH:mm')
                    OsBuild = Get-OsBuildLevel -vmHost $vmHost.ComputerName
                    }

            [void]$vmHostsList.Add($hypervInfo)
        }

        return $vmHostsList
}

function Get-VHDXInfo {
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vms
        ) 

        $vhdxList = [System.Collections.ArrayList]@()

        foreach ($vm in $vms)
        {
            $vhdxs = Get-VHD -VMId $vm.VMId |`
                Select-Object ComputerName,Path,VhdFormat,VhdType,FileSize,Size,FragmentationPercentage
            
            foreach ($vhdx in $vhdxs)
            {
                $vhdxInfo = [PSCustomObject]@{
                Host = $vhdx.ComputerName
                VM = $vm.VMName
                Path = $vhdx.Path
                Format = $vhdx.VhdFormat
                Type = $vhdx.VhdType
                File_Size_GB = [System.Math]::round(($vhdx.FileSize/1GB),2)
                Size_GB = $vhdx.Size/1GB
                Frag_Perc = $vhdx.FragmentationPercentage
                }

                [void]$vhdxList.Add($vhdxInfo)
            }    
            
        }

        return $vhdxList
    
}

function Get-VmnetInfo {
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vms
        ) 
    
    $vmnetAdapterList = [System.Collections.ArrayList]@()

    foreach ($vm in $vms)
        {
            $vmnetadapts = Get-VMNetworkAdapter -vm $vm |`
                Select-Object MacAddress,Connected,VMName,IsSynthetic,IPAddresses,SwitchName,Status,VlanSetting
            
    foreach ($vmnetadapt in $vmnetadapts)
            {
                $vmnetAdaptInfo = [PSCustomObject]@{
                    VM = $vmnetadapt.VMName
                    MAC = $vmnetadapt.MacAddress
                    IP_Addr = $vmnetadapt.IPAddresses | Out-String
                    Connected = $vmnetadapt.Connected
                    vSwitch = $vmnetadapt.SwitchName
                    #Status = $vmnetadapt.Status.
                    Vlan_Mode = $vmnetadapt.VlanSetting.OperationMode
                    Vlan_Id = $vmnetadapt.VlanSetting.AccessVlanId

                }

        [void]$vmnetAdapterList.Add($vmnetAdaptInfo)
            }    
            
        }

    return $vmnetAdapterList    

}
function Get-CsvHealth
{
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $clusterNodes
        ) 

            $csvHealthList = [System.Collections.ArrayList]@()
        
            foreach ($clusterNode in $clusterNodes)
            {
                $csvHealth = Get-ClusterSharedVolumeState -Node $clusterNode.NodeName |`
                Select-Object Node,VolumeFriendlyName,Name,StateInfo

                foreach ($csv in $csvHealth)
                    {
                        $csvHealthVol = [PSCustomObject]@{
                            Node = $csv.Node
                            Volume = $csv.VolumeFriendlyName
                            Disk = $csv.Name
                            State = $csv.StateInfo
                        }
                    
                    [void]$csvHealthList.Add($csvHealthVol)
                    
                    }
            }

            return $csvHealthList
}

function Get-CsvSpaceUtilization {

    [CmdletBinding()]

        $csvSpaceList = [System.Collections.ArrayList]@()

        $csvVolumes = Get-Volume | Where-Object {$_.FileSystem -match "CSVFS"} |`
        Select-Object FileSystemLabel,HealthStatus,Size,SizeRemaining

        foreach ($csvVolume in $csvVolumes)
            {
                $csvInfo = [PSCustomObject]@{
                    Volume = $csvVolume.FileSystemLabel
                    Health = $csvVolume.HealthStatus
                    Size_GB = [System.Math]::round(($csvVolume.Size/1GB),2)
                    Free_GB = [System.Math]::round(($csvVolume.SizeRemaining/1GB),2)
                    Used_Perc = [System.Math]::round(((($csvVolume.Size - $csvVolume.SizeRemaining)/$csvVolume.Size)*100),2)
                }

                [void]$csvSpaceList.Add($csvInfo)
            }
        
        return $csvSpaceList
    
}

function Get-VmInfo
{
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vm
        ) 

    $vmInfo = [PSCustomObject]@{
        Name = $vm.Name
        Host = $vm.ComputerName
        Gen = $vm.Generation
        Version = $vm.Version
        vCPU = $vm.ProcessorCount
        RAM_Assigned = [System.Math]::round((($vm.MemoryAssigned)/1GB),2)
        RAM_Demand = [System.Math]::round((($vm.MemoryDemand)/1GB),2)
        Dyn_Memory = $vm.DynamicMemoryEnabled
        IP_Addr_NIC0 = $vm.NetworkAdapters[0].IPAddresses[0]
        Snapshots = (Get-VMSnapshot -VMName $vm.Name).Count
        Clustered = $vm.IsClustered
        State = $vm.State
        Heartbeat = $vm.Heartbeat
        Uptime = $vm.Uptime.ToString("dd\.hh\:mm")
        Replication = $vm.ReplicationState
        Creation_Time = $vm.CreationTime
        }

    return $vmInfo
}

function Get-SnapshotInfo {
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vms
        ) 
        
    $vmSnapshotsList = [System.Collections.ArrayList]@()

    foreach ($vm in $vms)
        {    
            $foundSnapshots = Get-VMSnapshot -VMName $vm.name
            
            if (($foundSnapshots).Count -gt 0)
                {
                    foreach ($foundSnapshot in $foundSnapshots)
                        {
                            $snapInfo = [PSCustomObject]@{
                            VM = $foundSnapshot.VMName
                            Name = $foundSnapshot.Name
                            Creation_Time = $foundSnapshot.CreationTime
                            Age_Days = $today.Subtract($foundSnapshot.CreationTime).Days
                            Parent_Snap = $foundSnapshot.ParentSnapshotName
                            }   

                            [void]$vmSnapshotsList.Add($snapInfo)
                        }
                }

            else 
                {
                            $snapInfo = [PSCustomObject]@{
                            VM = $vm.VMName
                            Name = "No snapshots found"
                            Creation_Time = "N/A"
                            Age_Days = "N/A"
                            Parent_Snap = "No snapshots found"
                            }   

                            [void]$vmSnapshotsList.Add($snapInfo)

                }


            }
    return $vmSnapshotsList        
}

function Get-ReplicationInfo {

    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vms
        ) 

    $replicationsList = [System.Collections.ArrayList]@()
        #Getting the Replication status
        $replicatedVms = ($vms | Where-Object {$_.ReplicationState -ne "Disabled"} | Select-Object -Unique)
        if (($replicatedVms).Count -gt 0)
            {
                foreach ($vm in $replicatedVms)
                {
                    $replication = Get-VmReplication -Vmname $vm.VMName |`
                        Select-Object Name,State,Health,LastReplicationTime,PrimaryServer,ReplicaServer,AuthType
                    [void]$replicationsList.Add($replication)
                    }
            }
        
        #Creating a dummy object to correctly format the HTML report with no replications
        else
        {
            $statusMsg = "No replicated VMs found!"
            $noReplicationInfo = [PSCustomObject]@{
                Replication_Infos = $statusMsg
                    }
            [void]$replicationsList.Add($noReplicationInfo)
        }

        return $replicationsList
}

function Get-OsNetAdapterInfo
{
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vmHosts
        ) 
        
         $osNetadaptsList = @()

        foreach ($vmHost in $vmHosts)
        {
            $osNetadaptsArray = @()

            $osNetadapts = Get-VMNetworkAdapter -ManagementOS -ComputerName $vmHost.Name |`
                    Select-Object ComputerName,Name,MacAddress,IPAddresses,SwitchName,Status,VlanSetting
                
            foreach ($osNetadapt in $osNetadapts)
            {
                $osNetAdaptInfo = [PSCustomObject]@{
                    Host = $osNetadapt.ComputerName
                    Name = $osNetadapt.Name
                    MAC = $osNetadapt.MacAddress
                    IP_Addr = Get-MgmtOsNicIpAddr -adapterName $osNetadapt.Name -vmHost $vmHost
                    vSwitch = $osNetadapt.SwitchName
                    Status = $osNetadapt.Status | Out-String
                    Vlan_Mode = $osNetadapt.VlanSetting.OperationMode
                    Vlan_Id = $osNetadapt.VlanSetting.AccessVlanId
                    }

                $osNetadaptsArray += $osNetAdaptInfo
            }

            $osNetadaptsList += $osNetadaptsArray

        }

    return $osNetadaptsList
}

function Get-VswitchInfo
{
       [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vmHosts
        ) 

        $vswitchesList = @() 

        foreach ($vmHost in $vmHosts)
        {
            $vswitches = Get-VMSwitch -ComputerName $vmHost.Name | `
                Select-Object ComputerName,Name,EmbeddedTeamingEnabled,SwitchType,AllowManagementOS

            $vswitchesArray = @()

            foreach ($vswitch in $vswitches)
                {
                    $vswitchInfo = [PSCustomObject]@{
                        Host = $vswitch.ComputerName
                        Virtual_Switch = $vswitch.Name
                        SET = $vswitch.EmbeddedTeamingEnabled
                        Uplinks = Get-VswitchMember -vswitch $vswitch.Name
                        Type = $vswitch.SwitchType
                        Mgmt_OS_Allowed = $vswitch.AllowManagementOS
                        }
                    
                    $vswitchesArray += $vswitchInfo
                }

            $vswitchesList += $vswitchesArray
        }

        return $vswitchesList
}

function Get-MgmtOsNicIpAddr
{
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $adapterName,
            [Parameter(Mandatory = $true)] $vmHost
        )  

        $ipAddr = (Get-NetIPAddress -CimSession $vmHost.Name |`
            Where-Object {$_.InterfaceAlias -like "*($($adapterName))" -and $_.AddressFamily -eq 'IPv4'}).IPAddress

        return $ipAddr
}

function Get-VswitchMember
{
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vswitch
        )

        $targetSwitch = Get-VMSwitch -Name $vswitch

        $vswitchMembers = ($targetSwitch.NetAdapterInterfaceDescriptions)

        $physNics = [System.Collections.ArrayList]@()

        foreach ($vswitchMember in $vswitchMembers)
            {
                $physNic = (Get-Netadapter -InterfaceDescription $vswitchMember).Name
                [void]$physNics.Add($physNic)
            }

        $vswitchPhysNics = $physNics | Out-String

        return $vswitchPhysNics

}

function Get-ClusterConfigInfo {
    
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false)] $cluster = "."
        )

    $cluster = Get-Cluster -Name $cluster | Select-Object Name,Domain

    $clusterConfigInfoList = [PSCustomObject]@{

        Cluster_Name = $cluster.Name
        Cluster_Domain = $cluster.Domain
    }

    return $clusterConfigInfoList
    
}

function Get-ClusterNetworksInfo {
    
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $false)] $cluster = "."
        )

    $clusterNets = Get-ClusterNetwork -Cluster $cluster | Select-Object Name,Role,State,Address,Autometric,Metric

    $clusterNetworksList = [System.Collections.ArrayList]@()

    foreach ($clusterNet in $clusterNets)
    {   
        $clusterNetInfo = [PSCustomObject]@{
            Name = $clusterNet.Name
            Role = $clusterNet.Role
            State = $clusterNet.State
            Address = $clusterNet.Address
            Autometric = $clusterNet.Autometric
            Metric = $clusterNet.Metric 
        }

        [void]$clusterNetworksList.Add($clusterNetInfo)
    }

    return $clusterNetworksList
    
}

function Get-OsBuildLevel
{
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $vmHost
        )

        $code = {
                
                Try
                {
                    $osBuildLevel = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name LCUver -Verbose).LCUVer 
                    return $osBuildLevel
                }
                Catch
                {
                    $osBuildLevel = "N/A"
                    return $osBuildLevel
                }
        }

        if ($vmHost -eq $env:COMPUTERNAME)
        {
            $osBuildLevel = Invoke-Command -ScriptBlock $code
        }

        else 
        {
            $osBuildLevel = Invoke-Command -ComputerName $vmHost -ScriptBlock $code -Verbose
        }

    

    return $osBuildLevel
}

function Import-SafeCreds {
    [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $encryptedSMTPCredsFile
        )  

        $credentials = Import-Clixml -Path $encryptedSMTPCredsFile
        return $credentials
    
}

function SendEmailReport-MSGraph
{
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $body
        )        
                #Checking if required module is present; if not, install it
                if (!(Get-Module -Name Microsoft.Graph -ListAvailable))
                {
                        Write-Host -ForegroundColor Yellow "MS Graph module missing, installing..."
                        Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
                        Import-Module -Name Microsoft.Graph
                }

                #Connect interactively to Microsoft Graph with required permissions
                Connect-MgGraph -NoWelcome -Scopes 'Mail.Send', 'Mail.Send.Shared'

                if ($ccrecipient)
                {
                        $params = @{
                        Message         = @{
                                Subject       = $subject
                                Body          = @{
                                ContentType = $type
                                Content     = $body
                                }
                                ToRecipients  = @(
                                @{
                                        EmailAddress = @{
                                        Address = $reportRecipient
                                        }
                                }
                                )
                                CcRecipients  = @(
                                @{
                                        EmailAddress = @{
                                        Address = $ccrecipient
                                        }
                                }
                                )
                        }
                        SaveToSentItems = $save
                        }
                }

                else 
                {
                        $params = @{
                        Message         = @{
                                Subject       = $subject
                                Body          = @{
                                ContentType = $type
                                Content     = $body
                                }
                                ToRecipients  = @(
                                @{
                                        EmailAddress = @{
                                        Address = $reportRecipient
                                        }
                                }
                                )
                        }
                        SaveToSentItems = $save
                        }
                }

                # Send message
                Send-MgUserMail -UserId $reportSender -BodyParameter $params
                
}

function SendEmailReport-Mailkit
{
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $body
        )   
        
        #Checking if required module is present; if not, install it
        if (!(Get-Module -ListAvailable -Name "Send-MailKitMessage")) 
        { 
                Write-Host -ForegroundColor Yellow "Send-MailKitMessage module missing, installing..."
                Install-Module -Name "Send-MailKitMessage" -Scope CurrentUser -Force           
        }
		
		Import-Module -Name "Send-MailKitMessage"
        
        $UseSecureConnectionIfAvailable = $true

        if ($encryptedSMTPCreds)
            {
                $Credentials = Import-SafeCreds -encryptedSMTPCredsFile $encryptedSMTPCredsFile
            }

        else 
            {
                $credentials = `
                        [System.Management.Automation.PSCredential]::new($smtpServerUser, `
                                (ConvertTo-SecureString -String $smtpServerPwd -AsPlainText -Force))
            }

        $from = [MimeKit.MailboxAddress]$reportSender
        
        $recipientList = [MimeKit.InternetAddressList]::new()
        $recipientList.Add([MimeKit.InternetAddress]$reportRecipient)
        
        if ($ccrecipient)
        {
                $ccList = [MimeKit.InternetAddressList]::new();
                $ccList.Add([MimeKit.InternetAddress]$ccrecipient);      
        }

        $Parameters = @{
                "UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable    
                "Credential" = $credentials
                "SMTPServer" = $smtpServer
                "Port" = $smtpServerPort
                "From" = $from
                "RecipientList" = $recipientList
                "CCList" = $ccList
                "Subject" = $subject
                "HTMLBody" = $body
                }    

        Send-MailKitMessage @Parameters                 
}