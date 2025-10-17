<#
    .SYNOPSIS
    Generates a comprehensive HTML report of the Hyper-V environment, including host details, 
    virtual machines, snapshots, replication status, VHDX files, network adapters, 
    and virtual switches. Optionally sends the report via email.
    
    .DESCRIPTION
    The Hyper-V-Report.ps1 script is designed to automate the collection and reporting of key metrics 
    and configuration details from a Hyper-V infrastructure. It supports both standalone and 
    (future) clustered deployments and provides detailed insights into:

    Host system resources and configuration
    Virtual machine specifications and states
    Snapshot inventory and age
    Replication status and health
    VHDX file properties and fragmentation
    VM and management OS network adapter configurations
    Virtual switch topology and uplinks

    The script generates an HTML report and can send it via email using either MS Graph or MailKit, 
    depending on the configuration. 
    It relies on external modular scripts (GlobalVariables.ps1, StyleCSS.ps1, HtmlCode.ps1, Functions.ps1) 
    for customization and formatting.

    .EXAMPLE
    .\Hyper-V-Report.ps1
#>

#Region Credits
#Author: Federico Lillacci
#Github: https://github.com/tsmagnum
#endregion

# Setup all paths required for script to run

#Scripted execution
$ScriptPath = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)

#Getting date (logging,timestamps, etc.)
$today = Get-Date
$launchTime = $today.ToString('ddMMyyyy-hhmm')

#Importing required assets
. ("$($ScriptPath)\GlobalVariables.ps1")
. ("$($ScriptPath)\StyleCSS.ps1")
. ("$($ScriptPath)\HtmlCode.ps1")
. ("$($ScriptPath)\Functions.ps1")

#Setting the report filename
$reportHtmlFile = $reportHtmlDir+"\$($reportHtmlName)_"+$launchTime+".html"

#Region Hosts
#Getting Host infos
if ($clusterDeployment)
    {
        #Doing stuff for cluster env - Coming soon...
    }

else 
    {
        $vmHostsList = @()
        $vmHost = Get-VMHost 

            $hostInfos =  Get-CimInstance Win32_OperatingSystem
            $vhdPathDrive = (Get-VMHost).VirtualHardDiskPath.Split(":")[0]
            $hypervInfo = [PSCustomObject]@{     
                    Name = $vmHost.ComputerName
                    LogicalCPU = $vmHost.LogicalProcessorCount
                    RAM_GB = [System.Math]::round((($vmHost.MemoryCapacity)/1GB),2)
                    Free_RAM_GB = [System.Math]::round((($hostInfos).FreePhysicalMemory/1MB),2)
                    VHD_Volume = $vhdPathDrive+":"
                    Free_VHD_Vol_GB = [System.Math]::round(((Get-Volume $vhdPathDrive).SizeRemaining/1GB),2)
                    OsVersion = $hostInfos.Version
            }

        $vmHostsList += $hypervInfo
    }
Write-Host -ForegroundColor Cyan "###### Hyper-V Hosts infos ######"
$vmHostsList | Format-Table 

#endregion

#Region VMs
#Getting VMs detailed infos
$vmsList =@()
$vms = Get-VM 

foreach ($vm in $vms)
    {
        $vmInfo = [PSCustomObject]@{
        Name = $vm.Name
        Gen = $vm.Generation
        Version = $vm.Version
        vCPU = $vm.ProcessorCount
        RAM_Assigned = [System.Math]::round((($vm.MemoryAssigned)/1GB),2)
        RAM_Demand = [System.Math]::round((($vm.MemoryDemand)/1GB),2)
        Dyn_Memory = $vm.DynamicMemoryEnabled
        IP_Addr_NIC0 = $vm.NetworkAdapters[0].IPAddresses[0]
        Snapshots = (Get-VMSnapshot -VMName $vm.Name).Count
        State = $vm.State
        Heartbeat = $vm.Heartbeat
        Uptime = $vm.Uptime.ToString("dd\.hh\:mm")
        Replication = $vm.ReplicationState
        Creation_Time = $vm.CreationTime
        }

        $vmsList += $vmInfo
    }

Write-Host -ForegroundColor Cyan "###### Virtual Machines infos ######"
$vmsList | Format-Table

#endregion

#Region Snapshots
#Getting Snapshots
$vmSnapshotsList = @()

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

                        $vmSnapshotsList += $snapInfo
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

                        $vmSnapshotsList += $snapInfo

            }

        }
Write-Host -ForegroundColor Cyan "###### Snapshots infos ######" 
$vmSnapshotsList | Format-Table 

#endregion

#Region replication
if($replicationInfoNeeded) 
    {
        $replicationsList = @()
        #Getting the Replication status
        $replicatedVms = ($vm | Where-Object {$_.ReplicationState -ne "Disabled"})
        if (($replicatedVms).Count -gt 0)
            {
                foreach ($vm in $replicatedVms)
                {
                    $replication = Get-VmReplication -Vmname $vm.VMName |`
                        Select-Object Name,State,Health,LastReplicationTime,PrimaryServer,ReplicaServer,AuthType
                    $replicationsList += $replication
                    }
                 
                Write-Host -ForegroundColor Cyan "###### Replication infos ######"
                $replicationsList | Format-Table
            }
        
        #Creating a dummy object to correctly format the HTML report with no replications
        else
        {
            $statusMsg = "No replicated VMs found!"
            Write-Host -ForegroundColor Cyan "###### Replication infos ######"
            Write-Host -ForegroundColor Yellow $statusMsg
            $noReplicationInfo = [PSCustomObject]@{
                Replication_Infos = $statusMsg
                    }
            $replicationsList += $noReplicationInfo
        }
    }
#endregion

#Region VHDX
if ($vhdxInfoNeeded)
    {
        $vhdxList = @()

        foreach ($vm in $vms)
        {
            $vhdxs = Get-VHD -VMId $vm.VMId |`
                Select-Object ComputerName,Path,VhdFormat,VhdType,FileSize,Size,FragmentationPercentage
            
            foreach ($vhdx in $vhdxs)
            {
                $vhdxInfo = [PSCustomObject]@{
                Host = $vhdx.ComputerName
                Path = $vhdx.Path
                Format = $vhdx.VhdFormat
                Type = $vhdx.VhdType
                File_Size_GB = [System.Math]::round(($vhdx.FileSize/1GB),2)
                Size_GB = $vhdx.Size/1GB
                Frag_Perc = $vhdx.FragmentationPercentage
                }

                $vhdxList += $vhdxInfo
            }    
            
        }
        Write-Host -ForegroundColor Cyan "###### VHDX infos ######"
        $vhdxList | Format-Table
    }
#endregion    

#Region VMNetworkAdapter
if ($vmnetInfoNeeded)
    {
        $vmnetAdapterList = @()

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

                $vmnetAdapterList += $vmnetAdaptInfo
            }    
            
        }
        Write-Host -ForegroundColor Cyan "###### VM Net Adapters infos ######"
        $vmnetAdapterList | Format-Table

    }
#endregion

#Region Management OS NetworkAdapter
if ($osNetInfoNeeded)
    {
        $osNetAdapterList = @()

            $osNetadapts = Get-VMNetworkAdapter -ManagementOS |`
                Select-Object Name,MacAddress,IPAddresses,SwitchName,Status,VlanSetting
            
            foreach ($osNetadapt in $osNetadapts)
            {
                $osNetAdaptInfo = [PSCustomObject]@{
                    Name = $osNetadapt.Name
                    MAC = $osNetadapt.MacAddress
                    IP_Addr = Get-MgmtOsNicIpAddr -adapterName $osNetadapt.Name
                    vSwitch = $osNetadapt.SwitchName
                    Status = $osNetadapt.Status | Out-String
                    Vlan_Mode = $osNetadapt.VlanSetting.OperationMode
                    Vlan_Id = $osNetadapt.VlanSetting.AccessVlanId

                    
                }
                
                $osNetAdapterList += $osNetAdaptInfo
            }    
           
        Write-Host -ForegroundColor Cyan "###### Management OS Adapters infos ######"
        $osNetAdapterList | Format-Table

    }
#endregion

#Region VirtualSwitch
if ($vswitchInfoNeeded)
    {
       $vswitchList = @() 

       $vswitches = Get-VMSwitch | `
        Select-Object ComputerName,Name,EmbeddedTeamingEnabled,SwitchType,AllowManagementOS

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

                $vswitchList += $vswitchInfo
            }
                
        Write-Host -ForegroundColor Cyan "###### Virtual Switches infos ######"
        $vswitchList | Format-Table
    }
#endregion

#Creating the HTML report
if ($reportHtmlRequired)
    {
        $dataHTML =@()

        $vmhostsHTML = $preContent + $titleHtmlHosts + ($vmHostsList | ConvertTo-Html -Fragment)
        $dataHTML += $vmhostsHTML

        $vmsHTML =  $titleHtmlVms + ($vmsList | ConvertTo-Html -Fragment)
        $dataHTML += $vmsHTML

        $snapshotsHTML = $titleHtmlSnapshots + ($vmSnapshotsList | ConvertTo-Html -Fragment)
        $dataHTML += $snapshotsHTML

        if ($replicationInfoNeeded)
        {
            $replicationHTML = $titleHtmlReplication + ($replicationsList | ConvertTo-Html -Fragment)
            $dataHTML += $replicationHTML
        }

        if ($vhdxList) 
        {
            $vhdxListHTML = $titleHtmlVhdx + ($vhdxList | ConvertTo-Html -Fragment)
            $dataHTML += $vhdxListHTML
        }

        if ($vmnetInfoNeeded)
        {
            $vmnetAdapterListHTML = $titleHtmlVmnetAdapter + ($vmnetAdapterList | ConvertTo-Html -Fragment)
            $dataHTML += $vmnetAdapterListHTML
        }

        if ($osNetInfoNeeded)
        {
            $osNetAdapterListHTML = $titleHtmlOsNetAdapter + ($osNetAdapterList | ConvertTo-Html -Fragment)
            $dataHTML += $osNetAdapterListHTML
        }

        if ($vswitchInfoNeeded)
                {
            $vswitchListHTML = $titleHtmlVswitch + ($vswitchList | ConvertTo-Html -Fragment)
            $dataHTML += $vswitchListHTML
        }

        $htmlReport = ConvertTo-Html -Head $header -Title $title -PostContent $postContent -Body $dataHTML 
        $htmlReport | Out-File $reportHtmlFile 

    }

#Sending the report via email
if ($emailReport -and $reportHtmlRequired)
    {   
        switch ($emailSystem) {
            msgraph 
                { SendEmailReport-MSGraph -body (Out-String -InputObject $htmlReport) }
            
            mailkit
                { SendEmailReport-Mailkit -body (Out-String -InputObject $htmlReport) }
                
            Default {Write-Host -ForegroundColor Yellow "You must select an email system, msgraph or mailkit"}
        }
        

    }
