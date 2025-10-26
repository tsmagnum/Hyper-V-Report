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
. ("$($ScriptPath)\HtmlCode.ps1")
. ("$($ScriptPath)\Functions.ps1")

#Setting the report style
switch ($reportStyle) {
    minimal 
        { $styleSheet = "StyleCSS-Minimal.ps1" }
    pro 
        { $styleSheet = "StyleCSS-Pro.ps1" }
    prodark 
        { $styleSheet = "StyleCSS-ProDark.ps1" }    
    colorful 
        { $styleSheet = "StyleCSS-Colorful.ps1" }
    Default 
        { $styleSheet = "StyleCSS-Professional.ps1" }
}

. ("$($ScriptPath)\Style\$styleSheet")

#Setting the report filename
$reportHtmlFile = $reportHtmlDir+"\$($reportHtmlName)_"+$launchTime+".html"

#Setting the encrypted creds file path
if ($emailReport)
{
    $encryptedSMTPCredsFile = "$($ScriptPath)\$encryptedSMTPCredsFileName"
}

#Region Hosts
$vmHosts = [System.Collections.ArrayList]@()

#Getting Host infos
if ($clusterDeployment)
    {
        $clusterNodes = Get-ClusterNode -Cluster .
        
        foreach ($clusterNode in $clusterNodes)
        {      
            $vmHost = Get-VMHost -ComputerName $clusterNode.NodeName

            [void]$vmHosts.Add($vmHost)
        }
        
        $vmHostsList = Get-VmHostInfo -vmHosts $vmHosts
    }

#Non clustered deployments
else 
    { 
        $vmHost = Get-VMHost 

        $vmHostsList = Get-VmHostInfo -vmHosts $vmHost
    }

Write-Host -ForegroundColor Cyan "###### Hyper-V Hosts infos ######"
$vmHostsList | Format-Table 

#endregion

#Region CSV Health - Only for clustered environments
if (($clusterDeployment) -and ($csvHealthInfoNeeded))
    {
        $csvHealthList = Get-CsvHealth -clusterNodes $clusterNodes

        Write-Host -ForegroundColor Cyan "###### CSV Health Info ######"
        $csvHealthList | Format-Table
    }
#endregion

#Region CSV Space Utilization - Only for clustered environments
if ($clusterDeployment)
    {
        $csvSpaceList = Get-CsvSpaceUtilization

        Write-Host -ForegroundColor Cyan "###### CSV Space Utilization ######"
        $csvSpaceList | Format-Table
    }
#endregion

#Region VMs
#Getting VMs detailed infos
$vmsList =@()

if ($clusterDeployment)
    {
        $vms = foreach ($clusterNode in $clusterNodes)
            {
                Get-VM -ComputerName $clusterNode.NodeName
            }
    }

#Non clustered deployments
else
    {
        $vms = Get-VM 
    }

foreach ($vm in $vms)
    {
        $vmInfo = Get-VmInfo -vm $vm

        $vmsList += $vmInfo
    }

Write-Host -ForegroundColor Cyan "###### Virtual Machines infos ######"
$vmsList | Format-Table

#endregion

#Region Snapshots
#Getting Snapshots
$vmSnapshotsList = Get-SnapshotInfo -vms $vms

Write-Host -ForegroundColor Cyan "###### Snapshots infos ######" 
$vmSnapshotsList | Format-Table 
#endregion

#Region replication
if($replicationInfoNeeded) 
    {
        $replicationsList = Get-ReplicationInfo -vms $vms

        Write-Host -ForegroundColor Cyan "###### Replication infos ######"
        $replicationsList | Format-Table
        
    }
#endregion

#Region VHDX
if ($vhdxInfoNeeded)
    {
        $vhdxList = Get-VHDXInfo -vms $vms
        
        Write-Host -ForegroundColor Cyan "###### VHDX infos ######"
        $vhdxList | Format-Table
    }
#endregion    

#Region VMNetworkAdapter
if ($vmnetInfoNeeded)
    {
        $vmnetAdapterList = Get-VmnetInfo -vms $vms

        Write-Host -ForegroundColor Cyan "###### VM Net Adapters infos ######"
        $vmnetAdapterList | Format-Table

    }
#endregion

#Region Management OS NetworkAdapter
if ($osNetInfoNeeded)
    {
        if ($clusterDeployment)
        { 
            $osNetAdapterList = Get-OsNetAdapterInfo -vmHost $vmHosts
        }
        
        else
        {
            $osNetAdapterList = Get-OsNetAdapterInfo -vmHost $vmHost
        }            
           
        Write-Host -ForegroundColor Cyan "###### Management OS Adapters infos ######"
        $osNetAdapterList | Format-Table

    }
#endregion

#Region VirtualSwitch
if ($vswitchInfoNeeded)
    {
      
        if ($clusterDeployment)
        {
            $vswitchesList = Get-VswitchInfo -vmHost $vmhosts
        }
        
        else
        {
            $vswitchesList = Get-VswitchInfo -vmHost $vmhost
        }
                
        Write-Host -ForegroundColor Cyan "###### Virtual Switches infos ######"
        $vswitchesList | Format-Table
    }
#endregion

#Region Cluster Configuration Info
if ($clusterDeployment -and $clusterConfigInfoNeeded)
{
    $clusterConfigInfoList = Get-ClusterConfigInfo
    
    Write-Host -ForegroundColor Cyan "###### Cluster config infos ######"
    $clusterConfigInfoList | Format-Table
}
#endregion

#Region Cluster Networks Info
if ($clusterDeployment -and $clusterNetworksInfoNeeded)
{
    $clusterNetworksList = Get-ClusterNetworksInfo

    Write-Host -ForegroundColor Cyan "###### Cluster networks infos ######"
    $clusterNetworksList | Format-Table
}
#endregion

############### Report and Email ###############

#Creating the HTML report
if ($reportHtmlRequired)
    {
        $dataHTML = [System.Collections.ArrayList]@()

        $vmhostsHTML = $preContent + $titleHtmlHosts + ($vmHostsList | ConvertTo-Html -Fragment)
        [void]$dataHTML.Add($vmhostsHTML)

        if (($clusterDeployment) -and ($csvHealthInfoNeeded))
        {
            $csvHealthHTML = $titleHtmlcsvHealth + ($csvHealthList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($csvHealthHTML)
        }

        if ($clusterDeployment)
        {
            $csvSpaceHTML = $titleHtmlcsvSpace + ($csvSpaceList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($csvSpaceHTML)
        }

        $vmsHTML =  $titleHtmlVms + ($vmsList | ConvertTo-Html -Fragment)
        [void]$dataHTML.Add($vmsHTML)

        $snapshotsHTML = $titleHtmlSnapshots + ($vmSnapshotsList | ConvertTo-Html -Fragment)
        [void]$dataHTML.Add($snapshotsHTML)

        if ($replicationInfoNeeded)
        {
            $replicationHTML = $titleHtmlReplication + ($replicationsList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($replicationHTML)
        }

        if ($vhdxList) 
        {
            $vhdxListHTML = $titleHtmlVhdx + ($vhdxList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($vhdxListHTML)
        }

        if ($vmnetInfoNeeded)
        {
            $vmnetAdapterListHTML = $titleHtmlVmnetAdapter + ($vmnetAdapterList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($vmnetAdapterListHTML)
        }

        if ($osNetInfoNeeded)
        {
            $osNetAdapterListHTML = $titleHtmlOsNetAdapter + ($osNetAdapterList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($osNetAdapterListHTML)
        }

        if ($vswitchInfoNeeded)
                {
            $vswitchesListHTML = $titleHtmlVswitch + ($vswitchesList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($vswitchesListHTML)
        }

        if ($clusterDeployment -and $clusterConfigInfoNeeded)
        {
            $clusterConfigInfoListHTML = $titleHtmlClusterConfig + ($clusterConfigInfoList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($clusterConfigInfoListHTML)
        }

                if ($clusterDeployment -and $clusterNetworksInfoNeeded)
        {
            $clusterNetworksListHTML = $titleHtmlClusterNetworks + ($clusterNetworksList | ConvertTo-Html -Fragment)
            [void]$dataHTML.Add($clusterNetworksListHTML)
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

