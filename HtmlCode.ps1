$preContent = "<h1>Hyper-V Status Report</h1>"
$postContent = "<p>Creation Date: $($today | Out-String) - 
    <a href='https://github.com/tsmagnum/Hyper-V_Report' target='_blank'>Hyper-V-Report ($scriptVersion) by F. Lillacci</a><p>"

$title = "Hyper-V Status Report"
$breakHtml = "</br>"

$titleHtmlHosts = "<h3>Hyper-V Server</h3>"
$titleHtmlcsvHealth = "<h3>CSV Health</h3>"
$titleHtmlcsvSpace = "<h3>CSV Space Utilization</h3>"
$titleHtmlVms = "<h3>Virtual Machines</h3>"
$titleHtmlSnapshots = "<h3>Snapshots</h3>"
$titleHtmlReplication = "<h3>Replication</h3>"
$titleHtmlVhdx = "<h3>VHDX Disks</h3>"
$titleHtmlVmnetAdapter = "<h3>VM Network Adatpers</h3>"
$titleHtmlOsNetAdapter = "<h3>Management OS Network Adatpers</h3>"
$titleHtmlVswitch = "<h3>Virtual Switches</h3>"
$titleHtmlClusterConfig = "<h3>Cluster Config</h3>"
$titleHtmlClusterNetworks = "<h3>Cluster Networks</h3>"