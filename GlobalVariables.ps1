#Script Info - do not modify the following line
$scriptVersion = "1.0"

#Cluster Environment - Coming soon...
$clusterDeployment = $false
$vmHosts = "hv1", "hv2"

#Reporting section
$reportHtmlRequired = $true #Set to $true to generate an HTML report
$reportHtmlName = "Hyper-V_Status_Report" #Report file name without extension, e.g., "MyReport"
$reportHtmlDir = "C:\Temp" #Directory to save the report (no trailing slash), e.g., "C:\Temp"

#Information to be included in the report
$replicationInfoNeeded = $true # Set to true to include replication details
$vhdxInfoNeeded = $true #Set to true to include detailed VHDX information
$vmnetInfoNeeded = $true #Set to true to include VM network adapter details
$osNetInfoNeeded = $true #Set to true to include Management OS network adapter details
$vswitchInfoNeeded = $true #Set to true to include virtual switch details

#Email Configuration
$emailReport = $true #Set to true to send the report via email
$emailSystem = "mailkit" #Choose the email system: "msgraph" or "mailkit"

$reportSender = "mySender@mydomain.com" #Sender email address (use quotes)
$reportRecipient = "myRecipient@mydomain.com" #Recipient email address (use quotes)
$ccrecipient = $null #CC email address (use quotes); leave as $null if not used
$subject = "Hyper-V Status Report" #Email subject line

#MS Graph Email specific configuration
$type = "HTML" #Choose between "HTML" or "TXT"
$save = $false #Set to true to save the email in the Sent Items folder

#MailKit Email specific configuration
$smtpServer = "mysmtpserver.domain.com"
$smtpServerPort = 587
$smtpAuthRequired = $true
#If SMTP authentication is required, set a username and password below.
#DO NOT USE A SENSITIVE OR PRIVILEGED ACCOUNT HERE!!!
$smtpServerUser = "smtpserver.user"
$smtpServerPwd = "mySecretPwd"

