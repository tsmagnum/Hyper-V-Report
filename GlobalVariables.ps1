#Script Info - please do not modify the following line
$scriptVersion = "2.0"

#Cluster Environment
$clusterDeployment = $false #Set to $true for a clustered environment

#Reporting section
$reportHtmlRequired = $true #Set to $true to generate an HTML report
$reportStyle = "prodark" #Choose between "minimal", "pro", "prodark" or "colorful"
$reportHtmlName = "Hyper-V_Status_Report" #Report file name without extension, e.g., "MyReport"
$reportHtmlDir = "C:\Temp" #Directory to save the report (no trailing slash), e.g., "C:\Temp"

#Information to be included in the report
$csvHealthInfoNeeded = $true #Set to true to include CSV health details
$replicationInfoNeeded = $true # Set to true to include replication details
$vhdxInfoNeeded = $true #Set to true to include detailed VHDX information
$vmnetInfoNeeded = $true #Set to true to include VM network adapter details
$osNetInfoNeeded = $true #Set to true to include Management OS network adapter details
$vswitchInfoNeeded = $true #Set to true to include virtual switch details
$clusterConfigInfoNeeded = $true #Set to true to include cluster config details
$clusterNetworksInfoNeeded = $true #Set to true to include cluster networks details

#Email Configuration
$emailReport = $true #Set to true to send the report via email
$emailSystem = "mailkit" #Choose the email system: "msgraph" or "mailkit"

$reportSender = "mySender@domain.com" #Sender email address (use quotes)
$reportRecipient = "myrecipient@domain.com" #Recipient email address (use quotes)
$ccrecipient = $null #CC email address (use quotes); leave as $null if not used
$subject = "Hyper-V Status Report" #Email subject line

#MS Graph Email specific configuration
$type = "HTML" #Choose between "HTML" or "TXT"
$save = $false #Set to true to save the email in the Sent Items folder

#MailKit-specific email configuration
$smtpServer = "mySmtp.server.com"
$smtpServerPort = 587
$smtpAuthRequired = $true
#It is recommended to use an encrypted XML file for SMTP credentials.
#Run the Save-SafeCreds.ps1 script to store your credentials in an encrypted XML file.
#Save the encrypted XML file in the script directory.
#Set the following variable to $true and enter the path to the XML file. 
#Please note: only the user encrypting the creds will be able to decrypt them!
$encryptedSMTPCreds = $true #set to true to use the encrypted XML file for the creds.
$encryptedSMTPCredsFileName = "EncryptedCreds.xml" #name of the encrypted creds file.
#If you prefer to store the credentials in plain text, set the username and password below.
#and set $encryptedSMTPCreds to $false
#
#DO NOT USE A SENSITIVE OR PRIVILEGED ACCOUNT HERE!!!
#This poses a security risk — use these credentials only for testing purposes.
$smtpServerUser = "smtpserver.user"
$smtpServerPwd = "mySecretPwd"

