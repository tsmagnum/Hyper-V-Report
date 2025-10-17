#DO NOT MODIFY
function Get-MgmtOsNicIpAddr
{
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $true)] $adapterName
        )  

        $ipAddr = (Get-NetIPAddress | Where-Object {$_.InterfaceAlias -like "*$($adapterName)*" -and $_.AddressFamily -eq 'IPv4'}).IPAddress

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

        $physNics = @()

        foreach ($vswitchMember in $vswitchMembers)
            {
                $physNic = (Get-Netadapter -InterfaceDescription $vswitchMember).Name
                $physNics += $physNic
            }

        $vswitchPhysNics = $physNics | Out-String

        return $vswitchPhysNics

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
                Import-Module -Name "Send-MailKitMessage"
        }
        
        $UseSecureConnectionIfAvailable = $true

        $credential = `
                [System.Management.Automation.PSCredential]::new($smtpServerUser, `
                        (ConvertTo-SecureString -String $smtpServerPwd -AsPlainText -Force))

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
                "Credential" = $credential
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