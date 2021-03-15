# send smtp email with OAUTH (modern authentication) to Exhange online in PowerShell 5.1
#
# based on the highly apreciated work of gscales @ https://github.com/gscales/Powershell-Scripts/blob/master/TLS-SMTP-Oauth-Mod.ps1
#
# Source: https://github.com/arveds/Send-O365MailMessage/
#
# the Powershell.Module from the powershell Gallery MSAL.PS is requiered.
# To use  
# `Install-Module MSAL.ps`  
# you need to run though this:
# `Update-Module`
# `Get-Module`
# check if PowerShellGet is higher than 1.0.0.1
# `Set-ExecutionPolicy RemoteSigned` <- needed for Module to run
# `Install-PackageProvider Nuget –force –verbose`
# `Install-Module -Name PowerShellGet -Force -AllowClobber`
# `exit` <- important
# close shell and ISE and check back again
#
# The Email Body is sent diretly plain via the SMTP DATA command
#
# The main function is called Send-O365MailMessage.
#
# you may also use the functions seperatly
# to incorporate the functions into your script simply dot-scource start the script from your ps1:
#
# . "<path ti script>\Send-O365MailMessage.ps1" (see: https://devblogs.microsoft.com/scripting/how-to-reuse-windows-powershell-functions-in-scripts/)
#
# anyone who might beautify the code and add comments is welcome
# 
# a registered app in the Azure AD Tenant is needed. make sure that the registers app has got the following properties:
#
# API Permissions: SMTP.Send
# 
# Authentication: Redirect URI: the msal.....//auth (MSAL (only) Uri must be checked
# Authentication: Allow public client flows: Yes
# Authentication: Supported Accout types: Single tenant (your directory)
#
#
# what you need for the function:
#
# ApplicationID (clientID)
#
# user consent. 
#    Either by direct user consent (use  Get-AccessTokenForSMTPSending with the Prompt switch for the first time
#    or the AAD Admin add User to the registered app as user in AAD
#
# # the syntax is similar to the Send-MailMessage function shipped with PowerShell
#
# Email Adresses can be added as 'Name somewhat <name@emaildomain.com' or 'anothername@emaildomain.com' as an array -to 'email1@domain.com','email2@domain.com'
#
# Credential: PSCredentialObject (username and password)
#

function Get-AccessTokenForSMTPSending {
    [CmdletBinding()]
    param (   
        [Parameter(Position = 1, Mandatory = $true)]
        [Object]
        $Credential,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $ClientId,
        [Parameter(Position = 3, Mandatory = $true)]
        [String]
        $RedirectURI,
        [Parameter(Position = 4, Mandatory = $false)]
        [String]
        $scopes = "https://outlook.office.com/SMTP.Send",
        [Parameter(Position = 5, Mandatory = $false)]
        [switch]
        $Prompt
    )
    Process {
        if (Get-Module -ListAvailable -Name MSAL.PS) {
            Write-Host "MSAL.PS Module exists" -ForegroundColor Green
            Import-Module MSAL.PS
           }
        else{
               Write-Host "The MSAL.PS module is required. Please install it. you may user Install-Module MSAL.PS" -ForegroundColor Red
               exit
        }
        $Domain = $Credential.UserName.Split('@')[1]
        $TenantId = (Invoke-WebRequest ("https://login.windows.net/" + $Domain + "/v2.0/.well-known/openid-configuration") | ConvertFrom-Json).token_endpoint.Split('/')[3]
 
        #if user consedt existis
        If(!$Prompt){
            $tokenRequest = Get-MsalToken -UserCredential $cred -ClientId $ClientId -TenantId $TenantId -Scopes $scopes
        }
        #For one time user consent:
        else{
            $tokenRequest = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -LoginHint $cred.UserName -Scopes $scopes -Interactive
        }
 
        $AccessToken = $tokenRequest.AccessToken

        return $AccessToken
		
    }
    
}

function Send-O365MailMessage{
    [CmdletBinding()]
    param (
        [Parameter(Position = 1, Mandatory = $true)]
        [Object]
        $Credential,
        [Parameter(Position = 2, Mandatory = $true)]
        [String]
        $ClientId,
        [Parameter(Position = 3, Mandatory = $false)]
        [String]
        $RedirectURI,
        [Parameter(Position = 4, Mandatory = $true)]
        [Object]
        $To,
        [Parameter(Position = 5, Mandatory = $true)]
        [String]
        $Subject,
        [Parameter(Position = 6, Mandatory = $true)]
        [String]
        $Body,
        [Parameter(Position = 7, Mandatory = $false)]
        [String]
        $AttachmentFileName,
        [Parameter(Mandatory = $false)]
        [Object]
        $Cc,
        [Parameter(Mandatory = $false)]
        [Object]
        $Bcc,
        [Parameter(Mandatory = $false)]
        [String]
        $From,
        [Parameter(Mandatory = $false)]
        [int]
        $Port = 587,
        [Parameter(Mandatory = $false)]
        [string]
        $SMTPServer = "smtp.office365.com",
        [Parameter(Mandatory = $false)]
        [Switch]
        $BodyAsHTML,
        [Parameter(Mandatory = $false)]
        [String]
        $Encoding = "UTF8",
        [Parameter(Mandatory = $false)]
        [String]
        $Priority = "Normal"

    )
    Process {       

        if([String]::IsNullOrEmpty($From)){
            $SendingAddress = $Credential.UserName
        }
        else{
            $SendingAddress = $From
        }

        if(![String]::IsNullOrEmpty($AttachmentFileName)){
            $attachment = New-Object System.Net.Mail.Attachment -ArgumentList $AttachmentFileName
            $mailMessage.Attachments.Add($attachment);
        }
        if([String]::IsNullOrEmpty($RedirectURI)){
            $RedirectURI = "msal" + $ClientId + "://auth" 
        } 
        # Building MailMessage Object       
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = New-Object System.Net.Mail.MailAddress($SendingAddress)
        Foreach ($item in $to) {
            If ($item.Contains("<")){
                $ADDRESSParts = $item.Split("<")
                $NamePart = $ADDRESSParts[0].Trim(" ")
                $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                $addressObj = New-Object System.Net.Mail.MailAddress($mailPart,$NamePart)
                $mailMessage.To.Add($addressObj)
            }Else
            {
            $mailPart = $item.Trim("<").Trim(">").Trim(" ")
            $addressObj = New-Object System.Net.Mail.MailAddress($mailPart)
            $mailMessage.To.Add($addressObj)
            }
        }
        if(![String]::IsNullOrEmpty($Cc)){
            Foreach ($item in $Cc) {
                If ($item.Contains("<")){
                    $ADDRESSParts = $item.Split("<")
                    $NamePart = $ADDRESSParts[0].Trim(" ")
                    $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                    $addressObj = New-Object System.Net.Mail.MailAddress($mailPart,$NamePart)
                    $mailMessage.CC.Add($addressObj)
                }Else
                {
                    $mailPart = $item.Trim("<").Trim(">").Trim(" ")
                    $addressObj = New-Object System.Net.Mail.MailAddress($mailPart)
                    $mailMessage.CC.Add($addressObj)
                }
            }
        }
        if(![String]::IsNullOrEmpty($Bcc)){
            Foreach ($item in $Bcc) {
                If ($item.Contains("<")){
                    $ADDRESSParts = $item.Split("<")
                    $NamePart = $ADDRESSParts[0].Trim(" ")
                    $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                    $addressObj = New-Object System.Net.Mail.MailAddress($mailPart,$NamePart)
                    $mailMessage.Bcc.Add($addressObj)
                }Else
                {
                    $mailPart = $item.Trim("<").Trim(">").Trim(" ")
                    $addressObj = New-Object System.Net.Mail.MailAddress($mailPart)
                    $mailMessage.Bcc.Add($addressObj)
                }
            }
        }
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body
        if($BodyAsHTML){
            $mailMessage.IsBodyHtml = $true
        }
        switch ($encoding){
            "ASCII" { $encodingObj = New-Object System.Text.ASCIIEncoding }
            "UTF8" { $encodingObj = New-Object System.Text.UTF8Encoding }
            "UniCode" { $encodingObj = New-Object System.Text.UnicodeEncoding }
            "UTF32" { $encodingObj = New-Object System.Text.UTF32Encoding }
            "UTF7" { $encodingObj = New-Object System.Text.UTF7Encoding }
        }
        #
        switch ($Priority ){
            "Normal" { $mailMessage.Priority = 0 }
            "High" { $mailMessage.Priority = 2  }
            "Low" { $mailMessage.Priority = 1  }
        }
        $mailMessage.BodyEncoding = $encodingObj
        $mailMessage.SubjectEncoding = $encodingObj
        #MailMessage Obj done

        # obj for multiple recipients needed for plain SMTP communication
        $RCPTObj = $mailMessage.to + $mailMessage.CC + $mailMessage.Bcc
        # used later in SMTP part

        # converting  MailMessage obj to Message String - somewhat magic
        $binding = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
        $MessageType = $mailMessage.GetType()
        $smtpClient = New-Object System.Net.Mail.SmtpClient
        $scType = $smtpClient.GetType()
        $booleanType = [System.Type]::GetType("System.Boolean")
        $assembly = $scType.Assembly
        $mailWriterType = $assembly.GetType("System.Net.Mail.MailWriter")
        $MemoryStream = New-Object -TypeName "System.IO.MemoryStream"
        $typeArray = ([System.Type]::GetType("System.IO.Stream"))
        $mailWriterConstructor = $mailWriterType.GetConstructor($binding ,$null, $typeArray, $null)
        [System.Array]$paramArray = ($MemoryStream)
        $mailWriter = $mailWriterConstructor.Invoke($paramArray)
        $doubleBool = $true
        $typeArray = ($mailWriter.GetType(),$booleanType,$booleanType)
        $sendMethod = $MessageType.GetMethod("Send", $binding, $null, $typeArray, $null)
        if ($null -eq $sendMethod) {
            $doubleBool = $false
            [System.Array]$typeArray = ($mailWriter.GetType(),$booleanType)
            $sendMethod = $MessageType.GetMethod("Send", $binding, $null, $typeArray, $null)
         }
        [System.Array]$typeArray = @()
        $closeMethod = $mailWriterType.GetMethod("Close", $binding, $null, $typeArray, $null)
        [System.Array]$sendParams = ($mailWriter,$true)
        if ($doubleBool) {
            [System.Array]$sendParams = ($mailWriter,$true,$true)
        }
        $sendMethod.Invoke($mailMessage,$binding,$null,$sendParams,$null)
        [System.Array]$closeParams = @()
        $MessageString = [System.Text.Encoding]::UTF8.GetString($MemoryStream.ToArray());
        $closeMethod.Invoke($mailWriter,$binding,$null,$closeParams,$null)
        [Void]$MemoryStream.Dispose()
        [Void]$mailMessage.Dispose()
        $MessageString = $MessageString.SubString($MessageString.IndexOf("MIME-Version:"))
        # Messagestring constructed
        
        # connect to SMTP server direct via network 
        $socket = new-object System.Net.Sockets.TcpClient($SMTPServer, $Port)
        $stream = $socket.GetStream()
        $streamWriter = new-object System.IO.StreamWriter($stream)
        $streamReader = new-object System.IO.StreamReader($stream)
        $streamWriter.AutoFlush = $true
        $sslStream = New-Object System.Net.Security.SslStream($stream)
        $sslStream.ReadTimeout = 30000
        $sslStream.WriteTimeout = 30000        
        $ConnectResponse = $streamReader.ReadLine();
        Write-Host($ConnectResponse)
        if(!$ConnectResponse.StartsWith("220")){
            throw "Error connecting to the SMTP Server"
        }
        $Domain = $SendingAddress.Split('@')[1]
        Write-Host(("helo " + $Domain)) -ForegroundColor Green
        $streamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $streamReader.ReadLine();
        Write-Host($ehloResponse)
        if (!$ehloResponse.StartsWith("250")){
            throw "Error in ehelo Response"
        }
        #connection established

        # starttls encryption
        Write-Host("STARTTLS") -ForegroundColor Green
        $streamWriter.WriteLine("STARTTLS");
        $startTLSResponse = $streamReader.ReadLine();
        Write-Host($startTLSResponse)
        $ccCol = New-Object System.Security.Cryptography.X509Certificates.X509CertificateCollection
        $sslStream.AuthenticateAsClient($SMTPServer,$ccCol,[System.Security.Authentication.SslProtocols]::Tls12,$false);        
        $SSLstreamReader = new-object System.IO.StreamReader($sslStream)
        $SSLstreamWriter = new-object System.IO.StreamWriter($sslStream)
        $SSLstreamWriter.AutoFlush = $true
        $SSLstreamWriter.WriteLine(("helo " + $Domain));
        $ehloResponse = $SSLstreamReader.ReadLine();
        Write-Host($ehloResponse)
        # starttls done

        # Authentication OAUTH
        $command = "AUTH XOAUTH2" 
        write-host -foregroundcolor DarkGreen $command
        $SSLstreamWriter.WriteLine($command) 
        $AuthLoginResponse = $SSLstreamReader.ReadLine()
        write-host ($AuthLoginResponse)
        $token = Get-AccessTokenForSMTPSending -Credential $Credential -ClientId $ClientId -RedirectURI $RedirectURI
        $SALSHeaderBytes = [System.Text.Encoding]::ASCII.GetBytes(("user=" + $Credential.UserName + [char]1 + "auth=Bearer " + $token + [char]1 + [char]1))
        $Base64AuthSALS = [Convert]::ToBase64String($SALSHeaderBytes)     
        #write-host -foregroundcolor DarkGreen $Base64AuthSALS
        $SSLstreamWriter.WriteLine($Base64AuthSALS)        
        $AuthResponse = $SSLstreamReader.ReadLine()
        write-host $AuthResponse
        # Auth done

        # Write Message via SMTP commands
        if($AuthResponse.StartsWith("235")){
            # Write FROM to SMTP Server
            $command = "MAIL FROM: <" + $SendingAddress + ">" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $FromResponse = $SSLstreamReader.ReadLine()
            write-host $FromResponse

            # Write all recipients to SMTP Server
            Foreach ($rcptitem in $RCPTObj){
                $command = "RCPT TO: <" + $rcptitem.Address + ">" 
                write-host -foregroundcolor DarkGreen $command
                $SSLstreamWriter.WriteLine($command) 
                $ToResponse = $SSLstreamReader.ReadLine()
                write-host $ToResponse
            }

            # Send Data
            $command = "DATA"
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            $DataResponse = $SSLstreamReader.ReadLine()
            write-host $DataResponse
            write-host -foregroundcolor DarkGreen $MessageString
            $SSLstreamWriter.WriteLine($MessageString) 
            $SSLstreamWriter.WriteLine(".") 
            $DataResponse = $SSLstreamReader.ReadLine()
            write-host $DataResponse
            $command = "QUIT" 
            write-host -foregroundcolor DarkGreen $command
            $SSLstreamWriter.WriteLine($command) 
            # ## Close the streams 
            $stream.Close() 
            $sslStream.Close()
            Write-Host("Done")
        }  
    }
}
