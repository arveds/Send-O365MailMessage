# send smtp email with OAUTH (modern authentication) to Exhange online in PowerShell 5.1
#
# based on the highly apreciated work of gscales @ https://github.com/gscales/Powershell-Scripts/blob/master/TLS-SMTP-Oauth-Mod.ps1
#
# dual-stack email sending, supports SMTP with OAUTH and GraphAPI with OAUTH
#
# Default is send via GraphAPI, with the -SendWithSMTP switch you can change to SMTP
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

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-AccessTokenForMailSending {
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
            $tokenRequest = Get-MsalToken -UserCredential $Credential -ClientId $ClientId -TenantId $TenantId -Scopes $scopes
        }
        #For one time user consent:
        else{
            $tokenRequest = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -LoginHint $cred.UserName -Scopes $scopes -Interactive
        }
 
        $AccessToken = $tokenRequest.AccessToken

        return $AccessToken
		
    }
    
}


function Get-ContentTypeFromFileName{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
            [string]
            $FileName
          )
    $mimeMappings = @{
                '.323' = 'text/h323'
                '.aaf' = 'application/octet-stream'
                '.aca' = 'application/octet-stream'
                '.accdb' = 'application/msaccess'
                '.accde' = 'application/msaccess'
                '.accdt' = 'application/msaccess'
                '.acx' = 'application/internet-property-stream'
                '.afm' = 'application/octet-stream'
                '.ai' = 'application/postscript'
                '.aif' = 'audio/x-aiff'
                '.aifc' = 'audio/aiff'
                '.aiff' = 'audio/aiff'
                '.application' = 'application/x-ms-application'
                '.art' = 'image/x-jg'
                '.asd' = 'application/octet-stream'
                '.asf' = 'video/x-ms-asf'
                '.asi' = 'application/octet-stream'
                '.asm' = 'text/plain'
                '.asr' = 'video/x-ms-asf'
                '.asx' = 'video/x-ms-asf'
                '.atom' = 'application/atom+xml'
                '.au' = 'audio/basic'
                '.avi' = 'video/x-msvideo'
                '.axs' = 'application/olescript'
                '.bas' = 'text/plain'
                '.bcpio' = 'application/x-bcpio'
                '.bin' = 'application/octet-stream'
                '.bmp' = 'image/bmp'
                '.c' = 'text/plain'
                '.cab' = 'application/octet-stream'
                '.calx' = 'application/vnd.ms-office.calx'
                '.cat' = 'application/vnd.ms-pki.seccat'
                '.cdf' = 'application/x-cdf'
                '.chm' = 'application/octet-stream'
                '.class' = 'application/x-java-applet'
                '.clp' = 'application/x-msclip'
                '.cmx' = 'image/x-cmx'
                '.cnf' = 'text/plain'
                '.cod' = 'image/cis-cod'
                '.cpio' = 'application/x-cpio'
                '.cpp' = 'text/plain'
                '.crd' = 'application/x-mscardfile'
                '.crl' = 'application/pkix-crl'
                '.crt' = 'application/x-x509-ca-cert'
                '.csh' = 'application/x-csh'
                '.css' = 'text/css'
                '.csv' = 'application/octet-stream'
                '.cur' = 'application/octet-stream'
                '.dcr' = 'application/x-director'
                '.deploy' = 'application/octet-stream'
                '.der' = 'application/x-x509-ca-cert'
                '.dib' = 'image/bmp'
                '.dir' = 'application/x-director'
                '.disco' = 'text/xml'
                '.dll' = 'application/x-msdownload'
                '.dll.config' = 'text/xml'
                '.dlm' = 'text/dlm'
                '.doc' = 'application/msword'
                '.docm' = 'application/vnd.ms-word.document.macroEnabled.12'
                '.docx' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                '.dot' = 'application/msword'
                '.dotm' = 'application/vnd.ms-word.template.macroEnabled.12'
                '.dotx' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template'
                '.dsp' = 'application/octet-stream'
                '.dtd' = 'text/xml'
                '.dvi' = 'application/x-dvi'
                '.dwf' = 'drawing/x-dwf'
                '.dwp' = 'application/octet-stream'
                '.dxr' = 'application/x-director'
                '.eml' = 'message/rfc822'
                '.emz' = 'application/octet-stream'
                '.eot' = 'application/octet-stream'
                '.eps' = 'application/postscript'
                '.etx' = 'text/x-setext'
                '.evy' = 'application/envoy'
                '.exe' = 'application/octet-stream'
                '.exe.config' = 'text/xml'
                '.fdf' = 'application/vnd.fdf'
                '.fif' = 'application/fractals'
                '.fla' = 'application/octet-stream'
                '.flr' = 'x-world/x-vrml'
                '.flv' = 'video/x-flv'
                '.gif' = 'image/gif'
                '.gtar' = 'application/x-gtar'
                '.gz' = 'application/x-gzip'
                '.h' = 'text/plain'
                '.hdf' = 'application/x-hdf'
                '.hdml' = 'text/x-hdml'
                '.hhc' = 'application/x-oleobject'
                '.hhk' = 'application/octet-stream'
                '.hhp' = 'application/octet-stream'
                '.hlp' = 'application/winhlp'
                '.hqx' = 'application/mac-binhex40'
                '.hta' = 'application/hta'
                '.htc' = 'text/x-component'
                '.htm' = 'text/html'
                '.html' = 'text/html'
                '.htt' = 'text/webviewhtml'
                '.hxt' = 'text/html'
                '.ico' = 'image/x-icon'
                '.ics' = 'application/octet-stream'
                '.ief' = 'image/ief'
                '.iii' = 'application/x-iphone'
                '.inf' = 'application/octet-stream'
                '.ins' = 'application/x-internet-signup'
                '.isp' = 'application/x-internet-signup'
                '.IVF' = 'video/x-ivf'
                '.jar' = 'application/java-archive'
                '.java' = 'application/octet-stream'
                '.jck' = 'application/liquidmotion'
                '.jcz' = 'application/liquidmotion'
                '.jfif' = 'image/pjpeg'
                '.jpb' = 'application/octet-stream'
                '.jpe' = 'image/jpeg'
                '.jpeg' = 'image/jpeg'
                '.jpg' = 'image/jpeg'
                '.js' = 'application/x-javascript'
                '.jsx' = 'text/jscript'
                '.latex' = 'application/x-latex'
                '.lit' = 'application/x-ms-reader'
                '.lpk' = 'application/octet-stream'
                '.lsf' = 'video/x-la-asf'
                '.lsx' = 'video/x-la-asf'
                '.lzh' = 'application/octet-stream'
                '.m13' = 'application/x-msmediaview'
                '.m14' = 'application/x-msmediaview'
                '.m1v' = 'video/mpeg'
                '.m3u' = 'audio/x-mpegurl'
                '.man' = 'application/x-troff-man'
                '.manifest' = 'application/x-ms-manifest'
                '.map' = 'text/plain'
                '.mdb' = 'application/x-msaccess'
                '.mdp' = 'application/octet-stream'
                '.me' = 'application/x-troff-me'
                '.mht' = 'message/rfc822'
                '.mhtml' = 'message/rfc822'
                '.mid' = 'audio/mid'
                '.midi' = 'audio/mid'
                '.mix' = 'application/octet-stream'
                '.mmf' = 'application/x-smaf'
                '.mno' = 'text/xml'
                '.mny' = 'application/x-msmoney'
                '.mov' = 'video/quicktime'
                '.movie' = 'video/x-sgi-movie'
                '.mp2' = 'video/mpeg'
                '.mp3' = 'audio/mpeg'
                '.mpa' = 'video/mpeg'
                '.mpe' = 'video/mpeg'
                '.mpeg' = 'video/mpeg'
                '.mpg' = 'video/mpeg'
                '.mpp' = 'application/vnd.ms-project'
                '.mpv2' = 'video/mpeg'
                '.ms' = 'application/x-troff-ms'
                '.msi' = 'application/octet-stream'
                '.mso' = 'application/octet-stream'
                '.mvb' = 'application/x-msmediaview'
                '.mvc' = 'application/x-miva-compiled'
                '.nc' = 'application/x-netcdf'
                '.nsc' = 'video/x-ms-asf'
                '.nws' = 'message/rfc822'
                '.ocx' = 'application/octet-stream'
                '.oda' = 'application/oda'
                '.odc' = 'text/x-ms-odc'
                '.ods' = 'application/oleobject'
                '.one' = 'application/onenote'
                '.onea' = 'application/onenote'
                '.onetoc' = 'application/onenote'
                '.onetoc2' = 'application/onenote'
                '.onetmp' = 'application/onenote'
                '.onepkg' = 'application/onenote'
                '.osdx' = 'application/opensearchdescription+xml'
                '.p10' = 'application/pkcs10'
                '.p12' = 'application/x-pkcs12'
                '.p7b' = 'application/x-pkcs7-certificates'
                '.p7c' = 'application/pkcs7-mime'
                '.p7m' = 'application/pkcs7-mime'
                '.p7r' = 'application/x-pkcs7-certreqresp'
                '.p7s' = 'application/pkcs7-signature'
                '.pbm' = 'image/x-portable-bitmap'
                '.pcx' = 'application/octet-stream'
                '.pcz' = 'application/octet-stream'
                '.pdf' = 'application/pdf'
                '.pfb' = 'application/octet-stream'
                '.pfm' = 'application/octet-stream'
                '.pfx' = 'application/x-pkcs12'
                '.pgm' = 'image/x-portable-graymap'
                '.pko' = 'application/vnd.ms-pki.pko'
                '.pma' = 'application/x-perfmon'
                '.pmc' = 'application/x-perfmon'
                '.pml' = 'application/x-perfmon'
                '.pmr' = 'application/x-perfmon'
                '.pmw' = 'application/x-perfmon'
                '.png' = 'image/png'
                '.pnm' = 'image/x-portable-anymap'
                '.pnz' = 'image/png'
                '.pot' = 'application/vnd.ms-powerpoint'
                '.potm' = 'application/vnd.ms-powerpoint.template.macroEnabled.12'
                '.potx' = 'application/vnd.openxmlformats-officedocument.presentationml.template'
                '.ppam' = 'application/vnd.ms-powerpoint.addin.macroEnabled.12'
                '.ppm' = 'image/x-portable-pixmap'
                '.pps' = 'application/vnd.ms-powerpoint'
                '.ppsm' = 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12'
                '.ppsx' = 'application/vnd.openxmlformats-officedocument.presentationml.slideshow'
                '.ppt' = 'application/vnd.ms-powerpoint'
                '.pptm' = 'application/vnd.ms-powerpoint.presentation.macroEnabled.12'
                '.pptx' = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                '.prf' = 'application/pics-rules'
                '.prm' = 'application/octet-stream'
                '.prx' = 'application/octet-stream'
                '.ps' = 'application/postscript'
                '.psd' = 'application/octet-stream'
                '.psm' = 'application/octet-stream'
                '.psp' = 'application/octet-stream'
                '.pub' = 'application/x-mspublisher'
                '.qt' = 'video/quicktime'
                '.qtl' = 'application/x-quicktimeplayer'
                '.qxd' = 'application/octet-stream'
                '.ra' = 'audio/x-pn-realaudio'
                '.ram' = 'audio/x-pn-realaudio'
                '.rar' = 'application/octet-stream'
                '.ras' = 'image/x-cmu-raster'
                '.rf' = 'image/vnd.rn-realflash'
                '.rgb' = 'image/x-rgb'
                '.rm' = 'application/vnd.rn-realmedia'
                '.rmi' = 'audio/mid'
                '.roff' = 'application/x-troff'
                '.rpm' = 'audio/x-pn-realaudio-plugin'
                '.rtf' = 'application/rtf'
                '.rtx' = 'text/richtext'
                '.scd' = 'application/x-msschedule'
                '.sct' = 'text/scriptlet'
                '.sea' = 'application/octet-stream'
                '.setpay' = 'application/set-payment-initiation'
                '.setreg' = 'application/set-registration-initiation'
                '.sgml' = 'text/sgml'
                '.sh' = 'application/x-sh'
                '.shar' = 'application/x-shar'
                '.sit' = 'application/x-stuffit'
                '.sldm' = 'application/vnd.ms-powerpoint.slide.macroEnabled.12'
                '.sldx' = 'application/vnd.openxmlformats-officedocument.presentationml.slide'
                '.smd' = 'audio/x-smd'
                '.smi' = 'application/octet-stream'
                '.smx' = 'audio/x-smd'
                '.smz' = 'audio/x-smd'
                '.snd' = 'audio/basic'
                '.snp' = 'application/octet-stream'
                '.spc' = 'application/x-pkcs7-certificates'
                '.spl' = 'application/futuresplash'
                '.src' = 'application/x-wais-source'
                '.ssm' = 'application/streamingmedia'
                '.sst' = 'application/vnd.ms-pki.certstore'
                '.stl' = 'application/vnd.ms-pki.stl'
                '.sv4cpio' = 'application/x-sv4cpio'
                '.sv4crc' = 'application/x-sv4crc'
                '.swf' = 'application/x-shockwave-flash'
                '.t' = 'application/x-troff'
                '.tar' = 'application/x-tar'
                '.tcl' = 'application/x-tcl'
                '.tex' = 'application/x-tex'
                '.texi' = 'application/x-texinfo'
                '.texinfo' = 'application/x-texinfo'
                '.tgz' = 'application/x-compressed'
                '.thmx' = 'application/vnd.ms-officetheme'
                '.thn' = 'application/octet-stream'
                '.tif' = 'image/tiff'
                '.tiff' = 'image/tiff'
                '.toc' = 'application/octet-stream'
                '.tr' = 'application/x-troff'
                '.trm' = 'application/x-msterminal'
                '.tsv' = 'text/tab-separated-values'
                '.ttf' = 'application/octet-stream'
                '.txt' = 'text/plain'
                '.u32' = 'application/octet-stream'
                '.uls' = 'text/iuls'
                '.ustar' = 'application/x-ustar'
                '.vbs' = 'text/vbscript'
                '.vcf' = 'text/x-vcard'
                '.vcs' = 'text/plain'
                '.vdx' = 'application/vnd.ms-visio.viewer'
                '.vml' = 'text/xml'
                '.vsd' = 'application/vnd.visio'
                '.vss' = 'application/vnd.visio'
                '.vst' = 'application/vnd.visio'
                '.vsto' = 'application/x-ms-vsto'
                '.vsw' = 'application/vnd.visio'
                '.vsx' = 'application/vnd.visio'
                '.vtx' = 'application/vnd.visio'
                '.wav' = 'audio/wav'
                '.wax' = 'audio/x-ms-wax'
                '.wbmp' = 'image/vnd.wap.wbmp'
                '.wcm' = 'application/vnd.ms-works'
                '.wdb' = 'application/vnd.ms-works'
                '.wks' = 'application/vnd.ms-works'
                '.wm' = 'video/x-ms-wm'
                '.wma' = 'audio/x-ms-wma'
                '.wmd' = 'application/x-ms-wmd'
                '.wmf' = 'application/x-msmetafile'
                '.wml' = 'text/vnd.wap.wml'
                '.wmlc' = 'application/vnd.wap.wmlc'
                '.wmls' = 'text/vnd.wap.wmlscript'
                '.wmlsc' = 'application/vnd.wap.wmlscriptc'
                '.wmp' = 'video/x-ms-wmp'
                '.wmv' = 'video/x-ms-wmv'
                '.wmx' = 'video/x-ms-wmx'
                '.wmz' = 'application/x-ms-wmz'
                '.wps' = 'application/vnd.ms-works'
                '.wri' = 'application/x-mswrite'
                '.wrl' = 'x-world/x-vrml'
                '.wrz' = 'x-world/x-vrml'
                '.wsdl' = 'text/xml'
                '.wvx' = 'video/x-ms-wvx'
                '.x' = 'application/directx'
                '.xaf' = 'x-world/x-vrml'
                '.xaml' = 'application/xaml+xml'
                '.xap' = 'application/x-silverlight-app'
                '.xbap' = 'application/x-ms-xbap'
                '.xbm' = 'image/x-xbitmap'
                '.xdr' = 'text/plain'
                '.xla' = 'application/vnd.ms-excel'
                '.xlam' = 'application/vnd.ms-excel.addin.macroEnabled.12'
                '.xlc' = 'application/vnd.ms-excel'
                '.xlm' = 'application/vnd.ms-excel'
                '.xls' = 'application/vnd.ms-excel'
                '.xlsb' = 'application/vnd.ms-excel.sheet.binary.macroEnabled.12'
                '.xlsm' = 'application/vnd.ms-excel.sheet.macroEnabled.12'
                '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                '.xlt' = 'application/vnd.ms-excel'
                '.xltm' = 'application/vnd.ms-excel.template.macroEnabled.12'
                '.xltx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template'
                '.xlw' = 'application/vnd.ms-excel'
                '.xml' = 'text/xml'
                '.xof' = 'x-world/x-vrml'
                '.xpm' = 'image/x-xpixmap'
                '.xps' = 'application/vnd.ms-xpsdocument'
                '.xsd' = 'text/xml'
                '.xsf' = 'text/xml'
                '.xsl' = 'text/xml'
                '.xslt' = 'text/xml'
                '.xsn' = 'application/octet-stream'
                '.xtp' = 'application/octet-stream'
                '.xwd' = 'image/x-xwindowdump'
                '.z' = 'application/x-compress'
                '.zip' = 'application/x-zip-compressed'
    }

    $extension = [System.IO.Path]::GetExtension($FileName)
    $contentType = $mimeMappings[$extension]
    if ([string]::IsNullOrEmpty($contentType))
    {
        return New-Object System.Net.Mime.ContentType
    }
    else
    {
        return New-Object System.Net.Mime.ContentType($contentType)
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
        [Parameter(Position = 4, Mandatory = $true)]
        [Object]
        $To,
        [Parameter(Position = 5, Mandatory = $true)]
        [String]
        $Subject,
        [Parameter(Position = 6, Mandatory = $true)]
        [String]
        $Body,
        [Parameter(Mandatory = $false)]
        [String]
        $RedirectURI,
        [Parameter(Mandatory = $false)]
        [String[]]
        $Attachments,
        [Parameter(Mandatory = $false)]
        [Collections.HashTable]
        $InlineAttachments,
        [Parameter(Mandatory = $false)]
        [string[]]
        $Cc,
        [Parameter(Mandatory = $false)]
        [string[]]
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
        $Priority = "Normal",
        [Parameter(Mandatory = $false)]
        [switch]
        $SendwithSMTP

    )
    Process {
        if([String]::IsNullOrEmpty($RedirectURI)){
            $RedirectURI = "msal" + $ClientId + "://auth" 
        }
    
    If(!$SendwithSMTP){
    #send with Graph API
    # we're creating an object to convert it to json later. 
    $MessageObj = [PSCustomObject]@{}
    $MessageObj | Add-Member -MemberType NoteProperty -Name 'subject' -Value $subject

    # Build the body section
    If($BodyAsHTML){
        $contentType = "html"
    }else{
        $contentType = "text"
    }
    $BodyObj = [PSCustomObject]@{
        contentType = $contentType
        content = $body
    }

    $MessageObj | Add-Member -MemberType NoteProperty -Name 'body' -Value $BodyObj
    
    # build recipient section
    # to section
    $emailAddressArr = @()
    Foreach ($item in $to) {
        If ($item.Contains("<")){
            $ADDRESSParts = $item.Split("<")
            $NamePart = $ADDRESSParts[0].Trim(" ")
            $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
            $emailObj = [PSCustomObject]@{
                address = $mailPart
                name = $NamePart
            }
        }Else{
            $mailPart = $item.Trim("<").Trim(">").Trim(" ")
            $emailObj = [PSCustomObject]@{
                address = $mailPart
                name = ""
            }
        }
         $emailAddressArr += [PSCustomObject]@{emailAddress = $emailObj}
    }
    $toRecipientsObj = @(
        $emailAddressArr
        )
    $MessageObj | Add-Member -MemberType NoteProperty -Name 'toRecipients' -Value $toRecipientsObj
    # cc section
    if(![String]::IsNullOrEmpty($Cc)){
        $emailAddressArr = @()
        Foreach ($item in $Cc) {
            If ($item.Contains("<")){
                $ADDRESSParts = $item.Split("<")
                $NamePart = $ADDRESSParts[0].Trim(" ")
                $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                $emailObj = [PSCustomObject]@{
                    address = $mailPart
                    name = $NamePart
                }
            }Else{
                $mailPart = $item.Trim("<").Trim(">").Trim(" ")
                $emailObj = [PSCustomObject]@{
                    address = $mailPart
                    name = ""
                }
            }
            $emailAddressArr += [PSCustomObject]@{emailAddress = $emailObj}
        }
        $ccRecipientsObj = @(
            $emailAddressArr
        )
        $MessageObj | Add-Member -MemberType NoteProperty -Name 'ccRecipients' -Value $ccRecipientsObj
    }
    # bcc section
    if(![String]::IsNullOrEmpty($Bcc)){
        $emailAddressArr = @()
        Foreach ($item in $Bcc) {
            If ($item.Contains("<")){
                $ADDRESSParts = $item.Split("<")
                $NamePart = $ADDRESSParts[0].Trim(" ")
                $mailPart = $ADDRESSParts[1].Trim(" ").TrimEnd(">")
                $emailObj = [PSCustomObject]@{
                    address = $mailPart
                    name = $NamePart
                }
            }Else{
                $mailPart = $item.Trim("<").Trim(">").Trim(" ")
                $emailObj = [PSCustomObject]@{
                    address = $mailPart
                    name = ""
                }
            }
            $emailAddressArr += [PSCustomObject]@{emailAddress = $emailObj}
        }
        $bccRecipientsObj = @(
            $emailAddressArr
        )
        $MessageObj | Add-Member -MemberType NoteProperty -Name 'bccRecipients' -Value $bccRecipientsObj
    }
    # create email attachments section
    $hasAttachments = $false
    if($null -ne $Attachments){
        $AttachmentsArr = @()
        $hasAttachments = $true
        Foreach ($item in $Attachments) {
            $EncodedFile =  [convert]::ToBase64String( [system.io.file]::readallbytes($item))
            $FileType = (Get-ContentTypeFromFileName -FileName $item).MediaType
            $attachedFileObj = [PSCustomObject]@{
                '@odata.type' = "#microsoft.graph.fileAttachment"
                name = $item.Split("\")[$item.Split("\").count -1]
                contentBytes = $EncodedFile
                contentType = "application/octet-stream" #$FileType
            }
            $AttachmentsArr += $attachedFileObj
        }
    }
    # create section for inline attachemnts aka embedded files
    if($null -ne $InlineAttachments){
        $InlineAttachmentsArr = @()
        $hasAttachments = $true
        Foreach ($item in $InlineAttachments.GetEnumerator()) {
            $FileName = $item.Value.ToString()
            $EncodedFile =  [convert]::ToBase64String( [system.io.file]::readallbytes($FileName))
            $FileType = (Get-ContentTypeFromFileName -FileName $FileName).MediaType
            $InlineattachedFileObj = [PSCustomObject]@{
                '@odata.type' = "#microsoft.graph.fileAttachment"
                name = $FileName.Split("\")[$FileName.Split("\").count -1]
                contentBytes = $EncodedFile
                contentID = $item.Key
                isInline = $true
                contentType = $FileType
            }
            $InlineAttachmentsArr += $InlineattachedFileObj
        }
    }
    # both tyes are attachements from a technical view so we combine both
    If ($hasAttachments){
        $allAttachments = $AttachmentsArr + $InlineAttachmentsArr
        $MessageObj | Add-Member -MemberType NoteProperty -Name 'attachments' -Value $allAttachments
    }
    
    # message priority part
    $ChangePrority = $false
    switch ($Priority ){
        "High" { $ChangePrority = $true }
        "Low" { $ChangePrority = $true }  
    }
    If ($ChangePrority){
        $MessageObj | Add-Member -MemberType NoteProperty -Name 'importance' -Value $Priority
    }
    $fullObj = [PSCustomObject]@{message = $MessageObj}
    
    # convert all the stuff from above to json
    $jsonObj = $fullObj | ConvertTo-Json -Depth 50

    $ApiUrl = "https://graph.microsoft.com/v1.0/me/sendMail"
    
    # get the Token
    $token = Get-AccessTokenForMailSending -Credential $Credential -ClientId $ClientId -RedirectURI $RedirectURI -scopes "Mail.Send"
    
    # ...aaaaaand send
    Invoke-RestMethod -Headers @{Authorization = "Bearer $token"} -Uri $ApiUrl -Method Post -Body $jsonObj -ContentType "application/json"

    }
    Else{       
    # send with SMTP OAUTH 
    	# fill the from field
        if([String]::IsNullOrEmpty($From)){
            $SendingAddress = $Credential.UserName
        }
        else{
            $SendingAddress = $From
        } 
        # Building MailMessage Object       
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = New-Object System.Net.Mail.MailAddress($SendingAddress)
        # to section
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
	# cc section
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
	# bcc section
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
        
	#$mailMessage.Body = $Body <- this might be expected, but we use the alternate view
	#create the body
        if ($BodyAsHtml)
        {
            $bodyPart = [Net.Mail.AlternateView]::CreateAlternateViewFromString($Body, 'text/html')
            $mailMessage.IsBodyHtml = $true
        }
        else
        {
            $bodyPart = [Net.Mail.AlternateView]::CreateAlternateViewFromString($Body, 'text/plain')
        }   
        $mailmessage.AlternateViews.Add($bodyPart)
	
	# setup inline attachments aka embedded files
        if ($null -ne $InlineAttachments){
            foreach ($entry in $InlineAttachments.GetEnumerator()){
                $file = $entry.Value.ToString()
                if ([string]::IsNullOrEmpty($file)){
                    $mailmessage.Dispose()
                    throw "Send-MailMessage: Values in the InlineAttachments table cannot be null."
                }
                try{
                    $contentType = Get-ContentTypeFromFileName -FileName $file
                    $attachment = New-Object Net.Mail.LinkedResource($file, $contentType)
                    $attachment.ContentId = $entry.Key
                    $bodyPart.LinkedResources.Add($attachment)
                }
                catch{
                    $mailmessage.Dispose()
                    throw
                }
            }
        }
	# setup attachments
        if ($null -ne $Attachments){
            foreach ($file in $Attachments){
                try{
                    $contentType = Get-ContentTypeFromFileName -FileName $file
                    $mailmessage.Attachments.Add((New-Object Net.Mail.Attachment($file, $contentType)))
                }
                catch{
                    $mailmessage.Dispose()
                    throw
                }
            }
        }
	
	#set the encoding
        switch ($encoding){
            "ASCII" { $encodingObj = New-Object System.Text.ASCIIEncoding }
            "UTF8" { $encodingObj = New-Object System.Text.UTF8Encoding }
            "UniCode" { $encodingObj = New-Object System.Text.UnicodeEncoding }
            "UTF32" { $encodingObj = New-Object System.Text.UTF32Encoding }
            "UTF7" { $encodingObj = New-Object System.Text.UTF7Encoding }
        }
        #set email priority
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
        $token = Get-AccessTokenForMailSending -Credential $Credential -ClientId $ClientId -RedirectURI $RedirectURI
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
}
