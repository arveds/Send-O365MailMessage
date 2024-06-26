# Disclaimer
This solution is outdated, there are more elegant and easier ways to send emails via the Graph API. Please, look for other solutions.
Sending with SMTP is disabled by Microsoft, so this part is obsolete and not working any longer.

# Send-O365MailMessage
    
This is inspried fantastic script here: https://github.com/gscales/Powershell-Scripts/blob/master/TLS-SMTP-Oauth-Mod.ps1  
**Kudos to gscales**. 
Meanwhile there are so many changes, that this wasn't forked, because of the different intention to use the script.  
  
## Deprecation of Exchange online SMTP Basic auth  
Microsoft disabled SMTP Basic auth in the second half of 2021 see [here](https://docs.microsoft.com/en-us/lifecycle/announcements/exchange-online-basic-auth-deprecated), [here](https://developer.microsoft.com/de-de/office/blogs/deferred-end-of-support-date-for-basic-authentication-in-exchange-online/) or in the [MessageCenter MC237741](https://app.cloudscout.one/evergreen-item/mc237741/) [MessageCenter MC204828](https://app.cloudscout.one/evergreen-item/mc204828/)

For PowerShell sending Emails a soulution is needed.  

## Challange
The fact SMTP Basic Auth will be disabled, leads -for PowerShell scripts that are required to send Emails via SMTP- to some issues.  
The change from SMTP basic Auth to XOAUTH2 has lots of implications, caveats and obstacles.  
While looking myself for a solution the internet gave me the advice to use another external mail provider which allows STMP  in not using the internal SMTP solution. But this will decrease security instead of increasing it. And I not sure how many developers already went down this road.  
  
The recommended way by Microsoft is to use “MailKit” as a library together with MSAL. The downside here is, that MailKit does not work with PowerShell 5.1. For MailKit PowerShell 7 is needed. But it is no option to install PowerShell 7 on every system and migrate each and every PowerShell 5 script to 7.  
The good news are, that there is a PowerShell module from the PowerShell Gallery MSAL.PS that works on PowerShell 5.1 and uses MSAL as library.  

## Solution
Long story told short: 
### Dual Stack  
there are two options to solve this, the GraphAPI and SMTP OAUTH. I do not like the idea to use EWS, because deprecation of EWS by Microsoft is in the air.  
The default is sending with the GraphAPI. With the -SendWithSMTP switch you can change to send via SMTP/OAUTH.  

### Azure AD App registration
The solution I propose is to register one(!) App for all clients in AzureAD (Public App / User credentials)  
With the following properties:  
API Permissions: SMTP.Send and Mail.Send
Authentication: Redirect URI: the msal.....//auth (MSAL (only) Uri must be checked  
Authentication: Allow public client flows: Yes  
Authentication: Supported Accout types: Single tenant (your directory)  
  
No client secret needed.  
The respective user account for sending emails has to give user consent or must be added to the app by an AAD Admin as user.  
This app can be used by any solution authentication through user credential (Username / Password)  

### MSAL.PS Setup
For PowerShell 5.1 which is included in every Windows, I set up a script to provide a proof-of concept code working with MSAL (and MSAL.PS to make things easier).  

the Powershell.Module from the powershell Gallery MSAL.PS is requiered.
To use  
`Install-Module MSAL.ps`  
you might need to run though this:  
`Update-Module`   
`Get-Module`  
check if PowerShellGet is higher than 1.0.0.1  
`Set-ExecutionPolicy RemoteSigned` <- needed for Module to run  
`Install-PackageProvider Nuget –force –verbose`  
`Install-Module -Name PowerShellGet -Force -AllowClobber`  
`Exit` <- important  
close shell and ISE and check back again  
  
### Send-O365MailMessage.ps1  
This script implements -more or less- a handcrafted SMTP STARTTLS XOAUTH2 client. And a GraphAPI based solution 
  
If there are better ways, any solution is very welcome  
  
The syntax is very similar to the Send-MailMessage.  
  
you may also use the other functions seperatly  
to incorporate the functions into your script simply dot-source start the script from your ps1:  
  
`. "<path to script>\Send-O365MailMessage.ps1"` (see: https://devblogs.microsoft.com/scripting/how-to-reuse-windows-powershell-functions-in-scripts/)  
  

## Syntax
Syntax:  
   
 -Credential  -> PSCredentialObject (Username and Password of the sending Office 365 account) (mandatory)  
 -ClientID    -> The Application (client) ID of the registered Azure AD App (mandatory)  
 -Subject     -> Email Subject (mandantory)  
 -Body        -> Email content (mandantory)  
 -To  	      -> Recipent email (mandantory)  
 -Cc          -> cc (optional)  
 -Bcc         -> bcc (optional)  
 ---------   -> Email Adresses can be added as `'Name somewhat <name@emaildomain.com>'` or `'anothername@emaildomain.com'` or as an array `-to 'email1@domain.com','Its me <email2@domain.com>'`  
 -RedirectURI -> RedirectURI as configured in the registed Azure AD App (optional) defaults to msal$ClientID://auth  
 -Attachments -> File Path(s) (optional)  
 -SMTPServer  -> defaults to smtp.office365.com (optional)  
 -BodyAsHTML  -> Switch indicates if Body is in html  
 -Encoding    -> Encoding of the subject and the body one of "ASCII","UTF8","UniCode","UTF32","UTF7" (optional) defaults to "UTF8"  
 -From        -> From email adress. (optional) defaults to UserName from PSCredential Object   
 -Priority    -> one of "Low", "High", "Normal" [string]. (optional) defaults to "Normal"  
 -InlineAttachment -> Hashtable like `@{image1 = "C:\myimage.png"}` The name can be referenced in the html body to inline the file like `<img src="cid:image1">`.  
 
 no longer working due to Microsofts SMTP switch off:  
 -SendWithSMTP -> change the sending from GraphAPI to SMTP w/ OAUTH (optional)  
 -Port        -> SMTP Port (optional). defaults to 587  
  
