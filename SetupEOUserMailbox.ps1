#sets login info for 365 pulling password from encrypted text file.
$uLogin = [string]"<365 admin login>"
$uPass = cat C:\PowerShell.Scripts\archiveP.txt|ConvertTo-SecureString
$uCredentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $uLogin, $uPass

#items that must be manually set, Search the document for:
#<365 admin login>
#<your Domain name>
#<possible UPN>
#<primary SMTP address domain>
#<alias smtp address domain>


do{
##stores username , eg. jsmith, in variable $uName
$uName = $null

$uName = Read-Host "Please enter the user's Sam Account Name"

#Checks user SAM Name entered properly
Try{
	Get-Aduser $uName -ErrorAction SilentlyContinue
}
Catch{
	Write-Host "The user you have specified does not exist in AD."`n"Please Check the users UserName and try again"`n"Press any key to exit this script ..." -ForeGroundColor Red -BackgroundColor Black
	$keyPress = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	Exit
}

##Imports the module to sync the user to Office 365 and forces a sync
Import-Module dirsync
start-onlinecoexistencesync

## forces powershell to wait 30 seconds before continuing, this allows times for the new user to be added to the cloud
Write-Host "The script will pause for 30 seconds to allow the user information to be created online"
Start-Sleep -s 30

##Connect to MSOL by asking for login credentials used to sign into Office 365
Connect-MsolService -Credential $uCredentials

#adjust time to allow proper connection otherwise script may error out
Start-Sleep -s 30

#Sets user UPN in AD to simplify sign in on 365.
Set-ADUser $uName -UserPrincipalName "$uName@<your Domain name>"

start-onlinecoexistencesync
start-sleep -s 300

##Sets user UPN/login for 365 if not carried over from previous step, if it's possible the users UPN is already set in 365 and different from their AD UPN add them as needed. 
##If immediate access is not needed, comment out this section and the users login should begin after the mail sync.
Set-MsolUserPrincipalName -UserPrincipalName "$uName@<your Domain name>.onmicrosoft.com" -NewUserPrincipalName "$uName@<your domain name>" -ErrorAction SilentlyContinue
Set-MsolUserPrincipalName -UserPrincipalName "$uName@<possible UPN>" -NewUserPrincipalName "$uName@<your Domain name>" -ErrorAction SilentlyContinue

#Sets the users default email address while adding aliases. The default sending address uses 'SMTP' while aliases use 'smtp'.
#If a default SMTP address has already been set-up, it is recommended to remove before running this script.
#If need to remove once of two primary SMTP objects arise use the following command: Set-Aduser <SAM Account Name> -Remove @{proxyaddresses ="SMTP:<address>"}.
Set-Aduser $uName -Add @{proxyaddresses = "SMTP:$uName@<primary SMTP address domain"}
#Set-Aduser $uName -Add @{proxyaddresses = "smtp:$uName@<alias smtp address domain>"}


#$UserPN stores the users UPN from Office 365
$UserPN = [string]"$uName@<your domain name>"

##Sets user location to us and adds the license for Exchange Online
Set-MsolUser -UserPrincipalName $UserPN -UsageLocation US #Set your own 2 character country code here
#All licenses can be found use 'Get-MsolAccountSku' in powershell after using the 'connect-msolservice' command
Set-MsolUserLicense -UserPrincipalName $UserPN -AddLicenses <Your License Name>

#This Section connects a PSSession to Exchange online if one is not already started
If (!(Get-Command "Get-Mailbox" -ErrorAction SilentlyContinue)) {
      Try {
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $uCredentials -Authentication Basic -AllowRedirection
            Import-PSSession $Session  -ErrorAction Stop
      } Catch {
            Write-Host "An error occurred while establishing a remote PowerShell session with Exchange Online." -ForegroundColor Red -BackgroundColor Black
            Exit
      }
}


##Forces a second wait for mailbox to be built/assigned
Write-Host "The Script will pause to allow the user's MailBox to be created"
Start-Sleep -s 30

if(!(Get-Mailbox $UserPN -ErrorAction SilentlyContinue)){
Clear
Write-Host "The users mailbox has not been created yet."`n "We will continue to run this check every 15 Seconds until the mailbox is created or 30 minutes have passed." -ForeGroundColor Red -BackgroundColor Black
Write-Host
Write-Host
Write-Host "Would you like the script to continue?... (Y)es (N)o: " -ForegroundColor Yellow -BackgroundColor Black -NoNewLine
$endScript = read-host 
if($endScript -contains "N"){
	Exit
}
}

$attemptTime = 0
while(!(Get-MailBox $UserPN -ErrorAction SilentlyContinue)){	
	Start-Sleep -s 15
	$attemptTime++
		if($attempTime -eq 120){
			Clear
			$endScript
			write-host "It has been 30 minutes, please check that the user is created and has the license applied in Office 365" -ForeGroundColor Red -BackgroundColor Black
			write-host
			write-host
			write-host "Would you like the script to continue?... (Y)es (N)o: " -ForeGroundColor Yellow -BackgroundColor Black -NoNewLine
			$endScript = read-host
			if($endScript -contains "N"){
				Exit
			}
			Clear
			write-host "After you have confirmed the users Mailbox has been created with the correct SMTP address press any key to continue ..." -ForeGroundColor Yellow -BackgroundColor Black
			$keyPress = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
		}
}

##if the users SMTP address is not set properly in the previous step, or previously set this step will set an incorrect SMTP address.


##This section for setting OWA signature was written and posted on http://www.davethijssen.nl/2014/05/set-uniform-individual-outlook-web-app.html by Dave Thijssen. A more detailed setup can be found there.
## Sets Users Signature in OWA
$ScriptRoot = Split-Path $MyInvocation.MyCommand.Definition

#This section creates variables and designates the signature file location, and forces use of this signature in OWA
$SignatureFileName = ($ScriptRoot + "\Signature.html")
$SignatureHtml = Get-Content $SignatureFileName | Out-String
$AutoAddSignature = $true
#Collects user information for the signature, most information must be set in AD and synced to 365.
$MailBoxUser = Get-Mailbox -Identity $UserPN -RecipientTypeDetails UserMailbox | Get-MailboxMessageConfiguration
#stores the information from $MailBoxUser in a secondary variable and replaces Identifiable variables with the users information from AD.
$user = Get-User $MailBoxUser.Identity
Set-MailboxMessageConfiguration -Identity ($MailBoxUser.Identity) `
    -AutoAddSignature $AutoAddSignature `
    -SignatureHtml ($SignatureHtml   -replace "%%DisplayName%%", $user.DisplayName `
    -replace "%%Title%%", $user.Title `
    -replace "%%PhoneNumber%%", $user.Phone `
    -replace "%%Email%%", $user.WindowsEmailAddress `
    -replace "%%Company%%", $user.Company
    )
	
	
#Disables the mail protocols for the user
Set-CASMailbox $uName -PopEnabled $False -ErrorAction Stop
Set-CASMailbox $uName -ImapEnabled $False -ErrorAction Stop



#Syncs all changes made
Start-OnlineCoexistenceSync

##This cleans up the window and shows the license, the total amount purchased, and the amount used.
Clear
write-host "The Following Information is your Licenses and their total amount and the amount used..."
Get-MsolAccountSku|Select-Object AccountSkuId,ActiveUnits,ConsumedUnits

##Holds until the user presses any key
Write-Host
Write-Host
Write-Host
Write-Host "Press any key to continue ..."
$keyPress = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

Clear

$title = "New Email Accounts"
$message = "Would you like to run this Script again?"

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Runs the script again."

$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Stops the loop."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$result = $host.ui.PromptForChoice($title,$message,$options, 0)

switch ($result){
0 {"You have selected yes";start-sleep -s 5 -;Clear}
1 {"You have selected no"}
}

If($result -eq 1){
Remove-PSSession $Session
}

}
while($result -eq 0)