#creates a txt file in the powershell.scripts folder on the C drive containing an encrypted password.
#this is used to store a password on the local machine to allow for automation of Office 365 scripts.

$uPass = read-host -AsSecureString
$uPass | ConvertFrom-SecureString |Out-File -FilePath "C:\Powershell.Scripts\archiveP.txt"