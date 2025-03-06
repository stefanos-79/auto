# Automate mailbox archive expansion

$myEmailAddress = "o-pgeorgiadis@deloitte.onmicrosoft.com"
$emailFilePath = "C:\MyScripts\UserEmailAddress.txt"
$userEmailAddress = Get-Content -Path $emailFilePath | Select-Object -First 1
$username = $userEmailAddress.Split('@')[0]

# Connect to Exchange Online | Replace "" with $myEmailAddress to connect directly to your o-account
Connect-ExchangeOnline -UserPrincipalName ""

# Get current archive size for the user
$archiveSize = (Get-MailboxStatistics -Archive $username).TotalItemSize
#$archiveSize = Get-MailboxStatistics -Archive $username | select TotalItemSize,DisplayName
Write-Host "Current archive size for $username : $archiveSize"

# Enable auto-expanding archive option
Enable-Mailbox $userEmailAddress -AutoExpandingArchive
Write-Host "Auto-expanding archive enabled for $userEmailAddress"

# Verify auto-expand archive option
$userMailboxSettings = Get-Mailbox $userEmailAddress | FL AutoExpandingArchiveEnabled
Write-Host "Auto-expanding archive enabled status: $($userMailboxSettings.AutoExpandingArchiveEnabled)"

# Disconnect from Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online."