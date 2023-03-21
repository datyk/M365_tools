<#
This script will export a list of users with Inbox rule as a CSV file
Useful to check if users have a malicious inbox rule
#>

Connect-ExchangeOnline

$MainDomain = Get-AcceptedDomain | Where-Object Default -Match True | Select-Object -ExpandProperty Name
$Mailboxes = Get-Mailbox -ResultSize Unlimited
$Folder = "C:\temp"
$OutputFile = ("$Folder\$MainDomain-MailboxRules.csv")

Try {
foreach ($Mailbox in $Mailboxes) {
    Write-Host "Checking" $Mailbox.DisplayName - $Mailbox.UserPrincipalName -ForegroundColor Green
    Get-InboxRule -Mailbox $Mailbox.UserPrincipalName | Select-Object MailboxOwnerID,Name,Description,Enabled,RedirectTo,MoveToFolder,ForwardTo | Export-CSV $OutputFile -NoTypeInformation -Append -Encoding UTF8
}
    Write-Host `n"Done! Sucessfully exported to $OutputFile"`n -ForegroundColor Green
}
Catch {
    Write-Warning "An error occured! Is the folder '$Folder' created?"
}