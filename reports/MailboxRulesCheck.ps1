# Husk å koble til :)
# Connect-ExchangeOnline


$Folder = "C:\temp"
$OutputFile = ("$Folder\$MainDomain-MailboxRules.csv")
$MainDomain = Get-AcceptedDomain | Where-Object Default -Match True | Select-Object -ExpandProperty Name
$Mailboxes = Get-Mailbox -ResultSize Unlimited

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