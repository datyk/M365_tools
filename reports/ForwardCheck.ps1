<#
This script will export a list of users and mark users with email forwarding as a CSV file.

#>
Connect-ExchangeOnline
$Folder = "C:\temp"
$OutputFile = ("$Folder\$MainDomain-FwdCheck.csv")
$MainDomain = Get-AcceptedDomain | Where-Object Default -Match True | Select-Object -ExpandProperty Name
$FwdCheck = get-mailbox -resultsize unlimited | Select-Object PrimarySMTPAddress,ForwardingSMTPAddress, DeliverToMailboxAndForward

$FwdCheck |  Export-CSV $OutputFile -NoTypeInformation -Append -Encoding UTF8