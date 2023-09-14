# Created by:       Datyk
# DESCRIPTION:      Used to check mailbox rules on all mailboxes in a tennantand save the result as a excel spreadsheet, useful for checking for risky rules
# Requierd modules: ExchangeOnlineManagement, ImportExcel
# Error             You might set executionpolicy "Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser


# Check Requierd Modules

#Check for EXO v2 module inatallation
$Module = Get-Module ExchangeOnlineManagement -ListAvailable
if($Module.count -eq 0) 
{ 
 Write-Host Exchange Online PowerShell V2 module is not available  -ForegroundColor yellow  
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
 if($Confirm -match "[yY]") 
 { 
  Write-host "Installing Exchange Online PowerShell module"
  Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
 } 
 else 
 { 
  Write-Host EXO V2 module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet. 
  Exit
 }
}

#Check for ImportExcel module inatallation
$Module = Get-Module ImportExcel -ListAvailable
if($Module.count -eq 0) 
{ 
 Write-Host ImportExcel module is not available  -ForegroundColor yellow  
 $Confirm= Read-Host Are you sure you want to install module? [Y] Yes [N] No 
 if($Confirm -match "[yY]") 
 { 
  Write-host "Installing ImportExcel module"
  Install-Module ImportExcel -Repository PSGallery -AllowClobber -Force
 } 
 else 
 { 
  Write-Host ImportExcel module is required to export the result .Please install module using Install-Module ImportExcel cmdlet. 
  Exit
 }
} 

Connect-ExchangeOnline

$MainDomain = Get-AcceptedDomain | Where-Object Default -Match True | Select-Object -ExpandProperty Name
$Folder = "C:\Reports\MailboxRules"
$OutputFile = ("$Folder\$MainDomain-MailboxRules-" + (Get-Date).ToString("dd.MM.yyyy-hh.mm") + ".xlsx")
$Mailboxes = Get-Mailbox -ResultSize Unlimited

If(!(test-path $Folder))
{
    New-Item -ItemType Directory -Force -Path $Folder
}

Try {
    $Results = @()
    ForEach ($Mailbox in $Mailboxes) {
        Write-Host "Checking" $Mailbox.DisplayName - $Mailbox.UserPrincipalName -ForegroundColor Green

        $InboxRules = Get-InboxRule -Mailbox $Mailbox.UserPrincipalName

        If ($InboxRules) {
            ForEach ($InboxRule in $InboxRules) {
                $Results += New-Object PSObject -Property $([Ordered]@{
                        "MailboxOwnerID" = $InboxRule.MailboxOwnerID
                        "Name"           = $InboxRule.Name
                        "Description"    = $InboxRule.Description
                        "Enabled"        = $InboxRule.Enabled
                        "RedirectTo"     = $InboxRule.RedirectTo
                        "MoveToFolder"   = $InboxRule.MoveToFolder
                        "ForwardTo"      = $InboxRule.ForwardTo
                    })
            }
        }
    }
    $Results | Export-Excel -Path $OutputFile -WorksheetName "MailboxRules" -Append -AutoSize -TableStyle Medium7
    Write-Host `n"Done! Sucessfully exported to $OutputFile"`n -ForegroundColor Green
}
Catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Warning "An error occured! Is the folder '$Folder' created?"
}