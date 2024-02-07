# Created by:       github.com/datyk
# DESCRIPTION:      Used to check mailbox rules on all mailboxes in a tennantand save the result as a .csv file, useful for checking for risky rules
# Requierd modules: ExchangeOnlineManagement, ImportExcel
# Error             You might set executionpolicy "Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser


# Function to check module installation and prompt for installation if not present
function CheckAndInstallModule($moduleName) {
    $Module = Get-Module -Name $moduleName -ListAvailable
    if ($Module.count -eq 0) {
        Write-Host "$moduleName module is not available" -ForegroundColor Yellow
        $Confirm = Read-Host "Are you sure you want to install module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-host "Installing $moduleName module"
            Install-Module -Name $moduleName -Repository PSGallery -AllowClobber -Force
        } else {
            Write-Host "$moduleName module is required for this script. Please install module using Install-Module $moduleName cmdlet."
            return $false
        }
    }
    return $true
}

# Check for Exchange Online Management and ImportExcel modules
if (-not (CheckAndInstallModule "ExchangeOnlineManagement") -or -not (CheckAndInstallModule "ImportExcel")) {
    Exit
}

# Connect to Exchange Online
Connect-ExchangeOnline

# Get domain and prepare output file path
$MainDomain = Get-AcceptedDomain | Where-Object Default -Match True | Select-Object -ExpandProperty Name
$Folder = "C:\Reports\MailboxRules"
$OutputFile = Join-Path $Folder -ChildPath "$MainDomain-MailboxRules-$(Get-Date -Format 'dd.MM.yyyy-HH.mm').xlsx"

# Ensure the directory exists and inform the user
If (-not (Test-Path -Path $Folder)) {
    Write-Host "Directory '$Folder' does not exist. Creating it now..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $Folder
} else {
    Write-Host "Directory '$Folder' found." -ForegroundColor Green
}

Try {
    $Results = @()
    $Mailboxes = Get-Mailbox -ResultSize Unlimited

    ForEach ($Mailbox in $Mailboxes) {
        Write-Host "Checking" $Mailbox.DisplayName - $Mailbox.UserPrincipalName -ForegroundColor Green

        $InboxRules = Get-InboxRule -Mailbox $Mailbox.UserPrincipalName -IncludeHidden


        foreach ($InboxRule in $InboxRules) {
            $Results += [PSCustomObject][Ordered]@{
                "MailboxOwnerID" = $InboxRule.MailboxOwnerID
                "Name"           = $InboxRule.Name
                "Description"    = $InboxRule.Description
                "Enabled"        = $InboxRule.Enabled
                "RedirectTo"     = $InboxRule.RedirectTo
                "MoveToFolder"   = $InboxRule.MoveToFolder
                "ForwardTo"      = $InboxRule.FowardTo
            }
        }
    }

    $Results | Export-Excel -Path $OutputFile -WorksheetName "MailboxRules" -Append -AutoSize -TableStyle Medium7
    Write-Host "`nDone! Successfully exported to $OutputFile`n" -ForegroundColor Green
}
Catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Warning "An error occurred! Please check if the folder '$Folder' exists and is writable."
}