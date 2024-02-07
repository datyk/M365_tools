# Created by:    github.com/datyk
# DESCRIPTION:   Used to check if mailboxes in a tennant have forwarding turned on and save the result as a .xlsx file
# Problems       It might be neccesry to set executionpolicy for the script to work Example: "Set-ExecutionPolicy -ExecutionPolicy bypass"   


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
            Write-Host "$moduleName module is required for this script. Please install the module using Install-Module $moduleName cmdlet."
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
$Folder = "C:\reports\fwdcheck"
$OutputFile = Join-Path $Folder -ChildPath "$MainDomain-fwdcheck-$(Get-Date -Format 'dd.MM.yyyy-HH.mm').xlsx"

# Ensure the directory exists and inform the user
If (-not (Test-Path -Path $Folder)) {
    Write-Host "Directory '$Folder' does not exist. Creating it now..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $Folder
} else {
    Write-Host "Directory '$Folder' found." -ForegroundColor Green
}

# Perform forwarding check and export to Excel
Try {
    $FwdCheck = Get-Mailbox -ResultSize Unlimited | 
                Select-Object PrimarySMTPAddress, ForwardingSMTPAddress, DeliverToMailboxAndForward
    $FwdCheck | Export-Excel -TableName fwdcheck -TableStyle Medium13 -Path $OutputFile
    Write-Host "Export completed successfully. Data saved to $OutputFile" -ForegroundColor Green
}
Catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Warning "An error occurred during export. Please ensure you have the necessary permissions and the path '$Folder' is accessible."
}
