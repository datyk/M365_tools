<#
This script will export a list of users with roles in a tenant to an Excel file.

#>

if (!(Get-Module -ListAvailable -Name MSOnline)) {
    Write-Host "`nModule does not exist - please install the MSOnline module, and retry.`n" -ForegroundColor Red
}
Import-Module MSOnline -Force -ErrorAction Stop

if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "`nModule does not exist - please install the ImportExcel module, and retry." -ForegroundColor Red
}
Import-Module ImportExcel -Force -ErrorAction Stop


Write-Host "`nConnect to Tenant" -ForegroundColor Yellow
Connect-MsolService -ErrorAction Stop

$TenantName = (Get-MsolCompanyInformation).DisplayName
Write-Host "Successfully connected to: $TenantName`n" -ForegroundColor Green

$OutputFolder = "C:\Temp\Roles\"
$OutputFile = ("$OutputFolder" + "$TenantName" + " - Roles " + (Get-Date).ToString("dd.MM.yyyy") + ".xlsx") -replace "/", ""


if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
}


$Results = @()
$Roles = Get-MsolRole
ForEach ($Role in $Roles) {
    $Users = Get-MsolRoleMember -TenantId $Tenant.TenantId -RoleObjectId $Role.ObjectId
    ForEach ($User in $Users) {
        if ($User) {
            $Results += New-Object PSObject -Property $([Ordered]@{
                    "Tenant"        = $TenantName
                    "Display Name"  = $User.DisplayName
                    "Email Address" = $User.EmailAddress
                    "Role"          = $Role.Name
                })
        }
    }
}
$Results | Export-Excel -Path $OutputFile -WorksheetName "Roles" -Append -AutoSize -TableStyle Medium7

Write-Host "`n`nAll data has been exported to: $OutputFile`n" -ForegroundColor Green