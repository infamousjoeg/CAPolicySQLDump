#############################
#### CAPolicySQLDump.ps1 ####
#############################

Import-Module PoShPACLI

# VARIABLES (CHANGE HERE)
$pacliFolder = "C:\PACLI"
$vaultName = "VAULT"
$vaultAddress = "192.168.2.150"
$username = "Administrator"
$userIni = "C:\PACLI\user.ini"
$safe = "PasswordManager_info"
$policyCSV = "C:\PACLI\policy.csv"

# EXECUTE
Initialize-PoShPACLI -pacliFolder $pacliFolder

Start-PACLI

Add-VaultDefinition -vault $vaultName -address $vaultAddress

Connect-Vault -vault $vaultName -user $username -logonFile $userIni

Open-Safe -vault $vaultName -user $username -safe $safe

$csvFile = Import-CSV -Path $policyCSV

foreach ($row in $csvFile) {
    Get-File -vault $vaultName -user $username -safe $safe -folder Root\Policies -file $row.PolicyFileName -localFolder $env:TEMP -localFile "${row.PolicyFileName}.tmp" -evenIfLocked
    $fileReceived = Get-Content "${env:TEMP}/${row.PolicyFileName}.tmp"
    $PolicyID = $fileReceived | Select-String -Pattern "PolicyID=" -Encoding ascii | Select-Object -Last 1
    $PolicyName = $fileReceived | Select-String -Pattern "PolicyName=" -Encoding ascii | Select-Object -Last 1
    Write-Host "PolicyID:   ${PolicyID}"
    Write-Host "PolicyName: ${PolicyName}"
}