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

# EXECUTE
Initialize-PoShPACLI -pacliFolder $pacliFolder

Start-PACLI

Add-VaultDefinition -vault $vaultName -address $vaultAddress

Connect-Vault -vault $vaultName -user $username -logonFile $userIni

Open-Safe -vault $vaultName -user $username -safe $safe

$files = Get-FilesList -vault $vaultName -user $username -safe $safe -folder "Root"

foreach ($file in $files) {
    Write-Host $file
}