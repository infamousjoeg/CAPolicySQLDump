#############################
#### CAPolicySQLDump.ps1 ####
#############################

Import-Module PoShPACLI | Out-Null

# VARIABLES (CHANGE HERE)
## Path to PACLI.exe
$pacliFolder = "C:\PACLI"
## Name of Vault (anything works)
$vaultName = "VAULT"
## IP Address of Vault
$vaultAddress = "192.168.2.150"
## Username to connect to Vault as
$username = "Administrator"
## Path to user.ini created by CreateCredFile
$userIni = "C:\PACLI\user.ini"
## Name of safe to work in
$safe = "PasswordManager_info"
## Path to policy.csv file containing names of Policy-{Platform}.ini files
$policyCSV = "C:\PACLI\policy.csv"

# EXECUTE
## Initialize PACLI
$init = Initialize-PoShPACLI -pacliFolder $pacliFolder
if (!$init) {Write-Host "Failed to initialize PACLI." -ForegroundColor Red; break}
else {Write-Host "[ PACLI Initialized ]" -ForegroundColor Yellow}

## Start PACLI Process
$start = Start-PACLI
if (!$start) {
    Write-Host "Failed to start PACLI. Restarting..." -ForegroundColor Yellow
    $stop = Stop-PACLI
    if (!$stop) {
        Write-Host "Failed to stop PACLI. Please manually stop PACLI and retry script." -ForegroundColor Red
        break
    }
    else {
        Write-Host "PACLI has been stopped. Re-attempting start..." -ForegroundColor Yellow
        $restart = Start-PACLI
        if (!$restart) {Write-Host "Failed to restart PACLI. Please make sure PACLI process is killed and try again." -ForegroundColor Red; break}
    }
}
else {Write-Host "[ PACLI Started ]" -ForegroundColor Yellow}

## Define Vault Parameters
$vaultDef = Add-VaultDefinition -vault $vaultName -address $vaultAddress
if (!$vaultDef) {Write-Host "Failed to define Vault. Address: ${vaultAddress}" -ForegroundColor Red; break}
else {Write-Host "[ Vault Definition Set ]" -ForegroundColor Yellow}

## Connect to Vault as User (Logon)
$connect = Connect-Vault -vault $vaultName -user $username -logonFile $userIni
if (!$connect) {Write-Host "Failed to connect to Vault at address ${vaultAddress}." -ForegroundColor Red; break}
else {Write-Host "[ Connected to ${vaultName} ]" -ForegroundColor Yellow}

## Open safe to work in
$openSafe = Open-Safe -vault $vaultName -user $username -safe $safe
if (!$openSafe) {Write-Host "Failed to open safe ${safe}." -ForegroundColor Red; break}
else {Write-Host "[ ${safe} Safe Opened ]" -ForegroundColor Yellow}

## Import CSV file for reading
try {
    $csvFile = Import-CSV -Path $policyCSV
    $csvCount = $csvFile.Count
    Write-Host "[ Imported ${csvCount} rows from ${policyCSV} ]" -ForegroundColor Yellow
}
catch {
    Write-Host "Failed to import CSV file located at ${policyCSV}." -ForegroundColor Red
    break
}

foreach ($row in $csvFile) {
    ## Retrieve file from Vault to local temp path
    $getFile = Get-File -vault $vaultName -user $username -safe $safe -folder Root\Policies -file $row.PolicyFileName -localFolder $env:TEMP -localFile "${row.PolicyFileName}.tmp" -evenIfLocked

    ## Error handling for file retrieve
    if (!$getFile) {Write-Host "Failed to retrieve file ${row} from safe ${safe}." -ForegroundColor Red; break}
    else {Write-Host "[ Received file ${row} ]" -ForegroundColor Yellow}

    ## Get content of file retrieved and prune for parts needed
    $fileReceived = Get-Content "${env:TEMP}/${row.PolicyFileName}.tmp"
    $PolicyID = $fileReceived | Select-String -Pattern "PolicyID=" -Encoding ascii | Select-Object -Last 1
    $PolicyName = $fileReceived | Select-String -Pattern "PolicyName=" -Encoding ascii | Select-Object -Last 1
    
    ## TODO: Add trimming/pruning to remove all except PolicyID and PolicyName.
    $strPolicyID = $PolicyID.Line.Split(";")
    $strPolicyName = $PolicyName.Line.Split(";")
    $strPolicyID = $strPolicyID.Split("=")
    $strPolicyName = $strPolicyName.Split("=")
    
    $PolicyID = $strPolicyID[1].Trim()
    $PolicyName = $strPolicyName[1].Trim()

    Write-Host "PolicyID:   ${PolicyID}"
    Write-Host "PolicyName: ${PolicyName}"

    Remove-Variable -Name getFile, fileReceived, PolicyID, PolicyName
}

## Disconnect user from Vault (Logoff)
$disconnect = Disconnect-Vault -vault $vaultName -user $username
if (!$disconnect) {Write-Host "Failed to disconnect from ${vaultName}." -ForegroundColor Red; break}
else {Write-Host "[ Disconnected from ${vaultName} ]" -ForegroundColor Yellow}

## Stop PACLI process and close
$stop = Stop-PACLI
if (!$stop) {Write-Host "Failed to stop PACLI. Please manually stop PACLI and retry script." -ForegroundColor Red; break}
else {Write-Host "[ Stopped PACLI successfully ]" -ForegroundColor Yellow}