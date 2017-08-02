#############################
#### CAPolicySQLDump.ps1 ####
#############################

Import-Module PoShPACLI | Out-Null

#############
# VARIABLES #
#############

## Load Settings.xml
[xml]$ConfigFile = Get-Content -Path "Settings.xml"
## Counter - DO NOT TOUCH
$counter = 0
## XML Settings Declarations
$pacliFolder = $ConfigFile.Settings.PACLI.Folder
$vaultName = $ConfigFile.Settings.Vault.Name
$vaultAddress = $ConfigFile.Settings.Vault.Address
$vaultUsername = $ConfigFile.Settings.PACLI.Username
$vaultUserIni = $ConfigFile.Settings.PACLI.CredFile
$vaultSafe = $ConfigFile.Settings.Vault.Safe
$sqlTable = $ConfigFile.Settings.SQL.Table
$sqlPolIDCol = $ConfigFile.Settings.SQL.Column.PolicyID
$sqlPolNameCol = $ConfigFile.Settings.SQL.Column.PolicyName
$sqlServer = $ConfigFile.Settings.SQL.Server
$sqlDatabase = $ConfigFile.Settings.SQL.Database
$sqlUsername = $ConfigFile.Settings.SQL.Username
$sqlPassword = $ConfigFile.Settings.SQL.Password


#############
#  EXECUTE  #
#############

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
if (!$vaultDef) {Write-Host "Failed to define Vault. Address: "$vaultAddress"" -ForegroundColor Red; break}
else {Write-Host "[ Vault Definition Set ]" -ForegroundColor Yellow}

## Connect to Vault as User (Logon)
$connect = Connect-Vault -vault $vaultName -user $vaultUsername -logonFile $vaultUserIni
if (!$connect) {Write-Host "Failed to connect to Vault at address "$vaultAddress"." -ForegroundColor Red; break}
else {Write-Host "[ Connected to "$vaultName" ]" -ForegroundColor Yellow}

## Open safe to work in
$openSafe = Open-Safe -vault $vaultName -user $vaultUsername -safe $vaultSafe
if (!$openSafe) {Write-Host "Failed to open safe "$vaultSafe"." -ForegroundColor Red; break}
else {Write-Host "[ "$vaultSafe" Safe Opened ]" -ForegroundColor Yellow}

## Revert to CLI PACLI to use FILESLIST correctly and retrieve all filenames from safe
$filesList = C:\PACLI\PACLI.exe FILESLIST VAULT='TEST' USER='Administrator' SAFE='PasswordManager_info' FOLDER='Root\Policies' OUTPUT'(NAME)'

foreach ($file in $filesList) {
    ## Increment counter
    $counter++

    ## Retrieve file from Vault to local temp path
    $getFile = Get-File -vault $vaultName -user $vaultUsername -safe $vaultSafe -folder Root\Policies -file $row.PolicyFileName -localFolder $env:TEMP -localFile "${row.PolicyFileName}.tmp" -evenIfLocked

    ## Error handling for file retrieve
    if (!$getFile) {Write-Host "Failed to retrieve file ${row} from safe ${vaultSafe}." -ForegroundColor Red; break}
    else {Write-Host "[ Received file ${counter} of ${csvCount} ]" -ForegroundColor Yellow}

    ## Get content of file retrieved and prune for parts needed
    $fileReceived = Get-Content "${env:TEMP}/${row.PolicyFileName}.tmp"
    $PolicyID = $fileReceived | Select-String -Pattern "PolicyID=" -Encoding ascii | Select-Object -Last 1
    $PolicyName = $fileReceived | Select-String -Pattern "PolicyName=" -Encoding ascii | Select-Object -Last 1
    
    ## Trim everything but PolicyID and PolicyName
    $strPolicyID = $PolicyID.Line.Split(";")
    $strPolicyName = $PolicyName.Line.Split(";")
    $strPolicyID = $strPolicyID.Split("=")
    $strPolicyName = $strPolicyName.Split("=")
    $PolicyID = $strPolicyID[1].Trim()
    $PolicyName = $strPolicyName[1].Trim()
    
    ## DEBUGGING
    Write-Host "PolicyID:   ${PolicyID}"
    Write-Host "PolicyName: ${PolicyName}"

    ## SQL query to run
    $sqlQuery = "INSERT INTO ${ConfigFile.Settings.SQL.Table} (${ConfigFile.Settings.SQL.Column.PolicyID},${ConfigFile.Settings.SQL.Column.PolicyName}) VALUES ('${PolicyID}','${PolicyName}'); "

    try {
        ## Begin SQL Connection
        $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $sqlConnection.ConnectionString = "Server=${ConfigFile.Settings.SQL.Server};Database=${ConfigFile.Settings.SQL.Database};Integrated Security=False; User ID=${ConfigFile.Settings.SQL.Username}; Password=${ConfigFile.Settings.SQL.Password};"
        $sqlConnection.Open()
        if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
            Write-Host "Connection to the DB could not be established." -ForegroundColor Red
            break
        }

        ## Send SQL Query
        $sqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $sqlCmd.Connection = $sqlConnection
        $sqlCmd.CommandText = $sqlQuery

        ## Other
        $sqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $sqlAdapter.SelectCommand = $sqlCmd
        $dataSet = New-Object System.Data.DataSet
        $sqlAdapter.Fill($dataSet)

        if (sqlConnection.State -eq [Data.ConnectionState]::Open) {
            $sqlConnection.Close()
        }
    }
    catch {
        Write-Host "Errors occurred during SQL connection/query execution." -ForegroundColor Red
        break
    }

    ## Remove variables before loop to ensure empty variables
    Remove-Variable -Name getFile, fileReceived, PolicyID, PolicyName
}

## Disconnect user from Vault (Logoff)
$disconnect = Disconnect-Vault -vault $vaultName -user $vaultUsername
if (!$disconnect) {Write-Host "Failed to disconnect from "$vaultName"." -ForegroundColor Red; break}
else {Write-Host "[ Disconnected from "$vaultName" ]" -ForegroundColor Yellow}

## Stop PACLI process and close
$stop = Stop-PACLI
if (!$stop) {Write-Host "Failed to stop PACLI. Please manually stop PACLI and retry script." -ForegroundColor Red; break}
else {Write-Host "[ Stopped PACLI successfully ]" -ForegroundColor Yellow}