#############################
#### CAPolicySQLDump.ps1 ####
#############################

## Import PoShPACLI to ease use
Import-Module PoShPACLI | Out-Null

#############
# VARIABLES #
#############

## Load Settings.xml
[xml]$ConfigFile = Get-Content -Path "Settings.xml"
##### DO NOT CHANGE BELOW #####
## Counter
$counter = 0
$errColor = "Green"
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
Try {Start-PVPACLI -ErrorAction Stop -Verbose}
Catch {
	Write-Warning $(($error[0].Exception).Message).ToString()
	Write-Verbose "Failed to start PACLI. Please make sure PACLI process is killed and try again." -Verbose
	Return
}

## Define Vault Parameters
New-PVVaultDefinition -vault $vaultName -address $vaultAddress

## Connect to Vault as User (Logon)
Try {
	Connect-PVVault -vault $vaultName -user $vaultUsername -logonFile $vaultUserIni -ErrorAction Stop
} Catch {
	Write-Warning $(($error[0].Exception).Message).ToString()
	Write-Verbose "Failed to connect to Vault at address ${vaultAddress}." -Verbose
	Return
}

## Open safe to work in
Try {
	Open-PVSafe -vault $vaultName -user $vaultUsername -safe $vaultSafe -ErrorAction Stop | Out-Null
} Catch {
	Write-Warning $(($error[0].Exception).Message).ToString()
	Write-Verbose "Failed to open safe ${vaultSafe}." -Verbose
	Return
}

Try {
	$filesList = Get-PVFileList -vault $vaultName -user $vaultUsername -safe $vaultSafe -folder 'Root\Policies'
} Catch {
	Write-Warning $(($error[0].Exception).Message).ToString()
	Write-Verbose "Failed to retrieve file list from safe ${vaultSafe}." -Verbose
	Return
} Finally {

	##If Results
	if($filesList) {
		## Loop through each file given back
		foreach ($file in ($filesList| Select-Object -ExpandProperty Name)) {


			## Increment counter
			$counter++
			Write-Verbose "Getting File: $File ($counter of $($FilesList.count))" -Verbose
			Try {
				Get-PVFile -vault $vaultName -user $vaultUsername -safe $vaultSafe -folder 'Root\Policies' -file $file `
					-localFolder $env:TEMP -localFile "$file.tmp" -evenIfLocked -Verbose -ErrorAction Stop
			} Catch {
				Write-Warning $(($error[0].Exception).Message).ToString()
				Write-Warning "Failed to retrieve file $file from safe ${vaultSafe}." -Verbose
			}

			## Get content of file retrieved and prune for parts needed
			$fileReceived = Get-Content "${env:TEMP}/$file.tmp"
			$PolicyID = $fileReceived | Select-String -Pattern "PolicyID=" -Encoding ascii | Select-Object -Last 1
			$PolicyName = $fileReceived | Select-String -Pattern "PolicyName=" -Encoding ascii | Select-Object -Last 1

			## Trim everything but PolicyID and PolicyName
			$strPolicyID = $PolicyID.Line.Split(";")
			$strPolicyName = $PolicyName.Line.Split(";")
			$strPolicyID = $strPolicyID.Split("=")
			$strPolicyName = $strPolicyName.Split("=")
			$PolicyID = $strPolicyID[1].Trim()
			$PolicyName = $strPolicyName[1].Trim()

			## SQL query to run
			$sqlQuery = "INSERT INTO ${sqlTable} (${sqlPolIDCol},${sqlPolNameCol}) VALUES ('${PolicyID}','${PolicyName}'); "

			try {
				## Begin SQL Connection
				$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
				$sqlConnection.ConnectionString = "Server=${sqlServer};Database=${sqlDatabase};Integrated Security=False; User ID=${sqlUsername}; Password=${sqlPassword};"
				$sqlConnection.Open()
				if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
					Write-Host "Connection to the server\instance ${sqlServer} or database ${sqlDatabase} could not be established." -ForegroundColor Red
					break
				}

				## Send SQL Query
				$sqlCmd = New-Object System.Data.SqlClient.SqlCommand
				$sqlCmd.Connection = $sqlConnection
				$sqlCmd.CommandText = $sqlQuery

				## Close SQL Connection
				if (sqlConnection.State -eq [Data.ConnectionState]::Open) {
					$sqlConnection.Close()
				}
			}
			## If error occurs, do below
			catch {
				Write-Host "Errors occurred during SQL connection/query execution." -ForegroundColor Red
				$errColor = "Red"
				break
			}

			## Echo row entries in color based on entry status
			Write-Host "${sqlPolIDCol}:   ${PolicyID}" -ForegroundColor $errColor
			Write-Host "${sqlPolNameCol}: ${PolicyName}" -ForegroundColor $errColor

			## Remove variables before loop to ensure empty variables
			Remove-Variable -Name fileReceived, PolicyID, PolicyName

			##Remove Local Copy of Policy File
			Remove-Item -Path "$env:TEMP\$file.tmp" -Verbose

		}
	}
}

# Close safe after using
Try {
	Close-PVSafe -vault $vaultName -user $vaultUsername -safe $vaultSafe -Verbose -ErrorAction Stop
} Catch {
	Write-Warning $(($error[0].Exception).Message).ToString()
	Write-Verbose "Failed to close safe ${vaultSafe}." -Verbose
}

## Disconnect user from Vault (Logoff)
Try {
	Disconnect-PVVault -vault $vaultName -user $vaultUsername -Verbose -ErrorAction Stop
} Catch {
	Write-Warning $(($error[0].Exception).Message).ToString()
	Write-Verbose "Failed to logoff from vault ${vaultName}." -Verbose
}
## Stop PACLI process and close
Try {
	Stop-PVPACLI -Verbose
} Catch {
	Write-Warning $(($error[0].Exception).Message).ToString()
	Write-Verbose "PACLI TERM Command Failed" -Verbose
}