# CAPolicySQLDump
Takes PolicyID and PolicyName from every file in the PasswordManager_info safe and INSERT INTO capolicies table in mssql to cols caPolicyID and caPlatformName.

## Pre-Requisites
* This PowerShell script is unsigned and may throw warnings from PowerShell.  To prevent this, run `Set-ExecutionPolicy Unrestricted` in an elevated PowerShell console window.
* `git clone https://github.com/psPete/PoShPACLI.git` into `C:\Windows\system32\WindowsPowerShell\v1.0\Modules`
* Install CyberArk PACLI into `C:\PACLI` on the local machine this script is run from.

## Support
Please update the [Issues](https://github.com/infamousjoeg/CAPolicySQLDump/issues) on this repo for support.

## License
MIT