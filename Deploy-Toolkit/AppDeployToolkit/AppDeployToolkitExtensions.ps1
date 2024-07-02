<#
.SYNOPSIS
	This script is a template that allows you to extend the toolkit with your own custom functions.
    # LICENSE #
    PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows.
    Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
    This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
    You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is automatically dot-sourced by the AppDeployToolkitMain.ps1 script.
.NOTES
    Toolkit Exit Code Ranges:
    60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
    69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
    70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
)

##*===============================================
##* VARIABLE DECLARATION
##*===============================================

# Variables: Script
[string]$appDeployToolkitExtName = 'PSAppDeployToolkitExt'
[string]$appDeployExtScriptFriendlyName = 'App Deploy Toolkit Extensions'
[version]$appDeployExtScriptVersion = [version]'3.8.4'
[string]$appDeployExtScriptDate = '26/01/2021'
[hashtable]$appDeployExtScriptParameters = $PSBoundParameters

# Functions
[string]$AuditKeyPath = 'HKLM:\Software\DXC\Installed\'
[string]$AuditKey = ("$AuditKeyPath" + "$DevID")
[string]$KeyToRemove = ("$AuditKeyPath" + "$DevID")

##*===============================================
##* FUNCTION LISTINGS
##*===============================================

# <Your custom functions go here>

Function Audit-Key {

<# Create or delete audit key
   Audit-Key -Action "Create" will create the audit key.
   Audit-Key -Action "Delete" will delete the audit key.
#>
    [CmdletBinding()]
	Param (
        [Parameter(Mandatory=$true)]
		[ValidateSet('Create','Delete')]
		[string]$Action = 'Create'
    ) 

    Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {

        ## Create audit key parent path if path does not exist.
     IF ($Action -eq 'Create') {
        
        # Creating audit key properties
        
            Set-RegistryKey -Key "$AuditKey" -Name "Application Name" -Value "$appName" -Type "String"
            Set-RegistryKey -Key "$AuditKey" -Name "Application Version" -Value "$appVersion" -Type "String"
            Set-RegistryKey -Key "$AuditKey" -Name "Installed By" -Value "$env:USERNAME" -Type "String"
            Set-RegistryKey -Key "$AuditKey" -Name "Installed On" -Value "$currentDateTime" -Type "String"
            Set-RegistryKey -Key "$AuditKey" -Name "Installed From" -Value "$PSScriptRoot" -Type "String"
                
        } ELSE {
            
            # Deleting the audit key
            
            Remove-RegistryKey -Key "$KeyToRemove"
      
      } 

    }
}

##*===============================================
##* END FUNCTION LISTINGS
##*===============================================

##*===============================================
##* SCRIPT BODY
##*===============================================

If ($scriptParentPath) {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] dot-source invoked by [$(((Get-Variable -Name MyInvocation).Value).ScriptName)]" -Source $appDeployToolkitExtName
} Else {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] invoked directly" -Source $appDeployToolkitExtName
}

##*===============================================
##* END SCRIPT BODY
##*===============================================
