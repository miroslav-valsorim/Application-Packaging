<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
	# LICENSE #
	PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows.
	Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
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
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall','Repair')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Silent',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}

	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'COREL'
	[string]$appName = 'Corel Draw Graphic Suite'
	[string]$appVersion = '2022'
	[string]$appLang = 'DE'
	[string]$appRevision = 'R01'
	[string]$devID = ''
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '10/08/2022'
	[string]$appScriptAuthor = 'Miroslav Hristov'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = ''
	[string]$installTitle = ''

	##* Do not modify section below
	#region DoNotModify

	## Variables: Exit Code
	[int32]$mainExitCode = 0

	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.8.4'
	[string]$deployAppScriptDate = '26/01/2021'
	[hashtable]$deployAppScriptParameters = $psBoundParameters

	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}

	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
		Show-InstallationWelcome -CloseApps 'iexplore' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Installation tasks here>


		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'

		## Handle Zero-Config MSI Installations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) { $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ } }
		}

		## <Perform Installation tasks here>
		
		Copy-Item -Path "$dirFiles\CorelDRAW Graphics Suite 2022*" -Destination "$envProgramFiles\CorelDRAW Graphics Suite 2022" -Recurse
		
		Start-Process -FilePath "$envProgramFiles\CorelDRAW Graphics Suite 2022\Setup.exe" -Parameters" /qn DESKTOPSHORTCUTS=0"
		
		Start-Process -FilePath "$dirFiles\CorelDrawCurrentUserSettings.msi" /qn
		
		# Execute-MSI -Path '.msi' -Parameters '' -private:logname
		
		# Execute-Process -Path '.exe' -Parameters ''
		
		# Copy-Item -Path '$dirSupportFiles\'
		
		#Create Audit Key
		Audit-Key -Action 'Create'
		
		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## <Perform Post-Installation tasks here>
		
		## Display a message at the end of the install
		If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'You can customize text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait }
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
		Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Uninstallation tasks here>


		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		## Handle Zero-Config MSI Uninstallations
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
			Execute-MSI @ExecuteDefaultMSISplat
		}

		# <Perform Uninstallation tasks here>
		
		#Start-Process -FilePath "$envProgramFiles\CorelDRAW Graphics Suite 2022\Setup.exe" /x
		
        Execute-Process -Path "$envProgramFiles\CorelDRAW Graphics Suite 2022\Setup.exe" -Parameters '/x REMOVE=ALL REMOVE_GPL=1 REMOVE_SHELLEXT=1 /qn'
		
		Start-Sleep -Seconds 30
		
		Execute-MSI -Action 'Uninstall' -Path '{98CFADA3-527D-4A92-9160-EE463FCE95A5}' -Parameters "-qn" -private:APBK799_COREL_CorelDrawGraphicSuite_2022_Uninstall
		
		Execute-MSI -Action 'Uninstall' -Path '{76E381CE-5AD1-4A02-9CF4-B407B1BE9BE0}' -Parameters "-qn" -private:APBK799_COREL_CorelDrawGraphicSuite_2022_Uninstall
		
		Execute-MSI -Action 'Uninstall' -Path '{CF6AFE4D-5D61-43D0-8F75-A7B1C6FC0275}' -Parameters "-qn" -private:APBK799_COREL_CorelDrawGraphicSuite_CorelDrawCurrentUserSettings_2022_Uninstall
		
		Execute-MSI -Action 'Uninstall' -Path '{03389409-4F66-41A6-AD54-2AB04F7B82BB}' -Parameters "-qn" -private:APBK799_COREL_CorelDrawGraphicSuite_BRx64_2022_Uninstall
		
		Execute-MSI -Action 'Uninstall' -Path '{1E4B5F2C-0532-4CDA-AFCD-674E9C37521E}' -Parameters "-qn" -private:APBK799_COREL_CorelDrawGraphicSuite_SetupFiles_2022_Uninstall
		
		Execute-MSI -Action 'Uninstall' -Path '{06CD45E6-FF5E-4D8E-BC01-B276A90DADF2}' -Parameters "-qn" -private:APBK799_COREL_CorelDrawGraphicSuite_Ghostscript_2022_Uninstall
		
		
		$HKCURegistrySettings = {
		Remove-RegistryKey -Key 'HKEY_CURRENT_USER\Software\Corel' -SID $UserProfile.SID -Recurse		
	    }
	    Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings
		
		Execute-Process -Path "$envWinDir\System32\cscript.exe" -Parameters """$dirFiles\DeleteFilesAndFoldersFromUserProfilesAndSystem.vbs"""
		
		Remove-Folder -Path "$envProgramFiles\CorelDRAW Graphics Suite 2022"
	
		#EdgeWebViewerUninstallString
		#There are two folder paths, depends on the setup exe folder (104.0.1293.54 or 96.0.1054.34) run the commands below.
		#Execute-Process -Path "$envProgramFilesX86\Microsoft\EdgeWebView\Application\104.0.1293.54\Installer\Setup.exe" -Parameters '--uninstall --msedgewebview --system-level --verbose-logging --force-uninstall'
		#Execute-Process -Path "$envProgramFilesX86\Microsoft\EdgeWebView\Application\96.0.1054.34\Installer\Setup.exe" -Parameters '--uninstall --msedgewebview --system-level --verbose-logging --force-uninstall'
		#Remove-Folder -Path "$envProgramFilesX86\Microsoft\EdgeCore"
		
		#VisualStudioToolsUninstallString	
		#Execute-MSI -Action 'Uninstall' -Path '{9D6CE289-E12C-38BB-9999-E2377EC118B7}' -Parameters "-qn" -private:APBK799_COREL_VisualStudioTools_2022_Uninstall
		#Execute-MSI -Action 'Uninstall' -Path '{1edcd8d2-905a-4e93-bfdf-92ed5601528a}' -Parameters "-qn" -private:APBK799_COREL_VisualStudioTools_2022_Uninstall
		#Execute-MSI -Action 'Uninstall' -Path '{7C931D41-F302-3494-868C-320A4F4DD9F9}' -Parameters "-qn" -private:APBK799_COREL_VisualStudioTools_2022_Uninstall
	
		# Removing Audit Key
		Audit-Key -Action 'Delete'
		
		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'

		## <Perform Post-Uninstallation tasks here>


	}
	
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
