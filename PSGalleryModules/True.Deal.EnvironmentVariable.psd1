﻿@{
GUID="845286df-8df9-4cbb-89a5-0b29f5a01861"
Author="OtogawaKatsutoshi"
Copyright="Copyright (c) OtogawaKatsutoshi"
Description="Windows True Environment"
ModuleVersion="0.5.1.2"
# CompatiblePSEditions = @("Core", "Desktop")
PowerShellVersion="4.0"
NestedModules="PowerShell.Commands.True.Deal.EnvironmentVariable.dll"
HelpInfoURI = 'https://github.com/KatsutoshiOtogawa/True.Deal.EnvironmentVariable'
FunctionsToExport = @()
CmdletsToExport=@(
    "Get-WinEnvironmentVariable"
    ,"Set-WinEnvironmentVariable"
   )
PrivateData = @{

	PSData = @{
		Tags=@(
			"Environment"
			,"EnvironmentVariable"
			,"Windows"
		)
		LicenseUri ='https://github.com/KatsutoshiOtogawa/PowerShell.Commands.True.Deal.EnvironmentVariable/PowerShell.Commands.True.Deal.EnvironmentVariable/licenses/LICENSE.txt' 

		ProjectUri = 'https://github.com/KatsutoshiOtogawa/PowerShell.Commands.True.Deal.EnvironmentVariable'

		# Prerelease = ''

		ReleaseNotes = @'

'@
	}
}
}
