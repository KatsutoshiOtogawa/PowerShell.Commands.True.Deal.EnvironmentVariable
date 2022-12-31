@{
GUID="845286df-8df9-4cbb-89a5-0b29f5a01861"
Author="OtogawaKatsutoshi"
Copyright="Copyright (c) OtogawaKatsutoshi"
Description="Windows True Environment"
ModuleVersion="0.1.1.0"
Tags=@(
    "Environment"
    ,"EnvironmentVariable"
    ,"Windows"
)
CompatiblePSEditions = @("Core")
PowerShellVersion="7.2"
NestedModules="PowerShell.Commands.True.Deal.EnvironmentVariable.dll"
HelpInfoURI = 'https://github.com/KatsutoshiOtogawa/True.Deal.EnvironmentVariable'
FunctionsToExport = @()
CmdletsToExport=@(
    "Get-WinEnvironmentVariable"
    ,"Set-WinEnvironmentVariable"
   )
}
