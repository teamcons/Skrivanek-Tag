

if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }





ps2exe `
-inputFile $ScriptPath\..\main.ps1 `
-iconFile $ScriptPath\..\assets\stamp.ico `
-noConsole `
-noOutput `
-exitOnCancel `
-title "Tag! - Fuer Skrivanek" `
-description "Schnell dateien bearbeiten !" `
-company "Skrivanek GmbH" `
-copyright "CC0 1.0 Universal Stella - stella.menier@gmx.de" `
-version 0.9 `
-Verbose `
-outputFile $ScriptPath\..\Tag.exe


# Usage (ps1): powershell.exe -executionpolicy bypass -file "M:\4_BE\06_General information\Stella\tag_UNSTABLE.ps1" "%1"

# Usage: tag_UNSTABLE.exe "%1"
