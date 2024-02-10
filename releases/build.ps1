ps2exe `
-inputFile ..\src\tag_UNSTABLE.ps1 `
-iconFile ..\assets\stamp.ico `
-noConsole `
-noOutput `
-exitOnCancel `
-title "Tag! - Fuer Skrivanek" `
-description "Schnell dateien bearbeiten !" `
-company "Skrivanek GmbH" `
-copyright "CC0 1.0 Universal Stella - stella.menier@gmx.de" `
-version 0.9 `
-Verbose `
-outputFile ..\releases\Tag.exe


# Usage (ps1): powershell.exe -executionpolicy bypass -file "M:\4_BE\06_General information\Stella\tag_UNSTABLE.ps1" "%1"

# Usage: tag_UNSTABLE.exe "%1"
