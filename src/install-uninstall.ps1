
$REG_PATH = "Registry::HKEY_CURRENT_USER\Software\Classes\*\shell\Tag"


# True if exists
$IsInstalled = Test-Path -Path $REG_PATH


# Remove commands	
#Remove-Item -Path $REG_PATH\command
#Remove-Item -Path $REG_PATH



# Install commands
$ICO = Get-Item ..\assets\stamp.ico).FullName
$EXE = Get-Item .\tag_stable.ps1).FullName

New-Item -Path $REG_PATH
New-ItemProperty -Path $REG_PATH -Name "(Standard)" -Value "Tag - Schnell umbenennen"
New-ItemProperty -Path $REG_PATH -Name "Icon" -Value $ICO
New-ItemProperty -Path $REG_PATH -Name "MuiVerb" -Value "Tag - Schnell umbenennen"

New-Item -Path $REG_PATH\command
New-ItemProperty -Path $REG_PATH\command -Name "(Standard)" -Value "powershell.exe -noprofile -executionpolicy bypass -file \"$EXE\" \"%1\""




# Usage (ps1): powershell.exe -executionpolicy bypass -file "M:\4_BE\06_General information\Stella\tag_UNSTABLE.ps1" "%1"
# Usage: tag_UNSTABLE.exe "%1"
