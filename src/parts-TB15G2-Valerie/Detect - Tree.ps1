
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Auto renamer following ISO Format for Skrivanek GmbH
# Usage: powershell.exe -executionpolicy remotesigned -file "M:\4_BE\06_General information\Stella\tagv2.ps1" "%1"

############
#FIXME : DETECT NO MATTER HOW DEEP
param([String]$arg)
$path_to_file = $arg

$file = (Get-Item "$arg")
$parent = $file.Directory.Parent.Name
$filename = $file.Name

echo "[TAG - ISO RENAMER TOOL] For Skrivanek - By Stella Ménier - stella.menier@gmx.de
[PARAMETER] File: $filename"



############
# DETECT DIRCODE - REBUILD TREE



#EXAMPLE
# M:\9_JOBS_XTRF\2023_3000-3099\2023-3078_De Backer
#  cd "M:\9_JOBS_XTRF\2023_3000-3099\2023-3078_De Backer\01_orig"

if ($file.FullName.Contains('M:\9_JOBS_XTRF') )
{
	$dircode = ($file.DirectoryName -Split "\\")[3]
	$dircode = $dircode.SubString(0, 9)
	echo "[DETECTED] File in structure. Dircode: $dircode"


    $dirpath = 




}




elseif ( $file.Name.SubString(0, 9) -match "20[0-9][0-9]\-*" )
{
	$dircode = $file.Name.SubString(0, 9)
	echo "[DETECTED] Uprooted file, can relocate. Dircode: $dircode"




}





else
{
	echo "[FATAL] NO DIRCODE DETECTED - Cannot do my job"
	exit




}


##202[\d]_\d\d\d\d-\d\d\d\d



