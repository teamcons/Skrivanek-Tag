
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Auto renamer following ISO Format for Skrivanek GmbH
# Usage: powershell.exe -executionpolicy remotesigned -file "M:\4_BE\06_General information\Stella\tagv2.ps1" "%1"

############
#FIXME : DETECT NO MATTER HOW DEEP
param([String]$arg)
$filename = Split-Path -leaf $arg
$dirname = Split-Path -leaf ..
$dircode = $dirname.SubString(0, 10)
$path_to_file = $arg

############
# IF REVIEW JUST APPEND SUFFIX AND STOP




############
# IGNORE XLIFF

if ( $filename.Contains(".sdlxliff") )
{
	echo "TAG: Its an XLIFF ! Ignoring."
	exit
}



############
# ADD PROJECT CODE

if ( $filename.SubString(0, 10) -ne $dircode )
{
	echo "TAG: No directory code. Adding."
    $newname = ($dircode + $filename)
	Rename-Item -Path $path_to_file -NewName $newname #-WhatIf -Verbose
	$path_to_file = ((Split-Path -Path $path_to_file) + "\" + $newname)
	$filename = $newname
	
}


############
# IF ORIG, ADD ORIG

$parent = Split-Path -leaf .
if (( $parent -eq "01_orig" ) -and (! $filename.Contains("_orig") ))
{
	echo "TAG: Original file but no _orig detected. Adding."
	$extension = $filename[-5..-1] -join ""
	$filename = $filename -replace ".{5}$"	
	$newname = ($filename + "_orig" + $extension)
	Rename-Item -Path $path_to_file -NewName $newname #-WhatIf -Verbose
}



############
# IF COUNTRY CODE ADD FINAL

if (( $parent -match "[A-Z][A-Z]\-[A-Z][A-Z]" ) -and (! $filename.Contains("_Final") ))
{
	echo "TAG: FINAL. No final tag. Adding."
	$extension = $filename[-5..-1] -join ""
	$filename = $filename -replace ".{10}$"	
	$newname = ($filename + "_" + $parent + "_Final" + $extension)
	Rename-Item -Path $path_to_file -NewName $newname #-WhatIf -Verbose
}


# Move-Item -Path $path_to_file -Destination ..\..\..\07_to client


############
# IF FINAL BUT NO FINAL ADD FINAL
#
#if (( $parent -match "*_to client" ) -and (! $filename.Contains("_Final") ))
#{
#	echo "TAG: FINAL. No final tag. Unknown LCODE. Adding."
#	$extension = $filename[-5..-1] -join ""
#	$filename = $filename -replace .{10}$	
#	$newname = ($filename + "_" + "LCODE" + "_Final" + $extension)
#	Rename-Item -Path $path_to_file -NewName $newname #-WhatIf -Verbose
#}




############
# DEBUG

# pause