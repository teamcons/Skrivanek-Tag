
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Auto renamer following ISO Format for Skrivanek GmbH
# Usage: powershell.exe -executionpolicy remotesigned -file "M:\4_BE\06_General information\Stella\tagv2.ps1" "%1"

############
#FIXME : DETECT NO MATTER HOW DEEP

param([String]$arg)

$path_to_file = $arg
$file = (Get-Item $arg)
$parent = $file.Directory.Parent.Name
$filename = $file.Name

echo "[TAG - ISO RENAMER TOOL] For Skrivanek - By Stella Ménier - stella.menier@gmx.de
[PARAMETER] File: $filename"



############
# DETECT DIRCODE


if ($file.FullName.Contains('M:\9_JOBS_XTRF') )
{
	$dircode = ($file.DirectoryName -Split "\\")[3]
	$dircode = $dircode.SubString(0, 9)
	echo "[DETECTED] File in structure. Dircode: $dircode"
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



############
# IF REVIEW JUST APPEND SUFFIX AND STOP

if (( $file.DirectoryName -match "[A-Z][A-Z]\-[A-Z][A-Z]" -and ( $file.Name.EndsWith(".docx.review")) ))
{
	echo "[DETECTED] Bilingual doc."
	$newname = ( $file.DirectoryName + "__" + $file.Name )

	echo " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"

	# move it to correct place
	echo "TAG: Moving to '05_to proof'"
	Move-Item -Path "$file" -Destination "..\..\..\..\05_to proof"

	exit
}



############
# IGNORE XLIFF

if ( $file.Name.Contains(".sdlxliff") )
{
	echo "[DETECTED] Its an XLIFF ! Ignoring."
	exit
}



############
# ADD PROJECT CODE

if ( $file.Name.SubString(0, 9) -ne $dircode )        #-and (! $file.FullName.Contains("_Final") ) )
{
	echo "[DETECTED] $file.Name : No directory code. Adding"
        $newname = ($dircode + "_" + $file.Name)

	echo " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
}


############
# IF ORIG, ADD ORIG


if ( $file.DirectoryName.Contains("_orig") ) #-and (! $file.Name.Contains("_orig") ))
{
	echo "[DETECTED] Original file but no _orig detected. Adding."
	$newname = ($file.Basename + "_orig" + $file.Extension)

	echo " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
}



############
# IF COUNTRY CODE ADD FINAL

if (( $file.DirectoryName -match "[a-z][a-z]\-[A-Z][A-Z]" ) -and (! $file.Name.Contains("_Final") ))
{
	echo "[DETECTED] FINAL. No final tag. Adding."
	$newname = ($file.Basename.Replace("_orig","") + "_" + (Split-Path -Leaf $file.DirectoryName) + "_Final" + $file.Extension)

	echo " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"

	# And move it
	echo "[ACTION] Moving to '07_to client'"
	Move-Item -Path "$file" -Destination "..\..\07_to client\"
	exit

}





############
# IF FINAL BUT NO FINAL ADD FINAL

if (( $file.DirectoryName.Contains("_to client") ) -and (! $file.Name.Contains("_Final") ))
{
	echo "[DETECTED] FINAL. No final tag. Unknown LCODE. Adding."
	$newname = ($file.Basename.Replace("_orig","") + "_Final" + $file.Extension)

	echo " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
	exit
}




############
# IF PO/CO, SORT OUT

#if $path_to_file.Contains("Downloads")
#{
#	echo "[DETECTED] PO/CO - SORTING OUT"

	# Generate correct dircode
	#$dircode = $dircode.SubString(0, 7)       #$dircode.replace('_','-')
	# $destination = "M:\9_JOBS_XTRF\" + (Get-ChildItem "$dircode.SubString(0, 7)"  )
	# $destination = ( "M:\9_JOBS_XTRF\" + "$dircode" + "\00_info\" )	
	#echo "destination: $destination" 
	# Move-Item -Path $path_to_file -Destination $destination
	#exit
#}

exit
