
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Auto renamer following ISO Format for Skrivanek GmbH
# Usage: powershell.exe -executionpolicy bypass -file "M:\4_BE\06_General information\Stella\tag_UNSTABLE.ps1" "%1"

############ INITIALIZATION ############
# Get all important variables in place


param([String]$arg)
if (!$arg) 
{
        Add-Type -AssemblyName PresentationCore,PresentationFramework
	$msgBody = "Cannot detect filename :("
	[System.Windows.MessageBox]::Show($msgBody)
	exit
}

Write-Host "[TAG - ISO RENAMER TOOL] For Skrivanek - By Stella Ménier - stella.menier@gmx.de"

$path_to_file = $arg
$file = (Get-Item $arg)
$parent = $file.Directory.Parent.Name
$filename = $file.Name


Write-Host "[PARAMETER] File: $filename"
Write-Host ""


############ ENVIRONMENT DETECTION ############


$file_ninefirst_characters = $file.Name.SubString(0, 9)
echo "Attempt at dircode: $file_ninefirst_characters"
if ( $file_ninefirst_characters -match "20[0-9][0-9]\-[0-9][0-9][0-9][0-9]" )

{
	echo "[CASE DETECTED] Uprooted file, should relocate. Use a GUI.."
	# REBUILT THE WHOLE TREE
	$DIRCODE = $file_ninefirst_characters
	$BASEFOLDER = "M:\9_JOBS_XTRF\"
	$BASEFOLDER = $BASEFOLDER + $file_ninefirst_characters.Substring(0,4) + "_"
	$BASEFOLDER = $BASEFOLDER + $file_ninefirst_characters.Substring(5,2) + "00-" + $file_ninefirst_characters.Substring(5,2) + "99"
	$BASEFOLDER = $BASEFOLDER + "\"

	$BASEFOLDER = $BASEFOLDER + (Get-ChildItem -Path $BASEFOLDER -Directory -Filter "$DIRCODE*") + "\"
	echo "[DETECTED] Base folder: $BASEFOLDER"

	$STUDIO	= $BASEFOLDER + (Get-ChildItem -Path $BASEFOLDER -Filter *.sdlproj -Recurse -ErrorAction SilentlyContinue -Force)
	echo "[DETECTED] Studio: $STUDIO"

	$TO_CLIENT = $BASEFOLDER + (Get-ChildItem -Path $BASEFOLDER -Directory -Filter "*client*")
	echo "[DETECTED] To_Client: $TO_CLIENT"

	#dialog for folder

	#########################################################################
	############## GUI ################


	####### CHECK THE LANGUAGES : Does the file include one ?
	$LCODE_DETECTED = $false
	$allfolder =  Get-ChildItem -Path $STUDIO -Directory -Filter "*-*"
	foreach ($folder in $allfolder) {
		if ($filename.Contains($folder)) {
			echo "LCODE already in - Detected $folder"
			[bool] $LCODE_DETECTED = $true
			$LCODE = $folder } }

	echo "LCODE_DETECTED: $LCODE_DETECTED"



	#########################################################################

}



# NEW FILE, DO EVERYTHING
elseif (($file.FullName.Contains('Downloads') ) -and ( $file.Name.SubString(0, 8) -notmatch "20[0-9][0-9]\-20[0-9][0-9]_*" ) )
{
	$dircode = $file.Name.SubString(0, 8)
	echo "[CASE DETECTED] New file, new project! Dialog to create project here."

	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
	$form = New-Object System.Windows.Forms.Form
	$form.Text = 'Data Entry Form'
	$form.Size = New-Object System.Drawing.Size(280,160)
	$form.StartPosition = 'CenterScreen'
	$form.FormBorderStyle = 'FixedDialog'

	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Location = New-Object System.Drawing.Point(65,90)
	$okButton.Size = New-Object System.Drawing.Size(75,23)
	$okButton.Text = 'Add'
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $okButton
	$form.Controls.Add($okButton)
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Location = New-Object System.Drawing.Point(140,90)
	$cancelButton.Size = New-Object System.Drawing.Size(75,23)
	$cancelButton.Text = 'Cancel'
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $cancelButton
	$form.Controls.Add($cancelButton)

	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10,20)
	$label.Size = New-Object System.Drawing.Size(280,25)
	$label.Text = "Add project code for $filename"
	$form.Controls.Add($label)

	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Location = New-Object System.Drawing.Point(10,50)
	$textBox.Size = New-Object System.Drawing.Size(255,30)
	$form.Controls.Add($textBox)

	$form.Topmost = $true

	$form.Add_Shown({$textBox.Select()})
	$result = $form.ShowDialog()



	###################################### TODO	###################################### TODO
	# MORE ROBUST IF
	if ($result -eq [System.Windows.Forms.DialogResult]::OK)
	{
    		$newprojectname = $textBox.Text
	}
	else { exit }


	###################################### TODO	###################################### TODO
	###################################### TODO	###################################### TODO
	###################################### TODO	###################################### TODO

	# REBUILT THE WHOLE TREE
	$DIRCODE = $newprojectname.SubString(0, 9)
	$BASEFOLDER = "M:\9_JOBS_XTRF\"
	$BASEFOLDER = $BASEFOLDER + $DIRCODE.Substring(0,4) + "_"
	$BASEFOLDER = $BASEFOLDER + $DIRCODE.Substring(5,2) + "00-" + $DIRCODE.Substring(5,2) + "99"
	$BASEFOLDER = $BASEFOLDER + "\"
	$BASEFOLDER = $BASEFOLDER + $newprojectname
	echo "[CREATE] Base folder: $BASEFOLDER"

	# CREATE ALLLLL THE FOLDERS
	echo New-Item -ItemType Directory -Path $BASEFOLDER
	echo New-Item -ItemType Directory -Path $BASEFOLDER\00_info
	echo New-Item -ItemType Directory -Path $BASEFOLDER\01_orig
	echo New-Item -ItemType Directory -Path $BASEFOLDER\02_studio
	echo New-Item -ItemType Directory -Path $BASEFOLDER\03_to\ trans
	echo New-Item -ItemType Directory -Path $BASEFOLDER\04_from\ trans
	echo New-Item -ItemType Directory -Path $BASEFOLDER\05_to\ proof
	echo New-Item -ItemType Directory -Path $BASEFOLDER\06_from\ proof
	echo New-Item -ItemType Directory -Path $BASEFOLDER\07_to\ client


	#rename orig, project code,

	$newname = ($DIRCODE + "_" + $file.Basename + "_orig" + $file.Extension)
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
	
	echo "MOVE TO " + $BASEFOLDER\01_orig
	echo Move-Item -Path $file -Destination $BASEFOLDER\01_orig

	#add to bookmarks
	echo "TODO: Add $BASEFOLDER to bookmarks"


	exit
	###################################### TODO	###################################### TODO
	###################################### TODO	###################################### TODO
	###################################### TODO	###################################### TODO


}


else
{
	echo "[DETECTED] Already in structure."
}


##202[\d]_\d\d\d\d-\d\d\d\d






# If 
# If has a code and in DL
# If has no code and in DL

exit



############ EXCUTION ############


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
# IF NO PROJECT CODE, ADD PROJECT CODE

if ( $file.Name.SubString(0, 8) -ne $dircode )        #-and (! $file.FullName.Contains("_Final") ) )
{
	echo "[DETECTED] $file.Name : No directory code. Adding"
        $newname = ($dircode + "_" + $file.Name)

	echo " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
}


############
# IF ORIG, AND NO ORIG IN THE NAME, ADD ORIG


if (( $file.DirectoryName.Contains("_orig") ) -and (! $file.Name.Contains("_orig") ))
{
	echo "[DETECTED] Original file but no _orig detected. Adding."
	$newname = ($file.Basename + "_orig" + $file.Extension)

	echo " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
}



############
# IF IN COUNTRY CODE FOLDER, ADD COUNTRY CODE, ADD FINAL, MOVE IN FINAL

if (( $file.DirectoryName -match "[A-Z][A-Z]\-[A-Z][A-Z]" ) -and (! $file.Name.Contains("_Final") ))
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
	# echo "destination: $destination" 
	# Move-Item -Path $path_to_file -Destination $destination
	#exit
#}


############
# DEBUG

pause