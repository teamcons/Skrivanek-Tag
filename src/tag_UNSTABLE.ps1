#================================================================================================================================

#----------------INFO----------------
#
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Project creator for Skrivanek GmbH
#
# Usage add: powershell.exe -executionpolicy bypass -file ".\Rocketlaunch.ps1" "%1"
# Usage: Compiled form, just double-click.
#
#
#----------------STEPS----------------
#
# Initialization
# GUI
# Processing Input
# Build the project
# Bonus
#
#-------------------------------------


#===============================================
#                Initialization                =
#===============================================

# args
param([String]$arg)
if (!$arg) 
{
        Add-Type -AssemblyName System.Windows.Forms

	$ERRORTEXT="No! Usage: Right-click on file to process
->Open With->Program on your PC->This executable"

	$btn = [System.Windows.Forms.MessageBoxButtons]::OK
	$ico = [System.Windows.Forms.MessageBoxIcon]::Information

	Add-Type -AssemblyName System.Windows.Forms 
	[void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)
	exit
}



# Fancy !
Write-Output "======================"
Write-Output "=        TAG!        ="
Write-Output "======================"

Write-Output ""
Write-Output "For Skrivanek GmbH - Manage files really, really quick!"
Write-Output "CC0 Stella Ménier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output "Git: https://github.com/teamcons/Skrivanek-Tag"
Write-Output ""
Write-Output ""





#========================================
# Get all important variables in place 

$APPNAME = "Tag!"
$path_to_file = $arg
$file = (Get-Item $arg)
$parent = $file.Directory.Parent.Name
$filename = $file.Name


Write-Output "[PARAMETER] File: $filename"
Write-Output ""



# Project icon in Base 64, done using
# [Convert]::ToBase64String((Get-Content "..\assets\Skrivanek-Rocketlaunch-Icon.ico" -Encoding Byte))
$iconBase64 = "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"

# Big sentence innit





#======================================================
#                ENVIRONMENT DETECTION                =
#======================================================


# if it has a code
# --> Rebuild tree
# -> If its in Downloads, it is a downloaded package, we may have to move and rename it.

# if it doesnt have any code, it may be a new file we need to build a project for
# or something else. Then idk



$file_ninefirst_characters = $file.Name.SubString(0, 9)
Write-Output "[DETECTED] Attempt at dircode: $file_ninefirst_characters"


#==============
# It if has a code
if ( $file_ninefirst_characters -match "20[0-9][0-9]\-[0-9][0-9][0-9][0-9]" )
{


	# REBUILT THE WHOLE TREE
	$DIRCODE = $file_ninefirst_characters
	$BASEFOLDER = "M:\9_JOBS_XTRF\"
	$BASEFOLDER = -join($BASEFOLDER,$file_ninefirst_characters.Substring(0,4), "_")
	$BASEFOLDER = -join($BASEFOLDER,$file_ninefirst_characters.Substring(5,2),"00-",$file_ninefirst_characters.Substring(5,2),"99")
	$BASEFOLDER = -join($BASEFOLDER,"\")

	$BASEFOLDER = -join($BASEFOLDER,(Get-ChildItem -Path $BASEFOLDER -Directory -Filter "$DIRCODE*"),"\")
	Write-Output "[DETECTED] Base folder: $BASEFOLDER"


    cd "$BASEFOLDER"

    # Grab studio
    $STUDIO	= (Get-ChildItem *trados,*studio).FullName
	Write-Output "[DETECTED] Studio: $STUDIO"


    #Grab ToClient
	$TO_CLIENT = (Get-ChildItem *client).FullName
	Write-Output "[DETECTED] To_Client: $TO_CLIENT"



    if ( $file.Directory.FullName.Contains("Downloads") )
    {
    	Write-Output "[CASE DETECTED] Uprooted file, should relocate. Use a GUI.."



        #==============
	    ####### CHECK THE LANGUAGES : Does the file include one ?
        # By default, consider no.
        # Go in trados folder, and if theres one folder matching, yes.
        # OFFSET is there to leave place for asking. If we detect language code we dont need it


	    [bool]$LCODE_DETECTED = $false
        [int]$OFFSET = 90

        cd $STUDIO
	    $languages =  (dir "*-*" -Directory | Split-Path -leaf)

        # Program sorts out what is in the trados folder, minus non-language code folders
	    foreach ($folder in $languages) {
		    if ($filename.Contains($folder))
            {
			    Write-Output "LCODE already in - Detected $folder"
			    [bool] $LCODE_DETECTED = $true
			    $LCODE = $folder
                $OFFSET = 0
            }
         }
	    Write-Output "LCODE_DETECTED?: $LCODE_DETECTED"




        #==============
        # Create a window
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        # Lets look cool
        [void] [System.Windows.Forms.Application]::EnableVisualStyles() 

        $form = New-Object System.Windows.Forms.Form
        $form.FormBorderStyle = 'FixedDialog'
        $form.StartPosition = 'CenterScreen'
        $form.Text = "$APPNAME"
        $form.Size = New-Object System.Drawing.Size(265,(340 + $OFFSET))
        $form.MaximizeBox = $false

        $iconBytes = [Convert]::FromBase64String($iconBase64)
        # initialize a Memory stream holding the bytes
        $stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
        $form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))



        #==============
        # ICON
        # Remind what file we clickd

        $fileicon = [System.Drawing.Icon]::ExtractAssociatedIcon($file)
        $img = $fileicon.ToBitmap()
    
        $pictureBox = new-object Windows.Forms.PictureBox
        $pictureBox.Location = New-Object System.Drawing.Point(20,10)
        $pictureBox.Width = $img.Size.Width
        $pictureBox.Height = $img.Size.Height
        $pictureBox.Image = $img;
        $form.controls.add($pictureBox)

        $fileinfo = New-Object System.Windows.Forms.Label
        $fileinfo.Location = New-Object System.Drawing.Point(55,20)
        $fileinfo.Size = New-Object System.Drawing.Size(200,30)
        $fileinfo.Text = $file.Name
        $fileinfo.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Italic)
        $null = $form.Controls.Add($fileinfo)

        #==============
        # IF WE DETECTED LCODE, ALL FINE, ELSE OFFSET EVERYTHING AND ADD ALL LANGUAGE CODE


        if ($LCODE_DETECTED -eq $true)
        {
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(20,55)
            $label.Size = New-Object System.Drawing.Size(215,20)
            $label.Text = "Sprachcode detektiert: $LCODE"
            $null = $form.Controls.Add($label)
        }
        else
        {
        
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(20,55)
            $label.Size = New-Object System.Drawing.Size(215,20)
            $label.Text = 'Sprachcode Hinzuf gen:'
            $null = $form.Controls.Add($label)

            $listBox = New-Object System.Windows.Forms.ListBox
            $listBox.Location = New-Object System.Drawing.Point(20,75)
            $listBox.Size = New-Object System.Drawing.Size(215,20)
            $listBox.Height = 90

            # Get all codes + default
            [void] $listBox.Items.Add("(Kein sprachcode, danke!)") 
            $listBox.SelectedItem = "(Kein sprachcode, danke!)"

            foreach ($lang in $languages)
            {
                $null = $listBox.Items.Add($lang)
            }
            $form.Controls.Add($listBox)
        }

        #==============
        # WHAT FOLDER TO PUT THAT IN
        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(20,(80 + $OFFSET))
        $label2.Size = New-Object System.Drawing.Size(215,20)
        $label2.Text = 'Verschieben nach:'
        $form.Controls.Add($label2)


        $listBox2 = New-Object System.Windows.Forms.ListBox
        $listBox2.Location = New-Object System.Drawing.Point(20,(100 + $OFFSET))
        $listBox2.Size = New-Object System.Drawing.Size(215,120)
        $listBox2.Height = 120
        $listBox2.SelectedItem = ""

        # Get all subdirectories
        $allfolder =  Get-ChildItem -Path $BASEFOLDER -Directory
        foreach ($folder in $allfolder)
            { $null = $listBox2.Items.Add($folder) }
        $form.Controls.Add($listBox2) ######IF RELEVANT


        #==============
        # ASK IF OPEN IN TRADOS
        $CheckIfTrados = New-Object System.Windows.Forms.CheckBox        
        $CheckIfTrados.Location = New-Object System.Drawing.Point(20,(220 + $OFFSET))
        $CheckIfTrados.Size = New-Object System.Drawing.Size(215,25)
        $CheckIfTrados.Text = "Öffnen in Trados"
        $CheckIfTrados.UseVisualStyleBackColor = $True
        $CheckIfTrados.Checked = $True
        $form.Controls.Add($CheckIfTrados)

        #==============
        # OKCANCEL ETC
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(40, (255 + $OFFSET) )
        $OKButton.Size = New-Object System.Drawing.Size(80,25)
        $OKButton.Text = 'Verschieben!'
        $okButton.UseVisualStyleBackColor = $True
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $null = $form.Controls.Add($OKButton)
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(135, (255 + $OFFSET))
        $CancelButton.Size = New-Object System.Drawing.Size(80,25)
        $CancelButton.Text = 'Nö'
        $cancelButton.UseVisualStyleBackColor = $True
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)


        #######################
        # RESULT PROCESSING
        $form.Topmost = $true
        $result = $form.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
        {
            # Language is first item. Defaults to empty
            # WHERE is base folder + selected folder. If empty, just root dir
            $LANG = $listBox.SelectedItem
            $WHERE = -join($BASEFOLDER,$listBox2.SelectedItem)

            # No language code ? Empty then
            # Else add underscores to distinguish
            if ($LANG -eq "(Kein sprachcode, danke!)") { $LANG = "" }
            else { $LANG = -join("_",$LANG) }        

            # Get full path
            $newname = -join($file.BaseName,$LANG,$file.Extension)

            Write-Output "[MOVE] $file"
            Write-Output "[MOVE] to $WHERE\$newname"
            Move-Item -Path "$file" -Destination "$WHERE\$newname"


            # If we asked for Trados, open in Trados
            if ($CheckIfTrados.CheckState.ToString() -eq "Checked")
            {
                    Start-Process "$WHERE\$newname"
            }

            explorer $WHERE
            exit
        }
        else {exit } # End of processing Results.




    } #End of If A Download

    #==============
    # Not a download, already there. Nice. Carry on.
    else
    { Write-Output "[DETECTED] Already in structure."}


} # End of Has A Directory Code.

# NEW FILE, DO EVERYTHING
elseif (($file.FullName.Contains('Downloads') ) -and ( $file.Name.SubString(0, 8) -notmatch "20[0-9][0-9]\-20[0-9][0-9]_*" ) )
{
    Write-Output "[DETECTED] NO PROJECT - DO A NEW ONE (TODO)"
    exit

}

# CANNOT DETECT
else
{

	Write-Output "idk lol"
    Write-Output "ninefirst: $file_ninefirst_characters"

    Add-Type -AssemblyName PresentationCore,PresentationFramework
	$msgBody = "idk. ninefirst: $file_ninefirst_characters"
	[System.Windows.MessageBox]::Show($msgBody)

    exit
}



#============================================================
#                INDIVIDUAL FILES PROCESSING                =
#============================================================

#==============
# IF REVIEW JUST APPEND SUFFIX AND STOP

if (( $file.DirectoryName -match "[A-Z][A-Z]\-[A-Z][A-Z]" -and ( $file.Name.EndsWith(".docx.review")) ))
{
	Write-Output "[DETECTED] Bilingual doc."
	$newname = ( $file.DirectoryName + "__" + $file.Name )

	Write-Output " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"

	# move it to correct place
	Write-Output "TAG: Moving to '05_to proof'"
	Move-Item -Path "$file" -Destination "$BASEFOLDER"

	exit
}



#==============
# IGNORE XLIFF

if ( $file.Name.Contains(".sdlxliff") )
{
	Write-Output "[DETECTED] Its an XLIFF ! Ignoring."
	exit
}



#==============
# IF NO PROJECT CODE, ADD PROJECT CODE

if ( $file.Name.SubString(0, 8) -ne $dircode )        #-and (! $file.FullName.Contains("_Final") ) )
{
	Write-Output "[DETECTED] $file.Name : No directory code. Adding"
        $newname = -join($dircode,"_",$file.Name)

	Write-Output " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
}


#==============
# IF ORIG, AND NO ORIG IN THE NAME, ADD ORIG


if (( $file.DirectoryName.Contains("_orig") ) -and (! $file.Name.Contains("_orig") ))
{
	Write-Output "[DETECTED] Original file but no _orig detected. Adding."
	$newname = -join($file.Basename,"_orig",$file.Extension)

	Write-Output " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
}


#==============
# IF IN COUNTRY CODE FOLDER, ADD COUNTRY CODE, ADD FINAL, MOVE IN FINAL

if (( $file.DirectoryName -match "[A-Z][A-Z]\-[A-Z][A-Z]" ) -and (! $file.Name.Contains("_Final") ))
{
	Write-Output "[DETECTED] FINAL. No final tag. Adding."
	$newname = -join($file.Basename.Replace("_orig",""),"_",(Split-Path -Leaf $file.DirectoryName),"_Final",$file.Extension)

	Write-Output " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"

	# And move it
	Write-Output "[ACTION] Moving to $TO_CLIENT"
	Move-Item -Path "$file" -Destination "$TO_CLIENT"
    explorer $TO_CLIENT
	exit

}





#==============
# IF FINAL BUT NO FINAL ADD FINAL

if (( $file.DirectoryName.Contains("_to client") ) -and (! $file.Name.Contains("_Final") ))
{
	Write-Output "[DETECTED] FINAL. No final tag. Unknown LCODE. Adding."
	$newname = ($file.Basename.Replace("_orig","") + "_Final" + $file.Extension)

	Write-Output " -- Renaming $file.Name to $newname"
	Rename-Item -Path "$file" -NewName "$newname"
	$file = Get-Item "$newname"
	exit
}




#==============
# IF PO/CO, SORT OUT

#if $path_to_file.Contains("Downloads")
#{
#	Write-Output "[DETECTED] PO/CO - SORTING OUT"

	# Generate correct dircode
	#$dircode = $dircode.SubString(0, 7)       #$dircode.replace('_','-')
	# $destination = "M:\9_JOBS_XTRF\" + (Get-ChildItem "$dircode.SubString(0, 7)"  )
	# $destination = ( "M:\9_JOBS_XTRF\" + "$dircode" + "\00_info\" )	
	# Write-Output "destination: $destination" 
	# Move-Item -Path $path_to_file -Destination $destination
	#exit
#}