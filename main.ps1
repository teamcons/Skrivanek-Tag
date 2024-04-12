#================================================================================================================================

#----------------INFO----------------
#
# CC-BY-SA-NC Stella M�nier <stella.menier@gmx.de>
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
        $arg = "C:\Users\stella.Menier\Downloads\2024-0829_Yanmar_FR.sdlrpx"
<# 	$ERRORTEXT="No! Usage: Right-click on file to process
->Open With->Program on your PC->This executable"

	$btn = [System.Windows.Forms.MessageBoxButtons]::OK
	$ico = [System.Windows.Forms.MessageBoxIcon]::Information

	Add-Type -AssemblyName System.Windows.Forms 
	[void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)
	exit #>
}

# Fancy !
Write-Output "======================"
Write-Output "=        TAG!        ="
Write-Output "======================"

Write-Output ""
Write-Output "For Skrivanek GmbH - Manage files really, really quick!"
Write-Output "CC0 Stella M�nier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output "Git: https://github.com/teamcons/Skrivanek-Tag"
Write-Output ""
Write-Output ""


#========================================
# Grab script location in a way that is compatible with PS2EXE
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $global:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
else
    {$global:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
    if (!$ScriptPath){ $global:ScriptPath = "." } }




#========================================
# Get all important variables in place 

$APPNAME = "Tag!"
$file = (Get-Item $arg)

Write-Output "[PARAMETER] File: $file.Name"
Write-Output ""






#========================================
# Get all resources

# Allow having a fancing GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore,PresentationFramework
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

# Load assets
$script:icon                = New-Object system.drawing.icon $ScriptPath\assets\stamp.ico

Import-Module $ScriptPath\sources\ui-askdetails.ps1



#======================================================
#                ENVIRONMENT DETECTION                =
#======================================================


# if it has a code
# --> Rebuild tree
# -> If its in Downloads, it is a downloaded package, we may have to move and rename it.

# if it doesnt have any code, it may be a new file we need to build a project for
# or something else. Then idk







#==============
# It if has a code
if ( $file.Name -match "^20[0-9][0-9]\-[0-9][0-9][0-9][0-9]" )
{
    Write-Output "[DETECTED] Has a dircode"

	# REBUILT THE WHOLE TREE
    $file_ninefirst_characters = $file.Name.SubString(0, 9)
    Write-Output "[DETECTED] Attempt at dircode: $file_ninefirst_characters"

	$DIRCODE = $file_ninefirst_characters
	$BASEFOLDER = "M:\9_JOBS_XTRF\"
	$BASEFOLDER = -join($BASEFOLDER,$file_ninefirst_characters.Substring(0,4), "_")
	$BASEFOLDER = -join($BASEFOLDER,$file_ninefirst_characters.Substring(5,2),"00-",$file_ninefirst_characters.Substring(5,2),"99")
	$BASEFOLDER = -join($BASEFOLDER,"\")

	$BASEFOLDER = -join($BASEFOLDER,(Get-ChildItem -Path $BASEFOLDER -Directory -Filter "$DIRCODE*"),"\")
	Write-Output "[DETECTED] Base folder: $BASEFOLDER"
    Set-Location $BASEFOLDER


    # Grab info 
    $INFO	= (Get-ChildItem *info | Select-Object -First 1).Name
    try { Test-Path -Path "$BASEFOLDER\$INFO" }
    catch { $INFO = $BASEFOLDER }
    Write-Output "[DETECTED] Info: $INFO"

    # Grab orig
    $ORIG	= (Get-ChildItem *orig | Select-Object -First 1).Name
    try { Test-Path -Path "$ORIG"}
    catch {$ORIG = $BASEFOLDER}
    Write-Output "[DETECTED] Orig: $ORIG"

    # Grab studio
    $STUDIO	= (Get-ChildItem *trados,*studio | Select-Object -First 1).FullName


    # Sometimes the files are in a subfolder and theres several trados folders.
    #So if no SDL projet not immediately there, dig deeper
    if ((Get-ChildItem -Path $STUDIO -Filter "*.sdlproj") -eq $none )
    {
        $STUDIO	= (Get-ChildItem  -Path $STUDIO -Filter (-join($file.BaseName,"*.sdlproj")) -Recurse).Directory.FullName
    }

    # If the script cant find for shit, just grab basefolder
    try {Test-Path -Path "$BASEFOLDER\$STUDIO"}
    catch {$STUDIO = $BASEFOLDER}
    Write-Output "[DETECTED] Studio: $STUDIO"

    # Grab client
    $TO_CLIENT	= (Get-ChildItem *client | Select-Object -First 1).Name
    try {Test-Path -Path "$BASEFOLDER\$TO_CLIENT"}
    catch {$TO_CLIENT = $BASEFOLDER}
    Write-Output "[DETECTED] To client: $TO_CLIENT"



    if ( $file.Directory.FullName.Contains("Downloads") )
    {
    	Write-Output "[CASE DETECTED] Uprooted file, should relocate. Use a GUI.."



        #==============
	    ####### CHECK THE LANGUAGES : Does the file include one ?
        # By default, consider no.
        # Go in trados folder, and if theres one folder matching, yes.
        # OFFSET is there to leave place for asking. If we detect language code we dont need it


	    [bool]$LCODE_DETECTED = $false


        Set-Location $STUDIO

        # sometimes its in a subfolder
        if (Test-Path "*.sdlproj" )
            {
            $languages =  (Get-ChildItem "*-*" -Directory | Split-Path -leaf)
            }
        else {
            Set-Location (Get-ChildItem)
            $languages =  (Get-ChildItem "*-*" -Directory | Split-Path -leaf)
            <# Action when all if and elseif conditions are false #>
        }

        # Program sorts out what is in the trados folder, minus non-language code folders
	    foreach ($folder in $languages) {
		    if ($file.Name.Contains($folder))
            {
			    Write-Output "LCODE already in - Detected $folder"
			    [bool] $LCODE_DETECTED = $true
			    $LCODE = $folder
            }
         }
	    Write-Output "LCODE_DETECTED?: $LCODE_DETECTED"


        #==============
        # ICON
        # Remind what file we clickd

        $fileicon               = [System.Drawing.Icon]::ExtractAssociatedIcon($file)
        $img                    = $fileicon.ToBitmap()
        $pictureBox.Width       = $img.Size.Width
        $pictureBox.Height      = $img.Size.Height
        $pictureBox.Image       = $img;
        $fileinfo.Text          = $file.Name


        #==============
        # IF WE DETECTED LCODE, ALL FINE, ELSE OFFSET EVERYTHING AND ADD ALL LANGUAGE CODE
        if ($LCODE_DETECTED -eq $true)
        {
            # We detektened it
            $label.Text                     = "Sprachcode detektiert: $LCODE"
        }
        else
        {
            # We dint
            $label.Text                     = 'Sprachcode Hinzufügen:'

            # Make space
            [int]$OFFSET                    = 90
            $form.Height                   += $OFFSET
            $label2.top                    += $OFFSET
            $listBox2.top                  += $OFFSET
            $CheckIfTrados.top             += $OFFSET
            $OKButton.top                  += $OFFSET
            $CancelButton.top              += $OFFSET            

            # Get all codes + default
            [void]$listBox.Items.Add("(Kein sprachcode, danke!)") 
            $listBox.SelectedItem = $listBox.Items[0]

            foreach ($lang in $languages)
            { [void]$listBox.Items.Add($lang)}
            [void]$form.Controls.Add($listBox)
        }

        # Get all subdirectories
        $allfolder =  Get-ChildItem -Path $BASEFOLDER -Directory
        foreach ($folder in $allfolder)
            { $null = $listBox2.Items.Add($folder) }




        #######################
        # RESULT PROCESSING

        $result = $form.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
        {
            # Language is first item. Defaults to empty
            # WHERE is base folder + selected folder. If empty, just root dir
            $LANG = $listBox.SelectedItem
            $WHERE = -join($BASEFOLDER,$listBox2.SelectedItem)

            # No language code ? Empty then
            # Else add underscores to distinguish

            

            # Get full path
            if ($file.Name -match ".review.docx$" )
            {
                if ($LANG -eq "(Kein sprachcode, danke!)") { $LANG = "" }
                else { $LANG = -join($LANG,"__") }

                $newname = -join($LANG,$file.Name)
            }
            else {
                if ($LANG -eq "(Kein sprachcode, danke!)") { $LANG = "" }
                else { $LANG = -join("_",$LANG) }

                $newname = -join($file.BaseName,$LANG,$file.Extension)
            }



            Write-Output "[MOVE] $file"
            Write-Output "[MOVE] to $WHERE\$newname"
            Move-Item -Path "$file" -Destination "$WHERE\$newname"


            # If we asked for Trados, open in Trados
            if ($CheckIfTrados.CheckState.ToString() -eq "Checked")
            {
                   
                Write-Output "Starting Trados Studio..."
                # May not be where expected
                try {
                    Set-Location "C:\Program Files (x86)\Trados\Trados Studio\Studio17"
                    }
                catch { 
                    $TRADOSDIR = (Get-ChildItem -Path "C:\Program Files (x86)" -Filter *.sdlproj -Recurse -ErrorAction SilentlyContinue -Force -File).Directory.FullName
                    Write-Output "[DETECTED] Trados in $TRADOSDIR"
                    Set-Location $TRADOSDIR
                    }
            
                .\SDLTradosStudio.exe /openPackage $WHERE\$newname            
            }

            #explorer $WHERE
            exit
        }
        else {exit } # End of processing Results.




    } #End of If A Download

    #==============
    # Not a download, already there. Nice. Carry on.
    else
    { Write-Output "[DETECTED] Already in structure."
    }


} # End of Has A Directory Code.

# NEW FILE, DO EVERYTHING
elseif (($file.FullName -match "Downloads" ) -and ( $file.Name -notmatch "^20[0-9][0-9]\-20[0-9][0-9]_" ) )
{
    Write-Output "[DETECTED] NO PROJECT - DO A NEW ONE (TODO)"
    exit

}

# FILE IN TREE
elseif ($file.FullName -match "^M:\\9_JOBS_XTRF\\20[0-9][0-9]_[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]\\20[0-9][0-9]\-"  )
{

    Write-Output "[DETECTED] Case : In tree but no dircode"

    $foldername = $file.FullName.ToString().split('\\')[3]
    $DIRCODE = $foldername.SubString(0, 9)

	$BASEFOLDER = "M:\9_JOBS_XTRF\"
	$BASEFOLDER = -join($BASEFOLDER,"\",$file.FullName.ToString().split('\\')[2])
	$BASEFOLDER = -join($BASEFOLDER,"\",$foldername)

	Write-Output "[DETECTED] Base folder: $BASEFOLDER"
	Write-Output "[DETECTED] Dircode: $DIRCODE"

    Set-Location "$BASEFOLDER"

    # Grab info 
    $INFO	= (Get-ChildItem *info | Select-Object -First 1).Name
    try { Test-Path -Path "$BASEFOLDER\$INFO" }
    catch { $INFO = $BASEFOLDER }
    Write-Output "[DETECTED] Info: $INFO"

    # Grab orig
    $ORIG	= (Get-ChildItem *orig | Select-Object -First 1).Name
    try { Test-Path -Path "$ORIG"}
    catch {$ORIG = $BASEFOLDER}
    Write-Output "[DETECTED] Orig: $ORIG"

    # Grab studio
    $STUDIO	= (Get-ChildItem *trados,*studio | Select-Object -First 1).FullName


        # Sometimes the files are in a subfolder and theres several trados folders.
        #So if no SDL projet not immediately there, dig deeper
    if ((Get-ChildItem -Path $STUDIO -Filter "*.sdlproj") -eq $none )
    {
        $STUDIO	= (Get-ChildItem  -Path $STUDIO -Filter (-join($file.BaseName,"*.sdlproj")) -Recurse).Directory.FullName
    }

    # If the script cant find for shit, just grab basefolder
    try {Test-Path -Path "$STUDIO"}
    catch {$STUDIO = $BASEFOLDER}
    Write-Output "[DETECTED] Studio: $STUDIO"

    # Grab client
    $TO_CLIENT	= (Get-ChildItem *client | Select-Object -First 1).Name
    try {Test-Path -Path "$BASEFOLDER\$TO_CLIENT"}
    catch {$TO_CLIENT = $BASEFOLDER}
    Write-Output "[DETECTED] To client: $TO_CLIENT"

}


#============================================================
#                INDIVIDUAL FILES PROCESSING                =
#============================================================

Write-Output "[RESTORE PATH] Ready to go !"
Set-Location $file.Directory.FullName

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

if ( $file.Extension -match ".sdlxliff" )
{
	Write-Output "[DETECTED] Its an XLIFF ! Ignoring."
	exit
}



#==============
# IF NO PROJECT CODE, ADD PROJECT CODE

if ( $file.Name -notmatch $DIRCODE )        #-and (! $file.FullName.Contains("_Final") ) )
{
	Write-Output "[DETECTED] $file.Name : No directory code. Adding"
    $newname = -join($DIRCODE,"_",$file.Name)

	Rename-Item -Path $file.FullName -NewName "$newname"
    $file = Get-Item $newname
}


#==============
# IF ORIG, AND NO ORIG IN THE NAME, ADD ORIG


if (( $file.DirectoryName -match "$ORIG" ) -and (! $file.BaseName.Contains("_orig") ))
{
	Write-Output "[DETECTED] Original file but no _orig detected. Adding."
	$newname = -join($file.Basename,"_orig",$file.Extension)

	Rename-Item -Path $file.FullName -NewName "$newname"
    $file = Get-Item $newname
}


#==============
# IF IN COUNTRY CODE FOLDER, ADD COUNTRY CODE, ADD FINAL, MOVE IN FINAL

if ( (Split-Path -Leaf $file.DirectoryName) -match "[A-Z][A-Z]\-[A-Z][A-Z]" )
{
	Write-Output "[DETECTED] FINAL."

    if (! $file.Name.Contains("_Final") )
    {
        Write-Output "[RENAME] No final tag. Adding."
        $newname = -join($file.Basename.Replace("_orig",""),"_",(Split-Path -Leaf $file.DirectoryName),"_Final",$file.Extension)
        Rename-Item -Path $file.FullName -NewName "$newname"
        $file = Get-Item $newname
    }

	# And move it
	Write-Output "[ACTION] Moving to $TO_CLIENT"
	Move-Item -Path "$file" -Destination "$BASEFOLDER\$TO_CLIENT"

	exit

}





#==============
# IF FINAL BUT NO FINAL ADD FINAL

if (( $file.DirectoryName -match $TO_CLIENT ) -and (! $file.Name.Contains("_Final") ))
{
	Write-Output "[DETECTED] FINAL. No final tag. Unknown LCODE. Adding."
	$newname = -join($file.Basename.Replace("_orig","") + "_Final" + $file.Extension)

	Rename-Item -Path $file.FullName -NewName "$newname"
    $file = Get-Item $newname
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