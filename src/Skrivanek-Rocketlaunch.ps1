#================================================================================================================================

#----------------INFO----------------
#
# CC-BY-SA-NC Stella Ménier <stella.menier@gmx.de>
# Project creator for Skrivanek GmbH
#
# Usage: powershell.exe -executionpolicy bypass -file ".\Rocketlaunch.ps1"
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



# Fancy !
Write-Output "================================"
Write-Output "=        -ROCKETLAUNCH!        ="
Write-Output "================================"

Write-Output ""
Write-Output "For Skrivanek GmbH - Start new projects really, really quick!"
Write-Output "CC0 Stella Ménier, Project manager Skrivanek BELGIUM - <stella.menier@gmx.de>"
Write-Output "Git: https://github.com/teamcons/Skrivanek-Rocketlaunch"
Write-Output ""
Write-Output ""


#========================================
# Get all important variables in place 

Write-Output "[STARTUP] Getting all variables in place"
[string]$APPNAME = "-Rocketlaunch!"
[string]$PROJECTTEMPLATE = "Minimal"
[string]$LOAD_SOURCE_FROM = "$env:USERPROFILE\Downloads\"


# Load templates from a csv in same place as executable
[string]$LOAD_TEMPLATES_FROM = $MyInvocation.MyCommand.Path
[string]$ROOTSTRUCTURE = "M:\9_JOBS_XTRF\"
[regex]$CODEPATTERN = -join($YEAR,"-[0-9]")
[string]$YEAR = get-date –f yyyy

# TRADOS. TODO: MORE FLEXIBLE
[string]$NEWPROJECTICON = "C:\Program Files (x86)\Trados\Trados Studio\Studio17\StudioTips\de\00_WelcomeFlow\NewProject.ico"



#==========================================
# Try to predict what next number would be 
# Catch: have at least first part of code
try
{
    # 
    cd $ROOTSTRUCTURE
    cd (dir 2024_* -Directory | Select-Object -Last 1)   
    $PREDICT_CODE =  (dir -Directory | Select-Object -Last 1).Name.Substring(5,4)
    $PREDICT_CODE =  [int]$PREDICT_CODE + 1
    [string]$PREDICT_CODE =  -join($YEAR,"-",$PREDICT_CODE,"_")
    Write-Output "[PREDICTED] Next is $PREDICT_CODE"
}
catch
{
    $PREDICT_CODE = -join($YEAR,"-")
}


#================================
# Project icon in Base 64
# [Convert]::ToBase64String((Get-Content "..\assets\Skrivanek-Rocketlaunch-Icon.ico" -Encoding Byte))
[string]$iconBase64 = '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'

# phew


#==============================================================
#                GUI - Ask the Right Questions                =
#==============================================================


#= Layout
# Icon, name, input field
# Listbox
# Checkboxes
#=

#================
#= INITIAL WORK =

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Lets look cool
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

$form = New-Object System.Windows.Forms.Form
$form.Text = $APPNAME
$form.Size = New-Object System.Drawing.Size(265,375)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Topmost = $True

$iconBytes = [Convert]::FromBase64String($iconBase64)
# initialize a Memory stream holding the bytes
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)
$form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))



#==============
#= INPUT TEXT =

# FANCY ICON


$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(20,10)

$img = [System.Drawing.Image]::Fromfile($NEWPROJECTICON);
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img;

$form.controls.add($pictureBox)




# LABEL AND TEXT

# Label above input
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(90,25)
$label.Size = New-Object System.Drawing.Size(145,20)
$label.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Bold)
$label.Text = 'Projektname:'
$form.Controls.Add($label)

# Input box
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(90,50)
$textBox.Size = New-Object System.Drawing.Size(145,30)
$textBox.Text = $PREDICT_CODE
$form.Controls.Add($textBox)
$form.Add_Shown({$textBox.Select()})


#=====================
#= LIST OF TEMPLATES =

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(20,90)
$label2.Size = New-Object System.Drawing.Size(215,30)
$label2.Text = 'Welche Projektvorlage soll verwendet werden?'
$form.Controls.Add($label2)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(20,120)
$listBox.Size = New-Object System.Drawing.Size(215,60)
$listBox.Height = 120

[void] $listBox.Items.Add('Minimal')
$listBox.SelectedItem = "Minimal"

## LOAD FROM CSV HERE
[void] $listBox.Items.Add('Standard TEP')
[void] $listBox.Items.Add('Provider macht TEP')
[void] $listBox.Items.Add('Sworn Translation')
[void] $listBox.Items.Add('MemoQ')
[void] $listBox.Items.Add('Production')

$form.Controls.Add($listBox)


#===========
#= OPTIONS =


# Check if include original files
$CheckIfSourceFiles = New-Object System.Windows.Forms.CheckBox        
$CheckIfSourceFiles.Location = New-Object System.Drawing.Point(20,240)
$CheckIfSourceFiles.Size = New-Object System.Drawing.Size(215,25)
$CheckIfSourceFiles.Text = "Ausgangsdateien einbeziehen?"
$CheckIfSourceFiles.UseVisualStyleBackColor = $True
$CheckIfSourceFiles.Checked = $True
$form.Controls.Add($CheckIfSourceFiles)


# Check if start new trados project
$CheckIfTrados = New-Object System.Windows.Forms.CheckBox        
$CheckIfTrados.Location = New-Object System.Drawing.Point(20,265)
$CheckIfTrados.Size = New-Object System.Drawing.Size(215,25)
$CheckIfTrados.Text = "Ein neues Trados-Projekt beginnen?"
$CheckIfTrados.UseVisualStyleBackColor = $True
$CheckIfTrados.Checked = $True
$form.Controls.Add($CheckIfTrados)



#====================
#= OKCANCEL BUTTONS =

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(40,300)
$okButton.Size = New-Object System.Drawing.Size(80,25)
$okButton.Text = 'Los!'
$okButton.UseVisualStyleBackColor = $True
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(135,300)
$cancelButton.Size = New-Object System.Drawing.Size(80,25)
$cancelButton.Text = 'Nö'
$cancelButton.UseVisualStyleBackColor = $True
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)



#==============
#= WRAP IT UP =


$result = $form.ShowDialog()

[string]$PROJECTNAME = $textBox.Text 
$PROJECTTEMPLATE = $listBox.SelectedItem

# Cancel culture
# Close if cancel
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        Write-Output "[INPUT] Got Cancel. Aw. Exit."
        exit
    }
Write-Output "[INPUT] Got: $PROJECTNAME"




#==============================================================
#                     Processing Le input                     =
#==============================================================




# Empty, so go on with what was initially predicted
if ("$PROJECTNAME" -notmatch "[0-9]" )
{

    $PROJECTNAME = -join($PREDICT_CODE,$PROJECTNAME)
    Write-Output "Its words. Now: $PROJECTNAME"
}

# Remove invalid character, just in case
$PROJECTNAME = $PROJECTNAME.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
Write-Output "Removed invalid. Now: $PROJECTNAME"



# Code is formed correctly buuuut...
# Person started with invalid char - the slash from XTRF or some shit
# Test for underscore because invalid charactersvebeen stripped
if ($PROJECTNAME -match "^20[0-9][0-9]_[0-9]")
{
    [regex]$pattern = '_'
    $PROJECTNAME = $pattern.replace($PROJECTNAME, '-', 1)
    Write-Output "Slash, replace with dash. Now: $PROJECTNAME"
}





# If it matches the pattern
# We'll probably just have to insert missing zeros, if
if ($PROJECTNAME -match "$CODEPATTERN" )
{

    Write-Output "Case : Wellformed code"
    

    # just check if it is missing zeros
    if ($PROJECTNAME -match "^20[0-9][0-9]-[0-9][0-9][0-9]")
    {
        $PROJECTNAME = $PROJECTNAME.Insert(5,"0")
        Write-Output "Missing first zero. Now: $PROJECTNAME"

    }
    elseif ($PROJECTNAME -match "^20[0-9][0-9]\-[0-9][0-9]")
    {
        $PROJECTNAME = $PROJECTNAME.Insert(5,"00")
        Write-Output "Missing two zero. Now: $PROJECTNAME"
    }
    elseif ($PROJECTNAME -match "^20[0-9][0-9]\-[0-9]")
    {
        $PROJECTNAME = $PROJECTNAME.Insert(5,"000")
        Write-Output "Missing three zero. Now: $PROJECTNAME"
    }


}
else { Write-Output "$PROJECTNAME does not match $CODEPATTERN" }

# does NOT Start with year, but with numbers, or slash-numbers.
if ($PROJECTNAME -notmatch "^[0-9][0-9][0-9][0-9]-")
{
    Write-Output "Case : No year, adding year"

    # is it missing zeros
    if ($PROJECTNAME -match "^[0-9][0-9][0-9]")
    {
        $PROJECTNAME = -join("0",$PROJECTNAME)
        Write-Output "Missing first zero. Now: $PROJECTNAME"

    }
    elseif ($PROJECTNAME -match "^[0-9][0-9]")
    {
        $PROJECTNAME = -join("00",$PROJECTNAME)
        Write-Output "Missing two zero. Now: $PROJECTNAME"
    }
    elseif ($PROJECTNAME -match "^[0-9]")
    {
        $PROJECTNAME = -join("000",$PROJECTNAME)
        Write-Output "Missing three zero. Now: $PROJECTNAME"
    }


    # Add year
    $PROJECTNAME = -join($YEAR,"-",$PROJECTNAME)


}


###########################DEBUG
#Write-Output "DEBUG"
#Write-Output "IS CORRECT ?"
#Write-Output "$PROJECTNAME"
#exit
###########################DEBUG


##### Ultimate check
try { $DIRCODE = $PROJECTNAME.SubString(0, 9) }
catch {
	$ERRORTEXT="Projektcode ist unpassend !!!
Format: 20[0-9][0-9]\-[0-9][0-9][0-9][0-9] + Name
Angegeben: $PROJECTCODE"

	$btn = [System.Windows.Forms.MessageBoxButtons]::OK
	$ico = [System.Windows.Forms.MessageBoxIcon]::Information

	Add-Type -AssemblyName System.Windows.Forms 
	[void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)

exit }




## ULTIMATE ULTIMATE CHECK
if (($result -eq [System.Windows.Forms.DialogResult]::OK) -and ($DIRCODE -match "^20[0-9][0-9]\-[0-9][0-9][0-9][0-9]" ))
{

Write-Output "[INPUT] Project Name is $PROJECTNAME"
Write-Output "[INPUT] Project Template is $PROJECTTEMPLATE"

Write-Output "[INPUT] Include Source Files ?" $CheckIfSourceFiles.CheckState
Write-Output "[INPUT] Start new Trados project ?" $CheckIfTrados.CheckState

Write-Output "[DETECTED] Dircode is $DIRCODE"

}
else { 	$ERRORTEXT="Projektcode oder vorlage ist unpassend !!!
Format: 20[0-9][0-9]\-[0-9][0-9][0-9][0-9] + Name
Angegeben: $PROJECTCODE"

	$btn = [System.Windows.Forms.MessageBoxButtons]::OK
	$ico = [System.Windows.Forms.MessageBoxIcon]::Information

	Add-Type -AssemblyName System.Windows.Forms 
	[void] [System.Windows.Forms.MessageBox]::Show($ERRORTEXT,$APPNAME,$btn,$ico)

exit  }




#==============================================================
#                      Build The Project                      =
#==============================================================


# REBUILT THE WHOLE TREE
# Its in... year, underscore
$BASEFOLDER = -join($ROOTSTRUCTURE,$DIRCODE.Substring(0,4),"_")
$BASEFOLDER = -join($BASEFOLDER,$DIRCODE.Substring(5,2),"00-",$DIRCODE.Substring(5,2),"99")

# If the folder with project numbers in range do not exist, just create it lol
if (!(Test-Path $BASEFOLDER -PathType Container)) {
    Write-Output "[CREATE] Range folder in tree: $BASEFOLDER"
    New-Item -ItemType Directory -Force -Path "$BASEFOLDER"
}
$BASEFOLDER = -join($BASEFOLDER,"\",$PROJECTNAME)





Write-Output "[CREATE] Base folder: $BASEFOLDER"
New-Item -ItemType Directory -Path "$BASEFOLDER"


switch ( $PROJECTTEMPLATE )
{
    "Minimal"
        { $FOLDERS = @("00_info", "01_orig" ) }
    "Standard TEP"
        { $FOLDERS = @("00_info", "01_orig", "02_trados", "03_to trans", "04_from trans", "05_to proof", "06_from proof", "07_to client") }
    "Provider macht TEP"
        { $FOLDERS = @("00_info", "01_orig", "02_trados", "03_to TEP", "04_from TEP", "05_to client") }
    "Sworn Translation"
        { $FOLDERS = @("00_info", "01_orig", "02_to client") }
    "MemoQ"
        { $FOLDERS = @("00_info", "01_orig", "02_memoQ", "03_to client") }
    "Production"
        { $FOLDERS = @("00_info") }
}


# CREATE ALLLLL THE FOLDERS
foreach ($folder in $FOLDERS)
{
    Write-Output "[CREATE] folder: $BASEFOLDER\$folder"
    New-Item -ItemType Directory -Path "$BASEFOLDER\$folder"
}




# PIN TO EXPLORER
$o = new-object -com shell.application
$o.Namespace($BASEFOLDER).Self.InvokeVerb("pintohome")







#==================================================
#                      BONUS                      =
#==================================================



#==========================
#= INCLUDE ORIGINAL FILES =


# If user asked to include source files, include those in new folder, with naming conventions
if ($CheckIfSourceFiles.CheckState.ToString() -eq "Checked")
{

    Write-Output "[DETECTED] Load source files"



    # Grab source files
    $SOURCEFILES = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = $LOAD_SOURCE_FROM
        Multiselect = $true
        Title = $APPNAME
    }
    $null = $SOURCEFILES.ShowDialog()
    Write-Output "[INPUT] Got:"
    Write-Output $SOURCEFILES.FileNames



    # Rename and move source files
    foreach ($file in $SOURCEFILES.FileNames)
    {

        $truefile = Get-Item "$file"
        $newname = -join($DIRCODE,"_",$truefile.BaseName,"_orig",$truefile.Extension)

        # Move to orig if theres one. Else just at root of folder
        if (Test-Path "$BASEFOLDER\01_orig\" -PathType Container)
        {
            Write-Output "[MOVE] Move to $BASEFOLDER\01_orig\$newname"
            Move-Item -Path "$truefile" -Destination "$BASEFOLDER\01_orig\$newname"
        }
        else
        {
            Write-Output "[MOVE] Move to $BASEFOLDER\$newname"
            Move-Item -Path "$truefile" -Destination "$BASEFOLDER\$newname"
        }

    }




}



#==========================
#= START A TRADOS PROJECT =


# If user asked for trados, start it and fill what we can
if ($CheckIfTrados.CheckState.ToString() -eq "Checked")
{
	Write-Output "Starting Trados Studio..."

    # May not be where expected
    try {
        cd "C:\Program Files (x86)\Trados\Trados Studio\Studio17"
        }
    catch { 
        $TRADOSDIR = (Get-ChildItem -Path "C:\Program Files (x86)" -Filter *.sdlproj -Recurse -ErrorAction SilentlyContinue -Force -File).Directory.FullName
        Write-Output "[DETECTED] Trados in $TRADOSDIR"
        cd $TRADOSDIR
        }

    .\SDLTradosStudio.exe /createProject /name $PROJECTNAME



}


#=============
#= LAST STEP =

# OK NOW WE WORK
Write-Output "Starting Explorer..."
start-process explorer "$BASEFOLDER"

