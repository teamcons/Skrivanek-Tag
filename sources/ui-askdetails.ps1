

    #======================================================
    #                UI FOR ASKING DETAILS                =
    #======================================================




<# 



 #>




#================================================================
$form                       = New-Object System.Windows.Forms.Form
$form.Text                  = "$APPNAME"
$form.FormBorderStyle       = 'FixedDialog'
$form.StartPosition         = 'CenterScreen'
$form.Size                  = New-Object System.Drawing.Size(265,(340 + $OFFSET))
$form.MaximizeBox           = $false
$form.Topmost               = $true
$form.Icon                  = $icon



#================================================================
$pictureBox                     = new-object Windows.Forms.PictureBox
$pictureBox.Location            = New-Object System.Drawing.Point(20,10)

$fileinfo                       = New-Object System.Windows.Forms.Label
$fileinfo.Size                  = New-Object System.Drawing.Size(200,30)
$fileinfo.Location              = New-Object System.Drawing.Point(55,20)
$fileinfo.Font                  = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Italic)



#================================================================
# WHAT FOLDER TO PUT THAT IN
$label2                         = New-Object System.Windows.Forms.Label
$label2.Size                    = New-Object System.Drawing.Size(215,20)
$label2.Text                    = 'Verschieben nach:'


$listBox2                       = New-Object System.Windows.Forms.ListBox
$listBox2.Size                  = New-Object System.Drawing.Size(215,120)
$listBox2.Height                = 120
$listBox2.SelectedItem          = $listBox2.Items[0]


#================================================================
# ASK IF OPEN IN TRADOS
$CheckIfTrados                  = New-Object System.Windows.Forms.CheckBox
$CheckIfTrados.Size             = New-Object System.Drawing.Size(215,25)
$CheckIfTrados.Text             = "�ffnen in Trados"
$CheckIfTrados.Checked          = $True


#================================================================
# OKCANCEL ETC
$OKButton                       = New-Object System.Windows.Forms.Button
$OKButton.Size                  = New-Object System.Drawing.Size(80,25)
$OKButton.Text                  = 'Verschieben!'
$OKButton.DialogResult          = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton              = $OKButton


$CancelButton                   = New-Object System.Windows.Forms.Button
$CancelButton.Size              = New-Object System.Drawing.Size(80,25)
$CancelButton.Text              = 'N�'
$CancelButton.DialogResult      = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton              = $CancelButton











[void]$form.controls.add($pictureBox)
[void]$form.Controls.Add($fileinfo)

[void]$form.Controls.Add($label2)
[void]$form.Controls.Add($listBox2) ######IF RELEVANT
[void]$form.Controls.Add($CheckIfTrados)


[void]$form.Controls.Add($OKButton)
[void]$form.Controls.Add($CancelButton)