

    #======================================================
    #                UI FOR ASKING DETAILS                =
    #======================================================




<# 



 #>




#================================================================
$form                       = New-Object System.Windows.Forms.Form
$form.Text                  = $APPNAME
$form.FormBorderStyle       = 'FixedDialog'
$form.StartPosition         = 'CenterScreen'

$form.MaximizeBox           = $false
$form.Topmost               = $true
$form.Icon                  = $icon



#================================================================
$pictureBox                     = new-object Windows.Forms.PictureBox
$pictureBox.Location            = New-Object System.Drawing.Point(10,10)

$fileinfo                       = New-Object System.Windows.Forms.Label
$fileinfo.Size                  = New-Object System.Drawing.Size(200,30)
$fileinfo.Location              = New-Object System.Drawing.Point(45,20)
$fileinfo.Font                  = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Italic)


#================================================================
$label                          = New-Object System.Windows.Forms.Label
$label.Size                     = New-Object System.Drawing.Size(220,20)


$listBox                        = New-Object System.Windows.Forms.ListBox
$listBox.Size                   = New-Object System.Drawing.Size(220,20)
$listBox.Height = 90


#================================================================
# WHAT FOLDER TO PUT THAT IN
$label2                         = New-Object System.Windows.Forms.Label
$label2.Size                    = New-Object System.Drawing.Size(220,20)
$label2.Text                    = 'Verschieben nach:'


$listBox2                       = New-Object System.Windows.Forms.ListBox
$listBox2.Size                  = New-Object System.Drawing.Size(220,120)
$listBox2.Height                = 120
$listBox2.SelectedItem          = $listBox2.Items[0]


#================================================================
# ASK IF OPEN IN TRADOS
$CheckIfTrados                  = New-Object System.Windows.Forms.CheckBox
$CheckIfTrados.Size             = New-Object System.Drawing.Size(220,25)
$CheckIfTrados.Text             = "Öffnen in Trados"
$CheckIfTrados.Checked          = $True


#================================================================
# OKCANCEL ETC
$OKButton                       = New-Object System.Windows.Forms.Button
$OKButton.Size                  = New-Object System.Drawing.Size(90,25)
$OKButton.Text                  = 'Verschieben!'
$OKButton.DialogResult          = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton              = $OKButton


$CancelButton                   = New-Object System.Windows.Forms.Button
$CancelButton.Size              = New-Object System.Drawing.Size(90,25)
$CancelButton.Text              = 'Nein'
$CancelButton.DialogResult      = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton              = $CancelButton




[int]$leftalign = 10


$form.Height                      = 320
$label.top                    = 55
$listBox.top                    = 75
$label2.top                    = 80
$listBox2.top                  = 100
$CheckIfTrados.top             = 220
$OKButton.top                  = 255
$CancelButton.top              = 255


$form.Width                     = 255
$label.Left                     = $leftalign
$listBox.Left                   = $leftalign
$label2.Left                    = $leftalign
$listBox2.Left                  = $leftalign
$CheckIfTrados.Left             = $leftalign
$OKButton.Left                  = $leftalign + 15
$CancelButton.Left              = $leftalign + 125





[void]$form.controls.add($pictureBox)
[void]$form.Controls.Add($fileinfo)

[void]$form.Controls.Add($label)

[void]$form.Controls.Add($label2)
[void]$form.Controls.Add($listBox2) ######IF RELEVANT
[void]$form.Controls.Add($CheckIfTrados)

[void]$form.Controls.Add($OKButton)
[void]$form.Controls.Add($CancelButton)