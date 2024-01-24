


# If this one is set, ask for more info
[bool] $NeedMoreInfo = $true


# If we need more information, the form will need to be bigger
# And OK/Cancel buttons will need to be lower.
if ($NeedMoreInfo) { $OFFSET = 155 }
else { $OFFSET = 0 }



## Create a window
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = 'FixedDialog'

$form.Text = 'Tag - $File'
$form.Size = New-Object System.Drawing.Size(300,(200 + $OFFSET))
$form.StartPosition = 'CenterScreen'
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75, (140 + $OFFSET) )
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'Add'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150, (140 + $OFFSET))
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)


############ STANDARD FORM
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Language code to append:'
$form.Controls.Add($label)
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 100


# Get all codes
$allfolder =  Get-ChildItem -Path . -Directory -Filter "*-*"
foreach ($folder in $allfolder)
{ [void] $listBox.Items.Add($folder.Name) }
$form.Controls.Add($listBox)



########### IF RELEVANT
if ($NeedMoreInfo)
{

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,150)
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.Text = 'Folder to move it to:'
    $form.Controls.Add($label2)

    $listBox2 = New-Object System.Windows.Forms.ListBox
    $listBox2.Location = New-Object System.Drawing.Point(10,170)
    $listBox2.Size = New-Object System.Drawing.Size(260,120)
    $listBox2.Height = 120


    # Get all subdirectories
    $allfolder =  Get-ChildItem -Path .. -Directory
    foreach ($folder in $allfolder)
    { [void] $listBox2.Items.Add($folder) }
    $form.Controls.Add($listBox2) ######IF RELEVANT
}



# RESULT PROCESSING
$form.Topmost = $true
$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $listBox.SelectedItem
    $listBox2.SelectedItem #####IF RELEVANT
   

}
else {exit}