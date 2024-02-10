$olApp = new-object -comobject outlook.application
$namespace = $olApp.GetNamespace("MAPI")
$Folders = $namespace.Folders.Item("\\SKRIVANEK Belgium")

#parent folder
$Folder = $Folders.Folders.Item("Inbox")

#$Folder.folders.Add($subfolder)
  
echo $Folder



$olApp.Quit | Out-Null
[GC]::Collect()