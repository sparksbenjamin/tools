Function Get-Folder($initialDirectory=""){
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null



$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.Description = "Select a folder"
$foldername.rootfolder = "MyComputer"
$foldername.SelectedPath = $initialDirectory



if($foldername.ShowDialog() -eq "OK")
{
$folder += $foldername.SelectedPath
}
return $folder
}



$cred = Get-Credential
$usracct = $cred.UserName
$pass = $cred.Password
$loc = Get-Folder
$PassFile = "$loc\$usracct.txt"
$KeyFile = "$loc\Key_$usracct.key"
$Key = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
$Key | Out-File $KeyFile
$pass | ConvertFrom-SecureString -Key $Key | Out-File $PassFile
