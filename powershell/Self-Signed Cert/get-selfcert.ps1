$newcert = New-SelfSignedCertificate -CertStoreLocation cert:\LocalMachine\my -dnsname SRVRMSSQL10
$Thumb = $newcert.Thumbprint

$pwd=ConvertTo-SecureString "password1" -asplainText -force
$file="C:\temp\srvrmssql10.pfx"
$certpath = "cert:\LocalMachine\My\" + $Thumb
write-host $certpath
break
Export-PFXCertificate -cert $certpath.ToString() -file $file -Password $pwd
Import-PfxCertificate -FilePath $file cert:\LocalMachine\root -Password $pwd
