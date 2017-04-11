param([String]$policy=30)


If ($policy -eq 'restricted'){
 Set-ExecutionPolicy restricted
 Write-Host "restricted: Press any key to continue ..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")}
Else
 {Set-ExecutionPolicy unrestricted
 Write-Host "unrestricted: Press any key to continue ..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")}