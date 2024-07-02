$Source = "$env:WinDir\testpapka"

$users = (Get-Childitem "$envHomeDrive\users").name | where{$_.name -notmatch 'Public|default'}

Foreach ($user in $users){

Copy-Item -Path $Source -Destination "$envHomeDrive\users\$user\AppData\Roaming" -Recurse -Force

}