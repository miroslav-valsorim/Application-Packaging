$users = (Get-Childitem "$envHomeDrive\users").name | where{$_.name -notmatch 'Public|default'}

Foreach ($user in $users){

Remove-Item -Path "$envHomeDrive\users\$user\AppData\Roaming\testpapka" -Recurse -Force

}