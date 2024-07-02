function Copy-Directories 
			{
				param (
					[parameter(Mandatory = $true)] [string] $source,
					[parameter(Mandatory = $true)] [string] $destination        
				)

				try
				{
					Get-ChildItem -Path $source -Recurse -Force |
						Where-Object { $_.psIsContainer } |
						ForEach-Object { $_.FullName -replace [regex]::Escape($source), $destination } |
						ForEach-Object { $null = New-Item -ItemType Container -Path $_ }

					Get-ChildItem -Path $source -Recurse -Force |
						Where-Object { -not $_.psIsContainer } |
						Copy-Item -Force -Destination { $_.FullName -replace [regex]::Escape($source), $destination }
				}

				catch
				{
					Write-Host "$_"
				}
			}

		$source = "$env:WinDir\testpapka"
		$dest = "$env:AppData"

		Copy-Directories $source $dest