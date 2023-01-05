
Get-AppxPackage | Where-Object {
	! $_.IsFramework -and ! $_.IsResourcePackage -and ! $_.PackageFullName.Contains( "_neutral_" ) 
} | ForEach-Object {
	$appx = $_
	$manifest = Get-AppxPackageManifest $appx
	
	$applicationList = $manifest.Package.Applications.Application
	$application = $applicationList | Where-Object { $_.Id -eq "App" }
	
	if ( $application ) {
		@{
			appName = $appx.Name
			subAppId = $application.Id
			packageFullName = $appx.PackageFullName
			packageFamilyName = $appx.PackageFamilyName
			installLocation = $appx.InstallLocation
			version = $appx.Version
			iconPath = Join-Path -Path $appx.InstallLocation -ChildPath $manifest.package.properties.logo
			executable = Join-Path -Path $appx.InstallLocation -ChildPath $application.Executable
			#entryPoint = $application.EntryPoint	# Seems useless and makes no sense
			displayName = $manifest.Package.Properties.DisplayName
		}
	}
	else {
		#$done = $false

		$applicationList | ForEach-Object {
			$application = $_
			$id = $application.Id
			
			$displayName = $manifest.Package.Properties.DisplayName
			$altName = $application.VisualElements.DisplayName
			if ( $altName -and $altName -notLike "ms-resource:*" ) {
				$displayName = $altName
			}

			if ( $id -and $id -ne "PackageMetadata" -and $id -notLike "Global.*"  -and $id -notLike "*update*" ) {
				@{
					appName = $appx.Name
					subAppId = $application.Id
					packageFullName = $appx.PackageFullName
					packageFamilyName = $appx.PackageFamilyName
					installLocation = $appx.InstallLocation
					version = $appx.Version
					iconPath = Join-Path -Path $appx.InstallLocation -ChildPath $manifest.package.properties.logo
					executable = Join-Path -Path $appx.InstallLocation -ChildPath $application.Executable
					#entryPoint = $application.EntryPoint
					displayName = $displayName
				}
				
				#$done = $true
			}
		}
	}
} | ConvertTo-Json

