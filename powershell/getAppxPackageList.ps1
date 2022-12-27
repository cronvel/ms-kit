
Get-AppxPackage | Where-Object {
	! $_.IsFramework -and ! $_.IsResourcePackage -and ! $_.PackageFullName.Contains( "_neutral_" ) 
} | ForEach-Object {
	$manifest = Get-AppxPackageManifest $_
	
	$application = $manifest.Package.Applications.Application
	
	$application = $application | Where-Object { $_.Id -eq "App" }
	
	if ( $application ) {
		@{
			appName = $_.Name
			packageFullName = $_.PackageFullName
			packageFamilyName = $_.PackageFamilyName
			installLocation = $_.InstallLocation
			version = $_.Version
			iconPath = Join-Path -Path $_.InstallLocation -ChildPath $manifest.package.properties.logo
			executable = Join-Path -Path $_.InstallLocation -ChildPath $application.Executable
			entryPoint = $application.EntryPoint
			displayName = $manifest.Package.Properties.DisplayName
		}
	}
} | ConvertTo-Json

