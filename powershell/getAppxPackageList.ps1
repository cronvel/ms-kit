
$exit = 0

Get-AppxPackage | Where-Object {
	! $_.IsFramework -and ! $_.IsResourcePackage -and ! $_.PackageFullName.Contains( "_neutral_" )
} | ForEach-Object {
	$manifest = Get-AppxPackageManifest $_
	
	$application = $manifest.Package.Applications.Application
	
	$application = $application | Where-Object { $_.Id -eq "App" }
	
	if ( ! $exit -and $application ) {
		@{
			Name = $_.Name
			PackageFullName = $_.PackageFullName
			PackageFamilyName = $_.PackageFamilyName
			InstallLocation = $_.InstallLocation
			Version = $_.Version
			IconPath = Join-Path -Path $_.InstallLocation -ChildPath $manifest.package.properties.logo
			Executable = Join-Path -Path $_.InstallLocation -ChildPath $application.Executable
			EntryPoint = $application.EntryPoint
		}
		#$exit = 1
	}
} | ConvertTo-Json


