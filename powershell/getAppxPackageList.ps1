
Get-AppxPackage | Where-Object {
	! $_.IsFramework -and ! $_.IsResourcePackage -and ! $_.PackageFullName.Contains( "_neutral_" )
} | ForEach-Object {
	$manifest = Get-AppxPackageManifest $_
	@{
		Name = $_.Name
		PackageFullName = $_.PackageFullName
		PackageFamilyName = $_.PackageFamilyName
		InstallLocation = $_.InstallLocation
		Version = $_.Version
		IconPath = Join-Path -Path $_.InstallLocation -ChildPath $manifest.package.properties.logo
	}
} | ConvertTo-Json


