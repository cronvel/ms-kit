
# Get current appx package list
$rawList = Get-AppxPackage

# Filter, extends and convert to regular array of hashtable.
$list = @();
$rawList | ForEach-Object {
	# First, filter
	# Not useful to filter for IsBundle -eq $false since the first command already filter that out
	#if ( ($_.IsFramework -eq $true) -or ($_.IsResourcePackage -eq $true) ) { continue; }
	if ( ( $_.IsFramework -eq $false ) -and ( $_.IsResourcePackage -eq $false ) -and ( -not ( $_.PackageFullName.Contains( "_neutral_" ) ) ) ) {
		$manifest = Get-AppxPackageManifest $_;
		$element = @{};
		$current = $_;

		# Convert to a regular hashtable
		$current | Get-Member -MemberType *Property | ForEach-Object { $element.($_.name) = $current.($_.name); } 
		
		# Remove unwanted properties
		$element.remove("Dependencies");
		$element.remove("PackageUserInformation");
		$element.remove("ResourceId");

		# Add properties from the manifest
		$element.IconPath = Join-Path -Path $element.InstallLocation -ChildPath $manifest.package.properties.logo;

		$list += $element;
	}
}

$list | ConvertTo-Json
#$list.count

