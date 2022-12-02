#Get-AppxPackage | ConvertTo-Json


#Get-AppxPackage | ForEach-Object { $logo = ( Get-AppxPackageManifest $_).package.properties.logo ; echo @{logo = $logo ; Name = $_.Name } } | ConvertTo-Json 

# Not useful to filter for isBundle -eq $false since the first command already filter that out
$rawList = Get-AppxPackage *calc* | Where-Object {($_.isFramework -eq $false) -and ($_.isResourcePackage -eq $false)}

# This convert the objects into regular hashtable
$list = @();
$rawList | ForEach-Object {
	$element = @{};
	$_ | Get-Member -MemberType *Property | ForEach-Object { $element.($_.name) = $_; } 
	$list += $element;
}

$list | ForEach-Object{ $_.bob= "bill" }

$list
