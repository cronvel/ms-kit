Param(
	[Parameter(HelpMessage = "Specify the path to the file.")]
	[ValidateScript({Test-Path $_})]
	[string]$Path,

	[Switch]$Stdin,

	[Parameter(ValueFromPipeline)]$input
)

$shell = New-Object -ComObject WScript.Shell

if ( $Stdin ) {
	$input | ConvertFrom-Json | Foreach-Object {
		# This is ONE input, each input being a JSON
		$_ | Foreach-Object {
			$shell.CreateShortcut($_)
		}
	} | ConvertTo-Json
}
else {
	$shell.CreateShortcut($Path) | ConvertTo-Json
}

