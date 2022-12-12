Param(
    [Parameter(Position=0,Mandatory,HelpMessage = "Specify the path to the file.")]
    [ValidateScript({Test-Path $_})]
    [string[]]$pathList=@()
)

$shell = New-Object -ComObject WScript.Shell

$pathList | ForEach-Object {
	$shell.CreateShortcut($_)
} | ConvertTo-Json

