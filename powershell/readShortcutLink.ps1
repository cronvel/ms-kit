Param(
    [Parameter(Mandatory=$true,ValueFromRemainingArguments=$true,HelpMessage = "Specify the path to the file.")]
    [ValidateScript({Test-Path $_})]
    [string[]]$pathList=@()
)

$shell = New-Object -ComObject WScript.Shell

$pathList | ForEach-Object {
	$shell.CreateShortcut($_)
} | ConvertTo-Json

