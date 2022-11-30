Param(
    [Parameter(Position = 0, Mandatory,HelpMessage = "Specify the path to the file.")]
    [ValidateScript({Test-Path $_})]
    [string]$path
)

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($path) 
$shortcut | ConvertTo-Json
