#requires -version 5.1

# inspired by https://community.spiceworks.com/topic/592770-extract-icon-from-exe-powershell

[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Position = 0, Mandatory,HelpMessage = "Specify the path to the file.")]
    [ValidateScript({Test-Path $_})]
    [string]$Path,

    [Parameter(HelpMessage = "Specify the folder to save the file.")]
    [ValidateScript({Test-Path $_})]
    [string]$Destination = ".",

    [parameter(HelpMessage = "Specify an alternate base name for the new image file. Otherwise, the source name will be used.")]
    [ValidateNotNullOrEmpty()]
    [string]$Name,

    [Parameter(HelpMessage = "What format do you want to use? The default is png.")]
    [ValidateSet("ico","bmp","png","jpg","gif")]
    [string]$Format = "png"
)

$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, int size)
	 {
	  IntPtr icon;
	  IntPtr trashIcon;
	  if ( SHDefExtractIcon(file, number, 0, out icon, out trashIcon, size) >= 0 ) {
		return Icon.FromHandle( icon );
	  }
	  else {
		return null;
	  }

	

	 }
	 [DllImport("Shell32.dll", EntryPoint = "SHDefExtractIconW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int SHDefExtractIcon(string sFile, int iIndex, int uFlags, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int size);

	}
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

Write-Verbose "Starting $($MyInvocation.MyCommand)"

Try {
	Add-Type -AssemblyName System.Drawing -ErrorAction Stop
}
Catch {
	Write-Warning "Failed to import System.Drawing"
	Throw $_
}

Switch ($format) {
	"ico" {$ImageFormat = "icon"}
	"bmp" {$ImageFormat = "Bmp"}
	"png" {$ImageFormat = "Png"}
	"jpg" {$ImageFormat = "Jpeg"}
	"gif" {$ImageFormat = "Gif"}
}

$file = Get-Item $path
Write-Verbose "Processing $($file.fullname)"
#convert destination to file system path
$Destination = Convert-Path -path $Destination

if ($Name) {
	$base = $Name
}
else {
	$base = $file.BaseName
}

#construct the image file name
$out = Join-Path -Path $Destination -ChildPath "$base.$format"

Write-Verbose "Extracting $ImageFormat image to $out"
#$ico = [System.IconExtractor]::Extract("C:\Users\cedric\code\powershell\cal.lnk", 0, 0)
$ico = [System.IconExtractor]::Extract("C:\Users\cedric\code\powershell\firefox.exe", 0, 256)
#$ico = [System.IconExtractor]::Extract("shell32.dll", 12, 0)
#$ico = [System.IconExtractor2]::Extract($file.FullName, 1)
#$ico =  [System.Drawing.Icon]::ExtractAssociatedIcon($file.FullName)

#$form = New-Object System.Windows.Forms.Form
#$form.Icon = [System.IconExtractor]::Extract("shell32.dll", 42, $true)
#$form.ShowDialog()

if ($ico) {
	#WhatIf (target, action)
	if ($PSCmdlet.ShouldProcess($out, "Extract icon")) {
		$ico.ToBitmap().Save($Out,$Imageformat)
		Get-Item -path $out
	}
}
else {
	#this should probably never get called
	Write-Warning "No associated icon image found in $($file.fullname)"
}

Write-Verbose "Ending $($MyInvocation.MyCommand)"
