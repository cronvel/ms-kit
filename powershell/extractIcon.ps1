#requires -version 5.1

# inspired by https://community.spiceworks.com/topic/592770-extract-icon-from-exe-powershell

[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Position = 0, Mandatory,HelpMessage = "Specify the path to the file.")]
    [ValidateScript({Test-Path $_})]
    [string]$path,

    [Parameter(Position = 1, Mandatory,HelpMessage = "Specify the destination file to save the image.")]
    [string]$destination,

    [Parameter(Position = 2, Mandatory,HelpMessage = "What is the image index in the file?")]
    [int]$imageIndex,

    [Parameter(Position = 3, Mandatory,HelpMessage = "What size do you want to use?")]
    [int]$size,

    [Parameter(Position = 4, Mandatory,HelpMessage = "What format do you want to use?")]
    [ValidateSet("ico","bmp","png","jpg","gif")]
    [string]$format
)

$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System {
	public class IconExtractor {
		public static Icon Extract(string file, int number, int size) {
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

Try {
	Add-Type -AssemblyName System.Drawing -ErrorAction Stop
}
Catch {
	Write-Warning "Failed to import System.Drawing"
	Throw $_
}

Switch ($format) {
	"ico" {$imageFormat = "icon"}
	"bmp" {$imageFormat = "Bmp"}
	"png" {$imageFormat = "Png"}
	"jpg" {$imageFormat = "Jpeg"}
	"gif" {$imageFormat = "Gif"}
}

Write-Verbose "Processing $($file.fullname)"

Write-Verbose "Extracting $imageFormat image to $destination"
$ico = [System.IconExtractor]::Extract($path, $imageIndex, $size)

if ($ico) {
	#WhatIf (target, action)
	if ($PSCmdlet.ShouldProcess($destination, "Extract icon")) {
		$ico.ToBitmap().Save($destination,$imageFormat)
		#Get-Item -path $destination
	}
}
else {
	#this should probably never get called
	Write-Warning "No associated icon image found in $path"
}

Write-Verbose "Ending $($MyInvocation.MyCommand)"

