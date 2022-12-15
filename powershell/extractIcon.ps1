[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Position = 0 , HelpMessage = "Specify the path to the file.")]
    [ValidateScript({Test-Path $_})]
    [string]$Path,

    [Parameter(Position = 1 , HelpMessage = "Specify the destination file to save the image.")]
    [string]$Destination,

    [Parameter(Position = 2 , HelpMessage = "What is the image index in the file?")]
    [int]$Index,

    [Parameter(Position = 3 , HelpMessage = "What size do you want to use?")]
    [int]$Size,

    [Parameter(Position =4 , HelpMessage = "What image format do you want for output?")]
    [ValidateSet("ico","bmp","png","jpg","gif")]
    [string]$Format,

    [Switch]$Stdin
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

function Extract-Icon {
	Param(
		[Parameter(Position = 0 , HelpMessage = "Specify the path to the file.")]
		[ValidateScript({Test-Path $_})]
		[string]$Path,

		[Parameter(Position = 1 , HelpMessage = "Specify the destination file to save the image.")]
		[string]$Destination,

		[Parameter(Position = 2 , HelpMessage = "What is the image index in the file?")]
		[int]$Index,

		[Parameter(Position = 3 , HelpMessage = "What size do you want to use?")]
		[int]$Size,

		[Parameter(Position =4 , HelpMessage = "What image format do you want for output?")]
		[ValidateSet("ico","bmp","png","jpg","gif")]
		[string]$Format
	)

	Switch ($Format) {
		"ico" {$imageFormat = "icon"}
		"bmp" {$imageFormat = "Bmp"}
		"png" {$imageFormat = "Png"}
		"jpg" {$imageFormat = "Jpeg"}
		"gif" {$imageFormat = "Gif"}
	}

	$ico = [System.IconExtractor]::Extract($Path, $Index, $Size)

	if ($ico) {
		if ($PSCmdlet.ShouldProcess($Destination, "Extract icon")) {
			$ico.ToBitmap().Save($Destination,$imageFormat)
		}
	}
	else {
		Write-Warning "No associated icon image found in $Path"
	}
}

if ( $Stdin ) {
	$input | ConvertFrom-Json | Foreach-Object {
		echo "Path:" $Path
		echo "Destination:" $Destination
		echo "Index:" $Index
		echo "Size:" $Size
		echo "Format:" $Format
	}
}
else {
	echo "Path:" $Path
	echo "Destination:" $Destination
	echo "Index:" $Index
	echo "Size:" $Size
	echo "Format:" $Format
	
	Extract-Icon -Path $Path -Destination $Destination -Index $Index -Size $Size -Format $Format
}


