[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Position = 0 , HelpMessage = "Specify the source file.")]
    [ValidateScript({Test-Path $_})]
    [string]$Source,

    [Parameter(Position = 1 , HelpMessage = "Specify the destination file to save the image.")]
    [string]$Destination,

    [Parameter(Position = 2 , HelpMessage = "What is the image index in the file?")]
    [int]$Index,

    [Parameter(Position = 3 , HelpMessage = "What size do you want to use?")]
    [int]$Size,

    [Parameter(Position =4 , HelpMessage = "What image format do you want for output?")]
    [ValidateSet("ico","bmp","png","jpg","gif")]
    [string]$Format,

    [Switch]$Stdin,

	[Parameter(ValueFromPipeline)]$input
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
		[Parameter(Position = 0 , HelpMessage = "Specify the source file.")]
		[ValidateScript({Test-Path $_})]
		[string]$Source,

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

	try {
		$icon = [System.IconExtractor]::Extract($Source, $Index, $Size)

		if ( $icon ) {
			$icon.ToBitmap().Save( $Destination , $imageFormat )
			@{
				source = $Source ,
				destination = $Destination
			}
		}
		else {
			Write-Warning "No associated icon image found in $Source"
			@{
				source = $Source ,
				destination = null
			}
		}
	}
	catch {
		Write-Warning "Extract failed for $Source"
		@{
			source = $Source ,
			destination = null
		}
	}
}

if ( $Stdin ) {
	$input | ConvertFrom-Json | Foreach-Object {
		# This is ONE input, each input being a JSON
		$_ | Foreach-Object {
			#echo "object:" $_
			echo "Source:" $_.source
			echo "Destination:" $_.destination
			echo "Index:" $_.index
			echo "Size:" $_.size
			echo "Format:" $_.format

			Extract-Icon -Source $_.source -Destination $_.destination -Index $_.index -Size $_.size -Format $_.format
		}
	}
}
else {
	echo "Source:" $Source
	echo "Destination:" $Destination
	echo "Index:" $Index
	echo "Size:" $Size
	echo "Format:" $Format
	
	Extract-Icon -Source $Source -Destination $Destination -Index $Index -Size $Size -Format $Format
}


