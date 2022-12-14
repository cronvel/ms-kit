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

    [Parameter(Position = 4 , HelpMessage = "What image format do you want for output?")]
    [ValidateSet("ico","bmp","png","jpg","gif")]
    [string]$Format,

    [Switch]$Associated,

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

		[Parameter(Position = 4 , HelpMessage = "What image format do you want for output?")]
		[ValidateSet("ico","bmp","png","jpg","gif")]
		[string]$Format
	)

	Switch ( $Format ) {
		"ico" { $imageFormat = "icon" }
		"bmp" { $imageFormat = "Bmp" }
		"png" { $imageFormat = "Png" }
		"jpg" { $imageFormat = "Jpeg" }
		"gif" { $imageFormat = "Gif" }
	}

	try {
		$icon = [System.IconExtractor]::Extract($Source, $Index, $Size)

		if ( $icon ) {
			$icon.ToBitmap().Save( $Destination , $imageFormat )
			@{
				source = $Source
				destination = $Destination
				index = $Index
				size = $Size
				format = $Format
			}
		}
		else {
			Write-Warning "No icon could be extracted in $Source"
			@{
				source = $Source
				destination = $Destination
				index = $Index
				size = $Size
				format = $Format
				error = "No icon could be extracted"
			}
		}
	}
	catch {
		@{
			source = $Source
			destination = $Destination
			index = $Index
			size = $Size
			format = $Format
			error = "Extract icon failed"
		}
	}
}



function Extract-Associated-Icon {
	Param(
		[Parameter(Position = 0 , HelpMessage = "Specify the source file.")]
		[ValidateScript({Test-Path $_})]
		[string]$Source,

		[Parameter(Position = 1 , HelpMessage = "Specify the destination file to save the image.")]
		[string]$Destination,

		[Parameter(Position = 2 , HelpMessage = "What image format do you want for output?")]
		[ValidateSet("ico","bmp","png","jpg","gif")]
		[string]$Format
	)

	Switch ( $Format ) {
		"ico" { $imageFormat = "icon" }
		"bmp" { $imageFormat = "Bmp" }
		"png" { $imageFormat = "Png" }
		"jpg" { $imageFormat = "Jpeg" }
		"gif" { $imageFormat = "Gif" }
	}

	try {
		$icon = [System.Drawing.Icon]::ExtractAssociatedIcon( $Source )

		if ( $icon ) {
			$icon.ToBitmap().Save( $Destination , $imageFormat )
			@{
				source = $Source
				destination = $Destination
				format = $Format
			}
		}
		else {
			Write-Warning "No associated icon image found in $Source"
			@{
				source = $Source
				destination = $Destination
				format = $Format
				error = "No associated icon image found"
			}
		}
	}
	catch {
		@{
			source = $Source
			destination = $Destination
			format = $Format
			error = "Extract associated icon failed"
		}
	}
}



if ( $Stdin ) {
	$input | ConvertFrom-Json | Foreach-Object {
		# This is ONE input, each input being a JSON
		$_ | Foreach-Object {
			if ( $_.associated ) {
				Extract-Associated-Icon -Source $_.source -Destination $_.destination -Format $_.format
			}
			else {
				Extract-Icon -Source $_.source -Destination $_.destination -Index $_.index -Size $_.size -Format $_.format
			}
		}
	} | ConvertTo-Json
}
else {
	if ( $Associated ) {
		Extract-Associated-Icon -Source $Source -Destination $Destination -Format $Format | ConvertTo-Json
	}
	else {
		Extract-Icon -Source $Source -Destination $Destination -Index $Index -Size $Size -Format $Format | ConvertTo-Json
	}
}


