/*
	MS Kit

	Copyright (c) 2022 Cédric Ronvel

	The MIT License (MIT)

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in all
	copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.
*/

"use strict" ;



const os = require( 'os' ) ;
const path = require( 'path' ) ;
const fs = require( 'fs' ) ;
const fsKit = require( 'fs-kit' ) ;

const Promise = require( 'seventh' ) ;



const ms = {} ;
module.exports = ms ;

ms.powershell = require( './powershell.js' ) ;



// Lower-cased...
const percentVar = [
	'allusersprofile' ,
	'appdata' ,
	'chocolateyinstall' ,
	'commonprogramfiles' , 'commonprogramfiles(x86)' , 'commonprogramw6432' ,
	'driverdata' ,
	'homedrive' , 'homepath' ,
	'localappdata' ,
	'onedrive' ,
	'psmodulepath' ,
	'public' ,
	'programdata' , 'programfiles' , 'programfiles(x86)' , 'programw6432' ,
	'systemdrive' , 'systemroot' ,
	'temp' , 'tmp' ,
	'username' , 'userprofile' , 
	'windir'
] ;

const percentVarSet = new Set( percentVar ) ;

// From lower-case to the actual env name
const percentVarTable = {} ;

for ( let name of Object.keys( process.env ) ) {
	let lowerCaseName = name.toLowerCase() ;
	if ( percentVarSet.has( lowerCaseName ) ) {
		percentVarTable[ lowerCaseName ] = name ;
	}
}

ms.getPercentVar = name => process.env[ percentVarTable[ name.toLowerCase() ] ] ;



// Call path.resolve() and replace %UserProfile% and friends
ms.resolvePath = path_ => {
	path_ = path_.replace( /%([^%]+)+%/ , ( match , name ) => ms.getPercentVar( name ) || match ) ;
	path_ = path.resolve( path_ ) ;
	return path_ ;
} ;



ms.getStartMenuShortcutLinks_ = async ( options = {} ) => {
	var globalLnkPath = process.env.ProgramData + "\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var globalLnkList = await ms.findShortcutLinks( globalLnkPath ) ;
	console.log( "globalLnkList:" , globalLnkPath , globalLnkList ) ;

	var userLnkPath = ms.getPercentVar( 'userprofile' ) + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var userLnkList = await ms.findShortcutLinks( userLnkPath ) ;
	//C:\Users\cedric\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
	console.log( "userLnkList:" , userLnkPath , userLnkList ) ;
	
	var globalLnkDataList = await ms.powershell.readShortcutLink( globalLnkList ) ;
	var userLnkDataList = await ms.powershell.readShortcutLink( userLnkList ) ;
	
	var lnkDataList = [ ... globalLnkDataList , ... userLnkDataList ] ;
	
	if ( options.extractIconDir ) {
		options.extractIconDir = path.isAbsolute( options.extractIconDir ) ? options.extractIconDir :
		options.extractIconDir = path.resolve( options.extractIconDir ) ;
		console.error( "extraction to directory:" , options.extractIconDir ) ;
		await fsKit.ensurePath( options.extractIconDir ) ;
	}
	
	var success = 0 , failed = 0 ;

	for ( let lnkData of lnkDataList ) {
		let basename = path.basename( lnkData.linkPath ) ;
		lnkData.appName = basename.slice( 0 , - path.extname( basename ).length ) ;
		
		if ( options.extractIconDir && lnkData.iconPath ) {
			let extractedIconPath = path.join( options.extractIconDir , lnkData.appName ) + ".png" ;
			console.log( "\nExtracting" , lnkData.iconPath , "into" , extractedIconPath ) ;
			try {
				await ms.powershell.extractIcon( lnkData.iconPath , extractedIconPath , {
					index: lnkData.iconIndex || 0 ,
					size: options.extractIconSize || 256
				} ) ;
				lnkData.extractedIconPath = extractedIconPath ;
				success ++ ;
			}
			catch ( error ) {
				console.error( "Error:" , error , "\nIconLocation:" , lnkData.iconLocation , "\nCommand:" , error.cmd ) ;
				failed ++ ;
				//throw error ;
			}
		}
	}
	
	console.log( "Extraction -- success:" , success , "failed:" , failed ) ;
	return ;
	
	return lnkDataList ;
} ;



ms.getStartMenuShortcutLinks = async ( options = {} ) => {
	var globalLnkPath = process.env.ProgramData + "\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var globalLnkList = await ms.findShortcutLinks( globalLnkPath ) ;
	console.log( "globalLnkList:" , globalLnkPath , globalLnkList ) ;

	var userLnkPath = ms.getPercentVar( 'userprofile' ) + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var userLnkList = await ms.findShortcutLinks( userLnkPath ) ;
	//C:\Users\cedric\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
	console.log( "userLnkList:" , userLnkPath , userLnkList ) ;
	
	
	var globalLnkDataList = await ms.powershell.readShortcutLink( globalLnkList ) ;
	var userLnkDataList = await ms.powershell.readShortcutLink( userLnkList ) ;
	
	var lnkDataList = [ ... globalLnkDataList , ... userLnkDataList ] ;
	
	if ( options.extractIconDir ) {
		options.extractIconDir = path.isAbsolute( options.extractIconDir ) ? options.extractIconDir :
		options.extractIconDir = path.resolve( options.extractIconDir ) ;
		console.error( "extraction to directory:" , options.extractIconDir ) ;
		await fsKit.ensurePath( options.extractIconDir ) ;
	}
	
	//var success = 0 , failed = 0 ;

	for ( let lnkData of lnkDataList ) {
		let basename = path.basename( lnkData.linkPath ) ;
		lnkData.appName = basename.slice( 0 , - path.extname( basename ).length ) ;
	}
	
	if ( options.extractIconDir ) {
		let extractList = [] ;

		for ( let lnkData of lnkDataList ) {
			if ( lnkData.iconPath ) {
				let extractedIconPath = path.join( options.extractIconDir , lnkData.appName ) + ".png" ;
				console.log( "\nShould be extracting" , lnkData.iconPath , "into" , extractedIconPath ) ;
				extractList.push( {
					source: lnkData.iconPath ,
					destination: extractedIconPath ,
					index: lnkData.iconIndex || 0 ,
					size: options.extractIconSize || 256
				} ) ;
				lnkData.extractedIconPath = extractedIconPath ;
				//success ++ ;
			}
		}
		
		try {
			await ms.powershell.extractMultiIcon( extractList ) ;
		}
		catch ( error ) {
			console.error( "Error:" , error , "\nCommand:" , error.cmd ) ;
			//failed ++ ;
			//throw error ;
		}
		
	}
	
	//console.log( "Extraction -- success:" , success , "failed:" , failed ) ;
	return ;
	
	return lnkDataList ;
} ;



ms.findShortcutLinks = async ( dirPath ) => {
	//return fsKit.glob( dirPath + "\\*.lnk" ) ;
	return fsKit.glob( dirPath + "\\**\\*.lnk" ) ;
} ;

