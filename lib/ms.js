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



const msKit = {} ;
module.exports = msKit ;

msKit.powershell = require( './powershell.js' ) ;



const percentVar = [
	'homedrive' , 'homepath' , 'userprofile' , 
	'systemdrive' , 'systemroot' ,
	'programfiles' , 'programdata'
] ;

const percentVarSet = new Set( percentVar ) ;
const percentVarTable = {} ;

for ( let name of process.env ) {
	let lowerCaseName = name.toLowerCase() ;
	if ( percentVarSet.has( lowerCaseName ) ) {
		percentVarTable[ lowerCaseName ] = name ;
	}
}



// Call path.resolve() and replace %UserProfile% and friends
msKit.resolvePath = path_ => {
	path_ = path.resolve( path_ ) ;
	
	path_.replace( /%([^%])+%/ , ( match , name ) => {
		name = name.toLowerCase() ;

		if ( percentVarTable[ name ] ) {
			return process.env[ percentVarTable[ name ] ] ;
		}

		return match ;
	} ) ;
	return path_ ;
} ;



msKit.getStartMenuShortcutLinks = async ( options = {} ) => {
	var globalLnkPath = process.env.ProgramData + "\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var globalLnkList = await msKit.findShortcutLinks( globalLnkPath ) ;
	console.log( "globalLnkList:" , globalLnkPath , globalLnkList ) ;

	var userLnkPath = os.homedir() + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var userLnkList = await msKit.findShortcutLinks( userLnkPath ) ;
	//C:\Users\cedric\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
	console.log( "userLnkList:" , userLnkPath , userLnkList ) ;
	
	
	var globalLnkDataList = await msKit.powershell.readShortcutLink( globalLnkList ) ;
	var userLnkDataList = await msKit.powershell.readShortcutLink( userLnkList ) ;
	
	var lnkDataList = [ ... globalLnkDataList , ... userLnkDataList ] ;
	
	if ( options.extractIconDir ) {
		options.extractIconDir = path.isAbsolute( options.extractIconDir ) ? options.extractIconDir :
		options.extractIconDir = path.resolve( options.extractIconDir ) ;
		console.error( "extraction to:" , options.extractIconDir ) ;
		await fsKit.ensurePath( options.extractIconDir ) ;
	}
	
	for ( let lnkData of lnkDataList ) {
		let basename = path.basename( lnkData.FullName ) ;
		lnkData.AppName = basename.slice( 0 , - path.extname( basename ).length ) ;
		
		
		if ( options.extractIconDir && lnkData.IconPath ) {
			let extractedIconPath = path.join( options.extractIconDir , lnkData.AppName ) ;
			console.log( "Extracting" , lnkData.IconPath , "into" , lnkData.ExtractedIconPath ) ;
			try {
				await msKit.powershell.extractIcon( lnkData.IconPath , lnkData.ExtractedIconPath , {
					index: lnkData.IconIndex || 0 ,
					size: options.extractIconSize || 256
				} ) ;
				lnkData.ExtractedIconPath = extractedIconPath ;
			}
			catch ( error ) {
				console.error( "Error:" , error ) ;
				//throw error ;
			}
		}
	}
	
	return lnkDataList ;
} ;



msKit.findShortcutLinks = async ( dirPath ) => {
	//return fsKit.glob( dirPath + "\\*.lnk" ) ;
	return fsKit.glob( dirPath + "\\**\\*.lnk" ) ;
} ;

