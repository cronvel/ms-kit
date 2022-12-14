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
	path_ = path_.replace( /%([^%]+)+%/g , ( match , name ) => ms.getPercentVar( name ) || match ) ;

	// Path like C: are resolved by path.resolve() to current path -_-'
	if ( path_.match( /^[a-z]:$/i ) ) {
		path_ = path_ + '\\' ;
	}
	else {
		path_ = path.resolve( path_ ) ;
	}

	return path_ ;
} ;



ms.findShortcutLinks = async ( dirPath ) => {
	//return fsKit.glob( dirPath + "\\*.lnk" ) ;
	return fsKit.glob( dirPath + "\\**\\*.lnk" ) ;
} ;

