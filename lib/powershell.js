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



/*
	Notice: may need unrestricted execution policy to execute powershell scripts:
		Get-ExecutionPolicy
		Set-ExecutionPolicy Unrestricted
		Set-ExecutionPolicy Default
*/



const childProcess = require( 'child_process' ) ;
const path = require( 'path' ) ;
const fs = require( 'fs' ) ;
const fsKit = require( 'fs-kit' ) ;

const Promise = require( 'seventh' ) ;



const parseStdout = ( type , stdout ) => {
	if ( ! stdout ) {
		return ;
	}
	
	if ( ! parseStdout[ type ] ) {
		throw new Error( "parseStdout -- type not found: " + type ) ;
		return ;
	}
	
	try {
		var result = parseStdout[ type ]( stdout ) ;
	}
	catch ( error ) {
		console.error( "parseStdout -- Error parsing: '" + stdout + "'" ) ;
		throw error ;
	}
	
	return result ;
} ;

parseStdout.json = stdout => JSON.parse( stdout ) ;



const escapePowershellArg_ = arg => arg.replace(
	/["*~;()%?.:@\/\\ -]/g ,
	str => '`' + str
) ;

const escapePowershellArg = arg =>
	'"' + arg.replace( /["]/g , str => '\\' + str ) + '"' ;



const execScript = exports.execScript = async ( script , args = null , options = {} ) => {
	var powershellArgs ,
		promise = new Promise() ,
		scriptPath = path.join( __dirname , '..' , 'powershell' , script ) ;

	args = args === null ? [] :
		Array.isArray( args ) ? args :
		[ args ] ;
	
	powershellArgs = [
		"-ExecutionPolicy" , "Bypass" ,
		"-File" ,
		escapePowershellArg( scriptPath ) ,
		... args.map( escapePowershellArg )
	] ;
	
	childProcess.execFile( 
		"powershell.exe" ,
		powershellArgs ,
		{
			shell: "powershell.exe"
		} ,
		( error , stdout , stderr ) => {
			if ( error ) {
				//console.error( "Command failed:" , error.cmd ) ;
				promise.reject( error ) ;
				return ;
			}

			//console.log( "stdout:" , stdout ) ;
			//console.error( "stderr:" , stderr ) ;
			
			var result ;

			if ( options.parseStdout ) {
				try {
					result = parseStdout( options.parseStdout , stdout ) ;
				}
				catch ( error ) {
					promise.reject( error ) ;
				}
			}
			
			promise.resolve( result ) ;
		}
	) ;
	
	return promise ;
} ;



// Can be a single file path or an array of file path
exports.readShortcutLink = async ( shortcutPathList = [] ) => {
	var isList = true ;

	if ( ! Array.isArray( shortcutPathList ) ) {
		shortcutPathList = [ shortcutPathList ] ;
		isList = false ;
	}

	var lnkList = [] ;

	for ( let shortcutPath of shortcutPathList ) {
		shortcutPath = path.resolve( shortcutPath ) ;
		
		if ( ! shortcutPath.endsWith( ".lnk" ) ) {
			if ( ! isList ) {
				throw new Error( "The file provided is not a .lnk" ) ;
			}
			
			continue ;
		}

		if ( ! await fsKit.isFile( shortcutPath ) ) {
			if ( ! isList ) {
				throw new Error( "File not found: " + shortcutPath ) ;
			}
			
			continue ;
		}
		
		lnkList.push( shortcutPath ) ;
	}
	
	if ( ! lnkList.length ) { return [] ; }
	
	var dataList = await execScript( 'readShortcutLink.ps1' , lnkList , { parseStdout: 'json' } ) ;
	if ( ! Array.isArray( dataList ) ) { dataList = [ dataList ] ; }
	
	for ( let data of dataList ) {
		let splitted = data.IconLocation.split( ',' ) ;
		data.IconPath = splitted[ 0 ] || data.TargetPath ;
		data.IconIndex = splitted[ 1 ] || 0 ;
	}

	return isList ? dataList : dataList[ 0 ] ;
} ;



exports.extractIcon = async ( sourcePath , destinationPath , options = {} ) => {
	if ( sourcePath.endsWith( ".lnk" ) ) {
		// This is a .lnk, we can't extract icon from it,
		// however we can extract the link and find out the icon source
		let linkdata = await exports.readShortcutLink( sourcePath ) ;
		sourcePath = linkdata.IconPath ;
		options.index = linkdata.IconIndex ;
	}

	sourcePath = path.resolve( sourcePath ) ;
	destinationPath = path.resolve( destinationPath ) ;
	
	var format = options.format || path.extname( destinationPath ).slice( 1 ) ,
		size = options.size || 32 ,
		imageIndex = options.index || 0 ;
	
	if ( ! await fsKit.isFile( sourcePath ) ) {
		throw new Error( "File not found: " + sourcePath ) ;
	}
	
	return execScript( 'extractIcon.ps1' , [ sourcePath , destinationPath , imageIndex , size , format ] ) ;
} ;



exports.getAppxPackageList = async ( options = {} ) => {
	return execScript( 'getAppxPackageList.ps1' , null , { parseStdout: 'json' } ) ;
} ;

