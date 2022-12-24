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
const string = require( 'string-kit' ) ;

const Promise = require( 'seventh' ) ;



const ps = {} ;
module.exports = ps ;

const ms = require( './ms.js' ) ;



const parseStdout = ( type , stdout ) => {
	if ( ! stdout ) {
		return ;
	}

	if ( ! parseStdout[ type ] ) {
		throw new Error( "parseStdout -- type not found: " + type ) ;
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
parseStdout.text = stdout => stdout ;



const escapePowershellArg_ = arg => arg.replace( /["*~;()%?.:@/\\ -]/g , str => '`' + str ) ;
const escapePowershellArg = arg =>
	typeof arg === 'number' ? arg :
	'"' + arg.replace( /["]/g , str => '\\' + str ) + '"' ;



ps.execScript = async ( script , args = null , options = {} ) => {
	var powershellArgs ,
		promise = new Promise() ,
		scriptPath = escapePowershellArg( path.join( __dirname , '..' , 'powershell' , script ) ) ;

	//console.error( "args:" , args ) ;
	args = args === null ? [] :
		Array.isArray( args ) ? args.map( escapePowershellArg ) :
		[ escapePowershellArg( args ) ] ;

	powershellArgs = [
		"-ExecutionPolicy" , "Bypass" ,
		"-File" ,
		scriptPath ,
		... args
	] ;

	var child = childProcess.execFile(
		"powershell.exe" ,
		powershellArgs ,
		{
			shell: "powershell.exe"
		} ,
		( error , stdout , stderr ) => {
			if ( options.debug ) {
				console.log( "stdout:" , stdout ) ;
				console.error( "stderr:" , stderr ) ;
			}

			if ( error ) {
				//console.error( "Command failed:" , error.cmd ) ;
				promise.reject( error ) ;
				return ;
			}

			var result ;

			if ( options.parseStdout ) {
				try {
					result = parseStdout( options.parseStdout , stdout ) ;
				}
				catch ( error_ ) {
					promise.reject( error_ ) ;
				}
			}

			promise.resolve( result ) ;
		}
	) ;

	//child.on( 'spawn' , () => console.log( "Spawned!" ) ) ;
	//child.on( 'exit' , () => console.log( "Exited!" ) ) ;

	if ( options.pipeData ) {
		if ( typeof options.pipeData !== 'string' ) {
			child.stdin.write( JSON.stringify( options.pipeData ) ) ;
		}
		else {
			child.stdin.write( options.pipeData ) ;
		}
	}

	// Always close stdin, because if a powershell could support stdin data, it would hang forever
	child.stdin.end() ;

	return promise ;
} ;



// Can be a single file path or an array of file path
ps.readShortcutLink = async ( shortcutPathList = [] , useMap = false ) => {
	var isList = true ;

	if ( ! Array.isArray( shortcutPathList ) ) {
		shortcutPathList = [ shortcutPathList ] ;
		isList = false ;
	}

	var lnkList = [] ,
		map = {} ,
		index = 0 ;

	for ( let shortcutPath of shortcutPathList ) {
		map[ shortcutPath ] = null ;
		shortcutPath = ms.resolvePath( shortcutPath ) ;

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
		map[ shortcutPath ] = index ++ ;
	}

	if ( ! lnkList.length ) { return [] ; }

	var dataList = await ps.execScript( 'readShortcutLink.ps1' , lnkList , { parseStdout: 'json' } ) ;
	if ( ! Array.isArray( dataList ) ) { dataList = [ dataList ] ; }

	dataList = dataList.map( data => {
		let splittedIconLocation = data.IconLocation.split( ',' ) ;

		return {
			linkPath: data.FullName ,
			arguments: data.Arguments ,
			description: data.Description ,
			hotkey: data.Hotkey ,
			relativePath: data.RelativePath ,
			targetPath: data.TargetPath ,
			windowsStyle: data.WindowStyle ,
			workingDirectory: data.WorkingDirectory ,
			iconPath: ms.resolvePath( splittedIconLocation[ 0 ] || data.TargetPath ) ,
			iconIndex: + splittedIconLocation[ 1 ] || 0
		} ;
	} ) ;

	if ( ! useMap ) { return isList ? dataList : dataList[ 0 ] ; }

	for ( let shortcutPath in map ) {
		index = map[ shortcutPath ] ;
		if ( index !== null ) {
			map[ shortcutPath ] = dataList[ index ] ;
		}
	}

	return map ;
} ;



ps.extractIcon = async ( params = {} ) => {
	var sourcePath = params.source ;

	if ( sourcePath.endsWith( ".lnk" ) ) {
		// This is a .lnk, we can't extract icon from it,
		// however we can extract the link and find out the icon source
		let linkdata = await ps.readShortcutLink( sourcePath ) ;
		sourcePath = linkdata.iconPath ;
		params.index = linkdata.iconIndex ;
	}

	sourcePath = ms.resolvePath( sourcePath ) ;

	var destinationPath = ms.resolvePath( params.destination ) ,
		index = params.index || 0 ,
		size = params.size || 32 ,
		format = params.format || path.extname( destinationPath ).slice( 1 ) ;

	if ( ! await fsKit.isFile( sourcePath ) ) {
		throw new Error( "File not found: " + sourcePath ) ;
	}

	return ps.execScript(
		'extractIcon.ps1' ,
		[ '-Source' , sourcePath , '-Destination' , destinationPath , '-Index' , index , '-Size' , size , '-Format' , format ] ,
		{ parseStdout: 'json' }
	) ;
} ;



ps.extractIconBatch = async ( list , options = {} ) => {
	var lnkDataMap , lnkList = [] ;

	for ( let item of list ) {
		if ( item.source.endsWith( ".lnk" ) ) {
			lnkList.push( item.source ) ;
		}
	}

	if ( lnkList.length ) {
		lnkDataMap = await ps.readShortcutLink( lnkList , true ) ;
	}

	var filteredList = [] ;

	for ( let item of list ) {
		let format = item.format || options.format || path.extname( item.destination ).slice( 1 ) ;
		let size = item.size || options.size || 32 ;
		let destination = ms.resolvePath( item.destination ) ;

		if ( item.source.endsWith( ".lnk" ) ) {
			let lnkData = lnkDataMap?.[ item.source ] ;
			if ( lnkData !== null && await fsKit.isFile( lnkData.iconPath ) ) {
				filteredList.push( {
					destination ,
					size ,
					format ,
					source: ms.resolvePath( lnkData.iconPath ) ,
					index: lnkData.iconIndex || 0
				} ) ;
			}
		}
		else if ( await fsKit.isFile( item.source ) ) {
			filteredList.push( {
				destination ,
				size ,
				format ,
				source: ms.resolvePath( item.source ) ,
				index: item.index || 0
			} ) ;
		}
	}

	//console.log( "filteredList:" , filteredList ) ;

	return ps.execScript( 'extractIcon.ps1' , [ '-Stdin' ] , { pipeData: filteredList , parseStdout: 'json' } ) ;
} ;



ps.getStartMenuShortcutLinkList = async ( options = {} ) => {
	var globalLnkPath = ms.getPercentVar( 'programdata' ) + "\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var globalLnkList = await ms.findShortcutLinks( globalLnkPath ) ;
	//console.log( "globalLnkList:" , globalLnkPath , globalLnkList ) ;

	var userLnkPath = ms.getPercentVar( 'userprofile' ) + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs" ;
	var userLnkList = await ms.findShortcutLinks( userLnkPath ) ;
	//console.log( "userLnkList:" , userLnkPath , userLnkList ) ;


	var globalLnkDataList = await ms.powershell.readShortcutLink( globalLnkList ) ;
	var userLnkDataList = await ms.powershell.readShortcutLink( userLnkList ) ;

	var lnkDataList = [ ... globalLnkDataList , ... userLnkDataList ] ;

	if ( options.extractIconDir ) {
		options.extractIconDir = path.isAbsolute( options.extractIconDir ) ? options.extractIconDir :
			options.extractIconDir = path.resolve( options.extractIconDir ) ;
		//console.log( "extraction to directory:" , options.extractIconDir ) ;
		await fsKit.ensurePath( options.extractIconDir ) ;
	}

	for ( let lnkData of lnkDataList ) {
		let basename = path.basename( lnkData.linkPath ) ;
		lnkData.appName = basename.slice( 0 , -path.extname( basename ).length ) ;
	}

	// .targetPath is mandatory
	lnkDataList = lnkDataList.filter( lnkData => lnkData.targetPath ) ;

	if ( options.extractIconDir ) {
		let batch = [] ;

		for ( let lnkData of lnkDataList ) {
			if ( lnkData.iconPath ) {
				let extractedIconPath = path.join( options.extractIconDir , lnkData.appName ) + ".png" ;
				//console.log( "\nShould be extracting" , lnkData.iconPath , "into" , extractedIconPath ) ;
				batch.push( {
					source: lnkData.iconPath ,
					destination: extractedIconPath ,
					index: lnkData.iconIndex || 0 ,
					size: options.extractIconSize || 256
				} ) ;
				lnkData.extractedIconPath = null ;
			}
		}

		var jobs ;

		try {
			jobs = await ms.powershell.extractIconBatch( batch ) ;
		}
		catch ( error ) {
			console.error( "Error:" , error , "\nCommand:" , error.cmd ) ;
			//throw error ;
		}

		if ( jobs.length !== lnkDataList.length ) {
			throw new Error( "Jobs size mismatch: " + jobs.length + " ; lnks: " + lnkDataList.length ) ;
		}

		var success = 0 , failed = 0 ;

		for ( let i = 0 ; i < jobs.length ; i ++ ) {
			let job = jobs[ i ] ;

			if ( ! job.error ) {
				//success ++ ;
				lnkDataList[ i ].extractedIconPath = job.destination ;
			}
			//else { failed ++ ; }
		}
	}

	//console.log( "Extraction -- success:" , success , "failed:" , failed ) ;

	return lnkDataList ;
} ;



ps.getAppxPackageList = async ( options = {} ) => {
	options.iconSize = + options.iconSize || 256 ;

	var appDataList = await ps.execScript( 'getAppxPackageList.ps1' , null , { parseStdout: 'json' } ) ;

	for ( let appData of appDataList ) {
		if ( ! await fsKit.isFile( appData.iconPath ) ) {
			//console.log( ">>> Icon not found:" , appData.iconPath ) ;

			let iconDir = path.dirname( appData.iconPath ) ,
				iconExtension = path.extname( appData.iconPath ) ,
				iconBasename = path.basename( appData.iconPath , iconExtension ) ,
				globPattern = path.join( iconDir , iconBasename ) + '.scale-*' + iconExtension ;

			//console.log( "=> Trying a glob pattern:" , globPattern ) ;
			let candidateList = await fsKit.glob( globPattern ) ;
			//console.log( "   ... found" , candidateList.length , "candidates:" , candidateList ) ;

			if ( ! candidateList.length ) {
				appData.iconPath = null ;
			}
			else if ( candidateList.length === 1 ) {
				appData.iconPath = candidateList[ 0 ] ;
			}
			else {
				// /!\ be careful /!\ Glob always returns forward slashes, so we have to replace backslashes or it will match nothing!
				let regexpPattern =
						string.escape.regExp( path.join( iconDir , iconBasename ).replace( /\\/g , '/' ) + '.scale-' )
						//+ '([0-9]+)'
						+ '([0-9]+)(_[^.]*)?'
						+ string.escape.regExp( iconExtension ) ,
					regexp = new RegExp( regexpPattern , 'i' ) ,
					maxSizeScore = 0 ,
					chosen = null ,
					altMaxSizeScore = 0 ,
					altChosen = null ;

				//console.log( "=> RegExp Pattern:" , regexpPattern ) ;

				for ( let candidate of candidateList ) {
					let match = candidate.match( regexp ) ;
					//console.warn( "!!! Match:" , match ) ;
					if ( match ) {
						let size = parseInt( match[ 1 ] , 10 ) ,
							sizeScore = size <= options.iconSize ? size / options.iconSize : options.iconSize / size ;
						//console.log( "=> Size score:" , sizeScore ) ;

						if ( match[ 2 ] ) {
							// There is something like 'contrast-white' or 'contrast-black' or whatever in the name,
							// it has a lower priority and will be used only if there is no regular icon
							if ( sizeScore > altMaxSizeScore ) {
								altMaxSizeScore = sizeScore ;
								altChosen = candidate ;
							}
						}
						else {
							if ( sizeScore > maxSizeScore ) {
								maxSizeScore = sizeScore ;
								chosen = candidate ;
							}
						}
					}
				}

				appData.iconPath = chosen || altChosen || candidateList[ 0 ] ;
				//console.log( "Chosen icon:" , chosen ) ;
			}
		}
	}

	return appDataList ;
} ;

