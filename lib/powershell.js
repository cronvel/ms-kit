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



const childProcess = require( 'child_process' ) ;
const path = require( 'path' ) ;

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



const execScript = exports.execScript = async ( script , args , options = {} ) => {
	var promise = new Promise() ,
		scriptPath = path.join( __dirname , '..' , 'powershell' , script ) ;

	childProcess.execFile( 
		scriptPath ,
		args ,
		{
			shell: "powershell.exe"
		} ,
		( error , stdout , stderr ) => {
			if ( error ) {
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



exports.readShortcutLink = async ( shortcutPath ) => {
	shortcutPath = path.resolve( shortcutPath ) ;
	
	if ( ! shortcutPath.endsWith( ".lnk" ) ) {
		throw new Error( "The file provided is not a .lnk" ) ;
	}

	if ( ! await fsKit.isFile( shortcutPath ) ) {
		throw new Error( "File not found: " + shortcutPath ) ;
	}
	
	return execScript( 'readShortcutLink.ps1' , [ shortcutPath ] , { parseStdout: 'json' } ) ;
} ;

