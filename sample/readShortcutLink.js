#!/usr/bin/env node

"use strict" ; 

const ms = require( '..' ) ;

async function run() {
	var result = await ms.powershell.readShortcutLink(
		process.argv.length > 3 ? process.argv.slice( 2 ) : process.argv[ 2 ] ,
		{ resolvePath: true }
	) ;
	console.log( "Link:" , result ) ;
}

run() ;

