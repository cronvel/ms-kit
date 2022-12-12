#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	var result = await msKit.powershell.readShortcutLink(
		process.argv.length > 3 ? process.argv.slice( 2 ) : process.argv[ 2 ]
	) ;
	console.log( "Link:" , result ) ;
}

run() ;

