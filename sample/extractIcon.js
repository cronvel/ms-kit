#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	var result = await msKit.powershell.extractIcon(
		process.argv[ 2 ] ,
		process.argv[ 3 ] ,
		{
			index: process.argv[ 4 ] ,
			size: process.argv[ 5 ]
		}
	) ;
	//console.log( "Link:" , result ) ;
}

run() ;

