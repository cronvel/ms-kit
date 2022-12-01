#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	var result = await msKit.powershell.extractIcon(
		process.argv[ 2 ] ,
		process.argv[ 3 ] ,
		{
			size: process.argv[ 4 ]
		}
	) ;
	//console.log( "Link:" , result ) ;
}

run() ;

