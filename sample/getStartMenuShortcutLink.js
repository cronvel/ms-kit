#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	var result = await msKit.getStartMenuShortcutLinks( {
		extractIconDir: process.argv[ 2 ] || null
	} ) ;
	console.log( "Links:" , result ) ;
}

run() ;

