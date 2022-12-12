#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	var result = await msKit.getStartMenuShortcutLinks() ;
	console.log( "Links:" , result ) ;
}

run() ;

