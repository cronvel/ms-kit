#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	var result = await msKit.powershell.getAppxPackageList( { iconPath: true } ) ;
	console.log( "Appx packages:" , result ) ;
}

run() ;

