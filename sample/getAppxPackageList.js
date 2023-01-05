#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	var keyword = process.argv[ 2 ] ;
	var result = await msKit.powershell.getAppxPackageList( { iconPath: true } ) ;

	if ( keyword ) {
		keyword = keyword.toLowerCase() ;
		result = result.filter( p =>
			p.appName.toLowerCase().includes( keyword )
			|| p.subAppId.toLowerCase().includes( keyword )
		) ;
	}

	console.log( "Appx packages:" , result ) ;
	console.log( "Count:" , result.length ) ;
}

run() ;

