#!/usr/bin/env node

"use strict" ; 



const msKit = require( '..' ) ;



var list = [
	{
		source: "C:\\Users\\cedric\\code\\powershell\\firefox.exe.lnk" ,
		destination: "C:\\Users\\cedric\\code\\ms-kit\\sample\\extracted\\firefox.png" ,
		size: 256
	},
	{
		source: "C:\\Users\\cedric\\code\\powershell\\firefox2.exe.lnk" ,
		destination: "C:\\Users\\cedric\\code\\ms-kit\\sample\\extracted\\firefox2.png" ,
		size: 256
	},
	{
		source: "C:\\does\\not\\exist.lnk" ,
		destination: "C:\\does\\not\\exist.png" ,
		size: 256
	}
] ;



async function run() {
	var result = await msKit.powershell.extractIconBatch( list ) ;
	console.log( "Result:" , result ) ;
}

run() ;

