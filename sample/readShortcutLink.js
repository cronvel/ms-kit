#!/usr/bin/env node

"use strict" ; 

const msKit = require( '..' ) ;

async function run() {
	msKit.powershell.readShortcutLink( '../../powershell/firefox.exe.lnk' ) ;
}

run() ;

