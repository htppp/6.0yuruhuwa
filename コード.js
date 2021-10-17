
function myFunction() {
}

// シート数値計算の20行目の項目の列数
const cellColumnNum = {
	'自己軽減' : 3,
	'単体軽減' : 4,
	'全体軽減（無条件）' : 5,
	'全体軽減（魔法ダメ）' : 6,
	'バリア' : 7,
	'回復量(発動)' : 8,
	'回復量(1tick)' : 9,
	'hot発動数' : 10,
	'hot総回復量(参考)' : 11,
	'回復量(自己のみ、リキャ打ちしないもの)' : 12,
};

// グローバルなかんじのオブジェクト
var      spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 6.0Yurufuwa固定のスプレッドシートオブジェクト
// つまり大本のオブジェクト
var      sheet_calc  = spreadsheet.getSheetByName( '数値計算' ); // シート数値計算のオブジェクト
var      actions     = sheet_calc.getDataRange().getValues();    // 何だったか忘れた
var      sheet_tl    = spreadsheet.getSheetByName( 'TL' );       // シートTLのオブジェクト

// 技名からその技のデータが書いてある行の番号を取得する シート数値計算
// を参照する
function GetRowNumberOfAction( actionName ) {
	for( var i = 20; i < actions.length; i++ ) { // 行21からアクション名が入っている
		if( actions[ i ][ 1 ] === actionName ) { // 列Bにアクション名が入っている
			return i;
		}
	}
	return -1;
}

function test( cells ) {
	var cellsArray = sheet_tl.getRange( cells ).getValues();
	var str        = '';
	var actions    = sheet_calc.getDataRange().getValues();
	cellsArray.forEach( function( actionName ) {
		str += actionName[ 0 ] + ':' + GetRowNumberOfAction( actionName[ 0 ] ) + ':' + cellColumnNum[ '全体軽減（無条件）' ] + +actions.length + '\n';
	} );
	return str;
}

// 単体の軽減率を取得する
// cellsには A1B2のような形式でセルの範囲を指定する
function GetSingleUnitReductionRate( cells ) {
	var cellsArray = sheet_tl.getRange( cells ).getValues();
	var rate       = 1.0;
	cellsArray.forEach( function( actionName ) {
		console.log( 'actionNam:' + actionName[ 0 ] + ',' );
		rate *= actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '単体軽減' ] ];
		rate *= actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '自己軽減' ] ];
	} );
	return rate;
}

// 全体軽減（無条件）を取得する
// cellsには A1B2のような形式でセルの範囲を指定する
function GetOverallReductionRate( cells ) {
	var cellsArray = sheet_tl.getRange( cells ).getValues();
	var rate       = 1.0;
	cellsArray.forEach( function( actionName ) {
		rate *= actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '全体軽減（無条件）' ] ];
	} );
	return rate;
}

// 全体軽減（魔法ダメ）を取得する
// cellsには A1B2のような形式でセルの範囲を指定する
function GetOverallMagicReductionRate( cells ) {
	var cellsArray = sheet_tl.getRange( cells ).getValues();
	var rate       = 1.0;
	cellsArray.forEach( function( actionName ) {
		rate *= actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '全体軽減（魔法ダメ）' ] ];
	} );
	return rate;
}
