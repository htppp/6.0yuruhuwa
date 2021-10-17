// グローバルなかんじのオブジェクト
var spreadsheet     = SpreadsheetApp.getActiveSpreadsheet();    // 6.0Yurufuwa固定のスプレッドシートオブジェクト つまり大本のオブジェクト
var sheet_calc      = spreadsheet.getSheetByName( '数値計算' ); // シート数値計算のオブジェクト
var actions         = sheet_calc.getDataRange().getValues();    // 何だったか忘れた
var sheet_tl        = spreadsheet.getSheetByName( 'TL' );       // シートTLのオブジェクト

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

// 技名からその技のデータが書いてある行の番号を取得する シート数値計算を参照する
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
		if( actionName != '' ) {
			Logger.log( actionName );
			rate *= actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '単体軽減' ] ];
			rate *= actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '自己軽減' ] ];
		}
	} );
	return rate;
}

// 指定の軽減率を取得する
// cellsには A1B2のような形式でセルの範囲を指定する
// indexには"全体軽減（魔法ダメ）"などが入る
// GetReductionRate( cells, "全体軽減（無条件）" );   // 全体軽減（無条件）を取得する
// GetReductionRate( cells, "全体軽減（魔法ダメ）" ); // 全体軽減（魔法ダメ）を取得する
function GetReductionRate( cells, index ) {
	var cellsArray = sheet_tl.getRange( cells ).getValues();
	var rate       = 1.0;
	cellsArray.forEach( function( actionName ) {
		rate *= actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ index ] ];
	} );
	return rate;
}

// 選択したセルを強制的に再計算させます
function RecalcCell( editRow, editColumn ) {
	var value = sheet_tl.getRange( editRow, editColumn ).getValue(); // セルの値を取得
	if( value.slice( 0, 1 ) !== "=" ) { return; }
	sheet_tl.getRange( editRow, editColumn ).setValue( "" );    // 一度消す
	sheet_tl.getRange( editRow, editColumn ).setValue( value ); // 再度設定する
}

// 指定セルに関連したセルを再計算する
function CalcCell( editRow, editColumn ) {
	if( editRow >= 8 && editColumn >= 2 && editColumn <= 36 ) {
		// T1の軽減 H,I,J (8,9,10)
		// T2の軽減 N,O,P (14,15,16)
		// H1の軽減 U,V,W (21,22,23)
		// H1の軽減 AA,AB,AC (27,28,29)
		// DPSの軽減 AH,AI,AJ (34,35,36)
		let index = [ 8, 14, 21, 27, 34 ]; // T1, T2, H1, H2, DPSのバフ欄のそれぞれ一番左のセルの列番号
		index.forEach( function( i ) {
			if( editRow >= i && editRow >= i + 2 ) {
				RecalcCell( editRow, i + 3 ); // i = 8 なら K列
				RecalcCell( editRow, i + 4 ); // i = 8 なら L列
				RecalcCell( editRow, i + 5 ); // i = 8 なら M列
			}
		} );
	}
}

// function onEdit( e ) { // 何か操作されたとき呼ばれるコールバック関数
// 	//操作されたセルの情報 シート名、行、列を取得
// 	var sheet      = e.range.getSheet().getSheetName();
// 	var editRow    = e.range.getRow();
// 	var editColumn = e.range.getColumn();
// 	if( sheet == 'TL' ) {
// 		Logger.log( "Edited on R" + editRow + "C" + editColumn );
// 		CalcCell( editRow, editColumn ); // 編集があったセルに関連するセルを再計算する
// 	}
// }

function onSelectionChange( e ) {
	//操作されたセルの情報 シート名、行、列を取得
	var sheet      = e.range.getSheet().getSheetName();
	var editRow    = e.range.getRow();
	var editColumn = e.range.getColumn();
	if( sheet == 'TL' ) {
		Logger.log( "onSelectionChange on R" + editRow + "C" + editColumn );
		CalcCell( editRow, editColumn ); // 編集があったセルに関連するセルを再計算する
	}
}

// 以下 シート内の範囲選択したセルを再計算させるアドオン
// 引用元URL https://gist.github.com/katz/ab751588580469b35e08
// サイドバーとダイアログを表示する項目を含むカスタムメニューを追加します。
function onOpen( e ) {
	SpreadsheetApp.getUi().createAddonMenu().addItem( 'Re-calculate selected cells', 'recalculate' ).addToUi();
}

// アドオンがインストールされているときに実行されます。
// onOpen（）を呼び出して、メニューの作成やその他の初期化作業がすぐに行われるようにします。<
function onInstall( e ) {
	onOpen( e );
}

// スプレッドシートに選択したセルを強制的に再計算させます
function recalculate() {
	var activeRange            = SpreadsheetApp.getActiveRange();
	var originalFormulas       = activeRange.getFormulas();
	var originalValues         = activeRange.getValues();

	var valuesToEraseFormula   = [];
	var valuesToRestoreFormula = [];

	originalFormulas.forEach( function( outerVal, outerIdx ) {
		valuesToEraseFormula[ outerIdx ]   = []; // 数式を削除
		valuesToRestoreFormula[ outerIdx ] = []; // 再度数式を設定 これで再計算処理が走るらしい
		outerVal.forEach( function( innerVal, innerIdx ) {
			if( '' === innerVal ) { // 元々数式が設定されていないセルは空文字が返ってくるので喪との文字列を再設定する
				valuesToEraseFormula[ outerIdx ][ innerIdx ]   = originalValues[ outerIdx ][ innerIdx ];
				valuesToRestoreFormula[ outerIdx ][ innerIdx ] = originalValues[ outerIdx ][ innerIdx ];
			} else {
				valuesToEraseFormula[ outerIdx ][ innerIdx ]   = '';                                       // 一度消す
				valuesToRestoreFormula[ outerIdx ][ innerIdx ] = originalFormulas[ outerIdx ][ innerIdx ]; // 再度設定する
			}
		} )
	} )

	activeRange.setValues( valuesToEraseFormula );
	activeRange.setValues( valuesToRestoreFormula );
}
