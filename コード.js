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
			// Logger.log( "GetRowNumberOfAction(" + actionName + ") : return " + i );
			return i;
		}
		if( actions[ i ][ 2 ] === actionName ) { // 列Cにアクション名の略称が入っている
			// Logger.log( "GetRowNumberOfAction(" + actionName + ") : return " + i );
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
	// Logger.log( "GetSingleUnitReductionRate : called. cells : " + cells );
	var cellsArray = sheet_tl.getRange( cells ).getValues();
	// Logger.log( "GetSingleUnitReductionRate : cellsArray : " + cellsArray );
	var rate       = 1.0;
	cellsArray.forEach( function( actionName ) {
		// Logger.log( '|actionName:' + actionName[ 0 ] + ',' );
		// Logger.log( "|actions : " + actions );
		// Logger.log( "|GetRowNumberOfAction( actionName[ 0 ] ) : " + GetRowNumberOfAction( actionName[ 0 ] ) );
		// Logger.log( "|cellColumnNum[ '自己軽減' ] : " + cellColumnNum[ '自己軽減' ] );

		if( actionName[ 0 ] != '' ) { // 技名が入力されていなかった
			// Logger.log( "actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '単体軽減' ] ] = " + actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '単体軽減' ] ] );
			// Logger.log( "actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '自己軽減' ] ] = " + actions[ GetRowNumberOfAction( actionName[ 0 ] ) ][ cellColumnNum[ '自己軽減' ] ] );
			var r = GetRowNumberOfAction( actionName[ 0 ] );
			if( r === -1 ) return 1; // 見つからなかった
			var c1 = cellColumnNum[ '単体軽減' ];
			var c2 = cellColumnNum[ '自己軽減' ];
			rate *= actions[ r ][ c1 ];
			rate *= actions[ r ][ c2 ];
			// Logger.log( "r :" + r );
			// Logger.log( "c1 : " + c1 );
			// Logger.log( "c2 : " + c2 );
			// Logger.log( "actions[ r ] : " + actions[ r ] );
			// Logger.log( "actions[ r ][ c1 ] : " + actions[ r ][ c1 ] );
			// Logger.log( "actions[ r ][ c2 ] : " + actions[ r ][ c2 ] );
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
	// Logger.log( "GetReductionRate : called" );
	var cellsArray = sheet_tl.getRange( cells ).getValues();
	var rate       = 1.0;
	cellsArray.forEach( function( actionName ) {
		if( actionName[ 0 ] != '' ) { // 技名が入力されていなかった
			var r = GetRowNumberOfAction( actionName[ 0 ] );
			if( r === -1 ) return 1; // 見つからなかった
			var c = cellColumnNum[ index ];
			rate *= actions[ r ][ c ];
			Logger.log( "r :" + r );
			Logger.log( "c : " + c );
			Logger.log( "actions[ r ] : " + actions[ r ] );
			Logger.log( "actions[ r ][ c ] : " + actions[ r ][ c ] );
		}
	} );
	return rate;
}

// 選択したセルを強制的に再計算させます
function RecalcCell( editRow, editColumn ) {
	// var sheet_tl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'TL' );
	// Logger.log( "RecalcCell : called. " + "( editRow, editColumn ) = " + "( " + editRow + ", " + editColumn + " )" );
	// Logger.log( "sheet_tl.getName() " + sheet_tl.getName() );
	var r     = sheet_tl.getRange( editRow, editColumn );
	var value = r.getFormulas();
	// Logger.log( "r : " + r );
	// Logger.log( "value : " + value );
	if( value !== '' ) { return; }
	sheet_tl.getRange( editRow, editColumn ).setValue( "" );    // 一度消す
	sheet_tl.getRange( editRow, editColumn ).setValue( value ); // 再度設定する
	                                                            // Logger.log( "-------------------" );
}

// 指定セルに関連したセルを再計算する
function CalcCell( editRow, editColumn ) {

	var f1 = editRow >= 2;
	var f2 = editColumn >= 8;
	var f3 = editColumn <= 36;
	// Logger.log( "( f1 && f2 && f3 ) :" + ( f1 && f2 && f3 ) );

	if( f1 && f2 && f3 ) {
		// Logger.log( "CalcCell : ( f1 && f2 && f3 ) == true" );
		//  T1の軽減 H,I,J (8,9,10)
		//  T2の軽減 N,O,P (14,15,16)
		//  H1の軽減 U,V,W (21,22,23)
		//  H1の軽減 AA,AB,AC (27,28,29)
		//  DPSの軽減 AH,AI,AJ (34,35,36)
		let index = [ 8, 14, 21, 27, 34 ]; // T1, T2, H1, H2, DPSのバフ欄のそれぞれ一番左のセルの列番号
		index.forEach( function( i ) {
			var flag4 = editColumn >= i + 3;
			var flag5 = editColumn <= i + 5;
			// Logger.log( "i,  flag4 , flag5 ) : ( " + i + ", " + flag4 + ", " + flag5 + " )" );
			// Logger.log( "( flag4 && flag5 ) :" + ( flag4 && flag5 ) );
			if( flag4 && flag5 ) {
				// if( editColumn >= i && editColumn <= i + 2 ) {
				// Logger.log( "CalcCell : Call RecallCell(" + editRow + "," + i + 3 + ")" );
				RecalcCell( editRow, i + 3 ); // i = 8 なら K列
				RecalcCell( editRow, i + 4 ); // i = 8 なら L列
				RecalcCell( editRow, i + 5 ); // i = 8 なら M列
			}
		} );
	}
}

function onEdit( e ) { // 何か操作されたとき呼ばれるコールバック関数
	//操作されたセルの情報 シート名、行、列を取得
	var sheet      = e.range.getSheet().getSheetName();
	var editRow    = e.range.getRow();
	var editColumn = e.range.getColumn();
	if( sheet == 'TL' ) {
		// Logger.log( "onEdit : R" + editRow + "C" + editColumn );
		// Logger.log( "onEdit : Call RecallCell(" + editRow + "C" + editColumn + ")" );
		CalcCell( editRow, editColumn ); // 編集があったセルに関連するセルを再計算する
	}
	SpreadsheetApp.flush();
}

function onSelectionChange( e ) {
	//操作されたセルの情報 シート名、行、列を取得
	var sheet      = e.range.getSheet().getSheetName();
	var editRow    = e.range.getRow();
	var editColumn = e.range.getColumn();
	if( sheet == 'TL' ) {
		// Logger.log( "onSelectionChange : R" + editRow + "C" + editColumn );
		Logger.log( "onSelectionChange : Call RecallCell(" + editRow + "C" + editColumn + ")" );
		CalcCell( editRow, editColumn ); // 編集があったセルに関連するセルを再計算する
		var f1 = editRow === 1;
		var f2 = editColumn === 1;
		Logger.log( "editRow === 1    : " + f1 );
		Logger.log( "editColumn === 1 : " + f2 );
		if( f1 && f2 ) { SpreadsheetApp.flush(); }
	}
}

// 以下 シート内の範囲選択したセルを再計算させるアドオン // {{{
// 引用元URL https://gist.github.com/katz/ab751588580469b35e08
// サイドバーとダイアログを表示する項目を含むカスタムメニューを追加します。
function onOpen( e ) {
	SpreadsheetApp.getUi().createAddonMenu().addItem( 'Re-calculate selected cells', 'recalculate' ).addToUi();
	SpreadsheetApp.flush();
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
// }}}
