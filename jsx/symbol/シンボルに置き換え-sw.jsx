#target illustrator
#targetengine "SymbolReplaceSwEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメントに登録されているシンボルを一覧から選び、選択したオブジェクトをそのシンボルインスタンスへ置き換えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/シンボルに置き換え-sw.md

### Overview

Picks a symbol from the ones registered in the document and replaces the selected objects with instances of it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/シンボルに置き換え-sw.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "シンボルに置き換え-sw";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v0.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Toshiyuki Takahashi";          /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2015-12-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/シンボルに置き換え-sw.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/シンボルに置き換え-sw.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author Toshiyuki Takahashi
 * @discussion https://github.com/gau/object-to-symbol
 * @discussion https://graphicartsunit.tumblr.com/post/134802610854/object-to-symbol
 */

(function () {

	// =========================================
	// ユーザー設定 / User Settings
	// =========================================
	var MAX_VISIBLE_SYMBOLS = 20;      /* 一覧に一度に並べるシンボルの数（超えるとスクロールバー）/ symbols shown at once; more adds a scrollbar */
	var SLOW_SELECTION_THRESHOLD = 20; /* これを超える選択数では続けるか確認する / ask before running on more objects than this */

	// =========================================
	// レイアウト / Layout
	// =========================================
	var SYMBOL_RADIO_ROW_HEIGHT = 20;  /* シンボルのラジオボタン1行の高さ / height of one symbol radio row */
	var SYMBOL_RADIO_SPACING = 2;      /* ラジオボタンどうしの間隔 / spacing between the radios */
	var SYMBOL_SCROLLBAR_WIDTH = 16;   /* スクロールバーの幅 / scrollbar width */
	var SYMBOL_LIST_SPACING = 4;       /* ラジオの列とスクロールバーの間隔 / gap between the radios and the scrollbar */

	// =========================================
	// 実行中の設定 / Current settings
	// =========================================
	var replaceSettings = {
		symbolIndex: 0,  /* 使うシンボルの documentSymbols 上のインデックス / index of the symbol in documentSymbols */
		anchorIndex: 4   /* 揃える基準点 0〜8（4=中央）/ anchor to align on, 0..8 (4 = center) */
	};

	// ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

	var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
	var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
	var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

	/**
	 * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
	 * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
	 * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
	 * @param {Window} dialog - 対象のダイアログ
	 * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
	 * @returns {void}
	 */
	function prepareDialogWindow(dialog, storageKey) {
	    /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
	       When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
	    if (dialog.dialogWindowState) {
	        dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
	        dialog.dialogWindowState.avoidedLocation = null;
	        return;
	    }
	    var locationKey = "__" + storageKey + "_DialogLocation";
	    var previousOnShow = dialog.onShow;
	    var previousOnMove = dialog.onMove;
	    var previousOnClose = dialog.onClose;
	    var windowState = {
	        selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
	        screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
	        avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
	    };
	    dialog.dialogWindowState = windowState;

	    dialog.opacity = DIALOG_OPACITY;

	    /* 今の位置を記録する / Remember the current location */
	    function rememberDialogLocation() {
	        var currentLocation = [dialog.location[0], dialog.location[1]];
	        var avoidedLocation = windowState.avoidedLocation;
	        if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
	        $.global[locationKey] = currentLocation;
	    }

	    dialog.onShow = function () {
	        /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
	           On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
	        if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
	        if (previousOnShow) previousOnShow.apply(this, arguments);
	        /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
	           $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
	        var savedLocation = $.global[locationKey];
	        if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
	        if (windowState.selectionSpan) {
	            var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
	            if (avoidLeft !== null) {
	                dialog.location = [avoidLeft, dialog.location[1]];
	                /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
	                windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
	            }
	        }
	    };
	    dialog.onMove = function () {
	        if (previousOnMove) previousOnMove.apply(this, arguments);
	        rememberDialogLocation();
	    };
	    dialog.onClose = function () {
	        rememberDialogLocation();
	        /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
	           Returning false cancels the close, so pass the original onClose result through */
	        if (previousOnClose) return previousOnClose.apply(this, arguments);
	    };
	}

	/**
	 * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
	 * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
	 */
	function getSelectionViewSpan() {
	    try {
	        if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
	        var targetDoc = app.activeDocument;
	        var selectedItems = targetDoc.selection;
	        /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
	        if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
	        var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
	        var spanLeft = Infinity;
	        var spanRight = -Infinity;
	        for (var i = 0; i < itemCount; i++) {
	            var itemBounds = selectedItems[i].visibleBounds;
	            if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
	            if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
	        }
	        var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
	        var viewBounds = activeView.bounds;
	        var zoom = activeView.zoom;
	        var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
	        /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
	        var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
	        var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
	        if (right <= left) return null;
	        return { left: left, right: right, viewWidth: viewWidth };
	    } catch (e) {
	        /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
	        return null;
	    }
	}

	/**
	 * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
	 * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
	 * @param {number} dialogLeft - 今のダイアログの左端
	 * @param {number} dialogWidth - ダイアログの幅
	 * @param {number} screenWidth - 画面の幅
	 * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
	 * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
	 */
	function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
	    var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
	    var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
	    var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
	    if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

	    var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
	    var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
	    var fitsLeft = leftSideLeft >= 0;
	    var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
	    /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
	    var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
	    if (preferLeft && fitsLeft) return leftSideLeft;
	    if (fitsRight) return rightSideLeft;
	    if (fitsLeft) return leftSideLeft;
	    return null;
	}

	// ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

	// =========================================
	// ローカライズ / Localization
	// =========================================

	// ローカライズ（再利用パーツ） / Localization (reusable)

	/**
	 * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
	 * @returns {string} "ja" または "en"
	 */
	function getCurrentLang() {
	    return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
	}

	var uiLang = getCurrentLang();

	/**
	 * LABELS から今の UI 言語の文言を取り出す。
	 * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
	 * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
	 * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
	 */
	function getLabel(labelRef, placeholderValues) {
	    var labelEntry = labelRef;
	    if (typeof labelRef === "string") {
	        var labelPathKeys = labelRef.split(".");
	        labelEntry = LABELS;
	        for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
	            labelEntry = labelEntry[labelPathKeys[i]];
	        }
	    }
	    var labelString;
	    if (typeof labelEntry === "string") labelString = labelEntry;
	    else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
	    else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
	    else return (typeof labelRef === "string") ? labelRef : "";
	    return fillLabelPlaceholders(String(labelString), placeholderValues);
	}

	/**
	 * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
	 * @param {string|Object} labelRef - getLabel と同じ
	 * @param {Object|Array} [placeholderValues] - getLabel と同じ
	 * @returns {string} コロン付きの文言
	 */
	function labelText(labelRef, placeholderValues) {
	    return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
	}

	/**
	 * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
	 * @param {string|Object} labelRef - getLabel と同じ
	 * @param {string|number} value - コロンのあとに続ける値
	 * @returns {string} 項目名と値をつないだ文字列
	 */
	function labelValueText(labelRef, value) {
	    return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
	}

	/**
	 * 文言の {name} や %1 に値を差し込む
	 * @param {string} labelString - 文言
	 * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
	 * @returns {string} 差し込んだ文言
	 */
	function fillLabelPlaceholders(labelString, placeholderValues) {
	    if (placeholderValues == null) return labelString;
	    if (placeholderValues instanceof Array) {
	        /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
	        for (var i = placeholderValues.length; i >= 1; i--) {
	            labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
	        }
	        return labelString;
	    }
	    for (var placeholderKey in placeholderValues) {
	        if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
	        labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
	    }
	    return labelString;
	}

	// ローカライズ（再利用パーツ）ここまで / End of the reusable localization

	/* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
	var LABELS = {
		dialog: {
			title: { ja: 'シンボルに置き換え', en: 'Replace with Symbol' }
		},
		panel: {
			symbol: { ja: 'シンボル', en: 'Symbol' },
			anchor: { ja: '基準点', en: 'Reference Point' }
		},
		button: {
			ok:     { ja: '実行', en: 'Run' },
			cancel: { ja: 'キャンセル', en: 'Cancel' }
		},
		tooltip: {
			symbol: {
				ja: '選択したオブジェクトを、このシンボルのインスタンスに置き換えます。',
				en: 'Replaces the selected objects with an instance of this symbol.'
			},
			anchor: {
				ja: '置き換え前のオブジェクトとシンボルを、この基準点で揃えます。',
				en: 'Aligns the symbol with the original object at this reference point.'
			}
		},
		alert: {
			noDocument:  { ja: 'ドキュメントが開かれていません。', en: 'No document is open.' },
			noSelection: { ja: 'オブジェクトが選択されていません。', en: 'No objects are selected.' },
			noSymbols:   { ja: 'ドキュメントにシンボルがありません。', en: 'The document has no symbols.' },
			layerUnavailable: {
				ja: '現在のレイヤーがロックされているか、非表示になっています。',
				en: 'The current layer is locked or hidden.'
			},
			slowSelection: {
				ja: '%1個のオブジェクトが選択されています。処理に時間がかかることがあります。\n続けますか？',
				en: '%1 objects are selected. This may take a long time.\nContinue?'
			},
			error: {
				ja: 'エラーが発生したため、処理を実行できませんでした。\nエラー内容：',
				en: 'The operation could not be completed because of an error.\nError: '
			}
		}
	};

	// ボタン行（再利用パーツ） / Button row (reusable)

	var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
	var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
	var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
	var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

	/**
	 * ダイアログ下部のボタン行を作る。
	 * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
	 * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
	 * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
	 * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
	 */
	function addButtonRow(parent, rowOptions) {
	    var isCentered = !!(rowOptions && rowOptions.centered);
	    var btnRowGroup = parent.add("group");
	    btnRowGroup.orientation = "row";
	    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
	    btnRowGroup.spacing = BUTTON_ROW_SPACING;

	    if (isCentered) {
	        btnRowGroup.alignment = ["center", "bottom"];
	        btnRowGroup.alignChildren = ["center", "center"];
	        return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
	    }

	    btnRowGroup.alignment = ["fill", "bottom"];

	    var btnLeftGroup = btnRowGroup.add("group");
	    btnLeftGroup.alignChildren = ["left", "center"];
	    btnLeftGroup.spacing = BUTTON_ROW_SPACING;

	    /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
	    var spacer = btnRowGroup.add("group");
	    spacer.alignment = ["fill", "fill"];
	    spacer.minimumSize.width = 0;

	    var btnRightGroup = btnRowGroup.add("group");
	    btnRightGroup.alignChildren = ["right", "center"];
	    btnRightGroup.spacing = BUTTON_ROW_SPACING;

	    return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
	}

	/**
	 * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
	 * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
	 * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
	 * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
	 * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
	 * @returns {void}
	 */
	function alignRightOnlyButtonRow(buttonRow) {
	    if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
	    var dialogWindow = buttonRow.rowGroup.window;
	    dialogWindow.addEventListener("show", function () {
	        if (!buttonRow.leftGroup) return;
	        var btnRowGroup = buttonRow.rowGroup;
	        /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
	        if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
	        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
	        btnRowGroup.remove(buttonRow.leftGroup);
	        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
	        btnRowGroup.alignment = ["center", "bottom"];
	        btnRowGroup.alignChildren = ["center", "center"];
	        buttonRow.leftGroup = null;
	        dialogWindow.layout.layout(true);
	    });
	}

	// ボタン行（再利用パーツ）ここまで / End of the reusable button row

	// UIレイアウト（再利用パーツ） / UI layout (reusable)

	/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
	var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
	var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
	var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
	var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
	var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
	var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

	/**
	 * ウィンドウの共通設定
	 * @param {Window} targetWindow - 対象のウィンドウ
	 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
	 * @returns {void}
	 */
	function setupWindow(targetWindow, spacing) {
	    targetWindow.orientation = "column";
	    targetWindow.alignChildren = "fill";
	    targetWindow.margins = WINDOW_MARGINS;
	    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
	}

	/**
	 * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
	 * @param {Panel} targetPanel - 対象のパネル
	 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
	 * @returns {void}
	 */
	function setupPanel(targetPanel, spacing) {
	    targetPanel.orientation = "column";
	    targetPanel.alignChildren = ["fill", "top"];
	    targetPanel.alignment = "fill";
	    targetPanel.margins = PANEL_MARGINS;
	    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
	}

	/**
	 * タブの共通設定
	 * @param {Tab} targetTab - 対象のタブ
	 * @param {number} [spacing] - 要素間隔（省略時は変えない）
	 * @returns {void}
	 */
	function setupTab(targetTab, spacing) {
	    targetTab.orientation = "column";
	    targetTab.alignChildren = "fill";
	    targetTab.margins = TAB_MARGINS;
	    if (typeof spacing === "number") targetTab.spacing = spacing;
	}

	/**
	 * 横並びの行グループの共通設定（ボタン列など）。
	 * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
	 * @param {Group} rowGroup - 対象のグループ
	 * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
	 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
	 * @returns {void}
	 */
	function setupRow(rowGroup, rowAlignment, spacing) {
	    rowGroup.orientation = "row";
	    rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
	    rowGroup.alignChildren = ["left", "center"];
	    rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
	}

	/**
	 * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
	 * @param {Button} targetButton - 対象のボタン
	 * @param {number} trimPixels - 詰める量（px）
	 * @returns {void}
	 */
	function trimButtonHeight(targetButton, trimPixels) {
	    /* レイアウト前は size が無い / size is not set until the layout runs */
	    if (!targetButton.size) return;
	    targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
	}

	// UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

	// UI の明暗（再利用パーツ） / UI theme (reusable)

	/**
	 * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
	 * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
	 */
	function isDarkUI() {
	    try {
	        if (app.preferences && app.preferences.getRealPreference) {
	            return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
	        }
	        return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
	    } catch (e) {
	        return false;
	    }
	}

	// UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

	// 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)

	// -----------------------------------------
	// 基準点ウィジェットの寸法 / Anchor widget metrics
	// -----------------------------------------
	var ANCHOR_WIDGET_SIZE      = 66;   /* ウィジェット全体の一辺 / overall size of the widget */
	var ANCHOR_WIDGET_CELL_SIZE = 9;    /* □1個の一辺 / size of one square */
	var ANCHOR_WIDGET_CELL_GAP  = 7.5;  /* □どうしの間隔 / gap between squares */
	var ANCHOR_WIDGET_NONE      = -1;   /* 未選択のインデックス / index while nothing is selected */

	/* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
	   Cell names in row-major order, matching the Transformation enumeration */
	var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

	/* 中央(4)を除く外周の□どうしをつなぐケイ線 / Rules joining the outer squares (the center stands alone) */
	var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

	// -----------------------------------------
	// 基準点ウィジェットの配色 / Anchor widget colors
	// -----------------------------------------
	var ANCHOR_WIDGET_UI_DARK = isDarkUI();
	/* 枠線・ケイ線はグレー、選択セルの塗りはライトで濃いグレー・ダークで明るいグレー（既存スクリプトの配色を踏襲）。
	   無効時は同じ色を半透明にして背景へ沈める（不透明の薄いグレーだとダークUIで逆に明るく浮くため）
	   Gray rules; the selected fill is dark gray on light UI and light gray on dark UI (as in the existing scripts).
	   Disabled colors are translucent versions so they sink into any background */
	var ANCHOR_WIDGET_LINE_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 1]     : [0.42, 0.42, 0.42, 1];  /* 枠線・ケイ線 / rules */
	var ANCHOR_WIDGET_FILL_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 1]     : [0.27, 0.27, 0.27, 1];  /* 選択セルの塗り / selected fill */
	var ANCHOR_WIDGET_DIM_LINE_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 0.4]   : [0.42, 0.42, 0.42, 0.4];  /* 無効時の枠線 / rules when disabled */
	var ANCHOR_WIDGET_DIM_FILL_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 0.3]   : [0.27, 0.27, 0.27, 0.3];  /* 無効時の塗り / fill when disabled */

	// -----------------------------------------
	// ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
	// -----------------------------------------
	/**
	 * 基準点（3×3）を選ぶウィジェットを追加する。クリックしたセルを選び、onChange を呼ぶ
	 * @param {Group|Panel} parent - 追加先
	 * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
	 * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
	 * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）/ size（一辺。既定 66）
	 * @returns {Button} ウィジェット（値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
	 */
	function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
	    var anchorOptions = widgetOptions || {};
	    var widgetSize = anchorOptions.size || ANCHOR_WIDGET_SIZE;
	    var anchorWidget = parent.add("button", undefined, "");
	    anchorWidget.minimumSize = [widgetSize, widgetSize];
	    anchorWidget.preferredSize = [widgetSize, widgetSize];
	    anchorWidget.maximumSize = [widgetSize, widgetSize];
	    anchorWidget.isAnchorWidget = true; /* redrawAnchorWidgetsIn() の目印 / marker for redrawAnchorWidgetsIn() */
	    anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
	    anchorWidget.anchorDisabledCells = toAnchorCellFlags(anchorOptions.disabledCells);
	    anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(initialValue, anchorWidget.anchorAllowNone);
	    anchorWidget.onDraw = function () { drawAnchorWidget(anchorWidget); };
	    anchorWidget.onClick = function () {}; /* セルの判定は mousedown で行う / hit-testing happens in mousedown */

	    /* クリック座標（コントロール基準）を3分割してセルを判定する / split the control-relative click into thirds */
	    anchorWidget.addEventListener("mousedown", function (event) {
	        if (!isAnchorWidgetEnabledInTree(anchorWidget)) return;
	        var cellIndex = getAnchorCellAt(event.clientX, event.clientY, anchorWidget.size[0], anchorWidget.size[1]);
	        if (anchorWidget.anchorDisabledCells[cellIndex]) return;
	        anchorWidget.anchorWidgetIndex = cellIndex;
	        redrawAnchorWidget(anchorWidget);
	        if (onChange) onChange(cellIndex, anchorWidget);
	    });
	    return anchorWidget;
	}

	/**
	 * 選択中のセルのインデックスを返す
	 * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
	 * @returns {number} 0〜8（行優先）。未選択なら -1
	 */
	function getAnchorWidgetIndex(anchorWidget) {
	    return anchorWidget.anchorWidgetIndex;
	}

	/**
	 * 選択中のセルの名前を返す
	 * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
	 * @returns {string} "topLeft" など。未選択なら ""
	 */
	function getAnchorWidgetName(anchorWidget) {
	    return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
	}

	/**
	 * 選択するセルを変えて描き直す（onChange は呼ばない）
	 * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
	 * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
	 * @returns {void}
	 */
	function setAnchorWidgetValue(anchorWidget, anchorValue) {
	    anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
	    redrawAnchorWidget(anchorWidget);
	}

	/**
	 * ウィジェットの有効／無効を切り替えて描き直す（無効の間は薄く描き、クリックも無視する）
	 * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
	 * @param {boolean} isEnabled - 有効にするなら true
	 * @returns {void}
	 */
	function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
	    anchorWidget.enabled = isEnabled;
	    redrawAnchorWidget(anchorWidget);
	}

	/**
	 * 選べないセルを指定し直して描き直す（選択中のセルは変えない）
	 * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
	 * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
	 * @returns {void}
	 */
	function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
	    anchorWidget.anchorDisabledCells = toAnchorCellFlags(disabledCells);
	    redrawAnchorWidget(anchorWidget);
	}

	/**
	 * コンテナ以下にある基準点ウィジェットをすべて描き直す。パネルや行の enabled を切り替えたあとに呼ぶ
	 * @param {Object} container - パネル・グループ・ウィンドウなど
	 * @returns {void}
	 */
	function redrawAnchorWidgetsIn(container) {
	    if (container.isAnchorWidget) {
	        redrawAnchorWidget(container);
	        return;
	    }
	    if (!container.children) return;
	    for (var i = 0; i < container.children.length; i++) {
	        redrawAnchorWidgetsIn(container.children[i]);
	    }
	}

	// -----------------------------------------
	// 値の変換 / Value helpers
	// -----------------------------------------
	/**
	 * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
	 * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
	 * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
	 * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
	 */
	function resolveAnchorWidgetIndex(anchorValue, allowNone) {
	    if (typeof anchorValue === "string") {
	        for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
	            if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
	        }
	        return 4;
	    }
	    if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
	    if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
	        return anchorValue;
	    }
	    return 4;
	}

	/**
	 * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
	 * @param {number|string} anchorValue - 0〜8 か名前
	 * @returns {number[]} [横の割合, 縦の割合]
	 */
	function getAnchorRatio(anchorValue) {
	    var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
	    return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
	}

	/**
	 * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
	 * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
	 * @param {number|string} anchorValue - 0〜8 か名前
	 * @returns {number[]} [x, y]
	 */
	function getAnchorPointOnBounds(bounds, anchorValue) {
	    var anchorRatio = getAnchorRatio(anchorValue);
	    return [
	        bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
	        bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
	    ];
	}

	/**
	 * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
	 * 基準は効果を含まない境界（geometricBounds）
	 * @param {number|string} anchorValue - 0〜8 か名前
	 * @returns {Transformation} Transformation.TOPLEFT など
	 */
	function getAnchorTransformation(anchorValue) {
	    var transformations = [
	        Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
	        Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
	        Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
	    ];
	    return transformations[resolveAnchorWidgetIndex(anchorValue)];
	}

	/**
	 * symbols.add() に渡す登録点を返す（Illustrator 専用）
	 * @param {number|string} anchorValue - 0〜8 か名前
	 * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
	 */
	function getAnchorSymbolRegistrationPoint(anchorValue) {
	    var registrationPoints = [
	        SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
	        SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
	        SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
	    ];
	    return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
	}

	/**
	 * クリック位置からセルのインデックスを求める（ウィジェットを縦横3等分し、外にはみ出した座標は端のセルに寄せる）
	 * @param {number} clickX - コントロール基準の x
	 * @param {number} clickY - コントロール基準の y
	 * @param {number} widgetWidth - ウィジェットの幅
	 * @param {number} widgetHeight - ウィジェットの高さ
	 * @returns {number} 0〜8
	 */
	function getAnchorCellAt(clickX, clickY, widgetWidth, widgetHeight) {
	    var column = Math.min(2, Math.max(0, Math.floor(clickX / (widgetWidth / 3))));
	    var row = Math.min(2, Math.max(0, Math.floor(clickY / (widgetHeight / 3))));
	    return row * 3 + column;
	}

	/**
	 * セルのインデックスの配列を、9個の真偽値に直す
	 * @param {number[]} [cellIndexes] - セルのインデックスの配列
	 * @returns {boolean[]} 含まれるセルだけ true
	 */
	function toAnchorCellFlags(cellIndexes) {
	    var cellFlags = [false, false, false, false, false, false, false, false, false];
	    if (!cellIndexes) return cellFlags;
	    for (var i = 0; i < cellIndexes.length; i++) {
	        if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
	    }
	    return cellFlags;
	}

	// -----------------------------------------
	// 描画 / Drawing
	// -----------------------------------------
	/**
	 * ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立。選択セルだけ塗る）
	 * @param {Button} anchorWidget - 描くウィジェット
	 * @returns {void}
	 */
	function drawAnchorWidget(anchorWidget) {
	    var graphics = anchorWidget.graphics;
	    var widgetWidth = anchorWidget.size[0];
	    var widgetHeight = anchorWidget.size[1];
	    var cellSize = ANCHOR_WIDGET_CELL_SIZE;
	    var halfCell = cellSize / 2;
	    /* 自作描画は自動でディムにならないので、親までたどって判定する / custom drawing is not dimmed automatically */
	    var isEnabled = isAnchorWidgetEnabledInTree(anchorWidget);

	    /* ボタンの地をコントロールの地色で塗り、パネルに溶け込ませる（backgroundColor が無い環境では例外）
	       Paint the control's own background so the widget blends into the panel; throws where backgroundColor is missing */
	    try {
	        graphics.newPath();
	        graphics.rectPath(0, 0, widgetWidth, widgetHeight);
	        graphics.fillPath(graphics.backgroundColor);
	    } catch (e) {}

	    var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
	    var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
	    var originX = Math.round((widgetWidth - gridSize) / 2);
	    var originY = Math.round((widgetHeight - gridSize) / 2);
	    var cellPositions = [];
	    var i;
	    for (i = 0; i < 9; i++) {
	        cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
	    }

	    var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1);
	    for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
	        var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
	        var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
	        graphics.newPath();
	        if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
	            /* 横方向：右隣の□へ / horizontal: to the square on the right */
	            graphics.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
	            graphics.lineTo(cellB[0], cellB[1] + halfCell);
	        } else {
	            /* 縦方向：下の□へ / vertical: to the square below */
	            graphics.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
	            graphics.lineTo(cellB[0] + halfCell, cellB[1]);
	        }
	        graphics.strokePath(linePen);
	    }

	    for (i = 0; i < 9; i++) {
	        var isCellEnabled = isEnabled && !anchorWidget.anchorDisabledCells[i];
	        drawAnchorWidgetCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.anchorWidgetIndex, isCellEnabled);
	    }
	}

	/**
	 * □を1つ描く（選択中だけ塗り、枠は塗りの上に重ねる）
	 * @param {ScriptUIGraphics} graphics - 描画先
	 * @param {number} cellX - 左端
	 * @param {number} cellY - 上端
	 * @param {boolean} isSelected - 選択中なら true
	 * @param {boolean} isEnabled - 選べるセルなら true（false なら薄く描く）
	 * @returns {void}
	 */
	function drawAnchorWidgetCell(graphics, cellX, cellY, isSelected, isEnabled) {
	    var cellSize = ANCHOR_WIDGET_CELL_SIZE;
	    /* rectPath の前には毎回 newPath()（呼ばないとパスが累積して塗りが線画になる）
	       Always call newPath() before rectPath(), or paths accumulate and fills turn into outlines */
	    if (isSelected) {
	        graphics.newPath();
	        graphics.rectPath(cellX, cellY, cellSize, cellSize);
	        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_FILL_COLOR : ANCHOR_WIDGET_DIM_FILL_COLOR));
	    }
	    graphics.newPath();
	    graphics.rectPath(cellX, cellY, cellSize, cellSize);
	    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1));
	}

	/**
	 * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
	 * @param {Object} control - 対象のコントロール
	 * @returns {boolean} すべて有効なら true
	 */
	function isAnchorWidgetEnabledInTree(control) {
	    for (var node = control; node; node = node.parent) {
	        if (node.enabled === false) return false;
	    }
	    return true;
	}

	/**
	 * ウィジェットの onDraw を呼び直す。notify("onDraw") は環境によって例外や空振りになるため、隠して再表示して描き直させる
	 * @param {Button} anchorWidget - 描き直すウィジェット
	 * @returns {void}
	 */
	function redrawAnchorWidget(anchorWidget) {
	    anchorWidget.hide();
	    anchorWidget.show();
	}

	// 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget

	// =========================================
	// ダイアログ / Dialog
	// =========================================

	/**
	 * シンボルと基準点を選ぶダイアログを作る（左にシンボル一覧、右に基準点）
	 * @returns {{show: Function}} show() でダイアログを開くオブジェクト
	 */
	function createReplaceDialog() {
		var replaceDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
		setupWindow(replaceDialog);

		/* 左：シンボル一覧、右：基準点 / Left: symbol list, right: reference point */
		var columnsGroup = replaceDialog.add('group');
		columnsGroup.orientation = 'row';
		columnsGroup.alignChildren = ['fill', 'top'];
		columnsGroup.spacing = COLUMN_SPACING;

		var symbolPanel = columnsGroup.add('panel', undefined, labelText('panel.symbol'));
		setupPanel(symbolPanel);

		var anchorPanel = columnsGroup.add('panel', undefined, labelText('panel.anchor'));
		setupPanel(anchorPanel);
		anchorPanel.alignChildren = ['center', 'top'];

		var buttonRow = addButtonRow(replaceDialog);

		buildSymbolList(symbolPanel);
		buildAnchorWidget(anchorPanel);
		buildButtons(buttonRow.rightGroup);
		alignRightOnlyButtonRow(buttonRow);

		return {
			show: function () {
				prepareDialogWindow(replaceDialog, SCRIPT_NAME);
				replaceDialog.show();
			}
		};

		/**
		 * シンボル名のラジオボタンを並べる。MAX_VISIBLE_SYMBOLS を超える分はスクロールバーで送る
		 * @param {Panel} parent - 追加先のパネル
		 * @returns {void}
		 */
		function buildSymbolList(parent) {
			var symbolListGroup = parent.add('group');
			symbolListGroup.orientation = 'row';
			symbolListGroup.alignChildren = ['left', 'top'];
			symbolListGroup.spacing = SYMBOL_LIST_SPACING;

			var symbolRadioGroup = symbolListGroup.add('group');
			symbolRadioGroup.orientation = 'column';
			symbolRadioGroup.alignChildren = 'left';
			symbolRadioGroup.spacing = SYMBOL_RADIO_SPACING;

			var visibleCount = Math.min(symbolEntries.length, MAX_VISIBLE_SYMBOLS);
			var symbolRadios = [];
			for (var i = 0; i < visibleCount; i++) {
				var symbolRadio = symbolRadioGroup.add('radiobutton', undefined, '');
				symbolRadio.helpTip = getLabel('tooltip.symbol');
				symbolRadio.onClick = onSymbolRadioClick;
				symbolRadios.push(symbolRadio);
			}
			showSymbolEntries(symbolRadios, 0);
			symbolRadios[0].active = true;

			if (symbolEntries.length > MAX_VISIBLE_SYMBOLS) {
				addSymbolScrollbar(symbolListGroup, symbolRadios);
			}
		}

		/**
		 * ラジオボタンに、スクロール位置から始まるシンボルを割り当てる
		 * @param {RadioButton[]} symbolRadios - シンボルのラジオボタン
		 * @param {number} scrollOffset - 先頭に表示する symbolEntries のインデックス
		 * @returns {void}
		 */
		function showSymbolEntries(symbolRadios, scrollOffset) {
			for (var i = 0; i < symbolRadios.length; i++) {
				var symbolEntry = symbolEntries[scrollOffset + i];
				symbolRadios[i].text = symbolEntry.name;
				symbolRadios[i].symbolIndex = symbolEntry.index;
				symbolRadios[i].value = (symbolEntry.index === replaceSettings.symbolIndex);
			}
		}

		/**
		 * シンボル一覧の横にスクロールバーを付ける
		 * @param {Group} parent - 一覧のグループ
		 * @param {RadioButton[]} symbolRadios - シンボルのラジオボタン
		 * @returns {void}
		 */
		function addSymbolScrollbar(parent, symbolRadios) {
			var maxOffset = symbolEntries.length - MAX_VISIBLE_SYMBOLS;
			var symbolScrollbar = parent.add('scrollbar', undefined, 0, 0, maxOffset);
			symbolScrollbar.preferredSize.width = SYMBOL_SCROLLBAR_WIDTH;
			symbolScrollbar.preferredSize.height = SYMBOL_RADIO_ROW_HEIGHT * MAX_VISIBLE_SYMBOLS;
			symbolScrollbar.onChanging = function () {
				showSymbolEntries(symbolRadios, Math.round(symbolScrollbar.value));
			};
		}

		/**
		 * 基準点（9軸）のウィジェットを追加する
		 * @param {Panel} parent - 追加先のパネル
		 * @returns {void}
		 */
		function buildAnchorWidget(parent) {
			var anchorWidget = addAnchorWidget(parent, replaceSettings.anchorIndex, function (anchorIndex) {
				replaceSettings.anchorIndex = anchorIndex;
				runWithErrorAlert(previewReplace);
			});
			anchorWidget.helpTip = getLabel('tooltip.anchor');
		}

		/**
		 * ［キャンセル］［実行］ボタンを追加する
		 * @param {Group} parent - ボタン行の右のグループ
		 * @returns {void}
		 */
		function buildButtons(parent) {
			var btnCancel = parent.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
			var btnOK = parent.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

			btnOK.onClick = function () {
				runWithErrorAlert(function () {
					replaceSelectionWithSymbol(false);
					replaceDialog.close();
				});
			};
			btnCancel.onClick = function () {
				replaceDialog.close();
			};
		}

		/**
		 * シンボルのラジオボタンを選んだときに、設定を更新してプレビューする
		 * @returns {void}
		 */
		function onSymbolRadioClick() {
			replaceSettings.symbolIndex = this.symbolIndex;
			runWithErrorAlert(previewReplace);
		}

		/**
		 * 置き換えを実行して描画し、取り消してプレビューにする
		 * @returns {void}
		 */
		function previewReplace() {
			replaceSelectionWithSymbol(true);
			app.redraw();
			app.undo();
		}
	}

	/**
	 * 関数を実行し、例外が起きたら内容を表示する
	 * @param {Function} task - 実行する関数
	 * @returns {void}
	 */
	function runWithErrorAlert(task) {
		try {
			task();
		} catch (e) {
			alert(getLabel('alert.error') + e);
		}
	}

	// =========================================
	// 置き換え / Replace
	// =========================================

	/**
	 * 実行できる状態かを確かめ、できなければ理由を表示する
	 * @returns {boolean} 実行してよければ true
	 */
	function canRun() {
		if (!selectedItems.length) {
			alert(getLabel('alert.noSelection'));
			return false;
		}
		if (!symbolEntries.length) {
			alert(getLabel('alert.noSymbols'));
			return false;
		}
		if (!activeLayer.visible || activeLayer.locked) {
			alert(getLabel('alert.layerUnavailable'));
			return false;
		}
		if (selectedItems.length > SLOW_SELECTION_THRESHOLD) {
			return confirm(getLabel('alert.slowSelection', [selectedItems.length]));
		}
		return true;
	}

	/**
	 * 選択中のオブジェクトを、選んだシンボルのインスタンスに置き換える
	 * @param {boolean} isPreview - true なら元のオブジェクトを消さずに隠す（あとで取り消す前提）
	 * @returns {void}
	 */
	function replaceSelectionWithSymbol(isPreview) {
		var targetSymbol = documentSymbols[replaceSettings.symbolIndex];
		for (var i = 0; i < selectedItems.length; i++) {
			var targetItem = selectedItems[i];
			var newSymbolItem = activeLayer.symbolItems.add(targetSymbol);
			alignOnTarget(newSymbolItem, targetItem, replaceSettings.anchorIndex);
			if (isPreview) {
				targetItem.hidden = true;
			} else {
				newSymbolItem.selected = true;
				targetItem.remove();
			}
		}
	}

	/**
	 * シンボルインスタンスを、元のオブジェクトの境界ボックスに基準点で揃える
	 * @param {SymbolItem} symbolItem - 動かすシンボルインスタンス
	 * @param {PageItem} targetItem - 元のオブジェクト
	 * @param {number} anchorIndex - 基準点 0〜8（4=中央）
	 * @returns {void}
	 */
	function alignOnTarget(symbolItem, targetItem, anchorIndex) {
		var targetPoint = getAnchorPointOnBounds(targetItem.geometricBounds, anchorIndex);
		var symbolPoint = getAnchorPointOnBounds(symbolItem.geometricBounds, anchorIndex);
		symbolItem.translate(targetPoint[0] - symbolPoint[0], targetPoint[1] - symbolPoint[1]);
	}

	/**
	 * 選択をオブジェクトの配列として取り出す（文字ツールで文字を選択しているときは空）
	 * @param {Document} targetDoc - 対象のドキュメント
	 * @returns {PageItem[]} 選択中のオブジェクト
	 */
	function getSelectedItems(targetDoc) {
		var selection = targetDoc.selection;
		var items = [];
		/* 文字を選択しているときは TextRange が返り、length は文字数 / Selecting characters returns a TextRange whose length is the character count */
		if (!selection || selection.typename === 'TextRange') return items;
		for (var i = 0; i < selection.length; i++) {
			items.push(selection[i]);
		}
		return items;
	}

	/**
	 * シンボルを名前順（大文字・小文字を区別しない）に並べた一覧を作る。
	 * 比較関数つきの sort() は遅く並びも狂うので、名前と番号をつないだ文字列を引数なしの sort() で並べる
	 * @param {Symbols} symbolCollection - ドキュメントのシンボル
	 * @returns {{name: string, index: number}[]} 名前と symbolCollection 上のインデックス
	 */
	function getSortedSymbolEntries(symbolCollection) {
		var SORT_KEY_SEPARATOR = '\u0000';
		var sortKeys = [];
		var i;
		for (i = 0; i < symbolCollection.length; i++) {
			/* 同名は元の順に並ぶよう、番号を6桁にそろえる / Pad the index so equal names keep their original order */
			var paddedIndex = ('000000' + i).slice(-6);
			sortKeys.push(symbolCollection[i].name.toLowerCase() + SORT_KEY_SEPARATOR + paddedIndex);
		}
		sortKeys.sort();
		var symbolEntryList = [];
		for (i = 0; i < sortKeys.length; i++) {
			var symbolIndex = parseInt(sortKeys[i].split(SORT_KEY_SEPARATOR).pop(), 10);
			symbolEntryList.push({ name: symbolCollection[symbolIndex].name, index: symbolIndex });
		}
		return symbolEntryList;
	}

	// =========================================
	// メイン処理 / Main
	// =========================================
	/* LABELS と部品の定数がそろってから実行する / Run after LABELS and the parts' constants are set */
	if (!app.documents.length) {
		alert(getLabel('alert.noDocument'));
		return;
	}
	var activeDoc = app.activeDocument;
	var activeLayer = activeDoc.activeLayer;
	var selectedItems = getSelectedItems(activeDoc);
	var documentSymbols = activeDoc.symbols;
	var symbolEntries = getSortedSymbolEntries(documentSymbols);
	if (symbolEntries.length > 0) {
		replaceSettings.symbolIndex = symbolEntries[0].index;
	}

	if (canRun()) {
		createReplaceDialog().show();
	}
}());
