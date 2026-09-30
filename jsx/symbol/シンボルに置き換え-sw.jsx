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
var SCRIPT_VERSION  = "v0.5.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Toshiyuki Takahashi";          /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/シンボルに置き換え-sw.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/シンボルに置き換え-sw.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author Toshiyuki Takahashi
 * @discussion http://www.graphicartsunit.com/
 */

(function () {

	var SCRIPT_TITLE = 'シンボルに置き換え';

	var MAX_VISIBLE_SYMBOLS = 20;
	var RADIO_ROW_HEIGHT = 20;
	var SLOW_SELECTION_THRESHOLD = 20;
	var ERROR_PREFIX = 'エラーが発生して処理を実行できませんでした\nエラー内容：';

	// Settings
	var settings = {
		'symbolIndex': 0
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

	// UI dialog
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
		panel: {
			symbol: { ja: 'シンボル', en: 'Symbol' }
		},
		button: {
			ok:     { ja: '実行', en: 'Run' },
			cancel: { ja: 'キャンセル', en: 'Cancel' }
		},
		tooltip: {
			symbol: {
				ja: '選択したオブジェクトを、このシンボルのインスタンスに置き換えます。',
				en: 'Replaces the selected objects with an instance of this symbol.'
			}
		}
	};

	// ボタン行（再利用パーツ） / Button row (reusable)

	var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
	var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
		btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

	// ボタン行（再利用パーツ）ここまで / End of the reusable button row

	function createDialog() {
		var window = new Window('dialog', SCRIPT_TITLE + ' ' + SCRIPT_VERSION);

		var symbolPanel = window.add('panel', undefined, labelText('panel.symbol'));
		symbolPanel.alignment = 'left';
		symbolPanel.margins = [15, 20, 15, 10];
		symbolPanel.orientation = 'column';
		symbolPanel.alignChildren = 'left';

		var buttonRow = addButtonRow(window, { centered: true });

		var symbolRadios = buildSymbolList(symbolPanel);
		buildButtons(buttonRow.rowGroup);

		return {
			show: function () {
				prepareDialogWindow(window, SCRIPT_NAME);
				window.show();
			}
		};

		function buildSymbolList(parent) {
			var listContainer = parent.add('group');
			listContainer.orientation = 'row';
			listContainer.alignChildren = 'top';
			listContainer.spacing = 4;

			var radioGroup = listContainer.add('group');
			radioGroup.orientation = 'column';
			radioGroup.alignChildren = 'left';
			radioGroup.spacing = 2;

			var visibleCount = Math.min(symbolEntries.length, MAX_VISIBLE_SYMBOLS);
			var radios = [];
			for (var i = 0; i < visibleCount; i++) {
				var entry = symbolEntries[i];
				var radio = radioGroup.add('radiobutton', undefined, entry.name);
				radio.helpTip = getLabel('tooltip.symbol');
				radio.symbolIndex = entry.index;
				radio.value = (entry.index === settings.symbolIndex);
				radio.onClick = onRadioClick;
				radios.push(radio);
			}
			if (radios.length > 0) {
				radios[0].active = true;
			}

			if (symbolEntries.length > MAX_VISIBLE_SYMBOLS) {
				addScrollbar(listContainer, radios);
			}
			return radios;
		}

		function addScrollbar(parent, radios) {
			var maxOffset = symbolEntries.length - MAX_VISIBLE_SYMBOLS;
			var scrollbar = parent.add('scrollbar', undefined, 0, 0, maxOffset);
			scrollbar.preferredSize.width = 16;
			scrollbar.preferredSize.height = RADIO_ROW_HEIGHT * MAX_VISIBLE_SYMBOLS;
			scrollbar.onChanging = function () {
				var offset = Math.round(scrollbar.value);
				for (var idx = 0; idx < radios.length; idx++) {
					var entry = symbolEntries[offset + idx];
					radios[idx].text = entry.name;
					radios[idx].symbolIndex = entry.index;
					radios[idx].value = (entry.index === settings.symbolIndex);
				}
			};
		}

		function buildButtons(parent) {
			var btnCancel = parent.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
			var btnOK = parent.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

			btnOK.onClick = function () {
				try {
					replaceSelectionWithSymbol(false);
					window.close();
				} catch (e) {
					alert(ERROR_PREFIX + e);
				}
			};
			btnCancel.onClick = function () {
				window.close();
			};
		}

		function onRadioClick() {
			try {
				settings.symbolIndex = this.symbolIndex;
				previewReplace();
			} catch (e) {
				alert(ERROR_PREFIX + e);
			}
		}

		function previewReplace() {
			replaceSelectionWithSymbol(true);
			app.redraw();
			app.undo();
		}
	}

	// Pre-flight check before showing the dialog
	function canRun() {
		if (!activeDoc || selectedItems.length < 1) {
			alert('オブジェクトが選択されていません');
			return false;
		}
		if (!activeLayer.visible || activeLayer.locked) {
			alert('選択レイヤーがロックされているか非表示になっています');
			return false;
		}
		if (selectedItems.length > SLOW_SELECTION_THRESHOLD) {
			return confirm(selectedItems.length + '個のオブジェクトが選択されており、処理にとても時間がかかる可能性があります。継続しますか？');
		}
		return true;
	}

	// Main process
	function replaceSelectionWithSymbol(isPreview) {
		var targetItems = toItemArray(selectedItems);
		for (var i = 0; i < targetItems.length; i++) {
			var newSymbolItem = activeLayer.symbolItems.add(documentSymbols[settings.symbolIndex]);
			centerOnTarget(newSymbolItem, targetItems[i]);
			if (isPreview) {
				selectedItems[i].hidden = true;
			} else {
				newSymbolItem.selected = true;
				selectedItems[i].remove();
			}
		}
	}

	// Center the symbol item on the target item's bounding box
	function centerOnTarget(symbolItem, targetItem) {
		var t = targetItem.geometricBounds;
		var s = symbolItem.geometricBounds;
		symbolItem.top = (t[1] + t[3]) / 2 - (s[3] - s[1]) / 2;
		symbolItem.left = (t[0] + t[2]) / 2 - (s[2] - s[0]) / 2;
	}

	// Convert a collection (selection / PageItems) to a plain Array
	function toItemArray(collection) {
		var items = [];
		for (var i = 0; i < collection.length; i++) {
			items.push(collection[i]);
		}
		return items;
	}

	// Build {name, index} entries sorted by name (case-insensitive)
	function getSortedSymbolEntries(symbolCollection) {
		var entries = [];
		for (var i = 0; i < symbolCollection.length; i++) {
			entries.push({ name: symbolCollection[i].name, index: i });
		}
		entries.sort(function (a, b) {
			var an = a.name.toLowerCase();
			var bn = b.name.toLowerCase();
			if (an < bn) return -1;
			if (an > bn) return 1;
			return 0;
		});
		return entries;
	}
	// =========================================
	// メイン処理 / Main
	// =========================================
	/* LABELS と部品の定数がそろってから実行する / Run after LABELS and the parts' constants are set */
	var activeDoc = app.activeDocument;
	var activeLayer = activeDoc.activeLayer;
	var selectedItems = activeDoc.selection;
	var documentSymbols = activeDoc.symbols;
	var symbolEntries = getSortedSymbolEntries(documentSymbols);
	if (symbolEntries.length > 0) {
		settings.symbolIndex = symbolEntries[0].index;
	}

	if (canRun()) {
		createDialog().show();
	}
}());
