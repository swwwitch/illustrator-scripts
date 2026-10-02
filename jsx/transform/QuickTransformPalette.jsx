#target illustrator
#targetengine "QuickTransformPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの移動・複製と反転・回転を、アイコンのクリックで即時実行する常駐パレットです。
9軸の基準点・マージン・プレビュー境界の設定はすべての操作に共通で、Option＋クリックすると複製してから変形します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/QuickTransformPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n277bd0865986

### Overview

A persistent palette that moves, duplicates, flips and rotates the selection immediately on an icon click.
The nine-point reference, margin and preview-bounds settings are shared by every operation, and Option-clicking duplicates before transforming.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/QuickTransformPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "QuickTransformPalette";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/QuickTransformPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/QuickTransformPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n277bd0865986"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    (function () {

    	/* すでにパレットが開いていれば前面に出して終了。$.global を触る前に判定する / If a palette is already open, bring it forward and return; checked before touching $.global */
    	try {
    		if ($.global.__quickTransformPalette) {
    			$.global.__quickTransformPalette.show();
    			$.global.__quickTransformPalette.active = true;
    			return;
    		}
    	} catch (staleReferenceError) {
    		$.global.__quickTransformPalette = null; /* 参照が無効なら作り直す / stale reference: rebuild */
    	}

    	// =========================================
    	// ユーザー設定 / User settings
    	// =========================================
    	var DEFAULT_MARGIN         = 0;     /* マージン欄の初期値（定規の単位）/ Initial margin value (in ruler units) */
    	var DEFAULT_PREVIEW_BOUNDS = true;  /* プレビュー境界チェックの初期状態 / Initial state of the preview-bounds checkbox */
    	var ROTATE_ANGLE           = 90;    /* 回転アイコンのクリックで回す角度（度）。正＝反時計回り / Angle per rotate-icon click (deg); positive = CCW */

    	// =========================================
    	// レイアウト / Layout
    	// =========================================

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

    	var ICON_SIZE             = 30;               /* 方向・反転・回転アイコン1個の大きさ（px、両パネル共通）/ size of each direction / flip / rotate icon (px, shared) */
    	var ICON_LINE_WIDTH       = 1;                /* アイコン内の線画の線幅（ボタン枠・9軸グリッドと同じ1）/ stroke width of the line art inside each icon */
    	var ICON_GAP              = 7;                /* 反転・回転アイコンどうしの間隔 / gap between flip/rotate icons */
    	var ICON_COLUMNS_PER_ROW  = 2;                /* 反転・回転アイコンを何個ごとに改行するか / flip/rotate icons per row before wrapping */
    	var CROSS_GAP             = 2;                /* 方向ボタン（十字）どうしの間隔 / gap between direction buttons */
    	var GROUP_SPACING         = 12;               /* アイコン群と9軸ウィジェットの間隔 / gap between the icon grid and the anchor widget */
    	var LABEL_FIELD_SPACING   = 4;                /* ラベルと入力欄の間隔（既定は広すぎる）/ gap between a label and its field (the default looks too wide) */
    	var SLIDER_ROW_SPACING    = 6;                /* 角度表示とスライダーの間隔 / gap between the angle readout and the slider */
    	var ANGLE_LABEL_WIDTH     = 40;               /* 角度表示の幅 / width of the angle readout */
    	var SLIDER_WIDTH          = 120;              /* 回転スライダーの幅 / rotate slider width */
    	var SLIDER_ROW_HEIGHT     = 18;               /* 角度表示とスライダーの高さ / height of the angle readout and the slider */
    	var FIELD_CHARS           = 3;                /* 数値入力欄の文字数 / width of numeric fields */

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

    	var LABELS = {
    		dialog: {
    			title: { ja: "クイック変形", en: "Quick Transform" }
    		},
    		panel: {
    			direction: { ja: "移動・複製", en: "Move / Duplicate" },
    			flip: { ja: "反転・回転", en: "Flip / Rotate" },
    			options: { ja: "オプション", en: "Options" }
    		},
    		direction: {
    			up:    { ja: "上", en: "Up" },
    			left:  { ja: "左", en: "Left" },
    			right: { ja: "右", en: "Right" },
    			down:  { ja: "下", en: "Down" }
    		},
    		fieldLabel: {
    			margin: { ja: "マージン", en: "Margin" }
    		},
    		checkbox: {
    			previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
    		},
    		tooltip: {
    			flipHorizontal: { ja: "選択を水平方向に反転（基準点が基点）／ Option＋クリックで複製", en: "Flip the selection horizontally (about the anchor point) / Option-click to duplicate" },
    			flipVertical: { ja: "選択を垂直方向に反転（基準点が基点）／ Option＋クリックで複製", en: "Flip the selection vertically (about the anchor point) / Option-click to duplicate" },
    			rotateCCW: { ja: "選択を反時計回りに90°回転（基準点が基点）／ Option＋クリックで複製", en: "Rotate the selection 90° counterclockwise (about the anchor point) / Option-click to duplicate" },
    			rotateCW: { ja: "選択を時計回りに90°回転（基準点が基点）／ Option＋クリックで複製", en: "Rotate the selection 90° clockwise (about the anchor point) / Option-click to duplicate" },
    			anchor: { ja: "基準点（反転・回転の基点）", en: "Anchor point (pivot for flip / rotate)" },
    			moveDuplicate: { ja: "クリックで移動／ Option＋クリックで複製", en: "Click to move / Option-click to duplicate" },
    			duplicateInPlace: { ja: "複製", en: "Duplicate" },
    			rotateSlider: { ja: "スライダーで回転（-180〜180°・15°刻み／Shift＝90°刻み、正＝反時計回り、基準点が基点）", en: "Rotate with the slider (-180 to 180°, 15° steps / Shift = 90° steps, positive = CCW, about the anchor point)" },
    			stepUp: {
    				ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
    				en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
    			},
    			stepDown: {
    				ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
    				en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
    			},
    			stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
    			stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
    		}
    	};

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    	var FALLBACK_RULER_UNIT_INDEX = 2; /* 判別できないときは pt 扱い / Fall back to pt */

    	// =========================================
    	// 状態 / State
    	// =========================================
    	/* パレット本体は $.global.__quickTransformPalette に持たせる（IIFE 内の var では GC される）/ The palette itself lives on $.global.__quickTransformPalette (an IIFE-local var would be garbage-collected) */
    	var isBusy = false;              /* 委譲の再入防止 / Re-entrancy guard for delegation */
    	var pivotAnchorWidget = null;    /* 基準点（9軸）のウィジェット。値は getAnchorWidgetIndex() で 0..8 を行優先（0=左上, 4=中央, 8=右下）/ 9-axis anchor widget; getAnchorWidgetIndex() gives 0..8 row-major (0=top-left, 4=center, 8=bottom-right) */
    	var readTransformOptions = null; /* buildOptionsPanel() が公開する設定読取り関数（未生成の間は null）/ Settings reader published by buildOptionsPanel() (null until built) */

    	// =========================================
    	// worker（メインエンジンで実行する DOM 処理）/ Worker (DOM work run on the main engine)
    	// 注意: JSDoc・行コメント（//）禁止・/* */ のみ・各文はセミコロンで終える（toString が改行を消すため）
    	// Note: no JSDoc, no // comments, /* */ only, end every statement with ';' (toString strips newlines)
    	// =========================================
    	/* 選択に対して移動・複製を1回実行する（アイコンのクリックで即時適用）/ Apply the move/duplicate once to the selection (immediate on icon click) */
    	function workerApply(options) {
    		if (app.documents.length === 0) { return "NODOC"; };
    		var docSelection = app.activeDocument.selection;
    		if (!docSelection || docSelection.length === 0) { return "NOSEL"; };
    		try {
    			var selectionSize = workerSelectionSize(docSelection, options.usePreviewBounds);
    			/* 選択はあるが1つも境界を測れなかった場合は何もせず抜ける / Bail out when there is a selection but nothing could be measured */
    			if (!selectionSize) { return "NOBOUNDS"; };
    			var offset = workerOffset(options.direction, selectionSize.width, selectionSize.height, options.marginPt);
    			var createdItems = workerBuild(docSelection, offset, options.duplicate);
    			/* 複製時は元の選択を外し、生成した複製だけを選択状態にする / On duplicate, deselect the originals and select only the created copies */
    			if (options.duplicate && createdItems.length > 0) {
    				try { app.activeDocument.selection = createdItems; } catch (selectionError) {};
    			};
    			app.redraw();
    			return "OK";
    		} catch (e) {
    			return "ERR:" + e;
    		};
    	};

    	/* 選択範囲全体のサイズを求める。境界を取れない項目（ロック・非表示・ガイド等）は読み飛ばし、1つも測れなければ null を返す */
    	/* Get the overall size of the selection; skip items whose bounds fail (locked, hidden, guides, ...) and return null if nothing could be measured */
    	function workerSelectionSize(items, usePreviewBounds) {
    		var minLeft = null, maxRight = null, minBottom = null, maxTop = null;
    		for (var i = 0; i < items.length; i++) {
    			try {
    				var bounds = usePreviewBounds ? items[i].visibleBounds : items[i].geometricBounds;
    				if (minLeft === null || bounds[0] < minLeft) { minLeft = bounds[0]; };
    				if (maxTop === null || bounds[1] > maxTop) { maxTop = bounds[1]; };
    				if (maxRight === null || bounds[2] > maxRight) { maxRight = bounds[2]; };
    				if (minBottom === null || bounds[3] < minBottom) { minBottom = bounds[3]; };
    			} catch (boundsError) {};
    		};
    		if (minLeft === null) { return null; };
    		return { width: maxRight - minLeft, height: maxTop - minBottom };
    	};

    	/* 方向キーを1ステップの移動量（dx, dy）へ変換。幅／高さにマージンを加算 / Convert a direction key to a one-step offset, adding margin */
    	function workerOffset(direction, width, height, margin) {
    		if (direction === "up") { return [0, height + margin]; };
    		if (direction === "down") { return [0, -(height + margin)]; };
    		if (direction === "left") { return [-(width + margin), 0]; };
    		if (direction === "right") { return [width + margin, 0]; };
    		return [0, 0];
    	};

    	/* 複製（またはそのまま）を1つ分だけ移動し、対象になった項目の配列を返す。複製・移動できない項目は読み飛ばす */
    	/* Move by one step (duplicating or not) and return the affected items; skip items that cannot be duplicated or moved */
    	/* 複製できた項目は移動に失敗しても配列へ入れる（選択から漏れて置き去りにならないように）*/
    	/* A copy that was created is kept in the array even if the move fails, so it is not left behind outside the selection */
    	function workerBuild(items, offset, shouldDuplicate) {
    		var createdItems = [];
    		for (var i = 0; i < items.length; i++) {
    			try {
    				var target = shouldDuplicate ? items[i].duplicate() : items[i];
    				createdItems.push(target);
    				target.translate(offset[0], offset[1]);
    			} catch (buildError) {};
    		};
    		return createdItems;
    	};

    	/* 委譲する worker 関数の全登録（追加漏れ防止）/ All worker functions to delegate (avoid missing registrations) */
    	var WORKER_FUNCTIONS = [workerApply, workerSelectionSize, workerOffset, workerBuild];

    	// =========================================
    	// BridgeTalk 委譲 / BridgeTalk delegation
    	// =========================================
    	/**
    	 * 関数のソースから宣言行〜閉じ括弧行だけを切り出す
    	 * ExtendScript の toString() は改行を CR で返し、周辺のコメント断片まで巻き込むことがあるため、
    	 * 行区切りを LF に正規化したうえで関数本体だけを取り出す
    	 * @param {function} targetFunction - 文字列化する関数
    	 * @returns {string} 関数宣言だけのソース文字列
    	 */
    	function sliceFunctionSource(targetFunction) {
    		var lines = String(targetFunction).replace(/\r\n?/g, "\n").split("\n");
    		var firstIndex = -1;
    		var lastIndex = -1;
    		for (var i = 0; i < lines.length; i++) {
    			if (firstIndex < 0 && /^\s*function\s/.test(lines[i])) { firstIndex = i; }
    			if (firstIndex >= 0 && /^\s*\}[;\s]*$/.test(lines[i])) { lastIndex = i; }
    		}
    		if (firstIndex < 0) { return String(targetFunction); }
    		if (lastIndex < firstIndex) {
    			/* 1行で書かれた関数は、その行だけを取り出す / A function written on one line: keep just that line */
    			return /\}[;\s]*$/.test(lines[firstIndex]) ? lines[firstIndex] : lines.slice(firstIndex).join("\n");
    		}
    		return lines.slice(firstIndex, lastIndex + 1).join("\n");
    	}

    	/**
    	 * worker 関数群をソース文字列に連結する
    	 * @returns {string} 連結したソース文字列
    	 */
    	function buildWorkerSource() {
    		var source = "";
    		for (var i = 0; i < WORKER_FUNCTIONS.length; i++) {
    			source += sliceFunctionSource(WORKER_FUNCTIONS[i]) + "\n";
    		}
    		return source;
    	}

    	/**
    	 * メインエンジン（#targetengine 指定なし）へコードを送って同期実行する
    	 * 同期 send なので isBusy による再入防止が有効に働く
    	 * @param {string} bodyCode - メインエンジンで評価するコード
    	 * @returns {string} 結果マーカー（"OK" / "NODOC" / "NOSEL" / "NOBOUNDS" / "ERR:..."）
    	 */
    	function sendToMainEngine(bodyCode) {
    		var resultHolder = { result: "ERR:timeout" };
    		try {
    			var bridge = new BridgeTalk();
    			bridge.target = "illustrator";
    			bridge.body = bodyCode;
    			bridge.onResult = function (message) { resultHolder.result = String(message.body); };
    			bridge.onError = function (message) {
    				resultHolder.result = "ERR:" + String(message.body);
    				try { $.writeln(SCRIPT_NAME + " BridgeTalk error: " + message.body); } catch (e) {}
    			};
    			bridge.send(10); /* 完了まで待つ / wait for completion */
    		} catch (bridgeError) {
    			/* BridgeTalk が使えない環境ではこのエンジンで直接実行。worker の戻り値をそのまま結果にする（値を返さない body は "OK" 扱い）*/
    			/* Fallback: run directly in this engine, keeping the worker's own return value (a body that returns nothing counts as "OK") */
    			try {
    				var evalResult = eval(bodyCode);
    				resultHolder.result = (evalResult === undefined) ? "OK" : String(evalResult);
    			} catch (evalError) {
    				resultHolder.result = "ERR:" + evalError;
    			}
    		}
    		return resultHolder.result;
    	}

    	/**
    	 * オプションを JS オブジェクトリテラル文字列へ変換する
    	 * @param {{direction: string, marginPt: number, duplicate: boolean, usePreviewBounds: boolean}} options - 変換するオプション
    	 * @returns {string} オブジェクトリテラル文字列
    	 */
    	function optionsLiteral(options) {
    		return "{" +
    			"direction:'" + options.direction + "'," +
    			"marginPt:" + options.marginPt + "," +
    			"duplicate:" + (options.duplicate ? "true" : "false") + "," +
    			"usePreviewBounds:" + (options.usePreviewBounds ? "true" : "false") +
    		"}";
    	}

    	/**
    	 * 移動・複製の即時実行をメインエンジンへ委譲する
    	 * @param {{direction: string, marginPt: number, duplicate: boolean, usePreviewBounds: boolean}} options - 実行オプション
    	 * @returns {string} 結果マーカー（"OK" / "NODOC" / "NOSEL" / "NOBOUNDS" / "ERR:..."）
    	 */
    	function delegateApply(options) {
    		var payload = buildWorkerSource() + "workerApply(" + optionsLiteral(options) + ");";
    		return sendToMainEngine('eval(decodeURIComponent("' + encodeURIComponent(payload) + '"));');
    	}

    	// =========================================
    	// 反転・回転の委譲（メインエンジンで DOM を変形）/ Flip & rotate delegation (transform the DOM on the main engine)
    	// =========================================
    	/**
    	 * 選択を、9軸の基準点を基点に1回の合成行列で変形する（可視／幾何境界の測定→基準点→変形→マージン→再描画）
    	 * matrixCode が matrix を組み立てる。境界が取れない／変形できない項目（ロック・非表示・ガイド等）は読み飛ばす
    	 * duplicate=true のときは変形前に選択を複製し、複製側だけを変形して新しい選択にする（Option＋クリック）
    	 * マージンは変形後に基準点から中心の反対方向へ平行移動して足す（中心基点＝index 4 では 0 になり無視される）
    	 * @param {string} matrixCode - matrix を組み立てるコード片（anchorX / anchorY を参照できる）
    	 * @param {boolean} duplicate - true で複製してから変形する
    	 * @param {number} marginPt - 変形後に足す余白（pt）
    	 * @param {boolean} usePreviewBounds - true で可視境界、false で幾何境界から基準点を測る
    	 * @returns {void}
    	 */
    	function btTransformSelection(matrixCode, duplicate, marginPt, usePreviewBounds) {
    		var dupFlag = duplicate ? 'true' : 'false';
    		var boundsProp = usePreviewBounds ? 'visibleBounds' : 'geometricBounds';
    		var anchorIndex = pivotAnchorWidget ? getAnchorWidgetIndex(pivotAnchorWidget) : 4;
    		var col = anchorIndex % 3;
    		var row = Math.floor(anchorIndex / 3);
    		var margin = Number(marginPt) || 0;
    		var marginX = ((col === 0) ? -1 : ((col === 2) ? 1 : 0)) * margin;
    		var marginY = ((row === 0) ? 1 : ((row === 2) ? -1 : 0)) * margin;
    		/* 入れ子三項は ExtendScript が左結合で誤評価するため右結合を括弧で明示（無いと左＝中央・上＝中央になる）/ Parenthesize nested ternaries; ExtendScript misparses them left-associatively (else left==center, top==center) */
    		var anchorXExpr = (col === 0) ? "left" : ((col === 1) ? "((left+right)/2)" : "right");
    		var anchorYExpr = (row === 0) ? "top" : ((row === 1) ? "((top+bottom)/2)" : "bottom");
    		var body = '' +
    			'if(app.documents.length>0){' +
    			'var doc=app.activeDocument,selection=doc.selection;' +
    			'if(selection&&selection.length>0){' +
    			'var left=Infinity,top=-Infinity,right=-Infinity,bottom=Infinity,measured=false;' +
    			'for(var i=0;i<selection.length;i++){try{var b=selection[i].' + boundsProp + ';if(b[0]<left)left=b[0];if(b[1]>top)top=b[1];if(b[2]>right)right=b[2];if(b[3]<bottom)bottom=b[3];measured=true;}catch(e){}}' +
    			'if(measured){' +
    			'var anchorX=' + anchorXExpr + ',anchorY=' + anchorYExpr + ';' +
    			'var matrix=app.getIdentityMatrix();' + matrixCode +
    			'var targets=[];' +
    			'if(' + dupFlag + '){for(var i=0;i<selection.length;i++){try{targets.push(selection[i].duplicate());}catch(e){}}}else{for(var i=0;i<selection.length;i++){targets.push(selection[i]);}}' +
    			'for(var i=0;i<targets.length;i++){try{targets[i].transform(matrix,true,true,true,true,1,Transformation.DOCUMENTORIGIN);}catch(e){}}' +
    			'var marginX=' + marginX + ',marginY=' + marginY + ';' +
    			'if(marginX!==0||marginY!==0){for(var i=0;i<targets.length;i++){try{targets[i].translate(marginX,marginY);}catch(e){}}}' +
    			'if(' + dupFlag + '){try{doc.selection=targets;}catch(e){}}' +
    			'app.redraw();' +
    			'}}}';
    		sendToMainEngine(body);
    	}

    	/**
    	 * 選択を、9軸の基準点を基点に反転する（水平＝-100,100／垂直＝100,-100）
    	 * 係数は数値化して埋め込む（'1-' + (-1) だと "1--1" になりデクリメント解釈で構文エラーになるため）
    	 * @param {number} scaleX - X 方向の倍率（%）
    	 * @param {number} scaleY - Y 方向の倍率（%）
    	 * @param {boolean} duplicate - true で複製してから反転する
    	 * @param {number} marginPt - 反転後に足す余白（pt）
    	 * @param {boolean} usePreviewBounds - true で可視境界から基準点を測る
    	 * @returns {void}
    	 */
    	function btFlipSelection(scaleX, scaleY, duplicate, marginPt, usePreviewBounds) {
    		var scaleFractionX = Number(scaleX) / 100; /* -1 or 1 */
    		var scaleFractionY = Number(scaleY) / 100;
    		btTransformSelection(
    			'matrix.mValueA=' + scaleFractionX + ';matrix.mValueD=' + scaleFractionY + ';' +
    			'matrix.mValueTX=anchorX*' + (1 - scaleFractionX) + ';matrix.mValueTY=anchorY*' + (1 - scaleFractionY) + ';',
    			duplicate, marginPt, usePreviewBounds
    		);
    	}

    	/**
    	 * 選択を、9軸の基準点を基点に回転する
    	 * @param {number} angleDegrees - 回転角（度）。正＝反時計回り／負＝時計回り
    	 * @param {boolean} duplicate - true で複製してから回転する
    	 * @param {number} marginPt - 回転後に足す余白（pt）
    	 * @param {boolean} usePreviewBounds - true で可視境界から基準点を測る
    	 * @returns {void}
    	 */
    	function btRotateSelection(angleDegrees, duplicate, marginPt, usePreviewBounds) {
    		var radians = Number(angleDegrees) * Math.PI / 180;
    		var cosAngle = Math.cos(radians);
    		var sinAngle = Math.sin(radians);
    		var oneMinusCos = 1 - cosAngle;   /* 係数は数値化（"1--0.7" のような構文エラーを避ける）/ Precompute to avoid "1--0.7"-style syntax errors */
    		btTransformSelection(
    			'matrix.mValueA=' + cosAngle + ';matrix.mValueB=' + sinAngle + ';matrix.mValueC=' + (-sinAngle) + ';matrix.mValueD=' + cosAngle + ';' +
    			'matrix.mValueTX=anchorX*' + oneMinusCos + '+anchorY*' + sinAngle + ';' +
    			'matrix.mValueTY=anchorY*' + oneMinusCos + '-anchorX*' + sinAngle + ';',
    			duplicate, marginPt, usePreviewBounds
    		);
    	}

    	// =========================================
    	// 配色 / Colors
    	// =========================================
    	/* アイコンの配色（initIconColors() で UI 明暗から設定）/ Icon colors (set from the light/dark UI in initIconColors()) */
    	var iconColor, iconBorderColor, iconBaseBg, iconHoverBg;

    	/**
    	 * UI 明度（0..1）を取得する
    	 * @returns {number} 0〜1 にクランプした明度（取得失敗時は 0＝暗い側）
    	 */
    	function getUIBrightness() {
    		try {
    			var brightness = app.preferences.getRealPreference("uiBrightness");
    			if (brightness < 0) { brightness = 0; }
    			if (brightness > 1) { brightness = 1; }
    			return brightness;
    		} catch (e) {
    			return 0;
    		}
    	}

    	/**
    	 * UI が明るいテーマかを判定する
    	 * @returns {boolean} 明るいテーマなら true（取得失敗時は false＝暗い側）
    	 */
    	function isLightUI() {
    		return !isDarkUI();
    	}

    	/**
    	 * グレーの RGBA を作る
    	 * @param {number} value - 明度（0..1 にクランプ）
    	 * @returns {number[]} [r, g, b, a] の配列
    	 */
    	function grayColor(value) {
    		if (value < 0) { value = 0; }
    		if (value > 1) { value = 1; }
    		return [value, value, value, 1];
    	}

    	/**
    	 * UI の明暗に合わせてアイコン色・背景色・枠線を決める（showPalette() から呼ぶ）
    	 * @returns {void}
    	 */
    	function initIconColors() {
    		var lightUI = isLightUI();
    		var uiBrightness = getUIBrightness();
    		iconColor       = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
    		/* ライトは薄いグレーの枠、ダークは背景より少し明るいグレーの枠でボタンの輪郭を出す / Light: light gray border; dark: gray slightly brighter than the background so the edge shows */
    		iconBorderColor = lightUI ? [0.65, 0.65, 0.65, 1] : [0.45, 0.45, 0.45, 1];
    		iconBaseBg      = lightUI ? grayColor(uiBrightness)        : [0.28, 0.28, 0.28, 1];
    		/* マウスオーバー時の背景（ライトは少し暗く、ダークは少し明るく）/ Hover background (slightly darker in light, lighter in dark) */
    		iconHoverBg     = lightUI ? grayColor(uiBrightness - 0.10) : [0.38, 0.38, 0.38, 1];
    	}

    	// =========================================
    	// 描画ヘルパー / Drawing helpers
    	// =========================================
    	/**
    	 * 正方形のパスを作る（塗り／線は呼び出し側で行う）
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {number} x - 左端
    	 * @param {number} y - 上端
    	 * @param {number} size - 一辺の長さ
    	 * @returns {void}
    	 */
    	function squarePath(graphics, x, y, size) {
    		graphics.newPath();
    		graphics.moveTo(x, y);
    		graphics.lineTo(x + size, y);
    		graphics.lineTo(x + size, y + size);
    		graphics.lineTo(x, y + size);
    		graphics.closePath();
    	}

    	/**
    	 * 3点の三角形を塗り or 線で描く
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {Array<number[]>} points - 頂点3つの [x, y] 配列
    	 * @param {number[]} color - RGBA の配列
    	 * @param {boolean} fill - true で塗り、false で輪郭線
    	 * @returns {void}
    	 */
    	function drawTriangle(graphics, points, color, fill) {
    		graphics.newPath();
    		graphics.moveTo(points[0][0], points[0][1]);
    		graphics.lineTo(points[1][0], points[1][1]);
    		graphics.lineTo(points[2][0], points[2][1]);
    		graphics.closePath();
    		if (fill) {
    			graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
    		} else {
    			graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
    		}
    	}

    	/**
    	 * 水平または垂直の点線を描く
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {number} x1 - 始点 X
    	 * @param {number} y1 - 始点 Y
    	 * @param {number} x2 - 終点 X
    	 * @param {number} y2 - 終点 Y
    	 * @param {number[]} color - RGBA の配列
    	 * @returns {void}
    	 */
    	function drawDottedLine(graphics, x1, y1, x2, y2, color) {
    		var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
    		var isHorizontal = (y1 === y2);
    		var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
    		var dashStep = 3;

    		for (var pos = 0; pos < totalLength; pos += dashStep) {
    			graphics.newPath();
    			if (isHorizontal) {
    				graphics.moveTo(x1 + pos, y1);
    				graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
    			} else {
    				graphics.moveTo(x1, y1 + pos);
    				graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
    			}
    			graphics.strokePath(pen);
    		}
    	}

    	/**
    	 * 円弧を線分で近似して1本の連続パスで描く（継ぎ目が出ないよう単一パスにする）
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {number[]} color - RGBA の配列
    	 * @param {number} centerX - 中心 X
    	 * @param {number} centerY - 中心 Y
    	 * @param {number} radius - 半径
    	 * @param {number} startDeg - 開始角（度）
    	 * @param {number} endDeg - 終了角（度）
    	 * @param {number} mirrorSign - 1 でそのまま、-1 で X をミラー
    	 * @returns {void}
    	 */
    	function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
    		var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
    		graphics.newPath();
    		for (var i = 0; i <= segments; i++) {
    			var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
    			var x = centerX + mirrorSign * radius * Math.cos(rad);
    			var y = centerY + radius * Math.sin(rad);
    			if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
    		}
    		graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
    	}

    	/**
    	 * 円弧に沿って四角い点線を描く
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {number[]} color - RGBA の配列
    	 * @param {number} centerX - 中心 X
    	 * @param {number} centerY - 中心 Y
    	 * @param {number} radius - 半径
    	 * @param {number} startDeg - 開始角（度）
    	 * @param {number} endDeg - 終了角（度）
    	 * @param {number} mirrorSign - 1 でそのまま、-1 で X をミラー
    	 * @returns {void}
    	 */
    	function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
    		var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
    		var stepDeg = 13;
    		var dashHalf = 0.9;

    		for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
    			var rad = deg * Math.PI / 180;
    			var x = centerX + mirrorSign * radius * Math.cos(rad);
    			var y = centerY + radius * Math.sin(rad);
    			var tangentX = mirrorSign * Math.sin(rad);
    			var tangentY = -Math.cos(rad);

    			graphics.newPath();
    			graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
    			graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
    			graphics.strokePath(pen);
    		}
    	}

    	/**
    	 * 右向き矢印の1点を、方向キーに合わせて回転する
    	 * @param {string} directionKey - "up" / "left" / "right" / "down"
    	 * @param {number} x - 回転前の X
    	 * @param {number} y - 回転前の Y
    	 * @returns {number[]} 回転後の [x, y]
    	 */
    	function transformArrowPoint(directionKey, x, y) {
    		if (directionKey === 'left') { return [-x, y]; }
    		if (directionKey === 'down') { return [-y, x]; }
    		if (directionKey === 'up')   { return [y, -x]; }
    		return [x, y]; /* right */
    	}

    	/**
    	 * 方向キーの矢印を白抜き（アウトライン）で描画する
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {string} directionKey - "up" / "left" / "right" / "down"
    	 * @param {number} width - ボタンの幅
    	 * @param {number} height - ボタンの高さ
    	 * @param {number[]} color - RGBA の配列
    	 * @returns {void}
    	 */
    	function drawArrow(graphics, directionKey, width, height, color) {
    		var iconSize = Math.min(width, height);
    		var tip = iconSize * 0.32;
    		var shaft = iconSize * 0.11;
    		var headHalf = iconSize * 0.27;       /* 矢じりの半分の高さ（大きめ）/ half height of the arrowhead (larger) */
    		var headBase = tip - iconSize * 0.34; /* 矢じりの付け根（長め）/ base of the arrowhead (longer) */
    		var basePoints = [
    			[-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
    			[tip, 0],
    			[headBase, headHalf], [headBase, shaft], [-tip, shaft]
    		];
    		var centerX = width / 2, centerY = height / 2;
    		graphics.newPath();
    		for (var i = 0; i < basePoints.length; i++) {
    			var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
    			if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
    			else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
    		}
    		graphics.closePath();
    		/* 白抜き：塗らずに輪郭線だけ描く / Knockout: stroke the outline only, no fill */
    		graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
    	}

    	/**
    	 * 複製アイコン（重なる2つの四角。奥＝右上・手前＝左下）を白抜きで描画する
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {number} width - ボタンの幅
    	 * @param {number} height - ボタンの高さ
    	 * @param {number[]} color - 線の RGBA
    	 * @param {number[]} backgroundColor - 手前の四角の下の奥線を消す背景色の RGBA
    	 * @returns {void}
    	 */
    	function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
    		var squareSize = Math.min(width, height) * 0.40;
    		var shift = squareSize * 0.36;              /* 2枚のずらし量 / Offset between the two squares */
    		var pairSize = squareSize + shift;
    		var left = Math.round((width - pairSize) / 2);
    		var top = Math.round((height - pairSize) / 2);
    		var backX = left + shift, backY = top;      /* 奥の四角（右上）/ Back square (upper-right) */
    		var frontX = left, frontY = top + shift;    /* 手前の四角（左下）/ Front square (lower-left) */
    		var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);

    		/* 奥の四角を輪郭線で描く / Outline the back square */
    		squarePath(graphics, backX, backY, squareSize);
    		graphics.strokePath(pen);

    		/* 手前の四角の内側を背景色で塗り、奥の線を消してから輪郭を描く / Fill the front square with the background to erase the back lines, then outline it */
    		if (backgroundColor) {
    			squarePath(graphics, frontX, frontY, squareSize);
    			graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
    		}
    		squarePath(graphics, frontX, frontY, squareSize);
    		graphics.strokePath(pen);
    	}

    	/**
    	 * 反転アイコン（軸の点線＋向かい合う三角形）を描画する
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {string} iconType - "flipHorizontal" または "flipVertical"
    	 * @param {number} width - ボタンの幅
    	 * @param {number} height - ボタンの高さ
    	 * @returns {void}
    	 */
    	function drawFlipIcon(graphics, iconType, width, height) {
    		var color = iconColor;
    		var centerX = width / 2;
    		var centerY = height / 2;

    		if (iconType === "flipVertical") {
    			/* 横の点線を軸に、上向き／下向きの三角形を配置する / Up/down triangles about a horizontal dotted axis */
    			drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
    			drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
    			drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
    		} else {
    			/* 縦の点線を軸に、左向き／右向きの三角形を配置する / Left/right triangles about a vertical dotted axis */
    			drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
    			drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
    			drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
    		}
    	}

    	/**
    	 * 回転アイコン（実線弧＋点線弧＋矢じり）を描画する
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {number} width - ボタンの幅
    	 * @param {number} height - ボタンの高さ
    	 * @param {boolean} mirror - true で左右反転（時計回りの図柄）
    	 * @returns {void}
    	 */
    	function drawRotateIcon(graphics, width, height, mirror) {
    		var color = iconColor;
    		var centerX = width / 2;
    		var centerY = height / 2 + 1;
    		var radius = 7.5;
    		var mirrorSign = mirror ? -1 : 1;   /* 左右反転のときは x をミラーする / Mirror x when flipped */
    		var headDeg = 232;                  /* 矢じりの位置（左上）/ Arrowhead position (top-left) */

    		strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
    		/* 下側は四角い点線 / Square-dotted arc on the lower side */
    		drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);

    		/* 左上（反転時は右上）に大きめの矢じりを付ける / Add a larger arrowhead top-left (top-right when mirrored) */
    		var headRad = headDeg * Math.PI / 180;
    		var headX = centerX + radius * Math.cos(headRad);
    		var headY = centerY + radius * Math.sin(headRad);
    		var tangentX = Math.sin(headRad);   /* 反時計回り（角度が減る向き）の接線 / Tangent for the CCW (decreasing angle) direction */
    		var tangentY = -Math.cos(headRad);
    		var perpX = -tangentY;
    		var perpY = tangentX;
    		var tipForward = 4;       /* 矢じり先端の前方への張り出し / Arrowhead tip extent (forward) */
    		var tipBack = 2;          /* 矢じり後方への張り出し / Arrowhead extent (backward) */
    		var tipHalfWidth = 4.5;   /* 矢じりの片側の幅 / Arrowhead half width */

    		var arrowPoints = [
    			[headX + tangentX * tipForward, headY + tangentY * tipForward],
    			[headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
    			[headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
    		];

    		if (mirror) {
    			for (var i = 0; i < arrowPoints.length; i++) {
    				arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
    			}
    		}

    		drawTriangle(graphics, arrowPoints, color, true);
    	}

    	/**
    	 * ボタンの下地（背景＋枠線）を描く。描画に失敗したら OS 標準の見た目にフォールバックする
    	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
    	 * @param {number} width - ボタンの幅
    	 * @param {number} height - ボタンの高さ
    	 * @param {number[]} backgroundColor - 背景色の RGBA
    	 * @returns {void}
    	 */
    	function drawButtonBase(graphics, width, height, backgroundColor) {
    		try {
    			graphics.rectPath(0, 0, width, height);
    			graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
    			if (iconBorderColor) {
    				graphics.rectPath(0.5, 0.5, width - 1, height - 1);
    				graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconBorderColor, 1));
    			}
    		} catch (e) {
    			try { graphics.drawOSControl(); } catch (osControlError) {}
    		}
    	}

    	/**
    	 * ホバー状態に応じた背景色を返す
    	 * @param {Button} control - 対象のコントロール
    	 * @returns {number[]} 背景色の RGBA
    	 */
    	function hoverBackground(control) {
    		return (control.isHover === true) ? iconHoverBg : iconBaseBg;
    	}

    	/**
    	 * 反転・回転アイコンボタンを描画する
    	 * @param {Button} button - 対象のボタン（iconType を持つ）
    	 * @returns {void}
    	 */
    	function drawIconButton(button) {
    		var graphics = button.graphics;
    		var width = button.size[0];
    		var height = button.size[1];
    		drawButtonBase(graphics, width, height, hoverBackground(button));

    		if (button.iconType === "rotate") {
    			drawRotateIcon(graphics, width, height, false);
    		} else if (button.iconType === "rotateFlip") {
    			drawRotateIcon(graphics, width, height, true);
    		} else {
    			drawFlipIcon(graphics, button.iconType, width, height);
    		}
    	}

    	/**
    	 * 方向ボタン（矢印）を描画する
    	 * @param {Button} button - 対象のボタン（directionKey を持つ）
    	 * @returns {void}
    	 */
    	function drawDirectionButton(button) {
    		var graphics = button.graphics;
    		var width = button.size[0];
    		var height = button.size[1];
    		drawButtonBase(graphics, width, height, hoverBackground(button));
    		drawArrow(graphics, button.directionKey, width, height, iconColor);
    	}

    	/**
    	 * 中央の複製ボタンを描画する
    	 * @param {Button} button - 対象のボタン
    	 * @returns {void}
    	 */
    	function drawCenterButton(button) {
    		var graphics = button.graphics;
    		var width = button.size[0];
    		var height = button.size[1];
    		var backgroundColor = hoverBackground(button);
    		drawButtonBase(graphics, width, height, backgroundColor);
    		drawDuplicateGlyph(graphics, width, height, iconColor, backgroundColor);
    	}

    	// =========================================
    	// 実行 / Actions
    	// =========================================
    	/**
    	 * 再入防止つきで処理を実行する（連打・スライダードラッグ中の多重実行を防ぐ）
    	 * @param {function} action - 実行する処理
    	 * @returns {void}
    	 */
    	function runExclusive(action) {
    		if (isBusy) { return; }
    		isBusy = true;
    		try {
    			action();
    		} finally {
    			isBusy = false;
    		}
    	}

    	/**
    	 * Option（Alt）キーが押されているかを判定する
    	 * @returns {boolean} 押されていれば true
    	 */
    	function isAltPressed() {
    		try {
    			return ScriptUI.environment.keyboardState.altKey === true;
    		} catch (e) {
    			return false;
    		}
    	}

    	/**
    	 * オプションパネルからマージン／プレビュー境界を読む（パネル未生成のときは既定値）
    	 * @returns {{marginPt: number, usePreviewBounds: boolean}} 変形に使う設定
    	 */
    	function getTransformOptions() {
    		var settings = readTransformOptions ? readTransformOptions() : null;
    		return {
    			marginPt: (settings && settings.marginPt) || 0,
    			usePreviewBounds: !!(settings && settings.usePreviewBounds)
    		};
    	}

    	/**
    	 * 反転・回転アイコンに対応する変形を実行する（回転は1クリック＝ROTATE_ANGLE）
    	 * @param {string} actionName - "FLIP_HORIZONTAL" / "FLIP_VERTICAL" / "ROTATE" / "ROTATE_FLIP"
    	 * @param {boolean} duplicate - true で複製してから変形する
    	 * @returns {void}
    	 */
    	function runIconAction(actionName, duplicate) {
    		runExclusive(function () {
    			var settings = getTransformOptions();
    			if (actionName === "FLIP_HORIZONTAL") {
    				btFlipSelection(-100, 100, duplicate, settings.marginPt, settings.usePreviewBounds);
    			} else if (actionName === "FLIP_VERTICAL") {
    				btFlipSelection(100, -100, duplicate, settings.marginPt, settings.usePreviewBounds);
    			} else if (actionName === "ROTATE") {
    				btRotateSelection(ROTATE_ANGLE, duplicate, settings.marginPt, settings.usePreviewBounds);    /* 反時計回り / counterclockwise */
    			} else if (actionName === "ROTATE_FLIP") {
    				btRotateSelection(-ROTATE_ANGLE, duplicate, settings.marginPt, settings.usePreviewBounds);   /* 時計回り / clockwise */
    			}
    		});
    	}

    	/**
    	 * 指定方向へ移動・複製を即時実行する
    	 * @param {string} directionKey - "up" / "left" / "right" / "down"
    	 * @returns {void}
    	 */
    	function runDirectionAction(directionKey) {
    		runExclusive(function () {
    			var settings = getTransformOptions();
    			delegateApply({
    				direction: directionKey,
    				marginPt: settings.marginPt,
    				usePreviewBounds: settings.usePreviewBounds,
    				duplicate: isAltPressed()
    			});
    		});
    	}

    	/**
    	 * 選択オブジェクトを同じ座標に複製し、複製側を選択状態にする（方向 none＝オフセット0）
    	 * @returns {void}
    	 */
    	function runDuplicateInPlace() {
    		runExclusive(function () {
    			delegateApply({ direction: 'none', marginPt: 0, usePreviewBounds: false, duplicate: true });
    		});
    	}

    	// =========================================
    	// UI 部品 / UI helpers
    	// =========================================
    	/**
    	 * コントロールのサイズを固定する（最小・推奨・最大を同じ値でそろえる）
    	 * @param {Object} control - 対象のコントロール
    	 * @param {number} width - 幅
    	 * @param {number} height - 高さ
    	 * @returns {void}
    	 */
    	function fixControlSize(control, width, height) {
    		control.minimumSize = [width, height];
    		control.preferredSize = [width, height];
    		control.maximumSize = [width, height];
    	}

    	/**
    	 * コントロールを再描画する（notify は環境により例外を投げ得るので保護）
    	 * @param {Object} control - 対象のコントロール
    	 * @returns {void}
    	 */
    	function redrawControl(control) {
    		try { control.notify("onDraw"); } catch (e) {}
    	}

    	/**
    	 * マウスオーバーの状態を button.isHover に反映して再描画する（方向・反転回転アイコン共通）
    	 * @param {Button} button - 対象のボタン
    	 * @returns {void}
    	 */
    	function attachHover(button) {
    		try {
    			button.addEventListener("mouseover", function () { button.isHover = true; redrawControl(button); });
    			button.addEventListener("mouseout", function () { button.isHover = false; redrawControl(button); });
    		} catch (e) {}
    	}

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

    	// ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    	// -----------------------------------------
    	// ステップボタンの寸法・増減量 / Stepper metrics and steps
    	// -----------------------------------------
    	var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    	var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    	var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    	var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    	var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    	var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    	var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    	// -----------------------------------------
    	// ステップボタンの配色 / Stepper colors
    	// -----------------------------------------
    	var STEPPER_UI_DARK           = isDarkUI();
    	/* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
    	   ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
    	   UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
    	   dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    	var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    	var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    	var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    	var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    	var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    	var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    	var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    	// -----------------------------------------
    	// 数値欄を作る（外から呼ぶ関数） / Public API
    	// -----------------------------------------
    	/**
    	 * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
    	 * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
    	 * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
    	 * @param {Group|Panel} parent - 追加先
    	 * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
    	 *     step / min / max / integer（true で整数のみ）/ unit / onStep
    	 * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
    	 */
    	function addSteppedField(parent, fieldOptions) {
    	    var fieldRowGroup = parent.add("group");
    	    fieldRowGroup.orientation = "row";
    	    fieldRowGroup.alignChildren = ["left", "center"];
    	    fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

    	    var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
    	    if (fieldOptions.labelWidth) {
    	        fieldLabel.preferredSize.width = fieldOptions.labelWidth;
    	        fieldLabel.justify = "right";
    	    }

    	    /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
    	    var stepperInputGroup = fieldRowGroup.add("group");
    	    stepperInputGroup.orientation = "row";
    	    stepperInputGroup.alignChildren = ["left", "center"];
    	    stepperInputGroup.spacing = 0;
    	    stepperInputGroup.margins = 0;

    	    var numberInput;
    	    var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
    	    numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
    	    numberInput.characters = fieldOptions.characters || 6;
    	    numberInput.fieldLabel = fieldLabel;
    	    numberInput.stepperGroup = stepperGroup;

    	    /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
    	    bindSteppedArrowKeys(numberInput, stepperGroup);

    	    /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
    	    fieldLabel.addEventListener("click", function () {
    	        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
    	        numberInput.active = true;
    	    });

    	    /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
    	    numberInput.lastValidText = numberInput.text;
    	    numberInput.onChange = function () {
    	        var value = parseFloat(numberInput.text);
    	        if (isNaN(value)) {
    	            numberInput.text = numberInput.lastValidText;
    	            return;
    	        }
    	        writeSteppedValue(numberInput, value, fieldOptions);
    	    };
    	    return numberInput;
    	}

    	/**
    	 * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
    	 * @param {EditText} numberInput - addSteppedField() で作った入力欄
    	 * @param {boolean} isEnabled - 有効にするなら true
    	 * @returns {void}
    	 */
    	function setSteppedFieldEnabled(numberInput, isEnabled) {
    	    numberInput.enabled = isEnabled;
    	    numberInput.fieldLabel.enabled = isEnabled;
    	    numberInput.stepperGroup.enabled = isEnabled;
    	    /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
    	    for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
    	        redrawStepperGroup(numberInput.stepperGroup.children[i]);
    	    }
    	}

    	/**
    	 * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
    	 * @param {Group|Panel} parent - 追加先
    	 * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
    	 * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
    	 * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
    	 */
    	function addStepper(parent, getNumberInput, stepOptions) {
    	    var stepperGroup = parent.add("group");
    	    stepperGroup.orientation = "column";
    	    stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
    	    stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
    	    stepperGroup.alignment = ["left", "center"];

    	    /**
    	     * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
    	     * @param {number} direction - 増やすなら 1、減らすなら -1
    	     * @returns {void}
    	     */
    	    function stepBy(direction) {
    	        var numberInput = getNumberInput();
    	        if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
    	        var value = parseFloat(numberInput.text);
    	        if (isNaN(value)) value = 0;
    	        writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
    	        if (stepOptions.onStep) stepOptions.onStep(numberInput);
    	    }

    	    /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
    	    var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
    	    var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
    	    makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
    	    makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
    	    stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
    	    return stepperGroup;
    	}

    	/**
    	 * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
    	 * @param {EditText} numberInput - 対象の入力欄
    	 * @param {Group} stepperGroup - addStepper() で作った∧∨
    	 * @returns {void}
    	 */
    	function bindSteppedArrowKeys(numberInput, stepperGroup) {
    	    numberInput.addEventListener("keydown", function (event) {
    	        if (event.keyName !== "Up" && event.keyName !== "Down") return;
    	        stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
    	        event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
    	    });
    	}

    	// -----------------------------------------
    	// 値の計算 / Value helpers
    	// -----------------------------------------
    	/**
    	 * 押された修飾キーに応じて、1回分増減した値を返す
    	 * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
    	 * 整数の欄では option を無視して step の倍数へ）
    	 * @param {number} value - 元の値
    	 * @param {number} direction - 増やすなら 1、減らすなら -1
    	 * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
    	 * @returns {number} 増減した値（下限・上限は未適用）
    	 */
    	function computeSteppedValue(value, direction, stepOptions) {
    	    var keyState = ScriptUI.environment.keyboardState;
    	    if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
    	    if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
    	    return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    	}

    	/**
    	 * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
    	 * @param {number} value - 元の値
    	 * @param {number} multiple - 倍数の単位（例 10）
    	 * @param {number} direction - 上げるなら 1、下げるなら -1
    	 * @returns {number} 移した値
    	 */
    	function snapStepperToNextMultiple(value, multiple, direction) {
    	    /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
    	       round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
    	    var quotient = Math.round(value / multiple * 1e6) / 1e6;
    	    if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
    	    return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    	}

    	/**
    	 * 値を下限・上限の範囲に収める
    	 * @param {number} value - 数値
    	 * @param {Object} rangeOptions - min / max（どちらも省略可）
    	 * @returns {number} 範囲に収めた値
    	 */
    	function clampSteppedValue(value, rangeOptions) {
    	    if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
    	    if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
    	    return value;
    	}

    	/**
    	 * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
    	 * @param {EditText} numberInput - 書き込む入力欄
    	 * @param {number} value - 数値
    	 * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
    	 * @returns {void}
    	 */
    	function writeSteppedValue(numberInput, value, valueOptions) {
    	    numberInput.text = formatSteppedValue(value, valueOptions);
    	    numberInput.lastValidText = numberInput.text;
    	}

    	/**
    	 * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
    	 * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
    	 * @param {number} value - 数値
    	 * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
    	 * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
    	 */
    	function formatSteppedValue(value, valueOptions) {
    	    if (valueOptions.integer) value = Math.round(value);
    	    return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    	}

    	/**
    	 * 小数第2位で丸めた数値を文字列で返す
    	 * @param {number} value - 数値
    	 * @returns {string} 表示用の数値文字列
    	 */
    	function formatStepperNumber(value) {
    	    return String(Math.round(value * 100) / 100);
    	}

    	// -----------------------------------------
    	// ∧∨ボタンの描画 / Drawing
    	// -----------------------------------------
    	/**
    	 * 山形（∧／∨）の極小ボタンを作成する。
    	 * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
    	 * 継ぎ目に線は引かない
    	 * @param {Group|Panel} parent - 追加先
    	 * @param {string} direction - "up" または "down"
    	 * @param {Function} onClickFn - クリック時の処理
    	 * @returns {Group} ボタンとして使う group
    	 */
    	function makeStepperChevronButton(parent, direction, onClickFn) {
    	    var buttonWidth = STEPPER_BUTTON_WIDTH;
    	    var buttonHeight = STEPPER_BUTTON_HEIGHT;
    	    var isUp = (direction === "up");
    	    var chevronBox = parent.add("group");
    	    chevronBox.margins = 0;
    	    chevronBox.spacing = 0;
    	    chevronBox.preferredSize = [buttonWidth, buttonHeight];
    	    chevronBox.minimumSize = [buttonWidth, buttonHeight];
    	    chevronBox.maximumSize = [buttonWidth, buttonHeight];
    	    chevronBox.isPressed = false;
    	    chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

    	    chevronBox.onDraw = function () {
    	        var boxGraphics = chevronBox.graphics;
    	        /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
    	           Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
    	        var isDimmed = !isStepperEnabledInTree(chevronBox);

    	        /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
    	        var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
    	        boxGraphics.newPath();
    	        boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
    	        boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

    	        drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
    	        drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
    	    };

    	    /**
    	     * 押下状態を変えて描き直す
    	     * @param {boolean} isPressed - 押下中なら true
    	     * @returns {void}
    	     */
    	    function repaint(isPressed) {
    	        if (chevronBox.isPressed === isPressed) return;
    	        chevronBox.isPressed = isPressed;
    	        redrawStepperGroup(chevronBox);
    	    }
    	    chevronBox.addEventListener("mousedown", function () {
    	        if (!isStepperEnabledInTree(chevronBox)) return;
    	        repaint(true);
    	        if (onClickFn) onClickFn();
    	    });
    	    chevronBox.addEventListener("mouseup", function () { repaint(false); });
    	    /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
    	    chevronBox.addEventListener("mouseout", function () { repaint(false); });
    	    return chevronBox;
    	}

    	/**
    	 * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
    	 * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
    	 * @param {ScriptUIGraphics} boxGraphics - 描画先
    	 * @param {number} boxWidth - ボタンの幅
    	 * @param {number} boxHeight - ボタンの高さ
    	 * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
    	 * @param {number[]} frameColor - [r, g, b, a]
    	 * @returns {void}
    	 */
    	function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
    	    var frameLeft = 0.5;
    	    var frameRight = boxWidth - 0.5;
    	    var outerY = isUp ? 0.5 : boxHeight - 0.5;
    	    var seamY = isUp ? boxHeight : 0;
    	    var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
    	    var radius = STEPPER_CORNER_RADIUS;
    	    var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
    	    var angle, k;

    	    boxGraphics.newPath();
    	    boxGraphics.moveTo(frameLeft, seamY);
    	    /* 左の角丸 / left corner */
    	    for (k = 0; k <= arcSteps; k++) {
    	        angle = (Math.PI / 2) * k / arcSteps;
    	        boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
    	    }
    	    /* 右の角丸 / right corner */
    	    for (k = 0; k <= arcSteps; k++) {
    	        angle = (Math.PI / 2) * k / arcSteps;
    	        boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
    	    }
    	    boxGraphics.lineTo(frameRight, seamY);
    	    boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    	}

    	/**
    	 * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
    	 * @param {ScriptUIGraphics} boxGraphics - 描画先
    	 * @param {number} boxWidth - ボタンの幅
    	 * @param {number} boxHeight - ボタンの高さ
    	 * @param {boolean} isUp - ∧なら true、∨なら false
    	 * @param {number[]} chevronColor - [r, g, b, a]
    	 * @returns {void}
    	 */
    	function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
    	    var centerX = boxWidth / 2;
    	    var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
    	    var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
    	    var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
    	    boxGraphics.newPath();
    	    boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
    	    boxGraphics.lineTo(centerX, centerY + tipOffsetY);
    	    boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
    	    boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    	}

    	/**
    	 * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
    	 * @param {Object} control - 対象のコントロール
    	 * @returns {boolean} すべて有効なら true
    	 */
    	function isStepperEnabledInTree(control) {
    	    for (var node = control; node; node = node.parent) {
    	        if (!node.enabled) return false;
    	    }
    	    return true;
    	}

    	/**
    	 * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
    	 * @param {Object} container - 行・グループ・パネルなど
    	 * @returns {void}
    	 */
    	function redrawSteppersIn(container) {
    	    if (!container.children) return;
    	    for (var i = 0; i < container.children.length; i++) {
    	        var child = container.children[i];
    	        if (child.isStepperButton) redrawStepperGroup(child);
    	        else redrawSteppersIn(child);
    	    }
    	}

    	/**
    	 * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
    	 * @param {Group} targetGroup - 描き直す group
    	 * @returns {void}
    	 */
    	function redrawStepperGroup(targetGroup) {
    	    targetGroup.hide();
    	    targetGroup.show();
    	}

    	// ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    	// =========================================
    	// パネル構築 / Panel builders
    	// =========================================
    	/* 反転・回転の4アイコン（左右反転／上下反転／回転CCW／回転CW）/ The four flip/rotate icons */
    	var ICON_BUTTON_DEFS = [
    		{ name: "FLIP_HORIZONTAL", icon: "flipHorizontal", tooltip: "tooltip.flipHorizontal" },
    		{ name: "FLIP_VERTICAL",   icon: "flipVertical",   tooltip: "tooltip.flipVertical" },
    		{ name: "ROTATE",          icon: "rotate",         tooltip: "tooltip.rotateCCW" },
    		{ name: "ROTATE_FLIP",     icon: "rotateFlip",     tooltip: "tooltip.rotateCW" }
    	];

    	/**
    	 * 反転・回転アイコンボタンを1つ生成する
    	 * @param {Group} parentGroup - 追加先のグループ
    	 * @param {{name: string, icon: string, tooltip: string}} buttonDef - ボタン定義
    	 * @returns {void}
    	 */
    	function addIconButton(parentGroup, buttonDef) {
    		var button = parentGroup.add("button", undefined, "");
    		button.helpTip = getLabel(buttonDef.tooltip);
    		/* 移動・複製の方向ボタンと同じ大きさに合わせる / Match the size of the move/duplicate direction buttons */
    		fixControlSize(button, ICON_SIZE, ICON_SIZE);
    		button.iconType = buttonDef.icon;
    		button.isHover = false;
    		button.onDraw = function () { drawIconButton(this); };
    		/* Option＝複製してから変形 / Option = duplicate before transforming */
    		button.onClick = function () { runIconAction(buttonDef.name, isAltPressed()); };
    		attachHover(button);
    	}

    	/**
    	 * 方向ボタンを追加する（クリックでその方向へ移動、Option＋クリックで複製）
    	 * @param {Group} parentRow - 追加先の行グループ
    	 * @param {string} directionKey - "up" / "left" / "right" / "down"
    	 * @returns {void}
    	 */
    	function addDirectionButton(parentRow, directionKey) {
    		var button = parentRow.add('iconbutton', undefined, undefined, { style: 'toolbutton' });
    		fixControlSize(button, ICON_SIZE, ICON_SIZE);
    		button.directionKey = directionKey;
    		button.isHover = false;
    		button.helpTip = getLabel('direction.' + directionKey) + '  —  ' + getLabel('tooltip.moveDuplicate');
    		button.onDraw = function () { drawDirectionButton(this); };
    		button.onClick = function () { runDirectionAction(this.directionKey); };
    		attachHover(button);
    	}

    	/**
    	 * 中央の複製ボタンを追加する
    	 * @param {Group} parentRow - 追加先の行グループ
    	 * @returns {void}
    	 */
    	function addCenterButton(parentRow) {
    		var button = parentRow.add('iconbutton', undefined, undefined, { style: 'toolbutton' });
    		fixControlSize(button, ICON_SIZE, ICON_SIZE);
    		button.isHover = false;
    		button.helpTip = getLabel('tooltip.duplicateInPlace');
    		button.onDraw = function () { drawCenterButton(this); };
    		button.onClick = function () { runDuplicateInPlace(); };
    		attachHover(button);
    	}

    	/**
    	 * 十字レイアウトの空セルを追加する
    	 * @param {Group} parentRow - 追加先の行グループ
    	 * @returns {void}
    	 */
    	function addSpacerCell(parentRow) {
    		fixControlSize(parentRow.add('statictext', undefined, ''), ICON_SIZE, ICON_SIZE);
    	}

    	/**
    	 * 回転スライダー（-180〜180°・15°刻み／Shift＝90°刻み）をパネル最下部に追加する
    	 * 基点は9軸ウィジェットに従い、正＝反時計回り
    	 * @param {Panel} parentPanel - 追加先のパネル
    	 * @returns {void}
    	 */
    	function addRotateSlider(parentPanel) {
    		var sliderRow = parentPanel.add('group');
    		setupRow(sliderRow, 'fill', SLIDER_ROW_SPACING);

    		/* スライダーの左に現在角度を数字で表示（スナップ後の値）/ Numeric angle readout to the left of the slider (snapped value) */
    		var angleLabel = sliderRow.add('statictext', undefined, '0°', { justify: 'right' });
    		angleLabel.preferredSize = [ANGLE_LABEL_WIDTH, SLIDER_ROW_HEIGHT];

    		var rotateSlider = sliderRow.add('slider', undefined, 0, -180, 180);
    		rotateSlider.helpTip = getLabel('tooltip.rotateSlider');
    		rotateSlider.alignment = ['fill', 'center'];
    		rotateSlider.preferredSize = [SLIDER_WIDTH, SLIDER_ROW_HEIGHT];

    		/* 前回適用したスナップ角（差分回転の基準。ドラッグ開始時は 0）/ Last applied snapped angle (baseline for delta rotation; 0 at drag start) */
    		var previousSnapped = 0;

    		/**
    		 * スライダー値を刻み幅（通常15°、Shift 押下中は90°）にスナップする
    		 * @param {number} value - スライダーの生の値
    		 * @returns {number} -180〜180 に収めたスナップ角
    		 */
    		function snapSliderAngle(value) {
    			var step = 15;
    			try {
    				if (ScriptUI.environment.keyboardState.shiftKey === true) { step = 90; }
    			} catch (e) {}
    			var snapped = Math.round(value / step) * step;
    			if (snapped < -180) { snapped = -180; }
    			if (snapped > 180) { snapped = 180; }
    			return snapped;
    		}

    		/**
    		 * スナップ角まで前回位置との差分だけ回転する（正＝反時計回り）
    		 * isBusy 中はスキップし、差分は次のティックで取り戻す（previousSnapped は適用時のみ進める）
    		 * @param {number} snapped - スナップ後の角度
    		 * @returns {void}
    		 */
    		function applySliderRotation(snapped) {
    			if (snapped === previousSnapped) { return; }
    			runExclusive(function () {
    				var settings = getTransformOptions();
    				btRotateSelection(snapped - previousSnapped, false, settings.marginPt, settings.usePreviewBounds);
    				previousSnapped = snapped;
    			});
    		}

    		/* ドラッグ中：刻み境界を越えるたびに差分回転し、角度表示も更新 / While dragging: rotate by the delta at each step boundary and refresh the readout */
    		rotateSlider.onChanging = function () {
    			var snapped = snapSliderAngle(this.value);
    			angleLabel.text = snapped + '°';
    			applySliderRotation(snapped);
    		};

    		/* 離した時：最終スナップ角まで回してからスライダーと表示を 0 に戻す（オブジェクトは回った位置のまま）/ On release: finish rotating, then reset the slider and readout to 0 (the object keeps its rotation) */
    		rotateSlider.onChange = function () {
    			applySliderRotation(snapSliderAngle(this.value));
    			previousSnapped = 0;
    			this.value = 0;
    			angleLabel.text = '0°';
    		};
    	}

    	/**
    	 * 反転・回転パネル（アイコン2×2＋9軸ウィジェット＋回転スライダー）を追加する
    	 * @param {Window} targetWindow - 追加先のウィンドウ
    	 * @returns {void}
    	 */
    	function buildFlipPanel(targetWindow) {
    		var flipPanel = targetWindow.add('panel', undefined, getLabel('panel.flip'));
    		setupPanel(flipPanel, 6);

    		/* アイコン2×2（左）と9軸ウィジェット（右）を横並び。パネルは fill なので中央寄せは行側で指定 / 2x2 icons (left) and the 9-axis widget (right); the panel fills, so center on the row */
    		var flipRow = flipPanel.add('group');
    		setupRow(flipRow, 'center', GROUP_SPACING);

    		/* アイコンボタンを ICON_COLUMNS_PER_ROW 個ごとに改行して並べる / Lay out icon buttons, wrapping every ICON_COLUMNS_PER_ROW */
    		var iconGrid = flipRow.add('group');
    		iconGrid.orientation = 'column';
    		iconGrid.alignChildren = ['center', 'center'];
    		iconGrid.spacing = ICON_GAP;

    		var iconRow = null;
    		for (var iconIndex = 0; iconIndex < ICON_BUTTON_DEFS.length; iconIndex++) {
    			if ((iconIndex % ICON_COLUMNS_PER_ROW) === 0) {
    				iconRow = iconGrid.add('group');
    				setupRow(iconRow, 'center', ICON_GAP);
    			}
    			addIconButton(iconRow, ICON_BUTTON_DEFS[iconIndex]);
    		}

    		/* 9軸（3×3）の基準点ウィジェット（アイコンの右。反転・回転の基点を指定）/ 9-axis anchor widget (right of the icons; sets the flip/rotate pivot) */
    		pivotAnchorWidget = addAnchorWidget(flipRow, 4);
    		pivotAnchorWidget.helpTip = getLabel('tooltip.anchor');

    		addRotateSlider(flipPanel);
    	}

    	/**
    	 * 移動・複製パネル（方向の十字ボタン。クリックで移動／Option＋クリックで複製）を追加する
    	 * @param {Window} targetWindow - 追加先のウィンドウ
    	 * @returns {void}
    	 */
    	function buildMovePanel(targetWindow) {
    		var directionPanel = targetWindow.add('panel', undefined, getLabel('panel.direction'));
    		setupPanel(directionPanel, 6);

    		/* 十字ボタンは専用サブグループへ（行間を密に保つ）。パネルは fill なので中央寄せはグループ側で指定 / Keep the cross in its own subgroup (tight rows); the panel fills, so center on the group */
    		var crossGroup = directionPanel.add('group');
    		crossGroup.orientation = 'column';
    		crossGroup.alignChildren = 'center';
    		crossGroup.alignment = ['center', 'top'];
    		crossGroup.spacing = CROSS_GAP;

    		var topRow = crossGroup.add('group');
    		setupRow(topRow, 'center', CROSS_GAP);
    		addSpacerCell(topRow);
    		addDirectionButton(topRow, 'up');
    		addSpacerCell(topRow);

    		var middleRow = crossGroup.add('group');
    		setupRow(middleRow, 'center', CROSS_GAP);
    		addDirectionButton(middleRow, 'left');
    		addCenterButton(middleRow);
    		addDirectionButton(middleRow, 'right');

    		var bottomRow = crossGroup.add('group');
    		setupRow(bottomRow, 'center', CROSS_GAP);
    		addSpacerCell(bottomRow);
    		addDirectionButton(bottomRow, 'down');
    		addSpacerCell(bottomRow);
    	}

    	/**
    	 * オプションパネル（マージン／プレビュー境界）を追加し、設定読取り関数を readTransformOptions に公開する
    	 * @param {Window} targetWindow - 追加先のウィンドウ
    	 * @returns {void}
    	 */
    	function buildOptionsPanel(targetWindow) {
    		var optionsPanel = targetWindow.add('panel', undefined, getLabel('panel.options'));
    		setupPanel(optionsPanel, 6);

    		var marginGroup = optionsPanel.add('group');
    		setupRow(marginGroup, 'left', LABEL_FIELD_SPACING);
    		marginGroup.add('statictext', undefined, labelText('fieldLabel.margin'));
    		/* ∧∨と入力欄は隙間0で突き合わせる。負値を許可（マイナスで重なり方向へ）なので下限は付けない
    		   Butt the stepper against the field; no minimum, since negatives move toward overlap */
    		var marginStepperGroup = marginGroup.add('group');
    		setupRow(marginStepperGroup, 'left', 0);
    		marginStepperGroup.margins = 0;
    		var marginInput;
    		var marginStepper = addStepper(marginStepperGroup, function () { return marginInput; }, {});
    		marginInput = marginStepperGroup.add('edittext', undefined, String(DEFAULT_MARGIN));
    		marginInput.characters = FIELD_CHARS;
    		bindSteppedArrowKeys(marginInput, marginStepper); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */
    		var unitLabel = marginGroup.add('statictext', undefined, getUnitInfo().label);

    		var previewBoundsCheck = optionsPanel.add('checkbox', undefined, getLabel('checkbox.previewBounds'));
    		previewBoundsCheck.alignment = 'left'; /* パネルは fill なのでチェックボックスだけ左寄せに戻す / The panel fills, so pull the checkbox back to the left */
    		previewBoundsCheck.value = DEFAULT_PREVIEW_BOUNDS;

    		/**
    		 * UI からマージン／プレビュー境界を読む（不正値はここで 0 に落とし、単位表示も更新する）
    		 * @returns {{marginPt: number, usePreviewBounds: boolean}} 変形に使う設定
    		 */
    		function readSettings() {
    			var unitInfo = getUnitInfo();
    			unitLabel.text = unitInfo.label; /* 単位表示を更新 / refresh the unit label */
    			var marginValue = parseFloat(marginInput.text);
    			if (isNaN(marginValue)) { marginValue = 0; } /* 負値は許容（マイナスで重なり方向へ）/ Negatives allowed (moves toward overlap) */
    			marginInput.text = String(marginValue);
    			return {
    				marginPt: marginValue * unitInfo.pointsPerUnit,
    				usePreviewBounds: previewBoundsCheck.value
    			};
    		}
    		/* 移動・複製と反転・回転の双方から同じ設定を使えるよう公開 / Publish so both move-duplicate and flip-rotate reuse the same settings */
    		readTransformOptions = readSettings;
    	}

    	// =========================================
    	// パレット / Palette
    	// =========================================
    	/**
    	 * 常駐パレットを組み立てて表示する（重複起動の判定は IIFE 冒頭で済ませている）
    	 * @returns {void}
    	 */
    	function showPalette() {
    		/* UI の明暗からアイコンの配色を決定 / Decide icon colors from the light/dark UI */
    		initIconColors();

    		var win = new Window("palette", getLabel('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
    		setupWindow(win);

    		/* 1カラム：反転・回転 → 移動・複製 → オプションの順で縦積み / Single column: Flip-Rotate → Move-Duplicate → Options */
    		buildFlipPanel(win);
    		buildMovePanel(win);
    		buildOptionsPanel(win);

    		/* Esc で閉じる / Esc closes */
    		win.addEventListener('keydown', function (event) {
    			if (event.keyName === 'Escape') { win.close(); }
    		});
    		/* 閉じるとき：参照を解放（次回起動で作り直せるように）/ On close: release the reference so the next launch rebuilds */
    		win.onClose = function () {
    			$.global.__quickTransformPalette = null;
    			return true;
    		};

    		/* 常駐参照：GC 回避と二重起動の検出を兼ねる / Persistent reference: avoids GC and detects a second launch */
    		$.global.__quickTransformPalette = win;
    		win.layout.layout(true);
    		win.show();
    	}

    	showPalette();

    }());

})();
