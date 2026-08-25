#target illustrator
#targetengine "QuickTransformPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの移動・複製と反転・回転を、アイコンのクリックで即時実行する常駐パレットです。
9軸の基準点・マージン・プレビュー境界の設定はすべての操作に共通で、Option＋クリックすると複製してから変形します。

詳細は README を参照してください。

### Overview

A persistent palette that moves, duplicates, flips and rotates the selection immediately on an icon click.
The nine-point reference, margin and preview-bounds settings are shared by every operation, and Option-clicking duplicates before transforming.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "QuickTransformPalette";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/QuickTransformPalette.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/QuickTransformPalette.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n277bd0865986"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
	var WINDOW_MARGINS        = 15;               /* ウィンドウ外周の余白 / window margin */
	var WINDOW_SPACING        = 10;               /* ウィンドウ内の要素間隔 / window spacing */
	var PANEL_MARGINS         = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
	var PANEL_SPACING         = 8;                /* パネル内の要素間隔 / panel spacing */
	var ICON_SIZE             = 30;               /* 方向・反転・回転アイコン1個の大きさ（px、両パネル共通）/ size of each direction / flip / rotate icon (px, shared) */
	var ICON_LINE_WIDTH       = 1;                /* アイコン内の線画の線幅（ボタン枠・9軸グリッドと同じ1）/ stroke width of the line art inside each icon */
	var ICON_GAP              = 7;                /* 反転・回転アイコンどうしの間隔 / gap between flip/rotate icons */
	var ICON_COLUMNS_PER_ROW  = 2;                /* 反転・回転アイコンを何個ごとに改行するか / flip/rotate icons per row before wrapping */
	var CROSS_GAP             = 2;                /* 方向ボタン（十字）どうしの間隔 / gap between direction buttons */
	var ANCHOR_WIDGET_SIZE    = 66;               /* 9軸ウィジェット全体の大きさ / overall size of the 9-axis widget */
	var ANCHOR_CELL_SIZE      = 9;                /* 9軸の□1個のサイズ / size of one anchor square */
	var ANCHOR_CELL_GAP       = 7.5;              /* 9軸の□どうしの間隔 / gap between anchor squares */
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
	/**
	 * 現在の言語を判定する
	 * @returns {string} "ja" または "en"
	 */
	function getCurrentLanguage() {
		return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
	}
	var currentLanguage = getCurrentLanguage();

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
			rotateSlider: { ja: "スライダーで回転（-180〜180°・15°刻み／Shift＝90°刻み、正＝反時計回り、基準点が基点）", en: "Rotate with the slider (-180 to 180°, 15° steps / Shift = 90° steps, positive = CCW, about the anchor point)" }
		}
	};

	/**
	 * ラベルをドット区切りキーで取得する（現在の言語）
	 * @param {string} labelPath - "tooltip.anchor" のようなドット区切りキー
	 * @returns {string} ラベル文字列（見つからなければキーをそのまま返す）
	 */
	function getLabel(labelPath) {
		var keys = labelPath.split('.');
		var node = LABELS;
		for (var i = 0; i < keys.length; i++) {
			node = node[keys[i]];
			if (node === undefined || node === null) { return labelPath; }
		}
		return node[currentLanguage];
	}

	/**
	 * コロン付きのラベルを取得する（日本語は全角、英語は半角）
	 * @param {string} labelPath - ドット区切りキー
	 * @returns {string} コロンを付けたラベル文字列
	 */
	function getLabelWithColon(labelPath) {
		return getLabel(labelPath) + (currentLanguage === 'ja' ? '：' : ':');
	}

	// =========================================
	// 単位 / Units
	// =========================================
	/* 定規の単位ID（rulerType）順に、表示ラベルと pt 換算係数を並べる / Label and pt factor per ruler unit id (rulerType) */
	var RULER_UNITS = [
		{ label: "inch", factor: 72.0 },                /* 0 */
		{ label: "mm",   factor: 72.0 / 25.4 },         /* 1 */
		{ label: "pt",   factor: 1.0 },                 /* 2 */
		{ label: "pica", factor: 12.0 },                /* 3 */
		{ label: "cm",   factor: 72.0 / 2.54 },         /* 4 */
		{ label: "Q",    factor: 72.0 / 25.4 * 0.25 },  /* 5 */
		{ label: "px",   factor: 1.0 }                  /* 6 */
	];
	var FALLBACK_RULER_UNIT_INDEX = 2; /* 判別できないときは pt 扱い / Fall back to pt */

	/**
	 * 定規の単位からラベルと pt 換算係数を求める
	 * @returns {{label: string, factor: number}} 単位ラベルと pt 換算係数
	 */
	function getRulerUnitInfo() {
		var rulerType = app.preferences.getIntegerPreference("rulerType");
		return RULER_UNITS[rulerType] || RULER_UNITS[FALLBACK_RULER_UNIT_INDEX];
	}

	// =========================================
	// 状態 / State
	// =========================================
	/* パレット本体は $.global.__quickTransformPalette に持たせる（IIFE 内の var では GC される）/ The palette itself lives on $.global.__quickTransformPalette (an IIFE-local var would be garbage-collected) */
	var isBusy = false;              /* 委譲の再入防止 / Re-entrancy guard for delegation */
	var selectedAnchorIndex = 4;     /* 基準点（9軸）: 0..8 を行優先（0=左上, 4=中央, 8=右下）/ 9-axis anchor 0..8 row-major (0=top-left, 4=center, 8=bottom-right) */
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
		var col = selectedAnchorIndex % 3;
		var row = Math.floor(selectedAnchorIndex / 3);
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

	/* 9軸セルの枠線：薄いグレー（常時）/ Anchor-cell border: light gray (always) */
	var ANCHOR_LINE_COLOR = [0.6, 0.6, 0.6, 1];
	/* 9軸セルの選択時の塗り（通常時は塗らずパネル地色を見せる）。initIconColors() で UI 明暗に合わせて上書き / Fill for the selected anchor cell (unpainted otherwise); overwritten per light/dark UI in initIconColors() */
	var ANCHOR_SELECTED_FILL = [0.4, 0.4, 0.4, 1];

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
		return getUIBrightness() > 0.5;
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
		/* 選択セルの塗り：ライトは濃いグレー、ダークは明るいグレー（暗い地色でも□が見えるように）/ Selected-cell fill: dark gray in light UI, bright gray in dark UI */
		ANCHOR_SELECTED_FILL = lightUI ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
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

	/**
	 * 基準点セルの□を1つ描画する（通常は塗り透過・枠は薄いグレー／選択時は塗りあり・枠はそのまま）
	 * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
	 * @param {number} x - 左端
	 * @param {number} y - 上端
	 * @param {boolean} selected - 選択中なら true
	 * @returns {void}
	 */
	function drawAnchorCell(graphics, x, y, selected) {
		/* 枠を上に描くので塗りを先に行う / Fill first so the border draws on top */
		if (selected) {
			squarePath(graphics, x, y, ANCHOR_CELL_SIZE);
			graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, ANCHOR_SELECTED_FILL));
		}
		squarePath(graphics, x, y, ANCHOR_CELL_SIZE);
		graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, ANCHOR_LINE_COLOR, 1));
	}

	/* 中央(4)を除く外周の□どうしをつなぐケイ線の組み合わせ / Pairs of outer squares (center 4 excluded) joined by rules */
	var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

	/**
	 * 9軸ウィジェットを描画する（外周の□をケイ線でつなぐ・中央は独立）
	 * @param {Button} widget - 対象のウィジェット
	 * @returns {void}
	 */
	function drawAnchorWidget(widget) {
		var graphics = widget.graphics;
		var width = widget.size[0];
		var height = widget.size[1];

		/* 背景は塗らずコントロール地色（パネルと同色）で塗って透過に見せる / Paint the control's own background color so the widget looks transparent */
		try {
			graphics.rectPath(0, 0, width, height);
			graphics.fillPath(graphics.backgroundColor);
		} catch (e) {}

		var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
		var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
		var originX = Math.round((width - gridSize) / 2);
		var originY = Math.round((height - gridSize) / 2);

		/* 9セルの左上座標を先に求める / Precompute the top-left corner of all nine cells */
		var cellPositions = [];
		for (var index = 0; index < 9; index++) {
			cellPositions.push([originX + (index % 3) * cellStep, originY + Math.floor(index / 3) * cellStep]);
		}

		/* ケイ線も枠と同じ薄いグレーに揃える / Match the connecting rules to the light-gray cell borders */
		var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, ANCHOR_LINE_COLOR, 1);
		for (var i = 0; i < ANCHOR_CONNECTIONS.length; i++) {
			var cellA = cellPositions[ANCHOR_CONNECTIONS[i][0]];
			var cellB = cellPositions[ANCHOR_CONNECTIONS[i][1]];
			graphics.newPath();
			if (ANCHOR_CONNECTIONS[i][1] - ANCHOR_CONNECTIONS[i][0] === 1) {
				/* 横方向：右隣の□へ / Horizontal: to the square on the right */
				graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE, cellA[1] + ANCHOR_CELL_SIZE / 2);
				graphics.lineTo(cellB[0], cellB[1] + ANCHOR_CELL_SIZE / 2);
			} else {
				/* 縦方向：下の□へ / Vertical: to the square below */
				graphics.moveTo(cellA[0] + ANCHOR_CELL_SIZE / 2, cellA[1] + ANCHOR_CELL_SIZE);
				graphics.lineTo(cellB[0] + ANCHOR_CELL_SIZE / 2, cellB[1]);
			}
			graphics.strokePath(linePen);
		}

		for (var cellIndex = 0; cellIndex < cellPositions.length; cellIndex++) {
			drawAnchorCell(graphics, cellPositions[cellIndex][0], cellPositions[cellIndex][1], cellIndex === selectedAnchorIndex);
		}
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
	 * ウィンドウの共通設定を適用する
	 * @param {Window} targetWindow - 対象のウィンドウ
	 * @returns {void}
	 */
	function setupWindow(targetWindow) {
		targetWindow.orientation = "column";
		targetWindow.alignChildren = ["fill", "top"];
		targetWindow.margins = WINDOW_MARGINS;
		targetWindow.spacing = WINDOW_SPACING;
	}

	/**
	 * パネルの共通設定を適用する
	 * @param {Panel} targetPanel - 対象のパネル
	 * @param {number} [spacing] - パネル内の要素間隔
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
	 * 行グループの共通設定を適用する（alignment と alignChildren を必ず対で指定する）
	 * @param {Group} targetGroup - 対象のグループ
	 * @param {string} [alignment] - グループ自身の横方向の配置（既定は "left"）
	 * @param {number} [spacing] - グループ内の要素間隔
	 * @returns {void}
	 */
	function setupRow(targetGroup, alignment, spacing) {
		targetGroup.orientation = "row";
		targetGroup.alignment = [alignment || "left", "center"];
		targetGroup.alignChildren = ["left", "center"];
		targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

	/**
	 * 数値入力欄を ↑↓ キーで増減できるようにする（Shift＝±10・Option＝±0.1）
	 * @param {EditText} editText - 対象の入力欄
	 * @param {boolean} allowNegative - true で負値を許可する
	 * @returns {void}
	 */
	function changeValueByArrowKey(editText, allowNegative) {
		editText.addEventListener("keydown", function (event) {
			if (event.keyName !== "Up" && event.keyName !== "Down") { return; }
			var value = Number(editText.text);
			if (isNaN(value)) { return; }

			var keyboard = ScriptUI.environment.keyboardState;
			var useDecimal = keyboard.altKey === true;
			var isUp = (event.keyName === "Up");

			if (keyboard.shiftKey) {
				/* Shift 押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
				value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
			} else if (useDecimal) {
				/* Option 押下時は0.1単位で増減 / Step by 0.1 when Option is held */
				value += isUp ? 0.1 : -0.1;
			} else {
				value += isUp ? 1 : -1;
			}
			if (!allowNegative && value < 0) { value = 0; }

			/* Option 押下時のみ小数第1位まで、それ以外は整数に丸める / Round to 1 decimal with Option, to an integer otherwise */
			editText.text = useDecimal ? (Math.round(value * 10) / 10) : Math.round(value);
			event.preventDefault();
		});
	}

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
	 * 9軸（3×3）の基準点ウィジェットを生成する
	 * @param {Group} parentGroup - 追加先のグループ
	 * @returns {Button} 生成したウィジェット
	 */
	function addAnchorWidget(parentGroup) {
		var widget = parentGroup.add("button", undefined, "");
		widget.helpTip = getLabel('tooltip.anchor');
		fixControlSize(widget, ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE);
		widget.onDraw = function () { drawAnchorWidget(this); };
		/* クリックした 3×3 のセルを基準点に設定する（判定は mousedown で行い、クリック座標はコントロール基準）/ Set the anchor from the clicked 3x3 cell (hit-tested in mousedown; coords are control-relative) */
		try {
			widget.addEventListener("mousedown", function (event) {
				var col = Math.floor(event.clientX / (widget.size[0] / 3));
				var row = Math.floor(event.clientY / (widget.size[1] / 3));
				if (col < 0) { col = 0; }
				if (col > 2) { col = 2; }
				if (row < 0) { row = 0; }
				if (row > 2) { row = 2; }
				selectedAnchorIndex = row * 3 + col;
				redrawControl(widget);
			});
		} catch (e) {}
		return widget;
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
		setupPanel(flipPanel);

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
		addAnchorWidget(flipRow);

		addRotateSlider(flipPanel);
	}

	/**
	 * 移動・複製パネル（方向の十字ボタン。クリックで移動／Option＋クリックで複製）を追加する
	 * @param {Window} targetWindow - 追加先のウィンドウ
	 * @returns {void}
	 */
	function buildMovePanel(targetWindow) {
		var directionPanel = targetWindow.add('panel', undefined, getLabel('panel.direction'));
		setupPanel(directionPanel);

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
		setupPanel(optionsPanel);

		var marginGroup = optionsPanel.add('group');
		setupRow(marginGroup, 'left', LABEL_FIELD_SPACING);
		marginGroup.add('statictext', undefined, getLabelWithColon('fieldLabel.margin'));
		var marginInput = marginGroup.add('edittext', undefined, String(DEFAULT_MARGIN));
		marginInput.characters = FIELD_CHARS;
		var unitLabel = marginGroup.add('statictext', undefined, getRulerUnitInfo().label);
		changeValueByArrowKey(marginInput, true); /* 負値を許可（マイナスで重なり方向へ）/ allow negatives (moves toward overlap) */

		var previewBoundsCheck = optionsPanel.add('checkbox', undefined, getLabel('checkbox.previewBounds'));
		previewBoundsCheck.alignment = 'left'; /* パネルは fill なのでチェックボックスだけ左寄せに戻す / The panel fills, so pull the checkbox back to the left */
		previewBoundsCheck.value = DEFAULT_PREVIEW_BOUNDS;

		/**
		 * UI からマージン／プレビュー境界を読む（不正値はここで 0 に落とし、単位表示も更新する）
		 * @returns {{marginPt: number, usePreviewBounds: boolean}} 変形に使う設定
		 */
		function readSettings() {
			var unitInfo = getRulerUnitInfo();
			unitLabel.text = unitInfo.label; /* 単位表示を更新 / refresh the unit label */
			var marginValue = parseFloat(marginInput.text);
			if (isNaN(marginValue)) { marginValue = 0; } /* 負値は許容（マイナスで重なり方向へ）/ Negatives allowed (moves toward overlap) */
			marginInput.text = String(marginValue);
			return {
				marginPt: marginValue * unitInfo.factor,
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
