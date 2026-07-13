#target illustrator
#targetengine "MoveDuplicateObject"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの移動・複製と反転・回転を、アイコンのクリックで即時実行する常駐パレットです（プレビューや適用ボタンはありません）。

- 移動・複製: 上 / 左 / 右 / 下 の矢印アイコンをクリックでその方向へ移動。Option＋クリックで複製
- 反転・回転: アイコンボタンで左右反転／上下反転／45°回転（反時計回り・時計回り）を即時実行。回転は Shift＋クリックで90°単位。Option＋クリックで複製してから変形（Shift＋Option で90°複製回転）。アイコン右の9軸（3×3）ウィジェットで基点を指定
- オプション（マージン／プレビュー境界。移動・複製と反転・回転の両方に適用）
	- マージン: 変形後に基準点の反対方向へ足す余白（単位は定規に追従）。反転・回転では9軸が中心のときは無視
	- プレビュー境界: 線や効果を含む見た目の境界でサイズ・基点を計算

DOM を触る処理（選択取得・移動・複製・反転・回転）はメインエンジンへ BridgeTalk で委譲します。

### 謝辞

kenさん
https://x.com/ken_rainy/status/1472505526768783361

### Overview

A persistent palette that moves/duplicates and flips/rotates the selection immediately on an icon click (no preview, no Apply button).

- Move / duplicate: click an Up / Left / Right / Down arrow icon to move in that direction; Option-click to duplicate
- Flip & rotate: icon buttons flip horizontally/vertically and rotate 45° (CCW/CW), applied immediately; Shift-click rotates in 90° steps; Option-click duplicates before transforming (Shift+Option = duplicate + 90° rotation); a 9-axis (3x3) widget to the right of the icons sets the pivot
- Options (Margin / Preview bounds; applied to both move-duplicate and flip-rotate)
	- Margin: extra gap added after the transform, away from the anchor (unit follows the ruler); ignored for flip/rotate when the 9-axis anchor is centered
	- Preview bounds: uses the visible bounds (including stroke/effects) for size and pivot

DOM work (selection, move, duplicate, flip, rotate) is delegated to the main engine via BridgeTalk.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.0";

// =========================================
// ユーザー設定 / User Settings
// =========================================
var DEFAULT_MARGIN = 0;          /* 元と複製の間に足す余白（定規の単位）/ Extra gap between original and copy (in ruler units) */
var ROTATE_ANGLE = 45;           /* 回転アイコンの回転角（度）。正＝反時計回り / Rotate icon angle (deg); positive = CCW */
var ROTATE_ANGLE_QUARTER = 90;   /* Shift＋クリック時の回転角（度、90°単位）/ Rotate angle when Shift is held (deg; 90° step) */
var FLIP_COLUMNS_PER_ROW = 2;    /* 反転・回転アイコンを何個ごとに改行するか / Flip/rotate icons per row before wrapping */
var ICON_SIZE = 30;              /* 方向・反転・回転アイコン1個の大きさ（px、両パネル共通）/ Size of each direction / flip / rotate icon (px, shared by both panels) */

// =========================================
// ローカライズ / Localization
// =========================================
/* 現在の言語を判定 / Detect the current language */
function getCurrentLang() {
	return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
	dialog: {
		title: { ja: "クイック変形", en: "Quick Transform" }
	},
	panel: {
		direction: { ja: "移動・複製", en: "Move / Duplicate" },
		flip: { ja: "反転・回転", en: "Flip & Rotate" },
		options: { ja: "オプション", en: "Options" }
	},
	direction: {
		up: { ja: "上", en: "Up" },
		left: { ja: "左", en: "Left" },
		right: { ja: "右", en: "Right" },
		down: { ja: "下", en: "Down" }
	},
	field: {
		margin: { ja: "マージン", en: "Margin" }
	},
	checkbox: {
		previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
	},
	tooltip: {
		flipHorizontal: { ja: "選択を水平方向に反転（基準点が基点）／ Option＋クリックで複製", en: "Flip the selection horizontally (about the anchor point) / Option-click to duplicate" },
		flipVertical: { ja: "選択を垂直方向に反転（基準点が基点）／ Option＋クリックで複製", en: "Flip the selection vertically (about the anchor point) / Option-click to duplicate" },
		rotateCCW: { ja: "選択を反時計回りに45°回転（基準点が基点）／ Shift＋クリックで90°単位／ Option＋クリックで複製", en: "Rotate the selection 45° counterclockwise (about the anchor point) / Shift-click for 90° steps / Option-click to duplicate" },
		rotateCW: { ja: "選択を時計回りに45°回転（基準点が基点）／ Shift＋クリックで90°単位／ Option＋クリックで複製", en: "Rotate the selection 45° clockwise (about the anchor point) / Shift-click for 90° steps / Option-click to duplicate" },
		anchor: { ja: "基準点（反転・回転の基点）", en: "Anchor point (pivot for flip / rotate)" },
		moveDuplicate: { ja: "クリックで移動／ Option＋クリックで複製", en: "Click to move / Option-click to duplicate" }
	}
};

/* ラベルをドット区切りキーで取得（現在の言語、途中で見つからなければキーを返す）/ Get a label by dot-separated key (current language; return the key if not found) */
function L(pathKey) {
	var keys = pathKey.split('.');
	var node = LABELS;
	for (var i = 0; i < keys.length; i++) {
		node = node[keys[i]];
		if (node === undefined || node === null) { return pathKey; }
	}
	return node[currentLanguage];
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(pathKey) {
	return L(pathKey) + (currentLanguage === 'ja' ? '：' : ':');
}

// =========================================
// 単位 / Units
// =========================================
/* 定規の単位からラベルと pt 換算係数を求める / Resolve ruler unit label and pt factor */
function getRulerUnitInfo() {
	var rulerUnit = app.preferences.getIntegerPreference("rulerType");
	var unitLabel = "pt";
	var unitFactor = 1.0;

	switch (rulerUnit) {
		case 0: // inch
			unitLabel = "inch";
			unitFactor = 72.0;
			break;
		case 1: // mm
			unitLabel = "mm";
			unitFactor = 72.0 / 25.4;
			break;
		case 2: // pt
			unitLabel = "pt";
			unitFactor = 1.0;
			break;
		case 3: // pica
			unitLabel = "pica";
			unitFactor = 12.0;
			break;
		case 4: // cm
			unitLabel = "cm";
			unitFactor = 72.0 / 2.54;
			break;
		case 5: // Q
			unitLabel = "Q";
			unitFactor = 72.0 / 25.4 * 0.25;
			break;
		case 6: // px
			unitLabel = "px";
			unitFactor = 1.0;
			break;
		default:
			unitLabel = "pt";
			unitFactor = 1.0;
	}

	return { label: unitLabel, factor: unitFactor };
}

// =========================================
// 常駐参照 / Persistent reference
// =========================================
/* パレットを常駐エンジンの変数に保持して GC を回避（初期化子なしで再実行時も参照を維持）/ Keep the palette in an engine-persistent var to avoid GC (no initializer so re-runs keep the reference) */
var paletteWindow;
var isBusy = false; /* 委譲の再入防止 / Re-entrancy guard for delegation */

// =========================================
// worker（メインエンジンで実行する DOM 処理）/ Worker (DOM work run on the main engine)
// 注意: 行コメント（//）禁止・/* */ のみ・各文はセミコロンで終える（toString が改行を消すため）
// Note: no // comments, /* */ only, end every statement with ';' (toString strips newlines)
// =========================================
/* 選択に対して移動・複製を1回実行する（アイコンのクリックで即時適用）/ Apply the move/duplicate once to the selection (immediate on icon click) */
function workerApply(options) {
	if (app.documents.length === 0) { return "NODOC"; };
	var docSelection = app.activeDocument.selection;
	if (!docSelection || docSelection.length === 0) { return "NOSEL"; };
	try {
		var selectionSize = workerSelectionSize(docSelection, options.usePreviewBounds);
		var offset = workerOffset(options.direction, selectionSize.width, selectionSize.height, options.marginPt);
		var buildResult = workerBuild(docSelection, offset, options.duplicate);
		/* 複製時は元の選択を外し、生成した複製だけを選択状態にする / On duplicate, deselect the originals and select only the created copies */
		if (options.duplicate && buildResult && buildResult.items && buildResult.items.length > 0) {
			try { app.activeDocument.selection = buildResult.items; } catch (selectionError) { };
		};
		app.redraw();
		return "OK";
	} catch (e) {
		return "ERR:" + e;
	};
};

/* 選択範囲全体のサイズを求める / Get the overall size of the selection */
function workerSelectionSize(items, usePreviewBounds) {
	var minLeft = null, maxRight = null, minBottom = null, maxTop = null;
	for (var i = 0; i < items.length; i++) {
		var bounds = usePreviewBounds ? items[i].visibleBounds : items[i].geometricBounds;
		if (minLeft === null || bounds[0] < minLeft) { minLeft = bounds[0]; };
		if (maxTop === null || bounds[1] > maxTop) { maxTop = bounds[1]; };
		if (maxRight === null || bounds[2] > maxRight) { maxRight = bounds[2]; };
		if (minBottom === null || bounds[3] < minBottom) { minBottom = bounds[3]; };
	};
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

/* 複製（またはそのまま）を1つ分だけ移動し、生成物の配列を返す / Move by one step (duplicating or not) and return the created items */
function workerBuild(items, offset, shouldDuplicate) {
	var createdItems = [];
	if (shouldDuplicate) {
		for (var i = 0; i < items.length; i++) {
			var duplicatedItem = items[i].duplicate();
			duplicatedItem.translate(offset[0], offset[1]);
			createdItems.push(duplicatedItem);
		};
		return { mode: "dup", items: createdItems, offset: offset };
	};
	for (var k = 0; k < items.length; k++) {
		items[k].translate(offset[0], offset[1]);
		createdItems.push(items[k]);
	};
	return { mode: "move", items: createdItems, offset: offset };
};

/* 委譲する worker 関数の全登録（追加漏れ防止）/ All worker functions to delegate (avoid missing registrations) */
var WORKER_FUNCTIONS = [workerApply, workerBuild, workerSelectionSize, workerOffset];

// =========================================
// BridgeTalk 委譲 / BridgeTalk delegation
// =========================================
/* worker 関数群をソース文字列に連結 / Concatenate all worker functions into a source string */
function buildWorkerSource() {
	var source = "";
	for (var i = 0; i < WORKER_FUNCTIONS.length; i++) {
		source += WORKER_FUNCTIONS[i].toString() + "\n";
	}
	return source;
}

/* オプションを JS オブジェクトリテラル文字列へ / Serialize options into a JS object-literal string */
function optionsLiteral(options) {
	return "{" +
		"direction:'" + options.direction + "'," +
		"marginPt:" + options.marginPt + "," +
		"duplicate:" + (options.duplicate ? "true" : "false") + "," +
		"usePreviewBounds:" + (options.usePreviewBounds ? "true" : "false") +
		"}";
}

/* worker 呼び出し式をメインエンジンで実行し結果マーカーを返す（同期 send＋holder）/ Run a worker call expression on the main engine and return the result marker (sync send + holder) */
function delegateToMain(callExpression) {
	var payload = buildWorkerSource() + callExpression;
	var encoded = encodeURIComponent(payload);
	var resultHolder = { result: "ERR:timeout" };

	var bridge = new BridgeTalk();
	bridge.target = "illustrator";
	bridge.body = 'eval(decodeURIComponent("' + encoded + '"));';
	bridge.onResult = function (message) { resultHolder.result = String(message.body); };
	bridge.onError = function (message) { resultHolder.result = "ERR:" + String(message.body); };
	bridge.send(10);

	return resultHolder.result;
}

/* 移動・複製の即時実行を委譲 / Delegate the immediate move/duplicate */
function delegateApply(options) {
	return delegateToMain("workerApply(" + optionsLiteral(options) + ");");
}

// =========================================
// 反転・回転の委譲（メインエンジンで DOM を変形）/ Flip & rotate delegation (transform the DOM on the main engine)
// =========================================
/* 基準点（9軸）: 0..8 を行優先（0=左上, 4=中央, 8=右下）。反転・回転の基点に使う / 9-axis anchor 0..8 row-major (0=top-left, 4=center, 8=bottom-right); used as the flip/rotate pivot */
var selectedAnchorIndex = 4;

/* showPalette がマージン／プレビュー境界を読む関数をここへ公開（反転・回転からも同じ設定を利用）。未生成時は既定値 / showPalette publishes a reader for margin / preview-bounds here (flip/rotate reuse the same settings); defaults until built */
var readTransformOptions = null;

/* メインエンジン（target="illustrator"）へ本文コードを送って実行（同期 send＝完了まで待つので isBusy ガードが有効）/ Send body code to the main engine and run it (synchronous send so the isBusy guard is effective) */
function runInMainEngine(bodyCode) {
	try {
		var bridge = new BridgeTalk();
		bridge.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
		bridge.body = bodyCode;
		bridge.onError = function (message) {
			try { $.writeln("AiMoveDuplicateRotateReflect BridgeTalk error: " + message.body); } catch (e) { }
		};
		bridge.send(10); /* 完了まで待つ（移動・複製の delegateToMain と同じ）/ wait for completion (same as delegateToMain) */
	} catch (e) {
		/* BridgeTalk 不可時は同一エンジンで直接実行 / Fallback: run directly in this engine */
		try { eval(bodyCode); } catch (e2) { }
	}
}

/* 選択した 9 軸の基準点に対応する anchorX/anchorY の式（body 内の left/right/top/bottom を参照）を返す */
/* Return the anchorX/anchorY expressions for the selected 9-axis point (referencing left/right/top/bottom in the body) */
function getAnchorExpressions() {
	var col = selectedAnchorIndex % 3;
	var row = Math.floor(selectedAnchorIndex / 3);
	/* 入れ子三項は ExtendScript が左結合で誤評価するため右結合を括弧で明示（無いと左＝中央・上＝中央になる）/ Parenthesize nested ternaries; ExtendScript misparses them left-associatively (else left==center, top==center) */
	var xExpr = (col === 0) ? "left" : ((col === 1) ? "((left+right)/2)" : "right");
	var yExpr = (row === 0) ? "top" : ((row === 1) ? "((top+bottom)/2)" : "bottom");
	return "var anchorX=" + xExpr + ",anchorY=" + yExpr + ";";
}

/* 共通の委譲本体：可視バウンディング測定→基準点→合成行列を1回適用→再描画。matrixCode が matrix を組み立てる */
/* Shared delegation body: measure visible bounds -> anchor -> apply one matrix -> redraw; matrixCode builds the matrix */
/* visibleBounds が取れない/変形できない項目（ロック・非表示・ガイド等）は try/catch でスキップ。測定できた項目が無ければ中断 */
/* Skip items whose bounds/transform fail (locked, hidden, guides, ...) via try/catch; abort if nothing could be measured */
/* duplicate=true のとき、変形前に選択を複製して複製側だけを変形し、複製を新しい選択にする（Option＋クリック）*/
/* When duplicate=true, duplicate the selection first and transform only the copies, then select the copies (Option+click) */
/* マージンは基準点から中心の反対方向へ変形後に平行移動して足す（中心基点=index 4 は marginX/Y が 0 になり無視される）*/
/* Margin is added by translating the result away from center after the transform (center anchor = index 4 yields marginX/Y = 0, i.e. ignored) */
/* usePreviewBounds=true なら可視境界、false なら幾何境界で基準点を測る / Measure the anchor from visible bounds when usePreviewBounds=true, geometric bounds otherwise */
function btTransformSelection(matrixCode, duplicate, marginPt, usePreviewBounds) {
	var dupFlag = duplicate ? 'true' : 'false';
	var boundsProp = usePreviewBounds ? 'visibleBounds' : 'geometricBounds';
	var col = selectedAnchorIndex % 3;
	var row = Math.floor(selectedAnchorIndex / 3);
	var margin = Number(marginPt) || 0;
	var marginX = ((col === 0) ? -1 : ((col === 2) ? 1 : 0)) * margin;
	var marginY = ((row === 0) ? 1 : ((row === 2) ? -1 : 0)) * margin;
	var body = '' +
		'if(app.documents.length>0){' +
		'var doc=app.activeDocument,selection=doc.selection;' +
		'if(selection&&selection.length>0){' +
		'var left=Infinity,top=-Infinity,right=-Infinity,bottom=Infinity,measured=false;' +
		'for(var i=0;i<selection.length;i++){try{var b=selection[i].' + boundsProp + ';if(b[0]<left)left=b[0];if(b[1]>top)top=b[1];if(b[2]>right)right=b[2];if(b[3]<bottom)bottom=b[3];measured=true;}catch(e){}}' +
		'if(measured){' +
		getAnchorExpressions() +
		'var matrix=app.getIdentityMatrix();' + matrixCode +
		'var targets=[];' +
		'if(' + dupFlag + '){for(var i=0;i<selection.length;i++){try{targets.push(selection[i].duplicate());}catch(e){}}}else{for(var i=0;i<selection.length;i++){targets.push(selection[i]);}}' +
		'for(var i=0;i<targets.length;i++){try{targets[i].transform(matrix,true,true,true,true,1,Transformation.DOCUMENTORIGIN);}catch(e){}}' +
		'var marginX=' + marginX + ',marginY=' + marginY + ';' +
		'if(marginX!==0||marginY!==0){for(var i=0;i<targets.length;i++){try{targets[i].translate(marginX,marginY);}catch(e){}}}' +
		'if(' + dupFlag + '){try{doc.selection=targets;}catch(e){}}' +
		'app.redraw();' +
		'}}}';
	runInMainEngine(body);
}

/* 選択を、選択した 9 軸基準点を基点に反転（水平=-100,100／垂直=100,-100）。DOM 操作なのでメインエンジンへ委譲 */
/* Flip the selection about the selected 9-axis anchor (H=-100,100 / V=100,-100); delegated to the main engine since it touches the DOM */
/* 係数は数値化して埋め込む（'1-' + (-1) だと "1--1" になりデクリメント解釈で構文エラーになるため）*/
/* Coefficients are precomputed as numbers ('1-' + (-1) would yield "1--1", a decrement → syntax error) */
function btFlipSelection(scaleX, scaleY, duplicate, marginPt, usePreviewBounds) {
	var scaleFractionX = Number(scaleX) / 100; /* -1 or 1 */
	var scaleFractionY = Number(scaleY) / 100;
	btTransformSelection(
		'matrix.mValueA=' + scaleFractionX + ';matrix.mValueD=' + scaleFractionY + ';' +
		'matrix.mValueTX=anchorX*' + (1 - scaleFractionX) + ';matrix.mValueTY=anchorY*' + (1 - scaleFractionY) + ';',
		duplicate, marginPt, usePreviewBounds
	);
}

/* 選択を、選択した 9 軸基準点を基点に回転（angleDegrees：正＝反時計回り／負＝時計回り）。DOM 操作なのでメインエンジンへ委譲 */
/* Rotate the selection about the selected 9-axis anchor (angleDegrees: positive = CCW / negative = CW); delegated to the main engine since it touches the DOM */
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
// ボタン描画 / Button drawing
// =========================================
/* UI が明るいテーマかを判定（取得失敗時は暗い側）/ Detect a light UI theme (fall back to dark on failure) */
function isLightUI() {
	try {
		return app.preferences.getRealPreference('uiBrightness') > 0.5;
	} catch (e) {
		return false;
	}
}

/* 右向き矢印の各点を方向キーに合わせて回転 / Rotate a right-arrow point to match the direction key */
function transformArrowPoint(directionKey, x, y) {
	if (directionKey === 'left') { return [-x, y]; }
	if (directionKey === 'down') { return [-y, x]; }
	if (directionKey === 'up') { return [y, -x]; }
	return [x, y]; /* right */
}

/* 方向キーの矢印を白抜き（アウトライン）で描画 / Draw an outlined (knockout) arrow for the direction key */
function drawArrow(graphics, directionKey, w, h, color) {
	var iconSize = Math.min(w, h);
	var tip = iconSize * 0.32;
	var shaft = iconSize * 0.11;
	var headHalf = iconSize * 0.27;   /* 矢じりの半分の高さ（大きめ）/ half height of the arrowhead (larger) */
	var headBase = tip - iconSize * 0.34; /* 矢じりの付け根（長め）/ base of the arrowhead (longer) */
	var basePoints = [
		[-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
		[tip, 0],
		[headBase, headHalf], [headBase, shaft], [-tip, shaft]
	];
	var cx = w / 2, cy = h / 2;
	graphics.newPath();
	for (var i = 0; i < basePoints.length; i++) {
		var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
		if (i === 0) { graphics.moveTo(cx + point[0], cy + point[1]); }
		else { graphics.lineTo(cx + point[0], cy + point[1]); }
	}
	graphics.closePath();
	/* 白抜き：塗らずに輪郭線だけ描く / Knockout: stroke the outline only, no fill */
	graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1.2));
}

// =========================================
// 反転・回転（アイコン UI）/ Flip & Rotate (icon UI)
// =========================================
/* 4つのアイコンボタン定義（左右反転／上下反転／回転CCW／回転CW）/ The four icon buttons (flip H / flip V / rotate CCW / rotate CW) */
var iconButtonDefs = [
	{ name: "FLIP_HORIZONTAL", icon: "flipHorizontal", tooltip: "tooltip.flipHorizontal" },
	{ name: "FLIP_VERTICAL", icon: "flipVertical", tooltip: "tooltip.flipVertical" },
	{ name: "ROTATE", icon: "rotate", tooltip: "tooltip.rotateCCW" },
	{ name: "ROTATE_FLIP", icon: "rotateFlip", tooltip: "tooltip.rotateCW" }
];

/* アイコンの配色（showPalette() で UI 明暗から設定）/ Icon colors (set in showPalette() from the light/dark UI) */
var iconColor, iconBorderColor, iconBaseBg, iconHoverBg;

/* UI 明度(0..1)を取得、取得失敗時は 1（明るい）を返す / Get UI brightness (0..1); 1 on failure */
function getUIBrightness() {
	try {
		var brightness = app.preferences.getRealPreference("uiBrightness");
		if (brightness < 0) brightness = 0;
		if (brightness > 1) brightness = 1;
		return brightness;
	} catch (e) {
		return 1;
	}
}

/* グレーの RGBA を返す（0..1 にクランプ）/ Build a clamped gray RGBA */
function grayColor(value) {
	if (value < 0) value = 0;
	if (value > 1) value = 1;
	return [value, value, value, 1];
}

/* UI の明暗に合わせてアイコン色・背景色・枠線を決める（showPalette() から呼ぶ）/ Decide icon, background and border colors from the light/dark UI (called from showPalette()) */
function initIconColors() {
	var lightUI = isLightUI();
	var uiBrightness = getUIBrightness();
	iconColor = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
	/* ライトモードは背景を塗らず（ウィンドウ色）薄いグレーの枠を付ける。ダークは枠なし / Light: no fill (window color) + light gray border; dark: no border */
	iconBorderColor = lightUI ? [0.65, 0.65, 0.65, 1] : null;
	iconBaseBg = lightUI ? grayColor(uiBrightness) : [0.28, 0.28, 0.28, 1];
	/* マウスオーバー時の背景（ライトは少し暗く、ダークは少し明るく）/ Hover background (slightly darker in light, lighter in dark) */
	iconHoverBg = lightUI ? grayColor(uiBrightness - 0.10) : [0.38, 0.38, 0.38, 1];
}

/* コントロールを再描画（notify は環境により例外を投げ得るので保護）/ Redraw a control (notify can throw in some environments) */
function redrawControl(control) {
	try { control.notify("onDraw"); } catch (e) { }
}

/* マウスオーバーで背景を切り替えるためホバー状態を button.isHover に反映して再描画（方向・反転回転アイコン共通）/ Track hover in button.isHover and repaint so the background can change on mouseover (shared by direction and flip/rotate icons) */
function attachHover(button) {
	try {
		button.addEventListener("mouseover", function () { button.isHover = true; redrawControl(button); });
		button.addEventListener("mouseout", function () { button.isHover = false; redrawControl(button); });
	} catch (e) { }
}

/* アイコンに対応する変形を実行する（duplicate=true で複製してから変形。quarter=true で回転を90°単位に。マージン／プレビュー境界はオプションパネルの設定を使用）*/
/* Run the transform for an icon (duplicate first when duplicate=true; quarter=true rotates in 90° steps; margin / preview-bounds come from the Options panel) */
function handleIconAction(name, duplicate, quarter) {
	if (isBusy) { return; } /* 連打の再入防止（移動・複製と共通のガード）/ Guard against rapid re-entry (shared with move/duplicate) */
	isBusy = true;
	try {
		var settings = readTransformOptions ? readTransformOptions() : { marginPt: 0, usePreviewBounds: false };
		var marginPt = settings.marginPt || 0;
		var usePreviewBounds = settings.usePreviewBounds === true;
		var rotateAngle = quarter ? ROTATE_ANGLE_QUARTER : ROTATE_ANGLE; /* Shift で90°単位 / Shift → 90° step */
		if (name === "FLIP_HORIZONTAL") {
			btFlipSelection(-100, 100, duplicate, marginPt, usePreviewBounds);
		} else if (name === "FLIP_VERTICAL") {
			btFlipSelection(100, -100, duplicate, marginPt, usePreviewBounds);
		} else if (name === "ROTATE") {
			btRotateSelection(rotateAngle, duplicate, marginPt, usePreviewBounds);     /* 反時計回り / counterclockwise */
		} else if (name === "ROTATE_FLIP") {
			btRotateSelection(-rotateAngle, duplicate, marginPt, usePreviewBounds);    /* 時計回り / clockwise */
		}
	} finally {
		isBusy = false;
	}
}

/* アイコンボタンを1つ生成して登録 / Create and register one icon button */
function addIconButton(parentGroup, buttonDef) {
	var button = parentGroup.add("button", undefined, "");
	button.helpTip = L(buttonDef.tooltip);
	/* 移動・複製の方向ボタン（ICON_SIZE）と同じ大きさに合わせる / Match the size of the move/duplicate direction buttons (ICON_SIZE) */
	button.preferredSize = [ICON_SIZE, ICON_SIZE];
	button.minimumSize = [ICON_SIZE, ICON_SIZE];
	button.maximumSize = [ICON_SIZE, ICON_SIZE];
	button.iconType = buttonDef.icon;
	button.isHover = false;
	button.onDraw = function () {
		drawIcon(this);
	};
	/* Option＝複製、Shift＝回転を90°単位（回転アイコンのみ有効）/ Option = duplicate, Shift = 90° rotation step (only affects the rotate icons) */
	button.onClick = function () {
		var duplicate = false, quarter = false;
		try {
			var keyboard = ScriptUI.environment.keyboardState;
			duplicate = keyboard.altKey === true;
			quarter = keyboard.shiftKey === true;
		} catch (e) { }
		handleIconAction(buttonDef.name, duplicate, quarter);
	};
	attachHover(button);
}

/* 9軸（3×3）の基準点ウィジェットを生成 / Create the 9-axis (3x3) anchor widget */
function addAnchorWidget(parentGroup) {
	var widget = parentGroup.add("button", undefined, "");
	widget.helpTip = L('tooltip.anchor');
	widget.preferredSize = [66, 66];
	widget.minimumSize = [66, 66];
	widget.maximumSize = [66, 66];
	widget.onDraw = function () {
		drawAnchorWidget(this);
	};
	widget.onClick = function () {
		/* クリック位置の判定は mousedown 側で行う（onClick は何もしない）/ Cell hit-testing happens in mousedown; onClick is a no-op */
	};
	/* クリックした 3×3 のセルを基準点に設定（クリック座標はコントロール基準）/ Set the anchor from the clicked 3x3 cell (coords are control-relative) */
	try {
		widget.addEventListener("mousedown", function (event) {
			var cellWidth = widget.size[0] / 3;
			var cellHeight = widget.size[1] / 3;
			var col = Math.floor(event.clientX / cellWidth);
			var row = Math.floor(event.clientY / cellHeight);
			if (col < 0) col = 0;
			if (col > 2) col = 2;
			if (row < 0) row = 0;
			if (row > 2) row = 2;
			selectedAnchorIndex = row * 3 + col;
			redrawControl(widget);
		});
	} catch (e) { }
	return widget;
}

/* 9軸ウィジェットを描画（外周の□をケイ線でつなぐ・中央は独立）/ Draw the 9-axis widget (outer squares joined by rules; center stands alone) */
function drawAnchorWidget(widget) {
	var graphics = widget.graphics;
	var width = widget.size[0];
	var height = widget.size[1];

	try {
		/* 背景は塗らずコントロール地色（パネルと同色）で塗って透過に見せる / Paint the control's own background color (matches the panel) so the widget looks transparent */
		graphics.rectPath(0, 0, width, height);
		graphics.fillPath(graphics.backgroundColor);
	} catch (e0) { }

	var cellSize = 9;    /* 四角のサイズ（44→66 の 1.5 倍に合わせて拡大）/ Square size (scaled 1.5x with the 44→66 widget) */
	var cellGap = 7.5;   /* 四角どうしの間隔 / Gap between squares */
	var cellStep = cellSize + cellGap;
	var gridSize = cellSize * 3 + cellGap * 2;
	var originX = Math.round((width - gridSize) / 2);
	var originY = Math.round((height - gridSize) / 2);

	function anchorCellX(index) { return originX + (index % 3) * cellStep; }
	function anchorCellY(index) { return originY + Math.floor(index / 3) * cellStep; }

	/* 中央(4)を除く外周の□どうしをケイ線（隙間部分）でつなぐ / Join the outer squares (except center 4) with rules across the gaps */
	var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
	var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, iconColor, 1);
	for (var i = 0; i < connections.length; i++) {
		var cellA = connections[i][0];
		var cellB = connections[i][1];
		var cellAX = anchorCellX(cellA);
		var cellAY = anchorCellY(cellA);
		var cellBX = anchorCellX(cellB);
		var cellBY = anchorCellY(cellB);
		graphics.newPath();
		if (cellB - cellA === 1) {
			/* 横方向：右隣の□へ / Horizontal: to the square on the right */
			graphics.moveTo(cellAX + cellSize, cellAY + cellSize / 2);
			graphics.lineTo(cellBX, cellBY + cellSize / 2);
		} else {
			/* 縦方向：下の□へ / Vertical: to the square below */
			graphics.moveTo(cellAX + cellSize / 2, cellAY + cellSize);
			graphics.lineTo(cellBX + cellSize / 2, cellBY);
		}
		graphics.strokePath(linePen);
	}

	for (var index = 0; index < 9; index++) {
		drawAnchorCell(graphics, anchorCellX(index), anchorCellY(index), cellSize, index === selectedAnchorIndex);
	}
}

/* 基準点セルの□を1つ描画（全セルにグレー枠、選択中は枠の内側を塗る）/ Draw one anchor-cell square (gray border on every cell; fill inside the border when selected) */
function drawAnchorCell(graphics, x, y, size, selected) {
	/* まず全セル共通のグレー枠を描く（中央も他の□と同じ枠を持つ）/ First draw the shared gray border on every cell (center gets the same border as the others) */
	graphics.newPath();
	graphics.moveTo(x, y);
	graphics.lineTo(x + size, y);
	graphics.lineTo(x + size, y + size);
	graphics.lineTo(x, y + size);
	graphics.closePath();
	graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconColor, 1));

	/* 選択中は枠を残したまま内側を塗ってハイライト / When selected, fill inside the border so the border stays visible */
	if (selected) {
		var inset = 1.5;
		graphics.newPath();
		graphics.moveTo(x + inset, y + inset);
		graphics.lineTo(x + size - inset, y + inset);
		graphics.lineTo(x + size - inset, y + size - inset);
		graphics.lineTo(x + inset, y + size - inset);
		graphics.closePath();
		graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, iconColor));
	}
}

/* アイコンボタンの背景・枠線を描き、種類に応じた図柄を描画 / Draw the icon button background/border and dispatch to the right glyph */
function drawIcon(button) {
	var graphics = button.graphics;
	var width = button.size[0];
	var height = button.size[1];
	var backgroundColor = (button.isHover === true) ? iconHoverBg : iconBaseBg;

	try {
		graphics.rectPath(0, 0, width, height);
		graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));

		/* ライトモードは薄いグレーの枠線を描く / Light mode draws a light gray border */
		if (iconBorderColor) {
			graphics.rectPath(0.5, 0.5, width - 1, height - 1);
			graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconBorderColor, 1));
		}
	} catch (e1) {
		try {
			graphics.drawOSControl();
		} catch (e2) { }
	}

	if (button.iconType === "rotate") {
		drawRotateIcon(graphics, width, height, false);
	} else if (button.iconType === "rotateFlip") {
		drawRotateIcon(graphics, width, height, true);
	} else {
		drawFlipIcon(graphics, button.iconType, width, height);
	}
}

/* 反転アイコン（軸の点線＋向かい合う三角形）を描画 / Draw a flip icon (dotted axis + opposing triangles) */
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

/* 回転アイコン（一定幅の実線弧＋点線弧＋矢じり）を描画。mirror で左右反転 / Draw a rotate icon (uniform-width solid arc + dotted arc + arrowhead); mirror flips it horizontally */
function drawRotateIcon(graphics, width, height, mirror) {
	var color = iconColor;
	var centerX = width / 2;
	var centerY = height / 2 + 1;
	var radius = 7.5;
	var mirrorSign = mirror ? -1 : 1;   /* 左右反転のときは x をミラーする / Mirror x when flipped */
	var headDeg = 232;                  /* 矢じりの位置（左上）/ Arrowhead position (top-left) */

	/* 実線の弧は一定幅・単一パスでなめらかに描く / Solid arc: uniform width, single path for smoothness */
	strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign, 1.8, 1.8);
	/* 下側は四角い点線 / Square-dotted arc on the lower side */
	drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign, 1.8);

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

/* 円弧を線分で近似して描く。線幅一定なら継ぎ目の出ない単一パス、変化させる場合のみセグメントごとに描く（先細り）/
   Stroke an arc as line segments: a single seamless path when the width is constant, per-segment only when it tapers */
function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign, startWidth, endWidth) {
	var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));

	function arcX(ratio) { return centerX + mirrorSign * radius * Math.cos((startDeg + (endDeg - startDeg) * ratio) * Math.PI / 180); }
	function arcY(ratio) { return centerY + radius * Math.sin((startDeg + (endDeg - startDeg) * ratio) * Math.PI / 180); }

	/* 幅一定：1本の連続パスで描くのでガタつかない / Constant width: one continuous path, so no jaggedness */
	if (startWidth === endWidth) {
		graphics.newPath();
		graphics.moveTo(arcX(0), arcY(0));
		for (var i = 1; i <= segments; i++) {
			graphics.lineTo(arcX(i / segments), arcY(i / segments));
		}
		graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, startWidth));
		return;
	}

	/* 幅可変：セグメントごとに線幅を補間 / Variable width: interpolate per segment */
	for (var j = 1; j <= segments; j++) {
		var ratio = j / segments;
		var lineWidth = startWidth + (endWidth - startWidth) * ratio;
		graphics.newPath();
		graphics.moveTo(arcX((j - 1) / segments), arcY((j - 1) / segments));
		graphics.lineTo(arcX(ratio), arcY(ratio));
		graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, lineWidth));
	}
}

/* 円弧に沿って四角い点線を描く / Draw a square-dotted line along an arc */
function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign, dotWidth) {
	var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, dotWidth);
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

/* 3点の三角形を塗り or 線で描く / Draw a triangle from 3 points, filled or stroked */
function drawTriangle(graphics, points, color, fill) {
	graphics.newPath();
	graphics.moveTo(points[0][0], points[0][1]);
	graphics.lineTo(points[1][0], points[1][1]);
	graphics.lineTo(points[2][0], points[2][1]);
	graphics.closePath();
	if (fill) {
		graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
	} else {
		graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1.2));
	}
}

/* 水平または垂直の点線を描く / Draw a horizontal or vertical dotted line */
function drawDottedLine(graphics, x1, y1, x2, y2, color) {
	var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1);
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

// =========================================
// UI 部品 / UI helpers
// =========================================
/* ↑↓キーで値を増減（Shift=±10, Option=±0.1）。integerOnly で 0.1 無効、allowNegative で負値許可、onValueChange は増減時に呼ぶ / Arrow-key stepping (Shift=±10, Option=±0.1); integerOnly disables 0.1, allowNegative permits negatives, onValueChange fires on step */
function changeValueByArrowKey(editText, integerOnly, onValueChange, allowNegative) {
	editText.addEventListener("keydown", function (event) {
		var value = Number(editText.text);
		if (isNaN(value)) return;

		var keyboard = ScriptUI.environment.keyboardState;
		var useDecimal = keyboard.altKey && !integerOnly;
		var delta = 1;

		if (keyboard.shiftKey) {
			delta = 10;
			/* Shift 押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
			if (event.keyName === "Up") {
				value = Math.ceil((value + 1) / delta) * delta;
				event.preventDefault();
			} else if (event.keyName === "Down") {
				value = Math.floor((value - 1) / delta) * delta;
				if (!allowNegative && value < 0) value = 0;
				event.preventDefault();
			}
		} else if (useDecimal) {
			delta = 0.1;
			/* Option 押下時は0.1単位で増減 / Step by 0.1 when Option is held */
			if (event.keyName === "Up") {
				value += delta;
				event.preventDefault();
			} else if (event.keyName === "Down") {
				value -= delta;
				event.preventDefault();
			}
		} else {
			delta = 1;
			if (event.keyName === "Up") {
				value += delta;
				event.preventDefault();
			} else if (event.keyName === "Down") {
				value -= delta;
				if (!allowNegative && value < 0) value = 0;
				event.preventDefault();
			}
		}

		if (useDecimal) {
			/* 小数第1位までに丸め / Round to 1 decimal place */
			value = Math.round(value * 10) / 10;
		} else {
			/* 整数に丸め / Round to integer */
			value = Math.round(value);
		}

		editText.text = value;

		/* 増減したときだけ通知 / Notify only when actually stepped */
		if ((event.keyName === "Up" || event.keyName === "Down") && onValueChange) {
			onValueChange();
		}
	});
}

// =========================================
// パネル共通レイアウト / Shared panel layout
// =========================================
/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定（縦並び・子要素は横いっぱい・共通余白。spacing は任意）/ Apply shared panel layout (column, children fill width, shared margins; spacing optional) */
function setupPanel(panel, spacing) {
	panel.orientation = "column";
	panel.alignChildren = ["fill", "top"];
	panel.alignment = "fill";
	panel.margins = PANEL_MARGINS;
	panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// パネル構築 / Panel builders
// =========================================
/* 反転・回転パネル（アイコン2×2＋9軸ウィジェット。クリックで即時実行）を win に追加 / Add the Flip-Rotate panel (2x2 icons + 9-axis widget; runs immediately on click) to win */
function buildFlipPanel(win) {
	var flipPanel = win.add('panel', undefined, L('panel.flip'));
	setupPanel(flipPanel);

	/* アイコン2×2（左）と9軸ウィジェット（右）を横並び。パネルは fill なので中央寄せは行側で指定 / 2x2 icons (left) and the 9-axis widget (right); panel fills, so center on the row */
	var flipRow = flipPanel.add('group');
	flipRow.orientation = 'row';
	flipRow.alignChildren = ['center', 'center'];
	flipRow.alignment = ['center', 'top'];
	flipRow.spacing = 12;

	/* アイコンボタンを FLIP_COLUMNS_PER_ROW 個ごとに改行して並べる / Lay out icon buttons, wrapping every FLIP_COLUMNS_PER_ROW */
	var iconGrid = flipRow.add('group');
	iconGrid.orientation = 'column';
	iconGrid.alignChildren = ['center', 'center'];
	iconGrid.spacing = 7;

	var iconRow = null;
	for (var iconIndex = 0; iconIndex < iconButtonDefs.length; iconIndex++) {
		if ((iconIndex % FLIP_COLUMNS_PER_ROW) === 0) {
			iconRow = iconGrid.add('group');
			iconRow.orientation = 'row';
			iconRow.alignChildren = ['center', 'center'];
			iconRow.spacing = 7;
		}
		addIconButton(iconRow, iconButtonDefs[iconIndex]);
	}

	/* 9軸（3×3）の基準点ウィジェット（アイコンの右。反転・回転の基点を指定）/ 9-axis (3x3) anchor widget (to the right of the icons; sets the flip/rotate pivot) */
	addAnchorWidget(flipRow);
}

/* 移動・複製パネル（方向の十字ボタン。クリックで移動／Option＋クリックで複製）を win に追加し、方向ショートカット（e/d/s/f）も登録 / Add the Move-Duplicate panel (direction cross; click moves / Option-click duplicates) to win; also registers the direction shortcuts (e/d/s/f) */
function buildMovePanel(win) {
	var directionPanel = win.add('panel', undefined, L('panel.direction'));
	setupPanel(directionPanel);

	/* 方向ボタンは反転・回転アイコンに合わせる（グレー枠・白系背景・ホバーで背景変化・ダーク対応）/ Direction buttons match the flip/rotate icons (gray border, white-ish bg, hover background, dark-aware) */
	var uiColors = {
		background: iconBaseBg,
		hoverBackground: iconHoverBg,
		foreground: iconColor,
		border: iconBorderColor
	};

	var CELL_SIZE = ICON_SIZE; /* 方向ボタン1セルの大きさ（反転・回転アイコンと共通）/ Direction cell size (shared with the flip/rotate icons) */
	/* 方向実行のショートカット（上:e / 下:d / 左:s / 右:f）/ Direction shortcuts (up:e / down:d / left:s / right:f) */
	var directionShortcuts = { up: 'E', down: 'D', left: 'S', right: 'F' };

	/* ボタン1つを onDraw で描画（ホバーで背景変化）/ Draw one button via onDraw (background changes on hover) */
	function drawDirectionButton(control) {
		var graphics = control.graphics;
		var w = control.size[0], h = control.size[1];
		graphics.rectPath(0, 0, w, h);
		graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, (control.isHover === true) ? uiColors.hoverBackground : uiColors.background));
		if (uiColors.border) {
			graphics.rectPath(0.5, 0.5, w - 1, h - 1);
			graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, uiColors.border, 1));
		}
		drawArrow(graphics, control.directionKey, w, h, uiColors.foreground);
	}

	/* 指定方向へ即時実行。Option＋クリック（またはキー）で複製、通常は移動。マージン／プレビュー境界はオプションパネルの設定を使用 / Run immediately; Option duplicates, otherwise moves; margin / preview-bounds come from the Options panel */
	function executeDirection(directionKey) {
		if (isBusy) { return; }
		isBusy = true;
		try {
			var settings = readTransformOptions ? readTransformOptions() : { marginPt: 0, usePreviewBounds: false };
			var duplicate = false;
			try { duplicate = ScriptUI.environment.keyboardState.altKey === true; } catch (e) { }
			delegateApply({
				direction: directionKey,
				marginPt: settings.marginPt || 0,
				usePreviewBounds: settings.usePreviewBounds === true,
				duplicate: duplicate
			});
		} finally {
			isBusy = false;
		}
	}

	/* 空セルを追加 / Add an empty cell */
	function addSpacer(row) {
		var spacer = row.add('statictext', undefined, '');
		spacer.preferredSize = [CELL_SIZE, CELL_SIZE];
	}

	/* 方向ボタンを追加（クリックでその方向へ移動・複製を即時実行）/ Add a direction button (click runs move/duplicate in that direction immediately) */
	function addDirectionButton(row, directionKey) {
		var button = row.add('iconbutton', undefined, undefined, { style: 'toolbutton' });
		button.preferredSize = [CELL_SIZE, CELL_SIZE];
		button.directionKey = directionKey;
		button.isHover = false;
		button.helpTip = L('direction.' + directionKey)
			+ (directionShortcuts[directionKey] ? '  [' + directionShortcuts[directionKey] + ']' : '')
			+ '  —  ' + L('tooltip.moveDuplicate');
		button.onDraw = function () { drawDirectionButton(this); };
		button.onClick = function () { executeDirection(this.directionKey); };
		attachHover(button);
	}

	/* 十字ボタンは専用サブグループへ（行間を密に保つ）。パネルは fill なので中央寄せはグループ側で指定 / Keep the cross in its own subgroup (tight rows); panel fills, so center on the group */
	var crossGroup = directionPanel.add('group');
	crossGroup.orientation = 'column';
	crossGroup.alignChildren = 'center';
	crossGroup.alignment = ['center', 'top'];
	crossGroup.spacing = 2;

	var topRow = crossGroup.add('group');
	topRow.spacing = 2;
	addSpacer(topRow);
	addDirectionButton(topRow, 'up');
	addSpacer(topRow);

	var middleRow = crossGroup.add('group');
	middleRow.spacing = 2;
	addDirectionButton(middleRow, 'left');
	addSpacer(middleRow);
	addDirectionButton(middleRow, 'right');

	var bottomRow = crossGroup.add('group');
	bottomRow.spacing = 2;
	addSpacer(bottomRow);
	addDirectionButton(bottomRow, 'down');
	addSpacer(bottomRow);

	/* e/d/s/f でその方向へ即時実行（テキスト入力中は無視）/ e/d/s/f run that direction immediately (ignored while typing in a field) */
	win.addEventListener('keydown', function (event) {
		if (event.target && event.target.constructor && event.target.constructor.name === 'EditText') { return; }
		for (var directionKey in directionShortcuts) {
			if (directionShortcuts[directionKey] === event.keyName) {
				executeDirection(directionKey);
				event.preventDefault();
				return;
			}
		}
	});
}

/* オプションパネル（マージン／プレビュー境界）を win に追加し、設定読取り関数を readTransformOptions に公開 / Add the Options panel (margin / preview bounds) to win and publish the reader via readTransformOptions */
function buildOptionsPanel(win) {
	var FIELD_CHARS = 3;
	var optionsPanel = win.add('panel', undefined, L('panel.options'));
	setupPanel(optionsPanel);

	var marginGroup = optionsPanel.add('group');
	marginGroup.spacing = 4; /* ラベルと入力欄の間を詰める（既定spacingが広く「マージン：」の後に空いて見えるため）/ Tighten label-to-field gap (default spacing looks like a gap after "マージン：") */
	marginGroup.add('statictext', undefined, labelText('field.margin'));
	var marginInput = marginGroup.add('edittext', undefined, String(DEFAULT_MARGIN));
	marginInput.characters = FIELD_CHARS;
	var unitLabelText = marginGroup.add('statictext', undefined, getRulerUnitInfo().label);
	changeValueByArrowKey(marginInput, false, null, true); /* 負値を許可（値は実行時に読む）/ allow negatives (read at execution time) */

	var previewBoundsCheck = optionsPanel.add('checkbox', undefined, L('checkbox.previewBounds'));
	previewBoundsCheck.value = false; /* 既定は OFF（明示）/ default off (explicit) */

	/* UI からマージン／プレビュー境界を読み、不正値はここでクランプ / Read margin / preview-bounds from UI; clamp invalid values here */
	function readSettings() {
		var unitInfo = getRulerUnitInfo();
		unitLabelText.text = unitInfo.label; /* 単位表示を更新 / refresh the unit label */
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
/* 常駐パレットを表示（重複起動は排除：既存があれば前面に出して終了）/ Show the persistent palette (no duplicate launch: bring an existing one to front and return) */
function showPalette() {
	if (paletteWindow) {
		try {
			paletteWindow.show();
			paletteWindow.active = true;
			return; /* 既に開いているので作り直さない / already open: do not rebuild */
		} catch (staleReferenceError) {
			paletteWindow = null; /* 参照が無効なら作り直す / stale reference: rebuild */
		}
	}

	/* UI の明暗からアイコンの配色を決定 / Decide icon colors from the light/dark UI */
	initIconColors();

	var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
	win.orientation = 'column';
	win.alignChildren = 'fill';
	win.margins = 15;
	win.spacing = 10;

	/* 1カラム：反転・回転 → 移動・複製 → オプションの順で縦積み / Single column: Flip-Rotate → Move-Duplicate → Options */
	buildFlipPanel(win);
	buildMovePanel(win);
	buildOptionsPanel(win);

	/* Esc で閉じる / Esc closes */
	win.addEventListener('keydown', function (event) {
		if (event.keyName === 'Escape') { win.close(); }
	});
	/* 閉じるとき：参照を解放 / On close: release the reference */
	win.onClose = function () {
		paletteWindow = null;
		return true;
	};

	paletteWindow = win;
	win.layout.layout(true);
	win.show();
}

showPalette();
