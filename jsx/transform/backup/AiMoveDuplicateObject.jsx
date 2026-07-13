#target illustrator
#targetengine "MoveDuplicateObject"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、指定した方向へ複製して移動する常駐パレットです。

- 方向: 上 / 左 / 中央 / 右 / 下 をグラフィカルなボタンで選択（中央＝同じ場所）
- マージン: 元と複製の間に足す余白（単位は定規に追従）
- 繰り返し数: 複製する個数
- 複製: オフのときは元オブジェクトをそのまま移動
- プレビュー境界: 線や効果を含む見た目の境界で移動量を計算

DOM を触る処理（選択取得・移動・複製）はメインエンジンへ BridgeTalk で委譲します。

### 謝辞

kenさん
https://x.com/ken_rainy/status/1472505526768783361

### Overview

A persistent palette that duplicates the selected objects and moves them in a chosen direction.

- Direction: choose Up / Left / Center / Right / Down with graphical buttons (Center = same place)
- Margin: extra gap between original and copy (unit follows the ruler)
- Repeat: number of copies
- Duplicate: when off, moves the original objects as-is
- Preview bounds: uses the visible bounds (including stroke/effects) for the distance

DOM work (selection, move, duplicate) is delegated to the main engine via BridgeTalk.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ユーザー設定 / User Settings
// =========================================
var DEFAULT_DIRECTION = "right"; /* 初期の方向 / Initial direction */
var DEFAULT_DUPLICATE = true;    /* 既定で複製する / Duplicate by default */
var DEFAULT_MARGIN = 0;          /* 元と複製の間に足す余白（定規の単位）/ Extra gap between original and copy (in ruler units) */
var DEFAULT_REPEAT = 1;          /* 繰り返し数（複製する個数）/ Repeat count (number of copies) */

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
		title: { ja: "移動・複製", en: "Move / Duplicate" }
	},
	panel: {
		direction: { ja: "方向", en: "Direction" }
	},
	direction: {
		up:    { ja: "上", en: "Up" },
		left:  { ja: "左", en: "Left" },
		right: { ja: "右", en: "Right" },
		down:  { ja: "下", en: "Down" }
	},
	field: {
		margin: { ja: "マージン", en: "Margin" },
		repeat: { ja: "リピート", en: "Repeat" }
	},
	checkbox: {
		duplicate:     { ja: "複製", en: "Duplicate" },
		previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
	},
	button: {
		apply: { ja: "適用", en: "Apply" }
	},
	tooltip: {
		apply: { ja: "選択オブジェクトに適用（Esc で閉じる）", en: "Apply to the selection (Esc to close)" }
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
/* 直前のプレビューを取り消す（複製は削除、移動は元へ戻す）/ Undo the last preview (remove duplicates, or move originals back) */
function workerClearPreview() {
	var state = $.global.__moveDupPreview;
	if (!state) { return "OK"; };
	if (app.documents.length === 0) { $.global.__moveDupPreview = null; return "NODOC"; };
	try {
		if (state.mode === "dup") {
			for (var i = 0; i < state.items.length; i++) {
				try { state.items[i].remove(); } catch (removeError) {};
			};
		} else {
			for (var j = 0; j < state.items.length; j++) {
				state.items[j].translate(-state.offset[0], -state.offset[1]);
			};
		};
		app.redraw();
	} catch (e) {};
	$.global.__moveDupPreview = null;
	return "OK";
};

/* プレビューを表示（前回を取り消してから再適用し、生成物を記録）/ Show a preview (clear the previous one, re-apply, and record what was created) */
function workerPreview(options) {
	if (app.documents.length === 0) { return "NODOC"; };
	workerClearPreview();
	var docSelection = app.activeDocument.selection;
	if (!docSelection || docSelection.length === 0) { return "NOSEL"; };
	try {
		var selectionSize = workerSelectionSize(docSelection, options.usePreviewBounds);
		var offset = workerOffset(options.direction, selectionSize.width, selectionSize.height, options.marginPt);
		$.global.__moveDupPreview = workerBuild(docSelection, offset, options.repeat, options.duplicate);
		app.redraw();
		return "OK";
	} catch (e) {
		return "ERR:" + e;
	};
};

/* プレビューを確定（記録を破棄して生成物を残す）/ Commit the preview (drop the record, keep what was created) */
function workerCommit() {
	$.global.__moveDupPreview = null;
	return "OK";
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

/* 複製（またはそのまま）を移動し、取り消し用の状態を返す / Move duplicates (or originals) and return the state needed to undo it */
function workerBuild(items, offset, repeatCount, shouldDuplicate) {
	var createdItems = [];
	if (shouldDuplicate) {
		for (var step = 1; step <= repeatCount; step++) {
			for (var i = 0; i < items.length; i++) {
				var duplicatedItem = items[i].duplicate();
				duplicatedItem.translate(offset[0] * step, offset[1] * step);
				createdItems.push(duplicatedItem);
			};
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
var WORKER_FUNCTIONS = [workerPreview, workerCommit, workerClearPreview, workerBuild, workerSelectionSize, workerOffset];

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
		"repeat:" + options.repeat + "," +
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

/* プレビュー表示を委譲 / Delegate showing the preview */
function delegatePreview(options) {
	return delegateToMain("workerPreview(" + optionsLiteral(options) + ");");
}

/* プレビュー確定を委譲 / Delegate committing the preview */
function delegateCommit() {
	return delegateToMain("workerCommit();");
}

/* プレビュー取り消しを委譲 / Delegate clearing the preview */
function delegateCancel() {
	return delegateToMain("workerClearPreview();");
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

/* 矩形パスを作る（rectPath の引数順を避けて moveTo/lineTo で構築）/ Build a rectangle path via moveTo/lineTo */
function pathRect(graphics, x, y, w, h) {
	graphics.newPath();
	graphics.moveTo(x, y);
	graphics.lineTo(x + w, y);
	graphics.lineTo(x + w, y + h);
	graphics.lineTo(x, y + h);
	graphics.closePath();
}

/* 右向き矢印の各点を方向キーに合わせて回転 / Rotate a right-arrow point to match the direction key */
function transformArrowPoint(directionKey, x, y) {
	if (directionKey === 'left') { return [-x, y]; }
	if (directionKey === 'down') { return [-y, x]; }
	if (directionKey === 'up')   { return [y, -x]; }
	return [x, y]; /* right */
}

/* 方向キーの矢印を塗りで描画（pressOffset で押し込み時に少しずらす）/ Draw a filled arrow (pressOffset nudges it slightly when pressed) */
function drawArrow(graphics, directionKey, w, h, color, pressOffset) {
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
	var cx = w / 2 + (pressOffset || 0), cy = h / 2 + (pressOffset || 0);
	graphics.newPath();
	for (var i = 0; i < basePoints.length; i++) {
		var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
		if (i === 0) { graphics.moveTo(cx + point[0], cy + point[1]); }
		else { graphics.lineTo(cx + point[0], cy + point[1]); }
	}
	graphics.closePath();
	graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
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

	var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
	win.orientation = 'column';
	win.alignChildren = 'fill';
	win.margins = 15;
	win.spacing = 10;

	/* --- 方向パネル（3x3、十字位置にグラフィカルボタン、四隅は空）/ Direction panel (3x3, graphical buttons on the cross, empty corners) --- */
	var directionPanel = win.add('panel', undefined, L('panel.direction'));
	directionPanel.orientation = 'column';
	directionPanel.alignChildren = 'center';
	directionPanel.margins = 15;
	directionPanel.spacing = 2;

	var uiColors = isLightUI()
		? { background: [0.92, 0.92, 0.92, 1], foreground: [0.20, 0.20, 0.20, 1] }
		: { background: [0.26, 0.26, 0.26, 1], foreground: [0.85, 0.85, 0.85, 1] };
	uiColors.selectedBackground = [0.20, 0.50, 0.90, 1];
	uiColors.selectedForeground = [1, 1, 1, 1];
	uiColors.pressedBorder = [0.12, 0.34, 0.68, 1]; /* 押し込み枠（濃いアクセント）/ pressed border (darker accent) */

	var CELL_SIZE = 30; /* 方向ボタン1セルの大きさ（参照: AiQuickPrefsPalette-simple の 26px 前後）/ Direction cell size (ref: ~26px in AiQuickPrefsPalette-simple) */
	var selectedDirection = DEFAULT_DIRECTION;
	var directionButtons = [];
	var previewActive = false; /* 未確定のプレビューが表示中か / Whether an uncommitted preview is showing */
	/* 方向選択のショートカット（上:e / 下:d / 左:s / 右:f）/ Direction shortcuts (up:e / down:d / left:s / right:f) */
	var directionShortcuts = { up: 'E', down: 'D', left: 'S', right: 'F' };

	/* ボタン1つを onDraw で描画（選択中は押し込み風）/ Draw one button via onDraw (pressed look when selected) */
	function drawDirectionButton(control) {
		var graphics = control.graphics;
		var w = control.size[0], h = control.size[1];
		var isSelected = (control.directionKey === selectedDirection);

		/* 背景 / background */
		pathRect(graphics, 0, 0, w, h);
		graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isSelected ? uiColors.selectedBackground : uiColors.background));

		/* 選択中はインセット枠で押し込み表現 / Inset border to look pressed when selected */
		if (isSelected) {
			pathRect(graphics, 1.5, 1.5, w - 3, h - 3);
			graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, uiColors.pressedBorder, 1.5));
		}

		/* 矢印（押し込み中は1px ずらす）/ Arrow (nudged 1px while pressed) */
		var markColor = isSelected ? uiColors.selectedForeground : uiColors.foreground;
		drawArrow(graphics, control.directionKey, w, h, markColor, isSelected ? 1 : 0);
	}

	/* 指定キーだけ選択し、全ボタンを再描画して他を OFF 表示に、プレビューも更新 / Select the key, repaint all buttons, and refresh the preview */
	function selectDirection(directionKey) {
		selectedDirection = directionKey;
		for (var i = 0; i < directionButtons.length; i++) {
			/* hide()/show() で確実に onDraw を再実行 / hide()/show() reliably re-runs onDraw */
			directionButtons[i].hide();
			directionButtons[i].show();
		}
		requestPreview();
	}

	/* 空セルを追加 / Add an empty cell */
	function addSpacer(row) {
		var spacer = row.add('statictext', undefined, '');
		spacer.preferredSize = [CELL_SIZE, CELL_SIZE];
	}

	/* 方向ボタンを追加 / Add a direction button */
	function addDirectionButton(row, directionKey) {
		var button = row.add('iconbutton', undefined, undefined, { style: 'toolbutton' });
		button.preferredSize = [CELL_SIZE, CELL_SIZE];
		button.directionKey = directionKey;
		button.helpTip = L('direction.' + directionKey) + (directionShortcuts[directionKey] ? '  [' + directionShortcuts[directionKey] + ']' : '');
		button.onDraw = function () { drawDirectionButton(this); };
		button.onClick = function () { selectDirection(this.directionKey); };
		directionButtons.push(button);
	}

	var topRow = directionPanel.add('group');
	topRow.spacing = 2;
	addSpacer(topRow);
	addDirectionButton(topRow, 'up');
	addSpacer(topRow);

	var middleRow = directionPanel.add('group');
	middleRow.spacing = 2;
	addDirectionButton(middleRow, 'left');
	addSpacer(middleRow);
	addDirectionButton(middleRow, 'right');

	var bottomRow = directionPanel.add('group');
	bottomRow.spacing = 2;
	addSpacer(bottomRow);
	addDirectionButton(bottomRow, 'down');
	addSpacer(bottomRow);

	/* --- 数値入力（入力欄は3文字幅）/ Numeric inputs (3-char fields) --- */
	var FIELD_CHARS = 3;

	var repeatGroup = win.add('group');
	repeatGroup.add('statictext', undefined, labelText('field.repeat'));
	var repeatInput = repeatGroup.add('edittext', undefined, String(DEFAULT_REPEAT));
	repeatInput.characters = FIELD_CHARS;
	changeValueByArrowKey(repeatInput, true, requestPreview); /* リピートは整数専用 / Repeat count is integer-only */
	repeatInput.onChange = requestPreview;

	var marginGroup = win.add('group');
	marginGroup.add('statictext', undefined, labelText('field.margin'));
	var marginInput = marginGroup.add('edittext', undefined, String(DEFAULT_MARGIN));
	marginInput.characters = FIELD_CHARS;
	var unitLabelText = marginGroup.add('statictext', undefined, getRulerUnitInfo().label);
	changeValueByArrowKey(marginInput, false, requestPreview, true); /* 負値を許可 / allow negatives */
	marginInput.onChange = requestPreview; /* 手入力の確定でプレビュー更新 / Update preview when typed value is committed */

	/* --- オプション / Options --- */
	var optionGroup = win.add('group');
	optionGroup.orientation = 'column';
	optionGroup.alignChildren = 'left';
	var duplicateCheck = optionGroup.add('checkbox', undefined, L('checkbox.duplicate'));
	var previewBoundsCheck = optionGroup.add('checkbox', undefined, L('checkbox.previewBounds'));
	duplicateCheck.value = DEFAULT_DUPLICATE;
	previewBoundsCheck.value = false; /* 既定は OFF（明示）/ default off (explicit) */
	duplicateCheck.onClick = requestPreview;
	previewBoundsCheck.onClick = requestPreview;

	/* --- 適用ボタン（閉じるは ×/Esc に任せる）/ Apply button (close via × / Esc) --- */
	var buttonGroup = win.add('group');
	buttonGroup.alignment = 'right';
	var applyButton = buttonGroup.add('button', undefined, L('button.apply'));
	applyButton.helpTip = L('tooltip.apply');

	/* UI から設定を読み、負数・不正値はここでクランプ / Read settings from UI; clamp negatives and invalid values here */
	function readOptions() {
		var unitInfo = getRulerUnitInfo();
		unitLabelText.text = unitInfo.label; /* 単位表示を更新 / refresh the unit label */

		var marginValue = parseFloat(marginInput.text);
		if (isNaN(marginValue)) { marginValue = 0; } /* 負値は許容（マイナスで重なり方向へ）/ Negatives allowed (moves toward overlap) */
		marginInput.text = String(marginValue);

		var repeatValue = parseInt(repeatInput.text, 10);
		if (isNaN(repeatValue) || repeatValue < 1) { repeatValue = 1; }
		repeatInput.text = String(repeatValue);

		return {
			direction: selectedDirection,
			marginPt: marginValue * unitInfo.factor,
			repeat: repeatValue,
			duplicate: duplicateCheck.value,
			usePreviewBounds: previewBoundsCheck.value
		};
	}

	/* プレビューを更新（再入防止つき、成功時のみ有効フラグを立てる）/ Refresh the preview (re-entrancy guard; mark active only on success) */
	function requestPreview() {
		if (isBusy) { return; }
		isBusy = true;
		try {
			previewActive = (delegatePreview(readOptions()) === "OK");
		} finally {
			isBusy = false;
		}
	}

	/* 適用：最新設定でプレビューし直してから確定 / Apply: refresh the preview with the latest settings, then commit */
	function onApply() {
		if (isBusy) { return; }
		isBusy = true;
		try {
			delegatePreview(readOptions());
			delegateCommit();
			previewActive = false;
		} finally {
			isBusy = false;
		}
	}
	applyButton.onClick = onApply;

	/* Esc で閉じる、e/d/s/f で方向を選択（テキスト入力中は方向キーを無視）/ Esc closes; e/d/s/f pick a direction (ignored while typing in a field) */
	win.addEventListener('keydown', function (event) {
		if (event.keyName === 'Escape') { win.close(); return; }
		if (event.target && event.target.constructor && event.target.constructor.name === 'EditText') { return; }
		for (var directionKey in directionShortcuts) {
			if (directionShortcuts[directionKey] === event.keyName) {
				selectDirection(directionKey);
				event.preventDefault();
				return;
			}
		}
	});
	/* 閉じるとき：未確定プレビューは取り消し、参照を解放 / On close: cancel any uncommitted preview and release the reference */
	win.onClose = function () {
		if (previewActive) {
			delegateCancel();
			previewActive = false;
		}
		paletteWindow = null;
		return true;
	};

	paletteWindow = win;
	win.layout.layout(true);
	win.show();

	/* 選択があれば開いた直後にプレビュー / Preview right after opening if something is selected */
	requestPreview();
}

showPalette();
