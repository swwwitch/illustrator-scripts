#target illustrator
#targetengine "fxConvertToShape"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトに［形状に変換］のライブエフェクトを適用する常駐パレット。パレットで「長方形／楕円」と「値を指定（Absolute）／値を追加（Relative）」、幅・高さ（pt）を設定すると、選択にライブプレビューが反映される。DOM 操作は BridgeTalk でメインエンジンへ委譲する。

### Overview

A persistent palette that applies the "Convert to Shape" live effect to the selection. Pick Rectangle/Ellipse and Absolute/Relative sizing plus width/height (pt); the selection updates as a live preview. All DOM work is delegated to the main engine via BridgeTalk.

### 更新履歴 / Changelog

- v1.0.0: 常駐パレット化。形状（長方形／楕円）とサイズモード（値を指定／値を追加）＋幅・高さを設定し、ライブプレビュー付きで［形状に変換］を適用 ／ Persistent palette. Choose shape (rectangle/ellipse) and size mode (absolute/relative) plus width/height, apply "Convert to Shape" with a live preview.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ローカライズ / Localization
// =========================================

/* 表示言語を判定 / Detect display language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* 日英ラベル定義（カテゴリ構造）/ Japanese-English label definitions (categorized) */
var LABELS = {
    dialog: {
        title: { ja: "形状に変換", en: "Convert to Shape" }
    },
    panel: {
        main: { ja: "形状に変換", en: "Convert to Shape" },
        shape: { ja: "形状", en: "Shape" },
        options: { ja: "オプション", en: "Options" }
    },
    shape: {
        rectangle: { ja: "長方形", en: "Rectangle" },
        ellipse: { ja: "楕円", en: "Ellipse" }
    },
    option: {
        size: { ja: "サイズ", en: "Size" },
        absolute: { ja: "値を指定", en: "Absolute" },
        relative: { ja: "値を追加", en: "Relative" },
        width: { ja: "幅:", en: "Width:" },
        height: { ja: "高さ:", en: "Height:" }
    },
    pathfinder: {
        label: { ja: "パスファインダー", en: "Pathfinder" },
        none: { ja: "なし", en: "None" },
        add: { ja: "合体", en: "Unite" },
        intersect: { ja: "交差", en: "Intersect" },
        exclude: { ja: "中マド", en: "Exclude" },
        minusFront: { ja: "前面オブジェクトで型抜き", en: "Minus Front" },
        minusBack: { ja: "背面オブジェクトで型抜き", en: "Minus Back" },
        divide: { ja: "分割", en: "Divide" },
        trim: { ja: "刈り込み", en: "Trim" },
        merge: { ja: "合流", en: "Merge" },
        crop: { ja: "切り抜き", en: "Crop" },
        outline: { ja: "アウトライン", en: "Outline" }
    },
    button: {
        apply: { ja: "適用", en: "Apply" }
    },
    tip: {
        apply: { ja: "効果を確定します（Esc で閉じる）", en: "Commit the effect (Esc to close)" }
    },
    status: {
        ready: { ja: "オブジェクトを選択して設定してください。", en: "Select objects and adjust settings." },
        applied: { ja: "適用しました", en: "Applied" },
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
        error: { ja: "エラーが発生しました。", en: "An error occurred." }
    }
};

/* ドット区切りキーから LABELS のエントリを取得 / Resolve a LABELS entry from a dot-separated key */
function getLabelEntry(key) {
    var parts = key.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node && typeof node[parts[i]] !== "undefined") { node = node[parts[i]]; }
        else { return null; }
    }
    return node;
}

/* キーからローカライズ文字列を取得 / Get a localized string by key */
function L(key) {
    var entry = getLabelEntry(key);
    if (entry) {
        if (entry[currentLanguage]) return entry[currentLanguage];
        if (entry.en) return entry.en;
    }
    return key;
}

// =========================================
// 形状パラメータ（コントローラ側）/ Shape parameters (controller side)
// =========================================

/* 形状ごとのライブエフェクト定義パラメータ / Live-effect parameters per shape */
var SHAPES = {
    rectangle: { displayString: "Rectangle", shape: 0 },
    ellipse: { displayString: "Ellipse", shape: 2 }
};

/* パスファインダーの適用方式（none は適用なし）
   menu … app.executeMenuCommand（選択全体へ1回）／ effect … applyEffect（各オブジェクトへXMLで）
   分割・アウトラインのみ effect 方式、それ以外は menu 方式
   Pathfinder application method (none = no effect):
   menu … app.executeMenuCommand (once, whole selection) / effect … applyEffect (per object, via XML).
   Divide & Outline use the effect method; the rest use menu. */
var PATHFINDERS = {
    none: { menu: "", effect: "" },
    add: { menu: "Live Pathfinder Add", effect: "" },
    intersect: { menu: "Live Pathfinder Intersect", effect: "" },
    exclude: { menu: "Live Pathfinder Exclude", effect: "" },
    minusFront: { menu: "Live Pathfinder Subtract", effect: "" },
    minusBack: { menu: "Live Pathfinder Minus Back", effect: "" },
    divide: { menu: "", effect: "Adobe Pathfinder Divide" },
    trim: { menu: "Live Pathfinder Trim", effect: "" },
    merge: { menu: "Live Pathfinder Merge", effect: "" },
    crop: { menu: "Live Pathfinder Crop", effect: "" },
    outline: { menu: "", effect: "Adobe Pathfinder Outline" }
};

/* ラジオ生成順のパスファインダーキー / Pathfinder keys in radio order */
var PATHFINDER_KEYS = ["none", "add", "intersect", "exclude", "minusFront", "minusBack", "divide", "trim", "merge", "crop", "outline"];

// =========================================
// ワーカー関数（メインエンジンで実行）/ Worker functions (run on the main engine)
// 委譲先で eval されるため、// 行コメント禁止・/* */ のみ・必ずセミコロンで終える
// Delegated & eval'd on the main engine: no // comments, use /* */ only, always end with a semicolon
// =========================================

/* ［形状に変換］のライブエフェクト定義を組み立て（幅・高さは pt）/ Build the "Convert to Shape" live-effect XML (width/height in pt) */
function fxBuildShapeXML(shapeNum, displayString, absolute, width, height) {
    var relW = absolute ? 0 : width;
    var relH = absolute ? 0 : height;
    var absW = absolute ? width : 0;
    var absH = absolute ? height : 0;
    return '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString ' + displayString + ' I Shape ' + shapeNum + ' R RelWidth ' + relW + ' R RelHeight ' + relH + ' R AbsWidth ' + absW + ' R AbsHeight ' + absH + ' R Absolute ' + (absolute ? 1 : 0) + ' R CornerRadius 9 "/></LiveEffect>';
}

/* パスファインダーのライブエフェクト定義を組み立て / Build a pathfinder live-effect XML */
function fxBuildPathfinderXML(effectName) {
    return '<LiveEffect name="' + effectName + '"/>';
}

/* 前回プレビュー分を undo してから選択にシェイプ効果（＋任意のパスファインダー）を適用。結果は "OK:undo回数"
   シェイプ効果は各オブジェクトへ applyEffect。パスファインダーは effect 指定なら各オブジェクトへ applyEffect、
   command 指定なら選択全体へ executeMenuCommand で1回適用する（分割・アウトラインのみ effect 方式）
   Undo the previous preview, then apply the shape effect (plus an optional pathfinder). Returns "OK:undoSteps".
   Shape effect: per object via applyEffect. Pathfinder: per object via applyEffect when an effect name is given,
   otherwise once on the whole selection via executeMenuCommand (Divide & Outline use the effect method). */
function fxRunShapeEffect(shapeNum, displayString, absolute, width, height, pathfinderCommand, pathfinderEffect, undoCount) {
    if (app.documents.length === 0) { return "NODOC"; }
    var u = 0;
    for (u = 0; u < undoCount; u++) { app.undo(); }
    if (undoCount > 0) { app.redraw(); }
    var doc = app.activeDocument;
    var sel = doc.selection;
    if (!sel || sel.length === 0) { return "NOSEL"; }
    var xml = fxBuildShapeXML(shapeNum, displayString, absolute, width, height);
    var pfXml = pathfinderEffect ? fxBuildPathfinderXML(pathfinderEffect) : "";
    var steps = 0;
    var i = 0;
    for (i = 0; i < sel.length; i++) {
        try { sel[i].applyEffect(xml); steps = steps + 1; } catch (e) { }
        if (pfXml) {
            try { sel[i].applyEffect(pfXml); steps = steps + 1; } catch (e2) { }
        }
    }
    if (!pfXml && pathfinderCommand) {
        try { app.executeMenuCommand(pathfinderCommand); steps = steps + 1; } catch (e3) { }
    }
    app.redraw();
    return "OK:" + steps;
}

/* 指定回数だけ undo（プレビュー取り消し用）。ドキュメント確認を先に置く / Undo N times (to cancel a preview); check the document first */
function fxUndo(undoCount) {
    if (app.documents.length === 0) { return "NODOC"; }
    var i = 0;
    for (i = 0; i < undoCount; i++) { app.undo(); }
    app.redraw();
    return "OK";
}

/* 委譲するワーカー関数（追加漏れ防止のため全登録）/ Worker functions to delegate (register all to avoid omissions) */
var WORKER_FUNCS = [fxBuildShapeXML, fxBuildPathfinderXML, fxRunShapeEffect, fxUndo];

// =========================================
// BridgeTalk 委譲 / BridgeTalk delegation
// =========================================

/* 再入防止フラグ / Reentrancy guard */
var isBusy = false;

/* ワーカー関数群＋呼び出し式をメインエンジンへ同期送信し、マーカー文字列を返す
   Send the worker functions + a call expression to the main engine synchronously; return the marker string */
function callWorker(callExpression) {
    var source = "";
    for (var i = 0; i < WORKER_FUNCS.length; i++) {
        source += WORKER_FUNCS[i].toString() + "\n";
    }
    source += callExpression + ";";

    var encoded = encodeURIComponent(source);
    var holder = { result: "ERR:notrun" };

    var bridge = new BridgeTalk();
    bridge.target = "illustrator";
    bridge.body = 'eval(decodeURIComponent("' + encoded + '"));';
    bridge.onResult = function (resObj) { holder.result = String(resObj.body); };
    bridge.onError = function (errObj) { holder.result = "ERR:" + String(errObj.body); };
    bridge.send(10);

    return holder.result;
}

// =========================================
// オプション取得 / Options
// =========================================

/* コントロールの参照（showPalette で設定）/ Control references (populated in showPalette) */
var controls = {};

/* パレットUIから設定値を取得。手入力の負数・不正値はクランプ / Read settings from the palette; clamp negatives/invalid input */
function readOptions() {
    var width = parseFloat(controls.widthInput.text);
    var height = parseFloat(controls.heightInput.text);
    if (isNaN(width) || width < 0) width = 0;
    if (isNaN(height) || height < 0) height = 0;
    controls.widthInput.text = String(width);   // クランプ結果を反映 / Reflect the clamped value
    controls.heightInput.text = String(height);

    // 選択中のパスファインダーキーを取得（既定 none）/ Selected pathfinder key (default none)
    var pathfinderKey = "none";
    for (var i = 0; i < controls.pathfinderRadios.length; i++) {
        if (controls.pathfinderRadios[i].value) { pathfinderKey = controls.pathfinderRadios[i].pfKey; break; }
    }

    return {
        shapeKey: controls.rectangleRadio.value ? "rectangle" : "ellipse",
        absolute: controls.absoluteRadio.value,
        width: width,
        height: height,
        pathfinderKey: pathfinderKey
    };
}

// =========================================
// 適用 / プレビュー / 取り消し / Apply / Preview / Undo
// =========================================

/* プレビュー状態 / Preview state */
var previewActive = false;   // 未確定のプレビューが乗っているか / Whether an uncommitted preview is applied
var previewCount = 0;        // 直近プレビューで適用した効果数（undo 回数）/ Effects applied last (undo count)

/* 状況表示を更新 / Update the status line */
function setStatus(message) {
    if (controls.status) { controls.status.text = message; }
}

/* ワーカー結果（マーカー）をローカライズして status に反映 / Localize the worker marker and show it in the status */
function handleResult(result) {
    if (result.indexOf("OK") === 0) {
        var applied = parseInt(result.split(":")[1], 10);
        if (isNaN(applied)) applied = 0;
        previewActive = applied > 0;
        previewCount = applied;
        setStatus(L("status.applied") + " (" + applied + ")");
    } else if (result === "NODOC") {
        previewActive = false; previewCount = 0;
        setStatus(L("status.noDocument"));
    } else if (result.indexOf("NOSEL") === 0) {
        previewActive = false; previewCount = 0;
        setStatus(L("status.noSelection"));
    } else {
        previewActive = false; previewCount = 0;
        setStatus(L("status.error") + " " + result);
    }
}

/* 現在の設定でプレビュー適用（前回プレビューは undo してから再適用）/ Apply a preview with the current settings (undo the previous preview first) */
function runPreview() {
    if (isBusy) return;
    isBusy = true;
    try {
        var options = readOptions();
        var shape = SHAPES[options.shapeKey] || SHAPES.rectangle;
        var pathfinder = PATHFINDERS[options.pathfinderKey] || PATHFINDERS.none;
        var factor = controls.unitFactor || 1.0;
        var undoCount = previewActive ? previewCount : 0;
        // 表示単位の入力値を pt に換算して委譲 / Convert the entered unit values to pt before delegating
        var call = "fxRunShapeEffect(" +
            shape.shape + ",\"" + shape.displayString + "\"," +
            (options.absolute ? "true" : "false") + "," +
            (options.width * factor) + "," + (options.height * factor) + ",\"" +
            pathfinder.menu + "\",\"" + pathfinder.effect + "\"," + undoCount + ")";
        handleResult(callWorker(call));
    } finally {
        isBusy = false;
    }
}

/* ［適用］：現在のプレビューを確定（以後 onClose で取り消さない）/ Apply: commit the current preview (onClose won't undo it afterwards) */
function commitApply() {
    runPreview();
    previewActive = false; // 確定 / committed
    previewCount = 0;
}

/* 未確定プレビューを取り消す（× / Esc で閉じるとき）/ Undo the uncommitted preview (when closing via × / Esc) */
function undoLastPreview() {
    if (isBusy) return;
    isBusy = true;
    try {
        callWorker("fxUndo(" + previewCount + ")");
        previewActive = false;
        previewCount = 0;
    } finally {
        isBusy = false;
    }
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
// UI ヘルパー / UI helpers
// =========================================

/* ↑↓キーで値を増減（Shift=±10 かつ10の倍数へスナップ／Option=±0.1／単独=±1）
   修飾キーは event 優先＋keyboardState フォールバック（macOS の altKey 誤報対策）
   Increment/decrement the value with ↑/↓ (Shift = ±10 snapped to multiples of 10 / Option = ±0.1 / plain = ±1).
   Read modifiers from event first, falling back to keyboardState (works around the macOS altKey misreport). */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var shiftKey = (typeof event.shiftKey === "boolean") ? event.shiftKey : keyboard.shiftKey;
        var altKey = (typeof event.altKey === "boolean") ? event.altKey : keyboard.altKey;
        var delta = 1;

        if (shiftKey) {
            delta = 10;
            /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
            if (event.keyName === "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (altKey) {
            delta = 0.1;
            /* Optionキー押下時は0.1単位で増減 / Step by 0.1 when Option is held */
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
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (altKey) {
            /* 小数第1位までに丸め / Round to one decimal place */
            value = Math.round(value * 10) / 10;
        } else {
            /* 整数に丸め / Round to an integer */
            value = Math.round(value);
        }

        editText.text = value;
        if (editText.onChange) editText.onChange(); // ライブプレビューを更新 / Refresh the live preview
    });
}

// =========================================
// パレット / Palette
// =========================================

/* 常駐パレット参照キー（$.global に保持。スクリプト再実行で top-level var が初期化されても参照を失わない）
   Persistent palette reference key (kept on $.global so a re-run's top-level var reset can't drop the reference) */
var PALETTE_GLOBAL_KEY = "__fxConvertToShapePalette";

/* パレットを表示（多重起動防止：既存があれば閉じる）/ Show the palette (prevent duplicates: close any existing one) */
function showPalette() {
    // 二重起動防止：既存パレットを閉じてから作り直す / Prevent double launch: close any existing palette first
    if ($.global[PALETTE_GLOBAL_KEY]) {
        try { $.global[PALETTE_GLOBAL_KEY].close(); } catch (e0) { }
        $.global[PALETTE_GLOBAL_KEY] = null;
    }

    var win = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = 15;

    // 「形状に変換」パネルにすべてまとめる / Wrap everything in the "Convert to Shape" panel
    var mainPanel = win.add("panel", undefined, L("panel.main"));
    mainPanel.orientation = "column";
    mainPanel.alignChildren = "fill";
    mainPanel.margins = [16, 20, 16, 12];
    mainPanel.spacing = 8;

    // 形状：長方形／楕円（パネルは使わずラベル＋横並びラジオ）/ Shape: rectangle / ellipse (label + inline radios, no panel)
    var shapeGroup = mainPanel.add("group");
    shapeGroup.orientation = "row";
    shapeGroup.spacing = 12;
    shapeGroup.add("statictext", undefined, L("panel.shape"));
    var rectangleRadio = shapeGroup.add("radiobutton", undefined, L("shape.rectangle"));
    var ellipseRadio = shapeGroup.add("radiobutton", undefined, L("shape.ellipse"));
    rectangleRadio.value = true; // 既定は長方形 / Default: rectangle

    // オプション：値を指定／値を追加＋幅・高さ（パネルは使わずグループ）/ Options: Absolute / Relative + width / height (group, no panel)
    var sizePanel = mainPanel.add("group");
    sizePanel.orientation = "column";
    sizePanel.alignChildren = ["left", "top"];
    sizePanel.spacing = 6;
    var sizeRadioGroup = sizePanel.add("group");
    sizeRadioGroup.orientation = "row";
    sizeRadioGroup.spacing = 12;
    sizeRadioGroup.add("statictext", undefined, L("option.size"));
    var absoluteRadio = sizeRadioGroup.add("radiobutton", undefined, L("option.absolute"));
    var relativeRadio = sizeRadioGroup.add("radiobutton", undefined, L("option.relative"));
    relativeRadio.value = true; // 既定は値を追加 / Default: relative

    // 定規の単位を取得（ラベル表示＋pt換算係数）/ Get the ruler unit (label + pt factor)
    var unitInfo = getRulerUnitInfo();
    controls.unitFactor = unitInfo.factor;

    var LABEL_WIDTH = 40; // 幅・高さラベルの固定幅 / Fixed width for the width/height labels

    var widthGroup = sizePanel.add("group");
    var widthLabel = widthGroup.add("statictext", undefined, L("option.width"), { justify: "right" });
    widthLabel.preferredSize.width = LABEL_WIDTH;
    var widthInput = widthGroup.add("edittext", undefined, "0");
    widthInput.characters = 3;
    widthGroup.add("statictext", undefined, unitInfo.label);
    var heightGroup = sizePanel.add("group");
    var heightLabel = heightGroup.add("statictext", undefined, L("option.height"), { justify: "right" });
    heightLabel.preferredSize.width = LABEL_WIDTH;
    var heightInput = heightGroup.add("edittext", undefined, "0");
    heightInput.characters = 3;
    heightGroup.add("statictext", undefined, unitInfo.label);

    // ↑↓（Shift=±10／Option=±0.1）で増減 / Increment with ↑/↓ (Shift = ±10 / Option = ±0.1)
    changeValueByArrowKey(widthInput);
    changeValueByArrowKey(heightInput);

    // パスファインダー（パネル＋2カラムのラジオ）/ Pathfinder (panel + two columns of radios)
    var pathfinderPanel = mainPanel.add("panel", undefined, L("pathfinder.label"));
    pathfinderPanel.orientation = "row";
    pathfinderPanel.alignChildren = ["left", "top"];
    pathfinderPanel.margins = [16, 20, 16, 12];
    pathfinderPanel.spacing = 16;
    var pfColumnLeft = pathfinderPanel.add("group");
    pfColumnLeft.orientation = "column";
    pfColumnLeft.alignChildren = ["left", "top"];
    pfColumnLeft.spacing = 4;
    var pfColumnRight = pathfinderPanel.add("group");
    pfColumnRight.orientation = "column";
    pfColumnRight.alignChildren = ["left", "top"];
    pfColumnRight.spacing = 4;

    controls.pathfinderRadios = [];
    var pfHalf = Math.ceil(PATHFINDER_KEYS.length / 2);
    for (var p = 0; p < PATHFINDER_KEYS.length; p++) {
        var pfKey = PATHFINDER_KEYS[p];
        var pfColumn = (p < pfHalf) ? pfColumnLeft : pfColumnRight;
        var pfRadio = pfColumn.add("radiobutton", undefined, L("pathfinder." + pfKey));
        pfRadio.pfKey = pfKey;
        pfRadio.onClick = runPreview;
        controls.pathfinderRadios.push(pfRadio);
    }
    controls.pathfinderRadios[0].value = true; // 既定は「なし」/ Default: none

    // ［適用］（閉じるは × / Esc）/ Apply button (close via × / Esc)
    var buttonGroup = mainPanel.add("group");
    buttonGroup.alignment = "center";
    var applyButton = buttonGroup.add("button", undefined, L("button.apply"), { name: "ok" });
    applyButton.helpTip = L("tip.apply");

    // 状況表示（下部）/ Status line (bottom)
    var status = win.add("statictext", undefined, L("status.ready"));
    status.alignment = "fill";

    // 参照を保持 / Keep references
    controls.rectangleRadio = rectangleRadio;
    controls.ellipseRadio = ellipseRadio;
    controls.absoluteRadio = absoluteRadio;
    controls.relativeRadio = relativeRadio;
    controls.widthInput = widthInput;
    controls.heightInput = heightInput;
    controls.status = status;

    // 設定変更のたびにライブプレビュー / Live preview on each change
    rectangleRadio.onClick = runPreview;
    ellipseRadio.onClick = runPreview;
    absoluteRadio.onClick = runPreview;
    relativeRadio.onClick = runPreview;
    widthInput.onChange = runPreview;
    heightInput.onChange = runPreview;

    // ［適用］で確定 / Commit on Apply
    applyButton.onClick = commitApply;

    // Esc で閉じる / Close on Esc
    win.addEventListener("keydown", function (keyEvent) {
        if (keyEvent.keyName === "Escape") { win.close(); }
    });

    // 未確定のまま閉じたらプレビューを取り消す / Undo the preview if closed while uncommitted
    win.onClose = function () {
        if (previewActive) { undoLastPreview(); }
        return true;
    };

    $.global[PALETTE_GLOBAL_KEY] = win;
    win.center();
    win.show();
}

showPalette();
