#target illustrator
#targetengine "fxConvertToShape"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトに［形状に変換］のライブエフェクトを適用する常駐パレットです。
「長方形／楕円」と「値を指定／値を追加」、幅・高さを設定すると、選択にライブプレビューが反映されます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LEConvertToShapePalette.md

### Overview

A persistent palette that applies the "Convert to Shape" live effect to the selection.
Pick Rectangle or Ellipse, Absolute or Relative sizing, and the width and height; the selection updates as a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LEConvertToShapePalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LEConvertToShapePalette";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LEConvertToShapePalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LEConvertToShapePalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
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
        tooltip: {
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
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) { return "NOSEL"; }
        var xml = fxBuildShapeXML(shapeNum, displayString, absolute, width, height);
        var pfXml = pathfinderEffect ? fxBuildPathfinderXML(pathfinderEffect) : "";
        var steps = 0;
        var i = 0;
        for (i = 0; i < currentSelection.length; i++) {
            try { currentSelection[i].applyEffect(xml); steps = steps + 1; } catch (e) { }
            if (pfXml) {
                try { currentSelection[i].applyEffect(pfXml); steps = steps + 1; } catch (e2) { }
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
            setStatus(getLabel("status.applied") + " (" + applied + ")");
        } else if (result === "NODOC") {
            previewActive = false; previewCount = 0;
            setStatus(getLabel("status.noDocument"));
        } else if (result.indexOf("NOSEL") === 0) {
            previewActive = false; previewCount = 0;
            setStatus(getLabel("status.noSelection"));
        } else {
            previewActive = false; previewCount = 0;
            setStatus(getLabel("status.error") + " " + result);
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

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // UI ヘルパー / UI helpers
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

        var win = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        setupWindow(win);

        // 「形状に変換」パネルにすべてまとめる / Wrap everything in the "Convert to Shape" panel
        var mainPanel = win.add("panel", undefined, getLabel("panel.main"));
        setupPanel(mainPanel, 6);

        // 形状：長方形／楕円（パネルは使わずラベル＋横並びラジオ）/ Shape: rectangle / ellipse (label + inline radios, no panel)
        var shapeGroup = mainPanel.add("group");
        shapeGroup.orientation = "row";
        shapeGroup.spacing = 12;
        shapeGroup.add("statictext", undefined, getLabel("panel.shape"));
        var rectangleRadio = shapeGroup.add("radiobutton", undefined, getLabel("shape.rectangle"));
        var ellipseRadio = shapeGroup.add("radiobutton", undefined, getLabel("shape.ellipse"));
        rectangleRadio.value = true; // 既定は長方形 / Default: rectangle

        // オプション：値を指定／値を追加＋幅・高さ（パネルは使わずグループ）/ Options: Absolute / Relative + width / height (group, no panel)
        var sizePanel = mainPanel.add("group");
        sizePanel.orientation = "column";
        sizePanel.alignChildren = ["left", "top"];
        sizePanel.spacing = 6;
        var sizeRadioGroup = sizePanel.add("group");
        sizeRadioGroup.orientation = "row";
        sizeRadioGroup.spacing = 12;
        sizeRadioGroup.add("statictext", undefined, getLabel("option.size"));
        var absoluteRadio = sizeRadioGroup.add("radiobutton", undefined, getLabel("option.absolute"));
        var relativeRadio = sizeRadioGroup.add("radiobutton", undefined, getLabel("option.relative"));
        relativeRadio.value = true; // 既定は値を追加 / Default: relative

        // 定規の単位を取得（ラベル表示＋pt換算係数）/ Get the ruler unit (label + pt factor)
        var unitInfo = getUnitInfo("rulerType");
        controls.unitFactor = unitInfo.pointsPerUnit;

        var LABEL_WIDTH = 40; // 幅・高さラベルの固定幅 / Fixed width for the width/height labels

        /* ∧∨・↑↓で増減（0で止める。増減のたびにライブプレビュー） / step with the stepper or arrow keys (stops at 0, previews each step) */
        var sizeStepOptions = { min: 0, onStep: function () { runPreview(); } };

        var widthGroup = sizePanel.add("group");
        var widthLabel = widthGroup.add("statictext", undefined, labelText("option.width"), { justify: "right" });
        widthLabel.preferredSize.width = LABEL_WIDTH;
        var widthStepperGroup = widthGroup.add("group");
        widthStepperGroup.orientation = "row";
        widthStepperGroup.alignChildren = ["left", "center"];
        widthStepperGroup.spacing = 0;
        widthStepperGroup.margins = 0;
        var widthInput;
        var widthStepper = addStepper(widthStepperGroup, function () { return widthInput; }, sizeStepOptions);
        widthInput = widthStepperGroup.add("edittext", undefined, "0");
        widthInput.characters = 3;
        bindSteppedArrowKeys(widthInput, widthStepper);
        widthGroup.add("statictext", undefined, unitInfo.label);
        var heightGroup = sizePanel.add("group");
        var heightLabel = heightGroup.add("statictext", undefined, labelText("option.height"), { justify: "right" });
        heightLabel.preferredSize.width = LABEL_WIDTH;
        var heightStepperGroup = heightGroup.add("group");
        heightStepperGroup.orientation = "row";
        heightStepperGroup.alignChildren = ["left", "center"];
        heightStepperGroup.spacing = 0;
        heightStepperGroup.margins = 0;
        var heightInput;
        var heightStepper = addStepper(heightStepperGroup, function () { return heightInput; }, sizeStepOptions);
        heightInput = heightStepperGroup.add("edittext", undefined, "0");
        heightInput.characters = 3;
        bindSteppedArrowKeys(heightInput, heightStepper);
        heightGroup.add("statictext", undefined, unitInfo.label);

        // パスファインダー（パネル＋2カラムのラジオ）/ Pathfinder (panel + two columns of radios)
        var pathfinderPanel = mainPanel.add("panel", undefined, getLabel("pathfinder.label"));
        setupPanel(pathfinderPanel, COLUMN_SPACING);
        pathfinderPanel.orientation = "row"; /* 2カラムのラジオを横に並べる / two columns of radios side by side */
        pathfinderPanel.alignChildren = ["left", "top"];
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
            var pfRadio = pfColumn.add("radiobutton", undefined, getLabel("pathfinder." + pfKey));
            pfRadio.pfKey = pfKey;
            pfRadio.onClick = runPreview;
            controls.pathfinderRadios.push(pfRadio);
        }
        controls.pathfinderRadios[0].value = true; // 既定は「なし」/ Default: none

        // 状況表示 / Status line
        var status = win.add("statictext", undefined, getLabel("status.ready"));
        status.alignment = "fill";

        // ［適用］（閉じるは × / Esc）/ Apply button (close via × / Esc)
        var buttonRow = addButtonRow(win);
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLabel("button.apply"), { name: "ok" });
        btnApply.helpTip = getLabel("tip.apply");
        alignRightOnlyButtonRow(buttonRow);

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
        btnApply.onClick = commitApply;

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

})();
