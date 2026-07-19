#target illustrator
#targetengine "PresetManagerArtboardsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード関連の Illustrator 環境設定を、常駐パレットでまとめて切り替えるユーティリティです。設定は操作した時点で即時反映されます。

- 現在のアートボード：番号・名前を表示し、幅／高さ（定規単位）を確認、［ピクセルグリッドに最適化］で XYWH を整数値へ丸め
- アートボード名と枠線：名前表示／枠線カラー・幅／プリセット（デフォルト・強調・ライト）
- オプション：「裁ち落としを印刷」生成AIボタン表示／ロックまたは非表示オブジェクトを一緒に移動
- 下部ボタン：カンバスカラーの変更／ビデオ定規
- アクティブ時に Esc キーで閉じる
- DOM 読み書き（アートボード情報・リサイズ）は BridgeTalk でメインエンジンへ委譲

### Overview

A persistent-palette utility for batch-switching artboard-related Illustrator preferences. Every change applies immediately.

- Current Artboard: shows number/name, width/height (in the ruler unit), and an [Optimize to Pixel Grid] button that rounds XYWH to integers
- Artboard Name & Border: name visibility / border color & width / presets (Default, Emphasis, Light)
- Options: show the "Print Bleed" generative-AI button / move locked or hidden objects together
- Bottom buttons: Change Canvas Color / Video Ruler
- Press Esc to close while the palette is active
- DOM reads/writes (artboard info, resize) are delegated to the main engine via BridgeTalk

### 更新履歴 / Change Log

- v1.2.0: 標準フォーマットへ整理（概要ブロック、カテゴリ分けローカライズ、ブロックコメント、IIFE 化）。現在のアートボードパネル（番号・名前・幅高さ・ピクセルグリッド最適化）、カンバスカラー切替、Esc で閉じる、プリセットを枠線パネル下部へ移動、DOM 読み書きを BridgeTalk 委譲。
- v1.0.0 (20260323): 初期バージョン。

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardDisplayPresetManager";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                        /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";   /* 作者 / author */
var SCRIPT_RELEASED = "";                              /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                              /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var PANEL_MARGINS = [16, 20, 16, 12]; /* パネルの余白 / Panel margins */
    var PANEL_SPACING = 8;                /* パネル内の間隔 / Spacing inside panels */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のUI言語を取得 / Get the current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "アートボード関連の環境設定", en: "Artboard-Related Preferences" }
        },
        panel: {
            currentArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            artboard: { ja: "アートボード名と枠線", en: "Artboard Name & Border" },
            artboardBorder: { ja: "アートボードの枠線", en: "Artboard Border" },
            options: { ja: "オプション", en: "Options" }
        },
        label: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            strokeColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            strokeWidth: { ja: "ストロークの幅", en: "Stroke Width" }
        },
        checkbox: {
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            showPrintBleedAI: {
                ja: "「裁ち落としを印刷」生成AIボタンを表示",
                en: "Show the \"Print Bleed\" Generative AI Button"
            },
            moveLockedHidden: {
                ja: "ロックまたは非表示オブジェクトを一緒に移動",
                en: "Move Locked or Hidden Objects Together"
            }
        },
        button: {
            optimizePixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to Pixel Grid" },
            reload: { ja: "再読み込み", en: "Reload" },
            canvasColor: { ja: "カンバスカラーの変更", en: "Change Canvas Color" },
            videoRuler: { ja: "ビデオ定規", en: "Video Ruler" }
        },
        preset: {
            /* "default" は ES3 予約語のため引用符付きキーにする / "default" is an ES3 reserved word, so quote the key */
            "default": { ja: "デフォルト", en: "Default" },
            emphasis: { ja: "強調", en: "Emphasis" },
            light: { ja: "ライト", en: "Light" }
        },
        color: {
            lightBlue: { ja: "ライトブルー", en: "Light Blue" },
            red: { ja: "サーモンピンク", en: "Light Red" },
            green: { ja: "グリーン", en: "Green" },
            blue: { ja: "ミディアムブルー", en: "Medium Blue" },
            magenta: { ja: "マゼンタ", en: "Magenta" },
            cyan: { ja: "シアン", en: "Cyan" },
            grey: { ja: "ライトグレー", en: "Light Gray" },
            black: { ja: "ブラック", en: "Black" },
            yellow: { ja: "イエロー", en: "Yellow" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /* ドット区切りキーで LABELS を辿り現在言語の文言を返す（{slash}→/）/ Resolve a dot-path key in LABELS to the current-language text ({slash}→/) */
    function L(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return key;
            node = node[parts[i]];
        }
        if (node == null) return key;
        var text = node[currentLanguage] || node.en || "";
        return text.replace(/\{slash\}/g, "/");
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Unit
    // =========================================

    /* 定規単位（rulerType）→ ラベルと pt 換算係数 / Ruler unit (rulerType) -> label and pt factor */
    var RULER_UNITS = {
        0: { label: "in", factor: 72.0 },
        1: { label: "mm", factor: 72.0 / 25.4 },
        2: { label: "pt", factor: 1.0 },
        3: { label: "pica", factor: 12.0 },
        4: { label: "cm", factor: 72.0 / 2.54 },
        5: { label: "Q", factor: 72.0 / 25.4 * 0.25 },
        6: { label: "px", factor: 1.0 }
    };

    // =========================================
    // 環境設定アクセス / Preferences access
    // =========================================

    var prefs = app.preferences;

    /* 環境設定を取得（kind: "Real"|"Boolean"|"Integer"、失敗時は fb）/ Read a preference by kind (fallback fb on failure) */
    function getPref(kind, key, fb) {
        try { return prefs["get" + kind + "Preference"](key); } catch (e) { return fb; }
    }
    function getReal(key, fb) { return getPref("Real", key, fb); }
    function getBool(key, fb) { return getPref("Boolean", key, fb); }
    function getInt(key, fb) { return getPref("Integer", key, fb); }

    /* 現在の定規単位（ラベル＋pt換算係数）を取得 / Get the current ruler unit (label + pt factor) */
    function getRulerUnit() {
        var code = getInt("rulerType", 2);
        return RULER_UNITS[code] || RULER_UNITS[2];
    }

    /* pt 値を現在単位の表示文字列へ（小数2桁・末尾0除去）/ Convert a pt value to a current-unit display string */
    function ptToUnitText(pt, unit) {
        if (isNaN(pt)) return "";
        var v = pt / unit.factor;
        v = Math.round(v * 100) / 100;
        return String(v);
    }

    /* 数値を範囲内に収める / Clamp a number to a range */
    function clamp(n, min, max) {
        return Math.max(min, Math.min(max, n));
    }

    // =========================================
    // アートボード枠線カラー / Artboard border color
    // =========================================

    /* 枠線カラーのプリセット（ドロップダウンの並び順）/ Border color presets (dropdown order) */
    var STROKE_COLOR_PRESETS = [
        { labelKey: "color.lightBlue", r: 0.29, g: 0.52, b: 1.0 },
        { labelKey: "color.red",       r: 1.0,  g: 0.29, b: 0.29 },
        { labelKey: "color.green",     r: 0.0,  g: 0.65, b: 0.31 },
        { labelKey: "color.blue",      r: 0.0,  g: 0.45, b: 0.78 },
        { labelKey: "color.magenta",   r: 1.0,  g: 0.0,  b: 1.0 },
        { labelKey: "color.cyan",      r: 0.0,  g: 1.0,  b: 1.0 },
        { labelKey: "color.grey",      r: 0.65, g: 0.65, b: 0.65 },
        { labelKey: "color.black",     r: 0.0,  g: 0.0,  b: 0.0 },
        { labelKey: "color.yellow",    r: 1.0,  g: 1.0,  b: 0.0 }
    ];
    var STROKE_COLOR_BLACK_INDEX = 7;

    /* ドロップダウン用のカラー名配列を生成 / Build the color name list for the dropdown */
    function buildStrokeColorNames() {
        var names = [];
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            names.push(L(STROKE_COLOR_PRESETS[i].labelKey));
        }
        return names;
    }

    /* RGB に最も近いプリセットの index を返す / Return the index of the preset closest to the given RGB */
    function findClosestStrokeColor(r, g, b) {
        var bestIdx = 0;
        var bestDist = Infinity;
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            var p = STROKE_COLOR_PRESETS[i];
            var dist = Math.abs(p.r - r) + Math.abs(p.g - g) + Math.abs(p.b - b);
            if (dist < bestDist) {
                bestDist = dist;
                bestIdx = i;
            }
        }
        return bestIdx;
    }

    /* labelKey からカラープリセットの index を取得 / Find a color preset index by labelKey */
    function findColorIndexByKey(labelKey) {
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            if (STROKE_COLOR_PRESETS[i].labelKey === labelKey) return i;
        }
        return STROKE_COLOR_BLACK_INDEX;
    }

    // =========================================
    // 表示プリセット / Display presets
    // =========================================

    /* プリセット定義（適用・判定の両方で使う単一の真実）/ Preset definitions (single source for both apply and detect) */
    /* "default" は ES3 予約語のため引用符付きキー / "default" is an ES3 reserved word, so quote the key */
    var PRESETS = {
        "default": { showName: true,  colorKey: "color.black", strokeWidth: 1, printBleed: true,  moveLocked: false },
        emphasis:  { showName: false, colorKey: "color.red",   strokeWidth: 3, printBleed: false, moveLocked: true },
        light:     { showName: false, colorKey: "color.grey",  strokeWidth: 1, printBleed: false, moveLocked: true }
    };

    // =========================================
    // レイアウトヘルパー / Layout helpers
    // =========================================

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
    function trimButtonHeight(button, px) {
        try {
            var w = button.size.width;
            var h = button.size.height - px;
            /* preferredSize も固定しないと再レイアウトで戻るため両方設定 / Pin preferredSize too, or a relayout reverts it */
            button.preferredSize = [w, h];
            button.size = [w, h];
        } catch (e) {}
    }

    // =========================================
    // アートボード情報（BridgeTalk 委譲）/ Artboard info (BridgeTalk delegation)
    // 常駐パレットのイベントハンドラ内では DOM 接続を失うため、DOM の読み書きはメインエンジンへ委譲する
    // Inside persistent-palette event handlers the DOM connection is lost, so all DOM access is delegated to the main engine
    // =========================================

    /* フィールド区切り（アートボード名に現れにくい ASCII 文字列／エスケープ不要）/ Field separator (ASCII, unlikely in names, no escaping) */
    var AB_SEP = "<|>";

    /* メインエンジンへコードを委譲し、結果文字列を onResult へ渡す / Delegate code to the main engine; pass the result string to onResult */
    function delegateToMainEngine(bodyCode, onResult) {
        /* フォールバックで同一エンジン実行しても DOM 接続が無いため無意味。失敗時は空結果で UI を更新 */
        /* A same-engine fallback is pointless (no DOM connection here), so just yield an empty result on failure */
        try {
            var bt = new BridgeTalk();
            bt.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
            bt.body = bodyCode;
            bt.onResult = function (resObj) { if (onResult) onResult(String(resObj.body)); };
            bt.onError = function (message) { if (onResult) onResult(""); };
            bt.send();
        } catch (e) {
            if (onResult) onResult("");
        }
    }

    /* メインエンジンで実行するコードの本体を組み立て（read→丸め／リサイズ後に最新情報を返す）
       Build the code body run on the main engine (read, or round/resize then return the latest info).
       op: "read" | "round" | "resize"。resize は widthPt/heightPt を使用 */
    function buildArtboardBody(op, widthPt, heightPt) {
        var mutate = "";
        if (op === "round") {
            mutate = "ab.artboardRect=[Math.round(r[0]),Math.round(r[1]),Math.round(r[2]),Math.round(r[3])];r=ab.artboardRect;";
        } else if (op === "resize") {
            /* 負数連結による '--' 構文エラーを避けるため括弧で囲む / Wrap in parens to avoid '--' from negative numbers */
            mutate = "ab.artboardRect=[r[0],r[1],r[0]+(" + Number(widthPt) + "),r[1]-(" + Number(heightPt) + ")];r=ab.artboardRect;";
        }
        return "" +
            "(function(){try{" +
            "var d=app.activeDocument;" +
            "var i=d.artboards.getActiveArtboardIndex();" +
            "var ab=d.artboards[i];var r=ab.artboardRect;" +
            mutate +
            "return (i+1)+\"" + AB_SEP + "\"+ab.name+\"" + AB_SEP + "\"+(r[2]-r[0])+\"" + AB_SEP + "\"+(r[1]-r[3]);" +
            "}catch(e){return \"\";}})();";
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /* パレットを構築して表示 / Build and show the palette */
    function main() {

        /* 重複起動を避ける：既に開いていれば前面に出して終了 / Avoid duplicate launch: bring the existing palette forward and return */
        try {
            if ($.global.__artboardPresetPalette) {
                $.global.__artboardPresetPalette.show();
                return;
            }
        } catch (e) {
            $.global.__artboardPresetPalette = null;
        }

        var dlg = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        /* 常駐エンジンに参照を保持／閉じたらクリア / Hold the reference in the persistent engine; clear on close */
        $.global.__artboardPresetPalette = dlg;
        dlg.onClose = function () {
            $.global.__artboardPresetPalette = null;
        };

        var mainGroup = dlg.add("group");
        setupGroup(mainGroup, "column");

        // -----------------------------------------
        // 現在のアートボード / Current artboard
        // -----------------------------------------
        var panelCurrentArtboard = mainGroup.add("panel", undefined, L("panel.currentArtboard"));
        setupPanel(panelCurrentArtboard);

        /* 番号・名前（左右中央）/ Number and name (centered) */
        var abInfoRow = panelCurrentArtboard.add("group");
        abInfoRow.alignment = ["fill", "top"];
        abInfoRow.alignChildren = ["center", "center"];
        var abInfoText = abInfoRow.add("statictext", undefined, "—");
        abInfoText.characters = 28;
        abInfoText.justify = "center";

        /* 幅・高さ（横並び・天地中央）/ Width and height (side by side, vertically centered) */
        /* statictext は edittext より低く、ScriptUI(mac) では上寄せになりがちなので各要素に縦中央を明示
           statictext is shorter than edittext and tends to top-align on ScriptUI (mac), so center each control explicitly */
        function addSizeField(parentRow, labelKey, unitLabel) {
            var group = parentRow.add("group");
            group.orientation = "row";
            group.alignChildren = ["left", "center"];
            group.alignment = ["left", "center"];
            var label = group.add("statictext", undefined, labelText(labelKey));
            label.alignment = ["left", "center"];
            var field = group.add("edittext", undefined, "");
            field.characters = 5;
            field.alignment = ["left", "center"];
            var unit = group.add("statictext", undefined, unitLabel);
            unit.alignment = ["left", "center"];
            return { field: field, unit: unit };
        }

        var abSizeRow = panelCurrentArtboard.add("group");
        abSizeRow.orientation = "row";
        abSizeRow.alignChildren = ["left", "center"];
        abSizeRow.spacing = 16;

        var widthField = addSizeField(abSizeRow, "label.width", getRulerUnit().label);
        var etWidth = widthField.field;
        var stWidthUnit = widthField.unit;

        var heightField = addSizeField(abSizeRow, "label.height", getRulerUnit().label);
        var etHeight = heightField.field;
        var stHeightUnit = heightField.unit;

        /* ボタン行（ピクセルグリッド最適化 / 再読み込み）/ Button row (Optimize to pixel grid / Reload) */
        var abButtonRow = panelCurrentArtboard.add("group");
        abButtonRow.orientation = "row";
        abButtonRow.alignChildren = ["left", "center"];
        abButtonRow.alignment = ["left", "top"];
        /* ピクセルグリッドに最適化（XYWH を整数値へ）/ Optimize to pixel grid (round XYWH to integers) */
        var btnOptimize = abButtonRow.add("button", undefined, L("button.optimizePixelGrid"));
        /* 再読み込み（現在のアートボード情報を取り直す）/ Reload (re-read current artboard info) */
        var btnReload = abButtonRow.add("button", undefined, L("button.reload"));

        // -----------------------------------------
        // アートボード名と枠線 / Artboard name & border
        // -----------------------------------------
        var panelArtboard = mainGroup.add("panel", undefined, L("panel.artboard"));
        setupPanel(panelArtboard);

        var cbShowArtboardName = panelArtboard.add("checkbox", undefined, L("checkbox.showArtboardName"));

        /* 枠線サブパネル / Border sub-panel */
        var panelArtboardBorder = panelArtboard.add("panel", undefined, L("panel.artboardBorder"));
        setupPanel(panelArtboardBorder);

        /* ハイライトのカラー / Highlight color */
        var strokeColorRow = panelArtboardBorder.add("group");
        setupGroup(strokeColorRow, "row");
        strokeColorRow.add("statictext", undefined, labelText("label.strokeColor"));
        var ddStrokeColor = strokeColorRow.add("dropdownlist", undefined, buildStrokeColorNames());

        /* ストロークの幅（1〜4）/ Stroke width (1-4) */
        var strokeWidthRow = panelArtboardBorder.add("group");
        setupGroup(strokeWidthRow, "row");
        strokeWidthRow.add("statictext", undefined, labelText("label.strokeWidth"));
        var rbStrokeWidth1 = strokeWidthRow.add("radiobutton", undefined, "1");
        var rbStrokeWidth2 = strokeWidthRow.add("radiobutton", undefined, "2");
        var rbStrokeWidth3 = strokeWidthRow.add("radiobutton", undefined, "3");
        var rbStrokeWidth4 = strokeWidthRow.add("radiobutton", undefined, "4");
        var rbStrokeWidths = [rbStrokeWidth1, rbStrokeWidth2, rbStrokeWidth3, rbStrokeWidth4];

        /* プリセット（パネル最下部・左右中央）/ Presets (bottom of panel, centered) */
        var presetRow = panelArtboard.add("group");
        presetRow.orientation = "row";
        presetRow.alignment = ["center", "top"];
        var rbPresetDefault = presetRow.add("radiobutton", undefined, L("preset.default"));
        var rbPresetEmphasis = presetRow.add("radiobutton", undefined, L("preset.emphasis"));
        var rbPresetLight = presetRow.add("radiobutton", undefined, L("preset.light"));

        // -----------------------------------------
        // オプション / Options
        // -----------------------------------------
        var panelOptions = mainGroup.add("panel", undefined, L("panel.options"));
        setupPanel(panelOptions);

        var cbShowPrintBleedAI = panelOptions.add("checkbox", undefined, L("checkbox.showPrintBleedAI"));
        var cbMoveLockedHidden = panelOptions.add("checkbox", undefined, L("checkbox.moveLockedHidden"));

        // -----------------------------------------
        // 下部ボタン行（カンバスカラー / ビデオ定規）/ Bottom button row (Canvas color / Video ruler)
        // -----------------------------------------
        var outerGroup = mainGroup.add("group");
        outerGroup.orientation = "row";
        outerGroup.alignChildren = ["fill", "center"];
        outerGroup.alignment = ["fill", "bottom"];

        var leftGroup = outerGroup.add("group");
        leftGroup.orientation = "row";
        leftGroup.alignChildren = ["left", "center"];
        var btnCanvasColor = leftGroup.add("button", undefined, L("button.canvasColor"));

        var spacer = outerGroup.add("group");
        spacer.alignment = ["fill", "fill"];

        var rightGroup = outerGroup.add("group");
        rightGroup.orientation = "row";
        rightGroup.alignChildren = ["right", "center"];
        var btnVideoRuler = rightGroup.add("button", undefined, L("button.videoRuler"));

        // =========================================
        // 即時反映：環境設定の書き込み / Immediate apply: preference writes
        // =========================================

        /* アートボード表示後の強制再描画（値変更が即時反映されないため）/ Force a viewport refresh (changes don't repaint by themselves) */
        function refreshArtboardDisplay() {
            try {
                app.executeMenuCommand("zoomout");
                app.executeMenuCommand("zoomin");
            } catch (e) {}
        }

        /* 選択中のストローク幅（1〜4）を取得 / Get the selected stroke width (1-4) */
        function getSelectedStrokeWidth() {
            for (var i = 0; i < rbStrokeWidths.length; i++) {
                if (rbStrokeWidths[i].value) return i + 1;
            }
            return 1;
        }

        /* アートボード名・枠線カラー・幅を反映 / Apply artboard name, border color and width */
        function applyArtboardDisplaySettings() {
            prefs.setBooleanPreference("showArtboardLabelOnCanvas", cbShowArtboardName.value);
            var idx = ddStrokeColor.selection ? ddStrokeColor.selection.index : STROKE_COLOR_BLACK_INDEX;
            var c = STROKE_COLOR_PRESETS[idx];
            prefs.setRealPreference("ArtboardBBColorRed", c.r);
            prefs.setRealPreference("ArtboardBBColorGreen", c.g);
            prefs.setRealPreference("ArtboardBBColorBlue", c.b);
            prefs.setRealPreference("ArtboardBBWidth", getSelectedStrokeWidth());
            refreshArtboardDisplay();
        }

        /* オプション（生成AIボタン／一緒に移動）を反映 / Apply options (AI button / move together) */
        function applyOptionSettings() {
            prefs.setBooleanPreference("enablePrintBleedWidget", cbShowPrintBleedAI.value);
            prefs.setBooleanPreference("moveLockedAndHiddenArt", cbMoveLockedHidden.value);
        }

        /* すべての設定を反映 / Apply all settings */
        function applyAllSettings() {
            applyArtboardDisplaySettings();
            applyOptionSettings();
        }

        /* ストローク幅ラジオ（1〜4）を選択 / Select the stroke-width radio (1-4) */
        function setStrokeWidthRadio(width) {
            for (var i = 0; i < rbStrokeWidths.length; i++) {
                rbStrokeWidths[i].value = (i === width - 1);
            }
        }

        /* プリセットを UI へ適用 / Apply a preset to the UI */
        function applyPreset(presetKey) {
            var p = PRESETS[presetKey];
            if (!p) return;
            cbShowArtboardName.value = p.showName;
            ddStrokeColor.selection = findColorIndexByKey(p.colorKey);
            setStrokeWidthRadio(p.strokeWidth);
            cbShowPrintBleedAI.value = p.printBleed;
            cbMoveLockedHidden.value = p.moveLocked;
        }

        // =========================================
        // 現在のアートボード / Current artboard
        // =========================================

        /* 番号と名前から表示用ラベルを組み立て / Build the display label from number and name */
        function formatArtboardLabel(num, name) {
            var sep = (currentLanguage === "ja") ? "：" : ": ";
            return "#" + num + sep + name;
        }

        /* 委譲結果（番号<|>名前<|>幅pt<|>高さpt）を UI へ反映 / Apply the delegated result to the UI */
        /* showAlert=true のとき、空結果（ドキュメントなし）でアラート表示 / Alert on empty result when showAlert is true */
        function applyArtboardResult(result, showAlert) {
            var unit = getRulerUnit();
            stWidthUnit.text = unit.label;
            stHeightUnit.text = unit.label;
            var parts = result ? result.split(AB_SEP) : [];
            if (parts.length < 4) {
                abInfoText.text = "—";
                etWidth.text = "";
                etHeight.text = "";
                if (showAlert) alert(L("alert.noDocument"));
                return;
            }
            abInfoText.text = formatArtboardLabel(parts[0], parts[1]);
            etWidth.text = ptToUnitText(parseFloat(parts[2]), unit);
            etHeight.text = ptToUnitText(parseFloat(parts[3]), unit);
        }

        /* アートボード操作を委譲し、結果を UI へ反映 / Delegate an artboard op and apply the result */
        function runArtboardOp(op, widthPt, heightPt, showAlert) {
            delegateToMainEngine(buildArtboardBody(op, widthPt, heightPt), function (result) {
                applyArtboardResult(result, showAlert);
            });
        }

        /* 現在のアートボード情報を取得して UI へ反映 / Read current artboard info and apply to the UI */
        function refreshArtboardInfo() {
            runArtboardOp("read", 0, 0, false);
        }

        /* アクティブアートボードの XYWH を整数値へ丸めて表示更新 / Round the active artboard's XYWH and refresh the display */
        function optimizeArtboardPixelGrid() {
            runArtboardOp("round", 0, 0, true);
        }

        /* 幅・高さフィールドの値でアクティブアートボードをリサイズ / Resize the active artboard from the width/height fields */
        function applyArtboardSizeFromFields() {
            var unit = getRulerUnit();
            var w = parseFloat(etWidth.text);
            var h = parseFloat(etHeight.text);
            /* 不正値は現在値へ戻す / Restore current values on invalid input */
            if (isNaN(w) || isNaN(h) || w <= 0 || h <= 0) {
                refreshArtboardInfo();
                return;
            }
            runArtboardOp("resize", w * unit.factor, h * unit.factor, true);
        }

        // =========================================
        // 値反映：環境設定 → UI / Reflect preferences into the UI
        // =========================================

        /* 現在値に一致するプリセットキーを返す（なければ null）/ Return the preset key matching the current state (null if none) */
        function detectPresetKey(closestIdx, curWidth) {
            for (var key in PRESETS) {
                if (!PRESETS.hasOwnProperty(key)) continue;
                var p = PRESETS[key];
                if (cbShowArtboardName.value === p.showName &&
                    closestIdx === findColorIndexByKey(p.colorKey) &&
                    curWidth === p.strokeWidth &&
                    cbShowPrintBleedAI.value === p.printBleed &&
                    cbMoveLockedHidden.value === p.moveLocked) {
                    return key;
                }
            }
            return null;
        }

        /* 現在値からプリセットを判定して該当ラジオを選択 / Detect and select the matching preset radio */
        function detectPreset(closestIdx, curWidth) {
            var key = detectPresetKey(closestIdx, curWidth);
            rbPresetDefault.value = (key === "default");
            rbPresetEmphasis.value = (key === "emphasis");
            rbPresetLight.value = (key === "light");
        }

        /* 現在の環境設定を UI へ反映し、プリセットの一致を判定 / Reflect current preferences and detect the matching preset */
        function reflectPreferences() {
            cbShowArtboardName.value = !!getBool("showArtboardLabelOnCanvas", false);

            var cr = getReal("ArtboardBBColorRed", 0.0);
            var cg = getReal("ArtboardBBColorGreen", 0.0);
            var cb = getReal("ArtboardBBColorBlue", 0.0);
            var closestIdx = findClosestStrokeColor(cr, cg, cb);
            ddStrokeColor.selection = closestIdx;

            var curWidth = clamp(Math.round(getReal("ArtboardBBWidth", 1.0)), 1, 4);
            setStrokeWidthRadio(curWidth);

            cbShowPrintBleedAI.value = !!getBool("enablePrintBleedWidget", false);
            cbMoveLockedHidden.value = !!getBool("moveLockedAndHiddenArt", false);

            detectPreset(closestIdx, curWidth);
        }

        // =========================================
        // イベント設定 / Event wiring
        // =========================================

        rbPresetDefault.onClick = function () {
            applyPreset("default");
            applyAllSettings();
        };
        rbPresetEmphasis.onClick = function () {
            applyPreset("emphasis");
            applyAllSettings();
        };
        rbPresetLight.onClick = function () {
            applyPreset("light");
            applyAllSettings();
        };

        cbShowArtboardName.onClick = function () { applyArtboardDisplaySettings(); };
        cbShowPrintBleedAI.onClick = function () { applyOptionSettings(); };
        cbMoveLockedHidden.onClick = function () { applyOptionSettings(); };
        ddStrokeColor.onChange = function () { applyArtboardDisplaySettings(); };

        for (var sw = 0; sw < rbStrokeWidths.length; sw++) {
            rbStrokeWidths[sw].onClick = function () { applyArtboardDisplaySettings(); };
        }

        btnOptimize.onClick = function () {
            optimizeArtboardPixelGrid();
        };

        /* 再読み込み：現在のアートボード情報を取り直す / Reload: re-read the current artboard info */
        btnReload.onClick = function () {
            refreshArtboardInfo();
        };

        /* 幅・高さの確定でアートボードをリサイズ / Resize the artboard when width/height are committed */
        etWidth.onChange = function () { applyArtboardSizeFromFields(); };
        etHeight.onChange = function () { applyArtboardSizeFromFields(); };

        /* カンバスカラーの変更（uiCanvasIsWhite: 1=ON / 0=OFF）。値変更後は zoom 再描画で反映 / Toggle canvas color; refresh via zoom afterwards */
        btnCanvasColor.onClick = function () {
            var isWhite = getInt("uiCanvasIsWhite", 0) === 1;
            prefs.setIntegerPreference("uiCanvasIsWhite", isWhite ? 0 : 1);
            refreshArtboardDisplay();
        };

        btnVideoRuler.onClick = function () {
            try { app.executeMenuCommand("videoruler"); } catch (e) {}
        };

        /* アクティブ時に Esc で閉じる / Close on Esc while active */
        dlg.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") {
                dlg.close();
            }
        });

        /* 再アクティブ時：外部変更とアートボードの切り替えに追従 / On re-activate: follow external changes and artboard switches */
        dlg.onActivate = function () {
            reflectPreferences();
            refreshArtboardInfo();
        };

        // =========================================
        // 初期化と表示 / Initialize and show
        // =========================================
        reflectPreferences();
        refreshArtboardInfo();

        dlg.center();
        dlg.show();

        /* レイアウト確定後にボタン高さを 2px 詰める / Trim the button heights by 2px after layout */
        trimButtonHeight(btnOptimize, 2);
        trimButtonHeight(btnReload, 2);
    }

    main();

}());
