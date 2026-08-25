#target illustrator
#targetengine "ArtboardDisplayPresetManagerPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード関連のIllustrator環境設定を、常駐パレットでまとめて切り替えます。
アートボードのサイズ確認とリサイズ、名前表示と枠線、プリセット、カンバスカラーの切り替えに対応します。

詳細は README を参照してください。

### Overview

A persistent palette for switching the artboard-related Illustrator preferences.
It covers checking and resizing artboards, name and border display, presets, and the canvas color.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardDisplayPresetManager"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardDisplayPresetManager.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardDisplayPresetManager.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9eba8ab03170"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;               /* パネル内の要素間隔 / panel spacing */

    /* 常駐エンジンにパレット参照を保持するキー / Key holding the palette reference in the persistent engine */
    var PALETTE_GLOBAL_KEY = "__artboardDisplayPresetPalette";

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定を適用する（ボタン列・ラベル＋入力欄など）
     * statictext は edittext より低く ScriptUI(mac) では上寄せになりやすいため、天地中央を明示する
     * @param {Group} group - 対象グループ
     * @param {string} [alignment] - グループ自身の整列（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）
     * @param {Button} button - 対象ボタン
     * @param {number} px - 詰める高さ（px）
     * @returns {void}
     */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) {}
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を取得する
     * @returns {string} "ja" または "en"
     */
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
            artboardDisplay: { ja: "アートボード名と枠線", en: "Artboard Name & Border" },
            artboardBorder: { ja: "アートボードの枠線", en: "Artboard Border" },
            options: { ja: "オプション", en: "Options" }
        },
        label: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            borderColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            borderWidth: { ja: "ストロークの幅", en: "Stroke Width" }
        },
        checkbox: {
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            /* showPrintBleedAI は現在未使用（PRINT_BLEED_WIDGET を参照）/ showPrintBleedAI is currently unused (see the PRINT_BLEED_WIDGET note) */
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

    /**
     * ドット区切りキーで LABELS を辿り、現在言語の文言を返す（{slash}→/）
     * @param {string} labelPath - "panel.options" のようなドット区切りキー
     * @returns {string} 現在言語の文言（見つからない場合はキーをそのまま返す）
     */
    function L(labelPath) {
        var parts = labelPath.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return labelPath;
            node = node[parts[i]];
        }
        if (node == null) return labelPath;
        var text = node[currentLanguage] || node.en || "";
        return text.replace(/\{slash\}/g, "/");
    }

    /**
     * コロン付きラベルを返す（日本語は全角、英語は半角）
     * @param {string} labelPath - LABELS のドット区切りキー
     * @returns {string} コロンを付けた文言
     */
    function labelWithColon(labelPath) {
        return L(labelPath) + (currentLanguage === "ja" ? "：" : ":");
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
    var RULER_UNIT_POINT = RULER_UNITS[2];

    // =========================================
    // 環境設定アクセス / Preferences access
    // =========================================

    var prefs = app.preferences;

    /* -----------------------------------------
       PRINT_BLEED_WIDGET：「裁ち落としを印刷」生成AIボタンの表示切り替えについて

       enablePrintBleedWidget は setBooleanPreference で確実に書き換えられ、環境設定ダイアログの
       表示にも反映されるが、カンバス上のウィジェットは再評価されない。app.redraw()、ズーム、
       ツール切り替え、プレビュー／アウトライン切り替え、アートボード再設定、ドキュメント切り替えの
       いずれでも反映せず、スクリプトから即時反映させる手段が見つからなかった（Illustrator が
       起動時か環境設定ダイアログの適用時にしか読まないと思われる）。

       将来のバージョンで挙動が戻る可能性があるため、関連コードは削除せずコメントアウトで保留する。
       復活させるときは、この定数・LABELS の checkbox.showPrintBleedAI・buildOptionsPanel・
       applyPrintBleedWidgetSetting・applyOptionSettings・applyPresetToUI・detectPresetKey・
       reflectPreferences のコメントアウトを戻すこと。

       The preference is written reliably, but Illustrator never re-evaluates the canvas widget,
       and no scripted refresh triggers it. The item is parked (commented out) rather than removed.
       ----------------------------------------- */
    // var PRINT_BLEED_WIDGET_KEY = "enablePrintBleedWidget";

    /**
     * 環境設定を型指定で読み取る
     * @param {string} kind - "Real" | "Boolean" | "Integer"
     * @param {string} key - 環境設定キー
     * @param {*} fallback - 取得に失敗したときの値
     * @returns {*} 取得値（失敗時は fallback）
     */
    function getPref(kind, key, fallback) {
        try { return prefs["get" + kind + "Preference"](key); } catch (e) { return fallback; }
    }

    /**
     * 実数の環境設定を読み取る
     * @param {string} key - 環境設定キー
     * @param {number} fallback - 取得に失敗したときの値
     * @returns {number} 取得値
     */
    function getReal(key, fallback) { return getPref("Real", key, fallback); }

    /**
     * 真偽値の環境設定を読み取る
     * @param {string} key - 環境設定キー
     * @param {boolean} fallback - 取得に失敗したときの値
     * @returns {boolean} 取得値
     */
    function getBool(key, fallback) { return getPref("Boolean", key, fallback); }

    /**
     * 整数の環境設定を読み取る
     * @param {string} key - 環境設定キー
     * @param {number} fallback - 取得に失敗したときの値
     * @returns {number} 取得値
     */
    function getInt(key, fallback) { return getPref("Integer", key, fallback); }

    /**
     * 現在の定規単位（ラベル＋pt換算係数）を取得する
     * @returns {object} { label: string, factor: number }
     */
    function getRulerUnit() {
        return RULER_UNITS[getInt("rulerType", 2)] || RULER_UNIT_POINT;
    }

    /**
     * pt 値を現在単位の表示文字列へ変換する（小数2桁で丸め）
     * @param {number} pointValue - pt 値
     * @param {object} unit - getRulerUnit() の戻り値
     * @returns {string} 表示用文字列（数値でない場合は空文字）
     */
    function pointToUnitText(pointValue, unit) {
        if (isNaN(pointValue)) return "";
        return String(Math.round(pointValue / unit.factor * 100) / 100);
    }

    /**
     * 数値を範囲内に収める
     * @param {number} value - 対象の値
     * @param {number} min - 下限
     * @param {number} max - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clamp(value, min, max) {
        return Math.max(min, Math.min(max, value));
    }

    /**
     * メニューコマンドを実行する（未対応・ドキュメントなしでも無視する）
     * @param {string} command - メニューコマンド名
     * @returns {void}
     */
    function runMenuCommand(command) {
        try { app.executeMenuCommand(command); } catch (e) {}
    }

    /**
     * アートボード表示を強制的に再描画する（環境設定の変更は自動では反映されないため）
     * @returns {void}
     */
    function refreshArtboardDisplay() {
        runMenuCommand("zoomout");
        runMenuCommand("zoomin");
    }

    // =========================================
    // アートボード枠線カラー / Artboard border color
    // =========================================

    /* 枠線カラーのプリセット（ドロップダウンの並び順）/ Border color presets (dropdown order) */
    var BORDER_COLOR_PRESETS = [
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
    var BORDER_COLOR_BLACK_INDEX = 7;

    /* 枠線の太さの選択肢（1〜4）/ Border width choices (1-4) */
    var BORDER_WIDTH_CHOICES = [1, 2, 3, 4];

    /**
     * ドロップダウン用のカラー名配列を生成する
     * @returns {array} 現在言語のカラー名の配列
     */
    function buildBorderColorNames() {
        var names = [];
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            names.push(L(BORDER_COLOR_PRESETS[i].labelKey));
        }
        return names;
    }

    /**
     * 指定 RGB にもっとも近いカラープリセットの index を返す
     * @param {number} red - 赤成分（0〜1）
     * @param {number} green - 緑成分（0〜1）
     * @param {number} blue - 青成分（0〜1）
     * @returns {number} もっとも近いプリセットの index
     */
    function findClosestBorderColorIndex(red, green, blue) {
        var closestIndex = 0;
        var closestDistance = Infinity;
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            var preset = BORDER_COLOR_PRESETS[i];
            var distance = Math.abs(preset.r - red) + Math.abs(preset.g - green) + Math.abs(preset.b - blue);
            if (distance < closestDistance) {
                closestDistance = distance;
                closestIndex = i;
            }
        }
        return closestIndex;
    }

    /**
     * labelKey からカラープリセットの index を取得する
     * @param {string} labelKey - LABELS のカラーキー（例: "color.red"）
     * @returns {number} プリセットの index（見つからない場合はブラック）
     */
    function findBorderColorIndexByKey(labelKey) {
        for (var i = 0; i < BORDER_COLOR_PRESETS.length; i++) {
            if (BORDER_COLOR_PRESETS[i].labelKey === labelKey) return i;
        }
        return BORDER_COLOR_BLACK_INDEX;
    }

    // =========================================
    // 表示プリセット / Display presets
    // =========================================

    /* プリセット定義（適用・判定の両方で使う単一の真実）/ Preset definitions (single source for both apply and detect) */
    /* "default" は ES3 予約語のため引用符付きキー / "default" is an ES3 reserved word, so quote the key */
    /* printBleed は現在未使用（PRINT_BLEED_WIDGET を参照）/ printBleed is currently unused (see the PRINT_BLEED_WIDGET note) */
    var PRESET_KEYS = ["default", "emphasis", "light"];
    var PRESETS = {
        "default": { showName: true,  colorKey: "color.black", borderWidth: 1, printBleed: true,  moveLocked: false },
        emphasis:  { showName: false, colorKey: "color.red",   borderWidth: 3, printBleed: false, moveLocked: true },
        light:     { showName: false, colorKey: "color.grey",  borderWidth: 1, printBleed: false, moveLocked: true }
    };

    // =========================================
    // アートボード情報（BridgeTalk 委譲）/ Artboard info (BridgeTalk delegation)
    // 常駐パレットのイベントハンドラ内では DOM 接続を失うため、DOM の読み書きはメインエンジンへ委譲する
    // Inside persistent-palette event handlers the DOM connection is lost, so all DOM access is delegated to the main engine
    // =========================================

    /* フィールド区切り（アートボード名に現れにくい ASCII 文字列／エスケープ不要）/ Field separator (ASCII, unlikely in names, no escaping) */
    var ARTBOARD_FIELD_SEPARATOR = "<|>";

    /**
     * メインエンジンへコードを委譲し、結果文字列をコールバックへ渡す
     * フォールバックで同一エンジン実行しても DOM 接続が無いため、失敗時は空文字を返す
     * @param {string} bodyCode - メインエンジンで実行するコード
     * @param {function} [onResult] - 結果文字列を受け取るコールバック（省略可）
     * @returns {void}
     */
    function delegateToMainEngine(bodyCode, onResult) {
        if (typeof onResult !== "function") onResult = function () {};
        try {
            var bridge = new BridgeTalk();
            bridge.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
            bridge.body = bodyCode;
            bridge.onResult = function (response) { onResult(String(response.body)); };
            bridge.onError = function (message) { onResult(""); };
            bridge.send();
        } catch (e) {
            onResult("");
        }
    }

    /**
     * メインエンジンで実行するコード本体を組み立てる（read／丸め／リサイズ後に最新情報を返す）
     * @param {string} operation - "read" | "round" | "resize"
     * @param {number} [widthPoint] - resize 時の幅（pt）
     * @param {number} [heightPoint] - resize 時の高さ（pt）
     * @returns {string} メインエンジンへ送るコード
     */
    function buildArtboardBody(operation, widthPoint, heightPoint) {
        var mutation = "";
        if (operation === "round") {
            mutation = "ab.artboardRect=[Math.round(r[0]),Math.round(r[1]),Math.round(r[2]),Math.round(r[3])];r=ab.artboardRect;";
        } else if (operation === "resize") {
            /* 負数連結による '--' 構文エラーを避けるため括弧で囲む / Wrap in parens to avoid '--' from negative numbers */
            mutation = "ab.artboardRect=[r[0],r[1],r[0]+(" + Number(widthPoint) + "),r[1]-(" + Number(heightPoint) + ")];r=ab.artboardRect;";
        }
        return "" +
            "(function(){try{" +
            "var d=app.activeDocument;" +
            "var i=d.artboards.getActiveArtboardIndex();" +
            "var ab=d.artboards[i];var r=ab.artboardRect;" +
            mutation +
            "return (i+1)+\"" + ARTBOARD_FIELD_SEPARATOR + "\"+ab.name+\"" + ARTBOARD_FIELD_SEPARATOR + "\"+(r[2]-r[0])+\"" + ARTBOARD_FIELD_SEPARATOR + "\"+(r[1]-r[3]);" +
            "}catch(e){return \"\";}})();";
    }

    // =========================================
    // UI構築 / UI construction
    // =========================================

    /**
     * ラベル・入力欄・単位表示を並べたサイズ入力欄を作る
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} labelPath - ラベルの LABELS キー
     * @param {string} unitLabel - 単位ラベル（例: "mm"）
     * @returns {object} { input: EditText, unitText: StaticText }
     */
    function addSizeField(parentRow, labelPath, unitLabel) {
        var fieldGroup = parentRow.add("group");
        setupRow(fieldGroup, "left", 4);
        fieldGroup.add("statictext", undefined, labelWithColon(labelPath));
        var input = fieldGroup.add("edittext", undefined, "");
        input.characters = 5;
        var unitText = fieldGroup.add("statictext", undefined, unitLabel);
        return { input: input, unitText: unitText };
    }

    /**
     * 「現在のアートボード」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildCurrentArtboardPanel(parentWindow) {
        var currentArtboardPanel = parentWindow.add("panel", undefined, L("panel.currentArtboard"));
        setupPanel(currentArtboardPanel, 8);

        /* 番号・名前（左右中央）/ Number and name (centered) */
        var infoRow = currentArtboardPanel.add("group");
        setupRow(infoRow, "fill");
        infoRow.alignChildren = ["center", "center"];
        var infoText = infoRow.add("statictext", undefined, "—");
        infoText.characters = 28;
        infoText.justify = "center";

        /* 幅・高さ（横並び）/ Width and height (side by side) */
        var sizeRow = currentArtboardPanel.add("group");
        setupRow(sizeRow, "left", 16);
        var unitLabel = getRulerUnit().label;
        var widthField = addSizeField(sizeRow, "label.width", unitLabel);
        var heightField = addSizeField(sizeRow, "label.height", unitLabel);

        /* ボタン行（パネル幅いっぱいには広げない）/ Button row (do not stretch to the panel width) */
        var actionRow = currentArtboardPanel.add("group");
        setupRow(actionRow, "left");
        var optimizeButton = actionRow.add("button", undefined, L("button.optimizePixelGrid"));
        optimizeButton.alignment = "left";
        var reloadButton = actionRow.add("button", undefined, L("button.reload"));
        reloadButton.alignment = "left";

        return {
            artboardInfoText: infoText,
            widthInput: widthField.input,
            heightInput: heightField.input,
            widthUnitText: widthField.unitText,
            heightUnitText: heightField.unitText,
            optimizeButton: optimizeButton,
            reloadButton: reloadButton
        };
    }

    /**
     * 「アートボード名と枠線」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildArtboardDisplayPanel(parentWindow) {
        var displayPanel = parentWindow.add("panel", undefined, L("panel.artboardDisplay"));
        setupPanel(displayPanel);

        var showNameCheckbox = displayPanel.add("checkbox", undefined, L("checkbox.showArtboardName"));

        /* 枠線サブパネル / Border sub-panel */
        var borderPanel = displayPanel.add("panel", undefined, L("panel.artboardBorder"));
        setupPanel(borderPanel, 8);

        var colorRow = borderPanel.add("group");
        setupRow(colorRow, "left");
        colorRow.add("statictext", undefined, labelWithColon("label.borderColor"));
        var borderColorList = colorRow.add("dropdownlist", undefined, buildBorderColorNames());

        var widthRow = borderPanel.add("group");
        setupRow(widthRow, "left");
        widthRow.add("statictext", undefined, labelWithColon("label.borderWidth"));
        var borderWidthRadios = [];
        for (var i = 0; i < BORDER_WIDTH_CHOICES.length; i++) {
            borderWidthRadios.push(widthRow.add("radiobutton", undefined, String(BORDER_WIDTH_CHOICES[i])));
        }

        /* プリセット（パネル最下部・左右中央）/ Presets (bottom of panel, centered) */
        var presetRow = displayPanel.add("group");
        setupRow(presetRow, "center");
        var presetRadios = [];
        for (var j = 0; j < PRESET_KEYS.length; j++) {
            presetRadios.push(presetRow.add("radiobutton", undefined, L("preset." + PRESET_KEYS[j])));
        }

        return {
            showNameCheckbox: showNameCheckbox,
            borderColorList: borderColorList,
            borderWidthRadios: borderWidthRadios,
            presetRadios: presetRadios
        };
    }

    /**
     * 「オプション」パネルを構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネル内のコントロール
     */
    function buildOptionsPanel(parentWindow) {
        var optionsPanel = parentWindow.add("panel", undefined, L("panel.options"));
        setupPanel(optionsPanel, 8);
        return {
            /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
            // printBleedCheckbox: optionsPanel.add("checkbox", undefined, L("checkbox.showPrintBleedAI")),
            moveLockedHiddenCheckbox: optionsPanel.add("checkbox", undefined, L("checkbox.moveLockedHidden"))
        };
    }

    /**
     * 下部のボタン行（カンバスカラー／ビデオ定規）を構築する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} 行内のボタン
     */
    function buildFooterRow(parentWindow) {
        var footerRow = parentWindow.add("group");
        setupRow(footerRow, "fill");

        var canvasColorButton = footerRow.add("button", undefined, L("button.canvasColor"));
        canvasColorButton.alignment = ["left", "center"];

        var footerSpacer = footerRow.add("group");
        footerSpacer.alignment = ["fill", "fill"];

        var videoRulerButton = footerRow.add("button", undefined, L("button.videoRuler"));
        videoRulerButton.alignment = ["right", "center"];

        return { canvasColorButton: canvasColorButton, videoRulerButton: videoRulerButton };
    }

    /**
     * パレット全体を構築する
     * @returns {object} ウィンドウと全コントロールをまとめたオブジェクト
     */
    function buildPalette() {
        var paletteWindow = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(paletteWindow);

        var ui = { paletteWindow: paletteWindow };
        var panels = [
            buildCurrentArtboardPanel(paletteWindow),
            buildArtboardDisplayPanel(paletteWindow),
            buildOptionsPanel(paletteWindow),
            buildFooterRow(paletteWindow)
        ];
        for (var i = 0; i < panels.length; i++) {
            for (var controlName in panels[i]) {
                if (panels[i].hasOwnProperty(controlName)) ui[controlName] = panels[i][controlName];
            }
        }
        return ui;
    }

    // =========================================
    // 即時反映：UI → 環境設定 / Immediate apply: UI to preferences
    // =========================================

    /**
     * 選択中の枠線の太さ（1〜4）を取得する
     * @param {object} ui - パレットのコントロール一式
     * @returns {number} 選択中の太さ
     */
    function getSelectedBorderWidth(ui) {
        for (var i = 0; i < ui.borderWidthRadios.length; i++) {
            if (ui.borderWidthRadios[i].value) return BORDER_WIDTH_CHOICES[i];
        }
        return BORDER_WIDTH_CHOICES[0];
    }

    /**
     * アートボード名・枠線カラー・枠線の太さを環境設定へ反映する
     * @param {object} ui - パレットのコントロール一式
     * @returns {void}
     */
    function applyArtboardDisplaySettings(ui) {
        prefs.setBooleanPreference("showArtboardLabelOnCanvas", ui.showNameCheckbox.value);
        var colorIndex = ui.borderColorList.selection ? ui.borderColorList.selection.index : BORDER_COLOR_BLACK_INDEX;
        var color = BORDER_COLOR_PRESETS[colorIndex];
        prefs.setRealPreference("ArtboardBBColorRed", color.r);
        prefs.setRealPreference("ArtboardBBColorGreen", color.g);
        prefs.setRealPreference("ArtboardBBColorBlue", color.b);
        prefs.setRealPreference("ArtboardBBWidth", getSelectedBorderWidth(ui));
        refreshArtboardDisplay();
    }

    // PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note
    // 「裁ち落としを印刷」生成AIボタンの表示設定を、書き込み・再描画・読み戻しまで
    // メインエンジンへ委譲する（書き込めなかった場合はチェックを実際の値へ戻す）
    // function applyPrintBleedWidgetSetting(ui) {
    //     var enabled = ui.printBleedCheckbox.value;
    //     delegateToMainEngine("" +
    //         "(function(){" +
    //         "app.preferences.setBooleanPreference('" + PRINT_BLEED_WIDGET_KEY + "'," + (enabled ? "true" : "false") + ");" +
    //         "try{if(app.documents.length){app.activeDocument.activate();app.redraw();" +
    //         "app.executeMenuCommand('zoomout');app.executeMenuCommand('zoomin');}}catch(e){}" +
    //         "return String(app.preferences.getBooleanPreference('" + PRINT_BLEED_WIDGET_KEY + "'));" +
    //         "})();",
    //         function (result) {
    //             if (result === "true" || result === "false") {
    //                 ui.printBleedCheckbox.value = (result === "true");
    //             }
    //         });
    // }

    /**
     * オプション（ロックまたは非表示オブジェクトを一緒に移動）を環境設定へ反映する
     * @param {object} ui - パレットのコントロール一式
     * @returns {void}
     */
    function applyOptionSettings(ui) {
        // applyPrintBleedWidgetSetting(ui); /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        prefs.setBooleanPreference("moveLockedAndHiddenArt", ui.moveLockedHiddenCheckbox.value);
    }

    // =========================================
    // 値反映：環境設定 → UI / Reflect preferences into the UI
    // =========================================

    /**
     * 枠線の太さのラジオボタンを選択する
     * @param {object} ui - パレットのコントロール一式
     * @param {number} borderWidth - 選択する太さ（1〜4）
     * @returns {void}
     */
    function selectBorderWidthRadio(ui, borderWidth) {
        for (var i = 0; i < ui.borderWidthRadios.length; i++) {
            ui.borderWidthRadios[i].value = (BORDER_WIDTH_CHOICES[i] === borderWidth);
        }
    }

    /**
     * プリセットの内容を UI へ適用する
     * @param {object} ui - パレットのコントロール一式
     * @param {string} presetKey - PRESETS のキー
     * @returns {void}
     */
    function applyPresetToUI(ui, presetKey) {
        var preset = PRESETS[presetKey];
        if (!preset) return;
        ui.showNameCheckbox.value = preset.showName;
        ui.borderColorList.selection = findBorderColorIndexByKey(preset.colorKey);
        selectBorderWidthRadio(ui, preset.borderWidth);
        // ui.printBleedCheckbox.value = preset.printBleed; /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        ui.moveLockedHiddenCheckbox.value = preset.moveLocked;
    }

    /**
     * 現在の UI 状態に一致するプリセットキーを返す
     * @param {object} ui - パレットのコントロール一式
     * @param {number} colorIndex - 現在のカラープリセット index
     * @param {number} borderWidth - 現在の枠線の太さ
     * @returns {string|null} 一致するプリセットキー（なければ null）
     */
    function detectPresetKey(ui, colorIndex, borderWidth) {
        for (var i = 0; i < PRESET_KEYS.length; i++) {
            var preset = PRESETS[PRESET_KEYS[i]];
            if (ui.showNameCheckbox.value === preset.showName &&
                colorIndex === findBorderColorIndexByKey(preset.colorKey) &&
                borderWidth === preset.borderWidth &&
                /* ui.printBleedCheckbox.value === preset.printBleed && */ /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
                ui.moveLockedHiddenCheckbox.value === preset.moveLocked) {
                return PRESET_KEYS[i];
            }
        }
        return null;
    }

    /**
     * 現在の環境設定を UI へ反映し、一致するプリセットのラジオボタンを選択する
     * @param {object} ui - パレットのコントロール一式
     * @returns {void}
     */
    function reflectPreferences(ui) {
        ui.showNameCheckbox.value = !!getBool("showArtboardLabelOnCanvas", false);

        var colorIndex = findClosestBorderColorIndex(
            getReal("ArtboardBBColorRed", 0.0),
            getReal("ArtboardBBColorGreen", 0.0),
            getReal("ArtboardBBColorBlue", 0.0)
        );
        ui.borderColorList.selection = colorIndex;

        var borderWidth = clamp(Math.round(getReal("ArtboardBBWidth", 1.0)), 1, BORDER_WIDTH_CHOICES.length);
        selectBorderWidthRadio(ui, borderWidth);

        // ui.printBleedCheckbox.value = !!getBool(PRINT_BLEED_WIDGET_KEY, false); /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        ui.moveLockedHiddenCheckbox.value = !!getBool("moveLockedAndHiddenArt", false);

        var matchedKey = detectPresetKey(ui, colorIndex, borderWidth);
        for (var i = 0; i < PRESET_KEYS.length; i++) {
            ui.presetRadios[i].value = (PRESET_KEYS[i] === matchedKey);
        }
    }

    // =========================================
    // 現在のアートボード / Current artboard
    // =========================================

    /**
     * 委譲結果（番号<|>名前<|>幅pt<|>高さpt）を UI へ反映する
     * @param {object} ui - パレットのコントロール一式
     * @param {string} result - メインエンジンからの結果文字列
     * @param {boolean} alertOnEmpty - 空結果（ドキュメントなし）でアラートを出すか
     * @returns {void}
     */
    function applyArtboardResult(ui, result, alertOnEmpty) {
        var unit = getRulerUnit();
        ui.widthUnitText.text = unit.label;
        ui.heightUnitText.text = unit.label;

        var fields = result ? result.split(ARTBOARD_FIELD_SEPARATOR) : [];
        if (fields.length < 4) {
            ui.artboardInfoText.text = "—";
            ui.widthInput.text = "";
            ui.heightInput.text = "";
            if (alertOnEmpty) alert(L("alert.noDocument"));
            return;
        }
        var separator = (currentLanguage === "ja") ? "：" : ": ";
        ui.artboardInfoText.text = "#" + fields[0] + separator + fields[1];
        ui.widthInput.text = pointToUnitText(parseFloat(fields[2]), unit);
        ui.heightInput.text = pointToUnitText(parseFloat(fields[3]), unit);
    }

    /**
     * アートボード操作をメインエンジンへ委譲し、結果を UI へ反映する
     * @param {object} ui - パレットのコントロール一式
     * @param {string} operation - "read" | "round" | "resize"
     * @param {number} widthPoint - resize 時の幅（pt）
     * @param {number} heightPoint - resize 時の高さ（pt）
     * @param {boolean} alertOnEmpty - 空結果でアラートを出すか
     * @returns {void}
     */
    function runArtboardOperation(ui, operation, widthPoint, heightPoint, alertOnEmpty) {
        delegateToMainEngine(buildArtboardBody(operation, widthPoint, heightPoint), function (result) {
            applyArtboardResult(ui, result, alertOnEmpty);
        });
    }

    /**
     * 現在のアートボード情報を取得して UI へ反映する
     * @param {object} ui - パレットのコントロール一式
     * @returns {void}
     */
    function refreshArtboardInfo(ui) {
        runArtboardOperation(ui, "read", 0, 0, false);
    }

    /**
     * 幅・高さの入力値でアクティブアートボードをリサイズする（不正値は現在値へ戻す）
     * @param {object} ui - パレットのコントロール一式
     * @returns {void}
     */
    function resizeArtboardFromFields(ui) {
        var unit = getRulerUnit();
        var width = parseFloat(ui.widthInput.text);
        var height = parseFloat(ui.heightInput.text);
        if (isNaN(width) || isNaN(height) || width <= 0 || height <= 0) {
            refreshArtboardInfo(ui);
            return;
        }
        runArtboardOperation(ui, "resize", width * unit.factor, height * unit.factor, true);
    }

    // =========================================
    // イベント設定 / Event wiring
    // =========================================

    /**
     * パレットのイベントハンドラを設定する
     * @param {object} ui - パレットのコントロール一式
     * @returns {void}
     */
    function wirePaletteEvents(ui) {
        var i;

        /* プリセット：UI へ展開してから環境設定へ反映 / Presets: expand into the UI, then write */
        for (i = 0; i < PRESET_KEYS.length; i++) {
            ui.presetRadios[i].presetKey = PRESET_KEYS[i];
            ui.presetRadios[i].onClick = function () {
                applyPresetToUI(ui, this.presetKey);
                applyArtboardDisplaySettings(ui);
                applyOptionSettings(ui);
            };
        }

        ui.showNameCheckbox.onClick = function () { applyArtboardDisplaySettings(ui); };
        ui.borderColorList.onChange = function () { applyArtboardDisplaySettings(ui); };
        for (i = 0; i < ui.borderWidthRadios.length; i++) {
            ui.borderWidthRadios[i].onClick = function () { applyArtboardDisplaySettings(ui); };
        }

        // ui.printBleedCheckbox.onClick = function () { applyOptionSettings(ui); }; /* PRINT_BLEED_WIDGET を参照 / See the PRINT_BLEED_WIDGET note */
        ui.moveLockedHiddenCheckbox.onClick = function () { applyOptionSettings(ui); };

        /* ピクセルグリッドに最適化：XYWH を整数値へ丸める / Optimize: round XYWH to integers */
        ui.optimizeButton.onClick = function () { runArtboardOperation(ui, "round", 0, 0, true); };
        ui.reloadButton.onClick = function () { refreshArtboardInfo(ui); };

        /* 幅・高さの確定でアートボードをリサイズ / Resize the artboard when width/height are committed */
        ui.widthInput.onChange = function () { resizeArtboardFromFields(ui); };
        ui.heightInput.onChange = function () { resizeArtboardFromFields(ui); };

        /* カンバスカラーの変更（uiCanvasIsWhite: 1=白 / 0=グレー）/ Toggle the canvas color */
        ui.canvasColorButton.onClick = function () {
            prefs.setIntegerPreference("uiCanvasIsWhite", getInt("uiCanvasIsWhite", 0) === 1 ? 0 : 1);
            refreshArtboardDisplay();
        };
        ui.videoRulerButton.onClick = function () { runMenuCommand("videoruler"); };

        /* アクティブ時に Esc で閉じる / Close on Esc while active */
        ui.paletteWindow.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") ui.paletteWindow.close();
        });

        /* 再アクティブ時：外部変更とアートボードの切り替えに追従 / On re-activate: follow external changes */
        ui.paletteWindow.onActivate = function () {
            reflectPreferences(ui);
            refreshArtboardInfo(ui);
        };

        /* 常駐エンジンの参照は閉じたらクリア / Clear the persistent-engine reference on close */
        ui.paletteWindow.onClose = function () {
            $.global[PALETTE_GLOBAL_KEY] = null;
        };
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * すでに開いているパレットを返す（無効な参照・表示前の残骸は null）
     * 表示中のものだけを有効と見なす。構築途中でエラーになったウィンドウを掴むと、
     * イベント未配線のパレットを開き続けることになるため
     * @returns {Window|null} 既存のパレット
     */
    function getExistingPalette() {
        try {
            var palette = $.global[PALETTE_GLOBAL_KEY];
            return (palette && palette.visible === true) ? palette : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * パレットを構築して表示する（重複起動時は既存のパレットを前面に出す）
     * @returns {void}
     */
    function main() {
        var existingPalette = getExistingPalette();
        if (existingPalette) {
            existingPalette.show();
            return;
        }

        var ui = buildPalette();
        wirePaletteEvents(ui);
        reflectPreferences(ui);
        refreshArtboardInfo(ui);

        ui.paletteWindow.center();
        ui.paletteWindow.show();

        /* 構築と表示がすべて通ってから参照を保持する（途中で失敗した窓を残さない）
           Store the reference only after everything succeeded, so a half-built window is never kept */
        $.global[PALETTE_GLOBAL_KEY] = ui.paletteWindow;

        /* レイアウト確定後にボタン高さを 2px 詰める / Trim the button heights by 2px after layout */
        trimButtonHeight(ui.optimizeButton, 2);
        trimButtonHeight(ui.reloadButton, 2);
    }

    main();

}());
