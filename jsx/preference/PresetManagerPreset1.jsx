#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PresetManager の［プリセット1］と同じ設定一式を、Illustratorの環境設定へまとめて書き込みます。
ダイアログは表示せず、実行するとその場で反映されます。

詳細は README を参照してください。

### Overview

Writes the same set of values as PresetManager's Preset 1 into the Illustrator preferences.
There is no dialog; running it applies everything straight away.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManagerPreset1";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-09";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerPreset1.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerPreset1.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 設定値の定義 / Preference value definitions
    // =========================================

    /* アートボードのハイライトカラープリセット（RGB 0..1）。index を PRESET_STATE_1 で指定する */
    /* Artboard highlight color presets (RGB 0..1); PRESET_STATE_1 selects one by index */
    var ARTBOARD_COLOR_PRESETS = [
        { red: 0.29, green: 0.52, blue: 1.0 }, /* 0: ライトブルー / Light Blue */
        { red: 1.0, green: 0.29, blue: 0.29 }, /* 1: サーモンピンク / Light Red */
        { red: 0.0, green: 0.65, blue: 0.31 }, /* 2: グリーン / Green */
        { red: 0.0, green: 0.45, blue: 0.78 }, /* 3: ミディアムブルー / Medium Blue */
        { red: 1.0, green: 0.0, blue: 1.0 },   /* 4: マゼンタ / Magenta */
        { red: 0.0, green: 1.0, blue: 1.0 },   /* 5: シアン / Cyan */
        { red: 1.0, green: 1.0, blue: 1.0 },   /* 6: ホワイト / White */
        { red: 0.0, green: 0.0, blue: 0.0 },   /* 7: ブラック / Black */
        { red: 1.0, green: 1.0, blue: 0.0 }    /* 8: イエロー / Yellow */
    ];

    /* ガイドカラーの2択（RGB 0..1）/ The two selectable guide colors (RGB 0..1) */
    var GUIDE_COLOR_CYAN = { red: 0.0, green: 1.0, blue: 1.0 };
    var GUIDE_COLOR_LIGHT_BLUE = { red: 0.29, green: 0.52, blue: 1.0 };

    /* ガイドスタイルの値 / Guide style values */
    var GUIDE_STYLE_LINES = 0;
    var GUIDE_STYLE_DOTS = 1;

    /* スマートガイドの「スナップを制限」の値 / "Restrict snapping" values for Smart Guides */
    var SNAP_RANGE_CANVAS = 0;          /* カンバス全体 / Entire canvas */
    var SNAP_RANGE_ACTIVE_ARTBOARD = 1; /* アクティブなアートボードの内容 / Active artboard content */

    /* 「リンクを更新」の値 / "Update Links" values */
    var UPDATE_LINKS_AUTO = 0;
    var UPDATE_LINKS_MANUAL = 1;
    var UPDATE_LINKS_ASK = 2;

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 数値項目の許容範囲と既定値。範囲外の値は書き込み時に補正する */
    /* Allowed range and default for the numeric fields; out-of-range values are clamped on write */
    var NUMERIC_VALUE_RULES = {
        recentFonts: { min: 1, max: 30, defaultValue: 15 },     /* 表示数 1〜30 / Visible count 1-30 */
        historyStates: { min: 1, max: 1000, defaultValue: 100 } /* 想定範囲 1〜1000 / Expected range 1-1000 */
    };

    /* 書き込む設定一式（PresetManager の［プリセット1］と同じ内容）。値を変えたいときはここを編集する */
    /* 各行のコメントは「Illustrator の既定値 → このスクリプトが書き込む値」。既定値は PresetManager の［デフォルト］に準拠 */
    /* The set of values to write (same as [Preset 1] in PresetManager); each comment reads "Illustrator default -> value written here" */
    var PRESET_STATE_1 = {
        /* 一般 / General */
        richToolTips: false,              /* 詳細なツールヒントを表示 / Show Rich Tool Tips（既定 true → false）*/
        homeScreen: false,                /* 「ホーム画面」を表示 / Show the Home Screen（既定 true → false）*/
        legacyNewDoc: true,               /* 以前の「新規ドキュメント」インターフェイス / Legacy File > New（既定 false → true）*/
        printBleedWidget: false,          /* 「裁ち落としを印刷」生成AIボタン / Print Bleed AI buttons（既定 true → false）*/
        /* 選択範囲・アンカー表示 / Selection & Anchor Display */
        zoomToSelection: false,           /* 選択範囲へズーム / Zoom to Selection（既定 true → false）*/
        anchorSize: 7,                    /* アンカーポイントのサイズ 5/7/9/11 / Anchor Point Size（既定 5 → 7）*/
        /* パスに制限 / Limit to Path */
        objectPathOnly: false,            /* オブジェクトの選択範囲をパスに制限 / Object Selection by Path Only（既定 false のまま）*/
        textPathOnly: false,              /* テキストオブジェクトの選択範囲をパスに制限 / Type Object Selection by Path Only（既定 false のまま）*/
        /* アートボード / Artboard */
        moveLockedArt: true,              /* ロックまたは非表示オブジェクトを一緒に移動 / Move Locked and Hidden Artwork（既定 false → true）*/
        showArtboardName: false,          /* アートボード名を表示 / Show Artboard Name（既定 true → false）*/
        artboardColorIndex: 7,            /* ハイライトのカラー / Highlight Color（既定 0 ライトブルー → 7 ブラック）*/
        artboardStrokeWidth: 2,           /* ストロークの幅 1〜4 / Stroke Width（既定 1 → 2）*/
        /* テキスト / Text */
        autoSizeAreaText: true,           /* 新規エリア内文字の自動サイズ調整 / Auto Size New Area Type（既定 false → true）*/
        recentFontsEnabled: true,         /* 最近使用したフォントを表示 / Show recent fonts（既定 true のまま）*/
        recentFontsCount: 15,             /* 最近使用したフォントの表示数 1〜30 / Number of Recent Fonts（既定 10 → 15）*/
        missingGlyphProtection: false,    /* 見つからない字形の保護 / Enable Missing Glyph Protection（既定 true → false）*/
        alternateGlyph: false,            /* 選択された文字の異体字を表示 / Show Character Alternates（既定 true → false）*/
        /* ユーザーインターフェイス / User Interface */
        canvasWhite: true,                /* カンバスカラー / Canvas Color（既定 false UIに合わせる → true ホワイト）*/
        /* ガイド / Guides */
        guideColorIsLightBlue: true,      /* ガイドのカラー / Guide Color（既定 false シアン → true ライトブルー）*/
        guideStyleIsDots: false,          /* ガイドのスタイル / Guide Style（既定 false ライン のまま）*/
        /* スマートガイド：表示オプション / Smart Guides: display options */
        objectHighlighting: false,        /* オブジェクトのハイライト表示 / Object Highlighting（既定 true → false）*/
        /* スマートガイド：スナップを制限 / Smart Guides: restrict snapping */
        snapRange: SNAP_RANGE_CANVAS,     /* スナップ範囲 / Snap range（既定 0 カンバス全体 のまま）*/
        snapToIsolatedObjects: true,      /* 編集モードのオブジェクトにスナップ / Snap to isolated objects（既定 false → true）*/
        snapTolerance: 6,                 /* 許容値 / Snapping tolerance（既定 未確認 → 6）*/
        /* 詳細設定・グリッドにスナップ / Advanced & snap to grid */
        snapToGrid: true,                 /* グリッドに強制スナップ / Snap to grid（既定 true のまま）*/
        showSnapToGridGuides: true,       /* グリッドにスナップするときにガイドを表示 / Show visual guides（既定 true のまま）*/
        snapToPointTolerance: 1,          /* ポイントにスナップ許容値 / Snap to point tolerance（既定 未確認 → 1）*/
        /* パフォーマンス / Performance */
        animatedZoom: false,              /* アニメーションズーム / Animated Zoom（既定 true → false）*/
        historyStates: 50,                /* ヒストリー数 1〜1000 / History States（既定 100 → 50）*/
        realTimeDrawing: false,           /* リアルタイムの描画と編集 / Real-Time Drawing and Editing（既定 true → false）*/
        /* ファイル管理 / File Management */
        editOriginalSystemDefault: true,  /* 「オリジナルの編集」にシステムデフォルトを使用 / Use System Defaults（既定 false → true）*/
        autoActivateFonts: true,          /* Adobe Fonts を自動アクティベート / Auto-activate Adobe Fonts（既定 false → true）*/
        saveToCloud: false,               /* ファイルの保存先 / Save Location（既定 true クラウド → false コンピューター）*/
        updateLinks: UPDATE_LINKS_AUTO,   /* リンクを更新 / Update Links（既定 2 確認 → 0 自動）*/
        /* クリップボードの処理 / Clipboard Handling */
        includeSvgCode: true              /* SVGコードを含める / Include SVG Code（既定 false → true）*/
    };

    // =========================================
    // 環境設定キーの対応表 / Preference key bindings
    // =========================================

    /* 1項目=1キーで書けるキーの一覧。key=環境設定キー / presetField=PRESET_STATE_1 の項目名 */
    /* valueType: "bool"=真偽値 / "boolInt"=ON:1・OFF:0 の整数 / "invertedInt"=ON:0・OFF:1 の整数 / "int"=数値そのまま */
    /* Keys writable one-to-one; valueType selects how the preset value is stored */
    var PREFERENCE_BINDINGS = [
        { key: "showRichToolTips", presetField: "richToolTips", valueType: "bool" },
        { key: "Hello/ShowHomeScreenWS", presetField: "homeScreen", valueType: "bool" },
        { key: "Hello/NewDoc", presetField: "legacyNewDoc", valueType: "bool" },
        { key: "enablePrintBleedWidget", presetField: "printBleedWidget", valueType: "bool" },
        { key: "zoomToSelection", presetField: "zoomToSelection", valueType: "bool" },
        { key: "moveLockedAndHiddenArt", presetField: "moveLockedArt", valueType: "bool" },
        { key: "showArtboardLabelOnCanvas", presetField: "showArtboardName", valueType: "bool" },
        { key: "text/autoSizing", presetField: "autoSizeAreaText", valueType: "bool" },
        { key: "text/doFontLocking", presetField: "missingGlyphProtection", valueType: "bool" },
        { key: "text/enableAlternateGlyph", presetField: "alternateGlyph", valueType: "bool" },
        /* スマートガイド：表示オプション / Smart Guides: display options */
        { key: "smartGuides/showObjectHighlighting", presetField: "objectHighlighting", valueType: "boolInt" },
        /* スマートガイド：スナップを制限 / Smart Guides: restrict snapping */
        { key: "smartGuides/snapToActiveArtboardContent", presetField: "snapRange", valueType: "int" },
        { key: "smartGuides/snapToIsolatedObjects", presetField: "snapToIsolatedObjects", valueType: "boolInt" },
        { key: "tolerance", presetField: "snapTolerance", valueType: "int" },
        /* 詳細設定・グリッドにスナップ / Advanced & snap to grid */
        { key: "fSnapToGrid", presetField: "snapToGrid", valueType: "boolInt" },
        { key: "showVisualGuidesForSnapToGrid", presetField: "showSnapToGridGuides", valueType: "boolInt" },
        { key: "snappingTolerance", presetField: "snapToPointTolerance", valueType: "int" },
        { key: "Performance/AnimZoom", presetField: "animatedZoom", valueType: "bool" },
        { key: "LiveEdit_State_Machine", presetField: "realTimeDrawing", valueType: "bool" },
        { key: "useSysDefEdit", presetField: "editOriginalSystemDefault", valueType: "bool" },
        { key: "AutoActivateMissingFont", presetField: "autoActivateFonts", valueType: "bool" },
        { key: "plugin/FileClipboard/copySVGCode", presetField: "includeSvgCode", valueType: "bool" },
        { key: "hitShapeOnPreview", presetField: "objectPathOnly", valueType: "invertedInt" },
        { key: "hitTypeShapeOnPreview", presetField: "textPathOnly", valueType: "invertedInt" }
    ];

    // =========================================
    // 環境設定の書き込み / Preference writers
    // =========================================

    /**
     * 例外を握りつぶして処理を続行する共通ラッパー（このスクリプト唯一の try）
     * Shared guard that swallows exceptions so the run keeps going (the only try in the script)
     * @param {function} operation - 実行する処理
     * @param {string} actionName - ログに出す処理名
     * @param {string} targetDetail - ログに出す対象（キー名など）
     * @returns {*} operation の戻り値、失敗時は undefined
     */
    function runSafely(operation, actionName, targetDetail) {
        try {
            return operation();
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] " + actionName + ": " + targetDetail + " / " + e);
        }
    }

    /**
     * 整数の環境設定を書き込む（失敗しても処理を止めない）/ Write an integer preference (never throws)
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} preferenceValue - 書き込む値
     * @returns {void}
     */
    function safeSetInt(preferenceKey, preferenceValue) {
        runSafely(function () {
            app.preferences.setIntegerPreference(preferenceKey, preferenceValue);
        }, "safeSetInt failed", preferenceKey + " = " + preferenceValue);
    }

    /**
     * 真偽値の環境設定を書き込む / Write a boolean preference (never throws)
     * @param {string} preferenceKey - 環境設定キー
     * @param {boolean} preferenceValue - 書き込む値
     * @returns {void}
     */
    function safeSetBool(preferenceKey, preferenceValue) {
        runSafely(function () {
            app.preferences.setBooleanPreference(preferenceKey, preferenceValue);
        }, "safeSetBool failed", preferenceKey + " = " + preferenceValue);
    }

    /**
     * 実数の環境設定を書き込む / Write a real-number preference (never throws)
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} preferenceValue - 書き込む値
     * @returns {void}
     */
    function safeSetReal(preferenceKey, preferenceValue) {
        runSafely(function () {
            app.preferences.setRealPreference(preferenceKey, preferenceValue);
        }, "safeSetReal failed", preferenceKey + " = " + preferenceValue);
    }

    // =========================================
    // 数値ユーティリティ / Numeric utilities
    // =========================================

    /**
     * @typedef {object} NumericRule
     * @property {number} min - 許容する最小値
     * @property {number} max - 許容する最大値
     * @property {number} defaultValue - 数値として解釈できないときの既定値
     */

    /**
     * 値を min〜max に収める / Clamp a value into the min-max range
     * @param {number} targetValue - 対象の値
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @returns {number} 範囲内に収めた値
     */
    function clamp(targetValue, minValue, maxValue) {
        return Math.max(minValue, Math.min(maxValue, targetValue));
    }

    /**
     * ルール（min/max/defaultValue）に沿って整数へ補正 / Normalize a value against a min/max/default rule
     * @param {string|number} rawValue - 対象の値
     * @param {NumericRule} numericRule - 適用するルール
     * @returns {number} 補正後の整数
     */
    function clampIntToRule(rawValue, numericRule) {
        var parsedInt = parseInt(rawValue, 10);
        if (isNaN(parsedInt)) return numericRule.defaultValue;
        return clamp(parsedInt, numericRule.min, numericRule.max);
    }

    // =========================================
    // 設定値のヘルパー / Preference value helpers
    // =========================================

    /**
     * ガイドカラーを書き込む / Write the guide color
     * @param {{red: number, green: number, blue: number}} guideColor - 書き込むRGB（0〜1）
     * @returns {void}
     */
    function writeGuideColor(guideColor) {
        safeSetReal("Guide/Color/red", guideColor.red);
        safeSetReal("Guide/Color/green", guideColor.green);
        safeSetReal("Guide/Color/blue", guideColor.blue);
    }

    /**
     * アートボードのハイライトカラーを書き込む / Write the artboard highlight color
     * @param {number} colorIndex - ARTBOARD_COLOR_PRESETS の index
     * @returns {void}
     */
    function writeArtboardColor(colorIndex) {
        var artboardColor = ARTBOARD_COLOR_PRESETS[colorIndex] || ARTBOARD_COLOR_PRESETS[0];
        safeSetReal("ArtboardBBColorRed", artboardColor.red);
        safeSetReal("ArtboardBBColorGreen", artboardColor.green);
        safeSetReal("ArtboardBBColorBlue", artboardColor.blue);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 対応表のキーをまとめて書き込む / Write every key listed in PREFERENCE_BINDINGS
     * @param {object} presetState - PRESET_STATE_1 と同じ形の設定一式
     * @returns {void}
     */
    function applyPreferenceBindings(presetState) {
        for (var i = 0; i < PREFERENCE_BINDINGS.length; i++) {
            var binding = PREFERENCE_BINDINGS[i];
            var presetValue = presetState[binding.presetField];
            if (binding.valueType === "bool") {
                safeSetBool(binding.key, !!presetValue);
            } else if (binding.valueType === "boolInt") {
                safeSetInt(binding.key, presetValue ? 1 : 0);
            } else if (binding.valueType === "invertedInt") {
                safeSetInt(binding.key, presetValue ? 0 : 1); /* 0がON / 0 means ON */
            } else if (binding.valueType === "int") {
                safeSetInt(binding.key, presetValue);
            }
        }
    }

    /**
     * 設定一式を環境設定へ書き込む / Write a whole preset into the preferences
     * @param {object} presetState - PRESET_STATE_1 と同じ形の設定一式
     * @returns {void}
     */
    function applyPreset(presetState) {
        /* 対応表で扱えるキー / Keys covered by the binding table */
        applyPreferenceBindings(presetState);

        /* テキスト（0＝一覧を非表示）/ Text (0 = hide the list) */
        safeSetInt("text/recentFontMenu/showNEntries",
            presetState.recentFontsEnabled ? clampIntToRule(presetState.recentFontsCount, NUMERIC_VALUE_RULES.recentFonts) : 0);

        /* 選択範囲・アンカー表示 / Selection & anchor display */
        safeSetInt("anchorSizePref", presetState.anchorSize);

        /* ユーザーインターフェイス / User interface */
        safeSetInt("uiCanvasIsWhite", presetState.canvasWhite ? 1 : 0);

        /* ガイド / Guides */
        writeGuideColor(presetState.guideColorIsLightBlue ? GUIDE_COLOR_LIGHT_BLUE : GUIDE_COLOR_CYAN);
        safeSetInt("Guide/Style", presetState.guideStyleIsDots ? GUIDE_STYLE_DOTS : GUIDE_STYLE_LINES);

        /* パフォーマンス / Performance */
        safeSetInt("maximumUndoDepth", clampIntToRule(presetState.historyStates, NUMERIC_VALUE_RULES.historyStates));

        /* ファイル管理・クリップボード / File management & clipboard */
        safeSetBool("AdobeSaveAsCloudDocumentPreference", !!presetState.saveToCloud);
        safeSetInt("plugin/FileClipboard/linkoptions", presetState.updateLinks);

        /* アートボード / Artboard */
        writeArtboardColor(presetState.artboardColorIndex);
        safeSetReal("ArtboardBBWidth", presetState.artboardStrokeWidth);
    }

    /**
     * 環境設定の変更後に画面を強制再描画（redraw だけでは反映されないため）
     * Force a redraw after preference changes (redraw alone is unreliable)
     * ドキュメントが開いていないときは何もしない / Does nothing when no document is open
     * @returns {void}
     */
    function forceScreenRefresh() {
        if (app.documents.length === 0) return;
        runSafely(function () {
            app.executeMenuCommand("zoomout");
            app.executeMenuCommand("zoomin");
        }, "forceScreenRefresh failed", "zoomout/zoomin");
    }

    /**
     * ［プリセット1］を環境設定へ書き込み、画面を更新する / Write [Preset 1] and refresh the screen
     * @returns {void}
     */
    function main() {
        applyPreset(PRESET_STATE_1);
        forceScreenRefresh();
    }

    main();
})();
