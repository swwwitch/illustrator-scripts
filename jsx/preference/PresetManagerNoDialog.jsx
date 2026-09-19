#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

あらかじめ決めた環境設定一式を、ダイアログを表示せずにまとめて適用します。
冒頭の ACTIVE_PRESET で、適用するプリセットを切り替えます。

詳細は README を参照してください。

### Overview

Applies a fixed set of Illustrator preferences at once, without showing a dialog.
ACTIVE_PRESET at the top selects which preset gets written.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManagerNoDialog";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerNoDialog.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerNoDialog.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 適用するプリセット / Preset to apply
    // =========================================

    /* 実行時に書き込むプリセット。PRESET_STATES のキーから選ぶ */
    /* Which preset gets written; pick one of the PRESET_STATES keys */
    var ACTIVE_PRESET = "full"; /* "minimal" | "full" | "preset1" */

    // =========================================
    // 設定値の定義 / Preference value definitions
    // =========================================

    /* アートボードのハイライトカラープリセット（RGB 0..1）。index を artboardColorIndex で指定する */
    /* Artboard highlight color presets (RGB 0..1); artboardColorIndex selects one by index */
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

    /* スマートガイドの「スナップを制限」の値 / "Restrict snapping" values for Smart Guides */
    var SNAP_RANGE_CANVAS = 0;          /* カンバス全体 / Entire canvas */
    var SNAP_RANGE_ACTIVE_ARTBOARD = 1; /* アクティブなアートボードの内容 / Active artboard content */

    /* 「リンクを更新」の値 / "Update Links" values */
    var UPDATE_LINKS_AUTO = 0;
    var UPDATE_LINKS_MANUAL = 1;
    var UPDATE_LINKS_ASK = 2;

    /* 定規の単位 / Ruler units */
    var RULER_TYPE_MM = 1;

    /* ［シェイプ形成ツール］の［次のカラー］/ Shape Builder "Pick Color From" */
    var PAINT_FILLS_OBJECT = 0;  /* オブジェクト / Artwork */
    var PAINT_FILLS_SWATCH = 1;  /* スウォッチ / Color swatches */

    /* ユーザーインターフェイスの明るさ（連続値ではなく離散プリセット）/ UI brightness; discrete presets, not a continuous scale */
    var UI_BRIGHTNESS_DARK = 0;
    var UI_BRIGHTNESS_MEDIUM_DARK = 0.5;
    var UI_BRIGHTNESS_MEDIUM_LIGHT = 0.50999999046326;
    var UI_BRIGHTNESS_LIGHT = 1;

    /* キー増加の量（pt）/ Keyboard increment in points */
    var CURSOR_KEY_LENGTH_0_1MM = 0.2835; /* 0.1mm */

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 数値項目の許容範囲と既定値。範囲外の値は書き込み時に補正する */
    /* Allowed range and default for the numeric fields; out-of-range values are clamped on write */
    var NUMERIC_VALUE_RULES = {
        recentFonts: { min: 1, max: 30, defaultValue: 15 },     /* 表示数 1〜30 / Visible count 1-30 */
        historyStates: { min: 1, max: 1000, defaultValue: 100 } /* 想定範囲 1〜1000 / Expected range 1-1000 */
    };

    /* 書き込む設定一式。項目を省くとそのキーには触れない（Illustratorの現在値が残る）*/
    /* The presets; an omitted field leaves that preference untouched */
    var PRESET_STATES = {

        /* 定番だけを手早く。項目数を絞った版 / The short list: just the staples */
        minimal: {
            /* 一般 / General */
            richToolTips: false,              /* 詳細なツールヒントを表示 / Show Rich Tool Tips */
            homeScreen: false,                /* 「ホーム画面」を表示 / Show the Home Screen */
            legacyNewDoc: true,               /* 以前の「新規ドキュメント」インターフェイス / Legacy File > New */
            printBleedWidget: false,          /* 「裁ち落としを印刷」生成AIボタン / Print Bleed AI buttons */
            contextualTaskBar: false,         /* コンテキストタスクバー / Contextual Task Bar */
            pathfinderWarning: false,         /* パスファインダー適用時のアラート / Pathfinder group warning */
            /* 選択範囲・アンカー表示 / Selection & Anchor Display */
            zoomToSelection: false,           /* 選択範囲へズーム / Zoom to Selection */
            /* アートボード / Artboard */
            moveLockedArt: true,              /* ロックまたは非表示オブジェクトを一緒に移動 / Move Locked and Hidden Artwork */
            /* テキスト / Text */
            autoSizeAreaText: true,           /* 新規エリア内文字の自動サイズ調整 / Auto Size New Area Type */
            recentFontsEnabled: false,        /* 最近使用したフォントを表示 / Show recent fonts */
            missingGlyphProtection: false,    /* 見つからない字形の保護 / Enable Missing Glyph Protection */
            alternateGlyph: false,            /* 選択された文字の異体字を表示 / Show Character Alternates */
            /* ユーザーインターフェイス / User Interface */
            canvasWhite: true,                /* カンバスカラー / Canvas Color */
            /* パフォーマンス / Performance */
            animatedZoom: false,              /* アニメーションズーム / Animated Zoom */
            historyStates: 50,                /* ヒストリー数 / History States */
            realTimeDrawing: true,            /* リアルタイムの描画と編集 / Real-Time Drawing and Editing */
            /* ファイル管理 / File Management */
            autoActivateFonts: true,          /* Adobe Fonts を自動アクティベート / Auto-activate Adobe Fonts */
            saveToCloud: false,               /* ファイルの保存先 / Save Location */
            /* 再起動が必要 / Requires a restart */
            shapeBuilderPaintFills: PAINT_FILLS_OBJECT /* ［次のカラー］/ Shape Builder pick color from */
        },

        /* 環境設定をひととおり押さえた版 / The long list: a full sweep of the panels */
        full: {
            /* 一般 / General */
            cursorKeyLength: CURSOR_KEY_LENGTH_0_1MM, /* キー増加 / Keyboard Increment */
            richToolTips: false,              /* 詳細なツールヒントを表示 / Show Rich Tool Tips */
            homeScreen: false,                /* 「ホーム画面」を表示 / Show the Home Screen */
            legacyNewDoc: true,               /* 以前の「新規ドキュメント」インターフェイス / Legacy File > New */
            printBleedWidget: false,          /* 「裁ち落としを印刷」生成AIボタン / Print Bleed AI buttons */
            contextualTaskBar: false,         /* コンテキストタスクバー / Contextual Task Bar */
            pathfinderWarning: false,         /* パスファインダー適用時のアラート / Pathfinder group warning */
            /* 選択範囲・アンカー表示 / Selection & Anchor Display */
            zoomToSelection: false,           /* 選択範囲へズーム / Zoom to Selection */
            objectPathOnly: false,            /* オブジェクトの選択範囲をパスに制限 / Object Selection by Path Only */
            textPathOnly: false,              /* テキストオブジェクトの選択範囲をパスに制限 / Type Object Selection by Path Only */
            /* アートボード / Artboard */
            moveLockedArt: true,              /* ロックまたは非表示オブジェクトを一緒に移動 / Move Locked and Hidden Artwork */
            /* テキスト / Text */
            autoSizeAreaText: true,           /* 新規エリア内文字の自動サイズ調整 / Auto Size New Area Type */
            fillSampleText: false,            /* 新規テキストオブジェクトにサンプルテキストを割り付け / Fill New Type Objects With Placeholder Text */
            recentFontsEnabled: false,        /* 最近使用したフォントを表示 / Show recent fonts */
            missingGlyphProtection: false,    /* 見つからない字形の保護 / Enable Missing Glyph Protection */
            alternateGlyph: false,            /* 選択された文字の異体字を表示 / Show Character Alternates */
            listAutoDetection: false,         /* 入力中に箇条書きリストのレベルを自動更新 / Auto update bullet list level */
            /* スマートガイド / Smart Guides */
            objectHighlighting: false,        /* オブジェクトのハイライト表示 / Object Highlighting */
            /* ユーザーインターフェイス / User Interface */
            canvasWhite: true,                /* カンバスカラー / Canvas Color */
            uiBrightness: UI_BRIGHTNESS_LIGHT,/* ユーザーインターフェイスの明るさ / Brightness */
            /* パフォーマンス / Performance */
            animatedZoom: false,              /* アニメーションズーム / Animated Zoom */
            historyStates: 50,                /* ヒストリー数 / History States */
            realTimeDrawing: true,            /* リアルタイムの描画と編集 / Real-Time Drawing and Editing */
            /* ファイル管理 / File Management */
            editOriginalSystemDefault: true,  /* 「オリジナルの編集」にシステムデフォルトを使用 / Use System Defaults */
            autoActivateFonts: true,          /* Adobe Fonts を自動アクティベート / Auto-activate Adobe Fonts */
            saveToCloud: false,               /* ファイルの保存先 / Save Location */
            /* ブラックのアピアランス / Black Appearance */
            accurateBlackOnscreen: true,      /* すべてのブラックを正確に表示（スクリーン）/ Display All Blacks Accurately (screen) */
            accurateBlackExport: true,        /* すべてのブラックを正確に出力（プリント／書き出し）/ Output All Blacks Accurately (print & export) */
            /* 再起動が必要 / Requires a restart */
            rulerType: RULER_TYPE_MM,         /* 定規の単位 / Ruler units */
            shapeBuilderPaintFills: PAINT_FILLS_OBJECT /* ［次のカラー］/ Shape Builder pick color from */
        },

        /* PresetManager の［プリセット1］と同じ内容 / Mirrors [Preset 1] in PresetManager */
        preset1: {
            /* 一般 / General */
            richToolTips: false,              /* 詳細なツールヒントを表示 / Show Rich Tool Tips */
            homeScreen: false,                /* 「ホーム画面」を表示 / Show the Home Screen */
            legacyNewDoc: true,               /* 以前の「新規ドキュメント」インターフェイス / Legacy File > New */
            printBleedWidget: false,          /* 「裁ち落としを印刷」生成AIボタン / Print Bleed AI buttons */
            /* 選択範囲・アンカー表示 / Selection & Anchor Display */
            zoomToSelection: false,           /* 選択範囲へズーム / Zoom to Selection */
            anchorSize: 7,                    /* アンカーポイントのサイズ 5/7/9/11 / Anchor Point Size */
            objectPathOnly: false,            /* オブジェクトの選択範囲をパスに制限 / Object Selection by Path Only */
            textPathOnly: false,              /* テキストオブジェクトの選択範囲をパスに制限 / Type Object Selection by Path Only */
            /* アートボード / Artboard */
            moveLockedArt: true,              /* ロックまたは非表示オブジェクトを一緒に移動 / Move Locked and Hidden Artwork */
            showArtboardName: false,          /* アートボード名を表示 / Show Artboard Name */
            artboardColorIndex: 7,            /* ハイライトのカラー / Highlight Color */
            artboardStrokeWidth: 2,           /* ストロークの幅 1〜4 / Stroke Width */
            /* テキスト / Text */
            autoSizeAreaText: true,           /* 新規エリア内文字の自動サイズ調整 / Auto Size New Area Type */
            recentFontsEnabled: true,         /* 最近使用したフォントを表示 / Show recent fonts */
            recentFontsCount: 15,             /* 最近使用したフォントの表示数 1〜30 / Number of Recent Fonts */
            missingGlyphProtection: false,    /* 見つからない字形の保護 / Enable Missing Glyph Protection */
            alternateGlyph: false,            /* 選択された文字の異体字を表示 / Show Character Alternates */
            /* ユーザーインターフェイス / User Interface */
            canvasWhite: true,                /* カンバスカラー / Canvas Color */
            /* ガイド / Guides */
            guideColorIsLightBlue: true,      /* ガイドのカラー / Guide Color */
            guideStyleIsDots: false,          /* ガイドのスタイル / Guide Style */
            /* スマートガイド：表示オプション / Smart Guides: display options */
            objectHighlighting: false,        /* オブジェクトのハイライト表示 / Object Highlighting */
            /* スマートガイド：スナップを制限 / Smart Guides: restrict snapping */
            snapRange: SNAP_RANGE_CANVAS,     /* スナップ範囲 / Snap range */
            snapToIsolatedObjects: true,      /* 編集モードのオブジェクトにスナップ / Snap to isolated objects */
            snapTolerance: 6,                 /* 許容値 / Snapping tolerance */
            /* 詳細設定・グリッドにスナップ / Advanced & snap to grid */
            snapToGrid: true,                 /* グリッドに強制スナップ / Snap to grid */
            showSnapToGridGuides: true,       /* グリッドにスナップするときにガイドを表示 / Show visual guides */
            snapToPointTolerance: 1,          /* ポイントにスナップ許容値 / Snap to point tolerance */
            /* パフォーマンス / Performance */
            animatedZoom: false,              /* アニメーションズーム / Animated Zoom */
            historyStates: 50,                /* ヒストリー数 1〜1000 / History States */
            realTimeDrawing: false,           /* リアルタイムの描画と編集 / Real-Time Drawing and Editing */
            /* ファイル管理 / File Management */
            editOriginalSystemDefault: true,  /* 「オリジナルの編集」にシステムデフォルトを使用 / Use System Defaults */
            autoActivateFonts: true,          /* Adobe Fonts を自動アクティベート / Auto-activate Adobe Fonts */
            saveToCloud: false,               /* ファイルの保存先 / Save Location */
            updateLinks: UPDATE_LINKS_AUTO,   /* リンクを更新 / Update Links */
            /* クリップボードの処理 / Clipboard Handling */
            includeSvgCode: true              /* SVGコードを含める / Include SVG Code */
        }
    };

    // =========================================
    // 環境設定キーの対応表 / Preference key bindings
    // =========================================

    /* 1項目=1キーで書けるキーの一覧。key=環境設定キー / presetField=プリセットの項目名 */
    /* valueType: "bool"=真偽値 / "boolInt"=ON:1・OFF:0 の整数 / "invertedInt"=ON:0・OFF:1 の整数 / "int"=数値そのまま / "real"=実数 */
    /* numericRule を添えた項目は NUMERIC_VALUE_RULES で範囲内に補正してから書き込む */
    /* Keys writable one-to-one; valueType selects how the preset value is stored */
    var PREFERENCE_BINDINGS = [
        /* 一般 / General */
        { key: "cursorKeyLength", presetField: "cursorKeyLength", valueType: "real" },
        { key: "showRichToolTips", presetField: "richToolTips", valueType: "bool" },
        { key: "Hello/ShowHomeScreenWS", presetField: "homeScreen", valueType: "bool" },
        { key: "Hello/NewDoc", presetField: "legacyNewDoc", valueType: "bool" },
        { key: "enablePrintBleedWidget", presetField: "printBleedWidget", valueType: "bool" },
        { key: "ContextualTaskBarEnabled", presetField: "contextualTaskBar", valueType: "bool" },
        { key: "plugin/DontShowWarningAgain/ShowPathfinderGroupWarning", presetField: "pathfinderWarning", valueType: "bool" },
        /* 選択範囲・アンカー表示 / Selection & Anchor Display */
        { key: "zoomToSelection", presetField: "zoomToSelection", valueType: "bool" },
        { key: "anchorSizePref", presetField: "anchorSize", valueType: "int" },
        { key: "hitShapeOnPreview", presetField: "objectPathOnly", valueType: "invertedInt" },
        { key: "hitTypeShapeOnPreview", presetField: "textPathOnly", valueType: "invertedInt" },
        /* アートボード / Artboard */
        { key: "moveLockedAndHiddenArt", presetField: "moveLockedArt", valueType: "bool" },
        { key: "showArtboardLabelOnCanvas", presetField: "showArtboardName", valueType: "bool" },
        { key: "ArtboardBBWidth", presetField: "artboardStrokeWidth", valueType: "real" },
        /* テキスト / Text */
        { key: "text/autoSizing", presetField: "autoSizeAreaText", valueType: "bool" },
        { key: "text/fillWithDefaultText", presetField: "fillSampleText", valueType: "bool" },
        { key: "text/fillWithDefaultTextJP", presetField: "fillSampleText", valueType: "bool" },
        { key: "text/doFontLocking", presetField: "missingGlyphProtection", valueType: "bool" },
        { key: "text/enableAlternateGlyph", presetField: "alternateGlyph", valueType: "bool" },
        { key: "text/enableListAutoDetection", presetField: "listAutoDetection", valueType: "bool" },
        /* ユーザーインターフェイス / User Interface */
        { key: "uiCanvasIsWhite", presetField: "canvasWhite", valueType: "boolInt" },
        { key: "uiBrightness", presetField: "uiBrightness", valueType: "real" },
        /* ガイド / Guides */
        { key: "Guide/Style", presetField: "guideStyleIsDots", valueType: "boolInt" }, /* 0=ライン, 1=ポイント / 0=lines, 1=dots */
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
        /* パフォーマンス / Performance */
        { key: "Performance/AnimZoom", presetField: "animatedZoom", valueType: "bool" },
        { key: "maximumUndoDepth", presetField: "historyStates", valueType: "int", numericRule: NUMERIC_VALUE_RULES.historyStates },
        { key: "LiveEdit_State_Machine", presetField: "realTimeDrawing", valueType: "bool" },
        /* ファイル管理 / File Management */
        { key: "useSysDefEdit", presetField: "editOriginalSystemDefault", valueType: "bool" },
        { key: "AutoActivateMissingFont", presetField: "autoActivateFonts", valueType: "bool" },
        { key: "AdobeSaveAsCloudDocumentPreference", presetField: "saveToCloud", valueType: "bool" },
        { key: "plugin/FileClipboard/linkoptions", presetField: "updateLinks", valueType: "int" },
        /* クリップボードの処理 / Clipboard Handling */
        { key: "plugin/FileClipboard/copySVGCode", presetField: "includeSvgCode", valueType: "bool" },
        /* ブラックのアピアランス / Black Appearance */
        { key: "blackPreservation/Onscreen", presetField: "accurateBlackOnscreen", valueType: "invertedInt" },
        { key: "blackPreservation/Export", presetField: "accurateBlackExport", valueType: "invertedInt" },
        /* 再起動が必要 / Requires a restart */
        { key: "rulerType", presetField: "rulerType", valueType: "int" },
        { key: "Planar/MergeTool/PaintFills", presetField: "shapeBuilderPaintFills", valueType: "int" }
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
     * プリセットにその項目が書かれているか / Whether the preset defines this field
     * 書かれていない項目は書き込まず、Illustratorの現在値をそのまま残す
     * @param {object} presetState - プリセットの設定一式
     * @param {string} presetField - 項目名
     * @returns {boolean} 定義されていれば true
     */
    function hasField(presetState, presetField) {
        return typeof presetState[presetField] !== "undefined";
    }

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
     * @param {object} presetState - プリセットの設定一式
     * @returns {void}
     */
    function applyPreferenceBindings(presetState) {
        for (var i = 0; i < PREFERENCE_BINDINGS.length; i++) {
            var binding = PREFERENCE_BINDINGS[i];
            if (!hasField(presetState, binding.presetField)) continue;

            var presetValue = presetState[binding.presetField];
            if (binding.valueType === "bool") {
                safeSetBool(binding.key, !!presetValue);
            } else if (binding.valueType === "boolInt") {
                safeSetInt(binding.key, presetValue ? 1 : 0);
            } else if (binding.valueType === "invertedInt") {
                safeSetInt(binding.key, presetValue ? 0 : 1); /* 0がON / 0 means ON */
            } else if (binding.valueType === "int") {
                safeSetInt(binding.key, binding.numericRule ? clampIntToRule(presetValue, binding.numericRule) : presetValue);
            } else if (binding.valueType === "real") {
                safeSetReal(binding.key, presetValue);
            }
        }
    }

    /**
     * 1項目で複数キーを書く設定を反映 / Write the fields that map to more than one key
     * @param {object} presetState - プリセットの設定一式
     * @returns {void}
     */
    function applyCompositePreferences(presetState) {
        /* テキスト（0＝一覧を非表示）/ Text (0 = hide the list) */
        if (hasField(presetState, "recentFontsEnabled")) {
            safeSetInt("text/recentFontMenu/showNEntries",
                presetState.recentFontsEnabled ? clampIntToRule(presetState.recentFontsCount, NUMERIC_VALUE_RULES.recentFonts) : 0);
        }

        /* ガイド / Guides */
        if (hasField(presetState, "guideColorIsLightBlue")) {
            writeGuideColor(presetState.guideColorIsLightBlue ? GUIDE_COLOR_LIGHT_BLUE : GUIDE_COLOR_CYAN);
        }

        /* アートボード / Artboard */
        if (hasField(presetState, "artboardColorIndex")) {
            writeArtboardColor(presetState.artboardColorIndex);
        }
    }

    /**
     * 設定一式を環境設定へ書き込む / Write a whole preset into the preferences
     * @param {object} presetState - プリセットの設定一式
     * @returns {void}
     */
    function applyPreset(presetState) {
        applyPreferenceBindings(presetState);
        applyCompositePreferences(presetState);
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
     * ACTIVE_PRESET のプリセットを環境設定へ書き込み、画面を更新する / Write ACTIVE_PRESET and refresh the screen
     * @returns {void}
     */
    function main() {
        var presetState = PRESET_STATES[ACTIVE_PRESET];
        if (!presetState) {
            alert(SCRIPT_NAME + "\nACTIVE_PRESET の値が正しくありません: " + ACTIVE_PRESET);
            return;
        }
        applyPreset(presetState);
        forceScreenRefresh();
    }

    main();
})();
