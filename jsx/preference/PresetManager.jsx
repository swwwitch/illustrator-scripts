#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- Illustrator の主要な環境設定を、カテゴリ別に並べた1枚のダイアログでまとめて確認・変更します。
- ［デフォルト］／［プリセット1］を選ぶと、一式の設定値をUIに反映できます。
- 変更はすべて［OK］でまとめて書き込まれます。
- 対象カテゴリや環境設定キーの一覧は README を参照してください。

### Overview

- Review and change the major Illustrator preferences from a single dialog grouped by category.
- Selecting [Default] or [Preset 1] loads a whole set of values into the UI.
- All changes are written at once when [OK] is pressed.
- See the README for the covered categories and the preference keys.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManager";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManager.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManager.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n3b33862538f6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 数値入力欄の許容範囲と既定値 / Allowed range and default for numeric fields */
    var NUMERIC_INPUT_RULES = {
        recentFonts: { min: 0, max: 30, defaultValue: 15 },     /* 0で一覧を非表示 / 0 hides the list */
        historyStates: { min: 1, max: 1000, defaultValue: 100 } /* 想定範囲 1〜1000 / Expected range 1-1000 */
    };

    /* ダイアログの表示位置と不透明度 / Dialog position offset and opacity */
    var DIALOG_OFFSET_X = 300;
    var DIALOG_OFFSET_Y = 0;
    var DIALOG_OPACITY = 0.98;

    /* アートボードのストローク幅の選択肢 / Selectable artboard stroke widths */
    var ARTBOARD_STROKE_WIDTHS = [1, 2, 3, 4];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* UIロケールから言語(ja/en)を判定 / Detect UI language (ja/en) from locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義（カテゴリ別に構造化）/ Japanese-English label definitions (grouped by category) */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "環境設定をまとめて変更", en: "Illustrator Preferences Utility" }
        },
        /* パネル見出し / Panel titles */
        panel: {
            general: { ja: "［一般］カテゴリ", en: "[General] Category" },
            selectionAnchor: {
                ja: "［選択範囲・アンカー表示］カテゴリ",
                en: "[Selection & Anchor Display] Category"
            },
            artboard: { ja: "アートボード", en: "Artboard" },
            text: { ja: "［テキスト］カテゴリ", en: "[Text] Category" },
            guides: { ja: "ガイド", en: "Guides" },
            smartGuides: { ja: "スマートガイド", en: "Smart Guides" },
            userInterface: {
                ja: "［ユーザーインターフェイス］カテゴリ",
                en: "[User Interface] Category"
            },
            performance: { ja: "［パフォーマンス］カテゴリ", en: "[Performance] Category" },
            fileManagement: { ja: "［ファイル管理］カテゴリ", en: "[File Management] Category" },
            clipboard: { ja: "クリップボードの処理", en: "Clipboard Handling" },
            limitToPath: { ja: "パスに制限", en: "Limit to Path" }
        },
        /* チェックボックス / Checkboxes */
        checkbox: {
            richToolTips: { ja: "詳細なツールヒントを表示", en: "Show Rich Tool Tips" },
            homeScreen: {
                ja: "ドキュメントを開いていないときに「ホーム画面」を表示",
                en: "Show the Home Screen When No Documents Are Open"
            },
            legacyNewDoc: {
                ja: "以前の「新規ドキュメント」インターフェイスを使用",
                en: "Use Legacy \"File > New\" Interface"
            },
            printBleedWidget: {
                ja: "「裁ち落としを印刷」生成AIボタンを表示",
                en: "Show 'Print Bleed' generative AI buttons on Bleed"
            },
            moveLockedArt: {
                ja: "ロックまたは非表示オブジェクトを一緒に移動",
                en: "Move Locked and Hidden Artwork"
            },
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            zoomToSelection: { ja: "選択範囲へズーム", en: "Zoom to Selection" },
            objectPathOnly: {
                ja: "オブジェクトの選択範囲をパスに制限",
                en: "Object Selection by Path Only"
            },
            textPathOnly: {
                ja: "テキストオブジェクトの選択範囲をパスに制限",
                en: "Type Object Selection by Path Only"
            },
            autoSizeAreaText: {
                ja: "新規エリア内文字の自動サイズ調整",
                en: "Auto Size New Area Type"
            },
            recentFonts: { ja: "最近使用したフォントの表示数", en: "Number of Recent Fonts" },
            missingGlyphProtection: {
                ja: "見つからない字形の保護を有効にする",
                en: "Enable Missing Glyph Protection"
            },
            alternateGlyph: { ja: "選択された文字の異体字を表示", en: "Show Character Alternates" },
            objectHighlighting: { ja: "オブジェクトのハイライト表示", en: "Object Highlighting" },
            animatedZoom: { ja: "アニメーションズーム", en: "Animated Zoom" },
            realTimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-Time Drawing and Editing" },
            editOriginalSystemDefault: {
                ja: "「オリジナルの編集」にシステムデフォルトを使用",
                en: "Use System Defaults for ‘Edit Original’"
            },
            autoActivateFonts: { ja: "Adobe Fonts を自動アクティベート", en: "Auto-activate Adobe Fonts" },
            includeSvgCode: { ja: "SVGコードを含める", en: "Include SVG Code" }
        },
        /* 見出しラベル / Static labels */
        label: {
            preset: { ja: "プリセット{colon}", en: "Preset{colon}" },
            artboardColor: { ja: "ハイライトのカラー{colon}", en: "Highlight Color{colon}" },
            artboardStrokeWidth: { ja: "ストロークの幅{colon}", en: "Stroke Width{colon}" },
            guideColor: { ja: "カラー{colon}", en: "Color{colon}" },
            guideStyle: { ja: "スタイル{colon}", en: "Style{colon}" },
            brightness: { ja: "明るさ{colon}", en: "Brightness{colon}" },
            canvasColor: { ja: "カンバスカラー{colon}", en: "Canvas Color{colon}" },
            historyStates: { ja: "ヒストリー数{colon}", en: "History States{colon}" },
            saveLocation: { ja: "ファイルの保存先{colon}", en: "Save Location{colon}" },
            updateLinks: { ja: "リンクを更新{colon}", en: "Update Links{colon}" }
        },
        /* ドロップダウンの項目 / Dropdown items */
        dropdown: {
            presetCurrent: { ja: "現在の設定", en: "Current Settings" },
            presetDefault: { ja: "デフォルト", en: "Default" },
            preset1: { ja: "プリセット1", en: "Preset 1" },
            colorLightBlue: { ja: "ライトブルー", en: "Light Blue" },
            colorLightRed: { ja: "サーモンピンク", en: "Light Red" },
            colorGreen: { ja: "グリーン", en: "Green" },
            colorMediumBlue: { ja: "ミディアムブルー", en: "Medium Blue" },
            colorMagenta: { ja: "マゼンタ", en: "Magenta" },
            colorCyan: { ja: "シアン", en: "Cyan" },
            colorLightGray: { ja: "ライトグレー", en: "Light Gray" },
            colorBlack: { ja: "ブラック", en: "Black" },
            colorYellow: { ja: "イエロー", en: "Yellow" }
        },
        /* ラジオボタン / Radio buttons */
        radio: {
            guideColorCyan: { ja: "シアン", en: "Cyan" },
            guideColorLightBlue: { ja: "ライトブルー", en: "Light Blue" },
            guideStyleLines: { ja: "ライン", en: "Lines" },
            guideStyleDots: { ja: "点線", en: "Dots" },
            canvasMatch: { ja: "UIに合わせる", en: "Match Brightness" },
            canvasWhite: { ja: "ホワイト", en: "White" },
            saveToComputer: { ja: "コンピューター", en: "Computer" },
            saveToCloud: { ja: "クラウド", en: "Cloud" },
            updateLinksAuto: { ja: "自動", en: "Automatic" },
            updateLinksManual: { ja: "手動", en: "Manual" },
            updateLinksAsk: { ja: "確認", en: "Ask When Modified" }
        },
        /* 明るさスウォッチ / Brightness swatches */
        swatch: {
            dark: { ja: "暗", en: "Dark" },
            mediumDark: { ja: "やや暗", en: "Medium Dark" },
            mediumLight: { ja: "やや明", en: "Medium Light" },
            light: { ja: "明", en: "Light" }
        },
        /* ボタン / Buttons */
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        /* ヘルプチップ / Help tips */
        hint: {
            brightness: {
                ja: "インターフェイスカラーは直接反映できないため、［OK］後に環境設定（ユーザーインターフェイス）が開きます。矢印キー＋Return で確定してください",
                en: "Interface color can’t be applied directly, so Preferences opens after [OK]. Confirm with the arrow keys + Return."
            },
            historyStates: { ja: "ヒストリー数を設定", en: "Set history states" }
        },
        /* 警告メッセージ / Alerts */
        alert: {
            savePrefsFailed: {
                ja: "環境設定の保存に失敗しました: ",
                en: "Failed to save preferences: "
            }
        }
    };

    /* LABELS からドット区切りのパスで多言語テキストを取り出す / Look up localized text from LABELS by dot path */
    function L(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            node = node[parts[i]];
            if (!node) return path; /* 保険（パスをそのまま返す）/ Fallback: return the path itself */
        }
        var text = node[lang] || node.en;
        if (!text) return path;
        return applyUISymbols(text);
    }

    /* {colon} などのプレースホルダを言語別の記号に展開 / Expand placeholders like {colon} */
    function applyUISymbols(text) {
        return text.replace(/\{colon\}/g, (lang === "ja") ? "：" : ":");
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING = 6;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;               /* 2カラムの間隔 / gap between columns */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ラベル＋コントロールの横並び）/ Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : 10;
    }

    // =========================================
    // 環境設定の読み書き / Preference accessors
    // =========================================

    /* 整数の環境設定を書き込む（失敗しても処理を止めない）/ Write an integer preference (never throws) */
    function safeSetInt(key, value) {
        try {
            app.preferences.setIntegerPreference(key, value);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] safeSetInt failed: " + key + " = " + value + " / " + e);
        }
    }

    /* 真偽値の環境設定を書き込む / Write a boolean preference (never throws) */
    function safeSetBool(key, value) {
        try {
            app.preferences.setBooleanPreference(key, value);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] safeSetBool failed: " + key + " = " + value + " / " + e);
        }
    }

    /* 実数の環境設定を書き込む / Write a real-number preference (never throws) */
    function safeSetReal(key, value) {
        try {
            app.preferences.setRealPreference(key, value);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] safeSetReal failed: " + key + " = " + value + " / " + e);
        }
    }

    /* 整数の環境設定を読む。存在しないキーはfallback / Read an integer preference, falling back when missing */
    function safeGetInt(key, fallback) {
        try {
            return app.preferences.getIntegerPreference(key);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] safeGetInt fallback: " + key + " / " + e);
            return fallback;
        }
    }

    /* 真偽値の環境設定を読む / Read a boolean preference, falling back when missing */
    function safeGetBool(key, fallback) {
        try {
            return app.preferences.getBooleanPreference(key);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] safeGetBool fallback: " + key + " / " + e);
            return fallback;
        }
    }

    /* 実数の環境設定を読む / Read a real-number preference, falling back when missing */
    function safeGetReal(key, fallback) {
        try {
            return app.preferences.getRealPreference(key);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] safeGetReal fallback: " + key + " / " + e);
            return fallback;
        }
    }

    // =========================================
    // 数値ユーティリティ / Numeric utilities
    // =========================================

    /* 値を min〜max に収める / Clamp a value into the min-max range */
    function clamp(value, min, max) {
        return Math.max(min, Math.min(max, value));
    }

    /* 整数に変換。数値でなければ null / Parse as an integer, or null when not numeric */
    function parseIntOrNull(value) {
        var parsed = parseInt(value, 10);
        return isNaN(parsed) ? null : parsed;
    }

    /* ルール（min/max/def）に沿って整数へ補正 / Normalize a value against a min/max/def rule */
    function clampIntToRule(value, rule) {
        var parsed = parseIntOrNull(value);
        if (parsed === null) return rule.defaultValue;
        return clamp(parsed, rule.min, rule.max);
    }

    /* 入力欄の値を補正して書き戻す / Normalize an edittext value in place and return it */
    function normalizeIntInput(inputField, rule) {
        var normalized = clampIntToRule(inputField.text, rule);
        inputField.text = String(normalized);
        return normalized;
    }

    /* 許容誤差つきで比較 / Compare two numbers with a tolerance */
    function almostEqual(a, b, tolerance) {
        if (typeof tolerance !== "number") tolerance = 0.02;
        return Math.abs(a - b) < tolerance;
    }

    // =========================================
    // 設定値の定義 / Preference value definitions
    // =========================================

    /* アートボードのハイライトカラープリセット（RGB 0..1）/ Artboard highlight color presets (RGB 0..1) */
    var ARTBOARD_COLOR_PRESETS = [
        { label: L("dropdown.colorLightBlue"), red: 0.29, green: 0.52, blue: 1.0 },
        { label: L("dropdown.colorLightRed"), red: 1.0, green: 0.29, blue: 0.29 },
        { label: L("dropdown.colorGreen"), red: 0.0, green: 0.65, blue: 0.31 },
        { label: L("dropdown.colorMediumBlue"), red: 0.0, green: 0.45, blue: 0.78 },
        { label: L("dropdown.colorMagenta"), red: 1.0, green: 0.0, blue: 1.0 },
        { label: L("dropdown.colorCyan"), red: 0.0, green: 1.0, blue: 1.0 },
        { label: L("dropdown.colorLightGray"), red: 1.0, green: 1.0, blue: 1.0 },
        { label: L("dropdown.colorBlack"), red: 0.0, green: 0.0, blue: 0.0 },
        { label: L("dropdown.colorYellow"), red: 1.0, green: 1.0, blue: 0.0 }
    ];

    /* ガイドカラーの2択（RGB 0..1）/ The two selectable guide colors (RGB 0..1) */
    var GUIDE_COLOR_CYAN = { red: 0.0, green: 1.0, blue: 1.0 };
    var GUIDE_COLOR_LIGHT_BLUE = { red: 0.29, green: 0.52, blue: 1.0 };

    /* ガイドスタイルの値 / Guide style values */
    var GUIDE_STYLE_LINES = 0;
    var GUIDE_STYLE_DOTS = 1;

    /* 「リンクを更新」の値 / "Update Links" values */
    var UPDATE_LINKS_AUTO = 0;
    var UPDATE_LINKS_MANUAL = 1;
    var UPDATE_LINKS_ASK = 2;

    /* UI明るさの4プリセット：uiBrightness 値・スウォッチのシェード（RGB 0..1）・ラベル / Four UI-brightness presets: uiBrightness value, swatch shade (RGB 0..1), label */
    /* 値は連続値ではなく離散プリセット（0.5 と 0.50999999 が別段階）/ Values are discrete presets, not a continuous scale (0.5 vs 0.50999999 are distinct steps) */
    var BRIGHTNESS_LEVELS = [
        { value: 0.0, shade: [0.22, 0.22, 0.22], labelPath: "swatch.dark" },
        { value: 0.5, shade: [0.33, 0.33, 0.33], labelPath: "swatch.mediumDark" },
        { value: 0.50999999046326, shade: [0.70, 0.70, 0.70], labelPath: "swatch.mediumLight" },
        { value: 1.0, shade: [0.94, 0.94, 0.94], labelPath: "swatch.light" }
    ];
    var BRIGHTNESS_SWATCH_SIZE = 23;                        /* スウォッチの一辺(px) / Swatch side length (px) */
    var BRIGHTNESS_SELECTED_BORDER = [0.15, 0.5, 0.92];     /* 選択枠の青 / Blue selection border */
    var BRIGHTNESS_SWATCH_OUTLINE = [0.5, 0.5, 0.5];        /* 通常時の細枠 / Thin outline when not selected */
    var BRIGHTNESS_TOLERANCE = 0.001;                       /* プリセット同士の判定用（0.5 と 0.50999999 を区別）/ Preset comparison tolerance (keeps 0.5 and 0.50999999 distinct) */

    /* プリセットの識別子 / Preset identifiers */
    var PRESET_IDS = {
        current: "current",
        defaultPreset: "default",
        preset1: "preset1"
    };

    /* ［デフォルト］の設定一式。全項目を必ず定義する / [Default] preset; every field must be defined */
    var PRESET_STATE_DEFAULT = {
        richToolTips: true,
        homeScreen: true,
        legacyNewDoc: false,
        printBleedWidget: true,
        moveLockedArt: false,
        objectPathOnly: false,
        zoomToSelection: true,
        textPathOnly: false,
        showArtboardName: true,
        artboardColorIndex: 0,
        artboardStrokeWidth: 1,
        autoSizeAreaText: false,
        recentFontsEnabled: true,
        recentFontsCount: 10,
        missingGlyphProtection: true,
        alternateGlyph: true,
        objectHighlighting: true,
        animatedZoom: true,
        historyStates: 100,
        realTimeDrawing: true,
        editOriginalSystemDefault: false,
        autoActivateFonts: false,
        canvasWhite: false,
        guideColorIsLightBlue: false,
        saveToCloud: true,
        updateLinks: UPDATE_LINKS_ASK,
        includeSvgCode: false
    };

    /* ［プリセット1］の設定一式。全項目を必ず定義する / [Preset 1]; every field must be defined */
    var PRESET_STATE_1 = {
        richToolTips: false,
        homeScreen: false,
        legacyNewDoc: true,
        printBleedWidget: false,
        moveLockedArt: true,
        objectPathOnly: false,
        zoomToSelection: false,
        textPathOnly: false,
        showArtboardName: false,
        artboardColorIndex: 7,
        artboardStrokeWidth: 2,
        autoSizeAreaText: true,
        recentFontsEnabled: true,
        recentFontsCount: 15,
        missingGlyphProtection: false,
        alternateGlyph: false,
        objectHighlighting: false,
        animatedZoom: false,
        historyStates: 50,
        realTimeDrawing: false,
        editOriginalSystemDefault: true,
        autoActivateFonts: true,
        canvasWhite: true,
        guideColorIsLightBlue: true,
        saveToCloud: false,
        updateLinks: UPDATE_LINKS_AUTO,
        includeSvgCode: true
    };

    // =========================================
    // 設定値のヘルパー / Preference value helpers
    // =========================================

    /* 現在のRGBに最も近いハイライトカラーの index / Index of the highlight color nearest to the given RGB */
    function findNearestArtboardColorIndex(red, green, blue) {
        var nearestIndex = 0;
        var nearestDistance = Infinity;
        for (var i = 0; i < ARTBOARD_COLOR_PRESETS.length; i++) {
            var colorPreset = ARTBOARD_COLOR_PRESETS[i];
            var distance = Math.abs(colorPreset.red - red) +
                Math.abs(colorPreset.green - green) +
                Math.abs(colorPreset.blue - blue);
            if (distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /* uiBrightness 値に最も近いプリセットの index / Index of the brightness preset closest to a uiBrightness value */
    function findNearestBrightnessIndex(value) {
        var nearestIndex = 0;
        var nearestDiff = Infinity;
        for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
            var diff = Math.abs(BRIGHTNESS_LEVELS[i].value - value);
            if (diff < nearestDiff) {
                nearestDiff = diff;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /* 現在のガイドカラーがライトブルーか / Whether the current guide color is Light Blue */
    function isGuideColorLightBlue() {
        return almostEqual(safeGetReal("Guide/Color/red", GUIDE_COLOR_CYAN.red), GUIDE_COLOR_LIGHT_BLUE.red) &&
            almostEqual(safeGetReal("Guide/Color/green", GUIDE_COLOR_CYAN.green), GUIDE_COLOR_LIGHT_BLUE.green) &&
            almostEqual(safeGetReal("Guide/Color/blue", GUIDE_COLOR_CYAN.blue), GUIDE_COLOR_LIGHT_BLUE.blue);
    }

    /* ガイドカラーを書き込む / Write the guide color */
    function writeGuideColor(guideColor) {
        safeSetReal("Guide/Color/red", guideColor.red);
        safeSetReal("Guide/Color/green", guideColor.green);
        safeSetReal("Guide/Color/blue", guideColor.blue);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* ダイアログを組み立てて表示し、［OK］で環境設定を保存する / Build and show the dialog, then save preferences on [OK] */
    function main() {

        /* ダイアログ本体 / Dialog window */
        var prefsDialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(prefsDialog);

        /* 全体を縦に積むコンテナ / Vertical container for the whole dialog */
        var dialogContentGroup = prefsDialog.add("group");
        dialogContentGroup.orientation = "column";
        dialogContentGroup.alignChildren = "left";

        /* プリセット選択行（現在の設定 / デフォルト / プリセット1）/ Preset selector row (Current / Default / Preset 1) */
        var presetRow = dialogContentGroup.add("group");
        presetRow.name = "presetRow";
        presetRow.alignment = "center";
        presetRow.margins = [0, 10, 20, 20];
        setupRow(presetRow, "center");
        var presetSelectorGroup = presetRow.add("group");
        setupRow(presetSelectorGroup);
        presetSelectorGroup.add("statictext", undefined, L("label.preset"));
        var presetLabels = [
            L("dropdown.presetCurrent"),
            L("dropdown.presetDefault"),
            L("dropdown.preset1")
        ];
        var presetDropdown = presetSelectorGroup.add("dropdownlist", undefined, presetLabels);
        presetDropdown.selection = 0; /* 初期選択は「現在の設定」/ Default selection is "Current Settings" */

        /* パネルを左右2カラムに配置 / Two-column layout for the panels */
        var panelColumnsGroup = dialogContentGroup.add("group");
        panelColumnsGroup.orientation = "row";
        panelColumnsGroup.alignChildren = "top";
        panelColumnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = panelColumnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];

        var rightColumnGroup = panelColumnsGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = ["fill", "top"];

        /* ［一般］パネル（左）/ General panel (left) */
        var generalPanel = leftColumnGroup.add("panel", undefined, L("panel.general"));
        setupPanel(generalPanel);

        var richToolTipsCheckbox = generalPanel.add("checkbox", undefined, L("checkbox.richToolTips"));
        richToolTipsCheckbox.helpTip = L("checkbox.richToolTips");

        var homeScreenCheckbox = generalPanel.add("checkbox", undefined, L("checkbox.homeScreen"));
        homeScreenCheckbox.helpTip = L("checkbox.homeScreen");

        var legacyNewDocCheckbox = generalPanel.add("checkbox", undefined, L("checkbox.legacyNewDoc"));
        legacyNewDocCheckbox.helpTip = L("checkbox.legacyNewDoc");

        var printBleedWidgetCheckbox = generalPanel.add("checkbox", undefined, L("checkbox.printBleedWidget"));
        printBleedWidgetCheckbox.helpTip = L("checkbox.printBleedWidget");

        /* ［選択範囲・アンカー表示］パネル（左）/ Selection & Anchor Display panel (left) */
        var selectionAnchorPanel = leftColumnGroup.add("panel", undefined, L("panel.selectionAnchor"));
        setupPanel(selectionAnchorPanel);

        var zoomToSelectionCheckbox = selectionAnchorPanel.add("checkbox", undefined, L("checkbox.zoomToSelection"));
        zoomToSelectionCheckbox.helpTip = L("checkbox.zoomToSelection");

        /* ［アートボード］パネル（左）/ Artboard panel (left) */
        var artboardPanel = leftColumnGroup.add("panel", undefined, L("panel.artboard"));
        setupPanel(artboardPanel);

        var moveLockedArtCheckbox = artboardPanel.add("checkbox", undefined, L("checkbox.moveLockedArt"));
        moveLockedArtCheckbox.helpTip = L("checkbox.moveLockedArt");

        var showArtboardNameCheckbox = artboardPanel.add("checkbox", undefined, L("checkbox.showArtboardName"));
        showArtboardNameCheckbox.helpTip = L("checkbox.showArtboardName");

        /* ハイライトのカラー（9色のプリセットから選択）/ Highlight color (nine presets) */
        var artboardColorRow = artboardPanel.add("group");
        setupRow(artboardColorRow);
        artboardColorRow.add("statictext", undefined, L("label.artboardColor"));
        var artboardColorLabels = [];
        for (var i = 0; i < ARTBOARD_COLOR_PRESETS.length; i++) {
            artboardColorLabels.push(ARTBOARD_COLOR_PRESETS[i].label);
        }
        var artboardColorDropdown = artboardColorRow.add("dropdownlist", undefined, artboardColorLabels);

        /* ストロークの幅（1〜4）/ Stroke width (1-4) */
        var artboardStrokeWidthRow = artboardPanel.add("group");
        setupRow(artboardStrokeWidthRow);
        artboardStrokeWidthRow.add("statictext", undefined, L("label.artboardStrokeWidth"));
        var artboardStrokeWidthRadios = [];
        for (var i = 0; i < ARTBOARD_STROKE_WIDTHS.length; i++) {
            artboardStrokeWidthRadios.push(
                artboardStrokeWidthRow.add("radiobutton", undefined, String(ARTBOARD_STROKE_WIDTHS[i]))
            );
        }

        /* ［テキスト］パネル（左）/ Text panel (left) */
        var textPanel = leftColumnGroup.add("panel", undefined, L("panel.text"));
        setupPanel(textPanel);

        var autoSizeAreaTextCheckbox = textPanel.add("checkbox", undefined, L("checkbox.autoSizeAreaText"));
        autoSizeAreaTextCheckbox.helpTip = L("checkbox.autoSizeAreaText");

        /* 最近使用したフォントの表示数（チェックOFFで0＝非表示）/ Recent font count (unchecked means 0 = hidden) */
        var recentFontsRow = textPanel.add("group");
        setupRow(recentFontsRow);
        var recentFontsCheckbox = recentFontsRow.add("checkbox", undefined, L("checkbox.recentFonts"));
        recentFontsCheckbox.helpTip = L("checkbox.recentFonts");
        var recentFontsInput = recentFontsRow.add("edittext", undefined, "0");
        recentFontsInput.characters = 3;

        var missingGlyphProtectionCheckbox = textPanel.add("checkbox", undefined, L("checkbox.missingGlyphProtection"));
        missingGlyphProtectionCheckbox.helpTip = L("checkbox.missingGlyphProtection");

        var alternateGlyphCheckbox = textPanel.add("checkbox", undefined, L("checkbox.alternateGlyph"));
        alternateGlyphCheckbox.helpTip = L("checkbox.alternateGlyph");

        /* ［ユーザーインターフェイス］パネル（左）/ User Interface panel (left) */
        var userInterfacePanel = leftColumnGroup.add("panel", undefined, L("panel.userInterface"));
        setupPanel(userInterfacePanel);

        /* 明るさ（ラベル＋4段階のスウォッチ）。現在は非表示にしているが、復活できるよう残す */
        /* Brightness (label + four swatches); currently hidden but kept so it can be restored */
        // var brightnessRow = userInterfacePanel.add("group");
        // setupRow(brightnessRow);
        // brightnessRow.add("statictext", undefined, L("label.brightness"));
        // var brightnessSwatchGroup = brightnessRow.add("group");
        // setupRow(brightnessSwatchGroup, "left", 8);

        var brightnessSwatchButtons = [];
        var selectedBrightnessIndex = -1;
        var brightnessSwatchTouched = false; /* ユーザーがスウォッチを操作したか（未操作なら書き込まない）/ Whether the user clicked a swatch (no write when untouched) */

        /* スウォッチの描画ハンドラを作る：塗り＋（選択時のみ）青い枠 / Build a swatch draw handler: fill + (only when selected) a blue border */
        function createBrightnessSwatchDrawHandler(index) {
            return function () {
                var graphics = this.graphics;
                var swatchWidth = this.size[0];
                var swatchHeight = this.size[1];
                var brightnessLevel = BRIGHTNESS_LEVELS[index];

                /* シェードで塗りつぶし / Fill with the shade */
                var brush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, brightnessLevel.shade.concat(1));
                graphics.newPath();
                graphics.rectPath(0, 0, swatchWidth, swatchHeight);
                graphics.fillPath(brush);

                if (index === selectedBrightnessIndex) {
                    /* 選択：太めの青枠 / Selected: thicker blue border */
                    var selectedPen = graphics.newPen(graphics.PenType.SOLID_COLOR, BRIGHTNESS_SELECTED_BORDER.concat(1), 2);
                    graphics.newPath();
                    graphics.rectPath(1, 1, swatchWidth - 2, swatchHeight - 2);
                    graphics.strokePath(selectedPen);
                } else {
                    /* 非選択：細いグレー枠（明シェードを明背景でも視認）/ Not selected: thin gray outline (keep light swatches visible) */
                    var outlinePen = graphics.newPen(graphics.PenType.SOLID_COLOR, BRIGHTNESS_SWATCH_OUTLINE.concat(1), 1);
                    graphics.newPath();
                    graphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
                    graphics.strokePath(outlinePen);
                }
            };
        }

        /* 全スウォッチを再描画（hide/show で onDraw を確実に再実行し、旧選択枠を残さない＝排他表示）/ Force all swatches to repaint (hide/show reliably re-runs onDraw so the old selection border never lingers = exclusive) */
        function refreshBrightnessSwatches() {
            for (var i = 0; i < brightnessSwatchButtons.length; i++) {
                brightnessSwatchButtons[i].hide();
                brightnessSwatchButtons[i].show();
            }
        }

        /* uiBrightness 値から選択状態だけ更新（書き込みは［OK］時）/ Update the selection only, from a uiBrightness value (writes happen on [OK]) */
        function selectBrightnessSwatchByValue(value) {
            var nextIndex = findNearestBrightnessIndex(value);
            if (nextIndex === selectedBrightnessIndex) return;
            selectedBrightnessIndex = nextIndex;
            refreshBrightnessSwatches();
        }

        /* スウォッチを4つ生成（非表示中）/ Build the four swatches (currently hidden) */
        // for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
        //     var brightnessSwatchButton = brightnessSwatchGroup.add("iconbutton", undefined, undefined, { style: "toolbutton" });
        //     brightnessSwatchButton.preferredSize = [BRIGHTNESS_SWATCH_SIZE, BRIGHTNESS_SWATCH_SIZE];
        //     brightnessSwatchButton.alignment = ["left", "center"];
        //     brightnessSwatchButton.helpTip = L(BRIGHTNESS_LEVELS[i].labelPath) + " / " + L("hint.brightness");
        //     brightnessSwatchButton.onDraw = createBrightnessSwatchDrawHandler(i);
        //     brightnessSwatchButtons.push(brightnessSwatchButton);
        //     (function (index) {
        //         brightnessSwatchButtons[index].onClick = function () {
        //             /* UIのみ更新し、保存は［OK］時 / UI only; persist on [OK] */
        //             selectedBrightnessIndex = index;
        //             brightnessSwatchTouched = true;
        //             refreshBrightnessSwatches();
        //         };
        //     })(brightnessSwatchButtons.length - 1);
        // }

        /* カンバスカラー（UIに合わせる／ホワイト）/ Canvas color (Match Brightness / White) */
        var canvasColorRow = userInterfacePanel.add("group");
        setupRow(canvasColorRow);
        canvasColorRow.add("statictext", undefined, L("label.canvasColor"));
        var canvasMatchRadio = canvasColorRow.add("radiobutton", undefined, L("radio.canvasMatch"));
        var canvasWhiteRadio = canvasColorRow.add("radiobutton", undefined, L("radio.canvasWhite"));

        /* ［ガイド］パネル（右）/ Guides panel (right) */
        var guidesPanel = rightColumnGroup.add("panel", undefined, L("panel.guides"));
        setupPanel(guidesPanel);

        var GUIDE_LABEL_WIDTH = 80; /* ガイド内ラベルの共通幅 / Unified label width inside Guides */

        /* ガイドのカラー（シアン／ライトブルー）/ Guide color (Cyan / Light Blue) */
        var guideColorRow = guidesPanel.add("group");
        setupRow(guideColorRow);
        var guideColorLabel = guideColorRow.add("statictext", undefined, L("label.guideColor"));
        guideColorLabel.preferredSize = [GUIDE_LABEL_WIDTH, guideColorLabel.preferredSize ? guideColorLabel.preferredSize[1] : 20];
        guideColorLabel.justify = "right";
        var guideColorCyanRadio = guideColorRow.add("radiobutton", undefined, L("radio.guideColorCyan"));
        var guideColorLightBlueRadio = guideColorRow.add("radiobutton", undefined, L("radio.guideColorLightBlue"));

        /* ガイドのスタイル（ライン／点線）/ Guide style (Lines / Dots) */
        var guideStyleRow = guidesPanel.add("group");
        setupRow(guideStyleRow);
        var guideStyleLabel = guideStyleRow.add("statictext", undefined, L("label.guideStyle"));
        guideStyleLabel.preferredSize = [GUIDE_LABEL_WIDTH, guideStyleLabel.preferredSize ? guideStyleLabel.preferredSize[1] : 20];
        guideStyleLabel.justify = "right";
        var guideStyleLinesRadio = guideStyleRow.add("radiobutton", undefined, L("radio.guideStyleLines"));
        var guideStyleDotsRadio = guideStyleRow.add("radiobutton", undefined, L("radio.guideStyleDots"));

        /* ［スマートガイド］パネル（右）/ Smart Guides panel (right) */
        var smartGuidesPanel = rightColumnGroup.add("panel", undefined, L("panel.smartGuides"));
        setupPanel(smartGuidesPanel);

        var objectHighlightingCheckbox = smartGuidesPanel.add("checkbox", undefined, L("checkbox.objectHighlighting"));
        objectHighlightingCheckbox.helpTip = L("checkbox.objectHighlighting");

        /* ［パフォーマンス］パネル（右）/ Performance panel (right) */
        var performancePanel = rightColumnGroup.add("panel", undefined, L("panel.performance"));
        setupPanel(performancePanel);

        var animatedZoomCheckbox = performancePanel.add("checkbox", undefined, L("checkbox.animatedZoom"));
        animatedZoomCheckbox.helpTip = L("checkbox.animatedZoom");

        var historyStatesRow = performancePanel.add("group");
        setupRow(historyStatesRow);
        historyStatesRow.add("statictext", undefined, L("label.historyStates"));
        var historyStatesInput = historyStatesRow.add("edittext", undefined, String(NUMERIC_INPUT_RULES.historyStates.defaultValue));
        historyStatesInput.characters = 4;
        historyStatesInput.helpTip = L("hint.historyStates");

        var realTimeDrawingCheckbox = performancePanel.add("checkbox", undefined, L("checkbox.realTimeDrawing"));
        realTimeDrawingCheckbox.helpTip = L("checkbox.realTimeDrawing");

        /* ［ファイル管理］パネル（右）/ File Management panel (right) */
        var fileManagementPanel = rightColumnGroup.add("panel", undefined, L("panel.fileManagement"));
        setupPanel(fileManagementPanel);

        var editOriginalSystemDefaultCheckbox = fileManagementPanel.add("checkbox", undefined, L("checkbox.editOriginalSystemDefault"));
        editOriginalSystemDefaultCheckbox.helpTip = L("checkbox.editOriginalSystemDefault");

        var autoActivateFontsCheckbox = fileManagementPanel.add("checkbox", undefined, L("checkbox.autoActivateFonts"));
        autoActivateFontsCheckbox.helpTip = L("checkbox.autoActivateFonts");

        /* ファイルの保存先（コンピューター／クラウド）/ Save location (Computer / Cloud) */
        var saveLocationRow = fileManagementPanel.add("group");
        setupRow(saveLocationRow);
        saveLocationRow.add("statictext", undefined, L("label.saveLocation"));
        var saveToComputerRadio = saveLocationRow.add("radiobutton", undefined, L("radio.saveToComputer"));
        var saveToCloudRadio = saveLocationRow.add("radiobutton", undefined, L("radio.saveToCloud"));

        /* リンクを更新（自動／手動／確認）/ Update Links (Automatic / Manual / Ask) */
        var updateLinksRow = fileManagementPanel.add("group");
        setupRow(updateLinksRow);
        updateLinksRow.add("statictext", undefined, L("label.updateLinks"));
        var updateLinksAutoRadio = updateLinksRow.add("radiobutton", undefined, L("radio.updateLinksAuto"));
        var updateLinksManualRadio = updateLinksRow.add("radiobutton", undefined, L("radio.updateLinksManual"));
        var updateLinksAskRadio = updateLinksRow.add("radiobutton", undefined, L("radio.updateLinksAsk"));

        /* ［クリップボードの処理］パネル（右）/ Clipboard Handling panel (right) */
        var clipboardPanel = rightColumnGroup.add("panel", undefined, L("panel.clipboard"));
        setupPanel(clipboardPanel);

        var includeSvgCodeCheckbox = clipboardPanel.add("checkbox", undefined, L("checkbox.includeSvgCode"));
        includeSvgCodeCheckbox.helpTip = L("checkbox.includeSvgCode");

        /* ［パスに制限］パネル（右）/ Limit to Path panel (right) */
        var limitToPathPanel = rightColumnGroup.add("panel", undefined, L("panel.limitToPath"));
        setupPanel(limitToPathPanel);

        var objectPathOnlyCheckbox = limitToPathPanel.add("checkbox", undefined, L("checkbox.objectPathOnly"));
        objectPathOnlyCheckbox.helpTip = L("checkbox.objectPathOnly");

        var textPathOnlyCheckbox = limitToPathPanel.add("checkbox", undefined, L("checkbox.textPathOnly"));
        textPathOnlyCheckbox.helpTip = L("checkbox.textPathOnly");

        /* 環境設定キーとUIの単純な対応表。invertedInt は 0=ON / 1=OFF の反転キー */
        /* Simple preference-to-UI bindings; "invertedInt" keys store 0 = ON and 1 = OFF */
        var PREFERENCE_UI_BINDINGS = [
            { key: "showRichToolTips", valueType: "bool", control: richToolTipsCheckbox },
            { key: "Hello/ShowHomeScreenWS", valueType: "bool", control: homeScreenCheckbox },
            { key: "Hello/NewDoc", valueType: "bool", control: legacyNewDocCheckbox },
            { key: "enablePrintBleedWidget", valueType: "bool", control: printBleedWidgetCheckbox },
            { key: "moveLockedAndHiddenArt", valueType: "bool", control: moveLockedArtCheckbox },
            { key: "showArtboardLabelOnCanvas", valueType: "bool", control: showArtboardNameCheckbox },
            { key: "zoomToSelection", valueType: "bool", control: zoomToSelectionCheckbox },
            { key: "hitShapeOnPreview", valueType: "invertedInt", control: objectPathOnlyCheckbox },
            { key: "hitTypeShapeOnPreview", valueType: "invertedInt", control: textPathOnlyCheckbox },
            { key: "text/autoSizing", valueType: "bool", control: autoSizeAreaTextCheckbox },
            { key: "text/doFontLocking", valueType: "bool", control: missingGlyphProtectionCheckbox },
            { key: "text/enableAlternateGlyph", valueType: "bool", control: alternateGlyphCheckbox },
            { key: "smartGuides/showObjectHighlighting", valueType: "bool", control: objectHighlightingCheckbox },
            { key: "Performance/AnimZoom", valueType: "bool", control: animatedZoomCheckbox },
            { key: "LiveEdit_State_Machine", valueType: "bool", control: realTimeDrawingCheckbox },
            { key: "useSysDefEdit", valueType: "bool", control: editOriginalSystemDefaultCheckbox },
            { key: "AutoActivateMissingFont", valueType: "bool", control: autoActivateFontsCheckbox },
            { key: "plugin/FileClipboard/copySVGCode", valueType: "bool", control: includeSvgCodeCheckbox }
        ];

        /* 選択中のストローク幅を返す / Return the selected artboard stroke width */
        function getSelectedArtboardStrokeWidth() {
            for (var i = 0; i < artboardStrokeWidthRadios.length; i++) {
                if (artboardStrokeWidthRadios[i].value) return ARTBOARD_STROKE_WIDTHS[i];
            }
            return ARTBOARD_STROKE_WIDTHS[0];
        }

        /* ストローク幅のラジオを選択（1〜4を index に変換）/ Select the stroke-width radio for a width value (1-4) */
        function selectArtboardStrokeWidth(width) {
            var widthIndex = clamp(Math.round(width), 1, ARTBOARD_STROKE_WIDTHS.length) - 1;
            for (var i = 0; i < artboardStrokeWidthRadios.length; i++) {
                artboardStrokeWidthRadios[i].value = (i === widthIndex);
            }
        }

        /* 「リンクを更新」のラジオを選択 / Select the "Update Links" radio for a value */
        function selectUpdateLinks(value) {
            updateLinksAutoRadio.value = (value === UPDATE_LINKS_AUTO);
            updateLinksManualRadio.value = (value === UPDATE_LINKS_MANUAL);
            updateLinksAskRadio.value = (value !== UPDATE_LINKS_AUTO && value !== UPDATE_LINKS_MANUAL);
        }

        /* 「最近使用したフォント」のチェックと入力欄を同期 / Sync the recent-fonts checkbox and its input */
        function applyRecentFontsCount(count) {
            recentFontsCheckbox.value = (count > 0);
            recentFontsInput.text = String(count);
            recentFontsInput.enabled = (count > 0);
        }

        /* 現在の環境設定を読み込んでUIに反映 / Load the current preferences into the UI */
        function loadPreferencesIntoUI() {
            for (var i = 0; i < PREFERENCE_UI_BINDINGS.length; i++) {
                var binding = PREFERENCE_UI_BINDINGS[i];
                if (binding.valueType === "bool") {
                    binding.control.value = !!safeGetBool(binding.key, false);
                } else if (binding.valueType === "invertedInt") {
                    binding.control.value = (safeGetInt(binding.key, 1) === 0); /* 0がON / 0 means ON */
                }
            }

            /* 数値入力欄 / Numeric fields */
            historyStatesInput.text = String(clampIntToRule(safeGetInt("maximumUndoDepth", 0), NUMERIC_INPUT_RULES.historyStates));
            applyRecentFontsCount(clampIntToRule(safeGetInt("text/recentFontMenu/showNEntries", 0), NUMERIC_INPUT_RULES.recentFonts));

            /* アートボードのハイライト（Real値なので個別に読み込み）/ Artboard highlight (real values, read individually) */
            artboardColorDropdown.selection = findNearestArtboardColorIndex(
                safeGetReal("ArtboardBBColorRed", 0.0),
                safeGetReal("ArtboardBBColorGreen", 0.0),
                safeGetReal("ArtboardBBColorBlue", 0.0)
            );
            selectArtboardStrokeWidth(safeGetReal("ArtboardBBWidth", 1.0));

            /* ユーザーインターフェイス / User interface */
            var canvasIsWhite = (safeGetInt("uiCanvasIsWhite", 0) === 1);
            canvasWhiteRadio.value = canvasIsWhite;
            canvasMatchRadio.value = !canvasIsWhite;
            selectBrightnessSwatchByValue(safeGetReal("uiBrightness", 0.0));

            /* ガイド / Guides */
            var guideIsLightBlue = isGuideColorLightBlue();
            guideColorLightBlueRadio.value = guideIsLightBlue;
            guideColorCyanRadio.value = !guideIsLightBlue;
            var guideStyleIsDots = (safeGetInt("Guide/Style", GUIDE_STYLE_LINES) === GUIDE_STYLE_DOTS);
            guideStyleDotsRadio.value = guideStyleIsDots;
            guideStyleLinesRadio.value = !guideStyleIsDots;

            /* ファイル管理 / File management */
            var saveToCloud = safeGetBool("AdobeSaveAsCloudDocumentPreference", false);
            saveToCloudRadio.value = saveToCloud;
            saveToComputerRadio.value = !saveToCloud;
            selectUpdateLinks(safeGetInt("plugin/FileClipboard/linkoptions", UPDATE_LINKS_ASK));
        }

        /* プリセットの設定一式をUIに反映（環境設定には書き込まない）/ Load a preset into the UI (nothing is written to preferences) */
        function applyPresetStateToUI(presetState) {
            richToolTipsCheckbox.value = presetState.richToolTips;
            homeScreenCheckbox.value = presetState.homeScreen;
            legacyNewDocCheckbox.value = presetState.legacyNewDoc;
            printBleedWidgetCheckbox.value = presetState.printBleedWidget;
            moveLockedArtCheckbox.value = presetState.moveLockedArt;
            objectPathOnlyCheckbox.value = presetState.objectPathOnly;
            zoomToSelectionCheckbox.value = presetState.zoomToSelection;
            textPathOnlyCheckbox.value = presetState.textPathOnly;
            showArtboardNameCheckbox.value = presetState.showArtboardName;
            artboardColorDropdown.selection = presetState.artboardColorIndex;
            selectArtboardStrokeWidth(presetState.artboardStrokeWidth);
            autoSizeAreaTextCheckbox.value = presetState.autoSizeAreaText;
            applyRecentFontsCount(presetState.recentFontsEnabled ? presetState.recentFontsCount : 0);
            missingGlyphProtectionCheckbox.value = presetState.missingGlyphProtection;
            alternateGlyphCheckbox.value = presetState.alternateGlyph;
            objectHighlightingCheckbox.value = presetState.objectHighlighting;
            animatedZoomCheckbox.value = presetState.animatedZoom;
            historyStatesInput.text = String(presetState.historyStates);
            realTimeDrawingCheckbox.value = presetState.realTimeDrawing;
            editOriginalSystemDefaultCheckbox.value = presetState.editOriginalSystemDefault;
            autoActivateFontsCheckbox.value = presetState.autoActivateFonts;
            canvasWhiteRadio.value = presetState.canvasWhite;
            canvasMatchRadio.value = !presetState.canvasWhite;
            guideColorLightBlueRadio.value = presetState.guideColorIsLightBlue;
            guideColorCyanRadio.value = !presetState.guideColorIsLightBlue;
            saveToCloudRadio.value = presetState.saveToCloud;
            saveToComputerRadio.value = !presetState.saveToCloud;
            selectUpdateLinks(presetState.updateLinks);
            includeSvgCodeCheckbox.value = presetState.includeSvgCode;
        }

        /* プリセットIDに応じてUIを更新 / Update the UI for the given preset id */
        function applyPresetById(presetId) {
            if (presetId === PRESET_IDS.preset1) {
                applyPresetStateToUI(PRESET_STATE_1);
            } else if (presetId === PRESET_IDS.defaultPreset) {
                applyPresetStateToUI(PRESET_STATE_DEFAULT);
            } else {
                loadPreferencesIntoUI();
            }
        }

        /* ドロップダウンの選択位置をプリセットIDに変換 / Map the dropdown index to a preset id */
        function getPresetIdByIndex(selectionIndex) {
            if (selectionIndex === 1) return PRESET_IDS.defaultPreset;
            if (selectionIndex === 2) return PRESET_IDS.preset1;
            return PRESET_IDS.current;
        }

        /* ［OK］時に全項目を環境設定へ書き込む / Write every setting to the preferences on [OK] */
        function savePreferences() {
            safeSetBool("showRichToolTips", richToolTipsCheckbox.value);
            safeSetBool("Hello/ShowHomeScreenWS", homeScreenCheckbox.value);
            safeSetBool("Hello/NewDoc", legacyNewDocCheckbox.value);
            safeSetBool("enablePrintBleedWidget", printBleedWidgetCheckbox.value);
            safeSetBool("moveLockedAndHiddenArt", moveLockedArtCheckbox.value);
            safeSetBool("zoomToSelection", zoomToSelectionCheckbox.value);

            /* 「パスに制限」は 0=ON / 1=OFF の反転キー / "Limit to path" keys are inverted: 0 = ON, 1 = OFF */
            safeSetInt("hitShapeOnPreview", objectPathOnlyCheckbox.value ? 0 : 1);
            safeSetInt("hitTypeShapeOnPreview", textPathOnlyCheckbox.value ? 0 : 1);

            /* テキスト / Text */
            safeSetBool("text/autoSizing", autoSizeAreaTextCheckbox.value);
            safeSetInt("text/recentFontMenu/showNEntries",
                recentFontsCheckbox.value ? normalizeIntInput(recentFontsInput, NUMERIC_INPUT_RULES.recentFonts) : 0);
            safeSetBool("text/doFontLocking", missingGlyphProtectionCheckbox.value);
            safeSetBool("text/enableAlternateGlyph", alternateGlyphCheckbox.value);

            /* ユーザーインターフェイス / User interface */
            safeSetInt("uiCanvasIsWhite", canvasWhiteRadio.value ? 1 : 0);

            /* ガイド / Guides */
            writeGuideColor(guideColorLightBlueRadio.value ? GUIDE_COLOR_LIGHT_BLUE : GUIDE_COLOR_CYAN);
            safeSetInt("Guide/Style", guideStyleDotsRadio.value ? GUIDE_STYLE_DOTS : GUIDE_STYLE_LINES);
            safeSetBool("smartGuides/showObjectHighlighting", objectHighlightingCheckbox.value);

            /* パフォーマンス / Performance */
            safeSetBool("Performance/AnimZoom", animatedZoomCheckbox.value);
            safeSetInt("maximumUndoDepth", normalizeIntInput(historyStatesInput, NUMERIC_INPUT_RULES.historyStates));
            safeSetBool("LiveEdit_State_Machine", realTimeDrawingCheckbox.value);

            /* ファイル管理・クリップボード / File management & clipboard */
            safeSetBool("useSysDefEdit", editOriginalSystemDefaultCheckbox.value);
            safeSetBool("AutoActivateMissingFont", autoActivateFontsCheckbox.value);
            safeSetBool("AdobeSaveAsCloudDocumentPreference", !!saveToCloudRadio.value);
            safeSetInt("plugin/FileClipboard/linkoptions", getSelectedUpdateLinks());
            safeSetBool("plugin/FileClipboard/copySVGCode", includeSvgCodeCheckbox.value);

            /* アートボード / Artboard */
            safeSetBool("showArtboardLabelOnCanvas", showArtboardNameCheckbox.value);
            var selectedColor = ARTBOARD_COLOR_PRESETS[artboardColorDropdown.selection ? artboardColorDropdown.selection.index : 0];
            safeSetReal("ArtboardBBColorRed", selectedColor.red);
            safeSetReal("ArtboardBBColorGreen", selectedColor.green);
            safeSetReal("ArtboardBBColorBlue", selectedColor.blue);
            safeSetReal("ArtboardBBWidth", getSelectedArtboardStrokeWidth());
        }

        /* 選択中の「リンクを更新」の値を返す / Return the selected "Update Links" value */
        function getSelectedUpdateLinks() {
            if (updateLinksAutoRadio.value) return UPDATE_LINKS_AUTO;
            if (updateLinksManualRadio.value) return UPDATE_LINKS_MANUAL;
            return UPDATE_LINKS_ASK;
        }

        /* UIの明るさを書き込む。値の書き込みだけでは反映されないため、変更したときだけ書き込む */
        /* Write the UI brightness; writing the value alone does not apply it, so write only when it changed */
        /* @returns {boolean} 環境設定（ユーザーインターフェイス）を開く必要があるか / Whether Preferences (User Interface) must be opened */
        function saveBrightness() {
            var selectedLevel = BRIGHTNESS_LEVELS[selectedBrightnessIndex];
            if (!brightnessSwatchTouched || !selectedLevel) return false;
            if (almostEqual(safeGetReal("uiBrightness", 0.0), selectedLevel.value, BRIGHTNESS_TOLERANCE)) return false;
            safeSetReal("uiBrightness", selectedLevel.value);
            return true;
        }

        /* 環境設定の変更後に画面を強制再描画（redraw だけでは反映されないため）/ Force a redraw after preference changes (redraw alone is unreliable) */
        function forceScreenRefresh() {
            app.executeMenuCommand("zoomout");
            app.executeMenuCommand("zoomin");
        }

        /* 環境設定パネルを開く（モーダルを閉じてから呼ぶ）/ Open a Preferences panel (call after the modal dialog is closed) */
        function openPreferencePanel(menuCommand) {
            try {
                app.executeMenuCommand(menuCommand);
            } catch (e) {
                $.writeln("[" + SCRIPT_NAME + "] failed to open " + menuCommand + " / " + e);
            }
        }

        /* プリセット選択でUIを差し替え / Swap the UI when a preset is selected */
        presetDropdown.onChange = function () {
            if (!presetDropdown.selection) return;
            applyPresetById(getPresetIdByIndex(presetDropdown.selection.index));
            app.redraw();
        };

        /* チェックOFFで入力欄を無効化（0＝非表示として保存）/ Disable the field when unchecked (saved as 0 = hidden) */
        recentFontsCheckbox.onClick = function () {
            recentFontsInput.enabled = recentFontsCheckbox.value;
            if (recentFontsCheckbox.value) {
                normalizeIntInput(recentFontsInput, NUMERIC_INPUT_RULES.recentFonts);
            } else {
                recentFontsInput.text = "0";
            }
        };

        recentFontsInput.onChange = function () {
            normalizeIntInput(recentFontsInput, NUMERIC_INPUT_RULES.recentFonts);
        };

        historyStatesInput.onChange = function () {
            normalizeIntInput(historyStatesInput, NUMERIC_INPUT_RULES.historyStates);
        };

        /* 下部ボタン行（キャンセル／OK）/ Bottom button row (Cancel / OK) */
        var bottomButtonRow = dialogContentGroup.add("group");
        bottomButtonRow.orientation = "row";
        bottomButtonRow.alignChildren = ["fill", "center"];
        bottomButtonRow.alignment = ["fill", "bottom"];

        /* 伸縮するスペーサーでボタンを右寄せ / Flexible spacer that pushes the buttons to the right */
        var flexibleSpacer = bottomButtonRow.add("group");
        flexibleSpacer.alignment = ["fill", "fill"];

        var okCancelGroup = bottomButtonRow.add("group");
        okCancelGroup.orientation = "row";
        okCancelGroup.alignChildren = ["right", "center"];
        okCancelGroup.spacing = 10;
        var cancelButton = okCancelGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var okButton = okCancelGroup.add("button", undefined, L("button.ok"), { name: "ok" });

        okButton.onClick = function () {
            try {
                savePreferences();
                var willOpenUserInterface = saveBrightness();
                forceScreenRefresh();
                prefsDialog.close();
                /* 環境設定パネルはモーダルダイアログを閉じてから開く（表示中は開けないため）*/
                /* Open the Preferences panel only after this modal dialog is closed (it cannot open while it is up) */
                if (willOpenUserInterface) openPreferencePanel("UIPref");
            } catch (e) {
                alert(L("alert.savePrefsFailed") + e);
            }
        };

        /* 初期状態として現在の環境設定を読み込む / Load the current preferences as the initial UI state */
        applyPresetById(PRESET_IDS.current);

        /* ダイアログの表示位置をずらす / Shift the dialog position */
        function shiftDialogPosition(targetDialog, offsetX, offsetY) {
            targetDialog.onShow = function () {
                targetDialog.location = [targetDialog.location[0] + offsetX, targetDialog.location[1] + offsetY];
            };
        }

        /* ダイアログの不透明度を設定 / Set the dialog opacity */
        function setDialogOpacity(targetDialog, opacityValue) {
            targetDialog.opacity = opacityValue;
        }

        setDialogOpacity(prefsDialog, DIALOG_OPACITY);
        shiftDialogPosition(prefsDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        prefsDialog.show();
    }

    main();
})();
