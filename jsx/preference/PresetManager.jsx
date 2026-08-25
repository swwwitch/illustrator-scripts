#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorの主要な環境設定を、カテゴリ別に並べた1枚のダイアログでまとめて確認・変更します。
［デフォルト］／［プリセット1］を選ぶと一式の設定値をUIに反映でき、変更は［OK］でまとめて書き込まれます。

詳細は README を参照してください。

### Overview

Reviews and changes the main Illustrator preferences from a single dialog laid out by category.
Default and Preset 1 fill the whole UI with a set of values, and everything is written at once when you click OK.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManager";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-01";                   /* 更新日 / last updated */

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
    /* 「0＝非表示」はチェックボックスOFFで表現するため、入力欄の最小値は1 / "0 = hidden" is expressed by unchecking the box, so the field itself starts at 1 */
    var NUMERIC_INPUT_RULES = {
        recentFonts: { min: 1, max: 30, defaultValue: 15 },     /* 表示数 1〜30 / Visible count 1-30 */
        historyStates: { min: 1, max: 1000, defaultValue: 100 } /* 想定範囲 1〜1000 / Expected range 1-1000 */
    };

    /* 明るさ（4段階スウォッチ）のUIを表示するか。［OK］後に環境設定を開く必要があるため既定は非表示 */
    /* Whether to show the brightness swatches; off by default because it forces Preferences to open after [OK] */
    var SHOW_BRIGHTNESS_UI = false;

    /* ダイアログの不透明度 / Dialog opacity */
    var DIALOG_OPACITY = 0.98;

    /* アートボードのストローク幅の選択肢 / Selectable artboard stroke widths */
    var ARTBOARD_STROKE_WIDTHS = [1, 2, 3, 4];

    /* アンカーポイントのサイズ（スライダー4段階）が anchorSizePref に書き込む値。既定は先頭の5 */
    /* Values written to anchorSizePref by the four-step anchor point size slider; 5 is the default */
    var ANCHOR_SIZE_LEVELS = [5, 7, 9, 11];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIロケールから言語(ja/en)を判定 / Detect UI language (ja/en) from locale
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLang = getCurrentLang();

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
            homeScreen: { ja: "「ホーム画面」を表示", en: "Show the Home Screen" },
            legacyNewDoc: {
                ja: "以前の「新規ドキュメント」インターフェイス",
                en: "Legacy \"File > New\" Interface"
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
            anchorSize: { ja: "アンカーポイントのサイズ{colon}", en: "Anchor Point Size{colon}" },
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
            colorWhite: { ja: "ホワイト", en: "White" },
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
        /* ボタン / Buttons（OKは日英同一なのでリテラル）/ ("OK" is identical in both languages, so it stays a literal) */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        /* ヘルプチップ / Help tips */
        hint: {
            brightness: {
                ja: "インターフェイスカラーは直接反映できないため、［OK］後に環境設定（ユーザーインターフェイス）が開きます。矢印キー＋Return で確定してください",
                en: "Interface color can’t be applied directly, so Preferences opens after [OK]. Confirm with the arrow keys + Return."
            },
            historyStates: { ja: "ヒストリー数を設定", en: "Set history states" },
            anchorSize: {
                ja: "アンカーポイント・ハンドル・バウンディングボックスの表示サイズ（4段階）",
                en: "Display size of anchor points, handles and the bounding box (four steps)"
            },
            homeScreen: {
                ja: "ドキュメントを開いていないときに「ホーム画面」を表示",
                en: "Show the Home Screen When No Documents Are Open"
            },
            legacyNewDoc: {
                ja: "以前の「新規ドキュメント」インターフェイスを使用",
                en: "Use Legacy \"File > New\" Interface"
            }
        },
        /* 警告メッセージ / Alerts */
        alert: {
            savePrefsFailed: {
                ja: "環境設定の保存に失敗しました: ",
                en: "Failed to save preferences: "
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで多言語テキストを取り出す / Look up localized text from LABELS by dot path
     * @param {string} labelPath - "checkbox.richToolTips" のようなドット区切りのパス
     * @returns {string} 現在の言語のテキスト（見つからない場合はパス文字列）
     */
    function getLabel(labelPath) {
        var pathSegments = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathSegments.length; i++) {
            labelNode = labelNode[pathSegments[i]];
            if (!labelNode) return labelPath; /* 保険（パスをそのまま返す）/ Fallback: return the path itself */
        }
        var localizedText = labelNode[currentLang] || labelNode.en;
        if (!localizedText) return labelPath;
        return applyUISymbols(localizedText);
    }

    /**
     * {colon} などのプレースホルダを言語別の記号に展開 / Expand placeholders like {colon}
     * @param {string} templateText - 展開前のテキスト
     * @returns {string} 展開後のテキスト
     */
    function applyUISymbols(templateText) {
        return templateText.replace(/\{colon\}/g, (currentLang === "ja") ? "：" : ":");
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

    /**
     * ウィンドウの共通設定 / Apply shared window layout
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [rowSpacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, rowSpacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof rowSpacing === "number") ? rowSpacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定 / Apply shared panel layout
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [rowSpacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, rowSpacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof rowSpacing === "number") ? rowSpacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定（ラベル＋コントロールの横並び）/ Apply a horizontal row group
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [groupAlignment] - グループ自体の配置（省略時は "left"）
     * @param {number} [controlSpacing] - 要素間隔（省略時は 10）
     * @returns {void}
     */
    function setupRow(targetGroup, groupAlignment, controlSpacing) {
        targetGroup.orientation = "row";
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.alignment = groupAlignment || "left";
        targetGroup.spacing = (typeof controlSpacing === "number") ? controlSpacing : 10;
    }

    /**
     * ラベル付きチェックボックスを追加 / Add a labeled checkbox
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelPath - LABELS のドット区切りパス
     * @param {string} [hintPath] - ツールチップ用のパス。省略時はラベルをそのまま使う
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, labelPath, hintPath) {
        var newCheckbox = parentContainer.add("checkbox", undefined, getLabel(labelPath));
        newCheckbox.helpTip = getLabel(hintPath || labelPath);
        return newCheckbox;
    }

    /**
     * 見出し付きパネルを追加 / Add a titled panel with the shared layout
     * @param {Group} parentColumn - 追加先のカラム
     * @param {string} labelPath - LABELS のドット区切りパス
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentColumn, labelPath) {
        var newPanel = parentColumn.add("panel", undefined, getLabel(labelPath));
        setupPanel(newPanel);
        return newPanel;
    }

    /**
     * パネルを縦に積むカラムを追加 / Add a column that stacks panels vertically
     * @param {Group} parentContainer - 追加先のグループ
     * @returns {Group} 追加したカラムグループ
     */
    function addColumn(parentContainer) {
        var columnGroup = parentContainer.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        return columnGroup;
    }

    /**
     * ラベル＋コントロールを並べる行を追加 / Add a row that starts with a static label
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelPath - LABELS のドット区切りパス
     * @param {number} [labelWidth] - ラベル幅。指定すると右揃えで幅を固定
     * @returns {Group} 追加した行グループ（続けてコントロールを add する）
     */
    function addLabeledRow(parentContainer, labelPath, labelWidth) {
        var rowGroup = parentContainer.add("group");
        setupRow(rowGroup);
        var rowLabel = rowGroup.add("statictext", undefined, getLabel(labelPath));
        if (typeof labelWidth === "number") {
            rowLabel.preferredSize = [labelWidth, -1]; /* 高さは自動 / height stays automatic */
            rowLabel.justify = "right";
        }
        return rowGroup;
    }

    /**
     * 排他の2択ラジオをまとめて設定 / Set a mutually exclusive pair of radio buttons
     * @param {RadioButton} radioWhenTrue - 条件が真のときに選ぶラジオ
     * @param {RadioButton} radioWhenFalse - 条件が偽のときに選ぶラジオ
     * @param {boolean} isConditionMet - 条件の判定結果
     * @returns {void}
     */
    function setRadioPair(radioWhenTrue, radioWhenFalse, isConditionMet) {
        radioWhenTrue.value = !!isConditionMet;
        radioWhenFalse.value = !isConditionMet;
    }

    // =========================================
    // 環境設定の読み書き / Preference accessors
    // =========================================

    /* 注意：Illustrator は未登録のキーを読んでも例外を投げず 0 / false を返すため、
       下記の fallbackValue は「例外が出た場合」にしか効かない。0 が正当な値でないキーは
       safeGetPositiveInt() を使うこと */
    /* Note: Illustrator returns 0 / false for unknown keys instead of throwing, so the
       fallbacks below only apply on exceptions. Use safeGetPositiveInt() for keys where 0
       is not a valid value */

    /**
     * 例外を握りつぶして処理を続行する共通ラッパー（このスクリプト唯一の try）
     * Shared guard that swallows exceptions so the dialog keeps working (the only try in the script)
     * @param {function} operation - 実行する処理
     * @param {string} actionName - ログに出す処理名
     * @param {string} targetDetail - ログに出す対象（キー名など）
     * @param {*} [fallbackValue] - 失敗したときに返す値
     * @returns {*} operation の戻り値、失敗時は fallbackValue
     */
    function runSafely(operation, actionName, targetDetail, fallbackValue) {
        try {
            return operation();
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] " + actionName + ": " + targetDetail + " / " + e);
            return fallbackValue;
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

    /**
     * 整数の環境設定を読む。読み取りに失敗したら fallbackValue / Read an integer preference, falling back on failure
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} fallbackValue - 読み取りに失敗したときの値
     * @returns {number} 環境設定の値または fallbackValue
     */
    function safeGetInt(preferenceKey, fallbackValue) {
        return runSafely(function () {
            return app.preferences.getIntegerPreference(preferenceKey);
        }, "safeGetInt fallback", preferenceKey, fallbackValue);
    }

    /**
     * 正の整数として環境設定を読む。0以下は「キーが未登録」とみなして fallbackValue を返す
     * Read an integer preference as a positive number; zero or less means "key not present"
     * @param {string} preferenceKey - 環境設定キー（0以下が正当な値ではないキーに限る）
     * @param {number} fallbackValue - キーが未登録・読み取り失敗のときの値
     * @returns {number} 正の整数、または fallbackValue
     */
    function safeGetPositiveInt(preferenceKey, fallbackValue) {
        var storedValue = safeGetInt(preferenceKey, fallbackValue);
        return (storedValue > 0) ? storedValue : fallbackValue;
    }

    /**
     * 真偽値の環境設定を読む / Read a boolean preference, falling back on failure
     * @param {string} preferenceKey - 環境設定キー
     * @param {boolean} fallbackValue - 読み取りに失敗したときの値
     * @returns {boolean} 環境設定の値または fallbackValue
     */
    function safeGetBool(preferenceKey, fallbackValue) {
        return runSafely(function () {
            return app.preferences.getBooleanPreference(preferenceKey);
        }, "safeGetBool fallback", preferenceKey, fallbackValue);
    }

    /**
     * 実数の環境設定を読む / Read a real-number preference, falling back on failure
     * @param {string} preferenceKey - 環境設定キー
     * @param {number} fallbackValue - 読み取りに失敗したときの値
     * @returns {number} 環境設定の値または fallbackValue
     */
    function safeGetReal(preferenceKey, fallbackValue) {
        return runSafely(function () {
            return app.preferences.getRealPreference(preferenceKey);
        }, "safeGetReal fallback", preferenceKey, fallbackValue);
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
     * 整数に変換。数値でなければ null / Parse as an integer, or null when not numeric
     * @param {string|number} rawValue - 変換する値
     * @returns {number|null} 整数、または変換できない場合は null
     */
    function parseIntOrNull(rawValue) {
        var parsedInt = parseInt(rawValue, 10);
        return isNaN(parsedInt) ? null : parsedInt;
    }

    /**
     * ルール（min/max/defaultValue）に沿って整数へ補正 / Normalize a value against a min/max/default rule
     * @param {string|number} rawValue - 対象の値
     * @param {NumericRule} numericRule - 適用するルール
     * @returns {number} 補正後の整数
     */
    function clampIntToRule(rawValue, numericRule) {
        var parsedInt = parseIntOrNull(rawValue);
        if (parsedInt === null) return numericRule.defaultValue;
        return clamp(parsedInt, numericRule.min, numericRule.max);
    }

    /**
     * 入力欄の値を補正して書き戻す / Normalize an edittext value in place and return it
     * @param {EditText} inputField - 対象の入力欄
     * @param {NumericRule} numericRule - 適用するルール
     * @returns {number} 補正後の整数
     */
    function normalizeIntInput(inputField, numericRule) {
        var normalizedValue = clampIntToRule(inputField.text, numericRule);
        inputField.text = String(normalizedValue);
        return normalizedValue;
    }

    /**
     * 数値配列のうち目標値に最も近い要素の index を返す（離散的な選択肢へのスナップ用）
     * Index of the array element closest to a target value (snaps to discrete choices)
     * @param {number[]} candidateValues - 選択肢の数値配列
     * @param {number} targetValue - 目標値
     * @returns {number} 最も近い要素の index（配列が空なら 0）
     */
    function findNearestIndex(candidateValues, targetValue) {
        var nearestIndex = 0;
        var nearestDifference = Infinity;
        for (var i = 0; i < candidateValues.length; i++) {
            var difference = Math.abs(candidateValues[i] - targetValue);
            if (difference < nearestDifference) {
                nearestDifference = difference;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /**
     * 許容誤差つきで比較 / Compare two numbers with a tolerance
     * @param {number} firstValue - 比較する値
     * @param {number} secondValue - 比較する値
     * @param {number} [tolerance] - 許容誤差（省略時は 0.02）
     * @returns {boolean} 許容誤差の範囲内なら true
     */
    function almostEqual(firstValue, secondValue, tolerance) {
        if (typeof tolerance !== "number") tolerance = 0.02;
        return Math.abs(firstValue - secondValue) < tolerance;
    }

    // =========================================
    // 設定値の定義 / Preference value definitions
    // =========================================

    /* アートボードのハイライトカラープリセット（RGB 0..1）/ Artboard highlight color presets (RGB 0..1) */
    var ARTBOARD_COLOR_PRESETS = [
        { label: getLabel("dropdown.colorLightBlue"), red: 0.29, green: 0.52, blue: 1.0 },
        { label: getLabel("dropdown.colorLightRed"), red: 1.0, green: 0.29, blue: 0.29 },
        { label: getLabel("dropdown.colorGreen"), red: 0.0, green: 0.65, blue: 0.31 },
        { label: getLabel("dropdown.colorMediumBlue"), red: 0.0, green: 0.45, blue: 0.78 },
        { label: getLabel("dropdown.colorMagenta"), red: 1.0, green: 0.0, blue: 1.0 },
        { label: getLabel("dropdown.colorCyan"), red: 0.0, green: 1.0, blue: 1.0 },
        { label: getLabel("dropdown.colorWhite"), red: 1.0, green: 1.0, blue: 1.0 },
        { label: getLabel("dropdown.colorBlack"), red: 0.0, green: 0.0, blue: 0.0 },
        { label: getLabel("dropdown.colorYellow"), red: 1.0, green: 1.0, blue: 0.0 }
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
        anchorSize: 5, /* anchorSizePref の値（5/7/9/11）/ anchorSizePref value (5/7/9/11) */
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
        guideStyleIsDots: false,
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
        anchorSize: 7, /* anchorSizePref の値（5/7/9/11）/ anchorSizePref value (5/7/9/11) */
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
        guideStyleIsDots: false,
        saveToCloud: false,
        updateLinks: UPDATE_LINKS_AUTO,
        includeSvgCode: true
    };

    // =========================================
    // 設定値のヘルパー / Preference value helpers
    // =========================================

    /**
     * 指定RGBに最も近いハイライトカラーの index / Index of the highlight color nearest to the given RGB
     * @param {number} red - 赤成分（0〜1）
     * @param {number} green - 緑成分（0〜1）
     * @param {number} blue - 青成分（0〜1）
     * @returns {number} ARTBOARD_COLOR_PRESETS の index
     */
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

    /**
     * uiBrightness 値に最も近いプリセットの index / Index of the brightness preset closest to a uiBrightness value
     * @param {number} brightnessValue - uiBrightness の値
     * @returns {number} BRIGHTNESS_LEVELS の index
     */
    function findNearestBrightnessIndex(brightnessValue) {
        var brightnessValues = [];
        for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
            brightnessValues.push(BRIGHTNESS_LEVELS[i].value);
        }
        return findNearestIndex(brightnessValues, brightnessValue);
    }

    /**
     * 現在のガイドカラーがライトブルーか / Whether the current guide color is Light Blue
     * @returns {boolean} ライトブルーなら true
     */
    function isGuideColorLightBlue() {
        return almostEqual(safeGetReal("Guide/Color/red", GUIDE_COLOR_CYAN.red), GUIDE_COLOR_LIGHT_BLUE.red) &&
            almostEqual(safeGetReal("Guide/Color/green", GUIDE_COLOR_CYAN.green), GUIDE_COLOR_LIGHT_BLUE.green) &&
            almostEqual(safeGetReal("Guide/Color/blue", GUIDE_COLOR_CYAN.blue), GUIDE_COLOR_LIGHT_BLUE.blue);
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

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを組み立てて表示し、［OK］で環境設定を保存する / Build and show the dialog, then save preferences on [OK]
     * @returns {void}
     */
    function main() {

        /* 環境設定キー・UI・プリセット項目の対応表。UIを組み立てながら bindCheckbox() が積む */
        /* Bindings between preference key, control and preset field; filled in by bindCheckbox() while the UI is built */
        /* 読み込み・保存・プリセット適用の3か所がこの1つの表を参照する / Load, save and preset apply all read this single table */
        var PREFERENCE_UI_BINDINGS = [];

        /**
         * チェックボックスを追加し、対応表に登録する / Add a checkbox and register it in the binding table
         * @param {Panel|Group} parentPanel - 追加先
         * @param {{label: string, key: string, preset: string, type: string, hint: string}} checkboxSpec
         *        label=ラベルのパス / key=環境設定キー / preset=プリセット項目名 /
         *        type=省略時 "bool"、0=ON・1=OFF の反転キーは "invertedInt" / hint=ツールチップのパス（省略可）
         * @returns {Checkbox} 追加したチェックボックス
         */
        function bindCheckbox(parentPanel, checkboxSpec) {
            var newCheckbox = addCheckbox(parentPanel, checkboxSpec.label, checkboxSpec.hint);
            PREFERENCE_UI_BINDINGS.push({
                key: checkboxSpec.key,
                valueType: checkboxSpec.type || "bool",
                control: newCheckbox,
                presetField: checkboxSpec.preset
            });
            return newCheckbox;
        }

        /* ダイアログ本体 / Dialog window */
        var preferencesDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(preferencesDialog);

        /* 全体を縦に積むコンテナ / Vertical container for the whole dialog */
        var dialogContentGroup = preferencesDialog.add("group");
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
        presetSelectorGroup.add("statictext", undefined, getLabel("label.preset"));
        var presetLabels = [
            getLabel("dropdown.presetCurrent"),
            getLabel("dropdown.presetDefault"),
            getLabel("dropdown.preset1")
        ];
        var presetDropdown = presetSelectorGroup.add("dropdownlist", undefined, presetLabels);
        presetDropdown.selection = 0; /* 初期選択は「現在の設定」/ Default selection is "Current Settings" */

        /* パネルを左右2カラムに配置 / Two-column layout for the panels */
        var panelColumnsGroup = dialogContentGroup.add("group");
        panelColumnsGroup.orientation = "row";
        panelColumnsGroup.alignChildren = "top";
        panelColumnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = addColumn(panelColumnsGroup);
        var rightColumnGroup = addColumn(panelColumnsGroup);

        /* ［一般］パネル（左）/ General panel (left) */
        var generalPanel = addPanel(leftColumnGroup, "panel.general");
        bindCheckbox(generalPanel, { label: "checkbox.richToolTips", key: "showRichToolTips", preset: "richToolTips" });
        bindCheckbox(generalPanel, { label: "checkbox.homeScreen", key: "Hello/ShowHomeScreenWS", preset: "homeScreen", hint: "hint.homeScreen" });
        bindCheckbox(generalPanel, { label: "checkbox.legacyNewDoc", key: "Hello/NewDoc", preset: "legacyNewDoc", hint: "hint.legacyNewDoc" });
        bindCheckbox(generalPanel, { label: "checkbox.printBleedWidget", key: "enablePrintBleedWidget", preset: "printBleedWidget" });

        /* ［選択範囲・アンカー表示］パネル（左）/ Selection & Anchor Display panel (left) */
        var selectionAnchorPanel = addPanel(leftColumnGroup, "panel.selectionAnchor");
        bindCheckbox(selectionAnchorPanel, { label: "checkbox.zoomToSelection", key: "zoomToSelection", preset: "zoomToSelection" });

        /* アンカーポイントのサイズ（4段階スライダー）/ Anchor point size (four-step slider) */
        var anchorSizeRow = addLabeledRow(selectionAnchorPanel, "label.anchorSize");
        var anchorSizeSlider = anchorSizeRow.add("slider", undefined, 1, 1, ANCHOR_SIZE_LEVELS.length);
        anchorSizeSlider.preferredSize = [110, -1];
        anchorSizeSlider.helpTip = getLabel("hint.anchorSize");

        /* ［アートボード］パネル（左）/ Artboard panel (left) */
        var artboardPanel = addPanel(leftColumnGroup, "panel.artboard");
        bindCheckbox(artboardPanel, { label: "checkbox.moveLockedArt", key: "moveLockedAndHiddenArt", preset: "moveLockedArt" });
        bindCheckbox(artboardPanel, { label: "checkbox.showArtboardName", key: "showArtboardLabelOnCanvas", preset: "showArtboardName" });

        /* ハイライトのカラー（9色のプリセットから選択）/ Highlight color (nine presets) */
        var artboardColorRow = addLabeledRow(artboardPanel, "label.artboardColor");
        var artboardColorLabels = [];
        for (var i = 0; i < ARTBOARD_COLOR_PRESETS.length; i++) {
            artboardColorLabels.push(ARTBOARD_COLOR_PRESETS[i].label);
        }
        var artboardColorDropdown = artboardColorRow.add("dropdownlist", undefined, artboardColorLabels);

        /* ストロークの幅（1〜4）/ Stroke width (1-4) */
        var artboardStrokeWidthRow = addLabeledRow(artboardPanel, "label.artboardStrokeWidth");
        var artboardStrokeWidthRadios = [];
        for (var i = 0; i < ARTBOARD_STROKE_WIDTHS.length; i++) {
            artboardStrokeWidthRadios.push(
                artboardStrokeWidthRow.add("radiobutton", undefined, String(ARTBOARD_STROKE_WIDTHS[i]))
            );
        }

        /* ［テキスト］パネル（左）/ Text panel (left) */
        var textPanel = addPanel(leftColumnGroup, "panel.text");
        bindCheckbox(textPanel, { label: "checkbox.autoSizeAreaText", key: "text/autoSizing", preset: "autoSizeAreaText" });

        /* 最近使用したフォントの表示数（チェックOFFで0＝非表示）/ Recent font count (unchecked means 0 = hidden) */
        var recentFontsRow = textPanel.add("group");
        setupRow(recentFontsRow);
        var recentFontsCheckbox = addCheckbox(recentFontsRow, "checkbox.recentFonts");
        var recentFontsInput = recentFontsRow.add("edittext", undefined, "0");
        recentFontsInput.characters = 3;

        bindCheckbox(textPanel, { label: "checkbox.missingGlyphProtection", key: "text/doFontLocking", preset: "missingGlyphProtection" });
        bindCheckbox(textPanel, { label: "checkbox.alternateGlyph", key: "text/enableAlternateGlyph", preset: "alternateGlyph" });

        /* ［ユーザーインターフェイス］パネル（左）/ User Interface panel (left) */
        var userInterfacePanel = addPanel(leftColumnGroup, "panel.userInterface");

        /* 明るさ（ラベル＋4段階のスウォッチ）。SHOW_BRIGHTNESS_UI で表示を切り替える */
        /* Brightness (label + four swatches); toggled by SHOW_BRIGHTNESS_UI */
        var brightnessSwatchGroup = null;
        if (SHOW_BRIGHTNESS_UI) {
            var brightnessRow = addLabeledRow(userInterfacePanel, "label.brightness");
            brightnessSwatchGroup = brightnessRow.add("group");
            setupRow(brightnessSwatchGroup, "left", 8);
        }

        var brightnessSwatchButtons = [];
        var selectedBrightnessIndex = -1;
        var brightnessSwatchTouched = false; /* ユーザーがスウォッチを操作したか（未操作なら書き込まない）/ Whether the user clicked a swatch (no write when untouched) */

        /**
         * スウォッチの描画ハンドラを作る：塗り＋（選択時のみ）青い枠
         * Build a swatch draw handler: fill + (only when selected) a blue border
         * @param {number} levelIndex - BRIGHTNESS_LEVELS の index
         * @returns {function} onDraw に割り当てる関数
         */
        function createBrightnessSwatchDrawHandler(levelIndex) {
            return function () {
                var swatchGraphics = this.graphics;
                var swatchWidth = this.size[0];
                var swatchHeight = this.size[1];
                var brightnessLevel = BRIGHTNESS_LEVELS[levelIndex];

                /* シェードで塗りつぶし / Fill with the shade */
                var fillBrush = swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, brightnessLevel.shade.concat(1));
                swatchGraphics.newPath();
                swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
                swatchGraphics.fillPath(fillBrush);

                if (levelIndex === selectedBrightnessIndex) {
                    /* 選択：太めの青枠 / Selected: thicker blue border */
                    var selectedPen = swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, BRIGHTNESS_SELECTED_BORDER.concat(1), 2);
                    swatchGraphics.newPath();
                    swatchGraphics.rectPath(1, 1, swatchWidth - 2, swatchHeight - 2);
                    swatchGraphics.strokePath(selectedPen);
                } else {
                    /* 非選択：細いグレー枠（明シェードを明背景でも視認）/ Not selected: thin gray outline (keep light swatches visible) */
                    var outlinePen = swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, BRIGHTNESS_SWATCH_OUTLINE.concat(1), 1);
                    swatchGraphics.newPath();
                    swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
                    swatchGraphics.strokePath(outlinePen);
                }
            };
        }

        /**
         * 全スウォッチを再描画（hide/show で onDraw を確実に再実行し、旧選択枠を残さない＝排他表示）
         * Force all swatches to repaint (hide/show reliably re-runs onDraw so the old selection border never lingers = exclusive)
         * @returns {void}
         */
        function refreshBrightnessSwatches() {
            for (var i = 0; i < brightnessSwatchButtons.length; i++) {
                brightnessSwatchButtons[i].hide();
                brightnessSwatchButtons[i].show();
            }
        }

        /**
         * uiBrightness 値から選択状態だけ更新（書き込みは［OK］時）
         * Update the selection only, from a uiBrightness value (writes happen on [OK])
         * @param {number} brightnessValue - uiBrightness の値
         * @returns {void}
         */
        function selectBrightnessSwatchByValue(brightnessValue) {
            var nearestIndex = findNearestBrightnessIndex(brightnessValue);
            if (nearestIndex === selectedBrightnessIndex) return;
            selectedBrightnessIndex = nearestIndex;
            refreshBrightnessSwatches();
        }

        /* スウォッチを4つ生成 / Build the four swatches */
        if (SHOW_BRIGHTNESS_UI) {
            for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
                var brightnessSwatchButton = brightnessSwatchGroup.add("iconbutton", undefined, undefined, { style: "toolbutton" });
                brightnessSwatchButton.preferredSize = [BRIGHTNESS_SWATCH_SIZE, BRIGHTNESS_SWATCH_SIZE];
                brightnessSwatchButton.alignment = ["left", "center"];
                brightnessSwatchButton.helpTip = getLabel(BRIGHTNESS_LEVELS[i].labelPath) + " / " + getLabel("hint.brightness");
                brightnessSwatchButton.onDraw = createBrightnessSwatchDrawHandler(i);
                brightnessSwatchButtons.push(brightnessSwatchButton);
                (function (index) {
                    brightnessSwatchButtons[index].onClick = function () {
                        /* UIのみ更新し、保存は［OK］時 / UI only; persist on [OK] */
                        selectedBrightnessIndex = index;
                        brightnessSwatchTouched = true;
                        refreshBrightnessSwatches();
                    };
                })(brightnessSwatchButtons.length - 1);
            }
        }

        /* カンバスカラー（UIに合わせる／ホワイト）/ Canvas color (Match Brightness / White) */
        var canvasColorRow = addLabeledRow(userInterfacePanel, "label.canvasColor");
        var canvasMatchRadio = canvasColorRow.add("radiobutton", undefined, getLabel("radio.canvasMatch"));
        var canvasWhiteRadio = canvasColorRow.add("radiobutton", undefined, getLabel("radio.canvasWhite"));

        /* ［ガイド］パネル（右）/ Guides panel (right) */
        var guidesPanel = addPanel(rightColumnGroup, "panel.guides");

        var GUIDE_LABEL_WIDTH = 80; /* ガイド内ラベルの共通幅 / Unified label width inside Guides */

        /* ガイドのカラー（シアン／ライトブルー）/ Guide color (Cyan / Light Blue) */
        var guideColorRow = addLabeledRow(guidesPanel, "label.guideColor", GUIDE_LABEL_WIDTH);
        var guideColorCyanRadio = guideColorRow.add("radiobutton", undefined, getLabel("radio.guideColorCyan"));
        var guideColorLightBlueRadio = guideColorRow.add("radiobutton", undefined, getLabel("radio.guideColorLightBlue"));

        /* ガイドのスタイル（ライン／点線）/ Guide style (Lines / Dots) */
        var guideStyleRow = addLabeledRow(guidesPanel, "label.guideStyle", GUIDE_LABEL_WIDTH);
        var guideStyleLinesRadio = guideStyleRow.add("radiobutton", undefined, getLabel("radio.guideStyleLines"));
        var guideStyleDotsRadio = guideStyleRow.add("radiobutton", undefined, getLabel("radio.guideStyleDots"));

        /* ［スマートガイド］パネル（右）/ Smart Guides panel (right) */
        var smartGuidesPanel = addPanel(rightColumnGroup, "panel.smartGuides");
        bindCheckbox(smartGuidesPanel, { label: "checkbox.objectHighlighting", key: "smartGuides/showObjectHighlighting", preset: "objectHighlighting" });

        /* ［パフォーマンス］パネル（右）/ Performance panel (right) */
        var performancePanel = addPanel(rightColumnGroup, "panel.performance");
        bindCheckbox(performancePanel, { label: "checkbox.animatedZoom", key: "Performance/AnimZoom", preset: "animatedZoom" });

        var historyStatesRow = addLabeledRow(performancePanel, "label.historyStates");
        var historyStatesInput = historyStatesRow.add("edittext", undefined, String(NUMERIC_INPUT_RULES.historyStates.defaultValue));
        historyStatesInput.characters = 4;
        historyStatesInput.helpTip = getLabel("hint.historyStates");

        bindCheckbox(performancePanel, { label: "checkbox.realTimeDrawing", key: "LiveEdit_State_Machine", preset: "realTimeDrawing" });

        /* ［ファイル管理］パネル（右）/ File Management panel (right) */
        var fileManagementPanel = addPanel(rightColumnGroup, "panel.fileManagement");
        bindCheckbox(fileManagementPanel, { label: "checkbox.editOriginalSystemDefault", key: "useSysDefEdit", preset: "editOriginalSystemDefault" });
        bindCheckbox(fileManagementPanel, { label: "checkbox.autoActivateFonts", key: "AutoActivateMissingFont", preset: "autoActivateFonts" });

        /* ファイルの保存先（コンピューター／クラウド）/ Save location (Computer / Cloud) */
        var saveLocationRow = addLabeledRow(fileManagementPanel, "label.saveLocation");
        var saveToComputerRadio = saveLocationRow.add("radiobutton", undefined, getLabel("radio.saveToComputer"));
        var saveToCloudRadio = saveLocationRow.add("radiobutton", undefined, getLabel("radio.saveToCloud"));

        /* リンクを更新（自動／手動／確認）/ Update Links (Automatic / Manual / Ask) */
        var updateLinksRow = addLabeledRow(fileManagementPanel, "label.updateLinks");
        var updateLinksAutoRadio = updateLinksRow.add("radiobutton", undefined, getLabel("radio.updateLinksAuto"));
        var updateLinksManualRadio = updateLinksRow.add("radiobutton", undefined, getLabel("radio.updateLinksManual"));
        var updateLinksAskRadio = updateLinksRow.add("radiobutton", undefined, getLabel("radio.updateLinksAsk"));

        /* ［クリップボードの処理］パネル（右）/ Clipboard Handling panel (right) */
        var clipboardPanel = addPanel(rightColumnGroup, "panel.clipboard");
        bindCheckbox(clipboardPanel, { label: "checkbox.includeSvgCode", key: "plugin/FileClipboard/copySVGCode", preset: "includeSvgCode" });

        /* ［パスに制限］パネル（右）。0=ON / 1=OFF の反転キー / Limit to Path panel; these keys store 0 = ON and 1 = OFF */
        var limitToPathPanel = addPanel(rightColumnGroup, "panel.limitToPath");
        bindCheckbox(limitToPathPanel, { label: "checkbox.objectPathOnly", key: "hitShapeOnPreview", preset: "objectPathOnly", type: "invertedInt" });
        bindCheckbox(limitToPathPanel, { label: "checkbox.textPathOnly", key: "hitTypeShapeOnPreview", preset: "textPathOnly", type: "invertedInt" });

        /**
         * 選択中のストローク幅を返す / Return the selected artboard stroke width
         * @returns {number} 選択中の幅（未選択なら先頭の値）
         */
        function getSelectedArtboardStrokeWidth() {
            for (var i = 0; i < artboardStrokeWidthRadios.length; i++) {
                if (artboardStrokeWidthRadios[i].value) return ARTBOARD_STROKE_WIDTHS[i];
            }
            return ARTBOARD_STROKE_WIDTHS[0];
        }

        /**
         * ストローク幅のラジオを選択（最も近い選択肢に丸める）
         * Select the stroke-width radio, snapping to the nearest available width
         * @param {number} strokeWidth - ストローク幅
         * @returns {void}
         */
        function selectArtboardStrokeWidth(strokeWidth) {
            var widthIndex = findNearestIndex(ARTBOARD_STROKE_WIDTHS, strokeWidth);
            for (var i = 0; i < artboardStrokeWidthRadios.length; i++) {
                artboardStrokeWidthRadios[i].value = (i === widthIndex);
            }
        }

        /**
         * アンカーポイントサイズのスライダーを指定段階に合わせる / Move the anchor size slider to a step
         * @param {number} anchorStep - 1〜ANCHOR_SIZE_LEVELS.length の段階番号
         * @returns {void}
         */
        function applyAnchorSizeStep(anchorStep) {
            anchorSizeSlider.value = clamp(Math.round(anchorStep), 1, ANCHOR_SIZE_LEVELS.length);
        }

        /**
         * anchorSizePref の値から最も近い段階にスナップ / Snap the slider to the step nearest a stored value
         * @param {number} anchorSizeValue - anchorSizePref の値
         * @returns {void}
         */
        function selectAnchorSize(anchorSizeValue) {
            applyAnchorSizeStep(findNearestIndex(ANCHOR_SIZE_LEVELS, anchorSizeValue) + 1);
        }

        /**
         * 選択中の段階に対応する anchorSizePref の値を返す
         * Return the anchorSizePref value for the selected step
         * @returns {number} ANCHOR_SIZE_LEVELS のいずれかの値
         */
        function getSelectedAnchorSize() {
            var selectedStep = clamp(Math.round(anchorSizeSlider.value), 1, ANCHOR_SIZE_LEVELS.length);
            return ANCHOR_SIZE_LEVELS[selectedStep - 1];
        }

        /**
         * 「リンクを更新」のラジオを選択 / Select the "Update Links" radio for a value
         * @param {number} updateLinksValue - UPDATE_LINKS_* のいずれか
         * @returns {void}
         */
        function selectUpdateLinks(updateLinksValue) {
            updateLinksAutoRadio.value = (updateLinksValue === UPDATE_LINKS_AUTO);
            updateLinksManualRadio.value = (updateLinksValue === UPDATE_LINKS_MANUAL);
            updateLinksAskRadio.value = (updateLinksValue !== UPDATE_LINKS_AUTO && updateLinksValue !== UPDATE_LINKS_MANUAL);
        }

        /**
         * 「最近使用したフォント」のチェックと入力欄を同期
         * Sync the recent-fonts checkbox and its input
         * @param {number} recentFontCount - 表示数（0以下はチェックOFF＝非表示）
         * @returns {void}
         */
        function applyRecentFontsCount(recentFontCount) {
            var isListVisible = (recentFontCount > 0);
            recentFontsCheckbox.value = isListVisible;
            /* ONのときだけルール（1〜30）で補正。OFFは0を表示 / Normalize only when enabled; show 0 when off */
            recentFontsInput.text = String(isListVisible ? clampIntToRule(recentFontCount, NUMERIC_INPUT_RULES.recentFonts) : 0);
            recentFontsInput.enabled = isListVisible;
        }

        /**
         * 対応表のチェックボックスを環境設定から復元 / Restore the table-driven checkboxes from the preferences
         * @returns {void}
         */
        function loadBindingsIntoUI() {
            for (var i = 0; i < PREFERENCE_UI_BINDINGS.length; i++) {
                var binding = PREFERENCE_UI_BINDINGS[i];
                if (binding.valueType === "bool") {
                    binding.control.value = !!safeGetBool(binding.key, false);
                } else if (binding.valueType === "invertedInt") {
                    binding.control.value = (safeGetInt(binding.key, 1) === 0); /* 0がON / 0 means ON */
                }
            }
        }

        /**
         * 数値入力欄を環境設定から復元 / Restore the numeric fields from the preferences
         * @returns {void}
         */
        function loadNumericInputsIntoUI() {
            /* ヒストリー数は 0 を取り得ないので、0＝キー未登録とみなして既定値に戻す */
            /* History states can never be 0, so treat 0 as "key not present" and fall back to the default */
            historyStatesInput.text = String(clampIntToRule(
                safeGetPositiveInt("maximumUndoDepth", NUMERIC_INPUT_RULES.historyStates.defaultValue),
                NUMERIC_INPUT_RULES.historyStates
            ));
            /* 0 は「一覧を非表示」なのでそのまま渡す / 0 means "hide the list", so pass it through */
            applyRecentFontsCount(safeGetInt("text/recentFontMenu/showNEntries", 0));
        }

        /**
         * ドロップダウン・ラジオを環境設定から復元 / Restore the dropdowns and radio groups from the preferences
         * @returns {void}
         */
        function loadChoicesIntoUI() {
            /* アートボードのハイライト（Real値なので個別に読み込み）/ Artboard highlight (real values, read individually) */
            artboardColorDropdown.selection = findNearestArtboardColorIndex(
                safeGetReal("ArtboardBBColorRed", 0.0),
                safeGetReal("ArtboardBBColorGreen", 0.0),
                safeGetReal("ArtboardBBColorBlue", 0.0)
            );
            selectArtboardStrokeWidth(safeGetReal("ArtboardBBWidth", 1.0));

            /* 選択範囲・アンカー表示（キー未登録の 0 は最小値の段階に落ちる）/ Selection & anchor display (a missing key reads 0 and snaps to the smallest step) */
            selectAnchorSize(safeGetInt("anchorSizePref", ANCHOR_SIZE_LEVELS[0]));

            /* ユーザーインターフェイス / User interface */
            setRadioPair(canvasWhiteRadio, canvasMatchRadio, safeGetInt("uiCanvasIsWhite", 0) === 1);
            selectBrightnessSwatchByValue(safeGetReal("uiBrightness", 0.0));

            /* ガイド / Guides */
            setRadioPair(guideColorLightBlueRadio, guideColorCyanRadio, isGuideColorLightBlue());
            setRadioPair(guideStyleDotsRadio, guideStyleLinesRadio,
                safeGetInt("Guide/Style", GUIDE_STYLE_LINES) === GUIDE_STYLE_DOTS);

            /* ファイル管理 / File management */
            setRadioPair(saveToCloudRadio, saveToComputerRadio,
                safeGetBool("AdobeSaveAsCloudDocumentPreference", false));
            selectUpdateLinks(safeGetInt("plugin/FileClipboard/linkoptions", UPDATE_LINKS_ASK));
        }

        /**
         * 現在の環境設定を読み込んでUIに反映 / Load the current preferences into the UI
         * @returns {void}
         */
        function loadPreferencesIntoUI() {
            loadBindingsIntoUI();
            loadNumericInputsIntoUI();
            loadChoicesIntoUI();
        }

        /**
         * プリセットの設定一式をUIに反映（環境設定には書き込まない）
         * Load a preset into the UI (nothing is written to preferences)
         * @param {object} presetState - PRESET_STATE_DEFAULT / PRESET_STATE_1 と同じ形の設定一式
         * @returns {void}
         */
        function applyPresetStateToUI(presetState) {
            /* チェックボックスは対応表の presetField 経由 / Checkboxes come from the table's presetField */
            for (var i = 0; i < PREFERENCE_UI_BINDINGS.length; i++) {
                var binding = PREFERENCE_UI_BINDINGS[i];
                binding.control.value = !!presetState[binding.presetField];
            }

            /* 対応表に載らない項目 / Fields the table doesn't cover */
            artboardColorDropdown.selection = presetState.artboardColorIndex;
            selectArtboardStrokeWidth(presetState.artboardStrokeWidth);
            selectAnchorSize(presetState.anchorSize);
            applyRecentFontsCount(presetState.recentFontsEnabled ? presetState.recentFontsCount : 0);
            historyStatesInput.text = String(presetState.historyStates);
            setRadioPair(canvasWhiteRadio, canvasMatchRadio, presetState.canvasWhite);
            setRadioPair(guideColorLightBlueRadio, guideColorCyanRadio, presetState.guideColorIsLightBlue);
            setRadioPair(guideStyleDotsRadio, guideStyleLinesRadio, presetState.guideStyleIsDots);
            setRadioPair(saveToCloudRadio, saveToComputerRadio, presetState.saveToCloud);
            selectUpdateLinks(presetState.updateLinks);
        }

        /**
         * プリセットIDに応じてUIを更新 / Update the UI for the given preset id
         * @param {string} presetId - PRESET_IDS のいずれか
         * @returns {void}
         */
        function applyPresetById(presetId) {
            if (presetId === PRESET_IDS.preset1) {
                applyPresetStateToUI(PRESET_STATE_1);
            } else if (presetId === PRESET_IDS.defaultPreset) {
                applyPresetStateToUI(PRESET_STATE_DEFAULT);
            } else {
                loadPreferencesIntoUI();
            }
        }

        /**
         * ドロップダウンの選択位置をプリセットIDに変換 / Map the dropdown index to a preset id
         * @param {number} selectionIndex - ドロップダウンの index
         * @returns {string} PRESET_IDS のいずれか
         */
        function getPresetIdByIndex(selectionIndex) {
            if (selectionIndex === 1) return PRESET_IDS.defaultPreset;
            if (selectionIndex === 2) return PRESET_IDS.preset1;
            return PRESET_IDS.current;
        }

        /**
         * 対応表（PREFERENCE_UI_BINDINGS）のキーをまとめて書き込む
         * Write every key listed in PREFERENCE_UI_BINDINGS
         * @returns {void}
         */
        function savePreferenceBindings() {
            for (var i = 0; i < PREFERENCE_UI_BINDINGS.length; i++) {
                var binding = PREFERENCE_UI_BINDINGS[i];
                if (binding.valueType === "bool") {
                    safeSetBool(binding.key, !!binding.control.value);
                } else if (binding.valueType === "invertedInt") {
                    safeSetInt(binding.key, binding.control.value ? 0 : 1); /* 0がON / 0 means ON */
                }
            }
        }

        /**
         * ［OK］時に全項目を環境設定へ書き込む / Write every setting to the preferences on [OK]
         * @returns {void}
         */
        function savePreferences() {
            /* 対応表で扱えるキー（チェックボックス）/ Keys covered by the binding table (checkboxes) */
            savePreferenceBindings();

            /* テキスト / Text */
            safeSetInt("text/recentFontMenu/showNEntries",
                recentFontsCheckbox.value ? normalizeIntInput(recentFontsInput, NUMERIC_INPUT_RULES.recentFonts) : 0);

            /* 選択範囲・アンカー表示 / Selection & anchor display */
            safeSetInt("anchorSizePref", getSelectedAnchorSize());

            /* ユーザーインターフェイス / User interface */
            safeSetInt("uiCanvasIsWhite", canvasWhiteRadio.value ? 1 : 0);

            /* ガイド / Guides */
            writeGuideColor(guideColorLightBlueRadio.value ? GUIDE_COLOR_LIGHT_BLUE : GUIDE_COLOR_CYAN);
            safeSetInt("Guide/Style", guideStyleDotsRadio.value ? GUIDE_STYLE_DOTS : GUIDE_STYLE_LINES);

            /* パフォーマンス / Performance */
            safeSetInt("maximumUndoDepth", normalizeIntInput(historyStatesInput, NUMERIC_INPUT_RULES.historyStates));

            /* ファイル管理・クリップボード / File management & clipboard */
            safeSetBool("AdobeSaveAsCloudDocumentPreference", !!saveToCloudRadio.value);
            safeSetInt("plugin/FileClipboard/linkoptions", getSelectedUpdateLinks());

            /* アートボード / Artboard */
            var selectedArtboardColor = ARTBOARD_COLOR_PRESETS[artboardColorDropdown.selection ? artboardColorDropdown.selection.index : 0];
            safeSetReal("ArtboardBBColorRed", selectedArtboardColor.red);
            safeSetReal("ArtboardBBColorGreen", selectedArtboardColor.green);
            safeSetReal("ArtboardBBColorBlue", selectedArtboardColor.blue);
            safeSetReal("ArtboardBBWidth", getSelectedArtboardStrokeWidth());
        }

        /**
         * 選択中の「リンクを更新」の値を返す / Return the selected "Update Links" value
         * @returns {number} UPDATE_LINKS_* のいずれか
         */
        function getSelectedUpdateLinks() {
            if (updateLinksAutoRadio.value) return UPDATE_LINKS_AUTO;
            if (updateLinksManualRadio.value) return UPDATE_LINKS_MANUAL;
            return UPDATE_LINKS_ASK;
        }

        /**
         * UIの明るさを書き込む。値の書き込みだけでは反映されないため、変更したときだけ書き込む
         * Write the UI brightness; writing the value alone does not apply it, so write only when it changed
         * @returns {boolean} 環境設定（ユーザーインターフェイス）を開く必要があるか
         */
        function saveBrightness() {
            var selectedBrightnessLevel = BRIGHTNESS_LEVELS[selectedBrightnessIndex];
            if (!brightnessSwatchTouched || !selectedBrightnessLevel) return false;
            if (almostEqual(safeGetReal("uiBrightness", 0.0), selectedBrightnessLevel.value, BRIGHTNESS_TOLERANCE)) return false;
            safeSetReal("uiBrightness", selectedBrightnessLevel.value);
            return true;
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
         * 環境設定パネルを開く（モーダルを閉じてから呼ぶ）
         * Open a Preferences panel (call after the modal dialog is closed)
         * @param {string} menuCommand - メニューコマンド名（例: "UIPref"）
         * @returns {void}
         */
        function openPreferencePanel(menuCommand) {
            runSafely(function () {
                app.executeMenuCommand(menuCommand);
            }, "openPreferencePanel failed", menuCommand);
        }

        /* プリセット選択でUIを差し替え / Swap the UI when a preset is selected */
        presetDropdown.onChange = function () {
            if (!presetDropdown.selection) return;
            applyPresetById(getPresetIdByIndex(presetDropdown.selection.index));
            app.redraw();
        };

        /* チェックOFFで入力欄を無効化（0＝非表示として保存）/ Disable the field when unchecked (saved as 0 = hidden) */
        recentFontsCheckbox.onClick = function () {
            if (!recentFontsCheckbox.value) {
                applyRecentFontsCount(0);
                return;
            }
            /* ONに戻したとき 0 のままでは矛盾するので既定値を入れる / Restore the default instead of leaving 0 when re-enabled */
            var currentFontCount = parseIntOrNull(recentFontsInput.text);
            if (currentFontCount === null || currentFontCount < NUMERIC_INPUT_RULES.recentFonts.min) {
                currentFontCount = NUMERIC_INPUT_RULES.recentFonts.defaultValue;
            }
            applyRecentFontsCount(currentFontCount);
        };

        recentFontsInput.onChange = function () {
            normalizeIntInput(recentFontsInput, NUMERIC_INPUT_RULES.recentFonts);
        };

        /* ドラッグを離したときに段階へスナップ / Snap to a step when the drag is released */
        anchorSizeSlider.onChange = function () {
            applyAnchorSizeStep(anchorSizeSlider.value);
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
        var cancelButton = okCancelGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = okCancelGroup.add("button", undefined, "OK", { name: "ok" });

        okButton.onClick = function () {
            var willOpenUserInterface = false;
            try {
                savePreferences();
                willOpenUserInterface = saveBrightness();
            } catch (e) {
                alert(getLabel("alert.savePrefsFailed") + e);
            }
            /* 書き込みの成否にかかわらずダイアログは閉じる / Close the dialog whether or not the writes succeeded */
            forceScreenRefresh();
            preferencesDialog.close();
            /* 環境設定パネルはモーダルダイアログを閉じてから開く（表示中は開けないため）*/
            /* Open the Preferences panel only after this modal dialog is closed (it cannot open while it is up) */
            if (willOpenUserInterface) openPreferencePanel("UIPref");
        };

        /* 初期状態として現在の環境設定を読み込む / Load the current preferences as the initial UI state */
        applyPresetById(PRESET_IDS.current);

        /**
         * ダイアログの不透明度を設定 / Set the dialog opacity
         * @param {Window} targetDialog - 対象のダイアログ
         * @param {number} opacityValue - 不透明度（0〜1）
         * @returns {void}
         */
        function setDialogOpacity(targetDialog, opacityValue) {
            targetDialog.opacity = opacityValue;
        }

        setDialogOpacity(preferencesDialog, DIALOG_OPACITY);
        preferencesDialog.show();
    }

    main();
})();
