#target illustrator
#targetengine "PDFAIImporterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PDF/AI ファイルを指定したページ範囲で読み込み、現在のドキュメントまたは新規ドキュメント上に配置します。
各ページを個別のアートボードとして並べるか、アートボードを追加せずオブジェクトとして配置するかを選べます。横長の見開きページは左右に分割できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAIImporter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n42595650216f

### Overview

Imports a PDF/AI file over a given page range and places the pages in the current or a new document.
The pages can be laid out as one artboard each, or placed as objects without adding artboards. Landscape spreads can be split left and right.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAIImporter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PDFAIImporter";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PDFAIImporter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PDFAIImporter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n42595650216f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    // アートボードごと配置時のデフォルト間隔（pt）/ Default gap for per-artboard placement (pt)
    var DEFAULT_ARTBOARD_GAP = 100;

    // 自動列のときの最大行幅（pt）。これを超えると次行へ折り返す / Max row width for Auto columns (pt); wraps beyond this
    var MAX_AUTO_ROW_WIDTH = (220 / 2) * 72; // 7920 pt

    /* 見開きと判定する横長比（幅 > 高さ × この値） / Aspect ratio that marks a page as a spread (width > height x this) */
    var SPREAD_ASPECT_RATIO = 1.2;

    /* 綴じ方向の判定で読むPDFの行数 / Number of PDF lines scanned for the binding direction */
    var BINDING_SCAN_LINE_LIMIT = 200;

    /* 新規ドキュメントのラスタライズ効果解像度（ppi） / Raster effects resolution of a new document (ppi) */
    var NEW_DOCUMENT_RASTER_RESOLUTION = 300;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 */

    var LABELS = {
        dialog: {
            title: { ja: "PDF/AI配置", en: "PDF/AI Placement" },
            pickFile: { ja: "PDF/AIを選択してください", en: "Select a PDF/AI" }
        },
        panel: {
            source: { ja: "読み込みファイル", en: "Source File" },
            pages: { ja: "対象ページ", en: "Pages" },
            mode: { ja: "配置方法", en: "Placement Method" },
            placement: { ja: "レイアウト", en: "Layout" },
            option: { ja: "オプション", en: "Options" },
            spread: { ja: "見開き", en: "Spreads" },
            destination: { ja: "配置先", en: "Destination" },
            kei: { ja: "枠線", en: "Stroke" }
        },
        radio: {
            rangeAll: { ja: "全ページ", en: "All Pages" },
            rangeFirst: { ja: "先頭ページのみ", en: "First Page Only" },
            rangeCustom: { ja: "指定ページ", en: "Custom Pages" },
            perArtboard: { ja: "アートボードごと", en: "Per Artboard" },
            ignoreArtboard: { ja: "アートボードを無視", en: "Place as Objects" },
            keiNone: { ja: "なし", en: "None" },
            keiClipGroup: { ja: "枠線を追加", en: "Add stroke" },
            evenPageRight: { ja: "右", en: "Right" },
            evenPageLeft: { ja: "左", en: "Left" },
            destinationCurrent: { ja: "現在のドキュメント", en: "Current document" },
            destinationNew: { ja: "新規ドキュメント", en: "New document" },
            colorModeCMYK: { ja: "CMYK", en: "CMYK" },
            colorModeRGB: { ja: "RGB", en: "RGB" }
        },
        checkbox: {
            roundCorner: { ja: "角丸", en: "Round corners" },
            splitSpreads: { ja: "左右に分割", en: "Split left and right" }
        },
        dropdown: {
            cropArt: { ja: "アート", en: "Art" },
            cropCrop: { ja: "トリミング", en: "Crop" },
            cropTrim: { ja: "仕上がり", en: "Trim" },
            cropBleed: { ja: "裁ち落とし", en: "Bleed" }
        },
        label: {
            notSelected: { ja: "未指定", en: "Not selected" },
            scale: { ja: "倍率", en: "Scale" },
            scaleUnit: { ja: "%", en: "%" },
            gap: { ja: "間隔", en: "Gap" },
            gapUnit: { ja: "pt", en: "pt" },
            columns: { ja: "列数", en: "Columns" },
            columnsAuto: { ja: "自動", en: "Auto" },
            rows: { ja: "行数", en: "Rows" },
            cropMode: { ja: "トリミング", en: "Crop to" },
            evenPage: { ja: "偶数ページ", en: "Even pages" },
            colorMode: { ja: "カラーモード", en: "Color mode" },
            errorDetails: { ja: "詳細", en: "Details" }
        },
        button: {
            selectFile: { ja: "ファイル指定", en: "Select File" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            source: { ja: "PDF または AI ファイルを選択（複数ページ対応）", en: "Choose a PDF or AI file (multi-page supported)" },
            customRange: { ja: "例: 1-10, 1,3,5", en: "e.g. 1-10, 1,3,5" },
            totalPages: { ja: "配置されるページ数 / 総ページ数", en: "Pages to place / total pages" },
            columns: { ja: "1 行あたりの列数。自動はカンバス右端で折り返し", en: "Columns per row; Auto wraps at the canvas edge" },
            estimate: {
                ja: "配置に必要な行数（列数が自動のとき、見開きを分割するときは不定）",
                en: "Rows needed for the layout (unknown when Columns is Auto or spreads are split)"
            },
            gapPerArtboard: { ja: "アートボードの間隔", en: "Artboard gap" },
            gapIgnoreArtboard: { ja: "配置するオブジェクトの間隔", en: "Placed object gap" },
            scale: { ja: "配置倍率（［アートボードを無視］のときのみ有効）", en: "Placement scale (only when ignoring artboards)" },
            perArtboard: { ja: "各ページをアートボードとして並べる（倍率 100%）", en: "Lay out each page as an artboard (100%)" },
            ignoreArtboard: { ja: "アートボードを追加せずオブジェクトとして配置", en: "Place as objects without adding artboards" },
            roundCorner: { ja: "外接矩形の角を丸める", en: "Round the corners of the bounding rectangle" },
            cropMode: {
                ja: "PDF のどのボックスを基準に配置するかを選びます。AI ファイルでは使いません",
                en: "Which PDF box the pages are placed from. Not used for AI files"
            },
            splitSpreads: {
                ja: "横長のページを見開きとみなし、左右2つに切り分けて並べます",
                en: "Treats landscape pages as spreads and lays them out as two halves"
            },
            evenPage: {
                ja: "見開きを分割したとき、偶数ページを置く側です。PDF の綴じ方向から自動で設定します",
                en: "Which side the even pages go to when a spread is split. Detected from the PDF binding direction"
            },
            destinationCurrent: { ja: "現在のドキュメントに配置します", en: "Places into the current document" },
            destinationNew: { ja: "新規ドキュメントを作り、そこに配置します", en: "Creates a new document and places into it" },
            colorMode: {
                ja: "作成する新規ドキュメントのカラーモードです。初期値は現在のドキュメントに合わせます",
                en: "Color mode of the new document. Defaults to that of the current document"
            }
        },
        alert: {
            needDoc: {
                ja: "ドキュメントを開いてから実行してください。",
                en: "Please open a document before running."
            },
            needFile: {
                ja: "先に［ファイル指定］で読み込みファイルを選択してください。",
                en: "Please select a source file first."
            },
            placeError: {
                ja: "配置中にエラーが発生しました。",
                en: "An error occurred while placing the pages."
            },
            linkUnknown: {
                ja: "画像のリンク先が不明でした。",
                en: "Image link not found."
            },
            pageCountFail: {
                ja: "リンクされたPDF/AIファイルのページ数を取得できませんでした。",
                en: "Could not determine the page count of the linked PDF/AI file."
            },
            pickPdfAi: {
                ja: "PDFまたはAIファイルを選択してください。",
                en: "Please select a PDF or AI file."
            },
            someSkipped: {
                ja: "一部のページを配置できなかったため、スキップしました。",
                en: "Some pages could not be placed and were skipped."
            }
        }
    };

    // =========================================
    // 単位 / Unit
    // =========================================

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /**
     * パネルへ共通のレイアウト設定を適用する
     * @param {object} panel - 対象のパネル
     * @param {number} spacing - 子要素の間隔。省略時は既定値
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
     * グループへ共通のレイアウト設定を適用する（row/column で整列を切り替え）
     * @param {object} group - 対象のグループ
     * @param {string} orientation - "row" または "column"
     * @param {number} spacing - 子要素の間隔。省略時は既定値
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 例外オブジェクトから表示用の詳細文字列を取り出す
     * @param {object} e - 例外オブジェクトまたは文字列
     * @returns {string} 表示用の詳細文字列
     */
    function SC_getErrorDetailText(e) {
        if (e === undefined || e === null) return "";
        try {
            if (typeof e === "string") return e;
            if (e && e.message) return String(e.message);
            return String(e);
        } catch (err) {
            return "";
        }
    }

    /**
     * ラベルエントリと任意のエラーをアラート表示する
     * @param {object} entry - ja / en を持つラベル定義
     * @param {object} e - 併記する例外オブジェクト。省略可
     * @returns {void}
     */
    function SC_alert(entry, e) {
        try {
            var msg = getLabel(entry);
            var detail = SC_getErrorDetailText(e);
            if (detail) msg += "\n\n" + labelText(LABELS.label.errorDetails) + "\n" + detail;
            alert(msg);
        } catch (err) { }
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ========================
    // 配置処理ヘルパー
    // - 配置前のページ番号指定
    // - 配置サイズの計測
    // - レイアウト計算
    // - アートボードの作成 / 更新
    // - 単ページの配置
    // - 指定範囲が見える表示倍率への調整
    // ========================

    /**
     * PDF 取り込みページ番号を環境設定にセットする
     * @param {number} pageNum - 取り込むページ番号
     * @returns {void}
     */
    function placementSetImportPageNumber(pageNum) {
        var n = parseInt(pageNum, 10);
        if (isNaN(n) || n < 1) n = 1;
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", n);
        } catch (e) { }
    }

    /**
     * 取り込みページ番号を 1 に戻す
     * @returns {void}
     */
    function placementResetImportPageNumber() {
        try {
            app.preferences.setIntegerPreference("plugin/PDFImport/PageNumber", 1);
        } catch (e) { }
    }

    /**
     * 指定ページを一時配置してサイズを計測する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObj - 読み込み元ファイル
     * @param {number} pageNum - 計測するページ番号
     * @param {number} cropMode - トリミング指定（CropTo の値）
     * @returns {object} width と height を持つオブジェクト
     */
    function placementMeasurePlacedPageSize(targetDoc, fileObj, pageNum, cropMode) {
        if (isPdfLikeFile(fileObj)) {
            SC_setPdfCropPreference(cropMode);
        }
        placementSetImportPageNumber(pageNum);

        var measureItem = null;
        try {
            measureItem = targetDoc.placedItems.add();
            measureItem.file = fileObj;
            return {
                width: measureItem.width,
                height: measureItem.height
            };
        } finally {
            if (measureItem) {
                try { measureItem.remove(); } catch (e) { }
            }
        }
    }

    /**
     * 各ページを一度ずつ一時配置して、素のサイズを計測する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObj - 読み込み元ファイル
     * @param {Array} targetPages - 対象ページ番号の配列
     * @param {number} cropMode - トリミング指定（CropTo の値）
     * @returns {Array} 計測結果の配列。計測できなかったページは null
     */
    function placementMeasurePages(targetDoc, fileObj, targetPages, cropMode) {
        var measured = [];
        for (var i = 0; i < targetPages.length; i++) {
            var pageNum = parseInt(targetPages[i], 10);
            if (isNaN(pageNum) || pageNum < 1) pageNum = 1;
            try {
                var size = placementMeasurePlacedPageSize(targetDoc, fileObj, pageNum, cropMode);
                measured.push({ page: pageNum, width: size.width, height: size.height });
            } catch (e) {
                measured.push(null); // 計測できないページ / Page that cannot be measured
            }
        }
        return measured;
    }

    /**
     * 見開き（横長）のページを左右の半ページ2つに分ける（DOM は参照しない）
     * 偶数ページが右なら右半分、左なら左半分を先に並べる
     * @param {Array} measured - placementMeasurePages() の戻り値
     * @param {boolean} evenPageOnRight - 偶数ページを右に置くなら true
     * @returns {Array} 半ページには half（"left" / "right"）が付いた計測結果の配列
     */
    function placementSplitSpreads(measured, evenPageOnRight) {
        var pieces = [];
        for (var i = 0; i < measured.length; i++) {
            var m = measured[i];
            if (!m || m.width <= m.height * SPREAD_ASPECT_RATIO) {
                pieces.push(m);
                continue;
            }
            var halfWidth = m.width / 2;
            var firstHalf = evenPageOnRight ? "right" : "left";
            var secondHalf = evenPageOnRight ? "left" : "right";
            pieces.push({ page: m.page, half: firstHalf, width: halfWidth, height: m.height });
            pieces.push({ page: m.page, half: secondHalf, width: halfWidth, height: m.height });
        }
        return pieces;
    }

    /**
     * 計測結果から各ページの配置位置と全体サイズを求める（DOM は参照しない）
     * @param {Array} measured - placementMeasurePages() の戻り値
     * @param {number} scaleFactor - 配置倍率（1 = 100%）
     * @param {number} gap - ページ間の間隔（pt）
     * @param {number} colsPerRow - 1 行あたりの列数。0 は自動
     * @returns {object} slots（配置位置の配列。半ページは half 付き）と width / height を持つオブジェクト
     */
    function placementBuildLayout(measured, scaleFactor, gap, colsPerRow) {
        var slots = [];
        var nextX = 0;
        var topY = 0;
        var colCount = 0;
        var rowMaxH = 0; // 現在の行で最も高いページの高さ / Tallest page height in the current row
        var maxRight = 0;
        var minBottom = 0;

        for (var i = 0; i < measured.length; i++) {
            var m = measured[i];
            if (!m) {
                slots.push(null);
                continue;
            }

            var pageW = m.width * scaleFactor;
            var pageH = m.height * scaleFactor;

            // 指定列数に達した、または行幅が上限を超えるとき次の行へ折り返す
            // Wrap to the next row when the column count or the row width limit is reached
            var wrapByCol = colsPerRow > 0 && colCount >= colsPerRow;
            var wrapByEdge = colsPerRow === 0 && nextX > 0 && nextX + pageW > MAX_AUTO_ROW_WIDTH;
            if (wrapByCol || wrapByEdge) {
                topY -= rowMaxH + gap;
                nextX = 0;
                rowMaxH = 0;
                colCount = 0;
            }

            slots.push({ page: m.page, half: m.half || null, x: nextX, y: topY, width: pageW, height: pageH });

            if (nextX + pageW > maxRight) maxRight = nextX + pageW;
            if (topY - pageH < minBottom) minBottom = topY - pageH;

            nextX += pageW + gap;
            colCount++;
            if (pageH > rowMaxH) rowMaxH = pageH;
        }

        return { slots: slots, width: maxRight, height: -minBottom };
    }

    /**
     * 1 枚目はアクティブアートボードを再利用し、以降は新規追加する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {number} activeIdx - アクティブアートボードのインデックス
     * @param {Array} abRect - アートボードの矩形 [left, top, right, bottom]
     * @param {number} abCount - これまでに用意したアートボード数
     * @returns {number} 更新後のアートボード数
     */
    function placementUseOrAddArtboard(targetDoc, activeIdx, abRect, abCount) {
        if (abCount === 0) {
            targetDoc.artboards[activeIdx].artboardRect = abRect;
        } else {
            targetDoc.artboards.add(abRect);
        }
        return abCount + 1;
    }

    /**
     * ページを配置する（必要なら倍率を適用）
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObj - 読み込み元ファイル
     * @param {number} pageNum - 配置するページ番号
     * @param {Array} pos - 配置位置 [left, top]
     * @param {number} cropMode - トリミング指定（CropTo の値）
     * @param {number} scalePct - 配置倍率（%）
     * @returns {PlacedItem} 配置したアイテム
     */
    function placementPlacePage(targetDoc, fileObj, pageNum, pos, cropMode, scalePct) {
        if (isPdfLikeFile(fileObj)) {
            SC_setPdfCropPreference(cropMode);
        }
        placementSetImportPageNumber(pageNum);

        var item = targetDoc.placedItems.add();
        item.file = fileObj;

        if (typeof scalePct === "number" && scalePct !== 100) {
            item.resize(scalePct, scalePct);
        }
        // resize() の基準点は position と無関係なため、スケール後に左上位置を確定させる
        // resize() uses its own anchor, so set the top-left position after scaling
        item.position = pos;
        return item;
    }

    /**
     * 矩形群 [left, top, right, bottom] の和を求める
     * @param {Array} rects - 矩形の配列
     * @returns {Array} 外接矩形。求められない場合は null
     */
    function placementUnionRects(rects) {
        var union = null;
        for (var i = 0; i < rects.length; i++) {
            var r = rects[i];
            if (!r || r.length !== 4) continue;
            if (!union) {
                union = [r[0], r[1], r[2], r[3]];
                continue;
            }
            if (r[0] < union[0]) union[0] = r[0];
            if (r[1] > union[1]) union[1] = r[1];
            if (r[2] > union[2]) union[2] = r[2];
            if (r[3] < union[3]) union[3] = r[3];
        }
        return union;
    }

    /**
     * 全アートボードの外接範囲を求める
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {Array} 外接矩形。求められない場合は null
     */
    function placementGetArtboardsBounds(targetDoc) {
        if (!targetDoc || !targetDoc.artboards) return null;
        var rects = [];
        for (var i = 0; i < targetDoc.artboards.length; i++) {
            rects.push(targetDoc.artboards[i].artboardRect);
        }
        return placementUnionRects(rects);
    }

    /**
     * 複数アイテムの可視境界の和を求める
     * @param {Array} items - 対象アイテムの配列
     * @returns {Array} 外接矩形。求められない場合は null
     */
    function placementGetItemsVisibleBounds(items) {
        if (!items) return null;
        var rects = [];
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) continue;
            try {
                rects.push(it.visibleBounds);
            } catch (e) {
                try { rects.push(it.geometricBounds); } catch (err) { }
            }
        }
        return placementUnionRects(rects);
    }

    /**
     * 指定した範囲全体が見えるよう表示倍率と表示位置を調整する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {Array} bounds - 表示したい矩形 [left, top, right, bottom]
     * @returns {void}
     */
    function placementFitBoundsInView(targetDoc, bounds) {
        if (!targetDoc || !bounds) return;

        try {
            app.activeDocument = targetDoc;
        } catch (e) { }

        try {
            var view = targetDoc.activeView;
            if (!view) return;

            var unionWidth = bounds[2] - bounds[0];
            var unionHeight = bounds[1] - bounds[3];
            if (unionWidth <= 0 || unionHeight <= 0) return;

            var currentBounds = view.bounds;
            var currentWidth = currentBounds[2] - currentBounds[0];
            var currentHeight = currentBounds[1] - currentBounds[3];
            var currentZoom = view.zoom;
            if (!(currentWidth > 0) || !(currentHeight > 0) || !(currentZoom > 0)) return;

            var paddingScale = 0.9;
            var zoomX = currentZoom * (currentWidth / unionWidth);
            var zoomY = currentZoom * (currentHeight / unionHeight);

            view.centerPoint = [
                (bounds[0] + bounds[2]) / 2,
                (bounds[1] + bounds[3]) / 2
            ];
            view.zoom = Math.min(zoomX, zoomY) * paddingScale;
        } catch (e) { }
    }

    /**
     * 最大キャンバス範囲を求める（OMOTI 氏のアイデア）
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {Array} キャンバスの矩形 [left, top, right, bottom]
     */
    function getLargestCanvasBounds(targetDoc) {
        var LARGEST_SIZE = 16383;
        try {
            var tempLayer = targetDoc.layers.add();
            var tempText = tempLayer.textFrames.add();
            var left = tempText.matrix.mValueTX;
            var top = tempText.matrix.mValueTY;
            tempLayer.remove();
            return [left, top, left + LARGEST_SIZE, top - LARGEST_SIZE];
        } catch (e) {
            var half = LARGEST_SIZE / 2;
            return [-half, half, half, -half];
        }
    }

    /**
     * 配置先の新規ドキュメントを作成してアクティブにする。サイズは最初に並べるページに合わせる
     * @param {DocumentColorSpace} colorSpace - カラーモード
     * @param {object} layout - placementBuildLayout() の戻り値
     * @param {Document} fallbackDoc - 並べるページが無いときにサイズを借りるドキュメント
     * @returns {Document} 作成したドキュメント
     */
    function createOutputDocument(colorSpace, layout, fallbackDoc) {
        var firstSlot = null;
        for (var i = 0; i < layout.slots.length && !firstSlot; i++) {
            firstSlot = layout.slots[i];
        }
        var docWidth = firstSlot ? firstSlot.width : fallbackDoc.width;
        var docHeight = firstSlot ? firstSlot.height : fallbackDoc.height;

        var outputDoc = app.documents.add(colorSpace, docWidth, docHeight);

        /* ラスタライズ効果解像度の設定に失敗しても配置は続ける / Keep placing even if the raster resolution cannot be set */
        try {
            var rasterSettings = outputDoc.rasterEffectSettings;
            rasterSettings.resolution = NEW_DOCUMENT_RASTER_RESOLUTION;
            outputDoc.rasterEffectSettings = rasterSettings;
        } catch (e) { }

        app.activeDocument = outputDoc;
        return outputDoc;
    }

    // =========================================
    // ケイ処理ヘルパー
    // =========================================

    /**
     * 配置アイテムの外接矩形（または指定した矩形）でクリッピングマスクグループを作成する
     * @param {PlacedItem} placedItem - 対象の配置アイテム
     * @param {number[]} [clipRect] - マスクの矩形 [left, top, right, bottom]。省略時は配置アイテムの外接矩形
     * @returns {GroupItem} 作成したクリッピンググループ
     */
    function keiCreateClippingMaskGroup(placedItem, clipRect) {
        var targetLayer = placedItem.layer;
        var rect = clipRect
            ? targetLayer.pathItems.rectangle(clipRect[1], clipRect[0], clipRect[2] - clipRect[0], clipRect[1] - clipRect[3])
            : targetLayer.pathItems.rectangle(placedItem.top, placedItem.left, placedItem.width, placedItem.height);
        rect.stroked = false;
        rect.filled = false;

        var groupItem = targetLayer.groupItems.add();
        placedItem.moveToBeginning(groupItem);
        rect.moveToBeginning(groupItem);
        groupItem.clipped = true;

        return groupItem;
    }

    /**
     * 角丸 LiveEffect を適用する
     * @param {PageItem} targetItem - 対象アイテム
     * @param {number} radius - 角丸の半径（pt）
     * @returns {void}
     */
    function keiApplyRoundCornersLiveEffect(targetItem, radius) {
        if (!targetItem) return;
        var r = Number(radius);
        if (isNaN(r) || r <= 0) return;
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + r + ' "/></LiveEffect>';
        try {
            targetItem.applyEffect(xml);
        } catch (e) { }
    }

    /**
     * 配置アイテムにケイ処理を適用する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {PageItem} placedItem - 対象の配置アイテム。見開きの半ページは切り抜き済みのクリップグループ
     * @param {object} keiOpts - mode / roundCorners / roundRadius を持つ設定
     * @returns {PageItem} 処理後のアイテム。処理しない場合は元のアイテム
     */
    function keiApplyToPlacedItem(targetDoc, placedItem, keiOpts) {
        if (!keiOpts || !placedItem) return placedItem;

        if (keiOpts.mode === 'clipGroup') {
            /* 半ページはすでにクリップグループなので包み直さない / A half page is already a clip group */
            var isClipGroup = placedItem.typename === "GroupItem" && placedItem.clipped;
            var group = isClipGroup ? placedItem : keiCreateClippingMaskGroup(placedItem);

            if (keiOpts.roundCorners && keiOpts.roundRadius > 0) {
                keiApplyRoundCornersLiveEffect(group, keiOpts.roundRadius);
            }

            targetDoc.selection = [group];
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            app.executeMenuCommand('Live Pathfinder Exclude');
            targetDoc.selection = null;
            return group;
        }

        return placedItem;
    }

    // =========================================
    // 配置時のトリミング設定
    // 効くのは plugin/PDFImport/CropTo。値は 0=アート / 1=トリミング（CropBox）/ 2=仕上がり（TrimBox）/ 3=裁ち落とし / 4=メディア（実測）
    // =========================================

    // トリミングの初期選択。ドロップダウンの番号が CropTo の値と一致する（1 = トリミング / CropBox）。AI ファイルは常にこの値
    // Default crop; the dropdown index equals the CropTo value (1 = Crop / CropBox). AI files always use it
    var DEFAULT_CROP_MODE = 1;

    // PDF 配置時の crop プリファレンスキー（環境差を吸収するため複数試行）/ Crop preference keys (multiple keys tried for version differences)
    var PDF_CROP_PREFERENCE_KEYS = [
        "plugin/PDFImport/CropToBox",
        "plugin/PDFImport/CropTo",
        "plugin/PDFImport/CropBox",
        "plugin/PDFImport/CropToType"
    ];

    /**
     * crop 種別を複数キーに設定する（環境差を吸収）
     * @param {number} cropVal - 設定する crop 種別
     * @returns {void}
     */
    function SC_setPdfCropPreference(cropVal) {
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            try {
                app.preferences.setIntegerPreference(PDF_CROP_PREFERENCE_KEYS[i], cropVal);
            } catch (e) { }
        }
    }

    /**
     * 現在の crop プリファレンスを退避する
     * @returns {object} キーと値を持つスナップショット
     */
    function SC_snapshotPdfCropPreference() {
        var snapshot = {};
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            try {
                snapshot[PDF_CROP_PREFERENCE_KEYS[i]] = app.preferences.getIntegerPreference(PDF_CROP_PREFERENCE_KEYS[i]);
            } catch (e) { }
        }
        return snapshot;
    }

    /**
     * 退避した crop プリファレンスを復元する
     * @param {object} snapshot - SC_snapshotPdfCropPreference() の戻り値
     * @returns {void}
     */
    function SC_restorePdfCropPreference(snapshot) {
        if (!snapshot) return;
        for (var i = 0; i < PDF_CROP_PREFERENCE_KEYS.length; i++) {
            var prefKey = PDF_CROP_PREFERENCE_KEYS[i];
            if (snapshot[prefKey] === undefined) continue;
            try {
                app.preferences.setIntegerPreference(prefKey, snapshot[prefKey]);
            } catch (e) { }
        }
    }

    // ============================================================
    // ページ数取得ヘルパー
    // - リンクされた PDF/AI の総ページ数（最終ページ番号）を推定
    // - ファイルが明示指定された場合は、一時配置して既存の選択ベースの
    //   ページ数取得ロジックを再利用し、取得後すぐに削除
    // - 現在の選択に対してリンク変更や内容変更は行わない
    // ============================================================

    /**
     * 拡張子から PDF/AI ファイルかどうかを判定する
     * @param {File} f - 対象ファイル
     * @returns {boolean} PDF または AI なら true
     */
    function isPdfLikeFile(f) {
        return /\.(?:pdf|ai)$/i.test(String((f && f.name) || ""));
    }

    /**
     * 拡張子から PDF ファイルかどうかを判定する（トリミングは PDF でだけ選べる）
     * @param {File} f - 対象ファイル
     * @returns {boolean} PDF なら true
     */
    function isPdfFile(f) {
        return /\.pdf$/i.test(String((f && f.name) || ""));
    }

    /**
     * 綴じ方向を判定する。先頭の BINDING_SCAN_LINE_LIMIT 行に /Direction /R2L があれば右綴じとみなす
     * @param {File} f - 対象ファイル
     * @returns {boolean} 右綴じ（偶数ページが右）なら true。読めないときは false
     */
    function isRightBoundFile(f) {
        var isRightBound = false;
        try {
            f.open('r');
            for (var lineIndex = 0; lineIndex < BINDING_SCAN_LINE_LIMIT && !f.eof; lineIndex++) {
                if (/\/Direction\s*\/R2L/.test(f.readln())) {
                    isRightBound = true;
                    break;
                }
            }
        } catch (e) {
            isRightBound = false;
        } finally {
            try { f.close(); } catch (err) { }
        }
        return isRightBound;
    }

    /**
     * 選択範囲から最初の PlacedItem を再帰的に探す
     * @param {Array} items - 探索対象のアイテム配列
     * @returns {PlacedItem} 見つかった配置アイテム。無ければ null
     */
    function pageCountFindFirstPlacedItem(items) {
        if (!items || items.length <= 0) return null;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            var typeName = (item.constructor && item.constructor.name) ? item.constructor.name : '';
            if (typeName === 'PlacedItem') return item;
            if (typeName === 'GroupItem') {
                var hit = pageCountFindFirstPlacedItem(item.pageItems);
                if (hit) return hit;
            }
        }
        return null;
    }

    /**
     * PDF/AI ファイルから総ページ数を推定する
     * @param {File} file - 対象ファイル
     * @returns {number} 推定した総ページ数。取得できない場合は undefined
     */
    function getPageLengthFromFile(file) {
        var countInDictPattern = /<<\/Count\s(\d+)/;
        var pageRefPattern1 = /<<\/Type\/Page\/Parent/;
        var pageRefPattern2 = /\/Type\s\/Page\s/;
        var pageRefPattern3 = /\/StructParents\s\d+.*\/Type\/Page>>/;
        var linearizedCountPattern = /<<\/Linearized\s.+\/N\s(\d+)\/T\s.+>>/;
        var pagesNodePattern = /\/Type\/Pages/;
        var countPattern = /\/Count\s(\d+)/;

        var pageCount;
        var line;
        var pageRefCount = 0;
        var declaredCount = 0;
        var hasDeclaredCount = false;

        try {
            file.open('r');
            while (!file.eof) {
                line = file.readln();
                if (countInDictPattern.test(line) || linearizedCountPattern.test(line)) {
                    declaredCount = Number(RegExp.$1);
                    if (declaredCount > 0) {
                        pageCount = declaredCount;
                        hasDeclaredCount = true;
                        break;
                    }
                }
                if (pageRefPattern1.test(line) || pageRefPattern2.test(line) || pageRefPattern3.test(line)) pageRefCount++;
                if (pagesNodePattern.test(line)) {
                    line = file.readln();
                    if (countPattern.test(line)) {
                        declaredCount = Number(RegExp.$1);
                        if (pageRefCount < declaredCount) pageRefCount = declaredCount;
                    }
                }
            }
            // 宣言値を取得できたときは、数え上げた概算で上書きしない
            // Keep the declared count; do not overwrite it with the scanned estimate
            if (!hasDeclaredCount && pageRefCount > 0) pageCount = pageRefCount;
        } catch (e) {
            SC_alert(LABELS.alert.pageCountFail, e);
        } finally {
            try { file.close(); } catch (err) { }
        }
        return pageCount;
    }

    /**
     * 選択中の配置画像のリンク元から総ページ数を取得する
     * @param {Array} selectionItems - 対象アイテムの配列
     * @returns {number} 総ページ数。取得できない場合は null
     */
    function getLastPageFromSelection(selectionItems) {
        var placed = pageCountFindFirstPlacedItem(selectionItems);
        if (!placed) return null;

        var f = placed.file;
        if (!f) {
            SC_alert(LABELS.alert.linkUnknown);
            return null;
        }

        if (!isPdfLikeFile(f)) return null;

        var last = getPageLengthFromFile(f);
        if (!last || isNaN(Number(last)) || Number(last) <= 0) {
            SC_alert(LABELS.alert.pageCountFail);
            return null;
        }

        return Number(last);
    }

    /**
     * 指定ファイル（または現在の選択）から総ページ数を取得する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {File} fileObjOrNull - 読み込み元ファイル。null なら選択から取得
     * @param {function} setPathTextFn - ファイルパス表示を更新するコールバック
     * @returns {number} 総ページ数。取得できない場合は null
     */
    function updatePageCountFromPlacedOrFile(targetDoc, fileObjOrNull, setPathTextFn) {
        var placedTemp = null;

        try {
            // ファイル指定時は一時配置して、既存の選択ベースのページ数取得ロジックを再利用
            if (fileObjOrNull) {
                if (!isPdfLikeFile(fileObjOrNull)) {
                    SC_alert(LABELS.alert.pickPdfAi);
                    return null;
                }

                try {
                    placedTemp = targetDoc.placedItems.add();
                    placedTemp.file = fileObjOrNull;

                    // 一時配置物は、いったん表示中ビューの左上付近へ置く
                    try {
                        var vb = targetDoc.activeView && targetDoc.activeView.bounds ? targetDoc.activeView.bounds : null;
                        if (vb && vb.length === 4) {
                            placedTemp.position = [vb[0], vb[1]];
                        }
                    } catch (e) { }

                    // 既存ロジックをそのまま使って総ページ数を取得
                    var lastFromFile = getLastPageFromSelection([placedTemp]);
                    if (setPathTextFn) setPathTextFn(fileObjOrNull);
                    return lastFromFile;
                } finally {
                    if (placedTemp) {
                        try { placedTemp.remove(); } catch (e) { }
                    }
                }
            }

            // ファイル未指定時は現在の選択から取得（PlacedItem が選択されている前提）
            var last = getLastPageFromSelection(targetDoc.selection);

            // 選択中の配置画像があれば、そのファイルパス表示も更新
            try {
                var placedSel = pageCountFindFirstPlacedItem(targetDoc.selection);
                if (placedSel && placedSel.file && setPathTextFn) setPathTextFn(placedSel.file);
            } catch (e) { }

            return last;
        } catch (e) {
            SC_alert(LABELS.alert.pageCountFail, e);
            return null;
        }
    }

    /**
     * 入力文字列（例: "1-20", "1,3,5"）をページ番号配列へ変換する
     * 重複と範囲外（1..maxPages 以外）は除外する
     * @param {string} inputStr - ページ範囲の文字列
     * @param {number} maxPages - 上限ページ数。0 以下なら上限なし
     * @returns {Array} ページ番号の配列
     */
    function parsePageNumbers(inputStr, maxPages) {
        var result = [];
        var seen = {};
        var hasMax = (typeof maxPages === "number" && maxPages > 0);

        /**
         * 1 件のページ番号を検証して追加する
         * @param {number} num - ページ番号
         * @returns {void}
         */
        function addPage(num) {
            if (isNaN(num) || num < 1) return;
            if (hasMax && num > maxPages) return;
            var key = "p" + num;
            if (seen[key]) return;
            seen[key] = true;
            result.push(num);
        }

        var parts = String(inputStr).split(',');
        for (var i = 0; i < parts.length; i++) {
            var part = parts[i].replace(/^\s+|\s+$/g, '');
            if (part === '') continue;

            if (part.indexOf('-') > -1) {
                var bounds = part.split('-');
                var start = parseInt(bounds[0], 10);
                var end = parseInt(bounds[1], 10);
                if (!isNaN(start) && !isNaN(end)) {
                    var min = Math.min(start, end);
                    var max = Math.max(start, end);
                    for (var j = min; j <= max; j++) {
                        addPage(j);
                    }
                }
            } else {
                addPage(parseInt(part, 10));
            }
        }
        return result;
    }

    /**
     * メイン処理。ダイアログを構築して配置を実行する
     * @returns {void}
     */
    function main() {

        if (app.documents.length === 0) {
            SC_alert(LABELS.alert.needDoc);
            return;
        }

        var doc = app.activeDocument;
        // 現在の読み込み対象ファイル。未選択時は null。
        var sourceFile = null;
        var detectedRangeText = "";
        var customRangeText = "";
        // 直前に［指定ページ］が選択されていたか / Whether custom range was the previous mode
        var wasCustomRange = false;
        // OK 押下時の配置指示。ダイアログを閉じてから実行する / Placement request, run after the dialog closes
        var placementRequest = null;
        // 進捗表示用のパレット / Progress palette
        var progressWin = null;
        var progressBar = null;

        // ------------------------
        // UI構築
        // ------------------------
        var win = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        win.alignChildren = "fill";

        var bodyGroup = win.add("group");
        bodyGroup.orientation = "row";
        bodyGroup.alignChildren = ["fill", "top"];

        var leftColumnGroup = bodyGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "fill";

        var rightColumnGroup = bodyGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "fill";

        var sourcePanel = leftColumnGroup.add('panel', undefined, getLabel(LABELS.panel.source));
        setupPanel(sourcePanel);
        var btnBrowse = sourcePanel.add('button', undefined, getLabel(LABELS.button.selectFile));
        // パネルの alignChildren は fill なので、ボタンだけ自前の幅で左寄せにする
        // The panel fills its children, so keep the button at its natural width
        btnBrowse.alignment = ["left", "center"];
        btnBrowse.helpTip = getLabel(LABELS.tooltip.source);
        var stSourceName = sourcePanel.add('statictext', undefined, getLabel(LABELS.label.notSelected));
        stSourceName.characters = 16;

        var pagesPanel = leftColumnGroup.add('panel', undefined, getLabel(LABELS.panel.pages));
        setupPanel(pagesPanel);
        var rangeModeGroup = pagesPanel.add('group');
        setupGroup(rangeModeGroup, 'column');
        var rbRangeAll = rangeModeGroup.add('radiobutton', undefined, getLabel(LABELS.radio.rangeAll));
        var rbRangeFirst = rangeModeGroup.add('radiobutton', undefined, getLabel(LABELS.radio.rangeFirst));
        var rbRangeCustom = rangeModeGroup.add('radiobutton', undefined, labelText(LABELS.radio.rangeCustom));
        // 初期状態は「全ページ」を選択
        rbRangeAll.value = true;

        var rangeInputGroup = pagesPanel.add('group');
        setupGroup(rangeInputGroup, 'row');
        var etRange = rangeInputGroup.add('edittext', undefined, '');
        etRange.characters = 10;
        // 初期選択は「全ページ」なので無効から始める / Starts disabled: the default mode is All Pages
        etRange.enabled = false;
        etRange.helpTip = getLabel(LABELS.tooltip.customRange);

        var totalPagesGroup = pagesPanel.add('group');
        totalPagesGroup.orientation = 'row';
        totalPagesGroup.alignChildren = ['right', 'center'];
        totalPagesGroup.alignment = ['fill', 'top'];
        var stTotalPages = totalPagesGroup.add('statictext', undefined, '');
        stTotalPages.characters = 8;
        stTotalPages.justify = 'right';
        stTotalPages.alignment = ['right', 'center'];
        stTotalPages.helpTip = getLabel(LABELS.tooltip.totalPages);

        var methodPanel = leftColumnGroup.add("panel", undefined, getLabel(LABELS.panel.mode));
        setupPanel(methodPanel);
        var methodGroup = methodPanel.add("group");
        setupGroup(methodGroup, "column");
        var rbPerArtboard = methodGroup.add("radiobutton", undefined, getLabel(LABELS.radio.perArtboard));
        rbPerArtboard.helpTip = getLabel(LABELS.tooltip.perArtboard);
        var rbIgnoreArtboard = methodGroup.add("radiobutton", undefined, getLabel(LABELS.radio.ignoreArtboard));
        rbIgnoreArtboard.helpTip = getLabel(LABELS.tooltip.ignoreArtboard);
        rbPerArtboard.value = true;

        var destinationPanel = leftColumnGroup.add("panel", undefined, getLabel(LABELS.panel.destination));
        setupPanel(destinationPanel);
        var destinationGroup = destinationPanel.add("group");
        setupGroup(destinationGroup, "column");
        var rbDestinationCurrent = destinationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.destinationCurrent));
        rbDestinationCurrent.helpTip = getLabel(LABELS.tooltip.destinationCurrent);
        var rbDestinationNew = destinationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.destinationNew));
        rbDestinationNew.helpTip = getLabel(LABELS.tooltip.destinationNew);
        rbDestinationCurrent.value = true;

        var colorModeGroup = destinationPanel.add("group");
        setupGroup(colorModeGroup, "row");
        var stColorModeLabel = colorModeGroup.add("statictext", undefined, labelText(LABELS.label.colorMode));
        var rbColorCMYK = colorModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.colorModeCMYK));
        var rbColorRGB = colorModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.colorModeRGB));
        stColorModeLabel.helpTip = rbColorCMYK.helpTip = rbColorRGB.helpTip = getLabel(LABELS.tooltip.colorMode);
        /* 現在のドキュメントのカラーモードを初期値にする / Default to the current document's color mode */
        if (doc.documentColorSpace === DocumentColorSpace.RGB) rbColorRGB.value = true;
        else rbColorCMYK.value = true;

        var layoutPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.placement));
        setupPanel(layoutPanel);

        var columnsGroup = layoutPanel.add("group");
        setupGroup(columnsGroup, "row");
        var stColsLabel = columnsGroup.add("statictext", undefined, labelText(LABELS.label.columns));
        stColsLabel.justify = "right";
        var colsFieldGroup = columnsGroup.add("group"); /* ∧∨と入力欄は隙間0で突き合わせる / stepper butts the field */
        colsFieldGroup.orientation = "row";
        colsFieldGroup.alignChildren = ["left", "center"];
        colsFieldGroup.spacing = 0;
        colsFieldGroup.margins = 0;
        /* 「自動」で↑を押すと1、1から↓で0になり onChange で「自動」へ戻る / Up from Auto gives 1; Down from 1 returns to Auto via onChange */
        var etCols;
        var colsStepper = addStepper(colsFieldGroup, function () { return etCols; }, {
            integer: true, min: 0, onStep: notifyFieldChange
        });
        etCols = colsFieldGroup.add("edittext", undefined, getLabel(LABELS.label.columnsAuto));
        etCols.characters = 5;
        bindSteppedArrowKeys(etCols, colsStepper);
        etCols.helpTip = getLabel(LABELS.tooltip.columns);
        stColsLabel.helpTip = getLabel(LABELS.tooltip.columns);

        var rowsGroup = layoutPanel.add("group");
        setupGroup(rowsGroup, "row");
        var stRowsLabel = rowsGroup.add("statictext", undefined, labelText(LABELS.label.rows));
        stRowsLabel.justify = "right";
        var stRowsValue = rowsGroup.add("statictext", undefined, "");
        stRowsValue.characters = 5;
        stRowsLabel.helpTip = getLabel(LABELS.tooltip.estimate);
        stRowsValue.helpTip = getLabel(LABELS.tooltip.estimate);

        var gapGroup = layoutPanel.add("group");
        setupGroup(gapGroup, "row");
        var stGapLabel = gapGroup.add("statictext", undefined, labelText(LABELS.label.gap));
        stGapLabel.justify = "right";
        var gapFieldGroup = gapGroup.add("group"); /* ∧∨と入力欄は隙間0で突き合わせる / stepper butts the field */
        gapFieldGroup.orientation = "row";
        gapFieldGroup.alignChildren = ["left", "center"];
        gapFieldGroup.spacing = 0;
        gapFieldGroup.margins = 0;
        var etArtboardGap;
        var gapStepper = addStepper(gapFieldGroup, function () { return etArtboardGap; }, {
            min: 0, onStep: notifyFieldChange
        });
        etArtboardGap = gapFieldGroup.add("edittext", undefined, String(DEFAULT_ARTBOARD_GAP));
        etArtboardGap.characters = 5;
        bindSteppedArrowKeys(etArtboardGap, gapStepper);
        var stGapUnit = gapGroup.add("statictext", undefined, getLabel(LABELS.label.gapUnit));

        // 倍率は［アートボードを無視］のときだけ効く条件付き項目なので、レイアウトから分ける
        // Scale only applies when ignoring artboards, so keep it out of the Layout panel
        var optionPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.option));
        setupPanel(optionPanel);

        var scaleGroup = optionPanel.add("group");
        setupGroup(scaleGroup, "row");
        var stScaleLabel = scaleGroup.add("statictext", undefined, labelText(LABELS.label.scale));
        stScaleLabel.justify = "right";
        var scaleFieldGroup = scaleGroup.add("group"); /* ∧∨と入力欄は隙間0で突き合わせる / stepper butts the field */
        scaleFieldGroup.orientation = "row";
        scaleFieldGroup.alignChildren = ["left", "center"];
        scaleFieldGroup.spacing = 0;
        scaleFieldGroup.margins = 0;
        var etScale;
        var scaleStepper = addStepper(scaleFieldGroup, function () { return etScale; }, {
            min: 1, onStep: notifyFieldChange
        });
        etScale = scaleFieldGroup.add("edittext", undefined, "100");
        etScale.characters = 5;
        bindSteppedArrowKeys(etScale, scaleStepper);
        etScale.helpTip = getLabel(LABELS.tooltip.scale);
        stScaleLabel.helpTip = getLabel(LABELS.tooltip.scale);
        var stScaleUnit = scaleGroup.add("statictext", undefined, getLabel(LABELS.label.scaleUnit));

        var cropGroup = optionPanel.add("group");
        setupGroup(cropGroup, "row");
        var stCropLabel = cropGroup.add("statictext", undefined, labelText(LABELS.label.cropMode));
        stCropLabel.justify = "right";
        /* 並びは CropTo の値の順（0 アート / 1 トリミング / 2 仕上がり / 3 裁ち落とし） / Ordered by CropTo value */
        var ddCropMode = cropGroup.add("dropdownlist", undefined, [
            getLabel(LABELS.dropdown.cropArt),
            getLabel(LABELS.dropdown.cropCrop),
            getLabel(LABELS.dropdown.cropTrim),
            getLabel(LABELS.dropdown.cropBleed)
        ]);
        ddCropMode.selection = DEFAULT_CROP_MODE;
        stCropLabel.helpTip = ddCropMode.helpTip = getLabel(LABELS.tooltip.cropMode);

        var spreadPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.spread));
        setupPanel(spreadPanel);
        var cbSplitSpreads = spreadPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.splitSpreads));
        cbSplitSpreads.helpTip = getLabel(LABELS.tooltip.splitSpreads);
        var evenPageGroup = spreadPanel.add("group");
        setupGroup(evenPageGroup, "row");
        var stEvenPageLabel = evenPageGroup.add("statictext", undefined, labelText(LABELS.label.evenPage));
        var rbEvenPageRight = evenPageGroup.add("radiobutton", undefined, getLabel(LABELS.radio.evenPageRight));
        var rbEvenPageLeft = evenPageGroup.add("radiobutton", undefined, getLabel(LABELS.radio.evenPageLeft));
        rbEvenPageRight.value = true;
        stEvenPageLabel.helpTip = rbEvenPageRight.helpTip = rbEvenPageLeft.helpTip = getLabel(LABELS.tooltip.evenPage);

        var keiPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.kei));
        setupPanel(keiPanel);
        var keiModeGroup = keiPanel.add("group");
        setupGroup(keiModeGroup, "column");
        var rbKeiNone = keiModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.keiNone));
        var rbKeiClipGroup = keiModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.keiClipGroup));
        rbKeiNone.value = true;

        var roundCornerGroup = keiPanel.add("group");
        setupGroup(roundCornerGroup, "row");
        var cbRoundCorner = roundCornerGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.roundCorner));
        var etRoundCorner = roundCornerGroup.add("edittext", undefined, "3");
        etRoundCorner.characters = 5;
        var stRoundCornerUnit = roundCornerGroup.add("statictext", undefined, getUnitInfo().label);
        cbRoundCorner.value = false;
        etRoundCorner.enabled = false;
        cbRoundCorner.helpTip = getLabel(LABELS.tooltip.roundCorner);
        etRoundCorner.helpTip = getLabel(LABELS.tooltip.roundCorner);

        // === ボタンエリア（左スペーサー／右キャンセル・OK）/ Button area (spacer left, cancel+ok right)
        var buttonRow = addButtonRow(win);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOk = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        // ------------------------
        // UI状態の更新
        // ------------------------

        /**
         * 読み込みファイルの有無に応じて各パネルの有効／無効を切り替える
         * @returns {void}
         */
        function updatePanelEnabledState() {
            var enabled = !!sourceFile;
            pagesPanel.enabled = enabled;
            methodPanel.enabled = enabled;
            destinationPanel.enabled = enabled;
            layoutPanel.enabled = enabled;
            optionPanel.enabled = enabled;
            spreadPanel.enabled = enabled;
            keiPanel.enabled = enabled;
            updateCropEnabledState();
            /* ∧∨のディム表示を切り替える / update stepper dimming */
            redrawSteppersIn(layoutPanel);
            redrawSteppersIn(optionPanel);
        }

        /**
         * トリミングの有効／無効を切り替える（PDF のときだけ選べる）
         * @returns {void}
         */
        function updateCropEnabledState() {
            var enabled = isPdfFile(sourceFile);
            stCropLabel.enabled = enabled;
            ddCropMode.enabled = enabled;
        }

        /**
         * 偶数ページの位置の有効／無効を切り替える（見開きを分割するときだけ選べる）
         * @returns {void}
         */
        function updateEvenPageEnabledState() {
            evenPageGroup.enabled = cbSplitSpreads.value;
        }

        /**
         * カラーモードの有効／無効を切り替える（新規ドキュメントのときだけ選べる）
         * @returns {void}
         */
        function updateColorModeEnabledState() {
            colorModeGroup.enabled = rbDestinationNew.value;
        }

        /**
         * ページ範囲入力欄の有効／無効と表示内容を切り替える
         * @returns {void}
         */
        function updateRangeEnabledState() {
            if (rbRangeCustom.value) {
                etRange.enabled = true;
                etRange.text = customRangeText || detectedRangeText || '';
            } else {
                // ［指定ページ］から離れるときだけ入力内容を退避する
                // Keep the typed range only when leaving custom mode
                if (wasCustomRange) customRangeText = etRange.text;
                etRange.enabled = false;
                etRange.text = '';
            }
            wasCustomRange = rbRangeCustom.value;
        }

        /**
         * 検出済みの総ページ数を取得する
         * @returns {number} 総ページ数。未検出なら 0
         */
        function getDetectedTotalPages() {
            if (!detectedRangeText) return 0;
            var m = detectedRangeText.match(/^(?:1-)?(\d+)$/);
            return m ? (parseInt(m[1], 10) || 0) : 0;
        }

        /**
         * 現在の設定で配置されるページ数を求める
         * @param {number} totalPages - 総ページ数
         * @returns {number} 配置されるページ数
         */
        function getPlacedPageCount(totalPages) {
            if (rbRangeAll.value) return totalPages;
            if (rbRangeFirst.value) return totalPages > 0 ? 1 : 0;
            return parsePageNumbers(customRangeText || detectedRangeText || '', totalPages).length;
        }

        /**
         * 「配置ページ数 / 総ページ数」の表示を更新する
         * @returns {void}
         */
        function updateTotalPagesLabel() {
            var totalPages = getDetectedTotalPages();
            var placedPages = getPlacedPageCount(totalPages);
            stTotalPages.text = totalPages > 0 ? (placedPages + ' / ' + totalPages) : '';
        }

        /**
         * 行数の表示を更新する
         * 列数が自動のときは折り返し位置がページ幅次第で、見開きを分割するときは並ぶ数が計測するまで分からないため、空欄にしてディム表示にする
         * @returns {void}
         */
        function updateRowsInfo() {
            var colsPerRow = getColsPerRow();
            var placedPages = getPlacedPageCount(getDetectedTotalPages());
            var known = colsPerRow > 0 && !cbSplitSpreads.value;

            stRowsLabel.enabled = known;
            stRowsValue.enabled = known;
            stRowsValue.text = (known && placedPages > 0) ? String(Math.ceil(placedPages / colsPerRow)) : '';
        }

        /**
         * 倍率欄の有効／無効を切り替える
         * @returns {void}
         */
        function updateScaleEnabledState() {
            var enabled = !rbPerArtboard.value;
            etScale.enabled = enabled;
            stScaleLabel.enabled = enabled;
            stScaleUnit.enabled = enabled;
            scaleStepper.enabled = enabled;
            redrawSteppersIn(scaleStepper); /* ∧∨のディム表示を切り替える / update stepper dimming */
        }

        /**
         * 配置方法に合わせて間隔欄のツールチップを切り替える
         * @returns {void}
         */
        function updateGapHelpTip() {
            var helpText = rbPerArtboard.value ? getLabel(LABELS.tooltip.gapPerArtboard) : getLabel(LABELS.tooltip.gapIgnoreArtboard);
            stGapLabel.helpTip = helpText;
            etArtboardGap.helpTip = helpText;
            stGapUnit.helpTip = helpText;
        }

        // ------------------------
        // UI表示補助
        // ------------------------

        /**
         * ファイル表示用の名前とパスを求める
         * @param {File} f - 対象ファイル
         * @returns {object} name と path を持つオブジェクト
         */
        function getFileDisplayInfo(f) {
            if (!f) {
                return {
                    name: getLabel(LABELS.label.notSelected),
                    path: ''
                };
            }

            var nameText;
            var pathText;

            try {
                nameText = decodeURIComponent(f.name);
            } catch (e) {
                nameText = String(f.name || getLabel(LABELS.label.notSelected));
            }

            try {
                pathText = decodeURIComponent(f.fsName);
            } catch (e) {
                pathText = String(f.fsName || '');
            }

            return {
                name: nameText,
                path: pathText
            };
        }

        /**
         * File.openDialog のフィルター。フォルダーは辿れるよう通す
         * @param {File} fileObj - 判定対象
         * @returns {boolean} 表示するなら true
         */
        function isPdfOrAiFile(fileObj) {
            if (!fileObj) return false;
            if (fileObj instanceof Folder) return true;
            return isPdfLikeFile(fileObj);
        }

        /**
         * 読み込み対象ファイルを設定し、名前とパスの表示を更新する
         * @param {File} f - 読み込み元ファイル。null で未指定に戻す
         * @returns {void}
         */
        function setPathText(f) {
            sourceFile = f || null;
            updatePanelEnabledState();

            /* 綴じ方向から偶数ページの位置を設定する / Set the even-page side from the binding direction */
            if (sourceFile) {
                if (isRightBoundFile(sourceFile)) rbEvenPageRight.value = true;
                else rbEvenPageLeft.value = true;
            }

            var info = getFileDisplayInfo(f);
            stSourceName.text = info.name;
            stSourceName.helpTip = info.path;
        }

        /**
         * ∧∨・↑↓キーで値を変えたあと、入力欄の onChange を呼ぶ（列数の「自動」への正規化・行数の更新）
         * @param {EditText} numberInput - 値を変えた入力欄
         * @returns {void}
         */
        function notifyFieldChange(numberInput) {
            if (typeof numberInput.onChange === "function") numberInput.onChange();
        }

        // ------------------------
        // 配置オプション取得
        // ------------------------

        /**
         * UI で指定した 1 行あたりの列数を取得する
         * @returns {number} 列数。自動のときは 0
         */
        function getColsPerRow() {
            var v = parseInt(etCols.text, 10);
            if (isNaN(v) || v <= 0) return 0;
            return v;
        }

        /**
         * UI で指定したページ間隔を取得する
         * @returns {number} 間隔（pt）
         */
        function getArtboardGap() {
            var v = parseFloat(etArtboardGap.text);
            if (isNaN(v) || v < 0) v = DEFAULT_ARTBOARD_GAP;
            return v;
        }

        /**
         * UI で指定した配置倍率を取得する
         * @returns {number} 配置倍率（%）
         */
        function getScaleFromUI() {
            var v = parseFloat(etScale.text);
            if (isNaN(v) || v <= 0) v = 100;
            return v;
        }

        /**
         * UI で指定したケイ処理の設定を取得する
         * @returns {object} ケイ処理の設定。処理しない場合は null
         */
        function getKeiOptionsFromUI() {
            if (rbKeiNone.value) {
                return null;
            }
            var radius = parseFloat(etRoundCorner.text);
            if (isNaN(radius) || radius < 0) radius = 0;
            return {
                mode: 'clipGroup',
                roundCorners: cbRoundCorner.value,
                roundRadius: radius * getUnitInfo().pointsPerUnit
            };
        }

        /**
         * UI で指定したトリミングを CropTo の値で取得する。AI ファイルは DEFAULT_CROP_MODE
         * @returns {number} CropTo の値
         */
        function getCropModeFromUI() {
            if (!isPdfFile(sourceFile) || !ddCropMode.selection) return DEFAULT_CROP_MODE;
            return ddCropMode.selection.index;
        }

        /**
         * UI で指定した対象ページ番号の配列を取得する
         * @returns {Array} ページ番号の配列
         */
        function getTargetPagesFromUI() {
            if (rbRangeAll.value) {
                var lastPage = getDetectedTotalPages();
                if (!lastPage || lastPage < 1) lastPage = 1;
                var pages = [];
                for (var p = 1; p <= lastPage; p++) {
                    pages.push(p);
                }
                return pages;
            }
            if (rbRangeFirst.value) {
                return [1];
            }
            var custom = parsePageNumbers(etRange.text, getDetectedTotalPages());
            return (custom && custom.length > 0) ? custom : [1];
        }

        // ------------------------
        // 進捗表示
        // ダイアログを閉じた後に使うため、独立したパレットで表示する
        // Shown in its own palette because the dialog is already closed
        // ------------------------

        /**
         * 進捗表示用のパレットを開く。開けなくても処理は継続する
         * @returns {void}
         */
        function showProgress() {
            try {
                progressWin = new Window("palette", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
                progressWin.alignChildren = "fill";
                progressBar = progressWin.add("progressbar", undefined, 0, 100);
                progressBar.preferredSize.width = 300;
                progressBar.preferredSize.height = 8;
                progressWin.show();
                progressWin.update();
            } catch (e) {
                progressWin = null;
                progressBar = null;
            }
        }

        /**
         * 進捗を更新する
         * @param {number} current - 処理済みの件数
         * @param {number} total - 総件数
         * @returns {void}
         */
        function updateProgress(current, total) {
            if (!progressWin || !progressBar || !total) return;
            try {
                progressBar.value = Math.round((current / total) * 100);
                progressWin.update();
            } catch (e) { }
        }

        /**
         * 進捗表示を閉じる
         * @returns {void}
         */
        function hideProgress() {
            if (!progressWin) return;
            try { progressWin.close(); } catch (e) { }
            progressWin = null;
            progressBar = null;
        }

        // ------------------------
        // 実行処理
        // ------------------------

        /**
         * 指定ページを配置する。アートボードごと／無視、現在／新規ドキュメント、見開きの分割に対応
         * @param {object} request - OK で控えた配置指示
         * @param {Array} request.pages - 対象ページ番号の配列
         * @param {object} request.keiOpts - ケイ処理の設定。null で処理なし
         * @param {boolean} request.perArtboard - true でページごとにアートボードを作成
         * @param {number} request.scale - 配置倍率（%）。perArtboard が true のときは無視
         * @param {number} request.cropMode - トリミング（CropTo の値）
         * @param {boolean} request.splitSpreads - 見開きを左右に分割するなら true
         * @param {boolean} request.evenPageOnRight - 偶数ページを右に置くなら true
         * @param {boolean} request.placeInNewDoc - 新規ドキュメントに配置するなら true
         * @param {DocumentColorSpace} request.colorSpace - 新規ドキュメントのカラーモード
         * @returns {void}
         */
        function placePages(request) {
            var targetPages = request.pages;
            var keiOpts = request.keiOpts;
            var perArtboard = request.perArtboard;
            var scale = request.scale;
            var cropMode = request.cropMode;
            var scaleFactor = perArtboard ? 1 : (scale / 100);
            var gap = getArtboardGap();
            var colsPerRow = getColsPerRow();

            // crop 環境設定は計測で書き換わるため、計測より前に退避する
            // Snapshot the crop preferences before measuring, which overwrites them
            var cropSnapshot = SC_snapshotPdfCropPreference();
            var targetDoc = doc;
            var abCount = 0;
            var skippedCount = 0;
            var completed = false;
            var placedItems = [];
            showProgress();

            try {
                // 全ページを一度だけ計測し、レイアウト全体をカンバス中心へ揃える
                // Measure every page once, then center the whole layout on the canvas
                var measured = placementMeasurePages(doc, sourceFile, targetPages, cropMode);
                if (request.splitSpreads) measured = placementSplitSpreads(measured, request.evenPageOnRight);
                var layout = placementBuildLayout(measured, scaleFactor, gap, colsPerRow);
                if (request.placeInNewDoc) targetDoc = createOutputDocument(request.colorSpace, layout, doc);
                var activeIdx = targetDoc.artboards.getActiveArtboardIndex();
                var canvas = getLargestCanvasBounds(targetDoc);
                var startX = (canvas[0] + canvas[2]) / 2 - layout.width / 2;
                var baseTop = (canvas[1] + canvas[3]) / 2 + layout.height / 2;

                for (var i = 0; i < layout.slots.length; i++) {
                    var slot = layout.slots[i];
                    if (!slot) {
                        skippedCount++; // 計測できなかったページ / Page that could not be measured
                    } else {
                        try {
                            var left = startX + slot.x;
                            var top = baseTop + slot.y;

                            var slotRect = [left, top, left + slot.width, top - slot.height];

                            if (perArtboard) {
                                abCount = placementUseOrAddArtboard(targetDoc, activeIdx, slotRect, abCount);
                            }

                            /* 右半分は見開き全体を半ページ分左へずらして置き、枠の外をマスクで隠す
                               The right half is placed one half-width to the left, and the rest is masked out */
                            var placeLeft = (slot.half === "right") ? left - slot.width : left;
                            var item = placementPlacePage(targetDoc, sourceFile, slot.page, [placeLeft, top], cropMode, perArtboard ? 100 : scale);
                            if (slot.half) item = keiCreateClippingMaskGroup(item, slotRect);
                            if (keiOpts) {
                                try { item = keiApplyToPlacedItem(targetDoc, item, keiOpts); } catch (e) { }
                            }
                            placedItems.push(item);
                        } catch (e) {
                            skippedCount++; // 失敗ページはスキップして継続 / Skip the failed page and continue
                        }
                    }
                    updateProgress(i + 1, layout.slots.length);
                }
                completed = true;
            } catch (e) {
                SC_alert(LABELS.alert.placeError, e);
            } finally {
                try { placementResetImportPageNumber(); } catch (e) { }
                try { SC_restorePdfCropPreference(cropSnapshot); } catch (e) { }
                hideProgress();
            }

            if (!completed) return;

            if (perArtboard) {
                placementFitBoundsInView(targetDoc, placementGetArtboardsBounds(targetDoc));
            } else {
                // 配置したオブジェクトを選択して表示を合わせる / Select the placed objects and fit the view
                try { targetDoc.selection = null; } catch (e) { }
                for (var j = 0; j < placedItems.length; j++) {
                    try { placedItems[j].selected = true; } catch (e) { }
                }
                placementFitBoundsInView(targetDoc, placementGetItemsVisibleBounds(placedItems));
            }

            if (skippedCount > 0) SC_alert(LABELS.alert.someSkipped);
        }

        // ------------------------
        // UIイベント配線
        // ------------------------

        /**
         * ダイアログのイベントハンドラーを割り当てる
         * @returns {void}
         */
        function bindDialogEvents() {

            /**
             * ページ範囲まわりの表示をまとめて更新する
             * @returns {void}
             */
            function refreshRangeInfo() {
                updateRangeEnabledState();
                updateTotalPagesLabel();
                updateRowsInfo();
            }

            /**
             * 角丸まわりの有効／無効を切り替える
             * @returns {void}
             */
            function updateKeiRoundEnabled() {
                var keiActive = !rbKeiNone.value;
                cbRoundCorner.enabled = keiActive;
                etRoundCorner.enabled = keiActive && cbRoundCorner.value;
            }

            btnBrowse.onClick = function () {
                var f = File.openDialog(getLabel(LABELS.dialog.pickFile), isPdfOrAiFile);
                if (!f) return;
                var last = updatePageCountFromPlacedOrFile(doc, f, setPathText);
                detectedRangeText = last ? ('1-' + last) : '';
                if (!customRangeText) {
                    customRangeText = detectedRangeText;
                }
                refreshRangeInfo();
            };

            rbRangeAll.onClick = refreshRangeInfo;
            rbRangeFirst.onClick = refreshRangeInfo;
            rbRangeCustom.onClick = refreshRangeInfo;

            etRange.onChanging = function () {
                if (!rbRangeCustom.value) return;
                customRangeText = etRange.text;
                updateTotalPagesLabel();
                updateRowsInfo();
            };
            etRange.onChange = etRange.onChanging;

            etCols.onChanging = updateRowsInfo;
            etCols.onChange = function () {
                // 0 以下や数値でない入力は自動扱いなので、確定時に表示も「自動」へ揃える
                // Zero or non-numeric input means Auto; normalize the field text on commit
                var autoText = getLabel(LABELS.label.columnsAuto);
                if (getColsPerRow() === 0 && etCols.text !== autoText) etCols.text = autoText;
                updateRowsInfo();
            };

            rbPerArtboard.onClick = function () {
                updateScaleEnabledState();
                updateGapHelpTip();
                updateRowsInfo();
            };
            rbIgnoreArtboard.onClick = rbPerArtboard.onClick;

            cbSplitSpreads.onClick = function () {
                updateEvenPageEnabledState();
                updateRowsInfo();
            };

            rbDestinationCurrent.onClick = updateColorModeEnabledState;
            rbDestinationNew.onClick = updateColorModeEnabledState;

            rbKeiNone.onClick = updateKeiRoundEnabled;
            rbKeiClipGroup.onClick = updateKeiRoundEnabled;
            cbRoundCorner.onClick = updateKeiRoundEnabled;
            updateKeiRoundEnabled();

            btnCancel.onClick = function () {
                win.close(2);
            };

            btnOk.onClick = function () {
                if (!sourceFile) {
                    SC_alert(LABELS.alert.needFile);
                    return;
                }
                // 配置指示だけ控えて閉じる。実行は win.show() の後
                // Record the request and close; the placement runs after win.show()
                var perArtboard = rbPerArtboard.value;
                placementRequest = {
                    pages: getTargetPagesFromUI(),
                    keiOpts: getKeiOptionsFromUI(),
                    perArtboard: perArtboard,
                    scale: perArtboard ? 100 : getScaleFromUI(),
                    cropMode: getCropModeFromUI(),
                    splitSpreads: cbSplitSpreads.value,
                    evenPageOnRight: rbEvenPageRight.value,
                    placeInNewDoc: rbDestinationNew.value,
                    colorSpace: rbColorRGB.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK
                };
                win.close(1);
            };
        }

        // ------------------------
        // 初期状態反映
        // ------------------------

        /**
         * ダイアログの初期状態を反映する
         * @returns {void}
         */
        function initializeDialogState() {
            updatePanelEnabledState();

            var last = updatePageCountFromPlacedOrFile(doc, null, setPathText);
            detectedRangeText = last ? ('1-' + last) : '';
            if (!customRangeText) {
                customRangeText = detectedRangeText;
            }
            updateRangeEnabledState();
            updateTotalPagesLabel();
            updateScaleEnabledState();
            updateGapHelpTip();
            updateEvenPageEnabledState();
            updateColorModeEnabledState();
            updateRowsInfo();

            stRoundCornerUnit.text = getUnitInfo().label;
        }

        bindDialogEvents();
        initializeDialogState();

        prepareDialogWindow(win, SCRIPT_NAME);
        if (win.show() !== 1 || !placementRequest) return;

        // モーダル表示中は executeMenuCommand が無視されることがあり、閉じ処理も詰まりやすい
        // Run after the dialog closes: executeMenuCommand can be ignored while a modal dialog is up
        placePages(placementRequest);

    }

    main();

}());
