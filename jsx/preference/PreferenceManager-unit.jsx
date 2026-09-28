#target illustrator
#targetengine "PreferenceManager-unitEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorの各種環境設定をダイアログから変更します。
単位、文字設定、変形／整列設定などを1つのパネルで調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManager-unit.md

### Overview

Changes a range of Illustrator preferences from a dialog.
Units, text settings and transform/align settings are all adjusted in a single panel.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager-unit.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PreferenceManager-unit";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManager-unit.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager-unit.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* モードのラジオボタンで切り替える単位コードと増減量（増減量は切り替え後の単位での値）
       Unit codes and increments applied by each mode radio (increments are in the new units) */
    var UNIT_MODE_PRESETS = {
        printPt: {
            units: { "rulerType": 1, "strokeUnits": 2, "text/units": 2, "text/asianunits": 2 },     /* mm / pt / pt / pt */
            increments: { "cursorKeyLength": 0.1, "ovalRadius": 1, "text/sizeIncrement": 1, "text/riseIncrement": 0.1 }
        },
        printQ: {
            units: { "rulerType": 1, "strokeUnits": 1, "text/units": 5, "text/asianunits": 5 },     /* mm / mm / Q / H */
            increments: { "cursorKeyLength": 1, "ovalRadius": 2, "text/sizeIncrement": 1, "text/riseIncrement": 0.1 }
        },
        onscreen: {
            units: { "rulerType": 6, "strokeUnits": 6, "text/units": 6, "text/asianunits": 6 },     /* px */
            increments: { "cursorKeyLength": 1, "ovalRadius": 1, "text/sizeIncrement": 1, "text/riseIncrement": 0.5 }
        }
    };

    // =========================================
    // レイアウト / Layout
    // =========================================
    var MODE_ROW_MARGINS              = [15, 10, 15, 10];   /* モード行の余白 [左,上,右,下] */
    var PANEL_MARGINS                 = [8, 20, 8, 15];     /* パネル余白 [左,上,右,下] */
    var FIELD_LABEL_CHARACTERS        = 12;                 /* 項目名の幅 / field label width */
    var VALUE_FIELD_CHARACTERS        = 4;                  /* 数値欄の幅 / numeric field width */
    var UNIT_LABEL_CHARACTERS         = 4;                  /* 単位表示の幅 / unit label width */
    var UNIT_DROPDOWN_CHARACTERS      = 9;                  /* 単位ドロップダウンの幅 / unit dropdown width */
    var RECENT_FONTS_FIELD_CHARACTERS = 3;                  /* 最近使用したフォントの件数欄の幅 */
    var BUTTON_WIDTH                  = 90;                 /* ボタンの幅 / button width */

    /**
     * 縦並びのパネルを追加する
     * @param {Group} parent - 追加先
     * @param {string} labelPath - パネル見出しのラベルパス
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelPath) {
        var panel = parent.add('panel', undefined, getLabel(labelPath));
        panel.orientation = 'column';
        panel.alignChildren = ['left', 'top'];
        panel.margins = PANEL_MARGINS;
        return panel;
    }

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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 環境設定の項目 / Preference items
    // =========================================

    /* 単位のドロップダウン（上から順に並ぶ）/ Unit dropdowns, top to bottom */
    var UNIT_DROPDOWN_DEFINITIONS = [
        { prefKey: "rulerType",       labelPath: "fieldLabel.generalUnit", tipPath: "tooltip.generalUnit" },
        { prefKey: "strokeUnits",     labelPath: "fieldLabel.strokeUnit",  tipPath: "tooltip.strokeUnit" },
        { prefKey: "text/units",      labelPath: "fieldLabel.typeUnit",    tipPath: "tooltip.typeUnit" },
        { prefKey: "text/asianunits", labelPath: "fieldLabel.asianUnit",   tipPath: "tooltip.asianUnit" }
    ];

    /* 増減量の数値欄。値は pt で保存し、unitKey の単位で表示する
       Increment fields; values are stored in pt and shown in the unit of unitKey */
    var INCREMENT_FIELD_DEFINITIONS = [
        { panelId: "general", valueKey: "cursorKeyLength",    unitKey: "rulerType",       labelPath: "fieldLabel.keyIncrement",      tipPath: "tooltip.keyIncrement" },
        { panelId: "general", valueKey: "ovalRadius",         unitKey: "rulerType",       labelPath: "fieldLabel.cornerRadius",      tipPath: "tooltip.cornerRadius" },
        { panelId: "text",    valueKey: "text/sizeIncrement", unitKey: "text/units",      labelPath: "fieldLabel.sizeIncrement",     tipPath: "tooltip.sizeIncrement" },
        { panelId: "text",    valueKey: "text/riseIncrement", unitKey: "text/asianunits", labelPath: "fieldLabel.baselineIncrement", tipPath: "tooltip.baselineIncrement" }
    ];

    /* ON/OFF を真偽値で持つ環境設定のチェックボックス（パネルごと）
       Checkboxes bound to boolean preferences, per panel */
    var FONT_CHECKBOX_DEFINITIONS = [
        { prefKey: "text/useEnglishFontNames", labelPath: "checkbox.fontNamesInEnglish", tipPath: "tooltip.fontNamesInEnglish" }
    ];
    var GLYPH_BOUNDS_CHECKBOX_DEFINITIONS = [
        { prefKey: "EnableActualPointTextSpaceAlign", labelPath: "checkbox.pointType", tipPath: "tooltip.glyphBounds" },
        { prefKey: "EnableActualAreaTextSpaceAlign",  labelPath: "checkbox.areaType",  tipPath: "tooltip.glyphBounds" }
    ];
    var TRANSFORM_CHECKBOX_DEFINITIONS = [
        { prefKey: "includeStrokeInBounds", labelPath: "checkbox.previewBounds",    tipPath: "tooltip.previewBounds" },
        { prefKey: "transformPatterns",     labelPath: "checkbox.transformPattern", tipPath: "tooltip.transformPattern" }
    ];
    var TRANSFORM_TAIL_CHECKBOX_DEFINITIONS = [
        { prefKey: "scaleLineWeight",       labelPath: "checkbox.scaleStroke",      tipPath: "tooltip.scaleStroke" },
        { prefKey: "LiveEdit_State_Machine", labelPath: "checkbox.realtimeDrawing", tipPath: "tooltip.realtimeDrawing" }
    ];

    /* 「角を拡大・縮小」の環境設定値（1=ON, 2=OFF）/ Values of the scale-corners preference */
    var SCALE_CORNERS_ON = 1;
    var SCALE_CORNERS_OFF = 2;

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

    var LABELS = {
        dialog: {
            title: { ja: "まとめて環境設定", en: "Preferences" }
        },
        panel: {
            units: { ja: "単位", en: "Units" },
            general: { ja: "一般", en: "General" },
            textIncrement: { ja: "テキスト", en: "Text" },
            text: { ja: "テキスト", en: "Text" },
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            transform: { ja: "変形と整列", en: "Transform & Align" }
        },
        radio: {
            printPt: { ja: "プリント（pt）", en: "Print (pt)" },
            printQ: { ja: "プリント（Q）", en: "Print (Q)" },
            onscreen: { ja: "オンスクリーン（px）", en: "Onscreen (px)" }
        },
        fieldLabel: {
            generalUnit: { ja: "一般", en: "General" },
            strokeUnit: { ja: "線", en: "Stroke" },
            typeUnit: { ja: "文字", en: "Type" },
            asianUnit: { ja: "東アジア言語", en: "East Asian Type" },
            keyIncrement: { ja: "キー増加", en: "Keyboard Increment" },
            cornerRadius: { ja: "角丸の半径", en: "Corner Radius" },
            sizeIncrement: { ja: "サイズ/行送り", en: "Size/Leading" },
            baselineIncrement: { ja: "ベースライン", en: "Baseline Shift" }
        },
        checkbox: {
            fontNamesInEnglish: { ja: "フォント名を英語表記", en: "Show Font Names in English" },
            recentFonts: { ja: "最近使用したフォント", en: "Recent Fonts" },
            pointType: { ja: "ポイント文字", en: "Point Type" },
            areaType: { ja: "エリア内文字", en: "Area Type" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
            transformPattern: { ja: "パターンを変形", en: "Transform Pattern Tiles" },
            scaleCorners: { ja: "角を拡大・縮小", en: "Scale Corners" },
            scaleStroke: { ja: "線幅と効果も拡大・縮小", en: "Scale Strokes & Effects" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-time Drawing & Editing" }
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
            printPt: {
                ja: "一般=mm、線=pt、文字=pt、東アジア言語のオプション=pt にまとめて切り替えます。",
                en: "Sets General=mm, Stroke=pt, Text=pt, East Asian=pt."
            },
            printQ: {
                ja: "一般=mm、線=mm、文字=Q、東アジア言語のオプション=H にまとめて切り替えます。",
                en: "Sets General=mm, Stroke=mm, Text=Q, East Asian=H."
            },
            onscreen: {
                ja: "一般・線・文字・東アジア言語のオプションをすべて px に切り替えます。",
                en: "Sets General, Stroke, Text and East Asian all to px."
            },
            generalUnit: { ja: "定規やパネルに表示される、既定の長さの単位です。", en: "Default unit shown on rulers and panels." },
            strokeUnit: { ja: "線幅の入力・表示に使う単位です。", en: "Unit used for stroke weights." },
            typeUnit: { ja: "フォントサイズや行送りに使う単位です。", en: "Unit used for font size and leading." },
            asianUnit: { ja: "東アジア言語のオプションで使う単位です。", en: "Unit used for East Asian typography options." },
            keyIncrement: { ja: "矢印キー1回で動く距離です。", en: "How far one arrow key press moves things." },
            cornerRadius: { ja: "角丸ツールの既定の半径です。", en: "Default radius used by the rounded rectangle tool." },
            sizeIncrement: { ja: "文字サイズ・行送りを増減する1回ぶんの量です。", en: "How much one step changes the type size or leading." },
            baselineIncrement: { ja: "ベースラインシフトを増減する1回ぶんの量です。", en: "How much one step changes the baseline shift." },
            fontNamesInEnglish: { ja: "フォント名を英語表記で表示します。", en: "Shows font names in English." },
            recentFonts: {
                ja: "フォントメニューの先頭に並ぶ「最近使用したフォント」の表示件数です。0 で非表示になります。",
                en: "How many recently used fonts appear at the top of the font menu. 0 hides the list."
            },
            glyphBounds: { ja: "整列の基準を、仮想ボディではなく字形の実際の輪郭にします。", en: "Aligns text by the actual glyph outlines instead of the em box." },
            previewBounds: { ja: "線幅や効果を含めた見た目の端を、オブジェクトの境界として扱います。", en: "Treats the visible edges including strokes and effects as the object bounds." },
            transformPattern: { ja: "オブジェクトを変形したとき、塗りのパターンも一緒に変形します。", en: "Transforms the pattern fill along with the object." },
            scaleCorners: { ja: "拡大・縮小したとき、ライブコーナーの角丸も一緒に変わります。", en: "Scales live corner radii along with the object." },
            scaleStroke: { ja: "拡大・縮小したとき、線幅と効果も一緒に変わります。", en: "Scales stroke weights and effects along with the object." },
            realtimeDrawing: { ja: "ドラッグ中もオブジェクトの結果を表示しながら描画・編集します。", en: "Draws and edits with a live result while dragging." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

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

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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

    // =========================================
    // ダイアログの部品 / Dialog parts
    // =========================================

    /**
     * 右揃えの項目名つきの行を追加する
     * @param {Panel} parent - 追加先
     * @param {string} labelPath - 項目名のラベルパス
     * @returns {Group} 追加した行
     */
    function addFieldRow(parent, labelPath) {
        var fieldRow = parent.add('group');
        fieldRow.orientation = 'row';
        var fieldLabel = fieldRow.add('statictext', undefined, labelText(labelPath));
        fieldLabel.characters = FIELD_LABEL_CHARACTERS;
        fieldLabel.justify = 'right';
        return fieldRow;
    }

    /**
     * 単位のドロップダウンを追加する（項目の添字が単位コード）
     * @param {Panel} parent - 追加先
     * @param {{prefKey: string, labelPath: string, tipPath: string}} definition - ドロップダウンの定義
     * @returns {DropDownList} 追加したドロップダウン
     */
    function addUnitDropdown(parent, definition) {
        var prefKey = definition.prefKey;
        var unitRow = addFieldRow(parent, definition.labelPath);
        var unitDropdown = unitRow.add('dropdownlist', undefined, []);
        unitDropdown.characters = UNIT_DROPDOWN_CHARACTERS;
        unitDropdown.helpTip = getLabel(definition.tipPath);

        /* 単位コード5は環境設定キーによって Q か H / Code 5 reads Q or H depending on the key */
        for (var code = 0; code < UNITS.length; code++) {
            unitDropdown.add('item', (code === 5 && HA_UNIT_PREF_KEYS[prefKey]) ? "H" : UNITS[code].label);
        }

        var currentCode = app.preferences.getIntegerPreference(prefKey);
        unitDropdown.selection = UNITS[currentCode] ? currentCode : 2;

        unitDropdown.onChange = function () {
            if (unitDropdown.selection) {
                app.preferences.setIntegerPreference(prefKey, unitDropdown.selection.index);
            }
        };
        return unitDropdown;
    }

    /**
     * 増減量の数値欄（項目名・数値・単位）を追加する
     * @param {Panel} parent - 追加先
     * @param {{valueKey: string, unitKey: string, labelPath: string, tipPath: string}} definition - 数値欄の定義
     * @param {Object[]} incrementFields - 数値欄の部品の一覧（追加した欄をここに積む）
     * @returns {void}
     */
    function addIncrementField(parent, definition, incrementFields) {
        var fieldRow = addFieldRow(parent, definition.labelPath);

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldRow.add('group');
        stepperFieldGroup.orientation = 'row';
        stepperFieldGroup.alignChildren = ['left', 'center'];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        /* 下限0。増減後は onChange で環境設定へ書き込む / minimum 0; onChange writes the preference */
        var stepperGroup = addStepper(stepperFieldGroup, function () { return valueInput; }, {
            min: 0,
            onStep: function () { valueInput.notify("onChange"); }
        });
        var valueInput = stepperFieldGroup.add('edittext', undefined, "");
        valueInput.helpTip = getLabel(definition.tipPath);
        valueInput.characters = VALUE_FIELD_CHARACTERS;
        var unitText = fieldRow.add('statictext', undefined, getUnitInfo(definition.unitKey).label);
        unitText.characters = UNIT_LABEL_CHARACTERS;

        var incrementField = { definition: definition, input: valueInput, unitText: unitText };
        refreshIncrementValue(incrementField);
        incrementFields.push(incrementField);

        valueInput.onChange = function () {
            var value = parseFloat(valueInput.text);
            if (!isNaN(value)) {
                app.preferences.setRealPreference(definition.valueKey, value * getUnitInfo(definition.unitKey).pointsPerUnit);
                /* 単位ドロップダウンの変更も拾えるよう全欄を表示し直す / Refresh every field so unit changes show up too */
                for (var i = 0; i < incrementFields.length; i++) {
                    refreshIncrementValue(incrementFields[i]);
                }
            }
        };
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(valueInput, stepperGroup);
    }

    /**
     * 数値欄に環境設定の値を現在の単位で表示し直す
     * @param {{definition: Object, input: EditText}} incrementField - 数値欄の部品
     * @returns {void}
     */
    function refreshIncrementValue(incrementField) {
        var definition = incrementField.definition;
        var valuePt = app.preferences.getRealPreference(definition.valueKey);
        incrementField.input.text = (valuePt / getUnitInfo(definition.unitKey).pointsPerUnit).toFixed(1);
    }

    /**
     * 真偽値の環境設定に連動するチェックボックスを追加する
     * @param {Panel} parent - 追加先
     * @param {{prefKey: string, labelPath: string, tipPath: string}} definition - チェックボックスの定義
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addPrefCheckbox(parent, definition) {
        var prefCheckbox = parent.add('checkbox', undefined, getLabel(definition.labelPath));
        prefCheckbox.helpTip = getLabel(definition.tipPath);
        prefCheckbox.value = app.preferences.getBooleanPreference(definition.prefKey);
        prefCheckbox.onClick = function () {
            app.preferences.setBooleanPreference(definition.prefKey, prefCheckbox.value === true);
        };
        return prefCheckbox;
    }

    /**
     * 定義の並びどおりにチェックボックスを追加する
     * @param {Panel} parent - 追加先
     * @param {Object[]} definitions - チェックボックスの定義の配列
     * @returns {void}
     */
    function addPrefCheckboxes(parent, definitions) {
        for (var i = 0; i < definitions.length; i++) {
            addPrefCheckbox(parent, definitions[i]);
        }
    }

    /**
     * 「最近使用したフォント」の表示件数（チェックボックス＋件数欄）を追加する
     * @param {Panel} parent - 追加先
     * @returns {void}
     */
    function addRecentFontsRow(parent) {
        var prefKey = "text/recentFontMenu/showNEntries";
        var currentCount = app.preferences.getIntegerPreference(prefKey);
        var recentFontsRow = parent.add('group');
        recentFontsRow.orientation = 'row';

        var recentFontsCheckbox = recentFontsRow.add('checkbox', undefined, getLabel("checkbox.recentFonts"));
        recentFontsCheckbox.helpTip = getLabel("tooltip.recentFonts");
        recentFontsCheckbox.value = (currentCount > 0);

        var recentFontsInput = recentFontsRow.add('edittext', undefined, currentCount.toString());
        recentFontsInput.helpTip = getLabel("tooltip.recentFonts");
        recentFontsInput.characters = RECENT_FONTS_FIELD_CHARACTERS;
        recentFontsInput.enabled = recentFontsCheckbox.value;

        recentFontsCheckbox.onClick = function () {
            if (recentFontsCheckbox.value) {
                /* 0 や非数値のままONにしたら1件にする / Turning on with 0 or a non-number shows one entry */
                var count = parseInt(recentFontsInput.text, 10);
                if (count === 0 || isNaN(count)) {
                    recentFontsInput.text = "1";
                }
                recentFontsInput.enabled = true;
                app.preferences.setIntegerPreference(prefKey, parseInt(recentFontsInput.text, 10));
            } else {
                recentFontsInput.enabled = false;
                recentFontsInput.text = "0";
                app.preferences.setIntegerPreference(prefKey, 0);
            }
        };

        recentFontsInput.onChange = function () {
            var count = parseInt(recentFontsInput.text, 10);
            if (!isNaN(count)) {
                app.preferences.setIntegerPreference(prefKey, count);
                recentFontsCheckbox.value = (count > 0);
                recentFontsInput.enabled = recentFontsCheckbox.value;
            }
        };
    }

    /**
     * 「角を拡大・縮小」のチェックボックスを追加する（環境設定は 1=ON, 2=OFF の整数）
     * @param {Panel} parent - 追加先
     * @returns {void}
     */
    function addScaleCornersCheckbox(parent) {
        var prefKey = "policyForPreservingCorners";
        var scaleCornersCheckbox = parent.add('checkbox', undefined, getLabel("checkbox.scaleCorners"));
        scaleCornersCheckbox.helpTip = getLabel("tooltip.scaleCorners");
        scaleCornersCheckbox.value = (app.preferences.getIntegerPreference(prefKey) === SCALE_CORNERS_ON);
        scaleCornersCheckbox.onClick = function () {
            app.preferences.setIntegerPreference(prefKey, scaleCornersCheckbox.value ? SCALE_CORNERS_ON : SCALE_CORNERS_OFF);
        };
    }

    // =========================================
    // モード切り替え / Unit modes
    // =========================================

    /**
     * モードのプリセットどおりに単位と増減量を書き込み、表示を更新する
     * @param {string} modeKey - "printPt" / "printQ" / "onscreen"
     * @param {Object} unitDropdowns - 環境設定キー → 単位ドロップダウン
     * @param {Object[]} incrementFields - 増減量の数値欄の部品
     * @returns {void}
     */
    function applyUnitMode(modeKey, unitDropdowns, incrementFields) {
        var preset = UNIT_MODE_PRESETS[modeKey];
        var prefKey;
        var i;

        for (prefKey in unitDropdowns) {
            unitDropdowns[prefKey].selection = preset.units[prefKey];
        }

        /* 増減量はプリセットの単位で与え、pt に直して保存 / Increments are given in the preset units and stored in pt */
        for (i = 0; i < INCREMENT_FIELD_DEFINITIONS.length; i++) {
            var definition = INCREMENT_FIELD_DEFINITIONS[i];
            var unitCode = preset.units[definition.unitKey];
            app.preferences.setRealPreference(definition.valueKey, preset.increments[definition.valueKey] * UNITS[unitCode].pointsPerUnit);
        }

        /* 単位の環境設定を確実に書き込む / Make sure the unit preferences are written */
        for (prefKey in unitDropdowns) {
            if (unitDropdowns[prefKey].selection) {
                unitDropdowns[prefKey].onChange();
            }
        }

        for (i = 0; i < incrementFields.length; i++) {
            incrementFields[i].unitText.text = getUnitInfo(incrementFields[i].definition.unitKey).label;
            refreshIncrementValue(incrementFields[i]);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * モード選択のラジオボタン行を追加する
     * @param {Window} dialog - ダイアログ
     * @param {Object} unitDropdowns - 環境設定キー → 単位ドロップダウン
     * @param {Object[]} incrementFields - 増減量の数値欄の部品
     * @returns {RadioButton[]} 追加したラジオボタン
     */
    function addModeRadios(dialog, unitDropdowns, incrementFields) {
        var modeRow = dialog.add('group');
        modeRow.orientation = 'row';
        modeRow.alignChildren = ['center', 'center'];
        modeRow.alignment = ['center', 'top'];
        modeRow.margins = MODE_ROW_MARGINS;

        var modeKeys = ["printPt", "printQ", "onscreen"];
        var modeRadios = [];
        for (var i = 0; i < modeKeys.length; i++) {
            (function (modeKey) {
                var modeRadio = modeRow.add('radiobutton', undefined, getLabel("radio." + modeKey));
                modeRadio.helpTip = getLabel("tooltip." + modeKey);
                modeRadio.onClick = function () {
                    applyUnitMode(modeKey, unitDropdowns, incrementFields);
                };
                modeRadios.push(modeRadio);
            })(modeKeys[i]);
        }
        return modeRadios;
    }

    /**
     * 環境設定ダイアログを作って表示する
     * @returns {void}
     */
    function main() {
        var dialog = new Window('dialog');
        dialog.text = getLabel("dialog.title") + " " + SCRIPT_VERSION;
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];

        /* ラジオ行は先頭に置くが、押したときに参照する部品はあとで埋める
           The mode row comes first; the controls it updates are filled in below */
        var unitDropdowns = {};
        var incrementFields = [];
        var modeRadios = addModeRadios(dialog, unitDropdowns, incrementFields);

        /* 表示時はどのモードも未選択にする / No mode is selected when the dialog opens */
        dialog.onShow = function () {
            for (var i = 0; i < modeRadios.length; i++) {
                modeRadios[i].value = false;
            }
        };

        /* 2カラムのメインコンテナ / Two-column main container */
        var columnsGroup = dialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignChildren = ['fill', 'top'];

        /* 左カラム：単位と増減量 / Left column: units and increments */
        var leftColumn = columnsGroup.add('group');
        leftColumn.orientation = 'column';
        leftColumn.alignChildren = ['fill', 'top'];

        var unitsPanel = addPanel(leftColumn, "panel.units");
        var incrementPanels = {
            general: addPanel(leftColumn, "panel.general"),
            text: addPanel(leftColumn, "panel.textIncrement")
        };

        var i;
        for (i = 0; i < INCREMENT_FIELD_DEFINITIONS.length; i++) {
            var fieldDefinition = INCREMENT_FIELD_DEFINITIONS[i];
            addIncrementField(incrementPanels[fieldDefinition.panelId], fieldDefinition, incrementFields);
        }
        for (i = 0; i < UNIT_DROPDOWN_DEFINITIONS.length; i++) {
            var dropdownDefinition = UNIT_DROPDOWN_DEFINITIONS[i];
            unitDropdowns[dropdownDefinition.prefKey] = addUnitDropdown(unitsPanel, dropdownDefinition);
        }

        /* 右カラム：テキスト・字形の境界・変形と整列 / Right column: text, glyph bounds, transform */
        var rightColumn = columnsGroup.add('group');
        rightColumn.orientation = 'column';
        rightColumn.alignChildren = ['fill', 'top'];

        var textPanel = addPanel(rightColumn, "panel.text");
        addPrefCheckboxes(textPanel, FONT_CHECKBOX_DEFINITIONS);
        addRecentFontsRow(textPanel);

        addPrefCheckboxes(addPanel(rightColumn, "panel.glyphBounds"), GLYPH_BOUNDS_CHECKBOX_DEFINITIONS);

        var transformPanel = addPanel(rightColumn, "panel.transform");
        addPrefCheckboxes(transformPanel, TRANSFORM_CHECKBOX_DEFINITIONS);
        addScaleCornersCheckbox(transformPanel);
        addPrefCheckboxes(transformPanel, TRANSFORM_TAIL_CHECKBOX_DEFINITIONS);

        /* ボタン行（中央）/ Button row (centered) */
        var buttonRow = addButtonRow(dialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add('button', undefined, getLabel("button.cancel"), { name: 'cancel' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        var btnOK = buttonRow.rowGroup.add('button', undefined, getLabel("button.ok"), { name: 'ok' });
        btnOK.preferredSize.width = BUTTON_WIDTH;

        prepareDialogWindow(dialog, SCRIPT_NAME);
        dialog.show();
    }

    main();

})();
