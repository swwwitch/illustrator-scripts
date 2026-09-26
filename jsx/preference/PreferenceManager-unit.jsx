#target illustrator
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
var SCRIPT_VERSION  = "v1.2.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

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
    var DIALOG_OPACITY                = 0.97;               /* ダイアログの不透明度 / dialog opacity */
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

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

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

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.units" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 入力補助 / Input helpers
    // =========================================

    /**
     * ↑↓キーで数値欄の値を増減する（Shift=10刻み、Option=0.1刻み）
     * @param {EditText} editText - 対象の数値欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            var isDown = (event.keyName == "Down");

            if (keyboard.shiftKey) {
                /* 10の倍数にスナップ / Snap to multiples of 10 */
                if (isUp) {
                    value = Math.ceil((value + 1) / 10) * 10;
                    event.preventDefault();
                } else if (isDown) {
                    value = Math.floor((value - 1) / 10) * 10;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                if (isUp) {
                    value += 0.1;
                    event.preventDefault();
                } else if (isDown) {
                    value -= 0.1;
                    event.preventDefault();
                }
            } else {
                if (isUp) {
                    value += 1;
                    event.preventDefault();
                } else if (isDown) {
                    value -= 1;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            /* Option 時は小数第1位、それ以外は整数に丸める / Round to 0.1 with Option, otherwise to integers */
            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);

            editText.text = value;
            editText.notify("onChange"); /* 値変更をトリガー / Trigger value change */
        });
    }

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
        var valueInput = fieldRow.add('edittext', undefined, "");
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
        /* ↑↓キー操作を適用 / Apply arrow key value change */
        changeValueByArrowKey(valueInput);
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
        dialog.opacity = DIALOG_OPACITY;

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
        var btnRowGroup = dialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['center', 'center'];
        btnRowGroup.alignment = ['center', 'bottom'];

        var btnCancel = btnRowGroup.add('button', undefined, getLabel("button.cancel"), { name: 'cancel' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        var btnOK = btnRowGroup.add('button', undefined, getLabel("button.ok"), { name: 'ok' });
        btnOK.preferredSize.width = BUTTON_WIDTH;

        dialog.show();
    }

    main();

})();
