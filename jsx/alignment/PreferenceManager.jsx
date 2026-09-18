#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorの各種環境設定をダイアログから変更します。
単位、文字設定、変形／整列設定などを1つのパネルで調整できます。

詳細は README を参照してください。

### Overview

Changes a range of Illustrator preferences from a dialog.
Units, text settings and transform/align settings are all adjusted in a single panel.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PreferenceManager";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-08-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2024-08-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManager.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X = 300;              /* ダイアログの表示位置：右(+)／左(-) / shift right (+) / left (-) */
    var DIALOG_OFFSET_Y = 0;                /* ダイアログの表示位置：下(+)／上(-) / shift down (+) / up (-) */
    var DIALOG_OPACITY = 0.97;              /* ダイアログの不透明度 0.0 - 1.0 */
    var PANEL_MARGINS = [15, 20, 15, 15];   /* パネル余白 [左,上,右,下] / panel margins */
    var MODE_ROW_MARGINS = [15, 10, 15, 10];/* モード選択行の余白 / margins of the mode row */
    var UNIT_LABEL_CHARS = 14;              /* 単位ラベルの予約幅（文字数）/ reserved width of the unit labels */
    var RECENT_COUNT_CHARS = 3;             /* 表示件数欄の幅（文字数）/ width of the count field */
    var BUTTON_WIDTH = 90;

    // =========================================
    // 単位の定義 / Unit tables
    // =========================================

    /* 単位コードとラベルの対応 / Unit code to label */
    var UNIT_LABEL_BY_CODE = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    /* 単位を設定する環境設定キーと、対応するラベル / Unit preference keys with their labels */
    var UNIT_PREF_KEYS = [
        { prefKey: "rulerType",        labelPath: "fieldLabel.generalUnit", tooltipPath: "tooltip.generalUnit" },
        { prefKey: "strokeUnits",      labelPath: "fieldLabel.strokeUnit",  tooltipPath: "tooltip.strokeUnit" },
        { prefKey: "text/units",       labelPath: "fieldLabel.textUnit",    tooltipPath: "tooltip.textUnit" },
        { prefKey: "text/asianunits",  labelPath: "fieldLabel.asianUnit",   tooltipPath: "tooltip.asianUnit" }
    ];

    /* モードごとの単位プリセット。値は UNIT_LABEL_BY_CODE のラベル
       Unit presets per mode; values are labels from UNIT_LABEL_BY_CODE */
    var UNIT_PRESETS = {
        printPt: {
            "rulerType": "mm",
            "strokeUnits": "pt",
            "text/units": "pt",
            "text/asianunits": "pt"
        },
        printQ: {
            "rulerType": "mm",
            "strokeUnits": "mm",
            "text/units": "Q/H",
            "text/asianunits": "Q/H"
        },
        onscreen: {
            "rulerType": "px",
            "strokeUnits": "px",
            "text/units": "px",
            "text/asianunits": "px"
        }
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "まとめて環境設定", en: "Preferences" }
        },
        radio: {
            modePrintPt:  { ja: "プリント（pt）", en: "Print (pt)" },
            modePrintQ:   { ja: "プリント（Q）", en: "Print (Q)" },
            modeOnscreen: { ja: "オンスクリーン", en: "Onscreen" }
        },
        panel: {
            units:       { ja: "単位", en: "Units" },
            text:        { ja: "テキスト", en: "Text" },
            transform:   { ja: "変形と整列", en: "Transform & Align" },
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" }
        },
        fieldLabel: {
            generalUnit: { ja: "一般", en: "General" },
            strokeUnit:  { ja: "線", en: "Stroke" },
            textUnit:    { ja: "文字", en: "Text" },
            asianUnit:   { ja: "東アジア言語のオプション", en: "East Asian" }
        },
        checkbox: {
            fontEnglish:     { ja: "フォント名を英語表記", en: "Show font names in English" },
            recentFonts:     { ja: "最近使用したフォント", en: "Recent fonts" },
            previewBounds:   { ja: "プレビュー境界", en: "Preview Bounds" },
            transformPattern:{ ja: "パターンを変形", en: "Transform Patterns" },
            scaleCorners:    { ja: "角を拡大・縮小", en: "Scale Corners" },
            scaleStroke:     { ja: "線幅を拡大・縮小", en: "Scale Strokes" },
            scaleEffects:    { ja: "効果を拡大・縮小", en: "Scale Effects" },
            realtimeDrawing: { ja: "リアルタイムに描画と編集", en: "Real-time Drawing & Editing" },
            pointText:       { ja: "ポイント文字", en: "Point Text" },
            areaText:        { ja: "エリア内文字", en: "Area Text" }
        },
        unit: {
            recentFontsCount: { ja: "件", en: "items" }
        },
        tooltip: {
            modePrintPt:  { ja: "一般=mm、線=pt、文字=pt、東アジア言語のオプション=pt にまとめて切り替えます。", en: "Sets General=mm, Stroke=pt, Text=pt, East Asian=pt." },
            modePrintQ:   { ja: "一般=mm、線=mm、文字=Q/H、東アジア言語のオプション=Q/H にまとめて切り替えます。", en: "Sets General=mm, Stroke=mm, Text=Q/H, East Asian=Q/H." },
            modeOnscreen: { ja: "一般・線・文字・東アジア言語のオプションをすべて px に切り替えます。", en: "Sets General, Stroke, Text and East Asian all to px." },
            generalUnit:  { ja: "定規やパネルに表示される、既定の長さの単位です。", en: "Default unit shown on rulers and panels." },
            strokeUnit:   { ja: "線幅の入力・表示に使う単位です。", en: "Unit used for stroke weights." },
            textUnit:     { ja: "フォントサイズや行送りに使う単位です。", en: "Unit used for font size and leading." },
            asianUnit:    { ja: "東アジア言語のオプションで使う単位です。", en: "Unit used for East Asian typography options." },
            recentFonts:  { ja: "フォントメニューの先頭に並ぶ「最近使用したフォント」の表示件数です。0 で非表示になります。", en: "How many recently used fonts appear at the top of the font menu. 0 hides the list." },
            previewBounds:{ ja: "線幅や効果を含めた見た目の端を、オブジェクトの境界として扱います。", en: "Treats the visible edges including strokes and effects as the object bounds." },
            glyphBounds:  { ja: "整列の基準を、仮想ボディではなく字形の実際の輪郭にします。", en: "Aligns text by the actual glyph outlines instead of the em box." }
        },
        button: {
            ok:    { ja: "OK", en: "OK" },
            close: { ja: "閉じる", en: "Close" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.units" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split('.');
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのドット区切りキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ": ");
    }

    // =========================================
    // UI部品 / UI helpers
    // =========================================

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} panel - 対象のパネル
     * @returns {void}
     */
    function setupPanel(panel) {
        panel.orientation = 'column';
        panel.alignChildren = ['left', 'top'];
        panel.margins = PANEL_MARGINS;
    }

    /**
     * ダイアログの表示位置をずらす
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} offsetX - 横方向のオフセット
     * @param {number} offsetY - 縦方向のオフセット
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            targetDialog.location = [
                targetDialog.location[0] + offsetX,
                targetDialog.location[1] + offsetY
            ];
        };
    }

    /**
     * 真偽値の環境設定に連動するチェックボックスを追加する
     * @param {Panel|Group} parentContainer - 追加先のコンテナ
     * @param {string} labelPath - ラベルのドット区切りキー
     * @param {string} prefKey - 環境設定キー
     * @param {string} [tooltipPath] - ツールチップのドット区切りキー
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addPreferenceCheckbox(parentContainer, labelPath, prefKey, tooltipPath) {
        var checkbox = parentContainer.add('checkbox', undefined, getLabel(labelPath));
        if (tooltipPath) checkbox.helpTip = getLabel(tooltipPath);
        checkbox.value = app.preferences.getBooleanPreference(prefKey);
        checkbox.onClick = function () {
            app.preferences.setBooleanPreference(prefKey, checkbox.value === true);
        };
        return checkbox;
    }

    // =========================================
    // 単位のプルダウン / Unit dropdowns
    // =========================================

    /* 環境設定キーごとのプルダウン。モード切り替えから参照する
       Dropdowns by preference key, referenced when switching modes */
    var unitDropdowns = {};

    /**
     * 単位のプルダウンを1行追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {object} unitPrefEntry - UNIT_PREF_KEYS の1項目
     * @returns {void}
     */
    function addUnitDropdown(parentPanel, unitPrefEntry) {
        var unitRowGroup = parentPanel.add('group');
        unitRowGroup.orientation = 'row';

        var unitLabel = unitRowGroup.add('statictext', undefined, labelText(unitPrefEntry.labelPath));
        unitLabel.characters = UNIT_LABEL_CHARS;
        unitLabel.justify = 'right';

        var unitDropdown = unitRowGroup.add('dropdownlist', undefined, []);
        unitDropdown.helpTip = getLabel(unitPrefEntry.tooltipPath);
        for (var unitCode in UNIT_LABEL_BY_CODE) {
            unitDropdown.add('item', UNIT_LABEL_BY_CODE[unitCode]);
        }

        var currentCode = app.preferences.getIntegerPreference(unitPrefEntry.prefKey);
        unitDropdown.selection = unitDropdown.find(UNIT_LABEL_BY_CODE[currentCode]) || unitDropdown.find("pt");

        unitDropdown.onChange = function () {
            var unitCodeForLabel = findUnitCodeByLabel(unitDropdown.selection.text);
            if (unitCodeForLabel !== null) {
                app.preferences.setIntegerPreference(unitPrefEntry.prefKey, unitCodeForLabel);
            }
        };

        unitDropdowns[unitPrefEntry.prefKey] = unitDropdown;
    }

    /**
     * 単位ラベルから環境設定に渡す単位コードを求める
     * @param {string} unitLabel - "mm" などの単位ラベル
     * @returns {number|null} 単位コード。対応するものがなければ null
     */
    function findUnitCodeByLabel(unitLabel) {
        for (var unitCode in UNIT_LABEL_BY_CODE) {
            if (UNIT_LABEL_BY_CODE[unitCode] === unitLabel) return parseInt(unitCode, 10);
        }
        return null;
    }

    /**
     * モードのプリセットに合わせて単位のプルダウンをまとめて切り替える
     * @param {string} unitPresetName - UNIT_PRESETS のキー（"printPt" / "printQ" / "onscreen"）
     * @returns {void}
     */
    function applyUnitPreset(unitPresetName) {
        var unitPreset = UNIT_PRESETS[unitPresetName];
        if (!unitPreset) return;

        for (var prefKey in unitPreset) {
            var unitDropdown = unitDropdowns[prefKey];
            if (!unitDropdown) continue;
            unitDropdown.selection = unitDropdown.find(unitPreset[prefKey]);
            /* 選択の差し替えでは onChange が呼ばれないため、明示的に環境設定へ反映する
               Assigning selection does not fire onChange, so push it to the preferences by hand */
            if (unitDropdown.selection) unitDropdown.onChange();
        }
    }

    // =========================================
    // パネルの組み立て / Panels
    // =========================================

    /**
     * ［単位］パネルを追加する
     * @param {Group} parentColumn - 追加先のカラム
     * @returns {void}
     */
    function addUnitsPanel(parentColumn) {
        var unitsPanel = parentColumn.add('panel', undefined, getLabel('panel.units'));
        setupPanel(unitsPanel);

        for (var i = 0; i < UNIT_PREF_KEYS.length; i++) {
            addUnitDropdown(unitsPanel, UNIT_PREF_KEYS[i]);
        }
    }

    /**
     * ［テキスト］パネルを追加する
     * @param {Group} parentColumn - 追加先のカラム
     * @returns {void}
     */
    function addTextPanel(parentColumn) {
        var textPanel = parentColumn.add('panel', undefined, getLabel('panel.text'));
        setupPanel(textPanel);

        addPreferenceCheckbox(textPanel, 'checkbox.fontEnglish', 'text/useEnglishFontNames');
        addRecentFontsRow(textPanel);
    }

    /**
     * ［最近使用したフォント］の表示件数の行を追加する
     * チェックボックスと件数欄は同じ環境設定（0 で非表示）を別の見え方で扱う。
     * @param {Panel} parentPanel - 追加先のパネル
     * @returns {void}
     */
    function addRecentFontsRow(parentPanel) {
        var RECENT_FONTS_PREF_KEY = "text/recentFontMenu/showNEntries";
        var currentCount = app.preferences.getIntegerPreference(RECENT_FONTS_PREF_KEY);

        var recentFontsRowGroup = parentPanel.add('group');
        recentFontsRowGroup.orientation = 'row';

        var recentFontsCheckbox = recentFontsRowGroup.add('checkbox', undefined, getLabel('checkbox.recentFonts'));
        recentFontsCheckbox.helpTip = getLabel('tooltip.recentFonts');
        recentFontsCheckbox.value = (currentCount > 0);

        var recentFontsCountInput = recentFontsRowGroup.add('edittext', undefined, currentCount.toString());
        recentFontsCountInput.helpTip = getLabel('tooltip.recentFonts');
        recentFontsCountInput.characters = RECENT_COUNT_CHARS;
        recentFontsCountInput.enabled = recentFontsCheckbox.value;

        recentFontsRowGroup.add('statictext', undefined, getLabel('unit.recentFontsCount'));

        recentFontsCheckbox.onClick = function () {
            if (!recentFontsCheckbox.value) {
                recentFontsCountInput.enabled = false;
                recentFontsCountInput.text = "0";
                app.preferences.setIntegerPreference(RECENT_FONTS_PREF_KEY, 0);
                return;
            }

            /* ONに戻したとき 0 のままでは非表示のままなので 1 から始める / 0 would keep the list hidden */
            var enteredCount = parseInt(recentFontsCountInput.text, 10);
            if (isNaN(enteredCount) || enteredCount === 0) {
                recentFontsCountInput.text = "1";
                enteredCount = 1;
            }
            recentFontsCountInput.enabled = true;
            app.preferences.setIntegerPreference(RECENT_FONTS_PREF_KEY, enteredCount);
        };

        recentFontsCountInput.onChange = function () {
            var enteredCount = parseInt(recentFontsCountInput.text, 10);
            if (isNaN(enteredCount)) return;

            app.preferences.setIntegerPreference(RECENT_FONTS_PREF_KEY, enteredCount);
            recentFontsCheckbox.value = (enteredCount > 0);
            recentFontsCountInput.enabled = recentFontsCheckbox.value;
        };
    }

    /**
     * ［字形の境界に整列］パネルを追加する
     * @param {Group} parentColumn - 追加先のカラム
     * @returns {void}
     */
    function addGlyphBoundsPanel(parentColumn) {
        var glyphBoundsPanel = parentColumn.add('panel', undefined, getLabel('panel.glyphBounds'));
        setupPanel(glyphBoundsPanel);

        addPreferenceCheckbox(glyphBoundsPanel, 'checkbox.pointText', 'EnableActualPointTextSpaceAlign', 'tooltip.glyphBounds');
        addPreferenceCheckbox(glyphBoundsPanel, 'checkbox.areaText', 'EnableActualAreaTextSpaceAlign', 'tooltip.glyphBounds');
    }

    /**
     * ［変形と整列］パネルを追加する
     * @param {Group} parentColumn - 追加先のカラム
     * @returns {void}
     */
    function addTransformPanel(parentColumn) {
        var transformPanel = parentColumn.add('panel', undefined, getLabel('panel.transform'));
        setupPanel(transformPanel);

        addPreferenceCheckbox(transformPanel, 'checkbox.previewBounds', 'includeStrokeInBounds', 'tooltip.previewBounds');
        addPreferenceCheckbox(transformPanel, 'checkbox.transformPattern', 'transformPatterns');
        addPreferenceCheckbox(transformPanel, 'checkbox.scaleCorners', 'scaleCorners');
        addPreferenceCheckbox(transformPanel, 'checkbox.scaleStroke', 'scaleLineWeight');
        addPreferenceCheckbox(transformPanel, 'checkbox.scaleEffects', 'includeStrokeInBounds');
        addPreferenceCheckbox(transformPanel, 'checkbox.realtimeDrawing', 'realTimeDrawing');
    }

    /**
     * モード選択のラジオボタン行を追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addModeRow(targetDialog) {
        var modeRowGroup = targetDialog.add('group');
        modeRowGroup.orientation = 'row';
        modeRowGroup.alignChildren = ['center', 'center'];
        modeRowGroup.alignment = ['center', 'top'];
        modeRowGroup.margins = MODE_ROW_MARGINS;

        var modeEntries = [
            { presetName: 'printPt',  labelPath: 'radio.modePrintPt',  tooltipPath: 'tooltip.modePrintPt' },
            { presetName: 'printQ',   labelPath: 'radio.modePrintQ',   tooltipPath: 'tooltip.modePrintQ' },
            { presetName: 'onscreen', labelPath: 'radio.modeOnscreen', tooltipPath: 'tooltip.modeOnscreen' }
        ];

        for (var i = 0; i < modeEntries.length; i++) {
            addModeRadio(modeRowGroup, modeEntries[i], i === 0);
        }
    }

    /**
     * モード選択のラジオボタンを1つ追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {object} modeEntry - presetName / labelPath / tooltipPath を持つ定義
     * @param {boolean} isDefault - 既定で選択状態にする場合は true
     * @returns {void}
     */
    function addModeRadio(parentGroup, modeEntry, isDefault) {
        var modeRadio = parentGroup.add('radiobutton', undefined, getLabel(modeEntry.labelPath));
        modeRadio.helpTip = getLabel(modeEntry.tooltipPath);
        modeRadio.value = isDefault;
        modeRadio.onClick = function () {
            applyUnitPreset(modeEntry.presetName);
        };
    }

    /**
     * ボタンエリアを追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addButtonRow(targetDialog) {
        var btnRowGroup = targetDialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['center', 'center'];
        btnRowGroup.alignment = ['center', 'bottom'];

        /* 設定はクリックのたびに即反映されるため、取り消しではなく「閉じる」
           Preferences are written on each click, so this closes rather than cancels */
        var btnClose = btnRowGroup.add('button', undefined, getLabel('button.close'), { name: 'cancel' });
        btnClose.preferredSize.width = BUTTON_WIDTH;
        btnClose.onClick = function () {
            targetDialog.close();
        };

        var btnOK = btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnOK.preferredSize.width = BUTTON_WIDTH;
        btnOK.onClick = function () {
            targetDialog.close();
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを組み立てて表示する
     * @returns {void}
     */
    function main() {
        var preferencesDialog = new Window('dialog');
        preferencesDialog.text = getLabel('dialog.title') + ' ' + SCRIPT_VERSION;
        preferencesDialog.orientation = 'column';
        preferencesDialog.alignChildren = ['fill', 'top'];
        preferencesDialog.opacity = DIALOG_OPACITY;
        shiftDialogPosition(preferencesDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

        addModeRow(preferencesDialog);

        var columnsGroup = preferencesDialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignChildren = ['fill', 'top'];

        var leftColumn = columnsGroup.add('group');
        leftColumn.orientation = 'column';
        leftColumn.alignChildren = ['fill', 'top'];

        var rightColumn = columnsGroup.add('group');
        rightColumn.orientation = 'column';
        rightColumn.alignChildren = ['fill', 'top'];

        addGlyphBoundsPanel(rightColumn);
        addUnitsPanel(leftColumn);
        addTextPanel(leftColumn);
        addTransformPanel(rightColumn);

        addButtonRow(preferencesDialog);

        preferencesDialog.show();
    }

    main();

})();
