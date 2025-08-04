#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name：
PreferenceManager.jsx

### 概要 / Overview：
- Illustrator の各種環境設定をダイアログボックスから変更可能にします。
- 単位、文字設定、変形／整列設定などをワンパネルで調整できます。
- Allows changing various Illustrator preferences via a dialog box.
- Units, text settings, transform/align settings can be adjusted in one panel.

### 主な機能 / Main Features：
- 単位（一般・線・文字・東アジア言語）の設定 / Set units (general, stroke, text, East Asian)
- 最近使用したフォント数とフォント表記切替 / Recent fonts count and font name localization
- 変形と整列の各種オプション / Various transform and align options
- 字形の境界に整列の設定 / Align to glyph bounds setting
- モード切替（プリント pt／プリント Q／オンスクリーン） / Mode switch (Print pt / Print Q / Onscreen)

### 参考・謝辞

https://judicious-night-bca.notion.site/app-getIntegerPreference-e4088e6caef64b3ba5b801d86fba7877

### 更新履歴 / Change Log：
- v1.0 (20250804): 初期バージョン / Initial version
- v1.1 (20250804): ダイアログを2カラムに改修、単位とフォント設定を追加 / Dialog changed to two columns; added units and font settings
- v1.2 (20250804): 角の拡大のロジックを修正 / Fixed logic for corner scaling
*/


/* スクリプトバージョン / Script Version */
var SCRIPT_VERSION = "v1.2";

/* 現在のUI言語を取得 / Get the current UI language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* UIラベル定義 / UI Label Definitions */
var LABELS = {
    dialogTitle: {
        ja: "まとめて環境設定 " + SCRIPT_VERSION,
        en: "Preferences " + SCRIPT_VERSION
    },
    modePrintPt: {
        ja: "プリント（pt）",
        en: "Print (pt)"
    },
    modePrintQ: {
        ja: "プリント（Q）",
        en: "Print (Q)"
    },
    modeOnscreen: {
        ja: "オンスクリーン（px）",
        en: "Onscreen (px)"
    },
    unitsTitle: {
        ja: "単位",
        en: "Units"
    },
    // --- Inserted localization for unit dropdowns ---
    generalUnit: {
        ja: "一般",
        en: "General"
    },
    strokeUnit: {
        ja: "線",
        en: "Stroke"
    },
    textUnit: {
        ja: "文字",
        en: "Type"
    },
    asianUnit: {
        ja: "東アジア言語",
        en: "East Asian Type"
    },
    // -----------------------------------------------
    textTitle: {
        ja: "テキスト",
        en: "Text"
    },
    fontEnglish: {
        ja: "フォント名を英語表記",
        en: "Show Font Names in English"
    },
    recentFonts: {
        ja: "最近使用したフォント",
        en: "Recent Fonts"
    },
    transformTitle: {
        ja: "変形と整列",
        en: "Transform & Align"
    },
    previewBounds: {
        ja: "プレビュー境界",
        en: "Preview Bounds"
    },
    transformPattern: {
        ja: "パターンを変形",
        en: "Transform Pattern Tiles"
    },
    scaleCorners: {
        ja: "角を拡大・縮小",
        en: "Scale Corners"
    },
    scaleStroke: {
        ja: "線幅と効果も拡大・縮小",
        en: "Scale Strokes & Effects"
    },
    realtimeDrawing: {
        ja: "リアルタイムの描画と編集",
        en: "Real-time Drawing & Editing"
    },
    glyphBounds: {
        ja: "字形の境界に整列",
        en: "Align to Glyph Bounds"
    },
    pointText: {
        ja: "ポイント文字",
        en: "Point Type"
    },
    areaText: {
        ja: "エリア内文字",
        en: "Area Type"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    leadingLabel: {
        ja: "サイズ/行送り：",
        en: "Size/Leading:"
    },
    baselineLabel: {
        ja: "ベースライン：",
        en: "Baseline Shift:"
    },
    // ---- Added missing localized labels ----
    generalTitle: {
        ja: "一般",
        en: "General"
    },
    keyInputLabel: {
        ja: "キー入力：",
        en: "Keyboard Increment::"
    },
    cornerRadiusLabel: {
        ja: "角丸の半径：",
        en: "Corner Radius:"
    },
    textDetailTitle: {
        ja: "テキスト",
        en: "Text"
    }
};


/* ↑↓キーで値を増減 / Change value with arrow keys */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal place */
        } else {
            value = Math.round(value);
        }

        editText.text = value;
        editText.notify("onChange"); /* 値変更をトリガー / Trigger value change */
    });
}

/* 単位コードとラベルのマップ / Unit Code to Label Map */
var unitLabelMap = {
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

/* 単位換算ユーティリティ / Unit conversion utilities */
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0; // in
        case 1:
            return 72.0 / 25.4; // mm
        case 2:
            return 1.0; // pt
        case 3:
            return 12.0; // pica
        case 4:
            return 72.0 / 2.54; // cm
        case 5:
            return 72.0 / 25.4 * 0.25; // Q or H
        case 6:
            return 1.0; // px
        case 7:
            return 72.0 * 12.0; // ft/in
        case 8:
            return 72.0 / 25.4 * 1000.0; // m
        case 9:
            return 72.0 * 36.0; // yd
        case 10:
            return 72.0 * 12.0; // ft
        default:
            return 1.0;
    }
}

function convertFromPt(valuePt, unitCode) {
    return valuePt / getPtFactorFromUnitCode(unitCode);
}

function convertToPt(valueUnit, unitCode) {
    return valueUnit * getPtFactorFromUnitCode(unitCode);
}



/* 複数チェックボックスを環境設定キーにバインド / Bind multiple checkboxes to preference keys */
function bindCheckboxes(pairs) {
    for (var i = 0; i < pairs.length; i++) {
        (function(pair) {
            pair.checkbox.onClick = function() {
                app.preferences.setBooleanPreference(pair.prefKey, pair.checkbox.value === true);
            };
        })(pairs[i]);
    }
}

/* ダイアログ位置を調整 / Adjust dialog position */
function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function() {
        var currentX = dlg.location[0];
        var currentY = dlg.location[1];
        dlg.location = [currentX + offsetX, currentY + offsetY];
    };
}

/* ダイアログの透明度を設定 / Set dialog opacity */
function setDialogOpacity(dlg, opacityValue) {
    dlg.opacity = opacityValue;
}

/* メイン処理 / Main entry point */
function main() {
    var pref = app.preferences;

    var dialog = new Window('dialog');
    dialog.text = LABELS.dialogTitle[lang];
    dialog.orientation = 'column';
    dialog.alignChildren = ['fill', 'top'];


    /* ダイアログ位置と透明度の調整 / Adjust dialog position and opacity */
    var offsetX = 300;
    var dialogOpacity = 0.97;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    /* モード選択ラジオボタン / Mode Selection Radio Buttons */
    var modeGroup = dialog.add('group');
    modeGroup.orientation = 'row';
    modeGroup.alignChildren = ['center', 'center']; // 中央揃え
    modeGroup.alignment = ['center', 'top']; // グループ自体を中央に配置
    modeGroup.margins = [15, 10, 15, 10];

    var radioPrintPt = modeGroup.add('radiobutton', undefined, LABELS.modePrintPt[lang]);
    var radioPrintQ = modeGroup.add('radiobutton', undefined, LABELS.modePrintQ[lang]);
    var radioOnscreen = modeGroup.add('radiobutton', undefined, LABELS.modeOnscreen[lang]);

    /* デフォルトをプリント（pt）に設定 / Default selection */
    radioPrintPt.value = true;

    /* 2カラムのメインコンテナ / Two-column main container */
    var mainGroup = dialog.add('group');
    mainGroup.orientation = 'row';
    mainGroup.alignChildren = ['fill', 'top'];

    /* 左カラム / Left column */
    var leftColumn = mainGroup.add('group');
    leftColumn.orientation = 'column';
    leftColumn.alignChildren = ['fill', 'top'];

    /* 単位パネルを追加 / Add "Units" panel */
    var unitsPanel = leftColumn.add('panel', undefined, LABELS.unitsTitle[lang]);
    unitsPanel.orientation = 'column';
    unitsPanel.alignChildren = ['left', 'top'];
    unitsPanel.margins = [8, 20, 8, 15];

    /* 一般パネルを追加 / Add "General" panel */
    var generalPanel = leftColumn.add('panel', undefined, LABELS.generalTitle[lang]);
    generalPanel.orientation = 'column';
    generalPanel.alignChildren = ['left', 'top'];
    generalPanel.margins = [8, 20, 8, 15];

    // 単位ラベル取得ヘルパー
    function getUnitLabel() {
        var code = app.preferences.getIntegerPreference("rulerType");
        return unitLabelMap[code] || "pt";
    }

    // キー入力
    var groupKeyInput = generalPanel.add('group');
    groupKeyInput.orientation = 'row';
    var labelKey = groupKeyInput.add('statictext', undefined, LABELS.keyInputLabel[lang]);
    labelKey.characters = 12;
    labelKey.justify = 'right';
    var unitCodeKey = app.preferences.getIntegerPreference("rulerType");
    var keyValuePt = pref.getRealPreference("cursorKeyLength");
    var keyValue = convertFromPt(keyValuePt, unitCodeKey);
    var inputKey = groupKeyInput.add('edittext', undefined, keyValue.toFixed(1));
    inputKey.characters = 4;
    var unitLabelKey = groupKeyInput.add('statictext', undefined, getUnitLabel());
    unitLabelKey.characters = 4; // 幅を広げる
    inputKey.onChange = function() {
        var value = parseFloat(inputKey.text);
        if (!isNaN(value)) {
            var ptValue = convertToPt(value, app.preferences.getIntegerPreference("rulerType"));
            pref.setRealPreference("cursorKeyLength", ptValue);
            refreshValues();
        }
    };
    /* ↑↓キー操作を適用 / Apply arrow key value change */
    changeValueByArrowKey(inputKey);


    // 角丸の半径
    var groupCornerRadius = generalPanel.add('group');
    groupCornerRadius.orientation = 'row';
    var labelCorner = groupCornerRadius.add('statictext', undefined, LABELS.cornerRadiusLabel[lang]);
    labelCorner.characters = 12;
    labelCorner.justify = 'right';
    var cornerUnitCode = app.preferences.getIntegerPreference("rulerType");
    var cornerValuePt = pref.getRealPreference("ovalRadius");
    var cornerValue = convertFromPt(cornerValuePt, cornerUnitCode);
    var inputCornerRadius = groupCornerRadius.add('edittext', undefined, cornerValue.toFixed(1));
    inputCornerRadius.characters = 4;
    var unitLabelCorner = groupCornerRadius.add('statictext', undefined, getUnitLabel());
    unitLabelCorner.characters = 4; // 幅を広げる
    inputCornerRadius.onChange = function() {
        var value = parseFloat(inputCornerRadius.text);
        if (!isNaN(value)) {
            var ptValue = convertToPt(value, app.preferences.getIntegerPreference("rulerType"));
            pref.setRealPreference("ovalRadius", ptValue);
            refreshValues();
        }
    };
    /* ↑↓キー操作を適用 / Apply arrow key value change */
    changeValueByArrowKey(inputCornerRadius);


    /* テキスト詳細パネルを追加 / Add "Text Details" panel */
    // --- Unit label helpers for text detail panel ---
    function getTextUnitLabel() {
        var code = app.preferences.getIntegerPreference("text/units");
        return unitLabelMap[code] || "pt";
    }

    function getAsianUnitLabel() {
        var code = app.preferences.getIntegerPreference("text/asianunits");
        return unitLabelMap[code] || "pt";
    }
    var textDetailPanel = leftColumn.add('panel', undefined, LABELS.textDetailTitle[lang]);
    textDetailPanel.orientation = 'column';
    textDetailPanel.alignChildren = ['left', 'top'];
    textDetailPanel.margins = [8, 20, 8, 15];

    // サイズ行送り
    var groupLeading = textDetailPanel.add('group');
    groupLeading.orientation = 'row';
    var labelLeading = groupLeading.add('statictext', undefined, LABELS.leadingLabel[lang]);
    labelLeading.characters = 12;
    labelLeading.justify = 'right';
    var sizeUnitCode = app.preferences.getIntegerPreference("text/units");
    var sizeValuePt = pref.getRealPreference("text/sizeIncrement");
    var sizeValue = convertFromPt(sizeValuePt, sizeUnitCode);
    var inputLeading = groupLeading.add('edittext', undefined, sizeValue.toFixed(1));
    inputLeading.characters = 4;
    var unitLabelLeading = groupLeading.add('statictext', undefined, getTextUnitLabel());
    unitLabelLeading.characters = 4; // 幅を広げる
    inputLeading.onChange = function() {
        var value = parseFloat(inputLeading.text);
        if (!isNaN(value)) {
            var ptValue = convertToPt(value, app.preferences.getIntegerPreference("text/units"));
            pref.setRealPreference("text/sizeIncrement", ptValue);
            refreshValues();
        }
    };
    /* ↑↓キー操作を適用 / Apply arrow key value change */
    changeValueByArrowKey(inputLeading);



    // ベースラインシフト
    var groupBaseline = textDetailPanel.add('group');
    groupBaseline.orientation = 'row';
    var labelBaseline = groupBaseline.add('statictext', undefined, LABELS.baselineLabel[lang]);
    labelBaseline.characters = 12;
    labelBaseline.justify = 'right';
    var baselineUnitCode = app.preferences.getIntegerPreference("text/asianunits");
    var baselineValuePt = pref.getRealPreference("text/riseIncrement");
    var baselineValue = convertFromPt(baselineValuePt, baselineUnitCode);
    var inputBaseline = groupBaseline.add('edittext', undefined, baselineValue.toFixed(1));
    inputBaseline.characters = 4;
    var unitLabelBaseline = groupBaseline.add('statictext', undefined, getAsianUnitLabel());
    unitLabelBaseline.characters = 4; // 幅を広げる
    inputBaseline.onChange = function() {
        var value = parseFloat(inputBaseline.text);
        if (!isNaN(value)) {
            var ptValue = convertToPt(value, app.preferences.getIntegerPreference("text/asianunits"));
            pref.setRealPreference("text/riseIncrement", ptValue);
            refreshValues();
        }
    };
    /* ↑↓キー操作を適用 / Apply arrow key value change */
    changeValueByArrowKey(inputBaseline);


    var unitDropdowns = {};

    function createUnitDropdown(parent, label, prefKey) {
        var group = parent.add('group');
        group.orientation = 'row';
        var labelControl = group.add('statictext', undefined, label + "：");
        labelControl.characters = 12;
        labelControl.justify = 'right';

        var dropdown = group.add('dropdownlist', undefined, []);
        dropdown.characters = 9; // ← 幅を6文字分に指定
        /* ラベルを追加 */
        for (var code in unitLabelMap) {
            var labelText = unitLabelMap[code];
            dropdown.add('item', labelText);
        }

        var currentCode = app.preferences.getIntegerPreference(prefKey);
        dropdown.selection = dropdown.find(unitLabelMap[currentCode]) || dropdown.find("pt");

        dropdown.onChange = function() {
            var selectedLabel = dropdown.selection.text;
            for (var c in unitLabelMap) {
                if (unitLabelMap[c] === selectedLabel) {
                    app.preferences.setIntegerPreference(prefKey, parseInt(c, 10));
                    break;
                }
            }
        };

        unitDropdowns[prefKey] = dropdown;
    }

    /* 各プルダウンを作成 */
    createUnitDropdown(unitsPanel, LABELS.generalUnit[lang], "rulerType");
    createUnitDropdown(unitsPanel, LABELS.strokeUnit[lang], "strokeUnits");
    createUnitDropdown(unitsPanel, LABELS.textUnit[lang], "text/units");
    createUnitDropdown(unitsPanel, LABELS.asianUnit[lang], "text/asianunits");

    function setUnitsForMode(mode) {
        if (mode === "printPt") {
            unitDropdowns["rulerType"].selection = unitDropdowns["rulerType"].find("mm"); // 一般は mm
            unitDropdowns["strokeUnits"].selection = unitDropdowns["strokeUnits"].find("pt");
            unitDropdowns["text/units"].selection = unitDropdowns["text/units"].find("pt");
            unitDropdowns["text/asianunits"].selection = unitDropdowns["text/asianunits"].find("pt");

            // 値を更新（すべてptで保存、UIはmm/pt表示）
            var rulerCode = 1; // mm
            var textCode = 2; // pt
            var asianCode = 2; // pt
            pref.setRealPreference("cursorKeyLength", convertToPt(0.1, rulerCode));
            pref.setRealPreference("ovalRadius", convertToPt(1, rulerCode));
            pref.setRealPreference("text/sizeIncrement", convertToPt(1.0, textCode));
            pref.setRealPreference("text/riseIncrement", convertToPt(0.1, asianCode));

        } else if (mode === "printQ") {
            unitDropdowns["rulerType"].selection = unitDropdowns["rulerType"].find("mm");
            unitDropdowns["strokeUnits"].selection = unitDropdowns["strokeUnits"].find("mm");
            unitDropdowns["text/units"].selection = unitDropdowns["text/units"].find("Q/H");
            unitDropdowns["text/asianunits"].selection = unitDropdowns["text/asianunits"].find("Q/H");

            // 値を更新（すべてptで保存、UIはmm/Q/H表示）
            var rulerCodeQ = 1; // mm
            var textCodeQ = 5; // Q/H
            var asianCodeQ = 5; // Q/H
            pref.setRealPreference("cursorKeyLength", convertToPt(1, rulerCodeQ));
            pref.setRealPreference("ovalRadius", convertToPt(2, rulerCodeQ));
            pref.setRealPreference("text/sizeIncrement", convertToPt(1, textCodeQ));
            pref.setRealPreference("text/riseIncrement", convertToPt(0.1, asianCodeQ));

        } else if (mode === "onscreen") {
            unitDropdowns["rulerType"].selection = unitDropdowns["rulerType"].find("px");
            unitDropdowns["strokeUnits"].selection = unitDropdowns["strokeUnits"].find("px");
            unitDropdowns["text/units"].selection = unitDropdowns["text/units"].find("px");
            unitDropdowns["text/asianunits"].selection = unitDropdowns["text/asianunits"].find("px");

            // 値を更新
            var pxCode = 6;
            pref.setRealPreference("cursorKeyLength", convertToPt(1, pxCode));
            pref.setRealPreference("ovalRadius", convertToPt(1, pxCode));
            pref.setRealPreference("text/sizeIncrement", convertToPt(1, pxCode));
            pref.setRealPreference("text/riseIncrement", convertToPt(0.5, pxCode));
        }

        /* Trigger onChange manually to update preferences */
        for (var key in unitDropdowns) {
            if (unitDropdowns[key].selection) {
                unitDropdowns[key].onChange();
            }
        }

        refreshUnitLabels();
        refreshValues();
    }

    // --- Helper to refresh unit labels in panels ---
    function refreshUnitLabels() {
        // Update General panel unit labels
        groupKeyInput.children[groupKeyInput.children.length - 1].text = getUnitLabel();
        groupCornerRadius.children[groupCornerRadius.children.length - 1].text = getUnitLabel();
        // Update Text Detail panel unit labels
        groupLeading.children[groupLeading.children.length - 1].text = getTextUnitLabel();
        groupBaseline.children[groupBaseline.children.length - 1].text = getAsianUnitLabel();
    }

    // --- Helper to refresh numeric field values with unit conversion ---
    function refreshValues() {
        var unitCodeKey = app.preferences.getIntegerPreference("rulerType");
        inputKey.text = convertFromPt(pref.getRealPreference("cursorKeyLength"), unitCodeKey).toFixed(1);

        var cornerUnitCode = app.preferences.getIntegerPreference("rulerType");
        inputCornerRadius.text = convertFromPt(pref.getRealPreference("ovalRadius"), cornerUnitCode).toFixed(1);

        var sizeUnitCode = app.preferences.getIntegerPreference("text/units");
        inputLeading.text = convertFromPt(pref.getRealPreference("text/sizeIncrement"), sizeUnitCode).toFixed(1);

        var baselineUnitCode = app.preferences.getIntegerPreference("text/asianunits");
        inputBaseline.text = convertFromPt(pref.getRealPreference("text/riseIncrement"), baselineUnitCode).toFixed(1);
    }

    radioPrintPt.onClick = function() {
        setUnitsForMode("printPt");
        refreshUnitLabels();
    };
    radioPrintQ.onClick = function() {
        setUnitsForMode("printQ");
        refreshUnitLabels();
    };
    radioOnscreen.onClick = function() {
        setUnitsForMode("onscreen");
        refreshUnitLabels();
    };


    /* 右カラム / Right column */
    var rightColumn = mainGroup.add('group');
    rightColumn.orientation = 'column';
    rightColumn.alignChildren = ['fill', 'top'];

    /* テキストパネルを追加 / Add "Text" panel */
    var textPanel = rightColumn.add('panel', undefined, LABELS.textTitle[lang]);
    textPanel.orientation = 'column';
    textPanel.alignChildren = ['left', 'top'];
    textPanel.margins = [8, 20, 8, 15];

    var checkboxFontEnglish = textPanel.add('checkbox', undefined, LABELS.fontEnglish[lang]);
    checkboxFontEnglish.value = app.preferences.getBooleanPreference("text/useEnglishFontNames");
    checkboxFontEnglish.onClick = function() {
        app.preferences.setBooleanPreference("text/useEnglishFontNames", checkboxFontEnglish.value === true);
    };

    /* --- 新しい 最近使用したフォントの表示数 ロジック（UI調整） --- */
    var currentRecentCount = app.preferences.getIntegerPreference("text/recentFontMenu/showNEntries");
    var groupRecentFonts = textPanel.add('group');
    groupRecentFonts.orientation = 'row';

    var checkboxRecentFonts = groupRecentFonts.add('checkbox', undefined, LABELS.recentFonts[lang]);
    checkboxRecentFonts.value = (currentRecentCount > 0);

    var inputRecentFonts = groupRecentFonts.add('edittext', undefined, currentRecentCount.toString());
    inputRecentFonts.characters = 3;
    inputRecentFonts.enabled = checkboxRecentFonts.value;

    checkboxRecentFonts.onClick = function() {
        if (checkboxRecentFonts.value) {
            if (parseInt(inputRecentFonts.text, 10) === 0 || isNaN(parseInt(inputRecentFonts.text, 10))) {
                inputRecentFonts.text = "1";
            }
            inputRecentFonts.enabled = true;
            app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", parseInt(inputRecentFonts.text, 10));
        } else {
            inputRecentFonts.enabled = false;
            inputRecentFonts.text = "0";
            app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", 0);
        }
    };

    inputRecentFonts.onChange = function() {
        var value = parseInt(inputRecentFonts.text, 10);
        if (!isNaN(value)) {
            app.preferences.setIntegerPreference("text/recentFontMenu/showNEntries", value);
            checkboxRecentFonts.value = (value > 0);
            inputRecentFonts.enabled = checkboxRecentFonts.value;
        }
    };


    var glyphPanel = rightColumn.add('panel', undefined, LABELS.glyphBounds[lang]);
    glyphPanel.orientation = 'column';
    glyphPanel.alignChildren = ['left', 'top'];
    glyphPanel.margins = [8, 20, 8, 15];

    var checkboxPoint = glyphPanel.add('checkbox', undefined, LABELS.pointText[lang]);
    checkboxPoint.value = app.preferences.getBooleanPreference('EnableActualPointTextSpaceAlign');

    var checkboxArea = glyphPanel.add('checkbox', undefined, LABELS.areaText[lang]);
    checkboxArea.value = app.preferences.getBooleanPreference('EnableActualAreaTextSpaceAlign');

    bindCheckboxes([{
            checkbox: checkboxPoint,
            prefKey: 'EnableActualPointTextSpaceAlign'
        },
        {
            checkbox: checkboxArea,
            prefKey: 'EnableActualAreaTextSpaceAlign'
        }
    ]);



    /* その他パネルを追加 / Add "Transform & Align" panel */
    var otherPanel = rightColumn.add('panel', undefined, LABELS.transformTitle[lang]);
    otherPanel.orientation = 'column';
    otherPanel.alignChildren = ['left', 'top'];
    otherPanel.margins = [8, 20, 8, 15];

    //　プレビュー境界
    var checkboxPreview = otherPanel.add('checkbox', undefined, LABELS.previewBounds[lang]);
    checkboxPreview.value = app.preferences.getBooleanPreference("includeStrokeInBounds");
    checkboxPreview.onClick = function() {
        app.preferences.setBooleanPreference("includeStrokeInBounds", checkboxPreview.value === true);
    };


    // パターンを変形
    var checkboxPattern = otherPanel.add('checkbox', undefined, LABELS.transformPattern[lang]);
    checkboxPattern.value = app.preferences.getBooleanPreference("transformPatterns");
    checkboxPattern.onClick = function() {
        app.preferences.setBooleanPreference("transformPatterns", checkboxPattern.value === true);
    };

    // 角を拡大・縮小
    var checkboxCorner = otherPanel.add('checkbox', undefined, LABELS.scaleCorners[lang]);
    // 初期値を取得（1=ON, 2=OFF）
    checkboxCorner.value = (app.preferences.getIntegerPreference("policyForPreservingCorners") === 1);
    checkboxCorner.onClick = function() {
        app.preferences.setIntegerPreference(
            "policyForPreservingCorners",
            checkboxCorner.value ? 1 : 2
        );
    };

    /* 線幅と効果も拡大・縮小 */
    var checkboxStroke = otherPanel.add('checkbox', undefined, LABELS.scaleStroke[lang]);
    checkboxStroke.value = app.preferences.getBooleanPreference("scaleLineWeight");
    checkboxStroke.onClick = function() {
        app.preferences.setBooleanPreference("scaleLineWeight", checkboxStroke.value === true);
    };

    /* リアルタイムの描画と編集 */
    var checkboxRealtime = otherPanel.add('checkbox', undefined, LABELS.realtimeDrawing[lang]);
    checkboxRealtime.value = app.preferences.getBooleanPreference("LiveEdit_State_Machine");
    checkboxRealtime.onClick = function() {
        app.preferences.setBooleanPreference("LiveEdit_State_Machine", checkboxRealtime.value === true);
    };


    var group2 = dialog.add('group', undefined, {
        name: 'group2'
    });
    group2.orientation = 'row';
    group2.alignChildren = ['center', 'center']; /* 中央揃え */
    group2.alignment = ['center', 'bottom']; /* ダイアログ内で中央に配置 */

    var cancelBtn = group2.add('button', undefined, LABELS.cancel[lang], {
        name: 'cancel'
    });
    cancelBtn.preferredSize.width = 90;

    var okBtn = group2.add('button', undefined, LABELS.ok[lang], {
        name: 'ok'
    });
    okBtn.preferredSize.width = 90;

    dialog.show();
}

main();