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

### 参考・謝辞 / References & Thanks：
https://github.com/ten-A/Extend_Script_experimentals/blob/master/preferencesKeeper.jsx

### 更新履歴 / Change Log：
- v1.0 (20240801): 初期バージョン / Initial version
*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.0";

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
        ja: "オンスクリーン",
        en: "Onscreen"
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
        en: "Text"
    },
    asianUnit: {
        ja: "東アジア言語のオプション",
        en: "East Asian"
    },
    // -----------------------------------------------
    textTitle: {
        ja: "テキスト",
        en: "Text"
    },
    fontEnglish: {
        ja: "フォント名を英語表記",
        en: "Show font names in English"
    },
    recentFonts: {
        ja: "最近使用したフォント",
        en: "Recent fonts"
    },
    recentFontsCount: {
        ja: "件",
        en: "items"
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
        en: "Transform Patterns"
    },
    scaleCorners: {
        ja: "角を拡大・縮小",
        en: "Scale Corners"
    },
    scaleStroke: {
        ja: "線幅を拡大・縮小",
        en: "Scale Strokes"
    },
    scaleEffects: {
        ja: "効果を拡大・縮小",
        en: "Scale Effects"
    },
    realtimeDrawing: {
        ja: "リアルタイムに描画と編集",
        en: "Real-time Drawing & Editing"
    },
    glyphBounds: {
        ja: "字形の境界に整列",
        en: "Align to Glyph Bounds"
    },
    pointText: {
        ja: "ポイント文字",
        en: "Point Text"
    },
    areaText: {
        ja: "エリア内文字",
        en: "Area Text"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

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

    /* 右カラム / Right column */
    var rightColumn = mainGroup.add('group');
    rightColumn.orientation = 'column';
    rightColumn.alignChildren = ['fill', 'top'];

    /* ダイアログ位置と透明度の調整 / Adjust dialog position and opacity */
    var offsetX = 300;
    var dialogOpacity = 0.97;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    var glyphPanel = rightColumn.add('panel', undefined, LABELS.glyphBounds[lang]);
    glyphPanel.orientation = 'column';
    glyphPanel.alignChildren = ['left', 'top'];
    glyphPanel.margins = [15, 20, 15, 15];

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


    /* 単位パネルを追加 / Add "Units" panel */
    var unitsPanel = leftColumn.add('panel', undefined, LABELS.unitsTitle[lang]);
    unitsPanel.orientation = 'column';
    unitsPanel.alignChildren = ['left', 'top'];
    unitsPanel.margins = [15, 20, 15, 15];


    /* テキストパネルを追加 / Add "Text" panel */
    var textPanel = leftColumn.add('panel', undefined, LABELS.textTitle[lang]);
    textPanel.orientation = 'column';
    textPanel.alignChildren = ['left', 'top'];
    textPanel.margins = [15, 20, 15, 15];

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

    var staticLabel = groupRecentFonts.add('statictext', undefined, LABELS.recentFontsCount[lang]);

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


    var unitDropdowns = {};

    function createUnitDropdown(parent, label, prefKey) {
        var group = parent.add('group');
        group.orientation = 'row';
        var labelControl = group.add('statictext', undefined, label + "：");
        labelControl.characters = 14; /* 同一幅に設定 */
        labelControl.justify = 'right'; /* 右揃え */

        var dropdown = group.add('dropdownlist', undefined, []);
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
        } else if (mode === "printQ") {
            unitDropdowns["rulerType"].selection = unitDropdowns["rulerType"].find("mm");
            unitDropdowns["strokeUnits"].selection = unitDropdowns["strokeUnits"].find("mm");
            unitDropdowns["text/units"].selection = unitDropdowns["text/units"].find("Q/H");
            unitDropdowns["text/asianunits"].selection = unitDropdowns["text/asianunits"].find("Q/H");
        } else if (mode === "onscreen") {
            unitDropdowns["rulerType"].selection = unitDropdowns["rulerType"].find("px");
            unitDropdowns["strokeUnits"].selection = unitDropdowns["strokeUnits"].find("px");
            unitDropdowns["text/units"].selection = unitDropdowns["text/units"].find("px");
            unitDropdowns["text/asianunits"].selection = unitDropdowns["text/asianunits"].find("px");
        }
        /* Trigger onChange manually to update preferences */
        for (var key in unitDropdowns) {
            if (unitDropdowns[key].selection) {
                unitDropdowns[key].onChange();
            }
        }
    }

    radioPrintPt.onClick = function() {
        setUnitsForMode("printPt");
    };
    radioPrintQ.onClick = function() {
        setUnitsForMode("printQ");
    };
    radioOnscreen.onClick = function() {
        setUnitsForMode("onscreen");
    };

    /* その他パネルを追加 / Add "Transform & Align" panel */
    var otherPanel = rightColumn.add('panel', undefined, LABELS.transformTitle[lang]);
    otherPanel.orientation = 'column';
    otherPanel.alignChildren = ['left', 'top'];
    otherPanel.margins = [15, 20, 15, 15];

    var checkboxPreview = otherPanel.add('checkbox', undefined, LABELS.previewBounds[lang]);
    checkboxPreview.value = app.preferences.getBooleanPreference("includeStrokeInBounds");
    checkboxPreview.onClick = function() {
        app.preferences.setBooleanPreference("includeStrokeInBounds", checkboxPreview.value === true);
    };

    var checkboxPattern = otherPanel.add('checkbox', undefined, LABELS.transformPattern[lang]);
    checkboxPattern.value = app.preferences.getBooleanPreference("transformPatterns");
    checkboxPattern.onClick = function() {
        app.preferences.setBooleanPreference("transformPatterns", checkboxPattern.value === true);
    };

    var checkboxCorner = otherPanel.add('checkbox', undefined, LABELS.scaleCorners[lang]);
    checkboxCorner.value = app.preferences.getBooleanPreference("scaleCorners");
    checkboxCorner.onClick = function() {
        app.preferences.setBooleanPreference("scaleCorners", checkboxCorner.value === true);
    };

    /* 線幅を拡大・縮小 */
    var checkboxStroke = otherPanel.add('checkbox', undefined, LABELS.scaleStroke[lang]);
    checkboxStroke.value = app.preferences.getBooleanPreference("scaleLineWeight");
    checkboxStroke.onClick = function() {
        app.preferences.setBooleanPreference("scaleLineWeight", checkboxStroke.value === true);
    };

    /* 効果を拡大・縮小 */
    var checkboxEffect = otherPanel.add('checkbox', undefined, LABELS.scaleEffects[lang]);
    checkboxEffect.value = app.preferences.getBooleanPreference("includeStrokeInBounds");
    checkboxEffect.onClick = function() {
        app.preferences.setBooleanPreference("includeStrokeInBounds", checkboxEffect.value === true);
    };

    var checkboxRealtime = otherPanel.add('checkbox', undefined, LABELS.realtimeDrawing[lang]);
    checkboxRealtime.value = false;
    checkboxRealtime.onClick = function() {
        app.preferences.setBooleanPreference("realTimeDrawing", checkboxRealtime.value === true);
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
    cancelBtn.onClick = function() {
        dialog.close();
    };

    var okBtn = group2.add('button', undefined, LABELS.ok[lang], {
        name: 'ok'
    });
    okBtn.preferredSize.width = 90;
    okBtn.onClick = function() {
        dialog.close();
    };

    dialog.show();
}

main();