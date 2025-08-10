#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name：
PreferenceManagerForTransformAndAlign.jsx

### 概要 / Overview：
- Illustrator の各種環境設定をダイアログボックスから変更可能にします。
- 単位、文字設定、変形／整列設定などをワンパネルで調整できます。
- Allows changing various Illustrator preferences via a dialog box.
- Units, text settings, transform/align settings can be adjusted in one panel.

### 主な機能 / Main Features：
- 最近使用したフォント数とフォント表記切替 / Recent fonts count and font name localization
- 変形と整列の各種オプション / Various transform and align options
- 字形の境界に整列の設定 / Align to glyph bounds setting

### 更新履歴 / Change Log：
- v1.0 (20250804): 初期バージョン / Initial version
- v1.1 (20250804): ダイアログを2カラムに改修、単位とフォント設定を追加 / Dialog changed to two columns; added units and font settings
- v1.2 (20250804): 角の拡大のロジックを修正 / Fixed logic for corner scaling
- v1.3 (20250805): 単位と増減値のUIとロジックを削除 / Removed unit and increment UI and logic
*/


/* スクリプトバージョン / Script Version */
var SCRIPT_VERSION = "v1.3";

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
    }
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


    /* ダイアログ位置と透明度の調整 / Adjust dialog position and opacity */
    var offsetX = 300;
    var dialogOpacity = 0.97;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);


    /* 2カラムのメインコンテナ / Two-column main container */
    var mainGroup = dialog.add('group');
    mainGroup.orientation = 'row';
    mainGroup.alignChildren = ['fill', 'top'];


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

    /* 最近使用したフォントの表示数 ロジック（UI調整） / Recent fonts count logic (UI adjustment) */
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
                inputRecentFonts.text = "15";
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


    /* 字形の境界に整列パネル / Align to glyph bounds panel */
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



    /* 変形と整列パネル / Transform & Align panel */
    var otherPanel = rightColumn.add('panel', undefined, LABELS.transformTitle[lang]);
    otherPanel.orientation = 'column';
    otherPanel.alignChildren = ['left', 'top'];
    otherPanel.margins = [8, 20, 8, 15];

    /* プレビュー境界 / Preview bounds */
    var checkboxPreview = otherPanel.add('checkbox', undefined, LABELS.previewBounds[lang]);
    checkboxPreview.value = app.preferences.getBooleanPreference("includeStrokeInBounds");
    checkboxPreview.onClick = function() {
        app.preferences.setBooleanPreference("includeStrokeInBounds", checkboxPreview.value === true);
    };


    /* パターンを変形 / Transform pattern tiles */
    var checkboxPattern = otherPanel.add('checkbox', undefined, LABELS.transformPattern[lang]);
    checkboxPattern.value = app.preferences.getBooleanPreference("transformPatterns");
    checkboxPattern.onClick = function() {
        app.preferences.setBooleanPreference("transformPatterns", checkboxPattern.value === true);
    };

    /* 角を拡大・縮小 / Scale corners */
    var checkboxCorner = otherPanel.add('checkbox', undefined, LABELS.scaleCorners[lang]);
    /* 初期値を取得（1=ON, 2=OFF） / Get initial value (1=ON, 2=OFF) */
    checkboxCorner.value = (app.preferences.getIntegerPreference("policyForPreservingCorners") === 1);
    checkboxCorner.onClick = function() {
        app.preferences.setIntegerPreference(
            "policyForPreservingCorners",
            checkboxCorner.value ? 1 : 2
        );
    };

    /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
    var checkboxStroke = otherPanel.add('checkbox', undefined, LABELS.scaleStroke[lang]);
    checkboxStroke.value = app.preferences.getBooleanPreference("scaleLineWeight");
    checkboxStroke.onClick = function() {
        app.preferences.setBooleanPreference("scaleLineWeight", checkboxStroke.value === true);
    };

    /* リアルタイムの描画と編集 / Real-time drawing and editing */
    var checkboxRealtime = otherPanel.add('checkbox', undefined, LABELS.realtimeDrawing[lang]);
    checkboxRealtime.value = app.preferences.getBooleanPreference("LiveEdit_State_Machine");
    checkboxRealtime.onClick = function() {
        app.preferences.setBooleanPreference("LiveEdit_State_Machine", checkboxRealtime.value === true);
    };


    var group2 = dialog.add('group', undefined, {
        name: 'group2'
    });
    group2.orientation = 'row';
    group2.alignChildren = ['center', 'center']; /* 中央揃え / Center align */
    group2.alignment = ['center', 'bottom']; /* ダイアログ内で中央に配置 / Center in dialog */

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