#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartVerticalAlign

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- ポイント文字およびエリア内文字に対して［字形の境界に整列］を制御するスクリプト
- 整列位置（上・中央・下）をプレビューで即時確認可能

### 主な機能：

- ポイント文字／エリア内文字に対応
- 「プレビュー境界」のON/OFFで境界基準を切替
- キー入力（T/M/B）で整列を操作可能
- ダイアログ位置・透明度の調整機能

### 処理の流れ：

1. 選択中のオブジェクトからTextFrameを検出
2. チェックボックスで字形境界整列のON/OFFを設定
3. ラジオボタンまたはキー入力で整列方向を指定
4. プレビュー境界ON/OFFに応じてジオメトリ境界／プレビュー境界を適用
5. プレビューで即時に反映

### 更新履歴：

- v1.0 (20250804) : 初期バージョン

*/

/*
### Script Name:

SmartVerticalAlign

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Controls "Align to Glyph Bounds" for Point Text and Area Text
- Provides instant preview of alignment (Top / Center / Bottom)

### Key Features:

- Supports Point Text and Area Text
- Toggles between Preview Bounds and Geometric Bounds
- Allows T/M/B keyboard shortcuts for alignment
- Adjusts dialog position and opacity

### Workflow:

1. Detect TextFrames from current selection
2. Configure Glyph Bounds alignment via checkboxes
3. Specify alignment direction with radio buttons or keyboard shortcuts
4. Apply Geometric or Preview Bounds based on checkbox
5. Reflect instantly with preview

### Update History:

- v1.0 (20250804): Initial version

*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.0";

// 現在のUI言語を取得 / Get the current UI language
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

// UIラベル定義 / UI Label Definitions
var LABELS = {
    dialogTitle: {
        ja: "垂直方向の整列 " + SCRIPT_VERSION,
        en: "Vertical Alignment " + SCRIPT_VERSION
    },
    title: {
        en: 'Vertical Alignment',
        ja: '垂直方向の整列'
    },
    pointText: {
        en: 'Point Text',
        ja: 'ポイント文字'
    },
    areaText: {
        en: 'Area Text',
        ja: 'エリア内文字'
    },
    alignTitle: {
        en: 'Alignment',
        ja: '整列'
    },
    radioTop: {
        en: 'Top',
        ja: '上'
    },
    radioCenter: {
        en: 'Center',
        ja: '中央'
    },
    radioBottom: {
        en: 'Bottom',
        ja: '下'
    },
    previewBounds: {
        en: 'Preview Bounds',
        ja: 'プレビュー境界'
    },
    ok: {
        en: 'OK',
        ja: 'OK'
    },
    cancel: {
        en: 'Cancel',
        ja: 'キャンセル'
    }
}

// TextFrameかどうかを判定する / Check if object is a TextFrame
function isTextFrame(obj) {
    return obj && obj.typename === "TextFrame";
}

// 選択中のTextFrameを配列で取得 / Get all selected TextFrames
function getSelectedTextFrames() {
    var selection = app.activeDocument.selection;
    var frames = [];
    for (var i = 0; i < selection.length; i++) {
        if (isTextFrame(selection[i])) {
            frames.push(selection[i]);
        }
    }
    return frames;
}

// 複数チェックボックスを環境設定キーにバインド / Bind multiple checkboxes to preference keys
function bindCheckboxes(pairs) {
    for (var i = 0; i < pairs.length; i++) {
        (function(pair) {
            pair.checkbox.onClick = function() {
                app.preferences.setBooleanPreference(pair.prefKey, pair.checkbox.value === true);
            };
        })(pairs[i]);
    }
}

// メイン処理 / Main entry point
function main() {
    var pointText = 'EnableActualPointTextSpaceAlign';
    var areaText = 'EnableActualAreaTextSpaceAlign';
    var pref = app.preferences;
    var pointGlyph = pref.getBooleanPreference(pointText);
    var areaGlyph = pref.getBooleanPreference(areaText);
    showDialog(pointGlyph, areaGlyph);
}

// ダイアログUIを構築 / Build the dialog UI
function showDialog(pointGlyph, areaGlyph) {
    var pref = app.preferences;
    $.localize = true;
    var ui = LABELS;

    var dialog = new Window('dialog');
    dialog.text = LABELS.dialogTitle[lang];
    dialog.orientation = 'column';
    dialog.alignChildren = ['fill', 'top'];

    // ダイアログ位置と透明度の調整 / Adjust dialog position and opacity
    var offsetX = 300;
    var dialogOpacity = 0.97;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    var glyphPanel = dialog.add('panel', undefined, '字形の境界に整列');
    glyphPanel.orientation = 'column';
    glyphPanel.alignChildren = ['left', 'top'];
    glyphPanel.margins = [15, 20, 15, 15];

    var checkboxPoint = glyphPanel.add('checkbox', undefined, LABELS.pointText[lang]);
    checkboxPoint.value = pointGlyph;

    var checkboxArea = glyphPanel.add('checkbox', undefined, LABELS.areaText[lang]);
    checkboxArea.value = areaGlyph;

    bindCheckboxes([{
            checkbox: checkboxPoint,
            prefKey: 'EnableActualPointTextSpaceAlign'
        },
        {
            checkbox: checkboxArea,
            prefKey: 'EnableActualAreaTextSpaceAlign'
        }
    ]);
    // Override onClick handlers to trigger preview update
    checkboxPoint.onClick = function() {
        app.preferences.setBooleanPreference('EnableActualPointTextSpaceAlign', checkboxPoint.value);
        // 現在の整列を再実行 / Re-apply current alignment
        applyPreviewAlignment();
    };
    checkboxArea.onClick = function() {
        app.preferences.setBooleanPreference('EnableActualAreaTextSpaceAlign', checkboxArea.value);
        // 現在の整列を再実行 / Re-apply current alignment
        applyPreviewAlignment();
    };

    // 選択中のテキストタイプを判定し、該当しないチェックボックスを無効化 / Disable irrelevant checkboxes based on selected text type
    var frames = getSelectedTextFrames();
    if (frames.length > 0) {
        if (frames[0].kind === TextType.POINTTEXT) {
            checkboxArea.enabled = false; // エリア内文字をディム / Dim Area Text
        } else if (frames[0].kind === TextType.AREATEXT) {
            checkboxPoint.enabled = false; // ポイント文字をディム / Dim Point Text
        }
    }

    // 整列位置の選択肢（上・中央・下） / Alignment options (Top, Center, Bottom)
    var alignPanel = dialog.add('panel', undefined, LABELS.alignTitle[lang]);
    alignPanel.orientation = 'column';
    alignPanel.alignChildren = ['left', 'top'];
    alignPanel.margins = [15, 20, 15, 15];
    alignPanel.spacing = 5;

    var radioTop = alignPanel.add('radiobutton', undefined, LABELS.radioTop[lang]);
    var radioCenter = alignPanel.add('radiobutton', undefined, LABELS.radioCenter[lang]);
    var radioBottom = alignPanel.add('radiobutton', undefined, LABELS.radioBottom[lang]);

    // キー入力でラジオボタンを選択 / Select radio buttons with key input
    function addAlignKeyHandler(dlg, topRadio, centerRadio, bottomRadio) {
        dlg.addEventListener("keydown", function(event) {
            if (event.keyName === "T") {
                topRadio.value = true;
                centerRadio.value = false;
                bottomRadio.value = false;
                applyPreviewAlignment();
                event.preventDefault();
            } else if (event.keyName === "M") {
                topRadio.value = false;
                centerRadio.value = true;
                bottomRadio.value = false;
                applyPreviewAlignment();
                event.preventDefault();
            } else if (event.keyName === "B") {
                topRadio.value = false;
                centerRadio.value = false;
                bottomRadio.value = true;
                applyPreviewAlignment();
                event.preventDefault();
            }
        });
    }

    addAlignKeyHandler(dialog, radioTop, radioCenter, radioBottom);

    // デフォルトで「中央」を選択 / Default selection is "中央"
    radioCenter.value = true;

    // プレビュー境界チェックボックスを追加 / Add Preview Bounds checkbox
    var previewGroup = dialog.add('group');
    previewGroup.orientation = 'row';
    previewGroup.alignChildren = ['left', 'center'];
    previewGroup.margins = [15, 0, 15, 0];

    var checkboxPreview = previewGroup.add('checkbox', undefined, LABELS.previewBounds[lang]);
    checkboxPreview.value = false; // デフォルトはOFF / Default OFF
    // プレビュー境界チェックボックスのクリック時にIllustratorの環境設定を切り替え
    checkboxPreview.onClick = function() {
        if (checkboxPreview.value) {
            app.preferences.setBooleanPreference("includeStrokeInBounds", true); // ON時: プレビュー境界を含める
        } else {
            app.preferences.setBooleanPreference("includeStrokeInBounds", false); // OFF時: プレビュー境界を含めない
        }
        applyPreviewAlignment();
    };

    // 前回の整列状態を記録 / Track the last applied alignment
    var lastAlign = null;
    var lastSelectionLength = 0; // 前回の選択数を記録 / Track last selection count

    // 整列プレビュー処理 / Apply alignment preview immediately
    function applyPreviewAlignment() {
        var frames = getSelectedTextFrames();

        // 選択が変わったら lastAlign をリセット / Reset lastAlign if selection changed
        if (frames.length !== lastSelectionLength) {
            lastAlign = null;
            lastSelectionLength = frames.length;
        }

        var currentAlign = radioTop.value ? "Top" :
            radioCenter.value ? "Center" :
            radioBottom.value ? "Bottom" : null;

        lastAlign = currentAlign; // 選択変更時は常に実行 / Always execute on selection change

        if (frames.length > 0) {
            if (radioTop.value) {
                app.executeMenuCommand('Vertical Align Top');
            } else if (radioCenter.value) {
                app.executeMenuCommand('Vertical Align Center');
            } else if (radioBottom.value) {
                app.executeMenuCommand('Vertical Align Bottom');
            }
        }
        app.redraw();
    }

    // プレビュー用イベントリスナー / Add event listeners for preview
    radioTop.onClick = applyPreviewAlignment;
    radioCenter.onClick = applyPreviewAlignment;
    radioBottom.onClick = applyPreviewAlignment;

    var group2 = dialog.add('group', undefined, {
        name: 'group2'
    });
    group2.orientation = 'row';
    group2.alignChildren = ['fill', 'center'];
    group2.alignment = ['right', 'bottom'];
    group2.spacing = 10;
    group2.margins = [15, 15, 15, 15];

    var cancelBtn = group2.add('button', undefined, L("cancel"), {
        name: 'cancel'
    });
    cancelBtn.preferredSize.width = 90;
    cancelBtn.onClick = function() {
        dialog.close();
    };

    var okBtn = group2.add('button', undefined, L("ok"), {
        name: 'ok'
    });
    okBtn.preferredSize.width = 90;
    okBtn.onClick = function() {
        dialog.close();
    };

    dialog.show();
}

main();

// ラベル取得関数 / Label getter for localization
function L(key) {
    return LABELS[key] && LABELS[key][lang] ? LABELS[key][lang] : key;
}