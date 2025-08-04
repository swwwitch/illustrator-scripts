#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartVerticalAlign

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartVerticalAlign.jsx

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
- v1.1 (20250804) : ダイアログボックスを開くときのロジックを調整

*/

/*
### Script Name:

SmartVerticalAlign

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartVerticalAlign.jsx

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
- v1.1 (20250804): Adjusted dialog opening logic

*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.1";

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
    glyphBounds: {
        en: 'Align to Glyph Bounds',
        ja: '字形の境界に整列'
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

// テキストや図形などを対象とするか判定 / Check if object is a target
function isTargetObject(obj) {
    if (!obj) return false;
    return (obj.typename === "TextFrame" ||
            obj.typename === "PathItem" ||
            obj.typename === "GroupItem" ||
            obj.typename === "CompoundPathItem");
}

// 選択中の対象オブジェクトを配列で取得 / Get all selected target objects
function getSelectedObjects() {
    var selection = app.activeDocument.selection;
    var targets = [];
    for (var i = 0; i < selection.length; i++) {
        if (isTargetObject(selection[i])) {
            targets.push(selection[i]);
        }
    }
    return targets;
}

// 複数チェックボックスを環境設定キーにバインド / Bind multiple checkboxes to preference keys
function bindCheckboxes(pairs) {
    for (var i = 0; i < pairs.length; i++) {
        (function(pair) {
            pair.checkbox.onClick = function() {
                app.preferences.setBooleanPreference(pair.prefKey, pair.checkbox.value === true);
                applyPreviewAlignment(); // 共通で呼び出す
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

    var glyphPanel = dialog.add('panel', undefined, LABELS.glyphBounds[lang]);
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

    // まず選択オブジェクトを取得 / Get selected objects first
    var frames = getSelectedObjects();

    // 選択中のテキストタイプを判定し、該当しないチェックボックスを無効化 / Disable irrelevant checkboxes based on selected text type
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

    // 選択内容に応じたデフォルト設定 / Set defaults based on selection
    var hasText = false;
    var hasShape = false;
    for (var i = 0; i < frames.length; i++) {
        if (frames[i].typename === "TextFrame") {
            hasText = true;
        } else {
            hasShape = true;
        }
    }

    if (hasText && hasShape) {
        // ［A］テキストと図形の場合
        checkboxPoint.value = true;
        radioCenter.value = true;
        radioTop.value = false;
        radioBottom.value = false;
    } else if (hasText && !hasShape) {
        // ［B］テキストのみ
        checkboxPoint.value = false;
        radioBottom.value = true;
        radioTop.value = false;
        radioCenter.value = false;
    } else {
        // その他は中央に
        radioCenter.value = true;
        radioTop.value = false;
        radioBottom.value = false;
    }

    // キー入力でラジオボタンを選択 / Select radio buttons with key input
    function addAlignKeyHandler(dlg, radios) {
        dlg.addEventListener("keydown", function(event) {
            var keyMap = { T: 0, M: 1, B: 2 };
            if (keyMap[event.keyName] != null) {
                for (var i = 0; i < radios.length; i++) {
                    radios[i].value = (i === keyMap[event.keyName]);
                }
                applyPreviewAlignment();
                event.preventDefault();
            }
        });
    }

    addAlignKeyHandler(dialog, [radioTop, radioCenter, radioBottom]);

    // (removed: デフォルトで「中央」を選択 / Default selection is "中央")

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

    // 整列プレビュー処理 / Apply alignment preview immediately
    function applyPreviewAlignment() {
        var frames = getSelectedObjects();

        var currentAlign = radioTop.value ? "Top" :
            radioCenter.value ? "Center" :
            radioBottom.value ? "Bottom" : null;

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
    // group2.spacing = 10;
    // group2.margins = [15, 15, 15, 15];

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