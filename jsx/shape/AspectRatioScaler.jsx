#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

AdjustArtboardByRatio.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/AspectRatioScaler.jsx

### 概要：

- 選択オブジェクトを比率に基づいて変形
- プレビュー機能を備えたダイアログボックスUIに対応

### 主な機能：

- 16:9、1:1、カスタム比率から選択可能
- 「縦」または「横」基準で比率調整
- ピクセルグリッド整合オプション
- アートボードに変換オプション付き
- プレビュー反映と数値の矢印キー操作による調整

### 処理の流れ：

- ダイアログを表示し、比率・基準・整合オプションを選択
- プレビューで対象をリアルタイムにシミュレーション
- OKボタンで処理確定、アートボード作成（または調整）

### note：

https://note.com/dtp_tranist/n/n4a212e6eacf1

### 更新履歴：

- v1.0 (20250720) : 初期バージョン
- v1.1 (20250721) : アートボード変換・カスタム比率機能を追加
- v1.2 (20250722) : ダイアログ構成改善・ローカライズ・キー操作追加

*/

/*

### Script Name：
AdjustArtboardByRatio.jsx

### Readme (GitHub)：
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/AspectRatioScaler.jsx

### Description：

- Transforms selected objects based on aspect ratio
- Supports a preview-enabled dialog box UI

### Main Features：

- Choose from 16:9, 1:1, or custom ratio
- Ratio adjustment based on "height" or "width"
- Option to align to pixel grid
- Convert to artboard optionally
- Real-time preview and value adjustment via arrow keys

### Workflow：

- Display dialog to select ratio, basis, and options
- Simulate resizing in real time as preview
- Apply final settings by pressing OK

### Update History：

- v1.0 (20250720): Initial release
- v1.1 (20250721): Added artboard conversion & custom ratio
- v1.2 (20250722): Improved dialog structure, localization, and key input

*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.2";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "アスペクト比で調整 " + SCRIPT_VERSION,
        en: "Adjust by Aspect Ratio" + SCRIPT_VERSION
    },
    ratio169: {
        ja: "16:9",
        en: "16:9"
    },
    ratio11: {
        ja: "1:1（スクエア）",
        en: "1:1"
    },
    ratioA4: {
        ja: "A4（1:1.414）",
        en: "1:1.414"
    },
    ratioCustom: {
        ja: "カスタム",
        en: "Custom"
    },
    widthRatio: {
        ja: "幅比率",
        en: "Width Ratio"
    },
    heightRatio: {
        ja: "高さ比率",
        en: "Height Ratio"
    },
    baseLabel: {
        ja: "基準",
        en: "Base"
    },
    baseHeight: {
        ja: "高さ",
        en: "Height"
    },
    baseWidth: {
        ja: "幅",
        en: "Width"
    },
    alignToPixelGrid: {
        ja: "ピクセルグリッドに最適化",
        en: "Align to Pixel Grid"
    },
    convertToArtboard: {
        ja: "アートボードに変換",
        en: "Convert to Artboard"
    },
    aspectLabel: {
        ja: "アスペクト比",
        en: "Aspect Ratio"
    },
    previewError: {
        ja: "プレビュー更新エラー：",
        en: "Preview Update Error:"
    },
    run: {
        ja: "実行",
        en: "Apply"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

// Declare global UI variables
var baseWidthRadio, baseHeightRadio;
var ratio169, ratio11, ratioA4, ratioCustom;
var editTextWidth, editTextHeight;

// ラジオボタンの水平方向・垂直方向切り替えキーハンドラ追加
function addDirectionKeyHandler(dialog, horizontalRadio, verticalRadio) {
    dialog.addEventListener("keydown", function(event) {
        if (event.keyName == "H") {
            horizontalRadio.value = true;
            verticalRadio.value = false;
            event.preventDefault();
        } else if (event.keyName == "V") {
            horizontalRadio.value = false;
            verticalRadio.value = true;
            event.preventDefault();
        }
    });
}

/* プレビュー更新 / Update preview */
function updatePreview() {
    try {
        var selection = app.activeDocument.selection;
        if (!selection || selection.length === 0) return;

        var isWidthBase = baseWidthRadio.value;
        var ratioW = parseFloat(editTextWidth.text);
        var ratioH = parseFloat(editTextHeight.text);
        if (isNaN(ratioW) || isNaN(ratioH) || ratioW <= 0 || ratioH <= 0) return;

        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];

            var originalWidth = item.width;
            var originalHeight = item.height;

            if (ratioCustom.value) {
                if (isWidthBase) {
                    var newHeight = originalWidth * (ratioH / ratioW);
                    item.height = newHeight;
                } else {
                    var newWidth = originalHeight * (ratioW / ratioH);
                    item.width = newWidth;
                }
            } else if (ratio169.value) {
                if (isWidthBase) {
                    item.height = item.width / 1.777777;
                } else {
                    item.width = item.height * 1.777777;
                }
            } else if (ratio11.value) {
                var size = isWidthBase ? item.width : item.height;
                item.width = size;
                item.height = size;
            } else if (ratioA4.value) {
                if (isWidthBase) {
                    item.height = item.width * (297 / 210); // A4 aspect ratio: height = width * 1.414285
                } else {
                    item.width = item.height * (210 / 297); // A4 aspect ratio: width = height * 0.70707
                }
            }

            app.executeMenuCommand('Make Pixel Perfect');
        }
        app.redraw();
    } catch (e) {
        alert(LABELS.previewError[lang] + "\n" + e);
    }
}

/* ダイアログ作成 / Create dialog */
function createDialog() {
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

    var offsetX = 300;
    var dialogOpacity = 0.97;
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);
    dialog.alignChildren = "left";

    var topGroup = dialog.add("group");
    topGroup.orientation = "row";
    topGroup.alignChildren = "left";
    topGroup.alignChildren = ["fill", "top"];

    var aspectPanel = topGroup.add("panel", undefined, LABELS.aspectLabel[lang]);
    aspectPanel.orientation = "column";
    aspectPanel.alignChildren = "left";
    aspectPanel.margins = [15, 20, 15, 10];
    aspectPanel.alignment = ["fill", "top"];

    var aspectGroup = aspectPanel.add("group");
    aspectGroup.orientation = "column";
    aspectGroup.alignChildren = "left";
    ratio169 = aspectGroup.add("radiobutton", undefined, LABELS.ratio169[lang]);
    ratio11 = aspectGroup.add("radiobutton", undefined, LABELS.ratio11[lang]);
    ratioA4 = aspectGroup.add("radiobutton", undefined, LABELS.ratioA4[lang]);
    ratioCustom = aspectGroup.add("radiobutton", undefined, LABELS.ratioCustom[lang]);

    // Add event listener for A4 aspect ratio button
    ratioA4.onClick = function() {
        updatePreview();
    };

    var customRatioGroup = aspectPanel.add("group");
    customRatioGroup.orientation = "row";
    customRatioGroup.alignChildren = "left";

    // customRatioGroup.add("statictext", undefined, LABELS.widthRatio[lang]);
    editTextWidth = customRatioGroup.add("edittext", undefined, "3");
    editTextWidth.characters = 3;

    customRatioGroup.add("statictext", undefined, ":");
    editTextHeight = customRatioGroup.add("edittext", undefined, "2");
    editTextHeight.characters = 3;

    editTextWidth.enabled = false;
    editTextHeight.enabled = false;

    changeValueByArrowKey(editTextWidth);
    changeValueByArrowKey(editTextHeight);

    ratio169.value = true;

    var basePanel = topGroup.add("panel", undefined, LABELS.baseLabel[lang]);
    basePanel.orientation = "column";
    basePanel.alignChildren = "left";
    basePanel.margins = [15, 20, 15, 10];

    var baseGroup = basePanel.add("group");
    baseGroup.orientation = "column";
    baseGroup.alignChildren = "left";
    baseWidthRadio = baseGroup.add("radiobutton", undefined, LABELS.baseWidth[lang]);
    baseHeightRadio = baseGroup.add("radiobutton", undefined, LABELS.baseHeight[lang]);
    baseHeightRadio.value = true;

    var pixelGroup = dialog.add("group");
    var alignToPixel = pixelGroup.add("checkbox", undefined, LABELS.alignToPixelGrid[lang]);
    alignToPixel.value = true;

    var convertToArtboard = dialog.add("checkbox", undefined, LABELS.convertToArtboard[lang]);
    convertToArtboard.value = false;

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";

    var btnCancel = buttonGroup.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var btnOk = buttonGroup.add("button", undefined, LABELS.run[lang], {
        name: "ok"
    });

    return {
        dialog: dialog,
        ratio169: ratio169,
        ratio11: ratio11,
        ratioA4: ratioA4,
        ratioCustom: ratioCustom,
        baseVertical: baseHeightRadio,
        baseHorizontal: baseWidthRadio,
        alignToPixel: alignToPixel,
        convertToArtboard: convertToArtboard,
        btnOk: btnOk,
        btnCancel: btnCancel,
        customWidthInput: editTextWidth,
        customHeightInput: editTextHeight
    };
}

/* メイン処理 / Main function */
function main() {
    /* 選択したオブジェクトを取得する / Get selected objects */
    var selectedItems = app.activeDocument.selection;
    if (!selectedItems || selectedItems.length === 0) {
        alert(LABELS.cancel ? LABELS.cancel[lang] : "オブジェクトを選択してください。");
        return;
    }

    // プレビュー用コピーを作成し、元は非表示にする / Create preview copies and hide originals
    var previewCopies = [];
    var originalWidths = [];
    var originalHeights = [];
    for (var i = 0; i < selectedItems.length; i++) {
        var dup = selectedItems[i].duplicate();
        dup.hidden = false;
        dup.zOrder(ZOrderMethod.BRINGTOFRONT);
        previewCopies.push(dup);
        selectedItems[i].hidden = true;
        originalWidths.push(selectedItems[i].width); // 幅を保存 / Save width
        originalHeights.push(selectedItems[i].height); // 高さを保存 / Save height
    }

    var dialogResult = createDialog();

    // ラジオボタンのキーボード操作を有効化
    addDirectionKeyHandler(dialogResult.dialog, dialogResult.baseHorizontal, dialogResult.baseVertical);

    /* アスペクト比選択時のプレビュー更新 / Preview update on aspect ratio selection */
    dialogResult.ratio169.onClick = function() {
        dialogResult.customWidthInput.enabled = false;
        dialogResult.customHeightInput.enabled = false;
        applyAspect(previewCopies, 1.777777, dialogResult.baseVertical.value, originalWidths, originalHeights);
    };
    dialogResult.ratio11.onClick = function() {
        dialogResult.customWidthInput.enabled = false;
        dialogResult.customHeightInput.enabled = false;
        applyAspect(previewCopies, 1.0, dialogResult.baseVertical.value, originalWidths, originalHeights);
    };
    dialogResult.ratioCustom.onClick = function() {
        dialogResult.customWidthInput.enabled = true;
        dialogResult.customHeightInput.enabled = true;
        if (dialogResult.ratioCustom.value) {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            var r = (h === 0) ? 1 : w / h;
            applyAspect(previewCopies, r, dialogResult.baseVertical.value, originalWidths, originalHeights);
        }
    };

    /* カスタム比率入力時のプレビュー更新 / Preview update on custom ratio input */
    dialogResult.customWidthInput.onChanging = function() {
        if (dialogResult.ratioCustom.value) {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            var r = (h === 0) ? 1 : w / h;
            applyAspect(previewCopies, r, dialogResult.baseVertical.value, originalWidths, originalHeights);
        }
    };
    dialogResult.customHeightInput.onChanging = function() {
        if (dialogResult.ratioCustom.value) {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            var r = (h === 0) ? 1 : w / h;
            applyAspect(previewCopies, r, dialogResult.baseVertical.value, originalWidths, originalHeights);
        }
    };

    /* 基準切り替え時のプレビュー更新 / Preview update on base switch */
    dialogResult.baseVertical.onClick = function() {
        var ratio = dialogResult.ratio169.value ? 1.777777 :
            dialogResult.ratio11.value ? 1.0 :
            dialogResult.ratioA4.value ? (210 / 297) :
            (function() {
                var w = parseFloat(dialogResult.customWidthInput.text);
                var h = parseFloat(dialogResult.customHeightInput.text);
                return (h === 0) ? 1 : w / h;
            })();
        applyAspect(previewCopies, ratio, true, originalWidths, originalHeights);
    };
    dialogResult.baseHorizontal.onClick = function() {
        var ratio = dialogResult.ratio169.value ? 1.777777 :
            dialogResult.ratio11.value ? 1.0 :
            dialogResult.ratioA4.value ? (210 / 297) :
            (function() {
                var w = parseFloat(dialogResult.customWidthInput.text);
                var h = parseFloat(dialogResult.customHeightInput.text);
                return (h === 0) ? 1 : w / h;
            })();
        applyAspect(previewCopies, ratio, false, originalWidths, originalHeights);
    };

    /* 初期プレビュー / Initial preview */
    var initialRatio = dialogResult.ratio169.value ? 1.777777 : (dialogResult.ratio11.value ? 1.0 : (function() {
        var w = parseFloat(dialogResult.customWidthInput.text);
        var h = parseFloat(dialogResult.customHeightInput.text);
        return (h === 0) ? 1 : w / h;
    })());
    applyAspect(previewCopies, initialRatio, dialogResult.baseVertical.value, originalWidths, originalHeights);

    var result = dialogResult.dialog.show();

    /* OKボタンが押された場合 / If OK pressed */
    if (result === 1) {
        var ratio = dialogResult.ratio169.value ? 1.777777 : (dialogResult.ratio11.value ? 1.0 : (function() {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            return (h === 0) ? 1 : w / h;
        })());
        var isVertical = dialogResult.baseVertical.value;
        for (var i = 0; i < selectedItems.length; i++) {
            selectedItems[i].hidden = false;
            if (isVertical) {
                var h = originalHeights[i];
                selectedItems[i].width = h * ratio;
            } else {
                var w = originalWidths[i];
                selectedItems[i].height = w / ratio;
            }
            if (dialogResult.alignToPixel.value) {
                app.selection = [selectedItems[i]];
                app.executeMenuCommand('Make Pixel Perfect');
            }
            previewCopies[i].remove();
        }
        /* アートボードに変換 / Convert to artboard */
        if (dialogResult.convertToArtboard.value) {
            for (var i = 0; i < selectedItems.length; i++) {
                var item = selectedItems[i];
                var vb = item.visibleBounds; // [y1, x1, y2, x2]
                var abRect = [vb[0], vb[1], vb[2], vb[3]];
                app.activeDocument.artboards.add(abRect);
            }
        }
        app.selection = selectedItems;
    }
    /* キャンセル時は元に戻す / Restore on cancel */
    else {
        for (var i = 0; i < previewCopies.length; i++) {
            previewCopies[i].remove();
        }
        for (var i = 0; i < selectedItems.length; i++) {
            selectedItems[i].hidden = false;
        }
        return;
    }
}

/* アスペクト比適用 / Apply aspect ratio */
function applyAspect(items, ratio, isVertical, originalWidths, originalHeights) {
    for (var i = 0; i < items.length; i++) {
        if (isVertical) {
            items[i].height = originalHeights[i];
            items[i].width = originalHeights[i] * ratio;
        } else {
            items[i].width = originalWidths[i];
            items[i].height = originalWidths[i] / ratio;
        }
    }
    app.redraw();
}

/* 上下キーで数値変更を可能にする / Enable arrow key numeric input */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta;

        delta = 1;
        if (event.keyName == "Up") {
            value += delta;
            event.preventDefault();
        } else if (event.keyName == "Down") {
            value -= delta;
            if (value < 0) value = 0;
            event.preventDefault();
        }
        // 整数に丸め / Round to integer
        value = Math.round(value);

        editText.text = value;
        updatePreview(); // Reflect the value change in preview
    });
}

main();