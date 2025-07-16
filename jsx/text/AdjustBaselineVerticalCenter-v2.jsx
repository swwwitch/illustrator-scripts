#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

/*

### スクリプト名：

AdjustBaselineVerticalCenter-v2.jsx

### Readme （GitHub）：

https://github.com/creold/illustrator-scripts

### 概要：

- 指定した文字（複数可）を基準文字に合わせて縦方向（ベースライン）を調整
- 複数テキストフレームに一括適用可能

### 主な機能：

- 複数対象文字の指定と一括調整
- 自動的に最頻出記号を対象文字に設定
- 日本語／英語 UI 対応

### 処理の流れ：

1. ダイアログで文字を指定
2. アウトライン作成し中心Yを取得
3. 差分に基づきベースラインシフトを適用

### オリジナル、謝辞：

Egor Chistyakov https://x.com/tchegr

### 更新履歴：

- v1.0 (20250704): 初版リリース
- v1.6 (20250705): 複数対象文字対応、一括調整対応

---

### Script Name:

AdjustBaselineVerticalCenter-v2.jsx

### Readme (GitHub):

https://github.com/creold/illustrator-scripts

### Overview:

- Adjust vertical (baseline) position of specified characters to match a reference character
- Supports batch application to multiple text frames

### Main Features:

- Supports multiple target characters and batch adjustment
- Automatically detects most frequent symbol as default
- Japanese / English UI support

### Process Flow:

1. Specify characters in dialog
2. Create outlines and get center Y
3. Apply baseline shift based on difference

### Original / Acknowledgements:

Egor Chistyakov https://x.com/tchegr

### Update History:

- v1.0 (20250704): Initial release
- v1.6 (20250705): Support for multiple target characters and batch adjustment

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.6";

var LABELS = {
    dialogTitle: {
        ja: "ベースライン調整 " + SCRIPT_VERSION,
        en: "Adjust Baseline " + SCRIPT_VERSION
    },
    targetCharLabel: {
        ja: "対象文字:",
        en: "Target Character:"
    },
    baseCharLabel: {
        ja: "基準文字:",
        en: "Reference Character:"
    },
    okBtnLabel: {
        ja: "調整",
        en: "Adjust"
    },
    cancelBtnLabel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    selectFrameMsg: {
        ja: "テキストフレームを選択してください。",
        en: "Select one or more text frames."
    },
    docOpenMsg: {
        ja: "ドキュメントが開かれていません。",
        en: "No document open."
    },
    invalidCharMsg: {
        ja: "対象文字は1文字以上、基準文字は1文字を入力してください。",
        en: "Enter at least one target character and exactly one reference character."
    },
    notFoundMsg: {
        ja: "対象文字が含まれていません。",
        en: "Target character not found."
    },
    errorMsg: {
        ja: "エラー: ",
        en: "Error: "
    },
    resetBtnLabel: {
        ja: "リセット",
        en: "Reset"
    },
    shiftAmountLabel: {
        ja: "シフト量:",
        en: "Shift Amount:"
    },
    autoPanelTitle: {
        ja: "自動調整",
        en: "Auto Adjust"
    },
    numericErrorMsg: {
        ja: "シフト量は数値で入力してください。",
        en: "Shift amount must be a number."
    }
};

/* 言語判定 / Determine language from locale */
function getLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

/* 単位コードとラベルのマップ / Unit code to label map */
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/* 現在の単位ラベルを取得 / Get current unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("text/asianunits");
    return unitLabelMap[unitCode] || "pt";
}

/* 単位コードからポイント換算係数を取得 / Get point factor from unit code */
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

/* EditTextで上下キーによる値の増減を実装 / Enable arrow key increment/decrement on edittext */
function changeValueByArrowKey(editText, allowNegative, targetInput, textFrames) {
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
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10; // 小数第1位まで / Round to 1 decimal
        } else {
            value = Math.round(value); // 整数に丸め / Round to integer
        }

        if (!allowNegative && value < 0) value = 0;

        event.preventDefault();
        editText.text = value;
        if (typeof previewShiftAll === "function" && targetInput && textFrames) {
            previewShiftAll(targetInput, editText, textFrames);
        }
    });
}

/* プレビューとして全選択テキストのベースラインシフトを更新 / Preview baseline shift for all selected text */
function previewShiftAll(targetInput, shiftInput, textFrames) {
    var targetText = targetInput.text;
    var shiftValue = parseFloat(shiftInput.text);
    if (isNaN(shiftValue)) shiftValue = 0;

    if (!targetText || isNaN(shiftValue)) {
        resetBaselineShift(textFrames);
        lastTarget = "";
        lastShift = 0;
        app.redraw();
        return;
    }

    resetBaselineShift(textFrames);

    for (var j = 0; j < textFrames.length; j++) {
        var tf = textFrames[j];
        var contents = tf.contents;
        for (var c = 0; c < targetText.length; c++) {
            var ch = targetText.charAt(c);
            var pos = contents.indexOf(ch);
            while (pos !== -1) {
                tf.textRange.characters[pos].characterAttributes.baselineShift = shiftValue;
                pos = contents.indexOf(ch, pos + 1);
            }
        }
    }

    lastTarget = targetText;
    lastShift = shiftValue;
    app.redraw();
}

/* アイテムのジオメトリック境界から中心Y座標を取得 / Get center Y coordinate from geometric bounds */
function getCenterY(item) {
    var bounds = item.geometricBounds;
    return bounds[1] - (bounds[1] - bounds[3]) / 2;
}

/* 指定文字でアウトラインを作成し中心Y座標を取得 / Create outline for character and get center Y */
function createOutlineAndGetCenterY(textFrame, character) {
    var tempFrame = textFrame.duplicate();
    tempFrame.contents = character;
    tempFrame.filled = false;
    tempFrame.stroked = false;
    var outline = tempFrame.createOutline();
    var centerY = getCenterY(outline);
    outline.remove();
    return centerY;
}

/* 選択テキスト内の記号・非英数字の出現頻度を集計 / Count frequency of non-alphanumeric symbols */
function getSymbolFrequency(sel) {
    var charCount = {};
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename == "TextFrame") {
            var selText = item.contents;
            for (var j = 0; j < selText.length; j++) {
                var c = selText.charAt(j);
                if (!c.match(/^[A-Za-z0-9\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF]$/)) {
                    charCount[c] = (charCount[c] || 0) + 1;
                }
            }
        }
    }
    return charCount;
}

/* 全テキストフレームのベースラインシフトをリセット / Reset baseline shift for all characters */
function resetBaselineShift(textFrames) {
    for (var i = 0; i < textFrames.length; i++) {
        var tf = textFrames[i];
        var chars = tf.textRange.characters;
        for (var j = 0; j < chars.length; j++) {
            chars[j].characterAttributes.baselineShift = 0;
        }
    }
}

/* 対象文字と基準文字を入力するダイアログを表示 / Show dialog to input target and reference characters */
function showDialog() {
    var lang = getLang();
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];

    /* デフォルト対象文字を最頻出記号から抽出 / Default target from most frequent symbol */
    var defaultTarget = "";
    var sel = app.activeDocument.selection;
    if (sel && sel.length > 0) {
        var charCount = getSymbolFrequency(sel);
        var maxCount = 0;
        for (var key in charCount) {
            if (charCount[key] > maxCount) {
                maxCount = charCount[key];
                defaultTarget = key;
            }
        }
    }

    var inputGroup = mainGroup.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";
    inputGroup.margins = [15, 5, 15, 5];

    var targetGroup = inputGroup.add("group");
    targetGroup.add("statictext", undefined, LABELS.targetCharLabel[lang]);
    var targetInput = targetGroup.add("edittext", undefined, defaultTarget);
    targetInput.characters = 6;
    targetInput.active = true;

    var shiftGroup = inputGroup.add("group");
    shiftGroup.add("statictext", undefined, LABELS.shiftAmountLabel[lang]);
    var shiftInput = shiftGroup.add("edittext", undefined, "0");
    var unitLabel = shiftGroup.add("statictext", undefined, getCurrentUnitLabel());
    shiftInput.characters = 6;
    changeValueByArrowKey(shiftInput, true, targetInput, app.activeDocument.selection);
    shiftInput.onChanging = function() {
        previewShiftAll(targetInput, shiftInput, app.activeDocument.selection);
    };
    shiftInput.active = true;
    shiftGroup.margins = [0, 0, 0, 10];

    var autoPanel = inputGroup.add("panel", undefined, LABELS.autoPanelTitle[lang]);
    autoPanel.orientation = "column";
    autoPanel.alignChildren = "left";
    autoPanel.margins = [15, 20, 15, 5];

    var refGroup = autoPanel.add("group");
    refGroup.add("statictext", undefined, LABELS.baseCharLabel[lang]);
    var refInput = refGroup.add("edittext", undefined, "0");
    refInput.characters = 3;
    var okBtn = refGroup.add("button", [0, 0, 60, 25], LABELS.okBtnLabel[lang]);

    var buttonGroup = mainGroup.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.alignChildren = "fill";

    var finalOkBtn = buttonGroup.add("button", undefined, "OK", {
        name: "ok"
    });
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancelBtnLabel[lang]);

    buttonGroup.add("statictext", [0, 0, 0, 20], " "); // Spacer
    var resetBtn = buttonGroup.add("button", undefined, LABELS.resetBtnLabel[lang]);

    resetBtn.onClick = function() {
        var selection = app.activeDocument.selection;
        if (!selection || selection.length == 0) {
            alert(LABELS.selectFrameMsg[lang]);
            return;
        }
        resetBaselineShift(selection);
        shiftInput.text = "0";
        app.redraw();
    };

    /* 自動調整ボタンのクリック処理 / Auto adjust button click */
    okBtn.onClick = function() {
        if (targetInput.text.length == 0) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        if (refInput.text.length != 1) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        if (!app.documents.length) {
            alert(LABELS.docOpenMsg[lang]);
            return;
        }
        var selection = app.activeDocument.selection;
        if (!selection || selection.length == 0) {
            alert(LABELS.selectFrameMsg[lang]);
            return;
        }
        // 最初の該当文字で基準文字とのY座標差分を計算し、シフト量を設定 / Calculate offset for first valid character
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename == "TextFrame") {
                var contents = item.contents;
                for (var c = 0; c < targetInput.text.length; c++) {
                    var targetChar = targetInput.text.charAt(c);
                    if (contents.indexOf(targetChar) == -1) {
                        continue;
                    }
                    var refCenterY = createOutlineAndGetCenterY(item, refInput.text);
                    var targetCenterY = createOutlineAndGetCenterY(item, targetChar);
                    var yOffset = refCenterY - targetCenterY;
                    shiftInput.text = yOffset.toFixed(4);
                    previewShiftAll(targetInput, shiftInput, app.activeDocument.selection);
                    return;
                }
            }
        }
    };

    /* OKボタンのクリック処理 / Final OK button click */
    finalOkBtn.onClick = function() {
        if (targetInput.text.length == 0) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        if (refInput.text.length != 1) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        var shiftValue = Number(shiftInput.text);
        if (isNaN(shiftValue)) {
            alert(LABELS.numericErrorMsg[lang]);
            return;
        }
        dialog.close(1);
    };

    /* キャンセルボタンのクリック処理 / Cancel button click */
    cancelBtn.onClick = function() {
        dialog.close(0);
    };

    if (dialog.show() == 1) {
        var shiftValue = Number(shiftInput.text);
        return {
            target: targetInput.text,
            reference: refInput.text,
            shift: shiftValue
        };
    }
    return null;
}

/* メイン処理 / Main process */
function main() {
    try {
        if (app.documents.length == 0) {
            alert(LABELS.docOpenMsg[getLang()]);
            return;
        }

        var selection = app.activeDocument.selection;
        if (!selection || selection.length == 0) {
            alert(LABELS.selectFrameMsg[getLang()]);
            return;
        }

        var input = showDialog();
        if (!input) return;

        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename == "TextFrame") {
                var contents = item.contents;
                for (var c = 0; c < input.target.length; c++) {
                    var targetChar = input.target.charAt(c);
                    if (contents.indexOf(targetChar) == -1) {
                        continue;
                    }
                    var refCenterY = createOutlineAndGetCenterY(item, input.reference);
                    var targetCenterY = createOutlineAndGetCenterY(item, targetChar);
                    var yOffset = targetCenterY - refCenterY;
                    var chars = item.textRange.characters;
                    for (var j = 0; j < chars.length; j++) {
                        var ch = chars[j].contents;
                        if (ch && ch === targetChar) {
                            chars[j].characterAttributes.baselineShift = -yOffset;
                        }
                    }
                }
            }
        }

    } catch (e) {
        alert(LABELS.errorMsg[getLang()] + e);
    }
}

main();