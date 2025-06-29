#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：SmartBaselineShifter.jsx

    スクリプト概要：
    選択したテキストフレーム内の指定文字にベースラインシフト（ポイント単位）を個別適用します。
    ダイアログで文字列、シフト量（整数・小数）、符号、リセットオプションを設定し、即時プレビューが可能です。

    処理の流れ：
    1. テキストフレームを選択
    2. ダイアログで対象文字、シフト量、符号、リセットを設定
    3. プレビュー確認後、OKで確定

    対象：選択中のテキストフレーム（ポイント文字・エリア文字含む）
    更新日：2024-06-29
    限定条件：Illustrator 2025 以降推奨
*/

// 言語切り替え
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// ラベル定義（UIの表示順）
var LABELS = {
    dialogTitle: { ja: "ベースラインシフト", en: "Baseline Shift" },
    reset: { ja: "すべてをリセット", en: "Reset All" },
    targetText: { ja: "対象文字列:", en: "Target Text:" },
    shiftValue: { ja: "シフト量:", en: "Shift Value:" },
    sign: { ja: "符号:", en: "Sign:" },
    positive: { ja: "正", en: "Positive" },
    negative: { ja: "負", en: "Negative" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    alertOpenDoc: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
    alertSelectText: { ja: "テキストを選択してください。", en: "Please select text." },
    alertSelectFrame: { ja: "テキストフレームを選択してください。", en: "Please select text frame." },
    error: { ja: "処理中にエラーが発生しました: ", en: "An error occurred: " },
    previewLabel: { ja: "シフト量: ", en: "Shift Value: " }
};

function collectTextFrames(item, array) {
    if (item.typename === "TextFrame") {
        array.push(item);
    } else if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            collectTextFrames(item.pageItems[i], array);
        }
    }
}

function resetAllBaselineShift(frames) {
    for (var i = 0; i < frames.length; i++) {
        var tf = frames[i];
        var contents = tf.contents;
        for (var j = 0; j < contents.length; j++) {
            var range = tf.textRange.characters[j];
            range.characterAttributes.baselineShift = 0;
        }
    }
}

function main() {
    var lang = getCurrentLang();

    if (app.documents.length === 0) {
        alert(LABELS.alertOpenDoc[lang]);
        return;
    }

    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert(LABELS.alertSelectText[lang]);
        return;
    }

    var textFrames = [];
    for (var i = 0; i < doc.selection.length; i++) {
        collectTextFrames(doc.selection[i], textFrames);
    }

    if (textFrames.length === 0) {
        alert(LABELS.alertSelectFrame[lang]);
        return;
    }

    resetAllBaselineShift(textFrames);

    var defaultTarget = "";
    if (textFrames.length > 0) {
        var contents = textFrames[0].contents;
        var matches = contents.match(/[^0-9\s]/g);
        if (matches) {
            var uniqueChars = {};
            for (var i = 0; i < matches.length; i++) {
                uniqueChars[matches[i]] = true;
            }
            for (var c in uniqueChars) {
                defaultTarget += c;
            }
        }
    }

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.alignChildren = "fill";

    // リセットチェックボックス
    var resetCheckbox = dialog.add("checkbox", undefined, LABELS.reset[lang]);
    resetCheckbox.value = false;
    resetCheckbox.onClick = function() {
        if (resetCheckbox.value) {
            resetAllBaselineShift(textFrames);
            app.redraw();
        }
    };

    var targetGroup = dialog.add("group");
    targetGroup.add("statictext", undefined, LABELS.targetText[lang]);
    var targetInput = targetGroup.add("edittext", undefined, defaultTarget);
    targetInput.characters = 10;

    var shiftPanel = dialog.add("panel", undefined, LABELS.dialogTitle[lang]);
    shiftPanel.orientation = "column";
    shiftPanel.alignChildren = "left";
    shiftPanel.margins = [15, 20, 15, 10];

    var shiftGroup = shiftPanel.add("group");
    shiftGroup.add("statictext", undefined, LABELS.shiftValue[lang]);
    var shiftIntInput = shiftGroup.add("edittext", undefined, "0");
    shiftIntInput.characters = 3;
    shiftGroup.add("statictext", undefined, ".");
    var shiftDecInput = shiftGroup.add("edittext", undefined, "0");
    shiftDecInput.characters = 3;

    var signGroup = shiftPanel.add("group");
    signGroup.add("statictext", undefined, LABELS.sign[lang]);
    var signOptions = signGroup.add("radiobutton", undefined, LABELS.positive[lang]);
    var signOptionsNeg = signGroup.add("radiobutton", undefined, LABELS.negative[lang]);
    signOptions.value = true;
    signGroup.margins = [0, 0, 0, 10];

    var resultText = shiftPanel.add("statictext", undefined, LABELS.previewLabel[lang] + "0");
    resultText.characters = 10; // 表示幅を確保
    resultText.alignment = "center";

    signOptions.onClick = signOptionsNeg.onClick = function() {
        previewShiftAll();
    };

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang]);

    var lastTarget = "";
    var lastShift = 0;

    targetInput.onChanging = shiftIntInput.onChanging = shiftDecInput.onChanging = function() {
        previewShiftAll();
    };

    cancelBtn.onClick = function() {
        for (var j = 0; j < textFrames.length; j++) {
            resetBaselineShift(textFrames[j], lastTarget);
        }
        dialog.close();
    };

    okBtn.onClick = function() {
        dialog.close();
    };

    function previewShiftAll() {
        var targetText = targetInput.text;
        var intPart = parseInt(shiftIntInput.text, 10);
        var decPart = parseInt(shiftDecInput.text, 10);
        if (isNaN(intPart)) intPart = 0;
        if (isNaN(decPart)) decPart = 0;
        var shiftValue = parseFloat(intPart + "." + decPart);

        var displayValue = "";
        if (shiftValue === 0) {
            displayValue = "0.0";
        } else {
            displayValue = "" + Math.abs(shiftValue).toFixed(1);
        }
        if (signOptionsNeg.value && shiftValue > 0) {
            displayValue = "-" + displayValue;
        }
        resultText.text = "シフト量: " + displayValue;

        if (signOptionsNeg.value && shiftValue > 0) {
            shiftValue = -shiftValue;
        }

        if (!targetText || isNaN(shiftValue)) {
            for (var j = 0; j < textFrames.length; j++) {
                resetBaselineShift(textFrames[j], lastTarget);
            }
            lastTarget = targetText;
            lastShift = 0;
            app.redraw();
            return;
        }

        for (var j = 0; j < textFrames.length; j++) {
            resetBaselineShift(textFrames[j], lastTarget);
            for (var k = 0; k < targetText.length; k++) {
                var ch = targetText.charAt(k);
                applyShiftToTargetText(textFrames[j], ch, shiftValue);
            }
        }

        lastTarget = targetText;
        lastShift = shiftValue;
        app.redraw();
    }

    function applyShiftToTargetText(textFrame, targetStr, shift) {
        try {
            var contents = textFrame.contents;
            var index = contents.indexOf(targetStr);

            while (index !== -1) {
                var range = textFrame.textRange.characters[index];
                range.characterAttributes.baselineShift = shift;
                index = contents.indexOf(targetStr, index + 1);
            }
        } catch (e) {
            alert(LABELS.error[lang] + e);
        }
    }

    function resetBaselineShift(textFrame, targetStr) {
        if (!targetStr) return;
        try {
            var contents = textFrame.contents;
            for (var k = 0; k < targetStr.length; k++) {
                var ch = targetStr.charAt(k);
                var index = contents.indexOf(ch);

                while (index !== -1) {
                    var range = textFrame.textRange.characters[index];
                    range.characterAttributes.baselineShift = 0;
                    index = contents.indexOf(ch, index + 1);
                }
            }
        } catch (e) {
            // 無視
        }
    }

    dialog.show();
}

main();