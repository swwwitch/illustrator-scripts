#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartBaselineShifter.jsx

### 概要

- 選択したテキストフレーム内の指定文字に、ポイント単位のベースラインシフトを個別適用するIllustrator用スクリプトです。
- 対象文字列、シフト量（整数・小数）、符号、リセットオプションをダイアログで設定し、即時プレビューが可能です。

### 主な機能

- 対象文字の選択と指定
- シフト量の整数・小数単位指定、符号切り替え
- 全リセット機能
- 即時プレビューと元に戻す操作
- 日本語／英語インターフェース対応

### 処理の流れ

1. テキストフレームを選択
2. ダイアログで対象文字列とシフト量を設定
3. プレビューを確認
4. OKで確定、キャンセルで元に戻す

### 更新履歴

- v1.0.0 (20240629) : 初期バージョン
- v1.0.3 (20240629) : +/-ボタン追加
- v1.0.4 (20240629) : ダイアログ2カラム化、正規表現対応
- v1.0.5 (20240630) : TextRange選択用関数追加
- v1.0.6 (20240630) : 正規表現対応削除、微調整

---

### Script Name:

SmartBaselineShifter.jsx

### Overview

- An Illustrator script to individually apply baseline shift (in points) to specified characters in selected text frames.
- Allows setting target characters, shift amount (integer and decimal), sign, and reset options in a dialog with instant preview.

### Main Features

- Specify and select target characters
- Set shift amount in integer and decimal points, toggle sign
- Reset all baseline shifts
- Instant preview and undo functionality
- Japanese and English UI support

### Process Flow

1. Select text frames
2. Configure target characters and shift amount in dialog
3. Check preview
4. Confirm with OK or revert with Cancel

### Update History

- v1.0.0 (20240629): Initial version
- v1.0.3 (20240629): Added +/- buttons
- v1.0.4 (20240629): Two-column dialog layout, regex support
- v1.0.5 (20240630): Added function for TextRange selection
- v1.0.6 (20240630): Removed regex support, fine adjustments
*/


// 現在の環境言語を判定（日本語か英語）
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var LABELS = {
    dialogTitle: { ja: "ベースラインシフト", en: "Baseline Shift" },
    targetText: { ja: "対象文字列:", en: "Target Text:" },
    shiftValue: { ja: "シフト量:", en: "Shift Value:" },
    sign: { ja: "符号:", en: "Sign:" },
    positive: { ja: "正", en: "Positive" },
    negative: { ja: "負", en: "Negative" },
    resetAll: { ja: "すべてをリセット", en: "Reset All" },
    resetBtn: { ja: "初期化", en: "Reset" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    alertOpenDoc: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
    alertSelectText: { ja: "テキストを選択してください。", en: "Please select text." },
    alertSelectFrame: { ja: "テキストフレームを選択してください。", en: "Please select text frame." },
    error: { ja: "処理中にエラーが発生しました: ", en: "An error occurred: " },
    previewLabel: { ja: "シフト量: ", en: "Shift Value: " }
};

// 選択中の TextRange を含む単一の TextFrame を選択し直す
function selectSingleTextFrameFromTextRange() {
    if (app.selection.constructor.name === "TextRange") {
        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length === 1) {
            app.executeMenuCommand("deselectall"); // 現在の選択を解除
            app.selection = [textFramesInStory[0]]; // 該当の TextFrame を選択
            try {
                app.selectTool("Adobe Select Tool"); // 選択ツールに戻す
            } catch (e) {}
        }
    }
}

// 指定アイテムからテキストフレームを再帰的に収集
function collectTextFrames(item, array) {
    if (item.typename === "TextFrame") {
        array.push(item);
    } else if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            collectTextFrames(item.pageItems[i], array);
        }
    }
}

// 指定テキストフレーム群のベースラインシフトをリセット（対象文字列指定可）
function resetBaselineShift(frames, targetStr) {
    for (var i = 0; i < frames.length; i++) {
        var tf = frames[i];
        var contents = tf.contents;
        try {
            var chars = tf.textRange.characters;
            if (!targetStr) {
                // 全文字のベースラインシフトをリセット
                for (var j = 0; j < chars.length; j++) {
                    chars[j].characterAttributes.baselineShift = 0;
                }
            } else {
                // 対象文字列にマッチする箇所のみリセット（エスケープ済み正規表現使用）
                var regex = new RegExp(targetStr.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1"), "g");
                var match;
                while ((match = regex.exec(contents)) !== null) {
                    var index = match.index;
                    for (var k = 0; k < match[0].length; k++) {
                        chars[index + k].characterAttributes.baselineShift = 0;
                    }
                }
            }
        } catch (e) {
            // エラーは無視
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
    if (!doc.selection || doc.selection.length === 0) {
        alert(LABELS.alertSelectText[lang]);
        return;
    }

    selectSingleTextFrameFromTextRange();

    var textFrames = [];
    for (var i = 0; i < doc.selection.length; i++) {
        collectTextFrames(doc.selection[i], textFrames);
    }
    if (textFrames.length === 0) {
        alert(LABELS.alertSelectFrame[lang]);
        return;
    }

    resetBaselineShift(textFrames);

    // 対象文字列の初期値設定（数字と空白以外のユニーク文字）
    var defaultTarget = "";
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

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "row";

    var leftPanel = dialog.add("group");
    leftPanel.orientation = "column";
    leftPanel.alignChildren = "fill";

    var resetCheckbox = leftPanel.add("checkbox", undefined, LABELS.resetAll[lang]);
    resetCheckbox.value = false;
    resetCheckbox.onClick = function() {
        if (resetCheckbox.value) {
            resetBaselineShift(textFrames);
            lastTarget = "";
            lastShift = 0;
            shiftIntInput.text = "0";
            shiftDecInput.text = "0";
            signPositive.value = true;
            resultText.text = LABELS.previewLabel[lang] + "0.0";
            app.redraw();
        }
    };

    var targetPanel = leftPanel.add("panel", undefined, LABELS.targetText[lang]);
    targetPanel.orientation = "column";
    targetPanel.alignChildren = "left";
    targetPanel.margins = [15, 20, 15, 10];

    var targetInput = targetPanel.add("edittext", undefined, defaultTarget);
    targetInput.characters = 15;

    var shiftPanel = leftPanel.add("panel", undefined, LABELS.dialogTitle[lang]);
    shiftPanel.orientation = "column";
    shiftPanel.alignChildren = "left";
    shiftPanel.margins = [15, 20, 15, 10];

    var resultText = shiftPanel.add("statictext", undefined, LABELS.previewLabel[lang] + "0");
    resultText.characters = 16;
    resultText.alignment = "center";

    var signGroup = shiftPanel.add("group");
    signGroup.add("statictext", undefined, LABELS.sign[lang]);
    var signPositive = signGroup.add("radiobutton", undefined, LABELS.positive[lang]);
    var signNegative = signGroup.add("radiobutton", undefined, LABELS.negative[lang]);
    signPositive.value = true;
    signGroup.margins = [0, 10, 0, 0];

    var shiftGroup = shiftPanel.add("group");

    var shiftIntInput = shiftGroup.add("edittext", undefined, "0");
    shiftIntInput.characters = 3;

    var intSignGroup = shiftGroup.add("group");
    intSignGroup.orientation = "column";
    intSignGroup.alignment = "center";
    intSignGroup.spacing = 2;
    var plusIntBtn = intSignGroup.add("button", undefined, "+");
    plusIntBtn.preferredSize = [18, 18];
    var minusIntBtn = intSignGroup.add("button", undefined, "−");
    minusIntBtn.preferredSize = [18, 18];

    plusIntBtn.onClick = function() {
        var val = parseInt(shiftIntInput.text, 10);
        if (isNaN(val)) val = 0;
        shiftIntInput.text = (val + 1).toString();
        previewShiftAll();
    };
    minusIntBtn.onClick = function() {
        var val = parseInt(shiftIntInput.text, 10);
        if (isNaN(val)) val = 0;
        shiftIntInput.text = (val - 1).toString();
        previewShiftAll();
    };

    shiftGroup.add("statictext", undefined, ".");
    var shiftDecInput = shiftGroup.add("edittext", undefined, "0");
    shiftDecInput.characters = 3;

    var decSignGroup = shiftGroup.add("group");
    decSignGroup.orientation = "column";
    decSignGroup.alignment = "center";
    decSignGroup.spacing = 2;
    var plusDecBtn = decSignGroup.add("button", undefined, "+");
    plusDecBtn.preferredSize = [18, 18];
    var minusDecBtn = decSignGroup.add("button", undefined, "−");
    minusDecBtn.preferredSize = [18, 18];

    plusDecBtn.onClick = function() {
        var val = parseInt(shiftDecInput.text, 10);
        if (isNaN(val)) val = 0;
        if (val < 9) val++;
        shiftDecInput.text = val.toString();
        previewShiftAll();
    };
    minusDecBtn.onClick = function() {
        var val = parseInt(shiftDecInput.text, 10);
        if (isNaN(val)) val = 0;
        if (val > 0) val--;
        shiftDecInput.text = val.toString();
        previewShiftAll();
    };

    signPositive.onClick = signNegative.onClick = function() {
        previewShiftAll();
    };

    var rightPanel = dialog.add("group");
    rightPanel.orientation = "column";
    rightPanel.alignChildren = "center";

    var buttonGroup = rightPanel.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.alignment = "center";

    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang]);
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var spacer = buttonGroup.add("statictext", undefined, " ");
    spacer.preferredSize = [0, 120];
    var resetBtn = buttonGroup.add("button", undefined, LABELS.resetBtn[lang]);

    okBtn.preferredSize = [100, 30];
    cancelBtn.preferredSize = [100, 30];
    resetBtn.preferredSize = [100, 30];

    resetBtn.onClick = function() {
        signPositive.value = true;
        shiftIntInput.text = "0";
        shiftDecInput.text = "0";
        previewShiftAll();
        app.redraw();
    };

    var lastTarget = "";
    var lastShift = 0;

    targetInput.onChanging = shiftIntInput.onChanging = shiftDecInput.onChanging = function() {
        previewShiftAll();
    };

    cancelBtn.onClick = function() {
        resetBaselineShift(textFrames);
        dialog.close();
    };

    okBtn.onClick = function() {
        dialog.close();
    };

    // ベースラインシフトのプレビュー適用処理
    function previewShiftAll() {
        var targetText = targetInput.text;
        var intPart = parseInt(shiftIntInput.text, 10);
        var decPart = parseInt(shiftDecInput.text, 10);
        if (isNaN(intPart)) intPart = 0;
        if (isNaN(decPart)) decPart = 0;
        var shiftValue = parseFloat(intPart + "." + decPart);

        var displayVal = shiftValue === 0 ? "0.0" : Math.abs(shiftValue).toFixed(1);
        if (signNegative.value && shiftValue > 0) {
            displayVal = "-" + displayVal;
            shiftValue = -shiftValue;
        }
        resultText.text = LABELS.previewLabel[lang] + displayVal;

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
            // indexOfによる高速検索ロジックのみ
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

    dialog.show();
}

main(); 