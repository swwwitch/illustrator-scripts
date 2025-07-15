#target illustrator
$.localize = true;
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.7";

/*
### スクリプト名：

SmartBaselineShifter.jsx

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartBaselineShifter.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBaselineShifter.md

### 概要：

- 選択したテキストフレーム内の指定文字にポイント単位でベースラインシフトを個別適用
- ダイアログで対象文字列とシフト量を設定し、即時プレビュー可能

### 主な機能：

- 対象文字列の指定
- シフト量の整数・小数指定
- すべてをリセット
- 即時プレビュー、元に戻す操作
- 日本語／英語UI対応

#### 対象文字列の指定

- 選択しているテキストから、数字やスペースなどを除いたものが自動的に入ります。
- 編集可能です。
- 「:-」と入力すれば「:」と「-」の両方が対象になります。

#### 値の変更

- ↑↓キーで±1増減
- shiftキーを併用すると±10増減
- optionキーを併用すると±0.1増減

### 処理の流れ：

1. テキストフレームを選択
2. ダイアログで設定
3. プレビュー確認
4. OKで確定、キャンセルで元に戻す

### note

https://note.com/dtp_tranist/n/n2e19ad0bdb83

### 更新履歴：

- v1.0 (20240629) : 初期バージョン
- v1.3 (20240629) : +/-ボタン追加
- v1.4 (20240629) : ダイアログ2カラム化、正規表現対応
- v1.5 (20240630) : TextRange選択用関数追加
- v1.6 (20240630) : 正規表現対応削除、微調整
- v1.7 (20240630) : ↑↓キーでの値の変更機能追加、UIの再設計

---

### Script Name:

SmartBaselineShifter.jsx

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SmartBaselineShifter.jsx

### Readme (GitHub):

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBaselineShifter.md

### Overview:

- Apply baseline shift (points) individually to specified characters in selected text frames
- Configure target characters and shift amount in dialog with instant preview

### Main Features:

- Specify target characters
- Set shift amount as integer/decimal
- Reset all baseline shifts
- Instant preview and undo
- Japanese/English UI support

### Process Flow:

1. Select text frames
2. Configure in dialog
3. Check preview
4. Confirm with OK or revert with Cancel

### Update History:

- v1.0 (20240629): Initial version
- v1.3 (20240629): Added +/- buttons
- v1.4 (20240629): Two-column dialog layout, regex support
- v1.5 (20240630): Added function for TextRange selection
- v1.6 (20240630): Removed regex support, fine adjustments
*/

// 選択中の TextRange を含む単一の TextFrame を選択し直す / Re-select a single TextFrame containing the selected TextRange
function selectSingleTextFrameFromTextRange() {
    if (app.selection.constructor.name === "TextRange") {
        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length === 1) {
            app.executeMenuCommand("deselectall");
            app.selection = [textFramesInStory[0]];
            try {
                app.selectTool("Adobe Select Tool");
            } catch (e) {}
        }
    }
}

// 指定アイテムからテキストフレームを再帰的に収集 / Recursively collect TextFrames from specified item
function collectTextFrames(item, array) {
    if (item.typename === "TextFrame") {
        array.push(item);
    } else if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            collectTextFrames(item.pageItems[i], array);
        }
    }
}

// 指定テキストフレーム群のベースラインシフトをリセット（対象文字列指定可） / Reset baseline shift for specified TextFrames (optional target characters)
function resetBaselineShift(frames, targetStr) {
    for (var i = 0; i < frames.length; i++) {
        var tf = frames[i];
        var contents = tf.contents;
        try {
            var chars = tf.textRange.characters;
            if (!targetStr) {
                for (var j = 0; j < chars.length; j++) {
                    chars[j].characterAttributes.baselineShift = 0;
                }
            } else {
                for (var c = 0; c < targetStr.length; c++) {
                    var ch = targetStr.charAt(c);
                    var pos = contents.indexOf(ch);
                    while (pos !== -1) {
                        chars[pos].characterAttributes.baselineShift = 0;
                        pos = contents.indexOf(ch, pos + 1);
                    }
                }
            }
        } catch (e) {
            // エラーは無視 / Ignore errors
        }
    }
}

function changeValueByArrowKey(editText, onUpdate) {
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
            value = Math.round(value * 10) / 10;
        } else {
            value = Math.round(value);
        }

        editText.text = value;
        if (typeof onUpdate === "function") {
            onUpdate(editText.text);
        }
    });
}

var LABELS = {
    dialogTitle: "ベースラインシフトの調整 " + SCRIPT_VERSION,
    shiftSectionTitle: "ベースラインシフト",
    targetText: "対象文字列:",
    shiftValue: "シフト量:",
    resetAll: "すべてをリセット",
    cancel: "キャンセル",
    ok: "OK",
    alertOpenDoc: "ドキュメントを開いてください。",
    alertSelectText: "テキストを選択してください。",
    alertSelectFrame: "テキストフレームを選択してください。",
    error: "処理中にエラーが発生しました: "
};

// メイン処理 / Main process
function main() {

    if (app.documents.length === 0) {
        alert(LABELS.alertOpenDoc);
        return;
    }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert(LABELS.alertSelectText);
        return;
    }

    selectSingleTextFrameFromTextRange();

    var textFrames = [];
    for (var i = 0; i < doc.selection.length; i++) {
        collectTextFrames(doc.selection[i], textFrames);
    }
    if (textFrames.length === 0) {
        alert(LABELS.alertSelectFrame);
        return;
    }

    resetBaselineShift(textFrames);

    var defaultTarget = "";
    var contents = textFrames[0].contents;
    var matches = contents.match(/[^\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF0-9\s]/g);
    if (matches) {
        var uniqueChars = {};
        for (var i = 0; i < matches.length; i++) {
            uniqueChars[matches[i]] = true;
        }
        for (var c in uniqueChars) {
            defaultTarget += c;
        }
    }

    var dialog = new Window("dialog", LABELS.dialogTitle);
    dialog.orientation = "column";
    dialog.alignChildren = "top";

    var leftPanel = dialog.add("group");
    leftPanel.orientation = "column";
    leftPanel.alignChildren = "fill";

    var resetGroup = leftPanel.add("group");
    resetGroup.orientation = "row";
    resetGroup.alignment = "center";
    var resetCheckbox = resetGroup.add("checkbox", undefined, LABELS.resetAll);
    resetCheckbox.value = false;
    resetCheckbox.onClick = function() {
        if (resetCheckbox.value) {
            resetBaselineShift(textFrames);
            lastTarget = "";
            lastShift = 0;
            shiftInput.text = "0";
            app.redraw();
        }
    };

    var shiftPanel = leftPanel.add("panel", undefined, LABELS.shiftSectionTitle);
    shiftPanel.orientation = "column";
    shiftPanel.alignChildren = "fill";
    shiftPanel.margins = [15, 20, 15, 10];

    var targetGroup = shiftPanel.add("group");
    targetGroup.orientation = "row";
    var targetLabel = targetGroup.add("statictext", undefined, LABELS.targetText);
    targetLabel.preferredSize.width = 80;
    targetLabel.justify = "right";
    var targetInput = targetGroup.add("edittext", undefined, defaultTarget);
    targetInput.characters = 10;

    var shiftLabelGroup = shiftPanel.add("group");
    shiftLabelGroup.orientation = "row";
    var shiftLabel = shiftLabelGroup.add("statictext", undefined, LABELS.shiftValue);
    shiftLabel.preferredSize.width = 80;
    shiftLabel.justify = "right";
    var shiftInput = shiftLabelGroup.add("edittext", undefined, "0");
    shiftInput.characters = 5;

    var buttonGroup = leftPanel.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel);
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok);

    var lastTarget = "";
    var lastShift = 0;

    function previewShiftAll() {
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

    changeValueByArrowKey(shiftInput, previewShiftAll);
    shiftInput.active = true;

    targetInput.onChanging = shiftInput.onChanging = function() {
        previewShiftAll();
    };

    cancelBtn.onClick = function() {
        resetBaselineShift(textFrames);
        dialog.close();
    };

    okBtn.onClick = function() {
        dialog.close();
    };

    dialog.show();
}

main();