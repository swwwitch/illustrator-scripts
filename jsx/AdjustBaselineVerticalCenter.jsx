#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：AdjustBaselineVerticalCenter.jsx

概要:
選択したテキストフレーム内の指定文字を、基準文字に合わせて垂直方向（ベースライン）を調整します。
Adjusts the baseline of specified characters in one or multiple text frames to align with a reference character.

処理の流れ:
1. ダイアログで対象文字と基準文字を指定
2. 各文字のアウトライン化で中央Y座標を取得
3. ベースラインシフトで補正

対象:
- テキストフレーム（複数選択可）

限定条件:
- アウトライン済みや非テキストオブジェクトは対象外

オリジナルアイデア：
Egor Chistyakov https://x.com/tchegr

オリジナルからの変更点：
- 対象文字は自動入力（複数ある場合には最頻出記号を選択）
- 手動での上書き入力も可能
- 複数のテキストオブジェクトに対しても一括適用可能

Changes from original:
- Target character is auto-filled (if multiple, the most frequent symbol is selected)
- Manual override input is also possible
- Can be applied to multiple text objects at once

更新履歴:
- v1.0.0 (2025-07-04): 初版リリース
- v1.0.1 (2025-07-05): 対象文字の自動取得ロジック改善
- v1.0.2 (2025-07-06): 対象文字のみを選択しても実行できるように
*/

var LABELS = {
    dialogTitle: { ja: "ベースライン調整 v1.0.2", en: "Adjust Baseline" },
    infoTextMsg: { ja: "対象文字を縦方向に揃えます。", en: "Align target character vertically." },
    targetCharLabel: { ja: "対象文字:", en: "Target Character:" },
    baseCharLabel: { ja: "基準文字:", en: "Reference Character:" },
    selectFrameMsg: { ja: "テキストフレームを選択してください。", en: "Select one or more text frames." },
    docOpenMsg: { ja: "ドキュメントが開かれていません。", en: "No document open." },
    invalidCharMsg: { ja: "1文字ずつ入力してください。", en: "Enter exactly one character." },
    notFoundMsg: { ja: "対象文字が含まれていません。", en: "Target character not found." },
    errorMsg: { ja: "エラー: ", en: "Error: " },
    okBtnLabel: { ja: "調整", en: "Adjust" },
    cancelBtnLabel: { ja: "キャンセル", en: "Cancel" }
};

function getLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

function getCenterY(item) {
    var b = item.geometricBounds;
    return b[1] - (b[1] - b[3]) / 2;
}

function adjustBaseline(textFrame, targetChar, referenceChar) {
    if (textFrame.contents.indexOf(targetChar) === -1) {
        alert(LABELS.notFoundMsg[getLang()]);
        return;
    }

    var refFrame = textFrame.duplicate();
    refFrame.contents = referenceChar;
    refFrame.filled = false;
    refFrame.stroked = false;
    var refOutline = refFrame.createOutline();
    var refCenterY = getCenterY(refOutline);

    var targetFrame = textFrame.duplicate();
    targetFrame.contents = targetChar;
    targetFrame.filled = false;
    targetFrame.stroked = false;
    var targetOutline = targetFrame.createOutline();
    var targetCenterY = getCenterY(targetOutline);

    var yOffset = targetCenterY - refCenterY;

    var chars = textFrame.textRange.characters;
    for (var i = 0; i < chars.length; i++) {
        if (chars[i].contents === targetChar) {
            chars[i].characterAttributes.baselineShift = -yOffset;
        }
    }

    targetOutline.remove();
    refOutline.remove();
}

function runGetTextScript() {
    if (!app.documents.length || !app.selection.length) {
        alert(LABELS.docOpenMsg[getLang()]);
        return;
    }

    var selectionState = getSelectionState();
    var defaultTarget = "";

    if (selectionState === "editing") {
        app.executeMenuCommand("copy");
        var savedTextFrame = null;
        if (app.selection.constructor.name === "TextRange") {
            var storyFrames = app.selection.story.textFrames;
            if (storyFrames.length === 1) {
                app.executeMenuCommand("deselectall");
                savedTextFrame = storyFrames[0];
                app.selection = [savedTextFrame];
            }
        }
        app.executeMenuCommand("paste");
        var pasted = app.activeDocument.selection[0];
        if (pasted && pasted.contents) defaultTarget = pasted.contents;
        if (pasted) pasted.remove();

        if (savedTextFrame) {
            app.selection = [savedTextFrame];
            try { savedTextFrame.textRange.selected = true; } catch (e) {}
            app.selectTool("Adobe Select Tool");
            app.redraw();
        }
    } else if (selectionState === "selected") {
        if (app.selection && app.selection.length > 0) {
            var charCount = {};
            for (var i = 0; i < app.selection.length; i++) {
                var item = app.selection[i];
                if (item.typename == "TextFrame") {
                    var selText = item.contents;
                    for (var j = 0; j < selText.length; j++) {
                        var c = selText.charAt(j);
                        if (!c.match(/^[A-Za-z0-9\s]$/)) {
                            charCount[c] = (charCount[c] || 0) + 1;
                        }
                    }
                }
            }
            var maxCount = 0;
            for (var key in charCount) {
                if (charCount[key] > maxCount) {
                    maxCount = charCount[key];
                    defaultTarget = key;
                }
            }
        }
    }

    var result = showDialog(defaultTarget);
    if (result) {
        for (var i = 0; i < app.selection.length; i++) {
            var textFrame = app.selection[i];
            if (textFrame.typename === "TextFrame") {
                adjustBaseline(textFrame, result.target, result.reference);
            }
        }
    }
}

function showDialog(defaultTarget) {
    var lang = getLang();
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    dialog.add("statictext", undefined, LABELS.infoTextMsg[lang]);

    var inputGroup = dialog.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";
    inputGroup.margins = [15, 5, 15, 5];

    var targetGroup = inputGroup.add("group");
    targetGroup.add("statictext", undefined, LABELS.targetCharLabel[lang]);
    var targetInput = targetGroup.add("edittext", undefined, defaultTarget);
    targetInput.characters = 5;

    var refGroup = inputGroup.add("group");
    refGroup.add("statictext", undefined, LABELS.baseCharLabel[lang]);
    var refInput = refGroup.add("edittext", undefined, "0");
    refInput.characters = 5;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancelBtnLabel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.okBtnLabel[lang], { name: "ok" });

    okBtn.onClick = function () {
        if (targetInput.text.length !== 1 || refInput.text.length !== 1) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        dialog.close(1);
    };
    cancelBtn.onClick = function () { dialog.close(0); };

    if (dialog.show() == 1) {
        return { target: targetInput.text, reference: refInput.text };
    }
    return null;
}

function getSelectionState() {
    if (app.selection.constructor.name === "TextRange") return "editing";
    var sel = app.selection[0];
    return (sel && sel.typename === "TextFrame") ? "selected" : "none";
}

runGetTextScript();