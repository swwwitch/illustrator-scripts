#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
AddPageNumberFromTextSelection.jsx

概要：
_pagenumberレイヤー上で選択されたテキストを基準に、
すべてのアートボードに連番テキストを複製します。
ユーザーに開始番号・接頭辞・ゼロパディング・総ページ数表示を指定して、
アートボード数に応じたページ番号を作成します。

限定条件：
・選択テキストは_pagenumberレイヤー上のポイントテキストであること
・段落揃えの変更は行いません

作成日：2025-06-25
更新日: 2025-06-27
- v1.0.0 初版
- v1.0.1 テキストの複製ロジックを修正
- v1.0.2 ゼロ埋め機能、接頭辞、総ページ数表示機能を追加

Overview:
Based on the text selected in the _pagenumber layer,
this script duplicates the text as sequential page numbers on all artboards.
The user is prompted to enter a starting number, and page numbers are assigned according to the number of artboards.

Conditions:
- The selected text must be point text in the _pagenumber layer.
- Paragraph alignment will not be changed.

*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();
var LABELS = {
    dialogTitle: {
        ja: "ページ番号を追加",
        en: "Add Page Numbers"
    },
    promptMessage: {
        ja: "開始番号",
        en: "Starting number"
    },
    errorNotNumber: {
        ja: "有効な数字を入力してください",
        en: "Please enter a valid number"
    },
    errorInvalidSelection: {
        ja: "複製対象のテキストを_pagenumberレイヤーで選択してください",
        en: "Select text in the _pagenumber layer"
    }
};

function removeOtherTextFrames(layer, except) {
    for (var i = layer.textFrames.length - 1; i >= 0; i--) {
        var item = layer.textFrames[i];
        if (item !== except && item.typename === "TextFrame") {
            item.remove();
        }
    }
}

function getArtboardIndexByPosition(doc, pos) {
    for (var i = 0; i < doc.artboards.length; i++) {
        var abRect = doc.artboards[i].artboardRect;
        if (pos[0] >= abRect[0] && pos[0] <= abRect[2] && pos[1] <= abRect[1] && pos[1] >= abRect[3]) {
            return i;
        }
    }
    return -1;
}

function buildPageNumberString(num, maxDigits, prefix, zeroPad, totalPageNum, showTotal) {
    var numStr = String(num);
    if (zeroPad && numStr.length < maxDigits) {
        while (numStr.length < maxDigits) {
            numStr = "0" + numStr;
        }
    }
    numStr = prefix + numStr;
    if (showTotal) {
        numStr += "/" + totalPageNum;
    }
    return numStr;
}

function main() {
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    var prefixGroup = dialog.add("group");
    prefixGroup.orientation = "row";
    var prefixLabel = prefixGroup.add("statictext", undefined, "接頭辞");
    var prefixField = prefixGroup.add("edittext", undefined, "");
    prefixField.characters = 10;

    var inputGroup = dialog.add("group");
    inputGroup.orientation = "row";
    var label = inputGroup.add("statictext", undefined, LABELS.promptMessage[lang]);
    var inputField = inputGroup.add("edittext", undefined, "1");
    inputField.characters = 5;

    var zeroPadCheckbox = dialog.add("checkbox", undefined, "ゼロパディング");
    var totalPageCheckbox = dialog.add("checkbox", undefined, "総ページ数を表示");

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var cancelBtn = buttonGroup.add("button", undefined, "Cancel");
    var okBtn = buttonGroup.add("button", undefined, "OK");

    okBtn.onClick = function() {
        dialog.close(1);
    };
    cancelBtn.onClick = function() {
        dialog.close(0);
    };

    inputField.active = true;

    if (dialog.show() !== 1) {
        return;
    }

    var startNum = parseInt(inputField.text, 10);
    if (isNaN(startNum)) {
        alert(LABELS.errorNotNumber[lang]);
        return;
    }

    if (app.documents.length === 0) return;
    var doc = app.activeDocument;

    var pagenumberLayer;
    try {
        pagenumberLayer = doc.layers.getByName("_pagenumber");
    } catch (e) {
        pagenumberLayer = doc.layers.add();
        pagenumberLayer.name = "_pagenumber";
    }

    var sel = doc.selection;

    if (sel.length > 0 && sel[0].typename === "TextFrame") {
        sel[0].move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
    }

    removeOtherTextFrames(pagenumberLayer, sel[0]);

    var targetText = null;
    var abIndexToKeep = 0;
    var baseArtboard = doc.artboards[abIndexToKeep];
    var baseRect = baseArtboard.artboardRect;

    for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
        var tf = pagenumberLayer.textFrames[i];
        var pos = tf.position;
        if (pos[0] >= baseRect[0] && pos[0] <= baseRect[2] && pos[1] <= baseRect[1] && pos[1] >= baseRect[3]) {
            targetText = tf;
            break;
        }
    }

    if (!targetText) {
        for (var j = 1; j < doc.artboards.length; j++) {
            var abRect = doc.artboards[j].artboardRect;
            for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
                var tf = pagenumberLayer.textFrames[i];
                var pos = tf.position;
                if (pos[0] >= abRect[0] && pos[0] <= abRect[2] && pos[1] <= abRect[1] && pos[1] >= abRect[3]) {
                    targetText = tf;
                    break;
                }
            }
            if (targetText) break;
        }
    }

    if (!targetText) {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    var pos = targetText.position;
    var ab = doc.artboards[abIndexToKeep].artboardRect;
    doc.artboards.setActiveArtboardIndex(abIndexToKeep);

    var currentABIndex = getArtboardIndexByPosition(doc, pos);
    if (currentABIndex >= 0) {
        var currentAB = doc.artboards[currentABIndex].artboardRect;
        var offset = [ab[0] - currentAB[0], ab[1] - currentAB[1]];
        targetText.position = [pos[0] + offset[0], pos[1] + offset[1]];
    }

    removeOtherTextFrames(pagenumberLayer, targetText);

    if (targetText.layer.name !== "_pagenumber") {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    if (targetText.typename !== "TextFrame") {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    targetText.contents = String(startNum);

    targetText.selected = true;
    app.cut();
    app.executeMenuCommand('pasteInAllArtboard');

    var textFrames = [];
    for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
        var tf = pagenumberLayer.textFrames[i];
        if (tf !== targetText) {
            var tfPos = tf.position;
            var abIdx = getArtboardIndexByPosition(doc, tfPos);
            textFrames.push({frame: tf, abIdx: abIdx});
        }
    }

    textFrames.sort(function(a, b) { return a.abIdx - b.abIdx; });

    var maxNum = startNum + doc.artboards.length - 1;
    var maxDigits = String(maxNum).length;
    var prefix = prefixField.text;
    var zeroPad = zeroPadCheckbox.value;
    var showTotal = totalPageCheckbox.value;

    for (var i = 0; i < textFrames.length; i++) {
        var num = startNum + i;
        var numStr = buildPageNumberString(num, maxDigits, prefix, zeroPad, maxNum, showTotal);
        textFrames[i].frame.contents = numStr;
    }
}

main();