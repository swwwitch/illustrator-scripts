#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
AddPageNumberFromTextSelection.jsx
更新日: 2025-06-27

概要：
_pagenumberレイヤー上で選択されたテキストを基準に、
すべてのアートボードに連番テキストを複製します。
ユーザーに開始番号を入力してもらい、アートボード数に応じて番号を振ります。

限定条件：
・選択テキストは_pagenumberレイヤー上のポイントテキストであること
・段落揃えの変更は行いません

Overview:
Based on the text selected in the _pagenumber layer,
this script duplicates the text as sequential page numbers on all artboards.
The user is prompted to enter a starting number, and page numbers are assigned according to the number of artboards.

Conditions:
- The selected text must be point text in the _pagenumber layer.
- Paragraph alignment will not be changed.

Updated: 2025-06-27

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
        ja: "開始番号を入力してください",
        en: "Enter starting number"
    },
    errorNotNumber: {
        ja: "有効な数字を入力してください",
        en: "Please enter a valid number"
    },
    errorLayerHidden: {
        ja: "_pagenumberレイヤーが非表示またはロックされています",
        en: "_pagenumber layer is hidden or locked"
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

function main() {
    var startNum = prompt(LABELS.promptMessage[lang], "1");
    if (startNum === null) return;
    startNum = parseInt(startNum, 10);
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

    if (!pagenumberLayer.visible || pagenumberLayer.locked) {
        alert(LABELS.errorLayerHidden[lang]);
        return;
    }

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

    for (var i = 0; i < textFrames.length; i++) {
        var num = startNum + i;
        textFrames[i].frame.contents = String(num);
    }
}

main();