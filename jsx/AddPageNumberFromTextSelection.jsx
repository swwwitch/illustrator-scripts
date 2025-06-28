#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
AddPageNumberFromTextSelection.jsx

概要:
_pagenumberレイヤー上で選択されたテキストを基準に、
すべてのアートボードに連番テキストを複製するスクリプト。
ユーザーは開始番号・接頭辞・ゼロパディング・総ページ数表示を指定可能。
アートボード数に応じてページ番号を自動生成します。

制約:
- 複製元テキストは_pagenumberレイヤー上のポイントテキストであること
- 段落揃えの変更は行いません

作成日：2025-06-25
更新日: 2025-06-28
- v1.0.0 初版
- v1.0.1 テキスト複製ロジック修正
- v1.0.2 ゼロ埋め・接頭辞・総ページ数表示追加
- v1.0.3 プレビュー機能を追加

課題：
- プレビュー時、元のテキストが残ってしまうため、ダブったように見えます。OKボタンを押せば消えるのですが…

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
    errorInvalidSelection: {
        ja: "複製対象のテキストを_pagenumberレイヤーで選択してください",
        en: "Select text in the _pagenumber layer"
    }
};

// 指定レイヤー上の他のテキストフレームを削除（exceptを除く）
function removeOtherTextFrames(layer, except) {
    for (var i = layer.textFrames.length - 1; i >= 0; i--) {
        var item = layer.textFrames[i];
        if (item !== except && item.typename === "TextFrame") {
            item.remove();
        }
    }
}

// 座標からアートボードインデックスを取得
function getArtboardIndexByPosition(doc, pos) {
    for (var i = 0; i < doc.artboards.length; i++) {
        var abRect = doc.artboards[i].artboardRect;
        if (pos[0] >= abRect[0] && pos[0] <= abRect[2] && pos[1] <= abRect[1] && pos[1] >= abRect[3]) {
            return i;
        }
    }
    return -1;
}

// ページ番号文字列生成
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

// テキスト複製&内容更新（再利用性向上のため小関数化）
function generatePageNumbers(doc, pagenumberLayer, targetText, baseRect, startNum, prefix, zeroPad, showTotal) {
    removeOtherTextFrames(pagenumberLayer, targetText);
    var abCount = doc.artboards.length;
    var maxNum = startNum + abCount - 1;
    var maxDigits = String(maxNum).length;
    for (var i = 0; i < abCount; i++) {
        var abRect = doc.artboards[i].artboardRect;
        var newTF = targetText.duplicate(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
        newTF.position = [
            targetText.position[0] + (abRect[0] - baseRect[0]),
            targetText.position[1] + (abRect[1] - baseRect[1])
        ];
        var numStr = buildPageNumberString(startNum + i, maxDigits, prefix, zeroPad, maxNum, showTotal);
        newTF.contents = numStr;
    }
    app.redraw();
}

function main() {
    // ダイアログ作成
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    var prefixGroup = dialog.add("group");
    prefixGroup.orientation = "row";
    prefixGroup.add("statictext", undefined, "接頭辞");
    var prefixField = prefixGroup.add("edittext", undefined, "");
    prefixField.characters = 10;

    var inputGroup = dialog.add("group");
    inputGroup.orientation = "row";
    inputGroup.add("statictext", undefined, LABELS.promptMessage[lang]);
    var inputField = inputGroup.add("edittext", undefined, "1");
    inputField.characters = 5;

    var zeroPadCheckbox = dialog.add("checkbox", undefined, "ゼロパディング");
    var totalPageCheckbox = dialog.add("checkbox", undefined, "総ページ数を表示");

    if (app.documents.length === 0) return;
    var doc = app.activeDocument;

    // _pagenumberレイヤー取得または作成
    var pagenumberLayer;
    try {
        pagenumberLayer = doc.layers.getByName("_pagenumber");
    } catch (e) {
        pagenumberLayer = doc.layers.add();
        pagenumberLayer.name = "_pagenumber";
    }

    // 選択テキストを_pagenumberレイヤーへ移動
    var sel = doc.selection;
    if (sel.length > 0 && sel[0].typename === "TextFrame") {
        sel[0].move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
    }

    // 基準となるテキストフレームを取得
    var targetText = null;
    var abIndexToKeep = 0;
    var baseArtboard = doc.artboards[abIndexToKeep];
    var baseRect = baseArtboard.artboardRect;
    // アートボード上にあるテキストフレームを探す
    for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
        var tf = pagenumberLayer.textFrames[i];
        var pos = tf.position;
        if (pos[0] >= baseRect[0] && pos[0] <= baseRect[2] && pos[1] <= baseRect[1] && pos[1] >= baseRect[3]) {
            targetText = tf;
            break;
        }
    }
    // 1つ目のアートボード以外も検索
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

    // テキスト位置を1アートボード目基準に補正
    var pos = targetText.position;
    var ab = doc.artboards[abIndexToKeep].artboardRect;
    doc.artboards.setActiveArtboardIndex(abIndexToKeep);
    var currentABIndex = getArtboardIndexByPosition(doc, pos);
    if (currentABIndex >= 0) {
        var currentAB = doc.artboards[currentABIndex].artboardRect;
        var offset = [ab[0] - currentAB[0], ab[1] - currentAB[1]];
        targetText.position = [pos[0] + offset[0], pos[1] + offset[1]];
    }

    // レイヤー・型チェック
    if (targetText.layer.name !== "_pagenumber" || targetText.typename !== "TextFrame") {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    // プレビュー更新
    function previewUpdate() {
        var startNum = parseInt(inputField.text, 10);
        if (isNaN(startNum)) return;
        var prefix = prefixField.text;
        var zeroPad = zeroPadCheckbox.value;
        var showTotal = totalPageCheckbox.value;
        targetText.visible = false;
        generatePageNumbers(doc, pagenumberLayer, targetText, baseRect, startNum, prefix, zeroPad, showTotal);
    }

    zeroPadCheckbox.onClick = previewUpdate;
    totalPageCheckbox.onClick = previewUpdate;
    prefixField.onChanging = previewUpdate;
    inputField.onChanging = previewUpdate;

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    var cancelBtn = buttonGroup.add("button", undefined, "Cancel");
    var okBtn = buttonGroup.add("button", undefined, "OK");

    okBtn.onClick = function() {
        if (targetText && !targetText.locked && targetText.editable) {
            targetText.remove();
        }
        dialog.close(1);
    };
    cancelBtn.onClick = function() {
        targetText.visible = true;
        removeOtherTextFrames(pagenumberLayer, targetText);
        dialog.close(0);
    };

    previewUpdate();
    if (dialog.show() !== 1) {
        return;
    }
}

main();