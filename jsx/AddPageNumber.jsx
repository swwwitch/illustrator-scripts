#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
AddPageNumberFromTextSelection.jsx
更新日: 2025-06-27

概要：
「_pagenumber」レイヤーにある選択テキストを基準に、すべてのアートボードに連番テキストを複製します。

処理の流れ：
1. ユーザーに開始番号を入力
2. _pagenumberレイヤーで選択されたテキスト以外を削除
3. 元の段落揃えを記録し、一時的に左揃えに変更
4. 複製時に、最大桁数に応じて左方向に補正（1桁は補正なし）
5. 各アートボードに番号を設定し複製
6. 元のテキストを削除

限定条件：
・選択テキストは_pagenumberレイヤー上のポイントテキストであること
・段落揃えの設定は1段落目の属性に依存
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// -------------------------------
// 日英ラベル定義　Define label
// -------------------------------
var lang = getCurrentLang();
var LABELS = {
    dialogTitle: { ja: "ページ番号を追加", en: "Add Page Numbers" },
    promptMessage: { ja: "開始番号を入力してください", en: "Enter starting number" },
    promptDefault: { ja: "1", en: "1" },
    errorNotNumber: { ja: "有効な数字を入力してください", en: "Please enter a valid number" },
    errorInvalidSelection: { ja: "複製対象のテキストを_pagenumberレイヤーで選択してください", en: "Select text in the _pagenumber layer" },
    errorNotText: { ja: "選択されたオブジェクトがテキストではありません", en: "The selected object is not a text frame" },
    errorLayerHidden: { ja: "_pagenumberレイヤーが非表示またはロックされています", en: "_pagenumber layer is hidden or locked" },
    errorNoText: { ja: "アートボード上に有効なテキストが見つかりませんでした", en: "No valid text found on artboards" }
};

function getTextFrameWidth(textFrame) {
    if (textFrame.textRange.characters.length === 0) {
        return 0;
    }
    var bounds = textFrame.geometricBounds; // [top, left, bottom, right]
    return bounds[3] - bounds[1];
}

function alignTextRightKeepPosition(tf) {
    var originalLeft = tf.left;
    tf.textRange.paragraphAttributes.justification = Justification.RIGHT;
    var delta = originalLeft - tf.left;
    tf.left += delta;
}

function alignTextCenterKeepPosition(tf) {
    var originalLeft = tf.left;
    tf.textRange.paragraphAttributes.justification = Justification.CENTER;
    var delta = originalLeft - tf.left;
    tf.left += delta;
}

function main() {
    var startNum = prompt(LABELS.promptMessage[lang], LABELS.promptDefault[lang]);
    if (startNum === null) return; // キャンセル時は中断
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

    // 選択されたテキストを_pagenumberレイヤーに移動
    if (sel.length > 0 && sel[0].typename === "TextFrame") {
        sel[0].move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
    }

    // _pagenumberレイヤー内で選択中以外のテキストを削除
    for (var i = pagenumberLayer.textFrames.length - 1; i >= 0; i--) {
        var item = pagenumberLayer.textFrames[i];
        if (item !== sel[0] && item.typename === "TextFrame") {
            item.remove();
        }
    }

    // _pagenumberレイヤー制限およびアートボード1への移動処理
    var abIndexToKeep = 0;
    var baseArtboard = doc.artboards[abIndexToKeep];
    var baseRect = baseArtboard.artboardRect;

    var targetText = null;
    for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
        var tf = pagenumberLayer.textFrames[i];
        for (var j = 0; j < doc.artboards.length; j++) {
            var ab = doc.artboards[j].artboardRect;
            var pos = tf.position;
            if (pos[0] >= ab[0] && pos[0] <= ab[2] && pos[1] <= ab[1] && pos[1] >= ab[3]) {
                if (j === abIndexToKeep && !targetText) {
                    targetText = tf;
                }
            }
        }
    }

    // 見つからなければ他アートボードのテキストを順に探す
    if (!targetText) {
        for (var j = 1; j < doc.artboards.length; j++) {
            for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
                var tf = pagenumberLayer.textFrames[i];
                var ab = doc.artboards[j].artboardRect;
                var pos = tf.position;
                if (pos[0] >= ab[0] && pos[0] <= ab[2] && pos[1] <= ab[1] && pos[1] >= ab[3]) {
                    targetText = tf;
                    break;
                }
            }
            if (targetText) break;
        }
    }

    if (targetText) {
        // アートボード1に移動
        var pos = targetText.position;
        var ab = doc.artboards[abIndexToKeep].artboardRect;
        doc.artboards.setActiveArtboardIndex(abIndexToKeep);
        var currentAB = null;
        for (var k = 0; k < doc.artboards.length; k++) {
            var r = doc.artboards[k].artboardRect;
            if (pos[0] >= r[0] && pos[0] <= r[2] && pos[1] <= r[1] && pos[1] >= r[3]) {
                currentAB = r;
                break;
            }
        }
        if (currentAB) {
            var offset = [ab[0] - currentAB[0], ab[1] - currentAB[1]];
            targetText.position = [pos[0] + offset[0], pos[1] + offset[1]];
        }

        // 他のテキストを削除（TextFrame型のみ）
        for (var i = pagenumberLayer.textFrames.length - 1; i >= 0; i--) {
            var item = pagenumberLayer.textFrames[i];
            if (item !== targetText && item.typename === "TextFrame") {
                item.remove();
            }
        }

        sel = [targetText];
    } else {
        alert(LABELS.errorNoText[lang]);
        return;
    }

    // 複数選択されていた場合、_pagenumberレイヤー内のテキストのみ対象とし、最初のひとつだけを残し他は削除
    if (sel.length > 1) {
        var keep = null;
        for (var i = 0; i < sel.length; i++) {
            var item = sel[i];
            if (!keep && item.layer && item.layer.name === "_pagenumber" && item.typename === "TextFrame") {
                keep = item;
            } else if (item.typename === "TextFrame") {
                item.remove();
            }
        }
        if (keep) {
            app.selection = [keep];
            sel = [keep];
        }
    }

    // "_pagenumber"レイヤーを取得
    var layer = doc.layers.getByName("_pagenumber");
    if (!layer.visible || layer.locked) {
        alert(LABELS.errorLayerHidden[lang]);
        return;
    }

    if (sel.length === 0 || sel[0].layer.name !== "_pagenumber") {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    var baseText = sel[0];
    if (baseText.typename !== "TextFrame") {
        alert(LABELS.errorNotText[lang]);
        return;
    }

    // 他のテキストを削除（TextFrame型のみ）
    for (var i = layer.textFrames.length - 1; i >= 0; i--) {
        var item = layer.textFrames[i];
        if (item !== baseText && item.typename === "TextFrame") {
            item.remove();
        }
    }

    var pos = baseText.position;

    // 基準アートボードの情報取得
    var baseIndex = 0;
    var baseRect = doc.artboards[baseIndex].artboardRect;

    var originalJustification = baseText.paragraphs[0].paragraphAttributes.justification;
    var originalTop = baseText.top;
    var originalLeft = baseText.left;
    baseText.paragraphs[0].paragraphAttributes.justification = Justification.LEFT;
    baseText.top = originalTop;
    baseText.left = originalLeft;

    // 左補正を計算
    var tempText = baseText.duplicate();
    tempText.contents = "0";
    var textWidth = getTextFrameWidth(tempText);
    tempText.remove();

    var totalCount = startNum + doc.artboards.length - 1;
    var offsetX = 0;
    if (totalCount >= 100) {
        if (originalJustification === Justification.LEFT) {
            offsetX = -2 * textWidth;
        } else if (originalJustification === Justification.CENTER) {
            offsetX = -1 * textWidth;
        } else if (originalJustification === Justification.RIGHT) {
            offsetX = 0.5 * textWidth;
        }
    } else if (totalCount >= 10) {
        if (originalJustification === Justification.LEFT) {
            offsetX = -1 * textWidth;
        } else if (originalJustification === Justification.CENTER) {
            offsetX = -0.5 * textWidth;
        } else if (originalJustification === Justification.RIGHT) {
            offsetX = 0;
        }
    } else {
        offsetX = 0;
    }

    for (var i = 0; i < doc.artboards.length; i++) {
        var abRect = doc.artboards[i].artboardRect;
        var dx = abRect[0];
        var dy = abRect[1];
        var abPos = [
            dx + (pos[0] - baseRect[0]),
            dy + (pos[1] - baseRect[1])
        ];

        var dup = baseText.duplicate(layer, ElementPlacement.PLACEATBEGINNING);
        dup.contents = String(startNum + i);
        dup.position = [abPos[0] + offsetX, abPos[1]];
        dup.paragraphs[0].paragraphAttributes.justification = originalJustification;
    }
    
    if (originalJustification === Justification.RIGHT) {
        alignTextRightKeepPosition(baseText);
    } else if (originalJustification === Justification.CENTER) {
        alignTextCenterKeepPosition(baseText);
    } else {
        baseText.paragraphs[0].paragraphAttributes.justification = originalJustification;
    }

    baseText.remove();
}

main();