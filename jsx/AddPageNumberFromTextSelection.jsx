#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
### スクリプト名：

RenameArtboardsPlus.jsx

### 概要

- アートボード名を一括で柔軟に変更できるIllustrator用スクリプトです。
- 接頭辞、接尾辞、ファイル名、番号、アートボード名を自由に組み合わせて、高度な命名ルールに対応します。

### 主な機能

- 接頭辞・接尾辞・アートボード名の有無を自由に選択可能
- ファイル名の参照、区切り文字設定、プリセット保存・読み込みに対応
- 連番形式は「数字」「アルファベット（大文字／小文字）」を選択可能
- 開始番号から桁数（ゼロ埋め）を自動判定
- プレビュー機能で名称例を事前確認
- 最大15件のプレビュー表示
- OK／適用／キャンセルボタンを美しく配置
- ExtendScript（ES3）互換

### 処理の流れ

1. ダイアログで接頭辞、接尾辞、番号形式などを設定
2. プレビューで名称例を確認
3. 「適用」または「OK」をクリックしてアートボード名を一括変更

### 更新履歴

- v1.0.0 (20250420) : 初期バージョン作成
- v1.0.1 (20250430) : 開始番号から桁数自動判定を追加、プリセットlabel簡素化、ES3対応強化

---

### Script Name:

RenameArtboardsPlus.jsx

### Overview

- An Illustrator script to flexibly batch rename artboards.
- Allows combining prefixes, suffixes, file names, numbers, and artboard names to support advanced naming rules.

### Main Features

- Freely choose whether to include prefix, suffix, or original artboard name
- Supports file name reference, custom separators, preset save/load
- Numbering formats: "Number", "Alphabet (Uppercase)", or "Alphabet (Lowercase)"
- Automatically detect padding digits from start number (zero-padding)
- Preview feature to check example names beforehand
- Displays up to 15 preview items
- Beautifully aligned OK / Apply / Cancel buttons
- Compatible with ExtendScript (ES3)

### Process Flow

1. Configure prefix, suffix, numbering format, etc., in the dialog
2. Check example names using the preview
3. Click "Apply" or "OK" to batch rename artboards

### Update History

- v1.0.0 (20250420): Initial version created
- v1.0.1 (20250430): Added auto-detection of padding digits from start number, simplified preset label, improved ES3 support
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
    },
    prefixLabel: {
        ja: "接頭辞",
        en: "Prefix"
    },
    suffixLabel: {
        ja: "接尾辞",
        en: "Suffix"
    },
    zeroPadLabel: {
        ja: "ゼロ埋め",
        en: "Zero pad"
    },
    totalPageLabel: {
        ja: "総ページ",
        en: "Show total pages"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
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

// テキスト複製＆内容更新（再利用性向上のため関数化）
function generatePageNumbers(doc, pagenumberLayer, targetText, baseRect, startNum, prefix, zeroPad, showTotal, customDigits, suffix) {
    removeOtherTextFrames(pagenumberLayer, targetText);
    var abCount = doc.artboards.length;
    var maxNum = startNum + abCount - 1;
    var maxDigits = customDigits ? customDigits : String(maxNum).length;
    for (var i = 0; i < abCount; i++) {
        var abRect = doc.artboards[i].artboardRect;
        var newTF = targetText.duplicate(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
        newTF.position = [
            targetText.position[0] + (abRect[0] - baseRect[0]),
            targetText.position[1] + (abRect[1] - baseRect[1])
        ];
        var numStr = buildPageNumberString(startNum + i, maxDigits, prefix, zeroPad, maxNum, showTotal);
        numStr += suffix;
        newTF.contents = numStr;
    }
    app.redraw();
}

function main() {
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

    // 1つ目のアートボード上にあるテキストフレームを探す
    for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
        var tf = pagenumberLayer.textFrames[i];
        var pos = tf.position;
        if (pos[0] >= baseRect[0] && pos[0] <= baseRect[2] && pos[1] <= baseRect[1] && pos[1] >= baseRect[3]) {
            targetText = tf;
            break;
        }
    }
    // 見つからなければ他のアートボードも探索
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

    // ダイアログ作成
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    // 親グループ（横並び）
    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = "top";

    // 左カラム: 接頭辞
    var leftGroup = columnsGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = "left";
    leftGroup.add("statictext", undefined, LABELS.prefixLabel[lang]);
    var prefixField = leftGroup.add("edittext", undefined, "");
    prefixField.characters = 10;

    // 中央カラム: 開始番号 + ゼロ埋め
    var centerGroup = columnsGroup.add("group");
    centerGroup.orientation = "column";
    centerGroup.alignChildren = "left";
    var inputGroup = centerGroup.add("group");
    inputGroup.orientation = "column";
    inputGroup.add("statictext", undefined, LABELS.promptMessage[lang]);
    var inputField = inputGroup.add("edittext", undefined, "1");
    inputField.characters = 5;
    var zeroPadCheckbox = centerGroup.add("checkbox", undefined, LABELS.zeroPadLabel[lang]);

    // 右カラム: 接尾辞 + 総ページ数
    var rightGroup = columnsGroup.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = "left";
    rightGroup.add("statictext", undefined, LABELS.suffixLabel[lang]);
    var suffixField = rightGroup.add("edittext", undefined, "");
    suffixField.characters = 10;
    var totalPageCheckbox = rightGroup.add("checkbox", undefined, LABELS.totalPageLabel[lang]);

    // プレビュー更新関数
    function previewUpdate() {
        var startNumStr = inputField.text;
        var startNum = parseInt(startNumStr, 10);
        if (isNaN(startNum)) return;
        var prefix = prefixField.text;
        var zeroPad = zeroPadCheckbox.value;
        var showTotal = totalPageCheckbox.value;
        var suffix = suffixField.text;

        // 入力がゼロ始まりの場合、桁数を取得しゼロパディングを強制ON
        var customDigits = 0;
        if (startNumStr.match(/^0/)) {
            customDigits = startNumStr.length;
            zeroPad = true;
        }

        targetText.visible = false;
        generatePageNumbers(doc, pagenumberLayer, targetText, baseRect, startNum, prefix, zeroPad, showTotal, customDigits, suffix);
    }

    zeroPadCheckbox.onClick = previewUpdate;
    totalPageCheckbox.onClick = previewUpdate;
    prefixField.onChanging = previewUpdate;
    inputField.onChanging = previewUpdate;
    suffixField.onChanging = previewUpdate;

    // ボタン作成
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang]);

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