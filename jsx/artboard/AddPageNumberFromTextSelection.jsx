#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AddPageNumberFromTextSelection.jsx

### 概要

- _pagenumberレイヤーで選択したテキストを基準に、すべてのアートボードにページ番号テキストを複製・配置するIllustrator用スクリプトです。
- 開始番号、接頭辞、接尾辞、ゼロ埋め、総ページ数表示のカスタマイズが可能です。

### 主な機能

- 選択テキストを基にページ番号を生成
- 開始番号、接頭辞、接尾辞の指定
- ゼロパディング（ゼロ埋め）対応
- 総ページ数の表示オプション
- プレビュー機能
- 日本語／英語インターフェース対応

### 処理の流れ

1. _pagenumberレイヤーにテキストを選択
2. ダイアログで開始番号や接頭辞などを設定
3. プレビューを確認
4. OKでページ番号テキストを全アートボードに配置

### 更新履歴

- v1.0 (20240401) : 初期バージョン
- v1.1 (20240405) : テキスト複製ロジック修正
- v1.2 (20240410) : ゼロ埋め・接頭辞・総ページ数表示追加
- v1.3 (20240415) : プレビュー機能追加
- v1.4 (20240420) : 「001」形式のゼロ埋め対応
- v1.5 (20240425) : 接尾辞フィールド追加、UI改善
- v1.6 (20250810) : プレビュー開始時に元テキストを一時待避して非表示ではなく削除。キャンセル時は復元、OKで待避を破棄。
- v1.7 (20250815) : 開始番号を↑↓キーで変更可能に（Shiftで±10、Optionは未対応）

---

### Script Name:

AddPageNumberFromTextSelection.jsx

### Overview

- An Illustrator script to duplicate and place page number text on all artboards using selected text in the _pagenumber layer as a reference.
- Supports customizing starting number, prefix, suffix, zero-padding, and total page display.

### Main Features

- Generate page numbers based on selected text
- Specify starting number, prefix, and suffix
- Supports zero padding
- Option to show total pages
- Preview feature
- Japanese and English UI support

### Process Flow

1. Select text in the _pagenumber layer
2. Configure starting number, prefix, etc., in dialog
3. Check the preview
4. Click OK to place page numbers on all artboards

### Update History

- v1.0 (20240401): Initial version
- v1.1 (20240405): Fixed text duplication logic
- v1.2 (20240410): Added zero padding, prefix, and total page display
- v1.3 (20240415): Added preview feature
- v1.4 (20240420): Supported "001"-style zero padding
- v1.5 (20240425): Added suffix field, improved UI
- v1.6 (20250810): Remove original selection at preview start (backed up to temp layer). Restore on Cancel, discard on OK.
- v1.7 (20250815): Enabled changing starting number with Up/Down keys (Shift for ±10, no Option support)
*/

var SCRIPT_VERSION = "v1.7";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var TMP_LAYER_NAME = "_pagenumber_backup";
var backupTF = null; // 復元用の待避コピー / backup clone for cancel-restore
var originalRemoved = false; // プレビュー開始時に元テキストをremoveしたか / removed original at preview?

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    dialogTitle: {
        ja: "数値入力（上下キー対応） " + SCRIPT_VERSION,
        en: "Numeric Input (Arrow Key Supported) " + SCRIPT_VERSION
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
    },
    noDocument: {
        ja: "ドキュメントを開いてください。",
        en: "Please open a document."
    },
    invalidNumber: {
        ja: "開始番号に有効な数値を入力してください。",
        en: "Please enter a valid number for the starting number."
    }
};

/* ↑↓で値を変更（Shiftで±10、Optionは未対応） / Change value with arrow keys (Shift=±10, no Option) */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        var text = editText.text;
        var hadLeadingZeros = (text && text.length > 1 && text.charAt(0) === "0"); // 先頭ゼロ保持
        var digits = hadLeadingZeros ? text.length : 0;
        var value = Number(text);
        if (isNaN(value)) { value = 0; }

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.shiftKey ? 10 : 1;
        var handled = false;

        if (event.keyName === "Up") {
            value += delta;
            handled = true;
        } else if (event.keyName === "Down") {
            value -= delta;
            if (value < 0) { value = 0; }
            handled = true;
        }

        if (handled) {
            // 整数に丸め / Round to integer
            value = Math.round(value);

            // 先頭ゼロの桁数を維持 / Keep digit width when had leading zeros
            var out = String(value);
            if (digits > 0) {
                while (out.length < digits) { out = "0" + out; }
            }
            editText.text = out;
            // onChanging を通知して双方向で更新 / Notify onChanging to keep behaviors consistent
            try { if (editText.notify) { editText.notify('onChanging'); } } catch (eN) {}
            // 即時プレビュー更新 / Immediate preview update
            try { previewUpdate(); } catch (e2) {}
            try { app.redraw(); } catch (e3) {}

            // 既定動作抑止（キャレット移動防止） / Suppress default behavior
            if (event.preventDefault) { event.preventDefault(); }
            event.returnValue = false;
        }
    });
    // 念のため keyup でも更新 / Also update on keyup as a fallback
    editText.addEventListener("keyup", function (ev) {
        if (ev.keyName === "Up" || ev.keyName === "Down") {
            try { previewUpdate(); } catch (e) {}
            if (ev.preventDefault) { ev.preventDefault(); }
            ev.returnValue = false;
        }
    });
}

/* 指定レイヤー上の他のテキストフレームを削除（exceptを除く） / Remove other text frames on specified layer except 'except' */
function removeOtherTextFrames(layer, except) {
    for (var i = layer.textFrames.length - 1; i >= 0; i--) {
        var item = layer.textFrames[i];
        if (item !== except && item.typename === "TextFrame") {
            item.remove();
        }
    }
}

/* 座標からアートボードインデックスを取得 / Get artboard index by position */
function getArtboardIndexByPosition(doc, pos) {
    for (var i = 0; i < doc.artboards.length; i++) {
        var abRect = doc.artboards[i].artboardRect;
        if (pos[0] >= abRect[0] && pos[0] <= abRect[2] && pos[1] <= abRect[1] && pos[1] >= abRect[3]) {
            return i;
        }
    }
    return -1;
}

/* ページ番号文字列生成 / Build page number string */
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

/* テキスト複製＆内容更新（再利用性向上のため関数化） / Duplicate text & update contents (function for reusability) */
function generatePageNumbers(doc, pagenumberLayer, targetText, baseRect, startNum, prefix, zeroPad, showTotal, customDigits, suffix) {
    removeOtherTextFrames(pagenumberLayer, targetText);
    var abCount = doc.artboards.length;
    var maxNum = startNum + abCount - 1;
    var maxDigits = customDigits ? customDigits : String(maxNum).length;
    for (var i = 0; i < abCount; i++) {
        var abRect = doc.artboards[i].artboardRect;
        var newTF = targetText.duplicate(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
        // 複製物は必ず編集可能に / Ensure duplicates are editable
        try { newTF.locked = false; } catch (e) {}
        try { newTF.editable = true; } catch (e) {}
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
    if (app.documents.length === 0) {
        alert(LABELS.noDocument[lang]);
        return;
    }
    var doc = app.activeDocument;

    /* _pagenumberレイヤー取得または作成 / Get or create _pagenumber layer */
    var pagenumberLayer;
    try {
        pagenumberLayer = doc.layers.getByName("_pagenumber");
    } catch (e) {
        pagenumberLayer = doc.layers.add();
        pagenumberLayer.name = "_pagenumber";
    }

    /* 選択テキストを_pagenumberレイヤーへ移動 / Move selected text to _pagenumber layer */
    var sel = doc.selection;
    if (sel.length > 0 && sel[0].typename === "TextFrame") {
        sel[0].move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
    }

    /* 基準となるテキストフレームを取得 / Get reference text frame */
    var targetText = null;
    var abIndexToKeep = 0;
    var baseArtboard = doc.artboards[abIndexToKeep];
    var baseRect = baseArtboard.artboardRect;

    /* 1つ目のアートボード上にあるテキストフレームを探す / Search for text frame on first artboard */
    for (var i = 0; i < pagenumberLayer.textFrames.length; i++) {
        var tf = pagenumberLayer.textFrames[i];
        var pos = tf.position;
        if (pos[0] >= baseRect[0] && pos[0] <= baseRect[2] && pos[1] <= baseRect[1] && pos[1] >= baseRect[3]) {
            targetText = tf;
            break;
        }
    }
    /* 見つからなければ他のアートボードも探索 / If not found, search other artboards */
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

    /* テキスト位置を1アートボード目基準に補正 / Adjust text position relative to first artboard */
    var pos = targetText.position;
    var ab = doc.artboards[abIndexToKeep].artboardRect;
    doc.artboards.setActiveArtboardIndex(abIndexToKeep);
    var currentABIndex = getArtboardIndexByPosition(doc, pos);
    if (currentABIndex >= 0) {
        var currentAB = doc.artboards[currentABIndex].artboardRect;
        var offset = [ab[0] - currentAB[0], ab[1] - currentAB[1]];
        targetText.position = [pos[0] + offset[0], pos[1] + offset[1]];
    }

    /* レイヤー・型チェック / Layer and type check */
    if (targetText.layer.name !== "_pagenumber" || targetText.typename !== "TextFrame") {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    /* ダイアログ作成 / Create dialog */
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    /* 親グループ（横並び） / Parent group (horizontal layout) */
    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = "top";

    /* 左カラム: 接頭辞 / Left column: Prefix */
    var leftGroup = columnsGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = "left";
    leftGroup.add("statictext", undefined, LABELS.prefixLabel[lang]);
    var prefixField = leftGroup.add("edittext", undefined, "");
    prefixField.characters = 10;

    /* 中央カラム: 開始番号 + ゼロ埋め / Center column: Starting number + Zero pad */
    var centerGroup = columnsGroup.add("group");
    centerGroup.orientation = "column";
    centerGroup.alignChildren = "left";
    var inputGroup = centerGroup.add("group");
    inputGroup.orientation = "column";
    inputGroup.add("statictext", undefined, LABELS.promptMessage[lang]);
    var inputField = inputGroup.add("edittext", undefined, "1");
    inputField.characters = 5;
    // ↑↓/Shift+↑↓ 対応（Optionは未実装） / Enable Up/Down and Shift+Up/Down (no Option)
    changeValueByArrowKey(inputField);
    var zeroPadCheckbox = centerGroup.add("checkbox", undefined, LABELS.zeroPadLabel[lang]);

    /* 右カラム: 接尾辞 + 総ページ数 / Right column: Suffix + Total pages */
    var rightGroup = columnsGroup.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = "left";
    rightGroup.add("statictext", undefined, LABELS.suffixLabel[lang]);
    var suffixField = rightGroup.add("edittext", undefined, "");
    suffixField.characters = 10;
    var totalPageCheckbox = rightGroup.add("checkbox", undefined, LABELS.totalPageLabel[lang]);

    /* プレビュー更新関数 / Preview update function */
    function previewUpdate() {
        var startNumStr = inputField.text;
        var startNum = parseInt(startNumStr, 10);
        if (isNaN(startNum)) {
            alert(LABELS.invalidNumber[lang]);
            return;
        }
        var prefix = prefixField.text;
        var zeroPad = zeroPadCheckbox.value;
        var showTotal = totalPageCheckbox.value;
        var suffix = suffixField.text;

        /* 入力がゼロ始まりの場合、桁数を取得しゼロパディングを強制ON / If input starts with zero, get digit count and force zero padding ON */
        var customDigits = 0;
        if (startNumStr.match(/^0/)) {
            customDigits = startNumStr.length;
            zeroPad = true;
        }

        // プレビュー開始時に元テキストを待避して削除 / On first preview, backup & remove original
        if (!originalRemoved) {
            // 一時レイヤー取得 / Get or create temp layer
            var tmpLayer;
            try {
                tmpLayer = doc.layers.getByName(TMP_LAYER_NAME);
            } catch (e) {
                tmpLayer = doc.layers.add();
                tmpLayer.name = TMP_LAYER_NAME;
            }
            // 待避コピー作成（位置・外観を維持）/ Make backup clone (keeps position/appearance)
            backupTF = targetText.duplicate(tmpLayer, ElementPlacement.PLACEATBEGINNING);
            backupTF.locked = true;
            backupTF.visible = false;
            // 元テキスト削除 / Remove original
            if (!targetText.locked && targetText.editable) {
                try { targetText.remove(); } catch (e2) {}
            }
            originalRemoved = true;
            try { app.redraw(); } catch (e3) {}
        }
        // 以降は待避コピーをテンプレートに使用 / Use backup clone as template
        var templateTF = originalRemoved && backupTF ? backupTF : targetText;
        // 非表示化（保険）/ Ensure hidden (just in case)
        if (templateTF) { templateTF.visible = false; }
        generatePageNumbers(doc, pagenumberLayer, templateTF, baseRect, startNum, prefix, zeroPad, showTotal, customDigits, suffix);
        try { app.redraw(); } catch (e4) {}
    }

    zeroPadCheckbox.onClick = previewUpdate;
    totalPageCheckbox.onClick = previewUpdate;
    prefixField.onChanging = previewUpdate;
    inputField.onChanging = previewUpdate;
    suffixField.onChanging = previewUpdate;

    /* ボタン作成 / Create buttons */
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang]);

    okBtn.onClick = function() {
        // 念のため最終プレビューを反映 / Ensure latest preview before commit
        try { previewUpdate(); } catch (e0) {}
        // 元テキストはプレビュー時に削除済み。待避があれば破棄 / Original already removed at preview; discard backup if exists
        if (backupTF) {
            try { backupTF.locked = false; backupTF.remove(); } catch (e) {}
            backupTF = null;
        }
        // 空の一時レイヤーを削除 / Clean up temp layer if empty
        try {
            var tmp = doc.layers.getByName(TMP_LAYER_NAME);
            if (tmp.pageItems.length === 0) { tmp.remove(); }
        } catch (e2) {}
        dialog.close(1);
    };
    cancelBtn.onClick = function() {
        // 生成物をクリア / Clear generated numbers
        if (pagenumberLayer && targetText) {
            removeOtherTextFrames(pagenumberLayer, null);
        } else {
            removeOtherTextFrames(pagenumberLayer, null);
        }
        // 復元処理 / Restore original if we removed it
        if (originalRemoved && backupTF) {
            try {
                backupTF.locked = false;
                backupTF.visible = true;
                backupTF.move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING);
                // 復元後、参照をtargetTextへ差し替え / Hand back to targetText
                targetText = backupTF;
                backupTF = null;
            } catch (e) {}
        }
        // 空の一時レイヤーを削除 / Remove temp layer if empty
        try {
            var tmp = doc.layers.getByName(TMP_LAYER_NAME);
            if (tmp.pageItems.length === 0) { tmp.remove(); }
        } catch (e) {}
        dialog.close(0);
    };

    // ダイアログ表示時に［開始番号］へフォーカス / Focus starting number on dialog show
    dialog.onShow = function () {
        try { inputField.active = true; } catch (e) {}
    };
    previewUpdate();
    if (dialog.show() !== 1) {
        return;
    }
}

main();