#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中の2つのテキストオブジェクトの内容を入れ替えます。
ダイアログで入れ替える対象（文字列 / スタイル / 座標）を選択できます。

- 選択は2つ、かつ両方ともテキストオブジェクトである必要があります。
- 条件を満たさない場合はアラートを表示します。

### Overview

Swaps the contents of two selected text objects.
A dialog lets you choose what to swap (string / style / position).

- Exactly two objects must be selected, and both must be text objects.
- Shows an alert if the conditions are not met.

### 紹介記事

https://note.com/dtp_tranist/n/n071e09af28a7

*/

var SCRIPT_VERSION = "v1.0.0";

// =============================================================
// ローカライズ / Localization
// =============================================================

var lang = (function () {
    /* 日本語 / English */
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
})();

function L(obj) {
    return obj[lang] || obj.en;
}

var LABELS = {
    dialogTitle: { ja: "テキストの入れ替え " + SCRIPT_VERSION, en: "Swap Text " + SCRIPT_VERSION },
    panelTarget: { ja: "入れ替える対象", en: "Swap target" },
    modeContents: { ja: "文字列", en: "String" },
    modeFormat: { ja: "書式", en: "Format" },
    modePosition: { ja: "座標", en: "Position" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    needTwo: { ja: "テキストオブジェクトを2つ選択してください。", en: "Please select two text objects." },
    needText: { ja: "選択した2つは両方ともテキストオブジェクトである必要があります。", en: "Both selected objects must be text objects." }
};

// =============================================================
// ダイアログ / Dialog
// =============================================================

function showSwapDialog() {
    var dialog = new Window("dialog", L(LABELS.dialogTitle));
    dialog.alignChildren = "fill";

    var targetPanel = dialog.add("panel", undefined, L(LABELS.panelTarget));
    targetPanel.orientation = "column";
    targetPanel.alignChildren = "left";
    targetPanel.margins = [15, 20, 15, 15];

    var radioContents = targetPanel.add("radiobutton", undefined, L(LABELS.modeContents));
    var radioFormat = targetPanel.add("radiobutton", undefined, L(LABELS.modeFormat));
    var radioPosition = targetPanel.add("radiobutton", undefined, L(LABELS.modePosition));
    radioContents.value = true;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var cancelButton = buttonGroup.add("button", undefined, L(LABELS.cancel), { name: "cancel" });
    var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (dialog.show() !== 1) {
        return null;
    }

    if (radioFormat.value) return "format";
    if (radioPosition.value) return "position";
    return "contents";
}

// =============================================================
// 入れ替え処理 / Swap operations
// =============================================================

function swapContents(firstTextFrame, secondTextFrame) {
    var firstContents = firstTextFrame.contents;
    var secondContents = secondTextFrame.contents;
    firstTextFrame.contents = secondContents;
    secondTextFrame.contents = firstContents;
}

/*
   入れ替える書式（characterAttributes）の一覧。
   List of character attributes to swap. textFont はフォント＋スタイルを兼ねる。
*/
var FORMAT_ATTRIBUTE_KEYS = [
    "textFont",        /* フォント＋スタイル / font family + style */
    "size",            /* サイズ / size */
    "fillColor",       /* 文字カラー / text color */
    "strokeColor",     /* 線カラー / stroke color */
    "strokeWeight",    /* 線幅 / stroke weight */
    "tracking",        /* トラッキング / tracking */
    "leading",         /* 行送り / leading */
    "autoLeading",     /* 自動行送り / auto leading */
    "horizontalScale", /* 水平比率 / horizontal scale */
    "verticalScale",   /* 垂直比率 / vertical scale */
    "baselineShift",   /* ベースラインシフト / baseline shift */
    "capitalization"   /* 大文字小文字 / capitalization */
];

function captureFormatAttributes(textFrame) {
    var attributes = textFrame.textRange.characterAttributes;
    var captured = {};
    for (var i = 0; i < FORMAT_ATTRIBUTE_KEYS.length; i++) {
        var key = FORMAT_ATTRIBUTE_KEYS[i];
        try {
            captured[key] = attributes[key];
        } catch (e) {}
    }
    return captured;
}

function applyFormatAttributes(textFrame, captured) {
    var attributes = textFrame.textRange.characterAttributes;
    for (var i = 0; i < FORMAT_ATTRIBUTE_KEYS.length; i++) {
        var key = FORMAT_ATTRIBUTE_KEYS[i];
        if (!captured.hasOwnProperty(key)) continue;
        try {
            attributes[key] = captured[key];
        } catch (e) {}
    }
}

function swapFormat(firstTextFrame, secondTextFrame) {
    /* 両方の書式を先に取得してから入れ替える / Capture both before applying */
    var firstAttributes = captureFormatAttributes(firstTextFrame);
    var secondAttributes = captureFormatAttributes(secondTextFrame);
    applyFormatAttributes(firstTextFrame, secondAttributes);
    applyFormatAttributes(secondTextFrame, firstAttributes);
}

function swapPosition(firstTextFrame, secondTextFrame) {
    /*
       position はベースライン基準で上端/左端が崩れるため geometricBounds を使う。
       Use geometricBounds (not position) because TextFrame.position is baseline-based.
       geometricBounds = [left, top, right, bottom]
    */
    app.redraw(); // bounds が更新されない環境対策 / refresh stale bounds

    var firstBounds = firstTextFrame.geometricBounds;
    var secondBounds = secondTextFrame.geometricBounds;

    var deltaX = secondBounds[0] - firstBounds[0];
    var deltaY = secondBounds[1] - firstBounds[1];

    firstTextFrame.translate(deltaX, deltaY);
    secondTextFrame.translate(-deltaX, -deltaY);
}

// =============================================================
// メイン / Main
// =============================================================

function main() {
    if (app.documents.length === 0) {
        alert(L(LABELS.noDocument));
        return;
    }

    var doc = app.activeDocument;
    var selectedItems = doc.selection;

    if (selectedItems.length !== 2) {
        alert(L(LABELS.needTwo));
        return;
    }
    if (selectedItems[0].typename !== "TextFrame" || selectedItems[1].typename !== "TextFrame") {
        alert(L(LABELS.needText));
        return;
    }

    var mode = showSwapDialog();
    if (mode === null) {
        return;
    }

    var firstTextFrame = selectedItems[0];
    var secondTextFrame = selectedItems[1];

    if (mode === "format") {
        swapFormat(firstTextFrame, secondTextFrame);
    } else if (mode === "position") {
        swapPosition(firstTextFrame, secondTextFrame);
    } else {
        swapContents(firstTextFrame, secondTextFrame);
    }
}

main();
