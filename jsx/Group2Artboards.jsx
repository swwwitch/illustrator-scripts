#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
Group2Artboards.jsx

【日本語説明】
選択したグループオブジェクトの境界に指定したマージンを加え、その範囲をアートボードとして自動追加するIllustrator用スクリプトです。複数グループ選択に対応し、アートボード名の自動連番やファイル名参照、既存アートボードの削除も行えます。

【English Description】
This script creates new artboards in Adobe Illustrator around each selected group object, expanding their bounds by a user-specified margin. It supports multiple group selections, automatic sequential naming with optional file name prefix and custom symbols, and can delete existing artboards before adding new ones.

【更新履歴 / Change Log】
- v1.0.1 (2025-07-04) コメント整理と最適化 / Cleaned comments, optimized logic

【制限事項 / Limitations】
- グループ以外のオブジェクトは対象外です。Only group objects are processed.
- マージンは数値入力のみ対応。Margin value must be a number.
- Illustrator専用。For Adobe Illustrator only.
*/

// 単位コードから単位ラベルを取得
var unitLabelMap = {
  0: "in",
  1: "mm",
  2: "pt",
  3: "pica",
  4: "cm",
  5: "Q/H",
  6: "px",
  7: "ft/in",
  8: "m",
  9: "yd",
  10: "ft"
};

function getUnitLabel(code, prefKey) {
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitLabelMap[code] || "不明";
}

// 現在の言語を判定（日本語か英語）
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

// UIラベル（表示順に並べる）
var LABELS = {
    artboardPanel: { ja: "アートボード化", en: "Artboard" },
    previewBounds: { ja: "プレビュー境界", en: "Preview bounds" },
    margin: { ja: "マージン", en: "Margin" },
    deleteArtboards: { ja: "既存のアートボードを削除", en: "Delete existing artboards" },
    namePanel: { ja: "アートボード名", en: "Artboard Name" },
    useFileName: { ja: "ファイル名を参照", en: "Use file name" },
    prefix: { ja: "接頭辞", en: "Prefix" },
    symbol: { ja: "記号", en: "Symbol" },
    dash: { ja: "-", en: "-" },
    underscore: { ja: "_", en: "_" },
    none: { ja: "なし", en: "None" },
    startNumber: { ja: "開始番号", en: "Start Number" },
    zeroPadding: { ja: "ゼロ埋め", en: "Zero Padding" },
    example: { ja: "例：", en: "Example: " },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" }
};

// アートボード名を生成する共通関数
function buildArtboardName(prefix, symbol, seq, zeroPadding, useFileName, fileNameNoExt, padLen) {
    var seqNum = parseInt(seq, 10);
    if (isNaN(seqNum)) seqNum = 1;
    var numStr = seqNum.toString();
    if (zeroPadding && padLen > 0) {
        while (numStr.length < padLen) numStr = "0" + numStr;
    }
    if (useFileName) {
        if (prefix === "") {
            return fileNameNoExt + symbol + numStr;
        } else {
            return fileNameNoExt + symbol + prefix + symbol + numStr;
        }
    }
    return prefix + symbol + numStr;
}

function showDialog() {
    var dialog = new Window("dialog", LABELS.artboardPanel[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = [15, 20, 15, 10];
    dialog.spacing = 10;

    var rulerUnit = getUnitLabel(app.preferences.getIntegerPreference("rulerType"), "rulerType");

    var controlGroup = dialog.add("group");
    controlGroup.orientation = "column";
    controlGroup.alignChildren = "left";
    controlGroup.margins = [15,5, 15, 10];
    controlGroup.spacing = 10;

    var previewBoundsCheck = controlGroup.add("checkbox", undefined, LABELS.previewBounds[lang]);
    previewBoundsCheck.value = true;

    var marginGroup = controlGroup.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = "center";
    marginGroup.add("statictext", undefined, LABELS.margin[lang]);
    var marginInput = marginGroup.add("edittext", undefined, "0");
    marginInput.characters = 5;
    marginGroup.add("statictext", undefined, rulerUnit);

    var deleteArtboardsCheck = controlGroup.add("checkbox", undefined, LABELS.deleteArtboards[lang]);
    deleteArtboardsCheck.value = true;

    // アートボード名パネル
    var namePanel = dialog.add("panel");
    namePanel.text = LABELS.namePanel[lang];
    namePanel.orientation = "row";
    namePanel.alignChildren = "center";
    namePanel.margins = [15,25, 15, 10];

    var nameGroup = namePanel.add("group");
    nameGroup.orientation = "column";
    nameGroup.alignChildren = "left";

    var useFileNameCheck = nameGroup.add("checkbox", undefined, LABELS.useFileName[lang]);
    useFileNameCheck.value = false;

    var prefixRow = nameGroup.add("group");
    prefixRow.orientation = "row";
    prefixRow.alignChildren = "center";
    prefixRow.add("statictext", undefined, LABELS.prefix[lang]);
    var nameInput = prefixRow.add("edittext", undefined, "");
    nameInput.characters = 15;

    var symbolGroup = nameGroup.add("group");
    symbolGroup.orientation = "row";
    symbolGroup.alignChildren = "center";
    symbolGroup.add("statictext", undefined, LABELS.symbol[lang]);
    var radioDash = symbolGroup.add("radiobutton", undefined, LABELS.dash[lang]);
    var radioUnderscore = symbolGroup.add("radiobutton", undefined, LABELS.underscore[lang]);
    var radioNone = symbolGroup.add("radiobutton", undefined, LABELS.none[lang]);
    radioDash.value = true;

    var seqRow = nameGroup.add("group");
    seqRow.orientation = "row";
    seqRow.alignChildren = "center";
    seqRow.add("statictext", undefined, LABELS.startNumber[lang]);
    var seqInput = seqRow.add("edittext", undefined, "01");
    seqInput.characters = 5;
    var zeroPaddingCheck = seqRow.add("checkbox", undefined, LABELS.zeroPadding[lang]);
    zeroPaddingCheck.value = true;

    var previewText = nameGroup.add("statictext", undefined, "");
    previewText.alignment = "left";
    previewText.characters = 20;

    // プレビュー更新
    function updatePreview() {
        var prefix = nameInput.text;
        var symbol = radioDash.value ? "-" : (radioUnderscore.value ? "_" : "");
        var seq = seqInput.text;
        var zeroPadding = zeroPaddingCheck.value;
        var fileNameNoExt = "";
        if (useFileNameCheck.value && app && app.activeDocument && app.activeDocument.name) {
            var docName = app.activeDocument.name;
            var lastDot = docName.lastIndexOf(".");
            fileNameNoExt = lastDot > 0 ? docName.substring(0, lastDot) : docName;
        }
        previewText.text = LABELS.example[lang] + buildArtboardName(prefix, symbol, seq, zeroPadding, useFileNameCheck.value, fileNameNoExt, seq.length);
    }

    // イベント登録
    nameInput.onChanging = updatePreview;
    radioDash.onClick = updatePreview;
    radioUnderscore.onClick = updatePreview;
    radioNone.onClick = updatePreview;
    seqInput.onChanging = updatePreview;
    zeroPaddingCheck.onClick = updatePreview;
    useFileNameCheck.onClick = updatePreview;
    updatePreview();

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";
    buttonGroup.margins = [0, 10, 0, 10];
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    var dialogResult = null;
    cancelBtn.onClick = function() { dialog.close(); };
    okBtn.onClick = function() {
        var selectedSymbol = radioDash.value ? "-" : (radioUnderscore.value ? "_" : "");
        dialogResult = {
            marginValue: marginInput.text,
            deleteArtboards: deleteArtboardsCheck.value,
            usePreviewBounds: previewBoundsCheck.value,
            artboardName: nameInput.text,
            sequentialText: seqInput.text,
            zeroPadding: zeroPaddingCheck.value,
            symbol: selectedSymbol,
            useFileName: useFileNameCheck.value
        };
        dialog.close(1);
    };
    if (dialog.show() !== 1) return null;
    return dialogResult;
}

function main() {
    if (!app.documents.length) return;
    var selection = app.activeDocument.selection;
    if (!selection || selection.length === 0) return;

    var dialogResult = showDialog();
    if (!dialogResult) return;

    var doc = app.activeDocument;
    var margin = parseFloat(dialogResult.marginValue);
    if (isNaN(margin)) margin = 0;
    var initialCount = doc.artboards.length;

    var seqText = dialogResult.sequentialText;
    var seqNum = parseInt(seqText, 10);
    if (isNaN(seqNum)) seqNum = 1;

    var seqPadding = 0;
    if (dialogResult.zeroPadding) {
        seqPadding = seqText.length > 1 ? seqText.length : (seqNum + selection.length - 1).toString().length;
    }

    var fileNameNoExt = "";
    if (dialogResult.useFileName && doc && doc.name) {
        var lastDot = doc.name.lastIndexOf(".");
        fileNameNoExt = lastDot > 0 ? doc.name.substring(0, lastDot) : doc.name;
    }

    // 選択されたグループごとにアートボードを追加
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "GroupItem") {
            var bounds = dialogResult.usePreviewBounds ? selection[i].visibleBounds : selection[i].geometricBounds;
            var left = bounds[0] - margin;
            var top = bounds[1] + margin;
            var right = bounds[2] + margin;
            var bottom = bounds[3] - margin;
            try {
                doc.artboards.add([left, top, right, bottom]);
                var newName = buildArtboardName(
                    dialogResult.artboardName,
                    dialogResult.symbol,
                    seqNum,
                    dialogResult.zeroPadding,
                    dialogResult.useFileName,
                    fileNameNoExt,
                    seqPadding
                );
                doc.artboards[doc.artboards.length - 1].name = newName;
                seqNum++;
            } catch (e) {}
        }
    }

    // 既存アートボードを削除（新規追加後に実行）
    if (dialogResult.deleteArtboards) {
        for (var j = 0; j < initialCount; j++) {
            doc.artboards.remove(0);
        }
    }
}

main();