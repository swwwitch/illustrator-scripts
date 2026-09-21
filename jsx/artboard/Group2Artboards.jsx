#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したグループオブジェクトの境界に指定したマージンを加え、その範囲をアートボードとして追加します。
複数のグループ選択に対応し、名前の連番付与やファイル名参照、既存アートボードの削除オプションも備えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/Group2Artboards.md

### Overview

Adds an artboard covering the bounds of each selected group, plus a margin you specify.
Several groups can be processed at once, with options for sequential naming, using the file name, and removing the existing artboards.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/Group2Artboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "Group2Artboards";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/Group2Artboards.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/Group2Artboards.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */

    var LABELS = {
        dialog: {
            title: { ja: "アートボード化", en: "Artboard" }
        },
        panel: {
            artboard: { ja: "グループをアートボードに", en: "Convert Groups to Artboards" },
            name:     { ja: "アートボード名", en: "Artboard Name" }
        },
        fieldLabel: {
            margin:      { ja: "マージン", en: "Margin" },
            prefix:      { ja: "接頭辞", en: "Prefix" },
            symbol:      { ja: "記号", en: "Symbol" },
            startNumber: { ja: "開始番号", en: "Start Number" },
            example:     { ja: "例", en: "Example" }
        },
        checkbox: {
            previewBounds:   { ja: "プレビュー境界", en: "Preview bounds" },
            deleteArtboards: { ja: "既存のアートボードを削除", en: "Delete existing artboards" },
            useFileName:     { ja: "ファイル名を参照", en: "Use file name" },
            zeroPadding:     { ja: "ゼロ埋め", en: "Zero Padding" }
        },
        radio: {
            dash:       { ja: "-", en: "-" },
            underscore: { ja: "_", en: "_" },
            none:       { ja: "なし", en: "None" }
        },
        tooltip: {
            previewBounds: {
                ja: "線幅や効果を含めた見た目の端に合わせてアートボードを作ります。オフにするとパスの端が基準になります。",
                en: "Sizes each artboard to the visible edges including strokes and effects. Off uses the path edges."
            },
            margin:          { ja: "グループの外側に足す余白です。", en: "Extra space added around the group." },
            deleteArtboards: { ja: "作成する前に、いま開いているアートボードをすべて削除します。", en: "Removes every existing artboard before creating the new ones." },
            useFileName:     { ja: "接頭辞にドキュメントのファイル名（拡張子なし）を使います。", en: "Uses the document file name, without its extension, as the prefix." },
            prefix:          { ja: "アートボード名の先頭に付ける文字列です。", en: "Text placed at the start of each artboard name." },
            symbol:          { ja: "接頭辞と連番のあいだに入れる記号です。", en: "Character placed between the prefix and the number." },
            startNumber:     { ja: "連番の開始値です。", en: "The number the sequence starts from." },
            zeroPadding:     { ja: "開始番号の桁数にそろえて 0 を補います（01, 02, ...）。", en: "Pads the numbers with zeros to the width of the start number (01, 02, ...)." }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.name" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのドット区切りキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ": ");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /* アートボード名を生成する共通関数 / Common function to build artboard name */
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

    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
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
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;

            // 追加：プレビュー等の即時反映（onChanging があれば呼ぶ）
            try { if (typeof editText.onChanging === 'function') editText.onChanging(); } catch (e) {}
        });
    }

    function showDialog() {
        var dialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = [15, 20, 15, 10];
        dialog.spacing = 10;

        var rulerUnit = getUnitInfo("rulerType").label;

        var controlGroup = dialog.add("group");
        controlGroup.orientation = "column";
        controlGroup.alignChildren = "left";
        controlGroup.margins = [15, 5, 15, 10];
        controlGroup.spacing = 10;

        var previewBoundsCheck = controlGroup.add("checkbox", undefined, getLabel('checkbox.previewBounds'));
        previewBoundsCheck.helpTip = getLabel('tooltip.previewBounds');
        previewBoundsCheck.value = true;

        var marginGroup = controlGroup.add("group");
        marginGroup.orientation = "row";
        marginGroup.alignChildren = "center";
        marginGroup.add("statictext", undefined, labelText('fieldLabel.margin'));
        var marginInput = marginGroup.add("edittext", undefined, "0");
        marginInput.helpTip = getLabel('tooltip.margin');
        marginInput.characters = 5;
        marginGroup.add("statictext", undefined, rulerUnit);
        changeValueByArrowKey(marginInput);

        var deleteArtboardsCheck = controlGroup.add("checkbox", undefined, getLabel('checkbox.deleteArtboards'));
        deleteArtboardsCheck.helpTip = getLabel('tooltip.deleteArtboards');
        deleteArtboardsCheck.value = true;

        /* アートボード名パネル / Artboard name panel */
        var namePanel = dialog.add("panel");
        namePanel.text = getLabel('panel.name');
        namePanel.orientation = "row";
        namePanel.alignChildren = "center";
        namePanel.margins = [15, 25, 15, 10];

        var nameGroup = namePanel.add("group");
        nameGroup.orientation = "column";
        nameGroup.alignChildren = "left";

        var useFileNameCheck = nameGroup.add("checkbox", undefined, getLabel('checkbox.useFileName'));
        useFileNameCheck.helpTip = getLabel('tooltip.useFileName');
        useFileNameCheck.value = false;

        var prefixRow = nameGroup.add("group");
        prefixRow.orientation = "row";
        prefixRow.alignChildren = "center";
        prefixRow.add("statictext", undefined, labelText('fieldLabel.prefix'));
        var nameInput = prefixRow.add("edittext", undefined, "");
        nameInput.helpTip = getLabel('tooltip.prefix');
        nameInput.characters = 15;

        var symbolGroup = nameGroup.add("group");
        symbolGroup.orientation = "row";
        symbolGroup.alignChildren = "center";
        symbolGroup.add("statictext", undefined, labelText('fieldLabel.symbol'));
        var radioDash = symbolGroup.add("radiobutton", undefined, getLabel('radio.dash'));
        radioDash.helpTip = getLabel('tooltip.symbol');
        var radioUnderscore = symbolGroup.add("radiobutton", undefined, getLabel('radio.underscore'));
        radioUnderscore.helpTip = getLabel('tooltip.symbol');
        var radioNone = symbolGroup.add("radiobutton", undefined, getLabel('radio.none'));
        radioNone.helpTip = getLabel('tooltip.symbol');
        radioDash.value = true;

        var seqRow = nameGroup.add("group");
        seqRow.orientation = "row";
        seqRow.alignChildren = "center";
        seqRow.add("statictext", undefined, labelText('fieldLabel.startNumber'));
        var seqInput = seqRow.add("edittext", undefined, "01");
        seqInput.helpTip = getLabel('tooltip.startNumber');
        seqInput.characters = 5;
        changeValueByArrowKey(seqInput);
        var zeroPaddingCheck = seqRow.add("checkbox", undefined, getLabel('checkbox.zeroPadding'));
        zeroPaddingCheck.helpTip = getLabel('tooltip.zeroPadding');
        zeroPaddingCheck.value = true;

        var previewText = nameGroup.add("statictext", undefined, "");
        previewText.alignment = "left";
        previewText.characters = 20;

        /* プレビュー更新 / Update preview */
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
            previewText.text = getLabel('fieldLabel.example') + buildArtboardName(prefix, symbol, seq, zeroPadding, useFileNameCheck.value, fileNameNoExt, seq.length);
        }

        /* イベント登録 / Register events */
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
        var cancelBtn = buttonGroup.add("button", undefined, getLabel('button.cancel'));
        var okBtn = buttonGroup.add("button", undefined, getLabel('button.ok'), {
            name: "ok"
        });

        var dialogResult = null;
        cancelBtn.onClick = function() {
            dialog.close();
        };
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

        /* 選択されたグループごとにアートボードを追加 / Add artboards for each selected group */
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "GroupItem") {
                var bounds;
                if (dialogResult.usePreviewBounds && selection[i].clipped) {
                    /* クリップグループの場合はマスクパスのジオメトリを使う / For clipped groups, use mask path geometry */
                    var maskItem = selection[i].pageItems[0];
                    bounds = maskItem.geometricBounds;
                } else {
                    bounds = dialogResult.usePreviewBounds ? selection[i].visibleBounds : selection[i].geometricBounds;
                }

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

        /* 既存アートボードを削除（新規追加後に実行）/ Delete existing artboards (after adding new ones) */
        if (dialogResult.deleteArtboards) {
            for (var j = 0; j < initialCount; j++) {
                doc.artboards.remove(0);
            }
        }
    }

    main();

})();
