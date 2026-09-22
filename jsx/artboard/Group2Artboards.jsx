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
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/Group2Artboards.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/Group2Artboards.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値 / Dialog defaults */
    var DIALOG_DEFAULTS = {
        previewBounds: true,        /* プレビュー境界 / preview bounds */
        margin: "0",                /* マージン（定規単位の表示） / margin */
        deleteArtboards: true,      /* 既存のアートボードを削除 / delete existing artboards */
        useFileName: false,         /* ファイル名を参照 / use the file name */
        prefix: "",                 /* 接頭辞 / prefix */
        startNumber: "01",          /* 開始番号 / start number */
        zeroPadding: true           /* ゼロ埋め / zero padding */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = [15, 20, 15, 10];          /* ダイアログの余白 / dialog margins */
    var DIALOG_SPACING = 10;                        /* ダイアログ内の間隔 / dialog spacing */
    var OPTION_GROUP_MARGINS = [15, 5, 15, 10];     /* 上部オプションの余白 / top option group margins */
    var OPTION_GROUP_SPACING = 10;                  /* 上部オプションの間隔 / top option group spacing */
    var NAME_PANEL_MARGINS = [15, 25, 15, 10];      /* 「アートボード名」パネルの余白 / name panel margins */
    var BUTTON_ROW_MARGINS = [0, 10, 0, 10];        /* ボタン行の余白 / button row margins */
    var NUMBER_FIELD_CHARACTERS = 5;                /* マージン・開始番号欄の桁数 / margin & start number field width */
    var PREFIX_FIELD_CHARACTERS = 15;               /* 接頭辞欄の桁数 / prefix field width */
    var NAME_EXAMPLE_CHARACTERS = 20;               /* 名前の例の表示幅 / name example width */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
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
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのドット区切りキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
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

    // =========================================
    // アートボード名 / Artboard names
    // =========================================

    /**
     * アートボード名を生成する
     * @param {string} prefix - 接頭辞
     * @param {string} symbol - 区切り記号（"-" / "_" / ""）
     * @param {number|string} sequenceValue - 連番
     * @param {boolean} zeroPadding - ゼロ埋めするか
     * @param {boolean} useFileName - ファイル名を先頭に付けるか
     * @param {string} fileNameNoExt - 拡張子を除いたファイル名
     * @param {number} padLength - ゼロ埋めの桁数
     * @returns {string} アートボード名
     */
    function buildArtboardName(prefix, symbol, sequenceValue, zeroPadding, useFileName, fileNameNoExt, padLength) {
        var sequenceNumber = parseInt(sequenceValue, 10);
        if (isNaN(sequenceNumber)) sequenceNumber = 1;
        var numberText = sequenceNumber.toString();
        if (zeroPadding && padLength > 0) {
            while (numberText.length < padLength) numberText = "0" + numberText;
        }
        if (useFileName) {
            if (prefix === "") {
                return fileNameNoExt + symbol + numberText;
            }
            return fileNameNoExt + symbol + prefix + symbol + numberText;
        }
        return prefix + symbol + numberText;
    }

    /**
     * ドキュメントのファイル名から拡張子を除く
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {string} 拡張子を除いたファイル名（名前が無ければ空文字）
     */
    function getFileNameWithoutExtension(targetDocument) {
        if (!targetDocument || !targetDocument.name) return "";
        var documentName = targetDocument.name;
        var lastDot = documentName.lastIndexOf(".");
        return lastDot > 0 ? documentName.substring(0, lastDot) : documentName;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ↑↓キーで数値を増減し、入力欄の onChanging を呼ぶ
     * ↑↓で±1、Shift+↑↓で10の倍数にスナップ、Option+↑↓で±0.1（↑↓以外のキーでも値を丸め直す）。
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            if (isUp || event.keyName == "Down") {
                if (keyboard.shiftKey) {
                    /* Shiftキー押下時は10の倍数にスナップ / snap to multiples of 10 with Shift */
                    value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.max(0, Math.floor((value - 1) / 10) * 10);
                } else if (keyboard.altKey) {
                    /* Optionキー押下時は0.1単位で増減 / change by 0.1 with Option */
                    value += isUp ? 0.1 : -0.1;
                } else {
                    value = isUp ? value + 1 : Math.max(0, value - 1);
                }
                event.preventDefault();
            }

            /* Option 時は小数第1位、それ以外は整数に丸める / round to 0.1 with Option, otherwise to an integer */
            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
            editText.text = value;

            /* プレビュー等の即時反映（onChanging があれば呼ぶ） / refresh the preview through onChanging */
            if (typeof editText.onChanging === 'function') editText.onChanging();
        });
    }

    /**
     * 「項目名：」付きの横並びの行を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} labelPath - 項目名のラベルのパス
     * @returns {Group} 行のグループ
     */
    function addLabeledRow(parentGroup, labelPath) {
        var labeledRow = parentGroup.add("group");
        labeledRow.orientation = "row";
        labeledRow.alignChildren = "center";
        labeledRow.add("statictext", undefined, labelText(labelPath));
        return labeledRow;
    }

    /**
     * 「アートボード名」パネルを作る
     * @param {Window} group2ArtboardsDialog - ダイアログ
     * @returns {Object} パネル内のコントロール
     */
    function buildNamePanel(group2ArtboardsDialog) {
        var namePanel = group2ArtboardsDialog.add("panel");
        namePanel.text = getLabel('panel.name');
        namePanel.orientation = "row";
        namePanel.alignChildren = "center";
        namePanel.margins = NAME_PANEL_MARGINS;

        var nameColumn = namePanel.add("group");
        nameColumn.orientation = "column";
        nameColumn.alignChildren = "left";

        var nameControls = {};
        nameControls.useFileNameCheck = nameColumn.add("checkbox", undefined, getLabel('checkbox.useFileName'));
        nameControls.useFileNameCheck.helpTip = getLabel('tooltip.useFileName');
        nameControls.useFileNameCheck.value = DIALOG_DEFAULTS.useFileName;

        var prefixRow = addLabeledRow(nameColumn, 'fieldLabel.prefix');
        nameControls.prefixInput = prefixRow.add("edittext", undefined, DIALOG_DEFAULTS.prefix);
        nameControls.prefixInput.helpTip = getLabel('tooltip.prefix');
        nameControls.prefixInput.characters = PREFIX_FIELD_CHARACTERS;

        var symbolRow = addLabeledRow(nameColumn, 'fieldLabel.symbol');
        nameControls.dashRadio = symbolRow.add("radiobutton", undefined, getLabel('radio.dash'));
        nameControls.dashRadio.helpTip = getLabel('tooltip.symbol');
        nameControls.underscoreRadio = symbolRow.add("radiobutton", undefined, getLabel('radio.underscore'));
        nameControls.underscoreRadio.helpTip = getLabel('tooltip.symbol');
        nameControls.noSymbolRadio = symbolRow.add("radiobutton", undefined, getLabel('radio.none'));
        nameControls.noSymbolRadio.helpTip = getLabel('tooltip.symbol');
        nameControls.dashRadio.value = true;

        var startNumberRow = addLabeledRow(nameColumn, 'fieldLabel.startNumber');
        nameControls.startNumberInput = startNumberRow.add("edittext", undefined, DIALOG_DEFAULTS.startNumber);
        nameControls.startNumberInput.helpTip = getLabel('tooltip.startNumber');
        nameControls.startNumberInput.characters = NUMBER_FIELD_CHARACTERS;
        changeValueByArrowKey(nameControls.startNumberInput);
        nameControls.zeroPaddingCheck = startNumberRow.add("checkbox", undefined, getLabel('checkbox.zeroPadding'));
        nameControls.zeroPaddingCheck.helpTip = getLabel('tooltip.zeroPadding');
        nameControls.zeroPaddingCheck.value = DIALOG_DEFAULTS.zeroPadding;

        nameControls.nameExampleText = nameColumn.add("statictext", undefined, "");
        nameControls.nameExampleText.alignment = "left";
        nameControls.nameExampleText.characters = NAME_EXAMPLE_CHARACTERS;
        return nameControls;
    }

    /**
     * 選んでいる区切り記号を返す
     * @param {Object} nameControls - 「アートボード名」パネルのコントロール
     * @returns {string} "-" / "_" / ""
     */
    function getSelectedSymbol(nameControls) {
        if (nameControls.dashRadio.value) return "-";
        return nameControls.underscoreRadio.value ? "_" : "";
    }

    /**
     * 設定ダイアログを表示する
     * @returns {Object|null} 設定。キャンセル時は null
     */
    function showDialog() {
        var group2ArtboardsDialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        group2ArtboardsDialog.orientation = "column";
        group2ArtboardsDialog.alignChildren = "fill";
        group2ArtboardsDialog.margins = DIALOG_MARGINS;
        group2ArtboardsDialog.spacing = DIALOG_SPACING;

        var rulerUnit = getUnitInfo("rulerType").label;

        var optionGroup = group2ArtboardsDialog.add("group");
        optionGroup.orientation = "column";
        optionGroup.alignChildren = "left";
        optionGroup.margins = OPTION_GROUP_MARGINS;
        optionGroup.spacing = OPTION_GROUP_SPACING;

        var previewBoundsCheck = optionGroup.add("checkbox", undefined, getLabel('checkbox.previewBounds'));
        previewBoundsCheck.helpTip = getLabel('tooltip.previewBounds');
        previewBoundsCheck.value = DIALOG_DEFAULTS.previewBounds;

        var marginRow = addLabeledRow(optionGroup, 'fieldLabel.margin');
        var marginInput = marginRow.add("edittext", undefined, DIALOG_DEFAULTS.margin);
        marginInput.helpTip = getLabel('tooltip.margin');
        marginInput.characters = NUMBER_FIELD_CHARACTERS;
        marginRow.add("statictext", undefined, rulerUnit);
        changeValueByArrowKey(marginInput);

        var deleteArtboardsCheck = optionGroup.add("checkbox", undefined, getLabel('checkbox.deleteArtboards'));
        deleteArtboardsCheck.helpTip = getLabel('tooltip.deleteArtboards');
        deleteArtboardsCheck.value = DIALOG_DEFAULTS.deleteArtboards;

        /* アートボード名パネル / Artboard name panel */
        var nameControls = buildNamePanel(group2ArtboardsDialog);

        /* 名前の例を更新 / Update the name example */
        function updateNameExample() {
            var startText = nameControls.startNumberInput.text;
            var useFileName = nameControls.useFileNameCheck.value;
            var fileNameNoExt = useFileName ? getFileNameWithoutExtension(app.activeDocument) : "";
            var exampleName = buildArtboardName(nameControls.prefixInput.text, getSelectedSymbol(nameControls), startText,
                nameControls.zeroPaddingCheck.value, useFileName, fileNameNoExt, startText.length);
            /* 英語はコロンのあとに空白を入れる / add a space after the colon in English */
            nameControls.nameExampleText.text = labelText('fieldLabel.example') + (uiLang === "ja" ? "" : " ") + exampleName;
        }

        /* イベント登録 / Register events */
        nameControls.prefixInput.onChanging = updateNameExample;
        nameControls.dashRadio.onClick = updateNameExample;
        nameControls.underscoreRadio.onClick = updateNameExample;
        nameControls.noSymbolRadio.onClick = updateNameExample;
        nameControls.startNumberInput.onChanging = updateNameExample;
        nameControls.zeroPaddingCheck.onClick = updateNameExample;
        nameControls.useFileNameCheck.onClick = updateNameExample;
        updateNameExample();

        var btnRowGroup = group2ArtboardsDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        btnRowGroup.margins = BUTTON_ROW_MARGINS;
        var btnCancel = btnRowGroup.add("button", undefined, getLabel('button.cancel'));
        var btnOK = btnRowGroup.add("button", undefined, getLabel('button.ok'), {
            name: "ok"
        });

        var dialogResult = null;
        btnCancel.onClick = function () {
            group2ArtboardsDialog.close();
        };
        btnOK.onClick = function () {
            dialogResult = {
                marginValue: marginInput.text,
                deleteArtboards: deleteArtboardsCheck.value,
                usePreviewBounds: previewBoundsCheck.value,
                artboardName: nameControls.prefixInput.text,
                sequentialText: nameControls.startNumberInput.text,
                zeroPadding: nameControls.zeroPaddingCheck.value,
                symbol: getSelectedSymbol(nameControls),
                useFileName: nameControls.useFileNameCheck.value
            };
            group2ArtboardsDialog.close(1);
        };
        if (group2ArtboardsDialog.show() !== 1) return null;
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * グループからアートボードにする範囲を求める
     * プレビュー境界でクリップグループのときはマスクパスの幾何境界を使う。
     * @param {GroupItem} groupItem - 対象のグループ
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @param {number} margin - マージン
     * @returns {number[]} [left, top, right, bottom]
     */
    function getGroupArtboardRect(groupItem, usePreviewBounds, margin) {
        var bounds;
        if (usePreviewBounds && groupItem.clipped) {
            /* クリップグループの場合はマスクパスのジオメトリを使う / For clipped groups, use mask path geometry */
            bounds = groupItem.pageItems[0].geometricBounds;
        } else {
            bounds = usePreviewBounds ? groupItem.visibleBounds : groupItem.geometricBounds;
        }
        return [bounds[0] - margin, bounds[1] + margin, bounds[2] + margin, bounds[3] - margin];
    }

    /**
     * 選択したグループごとにアートボードを追加する
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0) return;

        var dialogResult = showDialog();
        if (!dialogResult) return;

        var margin = parseFloat(dialogResult.marginValue);
        if (isNaN(margin)) margin = 0;
        var initialCount = doc.artboards.length;

        var startText = dialogResult.sequentialText;
        var sequenceNumber = parseInt(startText, 10);
        if (isNaN(sequenceNumber)) sequenceNumber = 1;

        var padLength = 0;
        if (dialogResult.zeroPadding) {
            padLength = startText.length > 1 ? startText.length : (sequenceNumber + selectedItems.length - 1).toString().length;
        }

        var fileNameNoExt = dialogResult.useFileName ? getFileNameWithoutExtension(doc) : "";

        /* 選択されたグループごとにアートボードを追加 / Add artboards for each selected group */
        for (var i = 0; i < selectedItems.length; i++) {
            var groupItem = selectedItems[i];
            if (groupItem.typename !== "GroupItem") continue;

            var artboardRect = getGroupArtboardRect(groupItem, dialogResult.usePreviewBounds, margin);
            /* 追加できない範囲（キャンバス外など）はスキップ / skip rects Illustrator refuses (off-canvas, etc.) */
            try {
                doc.artboards.add(artboardRect);
                doc.artboards[doc.artboards.length - 1].name = buildArtboardName(
                    dialogResult.artboardName,
                    dialogResult.symbol,
                    sequenceNumber,
                    dialogResult.zeroPadding,
                    dialogResult.useFileName,
                    fileNameNoExt,
                    padLength
                );
                sequenceNumber++;
            } catch (e) {}
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
