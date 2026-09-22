#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Excel由来のIllustratorデータを、表組みとして扱いやすい状態に整形します。
クリッピングマスクの解除、テキストの整理と配置、セル背景の抽出、罫線の中心線化と均等配置をまとめて行います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExcelTableNormalizer.md

### Overview

Normalizes Illustrator artwork that came from Excel so that it is easier to work with as a table.
It releases clipping masks, tidies and aligns the text, extracts cell backgrounds, and converts the rules to evenly distributed center lines.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExcelTableNormalizer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExcelTableNormalizer";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExcelTableNormalizer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExcelTableNormalizer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var TEXT_LAYER_NAME = "_text_all";          /* テキストをまとめるレイヤー / Layer that collects the text */
    var PATH_LAYER_NAME = "_cell_rectangle";    /* セル背景をまとめるレイヤー / Layer that collects the cell backgrounds */

    // =========================================
    // 処理の中断 / Cancellation
    // =========================================

    /* 列ごとの配置ダイアログをキャンセルしたときに投げる目印 / Marker thrown when the column alignment dialog is cancelled */
    var USER_CANCELLED_COLUMN_ALIGNMENT = "USER_CANCELLED_COLUMN_ALIGNMENT";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 20;                    /* ダイアログの余白 / Dialog margins */
    var DIALOG_SPACING = 10;                    /* ダイアログ内の要素間隔 / Spacing inside the dialog */
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */
    var BUTTON_ROW_MARGINS = [0, 10, 0, 0];     /* ボタン行の余白 / Button row margins */
    var SAMPLE_LABEL_CHARS = 17;                /* 列ごとの配置：サンプル欄の幅（文字数） / Column alignment: sample width in characters */

    /**
     * パネルの共通レイアウトを設定する
     * @param {Panel} targetPanel - 設定するパネル
     * @param {number} [spacing] - 子要素の間隔
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Define Japanese-English labels */
    var LABELS = {
        dialog: {
            title: { ja: "Excelデータを整形", en: "Format Excel Data" },
            columnAlignmentTitle: { ja: "列ごとの配置", en: "Column Alignment" }
        },
        panel: {
            options: { ja: "オプション", en: "Options" },
            text: { ja: "テキスト", en: "Text" },
            cellBackground: { ja: "セル背景", en: "Cell Background" },
            rules: { ja: "罫線", en: "Rules" },
            placementMode: { ja: "罫線の配置", en: "Rule placement" },
            style: { ja: "スタイル", en: "Style" }
        },
        checkbox: {
            releaseMask: { ja: "クリッピングマスクを解除", en: "Release clipping masks" },
            removeSmallObjects: { ja: "小さいマークなどを削除", en: "Remove small marks" },
            moveTextToLayer: { ja: "専用レイヤーへ移動", en: "Move text to dedicated layer" },
            setTextK100: { ja: "印刷用の「黒」に", en: "Set text to print black" },
            removeDuplicateTexts: { ja: "重複テキストを削除", en: "Remove duplicate texts" },
            adjustCellBackground: { ja: "セル背景を抽出・統合", en: "Extract and merge cell backgrounds" },
            equalizeHeights: { ja: "罫線にスナップ", en: "Snap to rules" },
            centerline: { ja: "罫線を中心線化", en: "Convert rules to centerlines" },
            outerToRect: { ja: "外枠を長方形に変換", en: "Convert outer frame to rectangle" },
            rulesK100: { ja: "罫線を印刷用の「黒」に", en: "Set rules to print black" }
        },
        radio: {
            placementUniformForced: { ja: "行・列を均等に配置", en: "Distribute rows and columns evenly" },
            placementUniformMerged: { ja: "結合セルを考慮して均等配置", en: "Distribute evenly with merged cells" },
            placementKeepWidths: { ja: "列幅を保持", en: "Preserve column widths" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Stroke width" }
        },
        /* 列ごとの配置ダイアログの各行 / Rows of the column alignment dialog */
        columnRow: {
            columnNumber: { ja: "列目", en: "Column" },
            noSample: { ja: "サンプルなし", en: "No sample" }
        },
        tooltip: {
            releaseMask: {
                ja: "貼り付けた表に掛かっているクリッピングマスクを外します。",
                en: "Releases the clipping mask that comes with the pasted table."
            },
            removeSmallObjects: {
                ja: "ごく小さなゴミオブジェクトを削除します。",
                en: "Deletes the tiny stray objects that come with the paste."
            },
            moveTextToLayer: { ja: "テキストを専用のレイヤーへまとめて移します。", en: "Moves the text onto a layer of its own." },
            setTextK100: { ja: "テキストの色をスミ100%（K100）にします。", en: "Sets the text color to 100% black (K100)." },
            removeDuplicateTexts: {
                ja: "同じ位置に重なっている同じ文字列を1つにまとめます。",
                en: "Merges duplicated text that sits on top of itself."
            },
            adjustCellBackground: {
                ja: "セルの背景を、ケイ線に合わせて引き直します。",
                en: "Redraws the cell backgrounds to line up with the rules."
            },
            equalizeHeights: { ja: "同じ行のセルの高さをそろえます。", en: "Gives the cells in a row the same height." },
            centerline: {
                ja: "太い長方形のケイ線を、中心を通る1本の線に置き換えます。",
                en: "Replaces thick rectangular rules with a single centre line."
            },
            outerToRect: { ja: "表の外周のケイ線を1つの長方形にまとめます。", en: "Merges the outer rules into a single rectangle." },
            placementUniformForced: { ja: "セル幅を強制的に均等にそろえます。", en: "Forces every cell to the same width." },
            placementUniformMerged: {
                ja: "結合セルを考慮しながら、セル幅を均等にそろえます。",
                en: "Evens out the cell widths while respecting merged cells."
            },
            placementKeepWidths: { ja: "元の幅をそのまま保ちます。", en: "Keeps the original widths as they are." },
            strokeWidth: { ja: "ケイ線の太さです。", en: "Weight of the rules." },
            rulesK100: { ja: "ケイ線の色をスミ100%（K100）にします。", en: "Sets the rule color to 100% black (K100)." },
            columnAlignment: { ja: "この列のテキストの行揃えです。", en: "How the text in this column is aligned." },
            autoAlignment: {
                ja: "すべての列を最初の判定に戻します（数字が多い列は右、それ以外は左）。",
                en: "Resets every column to the detected default: right for mostly numeric columns, left otherwise."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            autoAlignment: { ja: "自動判定", en: "Auto" }
        },
        alert: {
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 共通ユーティリティ / Common utilities
    // =========================================

    /**
     * 関数を呼び、例外が出たら握りつぶして null を返す（失敗しても続行してよい DOM 操作用）
     * @param {Function} callback - 呼ぶ関数
     * @returns {*} 関数の戻り値。例外時は null
     */
    function safeCall(callback) {
        try {
            return callback();
        } catch (e) { }
        return null;
    }

    /**
     * 配列に同じ値（参照）が含まれているかどうか
     * @param {Array} itemList - 調べる配列
     * @param {*} targetItem - 探す値
     * @returns {boolean} 含まれていれば true
     */
    function containsItem(itemList, targetItem) {
        for (var i = 0; i < itemList.length; i++) {
            if (itemList[i] === targetItem) return true;
        }
        return false;
    }

    /**
     * アイテムを削除する（削除できないものは飛ばす）
     * @param {PageItem[]} pageItems - 削除するアイテム
     * @param {number} startIndex - この位置から後ろを削除する
     * @returns {void}
     */
    function removeItemsQuietly(pageItems, startIndex) {
        for (var i = startIndex; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            safeCall(function () { pageItem.remove(); });
        }
    }

    /**
     * 対象にするレイヤー（除外名に当たらず、ロックされておらず、表示中のもの）を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {string[]} excludeLayerNames - 除外するレイヤー名
     * @returns {Layer[]} 対象のレイヤー
     */
    function getTargetLayers(doc, excludeLayerNames) {
        var targetLayers = [];
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var docLayer = doc.layers[layerIndex];
            if (containsItem(excludeLayerNames, docLayer.name) || docLayer.locked || !docLayer.visible) continue;
            targetLayers.push(docLayer);
        }
        return targetLayers;
    }

    /**
     * 名前でレイヤーを探し、無ければ作る
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 見つけた（または作った）レイヤー
     */
    function findOrCreateLayer(doc, layerName) {
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            if (doc.layers[layerIndex].name === layerName) return doc.layers[layerIndex];
        }
        var newLayer = doc.layers.add();
        newLayer.name = layerName;
        return newLayer;
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

    /**
     * 線幅欄の初期値を返す
     * @param {string} unitLabel - 線の単位ラベル
     * @returns {string} pt なら "0.25"、それ以外は "0.1"
     */
    function getDefaultStrokeWidthText(unitLabel) {
        return (unitLabel === "pt") ? "0.25" : "0.1";
    }

    /**
     * 単位ラベルの値を pt に換算する（mm / cm / in 以外はそのままの数値を返す）
     * @param {number} value - 換算する値
     * @param {string} unitLabel - 単位ラベル
     * @returns {number} pt の値
     */
    function convertLengthToPoints(value, unitLabel) {
        switch (unitLabel) {
            case "mm": return value * 72 / 25.4;
            case "cm": return value * 72 / 2.54;
            case "in": return value * 72;
            case "pt": return value;
            case "px": return value; /* Illustrator では px ≒ pt / px equals pt in Illustrator */
            default: return value;
        }
    }

    /**
     * ↑↓キーで数値を増減する（Shift は 10 刻み、Option は 0.1 刻み）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is pressed
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減 / Change by 0.1 when Option is pressed
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め / Round to 1 decimal
                value = Math.round(value * 10) / 10;
            } else if (keyboard.shiftKey) {
                // Shiftキー押下時のみ整数に丸め / Round to integer only when Shift is pressed
                value = Math.round(value);
            }

            editText.text = String(value);
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、OK で整形処理を実行する
     * @returns {void}
     */
    function main() {
        safeCall(function () { app.executeMenuCommand('Colors8'); });
        var strokeUnitLabel = getUnitInfo("strokeUnits").label;
        var dialogControls = buildMainDialog(strokeUnitLabel);

        dialogControls.btnOK.onClick = function () {
            var normalizeOptions = readMainDialogOptions(dialogControls, strokeUnitLabel);
            dialogControls.mainDialog.close();
            runWithUndoRollback(function () {
                executeRelease(normalizeOptions);
            });
        };
        dialogControls.mainDialog.show();
    }

    /**
     * tooltip 付きのチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * tooltip 付きのラジオボタンを追加する
     * @param {Object} parentContainer - 追加先のパネルまたはグループ
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addOptionRadio(parentContainer, labelPath, tooltipPath) {
        var optionRadio = parentContainer.add("radiobutton", undefined, getLabel(labelPath));
        optionRadio.helpTip = getLabel(tooltipPath);
        return optionRadio;
    }

    /**
     * 整形オプションのダイアログを組み立てる
     * @param {string} strokeUnitLabel - 線の単位ラベル
     * @returns {Object} ダイアログ（mainDialog）と各コントロール
     */
    function buildMainDialog(strokeUnitLabel) {
        var mainDialog = new Window('dialog', getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        mainDialog.orientation = "column";
        mainDialog.alignChildren = "fill";
        mainDialog.spacing = DIALOG_SPACING;
        mainDialog.margins = DIALOG_MARGINS;

        var columnsGroup = mainDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";
        columnsGroup.spacing = DIALOG_SPACING;

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";
        leftColumn.spacing = DIALOG_SPACING;

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";
        rightColumn.spacing = DIALOG_SPACING;

        var optionsPanel = leftColumn.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel);
        var releaseMaskCheckbox = addOptionCheckbox(optionsPanel, "checkbox.releaseMask", "tooltip.releaseMask", true);
        var removeSmallCheckbox = addOptionCheckbox(optionsPanel, "checkbox.removeSmallObjects", "tooltip.removeSmallObjects", true);

        var textPanel = leftColumn.add("panel", undefined, getLabel("panel.text"));
        setupPanel(textPanel);
        var moveTextCheckbox = addOptionCheckbox(textPanel, "checkbox.moveTextToLayer", "tooltip.moveTextToLayer", true);
        var setK100Checkbox = addOptionCheckbox(textPanel, "checkbox.setTextK100", "tooltip.setTextK100", true);
        var removeDuplicateCheckbox = addOptionCheckbox(textPanel, "checkbox.removeDuplicateTexts", "tooltip.removeDuplicateTexts", true);

        var cellBgPanel = leftColumn.add("panel", undefined, getLabel("panel.cellBackground"));
        setupPanel(cellBgPanel);
        var adjustCellBgCheckbox = addOptionCheckbox(cellBgPanel, "checkbox.adjustCellBackground", "tooltip.adjustCellBackground", true);
        var equalizeHeightsCheckbox = addOptionCheckbox(cellBgPanel, "checkbox.equalizeHeights", "tooltip.equalizeHeights", true);
        equalizeHeightsCheckbox.enabled = adjustCellBgCheckbox.value;

        adjustCellBgCheckbox.onClick = function () {
            equalizeHeightsCheckbox.enabled = adjustCellBgCheckbox.value;
        };

        var rulesPanel = rightColumn.add("panel", undefined, getLabel("panel.rules"));
        setupPanel(rulesPanel);
        var centerlineCheckbox = addOptionCheckbox(rulesPanel, "checkbox.centerline", "tooltip.centerline", true);
        var outerToRectCheckbox = addOptionCheckbox(rulesPanel, "checkbox.outerToRect", "tooltip.outerToRect", true);

        var placementPanel = rulesPanel.add("panel", undefined, getLabel("panel.placementMode"));
        setupPanel(placementPanel, 6);
        var placementUniformForcedRadio = addOptionRadio(placementPanel, "radio.placementUniformForced", "tooltip.placementUniformForced");
        var placementUniformMergedRadio = addOptionRadio(placementPanel, "radio.placementUniformMerged", "tooltip.placementUniformMerged");
        var placementKeepWidthsRadio = addOptionRadio(placementPanel, "radio.placementKeepWidths", "tooltip.placementKeepWidths");
        placementKeepWidthsRadio.value = true;

        var stylePanel = rulesPanel.add("panel", undefined, getLabel("panel.style"));
        setupPanel(stylePanel);
        stylePanel.alignChildren = "fill";

        var strokeWidthGroup = stylePanel.add("group");
        strokeWidthGroup.orientation = "row";
        strokeWidthGroup.add("statictext", undefined, labelText("fieldLabel.strokeWidth"));
        var strokeWidthInput = strokeWidthGroup.add("edittext", undefined, getDefaultStrokeWidthText(strokeUnitLabel));
        strokeWidthInput.helpTip = getLabel("tooltip.strokeWidth");
        strokeWidthInput.characters = 5;
        changeValueByArrowKey(strokeWidthInput);
        strokeWidthGroup.add("statictext", undefined, strokeUnitLabel);

        var rulesK100Checkbox = addOptionCheckbox(stylePanel, "checkbox.rulesK100", "tooltip.rulesK100", true);

        var btnRowGroup = mainDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        btnRowGroup.margins = BUTTON_ROW_MARGINS;

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            mainDialog: mainDialog,
            releaseMaskCheckbox: releaseMaskCheckbox,
            removeSmallCheckbox: removeSmallCheckbox,
            moveTextCheckbox: moveTextCheckbox,
            setK100Checkbox: setK100Checkbox,
            removeDuplicateCheckbox: removeDuplicateCheckbox,
            adjustCellBgCheckbox: adjustCellBgCheckbox,
            equalizeHeightsCheckbox: equalizeHeightsCheckbox,
            centerlineCheckbox: centerlineCheckbox,
            outerToRectCheckbox: outerToRectCheckbox,
            placementUniformForcedRadio: placementUniformForcedRadio,
            placementUniformMergedRadio: placementUniformMergedRadio,
            strokeWidthInput: strokeWidthInput,
            rulesK100Checkbox: rulesK100Checkbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * ダイアログの値を読み取る（線幅は pt に換算。不正な値は初期値に戻す）
     * @param {Object} dialogControls - buildMainDialog() の戻り値
     * @param {string} strokeUnitLabel - 線の単位ラベル
     * @returns {Object} 整形オプション
     */
    function readMainDialogOptions(dialogControls, strokeUnitLabel) {
        var strokeWidthText = String(dialogControls.strokeWidthInput.text).replace(/,/g, ".");
        var strokeWidthValue = parseFloat(strokeWidthText);
        if (isNaN(strokeWidthValue) || strokeWidthValue <= 0) {
            strokeWidthValue = parseFloat(getDefaultStrokeWidthText(strokeUnitLabel));
        }
        return {
            releaseMask: dialogControls.releaseMaskCheckbox.value,
            moveText: dialogControls.moveTextCheckbox.value,
            setK100: dialogControls.setK100Checkbox.value,
            removeDuplicate: dialogControls.removeDuplicateCheckbox.value,
            removeSmall: dialogControls.removeSmallCheckbox.value,
            adjustCellBg: dialogControls.adjustCellBgCheckbox.value,
            equalizeHeights: dialogControls.equalizeHeightsCheckbox.value,
            centerline: dialogControls.centerlineCheckbox.value,
            placementMode: dialogControls.placementUniformForcedRadio.value ? "forced"
                : dialogControls.placementUniformMergedRadio.value ? "merged"
                    : "keepColumnWidths",
            outerToRect: dialogControls.outerToRectCheckbox.value,
            strokeWidthPt: convertLengthToPoints(strokeWidthValue, strokeUnitLabel),
            rulesK100: dialogControls.rulesK100Checkbox.value
        };
    }

    /**
     * 処理を実行し、列ごとの配置ダイアログでキャンセルされたら Undo で戻す
     * @param {Function} callback - 実行する処理
     * @returns {void}
     */
    function runWithUndoRollback(callback) {
        try {
            callback();
        } catch (e) {
            if (e === USER_CANCELLED_COLUMN_ALIGNMENT || (e && e.message === USER_CANCELLED_COLUMN_ALIGNMENT)) {
                safeCall(function () { app.undo(); });
                return;
            }
            throw e;
        }
    }

    /**
     * 選択オブジェクトに対して整形処理を順に実行する
     * @param {Object} normalizeOptions - readMainDialogOptions() の戻り値
     * @returns {void}
     */
    function executeRelease(normalizeOptions) {
        if (!app.documents.length || !app.activeDocument.selection.length) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = doc.selection;
        var ruleLayerExclusions = [TEXT_LAYER_NAME, PATH_LAYER_NAME];

        // 元の選択範囲の外形バウンズを保存（以後の処理で参照範囲をこれに限定する）
        // Snapshot original selection bounds so subsequent steps stay within the user's selection
        var initialSelectionBounds = computeUnionGeometricBounds(selectionItems);

        if (normalizeOptions.removeDuplicate) processClipGroupTexts(selectionItems);
        if (normalizeOptions.releaseMask) releaseClipGroups(selectionItems);
        if (normalizeOptions.removeSmall) removeSmallPathItems(doc);
        if (normalizeOptions.moveText) moveTextsToLayer(doc, TEXT_LAYER_NAME);
        if (normalizeOptions.setK100) applyK100ToAllTexts(doc);
        if (normalizeOptions.adjustCellBg) autoSelectAndMerge(PATH_LAYER_NAME, [TEXT_LAYER_NAME]);

        // セル背景のグループを強制解除（スナップ精度向上のため）/ Ungroup cell backgrounds for more accurate snapping
        ungroupAllInLayer(doc, PATH_LAYER_NAME);

        var finalItems = null;
        var centerlineApplied = false;
        if (normalizeOptions.centerline) {
            var centerLines = convertRectanglesToCenterLines(doc, ruleLayerExclusions);
            if (centerLines.length > 0) {
                centerLines = dedupeOverlappingLines(centerLines);
                applyPlacementMode(centerLines, normalizeOptions.placementMode, initialSelectionBounds);
                finalItems = dedupeOverlappingLines(centerLines);
                centerlineApplied = true;
            }
        }
        if (!finalItems) {
            finalItems = collectRuleLines(doc, ruleLayerExclusions);
        }
        if (finalItems.length > 0 && alignTextFramesByColumnDialog(doc, finalItems) === false) {
            throw new Error(USER_CANCELLED_COLUMN_ALIGNMENT);
        }

        // 外枠矩形化より前に、純粋な中心線群に対してセル背景をスナップする
        // Snap cell backgrounds while finalItems is still a homogeneous set of centerlines
        if (normalizeOptions.adjustCellBg && normalizeOptions.equalizeHeights && finalItems.length > 0) {
            snapCellBackgroundsAcrossLayers(doc, [TEXT_LAYER_NAME], finalItems, initialSelectionBounds);
        }

        // 外枠矩形化より前に線幅・色を確定させる
        // 新しい外枠矩形は referencePath（中心線）から strokeWidth/color を継承するため
        // Apply stroke styling before outerToRect so the new rectangle inherits the styled stroke
        applyUniformStrokeWidth(finalItems, normalizeOptions.strokeWidthPt);
        if (normalizeOptions.rulesK100) {
            applyK100ToLines(finalItems);
        }

        if (centerlineApplied && normalizeOptions.outerToRect && finalItems.length > 0) {
            convertOuterFrameToRectangle(finalItems);
        }
    }

    // =========================================
    // クリップグループとテキストの整理 / Clip groups and text cleanup
    // =========================================

    /**
     * 選択中のクリップグループからマスクパスを削除し、グループを解除する
     * @param {PageItem[]} selectionItems - 選択中のアイテム
     * @returns {void}
     */
    function releaseClipGroups(selectionItems) {
        for (var i = 0; i < selectionItems.length; i++) {
            var selectedItem = selectionItems[i];
            if (selectedItem.typename !== "GroupItem" || selectedItem.clipped !== true) continue;
            for (var j = 0; j < selectedItem.pageItems.length; j++) {
                var childItem = selectedItem.pageItems[j];
                if (childItem.typename === "PathItem" && childItem.clipping) {
                    childItem.remove();
                    break;
                }
            }
            ungroup(selectedItem);
        }
    }

    /**
     * テキストの最小サイズの半分より小さいパス（幅・高さとも）を削除する
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function removeSmallPathItems(doc) {
        var minimumTextSize = null;
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            var textFrame = doc.textFrames[textFrameIndex];
            /* 文字サイズを読めないテキストは飛ばす / skip text whose size cannot be read */
            try {
                var textSize = textFrame.textRange.characterAttributes.size;
                if (typeof textSize === 'number' && textSize > 0) {
                    if (minimumTextSize === null || textSize < minimumTextSize) minimumTextSize = textSize;
                }
            } catch (e) { }
        }
        if (minimumTextSize === null || minimumTextSize <= 0) return;

        var removalThreshold = minimumTextSize / 2;
        var pathItemsToRemove = [];
        for (var pathItemIndex = 0; pathItemIndex < doc.pathItems.length; pathItemIndex++) {
            var pathItem = doc.pathItems[pathItemIndex];
            /* 属性を読めないパスは残す / keep paths whose properties cannot be read */
            try {
                if (pathItem.clipping) continue;
                if (pathItem.locked || (pathItem.layer && pathItem.layer.locked)) continue;
                var pathBounds = pathItem.geometricBounds;
                var pathWidth = pathBounds[2] - pathBounds[0];
                var pathHeight = pathBounds[1] - pathBounds[3];
                if (pathWidth < removalThreshold && pathHeight < removalThreshold) {
                    pathItemsToRemove.push(pathItem);
                }
            } catch (e) { }
        }
        removeItemsQuietly(pathItemsToRemove, 0);
    }

    /**
     * 選択中のクリップグループごとに、重複テキストを削除してから残りを1つに連結する
     * @param {PageItem[]} selectionItems - 選択中のアイテム
     * @returns {void}
     */
    function processClipGroupTexts(selectionItems) {
        for (var selectionIndex = 0; selectionIndex < selectionItems.length; selectionIndex++) {
            var selectedItem = selectionItems[selectionIndex];
            if (selectedItem.typename !== "GroupItem" || selectedItem.clipped !== true) continue;
            deduplicateTextsInGroup(selectedItem);
            concatenateTextsInGroup(selectedItem);
        }
    }

    /**
     * グループ内で同じ文字列のテキストを1つだけ残して削除する（後ろにあるものを残す）
     * @param {GroupItem} clipGroup - 対象のグループ
     * @returns {void}
     */
    function deduplicateTextsInGroup(clipGroup) {
        var textFrames = [];
        collectTextFrames(clipGroup, textFrames);
        if (textFrames.length < 2) return;
        var seenTextContents = {};
        var textFramesToRemove = [];
        for (var textIndex = textFrames.length - 1; textIndex >= 0; textIndex--) {
            var textKey = "k:" + textFrames[textIndex].contents;
            if (seenTextContents[textKey]) {
                textFramesToRemove.push(textFrames[textIndex]);
            } else {
                seenTextContents[textKey] = true;
            }
        }
        removeItemsQuietly(textFramesToRemove, 0);
    }

    /**
     * グループ内のテキストを読み順（上から下、左から右）に1つへ連結する
     * @param {GroupItem} clipGroup - 対象のグループ
     * @returns {void}
     */
    function concatenateTextsInGroup(clipGroup) {
        var textFrames = [];
        collectTextFrames(clipGroup, textFrames);
        if (textFrames.length < 2) return;
        textFrames.sort(function (textFrameA, textFrameB) {
            var yA, yB, xA, xB;
            /* 位置を読めないものは並べ替えない / leave unreadable positions in place */
            try {
                yA = textFrameA.position[1]; yB = textFrameB.position[1];
                xA = textFrameA.position[0]; xB = textFrameB.position[0];
            } catch (e) { return 0; }
            if (Math.abs(yA - yB) > 0.5) return yB - yA;
            return xA - xB;
        });
        var combinedText = "";
        for (var textIndex = 0; textIndex < textFrames.length; textIndex++) {
            combinedText += textFrames[textIndex].contents;
        }
        /* 書き込めなければ連結をやめる / give up when the text cannot be written */
        try {
            textFrames[0].contents = combinedText;
        } catch (e) { return; }
        removeItemsQuietly(textFrames, 1);
    }

    /**
     * グループ内のテキストを再帰的に集める
     * @param {GroupItem} parentGroup - 走査するグループ
     * @param {TextFrame[]} collectedTextFrames - 見つけたテキストを追加する配列
     * @returns {void}
     */
    function collectTextFrames(parentGroup, collectedTextFrames) {
        for (var pageItemIndex = 0; pageItemIndex < parentGroup.pageItems.length; pageItemIndex++) {
            var pageItem = parentGroup.pageItems[pageItemIndex];
            if (pageItem.typename === "TextFrame") {
                collectedTextFrames.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                collectTextFrames(pageItem, collectedTextFrames);
            }
        }
    }

    /**
     * テキストの内容を読む
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string|null} 内容。読めなければ null
     */
    function readTextContents(textFrame) {
        /* 内容を読めないテキストは null / null when the contents cannot be read */
        try {
            return textFrame.contents;
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // 列ごとのテキスト配置 / Column text alignment
    // =========================================

    /**
     * 罫線から列を求め、列ごとの配置ダイアログで指定した揃え方でテキストを置き直す
     * 罫線から格子が作れないときは、数字のテキストを右端で揃えるだけにする
     * @param {Document} doc - 対象のドキュメント
     * @param {PathItem[]} ruleItems - 罫線
     * @returns {boolean} キャンセルされたら false
     */
    function alignTextFramesByColumnDialog(doc, ruleItems) {
        var grid = collectRuleCenterPositionsFromItems(ruleItems);
        if (grid.xPositions.length < 2 || grid.yPositions.length < 2) {
            alignNumericTextFramesRightByColumns(collectNumericTextFrames(doc));
            return true;
        }

        var textFramesByColumn = collectTextFramesByGridColumn(doc, grid);
        var columnCount = grid.xPositions.length - 1;

        var columnSamples = getColumnSamples(textFramesByColumn, columnCount);
        var columnAlignments = getDefaultColumnAlignments(textFramesByColumn, columnCount);
        var selectedAlignments = showColumnAlignmentDialog(columnSamples, columnAlignments);
        if (!selectedAlignments) return false;

        applyColumnTextAlignments(textFramesByColumn, grid, selectedAlignments, getTextAlignmentPadding(doc));
        return true;
    }

    /**
     * 数字（右揃えの対象）のテキストを集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} 数字のテキスト
     */
    function collectNumericTextFrames(doc) {
        var numericTextFrames = [];
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            var textFrame = doc.textFrames[textFrameIndex];
            if (isRightAlignNumericText(readTextContents(textFrame))) {
                numericTextFrames.push(textFrame);
            }
        }
        return numericTextFrames;
    }

    /**
     * 中心が格子の中にあるテキストを、列ごとに振り分ける
     * @param {Document} doc - 対象のドキュメント
     * @param {{xPositions: number[], yPositions: number[]}} grid - 罫線の格子
     * @returns {TextFrame[][]} 列ごとのテキスト
     */
    function collectTextFramesByGridColumn(doc, grid) {
        var textFramesByColumn = [];
        var columnCount = grid.xPositions.length - 1;
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            textFramesByColumn[columnIndex] = [];
        }

        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            var textFrame = doc.textFrames[textFrameIndex];
            var textCenter = getTextCenter(textFrame);
            if (!textCenter) continue;

            var textColumnIndex = findGridIntervalIndex(textCenter.x, grid.xPositions);
            var textRowIndex = findGridIntervalIndex(textCenter.y, grid.yPositions);
            if (textColumnIndex < 0 || textRowIndex < 0) continue;

            textFramesByColumn[textColumnIndex].push(textFrame);
        }
        return textFramesByColumn;
    }

    /**
     * 列ごとに、上から見て最初の空でないテキストをサンプルとして返す
     * @param {TextFrame[][]} textFramesByColumn - 列ごとのテキスト
     * @param {number} columnCount - 列の数
     * @returns {string[]} 列ごとのサンプル文字列
     */
    function getColumnSamples(textFramesByColumn, columnCount) {
        var columnSamples = [];
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            var columnTextFrames = (textFramesByColumn[columnIndex] || []).slice();
            sortTextFramesTopToBottom(columnTextFrames);

            var sampleText = "";
            for (var textIndex = 0; textIndex < columnTextFrames.length; textIndex++) {
                var normalizedContent = normalizeSampleText(readTextContents(columnTextFrames[textIndex]));
                if (normalizedContent !== "") {
                    sampleText = normalizedContent;
                    break;
                }
            }
            columnSamples.push(sampleText || getLabel("columnRow.noSample"));
        }
        return columnSamples;
    }

    /**
     * テキストを上から下、同じ高さなら左から右に並べる
     * @param {TextFrame[]} textFrames - 並べ替えるテキスト（その場で並べ替える）
     * @returns {void}
     */
    function sortTextFramesTopToBottom(textFrames) {
        textFrames.sort(function (textFrameA, textFrameB) {
            var boundsA = getTextBounds(textFrameA);
            var boundsB = getTextBounds(textFrameB);
            if (!boundsA || !boundsB) return 0;

            var topA = boundsA[1];
            var topB = boundsB[1];
            if (Math.abs(topA - topB) > 0.5) return topB - topA;

            var leftA = boundsA[0];
            var leftB = boundsB[0];
            return leftA - leftB;
        });
    }

    /**
     * サンプル文字列を1行に整え、16文字を超える分は「…」で省略する
     * @param {string|null} textContent - 元の文字列
     * @returns {string} 整えた文字列
     */
    function normalizeSampleText(textContent) {
        if (textContent === null || typeof textContent === "undefined") return "";
        var sampleText = String(textContent).replace(/[\r\n\t]+/g, " ").replace(/^\s+|\s+$/g, "");
        if (sampleText.length > 16) sampleText = sampleText.substring(0, 16) + "…";
        return sampleText;
    }

    /**
     * 列ごとの初期配置を判定する（数字が文字列以上に多い列は右、それ以外は左）
     * @param {TextFrame[][]} textFramesByColumn - 列ごとのテキスト
     * @param {number} columnCount - 列の数
     * @returns {string[]} 列ごとの "left" / "right"
     */
    function getDefaultColumnAlignments(textFramesByColumn, columnCount) {
        var columnAlignments = [];
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            var columnTextFrames = textFramesByColumn[columnIndex] || [];
            var numericCount = 0;
            var textCount = 0;
            for (var textIndex = 0; textIndex < columnTextFrames.length; textIndex++) {
                var textContents = readTextContents(columnTextFrames[textIndex]);
                if (isRightAlignNumericText(textContents)) {
                    numericCount++;
                } else if (normalizeSampleText(textContents) !== "") {
                    textCount++;
                }
            }
            columnAlignments.push(numericCount > 0 && numericCount >= textCount ? "right" : "left");
        }
        return columnAlignments;
    }

    /**
     * 列ごとの配置ダイアログに1列分の行を追加する
     * @param {Panel} rowsPanel - 追加先のパネル
     * @param {string} columnLabelText - 列番号の表示
     * @param {number} columnLabelChars - 列番号欄の幅（文字数）
     * @param {string} sampleLabelText - サンプルの表示
     * @returns {{left: RadioButton, center: RadioButton, right: RadioButton}} 左・中央・右のラジオ
     */
    function addColumnAlignmentRow(rowsPanel, columnLabelText, columnLabelChars, sampleLabelText) {
        var columnRow = rowsPanel.add("group");
        columnRow.orientation = "row";
        columnRow.alignChildren = "center";

        var columnLabel = columnRow.add("statictext", undefined, columnLabelText);
        columnLabel.characters = columnLabelChars;

        var leftRadio = addOptionRadio(columnRow, "radio.alignLeft", "tooltip.columnAlignment");
        var centerRadio = addOptionRadio(columnRow, "radio.alignCenter", "tooltip.columnAlignment");
        var rightRadio = addOptionRadio(columnRow, "radio.alignRight", "tooltip.columnAlignment");

        var sampleLabel = columnRow.add("statictext", undefined, sampleLabelText);
        sampleLabel.characters = SAMPLE_LABEL_CHARS;

        return { left: leftRadio, center: centerRadio, right: rightRadio };
    }

    /**
     * 列ごとの配置ダイアログを表示する
     * @param {string[]} columnSamples - 列ごとのサンプル文字列
     * @param {string[]} initialAlignments - 列ごとの初期配置
     * @returns {string[]|null} 列ごとの "left" / "center" / "right"。キャンセルなら null
     */
    function showColumnAlignmentDialog(columnSamples, initialAlignments) {
        var columnDialog = new Window('dialog', getLabel("dialog.columnAlignmentTitle"));
        columnDialog.orientation = "column";
        columnDialog.alignChildren = "fill";
        columnDialog.spacing = DIALOG_SPACING;
        columnDialog.margins = DIALOG_MARGINS;

        var rowsPanel = columnDialog.add("panel", undefined, getLabel("dialog.columnAlignmentTitle"));
        setupPanel(rowsPanel, 6);
        rowsPanel.alignChildren = "fill";

        var columnLabelTexts = [];
        var sampleLabelTexts = [];
        var columnLabelChars = 0;
        for (var labelIndex = 0; labelIndex < columnSamples.length; labelIndex++) {
            var columnLabelText = getColumnAlignmentColumnLabel(labelIndex);
            var rawSample = columnSamples[labelIndex] || "";
            columnLabelTexts.push(columnLabelText);
            sampleLabelTexts.push(rawSample.replace(/^[\(（]/, "").replace(/[\)）]$/, ""));
            if (columnLabelText.length > columnLabelChars) columnLabelChars = columnLabelText.length;
        }
        if (columnLabelChars < 4) columnLabelChars = 4;

        var radioRows = [];
        for (var columnIndex = 0; columnIndex < columnSamples.length; columnIndex++) {
            radioRows.push(addColumnAlignmentRow(rowsPanel, columnLabelTexts[columnIndex], columnLabelChars, sampleLabelTexts[columnIndex]));
            setColumnAlignmentRadioValue(radioRows[columnIndex], initialAlignments[columnIndex] || "left");
        }

        var btnRowGroup = columnDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = "fill";
        btnRowGroup.margins = BUTTON_ROW_MARGINS;

        // 左：自動判定 / Left: Auto
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignment = ["left", "center"];
        var btnAuto = btnLeftGroup.add("button", undefined, getLabel("button.autoAlignment"));
        btnAuto.helpTip = getLabel("tooltip.autoAlignment");

        // 中央：スペーサー / Center: spacer
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        // 右：キャンセル / OK / Right: Cancel / OK
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignment = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnAuto.onClick = function () {
            for (var i = 0; i < radioRows.length; i++) {
                setColumnAlignmentRadioValue(radioRows[i], initialAlignments[i] || "left");
            }
        };

        var selectedAlignments = null;
        btnOK.onClick = function () {
            selectedAlignments = [];
            for (var rowIndex = 0; rowIndex < radioRows.length; rowIndex++) {
                selectedAlignments.push(getColumnAlignmentRadioValue(radioRows[rowIndex]));
            }
            columnDialog.close();
        };
        btnCancel.onClick = function () {
            selectedAlignments = null;
            columnDialog.close();
        };

        columnDialog.show();
        return selectedAlignments;
    }

    /**
     * 列番号の表示を返す（日本語は「1列目」、英語は「Column 1」）
     * @param {number} columnIndex - 列の番号（0 始まり）
     * @returns {string} 列番号の表示
     */
    function getColumnAlignmentColumnLabel(columnIndex) {
        if (uiLang === "ja") {
            return String(columnIndex + 1) + getLabel("columnRow.columnNumber");
        }
        return getLabel("columnRow.columnNumber") + " " + String(columnIndex + 1);
    }

    /**
     * 配置のラジオを指定の値にする
     * @param {{left: RadioButton, center: RadioButton, right: RadioButton}} radioRow - 1列分のラジオ
     * @param {string} alignment - "left" / "center" / "right"
     * @returns {void}
     */
    function setColumnAlignmentRadioValue(radioRow, alignment) {
        radioRow.left.value = alignment === "left";
        radioRow.center.value = alignment === "center";
        radioRow.right.value = alignment === "right";
    }

    /**
     * 配置のラジオの値を読む
     * @param {{left: RadioButton, center: RadioButton, right: RadioButton}} radioRow - 1列分のラジオ
     * @returns {string} "left" / "center" / "right"
     */
    function getColumnAlignmentRadioValue(radioRow) {
        if (radioRow.center.value) return "center";
        if (radioRow.right.value) return "right";
        return "left";
    }

    /**
     * 昇順の座標の中で、値が入る区間の番号を返す
     * @param {number} value - 調べる値
     * @param {number[]} sortedPositions - 昇順の座標
     * @returns {number} 区間の番号。どこにも入らなければ -1
     */
    function findGridIntervalIndex(value, sortedPositions) {
        if (!sortedPositions || sortedPositions.length < 2) return -1;
        for (var positionIndex = 0; positionIndex < sortedPositions.length - 1; positionIndex++) {
            if (value >= sortedPositions[positionIndex] && value <= sortedPositions[positionIndex + 1]) {
                return positionIndex;
            }
        }
        return -1;
    }

    /**
     * テキストの中心座標を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {{x: number, y: number}|null} 中心座標。境界を読めなければ null
     */
    function getTextCenter(textFrame) {
        var textBounds = getTextBounds(textFrame);
        if (!textBounds) return null;
        return {
            x: (textBounds[0] + textBounds[2]) / 2,
            y: (textBounds[1] + textBounds[3]) / 2
        };
    }

    /**
     * テキストの境界を返す（visibleBounds、読めなければ geometricBounds）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[]|null} [left, top, right, bottom]。どちらも読めなければ null
     */
    function getTextBounds(textFrame) {
        try {
            return textFrame.visibleBounds;
        } catch (e) {
            try {
                return textFrame.geometricBounds;
            } catch (e2) {
                return null;
            }
        }
    }

    /**
     * テキストの右端を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number|null} 右端の X。境界を読めなければ null
     */
    function getTextRight(textFrame) {
        var textBounds = getTextBounds(textFrame);
        return textBounds ? textBounds[2] : null;
    }

    /**
     * 列の端からテキストまでの余白を返す（最頻の文字サイズの 35%、求められなければ 2pt）
     * @param {Document} doc - 対象のドキュメント
     * @returns {number} 余白（pt）
     */
    function getTextAlignmentPadding(doc) {
        var modeTextSize = getModeTextSize(doc);
        if (modeTextSize && modeTextSize > 0) return modeTextSize * 0.35;
        return 2;
    }

    /**
     * 列ごとの配置をテキストに適用する
     * @param {TextFrame[][]} textFramesByColumn - 列ごとのテキスト
     * @param {{xPositions: number[], yPositions: number[]}} grid - 罫線の格子
     * @param {string[]} columnAlignments - 列ごとの配置
     * @param {number} padding - 列の端からの余白（pt）
     * @returns {void}
     */
    function applyColumnTextAlignments(textFramesByColumn, grid, columnAlignments, padding) {
        for (var columnIndex = 0; columnIndex < columnAlignments.length; columnIndex++) {
            var columnTextFrames = textFramesByColumn[columnIndex] || [];
            for (var textIndex = 0; textIndex < columnTextFrames.length; textIndex++) {
                alignTextFrameToGridColumn(columnTextFrames[textIndex], grid, columnIndex, columnAlignments[columnIndex], padding);
            }
        }
    }

    /**
     * テキストを列内の指定位置（左・中央・右）に置く
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {{xPositions: number[]}} grid - 罫線の格子
     * @param {number} columnIndex - 列の番号
     * @param {string} alignment - "left" / "center" / "right"
     * @param {number} padding - 列の端からの余白（pt）
     * @returns {void}
     */
    function alignTextFrameToGridColumn(textFrame, grid, columnIndex, alignment, padding) {
        if (columnIndex < 0 || columnIndex >= grid.xPositions.length - 1) return;

        var columnLeft = grid.xPositions[columnIndex];
        var columnRight = grid.xPositions[columnIndex + 1];
        if (alignment === "right") {
            alignToRightEdge(textFrame, columnRight - padding);
        } else if (alignment === "center") {
            alignToCenterX(textFrame, (columnLeft + columnRight) / 2);
        } else {
            alignToLeftEdge(textFrame, columnLeft + padding);
        }
    }

    /**
     * 行揃えを設定してから、境界の基準位置が targetX に来るよう横に動かす
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string} justificationName - Justification の名前（"LEFT" / "CENTER" / "RIGHT"）
     * @param {Function} measureX - 境界から基準位置の X を求める関数
     * @param {number} targetX - 基準位置を合わせる X
     * @returns {void}
     */
    function alignTextFrameHorizontally(textFrame, justificationName, measureX, targetX) {
        if (!getTextBounds(textFrame)) return;

        safeCall(function () {
            textFrame.textRange.paragraphAttributes.justification = Justification[justificationName];
        });

        var textBounds = getTextBounds(textFrame);
        if (!textBounds) return;
        var dx = targetX - measureX(textBounds);
        if (Math.abs(dx) > 0.001) {
            safeCall(function () { textFrame.translate(dx, 0); });
        }
    }

    /**
     * 左揃えにして左端を合わせる
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {number} targetLeft - 左端の X
     * @returns {void}
     */
    function alignToLeftEdge(textFrame, targetLeft) {
        alignTextFrameHorizontally(textFrame, "LEFT", function (textBounds) { return textBounds[0]; }, targetLeft);
    }

    /**
     * 中央揃えにして中心を合わせる
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {number} targetCenterX - 中心の X
     * @returns {void}
     */
    function alignToCenterX(textFrame, targetCenterX) {
        alignTextFrameHorizontally(textFrame, "CENTER", function (textBounds) { return (textBounds[0] + textBounds[2]) / 2; }, targetCenterX);
    }

    /**
     * 右揃えにして右端を合わせる
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {number} targetRight - 右端の X
     * @returns {void}
     */
    function alignToRightEdge(textFrame, targetRight) {
        alignTextFrameHorizontally(textFrame, "RIGHT", function (textBounds) { return textBounds[2]; }, targetRight);
    }

    /**
     * 罫線の格子が作れないときの代わり：右端の近い数字テキストを列とみなし、列内の最も右に揃える
     * @param {TextFrame[]} numericTextFrames - 数字のテキスト
     * @returns {void}
     */
    function alignNumericTextFramesRightByColumns(numericTextFrames) {
        var rightEdgeColumns = [];
        var columnTolerance = 10;

        for (var textIndex = 0; textIndex < numericTextFrames.length; textIndex++) {
            var textFrame = numericTextFrames[textIndex];
            var textRight = getTextRight(textFrame);
            if (textRight === null) continue;

            var isAssigned = false;
            for (var columnIndex = 0; columnIndex < rightEdgeColumns.length; columnIndex++) {
                if (Math.abs(rightEdgeColumns[columnIndex].x - textRight) < columnTolerance) {
                    rightEdgeColumns[columnIndex].items.push(textFrame);
                    rightEdgeColumns[columnIndex].x = (rightEdgeColumns[columnIndex].x + textRight) / 2;
                    isAssigned = true;
                    break;
                }
            }

            if (!isAssigned) {
                rightEdgeColumns.push({ x: textRight, items: [textFrame] });
            }
        }

        for (var rightColumnIndex = 0; rightColumnIndex < rightEdgeColumns.length; rightColumnIndex++) {
            var columnTextFrames = rightEdgeColumns[rightColumnIndex].items;
            var maxRight = null;
            for (var measureIndex = 0; measureIndex < columnTextFrames.length; measureIndex++) {
                var itemRight = getTextRight(columnTextFrames[measureIndex]);
                if (itemRight === null) continue;
                if (maxRight === null || itemRight > maxRight) maxRight = itemRight;
            }
            if (maxRight === null) continue;
            for (var alignIndex = 0; alignIndex < columnTextFrames.length; alignIndex++) {
                alignToRightEdge(columnTextFrames[alignIndex], maxRight);
            }
        }
    }

    /**
     * 右揃えにする数字かどうかを判定する（空白・円記号・桁区切り・全角マイナス・▲を除いて数値になるもの）
     * @param {string|null} textContent - 文字列
     * @returns {boolean} 数字なら true
     */
    function isRightAlignNumericText(textContent) {
        if (textContent === null || typeof textContent === "undefined") return false;
        var normalizedText = String(textContent)
            .replace(/[\r\n\t ]+/g, "")
            .replace(/[￥¥]/g, "")
            .replace(/,/g, "")
            .replace(/，/g, "")
            .replace(/−/g, "-")
            .replace(/－/g, "-")
            .replace(/―/g, "-")
            .replace(/▲/g, "-");

        if (normalizedText === "") return false;
        return /^-?\d+(\.\d+)?$/.test(normalizedText);
    }

    // =========================================
    // テキストの色とレイヤー / Text color and layers
    // =========================================

    /**
     * すべてのテキストの文字色を K100 にする
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function applyK100ToAllTexts(doc) {
        var blackColor = getK100Black();
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            var textFrame = doc.textFrames[textFrameIndex];
            /* 色を変えられないテキストは飛ばす / skip text that cannot be recolored */
            try {
                textFrame.textRange.characterAttributes.fillColor = blackColor;
            } catch (e) { }
        }
    }

    /**
     * すべてのテキストを指定のレイヤーへ移す（レイヤーが無ければ作り、ロック解除・表示にする）
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - 移動先のレイヤー名
     * @returns {void}
     */
    function moveTextsToLayer(doc, layerName) {
        var targetLayer = findOrCreateLayer(doc, layerName);
        if (targetLayer.locked) targetLayer.locked = false;
        if (!targetLayer.visible) targetLayer.visible = true;

        var textFrames = [];
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            textFrames.push(doc.textFrames[textFrameIndex]);
        }
        for (var moveIndex = 0; moveIndex < textFrames.length; moveIndex++) {
            var textFrame = textFrames[moveIndex];
            if (textFrame.parent === targetLayer) continue;
            safeCall(function () { textFrame.move(targetLayer, ElementPlacement.PLACEATBEGINNING); });
        }
    }

    /**
     * グループを解除する（中身を親の末尾へ移してからグループを削除）
     * @param {GroupItem} groupItem - 解除するグループ
     * @returns {void}
     */
    function ungroup(groupItem) {
        var parentContainer = groupItem.parent;
        while (groupItem.pageItems.length > 0) {
            groupItem.pageItems[0].move(parentContainer, ElementPlacement.PLACEATEND);
        }
        groupItem.remove();
    }

    /**
     * K100 の CMYK カラーを作る
     * @returns {CMYKColor} K100 の色
     */
    function getK100Black() {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    // =========================================
    // 罫線の中心線化と配置 / Rule centerlines and placement
    // =========================================

    /**
     * 対象レイヤーの細長い長方形を中心線に置き換える
     * @param {Document} doc - 対象のドキュメント
     * @param {string[]} excludeLayerNames - 除外するレイヤー名
     * @returns {PathItem[]} 作成した中心線
     */
    function convertRectanglesToCenterLines(doc, excludeLayerNames) {
        var rectangleItems = [];
        var targetLayers = getTargetLayers(doc, excludeLayerNames);
        for (var layerIndex = 0; layerIndex < targetLayers.length; layerIndex++) {
            collectRectsForCenterline(targetLayers[layerIndex].pageItems, rectangleItems);
        }
        var centerLineItems = [];
        for (var rectangleIndex = 0; rectangleIndex < rectangleItems.length; rectangleIndex++) {
            /* 変換できない長方形は残す / leave rectangles that cannot be converted */
            try {
                var centerLine = convertRectToCenterLine(rectangleItems[rectangleIndex]);
                if (centerLine) centerLineItems.push(centerLine);
            } catch (e) { }
        }
        return centerLineItems;
    }

    /*
     目的:
     Excel由来の罫線は位置や本数が不揃いで、セルグリッドとして扱いにくい。
     この関数は線を水平／垂直に分類し、外接範囲から行列グリッドを再構築するための入口。
     配置モードに応じて、均等配置・結合セル対応・列幅保持のいずれかを適用する。

     Purpose:
     Lines imported from Excel are often inconsistent in position and count.
     This function classifies lines into horizontal/vertical and rebuilds a grid
     from their bounds, then delegates to a specific placement strategy.
    */
    /**
     * 配置モードに応じて罫線を並べ直す
     * @param {PathItem[]} lines - 2点の罫線
     * @param {string} placementMode - "forced" / "merged" / "keepColumnWidths"
     * @param {Object|null} outerBounds - 元の選択範囲の外形（left / top / right / bottom）
     * @returns {void}
     */
    function applyPlacementMode(lines, placementMode, outerBounds) {
        if (!lines || lines.length === 0) return;

        var horizontalLines = [];
        var verticalLines = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var lineInfo = getLineInfo(lines[lineIndex]);
            if (!lineInfo) continue;
            if (lineInfo.orient === "h") {
                horizontalLines.push({ path: lines[lineIndex], y: lineInfo.coord, minX: lineInfo.minSpan, maxX: lineInfo.maxSpan });
            } else {
                verticalLines.push({ path: lines[lineIndex], x: lineInfo.coord, minY: lineInfo.minSpan, maxY: lineInfo.maxSpan });
            }
        }
        if (horizontalLines.length === 0 || verticalLines.length === 0) return;

        var xRange = getValueRange(verticalLines, "x");
        var yRange = getValueRange(horizontalLines, "y");
        var lineBounds = { minX: xRange.min, maxX: xRange.max, minY: yRange.min, maxY: yRange.max };

        // 外周に罫線がない辺は、表全体の外形（セル背景＋罫線の合算バウンズ）まで拡張
        // For edges without rules, expand bounds to the full table outline (cell backgrounds + rules)
        if (outerBounds) {
            var edgeTolerance = 0.5;
            if (outerBounds.left < lineBounds.minX - edgeTolerance) lineBounds.minX = outerBounds.left;
            if (outerBounds.right > lineBounds.maxX + edgeTolerance) lineBounds.maxX = outerBounds.right;
            if (outerBounds.bottom < lineBounds.minY - edgeTolerance) lineBounds.minY = outerBounds.bottom;
            if (outerBounds.top > lineBounds.maxY + edgeTolerance) lineBounds.maxY = outerBounds.top;
        }

        if (placementMode === "forced") {
            alignLines(horizontalLines, verticalLines, lineBounds, true);
        } else if (placementMode === "merged") {
            alignLinesWithMergedCells(horizontalLines, verticalLines, lineBounds);
        } else if (placementMode === "keepColumnWidths") {
            alignLines(horizontalLines, verticalLines, lineBounds, false);
        }
    }

    /**
     * 線の記録から、指定プロパティの最小値と最大値を求める
     * @param {Object[]} lines - 線の記録（1本以上）
     * @param {string} propertyName - 調べるプロパティ名（x / y）
     * @returns {{min: number, max: number}} 最小値と最大値
     */
    function getValueRange(lines, propertyName) {
        var minValue = lines[0][propertyName];
        var maxValue = lines[0][propertyName];
        for (var lineIndex = 1; lineIndex < lines.length; lineIndex++) {
            if (lines[lineIndex][propertyName] < minValue) minValue = lines[lineIndex][propertyName];
            if (lines[lineIndex][propertyName] > maxValue) maxValue = lines[lineIndex][propertyName];
        }
        return { min: minValue, max: maxValue };
    }

    /**
     * 罫線を揃える：水平線は常に Y を等分。垂直線は evenColumns=true で X も等分、false で元の X を保持
     * @param {Object[]} horizontalLines - 水平線の記録
     * @param {Object[]} verticalLines - 垂直線の記録
     * @param {{minX: number, maxX: number, minY: number, maxY: number}} lineBounds - 表の外形
     * @param {boolean} evenColumns - 列も等分するかどうか
     * @returns {void}
     */
    function alignLines(horizontalLines, verticalLines, lineBounds, evenColumns) {
        var horizontalYPositions = distributePositions(horizontalLines, "y", lineBounds.minY, lineBounds.maxY);
        var verticalXPositions = evenColumns
            ? distributePositions(verticalLines, "x", lineBounds.minX, lineBounds.maxX)
            : null;

        var topY = Math.max(lineBounds.minY, lineBounds.maxY);
        var bottomY = Math.min(lineBounds.minY, lineBounds.maxY);

        for (var horizontalIndex = 0; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            setLineEndpointsSafe(horizontalLines[horizontalIndex].path,
                [lineBounds.minX, horizontalYPositions[horizontalIndex]],
                [lineBounds.maxX, horizontalYPositions[horizontalIndex]]);
        }
        for (var verticalIndex = 0; verticalIndex < verticalLines.length; verticalIndex++) {
            var lineX = verticalXPositions ? verticalXPositions[verticalIndex] : verticalLines[verticalIndex].x;
            setLineEndpointsSafe(verticalLines[verticalIndex].path, [lineX, topY], [lineX, bottomY]);
        }
    }

    /**
     * 線を minValue..maxValue で等分した位置を返す（線は昇順に並べ替える）。1本以下なら元の値を返す
     * @param {Object[]} lines - 線の記録（その場で並べ替える）
     * @param {string} propertyName - 位置のプロパティ名
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @returns {number[]} 新しい位置
     */
    function distributePositions(lines, propertyName, minValue, maxValue) {
        var positions = [];
        if (lines.length >= 2) {
            lines.sort(function (lineA, lineB) { return lineA[propertyName] - lineB[propertyName]; });
            var step = (maxValue - minValue) / (lines.length - 1);
            for (var i = 0; i < lines.length; i++) {
                positions.push(minValue + step * i);
            }
        } else {
            for (var j = 0; j < lines.length; j++) {
                positions.push(lines[j][propertyName]);
            }
        }
        return positions;
    }

    /**
     * 2点の線の両端を設定する（方向ハンドルもアンカーに揃える。失敗しても続行）
     * @param {PathItem} linePath - 2点の線
     * @param {number[]} startPoint - 始点 [x, y]
     * @param {number[]} endPoint - 終点 [x, y]
     * @returns {void}
     */
    function setLineEndpointsSafe(linePath, startPoint, endPoint) {
        safeCall(function () {
            linePath.pathPoints[0].anchor = startPoint;
            linePath.pathPoints[0].leftDirection = startPoint;
            linePath.pathPoints[0].rightDirection = startPoint;
            linePath.pathPoints[1].anchor = endPoint;
            linePath.pathPoints[1].leftDirection = endPoint;
            linePath.pathPoints[1].rightDirection = endPoint;
        });
    }

    /*
     目的:
     結合セルがある表では、単純な等間隔配置では列境界と一致しない。
     近接座標で線をグルーピングし、実際の列／行境界にスナップさせることで、
     結合セルを壊さずにグリッドを再構築する。

     Purpose:
     With merged cells, uniform spacing breaks true column boundaries.
     This groups lines by nearby coordinates and snaps them to detected
     row/column edges, preserving merged-cell structure.
    */
    /**
     * 結合セルに対応して罫線を揃える（同じ座標の線を1行／1列とし、両端は最寄りの交点に合わせる）
     * @param {Object[]} horizontalLines - 水平線の記録
     * @param {Object[]} verticalLines - 垂直線の記録
     * @param {{minX: number, maxX: number, minY: number, maxY: number}} lineBounds - 表の外形
     * @returns {void}
     */
    function alignLinesWithMergedCells(horizontalLines, verticalLines, lineBounds) {
        var coordinateTolerance = 5.0;
        var horizontalGroups = groupLinesByCoordinate(horizontalLines, "y", coordinateTolerance);
        var verticalGroups = groupLinesByCoordinate(verticalLines, "x", coordinateTolerance);

        var horizontalYPositions = distributePositions(horizontalGroups, "coord", lineBounds.minY, lineBounds.maxY);
        var verticalXPositions = distributePositions(verticalGroups, "coord", lineBounds.minX, lineBounds.maxX);

        for (var rowGroupIndex = 0; rowGroupIndex < horizontalGroups.length; rowGroupIndex++) {
            var rowY = horizontalYPositions[rowGroupIndex];
            var rowLines = horizontalGroups[rowGroupIndex].lines;
            for (var rowLineIndex = 0; rowLineIndex < rowLines.length; rowLineIndex++) {
                var horizontalLine = rowLines[rowLineIndex];
                var leftColumnIndex = findClosestIndexByProperty(horizontalLine.minX, verticalGroups, "coord");
                var rightColumnIndex = findClosestIndexByProperty(horizontalLine.maxX, verticalGroups, "coord");
                setLineEndpointsSafe(horizontalLine.path, [verticalXPositions[leftColumnIndex], rowY], [verticalXPositions[rightColumnIndex], rowY]);
            }
        }
        for (var columnGroupIndex = 0; columnGroupIndex < verticalGroups.length; columnGroupIndex++) {
            var columnX = verticalXPositions[columnGroupIndex];
            var columnLines = verticalGroups[columnGroupIndex].lines;
            for (var columnLineIndex = 0; columnLineIndex < columnLines.length; columnLineIndex++) {
                var verticalLine = columnLines[columnLineIndex];
                var topRowIndex = findClosestIndexByProperty(verticalLine.maxY, horizontalGroups, "coord");
                var bottomRowIndex = findClosestIndexByProperty(verticalLine.minY, horizontalGroups, "coord");
                setLineEndpointsSafe(verticalLine.path, [columnX, horizontalYPositions[topRowIndex]], [columnX, horizontalYPositions[bottomRowIndex]]);
            }
        }
    }

    /**
     * 線を座標の近さでグループにまとめる（昇順、coord はグループ内の平均）
     * @param {Object[]} lines - 線の記録
     * @param {string} propertyName - まとめる基準のプロパティ名（x / y）
     * @param {number} tolerance - 同じ座標とみなす差
     * @returns {{coord: number, lines: Object[]}[]} 座標ごとのグループ
     */
    function groupLinesByCoordinate(lines, propertyName, tolerance) {
        var sortedLines = lines.slice();
        sortedLines.sort(function (lineA, lineB) { return lineA[propertyName] - lineB[propertyName]; });
        var coordinateGroups = [];
        for (var i = 0; i < sortedLines.length; i++) {
            var currentLine = sortedLines[i];
            var currentCoordinate = currentLine[propertyName];
            if (coordinateGroups.length === 0 || Math.abs(currentCoordinate - coordinateGroups[coordinateGroups.length - 1].coord) > tolerance) {
                coordinateGroups.push({ coord: currentCoordinate, lines: [currentLine] });
            } else {
                var lastGroup = coordinateGroups[coordinateGroups.length - 1];
                lastGroup.lines.push(currentLine);
                var coordinateSum = 0;
                for (var k = 0; k < lastGroup.lines.length; k++) coordinateSum += lastGroup.lines[k][propertyName];
                lastGroup.coord = coordinateSum / lastGroup.lines.length;
            }
        }
        return coordinateGroups;
    }

    /**
     * 指定プロパティが targetValue に最も近い要素のインデックスを返す
     * @param {number} targetValue - 基準の値
     * @param {Object[]} items - 候補（1件以上）
     * @param {string} propertyName - 比べるプロパティ名
     * @returns {number} 最も近い要素のインデックス
     */
    function findClosestIndexByProperty(targetValue, items, propertyName) {
        var closestIndex = 0;
        var closestDistance = Math.abs(targetValue - items[0][propertyName]);
        for (var i = 1; i < items.length; i++) {
            var currentDistance = Math.abs(targetValue - items[i][propertyName]);
            if (currentDistance < closestDistance) {
                closestIndex = i;
                closestDistance = currentDistance;
            }
        }
        return closestIndex;
    }

    /**
     * 同じ向き・同じ座標（0.5pt 以内）で重なる線を1本にまとめる（長さは和集合）
     * @param {PathItem[]} lines - 線
     * @returns {PathItem[]} まとめた後の線
     */
    function dedupeOverlappingLines(lines) {
        if (!lines || lines.length < 2) return lines || [];
        var coordinateTolerance = 0.5;
        var lineGroups = [];

        for (var i = 0; i < lines.length; i++) {
            var lineInfo = getLineInfo(lines[i]);
            if (!lineInfo) {
                lineGroups.push({ orient: null, coord: 0, items: [{ path: lines[i], minSpan: 0, maxSpan: 0 }] });
                continue;
            }
            var matchedGroup = null;
            for (var groupIndex = 0; groupIndex < lineGroups.length; groupIndex++) {
                var candidateGroup = lineGroups[groupIndex];
                if (candidateGroup.orient !== lineInfo.orient) continue;
                if (Math.abs(candidateGroup.coord - lineInfo.coord) > coordinateTolerance) continue;
                matchedGroup = candidateGroup;
                break;
            }
            var groupEntry = { path: lines[i], minSpan: lineInfo.minSpan, maxSpan: lineInfo.maxSpan };
            if (matchedGroup) {
                matchedGroup.items.push(groupEntry);
            } else {
                lineGroups.push({ orient: lineInfo.orient, coord: lineInfo.coord, items: [groupEntry] });
            }
        }

        var keptLines = [];
        for (var keptIndex = 0; keptIndex < lineGroups.length; keptIndex++) {
            var lineGroup = lineGroups[keptIndex];
            if (lineGroup.items.length === 1 || lineGroup.orient === null) {
                keptLines.push(lineGroup.items[0].path);
                continue;
            }
            var unionMin = lineGroup.items[0].minSpan;
            var unionMax = lineGroup.items[0].maxSpan;
            for (var itemIndex = 1; itemIndex < lineGroup.items.length; itemIndex++) {
                if (lineGroup.items[itemIndex].minSpan < unionMin) unionMin = lineGroup.items[itemIndex].minSpan;
                if (lineGroup.items[itemIndex].maxSpan > unionMax) unionMax = lineGroup.items[itemIndex].maxSpan;
            }
            var keeperPath = lineGroup.items[0].path;
            if (lineGroup.orient === "h") {
                setLineEndpointsSafe(keeperPath, [unionMin, lineGroup.coord], [unionMax, lineGroup.coord]);
            } else {
                setLineEndpointsSafe(keeperPath, [lineGroup.coord, unionMax], [lineGroup.coord, unionMin]);
            }
            for (var removeIndex = 1; removeIndex < lineGroup.items.length; removeIndex++) {
                var duplicatePath = lineGroup.items[removeIndex].path;
                safeCall(function () { duplicatePath.remove(); });
            }
            keptLines.push(keeperPath);
        }
        return keptLines;
    }

    /**
     * 2点の線の向きと位置を返す
     * @param {PathItem} linePath - 線
     * @returns {{orient: string, coord: number, minSpan: number, maxSpan: number}|null} 水平（"h"）なら coord は Y、垂直（"v"）なら X。2点の線でなければ null
     */
    function getLineInfo(linePath) {
        /* 点を読めない線は null / null when the points cannot be read */
        try {
            if (!linePath.pathPoints || linePath.pathPoints.length !== 2) return null;
            var startAnchor = linePath.pathPoints[0].anchor;
            var endAnchor = linePath.pathPoints[1].anchor;
            var horizontalDistance = Math.abs(startAnchor[0] - endAnchor[0]);
            var verticalDistance = Math.abs(startAnchor[1] - endAnchor[1]);
            if (horizontalDistance >= verticalDistance) {
                return {
                    orient: "h",
                    coord: (startAnchor[1] + endAnchor[1]) / 2,
                    minSpan: Math.min(startAnchor[0], endAnchor[0]),
                    maxSpan: Math.max(startAnchor[0], endAnchor[0])
                };
            }
            return {
                orient: "v",
                coord: (startAnchor[0] + endAnchor[0]) / 2,
                minSpan: Math.min(startAnchor[1], endAnchor[1]),
                maxSpan: Math.max(startAnchor[1], endAnchor[1])
            };
        } catch (e) { return null; }
    }

    /**
     * 外周4本の線を1つの長方形に置き換える
     * @param {PathItem[]} lines - 罫線
     * @returns {PathItem[]} 外周の線を除き、長方形を加えた罫線（置き換えられなければ lines のまま）
     */
    function convertOuterFrameToRectangle(lines) {
        if (!lines || lines.length < 4) return lines || [];

        var horizontalLines = [];
        var verticalLines = [];
        for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
            var lineInfo = getLineInfo(lines[lineIndex]);
            if (!lineInfo) continue;
            if (lineInfo.orient === "h") {
                horizontalLines.push({ path: lines[lineIndex], y: lineInfo.coord });
            } else {
                verticalLines.push({ path: lines[lineIndex], x: lineInfo.coord });
            }
        }
        if (horizontalLines.length < 2 || verticalLines.length < 2) return lines;

        var topLine = horizontalLines[0], bottomLine = horizontalLines[0];
        for (var horizontalIndex = 1; horizontalIndex < horizontalLines.length; horizontalIndex++) {
            if (horizontalLines[horizontalIndex].y > topLine.y) topLine = horizontalLines[horizontalIndex];
            if (horizontalLines[horizontalIndex].y < bottomLine.y) bottomLine = horizontalLines[horizontalIndex];
        }
        var leftLine = verticalLines[0], rightLine = verticalLines[0];
        for (var verticalIndex = 1; verticalIndex < verticalLines.length; verticalIndex++) {
            if (verticalLines[verticalIndex].x < leftLine.x) leftLine = verticalLines[verticalIndex];
            if (verticalLines[verticalIndex].x > rightLine.x) rightLine = verticalLines[verticalIndex];
        }

        var outerCandidates = [topLine.path, bottomLine.path, leftLine.path, rightLine.path];
        var uniqueOuterLines = [];
        for (var candidateIndex = 0; candidateIndex < outerCandidates.length; candidateIndex++) {
            if (!containsItem(uniqueOuterLines, outerCandidates[candidateIndex])) uniqueOuterLines.push(outerCandidates[candidateIndex]);
        }

        var minX = leftLine.x, maxX = rightLine.x;
        var minY = bottomLine.y, maxY = topLine.y;

        var referencePath = topLine.path;
        var parentContainer = referencePath.parent;
        var rectanglePath;
        /* 長方形を作れなければ何もしない / do nothing when the rectangle cannot be created */
        try {
            rectanglePath = parentContainer.pathItems.rectangle(maxY, minX, maxX - minX, maxY - minY);
        } catch (e) {
            return lines;
        }
        rectanglePath.filled = false;
        rectanglePath.stroked = referencePath.stroked;
        if (referencePath.stroked) {
            safeCall(function () { rectanglePath.strokeColor = referencePath.strokeColor; });
            safeCall(function () { rectanglePath.strokeWidth = referencePath.strokeWidth; });
            safeCall(function () { rectanglePath.strokeCap = referencePath.strokeCap; });
        }
        safeCall(function () { rectanglePath.move(referencePath, ElementPlacement.PLACEAFTER); });

        removeItemsQuietly(uniqueOuterLines, 0);

        var remainingLines = [];
        for (var resultIndex = 0; resultIndex < lines.length; resultIndex++) {
            if (!containsItem(uniqueOuterLines, lines[resultIndex])) remainingLines.push(lines[resultIndex]);
        }
        remainingLines.push(rectanglePath);
        return remainingLines;
    }

    /**
     * 線幅をそろえる（線を有効にする）
     * @param {PageItem[]} items - 対象の線
     * @param {number} widthPt - 線幅（pt）
     * @returns {void}
     */
    function applyUniformStrokeWidth(items, widthPt) {
        if (!items || items.length === 0 || !widthPt || widthPt <= 0) return;
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var strokeItem = items[itemIndex];
            safeCall(function () {
                strokeItem.stroked = true;
                strokeItem.strokeWidth = widthPt;
            });
        }
    }

    /**
     * 線の色を K100 にする（線を有効にする）
     * @param {PageItem[]} items - 対象の線
     * @returns {void}
     */
    function applyK100ToLines(items) {
        if (!items || items.length === 0) return;
        var blackColor = getK100Black();
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var strokeItem = items[itemIndex];
            safeCall(function () {
                strokeItem.stroked = true;
                strokeItem.strokeColor = blackColor;
            });
        }
    }

    // =========================================
    // セル背景のスナップ / Snapping cell backgrounds
    // =========================================

    /**
     * レイヤー横断で、セル背景（塗りのあるクローズパス）を罫線の中心線にスナップする
     * @param {Document} doc - 対象のドキュメント
     * @param {string[]} excludeLayerNames - 除外するレイヤー名
     * @param {PathItem[]} ruleItems - 罫線
     * @param {Object|null} selectionBounds - 元の選択範囲の外形（これに収まるセル背景だけを対象にする）
     * @returns {number} スナップしたセル背景の数
     */
    function snapCellBackgroundsAcrossLayers(doc, excludeLayerNames, ruleItems, selectionBounds) {
        var ruleCenters = collectRuleCenterPositionsFromItems(ruleItems);
        if (ruleCenters.xPositions.length < 2 || ruleCenters.yPositions.length < 2) return 0;

        var rectangleItems = [];
        var targetLayers = getTargetLayers(doc, excludeLayerNames);
        for (var layerIndex = 0; layerIndex < targetLayers.length; layerIndex++) {
            collectCellBackgroundSnapTargetsInItems(targetLayers[layerIndex].pageItems, rectangleItems);
        }

        // 元の選択範囲外のセル背景は対象外にする
        // Limit cell backgrounds to those inside the original selection bounds
        if (selectionBounds) {
            rectangleItems = filterItemsWithinBounds(rectangleItems, selectionBounds, 1.0);
        }

        if (rectangleItems.length === 0) return 0;

        // 選択範囲の外周に罫線がない辺は、選択範囲の外形を仮想罫線として補う
        // For edges without rules, augment ruleCenters with virtual positions at the selection outline
        var virtualOuterBounds = selectionBounds || null;
        if (!virtualOuterBounds) {
            virtualOuterBounds = computeUnionGeometricBounds(ruleItems);
            virtualOuterBounds = mergeGeometricBounds(virtualOuterBounds, computeUnionGeometricBounds(rectangleItems));
        }
        augmentRuleCentersWithBounds(ruleCenters, virtualOuterBounds, 0.5);

        var snappedCount = 0;
        for (var rectangleIndex = 0; rectangleIndex < rectangleItems.length; rectangleIndex++) {
            /* 読めない・変形できない背景は飛ばす / skip backgrounds that cannot be read or reshaped */
            try {
                var rectangleItem = rectangleItems[rectangleIndex];
                var itemBounds = rectangleItem.geometricBounds;
                var snappedLeft = findNearestValue(itemBounds[0], ruleCenters.xPositions);
                var snappedTop = findNearestValue(itemBounds[1], ruleCenters.yPositions);
                var snappedRight = findNearestValue(itemBounds[2], ruleCenters.xPositions);
                var snappedBottom = findNearestValue(itemBounds[3], ruleCenters.yPositions);
                if (snappedLeft === null || snappedRight === null || snappedTop === null || snappedBottom === null) continue;
                if (snappedLeft >= snappedRight || snappedTop <= snappedBottom) continue;
                if (fitClosedPathToBounds(rectangleItem, snappedLeft, snappedTop, snappedRight, snappedBottom)) {
                    snappedCount++;
                }
            } catch (e) { }
        }
        return snappedCount;
    }

    /**
     * 任意のアンカー数のクローズパスを、目標の範囲へ比例して写す
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} left - 範囲の左端
     * @param {number} top - 範囲の上端
     * @param {number} right - 範囲の右端
     * @param {number} bottom - 範囲の下端
     * @returns {boolean} 変形できたとき true
     */
    function fitClosedPathToBounds(pathItem, left, top, right, bottom) {
        if (!pathItem || !pathItem.pathPoints || pathItem.pathPoints.length < 2) return false;

        var sourceBounds = pathItem.geometricBounds;
        var sourceLeft = sourceBounds[0];
        var sourceTop = sourceBounds[1];
        var sourceWidth = sourceBounds[2] - sourceLeft;
        var sourceHeight = sourceTop - sourceBounds[3];
        var targetWidth = right - left;
        var targetHeight = top - bottom;

        if (sourceWidth <= 0.0001 || sourceHeight <= 0.0001) return false;
        if (targetWidth <= 0.0001 || targetHeight <= 0.0001) return false;

        var scaleX = targetWidth / sourceWidth;
        var scaleY = targetHeight / sourceHeight;

        /**
         * 元の範囲の座標を、目標の範囲の座標に写す
         * @param {number[]} pointArray - [x, y]
         * @returns {number[]} 写した [x, y]
         */
        function mapPoint(pointArray) {
            return [
                left + (pointArray[0] - sourceLeft) * scaleX,
                top - (sourceTop - pointArray[1]) * scaleY
            ];
        }

        /* アンカーの書き込みに失敗したら失敗扱い / treat a failed anchor write as failure */
        try {
            var pathPointList = pathItem.pathPoints;
            for (var i = 0; i < pathPointList.length; i++) {
                var pathPoint = pathPointList[i];
                pathPoint.anchor = mapPoint(pathPoint.anchor);
                pathPoint.leftDirection = mapPoint(pathPoint.leftDirection);
                pathPoint.rightDirection = mapPoint(pathPoint.rightDirection);
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    /**
     * 指定レイヤー直下のグループを解除する（ロック中・非表示のレイヤーは何もしない）
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function ungroupAllInLayer(doc, layerName) {
        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var docLayer = doc.layers[layerIndex];
            if (docLayer.name !== layerName) continue;
            if (docLayer.locked || !docLayer.visible) return;
            ungroupGroupItems(docLayer.pageItems);
            break;
        }
    }

    /**
     * 並びの中のグループを後ろから解除する（入れ子のグループは解除しない）
     * @param {PageItem[]} pageItems - 対象の並び
     * @returns {void}
     */
    function ungroupGroupItems(pageItems) {
        for (var i = pageItems.length - 1; i >= 0; i--) {
            var pageItem = pageItems[i];
            /* 解除できないグループは残す / leave groups that cannot be released */
            try {
                if (pageItem.typename === "GroupItem") ungroup(pageItem);
            } catch (e) { }
        }
    }

    /**
     * セル背景としてスナップできるパスを再帰的に集める
     * @param {PageItem[]} pageItems - 走査するアイテム
     * @param {PathItem[]} collectedPathItems - 見つけたパスを追加する配列
     * @returns {void}
     */
    function collectCellBackgroundSnapTargetsInItems(pageItems, collectedPathItems) {
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            var pageItem = pageItems[itemIndex];
            /* 属性を読めないものは飛ばす / skip items whose properties cannot be read */
            try {
                if (!pageItem || pageItem.locked || pageItem.hidden) continue;
                if (pageItem.typename === "PathItem") {
                    if (isCellBackgroundSnapTarget(pageItem)) collectedPathItems.push(pageItem);
                } else if (pageItem.typename === "GroupItem") {
                    collectCellBackgroundSnapTargetsInItems(pageItem.pageItems, collectedPathItems);
                } else if (pageItem.typename === "CompoundPathItem") {
                    for (var pathIndex = 0; pathIndex < pageItem.pathItems.length; pathIndex++) {
                        var childPathItem = pageItem.pathItems[pathIndex];
                        if (isCellBackgroundSnapTarget(childPathItem)) collectedPathItems.push(childPathItem);
                    }
                }
            } catch (e) { }
        }
    }

    /**
     * セル背景としてスナップできるパス（塗りのあるクローズパスで、マスクではない）かどうか
     * @param {PathItem} pathItem - 判定するパス
     * @returns {boolean} 対象なら true
     */
    function isCellBackgroundSnapTarget(pathItem) {
        if (!pathItem || pathItem.typename !== "PathItem") return false;
        if (!pathItem.closed || pathItem.clipping) return false;
        if (!pathItem.pathPoints || pathItem.pathPoints.length < 2) return false;
        return !!pathItem.filled;
    }

    /**
     * 罫線から中心線の X・Y 座標を集め、近い値をまとめて昇順にする
     * @param {PageItem[]} ruleItems - 罫線
     * @returns {{xPositions: number[], yPositions: number[]}} 縦罫の X と横罫の Y
     */
    function collectRuleCenterPositionsFromItems(ruleItems) {
        var xPositions = [];
        var yPositions = [];
        for (var itemIndex = 0; itemIndex < ruleItems.length; itemIndex++) {
            collectRuleCenterPositionsFromItem(ruleItems[itemIndex], xPositions, yPositions);
        }
        return {
            xPositions: uniqueSortedValues(xPositions, 0.5),
            yPositions: uniqueSortedValues(yPositions, 0.5)
        };
    }

    /**
     * 1つのアイテムから罫線の中心座標を集める（2点の線は中心、4点の長方形は4辺）
     * @param {PageItem} pageItem - 対象のアイテム
     * @param {number[]} xPositions - X 座標を追加する配列
     * @param {number[]} yPositions - Y 座標を追加する配列
     * @returns {void}
     */
    function collectRuleCenterPositionsFromItem(pageItem, xPositions, yPositions) {
        /* 読めないアイテムは飛ばす / skip unreadable items */
        try {
            if (!pageItem || pageItem.locked || pageItem.hidden) return;
            if (pageItem.typename === "PathItem") {
                if (!pageItem.closed) {
                    var lineInfo = getLineInfo(pageItem);
                    if (lineInfo) (lineInfo.orient === "h" ? yPositions : xPositions).push(lineInfo.coord);
                } else if (pageItem.pathPoints && pageItem.pathPoints.length === 4) {
                    var itemBounds = pageItem.geometricBounds;
                    xPositions.push(itemBounds[0]);
                    xPositions.push(itemBounds[2]);
                    yPositions.push(itemBounds[1]);
                    yPositions.push(itemBounds[3]);
                }
            } else if (pageItem.typename === "GroupItem") {
                for (var childIndex = 0; childIndex < pageItem.pageItems.length; childIndex++) {
                    collectRuleCenterPositionsFromItem(pageItem.pageItems[childIndex], xPositions, yPositions);
                }
            } else if (pageItem.typename === "CompoundPathItem") {
                for (var pathIndex = 0; pathIndex < pageItem.pathItems.length; pathIndex++) {
                    collectRuleCenterPositionsFromItem(pageItem.pathItems[pathIndex], xPositions, yPositions);
                }
            }
        } catch (e) { }
    }

    /**
     * 罫線の中心座標に、外形の四辺を仮想の罫線として足す（すでに近くに線がある辺は足さない）
     * @param {{xPositions: number[], yPositions: number[]}} ruleCenters - 罫線の中心座標（その場で書き換える）
     * @param {Object|null} outerBounds - 外形（left / top / right / bottom）
     * @param {number} tolerance - 近いとみなす差
     * @returns {void}
     */
    function augmentRuleCentersWithBounds(ruleCenters, outerBounds, tolerance) {
        if (!outerBounds) return;
        var edgeTolerance = (typeof tolerance === "number" && tolerance > 0) ? tolerance : 0.5;
        var xPositions = ruleCenters.xPositions;
        var yPositions = ruleCenters.yPositions;

        if (xPositions.length === 0 || outerBounds.left < xPositions[0] - edgeTolerance) xPositions.unshift(outerBounds.left);
        if (xPositions.length === 0 || outerBounds.right > xPositions[xPositions.length - 1] + edgeTolerance) xPositions.push(outerBounds.right);
        if (yPositions.length === 0 || outerBounds.bottom < yPositions[0] - edgeTolerance) yPositions.unshift(outerBounds.bottom);
        if (yPositions.length === 0 || outerBounds.top > yPositions[yPositions.length - 1] + edgeTolerance) yPositions.push(outerBounds.top);
    }

    /**
     * 指定の範囲に収まるアイテムだけを残す
     * @param {PageItem[]} items - 対象のアイテム
     * @param {{left: number, top: number, right: number, bottom: number}} limitBounds - 範囲
     * @param {number} tolerance - はみ出しの許容量
     * @returns {PageItem[]} 範囲に収まるアイテム
     */
    function filterItemsWithinBounds(items, limitBounds, tolerance) {
        if (!items || items.length === 0 || !limitBounds) return items || [];
        var edgeTolerance = (typeof tolerance === "number" && tolerance >= 0) ? tolerance : 0;
        var itemsWithin = [];
        for (var i = 0; i < items.length; i++) {
            /* 境界を読めないものは外す / drop items whose bounds cannot be read */
            try {
                var pageItem = items[i];
                if (!pageItem) continue;
                var itemBounds = pageItem.geometricBounds;
                if (!itemBounds) continue;
                if (itemBounds[0] >= limitBounds.left - edgeTolerance &&
                    itemBounds[2] <= limitBounds.right + edgeTolerance &&
                    itemBounds[1] <= limitBounds.top + edgeTolerance &&
                    itemBounds[3] >= limitBounds.bottom - edgeTolerance) {
                    itemsWithin.push(pageItem);
                }
            } catch (e) { }
        }
        return itemsWithin;
    }

    /**
     * 複数アイテムの geometricBounds の和集合を返す（ロック中・非表示は除く）
     * @param {PageItem[]} items - 対象のアイテム
     * @returns {{left: number, top: number, right: number, bottom: number}|null} 和集合。対象が無ければ null
     */
    function computeUnionGeometricBounds(items) {
        if (!items || items.length === 0) return null;
        var unionBounds = null;
        for (var i = 0; i < items.length; i++) {
            /* 境界を読めないものは飛ばす / skip items whose bounds cannot be read */
            try {
                var pageItem = items[i];
                if (!pageItem || pageItem.locked || pageItem.hidden) continue;
                var itemBounds = pageItem.geometricBounds;
                if (!itemBounds) continue;
                if (!unionBounds) {
                    unionBounds = { left: itemBounds[0], top: itemBounds[1], right: itemBounds[2], bottom: itemBounds[3] };
                } else {
                    if (itemBounds[0] < unionBounds.left) unionBounds.left = itemBounds[0];
                    if (itemBounds[1] > unionBounds.top) unionBounds.top = itemBounds[1];
                    if (itemBounds[2] > unionBounds.right) unionBounds.right = itemBounds[2];
                    if (itemBounds[3] < unionBounds.bottom) unionBounds.bottom = itemBounds[3];
                }
            } catch (e) { }
        }
        return unionBounds;
    }

    /**
     * 2つの範囲の和集合を返す（片方が null ならもう片方）
     * @param {Object|null} firstBounds - 範囲（left / top / right / bottom）
     * @param {Object|null} secondBounds - 範囲（left / top / right / bottom）
     * @returns {Object|null} 和集合
     */
    function mergeGeometricBounds(firstBounds, secondBounds) {
        if (!firstBounds) return secondBounds;
        if (!secondBounds) return firstBounds;
        return {
            left: Math.min(firstBounds.left, secondBounds.left),
            top: Math.max(firstBounds.top, secondBounds.top),
            right: Math.max(firstBounds.right, secondBounds.right),
            bottom: Math.min(firstBounds.bottom, secondBounds.bottom)
        };
    }

    /**
     * 近い値（tolerance 以内）をまとめて昇順にする
     * @param {number[]} values - 値
     * @param {number} tolerance - 同じとみなす差
     * @returns {number[]} まとめた昇順の値
     */
    function uniqueSortedValues(values, tolerance) {
        var sortedValues = values.slice();
        sortedValues.sort(function (valueA, valueB) { return valueA - valueB; });
        var mergedValues = [];
        for (var valueIndex = 0; valueIndex < sortedValues.length; valueIndex++) {
            var currentValue = sortedValues[valueIndex];
            if (mergedValues.length === 0 || Math.abs(currentValue - mergedValues[mergedValues.length - 1]) > tolerance) {
                mergedValues.push(currentValue);
            } else {
                mergedValues[mergedValues.length - 1] = (mergedValues[mergedValues.length - 1] + currentValue) / 2;
            }
        }
        return mergedValues;
    }

    /**
     * 最も近い値を返す
     * @param {number} targetValue - 基準の値
     * @param {number[]} values - 候補
     * @returns {number|null} 最も近い値。候補が無ければ null
     */
    function findNearestValue(targetValue, values) {
        if (!values || values.length === 0) return null;
        var nearestValue = values[0];
        var nearestDistance = Math.abs(targetValue - nearestValue);
        for (var valueIndex = 1; valueIndex < values.length; valueIndex++) {
            var distance = Math.abs(targetValue - values[valueIndex]);
            if (distance < nearestDistance) {
                nearestValue = values[valueIndex];
                nearestDistance = distance;
            }
        }
        return nearestValue;
    }

    // =========================================
    // 罫線の収集 / Collecting rules
    // =========================================

    /**
     * 中心線にする候補（4点のクローズパス）を再帰的に集める
     * @param {PageItem[]} pageItems - 走査するアイテム
     * @param {PathItem[]} collectedPathItems - 見つけたパスを追加する配列
     * @returns {void}
     */
    function collectRectsForCenterline(pageItems, collectedPathItems) {
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            var pageItem = pageItems[itemIndex];
            /* 読めないものは飛ばす / skip unreadable items */
            try {
                if (pageItem.locked || pageItem.hidden) continue;
                if (pageItem.typename === "PathItem") {
                    if (pageItem.closed && pageItem.pathPoints.length === 4) collectedPathItems.push(pageItem);
                } else if (pageItem.typename === "GroupItem") {
                    collectRectsForCenterline(pageItem.pageItems, collectedPathItems);
                } else if (pageItem.typename === "CompoundPathItem") {
                    if (pageItem.pathItems.length === 1) {
                        var childPathItem = pageItem.pathItems[0];
                        if (childPathItem.closed && childPathItem.pathPoints.length === 4) collectedPathItems.push(childPathItem);
                    }
                }
            } catch (e) { }
        }
    }

    /**
     * 対象レイヤーから罫線（開いた2点のパス）を集める
     * @param {Document} doc - 対象のドキュメント
     * @param {string[]} excludeLayerNames - 除外するレイヤー名
     * @returns {PathItem[]} 罫線
     */
    function collectRuleLines(doc, excludeLayerNames) {
        var ruleLines = [];
        var targetLayers = getTargetLayers(doc, excludeLayerNames);
        for (var layerIndex = 0; layerIndex < targetLayers.length; layerIndex++) {
            collectRuleLinesInItems(targetLayers[layerIndex].pageItems, ruleLines);
        }
        return ruleLines;
    }

    /**
     * 罫線（開いた2点のパス）を再帰的に集める
     * @param {PageItem[]} pageItems - 走査するアイテム
     * @param {PathItem[]} collectedLines - 見つけた線を追加する配列
     * @returns {void}
     */
    function collectRuleLinesInItems(pageItems, collectedLines) {
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            /* 読めないものは飛ばす / skip unreadable items */
            try {
                if (pageItem.locked || pageItem.hidden) continue;
                if (pageItem.typename === "PathItem") {
                    if (!pageItem.closed && pageItem.pathPoints && pageItem.pathPoints.length === 2) {
                        collectedLines.push(pageItem);
                    }
                } else if (pageItem.typename === "GroupItem") {
                    collectRuleLinesInItems(pageItem.pageItems, collectedLines);
                } else if (pageItem.typename === "CompoundPathItem") {
                    for (var childIndex = 0; childIndex < pageItem.pathItems.length; childIndex++) {
                        var childPathItem = pageItem.pathItems[childIndex];
                        if (!childPathItem.closed && childPathItem.pathPoints && childPathItem.pathPoints.length === 2) {
                            collectedLines.push(childPathItem);
                        }
                    }
                }
            } catch (e) { }
        }
    }

    /**
     * 細長い長方形を、中心を通る線（線幅＝短辺）に置き換える（10度以内の傾きは先に補正）
     * @param {PathItem} rectPath - 4点のクローズパス
     * @returns {PathItem|null} 作成した中心線。細長くなければ null（元の長方形は残す）
     */
    function convertRectToCenterLine(rectPath) {
        var firstAnchor = rectPath.pathPoints[0].anchor;
        var secondAnchor = rectPath.pathPoints[1].anchor;
        var angleDeg = Math.atan2(secondAnchor[1] - firstAnchor[1], secondAnchor[0] - firstAnchor[0]) * 180 / Math.PI;
        if (angleDeg < 0) angleDeg += 360;
        var deviationDeg = angleDeg - Math.round(angleDeg / 90) * 90;
        var absDeviationDeg = Math.abs(deviationDeg);
        if (absDeviationDeg >= 0.5 && absDeviationDeg <= 10) {
            safeCall(function () { rectPath.rotate(-deviationDeg); });
        }

        var rectBounds = rectPath.geometricBounds;
        var left = rectBounds[0], top = rectBounds[1], right = rectBounds[2], bottom = rectBounds[3];
        var rectWidth = right - left;
        var rectHeight = top - bottom;

        var diffRatio = Math.abs(rectWidth - rectHeight) / Math.max(rectWidth, rectHeight);
        if (diffRatio < 0.05) return null;
        var shortSide = Math.min(rectWidth, rectHeight);
        var longSide = Math.max(rectWidth, rectHeight);
        if (shortSide * 1.5 > longSide) return null;

        var parentLayer = rectPath.layer;
        if (!parentLayer) return null;

        var centerLine = parentLayer.pathItems.add();
        centerLine.stroked = true;
        centerLine.filled = false;
        safeCall(function () { centerLine.strokeColor = rectPath.fillColor; });

        var startPoint = centerLine.pathPoints.add();
        var endPoint = centerLine.pathPoints.add();

        if (rectHeight <= rectWidth) {
            var centerY = (top + bottom) / 2;
            startPoint.anchor = [left, centerY];
            endPoint.anchor = [right, centerY];
        } else {
            var centerX = (left + right) / 2;
            startPoint.anchor = [centerX, top];
            endPoint.anchor = [centerX, bottom];
        }

        centerLine.strokeWidth = (rectHeight <= rectWidth) ? rectHeight : rectWidth;

        startPoint.leftDirection = startPoint.anchor;
        startPoint.rightDirection = startPoint.anchor;
        endPoint.leftDirection = endPoint.anchor;
        endPoint.rightDirection = endPoint.anchor;

        safeCall(function () { rectPath.remove(); });
        return centerLine;
    }

    // =========================================
    // セル背景の抽出と集約 / Extracting and merging cell backgrounds
    // =========================================

    /**
     * 最も多く使われている文字サイズを返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {number} 最頻の文字サイズ（テキストが無ければ 0）
     */
    function getModeTextSize(doc) {
        var sizeCounts = {};
        for (var textFrameIndex = 0; textFrameIndex < doc.textFrames.length; textFrameIndex++) {
            /* 読めないテキストは飛ばす / skip unreadable text */
            try {
                var textCharacters = doc.textFrames[textFrameIndex].characters;
                for (var charIndex = 0; charIndex < textCharacters.length; charIndex++) {
                    var sizeKey = Math.round(textCharacters[charIndex].size * 10) / 10;
                    sizeCounts[sizeKey] = (sizeCounts[sizeKey] || 0) + 1;
                }
            } catch (e) { }
        }
        var modeSize = 0, modeCount = 0;
        for (var sizeText in sizeCounts) {
            if (sizeCounts.hasOwnProperty(sizeText) && sizeCounts[sizeText] > modeCount) {
                modeCount = sizeCounts[sizeText];
                modeSize = parseFloat(sizeText);
            }
        }
        return modeSize;
    }

    /**
     * 4点のクローズパス（長方形とみなすもの）かどうか
     * @param {PageItem} pageItem - 判定するアイテム
     * @returns {boolean} 長方形とみなせるとき true
     */
    function isRectangleLike(pageItem) {
        if (!pageItem || pageItem.typename !== "PathItem") return false;
        if (!pageItem.closed) return false;
        return pageItem.pathPoints.length === 4;
    }

    /**
     * 対象レイヤーから、短辺が shortSideMin より大きい長方形を集める
     * @param {Document} doc - 対象のドキュメント
     * @param {string[]} excludeLayerNames - 除外するレイヤー名
     * @param {number} shortSideMin - 短辺の下限（これより大きいものだけ）
     * @returns {PathItem[]} 長方形
     */
    function collectRectangles(doc, excludeLayerNames, shortSideMin) {
        var rectangles = [];
        var targetLayers = getTargetLayers(doc, excludeLayerNames);
        for (var layerIndex = 0; layerIndex < targetLayers.length; layerIndex++) {
            collectRectanglesInItems(targetLayers[layerIndex].pageItems, shortSideMin, rectangles);
        }
        return rectangles;
    }

    /**
     * 短辺が shortSideMin より大きい長方形を再帰的に集める
     * @param {PageItem[]} pageItems - 走査するアイテム
     * @param {number} shortSideMin - 短辺の下限
     * @param {PathItem[]} collectedRectangles - 見つけた長方形を追加する配列
     * @returns {void}
     */
    function collectRectanglesInItems(pageItems, shortSideMin, collectedRectangles) {
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            var pageItem = pageItems[itemIndex];
            /* 読めないものは飛ばす / skip unreadable items */
            try {
                if (pageItem.typename === "GroupItem") {
                    collectRectanglesInItems(pageItem.pageItems, shortSideMin, collectedRectangles);
                } else if (isRectangleLike(pageItem)) {
                    var itemWidth = pageItem.width;
                    var itemHeight = pageItem.height;
                    var shortSide = (itemWidth < itemHeight) ? itemWidth : itemHeight;
                    if (shortSide > shortSideMin) collectedRectangles.push(pageItem);
                }
            } catch (e) { }
        }
    }

    /**
     * レイヤーを用意して最背面に送る
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 用意したレイヤー
     */
    function ensureBackLayer(doc, layerName) {
        var backLayer = findOrCreateLayer(doc, layerName);
        safeCall(function () { backLayer.zOrder(ZOrderMethod.SENDTOBACK); });
        return backLayer;
    }

    /**
     * 各アイテムと同じアピアランスのオブジェクトを「共通 > アピアランス」で選択し、まとめて返す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} items - 基準のアイテム
     * @returns {PageItem[]} 見つかったアイテム（重複なし）
     */
    function expandByAppearance(doc, items) {
        var foundItems = [];
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            /* メニューコマンドが使えない状態なら飛ばす / skip when the menu command is unavailable */
            try {
                doc.selection = null;
                items[itemIndex].selected = true;
                app.executeMenuCommand('Find Appearance menu item');
                for (var j = 0; j < doc.selection.length; j++) {
                    if (!containsItem(foundItems, doc.selection[j])) {
                        foundItems.push(doc.selection[j]);
                    }
                }
            } catch (e) { }
        }
        return foundItems;
    }

    /**
     * アイテムをレイヤーの先頭へ移す（後ろから順に。移せないものは飛ばす）
     * @param {PageItem[]} items - 移すアイテム
     * @param {Layer} targetLayer - 移動先のレイヤー
     * @returns {void}
     */
    function moveItemsToLayer(items, targetLayer) {
        for (var itemIndex = items.length - 1; itemIndex >= 0; itemIndex--) {
            var movingItem = items[itemIndex];
            safeCall(function () { movingItem.move(targetLayer, ElementPlacement.PLACEATBEGINNING); });
        }
    }

    /**
     * 選択を解除してから、指定のアイテムを選択する
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} items - 選択するアイテム
     * @returns {void}
     */
    function setSelection(doc, items) {
        doc.selection = null;
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var selectingItem = items[itemIndex];
            safeCall(function () { selectingItem.selected = true; });
        }
    }

    /*
     目的:
     セル背景（矩形）は分割・重複・見た目効果（アピアランス）で散在している。
     対象矩形を抽出→同一レイヤーへ集約→グループ化することで、
     後続処理（罫線スナップなど）が扱いやすい状態に正規化する。

     Purpose:
     Cell backgrounds are fragmented and affected by appearances.
     This collects target rectangles, moves them to a single layer, and groups
     them so downstream steps (rule snapping, etc.) can work with them cleanly.
    */
    /**
     * セル背景（短辺が最頻の文字サイズより大きい長方形）を集め、同じアピアランスのものも含めて1つのレイヤーにまとめてグループ化する
     * @param {string} layerName - まとめる先のレイヤー名
     * @param {string[]} extraExcludeLayerNames - ほかに除外するレイヤー名
     * @returns {void}
     */
    function autoSelectAndMerge(layerName, extraExcludeLayerNames) {
        var doc = app.activeDocument;

        var minShortSide = getModeTextSize(doc);
        if (!minShortSide || minShortSide <= 0) return;

        var rectangles = collectRectangles(doc, [layerName].concat(extraExcludeLayerNames), minShortSide);
        if (rectangles.length === 0) return;

        setSelection(doc, rectangles);

        var cellBackgroundItems = expandByAppearance(doc, rectangles);
        if (cellBackgroundItems.length === 0) return;

        var targetLayer = ensureBackLayer(doc, layerName);
        moveItemsToLayer(cellBackgroundItems, targetLayer);
        setSelection(doc, cellBackgroundItems);

        safeCall(function () { app.executeMenuCommand('group'); });
    }

    main();
})();
