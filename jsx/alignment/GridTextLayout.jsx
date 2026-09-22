#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストをグリッド状に整列して配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridTextLayout.md

### Overview

Arranges text objects into a grid layout.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridTextLayout.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GridTextLayout";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridTextLayout.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridTextLayout.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 同じ行とみなす上端Yの許容差（初期値）/ initial tolerance for grouping tops into a row */
    var DEFAULT_ROW_TOLERANCE = 6;
    /* 同じ列とみなす左端Xの許容差（初期値）/ initial tolerance for grouping lefts into a column */
    var DEFAULT_COLUMN_TOLERANCE = 6;

    /* 行間・列間の初期値 / initial gutters */
    var DEFAULT_GAP = 1;
    /* ［連動］の初期状態 / initial state of the Link checkbox */
    var DEFAULT_LINK_GAPS = true;
    /* ［背景の長方形を削除］の初期状態 / initial state of the delete-rectangles checkbox */
    var DEFAULT_DELETE_RECTANGLES = false;

    /* 背景の長方形を置くレイヤー名 / layer that holds the background rectangles */
    var BACKGROUND_LAYER_NAME = "_bg-rectangle";
    /* 背景の長方形のグレー濃度（K%）/ gray level of the background rectangles */
    var BACKGROUND_GRAY_VALUE = 15;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X = 300;                /* ダイアログの表示位置：右(+)／左(-) */
    var DIALOG_OFFSET_Y = 0;                  /* ダイアログの表示位置：下(+)／上(-) */
    var DIALOG_OPACITY = 0.97;                /* ダイアログの不透明度 0.0 - 1.0 */
    var PANEL_MARGINS = [15, 20, 15, 10];     /* パネル余白 [左,上,右,下] / panel margins */
    var GUTTER_PANEL_MARGINS = [15, 20, 15, 5]; /* ［ガター］だけ下を詰める / the Gutter panel has a tighter bottom */
    var SIZE_LABEL_WIDTH = 30;                /* 幅・高さラベルの共通幅 / common width of the size labels */
    var TOLERANCE_INPUT_CHARS = 3;
    var GAP_INPUT_CHARS = 3;
    var SIZE_INPUT_CHARS = 5;
    var LINK_SPACER_HEIGHT = 10;              /* ［連動］の上下に入れる余白 / padding above and below the Link checkbox */

    /**
     * 向き・子の揃え・余白を指定してパネルを追加する
     * @param {Window} parentWindow - 追加先
     * @param {string} panelTitle - パネル名
     * @param {string} orientation - "row" または "column"
     * @param {string|string[]} alignChildren - 子の揃え
     * @param {number[]} margins - 余白 [左,上,右,下]
     * @returns {Panel} 追加したパネル
     */
    function addDialogPanel(parentWindow, panelTitle, orientation, alignChildren, margins) {
        var dialogPanel = parentWindow.add("panel", undefined, panelTitle);
        dialogPanel.orientation = orientation;
        dialogPanel.alignChildren = alignChildren;
        dialogPanel.margins = margins;
        return dialogPanel;
    }

    /**
     * 項目名と数値入力欄を追加する
     * @param {Group|Panel} parentGroup - 追加先
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {string} initialText - 入力欄の初期値
     * @param {number} inputChars - 入力欄の文字数
     * @param {string} tooltipPath - 入力欄の tooltip のラベルのパス
     * @param {number} [labelWidth] - 指定すると項目名をこの幅で右揃えにする
     * @returns {EditText} 追加した入力欄
     */
    function addNumberField(parentGroup, labelPath, initialText, inputChars, tooltipPath, labelWidth) {
        var fieldLabel = parentGroup.add("statictext", undefined, labelText(labelPath));
        if (labelWidth) {
            fieldLabel.justify = "right";
            fieldLabel.minimumSize.width = labelWidth;
        }
        var numberInput = parentGroup.add("edittext", undefined, initialText);
        numberInput.characters = inputChars;
        numberInput.helpTip = getLabel(tooltipPath);
        return numberInput;
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
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グリッド化", en: "Grid" }
        },
        panel: {
            detection: { ja: "行・列の判定", en: "Row & Column Detection" },
            gutter:    { ja: "ガター", en: "Gutter" },
            region:    { ja: "全体サイズ", en: "Overall Size" },
            option:    { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            rowTolerance:    { ja: "行", en: "Row" },
            columnTolerance: { ja: "列", en: "Column" },
            rowGap:          { ja: "行間", en: "Row Gap" },
            columnGap:       { ja: "列間", en: "Column Gap" },
            width:           { ja: "幅", en: "Width" },
            height:          { ja: "高さ", en: "Height" }
        },
        checkbox: {
            linkGaps:        { ja: "連動", en: "Link" },
            deleteRectangles:{ ja: "背景の長方形を削除", en: "Delete background rectangles" },
            convertToArea:   { ja: "エリア内文字に変換", en: "Convert to area text" }
        },
        tooltip: {
            rowTolerance: {
                ja: "上端Yがこの値以内にそろっているテキストを、同じ行とみなします。",
                en: "Text objects whose tops fall within this distance are treated as one row."
            },
            columnTolerance: {
                ja: "左端Xがこの値以内にそろっているテキストを、同じ列とみなします。",
                en: "Text objects whose left edges fall within this distance are treated as one column."
            },
            rowGap:    { ja: "行と行のあいだにあける間隔です。", en: "Space left between rows." },
            columnGap: { ja: "列と列のあいだにあける間隔です。", en: "Space left between columns." },
            linkGaps:  { ja: "行間と同じ値を列間にも使います。オフにすると列間を個別に指定できます。", en: "Uses the row gap for the column gap too. Turn it off to set the column gap separately." },
            width:     { ja: "グリッド全体の幅です。既定値は選択範囲の幅です。", en: "Total width of the grid. Defaults to the width of the selection." },
            height:    { ja: "グリッド全体の高さです。既定値は選択範囲の高さです。", en: "Total height of the grid. Defaults to the height of the selection." },
            deleteRectangles: {
                ja: "OK のときに、下敷きにした長方形とそのレイヤーを削除します。テキストの配置はそのまま残ります。",
                en: "Removes the backing rectangles and their layer on OK. The text placement is kept."
            },
            convertToArea: { ja: "現在このオプションは使用できません。", en: "This option is currently unavailable." }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:  { ja: "テキストを選択してください。", en: "Select some text." },
            noTextFrames: { ja: "テキストが選択されていません。", en: "No text object is selected." }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.gutter" のようなドット区切りキー
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
     * @param {Object|string} labelSet - ラベル、またはラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 共通処理 / Shared helpers
    // =========================================

    /**
     * K指定のグレースケールカラーを作る
     * @param {number} grayValue - K値（0〜100）
     * @returns {GrayColor} 生成した色
     */
    function makeGrayColor(grayValue) {
        var grayColor = new GrayColor();
        grayColor.gray = grayValue;
        return grayColor;
    }

    /**
     * 名前のレイヤーを取得し、無ければ最背面に作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 取得または作成したレイヤー
     */
    function getOrCreateBackgroundLayer(doc, layerName) {
        /* getByName は見つからないと例外を投げるため、作成のきっかけとして受ける
           getByName throws when the layer is missing, which is the signal to create it */
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            var createdLayer = doc.layers.add();
            createdLayer.name = layerName;
            createdLayer.zOrder(ZOrderMethod.SENDTOBACK);
            return createdLayer;
        }
    }

    /**
     * 値の近いものがすでに記録されているか判定する
     * @param {number[]} recordedValues - 記録済みの座標
     * @param {number} value - 判定する座標
     * @param {number} tolerance - 同じとみなす許容差
     * @returns {boolean} 近い値がすでにあれば true
     */
    function hasNearbyValue(recordedValues, value, tolerance) {
        for (var i = 0; i < recordedValues.length; i++) {
            if (Math.abs(value - recordedValues[i]) < tolerance) return true;
        }
        return false;
    }

    /**
     * 選択テキストから行数・列数と全体の外接範囲を求める
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @param {number} rowTolerance - 同じ行とみなす許容差
     * @param {number} columnTolerance - 同じ列とみなす許容差
     * @returns {{unionBounds: number[], rowCount: number, columnCount: number}} 判定結果
     */
    function detectGridShape(textFrames, rowTolerance, columnTolerance) {
        var unionBounds = [Number.MAX_VALUE, -Number.MAX_VALUE, -Number.MAX_VALUE, Number.MAX_VALUE];
        var rowTops = [];
        var columnLefts = [];

        for (var i = 0; i < textFrames.length; i++) {
            var itemBounds = textFrames[i].visibleBounds;
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];

            if (!hasNearbyValue(rowTops, itemBounds[1], rowTolerance)) rowTops.push(itemBounds[1]);
            if (!hasNearbyValue(columnLefts, itemBounds[0], columnTolerance)) columnLefts.push(itemBounds[0]);
        }

        return { unionBounds: unionBounds, rowCount: rowTops.length, columnCount: columnLefts.length };
    }

    /**
     * セルの左上座標を返す
     * @param {number[]} unionBounds - 全体の外接範囲 [左, 上, 右, 下]
     * @param {object} cellMetrics - cellWidth / cellHeight / rowGap / columnGap を持つ寸法
     * @param {number} rowIndex - 行番号
     * @param {number} columnIndex - 列番号
     * @returns {{left: number, top: number}} セルの左上
     */
    function getCellOrigin(unionBounds, cellMetrics, rowIndex, columnIndex) {
        return {
            left: unionBounds[0] + columnIndex * (cellMetrics.cellWidth + cellMetrics.columnGap),
            top: unionBounds[1] - rowIndex * (cellMetrics.cellHeight + cellMetrics.rowGap)
        };
    }

    /**
     * テキストの中心が入っているセルを探す
     * @param {number[]} frameBounds - テキストの境界 [左, 上, 右, 下]
     * @param {number[]} unionBounds - 全体の外接範囲
     * @param {object} gridShape - rowCount / columnCount を持つ判定結果
     * @param {object} cellMetrics - cellWidth / cellHeight / rowGap / columnGap を持つ寸法
     * @returns {{left: number, top: number}|null} セルの左上。見つからなければ null
     */
    function findContainingCell(frameBounds, unionBounds, gridShape, cellMetrics) {
        var centerX = (frameBounds[0] + frameBounds[2]) / 2;
        var centerY = (frameBounds[1] + frameBounds[3]) / 2;

        for (var rowIndex = 0; rowIndex < gridShape.rowCount; rowIndex++) {
            for (var columnIndex = 0; columnIndex < gridShape.columnCount; columnIndex++) {
                var cellOrigin = getCellOrigin(unionBounds, cellMetrics, rowIndex, columnIndex);
                if (centerX >= cellOrigin.left && centerX <= cellOrigin.left + cellMetrics.cellWidth &&
                    centerY <= cellOrigin.top && centerY >= cellOrigin.top - cellMetrics.cellHeight) {
                    return cellOrigin;
                }
            }
        }
        return null;
    }

    /**
     * 各テキストを、その中心が入っているセルの中央へ移動する
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @param {number[]} unionBounds - 全体の外接範囲
     * @param {object} gridShape - rowCount / columnCount を持つ判定結果
     * @param {object} cellMetrics - cellWidth / cellHeight / rowGap / columnGap を持つ寸法
     * @returns {void}
     */
    function centerTextInCells(textFrames, unionBounds, gridShape, cellMetrics) {
        for (var i = 0; i < textFrames.length; i++) {
            var frameBounds = textFrames[i].visibleBounds;
            var cellOrigin = findContainingCell(frameBounds, unionBounds, gridShape, cellMetrics);
            if (!cellOrigin) continue;

            var frameWidth = frameBounds[2] - frameBounds[0];
            var frameHeight = frameBounds[1] - frameBounds[3];
            var targetLeft = cellOrigin.left + (cellMetrics.cellWidth - frameWidth) / 2;
            var targetTop = cellOrigin.top - (cellMetrics.cellHeight - frameHeight) / 2;

            textFrames[i].translate(targetLeft - frameBounds[0], targetTop - frameBounds[1]);
        }
    }

    // =========================================
    // 入力欄 / Numeric inputs
    // =========================================

    /**
     * ↑↓キーで数値欄の値を増減する（Shiftで10刻み、Optionで0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {function} onValueChanged - 値が変わったときに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onValueChanged) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === "Up");
            var newValue;

            if (keyboardState.shiftKey) {
                /* Shift：10の倍数へ丸めながら増減 / snap to multiples of 10 */
                newValue = isUp ? Math.ceil((currentValue + 1) / 10) * 10 : Math.floor((currentValue - 1) / 10) * 10;
            } else if (keyboardState.altKey) {
                newValue = Math.round((currentValue + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                newValue = Math.round(currentValue) + (isUp ? 1 : -1);
            }

            if (newValue < 0) newValue = 0;
            editText.text = newValue;
            event.preventDefault();

            if (onValueChanged) onValueChanged();
        });
    }

    /**
     * 数値入力欄に、連動先への反映とプレビュー更新、↑↓キー操作をまとめて結び付ける
     * @param {EditText} editText - 対象の入力欄
     * @param {object} [inputOptions] - linkedInput / linkCheckbox / onValueChanged
     * @returns {void}
     */
    function bindNumericInput(editText, inputOptions) {
        if (!inputOptions) inputOptions = {};

        /**
         * 連動が有効なら相手の入力欄にも同じ値を入れ、プレビューを更新する
         * @returns {void}
         */
        function syncAndPreview() {
            if (inputOptions.linkCheckbox && inputOptions.linkCheckbox.value && inputOptions.linkedInput) {
                inputOptions.linkedInput.text = editText.text;
            }
            if (inputOptions.onValueChanged) inputOptions.onValueChanged();
        }

        editText.onChanging = syncAndPreview;
        changeValueByArrowKey(editText, syncAndPreview);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択テキストの行・列を判定し、グリッドのプレビューを出すダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedObjects = doc.selection;
        if (!selectedObjects || selectedObjects.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var textFrames = [];
        for (var i = 0; i < selectedObjects.length; i++) {
            if (selectedObjects[i].typename === "TextFrame") textFrames.push(selectedObjects[i]);
        }
        if (textFrames.length === 0) {
            alert(getLabel("alert.noTextFrames"));
            return;
        }

        showGridDialog(doc, textFrames);
    }

    /**
     * グリッド化のダイアログを組み立てる（イベント結線は呼び出し側で行う）
     * @param {object} gridShape - unionBounds / rowCount / columnCount を持つ判定結果（幅・高さの初期値に使う）
     * @returns {object} ダイアログと各コントロールの参照
     */
    function buildGridDialog(gridShape) {
        var currentUnitLabel = getUnitInfo().label;

        var gridDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        gridDialog.orientation = "column";
        gridDialog.alignChildren = "left";
        gridDialog.opacity = DIALOG_OPACITY;

        /* 判定 / Detection */
        var detectionPanel = addDialogPanel(gridDialog, getLabel("panel.detection"), "row", ["left", "center"], PANEL_MARGINS);
        var rowToleranceInput = addNumberField(detectionPanel, "fieldLabel.rowTolerance",
            String(DEFAULT_ROW_TOLERANCE), TOLERANCE_INPUT_CHARS, "tooltip.rowTolerance");
        var columnToleranceInput = addNumberField(detectionPanel, "fieldLabel.columnTolerance",
            String(DEFAULT_COLUMN_TOLERANCE), TOLERANCE_INPUT_CHARS, "tooltip.columnTolerance");

        /* ガター / Gutter */
        var gutterPanel = addDialogPanel(gridDialog, getLabel("panel.gutter") + "（" + currentUnitLabel + "）",
            "row", "top", GUTTER_PANEL_MARGINS);

        var gutterLeftColumn = gutterPanel.add("group");
        gutterLeftColumn.orientation = "column";
        gutterLeftColumn.alignChildren = "left";

        var rowGapInput = addNumberField(gutterLeftColumn.add("group"), "fieldLabel.rowGap",
            String(DEFAULT_GAP), GAP_INPUT_CHARS, "tooltip.rowGap");
        var columnGapInput = addNumberField(gutterLeftColumn.add("group"), "fieldLabel.columnGap",
            String(DEFAULT_GAP), GAP_INPUT_CHARS, "tooltip.columnGap");

        var gutterRightColumn = gutterPanel.add("group");
        gutterRightColumn.orientation = "column";
        gutterRightColumn.alignChildren = ["fill", "fill"];

        var linkSpacerTop = gutterRightColumn.add("statictext", undefined, "");
        linkSpacerTop.minimumSize.height = LINK_SPACER_HEIGHT;

        var linkGapsCheckbox = gutterRightColumn.add("checkbox", undefined, getLabel("checkbox.linkGaps"));
        linkGapsCheckbox.helpTip = getLabel("tooltip.linkGaps");
        linkGapsCheckbox.value = DEFAULT_LINK_GAPS;
        columnGapInput.enabled = !DEFAULT_LINK_GAPS;

        var linkSpacerBottom = gutterRightColumn.add("statictext", undefined, "");
        linkSpacerBottom.minimumSize.height = LINK_SPACER_HEIGHT;

        /* 全体サイズ / Overall size */
        var regionPanel = addDialogPanel(gridDialog, getLabel("panel.region") + "（" + currentUnitLabel + "）",
            "row", "top", PANEL_MARGINS);
        var widthInput = addNumberField(regionPanel.add("group"), "fieldLabel.width",
            String(Math.round(gridShape.unionBounds[2] - gridShape.unionBounds[0])), SIZE_INPUT_CHARS, "tooltip.width", SIZE_LABEL_WIDTH);
        var heightInput = addNumberField(regionPanel.add("group"), "fieldLabel.height",
            String(Math.round(gridShape.unionBounds[1] - gridShape.unionBounds[3])), SIZE_INPUT_CHARS, "tooltip.height", SIZE_LABEL_WIDTH);

        /* オプション / Options */
        var optionPanel = addDialogPanel(gridDialog, getLabel("panel.option"), "column", "left", PANEL_MARGINS);

        var deleteRectanglesCheckbox = optionPanel.add("checkbox", undefined, getLabel("checkbox.deleteRectangles"));
        deleteRectanglesCheckbox.helpTip = getLabel("tooltip.deleteRectangles");
        deleteRectanglesCheckbox.value = DEFAULT_DELETE_RECTANGLES;

        var convertToAreaCheckbox = optionPanel.add("checkbox", undefined, getLabel("checkbox.convertToArea"));
        convertToAreaCheckbox.helpTip = getLabel("tooltip.convertToArea");
        convertToAreaCheckbox.value = false;
        convertToAreaCheckbox.enabled = false;

        /* ボタンエリア / Button row */
        var btnRowGroup = gridDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            gridDialog: gridDialog,
            rowToleranceInput: rowToleranceInput,
            columnToleranceInput: columnToleranceInput,
            rowGapInput: rowGapInput,
            columnGapInput: columnGapInput,
            linkGapsCheckbox: linkGapsCheckbox,
            widthInput: widthInput,
            heightInput: heightInput,
            deleteRectanglesCheckbox: deleteRectanglesCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 入力値からセルの寸法を求める。間隔は数値でなければ 0、幅・高さは正の数でなければ選択範囲の寸法
     * @param {object} dialogControls - buildGridDialog() の戻り値
     * @param {object} gridShape - unionBounds / rowCount / columnCount を持つ判定結果
     * @returns {object} cellWidth / cellHeight / rowGap / columnGap を持つ寸法
     */
    function readCellMetrics(dialogControls, gridShape) {
        var rowGap = Number(dialogControls.rowGapInput.text);
        var columnGap = Number(dialogControls.columnGapInput.text);
        if (isNaN(rowGap)) rowGap = 0;
        if (isNaN(columnGap)) columnGap = 0;

        var totalWidth = Number(dialogControls.widthInput.text);
        var totalHeight = Number(dialogControls.heightInput.text);
        if (isNaN(totalWidth) || totalWidth <= 0) totalWidth = gridShape.unionBounds[2] - gridShape.unionBounds[0];
        if (isNaN(totalHeight) || totalHeight <= 0) totalHeight = gridShape.unionBounds[1] - gridShape.unionBounds[3];

        return {
            rowGap: rowGap,
            columnGap: columnGap,
            cellWidth: (totalWidth - (gridShape.columnCount - 1) * columnGap) / gridShape.columnCount,
            cellHeight: (totalHeight - (gridShape.rowCount - 1) * rowGap) / gridShape.rowCount
        };
    }

    /**
     * セルごとに下敷きの長方形（グレー・線なし）を最背面に描く
     * @param {Layer} backgroundLayer - 描き込むレイヤー
     * @param {object} gridShape - unionBounds / rowCount / columnCount を持つ判定結果
     * @param {object} cellMetrics - cellWidth / cellHeight / rowGap / columnGap を持つ寸法
     * @returns {PathItem[]} 描いた長方形
     */
    function drawCellRectangles(backgroundLayer, gridShape, cellMetrics) {
        var cellRectangles = [];
        for (var rowIndex = 0; rowIndex < gridShape.rowCount; rowIndex++) {
            for (var columnIndex = 0; columnIndex < gridShape.columnCount; columnIndex++) {
                var cellOrigin = getCellOrigin(gridShape.unionBounds, cellMetrics, rowIndex, columnIndex);
                var cellRectangle = backgroundLayer.pathItems.rectangle(
                    cellOrigin.top, cellOrigin.left, cellMetrics.cellWidth, cellMetrics.cellHeight);
                cellRectangle.stroked = false;
                cellRectangle.filled = true;
                cellRectangle.fillColor = makeGrayColor(BACKGROUND_GRAY_VALUE);
                cellRectangle.zOrder(ZOrderMethod.SENDTOBACK);
                cellRectangles.push(cellRectangle);
            }
        }
        return cellRectangles;
    }

    /**
     * グリッド化のダイアログを表示し、入力に合わせてプレビューを更新する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @returns {void}
     */
    function showGridDialog(doc, textFrames) {
        var backgroundLayer = getOrCreateBackgroundLayer(doc, BACKGROUND_LAYER_NAME);
        var rowTolerance = DEFAULT_ROW_TOLERANCE;
        var columnTolerance = DEFAULT_COLUMN_TOLERANCE;
        var gridShape = detectGridShape(textFrames, rowTolerance, columnTolerance);

        var dialogControls = buildGridDialog(gridShape);
        var gridDialog = dialogControls.gridDialog;
        var rowGapInput = dialogControls.rowGapInput;
        var columnGapInput = dialogControls.columnGapInput;
        var linkGapsCheckbox = dialogControls.linkGapsCheckbox;

        /* プレビュー用の長方形 / Rectangles drawn for the preview */
        var previewRectangles = [];

        /**
         * プレビューの長方形をすべて消す
         * @returns {void}
         */
        function clearPreview() {
            for (var i = 0; i < previewRectangles.length; i++) {
                /* すでに消えている長方形を踏んでも止めない / a rectangle may already be gone */
                try {
                    previewRectangles[i].remove();
                } catch (e) {}
            }
            previewRectangles = [];
        }

        /**
         * 下敷きの長方形を引き直し、テキストを各セルの中央へ置き直す
         * @returns {void}
         */
        function updatePreview() {
            clearPreview();
            var cellMetrics = readCellMetrics(dialogControls, gridShape);
            previewRectangles = drawCellRectangles(backgroundLayer, gridShape, cellMetrics);
            centerTextInCells(textFrames, gridShape.unionBounds, gridShape, cellMetrics);
            app.redraw();
        }

        /**
         * 判定の許容値を読み直し、行数・列数を求め直してプレビューを更新する
         * @returns {void}
         */
        function refreshGridShape() {
            var enteredRowTolerance = Number(dialogControls.rowToleranceInput.text);
            var enteredColumnTolerance = Number(dialogControls.columnToleranceInput.text);
            if (!isNaN(enteredRowTolerance) && enteredRowTolerance >= 0) rowTolerance = enteredRowTolerance;
            if (!isNaN(enteredColumnTolerance) && enteredColumnTolerance >= 0) columnTolerance = enteredColumnTolerance;

            gridShape = detectGridShape(textFrames, rowTolerance, columnTolerance);
            updatePreview();
        }

        bindNumericInput(dialogControls.rowToleranceInput, { onValueChanged: refreshGridShape });
        bindNumericInput(dialogControls.columnToleranceInput, { onValueChanged: refreshGridShape });
        bindNumericInput(rowGapInput, { linkedInput: columnGapInput, linkCheckbox: linkGapsCheckbox, onValueChanged: updatePreview });
        bindNumericInput(columnGapInput, { onValueChanged: updatePreview });
        bindNumericInput(dialogControls.widthInput, { onValueChanged: updatePreview });
        bindNumericInput(dialogControls.heightInput, { onValueChanged: updatePreview });

        linkGapsCheckbox.onClick = function () {
            columnGapInput.enabled = !linkGapsCheckbox.value;
            if (linkGapsCheckbox.value) {
                columnGapInput.text = rowGapInput.text;
                updatePreview();
            }
        };

        dialogControls.btnCancel.onClick = function () {
            clearPreview();
            gridDialog.close();
        };

        dialogControls.btnOK.onClick = function () {
            updatePreview();
            if (dialogControls.deleteRectanglesCheckbox.value) {
                clearPreview();
                removeBackgroundLayer(doc);
            }
            gridDialog.close(1);
        };

        /* 位置をずらすのとフォーカス設定は onShow 1つにまとめる（2回代入すると先の指定が消える）
           Both the offset and the initial focus go in one onShow; a second assignment would drop the first */
        gridDialog.onShow = function () {
            gridDialog.location = [
                gridDialog.location[0] + DIALOG_OFFSET_X,
                gridDialog.location[1] + DIALOG_OFFSET_Y
            ];
            rowGapInput.active = true;
        };

        /* 初期プレビュー / Initial preview */
        columnGapInput.text = rowGapInput.text;
        updatePreview();

        gridDialog.show();
    }

    /**
     * 背景の長方形用レイヤーを削除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removeBackgroundLayer(doc) {
        /* すでに無い場合があるため、見つからない例外は受け流す / the layer may already be gone */
        try {
            doc.layers.getByName(BACKGROUND_LAYER_NAME).remove();
        } catch (e) {}
    }

    main();

})();
