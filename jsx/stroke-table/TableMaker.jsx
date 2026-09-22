#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームの見た目（位置・幅・行数）に合わせて、罫線と背景を自動生成します。
テキスト（文字・タブ・スタイル・タブストップ）には一切手を加えません。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TableMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4eaa14098858

### Overview

Generates rules and backgrounds that match the appearance — position, width and line count — of the selected text frame.
The text itself, including tabs, styles and tab stops, is never touched.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TableMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TableMaker";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TableMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TableMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4eaa14098858"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var MAX_SELECTION_COUNT    = 1000;         /* これ以上の選択では処理しない / skip when this many items are selected */
    var DEFAULT_SHAPE_MODE     = "topBottom";  /* ［形状］の初期値 "outerRect" / "topBottom" / "rowFills" / initial shape */
    var DEFAULT_VERTICAL_RULES = false;        /* ［縦罫］の初期値 / initial state of Vertical rules */
    var DEFAULT_HEADING        = true;         /* ［見出し］の初期値 / initial state of Heading */
    var HEADING_STROKE_SCALE   = 3;            /* 見出しを強調する罫線の倍率 / stroke multiplier for the heading rules */
    var RULE_BLACK             = 100;          /* 罫線の濃さ（K%）/ rule color (K%) */
    var ROW_FILL_BLACK_HEADING = 50;           /* 見出し行の背景（K%）/ heading row fill (K%) */
    var ROW_FILL_BLACK_ODD     = 30;           /* 1・3・5…行目の背景（K%）/ fill of the 1st, 3rd, 5th… rows (K%) */
    var ROW_FILL_BLACK_EVEN    = 10;           /* 2・4・6…行目の背景（K%）/ fill of the 2nd, 4th, 6th… rows (K%) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OFFSET_X    = 300;               /* ダイアログを右へずらす量 / horizontal dialog offset */
    var DIALOG_OFFSET_Y    = 0;                 /* ダイアログを下へずらす量 / vertical dialog offset */
    var DIALOG_OPACITY     = 0.98;              /* ダイアログの不透明度 / dialog opacity */
    var PANEL_MARGINS      = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / panel margins */
    var STROKE_WIDTH_CHARS = 5;                 /* 線幅の入力欄の文字数 / characters for the stroke width field */

    /**
     * ダイアログの表示位置をずらす
     * @param {Window} dlg - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    /**
     * ダイアログの不透明度を設定する
     * @param {Window} dlg - 対象のダイアログ
     * @param {number} opacityValue - 不透明度（0〜1）
     * @returns {void}
     */
    function setDialogOpacity(dlg, opacityValue) {
        try {
            dlg.opacity = opacityValue;
        } catch (e) {
            /* 不透明度を持たない環境では既定のまま / keep the default where opacity is unsupported */
        }
    }

    /**
     * ↑↓キーで値を増減する（Shift/Option対応）
     * - ↑↓: ±1
     * - Shift+↑↓: ±10（10の倍数へスナップ）
     * - Option+↑↓: ±0.1
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;

                // Shiftキー押下時は10の倍数にスナップ / Snap to tens when Shift is held
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }

            } else if (keyboard.altKey) {
                delta = 0.1;

                if (event.keyName === "Up") {
                    value += delta;
                } else {
                    value -= delta;
                }

            } else {
                delta = 1;

                if (event.keyName === "Up") {
                    value += delta;
                } else {
                    value -= delta;
                }
            }

            // 下限を0に / Clamp to 0
            if (value < 0) value = 0;

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
            } else {
                value = Math.round(value); /* 整数に丸め / Round to integer */
            }

            editText.text = String(value);
            event.preventDefault();
        });
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
     * 小数第3位までに丸める
     * @param {number} value - 丸める値
     * @returns {number} 丸めた値
     */
    function roundToThousandths(value) {
        return Math.round(value * 1000) / 1000;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義。radio と tooltip の形状キーは SHAPE_MODES と同じ名前
       Categorized Japanese-English labels; shape keys in radio / tooltip match SHAPE_MODES */
    var LABELS = {
        dialog: {
            title: { ja: SCRIPT_NAME + " " + SCRIPT_VERSION, en: SCRIPT_NAME + " " + SCRIPT_VERSION }
        },
        panel: {
            shape: { ja: "形状", en: "Shape" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Stroke width" }
        },
        radio: {
            outerRect: { ja: "外枠は長方形", en: "Rectangle border" },
            topBottom: { ja: "外枠なし", en: "No outer border" },
            rowFills:  { ja: "行ごとに長方形", en: "Row rectangles" }
        },
        checkbox: {
            verticalRules: { ja: "縦罫", en: "Vertical rules" },
            heading:       { ja: "見出し", en: "Heading" }
        },
        tooltip: {
            strokeWidth:   { ja: "表の幅です。", en: "Width of the table." },
            outerRect:     { ja: "表全体を1つの長方形で囲みます。", en: "Frames the whole table with a single rectangle." },
            topBottom:     { ja: "上下のケイ線だけを引きます。", en: "Draws only the top and bottom rules." },
            rowFills:      { ja: "行ごとに長方形を作ります。", en: "Creates a rectangle for each row." },
            verticalRules: { ja: "列の境に縦のケイ線を引きます。", en: "Draws a vertical rule between the columns." },
            heading:       { ja: "1行目を見出し行として扱います。", en: "Treats the first row as a header." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        }
    };

    /**
     * 表示言語に応じたラベル文字列を返す
     * @param {Object} labelNode - LABELS 内の { ja, en } ノード
     * @returns {string} 表示用の文言
     */
    function getLabel(labelNode) {
        if (!labelNode) return "";
        return labelNode[uiLang] || labelNode.en || labelNode.ja || "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelNode - LABELS 内の { ja, en } ノード
     * @returns {string} コロンを添えたラベル
     */
    function labelText(labelNode) {
        return getLabel(labelNode) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 形状の選択肢（ラジオボタンの並び順）/ Shape modes in radio button order */
    var SHAPE_MODES = ["outerRect", "topBottom", "rowFills"];

    /**
     * 罫線の設定ダイアログを表示する
     * @param {number} defaultStrokeWidth - 線幅の初期値（定規の単位）
     * @param {{code: number, label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {{strokeWidthPt: number, shapeMode: string, verticalRules: boolean, heading: boolean}|null} キャンセル時・線幅が数値でないときは null
     */
    function showTableSettingsDialog(defaultStrokeWidth, rulerUnit) {
        var dlg = new Window("dialog", getLabel(LABELS.dialog.title));
        setDialogOpacity(dlg, DIALOG_OPACITY);
        shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        var strokeWidthControls = addStrokeWidthRow(dlg, defaultStrokeWidth, rulerUnit.label);
        var shapeRadios = addShapePanel(dlg);
        var optionChecks = addOptionRow(dlg);
        addButtonRow(dlg);

        /**
         * 行ごとに長方形のときは線を引かないので、線幅と縦罫をディム表示にする
         * @returns {void}
         */
        function updateStrokeControls() {
            var drawsRules = !shapeRadios.rowFills.value;
            strokeWidthControls.txtStrokeWidth.enabled = drawsRules;
            strokeWidthControls.lblStrokeUnit.enabled = drawsRules;
            optionChecks.cbVerticalRules.enabled = drawsRules;
            if (!drawsRules) optionChecks.cbVerticalRules.value = false;
        }

        for (var i = 0; i < SHAPE_MODES.length; i++) {
            shapeRadios[SHAPE_MODES[i]].onClick = updateStrokeControls;
        }
        updateStrokeControls();

        if (dlg.show() !== 1) return null;
        return readTableSettings(strokeWidthControls, shapeRadios, optionChecks, rulerUnit);
    }

    /**
     * 線幅の行（項目名・入力欄・単位）を追加する
     * @param {Window} dlg - 追加先のダイアログ
     * @param {number} defaultStrokeWidth - 線幅の初期値（定規の単位）
     * @param {string} unitLabel - 単位の表示名
     * @returns {{txtStrokeWidth: EditText, lblStrokeUnit: StaticText}} 線幅の入力欄と単位表示
     */
    function addStrokeWidthRow(dlg, defaultStrokeWidth, unitLabel) {
        var strokeWidthRow = dlg.add("group");
        strokeWidthRow.orientation = "row";
        strokeWidthRow.alignChildren = ["left", "center"];

        strokeWidthRow.add("statictext", undefined, labelText(LABELS.fieldLabel.strokeWidth));
        var txtStrokeWidth = strokeWidthRow.add("edittext", undefined, String(defaultStrokeWidth));
        txtStrokeWidth.helpTip = getLabel(LABELS.tooltip.strokeWidth);
        txtStrokeWidth.characters = STROKE_WIDTH_CHARS;
        changeValueByArrowKey(txtStrokeWidth);

        var lblStrokeUnit = strokeWidthRow.add("statictext", undefined, unitLabel);
        return { txtStrokeWidth: txtStrokeWidth, lblStrokeUnit: lblStrokeUnit };
    }

    /**
     * ［形状］パネルを追加する
     * @param {Window} dlg - 追加先のダイアログ
     * @returns {Object} 形状名をキーにしたラジオボタンの集まり
     */
    function addShapePanel(dlg) {
        var shapePanel = dlg.add("panel", undefined, getLabel(LABELS.panel.shape));
        shapePanel.orientation = "column";
        shapePanel.alignChildren = ["left", "top"];
        shapePanel.margins = PANEL_MARGINS;

        var shapeRadios = {};
        for (var i = 0; i < SHAPE_MODES.length; i++) {
            var shapeMode = SHAPE_MODES[i];
            var rbShape = shapePanel.add("radiobutton", undefined, getLabel(LABELS.radio[shapeMode]));
            rbShape.helpTip = getLabel(LABELS.tooltip[shapeMode]);
            shapeRadios[shapeMode] = rbShape;
        }
        /* 設定名が違うときは先頭の形状にする / fall back to the first shape for an unknown name */
        (shapeRadios[DEFAULT_SHAPE_MODE] || shapeRadios[SHAPE_MODES[0]]).value = true;
        return shapeRadios;
    }

    /**
     * ［縦罫］［見出し］のチェックボックス行を追加する
     * @param {Window} dlg - 追加先のダイアログ
     * @returns {{cbVerticalRules: Checkbox, cbHeading: Checkbox}} 追加したチェックボックス
     */
    function addOptionRow(dlg) {
        var optionRow = dlg.add("group");
        optionRow.orientation = "row";
        optionRow.alignment = "center";                /* ダイアログの左右中央に置く / center in the dialog */
        optionRow.alignChildren = ["center", "center"];

        var cbVerticalRules = optionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.verticalRules));
        cbVerticalRules.helpTip = getLabel(LABELS.tooltip.verticalRules);
        cbVerticalRules.value = DEFAULT_VERTICAL_RULES;

        var cbHeading = optionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.heading));
        cbHeading.helpTip = getLabel(LABELS.tooltip.heading);
        cbHeading.value = DEFAULT_HEADING;

        return { cbVerticalRules: cbVerticalRules, cbHeading: cbHeading };
    }

    /**
     * ［キャンセル］［OK］のボタン行を追加する
     * @param {Window} dlg - 追加先のダイアログ
     * @returns {void}
     */
    function addButtonRow(dlg) {
        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["right", "bottom"];
        btnRowGroup.alignChildren = ["right", "center"];

        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        btnOK.active = true; /* 既定のボタンは右側の OK / OK on the right is the default */
    }

    /**
     * ダイアログの入力内容を設定値にまとめる
     * @param {{txtStrokeWidth: EditText, lblStrokeUnit: StaticText}} strokeWidthControls - 線幅の入力欄と単位表示
     * @param {Object} shapeRadios - 形状名をキーにしたラジオボタンの集まり
     * @param {{cbVerticalRules: Checkbox, cbHeading: Checkbox}} optionChecks - オプションのチェックボックス
     * @param {{code: number, label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {{strokeWidthPt: number, shapeMode: string, verticalRules: boolean, heading: boolean}|null} 線幅が数値でないときは null
     */
    function readTableSettings(strokeWidthControls, shapeRadios, optionChecks, rulerUnit) {
        var strokeWidth = Number(strokeWidthControls.txtStrokeWidth.text);
        if (isNaN(strokeWidth)) return null;

        var selectedShapeMode = SHAPE_MODES[0];
        for (var i = 0; i < SHAPE_MODES.length; i++) {
            if (shapeRadios[SHAPE_MODES[i]].value) selectedShapeMode = SHAPE_MODES[i];
        }

        return {
            strokeWidthPt: roundToThousandths(strokeWidth * rulerUnit.pointsPerUnit),
            shapeMode: selectedShapeMode,
            verticalRules: optionChecks.cbVerticalRules.value,
            heading: optionChecks.cbHeading.value
        };
    }

    // =========================================
    // 作図 / Drawing
    // =========================================

    /**
     * テキストフレームの見た目に合わせて罫線・背景を作る（テキストには手を加えない）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Layer} targetLayer - 作図先のレイヤー
     * @param {{strokeWidthPt: number, shapeMode: string, verticalRules: boolean, heading: boolean}} tableSettings - ダイアログの設定値
     * @returns {void}
     */
    function createTableForTextFrame(textFrame, targetLayer, tableSettings) {
        textFrame.selected = false;

        var paragraphs = textFrame.paragraphs;
        if (paragraphs.length === 0) return;

        var tableBox = measureTableBox(textFrame);

        if (tableSettings.shapeMode === "rowFills") {
            /* 行ごとの背景だけを作り、罫線・縦罫は引かない / fills only, no rules */
            createRowFills(tableBox, targetLayer, tableSettings.heading);
            return;
        }

        var strokeWidthPt = tableSettings.strokeWidthPt;
        /* 見出しがあるときは、外枠（上下の罫線）と見出しの下の罫線を太くする / thicker outer rules and heading rule */
        var accentWidthPt = tableSettings.heading ? strokeWidthPt * HEADING_STROKE_SCALE : strokeWidthPt;

        if (tableSettings.shapeMode === "topBottom") {
            createTopBottomRules(tableBox, targetLayer, accentWidthPt);
        } else {
            createOuterRectangle(tableBox, targetLayer, accentWidthPt);
        }
        createRowSeparators(tableBox, targetLayer, strokeWidthPt, accentWidthPt);

        if (tableSettings.verticalRules) {
            createColumnRules(tableBox, targetLayer, getColumnTabPositions(paragraphs), strokeWidthPt);
        }
    }

    /**
     * テキストフレームから表の外形を求める
     * 1段落を1行とし、行送りと文字サイズの差の半分を上下左右の余白にする
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{left: number, right: number, top: number, bottom: number, width: number, height: number, rowHeight: number, rowCount: number}} 表の外形
     */
    function measureTableBox(textFrame) {
        var bounds = textFrame.geometricBounds; /* [左, 上, 右, 下] / [left, top, right, bottom] */
        var paragraphs = textFrame.paragraphs;
        var firstAttributes = paragraphs[0].characterAttributes;

        var rowHeight = firstAttributes.leading;
        var cellPadding = (rowHeight - firstAttributes.size) / 2;
        var left = bounds[0] - cellPadding;
        var right = bounds[2] + cellPadding;
        var top = bounds[1] + cellPadding;
        var height = rowHeight * paragraphs.length;

        return {
            left: left,
            right: right,
            top: top,
            bottom: top - height,
            width: right - left,
            height: height,
            rowHeight: rowHeight,
            rowCount: paragraphs.length
        };
    }

    /**
     * 表全体を囲む長方形を作る
     * @param {Object} tableBox - measureTableBox() の戻り値
     * @param {Layer} targetLayer - 作図先のレイヤー
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function createOuterRectangle(tableBox, targetLayer, strokeWidthPt) {
        var outerRect = targetLayer.pathItems.rectangle(tableBox.top, tableBox.left, tableBox.width, tableBox.height);
        applyRuleStyle(outerRect, strokeWidthPt);
    }

    /**
     * 表の上端と下端に罫線を引く
     * @param {Object} tableBox - measureTableBox() の戻り値
     * @param {Layer} targetLayer - 作図先のレイヤー
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function createTopBottomRules(tableBox, targetLayer, strokeWidthPt) {
        createRuleLine(targetLayer, [tableBox.left, tableBox.top], [tableBox.right, tableBox.top], strokeWidthPt);
        createRuleLine(targetLayer, [tableBox.left, tableBox.bottom], [tableBox.right, tableBox.bottom], strokeWidthPt);
    }

    /**
     * 行の区切りに横罫を引く（1行目と2行目の間は firstWidthPt）
     * @param {Object} tableBox - measureTableBox() の戻り値
     * @param {Layer} targetLayer - 作図先のレイヤー
     * @param {number} strokeWidthPt - 線幅（pt）
     * @param {number} firstWidthPt - 1本目の区切りの線幅（pt）
     * @returns {void}
     */
    function createRowSeparators(tableBox, targetLayer, strokeWidthPt, firstWidthPt) {
        for (var row = 1; row < tableBox.rowCount; row++) {
            var separatorY = tableBox.top - tableBox.rowHeight * row;
            var separatorWidthPt = (row === 1) ? firstWidthPt : strokeWidthPt;
            createRuleLine(targetLayer, [tableBox.left, separatorY], [tableBox.right, separatorY], separatorWidthPt);
        }
    }

    /**
     * タブ位置に縦罫を引く
     * @param {Object} tableBox - measureTableBox() の戻り値
     * @param {Layer} targetLayer - 作図先のレイヤー
     * @param {number[]} tabPositions - 表の左端からの距離の配列
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function createColumnRules(tableBox, targetLayer, tabPositions, strokeWidthPt) {
        for (var i = 0; i < tabPositions.length; i++) {
            var ruleX = tableBox.left + tabPositions[i];
            createRuleLine(targetLayer, [ruleX, tableBox.top], [ruleX, tableBox.bottom], strokeWidthPt);
        }
    }

    /**
     * 行ごとの背景（塗りの長方形）を作る
     * @param {Object} tableBox - measureTableBox() の戻り値
     * @param {Layer} targetLayer - 作図先のレイヤー
     * @param {boolean} heading - 1行目を見出しの濃さにするか
     * @returns {void}
     */
    function createRowFills(tableBox, targetLayer, heading) {
        for (var row = 0; row < tableBox.rowCount; row++) {
            var rowTop = tableBox.top - tableBox.rowHeight * row;
            var rowRect = targetLayer.pathItems.rectangle(rowTop, tableBox.left, tableBox.width, tableBox.rowHeight);
            rowRect.stroked = false;
            rowRect.filled = true;

            /* row は0始まりなので、偶数の row が1・3・5…行目 / row is 0-based: even rows are the 1st, 3rd, 5th… */
            var fillBlack = (row % 2 === 0) ? ROW_FILL_BLACK_ODD : ROW_FILL_BLACK_EVEN;
            if (heading && row === 0) fillBlack = ROW_FILL_BLACK_HEADING;
            rowRect.fillColor = createBlackColor(fillBlack);

            selectAndSendToBack(rowRect);
        }
    }

    /**
     * 2点を結ぶ罫線を作る
     * @param {Layer} targetLayer - 作図先のレイヤー
     * @param {number[]} startPoint - 始点 [x, y]
     * @param {number[]} endPoint - 終点 [x, y]
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function createRuleLine(targetLayer, startPoint, endPoint, strokeWidthPt) {
        var ruleLine = targetLayer.pathItems.add();
        ruleLine.setEntirePath([startPoint, endPoint]);
        applyRuleStyle(ruleLine, strokeWidthPt);
    }

    /**
     * 罫線の見た目（塗りなし・黒の線）を適用する
     * @param {PathItem} ruleItem - 対象のパス
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function applyRuleStyle(ruleItem, strokeWidthPt) {
        ruleItem.filled = false;
        ruleItem.stroked = true;
        ruleItem.strokeWidth = strokeWidthPt;
        ruleItem.strokeColor = createBlackColor(RULE_BLACK);
        selectAndSendToBack(ruleItem);
    }

    /**
     * 作ったパスを選択し、レイヤーの最背面へ送る（テキストの背面に置くため）
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function selectAndSendToBack(pathItem) {
        pathItem.selected = true;
        pathItem.move(pathItem.layer, ElementPlacement.PLACEATEND);
    }

    /**
     * K だけの CMYK カラーを作る
     * @param {number} blackPercent - K の値（0〜100）
     * @returns {CMYKColor} 作ったカラー
     */
    function createBlackColor(blackPercent) {
        var blackColor = new CMYKColor();
        blackColor.cyan = 0;
        blackColor.magenta = 0;
        blackColor.yellow = 0;
        blackColor.black = blackPercent;
        return blackColor;
    }

    /**
     * 段落のタブストップから縦罫の位置（テキストフレーム左端からの距離）を集める
     * 段落ごとにタブストップが違うことがあるので、列ごとに最も右の位置を取る（テキストは読むだけ）
     * @param {Paragraphs} paragraphs - textFrame.paragraphs
     * @returns {number[]} タブ位置の配列（無ければ空）
     */
    function getColumnTabPositions(paragraphs) {
        var maxPositions = [];
        for (var p = 0; p < paragraphs.length; p++) {
            var tabStops = [];
            try {
                tabStops = paragraphs[p].paragraphAttributes.tabStops || [];
            } catch (e) {
                /* 段落属性を読めない段落は飛ばす（旧実装の防御を踏襲）/ skip paragraphs whose attributes cannot be read */
            }

            for (var i = 0; i < tabStops.length; i++) {
                var tabPosition = tabStops[i].position;
                if (!(tabPosition > 0)) continue;
                if (maxPositions[i] === undefined || tabPosition > maxPositions[i]) maxPositions[i] = tabPosition;
            }
        }

        /* 0以下しかない列は穴になるので詰める / drop holes left by columns without a positive stop */
        var tabPositions = [];
        for (var j = 0; j < maxPositions.length; j++) {
            if (maxPositions[j] !== undefined) tabPositions.push(maxPositions[j]);
        }
        return tabPositions;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択からテキストフレームとパスを振り分ける
     * @param {Array} selectedObjects - doc.selection
     * @returns {{textFrames: TextFrame[], pathItems: PathItem[]}} 種類ごとの配列
     */
    function collectTextFramesAndPaths(selectedObjects) {
        var selectedItems = { textFrames: [], pathItems: [] };
        for (var i = 0; i < selectedObjects.length; i++) {
            var selectedItem = selectedObjects[i];
            /* 文字の編集中は selection が TextRange になり、要素を持たない / a TextRange selection has no items */
            if (!selectedItem) continue;
            if (selectedItem.typename === "TextFrame") selectedItems.textFrames.push(selectedItem);
            if (selectedItem.typename === "PathItem") selectedItems.pathItems.push(selectedItem);
        }
        return selectedItems;
    }

    /**
     * 選択に含まれていたパス（前回の罫線など）を削除する
     * @param {PathItem[]} pathItems - 削除するパス
     * @returns {void}
     */
    function removeSelectedPaths(pathItems) {
        for (var i = 0; i < pathItems.length; i++) {
            pathItems[i].remove();
        }
    }

    /**
     * メイン処理
     * ドキュメント・選択・テキストフレームが無いときは何も表示せずに終了する
     * @returns {void}
     */
    function main() {
        try {
            if (!app.documents.length) return;
            var doc = app.activeDocument;

            var selectedObjects = doc.selection;
            if (!selectedObjects || selectedObjects.length === 0) return;
            if (selectedObjects.length >= MAX_SELECTION_COUNT) return;

            var selectedItems = collectTextFramesAndPaths(selectedObjects);
            if (selectedItems.textFrames.length === 0) return;

            /* 線幅の初期値は［キー入力］の移動距離 / the default stroke width is the keyboard increment */
            var rulerUnit = getUnitInfo("rulerType");
            var cursorKeyLengthPt = app.preferences.getRealPreference("cursorKeyLength");
            var defaultStrokeWidth = roundToThousandths(cursorKeyLengthPt / rulerUnit.pointsPerUnit);

            var tableSettings = showTableSettingsDialog(defaultStrokeWidth, rulerUnit);
            if (tableSettings === null) return;

            removeSelectedPaths(selectedItems.pathItems);

            for (var i = 0; i < selectedItems.textFrames.length; i++) {
                createTableForTextFrame(selectedItems.textFrames[i], doc.activeLayer, tableSettings);
            }
        } catch (e) {
            /* 仕様：alert は出さず、コンソールにだけ書く / by design, log to the console instead of an alert */
            $.writeln("エラー: " + e);
        }
    }

    main();

})();
