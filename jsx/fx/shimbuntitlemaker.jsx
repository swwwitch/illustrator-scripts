#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストまたはグループの左右に、二重線を描画します。
上下左右のマージンと二重線の間隔はダイアログボックスで指定できます。

### Overview

Draws a pair of double rules on the left and right sides of the selected text or group.
The top, bottom, left and right margins and the gap between the two rules are set in a dialog.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ShimbunTitleMaker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var DEFAULT_SIDE_MARGIN_PT = 8;      /* 左右マージンの既定値（pt）/ Default side margin */
    var DEFAULT_VERTICAL_MARGIN_PT = 0;  /* 上下マージンの既定値（pt）/ Default vertical margin */
    var DEFAULT_LINE_GAP_PT = 5;         /* 二重線の間隔の既定値（pt）/ Default gap between the two rules */
    var RULE_STROKE_WIDTH_PT = 0.6;      /* 二重線の線幅（pt）/ Stroke width of the rules */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 15;              /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10]; /* パネルの余白 [左,上,右,下] / panel margins */
    var FIELD_LABEL_WIDTH = 56;           /* 項目名の幅 / width of the field labels */
    var FIELD_CHARS = 5;                  /* 数値入力欄の幅（文字数）/ width of the numeric fields */
    var FIELD_ROW_SPACING = 6;            /* 入力行の間隔 / spacing inside a field row */
    var BUTTON_SPACING = 10;              /* ボタンの間隔 / spacing between buttons */

    /**
     * パネルを縦並び・左揃えにして共通の余白を付ける
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function applyPanelLayout(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "左右に二重線", en: "Side Double Rules" }
        },
        panel: {
            margin: { ja: "マージン", en: "Margins" },
            rule: { ja: "二重線", en: "Rules" }
        },
        fieldLabel: {
            marginTop: { ja: "上", en: "Top" },
            marginBottom: { ja: "下", en: "Bottom" },
            marginLeft: { ja: "左", en: "Left" },
            marginRight: { ja: "右", en: "Right" },
            lineGap: { ja: "間隔", en: "Gap" }
        },
        tooltip: {
            marginVertical: {
                ja: "二重線を対象の上端／下端から伸ばす量。マイナスで内側に縮みます。",
                en: "How far the rules extend beyond the top and bottom of the object. Negative values pull them inward."
            },
            marginHorizontal: {
                ja: "対象と内側の線との距離。マイナスで対象に重なります。",
                en: "Distance between the object and the inner rule. Negative values move it over the object."
            },
            lineGap: { ja: "内側の線と外側の線の間隔", en: "Distance between the inner and outer rule" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTarget: {
                ja: "テキストオブジェクトまたはグループを1つ選択してください。",
                en: "Select one text object or one group."
            }
        },
        itemName: {
            ruleGroup: { ja: "左右二重線", en: "Side rules" }
        }
    };

    /**
     * ラベルを表示言語で取得する
     * @param {string} labelPath - "panel.margin" のようなラベルのパス
     * @returns {string} 表示用の文字列
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length && labelNode; i++) {
            labelNode = labelNode[pathKeys[i]];
        }
        if (!labelNode) return labelPath;
        return labelNode[uiLang] || labelNode.en || labelNode.ja || labelPath;
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
    // 入力の読み取り / Input handling
    // =========================================

    /**
     * 数値入力欄を↑↓キーで増減できるようにする（Shiftで10、Optionで0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            /* Shiftで10単位、Optionで0.1単位 / Shift steps by 10, Option by 0.1 */
            var delta = keyboard.shiftKey ? 10 : (keyboard.altKey ? 0.1 : 1);

            if (keyboard.shiftKey) {
                /* 10の倍数にスナップ / snap to multiples of ten */
                value = (event.keyName == "Up") ?
                    Math.ceil((value + 1) / delta) * delta :
                    Math.floor((value - 1) / delta) * delta;
            } else {
                value += (event.keyName == "Up") ? delta : -delta;
            }
            event.preventDefault();

            /* Option時だけ小数第1位まで残す / keep one decimal only while Option is held */
            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);
        });
    }

    /**
     * 数値の入力文字列を読む。空白を除き、全角のカンマ・小数点を半角にし、
     * 「1,5」のような小数カンマは小数点に、「1,000」のような桁区切りは取り除く
     * @param {string} inputText - 入力された文字列
     * @returns {number} 読み取った数値（読めなければ NaN）
     */
    function parseLocaleNumber(inputText) {
        var normalizedText = String(inputText);
        normalizedText = normalizedText.replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        if (normalizedText.indexOf(",") >= 0 && normalizedText.indexOf(".") < 0) {
            normalizedText = normalizedText.replace(/,/g, ".");
        } else {
            normalizedText = normalizedText.replace(/,/g, "");
        }
        return Number(normalizedText);
    }

    /**
     * 入力欄の値を pt に換算する（読めないときは既定値）
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} fallbackPt - 読めなかったときに使う値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @param {boolean} allowNegative - 負の値を許すか
     * @returns {number} 換算した値（pt）
     */
    function readFieldAsPoints(inputField, fallbackPt, pointsPerUnit, allowNegative) {
        var enteredValue = parseLocaleNumber(inputField.text);
        if (isNaN(enteredValue)) return fallbackPt;
        if (!allowNegative && enteredValue < 0) return 0;
        return enteredValue * pointsPerUnit;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確認し、ダイアログの設定で二重線を描画する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var targetItem = getTargetItem(doc.selection);
        if (!targetItem) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        var ruleSettings = showDialog();
        if (!ruleSettings) return;

        drawSideRules(doc, targetItem, ruleSettings);
    }

    /**
     * 基準にできる選択かを判定して返す（テキスト1つ、またはグループ1つ）
     * @param {Object} currentSelection - ドキュメントの選択
     * @returns {PageItem|null} 基準にするオブジェクト（対象外なら null）
     */
    function getTargetItem(currentSelection) {
        if (!currentSelection || currentSelection.length !== 1) return null;

        var selectedItem = currentSelection[0];
        var isSupported = (
            selectedItem.typename === "TextFrame" ||
            selectedItem.typename === "GroupItem"
        );
        return isSupported ? selectedItem : null;
    }

    /**
     * 対象の左右に二重線を描画する
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {{marginTopPt: number, marginBottomPt: number, marginLeftPt: number, marginRightPt: number, lineGapPt: number}} ruleSettings - マージンと間隔（pt）
     * @returns {GroupItem} 生成した二重線のグループ
     */
    function drawSideRules(doc, targetItem, ruleSettings) {
        /* visibleBounds: [左, 上, 右, 下] / visibleBounds: [left, top, right, bottom] */
        var bounds = targetItem.visibleBounds;
        var lineTopY = bounds[1] + ruleSettings.marginTopPt;
        var lineBottomY = bounds[3] - ruleSettings.marginBottomPt;
        var innerLeftX = bounds[0] - ruleSettings.marginLeftPt;
        var innerRightX = bounds[2] + ruleSettings.marginRightPt;

        var ruleGroup = targetItem.layer.groupItems.add();
        ruleGroup.name = getLabel("itemName.ruleGroup");

        var strokeColor = createBlackColor(doc);
        var lineXPositions = [innerLeftX, innerRightX];
        /* 間隔が0なら内側の線に重なるので、外側の線は描かない / A zero gap would stack the rules, so skip the outer pair */
        if (ruleSettings.lineGapPt > 0) {
            lineXPositions.push(innerLeftX - ruleSettings.lineGapPt);
            lineXPositions.push(innerRightX + ruleSettings.lineGapPt);
        }
        for (var i = 0; i < lineXPositions.length; i++) {
            createRuleLine(ruleGroup, lineXPositions[i], lineTopY, lineBottomY, strokeColor);
        }

        /* 対象のすぐ背面へ / Move the rules just behind the target */
        ruleGroup.move(targetItem, ElementPlacement.PLACEAFTER);
        return ruleGroup;
    }

    /**
     * 縦の直線を1本描く
     * @param {GroupItem} parentGroup - 追加先のグループ
     * @param {number} lineX - 線の X 座標
     * @param {number} topY - 上端の Y 座標
     * @param {number} bottomY - 下端の Y 座標
     * @param {Object} strokeColor - 線の色
     * @returns {PathItem} 生成したパス
     */
    function createRuleLine(parentGroup, lineX, topY, bottomY, strokeColor) {
        var linePath = parentGroup.pathItems.add();
        linePath.setEntirePath([
            [lineX, topY],
            [lineX, bottomY]
        ]);
        linePath.stroked = true;
        linePath.filled = false;
        linePath.strokeWidth = RULE_STROKE_WIDTH_PT;
        linePath.strokeColor = strokeColor;
        linePath.strokeCap = StrokeCap.BUTTENDCAP;
        return linePath;
    }

    /**
     * ドキュメントのカラーモードに合わせた黒を返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {Object} RGBColor または CMYKColor
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var rgbBlack = new RGBColor();
            rgbBlack.red = 0;
            rgbBlack.green = 0;
            rgbBlack.blue = 0;
            return rgbBlack;
        }
        var cmykBlack = new CMYKColor();
        cmykBlack.cyan = 0;
        cmykBlack.magenta = 0;
        cmykBlack.yellow = 0;
        cmykBlack.black = 100;
        return cmykBlack;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログと各コントロールを作る
     * @param {string} unitLabel - 入力欄に添える単位の表示名
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {Object} ダイアログ（dialog）と各入力欄の参照
     */
    function buildDialog(unitLabel, pointsPerUnit) {
        var ruleDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        ruleDialog.orientation = "column";
        ruleDialog.alignChildren = ["fill", "top"];
        ruleDialog.margins = DIALOG_MARGINS;

        /* マージン（pt の既定値を現在の単位に換算して表示）/ Margins (defaults converted to the current unit) */
        var marginPanel = ruleDialog.add("panel", undefined, getLabel("panel.margin"));
        applyPanelLayout(marginPanel);

        var verticalDefault = formatUnitValue(DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit);
        var sideDefault = formatUnitValue(DEFAULT_SIDE_MARGIN_PT, pointsPerUnit);

        var marginTopInput = addNumericFieldRow(marginPanel, "fieldLabel.marginTop", verticalDefault, unitLabel, "tooltip.marginVertical", true);
        var marginBottomInput = addNumericFieldRow(marginPanel, "fieldLabel.marginBottom", verticalDefault, unitLabel, "tooltip.marginVertical", true);
        var marginLeftInput = addNumericFieldRow(marginPanel, "fieldLabel.marginLeft", sideDefault, unitLabel, "tooltip.marginHorizontal", true);
        var marginRightInput = addNumericFieldRow(marginPanel, "fieldLabel.marginRight", sideDefault, unitLabel, "tooltip.marginHorizontal", true);

        /* 二重線の間隔 / Gap between the two rules */
        var rulePanel = ruleDialog.add("panel", undefined, getLabel("panel.rule"));
        applyPanelLayout(rulePanel);

        var lineGapInput = addNumericFieldRow(
            rulePanel,
            "fieldLabel.lineGap",
            formatUnitValue(DEFAULT_LINE_GAP_PT, pointsPerUnit),
            unitLabel,
            "tooltip.lineGap",
            false
        );

        /* ボタンエリア（左：スペーサー／右：キャンセル・OK）/ Button row: spacer, then Cancel/OK */
        var btnRowGroup = ruleDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = ["fill", "top"];

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_SPACING;

        btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            dialog: ruleDialog,
            marginTopInput: marginTopInput,
            marginBottomInput: marginBottomInput,
            marginLeftInput: marginLeftInput,
            marginRightInput: marginRightInput,
            lineGapInput: lineGapInput
        };
    }

    /**
     * 「項目名＋数値入力欄＋単位」の行を生成する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} fieldLabelPath - 項目名のラベルパス
     * @param {string} initialText - 入力欄の初期値
     * @param {string} unitLabel - 入力欄の右に添える単位
     * @param {string} tooltipPath - helpTip のラベルパス
     * @param {boolean} allowNegative - 負の値を許すか
     * @returns {EditText} 生成した入力欄
     */
    function addNumericFieldRow(parentContainer, fieldLabelPath, initialText, unitLabel, tooltipPath, allowNegative) {
        var fieldRow = parentContainer.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        fieldRow.spacing = FIELD_ROW_SPACING;

        var fieldLabel = fieldRow.add("statictext", undefined, labelText(fieldLabelPath));
        fieldLabel.preferredSize = [FIELD_LABEL_WIDTH, -1];
        fieldLabel.justify = "right";

        var inputField = fieldRow.add("edittext", undefined, initialText);
        inputField.characters = FIELD_CHARS;
        inputField.helpTip = getLabel(tooltipPath);
        changeValueByArrowKey(inputField, allowNegative);

        fieldRow.add("statictext", undefined, unitLabel);
        return inputField;
    }

    /**
     * pt の値を現在の単位の表示文字列にする（小数第2位まで）
     * @param {number} valuePt - 値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {string} 入力欄に入れる文字列
     */
    function formatUnitValue(valuePt, pointsPerUnit) {
        var unitValue = valuePt / pointsPerUnit;
        return String(Math.round(unitValue * 100) / 100);
    }

    /**
     * ダイアログを表示して設定を取得する
     * @returns {Object|null} マージンと間隔（pt）。キャンセル時は null
     */
    function showDialog() {
        var unitInfo = getUnitInfo();
        var pointsPerUnit = unitInfo.pointsPerUnit;
        var dialogUI = buildDialog(unitInfo.label, pointsPerUnit);

        if (dialogUI.dialog.show() !== 1) return null;

        return {
            marginTopPt: readFieldAsPoints(dialogUI.marginTopInput, DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit, true),
            marginBottomPt: readFieldAsPoints(dialogUI.marginBottomInput, DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit, true),
            marginLeftPt: readFieldAsPoints(dialogUI.marginLeftInput, DEFAULT_SIDE_MARGIN_PT, pointsPerUnit, true),
            marginRightPt: readFieldAsPoints(dialogUI.marginRightInput, DEFAULT_SIDE_MARGIN_PT, pointsPerUnit, true),
            lineGapPt: readFieldAsPoints(dialogUI.lineGapInput, DEFAULT_LINE_GAP_PT, pointsPerUnit, false)
        };
    }

    main();
})();
