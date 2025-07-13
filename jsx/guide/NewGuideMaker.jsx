#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

NewGuideMaker.jsx

### 概要

- ダイアログで方向、位置、単位、対象（カンバス/アートボード）を指定してガイドを作成するスクリプト

### 主な機能

- 水平方向・垂直方向のガイド作成
- 位置と裁ち落とし（マージン）の指定
- 単位選択（px, pt, mm）
- 環境設定に基づいた初期単位設定
- 上下キーによる数値調整

### 処理の流れ

1. ダイアログ表示
2. 設定入力
3. OK 押下でガイド作成

### メモ

Photoshopの［新規ガイド］ダイアログを参考に、次の点を

- アートボード、カンバスの選択
- 裁ち落とし（マージン）の入力
- 水平・垂直の選択（Hキー、Vキー）
- ↑↓キー、shift + ↑↓での数値調整
- ガイドのプレビュー線を描画
- ガイド化時の色は設定不可（Illustratorの仕様）
- ガイド化時の線幅は 0.1pt
- ガイドは「_guide」レイヤーに作成し、ロック

### 更新履歴

- v1.0 (20250713) : 初期バージョン


### Script Name:

NewGuideMaker.jsx

### Overview

- Script to create guides in Illustrator by specifying direction, position, unit, and target (canvas or artboard) via dialog

### Features

- Create horizontal or vertical guides
- Specify position and margin (bleed)
- Unit selection (px, pt, mm)
- Auto get initial unit from preferences
- Increment/decrement values with up/down keys

### Flow

1. Show dialog
2. Input settings
3. Create guide on OK

### Change Log

- v1.0 (20250713): Initial version

*/

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    try {
        var lang = app.locale || app.language || "ja";
        if (lang.indexOf("en") === 0) {
            return "en";
        }
    } catch (e) {
        // fallback
    }
    return "ja";
}

var LABELS = {
    dialogTitle: {
        ja: "ガイド作成 " + SCRIPT_VERSION,
        en: "Create Guide " + SCRIPT_VERSION
    },
    target: { ja: "対象", en: "Target" },
    canvas: { ja: "カンバス", en: "Canvas" },
    artboard: { ja: "アートボード", en: "Artboard" },
    bleed: { ja: "裁ち落とし:", en: "Bleed:" },
    direction: { ja: "方向", en: "Direction" },
    horizontal: { ja: "水平方向", en: "Horizontal" },
    vertical: { ja: "垂直方向", en: "Vertical" },
    position: { ja: "位置:", en: "Position:" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
};

var lang = getCurrentLang();

var canvasSize = 227 * 72;

/* 
  ダイアログの構築 / Build dialog UI 
*/
function buildDialog() {
    var doc = app.activeDocument;

    // --- _guide レイヤー取得または作成・ロック解除 ---
    var guideLayer = null;
    var docLayers = doc.layers;
    var found = false;
    for (var i = 0; i < docLayers.length; i++) {
        if (docLayers[i].name === "_guide") {
            guideLayer = docLayers[i];
            found = true;
            break;
        }
    }
    if (!guideLayer) {
        guideLayer = docLayers.add();
        guideLayer.name = "_guide";
    }
    guideLayer.locked = false;

    var dlg = new Window("dialog");
    dlg.text = LABELS.dialogTitle[lang];
    dlg.orientation = "row";
    dlg.alignChildren = ["center", "top"];
    dlg.spacing = 20;
    dlg.margins = 16;

    var leftGroup = dlg.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = ["left", "top"];
    leftGroup.spacing = 10;

    // 対象選択パネル / Target selection panel
    var targetPanel = leftGroup.add("panel");
    targetPanel.text = LABELS.target[lang];
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["left", "top"];
    targetPanel.margins = [15, 20, 15, 10];

    var targetRadioGroup = targetPanel.add("group");
    targetRadioGroup.orientation = "column";
    targetRadioGroup.alignChildren = ["left", "center"];
    var canvasRadio = targetRadioGroup.add("radiobutton", undefined, LABELS.canvas[lang]);
    var artboardRadio = targetRadioGroup.add("radiobutton", undefined, LABELS.artboard[lang]);
    artboardRadio.value = true;

    // 裁ち落とし（マージン）入力 / Bleed (margin) input
    var marginGroup = targetPanel.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = ["left", "center"];
    marginGroup.spacing = 6;

    var marginLabel = marginGroup.add("statictext", undefined, LABELS.bleed[lang]);
    var marginEdit = marginGroup.add('edittext {characters: 5}');
    marginEdit.text = "0";
    changeValueByArrowKey(marginEdit, function(){ drawPreviewLine(); });

    var marginUnitLabel = marginGroup.add("statictext", undefined, getCurrentUnitLabel());

    marginEdit.enabled = true;

    // 方向選択パネル / Direction selection panel
    var directionPanel = leftGroup.add("panel");
    directionPanel.text = LABELS.direction[lang];
    directionPanel.orientation = "column";
    directionPanel.alignChildren = ["left", "top"];
    directionPanel.margins = [15, 20, 15, 10];

    var radioGroup = directionPanel.add("group");
    radioGroup.orientation = "row";
    radioGroup.alignChildren = ["left", "center"];
    var horizontalRadio = radioGroup.add("radiobutton", undefined, LABELS.horizontal[lang]);
    var verticalRadio = radioGroup.add("radiobutton", undefined, LABELS.vertical[lang]);
    horizontalRadio.value = true;

    // 位置と単位選択 / Position and unit selection
    var positionGroup = directionPanel.add("group");
    positionGroup.orientation = "row";
    positionGroup.alignChildren = ["left", "center"];
    positionGroup.margins = [0, 6, 0, 0];
    positionGroup.spacing = 4;

    var positionLabel = positionGroup.add("statictext", undefined, LABELS.position[lang]);
    var positionEdit = positionGroup.add('edittext {characters: 5}');
    positionEdit.text = "0";
    changeValueByArrowKey(positionEdit, function(){ drawPreviewLine(); });

    // 単位ドロップダウンを unitLabelMap ですべての単位に変更
    var unitOptions = [];
    for (var code in unitLabelMap) {
        if (unitLabelMap.hasOwnProperty(code)) {
            unitOptions.push(unitLabelMap[code]);
        }
    }
    var unitDropdown = positionGroup.add("dropdownlist", undefined, unitOptions);
    var currentLabel = getCurrentUnitLabel();
    var foundIndex = null;
    for (var i = 0; i < unitDropdown.items.length; i++) {
        if (unitDropdown.items[i].text === currentLabel) {
            foundIndex = i;
            break;
        }
    }
    if (foundIndex === null) {
        foundIndex = 0;
    }
    unitDropdown.selection = foundIndex;

    // ボタン群 / Button group
    var buttonGroup = dlg.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.alignChildren = ["left", "center"];
    buttonGroup.spacing = 10;

    var okButton = buttonGroup.add("button", undefined, LABELS.ok[lang]);
    okButton.preferredSize.width = 90;
    var cancelButton = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    cancelButton.preferredSize.width = 90;

    // プレビュー線保持用 / Preview line holder
    var previewLine = null;

    // プレビュー線を削除 / Remove preview line
    function removePreviewLine() {
        if (previewLine && !previewLine.locked && previewLine.layer && !previewLine.layer.locked) {
            try { previewLine.remove(); } catch(e){}
            previewLine = null;
        }
    }

    // プレビュー線を描画 / Draw preview line
    function drawPreviewLine() {
        removePreviewLine();
        var pos = parseFloat(positionEdit.text);
        if (isNaN(pos)) {
            return;
        }
        var unit = unitDropdown.selection.text;
        var posPt = convertToPt(pos, unit);

        // --- _guide レイヤー取得または作成・ロック管理 ---
        var guideLayer = null;
        var docLayers = doc.layers;
        var found = false;
        for (var i = 0; i < docLayers.length; i++) {
            if (docLayers[i].name === "_guide") {
                guideLayer = docLayers[i];
                found = true;
                break;
            }
        }
        var origLock = false;
        if (!guideLayer) {
            guideLayer = docLayers.add();
            guideLayer.name = "_guide";
        } else {
            origLock = guideLayer.locked;
            guideLayer.locked = false;
        }

        var path;
        if (canvasRadio.value) {
            var size = canvasSize;
            if (horizontalRadio.value) {
                path = guideLayer.pathItems.add();
                path.setEntirePath([
                    [-size, -posPt],
                    [size, -posPt]
                ]);
            } else {
                path = guideLayer.pathItems.add();
                path.setEntirePath([
                    [posPt, size],
                    [posPt, -size]
                ]);
            }
        } else {
            var margin = parseFloat(marginEdit.text);
            if (isNaN(margin)) {
                guideLayer.locked = origLock; // 念のため
                return;
            }
            var marginPt = convertToPt(margin, unit);
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            var abRect = ab.artboardRect;
            var abLeft = abRect[0];
            var abTop = abRect[1];
            var abRight = abRect[2];
            var abBottom = abRect[3];
            if (horizontalRadio.value) {
                path = guideLayer.pathItems.add();
                path.setEntirePath([
                    [abLeft - marginPt, abTop - posPt],
                    [abRight + marginPt, abTop - posPt]
                ]);
            } else {
                path = guideLayer.pathItems.add();
                path.setEntirePath([
                    [abLeft + posPt, abTop + marginPt],
                    [abLeft + posPt, abBottom - marginPt]
                ]);
            }
        }
        path.stroked = true;
        path.filled = false;
        path.strokeWidth = 1.0;
        path.strokeColor = makeBlackColor();
        path.guides = false; // プレビューはガイド化しない / Preview is not a guide
        previewLine = path;
        guideLayer.locked = origLock;
        app.redraw();
    }

    // プレビュー用ブラック(K100)カラー / Black color for preview (K100)
    function makeBlackColor() {
        var c = new GrayColor();
        c.gray = 100;
        return c;
    }

    // OKボタンクリック時の処理 / OK button click handler
    okButton.onClick = function() {
        // プレビュー線がある場合、それをガイド化 / Convert preview line to guide if exists
        if (previewLine && previewLine.layer && !previewLine.layer.locked) {
            previewLine.guides = true;
            previewLine.strokeWidth = 0.1;
            // 色はガイド化時は不要 / Color not needed when guide
        }
        previewLine = null;
        dlg.close();
    };

    // 対象ラジオボタンの切り替えで裁ち落とし入力の有効/無効切り替え / Enable/disable margin input based on target
    canvasRadio.onClick = function() {
        marginEdit.enabled = false;
        drawPreviewLine();
    };
    artboardRadio.onClick = function() {
        marginEdit.enabled = true;
        drawPreviewLine();
    };

    // プレビュー更新イベントを各UIに追加 / Add preview update events to UI elements
    positionEdit.addEventListener("changing", function(){ drawPreviewLine(); });
    marginEdit.addEventListener("changing", function(){ drawPreviewLine(); });
    unitDropdown.onChange = function(){ drawPreviewLine(); };
    horizontalRadio.onClick = function(){ drawPreviewLine(); };
    verticalRadio.onClick = function(){ drawPreviewLine(); };

    // --- H/Vキーで方向切り替え ---
    dlg.addEventListener("keydown", function(event) {
        if (event.keyName == "H" || event.keyName == "h") {
            horizontalRadio.value = true;
            verticalRadio.value = false;
            drawPreviewLine();
            event.preventDefault();
        } else if (event.keyName == "V" || event.keyName == "v") {
            horizontalRadio.value = false;
            verticalRadio.value = true;
            drawPreviewLine();
            event.preventDefault();
        }
    });

    // ダイアログ表示時に「位置」テキストフィールドにフォーカスを設定 / Focus position input on dialog show
    positionEdit.active = true;
    // 初回プレビュー / Initial preview
    drawPreviewLine();
    // ダイアログ閉じたらプレビュー線消す / Remove preview line on dialog close
    dlg.onClose = function() { removePreviewLine(); };
    return dlg;
}

/* メイン処理 / Main function */
function main() {
    var dlg = buildDialog();
    dlg.show();
}

main();

/*
  edittext に上下キーで値を増減させるイベントを追加（コールバック対応版）
  Add event to increment/decrement value in edittext by up/down arrow keys (with onUpdate callback)
*/
function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.shiftKey ? 10 : 1;
        if (event.keyName == "Up") {
            value += delta;
            event.preventDefault();
        } else if (event.keyName == "Down") {
            value -= delta;
            event.preventDefault();
        }
        editText.text = value;
        if (typeof onUpdate === "function") {
            onUpdate(editText.text);
        }
    });
}

/* 単位コードとラベルのマップ / Map of unit codes and labels */
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
/* 単位ラベルからコードへの逆引きマップ / Reverse map: label to code */
var unitLabelToCodeMap = {};
for (var code in unitLabelMap) {
    if (unitLabelMap.hasOwnProperty(code)) {
        unitLabelToCodeMap[unitLabelMap[code]] = parseInt(code, 10);
    }
}

/* 現在の単位ラベルを取得 / Get current unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

/* 単位ラベルから単位コードを取得 / Get unit code from unit label */
function getUnitCodeFromLabel(label) {
    return unitLabelToCodeMap[label] !== undefined ? unitLabelToCodeMap[label] : 2; // default pt
}

/* 単位コードに応じた pt 係数を取得 / Get pt factor from unit code */
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0; // in
        case 1:
            return 72.0 / 25.4; // mm
        case 2:
            return 1.0; // pt
        case 3:
            return 12.0; // pica
        case 4:
            return 72.0 / 2.54; // cm
        case 5:
            return 72.0 / 25.4 * 0.25; // Q or H
        case 6:
            return 1.0; // px
        case 7:
            return 72.0 * 12.0; // ft/in
        case 8:
            return 72.0 / 25.4 * 1000.0; // m
        case 9:
            return 72.0 * 36.0; // yd
        case 10:
            return 72.0 * 12.0; // ft
        default:
            return 1.0;
    }
}

/* 値と単位ラベルからptに変換（数値チェックを追加） / Convert value and unit label to pt with number check */
function convertToPt(value, label) {
    var numValue = Number(value);
    if (isNaN(numValue)) {
        return 0;
    }
    var unitCode = getUnitCodeFromLabel(label);
    var factor = getPtFactorFromUnitCode(unitCode);
    return numValue * factor;
}