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

Photoshopの［新規ガイド］ダイアログボックスをベースに、以下の要素を盛り込んでいます。

- アートボード、カンバスの選択
- アートボード選択時には裁ち落とし（マージン）を設定可能
- 水平・垂直をHキー、Vキーで切替
- ↑↓キー、shift + ↑↓での数値調整
- ガイドのプレビュー線を描画
- ガイド化時の色は設定不可（Illustratorの仕様）
- ガイドは「_guide」レイヤーに作成し、ロック

### note

https://note.com/dtp_tranist/n/n1085336d7265

### 更新履歴

- v1.0 (20250713) : 初期バージョン
- v1.1 (20250714) : レイヤー選択機能追加、リピート機能追加
- v1.2 (20250715) : shift + ↑↓キーで値の増減ロジックを調整

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
- v1.1 (20250714): Added layer selection and repeat functionality
- v1.2 (20250715): Adjusted increment/decrement logic for shift + up/down keys

*/

var SCRIPT_VERSION = "v1.2";

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

/*
  UIラベル定義 / UI label definitions
*/
/*
  UIラベル定義 / UI label definitions
  - UI順に並べる / Ordered as appears in UI
*/
var LABELS = {
    dialogTitle: { ja: "ガイド作成 " + SCRIPT_VERSION, en: "Create Guide " + SCRIPT_VERSION },
    target: { ja: "対象", en: "Target" },
    canvas: { ja: "カンバス", en: "Canvas" },
    artboard: { ja: "アートボード", en: "Artboard" },
    bleed: { ja: "はみ出し", en: "Bleed" },
    direction: { ja: "方向", en: "Direction" },
    horizontal: { ja: "水平方向", en: "Horizontal" },
    vertical: { ja: "垂直方向", en: "Vertical" },
    position: { ja: "位置:", en: "Position:" },
    layer: { ja: "作成レイヤー", en: "Target Layer" },
    guideLayer: { ja: "_guideレイヤー", en: "_guide Layer" },
    currentLayer: { ja: "現在のレイヤー", en: "Current Layer" },
    repeat: { ja: "リピート", en: "Repeat" },
    guideCount: { ja: "ガイド数", en: "Guide Count" },
    distance: { ja: "距離", en: "Distance" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertLocked: { ja: "アクティブレイヤーがロックされています。", en: "The active layer is locked." }
};

var lang = getCurrentLang();

var canvasSize = 227 * 72;

/* 
  ダイアログの構築 / Build dialog UI 
*/
/* ダイアログの構築 / Build dialog UI */
function buildDialog() {
    var doc = app.activeDocument;

    // --- Shared unit options and selection ---
    var unitOptions = [];
    for (var code in unitLabelMap) {
        if (unitLabelMap.hasOwnProperty(code)) {
            unitOptions.push(unitLabelMap[code]);
        }
    }
    var currentLabel = getCurrentUnitLabel();
    var foundIndex = null;
    for (var i = 0; i < unitOptions.length; i++) {
        if (unitOptions[i] === currentLabel) {
            foundIndex = i;
            break;
        }
    }
    if (foundIndex === null) {
        foundIndex = 0;
    }
    // --- End shared unit options ---

    /* _guide レイヤー取得または作成・ロック解除 / Get or create and unlock _guide layer */
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

    /* 対象選択パネル / Target selection panel */
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

    /* 裁ち落とし（マージン）入力 / Bleed (margin) input */
    var marginGroup = targetPanel.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = ["left", "center"];
    marginGroup.spacing = 6;

    var marginLabel = marginGroup.add("statictext", undefined, LABELS.bleed[lang]);
    var marginEdit = marginGroup.add('edittext {characters: 5}');
    marginEdit.text = "0";
    changeValueByArrowKey(marginEdit, function() {
        drawPreviewLine();
    });

    var marginUnitLabel = marginGroup.add("statictext", undefined, getCurrentUnitLabel());

    marginEdit.enabled = true;

    /* 方向選択パネル / Direction selection panel */
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

    /* 作成レイヤーパネル / Target layer panel */
    var layerPanel = leftGroup.add("panel");
    layerPanel.text = LABELS.layer[lang];
    layerPanel.orientation = "column";
    layerPanel.alignChildren = ["left", "top"];
    layerPanel.margins = [15, 20, 15, 10];

    var layerRadioGroup = layerPanel.add("group");
    layerRadioGroup.orientation = "column";
    layerRadioGroup.alignChildren = ["left", "center"];
    var guideLayerRadio = layerRadioGroup.add("radiobutton", undefined, LABELS.guideLayer[lang]);
    var currentLayerRadio = layerRadioGroup.add("radiobutton", undefined, LABELS.currentLayer[lang]);
    guideLayerRadio.value = true; // default

    /* レイヤーラジオボタン切り替え時にプレビュー線を即時更新 / Update preview immediately when switching layer radio */
    currentLayerRadio.onClick = function() {
        removePreviewLine();
        drawPreviewLine();
    };
    guideLayerRadio.onClick = function() {
        removePreviewLine();
        drawPreviewLine();
    };

    /* リピートパネル / Repeat panel */
    var repeatPanel = leftGroup.add("panel");
    repeatPanel.text = LABELS.repeat[lang];
    repeatPanel.orientation = "column";
    repeatPanel.alignChildren = ["left", "top"];
    repeatPanel.margins = [15, 20, 15, 10];

    var repeatRadioGroup = repeatPanel.add("group");
    repeatRadioGroup.orientation = "column";
    repeatRadioGroup.alignChildren = ["left", "center"];

    /* リピート数入力欄 / Repeat count input */
    var repeatCountGroup = repeatRadioGroup.add("group");
    repeatCountGroup.orientation = "row";
    repeatCountGroup.alignChildren = ["left", "center"];
    repeatCountGroup.add("statictext", undefined, LABELS.guideCount[lang]);
    var repeatCountEdit = repeatCountGroup.add('edittext {characters: 4}');
    repeatCountEdit.text = "1";
    changeValueByArrowKey(repeatCountEdit, function() {
        drawPreviewLine();
    });

    /* 距離入力欄 / Distance input */
    var repeatDistanceGroup = repeatRadioGroup.add("group");
    repeatDistanceGroup.orientation = "row";
    repeatDistanceGroup.alignChildren = ["left", "center"];
    repeatDistanceGroup.add("statictext", undefined, LABELS.distance[lang]);
    var repeatDistanceEdit = repeatDistanceGroup.add('edittext {characters: 4}');
    repeatDistanceEdit.text = "0";
    changeValueByArrowKey(repeatDistanceEdit, function() {
        drawPreviewLine();
    });
    /* 距離単位ドロップダウン / Distance unit dropdown */
    var distanceUnitDropdown = repeatDistanceGroup.add("dropdownlist", undefined, unitOptions);
    distanceUnitDropdown.selection = foundIndex;
    distanceUnitDropdown.onChange = function() {
        drawPreviewLine();
    };

    /* 位置と単位選択 / Position and unit selection */
    var positionGroup = directionPanel.add("group");
    positionGroup.orientation = "row";
    positionGroup.alignChildren = ["left", "center"];
    positionGroup.margins = [0, 6, 0, 0];
    positionGroup.spacing = 4;

    var positionLabel = positionGroup.add("statictext", undefined, LABELS.position[lang]);
    var positionEdit = positionGroup.add('edittext {characters: 5}');
    positionEdit.text = "0";
    changeValueByArrowKey(positionEdit, function() {
        drawPreviewLine();
    });

    /* 単位ドロップダウンを unitLabelMap ですべての単位に変更 / Populate unit dropdown with all units */
    var unitDropdown = positionGroup.add("dropdownlist", undefined, unitOptions);
    unitDropdown.selection = foundIndex;

    /* ボタン群 / Button group */
    var buttonGroup = dlg.add("group");
    buttonGroup.orientation = "column";
    buttonGroup.alignChildren = ["left", "center"];
    buttonGroup.spacing = 10;

    var okButton = buttonGroup.add("button", undefined, LABELS.ok[lang]);
    okButton.preferredSize.width = 90;
    var cancelButton = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    cancelButton.preferredSize.width = 90;

    /* プレビュー線保持用 / Preview line holder */
    var previewLine = null;

    /* プレビュー線を削除 / Remove preview line */
    function removePreviewLine() {
        if (previewLine) {
            // previewLine may be an array or a single path
            if (previewLine instanceof Array) {
                for (var i = 0; i < previewLine.length; i++) {
                    var p = previewLine[i];
                    if (p && !p.locked && p.layer && !p.layer.locked) {
                        try {
                            p.remove();
                        } catch (e) {}
                    }
                }
            } else if (previewLine && !previewLine.locked && previewLine.layer && !previewLine.layer.locked) {
                try {
                    previewLine.remove();
                } catch (e) {}
            }
            previewLine = null;
        }
    }

    /* プレビュー線を描画 / Draw preview line */
    function drawPreviewLine() {
        removePreviewLine();
        var pos = parseFloat(positionEdit.text);
        if (isNaN(pos)) {
            return;
        }
        var unit = unitDropdown.selection.text;
        var posPt = convertToPt(pos, unit);

        // リピート数・距離取得
        var repeatCount = parseInt(repeatCountEdit.text, 10);
        if (isNaN(repeatCount) || repeatCount < 1) repeatCount = 1;
        var repeatDistance = parseFloat(repeatDistanceEdit.text);
        if (isNaN(repeatDistance) || repeatDistance <= 0) repeatDistance = 0;
        var repeatDistancePt = convertToPt(repeatDistance, distanceUnitDropdown.selection.text);

        /* レイヤー選択ロジック / Layer selection logic */
        var targetLayer = null;
        var origLock = false;
        var guideLayer = null;
        var docLayers = doc.layers;
        // Find or create _guide layer for possible use and lock management
        for (var i = 0; i < docLayers.length; i++) {
            if (docLayers[i].name === "_guide") {
                guideLayer = docLayers[i];
                break;
            }
        }
        if (!guideLayer) {
            guideLayer = docLayers.add();
            guideLayer.name = "_guide";
        }
        origLock = guideLayer.locked;
        guideLayer.locked = false;

        if (guideLayerRadio.value) {
            // _guide layer logic (default)
            targetLayer = guideLayer;
        } else if (currentLayerRadio.value) {
            // Directly use the current active layer as target
            targetLayer = doc.activeLayer;
            if (targetLayer.locked) {
                alert(LABELS.alertLocked[lang]);
                if (guideLayerRadio.value) {
                    guideLayer.locked = origLock;
                }
                return;
            }
        }

        var previewPaths = [];
        for (var i = 0; i < repeatCount; i++) {
            var offset = i * repeatDistancePt;
            var newPosPt = posPt + offset;
            var p;
            if (canvasRadio.value) {
                var size = canvasSize;
                if (horizontalRadio.value) {
                    p = targetLayer.pathItems.add();
                    p.setEntirePath([
                        [-size, -newPosPt],
                        [size, -newPosPt]
                    ]);
                } else {
                    p = targetLayer.pathItems.add();
                    p.setEntirePath([
                        [newPosPt, size],
                        [newPosPt, -size]
                    ]);
                }
            } else {
                var margin = parseFloat(marginEdit.text);
                if (isNaN(margin)) {
                    if (guideLayerRadio.value) {
                        guideLayer.locked = origLock; // 念のため
                    }
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
                    p = targetLayer.pathItems.add();
                    p.setEntirePath([
                        [abLeft - marginPt, abTop - newPosPt],
                        [abRight + marginPt, abTop - newPosPt]
                    ]);
                } else {
                    p = targetLayer.pathItems.add();
                    p.setEntirePath([
                        [abLeft + newPosPt, abTop + marginPt],
                        [abLeft + newPosPt, abBottom - marginPt]
                    ]);
                }
            }
            p.stroked = true;
            p.filled = false;
            p.strokeWidth = 1.0;
            p.strokeColor = makePreviewColor();
            p.guides = false; // プレビューはガイド化しない / Preview is not a guide
            previewPaths.push(p);
        }
        /* シングルプレビュー時は previewLine = path 互換性 / Use single path for single preview */
        if (previewPaths.length === 1) {
            previewLine = previewPaths[0];
        } else {
            previewLine = previewPaths;
        }
        if (guideLayerRadio.value) {
            guideLayer.locked = origLock;
        }
        app.redraw();
    }

    /* プレビュー用カラー（カラーモードに応じて）/ Preview color depending on document color mode */
    function makePreviewColor() {
        var colorSpace = app.activeDocument.documentColorSpace;
        if (colorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 70;
            cmyk.magenta = 50;
            cmyk.yellow = 0;
            cmyk.black = 0;
            return cmyk;
        } else if (colorSpace === DocumentColorSpace.RGB) {
            var rgb = new RGBColor();
            rgb.red = 74;
            rgb.green = 132;
            rgb.blue = 255;
            return rgb;
        } else {
            // その他のカラーモード（念のためRGBに）
            var rgb2 = new RGBColor();
            rgb2.red = 74;
            rgb2.green = 132;
            rgb2.blue = 255;
            return rgb2;
        }
    }

    /* OKボタンクリック時の処理 / OK button click handler */
    okButton.onClick = function() {
        /* プレビュー線がある場合、それをガイド化 / Convert preview line to guide if exists */
        if (previewLine) {
            if (previewLine instanceof Array) {
                for (var i = 0; i < previewLine.length; i++) {
                    var p = previewLine[i];
                    if (p && p.layer && !p.layer.locked) {
                        p.guides = true;
                        p.strokeWidth = 0.1;
                    }
                }
            } else if (previewLine.layer && !previewLine.layer.locked) {
                previewLine.guides = true;
                previewLine.strokeWidth = 0.1;
            }
        }
        // "_guide"レイヤーが存在し、名前が"_guide"ならロック
        if (guideLayer && guideLayer.name === "_guide") {
            guideLayer.locked = true;
        }
        previewLine = null;
        dlg.close();
    };

    /* 対象ラジオボタンの切り替えで裁ち落とし入力の有効/無効切り替え / Enable/disable margin input based on target */
    canvasRadio.onClick = function() {
        marginEdit.enabled = false;
        drawPreviewLine();
    };
    artboardRadio.onClick = function() {
        marginEdit.enabled = true;
        drawPreviewLine();
    };

    /* プレビュー更新イベントを各UIに追加 / Add preview update events to UI elements */
    positionEdit.addEventListener("changing", function() {
        drawPreviewLine();
    });
    marginEdit.addEventListener("changing", function() {
        drawPreviewLine();
    });
    unitDropdown.onChange = function() {
        drawPreviewLine();
    };
    horizontalRadio.onClick = function() {
        drawPreviewLine();
    };
    verticalRadio.onClick = function() {
        drawPreviewLine();
    };
    /* リピート数・距離変更時もプレビュー更新 / Update preview on repeat count/distance change */
    repeatCountEdit.addEventListener("changing", function() {
        drawPreviewLine();
    });
    repeatDistanceEdit.addEventListener("changing", function() {
        drawPreviewLine();
    });

    /* H/Vキーで方向切り替え / Switch direction with H/V keys */
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

    /* ダイアログ表示時に「位置」テキストフィールドにフォーカスを設定 / Focus position input on dialog show */
    positionEdit.active = true;
    /* 初回プレビュー / Initial preview */
    drawPreviewLine();
    /* ダイアログ閉じたらプレビュー線消す / Remove preview line on dialog close */
    dlg.onClose = function() {
        removePreviewLine();
    };
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

        if (event.keyName == "Up" || event.keyName == "Down") {
            var delta = 1;
            if (keyboard.shiftKey) {
                // Shift押下時は「10の倍数」スナップ、制限なし
                value = Math.round(value / 10) * 10 + (event.keyName == "Up" ? 10 : -10);
            } else {
                delta = event.keyName == "Up" ? 1 : -1;
                value += delta;
            }

            event.preventDefault();
            editText.text = value;
            if (typeof onUpdate === "function") {
                onUpdate(editText.text);
            }
        }
    });
}

/* 現在の単位ラベルを取得 / Get current unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    if (unitLabelMap.hasOwnProperty(unitCode)) {
        return unitLabelMap[unitCode];
    } else {
        return "pt";
    }
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