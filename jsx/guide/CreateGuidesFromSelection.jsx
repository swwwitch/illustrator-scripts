#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

/*
### スクリプト名：

CreateGuidesFromSelection.jsx

### 概要

- Illustrator の選択オブジェクトからガイドを作成するスクリプト。
- ダイアログ上で上下左右中央のガイドを自由に指定して描画可能。

### 主な機能

- プレビュー境界または幾何境界の選択
- 一時アウトライン化とアピアランス分割による正確なテキスト処理
- オフセットと裁ち落とし指定
- 「_guide」レイヤー管理とガイド削除オプション
- クリップグループと複数選択対応

### 処理の流れ

- ダイアログボックスでオプションを設定
- 選択オブジェクトの外接矩形を取得
- ガイドを描画
- テキストオブジェクトは一時アウトライン化およびアピアランス展開して計算後復元
- 「_guide」レイヤーにガイドを追加後ロック

### note

https://note.com/dtp_tranist/n/nd1359cf41a2c

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGuidesFromSelection.md

### 更新履歴

- v1.0 (20250711) : 初期バージョン
- v1.1 (20250711) : 複数選択、クリップグループ、日英対応追加
- v1.2 (20250711) : プレビュー境界OFF時の安定化、テキスト処理修正
- v1.3 (20250711) : アートボード外のオブジェクト自動カンバス選択、テキストアウトライン処理改善
- v1.4 (20250711) : アートボード外のオブジェクト選択時のアラート削除、カンバス選択時のはみだし無効化
- v1.5 (20250711) : 左上、左下、右上、右下モードを追加（それぞれ2本のガイドを作成）
- v1.6 (20250712) : コードリファクタリング、ラジオボタンの表示切り替え機能追加
- v1.7 (20250712) : 複数オブジェクトごとにガイドを作成するオプション追加、アートボード基準のガイド作成機能改善
- v1.8 (20250712) : 中心線（垂直のみ）、中心線（水平のみ）の描画ロジックを追加
- v1.9 (20250712) : ↑↓キー、shift + ↑↓キーでの値の増減機能を追加

### Script Name:

CreateGuidesFromSelection.jsx

### Description

- Script to create guides from selected objects in Illustrator.
- Flexible dialog UI to specify top, bottom, left, right, and center guides.

### Main Features

- Use preview or geometric bounds
- Temporary outlining and appearance expansion for text objects
- Offset and bleed (margin) support
- "_guide" layer management and guide removal option
- Supports clip groups and multi-selection

### Workflow

- Configure options in dialog
- Get bounding box of selected objects
- Draw guides
- Temporarily outline and expand appearance for text, then restore
- Add guides to "_guide" layer and lock

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGuidesFromSelection.md

### Update History

- v1.0 (20250711): Initial version
- v1.1 (20250711): Multi-selection & clip group support, offset & bleed features, text outline support
- v1.2 (20250711): Added appearance expansion, UI improvements, enhanced error handling
- v1.3 (20250712): Code refactor and radio button visibility toggle
- v1.6 (20250712): refactored code, added radio button visibility toggle feature
- v1.7 (20250712): Added option to create guides per object, improved artboard-based guide creation
- v1.8 (20250712): Added center line (vertical only) and center line (horizontal only) drawing logic
- v1.9 (20250712): Added up/down arrow key and shift + up/down key value increment/decrement functionality

*/

// --- グローバル定義 / Global definitions ---

// スクリプトバージョン
var SCRIPT_VERSION = "v1.9";


// --- 右側ラジオボタン定義と出し入れ設定 / Add right-side radio buttons in dialog order ---
// Only keep show flags for dynamically shown buttons
var showRbTopBottom = false;
var showRbLeftRight = false;
var showRbTopLeft = true;
var showRbBottomLeft = false;
var showRbTopRight = false;
var showRbBottomRight = false;

// --- UIラベル定義（日本語 / 英語） / UI label definitions (Japanese / English) ---
/*
 * LABELS: UIラベル。UI出現順に定義 / UI labels, defined in dialog appearance order.
 */
var LABELS = {
    centerBothRb: {
        ja: "中心",
        en: "Center"
    },
    centerVRb: {
        ja: "中心線（垂直）",
        en: "Center line (vertical)"
    },
    centerHRb: {
        ja: "中心線（水平）",
        en: "Center line (horizontal)"
    },
    dialogTitle: {
        ja: "ガイド作成ツール " + SCRIPT_VERSION,
        en: "Guide Drawer " + SCRIPT_VERSION
    },
    targetPanel: {
        ja: "対象",
        en: "Target"
    },
    artboard: {
        ja: "アートボード",
        en: "Artboard"
    },
    canvas: {
        ja: "カンバス（擬似）",
        en: "Canvas"
    },
    axisGroup: {
        ja: "辺と中央",
        en: "Axis & Edges"
    },
    left: {
        ja: "左",
        en: "Left"
    },
    top: {
        ja: "上",
        en: "Top"
    },
    center: {
        ja: "中央",
        en: "Center"
    },
    bottom: {
        ja: "下",
        en: "Bottom"
    },
    right: {
        ja: "右",
        en: "Right"
    },
    allOn: {
        ja: "すべて",
        en: "All"
    },
    edges: {
        ja: "四辺",
        en: "Edges"
    },
    vertical: {
        ja: "上下",
        en: "Top & Bottom"
    },
    horizontal: {
        ja: "左右",
        en: "Left & Right"
    },
    topLeft: {
        ja: "左上",
        en: "Top Left"
    },
    bottomLeft: {
        ja: "左下",
        en: "Bottom Left"
    },
    topRight: {
        ja: "右上",
        en: "Top Right"
    },
    bottomRight: {
        ja: "右下",
        en: "Bottom Right"
    },
    clear: {
        ja: "クリア",
        en: "Clear"
    },
    usePreviewBounds: {
        ja: "プレビュー境界を使用",
        en: "Use Preview Bounds"
    },
    margin: {
        ja: "はみ出し",
        en: "Overflow"
    },
    deleteGuides: {
        ja: "「_guide」のガイドを削除",
        en: "Delete guides in \"_guide\""
    },
    offset: {
        ja: "オフセット",
        en: "Offset"
    },
    alertNoSelection: {
        ja: "オブジェクトを選択してください。",
        en: "Please select an object."
    },
    alertExpandError: {
        ja: "アピアランス展開中にエラーが発生しました。",
        en: "An error occurred while expanding appearance."
    },
    alertNoArtboard: {
        ja: "アートボードが存在しません。",
        en: "No artboard exists."
    },
    alertInvalidArtboard: {
        ja: "有効なアートボードが選択されていません。",
        en: "No valid artboard selected."
    },
    alertDeleteGuideError: {
        ja: "既存ガイド削除時にエラーが発生しました。",
        en: "An error occurred while deleting existing guides."
    },
    alertGuideError: {
        ja: "ガイド作成中にエラーが発生しました。",
        en: "An error occurred while creating guides."
    },
    okButton: {
        ja: "ガイドを描画",
        en: "Draw Guides"
    },
    cancelButton: {
        ja: "キャンセル",
        en: "Cancel"
    },
    individual: {
        ja: "オブジェクトごとにガイドを作成",
        en: "Create guides per object"
    }
};

// --- 単位コードとラベルのマップ / Unit code and label map ---
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/**
 * 現在の単位ラベルを取得
 * Get current ruler unit label
 */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

var canvasSize = 227 * 72; // カンバスサイズ（pt）/ canvas size (pt)

/**
 * _guideレイヤーを取得または作成
 * Get or create "_guide" layer.
 */
function getOrCreateGuideLayer() {
    var doc = app.activeDocument;
    var layer = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "_guide") {
            layer = doc.layers[i];
            break;
        }
    }
    if (!layer) {
        layer = doc.layers.add();
        layer.name = "_guide";
    }
    return layer;
}

/**
 * 選択オブジェクトからガイドを作成する
 * Create guides from selected objects.
 * @param {Object} options - ガイド描画の指定 / Guide drawing options
 * @param {boolean} useCanvas - カンバス基準か / Use canvas as reference
 * @param {number} offsetValue - オフセット値 / Offset value
 * @param {number} marginValue - 裁ち落とし値 / Bleed (margin) value
 */
function createGuidesFromSelection(options, useCanvas, offsetValue, marginValue) {
    var doc = app.activeDocument;
    var selItems = app.selection;
    var bounds;
    var isUsingArtboard = false;

    // --- Per-object guides block ---
    if (options.individual && selItems.length > 0) {
        var layer = getOrCreateGuideLayer();
        var wasLocked = layer.locked;
        if (wasLocked) layer.locked = false;

        for (var i = 0; i < selItems.length; i++) {
            var item = selItems[i];
            var b = options.usePreviewBounds ? item.visibleBounds : item.geometricBounds;

            var top = b[1] + offsetValue;
            var left = b[0] - offsetValue;
            var bottom = b[3] - offsetValue;
            var right = b[2] + offsetValue;
            var centerX = (left + right) / 2;
            var centerY = (top + bottom) / 2;

            var directions = [{
                    flag: options.left,
                    pos: left,
                    orientation: "vertical"
                },
                {
                    flag: options.right,
                    pos: right,
                    orientation: "vertical"
                },
                {
                    flag: options.top,
                    pos: top,
                    orientation: "horizontal"
                },
                {
                    flag: options.bottom,
                    pos: bottom,
                    orientation: "horizontal"
                }
            ];
            // Add center guides based on consistent logic
            if (options.center) {
                directions.push({
                    flag: true,
                    pos: centerX,
                    orientation: "vertical"
                });
                directions.push({
                    flag: true,
                    pos: centerY,
                    orientation: "horizontal"
                });
            }
            if (options.centerMode === "vertical") {
                directions.push({
                    flag: true,
                    pos: centerX,
                    orientation: "vertical"
                });
            }
            if (options.centerMode === "horizontal") {
                directions.push({
                    flag: true,
                    pos: centerY,
                    orientation: "horizontal"
                });
            }
            for (var d = 0; d < directions.length; d++) {
                if (directions[d].flag) {
                    createGuide(layer, directions[d].pos, directions[d].orientation, useCanvas, marginValue);
                }
            }
        }
        layer.locked = true;
        return;
    }
    // --- End per-object guides block ---

    if (selItems.length === 0) {
        // 選択がない場合はアートボード基準
        if (doc.artboards.length === 0) {
            alert(LABELS.alertNoArtboard);
            return;
        }
        var abIndex = doc.artboards.getActiveArtboardIndex();
        if (abIndex < 0 || abIndex >= doc.artboards.length) {
            alert(LABELS.alertInvalidArtboard);
            return;
        }
        var ab = doc.artboards[abIndex].artboardRect;
        bounds = [ab[0], ab[1], ab[2], ab[3]];
        isUsingArtboard = true;
    } else {
        // --- テキストオブジェクトのアウトライン化（プレビュー境界時のみ）およびアピアランス展開 / Outline and expand appearance if using preview bounds ---
        var textCopies = [];
        var originalTexts = [];
        if (options.usePreviewBounds) {
            var tempCopies = [];
            for (var i = 0; i < selItems.length; i++) {
                var item = selItems[i];
                if (item && item.typename === "TextFrame") {
                    var copy = item.duplicate();
                    item.hidden = true;
                    tempCopies.push(copy);
                    originalTexts.push(item);
                }
            }
            if (tempCopies.length > 0) {
                app.selection = null;
                for (var j = 0; j < tempCopies.length; j++) {
                    tempCopies[j].selected = true;
                }
                try {
                    app.executeMenuCommand('expandStyle');
                } catch (e) {
                    alert(LABELS.alertExpandError + "\n" + e.message);
                }
                for (var k = 0; k < tempCopies.length; k++) {
                    var outlined = tempCopies[k].createOutline();
                    if (outlined) {
                        textCopies.push(outlined);
                    } else {
                        textCopies.push(tempCopies[k]);
                    }
                }
                selItems = textCopies.length > 0 ? textCopies : selItems;
            }
        }

        var first = true;
        for (var i = 0; i < selItems.length; i++) {
            var item = selItems[i];
            var b;
            if (item.typename === "GroupItem" && item.clipped) {
                var mask = null;
                for (var j = 0; j < item.pageItems.length; j++) {
                    if (item.pageItems[j].clipping) {
                        mask = item.pageItems[j];
                        break;
                    }
                }
                if (mask) {
                    b = options.usePreviewBounds ? mask.visibleBounds : mask.geometricBounds;
                } else {
                    b = options.usePreviewBounds ? item.visibleBounds : item.geometricBounds;
                }
            } else {
                b = options.usePreviewBounds ? item.visibleBounds : item.geometricBounds;
            }
            if (first) {
                bounds = b.concat();
                first = false;
            } else {
                bounds[0] = Math.min(bounds[0], b[0]);
                bounds[1] = Math.max(bounds[1], b[1]);
                bounds[2] = Math.max(bounds[2], b[2]);
                bounds[3] = Math.min(bounds[3], b[3]);
            }
        }
    }

    var top = bounds[1] + offsetValue;
    var left = bounds[0] - offsetValue;
    var bottom = bounds[3] - offsetValue;
    var right = bounds[2] + offsetValue;
    var centerX = (left + right) / 2;
    var centerY = (top + bottom) / 2;

    var layer = getOrCreateGuideLayer();
    var wasLocked = layer.locked;
    if (wasLocked) layer.locked = false;

    var directions = [{
            flag: options.left,
            pos: left,
            orientation: "vertical"
        },
        {
            flag: options.right,
            pos: right,
            orientation: "vertical"
        },
        {
            flag: options.top,
            pos: top,
            orientation: "horizontal"
        },
        {
            flag: options.bottom,
            pos: bottom,
            orientation: "horizontal"
        }
    ];
    // Add center guides based on consistent logic
    if (options.center) {
        directions.push({
            flag: true,
            pos: centerX,
            orientation: "vertical"
        });
        directions.push({
            flag: true,
            pos: centerY,
            orientation: "horizontal"
        });
    }
    if (options.centerMode === "vertical") {
        directions.push({
            flag: true,
            pos: centerX,
            orientation: "vertical"
        });
    }
    if (options.centerMode === "horizontal") {
        directions.push({
            flag: true,
            pos: centerY,
            orientation: "horizontal"
        });
    }

    for (var i = 0; i < directions.length; i++) {
        if (directions[i].flag) {
            createGuide(layer, directions[i].pos, directions[i].orientation, useCanvas, marginValue);
        }
    }

    if (selItems.length !== 0 && textCopies && textCopies.length > 0) {
        for (var i = 0; i < textCopies.length; i++) {
            textCopies[i].remove();
            originalTexts[i].hidden = false;
        }
    }
    layer.locked = true;
}

/**
 * ガイド1本を作成
 * Create a single guide path.
 */
function createGuide(layer, pos, orientation, useCanvas, marginValue) {
    var doc = app.activeDocument;
    var guide;
    if (useCanvas) {
        // カンバス基準：±8000 pt固定 / Use ±8000pt for canvas
        var big = 8000;
        if (orientation === "horizontal") {
            guide = layer.pathItems.add();
            guide.setEntirePath([
                [-big, pos],
                [big, pos]
            ]);
        } else {
            guide = layer.pathItems.add();
            guide.setEntirePath([
                [pos, big],
                [pos, -big]
            ]);
        }
    } else {
        // アートボード基準 / Use artboard bounds
        if (doc.artboards.length === 0) {
            alert(LABELS.alertNoArtboard);
            return;
        }
        var abIndex = doc.artboards.getActiveArtboardIndex();
        if (abIndex < 0 || abIndex >= doc.artboards.length) {
            alert(LABELS.alertInvalidArtboard);
            return;
        }
        var ab = doc.artboards[abIndex].artboardRect;
        if (orientation === "horizontal") {
            guide = layer.pathItems.add();
            guide.setEntirePath([
                [ab[0] - marginValue, pos],
                [ab[2] + marginValue, pos]
            ]);
        } else {
            guide = layer.pathItems.add();
            guide.setEntirePath([
                [pos, ab[1] + marginValue],
                [pos, ab[3] - marginValue]
            ]);
        }
    }
    // Remove fill and stroke before converting to a guide
    guide.filled = false;
    guide.stroked = false;
    guide.guides = true;
}

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

/**
 * メインダイアログを構築・表示 / Build and show the main dialog
 */
function buildDialog() {
    // --- Declare global center mode variable ---
    var centerModeGlobal = "";
    var dialog = new Window("dialog");
    dialog.text = LABELS.dialogTitle;
    dialog.orientation = "column";
    dialog.alignChildren = ["center", "top"];
    dialog.spacing = 10;
    dialog.margins = [15, 15, 35, 15];

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["fill", "top"];
    mainGroup.spacing = 20;
    mainGroup.margins = 0;

    var leftGroup = mainGroup.add("group");
    leftGroup.orientation = "column";
    leftGroup.alignChildren = ["left", "top"];
    leftGroup.spacing = 10;

    var rightGroup = mainGroup.add("group");
    rightGroup.orientation = "column";
    rightGroup.alignChildren = ["left", "top"];
    rightGroup.spacing = 10;
    rightGroup.alignment = ["left", "center"];

    var targetPanel = leftGroup.add("panel", undefined, LABELS.targetPanel);
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["left", "top"];
    targetPanel.spacing = 10;
    targetPanel.margins = [10, 20, 20, 15];
    // カンバス→アートボードの順にラジオボタンを追加し、デフォルトをアートボードに
    var rbCanvas = targetPanel.add("radiobutton", undefined, LABELS.canvas);
    var rbArtboard = targetPanel.add("radiobutton", undefined, LABELS.artboard);
    rbArtboard.value = true;
    // --- 「裁ち落とし」グループを targetPanel 内に追加 / Add margin (bleed) group to targetPanel ---
    var marginGroup = targetPanel.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = ["left", "center"];
    marginGroup.add("statictext", undefined, LABELS.margin);
    var marginInput = marginGroup.add("edittext", undefined, "20");
    marginInput.characters = 3;
    marginGroup.add("statictext", undefined, getCurrentUnitLabel());
    changeValueByArrowKey(marginInput);

    var axisGroup = leftGroup.add("panel", undefined, undefined, {
        name: "axisGroup"
    });
    axisGroup.text = LABELS.axisGroup;
    axisGroup.orientation = "column";
    axisGroup.alignChildren = ["left", "top"];
    axisGroup.spacing = 10;
    axisGroup.margins = [10, 20, 10, 5];

    var diamondGroup = axisGroup.add("group", undefined, {
        name: "diamondGroup"
    });
    diamondGroup.orientation = "row";
    diamondGroup.alignChildren = ["left", "center"];
    diamondGroup.spacing = 20;
    diamondGroup.margins = 0;
    diamondGroup.alignment = ["fill", "top"];

    var colLeft = diamondGroup.add("group", undefined, {
        name: "colLeft"
    });
    colLeft.orientation = "row";
    colLeft.alignChildren = ["left", "center"];
    colLeft.spacing = 10;
    colLeft.margins = 0;
    var cbLeft = colLeft.add("checkbox", undefined, undefined, {
        name: "cbLeft"
    });
    cbLeft.text = LABELS.left;
    cbLeft.value = true;

    var colCenter = diamondGroup.add("group", undefined, {
        name: "colCenter"
    });

    colCenter.orientation = "column";
    colCenter.alignChildren = ["left", "center"];
    colCenter.spacing = 10;
    colCenter.margins = 0;
    var cbTop = colCenter.add("checkbox", undefined, undefined, {
        name: "cbTop"
    });
    cbTop.text = LABELS.top;
    cbTop.value = true;
    var cbCenter = colCenter.add("checkbox", undefined, undefined, {
        name: "cbCenter"
    });
    cbCenter.text = LABELS.center;
    cbCenter.value = false;
    // for horizontal center line
    var cbBottom = colCenter.add("checkbox", undefined, undefined, {
        name: "cbBottom"
    });
    cbBottom.text = LABELS.bottom;
    cbBottom.value = true;

    var colRight = diamondGroup.add("group", undefined, {
        name: "colRight"
    });
    colRight.orientation = "row";
    colRight.alignChildren = ["left", "center"];
    colRight.spacing = 10;
    colRight.margins = 0;
    var cbRight = colRight.add("checkbox", undefined, undefined, {
        name: "cbRight"
    });
    cbRight.text = LABELS.right;
    cbRight.value = true;

    var edgeGroup = axisGroup.add("group", undefined, {
        name: "edgeGroup"
    });
    edgeGroup.orientation = "row";
    edgeGroup.alignChildren = ["left", "center"];
    edgeGroup.spacing = 10;
    edgeGroup.margins = [0, 15, 0, 5];
    edgeGroup.alignment = ["center", "top"];



    function createRadioButton(group, label, show) {
        if (typeof show !== "undefined" && !show) return null;
        return group.add("radiobutton", undefined, label);
    }

    // Always-shown buttons: no show argument, no flags
    var rbAllOn = createRadioButton(rightGroup, LABELS.allOn);
    var rbShihen = createRadioButton(rightGroup, LABELS.edges);
    var rbTopBottom = createRadioButton(rightGroup, LABELS.vertical, showRbTopBottom);
    var rbLeftRight = createRadioButton(rightGroup, LABELS.horizontal, showRbLeftRight);
    var rbTopLeft = createRadioButton(rightGroup, LABELS.topLeft, showRbTopLeft);
    var rbBottomLeft = createRadioButton(rightGroup, LABELS.bottomLeft, showRbBottomLeft);
    var rbTopRight = createRadioButton(rightGroup, LABELS.topRight, showRbTopRight);
    var rbBottomRight = createRadioButton(rightGroup, LABELS.bottomRight, showRbBottomRight);
    var rbCenterBoth = createRadioButton(rightGroup, LABELS.centerBothRb);
    var rbCenterV = createRadioButton(rightGroup, LABELS.centerVRb);
    var rbCenterH = createRadioButton(rightGroup, LABELS.centerHRb);
    var rbClear = createRadioButton(rightGroup, LABELS.clear);
    // --- 中心ラジオボタンのイベントハンドラ ---
    if (rbCenterBoth) {
        rbCenterBoth.onClick = function() {
            if (rbCenterBoth.value) setCheckboxState(false, false, false, false, true);
        };
    }

    // --- Insert event handlers for center line radio buttons ---
    if (rbCenterV) {
        rbCenterV.onClick = function() {
            if (rbCenterV.value) {
                setCheckboxState(false, false, false, false, false);
                centerModeGlobal = "vertical";
            }
        };
    }
    if (rbCenterH) {
        rbCenterH.onClick = function() {
            if (rbCenterH.value) {
                setCheckboxState(false, false, false, false, false);
                centerModeGlobal = "horizontal";
            }
        };
    }

    // デフォルト選択は表示中の最初のボタンに自動設定（Shihen優先） / Default selection (prefer "Edges")
    if (rbShihen) {
        rbShihen.value = true;
    } else if (rbAllOn) {
        rbAllOn.value = true;
    }

    // --- チェックボックス一括設定関数 / Function to set checkboxes at once ---
    // チェックボックス群の状態をまとめてセット / Set checkboxes easily
    function setCheckboxState(l, t, r, b, c) {
        cbLeft.value = l;
        cbTop.value = t;
        cbRight.value = r;
        cbBottom.value = b;
        cbCenter.value = c;
    }

    if (rbShihen) {
        rbShihen.onClick = function() {
            if (rbShihen.value) setCheckboxState(true, true, true, true, false);
        };
    }
    if (rbAllOn) {
        rbAllOn.onClick = function() {
            if (rbAllOn.value) setCheckboxState(true, true, true, true, true);
        };
    }
    if (rbTopBottom) {
        rbTopBottom.onClick = function() {
            if (rbTopBottom.value) setCheckboxState(false, true, false, true, false);
        };
    }
    if (rbLeftRight) {
        rbLeftRight.onClick = function() {
            if (rbLeftRight.value) setCheckboxState(true, false, true, false, false);
        };
    }
    if (rbTopLeft) {
        rbTopLeft.onClick = function() {
            if (rbTopLeft.value) setCheckboxState(true, true, false, false, false);
        };
    }
    if (rbBottomLeft) {
        rbBottomLeft.onClick = function() {
            if (rbBottomLeft.value) setCheckboxState(true, false, false, true, false);
        };
    }
    if (rbTopRight) {
        rbTopRight.onClick = function() {
            if (rbTopRight.value) setCheckboxState(false, true, true, false, false);
        };
    }
    if (rbBottomRight) {
        rbBottomRight.onClick = function() {
            if (rbBottomRight.value) setCheckboxState(false, false, true, true, false);
        };
    }
    if (rbClear) {
        rbClear.onClick = function() {
            if (rbClear.value) setCheckboxState(false, false, false, false, false);
        };
    }

    var optionsGroup = dialog.add("group", undefined, {
        name: "optionsGroup"
    });
    optionsGroup.orientation = "column";
    optionsGroup.alignChildren = ["left", "center"];
    optionsGroup.margins = [10, 15, 10, 20];
    optionsGroup.spacing = 10;

    var cbUsePreview = optionsGroup.add("checkbox", undefined, LABELS.usePreviewBounds);
    cbUsePreview.value = true;
    var cbDeleteGuide = optionsGroup.add("checkbox", undefined, LABELS.deleteGuides);
    cbDeleteGuide.value = true;

    // --- 個別ガイドチェックボックス追加 / Add individual guide checkbox ---
    var cbIndividual = optionsGroup.add("checkbox", undefined, LABELS.individual);
    cbIndividual.value = false;

    // --- オフセット入力欄追加 / Add offset input field ---
    var offsetGroup = optionsGroup.add("group");
    offsetGroup.orientation = "row";
    offsetGroup.alignChildren = ["left", "center"];
    offsetGroup.add("statictext", undefined, LABELS.offset);
    var offsetInput = offsetGroup.add("edittext", undefined, "0");
    offsetInput.characters = 3;
    offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
    offsetInput.active = true;
    changeValueByArrowKey(offsetInput);
    // --- 「裁ち落とし」(marginInput) を「カンバス」選択時にディム表示する制御を追加 / Disable margin input if canvas is selected ---
    function updateMarginEnabled() {
        if (rbCanvas.value) {
            marginInput.enabled = false;
        } else {
            marginInput.enabled = true;
        }
    }
    rbArtboard.onClick = updateMarginEnabled;
    rbCanvas.onClick = updateMarginEnabled;
    // 初期状態に合わせる
    updateMarginEnabled();

    // --- 選択オブジェクトがアートボード外にある場合、自動的にカンバス選択 ---
    var doc = app.activeDocument;
    var selItems = app.selection;
    if (selItems.length > 0 && doc.artboards.length > 0) {
        var abIndex = doc.artboards.getActiveArtboardIndex();
        var ab = doc.artboards[abIndex].artboardRect;
        var abLeft = ab[0];
        var abTop = ab[1];
        var abRight = ab[2];
        var abBottom = ab[3];

        var allOutside = true;

        for (var i = 0; i < selItems.length; i++) {
            var item = selItems[i];
            var b = item.geometricBounds;
            if (!(b[0] > abRight || b[2] < abLeft || b[1] < abBottom || b[3] > abTop)) {
                allOutside = false;
                break;
            }
        }

        if (allOutside) {
            rbCanvas.value = true;
            rbArtboard.value = false;
            updateMarginEnabled();
        }
    }

    var btnGroup = dialog.add("group");
    btnGroup.alignment = ["center", "top"];
    var btnCancel = btnGroup.add("button", undefined, LABELS.cancelButton);
    var btnCreateGuides = btnGroup.add("button", undefined, LABELS.okButton, {
        name: "ok"
    });

    btnCancel.onClick = function() {
        dialog.close();
    };

    btnCreateGuides.onClick = function() {
        try {
            var options = {
                left: cbLeft.value,
                right: cbRight.value,
                top: cbTop.value,
                bottom: cbBottom.value,
                center: cbCenter.value,
                centerMode: centerModeGlobal,
                usePreviewBounds: cbUsePreview.value,
                individual: cbIndividual.value
            };
            var useCanvas = rbCanvas.value;

            // --- _guideレイヤー取得または作成 / Get or create "_guide" layer ---
            var layer = getOrCreateGuideLayer();
            var wasLocked = layer.locked;
            if (wasLocked) layer.locked = false;

            // --- 削除チェックON時、既存ガイド削除 / Remove existing guides if checked ---
            if (cbDeleteGuide.value) {
                try {
                    for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                        if (layer.pageItems[i].guides) {
                            layer.pageItems[i].remove();
                        }
                    }
                } catch (ex) {
                    alert(LABELS.alertDeleteGuideError + "\n" + ex.message);
                }
            }

            var offsetVal = parseFloat(offsetInput.text);
            if (isNaN(offsetVal)) offsetVal = 0;
            var marginVal = parseFloat(marginInput.text);
            if (isNaN(marginVal)) marginVal = 0;
            var unitCode = app.preferences.getIntegerPreference("rulerType");
            var ptFactor = getPtFactorFromUnitCode(unitCode);
            var offsetValPt = offsetVal * ptFactor;
            var marginValPt = marginVal * ptFactor;

            createGuidesFromSelection(options, useCanvas, offsetValPt, marginValPt);
            dialog.close();
        } catch (e) {
            alert(LABELS.alertGuideError + "\n" + (e && e.message ? e.message : e) + "\n" + (e && e.stack ? e.stack : ""));
        }
    };

    dialog.show();
}

buildDialog();

/**
 * テキストフィールドで上下矢印キーによる数値増減を可能にする
 * Enable value change with up/down arrow keys in an edittext
 */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            if (keyboard.shiftKey) {
                // Shift押下時は10の倍数スナップ
                value = Math.round(value / 10) * 10 + (event.keyName == "Up" ? 10 : -10);
            } else {
                var delta = event.keyName == "Up" ? 1 : -1;
                value += delta;
            }

            event.preventDefault();
            editText.text = value;
        }
    });
}