#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

CreateGuidesFromSelection

### 概要

- 選択したオブジェクトの中央、エッジ、左右のみ、上下のみを基準にガイドを作成します。
- プレビュー境界、はみだし、マージン、ガイド削除、クリップグループやテキストアウトライン化に対応。

### 主な機能

- モード切替（中央、エッジ、左右のみ、上下のみ）
- プレビュー境界使用オプション
- はみだし、マージン設定
- 「_guide」レイヤー内のガイド削除
- テキストアウトライン処理と復元
- 多言語対応（日本語・英語）

### 処理の流れ

1. ダイアログでオプション選択
2. 境界座標を取得
3. 指定モードでガイド作成

### 更新履歴

- v1.0 (20250711) : 初期バージョン
- v1.1 (20250711) : 複数選択、クリップグループ、日英対応追加
- v1.2 (20250711) : はみだし、マージン設定、左右/上下モード、ガイド削除機能追加

---

### Script Name:

CreateGuidesFromSelection

### Overview

- Creates guides based on center, edge, sides only, or top/bottom only of selected objects.
- Supports visible bounds, overflow, margin, clearing "_guide" layer, clip groups, and text outline handling.

### Features

- Mode switching (Center, Edge, Sides Only, Top/Bottom Only)
- Use visible bounds option
- Overflow and margin settings
- Clear guides in "_guide" layer
- Temporary text outline and restore
- Multi-language support (Japanese/English)

### Workflow

1. Choose options in dialog
2. Get object bounds
3. Create guides based on selected mode

### Update History

- v1.0 (20250711): Initial version
- v1.1 (20250711): Multiple selection, clip group, language support
- v1.2 (20250711): Added overflow & margin, sides/top-bottom modes, clear guides option
*/

var SCRIPT_VERSION = "v1.2";

// 言語取得 / Get current language
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// ラベル定義 / UI Label Definitions (ordered by UI appearance)
var lang = getCurrentLang();
var LABELS = {
    dialogTitle: {
        ja: "ガイド作成 " + SCRIPT_VERSION,
        en: "Create Guides " + SCRIPT_VERSION
    },
    panelTitle: {
        ja: "対象",
        en: "Target"
    },
    center: {
        ja: "中央",
        en: "Center"
    },
    edge: {
        ja: "エッジ",
        en: "Edge"
    },
    sidesOnly: {
        ja: "左右のみ",
        en: "Sides Only"
    },
    topBottomOnly: {
        ja: "上下のみ",
        en: "Top/Bottom Only"
    },
    useVisible: {
        ja: "プレビュー境界を使用",
        en: "Use visible bounds"
    },
    clearGuides: {
        ja: "「_guide」のガイドを削除",
        en: "Clear guides in '_guide' layer"
    },
    offset: {
        ja: "はみだし:",
        en: "Overflow:"
    },
    margin: {
        ja: "マージン:",
        en: "Margin:"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    alertSelect: {
        ja: "オブジェクトを選択してください。",
        en: "Please select an object."
    },
    error: {
        ja: "エラーが発生しました: ",
        en: "An error occurred: "
    }
};

// 単位コードとラベルのマップ / Unit code to label map
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

// 現在の単位ラベル取得 / Get current unit label
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

// 単位コードからpt変換係数取得 / Get pt factor from unit code
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0;
        case 1:
            return 72.0 / 25.4;
        case 2:
            return 1.0;
        case 3:
            return 12.0;
        case 4:
            return 72.0 / 2.54;
        case 5:
            return 72.0 / 25.4 * 0.25;
        case 6:
            return 1.0;
        case 7:
            return 72.0 * 12.0;
        case 8:
            return 72.0 / 25.4 * 1000.0;
        case 9:
            return 72.0 * 36.0;
        case 10:
            return 72.0 * 12.0;
        default:
            return 1.0;
    }
}

// bounds 計算用関数 / Function to calculate bounds
function calculateBounds(items, useVisible) {
    var bounds = useVisible ? items[0].visibleBounds.concat() : items[0].geometricBounds.concat();
    for (var i = 1; i < items.length; i++) {
        var itemBounds = useVisible ? items[i].visibleBounds : items[i].geometricBounds;
        if (itemBounds[0] < bounds[0]) bounds[0] = itemBounds[0];
        if (itemBounds[1] > bounds[1]) bounds[1] = itemBounds[1];
        if (itemBounds[2] > bounds[2]) bounds[2] = itemBounds[2];
        if (itemBounds[3] < bounds[3]) bounds[3] = itemBounds[3];
    }
    return bounds;
}

// 共通定義まとめ / Common definitions
var unitCode = app.preferences.getIntegerPreference("rulerType");
var factor = getPtFactorFromUnitCode(unitCode);
var selectionItems = app.selection;

// 中央基準ガイド作成 / Create center guides
function createCenterGuides(bounds, artboardRect, guideLayer) {
    var centerX = bounds[0] + ((bounds[2] - bounds[0]) / 2);
    var centerY = bounds[1] - ((bounds[1] - bounds[3]) / 2);

    var vGuide = guideLayer.pathItems.add();
    vGuide.setEntirePath([[centerX, artboardRect[1]], [centerX, artboardRect[3]]]);
    vGuide.guides = true;

    var hGuide = guideLayer.pathItems.add();
    hGuide.setEntirePath([[artboardRect[0], centerY], [artboardRect[2], centerY]]);
    hGuide.guides = true;
}

// エッジ基準ガイド作成 / Create edge guides
function createEdgeGuides(bounds, artboardRect, guideLayer, offsetValue, marginValue, mode) {
    var leftX = bounds[0] - marginValue;
    var rightX = bounds[2] + marginValue;
    var topY = bounds[1] + marginValue;
    var bottomY = bounds[3] - marginValue;

    var guideLeft = artboardRect[0] - offsetValue;
    var guideRight = artboardRect[2] + offsetValue;
    var guideTop = artboardRect[1] + offsetValue;
    var guideBottom = artboardRect[3] - offsetValue;

    if (mode === "full" || mode === "sides") {
        var leftGuide = guideLayer.pathItems.add();
        leftGuide.setEntirePath([[leftX, guideTop], [leftX, guideBottom]]);
        leftGuide.guides = true;

        var rightGuide = guideLayer.pathItems.add();
        rightGuide.setEntirePath([[rightX, guideTop], [rightX, guideBottom]]);
        rightGuide.guides = true;
    }

    if (mode === "full" || mode === "topBottom") {
        var topGuide = guideLayer.pathItems.add();
        topGuide.setEntirePath([[guideLeft, topY], [guideRight, topY]]);
        topGuide.guides = true;

        var bottomGuide = guideLayer.pathItems.add();
        bottomGuide.setEntirePath([[guideLeft, bottomY], [guideRight, bottomY]]);
        bottomGuide.guides = true;
    }
}

function main() {
    try {
        var isValid = true;
        var doc = app.activeDocument;
        var artboards = doc.artboards;
        var activeIndex = artboards.getActiveArtboardIndex();
        var artboardRect = artboards[activeIndex].artboardRect;

        // 「_guide」レイヤーを取得または作成 / Get or create "_guide" layer
        var guideLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "_guide") {
                guideLayer = doc.layers[i];
                break;
            }
        }
        if (!guideLayer) {
            guideLayer = doc.layers.add();
            guideLayer.name = "_guide";
        } else {
            guideLayer.locked = false;
        }

        if (selectionItems.length < 1) {
            alert(LABELS.alertSelect[lang]);
            isValid = false;
        }

        var rbCenter, rbEdge, rbSidesOnly, rbTopBottomOnly, cbUseVisible, cbClearGuides, offsetInput, marginInput;

        if (isValid) {
            // ダイアログ作成 / Create dialog
            var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
            dlg.orientation = "column";
            dlg.alignChildren = "left";

            var modePanel = dlg.add("panel", undefined, LABELS.panelTitle[lang]);
            modePanel.orientation = "row";
            modePanel.alignChildren = "top";
            modePanel.margins = [15, 20, 15, 10];

            var col1 = modePanel.add("group");
            col1.orientation = "column";
            col1.alignChildren = "left";

            var col2 = modePanel.add("group");
            col2.orientation = "column";
            col2.alignChildren = "left";

            rbCenter = col1.add("radiobutton", undefined, LABELS.center[lang]);
            rbEdge = col1.add("radiobutton", undefined, LABELS.edge[lang]);
            rbSidesOnly = col2.add("radiobutton", undefined, LABELS.sidesOnly[lang]);
            rbTopBottomOnly = col2.add("radiobutton", undefined, LABELS.topBottomOnly[lang]);
            rbCenter.value = false;
            rbEdge.value = true;

            // 排他的動作にするための onClick イベント
            rbCenter.onClick = function () {
                rbEdge.value = false;
                rbSidesOnly.value = false;
                rbTopBottomOnly.value = false;
            };
            rbEdge.onClick = function () {
                rbCenter.value = false;
                rbSidesOnly.value = false;
                rbTopBottomOnly.value = false;
            };
            rbSidesOnly.onClick = function () {
                rbCenter.value = false;
                rbEdge.value = false;
                rbTopBottomOnly.value = false;
            };
            rbTopBottomOnly.onClick = function () {
                rbCenter.value = false;
                rbEdge.value = false;
                rbSidesOnly.value = false;
            };

            cbUseVisible = dlg.add("checkbox", undefined, LABELS.useVisible[lang]);
            cbUseVisible.value = true;

            cbClearGuides = dlg.add("checkbox", undefined, LABELS.clearGuides[lang]);
            cbClearGuides.value = true;

            var offsetGroup = dlg.add("group");
            offsetGroup.add("statictext", undefined, LABELS.offset[lang]);
            offsetInput = offsetGroup.add("edittext", undefined, "6");
            offsetInput.characters = 3;
            offsetGroup.add("statictext", undefined, getCurrentUnitLabel());

            var marginGroup = dlg.add("group");
            marginGroup.add("statictext", undefined, LABELS.margin[lang]);
            marginInput = marginGroup.add("edittext", undefined, "0");
            marginInput.characters = 3;
            marginGroup.add("statictext", undefined, getCurrentUnitLabel());

            var btnGroup = dlg.add("group");
            btnGroup.alignment = "right";
            var cancelBtn = btnGroup.add("button", undefined, LABELS.cancel[lang]);
            var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang]);

            if (dlg.show() != 1) {
                isValid = false;
            }
        }

        if (!isValid) {
            guideLayer.locked = true;
            return;
        }

        if (cbClearGuides.value) {
            for (var i = guideLayer.pageItems.length - 1; i >= 0; i--) {
                guideLayer.pageItems[i].remove();
            }
        }

        var offsetValue = parseFloat(offsetInput.text);
        if (isNaN(offsetValue)) offsetValue = 0;
        offsetValue *= factor;

        var marginValue = parseFloat(marginInput.text);
        if (isNaN(marginValue)) marginValue = 0;
        marginValue *= factor;

        // テキストオブジェクトを事前に複製し、hidden 配列で管理 / Duplicate text frames and hide originals
        var textCopies = [];
        var originalTexts = [];
        for (var i = 0; i < selectionItems.length; i++) {
            var item = selectionItems[i];
            if (item.typename === "TextFrame") {
                var copy = item.duplicate();
                item.hidden = true;
                var outlined = copy.createOutline();
                textCopies.push(outlined);
                originalTexts.push(item);
            }
        }

        // クリップグループのマスクパス取得関数 / Get clipping mask path from group
        function getClipMask(groupItem) {
            for (var i = 0; i < groupItem.pageItems.length; i++) {
                if (groupItem.pageItems[i].clipping) {
                    return groupItem.pageItems[i];
                }
            }
            return null;
        }

        // itemsForBounds 構築処理の関数化 / Build items array for bounds calculation
        function buildItemsForBounds(selectionItems, textCopies) {
            var items = [];
            for (var i = 0; i < selectionItems.length; i++) {
                var item = selectionItems[i];
                if (item.typename !== "TextFrame") {
                    if (item.typename === "GroupItem" && item.clipped) {
                        var mask = getClipMask(item);
                        items.push(mask ? mask : item);
                    } else {
                        items.push(item);
                    }
                }
            }
            for (var i = 0; i < textCopies.length; i++) {
                items.push(textCopies[i]);
            }
            return items;
        }

        // bounds 計算用の itemsForBounds を構築 / Build items for bounds calculation
        var itemsForBounds = buildItemsForBounds(selectionItems, textCopies);
        // bounds 計算関数を使用 / Calculate bounds
        var bounds = calculateBounds(itemsForBounds, cbUseVisible.value);

        // オブジェクトがアートボード外にある場合、アラートを出す / Alert if objects are outside artboard
        if (
            bounds[2] < artboardRect[0] || // 右端がアートボード左より左 / right edge left of artboard left
            bounds[0] > artboardRect[2] || // 左端がアートボード右より右 / left edge right of artboard right
            bounds[1] < artboardRect[3] || // 上端がアートボード下より下 / top edge below artboard bottom
            bounds[3] > artboardRect[1]    // 下端がアートボード上より上 / bottom edge above artboard top
        ) {
            alert("オブジェクトがアートボード外にあります。");
            guideLayer.locked = true;
            // 元のテキストを再表示 / Restore original text visibility
            for (var i = 0; i < originalTexts.length; i++) {
                originalTexts[i].hidden = false;
            }
            // アウトライン化した複製オブジェクトを削除 / Remove outlined duplicates
            for (var i = 0; i < textCopies.length; i++) {
                textCopies[i].remove();
            }
            return;
        }

        if (rbCenter.value) {
            createCenterGuides(bounds, artboardRect, guideLayer);
        } else if (rbEdge.value) {
            createEdgeGuides(bounds, artboardRect, guideLayer, offsetValue, marginValue, "full");
        } else if (rbSidesOnly.value) {
            createEdgeGuides(bounds, artboardRect, guideLayer, offsetValue, marginValue, "sides");
        } else if (rbTopBottomOnly.value) {
            createEdgeGuides(bounds, artboardRect, guideLayer, offsetValue, marginValue, "topBottom");
        }
        guideLayer.locked = true;

        // ガイド作成後、元のテキストを再表示 / Restore original text visibility
        for (var i = 0; i < originalTexts.length; i++) {
            originalTexts[i].hidden = false;
        }
        // アウトライン化した複製オブジェクトを削除 / Remove outlined duplicates
        for (var i = 0; i < textCopies.length; i++) {
            textCopies[i].remove();
        }

    } catch (e) {
        alert(LABELS.error[lang] + e);
    }
}

main();