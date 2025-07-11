#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

CreateGuidesFromSelection

### 概要

- 選択オブジェクトの中央またはエッジを基準にアートボード全体にガイドを作成します。
- プレビュー境界の使用、はみだし、マージン設定が可能です。

### 主な機能

- 中央・エッジモード切替
- プレビュー境界使用オプション
- はみだし・マージン設定
- 既存ガイド削除オプション
- 多言語（日英）対応

### 処理の流れ

1. ダイアログでオプションを選択
2. オブジェクトの境界座標を取得
3. 中央またはエッジ基準でガイドを作成

### 更新履歴

- v1.0 (20250711) : 初期バージョン
- v1.1 (20250711) : 複数選択、クリップグループ対応、日英対応追加
- v1.2 (20250711) : はみだし・マージン設定、ガイド削除オプション追加

---

### Script Name:

CreateGuidesFromSelection

### Overview

- Creates guides across the artboard based on the center or edges of selected objects.
- Supports visible bounds, overflow, and margin options.

### Features

- Center or edge mode
- Use visible bounds option
- Overflow and margin settings
- Clear existing guides option
- Multi-language support (Japanese/English)

### Workflow

1. Select options in dialog
2. Get object bounds
3. Create guides based on selected mode

### Update History

- v1.0 (20250711): Initial version
- v1.1 (20250711): Multiple selection, clip group, language support
- v1.2 (20250711): Added overflow & margin, clear guides option
*/

var SCRIPT_VERSION = "v1.2";

// 言語取得 / Get current language
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// ラベル定義 / UI Label Definitions
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
    useVisible: {
        ja: "プレビュー境界を使用",
        en: "Use visible bounds"
    },
    clearGuides: {
        ja: "「_guide」レイヤー内のガイドを削除",
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

function createEdgeGuides(bounds, artboardRect, guideLayer, offsetValue, marginValue) {
    var leftX = bounds[0] - marginValue;
    var rightX = bounds[2] + marginValue;
    var topY = bounds[1] + marginValue;
    var bottomY = bounds[3] - marginValue;

    var guideLeft = artboardRect[0] - offsetValue;
    var guideRight = artboardRect[2] + offsetValue;
    var guideTop = artboardRect[1] + offsetValue;
    var guideBottom = artboardRect[3] - offsetValue;

    var leftGuide = guideLayer.pathItems.add();
    leftGuide.setEntirePath([[leftX, guideTop], [leftX, guideBottom]]);
    leftGuide.guides = true;

    var rightGuide = guideLayer.pathItems.add();
    rightGuide.setEntirePath([[rightX, guideTop], [rightX, guideBottom]]);
    rightGuide.guides = true;

    var topGuide = guideLayer.pathItems.add();
    topGuide.setEntirePath([[guideLeft, topY], [guideRight, topY]]);
    topGuide.guides = true;

    var bottomGuide = guideLayer.pathItems.add();
    bottomGuide.setEntirePath([[guideLeft, bottomY], [guideRight, bottomY]]);
    bottomGuide.guides = true;
}

function main() {
    try {
        var isValid = true;
        var doc = app.activeDocument;
        var artboards = doc.artboards;
        var activeIndex = artboards.getActiveArtboardIndex();
        var artboardRect = artboards[activeIndex].artboardRect;

        // 「_guide」レイヤーを取得または作成
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

        if (app.selection.length < 1) {
            alert(LABELS.alertSelect[lang]);
            isValid = false;
        }

        var rbCenter, rbEdge, cbUseVisible, cbClearGuides, offsetInput, marginInput;

        if (isValid) {
            // ダイアログ作成
            var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
            dlg.orientation = "column";
            dlg.alignChildren = "left";

            var modePanel = dlg.add("panel", undefined, LABELS.panelTitle[lang]);
            modePanel.orientation = "row";
            modePanel.alignChildren = "left";
            modePanel.margins = [15, 20, 15, 10];

            rbCenter = modePanel.add("radiobutton", undefined, LABELS.center[lang]);
            rbEdge = modePanel.add("radiobutton", undefined, LABELS.edge[lang]);
            rbCenter.value = false;
            rbEdge.value = true;

            cbUseVisible = dlg.add("checkbox", undefined, LABELS.useVisible[lang]);
            cbUseVisible.value = true;

            cbClearGuides = dlg.add("checkbox", undefined, LABELS.clearGuides[lang]);
            cbClearGuides.value = true;

            var offsetGroup = dlg.add("group");
            offsetGroup.add("statictext", undefined, LABELS.offset[lang]);
            offsetInput = offsetGroup.add("edittext", undefined, "6");
            offsetInput.characters = 5;
            offsetGroup.add("statictext", undefined, getCurrentUnitLabel());

            var marginGroup = dlg.add("group");
            marginGroup.add("statictext", undefined, LABELS.margin[lang]);
            marginInput = marginGroup.add("edittext", undefined, "0");
            marginInput.characters = 5;
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

        var unitCode = app.preferences.getIntegerPreference("rulerType");
        var factor = getPtFactorFromUnitCode(unitCode);

        var offsetValue = parseFloat(offsetInput.text);
        if (isNaN(offsetValue)) offsetValue = 0;
        offsetValue *= factor;

        var marginValue = parseFloat(marginInput.text);
        if (isNaN(marginValue)) marginValue = 0;
        marginValue *= factor;

        var selectionItems = app.selection;

        // テキストオブジェクトを事前に複製し、hidden 配列で管理
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

        // bounds 計算用の itemsForBounds を共通ロジックで決定
        var itemsForBounds = [];
        // itemsForBounds 構築
        if (textCopies.length > 0) {
            // 非テキストオブジェクトを含める
            for (var i = 0; i < selectionItems.length; i++) {
                var item = selectionItems[i];
                if (item.typename !== "TextFrame") {
                    // クリップグループの場合はマスクパスを使う
                    if (item.typename === "GroupItem" && item.clipped) {
                        var mask = null;
                        for (var j = 0; j < item.pageItems.length; j++) {
                            if (item.pageItems[j].clipping) {
                                mask = item.pageItems[j];
                                break;
                            }
                        }
                        if (mask) {
                            itemsForBounds.push(mask);
                        } else {
                            itemsForBounds.push(item);
                        }
                    } else {
                        itemsForBounds.push(item);
                    }
                }
            }
            // アウトライン化したテキストも含める
            for (var i = 0; i < textCopies.length; i++) {
                itemsForBounds.push(textCopies[i]);
            }
        } else {
            // テキストオブジェクトがない場合でもクリップグループ処理
            for (var i = 0; i < selectionItems.length; i++) {
                var item = selectionItems[i];
                if (item.typename === "GroupItem" && item.clipped) {
                    var mask = null;
                    for (var j = 0; j < item.pageItems.length; j++) {
                        if (item.pageItems[j].clipping) {
                            mask = item.pageItems[j];
                            break;
                        }
                    }
                    if (mask) {
                        itemsForBounds.push(mask);
                    } else {
                        itemsForBounds.push(item);
                    }
                } else {
                    itemsForBounds.push(item);
                }
            }
        }

        // bounds 計算共通化
        var bounds;
        if (cbUseVisible.value) {
            bounds = itemsForBounds[0].visibleBounds.concat();
        } else {
            bounds = itemsForBounds[0].geometricBounds.concat();
        }
        for (var i = 1; i < itemsForBounds.length; i++) {
            var item = itemsForBounds[i];
            var itemBounds = cbUseVisible.value ? item.visibleBounds : item.geometricBounds;
            if (itemBounds[0] < bounds[0]) bounds[0] = itemBounds[0];
            if (itemBounds[1] > bounds[1]) bounds[1] = itemBounds[1];
            if (itemBounds[2] > bounds[2]) bounds[2] = itemBounds[2];
            if (itemBounds[3] < bounds[3]) bounds[3] = itemBounds[3];
        }

        // オブジェクトがアートボード外にある場合、アラートを出す
        if (
            bounds[2] < artboardRect[0] || // 右端がアートボード左より左
            bounds[0] > artboardRect[2] || // 左端がアートボード右より右
            bounds[1] < artboardRect[3] || // 上端がアートボード下より下
            bounds[3] > artboardRect[1]    // 下端がアートボード上より上
        ) {
            alert("オブジェクトがアートボード外にあります。");
            guideLayer.locked = true;
            // 元のテキストを再表示
            for (var i = 0; i < originalTexts.length; i++) {
                originalTexts[i].hidden = false;
            }
            // アウトライン化した複製オブジェクトを削除
            for (var i = 0; i < textCopies.length; i++) {
                textCopies[i].remove();
            }
            return;
        }

        if (rbCenter.value) {
            createCenterGuides(bounds, artboardRect, guideLayer);
        } else if (rbEdge.value) {
            createEdgeGuides(bounds, artboardRect, guideLayer, offsetValue, marginValue);
        }
        guideLayer.locked = true;

        // ガイド作成後、元のテキストを再表示
        for (var i = 0; i < originalTexts.length; i++) {
            originalTexts[i].hidden = false;
        }
        // アウトライン化した複製オブジェクトを削除
        for (var i = 0; i < textCopies.length; i++) {
            textCopies[i].remove();
        }

    } catch (e) {
        alert(LABELS.error[lang] + e);
    }
}

main();