#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

DeleteOutsideArtboard.jsx

### 概要

- ドキュメント内のオブジェクトをアートボードとの重なり条件で判定し、外側のオブジェクトを削除または保管用レイヤーに移動するスクリプトです。
- 「現在のアートボードのみ」または「すべてのアートボード」を対象に選択できます。

### 主な機能

- アートボード外オブジェクトの削除
- 保管用レイヤーに移動オプション
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログで対象アートボードと移動オプションを選択
2. オブジェクトとアートボードの重なりを判定
3. 重なっていないオブジェクトを削除または保管用レイヤーに移動

### 更新履歴

- v1.0.0 (20250708) : 初期バージョン

---

### Script Name:

DeleteOutsideArtboard.jsx

### Overview

- This script checks objects in the document against artboards and deletes or moves objects outside to a backup layer.
- You can select "Current Artboard Only" or "All Artboards" as target.

### Main Features

- Delete objects outside artboards
- Option to move to backup layer
- Japanese/English interface support

### Workflow

1. Select target artboards and move option in dialog
2. Check object overlap with artboards
3. Remove or move non-overlapping objects

### Changelog

- v1.0.0 (20250708): Initial version
*/

// -------------------------------
// 言語判定関数 / Language detection function
// -------------------------------
function getCurrentLang() {
    // ロケールが日本語なら "ja"、それ以外は "en" を返す
    // Return "ja" if locale is Japanese, otherwise "en"
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

// -------------------------------
// 日英ラベル定義 / Define UI labels
// -------------------------------
var LABELS = {
    dialogTitle: {
        ja: "オブジェクトを削除",
        en: "Delete Objects"
    },
    infoText: {
        ja: "アートボード外のオブジェクトを削除",
        en: "Delete objects outside artboard"
    },
    currentArtboardOnly: {
        ja: "現在のアートボードのみ",
        en: "Current Artboard Only"
    },
    allArtboards: {
        ja: "すべてのアートボード",
        en: "All Artboards"
    },
    moveToBackup: {
        ja: "保管用レイヤーに移す",
        en: "Move to backup layer"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    deleteBtn: {
        ja: "削除",
        en: "Delete"
    }
};

// ダイアログを表示し、ユーザーの選択を取得
function showDialog() {
    // ダイアログタイトル（LABELSを利用）/ Dialog title (use LABELS)
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    // 上部に説明テキストラベル / Info text label at the top
    var infoText = dialog.add("statictext", undefined, LABELS.infoText[lang]);
    infoText.alignment = "left";

    // アートボード選択パネル / Artboard selection panel
    var abGroup = dialog.add("panel", undefined, (lang === "ja" ? "対象アートボード" : "Target Artboard"));
    abGroup.orientation = "column";
    abGroup.alignChildren = "left";
    abGroup.margins = [15, 20, 15, 10];

    var abRadios = {
        current: abGroup.add("radiobutton", undefined, LABELS.currentArtboardOnly[lang]),
        all: abGroup.add("radiobutton", undefined, LABELS.allArtboards[lang])
    };
    abRadios.all.value = true;

    // オプション（保管用レイヤー） / Option (backup layer)
    var optionGroup = dialog.add("group");
    optionGroup.orientation = "column";
    optionGroup.alignChildren = "center";
    var backupCheckbox = optionGroup.add("checkbox", undefined, LABELS.moveToBackup[lang]);
    backupCheckbox.value = false;

    // ボタン / Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.deleteBtn[lang], { name: "ok" });

    cancelBtn.onClick = function() {
        dialog.close();
    };
    okBtn.onClick = function() {
        var resultCode = abRadios.all.value ? 2 : 1;
        if (backupCheckbox.value) resultCode += 10;
        dialog.close(resultCode);
    };

    var result = dialog.show();
    return result;
}

// バウンディングボックスの重なり率（大きい方の面積に対する割合）を返す
// Return overlap ratio of bounding boxes (relative to larger area)
function getOverlapRatio(a, b) {
    var ax = Math.max(0, Math.min(a[2], b[2]) - Math.max(a[0], b[0]));
    var ay = Math.max(0, Math.min(a[1], b[1]) - Math.max(a[3], b[3]));
    var overlapArea = ax * ay;
    if (overlapArea <= 0) return 0;
    var areaA = (a[2] - a[0]) * (a[1] - a[3]);
    var areaB = (b[2] - b[0]) * (b[1] - b[3]);
    var maxArea = Math.max(areaA, areaB);
    return overlapArea / maxArea;
}

// オブジェクトがアートボードと重なっているか判定
// Check if object overlaps artboard
function isOverlappingArtboard(item, artboard) {
    var itemBounds = item.visibleBounds;
    var abBounds = artboard.artboardRect;
    var overlapRatio = getOverlapRatio(itemBounds, abBounds);
    return overlapRatio > 0;
}

// アートボード外オブジェクトを収集 / Collect objects outside artboards
function collectOutsideItems(items, artboards, currentOnly, ab, result) {
    for (var i = items.length - 1; i >= 0; i--) {
        var item = items[i];

        if (item.layer && item.layer.name === "// backup") {
            continue;
        }

        // 子アイテムを先に再帰処理
        if (item.typename === "GroupItem" && item.pageItems.length > 0) {
            collectOutsideItems(item.pageItems, artboards, currentOnly, ab, result);
        }

        if (item.locked) item.locked = false;
        if (!item.visible) item.visible = true;

        var overlaps = false;
        if (currentOnly) {
            if (isOverlappingArtboard(item, ab)) overlaps = true;
        } else {
            for (var j = 0; j < artboards.length; j++) {
                if (isOverlappingArtboard(item, artboards[j])) {
                    overlaps = true;
                    break;
                }
            }
        }

        if (!overlaps) {
            result.push(item);
        }
    }
}

// アートボード外オブジェクトを削除または保管用レイヤーに移動
// Remove or move objects outside artboards
function removeOutsideObjects(currentOnly, moveToBackup) {
    if (!app.documents.length) return;
    var doc = app.activeDocument;
    var artboards = doc.artboards;
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];

    var outsideItems = [];
    collectOutsideItems(doc.pageItems, artboards, currentOnly, ab, outsideItems);

    if (outsideItems.length === 0) {
        alert((lang === "ja" ? "削除対象のオブジェクトはありません。" : "There are no objects to delete."));
        return;
    }

    for (var i = 0; i < outsideItems.length; i++) {
        if (outsideItems[i].layer && outsideItems[i].layer.locked) {
            outsideItems[i].layer.locked = false;
        }
        if (outsideItems[i].locked) {
            outsideItems[i].locked = false;
        }

        if (moveToBackup) {
            moveItemToBackupLayer(outsideItems[i], doc);
        } else {
            outsideItems[i].remove();
        }
    }
}

// オブジェクトを保管用レイヤーに移動（ロック・非表示解除を含む、グループも対応）
// Move object to backup layer (unlock/show, supports group)
function moveItemToBackupLayer(item, doc) {
    var backupLayerName = "// backup";
    var backupLayer = null;
    try {
        backupLayer = doc.layers.getByName(backupLayerName);
    } catch(e) {
        backupLayer = doc.layers.add();
        backupLayer.name = backupLayerName;
    }

    if (backupLayer.locked) backupLayer.locked = false;
    if (!backupLayer.visible) backupLayer.visible = true;

    function unlockAndShowAll(target) {
        if (target.locked) target.locked = false;
        if (!target.visible) target.visible = true;
        if (typeof target.pageItems !== "undefined" && target.pageItems.length > 0) {
            for (var i = 0; i < target.pageItems.length; i++) {
                unlockAndShowAll(target.pageItems[i]);
            }
        }
    }
    unlockAndShowAll(item);

    item.move(backupLayer, ElementPlacement.PLACEATBEGINNING);

    backupLayer.visible = false;
}

// メイン処理 / Main process
function main() {
    var dialogResult = showDialog();
    if (dialogResult === 1) {
        removeOutsideObjects(true, false);
    } else if (dialogResult === 2) {
        removeOutsideObjects(false, false);
    } else if (dialogResult === 11) {
        removeOutsideObjects(true, true);
    } else if (dialogResult === 12) {
        removeOutsideObjects(false, true);
    }
}

main();
