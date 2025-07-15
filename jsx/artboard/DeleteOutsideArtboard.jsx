#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

/*

### スクリプト名：

DeleteOutsideArtboard.jsx

### 概要

- ドキュメント内のオブジェクトをアートボードとの重なり条件で判定し、外側のオブジェクトを削除または保管用レイヤーに移動します。
- 「現在のアートボードのみ」または「すべてのアートボード」を対象に選択できます。
- 新たに「すべてのオブジェクト」または「選択オブジェクト以外（現在のアートボード内）」を対象範囲として選択可能。
- 新たに「ロックされたオブジェクトを含む」オプションを追加（デフォルトON）。

### 主な機能

- デフォルトは「すべてのアートボード」かつ「すべてのオブジェクト」かつ「ロックされたオブジェクトを含む」。
- 「選択オブジェクト以外（現在のアートボード内）」モードでは、選択されたオブジェクトを除外して現在のアートボード内のオブジェクトを処理。
- オブジェクトを選択しておく必要はありません。
- ロックされたオブジェクト、非表示のオブジェクトも対象です。
- ［保管用レイヤーに移す］オプションをONにすると、「// backup」レイヤーに移動します。「// backup」レイヤーは非表示になります。
- ［ロックされたオブジェクトを含む］オプションはデフォルトでONです。

### 処理の流れ

1. ダイアログで対象範囲、対象アートボード、移動オプション、ロック含むオプションを選択
2. オブジェクトとアートボードの重なりを判定
3. 重なっていないオブジェクトを削除または保管用レイヤーに移動
4. 「選択オブジェクト以外」モードでは選択オブジェクトを除外（現在のアートボード内）
5. 「ロックされたオブジェクトを含む」モードではロックされたオブジェクトも処理対象

### 更新履歴

- v1.0 (20250708) : 初期バージョン
- v1.1 (20250708) : 微調整
- v1.2 (20250713) : 「選択オブジェクト以外（現在のアートボード内）」モードを追加
- v1.3 (20250713) : 「ロックされたオブジェクトを無視する」オプションを追加
- v1.4 (202507xx) : 「ロックされたオブジェクトを含む」オプションに変更しデフォルトONに

### Script Name:

DeleteOutsideArtboard.jsx

### Overview

- Checks objects in the document against artboards and deletes or moves objects outside to a backup layer.
- Allows selection of "Current Artboard Only" or "All Artboards" as target.
- Added new mode option to select "All Objects" or "Exclude Selected (Current Artboard)".
- Changed option to "Include locked objects" (default ON).

### Main Features

- Default is "All Artboards", "All Objects", and "Include locked objects" ON.
- "Exclude Selected (Current Artboard)" mode excludes selected objects within current artboard from deletion.
- No need to pre-select objects.
- Locked and hidden objects are included by default.
- Option to move outside objects to a hidden "// backup" layer.
- Option to exclude locked objects during processing by unchecking the box.

### Workflow

1. Select target scope, target artboards, move option, and include locked option in dialog
2. Check object overlap with artboards
3. Remove or move non-overlapping objects
4. Exclude selected objects in "Exclude Selected" mode (current artboard only)
5. Include or exclude locked objects based on option

### Changelog

- v1.0 (20250708): Initial version
- v1.1 (20250708): Minor adjustments
- v1.2 (20250713): Added 'Exclude Selected (Current Artboard)' mode option
- v1.3 (20250713): Added 'Ignore locked objects' option
- v1.4 (202507xx): Changed to 'Include locked objects' option with default ON

*/

// -------------------------------
// Localization labels
// -------------------------------
var lang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

var LABELS = {
    dialogTitle: { ja: "オブジェクトを削除", en: "Delete Objects" },
    scopePanel: { ja: "アートボード内", en: "Inside Artboard" },
    excludeSelectedRadio: { ja: "選択オブジェクトを残す", en: "Exclude Selected Objects" },
    allObjectsRadio: { ja: "すべて残す", en: "Keep All Objects" },
    outsidePanel: { ja: "アートボード外", en: "Outside Artboard" },
    deleteRadio: { ja: "削除", en: "Delete" },
    ignoreRadio: { ja: "無視（残す）", en: "Ignore (Keep)" },
    includeLockedCheckbox: { ja: "ロックされたオブジェクトを含む", en: "Include Locked Objects" },
    backupCheckbox: { ja: "保管用レイヤーに移す", en: "Move to Backup Layer" },
    cancelButton: { ja: "キャンセル", en: "Cancel" },
    okButton: { ja: "削除", en: "Delete" },
    alertNoSelection: { ja: "選択オブジェクトがありません。", en: "No objects selected." },
    alertNoTargets: { ja: "削除対象のオブジェクトはありません。", en: "No objects to delete." }
};

// ダイアログを表示し、ユーザーの選択を取得
function showDialog() {
    // ダイアログタイトル（日本語固定）
    var dialog = new Window("dialog", LABELS.dialogTitle);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    // 対象範囲パネル / Target scope panel
    var scopeGroup = dialog.add("panel", undefined, LABELS.scopePanel);
    scopeGroup.orientation = "column";
    scopeGroup.alignChildren = "left";
    scopeGroup.margins = [15, 20, 15, 10];

    var scopeRadios = {
        excludeSelected: scopeGroup.add("radiobutton", undefined, LABELS.excludeSelectedRadio),
        allObjects: scopeGroup.add("radiobutton", undefined, LABELS.allObjectsRadio)
    };

    // Set radio default based on selection
    var hasSelection = false;
    try {
        hasSelection = app.documents.length > 0 && app.activeDocument.selection && app.activeDocument.selection.length > 0;
    } catch (e) {
        hasSelection = false;
    }
    if (hasSelection) {
        scopeRadios.excludeSelected.value = true;
    } else {
        scopeRadios.allObjects.value = true;
    }

    // アートボード外処理パネル / Outside artboard action panel
    var abGroup = dialog.add("panel", undefined, LABELS.outsidePanel);
    abGroup.orientation = "column";
    abGroup.alignChildren = "left";
    abGroup.margins = [15, 20, 15, 10];

    var deleteRadio = abGroup.add("radiobutton", undefined, LABELS.deleteRadio);
    var ignoreRadio = abGroup.add("radiobutton", undefined, LABELS.ignoreRadio);
    ignoreRadio.value = true;

    // オプション（保管用レイヤー、ロック含む） / Option (backup layer, include locked)
    var optionGroup = dialog.add("group");
    optionGroup.orientation = "column";
    optionGroup.alignChildren = "left";
    optionGroup.margins = [15, 0, 15, 10];

    var ignoreLockedCheckbox = optionGroup.add("checkbox", undefined, LABELS.includeLockedCheckbox);
    ignoreLockedCheckbox.value = true;

    var backupCheckbox = optionGroup.add("checkbox", undefined, LABELS.backupCheckbox);
    backupCheckbox.value = false;

    // ボタン / Buttons
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancelButton);
    var okBtn = buttonGroup.add("button", undefined, LABELS.okButton, {
        name: "ok"
    });

    cancelBtn.onClick = function() {
        dialog.close();
    };
    okBtn.onClick = function() {
        var resultCode = 0;
        if (scopeRadios.excludeSelected.value) {
            resultCode += 100;
        }
        // Only two options: 削除 (deleteRadio) or 無視 (ignoreRadio)
        if (deleteRadio.value) {
            resultCode += 1;
        } else if (ignoreRadio.value) {
            resultCode += 3;
        }
        if (backupCheckbox.value) resultCode += 10;
        if (!ignoreLockedCheckbox.value) resultCode += 1000;
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

// 指定したオブジェクトが対象のアートボード群のいずれかと重なっているか判定
// Check if item overlaps any of the artboards
function checkOverlapWithArtboards(item, artboards) {
    for (var i = 0; i < artboards.length; i++) {
        if (isOverlappingArtboard(item, artboards[i])) {
            return true;
        }
    }
    return false;
}

// アートボード外オブジェクトを収集 / Collect objects outside artboards
function collectOutsideItems(items, artboards, currentOnly, ab, result, includeLocked) {
    for (var i = items.length - 1; i >= 0; i--) {
        var item = items[i];

        if (item.layer && item.layer.name === "// backup") {
            continue;
        }

        if (!includeLocked && item.locked) continue;

        if (item.locked) item.locked = false;
        if (!item.visible) item.visible = true;

        if (item.typename === "GroupItem") {
            var groupTarget = item;
            // If clipped group, use the clipping path for bounds
            if (item.clipped) {
                for (var k = 0; k < item.pageItems.length; k++) {
                    if (item.pageItems[k].clipping) {
                        groupTarget = item.pageItems[k];
                        break;
                    }
                }
            }

            var overlapsGroup = false;
            if (currentOnly) {
                overlapsGroup = isOverlappingArtboard(groupTarget, ab);
            } else {
                overlapsGroup = checkOverlapWithArtboards(groupTarget, artboards);
            }

            if (!overlapsGroup) {
                result.push(item);
                continue; // Skip inside items if group added
            }

            // If group overlaps, check inside
            collectOutsideItems(item.pageItems, artboards, currentOnly, ab, result, includeLocked);
            continue;
        }

        var overlaps = false;
        if (currentOnly) {
            overlaps = isOverlappingArtboard(item, ab);
        } else {
            overlaps = checkOverlapWithArtboards(item, artboards);
        }

        if (!overlaps) {
            result.push(item);
        }
    }
}

// アートボード外オブジェクトを削除または保管用レイヤーに移動
// Remove or move objects outside artboards
// mode: 0=All Objects, 1=Exclude Selected (Current Artboard)
function removeOutsideObjects(currentOnly, moveToBackup, mode, includeLocked) {
    if (!app.documents.length) return;
    var doc = app.activeDocument;
    var artboards = doc.artboards;
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];

    if (currentOnly === null) {
        // "無視" 選択時は何もしないで終了
        return;
    }

    var outsideItems = [];
    collectOutsideItems(doc.pageItems, artboards, currentOnly, ab, outsideItems, includeLocked);

    if (mode === 1) {
        // 「選択オブジェクトを残す」: Delete/move all objects inside current artboard except selected
        var sel = doc.selection;
        if (!sel || sel.length === 0) {
            alert(LABELS.alertNoSelection);
            return;
        }
        // Collect all objects inside current artboard (including locked/hidden as per includeLocked)
        var insideItems = [];

        function collectInsideItems(items) {
            for (var i = items.length - 1; i >= 0; i--) {
                var item = items[i];
                if (item.layer && item.layer.name === "// backup") continue;
                if (!includeLocked && item.locked) continue;
                if (item.locked) item.locked = false;
                if (!item.visible) item.visible = true;
                if (item.typename === "GroupItem") {
                    var groupTarget = item;
                    if (item.clipped) {
                        for (var k = 0; k < item.pageItems.length; k++) {
                            if (item.pageItems[k].clipping) {
                                groupTarget = item.pageItems[k];
                                break;
                            }
                        }
                    }
                    if (isOverlappingArtboard(groupTarget, ab)) {
                        insideItems.push(item);
                        // Also check inside group for further items
                        collectInsideItems(item.pageItems);
                    }
                    continue;
                }
                if (isOverlappingArtboard(item, ab)) {
                    insideItems.push(item);
                }
            }
        }
        collectInsideItems(doc.pageItems);
        // Exclude selected objects
        var selectedSet = {};
        for (var si = 0; si < sel.length; si++) {
            selectedSet[sel[si]] = true;
        }
        // For ExtendScript, compare by reference
        var filteredItems = [];
        for (var ii = 0; ii < insideItems.length; ii++) {
            var isSelected = false;
            for (var sj = 0; sj < sel.length; sj++) {
                if (insideItems[ii] === sel[sj]) {
                    isSelected = true;
                    break;
                }
            }
            if (!isSelected) {
                filteredItems.push(insideItems[ii]);
            }
        }
        outsideItems = filteredItems;
    }

    if (outsideItems.length === 0) {
        alert(LABELS.alertNoTargets);
        return;
    }

    for (var i = 0; i < outsideItems.length; i++) {
        var item = outsideItems[i];

        try {
            if (item.locked) item.locked = false;
            if (!item.visible) item.visible = true;
            if (item.layer && item.layer.locked) item.layer.locked = false;

            if (moveToBackup) {
                moveItemToBackupLayer(item, doc);
            } else {
                item.remove();
            }
        } catch (e) {
            $.writeln("Skipped invalid object: " + e);
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
    } catch (e) {
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
    if (dialogResult === 0 || dialogResult === undefined) return;

    var mode = 0;
    var includeLocked = true;
    if (dialogResult >= 1000) {
        includeLocked = false;
        dialogResult -= 1000;
    }
    if (dialogResult >= 100) {
        mode = 1;
        dialogResult -= 100;
    }

    var currentOnly = null;
    if (mode === 1) {
        // 「選択オブジェクトを残す」モードの場合は必ず現在のアートボードのみを対象とする
        currentOnly = true;
    } else {
        if (dialogResult === 1) {
            currentOnly = true;
        } else if (dialogResult === 3) {
            currentOnly = null; // 無視モード
        }
    }

    var moveToBackup = false;
    if (dialogResult >= 10) {
        moveToBackup = true;
        dialogResult -= 10;
    }

    if (currentOnly !== null) {
        removeOutsideObjects(currentOnly, moveToBackup, mode, includeLocked);
    }
}

main();