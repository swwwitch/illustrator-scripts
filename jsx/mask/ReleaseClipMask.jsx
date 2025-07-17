#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ReleaseClipMask.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ReleaseClipMask.jsx

### 概要：

- 選択オブジェクトのクリッピングマスクをモード別に解除できるIllustrator用スクリプト。
- 単純解除、パスのみ削除、画像のみ削除の3つの方法に対応。

### 主な機能：

- 単純に解除（パスと画像を残す）
- 配置画像を残してパスを削除
- パスを残して配置画像を削除
- パスにK100の塗り（不透明度15%）を適用するオプション付き
- 日本語／英語UI対応
- Q/W/Eキーによるモード切替ショートカット対応

### 処理の流れ：

1. ダイアログでモードとオプションを選択
2. 選択されたモードに従ってマスクを解除
3. 必要に応じてパスに塗りを適用

### note

https://note.com/dtp_tranist/n/nebc832e574f7

### 更新履歴：

- v1.0 (20250606) : 初期バージョン
- v1.1 (20250607) : 安定化と仕様調整
- v1.2 (20250717) : コメント整備

---

### Script Name:

ReleaseClipMask.jsx

### Readme (GitHub):

https://github.com/yourname/ai-scripts

### Overview:

- Illustrator script to release clipping masks based on selected modes.
- Supports simple release, remove mask path only, or remove masked image only.

### Main Features:

- Simple release (keep both path and image)
- Remove mask path (keep placed image)
- Remove masked image (keep path)
- Optional K100 fill color (15% opacity) to path
- Japanese/English UI support
- Q/W/E hotkeys for mode selection

### Process Flow:

1. Choose mode and option in dialog
2. Release clipping mask accordingly
3. Optionally apply fill color to path

### Update History:

- v1.0 (20250606) : Initial release
- v1.1 (20250607) : Stabilization and adjustments
- v1.2 (20250717) : Comments refactored
*/

var SCRIPT_VERSION = "v1.2";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Define Japanese-English labels */

    var LABELS = {
        dialogTitle: {
            ja: "クリッピングマスクの解除 " + SCRIPT_VERSION,
            en: "Release Clipping Mask " + SCRIPT_VERSION
        },
        simpleRelease: { ja: "単純に解除", en: "Simply release" },
        removePath: { ja: "配置画像を残して、パスを削除", en: "Remove mask path" },
        removeMasked: { ja: "パスを残して、配置画像を削除", en: "Remove masked object" },
        fillOnRelease: { ja: "解除時、パスにカラーを設定", en: "Set fill color to path when releasing" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
        releaseModeTitle: { ja: "解除方法", en: "Release Mode" }
    };

function onlyPathItemsInSelection(selectionItems) {
  for (var i = 0; i < selectionItems.length; i++) {
    var group = selectionItems[i];
    if (group.typename === "GroupItem" && group.clipped) {
      for (var j = 0; j < group.pageItems.length; j++) {
        if (group.pageItems[j].typename !== "PathItem") {
          return false;
        }
      }
    }
  }
  return true;
}

function main() {
    /* 日英ラベル定義 / Define Japanese-English labels */
   

    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";
    dlg.spacing = 10;
    dlg.margins = 20;

    var modeGroup = dlg.add("panel", undefined, LABELS.releaseModeTitle[lang]);
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";
    modeGroup.margins = [15, 20, 15, 15];

    var radioSimple = modeGroup.add("radiobutton", undefined, LABELS.simpleRelease[lang]);
    var radioRemovePath = modeGroup.add("radiobutton", undefined, LABELS.removePath[lang]);
    var radioRemoveMasked = modeGroup.add("radiobutton", undefined, LABELS.removeMasked[lang]);
    radioSimple.value = true;

    if (onlyPathItemsInSelection(app.activeDocument.selection)) {
      radioRemoveMasked.enabled = false;
    }

    /* 解除時に塗りカラーを適用するチェックボックスグループ / Checkbox group for fill color on release */
    var fillGroup = dlg.add("group");
    fillGroup.orientation = "column";
    fillGroup.alignChildren = "left";
    fillGroup.margins = [20, 5, 0, 10];
    var fillCheckbox = fillGroup.add("checkbox", undefined, LABELS.fillOnRelease[lang]);
    fillCheckbox.value = true;

    function updateFillCheckboxState() {
        fillCheckbox.enabled = !radioRemovePath.value;
    }
    radioSimple.onClick = updateFillCheckboxState;
    radioRemoveMasked.onClick = updateFillCheckboxState;
    radioRemovePath.onClick = updateFillCheckboxState;
    updateFillCheckboxState();

    var btnGroup = dlg.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "right";

    var btnCancel = btnGroup.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var btnOk = btnGroup.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    btnOk.onClick = function () {
        var mode = "simple";
        if (radioRemovePath.value) {
            mode = "removePath";
        } else if (radioRemoveMasked.value) {
            mode = "removeMasked";
        }

        dlg.close();
        executeRelease(mode, lang, fillCheckbox.value);
    };

    dlg.show();
}

function releaseSimpleMask(selectionItems) {
    for (var i = 0; i < selectionItems.length; i++) {
        var currentItem = selectionItems[i];
        if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
            currentItem.clipped = false;
            app.executeMenuCommand("ungroup");
        }
    }
}

function applyFillToPaths(item) {
  if (!item || !item.pageItems) return;

  // 一番上のパスアイテムにのみ適用 / Apply only to the topmost path item
  for (var i = 0; i < item.pageItems.length; i++) {
    var child = item.pageItems[i];
    if (child.typename === "PathItem") {
      child.filled = true;
      child.fillColor = getK100Black();
      child.opacity = 15;
      break;
    }
  }
}

function executeRelease(mode, lang, applyFill) {
    if (!app.documents.length || !app.activeDocument.selection.length) {
        alert((lang === "ja") ? "オブジェクトが選択されていません。" : "No objects selected.");
        return;
    }

    var selectionItems = app.activeDocument.selection;

    if (mode === "simple") {
        if (applyFill) {
            for (var i = 0; i < selectionItems.length; i++) {
                var currentItem = selectionItems[i];
                if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
                    currentItem.clipped = false;
                    applyFillToPaths(currentItem);
                    ungroup(currentItem);
                }
            }
        } else {
            releaseSimpleMask(selectionItems);
        }
        return;
    }

    if (mode === "removePath") {
        for (var i = 0; i < selectionItems.length; i++) {
            var currentItem = selectionItems[i];
            if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
                for (var j = 0; j < currentItem.pageItems.length; j++) {
                    var item = currentItem.pageItems[j];
                    if (item.typename === "PathItem" && item.clipping) {
                        item.remove();
                        break;
                    }
                }
                ungroup(currentItem);
            }
        }
    }

    if (mode === "removeMasked") {
        for (var i = 0; i < selectionItems.length; i++) {
            var currentItem = selectionItems[i];
            if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
                currentItem.clipped = false;

                var topPath = null;
                var fillColor, filled, opacity;
                for (var j = 0; j < currentItem.pageItems.length; j++) {
                    var item = currentItem.pageItems[j];
                    if (item.typename === "PathItem" && !topPath) {
                        topPath = item;
                        fillColor = item.fillColor;
                        filled = item.filled;
                        opacity = item.opacity;
                    }
                }

                removePlacedItems(currentItem);

                var allPaths = true;
                for (var k = 0; k < currentItem.pageItems.length; k++) {
                    if (currentItem.pageItems[k].typename !== "PathItem") {
                        allPaths = false;
                        break;
                    }
                }

                if (topPath && allPaths) {
                    topPath.fillColor = fillColor;
                    topPath.filled = filled;
                    topPath.opacity = opacity;
                }

                if (applyFill) {
                    applyFillToPaths(currentItem);
                }

                ungroup(currentItem);
            }
        }
    }
    return;
    
}

function removePlacedItems(groupItem) {
    var itemsToRemove = [];
    for (var i = 0; i < groupItem.pageItems.length; i++) {
        var item = groupItem.pageItems[i];
        if (item.typename === "PlacedItem") {
            itemsToRemove.push(item);
        }
    }
    for (var j = 0; j < itemsToRemove.length; j++) {
        itemsToRemove[j].remove();
    }
}

function getColor(r, g, b, c, m, y, k) {
    if (app.activeDocument.documentColorSpace == DocumentColorSpace.CMYK) {
        var tmpColor = new CMYKColor();
        tmpColor.cyan = c;
        tmpColor.magenta = m;
        tmpColor.yellow = y;
        tmpColor.black = k;
        return tmpColor;
    } else {
        var tmpColor = new RGBColor();
        tmpColor.red = r;
        tmpColor.green = g;
        tmpColor.blue = b;
        return tmpColor;
    }
}

function releaseAndStyleClippingGroup(groupItem) {
  groupItem.clipped = false;
  removePlacedItems(groupItem);
  applyFillToPaths(groupItem);
  ungroup(groupItem);
}

function getK100Black() {
  var cmykColor = new CMYKColor();
  cmykColor.cyan = 0;
  cmykColor.magenta = 0;
  cmykColor.yellow = 0;
  cmykColor.black = 100;
  return cmykColor;
}

function ungroup(groupItem) {
  var parent = groupItem.parent;
  while (groupItem.pageItems.length > 0) {
    groupItem.pageItems[0].move(parent, ElementPlacement.PLACEATEND);
  }
  groupItem.remove();
}

main();