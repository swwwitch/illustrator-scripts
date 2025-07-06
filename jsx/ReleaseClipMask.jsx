#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ReleaseClipMask.jsx

### 概要

- 選択オブジェクトのクリッピングマスクを複数のモードで解除するIllustrator用スクリプトです。
- 単純解除、マスクパス削除、マスク対象削除など柔軟な操作が可能です。

### 主な機能

- 単純解除（グループ維持）
- マスクパス削除（配置画像を残す）
- マスク対象削除（パスを残す）
- 解除時にパスへ塗りカラー設定オプション
- 日本語／英語インターフェース対応

### 処理の流れ

1. モードとオプションをダイアログで選択
2. 選択内容に応じてクリッピングマスクを解除
3. 必要に応じてパスにカラーを設定

### 更新履歴

- v1.0.0 (20250606) : 初期バージョン

---

### Script Name:

ReleaseClipMask.jsx

### Overview

- An Illustrator script to release clipping masks on selected objects with multiple flexible modes.
- Supports simple release, remove mask path, or remove masked content.

### Main Features

- Simple release (keep group)
- Remove mask path (keep placed content)
- Remove masked object (keep path)
- Option to set fill color on path when releasing
- Japanese and English UI support

### Process Flow

1. Select mode and option in dialog
2. Release clipping masks based on selected mode
3. Optionally set fill color to paths

### Update History

- v1.0.0 (20250606): Initial version
*/

function main() {
    // -------------------------------
    // 日英ラベル定義　Define label
    // -------------------------------
    var lang = getCurrentLang();
    var LABELS = {
        dialogTitle: { ja: "クリッピングマスクの解除", en: "Release Clipping Mask" },
        simpleRelease: { ja: "単純に解除", en: "Simply release" },
        removePath: { ja: "配置画像を残して、パスを削除", en: "Remove mask path" },
        removeMasked: { ja: "パスを残して、配置画像を削除", en: "Remove masked object" },
        fillOnRelease: { ja: "解除時、パスにカラーを設定", en: "Set fill color to path when releasing" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
        releaseModeTitle: { ja: "解除方法", en: "Release Mode" }
    };

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

    // Checkbox group for fill color on release
    var fillGroup = dlg.add("group");
    fillGroup.orientation = "column";
    fillGroup.alignChildren = "left";
    fillGroup.margins = [20, 5, 0, 10];
    var fillCheckbox = fillGroup.add("checkbox", undefined, LABELS.fillOnRelease[lang]);

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

function releaseSimpleMask(selectionItems, applyFill) {
    for (var i = 0; i < selectionItems.length; i++) {
        var currentItem = selectionItems[i];
        if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
            currentItem.clipped = false;
            for (var j = 0; j < currentItem.pageItems.length; j++) {
                if (currentItem.pageItems[j].clipping === true) {
                    if (applyFill) {
                        currentItem.pageItems[j].filled = true;
                        currentItem.pageItems[j].fillColor = getColor(0, 0, 0, 0, 0, 0, 100); // K100
                        currentItem.pageItems[j].opacity = 15;
                    }
                    break;
                }
            }
        }
    }
    app.executeMenuCommand("ungroup"); // Ungroup after release
}

function executeRelease(mode, lang, applyFill) {
    if (!app.documents.length || !app.activeDocument.selection.length) {
        alert((lang === "ja") ? "オブジェクトが選択されていません。" : "No objects selected.");
        return;
    }

    var selectionItems = app.activeDocument.selection;

    if (mode === "simple") {
        releaseSimpleMask(selectionItems, applyFill);
        return;
    }

    if (mode === "removePath") {
        for (var i = 0; i < selectionItems.length; i++) {
            var currentItem = selectionItems[i];

            if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
                currentItem.clipped = false;
                if (currentItem.pageItems.length > 0) {
                    currentItem.pageItems[0].remove(); // Remove mask path (topmost item)
                }
            }
        }
        app.executeMenuCommand("ungroup"); // Ungroup after release
    }

    if (mode === "removeMasked") {
        for (var i = 0; i < selectionItems.length; i++) {
            var currentItem = selectionItems[i];
            if (currentItem.typename === "GroupItem" && currentItem.clipped === true) {
                currentItem.clipped = false;
                for (var j = 0; j < currentItem.pageItems.length; j++) {
                    if (currentItem.pageItems[j].clipping === true) {
                        if (applyFill) {
                            currentItem.pageItems[j].filled = true;
                            currentItem.pageItems[j].fillColor = getColor(0, 0, 0, 0, 0, 0, 100); // K100
                            currentItem.pageItems[j].opacity = 15;
                        }
                        break;
                    }
                }
                var groupItems = currentItem.pageItems;
                if (groupItems.length > 0) {
                    var originalLayer = currentItem.layer;
                    groupItems[0].move(originalLayer, ElementPlacement.PLACEATEND);
                }
                // Remove remaining items (e.g. placed images)
                while (groupItems.length > 0) {
                    groupItems[0].remove();
                }
            }
        }
        return;
    }
}

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
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

main();