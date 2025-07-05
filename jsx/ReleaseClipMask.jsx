#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
ReleaseClipMask.jsx

クリッピングマスクの解除ダイアログ / Release Clipping Mask Dialog
===================================
選択中のオブジェクトに対し、指定したモードでクリッピングマスクを処理します。
Handles selected objects by releasing clipping masks based on selected mode.

モード / Mode：
- 単純に解除（Simply release）：クリップ状態を解除（グループは維持） / Release clipping, keep group
- パスを削除（Remove mask path）：マスク用パスのみ削除し、グループ解除 / Delete mask path and ungroup
- マスク対象を削除（Remove masked object）：マスク対象のみ削除し、パスは保持（グループ解除）/ Keep mask path and delete masked object

オプション / Option：
- 解除時、パスにカラーを設定 / Set fill color to path when releasing

更新日 / Last Update：2025-06-06
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