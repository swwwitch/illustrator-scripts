#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

/*

### スクリプト名：

ArtboardMaskAndRelease.jsx

### 概要

- すべてのアートボードに同じサイズの矩形を描画し、アートボード内のオブジェクトをマスクします。
- クリップグループ名をアートボード名に設定します。
- 複数アートボードにまたがるオブジェクトは複製され、それぞれのアートボードでマスクされます。
- マスクを解除する機能も搭載しており、解除後にグループを解除するオプションも用意されています。
- マージン値を指定してマスク範囲を調整できます。

### 主な機能

- アートボード単位でのマスク適用
- クリップグループ名をアートボード名に設定
- 複数アートボードにまたがるオブジェクトの複製処理
- マスク解除およびグループ解除オプション
- マージン設定
- 日本語 / 英語 UI 切り替え

### 処理の流れ

1. ダイアログでモード（マスク / 解除）とオプションを選択
2. 「マスク」選択時はアートボードごとに矩形を作成しオブジェクトをマスク
3. 「解除」選択時はクリッピングを解除し、必要に応じてグループを解除

### 更新履歴

- v1.0 (20250710) : 初期バージョン
- v1.1 (20250710) : アートボード外のオブジェクトを削除機能を追加、ロック／非表示オブジェクトのオプションを追加

---------------------------------------
### Script Name:

ArtboardMaskAndRelease.jsx

### Overview

- Draws the same-sized rectangle on all artboards and masks objects inside each artboard.
- Sets the clip group name to the artboard name.
- Duplicates objects that span multiple artboards and masks them on each.
- Includes a release function with an option to ungroup after release.
- Allows margin adjustment to control the mask area.

### Main Features

- Apply mask per artboard
- Set clip group name to artboard name
- Duplicate objects spanning multiple artboards
- Release mask and optional ungroup
- Margin setting
- Japanese / English UI support

### Workflow

1. Select mode (Mask / Release) and options in the dialog
2. When "Mask" is selected, create rectangles for each artboard and mask objects
3. When "Release" is selected, release clipping and optionally ungroup

### Change Log

- v1.0 (20250710) : Initial version
- v1.1 (20250710) : Added option to remove objects outside artboards, options for locked/hidden objects

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

    // -------------------------------
    // 日英ラベル定義 / Define labels
    // -------------------------------
    var LABELS = {
        dialogTitle: {
            ja: "アートボードでマスク " + SCRIPT_VERSION,
            en: "Mask Artboards " + SCRIPT_VERSION
        },
        modePanel: { ja: "モード", en: "Mode" },
        mask: { ja: "マスク", en: "Mask" },
        release: { ja: "解除", en: "Release" },
        maskOption: { ja: "マスクオプション", en: "Mask Options" },
        margin: { ja: "マージン", en: "Margin" },
        releaseOption: { ja: "解除オプション", en: "Release Options" },
        ungroup: { ja: "グループ解除", en: "Ungroup" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" },
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        removeOutside: { ja: "アートボード外のオブジェクトを削除", en: "Remove objects outside artboards" },
        includeLocked: { ja: "ロックされたオブジェクトを含める", en: "Include locked objects" },
        includeHidden: { ja: "非表示のオブジェクトを含める", en: "Include hidden objects" },
        ungroupLabel: { ja: "グループ解除", en: "Ungroup" },
        maskRelease: { ja: "マスク解除", en: "Release Mask" }
    };

function main() {

    if (app.documents.length == 0) {
        alert(LABELS.noDocument);
        return;
    }

    var dialog = new Window("dialog", LABELS.dialogTitle);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var panel = dialog.add("panel", undefined, LABELS.modePanel);
    panel.orientation = "row"; // 横並びに変更
    panel.alignChildren = "left";
    panel.margins = [15, 20, 15, 10];

    var rbMask = panel.add("radiobutton", undefined, LABELS.mask);
    var rbRelease = panel.add("radiobutton", undefined, LABELS.maskRelease);
    rbMask.value = true;

    // ラジオボタン切り替え時のパネル有効/無効制御
    rbMask.onClick = function() {
        marginGroup.enabled = true;
        releasePanel.enabled = false;
    };
    rbRelease.onClick = function() {
        marginGroup.enabled = false;
        releasePanel.enabled = true;
    };

    var marginGroup = dialog.add("panel", undefined, LABELS.maskOption);
    marginGroup.orientation = "column";
    marginGroup.alignChildren = "left";
    marginGroup.margins = [15, 20, 15, 10];
    var marginRow = marginGroup.add("group");
    marginRow.orientation = "row";
    marginRow.alignChildren = "left";
    marginRow.add("statictext", undefined, LABELS.margin + ":");
    // --- 単位ラベル追加 ---
    var unitLabelMap = {
        0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm", 5: "Q/H", 6: "px",
        7: "ft/in", 8: "m", 9: "yd", 10: "ft"
    };
    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return unitLabelMap[unitCode] || "pt";
    }

// Enable up/down arrow key increment/decrement on edittext inputs
function changeValueByArrowKey(editText, allowNegative) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            var isUp = event.keyName == "Up";
            var delta = 1;

            if (keyboard.shiftKey) {
                // 10の倍数にスナップ
                value = Math.floor(value / 10) * 10;
                delta = 10;
            }

            value += isUp ? delta : -delta;

            // 負数許可されない場合は0未満を禁止
            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;
        }
    });
}

    var marginInput = marginRow.add("edittext", undefined, "0");
    marginInput.characters = 5;
    marginInput.active = true;
    changeValueByArrowKey(marginInput, false);
    var unitLabel = getCurrentUnitLabel();
    marginRow.add("statictext", undefined, "(" + unitLabel + ")");
    
    var cbRemoveOutside = marginGroup.add("checkbox", undefined, LABELS.removeOutside);
    cbRemoveOutside.alignment = "left";

    var cbIncludeLocked = marginGroup.add("checkbox", undefined, LABELS.includeLocked);
    cbIncludeLocked.value = true; // デフォルトをONに設定

    // チェックボックス追加
    var cbIncludeHidden = marginGroup.add("checkbox", undefined, LABELS.includeHidden);
    cbIncludeHidden.value = true; // デフォルトをONに設定

    var releasePanel = dialog.add("panel", undefined, LABELS.releaseOption);
    releasePanel.orientation = "column";
    releasePanel.alignChildren = "left";
    releasePanel.margins = [15, 20, 15, 10];

    var cbUngroup = releasePanel.add("checkbox", undefined, LABELS.ungroupLabel);
    cbUngroup.value = true;
    releasePanel.enabled = false;

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";

    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel, { name: "cancel" });
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok, { name: "ok" });

    var result = dialog.show();

    if (result != 1) {
        return; // キャンセル時 / Cancel
    }

    var marginValue = parseFloat(marginInput.text);
    if (isNaN(marginValue)) {
        marginValue = 0;
    }

    if (rbMask.value) {
        applyMasks(marginValue, cbRemoveOutside.value, cbIncludeLocked.value, cbIncludeHidden.value);
    } else if (rbRelease.value) {
        releaseMasks(cbUngroup.value);
    }
}

function applyMasks(margin, removeOutside, includeLocked, includeHidden){
    var doc = app.activeDocument;
    var abCount = doc.artboards.length;

    var allItems = [];
    for (var j = 0; j < doc.pageItems.length; j++) {
        var item = doc.pageItems[j];
        var wasHidden = item.hidden;
        var includeItem = false;

        if (!item.locked || includeLocked) {
            if (!item.hidden || includeHidden) {
                includeItem = true;
            }
        }

        if (includeItem && item.parent.typename !== "GroupItem") {
            // hidden の場合、一時的に表示
            if (item.hidden && includeHidden) {
                item.hidden = false;
            }
            allItems.push(item);
        }

        // 元の hidden 状態に戻す（後の安全のため）
        if (includeHidden && wasHidden) {
            item.hidden = true;
        }
    }

    if (removeOutside) {
        for (var i = allItems.length - 1; i >= 0; i--) {
            var item = allItems[i];
            var isInsideAny = false;
            for (var abIdx = 0; abIdx < abCount; abIdx++) {
                var abRect = doc.artboards[abIdx].artboardRect;
                if (!(item.visibleBounds[2] < abRect[0] || item.visibleBounds[0] > abRect[2] || item.visibleBounds[3] > abRect[1] || item.visibleBounds[1] < abRect[3])) {
                    isInsideAny = true;
                    break;
                }
            }
            if (!isInsideAny) {
                item.remove();
                allItems.splice(i, 1);
            }
        }
    }

    for (var i = 0; i < abCount; i++) {
        var ab = doc.artboards[i];
        var abRect = ab.artboardRect;

        var abLeft   = abRect[0] - margin;
        var abTop    = abRect[1] + margin;
        var abRight  = abRect[2] + margin;
        var abBottom = abRect[3] - margin;

        var rect = doc.pathItems.rectangle(abTop, abLeft, abRight - abLeft, abTop - abBottom);
        rect.stroked = false;
        rect.filled = true;
        rect.fillColor = new NoColor();

        var targets = [];
        for (var k = 0; k < allItems.length; k++) {
            var item = allItems[k];
            var b = item.visibleBounds;
            if (!(b[2] < abLeft || b[0] > abRight || b[3] > abTop || b[1] < abBottom)) {
                var dup = item.duplicate();
                targets.push(dup);
            }
        }

        if (targets.length == 0) {
            rect.remove();
            continue;
        }

        var group = doc.groupItems.add();
        for (var m = 0; m < targets.length; m++) {
            targets[m].moveToBeginning(group);
        }

        rect.moveToBeginning(group);
        group.clipped = true;
        group.name = ab.name;
    }

    for (var n = allItems.length - 1; n >= 0; n--) {
        allItems[n].remove();
    }
}

function releaseMasks(ungroup){
    var doc = app.activeDocument;
    for (var i = doc.groupItems.length - 1; i >= 0; i--) {
        var group = doc.groupItems[i];
        if (group.clipped) {
            group.clipped = false;
            for (var j = group.pageItems.length - 1; j >= 0; j--) {
                var item = group.pageItems[j];
                if (item.clipping) {
                    item.remove();
                }
            }

            for (var j = group.pageItems.length - 1; j >= 0; j--) {
                var item = group.pageItems[j];
                if (item.typename === "PathItem" && item.filled == false && item.stroked == false) {
                    item.remove();
                }
            }

            // グループ内にオブジェクトが残っている場合、チェック時に解除
            if (ungroup) {
                group.selected = true;
                app.executeMenuCommand("ungroup");
            } else if (group.pageItems.length == 0) {
                group.remove();
            }
        }
    }
}

main();