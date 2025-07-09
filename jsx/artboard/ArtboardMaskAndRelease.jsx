#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

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

- v1.0.0 (20250710) : Initial version

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.0";

// -------------------------------
// 言語判定関数 / Function to get current language
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

(function () {
    // -------------------------------
    // 日英ラベル定義 / Define labels
    // -------------------------------
    var lang = getCurrentLang();
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
        ok: { ja: "OK", en: "OK" }
    };

    if (app.documents.length == 0) {
        alert(lang == "ja" ? "ドキュメントが開かれていません。" : "No document is open.");
        return;
    }

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var panel = dialog.add("panel", undefined, LABELS.modePanel[lang]);
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.margins = [15, 20, 15, 10];

    var rbMask = panel.add("radiobutton", undefined, LABELS.mask[lang]);
    var rbRelease = panel.add("radiobutton", undefined, LABELS.release[lang]);
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

    var marginGroup = dialog.add("panel", undefined, LABELS.maskOption[lang]);
    marginGroup.orientation = "row";
    marginGroup.alignChildren = "left";
    marginGroup.margins = [15, 20, 15, 10];
    marginGroup.add("statictext", undefined, LABELS.margin[lang] + ":");
    // --- 単位ラベル追加 ---
    var unitLabelMap = {
        0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm", 5: "Q/H", 6: "px",
        7: "ft/in", 8: "m", 9: "yd", 10: "ft"
    };
    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return unitLabelMap[unitCode] || "pt";
    }
    var unitLabel = getCurrentUnitLabel();
    marginGroup.add("statictext", undefined, "(" + unitLabel + ")");
    var marginInput = marginGroup.add("edittext", undefined, "0");
    marginInput.characters = 5;
    marginInput.active = true;

    var releasePanel = dialog.add("panel", undefined, LABELS.releaseOption[lang]);
    releasePanel.orientation = "column";
    releasePanel.alignChildren = "left";
    releasePanel.margins = [15, 20, 15, 10];

    var cbUngroup = releasePanel.add("checkbox", undefined, LABELS.ungroup[lang]);
    cbUngroup.value = true;
    releasePanel.enabled = false;

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";

    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    var result = dialog.show();

    if (result != 1) {
        return; // キャンセル時 / Cancel
    }

    var marginValue = parseFloat(marginInput.text);
    if (isNaN(marginValue)) {
        marginValue = 0;
    }

    if (rbMask.value) {
        applyMasks(marginValue);
    } else if (rbRelease.value) {
        releaseMasks(cbUngroup.value);
    }
})();

function applyMasks(margin){
    var doc = app.activeDocument;
    var abCount = doc.artboards.length;

    var allItems = [];
    for (var j = 0; j < doc.pageItems.length; j++) {
        var item = doc.pageItems[j];
        if (!item.locked && !item.hidden && item.parent.typename !== "GroupItem") {
            allItems.push(item);
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