#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

ResizeClipMask.jsx

### Readme （GitHub）：

（準備中）

### 概要：

- クリップマスク付きグループに対し、マージンを設定してマスクの矩形サイズを調整します。
- プレビュー機能を備え、OKで確定、Cancelで元に戻せます。

### 主な機能：

- マスクが矩形かどうかを判定し、対象のみ処理
- マージン量を±ボタン・±キー・±反転で直感的に操作可能
- キャンセル時には元のサイズへ復元

### 処理の流れ：

1. 選択オブジェクトからマスクパスを収集
2. プレビューとしてマージンを加えたサイズに一時変更
3. OKで確定、Cancelで元に戻す

### 更新履歴：

- v1.0 (20250710) : 初版
- v1.1 (20250718) : ↑↓入力に対応
- v1.2 (20250717) : プレビュー反映ロジックを調整

---
### Script Name:

ResizeClipMask.jsx

### Readme (GitHub):

(Coming soon)

### Overview:

- Adjusts rectangle clip mask size by applying margin to clipped groups
- Includes preview and cancel support

### Main Features:

- Detects rectangular clipping masks only
- Intuitive margin control via ± buttons, arrow keys, and polarity switch
- Restores original state when canceled

### Workflow:

1. Collect clipping masks from selected objects
2. Apply temporary margin as preview
3. Confirm with OK, restore with Cancel

### Update History:

- v1.0 (20250710): Initial version
- v1.1 (20250718): Arrow key input supported
- v1.2 (20250717): Adjusted preview application logic

*/

var SCRIPT_VERSION = "v1.2";

/* 現在の言語取得 / Get current language */
function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* UIラベル定義 / UI label definitions */
var LABELS = {
    notRect: { ja: "このマスクは長方形ではありません。処理をスキップします。", en: "This mask is not a rectangle. Skipping." },
    noMask: { ja: "マスクパスが見つかりませんでした。", en: "No mask path found." },
    noSelection: { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
    cancel: { ja: "キャンセル", en: "Cancel" },
    ok: { ja: "OK", en: "OK" },
    dialogTitle: {
        ja: "マスクパスのサイズ変更 " + SCRIPT_VERSION,
        en: "Resize Mask Path " + SCRIPT_VERSION
    },
    margin: { ja: "マージン", en: "Margin" }
};

/* 単位コードとラベルのマッピング / Unit code to label mapping */
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

/* 現在の単位ラベル取得 / Get current unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

/* 2点が等しいか判定 / Check if two points are equal */
function pointsEqual(p1, p2) {
    return p1[0] == p2[0] && p1[1] == p2[1];
}

/* プラスボタン処理 / Handle plus button */
function handlePlus(input) {
    var val = parseFloat(input.text);
    if (isNaN(val)) val = 0;
    var keyboard = ScriptUI.environment.keyboardState;
    var delta = keyboard.altKey ? 0.1 : 1;
    val += delta;
    val = keyboard.altKey ? Math.round(val * 10) / 10 : Math.round(val);
    input.text = String(val);
    if (typeof input.onChangeValue === "function") input.onChangeValue(val);
}

/* マイナスボタン処理 / Handle minus button */
function handleMinus(input) {
    var val = parseFloat(input.text);
    if (isNaN(val)) val = 0;
    var keyboard = ScriptUI.environment.keyboardState;
    var delta = keyboard.altKey ? 0.1 : 1;
    val -= delta;
    val = keyboard.altKey ? Math.round(val * 10) / 10 : Math.round(val);
    input.text = String(val);
    if (typeof input.onChangeValue === "function") input.onChangeValue(val);
}

/* 反転ボタン処理 / Handle swap button */
function handleSwap(input) {
    var val = parseFloat(input.text);
    if (isNaN(val)) val = 0;
    val = -val;
    input.text = String(val);
    if (typeof input.onChangeValue === "function") input.onChangeValue(val);
}

/* マージンダイアログ表示 / Show margin dialog */
function showMarginDialog(defaultValue, unitLabel, previewCallback) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "left";
    dlg.margins = 15;

    var inputGroup = dlg.add("group");
    inputGroup.add("statictext", undefined, LABELS.margin[lang] + " (" + unitLabel + "):");

    var inputSubGroup = inputGroup.add("group");
    inputSubGroup.orientation = "row";

    var input = inputSubGroup.add("edittext", undefined, defaultValue);
    input.characters = 4;
    changeValueByArrowKey(input);
    input.onChangeValue = previewCallback;

    var buttonGroup = inputSubGroup.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.spacing = 1;

    var plusMinusGroup = buttonGroup.add("group");
    plusMinusGroup.orientation = "column";
    plusMinusGroup.spacing = 1;

    var plusBtn = plusMinusGroup.add("button", [0, 0, 20, 15], "+");
    var minusBtn = plusMinusGroup.add("button", [0, 0, 20, 15], "-");

    var zeroGroup = buttonGroup.add("group");
    zeroGroup.orientation = "column";
    zeroGroup.alignChildren = ["center", "center"];

    var swapBtn = zeroGroup.add("button", [0, 0, 20, 31], "±");

    plusBtn.onClick = function() { handlePlus(input); };
    minusBtn.onClick = function() { handleMinus(input); };
    swapBtn.onClick = function() { handleSwap(input); };

    input.addEventListener("changing", function() {
        var val = parseFloat(input.text);
        if (!isNaN(val) && typeof previewCallback === "function") {
            previewCallback(val);
            app.redraw();
        }
    });

    var btns = dlg.add("group");
    btns.alignment = "right";
    var cancel = btns.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
    var ok = btns.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    input.active = true;
    var result = dlg.show();
    if (result != 1) {
        if (typeof previewCallback === "function") restoreOriginalMaskRects();
        return null;
    }

    var margin = parseFloat(input.text);
    if (isNaN(margin)) {
        alert(LABELS.noSelection[lang]);
        return null;
    }
    return margin;
}

/* 選択からマスクパスを抽出 / Collect mask paths from selection */
function collectMaskPaths(selection) {
    var masks = [];
    for (var i = 0; i < selection.length; i++) {
        var group = selection[i];
        if (group.typename === "GroupItem" && group.clipped) {
            for (var j = 0; j < group.pageItems.length; j++) {
                var item = group.pageItems[j];
                if (item.typename === "PathItem" && item.clipping) {
                    masks.push(item);
                    break;
                }
            }
        }
    }
    return masks;
}

/* マスク矩形の座標情報取得 / Get rectangle info from mask */
function getRectInfo(mask) {
    return {
        left: mask.left,
        top: mask.top,
        width: mask.width,
        height: mask.height
    };
}

/* 元のマスク矩形情報保存用 / Store original mask rectangles */
var originalRects = [];

/* マスクに一時的にマージンを適用 / Apply temporary margin to masks */
function applyTemporaryMarginToMasks(masks, margin) {
    restoreOriginalMaskRects(); // ★追加：まず元に戻す
    originalRects = [];
    for (var i = 0; i < masks.length; i++) {
        var mask = masks[i];
        if (!mask.clipping) continue;

        originalRects.push({
            mask: mask,
            top: mask.top,
            left: mask.left,
            width: mask.width,
            height: mask.height
        });

        mask.top += margin;
        mask.left -= margin;
        mask.width += margin * 2;
        mask.height += margin * 2;
    }
}

/* 元のマスクサイズに戻す / Restore original mask rectangles */
function restoreOriginalMaskRects() {
    for (var i = 0; i < originalRects.length; i++) {
        var info = originalRects[i];
        info.mask.top = info.top;
        info.mask.left = info.left;
        info.mask.width = info.width;
        info.mask.height = info.height;
    }
}

/* メイン処理 / Main function */
function main() {
    if (app.documents.length === 0) {
        alert(LABELS.noSelection[lang]);
        return;
    }
    var sel = app.activeDocument.selection;
    if (sel.length === 0) {
        alert(LABELS.noSelection[lang]);
        return;
    }
    var newSelection = collectMaskPaths(sel);
    var marginUnit = getCurrentUnitLabel();
    var defaultMarginValue = '0';
    var margin = showMarginDialog(defaultMarginValue, marginUnit, function(previewMargin) {
        applyTemporaryMarginToMasks(newSelection, previewMargin);
        app.redraw();
    });
    if (margin === null) return;
    app.redraw();
    if (newSelection.length === 0) {
        alert(LABELS.noMask[lang]);
    }
}
main();

/* 矢印キーで値を増減 / Change value by arrow key */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil(value / delta) * delta + delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor(value / delta) * delta - delta;
                event.preventDefault();
            }
        } else {
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        }

        value = Math.round(value);
        editText.text = value;
        if (typeof editText.onChangeValue === "function") {
            editText.onChangeValue(value);
            app.redraw();
        }
    });
}