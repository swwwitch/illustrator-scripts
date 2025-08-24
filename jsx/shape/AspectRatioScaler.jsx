#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

AspectRatioScaler.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択オブジェクトを指定したアスペクト比に合わせて変形します。
- プレビュー対応のダイアログと、縦置き／横置き（Portrait／Landscape）の切り替えに対応。
- 「サイズ＞横幅」に目標幅を入力（現在のルーラー単位：px/mm/pt など）。

### 主な機能：

- 比率プリセット：16:9、1:1、A4、カスタム
- 基準：縦置き／横置き（必要に応じて比率を自動反転）
- 目標幅指定＋単位に応じた丸め（px=整数、mm=0.1mm）
- ピクセルグリッド整合（任意）、アートボードへの変換（任意）

### 処理の流れ：

1) ダイアログで比率・基準・横幅・オプションを設定
2) プレビューでリアルタイム確認
3) ［実行］で選択オブジェクトへ適用（必要に応じてアートボード追加）

### 更新履歴：

- v1.0 (2025-07-20) : 初期バージョン
- v1.1 (2025-07-21) : アートボード変換・カスタム比率を追加
- v1.2 (2025-07-22) : ダイアログ構成・ローカライズ・キー入力を改善
- v1.3 (2025-07-23) : プレビュー安定化、UI微調整、サイズ入力＆丸め処理
- v1.4 (2025-08-24) : ロジック修正、ラベルのローカライズ整理、コメント整備

---

### Script Name:

AspectRatioScaler.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Transforms selected objects to a specified aspect ratio.
- Dialog with live preview and Portrait/Landscape switch.
- Enter target width under "Size" (uses current ruler unit: px/mm/pt, etc.).

### Main Features:

- Ratio presets: 16:9, 1:1, A4, Custom
- Orientation: Portrait/Landscape (auto-inverts ratio when needed)
- Target width with unit-aware rounding (px = integer, mm = 0.1mm)
- Optional: Align to Pixel Grid, Convert to Artboard

### Workflow:

1) Configure ratio, orientation, width, and options in the dialog
2) Preview updates in real time
3) Apply to the selection with [Apply]; optionally add an artboard

### Update History:

- v1.0 (2025-07-20): Initial release
- v1.1 (2025-07-21): Added artboard conversion & custom ratio
- v1.2 (2025-07-22): Improved dialog structure, localization, key input
- v1.3 (2025-07-23): Preview stability, minor UI tweaks, size input & rounding
- v1.4 (2025-08-24): Logic fixes, label localization cleanup, comment pass

*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.4";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    // Dialog title / ダイアログタイトル
    dialogTitle: {
        ja: "アスペクト比で調整",
        en: "Adjust by Aspect Ratio"
    },

    // Aspect panel / アスペクト比パネル
    aspectLabel: {
        ja: "アスペクト比",
        en: "Aspect Ratio"
    },
    ratio169: {
        ja: "16:9",
        en: "16:9"
    },
    ratio11: {
        ja: "1:1（スクエア）",
        en: "1:1"
    },
    ratioA4: {
        ja: "A4（1:1.414）",
        en: "1:1.414"
    },
    ratioCustom: {
        ja: "カスタム",
        en: "Custom"
    },

    // Base (orientation) panel / 基準（向き）パネル
    baseLabel: {
        ja: "基準",
        en: "Base"
    },
    baseWidth: {
        ja: "横置き",
        en: "Landscape"
    },
    baseHeight: {
        ja: "縦置き",
        en: "Portrait"
    },

    // Size panel / サイズパネル
    sizePanel: {
        ja: "サイズ",
        en: "Size"
    },
    labelWidth: {
        ja: "横幅",
        en: "Width"
    },

    // Options / オプション
    alignToPixelGrid: {
        ja: "ピクセルグリッドに最適化",
        en: "Align to Pixel Grid"
    },
    convertToArtboard: {
        ja: "アートボードに変換",
        en: "Convert to Artboard"
    },

    // Buttons / ボタン
    run: {
        ja: "実行",
        en: "Apply"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

// Localization helper
function L(key) {
    try {
        return LABELS[key][lang] || "";
    } catch (e) {
        return "";
    }
}


// Original sizes for preview/apply (global)
var __origW = [];
var __origH = [];

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

// 現在の単位ラベルを取得 / Get current unit label
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

// ルーラー設定に基づく pt 係数を取得 / Get pt factor for current ruler unit
function getPtFactorFromRuler() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    switch (unitCode) {
        case 0:
            /* in   */
            return 72.0;
        case 1:
            /* mm   */
            return 72.0 / 25.4;
        case 2:
            /* pt   */
            return 1.0;
        case 3:
            /* pica */
            return 12.0; // 1pc = 12pt
        case 4:
            /* cm   */
            return 72.0 / 2.54;
        case 5:
            /* Q/H  */
            return (72.0 / 25.4) * 0.25; // Q=0.25mm 相当
        case 6:
            /* px   */
            return 1.0; // 72ppi前提
        case 7:
            /* ft/in*/
            return 72.0; // in と同等扱い
        case 8:
            /* m    */
            return 72.0 / 0.0254;
        case 9:
            /* yd   */
            return 72.0 * 36.0;
        case 10:
            /* ft   */
            return 72.0 * 12.0;
        default:
            return 1.0;
    }
}

// 単位に応じた丸め（px=整数、mm=0.1mm刻み、その他=0.01pt刻み） / Unit-aware rounding
function roundForUnit(valPt) {
    try {
        var unit = app.preferences.getIntegerPreference("rulerType");
        if (unit === 6) { // px -> integer (1px = 1pt assumption)
            return Math.round(valPt);
        }
        if (unit === 1) { // mm -> 0.1mm steps
            var step = (72.0 / 25.4) * 0.1; // 0.1mm in pt
            return Math.round(valPt / step) * step;
        }
    } catch (e) {}
    // default: 0.01pt
    return Math.round(valPt * 100) / 100;
}

// 自動横幅の既定値（選択なしのとき）/ Default auto width when no selection
function getDefaultWidthTextForCurrentUnit() {
    try {
        var unit = app.preferences.getIntegerPreference("rulerType");
        if (unit === 6) { // px
            return "1000";
        }
        if (unit === 1) { // mm
            return "100";
        }
    } catch (e) {}
    return ""; // その他の単位は未指定 / leave empty for other units
}


/* ダイアログ作成 / Create dialog */
function createDialog() {
    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    // UI参照用のローカル変数 / Local variables for UI refs
    var baseWidthRadio, baseHeightRadio;
    var ratio169, ratio11, ratioA4, ratioCustom;
    var editTextWidth, editTextHeight;

    var offsetX = 300;
    var dialogOpacity = 0.97;
    var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);
    dialog.alignChildren = "left";

    var topGroup = dialog.add("group");
    topGroup.orientation = "row";
    topGroup.alignChildren = "left";
    topGroup.alignChildren = ["fill", "top"];

    // 2カラム構成：左=アスペクト、右=基準+サイズ
    // 2-column layout: left = Aspect, right = Base + Size
    var leftCol = topGroup.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["fill", "top"];

    var rightCol = topGroup.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];

    var aspectPanel = leftCol.add("panel", undefined, LABELS.aspectLabel[lang]);
    aspectPanel.orientation = "column";
    aspectPanel.alignChildren = "left";
    aspectPanel.margins = [15, 20, 15, 10];
    aspectPanel.alignment = ["fill", "top"];

    var aspectGroup = aspectPanel.add("group");
    aspectGroup.orientation = "column";
    aspectGroup.alignChildren = "left";
    ratio169 = aspectGroup.add("radiobutton", undefined, LABELS.ratio169[lang]);
    ratio11 = aspectGroup.add("radiobutton", undefined, LABELS.ratio11[lang]);
    ratioA4 = aspectGroup.add("radiobutton", undefined, LABELS.ratioA4[lang]);
    ratioCustom = aspectGroup.add("radiobutton", undefined, LABELS.ratioCustom[lang]);


    var customRatioGroup = aspectPanel.add("group");
    customRatioGroup.orientation = "row";
    customRatioGroup.alignChildren = "left";

    editTextWidth = customRatioGroup.add("edittext", undefined, "3");
    editTextWidth.characters = 5;

    customRatioGroup.add("statictext", undefined, ":");
    editTextHeight = customRatioGroup.add("edittext", undefined, "2");
    editTextHeight.characters = 5;

    editTextWidth.enabled = false;
    editTextHeight.enabled = false;

    changeValueByArrowKey(editTextWidth);
    changeValueByArrowKey(editTextHeight);

    ratio169.value = true;

    var basePanel = rightCol.add("panel", undefined, LABELS.baseLabel[lang]);
    basePanel.orientation = "column";
    basePanel.alignChildren = "left";
    basePanel.margins = [15, 20, 15, 10];
    basePanel.alignment = ["fill", "top"];

    var baseGroup = basePanel.add("group");
    baseGroup.orientation = "column";
    baseGroup.alignChildren = "left";
    baseWidthRadio = baseGroup.add("radiobutton", undefined, LABELS.baseWidth[lang]);
    baseHeightRadio = baseGroup.add("radiobutton", undefined, LABELS.baseHeight[lang]);
    baseWidthRadio.value = true; // default Landscape
    baseHeightRadio.value = false;

    // --- Size panel under Base ---
    var sizePanel = rightCol.add("panel", undefined, LABELS.sizePanel[lang]);
    sizePanel.orientation = "column";
    sizePanel.alignChildren = "left";
    sizePanel.margins = [15, 20, 15, 10];
    sizePanel.alignment = ["fill", "top"];

    var sizeRow = sizePanel.add("group");
    sizeRow.orientation = "row";
    sizeRow.alignChildren = ["left", "center"];

    var stWidthLabel = sizeRow.add("statictext", undefined, LABELS.labelWidth[lang]);
    var etWidthValue = sizeRow.add("edittext", undefined, "");
    etWidthValue.characters = 5; // 少し広め / slightly wider
    var stUnitLabel = sizeRow.add("statictext", undefined, getCurrentUnitLabel());

    var pixelGroup = dialog.add("group");
    pixelGroup.orientation = "column";
    pixelGroup.alignChildren = "left";
    var alignToPixel = pixelGroup.add("checkbox", undefined, LABELS.alignToPixelGrid[lang]);
    var __isPxRuler = false;
    try {
        __isPxRuler = (app.preferences.getIntegerPreference("rulerType") === 6);
    } catch (e) {}
    alignToPixel.value = __isPxRuler; // px時のみON、その他はOFF

    var convertToArtboard = pixelGroup.add("checkbox", undefined, LABELS.convertToArtboard[lang]);
    convertToArtboard.value = false;

    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";

    var btnCancel = buttonGroup.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var btnOk = buttonGroup.add("button", undefined, LABELS.run[lang], {
        name: "ok"
    });

    return {
        dialog: dialog,
        ratio169: ratio169,
        ratio11: ratio11,
        ratioA4: ratioA4,
        ratioCustom: ratioCustom,
        baseVertical: baseHeightRadio,
        baseHorizontal: baseWidthRadio,
        alignToPixel: alignToPixel,
        convertToArtboard: convertToArtboard,
        btnOk: btnOk,
        btnCancel: btnCancel,
        customWidthInput: editTextWidth,
        customHeightInput: editTextHeight,
        sizePanel: sizePanel,
        sizeWidthLabel: stWidthLabel,
        sizeWidthInput: etWidthValue,
        sizeUnitLabel: stUnitLabel,
    };
}

/* メイン処理 / Main function */
function main() {
    /* 選択したオブジェクトを取得する / Get selected objects */
    var selectedItems = app.activeDocument.selection;
    var isNoSelection = (!selectedItems || selectedItems.length === 0);

    // プレビュー用コピーを作成し、元は非表示にする / Create preview copies and hide originals (if any)
    var previewCopies = [];
    __origW = [];
    __origH = [];
    if (!isNoSelection) {
        for (var i = 0; i < selectedItems.length; i++) {
            var dup = selectedItems[i].duplicate();
            dup.hidden = false;
            dup.zOrder(ZOrderMethod.BRINGTOFRONT);
            previewCopies.push(dup);
            selectedItems[i].hidden = true;
            __origW.push(selectedItems[i].width); // 幅を保存 / Save width
            __origH.push(selectedItems[i].height); // 高さを保存 / Save height
        }
    }

    var dialogResult = createDialog();

    // 選択がない場合は横幅に自動入力（px:1000 / mm:100）
    // Auto-fill width when no selection (px:1000, mm:100)
    if (isNoSelection && (!dialogResult.sizeWidthInput.text || dialogResult.sizeWidthInput.text === "")) {
        var autoW = getDefaultWidthTextForCurrentUnit();
        if (autoW !== "") dialogResult.sizeWidthInput.text = autoW;
    }

    function getTargetWidthPt() {
        var txt = dialogResult.sizeWidthInput.text;
        if (!txt) return null;
        var v = parseFloat(txt);
        if (isNaN(v) || v <= 0) return null;
        return v * getPtFactorFromRuler();
    }

    // サイズパネル「横幅」入力のライブプレビュー / Live preview for width field
    dialogResult.sizeWidthInput.onChanging = function() {
        applyAspect(previewCopies, getCurrentRatio(), dialogResult.baseVertical.value, getTargetWidthPt());
    };

    // 何も選択されていない場合は、プレビュー用の長方形を新規作成 / Create a preview rectangle when no selection
    if (isNoSelection) {
        // 初期比率と目標幅を取得
        var initR;
        if (dialogResult.ratio169.value) initR = 1.777777;
        else if (dialogResult.ratio11.value) initR = 1.0;
        else if (dialogResult.ratioA4.value) initR = (210 / 297);
        else {
            var _w = parseFloat(dialogResult.customWidthInput.text);
            var _h = parseFloat(dialogResult.customHeightInput.text);
            initR = (isNaN(_w) || isNaN(_h) || _h === 0) ? 1 : _w / _h;
        }
        var wantPortraitInit = dialogResult.baseVertical.value;
        var rAdj = initR;
        if (wantPortraitInit && rAdj > 1) rAdj = 1 / rAdj;
        if (!wantPortraitInit && rAdj < 1) rAdj = 1 / rAdj;

        var targetW = getTargetWidthPt();
        if (targetW == null) targetW = 200; // デフォルト幅 200pt
        var targetH = roundForUnit(targetW / rAdj);

        var doc = app.activeDocument;
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
        var cx = (ab[0] + ab[2]) / 2;
        var cy = (ab[1] + ab[3]) / 2;
        var left = cx - targetW / 2;
        var top = cy + targetH / 2;

        var rect = doc.pathItems.rectangle(top, left, targetH, targetW);
        rect.stroked = false;
        rect.filled = true;

        previewCopies = [rect];
        __origW = [targetW];
        __origH = [targetH];
    }

    function getCurrentRatio() {
        if (dialogResult.ratio169.value) return 1.777777;
        if (dialogResult.ratio11.value) return 1.0;
        if (dialogResult.ratioA4.value) return (210 / 297);
        var w = parseFloat(dialogResult.customWidthInput.text);
        var h = parseFloat(dialogResult.customHeightInput.text);
        if (isNaN(w) || isNaN(h) || h === 0) return 1;
        return w / h;
    }

    /* アスペクト比選択時のプレビュー更新 / Preview update on aspect ratio selection */
    // 16:9
    dialogResult.ratio169.onClick = function() {
        dialogResult.customWidthInput.enabled = false;
        dialogResult.customHeightInput.enabled = false;
        // 16:9 選択時は「横置き」に切替 / Force Landscape on 16:9
        dialogResult.baseHorizontal.value = true;
        dialogResult.baseVertical.value = false;
        applyAspect(previewCopies, 1.777777, dialogResult.baseVertical.value, getTargetWidthPt());
    };

    // 1:1
    dialogResult.ratio11.onClick = function() {
        dialogResult.customWidthInput.enabled = false;
        dialogResult.customHeightInput.enabled = false;
        applyAspect(previewCopies, 1.0, dialogResult.baseVertical.value, getTargetWidthPt());
    };

    // A4
    dialogResult.ratioA4.onClick = function() {
        dialogResult.customWidthInput.enabled = false;
        dialogResult.customHeightInput.enabled = false;
        // A4 選択時は「縦置き」に切替 / Force Portrait on A4
        dialogResult.baseVertical.value = true;
        dialogResult.baseHorizontal.value = false;
        applyAspect(previewCopies, (210 / 297), dialogResult.baseVertical.value, getTargetWidthPt());
    };

    // カスタム
    dialogResult.ratioCustom.onClick = function() {
        dialogResult.customWidthInput.enabled = true;
        dialogResult.customHeightInput.enabled = true;
        if (dialogResult.ratioCustom.value) {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            var r = (h === 0) ? 1 : w / h;
            applyAspect(previewCopies, r, dialogResult.baseVertical.value, getTargetWidthPt());
        }
    };

    /* カスタム比率入力時のプレビュー更新 / Preview update on custom ratio input */
    dialogResult.customWidthInput.onChanging = function() {
        if (dialogResult.ratioCustom.value) {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            var r = (h === 0) ? 1 : w / h;
            applyAspect(previewCopies, r, dialogResult.baseVertical.value, getTargetWidthPt());

        }
    };
    dialogResult.customHeightInput.onChanging = function() {
        if (dialogResult.ratioCustom.value) {
            var w = parseFloat(dialogResult.customWidthInput.text);
            var h = parseFloat(dialogResult.customHeightInput.text);
            var r = (h === 0) ? 1 : w / h;
            applyAspect(previewCopies, r, dialogResult.baseVertical.value, getTargetWidthPt());
        }
    };

    dialogResult.baseVertical.onClick = function() {
        applyAspect(previewCopies, getCurrentRatio(), true, getTargetWidthPt());

    };
    dialogResult.baseHorizontal.onClick = function() {
        applyAspect(previewCopies, getCurrentRatio(), false, getTargetWidthPt());
    };

    /* 初期プレビュー / Initial preview */
    var initialRatio = dialogResult.ratio169.value ? 1.777777 : (dialogResult.ratio11.value ? 1.0 : (function() {
        var w = parseFloat(dialogResult.customWidthInput.text);
        var h = parseFloat(dialogResult.customHeightInput.text);
        return (h === 0) ? 1 : w / h;
    })());
    applyAspect(previewCopies, initialRatio, dialogResult.baseVertical.value, getTargetWidthPt());

    var result = dialogResult.dialog.show();

    if (result === 1) {
        if (isNoSelection) {
            // 新規作成したプレビュー矩形を最終物として扱う / Keep the preview rectangle as final
            var finalIt = previewCopies[0];
            // ピクセルグリッド整合 / Align to pixel grid (optional)
            if (dialogResult.alignToPixel.value) {
                app.selection = [finalIt];
                app.executeMenuCommand('Make Pixel Perfect');
            }
            // 必要に応じてアートボードを作成 / Convert to artboard if requested
            if (dialogResult.convertToArtboard.value) {
                var vb0 = finalIt.visibleBounds;
                var abRect0 = [vb0[0], vb0[1], vb0[2], vb0[3]];
                app.activeDocument.artboards.add(abRect0);
            }
            app.selection = [finalIt];
            app.redraw();
        } else {
            for (var i = 0; i < selectedItems.length; i++) {
                // 元オブジェクトを再表示 / Unhide original
                selectedItems[i].hidden = false;

                // プレビューの形状を反映（拡大縮小＋位置合わせ） / Apply by scaling and repositioning
                try {
                    var origW = __origW[i];
                    var origH = __origH[i];
                    var prevW = previewCopies[i].width;
                    var prevH = previewCopies[i].height;

                    // 安全ガード / guards
                    if (origW > 0 && origH > 0) {
                        var sx = (prevW / origW) * 100.0;
                        var sy = (prevH / origH) * 100.0;
                        // 中心基準で拡大縮小 / scale from center
                        selectedItems[i].resize(sx, sy);
                    }

                    // 位置合わせ（左上座標） / align position using top-left
                    try {
                        selectedItems[i].position = previewCopies[i].position;
                    } catch (pErr) {}
                } catch (e) {}

                // ピクセルグリッド整合 / Align to pixel grid (optional)
                if (dialogResult.alignToPixel.value) {
                    app.selection = [selectedItems[i]];
                    app.executeMenuCommand('Make Pixel Perfect');
                }

                // プレビューを削除 / Remove preview copy
                try {
                    previewCopies[i].remove();
                } catch (e3) {}
            }

            // 必要に応じてアートボードを作成 / Convert to artboard if requested
            if (dialogResult.convertToArtboard.value) {
                for (var j = 0; j < selectedItems.length; j++) {
                    var it = selectedItems[j];
                    var vb = it.visibleBounds;
                    var abRect = [vb[0], vb[1], vb[2], vb[3]];
                    app.activeDocument.artboards.add(abRect);
                }
            }

            // 最終的にオリジナルを選択状態に / Keep originals selected
            app.selection = selectedItems;
            app.redraw();
        }
    } else {
        // キャンセル時：プレビューを片付け、非表示化を解除 / On cancel, cleanup preview and unhide originals
        for (var i = 0; i < previewCopies.length; i++) {
            try {
                previewCopies[i].remove();
            } catch (e) {}
        }
        if (!isNoSelection) {
            for (var i = 0; i < selectedItems.length; i++) {
                selectedItems[i].hidden = false;
            }
        }
        return;
    }
}

/* アスペクト比適用 / Apply aspect ratio */
function applyAspect(items, ratio, wantPortrait, targetWidthPt) {
    // 向きのガード：縦置き=高さ>幅（ratio<1）、横置き=幅>高さ（ratio>1）
    // Orientation guard: Portrait => height>width (ratio<1), Landscape => width>height (ratio>1)
    var r = ratio;
    if (wantPortrait && r > 1) r = 1 / r;
    if (!wantPortrait && r < 1) r = 1 / r;

    var useTarget = (typeof targetWidthPt === 'number' && isFinite(targetWidthPt) && targetWidthPt > 0);

    for (var i = 0; i < items.length; i++) {
        var w = useTarget ? targetWidthPt : __origW[i];
        items[i].width = w;
        var h = w / r;
        items[i].height = roundForUnit(h);
    }
    app.redraw();
}

/* 上下キーで数値変更を可能にする / Enable arrow key numeric input */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta;

        delta = 1;
        if (event.keyName == "Up") {
            value += delta;
            event.preventDefault();
        } else if (event.keyName == "Down") {
            value -= delta;
            if (value < 0) value = 0;
            event.preventDefault();
        }
        // 整数に丸め / Round to integer
        value = Math.round(value);

        editText.text = value;
        if (typeof editText.onChanging === 'function') {
            try {
                editText.onChanging();
            } catch (e) {}
        }
    });
}

main();