#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを複製して背面に置き、オフセットパス→アウトライン→合体→拡張の順で処理して縁取りを作ります。
元オブジェクトと結果をグループ化し、Subtract を実行して白で塗りつぶします。

詳細は README を参照してください。

### Overview

Duplicates the selection behind itself and runs Offset Path, Outline Stroke, Unite and Expand to build an outline around it.
The original and the result are grouped, then Subtract is run and the result filled with white.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddOutlineOffsetPath";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-08-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddOutlineOffsetPath.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddOutlineOffsetPath.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
  dialogTitle: { ja: "アウトライン オフセットパス " + SCRIPT_VERSION, en: "Outline Offset Path " + SCRIPT_VERSION },
  offsetPanelTitle: { ja: "オフセット設定", en: "Offset settings" },
  joinPanelTitle: { ja: "角の形状", en: "Corner" },
  joinMiter: { ja: "マイター結合", en: "Miter Join" },
  joinRound: { ja: "ラウンド結合", en: "Round Join" },
  joinBevel: { ja: "ベベル結合", en: "Bevel Join" },
  cancel: { ja: "キャンセル", en: "Cancel" },
  ok: { ja: "OK", en: "OK" },
  alertNoSelection: { ja: "オブジェクトが選択されていません。", en: "Please select at least one object." },
  alertEnterNumeric: { ja: "数値を入力してください。", en: "Enter a numeric value." }
};

/* 単位サポート / Unit support */
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

function getCurrentUnitLabel() {
  var unitCode = app.preferences.getIntegerPreference("rulerType");
  return unitLabelMap[unitCode] || "pt";
}

function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0: return 72.0;                        // in
        case 1: return 72.0 / 25.4;                 // mm
        case 2: return 1.0;                         // pt
        case 3: return 12.0;                        // pica
        case 4: return 72.0 / 2.54;                 // cm
        case 5: return 72.0 / 25.4 * 0.25;          // Q or H
        case 6: return 1.0;                         // px (ドキュメント設定に依存)
        case 7: return 72.0 * 12.0;                 // ft/in
        case 8: return 72.0 / 25.4 * 1000.0;        // m
        case 9: return 72.0 * 36.0;                 // yd
        case 10: return 72.0 * 12.0;                // ft
        default: return 1.0;
    }
}

function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            // 小数第1位までに丸め
            value = Math.round(value * 10) / 10;
        } else {
            // 整数に丸め
            value = Math.round(value);
        }

        editText.text = value;
    });
}

function getSelectionWidthPt(items, useVisible) {
    var left = null, right = null;
    for (var i = 0; i < items.length; i++) {
        try {
            var b = useVisible ? items[i].visibleBounds : items[i].geometricBounds; // [left, top, right, bottom]
            if (!b) continue;
            if (left === null || b[0] < left) left = b[0];
            if (right === null || b[2] > right) right = b[2];
        } catch (e) {}
    }
    if (left === null || right === null) return 0;
    return right - left;
}

function main() {
    var doc = app.activeDocument;
    /* 未選択時のガード / Guard when nothing is selected */
    if (!doc || doc.selection.length === 0) {
        alert(LABELS.alertNoSelection[lang]);
        return;
    }
    /* オフセット入力ダイアログ / Offset input dialog */
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var unitLabel = getCurrentUnitLabel();
    var ptFactor = getPtFactorFromUnitCode(unitCode);

    var selWidthPt = getSelectionWidthPt(doc.selection, true); // true => visibleBounds
    var defaultOffsetInCurrentUnit = Math.max(0, Math.round(((selWidthPt / 30) / ptFactor) * 10) / 10);

    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.alignChildren = ["left", "top"];

    /* ダイアログの位置と透明度 / Dialog position & opacity */
    var offsetX = 20;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            try {
                var loc = dlg.location;
                dlg.location = [loc[0] + offsetX, loc[1] + offsetY];
            } catch (e) {}
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        try {
            dlg.opacity = opacityValue;
        } catch(e) {
            // opacity not supported on some platforms
        }
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, 0);

    var offsetPanel = dlg.add("panel", undefined, LABELS.offsetPanelTitle[lang]);
    offsetPanel.orientation = "row";
    offsetPanel.alignChildren = ["left", "center"];
    offsetPanel.margins = [15, 20, 15,10]
    var et = offsetPanel.add("edittext", undefined, String(defaultOffsetInCurrentUnit));
    var editTextWidth = et;
    changeValueByArrowKey(editTextWidth);
    offsetPanel.add("statictext", undefined, unitLabel);
    et.characters = 4;
    et.active = true;

    // Join type (jntp) radio buttons
    var joinGroup = dlg.add("panel", undefined, LABELS.joinPanelTitle[lang]);
    joinGroup.orientation = "column";
    joinGroup.alignChildren = ["left", "center"];
    joinGroup.margins = [15, 20, 15,10]
    var rbMiter = joinGroup.add("radiobutton", undefined, LABELS.joinMiter[lang]);
    var rbRound = joinGroup.add("radiobutton", undefined, LABELS.joinRound[lang]);
    var rbBevel = joinGroup.add("radiobutton", undefined, LABELS.joinBevel[lang]);
    rbMiter.value = false;
    rbRound.value = true;  // default = Round
    rbBevel.value = false;

    var btns = dlg.add("group");
    btns.orientation = "row";
    btns.alignment = ["center", "center"];
    btns.spacing = 10;
    var btnCancel = btns.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var btnOK = btns.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });
    if (dlg.show() !== 1) {
        return;
    }
    var offsetValueRaw = parseFloat(et.text);
    if (isNaN(offsetValueRaw)) {
        alert(LABELS.alertEnterNumeric[lang]);
        return;
    }
    var offsetValuePt = offsetValueRaw * ptFactor; // Convert to pt
    // joinTypes: 0 = Round, 1 = Bevel , 2 = Miter
    var jntp = rbRound.value ? 0 : (rbBevel.value ? 1 : 2);
    addOutline(doc.selection, offsetValuePt, jntp);
}

function addOutline(items, offsetValue, jntp) {
    var doc = app.activeDocument;
    var prevUIL = app.userInteractionLevel;

    function fillAllToWhite(node) {
        var white = new RGBColor();
        white.red = 255; white.green = 255; white.blue = 255;
        if (!node) return;
        try {
            if (node.typename === "PathItem") {
                node.filled = true; node.fillColor = white;
            } else if (node.typename === "CompoundPathItem") {
                for (var i = 0; i < node.pathItems.length; i++) {
                    node.pathItems[i].filled = true;
                    node.pathItems[i].fillColor = white;
                }
            } else if (node.typename === "GroupItem") {
                for (var j = 0; j < node.pageItems.length; j++) {
                    fillAllToWhite(node.pageItems[j]);
                }
            }
        } catch (e) {}
    }

    /* アラート抑止（終了時に復元）/ Suppress alerts (restore on exit) */
    try {
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var originals = [].concat(items);
        for (var idx = 0; idx < originals.length; idx++) {
            var original = originals[idx];
            if (!original) continue;
            try {
                // Reset selection for this cycle
                doc.selection = null;

                /* 複製して背面へ / Duplicate and send to back */
                var dup = original.duplicate();
                dup.zOrder(ZOrderMethod.SENDTOBACK);
                doc.selection = [dup];

                /* オフセットを Live Effect で適用 / Apply Offset Path (Live Effect) */
                var xmlstring = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetValue + ' I jntp ' + jntp + ' "/><\/LiveEffect>';
                doc.selection[0].applyEffect(xmlstring);

                /* アウトライン→合体→拡張 / Outline → Unite → Expand */
                app.executeMenuCommand("Live Outline Stroke");
                app.executeMenuCommand("Live Pathfinder Add");
                app.executeMenuCommand("expandStyle");
                app.executeMenuCommand('noCompoundPath');
                app.executeMenuCommand("Live Pathfinder Add");
                app.executeMenuCommand("expandStyle");

                /* 元と結果をグループ→Subtract→白塗り / Group original & result → Subtract → Fill white */
                var resultItem = (doc.selection && doc.selection.length) ? doc.selection[0] : null;
                if (resultItem) {
                    doc.selection = [original, resultItem];
                    app.executeMenuCommand("group");
                    app.executeMenuCommand('Live Pathfinder Subtract');
                    app.executeMenuCommand("expandStyle");

                    // keep group; fill recursively to white
                    if (doc.selection && doc.selection.length) {
                        resultItem = doc.selection[0];
                        doc.selection = [resultItem];
                        fillAllToWhite(resultItem);
                    }
                }
            } catch (e) {
                // skip this item and continue
            }
        }
    } finally {
        app.userInteractionLevel = prevUIL;
    }
}

main();
