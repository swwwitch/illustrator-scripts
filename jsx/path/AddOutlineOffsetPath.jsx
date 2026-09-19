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
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddOutlineOffsetPath.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddOutlineOffsetPath.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */

    var LABELS = {
      dialogTitle: { ja: "アウトライン オフセットパス " + SCRIPT_VERSION, en: "Outline Offset Path " + SCRIPT_VERSION },
      offsetPanelTitle: { ja: "オフセット設定", en: "Offset settings" },
      joinPanelTitle: { ja: "角の形状", en: "Corner" },
      tipOffset: { ja: "元のパスから外側（マイナスで内側）へ離す距離です。", en: "How far the new path sits outside the original. A negative value goes inside." },
      tipJoinMiter: { ja: "角を尖らせたまま結合します。鋭角では飛び出すことがあります。", en: "Keeps the corners pointed. Sharp angles can spike out." },
      tipJoinRound: { ja: "角を丸めて結合します。", en: "Rounds off the corners." },
      tipJoinBevel: { ja: "角を面取りして結合します。", en: "Cuts the corners off flat." },
      joinMiter: { ja: "マイター結合", en: "Miter Join" },
      joinRound: { ja: "ラウンド結合", en: "Round Join" },
      joinBevel: { ja: "ベベル結合", en: "Bevel Join" },
      cancel: { ja: "キャンセル", en: "Cancel" },
      ok: { ja: "OK", en: "OK" },
      alertNoSelection: { ja: "オブジェクトが選択されていません。", en: "Please select at least one object." },
      alertEnterNumeric: { ja: "数値を入力してください。", en: "Enter a numeric value." }
    };

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
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
            alert(LABELS.alertNoSelection[uiLang]);
            return;
        }
        /* オフセット入力ダイアログ / Offset input dialog */
        var rulerUnit = getUnitInfo("rulerType");
        var unitLabel = rulerUnit.label;
        var ptFactor = rulerUnit.pointsPerUnit;

        var selWidthPt = getSelectionWidthPt(doc.selection, true); // true => visibleBounds
        var defaultOffsetInCurrentUnit = Math.max(0, Math.round(((selWidthPt / 30) / ptFactor) * 10) / 10);

        var dlg = new Window("dialog", LABELS.dialogTitle[uiLang]);
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

        var offsetPanel = dlg.add("panel", undefined, LABELS.offsetPanelTitle[uiLang]);
        offsetPanel.orientation = "row";
        offsetPanel.alignChildren = ["left", "center"];
        offsetPanel.margins = [15, 20, 15,10]
        var et = offsetPanel.add("edittext", undefined, String(defaultOffsetInCurrentUnit));
        et.helpTip = LABELS.tipOffset[uiLang];
        var editTextWidth = et;
        changeValueByArrowKey(editTextWidth);
        offsetPanel.add("statictext", undefined, unitLabel);
        et.characters = 4;
        et.active = true;

        // Join type (jntp) radio buttons
        var joinGroup = dlg.add("panel", undefined, LABELS.joinPanelTitle[uiLang]);
        joinGroup.orientation = "column";
        joinGroup.alignChildren = ["left", "center"];
        joinGroup.margins = [15, 20, 15,10]
        var rbMiter = joinGroup.add("radiobutton", undefined, LABELS.joinMiter[uiLang]);
        rbMiter.helpTip = LABELS.tipJoinMiter[uiLang];
        var rbRound = joinGroup.add("radiobutton", undefined, LABELS.joinRound[uiLang]);
        rbRound.helpTip = LABELS.tipJoinRound[uiLang];
        var rbBevel = joinGroup.add("radiobutton", undefined, LABELS.joinBevel[uiLang]);
        rbBevel.helpTip = LABELS.tipJoinBevel[uiLang];
        rbMiter.value = false;
        rbRound.value = true;  // default = Round
        rbBevel.value = false;

        var btns = dlg.add("group");
        btns.orientation = "row";
        btns.alignment = ["center", "center"];
        btns.spacing = 10;
        var btnCancel = btns.add("button", undefined, LABELS.cancel[uiLang], {
            name: "cancel"
        });
        var btnOK = btns.add("button", undefined, LABELS.ok[uiLang], {
            name: "ok"
        });
        if (dlg.show() !== 1) {
            return;
        }
        var offsetValueRaw = parseFloat(et.text);
        if (isNaN(offsetValueRaw)) {
            alert(LABELS.alertEnterNumeric[uiLang]);
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

})();
