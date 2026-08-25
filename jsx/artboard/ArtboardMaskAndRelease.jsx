#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

すべてのアートボードに同じサイズの矩形を描画し、アートボード内のオブジェクトをマスクします。
クリップグループ名はアートボード名に設定します。

詳細は README を参照してください。

### Overview

Draws a rectangle the size of each artboard and uses it to mask the objects on that artboard.
Each clipping group is named after its artboard.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardMaskAndRelease";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-10";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardMaskAndRelease.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardMaskAndRelease.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    // -------------------------------
    // 日英ラベル定義 / Japanese-English label definitions
    // -------------------------------

    var LABELS = {
        dialogTitle: {
            ja: "アートボードでマスク " + SCRIPT_VERSION,
            en: "Mask Artboards " + SCRIPT_VERSION
        },
        modePanel: {
            ja: "モード",
            en: "Mode"
        },
        mask: {
            ja: "マスク",
            en: "Mask"
        },
        release: {
            ja: "解除",
            en: "Release"
        },
        maskOption: {
            ja: "マスクオプション",
            en: "Mask Options"
        },
        margin: {
            ja: "マージン",
            en: "Margin"
        },
        releaseOption: {
            ja: "解除オプション",
            en: "Release Options"
        },
        ungroup: {
            ja: "グループ解除",
            en: "Ungroup"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        noDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        removeOutside: {
            ja: "アートボード外のオブジェクトを削除",
            en: "Remove objects outside artboards"
        },
        includeLocked: {
            ja: "ロックされたオブジェクトを含める",
            en: "Include locked objects"
        },
        includeHidden: {
            ja: "非表示のオブジェクトを含める",
            en: "Include hidden objects"
        },
        ungroupLabel: {
            ja: "グループ解除",
            en: "Ungroup"
        },
        maskRelease: {
            ja: "マスク解除",
            en: "Release Mask"
        }
    };

    function main() {

        if (app.documents.length == 0) {
            alert(LABELS.noDocument[lang]);
            return;
        }

        var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        var panel = dialog.add("panel", undefined, LABELS.modePanel[lang]);
        panel.orientation = "row"; /* 横並びに変更 / Change to horizontal layout */
        panel.alignChildren = "left";
        panel.margins = [15, 20, 15, 10];

        var rbMask = panel.add("radiobutton", undefined, LABELS.mask[lang]);
        var rbRelease = panel.add("radiobutton", undefined, LABELS.maskRelease[lang]);
        rbMask.value = true;

        // ラジオボタン切り替え時のパネル有効/無効制御 / Enable/disable panels on radio button toggle
        rbMask.onClick = function() {
            marginGroup.enabled = true;
            releasePanel.enabled = false;
        };
        rbRelease.onClick = function() {
            marginGroup.enabled = false;
            releasePanel.enabled = true;
        };

        var marginGroup = dialog.add("panel", undefined, LABELS.maskOption[lang]);
        marginGroup.orientation = "column";
        marginGroup.alignChildren = "left";
        marginGroup.margins = [15, 20, 15, 10];
        var marginRow = marginGroup.add("group");
        marginRow.orientation = "row";
        marginRow.alignChildren = "left";
        marginRow.add("statictext", undefined, LABELS.margin[lang] + ":");
        // --- 単位ラベル追加 / Add unit label ---
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

        // Enable up/down arrow key increment/decrement on edittext inputs / EditTextの上下矢印キーで値を増減
        function changeValueByArrowKey(editText, allowNegative) {
            editText.addEventListener("keydown", function(event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;

                if (event.keyName == "Up" || event.keyName == "Down") {
                    var isUp = event.keyName == "Up";
                    var delta = 1;

                    if (keyboard.shiftKey) {
                        // 10の倍数にスナップ / Snap to multiples of 10
                        value = Math.floor(value / 10) * 10;
                        delta = 10;
                    }

                    value += isUp ? delta : -delta;

                    // 負数許可されない場合は0未満を禁止 / Disallow negative if not allowed
                    if (!allowNegative && value < 0) value = 0;

                    event.preventDefault();
                    editText.text = value;
                }
            });
        }

        var marginInput = marginRow.add("edittext", undefined, "0");
        marginInput.characters = 5;
        marginInput.active = true;
        changeValueByArrowKey(marginInput, true);
        var unitLabel = getCurrentUnitLabel();
        marginRow.add("statictext", undefined, "(" + unitLabel + ")");

        var cbRemoveOutside = marginGroup.add("checkbox", undefined, LABELS.removeOutside[lang]);
        cbRemoveOutside.alignment = "left";

        var cbIncludeLocked = marginGroup.add("checkbox", undefined, LABELS.includeLocked[lang]);
        cbIncludeLocked.value = true; /* デフォルトをONに設定 / Default ON */

        // チェックボックス追加 / Add checkbox
        var cbIncludeHidden = marginGroup.add("checkbox", undefined, LABELS.includeHidden[lang]);
        cbIncludeHidden.value = true; /* デフォルトをONに設定 / Default ON */

        var releasePanel = dialog.add("panel", undefined, LABELS.releaseOption[lang]);
        releasePanel.orientation = "column";
        releasePanel.alignChildren = "left";
        releasePanel.margins = [15, 20, 15, 10];

        var cbUngroup = releasePanel.add("checkbox", undefined, LABELS.ungroupLabel[lang]);
        cbUngroup.value = true;
        releasePanel.enabled = false;

        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "right";

        var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang], {
            name: "cancel"
        });
        var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang], {
            name: "ok"
        });

        var result = dialog.show();

        if (result != 1) {
            return; /* キャンセル時 / Cancel */
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

    function applyMasks(margin, removeOutside, includeLocked, includeHidden) {
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
                /* hidden の場合、一時的に表示 / Temporarily show if hidden */
                if (item.hidden && includeHidden) {
                    item.hidden = false;
                }
                allItems.push(item);
            }

            /* 元の hidden 状態に戻す（後の安全のため） / Restore original hidden state for safety */
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

            var abLeft = abRect[0] - margin;
            var abTop = abRect[1] + margin;
            var abRight = abRect[2] + margin;
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

    function releaseMasks(ungroup) {
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

                /* グループ内にオブジェクトが残っている場合、チェック時に解除 / Ungroup if checkbox checked and objects remain */
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

})();
