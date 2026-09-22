#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

すべてのアートボードに同じサイズの矩形を描画し、アートボード内のオブジェクトをマスクします。
クリップグループ名はアートボード名に設定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardMaskAndRelease.md

### Overview

Draws a rectangle the size of each artboard and uses it to mask the objects on that artboard.
Each clipping group is named after its artboard.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardMaskAndRelease.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardMaskAndRelease";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardMaskAndRelease.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardMaskAndRelease.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];    /* パネル余白 [左,上,右,下] / panel margins */
    var MARGIN_FIELD_CHARS = 5;              /* マージン欄の文字数 / width of the margin field */

    /**
     * EditText の上下矢印キーで値を増減する（Shift で 10 の倍数にスナップして ±10）
     * Enable up/down arrow key increment/decrement on edittext inputs
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すなら true
     * @returns {void}
     */
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

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードでマスク", en: "Mask Artboards" }
        },
        panel: {
            mode: { ja: "モード", en: "Mode" },
            maskOption: { ja: "マスクオプション", en: "Mask Options" },
            releaseOption: { ja: "解除オプション", en: "Release Options" }
        },
        radio: {
            mask: { ja: "マスク", en: "Mask" },
            release: { ja: "マスク解除", en: "Release Mask" }
        },
        fieldLabel: {
            margin: { ja: "マージン", en: "Margin" }
        },
        checkbox: {
            removeOutside: { ja: "アートボード外のオブジェクトを削除", en: "Remove objects outside artboards" },
            includeLocked: { ja: "ロックされたオブジェクトを含める", en: "Include locked objects" },
            includeHidden: { ja: "非表示のオブジェクトを含める", en: "Include hidden objects" },
            ungroup: { ja: "グループ解除", en: "Ungroup" }
        },
        tooltip: {
            mask: {
                ja: "各アートボードの範囲でクリッピングマスクを作り、はみ出した部分を隠します。",
                en: "Creates a clipping mask at each artboard and hides whatever sticks out."
            },
            release: { ja: "アートボードで作ったクリッピングマスクを外します。", en: "Releases the clipping masks made from the artboards." },
            margin: {
                ja: "マスクの範囲をアートボードより広げる量です。0 でアートボードぴったりになります。",
                en: "How far the mask extends past the artboard. 0 matches the artboard exactly."
            },
            removeOutside: {
                ja: "マスクからはみ出したオブジェクトを、隠すのではなく削除します。",
                en: "Deletes the objects outside the mask instead of hiding them."
            },
            includeLocked: { ja: "ロックされたオブジェクトも処理の対象にします。", en: "Includes locked objects." },
            includeHidden: { ja: "非表示のオブジェクトも処理の対象にします。", en: "Includes hidden objects." },
            ungroup: { ja: "マスクを外したあと、残ったグループも解除します。", en: "Ungroups what is left once the mask is released." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.mode" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - checkbox.* と tooltip.* に共通のキー
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelKey, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        optionCheckbox.helpTip = getLabel("tooltip." + labelKey);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * 縦並びのオプションパネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} titleLabelPath - パネル名のパス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentWindow, titleLabelPath) {
        var optionPanel = parentWindow.add("panel", undefined, getLabel(titleLabelPath));
        optionPanel.orientation = "column";
        optionPanel.alignChildren = "left";
        optionPanel.margins = PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * モード・マスクオプション・解除オプションのダイアログを組み立てる
     * @returns {object} maskDialog と、設定の読み取りに使うコントロール
     */
    function buildMaskDialog() {
        var maskDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        maskDialog.orientation = "column";
        maskDialog.alignChildren = "fill";

        var modePanel = maskDialog.add("panel", undefined, getLabel("panel.mode"));
        modePanel.orientation = "row"; /* 横並びに変更 / Change to horizontal layout */
        modePanel.alignChildren = "left";
        modePanel.margins = PANEL_MARGINS;

        var maskRadio = modePanel.add("radiobutton", undefined, getLabel("radio.mask"));
        maskRadio.helpTip = getLabel("tooltip.mask");
        var releaseRadio = modePanel.add("radiobutton", undefined, getLabel("radio.release"));
        releaseRadio.helpTip = getLabel("tooltip.release");
        maskRadio.value = true;

        /* マスクオプション / Mask options */
        var maskOptionPanel = addOptionPanel(maskDialog, "panel.maskOption");
        var marginRow = maskOptionPanel.add("group");
        marginRow.orientation = "row";
        marginRow.alignChildren = "left";
        marginRow.add("statictext", undefined, labelText("fieldLabel.margin"));

        var marginInput = marginRow.add("edittext", undefined, "0");
        marginInput.helpTip = getLabel("tooltip.margin");
        marginInput.characters = MARGIN_FIELD_CHARS;
        marginInput.active = true;
        changeValueByArrowKey(marginInput, true);
        marginRow.add("statictext", undefined, "(" + getUnitInfo().label + ")");

        var removeOutsideCheckbox = addOptionCheckbox(maskOptionPanel, "removeOutside", false);
        removeOutsideCheckbox.alignment = "left";
        var includeLockedCheckbox = addOptionCheckbox(maskOptionPanel, "includeLocked", true);   /* デフォルトをONに設定 / Default ON */
        var includeHiddenCheckbox = addOptionCheckbox(maskOptionPanel, "includeHidden", true);   /* デフォルトをONに設定 / Default ON */

        /* 解除オプション / Release options */
        var releaseOptionPanel = addOptionPanel(maskDialog, "panel.releaseOption");
        var ungroupCheckbox = addOptionCheckbox(releaseOptionPanel, "ungroup", true);
        releaseOptionPanel.enabled = false;

        /* ラジオボタン切り替え時のパネル有効/無効制御 / Enable/disable panels on radio button toggle */
        maskRadio.onClick = function() {
            maskOptionPanel.enabled = true;
            releaseOptionPanel.enabled = false;
        };
        releaseRadio.onClick = function() {
            maskOptionPanel.enabled = false;
            releaseOptionPanel.enabled = true;
        };

        var btnRowGroup = maskDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            maskDialog: maskDialog,
            maskRadio: maskRadio,
            releaseRadio: releaseRadio,
            marginInput: marginInput,
            removeOutsideCheckbox: removeOutsideCheckbox,
            includeLockedCheckbox: includeLockedCheckbox,
            includeHiddenCheckbox: includeHiddenCheckbox,
            ungroupCheckbox: ungroupCheckbox
        };
    }

    // =========================================
    // マスク / Mask
    // =========================================

    /**
     * 境界がアートボードの矩形と重なるか（接していても重なりとみなす）
     * @param {number[]} itemBounds - オブジェクトの境界 [左, 上, 右, 下]
     * @param {number[]} targetRect - 矩形 [左, 上, 右, 下]
     * @returns {boolean} 重なっていれば true
     */
    function isBoundsOverlapping(itemBounds, targetRect) {
        return !(itemBounds[2] < targetRect[0] || itemBounds[0] > targetRect[2] ||
            itemBounds[3] > targetRect[1] || itemBounds[1] < targetRect[3]);
    }

    /**
     * マスクの対象にする最上位（グループの子でない）オブジェクトを集める
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} includeLocked - ロック中も含めるなら true
     * @param {boolean} includeHidden - 非表示も含めるなら true
     * @returns {PageItem[]} 対象のオブジェクト
     */
    function collectMaskTargets(doc, includeLocked, includeHidden) {
        var targetItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var pageItem = doc.pageItems[i];
            var wasHidden = pageItem.hidden;
            var includeItem = false;

            if (!pageItem.locked || includeLocked) {
                if (!pageItem.hidden || includeHidden) {
                    includeItem = true;
                }
            }

            if (includeItem && pageItem.parent.typename !== "GroupItem") {
                /* hidden の場合、一時的に表示 / Temporarily show if hidden */
                if (pageItem.hidden && includeHidden) {
                    pageItem.hidden = false;
                }
                targetItems.push(pageItem);
            }

            /* 元の hidden 状態に戻す（後の安全のため） / Restore original hidden state for safety */
            if (includeHidden && wasHidden) {
                pageItem.hidden = true;
            }
        }
        return targetItems;
    }

    /**
     * どのアートボードにも重ならないオブジェクトを削除し、配列からも除く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - 対象のオブジェクト（この関数内で要素を減らす）
     * @returns {void}
     */
    function removeItemsOutsideArtboards(doc, targetItems) {
        for (var i = targetItems.length - 1; i >= 0; i--) {
            var pageItem = targetItems[i];
            var isInsideAny = false;
            for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
                if (isBoundsOverlapping(pageItem.visibleBounds, doc.artboards[artboardIndex].artboardRect)) {
                    isInsideAny = true;
                    break;
                }
            }
            if (!isInsideAny) {
                pageItem.remove();
                targetItems.splice(i, 1);
            }
        }
    }

    /**
     * 1つのアートボードの範囲（＋マージン）で、重なるオブジェクトの複製をクリップグループにまとめる
     * @param {Document} doc - 対象ドキュメント
     * @param {Artboard} artboard - 対象のアートボード
     * @param {PageItem[]} targetItems - マスクの対象
     * @param {number} marginPt - アートボードより広げる量
     * @returns {void}
     */
    function maskArtboard(doc, artboard, targetItems, marginPt) {
        var artboardRect = artboard.artboardRect;
        var maskLeft = artboardRect[0] - marginPt;
        var maskTop = artboardRect[1] + marginPt;
        var maskRight = artboardRect[2] + marginPt;
        var maskBottom = artboardRect[3] - marginPt;

        var maskRect = doc.pathItems.rectangle(maskTop, maskLeft, maskRight - maskLeft, maskTop - maskBottom);
        maskRect.stroked = false;
        maskRect.filled = true;
        maskRect.fillColor = new NoColor();

        var duplicatedItems = [];
        for (var i = 0; i < targetItems.length; i++) {
            if (isBoundsOverlapping(targetItems[i].visibleBounds, [maskLeft, maskTop, maskRight, maskBottom])) {
                duplicatedItems.push(targetItems[i].duplicate());
            }
        }

        if (duplicatedItems.length == 0) {
            maskRect.remove();
            return;
        }

        var clipGroup = doc.groupItems.add();
        for (var j = 0; j < duplicatedItems.length; j++) {
            duplicatedItems[j].moveToBeginning(clipGroup);
        }

        maskRect.moveToBeginning(clipGroup);
        clipGroup.clipped = true;
        clipGroup.name = artboard.name;
    }

    /**
     * すべてのアートボードでマスクを作り、元のオブジェクトを削除する
     * @param {number} marginPt - アートボードより広げる量
     * @param {boolean} removeOutside - アートボード外のオブジェクトを削除するなら true
     * @param {boolean} includeLocked - ロック中も含めるなら true
     * @param {boolean} includeHidden - 非表示も含めるなら true
     * @returns {void}
     */
    function applyMasks(marginPt, removeOutside, includeLocked, includeHidden) {
        var doc = app.activeDocument;
        var targetItems = collectMaskTargets(doc, includeLocked, includeHidden);

        if (removeOutside) {
            removeItemsOutsideArtboards(doc, targetItems);
        }

        for (var i = 0; i < doc.artboards.length; i++) {
            maskArtboard(doc, doc.artboards[i], targetItems, marginPt);
        }

        for (var j = targetItems.length - 1; j >= 0; j--) {
            targetItems[j].remove();
        }
    }

    // =========================================
    // 解除 / Release
    // =========================================

    /**
     * クリップグループのマスクを解除し、マスクパスと塗り・線なしのパスを削除する
     * @param {boolean} shouldUngroup - 残ったグループも解除するなら true
     * @returns {void}
     */
    function releaseMasks(shouldUngroup) {
        var doc = app.activeDocument;
        for (var i = doc.groupItems.length - 1; i >= 0; i--) {
            var clipGroup = doc.groupItems[i];
            if (clipGroup.clipped) {
                clipGroup.clipped = false;
                for (var j = clipGroup.pageItems.length - 1; j >= 0; j--) {
                    var childItem = clipGroup.pageItems[j];
                    if (childItem.clipping) {
                        childItem.remove();
                    }
                }

                for (var k = clipGroup.pageItems.length - 1; k >= 0; k--) {
                    var remainingItem = clipGroup.pageItems[k];
                    if (remainingItem.typename === "PathItem" && remainingItem.filled == false && remainingItem.stroked == false) {
                        remainingItem.remove();
                    }
                }

                /* グループ内にオブジェクトが残っている場合、チェック時に解除 / Ungroup if checkbox checked and objects remain */
                if (shouldUngroup) {
                    clipGroup.selected = true;
                    app.executeMenuCommand("ungroup");
                } else if (clipGroup.pageItems.length == 0) {
                    clipGroup.remove();
                }
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで選んだモードでマスクの作成または解除を行う
     * @returns {void}
     */
    function main() {

        if (app.documents.length == 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var dialogControls = buildMaskDialog();
        if (dialogControls.maskDialog.show() != 1) {
            return; /* キャンセル時 / Cancel */
        }

        var marginValue = parseFloat(dialogControls.marginInput.text);
        if (isNaN(marginValue)) {
            marginValue = 0;
        }

        if (dialogControls.maskRadio.value) {
            applyMasks(marginValue, dialogControls.removeOutsideCheckbox.value,
                dialogControls.includeLockedCheckbox.value, dialogControls.includeHiddenCheckbox.value);
        } else if (dialogControls.releaseRadio.value) {
            releaseMasks(dialogControls.ungroupCheckbox.value);
        }
    }

    main();

})();
