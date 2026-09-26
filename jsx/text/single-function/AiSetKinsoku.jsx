#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストに、禁則処理のプリセットを適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSetKinsoku.md

### Overview

Applies a kinsoku (line-breaking) preset to the selected text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSetKinsoku.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiSetKinsoku";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSetKinsoku.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSetKinsoku.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 禁則プリセット（表示順）。kinsokuName は paragraphAttributes.kinsoku に渡す値、labelPath は表示名の LABELS キー
       Kinsoku presets in display order: kinsokuName goes to paragraphAttributes.kinsoku, labelPath names the LABELS entry */
    var KINSOKU_PRESETS = [
        { kinsokuName: "None",    labelPath: "radio.none" },
        { kinsokuName: "Hard",    labelPath: "radio.hard" },
        { kinsokuName: "Soft",    labelPath: "radio.soft" },
        { kinsokuName: "Soft_v2", labelPath: "radio.softV2" }
    ];

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins [L,T,R,B] */
    var PANEL_SPACING = 6;                /* パネル内の要素間隔 / spacing inside the panel */
    var BUTTON_SPACING = 8;               /* ボタンの間隔 / spacing between buttons */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "禁則処理の設定", en: "Kinsoku Settings" }
        },
        panel: {
            preset: { ja: "禁則処理", en: "Kinsoku" }
        },
        radio: {
            none: { ja: "なし", en: "None" },
            hard: { ja: "強い禁則", en: "Hard" },
            soft: { ja: "弱い禁則", en: "Soft" },
            softV2: { ja: "弱い禁則 v2", en: "Soft v2" }
        },
        tooltip: {
            preset: {
                ja: "クリックすると、選択中のテキストにすぐ適用します",
                en: "Applies to the selected text as soon as you click"
            },
            close: {
                ja: "適用済みの禁則処理はそのままにして閉じます",
                en: "Closes the dialog and keeps the kinsoku already applied"
            }
        },
        button: {
            close: { ja: "閉じる", en: "Close" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Select text objects." }
        }
    };

    /**
     * ドット区切りのキーから現在の UI 言語のラベルを返す
     * @param {string} labelPath - "dialog.title" のようなキー
     * @returns {string} 現在の UI 言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // 禁則の適用 / Applying kinsoku
    // =========================================

    /**
     * テキストフレームに禁則を適用する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} kinsokuName - paragraphAttributes.kinsoku に渡す値
     * @returns {void}
     */
    function applyKinsokuToTextFrame(textFrame, kinsokuName) {
        /* 「なし」= "None" はスクリプトから設定できず例外になるので、そのまま続行する
           "None" cannot be set from a script and throws; carry on */
        try {
            textFrame.textRange.paragraphAttributes.kinsoku = kinsokuName;
        } catch (e) {}
    }

    /**
     * テキストフレームには直接、グループには中身へ再帰して禁則を適用する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {string} kinsokuName - paragraphAttributes.kinsoku に渡す値
     * @returns {void}
     */
    function applyKinsokuToItem(pageItem, kinsokuName) {
        if (pageItem.typename === "TextFrame") {
            applyKinsokuToTextFrame(pageItem, kinsokuName);
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                applyKinsokuToItem(pageItem.pageItems[i], kinsokuName);
            }
        }
    }

    /**
     * 選択中のオブジェクトすべてに禁則を適用して再描画する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} kinsokuName - paragraphAttributes.kinsoku に渡す値
     * @returns {void}
     */
    function applyKinsokuToSelection(doc, kinsokuName) {
        for (var i = 0; i < doc.selection.length; i++) {
            applyKinsokuToItem(doc.selection[i], kinsokuName);
        }
        app.redraw();
    }

    // =========================================
    // 現在値の読み取り / Reading the current value
    // =========================================

    /**
     * テキストフレームの現在の禁則を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {string} 禁則の値。「なし」のときは ""
     */
    function getKinsoku(textFrame) {
        /* 禁則「なし」の段落は getter が Error 9563 を投げる / The getter throws Error 9563 on "None" paragraphs */
        try {
            return textFrame.textRange.paragraphAttributes.kinsoku;
        } catch (e) {
            return "";
        }
    }

    /**
     * オブジェクト（グループ内を含む）から最初のテキストフレームを探す
     * @param {PageItem} pageItem - 探す対象
     * @returns {TextFrame|null} 見つかったテキストフレーム。無ければ null
     */
    function findFirstTextFrame(pageItem) {
        if (pageItem.typename === "TextFrame") return pageItem;
        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                var foundTextFrame = findFirstTextFrame(pageItem.pageItems[i]);
                if (foundTextFrame !== null) return foundTextFrame;
            }
        }
        return null;
    }

    /**
     * 選択の中で最初に見つかるテキストフレームの禁則を返す（初期選択の判定用）
     * @param {Document} doc - 対象のドキュメント
     * @returns {string} 禁則の値。テキストフレームが無いか「なし」のときは ""
     */
    function getSelectionKinsoku(doc) {
        for (var i = 0; i < doc.selection.length; i++) {
            var firstTextFrame = findFirstTextFrame(doc.selection[i]);
            if (firstTextFrame !== null) return getKinsoku(firstTextFrame);
        }
        return "";
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 禁則プリセットのラジオを並べ、クリックで即適用するようにする
     * @param {Window} kinsokuDialog - 追加先のダイアログ
     * @param {Document} doc - 対象のドキュメント
     * @param {string} currentKinsoku - 初期選択にする禁則の値
     * @returns {RadioButton[]} 作成したラジオボタン
     */
    function addPresetPanel(kinsokuDialog, doc, currentKinsoku) {
        var presetPanel = kinsokuDialog.add("panel", undefined, getLabel("panel.preset"));
        presetPanel.orientation = "column";
        presetPanel.alignChildren = "left";
        presetPanel.alignment = "fill";
        presetPanel.margins = PANEL_MARGINS;
        presetPanel.spacing = PANEL_SPACING;

        var presetRadios = [];
        for (var i = 0; i < KINSOKU_PRESETS.length; i++) {
            var presetRadio = presetPanel.add("radiobutton", undefined, getLabel(KINSOKU_PRESETS[i].labelPath));
            presetRadio.kinsokuName = KINSOKU_PRESETS[i].kinsokuName;
            presetRadio.helpTip = getLabel("tooltip.preset");
            presetRadio.onClick = function () {
                applyKinsokuToSelection(doc, this.kinsokuName);
            };
            presetRadios.push(presetRadio);
        }

        /* 現在の禁則に一致するラジオを選ぶ（一致しなければ先頭＝「なし」）/ Check the matching radio, or the first ("None") */
        var checkedRadio = presetRadios[0];
        for (var j = 0; j < presetRadios.length; j++) {
            if (presetRadios[j].kinsokuName === currentKinsoku) {
                checkedRadio = presetRadios[j];
                break;
            }
        }
        checkedRadio.value = true;

        return presetRadios;
    }

    /**
     * ［閉じる］［OK］のボタン行を追加する
     * @param {Window} kinsokuDialog - 追加先のダイアログ
     * @param {Document} doc - 対象のドキュメント
     * @param {RadioButton[]} presetRadios - プリセットのラジオボタン
     * @returns {void}
     */
    function addButtonRow(kinsokuDialog, doc, presetRadios) {
        var btnRowGroup = kinsokuDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["left", "center"];
        btnRowGroup.alignment = "right";
        btnRowGroup.spacing = BUTTON_SPACING;

        var btnClose = btnRowGroup.add("button", undefined, getLabel("button.close"));
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));
        btnClose.helpTip = getLabel("tooltip.close");

        btnClose.onClick = function () {
            kinsokuDialog.close(0);
        };

        btnOK.onClick = function () {
            for (var i = 0; i < presetRadios.length; i++) {
                if (presetRadios[i].value) {
                    applyKinsokuToSelection(doc, presetRadios[i].kinsokuName);
                    break;
                }
            }
            kinsokuDialog.close(1);
        };
    }

    /**
     * 禁則処理を選ぶダイアログを表示する
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function showKinsokuDialog(doc) {
        var kinsokuDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        kinsokuDialog.orientation = "column";
        kinsokuDialog.alignChildren = "fill";

        var presetRadios = addPresetPanel(kinsokuDialog, doc, getSelectionKinsoku(doc));
        addButtonRow(kinsokuDialog, doc, presetRadios);

        kinsokuDialog.center();
        kinsokuDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確かめてダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        showKinsokuDialog(doc);
    }

    main();

})();
