#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ApplyMojikumi
Name=文字組みアキ量設定を適用
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-06-23
Description-BEGIN
選択したテキストフレーム（グループ内も再帰的に処理）の全段落に、文字組みアキ量設定をダイアログボックスのラジオボタンから適用します。
ラジオボタンを選ぶと、選択中のテキストへすぐに反映されます。組み込みの8種類に加えて、ドキュメントのカスタムのセットも選べます。
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したテキストフレーム（グループ内も再帰的に処理）の全段落に、文字組みアキ量設定をダイアログボックスのラジオボタンから適用します。
ラジオボタンを選ぶと、選択中のテキストへすぐに反映されます。組み込みの8種類に加えて、ドキュメントのカスタムのセットも選べます。

note記事も参照してください。
https://note.com/dtp_tranist/n/nb9c089e378d2

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyMojikumi";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-16";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "文字組みアキ量設定を適用"; /* 内容（README なし） / what it does (no README) */

var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb9c089e378d2"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/ApplyMojikumi.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 組み込みの文字組みプリセット（mojikumiSet のインデックス。-1 は「なし」）。組み込みの要素は内部 ID の文字列なので、表示名はここに持つ
       Built-in mojikumi presets (mojikumiSet index; -1 = none). Built-in entries are internal ID strings, so the names live here */
    var mojikumiPresets = [
        { mojikumiIndex: -1, labelText: "なし" },
        { mojikumiIndex: 0, labelText: "行末約物全角/半角" },
        { mojikumiIndex: 1, labelText: "約物半角" },
        { mojikumiIndex: 2, labelText: "行末約物半角" },
        { mojikumiIndex: 3, labelText: "行末約物全角" },
        { mojikumiIndex: 4, labelText: "約物全角" },
        { mojikumiIndex: 5, labelText: "ツメ組み" },
        { mojikumiIndex: 6, labelText: "ベタ組み" }
    ];

    /* 組み込みの数（mojikumiSet のインデックス7以降はカスタムのセットで、要素の文字列がそのまま名前）
       Number of built-in sets (index 7 and later are custom sets whose string is the name) */
    var MOJIKUMI_BUILTIN_COUNT = 7;

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 前提チェック：ドキュメントが開かれているか / Precondition: a document is open */
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    /* 前提チェック：オブジェクトが選択されているか / Precondition: something is selected */
    if (doc.selection.length === 0) {
        alert("テキストオブジェクトを選択してください。");
        return;
    }

    /**
     * 組み込みのプリセットのあとに、ドキュメントのカスタムのセットを足した選択肢を返す
     * @returns {Array<{mojikumiIndex: number, labelText: string}>} 選択肢
     */
    function buildMojikumiChoices() {
        var choices = mojikumiPresets.slice(0);
        var mojikumiSets = doc.mojikumiSet;
        for (var i = MOJIKUMI_BUILTIN_COUNT; i < mojikumiSets.length; i++) {
            choices.push({ mojikumiIndex: i, labelText: String(mojikumiSets[i]) });
        }
        return choices;
    }

    /**
     * テキストフレーム内の全段落に適用する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {*} mojikumiValue - 代入する値（mojikumiSet の要素、または "なし"）
     * @returns {void}
     */
    function applyMojikumiToTextFrame(textFrame, mojikumiValue) {
        var paragraphs = textFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            /* 末尾の空の段落は取り出すと例外になる / A trailing empty paragraph throws when accessed */
            try { paragraphs[i].paragraphAttributes.mojikumi = mojikumiValue; } catch (e) { }
        }
    }

    /**
     * 種類で振り分ける（TextFrame は直接、GroupItem は再帰）
     * @param {PageItem} item - 対象のアイテム
     * @param {*} mojikumiValue - 代入する値
     * @returns {void}
     */
    function applyMojikumiToItem(item, mojikumiValue) {
        if (item.typename === "TextFrame") {
            applyMojikumiToTextFrame(item, mojikumiValue);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) applyMojikumiToItem(item.pageItems[i], mojikumiValue);
        }
    }

    /**
     * 選択全体に適用して再描画する
     * @param {number} mojikumiIndex - mojikumiSet のインデックス（-1 は「なし」）
     * @returns {void}
     */
    function applyMojikumiToSelection(mojikumiIndex) {
        /* 値を1回だけ解決して再帰へ渡す / Resolve the value once, then pass it down */
        var mojikumiValue = (mojikumiIndex === -1) ? "なし" : doc.mojikumiSet[mojikumiIndex];
        for (var i = 0; i < doc.selection.length; i++) {
            applyMojikumiToItem(doc.selection[i], mojikumiValue);
        }
        app.redraw();
    }

    /**
     * 選択の最初のテキストフレームの、先頭段落の文字組みアキ量設定のインデックスを返す
     * 読み取りは組み込みなら内部 ID、カスタムなら名前の文字列で、mojikumiSet の要素と一致する。
     * 「行末約物半角」(2) は「約物半角」(1) と同じ ID で読み戻るので 1 になる
     * @returns {number} mojikumiSet のインデックス（「なし」・読めないときは -1）
     */
    function readSelectedMojikumiIndex() {
        for (var i = 0; i < doc.selection.length; i++) {
            if (doc.selection[i].typename !== "TextFrame") continue;
            var mojikumiText;
            try {
                mojikumiText = String(doc.selection[i].paragraphs[0].paragraphAttributes.mojikumi);
            } catch (e) {
                /* 「なし」は読むと例外 / Reading "None" throws */
                return -1;
            }
            var mojikumiSets = doc.mojikumiSet;
            for (var j = 0; j < mojikumiSets.length; j++) {
                if (String(mojikumiSets[j]) === mojikumiText) return j;
            }
            return -1;
        }
        return -1;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てて表示する
     * @returns {void}
     */
    function showMojikumiDialog() {
        var dialog = new Window("dialog", "文字組みアキ量設定 " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        /* プリセット選択パネル / Preset selection panel */
        var presetPanel = dialog.add("panel", undefined, "設定を選択");
        presetPanel.orientation = "column";
        presetPanel.alignChildren = "left";
        presetPanel.alignment = "fill";
        presetPanel.margins = [16, 20, 16, 12];
        presetPanel.spacing = 6;

        /* 選択肢ごとにラジオボタンを生成（同じ親に置いて排他にする）/ One radio button per choice (same parent, so they are exclusive) */
        var mojikumiChoices = buildMojikumiChoices();
        var selectedIndex = readSelectedMojikumiIndex();
        var mojikumiRadioButtons = [];
        for (var i = 0; i < mojikumiChoices.length; i++) {
            var radioButton = presetPanel.add("radiobutton", undefined, mojikumiChoices[i].labelText);
            radioButton.mojikumiIndex = mojikumiChoices[i].mojikumiIndex;
            radioButton.value = (radioButton.mojikumiIndex === selectedIndex);
            radioButton.onClick = function () {
                applyMojikumiToSelection(this.mojikumiIndex);
            };
            mojikumiRadioButtons.push(radioButton);
        }

        /* 閉じるボタン / Close button */
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["left", "center"];
        buttonGroup.alignment = "right";
        buttonGroup.spacing = 8;

        var closeButton = buttonGroup.add("button", undefined, "閉じる");
        closeButton.onClick = function () {
            dialog.close();
        };

        dialog.center();
        dialog.show();
    }

    showMojikumiDialog();

})();
