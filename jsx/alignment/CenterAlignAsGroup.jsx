#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクトを一時的にグループ化したうえで、整列パネルの「水平方向中央に整列」「垂直方向中央に整列」をダイナミックアクション経由で実行します。実行中だけ「字形の境界に整列」をONにし、終了時に元の状態へ戻します。

詳細は README を参照してください。

### Overview

Temporarily groups the selection, then centers it horizontally and vertically by running
the Align panel commands through a dynamic action. "Align to Glyph Bounds" is turned on for the run and restored afterwards.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CenterAlignAsGroup";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-21";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterAlignAsGroup.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CenterAlignAsGroup.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // アクション設定 / Action settings
    // =========================================
    var ACTION_SET_NAME = "CenterAlignAsGroup"; /* アクションセット名 / action set name */
    var ACTION_NAME     = "Center";             /* アクション名 / action name */

    /* 名前は .aia 内の /name（16進とバイト数）と一致させること / Keep the names in sync with the hex in the definition */

    /* アクション定義（.aia 形式）/ Action definition */
    var ACTION_CODE = [
        "/version 3",
        "/name [ 18",
        "\t43656e746572416c69676e417347726f7570",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        "\t/name [ 6",
        "\t\t43656e746572",
        "\t]",
        "\t/keyIndex 0",
        "\t/colorIndex 0",
        "\t/isOpen 1",
        "\t/eventCount 2",
        "\t/event-1 {",
        "\t\t/useRulersIn1stQuadrant 0",
        "\t\t/internalName (ai_plugin_alignPalette)",
        "\t\t/localizedName [ 6",
        "\t\t\te695b4e58897",
        "\t\t]",
        "\t\t/isOpen 1",
        "\t\t/isOn 1",
        "\t\t/hasDialog 0",
        "\t\t/parameterCount 1",
        "\t\t/parameter-1 {",
        "\t\t\t/key 1954115685",
        "\t\t\t/showInPalette 4294967295",
        "\t\t\t/type (enumerated)",
        "\t\t\t/name [ 27",
        "\t\t\t\te6b0b4e5b9b3e696b9e59091e4b8ade5a4aee381abe695b4e58897",
        "\t\t\t]",
        "\t\t\t/value 2",
        "\t\t}",
        "\t}",
        "\t/event-2 {",
        "\t\t/useRulersIn1stQuadrant 0",
        "\t\t/internalName (ai_plugin_alignPalette)",
        "\t\t/localizedName [ 6",
        "\t\t\te695b4e58897",
        "\t\t]",
        "\t\t/isOpen 1",
        "\t\t/isOn 1",
        "\t\t/hasDialog 0",
        "\t\t/parameterCount 1",
        "\t\t/parameter-1 {",
        "\t\t\t/key 1954115685",
        "\t\t\t/showInPalette 4294967295",
        "\t\t\t/type (enumerated)",
        "\t\t\t/name [ 27",
        "\t\t\t\te59e82e79bb4e696b9e59091e4b8ade5a4aee381abe695b4e58897",
        "\t\t\t]",
        "\t\t\t/value 5",
        "\t\t}",
        "\t}",
        "}",
        ""
    ].join("\n");

    // =========================================
    // 日英ラベル定義 / Japanese-English label definitions
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:  { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
            genericError: { ja: "エラーが発生しました：", en: "An error occurred: " },
            multipleLayers: {
                ja: "複数のレイヤーにまたがって選択されています。\nグループ化するとレイヤーが1つにまとまり、解除しても元に戻らないため中止しました。",
                en: "The selection spans multiple layers.\nGrouping would merge them into one layer and ungrouping cannot undo that, so nothing was changed."
            }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('alert','noDocument')）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    // =========================================
    // 環境設定 / Preferences
    // =========================================

    /**
     * 「字形の境界に整列」の現在の状態を取得する
     * @returns {{point: boolean, area: boolean}} ポイント文字・エリア内文字それぞれのON/OFF
     */
    function getGlyphBoundsAlign() {
        return {
            point: app.preferences.getBooleanPreference("EnableActualPointTextSpaceAlign") === true,
            area: app.preferences.getBooleanPreference("EnableActualAreaTextSpaceAlign") === true
        };
    }

    /**
     * 「字形の境界に整列」をポイント文字・エリア内文字それぞれに設定する
     * @param {{point: boolean, area: boolean}} glyphBoundsState - 設定するON/OFF
     * @returns {void}
     */
    function setGlyphBoundsAlign(glyphBoundsState) {
        app.preferences.setBooleanPreference("EnableActualPointTextSpaceAlign", glyphBoundsState.point);
        app.preferences.setBooleanPreference("EnableActualAreaTextSpaceAlign", glyphBoundsState.area);
    }

    // =========================================
    // ダイナミックアクション / Dynamic action
    // =========================================

    /**
     * アクション定義を一時ファイル経由で読み込んで実行し、読み込んだアクションを破棄する
     * @param {string} actionSetName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} actionCode - .aia 形式のアクション定義
     * @returns {void}
     */
    function runDynamicAction(actionSetName, actionName, actionCode) {
        var actionFile = new File(Folder.temp + "/CenterAlignAsGroup.aia");
        actionFile.open("w");
        actionFile.write(actionCode);
        actionFile.close();

        /* 読み込み後はファイル不要 / The file is no longer needed once loaded */
        app.loadAction(actionFile);
        actionFile.remove();

        try {
            app.doScript(actionName, actionSetName, false);
        } finally {
            app.unloadAction(actionSetName, "");
        }
    }

    // =========================================
    // 選択の検査 / Selection checks
    // =========================================

    /**
     * 文字を部分選択している場合に、その文字を含むテキストオブジェクトを選択し直す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択し直したあとの選択内容
     */
    function selectTextFrameFromTextRange(doc) {
        var storyFrames = doc.selection.story.textFrames;
        var targetFrames = [];
        for (var i = 0; i < storyFrames.length; i++) {
            targetFrames.push(storyFrames[i]);
        }
        /* 文字編集を抜けてからテキストオブジェクトを選択 / Leave text editing, then select the frames */
        app.executeMenuCommand("deselectall");
        for (var j = 0; j < targetFrames.length; j++) {
            targetFrames[j].selected = true;
        }
        return doc.selection;
    }

    /**
     * レイヤーを一意に識別するキーを作る（サブレイヤーの親子関係も含める）
     * @param {Layer} layer - 対象レイヤー
     * @returns {string} 識別用のキー
     */
    function getLayerKey(layer) {
        var keyParts = [];
        var node = layer;
        while (node && node.typename === "Layer") {
            keyParts.push(node.zOrderPosition + ":" + node.name);
            node = node.parent;
        }
        return keyParts.join("/");
    }

    /**
     * 選択が複数のレイヤーにまたがっているか判定する
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {boolean} またがっていれば true
     */
    function spansMultipleLayers(selectedItems) {
        var firstKey = getLayerKey(selectedItems[0].layer);
        for (var i = 1; i < selectedItems.length; i++) {
            if (getLayerKey(selectedItems[i].layer) !== firstKey) {
                return true;
            }
        }
        return false;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確認し、複数選択時は一時的にグループ化してアクションを実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字を部分選択しているときは selection が TextRange になるため、テキストオブジェクトに置き換える
           A partial text selection comes back as a TextRange; promote it to the text object */
        if (selectedItems && !(selectedItems instanceof Array)) {
            selectedItems = selectTextFrameFromTextRange(doc);
        }
        if (!(selectedItems instanceof Array) || selectedItems.length === 0) {
            alert(getLabel("alert", "noSelection"));
            return;
        }

        /* 複数選択時はまとめて動かすため一時的にグループ化 / Group temporarily so the selection moves as one */
        var needsGroup = selectedItems.length > 1;
        /* レイヤーをまたぐ選択はグループ化でレイヤーが移動してしまうため中止 / Grouping across layers is not reversible */
        if (needsGroup && spansMultipleLayers(selectedItems)) {
            alert(getLabel("alert", "multipleLayers"));
            return;
        }
        /* 実行中だけONにして、終了時に元の状態へ戻す / Turn on for this run only, then restore */
        var previousGlyphBounds = getGlyphBoundsAlign();
        try {
            setGlyphBoundsAlign({ point: true, area: true });
            if (needsGroup) {
                app.executeMenuCommand("group");
            }
            runDynamicAction(ACTION_SET_NAME, ACTION_NAME, ACTION_CODE);
        } catch (e) {
            alert(getLabel("alert", "genericError") + e);
        } finally {
            if (needsGroup) {
                app.executeMenuCommand("ungroup");
            }
            setGlyphBoundsAlign(previousGlyphBounds);
        }
    }

    main();

})();
