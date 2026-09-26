#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクトとテキストを同時に選択して実行すると、オブジェクトの見た目をグラフィックスタイルとして登録し、テキストの文字列をそのスタイル名にします。
テキスト＋オブジェクトのグループ選択や、複数グループの一括処理にも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterGraphicStyleWithText.md

### Overview

Registers the appearance of the selected object as a graphic style, using the selected text's content as the style name.
A group containing one text frame plus one object, and several such groups at once, are handled as well.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegisterGraphicStyleWithText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RegisterGraphicStyleWithText"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterGraphicStyleWithText.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegisterGraphicStyleWithText.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 選択の仕分け / Selection parsing
    // =========================================

    /**
     * 前後の空白を除去する（ES3 には String.trim が無い）
     * @param {string} value - 対象の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimText(value) {
        return String(value).replace(/^\s+/, '').replace(/\s+$/, '');
    }

    /**
     * 選択からテキストとスタイル対象オブジェクトの組を取り出す
     * 対応: (a) テキスト1点＋オブジェクト1点 / (b) グループ1点（中身がテキスト＋オブジェクト）
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {{text: TextFrame, target: PageItem}|null} 見つかった組（なければ null）
     */
    function resolveTextAndStyleTarget(selectedItems) {
        /* グループ1点なら、その中身を仕分け対象にする / If a single group is selected, look inside it */
        var candidates = selectedItems;
        if (selectedItems.length === 1 && selectedItems[0].typename === 'GroupItem') {
            candidates = selectedItems[0].pageItems;
        }

        if (candidates.length !== 2) {
            return null;
        }

        /* テキストとスタイル対象オブジェクトに仕分け / Sort into text and style-target */
        var textItem = null;
        var styleTarget = null;
        for (var i = 0; i < candidates.length; i++) {
            if (candidates[i].typename === 'TextFrame') {
                textItem = candidates[i];
            } else {
                styleTarget = candidates[i];
            }
        }
        if (!textItem || !styleTarget) {
            return null;
        }
        return { text: textItem, target: styleTarget };
    }

    /**
     * すべてがグループかどうかを判定する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 1点以上あり、すべて GroupItem なら true
     */
    function isAllGroups(selectedItems) {
        if (selectedItems.length === 0) {
            return false;
        }
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename !== 'GroupItem') {
                return false;
            }
        }
        return true;
    }

    /**
     * 選択から登録ジョブ（テキスト＋スタイル対象オブジェクトの組）を集める
     * 対応: (a) テキスト1点＋オブジェクト1点 / (b) グループ1点 / (c) 複数グループ（グループごと）
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {Array<{text: TextFrame, target: PageItem}>} 登録ジョブの配列
     */
    function collectStyleJobs(selectedItems) {
        var jobs = [];
        if (selectedItems.length >= 2 && isAllGroups(selectedItems)) {
            /* 複数グループはグループごとに仕分け / Process each group separately */
            for (var i = 0; i < selectedItems.length; i++) {
                var groupJob = resolveTextAndStyleTarget([selectedItems[i]]);
                if (groupJob) {
                    jobs.push(groupJob);
                }
            }
        } else {
            var selectionJob = resolveTextAndStyleTarget(selectedItems);
            if (selectionJob) {
                jobs.push(selectionJob);
            }
        }
        return jobs;
    }

    // =========================================
    // グラフィックスタイル関連 / Graphic Style Helpers
    // =========================================

    /**
     * 「新規グラフィックスタイル」を名前なしで実行するアクションを一時ファイルに書き出してロードする
     * @returns {void}
     */
    function loadForceNewGraphicStyleAction() {
        var actionData = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382 a4e383ab ] /value 1 } } }';

        var actionFile = new File(Folder.temp.fsName + '/__tmp_register_style.aia');
        actionFile.open('w');
        actionFile.write(actionData);
        actionFile.close();
        app.loadAction(actionFile);
        actionFile.remove();
    }

    /**
     * ロード済みのアクションを実行し、現在の選択をグラフィックスタイルとして登録する
     * @returns {void}
     */
    function runForceNewGraphicStyleAction() {
        app.doScript('AddNewWithoutName', 'GraphicStyle', false);
    }

    /**
     * 一時的にロードしたアクションセットをアンロードする
     * @returns {void}
     */
    function unloadForceNewGraphicStyleAction() {
        app.unloadAction('GraphicStyle', '');
    }

    /**
     * 1ジョブ分のスタイルを登録する（アクションはあらかじめロード済みであること）
     * @param {Document} activeDoc - 対象ドキュメント
     * @param {GraphicStyles} graphicStyles - ドキュメントのグラフィックスタイル
     * @param {{text: TextFrame, target: PageItem}} job - 登録ジョブ
     * @returns {void}
     */
    function registerGraphicStyleFromJob(activeDoc, graphicStyles, job) {
        /* テキストからスタイル名を取得（空なら何もしない） / Get the style name from the text; skip when empty */
        var styleName = trimText(job.text.contents);
        if (styleName === '') {
            return;
        }

        /* 既存の同名スタイルを削除（getByName は見つからないと例外） / Remove the existing style; getByName throws when missing */
        try {
            graphicStyles.getByName(styleName).remove();
        } catch (e) { }

        /* スタイル対象オブジェクトだけを選択（テキストの見た目を含めない） / Select only the target so the text is not included */
        activeDoc.selection = [job.target];

        /* スタイルが追加された場合のみ、末尾を改名 / Rename the last style only when one was added */
        var beforeCount = graphicStyles.length;
        runForceNewGraphicStyleAction();
        if (graphicStyles.length > beforeCount) {
            graphicStyles[graphicStyles.length - 1].name = styleName;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から登録ジョブを集め、テキストの文字列を名前にしてグラフィックスタイルを登録する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            return;
        }

        var activeDoc = app.activeDocument;
        var graphicStyles = activeDoc.graphicStyles;
        var selectedItems = activeDoc.selection;

        var jobs = collectStyleJobs(selectedItems);
        if (jobs.length === 0) {
            return;
        }

        /* アクションをロード → ジョブごとに実行 → アンロード / Load, run per job, then unload */
        loadForceNewGraphicStyleAction();
        for (var i = 0; i < jobs.length; i++) {
            registerGraphicStyleFromJob(activeDoc, graphicStyles, jobs[i]);
        }
        unloadForceNewGraphicStyleAction();

        /* 元の選択に戻す / Restore the original selection */
        activeDoc.selection = selectedItems;
    }

    main();

})();
