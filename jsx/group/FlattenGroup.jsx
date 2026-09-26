#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの入れ子グループをすべて解除してから、1つのグループにまとめ直します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenGroup.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n36fbd4162721

### Overview

Releases every nested group in the selection and then regroups everything as a single group.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenGroup.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FlattenGroup";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenGroup.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenGroup.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n36fbd4162721"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 選択 / Selection
    // =========================================

    /**
     * 処理対象の選択オブジェクトを返す
     * @returns {PageItem[]|null} 選択オブジェクト（前面→背面の順）。ドキュメントがない・未選択・文字の選択中は null
     */
    function getValidSelection() {
        if (!app.documents.length) return null;
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る
           A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || !selectedItems.length || selectedItems.typename === "TextRange") return null;
        return selectedItems;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択の入れ子グループをすべて解除してから、1つのグループにまとめ直す
     * @returns {void}
     */
    function main() {
        if (!getValidSelection()) return;

        /* 入れ子も含めてすべて解除 / Release every group, nested ones included */
        app.executeMenuCommand("ungroupAll");

        /* 選択全体を1つのグループにまとめる / Group the whole selection as one */
        app.executeMenuCommand("group");
    }

    main();

})();
