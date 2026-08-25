#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトがグループ内にある場合、親グループを辿って所属レイヤーの直下へ移動します。
重ね順が反転しないよう、選択オブジェクトは逆順に処理します。

詳細は README を参照してください。

### Overview

Moves the selected objects out of their groups and directly onto the layer that owns them.
They are processed in reverse order so that the stacking order is preserved.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReleaseFromGroup";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-03-06";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseFromGroup.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseFromGroup.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

/**
 * 現在のUIロケールが日本語かどうかを判定する
 * @returns {string} 日本語環境なら "ja"、それ以外は "en"
 */
function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "オブジェクトを選択して実行してください。", en: "Please select objects and run the script." }
    }
};

/**
 * LABELS から現在のロケールに対応する文言を取得する
 * @param {string} key - LABELS のキー
 * @returns {string} ロケールに対応する文言。見つからない場合は英語、それも無ければキーをそのまま返す
 */
function L(path) {
    var parts = String(path).split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return path;
        node = node[parts[i]];
    }
    if (node == null) return path;
    if (node[lang] != null) return node[lang];
    return (node.en != null) ? node.en : path;
}

// main
var doc = app.documents.length && app.activeDocument;
if (!doc) return;

var sel = doc.selection;
if (!sel.length) return;

// 2. 選択されたすべてのオブジェクトを退避
var targets = [].slice.call(sel);

// 3. 重ね順が逆転しないように「逆順（後ろから）」で処理を行う
for (var j = targets.length - 1; j >= 0; j--) {
    var obj = targets[j];

    // 親を辿って最上位のレイヤーを探す（GroupItem 内にあるものだけ対象）
    var targetLayer = obj.parent;
    if (targetLayer.typename !== "GroupItem") continue;

    while (targetLayer && targetLayer.typename === "GroupItem") {
        targetLayer = targetLayer.parent;
    }

    // 見つかったレイヤーにオブジェクトを移動
    if (targetLayer && targetLayer.typename === "Layer") {
        // レイヤーの最前面に移動
        // 逆順で処理しているため、結果的に元の上下関係が維持される
        obj.move(targetLayer, ElementPlacement.PLACEATBEGINNING);

        // 移動後も選択状態を維持
        obj.selected = true;
    }
}

/* 選択ツールに戻す / Return to Selection Tool */
app.selectTool('Adobe Select Tool');

})();
