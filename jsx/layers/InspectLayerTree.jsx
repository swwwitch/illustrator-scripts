#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のレイヤーとサブレイヤーを階層順にたどり、名前・ロック状態・表示状態をアラートで一覧表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InspectLayerTree.md

### Overview

Walks the document's layers and sublayers in hierarchy order and lists each name with its locked and visible state in an alert.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InspectLayerTree.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InspectLayerTree";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InspectLayerTree.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InspectLayerTree.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "alert.noDocument" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * レイヤーを階層順にたどり、名前・ロック・表示の状態を1行ずつ追加する
     * @param {Layers} layers - たどるレイヤーのコレクション
     * @param {number} depth - 階層の深さ（字下げの段数）
     * @param {string[]} reportLines - 追加先の行
     * @returns {void}
     */
    function collectLayerLines(layers, depth, reportLines) {
        var indent = "";
        for (var d = 0; d < depth; d++) indent += "  ";

        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            reportLines.push(indent + "[" + layer.name + "]  locked=" + layer.locked + " visible=" + layer.visible);
            if (layer.layers && layer.layers.length > 0) {
                collectLayerLines(layer.layers, depth + 1, reportLines);
            }
        }
    }

    /**
     * レイヤー構成をアラートで一覧表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var reportLines = [];
        collectLayerLines(app.activeDocument.layers, 0, reportLines);
        alert(reportLines.join("\n"));
    }

    main();

})();
