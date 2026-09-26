#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

「Guides Preview for Trim View」レイヤーを探し、ロックと非表示を解除してから削除し、結果をアラートで報告します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RemoveTrimViewGuideLayer.md

### Overview

Finds the "Guides Preview for Trim View" layer, unlocks and unhides it, removes it, and reports the result in an alert.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RemoveTrimViewGuideLayer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RemoveTrimViewGuideLayer";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RemoveTrimViewGuideLayer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RemoveTrimViewGuideLayer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var TARGET_LAYER_NAME = "Guides Preview for Trim View";  /* 削除するレイヤー名 / name of the layer to remove */

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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            found: { ja: "対象発見: ", en: "Found: " },
            removed: { ja: "  → remove() 成功", en: "  → remove() succeeded" },
            removeFailed: { ja: "  → remove() 失敗: ", en: "  → remove() failed: " },
            notFound: { ja: "対象が見つかりませんでした", en: "The layer was not found." }
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
     * レイヤー名から先頭の「*」と前後の空白を除く（テンプレートレイヤーなどの表記ゆれ対策）
     * @param {string} layerName - レイヤー名
     * @returns {string} 正規化した名前
     */
    function normalizeLayerName(layerName) {
        return layerName.replace(/^\s*\*?\s*/, "").replace(/\s+$/, "");
    }

    /**
     * 対象のトップレベルレイヤーを削除し、結果を報告する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        var reportLines = [];
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var layer = doc.layers[i];
            if (normalizeLayerName(layer.name) !== TARGET_LAYER_NAME) continue;

            reportLines.push(getLabel("alert.found") + "[" + layer.name + "] locked=" + layer.locked + " visible=" + layer.visible);
            try {
                /* ロック・非表示のままでは削除できない / A locked or hidden layer cannot be removed */
                layer.locked = false;
                layer.visible = true;
                layer.remove();
                reportLines.push(getLabel("alert.removed"));
            } catch (e) {
                reportLines.push(getLabel("alert.removeFailed") + e.message + " (line " + e.line + ")");
            }
        }
        if (reportLines.length === 0) reportLines.push(getLabel("alert.notFound"));
        alert(reportLines.join("\n"));
    }

    main();

})();
