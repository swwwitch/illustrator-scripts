#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

現在の表示領域の中心に黒く塗った正方形を作成して選択し、「Convert to Shape」「Make Pixel Perfect」コマンドを適用します。

詳細は README を参照してください。

### Overview

Creates a black square at the center of the current view, selects it, and applies the
"Convert to Shape" and "Make Pixel Perfect" commands.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InsertNewRectangle";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewRectangle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n509eb6aa0a19"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var RECT_SIZE = 100; /* 作成する正方形の一辺 / side length of the square */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントがありません。", en: "No document is open." },
            noEditableLayer: {
                ja: "ロック解除かつ表示されているレイヤーがありません。",
                en: "No unlocked and visible layers found."
            }
        }
    };

    /**
     * 現在のUI言語に合わせた文言を返す
     * @param {object} labelEntry - ja / en を持つラベル定義
     * @returns {string} 表示する文言
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * ドキュメントのカラーモードに応じた黒を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラーオブジェクト
     */
    function getBlackFillColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 100;
            return cmyk;
        }
        var rgb = new RGBColor();
        rgb.red = 0;
        rgb.green = 0;
        rgb.blue = 0;
        return rgb;
    }

    /**
     * ロックされておらず表示されている最初のレイヤーを返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 編集可能なレイヤー。見つからない場合は null
     */
    function getUnlockedVisibleLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.locked && layer.visible) return layer;
        }
        return null;
    }

    /**
     * 作成先レイヤーを決め、必要ならアクティブレイヤーを切り替える
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 作成先レイヤー。見つからない場合は null
     */
    function resolveTargetLayer(doc) {
        var targetLayer = doc.activeLayer;
        if (!targetLayer.locked && targetLayer.visible) return targetLayer;

        /* ロック／非表示なら、順にロック解除かつ表示のレイヤーを探す / Fall back to the first editable layer */
        var editableLayer = getUnlockedVisibleLayer(doc);
        if (!editableLayer) return null;

        doc.activeLayer = editableLayer;
        return editableLayer;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 表示中心に正方形を作成し、選択してコマンドを適用する
     * @returns {void}
     */
    function main() {
        /* ドキュメント確認 / Ensure a document is open */
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;

        var targetLayer = resolveTargetLayer(doc);
        if (!targetLayer) {
            alert(getLabel(LABELS.alert.noEditableLayer));
            return;
        }

        /* 表示領域の中心座標を取得 / Get view center */
        var viewCenterX = doc.activeView.centerPoint[0];
        var viewCenterY = doc.activeView.centerPoint[1];

        /* 正方形を作成 / Create the square */
        var rect = targetLayer.pathItems.rectangle(
            viewCenterY + RECT_SIZE / 2,
            viewCenterX - RECT_SIZE / 2,
            RECT_SIZE,
            RECT_SIZE
        );

        /* カラーモードに応じた黒を設定し、線はなしに / Fill with black, no stroke */
        rect.fillColor = getBlackFillColor(doc);
        rect.stroked = false;

        /* 作成した正方形だけを選択 / Select the created square only */
        doc.selection = null;
        rect.selected = true;

        /* 選択オブジェクトにコマンドを適用 / Apply commands to the selection */
        app.executeMenuCommand("Convert to Shape");
        app.executeMenuCommand("Make Pixel Perfect");
    }

    main();

})();
