#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

「Guides Preview for Trim View」レイヤーを作成し、そのレイヤーをテンプレート化します。

詳細は README を参照してください。

### Overview

Creates the "Guides Preview for Trim View" layer and turns it into a template layer.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ガイドをテンプレートに";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ガイドをテンプレートに.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ガイドをテンプレートに.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================================
    // 設定 / Settings
    // =========================================================

    var NEW_LAYER_NAME = "Guides Preview for Trim View"; /* 作成するレイヤー名 / Name of the layer to create */

    // =========================================================
    // ローカライズ / Localization
    // =========================================================

    var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
    };

    function L(key) {
        if (LABELS[key] && LABELS[key][currentLanguage]) {
            return LABELS[key][currentLanguage];
        }
        return key;
    }

    // =========================================================
    // 関数 / Functions
    // =========================================================

    // 名前でレイヤーを検索（見つからなければ null） / Find a layer by name (null if missing)
    function findLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                return doc.layers[i];
            }
        }
        return null;
    }

    // ドキュメント内のガイド（パス）をすべて収集 / Collect all guide paths in the document
    function collectGuidePaths(doc) {
        var guidePaths = [];
        var documentPaths = doc.pathItems;
        for (var i = 0; i < documentPaths.length; i++) {
            if (documentPaths[i].guides === true) {
                guidePaths.push(documentPaths[i]);
            }
        }
        return guidePaths;
    }

    // ドキュメントのカラースペースに応じたマゼンタを生成 / Build a magenta color for the document's color space
    function createMagentaColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var rgb = new RGBColor();
            rgb.red = 255; rgb.green = 0; rgb.blue = 255;
            return rgb;
        }
        var cmyk = new CMYKColor();
        cmyk.cyan = 0; cmyk.magenta = 100; cmyk.yellow = 0; cmyk.black = 0;
        return cmyk;
    }

    // ガイドを解除して線をマゼンタにする / Release the guide and set its stroke to magenta
    function releaseGuideToMagenta(guidePath, magentaColor) {
        guidePath.guides = false;                 // ガイドを解除 / Release the guide
        guidePath.filled = false;                 // 塗りなし / No fill
        guidePath.stroked = true;                 // 線あり / Stroked
        guidePath.strokeColor = magentaColor;     // 線色をマゼンタに / Magenta stroke
        if (!guidePath.strokeWidth) {
            guidePath.strokeWidth = 1;            // 線幅が 0 のとき既定値を付与 / Give a default width when 0
        }
    }

    // レイヤーをテンプレート化 / Turn a layer into a template
    function applyTemplateLayer(layer) {
        layer.visible = true;               // 表示する
        layer.locked = true;                // ロックする
        layer.printable = false;            // 印刷しない
    }

    // =========================================================
    // 事前チェック / Pre-checks
    // =========================================================

    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }

    var activeDoc = app.activeDocument;

    // 前提条件：対象レイヤーが既にあれば削除して終了（トグル動作） / Precondition: if the layer already exists, remove it and stop (toggle)
    var existingLayer = findLayerByName(activeDoc, NEW_LAYER_NAME);
    if (existingLayer !== null) {
        existingLayer.locked = false; // ロック中だと削除できないため解除 / Unlock so it can be removed
        existingLayer.remove();
        app.executeMenuCommand('TrimView'); // トリミング表示を切り替え / Toggle Trim View
        return; // ガイドの収集以降は実行しない / Skip the guide-collection logic
    }

    // =========================================================
    // 全レイヤーのロック/表示状態を退避して一時解除 / Save and temporarily clear lock/visibility of all layers
    // =========================================================

    var savedLayerStates = [];
    for (var i = 0; i < activeDoc.layers.length; i++) {
        var layer = activeDoc.layers[i];
        savedLayerStates.push({ layer: layer, locked: layer.locked, visible: layer.visible });
        layer.locked = false;
        layer.visible = true;
    }

    function restoreLayerStates() {
        for (var i = 0; i < savedLayerStates.length; i++) {
            savedLayerStates[i].layer.locked = savedLayerStates[i].locked;
            savedLayerStates[i].layer.visible = savedLayerStates[i].visible;
        }
    }

    // =========================================================
    // ガイドの収集・複製（途中でエラーが出ても finally で必ず状態復帰） / Collect & duplicate guides (finally always restores state)
    // =========================================================

    var previewLayer = null;
    try {
        var guidePaths = collectGuidePaths(activeDoc);
        if (guidePaths.length === 0) {
            app.executeMenuCommand('TrimView'); // ガイドが無ければ表示トグルのみ / No guides: only toggle Trim View
            return;
        }

        // 新規レイヤーを作成し、ガイドを複製 / Create a new layer and duplicate guides into it
        previewLayer = activeDoc.layers.add();
        previewLayer.name = NEW_LAYER_NAME;

        var magentaColor = createMagentaColor(activeDoc);
        for (var i = 0; i < guidePaths.length; i++) {
            var duplicatedGuide = guidePaths[i].duplicate(previewLayer, ElementPlacement.PLACEATEND);
            releaseGuideToMagenta(duplicatedGuide, magentaColor); // ガイド解除＋マゼンタ化 / Release guide + magenta
        }
    } finally {
        // 元レイヤーの状態を復帰（新規レイヤーはこの後テンプレート化） / Restore original layers (the new layer is templated next)
        restoreLayerStates();
    }

    // =========================================================
    // テンプレートレイヤーの設定を適用 / Apply template layer settings
    // =========================================================

    applyTemplateLayer(previewLayer);

    app.executeMenuCommand('TrimView'); // トリミング表示を切り替え / Toggle Trim View

})();
