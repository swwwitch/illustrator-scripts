#target illustrator

/*
 * TrimViewWithGuidesToggle.jsx
 *
 * 概要 / Overview:
 *   トリミング表示（Trim View）と、ガイドのプレビューレイヤーをワンクリックでトグルする。
 *   Toggles Trim View together with a preview layer that mirrors the document's guides.
 *
 * 動作 / Behavior:
 *   - 「Guides Preview for Trim View」レイヤーが無い場合：
 *       ドキュメント内のガイドをすべて複製した新規レイヤーを作成し、
 *       複製したガイドはガイド解除＋マゼンタの線にして、レイヤーをテンプレート化する。
 *       最後にトリミング表示をオンにする。
 *   - 同レイヤーが既にある場合：そのレイヤーを削除し、トリミング表示をオフにする。
 *   - ガイドが1本も無い場合：トリミング表示の切り替えのみ行う。
 *
 *   - If the "Guides Preview for Trim View" layer is absent: create it, duplicate every
 *     guide into it as released magenta strokes, make it a template layer, then enable Trim View.
 *   - If the layer already exists: remove it and disable Trim View.
 *   - If there are no guides at all: just toggle Trim View.
 */

(function() {

    // =========================================================
    // 設定 / Settings
    // =========================================================

    var SCRIPT_VERSION = "v1.5.0";
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

    // ガイドを複製したプレビュー用テンプレートレイヤーを作成 / Build the preview template layer from duplicated guides
    function buildPreviewLayer(doc, guidePaths) {
        // 全レイヤーのロック/表示を一時解除（ロック・非表示レイヤー上のガイドも複製するため）
        // Temporarily unlock & show every layer so guides on locked/hidden layers can be duplicated
        var savedLayerStates = [];
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            savedLayerStates.push({ layer: layer, locked: layer.locked, visible: layer.visible });
            layer.locked = false;
            layer.visible = true;
        }

        var previewLayer = doc.layers.add();
        previewLayer.name = NEW_LAYER_NAME;

        try {
            var magentaColor = createMagentaColor(doc);
            for (var j = 0; j < guidePaths.length; j++) {
                var duplicatedGuide = guidePaths[j].duplicate(previewLayer, ElementPlacement.PLACEATEND);
                releaseGuideToMagenta(duplicatedGuide, magentaColor); // ガイド解除＋マゼンタ化 / Release guide + magenta
            }
        } finally {
            // 途中でエラーが出ても元レイヤーの状態を必ず復帰 / Always restore original layers, even on error
            for (var k = 0; k < savedLayerStates.length; k++) {
                savedLayerStates[k].layer.locked = savedLayerStates[k].locked;
                savedLayerStates[k].layer.visible = savedLayerStates[k].visible;
            }
        }

        applyTemplateLayer(previewLayer);
    }

    // =========================================================
    // 事前チェック / Pre-checks
    // =========================================================

    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }

    var activeDoc = app.activeDocument;

    // =========================================================
    // メイン処理（トグル動作） / Main (toggle behavior)
    // =========================================================

    var existingLayer = findLayerByName(activeDoc, NEW_LAYER_NAME);
    if (existingLayer) {
        // 既にあれば削除（プレビューを消してトリミング表示をオフへ） / Already there: remove it (turn the preview off)
        existingLayer.locked = false; // ロック中だと削除できないため解除 / Unlock so it can be removed
        existingLayer.remove();
    } else {
        // 無ければガイドを集めてプレビューレイヤーを作成 / Otherwise build the preview layer from the guides
        var guidePaths = collectGuidePaths(activeDoc);
        if (guidePaths.length > 0) {
            buildPreviewLayer(activeDoc, guidePaths);
        }
    }

    app.executeMenuCommand('TrimView'); // いずれの場合も最後にトリミング表示をトグル / Always toggle Trim View last

})();