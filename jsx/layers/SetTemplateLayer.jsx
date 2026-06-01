#target illustrator

(function() {

    // =========================================================
    // 設定 / Settings
    // =========================================================

    var SCRIPT_VERSION = "v1.0";
    var specifiedLayerName = "下絵"; /* 「指定」選択時の対象レイヤー名 / Target layer name for the "Specified" option */
    var COMMENT_PREFIX = "// "; /* レイヤー名に付ける接頭辞 / Prefix added to the layer name */

    // =========================================================
    // ローカライズ / Localization
    // =========================================================

    var lang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialogTitle:   { ja: "テンプレートレイヤー設定 " + SCRIPT_VERSION, en: "Template Layer Setup " + SCRIPT_VERSION },
        panelTarget:   { ja: "対象", en: "Target" },
        radioSelected: { ja: "選択しているレイヤー", en: "Selected layer" },
        radioSpecified:{ ja: "指定", en: "Specified" },
        prefixComment: { ja: "レイヤー名に「" + COMMENT_PREFIX + "」を付ける", en: "Prefix layer name with \"" + COMMENT_PREFIX + "\"" },
        cancel:        { ja: "キャンセル", en: "Cancel" }
    };

    function L(key) {
        return LABELS[key][lang];
    }

    // 名前でレイヤーを検索（見つからなければ null） / Find a layer by name (null if missing)
    function findLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                return doc.layers[i];
            }
        }
        return null;
    }

    // レイヤーをテンプレート化（任意で接頭辞を付与） / Turn a layer into a template (optionally prefix its name)
    function applyTemplateLayer(layer, addPrefix) {
        layer.visible = true;               // 表示する
        layer.locked = true;                // ロックする
        layer.printable = false;            // 印刷しない
        layer.dimPlacedImages = true;       // 配置した画像を薄く（暗く）表示する

        // レイヤー名に接頭辞を付ける（重複付与を防ぐ） / Prefix layer name (avoid duplicating)
        if (addPrefix && layer.name.indexOf(COMMENT_PREFIX) !== 0) {
            layer.name = COMMENT_PREFIX + layer.name;
        }
    }

    // =========================================================
    // 事前チェック / Pre-checks
    // =========================================================

    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var activeDoc = app.activeDocument;

    // =========================================================
    // ダイアログ / Dialog
    // =========================================================

    var dialog = new Window("dialog", L("dialogTitle"));
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = 16;
    dialog.spacing = 12;

    // --- パネル：対象 / Panel: Target ---
    var targetLayerPanel = dialog.add("panel", undefined, L("panelTarget"));
    targetLayerPanel.orientation = "column";
    targetLayerPanel.alignChildren = "left";
    targetLayerPanel.margins = [16, 20, 16, 12];
    targetLayerPanel.spacing = 8;

    var selectedLayerRadio = targetLayerPanel.add("radiobutton", undefined, L("radioSelected"));

    // 「指定 ____」を 1 行で構成 / Build "Specified ____" on one row
    var specifiedLayerRow = targetLayerPanel.add("group");
    specifiedLayerRow.orientation = "row";
    specifiedLayerRow.spacing = 2;
    var specifiedLayerRadio = specifiedLayerRow.add("radiobutton", undefined, L("radioSpecified"));
    var specifiedLayerInput = specifiedLayerRow.add("edittext", undefined, specifiedLayerName);
    specifiedLayerInput.characters = 8;

    // 「下絵」レイヤーがあれば「指定」、無ければ現在のレイヤーを既定に / Default to "Specified" if the layer exists, otherwise the current layer
    var hasSpecifiedLayer = (findLayerByName(activeDoc, specifiedLayerName) !== null);
    specifiedLayerRadio.value = hasSpecifiedLayer;
    selectedLayerRadio.value = !hasSpecifiedLayer;

    // ラジオは別コンテナのため手動で排他制御 / Radios live in different containers, so enforce exclusivity manually
    function selectTarget(useSelected) {
        selectedLayerRadio.value = useSelected;
        specifiedLayerRadio.value = !useSelected;
    }
    selectedLayerRadio.onClick     = function() { selectTarget(true); };
    specifiedLayerRadio.onClick    = function() { selectTarget(false); };
    specifiedLayerInput.onActivate = function() { selectTarget(false); };

    // --- チェックボックス：レイヤー名に接頭辞を付ける / Checkbox: prefix layer name ---
    var prefixCommentCheckbox = dialog.add("checkbox", undefined, L("prefixComment"));
    prefixCommentCheckbox.value = true; /* 既定でON / Default: ON */

    // --- ボタン / Buttons（Mac 規約：Cancel → OK） ---
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (dialog.show() !== 1) {
        return; /* キャンセル / Cancelled */
    }

    var useSelectedLayer = selectedLayerRadio.value;
    var addCommentPrefix = prefixCommentCheckbox.value;

    // 入力が空なら既定のレイヤー名を使用 / Fall back to the default layer name when blank
    var resolvedLayerName = specifiedLayerInput.text.replace(/^\s+|\s+$/g, "");
    if (resolvedLayerName === "") {
        resolvedLayerName = specifiedLayerName;
    }

    // =========================================================
    // 対象レイヤーの決定 / Resolve target layer
    // =========================================================

    var targetLayer = useSelectedLayer
        ? activeDoc.activeLayer
        : findLayerByName(activeDoc, resolvedLayerName);

    if (!targetLayer) {
        alert(useSelectedLayer
            ? "選択しているレイヤーが取得できませんでした。"
            : "「" + resolvedLayerName + "」レイヤーが見つかりません。");
        return;
    }

    // =========================================================
    // テンプレートレイヤーの設定を適用 / Apply template layer settings
    // =========================================================

    applyTemplateLayer(targetLayer, addCommentPrefix);

    // 報告用は接頭辞を除いた名前 / Report the name without the prefix
    var displayLayerName = targetLayer.name;
    if (displayLayerName.indexOf(COMMENT_PREFIX) === 0) {
        displayLayerName = displayLayerName.substring(COMMENT_PREFIX.length);
    }

    alert("「" + displayLayerName + "」レイヤーをテンプレートに設定しました。");

})();
