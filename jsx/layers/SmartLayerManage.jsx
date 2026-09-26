#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクトを指定したレイヤーへ一括で移動します。
対象は選択オブジェクト、全テキスト、すべて、すべて（強制）から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartLayerManage.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n95ec4929ae9d

### Overview

Moves objects to a chosen layer in bulk.
The scope can be the selection, all text, everything, or everything (forced).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartLayerManage.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartLayerManage";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartLayerManage.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartLayerManage.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n95ec4929ae9d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_NEW_LAYER_NAME = "New Layer";    /* 新規レイヤー名の初期値 / default name for a new layer */
    var DEFAULT_DELETE_EMPTY = true;             /* 空レイヤーを削除するかの初期値 / delete empty layers by default */
    var PROTECTED_LAYER_NAME = "bg";             /* 空でも削除しないレイヤー名 / layer name kept even when empty */
    var PROTECTED_LAYER_PREFIX = "//";           /* この接頭辞のレイヤーは空でも削除しない / prefix of layers kept even when empty */
    var TARGET_LAYER_COLOR = [79, 128, 255];     /* 移動先レイヤーのカラー（RGB）/ layer color given to the destination */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];        /* パネル余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 20;                     /* 列の間隔 / gap between columns */
    var NEW_LAYER_FIELD_CHARS = 12;              /* 新規レイヤー名欄の文字数 / width of the new layer name field */

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
        dialog: {
            title: { ja: "オブジェクトをレイヤーへ移動", en: "Move Objects to Layer" }
        },
        panel: {
            scope: { ja: "対象オブジェクト", en: "Target Objects" },
            targetLayer: { ja: "移動先レイヤー", en: "Target Layer" }
        },
        radio: {
            selected: { ja: "選択中", en: "Selected Objects" },
            allText: { ja: "全テキスト", en: "All Text" },
            allObjects: { ja: "すべて", en: "All Objects" },
            allForce: { ja: "すべて（強制）", en: "All (Force)" },
            newLayer: { ja: "新規", en: "New" }
        },
        checkbox: {
            deleteEmpty: { ja: "空レイヤーを削除", en: "Delete Empty Layers" }
        },
        tooltip: {
            selected: { ja: "選択しているオブジェクトだけを移動します。", en: "Moves only the selected objects." },
            allText: { ja: "ドキュメント内のテキストをすべて移動します。", en: "Moves every text object in the document." },
            allObjects: {
                ja: "ロックと非表示をすべて解除してから、すべてのオブジェクトを移動します。",
                en: "Unlocks and shows everything, then moves every object."
            },
            allForce: {
                ja: "［すべてのレイヤーを結合］を実行し、すべてを一番上のレイヤーにまとめます。",
                en: "Runs Flatten Artwork, merging everything into the top layer."
            },
            deleteEmpty: {
                ja: "移動したあと、中身が無くなったレイヤーを削除します（「bg」と「//」で始まるレイヤーは残します）。",
                en: "Deletes the layers left empty after the move (layers named \"bg\" or starting with \"//\" are kept)."
            },
            newLayer: {
                ja: "新しいレイヤーを作って、そこへ移動します。名前は右の欄で決めます。",
                en: "Creates a new layer and moves the objects there. The field on the right names it."
            },
            layerRadio: { ja: "このレイヤーへ移動します。", en: "Moves the objects to this layer." }
        },
        button: {
            move: { ja: "移動", en: "Move" },
            close: { ja: "閉じる", en: "Close" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noLayerName: { ja: "新しいレイヤー名を入力してください。", en: "Enter a name for the new layer." },
            noLayerSelected: { ja: "移動先レイヤーを選択してください。", en: "Please select a target layer." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
            error: { ja: "エラーが発生しました: ", en: "An error occurred: " }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.scope" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

    /**
     * 一時アクションで［すべてのレイヤーを結合］を実行する
     * @returns {void}
     */
    function flattenArtwork() {
        var actionSetName = "layer";
        var actionName = "flattenLayers";

        /* アクション定義 / Action definition */
        var actionCode = [
            "/version 3 /name [ 5 6c61796572 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 13 666c617474656e4c6179657273 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_Layer) /localizedName [ 9 e8a1a8e7a4ba203a20 ] /isOpen 0 /isOn 1 /hasDialog 0 /parameterCount 2 /parameter-1 { /key 1836411236 /showInPalette 4294967295 /type (integer) /value 14 } /parameter-2 { /key 1851878757 /showInPalette 4294967295 /type (ustring) /value [ 33 e38199e381b9e381a6e381aee383ace382a4e383a4e383bce38292e7b590e590 88 ] } } }"
        ].join("\n");

        var tempFile = new File(Folder.temp + "/temp_action.aia");
        tempFile.open("w");
        tempFile.write(actionCode);
        tempFile.close();

        /* 読み込み時点でパース済みなので、ここで一時ファイルを消しておく / The action is parsed on load, so remove the temp file now */
        app.loadAction(tempFile);
        tempFile.remove();

        try {
            app.doScript(actionName, actionSetName);
        } catch (e) {
            alert(getLabel("alert.error") + e);
        } finally {
            app.unloadAction(actionSetName, "");
        }
    }

    // =========================================
    // 収集・移動 / Collect and move
    // =========================================

    /**
     * コンテナ（レイヤー／グループ）の pageItems を再帰的に集める（グループの中身、レイヤーならサブレイヤーも）
     * @param {Layer|GroupItem} container - 対象のレイヤーまたはグループ
     * @param {PageItem[]} collectedItems - 追加先の配列
     * @returns {void}
     */
    function collectItemsRecursive(container, collectedItems) {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            collectedItems.push(item);
            if (item.typename === "GroupItem") {
                collectItemsRecursive(item, collectedItems);
            }
        }
        if (container.typename === "Layer") {
            for (var j = 0; j < container.layers.length; j++) {
                collectItemsRecursive(container.layers[j], collectedItems);
            }
        }
    }

    /**
     * レイヤーとサブレイヤーからテキストフレームを再帰的に集める
     * @param {Layer} layer - 対象レイヤー
     * @param {TextFrame[]} collectedItems - 追加先の配列
     * @returns {void}
     */
    function collectTextFramesRecursive(layer, collectedItems) {
        for (var i = 0; i < layer.textFrames.length; i++) {
            collectedItems.push(layer.textFrames[i]);
        }
        for (var j = 0; j < layer.layers.length; j++) {
            collectTextFramesRecursive(layer.layers[j], collectedItems);
        }
    }

    /**
     * アイテムのロック・非表示を解除して移動先レイヤーへ移す
     * @param {PageItem} item - 移動するアイテム
     * @param {Layer} targetLayer - 移動先レイヤー
     * @returns {void}
     */
    function prepareAndMoveItem(item, targetLayer) {
        try {
            item.locked = false;
            item.hidden = false;
            item.move(targetLayer, ElementPlacement.PLACEATEND);
        } catch (e) {
            /* 移動済みの親と一緒に動いた子や、ロックされたレイヤー内のものは飛ばす / Skip items that cannot be moved */
        }
    }

    /**
     * 空のレイヤーを再帰的に削除する（保護するレイヤー名・接頭辞は残す）
     * @param {Layer} layer - 対象レイヤー
     * @returns {void}
     */
    function deleteEmptyLayersRecursive(layer) {
        for (var i = layer.layers.length - 1; i >= 0; i--) {
            deleteEmptyLayersRecursive(layer.layers[i]);
        }
        if (layer.pageItems.length === 0 && layer.layers.length === 0 &&
            layer.name !== PROTECTED_LAYER_NAME && layer.name.indexOf(PROTECTED_LAYER_PREFIX) !== 0) {
            layer.remove();
        }
    }

    /**
     * レイヤーカラーを RGB で変更する（表示中でロックされていないときだけ）
     * @param {Layer} layer - 対象レイヤー
     * @param {number[]} rgbValues - [r, g, b]
     * @returns {void}
     */
    function setLayerColor(layer, rgbValues) {
        if (!layer.visible || layer.locked) return;
        var newColor = new RGBColor();
        newColor.red = rgbValues[0];
        newColor.green = rgbValues[1];
        newColor.blue = rgbValues[2];
        layer.color = newColor;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを追加して helpTip を付ける
     * @param {Panel|Group} parent - 追加先
     * @param {string} labelKey - radio / tooltip 共通のキー
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addScopeRadio(parent, labelKey) {
        var radio = parent.add("radiobutton", undefined, getLabel("radio." + labelKey));
        radio.helpTip = getLabel("tooltip." + labelKey);
        return radio;
    }

    /**
     * ダイアログを組み立てる
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} ダイアログと主要なコントロール
     */
    function buildDialog(doc) {
        var layers = doc.layers;
        var hasSelection = doc.selection && doc.selection.length > 0;

        var dialog = new Window("dialog", getLabel("dialog.title"));
        dialog.orientation = "row";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = COLUMN_SPACING;

        /* 左列：対象オブジェクト / Left column: scope */
        var leftColumn = dialog.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["left", "top"];

        var scopePanel = leftColumn.add("panel", undefined, getLabel("panel.scope"));
        scopePanel.orientation = "column";
        scopePanel.alignChildren = ["left", "top"];
        scopePanel.margins = PANEL_MARGINS;

        var controls = { dialog: dialog };
        controls.radioSelected = addScopeRadio(scopePanel, "selected");
        controls.radioAllText = addScopeRadio(scopePanel, "allText");
        controls.radioAllObjects = addScopeRadio(scopePanel, "allObjects");
        controls.radioAllForce = addScopeRadio(scopePanel, "allForce");

        /* 初期値は選択の有無で決める / Default depends on whether anything is selected */
        controls.radioSelected.value = hasSelection;
        controls.radioAllObjects.value = !hasSelection;

        var deleteEmptyRow = leftColumn.add("group");
        deleteEmptyRow.orientation = "row";
        deleteEmptyRow.alignChildren = ["center", "center"];
        deleteEmptyRow.margins = [5, 5, 0, 0];

        controls.deleteEmptyCheckbox = deleteEmptyRow.add("checkbox", undefined, getLabel("checkbox.deleteEmpty"));
        controls.deleteEmptyCheckbox.helpTip = getLabel("tooltip.deleteEmpty");
        controls.deleteEmptyCheckbox.value = DEFAULT_DELETE_EMPTY;

        /* 中央列：移動先レイヤー / Middle column: destination layer */
        var rightColumn = dialog.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];

        var layerPanel = rightColumn.add("panel", undefined, getLabel("panel.targetLayer"));
        layerPanel.orientation = "column";
        layerPanel.alignChildren = ["left", "top"];
        layerPanel.margins = PANEL_MARGINS;

        /* 先頭に「新規」ラジオと名前欄 / "New" radio and name field first */
        var newLayerRow = layerPanel.add("group");
        newLayerRow.orientation = "row";
        controls.radioNewLayer = addScopeRadio(newLayerRow, "newLayer");
        controls.newLayerNameInput = newLayerRow.add("edittext", undefined, DEFAULT_NEW_LAYER_NAME);
        controls.newLayerNameInput.helpTip = getLabel("tooltip.newLayer");
        controls.newLayerNameInput.characters = NEW_LAYER_FIELD_CHARS;

        /* 既存レイヤー（ロック中は選べない）/ Existing layers; locked ones are disabled */
        controls.layerRadios = [controls.radioNewLayer];
        var firstAvailableIndex = -1;
        for (var i = 0; i < layers.length; i++) {
            var layerRadio = layerPanel.add("radiobutton", undefined, layers[i].name);
            layerRadio.helpTip = getLabel("tooltip.layerRadio");
            controls.layerRadios.push(layerRadio);
            if (layers[i].locked) {
                layerRadio.enabled = false;
            } else if (firstAvailableIndex === -1) {
                firstAvailableIndex = i + 1; /* 「新規」の分だけずらす / Offset by the "New" radio */
            }
        }
        if (firstAvailableIndex !== -1) {
            controls.layerRadios[firstAvailableIndex].value = true;
        }

        bindLayerRadios(controls);

        /* 右列：ボタン / Right column: buttons */
        var btnColumnGroup = dialog.add("group");
        btnColumnGroup.orientation = "column";
        btnColumnGroup.alignChildren = ["fill", "top"];
        controls.btnMove = btnColumnGroup.add("button", undefined, getLabel("button.move"), { name: "ok" });
        controls.btnClose = btnColumnGroup.add("button", undefined, getLabel("button.close"));

        return controls;
    }

    /**
     * 移動先ラジオの排他と名前欄の有効状態をつなぐ（「新規」は別グループにあるため手動で排他する）
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {void}
     */
    function bindLayerRadios(controls) {
        var layerRadios = controls.layerRadios;

        function selectLayerRadio(selectedIndex) {
            for (var i = 0; i < layerRadios.length; i++) {
                layerRadios[i].value = (i === selectedIndex);
            }
            controls.newLayerNameInput.enabled = controls.radioNewLayer.value;
        }

        for (var i = 0; i < layerRadios.length; i++) {
            (function (radioIndex) {
                layerRadios[radioIndex].onClick = function () {
                    selectLayerRadio(radioIndex);
                };
            })(i);
        }
        controls.newLayerNameInput.enabled = controls.radioNewLayer.value;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログの指定から移動先レイヤーを決める（「新規」なら一番上に作る）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {Layer|null} 移動先レイヤー。指定が足りなければ null
     */
    function resolveTargetLayer(doc, controls) {
        if (controls.radioNewLayer.value) {
            var newLayerName = controls.newLayerNameInput.text;
            if (!newLayerName) {
                alert(getLabel("alert.noLayerName"));
                return null;
            }
            var newLayer = doc.layers.add();
            newLayer.name = newLayerName;
            newLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
            return newLayer;
        }
        for (var i = 1; i < controls.layerRadios.length; i++) {
            if (controls.layerRadios[i].value) {
                return doc.layers.getByName(controls.layerRadios[i].text);
            }
        }
        alert(getLabel("alert.noLayerSelected"));
        return null;
    }

    /**
     * 対象の指定に応じて移動するアイテムを集める（「すべて（強制）」はレイヤーを結合して空を返す）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {PageItem[]|null} 移動するアイテム。選択が無ければ null
     */
    function collectItemsToMove(doc, controls) {
        var itemsToMove = [];
        if (controls.radioSelected.value) {
            var currentSelection = doc.selection;
            if (!currentSelection || currentSelection.length === 0) {
                alert(getLabel("alert.noSelection"));
                return null;
            }
            return currentSelection;
        }
        if (controls.radioAllText.value) {
            for (var i = 0; i < doc.layers.length; i++) {
                collectTextFramesRecursive(doc.layers[i], itemsToMove);
            }
        } else if (controls.radioAllForce.value) {
            flattenArtwork();
        } else {
            app.executeMenuCommand("unlockAll");
            app.executeMenuCommand("showAll");
            for (var j = 0; j < doc.layers.length; j++) {
                collectItemsRecursive(doc.layers[j], itemsToMove);
            }
        }
        return itemsToMove;
    }

    /**
     * ［移動］の処理：移動先を決めて移し、空レイヤーを削除して移動先のカラーを変える
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} controls - buildDialog() のコントロール
     * @returns {void}
     */
    function runMove(doc, controls) {
        var targetLayer = resolveTargetLayer(doc, controls);
        if (!targetLayer) return;

        var itemsToMove = collectItemsToMove(doc, controls);
        if (!itemsToMove) return;

        for (var i = 0; i < itemsToMove.length; i++) {
            prepareAndMoveItem(itemsToMove[i], targetLayer);
        }

        if (controls.deleteEmptyCheckbox.value) {
            for (var j = doc.layers.length - 1; j >= 0; j--) {
                deleteEmptyLayersRecursive(doc.layers[j]);
            }
        }

        setLayerColor(targetLayer, TARGET_LAYER_COLOR);
        controls.dialog.close();
    }

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var controls = buildDialog(doc);

        controls.btnMove.onClick = function () {
            runMove(doc, controls);
        };
        controls.btnClose.onClick = function () {
            controls.dialog.close();
        };

        controls.dialog.show();
    }

    main();

})();
