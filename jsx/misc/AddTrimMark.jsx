#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト（単純な長方形1点）、現在のアートボード、またはすべてのアートボードを対象に、トンボを作成します。
実行時のダイアログで、対象と「ガイドを残す」「日本式トンボ」のON/OFFを選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMark.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n40e3e39cf9f2

### Overview

Creates trim marks for a selected object (a single simple rectangle), the current artboard, or every artboard.
A dialog picks the target and toggles "keep guides" and "Japanese-style trim marks".

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMark.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddTrimMark";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMark.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMark.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n40e3e39cf9f2"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var TRIM_LAYER_NAME = "トンボ";             /* 選択オブジェクト・現在のアートボード用のレイヤー / Layer for the selection or current artboard */
    var ARTBOARD_LAYER_PREFIX = "トンボ_";      /* すべてのアートボード用のレイヤー名の接頭辞 / Layer name prefix for all artboards */
    var ARTBOARD_FALLBACK_NAME = "アートボード"; /* アートボード名が空のときの代わり / Fallback when the artboard has no name */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */
    var BUTTON_WIDTH = 80;                  /* ボタンの幅 / Button width */

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = 'column';
        targetPanel.alignChildren = 'left';
        targetPanel.alignment = 'fill';
        targetPanel.margins = PANEL_MARGINS;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "トンボ作成", en: "Create Trim Marks" }
        },
        panel: {
            target: { ja: "トンボの対象", en: "Trim Mark Target" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            selection: { ja: "選択オブジェクト", en: "Selected Object" },
            currentArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All Artboards" }
        },
        checkbox: {
            keepGuides: { ja: "ガイドを残す", en: "Keep Guides" },
            japaneseTrim: { ja: "日本式トンボ", en: "Japanese-style Trim Marks" }
        },
        tooltip: {
            selection: {
                ja: "選択した長方形のまわりにトンボを作ります。水平・垂直な長方形のときだけ選べます。",
                en: "Draws trim marks around the selected rectangle. Available only for an axis-aligned rectangle."
            },
            currentArtboard: { ja: "現在のアートボードのまわりにトンボを作ります。", en: "Draws trim marks around the current artboard." },
            allArtboards: { ja: "すべてのアートボードのまわりにトンボを作ります。", en: "Draws trim marks around every artboard." },
            keepGuides: { ja: "トンボの位置にガイドを残します。", en: "Leaves guides at the trim mark positions." },
            japaneseTrim: {
                ja: "内トンボと外トンボが対になった日本式のトンボにします。オフだと欧文式になります。",
                en: "Uses Japanese-style trim marks with paired inner and outer marks. When off, Western-style marks are drawn."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneRectangle: {
                ja: "「選択オブジェクト」を使うには、単純な長方形を1つだけ選択してください。",
                en: 'To use "Selected Object", select exactly one simple rectangle.'
            },
            simpleRectangleOnly: {
                ja: "「選択オブジェクト」で使えるのは、単純な長方形だけです。",
                en: 'Only a simple rectangle can be used for "Selected Object".'
            },
            axisAlignedOnly: {
                ja: "「選択オブジェクト」で使えるのは、各辺が水平・垂直で、4点が直交している単純な長方形だけです。",
                en: 'Only a simple rectangle with horizontal/vertical edges and right-angle corners can be used for "Selected Object".'
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) {
                return labelPath;
            }
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * レイヤー名に使えない文字を「_」に置き換える
     * @param {string} rawLayerName - 元の名前
     * @returns {string} 置き換え後の名前
     */
    function sanitizeLayerName(rawLayerName) {
        return rawLayerName.replace(new RegExp('[\\\\/:*?"<>|]', 'g'), '_');
    }

    /**
     * 名前が一致する最上位レイヤーを返す。無ければ作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 見つけた、または作ったレイヤー
     */
    function getOrCreateLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
        var newLayer = doc.layers.add();
        newLayer.name = layerName;
        return newLayer;
    }

    /**
     * アートボードごとのトンボ用レイヤー名を作る
     * @param {Artboard} artboard - 対象アートボード
     * @returns {string} レイヤー名
     */
    function buildTrimLayerNameForArtboard(artboard) {
        return ARTBOARD_LAYER_PREFIX + sanitizeLayerName(artboard.name || ARTBOARD_FALLBACK_NAME);
    }

    /**
     * トンボ用レイヤーを取得し、元のロック・表示状態を控えてから編集できる状態にする
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @param {Object[]} layerStates - 控えの記録先（{ layer, locked, visible } を追加する）
     * @returns {Layer} 編集できる状態にしたレイヤー
     */
    function prepareTrimLayer(doc, layerName, layerStates) {
        var trimLayer = getOrCreateLayerByName(doc, layerName);
        layerStates.push({ layer: trimLayer, locked: trimLayer.locked, visible: trimLayer.visible });
        if (trimLayer.locked) {
            trimLayer.locked = false;
        }
        if (!trimLayer.visible) {
            trimLayer.visible = true;
        }
        return trimLayer;
    }

    /**
     * 控えておいたロック・表示状態をレイヤーに戻す
     * @param {Object[]} layerStates - prepareTrimLayer() で控えた記録
     * @returns {void}
     */
    function restoreLayerStates(layerStates) {
        for (var i = 0; i < layerStates.length; i++) {
            layerStates[i].layer.visible = layerStates[i].visible;
            layerStates[i].layer.locked = layerStates[i].locked;
        }
    }

    // =========================================
    // 長方形の判定 / Rectangle validation
    // =========================================

    /**
     * 2つの値が許容誤差（0.01）内で等しいかを返す
     * @param {number} valueA - 値A
     * @param {number} valueB - 値B
     * @returns {boolean} ほぼ等しければ true
     */
    function isNearlyEqual(valueA, valueB) {
        return Math.abs(valueA - valueB) < 0.01;
    }

    /**
     * 指定した角の近くにアンカーポイントがあるかを返す
     * @param {number[]} anchorXs - アンカーの X 座標
     * @param {number[]} anchorYs - アンカーの Y 座標
     * @param {number} cornerX - 角の X 座標
     * @param {number} cornerY - 角の Y 座標
     * @returns {boolean} 近くにあれば true
     */
    function hasAnchorNear(anchorXs, anchorYs, cornerX, cornerY) {
        for (var i = 0; i < anchorXs.length; i++) {
            if (isNearlyEqual(anchorXs[i], cornerX) && isNearlyEqual(anchorYs[i], cornerY)) {
                return true;
            }
        }
        return false;
    }

    /**
     * パスが各辺が水平・垂直な長方形（4点・閉じた直線パス）かを返す
     * @param {PathItem} pathItem - 対象パス
     * @returns {boolean} 水平・垂直な長方形なら true
     */
    function isAxisAlignedRectangle(pathItem) {
        var pathPoints, anchorXs, anchorYs, i, left, right, top, bottom;

        if (pathItem.pathPoints.length !== 4 || !pathItem.closed) {
            return false;
        }

        pathPoints = pathItem.pathPoints;
        anchorXs = [];
        anchorYs = [];

        for (i = 0; i < 4; i++) {
            anchorXs.push(pathPoints[i].anchor[0]);
            anchorYs.push(pathPoints[i].anchor[1]);

            if (!isNearlyEqual(pathPoints[i].leftDirection[0], pathPoints[i].anchor[0]) ||
                !isNearlyEqual(pathPoints[i].leftDirection[1], pathPoints[i].anchor[1]) ||
                !isNearlyEqual(pathPoints[i].rightDirection[0], pathPoints[i].anchor[0]) ||
                !isNearlyEqual(pathPoints[i].rightDirection[1], pathPoints[i].anchor[1])) {
                return false;
            }
        }

        left = Math.min.apply(null, anchorXs);
        right = Math.max.apply(null, anchorXs);
        top = Math.max.apply(null, anchorYs);
        bottom = Math.min.apply(null, anchorYs);

        /* 幅または高さがゼロの退化パスを除外 / Reject degenerate paths with zero width or height */
        if (isNearlyEqual(left, right) || isNearlyEqual(top, bottom)) {
            return false;
        }

        for (i = 0; i < 4; i++) {
            if (!isNearlyEqual(anchorXs[i], left) && !isNearlyEqual(anchorXs[i], right)) {
                return false;
            }
            if (!isNearlyEqual(anchorYs[i], top) && !isNearlyEqual(anchorYs[i], bottom)) {
                return false;
            }
        }

        /* 4隅がすべて揃っているかを許容誤差込みで確認 / Check that all four corners are present, within tolerance */
        return hasAnchorNear(anchorXs, anchorYs, left, top) &&
            hasAnchorNear(anchorXs, anchorYs, right, top) &&
            hasAnchorNear(anchorXs, anchorYs, right, bottom) &&
            hasAnchorNear(anchorXs, anchorYs, left, bottom);
    }

    /**
     * ガイドでもクリッピングパスでもないパスかを返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {boolean} 通常のパスなら true
     */
    function isSimpleRectanglePathItem(pageItem) {
        return pageItem.typename === 'PathItem' && !pageItem.guides && !pageItem.clipping;
    }

    /**
     * 「選択オブジェクト」でトンボの基準にできるかを返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {boolean} 基準にできれば true
     */
    function isEligibleTrimSource(pageItem) {
        return isSimpleRectanglePathItem(pageItem) && isAxisAlignedRectangle(pageItem);
    }

    /**
     * 選択が「選択オブジェクト」の条件を満たしていればその長方形を返す。満たさなければ理由を知らせて null
     * @param {Document} doc - 対象ドキュメント
     * @returns {PathItem|null} 選択中の長方形
     */
    function getSelectedSimpleRectangle(doc) {
        if (doc.selection.length !== 1) {
            alert(getLabel('alert.selectOneRectangle'));
            return null;
        }

        var selectedItem = doc.selection[0];

        if (!isSimpleRectanglePathItem(selectedItem)) {
            alert(getLabel('alert.simpleRectangleOnly'));
            return null;
        }

        if (!isAxisAlignedRectangle(selectedItem)) {
            alert(getLabel('alert.axisAlignedOnly'));
            return null;
        }

        return selectedItem;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * オプションのダイアログを表示して、選ばれた設定を返す
     * @param {Document} doc - 対象ドキュメント（初期選択の判定に使う）
     * @returns {Object|null} { targetType, keepGuides, japaneseStyle }。キャンセル時は null
     */
    function showOptionsDialog(doc) {
        var trimDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);

        var targetPanel = trimDialog.add('panel', undefined, getLabel('panel.target'));
        setupPanel(targetPanel);
        var selectionRadio = targetPanel.add('radiobutton', undefined, getLabel('radio.selection'));
        selectionRadio.helpTip = getLabel('tooltip.selection');
        var currentArtboardRadio = targetPanel.add('radiobutton', undefined, getLabel('radio.currentArtboard'));
        currentArtboardRadio.helpTip = getLabel('tooltip.currentArtboard');
        var allArtboardsRadio = targetPanel.add('radiobutton', undefined, getLabel('radio.allArtboards'));
        allArtboardsRadio.helpTip = getLabel('tooltip.allArtboards');

        var optionsPanel = trimDialog.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel);
        var keepGuidesCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.keepGuides'));
        keepGuidesCheckbox.helpTip = getLabel('tooltip.keepGuides');
        var japaneseStyleCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.japaneseTrim'));
        japaneseStyleCheckbox.helpTip = getLabel('tooltip.japaneseTrim');

        var btnRowGroup = trimDialog.add('group');
        btnRowGroup.alignment = 'right';
        var btnCancel = btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;

        if (doc.selection.length === 1 && isEligibleTrimSource(doc.selection[0])) {
            selectionRadio.value = true;
        } else {
            currentArtboardRadio.value = true;
        }
        keepGuidesCheckbox.value = true;

        /* 現在の環境設定を初期値にする / Seed from the current preference */
        japaneseStyleCheckbox.value = app.preferences.getBooleanPreference('cropMarkStyle');

        if (trimDialog.show() !== 1) {
            return null;
        }

        return {
            targetType: selectionRadio.value ? 'selection' : (allArtboardsRadio.value ? 'allArtboards' : 'artboard'),
            keepGuides: keepGuidesCheckbox.value,
            japaneseStyle: japaneseStyleCheckbox.value
        };
    }

    // =========================================
    // トンボの作成 / Trim marks
    // =========================================

    /**
     * パスの塗りと線をなしにする
     * @param {PathItem} pathItem - 対象パス
     * @returns {void}
     */
    function clearFillAndStroke(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = false;
    }

    /**
     * 基準のオブジェクトを選択して［トンボを作成］を実行し、基準をガイドにするか削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} trimLayer - トンボを作るレイヤー
     * @param {PathItem} sourceItem - トンボの基準にする長方形（複製または新規）
     * @param {boolean} keepGuides - true なら基準をガイドにして残す
     * @returns {void}
     */
    function createTrimMarksFromSource(doc, trimLayer, sourceItem, keepGuides) {
        doc.activeLayer = trimLayer;
        doc.selection = null;
        doc.selection = [sourceItem];
        app.executeMenuCommand('TrimMark v25');

        if (keepGuides) {
            sourceItem.guides = true;
        } else {
            sourceItem.remove();
        }
    }

    /**
     * アートボードと同じ大きさの長方形を作り、それを基準にトンボを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} trimLayer - トンボを作るレイヤー
     * @param {Artboard} artboard - 対象アートボード
     * @param {boolean} keepGuides - true なら基準の長方形をガイドにして残す
     * @returns {void}
     */
    function createTrimMarksForArtboard(doc, trimLayer, artboard, keepGuides) {
        var artboardBounds = artboard.artboardRect;
        var artboardRectItem = trimLayer.pathItems.rectangle(artboardBounds[1], artboardBounds[0], artboardBounds[2] - artboardBounds[0], artboardBounds[1] - artboardBounds[3]);
        clearFillAndStroke(artboardRectItem);
        createTrimMarksFromSource(doc, trimLayer, artboardRectItem, keepGuides);
    }

    /**
     * 選んだ対象にトンボを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} trimOptions - showOptionsDialog() の戻り値
     * @param {PathItem|null} selectedRectangle - 「選択オブジェクト」のときの長方形
     * @param {Layer|null} trimLayer - 「トンボ」レイヤー（すべてのアートボードのときは null）
     * @param {Object[]} layerStates - レイヤー状態の控えの記録先
     * @returns {void}
     */
    function createTrimMarks(doc, trimOptions, selectedRectangle, trimLayer, layerStates) {
        if (trimOptions.targetType === 'selection') {
            /* 選択オブジェクトを複製してトンボ用に準備 / Duplicate the selected object for trim marks */
            var duplicatedItem = selectedRectangle.duplicate(trimLayer, ElementPlacement.PLACEATBEGINNING);
            clearFillAndStroke(duplicatedItem);
            createTrimMarksFromSource(doc, trimLayer, duplicatedItem, trimOptions.keepGuides);
        } else if (trimOptions.targetType === 'allArtboards') {
            for (var i = 0; i < doc.artboards.length; i++) {
                var artboardLayer = prepareTrimLayer(doc, buildTrimLayerNameForArtboard(doc.artboards[i]), layerStates);
                createTrimMarksForArtboard(doc, artboardLayer, doc.artboards[i], trimOptions.keepGuides);
            }
        } else {
            /* アクティブなアートボード / The active artboard */
            createTrimMarksForArtboard(doc, trimLayer, doc.artboards[doc.artboards.getActiveArtboardIndex()], trimOptions.keepGuides);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで設定を受け取り、環境設定とレイヤー状態を控えてトンボを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;
        var trimOptions = showOptionsDialog(doc);
        if (!trimOptions) {
            return;
        }

        var selectedRectangle = null;
        var trimLayer = null;
        var layerStates = [];
        var originalActiveLayer = doc.activeLayer;
        var originalJapaneseStyle = app.preferences.getBooleanPreference('cropMarkStyle');
        var didCreateTrimMarks = false;

        /* レイヤーや環境設定を変更する前に選択オブジェクトを検証 / Validate the selection before touching layers or preferences */
        if (trimOptions.targetType === 'selection') {
            selectedRectangle = getSelectedSimpleRectangle(doc);
            if (!selectedRectangle) {
                return;
            }
        }

        if (trimOptions.targetType !== 'allArtboards') {
            /* 「トンボ」レイヤーを取得（なければ作成） / Get the "Trim" layer, or create it if missing */
            trimLayer = prepareTrimLayer(doc, TRIM_LAYER_NAME, layerStates);
        }

        /* 失敗しても環境設定・アクティブレイヤー・レイヤー状態は戻す / Restore the preference and layers even on failure */
        try {
            /* 日本式トンボの設定を切り替える / Set Japanese-style trim marks */
            app.preferences.setBooleanPreference('cropMarkStyle', trimOptions.japaneseStyle ? 1 : 0);
            createTrimMarks(doc, trimOptions, selectedRectangle, trimLayer, layerStates);
            didCreateTrimMarks = true;
        } finally {
            /* トンボを作成できたときだけ選択を解除 / Deselect only when trim marks were actually created */
            if (didCreateTrimMarks) {
                doc.selection = null;
            }

            /* 環境設定を元に戻す / Restore the preference */
            app.preferences.setBooleanPreference('cropMarkStyle', originalJapaneseStyle ? 1 : 0);

            /* ロックや非表示を戻す前にアクティブレイヤーを復帰 / Restore the active layer before re-locking or re-hiding */
            if (originalActiveLayer.visible) {
                doc.activeLayer = originalActiveLayer;
            }

            restoreLayerStates(layerStates);
        }
    }

    main();

})();
