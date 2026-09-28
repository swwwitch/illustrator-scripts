#target illustrator
#targetengine "SetTemplateLayerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定したレイヤーをテンプレート化し、レイヤー名に接頭辞を付けます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SetTemplateLayer.md

### Overview

Turns a chosen layer into a template layer and prefixes its name.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SetTemplateLayer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SetTemplateLayer";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SetTemplateLayer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SetTemplateLayer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================================
    // 設定 / Settings
    // =========================================================

    var specifiedLayerName = "下絵"; /* 「指定」選択時の対象レイヤー名 / Target layer name for the "Specified" option */
    var COMMENT_PREFIX = "// "; /* レイヤー名に付ける接頭辞 / Prefix added to the layer name */

    // =========================================================
    // ローカライズ / Localization
    // =========================================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialogTitle:   { ja: "テンプレートレイヤー設定 " + SCRIPT_VERSION, en: "Template Layer Setup " + SCRIPT_VERSION },
        panelTarget:   { ja: "対象", en: "Target" },
        radioSelected: { ja: "選択しているレイヤー", en: "Selected layer" },
        radioSpecified:{ ja: "指定", en: "Specified" },
        prefixComment: { ja: "レイヤー名に「" + COMMENT_PREFIX + "」を付ける", en: "Prefix layer name with \"" + COMMENT_PREFIX + "\"" },
        cancel:        { ja: "キャンセル", en: "Cancel" },
        tipSelected:   { ja: "レイヤーパネルで選択中のレイヤーをテンプレートにします。", en: "Turns the layer selected in the Layers panel into a template." },
        tipSpecified:  { ja: "名前で指定したレイヤーをテンプレートにします。無ければ作成します。", en: "Turns the layer with the given name into a template, creating it if needed." },
        tipPrefix:     { ja: "テンプレート化したレイヤーの名前に接頭辞を付けて、見分けやすくします。", en: "Prefixes the name of the template layer so it stands out." }
    };

    function getLabel(key) {
        return LABELS[key][uiLang];
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================================
    // ダイアログ / Dialog
    // =========================================================

    var dialog = new Window("dialog", getLabel("dialogTitle"));
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = 16;
    dialog.spacing = 12;

    // --- パネル：対象 / Panel: Target ---
    var targetLayerPanel = dialog.add("panel", undefined, getLabel("panelTarget"));
    targetLayerPanel.orientation = "column";
    targetLayerPanel.alignChildren = "left";
    targetLayerPanel.margins = [16, 20, 16, 12];
    targetLayerPanel.spacing = 8;

    var selectedLayerRadio = targetLayerPanel.add("radiobutton", undefined, getLabel("radioSelected"));
    selectedLayerRadio.helpTip = getLabel("tipSelected");

    // 「指定 ____」を 1 行で構成 / Build "Specified ____" on one row
    var specifiedLayerRow = targetLayerPanel.add("group");
    specifiedLayerRow.orientation = "row";
    specifiedLayerRow.spacing = 2;
    var specifiedLayerRadio = specifiedLayerRow.add("radiobutton", undefined, getLabel("radioSpecified"));
    specifiedLayerRadio.helpTip = getLabel("tipSpecified");
    var specifiedLayerInput = specifiedLayerRow.add("edittext", undefined, specifiedLayerName);
    specifiedLayerInput.helpTip = getLabel("tipSpecified");
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
    var prefixCommentCheckbox = dialog.add("checkbox", undefined, getLabel("prefixComment"));
    prefixCommentCheckbox.helpTip = getLabel("tipPrefix");
    prefixCommentCheckbox.value = true; /* 既定でON / Default: ON */

    // --- ボタン / Buttons（Mac 規約：Cancel → OK） ---
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    prepareDialogWindow(dialog, SCRIPT_NAME);
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
