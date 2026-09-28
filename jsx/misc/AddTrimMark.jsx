#target illustrator
#targetengine "AddTrimMarkEngine"
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
var SCRIPT_VERSION  = "v1.2.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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

        var buttonRow = addButtonRow(trimDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
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

        prepareDialogWindow(trimDialog, SCRIPT_NAME);
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
