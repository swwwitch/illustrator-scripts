#target illustrator
#targetengine "SwapNearestItemWithDialogboxEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトと、指定した方向（上下左右）にある最も近いオブジェクトの位置を入れ替えます。方向はダイアログで指定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItemWithDialogbox.md

### Overview

Swaps the selected object with the nearest object in a chosen direction, picked from a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItemWithDialogbox.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapNearestItemWithDialogbox"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItemWithDialogbox.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItemWithDialogbox.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Basic settings
    // =========================================

    /* 矢印キーと探索方向の対応 / Arrow key to search direction */
    var DIRECTION_BY_ARROW_KEY = {
        Up: "up",
        Down: "down",
        Left: "left",
        Right: "right"
    };

    /* 入れ替えの対象にできる typename / Typenames that can take part in a swap */
    var SWAPPABLE_TYPENAMES = [
        "PathItem", "CompoundPathItem", "GroupItem", "TextFrame",
        "PlacedItem", "RasterItem", "SymbolItem", "MeshItem",
        "PluginItem", "GraphItem"
    ];

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトの位置を入れ替え", en: "Swap Positions" }
        },
        status: {
            prompt: { ja: "矢印キーで入れ替え", en: "Press an arrow key to swap" },
            notFound: { ja: "見つかりません", en: "Nothing found" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            prompt: {
                ja: "矢印キーを押すと、その方向にある最も近いオブジェクトと位置を入れ替えます。Esc または Enter で閉じます。",
                en: "An arrow key swaps the selection with the nearest object in that direction. Esc or Enter closes this dialog."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOne: { ja: "1つのオブジェクトを選択してください。", en: "Select one object." },
            noPosition: { ja: "位置情報を持つオブジェクトを選択してください。", en: "Select an object that has a position." }
        }
    };

    // =========================================
    // 判定・計測 / Checks and measurement
    // =========================================

    /**
     * 入れ替えの対象にできるオブジェクトか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象にできる場合はtrue
     */
    function isSwappableItem(pageItem) {
        /* ガイドは typename が "PathItem" のため guides で判定する / Guides are PathItems, so test .guides */
        if (pageItem.typename === "PathItem" && pageItem.guides === true) return false;

        for (var i = 0; i < SWAPPABLE_TYPENAMES.length; i++) {
            if (SWAPPABLE_TYPENAMES[i] === pageItem.typename) return true;
        }
        return false;
    }

    /**
     * ロック・非表示のレイヤーに属していないか判定する
     * @param {Layer} layer - 判定するレイヤー
     * @returns {boolean} 触れる場合はtrue
     */
    function isUsableLayer(layer) {
        /* サブレイヤーは自身がロックされていなくても親レイヤー側でロックされることがある
           A sublayer can be locked or hidden by an ancestor layer even when its own flags are clear */
        for (var layerNode = layer; layerNode && layerNode.typename === "Layer"; layerNode = layerNode.parent) {
            if (layerNode.locked || layerNode.visible === false) return false;
        }
        return true;
    }

    /**
     * 配列に指定したオブジェクトが含まれるか判定する
     * @param {PageItem[]} candidateItems - 探す対象の配列
     * @param {PageItem} targetItem - 探すオブジェクト
     * @returns {boolean} 含まれる場合はtrue
     */
    function containsItem(candidateItems, targetItem) {
        if (!candidateItems) return false;
        for (var i = 0; i < candidateItems.length; i++) {
            if (candidateItems[i] === targetItem) return true;
        }
        return false;
    }

    /**
     * オブジェクトの境界を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {number[]} [left, top, right, bottom]
     */
    function getBounds(pageItem) {
        return pageItem.visibleBounds;
    }

    /**
     * オブジェクトの中心座標を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {number[]} [x, y]
     */
    function getCenter(pageItem) {
        var itemBounds = getBounds(pageItem);
        return [(itemBounds[0] + itemBounds[2]) / 2, (itemBounds[1] + itemBounds[3]) / 2];
    }

    // =========================================
    // 探索と入れ替え / Search and swap
    // =========================================

    /**
     * 候補として見てよいオブジェクトか判定する
     * @param {PageItem} candidateItem - 判定するオブジェクト
     * @param {PageItem} referenceItem - 基準オブジェクト
     * @param {PageItem[]} excludeItems - 候補から除外するオブジェクト（省略可）
     * @returns {boolean} 候補にできる場合はtrue
     */
    function isSearchCandidate(candidateItem, referenceItem, excludeItems) {
        if (candidateItem === referenceItem) return false;
        if (containsItem(excludeItems, candidateItem)) return false;
        if (candidateItem.locked || candidateItem.hidden) return false;
        if (!isUsableLayer(candidateItem.layer)) return false;
        if (!isSwappableItem(candidateItem)) return false;
        /* グループや複合パス内の子オブジェクトは除外（親のみ処理対象）/ Skip children of groups and compound paths */
        if (candidateItem.parent &&
            (candidateItem.parent.typename === "GroupItem" || candidateItem.parent.typename === "CompoundPathItem")) return false;
        return true;
    }

    /**
     * 探索方向における2つの境界の隙間を返す
     * 探索軸と直交する方向で範囲が重なっていない場合や、方向が合わない場合は null を返す。
     * @param {number[]} referenceBounds - 基準オブジェクトの境界
     * @param {number[]} candidateBounds - 候補オブジェクトの境界
     * @param {number} dx - 中心X の差（候補 − 基準）
     * @param {number} dy - 中心Y の差（候補 − 基準）
     * @param {string} searchDirection - 探索方向（"right" / "left" / "up" / "down"）
     * @returns {number|null} 隙間（重なっている場合は0）。対象外なら null
     */
    function getDirectionalGap(referenceBounds, candidateBounds, dx, dy, searchDirection) {
        /* 探索軸と直交する方向で範囲が重なっているか / Overlap on the axis perpendicular to the search */
        var verticalOverlap = !(referenceBounds[3] > candidateBounds[1] || referenceBounds[1] < candidateBounds[3]);
        var horizontalOverlap = !(referenceBounds[0] > candidateBounds[2] || referenceBounds[2] < candidateBounds[0]);

        var gap;
        if (searchDirection === "right" && dx > 0 && verticalOverlap) {
            gap = candidateBounds[0] - referenceBounds[2];
        } else if (searchDirection === "left" && dx < 0 && verticalOverlap) {
            gap = referenceBounds[0] - candidateBounds[2];
        } else if (searchDirection === "up" && dy > 0 && horizontalOverlap) {
            gap = candidateBounds[3] - referenceBounds[1];
        } else if (searchDirection === "down" && dy < 0 && horizontalOverlap) {
            gap = referenceBounds[3] - candidateBounds[1];
        } else {
            return null;
        }

        /* 重なっている場合は隙間0として扱う / Treat overlapping items as a zero gap */
        return (gap < 0) ? 0 : gap;
    }

    /**
     * 指定方向にある最も近いオブジェクトを検索する
     * 近さは探索方向の隙間で測る（中心間の直線距離では斜めのオブジェクトが勝ってしまう）。
     * @param {PageItem} referenceItem - 基準オブジェクト
     * @param {string} searchDirection - 探索方向（"right" / "left" / "up" / "down"）
     * @param {PageItem[]} excludeItems - 候補から除外するオブジェクト（省略可）
     * @returns {PageItem|null} 見つかったオブジェクト。なければnull
     */
    function findNearestObjectInDirection(referenceItem, searchDirection, excludeItems) {
        var documentItems = app.activeDocument.pageItems;
        var referenceBounds = getBounds(referenceItem);
        var referenceCenter = getCenter(referenceItem);

        var nearestItem = null;
        var minGap = Number.MAX_VALUE;
        var minDistance = Number.MAX_VALUE;

        for (var i = 0; i < documentItems.length; i++) {
            var candidateItem = documentItems[i];
            if (!isSearchCandidate(candidateItem, referenceItem, excludeItems)) continue;

            var candidateCenter = getCenter(candidateItem);
            var dx = candidateCenter[0] - referenceCenter[0];
            var dy = candidateCenter[1] - referenceCenter[1];

            var gap = getDirectionalGap(referenceBounds, getBounds(candidateItem), dx, dy, searchDirection);
            if (gap === null) continue;

            var distance = Math.sqrt(dx * dx + dy * dy);
            if (gap < minGap || (gap === minGap && distance < minDistance)) {
                minGap = gap;
                minDistance = distance;
                nearestItem = candidateItem;
            }
        }

        return nearestItem;
    }

    /**
     * 2つのオブジェクトを左右に入れ替える（全体の占有範囲と隙間を保つ）
     * @param {PageItem} itemA - 入れ替える一方
     * @param {PageItem} itemB - 入れ替えるもう一方
     * @returns {void}
     */
    function swapHorizontally(itemA, itemB) {
        var boundsA = getBounds(itemA);
        var boundsB = getBounds(itemB);

        /* itemAが左、itemBが右になるように並び替え / Order so that itemA is the left one */
        if (boundsA[0] > boundsB[0]) {
            var tempItem = itemA;
            itemA = itemB;
            itemB = tempItem;

            var tempBounds = boundsA;
            boundsA = boundsB;
            boundsB = tempBounds;
        }

        var widthB = boundsB[2] - boundsB[0];
        var gap = boundsB[0] - boundsA[2];

        itemB.left = boundsA[0];
        itemA.left = boundsA[0] + widthB + gap;
    }

    /**
     * 2つのオブジェクトを上下に入れ替える（全体の占有範囲と隙間を保つ）
     * @param {PageItem} itemA - 入れ替える一方
     * @param {PageItem} itemB - 入れ替えるもう一方
     * @returns {void}
     */
    function swapVertically(itemA, itemB) {
        var boundsA = getBounds(itemA);
        var boundsB = getBounds(itemB);

        /* itemAが上、itemBが下になるように並び替え / Order so that itemA is the upper one */
        if (boundsA[1] < boundsB[1]) {
            var tempItem = itemA;
            itemA = itemB;
            itemB = tempItem;

            var tempBounds = boundsA;
            boundsA = boundsB;
            boundsB = tempBounds;
        }

        var heightB = boundsB[1] - boundsB[3];
        var gap = boundsA[3] - boundsB[1];

        itemB.top = boundsA[1];
        itemA.top = boundsA[1] - heightB - gap;
    }

    /**
     * 選択中のオブジェクトと、探索方向で最も近いオブジェクトの位置を入れ替える
     * @param {string} searchDirection - 探索方向（"right" / "left" / "up" / "down"）
     * @returns {boolean} 入れ替えた場合はtrue
     */
    function swapWithNearestObject(searchDirection) {
        var selectedObjects = app.activeDocument.selection;
        /* 実行中に選択が変わることがあるため、毎回確認する（案内はmain側で済ませている）
           The selection can change while the dialog is open */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length !== 1) return false;

        var baseItem = selectedObjects[0];
        if (!isSwappableItem(baseItem)) return false;

        var nearestItem = findNearestObjectInDirection(baseItem, searchDirection, selectedObjects);
        if (!nearestItem) return false;

        if (searchDirection === "up" || searchDirection === "down") {
            swapVertically(baseItem, nearestItem);
        } else {
            swapHorizontally(baseItem, nearestItem);
        }
        return true;
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 矢印キーで入れ替えを実行するダイアログを表示する
     * @returns {void}
     */
    function showDialog() {
        var swapDialog = new Window("dialog", getLabel("dialog.title"));

        var statusText = swapDialog.add("statictext", undefined, getLabel("status.prompt"));
        statusText.alignment = "center";
        statusText.helpTip = getLabel("tooltip.prompt");

        var btnClose = swapDialog.add("button", undefined, getLabel("button.close"));
        btnClose.onClick = function () {
            swapDialog.close();
        };

        /* キーリピート中に多重実行しないための門 / Gate out repeats while one swap is running */
        var isSwapInProgress = false;

        swapDialog.addEventListener("keydown", function (event) {
            if (isSwapInProgress) return;

            if (event.keyName === "Escape" || event.keyName === "Enter" || event.keyName === "Return") {
                btnClose.notify();
                return;
            }

            var keyedDirection = DIRECTION_BY_ARROW_KEY[event.keyName];
            if (!keyedDirection) return;

            isSwapInProgress = true;

            /* 見つからないときはメッセージを差し替えて知らせる。連打を邪魔しないようアラートは出さない
               Report a miss by swapping the message; an alert would interrupt repeated key presses */
            statusText.text = swapWithNearestObject(keyedDirection) ? getLabel("status.prompt") : getLabel("status.notFound");
            app.redraw();
        }, false);

        swapDialog.addEventListener("keyup", function () {
            isSwapInProgress = false;
        }, false);

        btnClose.active = true;
        prepareDialogWindow(swapDialog, SCRIPT_NAME);
        swapDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択内容を確認してダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedObjects = doc.selection;

        /* 文字カーソルが立っていると selection は TextRange で、length は文字数になる
           A text caret gives a TextRange whose length counts characters */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length !== 1) {
            alert(getLabel("alert.selectOne"));
            return;
        }

        if (!isSwappableItem(selectedObjects[0])) {
            alert(getLabel("alert.noPosition"));
            return;
        }

        showDialog();
    }

    main();

})();
