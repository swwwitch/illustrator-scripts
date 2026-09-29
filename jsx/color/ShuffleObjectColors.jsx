#target illustrator
#targetengine "ShuffleObjectColorsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト（パス、グループ、複合シェイプ）の塗り・線カラーを再配色します。
RGB／CMYK／グレースケール／特色／グラデーション／パターンに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShuffleObjectColors.md

### Overview

Reshuffles the fill and stroke colors of the selected paths, groups and compound shapes.
RGB, CMYK, grayscale, spot colors, gradients and patterns are all supported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShuffleObjectColors.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ShuffleObjectColors";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShuffleObjectColors.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShuffleObjectColors.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // バージョンとローカライズ
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* ラベル定義 / Labels */
    var LABELS = {
        /* --- ダイアログ / Dialog --- */
        dialogTitle: { ja: "選択オブジェクトのカラーをシャッフル", en: "Shuffle Selected Object Colors" },

        /* --- パネル / Panels --- */
        panelTarget: { ja: "対象", en: "Target" },
        panelExclude: { ja: "除外", en: "Exclude" },

        /* --- 対象オプション / Target options --- */
        fill: { ja: "塗り", en: "Fill" },
        stroke: { ja: "線", en: "Stroke" },

        /* --- 除外オプション / Exclude options --- */
        black: { ja: "黒", en: "Black" },
        white: { ja: "白", en: "White" },

        /* --- 適用オプション / Apply options --- */
        random: { ja: "ランダム順", en: "Random order" },
        balance: { ja: "バランスを保持", en: "Preserve balance" },

        /* --- ボタン / Buttons --- */
        apply: { ja: "適用", en: "Apply" },
        cancel: { ja: "キャンセル", en: "Cancel" },

        /* --- ツールチップ / Tooltips --- */
        tipFill: { ja: "塗りカラーをシャッフルの対象にする", en: "Include fill colors in the shuffle" },
        tipStroke: { ja: "線カラーをシャッフルの対象にする", en: "Include stroke colors in the shuffle" },
        tipBlack: { ja: "黒（および近いカラー）は変更しない", en: "Leave black (and near-black) colors untouched" },
        tipWhite: { ja: "白は変更しない", en: "Leave white untouched" },
        tipRandom: { ja: "ON：ランダム順／OFF：元の順序で適用", en: "ON: random order / OFF: keep original order" },
        tipBalance: { ja: "同じカラーの出現回数を保持して配色", en: "Keep the original frequency of each color" },
        tipApply: { ja: "ダイアログを閉じずにプレビュー", en: "Preview without closing the dialog" },

        /* --- アラート / Alerts --- */
        alertSelect: { ja: "オブジェクトを選択してください。", en: "Please select some objects." },
        alertEmpty: { ja: "使用可能なカラーが見つかりません（白と黒以外）。", en: "No usable colors found (excluding white and black)." },
        alertChoice: { ja: "塗りまたは線のいずれかを選択してください。", en: "Please select either fill or stroke." }
    };

    // =========================================
    // UI 状態 / UI state
    // =========================================

    var fillCheckbox, strokeCheckbox;
    var blackCheckbox, whiteCheckbox;
    var randomCheckbox, balanceCheckbox;
    var previewApplied = false;

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    main();

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
        if (app.documents.length === 0 || !app.activeDocument.selection || app.activeDocument.selection.length === 0) {
            alert(getLabel("alertSelect"));
            return;
        }
        showDialog();
    }

    function showDialog() {
        var dialog = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "row";
        dialog.alignChildren = "top";

        var leftGroup = dialog.add("group");
        leftGroup.orientation = "column";
        leftGroup.alignChildren = "left";

        buildTargetPanel(leftGroup);
        buildExcludePanel(leftGroup);

        randomCheckbox = leftGroup.add("checkbox", undefined, getLabel("random"));
        randomCheckbox.value = true;
        randomCheckbox.helpTip = getLabel("tipRandom");

        balanceCheckbox = leftGroup.add("checkbox", undefined, getLabel("balance"));
        balanceCheckbox.value = true;
        balanceCheckbox.helpTip = getLabel("tipBalance");

        buildButtons(dialog);

        prepareDialogWindow(dialog, SCRIPT_NAME);
        dialog.show();
    }

    function buildTargetPanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panelTarget"));
        panel.preferredSize.width = 120;
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = [10, 25, 10, 10];

        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = "left";

        fillCheckbox = row.add("checkbox", undefined, getLabel("fill"));
        fillCheckbox.value = true;
        fillCheckbox.helpTip = getLabel("tipFill");

        strokeCheckbox = row.add("checkbox", undefined, getLabel("stroke"));
        strokeCheckbox.value = true;
        strokeCheckbox.helpTip = getLabel("tipStroke");
    }

    function buildExcludePanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panelExclude"));
        panel.preferredSize.width = 120;
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = [10, 25, 10, 10];

        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = "left";

        blackCheckbox = row.add("checkbox", undefined, getLabel("black"));
        blackCheckbox.value = true;
        blackCheckbox.helpTip = getLabel("tipBlack");

        whiteCheckbox = row.add("checkbox", undefined, getLabel("white"));
        whiteCheckbox.value = true;
        whiteCheckbox.helpTip = getLabel("tipWhite");
    }

    function buildButtons(dialog) {
        var rightGroup = dialog.add("group");
        rightGroup.orientation = "column";
        rightGroup.alignChildren = "right";

        /* Mac 規約: 縦並びは OK が上 / Vertical: OK on top */
        var okBtn = rightGroup.add("button", undefined, "OK");
        okBtn.preferredSize.width = 80;

        var cancelBtn = rightGroup.add("button", undefined, getLabel("cancel"));
        cancelBtn.preferredSize.width = 80;

        // スペーサー（縦に伸びる）
        var verticalSpacer = rightGroup.add("statictext", undefined, "");
        verticalSpacer.alignment = ["fill", "fill"];

        // 適用ボタン用グループ（右下）
        var applyGroup = rightGroup.add("group");
        applyGroup.orientation = "column";
        applyGroup.alignChildren = ["fill", "bottom"];

        var applyBtn = applyGroup.add("button", undefined, getLabel("apply"));
        applyBtn.preferredSize.width = 80;
        applyBtn.helpTip = getLabel("tipApply");

        applyBtn.onClick = function () {
            applyBtn.enabled = false;
            applyWithPreview();
            applyBtn.enabled = true;
        };

        okBtn.onClick = function () {
            applyWithPreview();
            dialog.close(1);
        };

        cancelBtn.onClick = function () {
            if (previewApplied) {
                app.undo();
                previewApplied = false;
            }
            dialog.close(0);
        };
    }

    function applyWithPreview() {
        if (previewApplied) {
            app.undo();
            previewApplied = false;
        }
        if (applyColors()) {
            previewApplied = true;
            app.redraw();
        }
    }

    // =========================================
    // カラー収集 / Color collection
    // =========================================

    function applyColors() {
        var changeFill = fillCheckbox.value;
        var changeStroke = strokeCheckbox.value;
        var excludeBlack = blackCheckbox.value;
        var excludeWhite = whiteCheckbox.value;
        var balancePreserve = balanceCheckbox.value;
        var randomize = randomCheckbox.value;

        if (!changeFill && !changeStroke) {
            alert(getLabel("alertChoice"));
            return false;
        }

        /* グループ・複合パスの中のパスまで集める / Collect paths inside groups and compound paths */
        var pathItems = collectSelectionPathItems(app.activeDocument.selection);

        var colorPool = buildColorPool(pathItems, changeFill, changeStroke, excludeBlack, excludeWhite, balancePreserve);
        if (colorPool.length === 0) {
            alert(getLabel("alertEmpty"));
            return false;
        }

        var shuffled = randomize ? shuffleArray(colorPool) : colorPool;
        reapplyColors(pathItems, shuffled, changeFill, changeStroke, excludeBlack, excludeWhite);
        return true;
    }

    function buildColorPool(pathItems, changeFill, changeStroke, excludeBlack, excludeWhite, balancePreserve) {
        var colorCountMap = {};
        var uniqueColors = {};

        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            if (changeFill && pathItem.filled) {
                addColorToPool(pathItem.fillColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve);
            }
            if (changeStroke && pathItem.stroked && pathItem.strokeWidth > 0) {
                addColorToPool(pathItem.strokeColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve);
            }
        }

        var pool = [];
        if (balancePreserve) {
            for (var countKey in colorCountMap) {
                var entry = colorCountMap[countKey];
                for (var r = 0; r < entry.count; r++) {
                    pool.push(cloneColor(entry.color));
                }
            }
        } else {
            for (var uniqueKey in uniqueColors) {
                pool.push(cloneColor(uniqueColors[uniqueKey]));
            }
        }
        return pool;
    }

    function addColorToPool(rawColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve) {
        var color = cloneColor(rawColor);
        if (!color) return;
        if (excludeBlack && isNearBlack(color)) return;
        if (excludeWhite && isPureWhite(color)) return;

        var colorKey = getColorKey(color);
        if (!colorKey) return;

        if (!uniqueColors[colorKey]) {
            uniqueColors[colorKey] = cloneColor(color);
        }
        if (balancePreserve) {
            if (!colorCountMap[colorKey]) {
                colorCountMap[colorKey] = { color: cloneColor(color), count: 0 };
            }
            colorCountMap[colorKey].count++;
        }
    }

    // =========================================
    // カラー適用 / Color application
    // =========================================

    function reapplyColors(pathItems, colorPool, changeFill, changeStroke, excludeBlack, excludeWhite) {
        var colorIndex = 0;
        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            colorIndex = applyToSlot(pathItem, true, changeFill, colorPool, colorIndex, excludeBlack, excludeWhite);
            colorIndex = applyToSlot(pathItem, false, changeStroke, colorPool, colorIndex, excludeBlack, excludeWhite);
        }
    }

    function applyToSlot(pathItem, isFill, changeFlag, colorPool, colorIndex, excludeBlack, excludeWhite) {
        if (!changeFlag) return colorIndex;

        var currentColor = cloneColor(isFill ? pathItem.fillColor : pathItem.strokeColor);
        if (!currentColor) return colorIndex;
        if (excludeBlack && isNearBlack(currentColor)) return colorIndex;
        if (excludeWhite && isPureWhite(currentColor)) return colorIndex;

        var nextColor = cloneColor(colorPool[colorIndex % colorPool.length]);
        if (isFill) {
            pathItem.filled = true;
            pathItem.fillColor = nextColor;
        } else {
            pathItem.stroked = true;
            pathItem.strokeColor = nextColor;
        }
        return colorIndex + 1;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    // Fisher–Yates シャッフル / Fisher–Yates shuffle
    function shuffleArray(arr) {
        var shuffled = arr.slice();
        for (var i = shuffled.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = shuffled[i];
            shuffled[i] = shuffled[j];
            shuffled[j] = temp;
        }
        return shuffled;
    }

    /**
     * 色オブジェクトのプロパティを写す
     * 種類によっては持っていないプロパティがあり、代入で例外になるため1つずつ受け流す。
     * @param {object} targetColor - 写し先の色
     * @param {object} sourceColor - 写し元の色
     * @param {string[]} propertyNames - 写すプロパティ名
     * @returns {void}
     */
    function copyColorProperties(targetColor, sourceColor, propertyNames) {
        for (var i = 0; i < propertyNames.length; i++) {
            try {
                targetColor[propertyNames[i]] = sourceColor[propertyNames[i]];
            } catch (e) {}
        }
    }

    function getColorKey(color) {
        if (color.typename === "RGBColor") {
            return "rgb:" + color.red + "," + color.green + "," + color.blue;
        }
        if (color.typename === "CMYKColor") {
            return "cmyk:" + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black;
        }
        if (color.typename === "GrayColor") {
            return "gray:" + color.gray;
        }
        if (color.typename === "SpotColor") {
            return "spot:" + color.spot.name + "@" + color.tint;
        }
        if (color.typename === "GradientColor") {
            return "grad:" + color.gradient.name;
        }
        if (color.typename === "PatternColor") {
            return "pat:" + color.pattern.name;
        }
        return null;
    }

    // カラー複製（RGB/CMYK/Gray/Spot/Gradient/Pattern 対応） / Clone color (RGB/CMYK/Gray/Spot/Gradient/Pattern)
    function cloneColor(color) {
        if (!color || !color.typename) return null;
        if (color.typename === "RGBColor") {
            var rgb = new RGBColor();
            rgb.red = color.red;
            rgb.green = color.green;
            rgb.blue = color.blue;
            return rgb;
        }
        if (color.typename === "CMYKColor") {
            var cmyk = new CMYKColor();
            cmyk.cyan = color.cyan;
            cmyk.magenta = color.magenta;
            cmyk.yellow = color.yellow;
            cmyk.black = color.black;
            return cmyk;
        }
        if (color.typename === "GrayColor") {
            var gr = new GrayColor();
            gr.gray = color.gray;
            return gr;
        }
        if (color.typename === "SpotColor") {
            /* 見当合わせ色（レジストレーション）はシャッフル対象外 / Skip registration */
            if (color.spot.colorType === ColorModel.REGISTRATION) return null;
            var sc = new SpotColor();
            sc.spot = color.spot;
            sc.tint = color.tint;
            return sc;
        }
        if (color.typename === "GradientColor") {
            var gc = new GradientColor();
            gc.gradient = color.gradient;
            copyColorProperties(gc, color, ["angle", "length", "origin", "hiliteAngle", "hiliteLength", "matrix"]);
            return gc;
        }
        if (color.typename === "PatternColor") {
            var pc = new PatternColor();
            pc.pattern = color.pattern;
            copyColorProperties(pc, color, ["matrix", "shiftAngle", "shiftDistance", "reflect",
                "reflectAngle", "rotation", "scaleFactor", "shearAngle", "shearAxis"]);
            return pc;
        }
        return null;
    }

    // 黒近似判定 / Near-black detection
    function isNearBlack(color) {
        if (color.typename === "RGBColor") {
            return color.red <= 51 && color.green <= 51 && color.blue <= 51;
        }
        if (color.typename === "CMYKColor") {
            return color.black >= 0.8 || (color.cyan <= 0.2 && color.magenta <= 0.2 && color.yellow <= 0.2 && color.black >= 0.7);
        }
        if (color.typename === "GrayColor") {
            return color.gray >= 80;
        }
        if (color.typename === "SpotColor") {
            /* tint 100% のときだけ基底色で判定 / Only when tint is 100% */
            return color.tint >= 99.999 && isNearBlack(color.spot.color);
        }
        return false;
    }

    // 純白判定 / Pure-white detection
    function isPureWhite(color) {
        if (color.typename === "RGBColor") {
            return color.red === 255 && color.green === 255 && color.blue === 255;
        }
        if (color.typename === "CMYKColor") {
            return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
        }
        if (color.typename === "GrayColor") {
            return color.gray === 0;
        }
        if (color.typename === "SpotColor") {
            /* tint 0% は紙色（白） / Tint 0% = paper (white) */
            if (color.tint <= 0.001) return true;
            return color.tint >= 99.999 && isPureWhite(color.spot.color);
        }
        return false;
    }

})();
