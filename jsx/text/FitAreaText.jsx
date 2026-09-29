#target illustrator
#targetengine "FitAreaTextEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字・パス上文字のあふれ（オーバーセット）を、文字サイズの縮小・拡大、
またはエリア内文字の高さの調整で解消します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAreaText.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8c2e2568a6b7

### Overview

Resolves overset text in the selected area type and path type, either by shrinking or growing
the font size, or by adjusting the height of the area type.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAreaText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitAreaText";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.3.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAreaText.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAreaText.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8c2e2568a6b7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 選択にエリア内文字があるとき、［エリア内文字の高さ調整］を初期状態でONにするか / Turn the height mode on by default when the selection contains area type */
    var DEFAULT_HEIGHT_MODE = true;

    /* 高さ調整のオプションの初期選択（"autoSize" = 自動サイズ調整／"adjustHeight" = 高さを調整）/ Default option of the height mode */
    var DEFAULT_HEIGHT_OPTION = "autoSize";

    /* ［文字サイズ：あふれ処理］の初期状態 / Default state of "Shrink Text to Fit" */
    var DEFAULT_SHRINK_TO_FIT = true;

    /* ［文字サイズ：ぴったり］の初期状態 / Default state of "Maximize Text Size" */
    var DEFAULT_MAXIMIZE_SIZE = true;

    /* 文字サイズを縮小する刻み（pt）/ Step used when shrinking the font size */
    var FONT_SIZE_STEP = 0.1;

    /* 縮小できる最小の文字サイズ（pt）/ Smallest font size the shrink loop may reach */
    var MIN_FONT_SIZE = 0.1;

    /* 縮小処理の上限回数（安全弁）/ Safety limit for the shrink loop */
    var MAX_SHRINK_ITERATIONS = 2000;

    /* 上限回数に達したときに警告を出すか / Alert when the shrink limit is reached */
    var ALERT_ON_SHRINK_LIMIT = true;

    /* ［ぴったり］で倍々に拡大する上限回数 / Limit of the doubling loop used by "Maximize" */
    var MAX_GROW_ITERATIONS = 25;

    /* ［ぴったり］で拡大できる上限の文字サイズ（pt）/ Largest font size the grow loop may reach */
    var MAX_FONT_SIZE = 100000;

    /* 元の文字サイズ・高さを控えるタグ名（データセットごとのリセットに使う）/ Tags holding the original values */
    var FONT_SIZE_TAG_NAME = "overset_text_default_size";
    var HEIGHT_TAG_NAME = "overset_text_default_height";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS        = 15;                /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING        = 10;                /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS         = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING         = 6;                 /* パネル内の要素間隔 / panel spacing */
    var OPTION_INDENT_MARGINS = [18, 0, 0, 7];     /* 入れ子オプションの字下げ [左,上,右,下] / indent of nested options */
    var OPTION_SPACING        = 4;                 /* 入れ子オプションの間隔 / spacing of nested options */

    /**
     * ウィンドウの共通レイアウトを設定する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["left", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルを追加し、共通レイアウトを設定する
     * @param {Object} parentContainer - 追加先のウィンドウまたはグループ
     * @param {Object} titleSet - ja/en を持つパネル名
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentContainer, titleSet) {
        var newPanel = parentContainer.add("panel", undefined, getLabel(titleSet));
        newPanel.orientation = "column";
        newPanel.alignChildren = ["left", "top"];
        newPanel.alignment = ["fill", "top"];
        newPanel.margins = PANEL_MARGINS;
        newPanel.spacing = PANEL_SPACING;
        return newPanel;
    }

    /**
     * 字下げした縦並びグループを追加する（入れ子のオプション用）
     * @param {Object} parentContainer - 追加先のパネルまたはグループ
     * @returns {Group} 追加したグループ
     */
    function addIndentedColumn(parentContainer) {
        var indentedGroup = parentContainer.add("group");
        indentedGroup.orientation = "column";
        indentedGroup.alignment = ["left", "top"];
        indentedGroup.alignChildren = ["left", "top"];
        indentedGroup.margins = OPTION_INDENT_MARGINS;
        indentedGroup.spacing = OPTION_SPACING;
        return indentedGroup;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
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
            title: { ja: "テキストの自動調整", en: "Auto Fit Text Frame" }
        },
        panel: {
            processing: { ja: "調整方法", en: "Adjustment Method" }
        },
        checkbox: {
            shrinkToFit: { ja: "文字サイズ：あふれを解消", en: "Shrink Text to Fit" },
            maximizeSize: { ja: "文字サイズ：最大まで拡大", en: "Maximize Text Size" },
            heightMode: { ja: "エリア内文字の高さ調整", en: "Adjust Area Text Height" }
        },
        radio: {
            adjustHeight: { ja: "高さを広げて固定", en: "Expand and fix" },
            autoSize: { ja: "自動サイズ調整", en: "Auto size" }
        },
        tooltip: {
            shrinkToFit: {
                ja: "あふれ（オーバーセット）がなくなるまで文字サイズを縮小します。",
                en: "Shrink the font size until the text no longer oversets."
            },
            maximizeSize: {
                ja: "いったん文字サイズを拡大してから、あふれない最大サイズまで詰めます。",
                en: "Grow the font size first, then shrink it to the largest size that still fits."
            },
            heightMode: {
                ja: "文字サイズではなく、エリア内文字の高さで調整します。選択にエリア内文字があるときだけ選べます。",
                en: "Adjust the height of area type instead of the font size. Available only when the selection contains area type."
            },
            adjustHeight: {
                ja: "自動サイズ調整を一時的にONにして、必要な分だけ高さを広げてから固定します（以後は自動で変わりません）。",
                en: "Turn Auto Size on and off again, expanding the frame just enough and then fixing that height."
            },
            autoSize: {
                ja: "エリア内文字に自動サイズ調整を適用します（拡張のみ）。",
                en: "Apply Auto Size to Area Text (expand only)."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            selectObject: {
                ja: "エリア内文字またはパス上文字を選択してください。",
                en: "Please select area type or path type."
            },
            noValidText: {
                ja: "選択中に処理可能なテキストがありません。\n（グループ内のテキストはグループごと選択でもOK）",
                en: "No valid text frames found in the selection.\n(Text inside groups can be processed by selecting the group.)"
            },
            noAreaText: {
                ja: "エリア内文字が選択されていません。\n高さの調整はエリア内文字のみ対応です。",
                en: "No area type found in the selection.\nHeight adjustment is supported for area type only."
            },
            selectMode: {
                ja: "調整方法を選択してください。\n（文字サイズ：あふれを解消 / 文字サイズ：最大まで拡大）",
                en: "Please select an adjustment method.\n(Shrink Text to Fit / Maximize Text Size)"
            },
            hardReturn: {
                ja: "改行コードが含まれているテキストには対応していません。\n対象テキスト：",
                en: "Text containing line breaks is not supported.\nTarget: "
            },
            shrinkLimit: {
                ja: "文字サイズの縮小が上限回数に達しました：\n",
                en: "Shrink iteration limit reached:\n"
            }
        },
        fallbackName: {
            unnamedText: { ja: "［名前なし］", en: "[Unnamed Text]" }
        }
    };

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 調整方法 / Adjustment modes */
    var ADJUST_MODE = {
        FONT_SIZE: "fontSize",   /* 文字サイズで調整 / adjust the font size */
        HEIGHT: "height",        /* エリア内文字の高さで調整 / adjust the area type height */
        AUTO_SIZE: "autoSize"    /* 自動サイズ調整を適用 / apply Auto Size */
    };

    /* 自動サイズ調整アクションの値 / Values of the Auto Size action */
    var AUTO_SIZE_ON = 1;
    var AUTO_SIZE_OFF = 2;

    // =========================================
    // テキストの収集 / Collecting text frames
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 処理対象にできるテキストフレームか判定する
     * @param {TextFrame} textFrame - 判定するテキストフレーム
     * @returns {boolean} 対象にできるとき true
     */
    function isAdjustableTextFrame(textFrame) {
        return ((textFrame.kind == TextType.PATHTEXT || textFrame.kind == TextType.AREATEXT) &&
            textFrame.editable && !textFrame.locked && !textFrame.hidden);
    }

    /**
     * 選択から処理対象のテキストフレームを重複なく集める（グループの中・文字カーソルの選択を含む）
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} 処理対象のテキストフレーム
     */
    function getSelectedTextFrames(doc) {
        if (!doc || !doc.selection || doc.selection.length === 0) return [];
        return collectSelectionItems(doc.selection, {
            accept: function (item) {
                /* 読めないプロパティがあれば対象外 / Unreadable properties exclude the item */
                try {
                    return item.typename === "TextFrame" && isAdjustableTextFrame(item);
                } catch (e) {
                    return false;
                }
            }
        });
    }

    /**
     * テキストフレームの配列からエリア内文字だけを取り出す
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {TextFrame[]} エリア内文字のみの配列
     */
    function filterAreaTextFrames(textFrames) {
        var areaFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            if (textFrames[i].kind == TextType.AREATEXT) areaFrames.push(textFrames[i]);
        }
        return areaFrames;
    }

    /**
     * テキストフレームの表示名を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {string} 名前（未設定のときは代替名）
     */
    function getTextFrameName(textFrame) {
        return textFrame.name ? textFrame.name : getLabel(LABELS.fallbackName.unnamedText);
    }

    // =========================================
    // あふれの判定 / Overset detection
    // =========================================

    /**
     * 表示されている行に収まらない文字があるか判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function hasHiddenCharacters(textFrame) {
        var lineCount = textFrame.lines.length;
        if (lineCount === 0) return (textFrame.characters.length > 0);

        var visibleCharacters = 0;
        for (var i = 0; i < lineCount; i++) {
            visibleCharacters += textFrame.lines[i].characters.length;
        }
        return (visibleCharacters < textFrame.characters.length);
    }

    /**
     * テキストフレームがあふれているか判定する
     * エリア内文字は overflows を優先し、パス上文字は表示行の文字数で判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function isOversetFrame(textFrame) {
        try {
            if (textFrame.kind == TextType.AREATEXT && typeof textFrame.overflows !== "undefined") {
                return !!textFrame.overflows;
            }
            return hasHiddenCharacters(textFrame);
        } catch (e) {
            return false;
        }
    }

    /**
     * 改行コードを含むテキストなら警告を出す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 改行コードを含み、処理を中止すべきとき true
     */
    function stopIfHardReturn(textFrame) {
        var hasHardReturn;
        try {
            hasHardReturn = /[\r\n]/.test(textFrame.contents);
        } catch (e) {
            return false;
        }
        if (hasHardReturn) {
            alert(getLabel(LABELS.alert.hardReturn) + getTextFrameName(textFrame));
            return true;
        }
        return false;
    }

    // =========================================
    // 元の値の記録とリセット / Recording and resetting the original values
    // =========================================

    /**
     * タグを名前で探す（getByName は見つからないと例外になるためここで受ける）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {Tag|null} 見つかったタグ（ないときは null）
     */
    function findTag(textFrame, tagName) {
        try {
            return textFrame.tags.getByName(tagName);
        } catch (e) {
            return null;
        }
    }

    /**
     * 値をタグに書き込む（同名のタグがあれば上書き）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @param {number} value - 控える値
     * @returns {void}
     */
    function saveValueToTag(textFrame, tagName, value) {
        var valueTag = findTag(textFrame, tagName);
        if (!valueTag) {
            valueTag = textFrame.tags.add();
            valueTag.name = tagName;
        }
        valueTag.value = value;
    }

    /**
     * タグに控えた値を読み出す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {number|null} 控えた値（ないときは null）
     */
    function readValueFromTag(textFrame, tagName) {
        var valueTag = findTag(textFrame, tagName);
        return valueTag ? (valueTag.value * 1) : null;
    }

    /**
     * タグを削除する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {void}
     */
    function removeValueTag(textFrame, tagName) {
        var valueTag = findTag(textFrame, tagName);
        if (valueTag) valueTag.remove();
    }

    /**
     * 処理中のデータセットが1件目か判定する
     * @param {Document} doc - 対象のドキュメント
     * @returns {boolean} 1件目のとき true
     */
    function isFirstDataSet(doc) {
        return (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[0]);
    }

    /**
     * 処理中のデータセットが最後の1件か判定する
     * @param {Document} doc - 対象のドキュメント
     * @returns {boolean} 最後の1件のとき true
     */
    function isLastDataSet(doc) {
        return (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[doc.dataSets.length - 1]);
    }

    /**
     * データセットの1件目なら元の値をタグに控え、控えた値があれば毎回そこへ戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @param {Function} readCurrentValue - テキストフレームから現在値を読む処理
     * @param {Function} writeSavedValue - テキストフレームへ値を書き戻す処理
     * @returns {void}
     */
    function resetToOriginalValue(doc, textFrames, tagName, readCurrentValue, writeSavedValue) {
        var i;
        if (isFirstDataSet(doc)) {
            for (i = 0; i < textFrames.length; i++) {
                saveValueToTag(textFrames[i], tagName, readCurrentValue(textFrames[i]));
            }
        }
        for (i = 0; i < textFrames.length; i++) {
            if (textFrames[i].contents === "") continue;
            var savedValue = readValueFromTag(textFrames[i], tagName);
            if (savedValue !== null) writeSavedValue(textFrames[i], savedValue);
        }
    }

    /**
     * 最後のデータセットまで終わったらタグを片付ける
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {void}
     */
    function removeTagsAfterLastDataSet(doc, textFrames, tagName) {
        if (!isLastDataSet(doc)) return;
        for (var i = 0; i < textFrames.length; i++) {
            removeValueTag(textFrames[i], tagName);
        }
    }

    // =========================================
    // 文字サイズの調整 / Adjusting the font size
    // =========================================

    /**
     * 文字サイズを読む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} 文字サイズ（pt）
     */
    function getFontSize(textFrame) {
        return textFrame.textRange.characterAttributes.size;
    }

    /**
     * 文字サイズを書き込む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 文字サイズ（pt）
     * @returns {void}
     */
    function setFontSize(textFrame, fontSize) {
        textFrame.textRange.characterAttributes.size = fontSize;
    }

    /**
     * 手動行送りのときだけ、行送りと文字サイズの比率を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number|null} 行送りの比率（自動行送りのときは null）
     */
    function getLeadingRatio(textFrame) {
        try {
            var textAttributes = textFrame.textRange.characterAttributes;
            if (textAttributes.autoLeading) return null;
            if (textAttributes.size > 0 && textAttributes.leading > 0) return textAttributes.leading / textAttributes.size;
        } catch (e) { }
        return null;
    }

    /**
     * 文字サイズに合わせて行送りを追従させる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 変更後の文字サイズ（pt）
     * @param {number|null} leadingRatio - 行送りの比率（null のときは何もしない）
     * @returns {void}
     */
    function applyLeading(textFrame, fontSize, leadingRatio) {
        if (leadingRatio === null) return;
        try {
            textFrame.textRange.characterAttributes.leading = fontSize * leadingRatio;
        } catch (e) { }
    }

    /**
     * 文字サイズを書き込み、手動行送りなら比率を保って追従させる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {number|null} leadingRatio - 行送りの比率（null のときは行送りに触らない）
     * @returns {void}
     */
    function setFontSizeWithLeading(textFrame, fontSize, leadingRatio) {
        setFontSize(textFrame, fontSize);
        applyLeading(textFrame, fontSize, leadingRatio);
    }

    /**
     * あふれがなくなるまで文字サイズを縮小する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 続行してよいとき true（改行コードを含むときは false）
     */
    function shrinkFontToFit(textFrame) {
        if (stopIfHardReturn(textFrame)) return false;
        if (textFrame.characters.length <= 0 || !isOversetFrame(textFrame)) return true;

        var leadingRatio = getLeadingRatio(textFrame);
        var iteration = 0;
        while (isOversetFrame(textFrame)) {
            var currentSize = getFontSize(textFrame);
            if (currentSize <= MIN_FONT_SIZE) break;

            var reducedSize = Math.max(MIN_FONT_SIZE, currentSize - FONT_SIZE_STEP);
            setFontSizeWithLeading(textFrame, reducedSize, leadingRatio);

            iteration++;
            if (iteration >= MAX_SHRINK_ITERATIONS) {
                if (ALERT_ON_SHRINK_LIMIT) alert(getLabel(LABELS.alert.shrinkLimit) + getTextFrameName(textFrame));
                break;
            }
        }
        return true;
    }

    /**
     * いったんあふれるまで拡大してから、あふれない最大サイズまで詰める
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 続行してよいとき true（改行コードを含むときは false）
     */
    function maximizeFontToFit(textFrame) {
        if (stopIfHardReturn(textFrame)) return false;
        if (textFrame.characters.length <= 0) return true;

        var leadingRatio = getLeadingRatio(textFrame);
        var originalSize = getFontSize(textFrame);

        /* あふれるまで倍々に拡大する / Double the size until it oversets */
        var grownSize = originalSize;
        for (var i = 0; i < MAX_GROW_ITERATIONS && !isOversetFrame(textFrame); i++) {
            grownSize = grownSize * 2;
            if (grownSize > MAX_FONT_SIZE) break;
            setFontSizeWithLeading(textFrame, grownSize, leadingRatio);
        }

        /* それでもあふれないときは元に戻す / Restore the original size when it never oversets */
        if (!isOversetFrame(textFrame)) {
            setFontSizeWithLeading(textFrame, originalSize, leadingRatio);
            return true;
        }

        return shrinkFontToFit(textFrame);
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 高さの調整（エリア内文字）/ Adjusting the height (area type)
    // =========================================

    /**
     * 自動サイズ調整をアクション経由で切り替える
     * @param {number} autoSizeValue - AUTO_SIZE_ON（ON）または AUTO_SIZE_OFF（OFF）
     * @returns {void}
     */
    function setAutoSizeByAction(autoSizeValue) {
        /* アクション定義（セット名 AreaType／アクション名 AutoSize）/ Action definition */
        var actionCode = [
            '/version 3',
            '/name [ 8 4172656154797065]',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '  /name [ 8 4175746f53697a65 ]',
            '  /keyIndex 0',
            '  /colorIndex 0',
            '  /isOpen 1',
            '  /eventCount 1',
            '  /event-1 {',
            '    /useRulersIn1stQuadrant 0',
            '    /internalName (adobe_SLOAreaTextDialog)',
            '    /localizedName [ 33',
            '      e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3',
            '    ]',
            '    /isOpen 1',
            '    /isOn 1',
            '    /hasDialog 0',
            '    /parameterCount 1',
            '    /parameter-1 {',
            '      /key 1952539754',
            '      /showInPalette 4294967295',
            '      /type (integer)',
            '      /value ' + String(autoSizeValue),
            '    }',
            '  }',
            '}'
        ].join("\n");

        /* 失敗したら従来どおり例外で処理を止める（解除と一時ファイルの削除は済んでいる）
           On failure, stop with an exception as before (the set and temp file are already cleaned up) */
        if (!runTemporaryAction(actionCode, "AreaType", "AutoSize")) {
            throw new Error("AutoSize action failed");
        }
    }

    /**
     * エリア内文字に自動サイズ調整を適用する（拡張のみ・OFFには戻さない）
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function applyAutoSize(textFrame) {
        app.activeDocument.selection = [textFrame];
        setAutoSizeByAction(AUTO_SIZE_ON);
    }

    /**
     * 自動サイズ調整をON→OFFして、必要な分だけ高さを広げて固定する
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function adjustHeightToFit(textFrame) {
        if (textFrame.characters.length <= 0) return;
        applyAutoSize(textFrame);
        setAutoSizeByAction(AUTO_SIZE_OFF);
    }

    // =========================================
    // 実行 / Processing
    // =========================================

    /**
     * エリア内文字だけを取り出す（1つもなければ警告する）
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @returns {TextFrame[]|null} エリア内文字の配列（1つもないときは null）
     */
    function getAreaTextTargets(textFrames) {
        var areaFrames = filterAreaTextFrames(textFrames);
        if (areaFrames.length === 0) {
            alert(getLabel(LABELS.alert.noAreaText));
            return null;
        }
        return areaFrames;
    }

    /**
     * テキストフレームを順に処理する（中止が返ったらそこで止める）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {Function} adjustFrame - 1つのテキストフレームを処理する関数
     * @returns {boolean} 最後まで処理できたとき true
     */
    function adjustEachFrame(textFrames, adjustFrame) {
        for (var i = 0; i < textFrames.length; i++) {
            if (adjustFrame(textFrames[i]) === false) return false;
        }
        return true;
    }

    /**
     * 自動サイズ調整を適用する（エリア内文字のみ）
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @returns {void}
     */
    function runAutoSize(textFrames) {
        var areaFrames = getAreaTextTargets(textFrames);
        if (!areaFrames) return;
        adjustEachFrame(areaFrames, applyAutoSize);
    }

    /**
     * エリア内文字の高さを中身に合わせる
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @returns {void}
     */
    function runHeightAdjust(doc, textFrames) {
        var areaFrames = getAreaTextTargets(textFrames);
        if (!areaFrames) return;

        resetToOriginalValue(doc, areaFrames, HEIGHT_TAG_NAME,
            function(textFrame) { return textFrame.height; },
            function(textFrame, height) { textFrame.height = height; });

        adjustEachFrame(areaFrames, adjustHeightToFit);
        removeTagsAfterLastDataSet(doc, areaFrames, HEIGHT_TAG_NAME);
    }

    /**
     * 文字サイズであふれを調整する（両方ONなら「最大まで拡大」→「あふれを解消」の順）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @param {boolean} doMaximize - ［文字サイズ：最大まで拡大］を実行するか
     * @param {boolean} doShrink - ［文字サイズ：あふれを解消］を実行するか
     * @returns {void}
     */
    function runFontSizeAdjust(doc, textFrames, doMaximize, doShrink) {
        if (textFrames.length === 0) {
            alert(getLabel(LABELS.alert.noValidText));
            return;
        }

        resetToOriginalValue(doc, textFrames, FONT_SIZE_TAG_NAME, getFontSize, setFontSize);

        if (doMaximize && !adjustEachFrame(textFrames, maximizeFontToFit)) return;
        if (doShrink && !adjustEachFrame(textFrames, shrinkFontToFit)) return;

        removeTagsAfterLastDataSet(doc, textFrames, FONT_SIZE_TAG_NAME);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * チェックボックスかラジオボタンを、LABELS のキーで tooltip 付きで追加する
     * @param {Panel|Group} parentContainer - 追加先のパネルかグループ
     * @param {string} controlType - "checkbox" / "radiobutton"
     * @param {string} labelKey - LABELS.checkbox（または LABELS.radio）と LABELS.tooltip のキー
     * @returns {Checkbox|RadioButton} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var labelCategory = (controlType === "radiobutton") ? LABELS.radio : LABELS.checkbox;
        var addedControl = parentContainer.add(controlType, undefined, getLabel(labelCategory[labelKey]));
        addedControl.helpTip = getLabel(LABELS.tooltip[labelKey]);
        return addedControl;
    }

    /**
     * ［調整方法］パネルを組み立て、中のコントロールを返す
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {boolean} hasAreaText - 選択にエリア内文字があるか
     * @returns {Object} パネル内のコントロール
     */
    function addProcessingPanel(parentWindow, hasAreaText) {
        var processingPanel = addPanel(parentWindow, LABELS.panel.processing);

        var shrinkCheckbox = addLabeledControl(processingPanel, "checkbox", "shrinkToFit");
        shrinkCheckbox.value = DEFAULT_SHRINK_TO_FIT;

        var maximizeCheckbox = addLabeledControl(processingPanel, "checkbox", "maximizeSize");
        maximizeCheckbox.value = DEFAULT_MAXIMIZE_SIZE;

        var heightModeCheckbox = addLabeledControl(processingPanel, "checkbox", "heightMode");
        heightModeCheckbox.enabled = hasAreaText;
        heightModeCheckbox.value = (hasAreaText && DEFAULT_HEIGHT_MODE);

        var heightOptionGroup = addIndentedColumn(processingPanel);
        var adjustHeightRadio = addLabeledControl(heightOptionGroup, "radiobutton", "adjustHeight");
        var autoSizeRadio = addLabeledControl(heightOptionGroup, "radiobutton", "autoSize");

        /* 高さ調整をOFFに戻したとき用に、文字サイズの選択を控える / Remember the font-size choices */
        var previousShrinkState = shrinkCheckbox.value;
        var previousMaximizeState = maximizeCheckbox.value;

        /**
         * 高さ調整のON/OFFに合わせて、各項目の有効・無効と選択状態を切り替える
         * @returns {void}
         */
        function updateHeightOptionState() {
            var isHeightMode = (heightModeCheckbox.value && hasAreaText);
            heightOptionGroup.enabled = isHeightMode;

            if (isHeightMode) {
                /* 高さ調整中は文字サイズの処理を止める / The font-size options are off while adjusting the height */
                previousShrinkState = shrinkCheckbox.value;
                previousMaximizeState = maximizeCheckbox.value;
                shrinkCheckbox.value = false;
                maximizeCheckbox.value = false;
                shrinkCheckbox.enabled = false;
                maximizeCheckbox.enabled = false;
                autoSizeRadio.value = (DEFAULT_HEIGHT_OPTION === "autoSize");
                adjustHeightRadio.value = !autoSizeRadio.value;
            } else {
                shrinkCheckbox.enabled = true;
                maximizeCheckbox.enabled = true;
                shrinkCheckbox.value = previousShrinkState;
                maximizeCheckbox.value = previousMaximizeState;
                adjustHeightRadio.value = false;
                autoSizeRadio.value = false;
            }
        }

        heightModeCheckbox.onClick = updateHeightOptionState;
        updateHeightOptionState();

        return {
            shrinkCheckbox: shrinkCheckbox,
            maximizeCheckbox: maximizeCheckbox,
            heightModeCheckbox: heightModeCheckbox,
            autoSizeRadio: autoSizeRadio
        };
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
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

    /**
     * パネルの選択内容を設定にまとめる（選択が足りないときは警告する）
     * @param {Object} processingControls - ［調整方法］パネルのコントロール
     * @returns {Object|null} 実行する処理の設定（選択が足りないときは null）
     */
    function readAdjustSettings(processingControls) {
        if (processingControls.heightModeCheckbox.value) {
            return { mode: processingControls.autoSizeRadio.value ? ADJUST_MODE.AUTO_SIZE : ADJUST_MODE.HEIGHT };
        }

        if (!processingControls.shrinkCheckbox.value && !processingControls.maximizeCheckbox.value) {
            alert(getLabel(LABELS.alert.selectMode));
            return null;
        }

        return {
            mode: ADJUST_MODE.FONT_SIZE,
            doMaximize: processingControls.maximizeCheckbox.value,
            doShrink: processingControls.shrinkCheckbox.value
        };
    }

    /**
     * ダイアログを表示し、選ばれた処理を返す
     * @param {boolean} hasAreaText - 選択にエリア内文字があるか
     * @returns {Object|null} 実行する処理の設定（キャンセル時は null）
     */
    function showDialog(hasAreaText) {
        var adjustDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(adjustDialog);

        var processingControls = addProcessingPanel(adjustDialog, hasAreaText);
        var buttonRow = addButtonRow(adjustDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        var selectedSettings = null;

        btnOK.onClick = function() {
            selectedSettings = readAdjustSettings(processingControls);
            if (selectedSettings) adjustDialog.close(1);
        };

        btnCancel.onClick = function() {
            adjustDialog.close(0);
        };

        prepareDialogWindow(adjustDialog, SCRIPT_NAME);
        adjustDialog.show();
        return selectedSettings;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 選択からテキストを集め、ダイアログで選ばれた処理を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.selectObject));
            return;
        }

        /* 処理中に選択が変わるため、ダイアログの前に対象を確定させる / Collect the targets before the selection changes */
        var textFrames = getSelectedTextFrames(doc);
        var hasAreaText = (filterAreaTextFrames(textFrames).length > 0);

        var adjustSettings = showDialog(hasAreaText);
        if (!adjustSettings) return;

        if (adjustSettings.mode === ADJUST_MODE.AUTO_SIZE) {
            runAutoSize(textFrames);
        } else if (adjustSettings.mode === ADJUST_MODE.HEIGHT) {
            runHeightAdjust(doc, textFrames);
        } else {
            runFontSizeAdjust(doc, textFrames, adjustSettings.doMaximize, adjustSettings.doShrink);
        }
    }

    main();

})();
