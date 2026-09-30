#target illustrator
#targetengine "AiBreakLinkEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したシンボルインスタンスのリンクを解除し、通常のオブジェクトとして扱える状態にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiBreakLink.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf729c53f4300

### Overview

Breaks the links of the selected symbol instances and turns them into regular editable objects.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiBreakLink.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiBreakLink";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiBreakLink.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiBreakLink.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf729c53f4300"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 起動時にオプションダイアログを表示する / Show the options dialog on launch */
    var SHOW_OPTIONS_DIALOG = true;

    /* 1アイテムだけのグループを自動的に解除する / Auto-ungroup one-item groups */
    var UNGROUP_SINGLE_ITEM_GROUP_DEFAULT = true;

    /* 解除結果グループを完全に（ネストごと）解除する / Fully ungroup the result group, including nested groups */
    var UNGROUP_ALL_DEFAULT = false;

    /* 解除後の名前に元のシンボル名を使う / Reuse the original symbol name */
    var INHERIT_SYMBOL_NAME_DEFAULT = true;

    /* 解除後の名前に接頭辞を付ける / Add a prefix to the result name */
    var USE_PREFIX_DEFAULT = true;

    /* 解除後の名前に付ける接頭辞 / Prefix for the result name */
    var UNLINKED_ITEM_NAME_PREFIX_DEFAULT = "symbol_unlinked_";

    /* 単一テキストの場合は内容を名前にする / Use text content as name for a single text */
    var USE_TEXT_CONTENT_AS_NAME_DEFAULT = true;

    /* ネストしたグループを解除する際の反復回数の上限 / Cap on ungroup iterations for nested groups */
    var MAX_UNGROUP_ITERATIONS = 50;

    // =========================================
    // ローカライズ / Localization
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "シンボルのリンクを解除", en: "Break Symbol Links" }
        },
        panel: {
            ungroupOptions: { ja: "オプション", en: "Options" },
            resultName: { ja: "解除後の名前", en: "Result Name" }
        },
        checkbox: {
            ungroupSingleItem: { ja: "1アイテムだけのグループを解除", en: "Ungroup one-item groups" },
            ungroupAll: { ja: "グループを完全に解除", en: "Ungroup completely" },
            inheritSymbolName: { ja: "元のシンボル名を使う", en: "Use original symbol name" },
            prefix: { ja: "名前の接頭辞", en: "Name prefix" },
            useTextContentAsName: { ja: "単一テキストは内容を名前にする", en: "Use single text content as name" }
        },
        tooltip: {
            ungroupSingleItem: {
                ja: "解除後に中身が1つだけのグループができた場合、そのグループを外して単体オブジェクトにします。",
                en: "If breaking produces a group containing only one item, remove that group and keep the item standalone."
            },
            ungroupAll: {
                ja: "解除結果のグループをネストも含めて完全に解除します。グループに付けた名前は失われます。",
                en: "Fully ungroup the result, including nested groups. Any name given to a group is lost."
            },
            inheritSymbolName: {
                ja: "解除後のオブジェクト名に、元のシンボル名を引き継ぎます。",
                en: "Carry over the original symbol name to the unlinked object."
            },
            prefix: {
                ja: "シンボル名の前に付ける文字列。例: symbol_unlinked_ボタン",
                en: "Text added before the symbol name, e.g. symbol_unlinked_Button."
            },
            useTextContentAsName: {
                ja: "解除結果が1つのテキストのとき、その内容をオブジェクト名にします（シンボル名より優先）。",
                en: "When the result is a single text object, use its content as the name (takes priority over the symbol name)."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        log: {
            moveStaticSubLayerItem: { ja: "static 解除結果のサブレイヤーアイテムを変換後グループへ移動", en: "Move static sublayer item into converted group" },
            removeConvertedSubLayer: { ja: "変換済みサブレイヤーの削除", en: "Remove converted sublayer" },
            moveSingleItemBeforeConvertedGroup: { ja: "単一アイテムを変換後グループの前へ移動", en: "Move single item before converted group" },
            removeSingleItemConvertedGroup: { ja: "単一アイテム化した変換後グループの削除", en: "Remove single-item converted group" },
            moveStaticBreakItemIntoResultGroup: { ja: "static 解除結果アイテムを結果グループへ移動", en: "Move static break item into result group" },
            selectDynamicBreakItem: { ja: "dynamic 解除結果アイテムを選択", en: "Select dynamic break item" },
            selectTargetSymbolItem: { ja: "解除対象のシンボルを選択", en: "Select target symbol item" },
            selectGeneratedResultItem: { ja: "生成された解除結果アイテムを選択", en: "Select generated break-result item" },
            ungroupAllResultGroup: { ja: "解除結果グループを完全に解除", en: "Fully ungroup the result group" },
            ungroupNestedDynamicBreakGroup: { ja: "dynamic 解除結果のネストグループを解除", en: "Ungroup nested dynamic break group" },
            moveDynamicBreakItemIntoResultGroup: { ja: "dynamic 解除結果アイテムを結果グループへ移動", en: "Move dynamic break item into result group" },
            detectBreakLinkResult: { ja: "breakLink 結果の判定", en: "Detect breakLink result" },
            mixedBreakLinkResult: {
                ja: "新規サブレイヤーと新規ページアイテムの両方が見つかりました。static として処理します。",
                en: "Both generated sublayers and page items were found. Treating the result as static."
            },
            noBreakLinkResult: {
                ja: "新規サブレイヤーも新規ページアイテムも見つかりませんでした。",
                en: "No generated sublayers or page items were found."
            },
            processSymbolItem: { ja: "シンボルのリンク解除・正規化・命名", en: "Break, normalize, and name symbol" }
        }
    };

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */

    /**
     * ウィンドウに共通のレイアウトを適用します。
     *
     * @param {object} win - 対象の Window。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用します。
     *
     * @param {object} panel - 対象の Panel。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びのグループ（ボタン列など）に共通のレイアウトを適用します。
     *
     * @param {object} group - 対象の Group。
     * @param {string} [alignment] - グループ自体の配置。省略時は "left"。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    (function () {

        if (app.documents.length === 0) {
            return;
        }

        var documentRef = app.activeDocument;

        if (documentRef.selection.length === 0) {
            return;
        }

        var ungroupSingleItemGroup = UNGROUP_SINGLE_ITEM_GROUP_DEFAULT;
        var ungroupAll = UNGROUP_ALL_DEFAULT;
        var inheritSymbolName = INHERIT_SYMBOL_NAME_DEFAULT;
        var usePrefix = USE_PREFIX_DEFAULT;
        var unlinkedItemNamePrefix = UNLINKED_ITEM_NAME_PREFIX_DEFAULT;
        var useTextContentAsName = USE_TEXT_CONTENT_AS_NAME_DEFAULT;

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        /**
         * オプションダイアログを表示し、選択された設定を返します。
         *
         * @returns {object} 設定オブジェクト。キャンセルされた場合は null。
         */
        function showOptionsDialog() {
            var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
            setupWindow(dialog);

            var ungroupPanel = dialog.add("panel", undefined, getLabel("panel.ungroupOptions"));
            setupPanel(ungroupPanel, 6);

            var ungroupSingleItemCheckbox = ungroupPanel.add("checkbox", undefined, getLabel("checkbox.ungroupSingleItem"));
            ungroupSingleItemCheckbox.value = UNGROUP_SINGLE_ITEM_GROUP_DEFAULT;
            ungroupSingleItemCheckbox.helpTip = getLabel("tooltip.ungroupSingleItem");

            var ungroupAllCheckbox = ungroupPanel.add("checkbox", undefined, getLabel("checkbox.ungroupAll"));
            ungroupAllCheckbox.value = UNGROUP_ALL_DEFAULT;
            ungroupAllCheckbox.helpTip = getLabel("tooltip.ungroupAll");

            var resultNamePanel = dialog.add("panel", undefined, getLabel("panel.resultName"));
            setupPanel(resultNamePanel, 6);

            var inheritNameCheckbox = resultNamePanel.add("checkbox", undefined, getLabel("checkbox.inheritSymbolName"));
            inheritNameCheckbox.value = INHERIT_SYMBOL_NAME_DEFAULT;
            inheritNameCheckbox.helpTip = getLabel("tooltip.inheritSymbolName");

            var prefixRow = resultNamePanel.add("group");
            setupRow(prefixRow, "left", 6);

            var prefixCheckbox = prefixRow.add("checkbox", undefined, labelText("checkbox.prefix"));
            prefixCheckbox.value = USE_PREFIX_DEFAULT;
            prefixCheckbox.helpTip = getLabel("tooltip.prefix");

            var prefixInput = prefixRow.add("edittext", undefined, UNLINKED_ITEM_NAME_PREFIX_DEFAULT);
            prefixInput.characters = 20;
            prefixInput.helpTip = getLabel("tooltip.prefix");

            var useTextNameCheckbox = resultNamePanel.add("checkbox", undefined, getLabel("checkbox.useTextContentAsName"));
            useTextNameCheckbox.value = USE_TEXT_CONTENT_AS_NAME_DEFAULT;
            useTextNameCheckbox.helpTip = getLabel("tooltip.useTextContentAsName");

            /**
             * チェックボックスの状態に応じて、各コントロールの有効・無効を切り替えます。
             *
             * @returns {void}
             */
            function syncEnabledStates() {
                var isFullUngroup = ungroupAllCheckbox.value;
                /* 完全解除が ON なら 1アイテム解除は無意味なので無効化 / When full ungroup is on, the single-item option is moot, so disable it */
                ungroupSingleItemCheckbox.enabled = !isFullUngroup;
                /* 完全解除が ON ならコンテナに付けた名前は破棄されるため、シンボル名系の命名を無効化 / When full ungroup is on, container names are discarded, so disable the symbol-name options */
                inheritNameCheckbox.enabled = !isFullUngroup;
                prefixCheckbox.enabled = !isFullUngroup && inheritNameCheckbox.value;
                prefixInput.enabled = prefixCheckbox.enabled && prefixCheckbox.value;
            }
            syncEnabledStates();
            ungroupSingleItemCheckbox.onClick = syncEnabledStates;
            ungroupAllCheckbox.onClick = syncEnabledStates;
            inheritNameCheckbox.onClick = syncEnabledStates;
            prefixCheckbox.onClick = syncEnabledStates;

            var buttonRow = addButtonRow(dialog);
            var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

            /* 表示直後にレイアウトを再計算して描画欠けを防ぐ / Recalculate layout on show to avoid partial rendering */
            dialog.onShow = function () {
                dialog.layout.layout(true);
                dialog.layout.resize();
            };

            alignRightOnlyButtonRow(buttonRow);
            prepareDialogWindow(dialog, SCRIPT_NAME);
            if (dialog.show() !== 1) return null;

            return {
                ungroupSingleItem: ungroupSingleItemCheckbox.value,
                ungroupAll: ungroupAllCheckbox.value,
                inheritSymbolName: inheritNameCheckbox.value,
                usePrefix: prefixCheckbox.value,
                unlinkedPrefix: prefixInput.text,
                useTextContentAsName: useTextNameCheckbox.value
            };
        }

        // =========================================
        // 安全な実行ヘルパー / Safe execution helpers
        // =========================================

        /**
         * コンテキスト付きでエラーを $.writeln に出力します。
         *
         * @param {string} context - どの処理で起きたかを示す文言。
         * @param {object} errorObject - エラーオブジェクトまたはメッセージ。
         * @returns {void}
         */
        function logScriptError(context, errorObject) {
            $.writeln("[" + SCRIPT_NAME + " " + SCRIPT_VERSION + "] " + context + ": " + errorObject);
        }

        /**
         * 処理を実行し、失敗した場合はログを出して続行します。
         * ロック・非表示・削除済みなど、Illustrator 側の状態で失敗しうる操作に使います。
         *
         * @param {function} operation - 実行する処理。
         * @param {string} context - 失敗時にログへ出す文言。
         * @returns {boolean} 成功したら true、失敗したら false。
         */
        function runSafely(operation, context) {
            try {
                operation();
                return true;
            } catch (e) {
                logScriptError(context, e);
                return false;
            }
        }

        /**
         * アイテムを削除します（失敗しても続行）。
         *
         * @param {object} item - 対象のアイテムまたはレイヤー。
         * @param {string} context - 失敗時にログへ出す文言。
         * @returns {boolean} 成功したら true。
         */
        function removeItemSafely(item, context) {
            return runSafely(function () {
                item.remove();
            }, context);
        }

        /**
         * アイテムを移動します（失敗しても続行）。
         *
         * @param {object} item - 対象のアイテム。
         * @param {object} destination - 移動先。
         * @param {object} placement - ElementPlacement の値。
         * @param {string} context - 失敗時にログへ出す文言。
         * @returns {boolean} 成功したら true。
         */
        function moveItemSafely(item, destination, placement, context) {
            return runSafely(function () {
                item.move(destination, placement);
            }, context);
        }

        /**
         * アイテムの選択状態を変更します（失敗しても続行）。
         *
         * @param {object} item - 対象のアイテム。
         * @param {boolean} isSelected - 選択するなら true。
         * @param {string} context - 失敗時にログへ出す文言。
         * @returns {boolean} 成功したら true。
         */
        function setItemSelectedSafely(item, isSelected, context) {
            return runSafely(function () {
                item.selected = isSelected;
            }, context);
        }

        /**
         * メニューコマンドを実行します（失敗しても続行）。
         *
         * @param {string} commandName - メニューコマンド名。
         * @param {string} context - 失敗時にログへ出す文言。
         * @returns {boolean} 成功したら true。
         */
        function executeMenuCommandSafely(commandName, context) {
            return runSafely(function () {
                app.executeMenuCommand(commandName);
            }, context);
        }

        // =========================================
        // 汎用ユーティリティ / Generic utilities
        // =========================================

        /**
         * ExtendScript のコレクションを通常の配列へコピーします。
         *
         * @param {object} collection - pageItems や selection などのコレクション。
         * @returns {array} コピーした配列。
         */
        function collectionToArray(collection) {
            var items = [];
            for (var i = 0; i < collection.length; i++) {
                items.push(collection[i]);
            }
            return items;
        }

        /**
         * breakLink 前後の差分から、新しく生成されたアイテムだけを抽出します。
         *
         * @param {object} currentItems - 現在のコレクションまたは配列。
         * @param {array} existingItems - 処理前に控えておいた配列。
         * @returns {array} 新しく生成されたアイテムの配列。
         */
        function collectGeneratedItems(currentItems, existingItems) {
            var generatedItems = [];
            for (var i = 0; i < currentItems.length; i++) {
                var currentItem = currentItems[i];
                var isExisting = false;
                for (var j = 0; j < existingItems.length; j++) {
                    if (existingItems[j] === currentItem) {
                        isExisting = true;
                        break;
                    }
                }
                if (!isExisting) generatedItems.push(currentItem);
            }
            return generatedItems;
        }

        /**
         * 指定アイテムを含む Layer を返します。
         *
         * @param {object} pageItem - 対象のアイテム。
         * @returns {object} 含まれる Layer。見つからない場合は null。
         */
        function getContainingLayer(pageItem) {
            var currentParent = pageItem.parent;
            while (currentParent && currentParent.typename !== "Layer") {
                currentParent = currentParent.parent;
            }
            return currentParent;
        }

        /**
         * 選択範囲に GroupItem が含まれるかを判定します。
         *
         * @param {object} selectedItems - selection または配列。
         * @returns {boolean} 含まれていれば true。
         */
        function containsGroupItem(selectedItems) {
            for (var i = 0; i < selectedItems.length; i++) {
                if (selectedItems[i].typename === "GroupItem") return true;
            }
            return false;
        }

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

        // =========================================
        // シンボルの収集 / Symbol collection
        // =========================================

        /**
         * 選択範囲全体から SymbolItem をまとめて集めます（グループ内も対象）。
         *
         * @param {object} selectedItems - selection。
         * @returns {array} SymbolItem の配列。
         */
        function collectSymbolItemsFromSelection(selectedItems) {
            return collectSelectionItems(selectedItems, {
                accept: function (item) {
                    return item.typename === "SymbolItem";
                }
            });
        }

        // =========================================
        // 解除結果の整理 / Break-result normalization
        // =========================================

        /**
         * 指定サブレイヤーを親 Layer 直下のグループへ変換します。
         * 中身が1アイテムだけのグループは、設定に応じて解除します。
         *
         * @param {object} parentLayer - 変換先の親 Layer。
         * @param {object} generatedSubLayer - breakLink で生成されたサブレイヤー。
         * @returns {object} 変換後のグループまたは単体アイテム。変換できない場合は null。
         */
        function convertSubLayerToGroup(parentLayer, generatedSubLayer) {
            if (!parentLayer || !generatedSubLayer) return null;

            flattenSubLayersToGroups(generatedSubLayer);

            var convertedGroup = parentLayer.groupItems.add();
            for (var i = generatedSubLayer.pageItems.length - 1; i >= 0; i--) {
                moveItemSafely(generatedSubLayer.pageItems[i], convertedGroup, ElementPlacement.PLACEATBEGINNING, getLabel("log.moveStaticSubLayerItem"));
            }

            removeItemSafely(generatedSubLayer, getLabel("log.removeConvertedSubLayer"));

            if (ungroupSingleItemGroup && convertedGroup.pageItems.length === 1) {
                var singleItem = convertedGroup.pageItems[0];
                moveItemSafely(singleItem, convertedGroup, ElementPlacement.PLACEBEFORE, getLabel("log.moveSingleItemBeforeConvertedGroup"));
                removeItemSafely(convertedGroup, getLabel("log.removeSingleItemConvertedGroup"));
                return singleItem;
            }

            return convertedGroup;
        }

        /**
         * 指定 Layer 配下のサブレイヤーを、再帰的にグループへ変換してフラット化します。
         *
         * @param {object} parentLayer - 対象の Layer。
         * @returns {void}
         */
        function flattenSubLayersToGroups(parentLayer) {
            for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
                convertSubLayerToGroup(parentLayer, parentLayer.layers[i]);
            }
        }

        /**
         * 複数のアイテムはグループにまとめ、1つだけならそのまま返します。
         *
         * @param {object} targetLayer - グループを作成する Layer。
         * @param {array} items - 対象のアイテム配列。
         * @param {string} logLabelPath - 移動失敗時にログへ出すラベルのドットパス。
         * @returns {object} 作成したグループまたは単体アイテム。アイテムがなければ null。
         */
        function groupItemsOrReturnSingle(targetLayer, items, logLabelPath) {
            if (items.length === 0) return null;
            if (items.length === 1) return items[0];

            var resultGroup = targetLayer.groupItems.add();
            for (var i = 0; i < items.length; i++) {
                moveItemSafely(items[i], resultGroup, ElementPlacement.PLACEATEND, getLabel(logLabelPath));
            }
            return resultGroup;
        }

        /**
         * static シンボル（サブレイヤー生成型）の解除結果を、グループまたは単体に整理します。
         *
         * @param {object} targetLayer - 解除対象があった Layer。
         * @param {array} itemsBeforeBreak - breakLink 前の Layer 直下のアイテム配列。
         * @param {array} generatedSubLayers - breakLink で生成されたサブレイヤーの配列。
         * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
         */
        function organizeStaticBreakResult(targetLayer, itemsBeforeBreak, generatedSubLayers) {
            if (!targetLayer) return null;

            for (var i = 0; i < generatedSubLayers.length; i++) {
                convertSubLayerToGroup(targetLayer, generatedSubLayers[i]);
            }

            var generatedItems = collectGeneratedItems(targetLayer.pageItems, itemsBeforeBreak);
            return groupItemsOrReturnSingle(targetLayer, generatedItems, "log.moveStaticBreakItemIntoResultGroup");
        }

        /**
         * dynamic シンボル（ページアイテム生成型）の解除結果を、
         * ungroup で平坦化したうえでグループまたは単体に整理します。
         *
         * @param {object} documentObject - 対象ドキュメント。
         * @param {object} targetLayer - 解除対象があった Layer。
         * @param {array} itemsBeforeBreak - breakLink 前の Layer 直下のアイテム配列。
         * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
         */
        function organizeDynamicBreakResult(documentObject, targetLayer, itemsBeforeBreak) {
            if (!targetLayer) return null;

            /* breakLink 直後の selection に依存せず、Layer 配下の差分で新規アイテムを特定 / Detect generated items by Layer reference differences instead of relying on the selection after breakLink */
            var generatedItems = collectGeneratedItems(targetLayer.pageItems, itemsBeforeBreak);
            if (generatedItems.length === 0) return null;

            documentObject.selection = null;
            for (var i = 0; i < generatedItems.length; i++) {
                setItemSelectedSafely(generatedItems[i], true, getLabel("log.selectDynamicBreakItem"));
            }

            /* ネストが深い場合に備え、ungroup の反復回数に上限を設ける / Cap the ungroup iterations to guard against deeply nested groups */
            var ungroupCount = 0;
            while (containsGroupItem(documentObject.selection) && ungroupCount < MAX_UNGROUP_ITERATIONS) {
                if (!executeMenuCommandSafely("ungroup", getLabel("log.ungroupNestedDynamicBreakGroup"))) break;
                ungroupCount++;
            }

            /* selection は移動中に変化するため、事前に配列へ退避 / Snapshot the selection into an array since it changes during moves */
            var brokenItems = collectionToArray(documentObject.selection);
            if (brokenItems.length === 0) return null;

            var resultLayer = getContainingLayer(brokenItems[0]) || targetLayer;
            return groupItemsOrReturnSingle(resultLayer, brokenItems, "log.moveDynamicBreakItemIntoResultGroup");
        }

        /**
         * 生成物の種類から、解除結果の型を判定します。
         *
         * @param {array} generatedSubLayers - 生成されたサブレイヤーの配列。
         * @param {array} generatedPageItems - 生成されたページアイテムの配列。
         * @returns {string} "static" / "dynamic" / "mixed" / "none" のいずれか。
         */
        function classifyBreakResult(generatedSubLayers, generatedPageItems) {
            var hasSubLayers = generatedSubLayers.length > 0;
            var hasPageItems = generatedPageItems.length > 0;

            if (hasSubLayers && !hasPageItems) return "static";
            if (!hasSubLayers && hasPageItems) return "dynamic";
            if (hasSubLayers && hasPageItems) return "mixed";
            return "none";
        }

        /**
         * 判定結果に応じて、static / dynamic の整理処理へ振り分けます。
         *
         * @param {object} documentObject - 対象ドキュメント。
         * @param {object} targetLayer - 解除対象があった Layer。
         * @param {array} itemsBeforeBreak - breakLink 前の Layer 直下のアイテム配列。
         * @param {array} generatedSubLayers - 生成されたサブレイヤーの配列。
         * @param {array} generatedPageItems - 生成されたページアイテムの配列。
         * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
         */
        function normalizeBreakResult(documentObject, targetLayer, itemsBeforeBreak, generatedSubLayers, generatedPageItems) {
            var breakResultType = classifyBreakResult(generatedSubLayers, generatedPageItems);

            if (breakResultType === "dynamic") {
                return organizeDynamicBreakResult(documentObject, targetLayer, itemsBeforeBreak);
            }

            if (breakResultType === "mixed") {
                logScriptError(getLabel("log.detectBreakLinkResult"), getLabel("log.mixedBreakLinkResult"));
            }

            if (breakResultType === "none") {
                logScriptError(getLabel("log.detectBreakLinkResult"), getLabel("log.noBreakLinkResult"));
                return null;
            }

            return organizeStaticBreakResult(targetLayer, itemsBeforeBreak, generatedSubLayers);
        }

        // =========================================
        // 解除後の命名 / Result naming
        // =========================================

        /**
         * 設定に応じて、接頭辞を付けた解除後の名前を組み立てます。
         *
         * @param {string} symbolName - 元のシンボル名。
         * @returns {string} 解除後のアイテム名。
         */
        function buildUnlinkedItemName(symbolName) {
            return (usePrefix ? unlinkedItemNamePrefix : "") + symbolName;
        }

        /**
         * 単一 TextFrame の文字列を取得します（改行は空白に置換）。
         *
         * @param {object} item - 対象のアイテム。
         * @returns {string} テキストの内容。対象外または空文字の場合は null。
         */
        function getSingleTextFrameContent(item) {
            if (!item) return null;

            var textFrame = null;
            if (item.typename === "TextFrame") {
                textFrame = item;
            } else if (item.typename === "GroupItem" && item.pageItems.length === 1 && item.pageItems[0].typename === "TextFrame") {
                textFrame = item.pageItems[0];
            }
            if (!textFrame) return null;

            var contents = textFrame.contents;
            if (contents === null || typeof contents === "undefined") return null;

            contents = contents.replace(/[\r\n]+/g, " ");
            return contents.length === 0 ? null : contents;
        }

        /**
         * 解除結果アイテムに名前を付けます（単一テキストの内容を元シンボル名より優先）。
         *
         * @param {object} resultItem - 解除結果のアイテム。
         * @param {string} symbolName - 元のシンボル名。
         * @returns {void}
         */
        function applyResultName(resultItem, symbolName) {
            if (!resultItem) return;

            if (useTextContentAsName) {
                var textContent = getSingleTextFrameContent(resultItem);
                if (textContent !== null) {
                    resultItem.name = textContent;
                    return;
                }
            }

            if (inheritSymbolName) {
                resultItem.name = buildUnlinkedItemName(symbolName);
            }
        }

        // =========================================
        // シンボル1つ分の処理 / Per-symbol processing
        // =========================================

        /**
         * breakLink 前の Layer 直下のアイテムとサブレイヤーを控えておきます。
         *
         * @param {object} targetLayer - 対象の Layer。
         * @returns {object} pageItems / subLayers を持つオブジェクト。
         */
        function snapshotLayerContents(targetLayer) {
            return {
                pageItems: targetLayer ? collectionToArray(targetLayer.pageItems) : [],
                subLayers: targetLayer ? collectionToArray(targetLayer.layers) : []
            };
        }

        /**
         * シンボルのリンクを解除し、生成物を整理して1つの結果アイテムにまとめます。
         *
         * @param {object} symbolItem - 対象の SymbolItem。
         * @param {object} documentObject - 対象ドキュメント。
         * @returns {object} 整理後のグループまたは単体アイテム。該当がなければ null。
         */
        function breakAndNormalizeSymbol(symbolItem, documentObject) {
            var targetLayer = getContainingLayer(symbolItem);
            var beforeBreak = snapshotLayerContents(targetLayer);

            symbolItem.breakLink();

            var generatedSubLayers = targetLayer ? collectGeneratedItems(targetLayer.layers, beforeBreak.subLayers) : [];
            var generatedPageItems = targetLayer ? collectGeneratedItems(targetLayer.pageItems, beforeBreak.pageItems) : [];

            return normalizeBreakResult(documentObject, targetLayer, beforeBreak.pageItems, generatedSubLayers, generatedPageItems);
        }

        /**
         * 解除結果グループをネストごと完全に解除し、展開後のアイテム配列を返します。
         *
         * @param {object} documentObject - 対象ドキュメント。
         * @param {object} resultItem - 解除結果のアイテム。
         * @returns {array} 展開後のアイテム配列。
         */
        function ungroupResultCompletely(documentObject, resultItem) {
            if (!resultItem) return [];
            if (resultItem.typename !== "GroupItem") return [resultItem];

            documentObject.selection = null;
            if (!setItemSelectedSafely(resultItem, true, getLabel("log.ungroupAllResultGroup"))) return [resultItem];

            /* ungroupAll はネストごと一括で解除する / ungroupAll dissolves the group and all nested groups at once */
            executeMenuCommandSafely("ungroupAll", getLabel("log.ungroupAllResultGroup"));

            /* 解除直後の選択が展開後アイテム / The selection right after ungroupAll holds the loose items */
            return collectionToArray(documentObject.selection);
        }

        /**
         * SymbolItem 1つ分の処理（解除・整理・命名・完全解除）をまとめます。
         *
         * @param {object} symbolItem - 対象の SymbolItem。
         * @param {object} documentObject - 対象ドキュメント。
         * @returns {array} 選択状態として残す結果アイテムの配列。
         */
        function processSymbolItem(symbolItem, documentObject) {
            if (symbolItem.locked || symbolItem.hidden) return [];

            /* 外側で初期選択を解除済みの前提で、この SymbolItem だけを処理対象として選択 / Select only this SymbolItem, assuming the outer flow already cleared the initial selection */
            if (!setItemSelectedSafely(symbolItem, true, getLabel("log.selectTargetSymbolItem"))) return [];

            var symbolName = symbolItem.symbol.name;
            var resultItem = breakAndNormalizeSymbol(symbolItem, documentObject);

            applyResultName(resultItem, symbolName);

            /* 完全解除が ON なら、命名後に結果グループをネストごと解除して展開後アイテムを返す / When full ungroup is on, dissolve the result group after naming and return the loose items */
            if (ungroupAll) return ungroupResultCompletely(documentObject, resultItem);

            return resultItem ? [resultItem] : [];
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        var symbolItems = collectSymbolItemsFromSelection(documentRef.selection);

        if (symbolItems.length === 0) {
            return;
        }

        if (SHOW_OPTIONS_DIALOG) {
            var dialogSettings = showOptionsDialog();
            if (!dialogSettings) {
                return;
            }
            ungroupSingleItemGroup = dialogSettings.ungroupSingleItem;
            ungroupAll = dialogSettings.ungroupAll;
            inheritSymbolName = dialogSettings.inheritSymbolName;
            usePrefix = dialogSettings.usePrefix;
            unlinkedItemNamePrefix = dialogSettings.unlinkedPrefix;
            useTextContentAsName = dialogSettings.useTextContentAsName;
        }

        /* 処理後の選択方針 / Post-processing selection policy */
        /* 元の選択は復元せず、解除後に生成されたアイテムを選択状態として残す / Do not restore the original selection; keep the generated unlinked items selected */
        /* 初期選択はここで一度だけ解除し、生成物を配列に集めて末尾でまとめて選択する / Clear the initial selection once here, collect generated items, and select them all at the end */
        documentRef.selection = null;

        var generatedResultItems = [];
        for (var i = 0; i < symbolItems.length; i++) {
            /* 1つのシンボルで失敗しても、残りのシンボルの処理は続行する / Keep processing the remaining symbols even if one of them fails */
            try {
                var producedItems = processSymbolItem(symbolItems[i], documentRef);
                for (var j = 0; j < producedItems.length; j++) {
                    generatedResultItems.push(producedItems[j]);
                }
            } catch (e) {
                logScriptError(getLabel("log.processSymbolItem"), e);
            }
        }

        /* 全シンボルの生成物を最終的に選択状態へ（static / dynamic / 完全解除で共通）/ Select every symbol's generated items at the end (uniform for static / dynamic / full-ungroup) */
        documentRef.selection = null;
        for (var k = 0; k < generatedResultItems.length; k++) {
            setItemSelectedSafely(generatedResultItems[k], true, getLabel("log.selectGeneratedResultItem"));
        }

    })();

})();
