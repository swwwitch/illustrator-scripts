#target illustrator
#targetengine "ReleaseEverythingEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトのグループ・複合パス・複合シェイプ・ブレンド・エンベロープ・リピートなど、解除できるものを入れ子までまとめて解除します。
クリップグループは単純に解除してマスクパスに塗りを付け、入れ子のグループがあるときだけダイアログで解除する深さを選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseEverything.md

### Overview

Releases everything that can be released in the selection, such as groups, compound paths, compound shapes, blends, envelopes, and repeats, including nested ones.
Clipping groups are released simply with a fill on the mask path; a dialog asks how deep to release only when there are nested groups.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseEverything.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReleaseEverything";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseEverything.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseEverything.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* クリップグループの解除方法（ダイアログは出さない）と、入れ子のグループのダイアログの初期値
       How clipping groups are released (no dialog) and the default of the nested-group dialog */
    var USER_DEFAULTS = {
        clipReleaseMode: "simple",     /* simple（マスクパスと内容を残す）| removePath（パスを削除）| removeContent（マスク内容を削除） */
        applyMaskFill: true,           /* 残したマスクパスに K100・PATH_FILL_OPACITY の塗りを付ける（removePath では無視） */
        releaseDepth: "all"            /* all | oneLevel */
    };

    /* 残したマスクパスに付ける塗りの不透明度（%）/ Fill opacity applied to the remaining mask path (%) */
    var PATH_FILL_OPACITY = 15;

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    // =========================================
    // ダイアログの位置 / Dialog position
    // =========================================

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

    var LABELS = {
        dialog: {
            title: { ja: "グループとマスクの解除", en: "Release Groups and Masks" }
        },
        panel: {
            releaseDepth: { ja: "入れ子のグループ", en: "Nested Groups" }
        },
        radio: {
            releaseAll: { ja: "すべて解除", en: "Release all levels" },
            releaseOneLevel: { ja: "1階層だけ解除", en: "Release one level only" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            releaseAll: {
                ja: "入れ子になったグループ・複合パス・複合シェイプ・ブレンドなどを、なくなるまで解除します。",
                en: "Release nested groups, compound paths, compound shapes, blends, and more until none remain."
            },
            releaseOneLevel: {
                ja: "選択したグループ・複合パス・複合シェイプ・ブレンドなどを1回だけ解除し、中のものは残します。",
                en: "Release the selected groups, compound paths, compound shapes, blends, and more once and keep what is inside them."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Select one or more objects." },
            releaseFailed: {
                ja: "{count}個のクリッピングマスクの解除に失敗しました。",
                en: "Failed to release {count} clipping mask(s)."
            }
        }
    };

    // =========================================
    // 選択・マスクの判定 / Selection & mask checks
    // =========================================

    /**
     * ドキュメントの選択を配列で取得する。文字の選択（TextRange）などは空配列
     * @param {Document} targetDoc - 対象のドキュメント
     * @returns {PageItem[]} 選択アイテムの配列
     */
    function getSelectionItems(targetDoc) {
        var selection = targetDoc.selection;
        if (!selection || typeof selection.length !== "number" || selection.typename) return [];
        /* ライブ選択を実配列へコピー（以降のDOM変更の影響を受けない）/ Copy the live selection so later DOM changes do not affect it */
        var selectionItems = [];
        for (var i = 0; i < selection.length; i++) {
            selectionItems.push(selection[i]);
        }
        return selectionItems;
    }

    /**
     * 指定したアイテムだけを選択し直す
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem[]} targetItems - 選択するアイテム
     * @returns {void}
     */
    function selectItems(targetDoc, targetItems) {
        targetDoc.selection = null;
        if (targetItems.length) targetDoc.selection = targetItems;
    }

    /**
     * グループを含むグループがあるか調べる（クリップグループも含む）
     * @param {PageItem[]} targetItems - 調べるアイテム
     * @returns {boolean} 1つでもあれば true
     */
    function containsNestedGroup(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            if (targetItems[i].typename === "GroupItem" && targetItems[i].groupItems.length > 0) return true;
        }
        return false;
    }

    /**
     * クリッピングマスクを持つグループかどうか
     * @param {PageItem} pageItem - 判定するアイテム
     * @returns {boolean} クリップグループなら true
     */
    function isClippingGroup(pageItem) {
        return pageItem.typename === "GroupItem" && pageItem.clipped === true;
    }

    /**
     * アイテムがクリッピングマスクとして働いているか。複合パスは中のパスの clipping を見る
     * @param {PageItem} pageItem - 判定するアイテム
     * @returns {boolean} マスクなら true
     */
    function isMaskItem(pageItem) {
        if (pageItem.typename === "PathItem") return pageItem.clipping === true;
        if (pageItem.typename === "CompoundPathItem") {
            var subPaths = pageItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                if (subPaths[i].clipping === true) return true;
            }
        }
        return false;
    }

    /**
     * クリップグループの子を、最初のマスクとそれ以外（マスク内容）に分ける
     * @param {GroupItem} clippingGroup - 対象のクリップグループ
     * @returns {{maskItem: PageItem|null, contentItems: PageItem[]}} マスク（無ければ null）とマスク内容
     */
    function splitMaskAndContent(clippingGroup) {
        var maskItem = null;
        var contentItems = [];
        for (var i = 0; i < clippingGroup.pageItems.length; i++) {
            var pageItem = clippingGroup.pageItems[i];
            if (!maskItem && isMaskItem(pageItem)) {
                maskItem = pageItem;
            } else {
                contentItems.push(pageItem);
            }
        }
        return { maskItem: maskItem, contentItems: contentItems };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 入れ子のグループをどこまで解除するかのパネルを作る
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{radioReleaseAll: RadioButton, radioOneLevel: RadioButton}} 作ったコントロール
     */
    function addDepthPanel(parentDialog) {
        var depthPanel = parentDialog.add("panel", undefined, getLabel("panel.releaseDepth"));
        setupPanel(depthPanel, 6);

        var radioReleaseAll = depthPanel.add("radiobutton", undefined, getLabel("radio.releaseAll"));
        var radioOneLevel = depthPanel.add("radiobutton", undefined, getLabel("radio.releaseOneLevel"));
        radioReleaseAll.helpTip = getLabel("tooltip.releaseAll");
        radioOneLevel.helpTip = getLabel("tooltip.releaseOneLevel");

        if (USER_DEFAULTS.releaseDepth === "oneLevel") radioOneLevel.value = true;
        else radioReleaseAll.value = true;

        return { radioReleaseAll: radioReleaseAll, radioOneLevel: radioOneLevel };
    }

    /**
     * 入れ子のグループをどこまで解除するかを尋ねるダイアログ
     * @returns {string|null} "all" | "oneLevel"。キャンセルなら null
     */
    function askReleaseDepth() {
        var releaseDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(releaseDialog);

        var depthControls = addDepthPanel(releaseDialog);

        var buttonRow = addButtonRow(releaseDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        prepareDialogWindow(releaseDialog, SCRIPT_NAME);
        if (releaseDialog.show() !== 1) return null;
        return depthControls.radioOneLevel.value ? "oneLevel" : "all";
    }

    // =========================================
    // クリッピングマスクの解除 / Clipping mask release
    // =========================================

    /**
     * @typedef {Object} ClipReleaseResult
     * @property {PageItem[]} contentItems - グループの外に出したマスク内容
     * @property {PageItem|null} maskItem - 残したマスクパス（削除したときや見つからないときは null）
     */

    /**
     * クリップグループを解除方法に従って解除し、出てきたアイテムをまとめて返す。
     * 1つ失敗しても残りは処理し、最後に件数と理由を知らせる
     * @param {GroupItem[]} clippingGroups - 対象のクリップグループ
     * @param {string} clipReleaseMode - "simple" | "removePath" | "removeContent"
     * @param {boolean} shouldApplyFill - 残ったマスクパスに塗りを付けるなら true
     * @returns {{contentItems: PageItem[], maskItems: PageItem[]}} 出てきたマスク内容と、残したマスクパス
     */
    function releaseClippingGroups(clippingGroups, clipReleaseMode, shouldApplyFill) {
        var releasedContent = [];
        var keptMasks = [];
        var failureReasons = [];
        for (var i = 0; i < clippingGroups.length; i++) {
            try {
                var clipResult;
                if (clipReleaseMode === "removePath") clipResult = releaseKeepContent(clippingGroups[i]);
                else if (clipReleaseMode === "removeContent") clipResult = releaseKeepMaskPath(clippingGroups[i], shouldApplyFill);
                else clipResult = releaseKeepBoth(clippingGroups[i], shouldApplyFill);
                releasedContent = releasedContent.concat(clipResult.contentItems);
                if (clipResult.maskItem) keptMasks.push(clipResult.maskItem);
            } catch (err) {
                /* ExtendScript は message でなく description のことがある / ExtendScript may use description instead of message */
                failureReasons.push(err.message || err.description || String(err));
            }
        }
        if (failureReasons.length > 0) {
            alert(getLabel("alert.releaseFailed", { count: failureReasons.length }) + "\n" + failureReasons.join("\n"));
        }
        return { contentItems: releasedContent, maskItems: keptMasks };
    }

    /**
     * マスクパスと内容の両方を残して解除する。マスクが見つからなくても解除し、塗りだけ省く
     * @param {GroupItem} clippingGroup - 対象のクリップグループ
     * @param {boolean} shouldApplyFill - マスクパスに塗りを付けるなら true
     * @returns {ClipReleaseResult} 出てきたアイテム
     */
    function releaseKeepBoth(clippingGroup, shouldApplyFill) {
        var maskParts = splitMaskAndContent(clippingGroup);
        clippingGroup.clipped = false;
        if (shouldApplyFill && maskParts.maskItem) applyMaskFill(maskParts.maskItem);
        moveChildrenOutAndRemove(clippingGroup);
        return { contentItems: maskParts.contentItems, maskItem: maskParts.maskItem };
    }

    /**
     * マスク内容を残し、マスクパスを削除する
     * @param {GroupItem} clippingGroup - 対象のクリップグループ
     * @returns {ClipReleaseResult} 出てきたアイテム
     * @throws {Error} マスクが見つからないとき（このグループは処理しない）
     */
    function releaseKeepContent(clippingGroup) {
        /* 壊す前にマスクを確定する / Identify the mask before any destructive step */
        var maskParts = splitMaskAndContent(clippingGroup);
        if (!maskParts.maskItem) throw new Error("Clipping mask item not found.");
        clippingGroup.clipped = false;
        maskParts.maskItem.remove();
        moveChildrenOutAndRemove(clippingGroup);
        return { contentItems: maskParts.contentItems, maskItem: null };
    }

    /**
     * マスクパスを残し、マスク内容をすべて削除する
     * @param {GroupItem} clippingGroup - 対象のクリップグループ
     * @param {boolean} shouldApplyFill - 残ったマスクパスに塗りを付けるなら true
     * @returns {ClipReleaseResult} 出てきたアイテム
     * @throws {Error} マスクが見つからないとき（全部消してしまわないよう中止）
     */
    function releaseKeepMaskPath(clippingGroup, shouldApplyFill) {
        var maskParts = splitMaskAndContent(clippingGroup);
        if (!maskParts.maskItem) throw new Error("Clipping mask item not found.");
        clippingGroup.clipped = false;
        for (var i = 0; i < maskParts.contentItems.length; i++) {
            maskParts.contentItems[i].remove();
        }
        if (shouldApplyFill) applyMaskFill(maskParts.maskItem);
        moveChildrenOutAndRemove(clippingGroup);
        return { contentItems: [], maskItem: maskParts.maskItem };
    }

    /**
     * マスクパスにK100・PATH_FILL_OPACITY の塗りを付ける。テキストのマスクなどは対象外
     * @param {PageItem} maskItem - マスクパス
     * @returns {void}
     */
    function applyMaskFill(maskItem) {
        if (maskItem.typename === "PathItem") {
            maskItem.filled = true;
            maskItem.fillColor = getK100Black();
            maskItem.opacity = PATH_FILL_OPACITY;
        } else if (maskItem.typename === "CompoundPathItem") {
            maskItem.opacity = PATH_FILL_OPACITY;
            var subPaths = maskItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                subPaths[i].filled = true;
                subPaths[i].fillColor = getK100Black();
            }
        }
    }

    /**
     * K100（スミベタ）の CMYK カラーを返す
     * @returns {CMYKColor} K=100、ほかは0
     */
    function getK100Black() {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    /**
     * グループの子をグループの直前へ出してから、空のグループを消す。
     * 子を親の末尾へ送ると周りとの重なり順が崩れるので、グループの直前に順に置く
     * @param {GroupItem} targetGroup - 解除するグループ
     * @returns {void}
     */
    function moveChildrenOutAndRemove(targetGroup) {
        while (targetGroup.pageItems.length > 0) {
            targetGroup.pageItems[0].move(targetGroup, ElementPlacement.PLACEBEFORE);
        }
        targetGroup.remove();
    }

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

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

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /**
     * ［複合シェイプを解除］の一時アクションを1回再生する（AiSmartPathfinderPalette と同じ組み立て）。
     * セット名・アクション名は毎回ユニークにし、localizedName は記録どおり「複合シェイプを解除」にする。
     * PluginItem のうち複合シェイプにだけ効き、ブレンド・エンベロープはそのまま
     * @returns {boolean} 再生できたら true
     */
    function playReleaseCompoundShapeAction() {
        var uniqueToken = "ReleaseEverything_release_" + new Date().getTime() + "_" + Math.floor(Math.random() * 100000);
        var actionSetName = uniqueToken + "_set";
        var actionName = uniqueToken + "_action";
        var actionSource = [
            "/version 3"
        ].concat(buildActionNameLines("", actionSetName), [
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {"
        ], buildActionNameLines(" ", actionName), [
            " /keyIndex 0",
            " /colorIndex 0",
            " /isOpen 0",
            " /eventCount 1",
            " /event-1 {",
            " /useRulersIn1stQuadrant 0",
            " /internalName (ai_release_compound_shape)"
        ], buildActionNameLines(" ", "複合シェイプを解除", "localizedName"), [
            " /isOpen 0",
            " /isOn 1",
            " /hasDialog 0",
            " /parameterCount 1",
            " /parameter-1 {",
            " /key 1919710053",
            " /showInPalette 4294967295",
            " /type (integer)",
            " /value 0",
            " }",
            " }",
            "}"
        ]).join("\n") + "\n";

        return runTemporaryAction(actionSource, actionSetName, actionName);
    }

    // =========================================
    // 構造の解除 / Structure release
    // =========================================

    /* 試す解除のメニューコマンド（先頭から）。当てはまらないコマンドは何もせず例外も出さないので、
       元のアイテムが選択から消えたかどうかで解除できたかを判断する
       Release commands tried in order. Commands that do not apply do nothing and throw nothing,
       so success is judged by whether the original item left the selection */
    var GROUP_RELEASE_COMMANDS = [
        "Partial Rearrange Release", /* クロスと重なり / Intertwine */
        "ungroup"                    /* グループ解除 / Ungroup */
    ];
    var PLUGIN_RELEASE_COMMANDS = [   /* 複合シェイプは先に一時アクションで試す / compound shapes are tried first with the temporary action */
        "Path Blend Release",        /* ブレンド / Blend */
        "Release Envelope",          /* エンベロープ / Envelope */
        "Release Planet X",          /* ライブペイント / Live Paint */
        "Release Image Tracing",     /* 画像トレース / Image Trace */
        "Partial Rearrange Release"  /* クロスと重なり / Intertwine */
    ];

    /**
     * アイテム1つだけを選択して1段解除し、出てきたアイテムを返す。
     * 種類の違うものをまとめて選んだままメニューコマンドを当てると
     * 「オブジェクトのグループを解除できません」で止まるので、必ず1つずつ選び直す
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem} targetItem - 解除するアイテム
     * @returns {PageItem[]|null} 解除で出てきたアイテム。解除の対象でない・解除できなかったときは null
     */
    function releaseItemOnce(targetDoc, targetItem) {
        var typeName = targetItem.typename;
        if (typeName === "CompoundPathItem") return releaseByMenuCommands(targetDoc, targetItem, ["noCompoundPath"]);
        if (typeName === "GroupItem") return releaseByMenuCommands(targetDoc, targetItem, GROUP_RELEASE_COMMANDS);
        /* リピートは独自の型（グリッドは GridRepeatItem、2026-10-04 実測）/ Repeats have their own type (GridRepeatItem for grids, measured) */
        if (/RepeatItem$/.test(typeName)) return releaseByMenuCommands(targetDoc, targetItem, ["Release Repeat Art"]);
        if (typeName !== "PluginItem") return null;

        /* PluginItem は複合シェイプ・ブレンド・エンベロープなどを見分けられないので、解除を順に当てる
           A PluginItem cannot be told apart (compound shape, blend, envelope, ...), so try each release in turn */
        return releaseBySingleSelection(targetDoc, targetItem, playReleaseCompoundShapeAction)
            || releaseByMenuCommands(targetDoc, targetItem, PLUGIN_RELEASE_COMMANDS);
    }

    /**
     * メニューコマンドを順に当て、最初に解除できたときの出てきたアイテムを返す
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem} targetItem - 解除するアイテム
     * @param {string[]} commandIds - 試すメニューコマンドの ID（先頭から）
     * @returns {PageItem[]|null} 解除で出てきたアイテム。どれでも解除されなかったときは null
     */
    function releaseByMenuCommands(targetDoc, targetItem, commandIds) {
        for (var i = 0; i < commandIds.length; i++) {
            var commandId = commandIds[i];
            var releasedItems = releaseBySingleSelection(targetDoc, targetItem, function () { app.executeMenuCommand(commandId); });
            if (releasedItems) return releasedItems;
        }
        return null;
    }

    /**
     * アイテム1つだけを選択して解除の処理を当て、解除できていれば出てきたアイテムを返す。
     * 解除の処理は当てはまらなくても例外を出さないので、元のアイテムが選択に残るかで判断する。
     * 消えたアイテムも参照すると読めてしまう（2026-10-04 実測）ので、存在の確認には使わない。
     * 当てはまらない処理で選択が外れたときも、解除されていないとみなす
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem} targetItem - 解除するアイテム
     * @param {function(): void} releaseCommand - 選択に対して解除を行う処理
     * @returns {PageItem[]|null} 解除で出てきたアイテム。解除されなかったときは null
     */
    function releaseBySingleSelection(targetDoc, targetItem, releaseCommand) {
        selectItems(targetDoc, [targetItem]);
        /* DOM で選び直した直後は画面に反映してからコマンドを当てる / Redraw after a DOM selection change before running the command */
        app.redraw();
        releaseCommand();
        /* ブレンドの解除などは画面を更新するまで確定しない（実測）/ Some releases, such as blends, are not committed until a redraw (measured) */
        app.redraw();

        var releasedItems = getSelectionItems(targetDoc);
        if (!releasedItems.length) return null;
        for (var i = 0; i < releasedItems.length; i++) {
            if (releasedItems[i] === targetItem) return null;
        }
        return releasedItems;
    }

    /**
     * 指定したアイテムを1つずつ、入れ子のものまで解除し尽くす
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem[]} targetItems - 解除するアイテム
     * @returns {PageItem[]} 解除し終えたアイテム
     */
    function releaseItemsAllLevels(targetDoc, targetItems) {
        var pendingItems = targetItems.slice();
        var resultItems = [];
        while (pendingItems.length) {
            var currentItem = pendingItems.shift();
            var releasedItems = releaseItemOnce(targetDoc, currentItem);
            /* 出てきたものを先頭に戻して続けて解除する / Put what came out back at the front and keep releasing */
            if (releasedItems) pendingItems = releasedItems.concat(pendingItems);
            else resultItems.push(currentItem);
        }
        return resultItems;
    }

    /**
     * 指定したアイテムを1つずつ1段だけ解除する（出てきたものは解除しない）
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem[]} targetItems - 解除するアイテム
     * @returns {PageItem[]} 解除後のアイテム（解除しなかったものも含む）
     */
    function releaseItemsOneLevel(targetDoc, targetItems) {
        var resultItems = [];
        for (var i = 0; i < targetItems.length; i++) {
            var releasedItems = releaseItemOnce(targetDoc, targetItems[i]);
            resultItems = resultItems.concat(releasedItems || [targetItems[i]]);
        }
        return resultItems;
    }

    // =========================================
    // テキストの回り込みの解除 / Text wrap release
    // =========================================

    /**
     * 解除し終えたアイテムのテキストの回り込みを解除する。
     * 中身が出てくる解除ではないので、構造の解除が済んでから1つずつ当てる
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem[]} targetItems - 対象のアイテム
     * @returns {void}
     */
    function releaseTextWraps(targetDoc, targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            releaseTextWrap(targetDoc, targetItems[i]);
        }
    }

    /**
     * テキストの回り込みが設定されていれば解除する
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {PageItem} targetItem - 対象のアイテム
     * @returns {void}
     */
    function releaseTextWrap(targetDoc, targetItem) {
        var isWrapped = false;
        try {
            isWrapped = targetItem.wrapped === true;
        } catch (e) {
            /* 回り込みを持たない種類のアイテム / item types without text wrap */
        }
        if (!isWrapped) return;
        selectItems(targetDoc, [targetItem]);
        app.executeMenuCommand("Release Text Wrap");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var activeDoc = app.activeDocument;
        var selectionItems = getSelectionItems(activeDoc);
        if (!selectionItems.length) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var clippingGroups = [];
        var otherItems = [];
        for (var i = 0; i < selectionItems.length; i++) {
            if (isClippingGroup(selectionItems[i])) clippingGroups.push(selectionItems[i]);
            else otherItems.push(selectionItems[i]);
        }
        var hasNestedGroup = containsNestedGroup(selectionItems);

        /* ダイアログは入れ子のグループがあるときだけ。クリップグループは USER_DEFAULTS の方法で解除する
           Ask only when there are nested groups; clipping groups are released as set in USER_DEFAULTS */
        var releaseDepth = "all";
        if (hasNestedGroup) {
            releaseDepth = askReleaseDepth();
            if (!releaseDepth) return;
        }

        var clipResult;
        if (releaseDepth === "oneLevel") {
            /* クリップグループの解除で1階層ぶんなので、ほかのアイテムだけを1階層解除する
               Releasing a clipping group already counts as one level, so only the other items are released once */
            var releasedOthers = releaseItemsOneLevel(activeDoc, otherItems);
            clipResult = releaseClippingGroups(clippingGroups, USER_DEFAULTS.clipReleaseMode, USER_DEFAULTS.applyMaskFill);
            var oneLevelItems = releasedOthers.concat(clipResult.contentItems, clipResult.maskItems);
            releaseTextWraps(activeDoc, oneLevelItems);
            selectItems(activeDoc, oneLevelItems);
            return;
        }

        /* 残したマスクパスは複合パスを崩さないよう、続く解除から外す
           Keep the remaining mask paths out of the following release so compound masks stay intact */
        clipResult = releaseClippingGroups(clippingGroups, USER_DEFAULTS.clipReleaseMode, USER_DEFAULTS.applyMaskFill);
        var releasedItems = releaseItemsAllLevels(activeDoc, otherItems.concat(clipResult.contentItems));
        var allLevelItems = releasedItems.concat(clipResult.maskItems);
        releaseTextWraps(activeDoc, allLevelItems);
        selectItems(activeDoc, allLevelItems);
    }

    main();
})();
