#target illustrator
#targetengine "DashGapCalculatorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパス（オープン／クローズ）の長さをもとに、分割数と間隔から破線の線分長を計算して適用します。
線分から間隔を逆算するモードや、ランダムパターン、開始位置（位相）の指定にも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashGapCalculator.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n868bedb96542

### Overview

Works out the dash length for the selected paths, open or closed, from their length, the number of divisions and the gap.
It can also derive the gap from the dash, produce random patterns, and set the starting phase.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DashGapCalculator.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DashGapCalculator";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-06";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashGapCalculator.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DashGapCalculator.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n868bedb96542"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ランダムモードで生成する線分・間隔の範囲（現在の線の単位）/ Random dash & gap range in the current stroke unit */
    var RANDOM_DASH_MAX     = 40;    /* 線分の最大値 / max dash */
    var RANDOM_GAP_MIN      = 3;     /* 間隔の最小値 / min gap */
    var RANDOM_GAP_MAX      = 3;     /* 間隔の最大値 / max gap */
    var RANDOM_ROUND_VALUES = false; /* 生成値を整数に丸めるか / round generated values */

    /* 線端「なし」のときに線分が消えないための最小値（単位コード別）/ Minimum dash for butt caps, per unit code */
    var RANDOM_DASH_MIN_BY_UNIT = {
        1: 1, /* mm */
        2: 2, /* pt */
        5: 4  /* Q / H */
    };
    var RANDOM_DASH_MIN_DEFAULT = 2;

    /* ランダムパターンの要素数（線分・間隔を3組）/ Entries in one random pattern */
    var RANDOM_PATTERN_LENGTH = 6;

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
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

    /* このダイアログ固有の寸法 / Sizes specific to this dialog */
    var OPTION_ROW_SPACING       = 20;           /* 下部チェックボックスの間隔 */
    var OFFSET_PANEL_SPACING     = 6;            /* 開始位置パネルの要素間隔（密） */
    var FIELD_LABEL_WIDTH        = 40;           /* 分割数・間隔・線分のラベル幅 */
    var NUMBER_FIELD_CHARS       = 4;            /* 数値入力欄の幅（文字数） */
    var LENGTH_FIELD_CHARS       = 6;            /* 単位付きの入力欄の幅（文字数） */
    var LABELLESS_CHECKBOX_WIDTH = 18;           /* ラベルなしチェックボックスの幅 */

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

    /**
     * ラベル付きパネルを生成し、共通レイアウトを適用する
     * @param {Window|Panel|Group} parentContainer - 追加先のコンテナ
     * @param {string} titleText - パネルのタイトル
     * @param {number} spacing - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, titleText, spacing) {
        var createdPanel = parentContainer.add("panel", undefined, titleText);
        setupPanel(createdPanel, spacing);
        return createdPanel;
    }

    /**
     * 横並びのグループを生成する
     * @param {Window|Panel|Group} parentContainer - 追加先のコンテナ
     * @param {string} alignment - 親の中での配置（省略時は "left"）
     * @param {number} spacing - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Group} 生成したグループ
     */
    function addRow(parentContainer, alignment, spacing) {
        var rowGroup = parentContainer.add("group");
        setupRow(rowGroup, alignment, spacing);
        return rowGroup;
    }

    /**
     * 縦積みのグループを生成する
     * @param {Window|Panel|Group} parentContainer - 追加先のコンテナ
     * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "top"]）
     * @param {string} alignment - 親の中での配置（省略時はコンテナ既定）
     * @returns {Group} 生成したグループ
     */
    function addColumn(parentContainer, alignChildren, alignment) {
        var columnGroup = parentContainer.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = alignChildren || ["fill", "top"];
        if (alignment) columnGroup.alignment = alignment;
        return columnGroup;
    }

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "破線計算機", en: "Dash Calculator (Gap→Dash)" }
        },
        panel: {
            pathInfo:   { ja: "選択中のパス情報", en: "Selected Path Info" },
            dashCalc:   { ja: "破線の計算", en: "Dash Calculation" },
            calcMethod: { ja: "計算方法", en: "Calculation" },
            offset:     { ja: "開始位置", en: "Offset" },
            partial:    { ja: "部分表示", en: "Partial Display" },
            cap:        { ja: "線端", en: "Cap" }
        },
        fieldLabel: {
            pathLength: { ja: "パスの長さ", en: "Path length" },
            segments:   { ja: "分割数", en: "Segments" },
            gap:        { ja: "間隔", en: "Gap" },
            dash:       { ja: "線分", en: "Dash" }
        },
        radio: {
            gapToDash:  { ja: "間隔→線分", en: "Gap→Dash" },
            dashToGap:  { ja: "線分→間隔", en: "Dash→Gap" },
            random:     { ja: "ランダム", en: "Random" },
            capButt:    { ja: "なし", en: "Butt" },
            capRound:   { ja: "丸型", en: "Round" },
            capProject: { ja: "突出", en: "Projecting" }
        },
        checkbox: {
            partialDisplay: { ja: "部分表示", en: "Partial Display" },
            adjustEnds:     { ja: "両端を調整", en: "Adjust ends" },
            reversePath:    { ja: "パスの方向反転", en: "Reverse Path Direction" }
        },
        button: {
            ok:        { ja: "OK", en: "OK" },
            cancel:    { ja: "キャンセル", en: "Cancel" },
            clearDash: { ja: "破線クリア", en: "Clear Dashes" }
        },
        tooltip: {
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            pathInfo: {
                ja: "先頭のパスの長さ。括弧内は選択しているパスの数",
                en: "Length of the first path. The number in parentheses is the count of selected paths"
            },
            segments: {
                ja: "パスをいくつに分けるか。線分＋間隔の繰り返し回数になります",
                en: "How many parts the path is divided into — the number of dash + gap cycles"
            },
            gap: {
                ja: "破線のすき間の長さ",
                en: "Length of the empty space between dashes"
            },
            dash: {
                ja: "破線の線の長さ",
                en: "Length of each dash"
            },
            gapToDash: {
                ja: "間隔を入力して線分の長さを求めます",
                en: "Enter the gap; the dash length is calculated"
            },
            dashToGap: {
                ja: "線分の長さを入力して間隔を求めます",
                en: "Enter the dash length; the gap is calculated"
            },
            random: {
                ja: "線分をランダムに割り当てます。パスごとに別の乱数を使い、クリックのたびに作り直します",
                en: "Assigns random dashes. Each path gets its own draw, and every click generates a new pattern"
            },
            useOffset: {
                ja: "破線の開始位置（位相）を指定します。パスの始点から数えます",
                en: "Sets the dash offset (phase), measured from the start point of the path"
            },
            offsetPreset: {
                ja: "1周期（線分＋間隔）に対する比率で開始位置を決めます",
                en: "Sets the offset as a fraction of one cycle (dash + gap)"
            },
            partialDisplay: {
                ja: "分割した1本分の線分だけを表示し、残りを隠します",
                en: "Shows only one of the divided dashes and hides the rest"
            },
            cap: {
                ja: "破線の線端の形。丸型・突出は線分の長さより少しはみ出します",
                en: "Shape of the dash ends. Round and Projecting extend slightly beyond the dash length"
            },
            adjustEnds: {
                ja: "オープンパスで、両端が線分で終わるように配分します（クローズパスでは使いません）",
                en: "On an open path, distributes the dashes so both ends finish with a dash (unused for closed paths)"
            },
            reversePath: {
                ja: "パスの向きを反転して、破線の開始位置を反対の端へ移します",
                en: "Reverses the path direction, moving the dash start to the other end"
            },
            clearDash: {
                ja: "破線を解除した状態をプレビューします。OKで確定します",
                en: "Previews the paths without dashes. Click OK to confirm"
            }
        },
        alert: {
            calcError:       { ja: "エラー", en: "Error" },
            noDocument:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:     { ja: "対象となるパス（線や円など）を選択してください。", en: "Select one path (open or closed)." },
            noPathSelected:  { ja: "パス（オープン/クローズ）を1つ選択してください。", en: "Select exactly one path (open or closed)." },
            segmentsInvalid: { ja: "分割数は1以上の整数を入力してください。", en: "Enter an integer of 1 or greater for Segments." },
            gapInvalid:      { ja: "間隔 (Gap) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Gap." },
            dashInvalid:     { ja: "線分 (Dash) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Dash." },
            offsetInvalid:   { ja: "開始位置 (Offset) は数値を入力してください。", en: "Enter a number for Offset." },
            gapTooLong: {
                ja: "間隔 (Gap) が長すぎます。線分がゼロまたはマイナスになってしまいます。\n(設定可能な最大Gap: ほぼ %1 %2)",
                en: "Gap is too long; dash would be zero or negative.\n(Max allowed Gap: about %1 %2)"
            },
            dashTooLong: {
                ja: "線分 (Dash) が長すぎます。間隔がゼロまたはマイナスになってしまいます。\n(設定可能な最大Dash: ほぼ %1 %2)",
                en: "Dash is too long; gap would be zero or negative.\n(Max allowed Dash: about %1 %2)"
            }
        }
    };

    // =========================================
    // 単位ユーティリティ / Unit utilities
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位値を pt に変換する
     * @param {number} value - 単位値
     * @param {Object} unitInfo - getUnitInfo() の戻り値
     * @returns {number} pt 値
     */
    function unitToPt(value, unitInfo) {
        return value * unitInfo.pointsPerUnit;
    }

    /**
     * pt 値を単位値に変換する
     * @param {number} ptValue - pt 値
     * @param {Object} unitInfo - getUnitInfo() の戻り値
     * @returns {number} 単位値
     */
    function ptToUnit(ptValue, unitInfo) {
        return ptValue / unitInfo.pointsPerUnit;
    }

    /**
     * 入力欄向けに数値を整形する（小数第3位まで、整数はそのまま）
     * @param {number} value - 表示したい数値
     * @returns {string} 整形後の文字列
     */
    function formatFieldNumber(value) {
        if (value == null || isNaN(value)) return "0";
        var rounded = Math.round(value * 1000) / 1000;
        if (Math.abs(rounded - Math.round(rounded)) < 1e-10) return String(Math.round(rounded));
        return String(rounded);
    }

    // =========================================
    // 前回値の記憶 / Session settings
    // =========================================

    var PREF_KEY = "DashCalcPrefs_GapToDash_v1";

    /* 保存する項目と ActionDescriptor の型 / Saved fields and their descriptor types */
    var PREF_FIELD_TYPES = {
        segments:    "Integer",
        gapPt:       "Double",
        dashPt:      "Double",
        offsetPt:    "Double",
        capMode:     "Integer",
        mode:        "Integer",
        reversePath: "Boolean",
        adjustEnds:  "Boolean",
        useOffset:   "Boolean"
    };

    /* ランダムパターン保存用のキー / Keys for the random pattern */
    var RANDOM_PREF_KEYS = ["rand0Pt", "rand1Pt", "rand2Pt", "rand3Pt", "rand4Pt", "rand5Pt"];

    /**
     * 前回値を読み込む
     * @returns {Object} 保存値のオブジェクト（読み込めない場合は null）
     */
    function loadPrefs() {
        try {
            /* 保存値が無いと getCustomOptions は例外を投げる / throws when nothing is saved */
            var descriptor = app.getCustomOptions(PREF_KEY);
            var savedPrefs = {};

            for (var fieldName in PREF_FIELD_TYPES) {
                var fieldKey = stringIDToTypeID(fieldName);
                if (descriptor.hasKey(fieldKey)) savedPrefs[fieldName] = descriptor["get" + PREF_FIELD_TYPES[fieldName]](fieldKey);
            }

            /* ランダムパターンは先頭から連続している分だけ読む（旧バージョンの4要素も許容） */
            var randomDashes = [];
            for (var i = 0; i < RANDOM_PREF_KEYS.length; i++) {
                var randomKey = stringIDToTypeID(RANDOM_PREF_KEYS[i]);
                if (!descriptor.hasKey(randomKey)) break;
                randomDashes.push(descriptor.getDouble(randomKey));
            }
            if (randomDashes.length > 0) savedPrefs.randPt = randomDashes;

            return savedPrefs;
        } catch (e) {
            return null;
        }
    }

    /**
     * 前回値を保存する
     * @param {Object} newPrefs - 保存する設定値
     * @returns {void}
     */
    function savePrefs(newPrefs) {
        try {
            var descriptor = new ActionDescriptor();
            for (var fieldName in PREF_FIELD_TYPES) {
                var fieldType = PREF_FIELD_TYPES[fieldName];
                var fieldValue = (fieldType === "Boolean") ? !!newPrefs[fieldName] : newPrefs[fieldName];
                descriptor["put" + fieldType](stringIDToTypeID(fieldName), fieldValue);
            }

            if (newPrefs.randPt && newPrefs.randPt.length === RANDOM_PATTERN_LENGTH) {
                for (var i = 0; i < RANDOM_PREF_KEYS.length; i++) {
                    descriptor.putDouble(stringIDToTypeID(RANDOM_PREF_KEYS[i]), newPrefs.randPt[i]);
                }
            }
            app.putCustomOptions(PREF_KEY, descriptor, true);
        } catch (e) { }
    }

    /**
     * 前回値の長さ（pt）を現在の単位に変換して返す
     * @param {number} storedPt - 保存されている長さ（pt）
     * @param {number} fallbackUnit - 保存値がないときの値（単位値）
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @returns {number} 長さ（単位値）
     */
    function toInitialUnit(storedPt, fallbackUnit, strokeUnit) {
        return (typeof storedPt === "number" && storedPt >= 0) ? ptToUnit(storedPt, strokeUnit) : fallbackUnit;
    }

    /**
     * 前回値の真偽値を返す
     * @param {boolean} storedFlag - 保存されている値
     * @param {boolean} fallbackFlag - 保存値がないときの値
     * @returns {boolean} 復元した値
     */
    function toInitialFlag(storedFlag, fallbackFlag) {
        return (typeof storedFlag === "boolean") ? storedFlag : fallbackFlag;
    }

    /**
     * 前回値の数値を返す
     * @param {number} storedNumber - 保存されている値
     * @param {number} fallbackNumber - 保存値がないときの値
     * @param {number} minValue - 許容する下限値
     * @returns {number} 復元した値
     */
    function toInitialNumber(storedNumber, fallbackNumber, minValue) {
        return (typeof storedNumber === "number" && storedNumber >= minValue) ? storedNumber : fallbackNumber;
    }

    /**
     * 前回値からダイアログの初期値を決める（保存値が無い項目は既定値）
     * @param {Object} savedPrefs - loadPrefs() の戻り値（読み込めない場合は null）
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @returns {Object} 初期値（長さは現在の線の単位）
     */
    function readInitialValues(savedPrefs, strokeUnit) {
        var storedPrefs = savedPrefs || {};
        return {
            segments:    toInitialNumber(storedPrefs.segments, 3, 1),
            capMode:     toInitialNumber(storedPrefs.capMode, 0, 0),
            mode:        toInitialNumber(storedPrefs.mode, 0, 0), /* 0:間隔→線分 / 1:線分→間隔 / 2:ランダム */
            gapUnit:     toInitialUnit(storedPrefs.gapPt, 5, strokeUnit),
            dashUnit:    toInitialUnit(storedPrefs.dashPt, 0, strokeUnit),
            offsetUnit:  toInitialUnit(storedPrefs.offsetPt, 0, strokeUnit),
            useOffset:   toInitialFlag(storedPrefs.useOffset, false),
            adjustEnds:  toInitialFlag(storedPrefs.adjustEnds, true),
            reversePath: toInitialFlag(storedPrefs.reversePath, false)
        };
    }

    // =========================================
    // 破線の計算 / Dash calculation
    // =========================================

    /**
     * 間隔から線分長と1周期（線分＋間隔）の長さを求める
     * クローズパスと「両端を調整」OFFは 1周期＝全長÷分割数、
     * オープンパスで「両端を調整」ONは両端が線分で終わるように配分する。
     * @param {number} segments - 分割数
     * @param {number} gapPt - 間隔（pt）
     * @param {number} pathLen - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスかどうか
     * @param {boolean} adjustEnds - 両端を調整するかどうか
     * @returns {Object} { dashPt:number, cyclePt:number }（計算できない場合は null）
     */
    function calcDashAndCyclePt(segments, gapPt, pathLen, isClosed, adjustEnds) {
        if (!(segments > 0)) return null;
        if (!(gapPt >= 0)) return null;

        if (isClosed || !adjustEnds) {
            var cyclePt = pathLen / segments;
            return { dashPt: cyclePt - gapPt, cyclePt: cyclePt };
        }

        /* 分割数＝線分の本数、間隔は（分割数−1）回 */
        if (segments === 1) return { dashPt: pathLen, cyclePt: pathLen + gapPt };

        var dashPt = (pathLen - gapPt * (segments - 1)) / segments;
        return { dashPt: dashPt, cyclePt: dashPt + gapPt };
    }

    /**
     * 線分長から間隔を逆算する
     * @param {number} segments - 分割数
     * @param {number} dashPt - 線分長（pt）
     * @param {number} pathLen - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスかどうか
     * @param {boolean} adjustEnds - 両端を調整するかどうか
     * @returns {number} 間隔（pt）。計算できない場合は null
     */
    function calcGapPtFromDashPt(segments, dashPt, pathLen, isClosed, adjustEnds) {
        if (!(segments > 0)) return null;
        if (!(dashPt >= 0)) return null;

        if (isClosed || !adjustEnds) return (pathLen / segments) - dashPt;

        if (segments === 1) return 0;
        return (pathLen - dashPt * segments) / (segments - 1);
    }

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
    };

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 右揃えの項目名を持つ行を追加する
     * @param {Window|Panel|Group} parentContainer - 追加先のコンテナ
     * @param {Object} labelNode - 項目名の LABELS ノード
     * @returns {Group} 追加した行
     */
    function addFieldRow(parentContainer, labelNode) {
        var fieldRow = addRow(parentContainer);
        var fieldLabel = fieldRow.add("statictext", undefined, labelText(labelNode));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return fieldRow;
    }

    /**
     * 間隔・線分の行を追加する（入力欄と計算結果の表示を重ね、計算方法に応じて切り替える）
     * @param {Window|Panel|Group} parentContainer - 追加先のコンテナ
     * @param {Object} labelNode - 項目名の LABELS ノード
     * @param {number} initialValue - 入力欄の初期値（単位値）
     * @param {string} fieldUnit - 入力欄の単位（例 " mm"）
     * @param {Object} tooltipNode - tooltip の LABELS ノード
     * @returns {{row: Group, field: EditText, resultLabel: StaticText}} 行・入力欄・結果表示
     */
    function addDashGapRow(parentContainer, labelNode, initialValue, fieldUnit, tooltipNode) {
        var fieldRow = addFieldRow(parentContainer, labelNode);

        var numberField;
        var stepperInputGroup = addStepperInputGroup(fieldRow, function () { return numberField; }, { step: 1, min: 0, unit: fieldUnit });

        /* 入力欄と計算結果を重ね、計算方法に応じて切り替える / the field and the computed result share one slot */
        var fieldStack = stepperInputGroup.add("group");
        fieldStack.orientation = "stack";

        numberField = fieldStack.add("edittext", undefined, formatFieldNumber(initialValue) + fieldUnit);
        numberField.characters = LENGTH_FIELD_CHARS;
        numberField.stepperGroup = stepperInputGroup.stepperGroup;
        bindSteppedArrowKeys(numberField, numberField.stepperGroup);

        var resultLabel = fieldStack.add("statictext", undefined, "");
        resultLabel.justify = "right";
        resultLabel.preferredSize.width = numberField.preferredSize.width;

        fieldRow.helpTip = numberField.helpTip = getLabel(tooltipNode);
        return { row: fieldRow, field: numberField, resultLabel: resultLabel };
    }

    /**
     * ∧∨を入れた行 group（spacing 0）を追加する。入力欄はこの group に続けて足し、∧∨と突き合わせる
     * @param {Group} parentRow - 追加先の行
     * @param {Function} getNumberInput - 対象の入力欄を返す関数
     * @param {Object} stepOptions - step / min / integer / unit（onStep は入力欄の _onArrowChange を呼ぶ形に固定）
     * @returns {Group} 追加した group（∧∨は .stepperGroup で参照できる）
     */
    function addStepperInputGroup(parentRow, getNumberInput, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        /* 増減後の処理（プレビュー更新）は、イベント設定時に入力欄の _onArrowChange へ入る / set later by the event wiring */
        stepOptions.onStep = function (numberInput) {
            if (typeof numberInput._onArrowChange === "function") numberInput._onArrowChange();
        };
        stepperInputGroup.stepperGroup = addStepper(stepperInputGroup, getNumberInput, stepOptions);
        return stepperInputGroup;
    }

    /**
     * 開始位置の入力欄と∧∨の有効／無効をまとめて切り替える
     * @param {EditText} numberInput - stepperGroup を持つ入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * 選択中のパス情報の表示文字列を作る
     * @param {number} pathLength - 先頭のパスの長さ（pt）
     * @param {number} pathCount - 選択しているパスの数
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @returns {string} 「パスの長さ：123.456 mm  (3)」形式の文字列
     */
    function formatPathInfoText(pathLength, pathCount, strokeUnit) {
        var pathInfoText = labelText(LABELS.fieldLabel.pathLength) + " " +
            ptToUnit(pathLength, strokeUnit).toFixed(3) + " " + strokeUnit.label;
        if (pathCount > 1) pathInfoText += "  (" + pathCount + ")";
        return pathInfoText;
    }

    /**
     * ダイアログを組み立てる（イベントは設定しない）
     * @param {string} pathInfoText - 選択中のパス情報の表示
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @param {Object} initialValues - readInitialValues() の戻り値
     * @returns {Object} ダイアログと操作に使うコントロール
     */
    function buildDashDialog(pathInfoText, strokeUnit, initialValues) {
        var dashDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dashDialog);

        /* 選択中のパス情報（全幅）*/
        var pathInfoPanel = addPanel(dashDialog, getLabel(LABELS.panel.pathInfo));
        var lblPathInfo = pathInfoPanel.add("statictext", undefined, pathInfoText);
        lblPathInfo.alignment = "center";
        lblPathInfo.helpTip = getLabel(LABELS.tooltip.pathInfo);

        /* 2カラム */
        var columnsGroup = dashDialog.add("group");
        setupRow(columnsGroup, "fill", COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "top"];

        var leftColumn = addColumn(columnsGroup);
        var rightColumn = addColumn(columnsGroup);

        /* 破線の計算（左カラム）*/
        var dashCalcPanel = addPanel(leftColumn, getLabel(LABELS.panel.dashCalc));
        var dashInputColumn = addColumn(dashCalcPanel, ["left", "top"], "left");

        /* 分割数 */
        var segmentsRow = addFieldRow(dashInputColumn, LABELS.fieldLabel.segments);
        var txtSegments;
        var segmentsInputGroup = addStepperInputGroup(segmentsRow, function () { return txtSegments; }, { step: 1, min: 1, integer: true });
        txtSegments = segmentsInputGroup.add("edittext", undefined, String(initialValues.segments));
        txtSegments.characters = NUMBER_FIELD_CHARS;
        txtSegments.stepperGroup = segmentsInputGroup.stepperGroup;
        bindSteppedArrowKeys(txtSegments, txtSegments.stepperGroup);
        segmentsRow.helpTip = txtSegments.helpTip = getLabel(LABELS.tooltip.segments);

        /* 間隔・線分 */
        /* 単位は入力欄に入れる（別の単位で入れても線の単位へ換算される） / the unit lives in the field */
        var strokeFieldUnit = " " + strokeUnit.label;
        var gapRow = addDashGapRow(dashInputColumn, LABELS.fieldLabel.gap, initialValues.gapUnit, strokeFieldUnit, LABELS.tooltip.gap);
        var dashRow = addDashGapRow(dashInputColumn, LABELS.fieldLabel.dash, initialValues.dashUnit, strokeFieldUnit, LABELS.tooltip.dash);

        /* 計算方法（左カラム）*/
        var calcMethodPanel = addPanel(leftColumn, getLabel(LABELS.panel.calcMethod));
        var calcModeColumn = addColumn(calcMethodPanel, ["left", "top"], "left");

        var rbModeGapToDash = calcModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.gapToDash));
        var rbModeDashToGap = calcModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.dashToGap));
        var rbModeRandom = calcModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.random));
        rbModeGapToDash.helpTip = getLabel(LABELS.tooltip.gapToDash);
        rbModeDashToGap.helpTip = getLabel(LABELS.tooltip.dashToGap);
        rbModeRandom.helpTip = getLabel(LABELS.tooltip.random);
        rbModeGapToDash.value = (initialValues.mode === 0);
        rbModeDashToGap.value = (initialValues.mode === 1);
        rbModeRandom.value = (initialValues.mode === 2);

        /* 開始位置（右カラム）*/
        var offsetPanel = addPanel(rightColumn, getLabel(LABELS.panel.offset), OFFSET_PANEL_SPACING);

        var offsetRow = addRow(offsetPanel);
        var chkUseOffset = offsetRow.add("checkbox", undefined, "");
        chkUseOffset.value = initialValues.useOffset;
        chkUseOffset.preferredSize.width = LABELLESS_CHECKBOX_WIDTH;

        var txtOffset;
        var offsetInputGroup = addStepperInputGroup(offsetRow, function () { return txtOffset; }, { step: 1, min: 0, unit: strokeFieldUnit });
        txtOffset = offsetInputGroup.add("edittext", undefined, formatFieldNumber(initialValues.offsetUnit) + strokeFieldUnit);
        txtOffset.characters = LENGTH_FIELD_CHARS;
        txtOffset.stepperGroup = offsetInputGroup.stepperGroup;
        bindSteppedArrowKeys(txtOffset, txtOffset.stepperGroup);
        setStepperInputEnabled(txtOffset, chkUseOffset.value);
        offsetRow.helpTip = chkUseOffset.helpTip = txtOffset.helpTip = getLabel(LABELS.tooltip.useOffset);

        /* 1周期（線分＋間隔）を基準にしたプリセット */
        var offsetPresetRow = addRow(offsetPanel);
        offsetPresetRow.enabled = chkUseOffset.value;
        var rbOffsetQuarter = offsetPresetRow.add("radiobutton", undefined, "1/4");
        var rbOffsetHalf = offsetPresetRow.add("radiobutton", undefined, "1/2");
        var rbOffsetThreeQuarter = offsetPresetRow.add("radiobutton", undefined, "3/4");
        offsetPresetRow.helpTip = rbOffsetQuarter.helpTip = rbOffsetHalf.helpTip = rbOffsetThreeQuarter.helpTip =
            getLabel(LABELS.tooltip.offsetPreset);

        /* 部分表示（右カラム）*/
        var partialPanel = addPanel(rightColumn, getLabel(LABELS.panel.partial));
        var chkPartialDisplay = partialPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.partialDisplay));
        chkPartialDisplay.alignment = "left";
        chkPartialDisplay.value = false;
        chkPartialDisplay.helpTip = getLabel(LABELS.tooltip.partialDisplay);

        /* 線端（右カラム）*/
        var capPanel = addPanel(rightColumn, getLabel(LABELS.panel.cap));
        var capRow = addRow(capPanel);
        var rbCapButt = capRow.add("radiobutton", undefined, getLabel(LABELS.radio.capButt));
        var rbCapRound = capRow.add("radiobutton", undefined, getLabel(LABELS.radio.capRound));
        var rbCapProject = capRow.add("radiobutton", undefined, getLabel(LABELS.radio.capProject));
        capRow.helpTip = rbCapButt.helpTip = rbCapRound.helpTip = rbCapProject.helpTip = getLabel(LABELS.tooltip.cap);
        if (initialValues.capMode === 1) rbCapRound.value = true;
        else if (initialValues.capMode === 2) rbCapProject.value = true;
        else rbCapButt.value = true;

        /* 両端を調整・パスの方向反転（中央）*/
        var pathOptionRow = addRow(dashDialog, "center", OPTION_ROW_SPACING);

        var chkAdjustEnds = pathOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.adjustEnds));
        chkAdjustEnds.value = initialValues.adjustEnds;
        chkAdjustEnds.helpTip = getLabel(LABELS.tooltip.adjustEnds);

        var chkReversePath = pathOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.reversePath));
        chkReversePath.value = initialValues.reversePath;
        chkReversePath.helpTip = getLabel(LABELS.tooltip.reversePath);

        /* ボタン（左：破線クリア／右：キャンセル・OK） / Buttons (left: Clear Dash, right: Cancel and OK) */
        var buttonRow = addButtonRow(dashDialog);
        var btnClearDash = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.clearDash));
        btnClearDash.helpTip = getLabel(LABELS.tooltip.clearDash);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            dialog: dashDialog,
            segmentsRow: segmentsRow,
            txtSegments: txtSegments,
            txtGap: gapRow.field,
            lblGapResult: gapRow.resultLabel,
            dashRow: dashRow.row,
            txtDash: dashRow.field,
            lblDashResult: dashRow.resultLabel,
            rbModeGapToDash: rbModeGapToDash,
            rbModeDashToGap: rbModeDashToGap,
            rbModeRandom: rbModeRandom,
            chkUseOffset: chkUseOffset,
            txtOffset: txtOffset,
            offsetPresetRow: offsetPresetRow,
            rbOffsetQuarter: rbOffsetQuarter,
            rbOffsetHalf: rbOffsetHalf,
            rbOffsetThreeQuarter: rbOffsetThreeQuarter,
            chkPartialDisplay: chkPartialDisplay,
            rbCapButt: rbCapButt,
            rbCapRound: rbCapRound,
            rbCapProject: rbCapProject,
            chkAdjustEnds: chkAdjustEnds,
            chkReversePath: chkReversePath,
            btnClearDash: btnClearDash,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
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
    // メイン処理 / Main
    // =========================================

    /**
     * バウンディングボックスをリセットし、エッジ表示を切り替える
     * 実行時と終了時に呼び、表示を元へ戻す。
     * @returns {void}
     */
    function resetBoundsAndToggleEdges() {
        app.executeMenuCommand('AI Reset Bounding Box');
        app.executeMenuCommand('edge');
    }

    /**
     * 線の状態（破線・線端・線の有無）を控える
     * @param {Array<PathItem>} targetPaths - 対象のパス
     * @returns {Array<Object>} パスと線の状態の組
     */
    function captureStrokeStates(targetPaths) {
        var strokeStates = [];
        for (var i = 0; i < targetPaths.length; i++) {
            var pathItem = targetPaths[i];
            strokeStates.push({
                item: pathItem,
                stroked: pathItem.stroked,
                strokeCap: pathItem.strokeCap,
                strokeDashes: (pathItem.strokeDashes && pathItem.strokeDashes.length) ? pathItem.strokeDashes.slice(0) : [],
                strokeDashOffset: (typeof pathItem.strokeDashOffset === "number") ? pathItem.strokeDashOffset : 0
            });
        }
        return strokeStates;
    }

    /**
     * ダイアログを表示し、対象のパスに破線を適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PathItem>} targetPaths - 対象のパス
     * @returns {void}
     */
    function showDashDialog(doc, targetPaths) {
        /* 先頭のパスをUI表示・計算の代表として扱う */
        var primaryPath = targetPaths[0];
        var primaryPathLength = primaryPath.length;
        var strokeUnit = getUnitInfo("strokeUnits");
        var strokeFieldUnit = " " + strokeUnit.label; /* 入力欄・結果表示に付ける単位 / unit shown in the fields */

        /* ダイアログを開く前の状態（キャンセル時に復元）*/
        var originalStates = captureStrokeStates(targetPaths);

        var closedByOK = false;
        var directionReversed = false;
        var isDashCleared = false;

        /* ランダムパターン（線分・間隔を3組、pt）/ Random pattern in pt */
        var randomDashesPt = null;

        var savedPrefs = loadPrefs();
        var initialValues = readInitialValues(savedPrefs, strokeUnit);

        var dialogControls = buildDashDialog(
            formatPathInfoText(primaryPathLength, targetPaths.length, strokeUnit), strokeUnit, initialValues);
        var dashDialog = dialogControls.dialog;
        var segmentsRow = dialogControls.segmentsRow;
        var txtSegments = dialogControls.txtSegments;
        var txtGap = dialogControls.txtGap;
        var lblGapResult = dialogControls.lblGapResult;
        var dashRow = dialogControls.dashRow;
        var txtDash = dialogControls.txtDash;
        var lblDashResult = dialogControls.lblDashResult;
        var rbModeGapToDash = dialogControls.rbModeGapToDash;
        var rbModeDashToGap = dialogControls.rbModeDashToGap;
        var rbModeRandom = dialogControls.rbModeRandom;
        var chkUseOffset = dialogControls.chkUseOffset;
        var txtOffset = dialogControls.txtOffset;
        var offsetPresetRow = dialogControls.offsetPresetRow;
        var rbOffsetQuarter = dialogControls.rbOffsetQuarter;
        var rbOffsetHalf = dialogControls.rbOffsetHalf;
        var rbOffsetThreeQuarter = dialogControls.rbOffsetThreeQuarter;
        var chkPartialDisplay = dialogControls.chkPartialDisplay;
        var rbCapButt = dialogControls.rbCapButt;
        var rbCapRound = dialogControls.rbCapRound;
        var rbCapProject = dialogControls.rbCapProject;
        var chkAdjustEnds = dialogControls.chkAdjustEnds;
        var chkReversePath = dialogControls.chkReversePath;
        var btnClearDash = dialogControls.btnClearDash;
        var btnCancel = dialogControls.btnCancel;
        var btnOK = dialogControls.btnOK;

        // -----------------------------------------
        // 表示用の書式 / Display formatting
        // -----------------------------------------

        /**
         * pt 値を結果表示用の文字列にする
         * @param {number} ptValue - pt 値
         * @returns {string} 小数第3位までの文字列
         */
        function ptToResultText(ptValue) {
            return ptToUnit(ptValue, strokeUnit).toFixed(3);
        }

        /**
         * pt 値を入力欄用の文字列にする（単位付き）
         * @param {number} ptValue - pt 値
         * @returns {string} 整形した文字列（例 "2.5 mm"）
         */
        function ptToFieldText(ptValue) {
            return formatFieldNumber(ptToUnit(ptValue, strokeUnit)) + strokeFieldUnit;
        }

        // -----------------------------------------
        // ランダムパターン / Random pattern
        // -----------------------------------------

        /**
         * 指定範囲の乱数（現在の線の単位）を返す
         * @param {number} minValue - 最小値
         * @param {number} maxValue - 最大値
         * @returns {number} 生成した長さ（0以上）
         */
        function getRandomLengthUnit(minValue, maxValue) {
            var randomLength = Math.random() * (maxValue - minValue) + minValue;
            if (RANDOM_ROUND_VALUES) randomLength = Math.round(randomLength);
            /* Illustrator は NaN や負値を受け付けない */
            if (isNaN(randomLength) || randomLength < 0) randomLength = 0;
            return randomLength;
        }

        /**
         * ランダムパターンを生成し直す
         * 線端が「なし」のときは、線分が消えないように単位ごとの最小値を使う。
         * @returns {void}
         */
        function recalcRandomPattern() {
            var dashMinUnit = rbCapButt.value
                ? (RANDOM_DASH_MIN_BY_UNIT[strokeUnit.code] || RANDOM_DASH_MIN_DEFAULT)
                : 0;

            randomDashesPt = [];
            for (var k = 0; k < RANDOM_PATTERN_LENGTH; k += 2) {
                randomDashesPt.push(unitToPt(getRandomLengthUnit(dashMinUnit, RANDOM_DASH_MAX), strokeUnit));
                randomDashesPt.push(unitToPt(getRandomLengthUnit(RANDOM_GAP_MIN, RANDOM_GAP_MAX), strokeUnit));
            }
        }

        /**
         * ランダムパターンを用意する（前回値があれば復元）
         * @returns {void}
         */
        function ensureRandomPattern() {
            if (randomDashesPt && randomDashesPt.length === RANDOM_PATTERN_LENGTH) return;

            var savedDashes = savedPrefs ? savedPrefs.randPt : null;
            if (savedDashes && savedDashes.length >= RANDOM_PATTERN_LENGTH) {
                randomDashesPt = savedDashes.slice(0, RANDOM_PATTERN_LENGTH);
                return;
            }
            /* v1.7以前の4要素は、最後の1組をコピーして6要素へ拡張 */
            if (savedDashes && savedDashes.length >= 4) {
                randomDashesPt = [savedDashes[0], savedDashes[1], savedDashes[2], savedDashes[3], savedDashes[0], savedDashes[1]];
                return;
            }
            recalcRandomPattern();
        }

        /**
         * 入力された間隔をランダムパターンの全ギャップへ反映する
         * @returns {void}
         */
        function applyGapToRandomPattern() {
            var gapUnit = parseFloat(txtGap.text);
            if (isNaN(gapUnit) || gapUnit < 0) gapUnit = 0;
            var gapPt = unitToPt(gapUnit, strokeUnit);

            ensureRandomPattern();
            for (var k = 1; k < RANDOM_PATTERN_LENGTH; k += 2) {
                randomDashesPt[k] = gapPt;
            }
        }

        /**
         * 新しいランダムパターンを作り、入力された間隔を反映して返す
         * @returns {Array<number>} 破線パターン（pt）
         */
        function nextRandomPattern() {
            recalcRandomPattern();
            applyGapToRandomPattern();
            return randomDashesPt.slice(0);
        }

        /**
         * ランダムパターンの線分を表示用テキストにする
         * @returns {string} 「線分1 / 線分2 / 線分3」形式の文字列
         */
        function getRandomDashText() {
            ensureRandomPattern();
            var dashTexts = [];
            for (var k = 0; k < RANDOM_PATTERN_LENGTH; k += 2) {
                dashTexts.push(ptToResultText(randomDashesPt[k]));
            }
            return dashTexts.join(" / ");
        }

        // -----------------------------------------
        // UIの状態 / UI state
        // -----------------------------------------

        /**
         * 選択中の計算方法を返す
         * @returns {number} 0:間隔→線分 / 1:線分→間隔 / 2:ランダム
         */
        function getCalcModeIndex() {
            if (rbModeRandom.value) return 2;
            return rbModeDashToGap.value ? 1 : 0;
        }

        /**
         * 選択中の線端を返す
         * @returns {number} 0:なし / 1:丸型 / 2:突出
         */
        function getCapModeIndex() {
            if (rbCapRound.value) return 1;
            return rbCapProject.value ? 2 : 0;
        }

        /**
         * 計算方法に応じて入力欄・結果表示の切り替えを行う
         * @returns {void}
         */
        function updateModeUI() {
            var isRandom = rbModeRandom.value;
            var isDashToGap = rbModeDashToGap.value;

            /* ランダムは分割数・線分を使わないため、間隔だけ入力できるようにする */
            txtDash.visible = isDashToGap && !isRandom;
            lblDashResult.visible = !txtDash.visible;
            txtGap.visible = !isDashToGap || isRandom;
            lblGapResult.visible = !txtGap.visible;
            /* 計算結果を表示している欄は入力できないので、∧∨も隠す / hide the stepper while the slot shows a result */
            txtDash.stepperGroup.visible = txtDash.visible;
            txtGap.stepperGroup.visible = txtGap.visible;

            segmentsRow.enabled = !isRandom;
            dashRow.enabled = !isRandom;
            redrawSteppersIn(segmentsRow);
            redrawSteppersIn(dashRow);
            chkPartialDisplay.enabled = !isRandom;
            chkAdjustEnds.enabled = !isRandom && !primaryPath.closed;

            if (isRandom) {
                /* 部分表示はランダムパターンと併用できない */
                chkPartialDisplay.value = false;
                applyGapToRandomPattern();
                lblDashResult.text = getRandomDashText();
                txtGap.active = true;
            }
        }

        /**
         * 開始位置（オフセット）の入力可否を切り替える
         * @returns {void}
         */
        function updateOffsetUI() {
            setStepperInputEnabled(txtOffset, chkUseOffset.value);
            offsetPresetRow.enabled = chkUseOffset.value;
        }

        // -----------------------------------------
        // 入力値の取得 / Read input values
        // -----------------------------------------

        /**
         * 入力された分割数を返す
         * @returns {number} 分割数（不正な場合は null）
         */
        function getSegmentsFromUI() {
            var segments = parseInt(txtSegments.text, 10);
            if (isNaN(segments) || segments <= 0) return null;
            return segments;
        }

        /**
         * 入力された開始位置（単位値）を返す
         * @returns {number} 開始位置。チェックOFFなら0、不正な場合は null
         */
        function getOffsetUnitFromUI() {
            if (!chkUseOffset.value) return 0;
            var offsetUnit = parseFloat(txtOffset.text);
            if (isNaN(offsetUnit) || offsetUnit < 0) return null;
            return offsetUnit;
        }

        /**
         * 現在の計算方法で使う入力欄の値が読めるかどうかを返す
         * @returns {boolean} 読める場合は true
         */
        function hasValidDashGapInput() {
            if (chkPartialDisplay.value) return true;
            var enteredValue = parseFloat(rbModeDashToGap.value ? txtDash.text : txtGap.text);
            return !isNaN(enteredValue) && enteredValue >= 0;
        }

        /**
         * 入力された線分長（pt）を返す
         * オープンパスで「両端を調整」ONかつ分割数1のときは、パス全長を線分長とする。
         * @param {number} pathLen - 対象パスの長さ（pt）
         * @param {boolean} isClosed - クローズパスかどうか
         * @param {number} segments - 分割数
         * @returns {number} 線分長（pt）。不正な場合は null
         */
        function getDashPtForPath(pathLen, isClosed, segments) {
            var dashUnit = parseFloat(txtDash.text);
            if (isNaN(dashUnit) || dashUnit < 0) return null;
            if (!isClosed && chkAdjustEnds.value && segments === 1) return pathLen;
            return unitToPt(dashUnit, strokeUnit);
        }

        /**
         * 1つのパスに適用する破線パターンを計算する
         * @param {number} pathLen - パスの長さ（pt）
         * @param {boolean} isClosed - クローズパスかどうか
         * @param {number} segments - 分割数
         * @returns {Object} { dashPt:number, gapPt:number, dashesPt:Array<number> }。計算できない場合は null
         */
        function calcDashPatternForPath(pathLen, isClosed, segments) {
            var adjustEnds = chkAdjustEnds.value;

            /* 部分表示：間隔0で計算し、線分1本だけを見せて残りは長い間隔で隠す */
            if (chkPartialDisplay.value) {
                var partialCycle = calcDashAndCyclePt(segments, 0, pathLen, isClosed, adjustEnds);
                if (!partialCycle || partialCycle.dashPt <= 0) return null;
                var partialDashPt = partialCycle.dashPt;
                return {
                    dashPt: partialDashPt,
                    gapPt: 0,
                    dashesPt: [0, 0, partialDashPt, pathLen + partialDashPt / 2]
                };
            }

            if (rbModeDashToGap.value) {
                var enteredDashPt = getDashPtForPath(pathLen, isClosed, segments);
                if (enteredDashPt == null) return null;
                var solvedGapPt = calcGapPtFromDashPt(segments, enteredDashPt, pathLen, isClosed, adjustEnds);
                if (solvedGapPt == null || solvedGapPt < 0) return null;
                return { dashPt: enteredDashPt, gapPt: solvedGapPt, dashesPt: [enteredDashPt, solvedGapPt] };
            }

            var gapUnit = parseFloat(txtGap.text);
            if (isNaN(gapUnit) || gapUnit < 0) return null;
            var enteredGapPt = unitToPt(gapUnit, strokeUnit);
            var dashCycle = calcDashAndCyclePt(segments, enteredGapPt, pathLen, isClosed, adjustEnds);
            if (!dashCycle || dashCycle.dashPt <= 0) return null;
            return { dashPt: dashCycle.dashPt, gapPt: enteredGapPt, dashesPt: [dashCycle.dashPt, enteredGapPt] };
        }

        /**
         * 1周期（線分＋間隔）の長さを単位値で返す
         * @returns {number} 1周期の長さ。計算できない場合は null
         */
        function getDashCycleUnit() {
            var segments = getSegmentsFromUI();
            if (segments == null) return null;

            if (rbModeRandom.value) {
                applyGapToRandomPattern();
                var totalPt = 0;
                for (var k = 0; k < RANDOM_PATTERN_LENGTH; k++) totalPt += randomDashesPt[k];
                return ptToUnit(totalPt, strokeUnit);
            }

            var dashPattern = calcDashPatternForPath(primaryPathLength, primaryPath.closed, segments);
            if (!dashPattern) return null;
            return ptToUnit(dashPattern.dashPt + dashPattern.gapPt, strokeUnit);
        }

        // -----------------------------------------
        // 開始位置のプリセット / Offset presets
        // -----------------------------------------

        /**
         * プリセット（1/4・1/2・3/4）の選択状態を入力値に合わせる
         * @param {number} offsetUnit - 現在の開始位置（単位値）
         * @returns {void}
         */
        function syncOffsetPreset(offsetUnit) {
            rbOffsetQuarter.value = rbOffsetHalf.value = rbOffsetThreeQuarter.value = false;

            var cycleUnit = getDashCycleUnit();
            if (cycleUnit == null) return;

            /* 単位換算で誤差が出るため、判定はゆるめにする */
            var tolerance = Math.max(0.001, Math.abs(cycleUnit) * 0.0005);

            if (Math.abs(offsetUnit - cycleUnit * 0.25) <= tolerance) rbOffsetQuarter.value = true;
            else if (Math.abs(offsetUnit - cycleUnit * 0.50) <= tolerance) rbOffsetHalf.value = true;
            else if (Math.abs(offsetUnit - cycleUnit * 0.75) <= tolerance) rbOffsetThreeQuarter.value = true;
        }

        /**
         * 1周期を基準にした開始位置プリセットを適用する
         * @param {number} fraction - 1周期に対する比率（0.25 / 0.5 / 0.75）
         * @returns {void}
         */
        function applyOffsetPreset(fraction) {
            var cycleUnit = getDashCycleUnit();
            if (cycleUnit == null) return;
            var offsetUnit = cycleUnit * fraction;
            txtOffset.text = formatFieldNumber(offsetUnit < 0 ? 0 : offsetUnit) + strokeFieldUnit;
            updatePreviewFromInput();
        }

        // -----------------------------------------
        // パスへの適用 / Apply to paths
        // -----------------------------------------

        /**
         * 対象パスすべてに処理を行う（個別の失敗は無視する）
         * @param {function} pathAction - 各パスに対して実行する処理
         * @returns {void}
         */
        function forEachTargetPath(pathAction) {
            for (var k = 0; k < targetPaths.length; k++) {
                try {
                    pathAction(targetPaths[k]);
                } catch (e) { }
            }
        }

        /**
         * 選択中の線端をパスへ設定する
         * @param {PathItem} pathItem - 対象のパス
         * @returns {void}
         */
        function applySelectedStrokeCap(pathItem) {
            if (rbCapRound.value) pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
            else if (rbCapProject.value) pathItem.strokeCap = StrokeCap.PROJECTINGENDCAP;
            else pathItem.strokeCap = StrokeCap.BUTTENDCAP;
        }

        /**
         * パスへ破線設定を適用する
         * @param {PathItem} pathItem - 対象のパス
         * @param {Array<number>} dashesPt - 破線パターン（pt）
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function applyStrokeDashes(pathItem, dashesPt, offsetPt) {
            pathItem.stroked = true;
            applySelectedStrokeCap(pathItem);
            pathItem.strokeDashOffset = offsetPt;
            pathItem.strokeDashes = dashesPt;
        }

        /**
         * 計算した破線を対象パスへ適用する（パスごとに長さを見て計算する）
         * @param {number} segments - 分割数
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function applyDashesToPaths(segments, offsetPt) {
            forEachTargetPath(function (pathItem) {
                var dashPattern = calcDashPatternForPath(pathItem.length, pathItem.closed, segments);
                if (dashPattern) applyStrokeDashes(pathItem, dashPattern.dashesPt, offsetPt);
            });
        }

        /**
         * ランダムな破線を対象パスへ適用する（パスごとに別の乱数を使う）
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function applyRandomDashesToPaths(offsetPt) {
            /* 結果表示用に代表のパターンを1回作る */
            nextRandomPattern();
            lblDashResult.text = getRandomDashText();

            forEachTargetPath(function (pathItem) {
                applyStrokeDashes(pathItem, nextRandomPattern(), offsetPt);
            });
        }

        /**
         * 対象パスの破線設定を解除する
         * @returns {void}
         */
        function clearDashesOnPaths() {
            forEachTargetPath(function (pathItem) {
                applyStrokeDashes(pathItem, [], 0);
            });
            app.redraw();
        }

        /**
         * パスの方向反転を現在の指定に合わせる
         * @param {boolean} shouldReverse - 反転させるかどうか
         * @returns {void}
         */
        function setReversePath(shouldReverse) {
            shouldReverse = !!shouldReverse;
            if (shouldReverse === directionReversed) return;

            try {
                /* パスの方向反転は選択に対して実行されるため、対象を選択してから実行する */
                doc.selection = targetPaths;
                app.executeMenuCommand('Reverse Path Direction');
                directionReversed = shouldReverse;
            } catch (e) {
                /* 失敗した場合はフラグを変更しない */
            }
        }

        /**
         * ダイアログを開く前の状態へ戻す
         * @returns {void}
         */
        function restoreOriginalState() {
            if (directionReversed) setReversePath(false);

            for (var k = 0; k < originalStates.length; k++) {
                var savedState = originalStates[k];
                try {
                    savedState.item.strokeDashes = savedState.strokeDashes.slice(0);
                    savedState.item.strokeDashOffset = savedState.strokeDashOffset;
                    savedState.item.strokeCap = savedState.strokeCap;
                    savedState.item.stroked = savedState.stroked;
                } catch (e) { }
            }
            app.redraw();
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * 線分・間隔の結果表示を更新し、隠れている入力欄も同期する
         * @param {number} segments - 分割数
         * @returns {boolean} 計算できた場合は true
         */
        function refreshDashGapDisplay(segments) {
            var resultLabel = rbModeDashToGap.value ? lblGapResult : lblDashResult;

            /* 入力途中で数値として読めないときは、結果表示を空にする */
            if (!hasValidDashGapInput()) {
                resultLabel.text = "";
                return false;
            }

            var dashPattern = calcDashPatternForPath(primaryPathLength, primaryPath.closed, segments);
            if (!dashPattern) {
                resultLabel.text = getLabel(LABELS.alert.calcError);
                return false;
            }

            lblDashResult.text = ptToResultText(dashPattern.dashPt) + strokeFieldUnit;
            lblGapResult.text = ptToResultText(dashPattern.gapPt) + strokeFieldUnit;

            /* 表示を切り替えたときにずれないよう、隠れている入力欄も同期する */
            if (chkPartialDisplay.value) {
                txtGap.text = "0" + strokeFieldUnit;
                txtDash.text = ptToFieldText(dashPattern.dashPt);
            } else if (rbModeDashToGap.value) {
                txtGap.text = ptToFieldText(dashPattern.gapPt);
                /* 全長ダッシュに置き換わった場合だけ入力欄を合わせる（入力中の値は書き換えない）*/
                if (!primaryPath.closed && chkAdjustEnds.value && segments === 1) {
                    txtDash.text = ptToFieldText(dashPattern.dashPt);
                }
            } else {
                txtDash.text = ptToFieldText(dashPattern.dashPt);
            }
            return true;
        }

        /**
         * 現在の設定でプレビューを更新する（アラートは出さない）
         * @returns {void}
         */
        function updatePreview() {
            if (isDashCleared) {
                lblDashResult.text = "";
                lblGapResult.text = "";
                clearDashesOnPaths();
                return;
            }

            var offsetUnit = getOffsetUnitFromUI();
            var segments = getSegmentsFromUI();
            if (offsetUnit == null || segments == null) {
                lblDashResult.text = "";
                lblGapResult.text = "";
                return;
            }

            var offsetPt = unitToPt(offsetUnit, strokeUnit);

            if (rbModeRandom.value) {
                applyRandomDashesToPaths(offsetPt);
                app.redraw();
                return;
            }

            if (!refreshDashGapDisplay(segments)) return;

            /* プリセット表示を手入力・分割数の変更に追従させる */
            syncOffsetPreset(offsetUnit);
            applyDashesToPaths(segments, offsetPt);
            app.redraw();
        }

        /**
         * 入力変更時にプレビューを更新する（破線クリア状態は解除する）
         * @returns {void}
         */
        function updatePreviewFromInput() {
            isDashCleared = false;
            updatePreview();
        }

        // -----------------------------------------
        // 確定処理 / Commit
        // -----------------------------------------

        /**
         * OK時の入力チェックを行い、保存用の線分・間隔（pt）を返す
         * @param {number} segments - 分割数
         * @returns {Object} { dashPt:number, gapPt:number }。エラー時は null（アラート表示済み）
         */
        function validateBeforeApply(segments) {
            if (!hasValidDashGapInput()) {
                alert(getLabel(rbModeDashToGap.value ? LABELS.alert.dashInvalid : LABELS.alert.gapInvalid));
                return null;
            }

            /* ランダムは分割数を使わないため、間隔が読めれば確定できる */
            if (rbModeRandom.value) {
                return {
                    dashPt: unitToPt(parseFloat(txtDash.text) || 0, strokeUnit),
                    gapPt: unitToPt(parseFloat(txtGap.text), strokeUnit)
                };
            }

            var dashPattern = calcDashPatternForPath(primaryPathLength, primaryPath.closed, segments);
            if (dashPattern) return dashPattern;

            /* 計算できないときは、設定できる最大値を知らせる */
            if (chkPartialDisplay.value) {
                alert(getLabel(LABELS.alert.calcError));
            } else if (rbModeDashToGap.value) {
                alert(getLabel(LABELS.alert.dashTooLong, [ptToResultText(primaryPathLength / segments), strokeUnit.label]));
            } else {
                var maxGapPt;
                if (primaryPath.closed || !chkAdjustEnds.value) maxGapPt = primaryPathLength / segments;
                else maxGapPt = (segments <= 1) ? primaryPathLength : primaryPathLength / (segments - 1);
                alert(getLabel(LABELS.alert.gapTooLong, [ptToResultText(maxGapPt), strokeUnit.label]));
                lblDashResult.text = getLabel(LABELS.alert.calcError);
            }
            return null;
        }

        /**
         * 現在のUI状態を前回値として保存する
         * @param {number} segments - 分割数
         * @param {number} dashPt - 線分長（pt）
         * @param {number} gapPt - 間隔（pt）
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function saveCurrentPrefs(segments, dashPt, gapPt, offsetPt) {
            savePrefs({
                segments: segments,
                gapPt: gapPt,
                dashPt: dashPt,
                offsetPt: offsetPt,
                capMode: getCapModeIndex(),
                mode: getCalcModeIndex(),
                reversePath: chkReversePath.value,
                adjustEnds: chkAdjustEnds.value,
                randPt: (rbModeRandom.value && randomDashesPt) ? randomDashesPt.slice(0) : null,
                useOffset: chkUseOffset.value
            });
        }

        // -----------------------------------------
        // イベント / Event handlers
        // -----------------------------------------

        txtSegments._onArrowChange = updatePreviewFromInput;
        txtGap._onArrowChange = updatePreviewFromInput;
        txtDash._onArrowChange = updatePreviewFromInput;
        txtOffset._onArrowChange = updatePreviewFromInput;

        txtSegments.onChanging = txtSegments.onChange = updatePreviewFromInput;
        txtGap.onChanging = txtGap.onChange = updatePreviewFromInput;
        txtDash.onChanging = txtDash.onChange = updatePreviewFromInput;
        txtOffset.onChanging = txtOffset.onChange = updatePreviewFromInput;

        rbCapButt.onClick = rbCapRound.onClick = rbCapProject.onClick = updatePreview;

        rbModeGapToDash.onClick = rbModeDashToGap.onClick = function () {
            updateModeUI();
            updatePreviewFromInput();
        };

        rbModeRandom.onClick = function () {
            /* クリックのたびにパターンを作り直す */
            recalcRandomPattern();
            updateModeUI();
            updatePreviewFromInput();
        };

        rbOffsetQuarter.onClick = function () { applyOffsetPreset(0.25); };
        rbOffsetHalf.onClick = function () { applyOffsetPreset(0.50); };
        rbOffsetThreeQuarter.onClick = function () { applyOffsetPreset(0.75); };

        chkUseOffset.onClick = function () {
            updateOffsetUI();
            updatePreviewFromInput();
        };

        chkPartialDisplay.onClick = function () {
            /* ONにしたら間隔を0にする */
            if (chkPartialDisplay.value) txtGap.text = "0" + strokeFieldUnit;
            updatePreviewFromInput();
        };

        chkAdjustEnds.onClick = updatePreviewFromInput;

        chkReversePath.onClick = function () {
            setReversePath(chkReversePath.value);
            updatePreview();
        };

        btnClearDash.onClick = function () {
            isDashCleared = true;
            updatePreview();
        };

        /* キャンセル：ダイアログを開く前の状態に戻して閉じる */
        btnCancel.onClick = function () {
            restoreOriginalState();
            dashDialog.close(0);
        };

        /* ×ボタンやEscで閉じた場合も、OK以外は復元する */
        dashDialog.onClose = function () {
            if (!closedByOK) restoreOriginalState();
            return true;
        };

        btnOK.onClick = function () {
            var segments = getSegmentsFromUI();

            /* 破線クリアを確定（他のUI状態は残して保存する）*/
            if (isDashCleared) {
                clearDashesOnPaths();
                saveCurrentPrefs(
                    (segments == null) ? initialValues.segments : segments,
                    unitToPt(initialValues.dashUnit, strokeUnit),
                    unitToPt(initialValues.gapUnit, strokeUnit),
                    0
                );
                closedByOK = true;
                dashDialog.close(1);
                return;
            }

            if (segments == null) {
                alert(getLabel(LABELS.alert.segmentsInvalid));
                return;
            }
            var offsetUnit = getOffsetUnitFromUI();
            if (offsetUnit == null) {
                alert(getLabel(LABELS.alert.offsetInvalid));
                return;
            }

            var appliedPattern = validateBeforeApply(segments);
            if (!appliedPattern) return;

            /* プレビューと同じ処理で対象のパスへ適用する */
            updatePreview();
            saveCurrentPrefs(segments, appliedPattern.dashPt, appliedPattern.gapPt, unitToPt(offsetUnit, strokeUnit));

            closedByOK = true;
            dashDialog.close(1);
        };

        // -----------------------------------------
        // 初期化と表示 / Initialize & show
        // -----------------------------------------

        /* 前回値（パス方向）を反映 */
        setReversePath(chkReversePath.value);

        updateOffsetUI();
        updateModeUI();
        updatePreview();

        /* ランダム以外は分割数の入力欄をアクティブにして表示する */
        if (!rbModeRandom.value) txtSegments.active = true;
        dashDialog.onShow = function () {
            if (rbModeRandom.value) txtGap.active = true;
            else txtSegments.active = true;
        };

        prepareDialogWindow(dashDialog, SCRIPT_NAME);
        dashDialog.show();
    }

    /**
     * ドキュメントと選択を確認し、ダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        /* 選択の最上位にあるパスだけを対象にする（グループ・複合パスの中は見ない） / Only top-level paths in the selection (not inside groups or compound paths) */
        var targetPaths = collectSelectionPathItems(doc.selection, { enterGroups: false, compoundPaths: "skip" });
        if (targetPaths.length === 0) {
            alert(getLabel(LABELS.alert.noPathSelected));
            return;
        }

        resetBoundsAndToggleEdges();
        try {
            showDashDialog(doc, targetPaths);
        } finally {
            resetBoundsAndToggleEdges();
        }
    }

    main();

})();
