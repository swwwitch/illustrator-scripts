#target illustrator
#targetengine "SmartVerticalAlignEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字およびエリア内文字の［字形の境界に整列］を切り替えながら、垂直方向（上・中央・下）に整列します。
チェックボックスやラジオボタンの操作はそのつどプレビューへ反映され、T / M / B キーでも整列位置を切り替えられます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartVerticalAlign.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9ee716675032

### Overview

Aligns objects vertically — top, center or bottom — while toggling Align to Glyph Bounds for point text and area text.
Every checkbox and radio button refreshes the preview, and the T, M and B keys switch the alignment position.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartVerticalAlign.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartVerticalAlign";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.11";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartVerticalAlign.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartVerticalAlign.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9ee716675032"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ［プレビュー境界］の初期状態 / Initial state of the Preview Bounds checkbox */
    var DEFAULT_USE_PREVIEW_BOUNDS = false;

    // =========================================
    // 環境設定キー / Preference keys
    // =========================================

    /* 境界にプレビュー境界（線幅・効果）を含めるか / Include stroke and effects in bounds */
    var PREF_KEY_INCLUDE_STROKE_IN_BOUNDS = 'includeStrokeInBounds';

    // =========================================
    // 整列・字形境界の定義 / Alignment and glyph-bounds tables
    // =========================================

    /* 整列位置。ラジオボタン・メニューコマンド・ショートカットキーを1か所で対応付ける
       / Alignment options: radio button, menu command and shortcut key in one table */
    var ALIGN_OPTIONS = [
        { labelKey: 'radio.top',    tooltipKey: 'tooltip.top',    menuCommand: 'Vertical Align Top',    shortcutKey: 'T' },
        { labelKey: 'radio.center', tooltipKey: 'tooltip.center', menuCommand: 'Vertical Align Center', shortcutKey: 'M' },
        { labelKey: 'radio.bottom', tooltipKey: 'tooltip.bottom', menuCommand: 'Vertical Align Bottom', shortcutKey: 'B' }
    ];
    var ALIGN_INDEX_CENTER = 1;
    var ALIGN_INDEX_BOTTOM = 2;

    /* ［字形の境界に整列］のチェックボックスと、対応する環境設定キー・文字種
       / Glyph-bounds checkboxes with their preference key and text kind */
    var GLYPH_BOUNDS_OPTIONS = [
        { labelKey: 'checkbox.pointText', tooltipKey: 'tooltip.pointText', prefKey: 'EnableActualPointTextSpaceAlign', textKind: TextType.POINTTEXT },
        { labelKey: 'checkbox.areaText',  tooltipKey: 'tooltip.areaText',  prefKey: 'EnableActualAreaTextSpaceAlign',  textKind: TextType.AREATEXT }
    ];
    var GLYPH_INDEX_POINT_TEXT = 0;

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

    var PREVIEW_ROW_MARGINS = [15, 0, 15, 0];  /* ［プレビュー境界］行の余白 */
    var BUTTON_WIDTH = 90;

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

    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

    /* ラベル定義（カテゴリ別）/ Label definitions (by category) */
    var LABELS = {
        dialog: {
            title: { ja: "垂直方向の整列", en: "Vertical Alignment" }
        },
        panel: {
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            alignment: { ja: "整列位置", en: "Alignment" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        },
        tooltip: {
            top:    { ja: "選択範囲の上端にそろえます（T キー）", en: "Align to the top of the selection (T)" },
            center: { ja: "選択範囲の上下中央にそろえます（M キー）", en: "Align to the vertical center of the selection (M)" },
            bottom: { ja: "選択範囲の下端にそろえます（B キー）", en: "Align to the bottom of the selection (B)" },
            pointText: {
                ja: "ポイント文字を、仮想ボディではなく字形の実際の輪郭でそろえます。",
                en: "Aligns point text by the actual glyph outlines instead of the em box."
            },
            areaText: {
                ja: "エリア内文字を、テキストエリアの枠ではなく字形の実際の輪郭でそろえます。",
                en: "Aligns area text by the actual glyph outlines instead of the text area frame."
            },
            previewBounds: {
                ja: "線幅や効果を含めた見た目の端を基準にそろえます（環境設定の［プレビュー境界を使用］を切り替えます）。",
                en: "Aligns by the visible edges including strokes and effects (toggles the Use Preview Bounds preference)."
            }
        },
        checkbox: {
            pointText: { ja: "ポイント文字", en: "Point Text" },
            areaText: { ja: "エリア内文字", en: "Area Text" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
        },
        radio: {
            top: { ja: "上", en: "Top" },
            center: { ja: "中央", en: "Center" },
            bottom: { ja: "下", en: "Bottom" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    // =========================================
    // 選択オブジェクトの取得 / Selection
    // =========================================

    /**
     * 整列の対象になるオブジェクトか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isAlignableItem(pageItem) {
        if (!pageItem) return false;
        return (pageItem.typename === "TextFrame" ||
            pageItem.typename === "PathItem" ||
            pageItem.typename === "GroupItem" ||
            pageItem.typename === "CompoundPathItem");
    }

    /**
     * 選択中の整列対象オブジェクトを配列で取得する
     * @returns {PageItem[]} 整列対象のオブジェクト
     */
    function getAlignableSelection() {
        var currentSelection = app.activeDocument.selection;
        var alignableItems = [];
        for (var i = 0; i < currentSelection.length; i++) {
            if (isAlignableItem(currentSelection[i])) {
                alignableItems.push(currentSelection[i]);
            }
        }
        return alignableItems;
    }

    // =========================================
    // 整列プレビュー / Alignment preview
    // =========================================

    /* ALIGN_OPTIONS と同じ並びのラジオボタン。UI構築時に埋める
       / Radio buttons in ALIGN_OPTIONS order, filled while the dialog is built */
    var alignRadios = [];

    /**
     * 指定したインデックスの整列位置だけを選択状態にする
     * @param {number} optionIndex - ALIGN_OPTIONS のインデックス
     * @returns {void}
     */
    function selectAlignOption(optionIndex) {
        for (var i = 0; i < alignRadios.length; i++) {
            alignRadios[i].value = (i === optionIndex);
        }
    }

    /**
     * 選択中の整列位置のメニューコマンドを返す
     * @returns {string|null} メニューコマンド名（未選択なら null）
     */
    function getSelectedAlignCommand() {
        for (var i = 0; i < alignRadios.length; i++) {
            if (alignRadios[i].value) return ALIGN_OPTIONS[i].menuCommand;
        }
        return null;
    }

    /**
     * 現在の設定で整列を実行し、プレビューへ即時反映する
     * @returns {void}
     */
    function applyPreviewAlignment() {
        var menuCommand = getSelectedAlignCommand();
        if (menuCommand && getAlignableSelection().length > 0) {
            /* 環境設定の変更を境界に反映させてから整列する（OFFに戻したときも効かせるため）
               / Redraw first so the new preference is reflected in the bounds, including when it is turned OFF */
            app.redraw();
            app.executeMenuCommand(menuCommand);
        }
        app.redraw();
    }

    // =========================================
    // UI構築 / Dialog
    // =========================================

    /**
     * ［字形の境界に整列］パネルを追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @param {PageItem[]} alignableItems - 選択中の整列対象オブジェクト
     * @returns {Checkbox[]} GLYPH_BOUNDS_OPTIONS と同じ並びのチェックボックス
     */
    function addGlyphBoundsPanel(targetDialog, alignableItems) {
        var glyphBoundsPanel = targetDialog.add('panel', undefined, getLabel('panel.glyphBounds'));
        setupPanel(glyphBoundsPanel, 6);

        /* 選択中の文字種。ポイント文字／エリア内文字のときだけ他方をディムする
           / Kind of the selected text; dims the checkbox for the other kind */
        var selectedTextKind = (alignableItems.length > 0) ? alignableItems[0].kind : null;

        var glyphBoundsCheckboxes = [];
        for (var i = 0; i < GLYPH_BOUNDS_OPTIONS.length; i++) {
            glyphBoundsCheckboxes.push(addGlyphBoundsCheckbox(glyphBoundsPanel, GLYPH_BOUNDS_OPTIONS[i], selectedTextKind));
        }
        return glyphBoundsCheckboxes;
    }

    /**
     * ［字形の境界に整列］のチェックボックスを1つ追加し、環境設定とプレビューに結び付ける
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {object} glyphOption - GLYPH_BOUNDS_OPTIONS の1項目
     * @param {TextType} selectedTextKind - 選択中の文字種（テキスト以外は null）
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addGlyphBoundsCheckbox(parentPanel, glyphOption, selectedTextKind) {
        var glyphBoundsCheckbox = parentPanel.add('checkbox', undefined, getLabel(glyphOption.labelKey));
        glyphBoundsCheckbox.helpTip = getLabel(glyphOption.tooltipKey);
        glyphBoundsCheckbox.value = app.preferences.getBooleanPreference(glyphOption.prefKey);

        /* 選択が別の文字種ならディム / Dim when the selection is the other text kind */
        if (selectedTextKind != null && selectedTextKind !== glyphOption.textKind) {
            glyphBoundsCheckbox.enabled = false;
        }

        /* ON/OFFで環境設定を書き換え、そのままプレビューを更新
           / Write the preference and refresh the preview on every toggle */
        glyphBoundsCheckbox.onClick = function () {
            app.preferences.setBooleanPreference(glyphOption.prefKey, glyphBoundsCheckbox.value === true);
            applyPreviewAlignment();
        };
        return glyphBoundsCheckbox;
    }

    /**
     * ［整列］パネルを追加し、ラジオボタンを alignRadios に登録する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addAlignmentPanel(targetDialog) {
        var alignmentPanel = targetDialog.add('panel', undefined, getLabel('panel.alignment'));
        setupPanel(alignmentPanel, 6);

        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            var alignRadio = alignmentPanel.add('radiobutton', undefined, getLabel(ALIGN_OPTIONS[i].labelKey));
            alignRadio.helpTip = getLabel(ALIGN_OPTIONS[i].tooltipKey);
            alignRadio.onClick = applyPreviewAlignment;
            alignRadios.push(alignRadio);
        }
    }

    /**
     * ［プレビュー境界］の行を追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addPreviewBoundsRow(targetDialog) {
        var previewBoundsRow = targetDialog.add('group');
        previewBoundsRow.orientation = 'row';
        previewBoundsRow.alignChildren = ['left', 'center'];
        previewBoundsRow.margins = PREVIEW_ROW_MARGINS;

        var previewBoundsCheckbox = previewBoundsRow.add('checkbox', undefined, getLabel('checkbox.previewBounds'));
        previewBoundsCheckbox.helpTip = getLabel('tooltip.previewBounds');
        previewBoundsCheckbox.value = DEFAULT_USE_PREVIEW_BOUNDS;
        previewBoundsCheckbox.onClick = function () {
            /* ONで線幅・効果を境界に含める / ON: include stroke and effects in the bounds */
            app.preferences.setBooleanPreference(PREF_KEY_INCLUDE_STROKE_IN_BOUNDS, previewBoundsCheckbox.value === true);
            applyPreviewAlignment();
        };
    }

    /**
     * T / M / B キーで整列位置を切り替えるショートカットの表を作る
     * @returns {Object} キー → 整列位置のラジオボタン
     */
    function buildAlignShortcutMap() {
        var shortcutMap = {};
        for (var i = 0; i < ALIGN_OPTIONS.length && i < alignRadios.length; i++) {
            shortcutMap[ALIGN_OPTIONS[i].shortcutKey] = alignRadios[i];
        }
        return shortcutMap;
    }

    /**
     * 選択内容に応じて整列位置とポイント文字チェックボックスの初期値を決める
     * @param {PageItem[]} alignableItems - 選択中の整列対象オブジェクト
     * @param {Checkbox} pointTextCheckbox - ポイント文字のチェックボックス
     * @returns {void}
     */
    function applyDefaultAlignment(alignableItems, pointTextCheckbox) {
        var hasTextFrame = false;
        var hasOtherItem = false;
        for (var i = 0; i < alignableItems.length; i++) {
            if (alignableItems[i].typename === "TextFrame") {
                hasTextFrame = true;
            } else {
                hasOtherItem = true;
            }
        }

        if (hasTextFrame && !hasOtherItem) {
            /* テキストのみ：下揃え / Text only: align bottom */
            pointTextCheckbox.value = false;
            selectAlignOption(ALIGN_INDEX_BOTTOM);
            return;
        }

        /* テキストと図形の混在、または図形のみ：中央揃え / Mixed or shapes only: align center */
        if (hasTextFrame) {
            pointTextCheckbox.value = true;
        }
        selectAlignOption(ALIGN_INDEX_CENTER);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを組み立てて表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var alignableItems = getAlignableSelection();

        var alignDialog = new Window('dialog');
        alignDialog.text = getLabel('dialog.title') + ' ' + SCRIPT_VERSION;
        setupWindow(alignDialog);

        var glyphBoundsCheckboxes = addGlyphBoundsPanel(alignDialog, alignableItems);
        addAlignmentPanel(alignDialog);
        addPreviewBoundsRow(alignDialog);
        /* ボタンエリア / Button row */
        var buttonRow = addButtonRow(alignDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnCancel.onClick = function () {
            alignDialog.close();
        };

        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnOK.preferredSize.width = BUTTON_WIDTH;
        btnOK.onClick = function () {
            alignDialog.close();
        };
        addKeyShortcuts(alignDialog, buildAlignShortcutMap());

        applyDefaultAlignment(alignableItems, glyphBoundsCheckboxes[GLYPH_INDEX_POINT_TEXT]);

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(alignDialog, SCRIPT_NAME);
        alignDialog.show();
    }

    main();

})();
