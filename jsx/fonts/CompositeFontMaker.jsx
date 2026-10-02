#target illustrator
#targetengine "CompositeFontMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

和文・かな・欧文のフォントと、かな・欧文のサイズ・ベースラインを指定して、合成フォントのファイルを作ります。
選択したテキスト（和文・かな・欧文が混じった1行でも可）から初期値を読み取ります。作った合成フォントは Illustrator の再起動後に使えます。
［InDesign にも作成］をオンにすると、起動中の InDesign にも同じ合成フォントを作ります（InDesign は再起動不要）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CompositeFontMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne0f78458ddd3

### 注意

- 合成フォント名は半角英数字と記号のみ、29文字まで（「/」「:」は不可。長すぎると Illustrator が起動しなくなるため）
- 特例文字セットは、既存の合成フォントからセット（名前と文字）を読み込んで使います。対象文字は［文字…］で編集できますが、セットを新しく作ることはできません
- バリアブルフォントは使えません（合成フォントに入れると、適用したときに Illustrator が落ちるため）。名前に「VF」「Var」が付くものは一覧に出さず、それ以外も［作成］時に判定して止めます
- InDesign の合成フォントはファイルではなく InDesign のスクリプトで作るため、InDesign を起動しておく必要があります

### Overview

Creates a composite font file from Japanese, Kana and Roman fonts and the size and baseline of Kana and Roman.
Initial values are read from the selected text (even a single line mixing Japanese, Kana and Roman). Restart Illustrator to use the new composite font.
With Also create in InDesign on, the same composite font is created in the running InDesign too (no restart needed there).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CompositeFontMaker.md

### Notes

- Composite font names are limited to 29 ASCII letters, digits and symbols (no "/" or ":"); longer names can keep Illustrator from launching
- Custom sets are loaded (name and characters) from an existing composite font; characters can be edited with Chars…, but new sets cannot be created
- Variable fonts cannot be used (a composite font containing one crashes Illustrator when applied). Names containing "VF" or "Var" are not listed; others are caught on Create
- The InDesign composite font is created through InDesign scripting, not as a file, so InDesign must be running

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CompositeFontMaker";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CompositeFontMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CompositeFontMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne0f78458ddd3"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_KANJI_FONT = "KozGoPr6N-Regular"; /* 漢字の初期フォント（PostScript名） / initial Kanji font */
    var DEFAULT_ROMAN_FONT = "MyriadPro-Regular"; /* 欧文の初期フォント（PostScript名） / initial Roman font */

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

    /* ダイアログ固有の寸法 / Dialog-specific sizes */
    var ROW_LABEL_WIDTH       = ($.locale.indexOf("ja") === 0) ? 40 : 70; /* 「和文：」などの項目名の幅（英語は長め） / row label width, wider for English */
    var FAMILY_DROPDOWN_WIDTH = 180; /* ファミリーのドロップダウン幅 / family dropdown width */
    var STYLE_DROPDOWN_WIDTH  = 100; /* スタイルのドロップダウン幅 / style dropdown width */
    var CUSTOM_LABEL_WIDTH    = 70;  /* 特例文字のパネルがあるときの項目名の幅（セット名・「読み込み：」が入る幅） / label width with the Custom Sets panel */
    var VALUE_COLUMN_WIDTH    = 95;  /* サイズ・ベースラインの列幅（∧∨＋入力欄） / size and baseline column width (stepper + field) */
    var NAME_FIELD_CHARACTERS = 26;  /* 合成フォント名（フォント名の部分）の最低幅（文字数） / minimum width of the name part, in characters */
    var NAME_LENGTH_WIDTH     = 50;  /* 名前の文字数表示の幅（「29 / 29」が入る幅） / width of the name length display */
    var NAME_ROW_LEFT_MARGIN  = 12;  /* 合成フォント名の行の左余白（1文字ほど） / left margin of the name row (about one character) */
    var CHAR_EDIT_SIZE        = [320, 120]; /* 対象文字の入力欄の大きさ / size of the character field */
    var CUSTOM_TOOLTIP_CHARS  = 40;  /* 特例文字セット名のツールチップに出す文字数 / characters shown in a custom set's tooltip */

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
    // 合成フォントのファイル形式 / Composite font file format
    // =========================================
    /* Illustrator が保存する合成フォントは、PostScript の再配置フォント（CID）に CFMA・RLBL・name テーブルを添えた
       sfnt 風のコンテナ。チェックサムは全テーブル固定値。標準6文字セットの範囲と RLBL は実ファイルから写した
       Illustrator saves a composite font as an sfnt-like container: a PostScript rearranged font (CID) plus
       CFMA / RLBL / name tables. Checksums are a fixed value. Ranges and RLBL are copied from real files */

    /* 文字セットごとの Unicode 範囲（0 の漢字は既定フォントなので範囲なし） / Unicode ranges per set (0 = Kanji, the base font) */
    var CHARSET_RANGES = [
        null,
        "3041-3093,309d-309e,30a1-30f6,30fc-30fe",
        "2014-2016,2018-2019,201c-201d,2025-2026,2032-2033,2225,3001-3002,3008-3011,3014-3015,301c,30fb,ff01,ff08-ff09,ff0c,ff0e-ff0f,ff1a-ff1b,ff1f,ff3b,ff3d,ff5b,ff5d-ff5e",
        "00a2-00a3,00a7-00a8,00ac,00b0-00b1,00b4,00b6,00d7,00f7,0391-03a1,03a3-03a9,03b1-03c1,03c3-03c9,0401,0410-044f,0451,2010,2020-2021,2030,203b,2103,212b,2190-2193,21d2,21d4,2200,2202-2203,2207-2208,220b,2212,221a,221d-221e,2220,2227-222c,2234-2235,223d,2252,2260-2261,2266-2267,226a-226b,2282-2283,2286-2287,22a5,2312,25a0-25a1,25b2-25b3,25bc-25bd,25c6-25c7,25cb,25ce-25cf,25ef,2605-2606,2640,2642,266a,266d,266f,3003,3006-3007,3012-3013,309b-309c,ff03-ff06,ff0a-ff0b,ff0d,ff10-ff19,ff1c-ff1e,ff20-ff3a,ff3c,ff3e-ff5a,ff5c,ffe3,ffe5",
        "0020-002f,003a-007f,00a0-00a1,00a4-00a6,00a9-00ab,00ad-00af,00b2-00b3,00b5,00b7-00d6,00d8-00f6,00f8-00ff,0131,0152-0153,0160-0161,0178,017d-017e,0192,02c6-02c7,02d8-02dd,2013,201a,201e,2022,2039-203a,2044,20ac,2122,2206,220f,2211,2248,2264-2265,25ca,fb00-fb06",
        "0030-0039"
    ];
    var RLBL_HEX = "0001000000000000000603900003000000010000000500000003000400050000ffff000000030000000400000004000900030004000d304130933041309d309e309d30a130f630a130fc30fe30fc0003000000150000000b001100030008001c201420162014201820192018201c201d201c202520262025203220332032222522252225300130023001300830113008301430153014301c301c301c30fb30fb30fbff01ff01ff01ff08ff09ff08ff0cff0cff0cff0eff0fff0eff1aff1bff1aff1fff1fff1fff3bff3bff3bff3dff3dff3dff5bff5bff5bff5dff5eff5d00030000004700000007002400030008002b00a200a300a200a700a800a700ac00ac00ac00b000b100b000b400b400b400b600b600b600d700d700d700f700f700f7039103a1039103a303a903a303b103c103b103c303c903c30401040104010410044f0410045104510451201020102010202020212020203020302030203b203b203b210321032103212b212b212b21902193219021d221d221d221d421d421d4220022002200220222032202220722082207220b220b220b221222122212221a221a221a221d221e221d2220222022202227222c2227223422352234223d223d223d225222522252226022612260226622672266226a226b226a22822283228222862287228622a522a522a523122312231225a025a125a025b225b325b225bc25bd25bc25c625c725c625cb25cb25cb25ce25cf25ce25ef25ef25ef260526062605264026402640264226422642266a266a266a266d266d266d266f266f266f300330033003300630073006301230133012309b309c309bff03ff06ff03ff0aff0bff0aff0dff0dff0dff10ff19ff10ff1cff1eff1cff20ff3aff20ff3cff3cff3cff3eff5aff3eff5cff5cff5cffe3ffe3ffe3ffe5ffe5ffe50001000000220000000a003300030008003d0020002f0020003a007f003a00a000a100a000a400a600a400a900ab00a900ad00af00ad00b200b300b200b500b500b500b700d600b700d800f600d800f800ff00f8013101310131015201530152016001610160017801780178017d017e017d01920192019202c602c702c602d802dd02d8201320132013201a201a201a201e201e201e2022202220222039203a203920442044204420ac20ac20ac212221222122220622062206220f220f220f22112211221122482248224822642265226425ca25ca25cafb00fb06fb0000010000000100000007004500030008004c0030003900304b616e6a698abf8e9a4b616e6182a982c850756e6374756174696f6e91538a7096f195a853796d626f6c7391538a708b4c8d86416c706861626574696394bc8a7089a295b64e756d6265727394bc8a7090948e9a00";
    var TABLE_CHECKSUM = 0xaf64c0a8; /* 全テーブル共通の固定値 / fixed value shared by every table */
    var CMAP_SUFFIX = "-UniJIS-UTF16-H";
    var CHARSET_COUNT = 6;
    /* 名前の上限。43文字（内部名90文字）では Illustrator が起動しなくなり、29文字（内部名62文字）は起動を確認済み
       Name limit: 43 characters kept Illustrator from launching; 29 (internal name 62) is confirmed to work */
    var MAX_NAME_LENGTH = 29;

    /**
     * 16進文字列をバイナリ文字列にする
     * @param {string} hex - 16進文字列
     * @returns {string} 1文字1バイトの文字列
     */
    function hexToBinary(hex) {
        var binary = "";
        for (var i = 0; i < hex.length; i += 2) binary += String.fromCharCode(parseInt(hex.substr(i, 2), 16));
        return binary;
    }

    /**
     * 16ビット符号なし整数をビッグエンディアンの2バイトにする
     * @param {number} value - 値
     * @returns {string} 2バイトの文字列
     */
    function uint16(value) {
        return String.fromCharCode((value >> 8) & 255, value & 255);
    }

    /**
     * 32ビット符号なし整数をビッグエンディアンの4バイトにする
     * @param {number} value - 値
     * @returns {string} 4バイトの文字列
     */
    function uint32(value) {
        return uint16((value >>> 16) & 65535) + uint16(value & 65535);
    }

    /**
     * 16.16 固定小数点の4バイトにする
     * @param {number} value - 値
     * @returns {string} 4バイトの文字列
     */
    function fixed1616(value) {
        return uint32(Math.round(value * 65536) >>> 0);
    }

    /**
     * 小数6桁の文字列にする（-0 は 0 にそろえる）
     * @param {number} value - 値
     * @returns {string} "1.000000" の形の文字列
     */
    function formatDecimal6(value) {
        var text = value.toFixed(6);
        return (text === "-0.000000") ? "0.000000" : text;
    }

    /**
     * 合成フォント名から内部名（ATC-＋名前の16進）を作る
     * @param {string} fontName - 合成フォント名（ASCII）
     * @returns {string} "ATC-73772d42" の形の内部名
     */
    function getInternalName(fontName) {
        var hexName = "";
        for (var i = 0; i < fontName.length; i++) hexName += ("0" + fontName.charCodeAt(i).toString(16)).slice(-2);
        return "ATC-" + hexName;
    }

    /**
     * 文字セットの範囲を [下限, 上限, 割り当て先] の配列で返す
     * @param {Object} fontSpec - 合成フォントの設定
     * @param {number} setIndex - 文字セットの番号（6以降は特例文字）
     * @returns {number[][]} 範囲の配列
     */
    function getCharsetRanges(fontSpec, setIndex) {
        if (setIndex >= CHARSET_COUNT) return fontSpec.customSets[setIndex - CHARSET_COUNT].ranges;
        return parseRangeText(CHARSET_RANGES[setIndex]);
    }

    /**
     * 範囲文字列（"3041-3093,309d" の形）を [下限, 上限, 割り当て先] の配列にする
     * @param {string} rangeText - 範囲文字列
     * @returns {number[][]} 範囲の配列
     */
    function parseRangeText(rangeText) {
        var ranges = [];
        var rangeTexts = rangeText.split(",");
        for (var i = 0; i < rangeTexts.length; i++) {
            var bounds = rangeTexts[i].split("-");
            var lowCode = parseInt(bounds[0], 16);
            ranges.push([lowCode, parseInt(bounds[1] || bounds[0], 16), lowCode]);
        }
        return ranges;
    }

    /**
     * 文字コードを4桁の16進にする
     * @param {number} code - 文字コード
     * @returns {string} "3041" の形の文字列
     */
    function toHex4(code) {
        return ("000" + code.toString(16)).slice(-4);
    }

    /**
     * 再配置フォントの PostScript 本文を作る
     * @param {string} internalName - 内部名
     * @param {Object} fontSpec - 合成フォントの設定
     * @returns {{text: string, startPos: number}} 本文と %ADOStartRearrangedFont の位置
     */
    function buildPostScriptBody(internalName, fontSpec) {
        var CR = "\r";
        var setCount = fontSpec.fonts.length;
        var cmapNames = [];
        var i, k;
        for (i = 0; i < setCount; i++) cmapNames.push(fontSpec.fonts[i] + CMAP_SUFFIX);

        var psText = "%!PS-Adobe-3.0 Resource-Font" + CR +
            "%ADOResourceSubCategory: RearrangedFont" + CR +
            "%%DocumentNeededResources: ProcSet CIDInit" + CR;
        for (i = 0; i < setCount; i++) psText += "%%+ Font " + cmapNames[i] + CR;
        psText += "%%IncludeResource: ProcSet CIDInit" + CR;
        for (i = 0; i < setCount; i++) psText += "%%IncludeResource: font " + cmapNames[i] + CR;
        psText += "%%BeginResource: Font " + internalName + CR + "%%Version: 1" + CR +
            "/CIDInit /ProcSet findresource begin" + CR;

        var startPos = psText.length;
        psText += "%ADOStartRearrangedFont" + CR + "/" + internalName + CR + "[" + CR;
        for (i = 0; i < setCount; i++) psText += "/" + cmapNames[i] + CR;
        psText += "] beginrearrangedfont" + CR;

        /* 漢字（0）は既定フォントなので、かな（1）以降に行列と範囲を書く。特例文字は標準の後ろに並べる
           Kanji (0) is the base font; custom sets follow the standard ones */
        for (i = 1; i < setCount; i++) {
            var scaleX = fontSpec.size[i] * fontSpec.hScale[i] / 10000;
            var scaleY = fontSpec.size[i] * fontSpec.vScale[i] / 10000;
            var shiftY = fontSpec.baseline[i] / 100;
            psText += i + " beginusematrix [" + formatDecimal6(scaleX) + " 0 0 " + formatDecimal6(scaleY) +
                " 0 " + formatDecimal6(shiftY) + "] endusematrix" + CR + i + " usefont" + CR;
            var ranges = getCharsetRanges(fontSpec, i);
            psText += ranges.length + " beginbfrange" + CR;
            for (k = 0; k < ranges.length; k++) {
                psText += "<" + toHex4(ranges[k][0]) + "> <" + toHex4(ranges[k][1]) + "> <" + toHex4(ranges[k][2]) + "> " + CR;
            }
            psText += "endbfrange" + CR;
        }
        psText += "endrearrangedfont" + CR + "end" + CR + "%%EndResource" + CR + "%%EOF" + CR;
        return { text: psText, startPos: startPos };
    }

    /**
     * name テーブルを作る
     * @param {string} internalName - 内部名
     * @param {string} fontName - 合成フォント名
     * @returns {string} テーブルのバイナリ
     */
    function buildNameTable(internalName, fontName) {
        var internalLength = internalName.length;
        /* format 0・4レコード・文字列は 6＋12×4＝54 バイト目から / format 0, 4 records, strings at offset 54 */
        return uint16(0) + uint16(4) + uint16(54) +
            uint16(1) + uint16(0) + uint16(0) + uint16(1) + uint16(internalLength) + uint16(0) +
            uint16(1) + uint16(1) + uint16(11) + uint16(2) + uint16(0) + uint16(internalLength) +
            uint16(1) + uint16(0) + uint16(0) + uint16(6) + uint16(internalLength) + uint16(internalLength) +
            uint16(1) + uint16(1) + uint16(11) + uint16(1) + uint16(fontName.length) + uint16(internalLength * 2) +
            internalName + internalName + fontName;
    }

    /**
     * RLBL テーブル（文字セットの名前と範囲）を作る
     * 標準6セットは実ファイルから写したものを使い、特例文字を後ろに足す
     * @param {Object} fontSpec - 合成フォントの設定
     * @returns {string} テーブルのバイナリ
     */
    function buildRlblTable(fontSpec) {
        var standardRlbl = hexToBinary(RLBL_HEX);
        var standardNamesOffset = (standardRlbl.charCodeAt(10) << 8) | standardRlbl.charCodeAt(11);
        var records = standardRlbl.substring(12, standardNamesOffset);
        var names = standardRlbl.substring(standardNamesOffset, standardRlbl.length - 1); /* 末尾の 0 を除く / drop the trailing 0 */
        for (var i = 0; i < fontSpec.customSets.length; i++) {
            var customSet = fontSpec.customSets[i];
            var nameOffset = names.length;
            names += customSet.nameBinary;
            /* 種別0・特例1・範囲数・0・英語名なし（長さ0、位置は日本語名と同じ）・3・日本語名の長さと位置
               kind 0, custom 1, range count, 0, no English name, 3, localized name length and offset */
            records += uint16(0) + uint16(1) + uint16(customSet.ranges.length) + uint16(0) +
                uint16(0) + uint16(nameOffset) + uint16(3) + uint16(customSet.nameBinary.length) + uint16(nameOffset);
            for (var k = 0; k < customSet.ranges.length; k++) {
                records += uint16(customSet.ranges[k][0]) + uint16(customSet.ranges[k][1]) + uint16(customSet.ranges[k][2]);
            }
        }
        var setCount = CHARSET_COUNT + fontSpec.customSets.length;
        return uint16(1) + uint16(0) + uint16(0) + uint16(0) + uint16(setCount) + uint16(12 + records.length) +
            records + names + String.fromCharCode(0);
    }

    /**
     * 合成フォントファイルのバイナリを作る
     * @param {Object} fontSpec - name / fonts（PostScript名）/ size・hScale・vScale・baseline（％）/ customSets（特例文字）
     * @returns {string} ファイル全体のバイナリ文字列
     */
    function buildCompositeFontBinary(fontSpec) {
        if (!fontSpec.customSets) fontSpec.customSets = [];
        var setCount = fontSpec.fonts.length;
        var internalName = getInternalName(fontSpec.name);
        var psBody = buildPostScriptBody(internalName, fontSpec);
        var i, k;

        var cidTable = uint16(1) + uint16(0) + uint16(1) + uint16(0x0247) +
            uint32(psBody.text.length) + uint32(psBody.text.length - psBody.startPos - 6) +
            uint16(0) + uint16(0) + uint16(0) + psBody.text;

        /* フラグは文字セット数×4語。値の意味は不明だが、実ファイルは6セットなら先頭16語、それ以上なら先頭8語が1
           Flags: 4 words per set; meaning unknown, real files set the first 16 words (6 sets) or 8 words (more sets) to 1 */
        var cfmaTable = uint16(1) + uint16(1) + uint16(1) + uint16(0) + uint16(setCount);
        var flagOnes = (setCount === CHARSET_COUNT) ? 16 : 8;
        for (i = 0; i < setCount * 4; i++) cfmaTable += uint16(i < flagOnes ? 1 : 0);
        /* サイズ・水平比率・垂直比率を文字セット順に16.16固定小数点で / size, h-scale, v-scale as 16.16 fixed */
        var scaleLists = [fontSpec.size, fontSpec.hScale, fontSpec.vScale];
        for (k = 0; k < scaleLists.length; k++) {
            for (i = 0; i < setCount; i++) cfmaTable += fixed1616(scaleLists[k][i]);
        }

        var rlblTable = buildRlblTable(fontSpec);
        var nameTable = buildNameTable(internalName, fontSpec.name);

        /* テーブル一覧はタグ順、本体は CID から並べる / directory in tag order, data starts with CID */
        var cidOffset = 12 + 16 * 4;
        var cfmaOffset = cidOffset + cidTable.length;
        var rlblOffset = cfmaOffset + cfmaTable.length;
        var nameOffset = rlblOffset + rlblTable.length;
        var header = "typ1" + uint16(4) + uint16(0x80) + uint16(3) + uint16(0x70) +
            "CFMA" + uint32(TABLE_CHECKSUM) + uint32(cfmaOffset) + uint32(cfmaTable.length) +
            "CID " + uint32(TABLE_CHECKSUM) + uint32(cidOffset) + uint32(cidTable.length) +
            "RLBL" + uint32(TABLE_CHECKSUM) + uint32(rlblOffset) + uint32(rlblTable.length) +
            "name" + uint32(TABLE_CHECKSUM) + uint32(nameOffset) + uint32(nameTable.length);
        return header + cidTable + cfmaTable + rlblTable + nameTable;
    }

    // =========================================
    // 保存先 / Save location
    // =========================================
    /**
     * Illustrator の合成フォントフォルダーを探す
     * @returns {Folder|null} 見つかったフォルダー。なければ null
     */
    function findCompositeFontFolder() {
        var majorVersion = parseInt(app.version, 10);
        var localeFolderPath = Folder.userData.fsName + "/Adobe/Adobe Illustrator " + majorVersion + "/" + app.locale;
        var folderNames = ["合成フォント", "Composite Fonts"];
        for (var i = 0; i < folderNames.length; i++) {
            var candidate = new Folder(localeFolderPath + "/" + folderNames[i]);
            if (candidate.exists) return candidate;
        }
        return null;
    }

    /**
     * 合成フォントのファイルを書き出す
     * @param {Folder} targetFolder - 保存先
     * @param {Object} fontSpec - 合成フォントの設定
     * @returns {boolean} 書き出したら true（上書きを取りやめたら false）
     */
    function writeCompositeFontFile(targetFolder, fontSpec) {
        var fontFile = new File(targetFolder.fsName + "/" + fontSpec.name);
        if (fontFile.exists && !confirm(getLabel("alert.overwrite").replace("%1", fontSpec.name))) return false;
        fontFile.encoding = "BINARY";
        if (!fontFile.open("w")) throw new Error(getLabel("alert.writeFailed") + fontFile.fsName);
        fontFile.write(buildCompositeFontBinary(fontSpec));
        fontFile.close();
        return true;
    }

    // =========================================
    // InDesign の合成フォント / InDesign composite font
    // =========================================
    /* InDesign の合成フォントは InDesign Defaults が持ち、CompositeFont フォルダーはその書き出し先。
       フォルダーにファイルを置いても取り込まれず、起動時に削除される（2026-09-28 実測）ので、BridgeTalk で DOM から作る
       InDesign keeps composite fonts in its defaults and only exports them to the CompositeFont folder;
       files placed there are deleted on launch, so the font is created through the DOM via BridgeTalk */
    /* 次の関数は toString() で InDesign に送る。JSDoc を付けると受け側の eval が失敗し、日本語は toString() で化けて
       コメントの終わりを壊すので、中には ASCII しか書かない。改行も落ちるので // コメントも使わない。
       直後にコメントがあると toString() がそこまで取り込むので、関数のすぐ後ろには文（INDESIGN_TIMEOUT）を置く。
       InDesign 側で合成フォントを作る。同名があれば、上書き指定のときだけ標準6セットを書き換え、特例文字を作り直す。
       標準6セットは add() の時点であるので書き換え、特例文字は足す。サイズ・ベースラインは％のまま渡す。漢字（0）は基準なので触らない。
       戻り値は "OK" / "EXISTS" / "MISSING|PS名,…" / "ERROR|メッセージ"（改行・タブは転送で化けるので使わない）
       The worker is sent with toString(): no JSDoc, and ASCII only inside (non-ASCII is garbled by toString()) */
    function inDesignCompositeFontWorker(idSpec) {
        try {
            var psNames = app.fonts.everyItem().postscriptName;
            var fontObjects = [];
            var missingNames = [];
            for (var i = 0; i < idSpec.fonts.length; i++) {
                var foundFont = null;
                for (var k = 0; k < psNames.length; k++) {
                    if (psNames[k] === idSpec.fonts[i]) {
                        foundFont = app.fonts[k];
                        break;
                    }
                }
                if (!foundFont) missingNames.push(idSpec.fonts[i]);
                fontObjects.push(foundFont);
            }
            if (missingNames.length > 0) return "MISSING|" + missingNames.join(",");

            var compositeFont = app.compositeFonts.itemByName(idSpec.name);
            if (compositeFont.isValid) {
                if (!idSpec.overwrite) return "EXISTS";
                for (var entryIndex = compositeFont.compositeFontEntries.length - 1; entryIndex >= idSpec.standardCount; entryIndex--) {
                    compositeFont.compositeFontEntries[entryIndex].remove();
                }
            } else {
                compositeFont = app.compositeFonts.add({ name: idSpec.name });
            }
            var fontEntries = compositeFont.compositeFontEntries;
            for (var setIndex = 0; setIndex < idSpec.fonts.length; setIndex++) {
                var fontEntry;
                if (setIndex < idSpec.standardCount) {
                    fontEntry = fontEntries[setIndex];
                } else {
                    fontEntry = fontEntries.add({ name: idSpec.customSets[setIndex - idSpec.standardCount].name });
                    fontEntry.customCharacters = idSpec.customSets[setIndex - idSpec.standardCount].chars;
                }
                fontEntry.appliedFont = fontObjects[setIndex];
                if (setIndex === 0) continue;
                fontEntry.relativeSize = idSpec.size[setIndex];
                fontEntry.baselineShift = idSpec.baseline[setIndex];
            }
            return "OK";
        } catch (e) {
            return "ERROR|" + e.message;
        }
    }
    var INDESIGN_TIMEOUT = 60; /* 応答を待つ秒数 / seconds to wait for InDesign */

    /**
     * InDesign がインストールされていれば BridgeTalk の宛先を返す
     * @returns {string|null} 宛先（例 "indesign-21.064"）。無ければ null
     */
    function getInDesignSpecifier() {
        return BridgeTalk.getSpecifier("indesign") || null;
    }

    /**
     * InDesign に合成フォントを作らせ、結果を待つ
     * @param {string} specifier - BridgeTalk の宛先
     * @param {Object} idSpec - inDesignCompositeFontWorker() に渡す設定
     * @returns {string} ワーカーの戻り値（応答が無ければ "ERROR|…"）
     */
    function sendToInDesign(specifier, idSpec) {
        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = specifier;
        /* 転送で「\」が二重になり文字列が壊れるので、設定は URI エンコードして英数字と記号だけで渡す
           The transfer doubles backslashes, so the settings travel URI-encoded */
        bridgeMessage.body = "(" + inDesignCompositeFontWorker.toString() + ")(eval(decodeURIComponent(\"" +
            encodeURIComponent(idSpec.toSource()) + "\")));";
        var resultText = null;
        bridgeMessage.onResult = function (reply) {
            resultText = String(reply.body);
        };
        bridgeMessage.onError = function (reply) {
            resultText = "ERROR|" + reply.body;
        };
        bridgeMessage.send(INDESIGN_TIMEOUT);
        /* send() が待たずに戻った場合に備えて、応答が届くまで回す / pump in case send() returned before the reply */
        var waitStart = new Date().getTime();
        while (resultText === null && new Date().getTime() - waitStart < INDESIGN_TIMEOUT * 1000) {
            BridgeTalk.pump();
            $.sleep(100);
        }
        return (resultText === null) ? "ERROR|" + getLabel("alert.inDesignTimeout") : resultText;
    }

    /**
     * InDesign に同じ設定の合成フォントを作る
     * @param {Object} fontSpec - 合成フォントの設定
     * @returns {string} 結果の説明（完了・取りやめ・失敗）
     */
    function createInDesignCompositeFont(fontSpec) {
        var specifier = getInDesignSpecifier();
        if (!specifier) return getLabel("alert.inDesignMissing");
        if (!BridgeTalk.isRunning(specifier)) return getLabel("alert.inDesignNotRunning");

        var customSets = [];
        for (var i = 0; i < fontSpec.customSets.length; i++) {
            customSets.push({ name: fontSpec.customSets[i].displayName, chars: rangesToText(fontSpec.customSets[i].ranges) });
        }
        var idSpec = {
            name: fontSpec.name, fonts: fontSpec.fonts, size: fontSpec.size, baseline: fontSpec.baseline,
            customSets: customSets, standardCount: CHARSET_COUNT, overwrite: false
        };
        var resultText = sendToInDesign(specifier, idSpec);
        if (resultText === "EXISTS") {
            if (!confirm(getLabel("alert.inDesignOverwrite").replace("%1", fontSpec.name))) return getLabel("alert.inDesignSkipped");
            idSpec.overwrite = true;
            resultText = sendToInDesign(specifier, idSpec);
        }
        if (resultText === "OK") return getLabel("alert.inDesignDone");
        var resultParts = resultText.split("|");
        if (resultParts[0] === "MISSING") return getLabel("alert.inDesignFontMissing") + resultParts.slice(1).join("|");
        return getLabel("alert.inDesignFailed") + resultParts.slice(1).join("|");
    }

    // =========================================
    // 特例文字の読み込み / Loading custom sets
    // =========================================
    /**
     * バイナリ文字列の位置から16ビット符号なし整数を読む（ビッグエンディアン）
     * @param {string} binary - バイナリ文字列
     * @param {number} offset - 位置
     * @returns {number} 値
     */
    function readUint16(binary, offset) {
        return (binary.charCodeAt(offset) << 8) | binary.charCodeAt(offset + 1);
    }

    /**
     * バイナリ文字列の位置から32ビット符号なし整数を読む（ビッグエンディアン）
     * @param {string} binary - バイナリ文字列
     * @param {number} offset - 位置
     * @returns {number} 値
     */
    function readUint32(binary, offset) {
        return readUint16(binary, offset) * 65536 + readUint16(binary, offset + 2);
    }

    /**
     * Shift-JIS のバイト列を文字列に直す（一時ファイルを Shift_JIS で読み直す）
     * @param {string} binary - Shift-JIS のバイト列
     * @param {string} fallbackText - 直せなかったときの文字列
     * @returns {string} 文字列
     */
    function decodeShiftJIS(binary, fallbackText) {
        var tempFile = new File(Folder.temp.fsName + "/" + SCRIPT_NAME + "-name.txt");
        var decodedText = "";
        try {
            tempFile.encoding = "BINARY";
            tempFile.open("w");
            tempFile.write(binary);
            tempFile.close();
            tempFile.encoding = "Shift_JIS";
            tempFile.open("r");
            decodedText = tempFile.read();
            tempFile.close();
            tempFile.remove();
        } catch (e) {}
        return decodedText || fallbackText;
    }

    /**
     * 合成フォントのファイルから特例文字のセットを読み出す
     * @param {File} fontFile - 合成フォントのファイル
     * @returns {Object[]} { nameBinary, displayName, ranges } の配列（特例文字が無ければ空）
     */
    function readCustomCharsets(fontFile) {
        var customSets = [];
        fontFile.encoding = "BINARY";
        if (!fontFile.open("r")) return customSets; /* 開けないときは例外ではなく false / open() returns false on failure */
        var fontBinary = fontFile.read();
        fontFile.close();
        if (fontBinary.substr(0, 4) !== "typ1") return customSets;

        var rlblOffset = -1;
        for (var tableIndex = 0; tableIndex < readUint16(fontBinary, 4); tableIndex++) {
            var entryOffset = 12 + 16 * tableIndex;
            if (fontBinary.substr(entryOffset, 4) === "RLBL") rlblOffset = readUint32(fontBinary, entryOffset + 8);
        }
        if (rlblOffset < 0) return customSets;

        /* 見出し12バイト（1・0・0・0・セット数・名前の位置）のあと、セットごとに9語＋範囲×3語
           12-byte header (1, 0, 0, 0, set count, names offset), then 9 words plus 3 words per range for each set */
        var setCount = readUint16(fontBinary, rlblOffset + 8);
        var namesStart = rlblOffset + readUint16(fontBinary, rlblOffset + 10);
        var recordPos = rlblOffset + 12;
        for (var i = 0; i < setCount; i++) {
            var isCustom = readUint16(fontBinary, recordPos + 2) === 1;
            var rangeCount = readUint16(fontBinary, recordPos + 4);
            var nameLength = readUint16(fontBinary, recordPos + 14);
            var nameOffset = readUint16(fontBinary, recordPos + 16);
            recordPos += 18;
            var ranges = [];
            for (var k = 0; k < rangeCount; k++) {
                ranges.push([readUint16(fontBinary, recordPos), readUint16(fontBinary, recordPos + 2), readUint16(fontBinary, recordPos + 4)]);
                recordPos += 6;
            }
            if (!isCustom) continue;
            var nameBinary = fontBinary.substr(namesStart + nameOffset, nameLength);
            customSets.push({
                nameBinary: nameBinary,
                displayName: decodeShiftJIS(nameBinary, getLabel("fallbackName.customSet") + (customSets.length + 1)),
                ranges: ranges
            });
        }
        return customSets;
    }

    /**
     * 合成フォントフォルダーから、特例文字を持つ合成フォントを集める
     * @param {Folder|null} fontFolder - 合成フォントのフォルダー
     * @returns {Object[]} { label, customSets } の配列（名前順）
     */
    function collectCustomSetSources(fontFolder) {
        var customSetSources = [];
        if (!fontFolder) return customSetSources;
        var fontFiles = fontFolder.getFiles(function (item) {
            return (item instanceof File) && item.name.charAt(0) !== ".";
        });
        for (var i = 0; i < fontFiles.length; i++) {
            var customSets = readCustomCharsets(fontFiles[i]);
            if (customSets.length === 0) continue;
            /* displayName は空で返ることがあるので、name（URIエンコード）をデコードして使う
               displayName can come back empty, so decode the URI-encoded name instead */
            var fileLabel = "";
            /* 不正なエスケープは URIError になる / malformed escapes raise URIError */
            try {
                fileLabel = String(decodeURI(fontFiles[i].name));
            } catch (e) {
                fileLabel = String(fontFiles[i].name);
            }
            customSetSources.push({ label: fileLabel || String(customSetSources.length + 1), customSets: customSets });
        }
        customSetSources.sort(function (firstSource, secondSource) {
            return (firstSource.label < secondSource.label) ? -1 : (firstSource.label > secondSource.label) ? 1 : 0;
        });
        return customSetSources;
    }

    /**
     * 文字コードが範囲の配列に含まれるか
     * @param {number} code - 文字コード
     * @param {number[][]} ranges - parseRangeText() の結果
     * @returns {boolean} 含まれれば true
     */
    function isCodeInRanges(code, ranges) {
        for (var i = 0; i < ranges.length; i++) {
            if (code >= ranges[i][0] && code <= ranges[i][1]) return true;
        }
        return false;
    }

    /**
     * 特例文字のセットが、和文・かな・欧文のどれに近いかを返す（初期フォントの決定用）
     * すべて半角欧文・半角数字なら欧文、すべてかなならかな、それ以外は和文
     * @param {number[][]} ranges - セットの範囲
     * @returns {string} "kanji" / "kana" / "roman"
     */
    function classifyCustomSet(ranges) {
        var romanRanges = parseRangeText(CHARSET_RANGES[4] + "," + CHARSET_RANGES[5]);
        var kanaRanges = parseRangeText(CHARSET_RANGES[1]);
        var allRoman = true;
        var allKana = true;
        for (var i = 0; i < ranges.length; i++) {
            for (var code = ranges[i][0]; code <= ranges[i][1]; code++) {
                if (!isCodeInRanges(code, romanRanges)) allRoman = false;
                if (!isCodeInRanges(code, kanaRanges)) allKana = false;
                if (!allRoman && !allKana) return "kanji";
            }
        }
        return allRoman ? "roman" : (allKana ? "kana" : "kanji");
    }

    /**
     * 特例文字のセットの文字を並べる（ツールチップ用、CUSTOM_TOOLTIP_CHARS 文字まで）
     * @param {number[][]} ranges - セットの範囲
     * @returns {string} 文字の並び
     */
    function describeCustomSet(ranges) {
        var charList = rangesToText(ranges);
        return (charList.length > CUSTOM_TOOLTIP_CHARS) ? charList.substr(0, CUSTOM_TOOLTIP_CHARS) + "..." : charList;
    }

    /**
     * 特例文字セットの文字をすべて並べる（編集用）
     * @param {number[][]} ranges - セットの範囲
     * @returns {string} 文字の並び
     */
    function rangesToText(ranges) {
        var charList = "";
        for (var i = 0; i < ranges.length; i++) {
            for (var code = ranges[i][0]; code <= ranges[i][1]; code++) charList += String.fromCharCode(code);
        }
        return charList;
    }

    /**
     * 入力された文字から特例文字セットの範囲を作る
     * Illustrator の特例文字と同じく、1文字ずつ入力順に並べる（重複・空白・改行は除く）
     * Like Illustrator's own custom sets, one range per character in input order (duplicates and whitespace dropped)
     * @param {string} charText - 入力された文字
     * @returns {number[][]} 範囲の配列
     */
    function textToRanges(charText) {
        var ranges = [];
        var seenCodes = {};
        for (var i = 0; i < charText.length; i++) {
            var code = charText.charCodeAt(i);
            if (code >= 0xd800 && code <= 0xdfff) continue; /* サロゲートペア（U+10000以降）は扱えない / no characters beyond the BMP */
            if (/\s/.test(charText.charAt(i)) || seenCodes[code]) continue;
            seenCodes[code] = true;
            ranges.push([code, code, code]);
        }
        return ranges;
    }

    /**
     * 特例文字セットの対象文字を表示・編集するダイアログ
     * @param {Object} customSet - 特例文字セット（ranges を書き換える）
     * @returns {boolean} 書き換えたら true
     */
    function editCustomSetCharacters(customSet) {
        var charDialog = new Window("dialog", getLabel("dialog.customChars").replace("%1", customSet.displayName));
        setupWindow(charDialog);
        var charInput = charDialog.add("edittext", undefined, rangesToText(customSet.ranges), { multiline: true, scrolling: true });
        charInput.preferredSize = [CHAR_EDIT_SIZE[0], CHAR_EDIT_SIZE[1]];
        charInput.helpTip = getLabel("tooltip.customChars");

        var charButtonRow = addButtonRow(charDialog);
        var btnCharCancel = charButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnCharOK = charButtonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });

        var editedRanges = null;
        btnCharOK.onClick = function () {
            editedRanges = textToRanges(charInput.text);
            if (editedRanges.length === 0) {
                alert(getLabel("alert.noCustomChars"));
                return;
            }
            charDialog.close(1);
        };
        charDialog.onShow = function () {
            charInput.active = true;
        };
        alignRightOnlyButtonRow(charButtonRow);
        prepareDialogWindow(charDialog, SCRIPT_NAME + "_customChars");
        if (charDialog.show() !== 1 || !editedRanges) return false;
        customSet.ranges = editedRanges;
        return true;
    }

    // =========================================
    // フォント一覧 / Font list
    // =========================================
    /**
     * バリアブルフォントかどうかを、PostScript名かファミリー名に「VF」「Var」が付くかで判定する（一覧からの除外用）
     * Judged by "VF" / "Var" in the PostScript or family name (used to filter the list)
     * @param {string} psName - PostScript名
     * @param {string} familyName - ファミリー名
     * @returns {boolean} バリアブルフォントとみなすなら true
     */
    function isVariableFontName(psName, familyName) {
        return /VF|Var/.test(psName) || /VF|Var/.test(familyName);
    }

    /**
     * バリアブルフォントかどうかを axisVector で判定する（値があるのはバリアブルフォントだけ）
     * 名前に「VF」「Var」が付かないもの（Bodoni Moda など）を拾うため、［作成］時に使う数書体だけに使う
     * Checks axisVector; used only for the few chosen fonts on Create, to catch VFs without "VF"/"Var" in the name
     * @param {string} psName - PostScript名
     * @returns {boolean} バリアブルフォントなら true
     */
    function hasVariableAxes(psName) {
        try {
            var axisVector = app.textFonts.getByName(psName).axisVector;
            return axisVector !== undefined && axisVector !== null && String(axisVector) !== "";
        } catch (e) {
            return false; /* axisVector が無いバージョン / versions without axisVector */
        }
    }

    /**
     * インストール済みフォントをファミリーごとにまとめる（合成フォント・バリアブルフォントは除く）
     * @returns {{families: string[], stylesByFamily: Object, familyByPsName: Object}} ファミリー一覧と対応表
     */
    function collectFontFamilies() {
        var stylesByFamily = {};
        var familyByPsName = {};
        var families = [];
        var textFonts = app.textFonts;
        for (var i = 0; i < textFonts.length; i++) {
            var textFont = textFonts[i];
            var psName = textFont.name;
            if (psName.indexOf("ATC-") === 0) continue;
            var familyName = textFont.family;
            if (isVariableFontName(psName, familyName)) continue; /* 合成フォントに入れると適用時に落ちる / crashes Illustrator when applied */
            if (!stylesByFamily.hasOwnProperty(familyName)) {
                stylesByFamily[familyName] = [];
                families.push(familyName);
            }
            stylesByFamily[familyName].push({ style: textFont.style, psName: psName });
            familyByPsName[psName] = familyName;
        }
        families.sort(); /* 比較関数なしの sort が速い / plain sort() is fast */
        return { families: families, stylesByFamily: stylesByFamily, familyByPsName: familyByPsName };
    }

    /**
     * 文字のフォント・文字サイズ・ベースラインシフトを返す
     * @param {TextRange} textCharacter - 対象の文字
     * @returns {{psName: string, size: number, baselineShift: number}} PostScript名・文字サイズ（pt）・ベースラインシフト（pt）
     */
    function getCharacterFormat(textCharacter) {
        var charAttributes = textCharacter.characterAttributes;
        return {
            psName: String(charAttributes.textFont.name),
            size: Number(charAttributes.size),
            baselineShift: Number(charAttributes.baselineShift) || 0
        };
    }

    /**
     * テキストの最初の空白でない文字の書式を返す
     * @param {TextRange} textRange - 対象のテキスト
     * @returns {{psName: string, size: number, baselineShift: number}|null} 書式。見つからなければ null
     */
    function getFirstCharacterFormat(textRange) {
        /* 環境に無いフォントの文字は textFont の読み取りで例外になることがある / textFont can throw for missing fonts */
        try {
            var textCharacters = textRange.characters;
            for (var i = 0; i < textCharacters.length; i++) {
                if (/\s/.test(textCharacters[i].contents)) continue;
                return getCharacterFormat(textCharacters[i]);
            }
        } catch (e) {}
        return null;
    }

    /**
     * 文字が和文・かな・欧文のどれにあたるかを返す
     * @param {string} charText - 1文字
     * @returns {string|null} "ideograph"（漢字）/ "kanji"（その他の全角）/ "kana" / "roman"。空白なら null
     */
    function classifyCharacter(charText) {
        if (!charText || /\s/.test(charText)) return null;
        var code = charText.charCodeAt(0);
        if (code <= 0x7E) return "roman";
        if (code >= 0x3041 && code <= 0x30FF && code !== 0x30FB) return "kana"; /* 中黒は約物 / middle dot is punctuation */
        if ((code >= 0x4E00 && code <= 0x9FFF) || (code >= 0x3400 && code <= 0x4DBF) || (code >= 0xF900 && code <= 0xFAFF) ||
            (code >= 0xD840 && code <= 0xD87F) || code === 0x3005) return "ideograph";
        return "kanji";
    }

    /**
     * 1行のテキストから、和文・かな・欧文の書式を拾う
     * 和文は最初の文字（漢字を優先し、漢字が無ければ全角約物・記号）、
     * かな・欧文は文字数が一番多い書式（同数なら先に出てきたもの）
     * @param {TextRange} textRange - 対象のテキスト
     * @returns {{kanji: Object, kana: Object, roman: Object}} 書式（見つからなければ null）
     */
    function getFormatsByCharacterClass(textRange) {
        var firstFormats = { ideograph: null, kanji: null };
        var formatTallies = { kana: {}, roman: {} };
        /* 環境に無いフォントの文字は textFont の読み取りで例外になることがある / textFont can throw for missing fonts */
        try {
            var textCharacters = textRange.characters;
            for (var i = 0; i < textCharacters.length; i++) {
                var charClass = classifyCharacter(textCharacters[i].contents);
                if (!charClass) continue;
                if (formatTallies[charClass]) {
                    tallyFormat(formatTallies[charClass], getCharacterFormat(textCharacters[i]));
                } else if (!firstFormats[charClass]) {
                    firstFormats[charClass] = getCharacterFormat(textCharacters[i]);
                }
            }
        } catch (e) {}
        return {
            kanji: firstFormats.ideograph || firstFormats.kanji,
            kana: getMostFrequentFormat(formatTallies.kana),
            roman: getMostFrequentFormat(formatTallies.roman)
        };
    }

    /**
     * 書式ごとの文字数を数える
     * @param {Object} formatTally - 書式のキーごとの { format, count, order }
     * @param {Object} charFormat - 文字の書式（psName / size / baselineShift）
     * @returns {void}
     */
    function tallyFormat(formatTally, charFormat) {
        var formatKey = charFormat.psName + "|" + charFormat.size + "|" + charFormat.baselineShift;
        if (!formatTally[formatKey]) {
            var entryCount = 0;
            for (var existingKey in formatTally) entryCount++;
            formatTally[formatKey] = { format: charFormat, count: 0, order: entryCount };
        }
        formatTally[formatKey].count++;
    }

    /**
     * 文字数が一番多い書式を返す（同数なら先に出てきたもの）
     * @param {Object} formatTally - tallyFormat() で数えた結果
     * @returns {Object|null} 書式。1文字も無ければ null
     */
    function getMostFrequentFormat(formatTally) {
        var bestEntry = null;
        for (var formatKey in formatTally) {
            var tallyEntry = formatTally[formatKey];
            if (!bestEntry || tallyEntry.count > bestEntry.count ||
                (tallyEntry.count === bestEntry.count && tallyEntry.order < bestEntry.order)) bestEntry = tallyEntry;
        }
        return bestEntry ? bestEntry.format : null;
    }

    /**
     * 選択中のテキストを上から順に並べる
     * 複数のテキストオブジェクトなら上端の順、1つ（または文字選択）なら段落の順
     * @returns {TextRange[]} 上から順のテキスト（空の段落は除く）
     */
    function getSelectedTextsTopDown() {
        var textFrames = [];
        var singleRange = null;
        if (app.documents.length === 0) return [];
        var docSelection = app.activeDocument.selection;
        if (docSelection.typename === "TextRange") {
            singleRange = docSelection;
        } else {
            for (var i = 0; i < docSelection.length; i++) {
                if (docSelection[i].typename === "TextFrame") textFrames.push(docSelection[i]);
            }
            if (textFrames.length === 1) singleRange = textFrames[0].textRange;
        }

        var orderedTexts = [];
        if (singleRange) {
            var textParagraphs = singleRange.paragraphs;
            for (var j = 0; j < textParagraphs.length; j++) {
                if (textParagraphs[j].characters.length > 0) orderedTexts.push(textParagraphs[j]);
            }
            return orderedTexts;
        }

        /* 上端の高い順（同じ高さなら左から）に並べる / sort by top edge, then left edge */
        textFrames.sort(function (frameA, frameB) {
            var boundsA = frameA.geometricBounds;
            var boundsB = frameB.geometricBounds;
            if (Math.abs(boundsA[1] - boundsB[1]) > 0.001) return boundsB[1] - boundsA[1];
            return boundsA[0] - boundsB[0];
        });
        for (var k = 0; k < textFrames.length; k++) orderedTexts.push(textFrames[k].textRange);
        return orderedTexts;
    }

    /**
     * 選択中のテキストから和文・かな・欧文のフォントを拾う
     * 上から「和文・欧文」の2つ、または「和文・かな・欧文」の3つとして読む。
     * 1行（1段落）だけなら、文字の種類ごとに読む（和文・かな・欧文が混じった1行に対応）
     * （一覧にあるかの確認は呼び出し側で行う）
     * @returns {{fonts: Object, sizes: Object, baselineShifts: Object, isKanaLinked: boolean}} kanji / kana / roman ごとの PostScript名・文字サイズ（pt）・ベースラインシフト（pt）（無ければ null）と、かなを和文にそろえるか
     */
    function getSelectedTextFonts() {
        var foundFonts = { kanji: null, kana: null, roman: null };
        var foundSizes = { kanji: null, kana: null, roman: null };
        var foundShifts = { kanji: null, kana: null, roman: null };
        var orderedTexts = getSelectedTextsTopDown();
        var foundFormats = {};
        var isKanaLinked = false;
        if (orderedTexts.length === 1) {
            foundFormats = getFormatsByCharacterClass(orderedTexts[0]);
            /* かなが無いか、和文と同じ書式なら、かなは和文と同じ / no kana, or formatted like Japanese: same as Japanese */
            var kanaFormat = foundFormats.kana;
            var kanjiFormat = foundFormats.kanji;
            if (kanaFormat && kanjiFormat && kanaFormat.psName === kanjiFormat.psName &&
                kanaFormat.size === kanjiFormat.size && kanaFormat.baselineShift === kanjiFormat.baselineShift) {
                foundFormats.kana = null;
            }
            isKanaLinked = !foundFormats.kana;
        } else {
            var fontKeys = (orderedTexts.length >= 3) ? ["kanji", "kana", "roman"] : ["kanji", "roman"];
            for (var i = 0; i < fontKeys.length && i < orderedTexts.length; i++) {
                foundFormats[fontKeys[i]] = getFirstCharacterFormat(orderedTexts[i]);
            }
            isKanaLinked = orderedTexts.length === 2; /* 2つなら「和文・欧文」 / two texts: Japanese and Roman */
        }
        for (var fontKey in foundFonts) {
            var charFormat = foundFormats[fontKey];
            if (!charFormat) continue;
            foundFonts[fontKey] = charFormat.psName;
            foundSizes[fontKey] = charFormat.size;
            foundShifts[fontKey] = charFormat.baselineShift;
        }
        return { fonts: foundFonts, sizes: foundSizes, baselineShifts: foundShifts, isKanaLinked: isKanaLinked };
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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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
        fieldLabel.addEventListener("click", function () {
            numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
            numberInput.active = true;
        });

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
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
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "合成フォントを作成", en: "Create Composite Font" },
            customChars: { ja: "%1 の対象文字", en: "Characters in %1" }
        },
        panel: {
            fonts: { ja: "構成フォント", en: "Component Fonts" },
            customSets: { ja: "特例文字セット", en: "Custom Character Sets" }
        },
        dropdown: {
            noSource: { ja: "なし", en: "None" }
        },
        fallbackName: {
            customSet: { ja: "特例文字セット", en: "Custom set " }
        },
        fieldLabel: {
            fontName: { ja: "合成フォント名", en: "Composite font name" },
            customSource: { ja: "読み込み", en: "Load from" }
        },
        columnLabel: {
            family: { ja: "フォント", en: "Font" },
            style: { ja: "スタイル", en: "Style" },
            size: { ja: "サイズ", en: "Size" },
            baseline: { ja: "ベースライン", en: "Baseline" }
        },
        fontGroup: {
            kanji: { ja: "和文", en: "Japanese" },
            kana: { ja: "かな", en: "Kana" },
            roman: { ja: "欧文", en: "Roman" }
        },
        tooltip: {
            fontName: {
                ja: "合成フォント名のフォント名の部分です。選んだフォントから「和文のPS名-かなのPS名-欧文のPS名」（ウエイトの部分は除く。かなが和文と同じなら省く）を自動で入れ、ウエイトと合わせて29文字に収まるよう切り詰めます。手で書き換えたら自動入力は止まり、空にすると再開します。半角英数字と記号のみ（「/」「:」は不可）。",
                en: "The font-name part of the composite font name. It fills in as \"JapanesePS-KanaPS-RomanPS\" (weights removed; Kana omitted when same as Japanese), cut so that it fits in 29 characters with the weight, until you edit it; clear it to resume. ASCII letters, digits and symbols only (no \"/\" or \":\")."
            },
            size: { ja: "和文（漢字）に対する大きさ（％）", en: "Size relative to the Japanese font (Kanji) (%)" },
            baseline: { ja: "ベースラインの移動量（％）。正の値で上がります。", en: "Baseline shift (%). Positive values move up." },
            kanji: { ja: "漢字・全角約物・全角記号のフォント", en: "Font for Kanji, punctuation and symbols" },
            kana: { ja: "ひらがな・カタカナのフォント", en: "Font for hiragana and katakana" },
            roman: { ja: "半角欧文・半角数字のフォント", en: "Font for alphabetic characters and numerals" },
            nameLength: { ja: "合成フォント名（ハイフンとウエイトを含む）の文字数と上限", en: "Length of the composite font name (including the hyphen and weight) and the limit" },
            customCharsButton: { ja: "対象文字を表示・編集します", en: "Show and edit the characters in this set" },
            customChars: {
                ja: "このセットに含める文字です。1文字ずつ入力します（重複・空白・改行は無視）。",
                en: "Characters in this set, one per character (duplicates, spaces and line breaks are ignored)."
            },
            weight: {
                ja: "ウエイト。和文のスタイル名を自動で入れます（空白・日本語は除く）。手で書き換えたら自動入力は止まり、空にすると再開します。空のままなら付けません。",
                en: "Weight. It fills in from the Japanese style name (spaces and non-ASCII removed) until you edit it; clear it to resume. Left empty, no weight is added."
            },
            customSource: {
                ja: "特例文字セットを読み込む合成フォントを選びます（特例文字セットを持つものだけ並びます）。選び直すと、入力中の値を残したままダイアログを開き直します。",
                en: "Choose a composite font to load custom character sets from (only fonts with custom character sets are listed). Changing it reopens the dialog, keeping your entries."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            inDesign: {
                ja: "起動中の InDesign にも同じ設定の合成フォントを作ります（InDesign はインストールされていないと選べません）。",
                en: "Also creates the same composite font in InDesign, which must be running (unavailable when InDesign is not installed)."
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        checkbox: {
            inDesign: { ja: "InDesign にも作成", en: "Also create in InDesign" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            create: { ja: "作成", en: "Create" },
            customChars: { ja: "文字...", en: "Chars..." }
        },
        alert: {
            variableFont: {
                ja: "「%1」はバリアブルフォントです。バリアブルフォントを入れた合成フォントは、適用すると Illustrator が落ちるため使えません。",
                en: "\"%1\" is a variable font. Composite fonts containing variable fonts crash Illustrator when applied, so it cannot be used."
            },
            nameTooLong: { ja: "合成フォント名は%1文字以内にしてください（現在%2文字）。", en: "Keep the name within %1 characters (currently %2)." },
            noCustomChars: { ja: "対象文字を1文字以上入力してください。", en: "Enter at least one character." },
            invalidName: { ja: "合成フォント名は半角英数字と記号で入力してください（「/」「:」は使えません）。", en: "Use ASCII letters, digits and symbols for the name (no \"/\" or \":\")." },
            invalidNumber: { ja: "サイズまたはベースラインの数値が正しくありません：", en: "Invalid size or baseline value: " },
            noFolder: { ja: "合成フォントの保存先フォルダーが見つかりません。フォルダーを選んでください。", en: "The composite font folder was not found. Choose a folder." },
            chooseFolder: { ja: "合成フォントの保存先", en: "Composite font folder" },
            overwrite: { ja: "「%1」は既にあります。上書きしますか？", en: "\"%1\" already exists. Overwrite it?" },
            writeFailed: { ja: "書き出せませんでした：", en: "Could not write: " },
            done: {
                ja: "合成フォント「%1」を作成しました。\nIllustrator を再起動すると使えるようになります。\n\n%2",
                en: "Created the composite font \"%1\".\nRestart Illustrator to use it.\n\n%2"
            },
            inDesignDone: { ja: "InDesign：作成しました（再起動は不要です）。", en: "InDesign: created (no restart needed)." },
            inDesignSkipped: { ja: "InDesign：作成を取りやめました。", en: "InDesign: skipped." },
            inDesignMissing: { ja: "InDesign：見つからないため作成しませんでした。", en: "InDesign: not found, so nothing was created." },
            inDesignNotRunning: {
                ja: "InDesign：起動していないため作成しませんでした。InDesign を起動してから実行してください。",
                en: "InDesign: not running, so nothing was created. Launch InDesign and run the script again."
            },
            inDesignOverwrite: {
                ja: "InDesign に合成フォント「%1」は既にあります。上書きしますか？",
                en: "InDesign already has the composite font \"%1\". Overwrite it?"
            },
            inDesignFontMissing: {
                ja: "InDesign：フォントが見つからないため作成しませんでした：",
                en: "InDesign: fonts not found, so nothing was created: "
            },
            inDesignFailed: { ja: "InDesign：作成できませんでした：", en: "InDesign: could not create: " },
            inDesignTimeout: { ja: "応答がありません", en: "no response" }
        }
    };

    /**
     * 文字列にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} itemName - 項目名
     * @returns {string} コロン付きの項目名
     */
    function appendColon(itemName) {
        return itemName + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /* フォントを共有する文字セットのまとまり / character sets that share one font */
    var FONT_GROUPS = [
        { key: "kanji", charsets: [0, 2, 3], adjustable: false }, /* 漢字・全角約物・全角記号（基準なので100％・0％固定） */
        { key: "kana", charsets: [1], adjustable: true },         /* かな */
        { key: "roman", charsets: [4, 5], adjustable: true }      /* 半角欧文・半角数字 */
    ];
    var ROMAN_GROUP_INDEX = 2;

    /**
     * スタイルのドロップダウンを、ファミリーに合わせて作り直す
     * @param {DropDownList} styleDropdown - スタイルのドロップダウン
     * @param {Object[]} styleItems - { style, psName } の配列
     * @param {string} [preferredPsName] - 選んでおく PostScript名
     * @returns {void}
     */
    function fillStyleDropdown(styleDropdown, styleItems, preferredPsName) {
        styleDropdown.removeAll();
        var selectIndex = 0;
        for (var i = 0; i < styleItems.length; i++) {
            styleDropdown.add("item", styleItems[i].style);
            if (styleItems[i].psName === preferredPsName) selectIndex = i;
        }
        if (styleItems.length > 0) styleDropdown.selection = selectIndex;
    }

    /**
     * PostScript名の末尾のウエイト（スタイル）部分を外す
     * 最後の「-」より後ろが、スタイル名（空白を除く）と同じか、ウエイトらしい語（Bold・W3・B など）のときだけ外す
     * Drops the trailing weight: the part after the last "-" when it equals the style name (spaces removed)
     * or looks like a weight (Bold, W3, B, ...)
     * @param {string} psName - PostScript名（例 "PA1MinchoStdN-Bold"）
     * @param {string|null} styleName - スタイル名（例 "B"）
     * @returns {string} ウエイトを外した名前（例 "PA1MinchoStdN"）
     */
    function stripWeightFromPsName(psName, styleName) {
        var hyphenPos = psName.lastIndexOf("-");
        if (hyphenPos <= 0) return psName;
        var lastPart = psName.substring(hyphenPos + 1);
        var styleKey = String(styleName || "").replace(/\s+/g, "").toLowerCase();
        var looksLikeWeight = /^(W\d+|[A-Z]{1,3})$/.test(lastPart) ||
            /^((Extra|Ultra|Semi|Demi|Ex)?(Thin|Hairline|Light|Regular|Book|Roman|Normal|Medium|Bold|Heavy|Black)|Italic|Oblique)(It|Italic|Oblique)?$/i.test(lastPart);
        return (lastPart.toLowerCase() === styleKey || looksLikeWeight) ? psName.substring(0, hyphenPos) : psName;
    }

    /**
     * 「ファミリー・スタイル」のフォント選択行を追加する
     * @param {Panel} parent - 追加先
     * @param {string} rowLabelText - 行の項目名（コロン込み）
     * @param {string} rowTip - 項目名のツールチップ
     * @param {number} rowLabelWidth - 項目名の幅
     * @param {Object} fontCatalog - collectFontFamilies() の結果
     * @param {string} initialPsName - 初期フォントの PostScript名
     * @param {Function} onFontChange - フォントを選び直したときに呼ぶ関数
     * @returns {{getPsName: Function, getStyle: Function, selectPsName: Function, rowGroup: Group}} PostScript名・スタイル名を返す関数（未選択なら null）、選ぶ関数と行
     */
    function addFontPickerRow(parent, rowLabelText, rowTip, rowLabelWidth, fontCatalog, initialPsName, onFontChange) {
        var fontRowGroup = parent.add("group");
        fontRowGroup.orientation = "row";
        fontRowGroup.alignChildren = ["left", "center"];

        var rowLabel = fontRowGroup.add("statictext", undefined, rowLabelText);
        rowLabel.preferredSize.width = rowLabelWidth;
        rowLabel.justify = "right";
        rowLabel.helpTip = rowTip;

        /* 項目は1件ずつ足す。選択テキストの DOM を読んだあとで大きな配列をまとめて渡すと Illustrator が落ちた
           Add items one by one: passing the large array at creation crashed Illustrator after reading the selection's DOM */
        var familyDropdown = fontRowGroup.add("dropdownlist", undefined, []);
        for (var i = 0; i < fontCatalog.families.length; i++) familyDropdown.add("item", fontCatalog.families[i]);
        familyDropdown.preferredSize.width = FAMILY_DROPDOWN_WIDTH;
        var styleDropdown = fontRowGroup.add("dropdownlist", undefined, []);
        styleDropdown.preferredSize.width = STYLE_DROPDOWN_WIDTH;

        var isSelecting = false; /* スクリプトから選び直している間は onChange を無視 / ignore onChange while selecting from code */
        familyDropdown.onChange = function () {
            if (isSelecting || !familyDropdown.selection) return;
            fillStyleDropdown(styleDropdown, fontCatalog.stylesByFamily[fontCatalog.families[familyDropdown.selection.index]], null);
            onFontChange();
        };
        styleDropdown.onChange = function () {
            if (isSelecting) return;
            onFontChange();
        };

        /**
         * PostScript名でフォントを選ぶ
         * @param {string} psName - PostScript名
         * @returns {void}
         */
        function selectPsName(psName) {
            var familyName = fontCatalog.familyByPsName[psName];
            if (!familyName) return;
            isSelecting = true;
            for (var i = 0; i < fontCatalog.families.length; i++) {
                if (fontCatalog.families[i] === familyName) {
                    familyDropdown.selection = i;
                    break;
                }
            }
            fillStyleDropdown(styleDropdown, fontCatalog.stylesByFamily[familyName], psName);
            isSelecting = false;
        }
        selectPsName(initialPsName);

        return {
            getPsName: function () {
                if (!familyDropdown.selection || !styleDropdown.selection) return null;
                var familyName = fontCatalog.families[familyDropdown.selection.index];
                return fontCatalog.stylesByFamily[familyName][styleDropdown.selection.index].psName;
            },
            getStyle: function () {
                return styleDropdown.selection ? String(styleDropdown.selection.text) : null;
            },
            selectPsName: selectPsName,
            rowGroup: fontRowGroup
        };
    }

    /**
     * 列の幅を固定したグループを足す（見出しと各行で列をそろえる）
     * @param {Group} parent - 追加先の行
     * @param {number} columnWidth - 列の幅
     * @returns {Group} 列のグループ
     */
    function addFixedColumn(parent, columnWidth) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "row";
        columnGroup.alignChildren = ["left", "center"];
        columnGroup.margins = 0;
        columnGroup.minimumSize.width = columnGroup.preferredSize.width = columnGroup.maximumSize.width = columnWidth;
        return columnGroup;
    }

    /**
     * 「フォント・スタイル・サイズ・ベースライン」の見出し行を足す
     * @param {Panel} parent - 追加先
     * @param {number} rowLabelWidth - 行の項目名の幅（その分を空ける）
     * @returns {void}
     */
    function addColumnHeaderRow(parent, rowLabelWidth) {
        var headerRowGroup = parent.add("group");
        headerRowGroup.orientation = "row";
        headerRowGroup.alignChildren = ["left", "center"];
        var headerSpecs = [
            { text: "", width: rowLabelWidth },
            { text: getLabel("columnLabel.family"), width: FAMILY_DROPDOWN_WIDTH },
            { text: getLabel("columnLabel.style"), width: STYLE_DROPDOWN_WIDTH },
            { text: getLabel("columnLabel.size"), width: VALUE_COLUMN_WIDTH },
            { text: getLabel("columnLabel.baseline"), width: VALUE_COLUMN_WIDTH }
        ];
        for (var i = 0; i < headerSpecs.length; i++) {
            addFixedColumn(headerRowGroup, headerSpecs[i].width).add("statictext", undefined, headerSpecs[i].text);
        }
    }

    /**
     * フォント行の右に「サイズ・ベースライン」の欄を足す
     * @param {Group} fontRowGroup - addFontPickerRow() で作った行
     * @param {number} initialSize - サイズの初期値（％）
     * @param {number} initialBaseline - ベースラインの初期値（％）
     * @param {boolean} isFixed - 固定値として無効表示にするなら true（和文）
     * @returns {{sizeInput: EditText, baselineInput: EditText}} 入力欄
     */
    function addSizeBaselineRow(fontRowGroup, initialSize, initialBaseline, isFixed) {
        var sizeInput = addSteppedField(addFixedColumn(fontRowGroup, VALUE_COLUMN_WIDTH), {
            label: "", text: initialSize + "%", characters: 5, step: 1, min: 1, unit: "%"
        });
        sizeInput.helpTip = getLabel("tooltip.size");
        var baselineInput = addSteppedField(addFixedColumn(fontRowGroup, VALUE_COLUMN_WIDTH), {
            label: "", text: initialBaseline + "%", characters: 5, step: 1, unit: "%"
        });
        baselineInput.helpTip = getLabel("tooltip.baseline");
        if (isFixed) {
            setSteppedFieldEnabled(sizeInput, false);
            setSteppedFieldEnabled(baselineInput, false);
        }
        return { sizeInput: sizeInput, baselineInput: baselineInput };
    }

    /**
     * 特例文字セット1つ分の行（フォント・サイズ・ベースライン・［文字…］）を追加する
     * 初期フォントは文字に応じて和文・かな・欧文から選び、サイズ・ベースラインもそのまとまりに合わせる
     * @param {Panel} customPanel - 追加先
     * @param {Object} customSet - 特例文字セット（［文字…］で ranges を書き換える）
     * @param {Object} dialogState - ダイアログの状態（fonts・sizes・baselines を初期値に使う）
     * @param {Object} fontCatalog - collectFontFamilies() の結果
     * @param {number} rowLabelWidth - 項目名の幅
     * @returns {{picker: Object, adjust: Object, label: string}} フォント選択・数値欄・行の名前
     */
    function addCustomSetRow(customPanel, customSet, dialogState, fontCatalog, rowLabelWidth) {
        var nearestGroup = classifyCustomSet(customSet.ranges);
        var customPicker = addFontPickerRow(customPanel, appendColon(customSet.displayName), describeCustomSet(customSet.ranges),
            rowLabelWidth, fontCatalog, dialogState.fonts[nearestGroup], function () {});
        var customAdjust = addSizeBaselineRow(customPicker.rowGroup,
            (nearestGroup === "kanji") ? 100 : dialogState.sizes[nearestGroup],
            (nearestGroup === "kanji") ? 0 : dialogState.baselines[nearestGroup], false);
        /* 対象文字の表示・編集ボタン。書き換えたら項目名のツールチップも更新 / show and edit the characters, then refresh the tooltip */
        var btnChars = customPicker.rowGroup.add("button", undefined, getLabel("button.customChars"));
        btnChars.helpTip = getLabel("tooltip.customCharsButton");
        btnChars.onClick = function () {
            if (editCustomSetCharacters(customSet)) customPicker.rowGroup.children[0].helpTip = describeCustomSet(customSet.ranges);
        };
        return { picker: customPicker, adjust: customAdjust, label: customSet.displayName };
    }

    /**
     * ［特例文字セット］パネル（読み込み元と、セットごとの行）を追加する。読み込み元の候補が無ければ何も足さない
     * @param {Window} fontDialog - 追加先のダイアログ
     * @param {Object} dialogState - ダイアログの状態（sourceIndex で読み込み元を選ぶ）
     * @param {Object[]} customSetSources - collectCustomSetSources() の結果
     * @param {Object} fontCatalog - collectFontFamilies() の結果
     * @param {number} rowLabelWidth - 項目名の幅
     * @returns {{sourceDropdown: DropDownList|null, customSets: Object[], customRows: Object[]}} 読み込み元の選択・セット・行
     */
    function addCustomSetsPanel(fontDialog, dialogState, customSetSources, fontCatalog, rowLabelWidth) {
        var customSets = (dialogState.sourceIndex > 0) ? customSetSources[dialogState.sourceIndex - 1].customSets : [];
        var customRows = [];
        if (customSetSources.length === 0) return { sourceDropdown: null, customSets: customSets, customRows: customRows };

        var customPanel = fontDialog.add("panel", undefined, getLabel("panel.customSets"));
        setupPanel(customPanel);
        customPanel.alignChildren = ["left", "center"];

        var sourceRowGroup = customPanel.add("group");
        sourceRowGroup.orientation = "row";
        sourceRowGroup.alignChildren = ["left", "center"];
        var sourceLabel = sourceRowGroup.add("statictext", undefined, labelText("fieldLabel.customSource"));
        sourceLabel.preferredSize.width = rowLabelWidth;
        sourceLabel.justify = "right";
        var sourceDropdown = sourceRowGroup.add("dropdownlist", undefined, []);
        sourceDropdown.add("item", getLabel("dropdown.noSource"));
        for (var i = 0; i < customSetSources.length; i++) sourceDropdown.add("item", customSetSources[i].label);
        sourceDropdown.preferredSize.width = FAMILY_DROPDOWN_WIDTH;
        sourceDropdown.helpTip = getLabel("tooltip.customSource");
        sourceDropdown.selection = dialogState.sourceIndex;

        if (customSets.length > 0) addColumnHeaderRow(customPanel, rowLabelWidth);
        for (var k = 0; k < customSets.length; k++) {
            customRows.push(addCustomSetRow(customPanel, customSets[k], dialogState, fontCatalog, rowLabelWidth));
        }
        return { sourceDropdown: sourceDropdown, customSets: customSets, customRows: customRows };
    }

    /**
     * 合成フォント名が使えるかを確かめる
     * @param {string} fontName - ハイフンでつないだ合成フォント名
     * @returns {string} 使えなければ警告文、使えれば空文字
     */
    function validateFontName(fontName) {
        if (!/^[\x20-\x7e]+$/.test(fontName) || /[\/:]/.test(fontName) || /^\s|\s$/.test(fontName)) return getLabel("alert.invalidName");
        if (fontName.length > MAX_NAME_LENGTH) {
            return getLabel("alert.nameTooLong").replace("%1", MAX_NAME_LENGTH).replace("%2", fontName.length);
        }
        return "";
    }

    /**
     * 数値欄からサイズとベースラインを読む
     * @param {Object} adjust - addSizeBaselineRow() の戻り値
     * @param {string} rowName - エラー表示用の行の名前
     * @returns {{size: number, baseline: number}|null} 値。数値でなければ null（知らせる）
     */
    function readAdjustValues(adjust, rowName) {
        var sizeValue = parseFloat(adjust.sizeInput.text);
        var baselineValue = parseFloat(adjust.baselineInput.text);
        if (isNaN(sizeValue) || sizeValue <= 0 || isNaN(baselineValue)) {
            alert(getLabel("alert.invalidNumber") + rowName);
            return null;
        }
        return { size: sizeValue, baseline: baselineValue };
    }

    /**
     * ダイアログの入力から合成フォントの設定を組み立てる。和文は100％・0％、比率はすべて100％で固定
     * 数値が正しくないとき・バリアブルフォントが入っているときは知らせて null を返す
     * @param {string} fontName - 合成フォント名
     * @param {Object} dialogInputs - fontPickers / adjustInputs / customRows / customSets / isKanaLinked
     * @returns {Object|null} 合成フォントの設定
     */
    function buildFontSpecFromDialog(fontName, dialogInputs) {
        var fontSpec = { name: fontName, fonts: [], size: [], baseline: [], hScale: [], vScale: [], customSets: dialogInputs.customSets };
        var fontPickers = dialogInputs.fontPickers;

        /**
         * 文字セット1つ分の値を書き込む
         * @param {number} charsetIndex - 文字セットの番号
         * @param {string} psName - PostScript名
         * @param {{size: number, baseline: number}} setValues - サイズとベースライン（％）
         * @returns {void}
         */
        function setCharsetValues(charsetIndex, psName, setValues) {
            fontSpec.fonts[charsetIndex] = psName;
            fontSpec.size[charsetIndex] = setValues.size;
            fontSpec.baseline[charsetIndex] = setValues.baseline;
            fontSpec.hScale[charsetIndex] = 100;
            fontSpec.vScale[charsetIndex] = 100;
        }

        for (var i = 0; i < FONT_GROUPS.length; i++) {
            var isLinkedKana = dialogInputs.isKanaLinked && FONT_GROUPS[i].key === "kana";
            var psName = fontPickers[isLinkedKana ? 0 : i].getPsName();
            if (!psName) return null;
            var groupValues = { size: 100, baseline: 0 }; /* ディムのかなは和文と同じ100％・0％ / dimmed Kana matches Japanese */
            if (dialogInputs.adjustInputs[i] && !isLinkedKana) {
                groupValues = readAdjustValues(dialogInputs.adjustInputs[i], getLabel("fontGroup." + FONT_GROUPS[i].key));
                if (!groupValues) return null;
            }
            for (var j = 0; j < FONT_GROUPS[i].charsets.length; j++) setCharsetValues(FONT_GROUPS[i].charsets[j], psName, groupValues);
        }
        /* 特例文字は標準6セットの後ろに並べる / custom sets follow the six standard sets */
        var customRows = dialogInputs.customRows;
        for (var k = 0; k < customRows.length; k++) {
            var customPsName = customRows[k].picker.getPsName();
            if (!customPsName) return null;
            var customValues = readAdjustValues(customRows[k].adjust, customRows[k].label);
            if (!customValues) return null;
            setCharsetValues(CHARSET_COUNT + k, customPsName, customValues);
        }
        /* バリアブルフォントを入れた合成フォントは、適用すると Illustrator が落ちる / variable fonts crash Illustrator when applied */
        for (var fontIndex = 0; fontIndex < fontSpec.fonts.length; fontIndex++) {
            if (hasVariableAxes(fontSpec.fonts[fontIndex])) {
                alert(getLabel("alert.variableFont").replace("%1", fontSpec.fonts[fontIndex]));
                return null;
            }
        }
        return fontSpec;
    }

    /**
     * ダイアログを表示して設定を受け取る
     * 特例文字の読み込み元を選び直したときは、入力中の値を dialogState に控えて { reload: true } を返す
     * @param {Object} fontCatalog - collectFontFamilies() の結果
     * @param {Object} dialogState - name / isAutoName / weight / isAutoWeight / isKanaLinked / fonts・sizes・baselines（kanji・kana・roman）/ sourceIndex / hasInDesign / forInDesign
     * @param {Object[]} customSetSources - collectCustomSetSources() の結果
     * @returns {Object|null} 合成フォントの設定、{ reload: true }、キャンセルなら null
     */
    function showCompositeFontDialog(fontCatalog, dialogState, customSetSources) {
        var fontDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(fontDialog);

        var nameRowGroup = fontDialog.add("group");
        nameRowGroup.orientation = "row";
        nameRowGroup.alignment = ["fill", "center"];
        nameRowGroup.alignChildren = ["left", "center"];
        nameRowGroup.margins = [NAME_ROW_LEFT_MARGIN, 0, 0, 0]; /* 左に1文字ほど空ける / indent by about one character */
        nameRowGroup.add("statictext", undefined, labelText("fieldLabel.fontName"));
        /* 合成フォント名は「フォント名の部分」と「ウエイト」の2欄に分け、ハイフンでつなぐ
           The name is split into a font-name part and a weight, joined with a hyphen */
        var nameInput = nameRowGroup.add("edittext", undefined, dialogState.isAutoName ? "" : dialogState.name);
        /* 最低でも NAME_FIELD_CHARACTERS 文字分の幅を取り、余りは下のパネルの幅まで伸ばす
           At least NAME_FIELD_CHARACTERS wide, stretching to the panel width */
        nameInput.characters = NAME_FIELD_CHARACTERS;
        nameInput.alignment = ["fill", "center"];
        nameInput.helpTip = getLabel("tooltip.fontName");
        nameRowGroup.add("statictext", undefined, "-");
        var weightInput = nameRowGroup.add("edittext", undefined, dialogState.isAutoWeight ? "" : dialogState.weight);
        weightInput.characters = 6;
        weightInput.helpTip = getLabel("tooltip.weight");
        /* つないだ名前の文字数（上限を超えたら赤） / length of the joined name, red when over the limit */
        var nameLengthText = nameRowGroup.add("statictext", undefined, "");
        nameLengthText.preferredSize.width = NAME_LENGTH_WIDTH;
        nameLengthText.helpTip = getLabel("tooltip.nameLength");
        /* 既定の foregroundColor は未設定（代入すると Error 47）なので、通常色のペンも明暗に合わせて作る
           The default foregroundColor is unset (assigning it back raises Error 47), so build the normal pen too */
        var nameLengthDefaultPen = nameLengthText.graphics.newPen(nameLengthText.graphics.PenType.SOLID_COLOR,
            STEPPER_UI_DARK ? [0.85, 0.85, 0.85, 1] : [0.1, 0.1, 0.1, 1], 1);
        var nameLengthOverPen = nameLengthText.graphics.newPen(nameLengthText.graphics.PenType.SOLID_COLOR, [0.9, 0.2, 0.2, 1], 1);

        /* フォント：和文・かな・欧文の3行 / fonts: one row each for Japanese, Kana and Roman */
        var fontsPanel = fontDialog.add("panel", undefined, getLabel("panel.fonts"));
        setupPanel(fontsPanel);
        fontsPanel.alignChildren = ["left", "center"];
        var fontPickers = [];

        /* フォント名の部分は「和文PS名-かなPS名-欧文PS名」（かなが和文と同じなら省く）、ウエイトは和文のスタイル名を自動で入れる。
           手で書き換えた欄は追従をやめ、空にすると再開する
           The name part follows "JapanesePS-KanaPS-RomanPS" (Kana omitted when same as Japanese) and the weight follows
           the Japanese style, until edited by hand; clearing a field resumes it */
        var lastAutoName = "";
        var lastAutoWeight = "";
        /**
         * 名前の自動入力と文字数表示を更新する
         * @returns {void}
         */
        function updateAutoName() {
            fillAutoName();
            updateNameLength();
        }

        /**
         * つないだ名前の文字数を「24 / 29」の形で表示する
         * @returns {void}
         */
        function updateNameLength() {
            var nameLength = getCombinedName().length;
            nameLengthText.text = nameLength + " / " + MAX_NAME_LENGTH;
            nameLengthText.graphics.foregroundColor = (nameLength > MAX_NAME_LENGTH) ? nameLengthOverPen : nameLengthDefaultPen;
        }

        /**
         * 選ばれているフォントから、フォント名の部分とウエイトを作って入れる（手で書き換えた欄はそのまま）
         * @returns {void}
         */
        function fillAutoName() {
            if (fontPickers.length < FONT_GROUPS.length) return; /* 行を作っている途中 / rows still being built */
            /* かなの行がディムの間は、和文と同じフォントを表示しておく / while Kana is dimmed, show the Japanese font there */
            if (dialogState.isKanaLinked && fontPickers[0].getPsName()) fontPickers[1].selectPsName(fontPickers[0].getPsName());
            if (weightInput.text === "" || weightInput.text === lastAutoWeight) {
                /* 半角英数字と記号以外（空白・日本語）は名前に使えないので落とす / drop characters not allowed in the name */
                lastAutoWeight = (fontPickers[0].getStyle() || "").replace(/[^\x21-\x7e]/g, "").replace(/[\/:]/g, "");
                weightInput.text = lastAutoWeight;
            }
            if (nameInput.text !== "" && nameInput.text !== lastAutoName) return;
            /* ウエイトは別の欄に入るので、PS名からウエイト（スタイル）の部分を外す / the weight has its own field, so drop it from the PS names */
            var baseNames = [];
            for (var i = 0; i < fontPickers.length; i++) {
                baseNames.push(stripWeightFromPsName(fontPickers[i].getPsName() || "", fontPickers[i].getStyle()));
            }
            var kanjiBase = baseNames[0];
            var kanaBase = baseNames[1];
            var romanBase = baseNames[ROMAN_GROUP_INDEX];
            lastAutoName = (kanaBase === kanjiBase) ? kanjiBase + "-" + romanBase : kanjiBase + "-" + kanaBase + "-" + romanBase;
            /* ウエイトと合わせて上限に収まるよう切り詰め、末尾の区切りを落とす / truncate so the name fits with the weight */
            var weightLength = weightInput.text ? weightInput.text.length + 1 : 0;
            lastAutoName = lastAutoName.substr(0, Math.max(1, MAX_NAME_LENGTH - weightLength)).replace(/[_\-\s]+$/, "");
            nameInput.text = lastAutoName;
        }

        /**
         * 2つの欄をハイフンでつないだ合成フォント名を返す（ウエイトが空ならフォント名の部分だけ）
         * @returns {string} 合成フォント名
         */
        function getCombinedName() {
            return weightInput.text ? nameInput.text + "-" + weightInput.text : nameInput.text;
        }
        weightInput.onChanging = function () {
            updateAutoName(); /* ウエイトの長さに合わせてフォント名の部分を詰め直す / refit the name part to the weight */
        };
        nameInput.onChanging = function () {
            updateNameLength();
        };

        /* ［フォント］と［特例文字］で項目名の幅をそろえる。特例文字のパネルがあるときはセット名の入る幅に広げる
           Use one label width in both panels, widened for set names when the Custom Sets panel is shown */
        var rowLabelWidth = (customSetSources.length > 0) ? CUSTOM_LABEL_WIDTH : ROW_LABEL_WIDTH;

        /* 見出しを最上部に置き、各行はフォント・スタイル・サイズ・ベースラインを横に並べる。和文のサイズ・ベースラインは固定
           Column headers on top; each row lays out font, style, size and baseline. The Japanese size/baseline stay fixed */
        addColumnHeaderRow(fontsPanel, rowLabelWidth);
        var adjustInputs = [];
        for (var i = 0; i < FONT_GROUPS.length; i++) {
            var groupKey = FONT_GROUPS[i].key;
            var groupPicker = addFontPickerRow(fontsPanel, labelText("fontGroup." + groupKey), getLabel("tooltip." + groupKey),
                rowLabelWidth, fontCatalog, dialogState.fonts[groupKey], updateAutoName);
            fontPickers.push(groupPicker);
            var groupAdjust = addSizeBaselineRow(groupPicker.rowGroup, dialogState.sizes[groupKey], dialogState.baselines[groupKey],
                !FONT_GROUPS[i].adjustable);
            adjustInputs.push(FONT_GROUPS[i].adjustable ? groupAdjust : null);
        }
        /* 選択したテキストが2つ（和文・欧文）のときは、かなは和文と同じにして行をディムにする
           With two selected texts (Japanese, Roman), Kana uses the Japanese font and its row is dimmed */
        if (dialogState.isKanaLinked) {
            fontPickers[1].rowGroup.enabled = false;
            setSteppedFieldEnabled(adjustInputs[1].sizeInput, false);
            setSteppedFieldEnabled(adjustInputs[1].baselineInput, false);
            redrawSteppersIn(fontPickers[1].rowGroup);
        }
        updateAutoName();

        /* 特例文字：読み込み元の合成フォントと、セットごとのフォント・サイズ・ベースライン
           Custom sets: the source composite font, then font, size and baseline per set */
        var customSetsUI = addCustomSetsPanel(fontDialog, dialogState, customSetSources, fontCatalog, rowLabelWidth);
        var sourceDropdown = customSetsUI.sourceDropdown;

        var buttonRow = addButtonRow(fontDialog);
        /* InDesign が無ければ無効 / disabled when InDesign is not installed */
        var inDesignCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.inDesign"));
        inDesignCheckbox.helpTip = getLabel("tooltip.inDesign");
        inDesignCheckbox.enabled = dialogState.hasInDesign;
        inDesignCheckbox.value = dialogState.hasInDesign && dialogState.forInDesign;
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.create"), { name: "ok" });

        var dialogResult = null;

        /* 読み込み元を選び直したら、入力中の値を控えて開き直す / reopen with the current entries when the source changes */
        if (sourceDropdown) {
            sourceDropdown.onChange = function () {
                if (!sourceDropdown.selection || sourceDropdown.selection.index === dialogState.sourceIndex) return;
                dialogState.isAutoName = (nameInput.text === "" || nameInput.text === lastAutoName);
                dialogState.name = nameInput.text;
                dialogState.isAutoWeight = (weightInput.text === "" || weightInput.text === lastAutoWeight);
                dialogState.weight = weightInput.text;
                for (var i = 0; i < FONT_GROUPS.length; i++) {
                    var groupKey = FONT_GROUPS[i].key;
                    dialogState.fonts[groupKey] = fontPickers[i].getPsName() || dialogState.fonts[groupKey];
                    if (!adjustInputs[i]) continue;
                    var sizeValue = parseFloat(adjustInputs[i].sizeInput.text);
                    var baselineValue = parseFloat(adjustInputs[i].baselineInput.text);
                    if (!isNaN(sizeValue)) dialogState.sizes[groupKey] = sizeValue;
                    if (!isNaN(baselineValue)) dialogState.baselines[groupKey] = baselineValue;
                }
                dialogState.sourceIndex = sourceDropdown.selection.index;
                dialogState.forInDesign = inDesignCheckbox.value;
                dialogResult = { reload: true };
                fontDialog.close(3);
            };
        }

        btnOK.onClick = function () {
            var fontName = getCombinedName();
            var nameProblem = validateFontName(fontName);
            if (nameProblem) {
                alert(nameProblem);
                return;
            }
            var fontSpec = buildFontSpecFromDialog(fontName, {
                fontPickers: fontPickers, adjustInputs: adjustInputs, isKanaLinked: dialogState.isKanaLinked,
                customRows: customSetsUI.customRows, customSets: customSetsUI.customSets
            });
            if (!fontSpec) return;
            fontSpec.forInDesign = inDesignCheckbox.enabled && inDesignCheckbox.value;
            dialogResult = fontSpec;
            fontDialog.close(1);
        };

        fontDialog.onShow = function () {
            nameInput.active = true;
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(fontDialog, SCRIPT_NAME);
        var closeCode = fontDialog.show();
        return (closeCode === 1 || closeCode === 3) ? dialogResult : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    /**
     * ダイアログの初期フォントを決める。選択中のテキストのフォントを優先し、一覧に無ければ（合成・バリアブル）初期フォント
     * @param {Object} selectedFonts - getSelectedTextFonts() の fonts（kanji / kana / roman の PostScript名か null）
     * @param {Object} fontCatalog - collectFontFamilies() の結果
     * @returns {{kanji: string, kana: string, roman: string}} PostScript名
     */
    function getInitialFonts(selectedFonts, fontCatalog) {
        var listedFonts = {};
        for (var fontKey in selectedFonts) {
            var selectedPsName = selectedFonts[fontKey];
            listedFonts[fontKey] = (selectedPsName && fontCatalog.familyByPsName[selectedPsName]) ? selectedPsName : null;
        }
        var kanjiPsName = listedFonts.kanji || DEFAULT_KANJI_FONT;
        if (!fontCatalog.familyByPsName[kanjiPsName]) kanjiPsName = fontCatalog.stylesByFamily[fontCatalog.families[0]][0].psName;
        var romanPsName = listedFonts.roman || DEFAULT_ROMAN_FONT;
        if (!fontCatalog.familyByPsName[romanPsName]) romanPsName = kanjiPsName;
        return {
            kanji: kanjiPsName,
            kana: listedFonts.kana || kanjiPsName, /* かなが見つからなければ漢字と同じ / falls back to the Kanji font */
            roman: romanPsName
        };
    }

    /**
     * かな・欧文のサイズとベースラインの初期値（％）を、選択中のテキストの和文との差から求める
     * 文字サイズが違えばその比率（例 10pt と 10.8pt → 108%）、ベースラインシフトが違えばその差を和文のサイズに対する比率に
     * （例 10pt で 1pt 上 → 10%）する。小数1桁に丸める
     * @param {Object} selectedText - getSelectedTextFonts() の結果
     * @returns {{sizes: Object, baselines: Object}} kanji / kana / roman ごとの値
     */
    function getInitialAdjustments(selectedText) {
        var sizes = { kanji: 100, kana: 100, roman: 100 };
        var baselines = { kanji: 0, kana: 0, roman: 0 };
        var kanjiSize = selectedText.sizes.kanji;
        if (!(kanjiSize > 0)) return { sizes: sizes, baselines: baselines };
        var kanjiShift = selectedText.baselineShifts.kanji || 0;
        var adjustKeys = ["kana", "roman"];
        for (var i = 0; i < adjustKeys.length; i++) {
            var groupSize = selectedText.sizes[adjustKeys[i]];
            if (groupSize > 0 && groupSize !== kanjiSize) sizes[adjustKeys[i]] = Math.round(groupSize / kanjiSize * 1000) / 10;
            var groupShift = selectedText.baselineShifts[adjustKeys[i]];
            if (groupShift !== null && groupShift !== kanjiShift) {
                baselines[adjustKeys[i]] = Math.round((groupShift - kanjiShift) / kanjiSize * 1000) / 10;
            }
        }
        return { sizes: sizes, baselines: baselines };
    }

    /**
     * 合成フォントを作成する
     * @returns {void}
     */
    function main() {
        /* 選択の読み取りはフォント一覧を作る前に済ませる（DOM を読むと一覧の文字列が壊れて落ちたため）
           Read the selection before building the font list; reading the DOM afterwards corrupted the list and crashed */
        var selectedText = getSelectedTextFonts();
        var fontCatalog = collectFontFamilies();
        if (fontCatalog.families.length === 0) return;

        var initialAdjustments = getInitialAdjustments(selectedText);
        /* 特例文字の読み込み元（特例文字を持つ合成フォント）/ composite fonts that carry custom sets */
        var targetFolder = findCompositeFontFolder();
        var customSetSources = collectCustomSetSources(targetFolder);

        var dialogState = {
            name: "",
            isAutoName: true,
            weight: "",
            isAutoWeight: true,
            fonts: getInitialFonts(selectedText.fonts, fontCatalog),
            sizes: initialAdjustments.sizes,
            baselines: initialAdjustments.baselines,
            sourceIndex: 0,
            isKanaLinked: selectedText.isKanaLinked,
            hasInDesign: getInDesignSpecifier() !== null,
            forInDesign: true /* InDesign があれば初期オン / on by default when InDesign is installed */
        };
        var fontSpec;
        do {
            fontSpec = showCompositeFontDialog(fontCatalog, dialogState, customSetSources);
        } while (fontSpec && fontSpec.reload);
        if (!fontSpec) return;

        if (!targetFolder) {
            alert(getLabel("alert.noFolder"));
            targetFolder = Folder.selectDialog(getLabel("alert.chooseFolder"));
            if (!targetFolder) return;
        }
        /* 書き出しの失敗（権限など）は File が例外で返すので、ここで受けて知らせる / file errors surface as exceptions */
        try {
            if (!writeCompositeFontFile(targetFolder, fontSpec)) return;
        } catch (e) {
            alert(e.message);
            return;
        }
        var doneMessage = getLabel("alert.done").replace("%1", fontSpec.name).replace("%2", targetFolder.fsName);
        if (fontSpec.forInDesign) doneMessage += "\n\n" + createInDesignCompositeFont(fontSpec);
        alert(doneMessage);
    }

    main();

})();
