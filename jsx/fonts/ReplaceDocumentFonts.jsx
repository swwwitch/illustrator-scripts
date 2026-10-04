#target illustrator
#targetengine "ReplaceDocumentFontsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント（または選択範囲）で使用中のフォントをファミリー／スタイル単位で一覧し、
選んだフォントを別のフォントへまとめて置き換えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceDocumentFonts.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ncc9330ba1f7d

### Overview

Lists the fonts used in the document (or the selection) by family and style,
and replaces the selected ones with another font in a single pass.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceDocumentFonts.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceDocumentFonts";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.2.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-03-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceDocumentFonts.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceDocumentFonts.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ncc9330ba1f7d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* スタイル行の字下げ / Indent used for style rows */
    var STYLE_ROW_INDENT = "　　";

    /* PostScript名表示の初期状態 / Initial state of the PostScript-name display */
    var SHOW_POSTSCRIPT_NAME_DEFAULT = false;

    /* ソートの初期状態（"name"：名前順／"countDesc"：使用数の降順／"countAsc"：使用数の昇順）/ Initial sort order ("name", "countDesc" or "countAsc") */
    var SORT_MODE_DEFAULT = "name";

    /* 文字・段落スタイルのフォントも置換するかの初期状態 / Initial state of replacing fonts in character and paragraph styles */
    var REPLACE_STYLE_FONTS_DEFAULT = true;

    /* 置換先にドキュメントで使っていないスタイルも出すかの初期状態（記憶せず毎回これで開く）/ Initial state of listing unused styles in the target list (not remembered; every run starts here) */
    var SHOW_ALL_TARGET_STYLES_DEFAULT = false;

    /* 置換元を合成フォントだけに絞るかの初期状態（記憶せず毎回これで開く）/ Initial state of listing only composite fonts as sources (not remembered; every run starts here) */
    var COMPOSITE_FONTS_ONLY_DEFAULT = false;

    /* リストの幅をフォント名に合わせるかの初期状態（OFF は固定幅の簡易表示）/ Initial state of fitting the list width to the font names (off: compact fixed width) */
    var FIT_LIST_WIDTH_DEFAULT = false;

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

    var LIST_LABEL_SPACING    = 6;    /* 見出しとリストの間隔 / spacing between a label and its list */
    var LIST_TOP_MARGIN       = 10;   /* リストの行の上の余白 / top margin above the list row */
    var LIST_BOTTOM_MARGIN    = 10;   /* リストと下のパネルのあいだに足す余白 / extra space between a list and the panel below */
    var LISTBOX_HEIGHT        = 300;  /* リストの高さ / list height */
    var LISTBOX_WIDTH_MIN     = 200;  /* リスト幅の下限 / minimum list width */
    var LISTBOX_WIDTH_MAX     = 600;  /* リスト幅の上限 / maximum list width */
    var LISTBOX_WIDTH_COMPACT = 240;  /* 簡易表示のリスト幅 / list width in the compact view */
    var LISTBOX_CHAR_WIDTH    = 9;    /* 1文字あたりの概算幅 / approximate width per character */
    var LISTBOX_WIDTH_PADDING = 5;    /* リスト幅の余裕 / extra width added to the list */

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

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^\uFEFF/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

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

    var LABELS = {
        dialog: {
            title: { ja: "ドキュメントフォントを置換", en: "Replace Document Fonts" }
        },
        fieldLabel: {
            sourceFonts: { ja: "置換元フォント（複数選択可）", en: "Source Fonts (Multiple Selection)" },
            targetFont: { ja: "置換先フォント", en: "Target Font" },
            sort: { ja: "並び順", en: "Sort" },
            scope: { ja: "対象", en: "Scope" }
        },
        checkbox: {
            postScriptName: { ja: "フォント名をPostScript名で表示", en: "Show font names as PostScript names" },
            fitListWidth: { ja: "リストの幅をフォント名に合わせる", en: "Fit list width to font names" },
            compositeFontsOnly: { ja: "合成フォントのみ", en: "Composite fonts only" },
            replaceStyleFonts: { ja: "文字・段落スタイルも置換", en: "Also replace in styles" },
            showAllTargetStyles: { ja: "使っていないフォントスタイルも表示", en: "Show unused font styles" }
        },
        panel: {
            options: { ja: "表示", en: "Display" },
            replaceOptions: { ja: "置換オプション", en: "Replace Options" }
        },
        dropdown: {
            sortMode: {
                name: { ja: "名前順", en: "Name" },
                countDesc: { ja: "使用数の多い順", en: "Most Used" },
                countAsc: { ja: "使用数の少ない順", en: "Least Used" }
            }
        },
        radio: {
            scopeDocument: { ja: "ドキュメント全体", en: "Entire Document" },
            scopeSelection: { ja: "選択範囲のみ", en: "Selection Only" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" },
            replaceAll: { ja: "すべて置換", en: "Replace All" },
            replace: { ja: "フォントを置換", en: "Replace Fonts" }
        },
        tooltip: {
            sourceFonts: {
                ja: "置換元のフォントを選びます。ファミリー名の行を選ぶと、そのファミリーのスタイルがすべて選ばれます。（ ）内は、そのフォントを使っているテキストオブジェクトの数です。",
                en: "Pick the fonts to replace. Selecting a family row selects every style in that family. The number in parentheses is how many text objects use the font."
            },
            targetFont: {
                ja: "置換先のフォントを選びます。ファミリー名の行は選べません。（ ）が付いていないスタイルは、ドキュメントで使っていないものです。",
                en: "Pick the font to replace them with. Family rows cannot be selected. Styles without a number in parentheses are not used in the document."
            },
            postScriptName: {
                ja: "ファミリー名とスタイル名の代わりに、PostScript名で一覧します。",
                en: "List the fonts by PostScript name instead of family and style."
            },
            fitListWidth: {
                ja: "フォント名が見切れないよう、リストの幅をいちばん長い名前に合わせます。OFFのときは固定幅で表示します。",
                en: "Widen the lists to the longest font name so no name is cut off. When off, the lists use a fixed width."
            },
            compositeFontsOnly: {
                ja: "置換元と置換先のリストに、合成フォントだけを並べます。［すべて置換］も合成フォントだけが対象になります。合成フォントを使っていないときは選べません。",
                en: "List only composite fonts in both the source and target lists. Replace All then affects composite fonts only. Unavailable when no composite font is in use."
            },
            scope: {
                ja: "置換する範囲を選びます。",
                en: "Choose which text to replace in."
            },
            scopeDocument: {
                ja: "ドキュメント内のすべてのテキストを対象にします。",
                en: "Target all text in the document."
            },
            scopeSelection: {
                ja: "実行時に選択していたテキストだけを対象にします（グループ内も含む）。テキストを選択していないときは選べません。",
                en: "Target only the text that was selected when the script started (including inside groups). Unavailable when no text is selected."
            },
            replaceStyleFonts: {
                ja: "置換元フォントが設定された文字スタイル・段落スタイルも、置換先フォントに書き換えます。［ドキュメント全体］のときだけ使えます。",
                en: "Also switch character and paragraph styles that use a source font to the target font. Available only with Entire Document."
            },
            showAllTargetStyles: {
                ja: "置換先のリストに、同じファミリーでドキュメントに使っていないスタイルも並べます（件数なしで表示）。",
                en: "Also list the styles of each family that the document does not use (shown without a count)."
            },
            sort: {
                ja: "リストの並び順を選びます。名前順はファミリー名、スタイル名の順に並べます。使用数で並べるとき、ファミリーはスタイルの使用数の合計で並べます。",
                en: "Choose the list order. Name sorts by family, then by style. When sorting by count, families are ordered by the total count of their styles."
            },
            replaceAll: {
                ja: "使用中のすべてのフォントを置換先フォントに置き換えます。置換先を選んでいないときは、置換元の1つ目のフォントにそろえます。",
                en: "Replace every font in use with the target font. With no target selected, the first source font is used instead."
            },
            replace: {
                ja: "選んだ置換元フォントを、置換先フォントに置き換えます。",
                en: "Replace the selected source fonts with the target font."
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noFontsFound: {
                ja: "ドキュメント内に使用中のフォントが見つかりません。",
                en: "No fonts in use were found in the document."
            },
            noSourceFont: {
                ja: "置換元フォントを1つ以上選択してください。",
                en: "Please select at least one source font."
            },
            noTargetFont: {
                ja: "置換先フォントを選択してください。",
                en: "Please select a target font."
            },
            selectFonts: {
                ja: "置換先フォントを選択してください。置換元だけを選んだときは、1つ目の置換元フォントにそろえます。",
                en: "Please select a target font. With only source fonts selected, everything is unified on the first source font."
            },
            targetNotFound: {
                ja: "フォントが見つかりません（%1）。",
                en: "Font not found (%1)."
            }
        }
    };

    /**
     * 件数を括弧で囲む（日本語は全角括弧、英語は半角括弧）
     * @param {number} count - 件数
     * @returns {string} 括弧付きの件数
     */
    function countSuffix(count) {
        return (uiLang === "ja") ? "（" + count + "）" : " (" + count + ")";
    }

    // =========================================
    // 状態 / State
    // =========================================

    var doc = null;
    var mainDialog = null;
    var sourceFontListBox = null;
    var targetFontListBox = null;
    var postScriptNameCheckbox = null;
    var fitListWidthCheckbox = null;
    var scopeDocumentRadio = null;
    var scopeSelectionRadio = null;

    /* ソートの種類（ポップアップの並び順）/ Sort modes, in popup order */
    var SORT_MODES = ["name", "countDesc", "countAsc"];

    /* ソートキーの桁数と、降順キーを作るときの上限 / Digits in a sort key, and the ceiling used to build descending keys */
    var SORT_KEY_DIGITS = 9;
    var SORT_COUNT_CEILING = 999999999;

    /* 設定の保存先（再起動しても残す）/ Settings store kept across restarts */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");

    /* 画面の状態（ラジオ・チェックの値は表示前に読み戻せないので控えておく）/ UI state, kept here because radio / checkbox values cannot be read back before show() */
    var currentSortMode = SORT_MODE_DEFAULT;
    var showsPostScriptName = SHOW_POSTSCRIPT_NAME_DEFAULT;
    var currentScope = "document";
    var replacesStyleFonts = REPLACE_STYLE_FONTS_DEFAULT;
    var showsAllTargetStyles = SHOW_ALL_TARGET_STYLES_DEFAULT;
    var fitsListWidth = FIT_LIST_WIDTH_DEFAULT;
    var showsCompositeFontsOnly = COMPOSITE_FONTS_ONLY_DEFAULT;
    var compositeFontsOnlyCheckbox = null;
    var showAllTargetStylesCheckbox = null;

    /* 実行時に選択していたテキストフレーム / Text frames selected at launch */
    var selectedTextFrames = [];
    var replaceStyleFontsCheckbox = null;

    /* 置換元リストに並べるフォント（ファミリー見出しを含む）/ Fonts listed in the source box, family headers included */
    var flatFontList = [];

    /* 置換先リストに並べるフォント（使っていないスタイルを含むことがある）/ Fonts listed in the target box, possibly with unused styles */
    var targetFontList = [];

    /* 最後に集めた使用中フォント（置換先だけ作り直すときに使う）/ Last collected fonts in use, reused when only the target list is rebuilt */
    var lastUsedFontMap = {};

    /* インストール済みフォントのファミリー別一覧（初めて要るときに作る）/ Installed fonts by family, built on first use */
    var installedFontsByFamily = null;

    /* フォント名ごとのTextRange / Text ranges tagged with their font name */
    var textRangeList = [];

    /* 選択の再帰更新を防ぐフラグ / Guard against recursive selection updates */
    var isUpdatingSelection = false;

    // =========================================
    // フォントの収集 / Collecting fonts
    // =========================================

    /**
     * 実際に使う対象を返す（選択範囲にテキストが無ければドキュメント全体）
     * @returns {string} "document" または "selection"
     */
    function getEffectiveScope() {
        return (currentScope === "selection" && selectedTextFrames.length > 0) ? "selection" : "document";
    }

    /**
     * 選択からテキストフレームを集める（グループ・クリップグループの中もたどる。文字を編集中ならそのストーリーのフレーム）
     * @param {*} selectionItems - doc.selection
     * @returns {Array<TextFrame>} 重複のないテキストフレームの配列
     */
    function collectSelectedTextFrames(selectionItems) {
        var textFrames = [];
        if (!selectionItems) return textFrames;

        /* 文字の選択中は TextRange が1つ返る / While editing text, the selection is a single TextRange */
        if (selectionItems.typename === "TextRange") {
            var storyFrames = selectionItems.story.textFrames;
            for (var s = 0; s < storyFrames.length; s++) textFrames.push(storyFrames[s]);
            return textFrames;
        }

        /**
         * ページアイテムをたどってテキストフレームを足す
         * @param {PageItem} pageItem - たどるアイテム
         * @returns {void}
         */
        function addTextFramesIn(pageItem) {
            if (pageItem.typename === "TextFrame") {
                for (var k = 0; k < textFrames.length; k++) {
                    if (textFrames[k] === pageItem) return;
                }
                textFrames.push(pageItem);
                return;
            }
            if (pageItem.typename !== "GroupItem") return;
            for (var c = 0; c < pageItem.pageItems.length; c++) addTextFramesIn(pageItem.pageItems[c]);
        }

        for (var i = 0; i < selectionItems.length; i++) addTextFramesIn(selectionItems[i]);
        return textFrames;
    }

    /**
     * ドキュメント内のTextRangeを走査し、使用中フォントをファミリー別に集める
     * @returns {object} ファミリー名 → フォント名 → {name, style, family, frameCount} のマップ
     */
    function collectUsedFonts() {
        var usedFontMap = {};
        textRangeList = [];

        var textFrames = (getEffectiveScope() === "selection") ? selectedTextFrames : doc.textFrames;
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var isLocked = textFrame.locked;
            var isHidden = textFrame.hidden;
            var ranges = textFrame.textRanges;
            var fontsInFrame = {};

            for (var j = 0; j < ranges.length; j++) {
                var range = ranges[j];
                if (range.length === 0) continue;

                /* 無効な範囲はフォントを取得できないのでスキップ / Skip ranges whose font cannot be read */
                var font = null;
                try {
                    font = range.characterAttributes.textFont;
                } catch (e) {
                    continue;
                }
                if (!font) continue;

                textRangeList.push({
                    range: range,
                    fontName: font.name,
                    isLocked: isLocked,
                    isHidden: isHidden
                });

                /* 同じTextFrame内では1フォントにつき1回だけ数える / Count a font once per text frame */
                if (fontsInFrame[font.name]) continue;
                fontsInFrame[font.name] = true;

                if (!usedFontMap[font.family]) usedFontMap[font.family] = {};
                if (!usedFontMap[font.family][font.name]) {
                    usedFontMap[font.family][font.name] = {
                        name: font.name,
                        style: font.style,
                        family: font.family,
                        frameCount: 1
                    };
                } else {
                    usedFontMap[font.family][font.name].frameCount++;
                }
            }
        }
        return usedFontMap;
    }

    /**
     * フォントやファミリーの配列を、選択中のソートで並べ替える
     * @param {Array<object>} entries - frameCount と名前を持つ要素の配列
     * @param {string} nameKey - 名前順に使うプロパティ名（"family"・"style"・"name"）
     * @returns {void}
     */
    function sortFontEntries(entries, nameKey) {
        /* 比較関数は使わず、並び順を表す文字列キーを引数なしの sort() で並べる / Sort string keys with a plain sort() instead of a comparator */
        var sortKeys = [];
        for (var i = 0; i < entries.length; i++) {
            var countKey = "";
            if (currentSortMode === "countDesc") countKey = padNumber(SORT_COUNT_CEILING - entries[i].frameCount);
            if (currentSortMode === "countAsc") countKey = padNumber(entries[i].frameCount);

            /* 使用数が同じときは名前順、名前も同じなら元の順 / Ties fall back to the name, then to the original order */
            sortKeys.push(countKey + "\u0001" + String(entries[i][nameKey]).toLowerCase() + "\u0001" + padNumber(i));
        }
        sortKeys.sort();

        var sortedEntries = [];
        for (var k = 0; k < sortKeys.length; k++) {
            var keyParts = sortKeys[k].split("\u0001");
            sortedEntries.push(entries[parseInt(keyParts[keyParts.length - 1], 10)]);
        }
        for (var n = 0; n < sortedEntries.length; n++) {
            entries[n] = sortedEntries[n];
        }
    }

    /**
     * 数値を桁をそろえた文字列にする（文字列の並びを数値の並びと一致させる）
     * @param {number} value - 0 以上 SORT_COUNT_CEILING 以下の整数
     * @returns {string} 先頭を0で埋めた文字列
     */
    function padNumber(value) {
        var padded = String(value);
        while (padded.length < SORT_KEY_DIGITS) padded = "0" + padded;
        return padded;
    }

    /**
     * 収集したフォントをリスト表示用の1次元配列にする
     * @param {object} usedFontMap - collectUsedFonts() が返したマップ
     * @returns {Array<object>} ファミリー見出しとスタイル行を並べた配列（PostScript名表示中は見出しなしの1フォント1行）
     */
    function buildFlatFontList(usedFontMap) {
        var fontList = [];
        var allFonts = [];
        var families = [];

        for (var family in usedFontMap) {
            var styles = usedFontMap[family];
            var familyEntry = { family: family, fonts: [], frameCount: 0 };
            for (var fontName in styles) {
                familyEntry.fonts.push(styles[fontName]);
                familyEntry.frameCount += styles[fontName].frameCount;
                allFonts.push(styles[fontName]);
            }
            families.push(familyEntry);
        }

        /* PostScript名表示のときは見出しを立てず、1フォント1行で並べる / In PostScript-name mode, list one row per font with no headers */
        if (showsPostScriptName) {
            sortFontEntries(allFonts, "name");
            for (var p = 0; p < allFonts.length; p++) {
                fontList.push(createFontRow(allFonts[p].name, allFonts[p]));
            }
            return fontList;
        }

        sortFontEntries(families, "family");
        for (var f = 0; f < families.length; f++) {
            var fonts = families[f].fonts;

            /* スタイルが1つだけのファミリーは見出しを立てず1行で見せる / Show single-style families on one row */
            if (fonts.length === 1) {
                fontList.push(createFontRow(fonts[0].family + " " + fonts[0].style, fonts[0]));
                continue;
            }

            /* 見出しにはスタイルの使用数の合計を添える / Show the total count of the styles on the header */
            fontList.push({ label: families[f].family + countSuffix(families[f].frameCount), family: families[f].family, isHeader: true });
            sortFontEntries(fonts, "style");
            for (var i = 0; i < fonts.length; i++) {
                fontList.push(createFontRow(STYLE_ROW_INDENT + fonts[i].style, fonts[i]));
            }
        }
        return fontList;
    }

    /**
     * 合成フォントかを返す（合成フォントの名前は「ATC-」＋合成フォント名の16進表記）
     * @param {string} fontName - フォント名（PostScript名）
     * @returns {boolean} 合成フォントなら true
     */
    function isCompositeFontName(fontName) {
        return fontName.indexOf("ATC-") === 0;
    }

    /**
     * 使用中フォントに合成フォントが含まれるかを返す
     * @param {object} usedFontMap - collectUsedFonts() が返したマップ
     * @returns {boolean} 1つでもあれば true
     */
    function hasCompositeFont(usedFontMap) {
        for (var family in usedFontMap) {
            for (var fontName in usedFontMap[family]) {
                if (isCompositeFontName(fontName)) return true;
            }
        }
        return false;
    }

    /**
     * ［合成フォントのみ］が ON なら、使用中フォントを合成フォントだけに絞る
     * @param {object} usedFontMap - collectUsedFonts() が返したマップ
     * @returns {object} 同じ形のマップ（OFF のときは渡したマップそのもの）
     */
    function filterCompositeFonts(usedFontMap) {
        if (!showsCompositeFontsOnly) return usedFontMap;

        var compositeFontMap = {};
        for (var family in usedFontMap) {
            for (var fontName in usedFontMap[family]) {
                if (!isCompositeFontName(fontName)) continue;
                if (!compositeFontMap[family]) compositeFontMap[family] = {};
                compositeFontMap[family][fontName] = usedFontMap[family][fontName];
            }
        }
        return compositeFontMap;
    }

    /**
     * 置換元リストの並びを作る（オプションが ON なら合成フォントだけ）
     * @param {object} usedFontMap - collectUsedFonts() が返したマップ
     * @returns {Array<object>} 置換元リストに並べる配列
     */
    function buildSourceFontList(usedFontMap) {
        return buildFlatFontList(filterCompositeFonts(usedFontMap));
    }

    /**
     * ［合成フォントのみ］を、合成フォントを使っているときだけ選べるようにする（使っていなければ OFF に戻す）
     * @returns {void}
     */
    function updateCompositeFontsOnlyAvailability() {
        var isAvailable = hasCompositeFont(lastUsedFontMap);
        if (!isAvailable) showsCompositeFontsOnly = false;
        if (!compositeFontsOnlyCheckbox) return;
        compositeFontsOnlyCheckbox.enabled = isAvailable;
        compositeFontsOnlyCheckbox.value = showsCompositeFontsOnly;
    }

    /**
     * 置換先リストの並びを作る（オプションが ON なら合成フォントだけにし、使っていないスタイルも足す）
     * @param {object} usedFontMap - collectUsedFonts() が返したマップ
     * @returns {Array<object>} 置換先リストに並べる配列
     */
    function buildTargetFontList(usedFontMap) {
        usedFontMap = filterCompositeFonts(usedFontMap);
        if (!showsAllTargetStyles) return buildFlatFontList(usedFontMap);

        var installedFonts = getInstalledFontsByFamily();
        var extendedFontMap = {};
        for (var family in usedFontMap) {
            extendedFontMap[family] = {};
            for (var fontName in usedFontMap[family]) {
                extendedFontMap[family][fontName] = usedFontMap[family][fontName];
            }
            var familyFonts = installedFonts[family] || [];
            for (var i = 0; i < familyFonts.length; i++) {
                if (extendedFontMap[family][familyFonts[i].name]) continue;
                extendedFontMap[family][familyFonts[i].name] = {
                    name: familyFonts[i].name,
                    style: familyFonts[i].style,
                    family: family,
                    frameCount: 0,
                    isUnused: true
                };
            }
        }
        return buildFlatFontList(extendedFontMap);
    }

    /**
     * インストール済みフォントをファミリー別にまとめる（1回だけ作って使い回す）
     * @returns {object} ファミリー名 → [{name, style}] のマップ
     */
    function getInstalledFontsByFamily() {
        if (installedFontsByFamily) return installedFontsByFamily;

        installedFontsByFamily = {};
        var textFonts = app.textFonts;
        for (var i = 0; i < textFonts.length; i++) {
            var textFont = textFonts[i];
            var family = textFont.family;
            if (!installedFontsByFamily[family]) installedFontsByFamily[family] = [];
            installedFontsByFamily[family].push({ name: textFont.name, style: textFont.style });
        }
        return installedFontsByFamily;
    }

    /**
     * リストの1行分（見出し以外）を作る
     * @param {string} labelBase - 件数を付ける前の表示名
     * @param {object} font - collectUsedFonts() のフォント情報
     * @returns {object} リストの1行分
     */
    function createFontRow(labelBase, font) {
        return {
            label: font.isUnused ? labelBase : labelBase + countSuffix(font.frameCount),
            name: font.name,
            family: font.family,
            style: font.style,
            isHeader: false
        };
    }

    /**
     * フォント一覧を集め直してリストを作り直す（選択はフォント名で引き継ぐ）
     * @returns {void}
     */
    function refreshFontList() {
        var previousSourceFontNames = getSelectedSourceFontNames();
        var previousTargetFontName = getSelectedTargetFontName();

        lastUsedFontMap = collectUsedFonts();
        updateCompositeFontsOnlyAvailability();
        flatFontList = buildSourceFontList(lastUsedFontMap);
        targetFontList = buildTargetFontList(lastUsedFontMap);
        populateFontListBoxes();
        restoreSelection(previousSourceFontNames, previousTargetFontName);
    }

    // =========================================
    // 置換処理 / Replacing fonts
    // =========================================

    /**
     * フォント名から TextFont を取得する
     * @param {string} fontName - フォント名（PostScript名）
     * @returns {TextFont|null} 見つからなければ null
     */
    function findFontByName(fontName) {
        try {
            return app.textFonts.getByName(fontName);
        } catch (e) {
            return null;
        }
    }

    /**
     * 置換元として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {Array<string>} フォント名の配列
     */
    function getSelectedSourceFontNames() {
        var fontNames = [];
        if (!sourceFontListBox || !sourceFontListBox.selection) return fontNames;

        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var listEntry = flatFontList[sourceFontListBox.selection[i].index];
            if (!listEntry.isHeader) fontNames.push(listEntry.name);
        }
        return fontNames;
    }

    /**
     * 置換先リストでフォント（見出し行以外）が選ばれているか調べる
     * @returns {boolean} 選ばれていれば true
     */
    function hasTargetFontSelection() {
        if (!targetFontListBox || !targetFontListBox.selection) return false;
        return !targetFontList[targetFontListBox.selection.index].isHeader;
    }

    /**
     * 置換先として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {string} 未選択・見出し行のときは空文字列
     */
    function getSelectedTargetFontName() {
        if (!hasTargetFontSelection()) return "";
        return targetFontList[targetFontListBox.selection.index].name;
    }

    /**
     * 置換先として選択されているフォントを取り出す
     * @returns {TextFont|null} 未選択・見出し行・未インストールの場合は null
     */
    function getSelectedTargetFont() {
        var fontName = getSelectedTargetFontName();
        if (fontName === "") return null;

        var targetFont = findFontByName(fontName);
        if (!targetFont) alert(getLabel(LABELS.alert.targetNotFound, [fontName]));
        return targetFont;
    }

    /**
     * 使用中フォントの名前をすべて集める
     * @param {string} [excludedFontName] - 除外するフォント名
     * @returns {Array<string>} フォント名の配列
     */
    function collectAllFontNames(excludedFontName) {
        var fontNames = [];
        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;
            if (flatFontList[i].name === excludedFontName) continue;
            fontNames.push(flatFontList[i].name);
        }
        return fontNames;
    }

    /**
     * 指定したフォントを置換先フォントに置き換える（ロック・非表示は対象外）
     * @param {Array<string>} sourceFontNames - 置換元のフォント名
     * @param {TextFont} targetFont - 置換先フォント
     * @returns {void}
     */
    function replaceFonts(sourceFontNames, targetFont) {
        if (!targetFont || sourceFontNames.length === 0) return;

        for (var i = 0; i < textRangeList.length; i++) {
            var entry = textRangeList[i];
            if (entry.isLocked || entry.isHidden) continue;

            for (var j = 0; j < sourceFontNames.length; j++) {
                if (entry.fontName === sourceFontNames[j]) {
                    entry.range.characterAttributes.textFont = targetFont;
                    break;
                }
            }
        }

        /* スタイルは範囲の外のテキストにも効くので、ドキュメント全体のときだけ / Styles affect text outside the scope, so only for the whole document */
        if (replacesStyleFonts && getEffectiveScope() === "document") {
            replaceStyleFonts(doc.characterStyles, sourceFontNames, targetFont);
            replaceStyleFonts(doc.paragraphStyles, sourceFontNames, targetFont);
        }
        refreshFontList();
        app.redraw();
    }

    /**
     * スタイルのうち、置換元フォントが設定されているものを置換先フォントに書き換える
     * @param {CharacterStyles|ParagraphStyles} styles - 文字スタイルまたは段落スタイルの集まり
     * @param {Array<string>} sourceFontNames - 置換元のフォント名
     * @param {TextFont} targetFont - 置換先のフォント
     * @returns {void}
     */
    function replaceStyleFonts(styles, sourceFontNames, targetFont) {
        for (var i = 0; i < styles.length; i++) {
            /* フォントを設定していないスタイルは読むと例外になる / Reading a style with no font set throws */
            var styleFontName = null;
            try {
                styleFontName = styles[i].characterAttributes.textFont.name;
            } catch (e) {
                continue;
            }
            for (var j = 0; j < sourceFontNames.length; j++) {
                if (styleFontName === sourceFontNames[j]) {
                    styles[i].characterAttributes.textFont = targetFont;
                    break;
                }
            }
        }
    }

    // =========================================
    // イベントハンドラー / Event handlers
    // =========================================

    /**
     * 置換元リストの選択を整える（見出し行はファミリー内の全スタイルに展開）
     * @returns {void}
     */
    function handleSourceFontSelection() {
        if (isUpdatingSelection || !sourceFontListBox.selection) return;

        /* 見出し行はファミリー内の全スタイルに展開する / Expand a family header to all its styles */
        var expandedSelection = [];
        var headerIndex = -1;
        var lastStyleIndex = -1;
        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var selectedIndex = sourceFontListBox.selection[i].index;
            var listEntry = flatFontList[selectedIndex];

            if (!listEntry.isHeader) {
                expandedSelection.push(sourceFontListBox.items[selectedIndex]);
                continue;
            }
            if (headerIndex === -1) headerIndex = selectedIndex;
            for (var j = 0; j < flatFontList.length; j++) {
                if (!flatFontList[j].isHeader && flatFontList[j].family === listEntry.family) {
                    expandedSelection.push(sourceFontListBox.items[j]);
                    if (headerIndex === selectedIndex) lastStyleIndex = j;
                }
            }
        }

        /* 見出しを含むときだけ選択を入れ直す（入れ直すとリストがスクロールする）/ Reassign only when a header was expanded, since reassigning scrolls the list */
        if (headerIndex !== -1) {
            isUpdatingSelection = true;
            sourceFontListBox.selection = expandedSelection;
            isUpdatingSelection = false;

            /* 動いた表示を、クリックしたファミリーが見える位置に戻す / Scroll back so the clicked family stays in view */
            if (lastStyleIndex !== -1) sourceFontListBox.revealItem(sourceFontListBox.items[lastStyleIndex]);
            sourceFontListBox.revealItem(sourceFontListBox.items[headerIndex]);
        }
    }

    /**
     * 置換先リストで見出し行が選ばれたら選択を解除する
     * @returns {void}
     */
    function handleTargetFontSelection() {
        if (isUpdatingSelection || !targetFontListBox.selection) return;
        if (targetFontList[targetFontListBox.selection.index].isHeader) {
            targetFontListBox.selection = null;
        }
    }

    /**
     * ［フォント名をPostScript名で表示］：表示形式を切り替えてリストを作り直す
     * @returns {void}
     */
    function handleDisplayModeChange() {
        showsPostScriptName = postScriptNameCheckbox.value;
        refreshFontList();
    }

    /**
     * ［合成フォントのみ］：2つのリストを作り直す（ドキュメントは走査し直さない）
     * @returns {void}
     */
    function handleCompositeFontsOnlyClick() {
        var previousSourceFontNames = getSelectedSourceFontNames();
        var previousTargetFontName = getSelectedTargetFontName();
        showsCompositeFontsOnly = compositeFontsOnlyCheckbox.value;
        flatFontList = buildSourceFontList(lastUsedFontMap);
        targetFontList = buildTargetFontList(lastUsedFontMap);
        populateFontListBoxes();
        restoreSelection(previousSourceFontNames, previousTargetFontName);
    }

    /**
     * option＋Tab：置換元リストと置換先リストのあいだでフォーカスを移す
     * @param {Object} keyEvent - keydown イベント
     * @returns {void}
     */
    function switchFontListFocus(keyEvent) {
        var nextListBox = (keyEvent.target === sourceFontListBox) ? targetFontListBox : sourceFontListBox;
        nextListBox.active = true;
    }

    /**
     * ［リストの幅をフォント名に合わせる］：リストの幅を変えてダイアログを配置し直す
     * @returns {void}
     */
    function handleFitListWidthClick() {
        fitsListWidth = fitListWidthCheckbox.value;
        var listBoxWidth = getListBoxWidth();
        var listBoxes = [sourceFontListBox, targetFontListBox];
        /* 配置済みのコントロールは size も入れないと幅が変わらない / A laid-out control keeps its size unless size is set too */
        for (var i = 0; i < listBoxes.length; i++) {
            listBoxes[i].preferredSize.width = listBoxWidth;
            listBoxes[i].size = [listBoxWidth, listBoxes[i].size.height];
        }
        mainDialog.layout.layout(true);

        /* 狭めたときはウィンドウが縮まないので、計算し直した大きさを入れて中身を合わせる / The window does not shrink on its own, so apply the recalculated size and refit the contents */
        mainDialog.size = [mainDialog.preferredSize.width, mainDialog.preferredSize.height];
        mainDialog.layout.resize();
    }

    /**
     * 対象のラジオ：対象を切り替えてリストを作り直す
     * @param {string} scope - "document" または "selection"
     * @returns {void}
     */
    function handleScopeChange(scope) {
        if (scope === currentScope) return;
        currentScope = scope;
        replaceStyleFontsCheckbox.enabled = (getEffectiveScope() === "document");
        refreshFontList();
    }

    /**
     * ［使っていないスタイルも表示］：置換先リストだけ作り直す（ドキュメントは走査し直さない）
     * @returns {void}
     */
    function handleShowAllTargetStylesClick() {
        var previousSourceFontNames = getSelectedSourceFontNames();
        var previousTargetFontName = getSelectedTargetFontName();
        showsAllTargetStyles = showAllTargetStylesCheckbox.value;
        targetFontList = buildTargetFontList(lastUsedFontMap);
        populateFontListBoxes();
        restoreSelection(previousSourceFontNames, previousTargetFontName);
    }

    /**
     * ［文字・段落スタイルも置換］：状態を控える
     * @returns {void}
     */
    function handleReplaceStyleFontsClick() {
        replacesStyleFonts = replaceStyleFontsCheckbox.value;
    }

    /**
     * ソートのポップアップ：並べ替えてリストを作り直す
     * @param {string} sortMode - "name"・"countDesc"・"countAsc"
     * @returns {void}
     */
    function handleSortModeChange(sortMode) {
        if (sortMode === currentSortMode) return;
        currentSortMode = sortMode;
        refreshFontList();
    }

    /**
     * ［フォントを置換］：選んだ置換元フォントを置換先フォントに置き換える
     * @returns {void}
     */
    function handleReplaceClick() {
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.noSourceFont));
            return;
        }
        if (!hasTargetFontSelection()) {
            alert(getLabel(LABELS.alert.noTargetFont));
            return;
        }
        replaceFonts(sourceFontNames, getSelectedTargetFont());
    }

    /**
     * ［全置換］：使用中のすべてのフォントを1つのフォントにそろえる
     * @returns {void}
     */
    function handleReplaceAllClick() {
        /* 置換先が選ばれていれば、それにすべてをそろえる / Unify on the target font when one is selected */
        var targetFont = getSelectedTargetFont();
        if (targetFont) {
            replaceFonts(collectAllFontNames(), targetFont);
            return;
        }

        /* 置換先がなければ、置換元の1つ目にそろえる / Otherwise unify on the first source font */
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.selectFonts));
            return;
        }

        var fallbackFont = findFontByName(sourceFontNames[0]);
        if (!fallbackFont) {
            alert(getLabel(LABELS.alert.targetNotFound, [sourceFontNames[0]]));
            return;
        }
        replaceFonts(collectAllFontNames(sourceFontNames[0]), fallbackFont);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見出し付きのフォントリストを1カラム分追加する
     * @param {Group} parent - 追加先のグループ
     * @param {object} labelSet - 見出しのラベル（ja/en）
     * @param {object} tooltipSet - ツールチップのラベル（ja/en）
     * @param {boolean} allowsMultiple - 複数選択を許可するか
     * @returns {{group: Group, listBox: ListBox}} カラムのグループ（下にパネルを足す先）と、追加したリストボックス
     */
    function addFontListColumn(parent, labelSet, tooltipSet, allowsMultiple) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];

        /* 見出しとリストは詰め、下のパネルとは LIST_BOTTOM_MARGIN だけ離す / Keep the label close to the list and push the panel below away */
        var listBlock = columnGroup.add("group");
        listBlock.orientation = "column";
        listBlock.alignChildren = ["fill", "top"];
        listBlock.spacing = LIST_LABEL_SPACING;
        listBlock.margins = [0, 0, 0, LIST_BOTTOM_MARGIN];
        listBlock.add("statictext", undefined, labelText(labelSet));

        var fontListBox = listBlock.add("listbox", undefined, [], { multiselect: allowsMultiple });
        fontListBox.preferredSize.height = LISTBOX_HEIGHT;
        fontListBox.tabEnabled = true;
        fontListBox.helpTip = getLabel(tooltipSet);
        return { group: columnGroup, listBox: fontListBox };
    }

    /**
     * ラジオボタンを1つ追加する
     * @param {Group} parent - 追加先のグループ（ラジオは同じ親の中だけ排他になる）
     * @param {boolean} isSelected - 最初から選んでおくか
     * @param {object} labelSet - ラベル（ja/en）
     * @param {object} tooltipSet - ツールチップ（ja/en）
     * @param {Function} onSelect - クリックされたときの処理
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addChoiceRadio(parent, isSelected, labelSet, tooltipSet, onSelect) {
        var choiceRadio = parent.add("radiobutton", undefined, getLabel(labelSet));
        choiceRadio.value = isSelected;
        choiceRadio.helpTip = getLabel(tooltipSet);
        choiceRadio.onClick = onSelect;
        return choiceRadio;
    }

    /**
     * ソートのポップアップを追加する
     * @param {Group} parent - 追加先のグループ
     * @returns {DropDownList} 追加したポップアップ
     */
    function addSortDropdown(parent) {
        var sortLabels = [];
        var selectedIndex = 0;
        for (var i = 0; i < SORT_MODES.length; i++) {
            sortLabels.push(getLabel(LABELS.dropdown.sortMode[SORT_MODES[i]]));
            if (SORT_MODES[i] === currentSortMode) selectedIndex = i;
        }
        var sortDropdown = parent.add("dropdownlist", undefined, sortLabels);
        sortDropdown.selection = selectedIndex;
        sortDropdown.helpTip = getLabel(LABELS.tooltip.sort);
        sortDropdown.onChange = function() {
            if (sortDropdown.selection) handleSortModeChange(SORT_MODES[sortDropdown.selection.index]);
        };
        return sortDropdown;
    }

    /**
     * 置換元リストの下に「表示」パネルを作る（PostScript名・リストの幅・合成フォントのみ・ソート）
     * @param {Group} parent - 置換元リストのカラム
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parent) {
        var optionPanel = parent.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionPanel, 6);

        postScriptNameCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.postScriptName));
        postScriptNameCheckbox.value = showsPostScriptName;
        postScriptNameCheckbox.helpTip = getLabel(LABELS.tooltip.postScriptName);

        fitListWidthCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.fitListWidth));
        fitListWidthCheckbox.value = fitsListWidth;
        fitListWidthCheckbox.helpTip = getLabel(LABELS.tooltip.fitListWidth);

        compositeFontsOnlyCheckbox = optionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.compositeFontsOnly));
        compositeFontsOnlyCheckbox.helpTip = getLabel(LABELS.tooltip.compositeFontsOnly);
        updateCompositeFontsOnlyAvailability();

        /* ソート（ポップアップ）/ Sort popup */
        var sortRow = optionPanel.add("group");
        sortRow.orientation = "row";
        sortRow.alignment = ["left", "top"];
        sortRow.alignChildren = ["left", "center"];
        var sortLabel = sortRow.add("statictext", undefined, labelText(LABELS.fieldLabel.sort));
        sortLabel.helpTip = getLabel(LABELS.tooltip.sort);
        addSortDropdown(sortRow);
        return optionPanel;
    }

    /**
     * ダイアログを組み立てる
     * @returns {void}
     */
    function buildDialog() {
        mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        /* 対象（ダイアログの左右中央）/ Scope row, centered in the dialog */
        var scopeRow = mainDialog.add("group");
        scopeRow.orientation = "row";
        scopeRow.alignment = ["center", "top"];
        scopeRow.alignChildren = ["left", "center"];
        var scopeLabel = scopeRow.add("statictext", undefined, labelText(LABELS.fieldLabel.scope));
        scopeLabel.helpTip = getLabel(LABELS.tooltip.scope);
        var effectiveScope = getEffectiveScope();
        scopeDocumentRadio = addChoiceRadio(scopeRow, effectiveScope === "document", LABELS.radio.scopeDocument, LABELS.tooltip.scopeDocument, function() {
            handleScopeChange("document");
        });
        scopeSelectionRadio = addChoiceRadio(scopeRow, effectiveScope === "selection", LABELS.radio.scopeSelection, LABELS.tooltip.scopeSelection, function() {
            handleScopeChange("selection");
        });
        scopeSelectionRadio.enabled = (selectedTextFrames.length > 0);

        var listGroup = mainDialog.add("group");
        listGroup.orientation = "row";
        listGroup.alignChildren = ["fill", "top"];
        listGroup.spacing = COLUMN_SPACING;
        listGroup.margins = [0, LIST_TOP_MARGIN, 0, 0];

        var sourceColumn = addFontListColumn(listGroup, LABELS.fieldLabel.sourceFonts, LABELS.tooltip.sourceFonts, true);
        sourceFontListBox = sourceColumn.listBox;
        addOptionPanel(sourceColumn.group);
        var targetColumn = addFontListColumn(listGroup, LABELS.fieldLabel.targetFont, LABELS.tooltip.targetFont, false);
        targetFontListBox = targetColumn.listBox;
        /* 置換オプション / Replace options */
        var replaceOptionPanel = targetColumn.group.add("panel", undefined, getLabel(LABELS.panel.replaceOptions));
        setupPanel(replaceOptionPanel, 6);

        showAllTargetStylesCheckbox = replaceOptionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.showAllTargetStyles));
        showAllTargetStylesCheckbox.value = showsAllTargetStyles;
        showAllTargetStylesCheckbox.helpTip = getLabel(LABELS.tooltip.showAllTargetStyles);

        /* スタイルは範囲の外にも効くので、ドキュメント全体のときだけ使える / Styles reach beyond the scope, so only for the whole document */
        replaceStyleFontsCheckbox = replaceOptionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.replaceStyleFonts));
        replaceStyleFontsCheckbox.value = replacesStyleFonts;
        replaceStyleFontsCheckbox.enabled = (effectiveScope === "document");
        replaceStyleFontsCheckbox.helpTip = getLabel(LABELS.tooltip.replaceStyleFonts);

        /* ボタンエリア（左端に閉じる、右側に置換系）/ Button row: Close on the far left, replace buttons on the right */
        var buttonRow = addButtonRow(mainDialog);
        var btnClose = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "cancel" });
        var btnReplaceAll = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replaceAll));
        var btnReplace = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replace), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        btnReplaceAll.helpTip = getLabel(LABELS.tooltip.replaceAll);
        btnReplace.helpTip = getLabel(LABELS.tooltip.replace);

        sourceFontListBox.onChange = handleSourceFontSelection;
        targetFontListBox.onChange = handleTargetFontSelection;
        postScriptNameCheckbox.onClick = handleDisplayModeChange;
        fitListWidthCheckbox.onClick = handleFitListWidthClick;
        compositeFontsOnlyCheckbox.onClick = handleCompositeFontsOnlyClick;
        replaceStyleFontsCheckbox.onClick = handleReplaceStyleFontsClick;
        showAllTargetStylesCheckbox.onClick = handleShowAllTargetStylesClick;
        btnReplaceAll.onClick = handleReplaceAllClick;
        btnReplace.onClick = handleReplaceClick;

        addKeyShortcuts(mainDialog, {
            /* どれもリストにフォーカスがあっても効かせる / All of these work while a list has focus */
            "Alt+D": { target: scopeDocumentRadio, inFields: true },
            "Alt+S": { target: scopeSelectionRadio, inFields: true },
            "Alt+P": { target: postScriptNameCheckbox, inFields: true },
            "Alt+L": { target: fitListWidthCheckbox, inFields: true },
            "Alt+C": { target: compositeFontsOnlyCheckbox, inFields: true },
            "Alt+Tab": { target: switchFontListFocus, inFields: true }
        }, { showInTip: true });
    }

    /**
     * 2つのリストにフォント一覧を流し込み、幅をそろえる
     * @returns {void}
     */
    function populateFontListBoxes() {
        var listBoxWidth = getListBoxWidth();

        isUpdatingSelection = true;
        sourceFontListBox.removeAll();
        targetFontListBox.removeAll();
        for (var i = 0; i < flatFontList.length; i++) {
            sourceFontListBox.add("item", flatFontList[i].label);
        }
        for (var t = 0; t < targetFontList.length; t++) {
            targetFontListBox.add("item", targetFontList[t].label);
        }
        isUpdatingSelection = false;

        sourceFontListBox.preferredSize.width = listBoxWidth;
        targetFontListBox.preferredSize.width = listBoxWidth;
    }

    /**
     * フォント名を手がかりに、作り直したリストの選択を元に戻す
     * @param {Array<string>} sourceFontNames - 置換元として選択されていたフォント名
     * @param {string} targetFontName - 置換先として選択されていたフォント名
     * @returns {void}
     */
    function restoreSelection(sourceFontNames, targetFontName) {
        var restoredSelection = [];
        var restoredTargetIndex = -1;

        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;

            for (var j = 0; j < sourceFontNames.length; j++) {
                if (flatFontList[i].name === sourceFontNames[j]) {
                    restoredSelection.push(sourceFontListBox.items[i]);
                    break;
                }
            }
        }
        for (var t = 0; t < targetFontList.length; t++) {
            if (!targetFontList[t].isHeader && targetFontList[t].name === targetFontName) {
                restoredTargetIndex = t;
                break;
            }
        }

        isUpdatingSelection = true;
        sourceFontListBox.selection = restoredSelection;
        targetFontListBox.selection = (restoredTargetIndex === -1) ? null : restoredTargetIndex;
        isUpdatingSelection = false;
    }

    /**
     * 今の表示方法でのリストの幅を返す（簡易表示は固定幅、合わせるときは長い方のリストに）
     * @returns {number} リストの幅（px）
     */
    function getListBoxWidth() {
        if (!fitsListWidth) return LISTBOX_WIDTH_COMPACT;
        return Math.max(calculateListBoxWidth(flatFontList), calculateListBoxWidth(targetFontList));
    }

    /**
     * いちばん長いラベルからリストの幅を見積もる
     * @param {Array<object>} fontList - リストに並べるフォント
     * @returns {number} リストの幅（px）
     */
    function calculateListBoxWidth(fontList) {
        var maxLength = 0;
        for (var i = 0; i < fontList.length; i++) {
            if (fontList[i].label.length > maxLength) maxLength = fontList[i].label.length;
        }
        var estimatedWidth = maxLength * LISTBOX_CHAR_WIDTH + LISTBOX_WIDTH_PADDING;
        return Math.min(LISTBOX_WIDTH_MAX, Math.max(LISTBOX_WIDTH_MIN, estimatedWidth));
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        doc = app.activeDocument;

        loadSettings();
        selectedTextFrames = collectSelectedTextFrames(doc.selection);
        /* テキストを選択して実行したときは［選択範囲のみ］で開く / Open in Selection Only when text is selected at launch */
        currentScope = (selectedTextFrames.length > 0) ? "selection" : "document";

        lastUsedFontMap = collectUsedFonts();
        flatFontList = buildFlatFontList(lastUsedFontMap);

        /* 選択範囲にフォントが無ければドキュメント全体で集め直す / Fall back to the whole document when the selection has no fonts */
        if (flatFontList.length === 0 && getEffectiveScope() === "selection") {
            selectedTextFrames = [];
            lastUsedFontMap = collectUsedFonts();
            flatFontList = buildFlatFontList(lastUsedFontMap);
        }
        if (flatFontList.length === 0) {
            alert(getLabel(LABELS.alert.noFontsFound));
            return;
        }
        targetFontList = buildTargetFontList(lastUsedFontMap);

        buildDialog();
        populateFontListBoxes();

        /* 先頭のフォントを選んでおく / Preselect the first font */
        sourceFontListBox.selection = 0;
        targetFontListBox.selection = 0;
        handleSourceFontSelection();

        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        mainDialog.show();
        saveSettings();
    }

    /**
     * 保存した設定を読み込み、画面の状態に反映する
     * @returns {void}
     */
    function loadSettings() {
        var savedSettings = settingsStore.load({
            showPostScriptName: SHOW_POSTSCRIPT_NAME_DEFAULT,
            sortMode: SORT_MODE_DEFAULT,
            replaceStyleFonts: REPLACE_STYLE_FONTS_DEFAULT,
            fitListWidth: FIT_LIST_WIDTH_DEFAULT
        });
        showsPostScriptName = savedSettings.showPostScriptName;
        replacesStyleFonts = savedSettings.replaceStyleFonts;
        fitsListWidth = savedSettings.fitListWidth;

        /* 知らない値は初期状態に戻す / Unknown values fall back to the defaults */
        var sortMode = savedSettings.sortMode;
        currentSortMode = SORT_MODE_DEFAULT;
        for (var i = 0; i < SORT_MODES.length; i++) {
            if (SORT_MODES[i] === sortMode) currentSortMode = sortMode;
        }
    }

    /**
     * 今の画面の状態を保存する（対象は実行時の選択で決まり、［合成フォントのみ］［使っていないスタイルも表示］は毎回 OFF で開くので保存しない）
     * @returns {void}
     */
    function saveSettings() {
        settingsStore.save({
            showPostScriptName: showsPostScriptName,
            sortMode: currentSortMode,
            replaceStyleFonts: replacesStyleFonts,
            fitListWidth: fitsListWidth
        });
    }

    main();

})();
