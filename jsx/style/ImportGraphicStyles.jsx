#target illustrator
#targetengine "ImportGraphicStylesEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ImportGraphicStyles
Name=AIファイルからグラフィックスタイルを取り込む
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.8.1
Release-Date=2025-08-14
Description-BEGIN
指定したAIファイルからグラフィックスタイルを取り込み、スタイル名のラジオボタンを自動生成して選択オブジェクトへ適用します。
指定したファイルは記憶され、次回以降は自動で参照します。

Imports the graphic styles from an AI file you choose, builds a radio button per style name, and applies the selected one to the selection.
The chosen file is remembered and reused on later runs.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

指定したAIファイルからグラフィックスタイルを取り込み、スタイル名のラジオボタンを自動生成して選択オブジェクトへ適用します。
指定したファイルは記憶され、次回以降は自動で参照します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportGraphicStyles.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n0b929db4a4ad

### Overview

Imports the graphic styles from an AI file you choose, builds a radio button per style name, and applies the selected one to the selection.
The chosen file is remembered and reused on later runs.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportGraphicStyles.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportGraphicStyles";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "AIファイルからグラフィックスタイルを取り込む"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportGraphicStyles.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportGraphicStyles.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0b929db4a4ad"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportGraphicStyles.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var IMPORT_LAYER_NAME = "// _imported"; /* 取り込むときに一時的に貼り付けるレイヤー / temporary layer the styles are pasted onto */

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

    var SOURCE_FILE_TEXT_WIDTH = 240; /* 読み込み元のファイル名の幅 / width of the source file name */

    // =========================================
    // ダイアログボックスの位置 / Dialog position
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
    // ファイルビューア / File viewer
    // =========================================

    // ファイルビューアで表示（再利用パーツ） / Show in file viewer (reusable)

    /**
     * Path Finder（起動中のとき）か Finder で、フォルダーを開くかファイルを選択して表示する。
     * 補助アプリ /Applications/OpenInFileViewer.app に一時ファイルでパスを渡して起動する。
     * 補助アプリは illustrator-scripts の helpers/OpenInFileViewer.applescript から作る
     * @param {File|Folder} targetItem - 開くフォルダーか、選択して表示するファイル
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function openInFileViewer(targetItem) {
        /* 定数は巻き上げで未定義にならないよう関数内に置く / Kept local so hoisting never leaves them undefined */
        var viewerAppPath = "/Applications/OpenInFileViewer.app";
        var pathFilePath = "/tmp/open_in_file_viewer_path.txt";

        if ($.os.indexOf("Mac") === -1) return false;
        /* .app は実体がディレクトリなので Folder でも確かめる / An .app is a directory, so check it as a Folder too */
        if (!new Folder(viewerAppPath).exists && !new File(viewerAppPath).exists) return false;

        var pathFile = new File(pathFilePath);
        var written = false;
        try {
            pathFile.encoding = "UTF-8";
            pathFile.lineFeed = "Unix";
            if (pathFile.open("w")) {
                /* fsName で ~ ではなく絶対パスを渡す / fsName gives the absolute POSIX path */
                written = pathFile.write(targetItem.fsName);
            }
        } catch (e) {
        } finally {
            try { pathFile.close(); } catch (closeError) {}
        }
        return written && new File(viewerAppPath).execute();
    }

    // ファイルビューアで表示（再利用パーツ）ここまで / End of the reusable file viewer

    // =========================================
    // ボタン行 / Button row
    // =========================================

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px。字の高さはこの6倍）/ size of one dot (px); glyphs are six dots tall */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：ドットフォント Chava の「v」「.」「0〜9」を 6 行のドット（# が点）に写したもの（「v」は5ドット幅・4段に描き直したもの）
       / Glyphs: the dots of "v", "." and "0"-"9" from the Chava pixel font, six rows each ("#" is a dot) */
    var VERSION_FOOTER_GLYPHS = {
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", "..#.."],
        ".": [".", ".", ".", ".", ".", "#"],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", ".###."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", "###"],
        "2": [".###.", "#...#", "....#", ".###.", "#....", "#####"],
        "3": [".###.", "#...#", "..##.", "....#", "#...#", ".###."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#."],
        "5": ["#####", "#....", "####.", "....#", "#...#", ".###."],
        "6": [".###.", "#....", "####.", "#...#", "#...#", ".###."],
        "7": ["#####", "....#", "...#.", "..#..", ".#...", ".#..."],
        "8": [".###.", "#...#", ".###.", "#...#", "#...#", ".###."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", ".###."]
    };

    /**
     * ウィンドウの下端にバージョン（SCRIPT_VERSION）を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        var versionText = String(SCRIPT_VERSION);
        var versionCanvas = footerGroup.add("group");
        versionCanvas.preferredSize = [measureVersionFooterDots(versionText) * VERSION_FOOTER_PIXEL_SIZE, 6 * VERSION_FOOTER_PIXEL_SIZE];
        versionCanvas.helpTip = versionText;
        versionCanvas.onDraw = function () {
            drawVersionFooterDots(this.graphics, versionText);
        };
        targetWindow.versionFooter = versionCanvas;
        return versionCanvas;
    }

    /**
     * バージョンの文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - バージョンの文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * バージョンの文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - バージョンの文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // =========================================
    // 設定の保存 / Settings store
    // =========================================

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

    /* 前回のスタイルファイルのパスを記憶する（Folder.userData/illustrator-scripts/ImportGraphicStyles.json）。
       旧 swwwitch_ImportGraphicStyles.txt は新しい保存が無いときだけ読み継ぐ
       Remembers the last style file; the old swwwitch_ImportGraphicStyles.txt is read only until the first save */
    var styleFileStore = createSettingsStore(SCRIPT_NAME, "persistent", {
        legacy: function () {
            return readSettingsLegacyFile(Folder.userData + "/swwwitch_ImportGraphicStyles.txt");
        }
    });
    var DEFAULT_SETTINGS = { styleFilePath: "" }; /* 既定値 / defaults */

    /**
     * 記憶しているスタイルファイルのパスを読み込む
     * @returns {string} パス。無ければ空文字
     */
    function loadSavedFilePath() {
        return styleFileStore.load(DEFAULT_SETTINGS).styleFilePath;
    }

    /**
     * スタイルファイルのパスを記憶する
     * @param {string} filePath - スタイルファイルのパス
     * @returns {void}
     */
    function saveFilePath(filePath) {
        styleFileStore.save({ styleFilePath: filePath });
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 取り込み用の一時レイヤーを返す（無ければ作る）。ロックと非表示は外す
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @returns {Layer} 一時レイヤー
     */
    function getOrCreateImportLayer(destinationDoc) {
        var importLayer;
        /* getByName は見つからないと例外 / getByName throws when the layer is missing */
        try {
            importLayer = destinationDoc.layers.getByName(IMPORT_LAYER_NAME);
        } catch (e) {
            importLayer = destinationDoc.layers.add();
            importLayer.name = IMPORT_LAYER_NAME;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    /**
     * パスから表示用のファイル名を返す
     * @param {string} filePath - ファイルのパス
     * @returns {string} ファイル名。デコードできなければパスのまま
     */
    function getDisplayFileName(filePath) {
        /* File.name は URI エンコードされていて、壊れた % の並びだと decodeURI が URIError を出す
           File.name is URI-encoded, and decodeURI throws URIError on a malformed % sequence */
        try {
            return decodeURI(new File(filePath).name);
        } catch (e) {
            return filePath;
        }
    }

    /**
     * 名前でグラフィックスタイルを探す
     * @param {Document} targetDoc - 探すドキュメント
     * @param {string} styleName - グラフィックスタイル名
     * @returns {GraphicStyle|null} グラフィックスタイル。無ければ null
     */
    function findGraphicStyle(targetDoc, styleName) {
        /* getByName は見つからないと例外 / getByName throws when the style is missing */
        try {
            return targetDoc.graphicStyles.getByName(styleName);
        } catch (e) {
            return null;
        }
    }

    /**
     * 配列に同じ DOM の参照が入っているか
     * @param {Array} referenceList - 調べる配列
     * @param {Object} targetReference - 探す参照
     * @returns {boolean} 入っていれば true
     */
    function containsReference(referenceList, targetReference) {
        for (var i = 0; i < referenceList.length; i++) {
            if (referenceList[i] === targetReference) return true;
        }
        return false;
    }

    /**
     * レイヤーとサブレイヤーをすべて集める（親が先、子が後）
     * @param {Document|Layer} layerContainer - 集めるレイヤーの親
     * @param {Layer[]} layerList - 追加先の配列
     * @returns {Layer[]} layerList
     */
    function collectLayersDeep(layerContainer, layerList) {
        for (var i = 0; i < layerContainer.layers.length; i++) {
            layerList.push(layerContainer.layers[i]);
            collectLayersDeep(layerContainer.layers[i], layerList);
        }
        return layerList;
    }

    /**
     * 控えておいた選択に戻す（文字の選択は戻せないので、そのときは選択を解除する）
     * @param {Document} targetDoc - 対象のドキュメント
     * @param {Array|TextRange|null} savedSelection - 控えておいた doc.selection
     * @returns {void}
     */
    function restoreSelection(targetDoc, savedSelection) {
        targetDoc.selection = (savedSelection && savedSelection.typename !== "TextRange") ? savedSelection : null;
    }

    /**
     * 読み込み元のファイルを Finder（Path Finder が起動中ならそちら）で選択して表示する。
     * ファイルが無くなっていれば、残っている親フォルダーを開く
     * @param {string} filePath - 読み込み元のファイルのパス
     * @returns {void}
     */
    function revealStyleFile(filePath) {
        var styleFile = new File(filePath);
        if (styleFile.exists) {
            if (!openInFileViewer(styleFile)) styleFile.parent.execute();
            return;
        }
        var parentFolder = styleFile.parent;
        if (parentFolder && parentFolder.exists) {
            if (!openInFileViewer(parentFolder)) parentFolder.execute();
            return;
        }
        alert(getLabel("alert.fileNotFound", { fileName: getDisplayFileName(filePath) }));
    }

    /**
     * option キーを押しているか
     * @returns {boolean} 押していれば true
     */
    function isOptionKeyDown() {
        return ScriptUI.environment.keyboardState.altKey === true;
    }

    /**
     * スタイル用の AI ファイルを選ばせる
     * @returns {string} 選んだファイルのパス。キャンセルなら空文字
     */
    function pickStyleFile() {
        var pickedFile = File.openDialog(getLabel("dialog.pickFile"), function (candidate) {
            return (candidate instanceof Folder) || /\.ai$/i.test(candidate.name);
        });
        return pickedFile ? pickedFile.fsName : "";
    }

    // =========================================
    // スタイルの取り込み / Import styles
    // =========================================

    /**
     * AI ファイルのグラフィックスタイルを取り込む。
     * ファイルを開いてアートボード内をコピーし、取り込み先に貼り付けてすぐ消す（グラフィックスタイルだけが残る）
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @param {string} filePath - スタイルの AI ファイルのパス
     * @returns {string[]|null} 取り込めたグラフィックスタイル名。ファイルが無ければ警告して null
     */
    function importStylesFrom(destinationDoc, filePath) {
        var styleFile = new File(filePath);
        if (!styleFile.exists) {
            alert(getLabel("alert.fileNotFound", { fileName: getDisplayFileName(filePath) }));
            return null;
        }
        var sourceStyleNames = copyArtboardFromStyleFile(styleFile);
        app.activeDocument = destinationDoc;
        pasteAndDiscard(destinationDoc);
        return filterRegisteredStyleNames(destinationDoc, sourceStyleNames);
    }

    /**
     * スタイルの AI ファイルを開き、グラフィックスタイル名を集めて作業アートボード内をコピーする。
     * スクリプトが開いたときだけ保存せずに閉じる（すでに開いていたものは閉じない）
     * @param {File} styleFile - スタイルの AI ファイル
     * @returns {string[]} 元ファイルのグラフィックスタイル名（index 0 の既定のスタイルは除く）
     */
    function copyArtboardFromStyleFile(styleFile) {
        /* 開いているファイルを app.open しても手前に出るだけで、ドキュメントの数は変わらない
           Opening a file that is already open only brings it to the front, so the document count stays the same */
        var documentCountBefore = app.documents.length;
        var styleSourceDoc = app.open(styleFile);
        var wasAlreadyOpen = (app.documents.length === documentCountBefore);

        var sourceStyleNames = [];
        for (var i = 1; i < styleSourceDoc.graphicStyles.length; i++) {
            sourceStyleNames.push(styleSourceDoc.graphicStyles[i].name);
        }

        var selectionBeforeCopy = styleSourceDoc.selection;
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        if (wasAlreadyOpen) restoreSelection(styleSourceDoc, selectionBeforeCopy);
        else styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        return sourceStyleNames;
    }

    /**
     * クリップボードの中身を取り込み先に貼り付け、すぐに消す（グラフィックスタイルだけがドキュメントに残る）。
     * ［コピー元のレイヤーにペースト］がオンだと一時レイヤー以外にも入るので、一時レイヤーの外に貼り付いたものと、
     * 貼り付けで増えたレイヤー・サブレイヤーも消す。最後に貼り付ける前の選択に戻す
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @returns {void}
     */
    function pasteAndDiscard(destinationDoc) {
        var selectionBeforePaste = destinationDoc.selection;
        var layersBeforePaste = collectLayersDeep(destinationDoc, []);

        var importLayer = getOrCreateImportLayer(destinationDoc);
        destinationDoc.activeLayer = importLayer;
        /* 選択を残したまま貼り付けると、選択中のものや文字が置き換わることがある / Pasting over a selection can replace the selected items or text */
        destinationDoc.selection = null;
        app.executeMenuCommand("paste");

        /* 一時レイヤーのものはレイヤーごと消えるので、外に貼り付いたものだけ消す / Items on the temporary layer go with it; remove only the ones pasted elsewhere */
        var pastedItems = destinationDoc.selection;
        if (pastedItems && pastedItems.typename !== "TextRange") {
            for (var i = pastedItems.length - 1; i >= 0; i--) {
                if (pastedItems[i].layer !== importLayer) pastedItems[i].remove();
            }
        }

        /* 子から先に消す。最後の1枚のレイヤーは消せない / Remove children first; the last top-level layer cannot be removed */
        var layersAfterPaste = collectLayersDeep(destinationDoc, []);
        for (var j = layersAfterPaste.length - 1; j >= 0; j--) {
            var candidateLayer = layersAfterPaste[j];
            if (candidateLayer !== importLayer && containsReference(layersBeforePaste, candidateLayer)) continue;
            if (candidateLayer.parent === destinationDoc && destinationDoc.layers.length <= 1) continue;
            candidateLayer.locked = false;
            candidateLayer.remove();
        }

        /* ダイアログボックスが選択を避けて開けるように、キャンセルしても選択が残るように戻す
           Restore the selection so the dialog can avoid it and so it survives a cancel */
        restoreSelection(destinationDoc, selectionBeforePaste);
    }

    /**
     * 取り込み先に実際に登録されたグラフィックスタイル名だけを返す
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @param {string[]} styleNames - 調べるグラフィックスタイル名
     * @returns {string[]} 登録されていた名前
     */
    function filterRegisteredStyleNames(destinationDoc, styleNames) {
        var registeredNames = [];
        for (var i = 0; i < styleNames.length; i++) {
            if (findGraphicStyle(destinationDoc, styleNames[i])) registeredNames.push(styleNames[i]);
        }
        return registeredNames;
    }

    // =========================================
    // ダイアログボックス / Dialog
    // =========================================

    /**
     * グラフィックスタイルのラジオボタンを追加する
     * @param {Group|Panel} parentContainer - 追加先のコンテナ
     * @param {string} styleName - グラフィックスタイル名
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addStyleRadio(parentContainer, styleName) {
        var styleRadio = parentContainer.add("radiobutton", undefined, styleName);
        styleRadio.helpTip = getLabel("tooltip.style");
        return styleRadio;
    }

    /**
     * グラフィックスタイルを選ぶダイアログボックスを開く。
     * ［ファイルを選択...］で別のファイルを選ぶと、その場で取り込んでラジオボタンを組み直す
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @param {string} initialFilePath - 最初に表示するスタイルファイルのパス（無ければ空文字）
     * @param {string[]} initialStyleNames - 最初に並べるグラフィックスタイル名
     * @returns {string|null} 選んだグラフィックスタイル名。キャンセルなら null
     */
    function runStyleDialog(destinationDoc, initialFilePath, initialStyleNames) {
        var dialogState = {
            filePath: initialFilePath,
            styleNames: initialStyleNames
        };

        var styleDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(styleDialog);

        /* グラフィックスタイルのパネル（読み込むたびにラジオボタンを組み直す）/ Graphic style panel (radios rebuilt on each load) */
        var stylePanel = styleDialog.add("panel", undefined, getLabel("panel.style"));
        setupPanel(stylePanel, 6);
        var styleRadios = [];

        /* 読み込み元のファイルのパネル（ボタンの下にファイル名を表示）/ Source file panel (file name shown below the button) */
        var sourceFilePanel = styleDialog.add("panel", undefined, getLabel("panel.sourceFile"));
        setupPanel(sourceFilePanel, 6);
        var chooseFileButton = sourceFilePanel.add("button", undefined, getLabel("button.chooseFile"));
        chooseFileButton.alignment = "left";
        chooseFileButton.helpTip = getLabel("tooltip.chooseFile");
        var sourceFileNameText = sourceFilePanel.add("statictext", undefined, "", { truncate: "middle" });
        sourceFileNameText.preferredSize.width = SOURCE_FILE_TEXT_WIDTH;

        /* ボタン行（Mac の並び：キャンセル → 適用）/ Buttons (Mac order: Cancel → Apply) */
        var buttonRow = addButtonRow(styleDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var applyButton = buttonRow.rightGroup.add("button", undefined, getLabel("button.apply"), { name: "ok" });

        /**
         * 読み込み元のファイル名の表示を更新する
         * @returns {void}
         */
        function refreshSourceFileText() {
            sourceFileNameText.text = dialogState.filePath ? getDisplayFileName(dialogState.filePath) : getLabel("status.noFileSelected");
            /* 長い名前は中ほどが省かれるので、ツールチップでフルパスを見せる / Long names are truncated in the middle, so the tooltip shows the full path */
            sourceFileNameText.helpTip = dialogState.filePath ? dialogState.filePath + "\n" + getLabel("tooltip.revealFile") : "";
        }

        /**
         * 今のグラフィックスタイル名でラジオボタンを組み直す
         * @returns {void}
         */
        function rebuildStyleRadios() {
            for (var i = stylePanel.children.length - 1; i >= 0; i--) stylePanel.remove(stylePanel.children[i]);
            styleRadios = [];
            for (var j = 0; j < dialogState.styleNames.length; j++) {
                styleRadios.push(addStyleRadio(stylePanel, dialogState.styleNames[j]));
            }
            if (styleRadios.length > 0) styleRadios[0].value = true;
            else stylePanel.add("statictext", undefined, getLabel("status.noStyles"));
            applyButton.enabled = (styleRadios.length > 0);
            styleDialog.layout.layout(true);
            styleDialog.layout.resize();
        }

        /* onClick で受ける（addEventListener は発火しない環境がある）/ Use onClick, not addEventListener */
        chooseFileButton.onClick = function () {
            /* option＋クリックは選び直さず、今のファイルを Finder で表示する（ダイアログボックスは開いたまま）
               Option-click reveals the current file in Finder instead of choosing another (the dialog stays open) */
            if (isOptionKeyDown() && dialogState.filePath) {
                revealStyleFile(dialogState.filePath);
                return;
            }
            var pickedPath = pickStyleFile();
            if (!pickedPath) return;
            var importedStyleNames = importStylesFrom(destinationDoc, pickedPath);
            if (importedStyleNames === null) return; /* 見つからないときは importStylesFrom が警告済み / already warned in importStylesFrom */
            dialogState.filePath = pickedPath;
            dialogState.styleNames = importedStyleNames;
            saveFilePath(pickedPath); /* 次回もこのファイルを使う / Reuse it on later runs */
            refreshSourceFileText();
            rebuildStyleRadios();
        };

        /* ファイル名の option＋クリックでも Finder で表示する（statictext には onClick が無い）
           Option-clicking the file name also reveals it (a statictext has no onClick) */
        sourceFileNameText.addEventListener("mousedown", function () {
            if (isOptionKeyDown() && dialogState.filePath) revealStyleFile(dialogState.filePath);
        });

        refreshSourceFileText();
        rebuildStyleRadios();

        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(styleDialog);
        prepareDialogWindow(styleDialog, SCRIPT_NAME);
        if (styleDialog.show() !== 1) return null;
        for (var k = 0; k < styleRadios.length; k++) {
            if (styleRadios[k].value) return dialogState.styleNames[k];
        }
        return null;
    }

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

    var LABELS = {
        /* ダイアログボックス / Dialog */
        dialog: {
            title: { ja: "グラフィックスタイルを適用", en: "Apply Graphic Style" },
            pickFile: { ja: "グラフィックスタイルを読み込む AI ファイルを選択", en: "Choose an AI file to load graphic styles from" }
        },
        /* パネル / Panels */
        panel: {
            style: { ja: "グラフィックスタイル", en: "Graphic Styles" },
            sourceFile: { ja: "読み込み元のファイル", en: "Source File" }
        },
        /* ボタン / Buttons */
        button: {
            chooseFile: { ja: "ファイルを選択...", en: "Choose File..." },
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" }
        },
        /* ツールチップ / Tooltips */
        tooltip: {
            chooseFile: {
                ja: "グラフィックスタイルを読み込む AI ファイルを選びます。選んだファイルは次回も使います。\noption＋クリックで、今のファイルを Finder で表示します。",
                en: "Chooses the AI file to load graphic styles from. The file is reused on later runs.\nOption-click reveals the current file in Finder."
            },
            revealFile: { ja: "option＋クリックで Finder で表示", en: "Option-click to reveal in Finder" },
            style: { ja: "選択中のオブジェクトに適用するグラフィックスタイルです。", en: "The graphic style applied to the selected objects." }
        },
        /* 状態の表示 / Status text */
        status: {
            noFileSelected: { ja: "（未選択）", en: "(None)" },
            noStyles: { ja: "［ファイルを選択...］でスタイルを読み込んでください。", en: "Click Choose File... to load styles." }
        },
        /* 警告 / Alerts */
        alert: {
            openDocFirst: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document first." },
            selectObjectFirst: {
                ja: "グラフィックスタイルを適用するオブジェクトを選択してから実行してください。",
                en: "Select the objects to apply the graphic style to, then run the script."
            },
            fileNotFound: {
                ja: "ファイルが見つかりません。\n{fileName}",
                en: "The file was not found.\n{fileName}"
            },
            styleNotFound: {
                ja: "グラフィックスタイル「{styleName}」が見つかりません。",
                en: "The graphic style \"{styleName}\" was not found."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 記憶しているファイルから取り込み、選んだグラフィックスタイルを選択中のオブジェクトに適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocFirst"));
            return;
        }
        var destinationDoc = app.activeDocument;

        /* 文字ツールで文字を選択しているときは TextRange が返り、オブジェクトは選ばれていない
           Selecting characters with the Type tool returns a TextRange, not objects */
        var currentSelection = destinationDoc.selection;
        if (!currentSelection || currentSelection.typename === "TextRange" || currentSelection.length === 0) {
            alert(getLabel("alert.selectObjectFirst"));
            return;
        }
        /* 取り込みの貼り付けで選択が変わるので控えておく / The paste during import changes the selection, so keep a copy */
        var selectedItems = [];
        for (var i = 0; i < currentSelection.length; i++) selectedItems.push(currentSelection[i]);

        /* 記憶しているファイルがあればダイアログボックスの前に取り込む。見つからなければ［ファイルを選択...］で選び直してもらう
           If a file is remembered, import it before the dialog; if it is missing, the user picks it again */
        var savedFilePath = loadSavedFilePath();
        var initialStyleNames = [];
        if (savedFilePath && new File(savedFilePath).exists) {
            initialStyleNames = importStylesFrom(destinationDoc, savedFilePath) || [];
        }

        var styleName = runStyleDialog(destinationDoc, savedFilePath, initialStyleNames);
        if (!styleName) return;

        var graphicStyle = findGraphicStyle(destinationDoc, styleName);
        if (!graphicStyle) {
            alert(getLabel("alert.styleNotFound", { styleName: styleName }));
            return;
        }

        for (var j = 0; j < selectedItems.length; j++) {
            /* 適用できないオブジェクトがあっても残りには適用する / Keep applying to the rest when one item rejects the style */
            try {
                graphicStyle.applyTo(selectedItems[j]);
            } catch (e) {}
        }

        destinationDoc.selection = selectedItems;
    }

    main();

})();
