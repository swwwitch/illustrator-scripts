#target illustrator
#targetengine "export-EventEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログで選んだアートボードを、アートボード名ごとのルールでPNG書き出しします。
背景の透明・白、倍率、書き出し対象外の判定は `buildExportJobs()` で定義します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export-Event.md

### Overview

Exports the artboards chosen in a dialog to PNG, using rules keyed on the artboard name.
Transparent or white background, scale, and exclusions are all defined in `buildExportJobs()`.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export-Event.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "export-Event";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/export-Event.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/export-Event.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /**
     * アートボード名から書き出しジョブの配列を作る（空配列は書き出し対象外）
     * @param {string} artboardName - アートボード名
     * @returns {Array<{scale: number, transparent: boolean, suffix: string}>} 書き出しジョブ
     */
    function buildExportJobs(artboardName) {
        /* シンボル一覧は除外 / Skip "シンボル一覧" */
        if (artboardName === "シンボル一覧") {
            return [];
        }
        /* Doorkeeper は 200% + 100% を白背景で / Doorkeeper: 200% and 100% with white background */
        if (artboardName === "Doorkeeper") {
            return [
                { scale: 200, transparent: false, suffix: "-200" },
                { scale: 100, transparent: false, suffix: "" }
            ];
        }
        /* title / title2 系は透明背景で 100% / title / title2 family: 100% with transparent background */
        if (/^(title|title2)(-|$)/.test(artboardName)) {
            return [{ scale: 100, transparent: true, suffix: "" }];
        }
        /* その他は白背景で 100% / Otherwise: 100% with white background */
        return [{ scale: 100, transparent: false, suffix: "" }];
    }

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PROGRESS_WIDTH      = 360;              /* 状況表示とバーの幅 / Width of the status text and bar */
    var PROGRESS_BAR_HEIGHT = 14;               /* バーの高さ / Bar height */

    var ARTBOARD_LIST_WIDTH   = 440;               /* アートボード一覧の幅 / Artboard list width */
    var ARTBOARD_LIST_HEIGHT  = 280;               /* アートボード一覧の高さ / Artboard list height */
    var ARTBOARD_LIST_COLUMNS = [36, 220, 100, 70]; /* 一覧の列幅 [番号,名前,倍率,背景] / Column widths [no., name, scale, background] */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title:         { ja: "書き出すアートボード", en: "Artboards to Export" },
            progressTitle: { ja: "PNG 書き出し中...", en: "Exporting PNG..." }
        },
        column: {
            number:     { ja: "#", en: "#" },
            name:       { ja: "アートボード名", en: "Artboard" },
            scale:      { ja: "倍率", en: "Scale" },
            background: { ja: "背景", en: "Background" }
        },
        option: {
            closeAfterExport: { ja: "書き出し後にドキュメントを閉じる", en: "Close the document after export" },
            revealFolder:     { ja: "書き出し後に保存先を開く", en: "Open the output folder after export" }
        },
        background: {
            transparent: { ja: "透明", en: "Transparent" },
            white:       { ja: "白", en: "White" }
        },
        status: {
            preparing:  { ja: "準備中...", en: "Preparing..." },
            cancelling: { ja: "キャンセル中...", en: "Cancelling..." },
            cancelled:  { ja: "キャンセルしました", en: "Cancelled" },
            done:       { ja: "完了", en: "Done" }
        },
        button: {
            selectAll:   { ja: "すべて選択", en: "Select All" },
            deselectAll: { ja: "選択解除", en: "Deselect All" },
            cancel:      { ja: "キャンセル", en: "Cancel" },
            ok:          { ja: "書き出し", en: "Export" }
        },
        alert: {
            saveFirst: { ja: "先にドキュメントを保存してください。", en: "Save the document first." },
            exportError: {
                ja: "アートボード「%1」の書き出し中にエラーが発生しました：\n",
                en: "An error occurred while exporting the artboard \"%1\":\n"
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで選んだアートボードを名前ごとのルールで PNG 書き出しする
     * @returns {void}
     */
    function exportArtboardsAsPng() {
        if (app.documents.length === 0) {
            return;
        }

        var activeDoc = app.activeDocument;
        /* 一度も保存していないドキュメントは保存先が定まらないため中止（fullName が例外になることがある）/ Abort when the document has never been saved (fullName may throw) */
        var outputFolder = null;
        try {
            outputFolder = activeDoc.fullName.parent;
        } catch (e) {}
        if (!outputFolder || !outputFolder.exists) {
            alert(getLabel("alert.saveFirst"));
            return;
        }

        /* 書き出し対象とジョブを一括算出（buildExportJobs の二度呼びを回避）/ Resolve targets and jobs in one pass */
        var exportPlan = buildExportPlan(activeDoc);
        if (exportPlan.totalJobs === 0) {
            return;
        }

        /* 書き出すアートボードをダイアログで選ぶ / Choose the artboards to export in a dialog */
        exportPlan = showArtboardDialog(exportPlan);
        if (!exportPlan) {
            return;
        }

        var wasSaved = activeDoc.saved;
        var baseFileName = activeDoc.name.replace(/\.ai$/i, "");
        var progress = createProgressWindow(exportPlan.totalJobs);
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var cancelled = false;
        var completedCount = 0;
        /* 途中で失敗しても警告表示の設定と進捗ウィンドウは戻す / Always restore alerts and close the progress window */
        try {
            for (var i = 0; i < exportPlan.artboardPlans.length && !cancelled; i++) {
                var artboardPlan = exportPlan.artboardPlans[i];
                activeDoc.artboards.setActiveArtboardIndex(artboardPlan.index);
                for (var j = 0; j < artboardPlan.exportJobs.length; j++) {
                    progress.update(completedCount, artboardPlan.name + artboardPlan.exportJobs[j].suffix);
                    /* キャンセルボタンが押されていれば中断 / Stop if the cancel button was pressed */
                    if (progress.isCancelled()) {
                        cancelled = true;
                        break;
                    }
                    exportArtboardAsPng(activeDoc, outputFolder, baseFileName, artboardPlan.name, artboardPlan.exportJobs[j]);
                    completedCount++;
                }
            }
            progress.update(completedCount, getLabel(cancelled ? "status.cancelled" : "status.done"));
        } finally {
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            progress.close();
        }

        /* 1枚でも書き出したら保存先を開く（キャンセル時も途中までの分を確認できるように）/ Open the output folder once anything was exported, even after cancelling */
        if (exportPlan.revealFolder && completedCount > 0) {
            if (!openInFileViewer(outputFolder)) outputFolder.execute();
        }

        /* 書き出し前に保存済みだったときだけ保存せずに閉じ、未保存の編集があれば保存を確認する / Close without saving only when it was saved before export; ask when there are unsaved edits */
        if (exportPlan.closeAfterExport && !cancelled) {
            activeDoc.close(wasSaved ? SaveOptions.DONOTSAVECHANGES : SaveOptions.PROMPTTOSAVECHANGES);
        }
    }

    // =========================================
    // アートボード選択ダイアログ / Artboard dialog
    // =========================================

    /**
     * 書き出しジョブの倍率を「200% / 100%」の形にまとめる
     * @param {Object[]} exportJobs - 書き出しジョブ
     * @returns {string} 倍率の表示
     */
    function formatJobScales(exportJobs) {
        var scaleTexts = [];
        for (var i = 0; i < exportJobs.length; i++) {
            scaleTexts.push(exportJobs[i].scale + "%");
        }
        return scaleTexts.join(" / ");
    }

    /**
     * 書き出しジョブの背景を「透明」「白」の形にまとめる（同じ背景は1回だけ）
     * @param {Object[]} exportJobs - 書き出しジョブ
     * @returns {string} 背景の表示
     */
    function formatJobBackgrounds(exportJobs) {
        var backgroundTexts = [];
        for (var i = 0; i < exportJobs.length; i++) {
            var backgroundText = getLabel(exportJobs[i].transparent ? "background.transparent" : "background.white");
            if (("|" + backgroundTexts.join("|") + "|").indexOf("|" + backgroundText + "|") < 0) {
                backgroundTexts.push(backgroundText);
            }
        }
        return backgroundTexts.join(" / ");
    }

    /**
     * 書き出し対象のアートボードを一覧で示し、選んだものだけの書き出し計画を返す
     * @param {{artboardPlans: Object[], totalJobs: number}} exportPlan - buildExportPlan() の結果
     * @returns {{artboardPlans: Object[], totalJobs: number, closeAfterExport: boolean, revealFolder: boolean}|null} 選んだアートボードの書き出し計画。キャンセル時は null
     */
    function showArtboardDialog(exportPlan) {
        var artboardDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(artboardDialog);

        /* Mac では1列目が長い文字列で広がるため、先頭は短い番号にする / The first column widens with long text on Mac, so it holds the short number */
        var artboardList = artboardDialog.add("listbox", undefined, [], {
            multiselect: true,
            numberOfColumns: 4,
            showHeaders: true,
            columnTitles: [getLabel("column.number"), getLabel("column.name"), getLabel("column.scale"), getLabel("column.background")],
            columnWidths: ARTBOARD_LIST_COLUMNS
        });
        artboardList.preferredSize = [ARTBOARD_LIST_WIDTH, ARTBOARD_LIST_HEIGHT];

        var artboardPlans = exportPlan.artboardPlans;
        var allIndexes = [];
        for (var i = 0; i < artboardPlans.length; i++) {
            var listItem = artboardList.add("item", String(artboardPlans[i].index + 1));
            listItem.subItems[0].text = artboardPlans[i].name;
            listItem.subItems[1].text = formatJobScales(artboardPlans[i].exportJobs);
            listItem.subItems[2].text = formatJobBackgrounds(artboardPlans[i].exportJobs);
            allIndexes.push(i);
        }

        var closeAfterExportCheckbox = artboardDialog.add("checkbox", undefined, getLabel("option.closeAfterExport"));
        closeAfterExportCheckbox.alignment = "left";
        closeAfterExportCheckbox.value = true;

        /* 保存先を開くのは macOS だけ / Opening the output folder is macOS only */
        var revealFolderCheckbox = null;
        if ($.os.indexOf("Mac") >= 0) {
            revealFolderCheckbox = artboardDialog.add("checkbox", undefined, getLabel("option.revealFolder"));
            revealFolderCheckbox.alignment = "left";
            revealFolderCheckbox.value = true;
        }

        var buttonRow = addButtonRow(artboardDialog);
        var btnSelectAll = buttonRow.leftGroup.add("button", undefined, getLabel("button.selectAll"));
        var btnDeselectAll = buttonRow.leftGroup.add("button", undefined, getLabel("button.deselectAll"));
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /* 1つも選んでいないときは書き出せない / Export needs at least one artboard */
        function updateOKButton() {
            btnOK.enabled = (artboardList.selection !== null && artboardList.selection.length > 0);
        }

        artboardList.onChange = updateOKButton;
        btnSelectAll.onClick = function () {
            artboardList.selection = allIndexes;
            updateOKButton();
        };
        btnDeselectAll.onClick = function () {
            artboardList.selection = null;
            updateOKButton();
        };

        /* 初期状態はすべて選択 / Everything starts selected */
        artboardList.selection = allIndexes;
        updateOKButton();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(artboardDialog, SCRIPT_NAME);
        if (artboardDialog.show() !== 1) {
            return null;
        }

        /* 一覧の並び（アートボード順）で選択を拾う / Collect the selection in list order (artboard order) */
        var selectedPlans = [];
        var totalJobs = 0;
        for (var j = 0; j < artboardList.items.length; j++) {
            if (!artboardList.items[j].selected) continue;
            selectedPlans.push(artboardPlans[j]);
            totalJobs += artboardPlans[j].exportJobs.length;
        }
        if (selectedPlans.length === 0) {
            return null;
        }
        return { artboardPlans: selectedPlans, totalJobs: totalJobs, closeAfterExport: closeAfterExportCheckbox.value,
            revealFolder: revealFolderCheckbox !== null && revealFolderCheckbox.value
        };
    }

    // =========================================
    // ルール判定 / Rule resolver
    // =========================================

    /**
     * 書き出し対象のアートボードとジョブ、ジョブの総数を一括で求める
     * @param {Document} doc - 対象ドキュメント
     * @returns {{artboardPlans: Array<{index: number, name: string, exportJobs: Object[]}>, totalJobs: number}} 書き出し計画
     */
    function buildExportPlan(doc) {
        var artboardPlans = [];
        var totalJobs = 0;
        var artboardCount = doc.artboards.length;
        for (var i = 0; i < artboardCount; i++) {
            var artboardName = doc.artboards[i].name;
            var exportJobs = buildExportJobs(artboardName);
            if (exportJobs.length === 0) {
                continue;
            }
            artboardPlans.push({ index: i, name: artboardName, exportJobs: exportJobs });
            totalJobs += exportJobs.length;
        }
        return { artboardPlans: artboardPlans, totalJobs: totalJobs };
    }

    // =========================================
    // 書き出しヘルパー / Export helper
    // =========================================

    /**
     * 1つのアートボードを指定倍率・背景で PNG 書き出しする（アクティブなアートボードが対象）
     * @param {Document} sourceDoc - 書き出すドキュメント
     * @param {Folder} outputFolder - 保存先フォルダー
     * @param {string} baseFileName - 拡張子を除いたドキュメント名
     * @param {string} artboardName - アートボード名
     * @param {{scale: number, transparent: boolean, suffix: string}} exportJob - 書き出しジョブ
     * @returns {void}
     */
    function exportArtboardAsPng(sourceDoc, outputFolder, baseFileName, artboardName, exportJob) {
        var exportOptions = new ExportOptionsPNG24();
        exportOptions.artBoardClipping = true;
        exportOptions.antiAliasing = true;
        exportOptions.transparency = exportJob.transparent;
        exportOptions.horizontalScale = exportJob.scale;
        exportOptions.verticalScale = exportJob.scale;

        var outputFileName = baseFileName + "-" + artboardName + exportJob.suffix + ".png";
        var outputFile = new File(outputFolder.fsName + "/" + outputFileName);

        /* 1枚の失敗で全体を止めない / One failed export does not stop the rest */
        try {
            sourceDoc.exportFile(outputFile, ExportType.PNG24, exportOptions);
        } catch (e) {
            alert(getLabel("alert.exportError").split("%1").join(artboardName) + e.message);
        }
    }

    // =========================================
    // 進捗表示 / Progress UI
    // =========================================

    /**
     * 状況表示・プログレスバー・キャンセルボタンを持つ進捗パレットを開く
     * @param {number} totalJobs - ジョブの総数
     * @returns {{isCancelled: Function, update: Function, close: Function}} 進捗パレットの操作
     */
    function createProgressWindow(totalJobs) {
        var progressWin = new Window("palette", getLabel("dialog.progressTitle") + " " + SCRIPT_VERSION, undefined, { closeButton: false });
        setupWindow(progressWin);

        var statusText = progressWin.add("statictext", undefined, getLabel("status.preparing"));
        statusText.preferredSize.width = PROGRESS_WIDTH;

        var progressBar = progressWin.add("progressbar", undefined, 0, totalJobs);
        progressBar.preferredSize = [PROGRESS_WIDTH, PROGRESS_BAR_HEIGHT];

        /* キャンセルボタン（押下でフラグを立て、ループ側が中断）/ Cancel button (sets a flag that the export loop checks) */
        var cancelled = false;
        var buttonRow = addButtonRow(progressWin);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnCancel.onClick = function () {
            cancelled = true;
            btnCancel.enabled = false;
            statusText.text = getLabel("status.cancelling");
            progressWin.update();
        };
        alignRightOnlyButtonRow(buttonRow);

        progressWin.show();

        return {
            isCancelled: function () {
                return cancelled;
            },
            update: function (value, statusLabel) {
                statusText.text = statusLabel + "  (" + value + " / " + totalJobs + ")";
                progressBar.value = value;
                progressWin.update();
            },
            close: function () {
                progressWin.close();
            }
        };
    }

    exportArtboardsAsPng();

})();
