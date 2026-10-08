#target illustrator
#targetengine "OpenSaveAsPDFAndCloseEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定したファイルを Illustrator で開き、ダイアログボックスで選んだ設定で PDF として保存して閉じます。
付属スクリプト「ドキュメントを PDF として保存.jsx」をもとにしています。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OpenSaveAsPDFAndClose.md

### Overview

Opens the specified files, saves each one as PDF with the settings chosen in a dialog, and closes it.
Based on the bundled sample script "Save as PDFs.jsx".

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OpenSaveAsPDFAndClose.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "OpenSaveAsPDFAndClose";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-08";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OpenSaveAsPDFAndClose.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OpenSaveAsPDFAndClose.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 最初に対象にするファイル（空ならダイアログボックスの［選択...］で選ぶ）/ Initial files (empty = choose in the dialog) */
    var TARGET_FILES = [
        // "~/Desktop/sample.ai",
        // "~/Desktop/sample2.eps"
    ];

    var OPENABLE_FILE_PATTERN = /\.(ai|eps|pdf|svg)$/i; /* ［選択...］で選べるファイル / Files selectable with "Choose..." */
    /* PDF プリセットの初期値（上から順に、あるものを使う。どれも無ければ一覧の先頭）/ Initial PDF preset: the first one found, else the first in the list */
    var DEFAULT_PDF_PRESETS = ["[Illustrator Default]", "[Illustrator 初期設定]"];

    /* ダイアログボックスの初期値（2回目からは前回の設定）/ Initial dialog values (later runs use the last ones) */
    var DEFAULT_SETTINGS = {
        pdfPreset: "",               /* 空なら DEFAULT_PDF_PRESETS から / Empty = DEFAULT_PDF_PRESETS */
        useBleed: false,             /* 裁ち落としを付ける / Add bleed */
        bleedMm: 3,                  /* 裁ち落とし（mm）/ Bleed (mm) */
        useTrimMarks: false,         /* トンボを付ける / Add trim marks */
        trimMarkType: "japanese",    /* "japanese" / "roman" */
        useCustomFolder: false,      /* 保存先を指定する / Use a custom folder */
        customFolderPath: "",        /* 指定した保存先 / Custom folder path */
        overwrite: false,            /* 同名の PDF を上書きする / Overwrite an existing PDF */
        noOverwriteRule: "number"    /* 上書きしないとき："number"（連番を付ける）/ "skip"（スキップ）/ When not overwriting */
    };

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

    var PDF_PRESET_WIDTH  = 220; /* PDF プリセットの幅 / PDF preset dropdown width */
    var FOLDER_PATH_WIDTH = 320; /* 保存先のパス表示の幅 / Width of the destination path */
    var LABEL_WIDTH       = 110; /* 行ラベルの幅 / Row label width */
    var BLEED_FIELD_CHARS = 5;   /* 裁ち落としの欄の文字数 / Bleed field characters */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title:        { ja: "開いて PDF として保存", en: "Open and Save as PDF" },
            openFiles:    { ja: "PDF に書き出すファイルを選択してください", en: "Choose the files to save as PDF" },
            chooseFolder: { ja: "保存先のフォルダーを選択してください", en: "Choose the destination folder" }
        },
        panel: {
            files:       { ja: "ファイル", en: "Files" },
            pdf:         { ja: "PDF", en: "PDF" },
            destination: { ja: "保存先", en: "Destination" }
        },
        fieldLabel: {
            fileCount:       { ja: "対象", en: "Files" },
            pdfPreset:       { ja: "PDF プリセット", en: "PDF Preset" },
            noOverwriteRule: { ja: "上書きしないとき", en: "If Not Overwriting" }
        },
        value: {
            fileCount: { ja: "%1 件", en: "%1" },
            noFolder:  { ja: "（未選択）", en: "(Not Chosen)" }
        },
        checkbox: {
            bleed:     { ja: "裁ち落とし", en: "Bleed" },
            trimMarks: { ja: "トンボ", en: "Trim Marks" },
            overwrite: { ja: "同名の PDF を上書き", en: "Overwrite Existing PDF" }
        },
        radio: {
            sameFolder:   { ja: "元のファイルと同じ場所", en: "Same as Source File" },
            customFolder: { ja: "指定", en: "Custom" }
        },
        choice: {
            japanese: { ja: "日本式", en: "Japanese" },
            roman:    { ja: "西洋式", en: "Roman" },
            number:   { ja: "連番を付ける", en: "Add a Number" },
            skip:     { ja: "スキップ", en: "Skip" }
        },
        unit: {
            mm: "mm"
        },
        button: {
            chooseFiles:  { ja: "選択...", en: "Choose..." },
            chooseFolder: { ja: "選択...", en: "Choose..." },
            cancel:       { ja: "キャンセル", en: "Cancel" },
            ok:           { ja: "保存", en: "Save" }
        },
        tooltip: {
            bleed:           { ja: "上下左右に同じ幅の裁ち落としを付けます（ドキュメントの裁ち落とし設定は使いません）", en: "Adds the same bleed on all four sides (the document's bleed setting is not used)" },
            noOverwriteRule: { ja: "連番：「名前-1.pdf」「名前-2.pdf」… の空いている名前で保存", en: "Add a Number: saves as the first free name, \"name-1.pdf\", \"name-2.pdf\"..." }
        },
        alert: {
            noFiles:       { ja: "対象のファイルがありません。", en: "No files to process." },
            noFolder:      { ja: "保存先のフォルダーを選択してください。", en: "Choose the destination folder." },
            invalidBleed:  { ja: "裁ち落としには 0 以上の数値を入力してください。", en: "Enter a number of 0 or more for the bleed." },
            done:          { ja: "%1 件を PDF として保存しました。", en: "Saved %1 file(s) as PDF." },
            skipped:       { ja: "スキップ (%1) :", en: "Skipped (%1):" },
            failed:        { ja: "失敗 (%1) :", en: "Failed (%1):" },
            missingFiles:  { ja: "見つからないファイルがあります（対象から外します）:\n%1", en: "Some files were not found and were left out:\n%1" },
            alreadyOpen:   { ja: "すでに開いています", en: "Already open" },
            pdfExists:     { ja: "同名の PDF があります", en: "PDF already exists" },
            folderNotMade: { ja: "保存先フォルダーを作成できません : %1", en: "Could not create the output folder: %1" }
        }
    };

    // =========================================
    // 本体 / Main
    // =========================================

    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");

    /**
     * ダイアログボックスで設定を選ばせ、ファイルを順に開いて PDF に保存し、閉じる
     * @returns {void}
     */
    function main() {
        var dialogResult = showSaveDialog(getInitialFiles(), settingsStore.load(DEFAULT_SETTINGS));
        if (!dialogResult) return;
        settingsStore.save(dialogResult.settings);

        var settings = dialogResult.settings;
        var files = dialogResult.files;
        var pdfOptions = getPdfOptions(settings);
        var savedInteraction = app.userInteractionLevel;
        var savedCount = 0;
        var skipped = [];
        var failed = [];
        var savedPdfPaths = {}; /* この実行で保存した PDF（上書きしない）/ PDFs saved in this run (never overwritten) */

        /* フォントやリンクの警告で止まらないようにする / Suppress alerts while opening */
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        try {
            for (var i = 0; i < files.length; i++) {
                var fileName = decodeURI(files[i].name);
                try {
                    var skipReason = saveFileAsPdf(files[i], pdfOptions, settings, savedPdfPaths);
                    if (skipReason) skipped.push(fileName + " : " + skipReason);
                    else savedCount++;
                } catch (e) {
                    failed.push(fileName + " : " + e.message);
                }
            }
        } finally {
            app.userInteractionLevel = savedInteraction;
        }

        var message = getLabel("alert.done", [savedCount]);
        if (skipped.length > 0) message += "\n\n" + getLabel("alert.skipped", [skipped.length]) + "\n" + skipped.join("\n");
        if (failed.length > 0) message += "\n\n" + getLabel("alert.failed", [failed.length]) + "\n" + failed.join("\n");
        alert(message);
    }

    /**
     * TARGET_FILES から存在するファイルを返す（見つからないものは知らせて外す）
     * @returns {File[]} 存在するファイル
     */
    function getInitialFiles() {
        var files = [];
        var missing = [];
        for (var i = 0; i < TARGET_FILES.length; i++) {
            var targetFile = new File(TARGET_FILES[i]);
            if (targetFile.exists) files.push(targetFile);
            else missing.push(TARGET_FILES[i]);
        }
        if (missing.length > 0) alert(getLabel("alert.missingFiles", [missing.join("\n")]));
        return files;
    }

    /**
     * Illustrator にある PDF プリセットの名前を返す
     * @returns {string[]} プリセット名（取れなければ空）
     */
    function getPdfPresetNames() {
        var presetNames = [];
        try {
            var presetList = app.PDFPresetsList;
            for (var i = 0; i < presetList.length; i++) presetNames.push(String(presetList[i]));
        } catch (e) {}
        return presetNames;
    }

    /**
     * 使う PDF プリセットを決める（希望の名前 → DEFAULT_PDF_PRESETS → 一覧の先頭の順）
     * @param {string[]} presetNames - getPdfPresetNames() の結果
     * @param {string} [preferredName] - 希望のプリセット名
     * @returns {string} プリセット名（プリセットが1つも無ければ空文字）
     */
    function resolvePdfPreset(presetNames, preferredName) {
        var candidates = [preferredName].concat(DEFAULT_PDF_PRESETS);
        for (var i = 0; i < candidates.length; i++) {
            for (var j = 0; j < presetNames.length; j++) {
                if (candidates[i] && presetNames[j] === candidates[i]) return presetNames[j];
            }
        }
        return presetNames.length > 0 ? presetNames[0] : "";
    }

    /**
     * ホームフォルダーを「~」に縮めたパスを返す
     * @param {string} fullPath - フルパス
     * @returns {string} 表示用のパス
     */
    function abbreviateHomePath(fullPath) {
        var homePath = Folder("~").fsName;
        return (fullPath.indexOf(homePath) === 0) ? "~" + fullPath.substring(homePath.length) : fullPath;
    }

    // =========================================
    // ダイアログボックス / Dialog
    // =========================================

    /**
     * 設定のダイアログボックスを表示する
     * @param {File[]} initialFiles - 最初に対象にするファイル
     * @param {Object} initialSettings - 設定の初期値（DEFAULT_SETTINGS と同じ形）
     * @returns {{files: File[], settings: Object}|null} 選んだファイルと設定（キャンセル時は null）
     */
    function showSaveDialog(initialFiles, initialSettings) {
        var targetFiles = initialFiles;
        var customFolderPath = initialSettings.customFolderPath;
        var presetNames = getPdfPresetNames();
        var trimMarkTypeKeys = ["japanese", "roman"];
        var noOverwriteRuleKeys = ["number", "skip"];

        var saveDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(saveDialog);

        /* ファイル / Files */
        var filesPanel = saveDialog.add("panel", undefined, getLabel("panel.files"));
        setupPanel(filesPanel);
        var fileCountRow = filesPanel.add("group");
        setupRow(fileCountRow, ["fill", "center"]);
        var fileCountText = fileCountRow.add("statictext", undefined, "");
        fileCountText.alignment = ["fill", "center"];
        var btnChooseFiles = fileCountRow.add("button", undefined, getLabel("button.chooseFiles"));

        /* PDF / PDF */
        var pdfPanel = saveDialog.add("panel", undefined, getLabel("panel.pdf"));
        setupPanel(pdfPanel, 6);
        var presetRow = pdfPanel.add("group");
        setupRow(presetRow);
        var presetLabel = presetRow.add("statictext", undefined, labelText("fieldLabel.pdfPreset"));
        presetLabel.preferredSize.width = LABEL_WIDTH;
        presetLabel.justify = "right";
        var presetDropdown = presetRow.add("dropdownlist", undefined, presetNames);
        presetDropdown.preferredSize.width = PDF_PRESET_WIDTH;

        var bleedRow = pdfPanel.add("group");
        setupRow(bleedRow);
        var chkBleed = bleedRow.add("checkbox", undefined, getLabel("checkbox.bleed"));
        chkBleed.preferredSize.width = LABEL_WIDTH;
        chkBleed.helpTip = getLabel("tooltip.bleed");
        var bleedField = bleedRow.add("edittext", undefined, String(initialSettings.bleedMm));
        bleedField.characters = BLEED_FIELD_CHARS;
        bleedRow.add("statictext", undefined, getLabel("unit.mm"));

        var trimMarksRow = pdfPanel.add("group");
        setupRow(trimMarksRow);
        var chkTrimMarks = trimMarksRow.add("checkbox", undefined, getLabel("checkbox.trimMarks"));
        chkTrimMarks.preferredSize.width = LABEL_WIDTH;
        var trimMarkTypeDropdown = trimMarksRow.add("dropdownlist", undefined, [getLabel("choice.japanese"), getLabel("choice.roman")]);

        /* 保存先 / Destination */
        var destinationPanel = saveDialog.add("panel", undefined, getLabel("panel.destination"));
        setupPanel(destinationPanel, 6);
        var destinationRadioRow = destinationPanel.add("group");
        setupRow(destinationRadioRow);
        var rbSameFolder = destinationRadioRow.add("radiobutton", undefined, getLabel("radio.sameFolder"));
        var rbCustomFolder = destinationRadioRow.add("radiobutton", undefined, getLabel("radio.customFolder"));
        var destinationPathRow = destinationPanel.add("group");
        setupRow(destinationPathRow);
        var destinationPathText = destinationPathRow.add("statictext", undefined, "", { truncate: "middle" });
        destinationPathText.preferredSize.width = FOLDER_PATH_WIDTH;
        var btnChooseFolder = destinationPathRow.add("button", undefined, getLabel("button.chooseFolder"));

        var chkOverwrite = destinationPanel.add("checkbox", undefined, getLabel("checkbox.overwrite"));
        var noOverwriteRow = destinationPanel.add("group");
        setupRow(noOverwriteRow);
        var noOverwriteLabel = noOverwriteRow.add("statictext", undefined, labelText("fieldLabel.noOverwriteRule"));
        noOverwriteLabel.helpTip = getLabel("tooltip.noOverwriteRule");
        var noOverwriteDropdown = noOverwriteRow.add("dropdownlist", undefined, [getLabel("choice.number"), getLabel("choice.skip")]);
        noOverwriteDropdown.helpTip = getLabel("tooltip.noOverwriteRule");

        var buttonRow = addButtonRow(saveDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /* 初期値 / Initial values */
        var presetName = resolvePdfPreset(presetNames, initialSettings.pdfPreset);
        for (var i = 0; i < presetNames.length; i++) {
            if (presetNames[i] === presetName) presetDropdown.selection = i;
        }
        chkBleed.value = initialSettings.useBleed;
        chkTrimMarks.value = initialSettings.useTrimMarks;
        trimMarkTypeDropdown.selection = Math.max(0, indexOfKey(trimMarkTypeKeys, initialSettings.trimMarkType));
        rbSameFolder.value = !initialSettings.useCustomFolder;
        rbCustomFolder.value = initialSettings.useCustomFolder;
        chkOverwrite.value = initialSettings.overwrite;
        noOverwriteDropdown.selection = Math.max(0, indexOfKey(noOverwriteRuleKeys, initialSettings.noOverwriteRule));

        /**
         * 対象の件数を表示し、ファイル名の一覧をツールチップにする
         * @returns {void}
         */
        function updateFileCount() {
            var fileNames = [];
            for (var k = 0; k < targetFiles.length; k++) fileNames.push(decodeURI(targetFiles[k].name));
            fileCountText.text = labelValueText("fieldLabel.fileCount", getLabel("value.fileCount", [targetFiles.length]));
            fileCountText.helpTip = fileNames.join("\n");
        }

        /**
         * チェックに合わせて欄を有効・無効にする
         * @returns {void}
         */
        function updateEnabled() {
            bleedField.enabled = chkBleed.value;
            trimMarkTypeDropdown.enabled = chkTrimMarks.value;
            noOverwriteLabel.enabled = !chkOverwrite.value;
            noOverwriteDropdown.enabled = !chkOverwrite.value;
        }

        /**
         * 保存先のパスを表示する（元のファイルと同じ場所なら1件目のフォルダー）
         * @returns {void}
         */
        function updateDestinationPath() {
            var shownPath;
            if (rbCustomFolder.value) shownPath = customFolderPath ? new Folder(customFolderPath).fsName : "";
            else shownPath = (targetFiles.length > 0) ? targetFiles[0].parent.fsName : "";
            destinationPathText.text = shownPath ? abbreviateHomePath(shownPath) : getLabel("value.noFolder");
            destinationPathText.helpTip = shownPath;
            btnChooseFolder.enabled = rbCustomFolder.value;
        }

        /**
         * 保存先のフォルダーを選ばせる
         * @returns {boolean} 選んだら true
         */
        function chooseCustomFolder() {
            var startFolder = customFolderPath ? new Folder(customFolderPath) : (targetFiles.length > 0 ? targetFiles[0].parent : Folder("~"));
            var chosenFolder = startFolder.selectDlg(getLabel("dialog.chooseFolder"));
            if (!chosenFolder) return false;
            customFolderPath = chosenFolder.fsName;
            return true;
        }

        btnChooseFiles.onClick = function () {
            var pickedFiles = File.openDialog(getLabel("dialog.openFiles"), function (f) {
                return (f instanceof Folder) || OPENABLE_FILE_PATTERN.test(f.name);
            }, true);
            if (!pickedFiles) return;
            targetFiles = (pickedFiles instanceof Array) ? pickedFiles : [pickedFiles];
            updateFileCount();
            updateDestinationPath();
        };
        chkBleed.onClick = updateEnabled;
        chkTrimMarks.onClick = updateEnabled;
        chkOverwrite.onClick = updateEnabled;
        rbSameFolder.onClick = updateDestinationPath;
        rbCustomFolder.onClick = function () {
            /* まだフォルダーが無ければ選ばせ、選ばなければ元に戻す / Ask for a folder if none; revert when cancelled */
            if (!customFolderPath && !chooseCustomFolder()) {
                rbCustomFolder.value = false;
                rbSameFolder.value = true;
            }
            updateDestinationPath();
        };
        btnChooseFolder.onClick = function () {
            if (chooseCustomFolder()) updateDestinationPath();
        };

        /* 入力を確かめてから閉じる / Check the input before closing */
        btnOK.onClick = function () {
            if (targetFiles.length === 0) {
                alert(getLabel("alert.noFiles"));
                return;
            }
            var bleedMm = Number(bleedField.text);
            if (chkBleed.value && (bleedField.text === "" || isNaN(bleedMm) || bleedMm < 0)) {
                alert(getLabel("alert.invalidBleed"));
                return;
            }
            if (rbCustomFolder.value && !customFolderPath) {
                alert(getLabel("alert.noFolder"));
                return;
            }
            saveDialog.close(1);
        };

        updateFileCount();
        updateEnabled();
        updateDestinationPath();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(saveDialog, SCRIPT_NAME);
        if (saveDialog.show() !== 1) return null;

        var chosenBleed = Number(bleedField.text);
        return {
            files: targetFiles,
            settings: {
                pdfPreset: presetDropdown.selection ? presetDropdown.selection.text : "",
                useBleed: chkBleed.value,
                bleedMm: isNaN(chosenBleed) ? initialSettings.bleedMm : chosenBleed,
                useTrimMarks: chkTrimMarks.value,
                trimMarkType: trimMarkTypeKeys[trimMarkTypeDropdown.selection.index],
                useCustomFolder: rbCustomFolder.value,
                customFolderPath: customFolderPath,
                overwrite: chkOverwrite.value,
                noOverwriteRule: noOverwriteRuleKeys[noOverwriteDropdown.selection.index]
            }
        };
    }

    /**
     * キーの配列から位置を返す
     * @param {string[]} keys - キーの配列
     * @param {string} key - 探すキー
     * @returns {number} 位置（無ければ -1）
     */
    function indexOfKey(keys, key) {
        for (var i = 0; i < keys.length; i++) {
            if (keys[i] === key) return i;
        }
        return -1;
    }

    // =========================================
    // 保存 / Save
    // =========================================

    /**
     * PDF 保存オプションを作る
     * @param {Object} settings - ダイアログボックスの設定
     * @returns {PDFSaveOptions} 保存オプション
     */
    function getPdfOptions(settings) {
        var pdfOptions = new PDFSaveOptions();
        /* プリセットを先に読み込み、裁ち落とし・トンボはそのあとで上書きする / Load the preset first, then override bleed and marks */
        if (settings.pdfPreset) pdfOptions.pDFPreset = settings.pdfPreset;

        var bleedPt = settings.useBleed ? new UnitValue(settings.bleedMm, "mm").as("pt") : 0;
        pdfOptions.bleedLink = true;
        pdfOptions.bleedOffsetRect = [bleedPt, bleedPt, bleedPt, bleedPt];

        pdfOptions.trimMarks = settings.useTrimMarks;
        if (settings.useTrimMarks) {
            pdfOptions.pageMarksType = (settings.trimMarkType === "roman") ? PageMarksTypes.Roman : PageMarksTypes.Japanese;
        }
        pdfOptions.viewAfterSaving = false;
        return pdfOptions;
    }

    /**
     * 1ファイルを開いて PDF に保存し、閉じる
     * @param {File} sourceFile - 開くファイル
     * @param {PDFSaveOptions} pdfOptions - PDF 保存オプション
     * @param {Object} settings - ダイアログボックスの設定
     * @param {Object} savedPdfPaths - この実行で保存した PDF のパス（{ fsName: true }）。保存したら書き足す
     * @returns {string} スキップした理由（保存したときは空文字）
     */
    function saveFileAsPdf(sourceFile, pdfOptions, settings, savedPdfPaths) {
        var pdfFile = getPdfFile(sourceFile, settings, savedPdfPaths);
        if (!pdfFile) return getLabel("alert.pdfExists");

        /* すでに開いているドキュメントは対象外（saveAs で編集中のドキュメントが PDF に切り替わるため）/ Skip documents already open */
        var documentCountBefore = app.documents.length;
        var doc = app.open(sourceFile);
        if (app.documents.length === documentCountBefore) return getLabel("alert.alreadyOpen");

        try {
            doc.saveAs(pdfFile, pdfOptions);
        } finally {
            doc.close(SaveOptions.DONOTSAVECHANGES);
        }
        savedPdfPaths[pdfFile.fsName] = true;
        return "";
    }

    /**
     * 保存先の PDF ファイルを返す。同名があるときは設定に従って連番を付けるか null を返す。
     * 元ファイル自身と、この実行で保存した PDF は上書きの設定にかかわらず上書きしない
     * @param {File} sourceFile - 元ファイル
     * @param {Object} settings - ダイアログボックスの設定
     * @param {Object} savedPdfPaths - この実行で保存した PDF のパス（{ fsName: true }）
     * @returns {File|null} 保存先ファイル（スキップするときは null）
     */
    function getPdfFile(sourceFile, settings, savedPdfPaths) {
        var fileName = decodeURI(sourceFile.name);
        var dotIndex = fileName.lastIndexOf(".");
        var baseName = (dotIndex > 0) ? fileName.substring(0, dotIndex) : fileName;

        var outputFolder = settings.useCustomFolder ? new Folder(settings.customFolderPath) : sourceFile.parent;
        if (!outputFolder.exists && !outputFolder.create()) {
            throw new Error(getLabel("alert.folderNotMade", [outputFolder.fsName]));
        }

        var pdfFile = new File(outputFolder.fsName + "/" + baseName + ".pdf");
        if (!pdfFile.exists) return pdfFile;
        /* 元の PDF に重ねて保存すると、2ページ目以降が消える / Saving over the source PDF would drop its pages after the first */
        var isProtected = (pdfFile.fsName === sourceFile.fsName) || savedPdfPaths[pdfFile.fsName];
        if (settings.overwrite && !isProtected) return pdfFile;
        if (settings.noOverwriteRule === "skip") return null;

        /* 「名前-1.pdf」から空いている番号を探す / Find the first free "name-N.pdf" */
        for (var n = 1; ; n++) {
            pdfFile = new File(outputFolder.fsName + "/" + baseName + "-" + n + ".pdf");
            if (!pdfFile.exists) return pdfFile;
        }
    }

    main();

})();
