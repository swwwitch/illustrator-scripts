#target illustrator
#targetengine "IconCatalogEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

リポジトリー内のスクリプトが onDraw で描いているアイコンを、スクリプトごとに一覧で表示します。
各アイコンの描画は元のスクリプトから抜き出したもので、ここだけで動きます（元のスクリプトは読み込みません）。

### Overview

Shows the icons that the scripts in this repository draw in onDraw, grouped by script.
Each icon's drawing is lifted from its original script and runs on its own (the originals are not loaded).

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "IconCatalog";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    /* 一覧の寸法 / Catalog metrics */
    var SLOT_COLUMNS     = 6;             /* 1行の枠の数 / slots per row */
    var SLOT_ROWS        = 5;             /* 枠の行数 / slot rows */
    var SLOT_COUNT       = SLOT_COLUMNS * SLOT_ROWS; /* 1ページの枠の数 / slots per page */
    var SLOT_ICON_AREA   = [92, 70];      /* アイコンを描く枠の大きさ（最大のアイコン 76×26・66×66 が収まる）/ icon area per slot */
    var SLOT_SPACING     = 6;             /* 枠どうしの間隔 / gap between slots */
    var SCRIPT_LIST_SIZE = [230, 440];    /* スクリプトのリストの大きさ / script list size */

    // =========================================
    // ダイアログ共通 / Dialog helpers
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

    /* アイコンの色（UI の明暗に合わせる）/ icon colors following the UI brightness */
    var CATALOG_UI_DARK = isDarkUI();
    var CATALOG_INK_COLOR    = CATALOG_UI_DARK ? [0.85, 0.85, 0.85, 1] : [0.25, 0.25, 0.25, 1]; /* 図形 / shapes */
    var CATALOG_GROUND_COLOR = CATALOG_UI_DARK ? [0.22, 0.22, 0.22, 1] : [1, 1, 1, 1];          /* アイコンの地 / icon ground */

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

    var LABELS = {
        dialog: {
            title: { ja: "onDraw アイコン一覧", en: "onDraw Icon Catalog" }
        },
        panel: {
            icons: { ja: "アイコン", en: "Icons" }
        },
        fieldLabel: {
            script: { ja: "スクリプト", en: "Script" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            scriptList: { ja: "アイコンを表示するスクリプトを選びます。「すべて」は全スクリプトのアイコンを順に表示します。", en: "Pick the script whose icons are shown. All shows every script's icons in turn." },
            iconSlot: { ja: "{script}：{name}（{width}×{height}）", en: "{script}: {name} ({width} x {height})" }
        },
        list: {
            allScripts: { ja: "すべて", en: "All" }
        },
        status: {
            summary: { ja: "{scripts} 本のスクリプト・{icons} 個のアイコン", en: "{icons} icons in {scripts} scripts" },
            page: { ja: "{page}/{pages}", en: "{page}/{pages}" }
        }
    };

    // =========================================
    // アイコンの一覧 / Icon catalog
    // =========================================
    // 各スクリプトの onDraw から抜き出した描画。draw(g, w, h, ink, ground) は自己完結していて、外の値を参照しない
    // Drawing code lifted from each script's onDraw. Each draw(g, w, h, ink, ground) is self-contained

    var ICON_CATALOG = [];

    /* ==== AdjustPairGap (jsx/alignment/AdjustPairGap.jsx) ==== */
    ICON_CATALOG.push({
        script: "AdjustPairGap",
        name: "行揃え：自動",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var buttonColors = { bg: ground, border: [0.62, 0.62, 0.62, 1], line: ink };
            function getJustifyLineWidths(iconType, longWidth, shortWidth) {
                if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(iconType, buttonWidth, lineWidth) {
                var margin = 5;
                if (iconType === "right") return buttonWidth - margin - lineWidth;
                if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var lineYs = [7, 11, 15, 19];
                var lineWidths = getJustifyLineWidths(iconType, 15, 10);
                for (var i = 0; i < lineYs.length; i++) {
                    var lineWidth = lineWidths[i];
                    var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
                    graphics.newPath();
                    graphics.moveTo(lineStartX, lineYs[i]);
                    graphics.lineTo(lineStartX + lineWidth, lineYs[i]);
                    graphics.strokePath(linePen);
                }
            }
            function drawAutoIcon(graphics, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var centerX = Math.round(buttonWidth / 2);
                var topY = 7;
                var bottomY = 19;
                var halfWidth = 5;
                var crossY = 15;
                var crossHalf = Math.round(halfWidth * (crossY - topY) / (bottomY - topY));
                graphics.newPath();
                graphics.moveTo(centerX - halfWidth, bottomY);
                graphics.lineTo(centerX, topY);
                graphics.lineTo(centerX + halfWidth, bottomY);
                graphics.strokePath(linePen);
                graphics.newPath();
                graphics.moveTo(centerX - crossHalf, crossY);
                graphics.lineTo(centerX + crossHalf, crossY);
                graphics.strokePath(linePen);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, buttonColors.bg));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, buttonColors.border, 1));
            drawAutoIcon(g, w, buttonColors.line);
        }
    });

    /* ==== AdjustPairGap (jsx/alignment/AdjustPairGap.jsx) ==== */
    ICON_CATALOG.push({
        script: "AdjustPairGap",
        name: "行揃え：左揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var buttonColors = { bg: ground, border: [0.62, 0.62, 0.62, 1], line: ink };
            function getJustifyLineWidths(iconType, longWidth, shortWidth) {
                if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(iconType, buttonWidth, lineWidth) {
                var margin = 5;
                if (iconType === "right") return buttonWidth - margin - lineWidth;
                if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var lineYs = [7, 11, 15, 19];
                var lineWidths = getJustifyLineWidths(iconType, 15, 10);
                for (var i = 0; i < lineYs.length; i++) {
                    var lineWidth = lineWidths[i];
                    var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
                    graphics.newPath();
                    graphics.moveTo(lineStartX, lineYs[i]);
                    graphics.lineTo(lineStartX + lineWidth, lineYs[i]);
                    graphics.strokePath(linePen);
                }
            }
            function drawAutoIcon(graphics, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var centerX = Math.round(buttonWidth / 2);
                var topY = 7;
                var bottomY = 19;
                var halfWidth = 5;
                var crossY = 15;
                var crossHalf = Math.round(halfWidth * (crossY - topY) / (bottomY - topY));
                graphics.newPath();
                graphics.moveTo(centerX - halfWidth, bottomY);
                graphics.lineTo(centerX, topY);
                graphics.lineTo(centerX + halfWidth, bottomY);
                graphics.strokePath(linePen);
                graphics.newPath();
                graphics.moveTo(centerX - crossHalf, crossY);
                graphics.lineTo(centerX + crossHalf, crossY);
                graphics.strokePath(linePen);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, buttonColors.bg));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, buttonColors.border, 1));
            drawJustifyIconLines(g, "left", w, buttonColors.line);
        }
    });

    /* ==== AdjustPairGap (jsx/alignment/AdjustPairGap.jsx) ==== */
    ICON_CATALOG.push({
        script: "AdjustPairGap",
        name: "行揃え：中央揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var buttonColors = { bg: ground, border: [0.62, 0.62, 0.62, 1], line: ink };
            function getJustifyLineWidths(iconType, longWidth, shortWidth) {
                if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(iconType, buttonWidth, lineWidth) {
                var margin = 5;
                if (iconType === "right") return buttonWidth - margin - lineWidth;
                if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var lineYs = [7, 11, 15, 19];
                var lineWidths = getJustifyLineWidths(iconType, 15, 10);
                for (var i = 0; i < lineYs.length; i++) {
                    var lineWidth = lineWidths[i];
                    var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
                    graphics.newPath();
                    graphics.moveTo(lineStartX, lineYs[i]);
                    graphics.lineTo(lineStartX + lineWidth, lineYs[i]);
                    graphics.strokePath(linePen);
                }
            }
            function drawAutoIcon(graphics, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var centerX = Math.round(buttonWidth / 2);
                var topY = 7;
                var bottomY = 19;
                var halfWidth = 5;
                var crossY = 15;
                var crossHalf = Math.round(halfWidth * (crossY - topY) / (bottomY - topY));
                graphics.newPath();
                graphics.moveTo(centerX - halfWidth, bottomY);
                graphics.lineTo(centerX, topY);
                graphics.lineTo(centerX + halfWidth, bottomY);
                graphics.strokePath(linePen);
                graphics.newPath();
                graphics.moveTo(centerX - crossHalf, crossY);
                graphics.lineTo(centerX + crossHalf, crossY);
                graphics.strokePath(linePen);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, buttonColors.bg));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, buttonColors.border, 1));
            drawJustifyIconLines(g, "center", w, buttonColors.line);
        }
    });

    /* ==== AdjustPairGap (jsx/alignment/AdjustPairGap.jsx) ==== */
    ICON_CATALOG.push({
        script: "AdjustPairGap",
        name: "行揃え：右揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var buttonColors = { bg: ground, border: [0.62, 0.62, 0.62, 1], line: ink };
            function getJustifyLineWidths(iconType, longWidth, shortWidth) {
                if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(iconType, buttonWidth, lineWidth) {
                var margin = 5;
                if (iconType === "right") return buttonWidth - margin - lineWidth;
                if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var lineYs = [7, 11, 15, 19];
                var lineWidths = getJustifyLineWidths(iconType, 15, 10);
                for (var i = 0; i < lineYs.length; i++) {
                    var lineWidth = lineWidths[i];
                    var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
                    graphics.newPath();
                    graphics.moveTo(lineStartX, lineYs[i]);
                    graphics.lineTo(lineStartX + lineWidth, lineYs[i]);
                    graphics.strokePath(linePen);
                }
            }
            function drawAutoIcon(graphics, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var centerX = Math.round(buttonWidth / 2);
                var topY = 7;
                var bottomY = 19;
                var halfWidth = 5;
                var crossY = 15;
                var crossHalf = Math.round(halfWidth * (crossY - topY) / (bottomY - topY));
                graphics.newPath();
                graphics.moveTo(centerX - halfWidth, bottomY);
                graphics.lineTo(centerX, topY);
                graphics.lineTo(centerX + halfWidth, bottomY);
                graphics.strokePath(linePen);
                graphics.newPath();
                graphics.moveTo(centerX - crossHalf, crossY);
                graphics.lineTo(centerX + crossHalf, crossY);
                graphics.strokePath(linePen);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, buttonColors.bg));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, buttonColors.border, 1));
            drawJustifyIconLines(g, "right", w, buttonColors.line);
        }
    });

    /* ==== AdjustPairGap (jsx/alignment/AdjustPairGap.jsx) ==== */
    ICON_CATALOG.push({
        script: "AdjustPairGap",
        name: "行揃え：均等配置（最終行左揃え）",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var buttonColors = { bg: ground, border: [0.62, 0.62, 0.62, 1], line: ink };
            function getJustifyLineWidths(iconType, longWidth, shortWidth) {
                if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(iconType, buttonWidth, lineWidth) {
                var margin = 5;
                if (iconType === "right") return buttonWidth - margin - lineWidth;
                if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var lineYs = [7, 11, 15, 19];
                var lineWidths = getJustifyLineWidths(iconType, 15, 10);
                for (var i = 0; i < lineYs.length; i++) {
                    var lineWidth = lineWidths[i];
                    var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
                    graphics.newPath();
                    graphics.moveTo(lineStartX, lineYs[i]);
                    graphics.lineTo(lineStartX + lineWidth, lineYs[i]);
                    graphics.strokePath(linePen);
                }
            }
            function drawAutoIcon(graphics, buttonWidth, lineColor) {
                var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
                var centerX = Math.round(buttonWidth / 2);
                var topY = 7;
                var bottomY = 19;
                var halfWidth = 5;
                var crossY = 15;
                var crossHalf = Math.round(halfWidth * (crossY - topY) / (bottomY - topY));
                graphics.newPath();
                graphics.moveTo(centerX - halfWidth, bottomY);
                graphics.lineTo(centerX, topY);
                graphics.lineTo(centerX + halfWidth, bottomY);
                graphics.strokePath(linePen);
                graphics.newPath();
                graphics.moveTo(centerX - crossHalf, crossY);
                graphics.lineTo(centerX + crossHalf, crossY);
                graphics.strokePath(linePen);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, buttonColors.bg));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, buttonColors.border, 1));
            drawJustifyIconLines(g, "full", w, buttonColors.line);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "水平・垂直方向中央に整列",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawCenterBothIcon(g, w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "水平方向中央に整列",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawCenterOneIcon(g, w, "horizontal", ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "垂直方向中央に整列",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawCenterOneIcon(g, w, "vertical", ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "左上へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "upLeft", w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "上へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "up", w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "右上へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "upRight", w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "左へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "left", w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "右へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "right", w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "左下へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "downLeft", w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "下へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "down", w, ink);
        }
    });

    /* ==== AiAlignToArtboardPalette (jsx/alignment/AiAlignToArtboardPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardPalette",
        name: "右下へ寄せる",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            var MOVE_ICON_RULES = {
                up: ["top"], down: ["bottom"], left: ["left"], right: ["right"],
                upLeft: ["top", "left"], upRight: ["top", "right"],
                downLeft: ["bottom", "left"], downRight: ["bottom", "right"]
            };
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getIconCenter(size) {
                return Math.round(size / 2) + 0.5;
            }
            function roundToOddLength(rawLength) {
                var rounded = Math.round(rawLength);
                if (rounded % 2 !== 0) { return rounded; }
                return (rounded > 1) ? rounded - 1 : 1;
            }
            function getBlockSize(size) {
                return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
            }
            function drawEdgeRule(graphics, size, side, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var isTopOrLeft = (side === "top" || side === "left");
                var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
                if (side === "top" || side === "bottom") {
                    fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
                } else {
                    fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
                }
                return rulePosition;
            }
            function drawCenterRule(graphics, size, ruleDirection, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var iconCenter = getIconCenter(size);
                if (ruleDirection === "vertical") {
                    fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
                } else {
                    fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
                }
            }
            function fillCenteredBlock(graphics, size, blockSize, color) {
                var iconCenter = getIconCenter(size);
                fillRect(graphics, iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2, blockSize[0], blockSize[1], color);
            }
            function drawCenterBothIcon(graphics, size, color) {
                drawCenterRule(graphics, size, "vertical", color);
                drawCenterRule(graphics, size, "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawCenterOneIcon(graphics, size, iconType, color) {
                drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
                fillCenteredBlock(graphics, size, getBlockSize(size), color);
            }
            function drawMoveIcon(graphics, directionKey, size, color) {
                var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
                ruleSides = MOVE_ICON_RULES[directionKey];
                if (!ruleSides) { return; }
                blockSize = getBlockSize(size);
                clearance = Math.round(size * ICON_RULE_CLEARANCE);
                iconCenter = getIconCenter(size);
                blockX = iconCenter - blockSize[0] / 2;
                blockY = iconCenter - blockSize[1] / 2;
                for (i = 0; i < ruleSides.length; i++) {
                    side = ruleSides[i];
                    rulePosition = drawEdgeRule(graphics, size, side, color);
                    if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                    if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                    if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                    if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
                }
                fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawMoveIcon(g, "downRight", w, ink);
        }
    });

    /* ==== AiAlignToArtboardOld (jsx/alignment/single-function/AiAlignToArtboardOld.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardOld",
        name: "左揃え",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_BAR_THICKNESS = 0.25;
            var ICON_BAR_GAP = 0.07;
            var ICON_BAR_LONG = 0.55;
            var ICON_BAR_SHORT = 0.35;
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getBarStackOrigin(size) {
                var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
                return Math.round((size - stackLength) / 2);
            }
            function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
                if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
                if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
                return Math.round(rulePosition - barLength / 2);
            }
            function getBarLengths(size, iconType) {
                var longBar = Math.round(size * ICON_BAR_LONG);
                var shortBar = Math.round(size * ICON_BAR_SHORT);
                return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
            }
            function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
                if (isVertical) {
                    fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
                } else {
                    fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
                }
            }
            function drawAlignIcon(graphics, size, alignMode, color, iconType) {
                var isVertical = (iconType === "vertical");
                var rulePosition = getRulePosition(size, alignMode);
                var barThickness = Math.round(size * ICON_BAR_THICKNESS);
                var barGap = Math.round(size * ICON_BAR_GAP);
                var barStackOrigin = getBarStackOrigin(size);
                var barLengths = getBarLengths(size, iconType);
                var clearance = Math.round(size * ICON_RULE_CLEARANCE);
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);
                for (var i = 0; i < barLengths.length; i++) {
                    fillOrientedRect(graphics, isVertical,
                        getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                        barStackOrigin + i * (barThickness + barGap),
                        barLengths[i], barThickness, color);
                }
            }
            function drawCenterBothIcon(graphics, size, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var center = Math.round(size / 2) + 0.5;
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
                var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
                var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
                fillRect(graphics,
                    Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                    blockWidth, blockHeight, color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawAlignIcon(g, w, "start", ink, "horizontal");
        }
    });

    /* ==== AiAlignToArtboardOld (jsx/alignment/single-function/AiAlignToArtboardOld.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardOld",
        name: "水平方向中央揃え",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_BAR_THICKNESS = 0.25;
            var ICON_BAR_GAP = 0.07;
            var ICON_BAR_LONG = 0.55;
            var ICON_BAR_SHORT = 0.35;
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getBarStackOrigin(size) {
                var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
                return Math.round((size - stackLength) / 2);
            }
            function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
                if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
                if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
                return Math.round(rulePosition - barLength / 2);
            }
            function getBarLengths(size, iconType) {
                var longBar = Math.round(size * ICON_BAR_LONG);
                var shortBar = Math.round(size * ICON_BAR_SHORT);
                return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
            }
            function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
                if (isVertical) {
                    fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
                } else {
                    fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
                }
            }
            function drawAlignIcon(graphics, size, alignMode, color, iconType) {
                var isVertical = (iconType === "vertical");
                var rulePosition = getRulePosition(size, alignMode);
                var barThickness = Math.round(size * ICON_BAR_THICKNESS);
                var barGap = Math.round(size * ICON_BAR_GAP);
                var barStackOrigin = getBarStackOrigin(size);
                var barLengths = getBarLengths(size, iconType);
                var clearance = Math.round(size * ICON_RULE_CLEARANCE);
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);
                for (var i = 0; i < barLengths.length; i++) {
                    fillOrientedRect(graphics, isVertical,
                        getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                        barStackOrigin + i * (barThickness + barGap),
                        barLengths[i], barThickness, color);
                }
            }
            function drawCenterBothIcon(graphics, size, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var center = Math.round(size / 2) + 0.5;
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
                var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
                var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
                fillRect(graphics,
                    Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                    blockWidth, blockHeight, color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawAlignIcon(g, w, "center", ink, "horizontal");
        }
    });

    /* ==== AiAlignToArtboardOld (jsx/alignment/single-function/AiAlignToArtboardOld.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardOld",
        name: "右揃え",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_BAR_THICKNESS = 0.25;
            var ICON_BAR_GAP = 0.07;
            var ICON_BAR_LONG = 0.55;
            var ICON_BAR_SHORT = 0.35;
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getBarStackOrigin(size) {
                var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
                return Math.round((size - stackLength) / 2);
            }
            function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
                if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
                if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
                return Math.round(rulePosition - barLength / 2);
            }
            function getBarLengths(size, iconType) {
                var longBar = Math.round(size * ICON_BAR_LONG);
                var shortBar = Math.round(size * ICON_BAR_SHORT);
                return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
            }
            function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
                if (isVertical) {
                    fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
                } else {
                    fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
                }
            }
            function drawAlignIcon(graphics, size, alignMode, color, iconType) {
                var isVertical = (iconType === "vertical");
                var rulePosition = getRulePosition(size, alignMode);
                var barThickness = Math.round(size * ICON_BAR_THICKNESS);
                var barGap = Math.round(size * ICON_BAR_GAP);
                var barStackOrigin = getBarStackOrigin(size);
                var barLengths = getBarLengths(size, iconType);
                var clearance = Math.round(size * ICON_RULE_CLEARANCE);
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);
                for (var i = 0; i < barLengths.length; i++) {
                    fillOrientedRect(graphics, isVertical,
                        getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                        barStackOrigin + i * (barThickness + barGap),
                        barLengths[i], barThickness, color);
                }
            }
            function drawCenterBothIcon(graphics, size, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var center = Math.round(size / 2) + 0.5;
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
                var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
                var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
                fillRect(graphics,
                    Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                    blockWidth, blockHeight, color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawAlignIcon(g, w, "end", ink, "horizontal");
        }
    });

    /* ==== AiAlignToArtboardOld (jsx/alignment/single-function/AiAlignToArtboardOld.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardOld",
        name: "水平・垂直方向中央揃え",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_BAR_THICKNESS = 0.25;
            var ICON_BAR_GAP = 0.07;
            var ICON_BAR_LONG = 0.55;
            var ICON_BAR_SHORT = 0.35;
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getBarStackOrigin(size) {
                var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
                return Math.round((size - stackLength) / 2);
            }
            function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
                if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
                if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
                return Math.round(rulePosition - barLength / 2);
            }
            function getBarLengths(size, iconType) {
                var longBar = Math.round(size * ICON_BAR_LONG);
                var shortBar = Math.round(size * ICON_BAR_SHORT);
                return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
            }
            function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
                if (isVertical) {
                    fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
                } else {
                    fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
                }
            }
            function drawAlignIcon(graphics, size, alignMode, color, iconType) {
                var isVertical = (iconType === "vertical");
                var rulePosition = getRulePosition(size, alignMode);
                var barThickness = Math.round(size * ICON_BAR_THICKNESS);
                var barGap = Math.round(size * ICON_BAR_GAP);
                var barStackOrigin = getBarStackOrigin(size);
                var barLengths = getBarLengths(size, iconType);
                var clearance = Math.round(size * ICON_RULE_CLEARANCE);
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);
                for (var i = 0; i < barLengths.length; i++) {
                    fillOrientedRect(graphics, isVertical,
                        getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                        barStackOrigin + i * (barThickness + barGap),
                        barLengths[i], barThickness, color);
                }
            }
            function drawCenterBothIcon(graphics, size, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var center = Math.round(size / 2) + 0.5;
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
                var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
                var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
                fillRect(graphics,
                    Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                    blockWidth, blockHeight, color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawCenterBothIcon(g, w, ink);
        }
    });

    /* ==== AiAlignToArtboardOld (jsx/alignment/single-function/AiAlignToArtboardOld.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardOld",
        name: "上揃え",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_BAR_THICKNESS = 0.25;
            var ICON_BAR_GAP = 0.07;
            var ICON_BAR_LONG = 0.55;
            var ICON_BAR_SHORT = 0.35;
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getBarStackOrigin(size) {
                var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
                return Math.round((size - stackLength) / 2);
            }
            function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
                if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
                if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
                return Math.round(rulePosition - barLength / 2);
            }
            function getBarLengths(size, iconType) {
                var longBar = Math.round(size * ICON_BAR_LONG);
                var shortBar = Math.round(size * ICON_BAR_SHORT);
                return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
            }
            function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
                if (isVertical) {
                    fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
                } else {
                    fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
                }
            }
            function drawAlignIcon(graphics, size, alignMode, color, iconType) {
                var isVertical = (iconType === "vertical");
                var rulePosition = getRulePosition(size, alignMode);
                var barThickness = Math.round(size * ICON_BAR_THICKNESS);
                var barGap = Math.round(size * ICON_BAR_GAP);
                var barStackOrigin = getBarStackOrigin(size);
                var barLengths = getBarLengths(size, iconType);
                var clearance = Math.round(size * ICON_RULE_CLEARANCE);
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);
                for (var i = 0; i < barLengths.length; i++) {
                    fillOrientedRect(graphics, isVertical,
                        getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                        barStackOrigin + i * (barThickness + barGap),
                        barLengths[i], barThickness, color);
                }
            }
            function drawCenterBothIcon(graphics, size, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var center = Math.round(size / 2) + 0.5;
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
                var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
                var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
                fillRect(graphics,
                    Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                    blockWidth, blockHeight, color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawAlignIcon(g, w, "start", ink, "vertical");
        }
    });

    /* ==== AiAlignToArtboardOld (jsx/alignment/single-function/AiAlignToArtboardOld.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardOld",
        name: "垂直方向中央揃え",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_BAR_THICKNESS = 0.25;
            var ICON_BAR_GAP = 0.07;
            var ICON_BAR_LONG = 0.55;
            var ICON_BAR_SHORT = 0.35;
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getBarStackOrigin(size) {
                var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
                return Math.round((size - stackLength) / 2);
            }
            function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
                if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
                if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
                return Math.round(rulePosition - barLength / 2);
            }
            function getBarLengths(size, iconType) {
                var longBar = Math.round(size * ICON_BAR_LONG);
                var shortBar = Math.round(size * ICON_BAR_SHORT);
                return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
            }
            function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
                if (isVertical) {
                    fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
                } else {
                    fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
                }
            }
            function drawAlignIcon(graphics, size, alignMode, color, iconType) {
                var isVertical = (iconType === "vertical");
                var rulePosition = getRulePosition(size, alignMode);
                var barThickness = Math.round(size * ICON_BAR_THICKNESS);
                var barGap = Math.round(size * ICON_BAR_GAP);
                var barStackOrigin = getBarStackOrigin(size);
                var barLengths = getBarLengths(size, iconType);
                var clearance = Math.round(size * ICON_RULE_CLEARANCE);
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);
                for (var i = 0; i < barLengths.length; i++) {
                    fillOrientedRect(graphics, isVertical,
                        getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                        barStackOrigin + i * (barThickness + barGap),
                        barLengths[i], barThickness, color);
                }
            }
            function drawCenterBothIcon(graphics, size, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var center = Math.round(size / 2) + 0.5;
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
                var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
                var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
                fillRect(graphics,
                    Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                    blockWidth, blockHeight, color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawAlignIcon(g, w, "center", ink, "vertical");
        }
    });

    /* ==== AiAlignToArtboardOld (jsx/alignment/single-function/AiAlignToArtboardOld.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAlignToArtboardOld",
        name: "下揃え",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_BAR_THICKNESS = 0.25;
            var ICON_BAR_GAP = 0.07;
            var ICON_BAR_LONG = 0.55;
            var ICON_BAR_SHORT = 0.35;
            var ICON_RULE_INSET = 0.12;
            var ICON_RULE_OFFSET = 0.17;
            var ICON_RULE_CLEARANCE = 0.05;
            var ICON_BLOCK_WIDTH = 0.38;
            var ICON_BLOCK_HEIGHT = 0.30;
            function fillRect(graphics, x, y, width, height, color) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + width, y);
                graphics.lineTo(x + width, y + height);
                graphics.lineTo(x, y + height);
                graphics.closePath();
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
            }
            function getRulePosition(size, alignMode) {
                if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
                if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
                return Math.round(size / 2) + 0.5;
            }
            function getBarStackOrigin(size) {
                var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
                return Math.round((size - stackLength) / 2);
            }
            function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
                if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
                if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
                return Math.round(rulePosition - barLength / 2);
            }
            function getBarLengths(size, iconType) {
                var longBar = Math.round(size * ICON_BAR_LONG);
                var shortBar = Math.round(size * ICON_BAR_SHORT);
                return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
            }
            function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
                if (isVertical) {
                    fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
                } else {
                    fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
                }
            }
            function drawAlignIcon(graphics, size, alignMode, color, iconType) {
                var isVertical = (iconType === "vertical");
                var rulePosition = getRulePosition(size, alignMode);
                var barThickness = Math.round(size * ICON_BAR_THICKNESS);
                var barGap = Math.round(size * ICON_BAR_GAP);
                var barStackOrigin = getBarStackOrigin(size);
                var barLengths = getBarLengths(size, iconType);
                var clearance = Math.round(size * ICON_RULE_CLEARANCE);
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);
                for (var i = 0; i < barLengths.length; i++) {
                    fillOrientedRect(graphics, isVertical,
                        getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                        barStackOrigin + i * (barThickness + barGap),
                        barLengths[i], barThickness, color);
                }
            }
            function drawCenterBothIcon(graphics, size, color) {
                var ruleInset = Math.round(size * ICON_RULE_INSET);
                var ruleLength = size - ruleInset * 2;
                var center = Math.round(size / 2) + 0.5;
                fillRect(graphics, center - 0.5, ruleInset, 1, ruleLength, color);
                fillRect(graphics, ruleInset, center - 0.5, ruleLength, 1, color);
                var blockWidth = Math.round(size * ICON_BLOCK_WIDTH);
                var blockHeight = Math.round(size * ICON_BLOCK_HEIGHT);
                fillRect(graphics,
                    Math.round(center - blockWidth / 2), Math.round(center - blockHeight / 2),
                    blockWidth, blockHeight, color);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawAlignIcon(g, w, "end", ink, "vertical");
        }
    });

    /* ==== ArtboardNavigatorPalette (jsx/artboard/ArtboardNavigatorPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "ArtboardNavigatorPalette",
        name: "最初のアートボード",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            function strokePolyline(graphics, glyphPen, points) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                for (var i = 1; i < points.length; i++) {
                    graphics.lineTo(points[i][0], points[i][1]);
                }
                graphics.strokePath(glyphPen);
            }
            function drawNavGlyph(graphics, iconType, glyphColor) {
                if (iconType === "list") {
                    var listPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.0);
                    var cellSize = 6;
                    var cellOrigins = [[5, 5], [15, 5], [5, 15], [15, 15]];
                    for (var i = 0; i < cellOrigins.length; i++) {
                        graphics.newPath();
                        graphics.rectPath(cellOrigins[i][0], cellOrigins[i][1], cellSize, cellSize);
                        graphics.strokePath(listPen);
                    }
                    return;
                }
                var glyphPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
                if (iconType === "first") {
                    strokePolyline(graphics, glyphPen, [[9, 7], [9, 19]]);
                    strokePolyline(graphics, glyphPen, [[18, 7], [13, 13], [18, 19]]);
                } else if (iconType === "last") {
                    strokePolyline(graphics, glyphPen, [[8, 7], [13, 13], [8, 19]]);
                    strokePolyline(graphics, glyphPen, [[17, 7], [17, 19]]);
                } else if (iconType === "prev") {
                    strokePolyline(graphics, glyphPen, [[16, 7], [10, 13], [16, 19]]);
                } else if (iconType === "next") {
                    strokePolyline(graphics, glyphPen, [[10, 7], [16, 13], [10, 19]]);
                }
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.62, 0.62, 0.62, 1], 1));
            drawNavGlyph(g, "first", ink);
        }
    });

    /* ==== ArtboardNavigatorPalette (jsx/artboard/ArtboardNavigatorPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "ArtboardNavigatorPalette",
        name: "前のアートボード",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            function strokePolyline(graphics, glyphPen, points) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                for (var i = 1; i < points.length; i++) {
                    graphics.lineTo(points[i][0], points[i][1]);
                }
                graphics.strokePath(glyphPen);
            }
            function drawNavGlyph(graphics, iconType, glyphColor) {
                if (iconType === "list") {
                    var listPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.0);
                    var cellSize = 6;
                    var cellOrigins = [[5, 5], [15, 5], [5, 15], [15, 15]];
                    for (var i = 0; i < cellOrigins.length; i++) {
                        graphics.newPath();
                        graphics.rectPath(cellOrigins[i][0], cellOrigins[i][1], cellSize, cellSize);
                        graphics.strokePath(listPen);
                    }
                    return;
                }
                var glyphPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
                if (iconType === "first") {
                    strokePolyline(graphics, glyphPen, [[9, 7], [9, 19]]);
                    strokePolyline(graphics, glyphPen, [[18, 7], [13, 13], [18, 19]]);
                } else if (iconType === "last") {
                    strokePolyline(graphics, glyphPen, [[8, 7], [13, 13], [8, 19]]);
                    strokePolyline(graphics, glyphPen, [[17, 7], [17, 19]]);
                } else if (iconType === "prev") {
                    strokePolyline(graphics, glyphPen, [[16, 7], [10, 13], [16, 19]]);
                } else if (iconType === "next") {
                    strokePolyline(graphics, glyphPen, [[10, 7], [16, 13], [10, 19]]);
                }
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.62, 0.62, 0.62, 1], 1));
            drawNavGlyph(g, "prev", ink);
        }
    });

    /* ==== ArtboardNavigatorPalette (jsx/artboard/ArtboardNavigatorPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "ArtboardNavigatorPalette",
        name: "全体表示",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            function strokePolyline(graphics, glyphPen, points) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                for (var i = 1; i < points.length; i++) {
                    graphics.lineTo(points[i][0], points[i][1]);
                }
                graphics.strokePath(glyphPen);
            }
            function drawNavGlyph(graphics, iconType, glyphColor) {
                if (iconType === "list") {
                    var listPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.0);
                    var cellSize = 6;
                    var cellOrigins = [[5, 5], [15, 5], [5, 15], [15, 15]];
                    for (var i = 0; i < cellOrigins.length; i++) {
                        graphics.newPath();
                        graphics.rectPath(cellOrigins[i][0], cellOrigins[i][1], cellSize, cellSize);
                        graphics.strokePath(listPen);
                    }
                    return;
                }
                var glyphPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
                if (iconType === "first") {
                    strokePolyline(graphics, glyphPen, [[9, 7], [9, 19]]);
                    strokePolyline(graphics, glyphPen, [[18, 7], [13, 13], [18, 19]]);
                } else if (iconType === "last") {
                    strokePolyline(graphics, glyphPen, [[8, 7], [13, 13], [8, 19]]);
                    strokePolyline(graphics, glyphPen, [[17, 7], [17, 19]]);
                } else if (iconType === "prev") {
                    strokePolyline(graphics, glyphPen, [[16, 7], [10, 13], [16, 19]]);
                } else if (iconType === "next") {
                    strokePolyline(graphics, glyphPen, [[10, 7], [16, 13], [10, 19]]);
                }
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.62, 0.62, 0.62, 1], 1));
            drawNavGlyph(g, "list", ink);
        }
    });

    /* ==== ArtboardNavigatorPalette (jsx/artboard/ArtboardNavigatorPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "ArtboardNavigatorPalette",
        name: "次のアートボード",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            function strokePolyline(graphics, glyphPen, points) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                for (var i = 1; i < points.length; i++) {
                    graphics.lineTo(points[i][0], points[i][1]);
                }
                graphics.strokePath(glyphPen);
            }
            function drawNavGlyph(graphics, iconType, glyphColor) {
                if (iconType === "list") {
                    var listPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.0);
                    var cellSize = 6;
                    var cellOrigins = [[5, 5], [15, 5], [5, 15], [15, 15]];
                    for (var i = 0; i < cellOrigins.length; i++) {
                        graphics.newPath();
                        graphics.rectPath(cellOrigins[i][0], cellOrigins[i][1], cellSize, cellSize);
                        graphics.strokePath(listPen);
                    }
                    return;
                }
                var glyphPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
                if (iconType === "first") {
                    strokePolyline(graphics, glyphPen, [[9, 7], [9, 19]]);
                    strokePolyline(graphics, glyphPen, [[18, 7], [13, 13], [18, 19]]);
                } else if (iconType === "last") {
                    strokePolyline(graphics, glyphPen, [[8, 7], [13, 13], [8, 19]]);
                    strokePolyline(graphics, glyphPen, [[17, 7], [17, 19]]);
                } else if (iconType === "prev") {
                    strokePolyline(graphics, glyphPen, [[16, 7], [10, 13], [16, 19]]);
                } else if (iconType === "next") {
                    strokePolyline(graphics, glyphPen, [[10, 7], [16, 13], [10, 19]]);
                }
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.62, 0.62, 0.62, 1], 1));
            drawNavGlyph(g, "next", ink);
        }
    });

    /* ==== ArtboardNavigatorPalette (jsx/artboard/ArtboardNavigatorPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "ArtboardNavigatorPalette",
        name: "最後のアートボード",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            function strokePolyline(graphics, glyphPen, points) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                for (var i = 1; i < points.length; i++) {
                    graphics.lineTo(points[i][0], points[i][1]);
                }
                graphics.strokePath(glyphPen);
            }
            function drawNavGlyph(graphics, iconType, glyphColor) {
                if (iconType === "list") {
                    var listPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.0);
                    var cellSize = 6;
                    var cellOrigins = [[5, 5], [15, 5], [5, 15], [15, 15]];
                    for (var i = 0; i < cellOrigins.length; i++) {
                        graphics.newPath();
                        graphics.rectPath(cellOrigins[i][0], cellOrigins[i][1], cellSize, cellSize);
                        graphics.strokePath(listPen);
                    }
                    return;
                }
                var glyphPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, 1.6);
                if (iconType === "first") {
                    strokePolyline(graphics, glyphPen, [[9, 7], [9, 19]]);
                    strokePolyline(graphics, glyphPen, [[18, 7], [13, 13], [18, 19]]);
                } else if (iconType === "last") {
                    strokePolyline(graphics, glyphPen, [[8, 7], [13, 13], [8, 19]]);
                    strokePolyline(graphics, glyphPen, [[17, 7], [17, 19]]);
                } else if (iconType === "prev") {
                    strokePolyline(graphics, glyphPen, [[16, 7], [10, 13], [16, 19]]);
                } else if (iconType === "next") {
                    strokePolyline(graphics, glyphPen, [[10, 7], [16, 13], [10, 19]]);
                }
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.62, 0.62, 0.62, 1], 1));
            drawNavGlyph(g, "last", ink);
        }
    });

    /* ==== AiFileFinder (jsx/files/AiFileFinder.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiFileFinder",
        name: "検索語のクリア（丸に×）",
        size: [20, 20],
        draw: function (g, w, h, ink, ground) {
            var CLEAR_CIRCLE_INSET = 2;
            var CLEAR_GLYPH_INSET = 6;
            var CLEAR_STROKE_WIDTH = 1.5;
            var buttonSize = w;
            g.newPath();
            g.rectPath(0, 0, buttonSize, buttonSize);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var glyphPen = g.newPen(g.PenType.SOLID_COLOR, ink, CLEAR_STROKE_WIDTH);
            var circleSize = buttonSize - CLEAR_CIRCLE_INSET * 2;
            g.newPath();
            g.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
            g.strokePath(glyphPen);
            var glyphStart = CLEAR_GLYPH_INSET;
            var glyphEnd = buttonSize - CLEAR_GLYPH_INSET;
            g.newPath();
            g.moveTo(glyphStart, glyphStart);
            g.lineTo(glyphEnd, glyphEnd);
            g.strokePath(glyphPen);
            g.newPath();
            g.moveTo(glyphEnd, glyphStart);
            g.lineTo(glyphStart, glyphEnd);
            g.strokePath(glyphPen);
        }
    });

    /* ==== FavoriteFontPickerPalette (jsx/fonts/FavoriteFontPickerPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "FavoriteFontPickerPalette",
        name: "検索語のクリア（丸に×）",
        size: [20, 20],
        draw: function (g, w, h, ink, ground) {
            var CLEAR_CIRCLE_INSET = 2;
            var CLEAR_GLYPH_INSET = 6;
            var CLEAR_STROKE_WIDTH = 1.5;
            var buttonSize = w;
            g.newPath();
            g.rectPath(0, 0, buttonSize, buttonSize);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var glyphPen = g.newPen(g.PenType.SOLID_COLOR, ink, CLEAR_STROKE_WIDTH);
            var circleSize = buttonSize - CLEAR_CIRCLE_INSET * 2;
            g.newPath();
            g.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
            g.strokePath(glyphPen);
            var glyphStart = CLEAR_GLYPH_INSET;
            var glyphEnd = buttonSize - CLEAR_GLYPH_INSET;
            g.newPath();
            g.moveTo(glyphStart, glyphStart);
            g.lineTo(glyphEnd, glyphEnd);
            g.strokePath(glyphPen);
            g.newPath();
            g.moveTo(glyphEnd, glyphStart);
            g.lineTo(glyphStart, glyphEnd);
            g.strokePath(glyphPen);
        }
    });

    /* ==== FavoriteFontPickerPalette (jsx/fonts/FavoriteFontPickerPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "FavoriteFontPickerPalette",
        name: "カスタムセット（外れている：しおり輪郭）",
        size: [33, 33],
        draw: function (g, w, h, ink, ground) {
            var isOn = false;
            var setNumber = 1;
            var SET_CARD_SIZE = 22;
            var SET_CARD_OFFSET = 5;
            var SET_CARD_STROKE_WIDTH = 2;
            var SET_BOOKMARK_WIDTH = 7;
            var SET_BOOKMARK_HEIGHT = 10;
            var SET_BOOKMARK_NOTCH = 3;
            var SET_BOOKMARK_INSET = 3;
            var SET_NUMBER_FONT_SIZE = 13;
            /* 外れているときは通常色（少し淡い）、入っているときはいちばん濃い色 / Muted when off, strongest when on */
            function mix(a, b, t) { return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t, 1]; }
            var glyphColor = isOn ? ink : mix(ink, ground, 0.4);
            function drawBookmark(left, top, isFilled) {
                var right = left + SET_BOOKMARK_WIDTH;
                var bottom = top + SET_BOOKMARK_HEIGHT;
                var notchTop = bottom - SET_BOOKMARK_NOTCH;
                if (!isFilled) {
                    var outlinePen = g.newPen(g.PenType.SOLID_COLOR, glyphColor, 1.5);
                    g.newPath();
                    g.moveTo(left, top);
                    g.lineTo(right, top);
                    g.lineTo(right, bottom);
                    g.lineTo(left + SET_BOOKMARK_WIDTH / 2, notchTop);
                    g.lineTo(left, bottom);
                    g.lineTo(left, top);
                    g.strokePath(outlinePen);
                    return;
                }
                var fillBrush = g.newBrush(g.BrushType.SOLID_COLOR, glyphColor);
                g.newPath();
                g.rectPath(left, top, SET_BOOKMARK_WIDTH, SET_BOOKMARK_HEIGHT - SET_BOOKMARK_NOTCH);
                g.fillPath(fillBrush);
                for (var y = notchTop; y < bottom; y++) {
                    var legWidth = SET_BOOKMARK_WIDTH / 2 - (y - notchTop + 0.5) * (SET_BOOKMARK_WIDTH / 2) / SET_BOOKMARK_NOTCH;
                    if (legWidth <= 0) break;
                    g.newPath();
                    g.rectPath(left, y, legWidth, 1);
                    g.fillPath(fillBrush);
                    g.newPath();
                    g.rectPath(right - legWidth, y, legWidth, 1);
                    g.fillPath(fillBrush);
                }
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var cardPen = g.newPen(g.PenType.SOLID_COLOR, glyphColor, SET_CARD_STROKE_WIDTH);
            var backLeft = Math.round((w - SET_CARD_SIZE - SET_CARD_OFFSET) / 2);
            var backTop = Math.round((h - SET_CARD_SIZE - SET_CARD_OFFSET) / 2);
            var frontLeft = backLeft + SET_CARD_OFFSET;
            var frontTop = backTop + SET_CARD_OFFSET;
            var backRight = backLeft + SET_CARD_SIZE;
            var backBottom = backTop + SET_CARD_SIZE;
            g.newPath();
            g.moveTo(frontLeft, backBottom);
            g.lineTo(backLeft, backBottom);
            g.lineTo(backLeft, backTop);
            g.lineTo(backRight, backTop);
            g.lineTo(backRight, frontTop);
            g.strokePath(cardPen);
            g.newPath();
            g.rectPath(frontLeft, frontTop, SET_CARD_SIZE, SET_CARD_SIZE);
            g.strokePath(cardPen);
            var bookmarkLeft = frontLeft + SET_CARD_SIZE - SET_BOOKMARK_INSET - SET_BOOKMARK_WIDTH;
            drawBookmark(bookmarkLeft, frontTop, isOn);
            var numberFont = g.font;
            try { numberFont = ScriptUI.newFont("dialog", "BOLD", SET_NUMBER_FONT_SIZE); } catch (e) {}
            var numberText = String(setNumber);
            var numberSize = null;
            try { numberSize = g.measureString(numberText, numberFont); } catch (e2) {}
            if (!numberSize || typeof numberSize[0] !== "number") numberSize = [8, 15];
            var numberPen = g.newPen(g.PenType.SOLID_COLOR, glyphColor, 1);
            var numberLeft = frontLeft + Math.round((bookmarkLeft - frontLeft - numberSize[0]) / 2);
            var numberTop = frontTop + SET_CARD_SIZE - numberSize[1] - 2;
            g.drawString(numberText, numberPen, numberLeft, numberTop, numberFont);
        }
    });

    /* ==== FavoriteFontPickerPalette (jsx/fonts/FavoriteFontPickerPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "FavoriteFontPickerPalette",
        name: "カスタムセット（入っている：しおり塗り）",
        size: [33, 33],
        draw: function (g, w, h, ink, ground) {
            var isOn = true;
            var setNumber = 1;
            var SET_CARD_SIZE = 22;
            var SET_CARD_OFFSET = 5;
            var SET_CARD_STROKE_WIDTH = 2;
            var SET_BOOKMARK_WIDTH = 7;
            var SET_BOOKMARK_HEIGHT = 10;
            var SET_BOOKMARK_NOTCH = 3;
            var SET_BOOKMARK_INSET = 3;
            var SET_NUMBER_FONT_SIZE = 13;
            /* 外れているときは通常色（少し淡い）、入っているときはいちばん濃い色 / Muted when off, strongest when on */
            function mix(a, b, t) { return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t, 1]; }
            var glyphColor = isOn ? ink : mix(ink, ground, 0.4);
            function drawBookmark(left, top, isFilled) {
                var right = left + SET_BOOKMARK_WIDTH;
                var bottom = top + SET_BOOKMARK_HEIGHT;
                var notchTop = bottom - SET_BOOKMARK_NOTCH;
                if (!isFilled) {
                    var outlinePen = g.newPen(g.PenType.SOLID_COLOR, glyphColor, 1.5);
                    g.newPath();
                    g.moveTo(left, top);
                    g.lineTo(right, top);
                    g.lineTo(right, bottom);
                    g.lineTo(left + SET_BOOKMARK_WIDTH / 2, notchTop);
                    g.lineTo(left, bottom);
                    g.lineTo(left, top);
                    g.strokePath(outlinePen);
                    return;
                }
                var fillBrush = g.newBrush(g.BrushType.SOLID_COLOR, glyphColor);
                g.newPath();
                g.rectPath(left, top, SET_BOOKMARK_WIDTH, SET_BOOKMARK_HEIGHT - SET_BOOKMARK_NOTCH);
                g.fillPath(fillBrush);
                for (var y = notchTop; y < bottom; y++) {
                    var legWidth = SET_BOOKMARK_WIDTH / 2 - (y - notchTop + 0.5) * (SET_BOOKMARK_WIDTH / 2) / SET_BOOKMARK_NOTCH;
                    if (legWidth <= 0) break;
                    g.newPath();
                    g.rectPath(left, y, legWidth, 1);
                    g.fillPath(fillBrush);
                    g.newPath();
                    g.rectPath(right - legWidth, y, legWidth, 1);
                    g.fillPath(fillBrush);
                }
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var cardPen = g.newPen(g.PenType.SOLID_COLOR, glyphColor, SET_CARD_STROKE_WIDTH);
            var backLeft = Math.round((w - SET_CARD_SIZE - SET_CARD_OFFSET) / 2);
            var backTop = Math.round((h - SET_CARD_SIZE - SET_CARD_OFFSET) / 2);
            var frontLeft = backLeft + SET_CARD_OFFSET;
            var frontTop = backTop + SET_CARD_OFFSET;
            var backRight = backLeft + SET_CARD_SIZE;
            var backBottom = backTop + SET_CARD_SIZE;
            g.newPath();
            g.moveTo(frontLeft, backBottom);
            g.lineTo(backLeft, backBottom);
            g.lineTo(backLeft, backTop);
            g.lineTo(backRight, backTop);
            g.lineTo(backRight, frontTop);
            g.strokePath(cardPen);
            g.newPath();
            g.rectPath(frontLeft, frontTop, SET_CARD_SIZE, SET_CARD_SIZE);
            g.strokePath(cardPen);
            var bookmarkLeft = frontLeft + SET_CARD_SIZE - SET_BOOKMARK_INSET - SET_BOOKMARK_WIDTH;
            drawBookmark(bookmarkLeft, frontTop, isOn);
            var numberFont = g.font;
            try { numberFont = ScriptUI.newFont("dialog", "BOLD", SET_NUMBER_FONT_SIZE); } catch (e) {}
            var numberText = String(setNumber);
            var numberSize = null;
            try { numberSize = g.measureString(numberText, numberFont); } catch (e2) {}
            if (!numberSize || typeof numberSize[0] !== "number") numberSize = [8, 15];
            var numberPen = g.newPen(g.PenType.SOLID_COLOR, glyphColor, 1);
            var numberLeft = frontLeft + Math.round((bookmarkLeft - frontLeft - numberSize[0]) / 2);
            var numberTop = frontTop + SET_CARD_SIZE - numberSize[1] - 2;
            g.drawString(numberText, numberPen, numberLeft, numberTop, numberFont);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "合体",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            fillRect(iconBrush, backX, backY, side, side);
            fillRect(iconBrush, frontX, frontY, side, side);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "前面型抜き",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            fillRect(iconBrush, backX, backY, side, side);
            fillRect(backgroundBrush, frontX, frontY, side, side);
            strokeRect(iconPen, frontX, frontY, side, side);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "交差",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            strokeRect(iconPen, backX, backY, side, side);
            strokeRect(iconPen, frontX, frontY, side, side);
            fillRect(iconBrush, frontX, frontY, overlapWidth, overlapHeight);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "中マド",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            fillRect(iconBrush, backX, backY, side, side);
            fillRect(iconBrush, frontX, frontY, side, side);
            fillRect(backgroundBrush, frontX, frontY, overlapWidth, overlapHeight);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "分割",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            fillRect(iconBrush, backX, backY, side, side);
            fillRect(iconBrush, frontX, frontY, side, side);
            strokeRect(backgroundPen, frontX, frontY, overlapWidth, overlapHeight);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "刈り込み",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            var trimSeam = 3;
            fillRect(iconBrush, backX, backY, side, side);
            fillRect(backgroundBrush, frontX - trimSeam, frontY - trimSeam, side + trimSeam, side + trimSeam);
            fillRect(iconBrush, frontX, frontY, side, side);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "合流",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            fillRect(iconBrush, backX, backY, side, side);
            fillRect(backgroundBrush, frontX - 2, frontY, 2, overlapHeight);
            fillRect(iconBrush, frontX, frontY, side, side);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "切り抜き",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            var cropBorder = 3;
            var cropBackPen = g.newPen(g.PenType.SOLID_COLOR, muted, cropBorder);
            strokeRect(cropBackPen, backX, backY, side, side);
            fillRect(iconBrush, frontX, frontY, side, side);
            fillRect(backgroundBrush, frontX + overlapWidth, frontY + cropBorder, side - cropBorder - overlapWidth, side - cropBorder * 2);
            fillRect(backgroundBrush, frontX + cropBorder, frontY + overlapHeight, side - cropBorder * 2, side - cropBorder - overlapHeight);
            var cropGap = cropBorder * 2;
            fillRect(backgroundBrush, (backX + side) - cropGap / 2, frontY - cropGap / 2, cropGap, cropGap);
            fillRect(backgroundBrush, frontX - cropGap / 2, (backY + side) - cropGap / 2, cropGap, cropGap);
            fillRect(iconBrush, frontX, frontY, overlapWidth, overlapHeight);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "アウトライン",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            strokeRect(iconPen, backX, backY, side, side);
            strokeRect(iconPen, frontX, frontY, side, side);
            var outlineGap = 5;
            fillRect(backgroundBrush, (backX + side) - outlineGap / 2, frontY - outlineGap / 2, outlineGap, outlineGap);
            fillRect(backgroundBrush, frontX - outlineGap / 2, (backY + side) - outlineGap / 2, outlineGap, outlineGap);
        }
    });

    /* ==== AiSmartPathfinderPalette (jsx/fx/AiSmartPathfinderPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiSmartPathfinderPalette",
        name: "背面型抜き",
        size: [46, 46],
        draw: function (g, w, h, ink, ground) {
            /* 中間色は元の配色（明:0.62 / 暗:0.55）に合わせ、ink と ground の中間 / muted sits halfway between ink and ground */
            var muted = [(ink[0] + ground[0]) / 2, (ink[1] + ground[1]) / 2, (ink[2] + ground[2]) / 2, 1];
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var backgroundBrush = g.newBrush(g.BrushType.SOLID_COLOR, ground);
            var iconPen = g.newPen(g.PenType.SOLID_COLOR, ink, 2);
            var backgroundPen = g.newPen(g.PenType.SOLID_COLOR, ground, 2);
            function fillRect(brush, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.fillPath(brush); }
            function strokeRect(pen, x, y, rw, rh) { g.newPath(); g.rectPath(x, y, rw, rh); g.strokePath(pen); }
            fillRect(backgroundBrush, 0, 0, w, h);
            var side = 22;
            var backX = 7, backY = 7;
            var frontX = 17, frontY = 17;
            var overlapWidth = (backX + side) - frontX;
            var overlapHeight = (backY + side) - frontY;
            fillRect(iconBrush, frontX, frontY, side, side);
            fillRect(backgroundBrush, backX, backY, side, side);
            strokeRect(iconPen, backX, backY, side, side);
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "台形：上辺を狭く",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0.2, 0], [0.8, 0], [0, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "台形：上辺を広く",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[-0.2, 0], [1.2, 0], [0, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "台形：下辺を狭く",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, 0], [0.2, 1], [0.8, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "台形：下辺を広く",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, 0], [-0.2, 1], [1.2, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "三角形：直角を左上に",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, 0], [0, 1], [0, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "三角形：直角を右上に",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, 0], [1, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "三角形：直角を左下に",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [0, 0], [0, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "三角形：直角を右下に",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[1, 0], [1, 0], [0, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "対角線：左上から右下へ",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [0, 0], [1, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "対角線：右上から左下へ",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[1, 0], [1, 0], [0, 1], [0, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：横にずらす（固定：左上）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, 0], [0.4, 1], [1.4, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：横にずらす（固定：右上）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, 0], [-0.4, 1], [0.6, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：横にずらす（固定：左下）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0.4, 0], [1.4, 0], [0, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：横にずらす（固定：右下）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[-0.4, 0], [0.6, 0], [0, 1], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：縦にずらす（固定：左上）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, 0.4], [0, 1], [1, 1.4]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：縦にずらす（固定：右上）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0.4], [1, 0], [0, 1.4], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：縦にずらす（固定：左下）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, 0], [1, -0.4], [0, 1], [1, 0.6]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== SmartFreeDistort (jsx/fx/SmartFreeDistort.jsx) ==== */
    ICON_CATALOG.push({
        script: "SmartFreeDistort",
        name: "平行四辺形：縦にずらす（固定：右下）",
        size: [40, 40],
        draw: function (g, w, h, ink, ground) {
            /* 変形後の4隅 [TL, TR, BL, BR]（アイコン用の誇張したサンプル量：台形0.20・シアー0.40）*/
            var corners = [[0, -0.4], [1, 0], [0, 0.6], [1, 1]];
            var SOURCE_CORNERS = [[0, 0], [1, 0], [0, 1], [1, 1]];
            var ICON_BUTTON_PADDING = 5;
            var ICON_TILE_LINE_WIDTH = 2;
            var ICON_SHAPE_SHRINK = 0.94;
            var ICON_SHAPE_LINE_WIDTH = 1;
            var ICON_DIAGONAL_LINE_WIDTH = 3.5;
            var DEGENERATE_AREA_THRESHOLD = 0.0001;
            var TILE_FILL = [0.91, 0.91, 0.91, 1];
            var TILE_BORDER = [0.76, 0.76, 0.76, 1];
            function toPerimeterOrder(c) { return [c[0], c[1], c[3], c[2]]; }
            function shrinkAboutCenter(points, factor) {
                var r = [];
                for (var i = 0; i < points.length; i++) r.push([0.5 + (points[i][0] - 0.5) * factor, 0.5 + (points[i][1] - 0.5) * factor]);
                return r;
            }
            function getPolygonArea(points) {
                var d = 0;
                for (var i = 0; i < points.length; i++) {
                    var n = points[(i + 1) % points.length];
                    d += points[i][0] * n[1] - n[0] * points[i][1];
                }
                return Math.abs(d) / 2;
            }
            function makeMapper(points, areaLeft, areaTop, areaSide) {
                var minX = points[0][0], minY = points[0][1], maxX = points[0][0], maxY = points[0][1];
                for (var i = 1; i < points.length; i++) {
                    minX = Math.min(minX, points[i][0]); minY = Math.min(minY, points[i][1]);
                    maxX = Math.max(maxX, points[i][0]); maxY = Math.max(maxY, points[i][1]);
                }
                var bw = maxX - minX, bh = maxY - minY;
                var scale = areaSide / Math.max(bw, bh);
                var offsetX = areaLeft + (areaSide - bw * scale) / 2 - minX * scale;
                var offsetY = areaTop + (areaSide - bh * scale) / 2 - minY * scale;
                return function (p) { return [offsetX + p[0] * scale, offsetY + p[1] * scale]; };
            }
            function tracePolygon(points, toPixel) {
                g.newPath();
                for (var i = 0; i < points.length; i++) {
                    var pp = toPixel(points[i]);
                    if (i === 0) g.moveTo(pp[0], pp[1]); else g.lineTo(pp[0], pp[1]);
                }
                g.closePath();
            }
            var areaSide = Math.min(w, h) - ICON_BUTTON_PADDING * 2;
            var areaLeft = (w - areaSide) / 2;
            var areaTop = (h - areaSide) / 2;
            var sourceSquare = toPerimeterOrder(SOURCE_CORNERS);
            var shape = shrinkAboutCenter(toPerimeterOrder(corners), ICON_SHAPE_SHRINK);
            var toPixel = makeMapper(sourceSquare.concat(shape), areaLeft, areaTop, areaSide);
            tracePolygon(sourceSquare, toPixel);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, TILE_FILL));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, TILE_BORDER, ICON_TILE_LINE_WIDTH));
            tracePolygon(shape, toPixel);
            if (getPolygonArea(shape) < DEGENERATE_AREA_THRESHOLD) {
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_DIAGONAL_LINE_WIDTH));
                return;
            }
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, ICON_SHAPE_LINE_WIDTH));
        }
    });

    /* ==== LinkedImageManagerPalette (jsx/link/LinkedImageManagerPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "LinkedImageManagerPalette",
        name: "前のアートボード（◀）",
        size: [22, 22],
        draw: function (g, w, h, ink, ground) {
            var direction = -1;
            /* 枠は三角より淡い（明: #666 / #999）。ink を ground 側へ寄せて作る / Frame is lighter than the mark */
            var frame = [ink[0] + (ground[0] - ink[0]) * 0.4, ink[1] + (ground[1] - ink[1]) * 0.4, ink[2] + (ground[2] - ink[2]) * 0.4, 1];
            var pen = g.newPen(g.PenType.SOLID_COLOR, frame, 1);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            g.newPath();
            g.moveTo(0.5, 0.5);
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(0.5, h - 0.5);
            g.closePath();
            g.strokePath(pen);
            var cx = w / 2, cy = h / 2;
            var half = Math.min(w, h) * 0.26;
            g.newPath();
            if (direction < 0) {
                g.moveTo(cx - half, cy);
                g.lineTo(cx + half, cy - half);
                g.lineTo(cx + half, cy + half);
            } else {
                g.moveTo(cx + half, cy);
                g.lineTo(cx - half, cy - half);
                g.lineTo(cx - half, cy + half);
            }
            g.closePath();
            g.fillPath(brush);
        }
    });

    /* ==== LinkedImageManagerPalette (jsx/link/LinkedImageManagerPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "LinkedImageManagerPalette",
        name: "次のアートボード（▶）",
        size: [22, 22],
        draw: function (g, w, h, ink, ground) {
            var direction = 1;
            /* 枠は三角より淡い（明: #666 / #999）。ink を ground 側へ寄せて作る / Frame is lighter than the mark */
            var frame = [ink[0] + (ground[0] - ink[0]) * 0.4, ink[1] + (ground[1] - ink[1]) * 0.4, ink[2] + (ground[2] - ink[2]) * 0.4, 1];
            var pen = g.newPen(g.PenType.SOLID_COLOR, frame, 1);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            g.newPath();
            g.moveTo(0.5, 0.5);
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(0.5, h - 0.5);
            g.closePath();
            g.strokePath(pen);
            var cx = w / 2, cy = h / 2;
            var half = Math.min(w, h) * 0.26;
            g.newPath();
            if (direction < 0) {
                g.moveTo(cx - half, cy);
                g.lineTo(cx + half, cy - half);
                g.lineTo(cx + half, cy + half);
            } else {
                g.moveTo(cx + half, cy);
                g.lineTo(cx - half, cy - half);
                g.lineTo(cx - half, cy + half);
            }
            g.closePath();
            g.fillPath(brush);
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：旗（未選択）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "flag";
            var isSelected = false;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：上昇（未選択）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "rise";
            var isSelected = false;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：直線（未選択）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "riseStraight";
            var isSelected = false;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：ギザギザ（未選択）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "jagged";
            var isSelected = false;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：旗（選択中）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "flag";
            var isSelected = true;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：上昇（選択中）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "rise";
            var isSelected = true;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：直線（選択中）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "riseStraight";
            var isSelected = true;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== TrimWithBreakLine (jsx/mask/TrimWithBreakLine.jsx) ==== */
    ICON_CATALOG.push({
        script: "TrimWithBreakLine",
        name: "切り口：ギザギザ（選択中）",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var STYLE_ICON_SIZE = 24;
            var STYLE_ICON_CUT_WIDTH = 2.5;
            var warpStyleKey = "jagged";
            var isSelected = true;
            function getStyleIconCutPoints(warpStyleKey) {
                var cutPoints = [];
                var stepCount = 24;
                var stepIndex;
                var ratioX;
                if (warpStyleKey === "flag") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
                    }
                } else if (warpStyleKey === "rise") {
                    for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                        ratioX = stepIndex / stepCount;
                        cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
                    }
                } else if (warpStyleKey === "riseStraight") {
                    cutPoints.push([0, 0.6], [1, 0.42]);
                } else {
                    var ridgeCount = 6;
                    for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                        cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
                    }
                }
                return cutPoints;
            }
            function sampleCutLine(cutPoints, ratioX) {
                for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
                    var startPoint = cutPoints[pointIndex - 1];
                    var endPoint = cutPoints[pointIndex];
                    if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
                    var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
                    return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
                }
                return { y: cutPoints[0][1], slope: 0 };
            }
            /* 未選択の図形はUIの明暗によらずグレー、選択中は前景色＋青枠 */
            var shapeBrush = g.newBrush(g.BrushType.SOLID_COLOR, isSelected ? ink : [0.55, 0.55, 0.55, 1]);
            var originX = Math.round((w - STYLE_ICON_SIZE) / 2);
            var originY = Math.round((h - STYLE_ICON_SIZE) / 2);
            var cutPoints = getStyleIconCutPoints(warpStyleKey);
            for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
                var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
                var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
                var cutY = cutSample.y * STYLE_ICON_SIZE;
                var upperHeight = cutY - halfGap;
                var lowerTop = cutY + halfGap;
                if (upperHeight > 0) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY, 1, upperHeight);
                    g.fillPath(shapeBrush);
                }
                if (lowerTop < STYLE_ICON_SIZE) {
                    g.newPath();
                    g.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                    g.fillPath(shapeBrush);
                }
            }
            if (isSelected) {
                g.newPath();
                g.rectPath(1, 1, w - 2, h - 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [0.2, 0.5, 0.95, 1], 2));
            }
        }
    });

    /* ==== AiCommandPrefLookup (jsx/misc/AiCommandPrefLookup.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiCommandPrefLookup",
        name: "コピー",
        size: [22, 22],
        draw: function (g, w, h, ink, ground) {
            function fillRoundedRect(iconGraphics, iconBrush, rectBounds, cornerRadius) {
                var left = rectBounds[0];
                var top = rectBounds[1];
                var width = rectBounds[2];
                var height = rectBounds[3];
                var diameter = cornerRadius * 2;
                var crossRects = [
                    [left + cornerRadius, top, width - diameter, height],
                    [left, top + cornerRadius, width, height - diameter]
                ];
                for (var i = 0; i < crossRects.length; i++) {
                    iconGraphics.newPath();
                    iconGraphics.rectPath(crossRects[i][0], crossRects[i][1], crossRects[i][2], crossRects[i][3]);
                    iconGraphics.fillPath(iconBrush);
                }
                var cornerOrigins = [
                    [left, top], [left + width - diameter, top],
                    [left, top + height - diameter], [left + width - diameter, top + height - diameter]
                ];
                for (var j = 0; j < cornerOrigins.length; j++) {
                    iconGraphics.newPath();
                    iconGraphics.ellipsePath(cornerOrigins[j][0], cornerOrigins[j][1], diameter, diameter);
                    iconGraphics.fillPath(iconBrush);
                }
            }
            var iconSize = w;
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var scaleUnit = iconSize / 22;
            function scaled(baseValue) {
                return Math.max(1, Math.round(baseValue * scaleUnit));
            }
            /* 背面の L 字 / Back sheet */
            var backBars = [[4, 7, 2, 11], [4, 16, 11, 2]];
            for (var i = 0; i < backBars.length; i++) {
                g.newPath();
                g.rectPath(scaled(backBars[i][0]), scaled(backBars[i][1]), scaled(backBars[i][2]), scaled(backBars[i][3]));
                g.fillPath(iconBrush);
            }
            /* 前面の角丸四角 / Front sheet */
            fillRoundedRect(g, iconBrush, [scaled(8), scaled(3), scaled(11), scaled(11)], scaled(2));
        }
    });

    /* ==== AiScriptLauncher (jsx/misc/AiScriptLauncher.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiScriptLauncher",
        name: "キーワードをクリア",
        size: [20, 20],
        draw: function (g, w, h, ink, ground) {
            var CLEAR_CIRCLE_INSET = 2;
            var CLEAR_GLYPH_INSET = 6;
            var CLEAR_STROKE_WIDTH = 1.5;
            var buttonSize = w;
            g.newPath();
            g.rectPath(0, 0, buttonSize, buttonSize);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var glyphPen = g.newPen(g.PenType.SOLID_COLOR, ink, CLEAR_STROKE_WIDTH);
            var circleSize = buttonSize - CLEAR_CIRCLE_INSET * 2;
            g.newPath();
            g.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
            g.strokePath(glyphPen);
            var glyphStart = CLEAR_GLYPH_INSET;
            var glyphEnd = buttonSize - CLEAR_GLYPH_INSET;
            g.newPath();
            g.moveTo(glyphStart, glyphStart);
            g.lineTo(glyphEnd, glyphEnd);
            g.strokePath(glyphPen);
            g.newPath();
            g.moveTo(glyphEnd, glyphStart);
            g.lineTo(glyphStart, glyphEnd);
            g.strokePath(glyphPen);
        }
    });

    /* ==== AiAnchorPointMarker (jsx/path/AiAnchorPointMarker.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiAnchorPointMarker",
        name: "基準点（中央を選択）",
        size: [56, 56],
        draw: function (g, w, h, ink, ground) {
            var registrationIndex = 4; /* 既定は中央 / default: center */
            var foreColor = ink;
            var cellSize = 8;
            var cellGap = 6;
            var cellStep = cellSize + cellGap;
            var gridSize = cellSize * 3 + cellGap * 2;
            var originX = Math.round((w - gridSize) / 2);
            var originY = Math.round((h - gridSize) / 2);
            function cellX(index) { return originX + (index % 3) * cellStep; }
            function cellY(index) { return originY + Math.floor(index / 3) * cellStep; }
            function buildCellPath(x, y, size) {
                g.newPath();
                g.moveTo(x, y);
                g.lineTo(x + size, y);
                g.lineTo(x + size, y + size);
                g.lineTo(x, y + size);
                g.closePath();
            }
            function drawRegistrationCell(x, y, size, selected) {
                if (selected) {
                    buildCellPath(x, y, size);
                    g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, foreColor));
                }
                buildCellPath(x, y, size);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, foreColor, 1));
            }
            /* 中央(4)を除く外周の□どうしをケイ線でつなぐ */
            var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
            var linePen = g.newPen(g.PenType.SOLID_COLOR, foreColor, 1);
            for (var i = 0; i < connections.length; i++) {
                var startCell = connections[i][0];
                var endCell = connections[i][1];
                g.newPath();
                if (endCell - startCell === 1) {
                    g.moveTo(cellX(startCell) + cellSize, cellY(startCell) + cellSize / 2);
                    g.lineTo(cellX(endCell), cellY(endCell) + cellSize / 2);
                } else {
                    g.moveTo(cellX(startCell) + cellSize / 2, cellY(startCell) + cellSize);
                    g.lineTo(cellX(endCell) + cellSize / 2, cellY(endCell));
                }
                g.strokePath(linePen);
            }
            for (var index = 0; index < 9; index++) {
                drawRegistrationCell(cellX(index), cellY(index), cellSize, index === registrationIndex);
            }
        }
    });

    /* ==== AspectRatioScaler (jsx/shape/AspectRatioScaler.jsx) ==== */
    ICON_CATALOG.push({
        script: "AspectRatioScaler",
        name: "向き：縦",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var ORIENT_FRAME_LONG = 30;
            var ORIENT_FRAME_SHORT = 23;
            var isPortrait = true;
            function fillRect(rect, color) {
                g.newPath();
                g.rectPath(rect[0], rect[1], rect[2], rect[3]);
                g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, color));
            }
            function strokeRect(rect, color, lineWidth) {
                g.newPath();
                g.rectPath(rect[0], rect[1], rect[2], rect[3]);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, color, lineWidth));
            }
            function fillEllipse(rect, color) {
                g.newPath();
                g.ellipsePath(rect[0], rect[1], rect[2], rect[3]);
                g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, color));
            }
            function drawPortraitFigure(frameRect, color) {
                var centerX = frameRect[0] + frameRect[2] / 2;
                var figureBottom = frameRect[1] + frameRect[3] - 3;
                var headSize = Math.round(frameRect[3] * 0.3);
                var headTop = frameRect[1] + Math.round(frameRect[3] * 0.18);
                var bodyWidth = Math.round(headSize * 1.9);
                var shoulderTop = headTop + headSize + 1;
                var shoulderHeight = Math.round(headSize * 0.9);
                fillEllipse([centerX - headSize / 2, headTop, headSize, headSize], color);
                fillEllipse([centerX - bodyWidth / 2, shoulderTop, bodyWidth, shoulderHeight], color);
                var torsoTop = shoulderTop + shoulderHeight / 2;
                fillRect([centerX - bodyWidth / 2, torsoTop, bodyWidth, figureBottom - torsoTop], color);
            }
            /* 地をパネルと同じ色で塗る / paint the background */
            fillRect([0, 0, w, h], ground);
            var frameWidth = isPortrait ? ORIENT_FRAME_SHORT : ORIENT_FRAME_LONG;
            var frameHeight = isPortrait ? ORIENT_FRAME_LONG : ORIENT_FRAME_SHORT;
            var frameRect = [
                Math.round((w - frameWidth) / 2),
                Math.round((h - frameHeight) / 2),
                frameWidth,
                frameHeight
            ];
            /* 未選択の色で描く（選択中は [0.22, 0.47, 0.9, 1] の青） */
            var iconColor = ink;
            strokeRect(frameRect, iconColor, 2);
            drawPortraitFigure(frameRect, iconColor);
        }
    });

    /* ==== AspectRatioScaler (jsx/shape/AspectRatioScaler.jsx) ==== */
    ICON_CATALOG.push({
        script: "AspectRatioScaler",
        name: "向き：横",
        size: [36, 36],
        draw: function (g, w, h, ink, ground) {
            var ORIENT_FRAME_LONG = 30;
            var ORIENT_FRAME_SHORT = 23;
            var isPortrait = false;
            function fillRect(rect, color) {
                g.newPath();
                g.rectPath(rect[0], rect[1], rect[2], rect[3]);
                g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, color));
            }
            function strokeRect(rect, color, lineWidth) {
                g.newPath();
                g.rectPath(rect[0], rect[1], rect[2], rect[3]);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, color, lineWidth));
            }
            function fillEllipse(rect, color) {
                g.newPath();
                g.ellipsePath(rect[0], rect[1], rect[2], rect[3]);
                g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, color));
            }
            function drawPortraitFigure(frameRect, color) {
                var centerX = frameRect[0] + frameRect[2] / 2;
                var figureBottom = frameRect[1] + frameRect[3] - 3;
                var headSize = Math.round(frameRect[3] * 0.3);
                var headTop = frameRect[1] + Math.round(frameRect[3] * 0.18);
                var bodyWidth = Math.round(headSize * 1.9);
                var shoulderTop = headTop + headSize + 1;
                var shoulderHeight = Math.round(headSize * 0.9);
                fillEllipse([centerX - headSize / 2, headTop, headSize, headSize], color);
                fillEllipse([centerX - bodyWidth / 2, shoulderTop, bodyWidth, shoulderHeight], color);
                var torsoTop = shoulderTop + shoulderHeight / 2;
                fillRect([centerX - bodyWidth / 2, torsoTop, bodyWidth, figureBottom - torsoTop], color);
            }
            /* 地をパネルと同じ色で塗る / paint the background */
            fillRect([0, 0, w, h], ground);
            var frameWidth = isPortrait ? ORIENT_FRAME_SHORT : ORIENT_FRAME_LONG;
            var frameHeight = isPortrait ? ORIENT_FRAME_LONG : ORIENT_FRAME_SHORT;
            var frameRect = [
                Math.round((w - frameWidth) / 2),
                Math.round((h - frameHeight) / 2),
                frameWidth,
                frameHeight
            ];
            /* 未選択の色で描く（選択中は [0.22, 0.47, 0.9, 1] の青） */
            var iconColor = ink;
            strokeRect(frameRect, iconColor, 2);
            drawPortraitFigure(frameRect, iconColor);
        }
    });

    /* ==== AiConnectorBuilder (jsx/stroke-table/AiConnectorBuilder.jsx) ==== */
    ICON_CATALOG.push({
        script: "AiConnectorBuilder",
        name: "線（矢印なし）",
        size: [24, 24],
        draw: function (g, w, h, ink, ground) {
            var arrowIndex = 0;
            var ICON_PADDING = 4;
            var ICON_BORDER_COLOR = [0.65, 0.65, 0.65, 1];
            var backColor = ground;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, backColor));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));
            var color = ink;
            var thickPen = g.newPen(g.PenType.SOLID_COLOR, color, 3);
            var thinPen = g.newPen(g.PenType.SOLID_COLOR, color, 1);
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, color);
            var backBrush = g.newBrush(g.BrushType.SOLID_COLOR, backColor);
            var left = ICON_PADDING;
            var right = w - ICON_PADDING;
            var centerY = Math.round(h / 2);
            var headSize = 5;
            var dotRadius = 4;
            function drawShaft(endX, pen) {
                g.newPath();
                g.moveTo(left, centerY);
                g.lineTo(endX, centerY);
                g.strokePath(pen);
            }
            if (arrowIndex === 0) {
                drawShaft(right, thickPen);
                return;
            }
            if (arrowIndex === 1) {
                drawShaft(right - headSize - 1, thickPen);
                g.newPath();
                g.moveTo(right - headSize - 1, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize - 1, centerY + headSize);
                g.closePath();
                g.fillPath(iconBrush);
                return;
            }
            if (arrowIndex === 2) {
                drawShaft(right, thinPen);
                g.newPath();
                g.moveTo(right - headSize, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize, centerY + headSize);
                g.strokePath(thinPen);
                return;
            }
            drawShaft(right - dotRadius * 2 + 1, thickPen);
            g.newPath();
            g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
            g.fillPath(arrowIndex === 3 ? iconBrush : backBrush);
            if (arrowIndex === 4) {
                g.newPath();
                g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, color, 2));
            }
        }
    });

    ICON_CATALOG.push({
        script: "AiConnectorBuilder",
        name: "実線の矢印",
        size: [24, 24],
        draw: function (g, w, h, ink, ground) {
            var arrowIndex = 1;
            var ICON_PADDING = 4;
            var ICON_BORDER_COLOR = [0.65, 0.65, 0.65, 1];
            var backColor = ground;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, backColor));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));
            var color = ink;
            var thickPen = g.newPen(g.PenType.SOLID_COLOR, color, 3);
            var thinPen = g.newPen(g.PenType.SOLID_COLOR, color, 1);
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, color);
            var backBrush = g.newBrush(g.BrushType.SOLID_COLOR, backColor);
            var left = ICON_PADDING;
            var right = w - ICON_PADDING;
            var centerY = Math.round(h / 2);
            var headSize = 5;
            var dotRadius = 4;
            function drawShaft(endX, pen) {
                g.newPath();
                g.moveTo(left, centerY);
                g.lineTo(endX, centerY);
                g.strokePath(pen);
            }
            if (arrowIndex === 0) {
                drawShaft(right, thickPen);
                return;
            }
            if (arrowIndex === 1) {
                drawShaft(right - headSize - 1, thickPen);
                g.newPath();
                g.moveTo(right - headSize - 1, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize - 1, centerY + headSize);
                g.closePath();
                g.fillPath(iconBrush);
                return;
            }
            if (arrowIndex === 2) {
                drawShaft(right, thinPen);
                g.newPath();
                g.moveTo(right - headSize, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize, centerY + headSize);
                g.strokePath(thinPen);
                return;
            }
            drawShaft(right - dotRadius * 2 + 1, thickPen);
            g.newPath();
            g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
            g.fillPath(arrowIndex === 3 ? iconBrush : backBrush);
            if (arrowIndex === 4) {
                g.newPath();
                g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, color, 2));
            }
        }
    });

    ICON_CATALOG.push({
        script: "AiConnectorBuilder",
        name: "線の矢印",
        size: [24, 24],
        draw: function (g, w, h, ink, ground) {
            var arrowIndex = 2;
            var ICON_PADDING = 4;
            var ICON_BORDER_COLOR = [0.65, 0.65, 0.65, 1];
            var backColor = ground;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, backColor));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));
            var color = ink;
            var thickPen = g.newPen(g.PenType.SOLID_COLOR, color, 3);
            var thinPen = g.newPen(g.PenType.SOLID_COLOR, color, 1);
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, color);
            var backBrush = g.newBrush(g.BrushType.SOLID_COLOR, backColor);
            var left = ICON_PADDING;
            var right = w - ICON_PADDING;
            var centerY = Math.round(h / 2);
            var headSize = 5;
            var dotRadius = 4;
            function drawShaft(endX, pen) {
                g.newPath();
                g.moveTo(left, centerY);
                g.lineTo(endX, centerY);
                g.strokePath(pen);
            }
            if (arrowIndex === 0) {
                drawShaft(right, thickPen);
                return;
            }
            if (arrowIndex === 1) {
                drawShaft(right - headSize - 1, thickPen);
                g.newPath();
                g.moveTo(right - headSize - 1, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize - 1, centerY + headSize);
                g.closePath();
                g.fillPath(iconBrush);
                return;
            }
            if (arrowIndex === 2) {
                drawShaft(right, thinPen);
                g.newPath();
                g.moveTo(right - headSize, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize, centerY + headSize);
                g.strokePath(thinPen);
                return;
            }
            drawShaft(right - dotRadius * 2 + 1, thickPen);
            g.newPath();
            g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
            g.fillPath(arrowIndex === 3 ? iconBrush : backBrush);
            if (arrowIndex === 4) {
                g.newPath();
                g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, color, 2));
            }
        }
    });

    ICON_CATALOG.push({
        script: "AiConnectorBuilder",
        name: "黒丸",
        size: [24, 24],
        draw: function (g, w, h, ink, ground) {
            var arrowIndex = 3;
            var ICON_PADDING = 4;
            var ICON_BORDER_COLOR = [0.65, 0.65, 0.65, 1];
            var backColor = ground;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, backColor));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));
            var color = ink;
            var thickPen = g.newPen(g.PenType.SOLID_COLOR, color, 3);
            var thinPen = g.newPen(g.PenType.SOLID_COLOR, color, 1);
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, color);
            var backBrush = g.newBrush(g.BrushType.SOLID_COLOR, backColor);
            var left = ICON_PADDING;
            var right = w - ICON_PADDING;
            var centerY = Math.round(h / 2);
            var headSize = 5;
            var dotRadius = 4;
            function drawShaft(endX, pen) {
                g.newPath();
                g.moveTo(left, centerY);
                g.lineTo(endX, centerY);
                g.strokePath(pen);
            }
            if (arrowIndex === 0) {
                drawShaft(right, thickPen);
                return;
            }
            if (arrowIndex === 1) {
                drawShaft(right - headSize - 1, thickPen);
                g.newPath();
                g.moveTo(right - headSize - 1, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize - 1, centerY + headSize);
                g.closePath();
                g.fillPath(iconBrush);
                return;
            }
            if (arrowIndex === 2) {
                drawShaft(right, thinPen);
                g.newPath();
                g.moveTo(right - headSize, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize, centerY + headSize);
                g.strokePath(thinPen);
                return;
            }
            drawShaft(right - dotRadius * 2 + 1, thickPen);
            g.newPath();
            g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
            g.fillPath(arrowIndex === 3 ? iconBrush : backBrush);
            if (arrowIndex === 4) {
                g.newPath();
                g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, color, 2));
            }
        }
    });

    ICON_CATALOG.push({
        script: "AiConnectorBuilder",
        name: "白丸",
        size: [24, 24],
        draw: function (g, w, h, ink, ground) {
            var arrowIndex = 4;
            var ICON_PADDING = 4;
            var ICON_BORDER_COLOR = [0.65, 0.65, 0.65, 1];
            var backColor = ground;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, backColor));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));
            var color = ink;
            var thickPen = g.newPen(g.PenType.SOLID_COLOR, color, 3);
            var thinPen = g.newPen(g.PenType.SOLID_COLOR, color, 1);
            var iconBrush = g.newBrush(g.BrushType.SOLID_COLOR, color);
            var backBrush = g.newBrush(g.BrushType.SOLID_COLOR, backColor);
            var left = ICON_PADDING;
            var right = w - ICON_PADDING;
            var centerY = Math.round(h / 2);
            var headSize = 5;
            var dotRadius = 4;
            function drawShaft(endX, pen) {
                g.newPath();
                g.moveTo(left, centerY);
                g.lineTo(endX, centerY);
                g.strokePath(pen);
            }
            if (arrowIndex === 0) {
                drawShaft(right, thickPen);
                return;
            }
            if (arrowIndex === 1) {
                drawShaft(right - headSize - 1, thickPen);
                g.newPath();
                g.moveTo(right - headSize - 1, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize - 1, centerY + headSize);
                g.closePath();
                g.fillPath(iconBrush);
                return;
            }
            if (arrowIndex === 2) {
                drawShaft(right, thinPen);
                g.newPath();
                g.moveTo(right - headSize, centerY - headSize);
                g.lineTo(right, centerY);
                g.lineTo(right - headSize, centerY + headSize);
                g.strokePath(thinPen);
                return;
            }
            drawShaft(right - dotRadius * 2 + 1, thickPen);
            g.newPath();
            g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
            g.fillPath(arrowIndex === 3 ? iconBrush : backBrush);
            if (arrowIndex === 4) {
                g.newPath();
                g.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, color, 2));
            }
        }
    });

    /* ==== FavoriteArrow (jsx/stroke-table/FavoriteArrow.jsx) ==== */
    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "線端：線端なし",
        size: [28, 22],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[340,360],"rects":[[123,20,320,340],[74,119,123,241]],"heads":[],"holes":[[99,144,172,216],[172,168,320,192]]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "線端：丸型線端",
        size: [28, 22],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[340,360],"rects":[[185,20,320,340]],"heads":[],"discs":[[185,180,160,"left"]],"holes":[[124,144,197,216],[197,168,320,192]]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "線端：突出線端",
        size: [28, 22],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[340,360],"rects":[[25,20,320,340]],"heads":[],"holes":[[123,144,196,216],[196,168,320,192]]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "角の形状：マイター結合",
        size: [28, 22],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[340,334],"rects":[[25,20,320,314]],"heads":[],"holes":[[246,241,320,314],[148,191,172,314],[123,118,197,191],[197,142,320,167]]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "角の形状：ラウンド結合",
        size: [28, 22],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[340,334],"rects":[[159,20,320,314],[25,154,159,314]],"heads":[],"discs":[[159,154,134,"topLeft"]],"holes":[[247,241,320,314],[148,191,173,314],[124,118,197,191],[197,142,320,167]]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "角の形状：ベベル結合",
        size: [28, 22],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[340,334],"rects":[[148,20,320,314],[25,142,50,314],[50,118,74,314],[74,93,99,314],[99,69,123,314],[123,44,148,314]],"heads":[],"holes":[[246,241,320,314],[148,191,172,314],[123,118,196,191],[196,142,320,167]]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "両端を調整：オフ（破線の長さを保持）",
        size: [36, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[563,450],"rects":[[110,53,281,110],[395,53,509,110],[452,110,509,282],[53,224,110,338],[53,338,167,395],[281,338,452,395]],"heads":[]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "両端を調整：オン（コーナーやパス先端に破線の先端を整列）",
        size: [36, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[563,450],"rects":[[53,53,225,110],[53,110,111,167],[338,53,510,110],[452,110,510,167],[53,281,111,338],[53,338,225,395],[452,281,510,338],[338,338,510,395]],"heads":[]};
            var inset = 3;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "先端位置：矢印の先端をパスの終点から配置",
        size: [36, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[573,420],"rects":[[140,114,293,139],[293,87,320,166],[155,250,335,309]],"heads":[[308,451,279,77]]};
            var inset = 0;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "先端位置：矢印の先端をパスの終点に配置",
        size: [36, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[573,420],"rects":[[172,114,440,139],[440,87,467,166],[187,250,365,309]],"heads":[[339,482,279,77]]};
            var inset = 0;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, undefined, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "よく使う矢印：なし（矢印0）",
        size: [76, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[883,383],"rects":[[70,179,812,205]],"heads":[]};
            var inset = 5;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, 22 / 383, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "よく使う矢印：矢印8",
        size: [76, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[883,383],"rects":[[150,154,812,230]],"heads":[[186,70,192,116]]};
            var inset = 5;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, 22 / 383, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "よく使う矢印：矢印11",
        size: [76, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[883,383],"rects":[[100,180,812,206]],"heads":[],"lines":[[[191,94],[73,193],[191,292],24]]};
            var inset = 5;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, 22 / 383, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "よく使う矢印：矢印27",
        size: [76, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[883,383],"rects":[[96,178,812,204],[66,109,100,274]],"heads":[]};
            var inset = 5;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, 22 / 383, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "よく使う矢印：矢印8（終点も同じ・両端）",
        size: [76, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[883,383],"rects":[[150,154,812,230]],"heads":[[186,70,192,116]]};
            iconShape = buildBothEndsShape(iconShape);
            var inset = 5;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, 22 / 383, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "よく使う矢印：矢印11（終点も同じ・両端）",
        size: [76, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[883,383],"rects":[[100,180,812,206]],"heads":[],"lines":[[[191,94],[73,193],[191,292],24]]};
            iconShape = buildBothEndsShape(iconShape);
            var inset = 5;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, 22 / 383, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "よく使う矢印：矢印27（終点も同じ・両端）",
        size: [76, 26],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            function addDiscPath(iconGraphics, centerX, centerY, radius, discSide) {
                var sliceCount = Math.max(8, Math.ceil(radius * 2));
                var sliceHeight = radius / sliceCount;
                var bottomSlices = (discSide === "left") ? sliceCount : 0;
                for (var k = -sliceCount; k < bottomSlices; k++) {
                    var sliceCenterY = (k + 0.5) * sliceHeight;
                    var halfWidth = Math.sqrt(Math.max(0, radius * radius - sliceCenterY * sliceCenterY));
                    iconGraphics.rectPath(centerX - halfWidth, centerY + k * sliceHeight, halfWidth, sliceHeight);
                }
            }
            function buildBothEndsShape(iconShape) {
                var shapeWidth = iconShape.size[0];
                var bothEndsShape = { size: iconShape.size, rects: [], heads: [], lines: [] };
                var i, j;
                for (i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    bothEndsShape.rects.push(shapeRect, [shapeWidth - shapeRect[2], shapeRect[1], shapeWidth - shapeRect[0], shapeRect[3]]);
                }
                for (i = 0; i < iconShape.heads.length; i++) {
                    var headShape = iconShape.heads[i];
                    bothEndsShape.heads.push(headShape, [shapeWidth - headShape[0], shapeWidth - headShape[1], headShape[2], headShape[3]]);
                }
                var shapeLines = iconShape.lines || [];
                for (i = 0; i < shapeLines.length; i++) {
                    var shapeLine = shapeLines[i];
                    var mirroredLine = [];
                    for (j = 0; j < shapeLine.length - 1; j++) mirroredLine.push([shapeWidth - shapeLine[j][0], shapeLine[j][1]]);
                    mirroredLine.push(shapeLine[shapeLine.length - 1]);
                    bothEndsShape.lines.push(shapeLine, mirroredLine);
                }
                return bothEndsShape;
            }
            function drawIconShapes(iconGraphics, iconShape, drawArea, inkColor, fixedScale, groundColor) {
                var iconScale = fixedScale || Math.min(drawArea[2] / iconShape.size[0], drawArea[3] / iconShape.size[1]);
                var scaleX = iconScale;
                var scaleY = iconScale;
                var originX = drawArea[0] + (drawArea[2] - iconShape.size[0] * iconScale) / 2;
                var originY = drawArea[1] + (drawArea[3] - iconShape.size[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < iconShape.rects.length; i++) {
                    var shapeRect = iconShape.rects[i];
                    var shapeLeft = Math.round(originX + shapeRect[0] * scaleX);
                    var shapeTop = Math.round(originY + shapeRect[1] * scaleY);
                    var shapeRight = Math.round(originX + shapeRect[2] * scaleX);
                    var shapeBottom = Math.round(originY + shapeRect[3] * scaleY);
                    iconGraphics.rectPath(shapeLeft, shapeTop, Math.max(1, shapeRight - shapeLeft), Math.max(1, shapeBottom - shapeTop));
                }
                for (var j = 0; j < iconShape.heads.length; j++) {
                    var headShape = iconShape.heads[j];
                    addArrowHeadPath(iconGraphics, originX + headShape[0] * scaleX, originX + headShape[1] * scaleX, originY + headShape[2] * scaleY, headShape[3] * scaleY);
                }
                var shapeDiscs = iconShape.discs || [];
                for (var d = 0; d < shapeDiscs.length; d++) {
                    var discShape = shapeDiscs[d];
                    addDiscPath(iconGraphics, originX + discShape[0] * iconScale, originY + discShape[1] * iconScale, discShape[2] * iconScale, discShape[3]);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));
                var shapeHoles = iconShape.holes || [];
                if (shapeHoles.length > 0 && groundColor) {
                    iconGraphics.newPath();
                    for (var hh = 0; hh < shapeHoles.length; hh++) {
                        var holeRect = shapeHoles[hh];
                        var holeLeft = Math.round(originX + holeRect[0] * scaleX);
                        var holeTop = Math.round(originY + holeRect[1] * scaleY);
                        iconGraphics.rectPath(holeLeft, holeTop,
                            Math.max(1, Math.round(originX + holeRect[2] * scaleX) - holeLeft), Math.max(1, Math.round(originY + holeRect[3] * scaleY) - holeTop));
                    }
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
                }
                var shapeLines = iconShape.lines || [];
                for (var k = 0; k < shapeLines.length; k++) {
                    var shapeLine = shapeLines[k];
                    var lineWidth = shapeLine[shapeLine.length - 1];
                    iconGraphics.newPath();
                    for (var m = 0; m < shapeLine.length - 1; m++) {
                        var pointX = originX + shapeLine[m][0] * scaleX;
                        var pointY = originY + shapeLine[m][1] * scaleY;
                        if (m === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, Math.max(1, lineWidth * iconScale)));
                }
            }
            /* 地と枠（drawOptionIcon と同じ。単独表示なので左の枠も描く）/ ground and frame as in drawOptionIcon */
            var hasLeftEdge = true;
            var groundLeft = hasLeftEdge ? 0 : -0.5;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            if (hasLeftEdge) {
                g.moveTo(0.5, h - 0.5);
                g.lineTo(0.5, 0.5);
            } else {
                g.moveTo(groundLeft, 0.5);
            }
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(hasLeftEdge ? 0.5 : groundLeft, h - 0.5);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, 1));
            var iconShape = {"size":[883,383],"rects":[[96,178,812,204],[66,109,100,274]],"heads":[]};
            iconShape = buildBothEndsShape(iconShape);
            var inset = 5;
            drawIconShapes(g, iconShape, [inset, inset, w - inset * 2, h - inset * 2], ink, 22 / 383, ground);
        }
    });

    ICON_CATALOG.push({
        script: "FavoriteArrow",
        name: "始点と終点を入れ替え（⇄）",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            function addArrowHeadPath(iconGraphics, baseX, tipX, centerY, halfHeight) {
                var sliceCount = 12;
                var sliceWidth = Math.abs(tipX - baseX) / sliceCount;
                var direction = (tipX > baseX) ? 1 : -1;
                for (var k = 0; k < sliceCount; k++) {
                    var sliceStart = baseX + direction * sliceWidth * k;
                    var sliceHalf = halfHeight * (1 - (k + 0.5) / sliceCount);
                    iconGraphics.rectPath(direction > 0 ? sliceStart : sliceStart - sliceWidth, centerY - sliceHalf, sliceWidth, sliceHalf * 2);
                }
            }
            /* オンのときだけ押し込んだ地と枠を描く。ここはオフの見た目 / pressed ground and frame only while on; this is the off look */
            var iconScale = Math.min(w, h) / 22;
            var offsetX = (w - 22 * iconScale) / 2;
            var offsetY = (h - 22 * iconScale) / 2;
            function mapX(x) { return offsetX + x * iconScale; }
            function mapY(y) { return offsetY + y * iconScale; }
            g.newPath();
            g.rectPath(mapX(5.5), mapY(7.2), 6 * iconScale, 2 * iconScale);
            addArrowHeadPath(g, mapX(11.5), mapX(17), mapY(8.2), 2.8 * iconScale);
            g.rectPath(mapX(9.3), mapY(13), 6 * iconScale, 2 * iconScale);
            addArrowHeadPath(g, mapX(9.3), mapX(3.8), mapY(14), 2.8 * iconScale);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ink));
        }
    });

    /* ==== AddBulletsAndNumbers (jsx/text/AddBulletsAndNumbers.jsx) ==== */
    ICON_CATALOG.push({
        script: "AddBulletsAndNumbers",
        name: "行揃え：左揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "left";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AddBulletsAndNumbers",
        name: "行揃え：中央揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "center";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AddBulletsAndNumbers",
        name: "行揃え：右揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "right";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AddBulletsAndNumbers",
        name: "行揃え：均等配置（最終行左揃え）",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "justifyLeft";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AddBulletsAndNumbers",
        name: "行揃え：両端揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "justifyAll";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    /* ==== AreaTypeToolkit (jsx/text/AreaTypeToolkit.jsx) ==== */
    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "行揃え：左揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "left";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "行揃え：中央揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "center";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "行揃え：右揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "right";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "行揃え：均等配置（最終行左揃え）",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "justifyLeft";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "行揃え：両端揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "justifyAll";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getJustifyLineWidths(t, longWidth, shortWidth) {
                if (t === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (t === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(t, buttonWidth, lineWidth) {
                var margin = 5;
                if (t === "right") return buttonWidth - margin - lineWidth;
                if (t === "center") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineStartX = getJustifyLineX(iconType, w, lineWidths[i]);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidths[i], rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "テキストの配置：上揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "top";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getAlignRowYs(t) {
                if (t === "center") return [9, 13, 17];
                if (t === "bottom") return [11, 15, 19];
                if (t === "justify") return [7, 13, 19];
                return [7, 11, 15];
            }
            var linePen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var lineWidth = 14;
            var lineStartX = Math.round((w - lineWidth) / 2);
            var rowYs = getAlignRowYs(iconType);
            for (var i = 0; i < rowYs.length; i++) {
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(linePen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "テキストの配置：中央揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "center";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getAlignRowYs(t) {
                if (t === "center") return [9, 13, 17];
                if (t === "bottom") return [11, 15, 19];
                if (t === "justify") return [7, 13, 19];
                return [7, 11, 15];
            }
            var linePen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var lineWidth = 14;
            var lineStartX = Math.round((w - lineWidth) / 2);
            var rowYs = getAlignRowYs(iconType);
            for (var i = 0; i < rowYs.length; i++) {
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(linePen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "テキストの配置：下揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "bottom";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getAlignRowYs(t) {
                if (t === "center") return [9, 13, 17];
                if (t === "bottom") return [11, 15, 19];
                if (t === "justify") return [7, 13, 19];
                return [7, 11, 15];
            }
            var linePen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var lineWidth = 14;
            var lineStartX = Math.round((w - lineWidth) / 2);
            var rowYs = getAlignRowYs(iconType);
            for (var i = 0; i < rowYs.length; i++) {
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(linePen);
            }
        }
    });

    ICON_CATALOG.push({
        script: "AreaTypeToolkit",
        name: "テキストの配置：均等配置",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "justify";
            var border = [0.62, 0.62, 0.62, 1];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));
            function getAlignRowYs(t) {
                if (t === "center") return [9, 13, 17];
                if (t === "bottom") return [11, 15, 19];
                if (t === "justify") return [7, 13, 19];
                return [7, 11, 15];
            }
            var linePen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var lineWidth = 14;
            var lineStartX = Math.round((w - lineWidth) / 2);
            var rowYs = getAlignRowYs(iconType);
            for (var i = 0; i < rowYs.length; i++) {
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(linePen);
            }
        }
    });

    /* ==== DynamicTextGenerator (jsx/text/DynamicTextGenerator.jsx) ==== */
    /* モード選択アイコン（ブロック／円／アーチ／下向き弓）。選択中の角丸の座布団は省き、未選択の見た目で描く */
    ICON_CATALOG.push({
        script: "DynamicTextGenerator",
        name: "ブロック",
        size: [40, 34],
        draw: function (g, w, h, ink, ground) {
            function fillRect(graphics, brush, x, y, rectWidth, rectHeight) {
                graphics.newPath();
                graphics.rectPath(x, y, rectWidth, rectHeight);
                graphics.fillPath(brush);
            }
            var centerX = w / 2;
            var centerY = h / 2;
            fillRect(g, g.newBrush(g.BrushType.SOLID_COLOR, ground), 0, 0, w, h);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, ink);
            var barWidth = 20;
            var barHeights = [6, 4, 7];
            var barGap = 2;
            var totalHeight = barGap * (barHeights.length - 1);
            var i;
            for (i = 0; i < barHeights.length; i++) totalHeight += barHeights[i];
            var y = centerY - totalHeight / 2;
            for (i = 0; i < barHeights.length; i++) {
                fillRect(g, brush, centerX - barWidth / 2, y, barWidth, barHeights[i]);
                y += barHeights[i] + barGap;
            }
        }
    });

    /* ==== DynamicTextGenerator (jsx/text/DynamicTextGenerator.jsx) ==== */
    ICON_CATALOG.push({
        script: "DynamicTextGenerator",
        name: "円",
        size: [40, 34],
        draw: function (g, w, h, ink, ground) {
            var MODE_ICON_RADIUS = 9;
            var MODE_ICON_STROKE = 2;
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var centerX = w / 2;
            var centerY = h / 2;
            var radius = MODE_ICON_RADIUS;
            g.newPath();
            g.ellipsePath(centerX - radius, centerY - radius, radius * 2, radius * 2);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, ink, MODE_ICON_STROKE));
        }
    });

    /* ==== DynamicTextGenerator (jsx/text/DynamicTextGenerator.jsx) ==== */
    ICON_CATALOG.push({
        script: "DynamicTextGenerator",
        name: "アーチ",
        size: [40, 34],
        draw: function (g, w, h, ink, ground) {
            var MODE_ICON_RADIUS = 9;
            var MODE_ICON_STROKE = 2;
            function arcPoint(cx, cy, r, deg) {
                var rad = deg * Math.PI / 180;
                return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
            }
            function strokeArc(graphics, pen, cx, cy, r, startAngle, endAngle) {
                var startPoint = arcPoint(cx, cy, r, startAngle);
                graphics.newPath();
                graphics.moveTo(startPoint[0], startPoint[1]);
                for (var i = 1; i <= 32; i++) {
                    var p = arcPoint(cx, cy, r, startAngle + (endAngle - startAngle) * (i / 32));
                    graphics.lineTo(p[0], p[1]);
                }
                graphics.strokePath(pen);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, MODE_ICON_STROKE);
            var radius = MODE_ICON_RADIUS;
            var overshoot = 20;
            var arcOffset = radius * (1 - Math.sin(overshoot * Math.PI / 180)) / 2;
            strokeArc(g, pen, w / 2, h / 2 + arcOffset, radius, 180 - overshoot, 360 + overshoot);
        }
    });

    /* ==== DynamicTextGenerator (jsx/text/DynamicTextGenerator.jsx) ==== */
    ICON_CATALOG.push({
        script: "DynamicTextGenerator",
        name: "下向き弓",
        size: [40, 34],
        draw: function (g, w, h, ink, ground) {
            var MODE_ICON_RADIUS = 9;
            var MODE_ICON_STROKE = 2;
            function arcPoint(cx, cy, r, deg) {
                var rad = deg * Math.PI / 180;
                return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
            }
            function strokeArc(graphics, pen, cx, cy, r, startAngle, endAngle) {
                var startPoint = arcPoint(cx, cy, r, startAngle);
                graphics.newPath();
                graphics.moveTo(startPoint[0], startPoint[1]);
                for (var i = 1; i <= 32; i++) {
                    var p = arcPoint(cx, cy, r, startAngle + (endAngle - startAngle) * (i / 32));
                    graphics.lineTo(p[0], p[1]);
                }
                graphics.strokePath(pen);
            }
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, MODE_ICON_STROKE);
            var radius = MODE_ICON_RADIUS;
            var overshoot = 20;
            var arcOffset = radius * (1 - Math.sin(overshoot * Math.PI / 180)) / 2;
            strokeArc(g, pen, w / 2, h / 2 - arcOffset, radius, -overshoot, 180 + overshoot);
        }
    });

    /* ==== UnifiedTypePalette (jsx/text/UnifiedTypePalette.jsx) ==== */
    /* 行揃えボタン。未選択の見た目。枠線はライトUIのみの描画で、色は地とインクの中間に置き換え */
    ICON_CATALOG.push({
        script: "UnifiedTypePalette",
        name: "左揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "left";
            function getJustifyLineWidths(type, longWidth, shortWidth) {
                if (type === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (type === "justifyLeft" || type === "justifyCenter" || type === "justifyRight") {
                    return [longWidth, longWidth, longWidth, shortWidth];
                }
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(type, buttonWidth, lineWidth) {
                var margin = 5;
                if (type === "right" || type === "justifyRight") return buttonWidth - margin - lineWidth;
                if (type === "center" || type === "justifyCenter") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var border = [
                ground[0] + (ink[0] - ground[0]) * 0.5,
                ground[1] + (ink[1] - ground[1]) * 0.5,
                ground[2] + (ink[2] - ground[2]) * 0.5,
                1
            ];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));

            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineWidth = lineWidths[i];
                var lineStartX = getJustifyLineX(iconType, w, lineWidth);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    /* ==== UnifiedTypePalette (jsx/text/UnifiedTypePalette.jsx) ==== */
    /* 行揃えボタン。未選択の見た目。枠線はライトUIのみの描画で、色は地とインクの中間に置き換え */
    ICON_CATALOG.push({
        script: "UnifiedTypePalette",
        name: "中央揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "center";
            function getJustifyLineWidths(type, longWidth, shortWidth) {
                if (type === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (type === "justifyLeft" || type === "justifyCenter" || type === "justifyRight") {
                    return [longWidth, longWidth, longWidth, shortWidth];
                }
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(type, buttonWidth, lineWidth) {
                var margin = 5;
                if (type === "right" || type === "justifyRight") return buttonWidth - margin - lineWidth;
                if (type === "center" || type === "justifyCenter") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var border = [
                ground[0] + (ink[0] - ground[0]) * 0.5,
                ground[1] + (ink[1] - ground[1]) * 0.5,
                ground[2] + (ink[2] - ground[2]) * 0.5,
                1
            ];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));

            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineWidth = lineWidths[i];
                var lineStartX = getJustifyLineX(iconType, w, lineWidth);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    /* ==== UnifiedTypePalette (jsx/text/UnifiedTypePalette.jsx) ==== */
    /* 行揃えボタン。未選択の見た目。枠線はライトUIのみの描画で、色は地とインクの中間に置き換え */
    ICON_CATALOG.push({
        script: "UnifiedTypePalette",
        name: "右揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "right";
            function getJustifyLineWidths(type, longWidth, shortWidth) {
                if (type === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (type === "justifyLeft" || type === "justifyCenter" || type === "justifyRight") {
                    return [longWidth, longWidth, longWidth, shortWidth];
                }
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(type, buttonWidth, lineWidth) {
                var margin = 5;
                if (type === "right" || type === "justifyRight") return buttonWidth - margin - lineWidth;
                if (type === "center" || type === "justifyCenter") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var border = [
                ground[0] + (ink[0] - ground[0]) * 0.5,
                ground[1] + (ink[1] - ground[1]) * 0.5,
                ground[2] + (ink[2] - ground[2]) * 0.5,
                1
            ];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));

            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineWidth = lineWidths[i];
                var lineStartX = getJustifyLineX(iconType, w, lineWidth);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    /* ==== UnifiedTypePalette (jsx/text/UnifiedTypePalette.jsx) ==== */
    /* 行揃えボタン。未選択の見た目。枠線はライトUIのみの描画で、色は地とインクの中間に置き換え */
    ICON_CATALOG.push({
        script: "UnifiedTypePalette",
        name: "均等配置（最終行左）",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "justifyLeft";
            function getJustifyLineWidths(type, longWidth, shortWidth) {
                if (type === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (type === "justifyLeft" || type === "justifyCenter" || type === "justifyRight") {
                    return [longWidth, longWidth, longWidth, shortWidth];
                }
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(type, buttonWidth, lineWidth) {
                var margin = 5;
                if (type === "right" || type === "justifyRight") return buttonWidth - margin - lineWidth;
                if (type === "center" || type === "justifyCenter") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var border = [
                ground[0] + (ink[0] - ground[0]) * 0.5,
                ground[1] + (ink[1] - ground[1]) * 0.5,
                ground[2] + (ink[2] - ground[2]) * 0.5,
                1
            ];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));

            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineWidth = lineWidths[i];
                var lineStartX = getJustifyLineX(iconType, w, lineWidth);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    /* ==== UnifiedTypePalette (jsx/text/UnifiedTypePalette.jsx) ==== */
    /* 行揃えボタン。未選択の見た目。枠線はライトUIのみの描画で、色は地とインクの中間に置き換え */
    ICON_CATALOG.push({
        script: "UnifiedTypePalette",
        name: "両端揃え",
        size: [26, 26],
        draw: function (g, w, h, ink, ground) {
            var iconType = "justifyAll";
            function getJustifyLineWidths(type, longWidth, shortWidth) {
                if (type === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
                if (type === "justifyLeft" || type === "justifyCenter" || type === "justifyRight") {
                    return [longWidth, longWidth, longWidth, shortWidth];
                }
                return [longWidth, shortWidth, longWidth, shortWidth];
            }
            function getJustifyLineX(type, buttonWidth, lineWidth) {
                var margin = 5;
                if (type === "right" || type === "justifyRight") return buttonWidth - margin - lineWidth;
                if (type === "center" || type === "justifyCenter") return Math.round((buttonWidth - lineWidth) / 2);
                return margin;
            }
            var border = [
                ground[0] + (ink[0] - ground[0]) * 0.5,
                ground[1] + (ink[1] - ground[1]) * 0.5,
                ground[2] + (ink[2] - ground[2]) * 0.5,
                1
            ];
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, border, 1));

            var pen = g.newPen(g.PenType.SOLID_COLOR, ink, 1.2);
            var rowYs = [7, 11, 15, 19];
            var lineWidths = getJustifyLineWidths(iconType, 15, 10);
            for (var i = 0; i < rowYs.length; i++) {
                var lineWidth = lineWidths[i];
                var lineStartX = getJustifyLineX(iconType, w, lineWidth);
                g.newPath();
                g.moveTo(lineStartX, rowYs[i]);
                g.lineTo(lineStartX + lineWidth, rowYs[i]);
                g.strokePath(pen);
            }
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "左右反転",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawFlipIcon(g, "flipHorizontal", w, h);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "上下反転",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawFlipIcon(g, "flipVertical", w, h);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "回転（反時計回り）",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawRotateIcon(g, w, h, false);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "回転（時計回り）",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawRotateIcon(g, w, h, true);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "上へ移動",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawArrow(g, "up", w, h, iconColor);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "左へ移動",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawArrow(g, "left", w, h, iconColor);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "右へ移動",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawArrow(g, "right", w, h, iconColor);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "下へ移動",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawArrow(g, "down", w, h, iconColor);
        }
    });

    /* ==== QuickTransformPalette (jsx/transform/QuickTransformPalette.jsx) ==== */
    ICON_CATALOG.push({
        script: "QuickTransformPalette",
        name: "同じ位置に複製",
        size: [30, 30],
        draw: function (g, w, h, ink, ground) {
            var ICON_LINE_WIDTH = 1;
            var iconColor = ink;
            function drawButtonBase(graphics, width, height, backgroundColor) {
                /* 枠線は UI の明暗で色が変わるため、地とインクの中間色に置き換え */
                var border = [
                    ground[0] + (ink[0] - ground[0]) * 0.35,
                    ground[1] + (ink[1] - ground[1]) * 0.35,
                    ground[2] + (ink[2] - ground[2]) * 0.35,
                    1
                ];
                graphics.newPath();
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                graphics.newPath();
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, border, 1));
            }
            function squarePath(graphics, x, y, size) {
                graphics.newPath();
                graphics.moveTo(x, y);
                graphics.lineTo(x + size, y);
                graphics.lineTo(x + size, y + size);
                graphics.lineTo(x, y + size);
                graphics.closePath();
            }
            function drawTriangle(graphics, points, color, fill) {
                graphics.newPath();
                graphics.moveTo(points[0][0], points[0][1]);
                graphics.lineTo(points[1][0], points[1][1]);
                graphics.lineTo(points[2][0], points[2][1]);
                graphics.closePath();
                if (fill) {
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
                } else {
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
                }
            }
            function drawDottedLine(graphics, x1, y1, x2, y2, color) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var isHorizontal = (y1 === y2);
                var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
                var dashStep = 3;
                for (var pos = 0; pos < totalLength; pos += dashStep) {
                    graphics.newPath();
                    if (isHorizontal) {
                        graphics.moveTo(x1 + pos, y1);
                        graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
                    } else {
                        graphics.moveTo(x1, y1 + pos);
                        graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
                    }
                    graphics.strokePath(pen);
                }
            }
            function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));
                graphics.newPath();
                for (var i = 0; i <= segments; i++) {
                    var rad = (startDeg + (endDeg - startDeg) * (i / segments)) * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    if (i === 0) { graphics.moveTo(x, y); } else { graphics.lineTo(x, y); }
                }
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign) {
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                var stepDeg = 13;
                var dashHalf = 0.9;
                for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
                    var rad = deg * Math.PI / 180;
                    var x = centerX + mirrorSign * radius * Math.cos(rad);
                    var y = centerY + radius * Math.sin(rad);
                    var tangentX = mirrorSign * Math.sin(rad);
                    var tangentY = -Math.cos(rad);
                    graphics.newPath();
                    graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
                    graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
                    graphics.strokePath(pen);
                }
            }
            function transformArrowPoint(directionKey, x, y) {
                if (directionKey === 'left') { return [-x, y]; }
                if (directionKey === 'down') { return [-y, x]; }
                if (directionKey === 'up')   { return [y, -x]; }
                return [x, y];
            }
            function drawArrow(graphics, directionKey, width, height, color) {
                var iconSize = Math.min(width, height);
                var tip = iconSize * 0.32;
                var shaft = iconSize * 0.11;
                var headHalf = iconSize * 0.27;
                var headBase = tip - iconSize * 0.34;
                var basePoints = [
                    [-tip, -shaft], [headBase, -shaft], [headBase, -headHalf],
                    [tip, 0],
                    [headBase, headHalf], [headBase, shaft], [-tip, shaft]
                ];
                var centerX = width / 2, centerY = height / 2;
                graphics.newPath();
                for (var i = 0; i < basePoints.length; i++) {
                    var point = transformArrowPoint(directionKey, basePoints[i][0], basePoints[i][1]);
                    if (i === 0) { graphics.moveTo(centerX + point[0], centerY + point[1]); }
                    else { graphics.lineTo(centerX + point[0], centerY + point[1]); }
                }
                graphics.closePath();
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH));
            }
            function drawDuplicateGlyph(graphics, width, height, color, backgroundColor) {
                var squareSize = Math.min(width, height) * 0.40;
                var shift = squareSize * 0.36;
                var pairSize = squareSize + shift;
                var left = Math.round((width - pairSize) / 2);
                var top = Math.round((height - pairSize) / 2);
                var backX = left + shift, backY = top;
                var frontX = left, frontY = top + shift;
                var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, ICON_LINE_WIDTH);
                squarePath(graphics, backX, backY, squareSize);
                graphics.strokePath(pen);
                if (backgroundColor) {
                    squarePath(graphics, frontX, frontY, squareSize);
                    graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));
                }
                squarePath(graphics, frontX, frontY, squareSize);
                graphics.strokePath(pen);
            }
            function drawFlipIcon(graphics, iconType, width, height) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2;
                if (iconType === "flipVertical") {
                    drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
                    drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
                    drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
                } else {
                    drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
                    drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
                    drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
                }
            }
            function drawRotateIcon(graphics, width, height, mirror) {
                var color = iconColor;
                var centerX = width / 2;
                var centerY = height / 2 + 1;
                var radius = 7.5;
                var mirrorSign = mirror ? -1 : 1;
                var headDeg = 232;
                strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign);
                drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign);
                var headRad = headDeg * Math.PI / 180;
                var headX = centerX + radius * Math.cos(headRad);
                var headY = centerY + radius * Math.sin(headRad);
                var tangentX = Math.sin(headRad);
                var tangentY = -Math.cos(headRad);
                var perpX = -tangentY;
                var perpY = tangentX;
                var tipForward = 4;
                var tipBack = 2;
                var tipHalfWidth = 4.5;
                var arrowPoints = [
                    [headX + tangentX * tipForward, headY + tangentY * tipForward],
                    [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
                    [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
                ];
                if (mirror) {
                    for (var i = 0; i < arrowPoints.length; i++) {
                        arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
                    }
                }
                drawTriangle(graphics, arrowPoints, color, true);
            }

            drawButtonBase(g, w, h, ground);
            drawDuplicateGlyph(g, w, h, iconColor, ground);
        }
    });

    /* ==== StepperButtons（部品） (jsx/_templates/StepperButtons.jsx) ==== */
    /* 地は半透明の重ね色（ダイアログ地に追従）のため ground で塗り、枠はインクを 10% に落として描く */
    ICON_CATALOG.push({
        script: "StepperButtons（部品）",
        name: "上へ（∧）",
        size: [20, 11],
        draw: function (g, w, h, ink, ground) {
            var isUp = true;
            var STEPPER_CORNER_RADIUS = 2;
            var fillColor = ground;
            var frameColor = [ink[0], ink[1], ink[2], 0.10];
            var chevronColor = ink;
            function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUpper, color) {
                var frameLeft = 0.5;
                var frameRight = boxWidth - 0.5;
                var outerY = isUpper ? 0.5 : boxHeight - 0.5;
                var seamY = isUpper ? boxHeight : 0;
                var towardSeam = isUpper ? 1 : -1;
                var radius = STEPPER_CORNER_RADIUS;
                var arcSteps = 4;
                var angle, k;
                boxGraphics.newPath();
                boxGraphics.moveTo(frameLeft, seamY);
                for (k = 0; k <= arcSteps; k++) {
                    angle = (Math.PI / 2) * k / arcSteps;
                    boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
                }
                for (k = 0; k <= arcSteps; k++) {
                    angle = (Math.PI / 2) * k / arcSteps;
                    boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
                }
                boxGraphics.lineTo(frameRight, seamY);
                boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, color, 1));
            }
            function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUpper, color) {
                var centerX = boxWidth / 2;
                var centerY = isUpper ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5;
                var halfWidth = 3.6;
                var tipOffsetY = isUpper ? -1.8 : 1.8;
                boxGraphics.newPath();
                boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
                boxGraphics.lineTo(centerX, centerY + tipOffsetY);
                boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
                boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, color, 1.2));
            }
            g.newPath();
            g.rectPath(1, isUp ? 1 : 0, w - 2, h - 1);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, fillColor));
            drawStepperFrame(g, w, h, isUp, frameColor);
            drawStepperChevron(g, w, h, isUp, chevronColor);
        }
    });

    /* ==== StepperButtons（部品） (jsx/_templates/StepperButtons.jsx) ==== */
    /* 地は半透明の重ね色（ダイアログ地に追従）のため ground で塗り、枠はインクを 10% に落として描く */
    ICON_CATALOG.push({
        script: "StepperButtons（部品）",
        name: "下へ（∨）",
        size: [20, 11],
        draw: function (g, w, h, ink, ground) {
            var isUp = false;
            var STEPPER_CORNER_RADIUS = 2;
            var fillColor = ground;
            var frameColor = [ink[0], ink[1], ink[2], 0.10];
            var chevronColor = ink;
            function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUpper, color) {
                var frameLeft = 0.5;
                var frameRight = boxWidth - 0.5;
                var outerY = isUpper ? 0.5 : boxHeight - 0.5;
                var seamY = isUpper ? boxHeight : 0;
                var towardSeam = isUpper ? 1 : -1;
                var radius = STEPPER_CORNER_RADIUS;
                var arcSteps = 4;
                var angle, k;
                boxGraphics.newPath();
                boxGraphics.moveTo(frameLeft, seamY);
                for (k = 0; k <= arcSteps; k++) {
                    angle = (Math.PI / 2) * k / arcSteps;
                    boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
                }
                for (k = 0; k <= arcSteps; k++) {
                    angle = (Math.PI / 2) * k / arcSteps;
                    boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
                }
                boxGraphics.lineTo(frameRight, seamY);
                boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, color, 1));
            }
            function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUpper, color) {
                var centerX = boxWidth / 2;
                var centerY = isUpper ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5;
                var halfWidth = 3.6;
                var tipOffsetY = isUpper ? -1.8 : 1.8;
                boxGraphics.newPath();
                boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
                boxGraphics.lineTo(centerX, centerY + tipOffsetY);
                boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
                boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, color, 1.2));
            }
            g.newPath();
            g.rectPath(1, isUp ? 1 : 0, w - 2, h - 1);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, fillColor));
            drawStepperFrame(g, w, h, isUp, frameColor);
            drawStepperChevron(g, w, h, isUp, chevronColor);
        }
    });

    /* ==== LinkToggle（部品） (jsx/_templates/LinkToggle.jsx) ==== */
    ICON_CATALOG.push({
        script: "LinkToggle（部品）",
        name: "連動（リンク）",
        size: [22, 22],
        draw: function (g, w, h, ink, ground) {
            var LINK_ICON_STROKE        = 1.5;
            var LINK_CUT_DIRECTION      = [1, 0];
            var LINK_HOOK_CUT_DIRECTION = [0, 1];
            var LINK_STRAND_COUNT       = 4;
            var LINK_SLASH_CLEARANCE    = 2.2;
            function buildLinkedChainStrokes() {
                var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
                    .concat([[14.5, 11.2]])
                    .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
                var ringStart = upperRing[0];
                var ringEnd = upperRing[upperRing.length - 1];
                var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
                var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
                var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);
                var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
                    var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
                    trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
                    return [trimmed];
                });
                var strokes = [];
                for (var i = 0; i < upperStrands.length; i++) {
                    strokes.push(upperStrands[i]);
                    strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
                }
                return strokes;
            }
            function buildStrandStrokes(centerline, clipStrand) {
                var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
                var strokes = [];
                for (var k = 0; k < LINK_STRAND_COUNT; k++) {
                    var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
                    var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
                    for (var j = 0; j < strandPieces.length; j++) {
                        if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
                    }
                }
                return strokes;
            }
            function extendPolylineEnds(points, length) {
                function extendBeyond(from, to) {
                    var dx = to[0] - from[0];
                    var dy = to[1] - from[1];
                    var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
                    return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
                }
                var lastIndex = points.length - 1;
                return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
            }
            function sideOfLine(point, linePoint, direction) {
                return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
            }
            function trimPolylineTail(points, cutPoint, direction, outsideSign) {
                var lastIndex = points.length - 1;
                var searchLimit = Math.max(0, lastIndex - 12);
                var index = lastIndex;
                while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
                if (index === lastIndex) return points.slice(0);
                var inside = points[index];
                var outside = points[index + 1];
                var insideSide = sideOfLine(inside, cutPoint, direction);
                var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
                return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
            }
            function buildUnlinkedChainStrokes() {
                var slashStart = [3.5, 3.5];
                var slashEnd = [18.5, 18.5];
                var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
                var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];
                function clipAroundSlash(strandPoints) {
                    return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
                }
                var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
                strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
                return strokes;
            }
            function densifyPoints(points) {
                var densePoints = [points[0]];
                for (var i = 1; i < points.length; i++) {
                    var from = points[i - 1];
                    var to = points[i];
                    var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
                    for (var j = 1; j <= steps; j++) {
                        densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
                    }
                }
                return densePoints;
            }
            function offsetPolyline(points, offset) {
                var shifted = [];
                for (var i = 0; i < points.length; i++) {
                    var before = points[Math.max(0, i - 1)];
                    var after = points[Math.min(points.length - 1, i + 1)];
                    var tangentX = after[0] - before[0];
                    var tangentY = after[1] - before[1];
                    var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
                    shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
                }
                return shifted;
            }
            function clipOutsideBand(points, lineStart, lineEnd, clearance) {
                var directionX = lineEnd[0] - lineStart[0];
                var directionY = lineEnd[1] - lineStart[1];
                var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);
                function signedDistance(point) {
                    return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
                }
                function interpolateAt(from, to, fromDistance, toDistance, boundary) {
                    var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
                    return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
                }
                var pieces = [];
                var currentPiece = [];
                for (var i = 0; i < points.length; i++) {
                    var distance = signedDistance(points[i]);
                    var isOutside = Math.abs(distance) >= clearance;
                    if (i > 0) {
                        var previousDistance = signedDistance(points[i - 1]);
                        var wasOutside = Math.abs(previousDistance) >= clearance;
                        if (wasOutside && !isOutside) {
                            currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                            if (currentPiece.length > 1) pieces.push(currentPiece);
                            currentPiece = [];
                        } else if (!wasOutside && isOutside) {
                            currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                        }
                    }
                    if (isOutside) currentPiece.push(points[i]);
                }
                if (currentPiece.length > 1) pieces.push(currentPiece);
                return pieces;
            }
            function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
                var arcSteps = 12;
                var arcPoints = [];
                for (var i = 0; i <= arcSteps; i++) {
                    var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
                    arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
                }
                return arcPoints;
            }
            function rotatePointsHalfTurn(points) {
                var rotated = [];
                for (var i = 0; i < points.length; i++) {
                    rotated.push([22 - points[i][0], 22 - points[i][1]]);
                }
                return rotated;
            }
            function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
                var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
                var offsetX = (iconWidth - 22 * iconScale) / 2;
                var offsetY = (iconHeight - 22 * iconScale) / 2;
                var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
                for (var i = 0; i < strokes.length; i++) {
                    var strokePoints = strokes[i].points;
                    iconGraphics.newPath();
                    for (var j = 0; j < strokePoints.length; j++) {
                        var pointX = offsetX + strokePoints[j][0] * iconScale;
                        var pointY = offsetY + strokePoints[j][1] * iconScale;
                        if (j === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
                }
            }

            /* 連動中は押し込んだボタンのように地と枠を描く（元は UI の明暗に追従する半透明の重ね色。インクを薄めて代用） */
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, [ink[0], ink[1], ink[2], 0.12]));
            g.newPath();
            g.rectPath(0.5, 0.5, w - 1, h - 1);
            g.strokePath(g.newPen(g.PenType.SOLID_COLOR, [ink[0], ink[1], ink[2], 0.07], 1));
            drawLinkIcon(g, w, h, true, ink);
        }
    });

    /* ==== LinkToggle（部品） (jsx/_templates/LinkToggle.jsx) ==== */
    ICON_CATALOG.push({
        script: "LinkToggle（部品）",
        name: "連動しない（リンク解除）",
        size: [22, 22],
        draw: function (g, w, h, ink, ground) {
            var LINK_ICON_STROKE        = 1.5;
            var LINK_CUT_DIRECTION      = [1, 0];
            var LINK_HOOK_CUT_DIRECTION = [0, 1];
            var LINK_STRAND_COUNT       = 4;
            var LINK_SLASH_CLEARANCE    = 2.2;
            function buildLinkedChainStrokes() {
                var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
                    .concat([[14.5, 11.2]])
                    .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
                var ringStart = upperRing[0];
                var ringEnd = upperRing[upperRing.length - 1];
                var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
                var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
                var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);
                var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
                    var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
                    trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
                    return [trimmed];
                });
                var strokes = [];
                for (var i = 0; i < upperStrands.length; i++) {
                    strokes.push(upperStrands[i]);
                    strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
                }
                return strokes;
            }
            function buildStrandStrokes(centerline, clipStrand) {
                var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
                var strokes = [];
                for (var k = 0; k < LINK_STRAND_COUNT; k++) {
                    var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
                    var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
                    for (var j = 0; j < strandPieces.length; j++) {
                        if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
                    }
                }
                return strokes;
            }
            function extendPolylineEnds(points, length) {
                function extendBeyond(from, to) {
                    var dx = to[0] - from[0];
                    var dy = to[1] - from[1];
                    var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
                    return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
                }
                var lastIndex = points.length - 1;
                return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
            }
            function sideOfLine(point, linePoint, direction) {
                return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
            }
            function trimPolylineTail(points, cutPoint, direction, outsideSign) {
                var lastIndex = points.length - 1;
                var searchLimit = Math.max(0, lastIndex - 12);
                var index = lastIndex;
                while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
                if (index === lastIndex) return points.slice(0);
                var inside = points[index];
                var outside = points[index + 1];
                var insideSide = sideOfLine(inside, cutPoint, direction);
                var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
                return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
            }
            function buildUnlinkedChainStrokes() {
                var slashStart = [3.5, 3.5];
                var slashEnd = [18.5, 18.5];
                var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
                var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];
                function clipAroundSlash(strandPoints) {
                    return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
                }
                var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
                strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
                return strokes;
            }
            function densifyPoints(points) {
                var densePoints = [points[0]];
                for (var i = 1; i < points.length; i++) {
                    var from = points[i - 1];
                    var to = points[i];
                    var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
                    for (var j = 1; j <= steps; j++) {
                        densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
                    }
                }
                return densePoints;
            }
            function offsetPolyline(points, offset) {
                var shifted = [];
                for (var i = 0; i < points.length; i++) {
                    var before = points[Math.max(0, i - 1)];
                    var after = points[Math.min(points.length - 1, i + 1)];
                    var tangentX = after[0] - before[0];
                    var tangentY = after[1] - before[1];
                    var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
                    shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
                }
                return shifted;
            }
            function clipOutsideBand(points, lineStart, lineEnd, clearance) {
                var directionX = lineEnd[0] - lineStart[0];
                var directionY = lineEnd[1] - lineStart[1];
                var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);
                function signedDistance(point) {
                    return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
                }
                function interpolateAt(from, to, fromDistance, toDistance, boundary) {
                    var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
                    return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
                }
                var pieces = [];
                var currentPiece = [];
                for (var i = 0; i < points.length; i++) {
                    var distance = signedDistance(points[i]);
                    var isOutside = Math.abs(distance) >= clearance;
                    if (i > 0) {
                        var previousDistance = signedDistance(points[i - 1]);
                        var wasOutside = Math.abs(previousDistance) >= clearance;
                        if (wasOutside && !isOutside) {
                            currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                            if (currentPiece.length > 1) pieces.push(currentPiece);
                            currentPiece = [];
                        } else if (!wasOutside && isOutside) {
                            currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                        }
                    }
                    if (isOutside) currentPiece.push(points[i]);
                }
                if (currentPiece.length > 1) pieces.push(currentPiece);
                return pieces;
            }
            function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
                var arcSteps = 12;
                var arcPoints = [];
                for (var i = 0; i <= arcSteps; i++) {
                    var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
                    arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
                }
                return arcPoints;
            }
            function rotatePointsHalfTurn(points) {
                var rotated = [];
                for (var i = 0; i < points.length; i++) {
                    rotated.push([22 - points[i][0], 22 - points[i][1]]);
                }
                return rotated;
            }
            function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
                var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
                var offsetX = (iconWidth - 22 * iconScale) / 2;
                var offsetY = (iconHeight - 22 * iconScale) / 2;
                var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
                for (var i = 0; i < strokes.length; i++) {
                    var strokePoints = strokes[i].points;
                    iconGraphics.newPath();
                    for (var j = 0; j < strokePoints.length; j++) {
                        var pointX = offsetX + strokePoints[j][0] * iconScale;
                        var pointY = offsetY + strokePoints[j][1] * iconScale;
                        if (j === 0) iconGraphics.moveTo(pointX, pointY);
                        else iconGraphics.lineTo(pointX, pointY);
                    }
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
                }
            }

            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));
            drawLinkIcon(g, w, h, false, ink);
        }
    });

    /* ==== AnchorWidget（部品） (jsx/_templates/AnchorWidget.jsx) ==== */
    /* 中央を選択した状態で描く。ケイ線・枠は元ではインクより少し薄いグレーのため、インクを地へ寄せた色で代用 */
    ICON_CATALOG.push({
        script: "AnchorWidget（部品）",
        name: "基準点（中央）",
        size: [66, 66],
        draw: function (g, w, h, ink, ground) {
            var ANCHOR_WIDGET_CELL_SIZE = 9;
            var ANCHOR_WIDGET_CELL_GAP  = 7.5;
            var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
            var selectedIndex = 4;
            var lineColor = [
                ink[0] + (ground[0] - ink[0]) * 0.2,
                ink[1] + (ground[1] - ink[1]) * 0.2,
                ink[2] + (ground[2] - ink[2]) * 0.2,
                1
            ];
            var fillColor = ink;
            var cellSize = ANCHOR_WIDGET_CELL_SIZE;
            var halfCell = cellSize / 2;
            var i;

            g.newPath();
            g.rectPath(0, 0, w, h);
            g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, ground));

            var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
            var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
            var originX = Math.round((w - gridSize) / 2);
            var originY = Math.round((h - gridSize) / 2);
            var cellPositions = [];
            for (i = 0; i < 9; i++) {
                cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
            }

            var linePen = g.newPen(g.PenType.SOLID_COLOR, lineColor, 1);
            for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
                var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
                var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
                g.newPath();
                if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
                    g.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
                    g.lineTo(cellB[0], cellB[1] + halfCell);
                } else {
                    g.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
                    g.lineTo(cellB[0] + halfCell, cellB[1]);
                }
                g.strokePath(linePen);
            }

            for (i = 0; i < 9; i++) {
                var cellX = cellPositions[i][0];
                var cellY = cellPositions[i][1];
                if (i === selectedIndex) {
                    g.newPath();
                    g.rectPath(cellX, cellY, cellSize, cellSize);
                    g.fillPath(g.newBrush(g.BrushType.SOLID_COLOR, fillColor));
                }
                g.newPath();
                g.rectPath(cellX, cellY, cellSize, cellSize);
                g.strokePath(g.newPen(g.PenType.SOLID_COLOR, lineColor, 1));
            }
        }
    });

    /* ==== IconButtons（部品） (jsx/_templates/IconButtons.jsx) ==== */
    /* 元は地を塗らない（ダイアログの地が透ける）ため、ground は使わない */
    ICON_CATALOG.push({
        script: "IconButtons（部品）",
        name: "保存",
        size: [24, 22],
        draw: function (g, w, h, ink, ground) {
            function fillIconButtonRects(iconGraphics, iconWidth, iconHeight, designSize, designRects, iconColor) {
                var iconScale = Math.min(iconWidth / designSize[0], iconHeight / designSize[1]);
                var originX = (iconWidth - designSize[0] * iconScale) / 2;
                var originY = (iconHeight - designSize[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < designRects.length; i++) {
                    var designRect = designRects[i];
                    if (designRect[2] <= designRect[0] || designRect[3] <= designRect[1]) continue;
                    iconGraphics.rectPath(originX + designRect[0] * iconScale, originY + designRect[1] * iconScale,
                        (designRect[2] - designRect[0]) * iconScale, (designRect[3] - designRect[1]) * iconScale);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
            }

            var designWidth = 1223;
            var trayHole = [78, 688, 230, 840];
            var sliceCount = 16;
            var designRects = [[535, 0, 688, 383]];
            var k;
            function addTrayRect(left, top, right, bottom) {
                if (right <= trayHole[0] || left >= trayHole[2] || bottom <= trayHole[1] || top >= trayHole[3]) {
                    designRects.push([left, top, right, bottom]);
                    return;
                }
                designRects.push([left, top, right, trayHole[1]], [left, trayHole[3], right, bottom],
                    [left, Math.max(top, trayHole[1]), trayHole[0], Math.min(bottom, trayHole[3])],
                    [trayHole[2], Math.max(top, trayHole[1]), right, Math.min(bottom, trayHole[3])]);
            }
            for (k = 0; k < sliceCount; k++) {
                var headHalf = 251 * (1 - (k + 0.5) / sliceCount);
                var headTop = 383 + (703 - 383) * k / sliceCount;
                designRects.push([612 - headHalf, headTop, 612 + headHalf, headTop + (703 - 383) / sliceCount]);
            }
            for (k = 0; k < sliceCount; k++) {
                var notchTop = 612 + (856 - 612) * k / sliceCount;
                var notchBottom = notchTop + (856 - 612) / sliceCount;
                var notchHalf = 229 * (1 - (k + 0.5) / sliceCount);
                addTrayRect(0, notchTop, 612 - notchHalf, notchBottom);
                addTrayRect(612 + notchHalf, notchTop, designWidth, notchBottom);
            }
            addTrayRect(0, 856, designWidth, 993);
            fillIconButtonRects(g, w, h, [designWidth, 993], designRects, ink);
        }
    });

    /* ==== IconButtons（部品） (jsx/_templates/IconButtons.jsx) ==== */
    ICON_CATALOG.push({
        script: "IconButtons（部品）",
        name: "削除（ゴミ箱）",
        size: [24, 22],
        draw: function (g, w, h, ink, ground) {
            function fillIconButtonRects(iconGraphics, iconWidth, iconHeight, designSize, designRects, iconColor) {
                var iconScale = Math.min(iconWidth / designSize[0], iconHeight / designSize[1]);
                var originX = (iconWidth - designSize[0] * iconScale) / 2;
                var originY = (iconHeight - designSize[1] * iconScale) / 2;
                iconGraphics.newPath();
                for (var i = 0; i < designRects.length; i++) {
                    var designRect = designRects[i];
                    if (designRect[2] <= designRect[0] || designRect[3] <= designRect[1]) continue;
                    iconGraphics.rectPath(originX + designRect[0] * iconScale, originY + designRect[1] * iconScale,
                        (designRect[2] - designRect[0]) * iconScale, (designRect[3] - designRect[1]) * iconScale);
                }
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
            }

            fillIconButtonRects(g, w, h, [756, 825], [
                [206, 0, 550, 70], [206, 70, 275, 137], [481, 70, 550, 137],
                [0, 137, 756, 207],
                [69, 207, 138, 825], [618, 207, 688, 825], [138, 756, 618, 825],
                [206, 275, 275, 687], [343, 275, 413, 687], [481, 275, 550, 687]
            ], ink);
        }
    });

    // =========================================
    // 表示 / Display
    // =========================================

    /**
     * 描画先の座標をずらして渡す ScriptUIGraphics の代わりを作る（アイコンを枠の中央に置くため）
     * @param {ScriptUIGraphics} baseGraphics - 本来の描画先
     * @param {number} offsetX - 横のずれ
     * @param {number} offsetY - 縦のずれ
     * @returns {Object} 同じ名前のメソッドを持つオブジェクト
     */
    function createOffsetGraphics(baseGraphics, offsetX, offsetY) {
        return {
            BrushType: baseGraphics.BrushType,
            PenType: baseGraphics.PenType,
            font: baseGraphics.font,
            foregroundColor: baseGraphics.foregroundColor,
            backgroundColor: baseGraphics.backgroundColor,
            newPath: function () { return baseGraphics.newPath(); },
            closePath: function () { return baseGraphics.closePath(); },
            moveTo: function (x, y) { return baseGraphics.moveTo(x + offsetX, y + offsetY); },
            lineTo: function (x, y) { return baseGraphics.lineTo(x + offsetX, y + offsetY); },
            rectPath: function (x, y, width, height) { return baseGraphics.rectPath(x + offsetX, y + offsetY, width, height); },
            ellipsePath: function (x, y, width, height) { return baseGraphics.ellipsePath(x + offsetX, y + offsetY, width, height); },
            fillPath: function (brush, path) { return baseGraphics.fillPath(brush, path); },
            strokePath: function (pen, path) { return baseGraphics.strokePath(pen, path); },
            newBrush: function (brushType, color) { return baseGraphics.newBrush(brushType, color); },
            newPen: function (penType, color, width) { return baseGraphics.newPen(penType, color, width); },
            newFont: function (name, style, size) { return baseGraphics.newFont ? baseGraphics.newFont(name, style, size) : ScriptUI.newFont(name, style, size); },
            measureString: function (text, font, width) { return baseGraphics.measureString(text, font, width); },
            drawString: function (text, pen, x, y, font) { return baseGraphics.drawString(text, pen, x + offsetX, y + offsetY, font); },
            drawImage: function (image, x, y, width, height) { return baseGraphics.drawImage(image, x + offsetX, y + offsetY, width, height); },
            drawOSControl: function () {},
            drawFocusRing: function () {}
        };
    }

    /**
     * アイコンの並びを 1ページ SLOT_COUNT 個に分けて、ページの並びに足す
     * @param {Object[]} catalogPages - 足し先
     * @param {string} pageTitle - リストに出す名前（スクリプト名か「すべて」）
     * @param {Object[]} pageEntries - アイコンの並び
     * @returns {void}
     */
    function pushCatalogPages(catalogPages, pageTitle, pageEntries) {
        var pageCount = Math.ceil(pageEntries.length / SLOT_COUNT);
        for (var page = 0; page < pageCount; page++) {
            catalogPages.push({
                script: pageTitle,
                page: page + 1,
                pages: pageCount,
                total: pageEntries.length,
                entries: pageEntries.slice(page * SLOT_COUNT, (page + 1) * SLOT_COUNT)
            });
        }
    }

    /**
     * 先頭に「すべて」、続けてスクリプトごとにアイコンをまとめ、1ページ SLOT_COUNT 個に分ける
     * @returns {Object[]} { script, page, pages, total, entries } の並び（「すべて」のあとはスクリプト名順）
     */
    function buildCatalogPages() {
        var scriptNames = [];
        var entriesByScript = {};
        for (var i = 0; i < ICON_CATALOG.length; i++) {
            var scriptName = ICON_CATALOG[i].script;
            if (!entriesByScript.hasOwnProperty(scriptName)) {
                entriesByScript[scriptName] = [];
                scriptNames.push(scriptName);
            }
            entriesByScript[scriptName].push(ICON_CATALOG[i]);
        }
        scriptNames.sort();
        /* 「すべて」はスクリプト名順に全アイコンを続けて並べる / All lists every icon, script by script */
        var allEntries = [];
        for (var j = 0; j < scriptNames.length; j++) allEntries = allEntries.concat(entriesByScript[scriptNames[j]]);
        var catalogPages = [];
        pushCatalogPages(catalogPages, getLabel("list.allScripts"), allEntries);
        for (var k = 0; k < scriptNames.length; k++) pushCatalogPages(catalogPages, scriptNames[k], entriesByScript[scriptNames[k]]);
        return catalogPages;
    }

    /**
     * 一覧のリストに出す項目名を作る（例「AdjustPairGap (6)」、複数ページなら「1/2」も付ける）
     * @param {Object} catalogPage - buildCatalogPages() の要素
     * @returns {string} 項目名
     */
    function formatPageName(catalogPage) {
        var pageName = catalogPage.script + " (" + catalogPage.total + ")";
        if (catalogPage.pages > 1) pageName += " " + getLabel("status.page", { page: catalogPage.page, pages: catalogPage.pages });
        return pageName;
    }

    /**
     * アイコンを描く枠を1つ作る（枠の中央にアイコンを、下に名前を出す）
     * @param {Group} parent - 追加先
     * @returns {Group} 枠（.iconEntry に表示するアイコン、.nameText に名前の欄）
     */
    function addIconSlot(parent) {
        var slotGroup = parent.add("group");
        slotGroup.orientation = "column";
        slotGroup.alignChildren = ["center", "top"];
        slotGroup.spacing = 2;
        var iconArea = slotGroup.add("group");
        iconArea.preferredSize = SLOT_ICON_AREA;
        iconArea.minimumSize = SLOT_ICON_AREA;
        iconArea.maximumSize = SLOT_ICON_AREA;
        iconArea.iconEntry = null;
        iconArea.onDraw = function () { drawIconSlot(iconArea); };
        var nameText = slotGroup.add("statictext", undefined, "");
        nameText.preferredSize.width = SLOT_ICON_AREA[0];
        nameText.justify = "center";
        iconArea.nameText = nameText;
        return iconArea;
    }

    /**
     * 枠に割り当てたアイコンを、枠の中央に描く（アイコンの範囲には地の色を敷く）
     * @param {Group} iconArea - addIconSlot() で作った枠
     * @returns {void}
     */
    function drawIconSlot(iconArea) {
        var iconEntry = iconArea.iconEntry;
        if (!iconEntry) return;
        var areaGraphics = iconArea.graphics;
        var iconWidth = iconEntry.size[0];
        var iconHeight = iconEntry.size[1];
        var offsetX = Math.round((SLOT_ICON_AREA[0] - iconWidth) / 2);
        var offsetY = Math.round((SLOT_ICON_AREA[1] - iconHeight) / 2);
        areaGraphics.newPath();
        areaGraphics.rectPath(offsetX, offsetY, iconWidth, iconHeight);
        areaGraphics.fillPath(areaGraphics.newBrush(areaGraphics.BrushType.SOLID_COLOR, CATALOG_GROUND_COLOR));
        try {
            iconEntry.draw(createOffsetGraphics(areaGraphics, offsetX, offsetY), iconWidth, iconHeight, CATALOG_INK_COLOR, CATALOG_GROUND_COLOR);
        } catch (e) {
            /* 抜き出した描画が環境の違いで失敗しても、ほかの枠は描き続ける / keep drawing the other slots if one fails */
            $.writeln(SCRIPT_NAME + ": " + iconEntry.script + " / " + iconEntry.name + ": " + e);
        }
    }

    /**
     * 選んだページのアイコンを枠に割り当て、描き直す
     * @param {Group[]} iconSlots - addIconSlot() で作った枠
     * @param {Object} catalogPage - buildCatalogPages() の要素
     * @returns {void}
     */
    function showCatalogPage(iconSlots, catalogPage) {
        for (var i = 0; i < iconSlots.length; i++) {
            var iconEntry = catalogPage ? (catalogPage.entries[i] || null) : null;
            iconSlots[i].iconEntry = iconEntry;
            iconSlots[i].nameText.text = iconEntry ? iconEntry.name : "";
            iconSlots[i].helpTip = iconEntry ? getLabel("tooltip.iconSlot", {
                script: iconEntry.script, name: iconEntry.name, width: iconEntry.size[0], height: iconEntry.size[1]
            }) : "";
            /* group には notify() が無いため、隠して再表示して描き直させる / groups have no notify(), so hide and show to repaint */
            iconSlots[i].hide();
            iconSlots[i].show();
        }
    }

    /**
     * 一覧のダイアログボックスを表示する（左にスクリプトのリスト、右にアイコンの枠）
     * @returns {void}
     */
    function showCatalogDialog() {
        var catalogPages = buildCatalogPages();
        var catalogDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(catalogDialog);

        var catalogColumns = catalogDialog.add("group");
        catalogColumns.orientation = "row";
        catalogColumns.alignChildren = ["fill", "fill"];
        catalogColumns.spacing = COLUMN_SPACING;

        var pageNames = [];
        for (var i = 0; i < catalogPages.length; i++) pageNames.push(formatPageName(catalogPages[i]));
        var scriptList = catalogColumns.add("listbox", undefined, pageNames);
        scriptList.preferredSize = SCRIPT_LIST_SIZE;
        scriptList.helpTip = getLabel("tooltip.scriptList");

        var iconPanel = catalogColumns.add("panel", undefined, getLabel("panel.icons"));
        setupPanel(iconPanel);
        var iconSlots = [];
        for (var row = 0; row < SLOT_ROWS; row++) {
            var slotRow = iconPanel.add("group");
            setupRow(slotRow, "left", SLOT_SPACING);
            for (var column = 0; column < SLOT_COLUMNS; column++) iconSlots.push(addIconSlot(slotRow));
        }

        var summaryText = catalogDialog.add("statictext", undefined, getLabel("status.summary", {
            scripts: countScripts(), icons: ICON_CATALOG.length
        }));
        summaryText.alignment = ["left", "top"];

        var buttonRow = addButtonRow(catalogDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.close"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        scriptList.onChange = function () {
            showCatalogPage(iconSlots, scriptList.selection ? catalogPages[scriptList.selection.index] : null);
        };
        if (catalogPages.length > 0) scriptList.selection = 0;

        prepareDialogWindow(catalogDialog, SCRIPT_NAME);
        catalogDialog.show();
    }

    /**
     * 一覧に入っているスクリプトの数を返す
     * @returns {number} スクリプトの数
     */
    function countScripts() {
        var seenScripts = {};
        var scriptCount = 0;
        for (var i = 0; i < ICON_CATALOG.length; i++) {
            if (seenScripts.hasOwnProperty(ICON_CATALOG[i].script)) continue;
            seenScripts[ICON_CATALOG[i].script] = true;
            scriptCount++;
        }
        return scriptCount;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    showCatalogDialog();

})();
