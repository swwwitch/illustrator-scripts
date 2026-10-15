#target illustrator
#targetengine "ColorNumberByWidthEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストのうち、全角の英数字と「！？」を赤、半角の英数字と「!?」を青、等幅字形の文字を緑に塗ります。ほかの文字はベースの色（黒）にします。
色はダイアログボックスの色見本をクリックして、カラーピッカーで変えられます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorNumberByWidth.md

### Overview

Colors full-width alphanumerics and "！？" red, half-width alphanumerics and "!?" blue, and characters set to fixed-width glyphs green in the selected text. All other characters get the base color (black).
Click a swatch in the dialog to change the color with the color picker.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorNumberByWidth.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ColorNumberByWidth";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorNumberByWidth.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorNumberByWidth.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* ベースの色（ほかの文字）の初期値 / initial base color for all other characters */
    var BASE_RGB  = [0, 0, 0];
    var BASE_CMYK = [0, 0, 0, 100];

    /* 全角の色の初期値 / initial color for full-width characters */
    var FULL_WIDTH_RGB  = [255, 0, 0];
    var FULL_WIDTH_CMYK = [0, 100, 100, 0];

    /* 半角の色の初期値 / initial color for half-width characters */
    var HALF_WIDTH_RGB  = [0, 153, 255];   /* #0099ff */
    var HALF_WIDTH_CMYK = [100, 40, 0, 0];

    /* 等幅字形の色の初期値 / initial color for fixed-width glyphs */
    var GLYPH_WIDTH_RGB  = [0, 170, 0];
    var GLYPH_WIDTH_CMYK = [80, 0, 100, 0];

    /* 全角：０-９ Ａ-Ｚ ａ-ｚ ！？、半角：0-9 A-Z a-z !? / full-width and half-width alphanumerics, ! and ? */
    var FULL_WIDTH_PATTERN = /[０-９Ａ-Ｚａ-ｚ！？]/;
    var HALF_WIDTH_PATTERN = /[0-9A-Za-z!?]/;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var COLOR_SWATCH_SIZE = [60, 20]; /* 色見本の大きさ / swatch size */

    /**
     * 複数のラベル（statictext）の幅を最長のものへ揃え、指定方向に揃えます。
     * 実際の描画幅（preferredSize.width）を測るため、文言長やロケールに依存しません。
     * @param {StaticText[]} labelControls - 幅を揃える statictext の配列。
     * @param {string} [justify] - 揃え方向 "left" | "center" | "right"（省略時は "right"）。
     * @returns {number} 揃えた後の共通ラベル幅（px）。
     */
    function alignLabelWidths(labelControls, justify) {
        if (!labelControls || !labelControls.length) return 0;
        justify = justify || 'right';

        /* 1パス目：各ラベルの自然幅を測り、最大値を求める / Pass 1: find the widest natural width */
        var maxLabelWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            /* preferredSize を返さない環境では実サイズで代用する / fall back to the laid-out size */
            var naturalWidth = labelControls[i].preferredSize.width || labelControls[i].size.width;
            if (naturalWidth > maxLabelWidth) maxLabelWidth = naturalWidth;
        }

        /* 2パス目：全ラベルを最大幅に固定し、揃え方向を適用 / Pass 2: apply common width and justification */
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxLabelWidth;
            labelControls[j].justify = justify;
        }

        return maxLabelWidth;
    }

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
    var LABELS = {
        dialog: {
            title: { ja: "全角・半角の英数字を塗り分け", en: "Color Alphanumerics by Width" }
        },
        panel: {
            color: { ja: "カラー", en: "Color" }
        },
        fieldLabel: {
            base: { ja: "ベース", en: "Base" },
            fullWidth: { ja: "全角", en: "Full-width" },
            halfWidth: { ja: "半角", en: "Half-width" },
            glyphWidth: { ja: "等幅字形", en: "Fixed-width glyphs" }
        },
        tooltip: {
            base: {
                ja: "全角・半角・等幅字形のどれにも当たらない文字の色。クリックするとカラーピッカーが開きます。",
                en: "Color for characters that are not full-width, half-width or fixed-width glyphs. Click to open the color picker."
            },
            fullWidth: {
                ja: "全角の英数字と「！？」の色。クリックするとカラーピッカーが開きます。",
                en: "Color for full-width alphanumerics and \"！？\". Click to open the color picker."
            },
            halfWidth: {
                ja: "半角の英数字と「!?」の色。クリックするとカラーピッカーが開きます。",
                en: "Color for half-width alphanumerics and \"!?\". Click to open the color picker."
            },
            glyphWidth: {
                ja: "等幅全角・半角・1/3・1/4字形にした文字の色（全角・半角より優先）。クリックするとカラーピッカーが開きます。",
                en: "Color for characters set to full-, half-, third- or quarter-width glyphs (takes priority over full- and half-width). Click to open the color picker."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

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

    // =========================================
    // 色 / Colors
    // =========================================

    /**
     * ドキュメントのカラーモードに合わせて色を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} rgbValues - R, G, B
     * @param {number[]} cmykValues - C, M, Y, K
     * @returns {Color} RGBColor または CMYKColor
     */
    function createDocumentColor(doc, rgbValues, cmykValues) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = cmykValues[0];
            cmykColor.magenta = cmykValues[1];
            cmykColor.yellow = cmykValues[2];
            cmykColor.black = cmykValues[3];
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = rgbValues[0];
        rgbColor.green = rgbValues[1];
        rgbColor.blue = rgbValues[2];
        return rgbColor;
    }

    /**
     * 色を色見本に描くための RGB（0〜1）にする（CMYK・グレーは簡易換算）
     * @param {Color} swatchColor - 色
     * @returns {number[]} [r, g, b]（0〜1）
     */
    function colorToScreenRgb(swatchColor) {
        if (swatchColor.typename === "RGBColor") return [swatchColor.red / 255, swatchColor.green / 255, swatchColor.blue / 255];
        if (swatchColor.typename === "GrayColor") return [1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100];
        if (swatchColor.typename === "CMYKColor") {
            var blackRatio = 1 - swatchColor.black / 100;
            return [
                (1 - swatchColor.cyan / 100) * blackRatio,
                (1 - swatchColor.magenta / 100) * blackRatio,
                (1 - swatchColor.yellow / 100) * blackRatio
            ];
        }
        /* スポットカラーなどは灰色で示す / show other color types in gray */
        return [0.5, 0.5, 0.5];
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * 選択からテキストフレームを集める（グループの中も含む）
     * @param {Array} items - 選択されたオブジェクト
     * @param {TextFrame[]} textFrames - 集めたテキストフレームの格納先
     * @returns {void}
     */
    function collectTextFrames(items, textFrames) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "TextFrame") {
                textFrames.push(item);
            } else if (item.typename === "GroupItem") {
                collectTextFrames(item.pageItems, textFrames);
            }
        }
    }

    /**
     * 選択から塗り分けるテキスト範囲を集める
     * @param {Object} currentSelection - doc.selection
     * @returns {TextRange[]} テキスト範囲
     */
    function collectTargetRanges(currentSelection) {
        /* テキスト編集モードでは selection が TextRange 単体になる / In text-edit mode the selection is a single TextRange */
        if (currentSelection.typename === "TextRange") return [currentSelection];
        var textFrames = [];
        collectTextFrames(currentSelection, textFrames);
        var targetRanges = [];
        for (var i = 0; i < textFrames.length; i++) {
            targetRanges.push(textFrames[i].textRange);
        }
        return targetRanges;
    }

    // =========================================
    // 着色 / Coloring
    // =========================================

    /**
     * 等幅字形（等幅全角・半角・1/3・1/4字形）の AlternateGlyphsForm を集める
     * @returns {Array} AlternateGlyphsForm の値
     */
    function getFixedWidthGlyphForms() {
        return [
            AlternateGlyphsForm.FULLWIDTH,
            AlternateGlyphsForm.HALFWIDTH,
            AlternateGlyphsForm.THIRDWIDTH,
            AlternateGlyphsForm.QUARTERWIDTH
        ];
    }

    /**
     * 文字が等幅字形になっているかを返す
     * @param {TextRange} targetCharacter - 1文字の範囲
     * @param {Array} glyphForms - getFixedWidthGlyphForms() の戻り値
     * @returns {boolean} 等幅字形なら true
     */
    function isFixedWidthGlyph(targetCharacter, glyphForms) {
        var glyphForm;
        /* 字形を読めない文字は等幅字形ではないとみなす / treat characters whose glyph form cannot be read as not fixed-width */
        try {
            glyphForm = targetCharacter.characterAttributes.alternateGlyphs;
        } catch (e) {
            return false;
        }
        for (var i = 0; i < glyphForms.length; i++) {
            if (glyphForm === glyphForms[i]) return true;
        }
        return false;
    }

    /**
     * テキスト範囲を塗り分ける（全体をベースの色にし、等幅字形の文字を緑、英数字と「!?」を全角は赤・半角は青）
     * @param {TextRange} targetRange - 対象のテキスト範囲
     * @param {{baseColor: Color, fullWidthColor: Color, halfWidthColor: Color, glyphWidthColor: Color}} chosenColors - 塗る色
     * @param {Array} glyphForms - getFixedWidthGlyphForms() の戻り値
     * @returns {void}
     */
    function colorCharactersInRange(targetRange, chosenColors, glyphForms) {
        var rangeCharacters = targetRange.characters;
        var characterCount = rangeCharacters.length;
        var rangeContents = targetRange.contents;
        /* contents と文字の番号がずれるときは1文字ずつ読む / read each character when the indices do not line up */
        var useContents = (rangeContents.length === characterCount);

        /* 先に全体をベースの色にし、当たる文字だけ塗り直す / paint the whole range with the base color first, then recolor matches */
        targetRange.characterAttributes.fillColor = chosenColors.baseColor;

        for (var i = 0; i < characterCount; i++) {
            var characterText = useContents ? rangeContents.charAt(i) : rangeCharacters[i].contents;
            var characterColor = null;
            /* 等幅字形は全角・半角より優先する / fixed-width glyphs take priority over full- and half-width */
            if (isFixedWidthGlyph(rangeCharacters[i], glyphForms)) {
                characterColor = chosenColors.glyphWidthColor;
            } else if (FULL_WIDTH_PATTERN.test(characterText)) {
                characterColor = chosenColors.fullWidthColor;
            } else if (HALF_WIDTH_PATTERN.test(characterText)) {
                characterColor = chosenColors.halfWidthColor;
            }
            if (!characterColor) continue;
            rangeCharacters[i].characterAttributes.fillColor = characterColor;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名と、クリックでカラーピッカーを開く色見本の1行を作る
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.fieldLabel / LABELS.tooltip のキー
     * @param {Color} initialColor - 最初の色
     * @returns {{label: StaticText, getColor: Function}} 項目名と、今の色を返す getColor を持つオブジェクト
     */
    function addColorSwatchRow(parentPanel, labelKey, initialColor) {
        var swatchColor = initialColor;
        var swatchRow = parentPanel.add("group");
        setupRow(swatchRow, "left", 8);

        var rowLabel = swatchRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));

        var colorSwatch = swatchRow.add("group");
        colorSwatch.preferredSize = COLOR_SWATCH_SIZE;
        colorSwatch.helpTip = getLabel("tooltip." + labelKey);
        colorSwatch.onDraw = function () {
            var swatchGraphics = colorSwatch.graphics;
            var swatchWidth = colorSwatch.size[0];
            var swatchHeight = colorSwatch.size[1];
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, colorToScreenRgb(swatchColor)));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1));
        };
        /* クリックで標準のカラーピッカーを開く。キャンセルは渡した色が返る
           Clicking opens the standard color picker; Cancel returns the color passed in */
        colorSwatch.addEventListener("click", function () {
            swatchColor = app.showColorPicker(swatchColor);
            colorSwatch.hide(); /* group には notify() が無いので、隠して再表示して描き直す / redraw the group */
            colorSwatch.show();
        });

        return {
            label: rowLabel,
            getColor: function () { return swatchColor; }
        };
    }

    /**
     * 色を選ぶダイアログボックスを表示する
     * @param {Document} doc - 対象ドキュメント
     * @returns {{baseColor: Color, fullWidthColor: Color, halfWidthColor: Color, glyphWidthColor: Color}|null} 選んだ色。キャンセルなら null
     */
    function showColorDialog(doc) {
        var colorDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(colorDialog);

        var colorPanel = colorDialog.add("panel", undefined, getLabel("panel.color"));
        setupPanel(colorPanel, 8);
        var baseRow = addColorSwatchRow(colorPanel, "base", createDocumentColor(doc, BASE_RGB, BASE_CMYK));
        var fullWidthRow = addColorSwatchRow(colorPanel, "fullWidth", createDocumentColor(doc, FULL_WIDTH_RGB, FULL_WIDTH_CMYK));
        var halfWidthRow = addColorSwatchRow(colorPanel, "halfWidth", createDocumentColor(doc, HALF_WIDTH_RGB, HALF_WIDTH_CMYK));
        var glyphWidthRow = addColorSwatchRow(colorPanel, "glyphWidth", createDocumentColor(doc, GLYPH_WIDTH_RGB, GLYPH_WIDTH_CMYK));
        alignLabelWidths([baseRow.label, fullWidthRow.label, halfWidthRow.label, glyphWidthRow.label], "right");

        var buttonRow = addButtonRow(colorDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        prepareDialogWindow(colorDialog, SCRIPT_NAME);
        if (colorDialog.show() !== 1) return null;
        return {
            baseColor: baseRow.getColor(),
            fullWidthColor: fullWidthRow.getColor(),
            halfWidthColor: halfWidthRow.getColor(),
            glyphWidthColor: glyphWidthRow.getColor()
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    var currentSelection = doc.selection;
    if (!currentSelection) return;

    var targetRanges = collectTargetRanges(currentSelection);
    if (targetRanges.length === 0) return;

    var chosenColors = showColorDialog(doc);
    if (!chosenColors) return;

    var glyphForms = getFixedWidthGlyphForms();
    for (var i = 0; i < targetRanges.length; i++) {
        colorCharactersInRange(targetRanges[i], chosenColors, glyphForms);
    }
})();
