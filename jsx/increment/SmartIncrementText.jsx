#target illustrator
#targetengine "SmartIncrementTextEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレーム内の数字・英字・日付・時刻を検出し、値を増分しながら下方向へ複製します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartIncrementText.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n5f25ed17b123

### Overview

Finds the digits, letters, dates or times in the selected text frame and duplicates it downwards,
incrementing the value each time.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartIncrementText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartIncrementText";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartIncrementText.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartIncrementText.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n5f25ed17b123"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_COPY_COUNT = 5;        /* ［複製数］の初期値 / initial number of copies */
    var DEFAULT_STEP = 1;              /* ［増分］の初期値 / initial step */
    var DEFAULT_PITCH_RATIO = 1.5;     /* 既定の送り＝文字サイズ×この倍率 / default pitch = font size * this */
    var DEFAULT_ZERO_PAD = true;       /* ［ゼロ埋め］の初期状態 / zero padding on by default */
    var DEFAULT_MERGE_ON_OK = false;   /* ［1つのテキストに結合］の初期状態 / merging off by default */
    var FALLBACK_FONT_SIZE_PT = 10;    /* 文字サイズを取得できないときの代替値 / fallback font size */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var FIELD_LABEL_WIDTH = 60;        /* 項目名の幅 / width of a row label */
    var NUMBER_FIELD_CHARS = 4;        /* 数値入力欄の文字数 / width of a numeric field */
    var START_FIELD_CHARS = 6;         /* ［開始値］入力欄の文字数 / width of the start field */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "連番複製", en: "Duplicate with Increment" }
        },
        fieldLabel: {
            copyCount: { ja: "複製数", en: "Copies" },
            stepValue: { ja: "増分", en: "Step" },
            gap: { ja: "アキ", en: "Gap" },
            incrementTarget: { ja: "増分対象", en: "Target" }
        },
        checkbox: {
            startOverride: { ja: "開始値", en: "Start value" },
            zeroPad: { ja: "ゼロ埋め", en: "Pad with zeros" },
            mergeOnOK: { ja: "1つのテキストに結合", en: "Merge into one text object" }
        },
        radio: {
            year: { ja: "年", en: "Year" },
            month: { ja: "月", en: "Month" },
            day: { ja: "日", en: "Day" },
            hour: { ja: "時", en: "Hour" },
            minute: { ja: "分", en: "Minute" },
            numberPrefix: { ja: "数字", en: "Num" },
            alphabetPrefix: { ja: "英字", en: "Alpha" }
        },
        tooltip: {
            copyCount: { ja: "作る複製の数です。元のテキストは含みません。", en: "How many copies to create, not counting the original." },
            stepValue: {
                ja: "1つ進むごとに足す数です。負の値で減らせます（0は1として扱います）。",
                en: "Amount added at each step. Negative values count down (0 is treated as 1)."
            },
            gap: {
                ja: "複製どうしのアキです。文字サイズに加算されます（結合するときは行送り＝文字サイズ＋アキ）。",
                en: "Space between copies, added on top of the font size (when merging, leading = font size + gap)."
            },
            incrementTarget: {
                ja: "テキストの中で増やす箇所です。数字や英字が複数あるときに選べます。",
                en: "Which part of the text to increment, when there is more than one number or letter."
            },
            startOverride: { ja: "元のテキストの値ではなく、指定した値から始めます。", en: "Starts from the value you enter instead of the one in the original text." },
            startValue: {
                ja: "元のテキストの値をこの値に置き換え、ここから増分します。英字が対象のときは1文字（A〜Z）で入力します。",
                en: "Replaces the value in the original text and counts on from it. Enter a single letter (A–Z) when a letter is the target."
            },
            zeroPad: {
                ja: "元の桁数に合わせて、頭に0を足します。最終値で桁が増えるときは桁数を広げます。",
                en: "Pads with leading zeros to match the original width, widening it when the last value needs another digit."
            },
            mergeOnOK: {
                ja: "［OK］で確定するとき、複製を改行でつないで元のテキストにまとめます。",
                en: "On OK, joins the copies to the original text with line breaks."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneTextFrame: { ja: "テキストフレームを1つだけ選択してください。", en: "Select exactly one text frame." },
            noTarget: {
                ja: "増分できる箇所が見つかりませんでした。\n英字は1文字（A〜Z）のときだけ増分できます。\n例：A1 の A は対象になりますが、AB1 や Ver1 の英字は対象になりません。",
                en: "Nothing to increment was found.\nLetters are incremented only when they stand alone (A–Z).\nExample: the A in A1 is a target, but the letters in AB1 or Ver1 are not."
            },
            noToken: {
                ja: "選択したテキストに半角数字も英字も含まれていません。\n数字や英字を含むテキスト（例：01、2025/11/21、19:00、A1）を選択してください。",
                en: "The selected text contains no digits or letters.\nSelect text that contains digits or letters (e.g. 01, 2025/11/21, 19:00, A1)."
            }
        }
    };

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ステップボタン / Stepper buttons
    // =========================================

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
    // 実行中の状態 / Run state
    // =========================================
    var sourceTextFrame = null;   /* 複製元のテキストフレーム / source text frame */
    var sourceText = "";          /* 実行前のテキスト（キャンセル時の復帰にも使う） / original contents, also used to restore */
    var literalSegments = [];     /* トークンの間の文字列 / literal text between tokens */
    var sourceTokens = [];        /* 数字・英字のトークン / digit and letter tokens */
    var tokenTypes = [];          /* "alpha1"（英字1文字）/ "num"（数字）/ "other" */
    var patternType = "generic";  /* "date_ymd" / "time_hm" / "generic" */
    var targetTokenIndices = [];  /* 増分対象の候補（トークンの位置） / candidate token positions */
    var targetTokenIndex = 0;     /* 選択中の増分対象 / token position being incremented */
    var targetDigitLength = 0;    /* 増分対象の元の桁数 / original width of the target */
    var weekdayInfo = { has: false }; /* 曜日表記の有無と囲み記号 / weekday notation, if any */
    var sourceFontSizePt = FALLBACK_FONT_SIZE_PT; /* 送り・行送りの基準 / base for the pitch and leading */
    var textUnitInfo = null;      /* 環境設定の文字の単位 / text unit from the preferences */
    var dialogControls = null;    /* buildIncrementDialog() の戻り値 / controls of the dialog */
    var previewItems = [];        /* プレビューで作った複製 / copies made for the preview */
    var previewSuspended = false; /* 復帰中はプレビューを作らない / no preview while restoring */
    var closedWithOK = false;     /* ［OK］で閉じたか / closed with OK */

    // =========================================
    // 元テキストの解析 / Analyze the source text
    // =========================================

    /* 曜日表記（括弧つき） / Weekday in brackets */
    var WEEKDAY_PATTERN = /([（(［\[])[ 　\t]*(日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[ 　\t]*([）)\]］])/;

    /**
     * 英字1文字のトークンかを判定する
     * @param {string} token - 判定するトークン
     * @returns {boolean} A〜Z／a〜z 1文字なら true
     */
    function isAlphaToken(token) {
        return /^[A-Za-z]$/.test(String(token));
    }

    /**
     * 数字・英字のまとまりを抽出し、リテラル部分とトークンに分ける
     * 例: "A1" -> ["A", "1"] / "AB1" -> ["AB", "1"]（AB は英字増分の対象外）
     * @param {string} text - 元のテキスト
     * @returns {{literalSegments: string[], tokens: string[], tokenTypes: string[]}} 分割結果（literalSegments はトークン数＋1個）
     */
    function splitTextIntoTokens(text) {
        var segments = [];
        var tokens = [];
        var types = [];
        var tokenPattern = /[A-Za-z]+|\d+/g;
        var tokenMatch;
        var literalStart = 0;

        while ((tokenMatch = tokenPattern.exec(text)) !== null) {
            segments.push(text.substring(literalStart, tokenMatch.index));
            tokens.push(tokenMatch[0]);
            types.push(isAlphaToken(tokenMatch[0]) ? "alpha1" : (/^\d+$/.test(tokenMatch[0]) ? "num" : "other"));
            literalStart = tokenMatch.index + tokenMatch[0].length;
        }
        segments.push(text.substring(literalStart));
        return { literalSegments: segments, tokens: tokens, tokenTypes: types };
    }

    /**
     * トークン配列からテキストを組み立て直す
     * @param {string[]} tokens - 差し替え後のトークン
     * @returns {string} 組み立てたテキスト
     */
    function rebuildText(tokens) {
        var rebuiltText = literalSegments[0];
        for (var i = 0; i < tokens.length; i++) {
            rebuiltText += String(tokens[i]) + literalSegments[i + 1];
        }
        return rebuiltText;
    }

    /**
     * 増分対象のラジオに表示する名前を返す
     * @param {string} kind - "date_ymd" / "time_hm" / "alpha1" / "num"
     * @param {number} position - 同じ種類の中での位置（1始まり）
     * @returns {string} ラジオのラベル
     */
    function getTargetRadioLabel(kind, position) {
        if (kind === "date_ymd") {
            if (position === 1) return getLabel(LABELS.radio.year);
            if (position === 2) return getLabel(LABELS.radio.month);
            return getLabel(LABELS.radio.day);
        }
        if (kind === "time_hm") {
            return getLabel((position === 1) ? LABELS.radio.hour : LABELS.radio.minute);
        }
        if (kind === "alpha1") return getLabel(LABELS.radio.alphabetPrefix) + position;
        return getLabel(LABELS.radio.numberPrefix) + position;
    }

    /**
     * 元テキストの書式（日付／時刻／一般）を判定し、増分対象の候補を組み立てる
     * @param {string} text - 元のテキスト
     * @param {string[]} types - トークンの種類
     * @returns {{patternType: string, radioLabels: string[], tokenIndices: number[]}} 判定結果
     */
    function detectIncrementTargets(text, types) {
        /* 和文年月日 / Japanese YMD (with optional weekday) */
        var dateJaPattern = /^(\d{4})年(\d{1,2})月(\d{1,2})日(?:[（(［\[]?(?:日|月|火|水|木|金|土)[）)\]］]?)?$/;
        /* スラッシュ区切り / Slash YMD (with optional weekday) */
        var dateSlashPattern = /^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:[ 　\t]*[（(［\[]?(?:日|月|火|水|木|金|土|Sun|Mon|Tue|Wed|Thu|Fri|Sat)[）)\]］]?)?$/;
        /* ドット区切り3要素 / Dot 3 parts */
        var dateDotPattern = /^(\d+)\.(\d+)\.(\d+)$/;
        /* 時刻 / Time */
        var timePattern = /^(\d{1,2}):(\d{2})$/;

        if (dateJaPattern.test(text) || dateSlashPattern.test(text) || dateDotPattern.test(text)) {
            return {
                patternType: "date_ymd",
                radioLabels: [getTargetRadioLabel("date_ymd", 1), getTargetRadioLabel("date_ymd", 2), getTargetRadioLabel("date_ymd", 3)],
                tokenIndices: [0, 1, 2]
            };
        }
        if (timePattern.test(text)) {
            return {
                patternType: "time_hm",
                radioLabels: [getTargetRadioLabel("time_hm", 1), getTargetRadioLabel("time_hm", 2)],
                tokenIndices: [0, 1]
            };
        }

        var radioLabels = [];
        var tokenIndices = [];
        var numberCount = 0;
        var alphabetCount = 0;
        for (var i = 0; i < types.length; i++) {
            /* "AB" や "Ver" のような2文字以上の英字は候補にしない / multi-letter tokens are not candidates */
            if (types[i] === "alpha1") {
                alphabetCount++;
                radioLabels.push(getTargetRadioLabel("alpha1", alphabetCount));
                tokenIndices.push(i);
            } else if (types[i] === "num") {
                numberCount++;
                radioLabels.push(getTargetRadioLabel("num", numberCount));
                tokenIndices.push(i);
            }
        }
        return { patternType: "generic", radioLabels: radioLabels, tokenIndices: tokenIndices };
    }

    /**
     * 既定の増分対象を返す（日付は日、時刻は分、一般は末尾の候補）
     * @param {string} detectedPattern - detectIncrementTargets() の patternType
     * @param {number[]} tokenIndices - 候補のトークンの位置
     * @returns {number} 増分対象のトークンの位置
     */
    function getDefaultTargetIndex(detectedPattern, tokenIndices) {
        if (detectedPattern === "date_ymd") return 2;
        if (detectedPattern === "time_hm") return 1;
        return tokenIndices[tokenIndices.length - 1];
    }

    /**
     * 増分対象を切り替え、元の桁数を控え直す
     * @param {number} tokenIndex - 増分対象のトークンの位置
     * @returns {void}
     */
    function setTargetTokenIndex(tokenIndex) {
        targetTokenIndex = tokenIndex;
        targetDigitLength = String(sourceTokens[tokenIndex]).length;
    }

    /**
     * 増分対象が英字1文字かを返す
     * @returns {boolean} 英字なら true
     */
    function isAlphaTarget() {
        return tokenTypes[targetTokenIndex] === "alpha1";
    }

    /**
     * テキストフレームの文字サイズを返す（読めないときは代替値）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fallbackSizePt - 代替値
     * @returns {number} 文字サイズ（pt）
     */
    function readFontSizePt(textFrame, fallbackSizePt) {
        var fontSizePt = textFrame.textRange.characterAttributes.size;
        return (isNaN(fontSizePt) || fontSizePt <= 0) ? fallbackSizePt : fontSizePt;
    }

    // =========================================
    // 増分の計算 / Increment helpers
    // =========================================

    /**
     * 先頭に0を足して桁を揃える（負の値は符号を残して数字だけを揃える）
     * @param {string|number} value - 対象の値
     * @param {number} length - 揃える桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function zeroPad(value, length) {
        var digits = String(value);
        var sign = "";
        if (digits.charAt(0) === "-") {
            sign = "-";
            digits = digits.substring(1);
        }
        while (digits.length < length) digits = "0" + digits;
        return sign + digits;
    }

    /**
     * 英字1文字を番号に変換する（A=1〜Z=26）
     * @param {string} alphaToken - 英字1文字
     * @returns {number|null} 番号。英字1文字でなければ null
     */
    function alphaToNumber(alphaToken) {
        var upperCased = String(alphaToken).toUpperCase();
        if (!/^[A-Z]$/.test(upperCased)) return null;
        return upperCased.charCodeAt(0) - 64;
    }

    /**
     * 番号を英字1文字に変換する（26を超えたらAに戻る）
     * @param {number} value - 番号
     * @param {boolean} isLowerCase - 小文字で返すか
     * @returns {string} 英字1文字
     */
    function numberToAlpha(value, isLowerCase) {
        var letterNumber = Math.floor(Number(value));
        if (!isFinite(letterNumber) || letterNumber <= 0) letterNumber = 1;
        var offset = ((letterNumber - 1) % 26 + 26) % 26;
        var alphaToken = String.fromCharCode(65 + offset);
        return isLowerCase ? alphaToken.toLowerCase() : alphaToken;
    }

    /**
     * トークンが小文字かを判定する
     * @param {string} token - 判定するトークン
     * @returns {boolean} 小文字なら true
     */
    function isLowerCaseToken(token) {
        return String(token) === String(token).toLowerCase();
    }

    /**
     * テキストに含まれる曜日表記を調べる
     * @param {string} text - 調べるテキスト
     * @returns {{has: boolean, style: string, left: string, right: string}} 曜日の有無と囲み記号
     */
    function detectWeekdayInfo(text) {
        var weekdayMatch = text.match(WEEKDAY_PATTERN);
        if (!weekdayMatch) return { has: false };
        return { has: true, style: (weekdayMatch[2].length === 1) ? "ja" : "en", left: weekdayMatch[1], right: weekdayMatch[3] };
    }

    /**
     * 日付に対応する曜日の文字を返す
     * @param {Date} date - 対象の日付
     * @param {string} style - "ja"（日〜土）または "en"（Sun〜Sat）
     * @returns {string} 曜日の文字
     */
    function weekdayTokenByDate(date, style) {
        var weekdaysJa = ["日", "月", "火", "水", "木", "金", "土"];
        var weekdaysEn = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
        return (style === "en") ? weekdaysEn[date.getDay()] : weekdaysJa[date.getDay()];
    }

    /**
     * テキスト内の曜日表記を日付に合わせて書き換える
     * @param {string} text - 対象のテキスト
     * @param {Date} date - 合わせる日付
     * @returns {string} 書き換えたテキスト
     */
    function applyWeekdayToText(text, date) {
        if (!weekdayInfo.has) return text;
        return text.replace(WEEKDAY_PATTERN, weekdayInfo.left + weekdayTokenByDate(date, weekdayInfo.style) + weekdayInfo.right);
    }

    /**
     * 年月日のトークンから日付を作る
     * @param {string[]} tokens - トークン配列
     * @returns {Date|null} 日付。数値として読めなければ null
     */
    function parseDateFromTokens(tokens) {
        if (tokens.length < 3) return null;
        var year = parseInt(tokens[0], 10);
        var month = parseInt(tokens[1], 10);
        var day = parseInt(tokens[2], 10);
        if (isNaN(year) || isNaN(month) || isNaN(day)) return null;

        /* 2桁年などは現在の世紀で補う / short years are completed with the current century */
        if (String(sourceTokens[0]).length < 4) {
            year = Math.floor((new Date()).getFullYear() / 100) * 100 + year;
        }
        return new Date(year, month - 1, day);
    }

    /**
     * 日付を元テキストの桁数に合わせたトークンにする
     * @param {Date} date - 対象の日付
     * @returns {string[]} 年・月・日のトークン
     */
    function formatDateToTokens(date) {
        var yearLength = String(sourceTokens[0]).length;
        var monthLength = String(sourceTokens[1]).length;
        var dayLength = String(sourceTokens[2]).length;

        var yearToken = String(date.getFullYear());
        if (yearLength < yearToken.length) yearToken = yearToken.slice(yearToken.length - yearLength);
        if (yearLength > yearToken.length) yearToken = zeroPad(yearToken, yearLength);

        var monthToken = String(date.getMonth() + 1);
        var dayToken = String(date.getDate());
        if (monthLength > 1) monthToken = zeroPad(monthToken, monthLength);
        if (dayLength > 1) dayToken = zeroPad(dayToken, dayLength);

        return [yearToken, monthToken, dayToken];
    }

    /**
     * 日付を年／月／日の単位で増減する
     * @param {Date} baseDate - 基準の日付
     * @param {number} unitIndex - 0=年 / 1=月 / 2=日
     * @param {number} step - 増減する量
     * @returns {Date} 増減後の日付
     */
    function addDateByUnit(baseDate, unitIndex, step) {
        var date = new Date(baseDate.getTime());
        if (unitIndex === 0) date.setFullYear(date.getFullYear() + step);
        else if (unitIndex === 1) date.setMonth(date.getMonth() + step);
        else date.setDate(date.getDate() + step);
        return date;
    }

    /**
     * 時分のトークンから時刻を作る
     * @param {string[]} tokens - トークン配列
     * @returns {{hour: number, minute: number}|null} 時刻。数値として読めなければ null
     */
    function parseTimeFromTokens(tokens) {
        if (tokens.length < 2) return null;
        var hour = parseInt(tokens[0], 10);
        var minute = parseInt(tokens[1], 10);
        if (isNaN(hour) || isNaN(minute)) return null;
        return { hour: hour, minute: minute };
    }

    /**
     * 時刻を時／分の単位で増減する（24時間で繰り上がる）
     * @param {{hour: number, minute: number}} baseTime - 基準の時刻
     * @param {number} unitIndex - 0=時 / 1=分
     * @param {number} step - 増減する量
     * @returns {{hour: number, minute: number}} 増減後の時刻
     */
    function addTimeByUnit(baseTime, unitIndex, step) {
        var totalMinutes = baseTime.hour * 60 + baseTime.minute + ((unitIndex === 0) ? step * 60 : step);
        totalMinutes = ((totalMinutes % 1440) + 1440) % 1440;
        return { hour: Math.floor(totalMinutes / 60), minute: totalMinutes % 60 };
    }

    /**
     * 時刻を元テキストの桁数に合わせたトークンにする
     * @param {{hour: number, minute: number}} time - 対象の時刻
     * @returns {string[]} 時・分のトークン
     */
    function formatTimeToTokens(time) {
        var hourLength = String(sourceTokens[0]).length;
        var minuteLength = String(sourceTokens[1]).length;
        var hourToken = String(time.hour);
        var minuteToken = String(time.minute);
        if (hourLength > 1) hourToken = zeroPad(hourToken, hourLength);
        if (minuteLength > 1) minuteToken = zeroPad(minuteToken, minuteLength);
        return [hourToken, minuteToken];
    }

    // =========================================
    // UIの値の取り出し / Reading the dialog values
    // =========================================

    /**
     * ［複製数］の値を取り出す
     * @returns {number} 複製数（数値でなければ NaN）
     */
    function readCopyCount() {
        return parseInt(dialogControls.countInput.text, 10);
    }

    /**
     * ［アキ］の値をポイントで取り出す
     * @returns {number} アキ（pt。数値でなければ NaN）
     */
    function readGapPt() {
        return parseFloat(dialogControls.gapInput.text) * textUnitInfo.pointsPerUnit;
    }

    /**
     * ［増分］の値を取り出す（0は1として扱う）
     * @returns {number} 増分の値
     */
    function readStepValue() {
        var stepValue = parseInt(String(dialogControls.stepInput.text), 10);
        return (isNaN(stepValue) || stepValue === 0) ? 1 : stepValue;
    }

    /**
     * ［開始値］に有効な値が入っているかを判定する
     * @returns {boolean} 使える値なら true
     */
    function isStartValueValid() {
        if (!dialogControls.startOverrideCheckbox.value) return true;
        var rawValue = String(dialogControls.startValueInput.text);
        if (isAlphaTarget()) return isAlphaToken(rawValue);
        return !isNaN(parseInt(rawValue, 10));
    }

    /**
     * ［開始値］を反映した、増分前のトークン配列を作る
     * @returns {string[]} トークン配列
     */
    function buildBaseTokens() {
        var tokens = sourceTokens.slice();
        if (!dialogControls.startOverrideCheckbox.value) return tokens;

        var rawValue = String(dialogControls.startValueInput.text);
        if (isAlphaTarget()) {
            if (isAlphaToken(rawValue)) tokens[targetTokenIndex] = rawValue;
        } else {
            var startNumber = parseInt(rawValue, 10);
            if (!isNaN(startNumber)) tokens[targetTokenIndex] = String(startNumber);
        }
        return tokens;
    }

    /**
     * ［ゼロ埋め］ON時の桁数を返す（最終値が桁上がりするときは広げる）
     * @param {number} baseNumber - 開始の値
     * @param {number} baseLength - 元の桁数
     * @returns {number} 使用する桁数
     */
    function getPadLength(baseNumber, baseLength) {
        var copyCount = readCopyCount();
        if (isNaN(copyCount) || copyCount < 0) return baseLength;
        var lastNumber = baseNumber + (copyCount * readStepValue());
        /* 桁数は符号を除いて数える / the minus sign does not count as a digit */
        var neededLength = Math.max(String(Math.abs(baseNumber)).length, String(Math.abs(lastNumber)).length);
        return (neededLength > baseLength) ? neededLength : baseLength;
    }

    /**
     * ［ゼロ埋め］の設定に従って数値を文字列にする
     * @param {number} value - 変換する値
     * @param {number} baseNumber - 開始の値
     * @param {number} baseLength - 元の桁数
     * @returns {string} 文字列にした値
     */
    function formatNumberToken(value, baseNumber, baseLength) {
        if (!dialogControls.zeroPadCheckbox.value) return String(value);
        return zeroPad(value, getPadLength(baseNumber, baseLength));
    }

    // =========================================
    // テキストの生成 / Building the text
    // =========================================

    /**
     * 対象トークンを数値として増分したテキストを作る
     * @param {string[]} tokens - 元にするトークン配列（書き換える）
     * @param {number} offset - 開始の値からの増分
     * @param {boolean} usePadding - ［ゼロ埋め］の設定を反映するか
     * @returns {string} 組み立てたテキスト
     */
    function buildNumberIncrementedText(tokens, offset, usePadding) {
        var baseNumber = parseInt(tokens[targetTokenIndex], 10);
        if (isNaN(baseNumber)) baseNumber = 0;
        var value = baseNumber + offset;
        tokens[targetTokenIndex] = usePadding ? formatNumberToken(value, baseNumber, targetDigitLength) : String(value);
        return rebuildText(tokens);
    }

    /**
     * 複製1つ分のテキストを作る
     * @param {string[]} baseTokens - 増分前のトークン配列
     * @param {number} offset - 開始の値からの増分（0で複製元と同じ値）
     * @returns {string} 複製に入れるテキスト
     */
    function buildIncrementedText(baseTokens, offset) {
        var tokens = baseTokens.slice();

        if (patternType === "date_ymd") {
            var baseDate = parseDateFromTokens(baseTokens);
            /* 暦として読めない並びは、対象トークンを数値として増分する / fall back to a plain number */
            if (!baseDate) return buildNumberIncrementedText(tokens, offset, false);
            var shiftedDate = addDateByUnit(baseDate, targetTokenIndex, offset);
            var dateTokens = formatDateToTokens(shiftedDate);
            tokens[0] = dateTokens[0];
            tokens[1] = dateTokens[1];
            tokens[2] = dateTokens[2];
            return applyWeekdayToText(rebuildText(tokens), shiftedDate);
        }

        if (patternType === "time_hm") {
            var baseTime = parseTimeFromTokens(baseTokens);
            if (!baseTime) return buildNumberIncrementedText(tokens, offset, false);
            var timeTokens = formatTimeToTokens(addTimeByUnit(baseTime, targetTokenIndex, offset));
            tokens[0] = timeTokens[0];
            tokens[1] = timeTokens[1];
            return rebuildText(tokens);
        }

        if (isAlphaTarget()) {
            var baseAlphaToken = String(baseTokens[targetTokenIndex]);
            var alphaNumber = alphaToNumber(baseAlphaToken);
            if (alphaNumber == null) alphaNumber = 1;
            tokens[targetTokenIndex] = numberToAlpha(alphaNumber + offset, isLowerCaseToken(baseAlphaToken));
            return rebuildText(tokens);
        }

        return buildNumberIncrementedText(tokens, offset, true);
    }

    /**
     * 現在のUI設定を反映した複製元のテキストを作る
     * @returns {string|null} 反映後のテキスト。［開始値］が不正なときは null
     */
    function buildSourceContents() {
        if (!isStartValueValid()) return null;

        /* 数字が対象で、ゼロ埋めも開始値の指定もないときは元の表記のまま
           Keep the original spelling when nothing reformats the number */
        if (patternType === "generic" && !isAlphaTarget() &&
            !dialogControls.zeroPadCheckbox.value && !dialogControls.startOverrideCheckbox.value) {
            return sourceText;
        }
        return buildIncrementedText(buildBaseTokens(), 0);
    }

    // =========================================
    // ドキュメントの更新 / Updating the document
    // =========================================

    /**
     * 複製元のテキストを、位置を保ったまま書き換える（contents の書き換えで位置がずれるため）
     * @param {string} text - 設定するテキスト
     * @returns {void}
     */
    function setSourceContents(text) {
        var savedPosition = sourceTextFrame.position;
        sourceTextFrame.contents = text;
        sourceTextFrame.position = savedPosition;
    }

    /**
     * 現在のUI設定を複製元のテキストに反映する
     * @returns {void}
     */
    function applySettingsToSourceText() {
        var sourceContents = buildSourceContents();
        if (sourceContents === null) return;
        setSourceContents(sourceContents);
    }

    /**
     * 指定した数だけ複製を作り、値を増分する
     * @returns {TextFrame[]} 作成した複製
     */
    function createIncrementedCopies() {
        var copyCount = readCopyCount();
        var gapPt = readGapPt();
        if (isNaN(copyCount) || isNaN(gapPt)) return [];

        /* ［アキ］は文字サイズに足すアキ → 実際の送り＝文字サイズ＋アキ
           The gap field holds the space added on top of the font size */
        var pitchPt = sourceFontSizePt + gapPt;
        var stepValue = readStepValue();
        var baseTokens = buildBaseTokens();
        var createdCopies = [];

        for (var i = 1; i <= copyCount; i++) {
            var textCopy = sourceTextFrame.duplicate();
            textCopy.top = sourceTextFrame.top - (pitchPt * i);
            textCopy.left = sourceTextFrame.left;
            textCopy.contents = buildIncrementedText(baseTokens, i * stepValue);
            createdCopies.push(textCopy);
        }
        return createdCopies;
    }

    /**
     * 複製を改行でつないで複製元の1つのテキストにまとめる
     * @returns {void}
     */
    function mergeCopiesIntoSource() {
        var savedPosition = sourceTextFrame.position;
        var mergedLines = [String(sourceTextFrame.contents)];
        for (var i = 0; i < previewItems.length; i++) {
            mergedLines.push(String(previewItems[i].contents));
        }
        sourceTextFrame.contents = mergedLines.join("\r");

        /* 行送り＝文字サイズ＋［アキ］ / Leading = font size + the gap field */
        var gapPt = readGapPt();
        if (isNaN(gapPt)) gapPt = 0;
        var leadingPt = readFontSizePt(sourceTextFrame, sourceFontSizePt) + gapPt;
        if (leadingPt > 0) {
            sourceTextFrame.textRange.characterAttributes.autoLeading = false;
            sourceTextFrame.textRange.characterAttributes.leading = leadingPt;
        }

        clearPreview();
        sourceTextFrame.position = savedPosition;
    }

    /**
     * 複製元のテキストと位置を実行前の状態に戻す
     * @returns {void}
     */
    function restoreSourceText() {
        previewSuspended = true;
        clearPreview();
        setSourceContents(sourceText);
        previewSuspended = false;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューで作った複製を取り除く
     * @returns {void}
     */
    function clearPreview() {
        for (var i = previewItems.length - 1; i >= 0; i--) {
            previewItems[i].remove();
        }
        previewItems = [];
    }

    /**
     * 現在のUI設定でプレビューを作り直す
     * @returns {void}
     */
    function updatePreview() {
        if (previewSuspended) return;
        applySettingsToSourceText();
        clearPreview();
        /* ［開始値］が不正な間はプレビューを出さず、［OK］も押せなくする
           No preview and no OK while the start value is invalid */
        var startValueValid = isStartValueValid();
        dialogControls.btnOK.enabled = startValueValid;
        if (startValueValid) previewItems = createIncrementedCopies();
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 横並びの行（左寄せ・上下中央）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Group} 作成した行
     */
    function addRow(parentGroup) {
        var newRow = parentGroup.add("group");
        newRow.orientation = "row";
        newRow.alignChildren = ["left", "center"];
        return newRow;
    }

    /**
     * 行の左端に置く項目名を作る
     * @param {Group} targetRow - 追加先の行
     * @param {Object} labelSet - 項目名のラベル定義
     * @returns {StaticText} 作成した項目名
     */
    function addRowLabel(targetRow, labelSet) {
        var rowLabel = targetRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * ∧∨付きの数値入力欄を追加する。↑↓キーも∧∨と同じ処理で増減し、増減のたびにプレビューを更新する
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {number} fieldChars - 入力欄の文字数
     * @param {Object} stepOptions - ∧∨の設定（min / integer）
     * @param {Object} tooltipSet - 入力欄のツールチップ定義
     * @returns {EditText} 作成した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSteppedInput(parentRow, initialText, fieldChars, stepOptions, tooltipSet) {
        /* .text への代入では onChanging が発火しないため、増減のたびにプレビューを更新する
           Setting .text does not fire onChanging, so refresh the preview after each step */
        stepOptions.onStep = function () { updatePreview(); };

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = fieldChars;
        numberInput.helpTip = getLabel(tooltipSet);
        numberInput.stepperGroup = stepperGroup;

        /* ∧∨を無効にしている間は↑↓キーも効かせない（英字が対象の［開始値］など）
           Arrow keys follow the stepper's enabled state, e.g. the start value for a letter target */
        var stepNumber = stepperGroup.stepBy;
        stepperGroup.stepBy = function (direction) {
            if (stepperGroup.enabled) stepNumber(direction);
        };
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 「項目名＋∧∨＋入力欄」の行を作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - 項目名のラベル定義
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - ∧∨の設定（min / integer）
     * @param {Object} tooltipSet - 入力欄のツールチップ定義
     * @returns {{fieldRow: Group, fieldInput: EditText}} 作成した行と入力欄
     */
    function addNumberFieldRow(parentGroup, labelSet, initialText, stepOptions, tooltipSet) {
        var fieldRow = addRow(parentGroup);
        addRowLabel(fieldRow, labelSet);
        var fieldInput = addSteppedInput(fieldRow, initialText, NUMBER_FIELD_CHARS, stepOptions, tooltipSet);
        return { fieldRow: fieldRow, fieldInput: fieldInput };
    }

    /**
     * チェックボックス1つの行を作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - チェックボックスのラベル定義
     * @param {Object} tooltipSet - ツールチップ定義
     * @param {boolean} initialValue - 初期状態
     * @returns {{checkboxRow: Group, checkbox: Checkbox}} 作成した行とチェックボックス
     */
    function addCheckboxRow(parentGroup, labelSet, tooltipSet, initialValue) {
        var checkboxRow = addRow(parentGroup);
        var rowCheckbox = checkboxRow.add("checkbox", undefined, getLabel(labelSet));
        rowCheckbox.helpTip = getLabel(tooltipSet);
        rowCheckbox.value = initialValue;
        return { checkboxRow: checkboxRow, checkbox: rowCheckbox };
    }

    /**
     * 増分対象のラジオの行を作る（候補が複数あるときだけ）
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string[]} radioLabels - ラジオの名前
     * @returns {RadioButton[]} 作成したラジオ（候補が1つなら空）
     */
    function addTargetRadioRow(parentGroup, radioLabels) {
        var targetRadios = [];
        if (radioLabels.length <= 1) return targetRadios;

        var targetRow = addRow(parentGroup);
        addRowLabel(targetRow, LABELS.fieldLabel.incrementTarget);

        var targetRadioGroup = addRow(targetRow);
        for (var i = 0; i < radioLabels.length; i++) {
            var targetRadio = targetRadioGroup.add("radiobutton", undefined, radioLabels[i]);
            targetRadio.value = (targetTokenIndices[i] === targetTokenIndex);
            targetRadio.helpTip = getLabel(LABELS.tooltip.incrementTarget);
            targetRadios.push(targetRadio);
        }
        return targetRadios;
    }

    /**
     * ダイアログを組み立てる（イベントはまだ付けない）
     * @param {string[]} targetRadioLabels - 増分対象のラジオの名前
     * @returns {Object} ダイアログと各コントロール
     */
    function buildIncrementDialog(targetRadioLabels) {
        var incrementDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(incrementDialog);

        var settingsGroup = incrementDialog.add("group");
        settingsGroup.orientation = "column";
        settingsGroup.alignChildren = "left";

        /* 複製数（0以上の整数） / Copies (integer, 0 or more) */
        var countInput = addNumberFieldRow(settingsGroup, LABELS.fieldLabel.copyCount, String(DEFAULT_COPY_COUNT),
            { integer: true, min: 0 }, LABELS.tooltip.copyCount).fieldInput;

        /* 増分（負の値で減らす整数） / Step (integer, negative counts down) */
        var stepInput = addNumberFieldRow(settingsGroup, LABELS.fieldLabel.stepValue, String(DEFAULT_STEP),
            { integer: true }, LABELS.tooltip.stepValue).fieldInput;

        /* アキ（文字サイズに足す量、0以上） / Gap added on top of the font size (0 or more) */
        var defaultGapPt = Math.max(0, sourceFontSizePt * (DEFAULT_PITCH_RATIO - 1));
        var gapRow = addNumberFieldRow(settingsGroup, LABELS.fieldLabel.gap, (defaultGapPt / textUnitInfo.pointsPerUnit).toFixed(1),
            { min: 0 }, LABELS.tooltip.gap);
        gapRow.fieldRow.add("statictext", undefined, textUnitInfo.label);

        /* 増分対象 / Increment target */
        var targetRadios = addTargetRadioRow(settingsGroup, targetRadioLabels);

        /* 開始値（0以上の整数。英字が対象のときは∧∨を使わない） / Start value (integer 0+, no stepper for letters) */
        var startOverrideRow = addCheckboxRow(settingsGroup, LABELS.checkbox.startOverride, LABELS.tooltip.startOverride, false);
        var startValueInput = addSteppedInput(startOverrideRow.checkboxRow, String(sourceTokens[targetTokenIndex]), START_FIELD_CHARS,
            { integer: true, min: 0 }, LABELS.tooltip.startValue);

        /* ゼロ埋め / Zero padding */
        var zeroPadCheckbox = addCheckboxRow(settingsGroup, LABELS.checkbox.zeroPad, LABELS.tooltip.zeroPad, DEFAULT_ZERO_PAD).checkbox;

        /* 確定時にテキストを結合 / Merge text on OK */
        var mergeOnOKCheckbox = addCheckboxRow(settingsGroup, LABELS.checkbox.mergeOnOK, LABELS.tooltip.mergeOnOK, DEFAULT_MERGE_ON_OK).checkbox;

        /* ボタン行 / Button row */
        var buttonRow = addButtonRow(incrementDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            dialog: incrementDialog,
            countInput: countInput,
            stepInput: stepInput,
            gapInput: gapRow.fieldInput,
            targetRadios: targetRadios,
            startOverrideCheckbox: startOverrideRow.checkbox,
            startValueInput: startValueInput,
            zeroPadCheckbox: zeroPadCheckbox,
            mergeOnOKCheckbox: mergeOnOKCheckbox,
            btnOK: btnOK
        };
    }

    /**
     * ［開始値］の入力欄と∧∨の有効／無効をそろえる（∧∨は数字が対象のときだけ）
     * @returns {void}
     */
    function updateStartValueFieldState() {
        var startValueInput = dialogControls.startValueInput;
        var isOverridden = dialogControls.startOverrideCheckbox.value;
        startValueInput.enabled = isOverridden;
        startValueInput.stepperGroup.enabled = isOverridden && !isAlphaTarget();
        redrawSteppersIn(startValueInput.stepperGroup);
    }

    // =========================================
    // イベント / Event handlers
    // =========================================

    /**
     * 増分対象のラジオが切り替わったときの処理
     * @param {number} radioIndex - 選択されたラジオの位置
     * @returns {void}
     */
    function onTargetChanged(radioIndex) {
        setTargetTokenIndex(targetTokenIndices[radioIndex]);
        dialogControls.startValueInput.text = String(sourceTokens[targetTokenIndex]);
        updateStartValueFieldState();
        updatePreview();
    }

    /**
     * ［OK］の処理。プレビューをそのまま結果として確定する
     * @returns {void}
     */
    function commitIncrementedCopies() {
        closedWithOK = true;
        applySettingsToSourceText();

        /* プレビューがそのまま結果になる。空のときだけ作り直す / the preview is the result */
        if (previewItems.length === 0) previewItems = createIncrementedCopies();
        if (dialogControls.mergeOnOKCheckbox.value) mergeCopiesIntoSource();

        /* 複製は確定したので、プレビューの管理対象から外す / keep the copies in the document */
        previewItems = [];
        dialogControls.dialog.close();
    }

    /**
     * ダイアログのイベントを付ける
     * @returns {void}
     */
    function bindDialogEvents() {
        var controls = dialogControls;

        for (var i = 0; i < controls.targetRadios.length; i++) {
            (function (radioIndex) {
                controls.targetRadios[radioIndex].onClick = function () { onTargetChanged(radioIndex); };
            })(i);
        }

        controls.startOverrideCheckbox.onClick = function () {
            updateStartValueFieldState();
            updatePreview();
        };
        controls.zeroPadCheckbox.onClick = updatePreview;

        controls.countInput.onChanging = updatePreview;
        controls.gapInput.onChanging = updatePreview;
        controls.stepInput.onChanging = updatePreview;
        controls.startValueInput.onChanging = updatePreview;

        controls.btnOK.onClick = commitIncrementedCopies;

        /* ［キャンセル］やタイトルバーの×で閉じたときは、プレビューを消して元のテキストに戻す / Also covers the close box */
        controls.dialog.onClose = function () {
            if (!closedWithOK) restoreSourceText();
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したテキストを解析し、ダイアログを開いて増分しながら複製する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / a TextRange selection has no [0] */
        if (selectedItems.typename === "TextRange" || selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") {
            alert(getLabel(LABELS.alert.selectOneTextFrame));
            return;
        }

        sourceTextFrame = selectedItems[0];
        sourceText = sourceTextFrame.contents;

        var splitResult = splitTextIntoTokens(sourceText);
        literalSegments = splitResult.literalSegments;
        sourceTokens = splitResult.tokens;
        tokenTypes = splitResult.tokenTypes;
        if (sourceTokens.length === 0) {
            alert(getLabel(LABELS.alert.noToken));
            return;
        }

        var incrementTargets = detectIncrementTargets(sourceText, tokenTypes);
        if (incrementTargets.tokenIndices.length === 0) {
            alert(getLabel(LABELS.alert.noTarget));
            return;
        }
        patternType = incrementTargets.patternType;
        targetTokenIndices = incrementTargets.tokenIndices;
        setTargetTokenIndex(getDefaultTargetIndex(patternType, targetTokenIndices));

        weekdayInfo = detectWeekdayInfo(sourceText);
        /* 複製の送りと結合時の行送りの基準 / base for the pitch and the merged leading */
        sourceFontSizePt = readFontSizePt(sourceTextFrame, FALLBACK_FONT_SIZE_PT);
        textUnitInfo = getUnitInfo("text/units");

        dialogControls = buildIncrementDialog(incrementTargets.radioLabels);
        bindDialogEvents();
        updateStartValueFieldState();
        updatePreview();
        prepareDialogWindow(dialogControls.dialog, SCRIPT_NAME);
        dialogControls.dialog.show();
    }

    main();

})();
