#target illustrator
#targetengine "DynamicTextGeneratorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの文字幅に合わせたパスを作り、アーチ・円・下向き弓のパス上文字に変換します。
パスに対して文字が占める割合を指定でき、各行の幅を最長行にそろえる「ブロック」も選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DynamicTextGenerator.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nb9e9082df5e5

### Overview

Builds a path sized to the width of the selected text and converts it into text on an arch, a circle, or a downward bow.
The share of the path the text occupies is configurable, and a "block" mode evens every line out to the width of the longest one.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGenerator.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DynamicTextGenerator";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DynamicTextGenerator.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGenerator.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb9e9082df5e5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 円モードで文字が円周に占める割合（カーブ0＝ゆるやかな弧／カーブ100＝ほぼ一周） */
    var CIRCLE_MIN_OCCUPANCY = 0.25;
    var CIRCLE_MAX_OCCUPANCY = 0.95;

    /* 起動時に選ばれるモード。空文字なら未選択で開く
       （modeBlock / modeCircle / modeArch / modeBow） */
    var DEFAULT_MODE = '';

    /* 結果が画面から外れたときに表示位置を合わせるか（ダイアログの初期値） */
    var ZOOM_TO_SELECTION = true;

    /* 結果が可視領域からはみ出したときに合わせる倍率（1で余白なし。小さいほど余白が増える） */
    var VIEW_FIT_RATIO = 0.9;

    /* 起動時のカーニング（keep＝そのまま／metrics＝メトリクス／optical＝オプティカル／mono＝和文等幅） */
    var DEFAULT_AUTO_KERNING = 'keep';

    /* 起動時の効果（EFFECTS の並び順：0＝虹／1＝歪み／2＝3Dリボン／3＝階段／4＝引力） */
    var DEFAULT_EFFECT_INDEX = 0;

    /* 起動時のフィット方法（none＝しない／fontSize＝文字サイズ／tracking＝トラッキング） */
    var DEFAULT_FIT_METHOD = 'fontSize';

    /* カーブスライダーの初期値（0＝直線に近い／100＝最も丸い） */
    var ARC_ROUNDNESS_DEFAULT = 100;

    /* 占有率スライダーの初期値（100＝パスの端まで文字を並べる） */
    var PATH_COVERAGE_DEFAULT = 100;
    /* 占有率スライダーの下限（％） */
    var PATH_COVERAGE_MIN = 30;

    /* カーブ・占有率スライダーを shift キー併用で動かすときの刻み */
    var SLIDER_SHIFT_STEP = 10;

    /* アーチ化の前に改行を削除するか（ダイアログの初期値） */
    var REMOVE_LINE_BREAKS = true;

    /* ブロック：行の分け方の初期値（keep＝そのまま／punctuation＝句読点で改行／count＝行数を指定） */
    var BLOCK_LINE_SPLIT = 'keep';
    /* ブロック：「行数を指定」の初期値 */
    var BLOCK_LINE_COUNT = 3;
    /* ブロック：改行位置とみなす和文の句読点（直後で改行する） */
    var BLOCK_PUNCTUATION = "、。，．！？";
    /* ブロック：改行位置とみなす欧文の句読点（うしろにスペースがあるときだけ改行する） */
    var BLOCK_PUNCTUATION_LATIN = ".,!?";
    /* ブロック：改行位置を送るときに読み飛ばす空白 */
    var BLOCK_BREAK_SPACES = " \t　";
    /* ブロック：句読点に続いていたら前の行に残す閉じ括弧・引用符 */
    var BLOCK_CLOSING_MARKS = "）」』】〉》〕｝］”’)]}\"'";
    /* ブロック：行末に残った句読点を削除するか（ダイアログの初期値） */
    var BLOCK_REMOVE_PUNCTUATION = false;

    /* ブロック：幅をそろえたあとに行送りを自動へ切り替えるか（ダイアログの初期値） */
    var BLOCK_AUTO_LEADING = true;
    /* ブロック：自動行送りの比率（％） */
    var BLOCK_AUTO_LEADING_AMOUNT = 100;
    /* ブロック：変倍率がこの範囲内なら誤差とみなして変更しない */
    var BLOCK_RATIO_EPSILON = 0.001;

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

    var MODE_ICON_SIZE    = [40, 34];  /* モードアイコンボタンの大きさ / size of a mode icon button */
    var MODE_ICON_RADIUS  = 9;         /* アイコンの円・円弧の半径 / radius of the circle and arcs */
    var MODE_ICON_STROKE  = 2;         /* アイコンの線幅 / stroke width of the icons */
    var LABEL_COLUMN_WIDTH = 88;       /* 各行の先頭ラベルの幅 / width of the leading label column */
    var SLIDER_WIDTH = 200;            /* スライダーの幅（全スライダー共通）/ width shared by every slider */
    var COVERAGE_VALUE_WIDTH = 34;     /* 占有率の数値表示の幅 / width of the coverage readout */
    var MODE_GROUP_MARGINS = [15, 5, 15, 5];   /* モード選択の余白 / margins of the mode selector */

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
            title: { ja: "ダイナミックテキスト", en: "Dynamic Text" }
        },
        panel: {
            blockOptions: { ja: "ブロックオプション", en: "Block Options" },
            arcOptions: { ja: "パス上文字オプション", en: "Type on a Path Options" },
            commonOptions: { ja: "共通オプション", en: "Common Options" }
        },
        mode: {
            block: { ja: "ブロック", en: "Block" },
            circle: { ja: "円", en: "Circle" },
            arch: { ja: "アーチ", en: "Arch" },
            bow: { ja: "下向き弓", en: "Bow Down" }
        },
        /* 項目名のコロンは labelText() で付ける / the colon is added by labelText() */
        fieldLabel: {
            lineSplit: { ja: "行の分け方", en: "Line breaks" },
            leading: { ja: "行送り", en: "Leading" },
            roundness: { ja: "カーブ", en: "Curve" },
            coverage: { ja: "占有率", en: "Coverage" },
            fit: { ja: "合わせ方", en: "Fit" },
            effect: { ja: "効果", en: "Effect" },
            autoKerning: { ja: "カーニング", en: "Kerning" },
            tracking: { ja: "トラッキング", en: "Tracking" }
        },
        radio: {
            lineSplitKeep: { ja: "そのまま", en: "Keep" },
            lineSplitPunctuation: { ja: "句読点で改行", en: "At punctuation" },
            lineSplitCount: { ja: "行数を指定", en: "Line count" },
            fitNone: { ja: "しない", en: "None" },
            fitByFontSize: { ja: "文字サイズ", en: "Font size" },
            fitByTracking: { ja: "トラッキング", en: "Tracking" },
            kerningKeep: { ja: "そのまま", en: "Keep" },
            kerningMetrics: { ja: "メトリクス", en: "Metrics" },
            kerningOptical: { ja: "オプティカル", en: "Optical" },
            kerningMono: { ja: "和文等幅", en: "Metrics - Roman Only" }
        },
        checkbox: {
            leadingAuto: { ja: "自動", en: "Auto" },
            removePunctuation: { ja: "行末の句読点を削除", en: "Drop punctuation at line ends" },
            removeLineBreaks: { ja: "改行を削除", en: "Remove line breaks" },
            zoomToSelection: { ja: "結果を画面内に表示", en: "Keep the result in view" }
        },
        /* 効果名は Illustrator の「パス上文字オプション」の表記に合わせる
           Effect names follow Illustrator's own Type on a Path Options dialog */
        dropdown: {
            effectRainbow: { ja: "虹形", en: "Rainbow" },
            effectDistort: { ja: "歪み", en: "Skew" },
            effectRibbon: { ja: "3D リボン", en: "3D Ribbon" },
            effectStep: { ja: "階段状", en: "Stair Step" },
            effectGravity: { ja: "重力", en: "Gravity" }
        },
        unit: {
            percent: { ja: "%", en: "%" },
            line: { ja: "行", en: "lines" }
        },
        button: {
            hiddenChar: { ja: "制御文字の表示", en: "Hidden Characters" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noText: { ja: "対象のテキストが見つかりません。", en: "No target text found." },
            pathFailed: { ja: "パスの生成に失敗しました。", en: "Failed to generate the path." },
            needTwoLines: { ja: "2行以上のテキストを選択してください。", en: "Select text with two or more lines." },
            selectMode: { ja: "モードを選択してください。", en: "Select a mode." }
        },
        tooltip: {
            modeBlock: {
                ja: "パス上文字にはせず、各行の文字サイズを変えて行の幅を最長行にそろえます。パス上文字やエリア内文字は、いったんポイント文字へ変換してからそろえます。",
                en: "Creates no path; scales each line's font size so every line matches the widest one. Type on a path and area type are converted to point text first."
            },
            modeCircle: {
                ja: "閉じた円形のパスに変換し、円周に沿わせます。文字は円の上側中央に配置されます。",
                en: "Converts to a closed circular path and flows the text around it, centered at the top of the circle."
            },
            modeArch: {
                ja: "上に膨らむパスに変換します。",
                en: "Converts the text to a path that bulges upward."
            },
            modeBow: {
                ja: "下に膨らむパスに変換します。",
                en: "Converts the text to a path that bulges downward."
            },
            lineSplit: {
                ja: "幅をそろえる前に、テキストを何行に分けるかを決めます。",
                en: "Decides how the text is split into lines before the widths are fitted."
            },
            lineSplitKeep: {
                ja: "いまの改行のまま、行ごとの幅をそろえます。",
                en: "Keeps the current line breaks and fits each line."
            },
            lineSplitPunctuation: {
                ja: "いまの改行をいったん外し、句読点のうしろで改行し直します。欧文の「.」「,」はうしろにスペースがあるときだけ区切ります。文字サイズもいちばん大きい値にそろえます。",
                en: "Drops the current line breaks and re-breaks after each punctuation mark. Latin marks such as \".\" and \",\" only break when a space follows. Font sizes are levelled to the largest one."
            },
            removePunctuation: {
                ja: "行の終わりに残った句読点を削除します。閉じ括弧や引用符は残します。",
                en: "Deletes the punctuation left at the end of each line. Closing brackets and quotes are kept."
            },
            lineSplitCount: {
                ja: "いまの改行をいったん外し、指定した行数へ均等に分け直します。近くに句読点があればそこで、なければ単語の切れ目で改行し、文字サイズもいちばん大きい値にそろえます。",
                en: "Drops the current line breaks and re-splits the text into the given number of lines, preferring nearby punctuation and falling back to word boundaries. Font sizes are levelled to the largest one."
            },
            leadingAuto: {
                ja: "幅をそろえたあとに、行送りを自動へ切り替えます。",
                en: "Switches leading to auto after the widths are fitted."
            },
            leadingAmount: {
                ja: "自動行送りの比率です。文字サイズに対する行送りの割合を指定します。",
                en: "The auto-leading ratio, given as a percentage of the font size."
            },
            roundness: {
                ja: "0で直線に近く、100でちょうど半円になります。円では、文字に対する円の大きさを決めます。shiftキーを押しながら操作すると10刻みになります。",
                en: "0 is almost straight; 100 makes an exact semicircle. In Circle mode it sets the size of the circle relative to the text. Hold Shift to move in steps of 10."
            },
            coverage: {
                ja: "パス全体のうち、文字が占める割合です。100でパスの端まで、30でパスの中央3割だけに文字が並びます。円では円周に対する割合になり、文字は上側中央に集まります。「合わせ方：トラッキング」のときは文字サイズを保つため、字間を詰めて短くします。shiftキーを押しながら操作すると10刻みになります。",
                en: "How much of the path the text covers. 100 reaches the path ends; 30 keeps the text within the middle third. In Circle mode it is measured against the circumference, so the text gathers at the top. With \"Fit: Tracking\" the font size is preserved, so the spacing is tightened instead. Hold Shift to move in steps of 10."
            },
            fit: {
                ja: "文字の長さをパスの長さに合わせる方法を選びます。",
                en: "Chooses how the length of the text is fitted to the length of the path."
            },
            fitNone: {
                ja: "パスの長さには合わせません。「占有率」を下げたぶんだけ文字サイズが小さくなります。",
                en: "Does not fit to the path. The font size only shrinks by however much Coverage is lowered."
            },
            fitByFontSize: {
                ja: "文字サイズを拡大・縮小してパスの長さに合わせます。比率で変倍するので、文字ごとのサイズ差は保たれます。",
                en: "Scales the font size up or down to the length of the path. Scaling is done by ratio, so mixed character sizes keep their relative differences."
            },
            fitByTracking: {
                ja: "文字サイズを変えず、字間を広げたり詰めたりしてパスの長さに合わせます。",
                en: "Keeps the font size and widens or tightens the letter spacing to the length of the path."
            },
            effect: {
                ja: "Illustratorの「パス上文字オプション」の効果を適用します。虹形＝1文字ずつパスに垂直に立てる標準の効果／歪み＝文字を垂直に保ったまま傾ける／3D リボン＝厚みのあるリボンのように見せる／階段状＝文字を回転させず水平に並べる／重力＝中心へ引き寄せる。",
                en: "Applies Illustrator's Type on a Path effect. Rainbow is the default, standing each character perpendicular to the path; Skew keeps the characters upright and slants them; 3D Ribbon gives them the thickness of a ribbon; Stair Step keeps them horizontal without rotating; Gravity pulls them toward the center."
            },
            removeLineBreaks: {
                ja: "パスに沿わせる前に改行を削除し、1行にまとめます。",
                en: "Removes line breaks and joins the text into a single line before it is flowed along the path."
            },
            kerningKeep: {
                ja: "カーニングの設定には手を触れません。",
                en: "Leaves the kerning settings untouched."
            },
            autoKerning: {
                ja: "文字幅に影響するため、幅をそろえる前に適用します。「メトリクス」を選ぶとプロポーショナルメトリクスもONになります。",
                en: "Applied before the widths are fitted, since it changes the glyph widths. Choosing Metrics also turns proportional metrics on."
            },
            kerningOptical: {
                ja: "字面を見て字間を自動調整します。メトリクス情報を持たないフォントに向きます。",
                en: "Adjusts the spacing automatically from the shapes of the glyphs. Suits fonts that carry no metrics information."
            },
            kerningMono: {
                ja: "欧文だけメトリクスを使い、和文は等幅のままにします。",
                en: "Uses metrics for Roman text only and leaves Japanese text monospaced."
            },
            tracking: {
                ja: "既存のトラッキング値に加算します。「合わせ方：トラッキング」を選んでいるときは自動調整にまかせるため使えません。",
                en: "Adds this value to the existing tracking. Unavailable while \"Fit: Tracking\" is selected, since it is set automatically."
            },
            trackingToggle: {
                ja: "ONでトラッキングを調整できます。OFFにすると0に戻ります。",
                en: "Enable to adjust tracking. Turning it off resets it to 0."
            },
            zoomToSelection: {
                ja: "変換した結果が画面から外れたときに、見える位置へ表示を移します。すでに見えているときは動かしません。",
                en: "Moves the view so the converted result stays visible. The view is left alone when it is already in sight."
            },
            hiddenChar: {
                ja: "制御文字の表示・非表示を切り替えます。改行がどこに入ったかを確かめるときに使います。",
                en: "Toggles the display of hidden characters, so you can check where the line breaks landed."
            },
            ok: {
                ja: "設定を適用して閉じます。",
                en: "Applies the settings and closes."
            },
            cancel: {
                ja: "プレビューを取り消して閉じます。",
                en: "Discards the preview and closes."
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
        }
    };

    // =========================================
    // モード・カーニング・効果の定義 / Mode, kerning and effect definitions
    // =========================================

    /* モード定義（アイコンの表示順）/ Mode definitions in icon display order
       - modeBlock          : パス上文字にせず、各行の幅を最長行にそろえる
       - modeCircle         : 閉じた円形パスに変換し、円周に沿わせる
       - modeArch / modeBow : パスの膨らむ向き（上／下） */
    var MODES = [
        { key: 'modeBlock', labelKey: 'block', tipKey: 'modeBlock' },
        { key: 'modeCircle', labelKey: 'circle', tipKey: 'modeCircle' },
        { key: 'modeArch', labelKey: 'arch', tipKey: 'modeArch' },
        { key: 'modeBow', labelKey: 'bow', tipKey: 'modeBow' }
    ];

    /* 自動カーニングの定義（表示順）。和文等幅は欧文のみメトリクス＝和文は等幅
       Auto-kerning definitions in display order; "mono" is metrics for Roman only */
    var KERNING_METHODS = [
        { key: 'keep', labelKey: 'kerningKeep', tipKey: 'kerningKeep', value: null },
        { key: 'metrics', labelKey: 'kerningMetrics', tipKey: 'autoKerning', value: AutoKernType.AUTO },
        { key: 'optical', labelKey: 'kerningOptical', tipKey: 'kerningOptical', value: AutoKernType.OPTICAL },
        { key: 'mono', labelKey: 'kerningMono', tipKey: 'kerningMono', value: AutoKernType.METRICSROMANONLY }
    ];

    /* 効果の定義（表示順）。command は Illustrator のメニューコマンド名
       Effect definitions in menu order; command is Illustrator's menu command */
    var EFFECTS = [
        { labelKey: 'effectRainbow', command: 'Rainbow' },
        { labelKey: 'effectDistort', command: 'Skew' },
        { labelKey: 'effectRibbon', command: '3D ribbon' },
        { labelKey: 'effectStep', command: 'Stair Step' },
        { labelKey: 'effectGravity', command: 'Gravity' }
    ];

    // =========================================
    // 状態 / State
    // =========================================

    var doc = null;                       /* 対象のドキュメント / active document */
    var baseSelection = [];               /* 開いたときの選択（プレビューの基準）/ selection snapshot for a stable preview */
    var targetTextFrames = [];            /* 変換対象のテキスト / target text frames */
    var selectedPaths = [];               /* テキストと一緒に選ばれていたパス / paths selected together with the text */
    var currentMode = DEFAULT_MODE;       /* 選択中のモード / current mode key */
    var previewTempItems = [];            /* プレビューで作った一時オブジェクト / items created during preview */
    var previewHiddenOriginals = [];      /* プレビュー中に隠した元のオブジェクト / originals hidden during preview */
    var arcRoundnessBeforeCircle = null;  /* 円モードに入る前のカーブ値（戻したときに復帰）/ curve value restored when leaving Circle mode */
    var trackingSyncLock = false;         /* トラッキング欄とスライダーの相互更新の抑止 / guards the edittext <-> slider sync */

    // =========================================
    // ダイアログのコントロール / Dialog controls
    // buildDialog() で作り、ハンドラと変換処理から参照する / created in buildDialog(), read by the handlers and the conversion
    // =========================================

    var dynamicTextDialog = null;
    var modeButtons = [];
    /* ブロックのオプション / Block options */
    var stLineSplit, rbLineSplitKeep, rbLineSplitPunctuation, rbLineSplitCount, etLineCount, stLineCountUnit;
    var cbRemovePunctuation, stLeading, cbAutoLeading, etLeadingAmount, stLeadingUnit;
    /* パス上文字のオプション / Type on a path options */
    var stArcRoundness, slArcRoundness, stPathCoverage, slPathCoverage, stPathCoverageValue;
    var stFit, rbFitNone, rbFitFontSize, rbFitTracking, stEffect, ddEffect, cbRemoveLineBreaks;
    /* 共通のオプション / Common options */
    var stAutoKerning = null;
    var kerningButtons = [];
    var stTracking, cbTracking, etTracking, slTracking;
    /* 表示とボタン / View option and buttons */
    var cbZoomToSelection, btnHiddenChar, btnCancel, btnOK;

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 文字列を数値にする（数値でなければ代わりの値）
     * @param {string} text - 元の文字列
     * @param {number} fallback - 数値でないときの値
     * @returns {number} 数値
     */
    function parseNumber(text, fallback) {
        var parsedNumber = Number(text);
        if (isNaN(parsedNumber)) return fallback;
        return parsedNumber;
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

    /**
     * ∧∨と数値入力欄を隙間0で突き合わせて追加し、↑↓キーも∧∨と同じ処理で増減させる
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {number} characters - 入力欄の文字数
     * @param {Object} stepOptions - ∧∨の設定（min / max / integer / onStep）
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addStepperInput(parentRow, initialText, characters, stepOptions) {
        var stepperFieldGroup = parentRow.add('group');
        stepperFieldGroup.orientation = 'row';
        stepperFieldGroup.alignChildren = ['left', 'center'];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperFieldGroup.add('edittext', undefined, initialText);
        numberInput.characters = characters;
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 入力欄の有効／無効を∧∨ごと切り替える（変わったときだけ∧∨を描き直す）
     * @param {EditText} numberInput - addStepperInput() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        if (numberInput.stepperGroup.enabled === isEnabled) return;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * 値をスライダーの範囲内に収める
     * @param {Slider} slider - 対象のスライダー
     * @param {number} value - 収めたい値
     * @returns {number} 範囲内に収めた値
     */
    function clampToSliderRange(slider, value) {
        if (value < slider.minvalue) return slider.minvalue;
        if (value > slider.maxvalue) return slider.maxvalue;
        return value;
    }

    /**
     * shift キーを押している間だけ、スライダーの値を SLIDER_SHIFT_STEP の倍数へそろえる
     * ScriptUI のスライダーは刻み幅を持たないため、値そのものを丸めて代入する。
     * @param {Slider} slider - 対象のスライダー
     * @returns {void}
     */
    function snapSliderWithShift(slider) {
        if (!ScriptUI.environment.keyboardState.shiftKey) return;

        var snapped = Math.round(slider.value / SLIDER_SHIFT_STEP) * SLIDER_SHIFT_STEP;
        slider.value = clampToSliderRange(slider, snapped);
    }

    /**
     * ↑↓←→キーでスライダーを1ずつ動かす（Shift＝SLIDER_SHIFT_STEP の倍数へスナップ）
     * @param {Slider} slider - 対象のスライダー
     * @param {Function} [onChanged] - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeSliderByArrowKey(slider, onChanged) {
        if (!slider) return;

        slider.addEventListener('keydown', function (event) {
            if (!event) return;

            var goingUp = (event.keyName === 'Up' || event.keyName === 'Right');
            var goingDown = (event.keyName === 'Down' || event.keyName === 'Left');
            if (!goingUp && !goingDown) return;

            var currentValue = Math.round(slider.value);

            if (ScriptUI.environment.keyboardState.shiftKey) {
                /* Shift は刻みの倍数にスナップ / Snap to multiples of the step for Shift */
                currentValue = goingUp
                    ? Math.ceil((currentValue + 1) / SLIDER_SHIFT_STEP) * SLIDER_SHIFT_STEP
                    : Math.floor((currentValue - 1) / SLIDER_SHIFT_STEP) * SLIDER_SHIFT_STEP;
            } else {
                currentValue = goingUp ? currentValue + 1 : currentValue - 1;
            }

            /* スライダーが自分で動かないよう既定の動作を止める / Prevent default arrow key behavior (the slider would move on its own) */
            event.preventDefault();
            slider.value = clampToSliderRange(slider, currentValue);

            if (onChanged) onChanged();
        });
    }

    // =========================================
    // アイコン描画 / Icon drawing
    // =========================================

    /**
     * UI が明るいテーマかどうかを判定する
     * @returns {boolean} 明るいUIなら true
     */
    function isLightUI() {
        /* 取得できないときは暗い側にフォールバック / fall back to dark when unavailable */
        try { return app.preferences.getRealPreference("uiBrightness") > 0.5; } catch (e) { }
        return false;
    }

    /**
     * UIの明暗に応じたアイコン色・背景色・選択色を返す
     * @returns {{icon: number[], bg: number[], selection: number[]}} 描画色（RGBA 0〜1）
     */
    function getIconColors() {
        if (isLightUI()) {
            return { icon: [0.20, 0.20, 0.20, 1], bg: [0.93, 0.93, 0.93, 1], selection: [0.78, 0.78, 0.78, 1] };
        }
        return { icon: [0.88, 0.88, 0.88, 1], bg: [0.27, 0.27, 0.27, 1], selection: [0.45, 0.45, 0.45, 1] };
    }

    /**
     * 無効（ディム）表示用に、アイコン色を背景色へ寄せて薄くする
     * @param {{icon: number[], bg: number[], selection: number[]}} colors - 通常色
     * @returns {{icon: number[], bg: number[], selection: number[]}} ディム色
     */
    function dimIconColors(colors) {
        var towardBackground = 0.6; /* 0=そのまま／1=背景色 / blend factor toward the background */
        function blendTowardBackground(color) {
            return [
                color[0] + (colors.bg[0] - color[0]) * towardBackground,
                color[1] + (colors.bg[1] - color[1]) * towardBackground,
                color[2] + (colors.bg[2] - color[2]) * towardBackground,
                1
            ];
        }
        return { icon: blendTowardBackground(colors.icon), bg: colors.bg, selection: blendTowardBackground(colors.selection) };
    }

    /**
     * 矩形を塗る
     * @param {object} graphics - ScriptUIGraphics
     * @param {object} brush - ブラシ
     * @param {number} x - 左
     * @param {number} y - 上
     * @param {number} rectWidth - 幅
     * @param {number} rectHeight - 高さ
     * @returns {void}
     */
    function fillRect(graphics, brush, x, y, rectWidth, rectHeight) {
        graphics.newPath();
        graphics.rectPath(x, y, rectWidth, rectHeight);
        graphics.fillPath(brush);
    }

    /**
     * 円弧上の座標を求める（0度＝右、角度は画面下向きに増える）
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radius - 半径
     * @param {number} angleDegrees - 角度（度）
     * @returns {number[]} [x, y]
     */
    function arcPoint(centerX, centerY, radius, angleDegrees) {
        var radians = angleDegrees * Math.PI / 180;
        return [centerX + radius * Math.cos(radians), centerY + radius * Math.sin(radians)];
    }

    /**
     * 現在のパスへ円弧を折れ線で追加する（始点へは呼び出し側で moveTo しておく）
     * @param {object} graphics - ScriptUIGraphics
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radius - 半径
     * @param {number} startAngle - 開始角（度）
     * @param {number} endAngle - 終了角（度）
     * @param {number} steps - 折れ線の分割数
     * @returns {void}
     */
    function appendArc(graphics, centerX, centerY, radius, startAngle, endAngle, steps) {
        for (var i = 1; i <= steps; i++) {
            var arcVertex = arcPoint(centerX, centerY, radius, startAngle + (endAngle - startAngle) * (i / steps));
            graphics.lineTo(arcVertex[0], arcVertex[1]);
        }
    }

    /**
     * 円弧を描く
     * @param {object} graphics - ScriptUIGraphics
     * @param {object} pen - ペン
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radius - 半径
     * @param {number} startAngle - 開始角（度）
     * @param {number} endAngle - 終了角（度）
     * @returns {void}
     */
    function strokeArc(graphics, pen, centerX, centerY, radius, startAngle, endAngle) {
        var startPoint = arcPoint(centerX, centerY, radius, startAngle);
        graphics.newPath();
        graphics.moveTo(startPoint[0], startPoint[1]);
        appendArc(graphics, centerX, centerY, radius, startAngle, endAngle, 32);
        graphics.strokePath(pen);
    }

    /**
     * 角丸の矩形を塗る（選択中のアイコンの座布団に使う）
     * @param {object} graphics - ScriptUIGraphics
     * @param {object} brush - ブラシ
     * @param {number} x - 左
     * @param {number} y - 上
     * @param {number} rectWidth - 幅
     * @param {number} rectHeight - 高さ
     * @param {number} radius - 角の半径
     * @returns {void}
     */
    function fillRoundedRect(graphics, brush, x, y, rectWidth, rectHeight, radius) {
        var CORNER_STEPS = 8;
        graphics.newPath();
        graphics.moveTo(x + radius, y);
        graphics.lineTo(x + rectWidth - radius, y);
        appendArc(graphics, x + rectWidth - radius, y + radius, radius, -90, 0, CORNER_STEPS);
        graphics.lineTo(x + rectWidth, y + rectHeight - radius);
        appendArc(graphics, x + rectWidth - radius, y + rectHeight - radius, radius, 0, 90, CORNER_STEPS);
        graphics.lineTo(x + radius, y + rectHeight);
        appendArc(graphics, x + radius, y + rectHeight - radius, radius, 90, 180, CORNER_STEPS);
        graphics.lineTo(x, y + radius);
        appendArc(graphics, x + radius, y + radius, radius, 180, 270, CORNER_STEPS);
        graphics.closePath();
        graphics.fillPath(brush);
    }

    /**
     * ブロックのアイコン（幅のそろった3本の帯）を描く
     * @param {object} graphics - ScriptUIGraphics
     * @param {object} brush - ブラシ
     * @param {number} centerX - アイコンの中心X
     * @param {number} centerY - アイコンの中心Y
     * @returns {void}
     */
    function drawBlockIcon(graphics, brush, centerX, centerY) {
        /* 幅は同じ・高さだけ違う帯で「行の幅がそろった状態」を表す
           Bars share one width and differ only in height: lines fitted to the same width */
        var barWidth = 20;
        var barHeights = [6, 4, 7];
        var barGap = 2;

        var totalHeight = barGap * (barHeights.length - 1);
        for (var i = 0; i < barHeights.length; i++) totalHeight += barHeights[i];

        var y = centerY - totalHeight / 2;
        for (var j = 0; j < barHeights.length; j++) {
            fillRect(graphics, brush, centerX - barWidth / 2, y, barWidth, barHeights[j]);
            y += barHeights[j] + barGap;
        }
    }

    /**
     * 円のアイコン（輪郭だけの正円）を描く
     * @param {object} graphics - ScriptUIGraphics
     * @param {object} pen - 線のペン
     * @param {number} centerX - アイコンの中心X
     * @param {number} centerY - アイコンの中心Y
     * @returns {void}
     */
    function drawCircleIcon(graphics, pen, centerX, centerY) {
        var radius = MODE_ICON_RADIUS;
        graphics.newPath();
        graphics.ellipsePath(centerX - radius, centerY - radius, radius * 2, radius * 2);
        graphics.strokePath(pen);
    }

    /**
     * アーチ・下向き弓のアイコン（半円より少し長い円弧）を描く
     * @param {object} graphics - ScriptUIGraphics
     * @param {object} pen - 線のペン
     * @param {number} centerX - アイコンの中心X
     * @param {number} centerY - アイコンの中心Y
     * @param {boolean} bulgesUp - 上に膨らむなら true
     * @returns {void}
     */
    function drawArcIcon(graphics, pen, centerX, centerY, bulgesUp) {
        var radius = MODE_ICON_RADIUS;
        /* 端が少し垂れた形にするため、半円より20度ずつ長く描く
           The arc runs 20 degrees past a half circle so the ends turn over slightly */
        var overshoot = 20;
        /* 描画範囲の中心をボタンの中心に合わせるための上下のずらし量 */
        var arcOffset = radius * (1 - Math.sin(overshoot * Math.PI / 180)) / 2;

        if (bulgesUp) {
            strokeArc(graphics, pen, centerX, centerY + arcOffset, radius, 180 - overshoot, 360 + overshoot);
        } else {
            strokeArc(graphics, pen, centerX, centerY - arcOffset, radius, -overshoot, 180 + overshoot);
        }
    }

    /**
     * モード選択アイコンを iconbutton に描画する
     * @param {object} control - 描画対象の iconbutton
     * @param {string} iconType - 描画種別（modeBlock / modeCircle / modeArch / modeBow）
     * @param {boolean} selected - 選択中なら true
     * @returns {void}
     */
    function drawModeIcon(control, iconType, selected) {
        var graphics = control.graphics;
        var colors = getIconColors();
        if (!control.enabled) colors = dimIconColors(colors);

        var iconWidth = control.size[0];
        var iconHeight = control.size[1];
        var centerX = iconWidth / 2;
        var centerY = iconHeight / 2;

        var iconBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.icon);
        var iconPen = graphics.newPen(graphics.PenType.SOLID_COLOR, colors.icon, MODE_ICON_STROKE);

        /* 背景を塗ってネイティブ枠を隠す / paint the background to hide the native frame */
        fillRect(graphics, graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.bg), 0, 0, iconWidth, iconHeight);
        /* 選択中は角丸の座布団を敷く / selected: draw a rounded backdrop */
        if (selected) {
            fillRoundedRect(graphics, graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.selection),
                0, 0, iconWidth, iconHeight, 6);
        }

        if (iconType === "modeBlock") {
            drawBlockIcon(graphics, iconBrush, centerX, centerY);
        } else if (iconType === "modeCircle") {
            drawCircleIcon(graphics, iconPen, centerX, centerY);
        } else if (iconType === "modeArch") {
            drawArcIcon(graphics, iconPen, centerX, centerY, true);
        } else if (iconType === "modeBow") {
            drawArcIcon(graphics, iconPen, centerX, centerY, false);
        }
    }

    /**
     * onDraw から iconType を束縛したクロージャを返す
     * @param {string} iconType - 描画種別
     * @returns {function} onDraw ハンドラ
     */
    function makeModeIconDrawer(iconType) {
        return function () {
            drawModeIcon(this, iconType, currentMode === iconType);
        };
    }

    /**
     * モードキーからモード名を返す
     * @param {string} modeKey - モードのキー
     * @returns {string} モード名（見つからない場合は空文字）
     */
    function getModeLabel(modeKey) {
        for (var i = 0; i < MODES.length; i++) {
            if (MODES[i].key === modeKey) return getLabel('mode.' + MODES[i].labelKey);
        }
        return '';
    }

    /**
     * モードのセル幅を、いちばん長いモード名に合わせて統一する（アイコンの間隔を揃えるため）
     * @param {object[]} cells - cell（縦グループ）と caption（モード名）を持つ配列
     * @returns {void}
     */
    function unifyModeCellWidths(cells) {
        var maxWidth = MODE_ICON_SIZE[0];
        for (var i = 0; i < cells.length; i++) {
            var measuredWidth = 0;
            try {
                var graphics = cells[i].caption.graphics;
                measuredWidth = graphics.measureString(cells[i].caption.text, graphics.font, 1000)[0];
            } catch (e) {
                measuredWidth = 0;
            }
            if (measuredWidth > maxWidth) maxWidth = measuredWidth;
        }
        for (var j = 0; j < cells.length; j++) {
            cells[j].cell.preferredSize.width = maxWidth;
            cells[j].caption.preferredSize.width = maxWidth;
        }
    }

    // =========================================
    // ダイアログ構築 / Building the dialog
    // =========================================

    /**
     * ラベル＋コントロールを横に並べる1行分のグループを追加する
     * @param {Panel|Group} parentGroup - 追加先のコンテナ
     * @returns {Group} 追加したグループ
     */
    function addFieldRow(parentGroup) {
        var fieldRow = parentGroup.add('group');
        fieldRow.orientation = 'row';
        fieldRow.alignChildren = ['left', 'center'];
        return fieldRow;
    }

    /**
     * 先頭ラベルのない行で、上の行と列を揃えるための空ラベルを置く
     * @param {Group} fieldRow - 対象の行グループ
     * @returns {void}
     */
    function addLabelSpacer(fieldRow) {
        setupLabelColumn(fieldRow.add('statictext', undefined, ''));
    }

    /**
     * 各行の先頭ラベルの幅を揃え、右寄せにする
     * @param {StaticText} rowLabel - 対象のラベル
     * @returns {void}
     */
    function setupLabelColumn(rowLabel) {
        rowLabel.preferredSize.width = LABEL_COLUMN_WIDTH;
        rowLabel.justify = 'right';
    }

    /**
     * オプションのパネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @param {string} titlePath - パネル名のラベルのパス
     * @returns {Panel} 追加したパネル
     */
    function addOptionsPanel(parentDialog, titlePath) {
        var optionsPanel = parentDialog.add('panel', undefined, getLabel(titlePath));
        setupPanel(optionsPanel);
        return optionsPanel;
    }

    /**
     * モード選択（アイコンで選ぶ）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function buildModeSelector(parentDialog) {
        var grpMode = parentDialog.add('group');
        grpMode.orientation = 'column';
        grpMode.alignChildren = ['fill', 'top'];
        grpMode.margins = MODE_GROUP_MARGINS;
        grpMode.spacing = 6;

        var grpModeIcons = grpMode.add('group');
        grpModeIcons.orientation = 'row';
        grpModeIcons.alignment = ['center', 'top'];
        grpModeIcons.spacing = 8;

        var modeCells = [];
        for (var modeIndex = 0; modeIndex < MODES.length; modeIndex++) {
            /* アイコンとその名前を縦に組にする / stack the icon and its name */
            var modeCell = grpModeIcons.add('group');
            modeCell.orientation = 'column';
            modeCell.alignChildren = 'center';
            modeCell.spacing = 5;

            var modeButton = modeCell.add('iconbutton', undefined, undefined, { style: 'toolbutton' });
            modeButton.preferredSize = MODE_ICON_SIZE;
            /* ヘルプチップは「モード名：説明」の形にする / help tip reads "name: description" */
            modeButton.helpTip = getModeLabel(MODES[modeIndex].key) + (uiLang === 'ja' ? '：' : ': ') + getLabel('tooltip.' + MODES[modeIndex].tipKey);
            modeButton.onDraw = makeModeIconDrawer(MODES[modeIndex].key);
            modeButtons.push(modeButton);

            var modeCaption = modeCell.add('statictext', undefined, getModeLabel(MODES[modeIndex].key));
            modeCaption.justify = 'center';
            modeCaption.helpTip = modeButton.helpTip;
            modeCells.push({ cell: modeCell, caption: modeCaption });
        }
        /* いちばん長いモード名に合わせてセル幅を統一し、アイコンの間隔を揃える
           Unify the cell widths to the longest name so the icons stay evenly spaced */
        unifyModeCellWidths(modeCells);
    }

    /**
     * ブロックのオプション（行の分け方・句読点・行送り）のパネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function buildBlockOptionsPanel(parentDialog) {
        var pnlBlockOptions = addOptionsPanel(parentDialog, 'panel.blockOptions');

        /* 行の分け方（ブロック専用）/ How to split lines (block mode only) */
        var grpLineSplit = addFieldRow(pnlBlockOptions);

        stLineSplit = grpLineSplit.add('statictext', undefined, labelText('fieldLabel.lineSplit'));
        setupLabelColumn(stLineSplit);
        stLineSplit.helpTip = getLabel('tooltip.lineSplit');
        rbLineSplitKeep = grpLineSplit.add('radiobutton', undefined, getLabel('radio.lineSplitKeep'));
        rbLineSplitKeep.helpTip = getLabel('tooltip.lineSplitKeep');
        rbLineSplitPunctuation = grpLineSplit.add('radiobutton', undefined, getLabel('radio.lineSplitPunctuation'));
        rbLineSplitPunctuation.helpTip = getLabel('tooltip.lineSplitPunctuation');

        /* 「行数を指定」は入力欄を伴うので次の行へ / the line-count choice carries an input, so it gets its own row */
        var grpLineCount = addFieldRow(pnlBlockOptions);

        /* 先頭は上の行と列を揃えるための空ラベル / empty label that keeps the column aligned */
        addLabelSpacer(grpLineCount);
        rbLineSplitCount = grpLineCount.add('radiobutton', undefined, getLabel('radio.lineSplitCount'));
        rbLineSplitCount.helpTip = getLabel('tooltip.lineSplitCount');
        /* 行数は1以上の整数 / line count: integer of 1 or more */
        etLineCount = addStepperInput(grpLineCount, String(BLOCK_LINE_COUNT), 4, {
            integer: true, min: 1,
            onStep: function () { refreshPreview(); }
        });
        etLineCount.helpTip = getLabel('tooltip.lineSplitCount');
        stLineCountUnit = grpLineCount.add('statictext', undefined, getLabel('unit.line'));
        stLineCountUnit.helpTip = getLabel('tooltip.lineSplitCount');

        setLineSplitMode(BLOCK_LINE_SPLIT);

        /* 行末の句読点の削除（改行を入れ直すときだけ使える）/ Drop the marks, only when the lines are re-cut */
        var grpRemovePunctuation = addFieldRow(pnlBlockOptions);

        /* 先頭は上の行と列を揃えるための空ラベル / empty label that keeps the column aligned */
        addLabelSpacer(grpRemovePunctuation);
        cbRemovePunctuation = grpRemovePunctuation.add('checkbox', undefined, getLabel('checkbox.removePunctuation'));
        cbRemovePunctuation.helpTip = getLabel('tooltip.removePunctuation');
        cbRemovePunctuation.value = BLOCK_REMOVE_PUNCTUATION;

        /* 行送り（ブロック専用）/ Leading (block mode only) */
        var grpLeading = addFieldRow(pnlBlockOptions);

        stLeading = grpLeading.add('statictext', undefined, labelText('fieldLabel.leading'));
        setupLabelColumn(stLeading);
        stLeading.helpTip = getLabel('tooltip.leadingAuto');
        cbAutoLeading = grpLeading.add('checkbox', undefined, getLabel('checkbox.leadingAuto'));
        cbAutoLeading.helpTip = getLabel('tooltip.leadingAuto');
        cbAutoLeading.value = BLOCK_AUTO_LEADING;
        /* 0以下は既定値に読み替えるので、∧∨は1で止める / non-positive values fall back to the default, so stop at 1 */
        etLeadingAmount = addStepperInput(grpLeading, String(BLOCK_AUTO_LEADING_AMOUNT), 6, {
            min: 1,
            onStep: function () { refreshPreview(); }
        });
        etLeadingAmount.helpTip = getLabel('tooltip.leadingAmount');
        stLeadingUnit = grpLeading.add('statictext', undefined, getLabel('unit.percent'));
        stLeadingUnit.helpTip = getLabel('tooltip.leadingAmount');
    }

    /**
     * パス上文字のオプション（カーブ・占有率・合わせ方・効果・改行の削除）のパネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function buildArcOptionsPanel(parentDialog) {
        var pnlArcOptions = addOptionsPanel(parentDialog, 'panel.arcOptions');

        /* カーブ / Roundness */
        var grpRoundness = addFieldRow(pnlArcOptions);

        stArcRoundness = grpRoundness.add('statictext', undefined, labelText('fieldLabel.roundness'));
        setupLabelColumn(stArcRoundness);
        stArcRoundness.helpTip = getLabel('tooltip.roundness');
        /* 0＝直線に近い、100＝最も丸い（初期値は最大）/ 0 = flat, 100 = roundest (defaults to the maximum) */
        slArcRoundness = grpRoundness.add('slider', undefined, ARC_ROUNDNESS_DEFAULT, 0, 100);
        slArcRoundness.preferredSize.width = SLIDER_WIDTH;
        slArcRoundness.helpTip = getLabel('tooltip.roundness');

        /* 占有率 / Coverage */
        var grpCoverage = addFieldRow(pnlArcOptions);

        stPathCoverage = grpCoverage.add('statictext', undefined, labelText('fieldLabel.coverage'));
        setupLabelColumn(stPathCoverage);
        stPathCoverage.helpTip = getLabel('tooltip.coverage');
        /* 100＝パスの端まで、PATH_COVERAGE_MIN＝パスの中央だけ / 100 reaches the path ends, PATH_COVERAGE_MIN keeps to the middle */
        slPathCoverage = grpCoverage.add('slider', undefined, PATH_COVERAGE_DEFAULT, PATH_COVERAGE_MIN, 100);
        /* カーブと同じ幅にそろえ、数値表示はそのうしろへ置く
           matches the Curve slider, with the readout placed after it */
        slPathCoverage.preferredSize.width = SLIDER_WIDTH;
        slPathCoverage.helpTip = getLabel('tooltip.coverage');
        stPathCoverageValue = grpCoverage.add('statictext', undefined, '');
        stPathCoverageValue.preferredSize.width = COVERAGE_VALUE_WIDTH;
        stPathCoverageValue.helpTip = getLabel('tooltip.coverage');

        /* 合わせ方：しない／文字サイズ＝サイズ変更／トラッキング＝サイズ維持で字間調整 / Fit */
        var grpFit = addFieldRow(pnlArcOptions);

        stFit = grpFit.add('statictext', undefined, labelText('fieldLabel.fit'));
        setupLabelColumn(stFit);
        stFit.helpTip = getLabel('tooltip.fit');
        rbFitNone = grpFit.add('radiobutton', undefined, getLabel('radio.fitNone'));
        rbFitNone.helpTip = getLabel('tooltip.fitNone');
        rbFitFontSize = grpFit.add('radiobutton', undefined, getLabel('radio.fitByFontSize'));
        rbFitFontSize.helpTip = getLabel('tooltip.fitByFontSize');
        rbFitTracking = grpFit.add('radiobutton', undefined, getLabel('radio.fitByTracking'));
        rbFitTracking.helpTip = getLabel('tooltip.fitByTracking');
        rbFitNone.value = (DEFAULT_FIT_METHOD === 'none');
        rbFitFontSize.value = (DEFAULT_FIT_METHOD === 'fontSize');
        rbFitTracking.value = (DEFAULT_FIT_METHOD === 'tracking');

        /* 効果 / Effect */
        var grpEffect = addFieldRow(pnlArcOptions);

        stEffect = grpEffect.add('statictext', undefined, labelText('fieldLabel.effect'));
        setupLabelColumn(stEffect);
        stEffect.helpTip = getLabel('tooltip.effect');

        var effectNames = [];
        for (var effectIndex = 0; effectIndex < EFFECTS.length; effectIndex++) {
            effectNames.push(getLabel('dropdown.' + EFFECTS[effectIndex].labelKey));
        }
        ddEffect = grpEffect.add('dropdownlist', undefined, effectNames);
        ddEffect.helpTip = getLabel('tooltip.effect');
        /* 既定はパス上文字の標準スタイルと同じ「虹」 / Default = Rainbow (Illustrator's own default) */
        ddEffect.selection = DEFAULT_EFFECT_INDEX;

        /* 改行の削除 / Remove line breaks */
        var grpRemoveLineBreaks = addFieldRow(pnlArcOptions);

        /* 先頭は上の行と列を揃えるための空ラベル / empty label that keeps the column aligned */
        addLabelSpacer(grpRemoveLineBreaks);
        cbRemoveLineBreaks = grpRemoveLineBreaks.add('checkbox', undefined, getLabel('checkbox.removeLineBreaks'));
        cbRemoveLineBreaks.helpTip = getLabel('tooltip.removeLineBreaks');
        cbRemoveLineBreaks.value = REMOVE_LINE_BREAKS;
    }

    /**
     * 共通のオプション（カーニング・トラッキング）のパネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function buildCommonOptionsPanel(parentDialog) {
        var pnlCommonOptions = addOptionsPanel(parentDialog, 'panel.commonOptions');

        /* カーニング（4つ横並びは幅を取るので2つずつ改行）/ Kerning, two choices per row */
        var kerningRow = null;
        for (var kerningIndex = 0; kerningIndex < KERNING_METHODS.length; kerningIndex++) {
            if (kerningIndex % 2 === 0) {
                kerningRow = addFieldRow(pnlCommonOptions);
                if (kerningIndex === 0) {
                    stAutoKerning = kerningRow.add('statictext', undefined, labelText('fieldLabel.autoKerning'));
                    setupLabelColumn(stAutoKerning);
                    stAutoKerning.helpTip = getLabel('tooltip.autoKerning');
                } else {
                    /* 先頭は上の行と列を揃えるための空ラベル / empty label that keeps the column aligned */
                    addLabelSpacer(kerningRow);
                }
            }

            var kerningButton = kerningRow.add('radiobutton', undefined,
                getLabel('radio.' + KERNING_METHODS[kerningIndex].labelKey));
            kerningButton.helpTip = getLabel('tooltip.' + KERNING_METHODS[kerningIndex].tipKey);
            kerningButtons.push(kerningButton);
        }
        setKerningMethod(DEFAULT_AUTO_KERNING);

        /* トラッキング / Tracking */
        var grpTracking = addFieldRow(pnlCommonOptions);

        stTracking = grpTracking.add('statictext', undefined, labelText('fieldLabel.tracking'));
        setupLabelColumn(stTracking);
        stTracking.helpTip = getLabel('tooltip.tracking');
        /* チェックOFFでトラッキング加算を無効化（値は0に固定）/ Checkbox OFF disables tracking (forced to 0) */
        cbTracking = grpTracking.add('checkbox', undefined, '');
        cbTracking.helpTip = getLabel('tooltip.trackingToggle');
        cbTracking.value = true;
        /* スライダーと同じ -100〜500 の整数。増減したらスライダーも追従させる / integer -100 to 500, kept in step with the slider */
        etTracking = addStepperInput(grpTracking, '0', 6, {
            integer: true, min: -100, max: 500,
            onStep: function () { syncTrackingFromEdit(); refreshPreview(); }
        });
        etTracking.helpTip = getLabel('tooltip.tracking');

        /* トラッキングのスライダーは幅を取るので次の行へ / the tracking slider needs room, so it gets its own row */
        var grpTrackingSlider = addFieldRow(pnlCommonOptions);

        /* 先頭は上の行と列を揃えるための空ラベル / empty label that keeps the column aligned */
        addLabelSpacer(grpTrackingSlider);
        slTracking = grpTrackingSlider.add('slider', undefined, 0, -100, 500);
        slTracking.preferredSize.width = SLIDER_WIDTH;
        slTracking.helpTip = getLabel('tooltip.tracking');
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

    /**
     * 表示の設定とフッター（制御文字・キャンセル・OK）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function buildFooter(parentDialog) {
        /* 表示の設定（ボタンの上に置く）/ View options, placed just above the buttons */
        var grpViewOptions = parentDialog.add('group');
        grpViewOptions.orientation = 'row';
        grpViewOptions.alignChildren = ['center', 'center'];
        /* グループ自体を中央に置く（ダイアログの fill を上書き）/ center the group itself, overriding the dialog's fill */
        grpViewOptions.alignment = ['center', 'top'];

        cbZoomToSelection = grpViewOptions.add('checkbox', undefined, getLabel('checkbox.zoomToSelection'));
        cbZoomToSelection.helpTip = getLabel('tooltip.zoomToSelection');
        cbZoomToSelection.value = ZOOM_TO_SELECTION;

        /* フッター / Footer */
        var buttonRow = addButtonRow(parentDialog);
        btnHiddenChar = buttonRow.leftGroup.add('button', undefined, getLabel('button.hiddenChar'));
        btnHiddenChar.helpTip = getLabel('tooltip.hiddenChar');
        btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'));
        btnCancel.helpTip = getLabel('tooltip.cancel');
        btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnOK.helpTip = getLabel('tooltip.ok');
    }

    /**
     * ダイアログを組み立てる（イベントは bindDialogEvents() で付ける）
     * @returns {void}
     */
    function buildDialog() {
        dynamicTextDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dynamicTextDialog);

        buildModeSelector(dynamicTextDialog);
        buildBlockOptionsPanel(dynamicTextDialog);
        buildArcOptionsPanel(dynamicTextDialog);
        buildCommonOptionsPanel(dynamicTextDialog);
        buildFooter(dynamicTextDialog);
    }

    // =========================================
    // 表示領域 / View
    // 同じ処理を再利用したいときは jsx/_templates/KeepInView.jsx を参照
    // see jsx/_templates/KeepInView.jsx to reuse this in another script
    // =========================================

    /**
     * 複数アイテムを囲む外接範囲を求める
     * @param {PageItem[]} pageItems - 対象アイテム
     * @returns {{left: number, top: number, right: number, bottom: number}|null} 外接範囲（求められない場合は null）
     */
    function getItemsBounds(pageItems) {
        var unionBounds = null;

        for (var i = 0; i < pageItems.length; i++) {
            var itemBounds;
            try {
                itemBounds = pageItems[i].visibleBounds; /* [left, top, right, bottom] */
            } catch (e) {
                continue;
            }
            if (unionBounds === null) {
                unionBounds = { left: itemBounds[0], top: itemBounds[1], right: itemBounds[2], bottom: itemBounds[3] };
                continue;
            }
            if (itemBounds[0] < unionBounds.left) unionBounds.left = itemBounds[0];
            if (itemBounds[1] > unionBounds.top) unionBounds.top = itemBounds[1];
            if (itemBounds[2] > unionBounds.right) unionBounds.right = itemBounds[2];
            if (itemBounds[3] < unionBounds.bottom) unionBounds.bottom = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 結果が可視領域に収まっていなければ、見えるように表示位置とズームを合わせる
     * すでに見えているときは何もしないので、操作のたびに画面が動くことはない
     * @param {PageItem[]} pageItems - 見えるようにしたいアイテム
     * @returns {void}
     */
    function ensureItemsVisible(pageItems) {
        if (!cbZoomToSelection.value) return;
        if (!pageItems || pageItems.length === 0) return;

        var targetBounds = getItemsBounds(pageItems);
        if (targetBounds === null) return;

        var activeView;
        try {
            activeView = doc.views[0];
        } catch (e) {
            return;
        }
        if (!activeView) return;

        var viewBounds = activeView.bounds; /* [left, top, right, bottom] */
        /* すでに全体が見えているなら動かさない / leave the view alone when everything is already visible */
        if (targetBounds.left >= viewBounds[0] && targetBounds.right <= viewBounds[2] &&
            targetBounds.top <= viewBounds[1] && targetBounds.bottom >= viewBounds[3]) return;

        var targetWidth = targetBounds.right - targetBounds.left;
        var targetHeight = targetBounds.top - targetBounds.bottom;
        var visibleWidth = viewBounds[2] - viewBounds[0];
        var visibleHeight = viewBounds[1] - viewBounds[3];

        /* 収まらないときだけズームアウトする。拡大はしない（操作のたびに倍率が変わると落ち着かない）
           Only zoom out when it does not fit; never zoom in, so the magnification stays predictable */
        var targetZoom = activeView.zoom;
        if (targetWidth > visibleWidth || targetHeight > visibleHeight) {
            var zoomByWidth = (targetWidth > 0) ? targetZoom * visibleWidth / targetWidth : targetZoom;
            var zoomByHeight = (targetHeight > 0) ? targetZoom * visibleHeight / targetHeight : targetZoom;
            targetZoom = Math.min(zoomByWidth, zoomByHeight) * VIEW_FIT_RATIO;
        }

        /* 中心を合わせてからズームする（ズームは中心を保つ）/ center first, then zoom about that center */
        activeView.centerPoint = [targetBounds.left + targetWidth / 2, targetBounds.top - targetHeight / 2];
        activeView.zoom = targetZoom;
    }

    // =========================================
    // プレビュー（Undoなし） / Preview (no undo)
    // =========================================

    /**
     * プレビューで作った一時オブジェクトを消し、隠した元のオブジェクトを表示に戻す
     * @returns {void}
     */
    function clearPreview() {
        /* 一時オブジェクトを削除（削除済みなら例外）/ Remove temp items (already removed ones throw) */
        for (var i = previewTempItems.length - 1; i >= 0; i--) {
            try { previewTempItems[i].remove(); } catch (e) { }
        }
        previewTempItems = [];

        /* 元のオブジェクトを表示に戻す / Restore originals visibility */
        for (var j = previewHiddenOriginals.length - 1; j >= 0; j--) {
            try { previewHiddenOriginals[j].hidden = false; } catch (e) { }
        }
        previewHiddenOriginals = [];
    }

    /**
     * プレビューのあいだ元のオブジェクトを隠す（clearPreview で戻す）
     * @param {PageItem} originalItem - 隠すオブジェクト
     * @returns {void}
     */
    function hideOriginalForPreview(originalItem) {
        if (!originalItem) return;
        for (var k = 0; k < previewHiddenOriginals.length; k++) {
            if (previewHiddenOriginals[k] === originalItem) return;
        }
        /* ロック中などで隠せないことがある / locked items may refuse */
        try {
            originalItem.hidden = true;
            previewHiddenOriginals.push(originalItem);
        } catch (e) { }
    }

    /**
     * 設定が変わったのでプレビューを貼り直す（プレビューは常時ON）
     * @returns {void}
     */
    function refreshPreview() {
        clearPreview();

        /* 開いたときの選択へ戻し、選択が変わってもプレビューが安定するようにする
           Restore base selection so preview stays stable even after selection changes */
        try { doc.selection = baseSelection; } catch (e) { }

        var currentSelection = [];
        try { currentSelection = doc.selection; } catch (e) { currentSelection = []; }
        if (!currentSelection || currentSelection.length === 0) {
            currentSelection = baseSelection;
        }
        targetTextFrames = collectSelectionTextFrames(currentSelection, TARGET_TEXT_OPTIONS);
        selectedPaths = collectSelectionPathItems(currentSelection, TARGET_PATH_OPTIONS);

        if (!targetTextFrames || targetTextFrames.length === 0) return;

        /* モードが選ばれるまでは何も変換しない / nothing to convert until a mode is picked */
        if (!isModeSelected()) return;

        generatePathText(false, true);
        app.redraw();
    }

    // =========================================
    // モードと有効・無効 / Modes and enabled states
    // =========================================

    /**
     * モードが選択されているか（起動直後は未選択）
     * @returns {boolean} 選択されていれば true
     */
    function isModeSelected() {
        return currentMode !== '';
    }

    /**
     * 円モードが選択されているか
     * @returns {boolean} 円モードなら true
     */
    function isCircleMode() {
        return currentMode === 'modeCircle';
    }

    /**
     * ブロックモードが選択されているか
     * @returns {boolean} ブロックモードなら true
     */
    function isBlockMode() {
        return currentMode === 'modeBlock';
    }

    /**
     * パス上文字を作るモードか（パスを作らないブロックと未選択は対象外）
     * @returns {boolean} パス上文字を作るなら true
     */
    function isPathTextMode() {
        return isModeSelected() && !isBlockMode();
    }

    /**
     * トラッキングでパス幅に合わせるモードか（ブロックは対象外）
     * @returns {boolean} 有効なら true
     */
    function isFitByTrackingActive() {
        return isPathTextMode() && rbFitTracking.value;
    }

    /**
     * 文字サイズでパス幅に合わせるモードか（ブロックは対象外）
     * @returns {boolean} 有効なら true
     */
    function isFitByFontSizeActive() {
        return isPathTextMode() && rbFitFontSize.value;
    }

    /**
     * モードアイコンを描き直す（選択状態の表示を更新する）
     * @returns {void}
     */
    function redrawModeIcons() {
        for (var i = 0; i < modeButtons.length; i++) {
            try { modeButtons[i].notify('onDraw'); } catch (e) { }
        }
    }

    /**
     * モードアイコンのクリック処理を、モードキーを束縛して返す
     * @param {string} modeKey - 選択するモードのキー
     * @returns {function} onClick ハンドラ
     */
    function makeModeSelector(modeKey) {
        return function () {
            if (currentMode === modeKey) return;
            currentMode = modeKey;
            onModeChanged();
        };
    }

    /**
     * パスを作らないブロックでは、パス幅に合わせる設定をディム表示する
     * @returns {void}
     */
    function updateFitEnabled() {
        var fitAvailable = isPathTextMode();
        stFit.enabled = fitAvailable;
        rbFitNone.enabled = fitAvailable;
        rbFitFontSize.enabled = fitAvailable;
        rbFitTracking.enabled = fitAvailable;
    }

    /**
     * ブロックはパス上文字を作らないため、アーチ用の行をまとめてディム表示する
     * @returns {void}
     */
    function updatePathTextControlsEnabled() {
        var arcActive = isPathTextMode();
        stArcRoundness.enabled = arcActive;
        slArcRoundness.enabled = arcActive;
        stPathCoverage.enabled = arcActive;
        slPathCoverage.enabled = arcActive;
        stPathCoverageValue.enabled = arcActive;
        stEffect.enabled = arcActive;
        ddEffect.enabled = arcActive;
        cbRemoveLineBreaks.enabled = arcActive;
    }

    /**
     * ブロックのオプションは、ブロックモードのときだけ操作できるようにする
     * @returns {void}
     */
    function updateBlockOptionsEnabled() {
        var blockActive = isBlockMode();

        stLineSplit.enabled = blockActive;
        rbLineSplitKeep.enabled = blockActive;
        rbLineSplitPunctuation.enabled = blockActive;
        rbLineSplitCount.enabled = blockActive;
        /* 行数の入力欄は「行数を指定」を選んでいるときだけ / the input follows the "line count" choice */
        var lineCountActive = blockActive && rbLineSplitCount.value;
        setStepperInputEnabled(etLineCount, lineCountActive);
        stLineCountUnit.enabled = lineCountActive;
        /* 句読点の削除は、改行を入れ直すときだけ / the marks are only dropped when the lines are re-cut */
        cbRemovePunctuation.enabled = blockActive && !rbLineSplitKeep.value;

        stLeading.enabled = blockActive;
        cbAutoLeading.enabled = blockActive;
        var leadingAmountActive = blockActive && cbAutoLeading.value;
        setStepperInputEnabled(etLeadingAmount, leadingAmountActive);
        stLeadingUnit.enabled = leadingAmountActive;
    }

    /**
     * カーニングの選択状態を切り替える
     * ラジオボタンが行ごとに別グループになり ScriptUI の自動排他が効かないため、明示的に設定する
     * @param {number} selectedIndex - 選択する KERNING_METHODS の添字
     * @returns {void}
     */
    function setKerningMethodByIndex(selectedIndex) {
        for (var i = 0; i < kerningButtons.length; i++) {
            kerningButtons[i].value = (i === selectedIndex);
        }
    }

    /**
     * カーニングの選択状態をキーで切り替える
     * @param {string} methodKey - KERNING_METHODS のキー
     * @returns {void}
     */
    function setKerningMethod(methodKey) {
        for (var i = 0; i < KERNING_METHODS.length; i++) {
            if (KERNING_METHODS[i].key === methodKey) {
                setKerningMethodByIndex(i);
                return;
            }
        }
        setKerningMethodByIndex(0);
    }

    /**
     * カーニングのクリック処理を、選択する添字を束縛して返す
     * @param {number} selectedIndex - 選択する KERNING_METHODS の添字
     * @returns {function} onClick ハンドラ
     */
    function makeKerningSelector(selectedIndex) {
        return function () {
            setKerningMethodByIndex(selectedIndex);
            refreshPreview();
        };
    }

    /**
     * カーニングは全モード共通。モードが選ばれるまではディム表示する
     * @returns {void}
     */
    function updateKerningEnabled() {
        var kerningAvailable = isModeSelected();
        stAutoKerning.enabled = kerningAvailable;
        for (var i = 0; i < kerningButtons.length; i++) kerningButtons[i].enabled = kerningAvailable;
    }

    /**
     * 「合わせ方：トラッキング」を選んでいる間は自動調整にまかせるので、手動トラッキング行をディム表示する
     * @returns {void}
     */
    function updateTrackingEnabled() {
        var trackingAvailable = isModeSelected() && !isFitByTrackingActive();
        stTracking.enabled = trackingAvailable;
        cbTracking.enabled = trackingAvailable;
        var manualTrackingActive = trackingAvailable && cbTracking.value;
        setStepperInputEnabled(etTracking, manualTrackingActive);
        slTracking.enabled = manualTrackingActive;
    }

    /**
     * 円モードではカーブを最大（＝ほぼ一周）にし、アーチに戻したら元の値へ戻す
     * @returns {void}
     */
    function syncRoundnessForMode() {
        if (isCircleMode()) {
            if (arcRoundnessBeforeCircle === null) arcRoundnessBeforeCircle = slArcRoundness.value;
            slArcRoundness.value = slArcRoundness.maxvalue;
        } else if (arcRoundnessBeforeCircle !== null) {
            slArcRoundness.value = arcRoundnessBeforeCircle;
            arcRoundnessBeforeCircle = null;
        }
    }

    /**
     * モードが変わったときに、表示と有効・無効をそろえてプレビューを貼り直す
     * @returns {void}
     */
    function onModeChanged() {
        redrawModeIcons();
        syncRoundnessForMode();
        updateFitEnabled();
        updatePathTextControlsEnabled();
        updateBlockOptionsEnabled();
        updateKerningEnabled();
        updateTrackingEnabled();
        refreshPreview();
    }

    /**
     * 合わせ方が変わったときの処理
     * @returns {void}
     */
    function onFitMethodChanged() {
        updateTrackingEnabled();
        refreshPreview();
    }

    /**
     * 行の分け方の選択状態を切り替える
     * ラジオボタンが2つのグループに分かれていて ScriptUI の自動排他が効かないため、明示的に設定する
     * @param {string} splitMode - keep／punctuation／count のいずれか
     * @returns {void}
     */
    function setLineSplitMode(splitMode) {
        rbLineSplitKeep.value = (splitMode === 'keep');
        rbLineSplitPunctuation.value = (splitMode === 'punctuation');
        rbLineSplitCount.value = (splitMode === 'count');
    }

    /**
     * 行の分け方のクリック処理を、選択する分け方を束縛して返す
     * @param {string} splitMode - 選択する分け方
     * @returns {function} onClick ハンドラ
     */
    function makeLineSplitSelector(splitMode) {
        return function () {
            setLineSplitMode(splitMode);
            updateBlockOptionsEnabled();
            refreshPreview();
        };
    }

    /**
     * 占有率スライダーの現在値を、右の数値表示へ反映する
     * @returns {void}
     */
    function syncPathCoverageValue() {
        stPathCoverageValue.text = Math.round(slPathCoverage.value) + getLabel('unit.percent');
    }

    /**
     * 占有率を変えたあとの共通処理（数値表示を合わせてプレビューを貼り直す）
     * @returns {void}
     */
    function onPathCoverageChanged() {
        syncPathCoverageValue();
        refreshPreview();
    }

    /**
     * トラッキングの入力欄の値を範囲内の整数にし、スライダーへ反映する
     * @returns {void}
     */
    function syncTrackingFromEdit() {
        if (trackingSyncLock) return;
        trackingSyncLock = true;
        var trackingValue = Math.max(-100, Math.min(500, Math.round(parseNumber(etTracking.text, 0))));
        etTracking.text = String(trackingValue);
        slTracking.value = trackingValue;
        trackingSyncLock = false;
    }

    /**
     * トラッキングのスライダーの値を入力欄へ反映する
     * @returns {void}
     */
    function syncTrackingFromSlider() {
        if (trackingSyncLock) return;
        trackingSyncLock = true;
        etTracking.text = String(Math.round(slTracking.value));
        trackingSyncLock = false;
    }

    // =========================================
    // イベント / Events
    // =========================================

    /**
     * ダイアログの各コントロールにイベントを付ける
     * @returns {void}
     */
    function bindDialogEvents() {
        for (var modeIndex = 0; modeIndex < modeButtons.length; modeIndex++) {
            modeButtons[modeIndex].onClick = makeModeSelector(MODES[modeIndex].key);
        }
        for (var kerningIndex = 0; kerningIndex < kerningButtons.length; kerningIndex++) {
            kerningButtons[kerningIndex].onClick = makeKerningSelector(kerningIndex);
        }

        /* ↑↓キーの増減は addStepperInput() で∧∨と同じ処理につないである / arrow keys are wired in addStepperInput() */

        /* カーブ：shift を押しながらのドラッグ・矢印キーで10刻み
           Curve: Shift snaps both dragging and the arrow keys to steps of 10 */
        slArcRoundness.onChanging = function () {
            snapSliderWithShift(slArcRoundness);
        };
        slArcRoundness.onChange = function () {
            snapSliderWithShift(slArcRoundness);
            refreshPreview();
        };
        changeSliderByArrowKey(slArcRoundness, refreshPreview);

        /* 占有率：ドラッグ中は数値だけ追従させ、離したときにプレビューを貼り直す
           the readout follows the drag; the preview is redrawn once the slider is released */
        slPathCoverage.onChanging = function () {
            snapSliderWithShift(slPathCoverage);
            syncPathCoverageValue();
        };
        slPathCoverage.onChange = function () {
            snapSliderWithShift(slPathCoverage);
            onPathCoverageChanged();
        };
        changeSliderByArrowKey(slPathCoverage, onPathCoverageChanged);

        rbFitNone.onClick = onFitMethodChanged;
        rbFitFontSize.onClick = onFitMethodChanged;
        rbFitTracking.onClick = onFitMethodChanged;
        ddEffect.onChange = refreshPreview;
        cbRemoveLineBreaks.onClick = refreshPreview;

        cbAutoLeading.onClick = function () {
            updateBlockOptionsEnabled();
            refreshPreview();
        };

        rbLineSplitKeep.onClick = makeLineSplitSelector('keep');
        rbLineSplitPunctuation.onClick = makeLineSplitSelector('punctuation');
        rbLineSplitCount.onClick = makeLineSplitSelector('count');
        cbRemovePunctuation.onClick = refreshPreview;
        etLineCount.onChange = refreshPreview;
        etLeadingAmount.onChange = refreshPreview;

        /* トラッキング UI 同期 / Tracking UI sync (edittext <-> slider) */
        etTracking.onChanging = function () { syncTrackingFromEdit(); refreshPreview(); };
        slTracking.onChanging = function () { syncTrackingFromSlider(); };
        slTracking.onChange = function () { syncTrackingFromSlider(); refreshPreview(); };
        cbTracking.onClick = function () {
            /* OFFにしたらトラッキングを0に戻す / Reset tracking to 0 when turned off */
            if (!cbTracking.value) {
                etTracking.text = '0';
                syncTrackingFromEdit();
            }
            updateTrackingEnabled();
            refreshPreview();
        };

        cbZoomToSelection.onClick = refreshPreview;

        btnHiddenChar.onClick = function () {
            /* 制御文字の表示はドキュメント側の設定なので、プレビューには手を触れない
               Hidden characters are a document-level setting, so the preview is left alone */
            try {
                app.executeMenuCommand('showHiddenChar');
                app.redraw();
            } catch (e) { }
        };

        btnCancel.onClick = function () {
            clearPreview();
            dynamicTextDialog.close(0);
        };

        btnOK.onClick = function () {
            /* モードを選ばずにOKされたら、閉じずに知らせる / stay open when no mode was picked */
            if (!isModeSelected()) {
                alert(getLabel('alert.selectMode'));
                return;
            }
            /* 一時オブジェクトを重ねないよう、プレビューを取り消してから本適用する / undo the preview before applying for real */
            clearPreview();
            if (!generatePathText(true, false)) {
                /* 何も適用できなかったので、設定を直せるよう開いたままにしてプレビューへ戻す
                   Nothing was applied: stay open so the settings can be fixed, and restore the preview */
                refreshPreview();
                return;
            }
            dynamicTextDialog.close(1);
        };
    }

    /**
     * 数値表示と各行の有効・無効を初期状態にそろえる
     * @returns {void}
     */
    function initDialogState() {
        syncPathCoverageValue();
        syncTrackingFromEdit();
        updateFitEnabled();
        updatePathTextControlsEnabled();
        updateBlockOptionsEnabled();
        updateKerningEnabled();
        updateTrackingEnabled();
    }

    // =========================================
    // テキスト・パス収集 / Collect text & paths
    // =========================================

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

    /* 変換対象のテキスト：ポイント文字・パス上文字・エリア内文字をグループの中まで（文字の選択は対象外）
       Target text: point, path and area text, including inside groups (text selections are not used) */
    var TARGET_TEXT_OPTIONS = { kinds: ["point", "path", "area"], textRangeToFrame: false, unique: false };

    /* 一緒に選択したパス：パスと複合パス（複合パスはまとめて1つ）をグループの中まで
       Paths selected alongside: paths and compound paths (as a whole), including inside groups */
    var TARGET_PATH_OPTIONS = { compoundPaths: "whole", textRangeToFrame: false, unique: false };

    // =========================================
    // パススタイル / Path style
    // =========================================

    /**
     * パス（複合パスなら中の各パス）に処理を行う（受け付けないものはスキップ）
     * @param {PageItem} pathItem - パスまたは複合パス
     * @param {Function} styleFn - PathItem を受け取る処理
     * @returns {void}
     */
    function forEachPathItem(pathItem, styleFn) {
        if (!pathItem || !styleFn) return;
        if (pathItem.typename === 'CompoundPathItem') {
            for (var subPathIndex = 0; subPathIndex < pathItem.pathItems.length; subPathIndex++) {
                try { styleFn(pathItem.pathItems[subPathIndex]); } catch (e) { }
            }
            return;
        }
        try { styleFn(pathItem); } catch (e) { }
    }

    /**
     * パスを見えなくする（塗りなし・線なし・線幅0）
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function styleInvisiblePath(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = false;
        pathItem.strokeWidth = 0;
    }

    /**
     * 生成したパスを見えないガイドにする（複合パスにも対応）
     * @param {PageItem} pathItem - 対象のパス
     * @returns {void}
     */
    function applyInvisiblePathStyle(pathItem) {
        forEachPathItem(pathItem, styleInvisiblePath);
    }

    // =========================================
    // アーチ生成 / Arc generation
    // =========================================

    /**
     * テキストと一緒に選ばれていたパスを残さない（プレビューでは隠し、本適用では削除）
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function removeOrHideSelectedPaths(previewMode) {
        if (!selectedPaths) return;
        for (var i = selectedPaths.length - 1; i >= 0; i--) {
            var selectedPath = selectedPaths[i];
            if (!selectedPath) continue;
            if (previewMode) {
                hideOriginalForPreview(selectedPath);
            } else {
                try { selectedPath.remove(); } catch (e) { }
            }
        }
    }

    /**
     * テキストフレームのすべての段落を中央揃えにする
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function applyCenterJustification(textFrame) {
        if (!textFrame) return;
        for (var i = 0; i < textFrame.paragraphs.length; i++) {
            /* 空段落など受け付けないものがあるので1段落ずつ守る / some paragraphs reject the change */
            try { textFrame.paragraphs[i].paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
        }
        try { textFrame.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
    }

    /**
     * テキストフレーム内の改行を削除して1行にまとめる
     * contents の置換では文字ごとの書式が失われるため、改行文字だけを後ろから削除する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function removeLineBreaks(textFrame) {
        if (!textFrame) return;
        var frameCharacters = textFrame.characters;
        for (var i = frameCharacters.length - 1; i >= 0; i--) {
            try {
                if (/^[\r\n]$/.test(frameCharacters[i].contents)) frameCharacters[i].remove();
            } catch (e) { }
        }
    }

    // =========================================
    // 効果・トラッキング / Effect & tracking
    // =========================================

    /**
     * 選ばれている効果のメニューコマンド名を返す
     * @returns {string|null} メニューコマンド名（未選択なら null）
     */
    function getSelectedEffectCommand() {
        if (!ddEffect.selection) return null;
        return EFFECTS[ddEffect.selection.index].command;
    }

    /**
     * 選ばれている効果をメニューコマンドで適用する（選択に対して働くので対象だけを選び直す）
     * @param {TextFrame} textFrame - 対象のパス上文字
     * @returns {void}
     */
    function applyPathTextEffect(textFrame) {
        var effectCommand = getSelectedEffectCommand();
        if (!effectCommand) return;

        /* メニューコマンドは選択に対して働くので、対象だけを選び直してから実行し、必ず元へ戻す
           The menu command acts on the selection, so isolate the target and always restore */
        var previousSelection = doc.selection;
        try {
            doc.selection = [];
            textFrame.selected = true;
            app.executeMenuCommand(effectCommand);
        } catch (e) { }
        try { doc.selection = previousSelection; } catch (e) { }
    }

    /**
     * フレームの textRange をすべて集める（無ければフレーム全体の textRange）
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {TextRange[]} textRange の配列
     */
    function collectTextRanges(textFrame) {
        var textRanges = [];
        try {
            if (textFrame.textRanges && textFrame.textRanges.length > 0) {
                for (var i = 0; i < textFrame.textRanges.length; i++) textRanges.push(textFrame.textRanges[i]);
            }
        } catch (e) { }
        if (textRanges.length === 0) {
            try { if (textFrame.textRange) textRanges = [textFrame.textRange]; } catch (e) { textRanges = []; }
        }
        return textRanges;
    }

    /**
     * フレーム内の各 textRange の文字属性に処理を行う（受け付けない範囲はスキップ）
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {Function} attributeAction - CharacterAttributes を受け取る処理
     * @returns {void}
     */
    function forEachRangeAttributes(textFrame, attributeAction) {
        var textRanges = collectTextRanges(textFrame);
        for (var rangeIndex = 0; rangeIndex < textRanges.length; rangeIndex++) {
            try {
                attributeAction(textRanges[rangeIndex].characterAttributes);
            } catch (e) {
                /* 受け付けない範囲はスキップ / skip ranges that reject the change */
            }
        }
    }

    /**
     * 選択中のカーニング方式を返す
     * @returns {object|null} AutoKernType の値。「そのまま」のときは null
     */
    function readKerningMethod() {
        for (var i = 0; i < kerningButtons.length; i++) {
            if (kerningButtons[i].value) return KERNING_METHODS[i].value;
        }
        return null;
    }

    /**
     * フレーム全体に自動カーニングを適用する
     * メトリクスのときだけプロポーショナルメトリクスもONにする（それ以外はOFF）
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function applyAutoKerning(textFrame) {
        var kerningMethod = readKerningMethod();
        /* 「そのまま」は既存の設定に手を触れない / "Keep" leaves the current settings alone */
        if (kerningMethod === null) return;

        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);

        forEachRangeAttributes(textFrame, function (charAttributes) {
            charAttributes.kerningMethod = kerningMethod;
            charAttributes.proportionalMetrics = useProportionalMetrics;
        });
    }

    /**
     * フレーム内の全 textRange のトラッキングに、指定量を加算する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number} trackingDelta - 加算する量（0なら何もしない）
     * @returns {void}
     */
    function addTrackingToFrame(textFrame, trackingDelta) {
        if (!textFrame || !trackingDelta) return;

        forEachRangeAttributes(textFrame, function (charAttributes) {
            charAttributes.tracking = charAttributes.tracking + trackingDelta;
        });
    }

    /**
     * ダイアログのトラッキング値を既存のトラッキングに加算する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function applyTrackingDelta(textFrame) {
        addTrackingToFrame(textFrame, Math.round(parseNumber(etTracking.text, 0)));
    }

    /**
     * 一時的なアウトラインで、描画されたテキストの範囲を測る
     * @param {TextFrame} sourceText - 測るテキスト
     * @returns {number[]|null} [L, T, R, B]（B は1行目のベースライン側）。測れなければ null
     */
    function measureTextBounds(sourceText) {
        try {
            /* [0] は1行目だけ（ベースラインに使う）、[1] は全体（左・上・右の端に使う）
               [0] holds only the first line (used for the baseline), [1] holds everything (left/top/right extents) */
            var measureTexts = [sourceText.duplicate(), sourceText.duplicate()];
            measureTexts[0].contents = '';
            for (var i = 0; i < sourceText.lines[0].length; i++) {
                sourceText.textRanges[i].duplicate(measureTexts[0]);
            }

            /* 改行を削除する設定のときは、実際にパスへ流し込む形＝1行につないだ状態で幅を測る。
               元の行のままだと最長行の幅しか得られず、パスが短すぎて文字が欠ける
               Measure the joined single line when the line breaks are going to be removed:
               otherwise the path would only span the widest line and the text would be clipped */
            if (cbRemoveLineBreaks.value) removeLineBreaks(measureTexts[1]);

            for (var k = 0; k < measureTexts.length; k++) {
                measureTexts[k] = measureTexts[k].createOutline();
            }

            var textBounds = measureTexts[1].geometricBounds; /* [L, T, R, B] */
            textBounds[3] = measureTexts[0].geometricBounds[3]; /* 1行目のベースライン / baseline from the first line */

            for (var copyIndex = 0; copyIndex < measureTexts.length; copyIndex++) {
                try { measureTexts[copyIndex].remove(); } catch (e) { }
            }
            return textBounds;
        } catch (e) {
            return null;
        }
    }

    /**
     * カーブスライダーの値を 0〜100 に収めて返す
     * @returns {number} カーブ（0〜100）
     */
    function readRoundnessPercent() {
        var percent = Number(slArcRoundness.value);
        if (isNaN(percent)) return ARC_ROUNDNESS_DEFAULT;
        return Math.max(0, Math.min(100, percent));
    }

    /**
     * 占有率スライダーの値を、パス長に対する比率として返す
     * 数値表示と結果を一致させるため、スライダーの値は整数に丸めて扱う。
     * @returns {number} 文字が占める割合（PATH_COVERAGE_MIN/100〜1）
     */
    function readPathCoverageRatio() {
        var percent = Math.round(slPathCoverage.value);
        if (isNaN(percent)) percent = PATH_COVERAGE_DEFAULT;
        return Math.max(PATH_COVERAGE_MIN, Math.min(100, percent)) / 100;
    }

    /**
     * 2点の直線パスを真円の円弧に変形する
     * カーブは弦の中点から頂点までの高さ（サジッタ）を決め、100で半幅＝半径、つまり弦を直径とする半円になる。
     * 円弧は頂点で2分割し、90度以下の区間ごとにベジェ曲線で近似する。
     * @param {PathItem} arcPath - ベースとなる2点の直線パス
     * @param {number} roundnessPercent - カーブ（0〜100）
     * @param {number} directionSign - 膨らむ向き（+1＝上／−1＝下）
     * @returns {void}
     */
    function applyArcHandles(arcPath, roundnessPercent, directionSign) {
        var startX = arcPath.pathPoints[0].anchor[0];
        var endX = arcPath.pathPoints[1].anchor[0];
        var baselineY = arcPath.pathPoints[0].anchor[1];
        var halfWidth = (endX - startX) / 2;
        if (!(halfWidth > 0)) return;

        var sagitta = halfWidth * (roundnessPercent / 100);
        if (!(sagitta > 0)) return; /* カーブ0は直線のまま */

        var radius = (sagitta * sagitta + halfWidth * halfWidth) / (2 * sagitta);
        /* 1区間の中心角。カーブ100で90度＝合計180度の半円になる */
        var segmentAngle = Math.atan2(halfWidth, radius - sagitta);
        /* 中心角θの円弧を1本のベジェで近似するハンドル長（θ=90度で半径×0.5523） */
        var handleLength = radius * (4 / 3) * Math.tan(segmentAngle / 4);
        /* 端点での接線方向へ倒したハンドルの成分 */
        var tangentX = handleLength * Math.cos(segmentAngle);
        var tangentY = handleLength * Math.sin(segmentAngle);

        var apexX = startX + halfWidth;
        var apexY = baselineY + sagitta * directionSign;

        arcPath.setEntirePath([
            [startX, baselineY],
            [apexX, apexY],
            [endX, baselineY]
        ]);

        var pathPoints = arcPath.pathPoints;
        pathPoints[0].rightDirection = [startX + tangentX, baselineY + tangentY * directionSign];
        pathPoints[1].leftDirection = [apexX - handleLength, apexY];
        pathPoints[1].rightDirection = [apexX + handleLength, apexY];
        pathPoints[2].leftDirection = [endX - tangentX, baselineY + tangentY * directionSign];
    }

    /**
     * 円モードで文字が占める円周の割合をカーブスライダーから求める
     * カーブ0＝大きい円（ゆるやかな弧）／カーブ100＝小さい円（ほぼ一周）
     * @returns {number} 円周に対する文字の割合（0〜1）
     */
    function readCircleOccupancy() {
        var ratio = readRoundnessPercent() / 100;
        return CIRCLE_MIN_OCCUPANCY + (CIRCLE_MAX_OCCUPANCY - CIRCLE_MIN_OCCUPANCY) * ratio;
    }

    /**
     * 4つのアンカーポイントで閉じた正円のパスを作成する
     * 始点を下（6時）に置き、下→左→上→右の時計回りにする。
     * 中央揃えの文字がパス長の中間＝円の上側中央に来るため、上向きの円弧文字になる。
     * @param {Layer} layer - パスを追加するレイヤー
     * @param {number} centerX - 円の中心X座標
     * @param {number} centerY - 円の中心Y座標
     * @param {number} radius - 円の半径
     * @returns {PathItem} 作成した閉じた円形パス
     */
    function createCirclePath(layer, centerX, centerY, radius) {
        var HANDLE_RATIO = 0.5522847498; /* ベジェ4分割で正円に近似する定数 / bezier constant for a 4-segment circle */
        var handleLength = radius * HANDLE_RATIO;

        var circlePath = layer.pathItems.add();
        circlePath.setEntirePath([
            [centerX, centerY - radius],  /* 下 / bottom */
            [centerX - radius, centerY],  /* 左 / left */
            [centerX, centerY + radius],  /* 上 / top */
            [centerX + radius, centerY]   /* 右 / right */
        ]);
        circlePath.closed = true;

        /* 各アンカーの方向線を接線方向へ倒して直線を円弧にする / tilt each handle along the tangent */
        var pathPoints = circlePath.pathPoints;
        pathPoints[0].leftDirection = [centerX + handleLength, centerY - radius];
        pathPoints[0].rightDirection = [centerX - handleLength, centerY - radius];
        pathPoints[1].leftDirection = [centerX - radius, centerY - handleLength];
        pathPoints[1].rightDirection = [centerX - radius, centerY + handleLength];
        pathPoints[2].leftDirection = [centerX - handleLength, centerY + radius];
        pathPoints[2].rightDirection = [centerX + handleLength, centerY + radius];
        pathPoints[3].leftDirection = [centerX + radius, centerY + handleLength];
        pathPoints[3].rightDirection = [centerX + radius, centerY - handleLength];

        try {
            circlePath.stroked = false;
            circlePath.filled = false;
        } catch (e) { }

        return circlePath;
    }

    /**
     * テキストの外接範囲から円形のパスを作成する
     * 文字幅が円周の一定割合になる半径を求め、円の頂点を元のベースライン位置に合わせる
     * @param {number[]} textBounds - 測定済みのテキスト範囲 [L, T, R, B]
     * @param {number} baselineY - 元テキストのベースラインY座標
     * @param {Layer} layer - パスを追加するレイヤー
     * @returns {PathItem|null} 作成した閉じた円形パス（幅が0のときは null）
     */
    function createCirclePathFromText(textBounds, baselineY, layer) {
        var textWidth = textBounds[2] - textBounds[0];
        if (!(textWidth > 0)) return null;

        var circumference = textWidth / readCircleOccupancy();
        var radius = circumference / (Math.PI * 2);
        var centerX = (textBounds[0] + textBounds[2]) / 2;

        return createCirclePath(layer, centerX, baselineY - radius, radius);
    }

    /**
     * ポイント文字の幅に合わせたアーチ（円モードでは円）のパスを作る
     * @param {TextFrame} sourceText - 変換元のポイント文字
     * @param {Layer} layer - パスを追加するレイヤー
     * @returns {PathItem|null} 作成したパス（作れなければ null）
     */
    function createArcPathFromText(sourceText, layer) {
        var baselineYMultiplier = 1.02;

        /* 空・不正なテキストは対象外 / Guard: empty / invalid text */
        try {
            if (!sourceText || sourceText.typename !== 'TextFrame') return null;
            if (!sourceText.lines || sourceText.lines.length === 0) return null;
            if (!sourceText.textRanges || sourceText.textRanges.length === 0) return null;
        } catch (e) {
            return null;
        }

        try {
            var textBounds = measureTextBounds(sourceText);
            if (!textBounds) return null;

            var baselineY = textBounds[3] * baselineYMultiplier;

            /* 円モードはアーチではなく閉じた円形パスを作る / Circle mode builds a closed circle instead */
            if (isCircleMode()) {
                return createCirclePathFromText(textBounds, baselineY, layer);
            }

            /* ベースラインに沿った直線のパス / Base straight path along the baseline */
            var arcPath = layer.pathItems.add();
            arcPath.setEntirePath([
                [textBounds[0], baselineY],
                [textBounds[2], baselineY]
            ]);
            try {
                arcPath.stroked = false;
                arcPath.filled = false;
            } catch (e) { }

            /* 直線を円弧に曲げる（上＝＋／下＝−）/ Bend the straight path into an arc */
            var directionSign = (currentMode === 'modeBow') ? -1 : 1;
            applyArcHandles(arcPath, readRoundnessPercent(), directionSign);

            return arcPath;
        } catch (e) {
            return null;
        }
    }

    /**
     * 1つのテキストから、パスとその上のパス上文字を作る
     * @param {TextFrame} sourceText - 変換元のポイント文字
     * @param {TextFrame} originalText - 選択されていた元のテキスト（重ね順の基準に使う）
     * @param {boolean} previewMode - プレビューなら true
     * @returns {TextFrame|null} 作成したパス上文字（パスを作れなかった場合は null）
     */
    function createPathTextFrom(sourceText, originalText, previewMode) {
        var currentLayer = originalText.layer;

        /* テキストの範囲からアーチ状のパスを作る / Create an arc-like path from the text bounds */
        var arcPath = createArcPathFromText(sourceText, currentLayer);
        if (!arcPath) return null;

        /* 作ったパスは見えないガイドにする / Generated arc path: invisible guide */
        applyInvisiblePathStyle(arcPath);
        if (previewMode) previewTempItems.push(arcPath);

        var textOnAPath = currentLayer.textFrames.pathText(arcPath);
        /* 重ね順を元のテキストに合わせる（ほかのオブジェクトの背面に隠れないように）/ Keep stacking position */
        try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (e) { }
        if (previewMode) previewTempItems.push(textOnAPath);

        /* 変換でスタイルが上書きされることがあるので、パス上文字のパスも見えなくする / Keep the text path invisible (AI may override its style) */
        if (textOnAPath.textPath) applyInvisiblePathStyle(textOnAPath.textPath);

        /* 変換元の textRange を複製する / Duplicate textRanges from the source text frame */
        for (var i = 0; i < sourceText.textRanges.length; i++) {
            sourceText.textRanges[i].duplicate(textOnAPath);
        }

        /* 改行の削除：1本のパスに沿わせるので、改行を消して1行にまとめる / Join into one line for a single path */
        if (cbRemoveLineBreaks.value) removeLineBreaks(textOnAPath);

        /* 行揃え：常に中央 ※ duplicate 後に適用しないと上書きされる / Always centered; must follow duplicate() */
        applyCenterJustification(textOnAPath);

        /* 自動カーニング ※ 文字幅が変わるので、トラッキングとフィットより先に適用する / Kerning comes before tracking and fitting */
        applyAutoKerning(textOnAPath);

        /* トラッキング（既存値 + 指定値）※ フィット前に適用してオーバーセット判定へ反映。
           「合わせ方：トラッキング」のときは手動値を加算しない（ディム表示と挙動を一致させる）
           Tracking is added before fitting, but not while "Fit: Tracking" is selected */
        if (!isFitByTrackingActive()) applyTrackingDelta(textOnAPath);

        /* 効果（パス上文字の効果をメニューコマンドで適用）/ Apply the path text effect */
        applyPathTextEffect(textOnAPath);

        return textOnAPath;
    }

    /**
     * 選択中のテキストをモードに合わせて変換する（ブロック、またはパスを作ってパス上文字に）
     * @param {boolean} showAlerts - 失敗時に警告を出すなら true
     * @param {boolean} previewMode - プレビューなら true
     * @returns {boolean} 1つでも変換できたら true
     */
    function generatePathText(showAlerts, previewMode) {
        var createdPathTexts = [];
        /* 変換元として一時的に作ったポイント文字。作り終えたら必ず取り除く
           Point-text stand-ins created as the source; always removed once the conversion is done */
        var temporarySources = [];

        /* モード未選択のまま呼ばれても何もしない / do nothing while no mode is picked */
        if (!isModeSelected()) return false;

        /* ブロックはパスを作らず、選択したテキストの行の幅をそろえるだけ / Block mode only fits the line widths */
        if (isBlockMode()) {
            return generateBlockText(showAlerts, previewMode);
        }

        /* テキストと一緒に選ばれていたパスは残さない / A path selected together with the text should not remain */
        removeOrHideSelectedPaths(previewMode);

        for (var j = 0; j < targetTextFrames.length; j++) {
            var originalText = targetTextFrames[j];

            /* パス上文字・エリア内文字は、いったん普通のテキスト（ポイント文字）へ戻してから作り直す。
               パス上文字は溢れて隠れていた行が戻り、エリア内文字は枠の折り返しが外れるため、
               どのモードでも同じ形のテキストを変換元にできる
               Path text and area type are turned back into plain point text first, so every mode
               starts from the same shape of text */
            var sourceText = originalText;
            if (!isPointTextFrame(originalText)) {
                sourceText = duplicateAsPointText(originalText);
                if (sourceText === null) {
                    if (showAlerts) alert(getLabel('alert.pathFailed'));
                    continue;
                }
                temporarySources.push(sourceText);
            }

            var textOnAPath = createPathTextFrom(sourceText, originalText, previewMode);
            if (textOnAPath === null) {
                if (showAlerts) alert(getLabel('alert.pathFailed'));
                continue;
            }
            createdPathTexts.push(textOnAPath);

            /* 元のテキストを削除する（プレビューでは隠す）/ Remove or hide the original text frame */
            if (previewMode) {
                hideOriginalForPreview(originalText);
            } else {
                originalText.remove();
                /* 作ったパス上文字を選択する / Select the created text on a path */
                textOnAPath.selected = true;
            }
        }

        /* 変換元の一時テキストを片づける / drop the temporary sources */
        for (var sourceIndex = temporarySources.length - 1; sourceIndex >= 0; sourceIndex--) {
            try { temporarySources[sourceIndex].remove(); } catch (e) { }
        }

        /* フィット：「しない」以外を選んだとき（ループ後にまとめて適用）。
           以下の処理はどれもフレームごとに例外を受け止めている
           Fit (unless None); each of these catches failures per frame */
        if (isFitByTrackingActive()) {
            /* 文字サイズを保ったまま、トラッキングでパス幅に合わせる */
            fitTextToPathByTracking(createdPathTexts);
        } else if (isFitByFontSizeActive()) {
            /* 文字サイズを変更してパス幅に合わせる（従来）*/
            fitTextToPathByFontSize(createdPathTexts);
        }

        /* 占有率：パスの端まで並んだ文字を、指定した割合ぶんまで詰めて中央へ寄せる */
        applyPathCoverage(createdPathTexts);

        /* 保険：フィットの設定にかかわらず、パスに収まらないぶんは縮めて文字を欠けさせない */
        preventOverset(createdPathTexts);

        /* 変換で位置が大きく変わるので、結果が画面から外れていたら見える位置へ / Bring the result into view */
        ensureItemsVisible(createdPathTexts);

        return createdPathTexts.length > 0;
    }

    // =========================================
    // ブロック / Block
    // =========================================

    /**
     * 改行や空白しか含まない行かどうかを判定する
     * @param {TextRange} textLine - 判定する行
     * @returns {boolean} 内容が空とみなせる場合 true
     */
    function isBlankLine(textLine) {
        return textLine.contents.replace(/[\r\n\x03\s　]/g, '').length === 0;
    }

    /**
     * 行の内容を一時テキストフレームへ複製し、アウトライン化して外形幅を測る
     * 一時オブジェクトは成否にかかわらず必ず削除する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {TextRange} textLine - 測定する行
     * @returns {number} 行の外形幅（pt）。測定できない場合は0
     */
    function measureLineWidth(targetDoc, textLine) {
        var tempTextFrame = null;
        var outlineGroup = null;
        var lineWidth = 0;

        try {
            tempTextFrame = targetDoc.textFrames.add();
            textLine.duplicate(tempTextFrame, ElementPlacement.INSIDE);
            /* createOutline() は元のテキストフレームを消費するので参照を手放す
               createOutline() consumes the source frame, so drop the reference */
            outlineGroup = tempTextFrame.createOutline();
            tempTextFrame = null;
            lineWidth = outlineGroup.width;
        } catch (outlineError) {
            /* アウトライン化できないときはフレーム幅で代用 / Fall back to the frame width */
            try { lineWidth = (tempTextFrame !== null) ? tempTextFrame.width : 0; } catch (e) { lineWidth = 0; }
        }

        /* 残っている一時オブジェクトを後始末する / Clean up whichever temporary object survived */
        try { if (outlineGroup !== null) outlineGroup.remove(); } catch (e) { }
        try { if (tempTextFrame !== null) tempTextFrame.remove(); } catch (e) { }

        return lineWidth;
    }

    /**
     * ストーリー内の指定範囲の文字サイズを変倍する（固定行送りの場合は行送りも追従させる）
     * @param {Story} story - 対象ストーリー
     * @param {number} startIndex - 開始文字インデックス
     * @param {number} endIndex - 終了文字インデックス（この位置は含まない）
     * @param {number} ratio - 変倍率
     * @returns {void}
     */
    function scaleCharacterSizes(story, startIndex, endIndex, ratio) {
        for (var i = startIndex; i < endIndex; i++) {
            try {
                var charAttributes = story.characters[i].characterAttributes;
                charAttributes.size *= ratio;
                /* 固定行送りのときだけ行が重ならないよう行送りも変倍する
                   Scale leading as well, but only when it is fixed */
                if (!charAttributes.autoLeading) {
                    charAttributes.leading *= ratio;
                }
            } catch (e) {
                /* 設定できない文字はスキップ / Skip characters that reject the change */
            }
        }
    }

    /**
     * 各行の文字サイズを最長行の幅にそろえる
     * 先に全行を測ってから変倍する（変倍で行が再合成されても対象がずれないようにするため）
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {boolean} 1行でも測定できた場合 true
     */
    function fitLinesToWidestLine(targetDoc, textFrame) {
        var frameLines = textFrame.lines;
        var lineMetrics = [];
        var maxWidth = 0;
        var i;

        /* 1. 全行の外形幅と最大幅を測る（この時点ではテキストを変更しない）
           1. Measure every line and the widest width; the text is not touched yet */
        for (i = 0; i < frameLines.length; i++) {
            var textLine = frameLines[i];
            var lineWidth = isBlankLine(textLine) ? 0 : measureLineWidth(targetDoc, textLine);
            lineMetrics.push({ start: textLine.start, end: textLine.end, width: lineWidth });
            if (lineWidth > maxWidth) {
                maxWidth = lineWidth;
            }
        }

        if (maxWidth === 0) {
            return false;
        }

        /* 2. 行ごとに変倍する。行オブジェクトではなくストーリー内の文字インデックスで指定して
              変倍による行の再合成の影響を受けないようにする
           2. Scale line by line, addressing characters by story index rather than by line object
              so re-composition during scaling cannot shift the target range */
        var story = textFrame.story;
        for (i = 0; i < lineMetrics.length; i++) {
            var lineMetric = lineMetrics[i];
            if (lineMetric.width <= 0) continue;

            var scaleRatio = maxWidth / lineMetric.width;
            /* 幅がほぼ同等の行はスキップ / Skip lines that already match the widest one */
            if (Math.abs(scaleRatio - 1) < BLOCK_RATIO_EPSILON) continue;

            scaleCharacterSizes(story, lineMetric.start, lineMetric.end, scaleRatio);
        }

        return true;
    }

    /**
     * テキストフレームの行送りを自動に切り替え、各段落の自動行送り比率を設定する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number} autoLeadingAmount - 自動行送りの比率（％）
     * @returns {void}
     */
    function applyAutoLeading(textFrame, autoLeadingAmount) {
        /* フレーム全体を自動行送りにする / Switch the whole frame to auto leading */
        try { textFrame.textRange.characterAttributes.autoLeading = true; } catch (e) { }

        var frameParagraphs = textFrame.paragraphs;
        for (var i = 0; i < frameParagraphs.length; i++) {
            try {
                frameParagraphs[i].characterAttributes.autoLeading = true;
                frameParagraphs[i].paragraphAttributes.autoLeadingAmount = autoLeadingAmount;
            } catch (e) {
                /* 空段落など設定できないものはスキップ / Skip paragraphs that reject the setting */
            }
        }
    }

    /**
     * 行送り入力欄から自動行送りの比率を読み取る
     * @returns {number} 自動行送りの比率（％）
     */
    function readLeadingAmount() {
        var amount = parseNumber(etLeadingAmount.text, BLOCK_AUTO_LEADING_AMOUNT);
        /* 0以下や数値でない入力は既定値に読み替える / fall back to the default for non-positive input */
        if (!(amount > 0)) amount = BLOCK_AUTO_LEADING_AMOUNT;
        return amount;
    }

    /**
     * 選ばれている行の分け方を返す
     * @returns {string} keep／punctuation／count のいずれか
     */
    function readLineSplitMode() {
        if (rbLineSplitPunctuation.value) return 'punctuation';
        if (rbLineSplitCount.value) return 'count';
        return 'keep';
    }

    /**
     * 「行数を指定」の入力値を読み取る
     * @returns {number} 分ける行数（1以上の整数）
     */
    function readLineCount() {
        var lineCount = Math.round(parseNumber(etLineCount.text, BLOCK_LINE_COUNT));
        if (!(lineCount > 0)) lineCount = BLOCK_LINE_COUNT;
        return lineCount;
    }

    /**
     * 指定位置に改行を挿入する（contents の書き換えでは文字ごとの書式が失われるため）
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number[]} positions - 挿入位置（文字インデックス）の昇順配列
     * @returns {void}
     */
    function insertLineBreaks(textFrame, positions) {
        /* 後ろから挿入すれば、まだ処理していない前側のインデックスがずれない
           Insert from the end so the not-yet-used earlier indexes stay valid */
        for (var i = positions.length - 1; i >= 0; i--) {
            try {
                textFrame.insertionPoints[positions[i]].characters.add('\r');
            } catch (e) { }
        }
    }

    /**
     * 句読点のうしろの改行位置を求める
     * @param {string} text - 対象の文字列（改行を外した状態）
     * @param {number} markIndex - 句読点とみなす文字の位置
     * @returns {number} 改行を入れる位置（改行しないときは -1）
     */
    function findBreakAfterPunctuation(text, markIndex) {
        if (markIndex < 0 || markIndex >= text.length) return -1;
        /* 末尾の句読点で改行すると空行ができるので入れない / no break after the final mark */
        if (markIndex + 1 >= text.length) return -1;

        var mark = text.charAt(markIndex);
        var isLatinMark = (BLOCK_PUNCTUATION_LATIN.indexOf(mark) >= 0);
        if (!isLatinMark && BLOCK_PUNCTUATION.indexOf(mark) < 0) return -1;

        /* 閉じ括弧・引用符が続くときは、それも前の行に残す（「〜です。」の 」 が行頭に落ちるのを防ぐ）*/
        var position = markIndex + 1;
        while (position < text.length && BLOCK_CLOSING_MARKS.indexOf(text.charAt(position)) >= 0) position++;
        if (position >= text.length) return -1;

        var nextChar = text.charAt(position);
        /* 欧文は 3.14 や e.g. で切ってしまわないよう、うしろのスペースを条件にする */
        if (isLatinMark && BLOCK_BREAK_SPACES.indexOf(nextChar) < 0) return -1;
        /* 「！？」のように句読点が続くときは、あとの句読点にゆずる */
        if (BLOCK_PUNCTUATION.indexOf(nextChar) >= 0 || BLOCK_PUNCTUATION_LATIN.indexOf(nextChar) >= 0) return -1;

        /* 次の行が空白で始まらないよう、うしろのスペースは前の行に残す */
        while (position < text.length && BLOCK_BREAK_SPACES.indexOf(text.charAt(position)) >= 0) position++;
        if (position >= text.length) return -1;

        return position;
    }

    /**
     * 単語の切れ目（スペース）のうしろの改行位置を求める
     * @param {string} text - 対象の文字列（改行を外した状態）
     * @param {number} spaceIndex - スペースとみなす文字の位置
     * @returns {number} 改行を入れる位置（改行しないときは -1）
     */
    function findBreakAfterSpace(text, spaceIndex) {
        if (spaceIndex < 0 || spaceIndex >= text.length) return -1;
        if (BLOCK_BREAK_SPACES.indexOf(text.charAt(spaceIndex)) < 0) return -1;

        /* 次の行が空白で始まらないよう、続くスペースはまとめて前の行に残す */
        var position = spaceIndex;
        while (position < text.length && BLOCK_BREAK_SPACES.indexOf(text.charAt(position)) >= 0) position++;
        if (position >= text.length) return -1;

        return position;
    }

    /**
     * 句読点のうしろを改行位置として拾う
     * @param {string} text - 対象の文字列（改行を外した状態）
     * @returns {number[]} 改行を入れる位置の配列
     */
    function findPunctuationBreaks(text) {
        var positions = [];
        for (var i = 0; i < text.length; i++) {
            var position = findBreakAfterPunctuation(text, i);
            if (position > 0) positions.push(position);
        }
        return positions;
    }

    /**
     * 目標位置の近くの区切りへ改行位置を寄せる
     * 句読点を優先し、見つからないときは単語の切れ目を使う（欧文で単語の途中が切れるのを防ぐ）
     * @param {string} text - 対象の文字列
     * @param {number} targetPosition - 均等割りで求めた位置
     * @param {number} snapRange - 前後に探す文字数
     * @returns {number} 実際に改行する位置
     */
    function snapToBreakPoint(text, targetPosition, snapRange) {
        for (var offset = 0; offset <= snapRange; offset++) {
            var forwardMark = findBreakAfterPunctuation(text, targetPosition + offset - 1);
            if (forwardMark > 0) return forwardMark;
            var backwardMark = findBreakAfterPunctuation(text, targetPosition - offset - 1);
            if (backwardMark > 0) return backwardMark;
        }
        for (var spaceOffset = 0; spaceOffset <= snapRange; spaceOffset++) {
            var forwardSpace = findBreakAfterSpace(text, targetPosition + spaceOffset - 1);
            if (forwardSpace > 0) return forwardSpace;
            var backwardSpace = findBreakAfterSpace(text, targetPosition - spaceOffset - 1);
            if (backwardSpace > 0) return backwardSpace;
        }
        return targetPosition;
    }

    /**
     * 指定行数へ均等に分ける改行位置を求める（近くの句読点、なければ単語の切れ目を優先）
     * @param {string} text - 対象の文字列（改行を外した状態）
     * @param {number} lineCount - 分ける行数
     * @returns {number[]} 改行を入れる位置の配列
     */
    function findEvenBreaks(text, lineCount) {
        var positions = [];
        var textLength = text.length;
        if (lineCount < 2 || textLength < lineCount) return positions;

        var charactersPerLine = textLength / lineCount;
        var snapRange = Math.floor(charactersPerLine / 3);
        var previousPosition = 0;

        for (var i = 1; i < lineCount; i++) {
            var position = snapToBreakPoint(text, Math.round(charactersPerLine * i), snapRange);
            /* 空行を作らないよう、前の改行位置より必ず後ろにする / never produce an empty line */
            if (position <= previousPosition) position = previousPosition + 1;
            if (position >= textLength) break;
            positions.push(position);
            previousPosition = position;
        }
        return positions;
    }

    /**
     * 「行末の句読点を削除」の指定を読み取る
     * @returns {boolean} 削除するなら true
     */
    function readRemovePunctuation() {
        return cbRemovePunctuation.value === true;
    }

    /**
     * 行の終わりに残った句読点を削除する（閉じ括弧・引用符・空白は残す）
     * 改行を入れたあとに実行する。contents の書き換えでは文字ごとの書式が失われるため1文字ずつ消す。
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function removeLineEndPunctuation(textFrame) {
        var text = '';
        try { text = textFrame.contents; } catch (e) { return; }

        /* うしろから見れば、まだ消していない前側のインデックスがずれない
           Scan from the end so the not-yet-used earlier indexes stay valid */
        var atLineEnd = true; /* ここから行末までが閉じ括弧・空白だけか */
        for (var i = text.length - 1; i >= 0; i--) {
            var currentChar = text.charAt(i);
            if (currentChar === '\r' || currentChar === '\n') { atLineEnd = true; continue; }
            if (!atLineEnd) continue;
            if (BLOCK_CLOSING_MARKS.indexOf(currentChar) >= 0 || BLOCK_BREAK_SPACES.indexOf(currentChar) >= 0) continue;
            if (BLOCK_PUNCTUATION.indexOf(currentChar) >= 0 || BLOCK_PUNCTUATION_LATIN.indexOf(currentChar) >= 0) {
                /* 「！？」のように続くときはまとめて消すので atLineEnd は下ろさない */
                try { textFrame.characters[i].remove(); } catch (e) { }
                continue;
            }
            atLineEnd = false;
        }
    }

    /**
     * ブロックにする前に、テキストの行の分け方を整える
     * 「そのまま」以外は、いまの改行をいったん外してから分け直す
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function applyLineSplit(textFrame) {
        var splitMode = readLineSplitMode();
        if (splitMode === 'keep') return;

        /* 行をまたいで混ざったサイズを先にそろえる / level the sizes before the lines are re-cut */
        unifyFontSize(textFrame);
        removeLineBreaks(textFrame);

        var text = '';
        try { text = textFrame.contents; } catch (e) { return; }
        if (text.length === 0) return;

        var positions = (splitMode === 'punctuation')
            ? findPunctuationBreaks(text)
            : findEvenBreaks(text, readLineCount());
        insertLineBreaks(textFrame, positions);

        /* 改行位置が決まったあとに消す（先に消すと位置がずれる）*/
        if (readRemovePunctuation()) removeLineEndPunctuation(textFrame);
    }

    /**
     * ポイント文字かどうかを判定する
     * @param {TextFrame} textFrame - 判定するテキストフレーム
     * @returns {boolean} ポイント文字なら true
     */
    function isPointTextFrame(textFrame) {
        if (!textFrame) return false;
        try { return textFrame.kind === TextType.POINTTEXT; } catch (e) { }
        return false;
    }

    /**
     * テキストフレームの中身をポイント文字へ写した複製を作る（元は残す）
     * エリア内文字の折り返しや、パス上文字が1行しか表示できない制約から外れるため、
     * 行の測定や行の分け直しはこの複製に対して行う
     * @param {TextFrame} sourceTextFrame - 写し取る元のテキストフレーム
     * @returns {TextFrame|null} 作成したポイント文字（失敗した場合は null）
     */
    function duplicateAsPointText(sourceTextFrame) {
        var pointText = null;
        try {
            var sourceBounds = sourceTextFrame.geometricBounds; /* [L, T, R, B] */
            pointText = sourceTextFrame.layer.textFrames.add();

            /* 中身をそのまま移して文字ごとの書式を保つ / carry the per-character formatting over */
            for (var i = 0; i < sourceTextFrame.textRanges.length; i++) {
                sourceTextFrame.textRanges[i].duplicate(pointText);
            }

            /* 行揃えは textRanges の複製では移らないので個別に写す
               Justification does not travel with the ranges, so copy it separately */
            try {
                pointText.textRange.paragraphAttributes.justification =
                    sourceTextFrame.textRange.paragraphAttributes.justification;
            } catch (e) { }

            /* 重ね順と位置を元のテキストに合わせる / keep the stacking order and position */
            try { pointText.move(sourceTextFrame, ElementPlacement.PLACEBEFORE); } catch (e) { }
            try { pointText.position = [sourceBounds[0], sourceBounds[1]]; } catch (e) { }

            return pointText;
        } catch (copyError) {
            try { if (pointText !== null) pointText.remove(); } catch (e) { }
            return null;
        }
    }

    /**
     * パス上文字・エリア内文字をポイント文字へ変換する（元のテキストフレームは取り除く）
     * Illustratorには解除のコマンドがないため作り直す
     * @param {TextFrame} textFrame - 変換するテキストフレーム
     * @returns {TextFrame|null} 作成したポイント文字（失敗した場合は null）
     */
    function convertToPointText(textFrame) {
        var pointText = duplicateAsPointText(textFrame);
        if (pointText === null) return null;
        try { textFrame.remove(); } catch (e) { }
        return pointText;
    }

    /**
     * フレーム内の文字サイズの最小値と最大値を返す
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {{smallest: number, largest: number}} 文字サイズの範囲（読めなければ 0）
     */
    function getFontSizeRange(textFrame) {
        var smallest = 0;
        var largest = 0;

        forEachRangeAttributes(textFrame, function (charAttributes) {
            var fontSize = charAttributes.size;
            if (smallest === 0 || fontSize < smallest) smallest = fontSize;
            if (fontSize > largest) largest = fontSize;
        });
        return { smallest: smallest, largest: largest };
    }

    /**
     * フレーム内の文字サイズの平均を返す
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {number} 文字サイズの平均（読めなければ0）
     */
    function getAverageFontSize(textFrame) {
        var total = 0;
        var counted = 0;

        forEachRangeAttributes(textFrame, function (charAttributes) {
            total += charAttributes.size;
            counted++;
        });
        return (counted > 0) ? (total / counted) : 0;
    }

    /**
     * フレーム全体の文字サイズに同じ比率を掛ける（文字ごとのサイズ差を保つ）
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number} ratio - 変倍率
     * @returns {void}
     */
    function scaleFontSize(textFrame, ratio) {
        forEachRangeAttributes(textFrame, function (charAttributes) {
            charAttributes.size = charAttributes.size * ratio;
        });
    }

    /**
     * フレーム全体の文字サイズを、いちばん大きい値にそろえる
     * 行を分け直すと、以前のブロック処理で行ごとに違っていたサイズが1行の中に混ざるため、
     * 分け直す前にそろえる。もともと均一なテキストでは何も変わらない
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function unifyFontSize(textFrame) {
        var largestSize = getFontSizeRange(textFrame).largest;
        if (!(largestSize > 0)) return;

        forEachRangeAttributes(textFrame, function (charAttributes) {
            charAttributes.size = largestSize;
        });
    }

    /**
     * ブロックを適用する
     * プレビューはUndoを使わない仕組みのため、複製へ適用して元のテキストを隠す
     * @param {boolean} showAlerts - 対象が1つもないときに警告を出すなら true
     * @param {boolean} previewMode - プレビューなら true
     * @returns {boolean} 1つでも適用できたら true
     */
    function generateBlockText(showAlerts, previewMode) {
        var appliedTexts = [];

        for (var i = 0; i < targetTextFrames.length; i++) {
            var originalText = targetTextFrames[i];

            /* プレビューはUndoを使わないため、複製へ適用して元のテキストは隠す
               Preview never undoes, so work on a duplicate and hide the original */
            var workingText = originalText;
            if (previewMode) {
                try {
                    workingText = originalText.duplicate();
                } catch (e) {
                    continue;
                }
            }

            /* ポイント文字以外は行ごとに測れないので、いったんポイント文字へ変換する。
               パス上文字はアーチや円にしたときに溢れて隠れていた行がここで戻り、
               エリア内文字は枠の折り返しが外れて段落がそのまま行になる
               Anything other than point text cannot be measured line by line, so convert it first */
            if (!isPointTextFrame(workingText)) {
                var convertedText = convertToPointText(workingText);
                if (convertedText === null) {
                    discardWorkingText(workingText, originalText, previewMode);
                    continue;
                }
                workingText = convertedText;
            }

            /* 幅をそろえる前に、行の分け方を整える / decide the lines before fitting their widths */
            applyLineSplit(workingText);

            /* 1行では最長行が自分自身になり変倍が起きないため対象外 / a single line has nothing to fit to */
            if (getLineAmount(workingText) <= 1) {
                discardWorkingText(workingText, originalText, previewMode);
                continue;
            }

            /* カーニングとトラッキングは行の幅を変えるので、幅をそろえる前に適用する
               Kerning and tracking change the line widths, so both come before the fitting */
            applyAutoKerning(workingText);
            applyTrackingDelta(workingText);

            if (!fitLinesToWidestLine(doc, workingText)) {
                discardWorkingText(workingText, originalText, previewMode);
                continue;
            }

            if (cbAutoLeading.value) applyAutoLeading(workingText, readLeadingAmount());

            if (previewMode) {
                previewTempItems.push(workingText);
                hideOriginalForPreview(originalText);
            } else {
                /* 変換でオブジェクトが差し替わっているので、選択を作り直したテキストへ移す */
                try { workingText.selected = true; } catch (e) { }
            }
            appliedTexts.push(workingText);
        }

        if (appliedTexts.length === 0) {
            if (showAlerts) alert(getLabel('alert.needTwoLines'));
            return false;
        }

        /* 行の分け直しで大きさが変わるので、結果が画面から外れていたら見える位置へ / Bring the result into view */
        ensureItemsVisible(appliedTexts);
        return true;
    }

    /**
     * 適用できなかった作業用テキストを片づける
     * プレビュー用の複製だけを取り除き、本適用で解除済みのテキストは残す（パス上文字へは戻せないため）
     * @param {TextFrame} workingText - 作業対象のテキストフレーム
     * @param {TextFrame} originalText - 元のテキストフレーム
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function discardWorkingText(workingText, originalText, previewMode) {
        if (!previewMode) return;
        if (workingText === originalText) return;
        try { workingText.remove(); } catch (e) { }
    }

    // =========================================
    // フィット / Fit
    // =========================================

    /**
     * 編集できるパス上文字か（フィットの対象。開いたパス・閉じたパスとも）
     * @param {TextFrame} textFrame - 判定するテキストフレーム
     * @returns {boolean} 対象なら true
     */
    function isEditablePathText(textFrame) {
        try {
            if (!textFrame || textFrame.typename !== 'TextFrame') return false;
            if (textFrame.kind !== TextType.PATHTEXT) return false;
            if (!textFrame.editable || textFrame.locked || textFrame.hidden) return false;
            return !!textFrame.textPath;
        } catch (e) { }
        return false;
    }

    /**
     * あふれ（表示される行に収まらない文字がある）かどうかを判定する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number} [lineAmount] - 表示される行数（省略時は1）
     * @returns {boolean} あふれていれば true
     */
    function isOverset(textFrame, lineAmount) {
        try {
            if (!textFrame) return false;

            if (textFrame.lines.length > 0) {
                var charactersOnVisibleLines = 0;

                if (typeof (lineAmount) === 'undefined' || lineAmount === null) {
                    lineAmount = 1;
                } else {
                    lineAmount = Math.floor(lineAmount);
                    if (lineAmount < 1) lineAmount = 1;
                    if (lineAmount > textFrame.lines.length) lineAmount = textFrame.lines.length;
                }

                for (var i = 0; i < lineAmount; i++) {
                    charactersOnVisibleLines += textFrame.lines[i].characters.length;
                }
                return (charactersOnVisibleLines < textFrame.characters.length);
            } else if (textFrame.characters.length > 0) {
                return true;
            }
        } catch (e) { }
        return false;
    }

    /**
     * フレームの表示行数を返す（常に1以上）
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {number} 行数
     */
    function getLineAmount(textFrame) {
        try {
            if (textFrame.lines && textFrame.lines.length > 0) return textFrame.lines.length;
        } catch (e) { }
        return 1;
    }

    /**
     * 文字サイズを変えずに、文字の長さをパス長の指定割合まで詰めるトラッキング量を求める
     * トラッキングは em の1/1000 単位なので、1文字あたり「文字サイズ×量/1000」だけ長さが変わる。
     * フィット直後は文字の長さ＝パス長とみなせるため、詰める量はパス長から求められる。
     * @param {TextFrame} textFrame - 対象のパス上文字
     * @param {number} coverageRatio - パス長に対する割合（0〜1）
     * @returns {number} 加算するトラッキング量（求められない場合は0）
     */
    function calcTrackingForCoverage(textFrame, coverageRatio) {
        var pathLength = 0;
        try { pathLength = textFrame.textPath.length; } catch (e) { pathLength = 0; }
        if (!(pathLength > 0)) return 0;

        var characterAmount = textFrame.characters.length;
        var averageFontSize = getAverageFontSize(textFrame);
        if (!(characterAmount > 0) || !(averageFontSize > 0)) return 0;

        var shortenBy = pathLength * (1 - coverageRatio);
        return -Math.round(shortenBy * 1000 / (averageFontSize * characterAmount));
    }

    /**
     * パス長のうち文字が占める割合を、フィットのあとに詰めて合わせる
     * 中央揃えなので、詰めたぶんだけ文字はパスの中央（円では上側中央）へ寄る。
     * 「合わせ方：トラッキング」は文字サイズを保つ設定なので、サイズではなく字間を詰めて短くする。
     * @param {TextFrame[]} pathTexts - 生成したパス上文字
     * @returns {void}
     */
    function applyPathCoverage(pathTexts) {
        if (!pathTexts || pathTexts.length === 0) return;

        var coverageRatio = readPathCoverageRatio();
        if (coverageRatio >= 1) return; /* 100％はパスの端まで＝何も詰めない */

        var keepFontSize = isFitByTrackingActive();

        for (var i = 0; i < pathTexts.length; i++) {
            var textFrame = pathTexts[i];
            if (!isEditablePathText(textFrame)) continue;

            try {
                if (textFrame.characters.length <= 0) continue;

                if (keepFontSize) {
                    addTrackingToFrame(textFrame, calcTrackingForCoverage(textFrame, coverageRatio));
                } else {
                    scaleFontSize(textFrame, coverageRatio);
                }
            } catch (e) { }
        }
    }

    /**
     * パスに収まらず文字が欠けるときだけ、文字サイズを縮めて収める（アーチ・円の保険）
     * 円のような閉じたパスも対象にし、すでに収まっている場合は何も変更しない。
     * 文字ごとのサイズ差を保つため、絶対値の代入ではなく比率で変倍する。
     * @param {TextFrame[]} pathTexts - 生成したパス上文字
     * @returns {void}
     */
    function preventOverset(pathTexts) {
        if (!pathTexts || pathTexts.length === 0) return;

        var shrinkOptions = {
            coarseRatio: 0.9,   /* 収まるまで一気に縮める比率 / coarse shrink ratio */
            fineRatio: 1.005,   /* 縮めすぎた分を戻す比率 / fine grow-back ratio */
            minFontSize: 0.5,
            maxCoarseIter: 40,
            maxFineIter: 25
        };

        for (var i = 0; i < pathTexts.length; i++) {
            var textFrame = pathTexts[i];
            if (!isEditablePathText(textFrame)) continue;

            try {
                if (textFrame.characters.length <= 0) continue;
                var lineAmount = getLineAmount(textFrame);
                /* 収まっているなら触らない（「しない」を選んだときの見た目を変えない）*/
                if (!isOverset(textFrame, lineAmount)) continue;

                var smallestSize = getFontSizeRange(textFrame).smallest;
                var appliedRatio = 1;
                var iterations = 0;

                /* 1. 収まるまで大きめの比率で縮める / shrink coarsely until it fits */
                while (isOverset(textFrame, lineAmount) && iterations < shrinkOptions.maxCoarseIter) {
                    if (smallestSize * appliedRatio * shrinkOptions.coarseRatio < shrinkOptions.minFontSize) break;
                    scaleFontSize(textFrame, shrinkOptions.coarseRatio);
                    appliedRatio *= shrinkOptions.coarseRatio;
                    iterations++;
                }

                /* 2. 縮めすぎた分を細かく戻し、あふれたら1段戻して確定 / grow back finely, then step back once */
                iterations = 0;
                while (!isOverset(textFrame, lineAmount) && iterations < shrinkOptions.maxFineIter) {
                    scaleFontSize(textFrame, shrinkOptions.fineRatio);
                    iterations++;
                }
                if (isOverset(textFrame, lineAmount)) scaleFontSize(textFrame, 1 / shrinkOptions.fineRatio);
            } catch (e) { }
        }
    }

    /**
     * 文字サイズでパスの端まで広げる（開いた／閉じたパスの両方）
     * あふれるまで拡大するところまでを担当し、収める側は preventOverset に任せる。
     * 絶対値を代入すると文字ごとのサイズ差が消えるため、比率で変倍する。
     * @param {TextFrame[]} pathTexts - 生成したパス上文字
     * @returns {boolean} 対象があれば true
     */
    function fitTextToPathByFontSize(pathTexts) {
        if (!pathTexts || pathTexts.length === 0) return false;

        var growOptions = {
            growRatio: 2,      /* あふれるまで一気に拡大する比率 / coarse grow ratio */
            maxGrowIter: 12,
            maxFontSize: 2000
        };

        for (var i = 0; i < pathTexts.length; i++) {
            var textFrame = pathTexts[i];
            if (!isEditablePathText(textFrame)) continue;

            try {
                if (textFrame.characters.length <= 0) continue;

                var lineAmount = getLineAmount(textFrame);
                /* すでにあふれているなら拡大は不要（preventOverset が収める）
                   Already overset: nothing to grow, preventOverset will pull it back */
                var iterations = 0;
                while (!isOverset(textFrame, lineAmount) && iterations < growOptions.maxGrowIter) {
                    if (getFontSizeRange(textFrame).largest * growOptions.growRatio > growOptions.maxFontSize) break;
                    scaleFontSize(textFrame, growOptions.growRatio);
                    iterations++;
                }
            } catch (e) { }
        }

        return true;
    }

    /**
     * 文字サイズを保ったまま、トラッキングだけでパスの端まで合わせる（開いた／閉じたパスの両方）
     * 大きな刻みであふれの境目を越え、細かい刻みで収まるいちばん広いトラッキングに落ち着かせる。
     * @param {TextFrame[]} pathTexts - 生成したパス上文字
     * @returns {boolean} 対象があれば true
     */
    function fitTextToPathByTracking(pathTexts) {
        if (!pathTexts || pathTexts.length === 0) return false;

        var trackingOptions = {
            coarseStep: 50,     /* 大きな刻み / tracking units per coarse step */
            fineStep: 1,        /* 細かい刻み / tracking units per fine step */
            minTracking: -1000, /* 加算の下限 / tightest allowed cumulative delta */
            maxTracking: 20000, /* 加算の上限 / loosest allowed cumulative delta */
            maxIter: 4000
        };

        /* 1つのパス上文字をトラッキングで合わせる / fit one path text by tracking */
        function fitByTracking(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);
                var appliedTracking = 0; /* これまでに加算した量 / cumulative tracking delta applied so far */
                var iterations;

                if (isOverset(textFrame, lineAmount)) {
                    /* 長すぎる：収まるまで大きな刻みで詰める / Too wide: tighten (coarse) until it fits */
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < trackingOptions.maxIter) {
                        if (appliedTracking - trackingOptions.coarseStep < trackingOptions.minTracking) break;
                        addTrackingToFrame(textFrame, -trackingOptions.coarseStep);
                        appliedTracking -= trackingOptions.coarseStep;
                        iterations++;
                    }
                    /* あふれるまで細かい刻みで広げ直す / Loosen back (fine) until it overflows again */
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < trackingOptions.maxIter) {
                        if (appliedTracking + trackingOptions.fineStep > trackingOptions.maxTracking) break;
                        addTrackingToFrame(textFrame, trackingOptions.fineStep);
                        appliedTracking += trackingOptions.fineStep;
                        iterations++;
                    }
                    /* 1刻み広げすぎたぶんを戻して収める / Stepped one fineStep too far: pull back once */
                    if (isOverset(textFrame, lineAmount)) {
                        addTrackingToFrame(textFrame, -trackingOptions.fineStep);
                        appliedTracking -= trackingOptions.fineStep;
                    }
                } else {
                    /* 余裕がある：あふれるまで大きな刻みで広げる / Fits with room: loosen (coarse) until it overflows */
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < trackingOptions.maxIter) {
                        if (appliedTracking + trackingOptions.coarseStep > trackingOptions.maxTracking) break;
                        addTrackingToFrame(textFrame, trackingOptions.coarseStep);
                        appliedTracking += trackingOptions.coarseStep;
                        iterations++;
                    }
                    /* 収まるまで細かい刻みで詰め直す / Tighten back (fine) until it fits */
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < trackingOptions.maxIter) {
                        if (appliedTracking - trackingOptions.fineStep < trackingOptions.minTracking) break;
                        addTrackingToFrame(textFrame, -trackingOptions.fineStep);
                        appliedTracking -= trackingOptions.fineStep;
                        iterations++;
                    }
                }
            } catch (e) { }
        }

        for (var i = 0; i < pathTexts.length; i++) {
            var textFrame = pathTexts[i];
            if (!isEditablePathText(textFrame)) continue;
            fitByTracking(textFrame);
        }

        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のテキストを集めてダイアログを開き、プレビューしながら変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }
        doc = app.activeDocument;
        var selectedItems = doc.selection;

        /* 開いたときの選択を控え、ダイアログ中に選択が変わってもプレビューを同じ対象から作る
           （文字の編集中は選択が配列でなく slice できない）
           Base selection snapshot for a stable preview (a text selection is not an array) */
        try { baseSelection = selectedItems.slice(0); } catch (e) { baseSelection = []; }

        targetTextFrames = collectSelectionTextFrames(selectedItems, TARGET_TEXT_OPTIONS);
        selectedPaths = collectSelectionPathItems(selectedItems, TARGET_PATH_OPTIONS);

        if (targetTextFrames.length === 0) {
            alert(getLabel('alert.noText'));
            return;
        }

        buildDialog();
        bindDialogEvents();
        initDialogState();

        /* 起動時に一度プレビュー / Auto-apply preview once on open */
        refreshPreview();

        prepareDialogWindow(dynamicTextDialog, SCRIPT_NAME);
        dynamicTextDialog.show();
    }

    main();

}());
