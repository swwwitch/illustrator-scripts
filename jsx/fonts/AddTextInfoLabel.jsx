#target illustrator
#targetengine "AddTextInfoLabelEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストオブジェクトの下または右に、フォント情報を表示するラベルを作成します。
詳細表示（エリア内文字）と簡易表示（ポイント文字）を切り替えられ、表示項目は個別に指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTextInfoLabel.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n607ef418877f

### Overview

Adds a label showing the font information below or beside the selected text objects.
You can switch between a detailed layout (area text) and a compact one (point text), and choose which items appear.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTextInfoLabel.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddTextInfoLabel";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.11";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTextInfoLabel.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTextInfoLabel.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n607ef418877f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 情報ラベルを入れるレイヤー名 / Layer that receives the info labels */
    var INFO_LAYER_NAME = "フォント情報";

    /* 情報ラベルの書式 / Info label text format */
    var LABEL_FONT_NAME = "HiraginoSans-W3";
    var LABEL_FONT_SIZE = 10;
    var LABEL_LEADING = 16;

    /* 一度に処理する選択数の上限（これ以上は何もしない） / Selections at or above this count are ignored */
    var MAX_SELECTION_COUNT = 1000;

    // =========================================
    // レイアウト / Layout

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

    // =========================================

    /* 一括切り替えボタンの大きさ / Bounds of the bulk toggle buttons */
    var TOGGLE_BUTTON_BOUNDS = [0, 0, 80, 24];

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "テキスト情報を追加", en: "Add Text Info Label" }
        },
        panel: {
            position: { ja: "位置", en: "Position" },
            mode: { ja: "表示形式", en: "Format" },
            info: { ja: "表示項目（詳細表示時のみ有効）", en: "Items (used by the detailed format only)" }
        },
        radio: {
            posBottom: { ja: "下", en: "Below" },
            posRight: { ja: "右", en: "Right" },
            modeCompact: { ja: "簡易版", en: "Compact" },
            modeFull: { ja: "詳細", en: "Detailed" }
        },
        checkbox: {
            fontName: { ja: "フォント名", en: "Font name" },
            postScript: { ja: "PSフォント名", en: "PostScript name" },
            fontStyle: { ja: "スタイル（ウェイト）", en: "Style (weight)" },
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            leading: { ja: "行送り", en: "Leading" },
            kerning: { ja: "カーニング", en: "Kerning" },
            proportional: { ja: "プロポーショナルメトリクス", en: "Proportional metrics" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            leadingPercent: { ja: "行送り（%）", en: "Leading (%)" }
        },
        tooltip: {
            posBottom: { ja: "テキストの下に情報ラベルを置きます。", en: "Places the info label below the text." },
            posRight: { ja: "テキストの右に情報ラベルを置きます。", en: "Places the info label to the right of the text." },
            modeCompact: { ja: "フォント名とサイズだけの短い表記にします。", en: "Writes a short label with just the font name and size." },
            modeFull: { ja: "下の［表示項目］で選んだ内容をすべて書き出します。", en: "Writes every item ticked under Items below." },
            info: { ja: "詳細表示のときに、ラベルへ書き出す項目を選びます。", en: "Picks which items go into the label when the detailed format is used." },
            kerning: {
                ja: "カーニングの方式（メトリクス／和文等幅／オプティカル／なし）を書き出します。",
                en: "Writes the kerning method (Metrics, Metrics - Roman Only, Optical, or none)."
            },
            leadingPercent: { ja: "行送りをフォントサイズに対する割合で書き出します。", en: "Writes the leading as a percentage of the font size." },
            minimalSet: {
                ja: "フォント名・スタイル・フォントサイズ・行送りだけをオンにします。",
                en: "Turns on only the font name, style, font size, and leading."
            }
        }
    };

    // =========================================
    // 表示項目 / Info items
    // =========================================

    /* 詳細表示の項目（チェックボックスの列・初期値・［最小セット］での値）
       Detailed-format items: checkbox column, initial value, value set by the minimal preset */
    var INFO_ITEMS = [
        { key: "fontName",       column: 0, initial: true,  minimal: true },
        { key: "postScript",     column: 0, initial: false, minimal: false },
        { key: "fontStyle",      column: 0, initial: true,  minimal: true },
        { key: "fontSize",       column: 0, initial: true,  minimal: true },
        { key: "leading",        column: 0, initial: true,  minimal: true },
        { key: "kerning",        column: 1, initial: true,  minimal: false, tooltip: "tooltip.kerning" },
        { key: "proportional",   column: 1, initial: true,  minimal: false },
        { key: "tracking",       column: 1, initial: false, minimal: false },
        { key: "tsume",          column: 1, initial: false, minimal: false },
        { key: "leadingPercent", column: 1, initial: false, minimal: false, tooltip: "tooltip.leadingPercent" }
    ];

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

    /* ラベル枠の余白は常に mm で指定する / label frame padding is always given in mm */
    var POINTS_PER_MM = UNITS[1].pointsPerUnit;

    // =========================================
    // テキスト情報の読み取り / Reading text attributes
    // =========================================

    /* 段落の区切り（CR） / Paragraph separator (CR) */
    var PARAGRAPH_BREAK = String.fromCharCode(13);

    /**
     * 読み取り関数を呼び、例外が出たら「不明」を返す
     * @param {Function} readValue - (textFrame, displayMode) を受け取る読み取り関数
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string} [displayMode] - "compact" または "full"
     * @returns {*} 読み取った値、または "不明"
     */
    function readOrUnknown(readValue, textFrame, displayMode) {
        /* 文字属性の読み取りは失敗しうる / reading character attributes may throw */
        try {
            return readValue(textFrame, displayMode);
        } catch (e) {
            return "不明";
        }
    }

    /**
     * フォントサイズを小数1桁の単位付き文字列で返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} フォントサイズ
     */
    function getFontSize(textFrame) {
        /* size は常に pt なので、文字サイズの単位に換算してから表示する / size is always in pt; convert to the type-size unit */
        var sizeUnit = getUnitInfo("text/units");
        var rawSize = textFrame.textRange.characterAttributes.size / sizeUnit.pointsPerUnit;
        var roundedSize = Math.round(rawSize * 10) / 10;
        return roundedSize + " " + sizeUnit.label;
    }

    /**
     * 小数が3桁以上なら2桁に丸め、それ以外はそのまま文字列にする
     * @param {number} value - 数値
     * @returns {string} 整形した数値
     */
    function formatUpToTwoDecimals(value) {
        var valueText = String(value);
        if (valueText.indexOf(".") === -1) {
            return valueText;
        }
        if (valueText.split(".")[1].length <= 2) {
            return valueText;
        }
        return String(Math.round(value * 100) / 100);
    }

    /**
     * 行送りを単位付き文字列で返す（詳細表示の自動行送りは「自動（…）」）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string} displayMode - "compact" または "full"
     * @returns {string} 行送り
     */
    function getLeading(textFrame, displayMode) {
        var textAttributes = textFrame.textRange.characterAttributes;
        /* leading は常に pt なので、行送りの単位に換算してから表示する / leading is always in pt; convert to the leading unit */
        var leadingUnit = getUnitInfo("text/asianunits");
        var leadingText = formatUpToTwoDecimals(textAttributes.leading / leadingUnit.pointsPerUnit) + " " + leadingUnit.label;

        if (textAttributes.autoLeading && displayMode !== "compact") {
            return "自動（" + leadingText + "）";
        }
        return leadingText;
    }

    /**
     * 行送りのフォントサイズに対する割合を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} 「120 %」のような割合（数値が取れなければ「不明」）
     */
    function getLeadingPercentage(textFrame) {
        var textAttributes = textFrame.textRange.characterAttributes;
        var fontSize = textAttributes.size;
        var leading = textAttributes.leading;

        /* 数値でない場合は不明 / Unknown unless both are numbers */
        if (typeof fontSize !== "number" || typeof leading !== "number" || fontSize === 0) {
            return "不明";
        }
        return roundToTwoDecimalFixed((leading / fontSize) * 100) + " %";
    }

    /**
     * 小数2桁に丸め、「.00」なら整数にする
     * @param {number} value - 数値
     * @returns {string} 整形した数値
     */
    function roundToTwoDecimalFixed(value) {
        var rounded = Math.round(value * 100) / 100;
        var fixedText = rounded.toFixed(2);
        if (fixedText.match(/\.00$/)) return String(parseInt(rounded, 10));
        return fixedText;
    }

    /**
     * カーニングの方式を日本語名で返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} カーニングの方式
     */
    function getKerningMethodText(textFrame) {
        switch (textFrame.textRange.characterAttributes.kerningMethod) {
            case AutoKernType.AUTO: return "メトリクス";
            case AutoKernType.METRICSROMANONLY: return "和文等幅";
            case AutoKernType.OPTICAL: return "オプティカル";
            default: return "なし";
        }
    }

    /**
     * トラッキングを返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number} トラッキング
     */
    function getTracking(textFrame) {
        return textFrame.textRange.characterAttributes.tracking;
    }

    /**
     * プロポーショナルメトリクスの状態を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} "ON" または "OFF"
     */
    function getProportionalMetrics(textFrame) {
        return textFrame.textRange.characterAttributes.proportionalMetrics ? "ON" : "OFF";
    }

    /**
     * 文字ツメを小数1桁までの文字列で返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {string} 文字ツメ（数値が取れなければ「なし」）
     */
    function getTsume(textFrame) {
        var tsume = textFrame.textRange.characterAttributes.Tsume;
        if (typeof tsume !== "number" || isNaN(tsume)) return "なし";

        var rounded = Math.round(tsume * 10) / 10;
        return (rounded % 1 === 0) ? String(rounded.toFixed(0)) : String(rounded.toFixed(1));
    }

    /**
     * 最初に見つかったフォントを返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {TextFont|null} フォント（見つからなければ null）
     */
    function getFirstAvailableFont(textFrame) {
        var characters = textFrame.textRange.characters;
        for (var i = 0; i < characters.length; i++) {
            var textFont = characters[i].characterAttributes.textFont;
            if (textFont) return textFont;
        }
        return null;
    }

    /**
     * ラベルに書き出すテキスト情報をまとめて読み取る
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {TextFont} sourceFont - 対象のフォント
     * @param {string} displayMode - "compact" または "full"
     * @returns {Object} テキスト情報
     */
    function readTextInfo(textFrame, sourceFont, displayMode) {
        return {
            fontSize: readOrUnknown(getFontSize, textFrame),
            leading: readOrUnknown(getLeading, textFrame, displayMode),
            leadingPercent: readOrUnknown(getLeadingPercentage, textFrame),
            kerning: readOrUnknown(getKerningMethodText, textFrame),
            tracking: readOrUnknown(getTracking, textFrame),
            proportional: readOrUnknown(getProportionalMetrics, textFrame),
            tsume: readOrUnknown(getTsume, textFrame),
            fontFamily: sourceFont.family,
            fontStyle: sourceFont.style,
            postScriptName: sourceFont.name
        };
    }

    // =========================================
    // 情報ラベルの作成 / Info label creation
    // =========================================

    /**
     * 簡易表示の3行のテキストを組み立てる
     * @param {Object} textInfo - readTextInfo() の結果
     * @returns {string} ラベルのテキスト
     */
    function buildCompactText(textInfo) {
        var fontLine = textInfo.fontFamily + " " + textInfo.fontStyle + "、" + textInfo.fontSize + " ↓" + textInfo.leading;
        var kerningLine = textInfo.kerning + "、プロポーショナルメトリクス：" + textInfo.proportional;
        var spacingLine = "トラッキング：" + textInfo.tracking + "、文字ツメ：" + textInfo.tsume + " %";
        return fontLine + PARAGRAPH_BREAK + kerningLine + PARAGRAPH_BREAK + spacingLine;
    }

    /**
     * 詳細表示のテキストを、選ばれた項目だけで組み立てる
     * @param {Object} textInfo - readTextInfo() の結果
     * @param {Object} includeItems - 項目キーごとの書き出す／書き出さない
     * @returns {string} ラベルのテキスト
     */
    function buildDetailedText(textInfo, includeItems) {
        /* 書き出す順 / Output order */
        var detailLines = [
            ["fontName", "・フォント名\t" + textInfo.fontFamily],
            ["postScript", "・PSフォント名\t" + textInfo.postScriptName],
            ["fontStyle", "・スタイル（ウェイト）\t" + textInfo.fontStyle],
            ["fontSize", "・フォントサイズ\t" + textInfo.fontSize],
            ["leading", "・行送り\t" + textInfo.leading],
            ["leadingPercent", "・行送り（%）\t" + textInfo.leadingPercent],
            ["kerning", "・カーニング\t" + textInfo.kerning],
            ["proportional", "・プロポーショナルメトリクス\t" + textInfo.proportional],
            ["tracking", "・トラッキング\t" + textInfo.tracking],
            ["tsume", "・文字ツメ\t" + textInfo.tsume + " %"]
        ];
        var infoLines = [];
        for (var i = 0; i < detailLines.length; i++) {
            if (includeItems[detailLines[i][0]]) infoLines.push(detailLines[i][1]);
        }
        return infoLines.join(PARAGRAPH_BREAK);
    }

    /**
     * 詳細表示のラベルをエリア内文字で作る（右揃えのタブとリーダー付き）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} textContent - ラベルのテキスト
     * @param {number[]} sourceBounds - 元のテキストの geometricBounds
     * @param {string} position - "bottom" または "right"
     * @returns {TextFrame} 作ったラベル
     */
    function createDetailedInfoFrame(doc, textContent, sourceBounds, position) {
        var lineCount = textContent.split(PARAGRAPH_BREAK).length;
        var lineHeight = LABEL_FONT_SIZE * 1.4;
        var frameHeight = lineHeight * (lineCount + 2) + 4 * POINTS_PER_MM;
        var frameWidth = 300;

        var frameLeft = (position === "right") ? sourceBounds[2] + 10 : sourceBounds[0];
        var frameTop = (position === "right") ? sourceBounds[3] + frameHeight : sourceBounds[3] - LABEL_FONT_SIZE;

        var framePath = doc.pathItems.rectangle(frameTop, frameLeft, frameWidth, frameHeight);
        var infoFrame = doc.textFrames.areaText(framePath);
        infoFrame.contents = textContent;
        infoFrame.spacing = 2 * POINTS_PER_MM;

        /* タブストップ（右揃え、位置400pt、リーダー…） / Right tab stop at 400pt with a "…" leader */
        var tabStop = new TabStopInfo();
        tabStop.position = 400;
        tabStop.alignment = TabStopAlignment.Right;
        tabStop.leader = "…";
        for (var i = 0; i < infoFrame.paragraphs.length; i++) {
            infoFrame.paragraphs[i].tabStops = [tabStop];
        }
        return infoFrame;
    }

    /**
     * 簡易表示のラベルをポイント文字で作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} textContent - ラベルのテキスト
     * @param {number[]} sourceBounds - 元のテキストの geometricBounds
     * @param {string} position - "bottom" または "right"
     * @returns {TextFrame} 作ったラベル
     */
    function createCompactInfoFrame(doc, textContent, sourceBounds, position) {
        var labelLeft = (position === "right") ? sourceBounds[2] + 20 : sourceBounds[0];
        var labelTop = (position === "right") ? sourceBounds[1] : sourceBounds[3] - LABEL_FONT_SIZE;

        var infoFrame = doc.textFrames.add();
        infoFrame.contents = textContent;
        infoFrame.position = [labelLeft, labelTop];
        /* 見た目の上端をそろえる / Align the visible top edge */
        var infoBounds = infoFrame.visibleBounds;
        infoFrame.translate(0, labelTop - infoBounds[1]);
        return infoFrame;
    }

    /**
     * ラベルに文字サイズ・行送り・揃え・フォントを設定する
     * @param {TextFrame} infoFrame - 作ったラベル
     * @param {string} displayMode - "compact" または "full"
     * @returns {void}
     */
    function applyInfoFrameStyle(infoFrame, displayMode) {
        var infoAttributes = infoFrame.textRange.characterAttributes;
        infoAttributes.autoLeading = false;
        infoAttributes.leading = LABEL_LEADING;
        infoFrame.textRange.characterAttributes.size = LABEL_FONT_SIZE;
        infoFrame.textRange.paragraphAttributes.justification =
            (displayMode === "full") ? Justification.RIGHT : Justification.LEFT;

        /* フォントが無ければ既定のまま / keep the default font when it is missing */
        try {
            infoFrame.textRange.characterAttributes.textFont = textFonts.getByName(LABEL_FONT_NAME);
        } catch (e) {}
    }

    /**
     * 名前でレイヤーを探し、無ければ作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} レイヤー
     */
    function getOrCreateLayer(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return doc.layers[i];
        }
        var newLayer = doc.layers.add();
        newLayer.name = layerName;
        return newLayer;
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタン2つを横に並べたパネルを追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} panelLabelPath - パネルタイトルの LABELS パス
     * @param {string[]} radioKeys - LABELS.radio / LABELS.tooltip のキー
     * @returns {Object} キーごとのラジオボタン
     */
    function addRadioPanel(parentGroup, panelLabelPath, radioKeys) {
        var radioPanel = parentGroup.add("panel", undefined, getLabel(panelLabelPath));
        setupPanel(radioPanel, 6);
        radioPanel.orientation = "row";
        radioPanel.alignChildren = ["left", "center"];
        var radios = {};
        for (var i = 0; i < radioKeys.length; i++) {
            var radio = radioPanel.add("radiobutton", undefined, getLabel("radio." + radioKeys[i]));
            radio.helpTip = getLabel("tooltip." + radioKeys[i]);
            radios[radioKeys[i]] = radio;
        }
        return radios;
    }

    /**
     * 表示項目パネル（2列のチェックボックスと一括切り替えボタン）を追加する
     * @param {Window} optionDialog - ダイアログ
     * @returns {{panel: Panel, checkboxes: Object}} パネルと項目キーごとのチェックボックス
     */
    function addInfoItemsPanel(optionDialog) {
        var infoPanel = optionDialog.add("panel", undefined, getLabel("panel.info"));
        infoPanel.helpTip = getLabel("tooltip.info");
        setupPanel(infoPanel, 6);
        infoPanel.alignChildren = ["left", "top"];

        var columnsGroup = infoPanel.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        var columns = [];
        for (var c = 0; c < 2; c++) {
            columns[c] = columnsGroup.add("group");
            columns[c].orientation = "column";
            columns[c].alignChildren = "left";
        }

        var checkboxes = {};
        for (var i = 0; i < INFO_ITEMS.length; i++) {
            var infoItem = INFO_ITEMS[i];
            var itemCheckbox = columns[infoItem.column].add("checkbox", undefined, getLabel("checkbox." + infoItem.key));
            if (infoItem.tooltip) itemCheckbox.helpTip = getLabel(infoItem.tooltip);
            itemCheckbox.value = infoItem.initial;
            checkboxes[infoItem.key] = itemCheckbox;
        }

        /**
         * 全項目のチェックを、項目ごとの値で設定する
         * @param {Function} getValue - 項目定義を受け取って値を返す関数
         * @returns {void}
         */
        function setAllItems(getValue) {
            for (var i = 0; i < INFO_ITEMS.length; i++) {
                checkboxes[INFO_ITEMS[i].key].value = getValue(INFO_ITEMS[i]);
            }
        }

        var toggleButtonGroup = infoPanel.add("group");
        toggleButtonGroup.orientation = "row";
        toggleButtonGroup.alignment = "left";
        var btnAllOn = toggleButtonGroup.add("button", TOGGLE_BUTTON_BOUNDS, "すべてON");
        var btnAllOff = toggleButtonGroup.add("button", TOGGLE_BUTTON_BOUNDS, "すべてOFF");
        var btnMinimal = toggleButtonGroup.add("button", TOGGLE_BUTTON_BOUNDS, "最小セット");
        btnMinimal.helpTip = getLabel("tooltip.minimalSet");

        btnAllOn.onClick = function () {
            setAllItems(function () { return true; });
        };
        btnAllOff.onClick = function () {
            setAllItems(function () { return false; });
        };
        btnMinimal.onClick = function () {
            setAllItems(function (infoItem) { return infoItem.minimal; });
        };

        return { panel: infoPanel, checkboxes: checkboxes };
    }

    /**
     * 設定ダイアログを表示し、選ばれた設定を返す
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showOptionDialog() {
        var optionDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(optionDialog);

        var topGroup = optionDialog.add("group");
        topGroup.orientation = "row";
        topGroup.alignChildren = ["left", "top"];
        topGroup.spacing = COLUMN_SPACING;

        var positionRadios = addRadioPanel(topGroup, "panel.position", ["posBottom", "posRight"]);
        positionRadios.posRight.value = true;

        var modeRadios = addRadioPanel(topGroup, "panel.mode", ["modeCompact", "modeFull"]);
        modeRadios.modeCompact.value = true;

        var infoItemsUI = addInfoItemsPanel(optionDialog);

        /* 表示項目は詳細表示のときだけ有効 / Items apply to the detailed format only */
        function updateInfoPanelEnabled() {
            infoItemsUI.panel.enabled = modeRadios.modeFull.value;
        }
        modeRadios.modeFull.onClick = updateInfoPanelEnabled;
        modeRadios.modeCompact.onClick = updateInfoPanelEnabled;
        updateInfoPanelEnabled();

        var buttonRow = addButtonRow(optionDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, "キャンセル", { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(optionDialog, SCRIPT_NAME);
        if (optionDialog.show() !== 1) return null;

        var includeItems = {};
        for (var i = 0; i < INFO_ITEMS.length; i++) {
            includeItems[INFO_ITEMS[i].key] = infoItemsUI.checkboxes[INFO_ITEMS[i].key].value;
        }
        return {
            position: positionRadios.posBottom.value ? "bottom" : "right",
            displayMode: modeRadios.modeFull.value ? "full" : "compact",
            includeItems: includeItems
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 文字を編集中で、そのストーリーが1つのテキストだけなら、テキストオブジェクトの選択に切り替える
     * @returns {void}
     */
    function selectFrameOfEditedText() {
        if (app.selection.constructor.name !== "TextRange") return;

        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length === 1) {
            app.executeMenuCommand("deselectall");
            app.selection = [textFramesInStory[0]];
            /* 選択ツールに切り替え（失敗しても続行） / Switch to the Selection tool; ignore failures */
            try { app.selectTool("Adobe Select Tool"); } catch (e) {}
        }
    }

    /**
     * 選択したテキストごとにフォント情報のラベルを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        selectFrameOfEditedText();

        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0 || selectedItems.length >= MAX_SELECTION_COUNT) return;

        var labelOptions = showOptionDialog();
        if (!labelOptions) return;

        var infoLayer = getOrCreateLayer(doc, INFO_LAYER_NAME);
        var generatedItems = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var sourceFrame = selectedItems[i];
            if (sourceFrame.typename !== "TextFrame") continue;

            var sourceFont = getFirstAvailableFont(sourceFrame);
            if (!sourceFont) continue;

            var textInfo = readTextInfo(sourceFrame, sourceFont, labelOptions.displayMode);
            var sourceBounds = sourceFrame.geometricBounds;
            var infoFrame;

            if (labelOptions.displayMode === "compact") {
                infoFrame = createCompactInfoFrame(doc, buildCompactText(textInfo), sourceBounds, labelOptions.position);
            } else {
                infoFrame = createDetailedInfoFrame(doc, buildDetailedText(textInfo, labelOptions.includeItems), sourceBounds, labelOptions.position);
            }
            applyInfoFrameStyle(infoFrame, labelOptions.displayMode);

            infoFrame.move(infoLayer, ElementPlacement.PLACEATBEGINNING);
            generatedItems.push(infoFrame);
            sourceFrame.selected = false;
        }

        if (generatedItems.length > 0) {
            app.selection = generatedItems;
        }
        app.redraw(); /* 画面再描画 / Redraw the screen */
    }

    main();

})();
