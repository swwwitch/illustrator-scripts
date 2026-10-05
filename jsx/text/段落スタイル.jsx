#target illustrator
#targetengine "ParagraphStyleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの体裁を、段落スタイルとして登録します。既存のスタイルを選んで上書きすることもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/段落スタイル.md

### Overview

Registers the formatting of the selected text as a paragraph style, or overwrites an existing style with it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/段落スタイル.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "段落スタイル";                       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-06";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/段落スタイル.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/段落スタイル.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle:       { ja: "段落スタイルの登録", en: "Register Paragraph Style" },
    panelMode:         { ja: "モード", en: "Mode" },
    radioNew:          { ja: "新規", en: "New" },
    radioOverwrite:    { ja: "上書き", en: "Overwrite" },
    panelNewName:      { ja: "新規スタイル名", en: "New Style Name" },
    panelOverwrite:    { ja: "上書きするスタイル", en: "Style to Overwrite" },
    cancel:            { ja: "キャンセル", en: "Cancel" },
    ok:                { ja: "OK", en: "OK" },
    defaultStyleName:  { ja: "新規段落スタイル", en: "New Paragraph Style" },
    tipModeNew:        { ja: "選択した段落の書式で、新しい段落スタイルを作ります。", en: "Creates a new paragraph style from the formatting of the selected paragraph." },
    tipModeOverwrite:  { ja: "選択した段落の書式で、既存の段落スタイルを上書きします。", en: "Overwrites an existing paragraph style with the formatting of the selected paragraph." },
    tipNewName:        { ja: "作る段落スタイルの名前です。", en: "Name of the paragraph style to create." },
    tipStyleList:      { ja: "上書きする段落スタイルを選びます。", en: "Choose the paragraph style to overwrite." },
    alertNoDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertSelectText:   { ja: "テキスト（またはテキストオブジェクト）を選択してから実行してください。", en: "Select text (or a text object) before running the script." },
    alertEnterName:    { ja: "スタイル名を入力してください。", en: "Enter a style name." },
    alertSelectStyle:  { ja: "上書きするスタイルを選択してください。", en: "Select the style to overwrite." },
    confirmOverwrite:  { ja: "同名のスタイル「{name}」が既に存在します。上書きしますか？", en: "A style named \u0022{name}\u0022 already exists. Overwrite it?" },
    doneCreated:       { ja: "新規段落スタイル「{name}」を作成し、適用しました。", en: "Created the paragraph style \u0022{name}\u0022 and applied it." },
    doneOverwritten:   { ja: "既存の段落スタイル「{name}」を現在の書式で上書き更新しました。", en: "Updated the existing paragraph style \u0022{name}\u0022 with the current formatting." },
    alertError:        { ja: "エラーが発生しました：\n", en: "An error occurred:\n" }
};

// ドキュメント内の段落スタイル名を取得（既定スタイル = index 0 は除外）
function collectParagraphStyleNames(doc) {
    var names = [];
    // Illustrator では doc.paragraphStyles[0] が常に既定スタイル（標準段落スタイル / Normal Paragraph Style）
    for (var i = 1; i < doc.paragraphStyles.length; i++) {
        names.push(doc.paragraphStyles[i].name);
    }
    return names;
}

// 段落に適用されている段落スタイル名を取得（既定スタイルなら null）
// Illustrator には paragraph.paragraphStyle が無いため paragraphStyles[0] を参照する
function getAppliedParagraphStyleName(doc, paragraph) {
    try {
        var applied = paragraph.paragraphStyles[0];
        if (!applied) {
            return null;
        }
        var defaultStyleName = doc.paragraphStyles[0].name;
        if (applied.name === defaultStyleName) {
            return null;
        }
        return applied.name;
    } catch (e) {
        return null;
    }
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

// 新規／上書きを選ぶダイアログを表示
// 戻り値: { mode: "new"|"overwrite", styleName: String } または null（キャンセル）
function showStyleDialog(styleNames, defaultName, currentStyleName) {
    var dialog = new Window("dialog", getLabel("dialogTitle"));
    setupWindow(dialog);

    // モード選択
    var modePanel = dialog.add("panel", undefined, getLabel("panelMode"));
    setupPanel(modePanel, 6);
    modePanel.orientation = "row";
    modePanel.alignChildren = ["left", "center"];
    var rbNew = modePanel.add("radiobutton", undefined, getLabel("radioNew"));
    rbNew.helpTip = getLabel("tipModeNew");
    var rbOverwrite = modePanel.add("radiobutton", undefined, getLabel("radioOverwrite"));
    rbOverwrite.helpTip = getLabel("tipModeOverwrite");

    // 新規スタイル名
    var newPanel = dialog.add("panel", undefined, getLabel("panelNewName"));
    setupPanel(newPanel);
    var nameField = newPanel.add("edittext", undefined, defaultName);
    nameField.characters = 30;
    nameField.helpTip = getLabel("tipNewName");

    // 上書きするスタイル
    var overwritePanel = dialog.add("panel", undefined, getLabel("panelOverwrite"));
    setupPanel(overwritePanel);
    var styleList = overwritePanel.add("listbox", undefined, styleNames);
    styleList.preferredSize.height = 140;
    styleList.helpTip = getLabel("tipStyleList");

    // 状態切り替え
    function updateState() {
        var isNew = rbNew.value;
        newPanel.enabled = isNew;
        nameField.enabled = isNew;
        overwritePanel.enabled = !isNew;
        styleList.enabled = !isNew;
    }
    rbNew.onClick = updateState;
    rbOverwrite.onClick = updateState;

    // 既存スタイルがなければ上書きは選べない
    var hasStyles = styleNames.length > 0;
    if (!hasStyles) {
        rbOverwrite.enabled = false;
    }

    // 初期選択：現在のスタイルがあれば上書きモードで選択、なければ新規
    var startAsOverwrite = false;
    if (hasStyles && currentStyleName) {
        for (var i = 0; i < styleNames.length; i++) {
            if (styleNames[i] === currentStyleName) {
                styleList.selection = i;
                startAsOverwrite = true;
                break;
            }
        }
    }
    if (startAsOverwrite) {
        rbOverwrite.value = true;
    } else {
        rbNew.value = true;
        if (hasStyles) {
            styleList.selection = 0;
        }
    }
    updateState();

    // ボタン（Mac 規約：Cancel → OK）
    var buttonRow = addButtonRow(dialog);
    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
    var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("ok"), { name: "ok" });

    var result = null;
    btnOK.onClick = function () {
        if (rbNew.value) {
            var name = nameField.text;
            if (!name || name.replace(/^\s+|\s+$/g, "") === "") {
                alert(getLabel("alertEnterName"));
                return;
            }
            result = { mode: "new", styleName: name };
        } else {
            if (!styleList.selection) {
                alert(getLabel("alertSelectStyle"));
                return;
            }
            result = { mode: "overwrite", styleName: styleList.selection.text };
        }
        dialog.close();
    };
    btnCancel.onClick = function () {
        dialog.close();
    };

    prepareDialogWindow(dialog, SCRIPT_NAME);
    dialog.show();
    return result;
}

function updateOrCreateParagraphStyle() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert(getLabel("alertNoDocument"));
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    // テキストの一部（TextRange）またはテキストオブジェクト（TextFrame）を許可
    var targetParagraph = null;
    if (selection.constructor.name === "TextRange") {
        // 文字ツールでテキストの一部を選択している場合
        targetParagraph = selection.paragraphs[0];
    } else if (selection.length > 0 && selection[0].constructor.name === "TextFrame") {
        // 選択ツールでテキストオブジェクトを選択している場合
        targetParagraph = selection[0].paragraphs[0];
    }

    if (!targetParagraph) {
        alert(getLabel("alertSelectText"));
        return;
    }

    // 選択している段落に現在適用されている段落スタイル名を取得
    var currentStyleName = getAppliedParagraphStyleName(doc, targetParagraph);

    // ダイアログで新規／上書きを選択
    var styleNames = collectParagraphStyleNames(doc);
    var choice = showStyleDialog(styleNames, getLabel("defaultStyleName"), currentStyleName);
    if (!choice) {
        return;
    }

    var targetStyle;
    var isNewStyle = false;

    if (choice.mode === "new") {
        // 同名のスタイルが既に存在するか確認
        try {
            targetStyle = doc.paragraphStyles.getByName(choice.styleName);
            var overwrite = confirm(getLabel("confirmOverwrite", { name: choice.styleName }), true);
            if (!overwrite) {
                return;
            }
        } catch (e) {
            // 存在しない場合は新規作成
            targetStyle = doc.paragraphStyles.add(choice.styleName);
            isNewStyle = true;
        }
    } else {
        // 選択した既存スタイルを上書き
        targetStyle = doc.paragraphStyles.getByName(choice.styleName);
    }

    // 選択された段落の属性をスタイルにコピー（上書き）
    try {
        var charAttr = targetParagraph.characterAttributes;
        var paraAttr = targetParagraph.paragraphAttributes;

        // 文字属性のコピー
        targetStyle.characterAttributes.textFont = charAttr.textFont;
        targetStyle.characterAttributes.size = charAttr.size;
        targetStyle.characterAttributes.leading = charAttr.leading;
        targetStyle.characterAttributes.tracking = charAttr.tracking;
        targetStyle.characterAttributes.fillColor = charAttr.fillColor;
        targetStyle.characterAttributes.strokeColor = new NoColor();

        // 段落属性のコピー
        targetStyle.paragraphAttributes.justification = paraAttr.justification;
        targetStyle.paragraphAttributes.firstLineIndent = paraAttr.firstLineIndent;
        targetStyle.paragraphAttributes.leftIndent = paraAttr.leftIndent;
        targetStyle.paragraphAttributes.rightIndent = paraAttr.rightIndent;
        targetStyle.paragraphAttributes.spaceBefore = paraAttr.spaceBefore;
        targetStyle.paragraphAttributes.spaceAfter = paraAttr.spaceAfter;

        // 選択していたテキストに、スタイルを再適用してオーバーライドをクリアする
        targetStyle.applyTo(targetParagraph, true);

        if (isNewStyle) {
            alert(getLabel("doneCreated", { name: targetStyle.name }));
        } else {
            alert(getLabel("doneOverwritten", { name: targetStyle.name }));
        }

    } catch (err) {
        alert(getLabel("alertError") + err.message);
    }
}

updateOrCreateParagraphStyle();

})();
