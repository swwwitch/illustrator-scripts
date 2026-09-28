#target illustrator
#targetengine "段落スタイルEngine"
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
var SCRIPT_NAME     = "段落スタイル";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/段落スタイル.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/段落スタイル.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

/* 表示言語を判定 / Detect the UI language */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var uiLang = getCurrentLang();

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

/**
 * 表示言語のラベルを取得する
 * @param {string} key - LABELS のキー
 * @param {string} [name] - ラベル内の {name} に差し込む文字列
 * @returns {string} 表示言語のテキスト
 */
function getLabel(key, name) {
    var entry = LABELS[key];
    if (!entry) return key;
    var text = entry[uiLang] || entry.en || key;
    return (name == null) ? text : text.replace("{name}", name);
}

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

// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
// ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
//
// 【移植手順 / How to port】
// 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
//    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
// 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
//    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
// 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
//    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
//      prepareDialogWindow(mainDialog, SCRIPT_NAME);
//      var dialogResult = mainDialog.show();
//    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
//    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
// 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
// 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
//    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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
        if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
// ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

// 新規／上書きを選ぶダイアログを表示
// 戻り値: { mode: "new"|"overwrite", styleName: String } または null（キャンセル）
function showStyleDialog(styleNames, defaultName, currentStyleName) {
    var dialog = new Window("dialog", getLabel("dialogTitle"));
    dialog.alignChildren = "fill";
    dialog.margins = 15;

    // モード選択
    var modePanel = dialog.add("panel", undefined, getLabel("panelMode"));
    modePanel.orientation = "row";
    modePanel.alignChildren = "left";
    modePanel.margins = [15, 20, 15, 15];
    var rbNew = modePanel.add("radiobutton", undefined, getLabel("radioNew"));
    rbNew.helpTip = getLabel("tipModeNew");
    var rbOverwrite = modePanel.add("radiobutton", undefined, getLabel("radioOverwrite"));
    rbOverwrite.helpTip = getLabel("tipModeOverwrite");

    // 新規スタイル名
    var newPanel = dialog.add("panel", undefined, getLabel("panelNewName"));
    newPanel.alignChildren = "fill";
    newPanel.margins = [15, 20, 15, 15];
    var nameField = newPanel.add("edittext", undefined, defaultName);
    nameField.characters = 30;
    nameField.helpTip = getLabel("tipNewName");

    // 上書きするスタイル
    var overwritePanel = dialog.add("panel", undefined, getLabel("panelOverwrite"));
    overwritePanel.alignChildren = "fill";
    overwritePanel.margins = [15, 20, 15, 15];
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
    var btnGroup = dialog.add("group");
    btnGroup.alignment = "right";
    var cancelBtn = btnGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, getLabel("ok"), { name: "ok" });

    var result = null;
    okBtn.onClick = function () {
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
    cancelBtn.onClick = function () {
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
            var overwrite = confirm(getLabel("confirmOverwrite", choice.styleName));
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
            alert(getLabel("doneCreated", targetStyle.name));
        } else {
            alert(getLabel("doneOverwritten", targetStyle.name));
        }

    } catch (err) {
        alert(getLabel("alertError") + err.message);
    }
}

updateOrCreateParagraphStyle();

})();
