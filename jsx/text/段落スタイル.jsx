#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの体裁を、段落スタイルとして登録します。既存のスタイルを選んで上書きすることもできます。

詳細は README を参照してください。

### Overview

Registers the formatting of the selected text as a paragraph style, or overwrites an existing style with it.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "段落スタイル";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

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
