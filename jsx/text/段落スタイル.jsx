#target illustrator

(function () {

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
    var dialog = new Window("dialog", "段落スタイルの登録");
    dialog.alignChildren = "fill";
    dialog.margins = 15;

    // モード選択
    var modePanel = dialog.add("panel", undefined, "モード");
    modePanel.orientation = "row";
    modePanel.alignChildren = "left";
    modePanel.margins = [15, 20, 15, 15];
    var rbNew = modePanel.add("radiobutton", undefined, "新規");
    var rbOverwrite = modePanel.add("radiobutton", undefined, "上書き");

    // 新規スタイル名
    var newPanel = dialog.add("panel", undefined, "新規スタイル名");
    newPanel.alignChildren = "fill";
    newPanel.margins = [15, 20, 15, 15];
    var nameField = newPanel.add("edittext", undefined, defaultName);
    nameField.characters = 30;

    // 上書きするスタイル
    var overwritePanel = dialog.add("panel", undefined, "上書きするスタイル");
    overwritePanel.alignChildren = "fill";
    overwritePanel.margins = [15, 20, 15, 15];
    var styleList = overwritePanel.add("listbox", undefined, styleNames);
    styleList.preferredSize.height = 140;

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
    var cancelBtn = btnGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, "OK", { name: "ok" });

    var result = null;
    okBtn.onClick = function () {
        if (rbNew.value) {
            var name = nameField.text;
            if (!name || name.replace(/^\s+|\s+$/g, "") === "") {
                alert("スタイル名を入力してください。");
                return;
            }
            result = { mode: "new", styleName: name };
        } else {
            if (!styleList.selection) {
                alert("上書きするスタイルを選択してください。");
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
        alert("ドキュメントが開かれていません。");
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
        alert("テキスト（またはテキストオブジェクト）を選択してから実行してください。");
        return;
    }

    // 選択している段落に現在適用されている段落スタイル名を取得
    var currentStyleName = getAppliedParagraphStyleName(doc, targetParagraph);

    // ダイアログで新規／上書きを選択
    var styleNames = collectParagraphStyleNames(doc);
    var choice = showStyleDialog(styleNames, "新規段落スタイル", currentStyleName);
    if (!choice) {
        return;
    }

    var targetStyle;
    var isNewStyle = false;

    if (choice.mode === "new") {
        // 同名のスタイルが既に存在するか確認
        try {
            targetStyle = doc.paragraphStyles.getByName(choice.styleName);
            var overwrite = confirm("同名のスタイル「" + choice.styleName + "」が既に存在します。上書きしますか？");
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
            alert("新規段落スタイル「" + targetStyle.name + "」を作成し、適用しました。");
        } else {
            alert("既存の段落スタイル「" + targetStyle.name + "」を現在の書式で上書き更新しました。");
        }

    } catch (err) {
        alert("エラーが発生しました:\n" + err.message);
    }
}

updateOrCreateParagraphStyle();

})();
