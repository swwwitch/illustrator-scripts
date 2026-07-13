#target illustrator

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 禁則プリセット（表示順）。labelText=UI表示名、kinsokuName=paragraphAttributes.kinsoku に渡す値 */
    var kinsokuPresets = [
        { kinsokuName: "None",    labelText: "なし" },
        { kinsokuName: "Hard",    labelText: "強い禁則" },
        { kinsokuName: "Soft",    labelText: "弱い禁則" },
        { kinsokuName: "Soft_v2", labelText: "弱い禁則 v2" }
    ];

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {
        alert("テキストオブジェクトを選択してください。");
        return;
    }

    /* テキストフレームに禁則を適用（「なし」= "None" は scripting では設定不可で例外になるため握りつぶす）*/
    function applyKinsokuToTextFrame(textFrame, kinsokuName) {
        try {
            textFrame.textRange.paragraphAttributes.kinsoku = kinsokuName;
        } catch (e) {
            // 「なし」は scripting から設定できない（アクション再生が必要）
        }
    }

    /* 種類で振り分け（TextFrame は直接、GroupItem は再帰）*/
    function applyKinsokuToItem(item, kinsokuName) {
        if (item.typename === "TextFrame") {
            applyKinsokuToTextFrame(item, kinsokuName);
        } else if (item.typename === "GroupItem") {
            applyKinsokuToGroup(item, kinsokuName);
        }
    }

    /* グループ内の各アイテムを再帰処理 */
    function applyKinsokuToGroup(groupItem, kinsokuName) {
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            applyKinsokuToItem(groupItem.pageItems[i], kinsokuName);
        }
    }

    /* 選択全体に適用して再描画 */
    function applyKinsokuToSelection(kinsokuName) {
        for (var i = 0; i < doc.selection.length; i++) {
            applyKinsokuToItem(doc.selection[i], kinsokuName);
        }
        app.redraw();
    }

    /* 現在の禁則を取得（禁則「なし」の段落は getter が Error 9563 を投げるため "" を返す）*/
    function getKinsoku(textFrame) {
        try {
            return textFrame.textRange.paragraphAttributes.kinsoku;
        } catch (e) {
            return "";
        }
    }

    /* 選択内から最初のテキストフレームを探す（初期選択の判定用）*/
    function findFirstTextFrame(item) {
        if (item.typename === "TextFrame") {
            return item;
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                var found = findFirstTextFrame(item.pageItems[i]);
                if (found !== null) {
                    return found;
                }
            }
        }
        return null;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    var kinsokuDialog = new Window("dialog", "禁則設定");
    kinsokuDialog.orientation = "column";
    kinsokuDialog.alignChildren = "fill";

    /* 禁則を選択する panel（ラジオをまとめる）*/
    var presetPanel = kinsokuDialog.add("panel", undefined, "禁則を選択");
    presetPanel.orientation = "column";
    presetPanel.alignChildren = "left";
    presetPanel.alignment = "fill";
    presetPanel.margins = [16, 20, 16, 12];
    presetPanel.spacing = 6;

    /* プリセットごとにラジオボタンを生成（選択で即適用）*/
    var radioButtons = [];
    for (var i = 0; i < kinsokuPresets.length; i++) {
        var radio = presetPanel.add("radiobutton", undefined, kinsokuPresets[i].labelText);
        radio.kinsokuName = kinsokuPresets[i].kinsokuName;
        radio.onClick = function () {
            applyKinsokuToSelection(this.kinsokuName);
        };
        radioButtons.push(radio);
    }

    /* 現在の禁則値を読んで初期選択に反映（一致しなければ先頭＝「なし」）*/
    var firstTextFrame = null;
    for (var i = 0; i < doc.selection.length; i++) {
        firstTextFrame = findFirstTextFrame(doc.selection[i]);
        if (firstTextFrame !== null) {
            break;
        }
    }

    var currentKinsoku = (firstTextFrame !== null) ? getKinsoku(firstTextFrame) : "";
    var selectedRadio = radioButtons[0];
    for (var i = 0; i < radioButtons.length; i++) {
        if (radioButtons[i].kinsokuName === currentKinsoku) {
            selectedRadio = radioButtons[i];
            break;
        }
    }
    selectedRadio.value = true;

    var dialogButtonGroup = kinsokuDialog.add("group");
    dialogButtonGroup.orientation = "row";
    dialogButtonGroup.alignChildren = ["left", "center"];
    dialogButtonGroup.alignment = "right";
    dialogButtonGroup.spacing = 8;

    var closeButton = dialogButtonGroup.add("button", undefined, "閉じる");
    var okButton = dialogButtonGroup.add("button", undefined, "OK");

    closeButton.onClick = function () {
        kinsokuDialog.close(0);
    };

    okButton.onClick = function () {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) {
                applyKinsokuToSelection(radioButtons[i].kinsokuName);
                break;
            }
        }
        kinsokuDialog.close(1);
    };

    kinsokuDialog.center();
    kinsokuDialog.show();

})();
