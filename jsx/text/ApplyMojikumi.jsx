#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

/*

### 概要

選択したテキストフレーム（グループ内も再帰的に処理）の全段落に、
文字組みアキ量設定（mojikumi）をダイアログのラジオボタンから適用します。
ラジオボタンを選ぶと、選択中のテキストへ即時に反映されます。
プリセットは「なし」「行末約物全角/半角」「約物半角」「行末約物半角」
「行末約物全角」「約物全角」「ツメ組み」「ベタ組み」の8種類です。

### 使い方

1. テキストオブジェクトを選択する
2. スクリプトを実行する
3. ダイアログで設定を選ぶと、選択中のテキストへ即適用される
4. ［閉じる］でダイアログを閉じる

### 紹介記事



*/

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 文字組みプリセット（mojikumiSet のインデックス。-1 は「なし」）/ Mojikumi presets (mojikumiSet index; -1 = none) */
    var mojikumiPresets = [
        { mojikumiIndex: -1, labelText: "なし" },
        { mojikumiIndex: 0, labelText: "行末約物全角/半角" },
        { mojikumiIndex: 1, labelText: "約物半角" },
        { mojikumiIndex: 2, labelText: "行末約物半角" },
        { mojikumiIndex: 3, labelText: "行末約物全角" },
        { mojikumiIndex: 4, labelText: "約物全角" },
        { mojikumiIndex: 5, labelText: "ツメ組み" },
        { mojikumiIndex: 6, labelText: "ベタ組み" }
    ];

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 前提チェック：ドキュメントが開かれているか / Precondition: a document is open */
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    /* 前提チェック：オブジェクトが選択されているか / Precondition: something is selected */
    if (doc.selection.length === 0) {
        alert("テキストオブジェクトを選択してください。");
        return;
    }

    /* テキストフレーム内の全段落に適用 / Apply to every paragraph in a text frame */
    function applyMojikumiToTextFrame(textFrame, mojikumiValue) {
        var paragraphs = textFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            paragraphs[i].paragraphAttributes.mojikumi = mojikumiValue;
        }
    }

    /* 種類で振り分け（TextFrame は直接、GroupItem は再帰）/ Dispatch by item type (TextFrame directly, GroupItem recursively) */
    function applyMojikumiToItem(item, mojikumiValue) {
        if (item.typename === "TextFrame") {
            applyMojikumiToTextFrame(item, mojikumiValue);
        } else if (item.typename === "GroupItem") {
            applyMojikumiToGroup(item, mojikumiValue);
        }
    }

    /* グループ内の各アイテムを再帰処理 / Recurse into group children */
    function applyMojikumiToGroup(groupItem, mojikumiValue) {
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            applyMojikumiToItem(groupItem.pageItems[i], mojikumiValue);
        }
    }

    /* 選択全体に適用して再描画 / Apply to the whole selection and redraw */
    function applyMojikumiToSelection(mojikumiIndex) {
        /* 値を1回だけ解決して再帰へ渡す（-1 は「なし」）/ Resolve the value once, then pass it down (-1 = none) */
        var mojikumiValue = (mojikumiIndex === -1) ? "なし" : doc.mojikumiSet[mojikumiIndex];
        for (var i = 0; i < doc.selection.length; i++) {
            applyMojikumiToItem(doc.selection[i], mojikumiValue);
        }
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ダイアログを組み立てて表示 / Build and show the dialog */
    function showMojikumiDialog() {
        var dialog = new Window("dialog", "文字組みアキ量設定  " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        /* プリセット選択パネル / Preset selection panel */
        var presetPanel = dialog.add("panel", undefined, "設定を選択");
        presetPanel.orientation = "column";
        presetPanel.alignChildren = "left";
        presetPanel.alignment = "fill";
        presetPanel.margins = [16, 20, 16, 12];
        presetPanel.spacing = 6;

        /* プリセットごとにラジオボタンを生成 / Build one radio button per preset */
        var mojikumiRadioButtons = [];
        for (var i = 0; i < mojikumiPresets.length; i++) {
            var radioButton = presetPanel.add("radiobutton", undefined, mojikumiPresets[i].labelText);
            radioButton.mojikumiIndex = mojikumiPresets[i].mojikumiIndex;
            mojikumiRadioButtons.push(radioButton);
        }
        mojikumiRadioButtons[0].value = true;

        /* 選択時に即適用 / Apply immediately on selection */
        for (var i = 0; i < mojikumiRadioButtons.length; i++) {
            mojikumiRadioButtons[i].onClick = function () {
                applyMojikumiToSelection(this.mojikumiIndex);
            };
        }

        /* 閉じるボタン / Close button */
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["left", "center"];
        buttonGroup.alignment = "right";
        buttonGroup.spacing = 8;

        var closeButton = buttonGroup.add("button", undefined, "閉じる");
        closeButton.onClick = function () {
            dialog.close();
        };

        dialog.center();
        dialog.show();
    }

    showMojikumiDialog();

})();
