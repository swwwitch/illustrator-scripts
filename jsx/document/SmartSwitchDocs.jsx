#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のIllustratorドキュメントが開いているときに、別のドキュメントへ素早く切り替えます。
2つだけ開いているときは自動で切り替え、3つ以上のときはダイアログのリストから選びます。

詳細は README を参照してください。

### Overview

Quickly switches to another Illustrator document when several are open.
With two documents it switches automatically; with three or more it offers a list in a dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartSwitchDocs";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v0.5.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-03-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-05-25";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSwitchDocs.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSwitchDocs.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // -------------------------------
    // 日英ラベル定義　Define labels for ja/en
    // -------------------------------
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();
    var LABELS = {
        dialogTitle: {
            ja: "ドキュメント切り替え",
            en: "Switch Document"
        },
        targetDocPanel: {
            ja: "切り替え先ドキュメント",
            en: "Target Document"
        },
        currentDocPanel: {
            ja: "現在のドキュメント",
            en: "Current Document"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        }
    };

    function main() {
        // ウィンドウを統合（すべてのドキュメントをタブ表示） // Consolidate all document windows into tabs
        app.executeMenuCommand('consolidateAllWindows');
        // 開いているドキュメント数を取得
        // Get the number of open documents
        var documentCount = app.documents.length;

        // ドキュメントが0件または1件の場合は処理を終了
        // Exit if no documents or only one is open
        if (documentCount === 0 || documentCount === 1) {
            return;
        }

        // ドキュメントが2件の場合は、アクティブでない方に自動で切り替え
        // If there are two documents, switch automatically to the non-active one
        if (documentCount === 2) {
            app.activeDocument = app.documents[1];
            return;
        }

        // アクティブなドキュメントを取得
        // Get the active document
        var currentDocument = app.activeDocument;

        // ダイアログウィンドウを作成
        // Create the dialog window
        var switchDialog = new Window("dialog", LABELS.dialogTitle[lang]);
        switchDialog.orientation = "column";
        switchDialog.alignChildren = "fill";

        // 切り替え先ドキュメントパネル（現在のドキュメントは除外）
        // Panel for target documents (excluding the current document)
        var targetDocPanel = switchDialog.add("panel", undefined, LABELS.targetDocPanel[lang]);
        targetDocPanel.orientation = "column";
        targetDocPanel.alignChildren = "fill";
        targetDocPanel.margins = [10, 15, 10, 10]; // 左, 上, 右, 下  // Left, Top, Right, Bottom

        var targetDocRefs = [];  // 切り替え対象のドキュメント参照  // References to target documents

        for (var i = 0; i < app.documents.length; i++) {
            var doc = app.documents[i];
            if (doc !== currentDocument) {
                targetDocRefs.push(doc);
            }
        }

        // ドキュメント名の配列を作成（表示用）
        // Create an array of document names for display
        var targetDocNames = [];
        for (var i = 0; i < targetDocRefs.length; i++) {
            targetDocNames.push(targetDocRefs[i].name);
        }

        // ドキュメント一覧（リストボックス）
        // multiselect: false はデフォルトなので省略可能
        // Document list (listbox). multiselect: false is default and can be omitted
        var docListBox = targetDocPanel.add("listbox", undefined, targetDocNames);
        docListBox.preferredSize = [300, 150];
        if (docListBox.items.length > 0) {
            docListBox.selection = 0;
        }
        // 初期選択時にプレビューとしてアクティブドキュメントを切り替え
        // Switch active document to preview when initially selected
        if (docListBox.selection) {
            app.activeDocument = targetDocRefs[docListBox.selection.index];
        }

        // 現在のドキュメント表示パネル
        // Panel displaying the current document
        var currentDocPanel = switchDialog.add("panel", undefined, LABELS.currentDocPanel[lang]);
        currentDocPanel.orientation = "column";
        currentDocPanel.alignChildren = "left";
        currentDocPanel.margins = [15, 20, 15, 15]; // 左, 上, 右, 下  // Left, Top, Right, Bottom

        // 現在のドキュメント名を表示
        // Display the name of the current document
        currentDocPanel.add("statictext", undefined, currentDocument.name);

        // リスト選択時にドキュメントを即時切り替え
        // Immediately switch document when list selection changes
        docListBox.onChange = function() {
            var selectedIdx = docListBox.selection ? docListBox.selection.index : -1;
            if (selectedIdx >= 0) {
                var selectedDoc = targetDocRefs[selectedIdx];
                if (app.activeDocument !== selectedDoc) {
                    app.activeDocument = selectedDoc;
                }
            }
            switchDialog.active = true; // フォーカスを戻す
        };

        // ボタンを横並びに配置（OKが右）
        // Arrange buttons horizontally (OK on the right)
        var buttonGroup = switchDialog.add("group");
        buttonGroup.alignment = "right";

        var cancelButton = buttonGroup.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
        cancelButton.onClick = function() {
            app.activeDocument = currentDocument;
            switchDialog.close();
        };

        var okButton = buttonGroup.add("button", undefined, LABELS.ok[lang], { name: "ok", isDefault: true });
        okButton.onClick = function() {
            var selectedIdx = docListBox.selection ? docListBox.selection.index : -1;
            if (selectedIdx >= 0) {
                app.activeDocument = targetDocRefs[selectedIdx];
            }
            switchDialog.close();
        };

        // ダイアログ表示時にリストにフォーカスを設定
        // Set focus to the list when the dialog is shown
        switchDialog.addEventListener("show", function() {
            docListBox.active = true;
        });

        // ダイアログを表示
        // Show the dialog
        switchDialog.show();
    }

    main();

})();
