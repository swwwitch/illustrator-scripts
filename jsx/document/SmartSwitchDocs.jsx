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
var SCRIPT_VERSION  = "v0.5.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-03-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSwitchDocs.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSwitchDocs.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd9c7b7c077fb"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 日英ラベル定義 / Japanese-English labels
    // =========================================

    /**
     * UIロケールに応じた言語コードを返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ドキュメント切り替え", en: "Switch Document" }
        },
        panel: {
            targetDoc:  { ja: "切り替え先ドキュメント", en: "Target Document" },
            currentDoc: { ja: "現在のドキュメント", en: "Current Document" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        tooltip: {
            preview: { ja: "オンのときは選択と同時に切り替え、オフのときは［OK］をクリックしてから切り替えます", en: "On: switch as soon as the selection changes. Off: switch only after OK is clicked" }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('panel','targetDoc')）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    // =========================================
    // ドキュメント操作 / Document helpers
    // =========================================

    /**
     * アクティブドキュメント以外の開いているドキュメントを集める
     * @param {Document} activeDoc - 現在アクティブなドキュメント
     * @returns {Document[]} 切り替え先候補のドキュメント
     */
    function collectTargetDocs(activeDoc) {
        var targetDocs = [];
        for (var i = 0; i < app.documents.length; i++) {
            if (app.documents[i] !== activeDoc) {
                targetDocs.push(app.documents[i]);
            }
        }
        return targetDocs;
    }

    /**
     * ドキュメントの表示名（リストボックス用）を取り出す
     * @param {Document[]} docs - 対象ドキュメント
     * @returns {string[]} ドキュメント名の配列
     */
    function getDocNames(docs) {
        var docNames = [];
        for (var i = 0; i < docs.length; i++) {
            docNames.push(docs[i].name);
        }
        return docNames;
    }

    /**
     * 指定したドキュメントがアクティブでなければアクティブにする
     * @param {Document} doc - アクティブにするドキュメント
     * @returns {void}
     */
    function activateDoc(doc) {
        if (doc && app.activeDocument !== doc) {
            app.activeDocument = doc;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 切り替え先を選ぶダイアログを表示し、選択に応じてドキュメントを切り替える
     * @param {Document} originalDoc - 起動時にアクティブだったドキュメント
     * @param {Document[]} targetDocs - 切り替え先候補のドキュメント
     * @returns {void}
     */
    function showSwitchDialog(originalDoc, targetDocs) {
        var switchDialog = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        switchDialog.orientation = "column";
        switchDialog.alignChildren = "fill";

        /* 切り替え先パネル（現在のドキュメントは除外）/ Target panel (current document excluded) */
        var targetDocPanel = switchDialog.add("panel", undefined, getLabel("panel", "targetDoc"));
        targetDocPanel.orientation = "column";
        targetDocPanel.alignChildren = "fill";
        targetDocPanel.margins = [10, 15, 10, 10]; /* 左, 上, 右, 下 / Left, Top, Right, Bottom */

        /* ドキュメント一覧（multiselect: false は既定なので省略）/ Document list (multiselect: false is the default) */
        var targetDocList = targetDocPanel.add("listbox", undefined, getDocNames(targetDocs));
        targetDocList.preferredSize = [300, 150];
        targetDocList.selection = 0;

        /**
         * リストで選択中のドキュメントを返す
         * @returns {Document} 選択中のドキュメント（未選択のときは null）
         */
        function getSelectedDoc() {
            var selectedItem = targetDocList.selection;
            return selectedItem ? targetDocs[selectedItem.index] : null;
        }

        /* 現在のドキュメントを表示するパネル / Panel showing the current document */
        var currentDocPanel = switchDialog.add("panel", undefined, getLabel("panel", "currentDoc"));
        currentDocPanel.orientation = "column";
        currentDocPanel.alignChildren = "left";
        currentDocPanel.margins = [15, 20, 15, 15]; /* 左, 上, 右, 下 / Left, Top, Right, Bottom */
        currentDocPanel.add("statictext", undefined, originalDoc.name);

        /* ボタンエリア（左：プレビュー／右：キャンセル・OK）/ Button area (left: preview, right: cancel and OK) */
        var btnRowGroup = switchDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "top"];
        btnRowGroup.alignChildren = ["fill", "center"];

        var previewCheckbox = btnRowGroup.add("checkbox", undefined, getLabel("checkbox", "preview"));
        previewCheckbox.alignment = ["left", "center"];
        previewCheckbox.helpTip = getLabel("tooltip", "preview");
        previewCheckbox.value = true; /* 既定はプレビューON / Preview is on by default */

        /* スペーサー：ボタンを右端へ押し出す / Spacer that pushes the buttons to the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
        spacer.maximumSize.height = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.alignChildren = ["right", "center"];

        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok", isDefault: true });

        /* プレビューONなら選択と同時に切り替え、OFFなら［OK］まで切り替えない / Switch on selection when preview is on; wait for OK when off */
        targetDocList.onChange = function() {
            if (previewCheckbox.value) {
                activateDoc(getSelectedDoc());
            }
            switchDialog.active = true; /* フォーカスを戻す / Restore focus */
        };

        /* ONにしたら選択中を表示、OFFにしたら元のドキュメントへ戻す / Show the selection when turned on, restore the original when turned off */
        previewCheckbox.onClick = function() {
            activateDoc(previewCheckbox.value ? getSelectedDoc() : originalDoc);
            switchDialog.active = true;
        };

        btnCancel.onClick = function() {
            activateDoc(originalDoc); /* 元のドキュメントへ戻す / Restore the original document */
            switchDialog.close();
        };

        btnOK.onClick = function() {
            activateDoc(getSelectedDoc());
            switchDialog.close();
        };

        /* 表示時にリストへフォーカスを移す / Move focus to the list when the dialog opens */
        switchDialog.addEventListener("show", function() {
            targetDocList.active = true;
        });

        if (previewCheckbox.value) {
            activateDoc(getSelectedDoc()); /* 初期選択をプレビュー / Preview the initial selection */
        }
        switchDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 開いているドキュメント数に応じて切り替え方法を分岐する
     * @returns {void}
     */
    function main() {
        /* ウィンドウを統合（すべてのドキュメントをタブ表示）/ Consolidate all document windows into tabs */
        app.executeMenuCommand('consolidateAllWindows');

        /* 0件または1件なら切り替え先がない / Nothing to switch to with zero or one document */
        if (app.documents.length < 2) {
            return;
        }

        var originalDoc = app.activeDocument;
        var targetDocs = collectTargetDocs(originalDoc);

        /* 2件ならダイアログを出さずにもう一方へ切り替え / With two documents, switch to the other one directly */
        if (targetDocs.length === 1) {
            activateDoc(targetDocs[0]);
            return;
        }

        showSwitchDialog(originalDoc, targetDocs);
    }

    main();

})();
