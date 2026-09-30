#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ウィンドウ・パネル・タブ・行グループの余白と間隔をそろえる、UIレイアウトの再利用テンプレートです。
余白と間隔の値を1か所にまとめ、setupWindow() / setupPanel() などで各コンテナに当てます。

### Overview

A reusable template that unifies the margins and spacing of windows, panels, tabs and row groups.
The values live in one place and are applied with setupWindow(), setupPanel() and friends.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "UILayout";                     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内の「レイアウト / Layout」ブロックに貼る。
    //    識別子は WINDOW_* / PANEL_* / COLUMN_SPACING / TAB_MARGINS / setupWindow / setupPanel / setupTab / setupRow / trimButtonHeight
    // 2. 独自の margins・spacing の直書きはやめ、ウィンドウは setupWindow()、パネルはすべて setupPanel() を通す
    //      setupWindow(mainDialog);
    //      setupPanel(placementPanel);       … 間隔は PANEL_SPACING
    //      setupPanel(optionsPanel, 6);      … ラジオボタンやチェックボックスが縦に並ぶパネルは詰める
    // 3. setupPanel() は alignChildren を fill にするので、パネル内のボタンは幅いっぱいに広がる。
    //    ボタンは広げず、alignment = "left" を付ける（行グループに入れるなら setupRow() が left にする）
    //      var btnImport = placementPanel.add("button", undefined, getLabel("button.import"));
    //      btnImport.alignment = "left";
    // 4. 2カラムの間隔は COLUMN_SPACING、タブは setupTab() を使う
    // 5. trimButtonHeight() はレイアウトが決まったあと（dialog.layout.layout(true) のあとや onShow）で呼ぶ
    // 6. 値をスクリプトごとに変えたいときは、部品の中の値だけを書き換える（関数はそのまま）

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
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "UIレイアウト", en: "UI Layout" }
        },
        panel: {
            placement: { ja: "配置", en: "Placement" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            left: { ja: "左", en: "Left" },
            center: { ja: "中央", en: "Center" },
            right: { ja: "右", en: "Right" }
        },
        fieldLabel: {
            name: { ja: "名前", en: "Name" }
        },
        button: {
            load: { ja: "読み込み…", en: "Load…" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        }
    };

    /**
     * 現在の UI 言語のラベルを返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} ラベル
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelEntry) {
        return getLabel(labelEntry) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // デモ / Demo
    // =========================================
    /**
     * デモのダイアログを表示する
     * @returns {void}
     */
    function showDemoDialog() {
        var demoDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(demoDialog);

        /* 入力欄とボタンのパネル。ボタンは広げない / Field and button panel; the button does not stretch */
        var placementPanel = demoDialog.add("panel", undefined, getLabel(LABELS.panel.placement));
        setupPanel(placementPanel);
        var nameRowGroup = placementPanel.add("group");
        setupRow(nameRowGroup, "fill", 6);
        nameRowGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.name));
        var nameInput = nameRowGroup.add("edittext", undefined, "");
        nameInput.alignment = ["fill", "center"];
        var btnLoad = placementPanel.add("button", undefined, getLabel(LABELS.button.load));
        btnLoad.alignment = "left";

        /* ラジオボタンが縦に並ぶパネルは間隔を詰める / Tighter spacing for a column of radio buttons */
        var optionsPanel = demoDialog.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(optionsPanel, 6);
        optionsPanel.add("radiobutton", undefined, getLabel(LABELS.radio.left)).value = true;
        optionsPanel.add("radiobutton", undefined, getLabel(LABELS.radio.center));
        optionsPanel.add("radiobutton", undefined, getLabel(LABELS.radio.right));

        var buttonRowGroup = demoDialog.add("group");
        setupRow(buttonRowGroup, "right", 10);
        buttonRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        buttonRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        demoDialog.show();
    }

    showDemoDialog();

})();
