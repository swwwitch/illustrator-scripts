#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

UIPartsPlain.jsx からステップボタンやハウスルール由来の処理を外し、ScriptUI の標準コントロールだけで組んだ比較用のテンプレートです。
基準点はラジオボタン9個、リンクアイコンはチェックボックスです。

### Overview

A comparison template that drops the stepper buttons and the house-rule extras from UIPartsPlain.jsx, using only standard ScriptUI controls.
The reference point uses nine radio buttons and the link toggle a checkbox.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "UIPartsUltraPlain";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【比較のためのメモ / Notes for the comparison】
    // ハウスルール（AGENTS.md）や共通部品の展開・一括差し替えの対象外
    //   余白・間隔・揃え … 指定なし（ScriptUI の既定のまま）
    //   数値欄           … 入力欄だけ。↑↓キーでの増減なし、幅は characters で確保
    //   基準点           … radiobutton 9個。行ごとの group に分かれるので、排他は手で管理する
    //   リンクアイコン   … checkbox
    //   文言             … ローカライズなし、日本語の直書き。コロンは半角
    //   ボタン行         … group にボタンを足すだけ

    var demoDialog = new Window("dialog", "UI パーツ " + SCRIPT_VERSION);
    var panelColumnsGroup = demoDialog.add("group");

    // =========================================
    // サイズ / Size
    // =========================================
    var sizePanel = panelColumnsGroup.add("panel", undefined, "サイズ");
    var sizeRowGroup = sizePanel.add("group");
    var sizeFieldsColumn = sizeRowGroup.add("group");
    sizeFieldsColumn.orientation = "column";
    var widthInput = addNumberField(sizeFieldsColumn, "幅:");
    var heightInput = addNumberField(sizeFieldsColumn, "高さ:");
    var linkToggle = sizeRowGroup.add("checkbox", undefined, "連動");
    linkToggle.helpTip = "幅と高さを同じ値にそろえる（クリックで切り替え）";
    linkToggle.value = true;

    widthInput.onChanging = function () { copyWhenLinked(widthInput, heightInput); };
    heightInput.onChanging = function () { copyWhenLinked(heightInput, widthInput); };
    linkToggle.onClick = function () { copyWhenLinked(widthInput, heightInput); };

    // =========================================
    // 基準点 / Reference point
    // =========================================
    var anchorPanel = panelColumnsGroup.add("panel", undefined, "基準点");
    var anchorGrid = anchorPanel.add("group");
    anchorGrid.orientation = "column";
    var anchorRadios = [];
    for (var row = 0; row < 3; row++) {
        var anchorRowGroup = anchorGrid.add("group");
        for (var column = 0; column < 3; column++) {
            var anchorRadio = anchorRowGroup.add("radiobutton", undefined, "");
            anchorRadio.helpTip = "基準点（拡大・縮小、回転の基点）";
            anchorRadio.anchorIndex = anchorRadios.length;
            anchorRadio.onClick = function () { selectAnchor(this.anchorIndex); };
            anchorRadios.push(anchorRadio);
        }
    }
    selectAnchor(4);

    // =========================================
    // 有効／無効・ボタン / Enabled toggle and buttons
    // =========================================
    var enableCheckbox = demoDialog.add("checkbox", undefined, "有効");
    enableCheckbox.helpTip = "OFF にすると、ステップボタン・基準点・リンクアイコンをディム表示にします";
    enableCheckbox.value = true;
    enableCheckbox.onClick = function () {
        /* パネルごと切り替えると中のコントロールもディム表示になる / disabling a panel dims its children */
        sizePanel.enabled = enableCheckbox.value;
        anchorPanel.enabled = enableCheckbox.value;
    };

    var buttonGroup = demoDialog.add("group");
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    demoDialog.show();

    // =========================================
    // 関数 / Functions
    // =========================================

    /* 「項目名・入力欄」の行を足す。確定時に「100 mm」の形へそろえ、数値でなければ直前の値に戻す
       add a label + field row; normalize to "100 mm" on commit, revert non-numbers */
    function addNumberField(parent, labelString) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.add("statictext", undefined, labelString);
        var numberInput = fieldRowGroup.add("edittext", undefined, "100 mm");
        numberInput.characters = 8;
        var lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (!isNaN(value)) lastValidText = Math.round(Math.max(value, 1) * 100) / 100 + " mm";
            numberInput.text = lastValidText;
        };
        return numberInput;
    }

    /* 連動中なら、変えた欄の値をもう一方へ写す / copy the value across while linked */
    function copyWhenLinked(sourceInput, targetInput) {
        if (linkToggle.value) targetInput.text = sourceInput.text;
    }

    /* ラジオは同じ親の中だけ排他なので、ほかの行の選択は手で外す / radios are exclusive only within one parent */
    function selectAnchor(anchorIndex) {
        for (var i = 0; i < anchorRadios.length; i++) {
            anchorRadios[i].value = (i === anchorIndex);
        }
    }

})();
