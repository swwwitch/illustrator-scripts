#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

単位と数値インクリメントを、ダイアログから指定・変更します。

詳細は README を参照してください。

### Overview

Sets the units and the numeric increment from a dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PreferenceManager-print-pt";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PreferenceManager-print-pt.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager-print-pt.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "単位とインクリメント設定", en: "Units and Increments" }
        },
        panel: {
            units:      { ja: "単位", en: "Units" },
            increments: { ja: "増減値", en: "Increments" }
        },
        radio: {
            modePt: { ja: "プリント（pt）", en: "Print (pt)" },
            modeQ:  { ja: "プリント（Q）", en: "Print (Q)" },
            modePx: { ja: "オンスクリーン（px）", en: "Onscreen (px)" }
        },
        fieldLabel: {
            general:  { ja: "一般", en: "General" },
            stroke:   { ja: "線", en: "Stroke" },
            type:     { ja: "文字", en: "Text" },
            asian:    { ja: "東アジア言語", en: "East Asian" },
            key:      { ja: "キー増加", en: "Keyboard increment" },
            radius:   { ja: "角丸の半径", en: "Corner radius" },
            size:     { ja: "フォントサイズ", en: "Font size" },
            baseline: { ja: "ベースラインシフト", en: "Baseline shift" }
        },
        tooltip: {
            modePt:   { ja: "一般=mm、線=pt、文字=pt にまとめて切り替えます。", en: "Sets General=mm, Stroke=pt, Text=pt." },
            modeQ:    { ja: "一般=mm、線=mm、文字=Q にまとめて切り替えます。", en: "Sets General=mm, Stroke=mm, Text=Q." },
            modePx:   { ja: "すべての単位を px に切り替えます。", en: "Sets every unit to px." },
            general:  { ja: "定規やパネルに表示される、既定の長さの単位です。", en: "Default unit shown on rulers and panels." },
            stroke:   { ja: "線幅の入力・表示に使う単位です。", en: "Unit used for stroke weights." },
            type:     { ja: "フォントサイズや行送りに使う単位です。", en: "Unit used for font size and leading." },
            asian:    { ja: "東アジア言語のオプションで使う単位です。", en: "Unit used for East Asian typography options." },
            key:      { ja: "矢印キー1回で動く距離です。", en: "How far one arrow key press moves things." },
            radius:   { ja: "角丸ツールの既定の半径です。", en: "Default radius used by the rounded rectangle tool." },
            size:     { ja: "文字サイズ・行送りを増減する1回ぶんの量です。", en: "How much one step changes the type size or leading." },
            baseline: { ja: "ベースラインシフトを増減する1回ぶんの量です。", en: "How much one step changes the baseline shift." }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.units" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのドット区切りキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ": ");
    }

    function main() {
        // --- 単位ラベル更新関数 ---
        // 左パネルの単位選択に応じて右パネルの単位表示を更新する
        function updateIncrementLabels() {
            lblKeyUnit.text      = ddGeneral.selection.text;
            lblRadiusUnit.text   = ddGeneral.selection.text;
            lblSizeUnit.text     = ddType.selection.text;
            lblBaselineUnit.text = ddType.selection.text;
        }

        // ドキュメントが開かれていない場合は処理を中止
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてから実行してください。");
            return;
        }

        // Illustratorのバージョン番号と現在の単位設定を取得
        var versionParts = app.version.split(".");
        var aiMajor = parseInt(versionParts[0], 10);
        var aiMinor = parseInt(versionParts[1] || "0", 10);

        /* 4つの環境設定キーで同じドロップダウンを使うため、単位コード5は「Q/H」の中立表記にする
           （他のスクリプトの UNITS / getUnitInfo とは用途が異なり、pt 換算は行わない）
           The same dropdown serves four preference keys, so unit code 5 uses the neutral "Q/H" label
           (unlike the UNITS / getUnitInfo tables elsewhere, this file does no pt conversion) */
        var unitOptions = ["pt","pc","in","mm","cm","Q/H","px"];

        // 現在の単位プリファレンスを取得
        var currentGeneral = getUnitKey("rulerType");
        var currentStroke  = getUnitKey("strokeUnits");
        var currentType    = getUnitKey("text/units");
        var currentAsian   = getUnitKey("text/asianunits");

        // ダイアログ作成
        var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dlg.alignChildren = "fill";

        // --- ダイアログ位置・透明度調整 ---
        var offsetX = 300;
        var dialogOpacity = 0.97;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function () {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            };
        }

        function setDialogOpacity(dlg, opacityValue) {
            dlg.opacity = opacityValue;
        }

        setDialogOpacity(dlg, dialogOpacity);
        shiftDialogPosition(dlg, offsetX, 0);

        // 左カラムラベル幅を約6文字分に短縮 (~60px)
        var labelWidthLeft = 80;
        var labelWidthRight = 120;

        // --- 単位プリセット選択用ラジオボタン ---
        // 「プリント（pt）」「プリント（Q）」「オンスクリーン（px）」の3種
        var modeGroup = dlg.add("group");
        modeGroup.orientation = "row";
        modeGroup.alignment = "center";
        var rbPt = modeGroup.add("radiobutton", undefined, getLabel("radio.modePt"));
        rbPt.helpTip = getLabel("tooltip.modePt");
        var rbQ  = modeGroup.add("radiobutton", undefined, getLabel("radio.modeQ"));
        rbQ.helpTip = getLabel("tooltip.modeQ");
        var rbPx = modeGroup.add("radiobutton", undefined, getLabel("radio.modePx"));
        rbPx.helpTip = getLabel("tooltip.modePx");
        rbPt.value = true;

        // プリント（pt）選択時の単位設定
        rbPt.onClick = function() {
            ddGeneral.selection = arrayIndexOf(unitOptions, "mm");
            ddStroke.selection  = arrayIndexOf(unitOptions, "pt");
            ddType.selection    = arrayIndexOf(unitOptions, "pt");
            ddAsian.selection   = arrayIndexOf(unitOptions, "pt");
            updateIncrementLabels();
        };

        // プリント（Q）選択時の単位設定
        rbQ.onClick = function() {
            ddGeneral.selection = arrayIndexOf(unitOptions, "mm");
            ddStroke.selection  = arrayIndexOf(unitOptions, "Q/H");
            ddType.selection    = arrayIndexOf(unitOptions, "Q/H");
            ddAsian.selection   = arrayIndexOf(unitOptions, "Q/H");
            updateIncrementLabels();
        };

        // オンスクリーン（px）選択時の単位設定
        rbPx.onClick = function() {
            ddGeneral.selection = arrayIndexOf(unitOptions, "px");
            ddStroke.selection  = arrayIndexOf(unitOptions, "pt");
            ddType.selection    = arrayIndexOf(unitOptions, "pt");
            ddAsian.selection   = arrayIndexOf(unitOptions, "px");
            updateIncrementLabels();
        };

        // --- ダイアログ本体（2カラム構成）---
        // 左：単位選択ドロップダウン
        // 右：数値入力フィールド
        var mainGroup = dlg.add("group");
        mainGroup.orientation = "row";

        // 左カラムをパネルに変更
        var leftPanel = mainGroup.add("panel", undefined, getLabel("panel.units"));
        leftPanel.orientation = "column";
        leftPanel.alignChildren = "left";

        // 右カラムをパネルに変更
        var rightPanel = mainGroup.add("panel", undefined, getLabel("panel.increments"));
        rightPanel.orientation = "column";
        rightPanel.alignChildren = "left";

        // 左カラムに単位選択4つ
        var grpGeneral = leftPanel.add("group");
        grpGeneral.orientation = "row";
        var lblGeneral = grpGeneral.add("statictext", undefined, labelText("fieldLabel.general"));
        lblGeneral.preferredSize.width = labelWidthLeft;
        lblGeneral.justify = "right";
        var ddGeneral = grpGeneral.add("dropdownlist", undefined, unitOptions);
        ddGeneral.helpTip = getLabel("tooltip.general");
        ddGeneral.selection = arrayIndexOf(unitOptions, currentGeneral);

        var grpStroke = leftPanel.add("group");
        grpStroke.orientation = "row";
        var lblStroke = grpStroke.add("statictext", undefined, labelText("fieldLabel.stroke"));
        lblStroke.preferredSize.width = labelWidthLeft;
        lblStroke.justify = "right";
        var ddStroke = grpStroke.add("dropdownlist", undefined, unitOptions);
        ddStroke.helpTip = getLabel("tooltip.stroke");
        ddStroke.selection = arrayIndexOf(unitOptions, currentStroke);

        var grpType = leftPanel.add("group");
        grpType.orientation = "row";
        var lblType = grpType.add("statictext", undefined, labelText("fieldLabel.type"));
        lblType.preferredSize.width = labelWidthLeft;
        lblType.justify = "right";
        var ddType = grpType.add("dropdownlist", undefined, unitOptions);
        ddType.helpTip = getLabel("tooltip.type");
        ddType.selection = arrayIndexOf(unitOptions, currentType);

        var grpAsian = leftPanel.add("group");
        grpAsian.orientation = "row";
        var lblAsian = grpAsian.add("statictext", undefined, labelText("fieldLabel.asian"));
        lblAsian.preferredSize.width = labelWidthLeft;
        lblAsian.justify = "right";
        var ddAsian = grpAsian.add("dropdownlist", undefined, unitOptions);
        ddAsian.helpTip = getLabel("tooltip.asian");
        ddAsian.selection = arrayIndexOf(unitOptions, currentAsian);

        // キー増加インクリメント
        var grpKey = rightPanel.add("group");
        grpKey.orientation = "row";
        var lblKey = grpKey.add("statictext", undefined, labelText("fieldLabel.key"));
        lblKey.preferredSize.width = labelWidthRight;
        lblKey.justify = "right";
        var etKeyValue = grpKey.add("edittext", undefined, "0.1");
        etKeyValue.helpTip = getLabel("tooltip.key");
        etKeyValue.characters = 5;
        var lblKeyUnit = grpKey.add("statictext", undefined, currentGeneral);
        lblKeyUnit.preferredSize.width = 40;

        // 角丸半径の増減値
        var grpRadius = rightPanel.add("group");
        grpRadius.orientation = "row";
        var lblRadius = grpRadius.add("statictext", undefined, labelText("fieldLabel.radius"));
        lblRadius.preferredSize.width = labelWidthRight;
        lblRadius.justify = "right";
        var etRadiusValue = grpRadius.add("edittext", undefined, "1");
        etRadiusValue.helpTip = getLabel("tooltip.radius");
        etRadiusValue.characters = 5;
        var lblRadiusUnit = grpRadius.add("statictext", undefined, currentGeneral);
        lblRadiusUnit.preferredSize.width = 40;

        // フォントサイズ増減値
        var grpSize = rightPanel.add("group");
        grpSize.orientation = "row";
        var lblSize = grpSize.add("statictext", undefined, labelText("fieldLabel.size"));
        lblSize.preferredSize.width = labelWidthRight;
        lblSize.justify = "right";
        var etSizeValue = grpSize.add("edittext", undefined, "1");
        etSizeValue.helpTip = getLabel("tooltip.size");
        etSizeValue.characters = 5;
        var lblSizeUnit = grpSize.add("statictext", undefined, currentType);
        lblSizeUnit.preferredSize.width = 40;

        // ベースラインシフト増減値
        var grpBaseline = rightPanel.add("group");
        grpBaseline.orientation = "row";
        var lblBaseline = grpBaseline.add("statictext", undefined, labelText("fieldLabel.baseline"));
        lblBaseline.preferredSize.width = labelWidthRight;
        lblBaseline.justify = "right";
        var etBaselineValue = grpBaseline.add("edittext", undefined, "0.1");
        etBaselineValue.helpTip = getLabel("tooltip.baseline");
        etBaselineValue.characters = 5;
        var lblBaselineUnit = grpBaseline.add("statictext", undefined, currentType);
        lblBaselineUnit.preferredSize.width = 40;

        // OK/キャンセル
        var btnGroup = dlg.add("group");
        btnGroup.alignment = "center";

        var btnCancel = btnGroup.add("button", undefined, "Cancel", {name: "cancel"});
        var btnOK = btnGroup.add("button", undefined, "OK", {name: "ok"});

        // ドロップダウン変更時に単位ラベルを更新
        ddGeneral.onChange = updateIncrementLabels;
        ddStroke.onChange = updateIncrementLabels;
        ddType.onChange = updateIncrementLabels;
        ddAsian.onChange = updateIncrementLabels;

        dlg.onShow = function() {
            updateIncrementLabels();
        };

        // 初期ラベル反映
        updateIncrementLabels();

        if (dlg.show() != 1) return; // Cancelなら中止

        // 値を取得
        var units_general = ddGeneral.selection.text;
        var units_stroke  = ddStroke.selection.text;
        var units_type    = ddType.selection.text;
        var units_asian   = ddAsian.selection.text;

        var num_key      = etKeyValue.text;
        var unit_key     = currentGeneral;
        var num_r        = etRadiusValue.text;
        var unit_r       = currentGeneral;
        var num_size     = etSizeValue.text;
        var unit_size    = currentType;
        var num_baseline = etBaselineValue.text;
        var unit_baseline = currentType;

        // 数値プリファレンスを設定
        setPref("cursorKeyLength", unit_key, num_key);
        setPref("ovalRadius", unit_r, num_r);
        setPref("text/sizeIncrement", unit_size, num_size);
        setPref("text/riseIncrement", unit_baseline, num_baseline);

        // 単位プリファレンスを設定（バージョンチェック不要のマッピング）
        setUnits("rulerType", units_general);
        setUnits("strokeUnits", units_stroke);
        setUnits("text/units", units_type);
        setUnits("text/asianunits", units_asian);

        // alert("単位と数値インクリメントを設定しました。");
    }

    /**
     * 単位をptに変換してプリファレンスに保存
     * 数値プリファレンスを設定
     */
    function setPref(prefKey, dstUnits, num) {
        var value = calcWithUnit(num, dstUnits);
        if (dstUnits !== "pt") {
            value = convertUnit(value, dstUnits, "pt");
        }
        if (value !== undefined) {
            app.preferences.setRealPreference(prefKey, value);
        }
    }

    /**
     * 単位プリファレンスを設定（バージョンごとのマッピング + フォールバック）
     */
    function setUnits(prefKey, dstUnits) {
        // バージョンごとのマッピング定義
        var DB_DEFAULT = {
            "pt": 2, "pc": 3, "in": 0, "mm": 1, "cm": 4, "Q/H": 5, "px": 6
        };
        var DB_ALT = {
            "pt": 3, "pc": 2, "in": 1, "mm": 0, "cm": 1, "Q/H": 6, "px": 5
        };

        var versionParts = app.version.split(".");
        var aiMajor = parseInt(versionParts[0], 10);
        var aiMinor = parseInt(versionParts[1] || "0", 10);

        // バージョン別に選択
        var db;
        if (aiMajor >= 25) {
            db = DB_DEFAULT;
        } else {
            db = DB_ALT;
        }

        if (db[dstUnits] !== undefined) {
            app.preferences.setIntegerPreference(prefKey, db[dstUnits]);
        } else {
            alert("未対応の単位: " + dstUnits + " (" + prefKey + ")。既定値 pt を使用します。");
            app.preferences.setIntegerPreference(prefKey, db["pt"]);
        }
    }

    /**
     * ユーザー入力を数値に変換（単位付き数値をサポート）
     * 単位付き文字列を計算
     */
    function calcWithUnit(str, defaultUnits) {
        if (!defaultUnits) defaultUnits = "mm";
        var newStr = str.replace(/[　\s]+/g, " ").replace(",", "");
        var regUnitValue = /([0-9]+(?:\.[0-9]+)?)( ?[a-zA-Z\/]+)/g;
        newStr = newStr.replace(regUnitValue, function (m0, m1, m2) {
            return convertUnit(m1, m2.replace(/^\s+|\s+$/g, ""), defaultUnits);
        });

        var result;
        try {
            result = eval(newStr);
        } catch (e) {
            result = undefined;
        }
        return result;
    }

    /**
     * Q/H（級／歯）をmm単位に変換し、UnitValueで指定単位に変換
     * 単位変換
     */
    function convertUnit(value, srcUnit, dstUnits) {
        var qReg = /^(?:Q\/H|[qh])$/i;
        var mmUnit = "mm";

        if (qReg.test(srcUnit)) {
            value = Number(value) * 0.25;
            srcUnit = mmUnit;
        }
        if (qReg.test(dstUnits)) {
            value = Number(value) * 4;
            dstUnits = mmUnit;
        }

        var res;
        try {
            res = new UnitValue(value, srcUnit).as(dstUnits);
        } catch (e) {
            res = undefined;
        }
        return res;
    }

    /**
     * 現在のプリファレンスIDを対応する単位キーに変換
     * 現在の単位プリファレンスのキーを取得
     */
    function getUnitKey(prefKey) {
        var id = app.preferences.getIntegerPreference(prefKey);
        var db = {
            0: "in",
            1: "mm",
            2: "pt",
            3: "pica",
            4: "cm",
            5: "Q/H",
            6: "px",
            7: "ft/in",
            8: "m",
            9: "yd",
            10: "ft"
        };
        return db[id] || "pt";
    }

    // ES3互換 indexOf
    function arrayIndexOf(arr, value) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] == value) return i;
        }
        return -1;
    }

    main();

})();
