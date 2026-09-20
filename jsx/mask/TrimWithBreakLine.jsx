#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクト（画像・グループなど）とパスを選択して実行すると、パスの範囲を取り除き、残りを指定の間隔に詰めます。パスが対象を横にまたぐなら上下、縦にまたぐなら左右に切り分け、切り口はワープ（旗・上昇）で曲げたり、省略線を引いたりできます。

詳細は README を参照してください。

### Overview

With an object (image, group, …) and a path selected, drops the area the path covers and closes the remaining parts up to a set gap. A path spanning the artwork horizontally splits it top and bottom, one spanning it vertically splits it left and right; the cut edge can be bent with a Flag or Rise warp and traced with break lines.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TrimWithBreakLine";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-21";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TrimWithBreakLine.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TrimWithBreakLine.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2483bd96e284"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================
var DEFAULT_GAP_MM       = 3;      /* 切り詰めたあとの上下パーツの間隔の初期値（mm） */
var DEFAULT_MASK_SCALE   = 100;    /* 帯の縦スケールの初期値（%） */
var DEFAULT_MASK_OFFSET  = 0;      /* 帯の上下位置の初期値（pt、プラスで下へ） */
var MIN_BAND_MARGIN      = 1;      /* 帯を画像の内側に保つ余白（pt） */
var DEFAULT_WARP_STYLE   = "flag"; /* ワープの初期スタイル（WARP_STYLES のキー） */
var DEFAULT_WARP_PERCENT = 3;      /* カーブの初期値（%） */
var MAX_WARP_PERCENT     = 100;    /* カーブの上限（%） */
var DEFAULT_ADD_RULE     = true;   /* 罫線を追加するかの初期値 */
var DEFAULT_RULE_DASHED  = false;  /* 罫線を破線にするかの初期値 */
var DEFAULT_GROUP_RULES  = true;   /* 罫線をパーツとグループ化するかの初期値 */
var DEFAULT_DASH_SEGMENTS = 20;    /* 破線の分割数（線分の本数）の初期値 */
var DEFAULT_ROUND_CAP    = false;  /* 罫線を丸形線端にするかの初期値 */
var DEFAULT_RULE_WIDTH   = 1;      /* 罫線の線幅の初期値（pt） */
var RULE_STROKE_GRAY     = 100;    /* 罫線の濃さ（0〜100のグレー） */
var TOLERANCE            = 0.001;  /* 座標比較の許容値（pt） */

// =========================================
// 切る方向 / Cut direction
// =========================================
/* 座標の添字と合わせる（0=X、1=Y）/ matches the index into [x, y] */
var AXIS_X = 0;  /* 縦長の図形で左右に切り分ける / a tall shape splits it left and right */
var AXIS_Y = 1;  /* 横長の図形で上下に切り分ける / a wide shape splits it top and bottom */

// =========================================
// ワープの種類 / Warp styles
// =========================================
/* DeformStyle はワープの並び順+1（Arc=1 … Flag=8 … Rise=11 … Twist=15） */
var WARP_STYLES = {
    flag: { warpName: "Flag", deformStyle: 8 },   /* 旗 / Flag */
    rise: { warpName: "Rise", deformStyle: 11 }   /* 上昇 / Rise */
};

// =========================================
// 単位 / Units
// =========================================

/* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
   Unit code -> display label and points per unit */
var UNITS = [
    { label: "in",    pointsPerUnit: 72 },                /* 0 */
    { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
    { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
    { label: "pica",  pointsPerUnit: 12 },                /* 3 */
    { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
    { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
    { label: "px",    pointsPerUnit: 1 },                 /* 6 */
    { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
    { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
    { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
    { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
];

/* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
   Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

/**
 * 環境設定キーの単位を返す
 * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
 * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
 */
function getUnitInfo(prefKey) {
    var unitKey = prefKey || "rulerType";
    var unitCode = app.preferences.getIntegerPreference(unitKey);
    /* 未知のコードは pt に寄せる / unknown codes fall back to points */
    var unit = UNITS[unitCode] || UNITS[2];
    /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
    var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
    return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
}

// =========================================
// レイアウト / Layout
// =========================================
var WINDOW_MARGINS        = 16;  /* ウィンドウ外周の余白 */
var WINDOW_SPACING        = 12;  /* ウィンドウ内の要素間隔 */
var ROW_SPACING           = 6;   /* 行内の要素間隔 */
var PANEL_MARGINS         = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] */
var PANEL_SPACING         = 6;   /* パネル内の要素間隔 */
var COLUMN_SPACING        = 12;  /* 2カラムの間隔 */
var LABEL_WIDTH           = 75;  /* 左カラムの項目名の幅 */
var RULE_LABEL_WIDTH      = 60;  /* 省略線パネルの項目名の幅（「分割数：」が収まる幅） */
var FIELD_CHARACTERS      = 5;   /* 数値入力欄の文字数 */
var BUTTON_ROW_TOP_MARGIN = 8;   /* ボタンエリアの上余白 */
var BUTTON_SPACING        = 8;   /* ボタン同士の間隔 */

(function () {

    /**
     * 現在のUI言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* UIラベル定義 / UI label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "トリミングと省略線 " + SCRIPT_VERSION, en: "Trim and Break Line " + SCRIPT_VERSION }
        },
        panel: {
            mask:      { ja: "マスクの図形", en: "Mask shape" },
            trim:      { ja: "トリミング", en: "Trim" },
            breakLine: { ja: "省略線", en: "Break line" }
        },
        fieldLabel: {
            maskHeight:   { ja: "高さ", en: "Height" },
            maskWidth:    { ja: "幅", en: "Width" },
            maskOffsetY:  { ja: "上下位置", en: "Offset" },
            maskOffsetX:  { ja: "左右位置", en: "Offset" },
            warpStyle:    { ja: "スタイル", en: "Style" },
            warpAmount:   { ja: "カーブ", en: "Bend" },
            gap:          { ja: "間隔", en: "Gap" },
            ruleStyle:    { ja: "線種", en: "Style" },
            dashSegments: { ja: "分割数", en: "Segments" },
            strokeWidth:  { ja: "線幅", en: "Weight" },
            strokeCap:    { ja: "線端", en: "Cap" }
        },
        radio: {
            flag:     { ja: "旗", en: "Flag" },
            rise:     { ja: "上昇", en: "Rise" },
            solid:    { ja: "実線", en: "Solid" },
            dashed:   { ja: "破線", en: "Dashed" },
            buttCap:  { ja: "なし", en: "None" },
            roundCap: { ja: "丸型", en: "Round" }
        },
        checkbox: {
            addRule:    { ja: "罫線を追加", en: "Add rules" },
            groupRules: { ja: "グループ化", en: "Group with parts" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        tooltip: {
            maskScale: {
                ja: "マスク用の図形（描いた長方形）の、切る方向の大きさです。100%で描いたとおりになります。",
                en: "Size of the shape used as the mask along the cut direction. 100% keeps it as drawn."
            },
            maskOffset: {
                ja: "マスク用の図形の位置です。プラスで下（右）へ、マイナスで上（左）へ動きます。",
                en: "Position of the shape used as the mask. A positive value moves it down (right), a negative one up (left)."
            },
            warpStyle: {
                ja: "切り口の形です。旗は波、上昇は右上がりのカーブになります。",
                en: "Shape of the cut edge. Flag waves, Rise curves upward to the right."
            },
            warpAmount: {
                ja: "切り口を曲げる量です（0〜" + MAX_WARP_PERCENT + "%）。0にすると直線で切ります。",
                en: "How much the cut edge bends (0-" + MAX_WARP_PERCENT + "%). 0 cuts along a straight line."
            },
            gap: {
                ja: "切り詰めたあとの、上下パーツのあいだの距離です。",
                en: "Distance between the upper and lower parts after closing up."
            },
            addRule: {
                ja: "切り口に沿った線を、上下それぞれに追加します。",
                en: "Adds a line along the cut edge of each part."
            },
            ruleStyle: {
                ja: "罫線の線種です。破線は分割数で線分の数を決めます。",
                en: "Line style of the rules. A dashed rule is set by the number of segments."
            },
            dashSegments: {
                ja: "線分の本数です。線分と間隔は同じ長さで、両端が線分で終わります。",
                en: "Number of dashes. Dash and gap are equal, and both ends finish with a dash."
            },
            strokeWidth: {
                ja: "罫線の線幅です（pt）。",
                en: "Stroke weight of the rules, in points."
            },
            strokeCap: {
                ja: "罫線の線端です。丸型にすると破線が丸くなります。",
                en: "Cap of the rules. A round cap makes the dashes rounded."
            },
            groupRules: {
                ja: "罫線を、その切り口のパーツとひとつのグループにまとめます。",
                en: "Groups each rule with the part it was cut from."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo:  { ja: "オブジェクトと、マスク用のパスを2つだけ選択してください。", en: "Select exactly two objects: the artwork and the path to use as the mask." },
            noBandPath: {
                ja: "マスク用のパスが選択されていません。\n選択中: ",
                en: "No path to use as the mask is selected.\nSelected: "
            },
            bandOutsideY: {
                ja: "マスク用の図形は、オブジェクトの上端・下端より内側に描いてください。",
                en: "Draw the mask shape inside the top and bottom edges of the object."
            },
            bandOutsideX: {
                ja: "マスク用の図形は、オブジェクトの左端・右端より内側に描いてください。",
                en: "Draw the mask shape inside the left and right edges of the object."
            }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "alert.noImage" のようなドット区切りキー
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
     * コロン付きラベルを返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ドット区切りキー
     * @returns {string} コロンを付けた項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // UI部品 / UI parts
    // =========================================

    /**
     * ラベル付きパネルを追加する
     * @param {Window|Group} parentGroup - 追加先
     * @param {string} labelPath - パネル名のラベルキー
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentGroup, labelPath) {
        var settingsPanel = parentGroup.add("panel", undefined, getLabel(labelPath));
        settingsPanel.orientation = "column";
        settingsPanel.alignChildren = ["fill", "top"];
        settingsPanel.margins = PANEL_MARGINS;
        settingsPanel.spacing = PANEL_SPACING;
        return settingsPanel;
    }

    /**
     * パネルを縦に並べる列グループを作る
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加した列グループ
     */
    function addColumn(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = WINDOW_SPACING;
        return columnGroup;
    }

    /**
     * 項目名と入力を並べる行グループを作る
     * @param {Window|Group} parentGroup - 追加先
     * @returns {Group} 追加した行グループ
     */
    function addFieldRow(parentGroup) {
        var fieldRow = parentGroup.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignment = ["left", "center"];
        fieldRow.alignChildren = ["left", "center"];
        fieldRow.spacing = ROW_SPACING;
        return fieldRow;
    }

    /**
     * 右揃えの項目名を行グループに追加する
     * @param {Group} fieldRow - 追加先の行グループ
     * @param {string} labelPath - ドット区切りキー
     * @param {number} [labelWidth] - 項目名の幅（省略時は LABEL_WIDTH）
     * @returns {StaticText} 追加した項目名
     */
    function addRowLabel(fieldRow, labelPath, labelWidth) {
        var rowLabel = fieldRow.add("statictext", undefined, labelText(labelPath));
        rowLabel.preferredSize.width = labelWidth || LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 数値入力の行を作る
     * @param {Window|Group} parentGroup - 追加先
     * @param {string} labelPath - 項目名のラベルキー
     * @param {string} tooltipPath - tooltipのラベルキー
     * @param {string} defaultText - 初期値の表示
     * @param {string} [unitText] - 欄の右に置く単位の表記
     * @param {number} [labelWidth] - 項目名の幅（省略時は LABEL_WIDTH）
     * @returns {{row: Group, field: EditText}} 追加した行と入力欄
     */
    function addNumberFieldRow(parentGroup, labelPath, tooltipPath, defaultText, unitText, labelWidth) {
        var fieldRow = addFieldRow(parentGroup);
        addRowLabel(fieldRow, labelPath, labelWidth);

        var numberField = fieldRow.add("edittext", undefined, defaultText);
        numberField.characters = FIELD_CHARACTERS;
        numberField.helpTip = getLabel(tooltipPath);

        if (unitText) fieldRow.add("statictext", undefined, unitText);
        return { row: fieldRow, field: numberField };
    }

    /**
     * ラジオボタンの行を作る
     * @param {Window|Group} parentGroup - 追加先
     * @param {string} labelPath - 項目名のラベルキー
     * @param {string} tooltipPath - tooltipのラベルキー
     * @param {string[]} optionLabelPaths - 選択肢のラベルキー
     * @param {number} selectedIndex - 初期選択の位置
     * @param {number} [labelWidth] - 項目名の幅（省略時は LABEL_WIDTH）
     * @returns {{row: Group, radios: RadioButton[]}} 追加した行とラジオボタン
     */
    function addRadioRow(parentGroup, labelPath, tooltipPath, optionLabelPaths, selectedIndex, labelWidth) {
        var fieldRow = addFieldRow(parentGroup);
        addRowLabel(fieldRow, labelPath, labelWidth);

        /* ラジオは同じ親の中だけで排他になる / radios are exclusive only within one parent */
        var radioGroup = fieldRow.add("group");
        radioGroup.orientation = "row";
        radioGroup.alignChildren = ["left", "center"];
        radioGroup.spacing = ROW_SPACING;

        var radioButtons = [];
        for (var optionIndex = 0; optionIndex < optionLabelPaths.length; optionIndex++) {
            var radioButton = radioGroup.add("radiobutton", undefined, getLabel(optionLabelPaths[optionIndex]));
            radioButton.helpTip = getLabel(tooltipPath);
            radioButton.value = (optionIndex === selectedIndex);
            radioButtons.push(radioButton);
        }
        return { row: fieldRow, radios: radioButtons };
    }

    /**
     * チェックボックスの行を作る
     * @param {Window|Group} parentGroup - 追加先
     * @param {string} labelPath - チェックボックスのラベルキー
     * @param {string} tooltipPath - tooltipのラベルキー
     * @param {boolean} defaultValue - 初期値
     * @param {boolean} [indent] - true で項目名の位置に合わせて字下げする
     * @param {number} [labelWidth] - 字下げ幅（省略時は LABEL_WIDTH）
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckboxRow(parentGroup, labelPath, tooltipPath, defaultValue, indent, labelWidth) {
        var fieldRow = addFieldRow(parentGroup);
        if (indent) {
            var rowIndent = fieldRow.add("statictext", undefined, "");
            rowIndent.preferredSize.width = labelWidth || LABEL_WIDTH;
        }

        var checkbox = fieldRow.add("checkbox", undefined, getLabel(labelPath));
        checkbox.helpTip = getLabel(tooltipPath);
        checkbox.value = defaultValue;
        return checkbox;
    }

    /**
     * ボタンエリアを作る（右揃え）
     * @param {Window} dialog - 追加先のダイアログ
     * @returns {{btnOK: Button, btnCancel: Button}} 追加したボタン
     */
    function addButtonRow(dialog) {
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 伸縮するスペーサーでボタンを右へ押し出す / a stretchable spacer pushes the buttons right */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_SPACING;

        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        return { btnOK: btnOK, btnCancel: btnCancel };
    }

    // =========================================
    // 入力値の扱い / Input values
    // =========================================

    /**
     * 数値を許容範囲に収める
     * @param {number} inputValue - 入力された値
     * @param {number|null} minValue - 下限（null で下限なし）
     * @param {number|null} maxValue - 上限（null で上限なし）
     * @param {number} fallbackValue - 数値として読めなかったときの値
     * @returns {number} 範囲に収めた値
     */
    function clampRange(inputValue, minValue, maxValue, fallbackValue) {
        if (isNaN(inputValue)) return fallbackValue;
        if (minValue !== null && inputValue < minValue) return minValue;
        if (maxValue !== null && inputValue > maxValue) return maxValue;
        return inputValue;
    }

    /**
     * マスクの縦スケールを1%以上に収める
     * @param {number} inputValue - 入力された値（%）
     * @returns {number} 1以上の値
     */
    function clampMaskScale(inputValue) {
        return clampRange(inputValue, 1, null, initialValues.maskScale);
    }

    /**
     * 上下位置を読み取る（上下どちらにもずらせる）
     * @param {number} inputValue - 入力された値（pt）
     * @returns {number} そのままの値（数値として読めなければ0）
     */
    function clampOffsetValue(inputValue) {
        return clampRange(inputValue, null, null, 0);
    }

    /**
     * カーブの量を許容範囲に収める
     * @param {number} inputValue - 入力された値（%）
     * @returns {number} 0〜MAX_WARP_PERCENT に収めた値
     */
    function clampWarpPercent(inputValue) {
        return clampRange(inputValue, 0, MAX_WARP_PERCENT, initialValues.warpPercent);
    }

    /**
     * 間隔を0以上に収める
     * @param {number} inputValue - 入力された値（現在の単位）
     * @returns {number} 0以上の値
     */
    function clampGapValue(inputValue) {
        return clampRange(inputValue, 0, null, 0);
    }

    /**
     * 破線の分割数を1以上の整数に収める
     * @param {number} inputValue - 入力された値
     * @returns {number} 1以上の整数
     */
    function clampDashSegments(inputValue) {
        return clampRange(Math.round(inputValue), 1, null, initialValues.dashSegments);
    }

    /**
     * 罫線の線幅を0以上に収める
     * @param {number} inputValue - 入力された値（pt）
     * @returns {number} 0以上の値
     */
    function clampRuleWidth(inputValue) {
        return clampRange(inputValue, 0, null, initialValues.ruleWidth);
    }

    /**
     * 入力欄に表示する数値に丸める（小数第2位まで）
     * @param {number} numberValue - 丸める値
     * @returns {string} 表示用の文字列
     */
    function formatFieldNumber(numberValue) {
        return String(Math.round(numberValue * 100) / 100);
    }

    /**
     * 入力欄の値を読み取り、そろえた値を欄に戻す
     * @param {EditText} inputField - 対象の入力欄
     * @param {function} clampValue - 値を許容範囲に収める処理
     * @returns {number} 読み取った値
     */
    function readFieldValue(inputField, clampValue) {
        var fieldValue = clampValue(Number(inputField.text));
        inputField.text = formatFieldNumber(fieldValue);
        return fieldValue;
    }

    /**
     * ↑↓キーで数値を増減する（Shift：10の倍数、Option：0.1刻み）
     * @param {EditText} inputField - 対象の入力欄
     * @param {function} clampValue - 値を許容範囲に収める処理
     * @param {function} onValueChanged - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(inputField, clampValue, onValueChanged) {
        inputField.addEventListener("keydown", function (event) {
            var direction = 0;
            if (event.keyName === "Up") direction = 1;
            else if (event.keyName === "Down") direction = -1;
            else return;

            var fieldValue = Number(inputField.text);
            if (isNaN(fieldValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;

            if (keyboardState.shiftKey) {
                /* 10の倍数に揃えながら増減 / snap to multiples of ten */
                fieldValue = (direction > 0) ?
                    Math.ceil((fieldValue + 1) / 10) * 10 :
                    Math.floor((fieldValue - 1) / 10) * 10;
            } else if (keyboardState.altKey) {
                fieldValue = Math.round((fieldValue + 0.1 * direction) * 10) / 10;
            } else {
                fieldValue = Math.round(fieldValue) + direction;
            }

            inputField.text = String(clampValue(fieldValue));
            event.preventDefault();
            if (typeof onValueChanged === "function") onValueChanged();
        });
    }

    /**
     * 数値欄に↑↓キーと入力確定をつなぐ
     * @param {EditText} inputField - 対象の入力欄
     * @param {function} clampValue - 値を許容範囲に収める処理
     * @param {function} onValueChanged - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function wireNumberField(inputField, clampValue, onValueChanged) {
        changeValueByArrowKey(inputField, clampValue, onValueChanged);
        inputField.onChange = onValueChanged;
    }

    /**
     * クリックで値が変わるコントロールに同じ処理をつなぐ
     * @param {Array} clickControls - ラジオボタンやチェックボックス
     * @param {function} onValueChanged - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function wireClickControls(clickControls, onValueChanged) {
        for (var controlIndex = 0; controlIndex < clickControls.length; controlIndex++) {
            clickControls[controlIndex].onClick = onValueChanged;
        }
    }

    // =========================================
    // 設定の保存 / Saved settings
    // =========================================

    var SETTINGS_FILE = new File(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");

    /**
     * 設定ファイルを key=value の連想配列として読み込む
     * @returns {object} 保存されていた値（無ければ空）
     */
    function loadSettings() {
        var savedValues = {};
        if (!SETTINGS_FILE.exists) return savedValues;

        /* 読めない場所・壊れたファイルでも実行は続ける / keep running even if the file cannot be read */
        try {
            if (!SETTINGS_FILE.open("r")) return savedValues;
            var content = SETTINGS_FILE.read();
            SETTINGS_FILE.close();

            var lines = content.split(/\r\n|\r|\n/);
            for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
                var separatorIndex = lines[lineIndex].indexOf("=");
                if (separatorIndex < 1) continue;
                savedValues[lines[lineIndex].substring(0, separatorIndex)] = lines[lineIndex].substring(separatorIndex + 1);
            }
        } catch (e) {
        }
        return savedValues;
    }

    /**
     * 次回のために設定を保存する（長さはptで持つ）
     * @param {object} buildSettings - 確定した設定
     * @returns {void}
     */
    function saveSettings(buildSettings) {
        /* 書き込めない環境でも処理は止めない / a failed write must not stop the script */
        try {
            var settingsFolder = SETTINGS_FILE.parent;
            if (!settingsFolder.exists) settingsFolder.create();
            if (!SETTINGS_FILE.open("w")) return;

            SETTINGS_FILE.write([
                "maskScale=" + buildSettings.maskScale,
                "maskOffsetPt=" + buildSettings.maskOffsetPt,
                "warpStyle=" + buildSettings.warpStyleKey,
                "warpPercent=" + buildSettings.warpPercent,
                "gapPt=" + buildSettings.gapPt,
                "addRule=" + (buildSettings.addRule ? "1" : "0"),
                "ruleDashed=" + (buildSettings.ruleDashed ? "1" : "0"),
                "dashSegments=" + buildSettings.dashSegments,
                "ruleWidthPt=" + buildSettings.ruleWidth,
                "roundCap=" + (buildSettings.roundCap ? "1" : "0"),
                "groupRules=" + (buildSettings.groupRules ? "1" : "0")
            ].join("\n"));
            SETTINGS_FILE.close();
        } catch (e) {
        }
    }

    /**
     * 保存値を数値として取り出す
     * @param {object} savedValues - 読み込んだ設定
     * @param {string} settingKey - キー
     * @param {number} fallbackValue - 無いときの値
     * @returns {number} 取り出した値
     */
    function settingNumber(savedValues, settingKey, fallbackValue) {
        var savedValue = Number(savedValues[settingKey]);
        return (savedValues[settingKey] === undefined || isNaN(savedValue)) ? fallbackValue : savedValue;
    }

    /**
     * 保存値を真偽として取り出す
     * @param {object} savedValues - 読み込んだ設定
     * @param {string} settingKey - キー
     * @param {boolean} fallbackValue - 無いときの値
     * @returns {boolean} 取り出した値
     */
    function settingBool(savedValues, settingKey, fallbackValue) {
        if (savedValues[settingKey] === undefined) return fallbackValue;
        return savedValues[settingKey] === "1";
    }

    /**
     * 保存値を文字列として取り出す
     * @param {object} savedValues - 読み込んだ設定
     * @param {string} settingKey - キー
     * @param {string} fallbackValue - 無いときの値
     * @returns {string} 取り出した値
     */
    function settingText(savedValues, settingKey, fallbackValue) {
        return (savedValues[settingKey] === undefined) ? fallbackValue : savedValues[settingKey];
    }

    // =========================================
    // 選択の確認 / Check the selection
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var currentSelection = doc.selection;

    if (!currentSelection || currentSelection.length !== 2) {
        alert(getLabel("alert.selectTwo"));
        return;
    }

    /**
     * マスクに使えるパスかどうかを返す
     * @param {PageItem} pageItem - 調べるアイテム
     * @returns {boolean} パスまたは複合パスなら true
     */
    function isMaskablePath(pageItem) {
        return (pageItem.typename === "PathItem" || pageItem.typename === "CompoundPathItem");
    }

    var firstItem = currentSelection[0];
    var secondItem = currentSelection[1];
    var firstIsPath = isMaskablePath(firstItem);
    var secondIsPath = isMaskablePath(secondItem);

    if (!firstIsPath && !secondIsPath) {
        alert(getLabel("alert.noBandPath") + firstItem.typename + ", " + secondItem.typename);
        return;
    }

    /* 対象はどんな種類でもよい。両方パスのときは前面にあるほうをマスクに使う
       Any type can be the artwork; when both are paths, the one in front becomes the mask */
    var bandIsFirst = firstIsPath && (!secondIsPath || firstItem.zOrderPosition > secondItem.zOrderPosition);
    var bandPath = bandIsFirst ? firstItem : secondItem;
    var targetItem = bandIsFirst ? secondItem : firstItem;

    /* geometricBounds は [left, top, right, bottom]、Y軸は上が大きい */
    var targetBounds = targetItem.geometricBounds;
    var bandBounds = bandPath.geometricBounds;

    /* 対象をより広くまたいでいる向きに切る。パス自体の縦横比ではなく、対象に対する割合で見る
       cut across whichever direction the path spans more of the artwork */
    var bandCoverX = (bandBounds[2] - bandBounds[0]) / (targetBounds[2] - targetBounds[0]);
    var bandCoverY = (bandBounds[1] - bandBounds[3]) / (targetBounds[1] - targetBounds[3]);
    var axisIndex = (bandCoverX >= bandCoverY) ? AXIS_Y : AXIS_X;

    /**
     * 切る方向の、座標が大きいほうの端を返す（縦なら上端、横なら右端）
     * @param {number[]} bounds - geometricBounds
     * @returns {number} 端の座標
     */
    function getAxisMax(bounds) {
        return (axisIndex === AXIS_Y) ? bounds[1] : bounds[2];
    }

    /**
     * 切る方向の、座標が小さいほうの端を返す（縦なら下端、横なら左端）
     * @param {number[]} bounds - geometricBounds
     * @returns {number} 端の座標
     */
    function getAxisMin(bounds) {
        return (axisIndex === AXIS_Y) ? bounds[3] : bounds[0];
    }

    /**
     * 切る方向と直交する側の、座標が大きいほうの端を返す
     * @param {number[]} bounds - geometricBounds
     * @returns {number} 端の座標
     */
    function getCrossMax(bounds) {
        return (axisIndex === AXIS_Y) ? bounds[2] : bounds[1];
    }

    /**
     * 切る方向と直交する側の、座標が小さいほうの端を返す
     * @param {number[]} bounds - geometricBounds
     * @returns {number} 端の座標
     */
    function getCrossMin(bounds) {
        return (axisIndex === AXIS_Y) ? bounds[0] : bounds[3];
    }

    var targetAxisMax = getAxisMax(targetBounds);
    var targetAxisMin = getAxisMin(targetBounds);
    var bandAxisMax = getAxisMax(bandBounds);
    var bandAxisMin = getAxisMin(bandBounds);

    if (bandAxisMax >= targetAxisMax - TOLERANCE || bandAxisMin <= targetAxisMin + TOLERANCE) {
        alert(getLabel((axisIndex === AXIS_Y) ? "alert.bandOutsideY" : "alert.bandOutsideX"));
        return;
    }

    var parentContainer = targetItem.parent;

    /* 切る方向と直交する幅は、最前面の図形（マスク用のパス）に合わせる。対象はこの幅にトリミングされる
       The path in front sets the size across the cut; the artwork is trimmed to it */
    var maskCrossMin = getCrossMin(bandBounds);
    var maskCrossMax = getCrossMax(bandBounds);
    var maskCrossSize = maskCrossMax - maskCrossMin;
    /* マスクはいったん対象と同じ大きさの矩形で作り、切り口側の辺だけを使う
       両方とも同じ辺を使うので、どのスタイルでも切り口は必ずかみ合う */
    var maskAxisSize = targetAxisMax - targetAxisMin;

    /* 上下はプラスで下へ、左右はプラスで右へ / down is positive for Y, right is positive for X */
    var offsetSign = (axisIndex === AXIS_Y) ? -1 : 1;

    /* 距離は定規の単位、線幅は線の単位で入力する
       distances use the ruler unit, the stroke weight uses the stroke unit */
    var rulerUnit = getUnitInfo();
    var strokeUnit = getUnitInfo("strokeUnits");

    /* 前回の設定があればそれを初期値にする / start from the settings saved last time */
    var savedValues = loadSettings();
    var initialValues = {
        maskScale:    settingNumber(savedValues, "maskScale", DEFAULT_MASK_SCALE),
        maskOffset:   settingNumber(savedValues, "maskOffsetPt", DEFAULT_MASK_OFFSET) / rulerUnit.pointsPerUnit,
        warpStyleKey: settingText(savedValues, "warpStyle", DEFAULT_WARP_STYLE),
        warpPercent:  settingNumber(savedValues, "warpPercent", DEFAULT_WARP_PERCENT),
        gap:          settingNumber(savedValues, "gapPt", DEFAULT_GAP_MM * 72 / 25.4) / rulerUnit.pointsPerUnit,
        addRule:      settingBool(savedValues, "addRule", DEFAULT_ADD_RULE),
        ruleDashed:   settingBool(savedValues, "ruleDashed", DEFAULT_RULE_DASHED),
        dashSegments: settingNumber(savedValues, "dashSegments", DEFAULT_DASH_SEGMENTS),
        ruleWidth:    settingNumber(savedValues, "ruleWidthPt", DEFAULT_RULE_WIDTH) / strokeUnit.pointsPerUnit,
        roundCap:     settingBool(savedValues, "roundCap", DEFAULT_ROUND_CAP),
        groupRules:   settingBool(savedValues, "groupRules", DEFAULT_GROUP_RULES)
    };

    // =========================================
    // アイテムの操作 / Item helpers
    // =========================================

    /**
     * アイテムを切る方向へ動かす
     * @param {PageItem} pageItem - 動かすアイテム
     * @param {number} delta - 動かす量（座標が大きくなる向きが正）
     * @returns {void}
     */
    function translateAlongAxis(pageItem, delta) {
        pageItem.translate((axisIndex === AXIS_X) ? delta : 0, (axisIndex === AXIS_Y) ? delta : 0);
    }

    /**
     * 点の座標を、切る方向だけ置き換えた組にする
     * @param {number[]} anchor - もとの座標 [x, y]
     * @param {number} axisValue - 切る方向の座標
     * @returns {number[]} 置き換えた座標
     */
    function withAxisValue(anchor, axisValue) {
        return (axisIndex === AXIS_Y) ? [anchor[0], axisValue] : [axisValue, anchor[1]];
    }

    /**
     * グレーの色を作る
     * @param {number} grayValue - 0〜100の濃さ
     * @returns {GrayColor} 作った色
     */
    function createGrayColor(grayValue) {
        var grayColor = new GrayColor();
        grayColor.gray = grayValue;
        return grayColor;
    }

    /**
     * 渡したアイテムだけを選択する
     * @param {Array} pageItems - 選択するアイテム
     * @returns {void}
     */
    function selectItems(pageItems) {
        doc.selection = null;
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            pageItems[itemIndex].selected = true;
        }
    }

    /**
     * マスクの中に並ぶパスを配列で返す（複合パスにも対応）
     * @param {PageItem} maskPath - 対象のマスクパス
     * @returns {PathItem[]} 中のパス
     */
    function getSubPaths(maskPath) {
        if (maskPath.typename !== "CompoundPathItem") return [maskPath];

        var subPaths = [];
        for (var pathIndex = 0; pathIndex < maskPath.pathItems.length; pathIndex++) {
            subPaths.push(maskPath.pathItems[pathIndex]);
        }
        return subPaths;
    }

    /**
     * 切り口側と反対側を分ける境目の座標を返す
     * @param {PageItem} maskPath - 対象のマスクパス
     * @returns {number} 切る方向の中央
     */
    function getMaskAxisMid(maskPath) {
        var maskBounds = maskPath.geometricBounds;
        return (getAxisMax(maskBounds) + getAxisMin(maskBounds)) / 2;
    }

    // =========================================
    // マスクの作成 / Build the masks
    // =========================================

    /**
     * パスにワープ効果を適用する
     * @param {PathItem} targetPath - 効果を適用するパス
     * @param {string} warpStyleKey - WARP_STYLES のキー（"flag" / "rise"）
     * @param {number} warpPercent - カーブの量（%）
     * @returns {void}
     */
    function applyWarp(targetPath, warpStyleKey, warpPercent) {
        var warpStyle = WARP_STYLES[warpStyleKey] || WARP_STYLES[DEFAULT_WARP_STYLE];
        var warpXml = '<LiveEffect name="Adobe Deform"><Dict data="S DisplayString Warp:' + warpStyle.warpName +
            ' I DeformStyle ' + warpStyle.deformStyle +
            ' B Rotate ' + ((axisIndex === AXIS_Y) ? 0 : 1) + ' R DeformValue ' + (warpPercent / 100) +
            ' R DeformHoriz 0 R DeformVert 0 "/></LiveEffect>';
        targetPath.applyEffect(warpXml);
    }

    /**
     * アピアランスを分割し、分割後のパスを返す
     * @param {PageItem} targetItem - 分割するアイテム
     * @returns {PageItem} 分割後のパス（1つだけを含むグループは中身を取り出す）
     */
    function expandAppearance(targetItem) {
        doc.selection = null;
        targetItem.selected = true;
        /* 選択を反映させてからコマンドを流す / let the new selection settle before the command */
        app.redraw();
        app.executeMenuCommand("expandStyle");

        var expandedItems = doc.selection;
        var expandedItem = (expandedItems && expandedItems.length) ? expandedItems[0] : targetItem;

        /* 分割結果がグループで返ることがあるので、中身のパスを取り出して空のグループを削除する */
        if (expandedItem.typename === "GroupItem") {
            var innerItem = expandedItem;
            while (innerItem.typename === "GroupItem" && innerItem.pageItems.length === 1) {
                innerItem = innerItem.pageItems[0];
            }
            if (innerItem !== expandedItem) {
                innerItem.move(parentContainer, ElementPlacement.PLACEATBEGINNING);
                expandedItem.remove();
                expandedItem = innerItem;
            }
        }

        doc.selection = null;
        return expandedItem;
    }

    /**
     * マスク用の矩形を作り、カーブの量に応じてワープを掛けて分割する
     * @param {number} axisMax - 切る方向の、座標が大きいほうの端
     * @param {number} axisSize - 切る方向の大きさ
     * @param {string} warpStyleKey - WARP_STYLES のキー
     * @param {number} warpPercent - カーブの量（%）
     * @returns {PageItem} マスクに使うパス
     */
    function createMaskPath(axisMax, axisSize, warpStyleKey, warpPercent) {
        /* rectangle(top, left, width, height) は常に上端・左端で指定する */
        var maskRect = (axisIndex === AXIS_Y) ?
            parentContainer.pathItems.rectangle(axisMax, maskCrossMin, maskCrossSize, axisSize) :
            parentContainer.pathItems.rectangle(maskCrossMax, axisMax - axisSize, axisSize, maskCrossSize);
        /* 塗りがないとワープを分割できない / the warp needs a filled path to expand */
        maskRect.filled = true;
        maskRect.fillColor = createGrayColor(100);
        maskRect.stroked = false;

        if (warpPercent === 0) return maskRect;

        applyWarp(maskRect, warpStyleKey, warpPercent);
        return expandAppearance(maskRect);
    }

    /**
     * 切り口の上下の中心が指定のY座標に来るようにマスクを動かす
     * 上昇のようにカーブが片寄るスタイルでも、帯の位置で切れるようにする
     * @param {PageItem} maskPath - 対象のマスクパス
     * @param {number} axisTarget - 切り口を合わせる座標
     * @returns {void}
     */
    function alignCutEdge(maskPath, axisTarget) {
        var axisMid = getMaskAxisMid(maskPath);
        var subPaths = getSubPaths(maskPath);
        var lowestValue = null;
        var highestValue = null;

        for (var pathIndex = 0; pathIndex < subPaths.length; pathIndex++) {
            var pathPoints = subPaths[pathIndex].pathPoints;
            for (var pointIndex = 0; pointIndex < pathPoints.length; pointIndex++) {
                var axisValue = pathPoints[pointIndex].anchor[axisIndex];
                if (axisValue > axisMid) continue;
                if (lowestValue === null || axisValue < lowestValue) lowestValue = axisValue;
                if (highestValue === null || axisValue > highestValue) highestValue = axisValue;
            }
        }

        if (lowestValue === null) return;
        translateAlongAxis(maskPath, axisTarget - (lowestValue + highestValue) / 2);
    }

    /**
     * そろえた点へ向かうハンドルを、隣り合う点から畳む
     * 残すと縦の辺がふくらんで余計なカーブになる
     * @param {PathPoints} pathPoints - 対象のパスの点
     * @param {boolean[]} isFlattened - 点ごとの、そろえたかどうか
     * @returns {void}
     */
    function retractHandlesAtFlattened(pathPoints, isFlattened) {
        var pointCount = pathPoints.length;
        for (var pointIndex = 0; pointIndex < pointCount; pointIndex++) {
            if (isFlattened[pointIndex]) continue;

            var edgePoint = pathPoints[pointIndex];
            var previousIndex = (pointIndex + pointCount - 1) % pointCount;
            var nextIndex = (pointIndex + 1) % pointCount;

            if (isFlattened[previousIndex]) {
                edgePoint.leftDirection = edgePoint.anchor;
                edgePoint.pointType = PointType.CORNER;
            }
            if (isFlattened[nextIndex]) {
                edgePoint.rightDirection = edgePoint.anchor;
                edgePoint.pointType = PointType.CORNER;
            }
        }
    }

    /**
     * 1つのパスについて、境目より外側の点を指定の座標にそろえて直線にする
     * @param {PathItem} targetPath - 対象のパス
     * @param {number} axisMid - 切り口と反対側を分ける境目の座標
     * @param {number} axisTarget - そろえる座標
     * @returns {void}
     */
    function flattenFarEdgePoints(targetPath, axisMid, axisTarget) {
        var pathPoints = targetPath.pathPoints;
        var isFlattened = [];

        for (var pointIndex = 0; pointIndex < pathPoints.length; pointIndex++) {
            var farPoint = pathPoints[pointIndex];
            if (farPoint.anchor[axisIndex] <= axisMid) {
                isFlattened.push(false);
                continue;
            }

            /* アンカーと両方のハンドルを同じ座標に置いて、直線のコーナーにする */
            var flatAnchor = withAxisValue(farPoint.anchor, axisTarget);
            farPoint.anchor = flatAnchor;
            farPoint.leftDirection = flatAnchor;
            farPoint.rightDirection = flatAnchor;
            farPoint.pointType = PointType.CORNER;
            isFlattened.push(true);
        }

        retractHandlesAtFlattened(pathPoints, isFlattened);
    }

    /**
     * 切り口の反対側の辺を、対象の端でまっすぐに切りそろえる
     * カーブは切り口側の辺にだけ残るので、マスクの外形は対象にぴったり収まる
     * @param {PageItem} maskPath - 対象のマスクパス
     * @param {number} axisTarget - そろえる座標（対象の端）
     * @returns {void}
     */
    function flattenFarEdge(maskPath, axisTarget) {
        var axisMid = getMaskAxisMid(maskPath);
        var subPaths = getSubPaths(maskPath);
        for (var pathIndex = 0; pathIndex < subPaths.length; pathIndex++) {
            flattenFarEdgePoints(subPaths[pathIndex], axisMid, axisTarget);
        }
    }

    // =========================================
    // 罫線の作成 / Build the rules
    // =========================================

    /**
     * 切り口側の点を、ひと続きになる順で拾う
     * @param {PathPoints} sourcePoints - 元のパスの点
     * @param {number} axisMid - 切り口と反対側を分ける境目の座標
     * @returns {number[]} 点の位置（切り口側がなければ空）
     */
    function getEdgePointIndexes(sourcePoints, axisMid) {
        var pointCount = sourcePoints.length;
        var isEdgePoint = [];
        var pointIndex;

        for (pointIndex = 0; pointIndex < pointCount; pointIndex++) {
            isEdgePoint.push(sourcePoints[pointIndex].anchor[axisIndex] <= axisMid);
        }

        /* 切り口の点がひと続きになる開始位置を探す / find where the run of edge points starts */
        var startIndex = -1;
        for (pointIndex = 0; pointIndex < pointCount; pointIndex++) {
            if (isEdgePoint[pointIndex] && !isEdgePoint[(pointIndex + pointCount - 1) % pointCount]) {
                startIndex = pointIndex;
                break;
            }
        }
        if (startIndex < 0) return [];

        var edgeIndexes = [];
        for (var stepIndex = 0; stepIndex < pointCount; stepIndex++) {
            var sourceIndex = (startIndex + stepIndex) % pointCount;
            if (!isEdgePoint[sourceIndex]) break;
            edgeIndexes.push(sourceIndex);
        }
        return edgeIndexes;
    }

    /**
     * 罫線の線の設定（太さ・色・線端）を適用する
     * @param {PathItem} rulePath - 対象の罫線
     * @param {object} ruleSettings - { ruleWidth: number, roundCap: boolean }
     * @returns {void}
     */
    function applyRuleStroke(rulePath, ruleSettings) {
        rulePath.filled = false;
        rulePath.stroked = true;
        rulePath.strokeWidth = ruleSettings.ruleWidth;
        rulePath.strokeColor = createGrayColor(RULE_STROKE_GRAY);
        rulePath.strokeCap = ruleSettings.roundCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
    }

    /**
     * パスの切り口側の辺だけを取り出して、開いたパス（罫線）を作る
     * @param {PathItem} sourcePath - 元になるマスクのパス
     * @param {number} axisMid - 切り口と反対側を分ける境目の座標
     * @param {object} ruleSettings - { ruleDashed: boolean, dashSegments: number, ruleWidth: number, roundCap: boolean }
     * @returns {PathItem|null} 作成した罫線のパス（取り出せないときは null）
     */
    function createEdgeRule(sourcePath, axisMid, ruleSettings) {
        var sourcePoints = sourcePath.pathPoints;
        var edgeIndexes = getEdgePointIndexes(sourcePoints, axisMid);
        if (!edgeIndexes.length) return null;

        var rulePath = parentContainer.pathItems.add();
        rulePath.closed = false;
        applyRuleStroke(rulePath, ruleSettings);

        for (var edgeIndex = 0; edgeIndex < edgeIndexes.length; edgeIndex++) {
            var sourcePoint = sourcePoints[edgeIndexes[edgeIndex]];
            var rulePoint = rulePath.pathPoints.add();
            rulePoint.anchor = sourcePoint.anchor;
            rulePoint.leftDirection = sourcePoint.leftDirection;
            rulePoint.rightDirection = sourcePoint.rightDirection;
            rulePoint.pointType = sourcePoint.pointType;
        }

        /* 両端の外向きハンドルは縦の辺に伸びていたものなので畳む */
        var firstPoint = rulePath.pathPoints[0];
        var lastPoint = rulePath.pathPoints[rulePath.pathPoints.length - 1];
        firstPoint.leftDirection = firstPoint.anchor;
        lastPoint.rightDirection = lastPoint.anchor;

        /* 線分が n 本、間隔が n−1 回で、罫線の長さちょうどに収める
           n dashes and n-1 gaps fill the whole rule */
        if (ruleSettings.ruleDashed) {
            var dashLength = rulePath.length / (ruleSettings.dashSegments * 2 - 1);
            rulePath.strokeDashes = [dashLength, dashLength];
        }

        return rulePath;
    }

    /**
     * マスクの切り口に沿った罫線を作る（複合パスは中のパスごとに作る）
     * @param {PageItem} maskPath - 元になるマスクのパス
     * @param {object} ruleSettings - { ruleDashed: boolean, dashSegments: number, ruleWidth: number, roundCap: boolean }
     * @returns {PathItem[]} 作成した罫線のパス
     */
    function createEdgeRules(maskPath, ruleSettings) {
        var axisMid = getMaskAxisMid(maskPath);
        var subPaths = getSubPaths(maskPath);
        var rulePaths = [];

        for (var pathIndex = 0; pathIndex < subPaths.length; pathIndex++) {
            var rulePath = createEdgeRule(subPaths[pathIndex], axisMid, ruleSettings);
            if (rulePath) rulePaths.push(rulePath);
        }
        return rulePaths;
    }

    // =========================================
    // 組み立て / Assemble the parts
    // =========================================

    /**
     * 画像を複製し、渡したパスでクリッピングマスクを作成する
     * @param {PageItem} maskPath - マスクに使うパス
     * @returns {GroupItem} 作成したクリップグループ
     */
    function createClippedPart(maskPath) {
        var partImage = targetItem.duplicate(parentContainer, ElementPlacement.PLACEATBEGINNING);

        var clipGroup = parentContainer.groupItems.add();
        maskPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        partImage.move(clipGroup, ElementPlacement.PLACEATEND);
        clipGroup.clipped = true;

        return clipGroup;
    }

    /**
     * 罫線をパーツとひとつのグループにまとめる
     * @param {GroupItem} clipGroup - クリップグループ
     * @param {PathItem[]} rulePaths - まとめる罫線
     * @returns {PageItem} まとめたグループ（罫線がなければクリップグループのまま）
     */
    function groupWithRules(clipGroup, rulePaths) {
        if (!rulePaths.length) return clipGroup;

        var partGroup = parentContainer.groupItems.add();
        for (var ruleIndex = 0; ruleIndex < rulePaths.length; ruleIndex++) {
            rulePaths[ruleIndex].move(partGroup, ElementPlacement.PLACEATEND);
        }
        clipGroup.move(partGroup, ElementPlacement.PLACEATEND);
        return partGroup;
    }

    /**
     * 帯を取り除いた上下のパーツを作る
     * 上側は下辺を帯の上端に、下側は同じ複製の下辺を帯の下端に合わせ、
     * それぞれ反対側の辺を画像の上端・下端に切りそろえる
     * @param {object} buildSettings - { maskScale, maskOffsetPt, warpStyleKey, warpPercent, gapPt, addRule, ruleDashed, dashSegments, ruleWidth, roundCap, groupRules }
     * @returns {PageItem[]} 作成したクリップグループと罫線
     */
    function buildParts(buildSettings) {
        /* 描いたマスク用の図形を、大きさと位置の設定で伸縮・移動した範囲を取り除く
           the mask shape as drawn, scaled and shifted by the mask settings */
        var bandCenter = (bandAxisMax + bandAxisMin) / 2 + offsetSign * buildSettings.maskOffsetPt;
        var bandHalfSize = (bandAxisMax - bandAxisMin) * buildSettings.maskScale / 200;
        var cutMax = bandCenter + bandHalfSize;
        var cutMin = bandCenter - bandHalfSize;

        /* 対象からはみ出すと両側のパーツが作れなくなるので、内側に収める */
        if (cutMax > targetAxisMax - MIN_BAND_MARGIN) cutMax = targetAxisMax - MIN_BAND_MARGIN;
        if (cutMin < targetAxisMin + MIN_BAND_MARGIN) cutMin = targetAxisMin + MIN_BAND_MARGIN;

        var upperMaskPath = createMaskPath(cutMax + maskAxisSize, maskAxisSize,
            buildSettings.warpStyleKey, buildSettings.warpPercent);

        /* カーブの中心を切り口に合わせる（上昇は切り口が片寄るため） */
        alignCutEdge(upperMaskPath, cutMax);

        var lowerMaskPath = upperMaskPath.duplicate(parentContainer, ElementPlacement.PLACEATBEGINNING);
        translateAlongAxis(lowerMaskPath, cutMin - cutMax);

        /* 罫線は切りそろえる前の辺から取り出す / take the rules before the far edge is flattened */
        var upperRules = buildSettings.addRule ? createEdgeRules(upperMaskPath, buildSettings) : [];
        var lowerRules = buildSettings.addRule ? createEdgeRules(lowerMaskPath, buildSettings) : [];

        flattenFarEdge(upperMaskPath, targetAxisMax);
        flattenFarEdge(lowerMaskPath, targetAxisMin);

        var upperPart = createClippedPart(upperMaskPath);
        var lowerPart = createClippedPart(lowerMaskPath);

        /* 残った側を寄せて、指定の間隔をあける。そちらの罫線も一緒に動かす */
        var lowerShift = cutMax - cutMin - buildSettings.gapPt;
        translateAlongAxis(lowerPart, lowerShift);

        var ruleIndex;
        for (ruleIndex = 0; ruleIndex < lowerRules.length; ruleIndex++) {
            translateAlongAxis(lowerRules[ruleIndex], lowerShift);
        }

        /* 罫線は切り口の上に出す / bring the rules in front of the parts */
        var rulePaths = upperRules.concat(lowerRules);
        for (ruleIndex = 0; ruleIndex < rulePaths.length; ruleIndex++) {
            rulePaths[ruleIndex].move(parentContainer, ElementPlacement.PLACEATBEGINNING);
        }

        if (buildSettings.groupRules) {
            return [groupWithRules(upperPart, upperRules), groupWithRules(lowerPart, lowerRules)];
        }
        return [upperPart, lowerPart].concat(rulePaths);
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* プレビューで作った物の控え / Items created for the preview */
    var previewItems = null;

    /**
     * プレビューを取り消して元の表示に戻す
     * @returns {void}
     */
    function clearPreview() {
        if (previewItems) {
            for (var itemIndex = 0; itemIndex < previewItems.length; itemIndex++) {
                previewItems[itemIndex].remove();
            }
            previewItems = null;
        }
        targetItem.hidden = false;
        bandPath.hidden = false;
    }

    /**
     * 現在の値でプレビューを作り直す
     * 複製は hidden 状態を引き継ぐので、元アイテムを表示したまま作ってから元を隠す
     * @param {object} buildSettings - { maskScale, maskOffsetPt, warpStyleKey, warpPercent, gapPt, addRule, ruleDashed, dashSegments, ruleWidth, roundCap, groupRules }
     * @returns {void}
     */
    function refreshPreview(buildSettings) {
        clearPreview();
        previewItems = buildParts(buildSettings);
        targetItem.hidden = true;
        bandPath.hidden = true;
        doc.selection = null;
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 「マスクの図形」パネルを組み立てる
     * @param {Group} parentGroup - 追加先の列グループ
     * @returns {object} パネルの入力コントロール
     */
    function buildMaskPanel(parentGroup) {
        var maskPanel = addPanel(parentGroup, "panel.mask");

        var sizeLabelPath = (axisIndex === AXIS_Y) ? "fieldLabel.maskHeight" : "fieldLabel.maskWidth";
        var offsetLabelPath = (axisIndex === AXIS_Y) ? "fieldLabel.maskOffsetY" : "fieldLabel.maskOffsetX";

        var scaleRow = addNumberFieldRow(maskPanel, sizeLabelPath, "tooltip.maskScale",
            formatFieldNumber(initialValues.maskScale), "%");
        var offsetRow = addNumberFieldRow(maskPanel, offsetLabelPath, "tooltip.maskOffset",
            formatFieldNumber(initialValues.maskOffset), rulerUnit.label);

        return { scaleField: scaleRow.field, offsetField: offsetRow.field };
    }

    /**
     * 「トリミング」パネルを組み立てる
     * @param {Group} parentGroup - 追加先の列グループ
     * @returns {object} パネルの入力コントロール
     */
    function buildTrimPanel(parentGroup) {
        var trimPanel = addPanel(parentGroup, "panel.trim");

        var warpStyleRow = addRadioRow(trimPanel, "fieldLabel.warpStyle", "tooltip.warpStyle",
            ["radio.flag", "radio.rise"], (initialValues.warpStyleKey === "rise") ? 1 : 0);
        var warpAmountRow = addNumberFieldRow(trimPanel, "fieldLabel.warpAmount", "tooltip.warpAmount",
            formatFieldNumber(initialValues.warpPercent), "%");
        var gapRow = addNumberFieldRow(trimPanel, "fieldLabel.gap", "tooltip.gap",
            formatFieldNumber(initialValues.gap), rulerUnit.label);

        return {
            styleRadios: warpStyleRow.radios,
            riseRadio: warpStyleRow.radios[1],
            warpAmountField: warpAmountRow.field,
            gapField: gapRow.field
        };
    }

    /**
     * 「省略線」パネルを組み立てる
     * @param {Group} parentGroup - 追加先の列グループ
     * @returns {object} パネルの入力コントロール
     */
    function buildBreakLinePanel(parentGroup) {
        var breakLinePanel = addPanel(parentGroup, "panel.breakLine");

        var addRuleCheckbox = addCheckboxRow(breakLinePanel, "checkbox.addRule", "tooltip.addRule", initialValues.addRule);
        var ruleStyleRow = addRadioRow(breakLinePanel, "fieldLabel.ruleStyle", "tooltip.ruleStyle",
            ["radio.solid", "radio.dashed"], initialValues.ruleDashed ? 1 : 0, RULE_LABEL_WIDTH);
        var dashSegmentsRow = addNumberFieldRow(breakLinePanel, "fieldLabel.dashSegments", "tooltip.dashSegments",
            formatFieldNumber(initialValues.dashSegments), "", RULE_LABEL_WIDTH);
        var strokeWidthRow = addNumberFieldRow(breakLinePanel, "fieldLabel.strokeWidth", "tooltip.strokeWidth",
            formatFieldNumber(initialValues.ruleWidth), strokeUnit.label, RULE_LABEL_WIDTH);
        var strokeCapRow = addRadioRow(breakLinePanel, "fieldLabel.strokeCap", "tooltip.strokeCap",
            ["radio.buttCap", "radio.roundCap"], initialValues.roundCap ? 1 : 0, RULE_LABEL_WIDTH);
        var groupRulesCheckbox = addCheckboxRow(breakLinePanel, "checkbox.groupRules", "tooltip.groupRules",
            initialValues.groupRules, true, RULE_LABEL_WIDTH);

        return {
            addRuleCheckbox: addRuleCheckbox,
            styleRow: ruleStyleRow.row,
            styleRadios: ruleStyleRow.radios,
            dashedRadio: ruleStyleRow.radios[1],
            segmentsRow: dashSegmentsRow.row,
            segmentsField: dashSegmentsRow.field,
            widthRow: strokeWidthRow.row,
            widthField: strokeWidthRow.field,
            capRow: strokeCapRow.row,
            capRadios: strokeCapRow.radios,
            roundCapRadio: strokeCapRow.radios[1],
            groupRulesCheckbox: groupRulesCheckbox
        };
    }

    /**
     * 切り口の形・間隔・省略線を指定するダイアログを表示する
     * @returns {object|null} 設定（キャンセル時は null）
     */
    function showSettingsDialog() {
        var dialog = new Window("dialog", getLabel("dialog.title"));
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = WINDOW_MARGINS;
        dialog.spacing = WINDOW_SPACING;

        /* 左に「マスクの図形」と「トリミング」、右に「省略線」を置く
           mask and trim settings on the left, break lines on the right */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumn = addColumn(columnsGroup);
        var rightColumn = addColumn(columnsGroup);

        var maskControls = buildMaskPanel(leftColumn);
        var trimControls = buildTrimPanel(leftColumn);
        var ruleControls = buildBreakLinePanel(rightColumn);
        var buttonControls = addButtonRow(dialog);

        /**
         * 罫線まわりの行の有効・無効を切り替える
         * @returns {void}
         */
        function updateRuleRows() {
            var addRule = ruleControls.addRuleCheckbox.value;
            ruleControls.styleRow.enabled = addRule;
            ruleControls.segmentsRow.enabled = addRule && ruleControls.dashedRadio.value;
            ruleControls.widthRow.enabled = addRule;
            ruleControls.capRow.enabled = addRule;
            ruleControls.groupRulesCheckbox.enabled = addRule;
        }

        /**
         * ダイアログの入力内容を読み取る（値は欄にそろえ直す）
         * @returns {object} 設定
         */
        function collectSettings() {
            var gapValue = readFieldValue(trimControls.gapField, clampGapValue);
            return {
                maskScale: readFieldValue(maskControls.scaleField, clampMaskScale),
                maskOffsetPt: readFieldValue(maskControls.offsetField, clampOffsetValue) * rulerUnit.pointsPerUnit,
                warpStyleKey: trimControls.riseRadio.value ? "rise" : "flag",
                warpPercent: readFieldValue(trimControls.warpAmountField, clampWarpPercent),
                gapPt: gapValue * rulerUnit.pointsPerUnit,
                addRule: ruleControls.addRuleCheckbox.value,
                ruleDashed: ruleControls.dashedRadio.value,
                dashSegments: readFieldValue(ruleControls.segmentsField, clampDashSegments),
                ruleWidth: readFieldValue(ruleControls.widthField, clampRuleWidth) * strokeUnit.pointsPerUnit,
                roundCap: ruleControls.roundCapRadio.value,
                groupRules: ruleControls.groupRulesCheckbox.value
            };
        }

        /**
         * 入力が変わったらプレビューを作り直す
         * @returns {void}
         */
        function onSettingChanged() {
            updateRuleRows();
            refreshPreview(collectSettings());
        }

        wireNumberField(maskControls.scaleField, clampMaskScale, onSettingChanged);
        wireNumberField(maskControls.offsetField, clampOffsetValue, onSettingChanged);
        wireNumberField(trimControls.warpAmountField, clampWarpPercent, onSettingChanged);
        wireNumberField(trimControls.gapField, clampGapValue, onSettingChanged);
        wireNumberField(ruleControls.segmentsField, clampDashSegments, onSettingChanged);
        wireNumberField(ruleControls.widthField, clampRuleWidth, onSettingChanged);
        wireClickControls(trimControls.styleRadios
            .concat([ruleControls.addRuleCheckbox])
            .concat(ruleControls.styleRadios)
            .concat(ruleControls.capRadios)
            .concat([ruleControls.groupRulesCheckbox]), onSettingChanged);

        var dialogSettings = null;
        buttonControls.btnOK.onClick = function () {
            dialogSettings = collectSettings();
            dialog.close(1);
        };
        buttonControls.btnCancel.onClick = function () {
            dialog.close(0);
        };

        /* 開いた時点のプレビューを先に出す / Show the preview before the dialog appears */
        updateRuleRows();
        refreshPreview(collectSettings());

        return (dialog.show() === 1) ? dialogSettings : null;
    }

    // =========================================
    // 実行 / Run
    // =========================================

    var buildSettings = showSettingsDialog();
    clearPreview();

    if (!buildSettings) {
        selectItems([targetItem, bandPath]);
        return;
    }

    saveSettings(buildSettings);

    /* 確定実行。プレビューと同じ関数を通し、元のオブジェクトとマスク用のパスは役目を終えるので削除する */
    var finalItems = buildParts(buildSettings);
    targetItem.remove();
    bandPath.remove();
    selectItems(finalItems);
})();
