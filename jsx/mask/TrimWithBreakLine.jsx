#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

画像と帯状のパスを選択して実行すると、パスの幅で画像をトリミングし、帯が重なる範囲を取り除いて、残った上下を指定の間隔に詰めます。切り口はワープ（旗・上昇）で曲げられ、切り口に沿った省略線も引けます。

詳細は README を参照してください。

### Overview

With an image and a band-shaped path selected, trims the image to the width of the path, drops the area the band covers and closes the remaining parts up to a set gap. The cut edge can be bent with a Flag or Rise warp, and break lines can be drawn along it.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TrimWithBreakLine";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
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
var DEFAULT_GAP_MM       = 5;      /* 切り詰めたあとの上下パーツの間隔の初期値（mm） */
var DEFAULT_WARP_STYLE   = "flag"; /* ワープの初期スタイル（WARP_STYLES のキー） */
var DEFAULT_WARP_PERCENT = 3;      /* カーブの初期値（%） */
var MAX_WARP_PERCENT     = 100;    /* カーブの上限（%） */
var DEFAULT_ADD_RULE     = true;   /* 罫線を追加するかの初期値 */
var DEFAULT_RULE_DASHED  = false;  /* 罫線を破線にするかの初期値 */
var DEFAULT_GROUP_RULES  = true;   /* 罫線をパーツとグループ化するかの初期値 */
var DEFAULT_DASH_SEGMENTS = 20;    /* 破線の分割数（線分の本数）の初期値 */
var DEFAULT_ROUND_CAP    = false;  /* 罫線を丸形線端にするかの初期値 */
var RULE_STROKE_WIDTH    = 1;      /* 罫線の太さ（pt） */
var RULE_STROKE_GRAY     = 100;    /* 罫線の濃さ（0〜100のグレー） */
var TOLERANCE            = 0.001;  /* 座標比較の許容値（pt） */

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
var LABEL_WIDTH           = 75;  /* 項目名の共通幅 */
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
            trim:      { ja: "トリミング", en: "Trim" },
            breakLine: { ja: "省略線", en: "Break line" }
        },
        fieldLabel: {
            warpStyle:    { ja: "スタイル", en: "Style" },
            warpAmount:   { ja: "カーブ", en: "Bend" },
            gap:          { ja: "間隔", en: "Gap" },
            ruleStyle:    { ja: "線種", en: "Style" },
            dashSegments: { ja: "分割数", en: "Segments" },
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
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo:   { ja: "画像とパスを2つだけ選択してください。", en: "Select exactly one image and one path." },
            noBandPath:  { ja: "削除範囲を示すパスが選択されていません。", en: "No path for the area to remove is selected." },
            noImage: {
                ja: "画像が選択されていません。\n選択中: ",
                en: "No image is selected.\nSelected: "
            },
            bandOutside: {
                ja: "削除する帯は、画像の上端・下端より内側に描いてください。",
                en: "Draw the band inside the top and bottom edges of the image."
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
     * @returns {StaticText} 追加した項目名
     */
    function addRowLabel(fieldRow, labelPath) {
        var rowLabel = fieldRow.add("statictext", undefined, labelText(labelPath));
        rowLabel.preferredSize.width = LABEL_WIDTH;
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
     * @returns {{row: Group, field: EditText}} 追加した行と入力欄
     */
    function addNumberFieldRow(parentGroup, labelPath, tooltipPath, defaultText, unitText) {
        var fieldRow = addFieldRow(parentGroup);
        addRowLabel(fieldRow, labelPath);

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
     * @returns {{row: Group, radios: RadioButton[]}} 追加した行とラジオボタン
     */
    function addRadioRow(parentGroup, labelPath, tooltipPath, optionLabelPaths, selectedIndex) {
        var fieldRow = addFieldRow(parentGroup);
        addRowLabel(fieldRow, labelPath);

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
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckboxRow(parentGroup, labelPath, tooltipPath, defaultValue, indent) {
        var fieldRow = addFieldRow(parentGroup);
        if (indent) {
            var rowIndent = fieldRow.add("statictext", undefined, "");
            rowIndent.preferredSize.width = LABEL_WIDTH;
        }

        var checkbox = fieldRow.add("checkbox", undefined, getLabel(labelPath));
        checkbox.helpTip = getLabel(tooltipPath);
        checkbox.value = defaultValue;
        return checkbox;
    }

    /**
     * ボタンエリアを作る（左右中央）
     * @param {Window} dialog - 追加先のダイアログ
     * @returns {{btnOK: Button, btnCancel: Button}} 追加したボタン
     */
    function addButtonRow(dialog) {
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.spacing = BUTTON_SPACING;

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        return { btnOK: btnOK, btnCancel: btnCancel };
    }

    // =========================================
    // 入力値の扱い / Input values
    // =========================================

    /**
     * 数値を許容範囲に収める
     * @param {number} inputValue - 入力された値
     * @param {number} minValue - 下限
     * @param {number|null} maxValue - 上限（null で上限なし）
     * @param {number} fallbackValue - 数値として読めなかったときの値
     * @returns {number} 範囲に収めた値
     */
    function clampRange(inputValue, minValue, maxValue, fallbackValue) {
        if (isNaN(inputValue)) return fallbackValue;
        if (inputValue < minValue) return minValue;
        if (maxValue !== null && inputValue > maxValue) return maxValue;
        return inputValue;
    }

    /**
     * カーブの量を許容範囲に収める
     * @param {number} inputValue - 入力された値（%）
     * @returns {number} 0〜MAX_WARP_PERCENT に収めた値
     */
    function clampWarpPercent(inputValue) {
        return clampRange(inputValue, 0, MAX_WARP_PERCENT, DEFAULT_WARP_PERCENT);
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
        return clampRange(Math.round(inputValue), 1, null, DEFAULT_DASH_SEGMENTS);
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

    var targetImage = null;
    var bandPath = null;
    var selectedTypeNames = [];

    /* 選択アイテムを画像と帯のパスに振り分ける / Sort the selection into image and band path */
    for (var i = 0; i < currentSelection.length; i++) {
        var selectedItem = currentSelection[i];
        selectedTypeNames.push(selectedItem.typename);

        /* PlacedItem はリンク画像、RasterItem は埋め込み画像 */
        if (selectedItem.typename === "PlacedItem" || selectedItem.typename === "RasterItem") {
            targetImage = selectedItem;
        } else if (selectedItem.typename === "PathItem" || selectedItem.typename === "CompoundPathItem") {
            bandPath = selectedItem;
        }
    }

    if (!targetImage) {
        alert(getLabel("alert.noImage") + selectedTypeNames.join(", "));
        return;
    }

    if (!bandPath) {
        alert(getLabel("alert.noBandPath"));
        return;
    }

    /* geometricBounds は [left, top, right, bottom]、Y軸は上が大きい */
    var imageBounds = targetImage.geometricBounds;
    var bandBounds = bandPath.geometricBounds;

    var imageTop = imageBounds[1];
    var imageBottom = imageBounds[3];

    var bandLeft = bandBounds[0];
    var bandTop = bandBounds[1];
    var bandRight = bandBounds[2];
    var bandBottom = bandBounds[3];

    if (bandTop >= imageTop - TOLERANCE || bandBottom <= imageBottom + TOLERANCE) {
        alert(getLabel("alert.bandOutside"));
        return;
    }

    var parentContainer = targetImage.parent;

    /* 幅は最前面の図形（帯のパス）に合わせる。画像はこの幅にトリミングされる
       The band path in front sets the width; the image is trimmed to it */
    var maskLeft = bandLeft;
    var maskWidth = bandRight - bandLeft;
    /* マスクはいったん画像と同じ高さの矩形で作り、下辺だけを切り口として使う
       上下とも同じ下辺を使うので、どのスタイルでも切り口は必ずかみ合う */
    var maskHeight = imageTop - imageBottom;

    /* 間隔は定規の単位で入力する / the gap is entered in the ruler unit */
    var rulerUnit = getUnitInfo();
    var defaultGapValue = DEFAULT_GAP_MM * (72 / 25.4) / rulerUnit.pointsPerUnit;

    // =========================================
    // アイテムの操作 / Item helpers
    // =========================================

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
     * 切り口側と反対側を分ける境目のY座標を返す
     * @param {PageItem} maskPath - 対象のマスクパス
     * @returns {number} マスクの上下の中央
     */
    function getMaskMidY(maskPath) {
        var maskBounds = maskPath.geometricBounds;
        return (maskBounds[1] + maskBounds[3]) / 2;
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
            ' B Rotate 0 R DeformValue ' + (warpPercent / 100) +
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
     * @param {number} rectTop - 矩形の上端のY座標
     * @param {string} warpStyleKey - WARP_STYLES のキー
     * @param {number} warpPercent - カーブの量（%）
     * @returns {PageItem} マスクに使うパス
     */
    function createMaskPath(rectTop, warpStyleKey, warpPercent) {
        var maskRect = parentContainer.pathItems.rectangle(rectTop, maskLeft, maskWidth, maskHeight);
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
     * @param {number} targetY - 切り口を合わせるY座標
     * @returns {void}
     */
    function alignCutEdge(maskPath, targetY) {
        var midY = getMaskMidY(maskPath);
        var subPaths = getSubPaths(maskPath);
        var lowestY = null;
        var highestY = null;

        for (var pathIndex = 0; pathIndex < subPaths.length; pathIndex++) {
            var pathPoints = subPaths[pathIndex].pathPoints;
            for (var pointIndex = 0; pointIndex < pathPoints.length; pointIndex++) {
                var anchorY = pathPoints[pointIndex].anchor[1];
                if (anchorY > midY) continue;
                if (lowestY === null || anchorY < lowestY) lowestY = anchorY;
                if (highestY === null || anchorY > highestY) highestY = anchorY;
            }
        }

        if (lowestY === null) return;
        maskPath.translate(0, targetY - (lowestY + highestY) / 2);
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
     * 1つのパスについて、境目より上の点を指定のY座標にそろえて直線にする
     * @param {PathItem} targetPath - 対象のパス
     * @param {number} midY - 切り口と反対側を分ける境目のY座標
     * @param {number} targetY - そろえるY座標
     * @returns {void}
     */
    function flattenFarEdgePoints(targetPath, midY, targetY) {
        var pathPoints = targetPath.pathPoints;
        var isFlattened = [];

        for (var pointIndex = 0; pointIndex < pathPoints.length; pointIndex++) {
            var farPoint = pathPoints[pointIndex];
            if (farPoint.anchor[1] <= midY) {
                isFlattened.push(false);
                continue;
            }

            /* アンカーと両方のハンドルを同じ座標に置いて、直線のコーナーにする */
            var flatAnchor = [farPoint.anchor[0], targetY];
            farPoint.anchor = flatAnchor;
            farPoint.leftDirection = flatAnchor;
            farPoint.rightDirection = flatAnchor;
            farPoint.pointType = PointType.CORNER;
            isFlattened.push(true);
        }

        retractHandlesAtFlattened(pathPoints, isFlattened);
    }

    /**
     * 切り口の反対側の辺を、画像の端でまっすぐに切りそろえる
     * カーブは切り口側の辺にだけ残るので、マスクの外形は画像にぴったり収まる
     * @param {PageItem} maskPath - 対象のマスクパス
     * @param {number} targetY - そろえるY座標（画像の上端または下端）
     * @returns {void}
     */
    function flattenFarEdge(maskPath, targetY) {
        var midY = getMaskMidY(maskPath);
        var subPaths = getSubPaths(maskPath);
        for (var pathIndex = 0; pathIndex < subPaths.length; pathIndex++) {
            flattenFarEdgePoints(subPaths[pathIndex], midY, targetY);
        }
    }

    // =========================================
    // 罫線の作成 / Build the rules
    // =========================================

    /**
     * 切り口側の点を、ひと続きになる順で拾う
     * @param {PathPoints} sourcePoints - 元のパスの点
     * @param {number} midY - 切り口と反対側を分ける境目のY座標
     * @returns {number[]} 点の位置（切り口側がなければ空）
     */
    function getEdgePointIndexes(sourcePoints, midY) {
        var pointCount = sourcePoints.length;
        var isEdgePoint = [];
        var pointIndex;

        for (pointIndex = 0; pointIndex < pointCount; pointIndex++) {
            isEdgePoint.push(sourcePoints[pointIndex].anchor[1] <= midY);
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
     * @param {object} ruleSettings - { roundCap: boolean }
     * @returns {void}
     */
    function applyRuleStroke(rulePath, ruleSettings) {
        rulePath.filled = false;
        rulePath.stroked = true;
        rulePath.strokeWidth = RULE_STROKE_WIDTH;
        rulePath.strokeColor = createGrayColor(RULE_STROKE_GRAY);
        rulePath.strokeCap = ruleSettings.roundCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
    }

    /**
     * パスの切り口側の辺だけを取り出して、開いたパス（罫線）を作る
     * @param {PathItem} sourcePath - 元になるマスクのパス
     * @param {number} midY - 切り口と反対側を分ける境目のY座標
     * @param {object} ruleSettings - { ruleDashed: boolean, dashSegments: number, roundCap: boolean }
     * @returns {PathItem|null} 作成した罫線のパス（取り出せないときは null）
     */
    function createEdgeRule(sourcePath, midY, ruleSettings) {
        var sourcePoints = sourcePath.pathPoints;
        var edgeIndexes = getEdgePointIndexes(sourcePoints, midY);
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
     * @param {object} ruleSettings - { ruleDashed: boolean, dashSegments: number, roundCap: boolean }
     * @returns {PathItem[]} 作成した罫線のパス
     */
    function createEdgeRules(maskPath, ruleSettings) {
        var midY = getMaskMidY(maskPath);
        var subPaths = getSubPaths(maskPath);
        var rulePaths = [];

        for (var pathIndex = 0; pathIndex < subPaths.length; pathIndex++) {
            var rulePath = createEdgeRule(subPaths[pathIndex], midY, ruleSettings);
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
        var partImage = targetImage.duplicate(parentContainer, ElementPlacement.PLACEATBEGINNING);

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
     * @param {object} buildSettings - { warpStyleKey, warpPercent, gapPt, addRule, ruleDashed, dashSegments, roundCap, groupRules }
     * @returns {PageItem[]} 作成したクリップグループと罫線
     */
    function buildParts(buildSettings) {
        var upperMaskPath = createMaskPath(bandTop + maskHeight, buildSettings.warpStyleKey, buildSettings.warpPercent);
        /* カーブの中心を帯の上端に合わせる（上昇は切り口が片寄るため） */
        alignCutEdge(upperMaskPath, bandTop);

        var lowerMaskPath = upperMaskPath.duplicate(parentContainer, ElementPlacement.PLACEATBEGINNING);
        lowerMaskPath.translate(0, bandBottom - bandTop);

        /* 罫線は切りそろえる前の辺から取り出す / take the rules before the far edge is flattened */
        var upperRules = buildSettings.addRule ? createEdgeRules(upperMaskPath, buildSettings) : [];
        var lowerRules = buildSettings.addRule ? createEdgeRules(lowerMaskPath, buildSettings) : [];

        flattenFarEdge(upperMaskPath, imageTop);
        flattenFarEdge(lowerMaskPath, imageBottom);

        var upperPart = createClippedPart(upperMaskPath);
        var lowerPart = createClippedPart(lowerMaskPath);

        /* 下側を上へ動かし、指定の間隔をあける。下側の罫線も一緒に動かす */
        var lowerShift = bandTop - bandBottom - buildSettings.gapPt;
        lowerPart.translate(0, lowerShift);

        var ruleIndex;
        for (ruleIndex = 0; ruleIndex < lowerRules.length; ruleIndex++) {
            lowerRules[ruleIndex].translate(0, lowerShift);
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
        targetImage.hidden = false;
        bandPath.hidden = false;
    }

    /**
     * 現在の値でプレビューを作り直す
     * 複製は hidden 状態を引き継ぐので、元アイテムを表示したまま作ってから元を隠す
     * @param {object} buildSettings - { warpStyleKey, warpPercent, gapPt, addRule, ruleDashed, dashSegments, roundCap, groupRules }
     * @returns {void}
     */
    function refreshPreview(buildSettings) {
        clearPreview();
        previewItems = buildParts(buildSettings);
        targetImage.hidden = true;
        bandPath.hidden = true;
        doc.selection = null;
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 「トリミング」パネルを組み立てる
     * @param {Window} dialog - 追加先のダイアログ
     * @returns {object} パネルの入力コントロール
     */
    function buildTrimPanel(dialog) {
        var trimPanel = addPanel(dialog, "panel.trim");

        var warpStyleRow = addRadioRow(trimPanel, "fieldLabel.warpStyle", "tooltip.warpStyle",
            ["radio.flag", "radio.rise"], (DEFAULT_WARP_STYLE === "rise") ? 1 : 0);
        var warpAmountRow = addNumberFieldRow(trimPanel, "fieldLabel.warpAmount", "tooltip.warpAmount",
            String(DEFAULT_WARP_PERCENT), "%");
        var gapRow = addNumberFieldRow(trimPanel, "fieldLabel.gap", "tooltip.gap",
            formatFieldNumber(defaultGapValue), rulerUnit.label);

        return {
            styleRadios: warpStyleRow.radios,
            riseRadio: warpStyleRow.radios[1],
            warpAmountField: warpAmountRow.field,
            gapField: gapRow.field
        };
    }

    /**
     * 「省略線」パネルを組み立てる
     * @param {Window} dialog - 追加先のダイアログ
     * @returns {object} パネルの入力コントロール
     */
    function buildBreakLinePanel(dialog) {
        var breakLinePanel = addPanel(dialog, "panel.breakLine");

        var addRuleCheckbox = addCheckboxRow(breakLinePanel, "checkbox.addRule", "tooltip.addRule", DEFAULT_ADD_RULE);
        var ruleStyleRow = addRadioRow(breakLinePanel, "fieldLabel.ruleStyle", "tooltip.ruleStyle",
            ["radio.solid", "radio.dashed"], DEFAULT_RULE_DASHED ? 1 : 0);
        var dashSegmentsRow = addNumberFieldRow(breakLinePanel, "fieldLabel.dashSegments", "tooltip.dashSegments",
            String(DEFAULT_DASH_SEGMENTS));
        var strokeCapRow = addRadioRow(breakLinePanel, "fieldLabel.strokeCap", "tooltip.strokeCap",
            ["radio.buttCap", "radio.roundCap"], DEFAULT_ROUND_CAP ? 1 : 0);
        var groupRulesCheckbox = addCheckboxRow(breakLinePanel, "checkbox.groupRules", "tooltip.groupRules",
            DEFAULT_GROUP_RULES, true);

        return {
            addRuleCheckbox: addRuleCheckbox,
            styleRow: ruleStyleRow.row,
            styleRadios: ruleStyleRow.radios,
            dashedRadio: ruleStyleRow.radios[1],
            segmentsRow: dashSegmentsRow.row,
            segmentsField: dashSegmentsRow.field,
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

        var trimControls = buildTrimPanel(dialog);
        var ruleControls = buildBreakLinePanel(dialog);
        var buttonControls = addButtonRow(dialog);

        /**
         * 罫線まわりの行の有効・無効を切り替える
         * @returns {void}
         */
        function updateRuleRows() {
            var addRule = ruleControls.addRuleCheckbox.value;
            ruleControls.styleRow.enabled = addRule;
            ruleControls.segmentsRow.enabled = addRule && ruleControls.dashedRadio.value;
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
                warpStyleKey: trimControls.riseRadio.value ? "rise" : "flag",
                warpPercent: readFieldValue(trimControls.warpAmountField, clampWarpPercent),
                gapPt: gapValue * rulerUnit.pointsPerUnit,
                addRule: ruleControls.addRuleCheckbox.value,
                ruleDashed: ruleControls.dashedRadio.value,
                dashSegments: readFieldValue(ruleControls.segmentsField, clampDashSegments),
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

        wireNumberField(trimControls.warpAmountField, clampWarpPercent, onSettingChanged);
        wireNumberField(trimControls.gapField, clampGapValue, onSettingChanged);
        wireNumberField(ruleControls.segmentsField, clampDashSegments, onSettingChanged);
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
        selectItems([targetImage, bandPath]);
        return;
    }

    /* 確定実行。プレビューと同じ関数を通し、元画像と帯のパスは役目を終えるので削除する */
    var finalItems = buildParts(buildSettings);
    targetImage.remove();
    bandPath.remove();
    selectItems(finalItems);
})();
