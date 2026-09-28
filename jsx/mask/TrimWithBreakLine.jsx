#target illustrator
#targetengine "TrimWithBreakLineEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクト（画像・グループなど）とパスを選択して実行すると、パスの範囲を取り除いて残りを指定の間隔に詰めます（パスが対象の端を覆っているときは、その側だけを残します）。切り口はワープで曲げたりギザギザにしたりでき、省略線も引けます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TrimWithBreakLine.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n2483bd96e284

### Overview

With an object (image, group, …) and a path selected, drops the area the path covers and closes the remaining parts up to a set gap (when the path covers an edge of the artwork, only that side is kept). The cut edge can be bent with a warp or made jagged, and traced with break lines.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TrimWithBreakLine.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TrimWithBreakLine";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

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
var DEFAULT_MASK_CROSS_SCALE = 100; /* 帯の、切り口に沿った方向のスケールの初期値（%） */
var DEFAULT_MASK_OFFSET  = 0;      /* 帯の上下位置の初期値（pt、プラスで下へ） */
var MIN_BAND_MARGIN      = 1;      /* 帯を画像の内側に保つ余白（pt） */
var DEFAULT_WARP_STYLE   = "flag"; /* ワープの初期スタイル（WARP_STYLES のキー） */
var DEFAULT_WARP_PERCENT = 3;      /* カーブの初期値（%） */
var MAX_WARP_PERCENT     = 100;    /* カーブの上限（%、マイナス側も同じ幅） */
var DEFAULT_JAGGED_SIZE  = 3;      /* ギザギザの大きさの初期値（pt） */
var DEFAULT_JAGGED_RIDGES = 10;    /* ギザギザの折り返しの初期値（回） */
var MAX_JAGGED_RIDGES    = 100;    /* 折り返しの上限（ジグザグ効果と同じ） */
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
    flag:         { warpName: "Flag", deformStyle: 8 },                  /* 旗 / Flag */
    rise:         { warpName: "Rise", deformStyle: 11 },                 /* 上昇 / Rise */
    riseStraight: { warpName: "Rise", deformStyle: 11, straight: true }, /* 直線（上昇を直線化）/ Straight (Rise, straightened) */
    jagged:       { jagged: true }                                       /* ギザギザ（ジグザグ効果）/ Jagged (Zig Zag effect) */
};

/* アイコンボタンに並べる順 / the order they appear as icon buttons */
var WARP_STYLE_KEYS = ["flag", "rise", "riseStraight", "jagged"];

/**
 * ジグザグ効果のXMLを返す（大きさは絶対値、折り返しは直線）
 * 大きさ0・折り返し0にすると、ワープのカーブを直線に置き換えられる
 * @param {number} amountPt - 大きさ（pt）
 * @param {number} ridges - 線分あたりの折り返し
 * @returns {string} LiveEffect のXML
 */
function createZigzagXml(amountPt, ridges) {
    return '<LiveEffect name="Adobe Zigzag"><Dict data="R amount ' + amountPt +
        ' R relAmount 0 R absoluteness 1 R ridges ' + ridges + ' R roundness 0 "/></LiveEffect>';
}

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
var RADIO_COLUMN_SPACING  = 2;   /* 縦に並べたラジオボタンの間隔 */
var STYLE_ICON_BUTTON_SIZE = 36;  /* 切り口の形のアイコンボタンの一辺 */
var STYLE_ICON_SIZE       = 24;  /* アイコンの図形（正方形）の一辺 */
var STYLE_ICON_CUT_WIDTH  = 2.5; /* アイコンの切り口の線幅 */
var STYLE_ICON_SPACING    = 4;   /* アイコンボタンの間隔 */
var STYLE_ICON_BOTTOM_MARGIN = 5; /* アイコンボタンの下の余白 */
var PANEL_MARGINS         = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] */
var PANEL_SPACING         = 6;   /* パネル内の要素間隔 */
var COLUMN_SPACING        = 12;  /* 2カラムの間隔 */
var LABEL_WIDTH           = 75;  /* 左カラムの項目名の幅 */
var RULE_LABEL_WIDTH      = 60;  /* 省略線パネルの項目名の幅（「分割数：」が収まる幅） */
var FIELD_CHARACTERS      = 4;   /* 数値入力欄の文字数 */

// =========================================
// ステップボタン / Stepper buttons
// =========================================
var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリック・shift＋↑↓でそろえる倍数 / Shift snaps to multiples of this */
var STEPPER_OPTION_STEP    = 0.1; /* option＋クリック・option＋↑↓の増減量 / Option step */

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* UIラベル定義 / UI label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "省略線でトリミング " + SCRIPT_VERSION, en: "Trim with Break Lines " + SCRIPT_VERSION }
        },
        panel: {
            mask:      { ja: "マスク用の図形", en: "Mask shape" },
            cutEdge:   { ja: "切り口", en: "Cut edge" },
            breakLine: { ja: "省略線", en: "Break lines" }
        },
        fieldLabel: {
            maskHeight:   { ja: "高さ", en: "Height" },
            maskWidth:    { ja: "幅", en: "Width" },
            maskOffsetY:  { ja: "上下位置", en: "Offset" },
            maskOffsetX:  { ja: "左右位置", en: "Offset" },
            warpAmount:   { ja: "カーブ", en: "Bend" },
            jaggedSize:   { ja: "大きさ", en: "Size" },
            jaggedRidges: { ja: "折り返し", en: "Ridges" },
            gap:          { ja: "間隔", en: "Gap" },
            ruleStyle:    { ja: "線種", en: "Line style" },
            dashSegments: { ja: "分割数", en: "Segments" },
            strokeWidth:  { ja: "線幅", en: "Weight" },
            strokeCap:    { ja: "線端", en: "Cap" }
        },
        radio: {
            flag:         { ja: "旗", en: "Flag" },
            rise:         { ja: "上昇", en: "Rise" },
            riseStraight: { ja: "直線", en: "Straight" },
            jagged:       { ja: "ギザギザ", en: "Jagged" },
            solid:        { ja: "実線", en: "Solid" },
            dashed:       { ja: "破線", en: "Dashed" },
            buttCap:      { ja: "なし", en: "Butt" },
            roundCap:     { ja: "丸型", en: "Round" }
        },
        checkbox: {
            addRule:    { ja: "省略線を追加", en: "Add break lines" },
            groupRules: { ja: "グループ化", en: "Group with parts" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        tooltip: {
            maskScale: {
                ja: "取り除く範囲の大きさです。100%でマスク用の図形のとおりになります。",
                en: "Size of the area to remove. 100% matches the mask shape."
            },
            maskCrossScale: {
                ja: "オブジェクトを切りそろえる範囲の大きさです。100%でマスク用の図形のとおりになります。",
                en: "Size the artwork is trimmed to. 100% matches the mask shape."
            },
            maskOffsetY: {
                ja: "マスク用の図形の位置です。プラスで下へ、マイナスで上へ動きます。",
                en: "Position of the shape used as the mask. A positive value moves it down, a negative one up."
            },
            maskOffsetX: {
                ja: "マスク用の図形の位置です。プラスで右へ、マイナスで左へ動きます。",
                en: "Position of the shape used as the mask. A positive value moves it right, a negative one left."
            },
            flag:         { ja: "切り口を波形にします。", en: "Makes the cut edge wavy." },
            rise:         { ja: "切り口を、片側へせり上がるカーブにします。", en: "Curves the cut edge up toward one side." },
            riseStraight: { ja: "切り口を斜めの直線にします。", en: "Slants the cut edge in a straight line." },
            jagged:       { ja: "切り口をギザギザにします。", en: "Makes the cut edge jagged." },
            warpAmount: {
                ja: "切り口を曲げる量です（-" + MAX_WARP_PERCENT + "〜" + MAX_WARP_PERCENT + "%）。直線では傾きの量になります。0でまっすぐに切り、マイナスで逆向きになります。",
                en: "How much the cut edge bends (-" + MAX_WARP_PERCENT + " to " + MAX_WARP_PERCENT + "%); for Straight, how much it slants. 0 cuts straight across; negative values reverse the direction."
            },
            jaggedSize: {
                ja: "ギザギザの山の高さです。0でまっすぐに切ります。",
                en: "Height of the jagged ridges. 0 cuts straight across."
            },
            jaggedRidges: {
                ja: "切り口の端から端までの折り返しの数です（0〜" + MAX_JAGGED_RIDGES + "）。0でまっすぐに切ります。",
                en: "Number of ridges across the cut edge (0 to " + MAX_JAGGED_RIDGES + "). 0 cuts straight across."
            },
            gap: {
                ja: "切り詰めたあとの、2つのパーツのあいだの距離です。片側だけ残すときは使いません。",
                en: "Distance between the two parts after closing up. Unused when only one side is kept."
            },
            addRule: {
                ja: "切り口に沿った省略線を、パーツごとに追加します。",
                en: "Adds a break line along the cut edge of each part."
            },
            ruleStyle: {
                ja: "省略線の線種です。破線は分割数で線分の数を決めます。",
                en: "Line style of the break lines. A dashed line is set by the number of segments."
            },
            dashSegments: {
                ja: "線分の本数です。線分と間隔は同じ長さで、両端が線分で終わります。",
                en: "Number of dashes. Dash and gap are equal, and both ends finish with a dash."
            },
            strokeWidth: {
                ja: "省略線の線幅です。単位は環境設定の「線」に従います。",
                en: "Stroke weight of the break lines. The unit follows the Stroke preference."
            },
            strokeCap: {
                ja: "省略線の線端です。丸型にすると線の両端（破線は各線分の端）が丸くなります。",
                en: "Cap of the break lines. A round cap rounds the line ends (each dash when dashed)."
            },
            groupRules: {
                ja: "省略線を、その切り口のパーツとひとつのグループにまとめます。",
                en: "Groups each break line with the part it was cut from."
            },
            maskPanel: {
                ja: "描いたマスク用の図形をもとに、取り除く範囲と残す範囲を調整します。",
                en: "Adjusts the areas to remove and keep, starting from the mask shape you drew."
            },
            cutEdgePanel: {
                ja: "切り口の形と、詰めたあとの間隔を指定します。",
                en: "Sets the shape of the cut edge and the gap after closing up."
            },
            breakLinePanel: {
                ja: "切り口に沿って引く線を設定します。",
                en: "Sets up the lines drawn along the cut edge."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo:  { ja: "オブジェクトと、マスク用のパスの2つを選択してください。", en: "Select exactly two objects: the artwork and the path to use as the mask." },
            noBandPath: {
                ja: "マスク用のパスが選択されていません。\n選択中：",
                en: "No path to use as the mask is selected.\nSelected: "
            },
            bandCoversAllY: {
                ja: "マスク用の図形が、オブジェクトの上端から下端までを覆っています。どちらかの端は空けてください。",
                en: "The mask shape covers the artwork from its top edge to its bottom edge. Leave one of the ends clear."
            },
            bandCoversAllX: {
                ja: "マスク用の図形が、オブジェクトの左端から右端までを覆っています。どちらかの端は空けてください。",
                en: "The mask shape covers the artwork from its left edge to its right edge. Leave one of the ends clear."
            }
        }
    };

    // =========================================
    // UI部品 / UI parts
    // =========================================

    /**
     * ラベル付きパネルを追加する
     * @param {Window|Group} parentGroup - 追加先
     * @param {string} labelPath - パネル名のラベルキー
     * @param {string} [tooltipPath] - tooltipのラベルキー
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentGroup, labelPath, tooltipPath) {
        var settingsPanel = parentGroup.add("panel", undefined, getLabel(labelPath));
        if (tooltipPath) settingsPanel.helpTip = getLabel(tooltipPath);
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
     * @param {string|string[]} [tooltipPath] - 入力欄と同じtooltipのラベルキー（配列のときは付けない）
     * @returns {StaticText} 追加した項目名
     */
    function addRowLabel(fieldRow, labelPath, labelWidth, tooltipPath) {
        var rowLabel = fieldRow.add("statictext", undefined, labelText(labelPath));
        rowLabel.preferredSize.width = labelWidth || LABEL_WIDTH;
        rowLabel.justify = "right";
        if (typeof tooltipPath === "string") rowLabel.helpTip = getLabel(tooltipPath);
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
     * @returns {{row: Group, field: EditText}} 追加した行と入力欄（∧∨は field.stepperGroup で参照できる）
     */
    function addNumberFieldRow(parentGroup, labelPath, tooltipPath, defaultText, unitText, labelWidth) {
        var fieldRow = addFieldRow(parentGroup);
        addRowLabel(fieldRow, labelPath, labelWidth, tooltipPath);

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var stepperGroup = addStepper(stepperFieldGroup);
        var numberField = stepperFieldGroup.add("edittext", undefined, defaultText);
        numberField.characters = FIELD_CHARACTERS;
        numberField.helpTip = getLabel(tooltipPath);
        numberField.stepperGroup = stepperGroup;

        if (unitText) fieldRow.add("statictext", undefined, unitText);
        return { row: fieldRow, field: numberField };
    }

    /**
     * 行の有効／無効を切り替え、行内の∧∨を描き直す（自作描画は自動でディムにならないため）
     * @param {Group} fieldRow - addNumberFieldRow() などで作った行
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setRowEnabled(fieldRow, isEnabled) {
        if (fieldRow.enabled === isEnabled) return;
        fieldRow.enabled = isEnabled;
        redrawSteppersIn(fieldRow);
    }

    /**
     * コントロール以下にある∧∨ボタンをすべて描き直す
     * @param {Object} container - 行やグループ
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var childIndex = 0; childIndex < container.children.length; childIndex++) {
            var child = container.children[childIndex];
            if (child.isStepperButton) redrawGroup(child);
            else redrawSteppersIn(child);
        }
    }

    // =========================================
    // ステップボタン / Stepper buttons
    // =========================================

    /* UIがダークテーマかどうか（取得できなければ明るいUI扱い） / whether the UI is dark */
    var STEPPER_UI_DARK = (function () {
        try {
            return app.preferences.getRealPreference("uiBrightness") <= 0.5;
        } catch (e) {
            return false;
        }
    })();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる
       colors are translucent overlays so they follow every UI brightness level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線 / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する。
     * 増減の中身は wireNumberField() が stepperGroup.stepBy に入れる
     * @param {Group} parentGroup - 追加先
     * @returns {Group} ∧∨をまとめた group
     */
    function addStepper(parentGroup) {
        var stepperGroup = parentGroup.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = 0;
        stepperGroup.alignment = ["left", "center"];
        stepperGroup.stepBy = null;

        stepperGroup.upButton = makeStepperChevronButton(stepperGroup, "up", function () {
            if (stepperGroup.stepBy) stepperGroup.stepBy(1);
        });
        stepperGroup.downButton = makeStepperChevronButton(stepperGroup, "down", function () {
            if (stepperGroup.stepBy) stepperGroup.stepBy(-1);
        });
        return stepperGroup;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group} parentGroup - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parentGroup, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parentGroup.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true;

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、行ごと無効なら薄い色で描く / dim when the row is disabled */
            var isDimmed = !isEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    /**
     * ラジオボタンの行を作る
     * @param {Window|Group} parentGroup - 追加先
     * @param {string|null} labelPath - 項目名のラベルキー（null で項目名なし）
     * @param {string|string[]} tooltipPath - tooltipのラベルキー（配列なら選択肢ごと）
     * @param {string[]} optionLabelPaths - 選択肢のラベルキー
     * @param {number} selectedIndex - 初期選択の位置
     * @param {number} [labelWidth] - 項目名の幅（省略時は LABEL_WIDTH）
     * @param {boolean} [isVertical] - 選択肢を縦に並べるか
     * @returns {{row: Group, radios: RadioButton[]}} 追加した行とラジオボタン
     */
    function addRadioRow(parentGroup, labelPath, tooltipPath, optionLabelPaths, selectedIndex, labelWidth, isVertical) {
        var fieldRow = addFieldRow(parentGroup);
        if (labelPath) {
            var rowLabel = addRowLabel(fieldRow, labelPath, labelWidth, tooltipPath);
            /* 縦に並べるときは、項目名を1つめの選択肢の高さに合わせる */
            if (isVertical) rowLabel.alignment = ["left", "top"];
        }

        /* ラジオは同じ親の中だけで排他になる / radios are exclusive only within one parent */
        var radioGroup = fieldRow.add("group");
        radioGroup.orientation = isVertical ? "column" : "row";
        radioGroup.alignChildren = ["left", "center"];
        radioGroup.spacing = isVertical ? RADIO_COLUMN_SPACING : ROW_SPACING;

        var radioButtons = [];
        for (var optionIndex = 0; optionIndex < optionLabelPaths.length; optionIndex++) {
            var radioButton = radioGroup.add("radiobutton", undefined, getLabel(optionLabelPaths[optionIndex]));
            radioButton.helpTip = getLabel((typeof tooltipPath === "string") ? tooltipPath : tooltipPath[optionIndex]);
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
     * マスクの、切り口に沿った方向のスケールを1%以上に収める
     * @param {number} inputValue - 入力された値（%）
     * @returns {number} 1以上の値
     */
    function clampMaskCrossScale(inputValue) {
        return clampRange(inputValue, 1, null, initialValues.maskCrossScale);
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
     * @returns {number} -MAX_WARP_PERCENT〜MAX_WARP_PERCENT に収めた値
     */
    function clampWarpPercent(inputValue) {
        return clampRange(inputValue, -MAX_WARP_PERCENT, MAX_WARP_PERCENT, initialValues.warpPercent);
    }

    /**
     * ギザギザの大きさを0以上に収める
     * @param {number} inputValue - 入力された値（現在の単位）
     * @returns {number} 0以上の値
     */
    function clampJaggedSize(inputValue) {
        return clampRange(inputValue, 0, null, initialValues.jaggedSize);
    }

    /**
     * 折り返しを0〜MAX_JAGGED_RIDGES の整数に収める
     * @param {number} inputValue - 入力された値
     * @returns {number} 0〜MAX_JAGGED_RIDGES の整数
     */
    function clampJaggedRidges(inputValue) {
        return clampRange(Math.round(inputValue), 0, MAX_JAGGED_RIDGES, initialValues.jaggedRidges);
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
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は次の整数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して次の整数へ）
     * @param {number} fieldValue - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {boolean} isInteger - 整数の欄なら true
     * @returns {number} 増減した値（範囲は未適用）
     */
    function computeSteppedValue(fieldValue, direction, isInteger) {
        var keyboardState = ScriptUI.environment.keyboardState;
        if (keyboardState.shiftKey) return snapToNextMultiple(fieldValue, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyboardState.altKey && !isInteger) return fieldValue + direction * STEPPER_OPTION_STEP;
        return snapToNextMultiple(fieldValue, 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（232→240、下げるときは 232→230、230→220）
     * @param {number} fieldValue - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapToNextMultiple(fieldValue, multiple, direction) {
        if (direction > 0) return Math.floor(fieldValue / multiple) * multiple + multiple;
        return Math.ceil(fieldValue / multiple) * multiple - multiple;
    }

    /**
     * 数値欄に∧∨・↑↓キー・入力確定をつなぐ。∧∨と↑↓キーは同じ増減処理を通す
     * @param {EditText} inputField - addNumberFieldRow() で作った入力欄
     * @param {function} clampValue - 値を許容範囲に収める処理
     * @param {function} onValueChanged - 値を変えたあとに呼ぶ処理
     * @param {boolean} [isInteger] - 整数の欄なら true（option の0.1刻みを使わない）
     * @returns {void}
     */
    function wireNumberField(inputField, clampValue, onValueChanged, isInteger) {
        var stepperGroup = inputField.stepperGroup;

        /**
         * 入力欄の値を1回分増減し、範囲に収めて書き戻す
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        stepperGroup.stepBy = function (direction) {
            if (!isEnabledInTree(inputField)) return; /* 行が無効の間は動かさない */
            var fieldValue = Number(inputField.text);
            if (isNaN(fieldValue)) fieldValue = 0;
            inputField.text = formatFieldNumber(clampValue(computeSteppedValue(fieldValue, direction, isInteger)));
            onValueChanged();
        };

        inputField.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
        inputField.onChange = onValueChanged;

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        stepperGroup.upButton.helpTip = getLabel(isInteger ? "tooltip.stepUpInteger" : "tooltip.stepUp");
        stepperGroup.downButton.helpTip = getLabel(isInteger ? "tooltip.stepDownInteger" : "tooltip.stepDown");
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は SETTINGS_STORE_* / createSettingsStore / readSettingsLegacyFile / readSettingsLegacyPreference / settingsStore*
    // 2. 寿命は今のスクリプトに合わせて選ぶ。
    //      "session"    … $.global に置く。Illustrator を終了するまで残る。#targetengine が必須（無いと毎回消える）
    //      "persistent" … Folder.userData/illustrator-scripts/<storeName>.json に書く。再起動しても残る
    //    storeName はふつう SCRIPT_NAME。ダイアログの位置は DialogPosition の部品が持つので、ここには入れない
    // 3. 既定値を1か所にまとめ、load で受け取る。戻り値は毎回新しいオブジェクト（書き換えても保存されない）
    //      var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
    //      var DEFAULT_SETTINGS = { widthPt: 10, addFrame: true, modeKey: "fit", corners: { tl: 0, tr: 0 } };
    //      var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
    //      …OK で閉じたら…
    //      settingsStore.save({ widthPt: …, addFrame: …, modeKey: …, corners: { tl: …, tr: … } });
    //    型は既定値に合わせる（数値の既定値には "12" も 12 として読む。真偽は "1"/"0"/"true"/"false" も読む）。
    //    合わない値・既定値に無い項目は捨てて既定値を使う。{} と null の既定値は中身を問わずそのまま受け取る
    //    （名前をキーにしたプリセット集など）。配列は配列ならそのまま受け取る
    // 4. 保存できるのは文字列・数値・真偽・null と、その配列・入れ子のオブジェクトだけ。
    //    DOM オブジェクト・File・関数は入れない（パスは fsName の文字列で持つ）。長さは pt で持つ
    // 5. 旧形式の設定を読み継ぐときは、3つ目の引数に legacy 関数を渡す。
    //    新しい保存が1度も無いとき（ファイルが無い・$.global に無い）だけ呼ばれ、戻り値を保存値として既定値と突き合わせる。
    //    旧ファイル・旧キーは消さない。キー名が変わったときは legacy の中で詰め替える
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");  … key=value / toSource / JSON を自動判別
    //      } });
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyPreference("SmartTextFindReplace/settings");  … app.preferences の文字列
    //      } });
    // 6. clear() は保存を消す。legacy を渡したストアでは空の保存（{}）を書き、旧設定が戻ってこないようにする
    // 7. 失敗は例外にせず、load は既定値、save は false を返す（$.writeln に理由を出す）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^﻿/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 保存先は Folder.userData/illustrator-scripts/TrimWithBreakLine.json。
       旧版（v1.2.1 まで）の settings.txt は最初の1回だけ読み継ぐ（キー名 warpStyle → warpStyleKey）
       Stored in Folder.userData/illustrator-scripts/TrimWithBreakLine.json; the old settings.txt is read once */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent", {
        legacy: function () {
            var legacyValues = readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");
            if (legacyValues && legacyValues.warpStyle !== undefined) legacyValues.warpStyleKey = legacyValues.warpStyle;
            return legacyValues;
        }
    });

    /**
     * 次回のために設定を保存する（長さはptで持つ）
     * @param {object} buildSettings - 確定した設定
     * @returns {void}
     */
    function saveSettings(buildSettings) {
        settingsStore.save({
            maskScale: buildSettings.maskScale,
            maskCrossScale: buildSettings.maskCrossScale,
            maskOffsetPt: buildSettings.maskOffsetPt,
            warpStyleKey: buildSettings.warpStyleKey,
            warpPercent: buildSettings.warpPercent,
            jaggedSizePt: buildSettings.jaggedSizePt,
            jaggedRidges: buildSettings.jaggedRidges,
            gapPt: buildSettings.gapPt,
            addRule: !!buildSettings.addRule,
            ruleDashed: !!buildSettings.ruleDashed,
            dashSegments: buildSettings.dashSegments,
            ruleWidthPt: buildSettings.ruleWidth,
            roundCap: !!buildSettings.roundCap,
            groupRules: !!buildSettings.groupRules
        });
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

    /* 帯が片側の端まで覆っているときは、その側だけを残して反対側を切り落とす
       When the band reaches one end, keep that side only and trim the rest away */
    var coversAxisMax = (bandAxisMax >= targetAxisMax - TOLERANCE);
    var coversAxisMin = (bandAxisMin <= targetAxisMin + TOLERANCE);
    var keepOneSide = (coversAxisMax || coversAxisMin);

    if (coversAxisMax && coversAxisMin) {
        alert(getLabel((axisIndex === AXIS_Y) ? "alert.bandCoversAllY" : "alert.bandCoversAllX"));
        return;
    }

    var parentContainer = targetItem.parent;

    /* 切る方向と直交する幅は、最前面の図形（マスク用のパス）に合わせ、設定で伸縮する。対象はこの幅にトリミングされる
       The path in front sets the size across the cut, scaled by the settings; the artwork is trimmed to it */
    var bandCrossMin = getCrossMin(bandBounds);
    var bandCrossMax = getCrossMax(bandBounds);
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
    var savedValues = settingsStore.load({
        maskScale: DEFAULT_MASK_SCALE,
        maskCrossScale: DEFAULT_MASK_CROSS_SCALE,
        maskOffsetPt: DEFAULT_MASK_OFFSET,
        warpStyleKey: DEFAULT_WARP_STYLE,
        warpPercent: DEFAULT_WARP_PERCENT,
        jaggedSizePt: DEFAULT_JAGGED_SIZE,
        jaggedRidges: DEFAULT_JAGGED_RIDGES,
        gapPt: DEFAULT_GAP_MM * UNITS[1].pointsPerUnit,
        addRule: DEFAULT_ADD_RULE,
        ruleDashed: DEFAULT_RULE_DASHED,
        dashSegments: DEFAULT_DASH_SEGMENTS,
        ruleWidthPt: DEFAULT_RULE_WIDTH,
        roundCap: DEFAULT_ROUND_CAP,
        groupRules: DEFAULT_GROUP_RULES
    });
    var initialValues = {
        maskScale:    savedValues.maskScale,
        maskCrossScale: savedValues.maskCrossScale,
        maskOffset:   savedValues.maskOffsetPt / rulerUnit.pointsPerUnit,
        warpStyleKey: savedValues.warpStyleKey,
        warpPercent:  savedValues.warpPercent,
        jaggedSize:   savedValues.jaggedSizePt / rulerUnit.pointsPerUnit,
        jaggedRidges: savedValues.jaggedRidges,
        gap:          savedValues.gapPt / rulerUnit.pointsPerUnit,
        addRule:      savedValues.addRule,
        ruleDashed:   savedValues.ruleDashed,
        dashSegments: savedValues.dashSegments,
        ruleWidth:    savedValues.ruleWidthPt / strokeUnit.pointsPerUnit,
        roundCap:     savedValues.roundCap,
        groupRules:   savedValues.groupRules
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
     * ワープの設定を返す
     * @param {string} warpStyleKey - WARP_STYLES のキー
     * @returns {object} ワープの設定（未知のキーは初期スタイル）
     */
    function getWarpStyle(warpStyleKey) {
        return WARP_STYLES[warpStyleKey] || WARP_STYLES[DEFAULT_WARP_STYLE];
    }

    /**
     * パスにワープ効果を適用する
     * @param {PathItem} targetPath - 効果を適用するパス
     * @param {string} warpStyleKey - WARP_STYLES のキー（"flag" / "rise" / "riseStraight"。"jagged" は対象外）
     * @param {number} warpPercent - カーブの量（%）
     * @returns {void}
     */
    function applyWarp(targetPath, warpStyleKey, warpPercent) {
        var warpStyle = getWarpStyle(warpStyleKey);
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
     * 切る方向と、切り口に沿った方向の座標から点を作る
     * @param {number} axisValue - 切る方向の座標
     * @param {number} crossValue - 切り口に沿った方向の座標
     * @returns {number[]} [x, y]
     */
    function toAnchor(axisValue, crossValue) {
        return (axisIndex === AXIS_Y) ? [crossValue, axisValue] : [axisValue, crossValue];
    }

    /**
     * パスの末尾に、ハンドルのないコーナーの点を足す
     * @param {PathItem} targetPath - 対象のパス
     * @param {number[]} anchor - 点の座標
     * @returns {void}
     */
    function addCornerPoint(targetPath, anchor) {
        var cornerPoint = targetPath.pathPoints.add();
        cornerPoint.anchor = anchor;
        cornerPoint.leftDirection = anchor;
        cornerPoint.rightDirection = anchor;
        cornerPoint.pointType = PointType.CORNER;
    }

    /**
     * 切り口だけがギザギザのマスクを作る
     * 矩形に掛けると横の辺までギザギザになるので、切り口の線にだけジグザグを掛けてから反対側の角を足して閉じる
     * @param {number} axisMax - 切る方向の、座標が大きいほうの端（反対側の辺）
     * @param {number} axisSize - 切る方向の大きさ
     * @param {{crossMin: number, crossMax: number}} crossRange - 切り口に沿った方向の範囲
     * @param {number} jaggedSizePt - ギザギザの大きさ（pt）
     * @param {number} jaggedRidges - 折り返し
     * @returns {PageItem} マスクに使うパス
     */
    function createJaggedMaskPath(axisMax, axisSize, crossRange, jaggedSizePt, jaggedRidges) {
        var cutAxis = axisMax - axisSize;
        var edgeLine = parentContainer.pathItems.add();
        edgeLine.setEntirePath([toAnchor(cutAxis, crossRange.crossMin), toAnchor(cutAxis, crossRange.crossMax)]);
        edgeLine.closed = false;
        edgeLine.filled = false;
        /* 線がないと分割で何も残らない / an unpainted path leaves nothing to expand */
        edgeLine.stroked = true;
        edgeLine.strokeColor = createGrayColor(100);
        edgeLine.applyEffect(createZigzagXml(jaggedSizePt, jaggedRidges));

        var maskPath = expandAppearance(edgeLine);
        maskPath.stroked = false;
        maskPath.filled = true;
        maskPath.fillColor = createGrayColor(100);

        /* 分割で点の向きが逆になっても辺が交差しないよう、末尾の点に近い角から足す
           add the nearer far corner first so the outline never crosses itself */
        var pathPoints = maskPath.pathPoints;
        var crossIndex = 1 - axisIndex;
        var endsAtCrossMax = pathPoints[pathPoints.length - 1].anchor[crossIndex] > pathPoints[0].anchor[crossIndex];
        addCornerPoint(maskPath, toAnchor(axisMax, endsAtCrossMax ? crossRange.crossMax : crossRange.crossMin));
        addCornerPoint(maskPath, toAnchor(axisMax, endsAtCrossMax ? crossRange.crossMin : crossRange.crossMax));
        maskPath.closed = true;
        return maskPath;
    }

    /**
     * マスク用の矩形を作り、カーブの量に応じてワープを掛けて分割する
     * @param {number} axisMax - 切る方向の、座標が大きいほうの端
     * @param {number} axisSize - 切る方向の大きさ
     * @param {{crossMin: number, crossMax: number}} crossRange - 切り口に沿った方向の範囲
     * @param {object} buildSettings - { warpStyleKey, warpPercent, jaggedSizePt, jaggedRidges }
     * @returns {PageItem} マスクに使うパス
     */
    function createMaskPath(axisMax, axisSize, crossRange, buildSettings) {
        var warpStyleKey = buildSettings.warpStyleKey;
        var warpPercent = buildSettings.warpPercent;
        var isJagged = getWarpStyle(warpStyleKey).jagged;
        if (isJagged && buildSettings.jaggedSizePt > 0 && buildSettings.jaggedRidges > 0) {
            return createJaggedMaskPath(axisMax, axisSize, crossRange, buildSettings.jaggedSizePt, buildSettings.jaggedRidges);
        }

        var crossSize = crossRange.crossMax - crossRange.crossMin;
        /* rectangle(top, left, width, height) は常に上端・左端で指定する */
        var maskRect = (axisIndex === AXIS_Y) ?
            parentContainer.pathItems.rectangle(axisMax, crossRange.crossMin, crossSize, axisSize) :
            parentContainer.pathItems.rectangle(crossRange.crossMax, axisMax - axisSize, axisSize, crossSize);
        /* 塗りがないとワープを分割できない / the warp needs a filled path to expand */
        maskRect.filled = true;
        maskRect.fillColor = createGrayColor(100);
        maskRect.stroked = false;

        /* ギザギザの大きさか折り返しが0のときもまっすぐに切る / a zero-size jagged edge is straight */
        if (isJagged || warpPercent === 0) return maskRect;

        applyWarp(maskRect, warpStyleKey, warpPercent);
        /* 直線のスタイルは、ワープのカーブをジグザグで直線に置き換えてから分割する */
        if (getWarpStyle(warpStyleKey).straight) maskRect.applyEffect(createZigzagXml(0, 0));
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
     * 切り口を対象の内側に収める
     * @param {number} axisValue - 切り口の座標
     * @returns {number} 収めた座標
     */
    function clampCutToTarget(axisValue) {
        if (axisValue > targetAxisMax - MIN_BAND_MARGIN) return targetAxisMax - MIN_BAND_MARGIN;
        if (axisValue < targetAxisMin + MIN_BAND_MARGIN) return targetAxisMin + MIN_BAND_MARGIN;
        return axisValue;
    }

    /**
     * 大きさ・位置の設定を反映した帯の範囲を返す
     * 描いたマスク用の図形を、設定で伸縮・移動した範囲が取り除く部分になる
     * @param {object} buildSettings - { maskScale, maskOffsetPt, ... }
     * @returns {{cutMax: number, cutMin: number}} 切り口の座標
     */
    function getCutRange(buildSettings) {
        var bandCenter = (bandAxisMax + bandAxisMin) / 2 + offsetSign * buildSettings.maskOffsetPt;
        var bandHalfSize = (bandAxisMax - bandAxisMin) * buildSettings.maskScale / 200;
        return {
            cutMax: clampCutToTarget(bandCenter + bandHalfSize),
            cutMin: clampCutToTarget(bandCenter - bandHalfSize)
        };
    }

    /**
     * 大きさの設定を反映した、切り口に沿った方向の範囲を返す
     * 描いたマスク用の図形の中心を基準に伸縮する。対象はこの範囲にトリミングされる
     * @param {object} buildSettings - { maskCrossScale, ... }
     * @returns {{crossMin: number, crossMax: number}} 範囲の両端の座標
     */
    function getCrossRange(buildSettings) {
        var crossCenter = (bandCrossMax + bandCrossMin) / 2;
        var crossHalfSize = (bandCrossMax - bandCrossMin) * buildSettings.maskCrossScale / 200;
        return { crossMin: crossCenter - crossHalfSize, crossMax: crossCenter + crossHalfSize };
    }

    /**
     * 罫線を切り口の上に出す
     * @param {PathItem[]} rulePaths - 移動する罫線
     * @returns {void}
     */
    function bringRulesToFront(rulePaths) {
        for (var ruleIndex = 0; ruleIndex < rulePaths.length; ruleIndex++) {
            rulePaths[ruleIndex].move(parentContainer, ElementPlacement.PLACEATBEGINNING);
        }
    }

    /**
     * 帯が覆っている側だけを残す（反対側は切り落とす）
     * 切り口は帯の内側の辺、反対側の辺は対象の端に切りそろえる
     * @param {object} buildSettings - buildParts と同じ設定
     * @param {{cutMax: number, cutMin: number}} cutRange - 切り口の座標
     * @param {{crossMin: number, crossMax: number}} crossRange - 切り口に沿った方向の範囲
     * @returns {PageItem[]} 作成したクリップグループと罫線
     */
    function buildSinglePart(buildSettings, cutRange, crossRange) {
        /* 覆われていない側にある帯の辺で切る / cut at the band edge on the side it does not cover */
        var cutAxis = coversAxisMax ? cutRange.cutMin : cutRange.cutMax;

        var maskPath = createMaskPath(cutAxis + maskAxisSize, maskAxisSize, crossRange, buildSettings);
        alignCutEdge(maskPath, cutAxis);

        /* 罫線は切りそろえる前の辺から取り出す / take the rule before the far edge is flattened */
        var rulePaths = buildSettings.addRule ? createEdgeRules(maskPath, buildSettings) : [];
        flattenFarEdge(maskPath, coversAxisMax ? targetAxisMax : targetAxisMin);

        var part = createClippedPart(maskPath);
        bringRulesToFront(rulePaths);

        if (buildSettings.groupRules) return [groupWithRules(part, rulePaths)];
        return [part].concat(rulePaths);
    }

    /**
     * 帯を取り除いたパーツを作る
     * 両側に切り分けるときは、上側は下辺を帯の上端に、下側は同じ複製の下辺を帯の下端に合わせ、
     * それぞれ反対側の辺を画像の上端・下端に切りそろえる
     * @param {object} buildSettings - { maskScale, maskCrossScale, maskOffsetPt, warpStyleKey, warpPercent, jaggedSizePt, jaggedRidges, gapPt, addRule, ruleDashed, dashSegments, ruleWidth, roundCap, groupRules }
     * @returns {PageItem[]} 作成したクリップグループと罫線
     */
    function buildParts(buildSettings) {
        var cutRange = getCutRange(buildSettings);
        var crossRange = getCrossRange(buildSettings);
        if (keepOneSide) return buildSinglePart(buildSettings, cutRange, crossRange);

        var cutMax = cutRange.cutMax;
        var cutMin = cutRange.cutMin;

        var upperMaskPath = createMaskPath(cutMax + maskAxisSize, maskAxisSize, crossRange, buildSettings);

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

        /* 片側を寄せて、指定の間隔をあける。上下は上側、左右は左側を動かさず、寄せる側の罫線も一緒に動かす
           close up to the gap, keeping the top part (Y) or the left part (X) in place */
        var closeDistance = cutMax - cutMin - buildSettings.gapPt;
        var moveUpperPart = (axisIndex === AXIS_X);
        var movingPart = moveUpperPart ? upperPart : lowerPart;
        var movingRules = moveUpperPart ? upperRules : lowerRules;
        var partShift = moveUpperPart ? -closeDistance : closeDistance;
        translateAlongAxis(movingPart, partShift);

        for (var ruleIndex = 0; ruleIndex < movingRules.length; ruleIndex++) {
            translateAlongAxis(movingRules[ruleIndex], partShift);
        }

        /* 罫線は切り口の上に出す / bring the rules in front of the parts */
        var rulePaths = upperRules.concat(lowerRules);
        bringRulesToFront(rulePaths);

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
     * @param {object} buildSettings - { maskScale, maskCrossScale, maskOffsetPt, warpStyleKey, warpPercent, jaggedSizePt, jaggedRidges, gapPt, addRule, ruleDashed, dashSegments, ruleWidth, roundCap, groupRules }
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 「マスク用の図形」パネルを組み立てる
     * @param {Group} parentGroup - 追加先の列グループ
     * @returns {object} パネルの入力コントロール
     */
    function buildMaskPanel(parentGroup) {
        var maskPanel = addPanel(parentGroup, "panel.mask", "tooltip.maskPanel");

        var isVerticalCut = (axisIndex === AXIS_Y);
        var offsetLabelPath = isVerticalCut ? "fieldLabel.maskOffsetY" : "fieldLabel.maskOffsetX";
        var offsetTooltipPath = isVerticalCut ? "tooltip.maskOffsetY" : "tooltip.maskOffsetX";
        var scaleText = formatFieldNumber(initialValues.maskScale);
        var crossScaleText = formatFieldNumber(initialValues.maskCrossScale);

        /* 幅・高さの順に並べる。上下に切るときは幅、左右に切るときは高さが切り口に沿った方向になる
           width comes before height; which one runs along the cut edge depends on the direction */
        var scaleRow, crossScaleRow;
        if (isVerticalCut) {
            crossScaleRow = addNumberFieldRow(maskPanel, "fieldLabel.maskWidth", "tooltip.maskCrossScale", crossScaleText, "%");
            scaleRow = addNumberFieldRow(maskPanel, "fieldLabel.maskHeight", "tooltip.maskScale", scaleText, "%");
        } else {
            scaleRow = addNumberFieldRow(maskPanel, "fieldLabel.maskWidth", "tooltip.maskScale", scaleText, "%");
            crossScaleRow = addNumberFieldRow(maskPanel, "fieldLabel.maskHeight", "tooltip.maskCrossScale", crossScaleText, "%");
        }
        var offsetRow = addNumberFieldRow(maskPanel, offsetLabelPath, offsetTooltipPath,
            formatFieldNumber(initialValues.maskOffset), rulerUnit.label);

        return { scaleField: scaleRow.field, crossScaleField: crossScaleRow.field, offsetField: offsetRow.field };
    }

    /**
     * ワープのキーが、アイコンボタンの何番目かを返す
     * @param {string} warpStyleKey - WARP_STYLES のキー
     * @returns {number} アイコンボタンの位置（見つからなければ0）
     */
    function getWarpStyleIndex(warpStyleKey) {
        for (var styleIndex = 0; styleIndex < WARP_STYLE_KEYS.length; styleIndex++) {
            if (WARP_STYLE_KEYS[styleIndex] === warpStyleKey) return styleIndex;
        }
        return 0;
    }


    /**
     * IllustratorのUIが明るいテーマか判定する
     * @returns {boolean} 明るいテーマなら true
     */
    function isLightUI() {
        return app.preferences.getRealPreference("uiBrightness") > 0.5;
    }

    /* アイコンボタンの配色（UIの明暗で切り替える）/ icon button colors for the light and dark UI */
    var STYLE_ICON_COLORS = isLightUI() ? {
        shape: [0.55, 0.55, 0.55, 1], selectedShape: [0.13, 0.12, 0.11, 1],
        frame: [0.2, 0.5, 0.95, 1]
    } : {
        shape: [0.55, 0.55, 0.55, 1], selectedShape: [0.92, 0.92, 0.92, 1],
        frame: [0.3, 0.6, 1, 1]
    };

    /**
     * 切り口の形のアイコンに描く、切り口の線の点を返す（0〜1の比率、Yは下向き）
     * @param {string} warpStyleKey - WARP_STYLES のキー
     * @returns {number[][]} 点の並び
     */
    function getStyleIconCutPoints(warpStyleKey) {
        var cutPoints = [];
        var stepCount = 24;
        var stepIndex;
        var ratioX;
        if (warpStyleKey === "flag") {
            /* 1周期の波 / one period of a wave */
            for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                ratioX = stepIndex / stepCount;
                cutPoints.push([ratioX, 0.5 + 0.16 * Math.sin(ratioX * Math.PI * 2)]);
            }
        } else if (warpStyleKey === "rise") {
            /* 左から右へせり上がるS字 / an S-curve rising to the right */
            for (stepIndex = 0; stepIndex <= stepCount; stepIndex++) {
                ratioX = stepIndex / stepCount;
                cutPoints.push([ratioX, 0.68 - 0.36 * (1 - Math.cos(ratioX * Math.PI)) / 2]);
            }
        } else if (warpStyleKey === "riseStraight") {
            cutPoints.push([0, 0.6], [1, 0.42]);
        } else {
            /* ギザギザ / zigzag */
            var ridgeCount = 6;
            for (stepIndex = 0; stepIndex <= ridgeCount; stepIndex++) {
                cutPoints.push([stepIndex / ridgeCount, (stepIndex % 2 === 0) ? 0.44 : 0.56]);
            }
        }
        return cutPoints;
    }

    /**
     * 切り口の線の、指定のX（0〜1）でのYと傾きを返す（点の間は直線でつなぐ）
     * @param {number[][]} cutPoints - 切り口の線の点（Xの昇順）
     * @param {number} ratioX - X（0〜1）
     * @returns {{y: number, slope: number}} Y（0〜1）と傾き
     */
    function sampleCutLine(cutPoints, ratioX) {
        for (var pointIndex = 1; pointIndex < cutPoints.length; pointIndex++) {
            var startPoint = cutPoints[pointIndex - 1];
            var endPoint = cutPoints[pointIndex];
            if (ratioX > endPoint[0] && pointIndex < cutPoints.length - 1) continue;
            var slope = (endPoint[1] - startPoint[1]) / (endPoint[0] - startPoint[0]);
            return { y: startPoint[1] + (ratioX - startPoint[0]) * slope, slope: slope };
        }
        return { y: cutPoints[0][1], slope: 0 };
    }

    /**
     * 切り口の形のアイコンボタンを描く（背景は塗らない）
     * 多角形は塗れないので、1px幅の縦の短冊を、切り口の上と下に分けて塗る
     * @param {Button} iconButton - 対象のボタン（warpStyleKey と isSelected を持つ）
     * @returns {void}
     */
    function drawStyleIconButton(iconButton) {
        var graphics = iconButton.graphics;
        var buttonWidth = iconButton.size[0];
        var buttonHeight = iconButton.size[1];
        var isSelected = iconButton.isSelected;
        var shapeBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR,
            isSelected ? STYLE_ICON_COLORS.selectedShape : STYLE_ICON_COLORS.shape);

        var originX = Math.round((buttonWidth - STYLE_ICON_SIZE) / 2);
        var originY = Math.round((buttonHeight - STYLE_ICON_SIZE) / 2);
        var cutPoints = getStyleIconCutPoints(iconButton.warpStyleKey);

        for (var columnIndex = 0; columnIndex < STYLE_ICON_SIZE; columnIndex++) {
            var cutSample = sampleCutLine(cutPoints, (columnIndex + 0.5) / STYLE_ICON_SIZE);
            /* 斜めでも切り口の太さがそろうよう、縦の幅を傾きで広げる / widen the vertical gap on slopes */
            var halfGap = STYLE_ICON_CUT_WIDTH / 2 * Math.sqrt(1 + cutSample.slope * cutSample.slope);
            var cutY = cutSample.y * STYLE_ICON_SIZE;
            var upperHeight = cutY - halfGap;
            var lowerTop = cutY + halfGap;

            if (upperHeight > 0) {
                graphics.newPath();
                graphics.rectPath(originX + columnIndex, originY, 1, upperHeight);
                graphics.fillPath(shapeBrush);
            }
            if (lowerTop < STYLE_ICON_SIZE) {
                graphics.newPath();
                graphics.rectPath(originX + columnIndex, originY + lowerTop, 1, STYLE_ICON_SIZE - lowerTop);
                graphics.fillPath(shapeBrush);
            }
        }

        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(1, 1, buttonWidth - 2, buttonHeight - 2);
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, STYLE_ICON_COLORS.frame, 2));
        }
    }

    /**
     * 切り口の形を選ぶアイコンボタンを横に並べる（ラジオボタンの代わり）
     * @param {Window|Group} parentGroup - 追加先
     * @param {number} selectedIndex - 初期選択の位置（WARP_STYLE_KEYS 上）
     * @param {function} onValueChanged - 選択が変わったあとに呼ぶ処理
     * @returns {{row: Group, getSelectedKey: function}} 追加した行と、選択中のキーを返す関数
     */
    function addStyleIconRow(parentGroup, selectedIndex, onValueChanged) {
        var iconRow = addFieldRow(parentGroup);
        iconRow.spacing = STYLE_ICON_SPACING;
        iconRow.margins = [0, 0, 0, STYLE_ICON_BOTTOM_MARGIN];

        var iconButtons = [];

        /**
         * 指定の位置のアイコンを選択状態にして描き直す
         * @param {number} targetIndex - 選ぶ位置
         * @returns {void}
         */
        function selectIconAt(targetIndex) {
            for (var buttonIndex = 0; buttonIndex < iconButtons.length; buttonIndex++) {
                iconButtons[buttonIndex].isSelected = (buttonIndex === targetIndex);
                iconButtons[buttonIndex].notify("onDraw");
            }
        }

        for (var styleIndex = 0; styleIndex < WARP_STYLE_KEYS.length; styleIndex++) {
            var warpStyleKey = WARP_STYLE_KEYS[styleIndex];
            var iconButton = iconRow.add("button", undefined, "");
            iconButton.preferredSize = [STYLE_ICON_BUTTON_SIZE, STYLE_ICON_BUTTON_SIZE];
            iconButton.helpTip = labelText("radio." + warpStyleKey) + getLabel("tooltip." + warpStyleKey);
            iconButton.warpStyleKey = warpStyleKey;
            iconButton.styleIndex = styleIndex;
            iconButton.isSelected = (styleIndex === selectedIndex);
            iconButton.onDraw = function () {
                drawStyleIconButton(this);
            };
            iconButton.onClick = function () {
                if (this.isSelected) return;
                selectIconAt(this.styleIndex);
                onValueChanged();
            };
            iconButtons.push(iconButton);
        }

        return {
            row: iconRow,
            getSelectedKey: function () {
                for (var buttonIndex = 0; buttonIndex < iconButtons.length; buttonIndex++) {
                    if (iconButtons[buttonIndex].isSelected) return iconButtons[buttonIndex].warpStyleKey;
                }
                return DEFAULT_WARP_STYLE;
            }
        };
    }

    /**
     * 「切り口」パネルを組み立てる
     * @param {Group} parentGroup - 追加先の列グループ
     * @param {function} onStyleChanged - 切り口の形が変わったあとに呼ぶ処理
     * @returns {object} パネルの入力コントロール
     */
    function buildCutEdgePanel(parentGroup, onStyleChanged) {
        var cutEdgePanel = addPanel(parentGroup, "panel.cutEdge", "tooltip.cutEdgePanel");

        /* 形は項目名なしのアイコンで横に並べ、パネルの左右中央に置く / the styles sit in a centered row of icons without a label */
        var warpStyleRow = addStyleIconRow(cutEdgePanel, getWarpStyleIndex(initialValues.warpStyleKey), onStyleChanged);
        warpStyleRow.row.alignment = ["center", "center"];
        var warpAmountRow = addNumberFieldRow(cutEdgePanel, "fieldLabel.warpAmount", "tooltip.warpAmount",
            formatFieldNumber(initialValues.warpPercent), "%");
        var jaggedSizeRow = addNumberFieldRow(cutEdgePanel, "fieldLabel.jaggedSize", "tooltip.jaggedSize",
            formatFieldNumber(initialValues.jaggedSize), rulerUnit.label);
        var jaggedRidgesRow = addNumberFieldRow(cutEdgePanel, "fieldLabel.jaggedRidges", "tooltip.jaggedRidges",
            formatFieldNumber(initialValues.jaggedRidges), "");
        var gapRow = addNumberFieldRow(cutEdgePanel, "fieldLabel.gap", "tooltip.gap",
            formatFieldNumber(initialValues.gap), rulerUnit.label);

        /* 片側だけ残すときはパーツがひとつなので、間隔は使わない */
        gapRow.row.enabled = !keepOneSide;

        return {
            getStyleKey: warpStyleRow.getSelectedKey,
            warpAmountRow: warpAmountRow.row,
            warpAmountField: warpAmountRow.field,
            jaggedSizeRow: jaggedSizeRow.row,
            jaggedSizeField: jaggedSizeRow.field,
            jaggedRidgesRow: jaggedRidgesRow.row,
            jaggedRidgesField: jaggedRidgesRow.field,
            gapField: gapRow.field
        };
    }

    /**
     * 「省略線」パネルを組み立てる
     * @param {Group} parentGroup - 追加先の列グループ
     * @returns {object} パネルの入力コントロール
     */
    function buildBreakLinePanel(parentGroup) {
        var breakLinePanel = addPanel(parentGroup, "panel.breakLine", "tooltip.breakLinePanel");

        var addRuleCheckbox = addCheckboxRow(breakLinePanel, "checkbox.addRule", "tooltip.addRule", initialValues.addRule);
        var strokeWidthRow = addNumberFieldRow(breakLinePanel, "fieldLabel.strokeWidth", "tooltip.strokeWidth",
            formatFieldNumber(initialValues.ruleWidth), strokeUnit.label, RULE_LABEL_WIDTH);
        var strokeCapRow = addRadioRow(breakLinePanel, "fieldLabel.strokeCap", "tooltip.strokeCap",
            ["radio.buttCap", "radio.roundCap"], initialValues.roundCap ? 1 : 0, RULE_LABEL_WIDTH, true);
        var ruleStyleRow = addRadioRow(breakLinePanel, "fieldLabel.ruleStyle", "tooltip.ruleStyle",
            ["radio.solid", "radio.dashed"], initialValues.ruleDashed ? 1 : 0, RULE_LABEL_WIDTH, true);
        var dashSegmentsRow = addNumberFieldRow(breakLinePanel, "fieldLabel.dashSegments", "tooltip.dashSegments",
            formatFieldNumber(initialValues.dashSegments), "", RULE_LABEL_WIDTH);
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

        /* 左に「マスク用の図形」と「切り口」、右に「省略線」を置く
           mask and cut-edge settings on the left, break lines on the right */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumn = addColumn(columnsGroup);
        var rightColumn = addColumn(columnsGroup);

        var maskControls = buildMaskPanel(leftColumn);
        var cutEdgeControls = buildCutEdgePanel(leftColumn, onSettingChanged);
        var ruleControls = buildBreakLinePanel(rightColumn);
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 罫線まわりの行の有効・無効を切り替える
         * @returns {void}
         */
        function updateRuleRows() {
            var addRule = ruleControls.addRuleCheckbox.value;
            ruleControls.styleRow.enabled = addRule;
            setRowEnabled(ruleControls.segmentsRow, addRule && ruleControls.dashedRadio.value);
            setRowEnabled(ruleControls.widthRow, addRule);
            ruleControls.capRow.enabled = addRule;
            ruleControls.groupRulesCheckbox.enabled = addRule;
        }

        /**
         * 選んだ切り口の形に合わせて、カーブとギザギザの行を切り替える
         * @returns {void}
         */
        function updateCutEdgeRows() {
            var isJagged = getWarpStyle(cutEdgeControls.getStyleKey()).jagged === true;
            setRowEnabled(cutEdgeControls.warpAmountRow, !isJagged);
            setRowEnabled(cutEdgeControls.jaggedSizeRow, isJagged);
            setRowEnabled(cutEdgeControls.jaggedRidgesRow, isJagged);
        }

        /**
         * ダイアログの入力内容を読み取る（値は欄にそろえ直す）
         * @returns {object} 設定
         */
        function collectSettings() {
            var gapValue = readFieldValue(cutEdgeControls.gapField, clampGapValue);
            return {
                maskScale: readFieldValue(maskControls.scaleField, clampMaskScale),
                maskCrossScale: readFieldValue(maskControls.crossScaleField, clampMaskCrossScale),
                maskOffsetPt: readFieldValue(maskControls.offsetField, clampOffsetValue) * rulerUnit.pointsPerUnit,
                warpStyleKey: cutEdgeControls.getStyleKey(),
                warpPercent: readFieldValue(cutEdgeControls.warpAmountField, clampWarpPercent),
                jaggedSizePt: readFieldValue(cutEdgeControls.jaggedSizeField, clampJaggedSize) * rulerUnit.pointsPerUnit,
                jaggedRidges: readFieldValue(cutEdgeControls.jaggedRidgesField, clampJaggedRidges),
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
            updateCutEdgeRows();
            updateRuleRows();
            refreshPreview(collectSettings());
        }

        wireNumberField(maskControls.scaleField, clampMaskScale, onSettingChanged);
        wireNumberField(maskControls.crossScaleField, clampMaskCrossScale, onSettingChanged);
        wireNumberField(maskControls.offsetField, clampOffsetValue, onSettingChanged);
        wireNumberField(cutEdgeControls.warpAmountField, clampWarpPercent, onSettingChanged);
        wireNumberField(cutEdgeControls.jaggedSizeField, clampJaggedSize, onSettingChanged);
        wireNumberField(cutEdgeControls.jaggedRidgesField, clampJaggedRidges, onSettingChanged, true);
        wireNumberField(cutEdgeControls.gapField, clampGapValue, onSettingChanged);
        wireNumberField(ruleControls.segmentsField, clampDashSegments, onSettingChanged, true);
        wireNumberField(ruleControls.widthField, clampRuleWidth, onSettingChanged);
        wireClickControls([ruleControls.addRuleCheckbox]
            .concat(ruleControls.styleRadios)
            .concat(ruleControls.capRadios)
            .concat([ruleControls.groupRulesCheckbox]), onSettingChanged);

        var dialogSettings = null;
        btnOK.onClick = function () {
            dialogSettings = collectSettings();
            dialog.close(1);
        };
        btnCancel.onClick = function () {
            dialog.close(0);
        };

        /* プレビューで選択を外す前に選択範囲を測る / Measure the selection before the preview clears it */
        prepareDialogWindow(dialog, SCRIPT_NAME);

        /* 開いた時点のプレビューを先に出す / Show the preview before the dialog appears */
        updateCutEdgeRows();
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
