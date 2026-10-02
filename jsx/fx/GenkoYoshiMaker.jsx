#target illustrator
#targetengine "GenkoYoshiMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストを1文字＝1マスに組み直し、マスの区切りと行の両側の罫線、十字線を引きます。縦組み・横組み、ポイント文字・エリア内文字のいずれにも対応し、組み方向は変換しません。
比率・自動行送り・罫線の濃度・太罫・十字線などはダイアログで指定し、結果はその場でプレビューされます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenkoYoshiMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n036c34760079

### Overview

Sets the selected text one character to a cell, then draws the cell rules, the rules along both
sides of each line, and the crosshairs. Vertical and horizontal, point and area text are all
supported, and the writing direction is left as it is. A dialog sets the scale, the auto leading,
the rule density, the emphasis and the crosshairs, and previews the result live.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GenkoYoshiMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GenkoYoshiMaker";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenkoYoshiMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GenkoYoshiMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n036c34760079"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 言語とラベル
    // Language and labels
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrame: { ja: "テキストオブジェクトを1つ選択してください。", en: "Please select a single text object." },
            noCharacter: { ja: "テキストに文字が入力されていません。", en: "The selected text is empty." },
            noFontSize: { ja: "文字サイズを取得できませんでした。", en: "Could not read the font size." },
            skipped: { ja: "設定できなかった属性：", en: "Attributes that could not be set:" },
            presetName: { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
            presetSaved: {
                ja: "プリセットを追加しました（このセッションのみ）。\n\nPRESETS に貼り付けると残せます:\n",
                en: "Preset added for this session.\n\nPaste this into PRESETS to keep it:\n"
            }
        },
        itemName: {
            layerName: { ja: "原稿用紙罫線", en: "Manuscript Grid" },
            groupName: { ja: "罫線", en: "Rules" }
        },
        dialog: {
            title: { ja: "原稿用紙メーカー", en: "Manuscript Paper Maker" }
        },
        panel: {
            overall: { ja: "全体", en: "Overall" },
            text: { ja: "文字", en: "Characters" },
            /* 行の両側に添える罫線。縦組みでは縦罫、横組みでは横罫 */
            lineRule: {
                vertical: { ja: "縦罫", en: "Vertical rules" },
                horizontal: { ja: "横罫", en: "Horizontal rules" }
            },
            /* マスの区切りの罫線。縦組みでは横罫、横組みでは縦罫 */
            cellRule: {
                vertical: { ja: "横罫", en: "Horizontal rules" },
                horizontal: { ja: "縦罫", en: "Vertical rules" }
            },
            cross: { ja: "十字線", en: "Crosshairs" }
        },
        label: {
            preset: { ja: "プリセット", en: "Preset" },
            scale: { ja: "比率", en: "Scale" },
            autoLeading: { ja: "自動行送り", en: "Auto leading" },
            extension: { ja: "はみ出し", en: "Extension" },
            lineColor: { ja: "罫線の濃度", en: "Rule density" },
            emphasisRatio: { ja: "太罫の比率", en: "Emphasis ratio" },
            emphasisTarget: { ja: "太罫の対象", en: "Emphasis applies to" },
            extraCells: { ja: "マスを追加", en: "Add cells" },
            extraLines: { ja: "行を追加", en: "Add lines" },
            emphasis: { ja: "強調", en: "Emphasis" },
            /* 太罫の対象のうち、マスの区切りに付ける補足 */
            emphasisTargetSuffix: { ja: "（n文字ごと）", en: " (at the interval)" },
            lineStyle: { ja: "線種", en: "Style" },
            segments: { ja: "分割数", en: "Segments" },
            emphasisPrefix: { ja: "", en: "Thicker every" },
            emphasisSuffix: { ja: "文字ごとに太く", en: "characters" }
        },
        checkbox: {
            cellRects: { ja: "文字ごとに長方形を作成", en: "Draw a rectangle for each character" },
            adjustSize: { ja: "文字の比率を調整", en: "Adjust the character scale" },
            sideRule: { ja: "いちばん外に1本追加", en: "Add one more outside" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        preset: {
            standard: { ja: "標準", en: "Standard" },
            noSideRule: { ja: "外側の罫線なし", en: "No outer rules" },
            noCross: { ja: "十字線なし", en: "No crosshairs" },
            noCrossWithMargin: { ja: "十字線なし（余白つき）", en: "No crosshairs, with margin" },
            rulesOnly: { ja: "シンプル", en: "Simple" },
            rectangles: { ja: "長方形のマス", en: "Rectangle cells" },
            custom: { ja: "カスタム", en: "Custom" }
        },
        radio: {
            none: { ja: "なし", en: "None" },
            solid: { ja: "実線", en: "Solid" },
            dashed: { ja: "破線", en: "Dashed" }
        },
        tooltip: {
            preset: {
                ja: "設定をまとめて切り替えます。値を手で変えると「カスタム」になります。",
                en: "Fills the whole dialog at once. Editing any value switches to Custom."
            },
            cellRects: {
                ja: "1マスにつき長方形を1つ作ります。マスの区切りと行の両側の罫線は引きません。罫線の濃度・線種と、十字線の設定はそのまま使います。",
                en: "Creates one rectangle per cell instead of the cell rules and the rules along each line. The rule density, the cell-rule style and the crosshairs still apply."
            },
            lineColor: {
                ja: "罫線の濃度。shiftを押しながらドラッグすると10%刻みになります。十字線の濃度は別（30%）です。",
                en: "Density of the rules. Hold Shift to snap to 10% steps. The crosshairs keep their own density (30%)."
            },
            emphasisRatio: {
                ja: "太くする罫線の太さ。ふつうの罫線（0.1mm）に対する比率で指定します。",
                en: "Weight of the emphasized rules, as a ratio of the normal ones (0.1mm)."
            },
            emphasisTarget: {
                ja: "太罫の比率を使う罫線を選びます。行の両側の罫線と、n文字ごとに太くする区切りが対象です。",
                en: "Picks which rules use the emphasis ratio: the rules along each line, and the cell rules thickened at the given interval."
            },
            extraCells: {
                ja: "文字の前後（上下）に足す空のマスの数。",
                en: "Empty cells added before and after the text."
            },
            extraLines: {
                ja: "テキストの行の外側に足す、空の行の数。マスの1/4だけアキを空けて並べます。",
                en: "Empty lines added outside the text, spaced by a quarter of a cell."
            },
            adjustSize: {
                ja: "文字の比率を変えてマスに合わせます。オフでも、送り（トラッキング・ツメ・アキ・カーニング）は今の比率に合わせてそろえます。",
                en: "Scales the characters to the cells. When off, the advance is still matched to the current scaling."
            },
            scale: {
                ja: "文字の水平比率・垂直比率。下げた分はトラッキングで送りに戻します。",
                en: "Horizontal and vertical scale of the characters; the tracking gives the advance back."
            },
            autoLeading: {
                ja: "行と行の送り。文字サイズに対する比率で、マス1つ分（100%）を超えた分が行と行のアキになります。外側の罫線もこのアキの位置に引かれます。",
                en: "Distance from line to line as a percentage of the font size; anything over one cell (100%) becomes the gutter, and the outer rules follow it."
            },
            extension: {
                ja: "縦罫をマスの上下へ伸ばす量。",
                en: "How far the vertical rules run past the cells."
            },
            sideRule: {
                ja: "いちばん外側の行の外へ、マスの1/4だけ離して縦罫を足します。",
                en: "Adds a vertical rule a quarter of a cell outside the outermost line."
            },
            emphasis: {
                ja: "指定した文字数ごとに横罫を太くします。",
                en: "Thickens the horizontal rule at the given interval."
            },
            cellRuleStyle: {
                ja: "マスの区切りを実線と破線から選びます。",
                en: "Draws the cell rules solid or dashed."
            },
            crossStyle: {
                ja: "各マスの中央に引く十字線の線種です。「なし」にすると引きません。",
                en: "Style of the crosshairs at the center of every cell; None draws nothing."
            },
            crossSegments: {
                ja: "十字線1本あたりの線分の本数。0か3以上の奇数で、線分と間隔は同じ長さになります。",
                en: "Dashes per crosshair line: zero or an odd number of 3 or more; the dash and the gap are equal."
            },
            preview: {
                ja: "ダイアログを開いたまま、確定後の状態を表示します。",
                en: "Shows the finished result while the dialog stays open."
            },
            savePreset: {
                ja: "いまの設定に名前を付けてプリセットへ加えます。次回も使うには、表示されるコードを PRESETS に貼り付けてください。",
                en: "Adds the current settings to the preset list. Paste the code it shows into PRESETS to keep it for next time."
            },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        unit: {
            mm: { ja: "mm", en: "mm" },
            percent: { ja: "%", en: "%" },
            characters: { ja: "文字", en: "characters" },
            lines: { ja: "行", en: "lines" }
        },
        button: {
            add: { ja: "追加", en: "Add" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        }
    };

    /**
     * 書字方向に合うほうのラベル定義を取り出す
     * @param {object} labelEntry vertical / horizontal を持つラベル定義
     * @param {boolean} isVertical 縦組みかどうか
     * @return {object} ja / en を持つラベル定義
     */
    function orientedEntry(labelEntry, isVertical) {
        return isVertical ? labelEntry.vertical : labelEntry.horizontal;
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

    /* 原稿用紙の寸法は常に mm で扱う / manuscript dimensions are always handled in mm */
    var POINTS_PER_MM = UNITS[1].pointsPerUnit;

    /**
     * ミリメートルをポイントに変換する
     * @param {number} mm ミリメートル
     * @return {number} ポイント
     */
    function mmToPt(mm) {
        return mm * POINTS_PER_MM;
    }

    /**
     * ポイントをミリメートルに変換する
     * @param {number} pt ポイント
     * @return {number} ミリメートル
     */
    function ptToMm(pt) {
        return pt / POINTS_PER_MM;
    }

    /**
     * 文字サイズから伸張の既定値を求める（文字サイズの1/4、小数第1位に丸める）
     * @param {number} fontSize 文字サイズ（pt）
     * @return {number} 伸張（mm）
     */
    function getDefaultExtensionMM(fontSize) {
        return Math.round(ptToMm(fontSize / 4) * 10) / 10;
    }

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* プレビュー用レイヤー名（一時的に作って消す）/ Name of the temporary preview layer */
    var PREVIEW_LAYER_NAME = "__ManuscriptGrid_Preview";

    /* 旧版が前回の設定を置いていた $.global のキー（1度だけ読み継ぐ）
       / The $.global key older versions used for the last settings (read once) */
    var LEGACY_SESSION_KEY = "__GenkoYoshiMakerSession";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var LAYOUT = {
        cellStrokeMM: 0.1,          /* マスの区切りの太さ（mm）/ weight of the cell rules (mm) */
        cellDashMM: [1, 1],         /* マスの区切りを破線にしたときの線分と間隔（mm）/ dash and gap of the dashed cell rules (mm) */
        crossStrokeMM: 0.1,         /* 十字線の太さ（mm）/ crosshair weight (mm) */
        crossGray: 30,              /* 十字線の濃度（%、100で黒）/ crosshair density (%, 100 is black) */
        gridStartOffset: 0          /* 1本目の位置の微調整（pt、プラスで書字方向の手前へ）/ nudge of the first rule (pt) */
    };

    /* ダイアログの初期値 / Dialog defaults */
    var DEFAULTS = {
        cellRects: false,        /* マスを1つずつ長方形にするか / draw one rectangle per cell instead of the rules */
        strokeGray: 50,          /* 罫線の濃度（%、100で黒）/ rule density (%, 100 is black) */
        emphasisRatio: 250,      /* 太罫の太さ（ふつうの罫線に対する%）/ weight of the emphasized rules (% of the normal ones) */
        emphasizeLineRules: true,/* 行の両側の罫線を太罫にするか / use the emphasis ratio for the rules along each line */
        emphasizeCellRules: true,/* n文字ごとの区切りを太罫にするか / use the emphasis ratio for the interval cell rules */
        extraCells: 0,           /* 文字の前後に足すマスの数 / empty cells added before and after the text */
        extraLines: 0,           /* テキストの行の外側に足す、空の行の数 / empty lines added outside the text */
        adjustSize: true,        /* 文字を比率とトラッキングでマスに合わせるか / fit the characters to the cells */
        scale: 90,               /* 水平・垂直比率（%）/ horizontal and vertical scale (%) */
        autoLeading: 125,        /* 自動行送り（%）。100%を超えた分が行と行のアキ / auto leading (%); anything over 100% is the gutter */
        extensionMM: null,       /* 罫線を書字方向へ伸ばす量（mm）。null は文字サイズの1/4 / null means a quarter of the font size */
        dashedCellRules: false,  /* マスの区切りを破線にするか / draw the cell rules dashed */
        emphasisEnabled: true,   /* 一定間隔のマスの区切りを太くするか / thicken the cell rules at a fixed interval */
        emphasisEvery: 5,        /* 太くする間隔（文字数）/ interval of the thickened rules (characters) */
        showSideRules: true,     /* いちばん外に罫線をもう1本添えるか / add one more rule outside */
        crossStyle: "dashed",    /* 十字線の線種（none / solid / dashed）/ crosshair style */
        crossSegments: 9,        /* 十字線1本あたりの線分の本数（線分と間隔は同じ長さ）/ dashes per crosshair line (dash and gap are equal) */
        preview: true            /* プレビューを表示するか / show the live preview */
    };

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // 自動サイズ調整のアクション / Auto-size action
    // =========================================

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /* エリア内文字の自動サイズ調整は DOM から設定できないため、
       その場で作ったアクションを実行する（AreaTypeToolkit.jsx 参考）
       / Auto Size cannot be set from the DOM, so a temporary action is loaded and run */
    var AUTO_SIZE_ACTION_SET = "GenkoYoshiMaker_AutoSize";
    var AUTO_SIZE_ACTION_NAME = "AutoSizeOn";

    /**
     * アクション名のブロックを組み立てる
     * @param {string} name アクション名
     * @return {string} /name [ <長さ> <16進> ] のブロック
     */
    function buildActionNameBlock(name) {
        var nameHex = toActionHex(name).toUpperCase();
        return "/name [ " + (nameHex.length / 2) + " " + nameHex + " ]";
    }

    /**
     * 自動サイズ調整のアクション定義（.aia 文字列）を組み立てる
     * @return {string} アクションセットの定義
     */
    function buildAutoSizeAia() {
        /* 1行の文字列にすると読みづらいので、配列にして join する */
        var definition = [
            "/version 3",
            buildActionNameBlock(AUTO_SIZE_ACTION_SET),
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            " " + buildActionNameBlock(AUTO_SIZE_ACTION_NAME),
            " /keyIndex 0",
            " /colorIndex 0",
            " /isOpen 1",
            " /eventCount 1",
            " /event-1 {",
            " /useRulersIn1stQuadrant 0",
            " /internalName (adobe_SLOAreaTextDialog)",
            " /localizedName [ 33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3 ]",
            " /isOpen 0",
            " /isOn 1",
            " /hasDialog 0",
            " /parameterCount 1",
            " /parameter-1 {",
            " /key 1952539754",
            " /showInPalette 4294967295",
            " /type (integer)",
            " /value 1",
            " }",
            " }",
            "}"
        ];

        return definition.join("");
    }

    /**
     * エリア内文字の枠を文字に合わせる（自動サイズ調整）
     * ダイアログを開いたままアクションを実行すると不安定なため、開く前と確定時だけ呼ぶ
     * @param {Document} doc 対象のドキュメント
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {void}
     */
    function fitAreaTextFrame(doc, textFrame) {
        if (!isAreaText(textFrame)) {
            return;
        }

        doc.selection = [textFrame];

        try {
            app.doScript(AUTO_SIZE_ACTION_NAME, AUTO_SIZE_ACTION_SET, false);
        } catch (actionError) {
            /* アクションを実行できない環境では枠のままにする / Leave the frame as it is when the action cannot run */
        }

        app.redraw();
    }

    // 設定の保存（再利用パーツ） / Settings store (reusable)

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
            return textFile.read().replace(/^\uFEFF/, "");
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
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
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

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    // =========================================
    // セッションの記憶 / Session memory
    // =========================================

    /* 前回の設定は Illustrator を終了するまで残す（#targetengine が必須）
       / The last settings last until Illustrator quits (needs #targetengine) */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_KEY] || null; }
    });

    /* values はプリセットと同じ形の設定値（null は未保存）/ values has the preset shape (null = nothing saved) */
    var DEFAULT_SESSION = { values: null, presetIndex: 0 };

    /**
     * 前回の設定を取り出す（同じIllustratorのセッション内だけ）
     * @return {object|null} { values, presetIndex }。まだ無いときは null
     */
    function loadSession() {
        var session = settingsStore.load(DEFAULT_SESSION);
        return settingsStoreIsPlainObject(session.values) ? session : null;
    }

    /**
     * 今回の設定を覚えておく（同じIllustratorのセッション内だけ）
     * @param {object} values プリセットと同じ形の設定値
     * @param {number} presetIndex 選んでいたプリセットの位置
     * @return {void}
     */
    function saveSession(values, presetIndex) {
        settingsStore.save({ values: values, presetIndex: presetIndex });
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /* ダイアログの設定をまとめて切り替える定義。settings が null のものは「カスタム」
       / Each preset fills the dialog at once; the one with null settings is "Custom" */
    var PRESETS = [
        { label: LABELS.preset.standard, settings: DEFAULTS },
        {
            label: LABELS.preset.noSideRule,
            settings: {
                cellRects: false, strokeGray: 50, emphasisRatio: 250, emphasizeLineRules: true, emphasizeCellRules: true, extraCells: 0, extraLines: 0, adjustSize: true, scale: 90, autoLeading: 125, extensionMM: null, showSideRules: false,
                dashedCellRules: false, emphasisEnabled: true, emphasisEvery: 5,
                crossStyle: "dashed", crossSegments: 9
            }
        },
        {
            label: LABELS.preset.noCross,
            settings: {
                cellRects: false, strokeGray: 50, emphasisRatio: 250, emphasizeLineRules: true, emphasizeCellRules: true, extraCells: 0, extraLines: 0, adjustSize: true, scale: 90, autoLeading: 125, extensionMM: null, showSideRules: true,
                dashedCellRules: false, emphasisEnabled: true, emphasisEvery: 5,
                crossStyle: "none", crossSegments: 9
            }
        },
        {
            label: LABELS.preset.noCrossWithMargin,
            settings: {
                cellRects: false, strokeGray: 50, emphasisRatio: 250, emphasizeLineRules: true, emphasizeCellRules: true, extraCells: 1, extraLines: 1, adjustSize: true, scale: 90, autoLeading: 125, extensionMM: null, showSideRules: true,
                dashedCellRules: false, emphasisEnabled: true, emphasisEvery: 5,
                crossStyle: "none", crossSegments: 9
            }
        },
        {
            label: LABELS.preset.rulesOnly,
            settings: {
                cellRects: false, strokeGray: 50, emphasisRatio: 250, emphasizeLineRules: true, emphasizeCellRules: true, extraCells: 0, extraLines: 0, adjustSize: true, scale: 90, autoLeading: 125, extensionMM: 0, showSideRules: false,
                dashedCellRules: false, emphasisEnabled: false, emphasisEvery: 5,
                crossStyle: "none", crossSegments: 9
            }
        },
        {
            label: LABELS.preset.rectangles,
            settings: {
                cellRects: true, strokeGray: 50, emphasisRatio: 250, emphasizeLineRules: true, emphasizeCellRules: true, extraCells: 0, extraLines: 0, adjustSize: true, scale: 100, autoLeading: 125, extensionMM: null, showSideRules: true,
                dashedCellRules: false, emphasisEnabled: true, emphasisEvery: 5,
                crossStyle: "none", crossSegments: 9
            }
        },
        { label: LABELS.preset.custom, settings: null }
    ];

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

    /* ダイアログの寸法 / Dialog metrics */
    var UI = {
        labelWidth: (uiLang === "ja") ? 90 : 130,  /* 項目名の幅（px）/ width of the row labels (px) */
        saveButtonWidth: 70                        /* ［追加］ボタンの幅（px）/ width of the Add button (px) */
    };

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

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
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
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
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // =========================================
    // 文字属性 / Character attributes
    // =========================================

    /**
     * テキスト全体に文字属性を1つ設定する
     * 取得済みの characterAttributes を使い回すと Error 9503 や片側無視が起きるため、毎回取り直す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {string} attributeName 属性名
     * @param {number|boolean|object} value 設定する値
     * @return {boolean} 設定できたかどうか
     */
    function setCharacterAttribute(textFrame, attributeName, value) {
        try {
            textFrame.textRange.characterAttributes[attributeName] = value;
            return true;
        } catch (attributeError) {
            return false;
        }
    }

    /**
     * 段落ごとに文字属性を1つ設定する
     * 文字ツメ・プロポーショナルメトリクス・自動カーニングは、テキスト全体へ書いても効かない
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {string} attributeName 属性名
     * @param {number|boolean|object} value 設定する値
     * @return {boolean} 設定できたかどうか
     */
    function setCharacterAttributeByParagraph(textFrame, attributeName, value) {
        try {
            var paragraphs = textFrame.textRange.paragraphs;

            for (var i = 0; i < paragraphs.length; i++) {
                paragraphs[i].characterAttributes[attributeName] = value;
            }

            return true;
        } catch (attributeError) {
            return false;
        }
    }

    /**
     * 段落ごとに段落属性を1つ設定する
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {string} attributeName 属性名
     * @param {number|boolean|object} value 設定する値
     * @return {boolean} 設定できたかどうか
     */
    function setParagraphAttribute(textFrame, attributeName, value) {
        try {
            var paragraphs = textFrame.textRange.paragraphs;

            for (var i = 0; i < paragraphs.length; i++) {
                paragraphs[i].paragraphAttributes[attributeName] = value;
            }

            return true;
        } catch (attributeError) {
            return false;
        }
    }

    /**
     * 手動カーニングを全文字で0に戻す
     * カーニングは文字ごとに持つため、範囲へまとめて書いても残りが消えない
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {boolean} 設定できたかどうか
     */
    function clearManualKerning(textFrame) {
        try {
            var characters = textFrame.textRange.characters;

            for (var i = 0; i < characters.length; i++) {
                characters[i].kerning = 0;
            }

            return true;
        } catch (kerningError) {
            return false;
        }
    }

    /**
     * 1行ぶんの送りを求める（マス1つ＋行と行のアキ）
     * @param {number} cellSize マスの一辺（pt）
     * @param {number} autoLeading 自動行送り（%）
     * @return {number} 行送り（pt）
     */
    function getLinePitch(cellSize, autoLeading) {
        return cellSize * autoLeading / 100;
    }

    /**
     * 行と行のアキを求める（行送りからマス1つ分を引いた残り）
     * @param {number} cellSize マスの一辺（pt）
     * @param {number} autoLeading 自動行送り（%）
     * @return {number} アキ（pt）
     */
    function getLineGutter(cellSize, autoLeading) {
        return getLinePitch(cellSize, autoLeading) - cellSize;
    }

    /**
     * 比率を下げた分の文字送りを打ち消すトラッキングを求める
     * @param {number} scale 水平・垂直比率（%）
     * @return {number} トラッキング（1/1000em）
     */
    function getTracking(scale) {
        return 100000 / scale - 1000;
    }

    /**
     * 1文字＝1マスになるよう文字属性をそろえる（1文字目のカーニングは実測後に入れる）
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @param {number} scale 水平・垂直比率（%）
     * @return {array} 設定できなかった属性名の配列
     */
    function alignCharactersToCells(textFrame, settings, scale) {
        var applyScale = settings.adjustSize;
        var failed = [];

        /* 文字前後のアキは「自動」（-1）。0は「アキなし」で別物
           / Aki before and after is set to auto (-1); zero would mean "no aki" */
        var rangeSettings = [
            ["tracking", getTracking(scale)],
            ["akiLeft", -1],
            ["akiRight", -1]
        ];

        /* 比率を変えない設定でも、送りは今の比率に合わせてそろえる
           / Even when the scaling stays, the advance is still matched to it */
        if (applyScale) {
            rangeSettings.push(["horizontalScale", scale]);
            rangeSettings.push(["verticalScale", scale]);
        }

        for (var i = 0; i < rangeSettings.length; i++) {
            if (!setCharacterAttribute(textFrame, rangeSettings[i][0], rangeSettings[i][1])) {
                failed.push(rangeSettings[i][0]);
            }
        }

        /* 文字ツメ・プロポーショナルメトリクス・自動カーニングは段落単位で書く
           自動カーニングは「和文等幅」。手動カーニングは1文字目だけ別に入れる
           / Tsume, proportional metrics and auto kerning are written per paragraph */
        var paragraphSettings = [
            ["Tsume", 0],
            ["proportionalMetrics", false],
            ["kerningMethod", AutoKernType.NOAUTOKERN],
            /* 行送りは自動行送りに任せ、その比率をマス＋アキに合わせる
               / The leading is left on auto, with its percentage set to one cell plus the gutter */
            ["autoLeading", true]
        ];

        for (var j = 0; j < paragraphSettings.length; j++) {
            if (!setCharacterAttributeByParagraph(textFrame, paragraphSettings[j][0], paragraphSettings[j][1])) {
                failed.push(paragraphSettings[j][0]);
            }
        }

        if (!setParagraphAttribute(textFrame, "autoLeadingAmount", settings.autoLeading)) {
            failed.push("autoLeadingAmount");
        }

        /* 手動カーニングはいったん全文字0に戻す。1文字目だけ実測のあとに入れ直す
           / Clear the manual kerning first; the first character gets its value after the measuring */
        if (!clearManualKerning(textFrame)) {
            failed.push("kerning");
        }

        return failed;
    }

    /**
     * 各行の1文字目にトラッキングの1/4のカーニングを入れて、字面をマスの中央へ落とす
     * 実測より先に入れると罫線ごとずれるため、寸法を測ったあとに呼ぶ
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {number} scale 水平・垂直比率（%）
     * @return {boolean} 設定できたかどうか
     */
    function applyLineStartKerning(textFrame, scale) {
        var kerning = getTracking(scale) / 4;

        try {
            var lines = textFrame.lines;

            for (var i = 0; i < lines.length; i++) {
                var firstCharacter = lines[i].characters[0];

                /* カーニングは characterAttributes ではなく文字そのものが持つ。単位はトラッキングと同じ1/1000em
                   / Kerning lives on the character itself, not on characterAttributes, in the same 1/1000 em unit */
                firstCharacter.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                firstCharacter.kerning = kerning;
            }

            return true;
        } catch (kerningError) {
            return false;
        }
    }

    /**
     * 罫線に合わせる比率を決める（比率を変えない設定のときは、いまの比率を使う）
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @return {number} 水平・垂直比率（%）
     */
    function getEffectiveScale(textFrame, settings) {
        if (settings.adjustSize) {
            return settings.scale;
        }

        /* 送りは書字方向の比率で決まる / The advance follows the scale along the writing direction */
        var characterAttributes = textFrame.textRange.characters[0].characterAttributes;
        var currentScale = isVerticalText(textFrame) ? characterAttributes.verticalScale : characterAttributes.horizontalScale;

        return (currentScale > 0) ? currentScale : 100;
    }

    /**
     * 行揃えを書字方向の先頭（縦組みは上、横組みは左）にそろえる
     * 行揃えを変えると文字が動くので、書き出し位置を元に戻す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {boolean} 設定できたかどうか
     */
    function alignParagraphsToStart(textFrame) {
        var before = textFrame.geometricBounds;

        if (!setParagraphAttribute(textFrame, "justification", Justification.LEFT)) {
            return false;
        }

        app.redraw();

        var after = textFrame.geometricBounds;

        if (isVerticalText(textFrame)) {
            textFrame.translate(0, before[1] - after[1]);
        } else {
            textFrame.translate(before[0] - after[0], 0);
        }

        return true;
    }

    /**
     * 文字をマス目に合わせた状態にする（組み方向はそのまま）
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @param {number} scale 罫線に合わせる比率（%）
     * @return {array} 設定できなかった属性名の配列
     */
    function applyGridAttributes(textFrame, settings, scale) {
        var failed = alignCharactersToCells(textFrame, settings, scale);

        if (!alignParagraphsToStart(textFrame)) {
            failed.push("justification");
        }

        return failed;
    }

    // =========================================
    // 実測 / Measuring
    // =========================================

    /**
     * テキストが縦組みかどうかを返す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {boolean} 縦組みかどうか
     */
    function isVerticalText(textFrame) {
        return textFrame.orientation === TextOrientation.VERTICAL;
    }

    /**
     * テキストがエリア内文字かどうかを返す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {boolean} エリア内文字かどうか
     */
    function isAreaText(textFrame) {
        return textFrame.kind === TextType.AREATEXT;
    }

    /**
     * 行数と、1行あたりの最大文字数を数える
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {object} { lineCount, maxCharacters }（取得できないときは0）
     */
    function countGridCells(textFrame) {
        var lineCount = 0;
        var maxCharacters = 0;

        try {
            var lines = textFrame.lines;
            lineCount = lines.length;

            for (var i = 0; i < lines.length; i++) {
                var characterCount = lines[i].characters.length;

                if (characterCount > maxCharacters) {
                    maxCharacters = characterCount;
                }
            }
        } catch (lineError) {
            return { lineCount: 0, maxCharacters: 0 };
        }

        return { lineCount: lineCount, maxCharacters: maxCharacters };
    }

    /**
     * 罫線を引くのに必要な寸法を実測する（属性をそろえたあとに呼ぶ）
     * originX / originY は、1行目の1マス目の左上。cellCount は書字方向のマス数
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @return {object|null} { isVertical, originX, originY, cellSize, cellCount, lineCount }。文字サイズを取得できないときは null
     */
    function measureGrid(textFrame, settings) {
        /* 先頭文字の文字サイズを基準にする（比率を変えても size は元の値のまま） */
        var fontSize = textFrame.textRange.characters[0].characterAttributes.size;

        if (!fontSize || fontSize <= 0) {
            return null;
        }

        var isVertical = isVerticalText(textFrame);
        var bounds = textFrame.geometricBounds;
        var cellCounts = countGridCells(textFrame);

        var lineCount = (cellCounts.lineCount > 0) ? cellCounts.lineCount : 1;

        /* 最後の文字の先にも罫線を引くため、マス数は1行の文字数から求める */
        var cellCount = cellCounts.maxCharacters;

        if (cellCount <= 0) {
            var runLength = isVertical ? (bounds[1] - bounds[3]) : (bounds[2] - bounds[0]);
            cellCount = Math.ceil(runLength / fontSize);
        }

        /* マスは文字サイズの正方形。字面は比率を下げた分だけ狭いので、
           行が並ぶ向きの寸法は実測せず、行送り×（行数−1）＋マス1つで求めて中央にそろえる
           / Cells are squares of the font size; the glyphs are narrower, so the line axis is centered instead of measured */
        var lineSize = getLinePitch(fontSize, settings.autoLeading) * (lineCount - 1) + fontSize;

        var metrics = {
            isVertical: isVertical,
            cellSize: fontSize, /* マスの一辺。文字サイズそのもので、比率は反映しない */
            cellCount: cellCount,
            lineCount: lineCount
        };

        /* ポイント文字は字面の中央に、エリア内文字は枠の左（横組みは上）にそろえる
           / Point text is centered on the glyphs; area text follows the frame's left (top for horizontal) edge */
        var isArea = isAreaText(textFrame);

        if (isVertical) {
            metrics.originX = isArea ? bounds[0] : (bounds[0] + bounds[2]) / 2 - lineSize / 2;
            metrics.originY = bounds[1] + LAYOUT.gridStartOffset;
        } else {
            metrics.originX = bounds[0] - LAYOUT.gridStartOffset;
            metrics.originY = isArea ? bounds[1] : (bounds[1] + bounds[3]) / 2 + lineSize / 2;
        }

        return metrics;
    }

    /**
     * テキストをマス目に合わせ、罫線の寸法を測る
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @param {boolean} fitFrame エリア内文字の枠を合わせ直すか（ダイアログを開いている間は false）
     * @return {object} { metrics, failed } 実測した寸法と、設定できなかった属性名
     */
    function fitTextToGrid(textFrame, settings, fitFrame) {
        var scale = getEffectiveScale(textFrame, settings);
        var failed = applyGridAttributes(textFrame, settings, scale);

        /* 属性変更後の再組版を反映させてから境界を読む */
        app.redraw();

        if (fitFrame) {
            fitAreaTextFrame(app.activeDocument, textFrame);
        }

        var metrics = measureGrid(textFrame, settings);

        /* 行頭のカーニングは実測のあとに入れる（先に入れると罫線ごとずれてしまう） */
        if (metrics && !applyLineStartKerning(textFrame, scale)) {
            failed.push("kerning");
        }

        return { metrics: metrics, failed: failed };
    }

    // =========================================
    // 罫線 / Rules
    // =========================================

    /**
     * 罫線用のレイヤーを用意する（同名があれば再利用）
     * @param {Document} doc 対象のドキュメント
     * @param {string} layerName レイヤー名
     * @return {Layer} 罫線用のレイヤー
     */
    function prepareRuleLayer(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                var existingLayer = doc.layers[i];
                existingLayer.locked = false;
                existingLayer.visible = true;
                return existingLayer;
            }
        }

        var createdLayer = doc.layers.add();
        createdLayer.name = layerName;
        return createdLayer;
    }

    /**
     * その罫線グループが対象テキストのものか（テキストの中心を囲んでいるか）を調べる
     * 罫線はテキストより広いので、中心が入っていれば対象とみなせる
     * 同じレイヤーに別のテキストの罫線が並ぶため、名前だけでは見分けられない
     * @param {GroupItem} group 罫線グループ
     * @param {array} bounds 対象テキストの境界 [left, top, right, bottom]
     * @return {boolean} 対象テキストに引いた罫線かどうか
     */
    function coversTextCenter(group, bounds) {
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        var groupBounds = group.geometricBounds;

        return centerX >= groupBounds[0] && centerX <= groupBounds[2] &&
            centerY <= groupBounds[1] && centerY >= groupBounds[3];
    }

    /**
     * 対象テキストに前回引いた罫線グループだけ削除する（ほかのテキストの罫線は残す）
     * @param {Layer} layer 対象のレイヤー
     * @param {string} groupName グループ名
     * @param {array} bounds 対象テキストの境界 [left, top, right, bottom]
     * @return {void}
     */
    function removeExistingGroup(layer, groupName, bounds) {
        for (var i = layer.groupItems.length - 1; i >= 0; i--) {
            var group = layer.groupItems[i];

            if (group.name === groupName && coversTextCenter(group, bounds)) {
                group.remove();
            }
        }
    }

    /**
     * ドキュメントのカラーモードに合わせたグレーの色をつくる
     * @param {Document} doc 対象のドキュメント
     * @param {number} grayPercent 濃度（%、100で黒）
     * @return {object} GrayColor または RGBColor
     */
    function createGrayColor(doc, grayPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var rgbLevel = Math.round(255 * (100 - grayPercent) / 100);
            var rgbColor = new RGBColor();
            rgbColor.red = rgbLevel;
            rgbColor.green = rgbLevel;
            rgbColor.blue = rgbLevel;
            return rgbColor;
        }

        var grayColor = new GrayColor();
        grayColor.gray = grayPercent;
        return grayColor;
    }

    /**
     * 設定に合わせて罫線と十字線の色をつくる
     * @param {Document} doc 対象のドキュメント
     * @param {object} settings ダイアログで決めた設定
     * @return {object} { rule, cross } 罫線と十字線の色
     */
    function createRuleColors(doc, settings) {
        return {
            rule: createGrayColor(doc, settings.strokeGray),
            cross: createGrayColor(doc, LAYOUT.crossGray)
        };
    }

    /**
     * 罫線を1本引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {array} from 始点 [x, y]
     * @param {array} to 終点 [x, y]
     * @param {number} strokeWidth 線幅（pt）
     * @param {array} dashes 破線パターン（実線は空配列）
     * @param {object} color 罫線の色
     * @param {StrokeCap} strokeCap 線端（省略可）
     * @return {void}
     */
    function drawRule(group, from, to, strokeWidth, dashes, color, strokeCap) {
        var rulePath = group.pathItems.add();
        rulePath.setEntirePath([from, to]);

        rulePath.filled = false;
        rulePath.stroked = true;
        rulePath.strokeWidth = strokeWidth;
        rulePath.strokeDashes = dashes;
        rulePath.strokeColor = color;

        if (strokeCap) {
            rulePath.strokeCap = strokeCap;
        }
    }

    /**
     * 十字線の分割数を0か3以上の奇数に丸める
     * 偶数だとマスの中央が間隔になり、十字線が交差しない
     * @param {number} value 入力値
     * @param {boolean} roundUp 大きい側へ丸めるか
     * @return {number} 0、または3以上の奇数
     */
    function normalizeCrossSegments(value, roundUp) {
        var segments = Math.round(value);

        if (isNaN(segments) || segments <= 0) {
            return 0;
        }

        if (segments < 3) {
            return roundUp ? 3 : 0;
        }

        if (segments % 2 === 1) {
            return segments;
        }

        return roundUp ? segments + 1 : segments - 1;
    }

    /**
     * マス1辺の長さに合わせて、十字線の破線パターンを求める
     * 線分と間隔は同じ長さ。分割数＝線分の本数で、両端が線分で終わるように配分する
     * @param {object} settings ダイアログで決めた設定
     * @param {number} cellSize マスの一辺（pt）
     * @return {array} 破線パターン（実線は空配列）
     */
    function getCrossDashes(settings, cellSize) {
        /* 分割数0は破線にしない（実線として引く）/ Zero segments draws a solid line */
        if (settings.crossStyle !== "dashed" || settings.crossSegments < 1) {
            return [];
        }

        /* 線分が n 本、間隔が n−1 回で、合わせてマス1辺 / n dashes and n-1 gaps fill one cell */
        var dash = cellSize / (settings.crossSegments * 2 - 1);

        return [dash, dash];
    }

    /**
     * ブロック1つ分の各マスの中央に十字線を引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect ブロックの矩形 { left, right, top, bottom }
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 十字線の色
     * @return {void}
     */
    function drawCrosshairs(group, rect, grid, settings, color) {
        var cellSize = grid.cellSize;
        var strokeWidth = mmToPt(LAYOUT.crossStrokeMM);
        var dashes = getCrossDashes(settings, cellSize);

        var columnCount = Math.round((rect.right - rect.left) / cellSize);
        var rowCount = Math.round((rect.top - rect.bottom) / cellSize);

        for (var column = 0; column < columnCount; column++) {
            var cellLeft = rect.left + cellSize * column;
            var cellRight = cellLeft + cellSize;
            var centerX = cellLeft + cellSize / 2;

            for (var i = 0; i < rowCount; i++) {
                var cellTop = rect.top - cellSize * i;
                var cellBottom = cellTop - cellSize;
                var centerY = cellTop - cellSize / 2;

                drawRule(group, [centerX, cellTop], [centerX, cellBottom], strokeWidth, dashes, color);
                drawRule(group, [cellLeft, centerY], [cellRight, centerY], strokeWidth, dashes, color);
            }
        }
    }

    /**
     * 文字の前後に足すマスを反映した寸法を返す（書字方向の手前へ原点をずらす）
     * @param {object} metrics 実測した寸法
     * @param {object} settings ダイアログで決めた設定
     * @return {object} 前後にマスを足した寸法
     */
    function expandGrid(metrics, settings) {
        var startShift = metrics.cellSize * settings.extraCells;

        return {
            isVertical: metrics.isVertical,
            originX: metrics.isVertical ? metrics.originX : metrics.originX - startShift,
            originY: metrics.isVertical ? metrics.originY + startShift : metrics.originY,
            cellSize: metrics.cellSize,
            cellCount: metrics.cellCount + settings.extraCells * 2,
            lineCount: metrics.lineCount
        };
    }

    /**
     * 罫線を引く行の位置を並べる
     * テキストの行と、その外側に足す行を、1行ぶんの送りで等間隔に置く
     * 値は、行が並ぶ向きに測った原点からの距離（pt）
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @return {array} 行の位置（pt）の配列
     */
    function buildGridBlocks(grid, settings) {
        var linePositions = [];
        var linePitch = getLinePitch(grid.cellSize, settings.autoLeading);

        for (var i = 0; i < grid.lineCount; i++) {
            linePositions.push(linePitch * i);
        }

        for (var j = 1; j <= settings.extraLines; j++) {
            linePositions.push(-linePitch * j);
            linePositions.push(linePitch * (grid.lineCount - 1 + j));
        }

        return linePositions;
    }

    /**
     * 1行ぶんの矩形を求める
     * @param {object} grid 前後のマスを足した寸法
     * @param {number} linePos 行が並ぶ向きに測った原点からの距離（pt）
     * @return {object} { left, right, top, bottom }
     */
    function getBlockRect(grid, linePos) {
        var runSize = grid.cellSize * grid.cellCount;

        if (grid.isVertical) {
            var left = grid.originX + linePos;

            return { left: left, right: left + grid.cellSize, top: grid.originY, bottom: grid.originY - runSize };
        }

        var top = grid.originY - linePos;

        return { left: grid.originX, right: grid.originX + runSize, top: top, bottom: top - grid.cellSize };
    }

    /**
     * すべての行を囲む矩形を求める
     * @param {object} grid 前後のマスを足した寸法
     * @param {array} linePositions 行の位置の配列
     * @return {object} { left, right, top, bottom }
     */
    function getBlocksRect(grid, linePositions) {
        var rect = getBlockRect(grid, linePositions[0]);

        for (var i = 1; i < linePositions.length; i++) {
            var blockRect = getBlockRect(grid, linePositions[i]);

            if (blockRect.left < rect.left) {
                rect.left = blockRect.left;
            }

            if (blockRect.right > rect.right) {
                rect.right = blockRect.right;
            }

            if (blockRect.top > rect.top) {
                rect.top = blockRect.top;
            }

            if (blockRect.bottom < rect.bottom) {
                rect.bottom = blockRect.bottom;
            }
        }

        return rect;
    }

    /**
     * 罫線の太さを求める（太罫の対象なら比率を掛ける）
     * @param {object} settings ダイアログで決めた設定
     * @param {boolean} isEmphasized 太罫にするかどうか
     * @return {number} 線幅（pt）
     */
    function getRuleWidth(settings, isEmphasized) {
        var ruleWidth = mmToPt(LAYOUT.cellStrokeMM);

        return isEmphasized ? ruleWidth * settings.emphasisRatio / 100 : ruleWidth;
    }

    /**
     * マスの区切りの破線パターンを求める
     * @param {object} settings ダイアログで決めた設定
     * @return {array} 破線パターン（実線は空配列）
     */
    function getCellDashes(settings) {
        if (!settings.dashedCellRules) {
            return [];
        }

        return [mmToPt(LAYOUT.cellDashMM[0]), mmToPt(LAYOUT.cellDashMM[1])];
    }

    /**
     * ブロック1つ分のマスを、1マスにつき1つの長方形として作る
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect ブロックの矩形 { left, right, top, bottom }
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawCellRects(group, rect, grid, settings, color) {
        var cellSize = grid.cellSize;
        var strokeWidth = getRuleWidth(settings, false);
        var dashes = getCellDashes(settings);

        var columnCount = Math.round((rect.right - rect.left) / cellSize);
        var rowCount = Math.round((rect.top - rect.bottom) / cellSize);

        for (var column = 0; column < columnCount; column++) {
            var left = rect.left + cellSize * column;

            for (var i = 0; i < rowCount; i++) {
                var top = rect.top - cellSize * i;
                var cellRect = group.pathItems.rectangle(top, left, cellSize, cellSize);

                cellRect.filled = false;
                cellRect.stroked = true;
                cellRect.strokeWidth = strokeWidth;
                cellRect.strokeDashes = dashes;
                cellRect.strokeColor = color;
            }
        }
    }

    /**
     * マスの区切りを引く（縦組みでは横罫、横組みでは縦罫）
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect ブロックの矩形
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawCellRules(group, rect, grid, settings, color) {
        var cellWidth = getRuleWidth(settings, false);
        var emphasisWidth = getRuleWidth(settings, settings.emphasizeCellRules);
        var dashes = getCellDashes(settings);

        /* 誤差を溜めないよう、都度 原点から引いた位置を使う。
           太くする位置は、前後に足したマスではなく1文字目から数える
           / The emphasis is counted from the first character, not from the added cells */
        for (var i = 0; i <= grid.cellCount; i++) {
            var cellIndex = i - settings.extraCells;
            var isEmphasized = (settings.emphasisEvery > 0) && (cellIndex % settings.emphasisEvery === 0);
            var strokeWidth = isEmphasized ? emphasisWidth : cellWidth;

            if (grid.isVertical) {
                var y = rect.top - grid.cellSize * i;

                drawRule(group, [rect.left, y], [rect.right, y], strokeWidth, dashes, color);
            } else {
                var x = rect.left + grid.cellSize * i;

                drawRule(group, [x, rect.top], [x, rect.bottom], strokeWidth, dashes, color);
            }
        }
    }

    /**
     * 行の両側の罫線を引く（縦組みでは縦罫、横組みでは横罫）。書字方向の前後へ伸ばす
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect ブロックの矩形
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawLineRules(group, rect, grid, settings, color) {
        var strokeWidth = getRuleWidth(settings, settings.emphasizeLineRules);

        /* 突出線端にして、マスの角までしっかり届かせる / Projecting caps reach the corners of the cells */
        var strokeCap = StrokeCap.PROJECTINGENDCAP;

        if (grid.isVertical) {
            var top = rect.top + settings.extension;
            var bottom = rect.bottom - settings.extension;

            drawRule(group, [rect.left, top], [rect.left, bottom], strokeWidth, [], color, strokeCap);
            drawRule(group, [rect.right, top], [rect.right, bottom], strokeWidth, [], color, strokeCap);
            return;
        }

        var left = rect.left - settings.extension;
        var right = rect.right + settings.extension;

        drawRule(group, [left, rect.top], [right, rect.top], strokeWidth, [], color, strokeCap);
        drawRule(group, [left, rect.bottom], [right, rect.bottom], strokeWidth, [], color, strokeCap);
    }

    /**
     * いちばん外側の行の外へ、マスの1/4だけ離して罫線を添える
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect すべてのブロックを囲む矩形
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawSideRules(group, rect, grid, settings, color) {
        var sideOffset = getLineGutter(grid.cellSize, settings.autoLeading);
        var strokeWidth = getRuleWidth(settings, settings.emphasizeLineRules);
        var strokeCap = StrokeCap.PROJECTINGENDCAP;

        if (grid.isVertical) {
            var top = rect.top + settings.extension;
            var bottom = rect.bottom - settings.extension;

            drawRule(group, [rect.left - sideOffset, top], [rect.left - sideOffset, bottom], strokeWidth, [], color, strokeCap);
            drawRule(group, [rect.right + sideOffset, top], [rect.right + sideOffset, bottom], strokeWidth, [], color, strokeCap);
            return;
        }

        var left = rect.left - settings.extension;
        var right = rect.right + settings.extension;

        drawRule(group, [left, rect.top + sideOffset], [right, rect.top + sideOffset], strokeWidth, [], color, strokeCap);
        drawRule(group, [left, rect.bottom - sideOffset], [right, rect.bottom - sideOffset], strokeWidth, [], color, strokeCap);
    }

    /**
     * 罫線をひとそろい引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} metrics 実測した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} colors 罫線と十字線の色 { rule, cross }
     * @return {void}
     */
    function drawGrid(group, metrics, settings, colors) {
        var grid = expandGrid(metrics, settings);
        var linePositions = buildGridBlocks(grid, settings);

        for (var i = 0; i < linePositions.length; i++) {
            var rect = getBlockRect(grid, linePositions[i]);

            /* 十字線は罫線より先に引いて背面へ回す */
            if (settings.crossStyle !== "none") {
                drawCrosshairs(group, rect, grid, settings, colors.cross);
            }

            /* 長方形にするときは、マスの区切りも行の両側の罫線も長方形が兼ねる
               / The rectangles stand in for both the cell rules and the rules along each line */
            if (settings.cellRects) {
                drawCellRects(group, rect, grid, settings, colors.rule);
                continue;
            }

            drawCellRules(group, rect, grid, settings, colors.rule);
            drawLineRules(group, rect, grid, settings, colors.rule);
        }

        if (settings.showSideRules && !settings.cellRects) {
            drawSideRules(group, getBlocksRect(grid, linePositions), grid, settings, colors.rule);
        }
    }

    /**
     * レイヤーごと最背面へ送る
     * @param {Layer} layer 対象のレイヤー
     * @return {void}
     */
    function sendLayerToBack(layer) {
        try {
            layer.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (zOrderError) {
            layer.move(layer.parent, ElementPlacement.PLACEATEND);
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * 対象テキストに前回引いた罫線グループを集める（プレビュー中は隠して重ねないため）
     * @param {Document} doc 対象のドキュメント
     * @param {string} layerName レイヤー名
     * @param {string} groupName グループ名
     * @param {array} bounds 対象テキストの境界 [left, top, right, bottom]
     * @return {array} 表示中の罫線グループの配列
     */
    function findExistingGroups(doc, layerName, groupName, bounds) {
        var existingGroups = [];

        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];

            if (layer.name !== layerName || layer.locked) {
                continue;
            }

            for (var j = 0; j < layer.groupItems.length; j++) {
                var group = layer.groupItems[j];

                if (group.name === groupName && !group.hidden && coversTextCenter(group, bounds)) {
                    existingGroups.push(group);
                }
            }
        }

        return existingGroups;
    }

    /**
     * プレビュー用のレイヤーが残っていれば消す
     * @param {Document} doc 対象のドキュメント
     * @return {void}
     */
    function removePreviewLayer(doc) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name === PREVIEW_LAYER_NAME) {
                doc.layers[i].remove();
            }
        }
    }

    /**
     * 原本と複製の表示を入れ替える（プレビュー中は前回の罫線も隠す）
     * @param {object} previewState beginPreview が返した状態
     * @param {boolean} visible プレビューを見せるかどうか
     * @return {void}
     */
    function togglePreviewItems(previewState, visible) {
        previewState.textFrame.hidden = visible;
        previewState.previewFrame.hidden = !visible;

        for (var i = 0; i < previewState.existingGroups.length; i++) {
            previewState.existingGroups[i].hidden = visible;
        }
    }

    /**
     * プレビューの下ごしらえ（原本は隠し、確定後と同じ状態にした複製を代わりに見せる）
     * @param {Document} doc 対象のドキュメント
     * @param {TextFrame} textFrame 原本のテキストフレーム
     * @return {object} 片付けに使う状態
     */
    function beginPreview(doc, textFrame) {
        var previewState = {
            doc: doc,
            textFrame: textFrame,
            activeLayer: doc.activeLayer,
            existingGroups: findExistingGroups(doc, getLabel(LABELS.itemName.layerName), getLabel(LABELS.itemName.groupName), textFrame.geometricBounds),
            isVertical: isVerticalText(textFrame),
            previewFrame: textFrame.duplicate()
        };

        togglePreviewItems(previewState, true);

        return previewState;
    }

    /**
     * プレビュー用の複製を今の設定に合わせ直し、寸法を測り直す
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @param {object} settings ダイアログで決めた設定
     * @return {void}
     */
    function reshapePreviewFrame(previewContext, settings) {
        var previewState = previewContext.previewState;

        if (previewContext.appliedShape.adjustSize === settings.adjustSize && previewContext.appliedShape.scale === settings.scale) {
            return;
        }

        /* 文字属性は元へ戻せないので、比率を変えるたびに複製を作り直す
           / Character attributes cannot be rolled back, so rebuild the copy whenever the scaling changes */
        previewState.previewFrame.remove();
        previewState.previewFrame = previewState.textFrame.duplicate();
        previewState.previewFrame.hidden = false;

        var fitResult = fitTextToGrid(previewState.previewFrame, settings, false);

        if (fitResult.metrics) {
            previewContext.metrics = fitResult.metrics;
        }

        previewContext.appliedShape = { adjustSize: settings.adjustSize, scale: settings.scale };
    }

    /**
     * プレビューを片付けて元の状態に戻す（画面の更新はスクリプト終了時に任せる）
     * @param {object} previewState beginPreview が返した状態
     * @return {void}
     */
    function endPreview(previewState) {
        removePreviewLayer(previewState.doc);
        togglePreviewItems(previewState, false);
        previewState.previewFrame.remove();

        /* プレビューでレイヤーと選択が変わるので、控えておいた状態に戻す */
        previewState.doc.activeLayer = previewState.activeLayer;
        previewState.doc.selection = [previewState.textFrame];
    }

    /**
     * 現在の設定で罫線をプレビューする
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @param {object} settings ダイアログで決めた設定
     * @return {void}
     */
    function drawPreviewGrid(previewContext, settings) {
        var doc = previewContext.previewState.doc;

        togglePreviewItems(previewContext.previewState, true);
        reshapePreviewFrame(previewContext, settings);
        removePreviewLayer(doc);

        if (!previewContext.metrics) {
            app.redraw();
            return;
        }

        var previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;

        drawGrid(previewLayer.groupItems.add(), previewContext.metrics, settings, createRuleColors(doc, settings));

        /* 本処理と同じく、テキストの下に回す */
        sendLayerToBack(previewLayer);

        app.redraw();
    }

    /**
     * プレビューを消して、実行前の見た目に戻す
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @return {void}
     */
    function clearPreviewGrid(previewContext) {
        removePreviewLayer(previewContext.previewState.doc);
        togglePreviewItems(previewContext.previewState, false);

        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 濃度を表示用の文字列にする
     * @param {number} grayPercent 濃度（%、100で黒）
     * @return {string} 表示用の文字列
     */
    function formatDensity(grayPercent) {
        return Math.round(grayPercent) + "%";
    }

    /**
     * 項目名のラベルを1つ追加する
     * @param {object} parent 追加先のグループ
     * @param {object} labelEntry ja / en を持つラベル定義
     * @return {object} 追加した statictext
     */
    function addRowLabel(parent, labelEntry) {
        var label = parent.add("statictext", undefined, labelText(labelEntry));
        label.justify = "right";
        label.preferredSize.width = UI.labelWidth;
        return label;
    }

    /**
     * ダイアログのパネルを1つ追加する
     * @param {object} parent 追加先のウィンドウ
     * @param {object} labelEntry ja / en を持つラベル定義
     * @return {object} 追加したパネル
     */
    function addPanel(parent, labelEntry) {
        var panel = parent.add("panel", undefined, getLabel(labelEntry));
        setupPanel(panel, 6);
        return panel;
    }

    /**
     * ダイアログの1行分のグループを追加する
     * @param {object} parent 追加先のパネルまたはウィンドウ
     * @return {object} 追加したグループ
     */
    function addRow(parent) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["fill", "center"];
        row.alignChildren = ["left", "center"];
        return row;
    }

    /**
     * 複数のコントロールに同じツールチップを付ける
     * @param {array} targets 対象のコントロール
     * @param {object} tooltipEntry ツールチップのラベル定義（不要なときは null）
     * @return {void}
     */
    function setTooltip(targets, tooltipEntry) {
        if (!tooltipEntry) {
            return;
        }

        var tooltip = getLabel(tooltipEntry);

        for (var i = 0; i < targets.length; i++) {
            targets[i].helpTip = tooltip;
        }
    }

    /**
     * 項目名・数値入力欄・単位をまとめた行を追加する
     * @param {object} parent 追加先のパネル
     * @param {object} labelEntry ja / en を持つラベル定義
     * @param {number} initialValue 入力欄の初期値
     * @param {object} unitEntry 単位のラベル定義（単位がないときは null）
     * @param {object} tooltipEntry ツールチップのラベル定義（不要なときは null）
     * @param {object} stepOptions ∧∨の設定（min / integer。onStep はあとから bindDialogHandlers で入れる）
     * @return {object} { row, input } 追加した行と入力欄
     */
    function addNumberRow(parent, labelEntry, initialValue, unitEntry, tooltipEntry, stepOptions) {
        var row = addRow(parent);
        var label = addRowLabel(row, labelEntry);

        var input = addStepperField(row, String(initialValue), 5, stepOptions);

        if (unitEntry) {
            row.add("statictext", undefined, getLabel(unitEntry));
        }

        setTooltip([label, input], tooltipEntry);

        return { row: row, input: input };
    }

    /**
     * ∧∨と数値入力欄を隙間0で突き合わせて追加する
     * @param {object} parent 追加先の行
     * @param {string} initialText 入力欄の初期値
     * @param {number} characters 入力欄の文字数
     * @param {object} stepOptions ∧∨の設定（min / integer。onStep はあとから入れる）
     * @return {object} 追加した edittext（∧∨は .stepperGroup、設定は .stepOptions で参照できる）
     */
    function addStepperField(parent, initialText, characters, stepOptions) {
        var fieldGroup = parent.add("group");
        fieldGroup.orientation = "row";
        fieldGroup.alignChildren = ["left", "center"];
        fieldGroup.spacing = 0;
        fieldGroup.margins = 0;

        var input;
        /* 増減前の値を控えておく（十字線の分割数で、増減の向きを知るため） / remember the value before stepping */
        var stepperGroup = addStepper(fieldGroup, function () {
            input.textBeforeStep = input.text;
            return input;
        }, stepOptions);
        input = fieldGroup.add("edittext", undefined, initialText);
        input.characters = characters;
        input.stepperGroup = stepperGroup;
        input.stepOptions = stepOptions;
        return input;
    }

    /**
     * 項目名の幅だけ字下げしたチェックボックスの行を追加する
     * @param {object} parent 追加先のパネル
     * @param {object} labelEntry ja / en を持つラベル定義
     * @param {boolean} initialValue 初期値
     * @param {object} tooltipEntry ツールチップのラベル定義（不要なときは null）
     * @return {object} 追加した checkbox
     */
    function addIndentedCheckbox(parent, labelEntry, initialValue, tooltipEntry) {
        var row = addRow(parent);
        row.add("statictext", undefined, "").preferredSize.width = UI.labelWidth;

        var checkbox = row.add("checkbox", undefined, getLabel(labelEntry));
        checkbox.value = initialValue;

        setTooltip([checkbox], tooltipEntry);

        return checkbox;
    }

    /**
     * 項目名と排他ラジオをまとめた行を追加する
     * ラジオは同じ親の中でしか排他にならないため、1つのグループへまとめて入れる
     * @param {object} parent 追加先のパネル
     * @param {object} labelEntry ja / en を持つラベル定義
     * @param {array} radioEntries ラジオのラベル定義の配列
     * @param {number} selectedIndex 最初に選ぶラジオの位置
     * @param {object} tooltipEntry ツールチップのラベル定義（不要なときは null）
     * @return {object} { row, radios } 追加した行とラジオの配列
     */
    function addRadioRow(parent, labelEntry, radioEntries, selectedIndex, tooltipEntry) {
        var row = addRow(parent);
        var label = addRowLabel(row, labelEntry);

        var radioGroup = row.add("group");
        radioGroup.orientation = "row";
        radioGroup.alignChildren = ["left", "center"];

        var radios = [label];

        for (var i = 0; i < radioEntries.length; i++) {
            var radio = radioGroup.add("radiobutton", undefined, getLabel(radioEntries[i]));
            radio.value = (i === selectedIndex);
            radios.push(radio);
        }

        setTooltip(radios, tooltipEntry);

        /* 先頭のラベルは戻さない / The label is only there for the tooltip */
        radios.shift();

        return { row: row, radios: radios };
    }

    /**
     * 入力欄の数値を読む（読めないときや下限を下回るときは既定値に寄せる）
     * @param {object} editText 対象の edittext
     * @param {number} minimum 下限値
     * @param {number} fallback 読めないときの値
     * @param {boolean} asInteger 整数に丸めるか
     * @return {number} 読み取った値
     */
    function readNumberField(editText, minimum, fallback, asInteger) {
        var value = Number(editText.text);

        if (isNaN(value)) {
            return fallback;
        }

        if (asInteger) {
            value = Math.round(value);
        }

        return (value < minimum) ? fallback : value;
    }

    /**
     * プリセット一覧をドロップダウンへ流し込む
     * @param {object} presetDropdown 対象のドロップダウン
     * @return {void}
     */
    function fillPresetDropdown(presetDropdown) {
        presetDropdown.removeAll();

        for (var i = 0; i < PRESETS.length; i++) {
            presetDropdown.add("item", getLabel(PRESETS[i].label))._preset = PRESETS[i];
        }
    }

    /**
     * 値をソース表記にする（プリセットのコード出力用）
     * @param {number|boolean|string} value 設定値
     * @return {string} ソース表記
     */
    function formatCodeValue(value) {
        if (typeof value === "string") {
            return '"' + value + '"';
        }

        if (typeof value === "boolean") {
            return value ? "true" : "false";
        }

        return String(value);
    }

    /**
     * プリセットを PRESETS へ貼り付けられるコードに変換する
     * @param {string} presetName プリセット名
     * @param {object} values プリセットの設定値
     * @return {string} 組み込み用のコード
     */
    function presetToCode(presetName, values) {
        var valueLines = [];

        for (var key in values) {
            if (values.hasOwnProperty(key)) {
                valueLines.push("                " + key + ": " + formatCodeValue(values[key]));
            }
        }

        return ",\n        {\n" +
            '            label: { ja: "' + presetName + '", en: "' + presetName + '" },\n' +
            "            settings: {\n" + valueLines.join(",\n") + "\n            }\n        }";
    }

    /**
     * プリセットを選ぶ行（ドロップダウンと［保存］ボタン）を追加する
     * @param {object} dialog 対象のウィンドウ
     * @return {object} { presetDropdown, saveButton }
     */
    function addPresetRow(dialog) {
        var presetRow = addRow(dialog);
        var presetLabel = addRowLabel(presetRow, LABELS.label.preset);

        var presetDropdown = presetRow.add("dropdownlist");
        presetDropdown.alignment = ["fill", "center"];

        fillPresetDropdown(presetDropdown);
        presetDropdown.selection = 0;

        var saveButton = presetRow.add("button", undefined, getLabel(LABELS.button.add));
        saveButton.preferredSize.width = UI.saveButtonWidth;

        setTooltip([presetLabel, presetDropdown], LABELS.tooltip.preset);
        setTooltip([saveButton], LABELS.tooltip.savePreset);

        return { presetDropdown: presetDropdown, saveButton: saveButton };
    }

    /**
     * ［全体］パネルを追加する
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @return {object} 全体パネルのコントロール
     */
    function addOverallPanel(dialog, isVertical) {
        var overallPanel = addPanel(dialog, LABELS.panel.overall);

        var cellRectsCheckbox = addIndentedCheckbox(overallPanel, LABELS.checkbox.cellRects, DEFAULTS.cellRects, LABELS.tooltip.cellRects);

        var colorRow = addRow(overallPanel);
        var colorLabel = addRowLabel(colorRow, LABELS.label.lineColor);

        var colorSlider = colorRow.add("slider", undefined, DEFAULTS.strokeGray, 0, 100);
        colorSlider.alignment = ["fill", "center"];

        /* 幅は作成時の文字列で確保しておく / Reserve the width with the longest text */
        var colorValueLabel = colorRow.add("statictext", undefined, formatDensity(100));
        colorValueLabel.text = formatDensity(DEFAULTS.strokeGray);

        setTooltip([colorLabel, colorSlider], LABELS.tooltip.lineColor);

        var emphasisRatioRow = addNumberRow(overallPanel, LABELS.label.emphasisRatio, DEFAULTS.emphasisRatio, LABELS.unit.percent, LABELS.tooltip.emphasisRatio, { min: 1 });

        var emphasisTargetRow = addRow(overallPanel);
        var emphasisTargetLabel = addRowLabel(emphasisTargetRow, LABELS.label.emphasisTarget);

        var emphasisTargetGroup = emphasisTargetRow.add("group");
        emphasisTargetGroup.orientation = "row";
        emphasisTargetGroup.alignChildren = ["left", "center"];

        var emphasizeLineRulesCheckbox = emphasisTargetGroup.add("checkbox", undefined, getLabel(orientedEntry(LABELS.panel.lineRule, isVertical)));
        emphasizeLineRulesCheckbox.value = DEFAULTS.emphasizeLineRules;

        var cellRuleTargetLabel = getLabel(orientedEntry(LABELS.panel.cellRule, isVertical)) + getLabel(LABELS.label.emphasisTargetSuffix);
        var emphasizeCellRulesCheckbox = emphasisTargetGroup.add("checkbox", undefined, cellRuleTargetLabel);
        emphasizeCellRulesCheckbox.value = DEFAULTS.emphasizeCellRules;

        setTooltip([emphasisTargetLabel, emphasizeLineRulesCheckbox, emphasizeCellRulesCheckbox], LABELS.tooltip.emphasisTarget);
        var extraCellsRow = addNumberRow(overallPanel, LABELS.label.extraCells, DEFAULTS.extraCells, LABELS.unit.characters, LABELS.tooltip.extraCells, { integer: true, min: 0 });
        var extraLinesRow = addNumberRow(overallPanel, LABELS.label.extraLines, DEFAULTS.extraLines, LABELS.unit.lines, LABELS.tooltip.extraLines, { integer: true, min: 0 });

        return {
            cellRectsCheckbox: cellRectsCheckbox,
            colorSlider: colorSlider,
            colorValueLabel: colorValueLabel,
            emphasisRatioRow: emphasisRatioRow.row,
            emphasisRatioInput: emphasisRatioRow.input,
            emphasisTargetRow: emphasisTargetRow,
            emphasizeLineRulesCheckbox: emphasizeLineRulesCheckbox,
            emphasizeCellRulesCheckbox: emphasizeCellRulesCheckbox,
            extraCellsInput: extraCellsRow.input,
            extraLinesInput: extraLinesRow.input
        };
    }

    /**
     * ［文字］パネルを追加する
     * @param {object} dialog 対象のウィンドウ
     * @return {object} { adjustSizeCheckbox, scaleRow, scaleInput }
     */
    function addTextPanel(dialog) {
        var textPanel = addPanel(dialog, LABELS.panel.text);

        var adjustSizeCheckbox = addIndentedCheckbox(textPanel, LABELS.checkbox.adjustSize, DEFAULTS.adjustSize, LABELS.tooltip.adjustSize);
        var scaleRow = addNumberRow(textPanel, LABELS.label.scale, DEFAULTS.scale, LABELS.unit.percent, LABELS.tooltip.scale, { min: 1 });
        var autoLeadingRow = addNumberRow(textPanel, LABELS.label.autoLeading, DEFAULTS.autoLeading, LABELS.unit.percent, LABELS.tooltip.autoLeading, { min: 1 });

        return {
            adjustSizeCheckbox: adjustSizeCheckbox,
            scaleRow: scaleRow.row,
            scaleInput: scaleRow.input,
            autoLeadingInput: autoLeadingRow.input
        };
    }

    /**
     * 行の両側の罫線のパネルを追加する（縦組みでは［縦罫］、横組みでは［横罫］）
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @param {number} defaultExtensionMM 伸張の既定値（mm）
     * @return {object} { extensionInput, sideRuleCheckbox }
     */
    function addLineRulePanel(dialog, isVertical, defaultExtensionMM) {
        var linePanel = addPanel(dialog, orientedEntry(LABELS.panel.lineRule, isVertical));

        var extensionRow = addNumberRow(linePanel, LABELS.label.extension, defaultExtensionMM, LABELS.unit.mm, LABELS.tooltip.extension, { min: 0 });
        var sideRuleCheckbox = addIndentedCheckbox(linePanel, LABELS.checkbox.sideRule, DEFAULTS.showSideRules, LABELS.tooltip.sideRule);

        return {
            linePanel: linePanel,
            extensionInput: extensionRow.input,
            sideRuleCheckbox: sideRuleCheckbox
        };
    }

    /**
     * マスの区切りのパネルを追加する（縦組みでは［横罫］、横組みでは［縦罫］）
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @return {object} { cellSolidRadio, cellDashedRadio, emphasisCheckbox, emphasisInput }
     */
    function addCellRulePanel(dialog, isVertical) {
        var cellPanel = addPanel(dialog, orientedEntry(LABELS.panel.cellRule, isVertical));

        var styleRow = addRadioRow(cellPanel, LABELS.label.lineStyle,
            [LABELS.radio.solid, LABELS.radio.dashed], DEFAULTS.dashedCellRules ? 1 : 0, LABELS.tooltip.cellRuleStyle);

        var emphasisRow = addRow(cellPanel);
        var emphasisLabel = addRowLabel(emphasisRow, LABELS.label.emphasis);

        var emphasisCheckbox = emphasisRow.add("checkbox", undefined, "");
        emphasisCheckbox.value = DEFAULTS.emphasisEnabled;

        var emphasisPrefix = getLabel(LABELS.label.emphasisPrefix);

        if (emphasisPrefix.length > 0) {
            emphasisRow.add("statictext", undefined, emphasisPrefix);
        }

        var emphasisInput = addStepperField(emphasisRow, String(DEFAULTS.emphasisEvery), 3, { integer: true, min: 1 });

        emphasisRow.add("statictext", undefined, getLabel(LABELS.label.emphasisSuffix));
        setTooltip([emphasisLabel, emphasisCheckbox, emphasisInput], LABELS.tooltip.emphasis);

        return {
            cellSolidRadio: styleRow.radios[0],
            cellDashedRadio: styleRow.radios[1],
            emphasisRow: emphasisRow,
            emphasisCheckbox: emphasisCheckbox,
            emphasisInput: emphasisInput
        };
    }

    /**
     * ［十字線］パネルを追加する
     * @param {object} dialog 対象のウィンドウ
     * @return {object} { crossNoneRadio, crossSolidRadio, crossDashedRadio, crossSegmentsRow, crossSegmentsInput }
     */
    function addCrossPanel(dialog) {
        var crossPanel = addPanel(dialog, LABELS.panel.cross);

        var crossStyleRow = addRadioRow(crossPanel, LABELS.label.lineStyle,
            [LABELS.radio.none, LABELS.radio.solid, LABELS.radio.dashed], crossStyleIndex(DEFAULTS.crossStyle), LABELS.tooltip.crossStyle);

        var crossSegmentsRow = addNumberRow(crossPanel, LABELS.label.segments, DEFAULTS.crossSegments, null, LABELS.tooltip.crossSegments, { integer: true, min: 0 });

        return {
            crossNoneRadio: crossStyleRow.radios[0],
            crossSolidRadio: crossStyleRow.radios[1],
            crossDashedRadio: crossStyleRow.radios[2],
            crossSegmentsRow: crossSegmentsRow.row,
            crossSegmentsInput: crossSegmentsRow.input
        };
    }

    /**
     * ダイアログのコントロールを組み立てる
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @param {number} defaultExtensionMM 伸張の既定値（mm）
     * @return {object} コントロール一式
     */
    function buildDialogControls(dialog, isVertical, defaultExtensionMM) {
        var preset = addPresetRow(dialog);
        var overall = addOverallPanel(dialog, isVertical);
        var text = addTextPanel(dialog);
        var lineRule = addLineRulePanel(dialog, isVertical, defaultExtensionMM);
        var cellRule = addCellRulePanel(dialog, isVertical);
        var cross = addCrossPanel(dialog);

        /* ボタンエリア（左にプレビュー、右にキャンセル・OK） / Button row: preview on the left, Cancel and OK on the right */
        var buttonRow = addButtonRow(dialog);
        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.value = DEFAULTS.preview;
        setTooltip([previewCheckbox], LABELS.tooltip.preview);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            presetDropdown: preset.presetDropdown,
            saveButton: preset.saveButton,
            cellRectsCheckbox: overall.cellRectsCheckbox,
            colorSlider: overall.colorSlider,
            colorValueLabel: overall.colorValueLabel,
            emphasisRatioRow: overall.emphasisRatioRow,
            emphasisRatioInput: overall.emphasisRatioInput,
            emphasisTargetRow: overall.emphasisTargetRow,
            emphasizeLineRulesCheckbox: overall.emphasizeLineRulesCheckbox,
            emphasizeCellRulesCheckbox: overall.emphasizeCellRulesCheckbox,
            extraCellsInput: overall.extraCellsInput,
            extraLinesInput: overall.extraLinesInput,
            adjustSizeCheckbox: text.adjustSizeCheckbox,
            scaleRow: text.scaleRow,
            scaleInput: text.scaleInput,
            autoLeadingInput: text.autoLeadingInput,
            linePanel: lineRule.linePanel,
            extensionInput: lineRule.extensionInput,
            sideRuleCheckbox: lineRule.sideRuleCheckbox,
            cellSolidRadio: cellRule.cellSolidRadio,
            cellDashedRadio: cellRule.cellDashedRadio,
            emphasisRow: cellRule.emphasisRow,
            emphasisCheckbox: cellRule.emphasisCheckbox,
            emphasisInput: cellRule.emphasisInput,
            crossNoneRadio: cross.crossNoneRadio,
            crossSolidRadio: cross.crossSolidRadio,
            crossDashedRadio: cross.crossDashedRadio,
            crossSegmentsRow: cross.crossSegmentsRow,
            crossSegmentsInput: cross.crossSegmentsInput,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * 十字線の線種に対応するラジオの位置を返す
     * @param {string} crossStyle 線種（none / solid / dashed）
     * @return {number} ラジオの位置
     */
    function crossStyleIndex(crossStyle) {
        if (crossStyle === "none") {
            return 0;
        }

        return (crossStyle === "solid") ? 1 : 2;
    }

    /**
     * チェックやラジオの状態に合わせて、入力欄の有効・無効を切り替える
     * @param {object} controls コントロール一式
     * @return {void}
     */
    function updateDialogState(controls) {
        /* 長方形にすると太罫も行の両側の罫線も出番がなくなる
           / The rectangles leave no room for the emphasis or the rules along each line */
        var usesRules = !controls.cellRectsCheckbox.value;

        controls.emphasisRatioRow.enabled = usesRules;
        controls.emphasisTargetRow.enabled = usesRules;
        controls.linePanel.enabled = usesRules;
        controls.emphasisRow.enabled = usesRules;

        controls.scaleRow.enabled = controls.adjustSizeCheckbox.value;
        controls.emphasisInput.enabled = controls.emphasisCheckbox.value;
        controls.emphasisInput.stepperGroup.enabled = controls.emphasisInput.enabled;
        controls.crossSegmentsRow.enabled = controls.crossDashedRadio.value;

        /* ∧∨は自作描画なので、行やパネルの有効／無効を変えたら描き直す / redraw the custom-drawn steppers */
        var stepperContainers = [controls.emphasisRatioRow, controls.linePanel, controls.emphasisRow, controls.scaleRow, controls.crossSegmentsRow];
        for (var i = 0; i < stepperContainers.length; i++) {
            redrawSteppersIn(stepperContainers[i]);
        }
    }

    /**
     * ダイアログの入力内容を、プリセットと同じ形で読む
     * @param {object} controls コントロール一式
     * @return {object} プリセットの設定値
     */
    function readPresetValues(controls) {
        return {
            cellRects: controls.cellRectsCheckbox.value,
            strokeGray: Math.round(controls.colorSlider.value),
            emphasisRatio: readNumberField(controls.emphasisRatioInput, 1, DEFAULTS.emphasisRatio, false),
            emphasizeLineRules: controls.emphasizeLineRulesCheckbox.value,
            emphasizeCellRules: controls.emphasizeCellRulesCheckbox.value,
            extraCells: readNumberField(controls.extraCellsInput, 0, 0, true),
            extraLines: readNumberField(controls.extraLinesInput, 0, 0, true),
            adjustSize: controls.adjustSizeCheckbox.value,
            scale: readNumberField(controls.scaleInput, 1, DEFAULTS.scale, false),
            autoLeading: readNumberField(controls.autoLeadingInput, 1, DEFAULTS.autoLeading, false),
            extensionMM: readNumberField(controls.extensionInput, 0, 0, false),
            dashedCellRules: controls.cellDashedRadio.value,
            emphasisEnabled: controls.emphasisCheckbox.value,
            emphasisEvery: readNumberField(controls.emphasisInput, 1, 0, true),
            showSideRules: controls.sideRuleCheckbox.value,
            crossStyle: controls.crossNoneRadio.value ? "none" : (controls.crossSolidRadio.value ? "solid" : "dashed"),
            crossSegments: normalizeCrossSegments(Number(controls.crossSegmentsInput.text), true)
        };
    }

    /**
     * ダイアログの入力内容を、罫線を引くための設定にまとめる
     * @param {object} controls コントロール一式
     * @return {object} 罫線を引くための設定
     */
    function readDialogSettings(controls) {
        var values = readPresetValues(controls);

        return {
            cellRects: values.cellRects,
            strokeGray: values.strokeGray,
            emphasisRatio: values.emphasisRatio,
            emphasizeLineRules: values.emphasizeLineRules,
            emphasizeCellRules: values.emphasizeCellRules,
            extraCells: values.extraCells,
            extraLines: values.extraLines,
            adjustSize: values.adjustSize,
            scale: values.scale,
            autoLeading: values.autoLeading,
            extension: mmToPt(values.extensionMM),
            dashedCellRules: values.dashedCellRules,
            emphasisEvery: values.emphasisEnabled ? values.emphasisEvery : 0,
            showSideRules: values.showSideRules,
            crossStyle: values.crossStyle,
            crossSegments: values.crossSegments
        };
    }

    /**
     * プリセットの値をダイアログへ流し込む
     * @param {object} controls コントロール一式
     * @param {object} values プリセットの設定値
     * @param {number} defaultExtensionMM 伸張の既定値（mm）
     * @return {void}
     */
    function applyPresetValues(controls, values, defaultExtensionMM) {
        controls.cellRectsCheckbox.value = values.cellRects;
        controls.colorSlider.value = values.strokeGray;
        controls.colorValueLabel.text = formatDensity(values.strokeGray);
        controls.emphasisRatioInput.text = String(values.emphasisRatio);
        controls.emphasizeLineRulesCheckbox.value = values.emphasizeLineRules;
        controls.emphasizeCellRulesCheckbox.value = values.emphasizeCellRules;
        controls.extraCellsInput.text = String(values.extraCells);
        controls.extraLinesInput.text = String(values.extraLines);

        controls.adjustSizeCheckbox.value = values.adjustSize;
        controls.scaleInput.text = String(values.scale);
        controls.autoLeadingInput.text = String(values.autoLeading);

        /* null は「文字サイズの1/4」/ null means a quarter of the font size */
        controls.extensionInput.text = String((values.extensionMM === null) ? defaultExtensionMM : values.extensionMM);
        controls.sideRuleCheckbox.value = values.showSideRules;

        controls.cellSolidRadio.value = !values.dashedCellRules;
        controls.cellDashedRadio.value = values.dashedCellRules;

        controls.emphasisCheckbox.value = values.emphasisEnabled;
        controls.emphasisInput.text = String(values.emphasisEvery);

        controls.crossNoneRadio.value = (values.crossStyle === "none");
        controls.crossSolidRadio.value = (values.crossStyle === "solid");
        controls.crossDashedRadio.value = (values.crossStyle === "dashed");
        controls.crossSegmentsInput.text = String(values.crossSegments);

        updateDialogState(controls);
    }

    /**
     * ダイアログのイベントをつなぐ
     * @param {object} controls コントロール一式
     * @param {object} callbacks { onPresetSelected, onPresetSaved, onSettingChanged, onPreviewToggled }
     * @return {void}
     */
    function bindDialogHandlers(controls, callbacks) {
        /* チェックやラジオは、ディム状態を直してからプレビューを更新する */
        function onToggle() {
            updateDialogState(controls);
            callbacks.onSettingChanged();
        }

        controls.presetDropdown.onChange = callbacks.onPresetSelected;
        controls.saveButton.onClick = callbacks.onPresetSaved;

        controls.colorSlider.onChanging = function () {
            var grayPercent = Math.round(controls.colorSlider.value);

            /* shiftを押している間は10%刻みにする / Shift snaps the slider to steps of 10% */
            if (ScriptUI.environment.keyboardState.shiftKey) {
                grayPercent = Math.round(grayPercent / 10) * 10;
                controls.colorSlider.value = grayPercent;
            }

            controls.colorValueLabel.text = formatDensity(grayPercent);
        };

        /* ドラッグ中に引き直すと重いので、離したタイミングで更新する / Redraw once the drag ends */
        controls.colorSlider.onChange = callbacks.onSettingChanged;

        /* 下限・整数かどうかは欄を作るときの stepOptions に入れてある / min and integer live in each field's stepOptions */
        var numberFields = [
            controls.emphasisRatioInput,
            controls.extraCellsInput,
            controls.extraLinesInput,
            controls.scaleInput,
            controls.autoLeadingInput,
            controls.extensionInput,
            controls.emphasisInput
        ];

        for (var i = 0; i < numberFields.length; i++) {
            numberFields[i].onChanging = callbacks.onSettingChanged;
            /* 値の代入では onChanging が呼ばれないため、∧∨・↑↓キーで増減したらプレビューを更新する */
            numberFields[i].stepOptions.onStep = callbacks.onSettingChanged;
            bindSteppedArrowKeys(numberFields[i], numberFields[i].stepperGroup);
        }

        var toggles = [
            controls.cellRectsCheckbox,
            controls.emphasizeLineRulesCheckbox,
            controls.emphasizeCellRulesCheckbox,
            controls.adjustSizeCheckbox,
            controls.sideRuleCheckbox,
            controls.emphasisCheckbox,
            controls.cellSolidRadio,
            controls.cellDashedRadio,
            controls.crossNoneRadio,
            controls.crossSolidRadio,
            controls.crossDashedRadio
        ];

        for (var j = 0; j < toggles.length; j++) {
            toggles[j].onClick = onToggle;
        }

        controls.crossSegmentsInput.onChanging = callbacks.onSettingChanged;
        /* 分割数は 0 → 3 → 5 → 7 と動かす。増減の向きは増減前の値と比べて決める / step 0, 3, 5, 7 in the direction moved */
        controls.crossSegmentsInput.stepOptions.onStep = function (input) {
            var steppedValue = Number(input.text);
            var valueBefore = Number(input.textBeforeStep);
            var isUp = isNaN(valueBefore) || steppedValue > valueBefore;
            input.text = String(normalizeCrossSegments(steppedValue, isUp));
            callbacks.onSettingChanged();
        };
        bindSteppedArrowKeys(controls.crossSegmentsInput, controls.crossSegmentsInput.stepperGroup);

        /* 入力欄を離れたら、実際に使う値（0か奇数）に揃える / Snap the field to the value actually used */
        controls.crossSegmentsInput.onChange = function () {
            controls.crossSegmentsInput.text = String(normalizeCrossSegments(Number(controls.crossSegmentsInput.text), true));
            callbacks.onSettingChanged();
        };

        controls.previewCheckbox.onClick = callbacks.onPreviewToggled;
    }

    /**
     * 設定ダイアログを表示する
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @return {object|null} 設定値。キャンセルしたときは null
     */
    function showSettingsDialog(previewContext) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var defaultExtensionMM = getDefaultExtensionMM(previewContext.fontSize);
        var controls = buildDialogControls(dialog, previewContext.previewState.isVertical, defaultExtensionMM);

        /* 前回の設定があれば、ハンドラをつなぐ前に流し込む（onChangeを呼ばせない）
           / Restore the last settings before the handlers are wired, so no events fire */
        var session = loadSession();

        if (session) {
            applyPresetValues(controls, session.values, defaultExtensionMM);

            if (session.presetIndex < controls.presetDropdown.items.length) {
                controls.presetDropdown.selection = session.presetIndex;
            }
        }

        /**
         * プレビューの表示・非表示を現在の設定に合わせる
         * @return {void}
         */
        function refreshPreview() {
            if (controls.previewCheckbox.value) {
                drawPreviewGrid(previewContext, readDialogSettings(controls));
            } else {
                clearPreviewGrid(previewContext);
            }
        }

        /**
         * 手で値を変えたら、プリセットの選択を「カスタム」へ戻してプレビューを更新する
         * @return {void}
         */
        function onSettingChanged() {
            /* カスタムは末尾に置いてある / Custom is the last item */
            controls.presetDropdown.selection = controls.presetDropdown.items.length - 1;
            refreshPreview();
        }

        /**
         * 選んだプリセットをダイアログへ流し込む
         * @return {void}
         */
        function onPresetSelected() {
            var preset = controls.presetDropdown.selection ? controls.presetDropdown.selection._preset : null;

            if (!preset || !preset.settings) {
                return;
            }

            applyPresetValues(controls, preset.settings, defaultExtensionMM);
            refreshPreview();
        }

        /**
         * いまの設定に名前を付けてプリセットへ加える（残すためのコードは通知で出す）
         * @return {void}
         */
        function onPresetSaved() {
            var currentName = controls.presetDropdown.selection ? controls.presetDropdown.selection.text : "";
            var presetName = prompt(getLabel(LABELS.alert.presetName), currentName);

            if (presetName === null) {
                return;
            }

            presetName = String(presetName).replace(/^\s+|\s+$/g, "");

            if (!presetName) {
                return;
            }

            var values = readPresetValues(controls);

            /* 「カスタム」は末尾に置いたままにする / Keep Custom at the end of the list */
            PRESETS.splice(PRESETS.length - 1, 0, { label: { ja: presetName, en: presetName }, settings: values });

            fillPresetDropdown(controls.presetDropdown);
            controls.presetDropdown.selection = PRESETS.length - 2;

            alert(getLabel(LABELS.alert.presetSaved) + presetToCode(presetName, values));
        }

        bindDialogHandlers(controls, {
            onPresetSelected: onPresetSelected,
            onPresetSaved: onPresetSaved,
            onSettingChanged: onSettingChanged,
            onPreviewToggled: refreshPreview
        });

        /* プレビューの初期状態をダイアログを開く前に反映しておく */
        updateDialogState(controls);
        refreshPreview();

        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();

        if (dialogResult !== 1) {
            return null;
        }

        var selectedPreset = controls.presetDropdown.selection;
        saveSession(readPresetValues(controls), selectedPreset ? selectedPreset.index : 0);

        return readDialogSettings(controls);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 決まった設定でテキストを整え、罫線を引く
     * @param {Document} doc 対象のドキュメント
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @return {void}
     */
    function applyGridToText(doc, textFrame, settings) {
        var fitResult = fitTextToGrid(textFrame, settings, true);

        if (!fitResult.metrics) {
            alert(getLabel(LABELS.alert.noFontSize));
            return;
        }

        var layerName = getLabel(LABELS.itemName.layerName);
        var groupName = getLabel(LABELS.itemName.groupName);

        var ruleLayer = prepareRuleLayer(doc, layerName);
        removeExistingGroup(ruleLayer, groupName, textFrame.geometricBounds);

        var ruleGroup = ruleLayer.groupItems.add();
        ruleGroup.name = groupName;

        drawGrid(ruleGroup, fitResult.metrics, settings, createRuleColors(doc, settings));

        /* グループの zOrder だけではテキストの下に回らないため、レイヤーごと最背面へ送る */
        sendLayerToBack(ruleLayer);

        /* 成功時は通知しない。設定できなかった属性があるときだけ知らせる */
        if (fitResult.failed.length > 0) {
            alert(getLabel(LABELS.alert.skipped) + " " + fitResult.failed.join(", "));
        }
    }

    /**
     * 選択のエッジ表示を切り替える（プレビューを見やすくする）
     * @return {void}
     */
    function toggleSelectionEdges() {
        if (app.documents.length === 0) {
            return;
        }

        app.executeMenuCommand("edge");
    }

    /**
     * 選択したテキストに罫線を引く
     * @return {void}
     */
    function runGridMaker() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (!selectedItems || selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") {
            alert(getLabel(LABELS.alert.noTextFrame));
            return;
        }

        var textFrame = selectedItems[0];

        if (textFrame.textRange.characters.length === 0) {
            alert(getLabel(LABELS.alert.noCharacter));
            return;
        }

        /* 先頭文字の文字サイズを基準にするので、先に読めることを確かめる */
        var fontSize = textFrame.textRange.characters[0].characterAttributes.size;

        if (!fontSize || fontSize <= 0) {
            alert(getLabel(LABELS.alert.noFontSize));
            return;
        }

        /* エリア内文字は、プレビューを作る前に枠を文字へ合わせておく
           / Fit the area text frame before the preview copy is made */
        if (isAreaText(textFrame)) {
            /* セットは main() の最後に解除する。読み込めなければ枠のままにする
               / The set is unloaded at the end of main(); if it cannot be loaded, the frame stays as it is */
            loadTemporaryActionSet(buildAutoSizeAia(), AUTO_SIZE_ACTION_SET);
            fitAreaTextFrame(doc, textFrame);
        }

        /* 原本は触らず、確定後と同じ状態にした複製でプレビューする / Preview on a copy, leaving the original untouched */
        var previewState = beginPreview(doc, textFrame);

        var previewContext = {
            previewState: previewState,
            fontSize: fontSize,
            appliedShape: { adjustSize: null, scale: 0 }, /* 最初の描画で必ず作り直す / forces the first reshape */
            metrics: null
        };

        var settings = null;

        /* 途中で失敗しても原本を隠したままにしない / Never leave the original hidden if something fails */
        try {
            settings = showSettingsDialog(previewContext);
        } finally {
            endPreview(previewState);
        }

        if (!settings) {
            return;
        }

        applyGridToText(doc, textFrame, settings);
    }

    /**
     * エッジ表示を戻せるよう、本体を挟んで実行する
     * @return {void}
     */
    function main() {
        toggleSelectionEdges();

        try {
            runGridMaker();
        } finally {
            unloadTemporaryActionSet(AUTO_SIZE_ACTION_SET);
            toggleSelectionEdges();
        }
    }

    main();
})();
