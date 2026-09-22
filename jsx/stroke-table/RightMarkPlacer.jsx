#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した複数オブジェクトを左から順に見て、隣り合うオブジェクト同士のアキの中央に記号を配置します。
記号は9種類から選べ、高さ・幅・線幅・位置をダイアログで調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RightMarkPlacer.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nebac730ec187

### Overview

Scans the selected objects from left to right and places a mark in the middle of the gap between each adjacent pair.
Nine marks are available, with height, width, stroke weight and position set from the dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RightMarkPlacer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RightMarkPlacer";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RightMarkPlacer.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RightMarkPlacer.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nebac730ec187"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 記号の色（CMYK）/ Mark color (CMYK) */
    var MARK_COLOR_CMYK = [0, 0, 0, 100];

    /* 高さ（％）の上限 / Maximum height percentage */
    var MAX_HEIGHT_PERCENT = 200;

    /* 線幅の下限（pt）/ Minimum stroke width in points */
    var MIN_STROKE_WIDTH_PT = 0.25;

    /* ▶ の凹みの上限（幅に対する割合）/ Maximum inset as a ratio of the width */
    var MAX_INSET_RATIO = 0.8;

    /* ▶ の角丸半径（幅と高さの小さい方に対する割合）/ Rounded-corner radius as a ratio of the smaller side */
    var TRIANGLE_CORNER_RADIUS_RATIO = 0.12;

    /* ＿\ の斜線の角度（度）/ Slash angle in degrees */
    var SLASH_ANGLE_DEFAULT = 35;
    var SLASH_ANGLE_MAX = 89;

    /* ➡ の矢じりの天地を求める、線幅に対する倍率 / Height of the solid arrowhead as a multiple of the stroke width */
    var SOLID_ARROW_HEIGHT_TO_STROKE_RATIO = 3;

    /* → ➡ の矢じりの奥行きが、幅に占める割合の上限 / Maximum arrowhead depth as a ratio of the width */
    var MAX_ARROW_HEAD_RATIO = 0.9;

    /* 幅を自動計算するときの、アキに対する割合 / Ratio of the gap used when the width is calculated automatically */
    var AUTO_WIDTH_GAP_RATIO = 0.7;
    var AUTO_WIDTH_GAP_RATIO_SMALL = 0.35;

    /* 山形の幅を自動計算するときの、高さに対する割合 / Ratio of the height used when a chevron width is calculated automatically */
    var AUTO_WIDTH_CHEVRON_HEIGHT_RATIO = 0.5;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var RADIO_SPACING  = 6;                  /* ラジオボタンの間隔 / gap between radio buttons */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var COLUMN_PANEL_SPACING = 10;           /* カラム内のパネルの間隔 / gap between panels in a column */
    var FIELD_ROW_SPACING = 8;               /* ラベルと入力欄の間隔 / gap between a label and its field */
    var LABEL_COLUMN_WIDTH = 60;             /* ラベル列の幅 / width of the label column */
    var FIELD_CHARACTERS = 4;                /* 入力欄の幅（文字数）/ field width in characters */
    var MIRROR_ROW_TOP_MARGIN = 6;           /* ［左右逆］の上余白 / top margin above the mirror checkbox */
    var OPTION_CHECKBOX_INDENT = 40;         /* ［オプション］のチェックボックスの字下げ / indent of the option checkboxes */
    var OPTION_CHECKBOX_TOP_MARGIN = 10;     /* ［オプション］のチェックボックスの上余白 / top margin above the option checkboxes */
    var BUTTON_ROW_TOP_MARGIN = 10;          /* ボタン行の上余白 / top margin above the button row */

    /**
     * ウィンドウに共通のレイアウトを適用します。
     *
     * @param {Window} targetWindow - 対象のウィンドウ。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用します。
     *
     * @param {Panel} targetPanel - 対象のパネル。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
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
     * 縦並びのグループ（カラム）に共通のレイアウトを適用します。
     *
     * @param {Group} columnGroup - 対象のグループ。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupColumn(columnGroup, spacing) {
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びのグループ（入力行など）に共通のレイアウトを適用します。
     *
     * @param {Group} rowGroup - 対象のグループ。
     * @param {string} [alignment] - グループ自体の配置。省略時は "left"。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    var rulerUnitInfo = getUnitInfo("rulerType");
    var strokeUnitInfo = getUnitInfo("strokeUnits");

    /**
     * 単位の数値を pt に換算します。
     *
     * @param {number} value - 単位付きの数値。
     * @param {object} unitInfo - 単位情報。
     * @returns {number} pt 値。
     */
    function convertValueToPt(value, unitInfo) {
        return value * unitInfo.pointsPerUnit;
    }

    /**
     * pt を単位の数値に換算します。
     *
     * @param {number} valuePt - pt 値。
     * @param {object} unitInfo - 単位情報。
     * @returns {number} 単位付きの数値。
     */
    function convertPtToUnitValue(valuePt, unitInfo) {
        return valuePt / unitInfo.pointsPerUnit;
    }

    /* 表示桁数 / Decimal places used for display */
    var DISPLAY_DECIMALS = 2;          /* 既定の桁数 / default decimal places */
    var DISPLAY_DECIMALS_MAX = 5;      /* 桁数の上限 / maximum decimal places */
    var DISPLAY_DECIMALS_PT_FACTOR = 3; /* この換算係数までは既定の桁数 / units up to this factor keep the default */

    /**
     * pt 値を小数点以下2桁へ丸めます。
     *
     * @param {number} value - 丸める値。
     * @returns {number} 丸めた値。
     */
    function roundDisplayValue(value) {
        return Math.round(value * 100) / 100;
    }

    /**
     * 単位に応じた表示桁数を求めます。
     * inch のように 1単位が大きい単位では、2桁だと pt 換算で精度が足りないため桁数を増やします。
     *
     * @param {object} [unitInfo] - 単位情報。
     * @returns {number} 小数点以下の桁数。
     */
    function getDisplayDecimals(unitInfo) {
        if (!unitInfo || !(unitInfo.pointsPerUnit > DISPLAY_DECIMALS_PT_FACTOR)) return DISPLAY_DECIMALS;
        var extraDigits = Math.ceil(Math.log(unitInfo.pointsPerUnit) / Math.LN10);
        return Math.min(DISPLAY_DECIMALS + extraDigits, DISPLAY_DECIMALS_MAX);
    }

    /**
     * 単位に合わせた桁数で丸めます。
     *
     * @param {number} value - 単位付きの数値。
     * @param {object} [unitInfo] - 単位情報。
     * @returns {number} 丸めた値。
     */
    function roundDisplayValueForUnit(value, unitInfo) {
        var scale = Math.pow(10, getDisplayDecimals(unitInfo));
        return Math.round(value * scale) / scale;
    }

    /**
     * 矢印キーで増減する量を、単位に合わせて求めます。
     * pt・px・mm・Q・H は 1単位、inch のように 1単位が大きい単位はより細かい刻みにします。
     *
     * @param {object} [unitInfo] - 単位情報。％や度など単位のない入力欄では省略します。
     * @returns {number} 増減量。
     */
    function getArrowKeyStep(unitInfo) {
        if (!unitInfo || !(unitInfo.pointsPerUnit > 0)) return 1;
        var exponent = Math.max(0, Math.round(Math.log(unitInfo.pointsPerUnit) / Math.LN10));
        return Math.pow(10, -exponent);
    }

    /**
     * pt 値を単位に換算して入力欄へ表示します。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {number} valuePt - pt 値。
     * @param {object} unitInfo - 単位情報。
     * @returns {void}
     */
    function setFieldFromPt(editText, valuePt, unitInfo) {
        editText.text = String(roundDisplayValueForUnit(convertPtToUnitValue(valuePt, unitInfo), unitInfo));
    }

    /**
     * 入力欄の値を pt として読み取ります。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {object} unitInfo - 単位情報。
     * @param {boolean} allowNegative - 負の値を許可するなら true。
     * @returns {number} pt 値。数値でない場合と、許可していない負の値の場合は NaN。
     */
    function parseFieldToPt(editText, unitInfo, allowNegative) {
        var value = parseFloat(editText.text);
        if (isNaN(value)) return NaN;
        /* 負の値を黙って 0 にせず、呼び出し元でエラーとして扱えるようにする / Report it instead of silently clamping to zero */
        if (!allowNegative && value < 0) return NaN;
        return convertValueToPt(value, unitInfo);
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定します。
     *
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            /* ＋ や × のような無方向の形状と［左右逆］があるため、向きを含めない名前にしています / Kept direction-neutral: some shapes have no direction and the mark can be mirrored */
            title: { ja: "オブジェクト間に記号を配置", en: "Place Marks Between Objects" }
        },
        panel: {
            shape: { ja: "形状", en: "Shape" },
            capStyle: { ja: "先端", en: "End Style" },
            options: { ja: "オプション", en: "Options" },
            position: { ja: "位置調整", en: "Position" }
        },
        radio: {
            arrowSlash: { ja: "＿\\", en: "─\\" },
            capNone: { ja: "なし", en: "None" },
            capRound: { ja: "丸型", en: "Round" }
        },
        checkbox: {
            mirrorHorizontal: { ja: "左右逆", en: "Mirror horizontally" },
            flatChevron: { ja: "天地を水平に", en: "Keep top and bottom edges horizontal" },
            roundCorners: { ja: "角丸", en: "Rounded corners" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        /* 入力欄の前に置く項目名。区切りのコロンは labelText() で付ける / Field captions; labelText() adds the colon */
        fieldLabel: {
            height: { ja: "高さ", en: "Height" },
            width: { ja: "幅", en: "Width" },
            gap: { ja: "間隔", en: "Gap" },
            inset: { ja: "凹み", en: "Inset" },
            strokeWidth: { ja: "線幅", en: "Stroke" },
            angle: { ja: "角度", en: "Angle" },
            offsetX: { ja: "左右", en: "Horizontal" },
            offsetY: { ja: "上下", en: "Vertical" }
        },
        tooltip: {
            solidArrow: { ja: "矢じりの天地は線幅の{0}倍になります。", en: "The arrowhead is {0}× the stroke width tall." },
            capNone: { ja: "ショートカット：F", en: "Shortcut: F" },
            capRound: { ja: "ショートカット：R", en: "Shortcut: R" },
            mirrorHorizontal: { ja: "ショートカット：V", en: "Shortcut: V" },
            height: {
                ja: "隣り合う2つのオブジェクト全体の高さに対する割合（{0}%まで）",
                en: "Percentage of the overall height of the two adjacent objects (up to {0}%)"
            },
            width: { ja: "空欄にすると自動計算に戻ります", en: "Clear to calculate automatically again" },
            inset: { ja: "幅の{0}%が上限です", en: "Max {0}% of width" },
            gap: {
                ja: ">> の2つの山形の間隔。負の値で重なります",
                en: "Space between the two chevrons of >>. Negative values overlap them"
            },
            angle: { ja: "＿\\ の斜線の角度（{0}°まで）", en: "Slash angle of ─\\ (up to {0}°)" },
            flatChevron: {
                ja: "> と >> を、天地が水平な塗りの形状で作成します",
                en: "Draws > and >> as filled shapes with horizontal top and bottom edges"
            },
            roundCorners: {
                ja: "▶ の角を丸くします。半径は記号の大きさから自動で決まります",
                en: "Rounds the corners of ▶. The radius is set automatically from the mark size"
            },
            offsetX: { ja: "記号を左右にずらします。正の値で右へ移動します", en: "Shifts the marks horizontally. Positive values move them right" },
            offsetY: { ja: "記号を上下にずらします。正の値で上へ移動します", en: "Shifts the marks vertically. Positive values move them up" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            openDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectTwoObjects: { ja: "オブジェクトを2つ以上選択してください。", en: "Please select two or more objects." },
            lockedLayer: {
                ja: "作業レイヤーがロックまたは非表示です。解除してから実行してください。",
                en: "The active layer is locked or hidden. Please unlock and show it, then run again."
            },
            positiveNumber: { ja: "正の数値を入力してください。", en: "Please enter a positive value." },
            maxHeight: { ja: "{0}% 以下の値を入力してください。", en: "Please enter a value of {0}% or less." },
            noGap: {
                ja: "隣り合うオブジェクト間に作成できるアキがありません。",
                en: "There is no usable gap between adjacent objects."
            },
            strokePositive: {
                ja: "線幅は {0} {1} 以上の値を入力してください。",
                en: "Please enter a stroke width of {0} {1} or greater."
            },
            widthPositive: { ja: "幅は 0 以上の値を入力してください。", en: "Please enter a width value of 0 or greater." },
            invalidValue: { ja: "入力値を確認してください。", en: "Please check the input values." }
        },
        log: {
            measureTextBounds: { ja: "テキストの計測用アウトライン化", en: "Outline text for measurement" },
            removeMeasurementCopy: { ja: "計測用複製の削除", en: "Remove measurement copy" },
            applyRoundCorners: { ja: "角丸効果の適用", en: "Apply rounded-corners effect" },
            restoreSelection: { ja: "選択状態の復元", en: "Restore selection" },
            removePreviewItem: { ja: "プレビューの削除", en: "Remove preview item" },
            mergeSolidArrow: { ja: "➡ の合成", en: "Merge the solid arrow" }
        }
    };

    /**
     * ドットパスで指定したラベルを、現在の言語で取得します。
     *
     * @param {string} labelPath - ラベルのドットパス（例 "panel.shape"）。
     * @returns {string} 現在の言語のラベル。見つからない場合は英語、それもなければ labelPath。
     */
    function getLabel(labelPath) {
        var pathParts = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathParts.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[pathParts[i]];
        }
        if (labelNode && labelNode[uiLang]) return labelNode[uiLang];
        if (labelNode && labelNode.en) return labelNode.en;
        return labelPath;
    }

    /**
     * 項目名に、言語に合わせたコロンを付けて返します（日本語は全角、英語は半角）。
     *
     * @param {string} labelPath - ラベルのドットパス。
     * @returns {string} コロン付きの項目名。
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * ラベル中の {0} {1} … を、与えた値で置き換えます。
     *
     * @param {string} labelPath - ラベルのドットパス。
     * @param {Array} [replacements] - 差し込む値の配列。
     * @returns {string} 置き換え後の文言。
     */
    function formatLabel(labelPath, replacements) {
        var formattedText = getLabel(labelPath);
        if (!replacements) return formattedText;
        for (var i = 0; i < replacements.length; i++) {
            formattedText = formattedText.replace("{" + i + "}", String(replacements[i]));
        }
        return formattedText;
    }

    // =========================================
    // 汎用ユーティリティ / Generic utilities
    // =========================================

    /**
     * コンテキスト付きでエラーを $.writeln に出力します。
     *
     * @param {string} logContext - どの処理で起きたかを示す文言。
     * @param {object} errorObject - エラーオブジェクトまたはメッセージ。
     * @returns {void}
     */
    function logScriptError(logContext, errorObject) {
        $.writeln("[" + SCRIPT_NAME + " " + SCRIPT_VERSION + "] " + logContext + ": " + errorObject);
    }

    /**
     * 処理を実行し、失敗した場合はログを出して続行します。
     *
     * @param {Function} operation - 実行する処理。
     * @param {string} logContext - 失敗時にログへ出す文言。
     * @returns {boolean} 成功したら true、失敗したら false。
     */
    function runSafely(operation, logContext) {
        try {
            operation();
            return true;
        } catch (e) {
            logScriptError(logContext, e);
            return false;
        }
    }

    /**
     * ExtendScript のコレクションを通常の配列へコピーします。
     *
     * @param {object} collection - selection などのコレクション。
     * @returns {Array} コピーした配列。
     */
    function collectionToArray(collection) {
        var copiedItems = [];
        if (!collection) return copiedItems;
        for (var i = 0; i < collection.length; i++) {
            copiedItems.push(collection[i]);
        }
        return copiedItems;
    }

    /**
     * アイテムを削除します。削除済みのアイテムを渡しても止まらないよう、失敗はログだけにします。
     *
     * @param {PageItem} pageItem - 対象のアイテム。null なら何もしません。
     * @param {string} logContext - 失敗時にログへ出す文言。
     * @returns {void}
     */
    function removeItemSafely(pageItem, logContext) {
        if (!pageItem) return;
        runSafely(function () {
            pageItem.remove();
        }, logContext);
    }

    /**
     * 選択状態をまとめて復元します。
     * メニューコマンドで置き換わったアイテムが混ざっても止まらないよう、1つずつ選択し直します。
     *
     * @param {PageItem[]} pageItems - 選択し直すアイテムの配列。
     * @returns {void}
     */
    function restoreSelection(pageItems) {
        app.activeDocument.selection = null;
        for (var i = 0; i < pageItems.length; i++) {
            (function (pageItem) {
                runSafely(function () {
                    pageItem.selected = true;
                }, getLabel("log.restoreSelection"));
            })(pageItems[i]);
        }
    }

    // =========================================
    // 選択オブジェクトの計測 / Measuring the selection
    // =========================================

    /*
     * 同じアイテムを何度も計測しないためのキャッシュ。
     * モーダルダイアログの表示中は選択オブジェクトが変化しないため、閉じるまで保持します。
     * テキストの計測はアウトライン化を伴うので、プレビューの更新ごとに測り直すと重くなります。
     *
     * Cache so each item is measured only once.
     * The selection cannot change while the modal dialog is up, so it is kept until the dialog closes.
     * Measuring text requires outlining it, which is too slow to repeat on every preview refresh.
     */
    var measuredBoundsCache = [];

    /**
     * テキストをアウトライン化した状態のバウンズを求めます。
     * 複製をアウトライン化して計測し、計測用のアイテムは必ず削除します。
     *
     * @param {TextFrame} textFrame - 対象のテキスト。
     * @returns {number[]} バウンズ。失敗した場合は元のバウンズ。
     */
    function measureOutlinedTextBounds(textFrame) {
        var duplicatedText = null;
        var outlinedText = null;
        try {
            duplicatedText = textFrame.duplicate();
            outlinedText = duplicatedText.createOutline();
            return outlinedText.geometricBounds;
        } catch (e) {
            logScriptError(getLabel("log.measureTextBounds"), e);
            return textFrame.geometricBounds;
        } finally {
            /* createOutline() は複製を消費するので、アウトラインができたらそちらだけを消す / createOutline() consumes the duplicate, so remove the outline once it exists */
            removeItemSafely(outlinedText || duplicatedText, getLabel("log.removeMeasurementCopy"));
        }
    }

    /**
     * 計測用のバウンズを取得します（テキストはアウトライン化した形状で計測）。
     * 戻り値はキャッシュそのものなので、呼び出し側では書き換えません。
     *
     * @param {PageItem} pageItem - 対象のアイテム。
     * @returns {number[]} バウンズ。
     */
    function getItemMeasurementBounds(pageItem) {
        for (var i = 0; i < measuredBoundsCache.length; i++) {
            if (measuredBoundsCache[i].item === pageItem) return measuredBoundsCache[i].bounds;
        }

        var measuredBounds = (pageItem.typename === "TextFrame")
            ? measureOutlinedTextBounds(pageItem)
            : pageItem.geometricBounds;

        measuredBoundsCache.push({ item: pageItem, bounds: measuredBounds });
        return measuredBounds;
    }

    /**
     * アイテムの中心の X 座標を求めます。
     *
     * @param {PageItem} pageItem - 対象のアイテム。
     * @returns {number} 中心の X 座標。
     */
    function getItemCenterX(pageItem) {
        var itemBounds = getItemMeasurementBounds(pageItem);
        return (itemBounds[0] + itemBounds[2]) / 2;
    }

    /**
     * アイテムの配列を、中心の X 座標で左から右の順に並べ替えます（配列そのものを並べ替えます）。
     *
     * @param {PageItem[]} pageItems - 対象のアイテムの配列。
     * @returns {PageItem[]} 並べ替えた配列。
     */
    function sortItemsLeftToRight(pageItems) {
        return pageItems.sort(function (leftItem, rightItem) {
            return getItemCenterX(leftItem) - getItemCenterX(rightItem);
        });
    }

    /**
     * 隣り合う2つのオブジェクトから、記号を置く位置とサイズを求めます。
     *
     * @param {PageItem} leftItem - 左側のアイテム。
     * @param {PageItem} rightItem - 右側のアイテム。
     * @param {number} heightPercent - 高さ（％）。null なら合計高さをそのまま使います。
     * @param {number} widthPt - 幅（pt）。0以下なら高さから決めます。
     * @param {number} offsetX - 左右の位置調整（pt）。
     * @param {number} offsetY - 上下の位置調整（pt）。
     * @returns {object} 配置情報。アキがない場合は null。
     */
    function computePlacementBetweenItems(leftItem, rightItem, heightPercent, widthPt, offsetX, offsetY) {
        var firstBounds = getItemMeasurementBounds(leftItem);
        var secondBounds = getItemMeasurementBounds(rightItem);

        var leftBounds = (firstBounds[0] < secondBounds[0]) ? firstBounds : secondBounds;
        var rightBounds = (firstBounds[0] < secondBounds[0]) ? secondBounds : firstBounds;

        var gapLeft = leftBounds[2];
        var gapRight = rightBounds[0];
        if (gapRight - gapLeft <= 0) return null;

        var top = Math.max(firstBounds[1], secondBounds[1]);
        var bottom = Math.min(firstBounds[3], secondBounds[3]);
        var totalHeight = top - bottom;
        var usesFullHeight = (heightPercent === null || typeof heightPercent === "undefined");
        var markHeight = usesFullHeight ? totalHeight : (totalHeight * (heightPercent / 100));

        return {
            gapLeft: gapLeft,
            gapRight: gapRight,
            centerX: (gapLeft + gapRight) / 2 + (offsetX || 0),
            centerY: (top + bottom) / 2 + (offsetY || 0),
            totalHeight: totalHeight,
            height: markHeight,
            width: (widthPt > 0) ? widthPt : markHeight,
            /* 矢印系は幅未指定のとき、アキに対する割合で決める / Arrow shapes fall back to a ratio of the gap when no width is set */
            arrowWidth: (widthPt > 0) ? widthPt : ((gapRight - gapLeft) * AUTO_WIDTH_GAP_RATIO)
        };
    }

    /**
     * 隣り合う組を走査し、最も狭いアキと最も低い合計高さを求めます。
     *
     * @param {PageItem[]} sortedItems - 左から右の順に並べたアイテム。
     * @returns {object} minGapWidth / minTotalHeight を持つオブジェクト。アキのある組がなければ null。
     */
    function measureNarrowestGap(sortedItems) {
        var minGapWidth = null;
        var minTotalHeight = null;

        for (var i = 0; i < sortedItems.length - 1; i++) {
            var placement = computePlacementBetweenItems(sortedItems[i], sortedItems[i + 1], null, 0, 0, 0);
            if (!placement) continue;

            var gapWidth = placement.gapRight - placement.gapLeft;
            if (minGapWidth === null || gapWidth < minGapWidth) {
                minGapWidth = gapWidth;
            }
            if (minTotalHeight === null || placement.totalHeight < minTotalHeight) {
                minTotalHeight = placement.totalHeight;
            }
        }

        if (minGapWidth === null) return null;
        return { minGapWidth: minGapWidth, minTotalHeight: minTotalHeight };
    }

    // =========================================
    // 形状定義 / Shape definitions
    // =========================================

    /* 形状の設定の既定値。各形状には、既定と違う項目だけを書きます / Shared defaults; each shape lists only what differs */
    var SHAPE_CONFIG_DEFAULTS = {
        symbol: "",                   /* ラジオボタンの表記 / radio button label */
        helpTip: "",                  /* ラジオボタンの tooltip / radio button tooltip */
        enableCapPanel: true,         /* ［先端］パネル / End Style panel */
        enableStrokeInput: true,      /* ［線幅］（有効なら下限も検証）/ stroke width, validated against the minimum */
        enableHeightInput: false,     /* ［高さ］/ height */
        enableMirror: false,          /* ［左右逆］/ mirror */
        enableGap: false,             /* ［間隔］/ gap */
        enableInset: false,           /* ［凹み］/ inset */
        enableRoundCorners: false,    /* ［角丸］/ rounded corners */
        enableAngle: false,           /* ［角度］/ angle */
        enableFlatChevron: false,     /* ［天地を水平に］/ flat chevron */
        changesSelection: false,      /* 作成時にメニューコマンドで選択を変えるか / whether creation drives menu commands on the selection */
        defaultHeightPercent: 0,
        defaultGap: -1,
        defaultStrokePt: 0.6,
        defaultAngle: 0,
        /* 幅の自動計算：basis が "gap" なら最も狭いアキ、"height" なら高さ（％）を掛けた合計高さに ratio を掛ける
           Automatic width: ratio × the narrowest gap ("gap"), or × the total height scaled by the height percentage ("height") */
        autoWidth: { basis: "gap", ratio: AUTO_WIDTH_GAP_RATIO },
        createMark: null              /* 記号を作成する関数 / function that draws the mark */
    };

    /**
     * 既定値に、形状ごとの設定を重ねたオブジェクトを返します。
     *
     * @param {object} overrides - 既定値と違う項目。
     * @returns {object} 形状の設定。
     */
    function withShapeDefaults(overrides) {
        var shapeConfig = {};
        var settingKey;
        for (settingKey in SHAPE_CONFIG_DEFAULTS) {
            if (SHAPE_CONFIG_DEFAULTS.hasOwnProperty(settingKey)) shapeConfig[settingKey] = SHAPE_CONFIG_DEFAULTS[settingKey];
        }
        for (settingKey in overrides) {
            if (overrides.hasOwnProperty(settingKey)) shapeConfig[settingKey] = overrides[settingKey];
        }
        return shapeConfig;
    }

    /* 形状ごとの設定 / Per-shape settings */
    var SHAPE_CONFIG = {
        triangle: withShapeDefaults({
            symbol: "▶",
            enableCapPanel: false,
            /* 塗りだけで描くので線幅は使いません / Drawn as a fill only, so the stroke width is unused */
            enableStrokeInput: false,
            enableHeightInput: true,
            enableMirror: true,
            enableInset: true,
            enableRoundCorners: true,
            defaultHeightPercent: 30,
            defaultStrokePt: 0.3,
            autoWidth: { basis: "height", ratio: 1 },
            createMark: createTriangleMark
        }),
        chevron: withShapeDefaults({
            symbol: ">",
            enableHeightInput: true,
            enableMirror: true,
            enableFlatChevron: true,
            defaultHeightPercent: 50,
            autoWidth: { basis: "height", ratio: AUTO_WIDTH_CHEVRON_HEIGHT_RATIO },
            createMark: createChevronMark
        }),
        doubleChevron: withShapeDefaults({
            symbol: ">>",
            enableHeightInput: true,
            enableMirror: true,
            enableGap: true,
            enableFlatChevron: true,
            defaultHeightPercent: 50,
            defaultGap: 0,
            /* 幅は山形1つ分。全体の幅は「幅×2＋間隔」になります / The width is per chevron; the pair spans width × 2 + gap */
            autoWidth: { basis: "height", ratio: AUTO_WIDTH_CHEVRON_HEIGHT_RATIO },
            createMark: createDoubleChevronMark
        }),
        dash: withShapeDefaults({
            symbol: "─",
            createMark: createDashMark
        }),
        arrow: withShapeDefaults({
            symbol: "→",
            enableHeightInput: true,
            enableMirror: true,
            defaultHeightPercent: 30,
            createMark: createArrowMark
        }),
        solidArrow: withShapeDefaults({
            symbol: "➡",
            helpTip: formatLabel("tooltip.solidArrow", [SOLID_ARROW_HEIGHT_TO_STROKE_RATIO]),
            /* 天地は線幅から決まるため、高さ（％）は使いません（enableHeightInput は既定の false）/ The height comes from the stroke width, so the percentage stays disabled */
            enableCapPanel: false,
            enableMirror: true,
            changesSelection: true,
            defaultStrokePt: 3.6,
            createMark: createSolidArrowMark
        }),
        arrowSlash: withShapeDefaults({
            symbol: getLabel("radio.arrowSlash"),
            enableHeightInput: true,
            enableMirror: true,
            enableAngle: true,
            defaultHeightPercent: 50,
            defaultAngle: SLASH_ANGLE_DEFAULT,
            createMark: createArrowSlashMark
        }),
        plus: withShapeDefaults({
            symbol: "＋",
            autoWidth: { basis: "gap", ratio: AUTO_WIDTH_GAP_RATIO_SMALL },
            createMark: createPlusMark
        }),
        multiply: withShapeDefaults({
            symbol: "×",
            autoWidth: { basis: "gap", ratio: AUTO_WIDTH_GAP_RATIO_SMALL },
            createMark: createMultiplyMark
        })
    };

    /* ［形状］パネルに並べる順。先頭が初期選択 / Order in the Shape panel; the first one starts selected */
    var SHAPE_ORDER = ["triangle", "chevron", "doubleChevron", "dash", "arrow", "solidArrow", "arrowSlash", "plus", "multiply"];

    // =========================================
    // パス作成のヘルパー / Path creation helpers
    // =========================================

    /**
     * 記号を作成するレイヤー（作業レイヤー）を返します。
     *
     * @returns {Layer} 作業レイヤー。
     */
    function getTargetLayer() {
        return app.activeDocument.activeLayer;
    }

    /**
     * 記号の色を作成します。
     *
     * @returns {CMYKColor} 記号の色。
     */
    function createMarkColor() {
        var markColor = new CMYKColor();
        markColor.cyan = MARK_COLOR_CMYK[0];
        markColor.magenta = MARK_COLOR_CMYK[1];
        markColor.yellow = MARK_COLOR_CMYK[2];
        markColor.black = MARK_COLOR_CMYK[3];
        return markColor;
    }

    /**
     * 線端（と角の形状）を設定します。
     *
     * @param {PathItem} pathItem - 対象のパス。
     * @param {object} endStyle - round（丸型にするか）/ join（角の形状も設定するか）/ flatWhenNotRound（丸型でないとき線端なし・マイターを明示するか）。
     * @returns {void}
     */
    function applyStrokeEnds(pathItem, endStyle) {
        if (endStyle.round) {
            pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
            if (endStyle.join) pathItem.strokeJoin = StrokeJoin.ROUNDENDJOIN;
            return;
        }
        if (endStyle.flatWhenNotRound) {
            pathItem.strokeCap = StrokeCap.BUTTENDCAP;
            if (endStyle.join) pathItem.strokeJoin = StrokeJoin.MITERENDJOIN;
        }
    }

    /**
     * 線だけの開いたパスを作成します。
     *
     * @param {number[][]} points - アンカーポイントの配列。
     * @param {number} strokeWidthPt - 線幅（pt）。
     * @param {object} [endStyle] - applyStrokeEnds に渡す線端の指定。省略時は線端を変えません。
     * @returns {PathItem} 作成したパス。
     */
    function createStrokedPath(points, strokeWidthPt, endStyle) {
        var strokedPath = getTargetLayer().pathItems.add();
        strokedPath.setEntirePath(points);
        strokedPath.closed = false;
        strokedPath.filled = false;
        strokedPath.stroked = true;
        strokedPath.strokeWidth = strokeWidthPt;
        strokedPath.strokeColor = createMarkColor();
        if (endStyle) applyStrokeEnds(strokedPath, endStyle);
        return strokedPath;
    }

    /**
     * 2点を結ぶ直線を、入力値の線幅と線端で作成します。
     *
     * @param {number[]} startPoint - 始点。
     * @param {number[]} endPoint - 終点。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createStrokedLine(startPoint, endPoint, markSettings) {
        return createStrokedPath([startPoint, endPoint], markSettings.strokeWidthPt, { round: markSettings.roundEnds });
    }

    /**
     * 塗りだけの閉じたパスを作成します。
     *
     * @param {number[][]} points - アンカーポイントの配列。
     * @returns {PathItem} 作成したパス。
     */
    function createFilledPath(points) {
        var filledPath = getTargetLayer().pathItems.add();
        filledPath.setEntirePath(points);
        filledPath.closed = true;
        filledPath.filled = true;
        filledPath.fillColor = createMarkColor();
        filledPath.stroked = false;
        return filledPath;
    }

    /**
     * 複数のアイテムを1つのグループにまとめます。
     *
     * @param {PageItem[]} pageItems - まとめるアイテムの配列。
     * @returns {GroupItem} 作成したグループ。
     */
    function groupPageItems(pageItems) {
        var markGroup = getTargetLayer().groupItems.add();
        for (var i = 0; i < pageItems.length; i++) {
            pageItems[i].move(markGroup, ElementPlacement.INSIDE);
        }
        return markGroup;
    }

    /**
     * 角丸のライブエフェクトを適用します。
     * 効果はプラグイン側で解釈されるため、適用に失敗しても角丸なしで続行します。
     *
     * @param {PageItem} pageItem - 対象のアイテム。
     * @param {number} radiusPt - 角丸の半径（pt）。
     * @returns {void}
     */
    function applyRoundCornersEffect(pageItem, radiusPt) {
        runSafely(function () {
            pageItem.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>');
        }, getLabel("log.applyRoundCorners"));
    }

    // =========================================
    // 形状の作成 / Shape creation
    // =========================================
    // 各 create…Mark() は、配置情報（computePlacementBetweenItems の戻り値）と
    // 検証済みの入力値（readMarkSettings の戻り値）を受け取ります。
    // Each create…Mark() takes the placement (from computePlacementBetweenItems)
    // and the validated settings (from readMarkSettings).

    /**
     * ▶ の角丸半径を求めます。
     *
     * @param {number} width - 幅（pt）。
     * @param {number} height - 高さ（pt）。
     * @param {number} insetPt - 凹み（pt）。
     * @returns {number} 角丸の半径（pt）。
     */
    function getTriangleCornerRadius(width, height, insetPt) {
        var radius = Math.min(width, height) * TRIANGLE_CORNER_RADIUS_RATIO;
        if (insetPt > 0) {
            radius = Math.min(radius, insetPt * 0.45);
        }
        return Math.max(1, radius);
    }

    /**
     * ▶（塗りの三角形）を作成します。凹みを指定すると左辺がへこんだ形になります。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createTriangleMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var width = placement.width;
        var height = placement.height;
        var inset = Math.max(0, Math.min(markSettings.insetPt, width * MAX_INSET_RATIO));
        var leftX = centerX - width / 2;
        var rightX = centerX + width / 2;
        var halfHeight = height / 2;

        var points = (inset > 0)
            ? [[rightX, centerY], [leftX, centerY + halfHeight], [leftX + inset, centerY], [leftX, centerY - halfHeight]]
            : [[rightX, centerY], [leftX, centerY + halfHeight], [leftX, centerY - halfHeight]];

        var trianglePath = createFilledPath(points);
        if (markSettings.roundCorners) {
            applyRoundCornersEffect(trianglePath, getTriangleCornerRadius(width, height, inset));
        }
        return trianglePath;
    }

    /**
     * 矢じりの奥行き（水平方向の長さ）を求めます。
     * 高さの半分を基本としつつ、矢じりが幅からはみ出さないよう上限を設けます。
     *
     * @param {number} width - 幅（pt）。
     * @param {number} height - 高さ（pt）。
     * @returns {number} 矢じりの奥行き（pt）。
     */
    function getArrowHeadDepth(width, height) {
        var headDepth = height / 2;
        if (width > 0) headDepth = Math.min(headDepth, width * MAX_ARROW_HEAD_RATIO);
        return headDepth;
    }

    /**
     * →（線の矢印）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createArrowMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var width = placement.arrowWidth;
        var headHalfHeight = placement.height / 2;
        var headDepth = getArrowHeadDepth(width, placement.height);
        var tipX = centerX + width / 2;

        var shaft = createStrokedLine([centerX - width / 2, centerY], [tipX, centerY], markSettings);

        var head = createStrokedPath([
            [tipX - headDepth, centerY + headHalfHeight],
            [tipX, centerY],
            [tipX - headDepth, centerY - headHalfHeight]
        ], markSettings.strokeWidthPt, { round: markSettings.roundEnds, join: true });

        return groupPageItems([shaft, head]);
    }

    /**
     * ➡（塗りの矢印）を作成します。
     * 高さ（％）は使わず、線幅を軸の太さとし、矢じりの天地を線幅から求めます。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。
     */
    function createSolidArrowMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var width = placement.arrowWidth;
        var shaftWidthPt = markSettings.strokeWidthPt;
        var height = shaftWidthPt * SOLID_ARROW_HEIGHT_TO_STROKE_RATIO;
        var headHalfHeight = height / 2;
        var headDepth = getArrowHeadDepth(width, height);
        var tipX = centerX + width / 2;

        var shaft = createStrokedPath([
            [centerX - width / 2, centerY],
            [tipX - headDepth, centerY]
        ], shaftWidthPt);

        var head = createFilledPath([
            [tipX - headDepth, centerY - headHalfHeight],
            [tipX, centerY],
            [tipX - headDepth, centerY + headHalfHeight]
        ]);

        return mergeSolidArrowParts(shaft, head);
    }

    /**
     * ➡ の軸と矢じりを、アウトライン化と合体で1つのパスにまとめます。
     * 選択状態を使うメニューコマンドを呼ぶため、呼び出し側（createMarks）で選択を退避・復元します。
     * メニューコマンドは失敗しても例外の中身が読めないので、ここでは受け止めてフォールバックへ回します。
     * 合成できなかった場合も、軸と矢じりを取り残さないよう1つのグループにまとめて返します。
     *
     * @param {PathItem} shaft - 軸（線）。
     * @param {PathItem} head - 矢じり（塗り）。
     * @returns {PageItem} 合成後のアイテム。
     */
    function mergeSolidArrowParts(shaft, head) {
        var activeDoc = app.activeDocument;
        var outlinedShaft = shaft;

        try {
            activeDoc.selection = [shaft];
            app.executeMenuCommand("Live Outline Stroke");
            if (activeDoc.selection.length > 0) {
                outlinedShaft = activeDoc.selection[0];
            }

            activeDoc.selection = [outlinedShaft, head];
            app.executeMenuCommand("group");
            app.executeMenuCommand("Live Pathfinder Add");

            /* 1つにまとまったときだけ成功とみなす / Treat it as merged only when a single item is left */
            if (activeDoc.selection.length === 1) {
                return activeDoc.selection[0];
            }
        } catch (e) {
            logScriptError(getLabel("log.mergeSolidArrow"), e);
        }

        /* メニューコマンドが効かなかった場合のフォールバック / Fallback when the menu commands did not take effect */
        var fallbackGroup = null;
        runSafely(function () {
            fallbackGroup = groupPageItems([outlinedShaft, head]);
        }, getLabel("log.mergeSolidArrow"));

        return fallbackGroup || outlinedShaft;
    }

    /**
     * ＿\（横線＋斜線）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createArrowSlashMark(placement, markSettings) {
        var centerY = placement.centerY;
        var width = placement.arrowWidth;
        var rise = placement.height / 2;
        var leftX = placement.centerX - width / 2;
        var tipX = placement.centerX + width / 2;

        var slashAngle = markSettings.angleDeg;
        if (isNaN(slashAngle) || slashAngle <= 0) slashAngle = SLASH_ANGLE_DEFAULT;
        if (slashAngle >= SLASH_ANGLE_MAX) slashAngle = SLASH_ANGLE_MAX;

        var slashDx = rise / Math.tan(slashAngle * Math.PI / 180);
        if (!isFinite(slashDx) || slashDx <= 0) slashDx = rise;

        var slashTopX = tipX - slashDx;
        if (slashTopX <= leftX) {
            slashTopX = leftX + Math.max(markSettings.strokeWidthPt, width * 0.15);
        }

        return createStrokedPath([
            [leftX, centerY],
            [tipX, centerY],
            [slashTopX, centerY + rise]
        ], markSettings.strokeWidthPt, { round: markSettings.roundEnds, join: true, flatWhenNotRound: true });
    }

    /**
     * 天地を水平にした > の腕を1本作成します。
     *
     * @param {number} startX - 左端の X 座標。
     * @param {number} startY - 左端の Y 座標。
     * @param {number} tipX - 先端の X 座標。
     * @param {number} tipY - 先端の Y 座標。
     * @param {number} thickness - 腕の太さ（pt）。
     * @returns {PathItem} 作成したパス。
     */
    function createFlatChevronArm(startX, startY, tipX, tipY, thickness) {
        return createFilledPath([
            [startX, startY],
            [startX + thickness, startY],
            [tipX, tipY],
            [tipX - thickness, tipY]
        ]);
    }

    /**
     * > を1つ作成します。［天地を水平に］がONなら塗りの形状、OFFなら線で描きます。
     *
     * @param {number} centerX - 中心の X 座標。
     * @param {number} centerY - 中心の Y 座標。
     * @param {number} width - 山形1つ分の幅（pt）。
     * @param {number} height - 高さ（pt）。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。
     */
    function createSingleChevron(centerX, centerY, width, height, markSettings) {
        var leftX = centerX - width / 2;
        var tipX = centerX + width / 2;
        var halfHeight = height / 2;

        if (markSettings.flatChevron) {
            return groupPageItems([
                createFlatChevronArm(leftX, centerY + halfHeight, tipX, centerY, markSettings.strokeWidthPt),
                createFlatChevronArm(leftX, centerY - halfHeight, tipX, centerY, markSettings.strokeWidthPt)
            ]);
        }

        return createStrokedPath([
            [leftX, centerY + halfHeight],
            [tipX, centerY],
            [leftX, centerY - halfHeight]
        ], markSettings.strokeWidthPt, { round: markSettings.roundEnds, join: true });
    }

    /**
     * >（山形）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。
     */
    function createChevronMark(placement, markSettings) {
        return createSingleChevron(placement.centerX, placement.centerY, placement.width, placement.height, markSettings);
    }

    /**
     * >>（二重の山形）を作成します。幅は山形1つ分で、2つの間を［間隔］だけ空けます。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createDoubleChevronMark(placement, markSettings) {
        var centerOffset = (placement.width + markSettings.chevronGapPt) / 2;
        return groupPageItems([
            createSingleChevron(placement.centerX - centerOffset, placement.centerY, placement.width, placement.height, markSettings),
            createSingleChevron(placement.centerX + centerOffset, placement.centerY, placement.width, placement.height, markSettings)
        ]);
    }

    /**
     * ─（横線）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PathItem} 作成したパス。
     */
    function createDashMark(placement, markSettings) {
        var halfWidth = placement.width / 2;
        return createStrokedLine(
            [placement.centerX - halfWidth, placement.centerY],
            [placement.centerX + halfWidth, placement.centerY],
            markSettings
        );
    }

    /**
     * ＋（十字）を作成します。
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createPlusMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var halfWidth = placement.width / 2;

        return groupPageItems([
            createStrokedLine([centerX - halfWidth, centerY], [centerX + halfWidth, centerY], markSettings),
            createStrokedLine([centerX, centerY + halfWidth], [centerX, centerY - halfWidth], markSettings)
        ]);
    }

    /**
     * ×（斜め十字）を作成します。＋ を45°回転させた形です。
     *
     * ほかの形状と違い、幅は実寸ではなく ＋ と同じ腕の長さを表します（描画幅は幅 × 0.707）。
     * 実寸で揃えると ＋ より大きく見えるため、意図的にこの基準にしています。
     *
     * Unlike the other shapes, the width here is the arm length shared with the plus sign, not the drawn width.
     * Matching the drawn width would make it look larger than the plus sign, so this is intentional.
     *
     * @param {object} placement - 配置情報。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {GroupItem} 作成したグループ。
     */
    function createMultiplyMark(placement, markSettings) {
        var centerX = placement.centerX;
        var centerY = placement.centerY;
        var diagonalOffset = (placement.width / 2) * Math.SQRT2 / 2;

        return groupPageItems([
            createStrokedLine([centerX - diagonalOffset, centerY + diagonalOffset], [centerX + diagonalOffset, centerY - diagonalOffset], markSettings),
            createStrokedLine([centerX - diagonalOffset, centerY - diagonalOffset], [centerX + diagonalOffset, centerY + diagonalOffset], markSettings)
        ]);
    }

    // =========================================
    // 記号の作成 / Creating marks
    // =========================================

    /**
     * 隣り合う2つのオブジェクトの間に記号を1つ作成し、［左右逆］がONなら反転します。
     *
     * @param {PageItem} leftItem - 左側のアイテム。
     * @param {PageItem} rightItem - 右側のアイテム。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem} 作成したアイテム。アキがない場合は null。
     */
    function createMarkBetweenItems(leftItem, rightItem, markSettings) {
        var placement = computePlacementBetweenItems(leftItem, rightItem, markSettings.heightPercent, markSettings.widthPt, markSettings.offsetX, markSettings.offsetY);
        if (!placement) return null;

        var markItem = markSettings.shapeConfig.createMark(placement, markSettings);
        if (markSettings.mirror) {
            markItem.resize(-100, 100, true, true, true, true, 100, Transformation.CENTER);
        }
        return markItem;
    }

    /**
     * 隣り合うすべての組に記号を作成します。
     *
     * @param {PageItem[]} sortedItems - 左から右の順に並べたアイテム。
     * @param {object} markSettings - 検証済みの入力値。
     * @returns {PageItem[]} 作成したアイテムの配列。
     */
    function createMarks(sortedItems, markSettings) {
        var createdItems = [];

        /* ➡ だけは選択状態を使うメニューコマンドを呼ぶため、ここで1回だけ退避・復元する / Only the solid arrow drives menu commands, so save and restore the selection once */
        var previousSelection = markSettings.shapeConfig.changesSelection
            ? collectionToArray(app.activeDocument.selection)
            : null;

        try {
            for (var i = 0; i < sortedItems.length - 1; i++) {
                var markItem = createMarkBetweenItems(sortedItems[i], sortedItems[i + 1], markSettings);
                if (markItem) createdItems.push(markItem);
            }
        } finally {
            if (previousSelection) restoreSelection(previousSelection);
        }
        return createdItems;
    }

    // =========================================
    // ダイアログの構築 / Building the dialog
    // =========================================

    /**
     * 項目名・入力欄・単位を1行にまとめて追加します。
     *
     * @param {Panel} parentPanel - 追加先のパネル。
     * @param {string} labelPath - 項目名のドットパス。
     * @param {string} unitText - 入力欄の右に置く単位表記。
     * @param {string} [initialText] - 入力欄の初期値。省略時は空欄。
     * @returns {EditText} 追加した入力欄。行のグループは parent で参照できます。
     */
    function addLabeledField(parentPanel, labelPath, unitText, initialText) {
        var fieldRowGroup = parentPanel.add("group");
        setupRow(fieldRowGroup, "left", FIELD_ROW_SPACING);

        var fieldCaption = fieldRowGroup.add("statictext", undefined, labelText(labelPath));
        fieldCaption.preferredSize = [LABEL_COLUMN_WIDTH, -1];
        fieldCaption.justify = "right";

        var inputField = fieldRowGroup.add("edittext", undefined, initialText || "");
        inputField.characters = FIELD_CHARACTERS;

        fieldRowGroup.add("statictext", undefined, unitText);

        return inputField;
    }

    /**
     * ［形状］パネルを作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildShapePanel(parentColumn, dialogControls) {
        var shapePanel = parentColumn.add("panel", undefined, getLabel("panel.shape"));
        setupPanel(shapePanel, RADIO_SPACING);

        dialogControls.shapeRadios = {};
        for (var i = 0; i < SHAPE_ORDER.length; i++) {
            var shapeConfig = SHAPE_CONFIG[SHAPE_ORDER[i]];
            var shapeRadio = shapePanel.add("radiobutton", undefined, shapeConfig.symbol);
            if (shapeConfig.helpTip) shapeRadio.helpTip = shapeConfig.helpTip;
            dialogControls.shapeRadios[SHAPE_ORDER[i]] = shapeRadio;
        }
        dialogControls.shapeRadios[SHAPE_ORDER[0]].value = true;

        var mirrorGroup = shapePanel.add("group");
        setupRow(mirrorGroup, "left");
        mirrorGroup.margins = [0, MIRROR_ROW_TOP_MARGIN, 0, 0];

        dialogControls.mirrorCheckbox = mirrorGroup.add("checkbox", undefined, getLabel("checkbox.mirrorHorizontal"));
        dialogControls.mirrorCheckbox.helpTip = getLabel("tooltip.mirrorHorizontal");
    }

    /**
     * ［先端］パネルを作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildCapStylePanel(parentColumn, dialogControls) {
        var capStylePanel = parentColumn.add("panel", undefined, getLabel("panel.capStyle"));
        setupPanel(capStylePanel, RADIO_SPACING);

        dialogControls.capStylePanel = capStylePanel;
        dialogControls.capNoneRadio = capStylePanel.add("radiobutton", undefined, getLabel("radio.capNone"));
        dialogControls.capNoneRadio.helpTip = getLabel("tooltip.capNone");
        dialogControls.capRoundRadio = capStylePanel.add("radiobutton", undefined, getLabel("radio.capRound"));
        dialogControls.capRoundRadio.helpTip = getLabel("tooltip.capRound");
        dialogControls.capNoneRadio.value = true;
    }

    /**
     * ［オプション］パネルを作成します。
     * 入力欄の初期値は、形状に合わせて applyShapeDefaults() で入れます。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildOptionsPanel(parentColumn, dialogControls) {
        var optionsPanel = parentColumn.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, FIELD_ROW_SPACING);

        dialogControls.heightField = addLabeledField(optionsPanel, "fieldLabel.height", "%");
        dialogControls.heightField.helpTip = formatLabel("tooltip.height", [MAX_HEIGHT_PERCENT]);
        dialogControls.heightField.active = true;

        dialogControls.widthField = addLabeledField(optionsPanel, "fieldLabel.width", rulerUnitInfo.label);
        dialogControls.widthField.helpTip = getLabel("tooltip.width");

        dialogControls.insetField = addLabeledField(optionsPanel, "fieldLabel.inset", rulerUnitInfo.label);
        dialogControls.insetField.helpTip = formatLabel("tooltip.inset", [roundDisplayValue(MAX_INSET_RATIO * 100)]);

        dialogControls.gapField = addLabeledField(optionsPanel, "fieldLabel.gap", rulerUnitInfo.label);
        dialogControls.gapField.helpTip = getLabel("tooltip.gap");

        dialogControls.strokeField = addLabeledField(optionsPanel, "fieldLabel.strokeWidth", strokeUnitInfo.label);

        dialogControls.angleField = addLabeledField(optionsPanel, "fieldLabel.angle", "°");
        dialogControls.angleField.helpTip = formatLabel("tooltip.angle", [SLASH_ANGLE_MAX]);

        var optionCheckboxGroup = optionsPanel.add("group");
        optionCheckboxGroup.orientation = "column";
        optionCheckboxGroup.alignChildren = ["left", "top"];
        optionCheckboxGroup.spacing = FIELD_ROW_SPACING;
        optionCheckboxGroup.margins = [OPTION_CHECKBOX_INDENT, OPTION_CHECKBOX_TOP_MARGIN, 0, 0];

        dialogControls.flatChevronCheckbox = optionCheckboxGroup.add("checkbox", undefined, getLabel("checkbox.flatChevron"));
        dialogControls.flatChevronCheckbox.helpTip = getLabel("tooltip.flatChevron");

        dialogControls.roundCornersCheckbox = optionCheckboxGroup.add("checkbox", undefined, getLabel("checkbox.roundCorners"));
        dialogControls.roundCornersCheckbox.helpTip = getLabel("tooltip.roundCorners");
        dialogControls.roundCornersCheckbox.value = true;
    }

    /**
     * ［位置調整］パネルを作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildPositionPanel(parentColumn, dialogControls) {
        var positionPanel = parentColumn.add("panel", undefined, getLabel("panel.position"));
        setupPanel(positionPanel, FIELD_ROW_SPACING);

        dialogControls.offsetXField = addLabeledField(positionPanel, "fieldLabel.offsetX", rulerUnitInfo.label, "0");
        dialogControls.offsetXField.helpTip = getLabel("tooltip.offsetX");
        dialogControls.offsetYField = addLabeledField(positionPanel, "fieldLabel.offsetY", rulerUnitInfo.label, "0");
        dialogControls.offsetYField.helpTip = getLabel("tooltip.offsetY");
    }

    /**
     * ［プレビュー］とボタンの行を作成します。
     * ［キャンセル］は name: "cancel" の既定動作で閉じ、プレビューは onClose で片付けます。
     *
     * @param {Window} markDialog - 追加先のダイアログ。
     * @param {object} dialogControls - コントロールを登録するオブジェクト。
     * @returns {void}
     */
    function buildButtonRow(markDialog, dialogControls) {
        var btnRowGroup = markDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        dialogControls.previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogControls.btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    }

    /**
     * ダイアログの中身を組み立て、コントロールをまとめて返します。
     *
     * @param {Window} markDialog - 対象のダイアログ。
     * @returns {object} すべてのコントロールを持つオブジェクト。
     */
    function buildDialogControls(markDialog) {
        var columnsGroup = markDialog.add("group");
        setupRow(columnsGroup, "fill", COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "top"];

        var leftColumn = columnsGroup.add("group");
        setupColumn(leftColumn, COLUMN_PANEL_SPACING);

        var rightColumn = columnsGroup.add("group");
        setupColumn(rightColumn, COLUMN_PANEL_SPACING);

        var dialogControls = {};
        buildShapePanel(leftColumn, dialogControls);
        buildCapStylePanel(leftColumn, dialogControls);
        buildOptionsPanel(rightColumn, dialogControls);
        buildPositionPanel(rightColumn, dialogControls);
        buildButtonRow(markDialog, dialogControls);
        return dialogControls;
    }

    // =========================================
    // ダイアログの状態と入力値 / Dialog state and input values
    // =========================================

    /**
     * 選択中の形状キーを返します。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @returns {string} SHAPE_CONFIG のキー。
     */
    function getSelectedShapeKey(dialogControls) {
        for (var i = 0; i < SHAPE_ORDER.length; i++) {
            if (dialogControls.shapeRadios[SHAPE_ORDER[i]].value) return SHAPE_ORDER[i];
        }
        return SHAPE_ORDER[0];
    }

    /**
     * ［先端］を［丸型］か［なし］に切り替えます。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @param {boolean} isRound - ［丸型］にするなら true。
     * @returns {void}
     */
    function setRoundCaps(dialogControls, isRound) {
        dialogControls.capRoundRadio.value = isRound;
        dialogControls.capNoneRadio.value = !isRound;
    }

    /**
     * チェックボックスの有効・無効を切り替え、無効にしたときはOFFに戻します。
     *
     * @param {Checkbox} checkbox - 対象のチェックボックス。
     * @param {boolean} isAvailable - 有効にするなら true。
     * @returns {void}
     */
    function setCheckboxAvailability(checkbox, isAvailable) {
        checkbox.enabled = isAvailable;
        if (!isAvailable) checkbox.value = false;
    }

    /**
     * 形状に応じて、各コントロールの有効・無効を切り替えます。
     * 入力欄は、項目名と単位表記もまとめてグレーにするため行ごと切り替えます。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @param {object} shapeConfig - 形状の設定。
     * @returns {void}
     */
    function applyShapeEnabledStates(dialogControls, shapeConfig) {
        dialogControls.capStylePanel.enabled = shapeConfig.enableCapPanel;
        if (!shapeConfig.enableCapPanel) setRoundCaps(dialogControls, false);

        dialogControls.heightField.parent.enabled = shapeConfig.enableHeightInput;
        dialogControls.strokeField.parent.enabled = shapeConfig.enableStrokeInput;
        dialogControls.gapField.parent.enabled = shapeConfig.enableGap;
        dialogControls.insetField.parent.enabled = shapeConfig.enableInset;
        dialogControls.angleField.parent.enabled = shapeConfig.enableAngle;

        setCheckboxAvailability(dialogControls.flatChevronCheckbox, shapeConfig.enableFlatChevron);
        setCheckboxAvailability(dialogControls.mirrorCheckbox, shapeConfig.enableMirror);
        setCheckboxAvailability(dialogControls.roundCornersCheckbox, shapeConfig.enableRoundCorners);
    }

    /**
     * 入力エラーを知らせ、対象の入力欄にフォーカスを移します。
     *
     * @param {EditText} invalidField - 対象の入力欄。
     * @param {string} alertPath - 表示するメッセージのドットパス。
     * @param {boolean} showAlert - 警告を表示するなら true。
     * @param {Array} [replacements] - メッセージに差し込む値。
     * @returns {null} 呼び出し元がそのまま返せるよう null を返します。
     */
    function reportInvalidValue(invalidField, alertPath, showAlert, replacements) {
        if (showAlert) {
            alert(formatLabel(alertPath, replacements));
            invalidField.active = true;
        }
        return null;
    }

    /**
     * ダイアログの入力値をまとめて読み取り、検証します。
     *
     * @param {object} dialogControls - ダイアログのコントロール。
     * @param {boolean} showAlert - 不正な値のときに警告を表示するなら true。
     * @returns {object} 検証済みの入力値（寸法は pt）。不正な場合は null。
     */
    function readMarkSettings(dialogControls, showAlert) {
        var shapeConfig = SHAPE_CONFIG[getSelectedShapeKey(dialogControls)];
        var heightPercent = null;
        var widthPt = parseFieldToPt(dialogControls.widthField, rulerUnitInfo, false);
        var insetPt = 0;
        var offsetX = parseFieldToPt(dialogControls.offsetXField, rulerUnitInfo, true);
        var offsetY = parseFieldToPt(dialogControls.offsetYField, rulerUnitInfo, true);
        var strokeWidthPt = parseFieldToPt(dialogControls.strokeField, strokeUnitInfo, false);
        var angleDeg = shapeConfig.enableAngle ? parseFloat(dialogControls.angleField.text) : 0;
        var chevronGapPt = parseFieldToPt(dialogControls.gapField, rulerUnitInfo, true);

        if (shapeConfig.enableInset) {
            insetPt = parseFieldToPt(dialogControls.insetField, rulerUnitInfo, false);
            if (isNaN(insetPt)) {
                return reportInvalidValue(dialogControls.insetField, "alert.invalidValue", showAlert);
            }
            /* 幅が明示されている場合は、その割合まで凹みを抑える / Clamp the inset when the width is set explicitly */
            if (widthPt > 0) {
                insetPt = Math.min(insetPt, widthPt * MAX_INSET_RATIO);
            }
        }

        if (shapeConfig.enableHeightInput) {
            heightPercent = parseFloat(dialogControls.heightField.text);
            if (isNaN(heightPercent) || heightPercent <= 0) {
                return reportInvalidValue(dialogControls.heightField, "alert.positiveNumber", showAlert);
            }
            if (heightPercent > MAX_HEIGHT_PERCENT) {
                return reportInvalidValue(dialogControls.heightField, "alert.maxHeight", showAlert, [MAX_HEIGHT_PERCENT]);
            }
        }

        if (isNaN(widthPt)) {
            return reportInvalidValue(dialogControls.widthField, "alert.widthPositive", showAlert);
        }
        if (isNaN(offsetX)) {
            return reportInvalidValue(dialogControls.offsetXField, "alert.invalidValue", showAlert);
        }
        if (isNaN(offsetY)) {
            return reportInvalidValue(dialogControls.offsetYField, "alert.invalidValue", showAlert);
        }
        if (isNaN(strokeWidthPt)) {
            return reportInvalidValue(dialogControls.strokeField, "alert.invalidValue", showAlert);
        }
        /* 塗りだけの形状は線幅を使わないので、下限を求めない / Fill-only shapes ignore the stroke width, so the minimum does not apply */
        if (shapeConfig.enableStrokeInput && strokeWidthPt < MIN_STROKE_WIDTH_PT) {
            var minStrokeText = roundDisplayValueForUnit(convertPtToUnitValue(MIN_STROKE_WIDTH_PT, strokeUnitInfo), strokeUnitInfo);
            return reportInvalidValue(dialogControls.strokeField, "alert.strokePositive", showAlert, [minStrokeText, strokeUnitInfo.label]);
        }
        if (shapeConfig.enableAngle && isNaN(angleDeg)) {
            return reportInvalidValue(dialogControls.angleField, "alert.invalidValue", showAlert);
        }

        return {
            shapeConfig: shapeConfig,
            heightPercent: heightPercent,
            widthPt: widthPt,
            insetPt: insetPt,
            offsetX: offsetX,
            offsetY: offsetY,
            strokeWidthPt: strokeWidthPt,
            angleDeg: angleDeg,
            chevronGapPt: isNaN(chevronGapPt) ? 0 : chevronGapPt,
            roundEnds: dialogControls.capRoundRadio.value,
            roundCorners: dialogControls.roundCornersCheckbox.value,
            flatChevron: dialogControls.flatChevronCheckbox.value,
            mirror: dialogControls.mirrorCheckbox.value
        };
    }

    /**
     * 矢印キーで入力値を増減できるようにします。
     * 基本の増減量は単位に合わせて決まり、Shift はその10倍、Option（Alt）は10分の1で増減します。
     * 増減したあとは入力欄の onChanging を呼び、手入力と同じ後処理を通します。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {boolean} allowNegative - 負の値を許可するなら true。
     * @param {object} [unitInfo] - 単位情報。％や度など単位のない入力欄では省略します。
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, unitInfo) {
        editText.addEventListener("keydown", function (event) {
            var isUp = (event.keyName === "Up");
            var isDown = (event.keyName === "Down");
            if (!isUp && !isDown) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var step = getArrowKeyStep(unitInfo);
            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.shiftKey) {
                /* 増減量10個分の倍数へ丸めながら増減 / Snap to multiples of ten steps while stepping */
                var coarseStep = step * 10;
                value = isUp
                    ? Math.ceil((value + step) / coarseStep) * coarseStep
                    : Math.floor((value - step) / coarseStep) * coarseStep;
            } else {
                var fineStep = keyboard.altKey ? step / 10 : step;
                value = isUp ? value + fineStep : value - fineStep;
                /* 増減量の刻みに揃える / Snap to the step grid */
                value = Math.round(value / fineStep) * fineStep;
            }
            event.preventDefault();

            if (!allowNegative && value < 0) value = 0;

            editText.text = String(roundDisplayValueForUnit(value, unitInfo));
            if (editText.onChanging) editText.onChanging();
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確認してダイアログを表示します。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocument"));
            return;
        }

        var activeDoc = app.activeDocument;

        /* ロック・非表示のレイヤーには作成できないので、先に知らせる / Nothing can be created on a locked or hidden layer */
        if (activeDoc.activeLayer.locked || !activeDoc.activeLayer.visible) {
            alert(getLabel("alert.lockedLayer"));
            return;
        }

        var sortedItems = collectionToArray(activeDoc.selection);
        if (sortedItems.length < 2) {
            alert(getLabel("alert.selectTwoObjects"));
            return;
        }
        sortItemsLeftToRight(sortedItems);

        /* どの組にもアキがなければ記号を作成できないので、ダイアログを出す前に知らせる / Fail early when no pair has a usable gap */
        var narrowestGap = measureNarrowestGap(sortedItems);
        if (!narrowestGap) {
            alert(getLabel("alert.noGap"));
            return;
        }

        showMarkDialog(activeDoc, sortedItems, narrowestGap);
    }

    /**
     * ダイアログを表示し、プレビューと［OK］で記号を作成します。
     *
     * @param {Document} activeDoc - 対象のドキュメント。
     * @param {PageItem[]} sortedItems - 左から右の順に並べたアイテム。
     * @param {object} narrowestGap - measureNarrowestGap() の戻り値（幅の自動計算に使います）。
     * @returns {void}
     */
    function showMarkDialog(activeDoc, sortedItems, narrowestGap) {
        var markDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(markDialog);

        var dialogControls = buildDialogControls(markDialog);

        /* 幅を手入力したかどうか（手入力後は自動計算しない）/ Whether the width was typed in (auto-calculation stops once it is) */
        var widthManuallySet = false;

        /* プレビューで作成したアイテム / Items created for the preview */
        var previewItems = [];

        /**
         * 現在の形状と高さから、幅の自動計算値（pt）を求めます。
         *
         * @returns {number} 幅（pt）。高さが不正で求められない場合は null。
         */
        function computeAutoWidthPt() {
            var shapeConfig = SHAPE_CONFIG[getSelectedShapeKey(dialogControls)];
            var heightPercent = parseFloat(dialogControls.heightField.text);
            if (shapeConfig.enableHeightInput && (isNaN(heightPercent) || heightPercent <= 0)) return null;

            var basePt = (shapeConfig.autoWidth.basis === "gap")
                ? narrowestGap.minGapWidth
                : narrowestGap.minTotalHeight * (heightPercent / 100);
            return roundDisplayValue(basePt * shapeConfig.autoWidth.ratio);
        }

        /**
         * 幅の入力欄に、自動計算した値を表示します。
         *
         * @returns {void}
         */
        function applyAutoWidthToField() {
            var autoWidthPt = computeAutoWidthPt();
            if (autoWidthPt === null) return;
            setFieldFromPt(dialogControls.widthField, autoWidthPt, rulerUnitInfo);
        }

        /**
         * 幅を自動計算に戻します。
         *
         * @returns {void}
         */
        function resetWidthToAuto() {
            widthManuallySet = false;
            applyAutoWidthToField();
        }

        /**
         * 実際に使われる幅（pt）を求めます。手入力があればその値を優先します。
         *
         * @returns {number} 幅（pt）。求められない場合は 0。
         */
        function getEffectiveWidthPt() {
            var typedWidthPt = parseFieldToPt(dialogControls.widthField, rulerUnitInfo, false);
            if (typedWidthPt > 0) return typedWidthPt;

            var autoWidthPt = computeAutoWidthPt();
            return (autoWidthPt > 0) ? autoWidthPt : 0;
        }

        /**
         * 凹みが幅の割合を超えないよう、入力欄の値を抑えます。
         *
         * @returns {void}
         */
        function clampInsetToWidth() {
            var insetValue = parseFloat(dialogControls.insetField.text);
            if (isNaN(insetValue) || insetValue < 0) return;

            var effectiveWidthPt = getEffectiveWidthPt();
            var maxInsetPt = effectiveWidthPt * MAX_INSET_RATIO;
            if (effectiveWidthPt > 0 && convertValueToPt(insetValue, rulerUnitInfo) > maxInsetPt) {
                setFieldFromPt(dialogControls.insetField, maxInsetPt, rulerUnitInfo);
            }
        }

        /**
         * 選択中の形状に合わせて、入力欄の初期値とコントロールの有効・無効を更新します。
         * 形状ごとに適切な値が違うため、切り替えるたびにその形状の初期値へ戻します。
         *
         * @returns {void}
         */
        function applyShapeDefaults() {
            var shapeConfig = SHAPE_CONFIG[getSelectedShapeKey(dialogControls)];

            if (shapeConfig.enableHeightInput) {
                dialogControls.heightField.text = String(shapeConfig.defaultHeightPercent);
            }
            resetWidthToAuto();
            setFieldFromPt(dialogControls.gapField, shapeConfig.defaultGap, rulerUnitInfo);
            setFieldFromPt(dialogControls.insetField, 0, rulerUnitInfo);
            setFieldFromPt(dialogControls.strokeField, shapeConfig.defaultStrokePt, strokeUnitInfo);
            dialogControls.angleField.text = String(shapeConfig.defaultAngle);

            applyShapeEnabledStates(dialogControls, shapeConfig);
        }

        /**
         * 形状を切り替えたときに、初期値と位置調整をリセットします。
         *
         * @returns {void}
         */
        function handleShapeChange() {
            dialogControls.offsetXField.text = "0";
            dialogControls.offsetYField.text = "0";
            applyShapeDefaults();
            updatePreview();
        }

        /**
         * プレビューで作成したアイテムを削除します。
         *
         * @returns {void}
         */
        function removePreview() {
            if (previewItems.length === 0) return;

            for (var i = 0; i < previewItems.length; i++) {
                removeItemSafely(previewItems[i], getLabel("log.removePreviewItem"));
            }
            previewItems = [];
            app.redraw();
        }

        /**
         * プレビューを作り直します。
         *
         * @returns {void}
         */
        function updatePreview() {
            removePreview();
            if (!dialogControls.previewCheckbox.value) return;

            var markSettings = readMarkSettings(dialogControls, false);
            if (!markSettings) return;

            previewItems = createMarks(sortedItems, markSettings);
            app.redraw();
        }

        /**
         * ［OK］で記号を確定します。プレビューがあれば、それをそのまま結果にします。
         *
         * @returns {void}
         */
        function confirmMarks() {
            var markSettings = readMarkSettings(dialogControls, true);
            if (!markSettings) return;

            /* onClose で消されないよう、プレビューを手放してから閉じる / Take over the preview so onClose does not remove it */
            var createdItems = previewItems;
            previewItems = [];
            markDialog.close(1);

            if (createdItems.length === 0) {
                createdItems = createMarks(sortedItems, markSettings);
                if (createdItems.length === 0) {
                    alert(getLabel("alert.noGap"));
                    return;
                }
            }
            activeDoc.selection = createdItems;
        }

        /**
         * 入力欄の変更と矢印キーの増減を登録します。
         *
         * @returns {void}
         */
        function bindFieldEvents() {
            dialogControls.heightField.onChanging = function () {
                if (!widthManuallySet) applyAutoWidthToField();
                updatePreview();
            };

            dialogControls.widthField.onChanging = function () {
                if (dialogControls.widthField.text === "") {
                    resetWidthToAuto();
                } else {
                    widthManuallySet = true;
                }
                updatePreview();
            };

            dialogControls.insetField.onChanging = function () {
                clampInsetToWidth();
                updatePreview();
            };

            var previewOnlyFields = [
                dialogControls.gapField,
                dialogControls.strokeField,
                dialogControls.angleField,
                dialogControls.offsetXField,
                dialogControls.offsetYField
            ];
            for (var i = 0; i < previewOnlyFields.length; i++) {
                previewOnlyFields[i].onChanging = updatePreview;
            }

            /* ％と度は単位を持たないため、増減量は1固定 / Percent and degrees have no unit, so they step by one */
            changeValueByArrowKey(dialogControls.heightField, false);
            changeValueByArrowKey(dialogControls.angleField, false);
            changeValueByArrowKey(dialogControls.widthField, false, rulerUnitInfo);
            changeValueByArrowKey(dialogControls.insetField, false, rulerUnitInfo);
            changeValueByArrowKey(dialogControls.strokeField, false, strokeUnitInfo);
            changeValueByArrowKey(dialogControls.gapField, true, rulerUnitInfo);
            changeValueByArrowKey(dialogControls.offsetXField, true, rulerUnitInfo);
            changeValueByArrowKey(dialogControls.offsetYField, true, rulerUnitInfo);
        }

        /**
         * 形状・先端・チェックボックスのクリックを登録します。
         *
         * @returns {void}
         */
        function bindOptionEvents() {
            for (var i = 0; i < SHAPE_ORDER.length; i++) {
                dialogControls.shapeRadios[SHAPE_ORDER[i]].onClick = handleShapeChange;
            }

            var previewToggles = [
                dialogControls.capNoneRadio,
                dialogControls.capRoundRadio,
                dialogControls.mirrorCheckbox,
                dialogControls.flatChevronCheckbox,
                dialogControls.roundCornersCheckbox,
                dialogControls.previewCheckbox
            ];
            for (var j = 0; j < previewToggles.length; j++) {
                previewToggles[j].onClick = updatePreview;
            }
        }

        /**
         * キーボードショートカット（F：先端なし、R：丸型、V：左右逆）を登録します。
         *
         * @returns {void}
         */
        function bindShortcutKeys() {
            var shortcutActions = {
                F: function () {
                    if (!dialogControls.capStylePanel.enabled) return false;
                    setRoundCaps(dialogControls, false);
                    return true;
                },
                R: function () {
                    if (!dialogControls.capStylePanel.enabled) return false;
                    setRoundCaps(dialogControls, true);
                    return true;
                },
                V: function () {
                    if (!dialogControls.mirrorCheckbox.enabled) return false;
                    dialogControls.mirrorCheckbox.value = !dialogControls.mirrorCheckbox.value;
                    return true;
                }
            };

            markDialog.addEventListener("keydown", function (event) {
                /* Cmd+V などの修飾キー付きの入力は横取りしない / Do not swallow modified keystrokes such as Cmd+V */
                var keyboard = ScriptUI.environment.keyboardState;
                if (keyboard.metaKey || keyboard.ctrlKey || keyboard.altKey || keyboard.shiftKey) return;

                var shortcutAction = shortcutActions[event.keyName];
                if (!shortcutAction || !shortcutAction()) return;
                updatePreview();
                event.preventDefault();
            });
        }

        bindFieldEvents();
        bindOptionEvents();
        bindShortcutKeys();
        dialogControls.btnOK.onClick = confirmMarks;
        markDialog.onClose = removePreview;

        applyShapeDefaults();

        /* 表示前にレイアウトを確定して描画欠けを防ぐ / Settle the layout before showing to avoid partial rendering */
        markDialog.layout.layout(true);
        markDialog.layout.resize();

        markDialog.show();
    }

    main();

})();
