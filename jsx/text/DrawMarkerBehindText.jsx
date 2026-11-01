#target illustrator
#targetengine "DrawMarkerBehindTextEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.DrawMarkerBehindText
Name=文字にマーカー・アンダーライン・取り消し線を引く
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-11-01
Description-BEGIN
文字ツールで選んだ文字やテキストオブジェクトの背面に、字面に合わせてマーカー・アンダーライン・取り消し線などの線を引きます。
種類・カラー・線幅はダイアログボックスで選び、設定はプリセットとして保存できます。

Draws a marker, underline, strikethrough or similar line behind the characters selected with the Type tool or behind text objects, fitted to their glyphs.
Choose the style, color and weight in the dialog box, and save the settings as presets.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

文字ツールで選んだ文字やテキストオブジェクトの背面に、字面に合わせてマーカー・アンダーライン・取り消し線などの線を引きます。
種類・カラー・線幅はダイアログボックスで選び、設定はプリセットとして保存できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawMarkerBehindText.md

### Overview

Draws a marker, underline, strikethrough or similar line behind the characters selected with the Type tool or behind text objects, fitted to their glyphs.
Choose the style, color and weight in the dialog box, and save the settings as presets.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawMarkerBehindText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DrawMarkerBehindText";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-01";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "文字にマーカー・アンダーライン・取り消し線を引く"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawMarkerBehindText.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawMarkerBehindText.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DrawMarkerBehindText.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author デザインゴリラ（オリジナルアイデア） / original idea
 * @discussion https://note.com/gorilla_gg/n/naa26b079cb75
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［黄色］［赤］のカラーと、［カスタム］の最初のカラー（色見本のクリックで変えられる）
       Colors of Yellow and Red, and the initial Custom color (changeable by clicking the swatch) */
    var MARKER_COLOR_PRESETS = {
        yellow: { cmyk: [0, 0, 100, 0], rgb: [255, 241, 0] },   /* 黄色（Y100） / yellow (Y100) */
        red:    { cmyk: [0, 100, 100, 0], rgb: [230, 0, 18] },  /* 赤（M100Y100） / red (M100Y100) */
        custom: { cmyk: [100, 0, 0, 0], rgb: [0, 160, 233] }    /* カスタムの最初の色（C100） / initial custom color (C100) */
    };

    /* 開いたときに選んでおくカラー（"yellow" / "red" / "textColor" / "custom"） / Color chosen when the dialog box opens */
    var DEFAULT_COLOR_CHOICE = "yellow";

    /* ［文字色］の線の不透明度の初期値（%） / Initial opacity (%) of the line in the text color */
    var TEXT_COLOR_OPACITY = 30;

    /* 線幅の初期値：字面の高さに対する比率 / Initial line width as a ratio of the glyph height */
    var MARKER_WEIGHT_RATIO = 0.4;

    /* 線の種類の既定（"underline"：アンダーライン / "marker"：マーカー / "straightShear"：直線＋シアー / "zigzag"：ギザギザ /
       "fill"：塗りつぶし / "box"：長方形 / "charBox"：1文字ずつの長方形 / "strikethrough"：取り消し線） / Default line style */
    var DEFAULT_LINE_STYLE = "marker";

    /* ［長方形］［1文字ずつの長方形］の余白の初期値：字面の高さに対する比率 / Initial box padding as a ratio of the glyph height */
    var BOX_PADDING_RATIO = 0.08;

    /* ［文字を白抜きに］の初期値と、白抜きにする色 / Initial state of White Text, and the color it applies */
    var DEFAULT_WHITE_TEXT = false;
    var WHITE_TEXT_COLOR = { rgb: [255, 255, 255], cmyk: [0, 0, 0, 0] };

    /* ［長方形］［1文字ずつの長方形］の角丸の半径の初期値（pt。0 で角丸なし） / Initial corner radius of the boxes (pt; 0 for square corners) */
    var DEFAULT_CORNER_RADIUS = 0;

    /* ［直線＋シアー］の角度の初期値（度。プラスで右上がり） / Initial shear angle in degrees (positive slants to the upper right) */
    var DEFAULT_SHEAR_ANGLE = 30;

    /* ギザギザ（ジグザグ効果）の初期値 / Initial values of the zigzag (Zig Zag effect) */
    var DEFAULT_ZIGZAG_SIZE = 3;      /* 大きさ（pt） / size (pt) */
    var DEFAULT_ZIGZAG_RIDGES = 10;   /* 折り返し（線の端から端まで） / ridges from end to end */

    /* 線の既定（"solidLine"：直線 / "dashedLine"：破線 / "dottedLine"：ドット点線） / Default line type */
    var DEFAULT_DASH_STYLE = "solidLine";

    /* 破線の線分と間隔、ドット点線の点の間隔（中心から中心）：線幅に対する比率
       Dash and gap of Dashed, and dot pitch (center to center) of Dotted, as ratios of the line width */
    var DASHED_LINE_RATIOS = [2, 1];
    var DOTTED_LINE_PITCH_RATIO = 2;

    /* 線端の既定（"buttCap"：なし / "roundCap"：丸型） / Default line cap */
    var DEFAULT_STROKE_CAP = "buttCap";

    /* ダイアログボックスを開いたときにプレビューするか / Preview when the dialog box opens */
    var PREVIEW_ON_OPEN = true;

    /* 最初から用意するプリセット（ドロップダウンの先頭にこの順で並ぶ。削除はできず、同じ名前で保存したものが優先）。
       書き方は保存したプリセットと同じで、長さは〜Pt のかわりに〜Ratio（先頭の対象の字面の高さに対する比率）でも書ける。
       customColor は { rgb, cmyk } で、ドキュメントのカラーモードに合わせて使い分ける
       Built-in presets (listed first in this order; they cannot be deleted, and a saved preset of the same name wins).
       Written like saved presets; lengths may use ...Ratio (of the first target's glyph height) instead of ...Pt.
       customColor is { rgb, cmyk }, picked by the document color mode */
    var BUILTIN_PRESETS = [
        {
            name: { ja: "黄色のマーカー", en: "Yellow Marker" },
            settings: { lineStyle: "marker", colorChoice: "yellow", strokeWidthRatio: 0.4, verticalOffsetPt: 0, dashStyle: "solidLine", strokeCap: "buttCap", whiteText: false }
        },
        {
            name: { ja: "斜めの蛍光ペン", en: "Slanted Highlighter" },
            settings: { lineStyle: "straightShear", shearAngle: 30, colorChoice: "yellow", strokeWidthRatio: 0.5, verticalOffsetPt: 0, dashStyle: "solidLine", strokeCap: "buttCap", whiteText: false }
        },
        {
            name: { ja: "赤のアンダーライン", en: "Red Underline" },
            settings: { lineStyle: "underline", colorChoice: "red", strokeWidthRatio: 0.06, verticalOffsetPt: 0, dashStyle: "solidLine", strokeCap: "buttCap", whiteText: false }
        },
        {
            name: { ja: "赤の波線", en: "Red Wavy Line" },
            settings: {
                lineStyle: "zigzag", zigzagSizeRatio: 0.06, zigzagRidges: 12, colorChoice: "red",
                strokeWidthRatio: 0.04, verticalOffsetPt: 0, dashStyle: "solidLine", strokeCap: "roundCap", whiteText: false
            }
        },
        {
            name: { ja: "うすい文字色の塗りつぶし", en: "Light Text-Color Fill" },
            settings: { lineStyle: "fill", colorChoice: "textColor", textColorOpacity: 15, verticalOffsetPt: 0, dashStyle: "solidLine", strokeCap: "buttCap", whiteText: false }
        },
        {
            name: { ja: "1文字ずつの長方形（ピンク）", en: "Box per Character (Pink)" },
            settings: {
                lineStyle: "charBox", colorChoice: "custom", customColor: { rgb: [214, 96, 112], cmyk: [0, 70, 40, 10] },
                boxPaddingRatio: 0.08, cornerRadiusPt: 0, verticalOffsetPt: 0, whiteText: true
            }
        },
        {
            name: { ja: "角丸の長方形（青）", en: "Rounded Box (Blue)" },
            settings: {
                lineStyle: "box", colorChoice: "custom", customColor: { rgb: [79, 155, 216], cmyk: [65, 25, 0, 0] },
                boxPaddingRatio: 0.12, cornerRadiusRatio: 0.2, verticalOffsetPt: 0, whiteText: true
            }
        },
        {
            name: { ja: "文字色の取り消し線", en: "Text-Color Strikethrough" },
            settings: { lineStyle: "strikethrough", colorChoice: "textColor", textColorOpacity: 100, strokeWidthRatio: 0.06, verticalOffsetPt: 0, dashStyle: "solidLine", strokeCap: "buttCap", whiteText: false }
        }
    ];

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

    /**
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

    /**
     * pt の値を単位の値に換算する
     * @param {number} ptValue - pt の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} 単位の値
     */
    function ptToUnit(ptValue, unit) {
        return ptValue / unit.pointsPerUnit;
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

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

    var COLOR_SWATCH_SIZE = [30, 20];      /* 色見本の大きさ / swatch size */
    var PRESET_DROPDOWN_WIDTH = 140;       /* プリセットのドロップダウンの幅 / preset dropdown width */
    var PRESET_ICON_SIZE = [20, 18];       /* プリセットの保存・削除アイコンの大きさ / preset save and delete icon size */
    var PRESET_NAME_CHARS = 20;            /* プリセット名の入力欄の文字数 / preset name field width */

    /**
     * 複数のラベル（statictext）の幅を最長のものへ揃え、指定方向に揃えます。
     * 実際の描画幅（preferredSize.width）を測るため、文言長やロケールに依存しません。
     * @param {StaticText[]} labelControls - 幅を揃える statictext の配列。
     * @param {string} [justify] - 揃え方向 "left" | "center" | "right"（省略時は "right"）。
     * @returns {number} 揃えた後の共通ラベル幅（px）。
     */
    function alignLabelWidths(labelControls, justify) {
        if (!labelControls || !labelControls.length) return 0;
        justify = justify || 'right';

        /* 1パス目：各ラベルの自然幅を測り、最大値を求める / Pass 1: find the widest natural width */
        var maxLabelWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            /* preferredSize を返さない環境では実サイズで代用する / fall back to the laid-out size */
            var naturalWidth = labelControls[i].preferredSize.width || labelControls[i].size.width;
            if (naturalWidth > maxLabelWidth) maxLabelWidth = naturalWidth;
        }

        /* 2パス目：全ラベルを最大幅に固定し、揃え方向を適用 / Pass 2: apply common width and justification */
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxLabelWidth;
            labelControls[j].justify = justify;
        }

        return maxLabelWidth;
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
    };

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
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

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
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
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
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
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
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
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
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
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
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
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
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

    // アイコンのボタン（再利用パーツ） / Icon buttons (reusable)

    var ICON_BUTTON_SIZE = [24, 22]; /* 既定の大きさ / default size */
    var ICON_BUTTON_UI_DARK = isDarkUI();
    var ICON_BUTTON_COLOR     = ICON_BUTTON_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* 絵の色 / icon color */
    var ICON_BUTTON_DIM_COLOR = ICON_BUTTON_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の色 / color when disabled */

    /**
     * onDraw で絵を描くアイコンのボタンを追加する。押すと onClick を呼ぶ（無効の間は押せず、薄く描く）
     * @param {Group} parent - 追加先
     * @param {number[]} iconSize - [幅, 高さ]
     * @param {Function} drawIcon - 絵を描く関数 (iconGraphics, iconWidth, iconHeight, iconColor)
     * @returns {Group} アイコン
     */
    function addIconButton(parent, iconSize, drawIcon) {
        var iconButton = parent.add("group");
        iconButton.preferredSize = iconSize;
        iconButton.minimumSize = iconSize;
        iconButton.maximumSize = iconSize;
        iconButton.onDraw = function () {
            var iconColor = isIconButtonEnabledInTree(iconButton) ? ICON_BUTTON_COLOR : ICON_BUTTON_DIM_COLOR;
            drawIcon(iconButton.graphics, iconSize[0], iconSize[1], iconColor);
        };
        iconButton.addEventListener("mousedown", function () {
            if (!isIconButtonEnabledInTree(iconButton)) return;
            if (typeof iconButton.onClick === "function") iconButton.onClick();
        });
        return iconButton;
    }

    /**
     * アイコンのボタンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} iconButton - addIconButton() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setIconButtonEnabled(iconButton, isEnabled) {
        if (iconButton.enabled === isEnabled) return;
        iconButton.enabled = isEnabled;
        /* group には notify() が無いため、隠して再表示して描き直させる / groups have no notify(), so hide and show to repaint */
        iconButton.hide();
        iconButton.show();
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isIconButtonEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * 絵の座標の長方形の並びを、縦横比を保って中央に置いて塗る（ScriptUI は多角形を塗れないため、斜めも長方形の並びで描く）
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} designSize - 絵の [幅, 高さ]
     * @param {number[][]} designRects - 長方形 [左, 上, 右, 下] の並び
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function fillIconButtonRects(iconGraphics, iconWidth, iconHeight, designSize, designRects, iconColor) {
        var iconScale = Math.min(iconWidth / designSize[0], iconHeight / designSize[1]);
        var originX = (iconWidth - designSize[0] * iconScale) / 2;
        var originY = (iconHeight - designSize[1] * iconScale) / 2;
        iconGraphics.newPath();
        for (var i = 0; i < designRects.length; i++) {
            var designRect = designRects[i];
            if (designRect[2] <= designRect[0] || designRect[3] <= designRect[1]) continue;
            iconGraphics.rectPath(originX + designRect[0] * iconScale, originY + designRect[1] * iconScale,
                (designRect[2] - designRect[0]) * iconScale, (designRect[3] - designRect[1]) * iconScale);
        }
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor));
    }

    /**
     * 保存アイコン（トレイに下向きの矢印）を描く。1223×993 の絵。
     * 矢じりとトレイの V 字の切り欠きは細い長方形を並べ、トレイの四角い穴は塗らずに残す
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawSaveIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var designWidth = 1223;
        var trayHole = [78, 688, 230, 840]; /* トレイの四角い穴 [左, 上, 右, 下] / the tray's square hole */
        var sliceCount = 16;
        var designRects = [[535, 0, 688, 383]]; /* 矢印の軸 / arrow shaft */
        var k;

        /** 穴と重なる部分を除いて長方形を足す / add a rectangle minus the tray hole */
        function addTrayRect(left, top, right, bottom) {
            if (right <= trayHole[0] || left >= trayHole[2] || bottom <= trayHole[1] || top >= trayHole[3]) {
                designRects.push([left, top, right, bottom]);
                return;
            }
            designRects.push([left, top, right, trayHole[1]], [left, trayHole[3], right, bottom],
                [left, Math.max(top, trayHole[1]), trayHole[0], Math.min(bottom, trayHole[3])],
                [trayHole[2], Math.max(top, trayHole[1]), right, Math.min(bottom, trayHole[3])]);
        }

        /* 下向きの矢じり（上の付け根から先端へ細っていく）/ the downward head */
        for (k = 0; k < sliceCount; k++) {
            var headHalf = 251 * (1 - (k + 0.5) / sliceCount);
            var headTop = 383 + (703 - 383) * k / sliceCount;
            designRects.push([612 - headHalf, headTop, 612 + headHalf, headTop + (703 - 383) / sliceCount]);
        }
        /* トレイ：上辺に V 字の切り欠き（幅 458 から下へ細り、856 で閉じる）/ tray with a V notch along the top */
        for (k = 0; k < sliceCount; k++) {
            var notchTop = 612 + (856 - 612) * k / sliceCount;
            var notchBottom = notchTop + (856 - 612) / sliceCount;
            var notchHalf = 229 * (1 - (k + 0.5) / sliceCount);
            addTrayRect(0, notchTop, 612 - notchHalf, notchBottom);
            addTrayRect(612 + notchHalf, notchTop, designWidth, notchBottom);
        }
        addTrayRect(0, 856, designWidth, 993);
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [designWidth, 993], designRects, iconColor);
    }

    /**
     * 削除アイコン（ゴミ箱）を描く。756×825 の絵
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawTrashIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [756, 825], [
            [206, 0, 550, 70], [206, 70, 275, 137], [481, 70, 550, 137],     /* 取っ手 / handle */
            [0, 137, 756, 207],                                               /* ふた / lid */
            [69, 207, 138, 825], [618, 207, 688, 825], [138, 756, 618, 825], /* 本体 / body */
            [206, 275, 275, 687], [343, 275, 413, 687], [481, 275, 550, 687] /* 縦の線 / ribs */
        ], iconColor);
    }

    // アイコンのボタン（再利用パーツ）ここまで / End of the reusable icon buttons

    // =========================================
    // プリセットの保存 / Preset storage
    // =========================================

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

    /* プリセットは名前をキーにした集まり（既定値 {} で中身を問わず受け取る）。再起動しても残す
       Presets: a map keyed by name (the {} default accepts any content), kept across restarts */
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "Presets", "persistent");

    // =========================================
    // ローカライズ / Localization
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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    var LABELS = {
        dialog: {
            title:      { ja: "マーカーを引く", en: "Draw Marker" },
            presetSave: { ja: "プリセットを保存", en: "Save Preset" }
        },
        panel: {
            lineStyle: { ja: "種類", en: "Style" },
            settings:  { ja: "設定", en: "Settings" }
        },
        dropdown: {
            presetPlaceholder: { ja: "---", en: "---" }
        },
        fieldLabel: {
            preset:       { ja: "プリセット", en: "Preset" },
            presetName:   { ja: "プリセット名", en: "Preset name" },
            shearAngle:   { ja: "角度", en: "Angle" },
            zigzagSize:   { ja: "大きさ", en: "Size" },
            zigzagRidges: { ja: "折り返し", en: "Ridges" },
            boxPadding:   { ja: "余白", en: "Padding" },
            cornerRadius: { ja: "角丸の半径", en: "Corner Radius" },
            opacity:     { ja: "不透明度", en: "Opacity" },
            verticalOffset: { ja: "上下位置", en: "Vertical Offset" },
            color:       { ja: "カラー", en: "Color" },
            strokeWidth: { ja: "線幅", en: "Weight" },
            dashStyle:   { ja: "線", en: "Line" },
            strokeCap:   { ja: "線端", en: "Cap" }
        },
        checkbox: {
            preview:   { ja: "プレビュー", en: "Preview" },
            whiteText: { ja: "文字を白抜きに", en: "White Text" }
        },
        radio: {
            underline: { ja: "アンダーライン", en: "Underline" },
            marker:    { ja: "マーカー", en: "Marker" },
            straightShear: { ja: "直線＋シアー", en: "Straight + Shear" },
            zigzag:    { ja: "ギザギザ", en: "Zigzag" },
            fill:      { ja: "塗りつぶし", en: "Fill" },
            box:       { ja: "長方形", en: "Box" },
            charBox:   { ja: "1文字ずつの長方形", en: "Box per Character" },
            strikethrough: { ja: "取り消し線", en: "Strikethrough" },
            yellow:   { ja: "黄色", en: "Yellow" },
            red:       { ja: "赤", en: "Red" },
            textColor: { ja: "文字色", en: "Text Color" },
            custom:    { ja: "カスタム", en: "Custom" },
            buttCap:  { ja: "なし", en: "Butt" },
            roundCap: { ja: "丸型", en: "Round" },
            solidLine:  { ja: "直線", en: "Solid" },
            dashedLine: { ja: "破線", en: "Dashed" },
            dottedLine: { ja: "ドット点線", en: "Dotted" }
        },
        tooltip: {
            swatch:    { ja: "クリックしてカラーピッカーで変更", en: "Click to change in the Color Picker" },
            textColor: {
                ja: "対象の先頭の文字の塗りを使います（テキストオブジェクトごと）",
                en: "Uses the fill of the first target character (per text object)"
            },
            opacity:   { ja: "［文字色］の線の不透明度です", en: "Opacity of the line in Text Color" },
            verticalOffset: {
                ja: "種類ごとの位置から線を上下にずらす量です。プラスで上、マイナスで下へずれます",
                en: "Moves the line up or down from the style's position; positive moves it up, negative down"
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            roundCap:  { ja: "線の両端が線幅の半分ずつ外へ伸びます", en: "Each end extends by half the line width" },
            dashedLine: { ja: "線幅の2倍の線分を、線幅の間隔で並べます", en: "Dashes twice the line width, spaced by the line width" },
            dottedLine: {
                ja: "線幅の大きさの丸い点を、線幅の2倍の間隔で並べます（線端は丸型になります）",
                en: "Round dots as large as the line width, every two line widths (the cap becomes round)"
            },
            whiteText: {
                ja: "線や長方形を描いたあと、対象の文字の塗りを白にします",
                en: "Sets the fill of the target characters to white after drawing"
            },
            preview:   { ja: "プレビュー中は文字の選択を外し、閉じると元に戻します", en: "Deselects the text while previewing and restores it when the dialog box closes" },
            underline: { ja: "字面のすぐ下に、線の上端を合わせて引きます", en: "Draws the line just below the glyphs, its top edge on their bottom" },
            marker:    { ja: "字面の下端を線の中心にして引きます", en: "Draws the line centered on the bottom edge of the glyphs" },
            straightShear: { ja: "線に［パスの自由変形］効果を掛けて、両端を斜めにします", en: "Applies the Free Distort effect to slant both ends of the line" },
            shearAngle: {
                ja: "線の両端を傾ける角度です（-80〜80°）。プラスで右上がり、マイナスで左上がりになります",
                en: "Angle of the slanted ends (-80 to 80°); positive slants to the upper right, negative to the upper left"
            },
            zigzag:    { ja: "線に［ジグザグ］効果を掛けます", en: "Applies the Zig Zag effect to the line" },
            fill:      { ja: "字面の範囲を塗りつぶします（線幅は字面の高さ）", en: "Fills the glyph area (the line width is the glyph height)" },
            charBox: {
                ja: "1文字ずつ、字面を囲む長方形を描きます（高さは行の字面にそろえ、幅は高さ以上）",
                en: "Draws a box around each character (as tall as the line's glyphs, at least as wide as tall)"
            },
            box: {
                ja: "字面を囲む長方形を、行ごとに1つ描きます",
                en: "Draws one box around the glyphs of each line"
            },
            boxPadding:   { ja: "［長方形］［1文字ずつの長方形］の、字面のまわりの余白です", en: "Space around the glyphs in Box and Box per Character" },
            cornerRadius: {
                ja: "［長方形］［1文字ずつの長方形］の角丸の半径です。0 で角丸なし",
                en: "Corner radius of Box and Box per Character; 0 for square corners"
            },
            strikethrough: { ja: "字面の上下中央に線を引きます", en: "Draws the line through the vertical center of the glyphs" },
            preset:       { ja: "保存した設定を読み込みます", en: "Loads a saved set of settings" },
            presetSave:   { ja: "今の設定に名前を付けて保存します", en: "Saves the current settings under a name" },
            presetDelete: {
                ja: "選んでいるプリセットを削除します（最初からあるものは削除できません）",
                en: "Deletes the selected preset (built-in presets cannot be deleted)"
            },
            presetName:   { ja: "保存する設定の名前です。同じ名前は上書きします", en: "Name the settings are saved under. The same name is overwritten" },
            zigzagSize:   { ja: "ギザギザの山の高さです", en: "Height of the zigzag ridges" },
            zigzagRidges: { ja: "線の端から端までの折り返しの数です", en: "Number of ridges from one end of the line to the other" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextSelection: {
                ja: "文字ツールで文字を選択するか、テキストオブジェクトを選択してください。",
                en: "Select characters with the Type tool, or select text objects."
            },
            noGlyphs: {
                ja: "選択した文字に字形がありません。\nスペースや改行だけの範囲にはマーカーを引けません。",
                en: "The selected characters have no glyphs.\nA marker cannot be drawn over spaces or line breaks only."
            },
            presetSaveFailed: { ja: "プリセットを保存できませんでした。", en: "Could not save the preset." }
        },
        confirm: {
            presetOverwrite: { ja: "「{name}」を上書きしますか？", en: "Overwrite \"{name}\"?" },
            presetDelete:    { ja: "「{name}」を削除しますか？", en: "Delete \"{name}\"?" }
        }
    };

    // =========================================
    // 色 / Colors
    // =========================================

    /**
     * ドキュメントのカラーモードに合わせて色を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} rgbValues - R, G, B（0〜255）
     * @param {number[]} cmykValues - C, M, Y, K（0〜100）
     * @returns {Color} RGBColor または CMYKColor
     */
    function createDocumentColor(doc, rgbValues, cmykValues) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = cmykValues[0];
            cmykColor.magenta = cmykValues[1];
            cmykColor.yellow = cmykValues[2];
            cmykColor.black = cmykValues[3];
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = rgbValues[0];
        rgbColor.green = rgbValues[1];
        rgbColor.blue = rgbValues[2];
        return rgbColor;
    }

    /**
     * 色を色見本に描くための RGB（0〜1）にする（CMYK・グレーは簡易換算）
     * @param {Color} swatchColor - 色
     * @returns {number[]} [r, g, b]（0〜1）
     */
    function colorToScreenRgb(swatchColor) {
        if (swatchColor.typename === "RGBColor") return [swatchColor.red / 255, swatchColor.green / 255, swatchColor.blue / 255];
        if (swatchColor.typename === "GrayColor") return [1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100, 1 - swatchColor.gray / 100];
        if (swatchColor.typename === "CMYKColor") {
            var blackRatio = 1 - swatchColor.black / 100;
            return [
                (1 - swatchColor.cyan / 100) * blackRatio,
                (1 - swatchColor.magenta / 100) * blackRatio,
                (1 - swatchColor.yellow / 100) * blackRatio
            ];
        }
        /* スポットカラーなどは灰色で示す / show other color types in gray */
        return [0.5, 0.5, 0.5];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * クリックで標準のカラーピッカーを開く色見本を作る
     * @param {Group} parent - 追加先
     * @param {Color} initialColor - 最初の色
     * @param {Function} onColorPicked - 色を選んだあとに、選んだ色を渡して呼ぶ関数
     * @returns {Group} 色見本
     */
    function addColorSwatch(parent, initialColor, onColorPicked) {
        var swatchColor = initialColor;
        var colorSwatch = parent.add("group");
        colorSwatch.preferredSize = COLOR_SWATCH_SIZE;
        colorSwatch.helpTip = getLabel("tooltip.swatch");
        colorSwatch.onDraw = function () {
            var swatchGraphics = colorSwatch.graphics;
            var swatchWidth = colorSwatch.size[0];
            var swatchHeight = colorSwatch.size[1];
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, colorToScreenRgb(swatchColor)));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 1));
        };
        /* キャンセルは渡した色が返る。group には notify() が無いので、隠して再表示して描き直す
           Cancel returns the color passed in; a group has no notify(), so hide and show it to redraw */
        colorSwatch.addEventListener("click", function () {
            swatchColor = app.showColorPicker(swatchColor);
            colorSwatch.hide();
            colorSwatch.show();
            onColorPicked(swatchColor);
        });
        /* プリセットから色を差し替える / Replace the color from a preset */
        colorSwatch.setSwatchColor = function (newColor) {
            swatchColor = newColor;
            colorSwatch.hide();
            colorSwatch.show();
        };
        return colorSwatch;
    }

    /**
     * 項目名と［黄色］［赤］［文字色］［カスタム］のラジオボタンを縦に並べた、カラーを選ぶ欄を作る。
     * ［カスタム］の横には、クリックでカラーピッカーを開く色見本を置く
     * @param {Window|Panel|Group} parent - 追加先
     * @param {Document} doc - 対象ドキュメント（カラーモード）
     * @param {boolean} isTextColorUsable - ［文字色］を使えるか（false なら無効にする）
     * @param {Function} onColorChange - カラーを選び直したあとに呼ぶ関数
     * @returns {{label: StaticText, getColorOptions: Function, getChoice: Function, setChoice: Function, getCustomColor: Function, setCustomColor: Function}}
     *     項目名、{ color, useTextColor, opacity } を返す getColorOptions、選んだキーの読み書き、［カスタム］の色の読み書き、
     *     ［文字色］の不透明度の入力欄 opacityInput
     */
    function addColorChoiceRow(parent, doc, isTextColorUsable, onColorChange) {
        var colorRow = parent.add("group");
        setupRow(colorRow, "left", 8);
        colorRow.alignChildren = ["left", "top"];
        var colorLabel = colorRow.add("statictext", undefined, labelText("fieldLabel.color"));

        var choiceColumn = colorRow.add("group");
        choiceColumn.orientation = "column";
        choiceColumn.alignChildren = ["left", "center"];
        choiceColumn.spacing = 6;

        var choiceColors = {
            yellow: createDocumentColor(doc, MARKER_COLOR_PRESETS.yellow.rgb, MARKER_COLOR_PRESETS.yellow.cmyk),
            red: createDocumentColor(doc, MARKER_COLOR_PRESETS.red.rgb, MARKER_COLOR_PRESETS.red.cmyk),
            custom: createDocumentColor(doc, MARKER_COLOR_PRESETS.custom.rgb, MARKER_COLOR_PRESETS.custom.cmyk)
        };
        /* show() 前に代入した value は読み戻せないので、選択は変数で持つ / value set before show() cannot be read back, so keep the choice in a variable */
        var chosenKey = (DEFAULT_COLOR_CHOICE === "textColor" && !isTextColorUsable) ? "yellow" : DEFAULT_COLOR_CHOICE;
        var choiceRadios = {};

        /* ［カスタム］は色見本と同じ行で親が別なので、排他は手で取る / Custom sits in another group with its swatch, so exclusivity is handled here */
        function checkChoiceRadio(choiceKey) {
            chosenKey = choiceKey;
            for (var radioKey in choiceRadios) {
                if (choiceRadios.hasOwnProperty(radioKey)) choiceRadios[radioKey].value = (radioKey === choiceKey);
            }
            /* ［不透明度］は［文字色］のときだけ使える / Opacity is used only for Text Color */
            if (opacityInput) setSteppedFieldEnabled(opacityInput, choiceKey === "textColor");
        }

        function selectChoice(choiceKey) {
            checkChoiceRadio(choiceKey);
            onColorChange();
        }

        function addChoiceRadio(radioParent, choiceKey) {
            var choiceRadio = radioParent.add("radiobutton", undefined, getLabel("radio." + choiceKey));
            choiceRadio.value = (choiceKey === chosenKey);
            choiceRadio.onClick = function () { selectChoice(choiceKey); };
            choiceRadios[choiceKey] = choiceRadio;
            return choiceRadio;
        }

        addChoiceRadio(choiceColumn, "yellow");
        addChoiceRadio(choiceColumn, "red");
        var opacityInput = null;
        var textColorRow = choiceColumn.add("group");
        setupRow(textColorRow, "left", 8);
        var textColorRadio = addChoiceRadio(textColorRow, "textColor");
        textColorRadio.helpTip = getLabel("tooltip.textColor");
        textColorRadio.enabled = isTextColorUsable;
        opacityInput = addMarkerNumberField(textColorRow, "opacity", {
            text: TEXT_COLOR_OPACITY + "%", characters: 5, step: 5, min: 0, max: 100, unit: "%"
        }, onColorChange);
        setSteppedFieldEnabled(opacityInput, chosenKey === "textColor");

        var customRow = choiceColumn.add("group");
        setupRow(customRow, "left", 8);
        addChoiceRadio(customRow, "custom");
        /* 色見本で色を変えたら［カスタム］を選んだことにする / Changing the swatch also chooses Custom */
        var customSwatch = addColorSwatch(customRow, choiceColors.custom, function (pickedColor) {
            choiceColors.custom = pickedColor;
            selectChoice("custom");
        });

        return {
            label: colorLabel,
            /* ［文字色］の色は対象ごとに違うので、線を引くときに対象から取る / Text Color differs per target, so it is read from each target when drawing */
            getColorOptions: function () {
                if (chosenKey === "textColor") return { color: null, useTextColor: true, opacity: parseFloat(opacityInput.text) };
                return { color: choiceColors[chosenKey], useTextColor: false, opacity: 100 };
            },
            getChoice: function () { return chosenKey; },
            /* 知らないキーと、使えない［文字色］は［黄色］にする / unknown keys and an unusable Text Color fall back to Yellow */
            setChoice: function (choiceKey) {
                var isUsable = choiceRadios.hasOwnProperty(choiceKey) && (choiceKey !== "textColor" || isTextColorUsable);
                checkChoiceRadio(isUsable ? choiceKey : "yellow");
            },
            opacityInput: opacityInput,
            getCustomColor: function () { return choiceColors.custom; },
            setCustomColor: function (newColor) {
                choiceColors.custom = newColor;
                customSwatch.setSwatchColor(newColor);
            }
        };
    }

    /**
     * 項目名・∧∨・入力欄の、数値を入れる1行を作る
     * @param {Window|Panel|Group} parent - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー（LABELS.tooltip に同じキーがあれば入力欄の helpTip にする）
     * @param {Object} fieldOptions - addSteppedField() の text / characters / step / min / max / integer / unit
     * @param {Function} onValueChange - 値を変えたあとに呼ぶ関数
     * @returns {EditText} 入力欄（項目名は .fieldLabel）
     */
    function addMarkerNumberField(parent, labelKey, fieldOptions, onValueChange) {
        fieldOptions.label = labelText("fieldLabel." + labelKey);
        fieldOptions.onStep = function () { onValueChange(); };
        var numberInput = addSteppedField(parent, fieldOptions);
        if (LABELS.tooltip[labelKey]) numberInput.helpTip = getLabel("tooltip." + labelKey);
        /* 部品の onChange が値を整えたあとでプレビューする（change リスナーは整える前に呼ばれる）
           Refresh after the part's onChange has normalized the value (a change listener runs before it) */
        var normalizeValue = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeValue();
            onValueChange();
        };
        return numberInput;
    }

    /**
     * 項目名とラジオボタンの1行を作る（ラジオは同じグループに入れて排他にする）
     * @param {Window|Panel|Group} parent - 追加先
     * @param {string|null} labelKey - LABELS.fieldLabel のキー（null なら項目名を置かない）
     * @param {string[]} choiceKeys - LABELS.radio のキー（並べる順）
     * @param {string} defaultKey - 最初に選んでおくキー
     * @param {Function} onChoiceChange - 選び直したあとに呼ぶ関数
     * @param {boolean} [isVertical] - true ならラジオを縦に並べる
     * @returns {{row: Group, label: StaticText, radios: Object, getChoice: Function, setChoice: Function}} 行のグループ、項目名、キーごとのラジオボタン、選んだキーの読み書き
     */
    function addRadioChoiceRow(parent, labelKey, choiceKeys, defaultKey, onChoiceChange, isVertical) {
        var choiceRow = parent.add("group");
        setupRow(choiceRow, "left", 8);
        var choiceLabel = labelKey ? choiceRow.add("statictext", undefined, labelText("fieldLabel." + labelKey)) : null;
        var radioParent = choiceRow;
        if (isVertical) {
            choiceRow.alignChildren = ["left", "top"];
            radioParent = choiceRow.add("group");
            radioParent.orientation = "column";
            radioParent.alignChildren = ["left", "center"];
            radioParent.spacing = 6;
        }
        /* show() 前に代入した value は読み戻せないので、選択は変数で持つ / value set before show() cannot be read back, so keep the choice in a variable */
        var chosenKey = defaultKey;
        var choiceRadios = {};
        for (var i = 0; i < choiceKeys.length; i++) {
            (function (choiceKey) {
                var choiceRadio = radioParent.add("radiobutton", undefined, getLabel("radio." + choiceKey));
                choiceRadio.value = (choiceKey === chosenKey);
                choiceRadio.onClick = function () {
                    chosenKey = choiceKey;
                    onChoiceChange();
                };
                choiceRadios[choiceKey] = choiceRadio;
            })(choiceKeys[i]);
        }
        return {
            row: choiceRow,
            label: choiceLabel,
            radios: choiceRadios,
            getChoice: function () { return chosenKey; },
            /* 知らないキーは無視する / ignore unknown keys */
            setChoice: function (choiceKey) {
                if (!choiceRadios.hasOwnProperty(choiceKey)) return;
                chosenKey = choiceKey;
                for (var radioKey in choiceRadios) {
                    if (choiceRadios.hasOwnProperty(radioKey)) choiceRadios[radioKey].value = (radioKey === choiceKey);
                }
            }
        };
    }

    /**
     * プレビューの線を出し入れするオブジェクトを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} markerTargets - collectMarkerTargets() の戻り値
     * @returns {{update: Function, clear: Function}} 線を引き直す update(markerOptions) と、消す clear()
     */
    function createMarkerPreview(doc, markerTargets) {
        var previewLines = [];
        /* ［文字を白抜きに］は、白くした複製を前面に置いて元のテキストを隠して見せる
           White Text is previewed with a whitened copy in front while the original is hidden */
        var whiteTextCopies = [];

        function clearPreviewLines() {
            for (var i = 0; i < previewLines.length; i++) previewLines[i].remove();
            previewLines = [];
            for (var j = 0; j < whiteTextCopies.length; j++) {
                whiteTextCopies[j].remove();
                markerTargets[j].sourceFrame.hidden = false;
            }
            whiteTextCopies = [];
        }

        return {
            update: function (markerOptions) {
                clearPreviewLines();
                /* 選択の反転表示やバウンディングボックスで線が見えにくいので、選択を外す（閉じたあと main() で戻す）
                   The selection highlight hides the line, so deselect it (main() restores it after closing) */
                doc.selection = null;
                previewLines = placeMarkerLines(markerTargets, markerOptions);
                if (!markerOptions.whiteText) return;
                var whiteColor = createDocumentColor(doc, WHITE_TEXT_COLOR.rgb, WHITE_TEXT_COLOR.cmyk);
                for (var i = 0; i < markerTargets.length; i++) {
                    var sourceFrame = markerTargets[i].sourceFrame;
                    /* 複製は hidden を引き継ぐので、隠す前に作る / a duplicate inherits hidden, so copy before hiding */
                    var textCopy = sourceFrame.duplicate(sourceFrame, ElementPlacement.PLACEBEFORE);
                    fillTargetSpan(textCopy, markerTargets[i], whiteColor);
                    sourceFrame.hidden = true;
                    whiteTextCopies.push(textCopy);
                }
            },
            clear: clearPreviewLines
        };
    }

    /**
     * 項目名・ドロップダウン・保存と削除のアイコンの、プリセットの1行を作る
     * @param {Window|Panel|Group} parent - 追加先
     * @returns {{label: StaticText, presetDropdown: DropDownList, presetSaveIcon: Group, presetDeleteIcon: Group}} 項目名とコントロール
     */
    function addPresetRow(parent) {
        var presetRow = parent.add("group");
        setupRow(presetRow, "center", 8);
        var presetLabel = presetRow.add("statictext", undefined, labelText("fieldLabel.preset"));
        var presetDropdown = presetRow.add("dropdownlist", undefined, []);
        presetDropdown.helpTip = getLabel("tooltip.preset");
        presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
        var presetSaveIcon = addIconButton(presetRow, PRESET_ICON_SIZE, drawSaveIcon);
        presetSaveIcon.helpTip = getLabel("tooltip.presetSave");
        var presetDeleteIcon = addIconButton(presetRow, PRESET_ICON_SIZE, drawTrashIcon);
        presetDeleteIcon.helpTip = getLabel("tooltip.presetDelete");
        return {
            label: presetLabel,
            presetDropdown: presetDropdown,
            presetSaveIcon: presetSaveIcon,
            presetDeleteIcon: presetDeleteIcon
        };
    }

    /**
     * 左に［種類］、右に［設定］（カラー・線幅・線端など）のパネルを並べてコントロールを作る
     * @param {Window} markerDialog - 追加先のダイアログボックス
     * @param {Document} doc - 対象ドキュメント
     * @param {number} glyphHeightPt - 線幅と余白の初期値のもとにする字面の高さ（pt）
     * @param {boolean} isTextColorUsable - ［文字色］を使えるか
     * @param {Function} onSettingChange - 設定を変えたあとに呼ぶ関数
     * @returns {Object} コントロールと単位 { lineStyleRow, shearAngleInput, zigzagSizeInput, zigzagRidgesInput, boxPaddingInput, cornerRadiusInput, colorRow, strokeWidthInput, verticalOffsetInput, dashStyleRow, strokeCapRow, getWhiteText, setWhiteText, rulerUnit, strokeUnit, labelColumns }
     *     （labelColumns は［設定］パネルの列ごとの項目名）
     */
    function addMarkerControls(markerDialog, doc, glyphHeightPt, isTextColorUsable, onSettingChange) {
        var rulerUnit = getUnitInfo();
        var strokeUnit = getUnitInfo("strokeUnits");
        var markerControls = { rulerUnit: rulerUnit, strokeUnit: strokeUnit };

        var columnGroup = markerDialog.add("group");
        columnGroup.orientation = "row";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = COLUMN_SPACING;
        var lineStylePanel = columnGroup.add("panel", undefined, getLabel("panel.lineStyle"));
        setupPanel(lineStylePanel);
        var settingsPanel = columnGroup.add("panel", undefined, getLabel("panel.settings"));
        setupPanel(settingsPanel);

        var lineStyleKeys = ["underline", "marker", "straightShear", "zigzag", "fill", "box", "charBox", "strikethrough"];
        /* 項目名はパネルの見出しが兼ねる / the panel title serves as the label */
        markerControls.lineStyleRow = addRadioChoiceRow(lineStylePanel, null, lineStyleKeys, DEFAULT_LINE_STYLE, function () {
            updateStyleFieldsEnabled(markerControls);
            onSettingChange();
        }, true);
        markerControls.lineStyleRow.row.alignment = ["left", "top"];
        for (var i = 0; i < lineStyleKeys.length; i++) {
            markerControls.lineStyleRow.radios[lineStyleKeys[i]].helpTip = getLabel("tooltip." + lineStyleKeys[i]);
        }
        /* ［設定］は左に共通の項目、右に種類ごとの項目の2列 / Settings: common items on the left, style-specific items on the right */
        settingsPanel.orientation = "row";
        settingsPanel.alignChildren = ["fill", "top"];
        settingsPanel.spacing = COLUMN_SPACING;
        var commonColumn = addSettingsColumn(settingsPanel);
        var styleColumn = addSettingsColumn(settingsPanel);

        markerControls.colorRow = addColorChoiceRow(commonColumn, doc, isTextColorUsable, onSettingChange);
        markerControls.strokeWidthInput = addMarkerNumberField(commonColumn, "strokeWidth", {
            text: formatStepperNumber(ptToUnit(glyphHeightPt * MARKER_WEIGHT_RATIO, strokeUnit)) + " " + strokeUnit.label,
            characters: 7, step: 1, min: 0.1, unit: " " + strokeUnit.label
        }, onSettingChange);
        markerControls.verticalOffsetInput = addMarkerNumberField(commonColumn, "verticalOffset", {
            text: "0 " + rulerUnit.label, characters: 7, step: 1, unit: " " + rulerUnit.label
        }, onSettingChange);
        markerControls.dashStyleRow = addRadioChoiceRow(commonColumn, "dashStyle", ["solidLine", "dashedLine", "dottedLine"], DEFAULT_DASH_STYLE, function () {
            updateStyleFieldsEnabled(markerControls);
            onSettingChange();
        });
        markerControls.dashStyleRow.radios.dashedLine.helpTip = getLabel("tooltip.dashedLine");
        markerControls.dashStyleRow.radios.dottedLine.helpTip = getLabel("tooltip.dottedLine");
        markerControls.strokeCapRow = addRadioChoiceRow(commonColumn, "strokeCap", ["buttCap", "roundCap"], DEFAULT_STROKE_CAP, onSettingChange);
        markerControls.strokeCapRow.radios.roundCap.helpTip = getLabel("tooltip.roundCap");


        /* show() 前に代入した value は読み戻せないので、オンオフは変数で持つ / value set before show() cannot be read back, so keep the state in a variable */
        var isWhiteText = DEFAULT_WHITE_TEXT;
        var whiteTextCheckbox = commonColumn.add("checkbox", undefined, getLabel("checkbox.whiteText"));
        whiteTextCheckbox.alignment = "left";
        whiteTextCheckbox.helpTip = getLabel("tooltip.whiteText");
        whiteTextCheckbox.value = isWhiteText;
        whiteTextCheckbox.onClick = function () {
            isWhiteText = whiteTextCheckbox.value;
            onSettingChange();
        };
        markerControls.getWhiteText = function () { return isWhiteText; };
        markerControls.setWhiteText = function (shouldWhiten) {
            isWhiteText = shouldWhiten;
            whiteTextCheckbox.value = shouldWhiten;
        };

        markerControls.shearAngleInput = addMarkerNumberField(styleColumn, "shearAngle", {
            text: DEFAULT_SHEAR_ANGLE + "°", characters: 7, step: 1, min: -80, max: 80, unit: "°"
        }, onSettingChange);
        markerControls.zigzagSizeInput = addMarkerNumberField(styleColumn, "zigzagSize", {
            text: formatStepperNumber(ptToUnit(DEFAULT_ZIGZAG_SIZE, rulerUnit)) + " " + rulerUnit.label,
            characters: 7, step: 1, min: 0.1, unit: " " + rulerUnit.label
        }, onSettingChange);
        markerControls.zigzagRidgesInput = addMarkerNumberField(styleColumn, "zigzagRidges", {
            text: String(DEFAULT_ZIGZAG_RIDGES), characters: 7, step: 1, min: 1, integer: true
        }, onSettingChange);
        markerControls.boxPaddingInput = addMarkerNumberField(styleColumn, "boxPadding", {
            text: formatStepperNumber(ptToUnit(glyphHeightPt * BOX_PADDING_RATIO, rulerUnit)) + " " + rulerUnit.label,
            characters: 7, step: 1, min: 0, unit: " " + rulerUnit.label
        }, onSettingChange);
        markerControls.cornerRadiusInput = addMarkerNumberField(styleColumn, "cornerRadius", {
            text: formatStepperNumber(ptToUnit(DEFAULT_CORNER_RADIUS, rulerUnit)) + " " + rulerUnit.label,
            characters: 7, step: 1, min: 0, unit: " " + rulerUnit.label
        }, onSettingChange);

        /* 項目名の幅は列ごとにそろえる / Label widths are aligned per column */
        markerControls.labelColumns = [
            [
                markerControls.colorRow.label, markerControls.strokeWidthInput.fieldLabel, markerControls.verticalOffsetInput.fieldLabel,
                markerControls.dashStyleRow.label, markerControls.strokeCapRow.label
            ],
            [
                markerControls.shearAngleInput.fieldLabel, markerControls.zigzagSizeInput.fieldLabel, markerControls.zigzagRidgesInput.fieldLabel,
                markerControls.boxPaddingInput.fieldLabel, markerControls.cornerRadiusInput.fieldLabel
            ]
        ];
        updateStyleFieldsEnabled(markerControls);
        return markerControls;
    }

    /**
     * ［設定］パネルの1列を作る
     * @param {Panel} settingsPanel - 追加先の［設定］パネル
     * @returns {Group} 列のグループ
     */
    function addSettingsColumn(settingsPanel) {
        var settingsColumn = settingsPanel.add("group");
        settingsColumn.orientation = "column";
        settingsColumn.alignChildren = ["fill", "top"];
        settingsColumn.alignment = ["fill", "top"];
        settingsColumn.spacing = PANEL_SPACING;
        return settingsColumn;
    }

    /**
     * ［角度］は［直線＋シアー］、［大きさ］［折り返し］は［ギザギザ］のときだけ使えるようにする。
     * ［塗りつぶし］は線幅が字面の高さで決まるので、［線幅］を使えなくする。
     * ［長方形］［1文字ずつの長方形］は線を引かないので、［線幅］［線］［線端］のかわりに［余白］［角丸の半径］を使えるようにする。
     * ［ドット点線］のときは［線端］を使えなくする
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {void}
     */
    function updateStyleFieldsEnabled(markerControls) {
        var lineStyle = markerControls.lineStyleRow.getChoice();
        var isBoxStyle = isBoxLineStyle(lineStyle);
        setSteppedFieldEnabled(markerControls.shearAngleInput, lineStyle === "straightShear");
        setSteppedFieldEnabled(markerControls.zigzagSizeInput, lineStyle === "zigzag");
        setSteppedFieldEnabled(markerControls.zigzagRidgesInput, lineStyle === "zigzag");
        setSteppedFieldEnabled(markerControls.boxPaddingInput, isBoxStyle);
        setSteppedFieldEnabled(markerControls.cornerRadiusInput, isBoxStyle);
        setSteppedFieldEnabled(markerControls.strokeWidthInput, lineStyle !== "fill" && !isBoxStyle);
        setRadiosEnabled(markerControls.dashStyleRow, !isBoxStyle);
        /* ドット点線は丸型の線端で点にするので、［線端］は選べない / Dotted needs round caps for its dots, so Cap is not selectable */
        setRadiosEnabled(markerControls.strokeCapRow, !isBoxStyle && markerControls.dashStyleRow.getChoice() !== "dottedLine");
    }

    /**
     * addRadioChoiceRow() の行のラジオボタンを、まとめて有効／無効にする
     * @param {Object} choiceRow - addRadioChoiceRow() の戻り値
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setRadiosEnabled(choiceRow, isEnabled) {
        for (var radioKey in choiceRow.radios) {
            if (choiceRow.radios.hasOwnProperty(radioKey)) choiceRow.radios[radioKey].enabled = isEnabled;
        }
    }

    /**
     * 線のかわりに長方形を描く種類かを返す
     * @param {string} lineStyle - 種類のキー
     * @returns {boolean} ［長方形］［1文字ずつの長方形］なら true
     */
    function isBoxLineStyle(lineStyle) {
        return lineStyle === "box" || lineStyle === "charBox";
    }

    /**
     * 今のダイアログボックスの指定を、線を引くための値にまとめる
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {Object} { color, useTextColor, opacity, strokeWidth（pt）, verticalOffset（pt）, lineStyle, shearAngle（度）, zigzagSize（pt）, zigzagRidges, boxPadding（pt）, cornerRadius（pt）, dashStyle, strokeCap, whiteText }
     */
    function readMarkerOptions(markerControls) {
        var colorOptions = markerControls.colorRow.getColorOptions();
        return {
            color: colorOptions.color,
            useTextColor: colorOptions.useTextColor,
            opacity: colorOptions.opacity,
            strokeWidth: unitToPt(parseFloat(markerControls.strokeWidthInput.text), markerControls.strokeUnit),
            verticalOffset: unitToPt(parseFloat(markerControls.verticalOffsetInput.text), markerControls.rulerUnit),
            lineStyle: markerControls.lineStyleRow.getChoice(),
            shearAngle: parseFloat(markerControls.shearAngleInput.text),
            zigzagSize: unitToPt(parseFloat(markerControls.zigzagSizeInput.text), markerControls.rulerUnit),
            zigzagRidges: parseInt(markerControls.zigzagRidgesInput.text, 10),
            boxPadding: unitToPt(parseFloat(markerControls.boxPaddingInput.text), markerControls.rulerUnit),
            cornerRadius: unitToPt(parseFloat(markerControls.cornerRadiusInput.text), markerControls.rulerUnit),
            dashStyle: markerControls.dashStyleRow.getChoice(),
            strokeCap: (markerControls.strokeCapRow.getChoice() === "roundCap") ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP,
            whiteText: markerControls.getWhiteText()
        };
    }

    /**
     * マーカーのカラーと線端を選ぶダイアログボックスを表示する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} markerTargets - collectMarkerTargets() の戻り値（線幅の初期値とプレビュー用）
     * @returns {Object|null} readMarkerOptions() の値。キャンセルなら null
     */
    function showMarkerDialog(doc, markerTargets) {
        var markerDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(markerDialog);

        var markerPreview = createMarkerPreview(doc, markerTargets);
        /* show() 前に代入した value は読み戻せないので、オンオフは変数で持つ / value set before show() cannot be read back, so keep the state in a variable */
        var isPreviewOn = PREVIEW_ON_OPEN;
        /* プリセットの書き込み中は、項目ごとのプレビューを止める / suspend per-field previews while a preset is written */
        var isApplyingPreset = false;

        var presetControls = addPresetRow(markerDialog);
        /* 線幅・余白の初期値と、最初からあるプリセットの比率は、先頭の対象の字面の高さから
           Initial width and padding, and the built-in preset ratios, come from the first target's glyph height */
        var firstBounds = markerTargets[0].glyphBounds;
        var glyphHeightPt = firstBounds.top - firstBounds.bottom;
        presetControls.builtinPresets = createBuiltinPresets(doc, glyphHeightPt);
        var markerControls = addMarkerControls(markerDialog, doc, glyphHeightPt, areTextColorsUsable(markerTargets), refreshPreview);
        for (var i = 0; i < markerControls.labelColumns.length; i++) alignLabelWidths(markerControls.labelColumns[i]);

        var buttonRow = addButtonRow(markerDialog);
        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = isPreviewOn;
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.onClick = function () {
            isPreviewOn = previewCheckbox.value;
            refreshPreview();
        };
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /* プレビューの線を今の指定で引き直す（オフなら消すだけ） / Redraw the preview line with the current settings (just remove it when off) */
        function refreshPreview() {
            if (isApplyingPreset) return;
            if (isPreviewOn) markerPreview.update(readMarkerOptions(markerControls));
            else markerPreview.clear();
            app.redraw();
        }

        /* ドロップダウンを作り直して名前を選ぶ（onChange は無視させる） / Refill the dropdown, ignoring its onChange */
        function refreshPresetDropdown(selectedName) {
            isApplyingPreset = true;
            fillPresetDropdown(presetControls, selectedName);
            isApplyingPreset = false;
        }

        presetControls.presetDropdown.onChange = function () {
            if (isApplyingPreset) return;
            var presetName = getSelectedPresetName(presetControls);
            setIconButtonEnabled(presetControls.presetDeleteIcon, isSavedPreset(presetName));
            var presetData = findPresetData(presetControls, presetName);
            if (!presetData) return;
            isApplyingPreset = true;
            applyPresetData(markerControls, presetData);
            isApplyingPreset = false;
            updateStyleFieldsEnabled(markerControls);
            refreshPreview();
        };
        presetControls.presetSaveIcon.onClick = function () {
            var savedName = saveCurrentPreset(markerControls, presetControls);
            if (savedName) refreshPresetDropdown(savedName);
        };
        presetControls.presetDeleteIcon.onClick = function () {
            if (deleteSelectedPreset(presetControls)) refreshPresetDropdown(null);
        };
        refreshPresetDropdown(null);

        if (isPreviewOn) refreshPreview();
        prepareDialogWindow(markerDialog, SCRIPT_NAME);
        var dialogResult = markerDialog.show();
        /* 確定でも消して、main() で引き直す / Removed on OK too; main() draws the final line */
        markerPreview.clear();
        if (dialogResult !== 1) return null;
        return readMarkerOptions(markerControls);
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /**
     * 保存したプリセットを返す
     * @returns {Object} プリセット名をキーにした設定の集まり（書き換えても保存されない）
     */
    function loadPresetMap() {
        return presetSettingsStore.load({});
    }

    /**
     * 保存したプリセットの名前を、名前順で返す
     * @param {Object} excludedNames - 除く名前をキーにした集まり（最初からあるプリセット。同じ名前はそちらの位置に並べる）
     * @returns {string[]} プリセット名
     */
    function getSavedPresetNames(excludedNames) {
        var presetMap = loadPresetMap();
        var presetNames = [];
        for (var presetName in presetMap) {
            if (presetMap.hasOwnProperty(presetName) && !excludedNames.hasOwnProperty(presetName)) presetNames.push(presetName);
        }
        presetNames.sort();
        return presetNames;
    }

    /**
     * プリセット名が保存したもの（削除できるもの）かを返す
     * @param {string|null} presetName - プリセット名
     * @returns {boolean} 保存したプリセットなら true
     */
    function isSavedPreset(presetName) {
        return !!presetName && loadPresetMap().hasOwnProperty(presetName);
    }

    /**
     * プリセット名の設定を返す（保存したものを、同じ名前の最初からあるものより優先する）
     * @param {Object} presetControls - addPresetRow() の戻り値（builtinPresets 付き）
     * @param {string|null} presetName - プリセット名
     * @returns {Object|null} プリセット。見つからなければ null
     */
    function findPresetData(presetControls, presetName) {
        if (!presetName) return null;
        var savedMap = loadPresetMap();
        if (savedMap.hasOwnProperty(presetName)) return savedMap[presetName];
        var builtinMap = presetControls.builtinPresets.map;
        return builtinMap.hasOwnProperty(presetName) ? builtinMap[presetName] : null;
    }

    /**
     * BUILTIN_PRESETS を、保存したプリセットと同じ形（長さは pt、色は serializeColor() の形）にする
     * @param {Document} doc - 対象ドキュメント（カラーモード）
     * @param {number} glyphHeightPt - 〜Ratio に掛ける字面の高さ（pt）
     * @returns {{names: string[], map: Object}} 今の UI 言語のプリセット名（並べる順）と、名前をキーにした設定
     */
    function createBuiltinPresets(doc, glyphHeightPt) {
        var builtinPresets = { names: [], map: {} };
        for (var i = 0; i < BUILTIN_PRESETS.length; i++) {
            var sourceSettings = BUILTIN_PRESETS[i].settings;
            var presetData = {};
            for (var settingKey in sourceSettings) {
                if (!sourceSettings.hasOwnProperty(settingKey)) continue;
                var settingValue = sourceSettings[settingKey];
                if (/Ratio$/.test(settingKey)) {
                    presetData[settingKey.replace(/Ratio$/, "Pt")] = settingValue * glyphHeightPt;
                } else if (settingKey === "customColor") {
                    presetData.customColor = serializeColor(createDocumentColor(doc, settingValue.rgb, settingValue.cmyk));
                } else {
                    presetData[settingKey] = settingValue;
                }
            }
            var presetName = getLabel(BUILTIN_PRESETS[i].name);
            builtinPresets.names.push(presetName);
            builtinPresets.map[presetName] = presetData;
        }
        return builtinPresets;
    }

    /**
     * 色を保存できる形にする（RGB・CMYK・グレーのみ）
     * @param {Color} sourceColor - 色
     * @returns {{model: string, values: number[]}|null} 保存する形。ほかの種類は null
     */
    function serializeColor(sourceColor) {
        if (sourceColor.typename === "RGBColor") return { model: "rgb", values: [sourceColor.red, sourceColor.green, sourceColor.blue] };
        if (sourceColor.typename === "CMYKColor") return { model: "cmyk", values: [sourceColor.cyan, sourceColor.magenta, sourceColor.yellow, sourceColor.black] };
        if (sourceColor.typename === "GrayColor") return { model: "gray", values: [sourceColor.gray] };
        return null;
    }

    /**
     * serializeColor() の形から色を作り直す
     * @param {{model: string, values: number[]}|null} colorData - 保存した色
     * @returns {Color|null} 色。読めなければ null
     */
    function deserializeColor(colorData) {
        if (!colorData || !(colorData.values instanceof Array)) return null;
        var colorValues = colorData.values;
        if (colorData.model === "rgb") {
            var rgbColor = new RGBColor();
            rgbColor.red = colorValues[0];
            rgbColor.green = colorValues[1];
            rgbColor.blue = colorValues[2];
            return rgbColor;
        }
        if (colorData.model === "cmyk") {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = colorValues[0];
            cmykColor.magenta = colorValues[1];
            cmykColor.yellow = colorValues[2];
            cmykColor.black = colorValues[3];
            return cmykColor;
        }
        if (colorData.model === "gray") {
            var grayColor = new GrayColor();
            grayColor.gray = colorValues[0];
            return grayColor;
        }
        return null;
    }

    /**
     * ダイアログボックスの値をプリセットの形にまとめる（長さは pt で持つ）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {Object} プリセット
     */
    function collectPresetData(markerControls) {
        var markerOptions = readMarkerOptions(markerControls);
        return {
            lineStyle: markerOptions.lineStyle,
            shearAngle: markerOptions.shearAngle,
            zigzagSizePt: markerOptions.zigzagSize,
            zigzagRidges: markerOptions.zigzagRidges,
            boxPaddingPt: markerOptions.boxPadding,
            cornerRadiusPt: markerOptions.cornerRadius,
            colorChoice:markerControls.colorRow.getChoice(),
            textColorOpacity: parseFloat(markerControls.colorRow.opacityInput.text),
            customColor: serializeColor(markerControls.colorRow.getCustomColor()),
            strokeWidthPt: markerOptions.strokeWidth,
            verticalOffsetPt: markerOptions.verticalOffset,
            dashStyle: markerOptions.dashStyle,
            strokeCap: markerControls.strokeCapRow.getChoice(),
            whiteText: markerControls.getWhiteText()
        };
    }

    /**
     * 数値欄に値を書き込む（欄の単位・下限・上限にそろえる）
     * @param {EditText} numberInput - addMarkerNumberField() で作った入力欄
     * @param {number} fieldValue - 欄の単位の値
     * @returns {void}
     */
    function setMarkerFieldValue(numberInput, fieldValue) {
        writeSteppedValue(numberInput, fieldValue, numberInput.stepperGroup.stepOptions);
    }

    /**
     * プリセットをダイアログボックスに書き込む（無い項目・読めない値は今のまま）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetData(markerControls, presetData) {
        /* 旧版の［直線］は［マーカー］と同じ位置 / the former Straight sits where Marker does */
        if (presetData.lineStyle) markerControls.lineStyleRow.setChoice(presetData.lineStyle === "straight" ? "marker" : presetData.lineStyle);
        if (typeof presetData.shearAngle === "number") setMarkerFieldValue(markerControls.shearAngleInput, presetData.shearAngle);
        if (typeof presetData.zigzagSizePt === "number") {
            setMarkerFieldValue(markerControls.zigzagSizeInput, ptToUnit(presetData.zigzagSizePt, markerControls.rulerUnit));
        }
        if (typeof presetData.zigzagRidges === "number") setMarkerFieldValue(markerControls.zigzagRidgesInput, presetData.zigzagRidges);
        if (typeof presetData.boxPaddingPt === "number") {
            setMarkerFieldValue(markerControls.boxPaddingInput, ptToUnit(presetData.boxPaddingPt, markerControls.rulerUnit));
        }
        if (typeof presetData.cornerRadiusPt === "number") {
            setMarkerFieldValue(markerControls.cornerRadiusInput, ptToUnit(presetData.cornerRadiusPt, markerControls.rulerUnit));
        }
        var customColor = deserializeColor(presetData.customColor);
        if (customColor) markerControls.colorRow.setCustomColor(customColor);
        if (typeof presetData.textColorOpacity === "number") setMarkerFieldValue(markerControls.colorRow.opacityInput, presetData.textColorOpacity);
        if (presetData.colorChoice) markerControls.colorRow.setChoice(presetData.colorChoice);
        if (typeof presetData.strokeWidthPt === "number") {
            setMarkerFieldValue(markerControls.strokeWidthInput, ptToUnit(presetData.strokeWidthPt, markerControls.strokeUnit));
        }
        if (typeof presetData.verticalOffsetPt === "number") {
            setMarkerFieldValue(markerControls.verticalOffsetInput, ptToUnit(presetData.verticalOffsetPt, markerControls.rulerUnit));
        }
        if (presetData.dashStyle) markerControls.dashStyleRow.setChoice(presetData.dashStyle);
        if (presetData.strokeCap) markerControls.strokeCapRow.setChoice(presetData.strokeCap);
        if (typeof presetData.whiteText === "boolean") markerControls.setWhiteText(presetData.whiteText);
    }

    /**
     * プリセットのドロップダウンを作り直し、指定の名前を選ぶ。
     * 最初からあるプリセット、区切り線、保存したプリセット（名前順）の順に並べる
     * @param {Object} presetControls - addPresetRow() の戻り値（builtinPresets 付き）
     * @param {string|null} selectedName - 選んでおくプリセット名（null なら「---」）
     * @returns {void}
     */
    function fillPresetDropdown(presetControls, selectedName) {
        var presetDropdown = presetControls.presetDropdown;
        var builtinNames = presetControls.builtinPresets.names;
        var presetNames = builtinNames.concat(getSavedPresetNames(presetControls.builtinPresets.map));
        var selectedIndex = 0;
        presetDropdown.removeAll();
        presetDropdown.add("item", getLabel("dropdown.presetPlaceholder"));
        for (var i = 0; i < presetNames.length; i++) {
            if (i > 0 && i === builtinNames.length) presetDropdown.add("separator");
            presetDropdown.add("item", presetNames[i]);
            if (presetNames[i] === selectedName) selectedIndex = presetDropdown.items.length - 1;
        }
        presetDropdown.selection = selectedIndex;
        setIconButtonEnabled(presetControls.presetDeleteIcon, isSavedPreset(selectedName) && selectedIndex > 0);
    }

    /**
     * 選んでいるプリセット名を返す
     * @param {Object} presetControls - addPresetRow() の戻り値
     * @returns {string|null} プリセット名。「---」なら null
     */
    function getSelectedPresetName(presetControls) {
        var presetSelection = presetControls.presetDropdown.selection;
        return (presetSelection && presetSelection.index > 0) ? presetSelection.text : null;
    }

    /**
     * プリセット名を尋ねる
     * @param {string} initialName - 入力欄に入れておく名前
     * @returns {string|null} プリセット名。キャンセルか空なら null
     */
    function showPresetNameDialog(initialName) {
        var nameDialog = new Window("dialog", getLabel("dialog.presetSave"));
        setupWindow(nameDialog);
        var nameRow = nameDialog.add("group");
        setupRow(nameRow, "left", 8);
        nameRow.add("statictext", undefined, labelText("fieldLabel.presetName"));
        var nameInput = nameRow.add("edittext", undefined, initialName);
        nameInput.helpTip = getLabel("tooltip.presetName");
        nameInput.characters = PRESET_NAME_CHARS;
        nameInput.active = true;

        var buttonRow = addButtonRow(nameDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(nameDialog, SCRIPT_NAME + "_presetName");
        if (nameDialog.show() !== 1) return null;

        var presetName = nameInput.text.replace(/^\s+|\s+$/g, "");
        if (!presetName || presetName === getLabel("dropdown.presetPlaceholder")) return null;
        return presetName;
    }

    /**
     * 今の設定を名前を付けて保存する（同じ名前は、最初からあるものも含めて確認してから上書き）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetControls - addPresetRow() の戻り値（選んでいるプリセット名を名前の入力欄に入れておく）
     * @returns {string|null} 保存したプリセット名。保存しなかったら null
     */
    function saveCurrentPreset(markerControls, presetControls) {
        var presetName = showPresetNameDialog(getSelectedPresetName(presetControls) || "");
        if (!presetName) return null;
        var presetMap = loadPresetMap();
        var isExistingName = presetMap.hasOwnProperty(presetName) || presetControls.builtinPresets.map.hasOwnProperty(presetName);
        if (isExistingName && !confirm(getLabel("confirm.presetOverwrite", { name: presetName }), true)) return null;
        presetMap[presetName] = collectPresetData(markerControls);
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel("alert.presetSaveFailed"));
            return null;
        }
        return presetName;
    }

    /**
     * 選んでいるプリセットを確認してから削除する（最初からあるものは削除しない）
     * @param {Object} presetControls - addPresetRow() の戻り値
     * @returns {boolean} 削除したら true
     */
    function deleteSelectedPreset(presetControls) {
        var presetName = getSelectedPresetName(presetControls);
        if (!isSavedPreset(presetName) || !confirm(getLabel("confirm.presetDelete", { name: presetName }), true)) return false;
        var presetMap = loadPresetMap();
        delete presetMap[presetName];
        if (!presetSettingsStore.save(presetMap)) {
            alert(getLabel("alert.presetSaveFailed"));
            return false;
        }
        return true;
    }

    // =========================================
    // 選択 / Selection
    // =========================================

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // =========================================
    // 字面の測定 / Glyph measurement
    // =========================================

    /**
     * アウトラインのグループを展開し、字形（パス・複合パス）だけを平らな配列にして返す
     * @param {GroupItem} outlineGroup - createOutline() の戻り値
     * @returns {PageItem[]} 字形のパスと複合パス
     */
    function flattenOutlineGlyphs(outlineGroup) {
        var glyphItems = [];
        var pendingItems = [outlineGroup];
        while (pendingItems.length > 0) {
            var currentItem = pendingItems.pop();
            if (currentItem.typename !== "GroupItem") {
                glyphItems.push(currentItem);
                continue;
            }
            for (var i = 0; i < currentItem.pageItems.length; i++) {
                pendingItems.push(currentItem.pageItems[i]);
            }
        }
        return glyphItems;
    }

    /**
     * 字形に塗りがあるかを返す（複合パスは構成パスの塗りで判定する）
     * @param {PageItem} glyphItem - 字形のパスか複合パス
     * @returns {boolean} 塗りがあれば true
     */
    function hasGlyphFill(glyphItem) {
        var fillSource = (glyphItem.typename === "CompoundPathItem") ? glyphItem.pathItems[0] : glyphItem;
        return !!fillSource && fillSource.filled && fillSource.fillColor.typename !== "NoColor";
    }

    /**
     * 範囲を、字形の geometricBounds まで広げる
     * @param {Object|null} glyphExtent - 今の範囲 { left, top, right, bottom }（null なら新しく作る）
     * @param {number[]} itemBounds - 字形の geometricBounds [左, 上, 右, 下]
     * @returns {Object} 広げた範囲
     */
    function extendGlyphExtent(glyphExtent, itemBounds) {
        if (!glyphExtent) return { left: itemBounds[0], top: itemBounds[1], right: itemBounds[2], bottom: itemBounds[3] };
        if (itemBounds[0] < glyphExtent.left) glyphExtent.left = itemBounds[0];
        if (itemBounds[1] > glyphExtent.top) glyphExtent.top = itemBounds[1];
        if (itemBounds[2] > glyphExtent.right) glyphExtent.right = itemBounds[2];
        if (itemBounds[3] < glyphExtent.bottom) glyphExtent.bottom = itemBounds[3];
        return glyphExtent;
    }

    /**
     * 塗りのある字形をすべて囲む範囲を返す
     * @param {PageItem[]} glyphItems - 字形のパスと複合パス
     * @returns {Object|null} 範囲 { left, top, right, bottom }。塗りのある字形が無ければ null
     */
    function getFilledGlyphsExtent(glyphItems) {
        var glyphExtent = null;
        for (var i = 0; i < glyphItems.length; i++) {
            if (hasGlyphFill(glyphItems[i])) glyphExtent = extendGlyphExtent(glyphExtent, glyphItems[i].geometricBounds);
        }
        return glyphExtent;
    }

    /**
     * 文字番号を、ドキュメントのカラーモードの色に置き換える（RGB は 256 進、CMYK は 100 進で C・M・Y に振る）
     * @param {number} charIndex - 文字番号（0始まり）
     * @param {boolean} isRgbDocument - RGB のドキュメントか
     * @returns {Color} 文字番号を表す色
     */
    function createCharIndexColor(charIndex, isRgbDocument) {
        if (isRgbDocument) {
            var rgbColor = new RGBColor();
            rgbColor.red = charIndex % 256;
            rgbColor.green = Math.floor(charIndex / 256) % 256;
            rgbColor.blue = Math.floor(charIndex / 65536) % 256;
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = charIndex % 100;
        cmykColor.magenta = Math.floor(charIndex / 100) % 100;
        cmykColor.yellow = Math.floor(charIndex / 10000) % 100;
        cmykColor.black = 0;
        return cmykColor;
    }

    /**
     * createCharIndexColor() の色から文字番号を読み戻す
     * @param {Color} indexColor - 字形の塗り
     * @param {boolean} isRgbDocument - RGB のドキュメントか
     * @returns {number} 文字番号（読めない色なら -1）
     */
    function readCharIndexColor(indexColor, isRgbDocument) {
        if (isRgbDocument && indexColor.typename === "RGBColor") {
            return Math.round(indexColor.red) + Math.round(indexColor.green) * 256 + Math.round(indexColor.blue) * 65536;
        }
        if (!isRgbDocument && indexColor.typename === "CMYKColor") {
            return Math.round(indexColor.cyan) + Math.round(indexColor.magenta) * 100 + Math.round(indexColor.yellow) * 10000;
        }
        return -1;
    }

    /**
     * 選んだ範囲の字面の範囲を測る。
     * 下書き用の複製では選んだ範囲の外の塗りを外しておき、アウトライン化して塗りの残った字形だけを囲む
     * @param {TextFrame} sourceFrame - 元のテキストフレーム
     * @param {number} spanStart - 選んだ範囲の先頭の文字番号（フレーム内、0始まり）
     * @param {number} spanLength - 選んだ範囲の文字数
     * @returns {Object|null} 字面の範囲 { left, top, right, bottom }。字形が無ければ null
     */
    function measureGlyphBounds(sourceFrame, spanStart, spanLength) {
        var draftFrame = sourceFrame.duplicate(sourceFrame, ElementPlacement.PLACEAFTER);
        var draftFill = new GrayColor();
        draftFill.gray = 100;

        /* 全体の塗りを外し、選んだ範囲だけ1つの TextRange として塗り直す
           Clear the fill of the whole copy, then refill only the selected span as one TextRange */
        draftFrame.textRange.characterAttributes.fillColor = new NoColor();
        var draftSpan = draftFrame.characters[spanStart];
        draftSpan.length = spanLength;
        draftSpan.characterAttributes.fillColor = draftFill;

        /* createOutline() は複製を消費するので、片付けるのは戻り値のグループ / createOutline() consumes the copy, so remove the returned group */
        var outlineGroup = draftFrame.createOutline();
        var glyphBounds = getFilledGlyphsExtent(flattenOutlineGlyphs(outlineGroup));
        outlineGroup.remove();
        return glyphBounds;
    }

    /**
     * 選んだ範囲の字面の範囲を1文字ずつ測る。
     * 下書き用の複製で選んだ範囲の文字を文字番号ごとに別の色で塗り、1回のアウトライン化で字形を文字に振り分ける
     * @param {TextFrame} sourceFrame - 元のテキストフレーム
     * @param {number} spanStart - 選んだ範囲の先頭の文字番号（フレーム内、0始まり）
     * @param {number} spanLength - 選んだ範囲の文字数
     * @returns {Object[]} 字面の範囲 { left, top, right, bottom } の配列（文字の順。スペースなど字形の無い文字は除く）
     */
    function measureCharGlyphBounds(sourceFrame, spanStart, spanLength) {
        /* 色はドキュメントのカラーモードのまま読み戻せるように、同じモードで作る / Use the document color mode so the colors read back unchanged */
        var isRgbDocument = (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB);
        var draftFrame = sourceFrame.duplicate(sourceFrame, ElementPlacement.PLACEAFTER);
        draftFrame.textRange.characterAttributes.fillColor = new NoColor();
        for (var i = 0; i < spanLength; i++) {
            draftFrame.characters[spanStart + i].characterAttributes.fillColor = createCharIndexColor(i, isRgbDocument);
        }

        var outlineGroup = draftFrame.createOutline();
        var glyphItems = flattenOutlineGlyphs(outlineGroup);
        var boundsByIndex = [];
        for (var j = 0; j < glyphItems.length; j++) {
            if (!hasGlyphFill(glyphItems[j])) continue;
            var fillSource = (glyphItems[j].typename === "CompoundPathItem") ? glyphItems[j].pathItems[0] : glyphItems[j];
            var charIndex = readCharIndexColor(fillSource.fillColor, isRgbDocument);
            if (charIndex < 0 || charIndex >= spanLength) continue;
            boundsByIndex[charIndex] = extendGlyphExtent(boundsByIndex[charIndex] || null, glyphItems[j].geometricBounds);
        }
        outlineGroup.remove();

        var charBoundsList = [];
        for (var k = 0; k < spanLength; k++) {
            if (boundsByIndex[k]) charBoundsList.push(boundsByIndex[k]);
        }
        return charBoundsList;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ［ジグザグ］効果の XML を返す（大きさは絶対値、折り返しの角は直線）
     * @param {number} amountPt - 大きさ（pt）
     * @param {number} ridges - 線分あたりの折り返し
     * @returns {string} LiveEffect の XML
     */
    function createZigzagXml(amountPt, ridges) {
        return '<LiveEffect name="Adobe Zigzag"><Dict data="R amount ' + amountPt +
            ' R relAmount 0 R absoluteness 1 R ridges ' + ridges + ' R roundness 0 "/></LiveEffect>';
    }

    /**
     * 上辺を横にずらす［パスの自由変形］効果の XML を返す。
     * 4隅は線の外形（線幅を含む範囲）を 0〜1 にした座標で、順序は [左上, 右上, 左下, 右下]
     * @param {number} topShiftRatio - 上辺を右へずらす量（外形の幅に対する比率。マイナスで左）
     * @returns {string} LiveEffect の XML
     */
    function createShearXml(topShiftRatio) {
        var sourceCorners = [[0, 0], [1, 0], [0, 1], [1, 1]];
        var destCorners = [[topShiftRatio, 0], [1 + topShiftRatio, 0], [0, 1], [1, 1]];
        var cornerData = "";
        for (var i = 0; i < 4; i++) {
            /* Y は Illustrator の座標系に合わせて符号を反転 / flip Y for Illustrator's axis */
            cornerData += "R src" + i + "h " + sourceCorners[i][0] + " R src" + i + "v " + (-sourceCorners[i][1]) + " ";
        }
        for (var j = 0; j < 4; j++) {
            cornerData += "R dst" + j + "h " + destCorners[j][0] + " R dst" + j + "v " + (-destCorners[j][1]) + " ";
        }
        return '<LiveEffect name="Adobe Free Distort"><Dict data="' + cornerData + '"/></LiveEffect>';
    }

    /**
     * パスに［効果］＞［パス］＞［パスのアウトライン］を掛ける。
     * XML の applyEffect では付かないので、パスだけを選んでメニューコマンドで掛け、選択を空に戻す（元の選択は main() で戻す）。
     * 入れ子のグループの中で実行すると一番外側のグループに付くので、レイヤー直下にあるうちに呼ぶ
     * @param {PathItem} targetPath - 対象のパス（レイヤー直下）
     * @returns {void}
     */
    function applyOutlineStrokeEffect(targetPath) {
        var doc = app.activeDocument;
        doc.selection = null;
        targetPath.selected = true;
        /* 直前の DOM の変更が画面に反映されていないと、メニューコマンドが空振りする / the command misfires unless the DOM changes are redrawn first */
        app.redraw();
        app.executeMenuCommand("Live Outline Stroke");
        /* 効果を確定させてから次の効果を足す / commit the effect before the next one is added */
        app.redraw();
        doc.selection = null;
    }

    /**
     * 元のテキストの直後（背面）へマーカーの線を引く。
     * ［塗りつぶし］［取り消し線］は字面の上下中央、［アンダーライン］は字面の下端に線の上端、ほかは字面の下端を線の中心にする
     * @param {Object} markerTarget - collectMarkerTargets() の要素 { sourceFrame, glyphBounds, textColor }
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {PathItem} 引いた線
     */
    function placeMarkerLine(markerTarget, markerOptions) {
        var sourceFrame = markerTarget.sourceFrame;
        var glyphBounds = markerTarget.glyphBounds;
        var markerLine = sourceFrame.layer.pathItems.add();
        var lineStyle = markerOptions.lineStyle;
        var lineY = glyphBounds.bottom;
        if (lineStyle === "fill" || lineStyle === "strikethrough") lineY = (glyphBounds.top + glyphBounds.bottom) / 2;
        else if (lineStyle === "underline") lineY = glyphBounds.bottom - markerOptions.strokeWidth / 2;
        /* ［上下位置］の分だけずらす（プラスで上） / Shift by the vertical offset (positive is up) */
        lineY += markerOptions.verticalOffset;
        markerLine.setEntirePath([[glyphBounds.left, lineY], [glyphBounds.right, lineY]]);
        markerLine.filled = false;
        markerLine.stroked = true;
        markerLine.strokeColor = markerOptions.useTextColor ? markerTarget.textColor : markerOptions.color;
        /* ［塗りつぶし］は字面の高さちょうどの線にする / Fill uses a line exactly as wide as the glyph height */
        markerLine.strokeWidth = (lineStyle === "fill") ? (glyphBounds.top - glyphBounds.bottom) : markerOptions.strokeWidth;
        /* ドット点線は長さ0の線分を丸型の線端で点にする / Dotted draws zero-length dashes as dots with round caps */
        var dashStyle = markerOptions.dashStyle;
        var lineWidth = markerLine.strokeWidth;
        markerLine.strokeCap = (dashStyle === "dottedLine") ? StrokeCap.ROUNDENDCAP : markerOptions.strokeCap;
        if (dashStyle === "dashedLine") markerLine.strokeDashes = [lineWidth * DASHED_LINE_RATIOS[0], lineWidth * DASHED_LINE_RATIOS[1]];
        else if (dashStyle === "dottedLine") markerLine.strokeDashes = [0, lineWidth * DOTTED_LINE_PITCH_RATIO];
        else markerLine.strokeDashes = [];
        markerLine.opacity = markerOptions.opacity;
        /* 効果は分割せずに残す。後から掛けたものほど下に並び、後で処理される
           Effects stay live; later ones go lower in the list and are processed later */
        /* ギザギザは中心線を折ってからアウトラインにする（線は1本なので、折り返しは端から端までの数）
           Zigzag folds the center line before it is outlined; the line has one segment, so ridges span end to end */
        if (lineStyle === "zigzag") {
            markerLine.applyEffect(createZigzagXml(markerOptions.zigzagSize, markerOptions.zigzagRidges));
        }
        applyOutlineStrokeEffect(markerLine);
        /* シアーはアウトラインにした面（線幅の高さがある）を変形する / Shear distorts the outlined shape, which has the stroke's height */
        if (lineStyle === "straightShear") {
            /* 線幅（外形の高さ）× tan(角度) だけ上辺をずらす。比率は線の外形の幅に対して取る
               Shift the top edge by stroke width x tan(angle), as a ratio of the outline width */
            var outlineBounds = markerLine.visibleBounds;
            var topShiftPt = markerOptions.strokeWidth * Math.tan(markerOptions.shearAngle * Math.PI / 180);
            markerLine.applyEffect(createShearXml(topShiftPt / (outlineBounds[2] - outlineBounds[0])));
        }
        /* 効果を掛け終えてから元のテキストの隣（グループの中かもしれない）へ移す / Move next to the text (possibly inside a group) after the effects are applied */
        markerLine.move(sourceFrame, ElementPlacement.PLACEAFTER);
        return markerLine;
    }

    /**
     * 1文字ずつの字面の範囲を行に分ける。
     * 文字の上下中央がそれまでの行の上端・下端の間から外れたら、次の行とみなす
     * @param {Object[]} charBoundsList - measureCharGlyphBounds() の戻り値（文字の順）
     * @returns {Object[]} 行 { left, top, right, bottom, chars }（範囲は行の字面全体、chars はその行の文字の範囲）
     */
    function groupCharBoundsByLine(charBoundsList) {
        var textLines = [];
        var currentLine = null;
        for (var i = 0; i < charBoundsList.length; i++) {
            var charBounds = charBoundsList[i];
            var centerY = (charBounds.top + charBounds.bottom) / 2;
            if (!currentLine || centerY > currentLine.top || centerY < currentLine.bottom) {
                currentLine = { left: charBounds.left, top: charBounds.top, right: charBounds.right, bottom: charBounds.bottom, chars: [] };
                textLines.push(currentLine);
            }
            extendGlyphExtent(currentLine, [charBounds.left, charBounds.top, charBounds.right, charBounds.bottom]);
            currentLine.chars.push(charBounds);
        }
        return textLines;
    }

    /**
     * 長方形を描く範囲（余白を足す前）を返す。
     * ［長方形］は行ごとに行の字面全体、［1文字ずつの長方形］は文字ごとに、高さを行の字面にそろえ、
     * 幅は字面の幅か高さの大きいほうで字面の左右中央に置く
     * @param {Object[]} charBoundsList - measureCharGlyphBounds() の戻り値（文字の順）
     * @param {boolean} isPerChar - 1文字ずつにするか
     * @returns {Object[]} 範囲 { left, top, right, bottom }
     */
    function getBoxAreas(charBoundsList, isPerChar) {
        var boxAreas = [];
        var textLines = groupCharBoundsByLine(charBoundsList);
        for (var i = 0; i < textLines.length; i++) {
            var textLine = textLines[i];
            if (!isPerChar) {
                boxAreas.push({ left: textLine.left, top: textLine.top, right: textLine.right, bottom: textLine.bottom });
                continue;
            }
            var lineHeight = textLine.top - textLine.bottom;
            for (var j = 0; j < textLine.chars.length; j++) {
                var charBounds = textLine.chars[j];
                var halfWidth = Math.max(charBounds.right - charBounds.left, lineHeight) / 2;
                var centerX = (charBounds.left + charBounds.right) / 2;
                boxAreas.push({ left: centerX - halfWidth, top: textLine.top, right: centerX + halfWidth, bottom: textLine.bottom });
            }
        }
        return boxAreas;
    }

    /**
     * 元のテキストの直後（背面）へ、字面を囲む長方形を描いてグループにする。
     * まわりに［余白］を足し、［角丸の半径］が 0 より大きければ角を丸める（半径は短い辺の半分まで）
     * @param {Object} markerTarget - collectMarkerTargets() の要素（1文字ずつの範囲は初回に測って charBounds に控える）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {GroupItem} 長方形のグループ
     */
    function placeTextBoxes(markerTarget, markerOptions) {
        if (!markerTarget.charBounds) {
            markerTarget.charBounds = measureCharGlyphBounds(markerTarget.sourceFrame, markerTarget.spanStart, markerTarget.spanLength);
        }
        var sourceFrame = markerTarget.sourceFrame;
        var boxColor = markerOptions.useTextColor ? markerTarget.textColor : markerOptions.color;
        var boxPadding = markerOptions.boxPadding;
        var boxGroup = sourceFrame.layer.groupItems.add();
        var boxAreas = getBoxAreas(markerTarget.charBounds, markerOptions.lineStyle === "charBox");
        for (var i = 0; i < boxAreas.length; i++) {
            var boxWidth = boxAreas[i].right - boxAreas[i].left + boxPadding * 2;
            var boxHeight = boxAreas[i].top - boxAreas[i].bottom + boxPadding * 2;
            var boxLeft = boxAreas[i].left - boxPadding;
            /* ［上下位置］の分だけずらす（プラスで上） / Shift by the vertical offset (positive is up) */
            var boxTop = boxAreas[i].top + boxPadding + markerOptions.verticalOffset;
            var cornerRadius = Math.min(markerOptions.cornerRadius, boxWidth / 2, boxHeight / 2);
            var textBox = (cornerRadius > 0)
                ? boxGroup.pathItems.roundedRectangle(boxTop, boxLeft, boxWidth, boxHeight, cornerRadius, cornerRadius)
                : boxGroup.pathItems.rectangle(boxTop, boxLeft, boxWidth, boxHeight);
            textBox.filled = true;
            textBox.fillColor = boxColor;
            textBox.stroked = false;
        }
        boxGroup.opacity = markerOptions.opacity;
        boxGroup.move(sourceFrame, ElementPlacement.PLACEAFTER);
        return boxGroup;
    }

    /**
     * すべての対象にマーカーの線（［長方形］［1文字ずつの長方形］なら長方形のグループ）を引く
     * @param {Object[]} markerTargets - collectMarkerTargets() の戻り値
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {PageItem[]} 引いた線と長方形のグループ
     */
    function placeMarkerLines(markerTargets, markerOptions) {
        var markerLines = [];
        for (var i = 0; i < markerTargets.length; i++) {
            if (isBoxLineStyle(markerOptions.lineStyle)) markerLines.push(placeTextBoxes(markerTargets[i], markerOptions));
            else markerLines.push(placeMarkerLine(markerTargets[i], markerOptions));
        }
        return markerLines;
    }

    /**
     * 対象の文字の範囲の塗りを変える
     * @param {TextFrame} textFrame - 対象の元のテキストフレーム、またはその複製
     * @param {Object} markerTarget - collectMarkerTargets() の要素（spanStart・spanLength を使う）
     * @param {Color} fillColor - 塗りの色
     * @returns {void}
     */
    function fillTargetSpan(textFrame, markerTarget, fillColor) {
        var targetSpan = textFrame.characters[markerTarget.spanStart];
        targetSpan.length = markerTarget.spanLength;
        targetSpan.characterAttributes.fillColor = fillColor;
    }

    /**
     * 文字の範囲の先頭の文字の塗りを返す
     * @param {TextRange} targetRange - 文字の範囲
     * @returns {Color|null} 塗りのカラー。塗りなしなら null
     */
    function getTextFillColor(targetRange) {
        var fillColor = targetRange.characters[0].characterAttributes.fillColor;
        return (fillColor && fillColor.typename !== "NoColor") ? fillColor : null;
    }

    /**
     * すべての対象で［文字色］が使えるかを返す
     * @param {Object[]} markerTargets - collectMarkerTargets() の戻り値
     * @returns {boolean} どの対象の先頭の文字にも塗りがあれば true
     */
    function areTextColorsUsable(markerTargets) {
        for (var i = 0; i < markerTargets.length; i++) {
            if (!markerTargets[i].textColor) return false;
        }
        return true;
    }

    /**
     * 選択からマーカーを引く対象を集める。
     * 文字ツールで選んだ文字ならその範囲、テキストオブジェクトならそれぞれのすべての文字を1つの対象にする
     * @param {TextRange|Array} docSelection - doc.selection
     * @returns {Object[]|null} 対象 { sourceFrame, spanStart, spanLength, glyphBounds, textColor } の配列（字形の無いものは除く）。
     *     テキストが選ばれていなければ null
     */
    function collectMarkerTargets(docSelection) {
        var markerTargets = [];

        function addMarkerTarget(sourceFrame, spanStart, spanLength, colorRange) {
            var glyphBounds = measureGlyphBounds(sourceFrame, spanStart, spanLength);
            if (!glyphBounds) return;
            markerTargets.push({
                sourceFrame: sourceFrame, spanStart: spanStart, spanLength: spanLength,
                glyphBounds: glyphBounds, textColor: getTextFillColor(colorRange)
            });
        }

        if (docSelection && docSelection.typename === "TextRange") {
            /* カーソルだけ（長さ 0）は対象外 / a bare insertion point (length 0) is skipped */
            if (docSelection.length === 0) return null;
            /* 選んだ範囲の位置を、ストーリーの先頭のフレームからの文字番号にする
               Convert the selection to character indexes counted from the story's first frame */
            var storyFrame = docSelection.story.textFrames[0];
            addMarkerTarget(storyFrame, docSelection.start - storyFrame.textRange.start, docSelection.length, docSelection);
            return markerTargets;
        }

        var textFrames = collectSelectionTextFrames(docSelection, { textRangeToFrame: false, skipLocked: true, skipHidden: true });
        if (textFrames.length === 0) return null;
        for (var i = 0; i < textFrames.length; i++) {
            /* 空のテキストは測るものが無い / an empty text frame has nothing to measure */
            if (textFrames[i].characters.length === 0) continue;
            addMarkerTarget(textFrames[i], 0, textFrames[i].characters.length, textFrames[i].textRange);
        }
        return markerTargets;
    }

    /**
     * プレビューで外した選択を戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {TextRange|Array} savedSelection - 実行時の doc.selection
     * @returns {void}
     */
    function restoreSelection(doc, savedSelection) {
        /* select() は今の選択に足すので、先に選択を空にする / select() adds to the current selection, so clear it first */
        doc.selection = null;
        if (savedSelection.typename === "TextRange") savedSelection.select();
        else doc.selection = savedSelection;
    }

    /**
     * 選んだ文字・テキストオブジェクトの背面にマーカーを引く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var savedSelection = doc.selection;

        var markerTargets = collectMarkerTargets(savedSelection);
        if (!markerTargets) {
            alert(getLabel("alert.noTextSelection"));
            return;
        }
        if (markerTargets.length === 0) {
            alert(getLabel("alert.noGlyphs"));
            return;
        }

        var markerOptions = showMarkerDialog(doc, markerTargets);
        if (markerOptions) {
            placeMarkerLines(markerTargets, markerOptions);
            if (markerOptions.whiteText) {
                var whiteColor = createDocumentColor(doc, WHITE_TEXT_COLOR.rgb, WHITE_TEXT_COLOR.cmyk);
                for (var i = 0; i < markerTargets.length; i++) fillTargetSpan(markerTargets[i].sourceFrame, markerTargets[i], whiteColor);
            }
        }
        restoreSelection(doc, savedSelection);
    }

    main();
})();
