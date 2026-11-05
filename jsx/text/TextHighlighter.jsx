#target illustrator
#targetengine "TextHighlighterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.DrawMarkerBehindText
Name=文字にマーカー・アンダーライン・取り消し線を引く
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.7.6
Release-Date=2026-11-01
Description-BEGIN
文字ツールで選んだ文字やテキストオブジェクトに、字面に合わせてマーカー・アンダーライン・取り消し線などの線や、長方形・円・六角形・傍点を描きます（文字の背面・前面を選べ、複数行は行ごと。行頭・行長・行送りの間隔もそろえられます）。
種類・カラー・線幅などはダイアログボックスで選び、文字のカラー・カーニング（図形の余白に合わせた前後のアキ）も一緒に変えられます。設定はプリセットとして保存できます。

Draws lines such as markers, underlines and strikethroughs, or boxes, circles, hexagons and emphasis dots, on the characters selected with the Type tool or on text objects, fitted to their glyphs (behind or in front of the text, line by line, with line starts, lengths and leading optionally aligned).
Choose the style, colors, weight and more in the dialog box, and change the text color and kerning (room before and after for the shape's padding) at the same time. The settings can be saved as presets.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

文字ツールで選んだ文字やテキストオブジェクトに、字面に合わせてマーカー・アンダーライン・取り消し線などの線や、長方形・円・六角形・傍点を描きます（文字の背面・前面を選べ、複数行は行ごと。行頭・行長・行送りの間隔もそろえられます）。
種類・カラー・線幅などはダイアログボックスで選び、文字のカラー・カーニング（図形の余白に合わせた前後のアキ）も一緒に変えられます。設定はプリセットとして保存できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextHighlighter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na1804afb7fe2

### Overview

Draws lines such as markers, underlines and strikethroughs, or boxes, circles, hexagons and emphasis dots, on the characters selected with the Type tool or on text objects, fitted to their glyphs (behind or in front of the text, line by line, with line starts, lengths and leading optionally aligned).
Choose the style, colors, weight and more in the dialog box, and change the text color and kerning (room before and after for the shape's padding) at the same time. The settings can be saved as presets.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextHighlighter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextHighlighter";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "文字にマーカー・アンダーライン・取り消し線を引く"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextHighlighter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextHighlighter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na1804afb7fe2"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextHighlighter.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author デザインゴリラ（字面の測り方） / glyph measurement
 * @discussion https://note.com/gorilla_gg/n/naa26b079cb75
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［黄色］［赤］［白］のカラーと、［カスタム］の最初のカラー（色見本のクリックで変えられる）
       Colors of Yellow, Red and White, and the initial Custom color (changeable by clicking the swatch) */
    var MARKER_COLOR_PRESETS = {
        yellow: { cmyk: [0, 0, 100, 0], rgb: [255, 241, 0] },   /* 黄色（Y100） / yellow (Y100) */
        red:    { cmyk: [0, 100, 100, 0], rgb: [230, 0, 18] },  /* 赤（M100Y100） / red (M100Y100) */
        white:  { cmyk: [0, 0, 0, 0], rgb: [255, 255, 255] },   /* 白 / white */
        custom: { cmyk: [100, 0, 0, 0], rgb: [0, 160, 233] }    /* カスタムの最初の色（C100） / initial custom color (C100) */
    };

    /* ［塗り］（長方形・円・六角形の塗り、塗りつぶし、傍点）と［線］のカラーに並べる選択肢（並べる順。"custom" は先頭に色見本付きで置く）
       Choices listed in Fill Color and Stroke Color (in order; "custom" goes first with its swatch) */
    var FILL_COLOR_CHOICES = ["textColor", "white", "custom"];
    var STROKE_COLOR_CHOICES = ["textColor", "yellow", "red", "custom"];

    /* 開いたときに選んでおく［塗り］のカラー（FILL_COLOR_CHOICES のどれか） / Fill Color chosen when the dialog box opens */
    var DEFAULT_FILL_COLOR_CHOICE = "textColor";

    /* 線・図形の不透明度の初期値（%） / Initial opacity (%) of the lines and shapes */
    var DEFAULT_OPACITY = 100;

    /* 線幅の初期値：字面の高さに対する比率 / Initial line width as a ratio of the glyph height */
    var MARKER_WEIGHT_RATIO = 0.4;

    /* 種類の既定（"underline"：アンダーライン / "marker"：マーカー / "straightShear"：直線＋シアー / "zigzag"：ギザギザ /
       "fill"：塗りつぶし / "strikethrough"：取り消し線 / "doubleStrikethrough"：二重取り消し線 /
       "box"：長方形 / "circle"：円 / "hexagon"：六角形 / "emphasisDots"：傍点） / Default style */
    var DEFAULT_LINE_STYLE = "marker";

    /* ［余白］の初期値：字面の高さに対する比率（長方形・円・六角形は字面のまわり、傍点は字面から点まで）
       Initial padding as a ratio of the glyph height (around the glyphs for boxes and frames, glyph-to-dot distance for emphasis dots) */
    var BOX_PADDING_RATIO = 0.08;

    /* ［二重取り消し線］の線の間隔の初期値：字面の高さに対する比率 / Initial gap of Double Strikethrough as a ratio of the glyph height */
    var LINE_GAP_RATIO = 0.2;

    /* ［傍点］の点の大きさ（直径）の初期値：字面の高さに対する比率 / Initial dot diameter of Emphasis Dots as a ratio of the glyph height */
    var DOT_SIZE_RATIO = 0.18;

    /* ［傍点］と［1文字ずつ］の図形で飛ばす文字（句読点・括弧・記号・スペース）
       Characters skipped by Emphasis Dots and the per-character shapes (punctuation, brackets, symbols, spaces) */
    var PER_CHAR_SKIP_CHARS = "、。，．,.・:;：；!?！？「」『』（）()［］[]【】〈〉《》〔〕…‥ー―-— 　\t\r\n";

    /* ［ラフ］の初期値：［手書き風］のオンオフ、サイズ（線や図形の大きさに対する％）。詳細は［折り返し］の値を使う
       Initial Roughen values: Hand-Drawn on/off and size (% of the line or shape); the detail comes from Ridges */
    var DEFAULT_ROUGHEN = false;
    var DEFAULT_ROUGHEN_SIZE_PERCENT = 1;

    /* 字面を測るアウトライン化の回数の上限。アウトライン直後の字形に読めない参照（MRAP）が混じったら作り直す
       Maximum outline attempts when measuring glyphs; retried when an outlined glyph cannot be read (MRAP) */
    var OUTLINE_ATTEMPT_LIMIT = 3;

    /* ［長方形］［円］［六角形］の［1文字ずつ］の初期値 / Initial state of Per Character for Box, Circle and Hexagon */
    var DEFAULT_PER_CHAR = false;

    /* ［長方形］［円］［六角形］の［塗り］［線］の［描画］の初期値 / Initial Draw of Fill and Stroke for Box, Circle and Hexagon */
    var DEFAULT_SHAPE_FILL = true;
    var DEFAULT_SHAPE_STROKE = false;

    /* 開いたときに選んでおく［線］のカラー（アンダーライン・マーカーなどの線と、長方形・円・六角形の線）（STROKE_COLOR_CHOICES のどれか）
       Stroke Color chosen when the dialog box opens (lines such as Underline and Marker, and the stroke of Box, Circle and Hexagon) */
    var DEFAULT_STROKE_COLOR_CHOICE = "yellow";

    /* ［自動カーニング］の初期値（"keepKerning"：変更しない / "zeroKerning"：0 / "metricsKerning"：メトリクス / "opticalKerning"：オプティカル）と、
       ［トラッキング］の初期値（空欄なら変更しない。1/1000 em）
       Initial Auto Kerning choice, and initial Tracking (blank leaves it unchanged; 1/1000 em) */
    var DEFAULT_KERNING_CHOICE = "keepKerning";
    var DEFAULT_TRACKING_TEXT = "";

    /* ［カーニング（前）］［カーニング（後）］の初期値（手動カーニング。空欄なら変更しない。1/1000 em）
       Initial Kerning (Before) / (After) (manual kerning; blank leaves it unchanged; 1/1000 em) */
    var DEFAULT_KERNING_BEFORE_TEXT = "";
    var DEFAULT_KERNING_AFTER_TEXT = "";

    /* ［余白に合わせてカーニングを自動調整］の初期値（［長方形］［円］［六角形］を行ごとに描くとき、カーニング（前・後）を［余白（横）］÷ 文字サイズ × 1000 にする。縦組みは［余白（縦）］）
       Initial state of Auto-Adjust Kerning to Padding (for Box, Circle and Hexagon per line, Kerning (Before) / (After) become Padding (H) / type size x 1000; Padding (V) in vertical text) */
    var DEFAULT_KERNING_FROM_PADDING = true;

    /* ［テキストカラー］に並べる選択肢（"keepTextColor" は文字の塗りを変えない。"custom" は先頭に色見本付きで置く）と、開いたときに選んでおくもの
       Choices listed in Text Color ("keepTextColor" leaves the fill unchanged; "custom" goes first with its swatch), and the one chosen when the dialog box opens */
    var TEXT_FILL_COLOR_CHOICES = ["keepTextColor", "white", "custom"];
    var DEFAULT_TEXT_FILL_COLOR_CHOICE = "keepTextColor";

    /* ［長方形］の角丸の半径の初期値（pt。0 で角丸なし） / Initial corner radius of the boxes (pt; 0 for square corners) */
    var DEFAULT_CORNER_RADIUS = 0;

    /* ［複数行］の初期値（複数行の［長方形］で、先頭の行の行頭側と最後の行の行末側だけを角丸にする）
       Initial state of Square Corners at Line Breaks (a multi-line Box rounds only the start of the first line and the end of the last) */
    var DEFAULT_SQUARE_LINE_JOINS = true;

    /* ［長方形］［円］［六角形］の回転の初期値（度。プラスで反時計回り） / Initial rotation of the shapes (degrees; positive is counterclockwise) */
    var DEFAULT_BOX_ROTATION = 0;

    /* ［直線＋シアー］の角度の初期値（度。プラスで右上がり） / Initial shear angle in degrees (positive slants to the upper right) */
    var DEFAULT_SHEAR_ANGLE = 30;

    /* ギザギザ（ジグザグ効果）の初期値。大きさは［手書き風］の揺れの大きさにも使う
       Initial values of the zigzag (Zig Zag effect); the size is also the Hand-Drawn roughness */
    var DEFAULT_ZIGZAG_SIZE = 3;      /* 大きさ（pt） / size (pt) */
    var DEFAULT_ZIGZAG_RIDGES = 12;   /* 折り返し（1インチあたり。［ラフ］効果の詳細） / ridges per inch (Roughen detail) */
    var DEFAULT_ZIGZAG_POINTS = "cornerPoints"; /* ポイント（"smoothPoints"：滑らか / "cornerPoints"：直線的） / points */

    /* 線の既定（"solidLine"：直線 / "dashedLine"：破線 / "dottedLine"：ドット） / Default line type */
    var DEFAULT_DASH_STYLE = "solidLine";

    /* 破線の線分の長さの目安と、［間隔］（破線の線分のあいだ・ドットの点のあいだ）の初期値：線幅に対する比率。
       破線は［分割数］が空欄ならこの長さに近い本数、ドットは点のあいだが［間隔］に近くなる数で、線の長さを割り振る（両端が線分・点で終わる）
       Typical dash length and initial Gap (between dashes or between dots) as ratios of the line width.
       Dashes near this length (with blank Segments) or dots near Gap are fitted so both ends finish with a dash or a dot */
    var DASHED_LINE_RATIOS = [2, 1];
    /* 破線の［分割数］の初期値（線ごとの線分の数。空欄なら自動） / Initial Segments per line for Dashed (blank fits automatically) */
    var DEFAULT_DASH_SEGMENTS_TEXT = "";
    /* 両端に点を置くドットの間隔の切り捨て単位（1/1000 pt）。長さ0の点は終点をわずかでも越えると描かれない
       Truncation unit for the dot pitch with dots at both ends; a zero-length dot past the end is not drawn */
    var DOT_GAP_PRECISION = 1000;

    /* 線端の既定（"buttCap"：なし / "roundCap"：丸型） / Default line cap */
    var DEFAULT_STROKE_CAP = "buttCap";

    /* ［行長］の［行頭］の初期値（複数行のとき、線や図形の始まりを、いちばん前から始まる行にそろえる。縦組みは上端）
       Initial state of Align Starts (lines and shapes start where the earliest line starts; the top in vertical text) */
    var DEFAULT_ALIGN_LINE_STARTS = false;

    /* ［行長］の［行頭行末］の初期値（複数行のとき、短い行の線や図形もいちばん長い行の長さにそろえる）
       Initial state of Align Both Ends (short lines get the length of the longest one) */
    var DEFAULT_EQUAL_LINE_LENGTHS = false;

    /* ［行の高さ］の［行送り基準］の初期値（行ごとの図形の高さ（縦組みは幅）を、字面ではなく行送りで決める）
       Initial state of Leading-Based under Line Height (the line shapes take their height, or width in vertical text, from the leading) */
    var DEFAULT_LEADING_BASED = false;

    /* ［余白（横）］［余白（縦）］の連動の初期値 / Initial link state of Padding (H) and Padding (V) */
    var DEFAULT_BOX_PADDING_LINKED = true;

    /* 重ね順の既定（"stackFront"：文字の前 / "stackBack"：文字の後ろ） / Default stacking order */
    var DEFAULT_STACK_ORDER = "stackBack";

    /* Illustrator を起動して最初に開いたときに選ぶプリセット（BUILTIN_PRESETS の名前）。2回目からは前回閉じたときの状態に戻す
       Preset chosen the first time after Illustrator starts (a BUILTIN_PRESETS name); later, the state at the last close comes back */
    var INITIAL_PRESET_NAME = { ja: "黄色のマーカー", en: "Yellow Marker" };

    /* 最初から用意するプリセット（ドロップダウンの先頭にこの順で並ぶ。削除はできず、同じ名前で保存したものが優先）。
       書き方は保存したプリセットと同じで、長さは〜Pt のかわりに〜Ratio（先頭の対象の字面の高さに対する比率）でも書ける。
       customColor・strokeCustomColor・textCustomColor は { rgb, cmyk } で、ドキュメントのカラーモードに合わせて使い分ける
       Built-in presets (listed first in this order; they cannot be deleted, and a saved preset of the same name wins).
       Written like saved presets; lengths may use ...Ratio (of the first target's glyph height) instead of ...Pt.
       customColor, strokeCustomColor and textCustomColor are { rgb, cmyk }, picked by the document color mode */
    var BUILTIN_PRESETS = [
        {
            name: { ja: "アンダーライン", en: "Underline" },
            settings: {
                lineStyle: "underline", stackOrder: "stackBack", strokeColorChoice: "textColor", strokeWidthPt: 0.65, dashStyle: "solidLine",
                strokeCap: "roundCap", opacity: 100, verticalOffsetPt: -1.134, roughen: false, alignLineStarts: false, equalLineLengths: false,
                leadingBased: false, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "ドットのアンダーライン", en: "Dotted Underline" },
            settings: {
                lineStyle: "underline", stackOrder: "stackBack", strokeColorChoice: "textColor", strokeWidthPt: 1, dashStyle: "dottedLine", dashGapPt: 2,
                strokeCap: "roundCap", opacity: 100, verticalOffsetPt: -1.701, roughen: false, alignLineStarts: false, equalLineLengths: false,
                leadingBased: false, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "黄色のマーカー", en: "Yellow Marker" },
            settings: {
                lineStyle: "marker", stackOrder: "stackBack", strokeColorChoice: "yellow", strokeWidthPt: 5, dashStyle: "solidLine", strokeCap: "roundCap",
                opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false, leadingBased: false,
                kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "黄色のマーカー（手書き風）", en: "Yellow Marker (Hand-Drawn)" },
            settings: {
                lineStyle: "marker", stackOrder: "stackBack", strokeColorChoice: "yellow", strokeWidthPt: 5, dashStyle: "solidLine", strokeCap: "roundCap",
                opacity: 100, verticalOffsetPt: 0, roughen: true, roughenSizePercent: 1, zigzagRidges: 10, alignLineStarts: false, equalLineLengths: false,
                leadingBased: false, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "斜めの蛍光ペン", en: "Slanted Highlighter" },
            settings: {
                lineStyle: "straightShear", stackOrder: "stackBack", strokeColorChoice: "yellow", strokeWidthPt: 5.63, dashStyle: "solidLine",
                strokeCap: "buttCap", shearAngle: 30, opacity: 100, verticalOffsetPt: 5.669, roughen: false, alignLineStarts: false, equalLineLengths: false,
                leadingBased: true, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "波線", en: "Wavy Line" },
            settings: {
                lineStyle: "zigzag", stackOrder: "stackBack", strokeColorChoice: "textColor", strokeWidthPt: 0.45, dashStyle: "solidLine", strokeCap: "roundCap",
                zigzagSizePt: 0.68, zigzagRidges: 12, zigzagPoints: "smoothPoints", opacity: 100, verticalOffsetPt: -2.835, roughen: false,
                alignLineStarts: false, equalLineLengths: false, leadingBased: false, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "",
                kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "ギザギザ", en: "Zigzag" },
            settings: {
                lineStyle: "zigzag", stackOrder: "stackBack", strokeColorChoice: "red", strokeWidthPt: 1, dashStyle: "solidLine", strokeCap: "buttCap",
                zigzagSizePt: 0.68, zigzagRidges: 30, zigzagPoints: "cornerPoints", opacity: 100, verticalOffsetPt: 1.984, roughen: false,
                alignLineStarts: false, equalLineLengths: false, leadingBased: true, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "",
                kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "文字色の取り消し線", en: "Text-Color Strikethrough" },
            settings: {
                lineStyle: "strikethrough", stackOrder: "stackBack", strokeColorChoice: "textColor", strokeWidthRatio: 0.18, dashStyle: "solidLine",
                strokeCap: "buttCap", opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false, leadingBased: false,
                kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "二重取り消し線", en: "Double Strikethrough" },
            settings: {
                lineStyle: "doubleStrikethrough", stackOrder: "stackBack", strokeColorChoice: "textColor", strokeWidthRatio: 0.05, dashStyle: "solidLine",
                strokeCap: "buttCap", lineGapRatio: 0.2, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false,
                leadingBased: false, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "角丸の長方形", en: "Rounded Box" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: false, boxRotation: 0, colorChoice: "custom",
                customColor: { rgb: [79, 155, 216], cmyk: [65, 25, 0, 0] }, boxPaddingPt: 4.819, boxPaddingVPt: -1.701, boxPaddingLinked: false,
                cornerRadiusPt: 5.669, squareLineJoins: true, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: true, equalLineLengths: false,
                leadingBased: true, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "white"
            }
        },
        {
            name: { ja: "行ごとに帯", en: "Band per Line" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: false, boxRotation: 0, colorChoice: "custom",
                customColor: { rgb: [79, 155, 216], cmyk: [65, 25, 0, 0] }, boxPaddingPt: 3.969, boxPaddingVPt: -0.85, boxPaddingLinked: false,
                cornerRadiusPt: 0, squareLineJoins: true, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: true, equalLineLengths: false,
                leadingBased: true, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "white"
            }
        },
        {
            name: { ja: "行ごとに帯（隣接）", en: "Band per Line (Touching)" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: false, boxRotation: 0, colorChoice: "custom",
                customColor: { rgb: [79, 155, 216], cmyk: [65, 25, 0, 0] }, boxPaddingPt: 3.969, boxPaddingVPt: 0, boxPaddingLinked: false, cornerRadiusPt: 0,
                squareLineJoins: false, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: true, equalLineLengths: false, leadingBased: true,
                kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "white"
            }
        },
        {
            name: { ja: "行ごとに帯（手書き風）", en: "Band per Line (Hand-Drawn)" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: false, boxRotation: 0, colorChoice: "custom",
                customColor: { rgb: [79, 155, 216], cmyk: [65, 25, 0, 0] }, boxPaddingPt: 3.969, boxPaddingVPt: -0.85, boxPaddingLinked: false,
                cornerRadiusPt: 0, squareLineJoins: true, opacity: 100, verticalOffsetPt: 0, roughen: true, roughenSizePercent: 0.7, zigzagRidges: 3,
                alignLineStarts: true, equalLineLengths: false, leadingBased: true, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "",
                kerningAfterText: "", textColorChoice: "white"
            }
        },
        {
            name: { ja: "囲み罫", en: "Frame" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: false, shapeStroke: true, isPerChar: false, boxRotation: 0, strokeColorChoice: "textColor",
                strokeWidthPt: 0.45, dashStyle: "solidLine", boxPaddingPt: 3.402, boxPaddingVPt: -2.551, boxPaddingLinked: false, cornerRadiusPt: 0,
                squareLineJoins: true, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false, leadingBased: true,
                kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "囲み罫（角丸）", en: "Rounded Frame" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: false, shapeStroke: true, isPerChar: false, boxRotation: 0, strokeColorChoice: "textColor",
                strokeWidthPt: 0.45, dashStyle: "solidLine", boxPaddingPt: 3.402, boxPaddingVPt: -2.551, boxPaddingLinked: false, cornerRadiusPt: 2.835,
                squareLineJoins: true, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false, leadingBased: true,
                kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "囲み罫（行ごと）", en: "Frame per Line" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: false, shapeStroke: true, isPerChar: false, boxRotation: 0, strokeColorChoice: "textColor",
                strokeWidthPt: 0.45, dashStyle: "solidLine", boxPaddingPt: 3.402, boxPaddingVPt: -2.551, boxPaddingLinked: false, cornerRadiusPt: 0,
                squareLineJoins: false, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: true, equalLineLengths: false, leadingBased: true,
                kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "囲み罫（行ごと）手書き風", en: "Frame per Line (Hand-Drawn)" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: false, shapeStroke: true, isPerChar: false, boxRotation: 0, strokeColorChoice: "textColor",
                strokeWidthPt: 0.45, dashStyle: "solidLine", boxPaddingPt: 3.402, boxPaddingVPt: -2.551, boxPaddingLinked: false, cornerRadiusPt: 0,
                squareLineJoins: false, opacity: 100, verticalOffsetPt: 0, roughen: true, roughenSizePercent: 0.4, zigzagRidges: 10, alignLineStarts: true,
                equalLineLengths: false, leadingBased: true, kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "",
                textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "墨消し", en: "Redaction" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: false, boxRotation: 0, colorChoice: "textColor",
                boxPaddingPt: 1.984, boxPaddingVPt: -0.567, boxPaddingLinked: false, cornerRadiusPt: 0, squareLineJoins: true, opacity: 80, verticalOffsetPt: 0,
                roughen: false, alignLineStarts: true, equalLineLengths: false, leadingBased: true, kerningChoice: "keepKerning", trackingText: "",
                kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
        },
        {
            name: { ja: "1文字ずつの長方形", en: "Box per Character" },
            settings: {
                lineStyle: "box", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: true, boxRotation: 0, colorChoice: "custom",
                customColor: { rgb: [214, 96, 112], cmyk: [0, 70, 40, 10] }, boxPaddingPt: 0.567, boxPaddingVPt: 0.567, boxPaddingLinked: true,
                cornerRadiusPt: 0, squareLineJoins: true, opacity: 100, verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false,
                leadingBased: false, kerningChoice: "zeroKerning", trackingText: "170", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "white"
            }
        },
        {
            name: { ja: "1文字ずつの円", en: "Circle per Character" },
            settings: {
                lineStyle: "circle", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: true, boxRotation: 0, colorChoice: "custom",
                customColor: { rgb: [214, 96, 112], cmyk: [0, 70, 40, 10] }, boxPaddingPt: 1.106, boxPaddingVPt: 1.106, boxPaddingLinked: true, opacity: 100,
                verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false, leadingBased: false, kerningChoice: "zeroKerning",
                trackingText: "170", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "white"
            }
        },
        {
            name: { ja: "1文字ずつの六角形", en: "Hexagon per Character" },
            settings: {
                lineStyle: "hexagon", stackOrder: "stackBack", shapeFill: true, shapeStroke: false, isPerChar: true, boxRotation: 30, colorChoice: "custom",
                customColor: { rgb: [79, 155, 216], cmyk: [65, 25, 0, 0] }, boxPaddingPt: 0.652, boxPaddingVPt: 0.652, boxPaddingLinked: true, opacity: 100,
                verticalOffsetPt: 0, roughen: false, alignLineStarts: false, equalLineLengths: false, leadingBased: false, kerningChoice: "zeroKerning",
                trackingText: "170", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "white"
            }
        },
        {
            name: { ja: "傍点", en: "Emphasis Dots" },
            settings: {
                lineStyle: "emphasisDots", stackOrder: "stackBack", colorChoice: "textColor", dotSizePt: 2.041, boxPaddingPt: 0.907, boxPaddingVPt: 0.907,
                boxPaddingLinked: true, opacity: 100, verticalOffsetPt: -3.969, alignLineStarts: false, equalLineLengths: false, leadingBased: true,
                kerningChoice: "keepKerning", trackingText: "", kerningBeforeText: "", kerningAfterText: "", textColorChoice: "keepTextColor"
            }
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
    var PRESET_LIST_SIZE = [170, 240];     /* プリセットのリストの大きさ（高さは［設定］パネルに合わせて伸びる） / preset list size (grows to the Settings panel height) */
    var LINE_STYLE_DROPDOWN_WIDTH = 170;   /* ［種類］のドロップダウンの幅 / Style dropdown width */
    var PRESET_ICON_SIZE = [20, 18];       /* プリセットの保存・削除アイコンの大きさ / preset save and delete icon size */
    var PRESET_NAME_CHARS = 20;            /* プリセット名の入力欄の文字数 / preset name field width */
    var STACKED_FIELD_LABEL_SPACING = 6;   /* 項目名を上に置いた欄の、項目名と入力欄の間隔 / gap between a label above a field and the field */

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px。字の高さはこの6倍）/ size of one dot (px); glyphs are six dots tall */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：ドットフォント Chava の「v」「.」「0〜9」を 6 行のドット（# が点）に写したもの（「v」は5ドット幅・4段に描き直したもの）
       / Glyphs: the dots of "v", "." and "0"-"9" from the Chava pixel font, six rows each ("#" is a dot) */
    var VERSION_FOOTER_GLYPHS = {
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", "..#.."],
        ".": [".", ".", ".", ".", ".", "#"],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", ".###."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", "###"],
        "2": [".###.", "#...#", "....#", ".###.", "#....", "#####"],
        "3": [".###.", "#...#", "..##.", "....#", "#...#", ".###."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#."],
        "5": ["#####", "#....", "####.", "....#", "#...#", ".###."],
        "6": [".###.", "#....", "####.", "#...#", "#...#", ".###."],
        "7": ["#####", "....#", "...#.", "..#..", ".#...", ".#..."],
        "8": [".###.", "#...#", ".###.", "#...#", "#...#", ".###."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", ".###."]
    };

    /**
     * ウィンドウの下端にバージョン（SCRIPT_VERSION）を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        var versionText = String(SCRIPT_VERSION);
        var versionCanvas = footerGroup.add("group");
        versionCanvas.preferredSize = [measureVersionFooterDots(versionText) * VERSION_FOOTER_PIXEL_SIZE, 6 * VERSION_FOOTER_PIXEL_SIZE];
        versionCanvas.helpTip = versionText;
        versionCanvas.onDraw = function () {
            drawVersionFooterDots(this.graphics, versionText);
        };
        targetWindow.versionFooter = versionCanvas;
        return versionCanvas;
    }

    /**
     * バージョンの文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - バージョンの文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * バージョンの文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - バージョンの文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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

    /**
     * 上書き保存アイコン（フロッピーディスク）を描く。800×800 の絵。右上の欠けた角は細い長方形を斜めに並べる
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - アイコンの幅
     * @param {number} iconHeight - アイコンの高さ
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawOverwriteIcon(iconGraphics, iconWidth, iconHeight, iconColor) {
        var designRects = [
            [0, 0, 620, 70], [0, 70, 70, 800], [730, 180, 800, 800], [70, 730, 730, 800], /* 外枠 / outer frame */
            [200, 70, 540, 240],                                                          /* シャッター / shutter */
            [150, 420, 650, 490], [150, 490, 220, 730], [580, 490, 650, 730]              /* ラベル / label */
        ];
        var sliceCount = 12;
        /* 右上の斜めの辺（(620, 0) から (800, 180)）。線の太さ 70 は横幅で約 99 / the bevelled corner */
        for (var k = 0; k < sliceCount; k++) {
            var sliceTop = 180 * k / sliceCount;
            var edgeX = 620 + sliceTop + 180 / sliceCount / 2;
            designRects.push([edgeX - 99, sliceTop, edgeX, sliceTop + 180 / sliceCount]);
        }
        fillIconButtonRects(iconGraphics, iconWidth, iconHeight, [800, 800], designRects, iconColor);
    }

    // アイコンのボタン（再利用パーツ）ここまで / End of the reusable icon buttons

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    var LINK_UI_DARK = isDarkUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @param {number[]} [iconSize] - アイコンの [幅, 高さ]（省略時は LINK_ICON_SIZE。絵は 22px 基準から拡大縮小する）
     * @param {number} [chainRatio] - 鎖の絵の大きさの比率（省略時は 1。枠・地の大きさは変えず、鎖だけ縮める）
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle, iconSize, chainRatio) {
        var toggleSize = iconSize || LINK_ICON_SIZE;
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = toggleSize;
        linkToggle.minimumSize = toggleSize;
        linkToggle.maximumSize = toggleSize;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = toggleSize[0];
            var iconHeight = toggleSize[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR, chainRatio);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @param {number} [chainRatio] - 鎖の大きさの比率（省略時は 1）。中央に置いたまま縮める
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

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
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "Presets", "persistent", {
        /* 旧名（DrawMarkerBehindText）で保存したプリセットを読み継ぐ / Carry over presets saved under the former name */
        legacy: function () {
            return readSettingsLegacyFile(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/DrawMarkerBehindTextPresets.json");
        }
    });

    /* 前回ダイアログボックスを閉じたときの状態（設定・選んでいたプリセット）。Illustrator の終了まで残す（#targetengine が必須）
       State at the last close of the dialog box (settings and chosen preset), kept until Illustrator quits (needs #targetengine) */
    var dialogSessionStore = createSettingsStore(SCRIPT_NAME, "session");

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
            title:      { ja: "文字に線・図形を描く", en: "Draw Lines and Shapes on Text" },
            presetSave: { ja: "プリセットを保存", en: "Save Preset" }
        },
        panel: {
            preset:     { ja: "プリセット", en: "Presets" },
            settings:   { ja: "設定", en: "Settings" },
            textFormat: { ja: "文字", en: "Text" }
        },
        dropdown: {
            underline:           { ja: "アンダーライン", en: "Underline" },
            marker:              { ja: "マーカー", en: "Marker" },
            straightShear:       { ja: "直線＋シアー", en: "Straight + Shear" },
            zigzag:              { ja: "ギザギザ", en: "Zigzag" },
            fill:                { ja: "塗りつぶし", en: "Fill" },
            strikethrough:       { ja: "取り消し線", en: "Strikethrough" },
            doubleStrikethrough: { ja: "二重取り消し線", en: "Double Strikethrough" },
            box:                 { ja: "長方形", en: "Box" },
            circle:              { ja: "円", en: "Circle" },
            hexagon:             { ja: "六角形", en: "Hexagon" },
            emphasisDots:        { ja: "傍点", en: "Emphasis Dots" }
        },
        fieldLabel: {
            lineStyle:      { ja: "種類", en: "Style" },
            presetName:     { ja: "プリセット名", en: "Preset Name" },
            fillColor:      { ja: "塗り", en: "Fill" },
            strokeColor:    { ja: "線", en: "Stroke" },
            opacity:        { ja: "不透明度", en: "Opacity" },
            strokeWidth:    { ja: "線幅", en: "Weight" },
            verticalOffset: { ja: "位置調整", en: "Offset" },
            dashStyle:      { ja: "線の種類", en: "Line Type" },
            dashSegments:   { ja: "分割数", en: "Segments" },
            dashGap:        { ja: "間隔", en: "Gap" },
            strokeCap:      { ja: "線端", en: "Cap" },
            stackOrder:     { ja: "重ね順", en: "Stacking Order" },
            shearAngle:     { ja: "角度", en: "Angle" },
            zigzagSize:     { ja: "大きさ", en: "Size" },
            zigzagRidges:   { ja: "折り返し", en: "Ridges" },
            zigzagPoints:   { ja: "ポイント", en: "Points" },
            boxPadding:     { ja: "余白（横）", en: "Padding (H)" },
            boxPaddingV:    { ja: "余白（縦）", en: "Padding (V)" },
            cornerRadius:   { ja: "角丸の半径", en: "Corner Radius" },
            boxRotation:    { ja: "回転", en: "Rotate" },
            lineGap:        { ja: "線の間隔", en: "Line Gap" },
            dotSize:        { ja: "点の大きさ", en: "Dot Size" },
            kerning:        { ja: "カーニング", en: "Kerning" },
            tracking:       { ja: "トラッキング", en: "Tracking" },
            kerningBefore:  { ja: "カーニング（前）", en: "Kerning (Before)" },
            kerningAfter:   { ja: "カーニング（後）", en: "Kerning (After)" },
            lineLength:     { ja: "行長", en: "Line Length" },
            lineHeight:     { ja: "行の高さ", en: "Line Height" },
            roughen:        { ja: "ラフ", en: "Roughen" },
            roughenSize:    { ja: "サイズ", en: "Size" },
            drawUnit:       { ja: "単位", en: "Unit" },
            textFillColor:  { ja: "テキストカラー", en: "Text Color" },
            lineJoins:      { ja: "連続処理", en: "Continuity" }
        },
        checkbox: {
            shapeFill:   { ja: "描画", en: "Draw" },
            shapeStroke: { ja: "描画", en: "Draw" },
            squareLineJoins: { ja: "複数行", en: "Multi-Line" },
            alignLineStarts: { ja: "行頭", en: "Align Starts" },
            equalLineLengths: { ja: "行頭行末", en: "Align Both Ends" },
            leadingBased: { ja: "行送り基準", en: "Leading-Based" },
            roughen:      { ja: "手書き風", en: "Hand-Drawn" },
            kerningFromPadding: { ja: "余白に合わせてカーニングを自動調整", en: "Auto-Adjust Kerning to Padding" }
        },
        radio: {
            yellow:         { ja: "黄色", en: "Yellow" },
            red:            { ja: "赤", en: "Red" },
            white:          { ja: "白", en: "White" },
            textColor:      { ja: "文字色", en: "Text Color" },
            keepTextColor:  { ja: "変更しない", en: "No Change" },
            custom:         { ja: "カスタム", en: "Custom" },
            solidLine:      { ja: "直線", en: "Solid" },
            dashedLine:     { ja: "破線", en: "Dashed" },
            dottedLine:     { ja: "ドット", en: "Dotted" },
            buttCap:        { ja: "なし", en: "Butt" },
            roundCap:       { ja: "丸型", en: "Round" },
            stackFront:     { ja: "前", en: "Front" },
            stackBack:      { ja: "後", en: "Back" },
            perLine:        { ja: "行", en: "Line" },
            perChar:        { ja: "1文字ずつ", en: "Per Character" },
            smoothPoints:   { ja: "滑らか", en: "Smooth" },
            cornerPoints:   { ja: "直線的", en: "Corner" },
            keepKerning:    { ja: "変更しない", en: "No Change" },
            zeroKerning:    { ja: "0", en: "0" },
            metricsKerning: { ja: "自動", en: "Auto" },
            opticalKerning: { ja: "オプティカル", en: "Optical" }
        },
        tooltip: {
            /* 種類（ドロップダウンの helpTip。選んでいる種類の説明） / Styles (the dropdown's helpTip follows the chosen style) */
            underline:     { ja: "字面のすぐ下に、線の上端を合わせて引きます", en: "Draws the line just below the glyphs, its top edge on their bottom" },
            marker:        { ja: "字面の下端を線の中心にして引きます", en: "Draws the line centered on the bottom edge of the glyphs" },
            straightShear: { ja: "線に［パスの自由変形］効果を掛けて、両端を斜めにします", en: "Applies the Free Distort effect to slant both ends of the line" },
            zigzag:        { ja: "線に［ラフ］効果でポイントを足してから、［ジグザグ］効果を掛けます", en: "Adds points to the line with the Roughen effect, then applies the Zig Zag effect" },
            fill:          { ja: "字面の範囲を塗りつぶします（線幅は字面の高さ）", en: "Fills the glyph area (the line width is the glyph height)" },
            strikethrough: { ja: "字面の上下中央に線を引きます", en: "Draws the line through the vertical center of the glyphs" },
            doubleStrikethrough: {
                ja: "字面の上下中央に2本の線を引きます（間隔は［線の間隔］）",
                en: "Draws two lines through the vertical center of the glyphs (spaced by Line Gap)"
            },
            box: {
                ja: "字面を囲む長方形を、行（縦組みは列）ごとに1つ描きます。［1文字ずつ］なら文字ごとの長方形です",
                en: "Draws one box around the glyphs of each line or column; with Per Character, a box per character"
            },
            circle: {
                ja: "字面を囲む楕円を、行（縦組みは列）ごとに1つ描きます。［1文字ずつ］なら文字ごとの円です",
                en: "Draws one ellipse around the glyphs of each line or column; with Per Character, a circle per character"
            },
            hexagon: {
                ja: "字面を囲む、左右（縦組みは上下）がとがった六角形を、行ごとに1つ描きます。［1文字ずつ］なら文字ごとの正六角形です",
                en: "Draws one hexagon pointed at the sides (top and bottom in vertical text) around each line; with Per Character, a regular hexagon per character"
            },
            emphasisDots: {
                ja: "1文字ずつ、字の上（縦組みは右）に丸い点を置きます。句読点・括弧・スペースには付けません",
                en: "Places a round dot above each character (right of it in vertical text); skips punctuation, brackets and spaces"
            },
            /* 設定 / Settings */
            perLine: { ja: "［長方形］［円］［六角形］を行（縦組みは列）ごとに1つ描きます", en: "Draws one Box, Circle or Hexagon per line (column in vertical text)" },
            perChar: {
                ja: "［長方形］［円］［六角形］を1文字ずつ描きます（高さは行の字面にそろえ、幅は高さ以上。円・六角形は長い辺に合わせます）。句読点・括弧・スペースは飛ばします",
                en: "Draws Box, Circle or Hexagon per character (as tall as the line's glyphs, at least as wide as tall; Circle and Hexagon fit the longer side); skips punctuation, brackets and spaces"
            },
            swatch:    { ja: "クリックしてカラーピッカーで変更", en: "Click to change in the Color Picker" },
            textColor: { ja: "対象の先頭の文字の塗りを使います（テキストオブジェクトごと）", en: "Uses the fill of the first target character (per text object)" },
            opacity:   { ja: "線・図形の不透明度です", en: "Opacity of the lines and shapes" },
            verticalOffset: {
                ja: "種類ごとの位置から、線や図形を上下にずらす量です。プラスで上、マイナスで下へずれます（縦組みは左右。プラスで左）",
                en: "Moves the line or shape up or down from the style's position; positive moves it up, negative down (left or right in vertical text; positive is left)"
            },
            dashedLine: {
                ja: "［分割数］の数の線分を［間隔］をあけて並べ、線の両端を線分で終わらせます",
                en: "Places Segments dashes spaced by Gap, so both ends finish with a dash"
            },
            dashSegments: {
                ja: "線ごとの破線の線分の数です。線分の長さは、線の両端が線分で終わるように決めます。空欄なら線幅の約2倍の線分になる数にします",
                en: "Number of dashes per line; their length makes both ends finish with a dash. Leave blank for dashes about twice the line width"
            },
            dashGap: { ja: "破線の線分のあいだ、ドットの点のあいだの長さです", en: "Space between dashes or between dots" },
            dottedLine: {
                ja: "線幅の大きさの丸い点を、点のあいだが［間隔］に近くなるように並べます。線の両端に点が来ます（線端は丸型になります）",
                en: "Round dots as large as the line width, spaced close to Gap, with dots at both ends (the cap becomes round)"
            },
            roundCap:    { ja: "線の両端が線幅の半分ずつ外へ伸びます", en: "Each end extends by half the line width" },
            stackFront:  { ja: "文字の前面に描きます", en: "Draws in front of the text" },
            stackBack:   { ja: "文字の背面に描きます", en: "Draws behind the text" },
            textFillColor: {
                ja: "線や図形を描いたあと、対象の文字の塗りをこのカラーにします",
                en: "Sets the fill of the target characters to this color after drawing"
            },
            shapeFill: { ja: "［長方形］［円］［六角形］の図形を、この行のカラーで塗ります", en: "Fills Box, Circle and Hexagon shapes with the color on this row" },
            shapeStroke: {
                ja: "［長方形］［円］［六角形］の図形に、この行のカラーで線を付けます（［線幅］［線の種類］を使います）",
                en: "Strokes Box, Circle and Hexagon shapes with the color on this row (uses Weight and Line Type)"
            },
            shearAngle: {
                ja: "線の両端を傾ける角度です（-80〜80°）。プラスで右上がり、マイナスで左上がりになります",
                en: "Angle of the slanted ends (-80 to 80°); positive slants to the upper right, negative to the upper left"
            },
            zigzagSize:   { ja: "ギザギザの山の高さです", en: "Height of the zigzag ridges" },
            zigzagRidges: {
                ja: "線の長さ1インチ（25.4 mm）あたりの折り返しの数です。［手書き風］では［ラフ］効果の詳細にも使います",
                en: "Number of ridges per inch (25.4 mm) of the line; also the Roughen detail for Hand-Drawn"
            },
            smoothPoints: { ja: "折り返しを曲線にして、波線にします", en: "Curves the ridges into a wavy line" },
            cornerPoints: { ja: "折り返しを角にして、ギザギザにします", en: "Keeps the ridges as sharp corners" },
            roughen: {
                ja: "線や図形に［ラフ］効果を掛けて、縁を手書きのように揺らします（ポイントはギザギザ。［傍点］には掛けません）",
                en: "Applies the Roughen effect to the lines and shapes so their edges wobble like hand drawing (corner points; not for Emphasis Dots)"
            },
            roughenSize:   { ja: "［ラフ］効果のサイズです（線や図形の大きさに対する％）", en: "Size of the Roughen effect (% of the line or shape)" },
            boxPadding: {
                ja: "［長方形］［円］［六角形］は字面のまわりの余白（左右・上下）、［傍点］は字面から点までの距離です（横組みは縦、縦組みは横）。マイナスで内側へ寄せます",
                en: "Space around the glyphs for Box, Circle and Hexagon (left/right and top/bottom); distance from the glyphs to the dots for Emphasis Dots (V in horizontal text, H in vertical text). Negative values pull inward"
            },
            boxPaddingLink: { ja: "［余白（横）］と［余白（縦）］を同じ値にします", en: "Keeps Padding (H) and Padding (V) the same" },
            cornerRadius: { ja: "［長方形］の角丸の半径です。0 で角丸なし", en: "Corner radius of Box; 0 for square corners" },
            alignLineStarts: {
                ja: "複数行のとき、線や図形の始まりを、いちばん前から始まる行にそろえます（縦組みは列の上端）。行末はそろえません",
                en: "For multi-line text, starts the lines and shapes where the earliest line starts (the top of the column in vertical text); the ends stay as they are"
            },
            equalLineLengths: {
                ja: "複数行のとき、文字数が少ない行の線や図形も、いちばん長い行の長さにそろえます（縦組みは列の長さ）",
                en: "For multi-line text, extends the lines and shapes of short lines to the length of the longest one (column length in vertical text)"
            },
            leadingBased: {
                ja: "線・図形・傍点の位置と大きさを、行の字面の高さ（縦組みは幅）ではなく行送りで決めます。どの行も同じ高さになり、行の中央にそろいます",
                en: "Positions and sizes lines, shapes and dots by the leading instead of the glyph height of each line (width in vertical text), so every line gets the same height, centered on the line"
            },
            squareLineJoins: {
                ja: "複数行にわたる［長方形］を1続きに見せます。先頭の行の左端と最後の行の右端だけを角丸にし、行の継ぎ目は角にして、［線］は継ぎ目の辺を開けます（縦組みは先頭の列の上端と最後の列の下端）",
                en: "Makes a Box spanning several lines look continuous: only the left end of the first line and the right end of the last are rounded, and the Stroke is left open where the lines continue (the top of the first column and the bottom of the last in vertical text)"
            },
            boxRotation: {
                ja: "［長方形］［円］［六角形］を、1つずつその中心で回転します（-360〜360°、プラスで反時計回り）",
                en: "Rotates each Box, Circle or Hexagon about its own center (-360 to 360°; positive is counterclockwise)"
            },
            lineGap: { ja: "［二重取り消し線］の2本の線のあいだの間隔です", en: "Space between the two lines of Double Strikethrough" },
            dotSize: { ja: "［傍点］の点の直径です", en: "Diameter of the emphasis dots" },
            /* 文字 / Text */
            zeroKerning: {
                ja: "自動カーニングを切り、手動のカーニングも 0 にします。プロポーショナルメトリクスはオフにします",
                en: "Turns auto kerning off and sets manual kerning to 0; turns Proportional Metrics off"
            },
            metricsKerning: { ja: "フォントのペアカーニングの情報（メトリクス）でカーニングします", en: "Kerns with the font's pair kerning (metrics)" },
            opticalKerning: { ja: "字形の形でカーニングします。プロポーショナルメトリクスはオフにします", en: "Kerns by the glyph shapes; turns Proportional Metrics off" },
            tracking: {
                ja: "対象の文字のトラッキングです（1/1000 em）。空欄なら変更しません",
                en: "Tracking of the target characters (1/1000 em); leave blank to keep it"
            },
            kerningBefore: {
                ja: "対象の先頭の文字と、その前の文字のあいだのカーニングです（1/1000 em）。空欄なら変更しません",
                en: "Kerning between the first target character and the one before it (1/1000 em); leave blank to keep it"
            },
            kerningAfter: {
                ja: "対象の最後の文字と、その次の文字のあいだのカーニングです（1/1000 em）。次の文字が無ければ使いません。空欄なら変更しません",
                en: "Kerning between the last target character and the next one (1/1000 em); unused when there is no next character. Leave blank to keep it"
            },
            kerningFromPadding: {
                ja: "［長方形］［円］［六角形］を行ごとに描くとき、カーニング（前・後）を［余白（横）］÷ 文字サイズ × 1000 にして、図形のぶん前後をあけます（縦組みは［余白（縦）］。文字サイズは対象の先頭と最後の文字）。欄には先頭の対象の値を出し、欄を変えると計算した値との差を微調整として足します",
                en: "For Box, Circle and Hexagon per line, sets Kerning (Before) / (After) to Padding (H) / type size x 1000 to make room for the shape (Padding (V) in vertical text; the size of the first and last target characters). The fields show the first target's values; editing them adds the difference as a fine adjustment"
            },
            /* プリセット / Presets */
            preset:     { ja: "クリックすると、その設定を読み込みます", en: "Click to load the settings" },
            presetSave: { ja: "今の設定に名前を付けて保存します", en: "Saves the current settings under a name" },
            presetOverwrite: { ja: "「{name}」に今の設定を上書き保存します", en: "Overwrites \"{name}\" with the current settings" },
            presetOverwriteNone: {
                ja: "選んでいるプリセット（設定を変える前に選んでいたもの）に、今の設定を上書き保存します",
                en: "Overwrites the selected preset (or the one selected before your changes) with the current settings"
            },
            presetDelete: {
                ja: "選んでいるプリセットを削除します（最初からあるものは削除できません）",
                en: "Deletes the selected preset (built-in presets cannot be deleted)"
            },
            presetName: { ja: "保存する設定の名前です。同じ名前は上書きします", en: "Name the settings are saved under. The same name is overwritten" },
            /* ボタン行・∧∨ / Button row and steppers */
            reset: {
                ja: "ダイアログボックスを開いたときの設定と、選んでいたプリセットに戻します。カーニング（前・後）とトラッキングは空欄にし、［余白に合わせてカーニングを自動調整］はオフにします（保存したプリセットはそのまま）",
                en: "Returns to the settings and the preset chosen when the dialog box opened, with Kerning (Before) / (After) and Tracking cleared and Auto-Adjust Kerning to Padding off (saved presets stay as they are)"
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
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset:  { ja: "リセット", en: "Reset" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextSelection: {
                ja: "文字ツールで文字を選択するか、テキストオブジェクトを選択してください。",
                en: "Select characters with the Type tool, or select text objects."
            },
            noGlyphs: {
                ja: "選択した文字に字形がありません。\nスペースや改行だけの範囲には線や図形を描けません。",
                en: "The selected characters have no glyphs.\nNothing can be drawn over spaces or line breaks only."
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
        return (doc.documentColorSpace === DocumentColorSpace.CMYK) ? createCmykColor(cmykValues) : createRgbColor(rgbValues);
    }

    /**
     * RGB の色を作る
     * @param {number[]} rgbValues - R, G, B（0〜255）
     * @returns {RGBColor} 色
     */
    function createRgbColor(rgbValues) {
        var rgbColor = new RGBColor();
        rgbColor.red = rgbValues[0];
        rgbColor.green = rgbValues[1];
        rgbColor.blue = rgbValues[2];
        return rgbColor;
    }

    /**
     * CMYK の色を作る
     * @param {number[]} cmykValues - C, M, Y, K（0〜100）
     * @returns {CMYKColor} 色
     */
    function createCmykColor(cmykValues) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cmykValues[0];
        cmykColor.magenta = cmykValues[1];
        cmykColor.yellow = cmykValues[2];
        cmykColor.black = cmykValues[3];
        return cmykColor;
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
            /* 親が無効なら薄く描く（自作描画は自動でディムにならない） / Draw faded when a parent is disabled (custom drawing is not dimmed automatically) */
            var swatchRgba = colorToScreenRgb(swatchColor).concat([isIconButtonEnabledInTree(colorSwatch) ? 1 : 0.3]);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, swatchRgba));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 1));
        };
        /* キャンセルは渡した色が返る。group には notify() が無いので、隠して再表示して描き直す
           Cancel returns the color passed in; a group has no notify(), so hide and show it to redraw */
        colorSwatch.addEventListener("click", function () {
            if (!isIconButtonEnabledInTree(colorSwatch)) return;
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
     * 項目名と、カラーのラジオボタン（［変更しない］［黄色］［赤］［白］［文字色］［カスタム］のうち choiceKeys のもの）を1行に並べた、カラーを選ぶ欄を作る。
     * 先頭に［カスタム］と、クリックでカラーピッカーを開く色見本を置く。
     * 選択肢に無い［黄色］［赤］［白］を書き込んだときは、その色を［カスタム］にして選ぶ
     * @param {Window|Panel|Group} parent - 追加先
     * @param {Document} doc - 対象ドキュメント（カラーモード）
     * @param {boolean} isTextColorUsable - ［文字色］を使えるか（false なら無効にする）
     * @param {Function} onColorChange - カラーを選び直したあとに呼ぶ関数
     * @param {string} labelKey - 項目名の LABELS.fieldLabel のキー
     * @param {string[]} choiceKeys - 並べるキー（"keepTextColor" / "yellow" / "red" / "white" / "textColor" / "custom"。"custom" はどこに書いても先頭に置く）
     * @param {string} defaultChoice - 最初に選んでおくキー
     * @param {Object} [drawOptionSpec] - 項目名とカラーのあいだに置くチェックボックス { checkboxKey, initialValue, onToggle }（addOptionCheckbox() に渡す）
     * @returns {{label: StaticText, drawOption: Object|null, getColorOptions: Function, getChoice: Function, setChoice: Function, getCustomColor: Function, setCustomColor: Function, setEnabled: Function}}
     *     項目名、チェックボックス（addOptionCheckbox() の戻り値。無ければ null）、{ color, useTextColor } を返す getColorOptions（［変更しない］なら color は null）、
     *     選んだキーの読み書き、［カスタム］の色の読み書き、カラーの有効／無効の切り替え
     */
    function addColorChoiceRow(parent, doc, isTextColorUsable, onColorChange, labelKey, choiceKeys, defaultChoice, drawOptionSpec) {
        var colorRow = parent.add("group");
        setupRow(colorRow, "left", 8);
        var colorLabel = colorRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        var drawOption = drawOptionSpec
            ? addOptionCheckbox(colorRow, drawOptionSpec.checkboxKey, drawOptionSpec.initialValue, drawOptionSpec.onToggle)
            : null;

        var choiceGroup = colorRow.add("group");
        setupRow(choiceGroup, "left", 8);

        var choiceColors = {};
        for (var presetKey in MARKER_COLOR_PRESETS) {
            if (!MARKER_COLOR_PRESETS.hasOwnProperty(presetKey)) continue;
            choiceColors[presetKey] = createDocumentColor(doc, MARKER_COLOR_PRESETS[presetKey].rgb, MARKER_COLOR_PRESETS[presetKey].cmyk);
        }
        /* ［文字色］が使えないとき・知らないキーの代わりに選ぶ、先頭の色 / The first plain color, used instead of an unusable Text Color or an unknown key */
        var fallbackKey = (choiceKeys[0] === "textColor") ? choiceKeys[1] : choiceKeys[0];
        var choiceRadios = {};

        /**
         * キーを、この欄で選べるキーに直す（［文字色］が使えなければ先頭の色、知らないキーも先頭の色）
         * @param {string} choiceKey - キー
         * @returns {string} 選べるキー
         */
        function toUsableKey(choiceKey) {
            var isListed = false;
            for (var i = 0; i < choiceKeys.length; i++) {
                if (choiceKeys[i] === choiceKey) isListed = true;
            }
            if (!isListed || (choiceKey === "textColor" && !isTextColorUsable)) return fallbackKey;
            return choiceKey;
        }
        /* show() 前に代入した value は読み戻せないので、選択は変数で持つ / value set before show() cannot be read back, so keep the choice in a variable */
        var chosenKey = toUsableKey(defaultChoice);

        /**
         * キーのラジオボタンだけをオンにする。［カスタム］は色見本と同じ行で親が別なので、排他は手で取る
         * @param {string} choiceKey - 選ぶキー
         * @returns {void}
         */
        function checkChoiceRadio(choiceKey) {
            chosenKey = choiceKey;
            for (var radioKey in choiceRadios) {
                if (choiceRadios.hasOwnProperty(radioKey)) choiceRadios[radioKey].value = (radioKey === choiceKey);
            }
        }

        /**
         * キーを選び、選び直したことを知らせる
         * @param {string} choiceKey - 選ぶキー
         * @returns {void}
         */
        function selectChoice(choiceKey) {
            checkChoiceRadio(choiceKey);
            onColorChange();
        }

        /**
         * カラーのラジオボタンを1つ足す
         * @param {Group} radioParent - 追加先
         * @param {string} choiceKey - キー（LABELS.radio のキー）
         * @returns {RadioButton} ラジオボタン
         */
        function addChoiceRadio(radioParent, choiceKey) {
            var choiceRadio = radioParent.add("radiobutton", undefined, getLabel("radio." + choiceKey));
            choiceRadio.value = (choiceKey === chosenKey);
            choiceRadio.onClick = function () { selectChoice(choiceKey); };
            choiceRadios[choiceKey] = choiceRadio;
            return choiceRadio;
        }

        /* ［カスタム］と色見本を先頭に置く / Custom and its swatch come first */
        var customRow = choiceGroup.add("group");
        setupRow(customRow, "left", 8);
        addChoiceRadio(customRow, "custom");
        /* 色見本で色を変えたら［カスタム］を選んだことにする / Changing the swatch also chooses Custom */
        var customSwatch = addColorSwatch(customRow, choiceColors.custom, function (pickedColor) {
            choiceColors.custom = pickedColor;
            selectChoice("custom");
        });

        for (var k = 0; k < choiceKeys.length; k++) {
            if (choiceKeys[k] === "custom") continue;
            var choiceRadio = addChoiceRadio(choiceGroup, choiceKeys[k]);
            if (choiceKeys[k] !== "textColor") continue;
            choiceRadio.helpTip = getLabel("tooltip.textColor");
            choiceRadio.enabled = isTextColorUsable;
        }

        return {
            label: colorLabel,
            drawOption: drawOption,
            /* ［文字色］の色は対象ごとに違うので、線を引くときに対象から取る / Text Color differs per target, so it is read from each target when drawing */
            getColorOptions: function () {
                if (chosenKey === "textColor") return { color: null, useTextColor: true };
                if (chosenKey === "keepTextColor") return { color: null, useTextColor: false };
                return { color: choiceColors[chosenKey], useTextColor: false };
            },
            getChoice: function () { return chosenKey; },
            /* 選択肢に無い［黄色］［赤］［白］は、その色を［カスタム］にして選ぶ。ほかの知らないキーと使えない［文字色］は先頭の色
               A preset color missing here becomes Custom; other unknown keys and an unusable Text Color fall back to the first color */
            setChoice: function (choiceKey) {
                if (!choiceRadios.hasOwnProperty(choiceKey) && choiceKey !== "custom" && MARKER_COLOR_PRESETS.hasOwnProperty(choiceKey)) {
                    choiceColors.custom = choiceColors[choiceKey];
                    customSwatch.setSwatchColor(choiceColors.custom);
                    checkChoiceRadio("custom");
                    return;
                }
                checkChoiceRadio(toUsableKey(choiceKey));
            },
            getCustomColor: function () { return choiceColors.custom; },
            setCustomColor: function (newColor) {
                choiceColors.custom = newColor;
                customSwatch.setSwatchColor(newColor);
            },
            /* カラーの選択肢をまとめて切り替え（項目名とチェックボックスはそのまま）、自作描画の色見本は描き直してディムにする
               Toggle the color choices (not the label or checkbox) and repaint the custom-drawn swatch */
            setEnabled: function (isEnabled) {
                if (choiceGroup.enabled === isEnabled) return;
                choiceGroup.enabled = isEnabled;
                customSwatch.hide();
                customSwatch.show();
            }
        };
    }

    /**
     * 項目名・∧∨・入力欄の、数値を入れる1行を作る
     * @param {Window|Panel|Group} parent - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー（LABELS.tooltip に同じキーがあれば入力欄の helpTip にする）
     * @param {Object} fieldOptions - addSteppedField() の text / characters / step / min / max / integer / unit。
     *     allowBlank が true なら空欄を残す（「変更しない」の意味に使う）
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
            if (fieldOptions.allowBlank && /^\s*$/.test(numberInput.text)) {
                numberInput.text = "";
                numberInput.lastValidText = "";
            } else {
                normalizeValue();
            }
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
     * プレビューの線・図形を出し入れするオブジェクトを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} markerTargets - collectMarkerTargets() の戻り値
     * @returns {{update: Function, clear: Function}} 描き直す update(markerOptions) と、消す clear()
     */
    function createMarkerPreview(doc, markerTargets) {
        var previewItems = [];
        /* 文字の書式を変えるときは、書式を変えた複製を前面に置いて元のテキストを隠して見せる（markerTargets と同じ順）
           Text formatting is previewed with a formatted copy in front while the original is hidden (same order as markerTargets) */
        var textCopies = [];

        /**
         * プレビューの線・図形と書式の複製を消し、隠した元のテキストを表示に戻す
         * @returns {void}
         */
        function clearPreviewItems() {
            for (var i = 0; i < previewItems.length; i++) previewItems[i].remove();
            previewItems = [];
            for (var j = 0; j < textCopies.length; j++) {
                textCopies[j].remove();
                markerTargets[j].sourceFrame.hidden = false;
            }
            textCopies = [];
        }

        return {
            update: function (markerOptions) {
                clearPreviewItems();
                /* 選択の反転表示やバウンディングボックスで線が見えにくいので、選択を外す（閉じたあと main() で戻す）
                   The selection highlight hides the line, so deselect it (main() restores it after closing) */
                doc.selection = null;
                if (!hasTextFormatChange(markerOptions)) {
                    previewItems = placeMarkers(markerTargets, markerOptions, null);
                    return;
                }
                /* 字面は書式を変えた複製で測る。複製は hidden を引き継ぐので、元を隠す前に作る
                   Measure the glyphs on the formatted copies; a duplicate inherits hidden, so copy before hiding */
                for (var i = 0; i < markerTargets.length; i++) {
                    var sourceFrame = markerTargets[i].sourceFrame;
                    var textCopy = sourceFrame.duplicate(sourceFrame, ElementPlacement.PLACEBEFORE);
                    applyTextFormat(textCopy, markerTargets[i], markerOptions);
                    textCopies.push(textCopy);
                }
                previewItems = placeMarkers(markerTargets, markerOptions, textCopies);
                for (var j = 0; j < markerTargets.length; j++) markerTargets[j].sourceFrame.hidden = true;
            },
            clear: clearPreviewItems
        };
    }

    /**
     * ［プリセット］パネル（プリセット名のリストと、保存・削除のアイコン）を作る
     * @param {Group} parent - 追加先（パネルを横に並べるグループ）
     * @returns {{presetList: ListBox, presetSaveIcon: Group, presetOverwriteIcon: Group, presetDeleteIcon: Group}} コントロール
     */
    function addPresetPanel(parent) {
        var presetPanel = parent.add("panel", undefined, getLabel("panel.preset"));
        setupPanel(presetPanel);
        presetPanel.alignChildren = ["fill", "fill"];
        var presetList = presetPanel.add("listbox", undefined, []);
        presetList.preferredSize = PRESET_LIST_SIZE;
        presetList.helpTip = getLabel("tooltip.preset");
        /* 左に削除、右に保存・上書き保存（あいだは伸びるスペーサー） / Delete on the left, Save and Overwrite on the right, with a stretching spacer between */
        var presetIconRow = presetPanel.add("group");
        setupRow(presetIconRow, ["fill", "bottom"], 8);
        var presetDeleteIcon = addIconButton(presetIconRow, PRESET_ICON_SIZE, drawTrashIcon);
        presetDeleteIcon.helpTip = getLabel("tooltip.presetDelete");
        var presetIconSpacer = presetIconRow.add("group");
        presetIconSpacer.alignment = ["fill", "fill"];
        presetIconSpacer.minimumSize.width = 0;
        var presetIconRightGroup = presetIconRow.add("group");
        setupRow(presetIconRightGroup, "right", 8);
        var presetSaveIcon = addIconButton(presetIconRightGroup, PRESET_ICON_SIZE, drawSaveIcon);
        presetSaveIcon.helpTip = getLabel("tooltip.presetSave");
        var presetOverwriteIcon = addIconButton(presetIconRightGroup, PRESET_ICON_SIZE, drawOverwriteIcon);
        presetOverwriteIcon.helpTip = getLabel("tooltip.presetOverwriteNone");
        return {
            presetList: presetList,
            presetSaveIcon: presetSaveIcon,
            presetOverwriteIcon: presetOverwriteIcon,
            presetDeleteIcon: presetDeleteIcon
        };
    }

    /**
     * ［設定］パネルと、その下の［文字］パネルを作る。［設定］は先頭に［種類］と、2つの列にまたがる［塗り］［線］の2行、
     * その下の左の列に線と種類ごとの効果（線幅・線の種類・角度など）、右の列に配置と大きさ（位置調整・余白・単位など）を置く
     * @param {Group} parent - 追加先（パネルを縦に並べるグループ）
     * @param {Document} doc - 対象ドキュメント
     * @param {number} glyphHeightPt - 線幅と余白の初期値のもとにする字面の高さ（pt）
     * @param {boolean} isTextColorUsable - ［文字色］を使えるか
     * @param {Function} onSettingChange - 設定を変えたあとに呼ぶ関数
     * @returns {Object} コントロールと単位 { lineStyleRow, shearAngleInput, zigzagSizeInput, zigzagRidgesInput, zigzagPointsRow, boxPaddingInput, cornerRadiusInput, boxRotationInput, lineGapInput, dotSizeInput, fillColorRow, strokeColorRow, opacityInput, strokeWidthInput, verticalOffsetInput, dashStyleRow, strokeCapRow, shapeFillOption, shapeStrokeOption, kerningRow, trackingInput, kerningBeforeInput, kerningAfterInput, textFillColorRow, perCharOption, rulerUnit, strokeUnit, labelColumns }
     *     （labelColumns は列ごとの項目名）
     */
    function addMarkerControls(parent, doc, glyphHeightPt, isTextColorUsable, onSettingChange) {
        var markerControls = { rulerUnit: getUnitInfo(), strokeUnit: getUnitInfo("strokeUnits") };
        /**
         * 種類や線の種類を変えたら、使える項目を切り替えてからプレビューする
         * @returns {void}
         */
        function onStyleChange() {
            updateStyleFieldsEnabled(markerControls);
            onSettingChange();
        }

        var settingsPanel = parent.add("panel", undefined, getLabel("panel.settings"));
        setupPanel(settingsPanel);
        addLineStyleDropdown(settingsPanel, markerControls, onStyleChange);
        /* ［重ね順］は［種類］の右に置く / Stacking Order sits right of Style */
        markerControls.stackOrderRow = addRadioChoiceRow(markerControls.lineStyleRow.row, "stackOrder", ["stackFront", "stackBack"], DEFAULT_STACK_ORDER, onSettingChange);
        markerControls.stackOrderRow.row.margins = [12, 0, 0, 0];
        markerControls.stackOrderRow.radios.stackFront.helpTip = getLabel("tooltip.stackFront");
        markerControls.stackOrderRow.radios.stackBack.helpTip = getLabel("tooltip.stackBack");
        /* ［塗り］［線］の2行は2つの列にまたがって置く。図形の塗り・線の［描画］は各行の先頭（切り替えたら使える項目も切り替える）
           Fill and Stroke span both columns; each row starts with Draw for shapes (toggling it updates the usable fields) */
        markerControls.fillColorRow = addColorChoiceRow(settingsPanel, doc, isTextColorUsable, onSettingChange, "fillColor", FILL_COLOR_CHOICES, DEFAULT_FILL_COLOR_CHOICE,
            { checkboxKey: "shapeFill", initialValue: DEFAULT_SHAPE_FILL, onToggle: onStyleChange });
        markerControls.strokeColorRow = addColorChoiceRow(settingsPanel, doc, isTextColorUsable, onSettingChange, "strokeColor", STROKE_COLOR_CHOICES, DEFAULT_STROKE_COLOR_CHOICE,
            { checkboxKey: "shapeStroke", initialValue: DEFAULT_SHAPE_STROKE, onToggle: onStyleChange });
        markerControls.shapeFillOption = markerControls.fillColorRow.drawOption;
        markerControls.shapeStrokeOption = markerControls.strokeColorRow.drawOption;
        var settingsColumnsGroup = settingsPanel.add("group");
        settingsColumnsGroup.orientation = "row";
        settingsColumnsGroup.alignChildren = ["fill", "top"];
        settingsColumnsGroup.spacing = COLUMN_SPACING;
        addStrokeSettingFields(addSettingsColumn(settingsColumnsGroup), markerControls, glyphHeightPt, onSettingChange, onStyleChange);
        addLayoutSettingFields(addSettingsColumn(settingsColumnsGroup), markerControls, glyphHeightPt, onSettingChange, onStyleChange);
        addTextFormatPanel(parent, markerControls, doc, onSettingChange, onStyleChange);

        /* 項目名の幅は列ごとにそろえる（列にまたがる行は左の列とそろえる） / Label widths are aligned per column (the spanning rows with the left column) */
        markerControls.labelColumns = [
            [
                markerControls.lineStyleRow.label, markerControls.fillColorRow.label, markerControls.strokeColorRow.label,
                markerControls.strokeWidthInput.fieldLabel, markerControls.dashStyleRow.label, markerControls.dashSegmentsInput.fieldLabel,
                markerControls.dashGapInput.fieldLabel, markerControls.strokeCapRow.label, markerControls.lineGapInput.fieldLabel,
                markerControls.shearAngleInput.fieldLabel, markerControls.zigzagSizeInput.fieldLabel, markerControls.zigzagRidgesInput.fieldLabel,
                markerControls.zigzagPointsRow.label, markerControls.roughenOption.label, markerControls.roughenSizeInput.fieldLabel
            ],
            [
                markerControls.opacityInput.fieldLabel, markerControls.verticalOffsetInput.fieldLabel,
                markerControls.boxPaddingInput.fieldLabel, markerControls.boxPaddingVInput.fieldLabel,
                markerControls.lineLengthLabel, markerControls.leadingBasedOption.label, markerControls.drawUnitRow.label,
                markerControls.cornerRadiusInput.fieldLabel, markerControls.squareLineJoinsOption.label, markerControls.boxRotationInput.fieldLabel,
                markerControls.dotSizeInput.fieldLabel
            ],
            [markerControls.textFillColorRow.label, markerControls.kerningRow.label]
        ];
        updateStyleFieldsEnabled(markerControls);
        return markerControls;
    }

    /**
     * ［設定］パネルの先頭に［種類］のドロップダウンを作り、markerControls.lineStyleRow に入れる。
     * ドロップダウンの helpTip は選んでいる種類の説明にする
     * @param {Panel} settingsPanel - ［設定］パネル
     * @param {Object} markerControls - 作ったコントロールを入れるオブジェクト
     * @param {Function} onStyleChange - 種類を選び直したあとに呼ぶ関数
     * @returns {void}
     */
    function addLineStyleDropdown(settingsPanel, markerControls, onStyleChange) {
        var lineStyleKeys = [
            "underline", "marker", "straightShear", "zigzag", "fill", "strikethrough", "doubleStrikethrough",
            "box", "circle", "hexagon", "emphasisDots"
        ];
        var lineStyleNames = [];
        for (var i = 0; i < lineStyleKeys.length; i++) lineStyleNames.push(getLabel("dropdown." + lineStyleKeys[i]));

        var lineStyleRow = settingsPanel.add("group");
        setupRow(lineStyleRow, "left", 8);
        var lineStyleLabel = lineStyleRow.add("statictext", undefined, labelText("fieldLabel.lineStyle"));
        var lineStyleDropdown = lineStyleRow.add("dropdownlist", undefined, lineStyleNames);
        lineStyleDropdown.preferredSize.width = LINE_STYLE_DROPDOWN_WIDTH;
        var chosenKey = DEFAULT_LINE_STYLE;
        /* selection の代入でも onChange が来るので、書き込み中は無視する / Assigning selection fires onChange too, so ignore it while writing */
        var isWritingSelection = false;

        /**
         * 種類を選んだ状態にする（知らないキーは無視）
         * @param {string} lineStyleKey - 種類のキー
         * @returns {void}
         */
        function selectLineStyle(lineStyleKey) {
            for (var j = 0; j < lineStyleKeys.length; j++) {
                if (lineStyleKeys[j] !== lineStyleKey) continue;
                chosenKey = lineStyleKey;
                isWritingSelection = true;
                lineStyleDropdown.selection = j;
                isWritingSelection = false;
                lineStyleDropdown.helpTip = getLabel("tooltip." + lineStyleKey);
                return;
            }
        }

        lineStyleDropdown.onChange = function () {
            if (isWritingSelection || !lineStyleDropdown.selection) return;
            chosenKey = lineStyleKeys[lineStyleDropdown.selection.index];
            lineStyleDropdown.helpTip = getLabel("tooltip." + chosenKey);
            onStyleChange();
        };
        selectLineStyle(DEFAULT_LINE_STYLE);
        markerControls.lineStyleRow = {
            row: lineStyleRow,
            label: lineStyleLabel,
            getChoice: function () { return chosenKey; },
            setChoice: selectLineStyle
        };
    }

    /**
     * ［設定］の右の列に、配置と大きさの項目（不透明度・位置調整・余白・行長・行送り・単位・角丸の半径・連続処理・回転・点の大きさ）を作り、markerControls に入れる
     * @param {Group} layoutColumn - 右の列
     * @param {Object} markerControls - 作ったコントロールを入れるオブジェクト（rulerUnit・strokeUnit 付き）
     * @param {number} glyphHeightPt - 余白・点の大きさの初期値のもとにする字面の高さ（pt）
     * @param {Function} onSettingChange - 設定を変えたあとに呼ぶ関数
     * @param {Function} onStyleChange - ［単位］を選び直したあとに呼ぶ関数（使える項目も切り替える）
     * @returns {void}
     */
    function addLayoutSettingFields(layoutColumn, markerControls, glyphHeightPt, onSettingChange, onStyleChange) {
        var rulerUnit = markerControls.rulerUnit;
        markerControls.opacityInput = addMarkerNumberField(layoutColumn, "opacity", {
            text: DEFAULT_OPACITY + "%", characters: 7, step: 5, min: 0, max: 100, unit: "%"
        }, onSettingChange);
        markerControls.verticalOffsetInput = addMarkerNumberField(layoutColumn, "verticalOffset", {
            text: "0 " + rulerUnit.label, characters: 7, step: 1, unit: " " + rulerUnit.label
        }, onSettingChange);
        addBoxPaddingFields(layoutColumn, markerControls, glyphHeightPt, onSettingChange);
        /* ［行長］は項目名の右に［行頭］［行頭行末］を縦に並べる / Line Length stacks Align Starts and Align Both Ends right of its label */
        var lineLengthRow = layoutColumn.add("group");
        setupRow(lineLengthRow, "left", 8);
        lineLengthRow.alignChildren = ["left", "top"];
        var lineLengthLabel = lineLengthRow.add("statictext", undefined, labelText("fieldLabel.lineLength"));
        var lineLengthOptionsGroup = lineLengthRow.add("group");
        lineLengthOptionsGroup.orientation = "column";
        lineLengthOptionsGroup.alignChildren = ["left", "center"];
        lineLengthOptionsGroup.spacing = 6;
        /* ［行頭行末］は行頭もそろえるので、［行頭］とはどちらか一方だけをオンにする / Align Both Ends includes Align Starts, so only one of them is on */
        markerControls.alignLineStartsOption = addOptionCheckbox(lineLengthOptionsGroup, "alignLineStarts", DEFAULT_ALIGN_LINE_STARTS && !DEFAULT_EQUAL_LINE_LENGTHS, function () {
            if (markerControls.alignLineStartsOption.getValue()) markerControls.equalLineLengthsOption.setValue(false);
            onSettingChange();
        });
        markerControls.equalLineLengthsOption = addOptionCheckbox(lineLengthOptionsGroup, "equalLineLengths", DEFAULT_EQUAL_LINE_LENGTHS, function () {
            if (markerControls.equalLineLengthsOption.getValue()) markerControls.alignLineStartsOption.setValue(false);
            onSettingChange();
        });
        markerControls.lineLengthLabel = lineLengthLabel;
        markerControls.leadingBasedOption = addLabeledOptionCheckbox(layoutColumn, "lineHeight", "leadingBased", DEFAULT_LEADING_BASED, onSettingChange);
        markerControls.drawUnitRow = addRadioChoiceRow(layoutColumn, "drawUnit", ["perLine", "perChar"], DEFAULT_PER_CHAR ? "perChar" : "perLine", onStyleChange);
        markerControls.drawUnitRow.radios.perLine.helpTip = getLabel("tooltip.perLine");
        markerControls.drawUnitRow.radios.perChar.helpTip = getLabel("tooltip.perChar");
        /* ほかの処理からは［1文字ずつ］のオンオフとして読み書きする / Read and written elsewhere as Per Character on or off */
        markerControls.perCharOption = {
            getValue: function () { return markerControls.drawUnitRow.getChoice() === "perChar"; },
            setValue: function (isPerChar) { markerControls.drawUnitRow.setChoice(isPerChar ? "perChar" : "perLine"); }
        };
        markerControls.cornerRadiusInput = addMarkerNumberField(layoutColumn, "cornerRadius", createLengthFieldOptions(DEFAULT_CORNER_RADIUS, rulerUnit, 0), onSettingChange);
        markerControls.squareLineJoinsOption = addLabeledOptionCheckbox(layoutColumn, "lineJoins", "squareLineJoins", DEFAULT_SQUARE_LINE_JOINS, onSettingChange);
        markerControls.boxRotationInput = addMarkerNumberField(layoutColumn, "boxRotation", {
            text: DEFAULT_BOX_ROTATION + "°", characters: 7, step: 1, min: -360, max: 360, unit: "°"
        }, onSettingChange);
        markerControls.dotSizeInput = addMarkerNumberField(layoutColumn, "dotSize", createLengthFieldOptions(glyphHeightPt * DOT_SIZE_RATIO, rulerUnit, 0.1), onSettingChange);
    }

    /**
     * ［文字］パネル（テキストカラー・自動カーニング・カーニング（前）・トラッキング・カーニング（後）・余白に合わせてカーニングを自動調整）を作り、markerControls に入れる。
     * テキストカラーとカーニングは1行ずつ、その下にカーニング（前）・トラッキング・カーニング（後）を、項目名を上にして3つ並べる
     * @param {Group} parent - 追加先
     * @param {Object} markerControls - 作ったコントロールを入れるオブジェクト
     * @param {Document} doc - 対象ドキュメント（カラーモード）
     * @param {Function} onSettingChange - 設定を変えたあとに呼ぶ関数
     * @param {Function} onStyleChange - ［余白に合わせてカーニングを自動調整］を切り替えたあとに呼ぶ関数（使える項目も切り替える）
     * @returns {void}
     */
    function addTextFormatPanel(parent, markerControls, doc, onSettingChange, onStyleChange) {
        var textFormatPanel = parent.add("panel", undefined, getLabel("panel.textFormat"));
        setupPanel(textFormatPanel);
        /* ［変更しない］は色ではないので、getColorOptions() の color は null になる / No Change is not a color, so getColorOptions() returns a null color */
        markerControls.textFillColorRow = addColorChoiceRow(textFormatPanel, doc, false, onSettingChange, "textFillColor", TEXT_FILL_COLOR_CHOICES, DEFAULT_TEXT_FILL_COLOR_CHOICE);
        markerControls.textFillColorRow.label.helpTip = getLabel("tooltip.textFillColor");
        markerControls.kerningRow = addRadioChoiceRow(textFormatPanel, "kerning",
            ["keepKerning", "zeroKerning", "metricsKerning", "opticalKerning"], DEFAULT_KERNING_CHOICE, onSettingChange);
        markerControls.kerningRow.radios.zeroKerning.helpTip = getLabel("tooltip.zeroKerning");
        markerControls.kerningRow.radios.metricsKerning.helpTip = getLabel("tooltip.metricsKerning");
        markerControls.kerningRow.radios.opticalKerning.helpTip = getLabel("tooltip.opticalKerning");

        /* 3つの欄と［余白に合わせてカーニングを自動調整］は、項目名の列のぶん右から始める（空の項目名の幅は、項目名をそろえたあとで合わせる）
           The three fields and Auto-Adjust Kerning to Padding start after the label column (the empty label gets its width after the labels are aligned) */
        var spacingRow = textFormatPanel.add("group");
        setupRow(spacingRow, "left", 8);
        spacingRow.alignChildren = ["left", "top"];
        markerControls.textSpacingIndent = spacingRow.add("statictext", undefined, "");
        var spacingBlock = spacingRow.add("group");
        spacingBlock.orientation = "column";
        spacingBlock.alignChildren = ["left", "top"];
        spacingBlock.spacing = PANEL_SPACING;
        var spacingFieldsGroup = spacingBlock.add("group");
        setupRow(spacingFieldsGroup, "left", COLUMN_SPACING);
        spacingFieldsGroup.alignChildren = ["left", "top"];

        /* ［余白に合わせてカーニングを自動調整］のとき、欄で変えた分は計算した値からの微調整として控える / With Auto-Adjust Kerning to Padding, edits are kept as adjustments to the computed values */
        markerControls.kerningAdjustments = { before: 0, after: 0 };
        markerControls.kerningBeforeInput = addStackedNumberField(spacingFieldsGroup, "kerningBefore", {
            text: DEFAULT_KERNING_BEFORE_TEXT, characters: 7, step: 10, min: -1000, max: 10000, allowBlank: true
        }, function () {
            rememberKerningAdjustment(markerControls, "before");
            onSettingChange();
        });
        markerControls.trackingInput = addStackedNumberField(spacingFieldsGroup, "tracking", {
            text: DEFAULT_TRACKING_TEXT, characters: 7, step: 10, min: -1000, max: 10000, allowBlank: true
        }, onSettingChange);
        markerControls.kerningAfterInput = addStackedNumberField(spacingFieldsGroup, "kerningAfter", {
            text: DEFAULT_KERNING_AFTER_TEXT, characters: 7, step: 10, min: -1000, max: 10000, allowBlank: true
        }, function () {
            rememberKerningAdjustment(markerControls, "after");
            onSettingChange();
        });
        markerControls.kerningFromPaddingOption = addOptionCheckbox(spacingBlock, "kerningFromPadding", DEFAULT_KERNING_FROM_PADDING, onStyleChange);
    }

    /**
     * 項目名を上、∧∨と入力欄を下に積んだ数値欄を作る（addMarkerNumberField() の行を縦向きにする）
     * @param {Group} parent - 追加先
     * @param {string} labelKey - addMarkerNumberField() と同じ
     * @param {Object} fieldOptions - addMarkerNumberField() と同じ
     * @param {Function} onValueChange - 値を変えたあとに呼ぶ関数
     * @returns {EditText} 入力欄（項目名は .fieldLabel）
     */
    function addStackedNumberField(parent, labelKey, fieldOptions, onValueChange) {
        var numberInput = addMarkerNumberField(parent, labelKey, fieldOptions, onValueChange);
        /* 項目名を上に置くのでコロンは付けない / No colon when the label sits above */
        numberInput.fieldLabel.text = getLabel("fieldLabel." + labelKey);
        var fieldRowGroup = numberInput.fieldLabel.parent;
        fieldRowGroup.orientation = "column";
        fieldRowGroup.alignChildren = ["left", "top"];
        fieldRowGroup.spacing = STACKED_FIELD_LABEL_SPACING;
        return numberInput;
    }

    /**
     * オンオフを変数で持つチェックボックスを作る（show() 前に代入した value は読み戻せないため）
     * @param {Group} parent - 追加先
     * @param {string} labelKey - LABELS.checkbox と LABELS.tooltip のキー
     * @param {boolean} initialValue - 最初のオンオフ
     * @param {Function} onValueChange - クリックで切り替えたあとに呼ぶ関数
     * @returns {{checkbox: Checkbox, getValue: Function, setValue: Function}} チェックボックスと、オンオフの読み書き
     */
    function addOptionCheckbox(parent, labelKey, initialValue, onValueChange) {
        var isChecked = initialValue;
        var optionCheckbox = parent.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        optionCheckbox.alignment = "left";
        optionCheckbox.helpTip = getLabel("tooltip." + labelKey);
        optionCheckbox.value = isChecked;
        optionCheckbox.onClick = function () {
            isChecked = optionCheckbox.value;
            onValueChange();
        };
        return {
            checkbox: optionCheckbox,
            getValue: function () { return isChecked; },
            setValue: function (shouldCheck) {
                isChecked = shouldCheck;
                optionCheckbox.value = shouldCheck;
            }
        };
    }

    /**
     * 項目名とチェックボックスを横に並べた1行を作る（例：「行長 : □揃える」）
     * @param {Group} parent - 追加先
     * @param {string} fieldLabelKey - LABELS.fieldLabel のキー（項目名）
     * @param {string} checkboxKey - LABELS.checkbox と LABELS.tooltip のキー
     * @param {boolean} initialValue - 最初のオンオフ
     * @param {Function} onValueChange - クリックで切り替えたあとに呼ぶ関数
     * @returns {{checkbox: Checkbox, label: StaticText, getValue: Function, setValue: Function}} addOptionCheckbox() の戻り値に項目名を足したもの
     */
    function addLabeledOptionCheckbox(parent, fieldLabelKey, checkboxKey, initialValue, onValueChange) {
        var optionRow = parent.add("group");
        setupRow(optionRow, "left", 8);
        var optionLabel = optionRow.add("statictext", undefined, labelText("fieldLabel." + fieldLabelKey));
        var labeledOption = addOptionCheckbox(optionRow, checkboxKey, initialValue, onValueChange);
        labeledOption.label = optionLabel;
        return labeledOption;
    }

    /**
     * 長さの欄（定規の単位）の指定を作る
     * @param {number} valuePt - 初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位（getUnitInfo() の戻り値）
     * @param {number} [minValue] - 下限（定規の単位。省くと下限なし）
     * @returns {Object} addMarkerNumberField() の fieldOptions
     */
    function createLengthFieldOptions(valuePt, rulerUnit, minValue) {
        return {
            text: formatStepperNumber(ptToUnit(valuePt, rulerUnit)) + " " + rulerUnit.label,
            characters: 7, step: 1, min: minValue, unit: " " + rulerUnit.label
        };
    }

    /**
     * ［設定］の左の列に、線と種類ごとの効果の項目（線幅・線の種類・分割数・間隔・線端・線の間隔・角度・大きさ・折り返し・ポイント・ラフ・サイズ）を作り、markerControls に入れる
     * @param {Group} strokeColumn - 左の列
     * @param {Object} markerControls - 作ったコントロールを入れるオブジェクト（rulerUnit・strokeUnit 付き）
     * @param {number} glyphHeightPt - 線幅・間隔・線の間隔の初期値のもとにする字面の高さ（pt）
     * @param {Function} onSettingChange - 設定を変えたあとに呼ぶ関数
     * @param {Function} onStyleChange - 線の種類・［手書き風］を切り替えたあとに呼ぶ関数（使える項目も切り替える）
     * @returns {void}
     */
    function addStrokeSettingFields(strokeColumn, markerControls, glyphHeightPt, onSettingChange, onStyleChange) {
        var rulerUnit = markerControls.rulerUnit;
        var strokeUnit = markerControls.strokeUnit;
        markerControls.strokeWidthInput = addMarkerNumberField(strokeColumn, "strokeWidth", {
            text: formatStepperNumber(ptToUnit(glyphHeightPt * MARKER_WEIGHT_RATIO, strokeUnit)) + " " + strokeUnit.label,
            characters: 7, step: 1, min: 0.1, unit: " " + strokeUnit.label
        }, onSettingChange);
        markerControls.dashStyleRow = addRadioChoiceRow(strokeColumn, "dashStyle", ["solidLine", "dashedLine", "dottedLine"], DEFAULT_DASH_STYLE, onStyleChange);
        markerControls.dashStyleRow.radios.dashedLine.helpTip = getLabel("tooltip.dashedLine");
        markerControls.dashStyleRow.radios.dottedLine.helpTip = getLabel("tooltip.dottedLine");
        markerControls.dashSegmentsInput = addMarkerNumberField(strokeColumn, "dashSegments", {
            text: DEFAULT_DASH_SEGMENTS_TEXT, characters: 7, step: 1, min: 1, integer: true, allowBlank: true
        }, onSettingChange);
        markerControls.dashGapInput = addMarkerNumberField(strokeColumn, "dashGap", {
            text: formatStepperNumber(ptToUnit(glyphHeightPt * MARKER_WEIGHT_RATIO * DASHED_LINE_RATIOS[1], strokeUnit)) + " " + strokeUnit.label,
            characters: 7, step: 1, min: 0, unit: " " + strokeUnit.label
        }, onSettingChange);
        markerControls.strokeCapRow = addRadioChoiceRow(strokeColumn, "strokeCap", ["buttCap", "roundCap"], DEFAULT_STROKE_CAP, onSettingChange);
        markerControls.strokeCapRow.radios.roundCap.helpTip = getLabel("tooltip.roundCap");
        markerControls.lineGapInput = addMarkerNumberField(strokeColumn, "lineGap", createLengthFieldOptions(glyphHeightPt * LINE_GAP_RATIO, rulerUnit, 0), onSettingChange);
        markerControls.shearAngleInput = addMarkerNumberField(strokeColumn, "shearAngle", {
            text: DEFAULT_SHEAR_ANGLE + "°", characters: 7, step: 1, min: -80, max: 80, unit: "°"
        }, onSettingChange);
        markerControls.zigzagSizeInput = addMarkerNumberField(strokeColumn, "zigzagSize", createLengthFieldOptions(DEFAULT_ZIGZAG_SIZE, rulerUnit, 0.1), onSettingChange);
        markerControls.zigzagRidgesInput = addMarkerNumberField(strokeColumn, "zigzagRidges", {
            text: String(DEFAULT_ZIGZAG_RIDGES), characters: 7, step: 1, min: 1, integer: true
        }, onSettingChange);
        markerControls.zigzagPointsRow = addRadioChoiceRow(strokeColumn, "zigzagPoints", ["smoothPoints", "cornerPoints"], DEFAULT_ZIGZAG_POINTS, onSettingChange);
        markerControls.zigzagPointsRow.radios.smoothPoints.helpTip = getLabel("tooltip.smoothPoints");
        markerControls.zigzagPointsRow.radios.cornerPoints.helpTip = getLabel("tooltip.cornerPoints");
        /* ［ラフ］：［手書き風］をオンにするとサイズを使い、詳細は［折り返し］の値を使う / Roughen: Size is used while Hand-Drawn is on; the detail is Ridges */
        markerControls.roughenOption = addLabeledOptionCheckbox(strokeColumn, "roughen", "roughen", DEFAULT_ROUGHEN, onStyleChange);
        markerControls.roughenSizeInput = addMarkerNumberField(strokeColumn, "roughenSize", {
            text: DEFAULT_ROUGHEN_SIZE_PERCENT + "%", characters: 7, step: 1, min: 0, max: 100, unit: "%"
        }, onSettingChange);
    }

    /**
     * ［余白（横）］［余白（縦）］の2欄と、2つを連動させるリンクアイコンを作り、markerControls に入れる。
     * 連動中は、片方を変えるともう片方を同じ値にする
     * @param {Group} paddingColumn - 追加先の列（右の列）
     * @param {Object} markerControls - 作ったコントロールを入れるオブジェクト（rulerUnit 付き）
     * @param {number} glyphHeightPt - 余白の初期値のもとにする字面の高さ（pt）
     * @param {Function} onSettingChange - 設定を変えたあとに呼ぶ関数
     * @returns {void}
     */
    function addBoxPaddingFields(paddingColumn, markerControls, glyphHeightPt, onSettingChange) {
        /* 2欄を縦に積んだ group とアイコンを横に並べる / Put the stacked fields and the icon side by side */
        var paddingRowGroup = paddingColumn.add("group");
        paddingRowGroup.orientation = "row";
        paddingRowGroup.alignChildren = ["left", "center"];
        paddingRowGroup.spacing = 6;
        var paddingFieldsGroup = paddingRowGroup.add("group");
        paddingFieldsGroup.orientation = "column";
        paddingFieldsGroup.alignChildren = ["left", "top"];
        paddingFieldsGroup.spacing = PANEL_SPACING;

        /**
         * 連動中なら、変えた欄の値をもう片方へ写す
         * @param {EditText} changedInput - 変えた欄
         * @param {EditText} pairedInput - もう片方の欄
         * @returns {void}
         */
        function syncPairedPadding(changedInput, pairedInput) {
            if (!markerControls.boxPaddingLink.value) return;
            pairedInput.text = changedInput.text;
            pairedInput.lastValidText = changedInput.text;
        }

        var initialPaddingPt = glyphHeightPt * BOX_PADDING_RATIO;
        markerControls.boxPaddingInput = addMarkerNumberField(paddingFieldsGroup, "boxPadding", createLengthFieldOptions(initialPaddingPt, markerControls.rulerUnit), function () {
            syncPairedPadding(markerControls.boxPaddingInput, markerControls.boxPaddingVInput);
            onSettingChange();
        });
        markerControls.boxPaddingVInput = addMarkerNumberField(paddingFieldsGroup, "boxPaddingV", createLengthFieldOptions(initialPaddingPt, markerControls.rulerUnit), function () {
            syncPairedPadding(markerControls.boxPaddingVInput, markerControls.boxPaddingInput);
            onSettingChange();
        });
        markerControls.boxPaddingVInput.helpTip = markerControls.boxPaddingInput.helpTip;
        /* 連動にしたときは［余白（横）］を［余白（縦）］へ写す / When linked, copy Padding (H) to Padding (V) */
        markerControls.boxPaddingLink = addLinkToggle(paddingRowGroup, DEFAULT_BOX_PADDING_LINKED, function () {
            syncPairedPadding(markerControls.boxPaddingInput, markerControls.boxPaddingVInput);
            onSettingChange();
        });
        markerControls.boxPaddingLink.helpTip = getLabel("tooltip.boxPaddingLink");
    }

    /**
     * ［設定］パネルの1列を作る
     * @param {Group} settingsColumnsGroup - 列を横に並べるグループ
     * @returns {Group} 列のグループ
     */
    function addSettingsColumn(settingsColumnsGroup) {
        var settingsColumn = settingsColumnsGroup.add("group");
        settingsColumn.orientation = "column";
        settingsColumn.alignChildren = ["fill", "top"];
        settingsColumn.alignment = ["fill", "top"];
        settingsColumn.spacing = PANEL_SPACING;
        return settingsColumn;
    }

    /**
     * 種類に合わせて、使う項目だけを使えるようにする。
     * ［線幅］は線を引く種類と、［線］の［描画］がオンの長方形の仲間、［線の種類］はさらに［塗りつぶし］でも使い、［線端］は線を引く種類だけ（［ドット］のときは丸型に固定）。
     * ［余白（横）］［余白（縦）］と連動のアイコンは長方形の仲間と［傍点］、［図形］［回転］［1文字ずつ］は長方形の仲間、［角丸の半径］は［長方形］で使う。
     * ［線］のカラーは線で引く種類（［塗りつぶし］を除く）と［線］の［描画］がオンの長方形の仲間、［塗り］のカラーは［塗りつぶし］［傍点］と［塗り］の［描画］がオンの長方形の仲間で使う。
     * ［余白に合わせてカーニングを自動調整］は行ごとの長方形の仲間で使い、オンの間はカーニング（前・後）の欄に計算した値（微調整込み）を出す
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {void}
     */
    function updateStyleFieldsEnabled(markerControls) {
        var lineStyle = markerControls.lineStyleRow.getChoice();
        var usesShape = isShapeStyle(lineStyle);
        var isDots = (lineStyle === "emphasisDots");
        var isLineStyle = !usesShape && !isDots;
        var isShapeFill = usesShape && markerControls.shapeFillOption.getValue();
        var isShapeStroke = usesShape && markerControls.shapeStrokeOption.getValue();
        var isStrokedStyle = isLineStyle || isShapeStroke;
        setSteppedFieldEnabled(markerControls.shearAngleInput, lineStyle === "straightShear");
        setSteppedFieldEnabled(markerControls.zigzagSizeInput, lineStyle === "zigzag");
        var usesRoughen = !isDots && markerControls.roughenOption.getValue();
        setSteppedFieldEnabled(markerControls.zigzagRidgesInput, lineStyle === "zigzag" || usesRoughen);
        setRadiosEnabled(markerControls.zigzagPointsRow, lineStyle === "zigzag");
        markerControls.roughenOption.checkbox.enabled = !isDots;
        setSteppedFieldEnabled(markerControls.roughenSizeInput, usesRoughen);
        setSteppedFieldEnabled(markerControls.boxPaddingInput, usesShape || isDots);
        setSteppedFieldEnabled(markerControls.boxPaddingVInput, usesShape || isDots);
        setLinkToggleEnabled(markerControls.boxPaddingLink, usesShape || isDots);
        setSteppedFieldEnabled(markerControls.cornerRadiusInput, lineStyle === "box");
        var usesSquareLineJoins = (lineStyle === "box" && !markerControls.perCharOption.getValue());
        markerControls.squareLineJoinsOption.checkbox.enabled = usesSquareLineJoins;
        markerControls.squareLineJoinsOption.label.enabled = usesSquareLineJoins;
        var usesEqualLineLengths = isLineStyle || (usesShape && !markerControls.perCharOption.getValue());
        markerControls.alignLineStartsOption.checkbox.enabled = usesEqualLineLengths;
        markerControls.equalLineLengthsOption.checkbox.enabled = usesEqualLineLengths;

        /* 線で引く種類は［線］、［塗りつぶし］［傍点］は［塗り］のカラー、長方形の仲間は［塗り］［線］の［描画］のオンに合わせて使う。
           ［描画］は長方形の仲間だけで使い、項目名は行のどこかを使う間は有効のまま
           Line styles use the Stroke color, Fill and Emphasis Dots the Fill color, shapes follow Draw; Draw is for shapes only,
           and the label stays enabled while any part of the row is in use */
        var isFillLine = (lineStyle === "fill");
        var usesFillColor = isFillLine || isDots || isShapeFill;
        var usesStrokeColor = (isLineStyle && !isFillLine) || isShapeStroke;
        markerControls.fillColorRow.setEnabled(usesFillColor);
        markerControls.strokeColorRow.setEnabled(usesStrokeColor);
        markerControls.shapeFillOption.checkbox.enabled = usesShape;
        markerControls.shapeStrokeOption.checkbox.enabled = usesShape;
        markerControls.fillColorRow.label.enabled = usesFillColor || usesShape;
        markerControls.strokeColorRow.label.enabled = usesStrokeColor || usesShape;
        setSteppedFieldEnabled(markerControls.boxRotationInput, usesShape);
        setRadiosEnabled(markerControls.drawUnitRow, usesShape);
        setSteppedFieldEnabled(markerControls.lineGapInput, lineStyle === "doubleStrikethrough");
        setSteppedFieldEnabled(markerControls.dotSizeInput, isDots);
        setSteppedFieldEnabled(markerControls.strokeWidthInput, isStrokedStyle && lineStyle !== "fill");
        setRadiosEnabled(markerControls.dashStyleRow, isStrokedStyle);
        /* ［分割数］は破線だけ、［間隔］は破線とドットで使う / Segments is for Dashed only; Gap is for Dashed and Dotted */
        var dashStyle = markerControls.dashStyleRow.getChoice();
        var isDashed = isStrokedStyle && dashStyle === "dashedLine";
        var isDotted = isStrokedStyle && dashStyle === "dottedLine";
        setSteppedFieldEnabled(markerControls.dashSegmentsInput, isDashed);
        setSteppedFieldEnabled(markerControls.dashGapInput, isDashed || isDotted);
        /* ドットは丸型の線端で点にするので、［線端］は選べない / Dotted needs round caps for its dots, so Cap is not selectable */
        setRadiosEnabled(markerControls.strokeCapRow, isLineStyle && dashStyle !== "dottedLine");
        markerControls.kerningFromPaddingOption.checkbox.enabled = usesShape && !markerControls.perCharOption.getValue();
        updatePaddingKerningFields(markerControls);
    }

    /**
     * ［余白に合わせてカーニングを自動調整］が効いているか（行ごとの長方形の仲間で、オン）を返す
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {boolean} 効いていれば true
     */
    function isKerningFromPaddingActive(markerControls) {
        return isShapeStyle(markerControls.lineStyleRow.getChoice()) && !markerControls.perCharOption.getValue()
            && markerControls.kerningFromPaddingOption.getValue();
    }

    /**
     * ［余白に合わせてカーニングを自動調整］が効いているとき、カーニング（前・後）の欄に、先頭の対象で計算した値に微調整を足して出す
     * @param {Object} markerControls - addMarkerControls() の戻り値（kerningReference 付き。無ければ何もしない）
     * @returns {void}
     */
    function updatePaddingKerningFields(markerControls) {
        if (!markerControls.kerningReference || !isKerningFromPaddingActive(markerControls)) return;
        var kerningFields = getPaddingKerningFields(markerControls);
        for (var edgeKey in kerningFields) {
            if (!kerningFields.hasOwnProperty(edgeKey)) continue;
            var kerningField = kerningFields[edgeKey];
            var kerningValue = calcPaddingKerning(kerningField.paddingPt, kerningField.fontSize, markerControls.kerningAdjustments[edgeKey]);
            kerningField.input.text = (kerningValue === null) ? "" : String(kerningValue);
            kerningField.input.lastValidText = kerningField.input.text;
        }
    }

    /**
     * ［余白に合わせてカーニングを自動調整］が効いているときに、カーニングの欄の値と計算した値の差を微調整として控える（空欄なら微調整なし）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {string} edgeKey - "before"（前）か "after"（後）
     * @returns {void}
     */
    function rememberKerningAdjustment(markerControls, edgeKey) {
        if (!markerControls.kerningReference || !isKerningFromPaddingActive(markerControls)) return;
        var kerningField = getPaddingKerningFields(markerControls)[edgeKey];
        var typedKerning = readOptionalNumber(kerningField.input);
        var computedKerning = calcPaddingKerning(kerningField.paddingPt, kerningField.fontSize, 0) || 0;
        markerControls.kerningAdjustments[edgeKey] = (typedKerning === null) ? 0 : typedKerning - computedKerning;
    }

    /**
     * ［余白に合わせてカーニングを自動調整］の計算に使う、カーニング（前・後）の欄・余白・先頭の対象の文字サイズを返す
     * @param {Object} markerControls - addMarkerControls() の戻り値（kerningReference 付き）
     * @returns {{before: Object, after: Object}} それぞれ { input, paddingPt, fontSize }
     */
    function getPaddingKerningFields(markerControls) {
        var kerningReference = markerControls.kerningReference;
        /* 縦組みは列の向きの［余白（縦）］ / Vertical text uses Padding (V), along the column */
        var paddingInput = kerningReference.isVertical ? markerControls.boxPaddingVInput : markerControls.boxPaddingInput;
        var paddingPt = readLengthPt(paddingInput, markerControls.rulerUnit);
        return {
            before: { input: markerControls.kerningBeforeInput, paddingPt: paddingPt, fontSize: kerningReference.firstSize },
            after: { input: markerControls.kerningAfterInput, paddingPt: paddingPt, fontSize: kerningReference.lastSize }
        };
    }

    /**
     * 余白の長さを手動カーニングの値にし、微調整を足す（余白 ÷ 文字サイズ × 1000 を整数に丸め、-1000〜10000 に収める）
     * @param {number} paddingPt - 余白（pt）
     * @param {number} fontSizePt - 文字サイズ（pt）
     * @param {number} adjustment - 微調整（1/1000 em）
     * @returns {number|null} カーニング（1/1000 em）。余白が 0 以下（かサイズが読めない）で微調整も無ければ null（変更しない）
     */
    function calcPaddingKerning(paddingPt, fontSizePt, adjustment) {
        var computedKerning = (paddingPt > 0 && fontSizePt > 0) ? Math.round(paddingPt / fontSizePt * 1000) : 0;
        if (computedKerning === 0 && !adjustment) return null;
        return Math.max(-1000, Math.min(10000, computedKerning + (adjustment || 0)));
    }

    /**
     * 対象の先頭と最後の文字のサイズを返す（［余白に合わせてカーニングを自動調整］の計算に使う）
     * @param {TextFrame} textFrame - 元のテキストフレーム、またはその複製
     * @param {Object} markerTarget - collectMarkerTargets() の要素（spanStart・spanLength）
     * @returns {{firstSize: number, lastSize: number}} 文字サイズ（pt）
     */
    function getTargetEdgeFontSizes(textFrame, markerTarget) {
        return {
            firstSize: textFrame.characters[markerTarget.spanStart].characterAttributes.size,
            lastSize: textFrame.characters[markerTarget.spanStart + markerTarget.spanLength - 1].characterAttributes.size
        };
    }

    /**
     * addRadioChoiceRow() の行のラジオボタンを、項目名ごとまとめて有効／無効にする
     * @param {Object} choiceRow - addRadioChoiceRow() の戻り値
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setRadiosEnabled(choiceRow, isEnabled) {
        if (choiceRow.label) choiceRow.label.enabled = isEnabled;
        for (var radioKey in choiceRow.radios) {
            if (choiceRow.radios.hasOwnProperty(radioKey)) choiceRow.radios[radioKey].enabled = isEnabled;
        }
    }

    /**
     * 線のかわりに図形（長方形・円・六角形）を描く種類かを返す
     * @param {string} lineStyle - 種類のキー
     * @returns {boolean} ［長方形］［円］［六角形］なら true
     */
    function isShapeStyle(lineStyle) {
        return lineStyle === "box" || lineStyle === "circle" || lineStyle === "hexagon";
    }

    /**
     * 範囲を囲む六角形の頂点を返す。
     * 1文字ずつなら、範囲の中心に、長い辺を平らな辺どうしの間隔にした正六角形（左右がとがる）。
     * 行ごとなら、範囲の左右（縦組みは上下）に 60° でとがった先を足した六角形
     * @param {number} areaLeft - 範囲の左端
     * @param {number} areaTop - 範囲の上端
     * @param {number} areaWidth - 範囲の幅
     * @param {number} areaHeight - 範囲の高さ
     * @param {boolean} isPerChar - 1文字ずつか
     * @param {boolean} isVertical - 縦組みか
     * @returns {number[][]} 頂点 [x, y] の配列
     */
    function getHexagonPoints(areaLeft, areaTop, areaWidth, areaHeight, isPerChar, isVertical) {
        var centerX = areaLeft + areaWidth / 2;
        var centerY = areaTop - areaHeight / 2;
        var areaRight = areaLeft + areaWidth;
        var areaBottom = areaTop - areaHeight;
        if (isPerChar) {
            var halfFlat = Math.max(areaWidth, areaHeight) / 2;
            var sideLength = halfFlat * 2 / Math.sqrt(3);
            return [
                [centerX - sideLength, centerY], [centerX - sideLength / 2, centerY + halfFlat], [centerX + sideLength / 2, centerY + halfFlat],
                [centerX + sideLength, centerY], [centerX + sideLength / 2, centerY - halfFlat], [centerX - sideLength / 2, centerY - halfFlat]
            ];
        }
        if (isVertical) {
            var tipHeight = areaWidth / (2 * Math.sqrt(3));
            return [
                [centerX, areaTop + tipHeight], [areaRight, areaTop], [areaRight, areaBottom],
                [centerX, areaBottom - tipHeight], [areaLeft, areaBottom], [areaLeft, areaTop]
            ];
        }
        var tipWidth = areaHeight / (2 * Math.sqrt(3));
        return [
            [areaLeft - tipWidth, centerY], [areaLeft, areaTop], [areaRight, areaTop],
            [areaRight + tipWidth, centerY], [areaRight, areaBottom], [areaLeft, areaBottom]
        ];
    }

    /**
     * 今のダイアログボックスの指定を、線を引くための値にまとめる
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {Object} { fillColor, fillUseTextColor, opacity, strokeWidth（pt）, verticalOffset（pt）, lineStyle, shearAngle（度）, zigzagSize（pt）, zigzagRidges, isZigzagSmooth, isRoughen, roughenSize（％）, boxPadding（横、pt）, boxPaddingV（縦、pt）, cornerRadius（pt）, boxRotation（度）, lineGap（pt）, dotSize（pt）, dashStyle, dashSegments（空欄なら null）, dashGap（pt）, strokeCap, isFront, isPerChar, squareLineJoins, alignLineStarts, equalLineLengths, isLeadingBased, isShapeFill, isShapeStroke, strokeColor, strokeUseTextColor, kerningChoice, tracking・kerningBefore・kerningAfter（空欄なら null）, kerningFromPadding, kerningAdjustBefore・kerningAdjustAfter（［余白に合わせてカーニングを自動調整］の微調整）, textFillColor（［変更しない］なら null） }
     */
    function readMarkerOptions(markerControls) {
        var fillColorOptions = markerControls.fillColorRow.getColorOptions();
        var strokeColorOptions = markerControls.strokeColorRow.getColorOptions();
        return {
            fillColor: fillColorOptions.color,
            fillUseTextColor: fillColorOptions.useTextColor,
            opacity: parseFloat(markerControls.opacityInput.text),
            strokeWidth: readLengthPt(markerControls.strokeWidthInput, markerControls.strokeUnit),
            dashSegments: readOptionalNumber(markerControls.dashSegmentsInput),
            dashGap: readLengthPt(markerControls.dashGapInput, markerControls.strokeUnit),
            verticalOffset: readLengthPt(markerControls.verticalOffsetInput, markerControls.rulerUnit),
            lineStyle: markerControls.lineStyleRow.getChoice(),
            shearAngle: parseFloat(markerControls.shearAngleInput.text),
            zigzagSize: readLengthPt(markerControls.zigzagSizeInput, markerControls.rulerUnit),
            isRoughen: (markerControls.lineStyleRow.getChoice() !== "emphasisDots" && markerControls.roughenOption.getValue()),
            roughenSize: parseFloat(markerControls.roughenSizeInput.text),
            zigzagRidges: parseInt(markerControls.zigzagRidgesInput.text, 10),
            isZigzagSmooth: (markerControls.zigzagPointsRow.getChoice() === "smoothPoints"),
            boxPadding: readLengthPt(markerControls.boxPaddingInput, markerControls.rulerUnit),
            boxPaddingV: readLengthPt(markerControls.boxPaddingVInput, markerControls.rulerUnit),
            cornerRadius: readLengthPt(markerControls.cornerRadiusInput, markerControls.rulerUnit),
            boxRotation: parseFloat(markerControls.boxRotationInput.text),
            lineGap: readLengthPt(markerControls.lineGapInput, markerControls.rulerUnit),
            dotSize: readLengthPt(markerControls.dotSizeInput, markerControls.rulerUnit),
            dashStyle: markerControls.dashStyleRow.getChoice(),
            strokeCap: (markerControls.strokeCapRow.getChoice() === "roundCap") ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP,
            isFront: (markerControls.stackOrderRow.getChoice() === "stackFront"),
            isPerChar: markerControls.perCharOption.getValue(),
            squareLineJoins: markerControls.squareLineJoinsOption.getValue(),
            alignLineStarts: markerControls.alignLineStartsOption.getValue(),
            equalLineLengths: markerControls.equalLineLengthsOption.getValue(),
            isLeadingBased: markerControls.leadingBasedOption.getValue(),
            isShapeFill: markerControls.shapeFillOption.getValue(),
            isShapeStroke: markerControls.shapeStrokeOption.getValue(),
            strokeColor: strokeColorOptions.color,
            strokeUseTextColor: strokeColorOptions.useTextColor,
            kerningChoice: markerControls.kerningRow.getChoice(),
            tracking: readOptionalNumber(markerControls.trackingInput),
            kerningBefore: readOptionalNumber(markerControls.kerningBeforeInput),
            kerningAfter: readOptionalNumber(markerControls.kerningAfterInput),
            kerningFromPadding: isKerningFromPaddingActive(markerControls),
            kerningAdjustBefore: markerControls.kerningAdjustments.before,
            kerningAdjustAfter: markerControls.kerningAdjustments.after,
            textFillColor: markerControls.textFillColorRow.getColorOptions().color
        };
    }

    /**
     * 長さの欄の値を pt で読む
     * @param {EditText} numberInput - addMarkerNumberField() で作った入力欄
     * @param {{pointsPerUnit: number}} fieldUnit - 欄の単位（getUnitInfo() の戻り値）
     * @returns {number} 値（pt）
     */
    function readLengthPt(numberInput, fieldUnit) {
        return unitToPt(parseFloat(numberInput.text), fieldUnit);
    }

    /**
     * 空欄にできる数値欄の値を読む
     * @param {EditText} numberInput - addMarkerNumberField() で作った入力欄（allowBlank）
     * @returns {number|null} 値。空欄なら null
     */
    function readOptionalNumber(numberInput) {
        return /^\s*$/.test(numberInput.text) ? null : parseFloat(numberInput.text);
    }

    /**
     * 項目名の幅を列ごとにそろえ、［文字］パネルの3つの欄の字下げを項目名の列の幅に合わせる
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {void}
     */
    function alignMarkerLabels(markerControls) {
        for (var i = 0; i < markerControls.labelColumns.length; i++) alignLabelWidths(markerControls.labelColumns[i]);
        /* 空の statictext は幅を測れないので、そろえた項目名の幅を写す / An empty statictext cannot be measured, so copy the aligned width */
        markerControls.textSpacingIndent.preferredSize.width = markerControls.kerningRow.label.preferredSize.width;
    }

    /**
     * 線・図形と文字の書式を選ぶダイアログボックスを表示し、選んでいるあいだプレビューする
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} markerTargets - collectMarkerTargets() の戻り値（線幅の初期値とプレビュー用）
     * @returns {Object|null} readMarkerOptions() の値。キャンセルなら null
     */
    function showMarkerDialog(doc, markerTargets) {
        var markerDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(markerDialog);

        var markerPreview = createMarkerPreview(doc, markerTargets);
        /* 前回閉じたときの状態（無ければ空） / State at the last close (empty if none) */
        var sessionState = dialogSessionStore.load({});

        /* 左に［プリセット］、右に［設定］と［文字］ / Presets on the left, Settings and Text on the right */
        var panelColumnsGroup = markerDialog.add("group");
        panelColumnsGroup.orientation = "row";
        panelColumnsGroup.alignChildren = ["fill", "fill"];
        panelColumnsGroup.spacing = COLUMN_SPACING;
        var presetControls = addPresetPanel(panelColumnsGroup);
        /* 線幅・余白の初期値と、最初からあるプリセットの比率は、先頭の対象の字面の高さから
           Initial width and padding, and the built-in preset ratios, come from the first target's glyph height */
        var glyphHeightPt = getFirstLineGlyphHeight(markerTargets[0]);
        presetControls.builtinPresets = createBuiltinPresets(doc, glyphHeightPt);
        var rightPanelsGroup = panelColumnsGroup.add("group");
        rightPanelsGroup.orientation = "column";
        rightPanelsGroup.alignChildren = ["fill", "top"];
        rightPanelsGroup.spacing = WINDOW_SPACING;
        var markerControls = addMarkerControls(rightPanelsGroup, doc, glyphHeightPt, areTextColorsUsable(markerTargets), onSettingEdited);
        alignMarkerLabels(markerControls);
        /* ［余白に合わせてカーニングを自動調整］の欄の値は、先頭の対象の文字サイズで計算する / Auto-Adjust Kerning to Padding shows values for the first target's type size */
        markerControls.kerningReference = getTargetEdgeFontSizes(markerTargets[0].sourceFrame, markerTargets[0]);
        markerControls.kerningReference.isVertical = markerTargets[0].isVertical;
        bindPresetControls(presetControls, markerControls, refreshPreview);
        restoreDialogState(presetControls, markerControls, sessionState.settings ? sessionState : null);
        /* ［リセット］で戻す、開いたときの状態（すべての項目を控える） / The opening state Reset returns to (every field) */
        var openingState = { settings: collectPresetData(markerControls), presetName: getSelectedPresetName(presetControls) };

        var buttonRow = addButtonRow(markerDialog);
        var resetButton = buttonRow.leftGroup.add("button", undefined, getLabel("button.reset"));
        resetButton.helpTip = getLabel("tooltip.reset");
        resetButton.onClick = function () {
            restoreDialogState(presetControls, markerControls, openingState);
            clearTextSpacing(markerControls);
            updateStyleFieldsEnabled(markerControls);
            refreshPreview();
        };
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * 設定を手で変えたら、プリセットの選択を外してから描き直す。
         * 選んだままだと、同じプリセットをクリックしても onChange が来ず、保存した値に戻せない
         * @returns {void}
         */
        function onSettingEdited() {
            presetControls.clearPresetSelection();
            refreshPreview();
        }

        /**
         * プレビューを今の指定で描き直す（プレビューは常にオン）
         * @returns {void}
         */
        function refreshPreview() {
            updatePaddingKerningFields(markerControls);
            markerPreview.update(readMarkerOptions(markerControls));
            app.redraw();
        }

        refreshPreview();
        addVersionFooter(markerDialog);
        prepareDialogWindow(markerDialog, SCRIPT_NAME);
        var dialogResult = markerDialog.show();
        /* ［OK］でも［キャンセル］でも、閉じたときの状態を覚える / Remember the state at close, whether OK or Cancel */
        dialogSessionStore.save({
            settings: collectPresetData(markerControls),
            presetName: getSelectedPresetName(presetControls)
        });
        /* 確定でも消して、main() で描き直す / Removed on OK too; main() draws the final result */
        markerPreview.clear();
        if (dialogResult !== 1) return null;
        return readMarkerOptions(markerControls);
    }

    /**
     * カーニング（前・後）とトラッキングを空欄（変更しない）にし、［余白に合わせてカーニングを自動調整］と微調整を外す（［リセット］用）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @returns {void}
     */
    function clearTextSpacing(markerControls) {
        var spacingInputs = [markerControls.kerningBeforeInput, markerControls.trackingInput, markerControls.kerningAfterInput];
        for (var i = 0; i < spacingInputs.length; i++) {
            spacingInputs[i].text = "";
            spacingInputs[i].lastValidText = "";
        }
        markerControls.kerningFromPaddingOption.setValue(false);
        markerControls.kerningAdjustments.before = 0;
        markerControls.kerningAdjustments.after = 0;
    }

    /**
     * ダイアログボックスを開いたときの設定にする。
     * 前回閉じたときの状態があれば、その設定と選んでいたプリセットに戻す。無ければ（Illustrator を起動して最初）INITIAL_PRESET_NAME のプリセットを選んで読み込む
     * @param {Object} presetControls - addPresetPanel() の戻り値（builtinPresets 付き、bindPresetControls() のあと）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object|null} sessionState - 前回閉じたときの状態 { settings, presetName }（無ければ null）
     * @returns {void}
     */
    function restoreDialogState(presetControls, markerControls, sessionState) {
        var presetName = sessionState ? sessionState.presetName : getLabel(INITIAL_PRESET_NAME);
        var presetData = sessionState ? sessionState.settings : findPresetData(presetControls, presetName);
        if (presetData) applyPresetData(markerControls, presetData);
        updateStyleFieldsEnabled(markerControls);
        /* リストは選ぶだけにする（選び直しで読み込むと、前回の調整が消える） / Only select in the list; reloading would drop the last adjustments */
        presetControls.refreshPresetList((typeof presetName === "string") ? presetName : null);
    }

    /**
     * 対象の1行目（縦組みは1列目）の字面の高さを返す（複数行でも、線幅などの初期値を1行分の大きさにする）
     * @param {Object} markerTarget - collectMarkerTargets() の要素
     * @returns {number} 字面の高さ（pt）
     */
    function getFirstLineGlyphHeight(markerTarget) {
        selectTargetLayout(markerTarget, "", markerTarget.sourceFrame);
        var textLines = groupCharBoundsByLine(getTargetCharBounds(markerTarget), markerTarget.isVertical);
        var firstBounds = (textLines.length > 0) ? textLines[0] : markerTarget.glyphBounds;
        /* 縦組みの列は上下が長いので、字面の大きさは幅で取る / A vertical column is tall, so take its width */
        return markerTarget.isVertical ? (firstBounds.right - firstBounds.left) : (firstBounds.top - firstBounds.bottom);
    }

    /**
     * プリセットのリスト・保存・削除の動作を付け、リストを埋める（作り直しの関数を presetControls.refreshPresetList に置く）
     * @param {Object} presetControls - addPresetPanel() の戻り値（builtinPresets 付き）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Function} onPresetApplied - プリセットを書き込んだあとに呼ぶ関数
     * @returns {void}
     */
    function bindPresetControls(presetControls, markerControls, onPresetApplied) {
        /* selection の代入でも onChange が来るので、作り直しの間は無視する / Assigning selection fires onChange too, so ignore it while refilling */
        var isRefillingList = false;
        /* プリセットを書き込んでいる間は、設定の変更で選択を外さない / Keep the selection while a preset is being written */
        var isApplyingPreset = false;

        /**
         * プリセットのリストを作り直す（作り直しの間の onChange は無視する）
         * @param {string|null} selectedName - 選んでおくプリセット名
         * @returns {void}
         */
        function refreshPresetList(selectedName) {
            isRefillingList = true;
            fillPresetList(presetControls, selectedName);
            isRefillingList = false;
        }

        presetControls.presetList.onChange = function () {
            if (isRefillingList) return;
            var presetName = getSelectedPresetName(presetControls);
            presetControls.editedPresetName = null;
            updatePresetIcons(presetControls);
            var presetData = findPresetData(presetControls, presetName);
            if (!presetData) return;
            /* 欄への代入ではイベントが来ないので、書き込んだあとでまとめて反映する / Writing the fields fires no events, so refresh once afterwards */
            isApplyingPreset = true;
            applyPresetData(markerControls, presetData);
            updateStyleFieldsEnabled(markerControls);
            isApplyingPreset = false;
            onPresetApplied();
        };
        /* 設定を変えたらリストの選択を外す（名前は保存のときの初期値に残す） / Deselect after an edit; the name stays as the default for saving */
        presetControls.clearPresetSelection = function () {
            if (isApplyingPreset || isRefillingList || !presetControls.presetList.selection) return;
            presetControls.editedPresetName = getSelectedPresetName(presetControls);
            isRefillingList = true;
            presetControls.presetList.selection = null;
            isRefillingList = false;
            updatePresetIcons(presetControls);
        };
        presetControls.presetSaveIcon.onClick = function () {
            var savedName = saveCurrentPreset(markerControls, presetControls);
            if (savedName) refreshPresetList(savedName);
        };
        presetControls.presetOverwriteIcon.onClick = function () {
            var savedName = overwritePreset(markerControls, presetControls);
            if (savedName) refreshPresetList(savedName);
        };
        presetControls.presetDeleteIcon.onClick = function () {
            if (deleteSelectedPreset(presetControls)) refreshPresetList(null);
        };
        presetControls.refreshPresetList = refreshPresetList;
        refreshPresetList(null);
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
     * @param {Object} presetControls - addPresetPanel() の戻り値（builtinPresets 付き）
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
                } else if (settingKey === "customColor" || settingKey === "strokeCustomColor" || settingKey === "textCustomColor") {
                    presetData[settingKey] = serializeColor(createDocumentColor(doc, settingValue.rgb, settingValue.cmyk));
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
        if (colorData.model === "rgb") return createRgbColor(colorData.values);
        if (colorData.model === "cmyk") return createCmykColor(colorData.values);
        if (colorData.model !== "gray") return null;
        var grayColor = new GrayColor();
        grayColor.gray = colorData.values[0];
        return grayColor;
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
            roughen: markerControls.roughenOption.getValue(),
            roughenSizePercent: markerOptions.roughenSize,
            zigzagRidges: markerOptions.zigzagRidges,
            zigzagPoints: markerControls.zigzagPointsRow.getChoice(),
            boxPaddingPt: markerOptions.boxPadding,
            boxPaddingVPt: markerOptions.boxPaddingV,
            boxPaddingLinked: markerControls.boxPaddingLink.value,
            cornerRadiusPt: markerOptions.cornerRadius,
            boxRotation: markerOptions.boxRotation,
            lineGapPt: markerOptions.lineGap,
            dotSizePt: markerOptions.dotSize,
            colorChoice: markerControls.fillColorRow.getChoice(),
            opacity: markerOptions.opacity,
            customColor: serializeColor(markerControls.fillColorRow.getCustomColor()),
            strokeWidthPt: markerOptions.strokeWidth,
            dashGapPt: markerOptions.dashGap,
            verticalOffsetPt: markerOptions.verticalOffset,
            dashStyle: markerOptions.dashStyle,
            strokeCap: markerControls.strokeCapRow.getChoice(),
            stackOrder: markerControls.stackOrderRow.getChoice(),
            isPerChar: markerControls.perCharOption.getValue(),
            squareLineJoins: markerControls.squareLineJoinsOption.getValue(),
            alignLineStarts: markerControls.alignLineStartsOption.getValue(),
            equalLineLengths: markerControls.equalLineLengthsOption.getValue(),
            leadingBased: markerControls.leadingBasedOption.getValue(),
            shapeFill: markerControls.shapeFillOption.getValue(),
            shapeStroke: markerControls.shapeStrokeOption.getValue(),
            strokeColorChoice: markerControls.strokeColorRow.getChoice(),
            strokeCustomColor: serializeColor(markerControls.strokeColorRow.getCustomColor()),
            kerningChoice: markerControls.kerningRow.getChoice(),
            trackingText: markerControls.trackingInput.text,
            kerningBeforeText: markerControls.kerningBeforeInput.text,
            kerningAfterText: markerControls.kerningAfterInput.text,
            kerningFromPadding: markerControls.kerningFromPaddingOption.getValue(),
            kerningAdjustBefore: markerControls.kerningAdjustments.before,
            kerningAdjustAfter: markerControls.kerningAdjustments.after,
            dashSegmentsText: markerControls.dashSegmentsInput.text,
            textColorChoice: markerControls.textFillColorRow.getChoice(),
            textCustomColor: serializeColor(markerControls.textFillColorRow.getCustomColor())
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
        applyPresetLineStyle(markerControls, presetData);
        applyPresetShapePaint(markerControls, presetData);
        applyPresetNumberFields(markerControls, presetData);
        applyPresetChoices(markerControls, presetData);
        applyPresetTextFormat(markerControls, presetData);
    }

    /**
     * プリセットの［種類］をダイアログボックスに書き込む（旧版の種類は今の種類と［1文字ずつ］に読み替える）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetLineStyle(markerControls, presetData) {
        /* 旧版の種類を今の種類と［1文字ずつ］［図形］に読み替える（［直線］は［マーカー］と同じ位置、［囲み罫］は線の［長方形］）
           Map former styles to the current style, Per Character and Shape (Straight sits where Marker does; Frame is a stroked Box) */
        var formerStyles = {
            straight: { lineStyle: "marker" },
            charBox: { lineStyle: "box", isPerChar: true },
            charCircle: { lineStyle: "circle", isPerChar: true },
            frame: { lineStyle: "box" },
            roughMarker: { lineStyle: "marker" }
        };
        var formerStyle = formerStyles.hasOwnProperty(presetData.lineStyle) ? formerStyles[presetData.lineStyle] : null;
        if (formerStyle) {
            markerControls.lineStyleRow.setChoice(formerStyle.lineStyle);
            if (formerStyle.isPerChar) markerControls.perCharOption.setValue(true);
        } else if (presetData.lineStyle) {
            markerControls.lineStyleRow.setChoice(presetData.lineStyle);
        }
    }

    /**
     * プリセットの数値（長さは pt から欄の単位へ換算）をダイアログボックスに書き込む
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetNumberFields(markerControls, presetData) {
        /* 数値の欄：[プリセットのキー, 入力欄, pt から換算する単位（null なら換算しない）]
           Number fields: [preset key, input, unit converted from pt (null for no conversion)] */
        var numberFields = [
            ["shearAngle", markerControls.shearAngleInput, null],
            ["zigzagSizePt", markerControls.zigzagSizeInput, markerControls.rulerUnit],
            ["zigzagRidges", markerControls.zigzagRidgesInput, null],
            ["roughenSizePercent", markerControls.roughenSizeInput, null],
            ["boxPaddingPt", markerControls.boxPaddingInput, markerControls.rulerUnit],
            /* 旧版のプリセットは［余白（縦）］も［余白］の値にする / Presets from earlier versions use Padding for both */
            [(typeof presetData.boxPaddingVPt === "number") ? "boxPaddingVPt" : "boxPaddingPt", markerControls.boxPaddingVInput, markerControls.rulerUnit],
            ["cornerRadiusPt", markerControls.cornerRadiusInput, markerControls.rulerUnit],
            ["boxRotation", markerControls.boxRotationInput, null],
            ["lineGapPt", markerControls.lineGapInput, markerControls.rulerUnit],
            ["dotSizePt", markerControls.dotSizeInput, markerControls.rulerUnit],
            ["opacity", markerControls.opacityInput, null],
            ["strokeWidthPt", markerControls.strokeWidthInput, markerControls.strokeUnit],
            ["dashGapPt", markerControls.dashGapInput, markerControls.strokeUnit],
            ["verticalOffsetPt", markerControls.verticalOffsetInput, markerControls.rulerUnit]
        ];
        for (var i = 0; i < numberFields.length; i++) {
            var presetValue = presetData[numberFields[i][0]];
            if (typeof presetValue !== "number") continue;
            var fieldUnit = numberFields[i][2];
            setMarkerFieldValue(numberFields[i][1], fieldUnit ? ptToUnit(presetValue, fieldUnit) : presetValue);
        }
        /* 旧版の不透明度は［文字色］のときだけ効いた / the former opacity applied only to Text Color */
        if (typeof presetData.opacity !== "number" && typeof presetData.textColorOpacity === "number") {
            setMarkerFieldValue(markerControls.opacityInput, presetData.colorChoice === "textColor" ? presetData.textColorOpacity : 100);
        }
        /* v1.5.0 の［詳細］は［折り返し］にまとめた（［ギザギザ］以外で手書き風のときだけ読む） / v1.5.0 Detail merged into Ridges (read only for Hand-Drawn outside Zigzag) */
        if (presetData.roughen === true && typeof presetData.roughenDetail === "number" && presetData.lineStyle !== "zigzag") {
            setMarkerFieldValue(markerControls.zigzagRidgesInput, presetData.roughenDetail);
        }
    }

    /**
     * プリセットの選択肢・チェックボックス・連動（［塗り］のカラー・線の種類・行長・行の高さ・ラフなど）をダイアログボックスに書き込む
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetChoices(markerControls, presetData) {
        var customColor = deserializeColor(presetData.customColor);
        if (customColor) markerControls.fillColorRow.setCustomColor(customColor);
        if (presetData.colorChoice) markerControls.fillColorRow.setChoice(presetData.colorChoice);
        if (presetData.zigzagPoints) markerControls.zigzagPointsRow.setChoice(presetData.zigzagPoints);
        if (presetData.dashStyle) markerControls.dashStyleRow.setChoice(presetData.dashStyle);
        if (presetData.strokeCap) markerControls.strokeCapRow.setChoice(presetData.strokeCap);
        if (presetData.stackOrder) markerControls.stackOrderRow.setChoice(presetData.stackOrder);
        if (typeof presetData.isPerChar === "boolean") markerControls.perCharOption.setValue(presetData.isPerChar);
        if (typeof presetData.squareLineJoins === "boolean") markerControls.squareLineJoinsOption.setValue(presetData.squareLineJoins);
        /* 前の版のプリセットには［行頭］が無いのでオフにする / Presets from earlier versions have no Align Starts, so it goes off */
        markerControls.alignLineStartsOption.setValue(presetData.alignLineStarts === true && presetData.equalLineLengths !== true);
        if (typeof presetData.equalLineLengths === "boolean") markerControls.equalLineLengthsOption.setValue(presetData.equalLineLengths);
        /* 縦横が別の値なら連動を切る（旧版のプリセットは連動） / Unlink when the two paddings differ (earlier presets stay linked) */
        if (typeof presetData.boxPaddingLinked === "boolean") setLinkToggleValue(markerControls.boxPaddingLink, presetData.boxPaddingLinked);
        else if (typeof presetData.boxPaddingPt === "number") setLinkToggleValue(markerControls.boxPaddingLink, true);
        if (typeof presetData.leadingBased === "boolean") markerControls.leadingBasedOption.setValue(presetData.leadingBased);
        /* 旧版の種類［手書き風］は［マーカー］の［手書き風］オン。持っていないプリセットはオフ / The former Hand-Drawn style is Marker with Hand-Drawn on; presets without it turn it off */
        markerControls.roughenOption.setValue(presetData.roughen === true || presetData.lineStyle === "roughMarker");
    }

    /**
     * プリセットの［文字］の設定（テキストカラー・カーニング・トラッキング・余白に合わせてカーニングを自動調整）をダイアログボックスに書き込む
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetTextFormat(markerControls, presetData) {
        if (presetData.kerningChoice) markerControls.kerningRow.setChoice(presetData.kerningChoice);
        /* 前の版のプリセットには無いので初期値にする / Presets from earlier versions lack it, so use the default */
        markerControls.kerningFromPaddingOption.setValue(typeof presetData.kerningFromPadding === "boolean" ? presetData.kerningFromPadding : DEFAULT_KERNING_FROM_PADDING);
        /* 微調整は文字サイズに左右されないよう、計算した値との差で持つ / Adjustments are kept as differences so they do not depend on the type size */
        markerControls.kerningAdjustments.before = (typeof presetData.kerningAdjustBefore === "number") ? presetData.kerningAdjustBefore : 0;
        markerControls.kerningAdjustments.after = (typeof presetData.kerningAdjustAfter === "number") ? presetData.kerningAdjustAfter : 0;
        /* 空欄にできる欄は、保存した文字列をそのまま入れる（［分割数］もここで扱う） / Blank-able fields take the saved text as is (Segments too) */
        var textFields = [
            ["trackingText", markerControls.trackingInput],
            ["kerningBeforeText", markerControls.kerningBeforeInput],
            ["kerningAfterText", markerControls.kerningAfterInput],
            ["dashSegmentsText", markerControls.dashSegmentsInput]
        ];
        for (var j = 0; j < textFields.length; j++) {
            var presetText = presetData[textFields[j][0]];
            if (typeof presetText !== "string") continue;
            textFields[j][1].text = presetText;
            textFields[j][1].lastValidText = presetText;
        }
        var textCustomColor = deserializeColor(presetData.textCustomColor);
        if (textCustomColor) markerControls.textFillColorRow.setCustomColor(textCustomColor);
        /* 旧版の［白抜きに］は［白］か［変更しない］に読み替える / The former White checkbox maps to White or No Change */
        if (presetData.textColorChoice) markerControls.textFillColorRow.setChoice(presetData.textColorChoice);
        else if (typeof presetData.whiteText === "boolean") markerControls.textFillColorRow.setChoice(presetData.whiteText ? "white" : "keepTextColor");
    }

    /**
     * プリセットの［塗り］［線］の［描画］と［線］のカラーをダイアログボックスに書き込む。
     * ［塗り］［線］の無い旧版の長方形の仲間は塗りだけ、旧版の［囲み罫］は線だけにする。
     * ［線］のカラーの無い旧版の［囲み罫］と線で引く種類（［塗りつぶし］を除く）は、旧版の［カラー］を線のカラーに使う
     * @param {Object} markerControls - addMarkerControls() の戻り値（種類は書き込み済み）
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetShapePaint(markerControls, presetData) {
        var isFormerFrame = (presetData.lineStyle === "frame");
        var isShapePreset = isShapeStyle(markerControls.lineStyleRow.getChoice());
        if (typeof presetData.shapeFill === "boolean") markerControls.shapeFillOption.setValue(presetData.shapeFill);
        else if (isShapePreset) markerControls.shapeFillOption.setValue(!isFormerFrame);
        if (typeof presetData.shapeStroke === "boolean") markerControls.shapeStrokeOption.setValue(presetData.shapeStroke);
        else if (isShapePreset) markerControls.shapeStrokeOption.setValue(isFormerFrame);
        var lineStyle = markerControls.lineStyleRow.getChoice();
        var usesFormerColor = !presetData.strokeColorChoice
            && (isFormerFrame || (!isShapePreset && lineStyle !== "fill" && lineStyle !== "emphasisDots"));
        var strokeCustomColor = deserializeColor(usesFormerColor ? presetData.customColor : presetData.strokeCustomColor);
        if (strokeCustomColor) markerControls.strokeColorRow.setCustomColor(strokeCustomColor);
        var strokeColorChoice = usesFormerColor ? presetData.colorChoice : presetData.strokeColorChoice;
        if (strokeColorChoice) markerControls.strokeColorRow.setChoice(strokeColorChoice);
    }

    /**
     * プリセットのリストを作り直し、指定の名前を選ぶ。
     * 最初からあるプリセット、保存したプリセット（名前順）の順に並べる
     * @param {Object} presetControls - addPresetPanel() の戻り値（builtinPresets 付き）
     * @param {string|null} selectedName - 選んでおくプリセット名（null なら選ばない）
     * @returns {void}
     */
    function fillPresetList(presetControls, selectedName) {
        var presetList = presetControls.presetList;
        var presetNames = presetControls.builtinPresets.names.concat(getSavedPresetNames(presetControls.builtinPresets.map));
        var selectedIndex = -1;
        presetList.removeAll();
        for (var i = 0; i < presetNames.length; i++) {
            presetList.add("item", presetNames[i]);
            if (presetNames[i] === selectedName) selectedIndex = i;
        }
        presetList.selection = (selectedIndex >= 0) ? selectedIndex : null;
        if (selectedIndex >= 0) presetControls.editedPresetName = null;
        updatePresetIcons(presetControls);
    }

    /**
     * 上書き保存するプリセット名を返す（選んでいるもの。設定を変えて選択が外れたなら、その前に選んでいたもの）
     * @param {Object} presetControls - addPresetPanel() の戻り値
     * @returns {string|null} プリセット名。どちらも無ければ null
     */
    function getOverwriteTargetName(presetControls) {
        return getSelectedPresetName(presetControls) || presetControls.editedPresetName || null;
    }

    /**
     * 上書き保存・削除のアイコンを、選んでいるプリセットに合わせて有効／無効にし、上書き保存のツールチップに名前を入れる
     * @param {Object} presetControls - addPresetPanel() の戻り値
     * @returns {void}
     */
    function updatePresetIcons(presetControls) {
        var targetName = getOverwriteTargetName(presetControls);
        setIconButtonEnabled(presetControls.presetOverwriteIcon, !!targetName);
        presetControls.presetOverwriteIcon.helpTip = targetName
            ? getLabel("tooltip.presetOverwrite", { name: targetName })
            : getLabel("tooltip.presetOverwriteNone");
        setIconButtonEnabled(presetControls.presetDeleteIcon, isSavedPreset(getSelectedPresetName(presetControls)));
    }

    /**
     * 選んでいるプリセット名を返す
     * @param {Object} presetControls - addPresetPanel() の戻り値
     * @returns {string|null} プリセット名。選んでいなければ null
     */
    function getSelectedPresetName(presetControls) {
        var presetSelection = presetControls.presetList.selection;
        return presetSelection ? presetSelection.text : null;
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
        if (!presetName) return null;
        return presetName;
    }

    /**
     * 今の設定を名前を付けて保存する（同じ名前は、最初からあるものも含めて確認してから上書き）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetControls - addPresetPanel() の戻り値（選んでいるプリセット名を名前の入力欄に入れておく）
     * @returns {string|null} 保存したプリセット名。保存しなかったら null
     */
    function saveCurrentPreset(markerControls, presetControls) {
        var presetName = showPresetNameDialog(getOverwriteTargetName(presetControls) || "");
        if (!presetName) return null;
        var presetMap = loadPresetMap();
        var isExistingName = presetMap.hasOwnProperty(presetName) || presetControls.builtinPresets.map.hasOwnProperty(presetName);
        if (isExistingName && !confirm(getLabel("confirm.presetOverwrite", { name: presetName }), true)) return null;
        presetMap[presetName] = collectPresetData(markerControls);
        return savePresetMap(presetMap) ? presetName : null;
    }

    /**
     * 選んでいるプリセット（設定を変える前に選んでいたもの）に、名前を尋ねず確認もせずに今の設定を上書き保存する。
     * 最初からあるプリセットなら、同じ名前の保存したプリセットになる（削除すると最初の設定に戻る）
     * @param {Object} markerControls - addMarkerControls() の戻り値
     * @param {Object} presetControls - addPresetPanel() の戻り値
     * @returns {string|null} 保存したプリセット名。保存しなかったら null
     */
    function overwritePreset(markerControls, presetControls) {
        var presetName = getOverwriteTargetName(presetControls);
        if (!presetName) return null;
        var presetMap = loadPresetMap();
        presetMap[presetName] = collectPresetData(markerControls);
        return savePresetMap(presetMap) ? presetName : null;
    }

    /**
     * 選んでいるプリセットを確認してから削除する（最初からあるものは削除しない）
     * @param {Object} presetControls - addPresetPanel() の戻り値
     * @returns {boolean} 削除したら true
     */
    function deleteSelectedPreset(presetControls) {
        var presetName = getSelectedPresetName(presetControls);
        if (!isSavedPreset(presetName) || !confirm(getLabel("confirm.presetDelete", { name: presetName }), true)) return false;
        var presetMap = loadPresetMap();
        delete presetMap[presetName];
        return savePresetMap(presetMap);
    }

    /**
     * プリセットの集まりを保存する（保存できなければ知らせる）
     * @param {Object} presetMap - プリセット名をキーにした設定の集まり
     * @returns {boolean} 保存できたら true
     */
    function savePresetMap(presetMap) {
        if (presetSettingsStore.save(presetMap)) return true;
        alert(getLabel("alert.presetSaveFailed"));
        return false;
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
            var itemType = "";
            /* 読めない参照は字形として返し、isGlyphReadable() で外す / Unreadable references are returned as glyphs and filtered by isGlyphReadable() */
            try { itemType = currentItem.typename; } catch (e) { /* 読めない / unreadable */ }
            if (itemType !== "GroupItem") {
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
        var fillSource = getGlyphFillSource(glyphItem);
        return !!fillSource && fillSource.filled && fillSource.fillColor.typename !== "NoColor";
    }

    /**
     * 字形の塗りを持つパスを返す（複合パスは先頭の構成パス）
     * @param {PageItem} glyphItem - 字形のパスか複合パス
     * @returns {PathItem} 塗りを読むパス
     */
    function getGlyphFillSource(glyphItem) {
        return (glyphItem.typename === "CompoundPathItem") ? glyphItem.pathItems[0] : glyphItem;
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
     * 字形の位置と塗りを読めるかを返す。アウトライン化の直後は、まれに読むと例外（MRAP）になる参照が混じる
     * @param {PageItem} glyphItem - 字形のパスか複合パス
     * @returns {boolean} 読めれば true
     */
    function isGlyphReadable(glyphItem) {
        try {
            glyphItem.geometricBounds;
            getGlyphFillSource(glyphItem).filled;
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 下書き用の複製をアウトライン化して、字形を平らな配列で返す。
     * 読めない字形が混じったら作り直し（OUTLINE_ATTEMPT_LIMIT 回まで）、それでも残るものは外す
     * @param {Function} createDraft - 下書き用の複製を作って返す関数（呼ぶたびに新しく作る）
     * @returns {{group: GroupItem, glyphItems: PageItem[]}} アウトラインのグループ（片付けは呼び出し側）と、読める字形
     */
    function outlineDraftGlyphs(createDraft) {
        for (var attempt = 1; ; attempt++) {
            /* createOutline() は複製を消費するので、片付けるのは戻り値のグループ / createOutline() consumes the copy, so remove the returned group */
            var outlineGroup = createDraft().createOutline();
            var glyphItems = flattenOutlineGlyphs(outlineGroup);
            var readableItems = [];
            for (var i = 0; i < glyphItems.length; i++) {
                if (isGlyphReadable(glyphItems[i])) readableItems.push(glyphItems[i]);
            }
            if (readableItems.length === glyphItems.length || attempt >= OUTLINE_ATTEMPT_LIMIT) {
                return { group: outlineGroup, glyphItems: readableItems };
            }
            outlineGroup.remove();
        }
    }

    /**
     * 文字番号を、ドキュメントのカラーモードの色に置き換える（RGB は 256 進、CMYK は 100 進で C・M・Y に振る）
     * @param {number} charIndex - 文字番号（0始まり）
     * @param {boolean} isRgbDocument - RGB のドキュメントか
     * @returns {Color} 文字番号を表す色
     */
    function createCharIndexColor(charIndex, isRgbDocument) {
        if (isRgbDocument) return createRgbColor([charIndex % 256, Math.floor(charIndex / 256) % 256, Math.floor(charIndex / 65536) % 256]);
        return createCmykColor([charIndex % 100, Math.floor(charIndex / 100) % 100, Math.floor(charIndex / 10000) % 100, 0]);
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
     * 字面を測るための下書き用の複製を、元のテキストの直後に作る（すべての文字の塗りを外す）
     * @param {TextFrame} sourceFrame - 元のテキストフレーム
     * @returns {TextFrame} 複製
     */
    function createDraftFrame(sourceFrame) {
        var draftFrame = sourceFrame.duplicate(sourceFrame, ElementPlacement.PLACEAFTER);
        draftFrame.textRange.characterAttributes.fillColor = new NoColor();
        return draftFrame;
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
        var draftFill = new GrayColor();
        draftFill.gray = 100;
        var outlinedDraft = outlineDraftGlyphs(function () {
            var draftFrame = createDraftFrame(sourceFrame);
            /* 選んだ範囲だけ1つの TextRange として塗り直す / Refill only the selected span as one TextRange */
            fillCharacterSpan(draftFrame, spanStart, spanLength, draftFill);
            return draftFrame;
        });
        var glyphBounds = getFilledGlyphsExtent(outlinedDraft.glyphItems);
        outlinedDraft.group.remove();
        return glyphBounds;
    }

    /**
     * 選んだ範囲の字面の範囲を1文字ずつ測る。
     * 下書き用の複製で選んだ範囲の文字を文字番号ごとに別の色で塗り、1回のアウトライン化で字形を文字に振り分ける
     * @param {TextFrame} sourceFrame - 元のテキストフレーム
     * @param {number} spanStart - 選んだ範囲の先頭の文字番号（フレーム内、0始まり）
     * @param {number} spanLength - 選んだ範囲の文字数
     * @returns {Object[]} 字面の範囲 { left, top, right, bottom, charIndex } の配列（文字の順。スペースなど字形の無い文字は除く。
     *     charIndex は範囲の先頭からの文字番号）
     */
    function measureCharGlyphBounds(sourceFrame, spanStart, spanLength) {
        /* 色はドキュメントのカラーモードのまま読み戻せるように、同じモードで作る / Use the document color mode so the colors read back unchanged */
        var isRgbDocument = (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB);
        var outlinedDraft = outlineDraftGlyphs(function () {
            var draftFrame = createDraftFrame(sourceFrame);
            for (var i = 0; i < spanLength; i++) {
                draftFrame.characters[spanStart + i].characterAttributes.fillColor = createCharIndexColor(i, isRgbDocument);
            }
            return draftFrame;
        });
        var glyphItems = outlinedDraft.glyphItems;
        var boundsByIndex = [];
        for (var j = 0; j < glyphItems.length; j++) {
            if (!hasGlyphFill(glyphItems[j])) continue;
            var charIndex = readCharIndexColor(getGlyphFillSource(glyphItems[j]).fillColor, isRgbDocument);
            if (charIndex < 0 || charIndex >= spanLength) continue;
            boundsByIndex[charIndex] = extendGlyphExtent(boundsByIndex[charIndex] || null, glyphItems[j].geometricBounds);
        }
        outlinedDraft.group.remove();

        var charBoundsList = [];
        for (var k = 0; k < spanLength; k++) {
            if (!boundsByIndex[k]) continue;
            boundsByIndex[k].charIndex = k;
            charBoundsList.push(boundsByIndex[k]);
        }
        return charBoundsList;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ［ジグザグ］効果の XML を返す（大きさは絶対値）
     * @param {number} amountPt - 大きさ（pt）
     * @param {number} ridges - 線分あたりの折り返し（0 ならアンカーポイントを交互にずらすだけ）
     * @param {boolean} isSmooth - ポイントを滑らかにするか（false なら直線的に）
     * @returns {string} LiveEffect の XML
     */
    function createZigzagXml(amountPt, ridges, isSmooth) {
        return '<LiveEffect name="Adobe Zigzag"><Dict data="R amount ' + amountPt +
            ' R relAmount 0 R absoluteness 1 R ridges ' + ridges + ' R roundness ' + (isSmooth ? 1 : 0) + ' "/></LiveEffect>';
    }

    /**
     * ［ラフ］効果の XML を返す（サイズはパーセント）。サイズ 0 なら形は変えず、詳細の数だけポイントを足す
     * @param {number} sizePercent - サイズ（線の大きさに対する％）
     * @param {number} detail - 詳細（1インチあたりの数）
     * @param {boolean} isSmooth - ポイントを丸くするか（false ならギザギザ）
     * @returns {string} LiveEffect の XML
     */
    function createRoughenXml(sizePercent, detail, isSmooth) {
        return '<LiveEffect name="Adobe Roughen"><Dict data="R asiz 0 R size ' + sizePercent + ' R absoluteness 0 R dtal ' + detail +
            ' R roundness ' + (isSmooth ? 1 : 0) + ' "/></LiveEffect>';
    }

    /**
     * ［手書き風］がオンなら、［ラフ］効果（ポイントはギザギザ）を掛ける
     * @param {PageItem} targetItem - 線や図形
     * @param {Object} markerOptions - readMarkerOptions() の戻り値（isRoughen・roughenSize・zigzagRidges）
     * @returns {void}
     */
    function applyRoughenEffect(targetItem, markerOptions) {
        if (!markerOptions.isRoughen) return;
        targetItem.applyEffect(createRoughenXml(markerOptions.roughenSize, markerOptions.zigzagRidges, false));
    }

    /**
     * 上辺と下辺（縦の線は左辺と右辺）を逆向きに半分ずつずらす［パスの自由変形］効果の XML を返す（傾けても中心は動かない）。
     * 4隅は線の外形（線幅を含む範囲）を 0〜1 にした座標で、順序は [左上, 右上, 左下, 右下]
     * @param {number} shiftRatio - 横の線は下辺に対して上辺を右へ、縦の線は左辺に対して右辺を下へずらす量
     *     （横の線は外形の幅、縦の線は外形の高さに対する比率。マイナスで逆向き）
     * @param {boolean} isVertical - 縦の線か
     * @returns {string} LiveEffect の XML
     */
    function createShearXml(shiftRatio, isVertical) {
        var halfShift = shiftRatio / 2;
        var sourceCorners = [[0, 0], [1, 0], [0, 1], [1, 1]];
        /* 縦の線は横の線を時計回りに 90° 回した形（上辺→右辺、右→下） / A vertical line is a horizontal one turned 90° clockwise */
        var destCorners = isVertical
            ? [[0, -halfShift], [1, halfShift], [0, 1 - halfShift], [1, 1 + halfShift]]
            : [[halfShift, 0], [1 + halfShift, 0], [-halfShift, 1], [1 - halfShift, 1]];
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
     * パスに［効果］＞［パス］＞［パスのアウトライン］を掛ける（XML の applyEffect。Dict の data は空白1つ）
     * @param {PathItem} targetPath - 対象のパス
     * @returns {void}
     */
    function applyOutlineStrokeEffect(targetPath) {
        targetPath.applyEffect('<LiveEffect name="Adobe Outline Stroke"><Dict data=" "/></LiveEffect>');
    }

    /**
     * ［重ね順］の位置（元のテキストの背面か前面）へマーカーの線を引く。
     * 行（縦組みは列）ごとに、その行の字面に合わせて引く（縦組みは縦の線）。2本以上ならグループにする
     * @param {Object} markerTarget - collectMarkerTargets() の要素（selectTargetLayout() で選んだあと）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {PathItem|GroupItem} 引いた線（2本以上ならグループ）
     */
    function placeMarkerLine(markerTarget, markerOptions) {
        var sourceFrame = markerTarget.sourceFrame;
        var isVertical = markerTarget.isVertical;
        var lineBoundsList = getMarkerLineBoundsList(markerTarget, markerOptions.isLeadingBased);
        if (markerOptions.equalLineLengths || markerOptions.alignLineStarts) lineBoundsList = alignLineEdges(lineBoundsList, isVertical, markerOptions.equalLineLengths);
        var markerLines = [];
        for (var i = 0; i < lineBoundsList.length; i++) {
            var lineBounds = lineBoundsList[i];
            /* ［塗りつぶし］は行の高さ（縦組みは幅）ちょうどの線にする / Fill uses a line exactly as wide as the line height (width in vertical text) */
            var glyphThickness = isVertical ? (lineBounds.right - lineBounds.left) : (lineBounds.top - lineBounds.bottom);
            var lineWidth = (markerOptions.lineStyle === "fill") ? glyphThickness : markerOptions.strokeWidth;
            var linePositions = getMarkerLinePositions(lineBounds, markerOptions, lineWidth, isVertical);
            for (var k = 0; k < linePositions.length; k++) {
                markerLines.push(drawMarkerLinePath(markerTarget, markerOptions, lineBounds, linePositions[k], lineWidth));
            }
        }
        /* 効果を掛け終えてから元のテキストの隣（グループの中かもしれない）へ移す / Move next to the text (possibly inside a group) after the effects are applied */
        var placedItem = markerLines[0];
        if (markerLines.length > 1) {
            placedItem = sourceFrame.layer.groupItems.add();
            for (var j = 0; j < markerLines.length; j++) markerLines[j].move(placedItem, ElementPlacement.PLACEATEND);
        }
        moveToStackPosition(placedItem, markerTarget, markerOptions);
        return placedItem;
    }

    /**
     * 線を引く字面の範囲を、行（縦組みは列）ごとに返す。1文字ずつ測れなかったときは選んだ範囲全体の字面1つ
     * @param {Object} markerTarget - collectMarkerTargets() の要素（selectTargetLayout() で選んだあと）
     * @param {boolean} isLeadingBased - 行の高さ（縦組みは幅）を行送りにするか
     * @returns {Object[]} 字面の範囲 { left, top, right, bottom } の配列
     */
    function getMarkerLineBoundsList(markerTarget, isLeadingBased) {
        var textLines = getTargetTextLines(markerTarget, isLeadingBased);
        return (textLines.length > 0) ? textLines : [markerTarget.activeLayout.glyphBounds];
    }

    /**
     * 行ごとの範囲の行頭（横組みは左端、縦組みは上端）を全行の外側にそろえる。alignsBothEnds なら行末（右端・下端）もそろえ、いちばん長い行の長さにする
     * @param {Object[]} lineBoundsList - 行ごとの範囲 { left, top, right, bottom }
     * @param {boolean} isVertical - 縦組みか
     * @param {boolean} alignsBothEnds - 行末もそろえるか
     * @returns {Object[]} そろえた範囲（新しいオブジェクト）
     */
    function alignLineEdges(lineBoundsList, isVertical, alignsBothEnds) {
        var outerExtent = null;
        for (var i = 0; i < lineBoundsList.length; i++) {
            var lineBounds = lineBoundsList[i];
            outerExtent = extendGlyphExtent(outerExtent, [lineBounds.left, lineBounds.top, lineBounds.right, lineBounds.bottom]);
        }
        var equalizedList = [];
        for (var j = 0; j < lineBoundsList.length; j++) {
            var sourceBounds = lineBoundsList[j];
            equalizedList.push(isVertical
                ? { left: sourceBounds.left, top: outerExtent.top, right: sourceBounds.right, bottom: alignsBothEnds ? outerExtent.bottom : sourceBounds.bottom }
                : { left: outerExtent.left, top: sourceBounds.top, right: alignsBothEnds ? outerExtent.right : sourceBounds.right, bottom: sourceBounds.bottom });
        }
        return equalizedList;
    }

    /**
     * 線の中心の位置（横組みは Y、縦組みは X 座標）を返す。
     * ［塗りつぶし］［取り消し線］［二重取り消し線］は字面の中央、［アンダーライン］は字面の下端（縦組みは右端）に線の端、
     * ほかは字面の下端（縦組みは右端）を線の中心にし、［位置調整］の分ずらす（縦組みは左右に、プラスで左）。
     * ［二重取り消し線］は中央から両側に［線の間隔］の半分と線幅の半分ずつ離した2本
     * @param {Object} glyphBounds - 字面の範囲 { left, top, right, bottom }
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @param {number} lineWidth - 線幅（pt）
     * @param {boolean} isVertical - 縦組みか
     * @returns {number[]} 線の中心の座標（1本か2本）
     */
    function getMarkerLinePositions(glyphBounds, markerOptions, lineWidth, isVertical) {
        var lineStyle = markerOptions.lineStyle;
        /* 縦組みの下線は字の右側に付くので、横組みの下端を右端に置き換える / Vertical underlines sit on the right, so the bottom edge becomes the right edge */
        var underSide = isVertical ? glyphBounds.right : glyphBounds.bottom;
        var outwardSign = isVertical ? 1 : -1;
        var linePosition = underSide;
        if (lineStyle === "fill" || lineStyle === "strikethrough" || lineStyle === "doubleStrikethrough") {
            linePosition = isVertical ? (glyphBounds.left + glyphBounds.right) / 2 : (glyphBounds.top + glyphBounds.bottom) / 2;
        } else if (lineStyle === "underline") {
            linePosition = underSide + outwardSign * lineWidth / 2;
        }
        /* ［位置調整］の分だけずらす（プラスで上。縦組みは左） / Shift by the offset (positive is up; left in vertical text) */
        linePosition += isVertical ? -markerOptions.verticalOffset : markerOptions.verticalOffset;
        if (lineStyle !== "doubleStrikethrough") return [linePosition];
        var lineOffset = (lineWidth + markerOptions.lineGap) / 2;
        return [linePosition + lineOffset, linePosition - lineOffset];
    }

    /**
     * レイヤーの直下に1本の線を描き、種類に応じた効果を掛ける（元のテキストの隣へは移さない）。
     * 横組みは左から右へ、縦組みは上から下へ引く
     * @param {Object} markerTarget - collectMarkerTargets() の要素（isVertical）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @param {Object} lineBounds - 線を引く字面の範囲 { left, top, right, bottom }
     * @param {number} linePosition - 線の中心の座標（横組みは Y、縦組みは X）
     * @param {number} lineWidth - 線幅（pt）
     * @returns {PathItem} 線
     */
    function drawMarkerLinePath(markerTarget, markerOptions, lineBounds, linePosition, lineWidth) {
        var lineStyle = markerOptions.lineStyle;
        var isVertical = markerTarget.isVertical;
        var markerLine = markerTarget.sourceFrame.layer.pathItems.add();
        markerLine.setEntirePath(isVertical
            ? [[linePosition, lineBounds.top], [linePosition, lineBounds.bottom]]
            : [[lineBounds.left, linePosition], [lineBounds.right, linePosition]]);
        markerLine.filled = false;
        markerLine.stroked = true;
        /* ［塗りつぶし］は面に見えるので［塗り］のカラー、ほかの線は［線］のカラー / Fill looks like an area and takes Fill Color; other lines take Stroke Color */
        markerLine.strokeColor = (lineStyle === "fill") ? getMarkerFillColor(markerTarget, markerOptions) : getMarkerStrokeColor(markerTarget, markerOptions);
        markerLine.strokeWidth = lineWidth;
        applyLineDashes(markerLine, markerOptions);
        markerLine.opacity = markerOptions.opacity;
        /* 効果は分割せずに残す。後から掛けたものほど下に並び、後で処理される
           Effects stay live; later ones go lower in the list and are processed later */
        /* ギザギザは中心線を折ってからアウトラインにする。先に［ラフ］効果で1インチあたり［折り返し］の数だけポイントを足し、
           ［ジグザグ］効果は折り返し 0 でそのポイントを交互にずらす（線の長さにかかわらず、山の間隔がそろう）
           Zigzag folds the center line before it is outlined: Roughen first adds Ridges points per inch,
           then Zig Zag with 0 ridges moves those points alternately, so the ridge pitch does not depend on the line length */
        if (lineStyle === "zigzag") {
            markerLine.applyEffect(createRoughenXml(0, markerOptions.zigzagRidges, false));
            markerLine.applyEffect(createZigzagXml(markerOptions.zigzagSize, 0, markerOptions.isZigzagSmooth));
        }
        /* ［手書き風］は中心線を揺らしてからアウトラインにする（縁ではなく線そのものが揺れる）
           Hand-Drawn roughens the center line before it is outlined, so the line itself wobbles rather than its edges */
        applyRoughenEffect(markerLine, markerOptions);
        applyOutlineStrokeEffect(markerLine);
        /* シアーはアウトラインにした面（線幅の太さがある）を変形する / Shear distorts the outlined shape, which has the stroke's thickness */
        if (lineStyle === "straightShear") {
            /* 両辺のずれの合計は 線幅 × tan(角度)。比率は線の外形の長さ（横の線は幅、縦の線は高さ）に対して取る
               The two edges differ by stroke width x tan(angle), as a ratio of the outline length */
            var outlineBounds = markerLine.visibleBounds;
            var outlineLength = isVertical ? (outlineBounds[1] - outlineBounds[3]) : (outlineBounds[2] - outlineBounds[0]);
            var edgeShiftPt = lineWidth * Math.tan(markerOptions.shearAngle * Math.PI / 180);
            markerLine.applyEffect(createShearXml(edgeShiftPt / outlineLength, isVertical));
        }
        return markerLine;
    }

    /**
     * 1文字ずつの字面の範囲を行（縦組みは列）に分ける。
     * 文字の中心が直前の文字より行頭側へ（直前の字面の半分を超えて）戻ったか（横組みは左、縦組みは上）、
     * それまでの行の字面より次の行の側へ出たら（横組みは下端より下、縦組みは左端より左）、次の行とみなす。
     * 行頭の「…」「-」のような小さい字形の上下にはみ出しても、縦中横の数字が横に並んでも分けない
     * @param {Object[]} charBoundsList - measureCharGlyphBounds() の戻り値（文字の順）
     * @param {boolean} isVertical - 縦組みか
     * @returns {Object[]} 行 { left, top, right, bottom, chars }（範囲は行の字面全体、chars はその行の文字の範囲）
     */
    function groupCharBoundsByLine(charBoundsList, isVertical) {
        var textLines = [];
        var currentLine = null;
        var previousCenterX = 0;
        var previousCenterY = 0;
        var previousHalfWidth = 0;
        var previousHalfHeight = 0;
        for (var i = 0; i < charBoundsList.length; i++) {
            var charBounds = charBoundsList[i];
            var centerX = (charBounds.left + charBounds.right) / 2;
            var centerY = (charBounds.top + charBounds.bottom) / 2;
            var isNewLine = !currentLine;
            /* 縦中横の「90」は「0」の中心がわずかに上に来るので、字面の半分までの戻りは同じ行にする
               In tate-chu-yoko "90" the "0" sits slightly higher, so a step back within half a glyph stays on the line */
            if (currentLine && isVertical) {
                isNewLine = (centerY > previousCenterY + previousHalfHeight || centerX < currentLine.left);
            } else if (currentLine) {
                isNewLine = (centerX < previousCenterX - previousHalfWidth || centerY < currentLine.bottom);
            }
            previousCenterX = centerX;
            previousCenterY = centerY;
            previousHalfWidth = (charBounds.right - charBounds.left) / 2;
            previousHalfHeight = (charBounds.top - charBounds.bottom) / 2;
            if (isNewLine) {
                currentLine = { left: charBounds.left, top: charBounds.top, right: charBounds.right, bottom: charBounds.bottom, chars: [] };
                textLines.push(currentLine);
            }
            extendGlyphExtent(currentLine, [charBounds.left, charBounds.top, charBounds.right, charBounds.bottom]);
            currentLine.chars.push(charBounds);
        }
        return textLines;
    }

    /**
     * 図形を描く範囲（余白を足す前）を返す。
     * 行ごとなら行の字面全体。1文字ずつなら文字ごとに、行の字面の高さ（縦組みは列の幅）にそろえ、
     * もう一方の向きは字面の大きさか行の高さの大きいほうで字面の中央に置く
     * @param {Object[]} textLines - getTargetTextLines() の戻り値
     * @param {boolean} isPerChar - 1文字ずつにするか
     * @param {boolean} isVertical - 縦組みか
     * @param {Function} [isSkippedChar] - 1文字ずつのとき、飛ばす文字なら true を返す関数（文字の範囲を渡す）
     * @returns {Object[]} 範囲 { left, top, right, bottom }
     */
    function getShapeAreas(textLines, isPerChar, isVertical, isSkippedChar) {
        var shapeAreas = [];
        for (var i = 0; i < textLines.length; i++) {
            var textLine = textLines[i];
            if (!isPerChar) {
                shapeAreas.push({ left: textLine.left, top: textLine.top, right: textLine.right, bottom: textLine.bottom });
                continue;
            }
            for (var j = 0; j < textLine.chars.length; j++) {
                var charBounds = textLine.chars[j];
                if (isSkippedChar && isSkippedChar(charBounds)) continue;
                if (isVertical) {
                    var halfHeight = Math.max(charBounds.top - charBounds.bottom, textLine.right - textLine.left) / 2;
                    var centerY = (charBounds.top + charBounds.bottom) / 2;
                    shapeAreas.push({ left: textLine.left, top: centerY + halfHeight, right: textLine.right, bottom: centerY - halfHeight });
                } else {
                    var halfWidth = Math.max(charBounds.right - charBounds.left, textLine.top - textLine.bottom) / 2;
                    var centerX = (charBounds.left + charBounds.right) / 2;
                    shapeAreas.push({ left: centerX - halfWidth, top: textLine.top, right: centerX + halfWidth, bottom: textLine.bottom });
                }
            }
        }
        return shapeAreas;
    }

    /**
     * 対象の字面を行（縦組みは列）に分けて返す。［行の高さ］の［行送り基準］なら、行の高さ（縦組みは幅）を行送りにし、
     * 行の文字の中心の中央値を行の中心にする（括弧・g・p・句読点に引っぱられない）。続けて並ぶ行は行送りちょうどの間隔に並べ直す。行の長さの向きは字面のまま
     * @param {Object} markerTarget - collectMarkerTargets() の要素（selectTargetLayout() で選んだあと）
     * @param {boolean} isLeadingBased - 行の高さを行送りにするか
     * @returns {Object[]} groupCharBoundsByLine() の戻り値（［行の高さ］の［行送り基準］なら範囲を書き換えたもの）
     */
    function getTargetTextLines(markerTarget, isLeadingBased) {
        var isVertical = markerTarget.isVertical;
        var textLines = groupCharBoundsByLine(getTargetCharBounds(markerTarget), isVertical);
        if (!isLeadingBased) return textLines;
        var lineCenters = [];
        var lineLeadings = [];
        for (var i = 0; i < textLines.length; i++) {
            var charCenters = [];
            var lineLeading = 0;
            for (var j = 0; j < textLines[i].chars.length; j++) {
                var charBounds = textLines[i].chars[j];
                charCenters.push(isVertical ? (charBounds.left + charBounds.right) / 2 : (charBounds.top + charBounds.bottom) / 2);
                lineLeading = Math.max(lineLeading, getCharLeading(markerTarget, charBounds.charIndex));
            }
            lineCenters.push(getMedian(charCenters));
            lineLeadings.push(lineLeading);
        }
        lineCenters = snapLineCentersToLeading(lineCenters, lineLeadings);
        for (var k = 0; k < textLines.length; k++) {
            var textLine = textLines[k];
            var lineCenter = lineCenters[k];
            lineLeading = lineLeadings[k];
            if (isVertical) {
                textLine.left = lineCenter - lineLeading / 2;
                textLine.right = lineCenter + lineLeading / 2;
            } else {
                textLine.top = lineCenter + lineLeading / 2;
                textLine.bottom = lineCenter - lineLeading / 2;
            }
        }
        return textLines;
    }

    /**
     * 続けて並ぶ行（縦組みは列）の中心を、行送りちょうどの間隔に並べ直す。
     * 文字の中心の中央値は、かなが多い行・欧文の行などで少しずれるので、そのままでは図形の間隔がそろわない。
     * 前の行との間隔が行送りに近い（差が行送りの 1/4 以内で、行送りも同じ）行を1続きとし、
     * 1続きの中では中心からのずれの中央値だけ全体をずらす。段落前後のアキや空行で離れたところは別の続きにする
     * @param {number[]} lineCenters - 行ごとの中心（横組みは Y、縦組みは X 座標。先頭の行から順に、座標は小さくなっていく）
     * @param {number[]} lineLeadings - 行ごとの行送り（pt）
     * @returns {number[]} 並べ直した中心（新しい配列）
     */
    function snapLineCentersToLeading(lineCenters, lineLeadings) {
        var snappedCenters = lineCenters.slice(0);
        var runStart = 0;
        for (var i = 1; i <= lineCenters.length; i++) {
            var lineLeading = lineLeadings[runStart];
            var continuesRun = (i < lineCenters.length && lineLeading > 0
                && Math.abs(lineLeadings[i] - lineLeading) < 0.01
                && Math.abs(lineCenters[i - 1] - lineCenters[i] - lineLeading) <= lineLeading / 4);
            if (continuesRun) continue;
            if (i - runStart > 1) {
                /* 1続きの先頭の行を基準に、行送りの何行分かを引いた位置からのずれ / Offsets from the run start minus whole leadings */
                var centerOffsets = [];
                for (var j = runStart; j < i; j++) centerOffsets.push(lineCenters[j] + (j - runStart) * lineLeading - lineCenters[runStart]);
                var runOffset = getMedian(centerOffsets);
                for (var k = runStart; k < i; k++) snappedCenters[k] = lineCenters[runStart] + runOffset - (k - runStart) * lineLeading;
            }
            runStart = i;
        }
        return snappedCenters;
    }

    /**
     * 文字の行送り（pt）を返す。自動行送りなら 文字サイズ × 自動行送り量。文字ごとに控え、プレビューの描き直しで読み直さない
     * @param {Object} markerTarget - collectMarkerTargets() の要素（charLeadings に控える）
     * @param {number} charIndex - 範囲の先頭からの文字番号
     * @returns {number} 行送り（pt）
     */
    function getCharLeading(markerTarget, charIndex) {
        if (!markerTarget.charLeadings) markerTarget.charLeadings = [];
        if (typeof markerTarget.charLeadings[charIndex] === "number") return markerTarget.charLeadings[charIndex];
        var targetChar = markerTarget.sourceFrame.characters[markerTarget.spanStart + charIndex];
        var charAttributes = targetChar.characterAttributes;
        var charLeading = charAttributes.autoLeading
            ? charAttributes.size * targetChar.paragraphAttributes.autoLeadingAmount / 100
            : charAttributes.leading;
        markerTarget.charLeadings[charIndex] = charLeading;
        return charLeading;
    }

    /**
     * 数の中央値を返す
     * @param {number[]} numberList - 数の配列（1つ以上）
     * @returns {number} 中央値
     */
    function getMedian(numberList) {
        /* 比較関数つきの sort() は遅いので、1行分の少ない数を挿入法で並べる / Comparator sort() is slow; insertion-sort the few values of a line */
        var sortedList = [];
        for (var i = 0; i < numberList.length; i++) {
            var insertAt = sortedList.length;
            while (insertAt > 0 && sortedList[insertAt - 1] > numberList[i]) insertAt--;
            sortedList.splice(insertAt, 0, numberList[i]);
        }
        var middleIndex = Math.floor(sortedList.length / 2);
        return (sortedList.length % 2) ? sortedList[middleIndex] : (sortedList[middleIndex - 1] + sortedList[middleIndex]) / 2;
    }

    /**
     * 対象の1文字ずつの字面の範囲を返す（今の書式の組み合わせで初回に測って控える）
     * @param {Object} markerTarget - collectMarkerTargets() の要素（selectTargetLayout() で選んだあと）
     * @returns {Object[]} measureCharGlyphBounds() の戻り値
     */
    function getTargetCharBounds(markerTarget) {
        var activeLayout = markerTarget.activeLayout;
        if (!activeLayout.charBounds) {
            activeLayout.charBounds = measureCharGlyphBounds(markerTarget.measureFrame, markerTarget.spanStart, markerTarget.spanLength);
        }
        return activeLayout.charBounds;
    }

    /**
     * ［重ね順］の位置（元のテキストの背面か前面）へ、字面を囲む図形（長方形・円・六角形）を描いてグループにする。
     * 図形は行ごとか1文字ずつで（［行の高さ］の［行送り基準］なら行ごとの高さは行送り）、まわりに［余白（横）］［余白（縦）］を足し、［位置調整］の分ずらす。［回転］が 0 でなければ1つずつその中心で回転する。
     * ［塗り］も［線］もオフなら、図形は描かない
     * @param {Object} markerTarget - collectMarkerTargets() の要素（1文字ずつの範囲は初回に測って charBounds に控える）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {GroupItem} 図形のグループ
     */
    function placeTextShapes(markerTarget, markerOptions) {
        var shapeGroup = markerTarget.sourceFrame.layer.groupItems.add();
        var fillColor = getMarkerFillColor(markerTarget, markerOptions);
        var strokeColor = getMarkerStrokeColor(markerTarget, markerOptions);
        var paddingH = markerOptions.boxPadding;
        var paddingV = markerOptions.boxPaddingV;
        var textLines = getTargetTextLines(markerTarget, markerOptions.isLeadingBased);
        var shapeAreas = getShapeAreas(textLines, markerOptions.isPerChar, markerTarget.isVertical, function (charBounds) {
            return isPerCharSkipped(markerTarget, charBounds);
        });
        if ((markerOptions.equalLineLengths || markerOptions.alignLineStarts) && !markerOptions.isPerChar) {
            shapeAreas = alignLineEdges(shapeAreas, markerTarget.isVertical, markerOptions.equalLineLengths);
        }
        /* 行ごとの長方形を1続きに見せるときは、行の継ぎ目側の角を丸めない / Keep the corners square where the lines continue */
        var joinsLines = (markerOptions.lineStyle === "box" && !markerOptions.isPerChar && markerOptions.squareLineJoins && shapeAreas.length > 1);
        for (var i = 0; i < shapeAreas.length; i++) {
            /* ［位置調整］の分だけずらす（プラスで上。縦組みは左） / Shift by the offset (positive is up; left in vertical text) */
            var shapeBounds = {
                left: shapeAreas[i].left - paddingH - (markerTarget.isVertical ? markerOptions.verticalOffset : 0),
                top: shapeAreas[i].top + paddingV + (markerTarget.isVertical ? 0 : markerOptions.verticalOffset),
                width: shapeAreas[i].right - shapeAreas[i].left + paddingH * 2,
                height: shapeAreas[i].top - shapeAreas[i].bottom + paddingV * 2,
                roundedCorners: joinsLines ? getLineJoinRoundedCorners(i, shapeAreas.length, markerTarget.isVertical) : [true, true, true, true]
            };
            /* 負の余白で幅か高さが無くなったもの、塗りも線も無いものは描かない / Skip collapsed shapes and shapes with neither fill nor stroke */
            if (shapeBounds.width <= 0 || shapeBounds.height <= 0 || (!markerOptions.isShapeFill && !markerOptions.isShapeStroke)) continue;
            var textShape = joinsLines && markerOptions.isShapeStroke
                ? addJoinedBoxShape(shapeGroup, shapeBounds, getLineJoinOpenEdges(i, shapeAreas.length, markerTarget.isVertical), fillColor, strokeColor, markerOptions)
                : addShapePath(shapeGroup, shapeBounds, markerTarget, markerOptions);
            /* 図形ごとに、その中心で回転する / Rotate each shape about its own center */
            if (markerOptions.boxRotation) textShape.rotate(markerOptions.boxRotation, true, false, false, false, Transformation.CENTER);
            if (textShape.typename === "PathItem") paintShapePath(textShape, fillColor, strokeColor, markerOptions);
            applyRoughenEffect(textShape, markerOptions);
        }
        shapeGroup.opacity = markerOptions.opacity;
        moveToStackPosition(shapeGroup, markerTarget, markerOptions);
        return shapeGroup;
    }

    /**
     * 範囲に、種類の形のパスを作る。
     * ［長方形］は［角丸の半径］が 0 より大きければ、roundedCorners の角を丸める（半径は短い辺の半分まで）。
     * ［円］は楕円（［1文字ずつ］なら長い辺を直径にした円）、［六角形］は getHexagonPoints() の形で、どちらも角丸は使わない
     * @param {GroupItem} shapeGroup - 追加先のグループ
     * @param {{left: number, top: number, width: number, height: number, roundedCorners: boolean[]}} shapeBounds - 図形の範囲（余白・位置調整を含む）と、
     *     角を丸めるかどうか [左上, 右上, 右下, 左下]
     * @param {Object} markerTarget - collectMarkerTargets() の要素（isVertical）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値（lineStyle・isPerChar・cornerRadius）
     * @returns {PathItem} 作ったパス
     */
    function addShapePath(shapeGroup, shapeBounds, markerTarget, markerOptions) {
        var shapeLeft = shapeBounds.left;
        var shapeTop = shapeBounds.top;
        var shapeWidth = shapeBounds.width;
        var shapeHeight = shapeBounds.height;
        if (markerOptions.lineStyle === "circle") {
            if (!markerOptions.isPerChar) return shapeGroup.pathItems.ellipse(shapeTop, shapeLeft, shapeWidth, shapeHeight);
            /* 直径は長い辺。範囲の中心にそろえる / The diameter is the longer side, centered on the area */
            var circleSize = Math.max(shapeWidth, shapeHeight);
            return shapeGroup.pathItems.ellipse(shapeTop - (shapeHeight - circleSize) / 2, shapeLeft - (circleSize - shapeWidth) / 2, circleSize, circleSize);
        }
        if (markerOptions.lineStyle === "hexagon") {
            var hexagonPath = shapeGroup.pathItems.add();
            hexagonPath.setEntirePath(getHexagonPoints(shapeLeft, shapeTop, shapeWidth, shapeHeight, markerOptions.isPerChar, markerTarget.isVertical));
            hexagonPath.closed = true;
            return hexagonPath;
        }
        var cornerRadius = Math.min(markerOptions.cornerRadius, shapeWidth / 2, shapeHeight / 2);
        if (!(cornerRadius > 0)) return shapeGroup.pathItems.rectangle(shapeTop, shapeLeft, shapeWidth, shapeHeight);
        var roundedCorners = shapeBounds.roundedCorners;
        if (roundedCorners[0] && roundedCorners[1] && roundedCorners[2] && roundedCorners[3]) {
            return shapeGroup.pathItems.roundedRectangle(shapeTop, shapeLeft, shapeWidth, shapeHeight, cornerRadius, cornerRadius);
        }
        return addPartlyRoundedRectangle(shapeGroup, shapeBounds, cornerRadius);
    }

    /**
     * 複数行の［長方形］の1行分を、行の継ぎ目の辺を開けて描く（［連続処理］の［複数行］で［線］の［描画］がオンのとき）。
     * 塗りは閉じた長方形、線は開けた辺を除いたオープンパスにして、1行分をグループにまとめる
     * @param {GroupItem} shapeGroup - 追加先のグループ
     * @param {{left: number, top: number, width: number, height: number, roundedCorners: boolean[]}} shapeBounds - 図形の範囲と、
     *     角を丸めるかどうか [左上, 右上, 右下, 左下]
     * @param {boolean[]} openEdges - 線を開ける辺 [上, 右, 下, 左]
     * @param {Color} fillColor - 塗りの色
     * @param {Color} strokeColor - 線の色
     * @param {Object} markerOptions - readMarkerOptions() の戻り値（isShapeFill・cornerRadius・線の指定）
     * @returns {GroupItem} 1行分のグループ
     */
    function addJoinedBoxShape(shapeGroup, shapeBounds, openEdges, fillColor, strokeColor, markerOptions) {
        var lineShapeGroup = shapeGroup.groupItems.add();
        var cornerRadius = Math.min(markerOptions.cornerRadius, shapeBounds.width / 2, shapeBounds.height / 2);
        if (markerOptions.isShapeFill) {
            var fillPath = addPartlyRoundedRectangle(lineShapeGroup, shapeBounds, cornerRadius);
            paintShapePath(fillPath, fillColor, strokeColor, markerOptions);
            fillPath.stroked = false;
        }
        var edgePaths = addOpenRectangleEdges(lineShapeGroup, shapeBounds, cornerRadius, openEdges);
        for (var i = 0; i < edgePaths.length; i++) {
            paintShapePath(edgePaths[i], fillColor, strokeColor, markerOptions);
            edgePaths[i].filled = false;
        }
        return lineShapeGroup;
    }

    /**
     * 複数行の長方形で、角を丸めるかどうかを返す。
     * 横組みは先頭の行の左側と最後の行の右側、縦組みは先頭の列の上側と最後の列の下側だけを丸め、あいだの継ぎ目は角にする
     * @param {number} lineIndex - 行（列）の番号（0始まり）
     * @param {number} lineCount - 行（列）の数（2以上）
     * @param {boolean} isVertical - 縦組みか
     * @returns {boolean[]} 丸めるなら true [左上, 右上, 右下, 左下]
     */
    function getLineJoinRoundedCorners(lineIndex, lineCount, isVertical) {
        var isFirst = (lineIndex === 0);
        var isLast = (lineIndex === lineCount - 1);
        if (isVertical) return [isFirst, isFirst, isLast, isLast];
        return [isFirst, isLast, isLast, isFirst];
    }

    /**
     * 複数行の長方形で、線を開ける（行の継ぎ目の）辺を返す。
     * 横組みは先頭の行以外の左辺と最後の行以外の右辺、縦組みは先頭の列以外の上辺と最後の列以外の下辺
     * @param {number} lineIndex - 行（列）の番号（0始まり）
     * @param {number} lineCount - 行（列）の数（2以上）
     * @param {boolean} isVertical - 縦組みか
     * @returns {boolean[]} 開けるなら true [上, 右, 下, 左]
     */
    function getLineJoinOpenEdges(lineIndex, lineCount, isVertical) {
        var isFirst = (lineIndex === 0);
        var isLast = (lineIndex === lineCount - 1);
        if (isVertical) return [!isFirst, false, !isLast, false];
        return [false, !isLast, false, !isFirst];
    }

    /**
     * 一部の角だけを丸めた長方形のパスを作る（円弧はベジェ曲線で近似）
     * @param {GroupItem} shapeGroup - 追加先のグループ
     * @param {{left: number, top: number, width: number, height: number, roundedCorners: boolean[]}} shapeBounds - 範囲と、
     *     角を丸めるかどうか [左上, 右上, 右下, 左下]
     * @param {number} cornerRadius - 角丸の半径（pt）
     * @returns {PathItem} 作ったパス
     */
    function addPartlyRoundedRectangle(shapeGroup, shapeBounds, cornerRadius) {
        var cornerAnchors = getRectangleCornerAnchors(shapeBounds, cornerRadius);
        var shapePath = shapeGroup.pathItems.add();
        for (var i = 0; i < cornerAnchors.length; i++) {
            for (var j = 0; j < cornerAnchors[i].length; j++) {
                addPathAnchor(shapePath, cornerAnchors[i][j][0], cornerAnchors[i][j][1], cornerAnchors[i][j][2]);
            }
        }
        shapePath.closed = true;
        return shapePath;
    }

    /**
     * 長方形の線を、開ける辺を除いたオープンパスで作る（行の継ぎ目の辺を開けて、複数行の囲み罫を1続きに見せる）。
     * 開けた辺で分かれるので、パスは1本か2本になる
     * @param {GroupItem} shapeGroup - 追加先のグループ
     * @param {{left: number, top: number, width: number, height: number, roundedCorners: boolean[]}} shapeBounds - 範囲と、
     *     角を丸めるかどうか [左上, 右上, 右下, 左下]
     * @param {number} cornerRadius - 角丸の半径（pt。0 なら角丸なし）
     * @param {boolean[]} openEdges - 開ける辺 [上, 右, 下, 左]（1つ以上 true）
     * @returns {PathItem[]} 作ったパス
     */
    function addOpenRectangleEdges(shapeGroup, shapeBounds, cornerRadius, openEdges) {
        var cornerAnchors = getRectangleCornerAnchors(shapeBounds, cornerRadius);
        /* 辺 k は角 k から角 k+1 へ（上・右・下・左）。開けた辺の次の角から1周たどり、開けた辺で切る
           Edge k runs from corner k to corner k+1 (top, right, bottom, left); walk around from the corner after an open edge */
        var firstOpenEdge = 0;
        while (!openEdges[firstOpenEdge]) firstOpenEdge++;
        var startCorner = (firstOpenEdge + 1) % 4;
        var edgePaths = [];
        var pieceCorners = [startCorner];
        for (var step = 0; step < 4; step++) {
            var edgeIndex = (startCorner + step) % 4;
            var nextCorner = (edgeIndex + 1) % 4;
            if (!openEdges[edgeIndex]) {
                pieceCorners.push(nextCorner);
                continue;
            }
            /* 辺を1本も含まない（角だけの）ものは描かない / Skip pieces with no edge */
            if (pieceCorners.length > 1) edgePaths.push(addCornerChainPath(shapeGroup, cornerAnchors, pieceCorners));
            pieceCorners = [nextCorner];
        }
        return edgePaths;
    }

    /**
     * 角を順につないだオープンパスを作る
     * @param {GroupItem} shapeGroup - 追加先のグループ
     * @param {Array[]} cornerAnchors - getRectangleCornerAnchors() の戻り値
     * @param {number[]} cornerIndexes - つなぐ角の番号（0：左上、1：右上、2：右下、3：左下）
     * @returns {PathItem} 作ったパス
     */
    function addCornerChainPath(shapeGroup, cornerAnchors, cornerIndexes) {
        var chainPath = shapeGroup.pathItems.add();
        for (var i = 0; i < cornerIndexes.length; i++) {
            var anchors = cornerAnchors[cornerIndexes[i]];
            for (var j = 0; j < anchors.length; j++) addPathAnchor(chainPath, anchors[j][0], anchors[j][1], anchors[j][2]);
        }
        chainPath.closed = false;
        return chainPath;
    }

    /**
     * 長方形の4つの角のアンカーを、時計回り（左上・右上・右下・左下）に返す。
     * 丸める角は円弧の両端の2つ（円弧はベジェ曲線で近似）、丸めない角は角の1つ
     * @param {{left: number, top: number, width: number, height: number, roundedCorners: boolean[]}} shapeBounds - 範囲と、
     *     角を丸めるかどうか [左上, 右上, 右下, 左下]
     * @param {number} cornerRadius - 角丸の半径（pt。0 なら角丸なし）
     * @returns {Array[]} 角ごとのアンカー [[アンカー, 前のハンドル, 次のハンドル], ...] の配列
     */
    function getRectangleCornerAnchors(shapeBounds, cornerRadius) {
        var leftX = shapeBounds.left;
        var topY = shapeBounds.top;
        var rightX = leftX + shapeBounds.width;
        var bottomY = topY - shapeBounds.height;
        /* 4つの角：角の位置と、角から隣り合う2辺へ向かう向き（時計回りで、入ってくる辺・出ていく辺の順）
           The corners, with directions along the incoming and outgoing edges (clockwise) */
        var corners = [
            { point: [leftX, topY], inward: [0, -1], outward: [1, 0] },
            { point: [rightX, topY], inward: [-1, 0], outward: [0, -1] },
            { point: [rightX, bottomY], inward: [0, 1], outward: [-1, 0] },
            { point: [leftX, bottomY], inward: [1, 0], outward: [0, 1] }
        ];
        /* 円弧を3次ベジェで近似するときのハンドルの長さの比率 / Handle length ratio for a cubic Bezier quarter circle */
        var handleLength = cornerRadius * 0.5523;
        var cornerAnchors = [];
        for (var i = 0; i < corners.length; i++) {
            var corner = corners[i];
            if (!(cornerRadius > 0) || !shapeBounds.roundedCorners[i]) {
                cornerAnchors.push([[corner.point, corner.point, corner.point]]);
                continue;
            }
            /* 入ってくる辺の上の円弧の始まりと、出ていく辺の上の円弧の終わり / Arc start on the incoming edge and arc end on the outgoing edge */
            var arcStart = [corner.point[0] + corner.inward[0] * cornerRadius, corner.point[1] + corner.inward[1] * cornerRadius];
            var arcEnd = [corner.point[0] + corner.outward[0] * cornerRadius, corner.point[1] + corner.outward[1] * cornerRadius];
            cornerAnchors.push([
                [arcStart, arcStart, [arcStart[0] - corner.inward[0] * handleLength, arcStart[1] - corner.inward[1] * handleLength]],
                [arcEnd, [arcEnd[0] - corner.outward[0] * handleLength, arcEnd[1] - corner.outward[1] * handleLength], arcEnd]
            ]);
        }
        return cornerAnchors;
    }

    /**
     * パスの最後にアンカーポイントを足す
     * @param {PathItem} targetPath - パス
     * @param {number[]} anchorPoint - アンカー [x, y]
     * @param {number[]} leftDirection - 前のアンカー側のハンドル [x, y]
     * @param {number[]} rightDirection - 次のアンカー側のハンドル [x, y]
     * @returns {void}
     */
    function addPathAnchor(targetPath, anchorPoint, leftDirection, rightDirection) {
        var pathPoint = targetPath.pathPoints.add();
        pathPoint.anchor = anchorPoint;
        pathPoint.leftDirection = leftDirection;
        pathPoint.rightDirection = rightDirection;
    }

    /**
     * 図形のパスに［塗り］［線］を付ける。線は［線幅］［線の種類］を使い、閉じたパスなのでドット以外は線端なしにする
     * @param {PathItem} textShape - 図形のパス
     * @param {Color} fillColor - 塗りの色
     * @param {Color} strokeColor - 線の色
     * @param {Object} markerOptions - readMarkerOptions() の戻り値（isShapeFill・isShapeStroke・strokeWidth・dashStyle）
     * @returns {void}
     */
    function paintShapePath(textShape, fillColor, strokeColor, markerOptions) {
        textShape.filled = markerOptions.isShapeFill;
        if (markerOptions.isShapeFill) textShape.fillColor = fillColor;
        textShape.stroked = markerOptions.isShapeStroke;
        if (!markerOptions.isShapeStroke) return;
        textShape.strokeColor = strokeColor;
        textShape.strokeWidth = markerOptions.strokeWidth;
        applyLineDashes(textShape, markerOptions);
        /* 閉じたパスなので、ドット以外は線端を問わない / A closed path: the cap matters only for Dotted */
        if (markerOptions.dashStyle !== "dottedLine") textShape.strokeCap = StrokeCap.BUTTENDCAP;
    }

    /**
     * 線の種類（直線・破線・ドット）と線端を設定する。破線は［分割数］［間隔］、ドットは［間隔］で並べる。
     * ドットは長さ0の線分を丸型の線端で点にする
     * @param {PathItem} markerLine - 線（線幅は設定済み）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値（dashStyle・dashSegments・dashGap・strokeCap）
     * @returns {void}
     */
    function applyLineDashes(markerLine, markerOptions) {
        var dashStyle = markerOptions.dashStyle;
        var lineWidth = markerLine.strokeWidth;
        markerLine.strokeCap = (dashStyle === "dottedLine") ? StrokeCap.ROUNDENDCAP : markerOptions.strokeCap;
        if (dashStyle === "dashedLine") markerLine.strokeDashes = calcFittedDashes(markerLine.length, markerLine.closed, lineWidth, markerOptions.dashGap, markerOptions.dashSegments);
        else if (dashStyle === "dottedLine") markerLine.strokeDashes = calcFittedDots(markerLine.length, markerLine.closed, lineWidth, markerOptions.dashGap);
        else markerLine.strokeDashes = [];
    }

    /**
     * 破線の線分と間隔を、パスの長さに合わせて求める（SmartStrokeSettings の［間隔 → 線分］［両端を調整］と同じ計算）。
     * 間隔は gapPt のまま、線分の本数は dashCount（null なら線分が線幅の DASHED_LINE_RATIOS[0] 倍に近くなる本数）で長さを割り振る。
     * オープンパスは両端が線分で終わるように線分 n 本・間隔 n−1 個、クローズパスは1周を n 等分する
     * @param {number} pathLength - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスか
     * @param {number} lineWidth - 線幅（pt）
     * @param {number} gapPt - 線分と線分のあいだ（pt）
     * @param {number|null} dashCount - 線分の本数（［分割数］。null なら自動）
     * @returns {number[]} strokeDashes に入れる [線分, 間隔]（pt）。オープンパスで線分が1本なら [] （実線）。
     *     間隔が長すぎて線分が入らないときは、線幅の DASHED_LINE_RATIOS[0] 倍の線分
     */
    function calcFittedDashes(pathLength, isClosed, lineWidth, gapPt, dashCount) {
        var cyclePt = lineWidth * DASHED_LINE_RATIOS[0] + gapPt;
        var fallbackDashes = [lineWidth * DASHED_LINE_RATIOS[0], gapPt];
        if (!(cyclePt > 0) || !(pathLength > 0)) return fallbackDashes;
        /* 両端を線分にするなら、全長＋間隔1つ分が周期の n 倍 / With dashes at both ends, length + one gap is n cycles */
        if (!(dashCount >= 1)) dashCount = Math.max(1, Math.round((isClosed ? pathLength : pathLength + gapPt) / cyclePt));
        if (!isClosed && dashCount === 1) return [];
        var dashPt = isClosed ? pathLength / dashCount - gapPt : (pathLength - gapPt * (dashCount - 1)) / dashCount;
        return (dashPt > 0) ? [dashPt, gapPt] : fallbackDashes;
    }

    /**
     * ドットの点の間隔を、パスの長さに合わせて求める（SmartStrokeSettings のドット点線［両端を調整］と同じ計算）。
     * 点のあいだが gapPt に近くなる点の数で長さを割り振る。
     * オープンパスは両端に点を置き（点は2つ以上）、クローズパスは1周を等分する
     * @param {number} pathLength - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスか
     * @param {number} lineWidth - 線幅（点の直径。pt）
     * @param {number} gapPt - 点と点のあいだ（pt。点の中心どうしは線幅＋gapPt）
     * @returns {number[]} strokeDashes に入れる [0, 点の中心どうしの間隔]（pt）
     */
    function calcFittedDots(pathLength, isClosed, lineWidth, gapPt) {
        var pitchPt = lineWidth + gapPt;
        if (!(pitchPt > 0) || !(pathLength > 0)) return [0, pitchPt];
        if (isClosed) return [0, pathLength / Math.max(1, Math.round(pathLength / pitchPt))];
        /* 両端に点：点 n 個と間隔 n−1 個 / Dots at both ends: n dots and n - 1 gaps */
        var dotCount = Math.max(2, Math.round((pathLength + pitchPt) / pitchPt));
        /* 最後の点が終点を越えないよう、間隔を切り捨てる / Truncate so the last dot stays inside the end */
        return [0, Math.floor(pathLength / (dotCount - 1) * DOT_GAP_PRECISION) / DOT_GAP_PRECISION];
    }

    /**
     * ［塗り］のカラーの色を返す（［文字色］なら対象の先頭の文字の塗り）
     * @param {Object} markerTarget - collectMarkerTargets() の要素
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {Color} 色
     */
    function getMarkerFillColor(markerTarget, markerOptions) {
        return markerOptions.fillUseTextColor ? markerTarget.textColor : markerOptions.fillColor;
    }

    /**
     * ［線］のカラーの色を返す（［文字色］なら対象の先頭の文字の塗り）
     * @param {Object} markerTarget - collectMarkerTargets() の要素
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {Color} 色
     */
    function getMarkerStrokeColor(markerTarget, markerOptions) {
        return markerOptions.strokeUseTextColor ? markerTarget.textColor : markerOptions.strokeColor;
    }

    /**
     * すべての対象にマーカーの線（長方形の仲間なら図形のグループ、［傍点］なら点のグループ）を描く
     * @param {Object[]} markerTargets - collectMarkerTargets() の戻り値
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @param {TextFrame[]|null} measureFrames - 字面を測るテキスト（markerTargets と同じ順。null なら元のテキスト）
     * @returns {PageItem[]} 描いた線と長方形のグループ
     */
    function placeMarkers(markerTargets, markerOptions, measureFrames) {
        var markerItems = [];
        var layoutKey = getTextLayoutKey(markerOptions);
        for (var i = 0; i < markerTargets.length; i++) {
            selectTargetLayout(markerTargets[i], layoutKey, measureFrames ? measureFrames[i] : markerTargets[i].sourceFrame);
            if (markerOptions.lineStyle === "emphasisDots") markerItems.push(placeEmphasisDots(markerTargets[i], markerOptions));
            else if (isShapeStyle(markerOptions.lineStyle)) markerItems.push(placeTextShapes(markerTargets[i], markerOptions));
            else markerItems.push(placeMarkerLine(markerTargets[i], markerOptions));
        }
        return markerItems;
    }

    /**
     * 描いた線・図形を［重ね順］の位置へ移す（元のテキストがグループの中なら、そのグループに入る）。
     * ［後］は元のテキストの直後（背面）、［前］は字面を測ったテキストの直前（前面）。
     * プレビューで書式を変えているときは、元の前面に置いた複製のさらに前になる
     * @param {PageItem} markerItem - 描いた線・図形
     * @param {Object} markerTarget - collectMarkerTargets() の要素（selectTargetLayout() で measureFrame を選んだあと）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {void}
     */
    function moveToStackPosition(markerItem, markerTarget, markerOptions) {
        if (markerOptions.isFront) markerItem.move(markerTarget.measureFrame, ElementPlacement.PLACEBEFORE);
        else markerItem.move(markerTarget.sourceFrame, ElementPlacement.PLACEAFTER);
    }

    /**
     * 1文字ずつの図形・傍点で飛ばす文字（PER_CHAR_SKIP_CHARS）かを返す
     * @param {Object} markerTarget - collectMarkerTargets() の要素
     * @param {Object} charBounds - measureCharGlyphBounds() の要素（charIndex 付き）
     * @returns {boolean} 飛ばす文字なら true
     */
    function isPerCharSkipped(markerTarget, charBounds) {
        var charText = markerTarget.sourceFrame.characters[markerTarget.spanStart + charBounds.charIndex].contents;
        return PER_CHAR_SKIP_CHARS.indexOf(charText) >= 0;
    }

    /**
     * 字の配置を変える書式（自動カーニング・トラッキング・カーニング（前・後））の組み合わせを、字面の測り直しを見分けるキーにする
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {string} キー（書式を変えないなら空文字）
     */
    function getTextLayoutKey(markerOptions) {
        /* ［余白に合わせてカーニングを自動調整］のカーニングは対象ごとに違うが、控えも対象ごとなので余白で見分ければよい
           Auto-Adjust Kerning to Padding kerning differs per target, but so does the cache, so the padding tells them apart */
        var kerningKey = markerOptions.kerningFromPadding
            ? ["padding", markerOptions.boxPadding, markerOptions.boxPaddingV, markerOptions.kerningAdjustBefore, markerOptions.kerningAdjustAfter]
            : [markerOptions.kerningBefore, markerOptions.kerningAfter];
        if (markerOptions.kerningChoice === "keepKerning" && markerOptions.tracking === null
            && !markerOptions.kerningFromPadding && markerOptions.kerningBefore === null && markerOptions.kerningAfter === null) return "";
        return [markerOptions.kerningChoice, markerOptions.tracking].concat(kerningKey).join("|");
    }

    /**
     * 文字の書式（自動カーニング・トラッキング・カーニング（前・後）・テキストカラー）を変えるかを返す
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {boolean} 変えるなら true
     */
    function hasTextFormatChange(markerOptions) {
        return markerOptions.textFillColor !== null || getTextLayoutKey(markerOptions) !== "";
    }

    /**
     * 対象の文字に書式（自動カーニング・トラッキング・カーニング（前・後）・テキストカラー）を設定する（変えない項目はそのまま）。
     * ［余白に合わせてカーニングを自動調整］なら、カーニング（前・後）は余白と対象の先頭・最後の文字のサイズから計算し、微調整を足す
     * @param {TextFrame} textFrame - 元のテキストフレーム、またはその複製
     * @param {Object} markerTarget - collectMarkerTargets() の要素（spanStart・spanLength）
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {void}
     */
    function applyTextFormat(textFrame, markerTarget, markerOptions) {
        var targetSpan = textFrame.characters[markerTarget.spanStart];
        targetSpan.length = markerTarget.spanLength;
        var spanAttributes = targetSpan.characterAttributes;
        var kerningChoice = markerOptions.kerningChoice;
        if (kerningChoice === "metricsKerning") spanAttributes.kerningMethod = AutoKernType.AUTO;
        else if (kerningChoice === "opticalKerning") spanAttributes.kerningMethod = AutoKernType.OPTICAL;
        else if (kerningChoice === "zeroKerning") {
            /* 自動を切っても手動の値は残るので、文字ごとに 0 にする / Manual values survive turning auto off, so zero each character */
            spanAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
            for (var i = 0; i < markerTarget.spanLength; i++) textFrame.characters[markerTarget.spanStart + i].kerning = 0;
        }
        /* ［0］［オプティカル］は字形本来の送りで詰めるので、プロポーショナルメトリクスを外す / Zero and Optical turn Proportional Metrics off */
        if (kerningChoice === "zeroKerning" || kerningChoice === "opticalKerning") spanAttributes.proportionalMetrics = false;
        if (markerOptions.tracking !== null) spanAttributes.tracking = markerOptions.tracking;
        var kerningBefore = markerOptions.kerningBefore;
        var kerningAfter = markerOptions.kerningAfter;
        /* ［余白に合わせてカーニングを自動調整］は対象ごとの文字サイズで計算する（縦組みは列の向きの［余白（縦）］） / Auto-Adjust Kerning to Padding uses each target's type size */
        if (markerOptions.kerningFromPadding) {
            var paddingPt = markerTarget.isVertical ? markerOptions.boxPaddingV : markerOptions.boxPadding;
            var edgeSizes = getTargetEdgeFontSizes(textFrame, markerTarget);
            kerningBefore = calcPaddingKerning(paddingPt, edgeSizes.firstSize, markerOptions.kerningAdjustBefore);
            kerningAfter = calcPaddingKerning(paddingPt, edgeSizes.lastSize, markerOptions.kerningAdjustAfter);
        }
        /* 手動カーニングはその文字の前に入るので、前の間隔は先頭の文字、後の間隔は範囲の次の文字に入れる
           Manual kerning goes before its character: the first character for the space before, the next one for the space after */
        if (kerningBefore !== null) setManualKerning(textFrame, markerTarget.spanStart, kerningBefore);
        if (kerningAfter !== null) setManualKerning(textFrame, markerTarget.spanStart + markerTarget.spanLength, kerningAfter);
        if (markerOptions.textFillColor !== null) spanAttributes.fillColor = markerOptions.textFillColor;
    }

    /**
     * 文字の前に手動カーニングを入れる（その位置の自動カーニングは切る）。文字が無ければ何もしない
     * @param {TextFrame} textFrame - テキストフレーム
     * @param {number} charIndex - 文字番号（フレーム内、0始まり）
     * @param {number} kerningValue - カーニング（1/1000 em）
     * @returns {void}
     */
    function setManualKerning(textFrame, charIndex, kerningValue) {
        if (charIndex >= textFrame.characters.length) return;
        var targetChar = textFrame.characters[charIndex];
        targetChar.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
        targetChar.kerning = kerningValue;
    }

    /**
     * 対象の字面の測り方を、書式の組み合わせ（layoutKey）に切り替える。初めての組み合わせなら measureFrame で測って控える
     * @param {Object} markerTarget - collectMarkerTargets() の要素（layouts・activeLayout・measureFrame を持たせる）
     * @param {string} layoutKey - getTextLayoutKey() の戻り値
     * @param {TextFrame} measureFrame - その書式になっているテキスト（元か複製）
     * @returns {void}
     */
    function selectTargetLayout(markerTarget, layoutKey, measureFrame) {
        if (!markerTarget.layouts.hasOwnProperty(layoutKey)) {
            var glyphBounds = measureGlyphBounds(measureFrame, markerTarget.spanStart, markerTarget.spanLength);
            markerTarget.layouts[layoutKey] = { glyphBounds: glyphBounds || markerTarget.glyphBounds, charBounds: null };
        }
        markerTarget.activeLayout = markerTarget.layouts[layoutKey];
        markerTarget.measureFrame = measureFrame;
    }

    /**
     * ［重ね順］の位置（元のテキストの背面か前面）へ、1文字ずつ丸い点（傍点）を置いてグループにする。
     * 横組みは行の字面（［行の高さ］の［行送り基準］なら行送り）の上端から［余白（縦）］離した上、縦組みは列の右端から［余白（横）］離した右に、字面の中央にそろえて置く。
     * PER_CHAR_SKIP_CHARS の文字には置かない
     * @param {Object} markerTarget - collectMarkerTargets() の要素
     * @param {Object} markerOptions - readMarkerOptions() の戻り値
     * @returns {GroupItem} 点のグループ
     */
    function placeEmphasisDots(markerTarget, markerOptions) {
        var sourceFrame = markerTarget.sourceFrame;
        var dotColor = getMarkerFillColor(markerTarget, markerOptions);
        var dotSize = markerOptions.dotSize;
        /* 点は字の上（縦組みは右）なので、横組みは［余白（縦）］、縦組みは［余白（横）］を使う / Dots sit above (right in vertical text): V padding for horizontal, H for vertical */
        var dotDistance = (markerTarget.isVertical ? markerOptions.boxPadding : markerOptions.boxPaddingV) + dotSize / 2;
        var dotGroup = sourceFrame.layer.groupItems.add();
        var textLines = getTargetTextLines(markerTarget, markerOptions.isLeadingBased);
        for (var i = 0; i < textLines.length; i++) {
            for (var j = 0; j < textLines[i].chars.length; j++) {
                var charBounds = textLines[i].chars[j];
                if (isPerCharSkipped(markerTarget, charBounds)) continue;
                var centerX = markerTarget.isVertical ? textLines[i].right + dotDistance : (charBounds.left + charBounds.right) / 2;
                var centerY = markerTarget.isVertical ? (charBounds.top + charBounds.bottom) / 2 : textLines[i].top + dotDistance;
                /* ［位置調整］の分だけずらす（プラスで上。縦組みは左） / Shift by the offset (positive is up; left in vertical text) */
                if (markerTarget.isVertical) centerX -= markerOptions.verticalOffset;
                else centerY += markerOptions.verticalOffset;
                var emphasisDot = dotGroup.pathItems.ellipse(centerY + dotSize / 2, centerX - dotSize / 2, dotSize, dotSize);
                emphasisDot.filled = true;
                emphasisDot.fillColor = dotColor;
                emphasisDot.stroked = false;
            }
        }
        dotGroup.opacity = markerOptions.opacity;
        moveToStackPosition(dotGroup, markerTarget, markerOptions);
        return dotGroup;
    }

    /**
     * 文字の範囲の塗りを変える（1つの TextRange として塗る）
     * @param {TextFrame} textFrame - テキストフレーム
     * @param {number} spanStart - 範囲の先頭の文字番号（フレーム内、0始まり）
     * @param {number} spanLength - 範囲の文字数
     * @param {Color} fillColor - 塗りの色
     * @returns {void}
     */
    function fillCharacterSpan(textFrame, spanStart, spanLength, fillColor) {
        var characterSpan = textFrame.characters[spanStart];
        characterSpan.length = spanLength;
        characterSpan.characterAttributes.fillColor = fillColor;
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
     * @returns {Object[]|null} 対象 { sourceFrame, spanStart, spanLength, glyphBounds, textColor, isVertical } の配列（字形の無いものは除く）。
     *     テキストが選ばれていなければ null
     */
    function collectMarkerTargets(docSelection) {
        var markerTargets = [];

        /**
         * 範囲の字面を測って対象に足す（字形が無ければ足さない）
         * @param {TextFrame} sourceFrame - 元のテキストフレーム
         * @param {number} spanStart - 範囲の先頭の文字番号（フレーム内、0始まり）
         * @param {number} spanLength - 範囲の文字数
         * @param {TextRange} colorRange - ［文字色］を読む範囲
         * @returns {void}
         */
        function addMarkerTarget(sourceFrame, spanStart, spanLength, colorRange) {
            var glyphBounds = measureGlyphBounds(sourceFrame, spanStart, spanLength);
            if (!glyphBounds) return;
            markerTargets.push({
                sourceFrame: sourceFrame, spanStart: spanStart, spanLength: spanLength,
                glyphBounds: glyphBounds, textColor: getTextFillColor(colorRange),
                isVertical: (sourceFrame.orientation === TextOrientation.VERTICAL),
                /* 書式の組み合わせごとの字面（"" は今の書式） / Glyph bounds per formatting combination ("" is the current formatting) */
                layouts: { "": { glyphBounds: glyphBounds, charBounds: null } }
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
     * 文字ツールで文字を選んでいるときは、選択ツールに持ち替えて文字の編集を抜ける。
     * 文字の編集中は、メニューコマンドの［パスのアウトライン］効果が線に付かないため
     * @param {TextRange|Array} docSelection - 実行時の doc.selection
     * @returns {string|null} 持ち替える前のツール名。持ち替えなかったら null
     */
    function leaveTextEditing(docSelection) {
        if (!docSelection || docSelection.typename !== "TextRange") return null;
        var previousToolName = "Adobe Type Tool";
        /* getSelectedToolName() は非公開のメソッドなので、無ければ文字ツールに戻す / undocumented; fall back to the Type tool */
        try {
            previousToolName = String(app.getSelectedToolName());
        } catch (e) {}
        app.selectTool("Adobe Select Tool");
        return previousToolName;
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
     * 選んだ文字・テキストオブジェクトにマーカーを引く
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

        var previousToolName = leaveTextEditing(savedSelection);
        /* 実行中は［境界線を隠す］を切り替え、終わったら（途中で止まっても）元に戻す。edge は表示・非表示の切り替え
           Toggle Hide Edges while running and toggle it back at the end, even on an error; "edge" is a toggle */
        app.executeMenuCommand("edge");
        try {
            var markerOptions = showMarkerDialog(doc, markerTargets);
            if (markerOptions) {
                /* 書式を先に変えて、変えたあとの字面で描く / Format the text first, then draw on the reformatted glyphs */
                if (hasTextFormatChange(markerOptions)) {
                    for (var i = 0; i < markerTargets.length; i++) applyTextFormat(markerTargets[i].sourceFrame, markerTargets[i], markerOptions);
                }
                placeMarkers(markerTargets, markerOptions, null);
            }
        } finally {
            app.executeMenuCommand("edge");
        }
        /* 文字を選び直す前に、文字ツールへ戻す / Return to the Type tool before reselecting the characters */
        if (previousToolName) app.selectTool(previousToolName);
        restoreSelection(doc, savedSelection);
    }

    main();
})();
