#target illustrator
#targetengine "SpeechBubbleMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SpeechBubbleMaker
Name=テキストに吹き出しのしっぽ付きの下線・囲みを描く
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.2.1
Release-Date=2026-11-02
Description-BEGIN
選択したテキストに、吹き出しのしっぽ（V字・斜線・すぼまり・U字・ループ・つの・思考）の付いた下線・L字・コの字・囲み・上下の線を描きます。線は上下左右のどの側にも描けます。
設定はダイアログボックスでプレビューしながら決め、プリセットとして保存できます。

Draws an underline, an L shape, a bracket, a box or top and bottom lines with a speech-bubble tail (a V, a slash, a taper, a U, a loop, a horn or a thought bubble) around the selected text. The line can go on any side.
Settings are previewed in a dialog box and can be saved as presets.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したテキストに、吹き出しのしっぽ（V字・斜線・すぼまり・U字・ループ・つの・思考）の付いた下線・L字・コの字・囲み・上下の線を描きます。線は上下左右のどの側にも描けます。
設定はダイアログボックスでプレビューしながら決め、プリセットとして保存できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SpeechBubbleMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8849614c960c

### Overview

Draws an underline, an L shape, a bracket, a box or top and bottom lines with a speech-bubble tail (a V, a slash, a taper, a U, a loop, a horn or a thought bubble) around the selected text. The line can go on any side.
Settings are previewed in a dialog box and can be saved as presets.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SpeechBubbleMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SpeechBubbleMaker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "テキストに吹き出しのしっぽ付きの下線・囲みを描く"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SpeechBubbleMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SpeechBubbleMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8849614c960c"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SpeechBubbleMaker.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 初めて開いたときの選択 / Initial choices */
    var DEFAULT_CHOICES = {
        shapeKey: "notch",          /* しっぽの形状（"notch" V字 / "slash" 斜線 / "curly" すぼまり / "round" U字 / "loop" ループ / "horn" つの / "thought" 思考） / tail shape */
        directionKey: "center",     /* しっぽの向き（"center" まっすぐ / "right" 右へ / "left" 左へ） / tail direction */
        rootKey: "center",          /* 斜線の根元（"center" 切れ目の中央 / "edge" 切れ目の端） / slash root */
        frameKey: "underline",      /* 線の形状（"underline" 下線 / "lShape" L字 / "bracket" コの字 / "box" 囲み / "twoLines" 上下） / frame shape */
        sideKey: "auto",            /* 線を描く側（"auto" は開くたびにテキストから決める＝横組みは下・縦組みは右 / "top" / "bottom" / "left" / "right"） / side to draw on; "auto" is resolved from the text each time */
        capKey: "butt",             /* 線端（"butt" なし / "round" 丸型） / stroke cap */
        colorKey: "text",           /* 線のカラー（"text" 文字色＝1文字目の塗り / "custom" カスタム） / line color */
        customColor: { model: "gray", values: [100] }, /* カスタムのカラー（"rgb" 0〜255 / "cmyk" 0〜100 / "gray" 0〜100） / custom color */
        position: 50,               /* しっぽの位置（線の始まりから％） / tail position (% from the start) */
        slashAngle: 45,             /* 斜線の角度（線から測る。90°で真下） / slash angle from the line (90° = straight) */
        useAreaFrame: true,         /* エリア内文字は字面ではなく枠を基準にする / area type is measured by its frame, not its glyphs */
        isCapsule: false,           /* カプセル形状（角丸の半径を最大にする） / capsule shape (the largest corner radius) */
        tipDot: false,              /* しっぽの先端に●を置く / dot at the tip of the tail */
        convertToAreaText: false,   /* 囲みのとき、テキストをエリア内文字にして中央に置く / make the text area type centered in the box */
        groupWithText: false,       /* テキストとグループ化する / group with the text */
        preview: true               /* プレビューする / preview */
    };

    /* 初めて開いたときの寸法。1文字目の文字サイズに対する比率
       Initial sizes, as ratios of the font size of the first character */
    var DEFAULT_SIZE_RATIOS = {
        gapWidth: 0.5,        /* 切れ目（V字の付け根の幅） / gap (base width of the V) */
        tailWidth: 0.5,       /* しっぽの横方向の長さ / horizontal run of the tail */
        tailHeight: 0.3,      /* しっぽの高さ / tail height */
        offset: 0.3,          /* 字面から線まで / distance from the glyphs to the line */
        extension: 0,         /* 字面の左右から線を伸ばす長さ / extension beyond the glyphs */
        tickHeight: 0.3,      /* コの字の立ち上がり / bracket tick height */
        cornerRadius: 0.2,    /* 角丸の半径 / corner radius */
        shadowDistance: 0,    /* 影の間隔（0 で影なし） / shadow distance (0 = none) */
        sideGapWidth: 0,      /* 囲みの左右の辺の中央の切れ目（0 で切らない） / gaps in the middle of the box's sides (0 = none) */
        slashShiftX: 0,       /* 線と離した斜線の左右の調整 / horizontal nudge of a detached slash */
        slashShiftY: 0,       /* 線と離した斜線の上下の調整 / vertical nudge of a detached slash */
        strokeWidth: 0.04     /* 線幅 / stroke weight */
    };

    // =========================================
    // プリセット / Presets
    // =========================================

    /* 最初からあるプリセット。定番、下線、コの字・L字、囲み、上下の順に並べ、null の所に区切り線を入れる
       寸法は1文字目の文字サイズに対する比率（書かない項目は DEFAULT_SIZE_RATIOS、選択は DEFAULT_CHOICES）
       Built-in presets: the staple first, then underlines, brackets and L shapes, boxes, and top & bottom lines; null draws a separator.
       Sizes are ratios of the first character's font size; missing items fall back to the defaults */
    var BUILTIN_PRESETS = [
        {
            /* 保存したプリセットから昇格（12pt・mm で作ったものを比率に直した） / promoted from a saved preset made at 12 pt in mm, converted to ratios */
            name: { ja: "シンプルな吹き出し", en: "Simple speech bubble" },
            choices: { shapeKey: "notch", directionKey: "center", rootKey: "edge", frameKey: "box", capKey: "butt" },
            position: 72,
            ratios: { gapWidth: 0.47, tailWidth: 0.59, tailHeight: 0.47, offset: 0.5, extension: 0.5, tickHeight: 0.31, cornerRadius: 0.31, shadowDistance: 0, strokeWidth: 0.04 }
        },
        {
            /* 保存したプリセットから昇格（12pt・mm で作ったものを比率に直した） / promoted from a saved preset made at 12 pt in mm, converted to ratios */
            name: { ja: "ひろがりカーリー", en: "Wide curly" },
            choices: { shapeKey: "curly", directionKey: "center", rootKey: "edge", frameKey: "underline", capKey: "round" },
            position: 48,
            ratios: { gapWidth: 1.65, tailWidth: 0.59, tailHeight: 0.47, offset: 0.47, extension: 0.24, tickHeight: 0.59, cornerRadius: 0, shadowDistance: 0, strokeWidth: 0.04 }
        },
        null, /* 区切り線 / separator */
        {
            name: { ja: "下線・V字", en: "Underline, V" },
            choices: { shapeKey: "notch", directionKey: "center", frameKey: "underline" },
            ratios: { gapWidth: 0.55, tailHeight: 0.28 }
        },
        {
            name: { ja: "下線・V字（右下）", en: "Underline, V down right" },
            choices: { shapeKey: "notch", directionKey: "right", frameKey: "underline" },
            position: 75,
            ratios: { gapWidth: 0.4, tailWidth: 0.6, tailHeight: 0.5 }
        },
        {
            name: { ja: "下線・斜線（右下）", en: "Underline, slash down right" },
            choices: { shapeKey: "slash", directionKey: "right", rootKey: "center", frameKey: "underline", capKey: "round" },
            ratios: { gapWidth: 0.5, tailWidth: 0.4, tailHeight: 0.4 }
        },
        {
            /* 保存したプリセットから昇格（12pt・pt で作ったものを比率に直した） / promoted from a saved preset made at 12 pt, converted to ratios */
            name: { ja: "下線（斜線・右下）切れ目なし", en: "Underline, slash down right, no gap" },
            choices: { shapeKey: "slash", directionKey: "right", rootKey: "center", frameKey: "underline", capKey: "butt" },
            position: 74,
            ratios: { gapWidth: 0, tailWidth: 0.4, tailHeight: 0.58, offset: 0.3, extension: 0.17, tickHeight: 0.3, cornerRadius: 0.2, shadowDistance: 0, strokeWidth: 0.04 }
        },
        {
            name: { ja: "下線・斜線（左下）", en: "Underline, slash down left" },
            choices: { shapeKey: "slash", directionKey: "left", rootKey: "center", frameKey: "underline", capKey: "round" },
            ratios: { gapWidth: 0.5, tailWidth: 0.4, tailHeight: 0.4 }
        },
        {
            name: { ja: "下線・U字", en: "Underline, U" },
            choices: { shapeKey: "round", directionKey: "center", frameKey: "underline", capKey: "round" },
            ratios: { gapWidth: 0.5, tailHeight: 0.35 }
        },
        {
            /* 広い切れ目から曲線でおりて輪を描く / sweeps down from a wide gap into a loop */
            name: { ja: "下線・ループ", en: "Underline, loop" },
            choices: { shapeKey: "loop", directionKey: "center", frameKey: "underline", capKey: "round" },
            ratios: { gapWidth: 1.6, tailHeight: 0.6, extension: 0.1 }
        },
        null, /* 区切り線 / separator */
        {
            name: { ja: "コの字・斜線（右下）", en: "Bracket, slash down right" },
            choices: { shapeKey: "slash", directionKey: "right", rootKey: "edge", frameKey: "bracket" },
            position: 55,
            ratios: { gapWidth: 0.6, tailWidth: 0.6, tailHeight: 0.6, extension: 0.6, tickHeight: 0.6, cornerRadius: 0 }
        },
        {
            name: { ja: "コの字（角丸）・V字（右下）", en: "Rounded bracket, V down right" },
            choices: { shapeKey: "notch", directionKey: "right", frameKey: "bracket" },
            position: 75,
            ratios: { gapWidth: 0.4, tailWidth: 0.6, tailHeight: 0.5, extension: 0.4, tickHeight: 0.4, cornerRadius: 0.25 }
        },
        {
            name: { ja: "カーリーブラケット（波括弧）", en: "Curly bracket" },
            choices: { shapeKey: "curly", directionKey: "center", frameKey: "bracket" },
            ratios: { gapWidth: 0.4, tailHeight: 0.2, extension: 0.3, tickHeight: 0.25, cornerRadius: 0.25 }
        },
        {
            /* 左へ出すV字の先端を付け根の左端の真下に置き、しっぽの側だけ立ち上げる / the V's tip sits under the left of its base; only the tail's end turns up */
            name: { ja: "L字・V字（左下）", en: "L shape, V down left" },
            choices: { shapeKey: "notch", directionKey: "left", frameKey: "lShape", capKey: "round" },
            position: 15,
            ratios: { gapWidth: 0.4, tailWidth: 0.4, tailHeight: 0.5, extension: 0.4, tickHeight: 0.4, cornerRadius: 0.15 }
        },
        null, /* 区切り線 / separator */
        {
            name: { ja: "囲み・斜線（右下）", en: "Box, slash down right" },
            choices: { shapeKey: "slash", directionKey: "right", rootKey: "edge", frameKey: "box" },
            ratios: { gapWidth: 0.6, tailWidth: 0.6, tailHeight: 0.6, offset: 0.5, extension: 0.5, cornerRadius: 0 }
        },
        {
            name: { ja: "囲み（角丸）・斜線（左下）", en: "Rounded box, slash down left" },
            choices: { shapeKey: "slash", directionKey: "left", rootKey: "edge", frameKey: "box" },
            ratios: { gapWidth: 0.6, tailWidth: 0.6, tailHeight: 0.6, offset: 0.5, extension: 0.5, cornerRadius: 0.3 }
        },
        {
            /* 保存したプリセットから昇格（12pt・pt で作ったものを比率に直した） / promoted from a saved preset made at 12 pt, converted to ratios */
            name: { ja: "囲み（カプセル）", en: "Capsule box" },
            choices: { shapeKey: "slash", directionKey: "right", rootKey: "edge", frameKey: "box", capKey: "butt" },
            position: 70,
            isCapsule: true,
            ratios: { gapWidth: 0.6, tailWidth: 0.6, tailHeight: 0.6, offset: 0.5, extension: 0.75, tickHeight: 0.3, cornerRadius: 0, shadowDistance: 0, strokeWidth: 0.04 }
        },
        {
            /* V字の付け根の右端を囲みの右下の角に合わせ、先端を角より外へ出す / the V's base ends at the bottom-right corner and its tip reaches past it */
            name: { ja: "囲み・右下角にしっぽ", en: "Box, tail at bottom-right corner" },
            choices: { shapeKey: "notch", directionKey: "right", frameKey: "box", capKey: "butt" },
            position: 100,
            ratios: { gapWidth: 0.5, tailWidth: 0.7, tailHeight: 0.6, offset: 0.5, extension: 0.5, cornerRadius: 0 }
        },
        {
            /* 左下の角に合わせた左右反転 / mirrored to the bottom-left corner */
            name: { ja: "囲み・左下角にしっぽ", en: "Box, tail at bottom-left corner" },
            choices: { shapeKey: "notch", directionKey: "left", frameKey: "box", capKey: "butt" },
            position: 0,
            ratios: { gapWidth: 0.5, tailWidth: 0.7, tailHeight: 0.6, offset: 0.5, extension: 0.5, cornerRadius: 0 }
        },
        {
            name: { ja: "囲み・V字（影）", en: "Box, V with shadow" },
            choices: { shapeKey: "notch", directionKey: "left", frameKey: "box" },
            position: 65,
            ratios: { gapWidth: 0.6, tailWidth: 0.9, tailHeight: 0.8, offset: 0.5, extension: 0.5, cornerRadius: 0, shadowDistance: 0.25 }
        },
        {
            name: { ja: "囲み（カプセル）・切れ目", en: "Capsule box with gaps" },
            choices: { shapeKey: "slash", directionKey: "left", rootKey: "edge", frameKey: "box", capKey: "round" },
            position: 20,
            isCapsule: true,
            ratios: { gapWidth: 0.6, tailWidth: 0.6, tailHeight: 0.6, offset: 0.5, extension: 0.6, sideGapWidth: 0.3, strokeWidth: 0.06 }
        },
        {
            name: { ja: "囲み（カプセル）・斜線の先に●", en: "Capsule box, slash with a dot" },
            choices: { shapeKey: "slash", directionKey: "right", rootKey: "center", frameKey: "box", capKey: "round" },
            position: 85,
            isCapsule: true,
            tipDot: true,
            ratios: { gapWidth: 0.5, tailHeight: 0.6, offset: 0.5, extension: 0.6, strokeWidth: 0.06 }
        },
        {
            /* ふくらみの左下に円を置く / the circle sits down left of the bump */
            name: { ja: "思考の吹き出し", en: "Thought bubble" },
            choices: { shapeKey: "thought", directionKey: "left", frameKey: "box", capKey: "round" },
            position: 70,
            isCapsule: true,
            ratios: { gapWidth: 0.85, tailWidth: 1.2, tailHeight: 1.2, offset: 1.1, extension: 1.8, strokeWidth: 0.07 }
        },
        {
            name: { ja: "囲み（角丸）・つの", en: "Rounded box, horn" },
            choices: { shapeKey: "horn", directionKey: "left", frameKey: "box", capKey: "butt" },
            position: 60,
            ratios: { gapWidth: 0.6, tailWidth: 0.9, tailHeight: 0.9, offset: 0.5, extension: 0.5, cornerRadius: 0.3 }
        },
        null, /* 区切り線 / separator */
        {
            /* 斜線を切れ目の右端から左下へ出す / the slash leaves the right of the gap toward the bottom left */
            name: { ja: "上下・斜線（左下）", en: "Top & bottom, slash down left" },
            choices: { shapeKey: "slash", directionKey: "left", rootKey: "edge", frameKey: "twoLines", capKey: "butt" },
            position: 52,
            ratios: { gapWidth: 0.6, tailHeight: 0.6, offset: 0.8, extension: 0.5 }
        }
    ];

    /* プリセットに入れる選択と寸法の項目 / Choices and sizes a preset holds */
    var PRESET_CHOICE_KEYS = ["shapeKey", "directionKey", "rootKey", "frameKey", "sideKey", "capKey"];
    var PRESET_SIZE_KEYS = ["gapWidth", "tailWidth", "tailHeight", "offset", "extension", "tickHeight", "cornerRadius", "shadowDistance", "sideGapWidth", "strokeWidth", "slashShiftX", "slashShiftY"];

    var MIN_SLASH_ANGLE = 5;       /* 斜線の角度の下限（度） / minimum slash angle (degrees) */
    var TAIL_MITER_LIMIT = 10;     /* V字の角をとがらせたままにする比率の上限 / miter limit that keeps the V corners pointed */
    var TAIL_TOLERANCE = 0.001;    /* 座標を同じと見なす許容値（pt） / tolerance for equal coordinates (pt) */
    var CORNER_HANDLE_RATIO = 0.5523; /* 角丸（4分の1円）のハンドルの長さの比率 / handle ratio of a quarter-circle corner */
    var MIN_AREA_TEXT_SIZE = 1;    /* エリア内文字の枠の最小の幅・高さ（pt） / minimum area type frame size (pt) */
    var THOUGHT_FILLET_RATIO = 0.4;  /* 思考のふくらみの付け根の丸み（ふくらみの半径に対する比率） / fillet of the thought bump, relative to its radius */
    var THOUGHT_CIRCLE_RATIO = 0.9;  /* 思考の円の半径（ふくらみの半径に対する比率） / radius of the thought circle, relative to the bump */
    var LOOP_RADIUS_RATIO = 0.32;    /* ループの輪の半径（高さに対する比率） / radius of the loop, relative to the height */
    var LOOP_CROSS_RATIO = 1.2;      /* ループの交差点の輪の中心からの高さ（輪の半径に対する比率） / height of the loop's crossing above its center, relative to its radius */
    var HORN_ROOT_HANDLE_RATIO = 0.5; /* つのの根元から真下へ出すハンドルの長さ（高さに対する比率） / handle straight out of the horn's root, relative to the height */
    var HORN_TIP_HANDLE_RATIO = 0.35; /* つのの先端のハンドルの長さ（根元への距離に対する比率） / handle at the horn's tip, relative to the way back to the root */
    var TIP_DOT_RATIO = 3;           /* 先端の●の直径（線幅に対する比率） / diameter of the tip dot, relative to the weight */
    var CUT_SAMPLE_COUNT = 32;       /* 切れ目の長さを測るときの曲線の分割数 / samples per curve when measuring the side gaps */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var FIELD_CHARACTERS = 7;        /* 数値欄の文字数 / characters of the numeric fields */
    var POSITION_SLIDER_RIGHT_INSET = 24; /* 位置のスライダーの右にあける余白 / room left at the right of the position slider */
    var SLASH_SHIFT_STEP = 0.5;      /* 斜線の左右・上下の調整の∧∨の刻み（欄の単位） / stepper step of the slash nudges (field unit) */
    var POSITION_SLIDER_SHIFT_STEP = 10; /* shift を押しながら動かしたときの位置のスライダーの刻み（％） / position slider step with Shift (%) */
    var POSITION_SLIDER_WIDTH = 100; /* 位置のスライダーの最小の幅（パネルの幅まで伸びる） / minimum width of the position slider (it stretches to the panel) */
    var SIDE_GRID_CELL_WIDTH = (String($.locale || "").indexOf("ja") === 0) ? 48 : 72; /* 配置の3行3列の1マスの幅（英語は語が長い） / cell width of the 3x3 side grid (English words are longer) */
    var SHAPE_GRID_CELL_WIDTH = (String($.locale || "").indexOf("ja") === 0) ? 72 : 80; /* しっぽの形状の3列の1マスの幅 / cell width of the 3-column tail shape grid */
    var COLOR_SWATCH_SIZE = [36, 18]; /* 線のカラーの色見本の大きさ / size of the line color swatch */
    var STACKED_RADIO_SPACING = 4;   /* 縦に並べたラジオボタンの間隔 / spacing of stacked radio buttons */
    var RADIO_GRID_SPACING = 2;      /* 表の形に並べたラジオボタンの行の間隔 / row spacing of radio button grids */
    var PRESET_SAVE_ICON_SIZE = [22, 22]; /* 保存アイコンの大きさ。横長の絵なので既定（24×22）より少し狭くして、ごみ箱と見た目をそろえる / the save icon is wide, so a little narrower than the default to match the trash icon */
    var PRESET_DROPDOWN_WIDTH = 220; /* プリセットのドロップダウンの幅 / width of the preset dropdown */
    var PRESET_ROW_SPACING = 6;      /* プリセットの行の間隔 / spacing in the preset row */
    var PRESET_NAME_CHARACTERS = 20; /* プリセット名の入力欄の文字数 / characters of the preset name field */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "吹き出しメーカー", en: "Speech Bubble Maker" },
            presetSave: { ja: "プリセットを保存", en: "Save Preset" }
        },
        panel: {
            tail: { ja: "しっぽ", en: "Tail" },
            line: { ja: "線/枠", en: "Line/Frame" },
            options: { ja: "テキストに関してのオプション", en: "Text Options" }
        },
        radio: {
            notch: { ja: "V字", en: "V shape" },
            slash: { ja: "斜線", en: "Slash" },
            curly: { ja: "すぼまり", en: "Taper" },
            round: { ja: "U字", en: "U shape" },
            loop: { ja: "ループ", en: "Loop" },
            horn: { ja: "つの", en: "Horn" },
            thought: { ja: "思考", en: "Thought" },
            straightDown: { ja: "まっすぐ", en: "Straight" },
            downRight: { ja: "右へ", en: "Right" },
            downLeft: { ja: "左へ", en: "Left" },
            rootCenter: { ja: "中央", en: "Center" },
            rootEdge: { ja: "端", en: "Edge" },
            underline: { ja: "下線", en: "Underline" },
            lShape: { ja: "L字", en: "L shape" },
            bracket: { ja: "コの字", en: "Bracket" },
            box: { ja: "囲み", en: "Box" },
            twoLines: { ja: "上下", en: "Top & bottom" },
            sideTop: { ja: "上", en: "Top" },
            sideBottom: { ja: "下", en: "Bottom" },
            sideLeft: { ja: "左", en: "Left" },
            sideRight: { ja: "右", en: "Right" },
            buttCap: { ja: "なし", en: "Butt" },
            roundCap: { ja: "丸型", en: "Round" },
            colorText: { ja: "文字色", en: "Text" },
            colorCustom: { ja: "カスタム", en: "Custom" } /* 画面には出さない（ツールチップ用の名前） / not shown; the swatch stands in for it */
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            presetName: { ja: "名前", en: "Name" },
            shape: { ja: "形状", en: "Shape" },
            direction: { ja: "向き", en: "Direction" },
            root: { ja: "根元", en: "Root" },
            slashShiftX: { ja: "左右の調整", en: "Nudge X" },
            slashShiftY: { ja: "上下の調整", en: "Nudge Y" },
            position: { ja: "位置", en: "Position" },
            gapWidth: { ja: "切れ目", en: "Gap" },
            tailWidth: { ja: "幅", en: "Width" },
            slashAngle: { ja: "角度", en: "Angle" },
            tailHeight: { ja: "高さ", en: "Height" },
            frame: { ja: "形状", en: "Shape" },
            side: { ja: "配置", en: "Side" },
            offset: { ja: "間隔", en: "Distance" },
            extension: { ja: "左右の伸び", en: "Extension" },
            tickHeight: { ja: "立ち上がり", en: "Ticks" },
            cornerRadius: { ja: "角丸の半径", en: "Corner radius" },
            shadowDistance: { ja: "影", en: "Shadow" },
            sideGapWidth: { ja: "左右の切れ目", en: "Side gaps" },
            strokeWidth: { ja: "線幅", en: "Weight" },
            cap: { ja: "線端", en: "Cap" },
            lineColor: { ja: "線のカラー", en: "Color" }
        },
        checkbox: {
            useAreaFrame: { ja: "エリア内文字は枠を基準にする", en: "Use the frame of area type" },
            isCapsule: { ja: "カプセル形状", en: "Capsule" },
            tipDot: { ja: "先端に●", en: "Dot at the tip" },
            convertToAreaText: { ja: "エリア内文字にして中央に配置", en: "Center as area type" },
            groupWithText: { ja: "テキストとグループ化", en: "Group with text" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            preset: {
                ja: "選ぶと設定をまとめて切り替えます。最初からあるプリセットの寸法は、1文字目の文字サイズに合わせます。",
                en: "Switches all the settings at once. Built-in presets size themselves to the font size of the first character."
            },
            presetSave: {
                ja: "今の設定に名前を付けて、プリセットとして保存します。",
                en: "Saves the current settings as a named preset."
            },
            presetDelete: {
                ja: "選んでいる保存したプリセットを削除します（最初からあるものは削除できません）。",
                en: "Deletes the selected saved preset (built-in presets cannot be deleted)."
            },
            presetName: { ja: "プリセットの名前です。", en: "Name of the preset." },
            notch: { ja: "線をV字に折ってしっぽを作ります。", en: "Bends the line into a V." },
            slash: { ja: "線から斜線を出します。", en: "Draws a slash out of the line." },
            curly: {
                ja: "カーリーブラケット（波括弧）の中央のように、線から曲線ですぼまってとがらせます。",
                en: "Tapers out of the line in curves to a point, like the middle of a curly bracket."
            },
            round: { ja: "線をU字に曲げてしっぽを作ります（［高さ］が深さ）。", en: "Bends the line into a U (Height is its depth)." },
            loop: {
                ja: "線から曲線でおりて、先で小さな輪を描きます。［高さ］は輪の下端までです。",
                en: "Curves out of the line into a small loop. Height reaches the bottom of the loop."
            },
            horn: {
                ja: "根元が太く先がとがった塗りのしっぽを、線に重ねます（線幅プロファイルのような形）。［切れ目］が根元の幅です。",
                en: "Lays a filled tail over the line, wide at the root and pointed at the end, like a width profile. Gap is the root width."
            },
            thought: {
                ja: "線に小さなふくらみを付け、その先に円を置きます（考えごとの吹き出し）。［切れ目］がふくらみの幅、［高さ］は円の下端までです。",
                en: "Adds a small bump to the line and a circle beyond it, as in a thought bubble. Gap is the bump's width; Height reaches the bottom of the circle."
            },
            straightDown: {
                ja: "V字・すぼまりの先端を切れ目の中央に置き、斜線は線と直角に出します（配置が下なら真下、上なら真上）。",
                en: "Puts the tip of a V or taper at the middle of the gap; a slash leaves the line at a right angle (straight down on the bottom, straight up on the top)."
            },
            downRight: {
                ja: "しっぽを右下へ出します（配置が上なら右上、左右なら下寄り）。",
                en: "The tail points down to the right (up to the right on the top, toward the bottom on the left or right)."
            },
            downLeft: {
                ja: "しっぽを左下へ出します（配置が上なら左上、左右なら上寄り）。",
                en: "The tail points down to the left (up to the left on the top, toward the top on the left or right)."
            },
            rootCenter: {
                ja: "斜線を切れ目の中央から出し、線とは離します。",
                en: "The slash starts at the middle of the gap, apart from the line."
            },
            rootEdge: {
                ja: "斜線を切れ目の端から出し、線とつなげます。",
                en: "The slash starts at the edge of the gap, joined to the line."
            },
            slashShiftX: {
                ja: "線と離した斜線を、線に沿ってずらします（プラスで右、配置が左右なら下）。向きを選ぶと、［右へ］は線幅の半分のマイナス、［左へ］はプラスが入ります。",
                en: "Shifts a detached slash along the line (positive moves it right, or down on the left or right side). Picking a direction sets it to minus half the weight for Right, plus for Left."
            },
            slashShiftY: {
                ja: "線と離した斜線を、線と直角にずらします（プラスで外側＝しっぽの出る向き、マイナスで文字の側）。",
                en: "Shifts a detached slash across the line (positive moves it outward, the way the tail points; negative toward the text)."
            },
            position: {
                ja: "線の始まりを0%、終わりを100%として、しっぽの位置を指定します。",
                en: "Where the tail sits along the line, from 0% (start) to 100% (end)."
            },
            positionSlider: {
                ja: "しっぽの位置を動かします（shift を押しながらで10%刻み）。",
                en: "Moves the tail along the line (hold Shift for 10% steps)."
            },
            gapWidth: {
                ja: "しっぽのところで線を切る幅です。V字・すぼまりは付け根の幅です。斜線は 0 で線を切りません。",
                en: "How much of the line is cut away at the tail: the base of a V or taper, or the gap at a slash (0 keeps the line whole)."
            },
            tailWidth: {
                ja: "向きが［右へ］［左へ］のときの、先端（思考は円、ループは輪）の横方向のずれです。付け根の反対側の端から測ります。",
                en: "How far the tip (the circle of a thought, the loop of a loop) moves sideways when Direction is Right or Left, measured from the far side of the base."
            },
            slashAngle: {
                ja: "向きが［右へ］［左へ］の斜線の、線から測った角度です（90°で線と直角）。長さは［高さ］で決まります。",
                en: "Angle of a slash when Direction is Right or Left, measured from the line (90° is at a right angle to it). Its length follows Height."
            },
            tailHeight: { ja: "線から外へ伸びるしっぽの長さです。", en: "How far the tail reaches out from the line." },
            underline: { ja: "字面の下に線を引きます。", en: "Draws a line below the glyphs." },
            bracket: { ja: "線の両端を文字の側へ立ち上げます。", en: "Turns both ends of the line up toward the text." },
            box: { ja: "字面を長方形で囲みます。", en: "Encloses the glyphs in a rectangle." },
            lShape: {
                ja: "しっぽに近いほうの端だけを、文字の側へ立ち上げます（位置が50%以下なら始まり側）。",
                en: "Turns up only the end nearer the tail toward the text (the start when Position is 50% or less)."
            },
            twoLines: {
                ja: "字面の両側に線を引きます。しっぽは配置の側の線に付きます。",
                en: "Draws lines on both sides of the glyphs; the tail sits on the line on the chosen side."
            },
            sideTop: { ja: "テキストの上に描きます（しっぽは上へ）。", en: "Draws above the text (the tail points up)." },
            sideBottom: { ja: "テキストの下に描きます。", en: "Draws below the text." },
            sideLeft: {
                ja: "テキストの左に描きます（線の始まりは上）。",
                en: "Draws left of the text (the line starts at the top)."
            },
            sideRight: {
                ja: "テキストの右に描きます（線の始まりは上）。",
                en: "Draws right of the text (the line starts at the top)."
            },
            offset: {
                ja: "字面から線までの距離です。囲みでは上下の余白です。",
                en: "Distance from the glyphs to the line; the top and bottom padding of a box."
            },
            extension: {
                ja: "字面の端から線を伸ばす長さです（マイナスで縮めます）。囲みでは左右の余白です。",
                en: "How far the line runs past the glyphs (negative values shorten it); the side padding of a box."
            },
            tickHeight: { ja: "コの字の両端を立ち上げる長さです。", en: "Length of the bracket's turned-up ends." },
            cornerRadius: {
                ja: "コの字・囲みの角の丸みです。0 で角をとがらせます。カプセル形状がオンのときは使いません。",
                en: "Corner radius of a bracket or box. 0 keeps the corners sharp. Unused while Capsule is on."
            },
            shadowDistance: {
                ja: "囲みの右と下に、この間隔で線を重ねます。0 で影なし。",
                en: "Adds a second line this far out on the right and bottom of a box. 0 for none."
            },
            sideGapWidth: {
                ja: "囲みの左右の辺の中央を、この長さだけ切ります。0 で切りません。",
                en: "Cuts this much out of the middle of the box's left and right sides. 0 for none."
            },
            strokeWidth: { ja: "線としっぽの線幅です。", en: "Stroke weight of the line and tail." },
            buttCap: { ja: "線の端をパスの端で切ります。", en: "Ends the stroke flush with the path ends." },
            roundCap: { ja: "線の端を丸めます（角も丸くなります）。", en: "Rounds the stroke ends (and the corners)." },
            colorText: {
                ja: "テキストごとに、1文字目の塗りと同じカラーにします。",
                en: "Uses the fill of each text's first character."
            },
            colorCustom: { ja: "右の色見本のカラーにします。", en: "Uses the color of the swatch on the right." },
            colorSwatch: {
                ja: "クリックすると、カラーピッカーで色を選べます。",
                en: "Click to choose a color in the Color Picker."
            },
            isCapsule: {
                ja: "角丸の半径をできるだけ大きくして、囲みの左右（コの字の両端）を半円にします。",
                en: "Uses the largest corner radius, so the ends of a box (or bracket) become half circles."
            },
            tipDot: {
                ja: "しっぽの先端に●を置きます（V字・斜線・すぼまり・U字）。直径は線幅の3倍です。",
                en: "Puts a dot at the tip of the tail (V, slash, taper, U). Its diameter is three times the weight."
            },
            useAreaFrame: {
                ja: "エリア内文字は、字面ではなく枠に沿って描きます（［間隔］［左右の伸び］は 0 として扱います）。",
                en: "Draws along the frame of area type instead of its glyphs (Distance and Extension count as 0)."
            },
            convertToAreaText: {
                ja: "囲みのとき、テキストを長方形と同じ大きさのエリア内文字にして、左右・天地中央に揃えます。",
                en: "With a box, turns the text into area type the size of the rectangle, centered both ways."
            },
            groupWithText: {
                ja: "線としっぽをテキストと一緒にグループにします。",
                en: "Groups the line and tail together with the text."
            },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel removes it."
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
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        confirm: {
            presetOverwrite: { ja: "「{name}」を上書きしますか？", en: "Overwrite \"{name}\"?" },
            presetDelete: { ja: "「{name}」を削除しますか？", en: "Delete \"{name}\"?" }
        },
        alert: {
            presetSaveFailed: { ja: "プリセットを保存できませんでした。", en: "Could not save the presets." },
            presetBuiltinName: {
                ja: "最初からあるプリセットと同じ名前は使えません。",
                en: "That name is used by a built-in preset."
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noText: { ja: "テキストを選択してください。", en: "Select some text." },
            skipped: {
                ja: "%1 個のテキストには描けませんでした（ロックされたグループの中、線の長さが 0 など）。",
                en: "Could not draw for %1 text object(s) (inside a locked group, a line of zero length, etc.)."
            },
            areaTextFailed: {
                ja: "エリア内文字の天地中央揃えを設定できませんでした。",
                en: "Could not center the area type vertically."
            }
        },
        fallbackName: {
            tailGroup: { ja: "吹き出しのしっぽ", en: "Speech bubble tail" },
            bubbleGroup: { ja: "吹き出し", en: "Speech bubble" }
        }
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
    // 設定の記憶 / Session memory
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
    // 線としっぽの形 / Line and tail geometry
    // =========================================

    /**
     * @typedef {Object} BubbleSettings
     * @property {string} shapeKey - しっぽの形状。"notch"（V字）/ "slash"（斜線）/ "curly"（すぼまり）/ "round"（U字）/ "loop"（ループ）/ "horn"（つの）/ "thought"（思考）
     * @property {string} directionKey - しっぽの向き。"center"（まっすぐ）/ "right"（右へ）/ "left"（左へ）
     * @property {string} rootKey - 斜線の根元。"center"（切れ目の中央）/ "edge"（切れ目の端）
     * @property {string} frameKey - 線の形状。"underline"（下線）/ "lShape"（L字）/ "bracket"（コの字）/ "box"（囲み）/ "twoLines"（上下）
     * @property {string} sideKey - 線を描く側。"top" / "bottom" / "left" / "right"（"auto" は横組みは下・縦組みは右）
     * @property {number} position - しっぽの位置（線の始まりから％）
     * @property {number} slashAngle - 斜線の角度（線から測った度。90 で真下）
     * @property {number} slashShiftX - 線と離した斜線を線に沿ってずらす量（pt、プラスで線の進む向き）
     * @property {number} slashShiftY - 線と離した斜線を線と直角にずらす量（pt、プラスで外側）
     * @property {number} gapWidth - しっぽのところで線を切る幅（pt）
     * @property {number} tailWidth - 向きが右へ・左へのV字・すぼまりの、先端の横方向のずれ（pt）
     * @property {number} tailHeight - しっぽが線から外へ伸びる長さ（pt）
     * @property {number} offset - 字面から線までの距離（pt）
     * @property {number} extension - 字面の端から線を伸ばす長さ（pt）
     * @property {number} tickHeight - コの字の立ち上がり（pt）
     * @property {number} cornerRadius - 角丸の半径（pt）
     * @property {number} shadowDistance - 影の間隔（pt、0 で影なし）
     * @property {number} sideGapWidth - 囲みの左右の辺の中央の切れ目（pt、0 で切らない）
     * @property {number} strokeWidth - 線幅（pt）
     * @property {string} capKey - 線端。"butt"（なし）/ "round"（丸型）
     * @property {string} colorKey - 線のカラー。"text"（文字色）/ "custom"（カスタム）
     * @property {{model: string, values: number[]}} customColor - カスタムのカラー（"rgb" / "cmyk" / "gray" と値）
     * @property {boolean} useAreaFrame - エリア内文字は枠を基準にするか
     * @property {boolean} isCapsule - カプセル形状（角丸の半径を、辺の長さで頭打ちになるまで大きくする）か
     * @property {boolean} tipDot - しっぽの先端に●を置くか
     * @property {boolean} convertToAreaText - 囲みのとき、テキストをエリア内文字にして中央に置くか
     * @property {boolean} groupWithText - テキストとグループ化するか
     * @property {boolean} preview - プレビューするか
     */

    /**
     * @typedef {Object} OutlineVertex
     * @property {number} x - x 座標
     * @property {number} y - y 座標
     * @property {number} radius - 角丸の半径（0 で角のまま）
     * @property {number[]} [leftHandle] - 前から入るハンドル [x, y]（すぼまりの曲線用。省略時はハンドルなし）
     * @property {number[]} [rightHandle] - 次へ出るハンドル [x, y]（同上）
     */

    /**
     * 値を下限・上限の範囲に収める（下限が上限を超えるときは下限を優先）
     * @param {number} value - 数値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲に収めた値
     */
    function clampNumber(value, minValue, maxValue) {
        if (value > maxValue) value = maxValue;
        if (value < minValue) value = minValue;
        return value;
    }

    /**
     * 頂点を作る
     * @param {number} x - x 座標
     * @param {number} y - y 座標
     * @param {number} [radius] - 角丸の半径
     * @param {number[]} [leftHandle] - 前から入るハンドル [x, y]
     * @param {number[]} [rightHandle] - 次へ出るハンドル [x, y]
     * @returns {OutlineVertex} 頂点
     */
    function makeVertex(x, y, radius, leftHandle, rightHandle) {
        return { x: x, y: y, radius: radius || 0, leftHandle: leftHandle || null, rightHandle: rightHandle || null };
    }

    /**
     * 2つの頂点が同じ位置か
     * @param {OutlineVertex} vertexA - 頂点A
     * @param {OutlineVertex} vertexB - 頂点B
     * @returns {boolean} 同じなら true
     */
    function isSameVertex(vertexA, vertexB) {
        return Math.abs(vertexA.x - vertexB.x) <= TAIL_TOLERANCE && Math.abs(vertexA.y - vertexB.y) <= TAIL_TOLERANCE;
    }

    /**
     * 続けて同じ位置にある頂点を1つにまとめる（角丸の半径は大きいほうを残す）
     * @param {OutlineVertex[]} vertices - 頂点の配列
     * @param {boolean} isClosed - 閉じたパスなら true（最後と最初も比べる）
     * @returns {OutlineVertex[]} 重なりを除いた頂点の配列
     */
    function removeDuplicateVertices(vertices, isClosed) {
        var uniqueVertices = [];
        for (var i = 0; i < vertices.length; i++) {
            var lastVertex = uniqueVertices[uniqueVertices.length - 1];
            if (lastVertex && isSameVertex(lastVertex, vertices[i])) {
                /* 入りのハンドルは前、出のハンドルは後ろから取る / incoming handle from the first, outgoing from the second */
                lastVertex.radius = Math.max(lastVertex.radius, vertices[i].radius);
                if (vertices[i].rightHandle) lastVertex.rightHandle = vertices[i].rightHandle;
                if (vertices[i].isCut) lastVertex.isCut = true;
                continue;
            }
            var uniqueVertex = makeVertex(vertices[i].x, vertices[i].y, vertices[i].radius, vertices[i].leftHandle, vertices[i].rightHandle);
            uniqueVertex.isCut = !!vertices[i].isCut;
            uniqueVertices.push(uniqueVertex);
        }
        if (isClosed && uniqueVertices.length > 1 && isSameVertex(uniqueVertices[0], uniqueVertices[uniqueVertices.length - 1])) {
            var closingVertex = uniqueVertices.pop();
            uniqueVertices[0].radius = Math.max(uniqueVertices[0].radius, closingVertex.radius);
            if (closingVertex.leftHandle) uniqueVertices[0].leftHandle = closingVertex.leftHandle;
            if (closingVertex.isCut) uniqueVertices[0].isCut = true;
        }
        return uniqueVertices;
    }

    /**
     * 円の頂点を4つ作る（上・右・下・左の順で、ハンドルで円にする）
     * @param {number} centerX - 中心の x
     * @param {number} centerY - 中心の y
     * @param {number} radius - 半径
     * @returns {OutlineVertex[]} 頂点（閉じたパス用）
     */
    function makeCircleVertices(centerX, centerY, radius) {
        var handleLength = radius * CORNER_HANDLE_RATIO;
        return [
            makeVertex(centerX, centerY + radius, 0, [centerX - handleLength, centerY + radius], [centerX + handleLength, centerY + radius]),
            makeVertex(centerX + radius, centerY, 0, [centerX + radius, centerY + handleLength], [centerX + radius, centerY - handleLength]),
            makeVertex(centerX, centerY - radius, 0, [centerX + handleLength, centerY - radius], [centerX - handleLength, centerY - radius]),
            makeVertex(centerX - radius, centerY, 0, [centerX - radius, centerY - handleLength], [centerX - radius, centerY + handleLength])
        ];
    }

    /**
     * 円弧をベジェにしたときの、始点から出るハンドルと終点に入るハンドルを返す（90°以下の円弧）
     * @param {number} centerX - 中心の x
     * @param {number} centerY - 中心の y
     * @param {number} radius - 半径
     * @param {number} startAngle - 始点の角度（ラジアン）
     * @param {number} endAngle - 終点の角度（ラジアン）
     * @returns {{start: number[], end: number[]}} ハンドルの位置
     */
    function computeArcHandles(centerX, centerY, radius, startAngle, endAngle) {
        var handleLength = radius * 4 / 3 * Math.tan(Math.abs(endAngle - startAngle) / 4);
        var turn = (endAngle > startAngle) ? 1 : -1;
        return {
            start: [
                centerX + radius * Math.cos(startAngle) - turn * handleLength * Math.sin(startAngle),
                centerY + radius * Math.sin(startAngle) + turn * handleLength * Math.cos(startAngle)
            ],
            end: [
                centerX + radius * Math.cos(endAngle) + turn * handleLength * Math.sin(endAngle),
                centerY + radius * Math.sin(endAngle) - turn * handleLength * Math.cos(endAngle)
            ]
        };
    }

    /**
     * しっぽの頂点を求める。線は y = lineY の水平線で、しっぽは y を減らす向き（外側）へ出す
     * @param {number} spanLeft - しっぽを置ける範囲の左端
     * @param {number} spanRight - しっぽを置ける範囲の右端
     * @param {number} lineY - 線の y
     * @param {BubbleSettings} bubbleSettings - 設定
     * @returns {Object} isBroken（線が切れるか）/ bottomVertices（切れないときの線の途中の頂点）/
     *     leftVertices・rightVertices（切れるときの左右の端の頂点）/ slashVertices（線と離した斜線）/ gapLeftX・gapRightX（切れ目の両端の x）/
     *     rightEdge・leftEdge（しっぽの右と左の辺 [根元, 先端]）/ tipPoint（●を置ける先端。無ければ null）/
     *     extraPolylines（線と別に描く頂点列。思考の円・つの）/ shadowTipY（影を止める y。省くと先端の y）
     */
    function computeTailVertices(spanLeft, spanRight, lineY, bubbleSettings) {
        var spanLength = Math.max(spanRight - spanLeft, 0);
        var halfGap = Math.max(0, Math.min(bubbleSettings.gapWidth, spanLength) / 2);
        /* しっぽが線からはみ出さないよう、内側へ寄せる / keep the tail within the line */
        var anchorX = clampNumber(spanLeft + spanLength * bubbleSettings.position / 100, spanLeft + halfGap, spanRight - halfGap);
        var gapLeft = makeVertex(anchorX - halfGap, lineY);
        var gapRight = makeVertex(anchorX + halfGap, lineY);
        var tipY = lineY - bubbleSettings.tailHeight;
        var direction = 0;
        if (bubbleSettings.directionKey === "right") direction = 1;
        if (bubbleSettings.directionKey === "left") direction = -1;
        var tailRun = bubbleSettings.tailWidth;
        var shapeKey = bubbleSettings.shapeKey;

        if (shapeKey !== "slash") {
            /* 右下は付け根の左端から、左下は右端から横に tailRun 離れた所が先端 / the tip sits tailRun from the far side of the base */
            var tipX = anchorX;
            if (direction > 0) tipX = gapLeft.x + tailRun;
            if (direction < 0) tipX = gapRight.x - tailRun;
            if (shapeKey === "round") return computeRoundTailVertices(gapLeft, gapRight, tipX, tipY);
            if (shapeKey === "loop") return computeLoopTailVertices(gapLeft, gapRight, tipX, tipY);
            if (shapeKey === "horn") return computeHornTailVertices(gapLeft, gapRight, tipX, tipY, bubbleSettings.strokeWidth);
            if (shapeKey === "thought") return computeThoughtTailVertices(gapLeft, gapRight, tipX, tipY);
            var notchTip = makeVertex(tipX, tipY);
            if (shapeKey === "curly") {
                /* カーリーブラケット（波括弧）の中央と同じく、付け根では線に沿って、先端では線と直角に入る曲線でとがらせる
                   Like the middle of a curly bracket: curves leave along the line and meet at the tip at right angles to it */
                var tipHandleY = tipY + (lineY - tipY) * CORNER_HANDLE_RATIO;
                gapLeft.rightHandle = [gapLeft.x + (tipX - gapLeft.x) * CORNER_HANDLE_RATIO, lineY];
                gapRight.leftHandle = [gapRight.x + (tipX - gapRight.x) * CORNER_HANDLE_RATIO, lineY];
                notchTip.leftHandle = [tipX, tipHandleY];
                notchTip.rightHandle = [tipX, tipHandleY];
            }
            return {
                isBroken: false,
                gapLeftX: gapLeft.x,
                gapRightX: gapRight.x,
                bottomVertices: [gapLeft, notchTip, gapRight],
                slashVertices: null,
                rightEdge: [gapRight, notchTip],
                leftEdge: [gapLeft, notchTip],
                tipPoint: notchTip
            };
        }

        /* 斜線の横の長さは、高さと角度から決める / the slash's run follows its height and angle */
        var slashAngle = clampNumber(bubbleSettings.slashAngle, MIN_SLASH_ANGLE, 90);
        tailRun = (slashAngle >= 90) ? 0 : bubbleSettings.tailHeight / Math.tan(slashAngle * Math.PI / 180);
        var hasGap = halfGap > TAIL_TOLERANCE;
        var isJoined = hasGap && direction !== 0 && bubbleSettings.rootKey === "edge";
        var rootX = anchorX;
        if (isJoined) rootX = (direction > 0) ? gapLeft.x : gapRight.x;
        /* 線と離した斜線だけ、左右・上下の調整を足す（つなげた斜線は線の端から動かさない）。上下はプラスで外側（y を減らす）
           Only a detached slash takes the nudge; a joined one stays on the line's end. Positive Y moves outward (decreasing y) */
        var shiftX = isJoined ? 0 : bubbleSettings.slashShiftX;
        var shiftY = isJoined ? 0 : -bubbleSettings.slashShiftY;
        var slashRoot = makeVertex(rootX + shiftX, lineY + shiftY);
        var slashTip = makeVertex(rootX + shiftX + direction * tailRun, tipY + shiftY);
        var slashEdge = [slashRoot, slashTip];
        if (!hasGap) {
            return { isBroken: false, gapLeftX: anchorX, gapRightX: anchorX, bottomVertices: [], slashVertices: slashEdge, rightEdge: slashEdge, leftEdge: slashEdge, tipPoint: slashTip };
        }
        return {
            isBroken: true,
            gapLeftX: gapLeft.x,
            gapRightX: gapRight.x,
            leftVertices: (isJoined && direction > 0) ? [gapLeft, slashTip] : [gapLeft],
            rightVertices: (isJoined && direction < 0) ? [slashTip, gapRight] : [gapRight],
            slashVertices: isJoined ? null : slashEdge,
            rightEdge: slashEdge,
            leftEdge: slashEdge,
            tipPoint: slashTip
        };
    }

    /**
     * U字のしっぽ：付け根から線と直角におりて、先端で線と平行になる曲線
     * @param {OutlineVertex} gapLeft - 付け根の左端
     * @param {OutlineVertex} gapRight - 付け根の右端
     * @param {number} tipX - 底の x
     * @param {number} tipY - 底の y
     * @returns {Object} computeTailVertices() と同じ形
     */
    function computeRoundTailVertices(gapLeft, gapRight, tipX, tipY) {
        var lineY = gapLeft.y;
        var sideHandleY = lineY - (lineY - tipY) * CORNER_HANDLE_RATIO;
        var roundBottom = makeVertex(tipX, tipY, 0,
            [tipX - (tipX - gapLeft.x) * CORNER_HANDLE_RATIO, tipY],
            [tipX + (gapRight.x - tipX) * CORNER_HANDLE_RATIO, tipY]);
        gapLeft.rightHandle = [gapLeft.x, sideHandleY];
        gapRight.leftHandle = [gapRight.x, sideHandleY];
        return {
            isBroken: false,
            gapLeftX: gapLeft.x,
            gapRightX: gapRight.x,
            bottomVertices: [gapLeft, roundBottom, gapRight],
            slashVertices: null,
            rightEdge: [gapRight, roundBottom],
            leftEdge: [gapLeft, roundBottom],
            tipPoint: roundBottom
        };
    }

    /**
     * ループのしっぽ：付け根から線に沿って出た曲線が交差点でおり、先端で輪を描いて同じ交差点から戻る
     * @param {OutlineVertex} gapLeft - 付け根の左端
     * @param {OutlineVertex} gapRight - 付け根の右端
     * @param {number} tipX - 輪の中心の x
     * @param {number} tipY - 輪の下端の y
     * @returns {Object} computeTailVertices() と同じ形
     */
    function computeLoopTailVertices(gapLeft, gapRight, tipX, tipY) {
        var lineY = gapLeft.y;
        var loopRadius = (lineY - tipY) * LOOP_RADIUS_RATIO;
        var loopCenterY = tipY + loopRadius;
        var crossY = loopCenterY + loopRadius * LOOP_CROSS_RATIO;
        var handleLength = loopRadius * CORNER_HANDLE_RATIO;
        /* 交差点は斜め45°で通る（入りは右下へ、出は右上へ） / the path crosses at 45° (down-right in, up-right out) */
        var crossHandle = loopRadius * 0.5;
        gapLeft.rightHandle = [gapLeft.x + (tipX - gapLeft.x) * CORNER_HANDLE_RATIO, lineY];
        gapRight.leftHandle = [gapRight.x + (tipX - gapRight.x) * CORNER_HANDLE_RATIO, lineY];
        var crossIn = makeVertex(tipX, crossY, 0, [tipX - crossHandle, crossY + crossHandle], [tipX + crossHandle, crossY - crossHandle]);
        var loopRight = makeVertex(tipX + loopRadius, loopCenterY, 0, [tipX + loopRadius, loopCenterY + handleLength], [tipX + loopRadius, loopCenterY - handleLength]);
        var loopBottom = makeVertex(tipX, tipY, 0, [tipX + handleLength, tipY], [tipX - handleLength, tipY]);
        var loopLeft = makeVertex(tipX - loopRadius, loopCenterY, 0, [tipX - loopRadius, loopCenterY - handleLength], [tipX - loopRadius, loopCenterY + handleLength]);
        var crossOut = makeVertex(tipX, crossY, 0, [tipX - crossHandle, crossY - crossHandle], [tipX + crossHandle, crossY + crossHandle]);
        return {
            isBroken: false,
            gapLeftX: gapLeft.x,
            gapRightX: gapRight.x,
            bottomVertices: [gapLeft, crossIn, loopRight, loopBottom, loopLeft, crossOut, gapRight],
            slashVertices: null,
            rightEdge: [gapRight, loopBottom],
            leftEdge: [gapLeft, loopBottom],
            tipPoint: null
        };
    }

    /**
     * つののしっぽ：線の内側の縁を根元に、真下へ出てから先端へ曲がる、塗りの形
     * @param {OutlineVertex} gapLeft - 根元の左端
     * @param {OutlineVertex} gapRight - 根元の右端
     * @param {number} tipX - 先端の x
     * @param {number} tipY - 先端の y
     * @param {number} strokeWidth - 線幅（根元を線の内側の縁まで上げる）
     * @returns {Object} computeTailVertices() と同じ形
     */
    function computeHornTailVertices(gapLeft, gapRight, tipX, tipY, strokeWidth) {
        var lineY = gapLeft.y;
        var rootY = lineY + strokeWidth / 2;
        var hornHeight = rootY - tipY;
        var rootHandleY = rootY - hornHeight * HORN_ROOT_HANDLE_RATIO;
        var anchorX = (gapLeft.x + gapRight.x) / 2;
        var tipHandle = [tipX + (anchorX - tipX) * HORN_TIP_HANDLE_RATIO, tipY + hornHeight * HORN_TIP_HANDLE_RATIO];
        var rootLeft = makeVertex(gapLeft.x, rootY, 0, null, [gapLeft.x, rootHandleY]);
        var hornTip = makeVertex(tipX, tipY, 0, tipHandle, tipHandle);
        var rootRight = makeVertex(gapRight.x, rootY, 0, [gapRight.x, rootHandleY], null);
        return {
            isBroken: false,
            gapLeftX: gapLeft.x,
            gapRightX: gapRight.x,
            bottomVertices: [],
            slashVertices: null,
            rightEdge: [gapRight, hornTip],
            leftEdge: [gapLeft, hornTip],
            tipPoint: null,
            extraPolylines: [{ vertices: [rootLeft, hornTip, rootRight], isClosed: true, isFilled: true }]
        };
    }

    /**
     * 思考のしっぽ：線に、付け根を丸めた半円のふくらみを付け、その先に円を置く
     * @param {OutlineVertex} gapLeft - ふくらみの左端
     * @param {OutlineVertex} gapRight - ふくらみの右端
     * @param {number} tipX - 円の中心の x
     * @param {number} tipY - 円の下端の y
     * @returns {Object} computeTailVertices() と同じ形
     */
    function computeThoughtTailVertices(gapLeft, gapRight, tipX, tipY) {
        var lineY = gapLeft.y;
        var anchorX = (gapLeft.x + gapRight.x) / 2;
        var halfGap = (gapRight.x - gapLeft.x) / 2;
        /* ふくらみ（半径 R、中心は線上）と付け根の丸み（半径 f）が接するよう、半分の幅 = √(R² + 2Rf) から R を決める
           The bump (radius R, centered on the line) and its fillets (radius f) meet tangentially: half the gap = √(R² + 2Rf) */
        var bumpRadius = halfGap / Math.sqrt(1 + 2 * THOUGHT_FILLET_RATIO);
        var circleRadius = bumpRadius * THOUGHT_CIRCLE_RATIO;
        var tailResult = {
            isBroken: false,
            gapLeftX: gapLeft.x,
            gapRightX: gapRight.x,
            bottomVertices: [],
            slashVertices: null,
            rightEdge: [makeVertex(anchorX + bumpRadius, lineY), makeVertex(anchorX + bumpRadius, lineY - bumpRadius)],
            leftEdge: [makeVertex(anchorX - bumpRadius, lineY), makeVertex(anchorX - bumpRadius, lineY - bumpRadius)],
            tipPoint: null,
            extraPolylines: [],
            shadowTipY: lineY - bumpRadius
        };
        if (bumpRadius <= TAIL_TOLERANCE) return tailResult;

        var filletRadius = bumpRadius * THOUGHT_FILLET_RATIO;
        /* 付け根の丸みの中心から、ふくらみの中心を見た角度 / angle from a fillet's center toward the bump's center */
        var filletAngle = Math.atan2(filletRadius, halfGap);
        var leftFillet = computeArcHandles(anchorX - halfGap, lineY - filletRadius, filletRadius, Math.PI / 2, filletAngle);
        var leftBump = computeArcHandles(anchorX, lineY, bumpRadius, filletAngle - Math.PI, -Math.PI / 2);
        var rightBump = computeArcHandles(anchorX, lineY, bumpRadius, -Math.PI / 2, -filletAngle);
        var rightFillet = computeArcHandles(anchorX + halfGap, lineY - filletRadius, filletRadius, Math.PI - filletAngle, Math.PI / 2);
        gapLeft.rightHandle = leftFillet.start;
        gapRight.leftHandle = rightFillet.end;
        var touchRatio = bumpRadius / (bumpRadius + filletRadius);
        var touchLeft = makeVertex(anchorX - halfGap * touchRatio, lineY - filletRadius * touchRatio, 0, leftFillet.end, leftBump.start);
        var bumpBottom = makeVertex(anchorX, lineY - bumpRadius, 0, leftBump.end, rightBump.start);
        var touchRight = makeVertex(anchorX + halfGap * touchRatio, lineY - filletRadius * touchRatio, 0, rightBump.end, rightFillet.start);
        tailResult.bottomVertices = [gapLeft, touchLeft, bumpBottom, touchRight, gapRight];
        tailResult.extraPolylines.push({ vertices: makeCircleVertices(tipX, tipY + circleRadius, circleRadius), isClosed: true });
        return tailResult;
    }

    /**
     * 辺の直線が y = targetY と交わる点の x を返す
     * @param {OutlineVertex} startVertex - 辺の始点
     * @param {OutlineVertex} endVertex - 辺の終点
     * @param {number} targetY - 求める y
     * @returns {number} 交点の x
     */
    function getXAtY(startVertex, endVertex, targetY) {
        var edgeRatio = (targetY - startVertex.y) / (endVertex.y - startVertex.y);
        return startVertex.x + (endVertex.x - startVertex.x) * edgeRatio;
    }

    /**
     * 囲みの影（右と下に重ねる線）の頂点列を返す。しっぽの右の辺は平行にずらし、先端の高さで止める
     * @param {Object} boxEdges - left / right / top / bottom / radius / bottomLeftRadius / bottomRightRadius（囲みの位置と角丸の半径。しっぽのかかった下の角は 0）
     * @param {Object} tailVertices - computeTailVertices() の戻り値
     * @param {number} tipY - しっぽの先端の y
     * @param {number} shadowDistance - 影の間隔
     * @returns {OutlineVertex[][]} 影の頂点列
     */
    function computeShadowPolylines(boxEdges, tailVertices, tipY, shadowDistance) {
        var shadowY = boxEdges.bottom - shadowDistance;
        var shadowRight = boxEdges.right + shadowDistance;
        /* 影の両端は、囲みの角丸が始まるアンカーポイントの位置でそろえる（角が小さいときは影の間隔ぶん内側）
           Both shadow ends line up with where the box's corner rounding starts (inset by the shadow distance for small corners) */
        var shadowEndX = boxEdges.left + Math.max(boxEdges.bottomLeftRadius, shadowDistance);
        var mainVertices = [
            makeVertex(shadowRight, boxEdges.top - Math.max(boxEdges.radius, shadowDistance)),
            makeVertex(shadowRight, shadowY, boxEdges.bottomRightRadius > 0 ? boxEdges.bottomRightRadius + shadowDistance : 0)
        ];
        /* しっぽが影より短いときは、下の線をそのまま通す / a tail shorter than the shadow distance does not break it */
        if (tipY >= shadowY - TAIL_TOLERANCE) {
            mainVertices.push(makeVertex(shadowEndX, shadowY));
            return [mainVertices];
        }

        /* しっぽの右の辺を、右（真横なら下）へ shadowDistance だけ平行にずらす / offset the tail's right edge outward */
        var edgeStart = tailVertices.rightEdge[0];
        var edgeEnd = tailVertices.rightEdge[1];
        var edgeLength = Math.sqrt(Math.pow(edgeEnd.x - edgeStart.x, 2) + Math.pow(edgeEnd.y - edgeStart.y, 2));
        var normalX = (edgeEnd.y - edgeStart.y) / edgeLength;
        var normalY = -(edgeEnd.x - edgeStart.x) / edgeLength;
        if (normalX < 0 || (normalX === 0 && normalY > 0)) {
            normalX = -normalX;
            normalY = -normalY;
        }
        var offsetStart = makeVertex(edgeStart.x + normalX * shadowDistance, edgeStart.y + normalY * shadowDistance);
        var offsetEnd = makeVertex(edgeEnd.x + normalX * shadowDistance, edgeEnd.y + normalY * shadowDistance);
        mainVertices.push(makeVertex(getXAtY(offsetStart, offsetEnd, shadowY), shadowY));
        mainVertices.push(makeVertex(getXAtY(offsetStart, offsetEnd, tipY), tipY));
        var shadowPolylines = [mainVertices];

        /* しっぽの左の辺の手前で、間隔をあけて下の線を止める / stop the bottom line short of the tail's left edge */
        var leftStopX = getXAtY(tailVertices.leftEdge[0], tailVertices.leftEdge[1], shadowY) - shadowDistance;
        if (leftStopX - shadowEndX > TAIL_TOLERANCE) {
            shadowPolylines.push([makeVertex(leftStopX, shadowY), makeVertex(shadowEndX, shadowY)]);
        }
        return shadowPolylines;
    }

    /**
     * 線としっぽの頂点列を求める（座標は、線が字面の下に来る向きにそろえた「ローカル座標」。y は上が正）
     * @param {Object} inkEdges - 字面の範囲 left / right / top / bottom（ローカル座標）
     * @param {BubbleSettings} bubbleSettings - 設定
     * @returns {{polylines: Array, boxEdges: Object|null}} polylines は { vertices, isClosed, isFilled } の配列、boxEdges は囲みの範囲（囲みのとき）
     */
    function computeBubblePolylines(inkEdges, bubbleSettings) {
        var frameLeft = inkEdges.left - bubbleSettings.extension;
        var frameRight = inkEdges.right + bubbleSettings.extension;
        var frameBottom = inkEdges.bottom - bubbleSettings.offset;
        var frameTop = inkEdges.top + bubbleSettings.offset;
        var frameKey = bubbleSettings.frameKey;
        if (frameRight - frameLeft <= TAIL_TOLERANCE) return { polylines: [], boxEdges: null };

        /* 角丸の半径は辺の長さで頭打ちにする / cap the corner radius by the side lengths */
        var cornerRadius = 0;
        var hasTicks = (frameKey === "bracket" || frameKey === "lShape");
        /* カプセル形状は、半径を大きく取ってから辺の長さで頭打ちにする（囲みは短い辺の半分、コの字・L字は立ち上がり） / a capsule takes the largest radius the sides allow */
        var requestedRadius = bubbleSettings.isCapsule ? Infinity : bubbleSettings.cornerRadius;
        if (hasTicks || frameKey === "box") cornerRadius = Math.max(0, Math.min(requestedRadius, (frameRight - frameLeft) / 2));
        if (frameKey === "box") cornerRadius = Math.min(cornerRadius, Math.max(0, (frameTop - frameBottom) / 2));
        if (hasTicks) cornerRadius = Math.min(cornerRadius, Math.max(0, bubbleSettings.tickHeight));
        var tickTop = frameBottom + bubbleSettings.tickHeight;
        /* L字は、しっぽに近いほうの端だけ立ち上げる / an L shape turns up only the end nearer the tail */
        var hasLeftTick = (frameKey === "bracket") || (frameKey === "lShape" && bubbleSettings.position <= 50);
        var hasRightTick = (frameKey === "bracket") || (frameKey === "lShape" && bubbleSettings.position > 50);

        /* しっぽは線の端まで置ける。しっぽが下の角の丸みにかかったら、その角だけ丸めずに角から出す。
           カプセル形状は形を保つため、これまでどおり角丸の内側に寄せる
           The tail may reach the ends of the line; a tail overlapping a bottom corner's rounding leaves that corner sharp.
           A capsule keeps its shape, so its tail stays inside the rounding */
        var tailInset = bubbleSettings.isCapsule ? cornerRadius : 0;
        var tailVertices = computeTailVertices(frameLeft + tailInset, frameRight - tailInset, frameBottom, bubbleSettings);
        var bottomLeftRadius = (tailVertices.gapLeftX < frameLeft + cornerRadius - TAIL_TOLERANCE) ? 0 : cornerRadius;
        var bottomRightRadius = (tailVertices.gapRightX > frameRight - cornerRadius + TAIL_TOLERANCE) ? 0 : cornerRadius;

        /* 下の辺の前後の頂点。囲みは一周してつながる / vertices before and after the bottom side; a box loops around */
        var beforeBottom = [makeVertex(frameLeft, frameBottom)];
        var afterBottom = [makeVertex(frameRight, frameBottom)];
        if (hasLeftTick) beforeBottom = [makeVertex(frameLeft, tickTop), makeVertex(frameLeft, frameBottom, bottomLeftRadius)];
        if (hasRightTick) afterBottom = [makeVertex(frameRight, frameBottom, bottomRightRadius), makeVertex(frameRight, tickTop)];
        if (frameKey === "box") {
            beforeBottom = [makeVertex(frameLeft, frameBottom, bottomLeftRadius)];
            afterBottom = [
                makeVertex(frameRight, frameBottom, bottomRightRadius),
                makeVertex(frameRight, frameTop, cornerRadius),
                makeVertex(frameLeft, frameTop, cornerRadius)
            ];
            if (bubbleSettings.sideGapWidth > TAIL_TOLERANCE) {
                /* 左右の辺の中央に切れ目の印を入れる（直線上の点なので形は変わらない） / mark the middle of the sides for the gaps; the points sit on the sides, so the shape is unchanged */
                var sideMiddleY = (frameTop + frameBottom) / 2;
                var rightMiddle = makeVertex(frameRight, sideMiddleY);
                var leftMiddle = makeVertex(frameLeft, sideMiddleY);
                rightMiddle.isCut = true;
                leftMiddle.isCut = true;
                afterBottom.splice(1, 0, rightMiddle);
                afterBottom.push(leftMiddle);
            }
        }

        var isBox = (frameKey === "box");
        var polylines = [];
        if (!tailVertices.isBroken) {
            polylines.push({ vertices: beforeBottom.concat(tailVertices.bottomVertices, afterBottom), isClosed: isBox });
        } else if (isBox) {
            /* 切れ目の右から一周して切れ目の左まで / from the right of the gap around to its left */
            polylines.push({ vertices: tailVertices.rightVertices.concat(afterBottom, beforeBottom, tailVertices.leftVertices), isClosed: false });
        } else {
            polylines.push({ vertices: beforeBottom.concat(tailVertices.leftVertices), isClosed: false });
            polylines.push({ vertices: tailVertices.rightVertices.concat(afterBottom), isClosed: false });
        }
        if (frameKey === "twoLines") polylines.push({ vertices: [makeVertex(frameLeft, frameTop), makeVertex(frameRight, frameTop)], isClosed: false });
        if (tailVertices.slashVertices) polylines.push({ vertices: tailVertices.slashVertices, isClosed: false });
        var extraPolylines = tailVertices.extraPolylines || [];
        for (var k = 0; k < extraPolylines.length; k++) polylines.push(extraPolylines[k]);
        if (bubbleSettings.tipDot && tailVertices.tipPoint) {
            var tipDotRadius = bubbleSettings.strokeWidth * TIP_DOT_RATIO / 2;
            polylines.push({ vertices: makeCircleVertices(tailVertices.tipPoint.x, tailVertices.tipPoint.y, tipDotRadius), isClosed: true, isFilled: true });
        }

        var boxEdges = isBox ? {
            left: frameLeft, right: frameRight, top: frameTop, bottom: frameBottom,
            radius: cornerRadius, bottomLeftRadius: bottomLeftRadius, bottomRightRadius: bottomRightRadius
        } : null;
        if (isBox && bubbleSettings.shadowDistance > TAIL_TOLERANCE) {
            /* 思考は円ではなくふくらみの先で止める / a thought stops at its bump, not at the circle */
            var shadowTipY = (typeof tailVertices.shadowTipY === "number") ? tailVertices.shadowTipY : frameBottom - bubbleSettings.tailHeight;
            var shadowPolylines = computeShadowPolylines(boxEdges, tailVertices, shadowTipY, bubbleSettings.shadowDistance);
            for (var i = 0; i < shadowPolylines.length; i++) polylines.push({ vertices: shadowPolylines[i], isClosed: false });
        }

        /* 重なった頂点をまとめ、2点に満たないものは描かない / merge duplicate vertices and drop pieces with fewer than two */
        var drawablePolylines = [];
        for (var j = 0; j < polylines.length; j++) {
            var uniqueVertices = removeDuplicateVertices(polylines[j].vertices, polylines[j].isClosed);
            if (uniqueVertices.length >= 2) drawablePolylines.push({ vertices: uniqueVertices, isClosed: polylines[j].isClosed, isFilled: !!polylines[j].isFilled });
        }
        return { polylines: drawablePolylines, boxEdges: boxEdges };
    }

    /**
     * 頂点列を、角丸をベジェのハンドルにしたパスの点の列にする
     * @param {OutlineVertex[]} vertices - 頂点
     * @param {boolean} isClosed - 閉じたパスなら true
     * @returns {Array<{anchor: number[], left: number[], right: number[], isCut: boolean}>} パスの点（座標とハンドル。isCut は左右の切れ目の印）
     */
    function buildBezierPoints(vertices, isClosed) {
        var pathPoints = [];
        var vertexCount = vertices.length;

        /**
         * ハンドルのない点を足す
         * @param {number} x - x 座標
         * @param {number} y - y 座標
         * @param {boolean} [isCut] - 左右の切れ目の印なら true
         * @returns {void}
         */
        function pushCornerPoint(x, y, isCut) {
            pathPoints.push({ anchor: [x, y], left: [x, y], right: [x, y], isCut: !!isCut });
        }

        for (var i = 0; i < vertexCount; i++) {
            var vertex = vertices[i];
            if (vertex.leftHandle || vertex.rightHandle) {
                /* すぼまりの曲線はハンドルを決めてある / tapered curves carry their own handles */
                pathPoints.push({
                    anchor: [vertex.x, vertex.y],
                    left: vertex.leftHandle || [vertex.x, vertex.y],
                    right: vertex.rightHandle || [vertex.x, vertex.y]
                });
                continue;
            }
            var hasNeighbors = isClosed || (i > 0 && i < vertexCount - 1);
            if (vertex.radius <= TAIL_TOLERANCE || !hasNeighbors) {
                pushCornerPoint(vertex.x, vertex.y, vertex.isCut);
                continue;
            }
            var prevVertex = vertices[(i - 1 + vertexCount) % vertexCount];
            var nextVertex = vertices[(i + 1) % vertexCount];
            var prevLength = Math.sqrt(Math.pow(prevVertex.x - vertex.x, 2) + Math.pow(prevVertex.y - vertex.y, 2));
            var nextLength = Math.sqrt(Math.pow(nextVertex.x - vertex.x, 2) + Math.pow(nextVertex.y - vertex.y, 2));
            /* 隣も角丸なら辺の半分まで / share a side with a rounded neighbor */
            var roundRadius = Math.min(
                vertex.radius,
                prevLength * (prevVertex.radius > 0 ? 0.5 : 1),
                nextLength * (nextVertex.radius > 0 ? 0.5 : 1)
            );
            if (roundRadius <= TAIL_TOLERANCE) {
                pushCornerPoint(vertex.x, vertex.y);
                continue;
            }
            var inX = vertex.x + (prevVertex.x - vertex.x) / prevLength * roundRadius;
            var inY = vertex.y + (prevVertex.y - vertex.y) / prevLength * roundRadius;
            var outX = vertex.x + (nextVertex.x - vertex.x) / nextLength * roundRadius;
            var outY = vertex.y + (nextVertex.y - vertex.y) / nextLength * roundRadius;
            /* 角に向かってハンドルを伸ばす / handles point toward the corner */
            pathPoints.push({
                anchor: [inX, inY],
                left: [inX, inY],
                right: [inX + (vertex.x - inX) * CORNER_HANDLE_RATIO, inY + (vertex.y - inY) * CORNER_HANDLE_RATIO]
            });
            pathPoints.push({
                anchor: [outX, outY],
                left: [outX + (vertex.x - outX) * CORNER_HANDLE_RATIO, outY + (vertex.y - outY) * CORNER_HANDLE_RATIO],
                right: [outX, outY]
            });
        }
        return mergeSameAnchorPoints(pathPoints, isClosed);
    }

    /**
     * 続けて同じ位置に来たパスの点を1つにまとめる（コの字の立ち上がりが角丸の半径と同じときなど）。
     * 入りのハンドルは前の点、出のハンドルは後ろの点から取る
     * @param {Array<{anchor: number[], left: number[], right: number[]}>} pathPoints - パスの点
     * @param {boolean} isClosed - 閉じたパスなら true（最後と最初も比べる）
     * @returns {Array<{anchor: number[], left: number[], right: number[]}>} まとめたパスの点
     */
    function mergeSameAnchorPoints(pathPoints, isClosed) {
        /**
         * 2つの点が同じ位置か
         * @param {Object} pointA - パスの点A
         * @param {Object} pointB - パスの点B
         * @returns {boolean} 同じなら true
         */
        function isSameAnchor(pointA, pointB) {
            return Math.abs(pointA.anchor[0] - pointB.anchor[0]) <= TAIL_TOLERANCE && Math.abs(pointA.anchor[1] - pointB.anchor[1]) <= TAIL_TOLERANCE;
        }
        var mergedPoints = [];
        for (var i = 0; i < pathPoints.length; i++) {
            var lastPoint = mergedPoints[mergedPoints.length - 1];
            if (lastPoint && isSameAnchor(lastPoint, pathPoints[i])) {
                lastPoint.right = pathPoints[i].right;
                if (pathPoints[i].isCut) lastPoint.isCut = true;
                continue;
            }
            mergedPoints.push(pathPoints[i]);
        }
        if (isClosed && mergedPoints.length > 1 && isSameAnchor(mergedPoints[0], mergedPoints[mergedPoints.length - 1])) {
            var closingPoint = mergedPoints.pop();
            mergedPoints[0].left = closingPoint.left;
            if (closingPoint.isCut) mergedPoints[0].isCut = true;
        }
        return mergedPoints;
    }

    // =========================================
    // 左右の切れ目 / Side gaps
    // =========================================

    /**
     * 2点を t で内分する
     * @param {number[]} pointA - 点A [x, y]
     * @param {number[]} pointB - 点B [x, y]
     * @param {number} t - 割合（0〜1）
     * @returns {number[]} 内分点 [x, y]
     */
    function lerpPoint(pointA, pointB, t) {
        return [pointA[0] + (pointB[0] - pointA[0]) * t, pointA[1] + (pointB[1] - pointA[1]) * t];
    }

    /**
     * 3次ベジェの区間を t で2つに分ける
     * @param {Array<number[]>} segment - [始点, 始点のハンドル, 終点のハンドル, 終点]
     * @param {number} t - 分ける位置（0〜1）
     * @returns {Array<Array<number[]>>} 前と後ろの区間
     */
    function splitSegment(segment, t) {
        var p01 = lerpPoint(segment[0], segment[1], t);
        var p12 = lerpPoint(segment[1], segment[2], t);
        var p23 = lerpPoint(segment[2], segment[3], t);
        var p012 = lerpPoint(p01, p12, t);
        var p123 = lerpPoint(p12, p23, t);
        var splitPoint = lerpPoint(p012, p123, t);
        return [[segment[0], p01, p012, splitPoint], [splitPoint, p123, p23, segment[3]]];
    }

    /**
     * 区間の始点から、曲線に沿って測った長さの累計を、CUT_SAMPLE_COUNT 等分の点ごとに返す
     * @param {Array<number[]>} segment - [始点, 始点のハンドル, 終点のハンドル, 終点]
     * @returns {number[]} 長さの累計（先頭は 0、末尾は区間の長さ）
     */
    function measureSegmentLengths(segment) {
        var lengths = [0];
        var previousPoint = segment[0];
        for (var i = 1; i <= CUT_SAMPLE_COUNT; i++) {
            var t = i / CUT_SAMPLE_COUNT;
            var u = 1 - t;
            var samplePoint = [
                u * u * u * segment[0][0] + 3 * u * u * t * segment[1][0] + 3 * u * t * t * segment[2][0] + t * t * t * segment[3][0],
                u * u * u * segment[0][1] + 3 * u * u * t * segment[1][1] + 3 * u * t * t * segment[2][1] + t * t * t * segment[3][1]
            ];
            lengths.push(lengths[i - 1] + Math.sqrt(Math.pow(samplePoint[0] - previousPoint[0], 2) + Math.pow(samplePoint[1] - previousPoint[1], 2)));
            previousPoint = samplePoint;
        }
        return lengths;
    }

    /**
     * 区間の始点から曲線に沿って targetLength の位置の t を返す
     * @param {number[]} lengths - measureSegmentLengths() の戻り値
     * @param {number} targetLength - 始点からの長さ
     * @returns {number} t（0〜1）
     */
    function findSegmentT(lengths, targetLength) {
        for (var i = 1; i < lengths.length; i++) {
            if (lengths[i] < targetLength) continue;
            var stepLength = lengths[i] - lengths[i - 1];
            var stepRatio = (stepLength > 0) ? (targetLength - lengths[i - 1]) / stepLength : 0;
            return (i - 1 + stepRatio) / CUT_SAMPLE_COUNT;
        }
        return 1;
    }

    /**
     * 区間の並びの始まりと終わりを、曲線に沿って指定の長さだけ削る（短い区間は丸ごと落とす）
     * @param {Array<Array<number[]>>} segments - 区間の配列
     * @param {number} startTrim - 始まりから削る長さ
     * @param {number} endTrim - 終わりから削る長さ
     * @returns {Array<Array<number[]>>} 削った区間の配列
     */
    function trimSegments(segments, startTrim, endTrim) {
        var trimmedSegments = segments.slice(0);
        var lengths;
        while (startTrim > TAIL_TOLERANCE && trimmedSegments.length) {
            lengths = measureSegmentLengths(trimmedSegments[0]);
            if (lengths[lengths.length - 1] <= startTrim) {
                startTrim -= lengths[lengths.length - 1];
                trimmedSegments.shift();
                continue;
            }
            trimmedSegments[0] = splitSegment(trimmedSegments[0], findSegmentT(lengths, startTrim))[1];
            startTrim = 0;
        }
        while (endTrim > TAIL_TOLERANCE && trimmedSegments.length) {
            var lastIndex = trimmedSegments.length - 1;
            lengths = measureSegmentLengths(trimmedSegments[lastIndex]);
            var segmentLength = lengths[lengths.length - 1];
            if (segmentLength <= endTrim) {
                endTrim -= segmentLength;
                trimmedSegments.pop();
                continue;
            }
            trimmedSegments[lastIndex] = splitSegment(trimmedSegments[lastIndex], findSegmentT(lengths, segmentLength - endTrim))[0];
            endTrim = 0;
        }
        return trimmedSegments;
    }

    /**
     * パスの点を、切れ目の印のところで切り、印の前後を halfCutLength ずつ削った開いたパスに分ける
     * @param {Array<{anchor: number[], left: number[], right: number[], isCut: boolean}>} pathPoints - buildBezierPoints() の戻り値
     * @param {boolean} isClosed - 閉じたパスなら true
     * @param {number} halfCutLength - 切れ目の長さの半分
     * @returns {Array<{points: Array, isClosed: boolean}>} 分けたパス（印が無ければ元のパス1つ）
     */
    function splitPathAtCuts(pathPoints, isClosed, halfCutLength) {
        var firstCutIndex = -1;
        var i;
        for (i = 0; i < pathPoints.length; i++) {
            if (pathPoints[i].isCut) {
                firstCutIndex = i;
                break;
            }
        }
        if (firstCutIndex < 0 || halfCutLength <= TAIL_TOLERANCE) return [{ points: pathPoints, isClosed: isClosed }];

        /* 閉じたパスは、最初の印から一周して同じ印に戻る開いたパスにする / a closed path becomes an open one from the first mark around to itself */
        var openPoints = pathPoints;
        if (isClosed) {
            openPoints = [];
            for (i = 0; i <= pathPoints.length; i++) openPoints.push(pathPoints[(firstCutIndex + i) % pathPoints.length]);
        }
        var segments = [];
        for (i = 0; i < openPoints.length - 1; i++) {
            segments.push([openPoints[i].anchor, openPoints[i].right, openPoints[i + 1].left, openPoints[i + 1].anchor]);
        }

        var pathPieces = [];
        var pieceStart = 0;
        for (i = 1; i < openPoints.length; i++) {
            if (!openPoints[i].isCut && i < openPoints.length - 1) continue;
            var pieceSegments = trimSegments(segments.slice(pieceStart, i),
                openPoints[pieceStart].isCut ? halfCutLength : 0, openPoints[i].isCut ? halfCutLength : 0);
            if (pieceSegments.length) pathPieces.push({ points: segmentsToPathPoints(pieceSegments), isClosed: false });
            pieceStart = i;
        }
        return pathPieces;
    }

    /**
     * 区間の並びを、開いたパスの点の配列に戻す
     * @param {Array<Array<number[]>>} segments - 区間の配列（1つ以上）
     * @returns {Array<{anchor: number[], left: number[], right: number[]}>} パスの点
     */
    function segmentsToPathPoints(segments) {
        var pathPoints = [{ anchor: segments[0][0], left: segments[0][0], right: segments[0][1] }];
        for (var i = 0; i < segments.length; i++) {
            var nextSegment = segments[i + 1];
            pathPoints.push({ anchor: segments[i][3], left: segments[i][2], right: nextSegment ? nextSegment[1] : segments[i][3] });
        }
        return pathPoints;
    }

    // =========================================
    // 向きの変換 / Orientation
    // =========================================

    /**
     * ページの座標を、線を描く側が下に来るローカル座標に変える。
     * 上は上下に裏返し、左は90°回して、右は回して裏返して、どれも文字の始まり（左・上）がローカルの左になるようにする
     * @param {number[]} pagePoint - ページの [x, y]
     * @param {string} sideKey - 線を描く側。"top" / "bottom" / "left" / "right"
     * @returns {number[]} ローカルの [x, y]
     */
    function pageToLocal(pagePoint, sideKey) {
        if (sideKey === "top") return [pagePoint[0], -pagePoint[1]];
        if (sideKey === "left") return [-pagePoint[1], pagePoint[0]];
        if (sideKey === "right") return [-pagePoint[1], -pagePoint[0]];
        return [pagePoint[0], pagePoint[1]];
    }

    /**
     * ローカル座標をページの座標に戻す（pageToLocal() の逆）
     * @param {number[]} localPoint - ローカルの [x, y]
     * @param {string} sideKey - 線を描く側。"top" / "bottom" / "left" / "right"
     * @returns {number[]} ページの [x, y]
     */
    function localToPage(localPoint, sideKey) {
        if (sideKey === "top") return [localPoint[0], -localPoint[1]];
        if (sideKey === "left") return [localPoint[1], -localPoint[0]];
        if (sideKey === "right") return [-localPoint[1], -localPoint[0]];
        return [localPoint[0], localPoint[1]];
    }

    /**
     * ページの範囲 [左, 上, 右, 下] を、ローカル座標の範囲にする
     * @param {number[]} pageBounds - [左, 上, 右, 下]
     * @param {string} sideKey - 線を描く側
     * @returns {{left: number, right: number, top: number, bottom: number}} ローカル座標の範囲
     */
    function boundsToLocalEdges(pageBounds, sideKey) {
        var cornerA = pageToLocal([pageBounds[0], pageBounds[1]], sideKey);
        var cornerB = pageToLocal([pageBounds[2], pageBounds[3]], sideKey);
        return {
            left: Math.min(cornerA[0], cornerB[0]),
            right: Math.max(cornerA[0], cornerB[0]),
            top: Math.max(cornerA[1], cornerB[1]),
            bottom: Math.min(cornerA[1], cornerB[1])
        };
    }

    /**
     * ローカル座標の範囲を、ページの範囲 [左, 上, 右, 下] に戻す
     * @param {{left: number, right: number, top: number, bottom: number}} localEdges - ローカル座標の範囲
     * @param {string} sideKey - 線を描く側
     * @returns {number[]} [左, 上, 右, 下]
     */
    function localEdgesToBounds(localEdges, sideKey) {
        var cornerA = localToPage([localEdges.left, localEdges.top], sideKey);
        var cornerB = localToPage([localEdges.right, localEdges.bottom], sideKey);
        return [
            Math.min(cornerA[0], cornerB[0]),
            Math.max(cornerA[1], cornerB[1]),
            Math.max(cornerA[0], cornerB[0]),
            Math.min(cornerA[1], cornerB[1])
        ];
    }

    // =========================================
    // テキストの測定 / Text measurement
    // =========================================

    /**
     * @typedef {Object} BubbleTarget
     * @property {TextFrame} frame - 対象のテキストフレーム
     * @property {number[]} inkBounds - 字面の範囲 [左, 上, 右, 下]（pt）
     * @property {boolean} isVertical - 縦組みなら true
     * @property {number} fontSize - 1文字目の文字サイズ（pt）
     * @property {Color} strokeColor - 線のカラー（1文字目の塗り）
     * @property {number[]|null} areaFrameBounds - エリア内文字の枠の範囲 [左, 上, 右, 下]（エリア内文字でなければ null）
     */

    /**
     * 字面（アウトラインの形）の範囲を測る。複製をアウトライン化して測り、すぐに消す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {number[]|null} [左, 上, 右, 下]（pt）。測れなければ null
     */
    function measureInkBounds(textFrame) {
        var duplicateFrame = null;
        var outlineGroup = null;
        var inkBounds = null;
        try {
            duplicateFrame = textFrame.duplicate();
            /* createOutline() は複製を消費するので、片付けるのは戻り値のグループ / createOutline() consumes the copy, so remove the returned group */
            outlineGroup = duplicateFrame.createOutline();
            var outlineBounds = outlineGroup.geometricBounds;
            inkBounds = [outlineBounds[0], outlineBounds[1], outlineBounds[2], outlineBounds[3]];
        } catch (e) {
            inkBounds = null;
        }
        if (outlineGroup) {
            try { outlineGroup.remove(); } catch (removeError) { /* 消せなくても測った値は使える / the measurement is still usable */ }
        } else if (duplicateFrame) {
            try { duplicateFrame.remove(); } catch (removeDuplicateError) { /* アウトライン化の途中で消費された / already consumed */ }
        }
        return inkBounds;
    }

    /**
     * 1文字目の塗りを返す（塗りが無ければ黒）
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {Color} 線に使うカラー
     */
    function readTextColor(textFrame) {
        /* 対象は空白だけでないテキストなので、1文字目はかならずある / targets are never blank, so the first character exists */
        var fillColor = textFrame.textRange.characters[0].characterAttributes.fillColor;
        if (fillColor.typename !== "NoColor") return fillColor;
        var blackColor = new GrayColor();
        blackColor.gray = 100;
        return blackColor;
    }

    /**
     * 1文字目の文字サイズを返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {number} 文字サイズ（pt）
     */
    function readFontSize(textFrame) {
        return textFrame.textRange.characters[0].characterAttributes.size;
    }

    /**
     * エリア内文字の枠の範囲を返す（エリア内文字の geometricBounds は枠のパスの範囲）
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {number[]|null} [左, 上, 右, 下]。エリア内文字でなければ null
     */
    function readAreaFrameBounds(textFrame) {
        if (textFrame.kind !== TextType.AREATEXT) return null;
        var frameBounds = textFrame.geometricBounds;
        return [frameBounds[0], frameBounds[1], frameBounds[2], frameBounds[3]];
    }

    /**
     * 縦組みのテキストか
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {boolean} 縦組みなら true
     */
    function isVerticalText(textFrame) {
        return textFrame.orientation === TextOrientation.VERTICAL;
    }

    /**
     * 選択から対象のテキストを集め、字面の範囲を測っておく（プレビューのたびに測り直さない）
     * @param {Document} doc - 対象のドキュメント
     * @returns {BubbleTarget[]} 対象（空白だけのテキストは除く）
     */
    function collectBubbleTargets(doc) {
        var textFrames = collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true });
        var bubbleTargets = [];
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!/\S/.test(textFrame.contents)) continue;
            var inkBounds = measureInkBounds(textFrame);
            if (!inkBounds) {
                /* 測れないときはフレームの範囲で代用する / fall back to the frame bounds */
                var frameBounds = textFrame.geometricBounds;
                inkBounds = [frameBounds[0], frameBounds[1], frameBounds[2], frameBounds[3]];
            }
            bubbleTargets.push({
                frame: textFrame,
                inkBounds: inkBounds,
                isVertical: isVerticalText(textFrame),
                fontSize: readFontSize(textFrame),
                strokeColor: readTextColor(textFrame),
                areaFrameBounds: readAreaFrameBounds(textFrame)
            });
        }
        return bubbleTargets;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * パスに線の設定を入れる（塗りなし。線端なしは角をとがらせ、丸型は角も丸める）
     * @param {PathItem} tailPath - 対象のパス
     * @param {Color} strokeColor - 線のカラー
     * @param {number} strokeWidth - 線幅（pt）
     * @param {string} capKey - 線端。"butt"（なし）/ "round"（丸型）
     * @returns {void}
     */
    function applyTailStroke(tailPath, strokeColor, strokeWidth, capKey) {
        var isRoundCap = (capKey === "round");
        tailPath.filled = false;
        tailPath.stroked = true;
        tailPath.strokeColor = strokeColor;
        tailPath.strokeWidth = strokeWidth;
        tailPath.strokeCap = isRoundCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
        tailPath.strokeJoin = isRoundCap ? StrokeJoin.ROUNDENDJOIN : StrokeJoin.MITERENDJOIN;
        tailPath.strokeMiterLimit = TAIL_MITER_LIMIT;
        tailPath.strokeDashes = [];
    }

    /**
     * パスに塗りの設定を入れる（線なし。つの・●用）
     * @param {PathItem} tailPath - 対象のパス
     * @param {Color} fillColor - 塗りのカラー
     * @returns {void}
     */
    function applyTailFill(tailPath, fillColor) {
        tailPath.stroked = false;
        tailPath.filled = true;
        tailPath.fillColor = fillColor;
    }

    /**
     * ローカル座標のパスの点からパスを作る（ページの座標に戻す）
     * @param {GroupItem} parentGroup - 追加先
     * @param {Array<{anchor: number[], left: number[], right: number[]}>} bezierPoints - パスの点（ローカル座標）
     * @param {boolean} isClosed - 閉じたパスなら true
     * @param {string} sideKey - 線を描く側
     * @returns {PathItem} 作ったパス
     */
    function addBezierPath(parentGroup, bezierPoints, isClosed, sideKey) {
        var anchors = [];
        var i;
        for (i = 0; i < bezierPoints.length; i++) anchors.push(localToPage(bezierPoints[i].anchor, sideKey));
        var polylinePath = parentGroup.pathItems.add();
        polylinePath.setEntirePath(anchors);
        polylinePath.closed = isClosed;
        for (i = 0; i < bezierPoints.length; i++) {
            var pathPoint = polylinePath.pathPoints[i];
            pathPoint.leftDirection = localToPage(bezierPoints[i].left, sideKey);
            pathPoint.rightDirection = localToPage(bezierPoints[i].right, sideKey);
            pathPoint.pointType = PointType.CORNER;
        }
        return polylinePath;
    }

    /**
     * テキストの組み方から線を描く側を決める（1つ目のテキストが縦組みなら右、横組みなら下）
     * @param {BubbleTarget[]} bubbleTargets - 対象
     * @returns {string} "right" / "bottom"
     */
    function getAutoSideKey(bubbleTargets) {
        return (bubbleTargets.length && bubbleTargets[0].isVertical) ? "right" : "bottom";
    }

    /**
     * 線を描く側を返す（"auto" なら横組みは下、縦組みは右）
     * @param {BubbleTarget} bubbleTarget - 対象
     * @param {BubbleSettings} bubbleSettings - 設定
     * @returns {string} "top" / "bottom" / "left" / "right"
     */
    function getLineSideKey(bubbleTarget, bubbleSettings) {
        var sideKey = bubbleSettings.sideKey;
        if (sideKey === "top" || sideKey === "bottom" || sideKey === "left" || sideKey === "right") return sideKey;
        return bubbleTarget.isVertical ? "right" : "bottom";
    }

    /**
     * 1つのテキストに、線としっぽをまとめたグループを描く（テキストのすぐ前面に置く）
     * @param {BubbleTarget} bubbleTarget - 対象
     * @param {BubbleSettings} bubbleSettings - 設定
     * @returns {{group: GroupItem, boxBounds: number[]|null}|null} 描いたグループと囲みの範囲 [左, 上, 右, 下]。描けなければ null
     */
    function drawBubble(bubbleTarget, bubbleSettings) {
        var textFrame = bubbleTarget.frame;
        var sideKey = getLineSideKey(bubbleTarget, bubbleSettings);
        /* エリア内文字は、設定に応じて枠の範囲を基準にする。枠に沿わせるので、間隔・左右の伸びは 0 にする
           Area type can be measured by its frame; the line then hugs the frame, so Distance and Extension are 0 */
        var shapeSettings = bubbleSettings;
        var baseBounds = bubbleTarget.inkBounds;
        if (bubbleSettings.useAreaFrame && bubbleTarget.areaFrameBounds) {
            baseBounds = bubbleTarget.areaFrameBounds;
            shapeSettings = copySettings(bubbleSettings);
            shapeSettings.offset = 0;
            shapeSettings.extension = 0;
        }
        var bubbleShape = computeBubblePolylines(boundsToLocalEdges(baseBounds, sideKey), shapeSettings);
        if (!bubbleShape.polylines.length) return null;

        var lineColor = (bubbleSettings.colorKey === "custom") ? colorDataToColor(bubbleSettings.customColor) : bubbleTarget.strokeColor;
        var tailGroup = textFrame.parent.groupItems.add();
        tailGroup.move(textFrame, ElementPlacement.PLACEBEFORE);
        tailGroup.name = getLabel(LABELS.fallbackName.tailGroup);
        for (var i = 0; i < bubbleShape.polylines.length; i++) {
            var polyline = bubbleShape.polylines[i];
            /* 角丸のハンドルを付けてから、左右の切れ目の印で切り分ける / add the corner handles, then cut at the side gap marks */
            var pathPieces = splitPathAtCuts(buildBezierPoints(polyline.vertices, polyline.isClosed), polyline.isClosed, bubbleSettings.sideGapWidth / 2);
            for (var j = 0; j < pathPieces.length; j++) {
                var piecePath = addBezierPath(tailGroup, pathPieces[j].points, pathPieces[j].isClosed, sideKey);
                if (polyline.isFilled) applyTailFill(piecePath, lineColor);
                else applyTailStroke(piecePath, lineColor, bubbleSettings.strokeWidth, bubbleSettings.capKey);
            }
        }
        return {
            group: tailGroup,
            boxBounds: bubbleShape.boxEdges ? localEdgesToBounds(bubbleShape.boxEdges, sideKey) : null
        };
    }

    /**
     * 設定を浅くコピーする
     * @param {BubbleSettings} bubbleSettings - 設定
     * @returns {BubbleSettings} コピー
     */
    function copySettings(bubbleSettings) {
        var settingsCopy = {};
        for (var settingKey in bubbleSettings) {
            if (bubbleSettings.hasOwnProperty(settingKey)) settingsCopy[settingKey] = bubbleSettings[settingKey];
        }
        return settingsCopy;
    }

    /**
     * 線のグループとテキストを、1つのグループにまとめる（線が前面）
     * @param {GroupItem} tailGroup - drawBubble() で描いたグループ
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {GroupItem} まとめたグループ
     */
    function groupTailWithText(tailGroup, textFrame) {
        var bubbleGroup = textFrame.parent.groupItems.add();
        bubbleGroup.move(textFrame, ElementPlacement.PLACEBEFORE);
        textFrame.move(bubbleGroup, ElementPlacement.PLACEATEND);
        tailGroup.move(bubbleGroup, ElementPlacement.PLACEATBEGINNING);
        bubbleGroup.name = getLabel(LABELS.fallbackName.bubbleGroup);
        return bubbleGroup;
    }

    /**
     * すべての対象に線としっぽを描く
     * @param {BubbleTarget[]} bubbleTargets - 対象
     * @param {BubbleSettings} bubbleSettings - 設定
     * @returns {{drawn: Array<{target: BubbleTarget, group: GroupItem, boxBounds: number[]|null}>, skippedCount: number}} 描いたものと、描けなかった数
     */
    function drawBubbles(bubbleTargets, bubbleSettings) {
        var drawnBubbles = [];
        var skippedCount = 0;
        for (var i = 0; i < bubbleTargets.length; i++) {
            try {
                var drawnBubble = drawBubble(bubbleTargets[i], bubbleSettings);
                if (!drawnBubble) {
                    skippedCount++;
                    continue;
                }
                drawnBubble.target = bubbleTargets[i];
                drawnBubbles.push(drawnBubble);
            } catch (e) {
                /* ロックされたグループの中などで描けない / e.g. inside a locked group */
                skippedCount++;
            }
        }
        return { drawn: drawnBubbles, skippedCount: skippedCount };
    }

    // =========================================
    // 線のカラー / Line color
    // =========================================

    /**
     * 保存用のカラーの値から Illustrator のカラーを作る
     * @param {{model: string, values: number[]}} colorData - "rgb"（0〜255）/ "cmyk"（0〜100）/ "gray"（0〜100）と値
     * @returns {Color} カラー（知らない形式なら黒）
     */
    function colorDataToColor(colorData) {
        var colorValues = colorData.values || [];
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
        var grayColor = new GrayColor();
        grayColor.gray = (colorData.model === "gray") ? colorValues[0] : 100;
        return grayColor;
    }

    /**
     * Illustrator のカラーを保存用の値にする（RGB・CMYK・グレー以外は null）
     * @param {Color} sourceColor - カラー
     * @returns {{model: string, values: number[]}|null} 保存用の値
     */
    function colorToColorData(sourceColor) {
        if (sourceColor.typename === "RGBColor") return { model: "rgb", values: [sourceColor.red, sourceColor.green, sourceColor.blue] };
        if (sourceColor.typename === "CMYKColor") return { model: "cmyk", values: [sourceColor.cyan, sourceColor.magenta, sourceColor.yellow, sourceColor.black] };
        if (sourceColor.typename === "GrayColor") return { model: "gray", values: [sourceColor.gray] };
        return null;
    }

    /**
     * カラーの値をドキュメントのカラーモード（RGB か CMYK）にそろえる。
     * カラーピッカーはグレーの色を渡すとグレースケールで開くので、開く前に通す（換算は簡易）
     * @param {{model: string, values: number[]}} colorData - 保存用のカラーの値
     * @returns {{model: string, values: number[]}} ドキュメントのカラーモードの値
     */
    function toDocumentColorData(colorData) {
        var isCmykDocument = (app.activeDocument.documentColorSpace === DocumentColorSpace.CMYK);
        var colorValues = colorData.values || [];
        if (isCmykDocument && colorData.model === "cmyk") return colorData;
        if (!isCmykDocument && colorData.model === "rgb") return colorData;
        if (isCmykDocument && colorData.model === "gray") return { model: "cmyk", values: [0, 0, 0, colorValues[0]] };
        var screenRgb = colorDataToScreenRgb(colorData);
        if (!isCmykDocument) return { model: "rgb", values: [Math.round(screenRgb[0] * 255), Math.round(screenRgb[1] * 255), Math.round(screenRgb[2] * 255)] };
        /* RGB から CMYK への簡易換算 / rough RGB to CMYK */
        var keyLevel = 1 - Math.max(screenRgb[0], screenRgb[1], screenRgb[2]);
        if (keyLevel >= 1) return { model: "cmyk", values: [0, 0, 0, 100] };
        return {
            model: "cmyk",
            values: [
                Math.round((1 - screenRgb[0] - keyLevel) / (1 - keyLevel) * 100),
                Math.round((1 - screenRgb[1] - keyLevel) / (1 - keyLevel) * 100),
                Math.round((1 - screenRgb[2] - keyLevel) / (1 - keyLevel) * 100),
                Math.round(keyLevel * 100)
            ]
        };
    }

    /**
     * 色見本に塗る RGB（0〜1）を返す。CMYK・グレーは簡易換算する
     * @param {{model: string, values: number[]}} colorData - 保存用のカラーの値
     * @returns {number[]} [r, g, b]（0〜1）
     */
    function colorDataToScreenRgb(colorData) {
        var colorValues = colorData.values || [];
        if (colorData.model === "rgb") return [colorValues[0] / 255, colorValues[1] / 255, colorValues[2] / 255];
        if (colorData.model === "cmyk") {
            var keyRatio = 1 - colorValues[3] / 100;
            return [(1 - colorValues[0] / 100) * keyRatio, (1 - colorValues[1] / 100) * keyRatio, (1 - colorValues[2] / 100) * keyRatio];
        }
        var grayLevel = 1 - ((colorData.model === "gray") ? colorValues[0] : 100) / 100;
        return [grayLevel, grayLevel, grayLevel];
    }

    /**
     * 描いたオブジェクトを消す
     * @param {PageItem[]} pageItems - 消すオブジェクト
     * @returns {void}
     */
    function removePageItems(pageItems) {
        for (var i = 0; i < pageItems.length; i++) {
            try { pageItems[i].remove(); } catch (e) { /* 消えている / already gone */ }
        }
    }

    // =========================================
    // エリア内文字 / Area type
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

    var AREA_TEXT_ACTION_SET_NAME = "SpeechBubbleMaker_AreaText"; /* アクションセット名 / action set name */

    /* エリア内文字のアクション（AreaTypeToolkit と同じイベント）。自動サイズ調整は 2=オフ、テキストの配置は 1=中央
       Area type actions (the same events as AreaTypeToolkit): auto size 2 = off, vertical alignment 1 = center */
    var AREA_TEXT_ACTIONS = {
        autoSizeOff: { name: "AutoSizeOff", internalName: "adobe_SLOAreaTextDialog", localizedName: "エリア内文字オプション", key: 1952539754, value: 2 },
        alignCenter: { name: "AlignCenter", internalName: "adobe_frameAlignment", localizedName: "エリア内文字のフレーム整列", key: 1717660782, value: 1 }
    };

    /**
     * AREA_TEXT_ACTIONS の2つのアクションを持つセットの定義を返す
     * @returns {string} アクション定義のテキスト
     */
    function buildAreaTextActionSet() {
        var actionKeys = ["autoSizeOff", "alignCenter"];
        var actionLines = ["/version 3"].concat(
            buildActionNameLines("", AREA_TEXT_ACTION_SET_NAME),
            ["/isOpen 1", "/actionCount " + actionKeys.length]
        );
        for (var i = 0; i < actionKeys.length; i++) {
            var actionDef = AREA_TEXT_ACTIONS[actionKeys[i]];
            actionLines = actionLines.concat(
                ["/action-" + (i + 1) + " {"],
                buildActionNameLines("\t", actionDef.name),
                [
                    "\t/keyIndex 0",
                    "\t/colorIndex 0",
                    "\t/isOpen 1",
                    "\t/eventCount 1",
                    "\t/event-1 {",
                    "\t\t/useRulersIn1stQuadrant 0",
                    "\t\t/internalName (" + actionDef.internalName + ")"
                ],
                buildActionNameLines("\t\t", actionDef.localizedName, "localizedName"),
                [
                    "\t\t/isOpen 0",
                    "\t\t/isOn 1",
                    "\t\t/hasDialog 0",
                    "\t\t/parameterCount 1",
                    "\t\t/parameter-1 {",
                    "\t\t\t/key " + actionDef.key,
                    "\t\t\t/showInPalette 4294967295",
                    "\t\t\t/type (integer)",
                    "\t\t\t/value " + actionDef.value,
                    "\t\t}",
                    "\t}",
                    "}"
                ]
            );
        }
        return actionLines.concat([""]).join("\n");
    }

    /**
     * ポイント文字をエリア内文字に変換する（変換で参照が置き換わるので、一時グループの中で変換して拾う）
     * @param {TextFrame} pointText - ポイント文字
     * @returns {TextFrame} 変換したエリア内文字（元のポイント文字は消す）
     */
    function convertPointTextToAreaText(pointText) {
        var holderGroup = pointText.parent.groupItems.add();
        holderGroup.move(pointText, ElementPlacement.PLACEBEFORE);
        pointText.duplicate(holderGroup, ElementPlacement.PLACEATEND).convertPointObjectToAreaObject();
        var areaText = holderGroup.textFrames[0];
        areaText.move(holderGroup, ElementPlacement.PLACEBEFORE);
        holderGroup.remove();
        pointText.remove();
        return areaText;
    }

    /**
     * エリア内文字の枠を、囲みの範囲と同じ大きさ・位置にする
     * @param {TextFrame} areaText - エリア内文字
     * @param {number[]} boxBounds - 囲みの範囲 [左, 上, 右, 下]
     * @returns {void}
     */
    function fitAreaTextToBox(areaText, boxBounds) {
        areaText.textPath.width = Math.max(boxBounds[2] - boxBounds[0], MIN_AREA_TEXT_SIZE);
        areaText.textPath.height = Math.max(boxBounds[1] - boxBounds[3], MIN_AREA_TEXT_SIZE);
        /* エリア内文字の geometricBounds は枠のパスの範囲 / The geometricBounds of area type is its frame path */
        var frameBounds = areaText.geometricBounds;
        areaText.translate(boxBounds[0] - frameBounds[0], boxBounds[1] - frameBounds[1]);
    }

    /**
     * 囲みの中のテキストを、囲みと同じ大きさのエリア内文字にして左右・天地中央に揃える（ダイアログを閉じてから呼ぶ）。
     * ポイント文字は変換し、エリア内文字はそのまま枠を合わせる。パス上文字はそのまま
     * @param {Array<{target: BubbleTarget, boxBounds: number[]|null}>} drawnBubbles - drawBubbles() の drawn
     * @returns {boolean} 天地中央のアクションを実行できたら true
     */
    function centerTextsInBoxes(drawnBubbles) {
        var fittedBubbles = [];
        var areaTexts = [];
        var i;
        for (i = 0; i < drawnBubbles.length; i++) {
            var drawnBubble = drawnBubbles[i];
            var textFrame = drawnBubble.target.frame;
            if (!drawnBubble.boxBounds || textFrame.kind === TextType.PATHTEXT) continue;
            if (textFrame.kind === TextType.POINTTEXT) {
                textFrame = convertPointTextToAreaText(textFrame);
                drawnBubble.target.frame = textFrame;
            }
            textFrame.spacing = 0;
            textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
            fitAreaTextToBox(textFrame, drawnBubble.boxBounds);
            fittedBubbles.push(drawnBubble);
            areaTexts.push(textFrame);
        }
        if (!areaTexts.length) return true;

        var doc = app.activeDocument;
        doc.selection = null;
        doc.selection = areaTexts;
        /* Illustrator に選択状態を確定させる / Let Illustrator commit the selection */
        app.redraw();
        if (!loadTemporaryActionSet(buildAreaTextActionSet(), AREA_TEXT_ACTION_SET_NAME)) return false;
        try {
            /* 変換したエリア内文字は自動サイズ調整がオンで枠が縮むので、切ってから合わせ直す / auto size shrinks the frame, so turn it off and refit */
            app.doScript(AREA_TEXT_ACTIONS.autoSizeOff.name, AREA_TEXT_ACTION_SET_NAME, false);
            for (i = 0; i < fittedBubbles.length; i++) fitAreaTextToBox(fittedBubbles[i].target.frame, fittedBubbles[i].boxBounds);
            app.doScript(AREA_TEXT_ACTIONS.alignCenter.name, AREA_TEXT_ACTION_SET_NAME, false);
            return true;
        } catch (e) {
            $.writeln("centerTextsInBoxes: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(AREA_TEXT_ACTION_SET_NAME);
        }
    }

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
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", "#...#", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
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

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
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
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
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

    // =========================================
    // 既定値とプリセットの保存 / Defaults and saved presets
    // =========================================

    /**
     * 初めて開いたときの設定を、1文字目の文字サイズから作る（寸法は単位の小数第1位で丸める）
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {Object} rulerUnit - 長さの単位（getUnitInfo() の戻り値）
     * @param {Object} strokeUnit - 線幅の単位（getUnitInfo("strokeUnits") の戻り値）
     * @returns {BubbleSettings} 既定の設定
     */
    function buildDefaultSettings(fontSize, rulerUnit, strokeUnit) {
        var defaultSettings = {};
        for (var settingKey in DEFAULT_CHOICES) {
            if (DEFAULT_CHOICES.hasOwnProperty(settingKey)) defaultSettings[settingKey] = DEFAULT_CHOICES[settingKey];
        }
        var defaultSizes = scaleSizeRatios(DEFAULT_SIZE_RATIOS, fontSize, rulerUnit, strokeUnit);
        for (var i = 0; i < PRESET_SIZE_KEYS.length; i++) defaultSettings[PRESET_SIZE_KEYS[i]] = defaultSizes[PRESET_SIZE_KEYS[i]];
        return defaultSettings;
    }

    /**
     * 斜線の左右の調整の既定値を返す。斜線の傾きぶん線幅の半分だけ戻す（右へはマイナス、左へはプラス、真下は 0）
     * @param {string} directionKey - しっぽの向き（"center" / "right" / "left"）
     * @param {number} strokeWidth - 線幅（pt）
     * @returns {number} 左右の調整（pt）
     */
    function getDefaultSlashShiftX(directionKey, strokeWidth) {
        if (directionKey === "right") return -strokeWidth / 2;
        if (directionKey === "left") return strokeWidth / 2;
        return 0;
    }

    /**
     * 文字サイズに対する比率から、寸法（pt）を作る。単位の小数第1位で丸め、書かれていない項目は DEFAULT_SIZE_RATIOS を使う
     * @param {Object} sizeRatios - 寸法のキーと比率
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {Object} rulerUnit - 長さの単位（getUnitInfo() の戻り値）
     * @param {Object} strokeUnit - 線幅の単位（getUnitInfo("strokeUnits") の戻り値）
     * @returns {Object} 寸法のキーと値（pt）
     */
    function scaleSizeRatios(sizeRatios, fontSize, rulerUnit, strokeUnit) {
        var scaledSizes = {};
        for (var i = 0; i < PRESET_SIZE_KEYS.length; i++) {
            var sizeKey = PRESET_SIZE_KEYS[i];
            var sizeRatio = (typeof sizeRatios[sizeKey] === "number") ? sizeRatios[sizeKey] : DEFAULT_SIZE_RATIOS[sizeKey];
            var sizeUnit = (sizeKey === "strokeWidth") ? strokeUnit : rulerUnit;
            scaledSizes[sizeKey] = unitToPt(Math.round(ptToUnit(fontSize * sizeRatio, sizeUnit) * 10) / 10, sizeUnit);
        }
        /* 線幅は 0 にしない / never a zero stroke */
        scaledSizes.strokeWidth = Math.max(scaledSizes.strokeWidth, unitToPt(0.1, strokeUnit));
        return scaledSizes;
    }

    /* 保存したプリセット。Illustrator を終了しても残す / Saved presets, kept across Illustrator sessions */
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "_presets", "persistent");

    /**
     * 保存したプリセットを読む
     * @returns {Object} プリセット名をキーにした { choices, position, slashAngle, useAreaFrame, isCapsule, tipDot, sizes } の集まり
     */
    function loadSavedPresets() {
        return presetSettingsStore.load({});
    }

    /**
     * 保存したプリセットを書く（書けなければ知らせる）
     * @param {Object} savedPresets - プリセット名をキーにした集まり
     * @returns {boolean} 書けたら true
     */
    function writeSavedPresets(savedPresets) {
        if (presetSettingsStore.save(savedPresets)) return true;
        alert(getLabel(LABELS.alert.presetSaveFailed));
        return false;
    }

    /**
     * 最初からあるプリセットの名前か（日本語・英語のどちらでも）
     * @param {string} presetName - プリセット名
     * @returns {boolean} 最初からあるものと同じ名前なら true
     */
    function isBuiltinPresetName(presetName) {
        for (var i = 0; i < BUILTIN_PRESETS.length; i++) {
            if (!BUILTIN_PRESETS[i]) continue;
            if (BUILTIN_PRESETS[i].name.ja === presetName || BUILTIN_PRESETS[i].name.en === presetName) return true;
        }
        return false;
    }

    /**
     * プリセットを設定の値にする。寸法は保存したプリセットなら pt のまま、最初からあるものなら文字サイズから作る
     * @param {Object} presetData - { choices, position, slashAngle, useAreaFrame, isCapsule, tipDot, sizes（pt）か ratios（比率）}
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {Object} rulerUnit - 長さの単位
     * @param {Object} strokeUnit - 線幅の単位
     * @returns {Object} 選択・位置・寸法の設定（BubbleSettings の一部）
     */
    function resolvePresetSettings(presetData, fontSize, rulerUnit, strokeUnit) {
        var presetSettings = {};
        var presetChoices = presetData.choices || {};
        var i;
        for (i = 0; i < PRESET_CHOICE_KEYS.length; i++) {
            var choiceKey = PRESET_CHOICE_KEYS[i];
            presetSettings[choiceKey] = (typeof presetChoices[choiceKey] === "string") ? presetChoices[choiceKey] : DEFAULT_CHOICES[choiceKey];
        }
        presetSettings.position = (typeof presetData.position === "number") ? presetData.position : DEFAULT_CHOICES.position;
        presetSettings.slashAngle = (typeof presetData.slashAngle === "number") ? presetData.slashAngle : DEFAULT_CHOICES.slashAngle;
        presetSettings.useAreaFrame = (typeof presetData.useAreaFrame === "boolean") ? presetData.useAreaFrame : DEFAULT_CHOICES.useAreaFrame;
        presetSettings.isCapsule = (typeof presetData.isCapsule === "boolean") ? presetData.isCapsule : DEFAULT_CHOICES.isCapsule;
        presetSettings.tipDot = (typeof presetData.tipDot === "boolean") ? presetData.tipDot : DEFAULT_CHOICES.tipDot;
        var scaledSizes = scaleSizeRatios(presetData.ratios || {}, fontSize, rulerUnit, strokeUnit);
        var savedSizes = presetData.sizes || {};
        for (i = 0; i < PRESET_SIZE_KEYS.length; i++) {
            var sizeKey = PRESET_SIZE_KEYS[i];
            presetSettings[sizeKey] = (typeof savedSizes[sizeKey] === "number") ? savedSizes[sizeKey] : scaledSizes[sizeKey];
        }
        /* 左右の調整を決めていない斜線のプリセットは、向きに合わせて線幅の半分ずらす / slash presets without a nudge get half the weight by direction */
        var hasSlashShiftX = (typeof savedSizes.slashShiftX === "number") || (presetData.ratios && typeof presetData.ratios.slashShiftX === "number");
        if (presetSettings.shapeKey === "slash" && !hasSlashShiftX) {
            presetSettings.slashShiftX = getDefaultSlashShiftX(presetSettings.directionKey, presetSettings.strokeWidth);
        }
        return presetSettings;
    }

    /**
     * プリセット名を尋ねる
     * @param {string} initialName - 入力欄に入れておく名前
     * @returns {string|null} プリセット名。キャンセルか空なら null
     */
    function showPresetNameDialog(initialName) {
        var nameDialog = new Window("dialog", getLabel(LABELS.dialog.presetSave));
        setupWindow(nameDialog);
        var nameRow = nameDialog.add("group");
        setupRow(nameRow, "left", PRESET_ROW_SPACING);
        nameRow.add("statictext", undefined, labelText(LABELS.fieldLabel.presetName));
        var nameInput = nameRow.add("edittext", undefined, initialName);
        nameInput.helpTip = getLabel(LABELS.tooltip.presetName);
        nameInput.characters = PRESET_NAME_CHARACTERS;
        nameInput.active = true;

        var buttonRow = addButtonRow(nameDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(nameDialog, SCRIPT_NAME + "_presetName");
        if (nameDialog.show() !== 1) return null;
        var presetName = nameInput.text.replace(/^\s+|\s+$/g, "");
        return presetName || null;
    }

    // =========================================
    // ダイアログのレイアウト補助 / Dialog layout helpers
    // =========================================

    /**
     * 横に並べた列の下端をそろえる。低い列の最後のパネルを伸ばす（中身は上揃えのまま）。
     * 一度レイアウトしてから測り、preferredSize と size の両方に入れる（preferredSize だけでは効かない）
     * @param {Array<{column: Object, lastPanel: Panel}>} columns - 列（パネル1つの列ならパネル自身）と、その列の最後のパネル
     * @returns {void}
     */
    function alignColumnBottoms(columns) {
        var dialogWindow = columns[0].lastPanel.window;
        dialogWindow.layout.layout(true);
        var maxHeight = 0;
        var i;
        for (i = 0; i < columns.length; i++) maxHeight = Math.max(maxHeight, columns[i].column.size.height);
        for (i = 0; i < columns.length; i++) {
            var lastPanel = columns[i].lastPanel;
            var panelHeight = lastPanel.size.height + maxHeight - columns[i].column.size.height;
            lastPanel.preferredSize.height = panelHeight;
            lastPanel.size = [lastPanel.size.width, panelHeight];
        }
        dialogWindow.layout.layout(true);
    }

    /**
     * 複数の項目名（statictext）の幅を、いちばん長いものに揃えて右揃えにする（実際の描画幅を測る）
     * @param {StaticText[]} labelControls - 揃える項目名
     * @returns {number} 揃えた幅（px）
     */
    function alignLabelWidths(labelControls) {
        var maxLabelWidth = 0;
        var i;
        for (i = 0; i < labelControls.length; i++) {
            /* preferredSize を返さない環境では実サイズで代用する。空の項目名（カプセル形状の行）はレイアウト前で size も無いので 0 とみなす
               Fall back to the laid-out size; an empty label (the capsule row) has no size before layout, so count it as 0 */
            var naturalWidth = labelControls[i].preferredSize.width;
            if (!naturalWidth && labelControls[i].size) naturalWidth = labelControls[i].size.width;
            naturalWidth = naturalWidth || 0;
            if (naturalWidth > maxLabelWidth) maxLabelWidth = naturalWidth;
        }
        for (i = 0; i < labelControls.length; i++) {
            labelControls[i].preferredSize.width = maxLabelWidth;
            labelControls[i].justify = "right";
        }
        return maxLabelWidth;
    }

    /**
     * 数値欄の値を読む（数値でなければ 0）
     * @param {EditText} numberInput - 入力欄
     * @returns {number} 欄の単位の値
     */
    function readFieldNumber(numberInput) {
        var fieldValue = parseFloat(numberInput.text);
        return isNaN(fieldValue) ? 0 : fieldValue;
    }

    // =========================================
    // ダイアログの部品 / Dialog controls
    // =========================================

    /**
     * 数値欄に値を書き込む（部品が直前の正しい値として控える lastValidText もそろえる）
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {string} fieldText - 入れる文字列（単位付き）
     * @returns {void}
     */
    function writeFieldText(numberInput, fieldText) {
        numberInput.text = fieldText;
        numberInput.lastValidText = fieldText;
    }

    /**
     * 直接入力を確定したときの処理を足す（部品の onChange で値をそろえたあとに呼ぶ）
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {Function} afterChange - そろえたあとに呼ぶ関数
     * @returns {void}
     */
    function chainAfterChange(numberInput, afterChange) {
        var normalizeOnChange = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeOnChange();
            afterChange();
        };
    }

    /**
     * 数値欄を足す（∧∨・↑↓キー・直接入力のどれで変えても onEdit を呼ぶ）
     * @param {Panel} parent - 追加先
     * @param {string} settingKey - 設定のキー（LABELS.fieldLabel・LABELS.tooltip も同じキー）
     * @param {string} fieldText - 初期値の文字列（単位付き）
     * @param {Object} fieldOptions - step / min / max / unit（addSteppedField() と同じ）
     * @param {Function} onEdit - 値を変えたときに呼ぶ関数
     * @returns {EditText} 入力欄（項目名は .fieldLabel）
     */
    function addNumberField(parent, settingKey, fieldText, fieldOptions, onEdit) {
        var numberInput = addSteppedField(parent, {
            label: labelText(LABELS.fieldLabel[settingKey]),
            text: fieldText,
            characters: FIELD_CHARACTERS, step: fieldOptions.step, min: fieldOptions.min, max: fieldOptions.max, unit: fieldOptions.unit,
            onStep: onEdit
        });
        numberInput.helpTip = getLabel(LABELS.tooltip[settingKey]);
        chainAfterChange(numberInput, onEdit);
        return numberInput;
    }

    /**
     * 長さの数値欄を足す（値は pt で受け取り、欄の単位を付けて表示する）
     * @param {Panel} parent - 追加先
     * @param {string} settingKey - 設定のキー
     * @param {number} valuePt - 初期値（pt）
     * @param {Object} unit - 欄の単位（getUnitInfo() の戻り値）
     * @param {number|undefined} minValue - 下限（欄の単位。下限なしは undefined）
     * @param {Function} onEdit - 値を変えたときに呼ぶ関数
     * @param {number} [step] - ∧∨・↑↓キーの増減量（省略時は 1）
     * @returns {EditText} 入力欄
     */
    function addLengthField(parent, settingKey, valuePt, unit, minValue, onEdit, step) {
        var unitSuffix = " " + unit.label;
        return addNumberField(parent, settingKey, formatStepperNumber(ptToUnit(valuePt, unit)) + unitSuffix,
            { step: step || 1, min: minValue, unit: unitSuffix }, onEdit);
    }

    /**
     * 項目名の付いた行を足す（項目名の幅はあとで alignLabelWidths() でそろえる）
     * @param {Panel} parent - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @returns {Group} 行（.rowLabel で項目名）
     */
    function addLabeledRow(parent, labelKey) {
        var labeledRow = parent.add("group");
        setupRow(labeledRow);
        labeledRow.rowLabel = labeledRow.add("statictext", undefined, labelText(LABELS.fieldLabel[labelKey]));
        return labeledRow;
    }

    /**
     * ラジオボタンの行に、設定の値で読み書きする関数を付ける（.radios の各ボタンは .settingValue を持つ）
     * @param {Group} radioRow - ラジオボタンの行
     * @param {string} selectedValue - 選んでおく設定の値（選択肢に無ければ先頭）
     * @returns {void}
     */
    function attachRadioValueAccessors(radioRow, selectedValue) {
        radioRow.getValue = function () {
            for (var i = 0; i < radioRow.radios.length; i++) {
                if (radioRow.radios[i].value) return radioRow.radios[i].settingValue;
            }
            return radioRow.radios[0].settingValue;
        };
        /* 選択肢に無い値なら先頭を選ぶ / pick the first when the value is not a choice */
        radioRow.setValue = function (settingValue) {
            var isFound = false;
            for (var i = 0; i < radioRow.radios.length; i++) {
                radioRow.radios[i].value = (radioRow.radios[i].settingValue === settingValue);
                if (radioRow.radios[i].value) isFound = true;
            }
            if (!isFound) radioRow.radios[0].value = true;
        };
        radioRow.setValue(selectedValue);
    }

    /**
     * ラジオボタンの行を足す。選択肢は [設定の値, LABELS.radio・LABELS.tooltip のキー] の組で渡す
     * @param {Panel} parent - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @param {string} selectedValue - 選んでおく設定の値
     * @param {Array<string[]>} choices - [設定の値, ラベルのキー] の配列
     * @param {Function} onChoice - 選び直したときに呼ぶ関数
     * @param {boolean} [isStacked] - true ならラジオボタンを縦に並べる（項目名は1つ目の高さ）
     * @returns {Group} 行（.getValue() / .setValue() / .radios / .rowLabel）
     */
    function addRadioRow(parent, labelKey, selectedValue, choices, onChoice, isStacked) {
        var radioRow = addLabeledRow(parent, labelKey);
        /* 縦に並べるときは同じ列のグループに入れる（同じ親なので排他は ScriptUI に任せられる） / stacked radios share one column group, so ScriptUI keeps them exclusive */
        var radioContainer = radioRow;
        if (isStacked) {
            radioRow.alignChildren = ["left", "top"];
            radioContainer = radioRow.add("group");
            radioContainer.orientation = "column";
            radioContainer.alignChildren = ["left", "center"];
            radioContainer.spacing = STACKED_RADIO_SPACING;
        }
        radioRow.radios = [];
        for (var i = 0; i < choices.length; i++) {
            var radioButton = radioContainer.add("radiobutton", undefined, getLabel(LABELS.radio[choices[i][1]]));
            radioButton.helpTip = getLabel(LABELS.tooltip[choices[i][1]]);
            radioButton.settingValue = choices[i][0];
            radioButton.onClick = onChoice;
            radioRow.radios.push(radioButton);
        }
        attachRadioValueAccessors(radioRow, selectedValue);
        return radioRow;
    }

    /**
     * ラジオボタンを表の形に並べる（項目名は真ん中の行の高さ）。行ごとに別のグループに入るので、排他はクリックで手で保つ
     * @param {Panel} parent - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @param {string} selectedValue - 選んでおく設定の値
     * @param {Array<Array<string[]|null>>} gridLayout - 行ごとの [設定の値, LABELS.radio・LABELS.tooltip のキー]（null は空きのマス）
     * @param {number} cellWidth - 1マスの幅
     * @param {Function} onChoice - 選び直したときに呼ぶ関数
     * @returns {Group} 行（addRadioRow() と同じ形）
     */
    function addRadioGrid(parent, labelKey, selectedValue, gridLayout, cellWidth, onChoice) {
        var gridRow = addLabeledRow(parent, labelKey);
        var gridColumn = gridRow.add("group");
        gridColumn.orientation = "column";
        gridColumn.alignChildren = ["left", "center"];
        gridColumn.spacing = RADIO_GRID_SPACING;
        gridRow.radios = [];

        /**
         * クリックしたラジオボタンだけを選び、ほかを外す
         * @param {RadioButton} clickedRadio - クリックしたラジオボタン
         * @returns {Function} onClick に入れる関数
         */
        function makeGridClickHandler(clickedRadio) {
            return function () {
                for (var k = 0; k < gridRow.radios.length; k++) gridRow.radios[k].value = (gridRow.radios[k] === clickedRadio);
                onChoice();
            };
        }

        for (var rowIndex = 0; rowIndex < gridLayout.length; rowIndex++) {
            var cellRow = gridColumn.add("group");
            cellRow.orientation = "row";
            cellRow.alignChildren = ["left", "center"];
            cellRow.spacing = 0;
            cellRow.margins = 0;
            for (var columnIndex = 0; columnIndex < gridLayout[rowIndex].length; columnIndex++) {
                var cellChoice = gridLayout[rowIndex][columnIndex];
                if (!cellChoice) {
                    cellRow.add("group").preferredSize = [cellWidth, 1];
                    continue;
                }
                var cellRadio = cellRow.add("radiobutton", undefined, getLabel(LABELS.radio[cellChoice[1]]));
                cellRadio.preferredSize.width = cellWidth;
                cellRadio.helpTip = getLabel(LABELS.tooltip[cellChoice[1]]);
                cellRadio.settingValue = cellChoice[0];
                cellRadio.onClick = makeGridClickHandler(cellRadio);
                gridRow.radios.push(cellRadio);
            }
        }
        attachRadioValueAccessors(gridRow, selectedValue);
        return gridRow;
    }

    /**
     * カスタムのカラーの色見本を足す。クリックで Illustrator 標準のカラーピッカーを開き、色を選んだら onPicked を呼ぶ
     * @param {Group} parent - 追加先（線のカラーの行）
     * @param {{data: Object}} colorState - 今のカスタムのカラー（.data を書き換える）
     * @param {Function} onPicked - 色を選んだときに呼ぶ関数
     * @returns {Group} 色見本
     */
    function addColorSwatch(parent, colorState, onPicked) {
        var swatchGroup = parent.add("group");
        swatchGroup.preferredSize = COLOR_SWATCH_SIZE;
        swatchGroup.minimumSize = COLOR_SWATCH_SIZE;
        swatchGroup.maximumSize = COLOR_SWATCH_SIZE;
        swatchGroup.helpTip = getLabel(LABELS.tooltip.colorSwatch);
        swatchGroup.onDraw = function () {
            var swatchGraphics = swatchGroup.graphics;
            var swatchRgb = colorDataToScreenRgb(colorState.data);
            /* newPath() を呼ばないとパスが積み重なる / without newPath() the paths accumulate */
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, COLOR_SWATCH_SIZE[0], COLOR_SWATCH_SIZE[1]);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, [swatchRgb[0], swatchRgb[1], swatchRgb[2], 1]));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, COLOR_SWATCH_SIZE[0] - 1, COLOR_SWATCH_SIZE[1] - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1));
        };
        swatchGroup.addEventListener("mousedown", function () {
            /* グレーで渡すとグレースケールで開くので、ドキュメントのカラーモードにして渡す。
               OK ならドキュメントのカラーモードの型、キャンセルなら渡した色がそのまま返る
               A gray color opens the picker in grayscale, so pass the document's color mode; OK returns that mode, Cancel the passed color */
            var pickedData = colorToColorData(app.showColorPicker(colorDataToColor(toDocumentColorData(colorState.data))));
            if (!pickedData || pickedData.toSource() === colorState.data.toSource()) return;
            colorState.data = pickedData;
            /* group には notify() が無いので、隠して表示し直して描き直す / groups have no notify(), so hide and show to redraw */
            swatchGroup.hide();
            swatchGroup.show();
            onPicked();
        });
        return swatchGroup;
    }

    /**
     * チェックボックスを足す
     * @param {Panel|Group} parent - 追加先
     * @param {string} settingKey - 設定のキー（LABELS.checkbox・LABELS.tooltip も同じキー）
     * @param {boolean} isChecked - 初期の状態
     * @returns {Checkbox} チェックボックス
     */
    function addSettingCheckbox(parent, settingKey, isChecked) {
        var settingCheckbox = parent.add("checkbox", undefined, getLabel(LABELS.checkbox[settingKey]));
        settingCheckbox.helpTip = getLabel(LABELS.tooltip[settingKey]);
        settingCheckbox.value = isChecked;
        return settingCheckbox;
    }

    // =========================================
    // ダイアログのパネル / Dialog panels
    // =========================================

    /**
     * @typedef {Object} DialogHandlers
     * @property {Function} onEdit - 数値・色を変えたとき（プリセットの選択を外してプレビュー）
     * @property {Function} onChoice - ラジオボタンを選び直したとき（使わない項目の切り替えとプレビュー）
     * @property {Function} onTailChoice - しっぽの形状・向きを選び直したとき（斜線の左右の調整を入れ直してから onChoice）
     */

    /**
     * 上部中央のプリセットの行を作る
     * @param {Window} bubbleDialog - ダイアログ
     * @returns {{dropdown: DropDownList, saveIcon: Group, deleteIcon: Group}} プリセットのコントロール
     */
    function buildPresetRow(bubbleDialog) {
        var presetRow = bubbleDialog.add("group");
        setupRow(presetRow, "center", PRESET_ROW_SPACING);
        presetRow.add("statictext", undefined, labelText(LABELS.fieldLabel.preset));
        var presetDropdown = presetRow.add("dropdownlist");
        presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
        presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);
        var presetSaveIcon = addIconButton(presetRow, PRESET_SAVE_ICON_SIZE, drawSaveIcon);
        presetSaveIcon.helpTip = getLabel(LABELS.tooltip.presetSave);
        var presetDeleteIcon = addIconButton(presetRow, ICON_BUTTON_SIZE, drawTrashIcon);
        presetDeleteIcon.helpTip = getLabel(LABELS.tooltip.presetDelete);
        return { dropdown: presetDropdown, saveIcon: presetSaveIcon, deleteIcon: presetDeleteIcon };
    }

    /**
     * ［しっぽ］パネルを作る
     * @param {Group} parent - 追加先（左の列）
     * @param {BubbleSettings} initialSettings - 開いたときの設定
     * @param {Object} rulerUnit - 長さの単位
     * @param {DialogHandlers} dialogHandlers - 変えたときの処理
     * @returns {Object} しっぽのコントロール（.panel と各行・欄）
     */
    function buildTailPanel(parent, initialSettings, rulerUnit, dialogHandlers) {
        var tailPanel = parent.add("panel", undefined, getLabel(LABELS.panel.tail));
        setupPanel(tailPanel);
        tailPanel.alignment = ["fill", "top"];
        var tailControls = { panel: tailPanel };
        tailControls.shapeRow = addRadioGrid(tailPanel, "shape", initialSettings.shapeKey, [
            [["notch", "notch"], ["slash", "slash"], ["curly", "curly"]],
            [["round", "round"], ["loop", "loop"], ["horn", "horn"]],
            [["thought", "thought"], null, null]
        ], SHAPE_GRID_CELL_WIDTH, dialogHandlers.onTailChoice);
        /* 形状は行が多いので、項目名を1行目の高さに置く / the shape grid is tall, so its label sits on the first row */
        tailControls.shapeRow.alignChildren = ["left", "top"];
        tailControls.directionRow = addRadioRow(tailPanel, "direction", initialSettings.directionKey,
            [["center", "straightDown"], ["right", "downRight"], ["left", "downLeft"]], dialogHandlers.onTailChoice);
        tailControls.slashAngleInput = addNumberField(tailPanel, "slashAngle", formatStepperNumber(initialSettings.slashAngle) + "°",
            { step: 5, min: MIN_SLASH_ANGLE, max: 90, unit: "°" }, dialogHandlers.onEdit);
        tailControls.rootRow = addRadioRow(tailPanel, "root", initialSettings.rootKey,
            [["center", "rootCenter"], ["edge", "rootEdge"]], dialogHandlers.onChoice);
        /* 線と離した斜線の微調整は、少しずつ動かせるよう細かい刻みにする / fine steps for nudging a detached slash */
        tailControls.slashShiftXInput = addLengthField(tailPanel, "slashShiftX", initialSettings.slashShiftX, rulerUnit, undefined, dialogHandlers.onEdit, SLASH_SHIFT_STEP);
        tailControls.slashShiftYInput = addLengthField(tailPanel, "slashShiftY", initialSettings.slashShiftY, rulerUnit, undefined, dialogHandlers.onEdit, SLASH_SHIFT_STEP);

        /* 位置の数値欄とスライダーは、どちらで変えてももう一方に写す / the position field and slider mirror each other */
        var positionInput = addNumberField(tailPanel, "position", formatStepperNumber(initialSettings.position) + "%",
            { step: 5, min: 0, max: 100, unit: "%" }, function () {
                positionSlider.value = readFieldNumber(positionInput);
                dialogHandlers.onEdit();
            });
        /* 位置の下の行にスライダーを置き、項目名の幅だけ字下げして数値欄とそろえる。右端は少しあける
           The slider sits on the next row, indented past the labels, with a little room on the right */
        var positionSliderRow = tailPanel.add("group");
        setupRow(positionSliderRow, ["fill", "center"], STEPPER_FIELD_SPACING);
        positionSliderRow.margins = [0, 0, POSITION_SLIDER_RIGHT_INSET, 0];
        var positionSliderIndent = positionSliderRow.add("group");
        var positionSlider = positionSliderRow.add("slider", undefined, initialSettings.position, 0, 100);
        positionSlider.preferredSize.width = POSITION_SLIDER_WIDTH;
        positionSlider.alignment = ["fill", "center"];
        positionSlider.helpTip = getLabel(LABELS.tooltip.positionSlider);
        positionSlider.onChanging = function () {
            /* shift を押しながらなら POSITION_SLIDER_SHIFT_STEP 刻みに吸着させる / snap to POSITION_SLIDER_SHIFT_STEP while Shift is held */
            var sliderValue = Math.round(positionSlider.value);
            if (ScriptUI.environment.keyboardState.shiftKey) {
                sliderValue = Math.round(positionSlider.value / POSITION_SLIDER_SHIFT_STEP) * POSITION_SLIDER_SHIFT_STEP;
                positionSlider.value = sliderValue;
            }
            writeFieldText(positionInput, formatStepperNumber(sliderValue) + "%");
        };
        positionSlider.onChange = function () {
            positionSlider.onChanging();
            dialogHandlers.onEdit();
        };
        tailControls.positionInput = positionInput;
        tailControls.positionSlider = positionSlider;

        tailControls.gapWidthInput = addLengthField(tailPanel, "gapWidth", initialSettings.gapWidth, rulerUnit, 0, dialogHandlers.onEdit);
        tailControls.tailWidthInput = addLengthField(tailPanel, "tailWidth", initialSettings.tailWidth, rulerUnit, 0, dialogHandlers.onEdit);
        tailControls.tailHeightInput = addLengthField(tailPanel, "tailHeight", initialSettings.tailHeight, rulerUnit, 0, dialogHandlers.onEdit);
        /* 先端の●は高さの下。項目名の列は空けて、欄の左端にそろえる / the tip dot checkbox sits under the height, aligned with the fields */
        var tipDotRow = tailPanel.add("group");
        setupRow(tipDotRow);
        tipDotRow.rowLabel = tipDotRow.add("statictext", undefined, "");
        tailControls.tipDotCheckbox = addSettingCheckbox(tipDotRow, "tipDot", initialSettings.tipDot);
        tailControls.tipDotCheckbox.onClick = dialogHandlers.onEdit;

        /* 項目名をいちばん長いものに揃えて右揃えにし、スライダーの字下げも同じ幅にする / align the labels to the widest and indent the slider likewise */
        positionSliderIndent.preferredSize.width = alignLabelWidths([
            tailControls.shapeRow.rowLabel, tailControls.directionRow.rowLabel, tailControls.slashAngleInput.fieldLabel,
            tailControls.rootRow.rowLabel, tailControls.slashShiftXInput.fieldLabel, tailControls.slashShiftYInput.fieldLabel, positionInput.fieldLabel, tailControls.gapWidthInput.fieldLabel,
            tailControls.tailWidthInput.fieldLabel, tailControls.tailHeightInput.fieldLabel, tipDotRow.rowLabel
        ]);
        return tailControls;
    }

    /**
     * ［線/枠］パネルを作る
     * @param {Group} parent - 追加先（2カラムの行）
     * @param {BubbleSettings} initialSettings - 開いたときの設定
     * @param {string} initialSideKey - 開いたときの配置（テキストの組み方から決めたもの）
     * @param {Object} rulerUnit - 長さの単位
     * @param {Object} strokeUnit - 線幅の単位
     * @param {DialogHandlers} dialogHandlers - 変えたときの処理
     * @returns {Object} 線/枠のコントロール（.panel と各行・欄、.colorState）
     */
    function buildLinePanel(parent, initialSettings, initialSideKey, rulerUnit, strokeUnit, dialogHandlers) {
        var linePanel = parent.add("panel", undefined, getLabel(LABELS.panel.line));
        setupPanel(linePanel);
        /* 縦の fill は効かず天地中央に置かれるので、上に揃える（下端は表示前にそろえる） / vertical fill centers the panel, so pin it to the top and align the bottoms before showing */
        linePanel.alignment = ["fill", "top"];
        var lineControls = { panel: linePanel, colorState: { data: initialSettings.customColor } };
        lineControls.frameRow = addRadioRow(linePanel, "frame", initialSettings.frameKey,
            [["underline", "underline"], ["lShape", "lShape"], ["bracket", "bracket"], ["box", "box"], ["twoLines", "twoLines"]], dialogHandlers.onChoice, true);
        /* 配置は上下左右の位置に合わせて3行3列に並べる / the sides sit where they draw, in a 3x3 grid */
        lineControls.sideRow = addRadioGrid(linePanel, "side", initialSideKey, [
            [null, ["top", "sideTop"], null],
            [["left", "sideLeft"], null, ["right", "sideRight"]],
            [null, ["bottom", "sideBottom"], null]
        ], SIDE_GRID_CELL_WIDTH, dialogHandlers.onChoice);
        lineControls.offsetInput = addLengthField(linePanel, "offset", initialSettings.offset, rulerUnit, undefined, dialogHandlers.onEdit);
        lineControls.extensionInput = addLengthField(linePanel, "extension", initialSettings.extension, rulerUnit, undefined, dialogHandlers.onEdit);
        lineControls.tickHeightInput = addLengthField(linePanel, "tickHeight", initialSettings.tickHeight, rulerUnit, 0, dialogHandlers.onEdit);
        lineControls.cornerRadiusInput = addLengthField(linePanel, "cornerRadius", initialSettings.cornerRadius, rulerUnit, 0, dialogHandlers.onEdit);
        /* カプセル形状は角丸の半径の下。項目名の列は空けて、欄の左端にそろえる / the capsule checkbox sits under the corner radius, aligned with the fields */
        var capsuleRow = linePanel.add("group");
        setupRow(capsuleRow);
        capsuleRow.rowLabel = capsuleRow.add("statictext", undefined, "");
        lineControls.capsuleCheckbox = addSettingCheckbox(capsuleRow, "isCapsule", initialSettings.isCapsule);
        lineControls.capsuleCheckbox.onClick = dialogHandlers.onChoice;
        lineControls.capsuleRow = capsuleRow;
        lineControls.shadowDistanceInput = addLengthField(linePanel, "shadowDistance", initialSettings.shadowDistance, rulerUnit, 0, dialogHandlers.onEdit);
        lineControls.sideGapWidthInput = addLengthField(linePanel, "sideGapWidth", initialSettings.sideGapWidth, rulerUnit, 0, dialogHandlers.onEdit);
        lineControls.strokeWidthInput = addLengthField(linePanel, "strokeWidth", initialSettings.strokeWidth, strokeUnit, 0.1, dialogHandlers.onEdit);

        var colorRow = addRadioRow(linePanel, "lineColor", initialSettings.colorKey,
            [["text", "colorText"], ["custom", "colorCustom"]], dialogHandlers.onChoice);
        /* カスタムのラジオボタンは文言を出さず、すぐ右の色見本をラベル代わりにする / the custom radio has no text; the swatch beside it is its label */
        colorRow.radios[1].text = "";
        addColorSwatch(colorRow, lineControls.colorState, function () {
            colorRow.setValue("custom");
            dialogHandlers.onEdit();
        });
        lineControls.colorRow = colorRow;
        lineControls.capRow = addRadioRow(linePanel, "cap", initialSettings.capKey,
            [["butt", "buttCap"], ["round", "roundCap"]], dialogHandlers.onChoice);

        alignLabelWidths([
            lineControls.frameRow.rowLabel, lineControls.sideRow.rowLabel, lineControls.offsetInput.fieldLabel,
            lineControls.extensionInput.fieldLabel, lineControls.tickHeightInput.fieldLabel, lineControls.cornerRadiusInput.fieldLabel, capsuleRow.rowLabel,
            lineControls.shadowDistanceInput.fieldLabel, lineControls.sideGapWidthInput.fieldLabel, lineControls.strokeWidthInput.fieldLabel, colorRow.rowLabel, lineControls.capRow.rowLabel
        ]);
        return lineControls;
    }

    /**
     * ［テキストに関してのオプション］パネルを作る
     * @param {Group} parent - 追加先（左の列）
     * @param {BubbleSettings} initialSettings - 開いたときの設定
     * @returns {Object} オプションのコントロール（.panel と各チェックボックス）
     */
    function buildTextOptionsPanel(parent, initialSettings) {
        var textOptionsPanel = parent.add("panel", undefined, getLabel(LABELS.panel.options));
        setupPanel(textOptionsPanel, 6);
        textOptionsPanel.alignment = ["fill", "top"];
        return {
            panel: textOptionsPanel,
            useAreaFrameCheckbox: addSettingCheckbox(textOptionsPanel, "useAreaFrame", initialSettings.useAreaFrame),
            convertToAreaTextCheckbox: addSettingCheckbox(textOptionsPanel, "convertToAreaText", initialSettings.convertToAreaText),
            groupWithTextCheckbox: addSettingCheckbox(textOptionsPanel, "groupWithText", initialSettings.groupWithText)
        };
    }

    // =========================================
    // ダイアログの値 / Dialog values
    // =========================================

    /**
     * ダイアログの値を pt の設定にまとめる
     * @param {Object} dialogControls - 各パネルのコントロールをまとめたもの
     * @param {Object} rulerUnit - 長さの単位
     * @param {Object} strokeUnit - 線幅の単位
     * @returns {BubbleSettings} 設定
     */
    function readControlSettings(dialogControls, rulerUnit, strokeUnit) {
        var bubbleSettings = {};
        for (var choiceKey in dialogControls.choiceRows) {
            if (dialogControls.choiceRows.hasOwnProperty(choiceKey)) bubbleSettings[choiceKey] = dialogControls.choiceRows[choiceKey].getValue();
        }
        for (var sizeKey in dialogControls.lengthFields) {
            if (!dialogControls.lengthFields.hasOwnProperty(sizeKey)) continue;
            var lengthField = dialogControls.lengthFields[sizeKey];
            bubbleSettings[sizeKey] = unitToPt(readFieldNumber(lengthField.input), lengthField.isStroke ? strokeUnit : rulerUnit);
        }
        bubbleSettings.position = readFieldNumber(dialogControls.positionInput);
        bubbleSettings.slashAngle = readFieldNumber(dialogControls.slashAngleInput);
        bubbleSettings.customColor = dialogControls.colorState.data;
        bubbleSettings.useAreaFrame = dialogControls.useAreaFrameCheckbox.value;
        bubbleSettings.isCapsule = dialogControls.capsuleCheckbox.value;
        bubbleSettings.tipDot = dialogControls.tipDotCheckbox.value;
        bubbleSettings.convertToAreaText = dialogControls.convertToAreaTextCheckbox.value;
        bubbleSettings.groupWithText = dialogControls.groupWithTextCheckbox.value;
        bubbleSettings.preview = dialogControls.previewCheckbox.value;
        return bubbleSettings;
    }

    /**
     * プリセットの設定をダイアログに入れる（選択・位置・角度・寸法と、枠を基準にするか）
     * @param {Object} dialogControls - 各パネルのコントロールをまとめたもの
     * @param {Object} presetSettings - resolvePresetSettings() の戻り値
     * @param {Object} rulerUnit - 長さの単位
     * @param {Object} strokeUnit - 線幅の単位
     * @returns {void}
     */
    function applySettingsToControls(dialogControls, presetSettings, rulerUnit, strokeUnit) {
        /* 線のカラーはプリセットに入らないので触らない / the line color is not part of a preset */
        for (var i = 0; i < PRESET_CHOICE_KEYS.length; i++) {
            dialogControls.choiceRows[PRESET_CHOICE_KEYS[i]].setValue(presetSettings[PRESET_CHOICE_KEYS[i]]);
        }
        for (var sizeKey in dialogControls.lengthFields) {
            if (!dialogControls.lengthFields.hasOwnProperty(sizeKey)) continue;
            var lengthField = dialogControls.lengthFields[sizeKey];
            var lengthUnit = lengthField.isStroke ? strokeUnit : rulerUnit;
            writeFieldText(lengthField.input, formatStepperNumber(ptToUnit(presetSettings[sizeKey], lengthUnit)) + " " + lengthUnit.label);
        }
        writeFieldText(dialogControls.positionInput, formatStepperNumber(presetSettings.position) + "%");
        dialogControls.positionSlider.value = presetSettings.position;
        writeFieldText(dialogControls.slashAngleInput, formatStepperNumber(presetSettings.slashAngle) + "°");
        dialogControls.useAreaFrameCheckbox.value = presetSettings.useAreaFrame;
        dialogControls.capsuleCheckbox.value = presetSettings.isCapsule;
        dialogControls.tipDotCheckbox.value = presetSettings.tipDot;
    }

    /**
     * 今の選択で使わない項目を無効にする
     * @param {Object} dialogControls - 各パネルのコントロールをまとめたもの
     * @param {boolean} isAllAreaText - 選択がすべてエリア内文字か
     * @returns {void}
     */
    function updateControlsEnabled(dialogControls, isAllAreaText) {
        var shapeKey = dialogControls.choiceRows.shapeKey.getValue();
        var isSlash = (shapeKey === "slash");
        var isStraight = (dialogControls.choiceRows.directionKey.getValue() === "center");
        var frameKey = dialogControls.choiceRows.frameKey.getValue();
        var lengthFields = dialogControls.lengthFields;
        dialogControls.choiceRows.rootKey.enabled = isSlash && !isStraight;
        /* 左右・上下の調整は、線と離した斜線（根元が中央か、真下）のときだけ / nudges apply only to a detached slash */
        var isDetachedSlash = isSlash && (isStraight || dialogControls.choiceRows.rootKey.getValue() === "center");
        setSteppedFieldEnabled(lengthFields.slashShiftX.input, isDetachedSlash);
        setSteppedFieldEnabled(lengthFields.slashShiftY.input, isDetachedSlash);
        /* 斜線は横の長さを角度で決める / a slash takes its run from the angle */
        setSteppedFieldEnabled(lengthFields.tailWidth.input, !isStraight && !isSlash);
        setSteppedFieldEnabled(dialogControls.slashAngleInput, !isStraight && isSlash);
        /* ●は先端が1点に決まるしっぽだけ / the dot needs a tail with a single tip */
        dialogControls.tipDotCheckbox.enabled = (shapeKey === "notch" || isSlash || shapeKey === "curly" || shapeKey === "round");
        var hasTicks = (frameKey === "bracket" || frameKey === "lShape");
        setSteppedFieldEnabled(lengthFields.tickHeight.input, hasTicks);
        /* カプセル形状のときは半径を自動で決めるので、角丸の半径は使わない / a capsule decides the radius itself */
        dialogControls.capsuleCheckbox.enabled = (hasTicks || frameKey === "box");
        setSteppedFieldEnabled(lengthFields.cornerRadius.input, (hasTicks || frameKey === "box") && !dialogControls.capsuleCheckbox.value);
        setSteppedFieldEnabled(lengthFields.shadowDistance.input, frameKey === "box");
        setSteppedFieldEnabled(lengthFields.sideGapWidth.input, frameKey === "box");
        dialogControls.convertToAreaTextCheckbox.enabled = (frameKey === "box");
        /* すべてエリア内文字で枠を基準にするときは、間隔・左右の伸びを使わない / unused when every text is area type measured by its frame */
        var isLengthFromGlyphs = !(dialogControls.useAreaFrameCheckbox.value && isAllAreaText);
        setSteppedFieldEnabled(lengthFields.offset.input, isLengthFromGlyphs);
        setSteppedFieldEnabled(lengthFields.extension.input, isLengthFromGlyphs);
    }

    /**
     * 設定をプリセットの形にする（寸法は pt）
     * @param {BubbleSettings} bubbleSettings - 設定
     * @returns {Object} { choices, position, slashAngle, useAreaFrame, isCapsule, tipDot, sizes }
     */
    function settingsToPresetData(bubbleSettings) {
        var presetData = {
            choices: {}, position: bubbleSettings.position, slashAngle: bubbleSettings.slashAngle, useAreaFrame: bubbleSettings.useAreaFrame,
            isCapsule: bubbleSettings.isCapsule, tipDot: bubbleSettings.tipDot, sizes: {}
        };
        var i;
        for (i = 0; i < PRESET_CHOICE_KEYS.length; i++) presetData.choices[PRESET_CHOICE_KEYS[i]] = bubbleSettings[PRESET_CHOICE_KEYS[i]];
        for (i = 0; i < PRESET_SIZE_KEYS.length; i++) presetData.sizes[PRESET_SIZE_KEYS[i]] = bubbleSettings[PRESET_SIZE_KEYS[i]];
        return presetData;
    }

    // =========================================
    // プリセットの操作 / Preset controls
    // =========================================

    /**
     * プリセットの行に、選ぶ・保存・削除の処理を付ける
     * @param {{dropdown: DropDownList, saveIcon: Group, deleteIcon: Group}} presetControls - buildPresetRow() の戻り値
     * @param {Function} readCurrentSettings - 今の設定を返す関数
     * @param {Function} applyPresetData - プリセットをダイアログに入れる関数（プリセットのデータを受け取る）
     * @returns {{clearSelection: Function, fill: Function}} 選択を外す関数と、ドロップダウンを作り直す関数
     */
    function bindPresetControls(presetControls, readCurrentSettings, applyPresetData) {
        var presetDropdown = presetControls.dropdown;
        var presetEntries = [];       /* ドロップダウンの項目ごとの { name, data, isSaved }（区切り線は null） / per item; null for separators */
        var isFillingDropdown = false; /* 作り直し・選択を外す間は onChange を無視する / ignore onChange while rebuilding or clearing */

        /**
         * 選んでいるプリセットの項目を返す
         * @returns {Object|null} { name, data, isSaved }。選んでいなければ null
         */
        function getSelectedPresetEntry() {
            return presetDropdown.selection ? presetEntries[presetDropdown.selection.index] : null;
        }

        /**
         * 削除のアイコンは、保存したプリセットを選んでいるときだけ使えるようにする
         * @returns {void}
         */
        function updatePresetIcons() {
            var selectedEntry = getSelectedPresetEntry();
            setIconButtonEnabled(presetControls.deleteIcon, !!(selectedEntry && selectedEntry.isSaved));
        }

        /**
         * ドロップダウンに項目か区切り線を足す
         * @param {Object|null} presetEntry - { name, data, isSaved }。null なら区切り線
         * @returns {void}
         */
        function addPresetEntry(presetEntry) {
            presetDropdown.add(presetEntry ? "item" : "separator", presetEntry ? presetEntry.name : undefined);
            presetEntries.push(presetEntry);
        }

        /**
         * ドロップダウンを作り直す（最初からあるもの、区切り線、保存したもの）
         * @param {string|null} selectedName - 選んでおく保存したプリセット名（なければ選ばない）
         * @returns {void}
         */
        function fillPresetDropdown(selectedName) {
            isFillingDropdown = true;
            presetDropdown.removeAll();
            presetEntries = [];
            var i;
            for (i = 0; i < BUILTIN_PRESETS.length; i++) {
                addPresetEntry(BUILTIN_PRESETS[i] ? { name: getLabel(BUILTIN_PRESETS[i].name), data: BUILTIN_PRESETS[i], isSaved: false } : null);
            }
            var savedPresets = loadSavedPresets();
            var savedNames = [];
            for (var presetName in savedPresets) {
                /* 最初からあるプリセットと同じ名前（昇格したもの）は出さない / skip saved presets named like a built-in (promoted ones) */
                if (savedPresets.hasOwnProperty(presetName) && !isBuiltinPresetName(presetName)) savedNames.push(presetName);
            }
            savedNames.sort();
            if (savedNames.length) addPresetEntry(null);
            for (i = 0; i < savedNames.length; i++) {
                addPresetEntry({ name: savedNames[i], data: savedPresets[savedNames[i]], isSaved: true });
                if (savedNames[i] === selectedName) presetDropdown.selection = presetDropdown.items.length - 1;
            }
            isFillingDropdown = false;
            updatePresetIcons();
        }

        /**
         * 設定を手で変えたら、プリセットの選択を外す
         * @returns {void}
         */
        function clearPresetSelection() {
            if (isFillingDropdown || !presetDropdown.selection) return;
            isFillingDropdown = true;
            presetDropdown.selection = null;
            isFillingDropdown = false;
            updatePresetIcons();
        }

        presetDropdown.onChange = function () {
            if (isFillingDropdown) return;
            updatePresetIcons();
            var selectedEntry = getSelectedPresetEntry();
            if (selectedEntry) applyPresetData(selectedEntry.data);
        };

        presetControls.saveIcon.onClick = function () {
            var selectedEntry = getSelectedPresetEntry();
            var presetName = showPresetNameDialog((selectedEntry && selectedEntry.isSaved) ? selectedEntry.name : "");
            if (!presetName) return;
            if (isBuiltinPresetName(presetName)) {
                alert(getLabel(LABELS.alert.presetBuiltinName));
                return;
            }
            var savedPresets = loadSavedPresets();
            /* 同じ名前の上書きは確認する（既定のボタンは［いいえ］） / confirm overwriting, defaulting to No */
            if (savedPresets.hasOwnProperty(presetName) && !confirm(getLabel(LABELS.confirm.presetOverwrite, { name: presetName }), true)) return;
            savedPresets[presetName] = settingsToPresetData(readCurrentSettings());
            if (writeSavedPresets(savedPresets)) fillPresetDropdown(presetName);
        };

        presetControls.deleteIcon.onClick = function () {
            var selectedEntry = getSelectedPresetEntry();
            if (!selectedEntry || !selectedEntry.isSaved) return;
            if (!confirm(getLabel(LABELS.confirm.presetDelete, { name: selectedEntry.name }), true)) return;
            var savedPresets = loadSavedPresets();
            delete savedPresets[selectedEntry.name];
            if (writeSavedPresets(savedPresets)) fillPresetDropdown(null);
        };

        return { clearSelection: clearPresetSelection, fill: fillPresetDropdown };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 吹き出しのダイアログを表示する。プレビューは線としっぽを描いては消す
     * @param {BubbleTarget[]} bubbleTargets - 対象
     * @param {BubbleSettings} initialSettings - 開いたときの設定
     * @param {Object} rulerUnit - 長さの単位（getUnitInfo() の戻り値）
     * @param {Object} strokeUnit - 線幅の単位（getUnitInfo("strokeUnits") の戻り値）
     * @returns {BubbleSettings|null} OK なら設定、キャンセルなら null
     */
    function showSpeechBubbleDialog(bubbleTargets, initialSettings, rulerUnit, strokeUnit) {
        var bubbleDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(bubbleDialog);
        var previewItems = [];
        var dialogHandlers = {
            onEdit: function () { refreshPreview(); },
            onChoice: function () {
                updateControlsEnabled(dialogControls, areaTextState.isAllAreaText);
                refreshPreview();
            },
            /* 形状・向きを選び直したら、斜線の左右の調整を向きに合わせて入れ直す / re-pick the slash nudge when the shape or direction changes */
            onTailChoice: function () {
                var currentSettings = readDialogSettings();
                if (currentSettings.shapeKey === "slash") {
                    var shiftX = getDefaultSlashShiftX(currentSettings.directionKey, currentSettings.strokeWidth);
                    writeFieldText(dialogControls.lengthFields.slashShiftX.input, formatStepperNumber(ptToUnit(shiftX, rulerUnit)) + " " + rulerUnit.label);
                }
                dialogHandlers.onChoice();
            }
        };

        var presetControls = buildPresetRow(bubbleDialog);

        /* 左に［しっぽ］と［テキストに関してのオプション］、右に［線/枠］の2カラム / two columns: tail and text options on the left, line/frame on the right */
        var columnsGroup = bubbleDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];
        leftColumn.alignment = ["fill", "top"];
        leftColumn.spacing = WINDOW_SPACING;
        var tailControls = buildTailPanel(leftColumn, initialSettings, rulerUnit, dialogHandlers);
        var textOptionControls = buildTextOptionsPanel(leftColumn, initialSettings);
        /* 配置は開くたびにテキストの組み方から選ぶ（横組みは下、縦組みは右） / the side is picked from the text orientation on every open */
        var lineControls = buildLinePanel(columnsGroup, initialSettings, getAutoSideKey(bubbleTargets), rulerUnit, strokeUnit, dialogHandlers);

        /* ボタン（左にプレビュー） / Buttons, with Preview on the left */
        var buttonRow = addButtonRow(bubbleDialog);
        var previewCheckbox = addSettingCheckbox(buttonRow.leftGroup, "preview", initialSettings.preview);
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /* 設定のキーと、それを入れる行・欄をまとめる / where each setting lives in the dialog */
        var dialogControls = {
            choiceRows: {
                shapeKey: tailControls.shapeRow, directionKey: tailControls.directionRow, rootKey: tailControls.rootRow,
                frameKey: lineControls.frameRow, sideKey: lineControls.sideRow, capKey: lineControls.capRow, colorKey: lineControls.colorRow
            },
            lengthFields: {
                gapWidth: { input: tailControls.gapWidthInput }, tailWidth: { input: tailControls.tailWidthInput },
                tailHeight: { input: tailControls.tailHeightInput }, offset: { input: lineControls.offsetInput },
                slashShiftX: { input: tailControls.slashShiftXInput }, slashShiftY: { input: tailControls.slashShiftYInput },
                extension: { input: lineControls.extensionInput }, tickHeight: { input: lineControls.tickHeightInput },
                cornerRadius: { input: lineControls.cornerRadiusInput }, shadowDistance: { input: lineControls.shadowDistanceInput },
                sideGapWidth: { input: lineControls.sideGapWidthInput },
                strokeWidth: { input: lineControls.strokeWidthInput, isStroke: true }
            },
            positionInput: tailControls.positionInput,
            positionSlider: tailControls.positionSlider,
            slashAngleInput: tailControls.slashAngleInput,
            colorState: lineControls.colorState,
            capsuleCheckbox: lineControls.capsuleCheckbox,
            tipDotCheckbox: tailControls.tipDotCheckbox,
            useAreaFrameCheckbox: textOptionControls.useAreaFrameCheckbox,
            convertToAreaTextCheckbox: textOptionControls.convertToAreaTextCheckbox,
            groupWithTextCheckbox: textOptionControls.groupWithTextCheckbox,
            previewCheckbox: previewCheckbox
        };
        var areaTextState = getAreaTextState(bubbleTargets);

        /**
         * 今の設定を読む
         * @returns {BubbleSettings} 設定
         */
        function readDialogSettings() {
            return readControlSettings(dialogControls, rulerUnit, strokeUnit);
        }

        /**
         * プレビューで描いたものを消す
         * @returns {void}
         */
        function clearPreview() {
            removePageItems(previewItems);
            previewItems = [];
        }

        /**
         * 今の設定でプレビューを描き直す
         * @returns {void}
         */
        function drawPreview() {
            clearPreview();
            if (previewCheckbox.value) {
                var drawnBubbles = drawBubbles(bubbleTargets, readDialogSettings()).drawn;
                for (var i = 0; i < drawnBubbles.length; i++) previewItems.push(drawnBubbles[i].group);
            }
            app.redraw();
        }

        /**
         * 設定を手で変えたとき：プリセットの選択を外してプレビューを描き直す
         * @returns {void}
         */
        function refreshPreview() {
            presetBinding.clearSelection();
            drawPreview();
        }

        var presetBinding = bindPresetControls(presetControls, readDialogSettings, function (presetData) {
            var presetSettings = resolvePresetSettings(presetData, bubbleTargets[0].fontSize, rulerUnit, strokeUnit);
            /* 配置を決めていないプリセットは、テキストの組み方から選ぶ / presets without a side follow the text orientation */
            if (presetSettings.sideKey === "auto") presetSettings.sideKey = getAutoSideKey(bubbleTargets);
            applySettingsToControls(dialogControls, presetSettings, rulerUnit, strokeUnit);
            updateControlsEnabled(dialogControls, areaTextState.isAllAreaText);
            drawPreview();
        });

        previewCheckbox.onClick = drawPreview;
        dialogControls.useAreaFrameCheckbox.onClick = dialogHandlers.onChoice;
        /* エリア内文字を選んでいないときは使わない / unused without area type in the selection */
        dialogControls.useAreaFrameCheckbox.enabled = areaTextState.hasAreaText;
        presetBinding.fill(null);
        updateControlsEnabled(dialogControls, areaTextState.isAllAreaText);

        /* 開く前にプレビューを描いておく / draw the preview before the dialog opens */
        if (previewCheckbox.value) drawPreview();

        alignColumnBottoms([{ column: leftColumn, lastPanel: textOptionControls.panel }, { column: lineControls.panel, lastPanel: lineControls.panel }]);
        addVersionFooter(bubbleDialog);
        prepareDialogWindow(bubbleDialog, SCRIPT_NAME);
        var dialogResult = bubbleDialog.show();
        var dialogSettings = readDialogSettings();
        clearPreview();
        app.redraw();
        return (dialogResult === 1) ? dialogSettings : null;
    }

    /**
     * 選択の中のエリア内文字の有無を調べる
     * @param {BubbleTarget[]} bubbleTargets - 対象
     * @returns {{hasAreaText: boolean, isAllAreaText: boolean}} 1つでもあるか、すべてそうか
     */
    function getAreaTextState(bubbleTargets) {
        var areaTextState = { hasAreaText: false, isAllAreaText: true };
        for (var i = 0; i < bubbleTargets.length; i++) {
            if (bubbleTargets[i].areaFrameBounds) areaTextState.hasAreaText = true;
            else areaTextState.isAllAreaText = false;
        }
        return areaTextState;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したテキストに線と吹き出しのしっぽを描く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        var bubbleTargets = collectBubbleTargets(doc);
        if (!bubbleTargets.length) {
            alert(getLabel(LABELS.alert.noText));
            return;
        }

        /* 実行中は［境界線を隠す］と［バウンディングボックスを隠す］を切り替えて、プレビューを見やすくする。
           どちらも反転しかできないので、終了時にもう一度切り替えて元に戻す
           Toggle Hide Edges and Hide Bounding Box while running so the preview is easier to see;
           both only toggle, so toggle them again on the way out to restore them */
        app.executeMenuCommand("edge");
        app.executeMenuCommand("AI Bounding Box Toggle");
        try {
            runSpeechBubbleMaker(doc, bubbleTargets);
        } finally {
            app.executeMenuCommand("AI Bounding Box Toggle");
            app.executeMenuCommand("edge");
        }
    }

    /**
     * ダイアログを出し、OK なら線としっぽを描いて、描いたものを選択する
     * @param {Document} doc - 対象のドキュメント
     * @param {BubbleTarget[]} bubbleTargets - 対象
     * @returns {void}
     */
    function runSpeechBubbleMaker(doc, bubbleTargets) {
        var rulerUnit = getUnitInfo();
        var strokeUnit = getUnitInfo("strokeUnits");
        var settingsStore = createSettingsStore(SCRIPT_NAME, "session");
        var initialSettings = settingsStore.load(buildDefaultSettings(bubbleTargets[0].fontSize, rulerUnit, strokeUnit));

        var bubbleSettings = showSpeechBubbleDialog(bubbleTargets, initialSettings, rulerUnit, strokeUnit);
        if (!bubbleSettings) return;
        settingsStore.save(bubbleSettings);

        var drawResult = drawBubbles(bubbleTargets, bubbleSettings);
        var isAreaTextCentered = true;
        if (bubbleSettings.frameKey === "box" && bubbleSettings.convertToAreaText) isAreaTextCentered = centerTextsInBoxes(drawResult.drawn);

        var resultItems = [];
        for (var i = 0; i < drawResult.drawn.length; i++) {
            var drawnBubble = drawResult.drawn[i];
            resultItems.push(bubbleSettings.groupWithText ? groupTailWithText(drawnBubble.group, drawnBubble.target.frame) : drawnBubble.group);
        }
        doc.selection = null;
        if (resultItems.length) doc.selection = resultItems;
        if (!isAreaTextCentered) alert(getLabel(LABELS.alert.areaTextFailed));
        if (drawResult.skippedCount > 0) alert(getLabel(LABELS.alert.skipped, [drawResult.skippedCount]));
    }

    main();

})();
