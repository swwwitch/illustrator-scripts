#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SwapNearestItemPalette
Meta-Default-language=ja
Name=オブジェクトを入れ替え（常駐パレット版）
Name[en]=Swap objects (floating palette)
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2025-07-06
Tags=alignment,整列
X-com.swwwitch-Script-Name=SwapNearestItemPalette
X-com.swwwitch-Updated=2026-11-13
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItemPalette.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItemPalette.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SwapNearestItemPalette.jsx
Description-BEGIN
[ja]
選択中のオブジェクトと、指定した方向（上下左右）にある最も近いオブジェクトの位置を入れ替えます。常駐パレットを開いたまま、矢印キーで方向を指定します。
[en]
Swaps the selected object with the nearest object in a chosen direction, picked with the arrow keys while a floating palette stays open.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Basic settings
    // =========================================

    /* 矢印キーと探索方向の対応 / Arrow key to search direction */
    var DIRECTION_BY_ARROW_KEY = {
        Up: "up",
        Down: "down",
        Left: "left",
        Right: "right"
    };

    /* 入れ替えの対象にできる typename / Typenames that can take part in a swap */
    var SWAPPABLE_TYPENAMES = [
        "PathItem", "CompoundPathItem", "GroupItem", "TextFrame",
        "PlacedItem", "RasterItem", "SymbolItem", "MeshItem",
        "PluginItem", "GraphItem"
    ];

    /* 共通エンジンの $.global に載せるパレットの参照名（CloseAllPalettes と同じ名前）
       Palette reference name on the shared engine's $.global (the same name as in CloseAllPalettes) */
    var PALETTE_GLOBAL_KEY = "__swapNearestItemPalette";

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ▲▼◀▶ ボタン / Arrow buttons */
    var ARROW_BUTTON_SIZE = [30, 24];   /* ボタン1つの大きさ [幅, 高さ]（px）/ size of one button [width, height] (px) */
    var ARROW_BUTTON_GAP = 2;           /* ボタンどうしの間隔（px）/ gap between buttons (px) */
    var ARROW_TRIANGLE_WIDTH = 10;      /* 三角形の底辺（px）/ base of the triangle (px) */
    var ARROW_TRIANGLE_DEPTH = 6;       /* 三角形の高さ（px）/ height of the triangle (px) */

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
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
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
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPTMETA の Version）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
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
        var versionText = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : "";
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトの位置を入れ替え", en: "Swap Positions" }
        },
        status: {
            prompt: { ja: "矢印キーで入れ替え", en: "Press an arrow key to swap" },
            notFound: { ja: "見つかりません", en: "Nothing found" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            prompt: {
                ja: "矢印キーを押すと、その方向にある最も近いオブジェクトと位置を入れ替えます。Esc または Enter で閉じます。",
                en: "An arrow key swaps the selection with the nearest object in that direction. Esc or Enter closes this palette."
            },
            arrow: {
                up: { ja: "上の最も近いオブジェクトと入れ替え", en: "Swap with the nearest object above" },
                down: { ja: "下の最も近いオブジェクトと入れ替え", en: "Swap with the nearest object below" },
                left: { ja: "左の最も近いオブジェクトと入れ替え", en: "Swap with the nearest object to the left" },
                right: { ja: "右の最も近いオブジェクトと入れ替え", en: "Swap with the nearest object to the right" }
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOne: { ja: "1つのオブジェクトを選択してください。", en: "Select one object." },
            noPosition: { ja: "位置情報を持つオブジェクトを選択してください。", en: "Select an object that has a position." }
        }
    };

    // =========================================
    // キーボードショートカット / Keyboard shortcuts
    // =========================================

    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

    // =========================================
    // 判定 / Checks
    // =========================================

    /**
     * 入れ替えの対象にできるオブジェクトか判定する（起動時の確認用。パレットからの入れ替えはメインエンジン側で判定する）
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象にできる場合はtrue
     */
    function isSwappableItem(pageItem) {
        /* ガイドは typename が "PathItem" のため guides で判定する / Guides are PathItems, so test .guides */
        if (pageItem.typename === "PathItem" && pageItem.guides === true) return false;

        for (var i = 0; i < SWAPPABLE_TYPENAMES.length; i++) {
            if (SWAPPABLE_TYPENAMES[i] === pageItem.typename) return true;
        }
        return false;
    }

    // =========================================
    // メインエンジンでの入れ替え（BridgeTalk で送る） / Swap in the main engine (sent through BridgeTalk)
    // =========================================

    /* 常駐パレットからは app.activeDocument に触れないので、探索と入れ替えは BridgeTalk でメインエンジンに任せる
       A persistent palette cannot touch app.activeDocument, so the search and swap run in the main engine through BridgeTalk */

    /* ワーカーの toString() を切り詰める目印 / Marker where the worker's toString() is cut */
    var SWAP_WORKER_END_MARKER = "/* @@swap-worker-end@@ */";

    /* メインエンジンの $.global に登録するワーカーの名前 / Name of the worker registered on the main engine's $.global */
    var SWAP_WORKER_GLOBAL_KEY = "__swapNearestItemWorker";

    /* swapNearestWorker は toString() で文字列にして送るので、JSDoc もコメントも書かない（中は ASCII だけ、// も使わない）。
       末尾の目印 @@swap-worker-end@@ で切り詰めて送る（消さないこと）。
       戻り値は "swapped"（入れ替えた）/ "notFound"（相手が無い）/ "noSelection"（1つ選択していない）/ "noDocument" / "error"
       swapNearestWorker is sent as toString() text: no JSDoc or comments inside (ASCII only, no //).
       It is cut at the end marker @@swap-worker-end@@ (keep it).
       Returns "swapped" / "notFound" / "noSelection" / "noDocument" / "error" */
    function swapNearestWorker(searchDirection, swappableTypeList) {
        var swappableTypes = {};
        var swappableTypeNames = swappableTypeList.split(",");
        for (var t = 0; t < swappableTypeNames.length; t++) swappableTypes[swappableTypeNames[t]] = true;

        function isSwappableItem(pageItem) {
            if (pageItem.typename === "PathItem" && pageItem.guides === true) return false;
            return swappableTypes[pageItem.typename] === true;
        }

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
            } catch (e) { }
            return false;
        }

        function getClipMaskItem(groupItem) {
            try {
                if (groupItem.clipped !== true) return null;
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

        function hasClippedDescendant(groupItem) {
            try {
                var groupChildren = groupItem.pageItems;
                for (var i = 0; i < groupChildren.length; i++) {
                    if (groupChildren[i].typename !== "GroupItem") continue;
                    if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
                }
            } catch (e) { }
            return false;
        }

        function getClipAwareBounds(item) {
            try {
                var measuredItem = item;
                if (item.typename === "GroupItem") {
                    var maskItem = getClipMaskItem(item);
                    if (maskItem) {
                        measuredItem = maskItem;
                    } else if (hasClippedDescendant(item)) {
                        var childBounds = getChildUnionBounds(item.pageItems);
                        if (childBounds) return childBounds;
                    }
                }
                var bounds = measuredItem.visibleBounds;
                return [bounds[0], bounds[1], bounds[2], bounds[3]];
            } catch (e) {
                return null;
            }
        }

        function getChildUnionBounds(childItems) {
            var unionBounds = null;
            for (var i = 0; i < childItems.length; i++) {
                var childItem = childItems[i];
                try {
                    if (childItem.hidden === true || childItem.guides === true) continue;
                } catch (e) { }
                var itemBounds = getClipAwareBounds(childItem);
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

        function getBoundsCenter(bounds) {
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        }

        function collectLayerItems(layerList, collectedItems) {
            for (var i = 0; i < layerList.length; i++) {
                var layer = layerList[i];
                if (layer.locked || layer.visible === false) continue;
                var layerItems = layer.pageItems;
                for (var j = 0; j < layerItems.length; j++) collectedItems.push(layerItems[j]);
                if (layer.layers.length > 0) collectLayerItems(layer.layers, collectedItems);
            }
            return collectedItems;
        }

        function containsItem(candidateItems, targetItem) {
            for (var i = 0; i < candidateItems.length; i++) {
                if (candidateItems[i] === targetItem) return true;
            }
            return false;
        }

        function isSearchCandidate(candidateItem, referenceItem, excludeItems) {
            if (candidateItem === referenceItem) return false;
            var parentType = candidateItem.parent.typename;
            if (parentType === "GroupItem" || parentType === "CompoundPathItem") return false;
            if (candidateItem.locked || candidateItem.hidden) return false;
            if (!isSwappableItem(candidateItem)) return false;
            if (containsItem(excludeItems, candidateItem)) return false;
            return true;
        }

        function getDirectionalGap(referenceBounds, candidateBounds, dx, dy) {
            var verticalOverlap = !(referenceBounds[3] > candidateBounds[1] || referenceBounds[1] < candidateBounds[3]);
            var horizontalOverlap = !(referenceBounds[0] > candidateBounds[2] || referenceBounds[2] < candidateBounds[0]);
            var gap;
            if (searchDirection === "right" && dx > 0 && verticalOverlap) {
                gap = candidateBounds[0] - referenceBounds[2];
            } else if (searchDirection === "left" && dx < 0 && verticalOverlap) {
                gap = referenceBounds[0] - candidateBounds[2];
            } else if (searchDirection === "up" && dy > 0 && horizontalOverlap) {
                gap = candidateBounds[3] - referenceBounds[1];
            } else if (searchDirection === "down" && dy < 0 && horizontalOverlap) {
                gap = referenceBounds[3] - candidateBounds[1];
            } else {
                return null;
            }
            return (gap < 0) ? 0 : gap;
        }

        function findNearestObjectInDirection(referenceItem, excludeItems) {
            var documentItems = collectLayerItems(app.activeDocument.layers, []);
            var referenceBounds = getClipAwareBounds(referenceItem);
            if (!referenceBounds) return null;
            var referenceCenter = getBoundsCenter(referenceBounds);
            var nearestItem = null;
            var minGap = Number.MAX_VALUE;
            var minDistance = Number.MAX_VALUE;
            for (var i = 0; i < documentItems.length; i++) {
                var candidateItem = documentItems[i];
                if (!isSearchCandidate(candidateItem, referenceItem, excludeItems)) continue;
                var candidateBounds = getClipAwareBounds(candidateItem);
                if (!candidateBounds) continue;
                var candidateCenter = getBoundsCenter(candidateBounds);
                var dx = candidateCenter[0] - referenceCenter[0];
                var dy = candidateCenter[1] - referenceCenter[1];
                var gap = getDirectionalGap(referenceBounds, candidateBounds, dx, dy);
                if (gap === null) continue;
                var distance = Math.sqrt(dx * dx + dy * dy);
                if (gap < minGap || (gap === minGap && distance < minDistance)) {
                    minGap = gap;
                    minDistance = distance;
                    nearestItem = candidateItem;
                }
            }
            return nearestItem;
        }

        function swapHorizontally(itemA, itemB) {
            var boundsA = getClipAwareBounds(itemA);
            var boundsB = getClipAwareBounds(itemB);
            if (boundsA[0] > boundsB[0]) {
                var tempItem = itemA;
                itemA = itemB;
                itemB = tempItem;
                var tempBounds = boundsA;
                boundsA = boundsB;
                boundsB = tempBounds;
            }
            var widthB = boundsB[2] - boundsB[0];
            var gap = boundsB[0] - boundsA[2];
            itemB.translate(boundsA[0] - boundsB[0], 0);
            itemA.translate(widthB + gap, 0);
        }

        function swapVertically(itemA, itemB) {
            var boundsA = getClipAwareBounds(itemA);
            var boundsB = getClipAwareBounds(itemB);
            if (boundsA[1] < boundsB[1]) {
                var tempItem = itemA;
                itemA = itemB;
                itemB = tempItem;
                var tempBounds = boundsA;
                boundsA = boundsB;
                boundsB = tempBounds;
            }
            var heightB = boundsB[1] - boundsB[3];
            var gap = boundsA[3] - boundsB[1];
            itemB.translate(0, boundsA[1] - boundsB[1]);
            itemA.translate(0, -(heightB + gap));
        }

        try {
            if (app.documents.length === 0) return "noDocument";
            var selectedObjects = app.activeDocument.selection;
            if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length !== 1) return "noSelection";
            var baseItem = selectedObjects[0];
            if (!isSwappableItem(baseItem)) return "noSelection";
            var nearestItem = findNearestObjectInDirection(baseItem, selectedObjects);
            if (!nearestItem) return "notFound";
            if (searchDirection === "up" || searchDirection === "down") {
                swapVertically(baseItem, nearestItem);
            } else {
                swapHorizontally(baseItem, nearestItem);
            }
            return "swapped";
        } catch (e) {
            return "error";
        }
        /* @@swap-worker-end@@ */
    }

    /**
     * ワーカー関数のソースを返す。toString() は前後のコメントまで取り込むことがあるので、
     * 関数の宣言から終わりの目印までを切り出し、閉じ括弧を足す
     * @returns {string} 送信する関数のソース
     */
    function getSwapWorkerSource() {
        var workerSource = swapNearestWorker.toString();
        var startIndex = workerSource.indexOf("function swapNearestWorker");
        if (startIndex > 0) workerSource = workerSource.substring(startIndex);
        var endIndex = workerSource.indexOf(SWAP_WORKER_END_MARKER);
        if (endIndex < 0) return workerSource;
        return workerSource.substring(0, endIndex) + "}";
    }

    /**
     * メインエンジンで評価する本文を組み立てる。最初の1回はワーカーを $.global に登録してから呼び、
     * 2回目からは登録済みのワーカーを短い呼び出しで使う（毎回ソースを送って評価し直さない）。
     * 登録が消えていたら "missingWorker" を返す
     * @param {string|null} workerSource - 登録するときは getSwapWorkerSource() の結果、登録済みなら null
     * @param {string} searchDirection - 探索方向（"right" / "left" / "up" / "down"）
     * @returns {string} BridgeTalk の本文（評価した値＝ワーカーの戻り値が結果になる）
     */
    function buildSwapRequestBody(workerSource, searchDirection) {
        /* 本文に \ を入れない（BridgeTalk の転送で二重になる）/ Keep backslashes out of the body (BridgeTalk doubles them) */
        var workerRef = "$.global." + SWAP_WORKER_GLOBAL_KEY;
        var workerCall = workerRef + "(\"" + searchDirection + "\", \"" + SWAPPABLE_TYPENAMES.join(",") + "\")";
        if (workerSource) return workerRef + " = (" + workerSource + "); " + workerCall + ";";
        return "(" + workerRef + " ? " + workerCall + " : \"missingWorker\");";
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

    // =========================================
    // ▲▼◀▶ ボタン / Arrow buttons
    // =========================================

    /* 文字の ▲▼ は上下で大きさやベースラインがそろわないので、onDraw で三角形を描く
       Glyph triangles do not line up in size or baseline, so the triangles are drawn in onDraw */
    var ARROW_UI_DARK = isDarkUI();
    var ARROW_FILL_COLOR     = ARROW_UI_DARK ? [1, 1, 1, 0.06] : [1, 1, 1, 0.60];  /* 地 / background */
    var ARROW_FRAME_COLOR    = ARROW_UI_DARK ? [1, 1, 1, 0.15] : [0, 0, 0, 0.20];  /* 枠線 / frame */
    var ARROW_PRESSED_COLOR  = ARROW_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var ARROW_TRIANGLE_COLOR = ARROW_UI_DARK ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70];  /* 三角形 / triangle */

    /**
     * 十字に並べた ▲▼◀▶ ボタンを足す（上段に▲、下段に◀▼▶）
     * @param {Group} parentGroup - 足す先
     * @param {function(string): void} onArrowPress - 押したときに探索方向（"up" / "down" / "left" / "right"）を受ける処理
     * @returns {void}
     */
    function addArrowPad(parentGroup, onArrowPress) {
        parentGroup.orientation = "column";
        parentGroup.alignChildren = ["center", "center"];
        parentGroup.spacing = ARROW_BUTTON_GAP;

        var topRow = parentGroup.add("group");
        setupRow(topRow, "center", ARROW_BUTTON_GAP);
        addArrowButton(topRow, "up", onArrowPress);

        var bottomRow = parentGroup.add("group");
        setupRow(bottomRow, "center", ARROW_BUTTON_GAP);
        addArrowButton(bottomRow, "left", onArrowPress);
        addArrowButton(bottomRow, "down", onArrowPress);
        addArrowButton(bottomRow, "right", onArrowPress);
    }

    /**
     * 三角形を描いたボタンを1つ足す（group を onDraw で描き、押した瞬間に処理を呼ぶ）
     * @param {Group} parentGroup - 足す先
     * @param {string} searchDirection - 向き（"up" / "down" / "left" / "right"）
     * @param {function(string): void} onArrowPress - 押したときの処理
     * @returns {Group} ボタンとして使う group
     */
    function addArrowButton(parentGroup, searchDirection, onArrowPress) {
        var arrowButton = parentGroup.add("group");
        arrowButton.preferredSize = [ARROW_BUTTON_SIZE[0], ARROW_BUTTON_SIZE[1]];
        arrowButton.minimumSize = [ARROW_BUTTON_SIZE[0], ARROW_BUTTON_SIZE[1]];
        arrowButton.maximumSize = [ARROW_BUTTON_SIZE[0], ARROW_BUTTON_SIZE[1]];
        arrowButton.helpTip = getLabel("tooltip.arrow." + searchDirection);
        arrowButton.isPressed = false;

        arrowButton.onDraw = function () {
            var buttonGraphics = arrowButton.graphics;
            var buttonWidth = ARROW_BUTTON_SIZE[0];
            var buttonHeight = ARROW_BUTTON_SIZE[1];
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonWidth - 1, buttonHeight - 1);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR,
                arrowButton.isPressed ? ARROW_PRESSED_COLOR : ARROW_FILL_COLOR));
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, ARROW_FRAME_COLOR, 1));
            drawArrowTriangle(buttonGraphics, buttonWidth, buttonHeight, searchDirection);
        };

        /**
         * 押下状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaintArrowButton(isPressed) {
            if (arrowButton.isPressed === isPressed) return;
            arrowButton.isPressed = isPressed;
            arrowButton.hide();
            arrowButton.show();
        }

        arrowButton.addEventListener("mousedown", function () {
            repaintArrowButton(true);
            onArrowPress(searchDirection);
        });
        arrowButton.addEventListener("mouseup", function () { repaintArrowButton(false); });
        /* 押したまま外へ出たときも押下色を残さない / Reset when the pointer leaves while pressed */
        arrowButton.addEventListener("mouseout", function () { repaintArrowButton(false); });
        return arrowButton;
    }

    /**
     * ボタンの中央に向きどおりの三角形を塗る
     * @param {ScriptUIGraphics} buttonGraphics - 描く先
     * @param {number} buttonWidth - ボタンの幅
     * @param {number} buttonHeight - ボタンの高さ
     * @param {string} searchDirection - 向き（"up" / "down" / "left" / "right"）
     * @returns {void}
     */
    function drawArrowTriangle(buttonGraphics, buttonWidth, buttonHeight, searchDirection) {
        var centerX = buttonWidth / 2;
        var centerY = buttonHeight / 2;
        var halfBase = ARROW_TRIANGLE_WIDTH / 2;
        var halfDepth = ARROW_TRIANGLE_DEPTH / 2;
        /* 頂点：底辺の両端と先端 / Vertices: both ends of the base, then the tip */
        var trianglePoints = {
            up: [[centerX - halfBase, centerY + halfDepth], [centerX + halfBase, centerY + halfDepth], [centerX, centerY - halfDepth]],
            down: [[centerX - halfBase, centerY - halfDepth], [centerX + halfBase, centerY - halfDepth], [centerX, centerY + halfDepth]],
            left: [[centerX + halfDepth, centerY - halfBase], [centerX + halfDepth, centerY + halfBase], [centerX - halfDepth, centerY]],
            right: [[centerX - halfDepth, centerY - halfBase], [centerX - halfDepth, centerY + halfBase], [centerX + halfDepth, centerY]]
        }[searchDirection];
        buttonGraphics.newPath();
        buttonGraphics.moveTo(trianglePoints[0][0], trianglePoints[0][1]);
        buttonGraphics.lineTo(trianglePoints[1][0], trianglePoints[1][1]);
        buttonGraphics.lineTo(trianglePoints[2][0], trianglePoints[2][1]);
        buttonGraphics.closePath();
        buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, ARROW_TRIANGLE_COLOR));
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * 矢印キーで入れ替えを実行する常駐パレットを表示する
     * @returns {void}
     */
    function showPalette() {
        var swapPalette = new Window("palette", getLabel("dialog.title"));
        setupWindow(swapPalette);

        var statusText = swapPalette.add("statictext", undefined, getLabel("status.prompt"));
        statusText.alignment = "center";
        statusText.helpTip = getLabel("tooltip.prompt");

        var arrowPadGroup = swapPalette.add("group");
        arrowPadGroup.alignment = "center";

        var buttonRow = addButtonRow(swapPalette);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("button.close"));
        btnClose.onClick = function () {
            swapPalette.close();
        };
        alignRightOnlyButtonRow(buttonRow);

        var swapWorkerSource = getSwapWorkerSource();
        var isWorkerInstalled = false; /* パレットを開くたびに登録し直す（修正後のワーカーを使う）/ Re-register on every open so edits take effect */
        var pendingSwapJob = null;     /* 返事待ちの BridgeTalk（届くまで GC されないよう持っておく）/ BridgeTalk awaiting its reply (kept so it is not collected) */
        var isArrowKeyHeld = false;    /* キーリピートで続けて入れ替えないための門 / Gate out key repeats while a key is held */

        /**
         * メインエンジンへ入れ替えを頼み、結果をステータスに出す
         * @param {string} searchDirection - 探索方向（"right" / "left" / "up" / "down"）
         * @returns {void}
         */
        function requestSwap(searchDirection) {
            var swapJob = new BridgeTalk();
            swapJob.target = BridgeTalk.appSpecifier; /* 自分のバージョンのメインエンジン / this version's main engine */
            swapJob.body = buildSwapRequestBody(isWorkerInstalled ? null : swapWorkerSource, searchDirection);
            swapJob.onResult = function (resultMessage) {
                pendingSwapJob = null;
                if (resultMessage.body === "missingWorker") {
                    /* メインエンジン側の登録が消えていたら、登録し直して送り直す / Re-register and resend when the worker is gone */
                    isWorkerInstalled = false;
                    requestSwap(searchDirection);
                    return;
                }
                isWorkerInstalled = true;
                /* 見つからないときはメッセージを差し替えて知らせる。連打を邪魔しないようアラートは出さない
                   Report a miss by swapping the message; an alert would interrupt repeated key presses */
                statusText.text = (resultMessage.body === "swapped") ? getLabel("status.prompt") : getLabel("status.notFound");
            };
            swapJob.onError = function (errorMessage) {
                pendingSwapJob = null;
                statusText.text = getLabel("status.notFound");
                /* 常駐パレットなのでアラートは出さず、デバッグ用に残す / No alert in a persistent palette; log for debugging */
                $.writeln(SCRIPT_NAME + " BridgeTalk error: " + errorMessage.body);
            };
            pendingSwapJob = swapJob;
            swapJob.send();
        }

        /**
         * 矢印キーを押したときの処理（押しっぱなしと返事待ちのあいだは受け付けない）
         * @param {string} searchDirection - 探索方向
         * @returns {void}
         */
        function handleArrowKey(searchDirection) {
            if (isArrowKeyHeld || pendingSwapJob) return;
            isArrowKeyHeld = true;
            requestSwap(searchDirection);
        }

        /* ▲▼◀▶ ボタン（返事待ちのあいだは受け付けない）/ Arrow buttons (ignored while a reply is pending) */
        addArrowPad(arrowPadGroup, function (searchDirection) {
            if (pendingSwapJob) return;
            requestSwap(searchDirection);
        });

        /**
         * 矢印キー1つ分のショートカットの処理を作る（ループの変数を閉じ込める）
         * @param {string} searchDirection - 探索方向
         * @returns {Function} addKeyShortcuts に渡す関数
         */
        function createArrowKeyHandler(searchDirection) {
            return function () {
                handleArrowKey(searchDirection);
            };
        }

        /**
         * ［閉じる］と同じ処理（パレットは Esc で閉じないので割り当てる）
         * @returns {void}
         */
        function closeFromKey() {
            btnClose.onClick();
        }

        var shortcutMap = {
            "Escape": { target: closeFromKey, inFields: true },
            "Enter": { target: closeFromKey, inFields: true },
            "Return": { target: closeFromKey, inFields: true }
        };
        for (var arrowKeyName in DIRECTION_BY_ARROW_KEY) {
            if (!DIRECTION_BY_ARROW_KEY.hasOwnProperty(arrowKeyName)) continue;
            shortcutMap[arrowKeyName] = createArrowKeyHandler(DIRECTION_BY_ARROW_KEY[arrowKeyName]);
        }
        addKeyShortcuts(swapPalette, shortcutMap);

        swapPalette.addEventListener("keyup", function () {
            isArrowKeyHeld = false;
        }, false);

        /* 閉じたら参照を外す（DOM には触れない）/ Drop the reference on close (never touch the DOM here) */
        swapPalette.onClose = function () {
            $.global[PALETTE_GLOBAL_KEY] = null;
        };

        btnClose.active = true;
        $.global[PALETTE_GLOBAL_KEY] = swapPalette;
        addVersionFooter(swapPalette);
        prepareDialogWindow(swapPalette, SCRIPT_NAME);
        swapPalette.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択内容を確認してパレットを表示する（開いていれば前面に出す）
     * @returns {void}
     */
    function main() {
        var openPalette = $.global[PALETTE_GLOBAL_KEY];
        if (openPalette) {
            try {
                openPalette.show();
                return;
            } catch (e) {
                /* 閉じたあとの参照が残っていたら作り直す / Rebuild when a stale reference is left */
                $.global[PALETTE_GLOBAL_KEY] = null;
            }
        }

        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedObjects = doc.selection;

        /* 文字カーソルが立っていると selection は TextRange で、length は文字数になる
           A text caret gives a TextRange whose length counts characters */
        if (!selectedObjects || selectedObjects.typename === "TextRange" || selectedObjects.length !== 1) {
            alert(getLabel("alert.selectOne"));
            return;
        }

        if (!isSwappableItem(selectedObjects[0])) {
            alert(getLabel("alert.noPosition"));
            return;
        }

        showPalette();
    }

    main();

})();
