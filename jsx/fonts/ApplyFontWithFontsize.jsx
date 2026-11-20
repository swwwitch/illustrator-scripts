#target illustrator
#targetengine "ApplyFontWithFontsizeEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ApplyFontWithFontsize
Name=行のテキストをフォント名（＋サイズ・行送り）とみなして行単位で適用
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.15
Release-Date=2026-06-06
Description-BEGIN
選択したテキストフレームの各行（段落）を「フォント名（＋サイズ・行送り）」の指定とみなし、行単位でフォントを適用します。
「ヒラギノ角ゴシック W3 12pt↓16pt」のようにサイズ・行送りを併記でき、併記のない行はフォントだけを適用します。

Reads each line (paragraph) of the selected text frames as a "font name (plus size and leading)" spec and applies it line by line.
Sizes and leading can be written after the name, as in "Hiragino Kaku Gothic W3 12pt↓16pt"; a line without them only changes the font.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したテキストフレームの各行（段落）を「フォント名（＋サイズ・行送り）」の指定とみなし、行単位でフォントを適用します。
「ヒラギノ角ゴシック W3 12pt↓16pt」のようにサイズ・行送りを併記でき、併記のない行はフォントだけを適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyFontWithFontsize.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n33d152e73f35

### Overview

Reads each line (paragraph) of the selected text frames as a "font name (plus size and leading)" spec and applies it line by line.
Sizes and leading can be written after the name, as in "Hiragino Kaku Gothic W3 12pt↓16pt"; a line without them only changes the font.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyFontWithFontsize.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyFontWithFontsize";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.15";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "行のテキストをフォント名（＋サイズ・行送り）とみなして行単位で適用"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyFontWithFontsize.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyFontWithFontsize.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n33d152e73f35"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/ApplyFontWithFontsize.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 特定の文字列に強制的に割り当てるフォント（照合前に置き換える）/ Forced font per line text */
    var CUSTOM_FONT_MAP = {
        "Jenson": "Adobe Jenson Pro",
        "Garamond": "Adobe Garamond Pro",
        "Myriad": "Myriad Pro",
        "Frutiger": "Neue Frutiger World",
        "FF DIN": "DIN 2014",
        "Minion": "Minion Pro"
    };

    /* 複数スタイルが候補になったときの優先順位（先頭ほど優先。小文字で比較）/ Style preference order */
    var STYLE_PRIORITY = ["bold", "semibold", "medium", "regular"];

    var MARKER_LAYER_NAME = "// missing-fonts";  /* 未適用の目印を置くレイヤー名 / marker layer name */
    var MARKER_OPACITY    = 35;                  /* 目印の不透明度（％）/ marker opacity in percent */

    var ZOOM_FIT_RATIO = 0.6;   /* ピッカー表示時、対象が表示領域を占める割合 / zoom fit ratio */
    var ZOOM_MIN       = 0.03;  /* Illustrator のズーム下限（3%）/ minimum zoom of Illustrator */
    var ZOOM_MAX       = 64;    /* Illustrator のズーム上限（6400%）/ maximum zoom of Illustrator */

    // =========================================
    // レイアウト設定 / Layout settings
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

    var PROGRESS_BAR_SIZE   = [320, 9];       /* プログレスバーの寸法 / size of the progress bar */
    var PROGRESS_TEXT_WIDTH = 320;            /* 進捗表示の幅 / width of the progress count label */
    var RESULT_FIELD_SIZE   = [380, 220];     /* 未適用一覧の寸法 / size of the unapplied list field */
    var PICKER_LABEL_WIDTH  = 70;             /* ピッカーの項目名の幅 / width of the picker row labels */
    var PICKER_FIELD_WIDTH  = 200;            /* ピッカーの入力欄・プルダウンの幅 / width of the picker fields */
    var PICKER_TARGET_WIDTH = 260;            /* 対象テキスト表示の幅 / width of the target text label */

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
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
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

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。基本情報ブロックの SCRIPT_* を読み、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 基本情報ブロックの値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値
     */
    function getVersionInfoValue(variableKey) {
        var infoValue;
        if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
        else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
        else if (variableKey === "readme") {
            if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
            else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
        }
        else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
        else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを Finder で表示する。RevealInFinder.app があればファイルを選択した状態で表示し、無ければ入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        var hasRevealApp = (File.fs === "Macintosh")
            && (new Folder(VERSION_INFO_REVEAL_APP_PATH).exists || new File(VERSION_INFO_REVEAL_APP_PATH).exists);
        /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
        if (hasRevealApp && writeVersionInfoRevealPath(scriptFile.fsName) && new File(VERSION_INFO_REVEAL_APP_PATH).execute()) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * RevealInFinder.app に渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(targetPath) {
        var pathFile = new File(VERSION_INFO_REVEAL_PATH_FILE);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
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
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noSelection: {
                ja: "テキストオブジェクトを選択してください。",
                en: "Please select a text object."
            }
        },
        dialog: {
            progressTitle: {
                ja: "フォントを適用中...",
                en: "Applying fonts..."
            },
            resultTitle: {
                ja: "適用結果",
                en: "Apply results"
            },
            tipResultList: {
                ja: "フォントを適用できなかったテキストの一覧です。選んでコピーできます。",
                en: "The text objects the font could not be applied to. You can select and copy them."
            },
            resultHeader: {
                ja: "フォントを適用できなかった文字列",
                en: "Strings with no matching font"
            },
            pickerTitle: {
                ja: "フォントを選択",
                en: "Choose font"
            }
        },
        panel: {
            applyFont: {
                ja: "適用するフォント",
                en: "Font to apply"
            }
        },
        fieldLabel: {
            targetText: {
                ja: "対象テキスト",
                en: "Target text"
            },
            search: {
                ja: "検索",
                en: "Search"
            },
            family: {
                ja: "フォント",
                en: "Font"
            },
            style: {
                ja: "スタイル",
                en: "Style"
            }
        },
        button: {
            copy: {
                ja: "クリップボードにコピー",
                en: "Copy to clipboard"
            },
            close: {
                ja: "閉じる",
                en: "Close"
            },
            apply: {
                ja: "適用",
                en: "Apply"
            },
            skip: {
                ja: "スキップ",
                en: "Skip"
            },
            quit: {
                ja: "終了",
                en: "Quit"
            }
        }
    };

    /* showFontPicker が「終了」で返す番兵（選択ループを打ち切る合図）/ Sentinel returned on Quit */
    var PICKER_QUIT = {};

    /* showFontPicker が「適用したが失敗した」ときに返す番兵（スキップと区別する）/ Sentinel for a failed apply */
    var PICKER_FAILED = {};

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

    /**
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

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
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    var fontIndex = null; /* 全インストールフォントの索引（選択チェック後に作成）/ index of installed fonts */

    var selectedItems = doc.selection;
    if (!selectedItems || selectedItems.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    /* 選択物からテキストフレームを再帰収集（グループ内も対象。ロック・非表示はグループごと除外。
       テキスト編集中の選択は対象にしない）/ Collect text frames recursively, skipping locked or hidden items
       with their contents; a text-editing selection is not a target */
    var targetTextFrames = collectSelectionTextFrames(selectedItems, {
        skipLocked: true,
        skipHidden: true,
        textRangeToFrame: false
    });

    if (targetTextFrames.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    /* 対象が確定してから索引化する（重い処理なので選択チェックの後）/ Build the index once the target is fixed */
    fontIndex = createFontIndex();

    /* フェーズ1：厳密一致だけ自動適用し、判断が要る行は保留にする / Phase 1: apply confident matches */
    var pendingLines = autoApplyFonts(targetTextFrames);

    /* フェーズ2：保留分を対話ピッカーで1件ずつ決める / Phase 2: resolve the queue interactively */
    var previousView = captureView(); /* ピッカーでズームするので現在の表示を控える / Keep the current view */
    var unapplied = resolvePendingLines(pendingLines);
    restoreView(previousView);

    /* 未適用の行を含むフレームに目印を置く / Mark the frames that still have unapplied lines */
    markUnappliedFrames(unapplied.frames);

    /* 適用できなかった文字列を一覧表示し、要求があればクリップボードへコピー / Show and optionally copy */
    if (unapplied.texts.length > 0 && showUnappliedDialog(unapplied.texts)) {
        copyTextToClipboard(unapplied.texts.join("\n"));
    }

    // =========================================
    // 処理フロー / Processing flow
    // =========================================

    /**
     * フェーズ1：全フレームを走査し、厳密一致のフォントだけを自動適用する
     * 判断が要る行（あいまい一致・未一致・適用失敗）は保留リストに積んで返す
     * @param {TextFrame[]} textFrameList - 対象のテキストフレーム
     * @returns {Object[]} 保留行（{ frame, index, lineText, fontName, fontSize, fontLeading, initialFont }）
     */
    function autoApplyFonts(textFrameList) {
        var pendingLines = [];
        var progress = createProgressPalette(textFrameList.length);

        /* 途中で例外が出てもパレットを残さない / Never leave the palette open on an error */
        try {
            for (var i = 0; i < textFrameList.length; i++) {
                progress.update(i + 1);
                autoApplyLinesInFrame(textFrameList[i], pendingLines);
            }
        } finally {
            progress.window.close();
        }

        return pendingLines;
    }

    /**
     * 1つのテキストフレームの各行を照合し、厳密一致は適用、それ以外は保留に積む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} pendingLines - 保留行の追加先
     * @returns {void}
     */
    function autoApplyLinesInFrame(textFrame, pendingLines) {
        /* 段落コレクションはキャッシュしない。適用のたびに参照が無効化され Error 1302 になるため
           / Fetch paragraphs live: applying a font invalidates a cached collection */
        var paragraphCount = textFrame.paragraphs.length;

        /* 下から上へたどると、適用によるインデックスのズレを避けられる / Loop upwards to keep indexes valid */
        for (var j = paragraphCount - 1; j >= 0; j--) {
            var lineText = textFrame.paragraphs[j].contents.replace(/^[\s　]+|[\s　]+$/g, "");
            if (lineText === "") continue; /* 空行・空白だけの行は対象外 / Skip blank lines */

            /* 行末に併記されたサイズ・行送りを切り出す（照合は名前部分だけで行う）/ Split off size and leading */
            var fontSpec = parseFontSpec(lineText);
            var fontMatch = findFont(fontSpec.name);

            var isApplied = fontMatch.confident
                && applyFontSpecToLine(textFrame, j, fontMatch.font, fontSpec.size, fontSpec.leading);

            if (!isApplied) {
                pendingLines.push({
                    frame: textFrame,
                    index: j,
                    lineText: lineText,
                    fontName: fontSpec.name,
                    fontSize: fontSpec.size,
                    fontLeading: fontSpec.leading,
                    initialFont: fontMatch.font
                });
            }
        }
    }

    /**
     * フェーズ2：保留行を対話ピッカーで1件ずつ決める
     * 同じフォント名は実行中に一度決めたら再質問せず、その結果を使い回す
     * @param {Object[]} pendingLines - 保留行
     * @returns {Object} 未適用の記録（{ frames: TextFrame[], texts: string[] }）
     */
    function resolvePendingLines(pendingLines) {
        var unapplied = { frames: [], texts: [] };
        var decidedFonts = {}; /* フォント名 -> TextFont（適用）/ null（スキップ）/ decision per font name */
        var fontPicker = null; /* ピッカーは初回だけ生成して使い回す / Build the picker only once */

        for (var i = 0; i < pendingLines.length; i++) {
            var pendingLine = pendingLines[i];

            /* 決定済みのフォント名は再質問しない。サイズ・行送りは行ごとの値を適用する
               / Reuse the earlier decision; size and leading still come from this line */
            if (decidedFonts.hasOwnProperty(pendingLine.fontName)) {
                if (!applyPendingLine(pendingLine, decidedFonts[pendingLine.fontName])) {
                    recordUnapplied(unapplied, pendingLine);
                }
                continue;
            }

            if (!fontPicker) fontPicker = createFontPickerDialog(fontIndex.families);
            var pickedFont = showFontPicker(fontPicker, pendingLine);

            /* 「終了」：残りの保留分をすべて未適用として記録し、ループを打ち切る / Quit stops the queue */
            if (pickedFont === PICKER_QUIT) {
                for (var q = i; q < pendingLines.length; q++) {
                    recordUnapplied(unapplied, pendingLines[q]);
                }
                break;
            }

            /* 適用に失敗した行は記録するだけで決定として残さない（同名の次の行では改めて確認する）
               / A failed apply is recorded but not remembered, so the next line asks again */
            if (pickedFont === PICKER_FAILED) {
                recordUnapplied(unapplied, pendingLine);
                continue;
            }

            decidedFonts[pendingLine.fontName] = pickedFont;
            if (!pickedFont) recordUnapplied(unapplied, pendingLine);
        }

        return unapplied;
    }

    /**
     * 未適用の行を含むフレームの背面に、赤・半透明の長方形を目印として置く
     * 実行のたびに古い目印レイヤーは削除し、作成後はレイヤーをロックする
     * @param {TextFrame[]} frameList - 未適用の行を含むフレーム
     * @returns {void}
     */
    function markUnappliedFrames(frameList) {
        removeLayerByName(MARKER_LAYER_NAME);
        if (frameList.length === 0) return;

        var previousActiveLayer = doc.activeLayer;
        var markerLayer = createMarkerLayer(MARKER_LAYER_NAME);

        for (var i = 0; i < frameList.length; i++) {
            createMarkerRect(frameList[i], markerLayer);
        }

        /* 目印を誤って動かさないようロックし、作業レイヤーは元へ戻す / Lock it and restore the active layer */
        markerLayer.locked = true;
        try { doc.activeLayer = previousActiveLayer; } catch (e) { }
    }

    /**
     * 未適用の行（スキップ・適用失敗）をフレーム・文字列として記録する
     * @param {Object} unapplied - 記録先（{ frames, texts }）
     * @param {Object} pendingLine - 対象の保留行
     * @returns {void}
     */
    function recordUnapplied(unapplied, pendingLine) {
        pushUnique(unapplied.texts, pendingLine.lineText);
        pushUnique(unapplied.frames, pendingLine.frame); /* 同一参照なので === で重複排除できる */
    }

    /**
     * 配列に未登録の値だけ追加する（重複防止）
     * @param {Array} targetList - 追加先の配列
     * @param {Object} newValue - 追加する値
     * @returns {void}
     */
    function pushUnique(targetList, newValue) {
        for (var i = 0; i < targetList.length; i++) {
            if (targetList[i] === newValue) return;
        }
        targetList.push(newValue);
    }

    // =========================================
    // フォントの適用 / Applying fonts
    // =========================================

    /**
     * 保留行に、決定したフォントとその行のサイズ・行送りを適用する
     * @param {Object} pendingLine - 対象の保留行
     * @param {TextFont} font - 適用するフォント（null なら適用しない）
     * @returns {boolean} 適用できたかどうか
     */
    function applyPendingLine(pendingLine, font) {
        return applyFontSpecToLine(pendingLine.frame, pendingLine.index, font,
            pendingLine.fontSize, pendingLine.fontLeading);
    }

    /**
     * 段落へフォント・サイズ・行送りを適用する（サイズ・行送りは指定があるものだけ）
     * 段落は使用直前にライブ取得する。保持した参照を使い回すと無効化され、
     * try/catch でも拾えないネイティブクラッシュを起こすため
     * 行送りは絶対値では入れず、行送り÷サイズの百分率を自動行送りに代入する
     * （例：12pt↓16pt → 16/12 ≒ 133.3%）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {number} paragraphIndex - 段落インデックス
     * @param {TextFont} font - 適用するフォント（null なら適用しない）
     * @param {number} sizePt - 文字サイズ（ポイント。null なら変更しない）
     * @param {number} leadingPt - 行送り（ポイント。null なら変更しない）
     * @returns {boolean} 適用できたかどうか
     */
    function applyFontSpecToLine(frame, paragraphIndex, font, sizePt, leadingPt) {
        if (!font) return false;

        try {
            var paragraph = frame.paragraphs[paragraphIndex];
            var attributes = paragraph.characterAttributes;

            attributes.textFont = font;
            if (sizePt) attributes.size = sizePt;
            if (leadingPt && sizePt) {
                paragraph.paragraphAttributes.autoLeadingAmount = (leadingPt / sizePt) * 100;
                attributes.autoLeading = true;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 行末に併記されたフォントサイズ・行送りを切り出し、{ name, size, leading } で返す
     * 例）"ヒラギノ角ゴシック W3 12pt↓16pt" → { name: "ヒラギノ角ゴシック W3", size: 12, leading: 16 }
     * 　　"Helvetica Neue, 14"            → { name: "Helvetica Neue", size: 14, leading: null }
     * 　　"DIN 2014"                      → { name: "DIN 2014", size: null, leading: null }
     * サイズは「単位付き」か「、」「,」「/」区切りのどちらかで必ずアンカーされるため、
     * フォント名末尾の数字（"DIN 2014" など）はサイズと誤認されない
     * @param {string} rawText - 行の文字列
     * @returns {Object} { name: string, size: number, leading: number }（size / leading は未指定なら null）
     */
    function parseFontSpec(rawText) {
        var text = String(rawText).replace(/^[\s　]+|[\s　]+$/g, "");
        var fontSpec = { name: text, size: null, leading: null };

        /* 末尾の行送りは任意：[区切り（↓含む）→ 数値 →（任意で単位）] / Trailing leading is optional */
        var trailingLeading = "(?:[ \\t　、,/↓]+([0-9]+(?:\\.[0-9]+)?)(pt|px|q)?)?$";

        /* パターン1：サイズに単位（pt/px/Q）が付くケース / Size with an explicit unit */
        var sizeWithUnit = new RegExp(
            "^([\\s\\S]*?)[ \\t　、,/]+([0-9]+(?:\\.[0-9]+)?)(pt|px|q)" + trailingLeading, "i");

        /* パターン2：単位は無いが「、」「,」「/」で区切られているケース / Size after a list separator */
        var sizeWithListSeparator = new RegExp(
            "^([\\s\\S]*?)[ \\t　]*[、,/][ \\t　]*([0-9]+(?:\\.[0-9]+)?)(pt|px|q)?" + trailingLeading, "i");

        /* groups: 1=名前 / 2=サイズ数値 / 3=サイズ単位 / 4=行送り数値 / 5=行送り単位 */
        var matched = text.match(sizeWithUnit) || text.match(sizeWithListSeparator);
        if (!matched) return fontSpec;

        var name = String(matched[1]).replace(/[\s　]+$/g, "");
        if (name === "") return fontSpec; /* 名前が空ならサイズ扱いしない / Not a size when the name is empty */

        var sizePt = parseSizeWithUnit(matched[2], matched[3]);
        if (sizePt === null) return fontSpec;

        fontSpec.name = name;
        fontSpec.size = sizePt;

        /* サイズが確定したときだけ行送りを見る / Read the leading only once the size is settled */
        if (matched[4]) fontSpec.leading = parseSizeWithUnit(matched[4], matched[5]);

        return fontSpec;
    }

    /**
     * 数値文字列＋単位をポイント値へ換算する（1Q = 0.25mm、px/pt/単位なしは 1px = 1pt）
     * @param {string} numberText - 数値の文字列
     * @param {string} unitText - 単位（pt / px / q / 未指定）
     * @returns {number} ポイント値。正の値でなければ null
     */
    function parseSizeWithUnit(numberText, unitText) {
        var value = parseFloat(numberText);
        if (!(value > 0)) return null;
        /* 行の指定に書かれた「Q」は環境設定によらず常に級として換算する / "Q" in a line is always read as kyu, regardless of preferences */
        if (unitText && String(unitText).toLowerCase() === "q") {
            return unitToPt(value, UNITS[5]); /* 1Q = 0.25mm → pt */
        }
        return value;
    }

    // =========================================
    // フォントの照合 / Font matching
    // =========================================

    /**
     * インストール済みフォントを索引化する（照合と一覧表示の両方で使う）
     * @returns {Object} フォント索引
     */
    function createFontIndex() {
        var installedFonts = app.textFonts;
        var index = {
            families: [],                /* 表示用のファミリー名（重複なし・ソート済み）*/
            stylesByFamily: {},          /* ファミリー名 -> スタイル名の配列 */
            fontByFamilyStyle: {},       /* ファミリー名＋スタイル名 -> TextFont */
            fontsByNormalizedFamily: {}, /* 正規化ファミリー名 -> TextFont の配列 */
            fontsByNormalizedFull: {},   /* 正規化「ファミリー名＋スタイル名」-> TextFont の配列 */
            normalizedFonts: [],         /* 部分一致の走査用 */
            customFontByNormalized: {}   /* 正規化した CUSTOM_FONT_MAP のキー -> 置換後のフォント名 */
        };

        for (var i = 0; i < installedFonts.length; i++) {
            var currentFont = installedFonts[i];
            var family = currentFont.family;
            var style = currentFont.style;

            if (!index.stylesByFamily.hasOwnProperty(family)) {
                index.stylesByFamily[family] = [];
                index.families.push(family);
            }
            index.stylesByFamily[family].push(style);
            index.fontByFamilyStyle[makeFamilyStyleKey(family, style)] = currentFont;

            var normalizedFamily = normalize(family);
            var normalizedFull = normalize(family + " " + style);
            pushToBucket(index.fontsByNormalizedFamily, normalizedFamily, currentFont);
            pushToBucket(index.fontsByNormalizedFull, normalizedFull, currentFont);

            index.normalizedFonts.push({
                font: currentFont,
                family: normalizedFamily,
                full: normalizedFull,
                name: normalize(currentFont.name)
            });
        }

        index.families.sort();
        for (var familyName in index.stylesByFamily) {
            if (index.stylesByFamily.hasOwnProperty(familyName)) {
                index.stylesByFamily[familyName].sort();
            }
        }

        /* カスタム置換ルールも正規化しておき、大文字小文字やスペースのゆれを吸収する
           / Normalize the custom map so case and spacing differences still hit */
        for (var customKey in CUSTOM_FONT_MAP) {
            if (CUSTOM_FONT_MAP.hasOwnProperty(customKey)) {
                index.customFontByNormalized[normalize(customKey)] = CUSTOM_FONT_MAP[customKey];
            }
        }

        return index;
    }

    /**
     * 索引のバケット（キー -> 配列）へ値を追加する
     * @param {Object} bucketMap - バケットを保持するオブジェクト
     * @param {string} key - キー
     * @param {Object} value - 追加する値
     * @returns {void}
     */
    function pushToBucket(bucketMap, key, value) {
        if (!bucketMap.hasOwnProperty(key)) bucketMap[key] = [];
        bucketMap[key].push(value);
    }

    /**
     * 素のオブジェクトを辞書として引く（"constructor" などで Object.prototype を拾わないようにする）
     * @param {Object} dictionary - 辞書として使うオブジェクト
     * @param {string} key - キー
     * @returns {Object} 対応する値（無ければ null）
     */
    function getOwn(dictionary, key) {
        return dictionary.hasOwnProperty(key) ? dictionary[key] : null;
    }

    /**
     * ファミリー名＋スタイル名の索引用キーを作る
     * @param {string} family - ファミリー名
     * @param {string} style - スタイル名
     * @returns {string} 索引用のキー
     */
    function makeFamilyStyleKey(family, style) {
        return family + " " + style;
    }

    /**
     * 指定ファミリーに属するスタイル名の一覧を返す
     * @param {string} familyName - ファミリー名
     * @returns {string[]} スタイル名の配列（無ければ空配列）
     */
    function stylesForFamily(familyName) {
        return getOwn(fontIndex.stylesByFamily, familyName) || [];
    }

    /**
     * ファミリー名＋スタイル名から TextFont を取得する
     * @param {string} familyName - ファミリー名
     * @param {string} styleName - スタイル名
     * @returns {TextFont} 該当するフォント（無ければ null）
     */
    function fontFor(familyName, styleName) {
        return getOwn(fontIndex.fontByFamilyStyle, makeFamilyStyleKey(familyName, styleName));
    }

    /**
     * 文字列に対応するフォントを、厳密一致 → あいまい一致の順に探す
     * @param {string} fontName - 行から切り出したフォント名
     * @returns {Object} { font: TextFont, confident: boolean }
     *   confident=true … 厳密一致。自動適用してよい
     *   confident=false 且つ font!=null … あいまい一致。ピッカーの初期値に使う
     *   font=null … 未一致。ピッカーで一から選ばせる
     */
    function findFont(fontName) {
        var targetName = getOwn(fontIndex.customFontByNormalized, normalize(fontName)) || fontName;

        var exactFont = findExactFont(targetName);
        if (exactFont) return { font: exactFont, confident: true };

        return { font: findFuzzyFont(targetName), confident: false };
    }

    /**
     * 厳密一致でフォントを探す（PostScript 名 → ファミリー名＋スタイル名 → ファミリー名）
     * 照合はスペース・ピリオドを除いた正規化文字列で行う
     * @param {string} fontName - 探すフォント名
     * @returns {TextFont} 該当するフォント（無ければ null）
     */
    function findExactFont(fontName) {
        try {
            return app.textFonts.getByName(fontName); /* PostScript 名の完全一致 */
        } catch (e) { }

        var query = normalize(fontName);

        var exactFull = getOwn(fontIndex.fontsByNormalizedFull, query);
        if (exactFull) return getBestStyle(exactFull);

        var exactFamily = getOwn(fontIndex.fontsByNormalizedFamily, query);
        if (exactFamily) return getBestStyle(exactFamily);

        return null;
    }

    /**
     * あいまい一致でフォントを探す（ファミリー名の部分一致 → フォント名全体の部分一致 → 先頭ワード）
     * 例）"Jenson" → "Adobe Jenson Pro"、"Myriad Pro Cond"（未インストール）→ "Myriad Pro"
     * @param {string} fontName - 探すフォント名
     * @returns {TextFont} 該当するフォント（無ければ null）
     */
    function findFuzzyFont(fontName) {
        var query = normalize(fontName);

        var partialFamily = collectIndexedFonts(function (fontInfo) {
            return fontInfo.family.indexOf(query) !== -1;
        });
        if (partialFamily.length > 0) return getBestStyle(partialFamily);

        var partialFull = collectIndexedFonts(function (fontInfo) {
            return fontInfo.full.indexOf(query) !== -1 || fontInfo.name.indexOf(query) !== -1;
        });
        if (partialFull.length > 0) return getBestStyle(partialFull);

        /* 最終手段：先頭ワードがファミリー名に含まれればOK。2文字以下は誤マッチ防止のため対象外
           / Last resort: match on the first word, ignoring words of two characters or fewer */
        var firstWord = String(fontName).toLowerCase()
            .replace(/[.　]+/g, " ").replace(/^\s+/, "").split(/\s+/)[0] || "";
        if (firstWord.length >= 3) {
            var looseFamily = collectIndexedFonts(function (fontInfo) {
                return fontInfo.family.indexOf(firstWord) !== -1;
            });
            if (looseFamily.length > 0) return getBestStyle(looseFamily);
        }

        return null;
    }

    /**
     * 正規化済みフォント索引から、条件に合うフォントを集めて返す
     * @param {function} isMatch - 判定関数（{ font, family, full, name } を受け取る）
     * @returns {TextFont[]} 条件に合ったフォント
     */
    function collectIndexedFonts(isMatch) {
        var matchedFonts = [];
        var normalizedFonts = fontIndex.normalizedFonts;
        for (var i = 0; i < normalizedFonts.length; i++) {
            if (isMatch(normalizedFonts[i])) {
                matchedFonts.push(normalizedFonts[i].font);
            }
        }
        return matchedFonts;
    }

    /**
     * フォント名照合用の正規化：小文字化し、空白（半角・全角）とピリオドを除去する
     * 「Bank Gothic」と「BankGothic」、「Mrs. Eaves」と「Mrs Eaves」のゆれを吸収する
     * @param {string} text - 対象の文字列
     * @returns {string} 正規化した文字列
     */
    function normalize(text) {
        return String(text)
            .toLowerCase()
            .replace(/[.\s　]+/g, "");
    }

    /**
     * 候補フォントから STYLE_PRIORITY の順で最適なスタイルを選ぶ
     * @param {TextFont[]} candidateFonts - 候補のフォント
     * @returns {TextFont} 選ばれたフォント（該当が無ければ候補の先頭）
     */
    function getBestStyle(candidateFonts) {
        for (var k = 0; k < STYLE_PRIORITY.length; k++) {
            for (var j = 0; j < candidateFonts.length; j++) {
                if (candidateFonts[j].style.toLowerCase() === STYLE_PRIORITY[k]) {
                    return candidateFonts[j];
                }
            }
        }
        return candidateFonts[0];
    }

    // =========================================
    // ダイアログ / Dialogs
    // =========================================

    /**
     * 適用中に表示するプログレスバー（パレット）を作成する
     * @param {number} totalCount - 対象の総数
     * @returns {Object} { window: Window, update: function }
     */
    function createProgressPalette(totalCount) {
        var progressWindow = new Window("palette", getLabel(LABELS.dialog.progressTitle));
        setupWindow(progressWindow);

        var progressBar = progressWindow.add("progressbar", undefined, 0, totalCount);
        progressBar.preferredSize = PROGRESS_BAR_SIZE;

        var countLabel = progressWindow.add("statictext", undefined, "0 / " + totalCount);
        countLabel.preferredSize.width = PROGRESS_TEXT_WIDTH;

        progressWindow.show();
        progressWindow.update();

        return {
            window: progressWindow,
            update: function (doneCount) {
                progressBar.value = doneCount;
                countLabel.text = doneCount + " / " + totalCount;
                progressWindow.update();
            }
        };
    }

    /**
     * 適用できなかった文字列を一覧表示する
     * コピーはダイアログを閉じてから行う（モーダル表示中はドキュメントを触らない）
     * @param {string[]} unappliedTexts - 未適用の文字列
     * @returns {boolean} クリップボードへのコピーが要求されたか
     */
    function showUnappliedDialog(unappliedTexts) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.resultTitle));
        setupWindow(dialog);

        dialog.add("statictext", undefined, labelText(LABELS.dialog.resultHeader));

        /* 一覧（読み取り専用・複数行・スクロール可）。手動で選択もできる / Read-only scrollable list */
        var listField = dialog.add("edittext", undefined, unappliedTexts.join("\n"),
            { multiline: true, scrolling: true, readonly: true });
        listField.helpTip = getLabel(LABELS.dialog.tipResultList);
        listField.preferredSize = RESULT_FIELD_SIZE;

        /* === ボタンエリア（左右分割：左=コピー／右=閉じる）=== */
        var buttonRow = addButtonRow(dialog);
        var btnCopy = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.copy));
        /* ESC・ウィンドウの閉じるボタンは 2 を返すため、コピーには別の値を使う
           / ESC and the close box return 2, so copy uses its own result code */
        btnCopy.onClick = function () { dialog.close(3); }; /* 3 = コピーして閉じる / copy and close */

        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "ok" });

        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(dialog);
        prepareDialogWindow(dialog, SCRIPT_NAME + "_result");
        return dialog.show() === 3;
    }

    /**
     * 文字列をクリップボードへコピーする
     * ExtendScript には直接コピーする API が無いため、一時テキストフレーム経由でコピーする
     * app.copy() は黙って無視されることがあるので、再描画とメニューコマンドを挟む
     * @param {string} textToCopy - コピーする文字列
     * @returns {void}
     */
    function copyTextToClipboard(textToCopy) {
        var previousSelection = doc.selection; /* コピー後に元へ戻すため控える / Keep it to restore later */
        var editableLayer = getEditableLayer();
        if (!editableLayer) return;

        var tempFrame = null;
        try {
            doc.activeLayer = editableLayer;
            tempFrame = doc.textFrames.add();
            tempFrame.contents = textToCopy;
            /* カンバスの範囲（±約16000pt）を超えると位置指定に失敗するため控えめに逃がす
               / Keep it inside the canvas bounds, otherwise the move fails */
            tempFrame.position = [-5000, -5000]; /* 画面外に逃がす / Move it off-canvas */

            app.redraw(); /* 追加直後のフレームは再描画しないとコピー対象にならない */
            app.executeMenuCommand("deselectall");
            tempFrame.selected = true;
            app.redraw();
            app.executeMenuCommand("copy"); /* app.copy() は黙って無視されることがある */
            app.redraw(); /* コピー確定前に削除すると空になる */
        } catch (e) {
        } finally {
            if (tempFrame) {
                try { tempFrame.remove(); } catch (err) { }
            }
            try { doc.selection = previousSelection; } catch (err) { doc.selection = null; }
        }
    }

    /**
     * 一時オブジェクトを置ける（ロックも非表示もされていない）レイヤーを返す
     * @returns {Layer} 編集できるレイヤー（無ければ null）
     */
    function getEditableLayer() {
        if (!doc.activeLayer.locked && doc.activeLayer.visible) return doc.activeLayer;

        for (var i = 0; i < doc.layers.length; i++) {
            if (!doc.layers[i].locked && doc.layers[i].visible) return doc.layers[i];
        }
        return null;
    }

    // =========================================
    // フォントピッカー / Font picker
    // =========================================

    /**
     * 保留行のフォントを対話的に選ばせる
     * ダイアログは作り直さず使い回し、ここでは対象テキストと選択のリセットだけ行う
     * モーダル表示中はドキュメントを変更しない（ライブプレビューはクラッシュ要因のため廃止）
     * @param {Object} fontPicker - createFontPickerDialog が返したピッカー
     * @param {Object} pendingLine - 対象の保留行
     * @returns {TextFont} 適用したフォント／スキップは null／適用失敗は PICKER_FAILED／終了は PICKER_QUIT
     */
    function showFontPicker(fontPicker, pendingLine) {
        zoomToFrame(pendingLine.frame); /* 対象を画面にフィット（モーダル表示前に行う）*/

        fontPicker.targetLabel.text = pendingLine.lineText;
        fontPicker.searchField.text = "";
        filterFamilyDropdown(fontPicker, "");
        initializePickerSelection(fontPicker, pendingLine.initialFont);

        /* 「適用」=1 /「スキップ」=2（閉じるを含む）/「終了」=3 */
        /* 開くたびに呼ぶ。2回目からは選択範囲を測り直すだけ / Call on every show; later calls only re-measure the selection */
        prepareDialogWindow(fontPicker.dialog, SCRIPT_NAME);
        var dialogResult = fontPicker.dialog.show();

        if (dialogResult === 1 && fontPicker.familyDropdown.selection && fontPicker.styleDropdown.selection) {
            var chosenFont = fontFor(fontPicker.familyDropdown.selection.text,
                fontPicker.styleDropdown.selection.text);
            /* 実適用はダイアログを閉じた後（＝モーダル表示外）なので安全 / Apply once the modal is closed */
            if (applyPendingLine(pendingLine, chosenFont)) return chosenFont;
            return PICKER_FAILED;
        }

        return (dialogResult === 3) ? PICKER_QUIT : null;
    }

    /**
     * フォントピッカーの UI を生成する
     * ボタンは name:"ok"/"cancel" なので、判定は dialog.show() の戻り値で行う
     * @param {string[]} familyNames - ファミリー名の一覧
     * @returns {Object} ピッカー（{ dialog, familyDropdown, styleDropdown, searchField, targetLabel, familyNames }）
     */
    function createFontPickerDialog(familyNames) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.pickerTitle));
        setupWindow(dialog);

        /* 対象テキスト（パネルの外）。値はピックごとに差し替えるので空で作り、幅だけ確保する */
        var targetRow = dialog.add("group");
        targetRow.add("statictext", undefined, labelText(LABELS.fieldLabel.targetText));
        var targetLabel = targetRow.add("statictext", undefined, "", { truncate: "end" });
        targetLabel.preferredSize.width = PICKER_TARGET_WIDTH;

        var applyFontPanel = dialog.add("panel", undefined, getLabel(LABELS.panel.applyFont));
        setupPanel(applyFontPanel);

        var searchField = addPickerRow(applyFontPanel, LABELS.fieldLabel.search, "edittext", "");

        /* 巨大配列での dropdownlist 生成を繰り返すと Illustrator が落ちるため、
           このダイアログ（と familyDropdown）は実行中に1回だけ生成して使い回す */
        var familyDropdown = addPickerRow(applyFontPanel, LABELS.fieldLabel.family, "dropdownlist", familyNames);
        var styleDropdown = addPickerRow(applyFontPanel, LABELS.fieldLabel.style, "dropdownlist", []);

        /* === ボタンエリア（左右分割：左=終了／右=スキップ・適用）=== */
        var buttonRow = addButtonRow(dialog);
        var btnQuit = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.quit));
        btnQuit.onClick = function () { dialog.close(3); }; /* 3 = 終了（保留分を打ち切る）/ quit the queue */

        var btnSkip = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.skip), { name: "cancel" });
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.apply), { name: "ok" });

        var fontPicker = {
            dialog: dialog,
            familyDropdown: familyDropdown,
            styleDropdown: styleDropdown,
            searchField: searchField,
            targetLabel: targetLabel,
            familyNames: familyNames
        };

        bindFontPickerEvents(fontPicker); /* イベントも生成時に1回だけ接続する / Bind the events once too */
        return fontPicker;
    }

    /**
     * ピッカーのパネルへ「項目名＋コントロール」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelNode - 項目名のラベル（{ ja, en }）
     * @param {string} controlType - コントロールの種類（"edittext" / "dropdownlist"）
     * @param {Object} initialValue - 初期値（edittext は文字列、dropdownlist は項目の配列）
     * @returns {Object} 追加したコントロール
     */
    function addPickerRow(parentPanel, labelNode, controlType, initialValue) {
        var row = parentPanel.add("group");
        var rowLabel = row.add("statictext", undefined, labelText(labelNode));
        rowLabel.preferredSize.width = PICKER_LABEL_WIDTH;

        var control = row.add(controlType, undefined, initialValue);
        control.preferredSize.width = PICKER_FIELD_WIDTH;
        return control;
    }

    /**
     * フォントピッカーの検索・ファミリー変更イベントを接続する
     * （ドキュメントは変更しない。実適用は「適用」確定後に行う）
     * @param {Object} fontPicker - 対象のピッカー
     * @returns {void}
     */
    function bindFontPickerEvents(fontPicker) {
        fontPicker.searchField.onChanging = function () {
            filterFamilyDropdown(fontPicker, fontPicker.searchField.text);
            populateStyleDropdown(fontPicker, null);
        };

        fontPicker.familyDropdown.onChange = function () {
            populateStyleDropdown(fontPicker, null);
        };
    }

    /**
     * ピッカーの初期選択（ファミリー・スタイル）を設定する
     * @param {Object} fontPicker - 対象のピッカー
     * @param {TextFont} initialFont - 初期値にするフォント（無ければ null）
     * @returns {void}
     */
    function initializePickerSelection(fontPicker, initialFont) {
        selectDropdownItemByText(fontPicker.familyDropdown,
            initialFont ? initialFont.family : fontPicker.familyNames[0]);
        if (!fontPicker.familyDropdown.selection) fontPicker.familyDropdown.selection = 0;

        populateStyleDropdown(fontPicker, initialFont ? initialFont.style : null);
    }

    /**
     * 選択中のファミリーに合わせてスタイル一覧を作り直し、選ぶべきスタイルを選択する
     * @param {Object} fontPicker - 対象のピッカー
     * @param {string} preferredStyleName - 優先して選ぶスタイル名（無ければ null）
     * @returns {void}
     */
    function populateStyleDropdown(fontPicker, preferredStyleName) {
        var styleDropdown = fontPicker.styleDropdown;
        styleDropdown.removeAll();
        if (!fontPicker.familyDropdown.selection) return;

        var styles = stylesForFamily(fontPicker.familyDropdown.selection.text);
        for (var i = 0; i < styles.length; i++) {
            styleDropdown.add("item", styles[i]);
        }

        if (preferredStyleName) selectDropdownItemByText(styleDropdown, preferredStyleName);
        if (!styleDropdown.selection && styleDropdown.items.length > 0) styleDropdown.selection = 0;
    }

    /**
     * 検索クエリ（部分一致・大文字小文字を無視）でファミリーのプルダウンを絞り込む
     * 絞り込み後も、可能なら直前に選択していたファミリーを選び直す
     * 空クエリで全件表示済みなら作り直さない（既存コントロールへの item 追加は安全だが、無駄を避ける）
     * @param {Object} fontPicker - 対象のピッカー
     * @param {string} searchQuery - 検索クエリ（空文字なら全件）
     * @returns {void}
     */
    function filterFamilyDropdown(fontPicker, searchQuery) {
        var familyDropdown = fontPicker.familyDropdown;
        var familyNames = fontPicker.familyNames;
        var needle = String(searchQuery).toLowerCase();

        if (needle === "" && familyDropdown.items.length === familyNames.length) return;

        var previousFamily = familyDropdown.selection ? familyDropdown.selection.text : null;
        familyDropdown.removeAll();
        for (var i = 0; i < familyNames.length; i++) {
            if (needle === "" || familyNames[i].toLowerCase().indexOf(needle) !== -1) {
                familyDropdown.add("item", familyNames[i]);
            }
        }

        if (previousFamily) selectDropdownItemByText(familyDropdown, previousFamily);
        if (!familyDropdown.selection && familyDropdown.items.length > 0) familyDropdown.selection = 0;
    }

    /**
     * プルダウンで指定テキストの項目を選択する（無ければ何もしない）
     * @param {DropDownList} dropdown - 対象のプルダウン
     * @param {string} itemText - 選択する項目のテキスト
     * @returns {void}
     */
    function selectDropdownItemByText(dropdown, itemText) {
        for (var i = 0; i < dropdown.items.length; i++) {
            if (dropdown.items[i].text === itemText) {
                dropdown.selection = i;
                return;
            }
        }
    }

    /**
     * 対象フレームが画面にフィットするようズーム＋センタリングする
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {void}
     */
    function zoomToFrame(frame) {
        try {
            var view = doc.activeView;
            var bounds = frame.geometricBounds; /* [left, top, right, bottom] */
            var frameWidth = bounds[2] - bounds[0];
            var frameHeight = bounds[1] - bounds[3];
            if (frameWidth <= 0 || frameHeight <= 0) return;

            var viewBounds = view.bounds; /* 現在の表示範囲（ドキュメント座標）/ current view in document coordinates */
            var viewWidth = viewBounds[2] - viewBounds[0];
            var viewHeight = viewBounds[1] - viewBounds[3];

            /* 表示領域の ZOOM_FIT_RATIO に収まる倍率を現在ズームから算出し、上下限に収める */
            var fitZoom = Math.min(viewWidth / frameWidth, viewHeight / frameHeight) * view.zoom * ZOOM_FIT_RATIO;
            view.zoom = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, fitZoom));
            view.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (e) { }
    }

    /**
     * 現在の表示（ズーム率・中心点）を控える
     * @returns {Object} 表示状態（{ zoom, centerPoint }。取得できなければ null）
     */
    function captureView() {
        try {
            var view = doc.activeView;
            return { zoom: view.zoom, centerPoint: view.centerPoint };
        } catch (e) {
            return null;
        }
    }

    /**
     * 控えておいた表示（ズーム率・中心点）へ戻す
     * @param {Object} viewState - captureView が返した表示状態（null なら何もしない）
     * @returns {void}
     */
    function restoreView(viewState) {
        if (!viewState) return;

        try {
            var view = doc.activeView;
            view.zoom = viewState.zoom;
            view.centerPoint = viewState.centerPoint;
        } catch (e) { }
    }

    // =========================================
    // 目印レイヤー / Marker layer
    // =========================================

    /**
     * 目印レイヤーを新規作成し、テキストの背面に来るよう最背面へ送る
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 作成したレイヤー
     */
    function createMarkerLayer(layerName) {
        var markerLayer = doc.layers.add();
        markerLayer.name = layerName;

        try {
            markerLayer.move(doc.layers[doc.layers.length - 1], ElementPlacement.PLACEAFTER);
        } catch (e) { }

        return markerLayer;
    }

    /**
     * テキストフレームの背面（＝目印レイヤー上）に、赤・半透明の長方形を作る
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {Layer} layer - 目印を置くレイヤー
     * @returns {void}
     */
    function createMarkerRect(frame, layer) {
        try {
            var bounds = frame.geometricBounds; /* [left, top, right, bottom]（線幅は含まない）*/
            var width = bounds[2] - bounds[0];
            var height = bounds[1] - bounds[3];
            if (width <= 0 || height <= 0) return;

            var markerRect = layer.pathItems.rectangle(bounds[1], bounds[0], width, height);
            markerRect.stroked = false;
            markerRect.filled = true;
            markerRect.fillColor = makeRedColor();
            markerRect.opacity = MARKER_OPACITY;
        } catch (e) { }
    }

    /**
     * ドキュメントのカラースペースに合わせた赤色を返す
     * @returns {Object} CMYKColor または RGBColor
     */
    function makeRedColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 100;
            cmyk.yellow = 100;
            cmyk.black = 0;
            return cmyk;
        }

        var rgb = new RGBColor();
        rgb.red = 255;
        rgb.green = 0;
        rgb.blue = 0;
        return rgb;
    }

    /**
     * 指定名のレイヤーがあれば削除する（ロックされていても解除してから削除する）
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function removeLayerByName(layerName) {
        var layer = getLayerByName(layerName);
        if (!layer) return;

        unlockContainer(layer);
        unlockPageItems(layer.pageItems);

        try {
            layer.remove();
        } catch (e) { }
    }

    /**
     * 指定名のレイヤーを取得する
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 該当するレイヤー（無ければ null）
     */
    function getLayerByName(layerName) {
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            return null;
        }
    }

    /**
     * レイヤー／ページアイテムのロックを解除し、表示状態に戻す
     * 表示のプロパティはレイヤーが visible、ページアイテムが hidden と異なる
     * @param {Object} target - Layer または PageItem
     * @returns {void}
     */
    function unlockContainer(target) {
        try {
            target.locked = false;
            if (target.typename === "Layer") {
                target.visible = true;
            } else {
                target.hidden = false;
            }
        } catch (e) { }
    }

    /**
     * ページアイテムを再帰的にロック解除・表示する
     * @param {Object} items - PageItems コレクション
     * @returns {void}
     */
    function unlockPageItems(items) {
        for (var i = 0; i < items.length; i++) {
            unlockContainer(items[i]);

            if (items[i].typename === "GroupItem") {
                unlockPageItems(items[i].pageItems);
            }
        }
    }
})();
