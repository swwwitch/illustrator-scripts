#targetengine "SwwwitchPalettes"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.AiMemoPalette
Name=メモ入力パレット
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.12
Release-Date=2026-06-15
Description-BEGIN
メモ入力用のフローティングパレットです。
選択オブジェクトやクリップボードからテキストを読み込み、空行・改行を整理して、テキストファイルへの保存やクリップボードへのコピーができます。

A floating palette for taking notes.
It can pull text in from the selection or the clipboard, tidy up blank lines and returns, and save the result to a text file or copy it back to the clipboard.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

メモ入力用のフローティングパレットです。
選択オブジェクトやクリップボードからテキストを読み込み、空行・改行を整理して、テキストファイルへの保存やクリップボードへのコピーができます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiMemoPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n41e91e4b1a09

### Overview

A floating palette for taking notes.
It can pull text in from the selection or the clipboard, tidy up blank lines and returns, and save the result to a text file or copy it back to the clipboard.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiMemoPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiMemoPalette";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.12";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "メモ入力パレット"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiMemoPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiMemoPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n41e91e4b1a09"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AiMemoPalette.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /**
     * 原案 / Original idea
     * @author こじらせたクマー (@nice_lotus120)
     * @discussion https://note.com/nice_lotus120/n/n6291a432b30d
     */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    /* パレットの透明度 / Palette opacity */
    var PALETTE_OPACITY = 0.97;

    /* 保存先モード / Save destination mode
       'dialog'  = 保存ダイアログで場所と名前を選ぶ（A）/ Choose location & name via the save dialog (A)
       'desktop' = 常にデスクトップへ memo-<ドキュメント名>-<yyyymmdd>.txt（B）/ Always save to the Desktop (B) */
    var SAVE_LOCATION_MODE = 'desktop'; // デフォルトは B / Default is B

    /* パレットを閉じたときに内容をクリアするか / Whether to clear the content when the palette is closed
       true  = 閉じると内容をクリア / Clear on close
       false = 内容を保持して次回復元（デフォルト）/ Keep content and restore next time (default) */
    var CLEAR_ON_CLOSE = false;

    /* 同じ行とみなす上端の差（pt）/ Difference in top edge treated as the same row (pt) */
    var SAME_ROW_TOLERANCE_PT = 10;

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

    var PALETTE_MIN_SIZE      = [300, 200];       /* パレットの最小サイズ [幅,高さ] */
    var PALETTE_INITIAL_WIDTH = 400;              /* 保存位置が無いときの初期幅 */
    var MEMO_AREA_MIN_SIZE    = [300, 100];       /* テキスト欄の最小サイズ [幅,高さ] */
    var BUTTON_SIZE           = [76, 24];         /* 標準ボタンのサイズ [幅,高さ] */
    var WIDE_BUTTON_SIZE      = [120, 24];        /* 長いラベルのボタンのサイズ [幅,高さ] */
    var BUTTON_HEIGHT         = 24;               /* 幅をラベルに合わせるボタンの高さ */

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
            title: { ja: "テキスト一時保管", en: "Text Stash" }
        },
        fieldLabel: {
            mode: { ja: "モード", en: "Mode" }
        },
        radio: {
            replace: { ja: "置き換え", en: "Replace" },
            append: { ja: "追加", en: "Append" }
        },
        button: {
            load: { ja: "読み込む", en: "Load" },
            save: { ja: "保存", en: "Save" },
            clear: { ja: "クリア", en: "Clear" },
            copyAll: { ja: "すべてをコピー", en: "Copy All" },
            removeBlanks: { ja: "空行削除", en: "Remove Blanks" },
            removeBreaks: { ja: "改行削除", en: "Remove Breaks" }
        },
        tooltip: {
            appendMode: {
                ja: "読み込んだテキストを、テキスト欄の末尾に空行を1つ挟んで追加します",
                en: "Adds the loaded text to the end of the memo, separated by one blank line"
            },
            replaceMode: {
                ja: "読み込んだテキストで、テキスト欄の内容を置き換えます",
                en: "Replaces the memo with the loaded text"
            },
            load: {
                ja: "選択オブジェクトのテキストを読み込みます（モードに従って追加／置き換え）。「追加」モードでは既存テキストとの間に空行を入れます",
                en: "Load text from the selected objects (append or replace per mode). In Append mode, a blank line is inserted between the existing and new text"
            },
            removeBlanks: {
                ja: "テキスト欄の空行（空白のみの行を含む）を削除します",
                en: "Remove blank lines (including whitespace-only lines) from the memo"
            },
            removeBreaks: {
                ja: "テキスト欄の改行をすべて削除して1行にまとめます",
                en: "Remove all line breaks in the memo, joining everything into one line"
            },
            clear: {
                ja: "テキスト欄を空にして再起動します",
                en: "Clear the memo and restart"
            },
            save: {
                ja: "メモをデスクトップにテキストファイルとして保存します",
                en: "Save the memo to the Desktop as a text file"
            },
            copyAll: {
                ja: "テキスト欄の内容をクリップボードへコピーします",
                en: "Copy the memo to the clipboard"
            }
        },
        alert: {
            savePrompt: {
                ja: "テキストファイル（.txt）の保存先を指定してください",
                en: "Choose where to save the text file (.txt)"
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrame: {
                ja: "選択されたオブジェクトにテキストフレームがありません。",
                en: "No text frame is found in the selection."
            },
            loadFailed: { ja: "テキストを追加できませんでした: ", en: "Could not load text: " },
            writeFailed: { ja: "ファイルを書き込めませんでした。", en: "Could not write the file." },
            clipboardEmpty: {
                ja: "選択オブジェクトもクリップボードのテキストもありません。",
                en: "No selected object and no clipboard text."
            },
            copyFailed: { ja: "クリップボードへコピーできませんでした。", en: "Could not copy to the clipboard." }
        },
        status: {
            unsaved: { ja: "(ドキュメント未保存)", en: "(Unsaved document)" }
        },
        footer: {
            source: { ja: "保存元", en: "Source" },
            savedAt: { ja: "保存日時", en: "Saved at" }
        }
    };

    // =========================================
    // メインエンジン用ワーカー / Main-engine workers
    // =========================================
    /* 以下の WORKER_* はメインエンジン側で評価される文字列。常駐エンジンの app はパレット表示中に
       DOM 接続を失い "there is no document" を投げるため、DOM に触る処理は BridgeTalk で
       メインエンジンへ渡す（結果は非同期）。結果は "マーカー:ペイロード" 形式で返し、
       マルチバイト欠落を防ぐためテキストは encodeURIComponent で受け渡す。
       These WORKER_* strings are evaluated in the main engine, which keeps a live DOM connection
       while the palette is up. Results come back asynchronously as "MARKER:payload". */

    /* アクティブドキュメントの取得 / Resolve the active document */
    var WORKER_TARGET_DOC =
        '    if (app.documents.length === 0) return "NODOC:";' +
        '    var targetDoc = app.activeDocument;';

    /* 選択の待避と復元（breakLink・貼り付けで選択が変わるため）/ Save and restore the selection */
    var WORKER_SELECTION =
        '    var savedSelection = [];' +
        '    function saveSelection() {' +
        '        savedSelection = [];' +
        '        for (var i = 0; i < targetDoc.selection.length; i++) savedSelection.push(targetDoc.selection[i]);' +
        '    }' +
        '    function restoreSelection() {' +
        '        try {' +
        '            targetDoc.selection = null;' +
        // 文字ツールで編集中の TextRange には selected が無いため飛ばす
        '            for (var i = 0; i < savedSelection.length; i++) {' +
        '                if (savedSelection[i].typename !== "TextRange") savedSelection[i].selected = true;' +
        '            }' +
        '        } catch (err) {}' +
        '    }';

    /* テキストの収集・整列・連結 / Collect, sort and join text */
    var WORKER_COLLECTOR =
        '    var collectedEntries = [];' +
        // シンボルの読み取り方は呼び出し側で差し替える（未設定ならシンボルは無視）
        '    var readSymbolItem = null;' +
        // 実質的に空のフレームは無視し、テキストと位置（左端・上端）を控える
        '    function pushFrame(textFrame) {' +
        '        var frameText = textFrame.contents;' +
        '        if (!frameText || frameText.replace(/[\\r\\n\\x03]/g, "").replace(/\\s+/g, "") === "") return;' +
        '        var frameBounds = textFrame.geometricBounds;' +
        '        collectedEntries.push({ text: frameText, left: frameBounds[0], top: frameBounds[1] });' +
        '    }' +
        // グループ・クリップグループの中もたどる
        '    function collectFrom(pageItem) {' +
        '        if (pageItem.typename === "TextFrame") { pushFrame(pageItem); return; }' +
        '        if (pageItem.typename === "SymbolItem") { if (readSymbolItem) readSymbolItem(pageItem); return; }' +
        '        if (pageItem.pageItems) { for (var i = 0; i < pageItem.pageItems.length; i++) collectFrom(pageItem.pageItems[i]); }' +
        '    }' +
        // カンバス上の位置で上から順（同じ高さは左から右）に並べて連結する
        '    function joinCollected(dedupe) {' +
        '        collectedEntries.sort(function (entryA, entryB) {' +
        '            if (Math.abs(entryB.top - entryA.top) <= ' + SAME_ROW_TOLERANCE_PT + ') return entryA.left - entryB.left;' +
        '            return entryB.top - entryA.top;' +
        '        });' +
        '        var seenTexts = {};' +
        '        var orderedTexts = [];' +
        '        for (var i = 0; i < collectedEntries.length; i++) {' +
        '            var seenKey = "k_" + collectedEntries[i].text;' +
        '            if (dedupe && seenTexts[seenKey]) continue;' +
        '            seenTexts[seenKey] = true;' +
        '            orderedTexts.push(collectedEntries[i].text);' +
        '        }' +
        '        return orderedTexts.join(String.fromCharCode(10));' +
        '    }';

    /* 選択オブジェクトからテキストを収集する（シンボル内も対象）/ Collect text from the selection (symbols included) */
    var WORKER_READ_SELECTION =
        '(function () {' +
        WORKER_TARGET_DOC +
        '    if (!targetDoc.selection || targetDoc.selection.length === 0) return "NOSEL:";' +
        WORKER_SELECTION +
        WORKER_COLLECTOR +
        // シンボルは一時レイヤーへ複製してから展開する（元のシンボルには触れない）
        '    var symbolReadLayer = null;' +
        '    function getSymbolReadLayer() {' +
        '        if (!symbolReadLayer) {' +
        '            symbolReadLayer = targetDoc.layers.add();' +
        '            symbolReadLayer.name = "__TextMemo_symbol_read__";' +
        '        }' +
        '        return symbolReadLayer;' +
        '    }' +
        // 複製なのでその場で breakLink してよい（入れ子シンボルも多段展開）
        '    function harvestDuplicate(pageItem) {' +
        '        if (!pageItem) return;' +
        '        if (pageItem.typename === "TextFrame") { pushFrame(pageItem); return; }' +
        '        if (pageItem.typename === "SymbolItem") {' +
        '            try {' +
        '                targetDoc.selection = null;' +
        '                pageItem.selected = true;' +
        '                pageItem.breakLink();' +
        '                var expandedItems = [];' +
        '                for (var i = 0; i < targetDoc.selection.length; i++) expandedItems.push(targetDoc.selection[i]);' +
        '                for (var j = 0; j < expandedItems.length; j++) harvestDuplicate(expandedItems[j]);' +
        '            } catch (err) {}' +
        '            return;' +
        '        }' +
        '        if (pageItem.pageItems) { for (var k = 0; k < pageItem.pageItems.length; k++) harvestDuplicate(pageItem.pageItems[k]); }' +
        '    }' +
        '    readSymbolItem = function (symbolItem) {' +
        '        harvestDuplicate(symbolItem.duplicate(getSymbolReadLayer(), ElementPlacement.PLACEATBEGINNING));' +
        '    };' +
        '    saveSelection();' +
        '    for (var i = 0; i < savedSelection.length; i++) collectFrom(savedSelection[i]);' +
        // 一時レイヤーは中身ごと削除する
        '    if (symbolReadLayer) { try { symbolReadLayer.remove(); } catch (err) {} }' +
        '    restoreSelection();' +
        '    if (collectedEntries.length === 0) return "NOTF:";' +
        '    return "OK:" + encodeURIComponent(joinCollected(false));' +
        '})();';

    /* クリップボードを貼り付けてテキストを読み取る（貼り付けたものは削除）/ Paste the clipboard, read the text, remove the pasted items */
    var WORKER_PASTE_CLIPBOARD =
        '(function () {' +
        WORKER_TARGET_DOC +
        WORKER_SELECTION +
        WORKER_COLLECTOR +
        // 貼り付け前に必ず選択解除する（解除しないと、貼り付かなかったときに元の選択を削除してしまう）
        '    saveSelection();' +
        '    targetDoc.selection = null;' +
        '    var pastedItems = [];' +
        '    try {' +
        '        app.executeMenuCommand("pasteInAllArtboard");' +
        // 文字ツールで編集中は TextRange が返る。削除すると選択中の文字自体が消えるため除外する
        '        for (var i = 0; i < targetDoc.selection.length; i++) {' +
        '            if (targetDoc.selection[i].typename !== "TextRange") pastedItems.push(targetDoc.selection[i]);' +
        '        }' +
        '    } catch (err) {}' +
        '    for (var j = 0; j < pastedItems.length; j++) collectFrom(pastedItems[j]);' +
        '    for (var k = 0; k < pastedItems.length; k++) { try { pastedItems[k].remove(); } catch (err) {} }' +
        '    restoreSelection();' +
        '    if (collectedEntries.length === 0) return "NOTX:";' +
        // pasteInAllArtboard はアートボードごとに複製を作るため重複を除く
        '    return "OK:" + encodeURIComponent(joinCollected(true));' +
        '})();';

    /**
     * テキストをクリップボードへ書き戻すワーカーのソースを組み立てる
     * Illustrator には app.system が無いため、一時テキストフレームを作ってコピーし、すぐ削除する
     * @param {string} memoText - コピーするテキスト
     * @returns {string} メインエンジンで評価するソース
     */
    function buildCopyWorkerSource(memoText) {
        return '(function () {' +
            '    var copyText = decodeURIComponent("' + encodeURIComponent(memoText) + '");' +
            // ドキュメントが無ければ一時ドキュメントを作り、コピー後に閉じる
            '    var usingTempDoc = (app.documents.length === 0);' +
            '    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;' +
            WORKER_SELECTION +
            '    if (!usingTempDoc) saveSelection();' +
            // ロック／非表示レイヤーでは textFrames.add() が 8705 で失敗するため一時的に解除する
            '    var editLayer = targetDoc.activeLayer;' +
            '    var layerWasLocked = editLayer.locked;' +
            '    var layerWasHidden = !editLayer.visible;' +
            '    var tempFrame = null;' +
            '    var copySucceeded = false;' +
            '    try {' +
            '        editLayer.locked = false;' +
            '        editLayer.visible = true;' +
            '        tempFrame = editLayer.textFrames.add();' +
            '        tempFrame.contents = copyText;' +
            // app.copy() は黙って無視されることがあるため、再描画を挟んでメニューコマンドでコピーする
            '        app.redraw();' +
            '        app.executeMenuCommand("deselectall");' +
            '        tempFrame.selected = true;' +
            '        app.redraw();' +
            '        app.executeMenuCommand("copy");' +
            '        app.redraw();' + // コピー確定前に削除すると空になる
            '        copySucceeded = true;' +
            '    } catch (err) {}' +
            '    if (tempFrame) { try { tempFrame.remove(); } catch (err) {} }' +
            '    editLayer.locked = layerWasLocked;' +
            '    editLayer.visible = !layerWasHidden;' +
            '    if (usingTempDoc) { try { targetDoc.close(SaveOptions.DONOTSAVECHANGES); } catch (err) {} }' +
            '    else restoreSelection();' +
            '    return copySucceeded ? "OK" : "ERR";' +
            '})();';
    }

    /**
     * ワーカーをメインエンジンで評価し、"マーカー:ペイロード" の結果を状態文字列にして返す
     * @param {string} workerSource - メインエンジンで評価するソース
     * @param {object} statusByMarker - マーカーと状態文字列の対応表
     * @param {function} onComplete - 完了コールバック onComplete(status, text)
     * @returns {void}
     */
    function callMainEngine(workerSource, statusByMarker, onComplete) {
        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = BridgeTalk.appSpecifier;
        bridgeMessage.body = workerSource;
        bridgeMessage.onResult = function (response) {
            var payload = response.body || '';
            var colonIndex = payload.indexOf(':');
            var marker = (colonIndex >= 0) ? payload.substring(0, colonIndex) : payload;
            var status = statusByMarker[marker];
            if (!status) {
                onComplete('error', payload);
                return;
            }
            onComplete(status, (status === 'ok' && colonIndex >= 0) ? decodeURIComponent(payload.substring(colonIndex + 1)) : '');
        };
        bridgeMessage.onError = function (response) {
            onComplete('error', response ? response.body : '');
        };
        bridgeMessage.send();
    }

    /**
     * 選択オブジェクトのテキストを取得する
     * @param {function} onComplete - 完了コールバック onComplete(status, text)
     *   status = 'ok' | 'nodoc' | 'nosel' | 'notextframe' | 'error'
     * @returns {void}
     */
    function fetchSelectedText(onComplete) {
        callMainEngine(WORKER_READ_SELECTION, {
            OK: 'ok',
            NODOC: 'nodoc',
            NOSEL: 'nosel',
            NOTF: 'notextframe'
        }, onComplete);
    }

    /**
     * クリップボードのテキストを取得する
     * @param {function} onComplete - 完了コールバック onComplete(status, text)
     *   status = 'ok' | 'nodoc' | 'notext' | 'error'
     * @returns {void}
     */
    function fetchClipboardText(onComplete) {
        callMainEngine(WORKER_PASTE_CLIPBOARD, {
            OK: 'ok',
            NODOC: 'nodoc',
            NOTX: 'notext'
        }, onComplete);
    }

    /**
     * テキストをクリップボードへコピーする
     * @param {string} memoText - コピーするテキスト
     * @param {function} onComplete - 完了コールバック onComplete(status) status = 'ok' | 'error'
     * @returns {void}
     */
    function copyTextToClipboard(memoText, onComplete) {
        callMainEngine(buildCopyWorkerSource(memoText), { OK: 'ok' }, onComplete);
    }

    // =========================================
    // 二重起動防止 / Prevent duplicate launch
    // =========================================
    /* 既にパレットが開いていれば前面に出して終了する（常駐エンジンなので二重に生成しない）/
       If the palette is already open, bring it to the front and stop (avoid a second instance on this persistent engine) */
    if ($.global.__TextMemoWindow) {
        /* 閉じたパレットの参照は失効していて例外になりうる / a stale reference to a closed palette can throw */
        try {
            if ($.global.__TextMemoWindow.visible) {
                $.global.__TextMemoWindow.active = true; // 前面へ / Bring to front
                return;
            }
        } catch (err) {
            // 参照が失効していれば通常どおり新規作成へ進む / If the reference is stale, fall through and create a new window
        }
    }

    // =========================================
    // セッション保持用変数 / Session state
    // =========================================
    var restoredMemoText = $.global.__TextMemoContent || '';
    var restoredPaletteBounds = $.global.__TextMemoBounds || null;

    /* アクティブドキュメント名（保存フッター・ファイル名用）/ Active document name (used in the footer and file name) */
    var sourceDocumentName = getLabel(LABELS.status.unsaved);
    var initialDocument = getActiveDocument();
    if (initialDocument) {
        sourceDocumentName = initialDocument.fullName ? decodeURI(initialDocument.fullName.name) : initialDocument.name;
    }

    /* 再起動（保存・クリア）由来の close では内容を消さないためのフラグ / Flag so a restart-triggered close does not clear the content */
    var isRestartingScript = false;

    // =========================================
    // UI 構築 / Build UI
    // =========================================
    var initialPaletteBounds = (restoredPaletteBounds && restoredPaletteBounds.length === 4) ? restoredPaletteBounds : undefined;
    var memoPalette = new Window('palette', getLabel(LABELS.dialog.title), initialPaletteBounds, {
        resizeable: true
    });
    memoPalette.opacity = PALETTE_OPACITY;
    setupWindow(memoPalette);
    memoPalette.minimumSize = PALETTE_MIN_SIZE;
    memoPalette.preferredSize.width = PALETTE_INITIAL_WIDTH; // 初回（保存位置がないとき）の幅 / Initial width when there is no saved position
    $.global.__TextMemoWindow = memoPalette; // 二重起動防止用に参照を保持 / Keep a reference so a second launch can detect this window

    /* モード選択（置き換え / 追加）/ Mode selection (replace / append) */
    var modeSelectGroup = memoPalette.add('group');
    modeSelectGroup.orientation = 'row';
    modeSelectGroup.alignment = ['center', 'top'];      // 左右中央 / horizontally centered
    modeSelectGroup.alignChildren = ['left', 'center']; // 天地中央 / vertically centered
    modeSelectGroup.add('statictext', undefined, labelText(LABELS.fieldLabel.mode));
    var appendModeRadio = modeSelectGroup.add('radiobutton', undefined, getLabel(LABELS.radio.append));
    var replaceModeRadio = modeSelectGroup.add('radiobutton', undefined, getLabel(LABELS.radio.replace));
    appendModeRadio.value = true; // デフォルトは追加 / Append by default
    appendModeRadio.helpTip = getLabel(LABELS.tooltip.appendMode);
    replaceModeRadio.helpTip = getLabel(LABELS.tooltip.replaceMode);

    /* 読み込み行（テキスト欄の上）/ Load row (above the text area) */
    var loadButtonRow = memoPalette.add('group');
    loadButtonRow.orientation = 'row';
    loadButtonRow.alignment = ['center', 'top'];
    loadButtonRow.alignChildren = ['center', 'center'];

    var btnLoad = loadButtonRow.add('button', undefined, getLabel(LABELS.button.load));
    var btnRemoveBlanks = loadButtonRow.add('button', undefined, getLabel(LABELS.button.removeBlanks));
    var btnRemoveBreaks = loadButtonRow.add('button', undefined, getLabel(LABELS.button.removeBreaks));
    btnLoad.preferredSize = BUTTON_SIZE;
    btnRemoveBlanks.preferredSize.height = BUTTON_HEIGHT; // 幅はラベルに合わせて自動 / Auto width to fit the label
    btnRemoveBreaks.preferredSize.height = BUTTON_HEIGHT; // 幅はラベルに合わせて自動 / Auto width to fit the label
    btnLoad.helpTip = getLabel(LABELS.tooltip.load);
    btnRemoveBlanks.helpTip = getLabel(LABELS.tooltip.removeBlanks);
    btnRemoveBreaks.helpTip = getLabel(LABELS.tooltip.removeBreaks);

    /* メモ入力テキストエリア / Memo text area */
    var memoTextArea = memoPalette.add('edittext', undefined, restoredMemoText, {
        multiline: true,
        scrolling: true
    });
    memoTextArea.minimumSize.width = MEMO_AREA_MIN_SIZE[0];
    memoTextArea.minimumSize.height = MEMO_AREA_MIN_SIZE[1];
    memoTextArea.alignment = ['fill', 'fill']; // 上下リサイズ対応 / Resize vertically

    /* 保存・コピー・クリア行（テキスト欄の下、3カラム）/ Save / Copy / Clear row (below the text area, 3 columns) */
    var btnRowGroup = memoPalette.add('group');
    btnRowGroup.orientation = 'row';
    btnRowGroup.alignment = ['fill', 'top']; // 行を幅いっぱいに広げて左右に振り分ける / Span full width to push columns left & right
    btnRowGroup.alignChildren = ['fill', 'center'];

    /* 左カラム：保存・すべてをコピー / Left column: Save & Copy All */
    var btnLeftGroup = btnRowGroup.add('group');
    btnLeftGroup.orientation = 'row';
    btnLeftGroup.alignment = ['left', 'center'];
    var btnSave = btnLeftGroup.add('button', undefined, getLabel(LABELS.button.save));
    btnSave.preferredSize = BUTTON_SIZE;
    btnSave.helpTip = getLabel(LABELS.tooltip.save);
    var btnCopyAll = btnLeftGroup.add('button', undefined, getLabel(LABELS.button.copyAll));
    btnCopyAll.preferredSize = WIDE_BUTTON_SIZE;
    btnCopyAll.helpTip = getLabel(LABELS.tooltip.copyAll);

    /* 中央カラム：スペーサー（余白を吸収して左右を振り分ける）/ Center column: spacer that absorbs slack */
    var spacer = btnRowGroup.add('group');
    spacer.alignment = ['fill', 'center'];

    /* 右カラム：クリア / Right column: Clear */
    var btnRightGroup = btnRowGroup.add('group');
    btnRightGroup.orientation = 'row';
    btnRightGroup.alignment = ['right', 'center'];
    var btnClear = btnRightGroup.add('button', undefined, getLabel(LABELS.button.clear));
    btnClear.preferredSize = BUTTON_SIZE;
    btnClear.helpTip = getLabel(LABELS.tooltip.clear);

    /* レイアウトとリサイズ / Layout and resize */
    memoPalette.layout.layout(true);
    memoPalette.onResize = function () {
        memoPalette.layout.resize();
        memoPalette.layout.layout(true);
    };

    /* 表示時に入力欄をアクティブ化 / Focus the field on show */
    memoPalette.onShow = function () {
        memoTextArea.active = true;
    };

    /* ウィンドウ移動時に位置を自動保存 / Auto-save position on move */
    memoPalette.onMove = function () {
        storePaletteBounds();
    };

    /**
     * パレットの位置・サイズをセッションに保存する
     * @returns {void}
     */
    function storePaletteBounds() {
        $.global.__TextMemoBounds = [
            memoPalette.bounds.left, memoPalette.bounds.top,
            memoPalette.bounds.right, memoPalette.bounds.bottom
        ];
    }

    // =========================================
    // ボタンアクション / Button actions
    // =========================================
    /* 選択オブジェクト（無ければクリップボード）のテキストを読み込む / Load text from the selection, falling back to the clipboard */
    btnLoad.onClick = function () {
        fetchSelectedText(function (status, loadedText) {
            if (status === 'nosel' || status === 'nodoc') {
                fetchClipboardText(function (clipboardStatus, clipboardText) {
                    if (clipboardStatus === 'nodoc') {
                        alert(getLabel(LABELS.alert.noDocument));
                        return;
                    }
                    if (clipboardStatus !== 'ok' || !clipboardText) {
                        alert(getLabel(LABELS.alert.clipboardEmpty));
                        return;
                    }
                    applyLoadedText(clipboardText);
                });
                return;
            }
            if (status === 'notextframe') {
                alert(getLabel(LABELS.alert.noTextFrame));
                return;
            }
            if (status !== 'ok') {
                alert(getLabel(LABELS.alert.loadFailed) + loadedText);
                return;
            }
            applyLoadedText(loadedText);
        });
    };

    /* メモをテキストファイルへ保存（パレットは開いたまま）/ Save the memo to a text file (the palette stays open) */
    btnSave.onClick = function () {
        var saveFile = resolveSaveFile();
        if (!saveFile) return;
        if (!writeTextFile(saveFile, buildSaveContent(memoTextArea.text))) {
            alert(getLabel(LABELS.alert.writeFailed));
            return;
        }
        // 保存後もパレットは開いたまま：内容・ウィンドウ位置をそのまま保持する（再起動しない）
        // Keep the palette open after saving: preserve the content and window position as-is (no restart)
        $.global.__TextMemoContent = memoTextArea.text; // フッターは含めず入力内容のみ保持 / Keep the typed memo only (without footer)
        storePaletteBounds();
    };

    /* メモをクリアして再起動 / Clear the memo, then restart */
    btnClear.onClick = function () {
        $.global.__TextMemoContent = '';
        storePaletteBounds();
        restartScript();
    };

    /* すべてをコピー：テキスト欄の内容をクリップボードへ / Copy All: copy the memo to the clipboard */
    btnCopyAll.onClick = function () {
        if (!memoTextArea.text) return;
        copyTextToClipboard(memoTextArea.text, function (status) {
            if (status !== 'ok') alert(getLabel(LABELS.alert.copyFailed));
        });
    };

    /* 空行削除：テキスト欄の空行を除去 / Remove Blanks: strip blank lines from the memo */
    btnRemoveBlanks.onClick = function () {
        setMemoText(removeBlankLines(memoTextArea.text));
    };

    /* 改行削除：テキスト欄の改行をすべて除去して1行にまとめる / Remove Breaks: strip all line breaks from the memo */
    btnRemoveBreaks.onClick = function () {
        setMemoText(removeLineBreaks(memoTextArea.text));
    };

    memoTextArea.onChanging = updateButtonState; // 入力中もリアルタイムに反映 / Update live while typing
    updateButtonState();

    /**
     * テキスト欄の内容を差し替え、セッションとボタンの状態を更新する
     * @param {string} newText - テキスト欄に入れる文字列
     * @returns {void}
     */
    function setMemoText(newText) {
        memoTextArea.text = newText;
        $.global.__TextMemoContent = newText;
        updateButtonState(); // プログラム的な変更では onChanging が発火しないため明示 / onChanging does not fire on programmatic changes
    }

    /**
     * 読み込んだテキストをモードに応じてテキスト欄へ反映する（濁点を NFC へ補正）
     * @param {string} loadedText - 読み込んだテキスト
     * @returns {void}
     */
    function applyLoadedText(loadedText) {
        var normalizedText = normalizeDakuten(loadedText);
        if (appendModeRadio.value) {
            // 既存テキストとの間に空行を1つ挟む / Insert one blank line between existing and new text
            setMemoText(memoTextArea.text + (memoTextArea.text ? '\n\n' : '') + normalizedText);
        } else {
            setMemoText(normalizedText);
        }
    }

    /**
     * テキストの有無・空行の有無でボタンの活性を切り替える
     * @returns {void}
     */
    function updateButtonState() {
        var memoText = memoTextArea.text;
        var hasText = (memoText.length > 0);
        btnClear.enabled = hasText;
        btnCopyAll.enabled = hasText;
        btnRemoveBlanks.enabled = hasText && hasBlankLines(memoText); // 空行が無ければディム / Dim when there are no blank lines
        btnRemoveBreaks.enabled = hasText && hasLineBreaks(memoText); // 改行が無ければディム / Dim when there are no line breaks
    }

    // =========================================
    // キー操作 & 閉じる / Keys & close
    // =========================================
    /* Escape でパレットを閉じる / Close the palette with Escape */
    memoPalette.addEventListener('keydown', function (keyEvent) {
        if (keyEvent.keyName === 'Escape') memoPalette.close();
    });

    /* 閉じたときの内容を CLEAR_ON_CLOSE で制御（再起動時は触れない）。ウィンドウ位置は常に保持 /
       On close, control the content via CLEAR_ON_CLOSE (untouched on restart); always keep the window position */
    memoPalette.onClose = function () {
        if (!isRestartingScript) {
            $.global.__TextMemoContent = CLEAR_ON_CLOSE ? '' : memoTextArea.text;
        }
        storePaletteBounds();
        $.global.__TextMemoWindow = null; // 参照を解放（次回起動で新規作成できるように）/ Release the reference so the next launch creates a fresh window
        return true;
    };

    /* 表示（既存位置がない場合のみ中央）/ Show the palette (center only when no saved position) */
    if (!initialPaletteBounds) memoPalette.center();
    addVersionFooter(memoPalette);
    memoPalette.show();

    // =========================================
    // テキスト整形 / Text utilities
    // =========================================
    /**
     * 改行を LF に正規化して行の配列へ分割する
     * @param {string} sourceText - 対象のテキスト
     * @returns {string[]} 行の配列
     */
    function splitLines(sourceText) {
        return sourceText.replace(/\r\n|\r/g, '\n').split('\n');
    }

    /**
     * 空行（空白のみの行を含む）かどうかを判定する
     * @param {string} line - 1行分のテキスト
     * @returns {boolean} 空行なら true
     */
    function isBlankLine(line) {
        return line.replace(/^\s+|\s+$/g, '') === '';
    }

    /**
     * 空行（空白のみの行を含む）を除去する
     * @param {string} sourceText - 対象のテキスト
     * @returns {string} 空行を除いたテキスト
     */
    function removeBlankLines(sourceText) {
        var lines = splitLines(sourceText);
        var keptLines = [];
        for (var i = 0; i < lines.length; i++) {
            if (!isBlankLine(lines[i])) keptLines.push(lines[i]);
        }
        return keptLines.join('\n');
    }

    /**
     * 空行が1つでもあるか判定する
     * @param {string} sourceText - 対象のテキスト
     * @returns {boolean} 空行があれば true
     */
    function hasBlankLines(sourceText) {
        var lines = splitLines(sourceText);
        for (var i = 0; i < lines.length; i++) {
            if (isBlankLine(lines[i])) return true;
        }
        return false;
    }

    /**
     * 改行をすべて除去して1行にまとめる
     * @param {string} sourceText - 対象のテキスト
     * @returns {string} 1行にまとめたテキスト
     */
    function removeLineBreaks(sourceText) {
        return sourceText.replace(/\r\n|\r|\n/g, '');
    }

    /**
     * 改行が1つでもあるか判定する
     * @param {string} sourceText - 対象のテキスト
     * @returns {boolean} 改行があれば true
     */
    function hasLineBreaks(sourceText) {
        return /\r|\n/.test(sourceText);
    }

    /**
     * 分解済みの濁点・半濁点（NFD）を合成形（NFC）へ補正する
     * macOS のファイル名やクリップボード由来の「か゛」を「が」に直す
     * @param {string} rawText - 補正前のテキスト
     * @returns {string} 補正後のテキスト
     */
    function normalizeDakuten(rawText) {
        return decodeURIComponent(composeDakutenOnEncoded(encodeURIComponent(rawText)));
    }

    /**
     * パーセントエンコード列上で分解済み濁点を合成済みへ置換する
     * @param {string} encodedText - encodeURIComponent 済みのテキスト
     * @returns {string} 置換後のエンコード列
     */
    function composeDakutenOnEncoded(encodedText) {
        var decomposedForms = ['%E3%81%86%E3%82%99', '%E3%81%8B%E3%82%99', '%E3%81%8D%E3%82%99', '%E3%81%8F%E3%82%99', '%E3%81%91%E3%82%99', '%E3%81%93%E3%82%99',
            '%E3%81%95%E3%82%99', '%E3%81%97%E3%82%99', '%E3%81%99%E3%82%99', '%E3%81%9B%E3%82%99', '%E3%81%9D%E3%82%99', '%E3%81%9F%E3%82%99',
            '%E3%81%A1%E3%82%99', '%E3%81%A4%E3%82%99', '%E3%81%A6%E3%82%99', '%E3%81%A8%E3%82%99', '%E3%81%AF%E3%82%99', '%E3%81%AF%E3%82%9A',
            '%E3%81%B2%E3%82%99', '%E3%81%B2%E3%82%9A', '%E3%81%B5%E3%82%99', '%E3%81%B5%E3%82%9A', '%E3%81%B8%E3%82%99', '%E3%81%B8%E3%82%9A',
            '%E3%81%BB%E3%82%99', '%E3%81%BB%E3%82%9A', '%E3%82%A6%E3%82%99', '%E3%82%AB%E3%82%99', '%E3%82%AD%E3%82%99', '%E3%82%AF%E3%82%99',
            '%E3%82%B1%E3%82%99', '%E3%82%B3%E3%82%99', '%E3%82%B5%E3%82%99', '%E3%82%B7%E3%82%99', '%E3%82%B9%E3%82%99', '%E3%82%BB%E3%82%99',
            '%E3%82%BD%E3%82%99', '%E3%82%BF%E3%82%99', '%E3%83%81%E3%82%99', '%E3%83%84%E3%82%99', '%E3%83%86%E3%82%99', '%E3%83%88%E3%82%99',
            '%E3%83%8F%E3%82%99', '%E3%83%8F%E3%82%9A', '%E3%83%92%E3%82%99', '%E3%83%92%E3%82%9A', '%E3%83%95%E3%82%99', '%E3%83%95%E3%82%9A',
            '%E3%83%98%E3%82%99', '%E3%83%98%E3%82%9A', '%E3%83%9B%E3%82%99', '%E3%83%9B%E3%82%9A'
        ];
        var composedForms = ['%E3%82%94', '%E3%81%8C', '%E3%81%8E', '%E3%81%90', '%E3%81%92', '%E3%81%94', '%E3%81%96', '%E3%81%98', '%E3%81%9A', '%E3%81%9C', '%E3%81%9E',
            '%E3%81%A0', '%E3%81%A2', '%E3%81%A5', '%E3%81%A7', '%E3%81%A9', '%E3%81%B0', '%E3%81%B1', '%E3%81%B3', '%E3%81%B4', '%E3%81%B6', '%E3%81%B7',
            '%E3%81%B9', '%E3%81%BA', '%E3%81%BC', '%E3%81%BD', '%E3%83%B4', '%E3%82%AC', '%E3%82%AE', '%E3%82%B0', '%E3%82%B2', '%E3%82%B4', '%E3%82%B6',
            '%E3%82%B8', '%E3%82%BA', '%E3%82%BC', '%E3%82%BE', '%E3%83%80', '%E3%83%82', '%E3%83%85', '%E3%83%87', '%E3%83%89', '%E3%83%90', '%E3%83%91',
            '%E3%83%93', '%E3%83%94', '%E3%83%96', '%E3%83%97'
        ];
        for (var i = 0; i < decomposedForms.length; i++) {
            encodedText = encodedText.replace(new RegExp(decomposedForms[i], 'ig'), composedForms[i]);
        }
        return encodedText;
    }

    // =========================================
    // 保存 / Save
    // =========================================
    /**
     * SAVE_LOCATION_MODE に従って保存先ファイルを決める
     * @returns {File|null} 保存先ファイル（保存ダイアログでキャンセルされたときは null）
     */
    function resolveSaveFile() {
        if (SAVE_LOCATION_MODE === 'desktop') {
            // B: 常にデスクトップへ memo-<ドキュメント名>-<yyyymmdd>.txt（日本語名対応で encodeURI）
            var fileName = 'memo-' + getDocumentBaseName() + '-' + getTimeStamp().substring(0, 8) + '.txt';
            return new File(Folder.desktop + '/' + encodeURI(fileName));
        }
        // A: 保存ダイアログで場所と名前を選ぶ
        var suggestedName = getDocumentBaseName() + '_' + getTimeStamp() + '.txt';
        var chosenFile = File.saveDialog(getLabel(LABELS.alert.savePrompt), suggestedName);
        if (!chosenFile) return null;
        return /\.txt$/i.test(chosenFile.name) ? chosenFile : new File(chosenFile.fsName + '.txt');
    }

    /**
     * 保存本文（フッター付き、改行を LF に正規化）を組み立てる
     * @param {string} memoText - テキスト欄の内容
     * @returns {string} ファイルへ書き込む文字列
     */
    function buildSaveContent(memoText) {
        var bodyText = memoText.replace(/\r\n|\r/g, '\n');
        var footerTimeStamp = getTimeStamp().replace(/(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})/, '$1-$2-$3 $4:$5:$6');
        return bodyText + '\n\n---\n[' + labelValueText(LABELS.footer.source, sourceDocumentName) + ']\n[' +
            labelValueText(LABELS.footer.savedAt, footerTimeStamp) + ']\n';
    }

    /**
     * UTF-8 でテキストファイルへ書き込む
     * @param {File} targetFile - 書き込み先ファイル
     * @param {string} fileContent - 書き込む文字列
     * @returns {boolean} 書き込めたら true
     */
    function writeTextFile(targetFile, fileContent) {
        targetFile.encoding = 'UTF-8';
        if (!targetFile.open('w')) return false;
        var writeSucceeded = targetFile.write(fileContent);
        targetFile.close();
        return writeSucceeded;
    }

    /**
     * タイムスタンプ YYYYMMDD_HHMMSS を返す
     * @returns {string} タイムスタンプ文字列
     */
    function getTimeStamp() {
        var currentDate = new Date();

        /**
         * 2桁になるよう 0 を補う
         * @param {number} numberValue - 対象の数値
         * @returns {string} 2桁の文字列
         */
        function padZero(numberValue) {
            return ('0' + numberValue).slice(-2);
        }
        return currentDate.getFullYear() + padZero(currentDate.getMonth() + 1) + padZero(currentDate.getDate()) + '_' +
            padZero(currentDate.getHours()) + padZero(currentDate.getMinutes()) + padZero(currentDate.getSeconds());
    }

    // =========================================
    // ドキュメント / Document helpers
    // =========================================
    /**
     * アクティブドキュメントを安全に取得する（起動直後・パレット表示前向け）
     * パレット表示中はこのエンジンの app が DOM 接続を失うため、DOM に触る処理は
     * fetchSelectedText などメインエンジン側へ委譲すること
     * @returns {Document|null} アクティブドキュメント（取得できなければ null）
     */
    function getActiveDocument() {
        try {
            return (app.documents.length > 0) ? app.activeDocument : null;
        } catch (err) {
            return null;
        }
    }

    /**
     * 保存ファイルの基底名（拡張子なし）を返す
     * パレット表示中は DOM 取得が失敗しうるため、起動時に控えた名前を使う
     * @returns {string} ファイルの基底名
     */
    function getDocumentBaseName() {
        if (sourceDocumentName === getLabel(LABELS.status.unsaved)) return 'Untitled';
        return sourceDocumentName.replace(/\.[^\.]+$/, '');
    }

    /**
     * スクリプト自身をメインエンジンで再評価して再起動する
     * @returns {void}
     */
    function restartScript() {
        var scriptFile = new File($.fileName);
        var bridgeMessage = new BridgeTalk();
        bridgeMessage.target = BridgeTalk.appSpecifier;
        bridgeMessage.body = '$.evalFile("' + scriptFile.fsName.replace(/\\/g, '\\\\') + '");';
        bridgeMessage.send(100);
        isRestartingScript = true; // この close では内容を消さない / This close must not clear the content
        memoPalette.close();
    }

})();
