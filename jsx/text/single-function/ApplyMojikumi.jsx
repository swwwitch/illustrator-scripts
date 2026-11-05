#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ApplyMojikumi
Name=文字組みアキ量設定を適用
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.1
Release-Date=2026-06-23
Description-BEGIN
選択したテキストフレーム（グループ内も再帰的に処理）の全段落に、文字組みアキ量設定をダイアログボックスのラジオボタンから適用します。
ラジオボタンを選ぶと、選択中のテキストへすぐに反映されます。組み込みの8種類に加えて、ドキュメントのカスタムのセットも選べます。
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したテキストフレーム（グループ内も再帰的に処理）の全段落に、文字組みアキ量設定をダイアログボックスのラジオボタンから適用します。
ラジオボタンを選ぶと、選択中のテキストへすぐに反映されます。組み込みの8種類に加えて、ドキュメントのカスタムのセットも選べます。

note記事も参照してください。
https://note.com/dtp_tranist/n/nb9c089e378d2

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyMojikumi";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "文字組みアキ量設定を適用"; /* 内容（README なし） / what it does (no README) */

var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb9c089e378d2"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/single-function/ApplyMojikumi.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 組み込みの文字組みプリセット（mojikumiSet のインデックス。-1 は「なし」）。組み込みの要素は内部 ID の文字列なので、表示名はここに持つ
       Built-in mojikumi presets (mojikumiSet index; -1 = none). Built-in entries are internal ID strings, so the names live here */
    var mojikumiPresets = [
        { mojikumiIndex: -1, labelText: "なし" },
        { mojikumiIndex: 0, labelText: "行末約物全角/半角" },
        { mojikumiIndex: 1, labelText: "約物半角" },
        { mojikumiIndex: 2, labelText: "行末約物半角" },
        { mojikumiIndex: 3, labelText: "行末約物全角" },
        { mojikumiIndex: 4, labelText: "約物全角" },
        { mojikumiIndex: 5, labelText: "ツメ組み" },
        { mojikumiIndex: 6, labelText: "ベタ組み" }
    ];

    /* 組み込みの数（mojikumiSet のインデックス7以降はカスタムのセットで、要素の文字列がそのまま名前）
       Number of built-in sets (index 7 and later are custom sets whose string is the name) */
    var MOJIKUMI_BUILTIN_COUNT = 7;

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 前提チェック：ドキュメントが開かれているか / Precondition: a document is open */
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    /* 前提チェック：オブジェクトが選択されているか / Precondition: something is selected */
    if (doc.selection.length === 0) {
        alert("テキストオブジェクトを選択してください。");
        return;
    }

    /**
     * 組み込みのプリセットのあとに、ドキュメントのカスタムのセットを足した選択肢を返す
     * @returns {Array<{mojikumiIndex: number, labelText: string}>} 選択肢
     */
    function buildMojikumiChoices() {
        var choices = mojikumiPresets.slice(0);
        var mojikumiSets = doc.mojikumiSet;
        for (var i = MOJIKUMI_BUILTIN_COUNT; i < mojikumiSets.length; i++) {
            choices.push({ mojikumiIndex: i, labelText: String(mojikumiSets[i]) });
        }
        return choices;
    }

    /**
     * テキストフレーム内の全段落に適用する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {*} mojikumiValue - 代入する値（mojikumiSet の要素、または "なし"）
     * @returns {void}
     */
    function applyMojikumiToTextFrame(textFrame, mojikumiValue) {
        var paragraphs = textFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            /* 末尾の空の段落は取り出すと例外になる / A trailing empty paragraph throws when accessed */
            try { paragraphs[i].paragraphAttributes.mojikumi = mojikumiValue; } catch (e) { }
        }
    }

    /**
     * 種類で振り分ける（TextFrame は直接、GroupItem は再帰）
     * @param {PageItem} item - 対象のアイテム
     * @param {*} mojikumiValue - 代入する値
     * @returns {void}
     */
    function applyMojikumiToItem(item, mojikumiValue) {
        if (item.typename === "TextFrame") {
            applyMojikumiToTextFrame(item, mojikumiValue);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) applyMojikumiToItem(item.pageItems[i], mojikumiValue);
        }
    }

    /**
     * 選択全体に適用して再描画する
     * @param {number} mojikumiIndex - mojikumiSet のインデックス（-1 は「なし」）
     * @returns {void}
     */
    function applyMojikumiToSelection(mojikumiIndex) {
        /* 値を1回だけ解決して再帰へ渡す / Resolve the value once, then pass it down */
        var mojikumiValue = (mojikumiIndex === -1) ? "なし" : doc.mojikumiSet[mojikumiIndex];
        for (var i = 0; i < doc.selection.length; i++) {
            applyMojikumiToItem(doc.selection[i], mojikumiValue);
        }
        app.redraw();
    }

    /**
     * 選択の最初のテキストフレームの、先頭段落の文字組みアキ量設定のインデックスを返す
     * 読み取りは組み込みなら内部 ID、カスタムなら名前の文字列で、mojikumiSet の要素と一致する。
     * 「行末約物半角」(2) は「約物半角」(1) と同じ ID で読み戻るので 1 になる
     * @returns {number} mojikumiSet のインデックス（「なし」・読めないときは -1）
     */
    function readSelectedMojikumiIndex() {
        for (var i = 0; i < doc.selection.length; i++) {
            if (doc.selection[i].typename !== "TextFrame") continue;
            var mojikumiText;
            try {
                mojikumiText = String(doc.selection[i].paragraphs[0].paragraphAttributes.mojikumi);
            } catch (e) {
                /* 「なし」は読むと例外 / Reading "None" throws */
                return -1;
            }
            var mojikumiSets = doc.mojikumiSet;
            for (var j = 0; j < mojikumiSets.length; j++) {
                if (String(mojikumiSets[j]) === mojikumiText) return j;
            }
            return -1;
        }
        return -1;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================


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
    /**
     * ダイアログを組み立てて表示する
     * @returns {void}
     */
    function showMojikumiDialog() {
        var dialog = new Window("dialog", "文字組みアキ量設定");
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        /* プリセット選択パネル / Preset selection panel */
        var presetPanel = dialog.add("panel", undefined, "設定を選択");
        presetPanel.orientation = "column";
        presetPanel.alignChildren = "left";
        presetPanel.alignment = "fill";
        presetPanel.margins = [16, 20, 16, 12];
        presetPanel.spacing = 6;

        /* 選択肢ごとにラジオボタンを生成（同じ親に置いて排他にする）/ One radio button per choice (same parent, so they are exclusive) */
        var mojikumiChoices = buildMojikumiChoices();
        var selectedIndex = readSelectedMojikumiIndex();
        var mojikumiRadioButtons = [];
        for (var i = 0; i < mojikumiChoices.length; i++) {
            var radioButton = presetPanel.add("radiobutton", undefined, mojikumiChoices[i].labelText);
            radioButton.mojikumiIndex = mojikumiChoices[i].mojikumiIndex;
            radioButton.value = (radioButton.mojikumiIndex === selectedIndex);
            radioButton.onClick = function () {
                applyMojikumiToSelection(this.mojikumiIndex);
            };
            mojikumiRadioButtons.push(radioButton);
        }

        /* 閉じるボタン / Close button */
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["left", "center"];
        buttonGroup.alignment = "right";
        buttonGroup.spacing = 8;

        var closeButton = buttonGroup.add("button", undefined, "閉じる");
        closeButton.onClick = function () {
            dialog.close();
        };

        dialog.center();
        addVersionFooter(dialog);
        dialog.show();
    }

    showMojikumiDialog();

})();
