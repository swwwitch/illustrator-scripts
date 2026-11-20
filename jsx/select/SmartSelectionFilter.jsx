#target illustrator
#targetengine "SmartSelectionFilterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SmartSelectionFilter
Name=選択オブジェクトを条件で絞り込み再選択
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.6
Release-Date=2026-05-02
Description-BEGIN
選択中のオブジェクトから、条件に合うテキストやパスだけを残して選択し直します。
文字種・行揃え・使用フォント・パスの開閉・塗りや線の状態を条件にでき、変更はダイアログを閉じずにカンバスへ反映されます。

Reselects only the text frames and paths in the current selection that match the conditions you set.
Text kind, justification, font, open/closed paths and fill/stroke state can all be used as conditions, and changes are reflected on the canvas without closing the dialog.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択中のオブジェクトから、条件に合うテキストやパスだけを残して選択し直します。
文字種・行揃え・使用フォント・パスの開閉・塗りや線の状態を条件にでき、変更はダイアログを閉じずにカンバスへ反映されます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSelectionFilter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n6203a03662d8

### Overview

Reselects only the text frames and paths in the current selection that match the conditions you set.
Text kind, justification, font, open/closed paths and fill/stroke state can all be used as conditions, and changes are reflected on the canvas without closing the dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSelectionFilter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartSelectionFilter";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "選択オブジェクトを条件で絞り込み再選択"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSelectionFilter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSelectionFilter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6203a03662d8"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SmartSelectionFilter.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    /* UI／エラーメッセージの日英ラベル定義 / Japanese-English UI and error message label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "選択オブジェクトの絞り込み",
            en: "Filter Selected Objects"
        },
        textPanel: {
            ja: "テキスト",
            en: "Text"
        },
        textObject: {
            ja: "ポイント文字",
            en: "Point Text"
        },
        areaText: {
            ja: "エリア内文字",
            en: "Area Text"
        },
        pathText: {
            ja: "パス上文字",
            en: "Path Text"
        },
        alignmentPanel: {
            ja: "行揃え",
            en: "Alignment"
        },
        tipNonSelectedNone: { ja: "絞り込みから外れたオブジェクトは、そのまま表示しておきます。", en: "Leaves the objects that fall outside the filter as they are." },
        tipNonSelectedHide: { ja: "絞り込みから外れたオブジェクトを一時的に隠します。OKで確定、キャンセルで元に戻ります。", en: "Temporarily hides the objects outside the filter. OK keeps it, Cancel restores them." },
        tipNonSelectedOpacity: { ja: "絞り込みから外れたオブジェクトの不透明度を下げて、対象を見分けやすくします。", en: "Dims the objects outside the filter so the matches stand out." },
        tipNonSelectedSlider: { ja: "外れたオブジェクトに使う不透明度（％）です。", en: "Opacity used for the objects outside the filter." },
        tipSimpleText: { ja: "テキストをまとめて選びます（ポイント文字・エリア内文字・パス上文字のすべて）。", en: "Selects all text at once: point, area and path text." },
        tipSimpleStrokeOnly: { ja: "線だけが設定されたパスを選びます（塗りなし）。", en: "Selects paths that have a stroke but no fill." },
        tipSimpleFillOnlyPath: { ja: "塗りだけが設定されたパスを選びます（線なし）。", en: "Selects paths that have a fill but no stroke." },
        tipPreviewButton: { ja: "いまの絞り込み結果を画面で確認します。もう一度押すと元の表示に戻ります。", en: "Shows the current filter result on the canvas. Press again to restore the view." },
        fontPanel: {
            ja: "フォント",
            en: "Font"
        },
        noFontItems: {
            ja: "対象フォントなし",
            en: "No target fonts"
        },
        alignLeft: {
            ja: "左揃え",
            en: "Left"
        },
        alignCenter: {
            ja: "中央揃え",
            en: "Center"
        },
        alignRight: {
            ja: "右揃え",
            en: "Right"
        },
        pathPanel: {
            ja: "パス",
            en: "Path"
        },
        openPath: {
            ja: "オープンパス",
            en: "Open Path"
        },
        horizontalLine: {
            ja: "水平線",
            en: "Horizontal Line"
        },
        verticalLine: {
            ja: "垂直線",
            en: "Vertical Line"
        },
        closedPathPanel: {
            ja: "クローズパス",
            en: "Closed Path"
        },
        closedPathFillOnly: {
            ja: "塗りのみ",
            en: "Fill Only"
        },
        closedPathStrokeOnly: {
            ja: "線のみ",
            en: "Stroke Only"
        },
        closedPathFillAndStroke: {
            ja: "塗りと線",
            en: "Fill and Stroke"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        btnOutlineOn: {
            ja: "アウトライン表示",
            en: "Outline"
        },
        btnOutlineOff: {
            ja: "プレビュー表示",
            en: "Preview"
        },
        nonSelectedPanel: {
            ja: "対象外オブジェクト",
            en: "Non-target Objects"
        },
        nonSelectedNone: {
            ja: "何もしない",
            en: "Do Nothing"
        },
        nonSelectedHide: {
            ja: "非表示",
            en: "Hide"
        },
        nonSelectedOpacity: {
            ja: "不透明度",
            en: "Opacity"
        },
        simplePanel: {
            ja: "まとめて選択",
            en: "Quick Select"
        },
        simpleText: {
            ja: "テキスト",
            en: "Text"
        },
        simpleStrokeOnly: {
            ja: "ケイのみ",
            en: "Stroke Only"
        },
        simpleFillOnlyPath: {
            ja: "塗りのみ",
            en: "Fill Only"
        },
        scopePanel: { ja: "対象スコープ", en: "Selection Scope" },
        scopeSelectedOnly: { ja: "選択直下のみ", en: "Selected objects only" },
        scopeIncludeGroupItems: { ja: "グループ内も対象に含める", en: "Include objects inside groups" },
        tipScope: { ja: "グループ内のオブジェクトを対象に含めるかを指定します。", en: "Choose whether to include objects inside groups." },
        noDocumentError: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        noSelectionError: {
            ja: "オブジェクトが選択されていません。",
            en: "No object is selected."
        }
    };

    /* エラー表示ヘルパー / Error display helper */
    function showError(key) {
        alert(getLabel(key));
    }

    /* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
    function labelWithCount(key, count) {
        if (uiLang === "ja") {
            return getLabel(key) + "（" + count + "）";
        }
        return getLabel(key) + " (" + count + ")";
    }

    /* 値付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with value (full-width JA parentheses, half-width EN parentheses) */
    function labelWithValue(key, valueText) {
        if (uiLang === "ja") {
            return getLabel(key) + "（" + valueText + "）";
        }
        return getLabel(key) + " (" + valueText + ")";
    }

    /* 任意テキストへの値付け（日本語は全角括弧、英語は半角括弧）/ Append value to arbitrary text (full-width JA parentheses, half-width EN parentheses) */
    function appendValueText(baseText, valueText) {
        if (uiLang === "ja") {
            return baseText + "（" + valueText + "）";
        }
        return baseText + " (" + valueText + ")";
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

    var LINE_ORIENTATION_EPSILON = 0.01;
    var DEFAULT_NON_SELECTED_OPACITY = 30;

    /* 選択を配列として退避 / Store selection as an array */
    function copySelection(selection) {
        var copiedSelection = [];
        for (var selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
            copiedSelection.push(selection[selectionIndex]);
        }
        return copiedSelection;
    }

    /* 退避した選択状態を復元 / Restore stored selection */
    function restoreSelection(items) {
        app.selection = null;
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            try {
                items[itemIndex].selected = true;
            } catch (e) { }
        }
    }

    /* チェックボックス変更時の処理を設定 / Bind checkbox update handlers */
    function bindCheckboxPreviewUpdates(checkboxes, onUpdate) {
        function bindOneCheckbox(sourceCheckbox) {
            sourceCheckbox._optionClickRequested = false;

            sourceCheckbox.addEventListener("mousedown", function (event) {
                sourceCheckbox._optionClickRequested = event.altKey === true;
            });

            sourceCheckbox.onClick = function () {
                if (sourceCheckbox._optionClickRequested) {
                    for (var checkboxIndex = 0; checkboxIndex < checkboxes.length; checkboxIndex++) {
                        checkboxes[checkboxIndex].value = sourceCheckbox.value;
                        checkboxes[checkboxIndex]._optionClickRequested = false;
                    }
                }

                sourceCheckbox._optionClickRequested = false;
                onUpdate();
            };
        }

        for (var checkboxIndex = 0; checkboxIndex < checkboxes.length; checkboxIndex++) {
            bindOneCheckbox(checkboxes[checkboxIndex]);
        }
    }

    /* ポイント文字かどうかを判定 / Check whether a text frame is point text */
    function isPointTextFrame(textFrame) {
        try {
            return textFrame.kind === TextType.POINTTEXT;
        } catch (e) {
            return false;
        }
    }

    /* エリア内文字かどうかを判定 / Check whether a text frame is area text */
    function isAreaTextFrame(textFrame) {
        try {
            return textFrame.kind === TextType.AREATEXT;
        } catch (e) {
            return false;
        }
    }

    /* パス上文字かどうかを判定 / Check whether a text frame is path text */
    function isPathTextFrame(textFrame) {
        try {
            return textFrame.kind === TextType.PATHTEXT;
        } catch (e) {
            return false;
        }
    }

    /* テキストの行揃えを取得 / Get text alignment */
    function getTextAlignmentType(textFrame) {
        try {
            var justification = textFrame.textRange.paragraphAttributes.justification;
            if (justification === Justification.LEFT) {
                return "left";
            }
            if (justification === Justification.CENTER) {
                return "center";
            }
            if (justification === Justification.RIGHT) {
                return "right";
            }
        } catch (e) { }

        return "other";
    }

    /* テキストのフォント名を取得 / Get text font display name (family + style) */
    function getTextFontName(textFrame) {
        try {
            var font = textFrame.textRange.characterAttributes.textFont;
            var family = "";
            var style = "";
            try { family = font.family || ""; } catch (e1) { }
            try { style = font.style || ""; } catch (e2) { }
            if (family && style) return family + " " + style;
            if (family) return family;
            try { return font.name; } catch (e3) { }
        } catch (e) { }

        return "";
    }

    /* フォント件数を追加 / Add font count */
    function addFontCount(fontCounts, fontName) {
        if (!fontName) {
            return;
        }
        if (!fontCounts[fontName]) {
            fontCounts[fontName] = 0;
        }
        fontCounts[fontName]++;
    }

    /* フォント名をソートして配列化 / Sort font names into an array */
    function getSortedFontNames(fontCounts) {
        var fontNames = [];
        for (var fontName in fontCounts) {
            if (fontCounts.hasOwnProperty(fontName)) {
                fontNames.push(fontName);
            }
        }
        fontNames.sort();
        return fontNames;
    }

    /* オープンパスの線方向を判定 / Detect line orientation of an open path
       - "horizontal": geometricBounds の高さ ≈ 0
       - "vertical":   geometricBounds の幅 ≈ 0
       - "other":      それ以外
    */
    function getOpenPathLineOrientation(pathItem) {
        try {
            var geometricBounds = pathItem.geometricBounds; // [left, top, right, bottom]
            if (!geometricBounds || geometricBounds.length < 4) return "other";
            var width = Math.abs(geometricBounds[2] - geometricBounds[0]);
            var height = Math.abs(geometricBounds[1] - geometricBounds[3]);
            if (height <= LINE_ORIENTATION_EPSILON && width > LINE_ORIENTATION_EPSILON) return "horizontal";
            if (width <= LINE_ORIENTATION_EPSILON && height > LINE_ORIENTATION_EPSILON) return "vertical";
        } catch (e) { }
        return "other";
    }

    /* クローズパスの塗りと線の状態を取得 / Get fill and stroke state of a closed path */
    function getClosedPathPaintType(pathItem) {
        var hasFill = false;
        var hasStroke = false;

        try {
            hasFill = pathItem.filled === true;
        } catch (e1) { }

        try {
            hasStroke = pathItem.stroked === true;
        } catch (e2) { }

        if (hasFill && hasStroke) {
            return "fillAndStroke";
        }
        if (hasFill) {
            return "fillOnly";
        }
        if (hasStroke) {
            return "strokeOnly";
        }

        return "none";
    }

    /**
     * クリップグループのマスク用パスかどうかを判定する
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {boolean} マスク用パスなら true
     */
    function isClippingPathItem(item) {
        try {
            return item.typename === "PathItem" && item.clipping === true;
        } catch (e) {
            return false;
        }
    }

    /* 選択内の各種オブジェクト数をカウント / Count object types in selection */
    function countSelectionItems(selection, includeGroupItems) {
        var counts = {
            text: 0,
            areaText: 0,
            pathText: 0,
            alignLeft: 0,
            alignCenter: 0,
            alignRight: 0,
            fonts: {},
            openPath: 0,
            horizontalLine: 0,
            verticalLine: 0,
            closedPathFillOnly: 0,
            closedPathStrokeOnly: 0,
            closedPathFillAndStroke: 0
        };

        countItemsRecursive(selection, counts, includeGroupItems);
        return counts;
    }

    /* 選択内の子要素も含めて再帰的にカウント / Count selected items recursively including children */
    function countItemsRecursive(items, counts, includeGroupItems) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var item = items[itemIndex];

            if (isClippingPathItem(item)) {
                continue;
            }

            if (item.typename === "TextFrame") {
                if (isPointTextFrame(item)) {
                    counts.text++;
                } else if (isAreaTextFrame(item)) {
                    counts.areaText++;
                } else if (isPathTextFrame(item)) {
                    counts.pathText++;
                }

                var alignmentType = getTextAlignmentType(item);
                if (alignmentType === "left") {
                    counts.alignLeft++;
                } else if (alignmentType === "center") {
                    counts.alignCenter++;
                } else if (alignmentType === "right") {
                    counts.alignRight++;
                }

                addFontCount(counts.fonts, getTextFontName(item));
            } else if (item.typename === "PathItem") {
                if (item.closed === true) {
                    var paintType = getClosedPathPaintType(item);
                    if (paintType === "fillOnly") {
                        counts.closedPathFillOnly++;
                    } else if (paintType === "strokeOnly") {
                        counts.closedPathStrokeOnly++;
                    } else if (paintType === "fillAndStroke") {
                        counts.closedPathFillAndStroke++;
                    }
                } else {
                    counts.openPath++;
                    var orientation = getOpenPathLineOrientation(item);
                    if (orientation === "horizontal") {
                        counts.horizontalLine++;
                    } else if (orientation === "vertical") {
                        counts.verticalLine++;
                    }
                }
            } else if (item.typename === "GroupItem" && includeGroupItems) {
                countItemsRecursive(item.pageItems, counts, includeGroupItems);
            }
        }
    }

    /* UIから選択条件を取得 / Read filter options from UI */
    function readFilterOptions(dialogUi) {
        return {
            keepText: dialogUi.cbText.value,
            keepAreaText: dialogUi.cbAreaText.value,
            keepPathText: dialogUi.cbPathText.value,
            keepAlignLeft: dialogUi.cbAlignLeft.value,
            keepAlignCenter: dialogUi.cbAlignCenter.value,
            keepAlignRight: dialogUi.cbAlignRight.value,
            selectedFonts: readSelectedFonts(dialogUi.fontCheckboxes),
            keepOpenPath: dialogUi.cbOpenPath.value,
            keepHorizontalLine: dialogUi.cbHorizontalLine.value,
            keepVerticalLine: dialogUi.cbVerticalLine.value,
            keepClosedPathFillOnly: dialogUi.cbClosedPathFillOnly.value,
            keepClosedPathStrokeOnly: dialogUi.cbClosedPathStrokeOnly.value,
            keepClosedPathFillAndStroke: dialogUi.cbClosedPathFillAndStroke.value,
            includeGroupItems: dialogUi.rbScopeIncludeGroupItems.value
        };
    }

    /* チェックされたフォント名を取得 / Read selected font names */
    function readSelectedFonts(fontCheckboxes) {
        var selectedFonts = {};
        for (var fontCheckboxIndex = 0; fontCheckboxIndex < fontCheckboxes.length; fontCheckboxIndex++) {
            if (fontCheckboxes[fontCheckboxIndex].checkbox.value) {
                selectedFonts[fontCheckboxes[fontCheckboxIndex].fontName] = true;
            }
        }
        return selectedFonts;
    }

    /* 1項目が残す対象かどうかを判定 / Check whether an item should remain selected */
    function shouldKeepItem(item, filterOptions) {
        /* マスク用パス自体は選択対象から除外 / Exclude the clipping mask path itself */
        if (isClippingPathItem(item)) {
            return false;
        }

        if (item.typename === "TextFrame") {
            var keepByTextType = false;
            if (isPointTextFrame(item)) {
                keepByTextType = filterOptions.keepText;
            } else if (isAreaTextFrame(item)) {
                keepByTextType = filterOptions.keepAreaText;
            } else if (isPathTextFrame(item)) {
                keepByTextType = filterOptions.keepPathText;
            }

            if (!keepByTextType) {
                return false;
            }

            if (filterOptions.keepAlignLeft || filterOptions.keepAlignCenter || filterOptions.keepAlignRight) {
                var alignmentType = getTextAlignmentType(item);
                if (alignmentType === "left" && !filterOptions.keepAlignLeft) {
                    return false;
                }
                if (alignmentType === "center" && !filterOptions.keepAlignCenter) {
                    return false;
                }
                if (alignmentType === "right" && !filterOptions.keepAlignRight) {
                    return false;
                }
                if (alignmentType !== "left" && alignmentType !== "center" && alignmentType !== "right") {
                    return false;
                }
            }

            if (hasSelectedFonts(filterOptions.selectedFonts)) {
                return filterOptions.selectedFonts[getTextFontName(item)] === true;
            }

            return true;
        }

        if (item.typename === "PathItem") {
            if (item.closed === true) {
                var paintType = getClosedPathPaintType(item);
                if (paintType === "fillOnly") {
                    return filterOptions.keepClosedPathFillOnly;
                }
                if (paintType === "strokeOnly") {
                    return filterOptions.keepClosedPathStrokeOnly;
                }
                if (paintType === "fillAndStroke") {
                    return filterOptions.keepClosedPathFillAndStroke;
                }
                return false;
            }
            var orientation = getOpenPathLineOrientation(item);
            if (orientation === "horizontal") {
                return filterOptions.keepOpenPath || filterOptions.keepHorizontalLine;
            }
            if (orientation === "vertical") {
                return filterOptions.keepOpenPath || filterOptions.keepVerticalLine;
            }
            return filterOptions.keepOpenPath;
        }

        return false;
    }

    /* フォント指定があるかどうかを判定 / Check whether any font is selected */
    function hasSelectedFonts(selectedFonts) {
        for (var fontName in selectedFonts) {
            if (selectedFonts.hasOwnProperty(fontName)) {
                return true;
            }
        }
        return false;
    }

    /* 選択状態を条件に合わせて反映 / Apply filter options to selection */
    function applyFilterToSelection(items, filterOptions) {
        app.selection = null;
        applyFilterToItemsRecursive(items, filterOptions);
    }

    /* 子要素も含めて選択状態を条件に合わせて反映 / Apply filter options recursively including children */
    function applyFilterToItemsRecursive(items, filterOptions) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var item = items[itemIndex];

            try {
                item.selected = shouldKeepItem(item, filterOptions);
            } catch (e1) { }

            if (item.typename === "GroupItem" && filterOptions.includeGroupItems) {
                try {
                    applyFilterToItemsRecursive(item.pageItems, filterOptions);
                } catch (e2) { }
            }
        }
    }

    /* 行揃え/フォントパネルの有効状態を更新 / Update alignment & font panel enabled state */
    function updateAlignmentPanelEnabled(dialogUi) {
        var hasTextTarget = dialogUi.cbText.value || dialogUi.cbAreaText.value || dialogUi.cbPathText.value;
        dialogUi.alignmentPanel.enabled = hasTextTarget;
        if (dialogUi.fontPanel) dialogUi.fontPanel.enabled = hasTextTarget;
    }

    /* 元の表示状態を退避 / Capture original visibility/opacity for items (recursive) */
    function captureVisualState(items, snapshot) {
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var item = items[itemIndex];
            var entry = { item: item };
            try { entry.hidden = item.hidden; } catch (e1) { entry.hidden = false; }
            try { entry.opacity = item.opacity; } catch (e2) { entry.opacity = 100; }
            snapshot.push(entry);
            if (item.typename === "GroupItem") {
                try { captureVisualState(item.pageItems, snapshot); } catch (e3) { }
            }
        }
    }

    /* 退避した表示状態を復元 / Restore captured visibility/opacity */
    function restoreVisualState(snapshot) {
        for (var snapshotIndex = 0; snapshotIndex < snapshot.length; snapshotIndex++) {
            var visualState = snapshot[snapshotIndex];
            try { visualState.item.hidden = visualState.hidden; } catch (e1) { }
            try { visualState.item.opacity = visualState.opacity; } catch (e2) { }
        }
    }

    /* 非選択オブジェクトに挙動を適用 / Apply behavior to non-selected items
       mode: "none" | "hide" | "opacity"
       opacityValue: number (0-100)
    */
    function applyNonSelectedBehavior(items, mode, opacityValue) {
        if (mode === "none") return;
        for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var item = items[itemIndex];
            var isSelected = false;
            try { isSelected = item.selected === true; } catch (e0) { }
            if (!isSelected) {
                try {
                    if (mode === "hide") {
                        item.hidden = true;
                    } else if (mode === "opacity") {
                        item.opacity = opacityValue;
                    }
                } catch (e) { }
            }
            if (item.typename === "GroupItem") {
                try { applyNonSelectedBehavior(item.pageItems, mode, opacityValue); } catch (e2) { }
            }
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

    function showFilterDialog(originalSelection) {
        var dialog = new Window("dialog", getLabel("dialogTitle"));
        setupWindow(dialog);

        /* フォント一覧はグループ内も含めて作り、件数だけスコープに合わせて書き換える
           Build the font list including group contents; only the counts follow the scope */
        var counts = countSelectionItems(originalSelection, true);

        /* 元の表示状態スナップショット / Snapshot original visual state */
        var visualSnapshot = [];
        captureVisualState(originalSelection, visualSnapshot);

        /* 上部：選択オブジェクト以外の挙動 / Top: behavior for non-selected items */
        var nonSelectedPanel = dialog.add("panel", undefined, getLabel("nonSelectedPanel"));
        setupPanel(nonSelectedPanel, 6);
        nonSelectedPanel.alignChildren = ["left", "center"];

        var nonSelectedRadioGroup = nonSelectedPanel.add("group");
        nonSelectedRadioGroup.orientation = "row";
        nonSelectedRadioGroup.alignChildren = ["left", "center"];
        nonSelectedRadioGroup.spacing = 12;

        var rbNonSelNone = nonSelectedRadioGroup.add("radiobutton", undefined, getLabel("nonSelectedNone"));
        rbNonSelNone.helpTip = getLabel("tipNonSelectedNone");
        var rbNonSelHide = nonSelectedRadioGroup.add("radiobutton", undefined, getLabel("nonSelectedHide"));
        rbNonSelHide.helpTip = getLabel("tipNonSelectedHide");
        var rbNonSelOpacity = nonSelectedRadioGroup.add("radiobutton", undefined, labelWithValue("nonSelectedOpacity", DEFAULT_NON_SELECTED_OPACITY + "%"));
        rbNonSelOpacity.helpTip = getLabel("tipNonSelectedOpacity");

        var nonSelectedOpacitySliderGroup = nonSelectedPanel.add("group");
        nonSelectedOpacitySliderGroup.orientation = "row";
        nonSelectedOpacitySliderGroup.alignChildren = ["center", "center"];
        nonSelectedOpacitySliderGroup.alignment = ["fill", "center"];

        var nonSelectedOpacitySlider = nonSelectedOpacitySliderGroup.add("slider", undefined, DEFAULT_NON_SELECTED_OPACITY, 0, 100);
        nonSelectedOpacitySlider.helpTip = getLabel("tipNonSelectedSlider");
        nonSelectedOpacitySlider.preferredSize.width = 220;
        rbNonSelNone.value = true;

        function getNonSelectedMode() {
            if (rbNonSelHide.value) return "hide";
            if (rbNonSelOpacity.value) return "opacity";
            return "none";
        }

        function getNonSelectedOpacityValue() {
            return Math.max(0, Math.min(100, Math.round(nonSelectedOpacitySlider.value)));
        }

        function updateNonSelectedOpacityDisplay() {
            rbNonSelOpacity.text = labelWithValue("nonSelectedOpacity", getNonSelectedOpacityValue() + "%");
        }

        function updateNonSelectedOpacitySliderEnabled() {
            nonSelectedOpacitySlider.enabled = rbNonSelOpacity.value === true;
        }

        /* 簡易モード：横並びチェックボックス / Quick mode: horizontal checkboxes */
        var simplePanel = dialog.add("panel", undefined, getLabel("simplePanel"));
        setupPanel(simplePanel);
        simplePanel.orientation = "row";
        simplePanel.alignChildren = ["left", "center"];
        var cbSimpleText = simplePanel.add("checkbox", undefined, getLabel("simpleText"));
        cbSimpleText.helpTip = getLabel("tipSimpleText");
        var cbSimpleStrokeOnly = simplePanel.add("checkbox", undefined, getLabel("simpleStrokeOnly"));
        cbSimpleStrokeOnly.helpTip = getLabel("tipSimpleStrokeOnly");
        var cbSimpleFillOnlyPath = simplePanel.add("checkbox", undefined, getLabel("simpleFillOnlyPath"));
        cbSimpleFillOnlyPath.helpTip = getLabel("tipSimpleFillOnlyPath");

        /* 対象スコープ / Selection scope */
        var scopePanel = dialog.add("panel", undefined, getLabel("scopePanel"));
        setupPanel(scopePanel);
        scopePanel.orientation = "row";
        scopePanel.alignChildren = ["left", "center"];
        var rbScopeSelectedOnly = scopePanel.add("radiobutton", undefined, getLabel("scopeSelectedOnly"));
        rbScopeSelectedOnly.helpTip = getLabel("tipScope");
        var rbScopeIncludeGroupItems = scopePanel.add("radiobutton", undefined, getLabel("scopeIncludeGroupItems"));
        rbScopeIncludeGroupItems.helpTip = getLabel("tipScope");
        rbScopeIncludeGroupItems.value = true;

        var columnGroup = dialog.add("group");
        columnGroup.orientation = "row";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.alignment = "fill";
        columnGroup.spacing = COLUMN_SPACING;

        var leftColumn = columnGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";
        leftColumn.alignment = ["fill", "top"];

        var rightColumn = columnGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";
        rightColumn.alignment = ["fill", "top"];

        var textPanel = leftColumn.add("panel", undefined, getLabel("textPanel"));
        setupPanel(textPanel, 6);
        var cbText = textPanel.add("checkbox", undefined, labelWithCount("textObject", counts.text));
        var cbAreaText = textPanel.add("checkbox", undefined, labelWithCount("areaText", counts.areaText));
        var cbPathText = textPanel.add("checkbox", undefined, labelWithCount("pathText", counts.pathText));

        var alignmentPanel = leftColumn.add("panel", undefined, getLabel("alignmentPanel"));
        setupPanel(alignmentPanel, 6);
        var cbAlignLeft = alignmentPanel.add("checkbox", undefined, labelWithCount("alignLeft", counts.alignLeft));
        var cbAlignCenter = alignmentPanel.add("checkbox", undefined, labelWithCount("alignCenter", counts.alignCenter));
        var cbAlignRight = alignmentPanel.add("checkbox", undefined, labelWithCount("alignRight", counts.alignRight));

        var pathPanel = rightColumn.add("panel", undefined, getLabel("pathPanel"));
        setupPanel(pathPanel, 6);
        var cbOpenPath = pathPanel.add("checkbox", undefined, labelWithCount("openPath", counts.openPath));
        var cbHorizontalLine = pathPanel.add("checkbox", undefined, labelWithCount("horizontalLine", counts.horizontalLine));
        var cbVerticalLine = pathPanel.add("checkbox", undefined, labelWithCount("verticalLine", counts.verticalLine));

        var closedPathPanel = rightColumn.add("panel", undefined, getLabel("closedPathPanel"));
        setupPanel(closedPathPanel, 6);
        var cbClosedPathFillOnly = closedPathPanel.add("checkbox", undefined, labelWithCount("closedPathFillOnly", counts.closedPathFillOnly));
        var cbClosedPathStrokeOnly = closedPathPanel.add("checkbox", undefined, labelWithCount("closedPathStrokeOnly", counts.closedPathStrokeOnly));
        var cbClosedPathFillAndStroke = closedPathPanel.add("checkbox", undefined, labelWithCount("closedPathFillAndStroke", counts.closedPathFillAndStroke));

        /* フォントパネル：カラム貫通（全幅） / Font panel spanning across columns */
        var fontPanel = dialog.add("panel", undefined, getLabel("fontPanel"));
        setupPanel(fontPanel, 6);
        var fontCheckboxes = [];
        var fontNames = getSortedFontNames(counts.fonts);
        for (var fontIndex = 0; fontIndex < fontNames.length; fontIndex++) {
            var fontName = fontNames[fontIndex];
            var fontCheckboxLabel = appendValueText(fontName, counts.fonts[fontName]);
            var fontCheckbox = fontPanel.add("checkbox", undefined, fontCheckboxLabel);
            fontCheckboxes.push({
                fontName: fontName,
                checkbox: fontCheckbox
            });
        }

        var dialogUi = {
            cbText: cbText,
            cbAreaText: cbAreaText,
            cbPathText: cbPathText,
            alignmentPanel: alignmentPanel,
            cbAlignLeft: cbAlignLeft,
            cbAlignCenter: cbAlignCenter,
            cbAlignRight: cbAlignRight,
            fontPanel: fontPanel,
            fontCheckboxes: fontCheckboxes,
            cbOpenPath: cbOpenPath,
            cbHorizontalLine: cbHorizontalLine,
            cbVerticalLine: cbVerticalLine,
            cbClosedPathFillOnly: cbClosedPathFillOnly,
            cbClosedPathStrokeOnly: cbClosedPathStrokeOnly,
            cbClosedPathFillAndStroke: cbClosedPathFillAndStroke,
            rbScopeIncludeGroupItems: rbScopeIncludeGroupItems
        };

        var filterCheckboxes = [
            cbText,
            cbAreaText,
            cbPathText,
            cbAlignLeft,
            cbAlignCenter,
            cbAlignRight,
            cbOpenPath,
            cbHorizontalLine,
            cbVerticalLine,
            cbClosedPathFillOnly,
            cbClosedPathStrokeOnly,
            cbClosedPathFillAndStroke
        ];

        for (var fontCheckboxIndex = 0; fontCheckboxIndex < fontCheckboxes.length; fontCheckboxIndex++) {
            filterCheckboxes.push(fontCheckboxes[fontCheckboxIndex].checkbox);
        }

        /* デフォルトでは全てチェックを外す / Uncheck all options by default */
        for (var checkboxIndex = 0; checkboxIndex < filterCheckboxes.length; checkboxIndex++) {
            filterCheckboxes[checkboxIndex].value = false;
        }

        /* テキスト系をまとめてデフォルトON / Enable all text-related options by default */
        cbSimpleText.value = true;
        cbText.value = true;
        cbAreaText.value = true;
        cbPathText.value = true;

        /* スコープに合わせて件数表示を更新 / Update the count labels for the current scope */
        function updateCountLabels() {
            var scopeCounts = countSelectionItems(originalSelection, rbScopeIncludeGroupItems.value);
            cbText.text = labelWithCount("textObject", scopeCounts.text);
            cbAreaText.text = labelWithCount("areaText", scopeCounts.areaText);
            cbPathText.text = labelWithCount("pathText", scopeCounts.pathText);
            cbAlignLeft.text = labelWithCount("alignLeft", scopeCounts.alignLeft);
            cbAlignCenter.text = labelWithCount("alignCenter", scopeCounts.alignCenter);
            cbAlignRight.text = labelWithCount("alignRight", scopeCounts.alignRight);
            cbOpenPath.text = labelWithCount("openPath", scopeCounts.openPath);
            cbHorizontalLine.text = labelWithCount("horizontalLine", scopeCounts.horizontalLine);
            cbVerticalLine.text = labelWithCount("verticalLine", scopeCounts.verticalLine);
            cbClosedPathFillOnly.text = labelWithCount("closedPathFillOnly", scopeCounts.closedPathFillOnly);
            cbClosedPathStrokeOnly.text = labelWithCount("closedPathStrokeOnly", scopeCounts.closedPathStrokeOnly);
            cbClosedPathFillAndStroke.text = labelWithCount("closedPathFillAndStroke", scopeCounts.closedPathFillAndStroke);
            for (var fontLabelIndex = 0; fontLabelIndex < fontCheckboxes.length; fontLabelIndex++) {
                var fontEntry = fontCheckboxes[fontLabelIndex];
                fontEntry.checkbox.text = appendValueText(fontEntry.fontName, scopeCounts.fonts[fontEntry.fontName] || 0);
            }
        }

        function updateSimpleCheckboxesFromDetailedSelection() {
            cbSimpleText.value = cbText.value === true && cbAreaText.value === true && cbPathText.value === true;
            cbSimpleStrokeOnly.value = cbOpenPath.value === true && cbClosedPathStrokeOnly.value === true;
            cbSimpleFillOnlyPath.value = cbClosedPathFillOnly.value === true;
        }

        function updateCanvasSelection() {
            updateAlignmentPanelEnabled(dialogUi);
            updateSimpleCheckboxesFromDetailedSelection();
            restoreVisualState(visualSnapshot);
            applyFilterToSelection(originalSelection, readFilterOptions(dialogUi));
            applyNonSelectedBehavior(originalSelection, getNonSelectedMode(), getNonSelectedOpacityValue());
            app.redraw();
        }

        rbNonSelNone.onClick = function () {
            updateNonSelectedOpacitySliderEnabled();
            updateCanvasSelection();
        };
        rbNonSelHide.onClick = function () {
            updateNonSelectedOpacitySliderEnabled();
            updateCanvasSelection();
        };
        rbNonSelOpacity.onClick = function () {
            updateNonSelectedOpacitySliderEnabled();
            updateCanvasSelection();
        };
        nonSelectedOpacitySlider.onChanging = function () {
            rbNonSelOpacity.value = true;
            updateNonSelectedOpacityDisplay();
            updateNonSelectedOpacitySliderEnabled();
            updateCanvasSelection();
        };
        nonSelectedOpacitySlider.onChange = function () {
            rbNonSelOpacity.value = true;
            updateNonSelectedOpacityDisplay();
            updateNonSelectedOpacitySliderEnabled();
            updateCanvasSelection();
        };

        rbScopeSelectedOnly.onClick = function () {
            updateCountLabels();
            updateCanvasSelection();
        };
        rbScopeIncludeGroupItems.onClick = function () {
            updateCountLabels();
            updateCanvasSelection();
        };

        cbSimpleText.onClick = function () {
            cbText.value = cbSimpleText.value;
            cbAreaText.value = cbSimpleText.value;
            cbPathText.value = cbSimpleText.value;
            updateCanvasSelection();
        };
        cbSimpleStrokeOnly.onClick = function () {
            cbOpenPath.value = cbSimpleStrokeOnly.value;
            cbClosedPathStrokeOnly.value = cbSimpleStrokeOnly.value;
            updateCanvasSelection();
        };
        cbSimpleFillOnlyPath.onClick = function () {
            cbClosedPathFillOnly.value = cbSimpleFillOnlyPath.value;
            updateCanvasSelection();
        };

        bindCheckboxPreviewUpdates(filterCheckboxes, updateCanvasSelection);

        /* ボタン行（左：アウトライン切り替え、右：キャンセル・OK） / Button row (left: outline toggle, right: Cancel and OK) */
        var buttonRow = addButtonRow(dialog);

        var isOutlineMode = false;
        var btnPreview = buttonRow.leftGroup.add("button", undefined, getLabel("btnOutlineOn"));
        btnPreview.helpTip = getLabel("tipPreviewButton");
        btnPreview.onClick = function () {
            try {
                app.executeMenuCommand("preview");
                isOutlineMode = !isOutlineMode;
                btnPreview.text = isOutlineMode ? getLabel("btnOutlineOff") : getLabel("btnOutlineOn");
            } catch (e) { }
        };

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("ok"), { name: "ok" });

        updateNonSelectedOpacitySliderEnabled();
        updateCanvasSelection();

        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(dialog);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();

        if (isOutlineMode) {
            try {
                app.executeMenuCommand("preview");
            } catch (e) { }
        }

        restoreVisualState(visualSnapshot);

        if (dialogResult !== 1) {
            restoreSelection(originalSelection);
            return null;
        }

        return readFilterOptions(dialogUi);
    }

    (function () {
        if (app.documents.length === 0) {
            showError("noDocumentError");
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = copySelection(doc.selection);

        if (originalSelection.length === 0) {
            showError("noSelectionError");
            return;
        }

        showFilterDialog(originalSelection);
    })();

})();
