#target illustrator

/*

### 概要

ダイアログ・パレットの左下にバージョン（SCRIPT_VERSION）を小さく薄く表示する再利用テンプレートです。
タイトルにはバージョンを入れず、この部品で下端に表示します。

### Overview

A reusable template that shows the version (SCRIPT_VERSION) small and dimmed at the bottom left of a dialog or palette.
The version is no longer put in the title; this part shows it at the bottom instead.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "VersionFooter";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-05";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内（ボタン行の部品のあと）に貼る。
    //    識別子は VERSION_FOOTER_* / addVersionFooter / measureVersionFooterDots / drawVersionFooterDots
    //    色の明暗に UITheme 部品の isDarkUI() を使うので、同じ階層に UITheme の部品が無ければ一緒に貼る（このファイルの冒頭の部品）
    // 2. ウィンドウのタイトルにはバージョンを入れない
    //      var mainDialog = new Window("dialog", getLabel("dialog.title"));
    // 3. 部品をすべて足したあと、show() の直前で addVersionFooter() を呼ぶ（ウィンドウは縦並び。何度呼んでも1回だけ足す）。
    //    prepareDialogWindow() を使うときはその前で呼ぶ
    //      addVersionFooter(mainDialog);
    //      mainDialog.show();
    // 4. 直前の行（ボタン行など）の下余白は 0 になり、ウィンドウの下余白は VERSION_FOOTER_BOTTOM_MARGIN になる。
    //    バージョンの上の間隔は、ウィンドウの間隔に VERSION_FOOTER_TOP_MARGIN を足したもの
    // 5. 確認・進捗などの小さいウィンドウには付けない（メインのダイアログ・パレットだけ）
    // 6. InDesign の ScriptUI は文字サイズの指定を無視する（statictext も onDraw の drawString も）ので、
    //    文字は使わず、ドットフォント Chava の形を写した「v」「.」「0〜9」を onDraw でドットとして描く（高さ6ドット、1ドット1px）。
    //    大きさは VERSION_FOOTER_PIXEL_SIZE、色は VERSION_FOOTER_COLOR で変える。字形に無い文字（「-」など）は飛ばす

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

    // =========================================
    // デモ / Demo
    // =========================================
    /**
     * デモのダイアログを表示する
     * @returns {void}
     */
    function showDemoDialog() {
        var demoDialog = new Window("dialog", "Version Footer");
        demoDialog.orientation = "column";
        demoDialog.alignChildren = ["fill", "top"];
        demoDialog.margins = 16;
        demoDialog.spacing = 12;

        var contentPanel = demoDialog.add("panel", undefined, "Content");
        contentPanel.preferredSize = [320, 80];

        var btnRowGroup = demoDialog.add("group");
        btnRowGroup.alignment = ["right", "bottom"];
        btnRowGroup.margins = [0, 5, 0, 14];
        btnRowGroup.add("button", undefined, "Cancel", { name: "cancel" });
        btnRowGroup.add("button", undefined, "OK", { name: "ok" });

        addVersionFooter(demoDialog);
        demoDialog.show();
    }

    showDemoDialog();

})();
