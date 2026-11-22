#target illustrator
#targetengine "SmartClipAndGroupEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SmartClipAndGroup
Meta-Default-language=ja
Name=オブジェクトのまとまりごとにグループ化したり、マスクする
Name[en]=Group or mask objects in logical sets
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.1
Release-Date=2024-06-05
Tags=group,グループ
X-com.swwwitch-Script-Name=SmartClipAndGroup
X-com.swwwitch-Updated=2026-11-20
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartClipAndGroup.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartClipAndGroup.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/nb23985473f80
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SmartClipAndGroup.jsx
Description-BEGIN
[ja]
選択オブジェクトを重なりや距離、縦の列・横の行、塗りと線（アピアランス）でまとまりに分け、まとまりごとにグループ化するか、クリッピングマスクを作成します（配置画像を1つずつクリップすることもできます）。
［OK］の前に、できるグループの数と範囲を件数表示とプレビューの枠で確認できます。
[en]
Splits the selection into clusters by overlap, distance, column, row or matching fill and stroke (appearance), and groups or clips each cluster (placed images can also be clipped one by one).
Before you click OK, a count and preview frames show the groups to be created.
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
    // ユーザー設定 / User Settings
    // =========================================

    /* しきい値スライダーの初期値と範囲（pt） / Default and range of the threshold slider (pt) */
    var DEFAULT_THRESHOLD = 10;
    var THRESHOLD_MIN     = 0;
    var THRESHOLD_MAX     = 100;

    /* 「間隔が空いたら分ける」の初期値 / Default of "Split at gaps" */
    var DEFAULT_SPLIT_AT_GAPS = false;

    /* 「プレビューを表示」の初期値 / Default of "Show preview" */
    var DEFAULT_SHOW_PREVIEW = true;

    /* 「重なり」や、列・行で揃っているとみなす距離（pt） / Gap treated as touching or aligned (pt) */
    var TOUCH_TOLERANCE = 0.01;

    /* スライダーのドラッグ中も件数とプレビューを更新する選択数の上限（超えると指を離したときだけ更新する）
       Max selection size for live updates while dragging; above it, the count and preview update on release only */
    var LIVE_COUNT_MAX_ITEMS = 300;

    /* プレビューの枠（赤・塗りなし・線 10 pt・不透明度 50%） / Preview frames: red, no fill, 10 pt stroke, 50% opacity */
    var PREVIEW_RGB          = [255, 0, 0];     /* RGB ドキュメントの線の色 / Stroke color in RGB documents */
    var PREVIEW_CMYK         = [0, 100, 100, 0]; /* CMYK ドキュメントの線の色 / Stroke color in CMYK documents */
    var PREVIEW_STROKE_WIDTH = 10;              /* 線幅（pt） / Stroke width (pt) */
    var PREVIEW_OPACITY      = 50;              /* 不透明度（%） / Opacity (%) */
    var PREVIEW_LAYER_NAME   = "SmartClipAndGroup Preview"; /* プレビュー用の一時レイヤー名 / Temporary preview layer name */

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

    /* ダイアログ固有の寸法 / Dialog-specific sizes */
    var RADIO_COLUMN_SPACING = 20;                /* ラジオボタンの列の間隔 / Gap between radio button columns */
    var SLIDER_WIDTH         = 150;               /* しきい値スライダーの幅 / Threshold slider width */
    var VALUE_TEXT_CHARS     = 5;                 /* しきい値表示の文字数 / Width of the threshold readout in characters */

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
        var versionText = getVersionInfoValue("version");
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
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA の項目を getScriptMetaValue() で直接読み（スクリプト情報の部品）、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPTMETA fields directly with getScriptMetaValue(),
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

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
     * 「このスクリプトについて」とフッターに出す値を返す。SCRIPTMETA の項目を getScriptMetaValue() で直接読み、名前は SCRIPT_NAME。
     * getScriptMetaValue() が無いスクリプトは従来の SCRIPT_* を typeof で読む（未定義でも例外にならない）
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値（バージョンは「v1.2.0」の形）。無ければ空文字
     */
    function getVersionInfoValue(variableKey) {
        var infoValue = "";
        if (variableKey === "name") {
            infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        } else if (typeof getScriptMetaValue === "function") {
            var isJapanese = (VERSION_INFO_LANG === "ja");
            if (variableKey === "version") infoValue = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : "";
            else if (variableKey === "feature") infoValue = getScriptMetaValue("Name");
            else if (variableKey === "readme") infoValue = getScriptMetaValue(isJapanese ? "X-com.swwwitch-README-JA" : "X-com.swwwitch-README-EN");
            else if (variableKey === "source") infoValue = getScriptMetaValue("X-com.swwwitch-Source-URL");
            else if (variableKey === "article") infoValue = getScriptMetaValue("X-com.swwwitch-Article-URL");
        } else {
            if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
            else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
            else if (variableKey === "readme") {
                if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
                else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
            }
            else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
            else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        }
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
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
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

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        dialog: {
            title: { ja: "クリップとグループ化", en: "Clip and Group" }
        },
        panel: {
            clip: { ja: "クリッピングマスク", en: "Clipping Mask" },
            grouping: { ja: "グループ化", en: "Grouping" }
        },
        radio: {
            clipWithFront: { ja: "最前面でクリップ", en: "Clip with Frontmost" },
            clipWithBack: { ja: "最背面でクリップ", en: "Clip with Backmost" },
            clipPlacedOnly: { ja: "配置画像のみをクリップ", en: "Clip Placed Images Only" },
            groupOverlap: { ja: "重なり", en: "Overlap" },
            groupProximity: { ja: "近接度", en: "Proximity" },
            groupVertical: { ja: "上下方向", en: "Vertical" },
            groupHorizontal: { ja: "左右方向", en: "Horizontal" },
            groupAppearance: { ja: "アピアランス", en: "Appearance" }
        },
        fieldLabel: {
            threshold: { ja: "しきい値", en: "Threshold" },
            selectedCount: { ja: "選択しているオブジェクト", en: "Selected objects" },
            groupCount: { ja: "グループ化", en: "Groups" },
            clipCount: { ja: "クリップ", en: "Clips" }
        },
        checkbox: {
            splitAtGaps: { ja: "間隔が空いたら分ける", en: "Split at gaps" },
            perArtboard: { ja: "アートボードごと", en: "Within each artboard" },
            showPreview: { ja: "プレビューを表示", en: "Show preview" }
        },
        tooltip: {
            clipWithFront: {
                ja: "重なっているまとまりごとに、最前面のパスを型にしてクリップします。",
                en: "Clips each cluster of overlapping objects, using its frontmost path as the mask."
            },
            clipWithBack: {
                ja: "重なっているまとまりごとに、最背面のパスを型にしてクリップします。",
                en: "Clips each cluster of overlapping objects, using its backmost path as the mask."
            },
            clipPlacedOnly: {
                ja: "画像を1つずつ、同じ大きさの長方形でクリップします。クリップグループを選ぶと、元のマスクだけを削除して中の画像をクリップし直します。",
                en: "Clips each image with a rectangle of the same size. For a clipping group, only the old mask is removed and the images inside are clipped again."
            },
            groupOverlap: {
                ja: "外枠が重なっているか接しているオブジェクトをグループにします。",
                en: "Groups objects whose bounding boxes overlap or touch."
            },
            groupProximity: {
                ja: "しきい値以内の距離にあるオブジェクトを、向きを問わずグループにします。",
                en: "Groups objects that sit within the threshold distance, in any direction."
            },
            groupVertical: {
                ja: "上下に並んでいるオブジェクトを、縦の列ごとにグループにします。左右のすき間がしきい値以内なら同じ列とみなします。",
                en: "Groups objects that line up vertically, column by column. Objects whose horizontal gap is within the threshold count as the same column."
            },
            groupHorizontal: {
                ja: "左右に並んでいるオブジェクトを、横の行ごとにグループにします。上下のすき間がしきい値以内なら同じ行とみなします。",
                en: "Groups objects that line up horizontally, row by row. Objects whose vertical gap is within the threshold count as the same row."
            },
            groupAppearance: {
                ja: "塗りと線（色・線幅）が同じオブジェクトを、位置を問わずグループにします。グループや画像など塗りと線を持たないものはグループにしません。",
                en: "Groups objects with the same fill and stroke (colors and stroke weight), wherever they are. Groups, images and other items without a fill and stroke are left as they are."
            },
            threshold: {
                ja: "同じグループとみなす距離です。大きくするとまとまりが粗くなります。",
                en: "How close objects must be to land in the same group. Larger values group more loosely."
            },
            splitAtGaps: {
                ja: "上下方向・左右方向で使います。ONにすると、列（行）の途中ですき間がしきい値より空いたところでグループを分けます。つなぐのは揃って並んでいるもの（列なら左右が重なるもの）だけです。OFFのときは距離を問わず、列（行）全体をまとめます。",
                en: "Used with Vertical and Horizontal. When on, a column (row) is split wherever the gap exceeds the threshold, and only aligned objects (horizontally overlapping, for columns) are linked. When off, whole columns (rows) are grouped regardless of distance."
            },
            perArtboard: {
                ja: "ONにすると、別のアートボードにあるオブジェクトは別々のグループにします。オブジェクトの中心がどのアートボードに含まれるかで判定します。",
                en: "When on, objects on different artboards go into separate groups. The artboard is determined by each object's center point."
            },
            showPreview: {
                ja: "ONにすると、作成されるグループの範囲を赤い枠で表示します。枠はダイアログを閉じると消えます。",
                en: "When on, red frames show where the groups will be created. The frames disappear when the dialog closes."
            },
            resultCount: {
                ja: "［OK］で作成されるグループの数です。クリッピングマスクのモードでは、作成されるクリップグループの数です。",
                en: "The number of groups OK will create. In the clipping mask modes, the number of clipping groups."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    // =========================================
    // モード / Modes
    // =========================================

    /* パネルごとのモードの並び（キーは LABELS.radio / LABELS.tooltip と共通）
       Mode order per panel; keys are shared with LABELS.radio and LABELS.tooltip */
    var CLIP_MODES  = ["clipWithFront", "clipWithBack", "clipPlacedOnly"];
    var GROUP_MODES = ["groupOverlap", "groupProximity", "groupVertical", "groupHorizontal", "groupAppearance"];
    var ALL_MODES   = CLIP_MODES.concat(GROUP_MODES);

    /* グループ化のモードごとに使う設定（クリップ系はどれも使わない）
       Settings each grouping mode uses; the clipping modes use none */
    var MODE_OPTIONS = {
        groupOverlap: { perArtboard: true },
        groupProximity: { threshold: true, perArtboard: true },
        groupVertical: { threshold: true, splitAtGaps: true, perArtboard: true },
        groupHorizontal: { threshold: true, splitAtGaps: true, perArtboard: true },
        groupAppearance: { perArtboard: true }
    };

    /* まとまりをクリップするモードと、型にするパスの側 / Cluster-clipping modes and which side supplies the mask */
    var MASK_SIDES = { clipWithFront: "front", clipWithBack: "back" };

    /* すき間の判定を通さず、面積を持つ重なりだけでつなぐ（すき間は0以上なので常に不成立）
       Link by area overlap only; gaps are never negative, so the gap test always fails */
    var OVERLAP_ONLY = -1;

    // =========================================
    // 選択 / Selection
    // =========================================

    /**
     * 処理対象の選択オブジェクトを返す
     * @returns {PageItem[]|null} 選択オブジェクト。ドキュメントがない・未選択・文字の選択中は null
     */
    function getValidSelection() {
        if (!app.documents.length) return null;
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る
           A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || !selectedItems.length || selectedItems.typename === "TextRange") return null;
        return selectedItems;
    }

    /**
     * 配置画像または埋め込み画像かどうかを返す
     * @param {PageItem} targetItem - 判定するオブジェクト
     * @returns {boolean} PlacedItem か RasterItem なら true
     */
    function isImageItem(targetItem) {
        return targetItem.typename === "PlacedItem" || targetItem.typename === "RasterItem";
    }

    /**
     * すべてが画像かどうかを返す
     * @param {PageItem[]|null} targetItems - 判定するオブジェクト
     * @returns {boolean} 1つ以上あり、すべてが画像なら true
     */
    function isAllImages(targetItems) {
        if (!targetItems || !targetItems.length) return false;
        for (var i = 0; i < targetItems.length; i++) {
            if (!isImageItem(targetItems[i])) return false;
        }
        return true;
    }

    /**
     * 選択を解除して、指定のオブジェクトを選択する
     * @param {PageItem[]} itemsToSelect - 選択するオブジェクト
     * @returns {void}
     */
    function selectOnly(itemsToSelect) {
        app.activeDocument.selection = null;
        for (var i = 0; i < itemsToSelect.length; i++) {
            itemsToSelect[i].selected = true;
        }
    }

    // =========================================
    // まとまりの判定 / Cluster detection
    // =========================================

    /**
     * @typedef {Object} ClusterRule まとまりのつなぎ方
     * @property {number} [maxGap] - つなぐ最大のすき間（pt）。OVERLAP_ONLY なら重なりだけでつなぐ
     * @property {boolean} [byAppearance] - 位置ではなく、塗りと線が同じものをまとめる
     * @property {string} [direction] - "vertical" なら縦の列、"horizontal" なら横の行で判定
     * @property {boolean} [splitAtGaps] - 列・行の途中で、すき間がしきい値より空いたところで分ける（揃っているものだけをつなぐ）
     * @property {boolean} [perArtboard] - 別のアートボードにあるものはつながない
     */

    /**
     * 2つの矩形が面積を持って重なっているかを返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @returns {boolean} 重なっていれば true
     */
    function boundsOverlap(boundsA, boundsB) {
        var overlapWidth = Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]);
        var overlapHeight = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
        return overlapWidth > 0 && overlapHeight > 0;
    }

    /**
     * 2つの矩形の左右のすき間を返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @returns {number} すき間（pt）。左右の範囲が重なっていれば 0
     */
    function getHorizontalGap(boundsA, boundsB) {
        return Math.max(0, boundsB[0] - boundsA[2], boundsA[0] - boundsB[2]);
    }

    /**
     * 2つの矩形の上下のすき間を返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @returns {number} すき間（pt）。上下の範囲が重なっていれば 0
     */
    function getVerticalGap(boundsA, boundsB) {
        return Math.max(0, boundsA[3] - boundsB[1], boundsB[3] - boundsA[1]);
    }

    /**
     * 列・行の判定で2つの矩形をつなぐかどうかを返す
     * @param {number} crossGap - 並ぶ向きと直交する向きのすき間（列なら左右）
     * @param {number} alongGap - 並ぶ向きのすき間（列なら上下）
     * @param {ClusterRule} clusterRule - つなぎ方
     * @returns {boolean} つなぐなら true
     */
    function areInSameLine(crossGap, alongGap, clusterRule) {
        /* 間隔が空いたら分ける：揃っていて、並ぶ向きのすき間がしきい値以内 / Split at gaps: aligned, and the gap along the line is within the threshold */
        if (clusterRule.splitAtGaps) return crossGap <= TOUCH_TOLERANCE && alongGap <= clusterRule.maxGap;
        /* 距離は問わない：直交する向きのずれがしきい値以内 / Any distance along the line; the cross offset is within the threshold */
        return crossGap <= clusterRule.maxGap;
    }

    /**
     * 2つの矩形を同じまとまりとしてつなぐかどうかを返す
     * @param {number[]} boundsA - geometricBounds [左, 上, 右, 下]
     * @param {number[]} boundsB - geometricBounds [左, 上, 右, 下]
     * @param {ClusterRule} clusterRule - つなぎ方
     * @returns {boolean} つなぐなら true
     */
    function areNeighbors(boundsA, boundsB, clusterRule) {
        var horizontalGap = getHorizontalGap(boundsA, boundsB);
        var verticalGap = getVerticalGap(boundsA, boundsB);
        if (clusterRule.direction === "vertical") return areInSameLine(horizontalGap, verticalGap, clusterRule);
        if (clusterRule.direction === "horizontal") return areInSameLine(verticalGap, horizontalGap, clusterRule);
        return boundsOverlap(boundsA, boundsB) || Math.max(horizontalGap, verticalGap) <= clusterRule.maxGap;
    }

    /**
     * 起点からつながっている矩形の番号をすべて集める（visited を更新する）
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @param {number} startIndex - 起点の番号
     * @param {boolean[]} visited - 集め済みの印
     * @param {ClusterRule} clusterRule - つなぎ方
     * @param {number[]|null} artboardIndexes - 各オブジェクトのアートボード番号（アートボードごとに分けないときは null）
     * @returns {number[]} つながっている番号（昇順）
     */
    function collectConnectedIndexes(boundsList, startIndex, visited, clusterRule, artboardIndexes) {
        var connectedIndexes = [];
        var pendingIndexes = [startIndex];
        visited[startIndex] = true;
        /* 再帰を使わず積み残しを順に処理する（大きなまとまりでスタックがあふれないように）
           Walk a work list instead of recursing so large clusters cannot overflow the stack */
        while (pendingIndexes.length) {
            var currentIndex = pendingIndexes.pop();
            connectedIndexes.push(currentIndex);
            for (var j = 0; j < boundsList.length; j++) {
                if (visited[j]) continue;
                if (artboardIndexes && artboardIndexes[j] !== artboardIndexes[currentIndex]) continue;
                if (!areNeighbors(boundsList[currentIndex], boundsList[j], clusterRule)) continue;
                visited[j] = true;
                pendingIndexes.push(j);
            }
        }
        connectedIndexes.sort(function (a, b) { return a - b; });
        return connectedIndexes;
    }

    /**
     * 矩形の並びを、重なりや距離でつながるまとまりに分ける
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @param {ClusterRule} clusterRule - つなぎ方
     * @param {number[]|null} artboardIndexes - 各オブジェクトのアートボード番号（アートボードごとに分けないときは null）
     * @returns {Array<number[]>} まとまりごとの番号の配列（各まとまりは昇順）
     */
    function collectClusterIndexes(boundsList, clusterRule, artboardIndexes) {
        var visited = [];
        for (var i = 0; i < boundsList.length; i++) {
            visited.push(false);
        }
        var clusterIndexes = [];
        for (var j = 0; j < boundsList.length; j++) {
            if (!visited[j]) clusterIndexes.push(collectConnectedIndexes(boundsList, j, visited, clusterRule, artboardIndexes));
        }
        return clusterIndexes;
    }

    /**
     * モードと設定に応じた、まとまりのつなぎ方を返す
     * @param {string} mode - モードのキー（「配置画像のみをクリップ」以外）
     * @param {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} groupingOptions - ダイアログの設定
     * @returns {ClusterRule|null} つなぎ方。該当しないモードは null
     */
    function getClusterRule(mode, groupingOptions) {
        var usedOptions = MODE_OPTIONS[mode] || {};
        var perArtboard = usedOptions.perArtboard === true && groupingOptions.perArtboard === true;
        switch (mode) {
            case "clipWithFront":
            case "clipWithBack":
                return { maxGap: OVERLAP_ONLY };
            case "groupOverlap":
                return { maxGap: TOUCH_TOLERANCE, perArtboard: perArtboard };
            case "groupProximity":
                return { maxGap: groupingOptions.threshold, perArtboard: perArtboard };
            case "groupAppearance":
                return { byAppearance: true, perArtboard: perArtboard };
            case "groupVertical":
            case "groupHorizontal":
                return {
                    maxGap: groupingOptions.threshold,
                    direction: (mode === "groupVertical") ? "vertical" : "horizontal",
                    splitAtGaps: groupingOptions.splitAtGaps === true,
                    perArtboard: perArtboard
                };
        }
        return null;
    }

    // =========================================
    // アピアランス / Appearance
    // =========================================

    /**
     * 色を、同じ色なら同じになる文字列にする
     * @param {Color} targetColor - 塗りまたは線の色
     * @returns {string} 色のキー
     */
    function getColorKey(targetColor) {
        if (!targetColor) return "none";
        switch (targetColor.typename) {
            case "RGBColor":
                return ["rgb", roundForKey(targetColor.red), roundForKey(targetColor.green), roundForKey(targetColor.blue)].join(",");
            case "CMYKColor":
                return ["cmyk", roundForKey(targetColor.cyan), roundForKey(targetColor.magenta), roundForKey(targetColor.yellow), roundForKey(targetColor.black)].join(",");
            case "GrayColor":
                return "gray," + roundForKey(targetColor.gray);
            case "LabColor":
                return ["lab", roundForKey(targetColor.l), roundForKey(targetColor.a), roundForKey(targetColor.b)].join(",");
            case "SpotColor":
                return "spot," + targetColor.spot.name + "," + roundForKey(targetColor.tint);
            case "PatternColor":
                return "pattern," + targetColor.pattern.name;
            case "GradientColor":
                return "gradient," + targetColor.gradient.name;
            case "NoColor":
                return "none";
        }
        return String(targetColor.typename);
    }

    /**
     * 数値をキー用に丸める（計算の誤差で別の色・線幅と判定しないように）
     * @param {number} value - 色の値や線幅
     * @returns {string} 小数第3位までに丸めた文字列
     */
    function roundForKey(value) {
        return String(Math.round(value * 1000) / 1000);
    }

    /**
     * 塗りと線の組み合わせのキーを作る
     * @param {boolean} isFilled - 塗りがあるか
     * @param {Color} fillColor - 塗りの色
     * @param {boolean} isStroked - 線があるか
     * @param {Color} strokeColor - 線の色
     * @param {number} strokeWeight - 線幅（pt）
     * @returns {string} アピアランスのキー
     */
    function buildAppearanceKey(isFilled, fillColor, isStroked, strokeColor, strokeWeight) {
        var fillKey = isFilled ? getColorKey(fillColor) : "none";
        var strokeKey = isStroked ? getColorKey(strokeColor) : "none";
        if (strokeKey !== "none") strokeKey += "/" + roundForKey(strokeWeight);
        return "fill:" + fillKey + "|stroke:" + strokeKey;
    }

    /**
     * オブジェクトの塗りと線のキーを返す。複合パスは中の先頭のパス、テキストは先頭の文字の塗りと線で見る
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {string|null} アピアランスのキー（グループ・画像など塗りと線を持たないもの、読めないものは null）
     */
    function getAppearanceKey(targetItem) {
        try {
            var pathItem = targetItem;
            if (targetItem.typename === "CompoundPathItem") {
                if (!targetItem.pathItems.length) return null;
                pathItem = targetItem.pathItems[0];
            }
            if (pathItem.typename === "PathItem") {
                return buildAppearanceKey(pathItem.filled, pathItem.fillColor, pathItem.stroked, pathItem.strokeColor, pathItem.strokeWidth);
            }
            if (targetItem.typename === "TextFrame") {
                if (!targetItem.characters.length) return null;
                var charAttributes = targetItem.characters[0].characterAttributes;
                var textFill = charAttributes.fillColor;
                var textStroke = charAttributes.strokeColor;
                return buildAppearanceKey(textFill.typename !== "NoColor", textFill, textStroke.typename !== "NoColor", textStroke, charAttributes.strokeWeight);
            }
        } catch (e) {
            /* 読めないものはまとめない / Leave unreadable items ungrouped */
        }
        return null;
    }

    /**
     * 選択の各オブジェクトのアピアランスのキーを返す（初めて使うときに読み、選択の情報に控える）
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @returns {Array<string|null>} キー（選択と同じ並び）
     */
    function getAppearanceKeys(selectionInfo) {
        if (!selectionInfo.appearanceKeys) {
            var appearanceKeys = [];
            for (var i = 0; i < selectionInfo.items.length; i++) {
                appearanceKeys.push(getAppearanceKey(selectionInfo.items[i]));
            }
            selectionInfo.appearanceKeys = appearanceKeys;
        }
        return selectionInfo.appearanceKeys;
    }

    /**
     * キーが同じものを1つのまとまりにする（キーが null のものは1つずつ別のまとまり）
     * @param {Array<string|null>} itemKeys - 各オブジェクトのキー
     * @param {number[]|null} artboardIndexes - 各オブジェクトのアートボード番号（アートボードごとに分けないときは null）
     * @returns {Array<number[]>} まとまりごとの番号の配列（各まとまりは昇順）
     */
    function collectKeyClusters(itemKeys, artboardIndexes) {
        var clusterIndexes = [];
        var clusterByKey = {};
        for (var i = 0; i < itemKeys.length; i++) {
            if (itemKeys[i] === null) {
                clusterIndexes.push([i]);
                continue;
            }
            /* 接頭辞を付け、Object の既存のプロパティ名とぶつからないようにする / Prefix so keys never collide with built-in property names */
            var clusterKey = "k:" + (artboardIndexes ? artboardIndexes[i] + "#" : "") + itemKeys[i];
            if (!clusterByKey.hasOwnProperty(clusterKey)) {
                clusterByKey[clusterKey] = [];
                clusterIndexes.push(clusterByKey[clusterKey]);
            }
            clusterByKey[clusterKey].push(i);
        }
        return clusterIndexes;
    }

    // =========================================
    // アートボード / Artboards
    // =========================================

    /**
     * 矩形の中心が含まれるアートボードの番号を返す
     * @param {number[]} bounds - geometricBounds [左, 上, 右, 下]
     * @param {Array<number[]>} artboardRects - 各アートボードの artboardRect
     * @returns {number} アートボードの番号。どこにも含まれなければ -1
     */
    function findArtboardIndex(bounds, artboardRects) {
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        for (var i = 0; i < artboardRects.length; i++) {
            var artboardRect = artboardRects[i];
            if (centerX >= artboardRect[0] && centerX <= artboardRect[2] && centerY <= artboardRect[1] && centerY >= artboardRect[3]) return i;
        }
        return -1;
    }

    /**
     * 各矩形が属するアートボードの番号を返す
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @returns {number[]} アートボードの番号（boundsList と同じ並び）
     */
    function getArtboardIndexes(boundsList) {
        var artboards = app.activeDocument.artboards;
        var artboardRects = [];
        for (var i = 0; i < artboards.length; i++) {
            artboardRects.push(artboards[i].artboardRect);
        }
        var artboardIndexes = [];
        for (var j = 0; j < boundsList.length; j++) {
            artboardIndexes.push(findArtboardIndex(boundsList[j], artboardRects));
        }
        return artboardIndexes;
    }

    /**
     * 2つ以上のアートボードにまたがっているかを返す（どのアートボードにもないものは数えない）
     * @param {number[]} artboardIndexes - 各オブジェクトのアートボード番号
     * @returns {boolean} またがっていれば true
     */
    function spansMultipleArtboards(artboardIndexes) {
        var firstIndex = -1;
        for (var i = 0; i < artboardIndexes.length; i++) {
            if (artboardIndexes[i] < 0) continue;
            if (firstIndex < 0) {
                firstIndex = artboardIndexes[i];
            } else if (artboardIndexes[i] !== firstIndex) {
                return true;
            }
        }
        return false;
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
    // 選択の読み取りと結果の見積もり / Selection snapshot and result plan
    // =========================================

    /**
     * @typedef {Object} SelectionInfo ダイアログを開く前に読み取った選択の情報
     * @property {PageItem[]} items - 選択オブジェクト（前面から背面の順）
     * @property {number} itemCount - 選択オブジェクトの数
     * @property {Array<number[]>} boundsList - 各オブジェクトの geometricBounds（クリップグループはマスクの範囲）
     * @property {boolean[]} isPath - 各オブジェクトがパスかどうか
     * @property {Array<number[]>} imageBoundsList - 「配置画像のみをクリップ」の対象になる画像の visibleBounds
     * @property {number[]} artboardIndexes - 各オブジェクトのアートボード番号
     * @property {boolean} hasMultipleArtboards - ドキュメントにアートボードが2つ以上あるか
     * @property {Array<string|null>|null} appearanceKeys - 各オブジェクトの塗りと線のキー（「アピアランス」を初めて選んだときに読む）
     */

    /**
     * 選択の情報を一度だけ読み取る（件数・プレビュー・実行で共有する）
     * @param {PageItem[]|null} selectedItems - 選択オブジェクト
     * @returns {SelectionInfo} 選択の情報
     */
    function readSelection(selectedItems) {
        var selectionInfo = {
            items: selectedItems || [],
            itemCount: 0,
            boundsList: [],
            isPath: [],
            imageBoundsList: [],
            artboardIndexes: [],
            hasMultipleArtboards: false,
            appearanceKeys: null
        };
        if (!selectedItems) return selectionInfo;

        selectionInfo.itemCount = selectedItems.length;
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            /* クリップグループはマスクの範囲で測る（隠れた部分でまとまりをつながない）
               Clip groups are measured by their mask, so hidden parts do not link clusters */
            selectionInfo.boundsList.push(getClipAwareBounds(selectedItem, false) || selectedItem.geometricBounds);
            selectionInfo.isPath.push(selectedItem.typename === "PathItem");
            /* 「配置画像のみをクリップ」の対象：選択中の画像と、クリップグループ直下の画像
               Targets of Clip Placed Images Only: selected images and images directly inside clipping groups */
            if (isImageItem(selectedItem)) {
                selectionInfo.imageBoundsList.push(selectedItem.visibleBounds);
            } else if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
                var childImages = getChildImages(selectedItem);
                for (var j = 0; j < childImages.length; j++) {
                    selectionInfo.imageBoundsList.push(childImages[j].visibleBounds);
                }
            }
        }
        selectionInfo.artboardIndexes = getArtboardIndexes(selectionInfo.boundsList);
        selectionInfo.hasMultipleArtboards = app.activeDocument.artboards.length > 1;
        return selectionInfo;
    }

    /**
     * 選択を、つなぎ方に従ってまとまりの番号に分ける
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @param {ClusterRule} clusterRule - つなぎ方
     * @returns {Array<number[]>} まとまりごとの番号の配列
     */
    function collectSelectionClusters(selectionInfo, clusterRule) {
        var artboardIndexes = clusterRule.perArtboard ? selectionInfo.artboardIndexes : null;
        if (clusterRule.byAppearance) return collectKeyClusters(getAppearanceKeys(selectionInfo), artboardIndexes);
        return collectClusterIndexes(selectionInfo.boundsList, clusterRule, artboardIndexes);
    }

    /**
     * まとまりにパスが含まれるかを返す
     * @param {boolean[]} isPath - 選択の各オブジェクトがパスかどうか
     * @param {number[]} memberIndexes - まとまりの番号
     * @returns {boolean} パスがあれば true
     */
    function hasPathMember(isPath, memberIndexes) {
        for (var i = 0; i < memberIndexes.length; i++) {
            if (isPath[memberIndexes[i]]) return true;
        }
        return false;
    }

    /**
     * まとまりの外枠（すべてを囲む矩形）を返す
     * @param {Array<number[]>} boundsList - 各オブジェクトの geometricBounds
     * @param {number[]} memberIndexes - まとまりの番号
     * @returns {number[]} 外枠 [左, 上, 右, 下]
     */
    function getUnionBounds(boundsList, memberIndexes) {
        var unionBounds = boundsList[memberIndexes[0]].slice(0);
        for (var i = 1; i < memberIndexes.length; i++) {
            var bounds = boundsList[memberIndexes[i]];
            unionBounds[0] = Math.min(unionBounds[0], bounds[0]);
            unionBounds[1] = Math.max(unionBounds[1], bounds[1]);
            unionBounds[2] = Math.max(unionBounds[2], bounds[2]);
            unionBounds[3] = Math.min(unionBounds[3], bounds[3]);
        }
        return unionBounds;
    }

    /**
     * 選ばれたモードで作成されるグループの外枠を返す（実行時と同じ判定を使う）
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @param {string} mode - モードのキー
     * @param {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} groupingOptions - ダイアログの設定
     * @returns {Array<number[]>} 作成されるグループごとの外枠
     */
    function planResultBounds(selectionInfo, mode, groupingOptions) {
        if (mode === "clipPlacedOnly") return selectionInfo.imageBoundsList;
        var clusterRule = getClusterRule(mode, groupingOptions);
        if (!clusterRule) return [];

        var clusterIndexes = collectSelectionClusters(selectionInfo, clusterRule);
        var resultBounds = [];
        for (var i = 0; i < clusterIndexes.length; i++) {
            if (clusterIndexes[i].length < 2) continue;
            /* クリップはパスを含むまとまりだけが対象 / Clipping needs a path in the cluster */
            if (MASK_SIDES[mode] && !hasPathMember(selectionInfo.isPath, clusterIndexes[i])) continue;
            resultBounds.push(getUnionBounds(selectionInfo.boundsList, clusterIndexes[i]));
        }
        return resultBounds;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューの線の色を作る（ドキュメントのカラーモードに合わせる）
     * @param {Document} doc - 対象のドキュメント
     * @returns {RGBColor|CMYKColor} 線の色
     */
    function createPreviewColor(doc) {
        var previewColor;
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            previewColor = new CMYKColor();
            previewColor.cyan = PREVIEW_CMYK[0];
            previewColor.magenta = PREVIEW_CMYK[1];
            previewColor.yellow = PREVIEW_CMYK[2];
            previewColor.black = PREVIEW_CMYK[3];
        } else {
            previewColor = new RGBColor();
            previewColor.red = PREVIEW_RGB[0];
            previewColor.green = PREVIEW_RGB[1];
            previewColor.blue = PREVIEW_RGB[2];
        }
        return previewColor;
    }

    /**
     * 作成されるグループの範囲を枠で示すプレビューを用意する（最上位の専用レイヤーに描き、閉じるときにレイヤーごと消す）
     * @param {Document} doc - 対象のドキュメント
     * @returns {{show: function(Array<number[]>): void, remove: function(): void}} プレビューの操作
     */
    function createGroupPreview(doc) {
        var previousActiveLayer = doc.activeLayer;
        var previewColor = createPreviewColor(doc);
        var previewLayer = null;

        /* 外枠ごとに枠を描く / Draw one frame per result */
        function show(resultBounds) {
            if (!previewLayer) {
                previewLayer = doc.layers.add();
                previewLayer.name = PREVIEW_LAYER_NAME;
                /* 最上位レイヤーの一番上に置き、どのオブジェクトより手前に枠を出す
                   Keep the layer at the very top so the frames sit in front of every object */
                previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
            }
            for (var i = previewLayer.pageItems.length - 1; i >= 0; i--) {
                previewLayer.pageItems[i].remove();
            }
            for (var j = 0; j < resultBounds.length; j++) {
                var bounds = resultBounds[j];
                var previewFrame = previewLayer.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]);
                previewFrame.filled = false;
                previewFrame.stroked = true;
                previewFrame.strokeColor = previewColor;
                previewFrame.strokeWidth = PREVIEW_STROKE_WIDTH;
                previewFrame.opacity = PREVIEW_OPACITY;
            }
            app.redraw();
        }

        /* プレビューのレイヤーを消して、アクティブレイヤーを戻す / Remove the preview layer and restore the active layer */
        function remove() {
            if (!previewLayer) return;
            previewLayer.remove();
            previewLayer = null;
            doc.activeLayer = previousActiveLayer;
            app.redraw();
        }

        return { show: show, remove: remove };
    }

    // =========================================
    // クリッピングマスク / Clipping mask
    // =========================================

    /**
     * まとまりの中から型にするパスを探す
     * @param {PageItem[]} clusterItems - まとまり（前面から背面の順）
     * @param {string} maskSide - "front" なら最前面、"back" なら最背面のパス
     * @returns {PathItem|null} 型にするパス。パスがなければ null
     */
    function findMaskPath(clusterItems, maskSide) {
        var fromBack = (maskSide === "back");
        for (var i = 0; i < clusterItems.length; i++) {
            var candidateItem = clusterItems[fromBack ? clusterItems.length - 1 - i : i];
            if (candidateItem.typename === "PathItem") return candidateItem;
        }
        return null;
    }

    /**
     * まとまりを、指定のパスを型にしたクリップグループにする
     * @param {PageItem[]} clusterItems - まとまり（前面から背面の順）
     * @param {PathItem|null} maskPath - 型にするパス
     * @returns {void}
     */
    function clipCluster(clusterItems, maskPath) {
        if (clusterItems.length <= 1 || !maskPath) return;

        /* 型にするパスがあった位置にグループを置く / Put the group where the mask path was */
        var clipGroup = app.activeDocument.groupItems.add();
        clipGroup.move(maskPath, ElementPlacement.PLACEBEFORE);
        maskPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);

        /* 残りを前面から順に末尾へ足し、元の重ね順を保つ / Append the rest front to back to keep the stacking order */
        for (var i = 0; i < clusterItems.length; i++) {
            if (clusterItems[i] !== maskPath) {
                clusterItems[i].move(clipGroup, ElementPlacement.PLACEATEND);
            }
        }
        clipGroup.clipped = true;
    }

    /**
     * まとまりごとにクリッピングマスクを作成する
     * @param {Array<PageItem[]>} clusters - まとまりの配列
     * @param {string} maskSide - "front" なら最前面、"back" なら最背面のパスを型にする
     * @returns {void}
     */
    function clipEachCluster(clusters, maskSide) {
        for (var i = 0; i < clusters.length; i++) {
            clipCluster(clusters[i], findMaskPath(clusters[i], maskSide));
        }
    }

    // =========================================
    // 配置画像のクリップ / Placed image clipping
    // =========================================

    /**
     * 画像を、表示範囲と同じ大きさの長方形でクリップする（画像があった位置に置く）
     * @param {PlacedItem|RasterItem} image - クリップする画像
     * @returns {GroupItem} 作成したクリップグループ
     */
    function maskImageWithBounds(image) {
        var imageLayer = image.layer;
        var wasLocked = imageLayer.locked;
        var wasTemplate = imageLayer.isTemplate;
        if (wasLocked) imageLayer.locked = false;
        if (wasTemplate) imageLayer.isTemplate = false;

        var bounds = image.visibleBounds;
        var maskRect = imageLayer.pathItems.rectangle(bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]);
        maskRect.stroked = false;
        maskRect.filled = false;

        /* 画像があった位置にグループを置き、画像→長方形の順に先頭へ入れて長方形を最前面にする
           Put the group where the image was; add the image, then the rectangle, so the rectangle ends up on top */
        var maskedGroup = imageLayer.groupItems.add();
        maskedGroup.move(image, ElementPlacement.PLACEBEFORE);
        image.moveToBeginning(maskedGroup);
        maskRect.moveToBeginning(maskedGroup);
        maskedGroup.clipped = true;

        if (wasLocked) imageLayer.locked = true;
        if (wasTemplate) imageLayer.isTemplate = true;
        return maskedGroup;
    }

    /**
     * グループの直下にある画像を返す
     * @param {GroupItem} parentGroup - 調べるグループ
     * @returns {Array<PlacedItem|RasterItem>} 画像（前面から背面の順）
     */
    function getChildImages(parentGroup) {
        var childImages = [];
        for (var i = 0; i < parentGroup.pageItems.length; i++) {
            if (isImageItem(parentGroup.pageItems[i])) childImages.push(parentGroup.pageItems[i]);
        }
        return childImages;
    }

    /**
     * クリップグループの元のマスクを削除し、中の画像を1つずつクリップし直す（画像以外はそのまま残す）
     * @param {GroupItem} clipGroup - クリップグループ
     * @returns {GroupItem[]} 作成したクリップグループ
     */
    function remaskClipGroup(clipGroup) {
        /* 元のマスクはクリップグループの最前面にある / The old mask is the topmost item of the clipping group */
        var oldMask = clipGroup.pageItems[0];
        clipGroup.clipped = false;
        if (!isImageItem(oldMask)) oldMask.remove();

        var childImages = getChildImages(clipGroup);
        var maskedGroups = [];
        for (var j = 0; j < childImages.length; j++) {
            maskedGroups.push(maskImageWithBounds(childImages[j]));
        }

        /* 画像以外が残っていなければ、作成したグループを前面から順に外へ出して元のグループを消す
           If nothing else is left, move the new groups out front to back and remove the original group */
        if (clipGroup.pageItems.length === maskedGroups.length) {
            for (var k = 0; k < maskedGroups.length; k++) {
                maskedGroups[k].move(clipGroup, ElementPlacement.PLACEBEFORE);
            }
            clipGroup.remove();
        }
        return maskedGroups;
    }

    /**
     * 選択中の画像とクリップグループ内の画像を、それぞれの大きさでクリップする
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {void}
     */
    function clipPlacedImages(selectedItems) {
        var maskedGroups = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
                maskedGroups = maskedGroups.concat(remaskClipGroup(selectedItem));
            } else if (isImageItem(selectedItem)) {
                maskedGroups.push(maskImageWithBounds(selectedItem));
            }
        }

        /* 作成したグループを選択する / Select the new groups */
        selectOnly(maskedGroups);
    }

    // =========================================
    // グループ化 / Grouping
    // =========================================

    /**
     * まとまりを1つのグループにする
     * @param {PageItem[]} clusterItems - まとまり（前面から背面の順）
     * @returns {GroupItem} 作成したグループ
     */
    function groupCluster(clusterItems) {
        /* 最前面のオブジェクトの位置にグループを置く / Put the group where the frontmost item was */
        var newGroup = app.activeDocument.groupItems.add();
        newGroup.move(clusterItems[0], ElementPlacement.PLACEBEFORE);

        /* 前面から順に末尾へ足し、元の重ね順を保つ / Append front to back to keep the stacking order */
        for (var i = 0; i < clusterItems.length; i++) {
            clusterItems[i].move(newGroup, ElementPlacement.PLACEATEND);
        }
        return newGroup;
    }

    /**
     * 2つ以上のオブジェクトを含むまとまりをグループ化し、作成したグループを選択する
     * @param {Array<PageItem[]>} clusters - まとまりの配列
     * @returns {void}
     */
    function groupEachCluster(clusters) {
        var newGroups = [];
        for (var i = 0; i < clusters.length; i++) {
            if (clusters[i].length > 1) newGroups.push(groupCluster(clusters[i]));
        }
        if (newGroups.length) selectOnly(newGroups);
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * しきい値を表示用の文字列にする
     * @param {number} thresholdValue - しきい値（pt）
     * @returns {string} 「10 pt」の形の文字列
     */
    function formatThreshold(thresholdValue) {
        return Math.round(thresholdValue) + " pt";
    }

    /**
     * モードのラジオボタンを並べたパネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} panelTitle - パネルの見出し
     * @param {string[]} modeKeys - 並べるモードのキー
     * @param {Object} modeRadios - 作成したラジオボタンをモードのキーで登録する先
     * @param {number} [columnCount] - 列の数（省略時は1列）。左上から横へ順に並べる
     * @returns {Panel} 追加したパネル
     */
    function addModePanel(parentWindow, panelTitle, modeKeys, modeRadios, columnCount) {
        var modePanel = parentWindow.add("panel", undefined, panelTitle);
        setupPanel(modePanel, 6);
        modePanel.alignChildren = "left";

        /* 複数列のときは列ごとのグループを横に並べる / For multiple columns, lay out one group per column side by side */
        var radioColumns = [modePanel];
        if (columnCount > 1) {
            var columnsRow = modePanel.add("group");
            columnsRow.orientation = "row";
            columnsRow.alignChildren = ["left", "top"];
            columnsRow.spacing = RADIO_COLUMN_SPACING;
            radioColumns = [];
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                var radioColumn = columnsRow.add("group");
                radioColumn.orientation = "column";
                radioColumn.alignChildren = "left";
                radioColumns.push(radioColumn);
            }
        }

        for (var i = 0; i < modeKeys.length; i++) {
            var modeKey = modeKeys[i];
            var modeRadio = radioColumns[i % radioColumns.length].add("radiobutton", undefined, getLabel(LABELS.radio[modeKey]));
            modeRadio.helpTip = getLabel(LABELS.tooltip[modeKey]);
            modeRadios[modeKey] = modeRadio;
        }
        return modePanel;
    }

    /**
     * しきい値の見出し・スライダー・値の表示を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @returns {{caption: StaticText, slider: Slider, valueText: StaticText}} 追加したコントロール
     */
    function addThresholdControls(parentPanel) {
        var thresholdTip = getLabel(LABELS.tooltip.threshold);

        var thresholdCaption = parentPanel.add("statictext", undefined, labelText(LABELS.fieldLabel.threshold));
        thresholdCaption.helpTip = thresholdTip;

        /* パネルの幅に合わせて伸ばし、値の表示をスライダーの中央に揃える / Stretch to the panel width so the readout stays centered under the slider */
        var thresholdSlider = parentPanel.add("slider", undefined, DEFAULT_THRESHOLD, THRESHOLD_MIN, THRESHOLD_MAX);
        thresholdSlider.preferredSize.width = SLIDER_WIDTH;
        thresholdSlider.alignment = "fill";
        thresholdSlider.helpTip = thresholdTip;

        var thresholdValueText = parentPanel.add("statictext", undefined, formatThreshold(DEFAULT_THRESHOLD));
        thresholdValueText.alignment = "center";
        thresholdValueText.characters = VALUE_TEXT_CHARS;
        thresholdValueText.helpTip = thresholdTip;

        return { caption: thresholdCaption, slider: thresholdSlider, valueText: thresholdValueText };
    }

    /**
     * チェックボックスを追加する
     * @param {Window|Panel} parentPanel - 追加先のダイアログかパネル
     * @param {string} optionKey - LABELS.checkbox / LABELS.tooltip のキー
     * @param {boolean} initialValue - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, optionKey, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(LABELS.checkbox[optionKey]));
        optionCheckbox.value = initialValue;
        optionCheckbox.helpTip = getLabel(LABELS.tooltip[optionKey]);
        return optionCheckbox;
    }

    /**
     * 選択数と、作成されるグループ数の表示を追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {number} itemCount - 選択しているオブジェクトの数
     * @returns {StaticText} 作成されるグループ数の表示
     */
    function addCountReadout(parentWindow, itemCount) {
        var countGroup = parentWindow.add("group");
        countGroup.orientation = "column";
        countGroup.alignment = "fill";
        countGroup.alignChildren = "fill";
        countGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.selectedCount) + itemCount);
        var resultCountText = countGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.groupCount) + 0);
        resultCountText.helpTip = getLabel(LABELS.tooltip.resultCount);
        return resultCountText;
    }

    /**
     * @typedef {Object} DialogControls ダイアログと、設定を読み書きするコントロール
     * @property {Window} dialogWindow - ダイアログ
     * @property {Object} modeRadios - モードのキーで引けるラジオボタン
     * @property {{caption: StaticText, slider: Slider, valueText: StaticText}} thresholdControls - しきい値の見出し・スライダー・値の表示
     * @property {Checkbox} splitAtGapsCheckbox - 「間隔が空いたら分ける」
     * @property {Checkbox} perArtboardCheckbox - 「アートボードごと」
     * @property {StaticText} resultCountText - 作成されるグループ数の表示
     * @property {Checkbox} previewCheckbox - 「プレビューを表示」
     */

    /**
     * ダイアログを組み立てる（表示はしない）
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @returns {DialogControls} ダイアログとコントロール
     */
    function buildDialog(selectionInfo) {
        var clipAndGroupDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(clipAndGroupDialog);
        clipAndGroupDialog.alignChildren = "left";

        var modeRadios = {};
        addModePanel(clipAndGroupDialog, getLabel(LABELS.panel.clip), CLIP_MODES, modeRadios);
        var groupingPanel = addModePanel(clipAndGroupDialog, getLabel(LABELS.panel.grouping), GROUP_MODES, modeRadios, 2);
        groupingPanel.alignment = "fill";  /* ダイアログの幅いっぱいに広げる / Stretch to the dialog width */
        var thresholdControls = addThresholdControls(groupingPanel);
        var splitAtGapsCheckbox = addOptionCheckbox(groupingPanel, "splitAtGaps", DEFAULT_SPLIT_AT_GAPS);
        /* 選択が複数のアートボードにまたがるときだけ最初からON / On by default only when the selection spans artboards */
        var perArtboardCheckbox = addOptionCheckbox(groupingPanel, "perArtboard",
            selectionInfo.hasMultipleArtboards && spansMultipleArtboards(selectionInfo.artboardIndexes));
        var resultCountText = addCountReadout(clipAndGroupDialog, selectionInfo.itemCount);
        var previewCheckbox = addOptionCheckbox(clipAndGroupDialog, "showPreview", DEFAULT_SHOW_PREVIEW);

        /* ボタンエリア / Button row */
        var buttonRow = addButtonRow(clipAndGroupDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            dialogWindow: clipAndGroupDialog,
            modeRadios: modeRadios,
            thresholdControls: thresholdControls,
            splitAtGapsCheckbox: splitAtGapsCheckbox,
            perArtboardCheckbox: perArtboardCheckbox,
            resultCountText: resultCountText,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * 選択中のモードのキーを返す
     * @param {Object} modeRadios - モードのキーで引けるラジオボタン
     * @returns {string|null} モードのキー（どれも選ばれていなければ null）
     */
    function getSelectedMode(modeRadios) {
        for (var i = 0; i < ALL_MODES.length; i++) {
            if (modeRadios[ALL_MODES[i]].value) return ALL_MODES[i];
        }
        return null;
    }

    /**
     * ダイアログの設定を読み取る（しきい値は表示と実行で同じく整数に丸める）
     * @param {DialogControls} dialogControls - ダイアログとコントロール
     * @param {boolean} hasMultipleArtboards - ドキュメントにアートボードが2つ以上あるか
     * @returns {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} ダイアログの設定
     */
    function readGroupingOptions(dialogControls, hasMultipleArtboards) {
        return {
            threshold: Math.round(dialogControls.thresholdControls.slider.value),
            splitAtGaps: dialogControls.splitAtGapsCheckbox.value,
            perArtboard: dialogControls.perArtboardCheckbox.value && hasMultipleArtboards
        };
    }

    /**
     * 選んだモードで使う設定だけを有効にする
     * @param {DialogControls} dialogControls - ダイアログとコントロール
     * @param {boolean} hasMultipleArtboards - ドキュメントにアートボードが2つ以上あるか
     * @returns {void}
     */
    function updateOptionControls(dialogControls, hasMultipleArtboards) {
        var usedOptions = MODE_OPTIONS[getSelectedMode(dialogControls.modeRadios)] || {};
        var thresholdEnabled = usedOptions.threshold === true;
        dialogControls.thresholdControls.caption.enabled = thresholdEnabled;
        dialogControls.thresholdControls.slider.enabled = thresholdEnabled;
        dialogControls.thresholdControls.valueText.enabled = thresholdEnabled;
        dialogControls.splitAtGapsCheckbox.enabled = usedOptions.splitAtGaps === true;
        dialogControls.perArtboardCheckbox.enabled = usedOptions.perArtboard === true && hasMultipleArtboards;
    }

    /**
     * ダイアログを表示して、選ばれたモードと設定を返す（開いている間は結果の範囲をプレビューする）
     * @param {string} initialMode - 最初に選んでおくモードのキー
     * @param {SelectionInfo} selectionInfo - 選択の情報
     * @returns {{mode: string, groupingOptions: Object}|null} キャンセル時は null
     */
    function showDialog(initialMode, selectionInfo) {
        var dialogControls = buildDialog(selectionInfo);
        var modeRadios = dialogControls.modeRadios;
        var thresholdControls = dialogControls.thresholdControls;
        var previewCheckbox = dialogControls.previewCheckbox;
        var hasMultipleArtboards = selectionInfo.hasMultipleArtboards;

        var groupPreview = selectionInfo.itemCount ? createGroupPreview(app.activeDocument) : null;
        previewCheckbox.enabled = (groupPreview !== null);

        /* 作成されるグループの数と範囲を表示する（モードか設定が変わったときだけ数え直す）
           Show the number and extent of the groups to be created; recount only when the mode or settings change */
        var lastResultKey = null;
        var lastResultBounds = [];
        function refreshResult() {
            var mode = getSelectedMode(modeRadios);
            var groupingOptions = readGroupingOptions(dialogControls, hasMultipleArtboards);
            var resultKey = [mode, groupingOptions.threshold, groupingOptions.splitAtGaps, groupingOptions.perArtboard].join(":");
            if (resultKey === lastResultKey) return;
            lastResultKey = resultKey;

            lastResultBounds = planResultBounds(selectionInfo, mode, groupingOptions);
            var countLabel = (MASK_SIDES[mode] || mode === "clipPlacedOnly") ? LABELS.fieldLabel.clipCount : LABELS.fieldLabel.groupCount;
            dialogControls.resultCountText.text = labelText(countLabel) + lastResultBounds.length;
            updatePreview();
        }

        /* 「プレビューを表示」がONのときだけ枠を描き、OFFなら消す / Draw the frames only while "Show preview" is on; remove them otherwise */
        function updatePreview() {
            if (!groupPreview) return;
            if (previewCheckbox.value) {
                groupPreview.show(lastResultBounds);
            } else {
                groupPreview.remove();
            }
        }

        /* ラジオボタンは2つのパネルに分かれているため、排他は手動で行う
           The radios span two panels, so exclusivity is handled by hand */
        function selectMode() {
            for (var i = 0; i < ALL_MODES.length; i++) {
                modeRadios[ALL_MODES[i]].value = (modeRadios[ALL_MODES[i]] === this);
            }
            updateOptionControls(dialogControls, hasMultipleArtboards);
            refreshResult();
        }

        for (var i = 0; i < ALL_MODES.length; i++) {
            modeRadios[ALL_MODES[i]].onClick = selectMode;
        }
        dialogControls.splitAtGapsCheckbox.onClick = refreshResult;
        dialogControls.perArtboardCheckbox.onClick = refreshResult;
        previewCheckbox.onClick = updatePreview;
        /* 選択が多いときは、ドラッグ中は値の表示だけ更新し、指を離したときに数える
           With a large selection, update only the readout while dragging and recount on release */
        thresholdControls.slider.onChanging = function () {
            thresholdControls.valueText.text = formatThreshold(thresholdControls.slider.value);
            if (selectionInfo.itemCount <= LIVE_COUNT_MAX_ITEMS) refreshResult();
        };
        thresholdControls.slider.onChange = refreshResult;

        modeRadios[initialMode].value = true;
        updateOptionControls(dialogControls, hasMultipleArtboards);
        refreshResult();

        addVersionFooter(dialogControls.dialogWindow);
        prepareDialogWindow(dialogControls.dialogWindow, SCRIPT_NAME);
        /* OK・キャンセルのどちらで閉じても、プレビューは実行前に消す / Remove the preview before running, however the dialog closes */
        var dialogReturnCode = dialogControls.dialogWindow.show();
        if (groupPreview) groupPreview.remove();
        if (dialogReturnCode !== 1) return null;
        return {
            mode: getSelectedMode(modeRadios),
            groupingOptions: readGroupingOptions(dialogControls, hasMultipleArtboards)
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選ばれたモードの処理を実行する
     * @param {string} mode - モードのキー
     * @param {{threshold: number, splitAtGaps: boolean, perArtboard: boolean}} groupingOptions - ダイアログの設定
     * @param {SelectionInfo} selectionInfo - ダイアログを開く前に読み取った選択の情報
     * @returns {void}
     */
    function runMode(mode, groupingOptions, selectionInfo) {
        if (!selectionInfo.itemCount) return;
        if (mode === "clipPlacedOnly") {
            clipPlacedImages(selectionInfo.items);
            return;
        }

        var clusterRule = getClusterRule(mode, groupingOptions);
        if (!clusterRule) return;

        /* プレビューと同じ判定でまとまりを作る / Build the clusters with the same rule as the preview */
        var clusterIndexes = collectSelectionClusters(selectionInfo, clusterRule);
        var clusters = [];
        for (var i = 0; i < clusterIndexes.length; i++) {
            var clusterItems = [];
            for (var j = 0; j < clusterIndexes[i].length; j++) {
                clusterItems.push(selectionInfo.items[clusterIndexes[i][j]]);
            }
            clusters.push(clusterItems);
        }

        if (MASK_SIDES[mode]) {
            clipEachCluster(clusters, MASK_SIDES[mode]);
        } else {
            groupEachCluster(clusters);
        }
    }

    /**
     * ダイアログを表示して、選ばれた処理を実行する
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        var selectionInfo = readSelection(selectedItems);
        /* すべて画像なら「配置画像のみをクリップ」を初期選択にする / Preselect "Clip Placed Images Only" when every item is an image */
        var initialMode = isAllImages(selectedItems) ? "clipPlacedOnly" : "groupProximity";
        var dialogSettings = showDialog(initialMode, selectionInfo);
        if (dialogSettings) runMode(dialogSettings.mode, dialogSettings.groupingOptions, selectionInfo);
    }

    main();

})();
