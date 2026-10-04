#target illustrator
#targetengine "session"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、指定した縦横比に合わせてサイズ変更します。
幅か高さを固定してもう一方を比率・長さ・％で決めるほか、幅・高さを別々に指定することもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4a212e6eacf1

### Overview

Resizes the selected objects to a chosen aspect ratio.
Fix the width or height and set the other side by ratio, length or percentage, or set width and height separately.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AspectRatioScaler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AspectRatioScaler";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AspectRatioScaler.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AspectRatioScaler.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4a212e6eacf1"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プリセットの比率（横 ÷ 縦）/ Preset ratios (width / height) */
    var RATIO_16_9   = 16 / 9;
    var RATIO_SQUARE = 1;
    var RATIO_A4     = 210 / 297;

    /* カスタム比率の初期値 / Initial custom ratio */
    var DEFAULT_CUSTOM_RATIO_WIDTH  = "3";
    var DEFAULT_CUSTOM_RATIO_HEIGHT = "2";

    /* 結果の縦横比を整数比で出すときの分母の上限と許容誤差（相対）/ Max denominator and relative tolerance when showing the resulting ratio as integers */
    var RATIO_PAIR_MAX_DENOMINATOR = 16;
    var RATIO_PAIR_TOLERANCE       = 0.002;

    /* 選択なしで作る長方形の幅（単位コード → 定規の単位での値）/ Width of the rectangle drawn with nothing selected (unit code -> value in ruler units) */
    var DEFAULT_RECT_WIDTH_BY_UNIT = { 1: 100, 6: 1000 };

    /* 上記にない単位のときの長方形の幅（pt）/ Rectangle width (pt) for other units */
    var FALLBACK_RECT_WIDTH_PT = 200;

    /* 基準点の初期値（0..8 を行優先、0=左上・4=中央・8=右下）/ Initial reference point (row-major 0..8; 0=top-left, 4=center, 8=bottom-right) */
    var DEFAULT_ANCHOR_INDEX = 4;

    /* ［ピクセルグリッドに最適化］の初期状態 / Initial state of "Make Pixel Perfect" */
    var DEFAULT_ALIGN_TO_PIXEL_GRID = true;

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

    var FIELD_CHARACTERS    = 5;                 /* 数値欄の幅（文字数）/ Numeric field width */
    var CUSTOM_RATIO_CHARS  = 3;                 /* カスタム比の欄の幅（文字数）/ Custom ratio field width */
    var PERCENT_CHARS       = 5;                 /* ％の欄の幅（文字数）/ Percent field width */
    var CUSTOM_RATIO_INDENT = 14;                /* カスタム比の欄の左インデント（約1文字）/ Left indent of the custom ratio fields (about one character) */
    var ORIENT_BUTTON_SIZE  = 36;                /* 向きアイコンのボタンの大きさ / Size of an orientation icon button */
    var ORIENT_FRAME_LONG   = 30;                /* 向きアイコンの枠の長辺 / Long side of the orientation icon frame */
    var ORIENT_FRAME_SHORT  = 23;                /* 向きアイコンの枠の短辺 / Short side of the orientation icon frame */
    var ORIENT_ANCHOR_GAP   = 24;                /* 向きアイコンと9軸の間隔 / Gap between the orientation icons and the 9-axis widget */

    /**
     * 見出し付きパネルを縦並びで追加する
     * @param {Group} parent - 追加先
     * @param {string} title - パネルの見出し
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, title, spacing) {
        var titledPanel = parent.add("panel", undefined, title);
        setupPanel(titledPanel, spacing);
        titledPanel.alignment = ["fill", "top"];
        return titledPanel;
    }

    /**
     * 子を縦に並べるグループを追加する
     * @param {Object} parent - 追加先のパネルまたはグループ
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parent) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "left";
        return columnGroup;
    }

    /**
     * 子を横に並べ、天地中央でそろえるグループを追加する
     * @param {Object} parent - 追加先のパネルまたはグループ
     * @returns {Group} 追加したグループ
     */
    function addRowGroup(parent) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = ["left", "top"];
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
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

    /* 入力された単位を欄の単位へ換算するための、1単位あたりのポイント数（値は UNITS 表と同じ。キーは小文字）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Points per unit for converting typed units into the field's unit (same values as the UNITS table; lowercase keys) */
    var STEPPER_POINTS_PER_UNIT = {
        "in": 72, "inch": 72, "mm": 72 / 25.4, "cm": 72 / 2.54, "m": 72 / 25.4 * 1000,
        "pt": 1, "px": 1, "p": 12, "pc": 12, "pica": 12,
        "q": 72 / 25.4 * 0.25, "h": 72 / 25.4 * 0.25, "ft": 72 * 12, "yd": 72 * 36
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
            if (!hasOperator && value === parseFloat(numberInput.text)) return;
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
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
     * 数値の後ろの単位は欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
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
        var fieldPointsPerUnit = STEPPER_POINTS_PER_UNIT[fieldUnitKey];
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
            var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var pointsPerUnit = STEPPER_POINTS_PER_UNIT[unitKey];
            if (pointsPerUnit === undefined || fieldPointsPerUnit === undefined) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = value * pointsPerUnit;
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldPointsPerUnit;
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
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

    // =========================================
    // 自作描画のウィジェット共通 / Custom-drawn widgets (shared)
    // =========================================

    /* UI が明るいテーマか（ステップボタンの判定を流用）/ Whether the UI uses a light theme (reuses the stepper's check) */
    var IS_LIGHT_UI = !STEPPER_UI_DARK;

    /**
     * 矩形を塗る（多角形は fillPath で塗れないので rectPath を使う）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} rect - [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function fillRect(graphics, rect, color) {
        graphics.newPath();
        graphics.rectPath(rect[0], rect[1], rect[2], rect[3]);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
    }

    /**
     * 矩形の輪郭を描く
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} rect - [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @param {number} lineWidth - 線幅
     * @returns {void}
     */
    function strokeRect(graphics, rect, color, lineWidth) {
        graphics.newPath();
        graphics.rectPath(rect[0], rect[1], rect[2], rect[3]);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, lineWidth));
    }

    /**
     * 楕円を塗る
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} rect - 外接矩形 [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function fillEllipse(graphics, rect, color) {
        graphics.newPath();
        graphics.ellipsePath(rect[0], rect[1], rect[2], rect[3]);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
    }

    /**
     * コントロールの地をパネルと同じ色で塗って透過に見せる
     * @param {Object} control - 対象のコントロール
     * @returns {void}
     */
    function paintControlBackground(control) {
        var graphics = control.graphics;
        /* backgroundColor が無い環境では例外 / Throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, control.size[0], control.size[1]);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}
    }

    /**
     * コントロールを再描画する
     * @param {Object} control - 対象のコントロール
     * @returns {void}
     */
    function redrawControl(control) {
        /* notify("onDraw") は環境により例外になる / Throws on some versions */
        try { control.notify("onDraw"); } catch (e) {}
    }

    // 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)

    // -----------------------------------------
    // 基準点ウィジェットの寸法 / Anchor widget metrics
    // -----------------------------------------
    var ANCHOR_WIDGET_SIZE      = 66;   /* ウィジェット全体の一辺 / overall size of the widget */
    var ANCHOR_WIDGET_CELL_SIZE = 9;    /* □1個の一辺 / size of one square */
    var ANCHOR_WIDGET_CELL_GAP  = 7.5;  /* □どうしの間隔 / gap between squares */
    var ANCHOR_WIDGET_NONE      = -1;   /* 未選択のインデックス / index while nothing is selected */

    /* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
       Cell names in row-major order, matching the Transformation enumeration */
    var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線 / Rules joining the outer squares (the center stands alone) */
    var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    // -----------------------------------------
    // 基準点ウィジェットの配色 / Anchor widget colors
    // -----------------------------------------
    var ANCHOR_WIDGET_UI_DARK = isDarkUI();
    /* 枠線・ケイ線はグレー、選択セルの塗りはライトで濃いグレー・ダークで明るいグレー（既存スクリプトの配色を踏襲）。
       無効時は同じ色を半透明にして背景へ沈める（不透明の薄いグレーだとダークUIで逆に明るく浮くため）
       Gray rules; the selected fill is dark gray on light UI and light gray on dark UI (as in the existing scripts).
       Disabled colors are translucent versions so they sink into any background */
    var ANCHOR_WIDGET_LINE_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 1]     : [0.42, 0.42, 0.42, 1];  /* 枠線・ケイ線 / rules */
    var ANCHOR_WIDGET_FILL_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 1]     : [0.27, 0.27, 0.27, 1];  /* 選択セルの塗り / selected fill */
    var ANCHOR_WIDGET_DIM_LINE_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.7, 0.7, 0.7, 0.4]   : [0.42, 0.42, 0.42, 0.4];  /* 無効時の枠線 / rules when disabled */
    var ANCHOR_WIDGET_DIM_FILL_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.9, 0.9, 0.9, 0.3]   : [0.27, 0.27, 0.27, 0.3];  /* 無効時の塗り / fill when disabled */

    // -----------------------------------------
    // ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 基準点（3×3）を選ぶウィジェットを追加する。クリックしたセルを選び、onChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
     * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
     * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）/ size（一辺。既定 66）
     * @returns {Button} ウィジェット（値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
     */
    function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
        var anchorOptions = widgetOptions || {};
        var widgetSize = anchorOptions.size || ANCHOR_WIDGET_SIZE;
        var anchorWidget = parent.add("button", undefined, "");
        anchorWidget.minimumSize = [widgetSize, widgetSize];
        anchorWidget.preferredSize = [widgetSize, widgetSize];
        anchorWidget.maximumSize = [widgetSize, widgetSize];
        anchorWidget.isAnchorWidget = true; /* redrawAnchorWidgetsIn() の目印 / marker for redrawAnchorWidgetsIn() */
        anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(anchorOptions.disabledCells);
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(initialValue, anchorWidget.anchorAllowNone);
        anchorWidget.onDraw = function () { drawAnchorWidget(anchorWidget); };
        anchorWidget.onClick = function () {}; /* セルの判定は mousedown で行う / hit-testing happens in mousedown */

        /* クリック座標（コントロール基準）を3分割してセルを判定する / split the control-relative click into thirds */
        anchorWidget.addEventListener("mousedown", function (event) {
            if (!isAnchorWidgetEnabledInTree(anchorWidget)) return;
            var cellIndex = getAnchorCellAt(event.clientX, event.clientY, anchorWidget.size[0], anchorWidget.size[1]);
            if (anchorWidget.anchorDisabledCells[cellIndex]) return;
            anchorWidget.anchorWidgetIndex = cellIndex;
            redrawAnchorWidget(anchorWidget);
            if (onChange) onChange(cellIndex, anchorWidget);
        });
        return anchorWidget;
    }

    /**
     * 選択中のセルのインデックスを返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {number} 0〜8（行優先）。未選択なら -1
     */
    function getAnchorWidgetIndex(anchorWidget) {
        return anchorWidget.anchorWidgetIndex;
    }

    /**
     * 選択中のセルの名前を返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {string} "topLeft" など。未選択なら ""
     */
    function getAnchorWidgetName(anchorWidget) {
        return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
    }

    /**
     * 選択するセルを変えて描き直す（onChange は呼ばない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * ウィジェットの有効／無効を切り替えて描き直す（無効の間は薄く描き、クリックも無視する）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
        anchorWidget.enabled = isEnabled;
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * 選べないセルを指定し直して描き直す（選択中のセルは変えない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
     * @returns {void}
     */
    function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(disabledCells);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * コンテナ以下にある基準点ウィジェットをすべて描き直す。パネルや行の enabled を切り替えたあとに呼ぶ
     * @param {Object} container - パネル・グループ・ウィンドウなど
     * @returns {void}
     */
    function redrawAnchorWidgetsIn(container) {
        if (container.isAnchorWidget) {
            redrawAnchorWidget(container);
            return;
        }
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            redrawAnchorWidgetsIn(container.children[i]);
        }
    }

    // -----------------------------------------
    // 値の変換 / Value helpers
    // -----------------------------------------
    /**
     * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
     * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
     * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
     * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
     */
    function resolveAnchorWidgetIndex(anchorValue, allowNone) {
        if (typeof anchorValue === "string") {
            for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
                if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
            }
            return 4;
        }
        if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
        if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
            return anchorValue;
        }
        return 4;
    }

    /**
     * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [横の割合, 縦の割合]
     */
    function getAnchorRatio(anchorValue) {
        var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    }

    /**
     * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
     * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [x, y]
     */
    function getAnchorPointOnBounds(bounds, anchorValue) {
        var anchorRatio = getAnchorRatio(anchorValue);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
            bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
        ];
    }

    /**
     * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
     * 基準は効果を含まない境界（geometricBounds）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {Transformation} Transformation.TOPLEFT など
     */
    function getAnchorTransformation(anchorValue) {
        var transformations = [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ];
        return transformations[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * symbols.add() に渡す登録点を返す（Illustrator 専用）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
     */
    function getAnchorSymbolRegistrationPoint(anchorValue) {
        var registrationPoints = [
            SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
        ];
        return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * クリック位置からセルのインデックスを求める（ウィジェットを縦横3等分し、外にはみ出した座標は端のセルに寄せる）
     * @param {number} clickX - コントロール基準の x
     * @param {number} clickY - コントロール基準の y
     * @param {number} widgetWidth - ウィジェットの幅
     * @param {number} widgetHeight - ウィジェットの高さ
     * @returns {number} 0〜8
     */
    function getAnchorCellAt(clickX, clickY, widgetWidth, widgetHeight) {
        var column = Math.min(2, Math.max(0, Math.floor(clickX / (widgetWidth / 3))));
        var row = Math.min(2, Math.max(0, Math.floor(clickY / (widgetHeight / 3))));
        return row * 3 + column;
    }

    /**
     * セルのインデックスの配列を、9個の真偽値に直す
     * @param {number[]} [cellIndexes] - セルのインデックスの配列
     * @returns {boolean[]} 含まれるセルだけ true
     */
    function toAnchorCellFlags(cellIndexes) {
        var cellFlags = [false, false, false, false, false, false, false, false, false];
        if (!cellIndexes) return cellFlags;
        for (var i = 0; i < cellIndexes.length; i++) {
            if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
        }
        return cellFlags;
    }

    // -----------------------------------------
    // 描画 / Drawing
    // -----------------------------------------
    /**
     * ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立。選択セルだけ塗る）
     * @param {Button} anchorWidget - 描くウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        var halfCell = cellSize / 2;
        /* 自作描画は自動でディムにならないので、親までたどって判定する / custom drawing is not dimmed automatically */
        var isEnabled = isAnchorWidgetEnabledInTree(anchorWidget);

        /* ボタンの地をコントロールの地色で塗り、パネルに溶け込ませる（backgroundColor が無い環境では例外）
           Paint the control's own background so the widget blends into the panel; throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
        var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);
        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1);
        for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
                graphics.lineTo(cellB[0], cellB[1] + halfCell);
            } else {
                /* 縦方向：下の□へ / vertical: to the square below */
                graphics.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
                graphics.lineTo(cellB[0] + halfCell, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < 9; i++) {
            var isCellEnabled = isEnabled && !anchorWidget.anchorDisabledCells[i];
            drawAnchorWidgetCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.anchorWidgetIndex, isCellEnabled);
        }
    }

    /**
     * □を1つ描く（選択中だけ塗り、枠は塗りの上に重ねる）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} cellX - 左端
     * @param {number} cellY - 上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {boolean} isEnabled - 選べるセルなら true（false なら薄く描く）
     * @returns {void}
     */
    function drawAnchorWidgetCell(graphics, cellX, cellY, isSelected, isEnabled) {
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        /* rectPath の前には毎回 newPath()（呼ばないとパスが累積して塗りが線画になる）
           Always call newPath() before rectPath(), or paths accumulate and fills turn into outlines */
        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(cellX, cellY, cellSize, cellSize);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_FILL_COLOR : ANCHOR_WIDGET_DIM_FILL_COLOR));
        }
        graphics.newPath();
        graphics.rectPath(cellX, cellY, cellSize, cellSize);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isAnchorWidgetEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false) return false;
        }
        return true;
    }

    /**
     * ウィジェットの onDraw を呼び直す。notify("onDraw") は環境によって例外や空振りになるため、隠して再表示して描き直させる
     * @param {Button} anchorWidget - 描き直すウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        anchorWidget.hide();
        anchorWidget.show();
    }

    // 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget

    // =========================================
    // 向きアイコン / Orientation icons
    // =========================================

    /* 未選択の枠と人物：グレー、選択中：青 / Unselected frame and figure: gray; selected: blue */
    var ORIENT_IDLE_COLOR = IS_LIGHT_UI ? [0.45, 0.45, 0.45, 1] : [0.7, 0.7, 0.7, 1];
    var ORIENT_SELECTED_COLOR = [0.22, 0.47, 0.9, 1];

    /**
     * 枠の中に胸像の人物を描く（頭は正円、肩は楕円、胴は枠の下端まで）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number[]} frameRect - 枠 [x, y, 幅, 高さ]
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function drawPortraitFigure(graphics, frameRect, color) {
        var centerX = frameRect[0] + frameRect[2] / 2;
        var figureBottom = frameRect[1] + frameRect[3] - 3;
        var headSize = Math.round(frameRect[3] * 0.3);
        var headTop = frameRect[1] + Math.round(frameRect[3] * 0.18);
        var bodyWidth = Math.round(headSize * 1.9);
        var shoulderTop = headTop + headSize + 1;
        var shoulderHeight = Math.round(headSize * 0.9);

        fillEllipse(graphics, [centerX - headSize / 2, headTop, headSize, headSize], color);
        fillEllipse(graphics, [centerX - bodyWidth / 2, shoulderTop, bodyWidth, shoulderHeight], color);
        var torsoTop = shoulderTop + shoulderHeight / 2;
        fillRect(graphics, [centerX - bodyWidth / 2, torsoTop, bodyWidth, figureBottom - torsoTop], color);
    }

    /**
     * 向きアイコンを描画する（横長または縦長の枠＋人物）
     * @param {Button} button - 対象のボタン
     * @returns {void}
     */
    function drawOrientationButton(button) {
        var graphics = button.graphics;
        paintControlBackground(button);

        var frameWidth = button.isPortrait ? ORIENT_FRAME_SHORT : ORIENT_FRAME_LONG;
        var frameHeight = button.isPortrait ? ORIENT_FRAME_LONG : ORIENT_FRAME_SHORT;
        var frameRect = [
            Math.round((button.size[0] - frameWidth) / 2),
            Math.round((button.size[1] - frameHeight) / 2),
            frameWidth,
            frameHeight
        ];
        var iconColor = button.isSelected ? ORIENT_SELECTED_COLOR : ORIENT_IDLE_COLOR;
        strokeRect(graphics, frameRect, iconColor, 2);
        drawPortraitFigure(graphics, frameRect, iconColor);
    }

    /**
     * 縦／横の向きアイコンを2つ並べて追加する（どちらか一方だけが isSelected=true）
     * 選び直すと group.onOrientationChange() を呼ぶ
     * @param {Object} parent - 追加先
     * @param {Object} portraitTip - 縦のツールチップ
     * @param {Object} landscapeTip - 横のツールチップ
     * @returns {Group} 追加したグループ（portraitButton / landscapeButton を持つ）
     */
    function addOrientationButtons(parent, portraitTip, landscapeTip) {
        var orientationGroup = addRowGroup(parent);

        /**
         * 向きアイコンのボタンを1つ追加する
         * @param {boolean} isPortrait - 縦なら true
         * @param {Object} tipSet - ツールチップ
         * @returns {Button} 追加したボタン
         */
        function addOrientationButton(isPortrait, tipSet) {
            var button = orientationGroup.add("button", undefined, "");
            button.helpTip = getLabel(tipSet);
            button.minimumSize = button.preferredSize = button.maximumSize = [ORIENT_BUTTON_SIZE, ORIENT_BUTTON_SIZE];
            button.isPortrait = isPortrait;
            button.isSelected = false;
            button.onDraw = function () { drawOrientationButton(this); };
            button.onClick = function () {
                orientationGroup.landscapeButton.isSelected = !isPortrait;
                orientationGroup.portraitButton.isSelected = isPortrait;
                redrawControl(orientationGroup.landscapeButton);
                redrawControl(orientationGroup.portraitButton);
                if (typeof orientationGroup.onOrientationChange === "function") orientationGroup.onOrientationChange();
            };
            return button;
        }

        /* 縦を左、横を右に並べる / Portrait on the left, landscape on the right */
        orientationGroup.portraitButton = addOrientationButton(true, portraitTip);
        orientationGroup.landscapeButton = addOrientationButton(false, landscapeTip);
        orientationGroup.landscapeButton.isSelected = true;
        return orientationGroup;
    }

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

    /* 定規の単位（実行中は変わらないので一度だけ読む）/ Ruler unit, read once */
    var RULER_UNIT = getUnitInfo();

    /**
     * 定規の単位に合わせて丸める（px は整数、mm は 0.1mm 刻み、その他は 0.01pt 刻み）
     * @param {number} valuePt - 値（pt）
     * @returns {number} 丸めた値（pt）
     */
    function roundForUnit(valuePt) {
        var unitCode = RULER_UNIT.code;
        if (unitCode === 6) return Math.round(valuePt); /* 1px = 1pt */
        if (unitCode === 1) {
            var stepPt = UNITS[1].pointsPerUnit * 0.1;
            return Math.round(valuePt / stepPt) * stepPt;
        }
        return Math.round(valuePt * 100) / 100;
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
            title: { ja: "縦横比を指定してサイズ変更", en: "Resize to Aspect Ratio" }
        },
        panel: {
            aspectRatio: { ja: "縦横比", en: "Aspect Ratio" },
            orientationAnchor: { ja: "向きと基準点", en: "Orientation & Reference Point" },
            size: { ja: "サイズ", en: "Size" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            ratioOriginal: { ja: "元の比率", en: "Original Ratio" },
            ratio16x9: { ja: "16:9", en: "16:9" },
            ratioSquare: { ja: "1:1（スクエア）", en: "1:1 (Square)" },
            ratioA4: { ja: "A4（1:1.414）", en: "A4 (1:1.414)" },
            ratioCustom: { ja: "カスタム", en: "Custom" },
            basisNone: { ja: "なし（自由）", en: "None (Free)" },
            basisHorizontal: { ja: "幅", en: "Width" },
            basisVertical: { ja: "高さ", en: "Height" }
        },
        fieldLabel: {
            basis: { ja: "固定", en: "Fixed" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        checkbox: {
            alignToPixelGrid: { ja: "ピクセルグリッドに最適化", en: "Make Pixel Perfect" },
            addArtboard: { ja: "アートボードを追加", en: "Add Artboard" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." }
        },
        tooltip: {
            anchor: {
                ja: "大きさを変えても動かない点です。クリックで選びます。",
                en: "The point that stays put when the size changes. Click to choose."
            },
            ratioOriginal: {
                ja: "各オブジェクトの元の縦横比を保ちます。向きの指定は使いません。",
                en: "Keeps each object's original ratio. The orientation setting is ignored."
            },
            ratioPreset: { ja: "よく使う比率です。", en: "Common ratios." },
            ratioCustom: { ja: "下の欄に任意の比率（横:縦）を入力します。", en: "Enter any ratio (width:height) in the fields below." },
            customWidth: { ja: "カスタム比の横の値です。", en: "The width part of the custom ratio." },
            customHeight: { ja: "カスタム比の縦の値です。", en: "The height part of the custom ratio." },
            resultRatio: {
                ja: "いまの結果の縦横比です。編集するには［カスタム］を選びます。",
                en: "The current resulting ratio. Choose Custom to edit it."
            },
            landscape: {
                ja: "横（ランドスケープ）：長い辺を横にします（1:1・［元の比率］・［固定：なし］では変わりません）。",
                en: "Landscape: puts the longer side horizontally (no effect on 1:1, Original Ratio, or Fixed: None)."
            },
            portrait: {
                ja: "縦（ポートレート）：長い辺を縦にします（1:1・［元の比率］・［固定：なし］では変わりません）。",
                en: "Portrait: puts the longer side vertically (no effect on 1:1, Original Ratio, or Fixed: None)."
            },
            basisNone: {
                ja: "比率を使わず、幅・高さをそれぞれ入力します（入力しない辺は元のまま）。",
                en: "Ignores the ratio; enter width and height separately (untouched sides stay as they are)."
            },
            basisHorizontal: {
                ja: "幅を元のまま固定し、高さを比率か入力した値で決めます。",
                en: "Keeps the width and sets the height from the ratio or the value you enter."
            },
            basisVertical: {
                ja: "高さを元のまま固定し、幅を比率か入力した値で決めます。",
                en: "Keeps the height and sets the width from the ratio or the value you enter."
            },
            sizeValue: {
                ja: "この辺の長さです。入力すると比率より優先し、結果の比率をカスタム欄に表示します。",
                en: "Length of this side. Entering it overrides the ratio; the result shows in the custom fields."
            },
            sizePercent: {
                ja: "この辺を、元の長さに対する％で指定します。入力すると比率より優先します。",
                en: "Sets this side as a percentage of its original length. Entering it overrides the ratio."
            },
            fixedValue: {
                ja: "固定している辺です。変えるには［固定］を切り替えます。",
                en: "This side is fixed. Change Fixed to edit it."
            },
            alignToPixelGrid: {
                ja: "確定時に［ピクセルグリッドに最適化］を実行し、結果をピクセルグリッドに合わせます。",
                en: "On OK, runs Make Pixel Perfect to align the result to the pixel grid."
            },
            addArtboard: {
                ja: "結果と同じ範囲にアートボードを追加します。オブジェクトは残ります。",
                en: "Adds an artboard matching the result. The objects stay in place."
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            reset: {
                ja: "［元の比率］を選び、ダイアログを開く前の大きさに戻します。",
                en: "Selects Original Ratio and restores the size from before the dialog opened."
            }
        },
        button: {
            reset: { ja: "リセット", en: "Reset" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを追加する
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 表示名
     * @param {Object} tipSet - ツールチップ
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parent, labelSet, tipSet) {
        var radioButton = parent.add("radiobutton", undefined, getLabel(labelSet));
        radioButton.helpTip = getLabel(tipSet);
        return radioButton;
    }

    /**
     * チェックボックスを追加する
     * @param {Object} parent - 追加先
     * @param {Object} labelSet - 表示名
     * @param {Object} tipSet - ツールチップ
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parent, labelSet, tipSet) {
        var checkbox = parent.add("checkbox", undefined, getLabel(labelSet));
        checkbox.helpTip = getLabel(tipSet);
        return checkbox;
    }

    /**
     * 数値入力欄を追加する
     * @param {Group} parent - 追加先
     * @param {string} initialText - 初期値
     * @param {Object} tipSet - ツールチップ
     * @param {number} [fieldCharacters] - 欄の幅（文字数）。省略時は FIELD_CHARACTERS
     * @param {string} [unit] - 欄の中に数値の後ろに出す単位（例 " mm"、"%"）。省略時は単位なし
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup、∧∨と欄を束ねた group は .parent）
     */
    function addNumberField(parent, initialText, tipSet, fieldCharacters, unit) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = parent.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var numberField;
        /* 負数は不可。増減後は onChanging を呼んでプレビューを更新する / no negatives; fire onChanging to refresh the preview */
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberField; }, {
            min: 0,
            unit: unit,
            onStep: function (steppedField) {
                if (typeof steppedField.onChanging === "function") steppedField.onChanging();
            }
        });
        numberField = stepperFieldGroup.add("edittext", undefined, initialText);
        numberField.helpTip = getLabel(tipSet);
        numberField.characters = fieldCharacters || FIELD_CHARACTERS;
        numberField.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberField, stepperGroup);
        return numberField;
    }

    /**
     * 入力欄と計算値の表示を同じ位置に重ねて追加する（setFieldValueEditable() で出し分ける。
     * 無効にした入力欄は Mac で文字が薄く読めないため）
     * @param {Group} parent - 追加先
     * @param {Object} fieldTipSet - 入力欄のツールチップ
     * @param {Object} valueTipSet - 計算値のツールチップ
     * @param {number} fieldCharacters - 欄の幅（文字数）
     * @param {string} [initialText] - 入力欄の初期値
     * @param {string} [unit] - 欄の中に出す単位（例 " mm"、"%"）。省略時は単位なし
     * @returns {{field: EditText, valueText: StaticText, unit: string}} 入力欄と計算値の表示、単位
     */
    function addFieldValueStack(parent, fieldTipSet, valueTipSet, fieldCharacters, initialText, unit) {
        var valueStack = parent.add("group");
        valueStack.orientation = "stack";
        valueStack.alignChildren = ["fill", "center"];
        var inputField = addNumberField(valueStack, initialText || "", fieldTipSet, fieldCharacters, unit);
        /* 計算値は∧∨の幅だけ右へずらし、入力欄と同じ位置に出す / indent by the stepper width to line up with the field */
        var valueTextGroup = valueStack.add("group");
        valueTextGroup.margins = [STEPPER_SIDE_MARGIN + STEPPER_BUTTON_WIDTH, 0, 0, 0];
        valueTextGroup.alignChildren = ["fill", "center"];
        var valueText = valueTextGroup.add("statictext", undefined, "");
        valueText.characters = fieldCharacters;
        valueText.helpTip = getLabel(valueTipSet);
        return { field: inputField, valueText: valueText, unit: unit || "" };
    }

    /**
     * 「項目名：［長さ mm］ ［％%］」の行を追加する（単位は欄の中に出す）（固定した辺は読めるだけの表示に切り替える）
     * @param {Object} parent - 追加先
     * @param {Object} labelSet - 項目名
     * @param {number} labelWidth - 項目名の幅（右揃えでそろえる）
     * @returns {{length: Object, percent: Object}} 長さと％の addFieldValueStack() の戻り値
     */
    function addSizeRow(parent, labelSet, labelWidth) {
        var sizeRow = addRowGroup(parent);
        var rowLabel = sizeRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = labelWidth;
        rowLabel.justify = "right";

        var lengthStack = addFieldValueStack(sizeRow, LABELS.tooltip.sizeValue, LABELS.tooltip.fixedValue, FIELD_CHARACTERS, "", " " + RULER_UNIT.label);
        var percentStack = addFieldValueStack(sizeRow, LABELS.tooltip.sizePercent, LABELS.tooltip.fixedValue, PERCENT_CHARS, "", "%");
        return { length: lengthStack, percent: percentStack };
    }

    /**
     * 行の長さと％を、入力欄（固定しない辺）か読めるだけの表示（固定した辺）に切り替える
     * @param {Object} sizeRow - addSizeRow() の戻り値
     * @param {boolean} editable - 入力欄を出すなら true
     * @returns {void}
     */
    function setSizeRowEditable(sizeRow, editable) {
        setFieldValueEditable(sizeRow.length, editable);
        setFieldValueEditable(sizeRow.percent, editable);
    }

    /**
     * 入力欄と計算値の表示を切り替える
     * @param {Object} fieldValue - addFieldValueStack() の戻り値
     * @param {boolean} editable - 入力欄を出すなら true
     * @returns {void}
     */
    function setFieldValueEditable(fieldValue, editable) {
        /* ∧∨ごと出し分ける / show or hide together with the stepper */
        fieldValue.field.parent.visible = editable;
        fieldValue.valueText.parent.visible = !editable;
    }

    /**
     * 「縦横比」パネルを作る（プリセットのラジオとカスタム比の欄）
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildRatioPanel(parent, dialogControls) {
        var ratioPanel = addPanel(parent, getLabel(LABELS.panel.aspectRatio));
        var ratioRadioGroup = addColumnGroup(ratioPanel);
        dialogControls.ratioOriginalRadio = addRadio(ratioRadioGroup, LABELS.radio.ratioOriginal, LABELS.tooltip.ratioOriginal);
        dialogControls.ratio16x9Radio = addRadio(ratioRadioGroup, LABELS.radio.ratio16x9, LABELS.tooltip.ratioPreset);
        dialogControls.ratioSquareRadio = addRadio(ratioRadioGroup, LABELS.radio.ratioSquare, LABELS.tooltip.ratioPreset);
        dialogControls.ratioA4Radio = addRadio(ratioRadioGroup, LABELS.radio.ratioA4, LABELS.tooltip.ratioPreset);
        dialogControls.ratioCustomRadio = addRadio(ratioRadioGroup, LABELS.radio.ratioCustom, LABELS.tooltip.ratioCustom);

        var customRatioGroup = addRowGroup(ratioPanel);
        customRatioGroup.margins = [CUSTOM_RATIO_INDENT, 0, 0, 0];
        /* ［カスタム］以外のときは、結果の縦横比を読める文字で出す / Outside Custom, show the resulting ratio as readable text */
        dialogControls.customWidthStack = addFieldValueStack(customRatioGroup, LABELS.tooltip.customWidth, LABELS.tooltip.resultRatio, CUSTOM_RATIO_CHARS, DEFAULT_CUSTOM_RATIO_WIDTH);
        customRatioGroup.add("statictext", undefined, ":");
        dialogControls.customHeightStack = addFieldValueStack(customRatioGroup, LABELS.tooltip.customHeight, LABELS.tooltip.resultRatio, CUSTOM_RATIO_CHARS, DEFAULT_CUSTOM_RATIO_HEIGHT);
        dialogControls.customWidthField = dialogControls.customWidthStack.field;
        dialogControls.customHeightField = dialogControls.customHeightStack.field;
    }

    /**
     * 「オプション」パネルを作る
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildOptionPanel(parent, dialogControls) {
        var optionPanel = addPanel(parent, getLabel(LABELS.panel.options), 6);
        dialogControls.alignToPixelCheckbox = addCheckbox(optionPanel, LABELS.checkbox.alignToPixelGrid, LABELS.tooltip.alignToPixelGrid);
        dialogControls.addArtboardCheckbox = addCheckbox(optionPanel, LABELS.checkbox.addArtboard, LABELS.tooltip.addArtboard);
    }

    /**
     * 「向きと基準点」パネルを作る（左に向きアイコン、右に9軸）
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildOrientationAnchorPanel(parent, dialogControls) {
        var orientationAnchorPanel = addPanel(parent, getLabel(LABELS.panel.orientationAnchor));
        var orientationAnchorRow = addRowGroup(orientationAnchorPanel);
        /* 左右に離さず、パネルの中央に寄せて並べる / Keep them together, centered in the panel */
        orientationAnchorRow.alignment = ["center", "top"];
        orientationAnchorRow.spacing = ORIENT_ANCHOR_GAP;
        dialogControls.orientationGroup = addOrientationButtons(orientationAnchorRow, LABELS.tooltip.portrait, LABELS.tooltip.landscape);
        /* addRowGroup() の top を上書きして、9軸と天地中央をそろえる / Override the top alignment to center vertically with the 9-axis widget */
        dialogControls.orientationGroup.alignment = ["left", "center"];
        /* クリックで widget.onAnchorChange() を呼ぶ（bindSettingsHandlers で設定）/ A click calls widget.onAnchorChange(), set in bindSettingsHandlers */
        dialogControls.anchorWidget = addAnchorWidget(orientationAnchorRow, DEFAULT_ANCHOR_INDEX, function (anchorIndex, anchorWidget) {
            if (typeof anchorWidget.onAnchorChange === "function") anchorWidget.onAnchorChange();
        });
        dialogControls.anchorWidget.helpTip = getLabel(LABELS.tooltip.anchor);
    }

    /**
     * 「サイズ」パネルを作る（固定する辺のラジオと、幅・高さの行）
     * @param {Group} parent - 追加先のカラム
     * @param {Object} dialogControls - コントロールの参照を書き込む先
     * @returns {void}
     */
    function buildSizePanel(parent, dialogControls) {
        var sizePanel = addPanel(parent, getLabel(LABELS.panel.size));

        /* 項目名の右にラジオを横に並べる（排他にするため同じ親へ）/ Radios in a row right of the label, in one parent to stay exclusive */
        var basisRow = addRowGroup(sizePanel);
        basisRow.add("statictext", undefined, labelText(LABELS.fieldLabel.basis));
        var basisRadioGroup = addRowGroup(basisRow);
        dialogControls.basisNoneRadio = addRadio(basisRadioGroup, LABELS.radio.basisNone, LABELS.tooltip.basisNone);
        dialogControls.basisHorizontalRadio = addRadio(basisRadioGroup, LABELS.radio.basisHorizontal, LABELS.tooltip.basisHorizontal);
        dialogControls.basisVerticalRadio = addRadio(basisRadioGroup, LABELS.radio.basisVertical, LABELS.tooltip.basisVertical);

        /* 幅・高さの両方を表示し、固定しない辺だけ編集できる / Show both; only the unfixed side is editable */
        var sizeLabelWidth = Math.max(
            sizePanel.graphics.measureString(labelText(LABELS.fieldLabel.width))[0],
            sizePanel.graphics.measureString(labelText(LABELS.fieldLabel.height))[0]
        );
        dialogControls.widthRow = addSizeRow(sizePanel, LABELS.fieldLabel.width, sizeLabelWidth);
        dialogControls.heightRow = addSizeRow(sizePanel, LABELS.fieldLabel.height, sizeLabelWidth);
    }

    /**
     * ダイアログを作成する（左：縦横比・オプション / 右：向きと基準点・サイズ）
     * @returns {Object} ダイアログと各コントロールの参照
     */
    function createDialog() {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = addColumnGroup(columnsGroup);
        var rightColumn = addColumnGroup(columnsGroup);
        leftColumn.alignChildren = ["fill", "top"];
        rightColumn.alignChildren = ["fill", "top"];

        var dialogControls = { dialog: dialog };
        buildRatioPanel(leftColumn, dialogControls);
        buildOptionPanel(leftColumn, dialogControls);
        buildOrientationAnchorPanel(rightColumn, dialogControls);
        buildSizePanel(rightColumn, dialogControls);

        /* ボタン行（左：リセット / 右：キャンセル・OK） / Button row (left: Reset, right: Cancel and OK) */
        var buttonRow = addButtonRow(dialog);
        var btnReset = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);
        dialogControls.btnReset = btnReset;
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        return dialogControls;
    }

    /**
     * ダイアログから選択中の比率（横 ÷ 縦）を読む
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {number|null} 比率。［元の比率］なら null、カスタムが数値でないか 0 以下なら 1
     */
    function readRatio(dialogControls) {
        if (dialogControls.ratioOriginalRadio.value) return null;
        if (dialogControls.ratio16x9Radio.value) return RATIO_16_9;
        if (dialogControls.ratioSquareRadio.value) return RATIO_SQUARE;
        if (dialogControls.ratioA4Radio.value) return RATIO_A4;
        var ratioWidth = parseFloat(dialogControls.customWidthField.text);
        var ratioHeight = parseFloat(dialogControls.customHeightField.text);
        if (isNaN(ratioWidth) || isNaN(ratioHeight) || ratioWidth <= 0 || ratioHeight <= 0) return 1;
        return ratioWidth / ratioHeight;
    }

    /**
     * 固定する辺の行を返す
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {Object|null} addSizeRow() の戻り値。［なし］なら null
     */
    function getFixedRow(dialogControls) {
        if (dialogControls.basisHorizontalRadio.value) return dialogControls.widthRow;
        if (dialogControls.basisVerticalRadio.value) return dialogControls.heightRow;
        return null;
    }

    /**
     * 欄の正の数値を読む
     * @param {EditText} numberField - 入力欄
     * @returns {number|null} 値。空欄・不正値・0以下なら null
     */
    function readPositiveNumber(numberField) {
        var fieldValue = parseFloat(numberField.text);
        return (isNaN(fieldValue) || fieldValue <= 0) ? null : fieldValue;
    }

    /**
     * 行を入力できるか（固定しない辺。［なし］なら幅・高さとも）
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} sizeRow - addSizeRow() の戻り値
     * @returns {boolean} 入力できるなら true
     */
    function isSizeRowEditable(dialogControls, sizeRow) {
        var fixedRow = getFixedRow(dialogControls);
        return fixedRow === null || fixedRow !== sizeRow;
    }

    /**
     * 行に入力された長さか％を読む（入力できない行や、まだ入力していない行は両方 null）
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} sizeRow - addSizeRow() の戻り値（inputMode に "length" / "percent" / null）
     * @returns {{lengthPt: (number|null), percent: (number|null)}} 長さ（pt）と％
     */
    function readSizeRowInput(dialogControls, sizeRow) {
        var rowInput = { lengthPt: null, percent: null };
        if (!isSizeRowEditable(dialogControls, sizeRow)) return rowInput;
        if (sizeRow.inputMode === "length") {
            var lengthValue = readPositiveNumber(sizeRow.length.field);
            if (lengthValue !== null) rowInput.lengthPt = lengthValue * RULER_UNIT.pointsPerUnit;
        } else if (sizeRow.inputMode === "percent") {
            rowInput.percent = readPositiveNumber(sizeRow.percent.field);
        }
        return rowInput;
    }

    /**
     * ダイアログの設定をまとめて読む。入力した長さか％は比率より優先する
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {{ratio: (number|null), wantPortrait: boolean, keepSize: boolean, fixByHeight: boolean, anchorIndex: number, widthInput: Object, heightInput: Object}} 設定
     */
    function readScaleSettings(dialogControls) {
        var fixedRow = getFixedRow(dialogControls);
        return {
            ratio: readRatio(dialogControls),
            wantPortrait: dialogControls.orientationGroup.portraitButton.isSelected,
            keepSize: (fixedRow === null),
            fixByHeight: (fixedRow === dialogControls.heightRow),
            anchorIndex: getAnchorWidgetIndex(dialogControls.anchorWidget),
            widthInput: readSizeRowInput(dialogControls, dialogControls.widthRow),
            heightInput: readSizeRowInput(dialogControls, dialogControls.heightRow)
        };
    }

    // =========================================
    // 計算とプレビュー / Calculation and preview
    // =========================================

    /**
     * 比率を向きに合わせて反転する
     * @param {number} ratio - 比率（横 ÷ 縦）
     * @param {boolean} wantPortrait - 縦置きなら true
     * @returns {number} 向きをそろえた比率
     */
    function orientRatio(ratio, wantPortrait) {
        if (wantPortrait && ratio > 1) return 1 / ratio;
        if (!wantPortrait && ratio < 1) return 1 / ratio;
        return ratio;
    }

    /**
     * 固定する辺の長さから、比率に合う幅と高さを求める
     * @param {number} orientedRatio - 向きをそろえた比率
     * @param {boolean} fixByHeight - 高さを固定するなら true
     * @param {number} baseSizePt - 固定する辺の長さ（pt）
     * @returns {{width: number, height: number}} 幅と高さ（pt）
     */
    function computeTargetSize(orientedRatio, fixByHeight, baseSizePt) {
        /* 比率から求める側だけを単位に合わせて丸める / Round only the side derived from the ratio */
        if (fixByHeight) {
            return { width: roundForUnit(baseSizePt * orientedRatio), height: baseSizePt };
        }
        return { width: baseSizePt, height: roundForUnit(baseSizePt / orientedRatio) };
    }

    /**
     * 入力した長さか％から辺の長さを求める
     * @param {Object} rowInput - readSizeRowInput() の戻り値
     * @param {number} originalLength - 元の長さ（pt）
     * @returns {number|null} 長さ（pt）。入力がなければ null
     */
    function resolveInputLength(rowInput, originalLength) {
        if (rowInput.lengthPt !== null) return rowInput.lengthPt;
        if (rowInput.percent !== null) return originalLength * rowInput.percent / 100;
        return null;
    }

    /**
     * プレビューの i 番目のアイテムの仕上がり寸法を求める
     * ［なし］は幅・高さを入力どおり（入力がなければ元のまま）。
     * 固定したときは、固定する辺は元のまま、もう一方は入力があればそれ、なければ比率から求める
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {number} itemIndex - アイテムの番号
     * @param {Object} scaleSettings - readScaleSettings() の戻り値
     * @returns {{width: number, height: number}} 幅と高さ（pt）
     */
    function computeItemTargetSize(preview, itemIndex, scaleSettings) {
        var originalWidth = preview.originalWidths[itemIndex];
        var originalHeight = preview.originalHeights[itemIndex];
        var inputWidth = resolveInputLength(scaleSettings.widthInput, originalWidth);
        var inputHeight = resolveInputLength(scaleSettings.heightInput, originalHeight);
        if (scaleSettings.keepSize) {
            return {
                width: (inputWidth === null) ? originalWidth : inputWidth,
                height: (inputHeight === null) ? originalHeight : inputHeight
            };
        }
        if (scaleSettings.fixByHeight && inputWidth !== null) return { width: inputWidth, height: originalHeight };
        if (!scaleSettings.fixByHeight && inputHeight !== null) return { width: originalWidth, height: inputHeight };

        /* ［元の比率］はアイテムごとの元の比率（向きで反転しない）/ Original: each item's own ratio, not flipped */
        var itemRatio;
        if (scaleSettings.ratio === null) {
            itemRatio = (originalWidth > 0 && originalHeight > 0) ? originalWidth / originalHeight : 1;
        } else {
            itemRatio = orientRatio(scaleSettings.ratio, scaleSettings.wantPortrait);
        }
        return computeTargetSize(itemRatio, scaleSettings.fixByHeight, scaleSettings.fixByHeight ? originalHeight : originalWidth);
    }

    /**
     * プレビューのアイテムに比率を当てる
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {Object} scaleSettings - readScaleSettings() の戻り値
     * @returns {void}
     */
    function applyRatioToPreview(preview, scaleSettings) {
        for (var i = 0; i < preview.items.length; i++) {
            var targetSize = computeItemTargetSize(preview, i, scaleSettings);
            var previewItem = preview.items[i];
            previewItem.width = targetSize.width;
            previewItem.height = targetSize.height;

            /* 基準点が元の位置に戻るように移動 / Move so the reference point returns to where it was */
            var anchorBefore = getAnchorPointOnBounds(preview.originalBounds[i], scaleSettings.anchorIndex);
            var anchorAfter = getAnchorPointOnBounds(previewItem.geometricBounds, scaleSettings.anchorIndex);
            previewItem.translate(anchorBefore[0] - anchorAfter[0], anchorBefore[1] - anchorAfter[1]);
        }
        app.redraw();
    }

    /**
     * 選択アイテムの複製をプレビュー用に作り、元は隠す
     * @param {Object[]} selectedItems - 選択アイテム
     * @returns {Object} { items, originalWidths, originalHeights, originalBounds }
     */
    function createPreviewFromSelection(selectedItems) {
        var preview = { items: [], originalWidths: [], originalHeights: [], originalBounds: [] };
        for (var i = 0; i < selectedItems.length; i++) {
            var previewCopy = selectedItems[i].duplicate();
            previewCopy.zOrder(ZOrderMethod.BRINGTOFRONT);
            preview.items.push(previewCopy);
            preview.originalWidths.push(selectedItems[i].width);
            preview.originalHeights.push(selectedItems[i].height);
            preview.originalBounds.push(selectedItems[i].geometricBounds);
            selectedItems[i].hidden = true;
        }
        return preview;
    }

    /**
     * 選択がないとき、アクティブなアートボードの中央に長方形を作ってプレビューにする
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} scaleSettings - readScaleSettings() の戻り値
     * @returns {Object} { items, originalWidths, originalHeights, originalBounds }
     */
    function createPreviewRectangle(doc, scaleSettings) {
        var defaultWidth = DEFAULT_RECT_WIDTH_BY_UNIT[RULER_UNIT.code];
        var rectWidthPt = defaultWidth ? defaultWidth * RULER_UNIT.pointsPerUnit : FALLBACK_RECT_WIDTH_PT;
        /* ［元の比率］は元が無いので 16:9 で作る / Original has no source here, so draw at 16:9 */
        var rectRatio = (scaleSettings.ratio === null) ? RATIO_16_9 : scaleSettings.ratio;
        var targetSize = computeTargetSize(orientRatio(rectRatio, scaleSettings.wantPortrait), false, rectWidthPt);

        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; /* [L,T,R,B] */
        var centerX = (artboardRect[0] + artboardRect[2]) / 2;
        var centerY = (artboardRect[1] + artboardRect[3]) / 2;
        var previewRect = doc.pathItems.rectangle(centerY + targetSize.height / 2, centerX - targetSize.width / 2, targetSize.width, targetSize.height);
        previewRect.stroked = false;
        previewRect.filled = true;

        return {
            items: [previewRect],
            originalWidths: [targetSize.width],
            originalHeights: [targetSize.height],
            originalBounds: [previewRect.geometricBounds]
        };
    }

    // =========================================
    // 確定と取り消し / Commit and cancel
    // =========================================

    /**
     * 確定後の仕上げ（ピクセル最適化・アートボード追加）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} resultItem - 仕上げるアイテム
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {void}
     */
    function applyFinishingOptions(doc, resultItem, dialogControls) {
        if (dialogControls.alignToPixelCheckbox.value) {
            doc.selection = [resultItem];
            app.executeMenuCommand("Make Pixel Perfect");
        }
        if (dialogControls.addArtboardCheckbox.value) {
            doc.artboards.add(resultItem.visibleBounds);
        }
    }

    /**
     * プレビューの大きさと位置を元のアイテムに移し、プレビューを消す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} selectedItems - 元の選択アイテム
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {void}
     */
    function commitToOriginals(doc, selectedItems, preview, dialogControls) {
        for (var i = 0; i < selectedItems.length; i++) {
            var originalItem = selectedItems[i];
            var previewCopy = preview.items[i];
            originalItem.hidden = false;

            /* 中心基準で拡大縮小し、左上をそろえる / Scale from center, then align the top-left */
            var originalWidth = preview.originalWidths[i];
            var originalHeight = preview.originalHeights[i];
            if (originalWidth > 0 && originalHeight > 0) {
                originalItem.resize(previewCopy.width / originalWidth * 100, previewCopy.height / originalHeight * 100);
            }
            originalItem.position = previewCopy.position;
            previewCopy.remove();

            applyFinishingOptions(doc, originalItem, dialogControls);
        }
        doc.selection = selectedItems;
    }

    /**
     * プレビューを消し、隠した元のアイテムを戻す
     * @param {Object[]} selectedItems - 元の選択アイテム
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @returns {void}
     */
    function cancelPreview(selectedItems, preview) {
        for (var i = 0; i < preview.items.length; i++) {
            preview.items[i].remove();
        }
        for (var j = 0; j < selectedItems.length; j++) {
            selectedItems[j].hidden = false;
        }
    }

    /**
     * 数値を小数第2位までの文字列にする
     * @param {number} value - 値
     * @returns {string} 文字列
     */
    function formatSizeNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    /**
     * プレビューの長さの表示用文字列を返す（対象が複数なら空欄）
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {boolean} isWidth - 幅なら true
     * @returns {string} 長さ（定規の単位）
     */
    function getPreviewLengthText(preview, isWidth) {
        if (preview.items.length !== 1) return "";
        var lengthPt = isWidth ? preview.items[0].width : preview.items[0].height;
        return formatSizeNumber(lengthPt / RULER_UNIT.pointsPerUnit);
    }

    /**
     * プレビューの元の長さに対する％の表示用文字列を返す（複数で値がそろわなければ空欄）
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @param {boolean} isWidth - 幅なら true
     * @returns {string} ％
     */
    function getPreviewPercentText(preview, isWidth) {
        var percentText = "";
        for (var i = 0; i < preview.items.length; i++) {
            var originalLength = isWidth ? preview.originalWidths[i] : preview.originalHeights[i];
            if (!(originalLength > 0)) return "";
            var currentLength = isWidth ? preview.items[i].width : preview.items[i].height;
            var itemPercentText = formatSizeNumber(currentLength / originalLength * 100);
            if (i > 0 && itemPercentText !== percentText) return "";
            percentText = itemPercentText;
        }
        return percentText;
    }

    /**
     * 入力欄と表示の両方に同じ文字列を入れる（空欄でなければ欄の単位を付ける）
     * @param {Object} fieldValue - addFieldValueStack() の戻り値
     * @param {string} valueText - 入れる文字列（数値のみ）
     * @returns {void}
     */
    function writeFieldValue(fieldValue, valueText) {
        if (valueText !== "") valueText += fieldValue.unit;
        fieldValue.field.text = valueText;
        fieldValue.valueText.text = valueText;
    }

    /**
     * 幅・高さの行に、プレビューの長さと％を表示する
     * 入力できる行で入力中のほう（長さか％）は書き換えず、もう一方だけをそろえる
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @returns {void}
     */
    function showSizeValues(dialogControls, preview) {
        var sizeSides = [
            { row: dialogControls.widthRow, isWidth: true },
            { row: dialogControls.heightRow, isWidth: false }
        ];
        for (var i = 0; i < sizeSides.length; i++) {
            var sizeRow = sizeSides[i].row;
            var editingInput = isSizeRowEditable(dialogControls, sizeRow) ? sizeRow.inputMode : null;
            if (editingInput !== "length") writeFieldValue(sizeRow.length, getPreviewLengthText(preview, sizeSides[i].isWidth));
            if (editingInput !== "percent") writeFieldValue(sizeRow.percent, getPreviewPercentText(preview, sizeSides[i].isWidth));
        }
    }

    /**
     * 比率を「横:縦」の数値の組にする（分母16までの整数比に近ければ整数、そうでなければ短い辺を1にした小数）
     * @param {number} ratio - 比率（横 ÷ 縦）
     * @returns {string[]} [横, 縦]
     */
    function toRatioPair(ratio) {
        for (var denominator = 1; denominator <= RATIO_PAIR_MAX_DENOMINATOR; denominator++) {
            var numerator = ratio * denominator;
            var roundedNumerator = Math.round(numerator);
            if (roundedNumerator > 0 && Math.abs(numerator - roundedNumerator) / numerator < RATIO_PAIR_TOLERANCE) {
                return [String(roundedNumerator), String(denominator)];
            }
        }
        return (ratio >= 1) ? [formatSizeNumber(ratio), "1"] : ["1", formatSizeNumber(1 / ratio)];
    }

    /**
     * プレビューの縦横比をカスタム比の欄に表示する（複数で比率がそろわなければ空欄）
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @returns {void}
     */
    function showResultRatio(dialogControls, preview) {
        var ratioPair = null;
        for (var i = 0; i < preview.items.length; i++) {
            var itemHeight = preview.items[i].height;
            var itemPair = (itemHeight > 0) ? toRatioPair(preview.items[i].width / itemHeight) : ["", ""];
            if (ratioPair && (itemPair[0] !== ratioPair[0] || itemPair[1] !== ratioPair[1])) {
                ratioPair = ["", ""];
                break;
            }
            ratioPair = itemPair;
        }
        var ratioStacks = [dialogControls.customWidthStack, dialogControls.customHeightStack];
        for (var j = 0; j < ratioStacks.length; j++) {
            /* 表示と、［カスタム］に切り替えたときの初期値の両方に使う / Used for display and as the starting value for Custom */
            writeFieldValue(ratioStacks[j], ratioPair[j]);
        }
    }

    /**
     * ダイアログの状態をそろえてプレビューを更新する
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} preview - { items, originalWidths, originalHeights, originalBounds }
     * @returns {void}
     */
    function refreshPreview(dialogControls, preview) {
        var isCustomRatio = dialogControls.ratioCustomRadio.value;
        setFieldValueEditable(dialogControls.customWidthStack, isCustomRatio);
        setFieldValueEditable(dialogControls.customHeightStack, isCustomRatio);
        setSizeRowEditable(dialogControls.widthRow, isSizeRowEditable(dialogControls, dialogControls.widthRow));
        setSizeRowEditable(dialogControls.heightRow, isSizeRowEditable(dialogControls, dialogControls.heightRow));
        var scaleSettings = readScaleSettings(dialogControls);
        applyRatioToPreview(preview, scaleSettings);
        showSizeValues(dialogControls, preview);
        /* ［カスタム］の欄は比率の入力元なので、比率を使わなかったときだけ結果を書く / Custom fields feed the ratio; overwrite them only when the ratio was not used */
        var hasSizeInput = scaleSettings.widthInput.lengthPt !== null || scaleSettings.widthInput.percent !== null
            || scaleSettings.heightInput.lengthPt !== null || scaleSettings.heightInput.percent !== null;
        if (!isCustomRatio || scaleSettings.keepSize || hasSizeInput) showResultRatio(dialogControls, preview);
    }

    /**
     * ラジオの組に、押したもの以外を外してから onSelect を呼ぶハンドラーを付ける
     * （クリック直後は前の選択が残って読めることがあるため、自分で排他にする）
     * @param {RadioButton[]} radioSet - 排他にするラジオ
     * @param {Function} onSelect - 選び直したときに呼ぶ関数
     * @returns {void}
     */
    function bindExclusiveRadios(radioSet, onSelect) {
        for (var i = 0; i < radioSet.length; i++) {
            radioSet[i].onClick = function () {
                for (var k = 0; k < radioSet.length; k++) radioSet[k].value = (radioSet[k] === this);
                onSelect();
            };
        }
    }

    /**
     * 設定を変えるたびにプレビューを更新するよう、各コントロールにハンドラーを付ける
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Function} onSettingsChange - 呼び出す関数
     * @returns {void}
     */
    function bindSettingsHandlers(dialogControls, onSettingsChange) {
        var sizeRows = [dialogControls.widthRow, dialogControls.heightRow];

        /* 固定する辺を変えたら入力をやめて元に戻す / Changing the fixed side drops the typed sizes */
        var onBasisChange = function () {
            clearSizeInputs(dialogControls);
            onSettingsChange();
        };
        /* 比率・向きを変えたら比率に戻す（［なし］は比率を使わないので入力を残す）/ Changing ratio or orientation returns to the ratio (None keeps the typed sizes) */
        var onRatioChange = function () {
            if (getFixedRow(dialogControls) !== null) clearSizeInputs(dialogControls);
            onSettingsChange();
        };
        bindExclusiveRadios([dialogControls.ratioOriginalRadio, dialogControls.ratio16x9Radio, dialogControls.ratioSquareRadio, dialogControls.ratioA4Radio, dialogControls.ratioCustomRadio], onRatioChange);
        bindExclusiveRadios([dialogControls.basisNoneRadio, dialogControls.basisHorizontalRadio, dialogControls.basisVerticalRadio], onBasisChange);
        dialogControls.customWidthField.onChanging = onRatioChange;
        dialogControls.customHeightField.onChanging = onRatioChange;
        dialogControls.orientationGroup.onOrientationChange = onRatioChange;
        dialogControls.anchorWidget.onAnchorChange = onSettingsChange;

        /* 行ごとに、最後に編集したほう（長さか％）で決める / Each row follows whichever field was edited last */
        for (var i = 0; i < sizeRows.length; i++) {
            bindSizeRowInput(sizeRows[i], onSettingsChange);
        }
    }

    /**
     * 行の長さ・％の欄に、編集したほうを inputMode に控えるハンドラーを付ける
     * @param {Object} sizeRow - addSizeRow() の戻り値
     * @param {Function} onSettingsChange - 呼び出す関数
     * @returns {void}
     */
    function bindSizeRowInput(sizeRow, onSettingsChange) {
        sizeRow.length.field.onChanging = function () {
            sizeRow.inputMode = "length";
            onSettingsChange();
        };
        sizeRow.percent.field.onChanging = function () {
            sizeRow.inputMode = "percent";
            onSettingsChange();
        };
    }

    /**
     * 幅・高さの入力をやめる（元の長さか比率に戻る）
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {void}
     */
    function clearSizeInputs(dialogControls) {
        dialogControls.widthRow.inputMode = null;
        dialogControls.heightRow.inputMode = null;
    }

    /* #targetengine 下の $.global は Illustrator の起動中は残るので、OK したときの設定をここに控える
       $.global survives while Illustrator runs under #targetengine; the settings at OK live here */
    var SAVED_STATE_KEY = "__AspectRatioScaler_State";

    /**
     * 初期値の設定を返す（開いたとき前回の値が無ければ、および［リセット］で使う）
     * @returns {Object} 設定（captureDialogState() と同じ形）
     */
    function getDefaultDialogState() {
        return {
            ratio: "original",
            customWidth: DEFAULT_CUSTOM_RATIO_WIDTH,
            customHeight: DEFAULT_CUSTOM_RATIO_HEIGHT,
            alignToPixel: DEFAULT_ALIGN_TO_PIXEL_GRID,
            addArtboard: false,
            portrait: false,
            anchorIndex: DEFAULT_ANCHOR_INDEX,
            basis: "vertical"
        };
    }

    /**
     * 比率のラジオと、控えるときの名前の対応を返す
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {Object} 名前 → ラジオ
     */
    function getRatioRadioMap(dialogControls) {
        return {
            original: dialogControls.ratioOriginalRadio,
            "16x9": dialogControls.ratio16x9Radio,
            square: dialogControls.ratioSquareRadio,
            a4: dialogControls.ratioA4Radio,
            custom: dialogControls.ratioCustomRadio
        };
    }

    /**
     * 固定する辺のラジオと、控えるときの名前の対応を返す
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {Object} 名前 → ラジオ
     */
    function getBasisRadioMap(dialogControls) {
        return {
            none: dialogControls.basisNoneRadio,
            horizontal: dialogControls.basisHorizontalRadio,
            vertical: dialogControls.basisVerticalRadio
        };
    }

    /**
     * ラジオの組のうち選ばれている名前を返す
     * @param {Object} radioMap - 名前 → ラジオ
     * @returns {string|null} 名前。どれも選ばれていなければ null
     */
    function getSelectedRadioName(radioMap) {
        for (var name in radioMap) {
            if (radioMap.hasOwnProperty(name) && radioMap[name].value) return name;
        }
        return null;
    }

    /**
     * ラジオの組を、名前のものだけ選んだ状態にする（知らない名前なら何もしない）
     * @param {Object} radioMap - 名前 → ラジオ
     * @param {string} selectedName - 選ぶラジオの名前
     * @returns {void}
     */
    function selectRadioByName(radioMap, selectedName) {
        if (!radioMap.hasOwnProperty(selectedName)) return;
        for (var name in radioMap) {
            if (radioMap.hasOwnProperty(name)) radioMap[name].value = (name === selectedName);
        }
    }

    /**
     * ダイアログの設定を控える形で読む（幅・高さの入力は対象ごとに違うので控えない）
     * @param {Object} dialogControls - createDialog() の戻り値
     * @returns {Object} 設定
     */
    function captureDialogState(dialogControls) {
        return {
            ratio: getSelectedRadioName(getRatioRadioMap(dialogControls)),
            customWidth: dialogControls.customWidthField.text,
            customHeight: dialogControls.customHeightField.text,
            alignToPixel: dialogControls.alignToPixelCheckbox.value,
            addArtboard: dialogControls.addArtboardCheckbox.value,
            portrait: dialogControls.orientationGroup.portraitButton.isSelected,
            anchorIndex: getAnchorWidgetIndex(dialogControls.anchorWidget),
            basis: getSelectedRadioName(getBasisRadioMap(dialogControls))
        };
    }

    /**
     * 各コントロールに設定を書き込み、幅・高さの入力をやめる
     * （幅・高さの欄はプレビューから書き直されるので触らない）
     * @param {Object} dialogControls - createDialog() の戻り値
     * @param {Object} dialogState - getDefaultDialogState() か captureDialogState() の戻り値
     * @returns {void}
     */
    function applyDialogState(dialogControls, dialogState) {
        selectRadioByName(getRatioRadioMap(dialogControls), dialogState.ratio);
        dialogControls.customWidthField.text = dialogState.customWidth;
        dialogControls.customHeightField.text = dialogState.customHeight;
        dialogControls.alignToPixelCheckbox.value = dialogState.alignToPixel;
        dialogControls.addArtboardCheckbox.value = dialogState.addArtboard;

        var orientationGroup = dialogControls.orientationGroup;
        orientationGroup.portraitButton.isSelected = dialogState.portrait;
        orientationGroup.landscapeButton.isSelected = !dialogState.portrait;
        redrawControl(orientationGroup.portraitButton);
        redrawControl(orientationGroup.landscapeButton);
        setAnchorWidgetValue(dialogControls.anchorWidget, dialogState.anchorIndex);

        selectRadioByName(getBasisRadioMap(dialogControls), dialogState.basis);
        clearSizeInputs(dialogControls);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        var hasSelection = (selectedItems && selectedItems.length > 0);
        if (!hasSelection) selectedItems = [];

        var dialogControls = createDialog();

        /* 前回 OK したときの設定で開く / Open with the settings from the last OK */
        applyDialogState(dialogControls, $.global[SAVED_STATE_KEY] || getDefaultDialogState());

        var preview = hasSelection
            ? createPreviewFromSelection(selectedItems)
            : createPreviewRectangle(doc, readScaleSettings(dialogControls));

        /* 設定が変わるたびにプレビューを更新 / Refresh the preview on every change */
        bindSettingsHandlers(dialogControls, function () { refreshPreview(dialogControls, preview); });
        refreshPreview(dialogControls, preview);

        dialogControls.btnReset.onClick = function () {
            applyDialogState(dialogControls, getDefaultDialogState());
            refreshPreview(dialogControls, preview);
        };

        prepareDialogWindow(dialogControls.dialog, SCRIPT_NAME);
        if (dialogControls.dialog.show() !== 1) {
            cancelPreview(selectedItems, preview);
            return;
        }
        $.global[SAVED_STATE_KEY] = captureDialogState(dialogControls);

        if (hasSelection) {
            commitToOriginals(doc, selectedItems, preview, dialogControls);
        } else {
            /* 作った長方形をそのまま残す / Keep the preview rectangle as the result */
            applyFinishingOptions(doc, preview.items[0], dialogControls);
            doc.selection = [preview.items[0]];
        }
        app.redraw();
    }

    main();

})();
