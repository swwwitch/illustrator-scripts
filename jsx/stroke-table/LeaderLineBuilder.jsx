#target illustrator
#targetengine "LeaderLineBuilderEngine"
#include "ColorPicker.jsx"

app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパスやグループの外接矩形から、指定した角度の引き出し線を作成し、元のオブジェクトと置き換えます。
角度・斜線の方向・線のスタイル・先端マーカー・フチを、ダイアログでプレビューしながら調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LeaderLineBuilder.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n506df641d5c5

### Overview

Builds a leader line at a chosen angle from the bounding box of the selected paths or groups and replaces the original with it.
Angle, diagonal direction, stroke style, end marker and outline are all adjusted with a preview in the dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LeaderLineBuilder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LeaderLineBuilder";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LeaderLineBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LeaderLineBuilder.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n506df641d5c5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ユーザー設定 / User Settings
    // =========================================

    /* 起動時の値。入力値が不正だったときのフォールバックにも使う */
    var DEFAULT_ANGLE         = 45;         /* 角度（度） */
    var DEFAULT_LINE_WIDTH_PT = 1;          /* 線幅（pt） */
    var DEFAULT_TIP_MARKER_SIZE_PT   = 3;          /* 線端の大きさ（pt） */
    var DEFAULT_TEXT_DIST_PT  = 72 / 25.4;  /* テキストとの距離（pt。1mm相当） */

    /* フチの太らせ方（いずれも線幅に対する倍率） */
    var EDGE_LINE_WIDTH_RATIO = 3;  /* 線のフチの線幅 */
    var EDGE_CIRCLE_EXPAND_RATIO    = 2;  /* 丸のフチの広がり */
    var EDGE_ARROW_EXPAND_RATIO  = 3;  /* 矢印のフチの広がり */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS    = 16;                /* ウィンドウ外周の余白 */
    var WINDOW_SPACING    = 12;                /* ウィンドウ内の要素間隔 */
    var PANEL_MARGINS     = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING     = 6;                 /* パネル内の要素間隔 */
    var COLUMN_SPACING    = 12;                /* 2カラムの間隔 */
    var COLOR_CHIP_SIZE       = [20, 20];          /* カラースウォッチの大きさ */
    var ZOOM_SLIDER_WIDTH = 240;               /* ズームスライダーの幅 */
    var ZOOM_MIN          = 0.1;               /* ズームの下限（倍率） */
    var ZOOM_MAX          = 8;                 /* ズームの上限（倍率） */

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * ダイアログウィンドウに共通のレイアウトを適用する
     * @param {Window} win - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウトを適用する
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びグループに共通のレイアウトを適用する
     * @param {Group} group - 対象のグループ
     * @param {Array<string>} [alignment] - 子要素の整列（省略時は ["left","center"]）
     * @param {number} [spacing] - 要素間隔
     * @returns {void}
     */
    function setupRowGroup(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = alignment || ["left", "center"];
        if (typeof spacing === "number") group.spacing = spacing;
    }

    /**
     * 共通レイアウトを適用したラベル付きパネルを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} labelText - パネルのラベル
     * @param {number} [spacing] - パネル内の要素間隔
     * @returns {Panel} 追加したパネル
     */
    function addLabeledPanel(parent, labelText, spacing) {
        var panel = parent.add("panel", undefined, labelText);
        setupPanel(panel, spacing);
        return panel;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
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
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 設定キーとラジオボタンの対応から、指定キーだけを選択状態にする
     * @param {Array<object>} radioMap - key / radio を持つ要素の配列
     * @param {string} key - 選択するキー
     * @param {string} fallbackKey - key が対応表にないときに選択するキー
     * @returns {void}
     */
    function selectRadioByKey(radioMap, key, fallbackKey) {
        var i;
        var matched = false;
        for (i = 0; i < radioMap.length; i++) {
            if (radioMap[i].key === key) matched = true;
        }
        if (!matched) key = fallbackKey;
        for (i = 0; i < radioMap.length; i++) {
            radioMap[i].radio.value = (radioMap[i].key === key);
        }
    }

    /**
     * 設定キーとラジオボタンの対応から、選択中のキーを取得する
     * @param {Array<object>} radioMap - key / radio を持つ要素の配列
     * @param {string} fallbackKey - どれも選択されていないときに返すキー
     * @returns {string} 選択中のキー
     */
    function getSelectedRadioKey(radioMap, fallbackKey) {
        for (var i = 0; i < radioMap.length; i++) {
            if (radioMap[i].radio.value) return radioMap[i].key;
        }
        return fallbackKey;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "引き出し線ビルダー", en: "Leader Line Builder" }
        },
        panel: {
            applyScope: { ja: "適用範囲", en: "Apply Scope" },
            angle:      { ja: "角度", en: "Angle" },
            direction:  { ja: "斜線の方向", en: "Diagonal Direction" },
            lineStyle:  { ja: "線のスタイル", en: "Line Style" },
            tipMarker:  { ja: "先端マーカー", en: "Tip Marker" },
            edge:       { ja: "フチ", en: "Edge" },
            text:       { ja: "テキスト", en: "Text" }
        },
        radio: {
            applyScopeAll:          { ja: "すべて更新", en: "Update all" },
            applyScopeKeepDir:      { ja: "方向は維持", en: "Keep each direction" },
            diagDirUpperLeft:       { ja: "左上", en: "Upper Left" },
            diagDirLowerLeft:       { ja: "左下", en: "Lower Left" },
            diagDirUpperRight:      { ja: "右上", en: "Upper Right" },
            diagDirLowerRight:      { ja: "右下", en: "Lower Right" },
            lineColorBlack:         { ja: "黒", en: "Black" },
            lineColorWhite:         { ja: "白", en: "White" },
            colorCustom:            { ja: "指定", en: "Custom" },
            strokeCapButt:          { ja: "なし", en: "None" },
            strokeCapRound:         { ja: "丸型", en: "Round" },
            tipMarkerNone:          { ja: "なし", en: "None" },
            tipMarkerCircle:        { ja: "円", en: "Circle" },
            tipMarkerArrow:         { ja: "矢印", en: "Arrow" },
            tipMarkerFill:          { ja: "塗り", en: "Filled" },
            tipMarkerOutline:       { ja: "線のみ", en: "Outlined" }
        },
        checkbox: {
            groupEnabled: { ja: "グループ化", en: "Group items" },
            edgeEnabled:  { ja: "フチを付ける", en: "Add edge" },
            lightMode:    { ja: "簡易", en: "Defer redraw" }
        },
        fieldLabel: {
            lineWidth:     { ja: "線幅", en: "Line Width" },
            strokeCap:     { ja: "線端", en: "Line End" },
            tipMarkerSize: { ja: "大きさ", en: "Size" },
            textDist:      { ja: "テキストとの距離", en: "Distance to text" },
            zoom:          { ja: "ズーム", en: "Zoom" },
            targetMark:    { ja: "基準", en: "Target" }
        },
        tooltip: {
            applyScopeAll: {
                ja: "ダイアログの設定をすべて適用します。",
                en: "Apply every setting in this dialog."
            },
            applyScopeKeepDir: {
                ja: "方向だけは、選択中の引き出し線それぞれに保存された向きを使います。",
                en: "Keep the direction stored in each selected leader line."
            },
            angleInput: {
                ja: "0より大きく90未満。↑↓で次の整数へ、Shift+↑↓で次の10の倍数へ、Option+↑↓で0.1ずつ変わります。",
                en: "Greater than 0 and less than 90. Arrow keys step to the next whole number, Shift to the next multiple of 10, Option by 0.1."
            },
            lengthInput: {
                ja: "↑↓で0.1ずつ、Shift+↑↓で次の10の倍数へ変わります。単位は環境設定に従います。",
                en: "Arrow keys step by 0.1, Shift to the next multiple of 10. The unit follows your preferences."
            },
            diagDir: {
                ja: "引き出し線が伸びる向き。左上なら、対象の左上へ引き出します。",
                en: "Where the leader line runs. Upper Left draws toward the upper left of the object."
            },
            colorChip: {
                ja: "クリックするとカラーピッカーが開きます。",
                en: "Click to open the color picker."
            },
            strokeCap: {
                ja: "線そのものの端の形です。先端マーカーとは別の設定です。",
                en: "The shape of the stroke ends. This is separate from the tip marker."
            },
            tipMarkerStyle: {
                ja: "矢印は塗りのみです。",
                en: "Arrows are always filled."
            },
            tipMarkerSize: {
                ja: "円なら直径、矢印なら長さです。",
                en: "Diameter for a circle, length for an arrow."
            },
            groupEnabled: {
                ja: "線・先端マーカー・フチを1つのグループにまとめます。あとから設定を変えるにはグループ化が必要です。",
                en: "Group the line, tip marker, and edge together. Grouping is required to re-apply settings later."
            },
            edgeEnabled: {
                ja: "線と先端マーカーの背面に、一回り太らせた同じ形を敷きます。",
                en: "Place a thicker copy of the line and tip marker behind them."
            },
            textDist: {
                ja: "テキストを一緒に選択したときだけ有効です。",
                en: "Available only when a text frame is also selected."
            },
            lightMode: {
                ja: "ドラッグ中は再描画せず、離した時点で反映します。",
                en: "Skip redrawing while dragging; apply when the slider is released."
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
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: {
                ja: "パスまたはグループを選択して実行してください。",
                en: "Please select a path or group and run the script."
            },
            noValidTargets: {
                ja: "引き出し線の基準にできるオブジェクトがありません。パス（2点以上）またはグループを選択してください。",
                en: "Nothing can be used as a leader line reference. Select a path with 2 or more points, or a group."
            },
            invalidAngle: {
                ja: "角度は0より大きく90未満の値を入力してください。",
                en: "Enter an angle greater than 0 and less than 90."
            }
        }
    };

    // =========================================
    // セッション記憶 / Session state
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は SETTINGS_STORE_* / createSettingsStore / readSettingsLegacyFile / readSettingsLegacyPreference / settingsStore*
    // 2. 寿命は今のスクリプトに合わせて選ぶ。
    //      "session"    … $.global に置く。Illustrator を終了するまで残る。#targetengine が必須（無いと毎回消える）
    //      "persistent" … Folder.userData/illustrator-scripts/<storeName>.json に書く。再起動しても残る
    //    storeName はふつう SCRIPT_NAME。ダイアログの位置は DialogPosition の部品が持つので、ここには入れない
    // 3. 既定値を1か所にまとめ、load で受け取る。戻り値は毎回新しいオブジェクト（書き換えても保存されない）
    //      var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
    //      var DEFAULT_SETTINGS = { widthPt: 10, addFrame: true, modeKey: "fit", corners: { tl: 0, tr: 0 } };
    //      var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
    //      …OK で閉じたら…
    //      settingsStore.save({ widthPt: …, addFrame: …, modeKey: …, corners: { tl: …, tr: … } });
    //    型は既定値に合わせる（数値の既定値には "12" も 12 として読む。真偽は "1"/"0"/"true"/"false" も読む）。
    //    合わない値・既定値に無い項目は捨てて既定値を使う。{} と null の既定値は中身を問わずそのまま受け取る
    //    （名前をキーにしたプリセット集など）。配列は配列ならそのまま受け取る
    // 4. 保存できるのは文字列・数値・真偽・null と、その配列・入れ子のオブジェクトだけ。
    //    DOM オブジェクト・File・関数は入れない（パスは fsName の文字列で持つ）。長さは pt で持つ
    // 5. 旧形式の設定を読み継ぐときは、3つ目の引数に legacy 関数を渡す。
    //    新しい保存が1度も無いとき（ファイルが無い・$.global に無い）だけ呼ばれ、戻り値を保存値として既定値と突き合わせる。
    //    旧ファイル・旧キーは消さない。キー名が変わったときは legacy の中で詰め替える
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");  … key=value / toSource / JSON を自動判別
    //      } });
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyPreference("SmartTextFindReplace/settings");  … app.preferences の文字列
    //      } });
    // 6. clear() は保存を消す。legacy を渡したストアでは空の保存（{}）を書き、旧設定が戻ってこないようにする
    // 7. 失敗は例外にせず、load は既定値、save は false を返す（$.writeln に理由を出す）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            return textFile.read().replace(/^﻿/, "");
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
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* #targetengine が生きている間だけダイアログの設定を保持する / Keep dialog settings while the #targetengine lives */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session");

    /* 保存する値の既定値 / Default values of the saved settings */
    var DEFAULT_SETTINGS = {
        angle: DEFAULT_ANGLE,
        radioAngle: 45,
        applyScope: "all",
        hasUserSetApplyScope: false,
        diagDir: "upperLeft",
        hDir: "left",
        vDir: "up",
        tipMarkerType: "circle",
        tipMarkerStyle: "fill",
        tipMarkerSize: DEFAULT_TIP_MARKER_SIZE_PT,
        strokeCapType: "round",
        groupEnabled: true,
        whiteEdge: false,
        edgeColor: "white",
        edgeColorHex: "#ffcc00",
        lineColor: "black",
        lineColorHex: "#ffcc00",
        lineWidth: DEFAULT_LINE_WIDTH_PT,
        textDist: DEFAULT_TEXT_DIST_PT
    };

    /* 今回の設定（main で読み込む）/ Current settings, loaded in main */
    var sessionSettings = null;

    // =========================================
    // 単位ユーティリティ / Unit utilities
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位値をptに変換する
     * @param {string|number} value - 単位付きの数値
     * @param {string} prefKey - 環境設定のキー
     * @returns {number} pt値（数値でない場合は NaN）
     */
    function unitValueToPt(value, prefKey) {
        var parsed = parseFloat(value);
        if (isNaN(parsed)) return NaN;
        return parsed * getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * ptを単位値に変換する
     * @param {string|number} valuePt - pt値
     * @param {string} prefKey - 環境設定のキー
     * @returns {number} 単位値（数値でない場合は NaN）
     */
    function ptValueToUnit(valuePt, prefKey) {
        var parsed = parseFloat(valuePt);
        if (isNaN(parsed)) return NaN;
        return parsed / getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * 指定桁で四捨五入する
     * @param {number} value - 対象の値
     * @param {number} [digits] - 小数桁数
     * @returns {number} 丸めた値
     */
    function roundTo(value, digits) {
        var factor = Math.pow(10, digits || 0);
        return Math.round(value * factor) / factor;
    }

    /**
     * 数値を表示用の文字列に整形する
     * @param {number} value - 対象の値
     * @param {number} [digits] - 小数桁数
     * @returns {string} 整形した文字列（数値でない場合は空文字）
     */
    function formatNumber(value, digits) {
        if (isNaN(value)) return "";
        var text = String(roundTo(value, digits || 0));
        text = text.replace(/(\.\d*?)0+$/, "$1");
        /* digits>=1 のとき最低1桁の小数を保持（1 → 1.0） */
        if ((digits || 0) >= 1 && text.indexOf(".") === -1) text += ".0";
        return text;
    }

    /**
     * pt値を現在の単位の表示用文字列に整形する
     * @param {string|number} valuePt - pt値
     * @param {string} prefKey - 環境設定のキー
     * @returns {string} 整形した文字列
     */
    function formatPtInCurrentUnit(valuePt, prefKey) {
        return formatNumber(ptValueToUnit(valuePt, prefKey), 2);
    }

    /**
     * pt値を線幅系の入力欄に表示する文字列に整形する
     * @param {string|number} valuePt - pt値
     * @param {number} fallbackValue - 値が不正だったときに表示する値
     * @returns {string} 現在の単位に換算した文字列
     */
    function formatUnitInput(valuePt, fallbackValue) {
        var displayValue = parseFloat(formatPtInCurrentUnit(valuePt, "strokeUnits"));
        if (isNaN(displayValue) || displayValue <= 0) displayValue = fallbackValue;
        return formatNumber(displayValue, 2);
    }

    /**
     * 線幅系の入力欄の値を、セッション記憶用のpt値に変換する
     * @param {string} unitText - 単位付きの数値
     * @param {number} fallbackPt - 値が不正だったときに使うpt値
     * @returns {string|number} pt値
     */
    function parseUnitInput(unitText, fallbackPt) {
        var valuePt = unitValueToPt(unitText, "strokeUnits");
        return (!isNaN(valuePt) && valuePt > 0) ? formatNumber(valuePt, 4) : fallbackPt;
    }

    // =========================================
    // ズーム操作 / View zoom
    // =========================================

    /**
     * 現在のビューの表示倍率と中心を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} ビュー・倍率・中心を持つ状態オブジェクト
     */
    function captureViewState(doc) {
        var state = { view: null, zoom: null, center: null };
        try {
            state.view = doc.activeView;
            state.zoom = state.view.zoom;
            state.center = state.view.centerPoint;
        } catch (e) { }
        return state;
    }

    /**
     * 控えておいたビューの表示倍率と中心を戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {object} state - captureViewState() の戻り値
     * @returns {void}
     */
    function restoreViewState(doc, state) {
        if (!state) return;
        try {
            var view = state.view || doc.activeView;
            if (view && state.zoom != null) view.zoom = state.zoom;
            if (view && state.center != null) view.centerPoint = state.center;
        } catch (e) { }
    }

    /**
     * ズームスライダーと「軽」チェックボックスをダイアログに追加する
     * 「軽」がONのときはドラッグ中に反映せず、離した時点でだけ倍率を適用する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {Document} doc - 対象ドキュメント
     * @param {string} labelText - スライダーのラベル
     * @param {object} initialState - captureViewState() の戻り値
     * @param {object} [options] - 表示と挙動のオプション
     * @returns {object} スライダーなどの参照と操作関数をまとめたオブジェクト
     */
    function addZoomControls(parent, doc, labelText, initialState, options) {
        options = options || {};
        var minZoom = (typeof options.min === "number") ? options.min : ZOOM_MIN;
        var maxZoom = (typeof options.max === "number") ? options.max : ZOOM_MAX;
        var sliderWidth = (typeof options.sliderWidth === "number") ? options.sliderWidth : ZOOM_SLIDER_WIDTH;
        var doRedraw = (options.redraw !== false);
        var showLightMode = (options.lightMode !== false);
        var lightModeLabel = options.lightModeLabel || "Light mode";
        var lightModeDefault = (options.lightModeDefault === true);

        var zoomGroup = parent.add("group");
        setupRowGroup(zoomGroup, ["center", "center"]);
        zoomGroup.alignment = "center";
        if (options.margins) zoomGroup.margins = options.margins;

        zoomGroup.add("statictext", undefined, String(labelText || "Zoom"));

        var initialZoom = 1;
        try {
            if (initialState && initialState.zoom != null) initialZoom = Number(initialState.zoom);
            else initialZoom = Number(doc.activeView.zoom);
        } catch (e) { }
        if (!initialZoom || isNaN(initialZoom)) initialZoom = 1;

        var zoomSlider = zoomGroup.add("slider", undefined, initialZoom, minZoom, maxZoom);
        zoomSlider.preferredSize.width = sliderWidth;

        var lightModeCheckbox = null;
        if (showLightMode) {
            lightModeCheckbox = zoomGroup.add("checkbox", undefined, String(lightModeLabel));
            lightModeCheckbox.value = lightModeDefault;
            if (options.lightModeTip) lightModeCheckbox.helpTip = String(options.lightModeTip);
        }

        /**
         * 「軽」がONかどうかを返す
         * @returns {boolean} ONなら true
         */
        function isLightMode() {
            return !!(lightModeCheckbox && lightModeCheckbox.value);
        }

        /**
         * 表示倍率を適用する
         * @param {number} zoom - 表示倍率
         * @returns {void}
         */
        function applyZoom(zoom) {
            try {
                var view = (initialState && initialState.view) ? initialState.view : doc.activeView;
                if (!view) return;
                view.zoom = zoom;
                if (doRedraw) app.redraw();
            } catch (e) { }
        }

        zoomSlider.onChanging = function () {
            if (isLightMode()) return;
            applyZoom(Number(zoomSlider.value));
        };

        zoomSlider.onChange = function () {
            applyZoom(Number(zoomSlider.value));
        };

        if (lightModeCheckbox) {
            lightModeCheckbox.onClick = function () {
                applyZoom(Number(zoomSlider.value));
            };
        }

        return {
            slider: zoomSlider,
            lightModeCheckbox: lightModeCheckbox,
            applyZoom: applyZoom,
            restoreInitial: function () { restoreViewState(doc, initialState); }
        };
    }

    // =========================================
    // 引き出し線の方向 / Leader line direction
    // =========================================

    /* 斜線の方向名と、水平・垂直の向きの対応 */
    var DIAG_DIRECTIONS = [
        { name: "upperLeft",  hDir: "right", vDir: "down" },
        { name: "lowerLeft",  hDir: "right", vDir: "up" },
        { name: "upperRight", hDir: "left",  vDir: "down" },
        { name: "lowerRight", hDir: "left",  vDir: "up" }
    ];

    /**
     * 斜線の方向名から水平・垂直の向きを求める
     * @param {string} diagDir - 方向名（"upperLeft" など）
     * @returns {object} hDir / vDir を持つオブジェクト
     */
    function hvDirFromDiagDir(diagDir) {
        for (var i = 0; i < DIAG_DIRECTIONS.length; i++) {
            if (DIAG_DIRECTIONS[i].name === diagDir) {
                return { hDir: DIAG_DIRECTIONS[i].hDir, vDir: DIAG_DIRECTIONS[i].vDir };
            }
        }
        return { hDir: "right", vDir: "up" };
    }

    /**
     * 水平・垂直の向きから斜線の方向名を求める
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @returns {string|null} 方向名（該当しない場合は null）
     */
    function diagDirFromHVDir(hDir, vDir) {
        for (var i = 0; i < DIAG_DIRECTIONS.length; i++) {
            if (DIAG_DIRECTIONS[i].hDir === hDir && DIAG_DIRECTIONS[i].vDir === vDir) {
                return DIAG_DIRECTIONS[i].name;
            }
        }
        return null;
    }

    /**
     * 斜線の方向ラジオから水平・垂直の向きを求める
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {object} hDir / vDir を持つオブジェクト
     */
    function getDiagDirValues(ui) {
        return hvDirFromDiagDir(getSelectedRadioKey(ui.radioMaps.diagDir, "lowerLeft"));
    }

    /**
     * 引き出し線を1つだけ選んでいるとき、保存済みの方向をセッション記憶へ戻す
     * @param {Array<PathItem|GroupItem>} targetItems - 対象アイテム
     * @returns {void}
     */
    function restoreDirFromSelection(targetItems) {
        if (targetItems.length !== 1) return;
        var target = targetItems[0];
        if (!isLeaderLineGroup(target)) return;
        var storedDir = getLeaderLineStoredDir(target);
        if (!storedDir) return;
        var diagDir = diagDirFromHVDir(storedDir.hDir, storedDir.vDir);
        if (!diagDir) return;
        sessionSettings.diagDir = diagDir;
        sessionSettings.hDir = storedDir.hDir;
        sessionSettings.vDir = storedDir.vDir;
    }

    // =========================================
    // アイテム操作 / Item helpers
    // =========================================

    /**
     * アイテムを安全に削除する
     * @param {object} item - 対象アイテム
     * @returns {void}
     */
    function safeRemove(item) {
        if (!item) return;
        try {
            if (item.parent) item.remove();
        } catch (e) { }
    }

    /**
     * アイテムを安全に選択する
     * @param {object} item - 対象アイテム
     * @returns {void}
     */
    function safeSelect(item) {
        if (!item) return;
        try {
            item.selected = true;
        } catch (e) { }
    }

    /**
     * アイテムのメモ（note）を安全に設定する
     * @param {object} item - 対象アイテム
     * @param {string} note - 設定する文字列
     * @returns {void}
     */
    function safeSetNote(item, note) {
        if (!item) return;
        item.note = note;
    }

    /**
     * 編集可能なレイヤーをアクティブにする
     * 対象アイテムのレイヤーが使えればそこへ、なければ最初の編集可能レイヤーへ切り替える
     * @param {Document} doc - 対象ドキュメント
     * @param {object} item - 対象アイテム
     * @returns {void}
     */
    function switchToEditableLayer(doc, item) {
        try {
            var itemLayer = item.layer;
            if (itemLayer && !itemLayer.locked && itemLayer.visible) {
                doc.activeLayer = itemLayer;
                return;
            }
        } catch (e) { }
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.locked && layer.visible) {
                doc.activeLayer = layer;
                return;
            }
        }
    }

    // =========================================
    // メモとタグ / Note and tag
    // =========================================

    /**
     * アイテムのタグに値を設定する（同名タグがあれば上書き）
     * @param {object} item - 対象アイテム
     * @param {string} name - タグ名
     * @param {string|number} value - 設定する値
     * @returns {void}
     */
    function setTagValue(item, name, value) {
        if (!item) return;
        try {
            for (var i = 0; i < item.tags.length; i++) {
                if (item.tags[i].name === name) {
                    item.tags[i].value = String(value);
                    return;
                }
            }
            var tag = item.tags.add();
            tag.name = name;
            tag.value = String(value);
        } catch (e) { }
    }

    /**
     * アイテムのタグから値を取得する
     * @param {object} item - 対象アイテム
     * @param {string} name - タグ名
     * @returns {string|null} タグの値（なければ null）
     */
    function getTagValue(item, name) {
        if (!item) return null;
        for (var i = 0; i < item.tags.length; i++) {
            if (item.tags[i].name === name) return item.tags[i].value;
        }
        return null;
    }

    /**
     * 引き出し線グループであることを示すメモ（note）を付ける
     * @param {GroupItem} item - 対象のグループ
     * @returns {void}
     */
    function setLeaderLineTag(item) {
        safeSetNote(item, "leader_line");
    }

    /**
     * 引き出し線の各パーツに再適用用のメモ（note）を付ける
     * @param {object} parts - assembleLeaderParts() の戻り値
     * @param {PathItem} edgePath - フチの線
     * @param {PathItem} mainPath - 本体の線
     * @param {PathItem} edgeTipMarker - フチの線端
     * @param {PathItem} mainTipMarker - 本体の線端
     * @param {boolean} hasEdge - フチがあるかどうか
     * @returns {void}
     */
    function tagLeaderParts(parts, edgePath, mainPath, edgeTipMarker, mainTipMarker, hasEdge) {
        if (hasEdge) {
            safeSetNote(parts.mainGroup, "leader_line_main");
            safeSetNote(parts.edgeGroup, "leader_line_edge");
        }

        safeSetNote(mainPath, "leader_line_main_path");
        safeSetNote(mainTipMarker, "leader_line_main_cap");
        safeSetNote(edgePath, "leader_line_edge_path");
        safeSetNote(edgeTipMarker, "leader_line_edge_cap");
    }

    /**
     * 再適用時に使う元オブジェクトの外接矩形をタグに保存する
     * @param {GroupItem} item - 対象のグループ
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {void}
     */
    function setLeaderLineBoundsTags(item, targetMetrics) {
        if (!item || !targetMetrics) return;
        setTagValue(item, "leader_line_x_left", formatNumber(targetMetrics.x_left, 4));
        setTagValue(item, "leader_line_x_right", formatNumber(targetMetrics.x_right, 4));
        setTagValue(item, "leader_line_y_top", formatNumber(targetMetrics.y_top, 4));
        setTagValue(item, "leader_line_y_bottom", formatNumber(targetMetrics.y_bottom, 4));
    }

    /**
     * 再適用時に使う斜線の方向をタグに保存する
     * @param {GroupItem} item - 対象のグループ
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @returns {void}
     */
    function setLeaderLineDirTags(item, hDir, vDir) {
        if (!item) return;
        setTagValue(item, "leader_line_hDir", hDir);
        setTagValue(item, "leader_line_vDir", vDir);
    }

    /**
     * タグに保存された斜線の方向を取得する
     * @param {GroupItem} item - 対象のグループ
     * @returns {object|null} hDir / vDir を持つオブジェクト（なければ null）
     */
    function getLeaderLineStoredDir(item) {
        var hDir = getTagValue(item, "leader_line_hDir");
        var vDir = getTagValue(item, "leader_line_vDir");
        if (!hDir || !vDir) return null;
        return { hDir: hDir, vDir: vDir };
    }

    /**
     * タグに保存された外接矩形を取得する
     * @param {GroupItem} item - 対象のグループ
     * @returns {object|null} x_left / x_right / y_top / y_bottom を持つ矩形（なければ null）
     */
    function getLeaderLineStoredBounds(item) {
        var xLeft = parseFloat(getTagValue(item, "leader_line_x_left"));
        var xRight = parseFloat(getTagValue(item, "leader_line_x_right"));
        var yTop = parseFloat(getTagValue(item, "leader_line_y_top"));
        var yBottom = parseFloat(getTagValue(item, "leader_line_y_bottom"));
        if (isNaN(xLeft) || isNaN(xRight) || isNaN(yTop) || isNaN(yBottom)) return null;
        return {
            x_left: xLeft,
            x_right: xRight,
            y_top: yTop,
            y_bottom: yBottom
        };
    }

    /**
     * 引き出し線グループかどうかを判定する
     * @param {object} item - 対象アイテム
     * @returns {boolean} 引き出し線グループなら true
     */
    function isLeaderLineGroup(item) {
        return item && item.typename === "GroupItem" && item.note === "leader_line";
    }

    /**
     * 引き出し線グループから線端（丸・矢印）のパスを取得する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PathItem|null} 線端のパス（なければ null）
     */
    function getLeaderLineTipMarker(groupItem) {
        for (var i = 0; i < groupItem.pathItems.length; i++) {
            var pathItem = groupItem.pathItems[i];
            if (pathItem.note === "leader_line_main_cap") return pathItem;
        }
        return null;
    }

    /**
     * 引き出し線グループから基準線（A線）を取得する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PathItem|null} 基準線（なければ null）
     */
    function getLeaderLineBasePath(groupItem) {
        var i, pathItem;

        /* 再適用時の基準線は常に A線のみとし、B/C/D は参照しない
           新構造なら note 付きの A線を最優先 */
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pathItem = groupItem.pathItems[i];
            if (pathItem.note === "leader_line_main_path") return pathItem;
        }

        /* 旧構造用フォールバック：白でない stroked の open path を優先 */
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pathItem = groupItem.pathItems[i];
            if (pathItem.stroked && !pathItem.closed && !isWhiteCMYKColor(pathItem.strokeColor)) return pathItem;
        }

        /* 最後のフォールバック：stroked の open path */
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pathItem = groupItem.pathItems[i];
            if (pathItem.stroked && !pathItem.closed) return pathItem;
        }

        return null;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。使わない関数も消さずに残してよい（互いに呼び合う）。
    //    識別子は SELECTION_ITEMS_TOLERANCE / normalizeSelectionItems / resolveTextRangeFrame /
    //    collectSelectionItems / getTextFrameKindKey / collectSelectionTextFrames / collectSelectionPathItems /
    //    isClipMaskItem / getClipMaskItem / hasClippedDescendant / readUsePreviewBoundsPreference /
    //    getClipAwareBounds / filterMeasurableChildren / getClipAwareUnionBounds / isNearlySameCoordinate / areBoundsNearlyEqual
    // 2. 選択は normalizeSelectionItems(doc.selection) で配列にする。文字カーソルの選択（TextRange）は
    //    配列ではなく1個で返り、しかも .length（文字数）を持つので、length だけで配列と見なさない
    // 3. テキストフレーム:
    //      var frames = collectSelectionTextFrames(doc.selection);                           // 全種類
    //      var frames = collectSelectionTextFrames(doc.selection, { kinds: ["point", "path"] });
    //    パス:
    //      var paths = collectSelectionPathItems(doc.selection);                             // 複合パスは中のパスへ
    //      var paths = collectSelectionPathItems(doc.selection, { compoundPaths: "whole", skipClipMasks: true });
    //    それ以外は collectSelectionItems(source, { accept: function (item) { … } }) で条件を書く
    // 4. 並びは選択と同じ前面→背面（グループの中も pageItems の順）。重なり順を使う処理はこの順を前提にしてよい
    // 5. doc.selection に代入し直す配列は skipLocked / skipHidden を true にする。
    //    ロック・非表示を選択に代入すると例外になり、中の子が選択に残る
    // 6. 境界は getClipAwareBounds(item, usePreviewBounds) / getClipAwareUnionBounds(items, usePreviewBounds)。
    //    usePreviewBounds を省くと環境設定の［プレビュー境界を使用］に従う。返り値は [左, 上, 右, 下] の新しい配列
    //    （書き換えても元のオブジェクトに影響しない）。測れないときは null
    // 7. 座標の一致・前後の判定は isNearlySameCoordinate / areBoundsNearlyEqual で許容値を挟む
    //    （吸着させた辺とガイドは 1e-12 ほどずれる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 外接矩形と線の属性 / Bounds and stroke
    // =========================================

    /**
     * GroupItem の参照用 PathItem を取得する
     * 通常は closed path を優先して返す（leader_line 再適用時の A線優先ロジックは getLeaderLineBasePath() 側で処理）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {Array<PathItem>} 参照用のパス配列
     */
    function getGroupReferencePathItems(groupItem) {
        var closedPaths = [];
        var allPaths = [];
        for (var k = 0; k < groupItem.pathItems.length; k++) {
            var pathItem = groupItem.pathItems[k];
            allPaths.push(pathItem);
            if (pathItem.closed) closedPaths.push(pathItem);
        }
        return closedPaths.length ? closedPaths : allPaths;
    }

    /**
     * PathItem のアンカー座標から外接矩形を取得する
     * @param {PathItem} pathItem - 対象のパス
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getPathAnchorBounds(pathItem) {
        var points = pathItem.pathPoints;
        if (!points || points.length === 0) {
            return {
                x_left: 0,
                x_right: 0,
                y_top: 0,
                y_bottom: 0
            };
        }

        var xMin = points[0].anchor[0], xMax = points[0].anchor[0];
        var yMin = points[0].anchor[1], yMax = points[0].anchor[1];
        for (var i = 1; i < points.length; i++) {
            var anchorX = points[i].anchor[0];
            var anchorY = points[i].anchor[1];
            if (anchorX < xMin) xMin = anchorX;
            if (anchorX > xMax) xMax = anchorX;
            if (anchorY < yMin) yMin = anchorY;
            if (anchorY > yMax) yMax = anchorY;
        }

        return {
            x_left: xMin,
            x_right: xMax,
            y_top: yMax,
            y_bottom: yMin
        };
    }

    /**
     * 線端パスの中心座標を取得する
     * @param {PathItem} capItem - 線端のパス
     * @returns {Array<number>|null} [x, y]（なければ null）
     */
    function getTipMarkerCenter(capItem) {
        if (!capItem) return null;
        var capBounds = capItem.geometricBounds; /* [left, top, right, bottom] */
        return [
            (capBounds[0] + capBounds[2]) / 2,
            (capBounds[1] + capBounds[3]) / 2
        ];
    }

    /**
     * 線端の中心を先端とみなして、引き出し線の外接矩形を求める
     * @param {PathItem} pathItem - 引き出し線のパス
     * @param {PathItem} capItem - 線端のパス
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getLeaderLineBoundsWithTipMarker(pathItem, capItem) {
        var bounds = getPathAnchorBounds(pathItem);
        if (!capItem) return bounds;

        var points = pathItem.pathPoints;
        if (!points || points.length < 2) return bounds;

        var p0 = points[0].anchor;
        var p1 = points[1].anchor;
        var p2 = points[points.length - 1].anchor;
        var capCenter = getTipMarkerCenter(capItem);
        if (!capCenter) return bounds;

        /* 線端に丸がある場合、丸の中心（= 丸の半分位置）を期待する端として扱う */
        var centerX = capCenter[0];
        var centerY = capCenter[1];

        var distanceToStart = Math.abs(p0[0] - centerX) + Math.abs(p0[1] - centerY);
        var distanceToEnd = Math.abs(p2[0] - centerX) + Math.abs(p2[1] - centerY);

        var tipStart = [p0[0], p0[1]];
        var bend = [p1[0], p1[1]];
        var tipEnd = [p2[0], p2[1]];

        if (distanceToStart <= distanceToEnd) {
            tipStart = [centerX, centerY];
        } else {
            tipEnd = [centerX, centerY];
        }

        return {
            x_left: Math.min(tipStart[0], bend[0], tipEnd[0]),
            x_right: Math.max(tipStart[0], bend[0], tipEnd[0]),
            y_top: Math.max(tipStart[1], bend[1], tipEnd[1]),
            y_bottom: Math.min(tipStart[1], bend[1], tipEnd[1])
        };
    }

    /**
     * GroupItem の外接矩形を取得する
     * 通常はグループ全体（クリップグループはマスク）、leader_line 再適用時は保存済み bounds を優先し、なければ旧構造から復元する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getGroupBounds(groupItem) {
        if (isLeaderLineGroup(groupItem)) {
            var storedBounds = getLeaderLineStoredBounds(groupItem);
            if (storedBounds) {
                return storedBounds;
            }

            var basePath = getLeaderLineBasePath(groupItem);
            if (basePath) {
                var mainTipMarker = getLeaderLineTipMarker(groupItem);
                return getLeaderLineBoundsWithTipMarker(basePath, mainTipMarker);
            }
        }

        /* クリップグループはマスクの範囲で測る / Clip groups are measured by their mask */
        var geoBounds = getClipAwareBounds(groupItem, false) || groupItem.geometricBounds; /* [left, top, right, bottom] */
        return {
            x_left: geoBounds[0],
            x_right: geoBounds[2],
            y_top: geoBounds[1],
            y_bottom: geoBounds[3]
        };
    }

    /**
     * 対象の線の属性を取得する
     * GroupItem は通常 closed path 優先、leader_line 再適用時は A線を優先する
     * @param {PathItem|GroupItem} item - 対象アイテム
     * @returns {object} stroked / strokeWidth / strokeColor を持つオブジェクト
     */
    function getStrokeInfo(item) {
        if (item.typename === "PathItem") {
            return {
                stroked: item.stroked,
                strokeWidth: item.stroked ? item.strokeWidth : DEFAULT_LINE_WIDTH_PT,
                strokeColor: item.stroked ? item.strokeColor : undefined
            };
        }

        /* 引き出し線グループの場合、A線（基準線）から取得 */
        if (isLeaderLineGroup(item)) {
            var leaderBasePath = getLeaderLineBasePath(item);
            if (leaderBasePath) {
                return {
                    stroked: leaderBasePath.stroked,
                    strokeWidth: leaderBasePath.stroked ? leaderBasePath.strokeWidth : DEFAULT_LINE_WIDTH_PT,
                    strokeColor: leaderBasePath.stroked ? leaderBasePath.strokeColor : undefined
                };
            }
        }

        /* GroupItemの場合、closed pathを優先して配下のPathItemを探す */
        var referencePaths = getGroupReferencePathItems(item);
        for (var k = 0; k < referencePaths.length; k++) {
            var pathItem = referencePaths[k];
            if (pathItem.stroked) {
                return {
                    stroked: true,
                    strokeWidth: pathItem.strokeWidth,
                    strokeColor: pathItem.strokeColor
                };
            }
        }
        return { stroked: false, strokeWidth: DEFAULT_LINE_WIDTH_PT, strokeColor: undefined };
    }

    /**
     * 各ターゲットの座標情報と線の属性を都度取得する
     * @param {Array<PathItem|GroupItem>} itemsToMeasure - 対象アイテムの配列
     * @returns {Array<object>} 外接矩形と線の属性をまとめた配列
     */
    function measureTargets(itemsToMeasure) {
        var data = [];
        for (var i = 0; i < itemsToMeasure.length; i++) {
            var item = itemsToMeasure[i];
            var x_left, x_right, y_bottom, y_top;

            if (item.typename === "PathItem") {
                /* PathItem：アンカーポイントから外接矩形を算出 */
                var pathBounds = getPathAnchorBounds(item);
                x_left = pathBounds.x_left;
                x_right = pathBounds.x_right;
                y_bottom = pathBounds.y_bottom;
                y_top = pathBounds.y_top;
            } else {
                /* GroupItem：通常はグループ全体、leader_line 再適用時は保存済み bounds を優先して使用 */
                var groupBounds = getGroupBounds(item);
                x_left = groupBounds.x_left;
                x_right = groupBounds.x_right;
                y_top = groupBounds.y_top;
                y_bottom = groupBounds.y_bottom;
            }

            var strokeInfo = getStrokeInfo(item);
            data.push({
                target: item,
                x_left: x_left,
                x_right: x_right,
                y_bottom: y_bottom,
                y_top: y_top,
                stroked: strokeInfo.stroked,
                strokeWidth: strokeInfo.strokeWidth,
                strokeColor: strokeInfo.strokeColor
            });
        }
        return data;
    }

    /**
     * テキストフレームの外接矩形を取得する
     * @param {TextFrame} textItem - 対象のテキストフレーム
     * @returns {object} x_left / x_right / y_top / y_bottom を持つ矩形
     */
    function getTextBounds(textItem) {
        var geoBounds = textItem.geometricBounds; /* [left, top, right, bottom] */
        return { x_left: geoBounds[0], x_right: geoBounds[2], y_top: geoBounds[1], y_bottom: geoBounds[3] };
    }

    // =========================================
    // カラー / Color
    // =========================================

    /**
     * HEXカラーを RGBColor に変換する
     * @param {string} hex - "#rrggbb" または "#rgb"
     * @returns {RGBColor|null} 変換したカラー（不正な場合は null）
     */
    function hexToRGBColor(hex) {
        hex = hex.replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        var red = parseInt(hex.substring(0, 2), 16);
        var green = parseInt(hex.substring(2, 4), 16);
        var blue = parseInt(hex.substring(4, 6), 16);
        if (isNaN(red) || isNaN(green) || isNaN(blue)) return null;
        var color = new RGBColor();
        color.red = red;
        color.green = green;
        color.blue = blue;
        return color;
    }

    /**
     * ドキュメントのカラースペースがCMYKかどうかを判定する
     * @param {Document} doc - 対象ドキュメント
     * @returns {boolean} CMYKなら true
     */
    function isCMYKDocument(doc) {
        return doc.documentColorSpace === DocumentColorSpace.CMYK;
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を作る
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラー
     */
    function createBlackColor(doc) {
        if (isCMYKDocument(doc)) {
            var cmykBlack = new CMYKColor();
            cmykBlack.cyan = 0;
            cmykBlack.magenta = 0;
            cmykBlack.yellow = 0;
            cmykBlack.black = 100;
            return cmykBlack;
        }
        var rgbBlack = new RGBColor();
        rgbBlack.red = 0;
        rgbBlack.green = 0;
        rgbBlack.blue = 0;
        return rgbBlack;
    }

    /**
     * ドキュメントのカラースペースに合わせた白を作る
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 白のカラー
     */
    function createWhiteColor(doc) {
        if (isCMYKDocument(doc)) {
            var cmykWhite = new CMYKColor();
            cmykWhite.cyan = 0;
            cmykWhite.magenta = 0;
            cmykWhite.yellow = 0;
            cmykWhite.black = 0;
            return cmykWhite;
        }
        var rgbWhite = new RGBColor();
        rgbWhite.red = 255;
        rgbWhite.green = 255;
        rgbWhite.blue = 255;
        return rgbWhite;
    }

    /**
     * RGB値を CMYKColor に変換する
     * @param {number} red - 赤（0-255）
     * @param {number} green - 緑（0-255）
     * @param {number} blue - 青（0-255）
     * @returns {CMYKColor} 変換したカラー
     */
    function rgbToCMYKColor(red, green, blue) {
        var redRatio = Math.max(0, Math.min(255, red)) / 255;
        var greenRatio = Math.max(0, Math.min(255, green)) / 255;
        var blueRatio = Math.max(0, Math.min(255, blue)) / 255;
        var blackRatio = 1 - Math.max(redRatio, greenRatio, blueRatio);
        var cyanRatio = 0, magentaRatio = 0, yellowRatio = 0;

        if (blackRatio < 1) {
            cyanRatio = (1 - redRatio - blackRatio) / (1 - blackRatio);
            magentaRatio = (1 - greenRatio - blackRatio) / (1 - blackRatio);
            yellowRatio = (1 - blueRatio - blackRatio) / (1 - blackRatio);
        }

        var color = new CMYKColor();
        color.cyan = Math.round(cyanRatio * 100);
        color.magenta = Math.round(magentaRatio * 100);
        color.yellow = Math.round(yellowRatio * 100);
        color.black = Math.round(blackRatio * 100);
        return color;
    }

    /**
     * HEXカラーをドキュメントのカラースペースに合わせて変換する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} hex - "#rrggbb" または "#rgb"
     * @returns {CMYKColor|RGBColor|null} 変換したカラー（不正な場合は null）
     */
    function hexToDocumentColor(doc, hex) {
        var rgb = hexToRGBColor(hex);
        if (!rgb) return null;
        if (isCMYKDocument(doc)) {
            return rgbToCMYKColor(rgb.red, rgb.green, rgb.blue);
        }
        return rgb;
    }

    /**
     * 白のCMYKカラーかどうかを判定する
     * @param {object} color - 対象のカラー
     * @returns {boolean} 白のCMYKカラーなら true
     */
    function isWhiteCMYKColor(color) {
        if (!color) return false;
        if (color.typename !== "CMYKColor") return false;
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    }

    /**
     * スウォッチ表示を指定のHEX値に合わせて更新する
     * @param {Panel} colorChip - 対象のスウォッチ
     * @param {string} hex - "#rrggbb" または "#rgb"
     * @returns {void}
     */
    function updateColorChip(colorChip, hex) {
        if (!colorChip) return;
        var rgb = hexToRGBColor(String(hex || ""));
        if (!rgb) return;
        var graphics = colorChip.graphics;
        graphics.backgroundColor = graphics.newBrush(
            graphics.BrushType.SOLID_COLOR,
            [rgb.red / 255, rgb.green / 255, rgb.blue / 255]
        );
    }

    /**
     * 生成条件からフチのカラーを求める
     * @param {Document} doc - 対象ドキュメント
     * @param {object} options - readLeaderOptions() の戻り値
     * @returns {CMYKColor|RGBColor} フチのカラー
     */
    function resolveEdgeColor(doc, options) {
        if (options.edgeColorMode === "white") {
            return createWhiteColor(doc);
        }
        var documentColor = hexToDocumentColor(doc, options.edgeColorHex);
        if (documentColor) return documentColor;
        return createWhiteColor(doc);
    }

    /**
     * 生成条件から線のカラーを求める
     * @param {Document} doc - 対象ドキュメント
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {CMYKColor|RGBColor} 線のカラー
     */
    function resolveLineColor(doc, options, targetMetrics) {
        if (options.lineColorMode === "black") {
            return createBlackColor(doc);
        }
        if (options.lineColorMode === "white") {
            return createWhiteColor(doc);
        }
        /* その他：HEXカラーコードから現在のドキュメント色空間に合わせて変換 */
        var documentColor = hexToDocumentColor(doc, options.lineColorHex);
        if (documentColor) return documentColor;
        /* 無効な値の場合は元のオブジェクトの線色を使用 */
        if (targetMetrics.strokeColor) return targetMetrics.strokeColor;
        return createBlackColor(doc);
    }

    /**
     * 生成条件から線幅（pt）を求める
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {number} 線幅（pt）
     */
    function resolveLineWidth(options, targetMetrics) {
        var widthPt = options.lineWidthPt;
        if (!isNaN(widthPt) && widthPt > 0) return widthPt;
        return targetMetrics.strokeWidth;
    }

    /**
     * 引き出し線1本分の描画スタイルをまとめて求める
     * @param {Document} doc - 対象ドキュメント
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {object} targetMetrics - measureTargets() の要素
     * @returns {object} lineColor / edgeColor / widthPt を持つオブジェクト
     */
    function resolveLeaderStyle(doc, options, targetMetrics) {
        return {
            lineColor: resolveLineColor(doc, options, targetMetrics),
            edgeColor: resolveEdgeColor(doc, options),
            widthPt: resolveLineWidth(options, targetMetrics)
        };
    }

    // =========================================
    // 引き出し線の座標計算 / Leader line geometry
    // =========================================

    /**
     * 引き出し線の座標を計算する
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @returns {Array<Array<number>>} 3点の座標配列
     */
    function calcLeaderPoints(targetMetrics, angleRad, hDir, vDir) {
        var height = targetMetrics.y_top - targetMetrics.y_bottom;
        var offset = height / Math.tan(angleRad);
        var bendX;
        if (vDir === "down") {
            /* 水平線が上、斜線が下に向かう */
            if (hDir === "left") {
                bendX = targetMetrics.x_left + offset;
                return [[targetMetrics.x_left, targetMetrics.y_bottom], [bendX, targetMetrics.y_top], [targetMetrics.x_right, targetMetrics.y_top]];
            }
            bendX = targetMetrics.x_right - offset;
            return [[targetMetrics.x_left, targetMetrics.y_top], [bendX, targetMetrics.y_top], [targetMetrics.x_right, targetMetrics.y_bottom]];
        }
        /* 水平線が下、斜線が上に向かう */
        if (hDir === "left") {
            bendX = targetMetrics.x_left + offset;
            return [[targetMetrics.x_left, targetMetrics.y_top], [bendX, targetMetrics.y_bottom], [targetMetrics.x_right, targetMetrics.y_bottom]];
        }
        bendX = targetMetrics.x_right - offset;
        return [[targetMetrics.x_left, targetMetrics.y_bottom], [bendX, targetMetrics.y_bottom], [targetMetrics.x_right, targetMetrics.y_top]];
    }

    /**
     * テキスト整列時の引き出し線座標を計算する
     * オブジェクトとテキストの位置関係から方向を自動判定し、水平端をテキストの対応する端にそろえ、
     * 折れ点のY座標をテキスト端から指定距離だけ離す
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {object} textBounds - テキストの外接矩形
     * @param {number} textDistPt - テキストとの距離（pt）
     * @returns {object} points（3点の座標配列）／hDir／vDir を持つオブジェクト
     */
    function calcLeaderPointsWithText(targetMetrics, angleRad, textBounds, textDistPt) {
        var bendY, horizontalEndX, tipX, tipY, diffY, diffX, bendX;

        /* オブジェクトとテキストの中心座標を比較して方向を自動判定
           （Illustratorは Y 上向き正: 下にあるほど Y が小さい） */
        var objectCenterX = (targetMetrics.x_left + targetMetrics.x_right) / 2;
        var objectCenterY = (targetMetrics.y_top + targetMetrics.y_bottom) / 2;
        var textCenterX = (textBounds.x_left + textBounds.x_right) / 2;
        var textCenterY = (textBounds.y_top + textBounds.y_bottom) / 2;
        var autoHDir = (objectCenterX < textCenterX) ? "left" : "right";
        var autoVDir = (objectCenterY < textCenterY) ? "down" : "up";

        /* オブジェクトがテキストより下 → 折れ点はテキスト下端-textDistPt
           オブジェクトがテキストより上 → 折れ点はテキスト上端+textDistPt */
        if (autoVDir === "down") {
            bendY = textBounds.y_bottom - textDistPt;
        } else {
            bendY = textBounds.y_top + textDistPt;
        }

        tipY = (autoVDir === "down") ? targetMetrics.y_bottom : targetMetrics.y_top;
        diffY = Math.abs(bendY - tipY);
        diffX = diffY / Math.tan(angleRad);

        if (autoHDir === "left") {
            /* オブジェクトがテキストより左: 水平端をテキスト右端へ、先端はオブジェクト左 */
            horizontalEndX = textBounds.x_right;
            tipX = targetMetrics.x_left;
            bendX = tipX + diffX;
            return { points: [[tipX, tipY], [bendX, bendY], [horizontalEndX, bendY]], hDir: autoHDir, vDir: autoVDir };
        }
        /* オブジェクトがテキストより右: 水平端をテキスト左端へ、先端はオブジェクト右 */
        horizontalEndX = textBounds.x_left;
        tipX = targetMetrics.x_right;
        bendX = tipX - diffX;
        return { points: [[horizontalEndX, bendY], [bendX, bendY], [tipX, tipY]], hDir: autoHDir, vDir: autoVDir };
    }

    /**
     * 引き出し線の座標と、その座標が示す方向を求める
     * テキスト整列時は方向を自動判定するため、指定した方向とは異なる結果を返すことがある
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {string} vDir - 垂直方向（"up" / "down"）
     * @param {object} [textBounds] - テキストの外接矩形（テキスト整列時のみ）
     * @param {number} [textDistPt] - テキストとの距離（pt）
     * @returns {object} points（3点の座標配列）／hDir／vDir を持つオブジェクト
     */
    function calcLeaderGeometry(targetMetrics, angleRad, hDir, vDir, textBounds, textDistPt) {
        if (textBounds) {
            return calcLeaderPointsWithText(targetMetrics, angleRad, textBounds, textDistPt);
        }
        return { points: calcLeaderPoints(targetMetrics, angleRad, hDir, vDir), hDir: hDir, vDir: vDir };
    }

    /**
     * 斜線の先端座標を取得する
     * @param {Array<Array<number>>} points - 引き出し線の座標配列
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @returns {Array<number>} 先端の座標 [x, y]
     */
    function getTipPoint(points, hDir) {
        return (hDir === "left") ? points[0] : points[2];
    }

    /**
     * 斜線の先端を指定長さだけ短縮する
     * @param {Array<Array<number>>} points - 引き出し線の座標配列（破壊的に更新）
     * @param {string} hDir - 水平方向（"left" / "right"）
     * @param {number} length - 短縮する長さ（pt）
     * @returns {void}
     */
    function shortenTip(points, hDir, length) {
        var tipIndex = (hDir === "left") ? 0 : 2;
        var tip = points[tipIndex];
        var bend = points[1];
        var dx = bend[0] - tip[0];
        var dy = bend[1] - tip[1];
        var distance = Math.sqrt(dx * dx + dy * dy);
        if (distance > 0) {
            points[tipIndex] = [tip[0] + dx / distance * length, tip[1] + dy / distance * length];
        }
    }

    /**
     * 矢印の3頂点を計算する
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {Array<number>} bendPt - 折れ点の座標 [x, y]
     * @param {number} arrowSize - 矢印の長さ（pt）
     * @returns {Array<Array<number>>|null} 先端・翼端・翼端の座標配列（長さが0の場合は null）
     */
    function calcArrowPoints(tipPt, bendPt, arrowSize) {
        /* tipPt → bendPt 方向のベクトル */
        var dx = bendPt[0] - tipPt[0];
        var dy = bendPt[1] - tipPt[1];
        var distance = Math.sqrt(dx * dx + dy * dy);
        if (distance === 0) return null;
        var unitX = dx / distance;
        var unitY = dy / distance;
        /* 矢印の2つの翼端を計算 */
        var halfWidth = arrowSize / 2;
        var backX = tipPt[0] + unitX * arrowSize;
        var backY = tipPt[1] + unitY * arrowSize;
        return [
            [tipPt[0], tipPt[1]],
            [backX + unitY * halfWidth, backY - unitX * halfWidth],
            [backX - unitY * halfWidth, backY + unitX * halfWidth]
        ];
    }

    // =========================================
    // 引き出し線の生成 / Leader line drawing
    // =========================================

    /**
     * 先端に円を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} diameter - 円の直径（pt）
     * @param {boolean} fillOnly - true なら塗り、false なら線で描く
     * @returns {PathItem} 追加した円
     */
    function addTipCircleMarker(doc, tipPt, style, diameter, fillOnly) {
        var radius = diameter / 2;
        var circle = doc.pathItems.ellipse(
            tipPt[1] + radius, tipPt[0] - radius, diameter, diameter
        );
        if (fillOnly) {
            /* ●（塗りのみ） */
            circle.filled = true;
            circle.fillColor = style.lineColor;
            circle.stroked = false;
        } else {
            /* ○（線のみ） */
            circle.filled = false;
            circle.stroked = true;
            circle.strokeWidth = style.widthPt;
            circle.strokeColor = style.lineColor;
        }
        return circle;
    }

    /**
     * 先端の円にフチを追加する（●／○ とも一回り大きい塗り円を背面に置く）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} diameter - 本体の円の直径（pt）
     * @returns {PathItem} 追加したフチの円
     */
    function addTipCircleMarkerEdge(doc, tipPt, style, diameter) {
        var expandedDiameter = diameter + style.widthPt * EDGE_CIRCLE_EXPAND_RATIO;
        var expandedRadius = expandedDiameter / 2;
        var edgeCircle = doc.pathItems.ellipse(
            tipPt[1] + expandedRadius, tipPt[0] - expandedRadius, expandedDiameter, expandedDiameter
        );
        edgeCircle.filled = true;
        edgeCircle.fillColor = style.edgeColor;
        edgeCircle.stroked = false;
        return edgeCircle;
    }

    /**
     * 先端に矢印を追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {Array<number>} bendPt - 折れ点の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} arrowSize - 矢印の長さ（pt）
     * @param {boolean} fillOnly - true なら塗り、false なら線で描く
     * @returns {PathItem|null} 追加した矢印（長さが0の場合は null）
     */
    function addTipArrow(doc, tipPt, bendPt, style, arrowSize, fillOnly) {
        var arrowPoints = calcArrowPoints(tipPt, bendPt, arrowSize);
        if (!arrowPoints) return null;

        var arrow = doc.pathItems.add();
        arrow.setEntirePath(arrowPoints);
        arrow.closed = true;
        if (fillOnly) {
            arrow.filled = true;
            arrow.fillColor = style.lineColor;
            arrow.stroked = false;
        } else {
            arrow.filled = false;
            arrow.stroked = true;
            arrow.strokeWidth = style.widthPt;
            arrow.strokeColor = style.lineColor;
        }
        return arrow;
    }

    /**
     * 先端の矢印にフチを追加する（本体と同じ重心を基準にスケールアップ）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<number>} tipPt - 先端の座標 [x, y]
     * @param {Array<number>} bendPt - 折れ点の座標 [x, y]
     * @param {object} style - resolveLeaderStyle() の戻り値
     * @param {number} arrowSize - 本体の矢印の長さ（pt）
     * @returns {PathItem|null} 追加したフチの矢印（長さが0の場合は null）
     */
    function addTipArrowEdge(doc, tipPt, bendPt, style, arrowSize) {
        var arrowPoints = calcArrowPoints(tipPt, bendPt, arrowSize);
        if (!arrowPoints) return null;
        /* 重心を求める */
        var centerX = (arrowPoints[0][0] + arrowPoints[1][0] + arrowPoints[2][0]) / 3;
        var centerY = (arrowPoints[0][1] + arrowPoints[1][1] + arrowPoints[2][1]) / 3;
        /* 重心を基準にスケールアップ */
        var scale = (arrowSize + style.widthPt * EDGE_ARROW_EXPAND_RATIO) / arrowSize;
        var scaledPoints = [];
        for (var i = 0; i < arrowPoints.length; i++) {
            scaledPoints.push([
                centerX + (arrowPoints[i][0] - centerX) * scale,
                centerY + (arrowPoints[i][1] - centerY) * scale
            ]);
        }

        var edgeArrow = doc.pathItems.add();
        edgeArrow.setEntirePath(scaledPoints);
        edgeArrow.closed = true;
        edgeArrow.filled = true;
        edgeArrow.fillColor = style.edgeColor;
        edgeArrow.stroked = false;
        return edgeArrow;
    }

    /**
     * 引き出し線の各パーツをグループに収める
     * @param {GroupItem} container - 収める先のグループ
     * @param {PathItem} edgePath - フチの線
     * @param {PathItem} mainPath - 本体の線
     * @param {PathItem} edgeTipMarker - フチの線端
     * @param {PathItem} mainTipMarker - 本体の線端
     * @param {boolean} hasEdge - フチがあるかどうか
     * @returns {object} edgeGroup / mainGroup を持つオブジェクト
     */
    function assembleLeaderParts(container, edgePath, mainPath, edgeTipMarker, mainTipMarker, hasEdge) {
        var edgeGroup = null;
        var mainGroup = container;

        if (hasEdge) {
            edgeGroup = container.groupItems.add();
            mainGroup = container.groupItems.add();

            if (mainPath) mainPath.move(mainGroup, ElementPlacement.PLACEATEND);
            if (mainTipMarker) mainTipMarker.move(mainGroup, ElementPlacement.PLACEATEND);

            if (edgePath) edgePath.move(edgeGroup, ElementPlacement.PLACEATEND);
            if (edgeTipMarker) edgeTipMarker.move(edgeGroup, ElementPlacement.PLACEATEND);
        } else {
            if (mainPath) mainPath.move(container, ElementPlacement.PLACEATEND);
            if (mainTipMarker) mainTipMarker.move(container, ElementPlacement.PLACEATEND);
        }

        return {
            edgeGroup: edgeGroup,
            mainGroup: mainGroup
        };
    }

    /**
     * ダイアログから引き出し線の生成条件を読み取る
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {Array<TextFrame>} textFrameTargets - 選択中のテキストフレーム
     * @returns {object} 生成条件をまとめたオブジェクト
     */
    function readLeaderOptions(ui, textFrameTargets) {
        var diagDirection = getDiagDirValues(ui);

        var markerSizePt = unitValueToPt(ui.tipMarkerSizeInput.text, "strokeUnits");
        if (isNaN(markerSizePt) || markerSizePt <= 0) markerSizePt = DEFAULT_TIP_MARKER_SIZE_PT;

        /* テキストを選択しているときだけテキスト整列を使う */
        var textBounds = null;
        var textDistPt = 0;
        if (textFrameTargets.length > 0) {
            textBounds = getTextBounds(textFrameTargets[0]);
            textDistPt = parseFloat(ui.textDistInput.text);
            if (isNaN(textDistPt) || textDistPt < 0) textDistPt = DEFAULT_TEXT_DIST_PT;
        }

        var options = {
            hDir: diagDirection.hDir,
            vDir: diagDirection.vDir,
            useStoredDir: ui.applyScopeKeepDirRadio.value,
            useCircleMarker: ui.tipMarkerCircleRadio.value,
            useArrowMarker: ui.tipMarkerArrowRadio.value,
            fillTipMarker: ui.tipMarkerFillRadio.value,
            hasEdge: ui.edgeEnabledCheck.value,
            roundStrokeCap: ui.strokeCapRoundRadio.value,
            markerSizePt: markerSizePt,
            textBounds: textBounds,
            textDistPt: textDistPt,
            lineColorMode: getSelectedRadioKey(ui.radioMaps.lineColor, "black"),
            lineColorHex: ui.lineColorHexHolder.text,
            edgeColorMode: getSelectedRadioKey(ui.radioMaps.edgeColor, "white"),
            edgeColorHex: ui.edgeColorHexHolder.text,
            lineWidthPt: unitValueToPt(ui.lineWidthInput.text, "strokeUnits")
        };
        options.hasTipMarker = options.useCircleMarker || options.useArrowMarker;
        return options;
    }

    /**
     * 引き出し線1本分のパーツ（フチ・本体・線端）を生成する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem|GroupItem} targetItem - 元になったアイテム
     * @param {object} targetMetrics - measureTargets() の要素
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @param {object} options - readLeaderOptions() の戻り値
     * @param {Array<object>} [createdArtItems] - 生成したアイテムを追記する配列（後始末用）
     * @returns {object} 各パーツと、実際に使った方向を持つオブジェクト
     */
    function buildLeaderLineParts(doc, targetItem, targetMetrics, angleRad, options, createdArtItems) {
        createdArtItems = createdArtItems || [];

        /* 「斜線の方向以外」のとき、各オブジェクトの保存済み方向を使用 */
        var itemHDir = options.hDir;
        var itemVDir = options.vDir;
        if (options.useStoredDir) {
            var storedDir = getLeaderLineStoredDir(targetItem);
            if (storedDir) {
                itemHDir = storedDir.hDir;
                itemVDir = storedDir.vDir;
            }
        }

        var style = resolveLeaderStyle(doc, options, targetMetrics);
        /* テキスト整列時は方向が自動判定されるので、以降は geometry が返した方向を使う */
        var geometry = calcLeaderGeometry(targetMetrics, angleRad, itemHDir, itemVDir, options.textBounds, options.textDistPt);
        var points = geometry.points;

        /* 先端座標は短縮前に取得 */
        var tipPointBeforeShorten = options.hasTipMarker ? getTipPoint(points, geometry.hDir).slice(0) : null;
        var bendPt = options.useArrowMarker ? points[1].slice(0) : null;

        if (options.useCircleMarker && !options.fillTipMarker) {
            /* 円の半径 + 円の線幅の半分で短縮（線が円の縁に接する） */
            shortenTip(points, geometry.hDir, options.markerSizePt / 2 + style.widthPt / 2);
        }
        if (options.useArrowMarker) {
            /* 矢印の長さ分だけ短縮 */
            shortenTip(points, geometry.hDir, options.markerSizePt);
        }

        var parts = {
            edgePath: null,
            mainPath: null,
            edgeTipMarker: null,
            mainTipMarker: null,
            hDir: geometry.hDir,
            vDir: geometry.vDir
        };

        /* フチ（最背面） */
        if (options.hasEdge) {
            parts.edgePath = doc.pathItems.add();
            createdArtItems.push(parts.edgePath);
            parts.edgePath.setEntirePath(points);
            parts.edgePath.stroked = true;
            parts.edgePath.strokeWidth = style.widthPt * EDGE_LINE_WIDTH_RATIO;
            parts.edgePath.strokeColor = style.edgeColor;
            parts.edgePath.filled = false;
            parts.edgePath.strokeCap = options.roundStrokeCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
        }

        parts.mainPath = doc.pathItems.add();
        createdArtItems.push(parts.mainPath);
        parts.mainPath.setEntirePath(points);
        parts.mainPath.stroked = true;
        parts.mainPath.strokeWidth = style.widthPt;
        parts.mainPath.strokeColor = style.lineColor;
        parts.mainPath.filled = false;
        parts.mainPath.strokeCap = options.roundStrokeCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

        if (options.useCircleMarker) {
            if (options.hasEdge) parts.edgeTipMarker = addTipCircleMarkerEdge(doc, tipPointBeforeShorten, style, options.markerSizePt);
            parts.mainTipMarker = addTipCircleMarker(doc, tipPointBeforeShorten, style, options.markerSizePt, options.fillTipMarker);
        }
        if (options.useArrowMarker) {
            if (options.hasEdge) parts.edgeTipMarker = addTipArrowEdge(doc, tipPointBeforeShorten, bendPt, style, options.markerSizePt);
            parts.mainTipMarker = addTipArrow(doc, tipPointBeforeShorten, bendPt, style, options.markerSizePt, options.fillTipMarker);
        }
        if (parts.edgeTipMarker) createdArtItems.push(parts.edgeTipMarker);
        if (parts.mainTipMarker) createdArtItems.push(parts.mainTipMarker);

        return parts;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ∧∨と入力欄をひと組で追加する（隙間0で突き合わせ、↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} parentRow - 追加先の行
     * @param {number} characters - 入力欄の文字数
     * @param {object} stepOptions - addStepper() に渡す増減の設定。onStep は bindNumberInputEvents() で入れる
     * @returns {EditText} 入力欄（∧∨は .stepperGroup、増減の設定は .stepOptions で参照できる）
     */
    function addStepperInput(parentRow, characters, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, "");
        numberInput.characters = characters;
        numberInput.stepperGroup = stepperGroup;
        numberInput.stepOptions = stepOptions;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 適用範囲パネルを追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addApplyScopePanel(parent, ui) {
        ui.applyScopePanel = addLabeledPanel(parent, getLabel("panel.applyScope"));
        setupRowGroup(ui.applyScopePanel, ["center", "center"]);
        ui.applyScopeAllRadio = ui.applyScopePanel.add("radiobutton", undefined, getLabel("radio.applyScopeAll"));
        ui.applyScopeAllRadio.helpTip = getLabel("tooltip.applyScopeAll");
        ui.applyScopeKeepDirRadio = ui.applyScopePanel.add("radiobutton", undefined, getLabel("radio.applyScopeKeepDir"));
        ui.applyScopeKeepDirRadio.helpTip = getLabel("tooltip.applyScopeKeepDir");
    }

    /**
     * 角度パネルを追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addAnglePanel(parent, ui) {
        ui.anglePanel = addLabeledPanel(parent, getLabel("panel.angle"));

        ui.angleInputRow = ui.anglePanel.add("group");
        ui.angleInputRow.alignment = ["center", "top"];
        ui.angleInput = addStepperInput(ui.angleInputRow, 4, { step: 1, min: 0 });
        ui.angleInput.helpTip = getLabel("tooltip.angleInput");
        ui.angleInputRow.add("statictext", undefined, "\u00B0");
        ui.angleInput.active = true;

        ui.anglePresetRow = ui.anglePanel.add("group");
        setupRowGroup(ui.anglePresetRow);
        ui.anglePreset30Radio = ui.anglePresetRow.add("radiobutton", undefined, "30\u00B0");
        ui.anglePreset45Radio = ui.anglePresetRow.add("radiobutton", undefined, "45\u00B0");
        ui.anglePreset60Radio = ui.anglePresetRow.add("radiobutton", undefined, "60\u00B0");
    }

    /**
     * 斜線の方向パネルを追加する
     * 左右2列のラジオの間に、対象を示す中央列を挟む
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addDiagDirPanel(parent, ui) {
        ui.diagDirPanel = addLabeledPanel(parent, getLabel("panel.direction"));

        ui.diagDirRow = ui.diagDirPanel.add("group");
        setupRowGroup(ui.diagDirRow, ["fill", "top"]);
        ui.diagDirRow.helpTip = getLabel("tooltip.diagDir");

        ui.diagDirLeftColumn = ui.diagDirRow.add("group");
        ui.diagDirLeftColumn.orientation = "column";
        ui.diagDirLeftColumn.alignChildren = ["fill", "top"];
        ui.diagDirUpperLeftRadio = ui.diagDirLeftColumn.add("radiobutton", undefined, getLabel("radio.diagDirUpperLeft"));
        ui.diagDirLeftColumn.add("statictext", undefined, " ");
        ui.diagDirLowerLeftRadio = ui.diagDirLeftColumn.add("radiobutton", undefined, getLabel("radio.diagDirLowerLeft"));

        ui.diagDirCenterColumn = ui.diagDirRow.add("group");
        ui.diagDirCenterColumn.orientation = "column";
        ui.diagDirCenterColumn.alignChildren = ["center", "center"];
        ui.diagDirCenterColumn.add("statictext", undefined, " ");
        ui.diagDirCenterColumn.add("statictext", undefined, getLabel("fieldLabel.targetMark"));

        ui.diagDirRightColumn = ui.diagDirRow.add("group");
        ui.diagDirRightColumn.orientation = "column";
        ui.diagDirRightColumn.alignChildren = ["fill", "top"];
        ui.diagDirUpperRightRadio = ui.diagDirRightColumn.add("radiobutton", undefined, getLabel("radio.diagDirUpperRight"));
        ui.diagDirRightColumn.add("statictext", undefined, " ");
        ui.diagDirLowerRightRadio = ui.diagDirRightColumn.add("radiobutton", undefined, getLabel("radio.diagDirLowerRight"));
    }

    /**
     * 線のスタイルパネル（色・線幅・線端の形状）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @param {string} strokeUnitLabel - 線幅の単位ラベル
     * @returns {void}
     */
    function addLineStylePanel(parent, ui, strokeUnitLabel) {
        ui.lineStylePanel = addLabeledPanel(parent, getLabel("panel.lineStyle"));

        ui.lineColorRow = ui.lineStylePanel.add("group");
        setupRowGroup(ui.lineColorRow);
        ui.lineColorBlackRadio = ui.lineColorRow.add("radiobutton", undefined, getLabel("radio.lineColorBlack"));
        ui.lineColorWhiteRadio = ui.lineColorRow.add("radiobutton", undefined, getLabel("radio.lineColorWhite"));
        ui.lineColorCustomRadio = ui.lineColorRow.add("radiobutton", undefined, getLabel("radio.colorCustom"));
        /* HEX値はコントロールではなく、text プロパティだけを持つ保持箱に入れる */
        ui.lineColorHexHolder = { text: sessionSettings.lineColorHex };
        ui.lineColorChip = ui.lineColorRow.add("panel", undefined, "");
        ui.lineColorChip.preferredSize = COLOR_CHIP_SIZE;
        ui.lineColorChip.helpTip = getLabel("tooltip.colorChip");

        ui.lineWidthRow = ui.lineStylePanel.add("group");
        setupRowGroup(ui.lineWidthRow);
        ui.lineWidthRow.add("statictext", undefined, getLabel("fieldLabel.lineWidth"));
        ui.lineWidthInput = addStepperInput(ui.lineWidthRow, 3, { step: 0.1, min: 0 });
        ui.lineWidthInput.helpTip = getLabel("tooltip.lengthInput");
        ui.lineWidthRow.add("statictext", undefined, strokeUnitLabel);

        ui.strokeCapRow = ui.lineStylePanel.add("group");
        setupRowGroup(ui.strokeCapRow);
        ui.strokeCapRow.helpTip = getLabel("tooltip.strokeCap");
        ui.strokeCapRow.add("statictext", undefined, getLabel("fieldLabel.strokeCap"));
        ui.strokeCapButtRadio = ui.strokeCapRow.add("radiobutton", undefined, getLabel("radio.strokeCapButt"));
        ui.strokeCapRoundRadio = ui.strokeCapRow.add("radiobutton", undefined, getLabel("radio.strokeCapRound"));
    }

    /**
     * 線端パネル（円・矢印の種類と大きさ）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @param {string} strokeUnitLabel - 大きさの単位ラベル
     * @returns {void}
     */
    function addTipMarkerPanel(parent, ui, strokeUnitLabel) {
        ui.tipMarkerPanel = addLabeledPanel(parent, getLabel("panel.tipMarker"));

        ui.tipMarkerTypeRow = ui.tipMarkerPanel.add("group");
        setupRowGroup(ui.tipMarkerTypeRow);
        ui.tipMarkerNoneRadio = ui.tipMarkerTypeRow.add("radiobutton", undefined, getLabel("radio.tipMarkerNone"));
        ui.tipMarkerCircleRadio = ui.tipMarkerTypeRow.add("radiobutton", undefined, getLabel("radio.tipMarkerCircle"));
        ui.tipMarkerArrowRadio = ui.tipMarkerTypeRow.add("radiobutton", undefined, getLabel("radio.tipMarkerArrow"));

        ui.tipMarkerStyleRow = ui.tipMarkerPanel.add("group");
        setupRowGroup(ui.tipMarkerStyleRow);
        ui.tipMarkerStyleRow.helpTip = getLabel("tooltip.tipMarkerStyle");
        ui.tipMarkerFillRadio = ui.tipMarkerStyleRow.add("radiobutton", undefined, getLabel("radio.tipMarkerFill"));
        ui.tipMarkerOutlineRadio = ui.tipMarkerStyleRow.add("radiobutton", undefined, getLabel("radio.tipMarkerOutline"));

        ui.tipMarkerSizeRow = ui.tipMarkerPanel.add("group");
        setupRowGroup(ui.tipMarkerSizeRow);
        ui.tipMarkerSizeRow.add("statictext", undefined, getLabel("fieldLabel.tipMarkerSize"));
        ui.tipMarkerSizeInput = addStepperInput(ui.tipMarkerSizeRow, 3, { step: 0.1, min: 0 });
        ui.tipMarkerSizeInput.helpTip = getLabel("tooltip.tipMarkerSize");
        ui.tipMarkerSizeRow.add("statictext", undefined, strokeUnitLabel);

        ui.groupItemsCheck = ui.tipMarkerPanel.add("checkbox", undefined, getLabel("checkbox.groupEnabled"));
        ui.groupItemsCheck.alignment = "left";
        ui.groupItemsCheck.helpTip = getLabel("tooltip.groupEnabled");
    }

    /**
     * テキストパネル（テキストとの距離）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addTextAlignPanel(parent, ui) {
        ui.textAlignPanel = addLabeledPanel(parent, getLabel("panel.text"));

        ui.textDistRow = ui.textAlignPanel.add("group");
        setupRowGroup(ui.textDistRow);
        ui.textDistRow.add("statictext", undefined, getLabel("fieldLabel.textDist"));
        ui.textDistInput = addStepperInput(ui.textDistRow, 4, { step: 0.1, min: 0 });
        ui.textDistInput.helpTip = getLabel("tooltip.textDist");
        ui.textDistRow.add("statictext", undefined, "pt");
    }

    /**
     * フチパネル（フチの有無と色）を追加する
     * @param {Window|Group|Panel} parent - 追加先
     * @param {object} ui - コントロールを格納するオブジェクト
     * @returns {void}
     */
    function addEdgePanel(parent, ui) {
        ui.edgePanel = addLabeledPanel(parent, getLabel("panel.edge"));

        ui.edgeEnabledCheck = ui.edgePanel.add("checkbox", undefined, getLabel("checkbox.edgeEnabled"));
        ui.edgeEnabledCheck.alignment = "left";
        ui.edgeEnabledCheck.helpTip = getLabel("tooltip.edgeEnabled");

        ui.edgeColorRow = ui.edgePanel.add("group");
        setupRowGroup(ui.edgeColorRow);
        ui.edgeColorWhiteRadio = ui.edgeColorRow.add("radiobutton", undefined, getLabel("radio.lineColorWhite"));
        ui.edgeColorCustomRadio = ui.edgeColorRow.add("radiobutton", undefined, getLabel("radio.colorCustom"));
        ui.edgeColorHexHolder = { text: sessionSettings.edgeColorHex };
        ui.edgeColorChip = ui.edgeColorRow.add("panel", undefined, "");
        ui.edgeColorChip.preferredSize = COLOR_CHIP_SIZE;
        ui.edgeColorChip.helpTip = getLabel("tooltip.colorChip");
    }

    /**
     * 設定キーとラジオボタンの対応表を作る
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {object} カテゴリごとの対応表
     */
    function buildRadioMaps(ui) {
        return {
            presetAngle: [
                { key: "30", radio: ui.anglePreset30Radio },
                { key: "45", radio: ui.anglePreset45Radio },
                { key: "60", radio: ui.anglePreset60Radio }
            ],
            applyScope: [
                { key: "all",             radio: ui.applyScopeAllRadio },
                { key: "exceptDirection", radio: ui.applyScopeKeepDirRadio }
            ],
            diagDir: [
                { key: "upperLeft",  radio: ui.diagDirUpperLeftRadio },
                { key: "lowerLeft",  radio: ui.diagDirLowerLeftRadio },
                { key: "upperRight", radio: ui.diagDirUpperRightRadio },
                { key: "lowerRight", radio: ui.diagDirLowerRightRadio }
            ],
            lineColor: [
                { key: "black", radio: ui.lineColorBlackRadio },
                { key: "white", radio: ui.lineColorWhiteRadio },
                { key: "other", radio: ui.lineColorCustomRadio }
            ],
            strokeCap: [
                { key: "none",  radio: ui.strokeCapButtRadio },
                { key: "round", radio: ui.strokeCapRoundRadio }
            ],
            tipMarkerType: [
                { key: "none",   radio: ui.tipMarkerNoneRadio },
                { key: "circle", radio: ui.tipMarkerCircleRadio },
                { key: "arrow",  radio: ui.tipMarkerArrowRadio }
            ],
            tipMarkerStyle: [
                { key: "fill",   radio: ui.tipMarkerFillRadio },
                { key: "stroke", radio: ui.tipMarkerOutlineRadio }
            ],
            edgeColor: [
                { key: "white", radio: ui.edgeColorWhiteRadio },
                { key: "other", radio: ui.edgeColorCustomRadio }
            ]
        };
    }

    /**
     * ダイアログを組み立てる
     * @returns {object} 各コントロールをまとめたオブジェクト
     */
    function buildDialogUI() {
        var ui = {};
        var strokeUnitLabel = getUnitInfo("strokeUnits").label;

        ui.dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(ui.dialogWindow);

        addApplyScopePanel(ui.dialogWindow, ui);

        ui.columnsRow = ui.dialogWindow.add("group");
        setupRowGroup(ui.columnsRow, ["fill", "top"], COLUMN_SPACING);

        ui.leftColumn = ui.columnsRow.add("group");
        ui.leftColumn.orientation = "column";
        ui.leftColumn.alignChildren = ["fill", "top"];

        addAnglePanel(ui.leftColumn, ui);
        addDiagDirPanel(ui.leftColumn, ui);

        ui.rightColumn = ui.columnsRow.add("group");
        ui.rightColumn.orientation = "column";
        ui.rightColumn.alignChildren = ["fill", "top"];

        addLineStylePanel(ui.rightColumn, ui, strokeUnitLabel);
        addTipMarkerPanel(ui.rightColumn, ui, strokeUnitLabel);
        addTextAlignPanel(ui.rightColumn, ui);
        addEdgePanel(ui.leftColumn, ui);

        ui.radioMaps = buildRadioMaps(ui);

        return ui;
    }

    /**
     * 前回の設定をダイアログへ反映する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {Array<PathItem|GroupItem>} targetItems - 対象アイテム
     * @param {Array<TextFrame>} textFrameTargets - 選択中のテキストフレーム
     * @returns {void}
     */
    function loadDialogSettingsToUI(ui, targetItems, textFrameTargets) {
        ui.angleInput.text = String(sessionSettings.angle);
        selectRadioByKey(ui.radioMaps.presetAngle, String(sessionSettings.radioAngle), "45");

        /* 複数選択で未設定のときは「斜線の方向以外」を初期選択にする */
        var applyScope = sessionSettings.applyScope;
        if (!sessionSettings.hasUserSetApplyScope && targetItems.length > 1) applyScope = "exceptDirection";
        selectRadioByKey(ui.radioMaps.applyScope, applyScope, "all");

        selectRadioByKey(ui.radioMaps.diagDir, sessionSettings.diagDir, "upperLeft");
        selectRadioByKey(ui.radioMaps.lineColor, sessionSettings.lineColor, "black");
        selectRadioByKey(ui.radioMaps.strokeCap, sessionSettings.strokeCapType, "round");
        selectRadioByKey(ui.radioMaps.tipMarkerType, sessionSettings.tipMarkerType, "none");
        selectRadioByKey(ui.radioMaps.tipMarkerStyle, sessionSettings.tipMarkerStyle, "fill");
        selectRadioByKey(ui.radioMaps.edgeColor, sessionSettings.edgeColor, "white");

        ui.lineColorHexHolder.text = sessionSettings.lineColorHex;
        updateColorChip(ui.lineColorChip, ui.lineColorHexHolder.text);
        ui.edgeColorHexHolder.text = sessionSettings.edgeColorHex;
        updateColorChip(ui.edgeColorChip, ui.edgeColorHexHolder.text);

        ui.lineWidthInput.text = formatUnitInput(sessionSettings.lineWidth, DEFAULT_LINE_WIDTH_PT);
        ui.tipMarkerSizeInput.text = formatUnitInput(sessionSettings.tipMarkerSize, DEFAULT_TIP_MARKER_SIZE_PT);

        ui.groupItemsCheck.value = !!sessionSettings.groupEnabled;
        ui.edgeEnabledCheck.value = !!sessionSettings.whiteEdge;

        var textDistValue = parseFloat(sessionSettings.textDist);
        if (isNaN(textDistValue) || textDistValue < 0) textDistValue = DEFAULT_TEXT_DIST_PT;
        ui.textDistInput.text = formatNumber(textDistValue, 2);

        ui.textAlignPanel.enabled = textFrameTargets.length > 0;
        redrawSteppersIn(ui.textAlignPanel);

        updateDirPanelEnabled(ui);
        updateTipMarkerControlsEnabled(ui);
        updateEdgeColorEnabled(ui);
    }

    /**
     * ダイアログの設定をセッション記憶へ保存する
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function saveDialogSettingsFromUI(ui) {
        var diagDirection = getDiagDirValues(ui);

        sessionSettings.angle = ui.angleInput.text;
        sessionSettings.radioAngle = parseFloat(getSelectedRadioKey(ui.radioMaps.presetAngle, "45"));
        sessionSettings.applyScope = getSelectedRadioKey(ui.radioMaps.applyScope, "all");
        sessionSettings.diagDir = getSelectedRadioKey(ui.radioMaps.diagDir, "lowerRight");
        sessionSettings.hDir = diagDirection.hDir;
        sessionSettings.vDir = diagDirection.vDir;

        sessionSettings.tipMarkerType = getSelectedRadioKey(ui.radioMaps.tipMarkerType, "none");
        sessionSettings.tipMarkerStyle = getSelectedRadioKey(ui.radioMaps.tipMarkerStyle, "fill");
        sessionSettings.tipMarkerSize = parseUnitInput(ui.tipMarkerSizeInput.text, DEFAULT_TIP_MARKER_SIZE_PT);
        sessionSettings.strokeCapType = getSelectedRadioKey(ui.radioMaps.strokeCap, "round");

        sessionSettings.groupEnabled = ui.groupItemsCheck.value;
        sessionSettings.whiteEdge = ui.edgeEnabledCheck.value;
        sessionSettings.edgeColor = getSelectedRadioKey(ui.radioMaps.edgeColor, "white");
        sessionSettings.edgeColorHex = ui.edgeColorHexHolder.text;

        sessionSettings.lineColor = getSelectedRadioKey(ui.radioMaps.lineColor, "black");
        sessionSettings.lineColorHex = ui.lineColorHexHolder.text;
        sessionSettings.lineWidth = parseUnitInput(ui.lineWidthInput.text, DEFAULT_LINE_WIDTH_PT);

        var savedTextDist = parseFloat(ui.textDistInput.text);
        sessionSettings.textDist = (!isNaN(savedTextDist) && savedTextDist >= 0) ? formatNumber(savedTextDist, 4) : DEFAULT_TEXT_DIST_PT;
        settingsStore.save(sessionSettings);
    }

    /**
     * 適用範囲に応じて斜線の方向パネルの有効・無効を切り替える
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function updateDirPanelEnabled(ui) {
        ui.diagDirPanel.enabled = ui.applyScopeAllRadio.value;
    }

    /**
     * 線端の選択に応じて関連コントロールの有効・無効を切り替える
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function updateTipMarkerControlsEnabled(ui) {
        var hasCap = ui.tipMarkerCircleRadio.value || ui.tipMarkerArrowRadio.value;
        if (ui.tipMarkerArrowRadio.value) {
            /* 矢印は塗りのみ */
            if (ui.tipMarkerOutlineRadio.value) {
                ui.tipMarkerOutlineRadio.value = false;
                ui.tipMarkerFillRadio.value = true;
            }
            ui.tipMarkerFillRadio.enabled = false;
            ui.tipMarkerOutlineRadio.enabled = false;
        } else {
            ui.tipMarkerFillRadio.enabled = hasCap;
            ui.tipMarkerOutlineRadio.enabled = hasCap;
        }
        ui.tipMarkerSizeInput.enabled = hasCap;
        ui.tipMarkerSizeInput.stepperGroup.enabled = hasCap;
        redrawSteppersIn(ui.tipMarkerSizeInput.stepperGroup);
        /* 線端がなくても、フチONなら本体とフチの2本になるためグループ化が意味を持つ */
        ui.groupItemsCheck.enabled = hasCap || ui.edgeEnabledCheck.value;
    }

    /**
     * フチのON/OFFに応じてフチ色のコントロールの有効・無効を切り替える
     * @param {object} ui - buildDialogUI() の戻り値
     * @returns {void}
     */
    function updateEdgeColorEnabled(ui) {
        var edgeEnabled = ui.edgeEnabledCheck.value;
        ui.edgeColorRow.enabled = edgeEnabled;
        ui.edgeColorChip.enabled = edgeEnabled;
    }

    /**
     * カラーピッカーを開き、選ばれた色をHEX保持欄とスウォッチに反映する
     * @param {object} hexHolder - HEX値を保持するオブジェクト
     * @param {Panel} colorChip - 対象のスウォッチ
     * @returns {boolean} 色が選ばれたら true
     */
    function pickColor(hexHolder, colorChip) {
        var pickedHex = ColorPicker.show(hexHolder.text.replace(/^#/, ""));
        if (!pickedHex) return false;
        hexHolder.text = "#" + pickedHex;
        updateColorChip(colorChip, hexHolder.text);
        return true;
    }

    /**
     * 数値入力欄のイベントを登録する（∧∨・↑↓キーと直接入力）
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindNumberInputEvents(ui, onChange) {
        var afterChange = function () { onChange(); };

        var numberInputs = [ui.angleInput, ui.lineWidthInput, ui.tipMarkerSizeInput, ui.textDistInput];
        for (var i = 0; i < numberInputs.length; i++) {
            numberInputs[i].stepOptions.onStep = afterChange;
            numberInputs[i].onChanging = afterChange;
        }

        ui.anglePreset30Radio.onClick = function () { ui.angleInput.text = "30"; onChange(); };
        ui.anglePreset45Radio.onClick = function () { ui.angleInput.text = "45"; onChange(); };
        ui.anglePreset60Radio.onClick = function () { ui.angleInput.text = "60"; onChange(); };
    }

    /**
     * 適用範囲と斜線の方向のイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindDiagDirEvents(ui, onChange) {
        /* 適用範囲はユーザーが明示的にクリックしたときだけ「手動設定済み」とみなす */
        var onScopeClick = function () {
            sessionSettings.hasUserSetApplyScope = true;
            updateDirPanelEnabled(ui);
            onChange();
        };
        ui.applyScopeAllRadio.onClick = onScopeClick;
        ui.applyScopeKeepDirRadio.onClick = onScopeClick;

        /* 方向ラジオは列をまたぐため自動排他が効かない。クリック時に自分だけを選択状態にする */
        for (var i = 0; i < ui.radioMaps.diagDir.length; i++) {
            (function (entry) {
                entry.radio.onClick = function () {
                    selectRadioByKey(ui.radioMaps.diagDir, entry.key, entry.key);
                    onChange();
                };
            })(ui.radioMaps.diagDir[i]);
        }
    }

    /**
     * 線のスタイル（色と線端の形状）のイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindLineStyleEvents(ui, onChange) {
        ui.lineColorChip.addEventListener("click", function () {
            if (!pickColor(ui.lineColorHexHolder, ui.lineColorChip)) return;
            selectRadioByKey(ui.radioMaps.lineColor, "other", "other");
            onChange();
        });
        ui.lineColorBlackRadio.onClick = function () { onChange(); };
        ui.lineColorWhiteRadio.onClick = function () { onChange(); };
        ui.lineColorCustomRadio.onClick = function () {
            pickColor(ui.lineColorHexHolder, ui.lineColorChip);
            onChange();
        };

        ui.strokeCapButtRadio.onClick = function () { onChange(); };
        ui.strokeCapRoundRadio.onClick = function () { onChange(); };
    }

    /**
     * 線端のイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindTipMarkerEvents(ui, onChange) {
        var onTypeClick = function () { updateTipMarkerControlsEnabled(ui); onChange(); };
        ui.tipMarkerNoneRadio.onClick = onTypeClick;
        ui.tipMarkerCircleRadio.onClick = onTypeClick;
        ui.tipMarkerArrowRadio.onClick = onTypeClick;

        ui.tipMarkerFillRadio.onClick = function () { onChange(); };
        ui.tipMarkerOutlineRadio.onClick = function () { onChange(); };
        ui.groupItemsCheck.onClick = function () { saveDialogSettingsFromUI(ui); };
    }

    /**
     * フチのイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数
     * @returns {void}
     */
    function bindEdgeEvents(ui, onChange) {
        ui.edgeEnabledCheck.onClick = function () {
            updateEdgeColorEnabled(ui);
            updateTipMarkerControlsEnabled(ui);
            onChange();
        };
        ui.edgeColorChip.addEventListener("click", function () {
            if (!pickColor(ui.edgeColorHexHolder, ui.edgeColorChip)) return;
            selectRadioByKey(ui.radioMaps.edgeColor, "other", "other");
            onChange();
        });
        ui.edgeColorWhiteRadio.onClick = function () { onChange(); };
        ui.edgeColorCustomRadio.onClick = function () {
            pickColor(ui.edgeColorHexHolder, ui.edgeColorChip);
            onChange();
        };
    }

    /**
     * ダイアログのイベントを登録する
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {function} onChange - 設定が変わったときに呼ぶ関数（プレビュー更新）
     * @returns {void}
     */
    function bindDialogEvents(ui, onChange) {
        bindNumberInputEvents(ui, onChange);
        bindDiagDirEvents(ui, onChange);
        bindLineStyleEvents(ui, onChange);
        bindTipMarkerEvents(ui, onChange);
        bindEdgeEvents(ui, onChange);

        ui.dialogWindow.onClose = function () {
            saveDialogSettingsFromUI(ui);
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から引き出し線の対象とテキストフレームを集める
     * PathItem は2点以上、GroupItem はそのまま対象とする
     * @param {object} selectedItems - ドキュメントの選択
     * @returns {object} leaderTargets / textTargets を持つオブジェクト
     */
    function collectTargets(selectedItems) {
        var leaderTargets = [];
        var textTargets = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var itemType = selectedItems[i].typename;
            if (itemType === "PathItem" && selectedItems[i].pathPoints.length >= 2) {
                leaderTargets.push(selectedItems[i]);
            } else if (itemType === "GroupItem") {
                leaderTargets.push(selectedItems[i]);
            } else if (itemType === "TextFrame") {
                textTargets.push(selectedItems[i]);
            }
        }

        return { leaderTargets: leaderTargets, textTargets: textTargets };
    }

    /**
     * 確定した設定で引き出し線を生成し、元オブジェクトと置き換える
     * @param {Document} doc - 対象ドキュメント
     * @param {object} ui - buildDialogUI() の戻り値
     * @param {Array<PathItem|GroupItem>} targetItems - 対象アイテム
     * @param {Array<TextFrame>} textFrameTargets - 選択中のテキストフレーム
     * @param {number} angleRad - 斜線の角度（ラジアン）
     * @returns {void}
     */
    function applyLeaderLines(doc, ui, targetItems, textFrameTargets, angleRad) {
        var leaderOptions = readLeaderOptions(ui, textFrameTargets);
        var targetMetricsList = measureTargets(targetItems);
        var itemsToSelect = [];
        var i;
        doc.selection = null;

        for (i = 0; i < targetItems.length; i++) {
            var sourceTarget = targetItems[i];
            var targetMetrics = targetMetricsList[i];
            var createdArtItems = [];
            switchToEditableLayer(doc, sourceTarget);

            try {
                var parts = buildLeaderLineParts(doc, sourceTarget, targetMetrics, angleRad, leaderOptions, createdArtItems);

                if (ui.groupItemsCheck.value) {
                    var leaderGroup = doc.groupItems.add();
                    createdArtItems.push(leaderGroup);
                    var assembled = assembleLeaderParts(leaderGroup, parts.edgePath, parts.mainPath, parts.edgeTipMarker, parts.mainTipMarker, leaderOptions.hasEdge);
                    tagLeaderParts(assembled, parts.edgePath, parts.mainPath, parts.edgeTipMarker, parts.mainTipMarker, leaderOptions.hasEdge);
                    setLeaderLineTag(leaderGroup);
                    setLeaderLineBoundsTags(leaderGroup, targetMetrics);
                    setLeaderLineDirTags(leaderGroup, parts.hDir, parts.vDir);
                    itemsToSelect.push(leaderGroup);
                } else {
                    if (parts.edgePath) itemsToSelect.push(parts.edgePath);
                    if (parts.edgeTipMarker) itemsToSelect.push(parts.edgeTipMarker);
                    if (parts.mainPath) itemsToSelect.push(parts.mainPath);
                    if (parts.mainTipMarker) itemsToSelect.push(parts.mainTipMarker);
                }

                /* 置き換え成功時のみ元オブジェクトを削除 */
                sourceTarget.remove();
            } catch (err) {
                for (var k = createdArtItems.length - 1; k >= 0; k--) {
                    safeRemove(createdArtItems[k]);
                }
                throw err;
            }
        }

        for (i = 0; i < itemsToSelect.length; i++) {
            safeSelect(itemsToSelect[i]);
        }
    }

    /**
     * 選択オブジェクトから引き出し線を作成する
     * @returns {void}
     */
    function main() {
        /* ドキュメントが開かれているか確認 / Check for an open document */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        /* 前回の設定を読み込む / Load the previous settings */
        sessionSettings = settingsStore.load(DEFAULT_SETTINGS);

        /* 選択オブジェクトのチェック / Check the selection */
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var targets = collectTargets(doc.selection);
        var targetItems = targets.leaderTargets;
        var textFrameTargets = targets.textTargets;

        if (targetItems.length === 0) {
            alert(getLabel("alert.noValidTargets"));
            return;
        }

        /* プレビュー用アイテムの配列（親グループ単位で管理） */
        var previewGroups = [];

        /**
         * プレビューを生成する
         * @param {number} angleDeg - 斜線の角度（度）
         * @param {object} ui - buildDialogUI() の戻り値
         * @returns {void}
         */
        function createPreview(angleDeg, ui) {
            var angleRad = angleDeg * Math.PI / 180;
            var targetMetricsList = measureTargets(targetItems);
            var options = readLeaderOptions(ui, textFrameTargets);

            for (var i = 0; i < targetMetricsList.length; i++) {
                switchToEditableLayer(doc, targetItems[i]);
                var parts = buildLeaderLineParts(doc, targetItems[i], targetMetricsList[i], angleRad, options);
                var previewGroup = doc.groupItems.add();
                assembleLeaderParts(previewGroup, parts.edgePath, parts.mainPath, parts.edgeTipMarker, parts.mainTipMarker, options.hasEdge);
                previewGroups.push(previewGroup);
            }
        }

        /**
         * プレビューを削除する
         * @returns {void}
         */
        function removePreview() {
            for (var i = previewGroups.length - 1; i >= 0; i--) {
                safeRemove(previewGroups[i]);
            }
            previewGroups = [];
        }

        /**
         * 対象オブジェクトの表示・非表示を切り替える
         * @param {boolean} visible - 表示するなら true
         * @returns {void}
         */
        function setTargetsVisible(visible) {
            for (var i = 0; i < targetItems.length; i++) {
                /* ロックされた対象では失敗するが、プレビューは続行する */
                try {
                    targetItems[i].hidden = !visible;
                } catch (e) { }
            }
        }

        var initialViewState = captureViewState(doc);

        /**
         * 入力値でプレビューを作り直す
         * @param {object} ui - buildDialogUI() の戻り値
         * @returns {void}
         */
        function updatePreview(ui) {
            removePreview();
            setTargetsVisible(true);

            var angleDeg = parseFloat(ui.angleInput.text);
            if (isNaN(angleDeg) || angleDeg <= 0 || angleDeg >= 90) {
                saveDialogSettingsFromUI(ui);
                app.redraw();
                return;
            }

            createPreview(angleDeg, ui);
            setTargetsVisible(false);
            saveDialogSettingsFromUI(ui);
            app.redraw();
        }

        var dialogUI = buildDialogUI();
        bindDialogEvents(dialogUI, function () { updatePreview(dialogUI); });
        restoreDirFromSelection(targetItems);
        loadDialogSettingsToUI(dialogUI, targetItems, textFrameTargets);

        var zoomControls = addZoomControls(dialogUI.dialogWindow, doc, getLabel("fieldLabel.zoom"), initialViewState, {
            min: ZOOM_MIN,
            max: ZOOM_MAX,
            sliderWidth: ZOOM_SLIDER_WIDTH,
            margins: [0, 0, 0, 10],
            redraw: true,
            lightMode: true,
            lightModeLabel: getLabel("checkbox.lightMode"),
            lightModeTip: getLabel("tooltip.lightMode"),
            lightModeDefault: false
        });
        /* ボタン行（左右中央） / Button row (centered) */
        var buttonRow = addButtonRow(dialogUI.dialogWindow, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        updatePreview(dialogUI);

        prepareDialogWindow(dialogUI.dialogWindow, SCRIPT_NAME);
        var dialogResult = dialogUI.dialogWindow.show();
        saveDialogSettingsFromUI(dialogUI);

        /* プレビューを削除して元のパスを復元 */
        removePreview();
        setTargetsVisible(true);
        app.redraw();

        if (dialogResult !== 1) {
            zoomControls.restoreInitial();
            return;
        }

        var angleDeg = parseFloat(dialogUI.angleInput.text);
        if (isNaN(angleDeg) || angleDeg <= 0 || angleDeg >= 90) {
            alert(getLabel("alert.invalidAngle"));
            app.redraw();
            return;
        }
        /* 確定：引き出し線を生成 / Build the leader lines */
        applyLeaderLines(doc, dialogUI, targetItems, textFrameTargets, angleDeg * Math.PI / 180);
    }

    main();

}());
