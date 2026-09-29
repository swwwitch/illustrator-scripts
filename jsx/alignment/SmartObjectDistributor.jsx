#target illustrator
#targetengine "SmartObjectDistributorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、指定した行数・列数のグリッドに沿って各セルの中央へ配置します。
配置先は「現在のアートボード」「最背面のオブジェクト」「_target レイヤーの長方形」から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectDistributor.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na3c45cea09b7

### Overview

Places the selected objects at the center of each cell of a grid with the rows and columns you specify.
The target area can be the current artboard, the backmost object, or a rectangle on the `_target` layer.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectDistributor.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectDistributor";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.10.3";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectDistributor.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectDistributor.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na3c45cea09b7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var CONFIG = {
        defaultGutter: 10,           /* 間隔の初期値（定規単位） / default gutter */
        defaultMargin: 10,           /* マージンの初期値（定規単位） / default margin */
        blackCellOpacity: 15,        /* 黒セルの不透明度（%） / opacity for black cells */
        whiteCellOpacity: 100,       /* 白セルの不透明度（%） / opacity for white cells */
        fallbackDivision: 5,         /* 分割数を決められないときの既定値 / fallback division */
        maxDivision: 100,            /* 行・列の上限 / max rows or columns */
        maxCellCount: 1000,          /* セル総数の上限 / max total cells */
        previewThrottleCells: 200,   /* この数を超えるセルでは入力中のプレビューを間引く / throttle preview above this many cells */
        previewThrottleMs: 200,      /* 間引く間隔（ミリ秒） / throttle window in milliseconds */
        parkingStep: 200,            /* 退避時の移動量（pt） / step when parking objects */
        overlapTolerance: 1,         /* 重なり判定の余裕（pt） / overlap tolerance */
        artboardBuffer: 10,          /* 他アートボードとの余裕（pt） / buffer around artboards */
        targetLayerName: "_target",                  /* 配置先レイヤー名 / target layer name */
        cellLayerName: "cell-background",            /* セル描画レイヤー名 / cell layer name */
        previewCellLayerName: "_Preview_Background", /* プレビュー用レイヤー名 / preview layer name */
        legacyPreviewLayerName: "_Preview_Guides"    /* 旧版のプレビュー用レイヤー名 / legacy preview layer */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /**
     * ウィンドウに共通のレイアウト設定を適用します。/ Apply shared window layout.
     *
     * @param {Window} targetWindow - 対象のウィンドウ。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウト設定を適用します。/ Apply shared panel layout.
     *
     * @param {Panel} targetPanel - 対象のパネル。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
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
     * 横並びの行グループを設定します（ボタン列など）。/ Apply a horizontal row group.
     *
     * @param {Group} targetGroup - 対象のグループ。
     * @param {string} [alignment] - 親に対する揃え。省略時は "left"。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = alignment || "left";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定 px 詰めます（レイアウト確定後に呼びます）。/ Trim a button's height.
     *
     * @param {Button} targetButton - 対象のボタン。
     * @param {number} trimPx - 詰める高さ（px）。
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPx) {
        try {
            targetButton.size = [targetButton.size.width, targetButton.size.height - trimPx];
        } catch (e) { /* レイアウト前は size が無いことがある / size may be missing before layout */ }
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
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

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 日英のラベル文言。
     *
     * @typedef {Object} LabelEntry
     * @property {string} ja - 日本語の文言。
     * @property {string} en - 英語の文言。
     */

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
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

    var LABELS = {
        dialog: {
            title: { ja: "グリッドに整列配置", en: "Arrange in Grid" }
        },
        panel: {
            placement: { ja: "配置先", en: "Placement Area" },
            division: { ja: "分割とマージン", en: "Divisions & Margin" },
            cellDrawing: { ja: "セルの扱い", en: "Cell Handling" }
        },
        fieldLabel: {
            rowCount: { ja: "行数", en: "Rows" },
            columnCount: { ja: "列数", en: "Columns" },
            gutter: { ja: "セル間隔", en: "Gutter" },
            margin: { ja: "マージン", en: "Margin" },
            cellColor: { ja: "カラー", en: "Color" },
            cellOpacity: { ja: "不透明度", en: "Opacity" }
        },
        radio: {
            targetArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            targetBackmost: { ja: "最背面のオブジェクト", en: "Backmost Object" },
            targetRectLayer: { ja: "「_target」レイヤーの長方形", en: "Rectangle in '_target' Layer" },
            keepCell: { ja: "長方形で残す", en: "Keep as Rectangle" },
            toGuide: { ja: "ガイド化", en: "Convert to Guides" },
            toArtboard: { ja: "アートボード化", en: "Convert to Artboards" },
            blackCell: { ja: "黒", en: "Black" },
            whiteCell: { ja: "白", en: "White" },
            transparentCell: { ja: "塗りなし", en: "No Fill" }
        },
        button: {
            transparencyGrid: { ja: "透明グリッド表示", en: "Transparency Grid" },
            randomize: { ja: "シャッフル", en: "Shuffle" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            targetUnavailable: { ja: "該当するオブジェクトがありません。", en: "No matching object was found." },
            targetBackmost: {
                ja: "表示されていてロックされていないオブジェクトのうち、最も背面にあるものを配置先にします。そのオブジェクト自体は配置しません。",
                en: "Uses the backmost visible, unlocked object as the area. That object itself is not placed."
            },
            targetRectLayer: {
                ja: "この長方形は処理中だけ非表示になり、配置対象には含まれません。",
                en: "This rectangle is hidden while the script runs and is never placed into a cell."
            },
            keepCell: {
                ja: "セルの長方形を「cell-background」レイヤーに残します。",
                en: "Keeps the cell rectangles on the \"cell-background\" layer."
            },
            toGuide: { ja: "セルの長方形をガイドに変換します。", en: "Converts the cell rectangles into guides." },
            toArtboard: {
                ja: "セルごとにアートボードを作成し、長方形は残しません。",
                en: "Creates one artboard per cell and keeps no rectangles."
            },
            division: {
                ja: "セル数より多いオブジェクトは、アートボードの外へ退避します。",
                en: "Objects beyond the number of cells are parked outside the artboard."
            },
            gutter: { ja: "行数・列数のいずれかが2以上のときに有効です。", en: "Available when the rows or columns are 2 or more." },
            margin: { ja: "配置先の四辺から内側に取る余白です。", en: "Inset taken from each edge of the placement area." },
            randomize: {
                ja: "セルへの割り当て順をシャッフルします（押すたびに変わります）。",
                en: "Shuffles the order in which objects fill the cells (changes on every click)."
            },
            transparencyGrid: {
                ja: "透明グリッドの表示を切り替えます。スクリプト終了時に元へ戻します。",
                en: "Toggles the transparency grid. It is restored when the script finishes."
            },
            /* ステップボタン用 / for the stepper */
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
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
            invalidGrid: {
                ja: "この設定ではグリッドを作れません。行数・列数・間隔・マージンを見直してください。",
                en: "These settings cannot form a grid. Check the rows, columns, gutter, and margin."
            },
            artboardError: { ja: "アートボードの作成中にエラーが発生しました。", en: "Error occurred while creating artboards." },
            artboardCreated: { ja: " 個のアートボードを作成しました。", en: " artboards created." }
        }
    };

    // =========================================
    // 位置の記録と復元 / Position bookkeeping
    // =========================================
    /* プレビューの巻き戻しは app.undo() を主、ここでの座標復元を保険とする。
       app.undo() が効けば履歴は伸びないが、1プレビューが複数の履歴ステップに
       分かれると 1 回では戻りきらない。その差分だけをここで埋める。
       位置が変わっていなければ translate しないので、undo が完全に効いた場合は
       履歴を 1 ステップも増やさない。
       Preview rollback relies on app.undo(); restoring centers here only fills the gap when one
       preview spans several history steps, and never moves items that are already in place. */

    /** 位置が変化したと見なす最小差分（pt） / Minimum delta treated as a real move. */
    var RESTORE_TOLERANCE = 0.001;

    /** @type {Array<number[]>} 記録した中心座標 [x, y] の配列。 */
    var originalCenters = [];

    /**
     * 各オブジェクトの中心座標を記録します。/ Record the center of each object.
     *
     * @param {Array<PageItem>} pageItems - 記録対象のオブジェクト。
     * @returns {void}
     */
    function saveOriginalCenters(pageItems) {
        originalCenters = [];
        for (var i = 0; i < pageItems.length; i++) {
            var bounds = pageItems[i].visibleBounds;
            originalCenters.push([(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2]);
        }
    }

    /**
     * 記録した中心座標へオブジェクトを戻します。/ Move objects back to the recorded centers.
     * 既に元の位置にあるものは動かさないため、履歴を無駄に増やしません。
     *
     * @param {Array<PageItem>} pageItems - 復元対象のオブジェクト（記録時と同じ順序）。
     * @returns {void}
     */
    function restoreOriginalCenters(pageItems) {
        for (var i = 0; i < pageItems.length && i < originalCenters.length; i++) {
            try {
                var bounds = pageItems[i].visibleBounds;
                var dx = originalCenters[i][0] - (bounds[0] + bounds[2]) / 2;
                var dy = originalCenters[i][1] - (bounds[1] + bounds[3]) / 2;

                /* undo が効いていれば差分は 0。ここで translate すると履歴が伸びる / Zero after a full undo; moving would add history */
                if (Math.abs(dx) < RESTORE_TOLERANCE && Math.abs(dy) < RESTORE_TOLERANCE) continue;
                pageItems[i].translate(dx, dy);
            } catch (e) {
                /* 参照が失われたオブジェクトは飛ばす / Skip items whose reference went stale */
            }
        }
    }

    // =========================================
    // 退避 / Parking
    // =========================================

    /**
     * 2つの矩形が重なるかを判定します（余裕を加味）。/ Test whether two rects overlap.
     *
     * @param {number[]} rectA - 矩形A [左, 上, 右, 下]。
     * @param {number[]} rectB - 矩形B [左, 上, 右, 下]。
     * @param {number} tolerance - 重なりと見なす余裕（pt）。
     * @returns {boolean} 重なっている場合は true。
     */
    function rectsOverlap(rectA, rectB, tolerance) {
        return !(rectA[2] < rectB[0] - tolerance || rectA[0] > rectB[2] + tolerance ||
            rectA[1] < rectB[3] - tolerance || rectA[3] > rectB[1] + tolerance);
    }

    /**
     * 対象領域や他アートボードと重ならない位置へオブジェクトを退避します。
     * / Park objects clear of the target area and other artboards.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {Array<PageItem>} pageItems - 退避するオブジェクト。
     * @param {number[]} avoidRect - 避けたい配置先の矩形 [左, 上, 右, 下]。
     * @returns {void}
     */
    function parkItemsOutside(doc, pageItems, avoidRect) {
        var activeIndex = doc.artboards.getActiveArtboardIndex();
        var activeRect = doc.artboards[activeIndex].artboardRect;

        /* アクティブアートボードと対象領域の和を「避けたい領域」とする / Avoid the union of the active artboard and the target area */
        var avoidArea = [
            Math.min(activeRect[0], avoidRect[0]),
            Math.max(activeRect[1], avoidRect[1]),
            Math.max(activeRect[2], avoidRect[2]),
            Math.min(activeRect[3], avoidRect[3])
        ];

        var parkedRects = [];
        var artboardTolerance = CONFIG.overlapTolerance + CONFIG.artboardBuffer;

        /**
         * 退避先が既存の退避先や他アートボードと衝突するかを判定します。
         * / Test a parking slot against parked items and artboards.
         *
         * @param {number[]} candidateRect - 退避先の候補矩形。
         * @returns {boolean} 使用できない場合は true。
         */
        function isSlotTaken(candidateRect) {
            for (var i = 0; i < parkedRects.length; i++) {
                if (rectsOverlap(candidateRect, parkedRects[i], CONFIG.overlapTolerance)) return true;
            }
            for (var j = 0; j < doc.artboards.length; j++) {
                if (j === activeIndex) continue;
                if (rectsOverlap(candidateRect, doc.artboards[j].artboardRect, artboardTolerance)) return true;
            }
            return false;
        }

        for (var i = 0; i < pageItems.length; i++) {
            var bounds = pageItems[i].visibleBounds;
            if (!rectsOverlap(bounds, avoidArea, 0)) continue;

            var dx = (avoidArea[2] - bounds[0]) + CONFIG.parkingStep;
            var slotRect;
            do {
                slotRect = [bounds[0] + dx, bounds[1], bounds[2] + dx, bounds[3]];
                if (!isSlotTaken(slotRect)) break;
                dx += CONFIG.parkingStep;
            } while (true);

            pageItems[i].translate(dx, 0);
            parkedRects.push(slotRect);
        }
    }

    // =========================================
    // 対象の検出 / Target detection
    // =========================================

    /* スクリプトが作る・使うレイヤー名の集合。最背面の検出ではこれらを飛ばす
       Layer names this script manages; skipped when looking for the backmost object */
    var SYSTEM_LAYER_NAMES = {};
    SYSTEM_LAYER_NAMES[CONFIG.targetLayerName] = true;
    SYSTEM_LAYER_NAMES[CONFIG.cellLayerName] = true;
    SYSTEM_LAYER_NAMES[CONFIG.previewCellLayerName] = true;
    SYSTEM_LAYER_NAMES[CONFIG.legacyPreviewLayerName] = true;

    /**
     * パスが枠として使える広がりを持つか判定します。/ Whether a path has usable extents for a frame.
     *
     * @param {PathItem} pathItem - 対象のパス。
     * @returns {boolean} 幅・高さともに正なら true。
     */
    function hasUsableArea(pathItem) {
        try {
            var bounds = pathItem.geometricBounds;
            return (bounds[2] - bounds[0] > 0) && (bounds[1] - bounds[3] > 0);
        } catch (e) {
            /* 空のパスなどは geometricBounds が例外 / geometricBounds throws on empty paths */
            return false;
        }
    }

    /**
     * パスが長方形か判定します（閉じた4点で、各点が角）。/ Whether a path is a rectangle.
     *
     * @param {PathItem} pathItem - 対象のパス。
     * @returns {boolean} 長方形とみなせる場合は true。
     */
    function isRectanglePath(pathItem) {
        try {
            if (!pathItem.closed || pathItem.pathPoints.length !== 4) return false;
            for (var i = 0; i < 4; i++) {
                var pathPoint = pathItem.pathPoints[i];
                /* ハンドルが出ていれば角丸などで長方形ではない / Handles mean rounded or curved corners */
                if (pathPoint.leftDirection[0] !== pathPoint.anchor[0]) return false;
                if (pathPoint.leftDirection[1] !== pathPoint.anchor[1]) return false;
                if (pathPoint.rightDirection[0] !== pathPoint.anchor[0]) return false;
                if (pathPoint.rightDirection[1] !== pathPoint.anchor[1]) return false;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 「_target」レイヤーの長方形を返します。/ Return the rectangle on the "_target" layer.
     * 長方形が無ければ、枠として使える閉じたパスで代用します。開いた線や面積のないパスは対象外です。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {PathItem|null} 見つかった長方形。無ければ null。
     */
    function findTargetLayerRectangle(doc) {
        var fallbackPath = null;
        for (var i = 0; i < doc.layers.length; i++) {
            var candidateLayer = doc.layers[i];
            if (candidateLayer.name !== CONFIG.targetLayerName) continue;

            for (var j = 0; j < candidateLayer.pathItems.length; j++) {
                var pathItem = candidateLayer.pathItems[j];
                /* 枠の上に引いた線などを配置先にしない / A stray line must not become the placement area */
                if (!hasUsableArea(pathItem)) continue;
                if (isRectanglePath(pathItem)) return pathItem;
                if (!fallbackPath && pathItem.closed) fallbackPath = pathItem;
            }
        }
        return fallbackPath;
    }

    /**
     * システムレイヤーを除く最背面のオブジェクトを返します。/ Return the backmost object, ignoring system layers.
     * 選択中のアイテムも候補に含めます（枠として使う場合の利便性を優先）。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {PageItem|null} 最背面のオブジェクト。無ければ null。
     */
    function findBackmostPageItem(doc) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var candidateLayer = doc.layers[i];
            if (!candidateLayer.visible || candidateLayer.locked || SYSTEM_LAYER_NAMES[candidateLayer.name]) continue;
            for (var j = candidateLayer.pageItems.length - 1; j >= 0; j--) {
                var pageItem = candidateLayer.pageItems[j];
                if (pageItem.hidden || pageItem.locked) continue;
                return pageItem;
            }
        }
        return null;
    }

    /**
     * 選択から、指定のアイテムを除いた配列を返します（順序は保持）。/ Copy the selection without one item.
     *
     * @param {Array<PageItem>|null} selectionItems - 選択。
     * @param {PageItem|null} excludedItem - 除くアイテム。
     * @returns {Array<PageItem>} 残ったオブジェクト。
     */
    function excludeItem(selectionItems, excludedItem) {
        var remainingItems = [];
        for (var i = 0; selectionItems && i < selectionItems.length; i++) {
            if (selectionItems[i] !== excludedItem) remainingItems.push(selectionItems[i]);
        }
        return remainingItems;
    }

    /**
     * セルが正方形に近くなる行数・列数を求めます。/ Find rows and columns that make cells nearly square.
     * 余白セルが行／列1本分以上になる構成は除外し、その範囲でセル比1に最も近いものを選びます。
     *
     * @param {number} itemCount - 配置するオブジェクトの数。
     * @param {number[]} targetRect - 配置先の矩形 [左, 上, 右, 下]。
     * @returns {{rowCount: number, columnCount: number}} 行数と列数。
     */
    function computeDefaultDivision(itemCount, targetRect) {
        var fallbackDivision = { rowCount: CONFIG.fallbackDivision, columnCount: CONFIG.fallbackDivision };
        if (!itemCount || itemCount <= 0) return fallbackDivision;
        if (itemCount === 1) return { rowCount: 1, columnCount: 1 };

        var areaWidth = targetRect[2] - targetRect[0];
        var areaHeight = targetRect[1] - targetRect[3];
        if (areaWidth <= 0 || areaHeight <= 0) return fallbackDivision;

        /* セル総数の上限を超える分はどのみち溢れるので、上限の範囲で割り付ける
           Anything beyond the cell cap overflows anyway, so lay out within the cap */
        var cellBudget = Math.min(itemCount, CONFIG.maxCellCount);

        var bestDivision = fallbackDivision;
        var bestAspectRatio = Infinity;
        for (var columnCount = 1; columnCount <= cellBudget; columnCount++) {
            var rowCount = Math.ceil(cellBudget / columnCount);
            var emptyCellCount = rowCount * columnCount - cellBudget;
            if (emptyCellCount > 0 && emptyCellCount >= Math.min(rowCount, columnCount)) continue;

            /* computeGridMetrics() が拒む組み合わせは初期値にしない / Never default to a grid the metrics reject */
            if (rowCount > CONFIG.maxDivision || columnCount > CONFIG.maxDivision) continue;
            if (rowCount * columnCount > CONFIG.maxCellCount) continue;

            var cellWidth = areaWidth / columnCount;
            var cellHeight = areaHeight / rowCount;
            var aspectRatio = (cellWidth > cellHeight) ? (cellWidth / cellHeight) : (cellHeight / cellWidth);
            if (aspectRatio < bestAspectRatio) {
                bestAspectRatio = aspectRatio;
                bestDivision = { rowCount: rowCount, columnCount: columnCount };
            }
        }
        return bestDivision;
    }

    // =========================================
    // レイヤー操作 / Layer helpers
    // =========================================

    /**
     * 指定名のレイヤーを探します。/ Find the layer with the given name.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - 探すレイヤー名。
     * @returns {Layer|null} 見つかったレイヤー。無ければ null。
     */
    function findLayerByName(doc, layerName) {
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            /* getByName は見つからないと例外 / getByName throws when the layer is missing */
            return null;
        }
    }

    /**
     * 指定名のレイヤーがあれば削除します。/ Remove the layer with the given name if present.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - 削除するレイヤー名。
     * @returns {void}
     */
    function removeLayerByName(doc, layerName) {
        var foundLayer = findLayerByName(doc, layerName);
        if (!foundLayer) return;
        try {
            foundLayer.locked = false;
            foundLayer.remove();
        } catch (e) {
            /* 消せない場合は何もしない / Nothing to do when the layer cannot be removed */
        }
    }

    /**
     * レイヤーを取得し、無ければ作成します。/ Get the layer, creating it when missing.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {string} layerName - レイヤー名。
     * @returns {Layer} 取得または作成したレイヤー（ロック解除・表示済み）。
     */
    function getOrCreateLayer(doc, layerName) {
        var foundLayer = findLayerByName(doc, layerName);
        if (!foundLayer) {
            foundLayer = doc.layers.add();
            foundLayer.name = layerName;
        }
        foundLayer.locked = false;
        /* これから描き込むレイヤーは必ず表示にする / A layer we are about to draw into must be visible */
        foundLayer.visible = true;
        return foundLayer;
    }

    /**
     * ロック中でも失敗しないようにレイヤーの表示状態を切り替えます。/ Toggle a layer without failing on locks.
     *
     * @param {Layer|null} layer - 対象のレイヤー。null なら何もしません。
     * @param {boolean} visible - 表示する場合は true。
     * @returns {void}
     */
    function setLayerVisible(layer, visible) {
        if (!layer) return;
        try {
            layer.visible = visible;
        } catch (e) { }
    }

    /**
     * ロック中でも失敗しないように表示状態を切り替えます。/ Toggle visibility without failing on locked items.
     *
     * @param {PageItem} pageItem - 対象のオブジェクト。
     * @param {boolean} hidden - 非表示にする場合は true。
     * @returns {void}
     */
    function setItemHidden(pageItem, hidden) {
        try {
            pageItem.hidden = hidden;
        } catch (e) { }
    }

    /**
     * セル描画レイヤーの長方形を選択します。/ Select the rectangles on the cell layer.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {boolean} 選択できた場合は true。レイヤーが無い場合は false。
     */
    function selectCellRectangles(doc) {
        var cellItems = [];
        try {
            var cellLayer = doc.layers.getByName(CONFIG.cellLayerName);
            for (var i = 0; i < cellLayer.pathItems.length; i++) cellItems.push(cellLayer.pathItems[i]);
        } catch (e) {
            return false;
        }
        doc.selection = cellItems;
        return true;
    }

    // =========================================
    // グリッド / Grid
    // =========================================

    /**
     * グリッドの寸法（すべて pt、矩形は [左, 上, 右, 下]）。
     *
     * @typedef {Object} GridMetrics
     * @property {number} rowCount - 行数。
     * @property {number} columnCount - 列数。
     * @property {number} originLeft - 1行1列目のセル左端。
     * @property {number} originTop - 1行1列目のセル上端。
     * @property {number} cellWidth - セルの幅。
     * @property {number} cellHeight - セルの高さ。
     * @property {number} gutter - セル間の間隔。
     * @property {number[]} targetRect - 配置先の矩形。
     */

    /**
     * 行数・列数・余白からグリッドの寸法を計算します。/ Compute the grid metrics.
     *
     * @param {number|null} rowCount - 行数。
     * @param {number|null} columnCount - 列数。
     * @param {number} margin - マージン（pt）。
     * @param {number} gutter - セル間隔（pt）。
     * @param {number[]} targetRect - 配置先の矩形 [左, 上, 右, 下]。
     * @returns {GridMetrics|null} グリッドの寸法。入力が無効な場合は null。
     */
    function computeGridMetrics(rowCount, columnCount, margin, gutter, targetRect) {
        if (rowCount === null || columnCount === null) return null;
        if (rowCount < 1 || columnCount < 1) return null;
        if (rowCount > CONFIG.maxDivision || columnCount > CONFIG.maxDivision) return null;
        if (rowCount * columnCount > CONFIG.maxCellCount) return null;

        var usableWidth = (targetRect[2] - margin) - (targetRect[0] + margin);
        var usableHeight = (targetRect[1] - margin) - (targetRect[3] + margin);
        var cellWidth = (usableWidth - (columnCount - 1) * gutter) / columnCount;
        var cellHeight = (usableHeight - (rowCount - 1) * gutter) / rowCount;

        /* マージンや間隔が大きすぎてセルが成立しない場合は描画しない / No grid when margins or gutters leave no room */
        if (!isFinite(cellWidth) || !isFinite(cellHeight) || cellWidth <= 0 || cellHeight <= 0) return null;

        return {
            rowCount: rowCount,
            columnCount: columnCount,
            originLeft: targetRect[0] + margin,
            originTop: targetRect[1] - margin,
            cellWidth: cellWidth,
            cellHeight: cellHeight,
            gutter: gutter,
            targetRect: targetRect
        };
    }

    /**
     * 全セルを行→列の順に走査します。/ Iterate cells in row-major order.
     *
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {function(number[]): (boolean|void)} handleCell - セル矩形を受け取る処理。false を返すと走査を中断します。
     * @returns {void}
     */
    function forEachCell(gridMetrics, handleCell) {
        for (var i = 0; i < gridMetrics.rowCount; i++) {
            var cellTop = gridMetrics.originTop - (gridMetrics.cellHeight + gridMetrics.gutter) * i;
            for (var j = 0; j < gridMetrics.columnCount; j++) {
                var cellLeft = gridMetrics.originLeft + (gridMetrics.cellWidth + gridMetrics.gutter) * j;
                var cellRect = [cellLeft, cellTop, cellLeft + gridMetrics.cellWidth, cellTop - gridMetrics.cellHeight];
                if (handleCell(cellRect) === false) return;
            }
        }
    }

    /**
     * 全セルの長方形を描画します。/ Draw rectangles for every cell.
     *
     * @param {Layer} cellLayer - 描画先のレイヤー。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @param {boolean} asGuide - ガイドに変換する場合は true。
     * @param {CMYKColor|RGBColor|null} fillColor - セルの塗り色。塗りなしは null。
     * @param {number|null} opacity - セルの不透明度（%）。指定しない場合は null。
     * @returns {void}
     */
    function drawCells(cellLayer, gridMetrics, asGuide, fillColor, opacity) {
        /* 省略された場合も「塗りなし」として扱う / Treat a missing argument as no fill */
        var hasFill = (fillColor !== null && fillColor !== undefined);

        forEachCell(gridMetrics, function (cellRect) {
            var cellRectangle = cellLayer.pathItems.rectangle(
                cellRect[1], cellRect[0], gridMetrics.cellWidth, gridMetrics.cellHeight);
            cellRectangle.stroked = false;
            cellRectangle.filled = hasFill;
            if (hasFill) cellRectangle.fillColor = fillColor;
            if (opacity !== null) cellRectangle.opacity = opacity;
            if (asGuide) cellRectangle.guides = true;
        });
        cellLayer.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 各セルの中央へオブジェクトを1つずつ配置します。/ Place one object at the center of each cell.
     *
     * @param {Array<PageItem>} pageItems - 配置するオブジェクト（配置順）。
     * @param {GridMetrics} gridMetrics - グリッドの寸法。
     * @returns {void}
     */
    function placeItemsInCells(pageItems, gridMetrics) {
        var itemIndex = 0;
        forEachCell(gridMetrics, function (cellRect) {
            if (itemIndex >= pageItems.length) return false;

            var bounds = pageItems[itemIndex].visibleBounds;
            var dx = (cellRect[0] + cellRect[2]) / 2 - (bounds[0] + bounds[2]) / 2;
            var dy = (cellRect[1] + cellRect[3]) / 2 - (bounds[1] + bounds[3]) / 2;
            pageItems[itemIndex].translate(dx, dy);
            itemIndex++;
        });
    }

    /**
     * セルの寸法からアートボードを作成します。/ Create artboards from the cell metrics.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {GridMetrics|null} gridMetrics - グリッドの寸法。
     * @param {number} baseArtboardIndex - 作成後にアクティブへ戻すアートボードの番号。
     * @returns {number} 作成したアートボードの数。
     */
    function createArtboardsFromCells(doc, gridMetrics, baseArtboardIndex) {
        if (!gridMetrics) return 0;

        var createdCount = 0;
        forEachCell(gridMetrics, function (cellRect) {
            try {
                doc.artboards.add(cellRect);
                createdCount++;
            } catch (e) {
                /* アートボード数の上限などで失敗した場合は続行 / Keep going when a single artboard fails */
            }
        });
        doc.artboards.setActiveArtboardIndex(baseArtboardIndex);
        return createdCount;
    }

    /**
     * ドキュメントのカラーモードに合わせた無彩色を返します。/ Return a gray matching the document color mode.
     *
     * @param {Document} doc - 対象ドキュメント。
     * @param {number} cmykBlack - CMYK のときのブラック値（0〜100）。
     * @param {number} rgbLevel - RGB のときの階調値（0〜255）。
     * @returns {CMYKColor|RGBColor} 生成したカラー。
     */
    function createGrayColor(doc, cmykBlack, rgbLevel) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = cmykBlack;
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = rgbLevel;
        rgbColor.green = rgbLevel;
        rgbColor.blue = rgbLevel;
        return rgbColor;
    }

    // =========================================
    // 配置順 / Placement order
    // =========================================

    /**
     * 配列に指定アイテムが含まれるかを判定します。/ Test whether the list contains the item.
     *
     * @param {Array<PageItem>} pageItems - 検索対象の配列。
     * @param {PageItem} pageItem - 探すオブジェクト。
     * @returns {boolean} 含まれる場合は true。
     */
    function containsItem(pageItems, pageItem) {
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i] === pageItem) return true;
        }
        return false;
    }

    /**
     * ランダム順を現在の配置対象に合わせ直します。/ Rebuild the random order for the current items.
     * 対象の切り替えでアイテムが増減しても破綻しないようにします。
     *
     * @param {Array<PageItem>} distributionItems - 現在の配置対象。
     * @param {Array<PageItem>|null} randomizedOrder - シャッフルで決めた順。未使用なら null。
     * @returns {Array<PageItem>} 配置順に並べたオブジェクト。
     */
    function orderByRandomizedOrder(distributionItems, randomizedOrder) {
        if (!randomizedOrder) return distributionItems;

        var orderedItems = [];
        for (var i = 0; i < randomizedOrder.length; i++) {
            if (containsItem(distributionItems, randomizedOrder[i])) orderedItems.push(randomizedOrder[i]);
        }
        for (var j = 0; j < distributionItems.length; j++) {
            if (!containsItem(orderedItems, distributionItems[j])) orderedItems.push(distributionItems[j]);
        }
        return orderedItems;
    }

    /**
     * 順序をシャッフルした新しい配列を返します（Fisher-Yates）。/ Return a shuffled copy.
     *
     * @param {Array<PageItem>} sourceItems - 元の配列。
     * @returns {Array<PageItem>} シャッフルした配列。
     */
    function createShuffledCopy(sourceItems) {
        var shuffledItems = [];
        for (var i = 0; i < sourceItems.length; i++) shuffledItems.push(sourceItems[i]);
        for (var k = shuffledItems.length - 1; k > 0; k--) {
            var swapIndex = Math.floor(Math.random() * (k + 1));
            var swapItem = shuffledItems[k];
            shuffledItems[k] = shuffledItems[swapIndex];
            shuffledItems[swapIndex] = swapItem;
        }
        return shuffledItems;
    }

    // =========================================
    // UIの組み立て / UI construction
    // =========================================

    /**
     * 対象パネルを作成します。/ Build the target panel.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @param {boolean} hasBackmostItem - 最背面のオブジェクトが存在する場合は true。
     * @param {boolean} hasTargetRect - 「_target」レイヤーの長方形が存在する場合は true。
     * @returns {{artboardRadio: RadioButton, backmostRadio: RadioButton, rectLayerRadio: RadioButton}} 配置先のラジオボタン。
     */
    function buildPlacementTargetPanel(parentWindow, hasBackmostItem, hasTargetRect) {
        var placementPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.placement));
        setupPanel(placementPanel, 6);
        placementPanel.alignChildren = ["left", "top"];

        var placementControls = {
            artboardRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetArtboard)),
            backmostRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetBackmost)),
            rectLayerRadio: placementPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetRectLayer))
        };

        placementControls.backmostRadio.enabled = hasBackmostItem;
        placementControls.rectLayerRadio.enabled = hasTargetRect;

        /* 選べない理由、または選んだときの挙動をツールチップで補う / Tooltips explain why an option is unavailable or what it does */
        placementControls.backmostRadio.helpTip = hasBackmostItem
            ? getLabel(LABELS.tooltip.targetBackmost)
            : getLabel(LABELS.tooltip.targetUnavailable);
        placementControls.rectLayerRadio.helpTip = hasTargetRect
            ? getLabel(LABELS.tooltip.targetRectLayer)
            : getLabel(LABELS.tooltip.targetUnavailable);

        if (hasTargetRect) {
            placementControls.rectLayerRadio.value = true;
        } else {
            placementControls.artboardRadio.value = true;
        }
        return placementControls;
    }

    /**
     * ラベルと入力欄の1行を作成します。/ Build one label-and-field row.
     *
     * @param {Group} parentColumn - 配置先のカラム。
     * @param {LabelEntry} labelEntry - 行ラベルの文言。
     * @param {number} initialValue - 入力欄の初期値。
     * @param {number} labelWidth - ラベルの幅（px）。
     * @param {string|null} unitSuffix - 入力欄の後ろに置く単位表記。不要なら null。
     * @param {LabelEntry|null} tooltipEntry - ラベルと入力欄に付けるツールチップ。不要なら null。
     * @param {Object} stepOptions - ∧∨の増減設定（addStepper() に渡す。個数の欄は integer: true）。
     * @returns {EditText} 作成した入力欄。
     */
    function addFieldRow(parentColumn, labelEntry, initialValue, labelWidth, unitSuffix, tooltipEntry, stepOptions) {
        var fieldRow = parentColumn.add("group");
        setupRow(fieldRow, "left", 4);
        fieldRow.alignChildren = "center";

        var fieldLabel = fieldRow.add("statictext", undefined, labelText(labelEntry));
        fieldLabel.preferredSize.width = labelWidth;
        if (unitSuffix) fieldLabel.justify = "right";

        var fieldInput = addSteppedInput(fieldRow, String(initialValue), stepOptions);
        fieldInput.characters = unitSuffix ? 4 : 3;
        if (unitSuffix) fieldRow.add("statictext", undefined, unitSuffix);

        if (tooltipEntry) {
            fieldLabel.helpTip = getLabel(tooltipEntry);
            fieldInput.helpTip = getLabel(tooltipEntry);
        }
        return fieldInput;
    }

    /**
     * ∧∨と入力欄を隙間なく並べて行に追加し、↑↓キーも∧∨と同じ処理で増減させます。/ Add a field with a stepper.
     * 増減後は入力欄の onChange（enableNumericField() で結線）を呼び、範囲の正規化とプレビュー更新を手入力と同じにします。
     *
     * @param {Group} parentRow - 配置先の行。
     * @param {string} initialText - 入力欄の初期値。
     * @param {Object} stepOptions - ∧∨の増減設定（addStepper() に渡す。個数の欄は integer: true）。
     * @returns {EditText} 作成した入力欄（∧∨は .stepperGroup で参照できる）。
     */
    function addSteppedInput(parentRow, initialText, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var inputField;
        stepOptions.onStep = function (numberInput) { if (numberInput.onChange) numberInput.onChange(); };
        var stepperGroup = addStepper(stepperInputGroup, function () { return inputField; }, stepOptions);
        inputField = stepperInputGroup.add("edittext", undefined, initialText);
        inputField.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(inputField, stepperGroup);
        return inputField;
    }

    /**
     * 入力欄の有効／無効を、左の∧∨ごと切り替えます（∧∨は自作描画なので描き直してディム表示をそろえます）。
     * Enable or disable a field together with its stepper.
     *
     * @param {EditText} inputField - addSteppedInput() で作った入力欄。
     * @param {boolean} isEnabled - 有効にする場合は true。
     * @returns {void}
     */
    function setSteppedInputEnabled(inputField, isEnabled) {
        /* 入力のたびに呼ばれるので、変わらないときは描き直さない / called on every keystroke; skip when unchanged */
        if (inputField.enabled === isEnabled && inputField.stepperGroup.enabled === isEnabled) return;
        inputField.enabled = isEnabled;
        inputField.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(inputField.stepperGroup);
    }

    /**
     * 分割とマージンのパネルを作成します（左：分割数／右：間隔）。/ Build the division panel.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @param {{rowCount: number, columnCount: number}} defaultDivision - 行数・列数の初期値。
     * @param {string} unitLabel - 定規単位のラベル。
     * @returns {{rowCountInput: EditText, columnCountInput: EditText, gutterInput: EditText, marginInput: EditText}} 入力欄。
     */
    function buildDivisionPanel(parentWindow, defaultDivision, unitLabel) {
        var divisionPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.division));
        setupPanel(divisionPanel, 6);
        divisionPanel.orientation = "row";
        divisionPanel.alignChildren = ["left", "top"];
        divisionPanel.spacing = COLUMN_SPACING;

        var divisionColumn = divisionPanel.add("group");
        divisionColumn.orientation = "column";
        divisionColumn.alignChildren = "left";
        divisionColumn.margins.right = COLUMN_SPACING;

        var spacingColumn = divisionPanel.add("group");
        spacingColumn.orientation = "column";
        spacingColumn.alignChildren = "left";

        /* ラベル幅は日本語環境と英語環境で個別に指定 / Label widths per language */
        var divisionLabelWidth = (uiLang === "ja") ? 45 : 60;
        var spacingLabelWidth = (uiLang === "ja") ? 70 : 75;

        return {
            /* 範囲の正規化は enableNumericField() の onChange が受け持つ / range clamping is done in enableNumericField() */
            rowCountInput: addFieldRow(divisionColumn, LABELS.fieldLabel.rowCount, defaultDivision.rowCount, divisionLabelWidth, null, LABELS.tooltip.division, { integer: true }),
            columnCountInput: addFieldRow(divisionColumn, LABELS.fieldLabel.columnCount, defaultDivision.columnCount, divisionLabelWidth, null, LABELS.tooltip.division, { integer: true }),
            gutterInput: addFieldRow(spacingColumn, LABELS.fieldLabel.gutter, CONFIG.defaultGutter, spacingLabelWidth, unitLabel, LABELS.tooltip.gutter, {}),
            marginInput: addFieldRow(spacingColumn, LABELS.fieldLabel.margin, CONFIG.defaultMargin, spacingLabelWidth, unitLabel, LABELS.tooltip.margin, {})
        };
    }

    /**
     * セル描画パネルを作成します。/ Build the cell drawing panel.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @returns {Object} セル描画モード、カラー、不透明度、透明グリッドボタンをまとめたオブジェクト。
     */
    function buildCellDrawingPanel(parentWindow) {
        var cellDrawingPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.cellDrawing));
        setupPanel(cellDrawingPanel, 6);

        var cellModeRow = cellDrawingPanel.add("group");
        setupRow(cellModeRow, "left");

        var cellColorRow = cellDrawingPanel.add("group");
        setupRow(cellColorRow, "left");
        cellColorRow.add("statictext", undefined, labelText(LABELS.fieldLabel.cellColor));

        var cellOpacityRow = cellDrawingPanel.add("group");
        setupRow(cellOpacityRow, "left");
        cellOpacityRow.alignChildren = "center";
        cellOpacityRow.add("statictext", undefined, labelText(LABELS.fieldLabel.cellOpacity));

        var cellControls = {
            keepCellRadio: cellModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.keepCell)),
            toGuideRadio: cellModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.toGuide)),
            toArtboardRadio: cellModeRow.add("radiobutton", undefined, getLabel(LABELS.radio.toArtboard)),
            blackCellRadio: cellColorRow.add("radiobutton", undefined, getLabel(LABELS.radio.blackCell)),
            whiteCellRadio: cellColorRow.add("radiobutton", undefined, getLabel(LABELS.radio.whiteCell)),
            transparentCellRadio: cellColorRow.add("radiobutton", undefined, getLabel(LABELS.radio.transparentCell)),
            opacityInput: addSteppedInput(cellOpacityRow, String(CONFIG.blackCellOpacity), {})
        };
        cellControls.keepCellRadio.value = true;
        cellControls.blackCellRadio.value = true;
        cellControls.opacityInput.characters = 4;

        /* モードごとの結果をツールチップで補う / Tooltips describe each mode's result */
        cellControls.keepCellRadio.helpTip = getLabel(LABELS.tooltip.keepCell);
        cellControls.toGuideRadio.helpTip = getLabel(LABELS.tooltip.toGuide);
        cellControls.toArtboardRadio.helpTip = getLabel(LABELS.tooltip.toArtboard);

        cellOpacityRow.add("statictext", undefined, "%");

        /* ボタンはパネル幅いっぱいに広げない / Keep the button from stretching across the panel */
        cellControls.transparencyGridButton = cellOpacityRow.add("button", undefined, getLabel(LABELS.button.transparencyGrid));
        cellControls.transparencyGridButton.alignment = "left";
        cellControls.transparencyGridButton.helpTip = getLabel(LABELS.tooltip.transparencyGrid);
        return cellControls;
    }

    /**
     * 下部のボタン列を作成します。/ Build the footer button row.
     *
     * @param {Window} parentWindow - 配置先のウィンドウ。
     * @returns {{btnShuffle: Button, btnCancel: Button, btnOK: Button}} 各ボタン。
     */
    function buildFooterRow(parentWindow) {
        var buttonRow = addButtonRow(parentWindow);
        var btnShuffle = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.randomize));
        btnShuffle.helpTip = getLabel(LABELS.tooltip.randomize);

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        var footerButtons = { btnShuffle: btnShuffle, btnCancel: btnCancel, btnOK: btnOK };
        return footerButtons;
    }

    // =========================================
    // 入力欄 / Input fields
    // =========================================

    /**
     * 入力欄の値を整数として読み取ります。/ Read an integer from a field.
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @returns {number|null} 読み取った整数。数値でない場合は null。
     */
    function readCount(inputField) {
        var parsedCount = parseInt(inputField.text, 10);
        return isFinite(parsedCount) ? parsedCount : null;
    }

    /**
     * 入力欄の値を pt 換算の長さとして読み取ります。/ Read a length in points from a field.
     * 負の値は配置先からはみ出したグリッドを生むため、0 で止めます。
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @param {number} pointsPerUnit - 定規単位 1 あたりの pt。
     * @returns {number} pt に換算した長さ。数値でない場合は 0。
     */
    function readLengthPt(inputField, pointsPerUnit) {
        var parsedLength = parseFloat(inputField.text);
        if (!isFinite(parsedLength)) return 0;
        return Math.max(0, parsedLength) * pointsPerUnit;
    }

    /**
     * 数値入力欄に範囲の正規化を組み込みます（∧∨・↑↓キーの増減も、この onChange を通ります）。
     * Wire range clamping into a numeric field; the stepper and arrow keys go through this onChange too.
     *
     * @param {EditText} inputField - 対象の入力欄。
     * @param {number} minValue - 下限値。
     * @param {number|null} maxValue - 上限値。上限なしは null。
     * @param {function(): void} onCommit - 値を変えたあとに呼ぶ処理。
     * @returns {void}
     */
    function enableNumericField(inputField, minValue, maxValue, onCommit) {
        /** 入力欄を書き換えている最中か（onChange の再入を防ぐ） / Guard against re-entering onChange. */
        var isNormalizing = false;

        /**
         * 値を下限・上限に収めます。/ Clamp a value into range.
         *
         * @param {number} value - 対象の値。
         * @returns {number} 範囲内に収めた値。
         */
        function clampValue(value) {
            var clamped = Math.max(minValue, value);
            return (maxValue === null) ? clamped : Math.min(maxValue, clamped);
        }

        /* 入力を確定した時点で範囲外の値を直し、表示と描画を一致させる
           Normalize out-of-range text on commit so the field matches what is drawn */
        inputField.onChange = function () {
            if (isNormalizing) return;

            var fieldValue = parseFloat(inputField.text);
            if (isFinite(fieldValue)) {
                var clamped = clampValue(fieldValue);
                if (clamped !== fieldValue) {
                    isNormalizing = true;
                    inputField.text = String(clamped);
                    isNormalizing = false;
                }
            }
            onCommit();
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示して配置処理を実行します。/ Show the dialog and run the distribution.
     *
     * @returns {void}
     */
    function showDistributeDialog() {
        var doc = app.activeDocument;

        /* 前回実行時のセル描画レイヤーは、OK を押すまで消さずに隠しておく。
           ここで削除するとキャンセルしても戻せない（app.undo() はここまで巻き戻せない）
           Hide the cell layer from a previous run instead of deleting it, so Cancel can restore it */
        var previousCellLayer = findLayerByName(doc, CONFIG.cellLayerName);
        var wasPreviousCellLayerVisible = previousCellLayer ? previousCellLayer.visible : false;
        setLayerVisible(previousCellLayer, false);

        /* 以後の計算は「ダイアログ起動時点のアクティブアートボード」を基準に固定 / Fix the base to the artboard active at launch */
        var baseArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var baseArtboardRect = doc.artboards[baseArtboardIndex].artboardRect;

        /* 対象候補の検出（この時点では変更を加えない）/ Detect target candidates without changing anything */
        var targetRectItem = findTargetLayerRectangle(doc);
        var backmostItem = findBackmostPageItem(doc);

        /* 選択内容はダイアログ起動時に固定する（順序も保持）。枠として使う「_target」矩形は移動対象から除く
           Freeze the selection (and its order) at launch, leaving out the _target frame rectangle */
        var placeableItems = excludeItem(doc.selection, targetRectItem);
        saveOriginalCenters(placeableItems);

        /* 初期ターゲット矩形（_target 矩形があればそれ、なければ現在のアートボード）/ Initial target rect */
        var initialTargetRect = targetRectItem ? targetRectItem.geometricBounds : baseArtboardRect;

        /* 「_target」レイヤーの矩形は対象として選ばれている間は非表示にする。
           元から隠されていた矩形を勝手に表示しないよう、元の状態を控えておく
           Hide the _target rectangle while it is the target, remembering its original state */
        var wasTargetRectHidden = targetRectItem ? targetRectItem.hidden : false;
        if (targetRectItem) setItemHidden(targetRectItem, true);

        var rulerUnit = getUnitInfo();
        var defaultDivision = computeDefaultDivision(placeableItems.length, initialTargetRect);
        var transparencyGridToggleCount = 0;

        var distributeDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(distributeDialog);

        var placementUI = buildPlacementTargetPanel(distributeDialog, backmostItem !== null, targetRectItem !== null);
        var divisionUI = buildDivisionPanel(distributeDialog, defaultDivision, rulerUnit.label);
        var cellUI = buildCellDrawingPanel(distributeDialog);
        var footerUI = buildFooterRow(distributeDialog);

        /** @type {Array<PageItem>|null} シャッフルで決めた配置順。未使用なら null。 */
        var randomizedOrder = null;
        /** @type {RadioButton} 「長方形で残す」に戻したときに復帰させるカラー選択。 */
        var lastCellColorRadio = cellUI.blackCellRadio;
        /** @type {boolean} app.undo() で剥がすべきプレビューが適用済みか。 */
        var hasUncommittedPreview = false;
        /** @type {boolean} OK／キャンセルで後始末済みか（onClose の二重実行を防ぐ）。 */
        var isCleanedUp = false;
        /** @type {number} 直近にプレビューを描き直した時刻（ミリ秒）。入力中の間引きに使います。 */
        var lastPreviewTime = 0;

        bindDialogEvents();

        /**
         * 行数・列数を変えたあとに間隔欄とプレビューを更新します。/ Refresh after a row or column change.
         *
         * @returns {void}
         */
        function handleDivisionChange() {
            syncGutterEnabled();
            updatePreview();
        }

        /* 行数・列数は 1 以上、上限はグリッドが成立する範囲まで / Rows and columns stay within a workable grid */
        enableNumericField(divisionUI.rowCountInput, 1, CONFIG.maxDivision, handleDivisionChange);
        enableNumericField(divisionUI.columnCountInput, 1, CONFIG.maxDivision, handleDivisionChange);
        /* 間隔・マージンは負にしない（配置先の外へはみ出す）/ Gutter and margin never go negative */
        enableNumericField(divisionUI.gutterInput, 0, null, updatePreview);
        enableNumericField(divisionUI.marginInput, 0, null, updatePreview);
        /* 不透明度は 0〜100%（描画側のクランプと表示を一致させる）/ Opacity matches what is actually drawn */
        enableNumericField(cellUI.opacityInput, 0, 100, updatePreview);

        /* 初期状態を反映してプレビューを表示 / Apply the initial state and show the preview */
        syncGutterEnabled();
        syncCellModeUI();
        updatePreview();

        distributeDialog.layout.layout(true);
        trimButtonHeight(cellUI.transparencyGridButton, 4);
        prepareDialogWindow(distributeDialog, SCRIPT_NAME);
        distributeDialog.show();

        /* 透明グリッドの表示状態を元へ戻す / Restore the transparency grid */
        if (transparencyGridToggleCount % 2 !== 0) {
            app.executeMenuCommand('TransparencyGrid Menu Item');
        }

        // -----------------------------------------
        // 入力値の取得 / Input readers
        // -----------------------------------------

        /**
         * 不透明度を 0〜100 に収めて読み取ります。/ Read the opacity clamped to 0-100.
         *
         * @returns {number|null} 不透明度（%）。数値でない場合は null。
         */
        function readOpacity() {
            var parsedOpacity = parseFloat(cellUI.opacityInput.text);
            if (!isFinite(parsedOpacity)) return null;
            return Math.max(0, Math.min(100, parsedOpacity));
        }

        /**
         * 現在のセル描画モードを返します。/ Return the current cell drawing mode.
         *
         * @returns {string} "keep"、"guide"、"artboard" のいずれか。
         */
        function getCellMode() {
            if (cellUI.toArtboardRadio.value) return "artboard";
            if (cellUI.toGuideRadio.value) return "guide";
            return "keep";
        }

        /**
         * 選択中の対象領域の矩形を返します。/ Return the rect of the selected target area.
         *
         * @returns {number[]} 配置先の矩形 [左, 上, 右, 下]。
         */
        function getTargetRect() {
            if (placementUI.backmostRadio.value && backmostItem) return backmostItem.geometricBounds;
            if (placementUI.rectLayerRadio.value && targetRectItem) return targetRectItem.geometricBounds;
            if (doc.artboards.length > baseArtboardIndex) return doc.artboards[baseArtboardIndex].artboardRect;
            return baseArtboardRect;
        }

        /**
         * 入力欄からグリッドの寸法を求めます。/ Compute the grid metrics from the fields.
         *
         * @returns {GridMetrics|null} グリッドの寸法。入力が無効な場合は null。
         */
        function readGridMetrics() {
            return computeGridMetrics(
                readCount(divisionUI.rowCountInput),
                readCount(divisionUI.columnCountInput),
                readLengthPt(divisionUI.marginInput, rulerUnit.pointsPerUnit),
                readLengthPt(divisionUI.gutterInput, rulerUnit.pointsPerUnit),
                getTargetRect()
            );
        }

        /**
         * セルの塗り色を返します。/ Return the cell fill color.
         *
         * @returns {CMYKColor|RGBColor|null} 塗り色。透過を選んでいる場合は null。
         */
        function getCellFillColor() {
            if (cellUI.blackCellRadio.value) return createGrayColor(doc, 100, 0);
            if (cellUI.whiteCellRadio.value) return createGrayColor(doc, 0, 255);
            return null;
        }

        // -----------------------------------------
        // 配置対象 / Distribution items
        // -----------------------------------------

        /**
         * 枠として使うアイテムを除いた配置対象を返します。/ Return the items to place, excluding frame items.
         *
         * @returns {Array<PageItem>} 配置対象のオブジェクト。
         */
        function getDistributionItems() {
            if (!(placementUI.backmostRadio.value && backmostItem)) return placeableItems;
            return excludeItem(placeableItems, backmostItem);
        }

        // -----------------------------------------
        // 描画と配置 / Drawing and placement
        // -----------------------------------------

        /**
         * セル描画と配置を実行します（プレビュー／本番共通）。/ Draw cells and place objects.
         *
         * @param {boolean} isPreview - プレビューとして描画する場合は true。
         * @param {function(): void} [onWillMutate] - ドキュメントを変更する直前に一度だけ呼ばれます。
         * @returns {boolean} ドキュメントを変更した場合は true。入力が無効で何もしなかった場合は false。
         */
        function renderDistribution(isPreview, onWillMutate) {
            /* 配置先の座標を読む前に、必ず元の位置へ戻しておく（通常は clearPreview() 済みで何も動かない）
               Restore positions before reading the target; normally a no-op after clearPreview() */
            restoreOriginalCenters(placeableItems);

            var cellMode = getCellMode();
            var distributionItems = orderByRandomizedOrder(getDistributionItems(), randomizedOrder);
            var gridMetrics = readGridMetrics();

            /* グリッドが成立しない入力では、ドキュメントに一切手を加えない / Leave the document untouched without a valid grid */
            if (!gridMetrics) return false;

            /* これ以降はドキュメントを変更する / Everything below mutates the document */
            if (onWillMutate) onWillMutate();

            /* セル数を超えたオブジェクトだけを対象領域の外へ退避する / Park only the objects beyond the cell count */
            var overflowItems = distributionItems.slice(gridMetrics.rowCount * gridMetrics.columnCount);
            if (overflowItems.length > 0) {
                parkItemsOutside(doc, overflowItems, gridMetrics.targetRect);
            }

            /* アートボード化のプレビューは、セルの範囲をガイドで示す / The artboard mode previews its cells as guides */
            if ((cellMode !== "artboard") || isPreview) {
                var layerName = isPreview ? CONFIG.previewCellLayerName : CONFIG.cellLayerName;
                var asGuide = (cellMode === "guide") || (cellMode === "artboard" && isPreview);
                drawCells(getOrCreateLayer(doc, layerName), gridMetrics, asGuide, getCellFillColor(), readOpacity());
            }

            placeItemsInCells(distributionItems, gridMetrics);
            return true;
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------
        /* app.undo() で直前のプレビューをヒストリごと取り除いてから描き直す。
           これをしないと、入力欄を 1 文字打つたびに数十〜数千ステップが積まれ、
           Illustrator の取り消し回数の上限を超えてユーザーの実行前履歴が失われる。

           ただし app.undo() は 1 回で 1 ステップしか戻さない。1 プレビューが
           複数ステップに分かれる場合は戻りきらないため、レイヤー削除と座標復元を
           保険として必ず併走させる（どちらも差分が無ければ何もしない）。

           undo の回数は「自分が積んだ 1 回分」に限る。ダイアログ表示前に
           cell-background レイヤーの非表示と _target 矩形の非表示という 2 つの
           変更を済ませており、そこまで巻き戻すと状態が壊れるため。
           Each preview is peeled off with one app.undo() so typing does not flood the history;
           layer removal and center restoring back it up, and undo never reaches the two changes
           made before the dialog opened. */

        /**
         * プレビューを消して元の状態に戻します。/ Discard the preview and restore the original state.
         *
         * @returns {void}
         */
        function clearPreview() {
            /* 1. 直前のプレビューをヒストリから取り除く / Peel the last preview off the history */
            if (hasUncommittedPreview) {
                hasUncommittedPreview = false;
                try {
                    app.undo();
                } catch (e) {
                    /* undo できない状態なら、以下の後始末に任せる / Fall back to the cleanup below */
                }
            }

            /* 2. undo で戻りきらなかった分だけを片付ける / Clean up whatever undo left behind */
            removeLayerByName(doc, CONFIG.previewCellLayerName);
            removeLayerByName(doc, CONFIG.legacyPreviewLayerName);
            restoreOriginalCenters(placeableItems);
        }

        /**
         * 現在の設定でプレビューを描き直します。/ Redraw the preview with the current settings.
         *
         * @returns {void}
         */
        function updatePreview() {
            try {
                clearPreview();
                renderDistribution(true, function () {
                    /* 変更が始まった時点で印を付ける。途中で例外が出ても undo で剥がせる / Mark first so undo can peel it even on failure */
                    hasUncommittedPreview = true;
                });
            } catch (e) {
                clearPreview();
            }
            app.redraw();
            lastPreviewTime = (new Date()).getTime();
        }

        /**
         * 入力中のプレビューを更新します（重いグリッドでは打鍵を間引きます）。
         * Refresh the preview while typing, throttling only when each redraw is slow.
         *
         * セル数が多いと1回の描き直しで数百〜数千の長方形を作り直すため、
         * 「100」と打つだけで同じ処理が3回走る。見送った分は入力欄の確定（onChange）で
         * 必ず描き直されるので、表示が取り残されたままにはならない。
         * A skipped redraw is always flushed when the field commits, so the preview never stays stale.
         *
         * @returns {void}
         */
        function updatePreviewWhileTyping() {
            var rowCount = readCount(divisionUI.rowCountInput);
            var columnCount = readCount(divisionUI.columnCountInput);
            var cellCount = (rowCount !== null && columnCount !== null) ? rowCount * columnCount : 0;

            /* 軽いグリッドは打鍵ごとに描き直す / Small grids stay fully live */
            if (cellCount <= CONFIG.previewThrottleCells) {
                updatePreview();
                return;
            }
            if ((new Date()).getTime() - lastPreviewTime < CONFIG.previewThrottleMs) return;
            updatePreview();
        }

        // -----------------------------------------
        // UI 状態の同期 / UI state
        // -----------------------------------------

        /**
         * 間隔欄の有効・無効を切り替えます。/ Enable or disable the gutter field.
         * 間隔は行・列いずれかが2以上のときだけ意味を持ちます。
         *
         * @returns {void}
         */
        function syncGutterEnabled() {
            var rowCount = readCount(divisionUI.rowCountInput);
            var columnCount = readCount(divisionUI.columnCountInput);

            /* 入力途中で両方が空のときは、現在の有効・無効を保つ / Keep the state while both fields are empty */
            if (rowCount === null && columnCount === null) return;
            setSteppedInputEnabled(divisionUI.gutterInput, rowCount > 1 || columnCount > 1);
        }

        /**
         * カラー選択を不透明度欄に反映します。/ Reflect the color choice in the opacity field.
         *
         * @param {boolean} resetsValue - カラーごとの既定値で上書きする場合は true。
         * @returns {void}
         */
        function syncOpacityEnabled(resetsValue) {
            setSteppedInputEnabled(cellUI.opacityInput, cellUI.blackCellRadio.value || cellUI.whiteCellRadio.value);
            if (!resetsValue) return;
            if (cellUI.blackCellRadio.value) {
                cellUI.opacityInput.text = String(CONFIG.blackCellOpacity);
            } else if (cellUI.whiteCellRadio.value) {
                cellUI.opacityInput.text = String(CONFIG.whiteCellOpacity);
            }
        }

        /**
         * セル描画モードに応じてUIを整えます。/ Update the UI for the current cell drawing mode.
         *
         * @returns {void}
         */
        function syncCellModeUI() {
            var usesFill = (getCellMode() === "keep");
            cellUI.blackCellRadio.enabled = usesFill;
            cellUI.whiteCellRadio.enabled = usesFill;
            cellUI.transparentCellRadio.enabled = usesFill;

            if (usesFill) {
                /* 「塗りなし」固定から、直前に選んでいたカラーへ戻す / Return to the color chosen before */
                lastCellColorRadio.value = true;
                syncOpacityEnabled(false);
                return;
            }
            /* ガイド化／アートボード化では塗りを持たないため塗りなしに固定 / Guides and artboards have no fill */
            cellUI.transparentCellRadio.value = true;
            setSteppedInputEnabled(cellUI.opacityInput, false);
        }

        /**
         * 対象を切り替えます（「_target」矩形は選択中だけ非表示）。/ Switch the target area.
         *
         * @returns {void}
         */
        function handleTargetChange() {
            /* 表示状態を変える前にプレビューを剥がす。順序を逆にすると、
               app.undo() がプレビューではなく setItemHidden を取り消してしまう
               Peel the preview first; otherwise app.undo() would revert setItemHidden instead */
            clearPreview();
            /* 元から隠されていた矩形は、対象を外しても隠したままにする / A rectangle hidden by the user stays hidden */
            if (targetRectItem) {
                setItemHidden(targetRectItem, wasTargetRectHidden || placementUI.rectLayerRadio.value === true);
            }
            updatePreview();
        }

        // -----------------------------------------
        // ボタン / Buttons
        // -----------------------------------------

        /**
         * 配置順をシャッフルしてプレビューを描き直します。/ Shuffle the placement order and redraw.
         *
         * @returns {void}
         */
        function shuffleDistributionOrder() {
            var distributionItems = getDistributionItems();
            if (distributionItems.length === 0) {
                alert(getLabel(LABELS.alert.noSelection));
                return;
            }
            randomizedOrder = createShuffledCopy(distributionItems);
            updatePreview();
        }

        /**
         * OK：プレビューを破棄して本番の配置を行い、ダイアログを閉じます。/ Commit the distribution and close.
         *
         * @returns {void}
         */
        function commitDistribution() {
            /* グリッドが成立しない入力では、閉じずに理由を知らせる / Explain instead of closing on an invalid grid */
            if (!readGridMetrics()) {
                alert(getLabel(LABELS.alert.invalidGrid));
                return;
            }

            var cellMode = getCellMode();
            var createdCount = -1;
            var artboardFailed = false;

            /* プレビューを undo で破棄してから本番処理へ（プレビューの痕跡も履歴も残さない）/ Undo the preview before the real run */
            clearPreview();

            /* 前回実行時のセル描画レイヤーは、実際に描き直すこの時点で捨てる
               Discard the previous run's cell layer only now that we are really redrawing */
            removeLayerByName(doc, CONFIG.cellLayerName);

            if (cellMode === "artboard") {
                try {
                    createdCount = createArtboardsFromCells(doc, readGridMetrics(), baseArtboardIndex);
                } catch (e) {
                    artboardFailed = true;
                }
            }
            /* 本番は undo の対象にしない（ユーザーの取り消し操作に委ねる）/ The real run is left to the user's own undo */
            renderDistribution(false);

            /* 「_target」レイヤーの矩形を元の表示状態へ戻す / Restore the _target rectangle's original state */
            if (targetRectItem) setItemHidden(targetRectItem, wasTargetRectHidden);

            /* セル描画を残した場合はその長方形を、それ以外は元の選択を選択状態にする / Select the kept cells, or the original selection */
            if (cellMode !== "keep" || !selectCellRectangles(doc)) {
                doc.selection = placeableItems;
            }

            isCleanedUp = true;
            app.redraw();
            distributeDialog.close(1);

            /* 0 個は入力値が無効でグリッドが成立しなかったケース / Zero means the grid was invalid */
            if (artboardFailed || createdCount === 0) {
                alert(getLabel(LABELS.alert.artboardError));
            } else if (createdCount > 0) {
                alert(createdCount + getLabel(LABELS.alert.artboardCreated));
            }
        }

        /**
         * プレビューを消し、ダイアログ表示前の状態へ戻します。/ Discard the preview and restore the pre-dialog state.
         *
         * @returns {void}
         */
        function discardDialogChanges() {
            clearPreview();
            if (targetRectItem) setItemHidden(targetRectItem, wasTargetRectHidden);
            /* 起動時に隠した前回のセル描画レイヤーを元へ戻す / Restore the cell layer hidden at launch */
            setLayerVisible(previousCellLayer, wasPreviousCellLayerVisible);
            app.redraw();
        }

        // -----------------------------------------
        // イベント / Event wiring
        // -----------------------------------------

        /**
         * ダイアログのイベントを結び付けます。/ Wire up the dialog events.
         *
         * @returns {void}
         */
        function bindDialogEvents() {
            placementUI.artboardRadio.onClick = handleTargetChange;
            placementUI.backmostRadio.onClick = handleTargetChange;
            placementUI.rectLayerRadio.onClick = handleTargetChange;

            cellUI.keepCellRadio.onClick = cellUI.toGuideRadio.onClick = cellUI.toArtboardRadio.onClick = function () {
                syncCellModeUI();
                updatePreview();
            };

            cellUI.blackCellRadio.onClick = cellUI.whiteCellRadio.onClick = cellUI.transparentCellRadio.onClick = function () {
                var selectedColorRadio = cellUI.whiteCellRadio.value ? cellUI.whiteCellRadio
                    : (cellUI.transparentCellRadio.value ? cellUI.transparentCellRadio : cellUI.blackCellRadio);

                /* 選択済みのラジオを押し直しても onClick は呼ばれる。カラーが変わっていなければ
                   入力済みの不透明度を既定値で上書きしない
                   A radio fires onClick even when re-clicked; keep a typed opacity when the color did not change */
                var hasColorChanged = (selectedColorRadio !== lastCellColorRadio);

                /* モードを往復してもカラー選択が失われないように覚えておく / Remember the color across mode switches */
                lastCellColorRadio = selectedColorRadio;
                syncOpacityEnabled(hasColorChanged);
                updatePreview();
            };

            divisionUI.rowCountInput.onChanging = divisionUI.columnCountInput.onChanging = function () {
                syncGutterEnabled();
                updatePreviewWhileTyping();
            };
            divisionUI.gutterInput.onChanging = updatePreviewWhileTyping;
            divisionUI.marginInput.onChanging = updatePreviewWhileTyping;
            cellUI.opacityInput.onChanging = updatePreviewWhileTyping;

            cellUI.transparencyGridButton.onClick = function () {
                app.executeMenuCommand('TransparencyGrid Menu Item');
                transparencyGridToggleCount++;
            };

            footerUI.btnShuffle.onClick = shuffleDistributionOrder;
            footerUI.btnOK.onClick = commitDistribution;

            footerUI.btnCancel.onClick = function () {
                discardDialogChanges();
                isCleanedUp = true;
                distributeDialog.close(0);
            };

            /* OK／キャンセルを経由せずに閉じた場合の保険（ESC やウィンドウを閉じた操作でプレビューを残さない）
               Safety net when the dialog closes without OK or Cancel (ESC, closing the window) */
            distributeDialog.onClose = function () {
                if (!isCleanedUp) {
                    isCleanedUp = true;
                    try {
                        discardDialogChanges();
                    } catch (e) {
                        /* 後始末の失敗でクローズを妨げない / Never block the close */
                    }
                }
                /* falsy を返すとクローズが取り消される実装があるため、必ず true を返す / Always return true so the close goes through */
                return true;
            };
        }
    }

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
    } else {
        showDistributeDialog();
    }

}());
