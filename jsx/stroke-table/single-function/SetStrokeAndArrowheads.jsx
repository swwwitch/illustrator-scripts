#target illustrator
#targetengine "SetStrokeAndArrowheadsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの線幅と矢印（始点／終点の形状・倍率・先端位置）をまとめて設定します。
矢印はDOMから操作できないため、一時アクションを生成して実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SetStrokeAndArrowheads.md

### Overview

Sets the stroke weight and the arrowheads — start and end shape, scale and tip position — on the selected objects.
Arrowheads cannot be reached from the DOM, so a temporary action is generated and played instead.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SetStrokeAndArrowheads.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SetStrokeAndArrowheads";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SetStrokeAndArrowheads.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SetStrokeAndArrowheads.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ユーザー設定 / User settings
    // =========================================

    /* 一時アクション / Temporary action */
    var ACTION_SET_NAME  = "SwwwitchTempStrokeSet";
    var ACTION_NAME      = "SwwwitchTempStroke";

    /* 既定値 / Defaults */
    var DEFAULT_STROKE_WIDTH = 5;            /* 線幅の初期値 / initial stroke width */
    var DEFAULT_ARROW_SCALE  = 100;          /* 倍率の初期値 / initial arrowhead scale */
    var ARROW_COUNT          = 39;           /* 矢印の種類数 / number of arrowhead presets */

    /* パラメータキー / Parameter keys（記録した .aia から採取 / taken from a recorded .aia） */
    var KEY_STROKE_WIDTH  = 2003072104;      /* 線幅 / stroke width */
    var KEY_CAP           = 1667330094;      /* 線端 / cap */
    var KEY_JOIN          = 1785686382;      /* 角の形状 / join */
    var KEY_DASH_INT      = 1684825454;      /* 破線（整数）/ dash (integer) */
    var KEY_DASH_BOOL     = 1684104298;      /* 破線（真偽）/ dash (boolean) */
    var KEY_ARROW_HEAD_1  = 1634231345;      /* ahd1: 始点の形状 / start arrowhead */
    var KEY_ARROW_HEAD_2  = 1634231346;      /* ahd2: 終点の形状 / end arrowhead */
    var KEY_ARROW_SCALE_1 = 1634951985;      /* asc1: 始点の倍率（推定値）/ start scale (estimated) */
    var KEY_ARROW_SCALE_2 = 1634951986;      /* asc2: 終点の倍率 / end scale */
    var KEY_ARROW_ALIGN   = 1634230636;      /* ahal: 矢印の配置 / tip alignment */
    var KEY_ALIGN         = 1634494318;      /* algn: 線の位置 / stroke alignment */

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

    /* UIレイアウト：コントロールの寸法 / UI layout: control metrics */
    var LABEL_WIDTH      = 40;               /* 行頭ラベルの幅 / row label width */
    var SHAPE_LIST_WIDTH = 90;               /* 形状プルダウンの幅 / shape dropdown width */
    var FIELD_CHARACTERS = 4;                /* 数値入力欄の文字数 / numeric field width */

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
            title: { ja: "線と矢印を設定", en: "Set Stroke and Arrowheads" }
        },
        panel: {
            stroke: { ja: "線", en: "Stroke" },
            arrow: { ja: "矢印", en: "Arrowheads" },
            start: { ja: "始点", en: "Start" },
            end: { ja: "終点", en: "End" },
            tipAlign: { ja: "先端位置", en: "Tip Alignment" }
        },
        field: {
            strokeWidth: { ja: "線幅", en: "Weight" },
            shape: { ja: "形状", en: "Shape" },
            scale: { ja: "倍率", en: "Scale" },
            unitPt: { ja: "pt", en: "pt" },
            unitPercent: { ja: "%", en: "%" }
        },
        tooltip: {
            strokeWidth: { ja: "線の太さです。", en: "Weight of the stroke." },
            linkEnds:    { ja: "始点と終点の矢印を同じ設定にします。", en: "Uses the same arrowhead settings at both ends." },
            shape:       { ja: "この端に付ける矢印の形です。", en: "The arrowhead shape used at this end." },
            scale:       { ja: "矢印の大きさ（％）です。", en: "Size of the arrowhead, in percent." },
            align:       { ja: "矢印をパスの端にどう合わせるかです。", en: "How the arrowhead lines up with the end of the path." },
            preview:     { ja: "結果を画面で確認します。キャンセルすると元に戻ります。", en: "Shows the result on the canvas. Cancel restores the original state." },
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
        checkbox: {
            linkEnds: { ja: "始点と終点を連動", en: "Link start and end" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        arrow: {
            none: { ja: "[なし]", en: "[None]" },
            prefix: { ja: "矢印 ", en: "Arrow " }
        },
        /* 先端位置：label は UI 表示用、name はアクションに埋め込む Illustrator の表示名 */
        /* Tip alignment: label is for the UI, name is the Illustrator label embedded in the action */
        tipAlign: {
            atEndLabel: {
                ja: "矢印の先端をパスの終点に配置",
                en: "Place arrow tip at end of path"
            },
            atEndName: {
                ja: "パスの終点に配置",
                en: "Place Arrow Tip At End of Path"
            },
            beyondEndLabel: {
                ja: "矢印の先端をパスの終点から配置",
                en: "Extend arrow tip beyond end of path"
            },
            beyondEndName: {
                ja: "パスの終点から配置",
                en: "Extend Arrow Tip Beyond End of Path"
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            },
            noSelection: {
                ja: "オブジェクトを選択してください。",
                en: "Please select at least one object."
            },
            invalidWidth: {
                ja: "線幅には 0 以上の数値を入力してください。",
                en: "Enter a stroke weight of 0 or greater."
            },
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." }
        }
    };

    // =========================================
    // 単位 / Units
    // =========================================

    var UNIT_POINT = 592476268;              /* ポイント / point（parameter /unit） */

    /* 先端位置の選択肢 / Tip alignment options（ahal の enumerated 値 / enumerated values） */
    var ARROW_ALIGN_OPTIONS = [
        { label: getLabel("tipAlign.atEndLabel"),     name: getLabel("tipAlign.atEndName"),     value: 0 },
        { label: getLabel("tipAlign.beyondEndLabel"), name: getLabel("tipAlign.beyondEndName"), value: 1 }
    ];

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* 線を持てるオブジェクトを再帰的に集める / Collect strokable items recursively */
    function collectStrokeTargets(items) {
        var targets = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            try {
                if (item.typename === "PathItem") {
                    targets.push(item);
                } else if (item.typename === "CompoundPathItem") {
                    targets = targets.concat(collectStrokeTargets(item.pathItems));
                } else if (item.typename === "GroupItem") {
                    targets = targets.concat(collectStrokeTargets(item.pageItems));
                }
            } catch (e) {}
        }
        return targets;
    }

    /*
     * プレビュー制御 / Preview controller
     * - 線幅は DOM で即時プレビュー（元の値を保存して書き戻すので undo に依存しない）
     * - 矢印は DOM に存在しないため、アクション実行 + app.undo() でしかプレビューできない
     * - Stroke width previews through the DOM; arrowheads require play + undo
     */
    function createPreviewController() {
        var targets = [];
        var originalStroked = [];
        var originalWidths = [];
        var isActionPreviewApplied = false;

        /* 現在の線設定を退避する / Capture the current stroke state */
        function capture() {
            targets = collectStrokeTargets(app.activeDocument.selection);
            originalStroked = [];
            originalWidths = [];
            for (var i = 0; i < targets.length; i++) {
                try {
                    originalStroked.push(targets[i].stroked);
                    originalWidths.push(targets[i].stroked ? targets[i].strokeWidth : 1);
                } catch (e) {
                    originalStroked.push(false);
                    originalWidths.push(1);
                }
            }
        }

        /* 退避した線設定を書き戻す / Restore the captured stroke state */
        function restoreStroke() {
            for (var i = 0; i < targets.length; i++) {
                try {
                    targets[i].stroked = originalStroked[i];
                    if (originalStroked[i]) targets[i].strokeWidth = originalWidths[i];
                } catch (e) {}
            }
        }

        /* アクションによるプレビューを undo で取り消す / Undo the action based preview */
        function clearActionPreview() {
            if (!isActionPreviewApplied) return;
            try { app.undo(); } catch (e) {}
            isActionPreviewApplied = false;
            /* undo でオブジェクト参照が無効になるため取り直す / References die on undo, re-capture */
            capture();
        }

        capture();

        return {
            /* 線幅のみ DOM で即時プレビュー / Preview stroke width only, via the DOM */
            previewStrokeWidth: function (strokeWidth) {
                clearActionPreview();
                if (isNaN(strokeWidth) || strokeWidth < 0) return;
                for (var i = 0; i < targets.length; i++) {
                    try {
                        targets[i].stroked = true;
                        targets[i].strokeWidth = strokeWidth;
                    } catch (e) {}
                }
                app.redraw();
            },

            /* 矢印を含む全設定をアクションで 1 回だけプレビュー / Preview all settings via the action */
            previewAction: function (settings) {
                clearActionPreview();
                restoreStroke();
                playStrokeAction(settings);
                isActionPreviewApplied = true;
                app.redraw();
            },

            /* アクションによるプレビューだけを取り消す / Clear only the action based preview */
            clearAction: function () {
                if (!isActionPreviewApplied) return;
                clearActionPreview();
                app.redraw();
            },

            /* すべてのプレビューを取り消して元の状態に戻す / Revert everything */
            reset: function () {
                clearActionPreview();
                restoreStroke();
                app.redraw();
            }
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名の行に、∧∨と入力欄を隙間0で突き合わせて追加する。↑↓キーも∧∨と同じ処理で増減する
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - addStepper() に渡す増減設定（min など）
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSteppedInput(parentRow, initialText, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = FIELD_CHARACTERS;
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 入力欄と∧∨の有効／無効をまとめて切り替え、∧∨を描き直す
     * @param {EditText} numberInput - addSteppedInput() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * ∧∨・↑↓キーで値を変えたあと、手入力と同じ処理を呼ぶ（プログラムからの書き換えではイベントが発火しないため）
     * @param {EditText} numberInput - 値を変えた入力欄
     * @returns {void}
     */
    function notifySteppedInput(numberInput) {
        if (typeof numberInput.onChanging === "function") numberInput.onChanging();
        /* フォーカスが無いまま∧∨で変えると onChange が来ないので、確定側のプレビューもここで更新する
           The stepper can change the value without focus, so run the commit handler too */
        if (typeof numberInput.onChange === "function") numberInput.onChange();
    }

    /* 矢印名の一覧を作る / Build the arrowhead name list */
    function buildArrowNameList() {
        var names = [getLabel("arrow.none")];
        for (var i = 1; i <= ARROW_COUNT; i++) {
            names.push(getLabel("arrow.prefix") + i);
        }
        return names;
    }

    /*
     * 設定用ダイアログを表示する / Show the settings dialog
     * @returns {object|null} 入力値。キャンセル時は null / entered values, or null when cancelled
     */
    function showSettingsDialog() {
        var arrowNames = buildArrowNameList();
        var preview = createPreviewController();

        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 線幅 / Stroke width */
        var strokePanel = dialog.add("panel", undefined, getLabel("panel.stroke"));
        setupPanel(strokePanel, 6);

        var strokeRow = strokePanel.add("group");
        setupRow(strokeRow, "left", 8);
        strokeRow.add("statictext", undefined, labelText("field.strokeWidth"));
        var strokeWidthInput = addSteppedInput(strokeRow, String(DEFAULT_STROKE_WIDTH), { min: 0, onStep: notifySteppedInput });
        strokeWidthInput.helpTip = getLabel("tooltip.strokeWidth");
        strokeRow.add("statictext", undefined, getLabel("field.unitPt"));

        /* 矢印 / Arrowheads */
        var arrowPanel = dialog.add("panel", undefined, getLabel("panel.arrow"));
        setupPanel(arrowPanel, 6);

        var linkCheckbox = arrowPanel.add("checkbox", undefined, getLabel("checkbox.linkEnds"));
        linkCheckbox.helpTip = getLabel("tooltip.linkEnds");
        linkCheckbox.value = false;
        linkCheckbox.alignment = "left";

        /* 始点／終点を 2 カラムで並べる / Lay out start and end in two columns */
        var arrowColumns = arrowPanel.add("group");
        arrowColumns.orientation = "row";
        arrowColumns.alignChildren = ["fill", "top"];
        arrowColumns.spacing = COLUMN_SPACING;

        /* 形状・倍率をひと組にしたカラムを作る / Build a shape + scale column */
        function addArrowColumn(parent, title) {
            var panel = parent.add("panel", undefined, title);
            setupPanel(panel, 6);

            var shapeRow = panel.add("group");
            setupRow(shapeRow, "left", 8);
            var shapeLabel = shapeRow.add("statictext", undefined, labelText("field.shape"));
            shapeLabel.preferredSize.width = LABEL_WIDTH;
            var shapeList = shapeRow.add("dropdownlist", undefined, arrowNames);
            shapeList.helpTip = getLabel("tooltip.shape");
            shapeList.selection = 0;
            shapeList.preferredSize.width = SHAPE_LIST_WIDTH;

            var scaleRow = panel.add("group");
            setupRow(scaleRow, "left", 8);
            var scaleLabel = scaleRow.add("statictext", undefined, labelText("field.scale"));
            scaleLabel.preferredSize.width = LABEL_WIDTH;
            /* 0 以下は既定の倍率に戻されるので、下限は 1 / values of 0 or less fall back to the default, so the minimum is 1 */
            var scaleInput = addSteppedInput(scaleRow, String(DEFAULT_ARROW_SCALE), { min: 1, onStep: notifySteppedInput });
            scaleInput.helpTip = getLabel("tooltip.scale");
            scaleRow.add("statictext", undefined, getLabel("field.unitPercent"));

            return { shapeList: shapeList, scaleInput: scaleInput };
        }

        var startColumn = addArrowColumn(arrowColumns, getLabel("panel.start"));
        var endColumn   = addArrowColumn(arrowColumns, getLabel("panel.end"));

        var startShapeList  = startColumn.shapeList;
        var startScaleInput = startColumn.scaleInput;
        var endShapeList    = endColumn.shapeList;
        var endScaleInput   = endColumn.scaleInput;

        /* 先端位置 / Tip alignment */
        var alignPanel = dialog.add("panel", undefined, getLabel("panel.tipAlign"));
        setupPanel(alignPanel, 6);

        var alignRadios = [];
        for (var a = 0; a < ARROW_ALIGN_OPTIONS.length; a++) {
            var radio = alignPanel.add("radiobutton", undefined, ARROW_ALIGN_OPTIONS[a].label);
            radio.helpTip = getLabel("tooltip.align");
            /* ボタン類はパネル幅いっぱいに広げない / Keep buttons at their natural width */
            radio.alignment = "left";
            alignRadios.push(radio);
        }
        alignRadios[0].value = true;

        /* ボタン（左：プレビュー／中央：スペーサー／右：キャンセル・OK） */
        /* Buttons (left: preview, center: spacer, right: cancel and OK) */
        var buttonRow = addButtonRow(dialog);
        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.value = false;
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });

        /* 連動時は始点の値を終点にコピー / Mirror start values onto end when linked */
        function applyLink() {
            if (!linkCheckbox.value) return;
            endShapeList.selection = startShapeList.selection.index;
            endScaleInput.text = startScaleInput.text;
        }

        /* 形状が [なし] のときは倍率を無効化。連動時は終点側を無効化 */
        /* Disable scale when the shape is [None]; disable the end column when linked */
        function syncEnabled() {
            var isLinked = linkCheckbox.value;
            var hasArrow = (startShapeList.selection.index !== 0) || (endShapeList.selection.index !== 0);

            setSteppedInputEnabled(startScaleInput, startShapeList.selection.index !== 0);
            endShapeList.enabled = !isLinked;
            setSteppedInputEnabled(endScaleInput, !isLinked && (endShapeList.selection.index !== 0));

            /* 矢印がひとつも無ければ先端位置は無効 / No arrowheads means no tip alignment */
            alignPanel.enabled = hasArrow;
        }

        /* 選択中の先端位置を返す / Return the selected tip alignment */
        function getSelectedAlign() {
            for (var i = 0; i < alignRadios.length; i++) {
                if (alignRadios[i].value) return ARROW_ALIGN_OPTIONS[i];
            }
            return ARROW_ALIGN_OPTIONS[0];
        }

        /* 連動と有効／無効をまとめて更新 / Refresh linking and enabled states */
        function refresh() {
            applyLink();
            syncEnabled();
        }

        /* 入力値を読み取る。線幅が不正なら null / Read input; null when the width is invalid */
        function readSettings() {
            var strokeWidth = parseFloat(strokeWidthInput.text);
            var startScale  = parseFloat(startScaleInput.text);
            var endScale    = parseFloat(endScaleInput.text);

            if (isNaN(strokeWidth) || strokeWidth < 0) return null;
            if (isNaN(startScale) || startScale <= 0) startScale = DEFAULT_ARROW_SCALE;
            if (isNaN(endScale)   || endScale   <= 0) endScale   = DEFAULT_ARROW_SCALE;

            return {
                strokeWidth: strokeWidth,
                startArrow: startShapeList.selection.text,
                startScale: startScale,
                endArrow: endShapeList.selection.text,
                endScale: endScale,
                arrowAlign: getSelectedAlign()
            };
        }

        /* プレビューを現在の入力値で更新する（矢印を含むためアクションを実行） */
        /* Refresh the preview with the current values (plays the action for arrowheads) */
        function updatePreview() {
            if (!previewCheckbox.value) {
                preview.reset();
                return;
            }
            var settings = readSettings();
            if (!settings) return;
            preview.previewAction(settings);
        }

        /* 値を変更したらアクションのプレビューを取り消す（確定時に貼り直す） */
        /* Drop the action preview while editing; it is reapplied on commit */
        function invalidatePreview() {
            if (previewCheckbox.value) preview.clearAction();
        }

        previewCheckbox.onClick = updatePreview;

        linkCheckbox.onClick = function () {
            refresh();
            updatePreview();
        };
        startShapeList.onChange = function () {
            refresh();
            updatePreview();
        };
        endShapeList.onChange = function () {
            syncEnabled();
            updatePreview();
        };

        startScaleInput.onChanging = function () {
            if (linkCheckbox.value) endScaleInput.text = startScaleInput.text;
            invalidatePreview();
        };
        startScaleInput.onChange = updatePreview;
        endScaleInput.onChanging = invalidatePreview;
        endScaleInput.onChange = updatePreview;

        /* 線幅は入力中も DOM で即時プレビュー、確定時にアクションで貼り直す */
        /* Stroke width previews live through the DOM, then via the action on commit */
        strokeWidthInput.onChanging = function () {
            if (previewCheckbox.value) preview.previewStrokeWidth(parseFloat(strokeWidthInput.text));
        };
        strokeWidthInput.onChange = updatePreview;

        for (var r = 0; r < alignRadios.length; r++) {
            alignRadios[r].onClick = updatePreview;
        }

        refresh();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var isAccepted = (dialog.show() === 1);

        /* プレビューを必ず取り消してから本適用に進む / Always revert the preview before applying */
        preview.reset();

        if (!isAccepted) return null;

        var settings = readSettings();
        if (!settings) {
            alert(getLabel("alert.invalidWidth"));
            return null;
        }
        return settings;
    }

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

    /*
     * 線幅・矢印を設定する一時アクションのソースを生成する
     * Build the source of a temporary action that sets stroke width and arrowheads
     * @param {string} setName    アクションセット名（ASCII 推奨）/ action set name
     * @param {string} actionName アクション名（ASCII 推奨）/ action name
     * @param {object} settings   showSettingsDialog() の戻り値 / result of showSettingsDialog()
     */
    function buildStrokeActionSource(setName, actionName, settings) {
        return ''
            + '/version 3\n'
            + buildNameLine(setName)
            + '/isOpen 1\n'
            + '/actionCount 1\n'
            + '/action-1 {\n'
            + '\t' + buildNameLine(actionName)
            + '\t/keyIndex 0\n'
            + '\t/colorIndex 0\n'
            + '\t/isOpen 1\n'
            + '\t/eventCount 1\n'
            + '\t/event-1 {\n'
            + '\t\t/useRulersIn1stQuadrant 0\n'
            + '\t\t/internalName (ai_plugin_setStroke)\n'
            + '\t\t/localizedName [ 0 \n\t\t]\n'
            + '\t\t/isOpen 1\n'
            + '\t\t/isOn 1\n'
            + '\t\t/hasDialog 0\n'
            + '\t\t/parameterCount 11\n'
            /* 可変 / variable */
            + buildUnitRealParam(1, KEY_STROKE_WIDTH, settings.strokeWidth, UNIT_POINT)
            + buildUStrParam(6, KEY_ARROW_HEAD_1, settings.startArrow)
            + buildUStrParam(7, KEY_ARROW_HEAD_2, settings.endArrow)
            + buildRealParam(8, KEY_ARROW_SCALE_1, settings.startScale)
            + buildRealParam(9, KEY_ARROW_SCALE_2, settings.endScale)
            + buildEnumParamByName(10, KEY_ARROW_ALIGN, settings.arrowAlign.name, settings.arrowAlign.value) /* 先端位置 / tip alignment */
            /* 以下は記録した .aia のまま / recorded as-is */
            + buildEnumParam(2, KEY_CAP,  'e4b8b8e59e8be7b79ae7abaf', 12, 1)             /* 線端: 丸型線端 / round cap */
            + buildEnumParam(3, KEY_JOIN, 'e383a9e382a6e383b3e38389e7b590e59088', 18, 1) /* 角の形状: ラウンド結合 / round join */
            + buildIntParam(4,  KEY_DASH_INT,  0)
            + buildBoolParam(5, KEY_DASH_BOOL, 0)
            + buildEnumParam(11, KEY_ALIGN, 'e4b8ade5a4ae', 6, 0)                        /* 線の位置: 中央 / center */
            + '\t}\n'
            + '}\n';
    }

    // --- パラメータ組み立てヘルパー / parameter builders ---

    /* パラメータブロックの外枠を作る / Wrap a parameter body in its block */
    function buildParamBlock(index, key, body) {
        return ''
            + '\t\t/parameter-' + index + ' {\n'
            + '\t\t\t/key ' + key + '\n'
            + '\t\t\t/showInPalette 4294967295\n'
            + body
            + '\t\t}\n';
    }

    /* 単位付き実数パラメータ / A unit real parameter */
    function buildUnitRealParam(index, key, value, unitCode) {
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (unit real)\n'
            + '\t\t\t/value ' + toRealString(value) + '\n'
            + '\t\t\t/unit ' + unitCode + '\n');
    }

    /* 列挙パラメータ（hex 指定）/ An enumerated parameter from a hex name */
    function buildEnumParam(index, key, nameHex, byteLength, value) {
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (enumerated)\n'
            + '\t\t\t/name [ ' + byteLength + ' \n\t\t\t\t' + nameHex + '\n\t\t\t]\n'
            + '\t\t\t/value ' + value + '\n');
    }

    /* enumerated パラメータを表示名から組み立てる（hex は自動生成）/ Enumerated parameter from a display name */
    function buildEnumParamByName(index, key, name, value) {
        var hex = toActionHex(name);
        return buildEnumParam(index, key, hex, hex.length / 2, value);
    }

    /* 整数パラメータ / An integer parameter */
    function buildIntParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (integer)\n\t\t\t/value ' + value + '\n');
    }

    /* 真偽値パラメータ / A boolean parameter */
    function buildBoolParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (boolean)\n\t\t\t/value ' + (value ? 1 : 0) + '\n');
    }

    /* 実数パラメータ / A real parameter */
    function buildRealParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (real)\n\t\t\t/value ' + toRealString(value) + '\n');
    }

    /* Unicode 文字列パラメータ / A ustring parameter */
    function buildUStrParam(index, key, value) {
        var hex = toActionHex(value);
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (ustring)\n'
            + '\t\t\t/value [ ' + (hex.length / 2) + ' \n\t\t\t\t' + hex + '\n\t\t\t]\n');
    }

    // --- 文字列・数値ユーティリティ / utilities ---

    /* アクションセット名・アクション名の /name 行を作る / Build a /name line */
    function buildNameLine(name) {
        var hex = toActionHex(name);
        return '/name [ ' + (hex.length / 2) + ' \n\t' + hex + '\n]\n';
    }

    /* 5 → "5.0" のように必ず小数点を含む文字列にする / Force a decimal point */
    function toRealString(value) {
        var text = String(Number(value));
        if (text.indexOf('.') === -1 && text.indexOf('e') === -1) text += '.0';
        return text;
    }

    // =========================================
    // 一時アクション実行 / Temporary action playback
    // =========================================

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /**
     * ［線］の設定アクションを1回実行する。失敗は従来どおり例外で伝える
     * @param {object} settings - showSettingsDialog() の戻り値
     * @returns {void}
     */
    function playStrokeAction(settings) {
        var source = buildStrokeActionSource(ACTION_SET_NAME, ACTION_NAME, settings);
        if (!runTemporaryAction(source, ACTION_SET_NAME, ACTION_NAME)) {
            throw new Error(getLabel("alert.actionFailed"));
        }
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    /* 前提条件を確認し、ダイアログの入力をアクションとして適用する */
    /* Verify preconditions, then apply the dialog input through the action */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        if (app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var settings = showSettingsDialog();
        if (!settings) return;

        playStrokeAction(settings);
    }

    main();

})();
