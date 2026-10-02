#target illustrator
#targetengine "SplitSpreadToSingleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

見開きページ相当のオブジェクトを検出し、左右2つの片ページに分割します。
対象は選択オブジェクトのみ、またはドキュメント内のすべてから選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitSpreadToSingle.md

### Overview

Detects spread-like objects and splits them into left and right single pages.
You can process only the selection, or every matching object in the document.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitSpreadToSingle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitSpreadToSingle";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitSpreadToSingle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitSpreadToSingle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 片ページ／見開きと判定する幅の許容差（基準ページ幅に対する比率） / Width tolerance for single / spread detection (ratio of the reference page width) */
    var PAGE_WIDTH_TOLERANCE = 0.15;

    /* 再配置の間隔の初期値 / Default rearrangement spacing */
    var DEFAULT_SPACING_TEXT = "20";

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

    /* 間隔欄の幅（文字数） / Width of the spacing fields in characters */
    var SPACING_FIELD_CHARACTERS = 5;

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
    // 識別文字列 / Markers
    // =========================================

    /* 分割したグループの note に書く識別子 / Identifier written to the note of split groups */
    var SPLIT_GROUP_NOTE_PREFIX = "__SplitSpreadToSingle__";

    /* 再配置中にアートボードごとのオブジェクトへ一時的に付ける目印 / Temporary marker added to notes while rearranging */
    var REARRANGE_NOTE_MARKER = "__SplitSpreadToSingle_Rearrange__";

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

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () {
            numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
            numberInput.active = true;
        });

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

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "見開きページを片ページに", en: "Split Spread Pages into Single Pages" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            evenPage: { ja: "偶数ページ", en: "Even Pages" },
            postProcess: { ja: "後処理", en: "Post-Process" }
        },
        radio: {
            selectionOnly: { ja: "選択したオブジェクトのみ", en: "Selected Objects Only" },
            all: { ja: "すべて", en: "All" },
            sideRight: { ja: "右", en: "Right" },
            sideLeft: { ja: "左", en: "Left" }
        },
        checkbox: {
            renameArtboards: { ja: "アートボード名を連番でリネーム", en: "Rename Artboards Sequentially" },
            rearrangeArtboards: { ja: "アートボードの再配置", en: "Rearrange Artboards" }
        },
        fieldLabel: {
            spacingHorizontal: { ja: "左右", en: "Horizontal" },
            spacingVertical: { ja: "上下", en: "Vertical" }
        },
        tooltip: {
            modeSelection: { ja: "選択しているアートボードだけを分割します。", en: "Splits only the selected artboards." },
            modeAll: { ja: "ドキュメント内のすべてのアートボードを分割します。", en: "Splits every artboard in the document." },
            sideRight: {
                ja: "偶数ページを見開きの右側として扱います。左綴じ（横書き）向けです。",
                en: "Treats even pages as the right-hand side of the spread, for left-bound documents."
            },
            sideLeft: {
                ja: "偶数ページを見開きの左側として扱います。右綴じ（縦書き）向けです。",
                en: "Treats even pages as the left-hand side of the spread, for right-bound documents."
            },
            rename: { ja: "分割後のアートボードに、ページ番号で名前を付け直します。", en: "Renames the resulting artboards with their page numbers." },
            rearrange: {
                ja: "分割後のアートボードを並べ直します。間隔は下の欄で指定します。",
                en: "Lays the resulting artboards out again. The fields below set the spacing."
            },
            spacingHorizontal: { ja: "並べ直すときの横方向の間隔です。", en: "Horizontal spacing used when the artboards are laid out." },
            spacingVertical: { ja: "並べ直すときの縦方向の間隔です。", en: "Vertical spacing used when the artboards are laid out." },
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
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        generatedName: {
            leftHalfGroup: { ja: "左半分", en: "Left_Half" },
            rightHalfGroup: { ja: "右半分", en: "Right_Half" },
            artboardPrefix: { ja: "アートボード ", en: "Artboard " }
        },
        alert: {
            invalidSpacing: { ja: "アートボードの再配置の間隔には数値を入力してください。", en: "Enter numeric values for artboard rearrangement spacing." },
            noSpreadFoundAll: {
                ja: "見開きページ相当の PlacedItem / GroupItem / RasterItem が見つかりませんでした。",
                en: "No spread-like PlacedItem / GroupItem / RasterItem was found."
            },
            selectSpreadObject: { ja: "分割したい見開きオブジェクトを選択してください。", en: "Select a spread object to split." },
            selectionIsSingle: {
                ja: "選択オブジェクトは片ページ相当です。見開きページ相当のオブジェクトを選択してください。",
                en: "The selected object looks like a single page. Select a spread-like object."
            },
            selectValidSpread: {
                ja: "見開きページ相当の PlacedItem / GroupItem / RasterItem を選択してください。",
                en: "Select a spread-like PlacedItem / GroupItem / RasterItem."
            }
        }
    };

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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見出し付きパネルを追加する（ラジオボタン・チェックボックスが並ぶので間隔を詰める）
     * @param {Window} parentWindow - 追加先
     * @param {string} titleLabelPath - パネルタイトルの LABELS パス
     * @param {string} orientation - "column" または "row"
     * @returns {Panel} 追加したパネル
     */
    function addTitledPanel(parentWindow, titleLabelPath, orientation) {
        var newPanel = parentWindow.add("panel", undefined, getLabel(titleLabelPath));
        setupPanel(newPanel, 6);
        newPanel.orientation = orientation;
        return newPanel;
    }

    /**
     * tooltip 付きのコントロールを追加する
     * @param {Object} parentGroup - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox" など
     * @param {string} textLabelPath - 表示文字の LABELS パス
     * @param {string} tooltipLabelPath - tooltip の LABELS パス
     * @returns {Object} 追加したコントロール
     */
    function addControlWithTip(parentGroup, controlType, textLabelPath, tooltipLabelPath) {
        var newControl = parentGroup.add(controlType, undefined, getLabel(textLabelPath));
        newControl.helpTip = getLabel(tooltipLabelPath);
        return newControl;
    }

    /**
     * 間隔の入力欄を項目名つきで追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelKey - LABELS.fieldLabel / LABELS.tooltip のキー
     * @returns {{label: StaticText, field: EditText}} 項目名と入力欄
     */
    function addSpacingField(parentGroup, labelKey) {
        var spacingLabel = parentGroup.add("statictext", undefined, getLabel("fieldLabel." + labelKey));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = parentGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var spacingField;
        var stepperGroup = addStepper(stepperInputGroup, function () { return spacingField; }, {});
        spacingField = stepperInputGroup.add("edittext", undefined, DEFAULT_SPACING_TEXT);
        spacingField.helpTip = getLabel("tooltip." + labelKey);
        spacingField.characters = SPACING_FIELD_CHARACTERS;
        /* setSteppedFieldEnabled() で項目名・∧∨ごと切り替えられるようにする / for setSteppedFieldEnabled() */
        spacingField.fieldLabel = spacingLabel;
        spacingField.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(spacingField, stepperGroup);
        return { label: spacingLabel, field: spacingField };
    }

    /**
     * 設定ダイアログを表示し、選ばれた設定を返す
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showSplitDialog() {
        var splitDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(splitDialog);

        var targetPanel = addTitledPanel(splitDialog, "panel.target", "column");
        addControlWithTip(targetPanel, "radiobutton", "radio.selectionOnly", "tooltip.modeSelection");
        var rbAll = addControlWithTip(targetPanel, "radiobutton", "radio.all", "tooltip.modeAll");
        rbAll.value = true;

        var evenPagePanel = addTitledPanel(splitDialog, "panel.evenPage", "row");
        var rbEvenRight = addControlWithTip(evenPagePanel, "radiobutton", "radio.sideRight", "tooltip.sideRight");
        var rbEvenLeft = addControlWithTip(evenPagePanel, "radiobutton", "radio.sideLeft", "tooltip.sideLeft");
        rbEvenLeft.value = true;

        var postProcessPanel = addTitledPanel(splitDialog, "panel.postProcess", "column");
        var cbRenameArtboards = addControlWithTip(postProcessPanel, "checkbox", "checkbox.renameArtboards", "tooltip.rename");
        cbRenameArtboards.value = true;
        var cbRearrangeArtboards = addControlWithTip(postProcessPanel, "checkbox", "checkbox.rearrangeArtboards", "tooltip.rearrange");
        cbRearrangeArtboards.value = true;

        var spacingGroup = postProcessPanel.add("group");
        spacingGroup.orientation = "row";
        spacingGroup.alignChildren = ["left", "center"];
        var horizontalSpacing = addSpacingField(spacingGroup, "spacingHorizontal");
        var verticalSpacing = addSpacingField(spacingGroup, "spacingVertical");

        /* 間隔欄は再配置がオンのときだけ有効 / Spacing fields apply only when rearranging */
        function updateSpacingEnabled() {
            var rearrangeEnabled = cbRearrangeArtboards.value;
            setSteppedFieldEnabled(horizontalSpacing.field, rearrangeEnabled);
            setSteppedFieldEnabled(verticalSpacing.field, rearrangeEnabled);
        }
        cbRearrangeArtboards.onClick = updateSpacingEnabled;
        updateSpacingEnabled();

        var buttonRow = addButtonRow(splitDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        prepareDialogWindow(splitDialog, SCRIPT_NAME);
        if (splitDialog.show() !== 1) return null;

        return {
            processAll: rbAll.value,
            evenOnRight: rbEvenRight.value,
            renameArtboards: cbRenameArtboards.value,
            rearrangeArtboards: cbRearrangeArtboards.value,
            spacingHorizontal: parseFloat(horizontalSpacing.field.text),
            spacingVertical: parseFloat(verticalSpacing.field.text)
        };
    }

    // =========================================
    // 見開きの判定 / Spread detection
    // =========================================

    /**
     * 分割できる種類のオブジェクトか（配置画像・グループ・ラスター画像）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isTargetItem(pageItem) {
        return pageItem && (
            pageItem.typename === "PlacedItem" ||
            pageItem.typename === "GroupItem" ||
            pageItem.typename === "RasterItem"
        );
    }

    /**
     * ドキュメント内の分割できる種類のオブジェクトをすべて集める
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 対象のオブジェクト
     */
    function collectTargetItems(doc) {
        var targetItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var pageItem = doc.pageItems[i];
            if (isTargetItem(pageItem)) {
                targetItems.push(pageItem);
            }
        }
        return targetItems;
    }

    /**
     * 最も重なり面積の大きいアートボードを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {number} アートボードの番号（重ならなければ -1）
     */
    function getPrimaryArtboardIndex(doc, pageItem) {
        var bounds = pageItem.geometricBounds;
        var bestIndex = -1;
        var bestArea = 0;

        for (var i = 0; i < doc.artboards.length; i++) {
            var abRect = doc.artboards[i].artboardRect;
            var overlapWidth = Math.min(bounds[2], abRect[2]) - Math.max(bounds[0], abRect[0]);
            var overlapHeight = Math.min(bounds[1], abRect[1]) - Math.max(bounds[3], abRect[3]);

            if (overlapWidth > 0 && overlapHeight > 0) {
                var overlapArea = overlapWidth * overlapHeight;
                if (overlapArea > bestArea) {
                    bestArea = overlapArea;
                    bestIndex = i;
                }
            }
        }
        return bestIndex;
    }

    /**
     * ドキュメント内の基準ページ幅（最小アートボード幅）を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {number} 基準ページ幅（アートボードが無ければ 0）
     */
    function getReferencePageWidth(doc) {
        var minWidth = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            var abRect = doc.artboards[i].artboardRect;
            var abWidth = abRect[2] - abRect[0];
            if (abWidth <= 0) continue;
            if (minWidth === null || abWidth < minWidth) {
                minWidth = abWidth;
            }
        }
        return (minWidth === null) ? 0 : minWidth;
    }

    /**
     * 幅の比率が目標にほぼ一致するか
     * @param {number} ratio - 幅の比率
     * @param {number} targetRatio - 目標の比率（1 = 片ページ、2 = 見開き）
     * @returns {boolean} 許容差の範囲内なら true
     */
    function isNearRatio(ratio, targetRatio) {
        return Math.abs(ratio - targetRatio) <= PAGE_WIDTH_TOLERANCE;
    }

    /**
     * 指定アートボードが片ページ相当か見開き相当かを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {number} artboardIndex - アートボードの番号
     * @param {number} referencePageWidth - 基準ページ幅
     * @returns {string} "single" / "spread" / "other"
     */
    function getArtboardType(doc, artboardIndex, referencePageWidth) {
        var abRect = doc.artboards[artboardIndex].artboardRect;
        var ratio = (abRect[2] - abRect[0]) / referencePageWidth;

        if (isNearRatio(ratio, 1)) return "single";
        if (isNearRatio(ratio, 2)) return "spread";
        return "other";
    }

    /**
     * オブジェクト幅と基準ページ幅から片ページ／見開き／その他を判定する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {{artboardIndex: number, pageType: string, artboardType: string}} 判定結果
     */
    function getPageTypeInfo(doc, pageItem) {
        var artboardIndex = getPrimaryArtboardIndex(doc, pageItem);
        var referencePageWidth = (artboardIndex < 0) ? 0 : getReferencePageWidth(doc);
        if (referencePageWidth <= 0) {
            return { artboardIndex: artboardIndex, pageType: "other", artboardType: "other" };
        }

        var bounds = pageItem.geometricBounds;
        var itemWidth = bounds[2] - bounds[0];
        var abRect = doc.artboards[artboardIndex].artboardRect;
        var ratio = itemWidth / referencePageWidth;
        var fitToArtboardRatio = itemWidth / (abRect[2] - abRect[0]);
        var artboardType = getArtboardType(doc, artboardIndex, referencePageWidth);
        var pageType = "other";

        /* 見開きアートボード上でアートボード幅にほぼ一致するオブジェクトは見開き扱い / Treat objects that match artboard width on spread artboards as spread */
        if (artboardType === "spread" && isNearRatio(fitToArtboardRatio, 1)) {
            pageType = "spread";
        } else if (artboardType === "single" && isNearRatio(ratio, 1)) {
            pageType = "single";
        } else if (artboardType === "spread" && isNearRatio(ratio, 2)) {
            pageType = "spread";
        }

        return { artboardIndex: artboardIndex, pageType: pageType, artboardType: artboardType };
    }

    /**
     * 見開きアートボード上の見開きオブジェクトか
     * @param {Object} pageTypeInfo - getPageTypeInfo() の結果
     * @returns {boolean} 分割の対象なら true
     */
    function isSpreadOnSpreadArtboard(pageTypeInfo) {
        return pageTypeInfo.artboardIndex >= 0 &&
            pageTypeInfo.artboardType === "spread" &&
            pageTypeInfo.pageType === "spread";
    }

    /**
     * ドキュメント内の見開きオブジェクトを集める（見つからなければ警告）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]|null} {item, abIndex} の配列（見つからなければ null）
     */
    function collectSpreadsInDocument(doc) {
        var spreads = [];
        var candidates = collectTargetItems(doc);
        for (var i = 0; i < candidates.length; i++) {
            var pageTypeInfo = getPageTypeInfo(doc, candidates[i]);
            if (isSpreadOnSpreadArtboard(pageTypeInfo)) {
                spreads.push({ item: candidates[i], abIndex: pageTypeInfo.artboardIndex });
            }
        }

        if (spreads.length === 0) {
            alert(getLabel("alert.noSpreadFoundAll"));
            return null;
        }
        return spreads;
    }

    /**
     * 選択から見開きオブジェクトを集める（見つからなければ理由に合わせて警告）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object[]|null} {item, abIndex} の配列（見つからなければ null）
     */
    function collectSpreadsInSelection(doc) {
        if (doc.selection.length < 1) {
            alert(getLabel("alert.selectSpreadObject"));
            return null;
        }

        var spreads = [];
        var singleCount = 0;
        var otherCount = 0;

        for (var i = 0; i < doc.selection.length; i++) {
            var selectedItem = doc.selection[i];
            if (!isTargetItem(selectedItem)) {
                otherCount++;
                continue;
            }
            var pageTypeInfo = getPageTypeInfo(doc, selectedItem);
            if (pageTypeInfo.artboardIndex < 0) {
                otherCount++;
            } else if (isSpreadOnSpreadArtboard(pageTypeInfo)) {
                spreads.push({ item: selectedItem, abIndex: pageTypeInfo.artboardIndex });
            } else if (pageTypeInfo.pageType === "single") {
                singleCount++;
            } else {
                otherCount++;
            }
        }

        if (spreads.length === 0) {
            if (singleCount > 0 && otherCount === 0) {
                alert(getLabel("alert.selectionIsSingle"));
            } else {
                alert(getLabel("alert.selectValidSpread"));
            }
            return null;
        }
        return spreads;
    }

    /**
     * アートボードごとに最初の1件だけを残し、番号の大きい順に並べる
     * （アートボードの挿入で番号がずれるため後ろから処理する）
     * @param {Object[]} spreads - {item, abIndex} の配列
     * @returns {Object[]} 処理順に並べた {item, abIndex} の配列
     */
    function buildWorkList(spreads) {
        var workList = [];
        var usedArtboardMap = {};
        for (var i = 0; i < spreads.length; i++) {
            if (usedArtboardMap[spreads[i].abIndex]) continue;
            usedArtboardMap[spreads[i].abIndex] = true;
            workList.push(spreads[i]);
        }
        workList.sort(function (a, b) { return b.abIndex - a.abIndex; });
        return workList;
    }

    // =========================================
    // 分割 / Splitting
    // =========================================

    /**
     * 矩形でクリップしたグループにオブジェクトを入れる
     * @param {PageItem} targetItem - クリップするオブジェクト
     * @param {number[]} clipRect - クリップ範囲 [left, top, right, bottom]
     * @param {string} groupName - グループ名
     * @returns {GroupItem} クリッピンググループ
     */
    function createClipGroup(targetItem, clipRect, groupName) {
        var parentContainer = targetItem.parent;

        var clipPath = parentContainer.pathItems.rectangle(
            clipRect[1],
            clipRect[0],
            clipRect[2] - clipRect[0],
            clipRect[1] - clipRect[3]
        );
        clipPath.stroked = false;
        clipPath.filled = false;

        var clipGroup = parentContainer.groupItems.add();
        clipGroup.name = groupName;

        targetItem.move(clipGroup, ElementPlacement.PLACEATEND);
        clipPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);

        clipPath.clipping = true;
        clipGroup.clipped = true;

        return clipGroup;
    }

    /**
     * 見開きオブジェクトを左右2つにクリップし、アートボードも2つに分ける
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} spreadItem - 見開きオブジェクト
     * @param {number} abIndex - 見開きアートボードの番号
     * @param {boolean} evenOnRight - 偶数ページを右にするか
     * @param {number} pageOffset - 分けたアートボードの間隔の半分
     * @returns {void}
     */
    function splitSpread(doc, spreadItem, abIndex, evenOnRight, pageOffset) {
        /* 元アートボードの位置とサイズを基準にする / Use the source artboard bounds as the base */
        var abRect = doc.artboards[abIndex].artboardRect;
        var left = abRect[0];
        var top = abRect[1];
        var right = abRect[2];
        var bottom = abRect[3];
        var centerX = left + (right - left) / 2;

        /* 複製（B） / Duplicate for right half */
        var rightHalfSource = spreadItem.duplicate();

        /* アートボード基準の左右ページ矩形 / Left and right page rects based on the artboard */
        var leftClipRect = [left, top, centerX, bottom];
        var rightClipRect = [centerX, top, right, bottom];

        /* 元オブジェクト(A) → 左半分、複製(B) → 右半分 / Original object to left half, duplicate to right half */
        var leftHalfGroup = createClipGroup(spreadItem, leftClipRect, getLabel("generatedName.leftHalfGroup"));
        var rightHalfGroup = createClipGroup(rightHalfSource, rightClipRect, getLabel("generatedName.rightHalfGroup"));

        var leftPageRect;
        var rightPageRect;
        var leftHalfTargetLeft;
        var rightHalfTargetLeft;

        if (evenOnRight) {
            rightPageRect = [centerX, top, right, bottom];
            leftPageRect = [left - pageOffset, top, centerX - pageOffset, bottom];
            /* 左半分（A）は右ページへ、右半分（B）は左ページへ移す / Move left half to right page and right half to left page */
            leftHalfTargetLeft = rightPageRect[0];
            rightHalfTargetLeft = leftPageRect[0];
        } else {
            /* 偶数ページが左（デフォルト） / Even pages on the left (default) */
            leftPageRect = [left, top, centerX, bottom];
            rightPageRect = [centerX + pageOffset, top, right + pageOffset, bottom];
            /* 左半分（A）は左ページのまま、右半分（B）は右ページへ移す / Keep left half on left page and move right half to right page */
            leftHalfTargetLeft = leftPageRect[0];
            rightHalfTargetLeft = rightPageRect[0];
        }

        /* 見た目の左→右に合わせて、元のアートボードを左ページ、追加アートボードを右ページにする / Original artboard becomes the left page, the inserted one the right page */
        doc.artboards[abIndex].artboardRect = leftPageRect;
        doc.artboards.setActiveArtboardIndex(abIndex);
        doc.artboards.insert(rightPageRect, abIndex + 1);

        /* クリップ範囲から配置先までの移動量 / Translation from the clip rects to the destination pages */
        leftHalfGroup.translate(leftHalfTargetLeft - leftClipRect[0], 0);
        rightHalfGroup.translate(rightHalfTargetLeft - rightClipRect[0], 0);

        leftHalfGroup.note = SPLIT_GROUP_NOTE_PREFIX + ":role=A";
        rightHalfGroup.note = SPLIT_GROUP_NOTE_PREFIX + ":role=B";

        leftHalfGroup.selected = true;
        rightHalfGroup.selected = true;
    }

    // =========================================
    // アートボード再配置 / Artboard rearrangement
    // =========================================

    /**
     * レイヤーとオブジェクトのロック・非表示をすべて解除する（入れ子もたどる）
     * @param {Object} container - Document / Layer / GroupItem
     * @returns {void}
     */
    function unlockAndUnhideAll(container) {
        if (!container) return;

        if (container.typename === "Document" || container.typename === "Layer") {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                var layer = layers[i];
                if (layer.locked) layer.locked = false;
                if (!layer.visible) layer.visible = true;
                unlockAndUnhideAll(layer);
            }
        }

        if (container.pageItems) {
            var pageItems = container.pageItems;
            for (var j = 0; j < pageItems.length; j++) {
                var pageItem = pageItems[j];
                if (pageItem.locked) pageItem.locked = false;
                if (pageItem.hidden) pageItem.hidden = false;
                if (pageItem.typename === "GroupItem") {
                    unlockAndUnhideAll(pageItem);
                }
            }
        }
    }

    /**
     * アートボードごとに乗っているオブジェクトを集める
     * （2つのアートボードにまたがるものは最初のアートボードだけに入れる）
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[][]} アートボード番号ごとのオブジェクト
     */
    function collectItemsByArtboard(doc) {
        var artboards = doc.artboards;
        var artboardItemMap = [];
        var markedItems = [];
        var originalNotes = [];

        for (var i = 0; i < artboards.length; i++) {
            artboards.setActiveArtboardIndex(i);
            doc.selection = null;
            doc.selectObjectsOnActiveArtboard();

            var itemsOnArtboard = [];
            for (var j = 0; j < doc.selection.length; j++) {
                var selectedItem = doc.selection[j];
                if (!selectedItem) continue;

                /* 目印付き＝前のアートボードで集めたもの / A marked item was already collected */
                var noteText = String(selectedItem.note || "");
                if (noteText.indexOf(REARRANGE_NOTE_MARKER) >= 0) continue;

                markedItems.push(selectedItem);
                originalNotes.push(noteText);
                selectedItem.note = noteText ? (noteText + "\n" + REARRANGE_NOTE_MARKER) : REARRANGE_NOTE_MARKER;
                itemsOnArtboard.push(selectedItem);
            }

            artboardItemMap[i] = itemsOnArtboard;
            doc.selection = null;
        }

        /* 目印を外して note を戻す / Restore the original notes */
        for (var k = 0; k < markedItems.length; k++) {
            markedItems[k].note = originalNotes[k];
        }
        return artboardItemMap;
    }

    /**
     * アートボードをページ順（見開きの並び）に並べ直し、乗っているオブジェクトも一緒に動かす
     * @param {Document} doc - 対象ドキュメント
     * @param {number} spacingX - 横方向の間隔
     * @param {number} spacingY - 縦方向の間隔
     * @param {boolean} evenOnRight - 偶数ページを右にするか
     * @returns {void}
     */
    function rearrangeArtboardsByPageOrder(doc, spacingX, spacingY, evenOnRight) {
        if (!doc || !doc.artboards || doc.artboards.length === 0) return;

        unlockAndUnhideAll(doc);

        var artboards = doc.artboards;
        var firstRect = artboards[0].artboardRect;
        var baseX = firstRect[0];
        var baseY = firstRect[1];
        var firstPageRight = !evenOnRight;
        var artboardItemMap = collectItemsByArtboard(doc);

        for (var i = 0; i < artboards.length; i++) {
            var artboard = artboards[i];
            var itemsOnArtboard = artboardItemMap[i] || [];
            var oldRect = artboard.artboardRect;
            var pageWidth = oldRect[2] - oldRect[0];
            var pageHeight = oldRect[1] - oldRect[3];
            var pageNumber = i + 1;
            var rowIndex;
            var newLeft;

            if (pageNumber === 1) {
                rowIndex = 0;
                newLeft = baseX;
            } else {
                var isThisPageRight;
                if (firstPageRight) {
                    isThisPageRight = (pageNumber % 2 !== 0);
                    newLeft = isThisPageRight ? baseX : (baseX - pageWidth - spacingX);
                } else {
                    isThisPageRight = (pageNumber % 2 === 0);
                    newLeft = isThisPageRight ? (baseX + pageWidth + spacingX) : baseX;
                }
                rowIndex = Math.floor(pageNumber / 2);
            }

            var newTop = baseY - rowIndex * (pageHeight + spacingY);
            var deltaX = newLeft - oldRect[0];
            var deltaY = newTop - oldRect[1];

            artboard.artboardRect = [newLeft, newTop, newLeft + pageWidth, newTop - pageHeight];

            if (deltaX !== 0 || deltaY !== 0) {
                for (var k = 0; k < itemsOnArtboard.length; k++) {
                    itemsOnArtboard[k].translate(deltaX, deltaY, true, true, true, true);
                }
            }
        }

        app.executeMenuCommand("fitall");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 分けたアートボードの間隔の半分を、新規ドキュメントの「アートボードの間隔」から求める
     * @returns {number} 間隔の半分（読めなければ 5）
     */
    function getPageOffset() {
        /* 環境設定が読めなければ既定値 / fall back when the preference cannot be read */
        try {
            return app.preferences.getRealPreference("artnewdialog/artboardSpacing") / 2;
        } catch (e) {
            return 5;
        }
    }

    /**
     * 見開きオブジェクトを片ページに分割する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            return;
        }

        var doc = app.activeDocument;

        var splitOptions = showSplitDialog();
        if (!splitOptions) return;

        if (splitOptions.rearrangeArtboards && (isNaN(splitOptions.spacingHorizontal) || isNaN(splitOptions.spacingVertical))) {
            alert(getLabel("alert.invalidSpacing"));
            return;
        }

        var pageOffset = getPageOffset();

        /* 対象オブジェクトの収集 / Collect target objects */
        var spreads = splitOptions.processAll ? collectSpreadsInDocument(doc) : collectSpreadsInSelection(doc);
        if (!spreads) return;

        var workList = buildWorkList(spreads);

        /* 各オブジェクトを処理 / Process each object */
        doc.selection = null;
        for (var i = 0; i < workList.length; i++) {
            splitSpread(doc, workList[i].item, workList[i].abIndex, splitOptions.evenOnRight, pageOffset);
        }

        /* アートボード名をリネーム（オプション） / Rename artboards (optional) */
        if (splitOptions.renameArtboards) {
            for (var k = 0; k < doc.artboards.length; k++) {
                doc.artboards[k].name = getLabel("generatedName.artboardPrefix") + (k + 1);
            }
        }

        /* アートボード再配置（オプション） / Rearrange artboards (optional) */
        if (splitOptions.rearrangeArtboards) {
            rearrangeArtboardsByPageOrder(doc, splitOptions.spacingHorizontal, splitOptions.spacingVertical, splitOptions.evenOnRight);
        }
    }

    main();

})();
