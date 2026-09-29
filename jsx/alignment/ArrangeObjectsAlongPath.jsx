#target illustrator
#targetengine "ArrangeObjectsAlongPathEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のオブジェクトを、選択範囲内の1本のパスに沿って等間隔に自動配置します。
基準にするパスは「自動（面積最大）／最前面／最背面」から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeObjectsAlongPath.md

### Overview

Distributes several objects evenly along a single path taken from the selection.
The reference path is chosen automatically by largest area, or as the frontmost or backmost object.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeObjectsAlongPath.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArrangeObjectsAlongPath";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeObjectsAlongPath.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeObjectsAlongPath.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 配置後、さらに接線方向へ回転する（簡易）/ Also rotate along the tangent after placing (simple) */
    var ROTATE_ALONG_TANGENT = false;
    /* true: 始点〜終点を含めて等分 / false: 端を避ける（開いたパス）/ true: include both ends, false: keep off the ends (open paths) */
    var USE_ENDPOINTS = true;
    /* 曲線1区間あたりのサンプル数（増やすほど精度が上がり重くなる）/ Samples per Bezier segment (more is precise but slower) */
    var SAMPLES_PER_SEGMENT = 30;
    /* ランダム間隔の強さの初期値（ステップに対する比率、0.1〜1.0）/ Initial random spacing strength (ratio of the step, 0.1-1.0) */
    var DEFAULT_SPACING_JITTER_RATIO = 0.4;
    /* 複製数の範囲 / Duplicate count range */
    var DUPLICATE_COUNT_MIN = 2;
    var DUPLICATE_COUNT_MAX = 20;

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* プレビュー用の一時レイヤー名 / Name of the temporary preview layer */
    var PREVIEW_LAYER_NAME = "__PREVIEW_ArrangeAlongPath";

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* 初めて開くときのダイアログの位置 / Dialog position on first open */
    var DIALOG_OFFSET_X = 300;
    var DIALOG_OFFSET_Y = 0;

    var COLUMN_SPACING = 15;                        /* 左右カラムの間隔 / Gap between the columns */
    var OUTER_PANEL_MARGINS = [15, 20, 15, 15];     /* 「対象パス」パネルの余白 / Margins of the Target Path panel */
    var PANEL_MARGINS = [15, 20, 15, 10];           /* その他のパネルの余白 / Margins of the other panels */
    var SLIDER_WIDTH = 180;                         /* スライダーの幅 / Slider width */

    /**
     * 表示時にダイアログを指定量ずらす
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
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

    // ボタン行（再利用パーツ） / Button row (reusable)

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

    /**
     * タイトル付きのパネルを追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} titlePath - パネルタイトルの LABELS パス
     * @param {string} orientation - "row" または "column"
     * @param {string|string[]} childAlignment - alignChildren に入れる値
     * @param {number[]} [panelMargins] - 余白（省略時は PANEL_MARGINS）
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentContainer, titlePath, orientation, childAlignment, panelMargins) {
        var optionPanel = parentContainer.add("panel", undefined, getLabel(titlePath));
        optionPanel.orientation = orientation;
        optionPanel.alignChildren = childAlignment;
        optionPanel.margins = panelMargins || PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * 上揃えの縦カラムを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加したカラム
     */
    function addColumn(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "fill";
        columnGroup.alignment = "top";
        return columnGroup;
    }

    /**
     * 余りの幅を吸う伸縮スペーサーを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {void}
     */
    function addStretchSpacer(parentGroup) {
        var stretchSpacer = parentGroup.add("statictext", undefined, "");
        stretchSpacer.alignment = "fill";
        stretchSpacer.minimumSize.width = 10;
        stretchSpacer.maximumSize.width = 10000;
    }

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

    var LABELS = {
        dialog: {
            title: { ja: "パスに沿って配置", en: "Arrange Objects Along Path" }
        },
        panel: {
            targetPath: { ja: "対象パス", en: "Target Path" },
            basePathRule: { ja: "基準", en: "Base Path" },
            basePathHandling: { ja: "パス処理", en: "Base Path Handling" },
            placeObjects: { ja: "配置するオブジェクト", en: "Objects to Arrange" },
            duplicate: { ja: "複製", en: "Duplicate" },
            rotation: { ja: "回転", en: "Rotation" },
            order: { ja: "順番", en: "Order" },
            spacing: { ja: "間隔", en: "Spacing" }
        },
        fieldLabel: {
            duplicateCount: { ja: "複製数", en: "Count" }
        },
        checkbox: {
            rotationFlip180: { ja: "反転", en: "Flip" },
            groupPlaced: { ja: "グループ化", en: "Group placed objects" },
            allRandom: { ja: "一括ランダム", en: "Random" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        radio: {
            autoLargest: { ja: "自動（面積最大）", en: "Auto (Largest)" },
            frontmost: { ja: "最前面", en: "Frontmost" },
            backmost: { ja: "最背面", en: "Backmost" },
            basePathModeNone: { ja: "何もしない", en: "Do nothing" },
            basePathModeHide: { ja: "「塗り／線」なし", en: "No fill / no stroke" },
            basePathModeDelete: { ja: "削除", en: "Delete" },
            rotationNone: { ja: "正立", en: "Upright" },
            rotationPerpendicular: { ja: "それぞれ垂直", en: "Perpendicular" },
            rotationPathPerpendicular: { ja: "パスに沿う（接線）", en: "Follow Path (Tangent)" },
            rotationRandom: { ja: "ランダム", en: "Random" },
            rotationAngle: { ja: "角度指定", en: "Angle" },
            orderCurrent: { ja: "正順", en: "Current" },
            orderReverse: { ja: "逆順", en: "Reverse" },
            orderRandom: { ja: "ランダム", en: "Random" },
            spacingEven: { ja: "均等（現状）", en: "Even (Current)" },
            spacingRandom: { ja: "ランダム", en: "Random" }
        },
        tooltip: {
            basePathRule: { ja: "選択の中から、どれを基準のパス（B）とみなすかを決めます。", en: "Which of the selected paths is treated as the base path (B)." },
            autoLargest: {
                ja: "面積が最も大きいパスを基準にします。開いたパスは外接矩形の面積で比べます。",
                en: "Uses the path with the largest area. Open paths are compared by their bounding box."
            },
            basePathHandling: { ja: "基準にしたパス（B）を、配置後にどう扱うかを決めます。", en: "What to do with the base path (B) once the objects are placed." },
            duplicateEnabled: {
                ja: "オンにすると、それぞれのオブジェクトを「複製数」の数まで増やしてから配置します。選択が2つのときは最初からオンです。",
                en: "When on, each object is multiplied up to the Count before arranging. Starts on when exactly two objects are selected."
            },
            duplicateCount: {
                ja: "選択したオブジェクトを複製して数を増やしてから配置します。1 なら複製しません。",
                en: "Duplicates the selection to this many copies before arranging. 1 means no duplication."
            },
            rotation: { ja: "配置したオブジェクトの向きの決め方です。", en: "How each placed object is rotated." },
            rotationNone: {
                ja: "オブジェクトの向きを変えません（［反転］がオンなら 180° 回転します）。",
                en: "Keeps each object's current rotation (turns it 180° when Flip is on)."
            },
            rotationPerpendicular: {
                ja: "基準パスの中心から放射状に、外向きに立つよう回転します。",
                en: "Rotates each object to stand outward, radiating from the center of the base path."
            },
            rotationPathPerpendicular: {
                ja: "パスの進む向き（接線）に合わせて回転します。",
                en: "Rotates each object to follow the direction of the path (its tangent)."
            },
            rotationRandom: {
                ja: "オブジェクトごとに −180°〜180° のランダムな角度だけ回転します。",
                en: "Rotates each object by a random angle between -180° and 180°."
            },
            rotationAngle: { ja: "「角度指定」を選んだときに適用する角度です。", en: "The angle applied when Angle is selected." },
            rotationFlip180: { ja: "回転に 180° を加えて、向きを反対にします。", en: "Adds 180° to the rotation to turn each object around." },
            order: { ja: "オブジェクトをパスに沿って並べる順序です。", en: "The order the objects are laid along the path." },
            orderCurrent: {
                ja: "重ね順で前面にあるものから、パスの始点側に並べます。",
                en: "Lays the objects from the start of the path in stacking order, frontmost first."
            },
            orderReverse: {
                ja: "重ね順で背面にあるものから、パスの始点側に並べます。",
                en: "Lays the objects from the start of the path in stacking order, backmost first."
            },
            spacing: { ja: "パス上に配置する間隔の決め方です。", en: "How the objects are spaced along the path." },
            spacingEven: { ja: "パスに沿って等間隔に配置します。", en: "Places the objects at equal intervals along the path." },
            spacingRandom: {
                ja: "等間隔の位置から、下のスライダーの強さでランダムにずらします。",
                en: "Shifts each position randomly away from even spacing, by the strength set with the slider below."
            },
            spacingJitter: {
                ja: "ランダム間隔のばらつきの強さです。右ほど均等な位置から大きくずれます。「ランダム」のときだけ使えます。",
                en: "How far the random spacing strays from even spacing; further right strays more. Available only with Random."
            },
            groupPlaced: { ja: "配置したオブジェクトを1つのグループにまとめます。", en: "Groups the placed objects into a single group." },
            allRandom: { ja: "順番・間隔・回転をまとめてランダムに設定します。", en: "Sets order, spacing and rotation all to random at once." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original state."
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントがありません。", en: "No document is open." },
            needSelection: {
                ja: "A（複数オブジェクト）とB（基準パス）を選択してください。基準パス（B）は「自動（面積最大）/ 最前面 / 最背面」で指定できます。",
                en: "Select A (objects) and B (a path). Choose the base path by Auto (largest), Frontmost, or Backmost."
            },
            noBasePath: {
                ja: "選択範囲に基準となるパス（B）が見つかりません。基準パスにしたいパス（PathItem）を含めて選択してください。",
                en: "No base path (B) was found. Include a PathItem to be used as the base path."
            },
            noItems: { ja: "配置対象（A）が見つかりません。", en: "No placeable objects (A) were found." },
            pathTooShort: { ja: "パスが短すぎます。", en: "The path is too short." },
            pathAnalyzeFailed: { ja: "パスの解析に失敗しました。", en: "Failed to analyze the path." },
            pathLengthZero: { ja: "パス長が0です。", en: "The path length is zero." }
        }
    };

    // =========================================
    // 配列と乱数 / Arrays and random numbers
    // =========================================

    /**
     * 配列の浅いコピーを返す
     * @param {Array} sourceList - 元の配列
     * @returns {Array} コピー
     */
    function copyArray(sourceList) {
        var copiedList = [];
        for (var i = 0; i < sourceList.length; i++) copiedList.push(sourceList[i]);
        return copiedList;
    }

    /**
     * シード付きの乱数関数を作る（xorshift32）
     * @param {number} seed - シード
     * @returns {Function} 0 以上 1 未満の数を返す関数
     */
    function createSeededRandom(seed) {
        var state = (seed | 0);
        if (state === 0) state = 123456789;
        return function () {
            state ^= (state << 13);
            state ^= (state >>> 17);
            state ^= (state << 5);
            return (state >>> 0) / 4294967296;
        };
    }

    /**
     * 現在時刻から並べ替え用のシードを作る
     * @returns {number} シード
     */
    function createShuffleSeed() {
        return (new Date().getTime() & 0x7fffffff) >>> 0;
    }

    /**
     * 配列をその場で並べ替える（Fisher-Yates）
     * @param {Array} targetList - 並べ替える配列
     * @param {Function} [randomFn] - 乱数関数（省略時は Math.random）
     * @returns {void}
     */
    function shuffleInPlace(targetList, randomFn) {
        for (var i = targetList.length - 1; i > 0; i--) {
            var randomValue = randomFn ? randomFn() : Math.random();
            var j = Math.floor(randomValue * (i + 1));
            var swappedItem = targetList[i];
            targetList[i] = targetList[j];
            targetList[j] = swappedItem;
        }
    }

    /**
     * 複製数を範囲内の整数にそろえる
     * @param {number|string} value - 入力値
     * @returns {number} DUPLICATE_COUNT_MIN〜DUPLICATE_COUNT_MAX の整数
     */
    function clampDuplicateCount(value) {
        var count = Math.round(Number(value));
        if (isNaN(count)) count = DUPLICATE_COUNT_MIN;
        if (count < DUPLICATE_COUNT_MIN) count = DUPLICATE_COUNT_MIN;
        if (count > DUPLICATE_COUNT_MAX) count = DUPLICATE_COUNT_MAX;
        return count;
    }

    /**
     * ランダム間隔の強さを 0.1〜1.0 にそろえる
     * @param {number} value - スライダーの値
     * @returns {number} 強さ
     */
    function clampJitterRatio(value) {
        var ratio = Number(value);
        if (isNaN(ratio)) ratio = DEFAULT_SPACING_JITTER_RATIO;
        if (ratio < 0.1) ratio = 0.1;
        if (ratio > 1.0) ratio = 1.0;
        return ratio;
    }

    // =========================================
    // 選択と重ね順 / Selection and stacking order
    // =========================================

    /**
     * 2点以上あるパスか
     * @param {PageItem} candidate - 調べるオブジェクト
     * @returns {boolean} パスなら true
     */
    function isPathItem(candidate) {
        return candidate && candidate.typename === "PathItem" && candidate.pathPoints && candidate.pathPoints.length >= 2;
    }

    /**
     * パスに沿って配置できる種類のオブジェクトか
     * @param {PageItem} candidate - 調べるオブジェクト
     * @returns {boolean} 配置できるなら true
     */
    function isPlaceableItem(candidate) {
        if (!candidate) return false;
        var typeName = candidate.typename;
        return (
            typeName === "PathItem" ||
            typeName === "CompoundPathItem" ||
            typeName === "GroupItem" ||
            typeName === "TextFrame" ||
            typeName === "PlacedItem" ||
            typeName === "RasterItem" ||
            typeName === "SymbolItem" ||
            typeName === "MeshItem"
        );
    }

    /**
     * 選択から、基準パス以外の配置できるオブジェクトを集める
     * @param {PageItem[]} selectionItems - 選択
     * @param {PathItem} basePath - 基準パス
     * @returns {PageItem[]} 配置するオブジェクト
     */
    function collectPlaceableItems(selectionItems, basePath) {
        var placeableItems = [];
        for (var i = 0; i < selectionItems.length; i++) {
            if (selectionItems[i] === basePath) continue;
            if (isPlaceableItem(selectionItems[i])) placeableItems.push(selectionItems[i]);
        }
        return placeableItems;
    }

    /**
     * 重ね順の位置を返す
     * @param {PageItem} pageItem - 対象
     * @returns {number|null} zOrderPosition。取れなければ null
     */
    function getZOrderPosition(pageItem) {
        var zOrder = null;
        try { zOrder = pageItem.zOrderPosition; } catch (e) { /* 取れないオブジェクトがある / not available for some items */ }
        if (zOrder === null || zOrder === undefined) return null;
        return zOrder;
    }

    /**
     * 重ね順（最前面→最背面）に並べた配列を返す。重ね順が取れないものは後ろに元の順で並べる
     * @param {PageItem[]} sourceItems - 並べるオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortByStackingOrder(sourceItems) {
        var decoratedItems = [];
        for (var i = 0; i < sourceItems.length; i++) {
            var zOrder = getZOrderPosition(sourceItems[i]);
            decoratedItems.push({ item: sourceItems[i], zOrder: zOrder, hasZOrder: (zOrder !== null), index: i });
        }

        decoratedItems.sort(function (a, b) {
            if (a.hasZOrder && b.hasZOrder) {
                if (a.zOrder === b.zOrder) return a.index - b.index;
                return b.zOrder - a.zOrder; /* 大きいほど前面 / larger is frontmost */
            }
            if (a.hasZOrder) return -1;
            if (b.hasZOrder) return 1;
            return a.index - b.index;
        });

        var sortedItems = [];
        for (var j = 0; j < decoratedItems.length; j++) sortedItems.push(decoratedItems[j].item);
        return sortedItems;
    }

    /**
     * 控えておいた順に並べ直す。控えに無いものは後ろに付ける
     * @param {PageItem[]} baseItems - 並べ直すオブジェクト
     * @param {PageItem[]|null} storedOrder - 控えておいた順
     * @returns {PageItem[]|null} 並べ直した配列。1つも一致しなければ null（別の選択とみなす）
     */
    function reorderByStoredOrder(baseItems, storedOrder) {
        if (!storedOrder || !storedOrder.length) return null;

        var usedFlags = [];
        for (var i = 0; i < baseItems.length; i++) usedFlags[i] = false;

        var reorderedItems = [];
        var matchedCount = 0;

        for (var s = 0; s < storedOrder.length; s++) {
            for (var j = 0; j < baseItems.length; j++) {
                if (!usedFlags[j] && baseItems[j] === storedOrder[s]) {
                    usedFlags[j] = true;
                    reorderedItems.push(baseItems[j]);
                    matchedCount++;
                    break;
                }
            }
        }

        /* 残りを後ろに付ける / Append the leftovers */
        for (var k = 0; k < baseItems.length; k++) {
            if (!usedFlags[k]) reorderedItems.push(baseItems[k]);
        }

        if (matchedCount === 0) return null;
        return reorderedItems;
    }

    /**
     * 外接矩形の面積を返す（線幅を含む visibleBounds を優先）
     * @param {PageItem} pageItem - 対象
     * @returns {number} 面積
     */
    function getBoundsArea(pageItem) {
        var bounds = null;
        try { bounds = pageItem.visibleBounds; } catch (e) { bounds = null; }
        if (!bounds) {
            try { bounds = pageItem.geometricBounds; } catch (e) { bounds = null; }
        }
        if (!bounds || bounds.length < 4) return 0;
        var width = Math.abs(bounds[2] - bounds[0]);
        var height = Math.abs(bounds[1] - bounds[3]);
        return width * height;
    }

    /**
     * 面積が最も大きいパスを返す。開いたパスは外接矩形の面積で比べる
     * @param {PageItem[]} selectionItems - 選択
     * @returns {PathItem|null} パス。無ければ null
     */
    function getLargestPathItem(selectionItems) {
        var largestPath = null;
        var largestArea = null;

        for (var i = 0; i < selectionItems.length; i++) {
            var candidate = selectionItems[i];
            if (!isPathItem(candidate)) continue;

            var area = null;
            if (!candidate.closed) {
                area = getBoundsArea(candidate);
            } else {
                /* 閉じたパスは area（向きで負になる）を使う / Closed paths use area, which can be negative */
                try {
                    area = Math.abs(candidate.area);
                } catch (e) {
                    area = null;
                }
                if (area === null || area === undefined || isNaN(area) || area <= 0) {
                    area = getBoundsArea(candidate);
                }
            }

            if (largestArea === null || area > largestArea) {
                largestArea = area;
                largestPath = candidate;
            }
        }
        return largestPath;
    }

    /**
     * 重ね順で最前面または最背面のパスを返す
     * 重ね順が取れないときは、最前面なら選択の最後、最背面なら選択の最初のパスを返す
     * @param {PageItem[]} selectionItems - 選択
     * @param {boolean} wantFrontmost - 最前面なら true、最背面なら false
     * @returns {PathItem|null} パス。無ければ null
     */
    function findPathItemByStacking(selectionItems, wantFrontmost) {
        var foundPath = null;
        var foundZOrder = null;

        for (var i = 0; i < selectionItems.length; i++) {
            var candidate = selectionItems[i];
            if (!isPathItem(candidate)) continue;

            var zOrder = getZOrderPosition(candidate);
            if (zOrder === null) continue;

            if (foundZOrder === null || (wantFrontmost ? zOrder > foundZOrder : zOrder < foundZOrder)) {
                foundZOrder = zOrder;
                foundPath = candidate;
            }
        }
        if (foundPath) return foundPath;

        if (wantFrontmost) {
            for (var j = selectionItems.length - 1; j >= 0; j--) {
                if (isPathItem(selectionItems[j])) return selectionItems[j];
            }
        } else {
            for (var k = 0; k < selectionItems.length; k++) {
                if (isPathItem(selectionItems[k])) return selectionItems[k];
            }
        }
        return null;
    }

    // =========================================
    // パスの形状 / Path geometry
    // =========================================

    /**
     * 2点間の距離を返す
     * @param {{x: number, y: number}} pointA - 点A
     * @param {{x: number, y: number}} pointB - 点B
     * @returns {number} 距離
     */
    function getDistance(pointA, pointB) {
        var dx = pointB.x - pointA.x;
        var dy = pointB.y - pointA.y;
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 3次ベジェ曲線上の点を返す
     * @param {{x: number, y: number}} startPoint - 始点
     * @param {{x: number, y: number}} control1 - 制御点1
     * @param {{x: number, y: number}} control2 - 制御点2
     * @param {{x: number, y: number}} endPoint - 終点
     * @param {number} t - 媒介変数（0〜1）
     * @returns {{x: number, y: number}} 曲線上の点
     */
    function cubicBezier(startPoint, control1, control2, endPoint, t) {
        var mt = 1 - t;
        var mt2 = mt * mt;
        var t2 = t * t;

        var x =
            startPoint.x * mt2 * mt +
            3 * control1.x * mt2 * t +
            3 * control2.x * mt * t2 +
            endPoint.x * t2 * t;
        var y =
            startPoint.y * mt2 * mt +
            3 * control1.y * mt2 * t +
            3 * control2.y * mt * t2 +
            endPoint.y * t2 * t;

        return { x: x, y: y };
    }

    /**
     * ベジェパスを折れ線に分割する
     * @param {PathItem} pathItem - パス
     * @param {number} samplesPerSegment - 1区間あたりのサンプル数
     * @returns {Array<{x: number, y: number}>} 折れ線の点
     */
    function buildPolylineFromBezierPath(pathItem, samplesPerSegment) {
        var pathPoints = pathItem.pathPoints;
        var isClosed = pathItem.closed;

        var polyline = [];
        var segmentCount = isClosed ? pathPoints.length : (pathPoints.length - 1);

        for (var i = 0; i < segmentCount; i++) {
            var startAnchor = pathPoints[i];
            var endAnchor = pathPoints[(i + 1) % pathPoints.length];

            var startPoint = { x: startAnchor.anchor[0], y: startAnchor.anchor[1] };
            var control1 = { x: startAnchor.rightDirection[0], y: startAnchor.rightDirection[1] };
            var control2 = { x: endAnchor.leftDirection[0], y: endAnchor.leftDirection[1] };
            var endPoint = { x: endAnchor.anchor[0], y: endAnchor.anchor[1] };

            for (var s = 0; s <= samplesPerSegment; s++) {
                /* 区間のつなぎ目は重複させない / Do not repeat the joint between segments */
                if (i > 0 && s === 0) continue;
                polyline.push(cubicBezier(startPoint, control1, control2, endPoint, s / samplesPerSegment));
            }
        }
        return polyline;
    }

    /**
     * 折れ線の各点までの累積長を返す
     * @param {Array<{x: number, y: number}>} polyline - 折れ線
     * @returns {number[]} 累積長（先頭は 0）
     */
    function getCumulativeLengths(polyline) {
        var cumulativeLengths = [0];
        for (var i = 1; i < polyline.length; i++) {
            cumulativeLengths[i] = cumulativeLengths[i - 1] + getDistance(polyline[i - 1], polyline[i]);
        }
        return cumulativeLengths;
    }

    /**
     * 始点からの距離にある折れ線上の点を返す
     * @param {Array<{x: number, y: number}>} polyline - 折れ線
     * @param {number[]} cumulativeLengths - 累積長
     * @param {number} distance - 始点からの距離
     * @returns {{x: number, y: number}} 点
     */
    function pointAtDistance(polyline, cumulativeLengths, distance) {
        if (distance <= 0) return { x: polyline[0].x, y: polyline[0].y };
        var totalLength = cumulativeLengths[cumulativeLengths.length - 1];
        if (distance >= totalLength) return { x: polyline[polyline.length - 1].x, y: polyline[polyline.length - 1].y };

        var index = 1;
        while (index < cumulativeLengths.length && cumulativeLengths[index] < distance) index++;

        var startLength = cumulativeLengths[index - 1];
        var endLength = cumulativeLengths[index];
        var ratio = (distance - startLength) / (endLength - startLength);

        var startPoint = polyline[index - 1];
        var endPoint = polyline[index];

        return {
            x: startPoint.x + (endPoint.x - startPoint.x) * ratio,
            y: startPoint.y + (endPoint.y - startPoint.y) * ratio
        };
    }

    /**
     * 始点からの距離にある接線の角度を返す（端でも安定させる）
     * @param {Array<{x: number, y: number}>} polyline - 折れ線
     * @param {number[]} cumulativeLengths - 累積長
     * @param {number} distance - 始点からの距離
     * @param {boolean} isClosed - 閉じたパスか
     * @returns {number} 角度（度）
     */
    function tangentDegAtDistance(polyline, cumulativeLengths, distance, isClosed) {
        var totalLength = cumulativeLengths[cumulativeLengths.length - 1];
        if (totalLength <= 0) return 0;

        if (distance < 0) distance = 0;
        if (distance > totalLength) distance = totalLength;

        /* cumulativeLengths[index] >= distance となる区間を探す（distance は全長以下なので index は末尾を超えない）
           Find the segment where the length reaches the distance; index never passes the last point */
        var index = 1;
        while (index < cumulativeLengths.length && cumulativeLengths[index] < distance) index++;

        var startPoint = polyline[index - 1];
        var endPoint = polyline[index];

        /* 長さ 0 の区間なら隣を使う（尖った点・重なった点で有効）/ Use a neighbor for a zero-length segment */
        if ((endPoint.x === startPoint.x) && (endPoint.y === startPoint.y)) {
            if (index + 1 < polyline.length) {
                endPoint = polyline[index + 1];
            } else if (isClosed && polyline.length > 2) {
                endPoint = polyline[1];
            }
        }

        return Math.atan2(endPoint.y - startPoint.y, endPoint.x - startPoint.x) * 180 / Math.PI;
    }

    /**
     * 均等配置の位置（始点からの距離）を返す
     * @param {number} totalLength - パスの長さ
     * @param {number} itemCount - 配置する数
     * @param {boolean} isClosed - 閉じたパスか
     * @param {boolean} useEndpoints - 開いたパスで始点・終点を含めるか
     * @returns {number[]} 距離の配列
     */
    function computeEvenDistances(totalLength, itemCount, isClosed, useEndpoints) {
        var distances = [];
        if (itemCount === 1) {
            distances.push(totalLength / 2);
        } else if (isClosed) {
            for (var i = 0; i < itemCount; i++) distances.push((totalLength * i) / itemCount);
        } else if (useEndpoints) {
            for (var j = 0; j < itemCount; j++) distances.push((totalLength * j) / (itemCount - 1));
        } else {
            for (var k = 0; k < itemCount; k++) distances.push(totalLength * (k + 0.5) / itemCount);
        }
        return distances;
    }

    /**
     * 均等配置の位置にばらつきを加えた位置を返す（始点から昇順）
     * @param {number} totalLength - パスの長さ
     * @param {number} itemCount - 配置する数
     * @param {boolean} isClosed - 閉じたパスか
     * @param {boolean} useEndpoints - 開いたパスで始点・終点を含めるか
     * @param {number} jitterRatio - ばらつきの強さ（間隔に対する比率）
     * @returns {number[]} 距離の配列
     */
    function computeJitteredDistances(totalLength, itemCount, isClosed, useEndpoints, jitterRatio) {
        /* 取りうる範囲 / Allowed range */
        var minDistance = 0;
        var maxDistance = totalLength;
        if (!isClosed && !useEndpoints) {
            var padding = totalLength * 0.02;
            if (padding > 0) {
                minDistance = padding;
                maxDistance = Math.max(padding, totalLength - padding);
            }
        }

        /* 均等配置から少しだけずらす / Start from even spacing and add a little jitter */
        var evenDistances = computeEvenDistances(totalLength, itemCount, isClosed, useEndpoints);
        var step = (itemCount <= 1) ? totalLength : (isClosed ? (totalLength / itemCount) : (useEndpoints ? (totalLength / (itemCount - 1)) : (totalLength / itemCount)));
        var jitter = step * jitterRatio;

        var distances = [];
        for (var i = 0; i < itemCount; i++) {
            var distance = evenDistances[i];
            /* 開いたパスで端を含めるときは両端を動かさない / Keep both ends fixed on an open path with endpoints */
            var isFixedEnd = !isClosed && useEndpoints && itemCount > 1 && (i === 0 || i === itemCount - 1);
            if (!isFixedEnd) distance += (Math.random() * 2 - 1) * jitter;

            if (distance < minDistance) distance = minDistance;
            if (distance > maxDistance) distance = maxDistance;
            distances.push(distance);
        }

        /* 間隔はばらつくが、並びはパスに沿って昇順 / Random gaps, but monotonic along the path */
        distances.sort(function (a, b) { return a - b; });
        return distances;
    }

    // =========================================
    // オブジェクトの操作 / Object operations
    // =========================================

    /**
     * 外接矩形の中心を返す
     * @param {PageItem} pageItem - 対象
     * @returns {{x: number, y: number}} 中心
     */
    function getItemCenter(pageItem) {
        var bounds = pageItem.geometricBounds; /* [left, top, right, bottom] */
        return { x: (bounds[0] + bounds[2]) / 2, y: (bounds[1] + bounds[3]) / 2 };
    }

    /**
     * 現在の回転角を返す
     * @param {PageItem} pageItem - 対象
     * @returns {number} 角度（度）。matrix を持たないオブジェクトは 0
     */
    function getItemRotationDeg(pageItem) {
        try {
            var itemMatrix = pageItem.matrix;
            return Math.atan2(itemMatrix.mValueB, itemMatrix.mValueA) * 180 / Math.PI;
        } catch (e) {
            /* matrix を持つのは配置画像・ラスター・テキストなどだけ / only some item types have a matrix */
            return 0;
        }
    }

    /**
     * 中心を基準に回転する（失敗しても続行）
     * @param {PageItem} pageItem - 対象
     * @param {number} angle - 回転角（度）
     * @returns {void}
     */
    function rotateItemSafely(pageItem, angle) {
        try {
            pageItem.rotate(angle, true, true, true, true, Transformation.CENTER);
        } catch (e) { }
    }

    /**
     * 指定の角度になるよう回転する
     * @param {PageItem} pageItem - 対象
     * @param {number} desiredDeg - 目標の角度（度）
     * @returns {void}
     */
    function rotateItemToDeg(pageItem, desiredDeg) {
        var delta = desiredDeg - getItemRotationDeg(pageItem);
        while (delta > 180) delta -= 360;
        while (delta < -180) delta += 360;
        rotateItemSafely(pageItem, delta);
    }

    /**
     * 塗りと線をなしにする
     * @param {PathItem} pathItem - 対象
     * @returns {void}
     */
    function clearFillAndStroke(pathItem) {
        try {
            pathItem.filled = false;
            pathItem.stroked = false;
        } catch (e) { }
    }

    /**
     * 回転の設定に従って、配置したオブジェクトを回転する
     * @param {PageItem} pageItem - 配置したオブジェクト
     * @param {{mode: string, angle: number, flip180: boolean}} rotationSettings - 回転の設定
     * @param {{x: number, y: number}} baseCenter - 基準パスの中心
     * @param {{polyline: Array, cumulativeLengths: number[], isClosed: boolean}} pathGeometry - パスの形状
     * @param {number} distance - 始点からの距離
     * @returns {void}
     */
    function rotatePlacedItem(pageItem, rotationSettings, baseCenter, pathGeometry, distance) {
        var flipAngle = rotationSettings.flip180 ? 180 : 0;
        var mode = rotationSettings.mode;

        if (mode === "angle") {
            rotateItemSafely(pageItem, (rotationSettings.angle || 0) + flipAngle);
        } else if (mode === "random") {
            rotateItemSafely(pageItem, (Math.random() * 360) - 180 + flipAngle);
        } else if (mode === "perp") {
            /* 基準パスの中心から外向きに立てる / Stand outward from the center of the base path */
            var itemCenter = getItemCenter(pageItem);
            var outwardDeg = Math.atan2(itemCenter.y - baseCenter.y, itemCenter.x - baseCenter.x) * 180 / Math.PI + 270;
            rotateItemToDeg(pageItem, outwardDeg + flipAngle);
        } else if (mode === "path_perp") {
            /* パスの接線方向に合わせる / Follow the path tangent */
            var tangentDeg = tangentDegAtDistance(pathGeometry.polyline, pathGeometry.cumulativeLengths, distance, pathGeometry.isClosed);
            rotateItemToDeg(pageItem, tangentDeg + flipAngle);
        } else if (flipAngle !== 0) {
            /* 「正立」でも反転は効かせる / Flip still applies in Upright mode */
            rotateItemSafely(pageItem, flipAngle);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを出し、OK なら選択したオブジェクトをパスに沿って配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            showDedupedAlert("alert.noDocument");
            return;
        }

        var doc = app.activeDocument;
        var initialSelection = doc.selection;

        /* 選択が0または1ならダイアログを出さずに終了 / Exit without the dialog when fewer than two objects are selected */
        if (!initialSelection || initialSelection.length < 2) {
            showDedupedAlert("alert.needSelection");
            return;
        }

        /* ランダム間隔の強さ（スライダーで変わる）/ Random spacing strength, changed by the slider */
        var spacingJitterRatio = DEFAULT_SPACING_JITTER_RATIO;

        /* プレビューの状態 / Preview state */
        var previewLayer = null;
        var previewSelectionSnapshot = null;   /* 元を隠して選択が外れても作り直せるように控える / kept for rebuilds while originals are hidden */
        var previewHiddenEntries = [];         /* { item, wasHidden } */

        /* ランダムの結果を控え、OK でプレビューと同じ結果にする / Keep random results so OK matches the latest preview */
        var lastRandomOrder = null;            /* PageItem[] */
        var lastRandomSpacing = null;          /* { totalLength: number, distances: number[] } */
        var lastShuffleSeed = null;            /* { seed: number, itemCount: number, duplicateCount: number } 複製込みのランダム順 */

        /* 警告の多重表示を防ぐ / Prevent stacked or repeated alerts */
        var alertLocked = false;
        var alertLastTimeByKey = {};
        var alertLastSignatureByKey = {};

        /* ダイアログのコントロール（buildDialog で作る）/ Dialog controls, created in buildDialog() */
        var rbAutoLargest, rbFrontmost, rbBackmost;
        var rbBasePathKeep, rbBasePathHide, rbBasePathDelete;
        var cbDuplicateEnabled, etDuplicateCount, sldDuplicateCount;
        var rbRotNone, rbRotPerp, rbRotPathPerp, rbRotRandom, rbRotAngle, etRotAngle, cbRotFlip180;
        var rbOrderCurrent, rbOrderReverse, rbOrderRandom;
        var rbSpacingEven, rbSpacingRandom, sldSpacingJitter;
        var cbGroupPlaced, cbAllRandom, cbPreview;

        var arrangeDialog = buildDialog();
        bindDialogEvents();
        updateDuplicateUI();
        updateRotationUI();
        updateSpacingUI();

        prepareDialogWindow(arrangeDialog, SCRIPT_NAME);
        var dialogResult = arrangeDialog.show();

        /* 閉じたらプレビューを片付ける / Clean up the preview on close */
        clearPreview();

        if (dialogResult !== 1) return;
        applyArrangement();

        // -----------------------------------------
        // 警告 / Alerts
        // -----------------------------------------

        /**
         * 警告の重複判定に使う、選択と基準パスの規則の要約を返す
         * @returns {string} 要約
         */
        function getSelectionSignature() {
            try {
                var selectionItems = doc && doc.selection ? doc.selection : null;
                if (!selectionItems || selectionItems.length === 0) return "sel:0";
                var countsByType = {};
                for (var i = 0; i < selectionItems.length; i++) {
                    var typeName = (selectionItems[i] && selectionItems[i].typename) ? selectionItems[i].typename : "?";
                    countsByType[typeName] = (countsByType[typeName] || 0) + 1;
                }
                var signatureParts = ["sel:" + selectionItems.length];
                for (var typeKey in countsByType) {
                    if (countsByType.hasOwnProperty(typeKey)) signatureParts.push(typeKey + ":" + countsByType[typeKey]);
                }
                /* 基準パスの規則（自動／最前面／最背面）も含める / Include the base path rule */
                var ruleName = (rbAutoLargest && rbAutoLargest.value) ? "auto" : ((rbFrontmost && rbFrontmost.value) ? "front" : "back");
                signatureParts.push("rule:" + ruleName);
                return signatureParts.join("|");
            } catch (e) {
                return "sel:?";
            }
        }

        /**
         * 警告を出す。同じ選択で同じ警告は 15 秒間出し直さない
         * @param {string} labelPath - 警告文の LABELS パス
         * @returns {void}
         */
        function showDedupedAlert(labelPath) {
            if (alertLocked) return;

            var now = new Date().getTime();
            var lastTime = alertLastTimeByKey[labelPath] || 0;
            var signature = getSelectionSignature();
            var lastSignature = alertLastSignatureByKey[labelPath] || "";

            if (signature === lastSignature && (now - lastTime) < 15000) return;

            alertLastTimeByKey[labelPath] = now;
            alertLastSignatureByKey[labelPath] = signature;

            alertLocked = true;
            alert(getLabel(labelPath));
            alertLocked = false;
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        /**
         * ダイアログを組み立てる
         * @returns {Window} ダイアログ
         */
        function buildDialog() {
            var arrangeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
            arrangeDialog.orientation = "column";
            arrangeDialog.alignChildren = "fill";

            shiftDialogPosition(arrangeDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

            /* 2カラム / Two columns */
            var columnsGroup = arrangeDialog.add("group");
            columnsGroup.orientation = "row";
            columnsGroup.alignChildren = ["fill", "top"];
            columnsGroup.spacing = COLUMN_SPACING;

            var leftColumn = addColumn(columnsGroup);
            var rightColumn = addColumn(columnsGroup);

            addTargetPathPanel(leftColumn);

            var placeObjectsPanel = addOptionPanel(rightColumn, "panel.placeObjects", "column", "fill");
            addDuplicatePanel(placeObjectsPanel);
            addRotationPanel(placeObjectsPanel);
            addOrderPanel(placeObjectsPanel);
            addSpacingPanel(placeObjectsPanel);
            addGroupOptionsRow(placeObjectsPanel);

            /* 下部のボタンエリア（左：プレビュー、右：キャンセル／OK）/ Button row (left: Preview, right: Cancel / OK) */
            var buttonRow = addButtonRow(arrangeDialog);
            cbPreview = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
            cbPreview.helpTip = getLabel("tooltip.preview");
            cbPreview.value = false;

            var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

            /* キャンセルで必ず閉じる / Always close on Cancel */
            btnCancel.onClick = function () {
                arrangeDialog.close(0);
            };
            return arrangeDialog;
        }

        /**
         * 「対象パス」パネル（基準にするパスとその扱い）を作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function addTargetPathPanel(parentColumn) {
            var targetPathPanel = addOptionPanel(parentColumn, "panel.targetPath", "column", "left", OUTER_PANEL_MARGINS);

            /* 基準にするパス / Which path is the base */
            var basePathRulePanel = addOptionPanel(targetPathPanel, "panel.basePathRule", "column", "left");
            basePathRulePanel.helpTip = getLabel("tooltip.basePathRule");
            rbAutoLargest = basePathRulePanel.add("radiobutton", undefined, getLabel("radio.autoLargest"));
            rbAutoLargest.helpTip = getLabel("tooltip.autoLargest");
            rbFrontmost = basePathRulePanel.add("radiobutton", undefined, getLabel("radio.frontmost"));
            rbBackmost = basePathRulePanel.add("radiobutton", undefined, getLabel("radio.backmost"));
            rbAutoLargest.value = true;

            /* 基準パスの扱い / What to do with the base path */
            var basePathHandlingPanel = addOptionPanel(targetPathPanel, "panel.basePathHandling", "column", "left");
            basePathHandlingPanel.helpTip = getLabel("tooltip.basePathHandling");
            rbBasePathKeep = basePathHandlingPanel.add("radiobutton", undefined, getLabel("radio.basePathModeNone"));
            rbBasePathHide = basePathHandlingPanel.add("radiobutton", undefined, getLabel("radio.basePathModeHide"));
            rbBasePathDelete = basePathHandlingPanel.add("radiobutton", undefined, getLabel("radio.basePathModeDelete"));
            rbBasePathHide.value = true;
        }

        /**
         * 「複製」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addDuplicatePanel(parentPanel) {
            var duplicatePanel = addOptionPanel(parentPanel, "panel.duplicate", "column", "fill");
            duplicatePanel.helpTip = getLabel("tooltip.duplicateCount");

            var duplicateCountRow = duplicatePanel.add("group");
            duplicateCountRow.orientation = "row";
            duplicateCountRow.alignChildren = ["left", "center"];
            duplicateCountRow.spacing = 10;

            cbDuplicateEnabled = duplicateCountRow.add("checkbox", undefined, "");
            /* 選択がちょうど2つのときだけ最初からオン / On by default only when exactly two objects are selected */
            cbDuplicateEnabled.value = (initialSelection.length === 2);
            cbDuplicateEnabled.preferredSize.width = 18;
            cbDuplicateEnabled.helpTip = getLabel("tooltip.duplicateEnabled");

            duplicateCountRow.add("statictext", undefined, getLabel("fieldLabel.duplicateCount"));
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var duplicateCountFieldGroup = duplicateCountRow.add("group");
            duplicateCountFieldGroup.orientation = "row";
            duplicateCountFieldGroup.alignChildren = ["left", "center"];
            duplicateCountFieldGroup.spacing = 0;
            duplicateCountFieldGroup.margins = 0;
            /* 増減後は onChange で範囲に収めてスライダーとプレビューをそろえる / onChange syncs the slider and preview */
            var duplicateCountStepper = addStepper(duplicateCountFieldGroup, function () { return etDuplicateCount; }, {
                integer: true, min: DUPLICATE_COUNT_MIN, max: DUPLICATE_COUNT_MAX,
                onStep: function (numberInput) { numberInput.onChange(); }
            });
            etDuplicateCount = duplicateCountFieldGroup.add("edittext", undefined, String(DUPLICATE_COUNT_MIN));
            etDuplicateCount.helpTip = getLabel("tooltip.duplicateCount");
            etDuplicateCount.characters = 4;
            etDuplicateCount.stepperGroup = duplicateCountStepper;
            bindSteppedArrowKeys(etDuplicateCount, duplicateCountStepper);

            var duplicateSliderRow = duplicatePanel.add("group");
            duplicateSliderRow.orientation = "row";
            duplicateSliderRow.alignChildren = ["left", "center"];

            sldDuplicateCount = duplicateSliderRow.add("slider", undefined, DUPLICATE_COUNT_MIN, DUPLICATE_COUNT_MIN, DUPLICATE_COUNT_MAX);
            sldDuplicateCount.preferredSize.width = SLIDER_WIDTH;
            sldDuplicateCount.helpTip = getLabel("tooltip.duplicateCount");
        }

        /**
         * 「回転」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addRotationPanel(parentPanel) {
            var rotationPanel = addOptionPanel(parentPanel, "panel.rotation", "column", "left");
            rotationPanel.helpTip = getLabel("tooltip.rotation");

            /* 回転のラジオは1つのグループにまとめる（排他は setRotationMode で管理）/ One group for the rotation radios; exclusivity is handled in setRotationMode() */
            var rotationRadioGroup = rotationPanel.add("group");
            rotationRadioGroup.orientation = "column";
            rotationRadioGroup.alignChildren = "left";

            rbRotNone = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationNone"));
            rbRotPerp = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationPerpendicular"));
            rbRotPerp.helpTip = getLabel("tooltip.rotationPerpendicular");
            rbRotPathPerp = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationPathPerpendicular"));
            rbRotPathPerp.helpTip = getLabel("tooltip.rotationPathPerpendicular");
            rbRotRandom = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationRandom"));
            rbRotRandom.helpTip = getLabel("tooltip.rotationRandom");
            rbRotNone.helpTip = getLabel("tooltip.rotationNone");
            rbRotNone.value = true;

            /* 角度指定（1行）/ Angle on one row */
            var rotationAngleRow = rotationRadioGroup.add("group");
            rotationAngleRow.orientation = "row";
            rotationAngleRow.alignChildren = ["left", "center"];
            rotationAngleRow.spacing = 6;

            rbRotAngle = rotationAngleRow.add("radiobutton", undefined, getLabel("radio.rotationAngle"));
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var rotationAngleFieldGroup = rotationAngleRow.add("group");
            rotationAngleFieldGroup.orientation = "row";
            rotationAngleFieldGroup.alignChildren = ["left", "center"];
            rotationAngleFieldGroup.spacing = 0;
            rotationAngleFieldGroup.margins = 0;
            /* 増減後は onChange で「角度指定」にしてプレビューを作り直す / onChange selects Angle mode and rebuilds the preview */
            var rotationAngleStepper = addStepper(rotationAngleFieldGroup, function () { return etRotAngle; }, {
                onStep: function (numberInput) { numberInput.onChange(); }
            });
            etRotAngle = rotationAngleFieldGroup.add("edittext", undefined, "0");
            etRotAngle.helpTip = getLabel("tooltip.rotationAngle");
            etRotAngle.characters = 3;
            etRotAngle.stepperGroup = rotationAngleStepper;
            bindSteppedArrowKeys(etRotAngle, rotationAngleStepper);
            rotationAngleRow.add("statictext", undefined, "°");

            cbRotFlip180 = rotationPanel.add("checkbox", undefined, getLabel("checkbox.rotationFlip180"));
            cbRotFlip180.helpTip = getLabel("tooltip.rotationFlip180");
            cbRotFlip180.value = false;
        }

        /**
         * 「順番」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addOrderPanel(parentPanel) {
            var orderPanel = addOptionPanel(parentPanel, "panel.order", "row", ["left", "center"]);
            orderPanel.helpTip = getLabel("tooltip.order");
            orderPanel.spacing = 12;

            rbOrderCurrent = orderPanel.add("radiobutton", undefined, getLabel("radio.orderCurrent"));
            rbOrderCurrent.helpTip = getLabel("tooltip.orderCurrent");
            rbOrderReverse = orderPanel.add("radiobutton", undefined, getLabel("radio.orderReverse"));
            rbOrderReverse.helpTip = getLabel("tooltip.orderReverse");
            rbOrderRandom = orderPanel.add("radiobutton", undefined, getLabel("radio.orderRandom"));
            rbOrderCurrent.value = true;
        }

        /**
         * 「間隔」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addSpacingPanel(parentPanel) {
            var spacingPanel = addOptionPanel(parentPanel, "panel.spacing", "column", "fill");
            spacingPanel.helpTip = getLabel("tooltip.spacing");

            var spacingRadioGroup = spacingPanel.add("group");
            spacingRadioGroup.orientation = "row";
            spacingRadioGroup.alignChildren = ["left", "center"];
            spacingRadioGroup.spacing = 12;

            rbSpacingEven = spacingRadioGroup.add("radiobutton", undefined, getLabel("radio.spacingEven"));
            rbSpacingRandom = spacingRadioGroup.add("radiobutton", undefined, getLabel("radio.spacingRandom"));
            rbSpacingEven.helpTip = getLabel("tooltip.spacingEven");
            rbSpacingRandom.helpTip = getLabel("tooltip.spacingRandom");
            rbSpacingEven.value = true;

            /* ばらつきの強さ（「ランダム」のときだけ有効）/ Jitter strength, enabled only with Random */
            var spacingSliderRow = spacingPanel.add("group");
            spacingSliderRow.orientation = "row";
            spacingSliderRow.alignChildren = ["left", "center"];

            sldSpacingJitter = spacingSliderRow.add("slider", undefined, spacingJitterRatio, 0.1, 1.0);
            sldSpacingJitter.preferredSize.width = SLIDER_WIDTH;
            sldSpacingJitter.helpTip = getLabel("tooltip.spacingJitter");
        }

        /**
         * 「グループ化」「一括ランダム」の行を中央寄せで作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addGroupOptionsRow(parentPanel) {
            var groupOptionsRow = parentPanel.add("group");
            groupOptionsRow.orientation = "row";
            groupOptionsRow.alignment = "fill";
            groupOptionsRow.alignChildren = ["center", "center"];

            addStretchSpacer(groupOptionsRow);

            var groupOptionsInner = groupOptionsRow.add("group");
            groupOptionsInner.orientation = "row";
            groupOptionsInner.alignChildren = ["center", "center"];
            groupOptionsInner.spacing = 12;

            cbGroupPlaced = groupOptionsInner.add("checkbox", undefined, getLabel("checkbox.groupPlaced"));
            cbGroupPlaced.helpTip = getLabel("tooltip.groupPlaced");
            cbGroupPlaced.value = true;

            cbAllRandom = groupOptionsInner.add("checkbox", undefined, getLabel("checkbox.allRandom"));
            cbAllRandom.helpTip = getLabel("tooltip.allRandom");
            cbAllRandom.value = false;

            addStretchSpacer(groupOptionsRow);
        }

        /**
         * ダイアログのイベントを結び付ける
         * @returns {void}
         */
        function bindDialogEvents() {
            /* プレビューを作り直すだけのコントロール / Controls that only rebuild the preview */
            var previewTriggers = [
                rbAutoLargest, rbFrontmost, rbBackmost,
                rbBasePathKeep, rbBasePathHide, rbBasePathDelete,
                cbRotFlip180,
                rbOrderCurrent, rbOrderReverse, rbOrderRandom
            ];
            for (var i = 0; i < previewTriggers.length; i++) {
                previewTriggers[i].onClick = function () { rebuildPreviewIfNeeded(); };
            }

            cbDuplicateEnabled.onClick = function () {
                if (cbDuplicateEnabled.value) {
                    /* オンにしたら複製数を最小値に戻す / Reset the count when turned on */
                    etDuplicateCount.text = String(DUPLICATE_COUNT_MIN);
                    sldDuplicateCount.value = DUPLICATE_COUNT_MIN;
                }
                updateDuplicateUI();
                rebuildPreviewIfNeeded();
            };
            sldDuplicateCount.onChanging = function () {
                /* ドラッグ中は数字だけ更新 / Update only the text while dragging */
                etDuplicateCount.text = String(clampDuplicateCount(sldDuplicateCount.value));
            };
            sldDuplicateCount.onChange = function () { syncDuplicateCountFromText(); };
            etDuplicateCount.onChange = function () { syncDuplicateCountFromText(); };

            rbRotNone.onClick = function () { onRotationModeClicked("none"); };
            rbRotPerp.onClick = function () { onRotationModeClicked("perp"); };
            rbRotPathPerp.onClick = function () { onRotationModeClicked("path_perp"); };
            rbRotAngle.onClick = function () { onRotationModeClicked("angle"); };
            rbRotRandom.onClick = function () { onRotationModeClicked("random"); };
            etRotAngle.onChange = function () {
                /* 角度を編集したら「角度指定」にする / Editing the angle selects Angle mode */
                setRotationMode("angle");
                rebuildPreviewIfNeeded();
            };

            rbSpacingEven.onClick = function () {
                lastRandomSpacing = null;
                updateSpacingUI();
                rebuildPreviewIfNeeded();
            };
            rbSpacingRandom.onClick = function () {
                updateSpacingUI();
                rebuildPreviewIfNeeded();
            };
            sldSpacingJitter.onChanging = function () {
                /* ドラッグ中は値だけ更新し、プレビューは作り直さない / Update the value only; no rebuild while dragging */
                spacingJitterRatio = clampJitterRatio(sldSpacingJitter.value);
            };
            sldSpacingJitter.onChange = function () {
                spacingJitterRatio = clampJitterRatio(sldSpacingJitter.value);
                /* 新しい強さでランダム間隔を作り直す / Regenerate the random spacing with the new strength */
                lastRandomSpacing = null;
                rebuildPreviewIfNeeded();
            };

            cbAllRandom.onClick = function () { onAllRandomClicked(); };
            cbPreview.onClick = function () { onPreviewClicked(); };
        }

        // -----------------------------------------
        // ダイアログの状態 / Dialog state
        // -----------------------------------------

        /**
         * 複製の欄を、チェックボックスに合わせて有効／無効にする（値はそのまま）
         * @returns {void}
         */
        function updateDuplicateUI() {
            var isEnabled = !!cbDuplicateEnabled.value;
            etDuplicateCount.enabled = isEnabled;
            etDuplicateCount.stepperGroup.enabled = isEnabled;
            redrawSteppersIn(etDuplicateCount.stepperGroup);
            sldDuplicateCount.enabled = isEnabled;
        }

        /**
         * 複製数の欄の値を範囲内に直してスライダーにそろえ、プレビューを作り直す
         * @returns {void}
         */
        function syncDuplicateCountFromText() {
            var count = clampDuplicateCount(etDuplicateCount.text);
            etDuplicateCount.text = String(count);
            sldDuplicateCount.value = count;
            rebuildPreviewIfNeeded();
        }

        /**
         * 複製数を返す（複製しないときは 1）
         * @returns {number} 複製数
         */
        function getDuplicateCount() {
            if (!cbDuplicateEnabled.value) return 1;
            return clampDuplicateCount(etDuplicateCount.text);
        }

        /**
         * 回転モードのラジオを設定する（別グループのラジオも含めて排他にする）
         * @param {string} mode - "none" / "angle" / "perp" / "path_perp" / "random"
         * @returns {void}
         */
        function setRotationMode(mode) {
            rbRotNone.value = (mode === "none");
            rbRotAngle.value = (mode === "angle");
            rbRotPerp.value = (mode === "perp");
            rbRotPathPerp.value = (mode === "path_perp");
            rbRotRandom.value = (mode === "random");
            updateRotationUI();
        }

        /**
         * 角度の欄を「角度指定」のときだけ有効にする
         * @returns {void}
         */
        function updateRotationUI() {
            etRotAngle.enabled = !!rbRotAngle.value;
            etRotAngle.stepperGroup.enabled = etRotAngle.enabled;
            redrawSteppersIn(etRotAngle.stepperGroup);
        }

        /**
         * ばらつきのスライダーを「ランダム」のときだけ有効にする
         * @returns {void}
         */
        function updateSpacingUI() {
            sldSpacingJitter.enabled = !!rbSpacingRandom.value;
        }

        /**
         * 回転の設定を読み取る
         * @returns {{mode: string, angle: number, flip180: boolean}} 回転の設定
         */
        function getRotationSettings() {
            var mode = "none";
            if (rbRotAngle.value) mode = "angle";
            else if (rbRotPerp.value) mode = "perp";
            else if (rbRotPathPerp.value) mode = "path_perp";
            else if (rbRotRandom.value) mode = "random";

            var angle = Number(etRotAngle.text);
            if (isNaN(angle)) angle = 0;

            return { mode: mode, angle: angle, flip180: !!cbRotFlip180.value };
        }

        /**
         * 並べる順番のモードを返す
         * @returns {string} "current" / "reverse" / "random"
         */
        function getOrderMode() {
            if (rbOrderReverse.value) return "reverse";
            if (rbOrderRandom.value) return "random";
            return "current";
        }

        /**
         * 間隔のモードを返す
         * @returns {string} "even" / "random"
         */
        function getSpacingMode() {
            return rbSpacingRandom.value ? "random" : "even";
        }

        /**
         * 回転のラジオをクリックしたときの処理
         * @param {string} mode - 回転モード
         * @returns {void}
         */
        function onRotationModeClicked(mode) {
            setRotationMode(mode);

            if (mode === "angle") {
                /* 角度指定にしたら 10° を入れ、すぐ ↑↓ で変えられるようにする / Default to 10° and focus the field so ↑↓ works right away */
                etRotAngle.text = "10";
                try {
                    etRotAngle.active = true;
                    etRotAngle.selection = [0, etRotAngle.text.length];
                } catch (e) { /* 環境によってフォーカス／選択範囲を設定できない / focus or selection may be unsupported */ }
            }

            rebuildPreviewIfNeeded();
        }

        /**
         * 「一括ランダム」：オンで順番・間隔・回転をランダムに、オフで既定に戻す
         * @returns {void}
         */
        function onAllRandomClicked() {
            var isAllRandom = !!cbAllRandom.value;

            setRotationMode(isAllRandom ? "random" : "none");
            if (isAllRandom) {
                rbOrderRandom.value = true;
                rbSpacingRandom.value = true;
            } else {
                rbOrderCurrent.value = true;
                rbSpacingEven.value = true;
            }
            updateSpacingUI();

            /* ランダムの結果を作り直させる / Force new random results */
            lastRandomOrder = null;
            lastRandomSpacing = null;
            if (isAllRandom) lastShuffleSeed = null;

            rebuildPreviewIfNeeded();
        }

        // -----------------------------------------
        // 基準パスと配置対象 / Base path and objects
        // -----------------------------------------

        /**
         * 選択されている規則で基準パス（B）を探す
         * @param {PageItem[]} selectionItems - 選択
         * @returns {PathItem|null} 基準パス
         */
        function findBasePath(selectionItems) {
            if (rbAutoLargest.value) return getLargestPathItem(selectionItems);
            return findPathItemByStacking(selectionItems, !!rbFrontmost.value);
        }

        /**
         * プレビューに使う選択を返す。元を隠して選択が外れたときは控えを使う
         * @returns {Array|null} 2つ以上の選択。足りなければ null
         */
        function getPreviewSourceSelection() {
            var sourceSelection = doc.selection;
            if ((!sourceSelection || sourceSelection.length < 2) && previewSelectionSnapshot && previewSelectionSnapshot.length >= 2) {
                sourceSelection = previewSelectionSnapshot;
            }
            if (!sourceSelection || sourceSelection.length < 2) return null;
            return sourceSelection;
        }

        /**
         * プレビューを作れない理由を返す
         * @returns {string} 警告文の LABELS パス。作れるなら空文字
         */
        function findPreviewProblem() {
            var sourceSelection = getPreviewSourceSelection();
            if (!sourceSelection) return "alert.needSelection";
            var basePath = findBasePath(sourceSelection);
            if (!basePath) return "alert.noBasePath";
            if (collectPlaceableItems(sourceSelection, basePath).length === 0) return "alert.noItems";
            return "";
        }

        /**
         * 並べる順番を適用した配列を返す。ランダムは OK のときにプレビューの順を使い回す
         * @param {PageItem[]} sourceItems - 並べるオブジェクト
         * @param {string} orderMode - "current" / "reverse" / "random"
         * @param {boolean} isPreview - プレビューか
         * @returns {PageItem[]} 並べた配列
         */
        function applyOrderToArray(sourceItems, orderMode, isPreview) {
            var orderedItems = sortByStackingOrder(sourceItems);

            if (orderMode !== "random") {
                lastRandomOrder = null;
                if (orderMode === "reverse") orderedItems.reverse();
                return orderedItems;
            }

            if (!isPreview) {
                var reusedItems = reorderByStoredOrder(orderedItems, lastRandomOrder);
                if (reusedItems) return reusedItems;
            }

            /* 新しく混ぜる（最新の順はプレビューが作る）/ New shuffle; the preview produces the latest order */
            var shuffledItems = copyArray(orderedItems);
            shuffleInPlace(shuffledItems);
            lastRandomOrder = copyArray(shuffledItems);
            return shuffledItems;
        }

        /**
         * 各オブジェクトを複製して、元を含めて duplicateCount 個ずつにする
         * @param {PageItem[]} sourceItems - 元のオブジェクト
         * @param {number} duplicateCount - 1つあたりの個数
         * @returns {PageItem[]} 元と複製を並べた配列
         */
        function expandWithDuplicates(sourceItems, duplicateCount) {
            var expandedItems = [];
            for (var i = 0; i < sourceItems.length; i++) {
                expandedItems.push(sourceItems[i]);
                for (var copyIndex = 1; copyIndex < duplicateCount; copyIndex++) {
                    var duplicatedItem = duplicateBeside(sourceItems[i]);
                    if (duplicatedItem) expandedItems.push(duplicatedItem);
                }
            }
            return expandedItems;
        }

        /**
         * オブジェクトを複製する。親の後ろに置けないときは現在のレイヤーの末尾に置く
         * @param {PageItem} sourceItem - 元のオブジェクト
         * @returns {PageItem|null} 複製。できなければ null
         */
        function duplicateBeside(sourceItem) {
            try {
                return sourceItem.duplicate(sourceItem.parent, ElementPlacement.PLACEAFTER);
            } catch (e) {
                try {
                    return sourceItem.duplicate(doc.activeLayer, ElementPlacement.PLACEATEND);
                } catch (err) {
                    return null;
                }
            }
        }

        // -----------------------------------------
        // 配置 / Arrange
        // -----------------------------------------

        /**
         * ランダム間隔の位置を返す。OK ではプレビューと同じ位置を使い回す
         * @param {number} totalLength - パスの長さ
         * @param {number} itemCount - 配置する数
         * @param {boolean} isClosed - 閉じたパスか
         * @param {boolean} isPreview - プレビューか
         * @returns {number[]} 距離の配列
         */
        function getRandomDistances(totalLength, itemCount, isClosed, isPreview) {
            if (!isPreview && lastRandomSpacing && lastRandomSpacing.distances && lastRandomSpacing.distances.length === itemCount &&
                Math.abs(Number(lastRandomSpacing.totalLength) - Number(totalLength)) < 0.01) {
                return copyArray(lastRandomSpacing.distances);
            }
            var distances = computeJitteredDistances(totalLength, itemCount, isClosed, USE_ENDPOINTS, clampJitterRatio(spacingJitterRatio));
            lastRandomSpacing = { totalLength: totalLength, distances: copyArray(distances) };
            return distances;
        }

        /**
         * オブジェクトをパスに沿って配置し、回転する
         * @param {PathItem} pathItem - 基準パス
         * @param {PageItem[]} targetItems - 配置するオブジェクト（この順に始点から並べる）
         * @param {{mode: string, angle: number, flip180: boolean}} rotationSettings - 回転の設定
         * @param {string} spacingMode - "even" / "random"
         * @param {boolean} isPreview - プレビューか
         * @returns {boolean} 配置できたら true
         */
        function arrangeAlongPath(pathItem, targetItems, rotationSettings, spacingMode, isPreview) {
            var pathPoints = pathItem.pathPoints;
            if (!pathPoints || pathPoints.length < 2) {
                showDedupedAlert("alert.pathTooShort");
                return false;
            }

            var polyline = buildPolylineFromBezierPath(pathItem, SAMPLES_PER_SEGMENT);
            if (polyline.length < 2) {
                showDedupedAlert("alert.pathAnalyzeFailed");
                return false;
            }

            var cumulativeLengths = getCumulativeLengths(polyline);
            var totalLength = cumulativeLengths[cumulativeLengths.length - 1];
            if (totalLength <= 0) {
                showDedupedAlert("alert.pathLengthZero");
                return false;
            }

            var itemCount = targetItems.length;
            var isClosedPath = !!pathItem.closed;
            var pathGeometry = { polyline: polyline, cumulativeLengths: cumulativeLengths, isClosed: isClosedPath };

            /* 「それぞれ垂直」の中心 / Center for the Perpendicular mode */
            var baseCenter = getItemCenter(pathItem);

            var distances;
            if (spacingMode === "random") {
                distances = getRandomDistances(totalLength, itemCount, isClosedPath, isPreview);
            } else {
                lastRandomSpacing = null;
                distances = computeEvenDistances(totalLength, itemCount, isClosedPath, USE_ENDPOINTS);
            }

            for (var j = 0; j < itemCount; j++) {
                var distance = distances[j];
                var position = pointAtDistance(polyline, cumulativeLengths, distance);

                var itemCenter = getItemCenter(targetItems[j]);
                var dx = position.x - itemCenter.x;
                var dy = position.y - itemCenter.y;
                targetItems[j].translate(dx, dy);

                rotatePlacedItem(targetItems[j], rotationSettings, baseCenter, pathGeometry, distance);

                if (ROTATE_ALONG_TANGENT) {
                    var aheadPosition = pointAtDistance(polyline, cumulativeLengths, Math.min(totalLength, distance + totalLength * 0.001));
                    var tangentAngle = Math.atan2(aheadPosition.y - position.y, aheadPosition.x - position.x) * 180 / Math.PI;
                    targetItems[j].rotate(tangentAngle, true, true, true, true, Transformation.CENTER);
                }
            }
            return true;
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * プレビュー中に隠したオブジェクトの表示状態を戻す
         * @returns {void}
         */
        function restoreHiddenItems() {
            for (var i = 0; i < previewHiddenEntries.length; i++) {
                try {
                    previewHiddenEntries[i].item.hidden = previewHiddenEntries[i].wasHidden;
                } catch (e) { /* 既に無いオブジェクト / the item may be gone */ }
            }
            previewHiddenEntries = [];
        }

        /**
         * プレビュー中だけ元のオブジェクトを隠す（元の表示状態を控える）
         * @param {PageItem} pageItem - 隠すオブジェクト
         * @returns {void}
         */
        function hideForPreview(pageItem) {
            if (!pageItem) return;
            for (var i = 0; i < previewHiddenEntries.length; i++) {
                if (previewHiddenEntries[i].item === pageItem) return;
            }
            var wasHidden = false;
            try { wasHidden = !!pageItem.hidden; } catch (e) { wasHidden = false; }
            previewHiddenEntries.push({ item: pageItem, wasHidden: wasHidden });
            try { pageItem.hidden = true; } catch (e) { }
        }

        /**
         * プレビューを消し、隠したオブジェクトと選択を戻す
         * @returns {void}
         */
        function clearPreview() {
            if (previewLayer) {
                try {
                    previewLayer.locked = false;
                    previewLayer.remove();
                } catch (e) { /* 既に消えている / already gone */ }
            }
            previewLayer = null;
            restoreHiddenItems();

            /* 元を隠したときに選択が外れていたら戻す / Restore the selection if hiding the originals dropped it */
            try {
                var currentSelection = doc.selection;
                if ((!currentSelection || currentSelection.length === 0) && previewSelectionSnapshot && previewSelectionSnapshot.length) {
                    doc.selection = previewSelectionSnapshot;
                }
            } catch (e) { }

            app.redraw();
        }

        /**
         * 作りかけのプレビュー用レイヤーを捨てる
         * @returns {void}
         */
        function discardPreviewLayer() {
            try { previewLayer.remove(); } catch (e) { }
            previewLayer = null;
        }

        /**
         * 各オブジェクトを duplicateCount 個ずつ、グループの末尾に複製する
         * @param {PageItem[]} sourceItems - 元のオブジェクト
         * @param {GroupItem} targetGroup - 複製先のグループ
         * @param {number} duplicateCount - 1つあたりの複製数
         * @returns {PageItem[]} 複製
         */
        function duplicateIntoGroup(sourceItems, targetGroup, duplicateCount) {
            var duplicatedItems = [];
            for (var i = 0; i < sourceItems.length; i++) {
                for (var copyIndex = 0; copyIndex < duplicateCount; copyIndex++) {
                    try {
                        duplicatedItems.push(sourceItems[i].duplicate(targetGroup, ElementPlacement.PLACEATEND));
                    } catch (e) { /* 複製できないものは飛ばす / skip items that cannot be duplicated */ }
                }
            }
            return duplicatedItems;
        }

        /**
         * プレビューを作り直す（複製を最前面の一時レイヤーに並べ、元は隠す）
         * @returns {boolean} 作れたら true
         */
        function buildPreview() {
            clearPreview();

            var sourceSelection = getPreviewSourceSelection();
            if (!sourceSelection) return false;
            previewSelectionSnapshot = copyArray(sourceSelection);

            var basePath = findBasePath(sourceSelection);
            if (!basePath) return false;

            var sourceItems = collectPlaceableItems(sourceSelection, basePath);
            if (sourceItems.length === 0) return false;

            var orderMode = getOrderMode();
            var orderedSourceItems = applyOrderToArray(sourceItems, orderMode, true);
            var duplicateCount = getDuplicateCount();

            /* プレビュー用レイヤーを最前面に作る / Create the preview layer on top */
            try {
                previewLayer = doc.layers.add();
                previewLayer.name = PREVIEW_LAYER_NAME;
            } catch (e) {
                previewLayer = null;
                return false;
            }

            /* 形状の計算用に基準パスを複製 / Duplicate the base path for the geometry */
            var previewGroup = null;
            var previewPath = null;
            try {
                previewGroup = previewLayer.groupItems.add();
                previewGroup.name = "Preview_ArrangedAlongPath";
                previewPath = basePath.duplicate(previewGroup, ElementPlacement.PLACEATBEGINNING);
            } catch (e) {
                discardPreviewLayer();
                return false;
            }

            /* 「塗り／線」なしはプレビュー側だけに反映 / Apply "no fill / no stroke" to the preview copy only */
            if (rbBasePathHide.value) clearFillAndStroke(previewPath);

            var previewItems = duplicateIntoGroup(orderedSourceItems, previewGroup, duplicateCount);
            if (previewItems.length === 0) {
                clearPreview();
                return false;
            }

            /* ランダム順で複製ありなら、複製も含めてまとめて混ぜ、OK 用にシードを控える / Shuffle all copies together and keep the seed for OK */
            if (orderMode === "random" && duplicateCount > 1) {
                lastShuffleSeed = { seed: createShuffleSeed(), itemCount: previewItems.length, duplicateCount: duplicateCount };
                shuffleInPlace(previewItems, createSeededRandom(lastShuffleSeed.seed));
            } else {
                lastShuffleSeed = null;
            }

            if (!arrangeAlongPath(previewPath, previewItems, getRotationSettings(), getSpacingMode(), true)) {
                clearPreview();
                return false;
            }

            /* 「削除」なら配置後にプレビューのパスを消す / With Delete, remove the preview path after arranging */
            if (rbBasePathDelete.value) {
                try { previewPath.remove(); } catch (e) { }
            }

            /* プレビュー中は元のオブジェクトを隠す / Hide the originals while previewing */
            hideForPreview(basePath);
            for (var h = 0; h < sourceItems.length; h++) hideForPreview(sourceItems[h]);

            /* 選択を戻す（複製で変わることがある）/ Restore the selection, which duplication may change */
            try { doc.selection = previewSelectionSnapshot; } catch (e) { }
            app.redraw();

            return true;
        }

        /**
         * プレビューがオンなら作り直す。作れなければプレビューをオフにする
         * @returns {void}
         */
        function rebuildPreviewIfNeeded() {
            if (!cbPreview.value) return;
            var isBuilt = false;
            try {
                isBuilt = buildPreview();
            } catch (e) { /* 想定外の DOM エラーでもダイアログは閉じない / keep the dialog alive on unexpected DOM errors */ }
            if (!isBuilt) {
                cbPreview.value = false;
                clearPreview();
            }
        }

        /**
         * 「プレビュー」をクリックしたときの処理。作る前に選択を確かめ、警告が重ならないようにする
         * @returns {void}
         */
        function onPreviewClicked() {
            if (!cbPreview.value) {
                clearPreview();
                return;
            }

            var problemPath = findPreviewProblem();
            if (problemPath) {
                showDedupedAlert(problemPath);
                cbPreview.value = false;
                clearPreview();
                return;
            }

            rebuildPreviewIfNeeded();
        }

        // -----------------------------------------
        // 実行 / Apply
        // -----------------------------------------

        /**
         * OK で閉じたあと、選択したオブジェクトをパスに沿って配置する
         * @returns {void}
         */
        function applyArrangement() {
            /* OK の時点の選択を取り直す / Re-read the selection at OK time */
            var currentSelection = doc.selection;

            if (!currentSelection || currentSelection.length < 2) {
                /* プレビュー中に選択が外れていたら控えから戻す / Fall back to the snapshot if the preview dropped the selection */
                if (previewSelectionSnapshot && previewSelectionSnapshot.length >= 2) {
                    try { doc.selection = previewSelectionSnapshot; } catch (e) { }
                    try { currentSelection = doc.selection; } catch (e) { }
                }
            }

            if (!currentSelection || currentSelection.length < 2) {
                showDedupedAlert("alert.needSelection");
                return;
            }

            /* B = 選ばれた規則で決まる基準パス / B = base path by the chosen rule */
            var basePath = findBasePath(currentSelection);
            if (!basePath) {
                showDedupedAlert("alert.noBasePath");
                return;
            }

            if (rbBasePathHide.value) clearFillAndStroke(basePath);

            /* A = 基準パス（B）以外で配置できるもの / A = placeable items other than the base path (B) */
            var arrangeItems = collectPlaceableItems(currentSelection, basePath);
            if (arrangeItems.length === 0) {
                showDedupedAlert("alert.noItems");
                return;
            }

            var orderMode = getOrderMode();
            arrangeItems = applyOrderToArray(arrangeItems, orderMode, false);

            var duplicateCount = getDuplicateCount();
            if (duplicateCount > 1) {
                arrangeItems = expandWithDuplicates(arrangeItems, duplicateCount);

                /* ランダム順なら複製も含めて混ぜる。数が同じならプレビューのシードを使う / Shuffle all copies; reuse the preview seed when the counts match */
                if (orderMode === "random") {
                    var canReuseSeed = (lastShuffleSeed !== null && lastShuffleSeed.itemCount === arrangeItems.length && lastShuffleSeed.duplicateCount === duplicateCount);
                    var shuffleSeed = canReuseSeed ? lastShuffleSeed.seed : createShuffleSeed();
                    shuffleInPlace(arrangeItems, createSeededRandom(shuffleSeed));
                }
            }

            if (!arrangeAlongPath(basePath, arrangeItems, getRotationSettings(), getSpacingMode(), false)) return;

            /* 配置したオブジェクトをグループ化（基準パスは含めない）/ Group the placed objects, excluding the base path */
            if (cbGroupPlaced.value) {
                try {
                    var arrangedGroup = doc.groupItems.add();
                    arrangedGroup.name = "ArrangedAlongPath";
                    for (var i = 0; i < arrangeItems.length; i++) {
                        arrangeItems[i].move(arrangedGroup, ElementPlacement.PLACEATEND);
                    }
                } catch (e) { }
            }

            /* 基準パスを削除 / Remove the base path */
            if (rbBasePathDelete.value) {
                try { basePath.remove(); } catch (e) { }
            }
        }
    }

    main();

})();
