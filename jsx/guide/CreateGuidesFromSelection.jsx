#target illustrator
#targetengine "CreateGuidesFromSelectionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト・アートボード・カンバスを基準にガイドを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGuidesFromSelection.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nd1359cf41a2c

### Overview

Creates guides based on the selection, the artboard, or the canvas.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGuidesFromSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CreateGuidesFromSelection";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.10.5";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGuidesFromSelection.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGuidesFromSelection.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd1359cf41a2c"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* ガイドを作成するレイヤー名 / Layer that receives the guides */
    var GUIDE_LAYER_NAME = "_guide";

    /* ライブプレビュー用の一時レイヤー名 / Temporary layer for the live preview */
    var PREVIEW_LAYER_NAME = "__GuidePreview__";

    /* カンバス基準時にガイドを伸ばす長さ（pt）/ Guide reach for the canvas target (pt) */
    var CANVAS_GUIDE_REACH = 8000;

    /* プリセットラジオボタンの表示 / Show flags for the preset radio buttons */
    var showPresetTopBottom = false;
    var showPresetLeftRight = false;
    var showPresetTopLeft = true;
    var showPresetBottomLeft = true;
    var showPresetTopRight = false;
    var showPresetBottomRight = false;

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

    var CROSS_SPACING  = 20;                 /* 十字レイアウトの間隔 / gap inside the cross layout */
    var ROW_SPACING    = 4;                  /* 行内の要素間隔 / gap inside a row */

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
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
            title: { ja: "選択オブジェクトからガイド作成", en: "Create Guides from Selection" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            preset: { ja: "プリセット", en: "Presets" },
            destination: { ja: "描画先", en: "Destination" },
            options: { ja: "オプション", en: "Options" },
            axis: { ja: "ガイド位置（辺と中央）", en: "Guide Positions (Edges & Center)" }
        },
        radio: {
            artboard: { ja: "アートボード", en: "Artboard" },
            canvas: { ja: "カンバス（擬似）", en: "Canvas (Pseudo)" },
            selectionLayer: { ja: "選択オブジェクトと同じレイヤー", en: "Same layer as the selection" },
            guideLayer: { ja: "「_guide」レイヤー", en: "\"_guide\" layer" },
            allOn: { ja: "すべて", en: "All" },
            edges: { ja: "四辺", en: "Edges" },
            vertical: { ja: "上下", en: "Top & Bottom" },
            horizontal: { ja: "左右", en: "Left & Right" },
            topLeft: { ja: "左上", en: "Top Left" },
            bottomLeft: { ja: "左下", en: "Bottom Left" },
            topRight: { ja: "右上", en: "Top Right" },
            bottomRight: { ja: "右下", en: "Bottom Right" },
            centerBoth: { ja: "中心", en: "Center" },
            centerVertical: { ja: "中心線（垂直）", en: "Center line (vertical)" },
            centerHorizontal: { ja: "中心線（水平）", en: "Center line (horizontal)" },
            clear: { ja: "クリア", en: "Clear" }
        },
        checkbox: {
            left: { ja: "左", en: "Left" },
            top: { ja: "上", en: "Top" },
            center: { ja: "中心", en: "Center" },
            bottom: { ja: "下", en: "Bottom" },
            right: { ja: "右", en: "Right" },
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" },
            deleteGuides: { ja: "「_guide」レイヤーのガイドを削除", en: "Delete guides in \"_guide\"" },
            individual: { ja: "オブジェクトごとに作成", en: "Create per object" },
            group: { ja: "描画するガイドをグループ化", en: "Group the guides to draw" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            extension: { ja: "延長", en: "Extend" },
            offset: { ja: "選択オブジェクトとのマージン", en: "Margin from selection" }
        },
        button: {
            draw: { ja: "ガイドを描画", en: "Draw Guides" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            expandError: {
                ja: "アピアランス展開中にエラーが発生しました。",
                en: "An error occurred while expanding appearance."
            },
            noArtboard: { ja: "アートボードが存在しません。", en: "No artboard exists." },
            invalidArtboard: { ja: "有効なアートボードが選択されていません。", en: "No valid artboard selected." },
            guideError: {
                ja: "ガイド作成中にエラーが発生しました。",
                en: "An error occurred while creating guides."
            },
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." }
        },
        tooltip: {
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
            canvas: {
                ja: "アートボード範囲ではなく、広いカンバス範囲にガイドを引きます。",
                en: "Draw guides across a wide pseudo-canvas area instead of the active artboard."
            },
            artboard: {
                ja: "アクティブなアートボードの範囲に合わせてガイドを引きます。",
                en: "Draw guides within the active artboard area."
            },
            selectionLayer: {
                ja: "選択オブジェクトと同じレイヤーにガイドを作成します。選択がないときはアクティブレイヤーに作成します。",
                en: "Create the guides on the same layer as the selection. With no selection, the active layer is used."
            },
            guideLayer: {
                ja: "「_guide」レイヤーにまとめてガイドを作成し、作成後はレイヤーをロックします。",
                en: "Create the guides on the \"_guide\" layer and lock that layer afterwards."
            },
            extension: {
                ja: "アートボード基準時に、ガイド線をアートボード外へ伸ばす量です。カンバス基準では使用しません。",
                en: "Extends guide lines beyond the artboard when using the artboard target. Not used for the canvas target."
            },
            offset: {
                ja: "対象の外側へガイドを離す距離です。0の場合は対象の辺・中心に作成します。",
                en: "Distance to move guides away from the target. Use 0 to place them on the target edges or center."
            },
            usePreviewBounds: {
                ja: "線幅や効果など、見た目上の境界を基準にします。テキストは一時的にアウトライン化して計算します（ライブプレビューでは省略）。",
                en: "Use visual bounds including strokes and effects. Text is temporarily outlined for calculation (skipped during live preview)."
            },
            deleteGuides: {
                ja: "実行前に「_guide」レイヤー内の既存ガイドを削除します。ほかのレイヤーのガイドは対象外です。",
                en: "Before drawing, delete existing guides in the \"_guide\" layer only. Guides on other layers are not affected."
            },
            individual: {
                ja: "ON：選択オブジェクトごとにガイドを作成。OFF：選択全体の外接でまとめて1組作成。選択が1つ以下のときは使用しません。",
                en: "On: one set of guides per selected object. Off: one set for the combined selection bounds. Not used with 0–1 objects selected."
            },
            group: {
                ja: "作成したガイドを1つのグループにまとめます。オブジェクトごとに作成する場合は、オブジェクト単位でグループ化します。",
                en: "Put the created guides into a group. With per-object guides, each object gets its own group."
            },
            preview: {
                ja: "確定前に、作成予定のガイド位置を色付きの仮線で表示します。",
                en: "Show colored temporary lines for the guide positions before committing."
            },
            edge: {
                left: { ja: "対象の左辺にガイドを作成", en: "Add a guide at the target's left edge" },
                top: { ja: "対象の上辺にガイドを作成", en: "Add a guide at the target's top edge" },
                center: { ja: "対象の中央（縦・横）にガイドを作成", en: "Add guides at the target's center (vertical & horizontal)" },
                bottom: { ja: "対象の下辺にガイドを作成", en: "Add a guide at the target's bottom edge" },
                right: { ja: "対象の右辺にガイドを作成", en: "Add a guide at the target's right edge" },
                soloHint: {
                    ja: "option（Alt）＋クリックで、この項目だけをオンにします。",
                    en: "Option/Alt-click to turn this one on and all the others off."
                }
            },
            preset: {
                allOn: { ja: "四辺＋中央をすべて選択", en: "Select all four edges and the center" },
                edges: { ja: "上下左右の四辺を選択", en: "Select all four edges" },
                topBottom: { ja: "上下の辺のみ選択", en: "Top and bottom edges only" },
                leftRight: { ja: "左右の辺のみ選択", en: "Left and right edges only" },
                topLeft: { ja: "左上（左＋上）を選択", en: "Top-left (left + top)" },
                bottomLeft: { ja: "左下（左＋下）を選択", en: "Bottom-left (left + bottom)" },
                topRight: { ja: "右上（右＋上）を選択", en: "Top-right (right + top)" },
                bottomRight: { ja: "右下（右＋下）を選択", en: "Bottom-right (right + bottom)" },
                centerBoth: { ja: "中央（縦横）のみ選択", en: "Center only (vertical & horizontal)" },
                centerVertical: { ja: "垂直の中心線のみ作成", en: "Vertical center line only" },
                centerHorizontal: { ja: "水平の中心線のみ作成", en: "Horizontal center line only" },
                clear: { ja: "すべて解除", en: "Clear all" }
            }
        }
    };

    // =========================================
    // 単位 / Units
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
     * 入力文字列を定規単位として解釈し pt に変換
     * @param {string} inputText - 入力欄の文字列
     * @returns {number} pt 値（数値でない場合は 0）
     */
    function rulerTextToPoints(inputText) {
        var inputValue = parseFloat(inputText);
        if (isNaN(inputValue)) inputValue = 0;
        return inputValue * getUnitInfo("rulerType").pointsPerUnit;
    }

    // =========================================
    // 境界の取得 / Bounds
    // =========================================

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

    /**
     * アクティブなアートボードの矩形を取得
     * @param {Document} doc - 対象ドキュメント
     * @returns {number[]|null} [左, 上, 右, 下]（取得できない場合は null）
     */
    function getActiveArtboardRect(doc) {
        if (doc.artboards.length === 0) return null;
        var artboardIndex = doc.artboards.getActiveArtboardIndex();
        if (artboardIndex < 0 || artboardIndex >= doc.artboards.length) return null;
        return doc.artboards[artboardIndex].artboardRect;
    }

    /**
     * プレビュー境界使用時、テキストを一時的にアウトライン化して境界計算用に差し替える
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
     * @returns {object} { items: 境界計算用オブジェクト配列, restore: 復元関数 }
     */
    function outlineTextForBounds(selectedItems, usePreviewBounds) {
        var passthrough = { items: selectedItems, restore: function () {} };
        if (!usePreviewBounds) return passthrough;

        var boundsItems = [];
        var duplicatedTexts = [];
        var originalTexts = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem && selectedItem.typename === "TextFrame") {
                duplicatedTexts.push(selectedItem.duplicate());
                originalTexts.push(selectedItem);
            } else {
                boundsItems.push(selectedItem);
            }
        }
        if (duplicatedTexts.length === 0) return passthrough;

        for (var j = 0; j < originalTexts.length; j++) originalTexts[j].hidden = true;

        /* 全選択解除は app.selection = null が確実（doc.selection = null は MRAP エラーになることがある）/ Use app.selection to deselect */
        app.selection = null;
        for (var k = 0; k < duplicatedTexts.length; k++) duplicatedTexts[k].selected = true;
        try {
            app.executeMenuCommand('expandStyle');
        } catch (e) {
            alert(getLabel("alert.expandError") + "\n" + e.message);
        }

        var outlinedTexts = [];
        for (var m = 0; m < duplicatedTexts.length; m++) {
            var outlined = duplicatedTexts[m].createOutline();
            outlinedTexts.push(outlined ? outlined : duplicatedTexts[m]);
        }
        return {
            items: boundsItems.concat(outlinedTexts),
            restore: function () {
                for (var n = 0; n < outlinedTexts.length; n++) {
                    outlinedTexts[n].remove();
                    originalTexts[n].hidden = false;
                }
            }
        };
    }

    /**
     * ガイドの基準となる外接矩形と、その元オブジェクトが載るレイヤーを取得
     * @param {PageItem[]} boundsItems - 境界計算に使うオブジェクト
     * @param {object} guideOptions - ガイド作成オプション
     * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
     * @returns {object[]} { bounds: number[], layer: Layer|null } の配列
     */
    function collectTargetEntries(boundsItems, guideOptions, artboardRect) {
        /* 選択が無いときはアートボードを基準にする / Fall back to the artboard when nothing is selected */
        if (boundsItems.length === 0) return artboardRect ? [{ bounds: artboardRect.concat(), layer: null }] : [];

        if (guideOptions.individual) {
            var targetEntries = [];
            for (var i = 0; i < boundsItems.length; i++) {
                targetEntries.push({
                    bounds: getClipAwareBounds(boundsItems[i], guideOptions.usePreviewBounds),
                    layer: boundsItems[i].layer
                });
            }
            return targetEntries;
        }
        var combinedBounds = getClipAwareUnionBounds(boundsItems, guideOptions.usePreviewBounds);
        /* まとめて1組にするときは先頭オブジェクトのレイヤーを代表にする / Use the first item's layer for the combined set */
        return combinedBounds ? [{ bounds: combinedBounds, layer: boundsItems[0].layer }] : [];
    }

    // =========================================
    // ガイドの線分 / Guide segments
    // =========================================

    /**
     * 境界とオプションから引くべきガイドの位置と向きを算出
     * @param {number[]} bounds - 基準の外接矩形 [左, 上, 右, 下]
     * @param {object} guideOptions - ガイド作成オプション
     * @param {number} offsetPt - 対象から離す距離（pt）
     * @returns {object[]} { position: number, orientation: string } の配列
     */
    function directionsFromBounds(bounds, guideOptions, offsetPt) {
        var topPosition = bounds[1] + offsetPt;
        var leftPosition = bounds[0] - offsetPt;
        var bottomPosition = bounds[3] - offsetPt;
        var rightPosition = bounds[2] + offsetPt;
        var centerX = (leftPosition + rightPosition) / 2;
        var centerY = (topPosition + bottomPosition) / 2;

        var guideDirections = [];
        if (guideOptions.left) guideDirections.push({ position: leftPosition, orientation: "vertical" });
        if (guideOptions.right) guideDirections.push({ position: rightPosition, orientation: "vertical" });
        if (guideOptions.top) guideDirections.push({ position: topPosition, orientation: "horizontal" });
        if (guideOptions.bottom) guideDirections.push({ position: bottomPosition, orientation: "horizontal" });
        if (guideOptions.center) {
            guideDirections.push({ position: centerX, orientation: "vertical" });
            guideDirections.push({ position: centerY, orientation: "horizontal" });
        }
        if (guideOptions.centerMode === "vertical") guideDirections.push({ position: centerX, orientation: "vertical" });
        if (guideOptions.centerMode === "horizontal") guideDirections.push({ position: centerY, orientation: "horizontal" });
        return guideDirections;
    }

    /**
     * ガイド1本分の線分を算出
     * @param {number} position - ガイドの座標
     * @param {string} orientation - "vertical" または "horizontal"
     * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
     * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
     * @returns {number[][]} [始点, 終点]
     */
    function guideSegment(position, orientation, drawSettings, artboardRect) {
        if (drawSettings.useCanvas) {
            return (orientation === "horizontal")
                ? [[-CANVAS_GUIDE_REACH, position], [CANVAS_GUIDE_REACH, position]]
                : [[position, CANVAS_GUIDE_REACH], [position, -CANVAS_GUIDE_REACH]];
        }
        var extensionPt = drawSettings.extensionPt;
        return (orientation === "horizontal")
            ? [[artboardRect[0] - extensionPt, position], [artboardRect[2] + extensionPt, position]]
            : [[position, artboardRect[1] + extensionPt], [position, artboardRect[3] - extensionPt]];
    }

    /**
     * 基準矩形1つ分のガイドの線分を算出
     * @param {number[]} bounds - 基準の外接矩形 [左, 上, 右, 下]
     * @param {object} guideOptions - ガイド作成オプション
     * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
     * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
     * @returns {number[][][]} 線分（[始点, 終点]）の配列
     */
    function segmentsForBounds(bounds, guideOptions, drawSettings, artboardRect) {
        var guideDirections = directionsFromBounds(bounds, guideOptions, drawSettings.offsetPt);
        var segments = [];
        for (var i = 0; i < guideDirections.length; i++) {
            segments.push(guideSegment(guideDirections[i].position, guideDirections[i].orientation, drawSettings, artboardRect));
        }
        return segments;
    }

    /**
     * 基準矩形ごとに「描画先レイヤーと線分」の組を作る（本描画とプレビューで共通）
     * @param {object[]} targetEntries - { bounds: number[], layer: Layer|null } の配列
     * @param {object} guideOptions - ガイド作成オプション
     * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
     * @param {number[]|null} artboardRect - アクティブなアートボードの矩形
     * @returns {object[]} { layer: Layer|null, segments: number[][][] } の配列
     */
    function buildDrawPlan(targetEntries, guideOptions, drawSettings, artboardRect) {
        var drawPlan = [];
        for (var i = 0; i < targetEntries.length; i++) {
            drawPlan.push({
                layer: targetEntries[i].layer,
                segments: segmentsForBounds(targetEntries[i].bounds, guideOptions, drawSettings, artboardRect)
            });
        }
        return drawPlan;
    }

    // =========================================
    // ガイドの作成 / Guide creation
    // =========================================

    /**
     * 名前でレイヤーを検索
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー（無ければ null）
     */
    function findLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return doc.layers[i];
        }
        return null;
    }

    /**
     * ガイド用レイヤーを取得（無ければ作成）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} ガイド用レイヤー
     */
    function getOrCreateGuideLayer(doc) {
        var guideLayer = findLayerByName(doc, GUIDE_LAYER_NAME);
        if (!guideLayer) {
            guideLayer = doc.layers.add();
            guideLayer.name = GUIDE_LAYER_NAME;
        }
        return guideLayer;
    }

    /**
     * レイヤー・グループ内の既存ガイドを削除（グループ化されたガイドも対象）
     * @param {Layer|GroupItem} guideContainer - 対象レイヤーまたはグループ
     * @returns {void}
     */
    function deleteGuidesInContainer(guideContainer) {
        for (var i = guideContainer.pageItems.length - 1; i >= 0; i--) {
            var containedItem = guideContainer.pageItems[i];
            if (containedItem.typename === "GroupItem") {
                deleteGuidesInContainer(containedItem);
                /* 中身が無くなったグループは残さない / Drop groups left empty */
                if (containedItem.pageItems.length === 0) containedItem.remove();
            } else if (containedItem.guides) {
                containedItem.remove();
            }
        }
    }

    /**
     * 線分の配列からガイドを作成
     * @param {Layer} targetLayer - 作成先レイヤー
     * @param {number[][][]} segments - 線分（[始点, 終点]）の配列
     * @param {boolean} groupGuides - 作成したガイドをグループにまとめるかどうか
     * @returns {void}
     */
    function addGuidesFromSegments(targetLayer, segments, groupGuides) {
        if (segments.length === 0) return;

        /* グループ化するときはガイドをグループ内に直接作成 / Create the guides inside a group when grouping */
        var guideContainer = groupGuides ? targetLayer.groupItems.add() : targetLayer;
        for (var i = 0; i < segments.length; i++) {
            var guidePath = guideContainer.pathItems.add();
            guidePath.setEntirePath(segments[i]);
            guidePath.filled = false;
            guidePath.stroked = false;
            guidePath.guides = true;
        }
    }

    /**
     * レイヤーのロック状態を記録して解除（記録済みなら何もしない）
     * @param {Layer} targetLayer - 対象レイヤー
     * @param {object[]} lockStates - { layer, wasLocked } の記録先
     * @returns {void}
     */
    function unlockLayerOnce(targetLayer, lockStates) {
        for (var i = 0; i < lockStates.length; i++) {
            if (lockStates[i].layer === targetLayer) return;
        }
        lockStates.push({ layer: targetLayer, wasLocked: targetLayer.locked });
        targetLayer.locked = false;
    }

    /**
     * 選択オブジェクト・アートボード・カンバスを基準にガイドを作成（本処理）
     * @param {object} guideOptions - ガイド作成オプション（各方向フラグ・プレビュー境界・個別作成・グループ化・描画先・既存削除）
     * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
     * @returns {void}
     */
    function createGuides(guideOptions, drawSettings) {
        var doc = app.activeDocument;
        var artboardRect = getActiveArtboardRect(doc);
        if (!drawSettings.useCanvas && !artboardRect) {
            alert(getLabel(doc.artboards.length === 0 ? "alert.noArtboard" : "alert.invalidArtboard"));
            return;
        }

        /* 境界計算のあいだだけテキストをアウトライン化 / Outline text only while measuring */
        var outlinedTextState = outlineTextForBounds(doc.selection, guideOptions.usePreviewBounds);
        var targetEntries = collectTargetEntries(outlinedTextState.items, guideOptions, artboardRect);
        /* 線分はアウトラインを戻す前に確定させる / Freeze the segments before restoring the outlined text */
        var drawPlan = buildDrawPlan(targetEntries, guideOptions, drawSettings, artboardRect);
        outlinedTextState.restore();

        /* 「_guide」レイヤーへまとめる場合は先に用意して既存ガイドを整理 / Prepare the "_guide" layer up front */
        var guideLayer = null;
        if (!guideOptions.drawOnSelectionLayer) {
            guideLayer = getOrCreateGuideLayer(doc);
            guideLayer.locked = false;
            if (guideOptions.deleteExisting) deleteGuidesInContainer(guideLayer);
        }

        var lockStates = [];
        for (var i = 0; i < drawPlan.length; i++) {
            var targetLayer = guideLayer || drawPlan[i].layer || doc.activeLayer;
            unlockLayerOnce(targetLayer, lockStates);
            addGuidesFromSegments(targetLayer, drawPlan[i].segments, guideOptions.groupGuides);
        }

        /* 「_guide」レイヤーはロックし、ほかは元の状態へ戻す / Lock "_guide", restore the others */
        for (var j = 0; j < lockStates.length; j++) {
            if (lockStates[j].layer !== guideLayer) lockStates[j].layer.locked = lockStates[j].wasLocked;
        }
        if (guideLayer) guideLayer.locked = true;
    }

    // =========================================
    // ライブプレビュー / Live preview
    // =========================================

    /**
     * プレビュー用の色を作成（ドキュメントのカラースペースに合わせる）
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} プレビュー線の色
     */
    function makePreviewColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 90;
            cmykColor.yellow = 0;
            cmykColor.black = 0;
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = 255;
        rgbColor.green = 0;
        rgbColor.blue = 255;
        return rgbColor;
    }

    /**
     * プレビュー用レイヤーを削除
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removePreviewLayer(doc) {
        var previewLayer = findLayerByName(doc, PREVIEW_LAYER_NAME);
        if (!previewLayer) return;
        previewLayer.locked = false;
        previewLayer.visible = true;
        previewLayer.remove();
    }

    /**
     * プレビュー用レイヤーを用意（既存があれば作り直す）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} プレビュー用レイヤー
     */
    function createPreviewLayer(doc) {
        removePreviewLayer(doc);
        var previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;
        return previewLayer;
    }

    /**
     * 線分を色付きの仮線としてレイヤーへ描画
     * @param {Layer} previewLayer - 描画先レイヤー
     * @param {number[][][]} segments - 線分（[始点, 終点]）の配列
     * @param {CMYKColor|RGBColor} previewColor - 仮線の色
     * @returns {void}
     */
    function drawPreviewSegments(previewLayer, segments, previewColor) {
        for (var i = 0; i < segments.length; i++) {
            var previewPath = previewLayer.pathItems.add();
            previewPath.setEntirePath([segments[i][0], segments[i][1]]);
            previewPath.filled = false;
            previewPath.stroked = true;
            previewPath.strokeColor = previewColor;
            previewPath.strokeWidth = 1;
        }
    }

    /**
     * 現在の設定からプレビュー線分を収集（テキストのアウトライン化は省略）
     * @param {object} guideOptions - ガイド作成オプション
     * @param {object} drawSettings - 描画設定 { useCanvas: boolean, offsetPt: number, extensionPt: number }
     * @returns {number[][][]} 線分（[始点, 終点]）の配列
     */
    function collectPreviewSegments(guideOptions, drawSettings) {
        var doc = app.activeDocument;
        var artboardRect = getActiveArtboardRect(doc);
        if (!drawSettings.useCanvas && !artboardRect) return [];

        var targetEntries = collectTargetEntries(doc.selection, guideOptions, artboardRect);
        var drawPlan = buildDrawPlan(targetEntries, guideOptions, drawSettings, artboardRect);
        var segments = [];
        for (var i = 0; i < drawPlan.length; i++) {
            segments = segments.concat(drawPlan[i].segments);
        }
        return segments;
    }

    // =========================================
    // UI 部品 / UI parts
    // =========================================

    /**
     * ツールチップ付きチェックボックスを追加
     * @param {Group|Panel} parentContainer - 追加先コンテナ
     * @param {string} labelKey - ラベルのドット区切りキー
     * @param {string} tooltipKey - ツールチップのドット区切りキー
     * @param {boolean} initialValue - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, labelKey, tooltipKey, initialValue) {
        var checkbox = parentContainer.add("checkbox", undefined, getLabel(labelKey));
        checkbox.helpTip = getLabel(tooltipKey);
        checkbox.value = initialValue;
        return checkbox;
    }

    /**
     * 「ラベル＋数値入力＋単位」の行を追加
     * @param {Group|Panel} parentContainer - 追加先コンテナ
     * @param {string} labelKey - ラベルのドット区切りキー
     * @param {string} tooltipKey - ツールチップのドット区切りキー
     * @param {string} initialValue - 入力欄の初期値
     * @returns {object} { row: Group, input: EditText }
     */
    function addUnitInputRow(parentContainer, labelKey, tooltipKey, initialValue) {
        var inputRow = parentContainer.add("group");
        setupRow(inputRow, "left", ROW_SPACING);
        var tooltipText = getLabel(tooltipKey);

        var rowLabel = inputRow.add("statictext", undefined, labelText(labelKey));
        rowLabel.helpTip = tooltipText;

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = inputRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var valueInput;
        /* マイナスも受け付ける。値を変えたら onChanging（プレビュー更新）を呼ぶ
           Negative values are allowed; run onChanging (preview refresh) after each step */
        var valueStepper = addStepper(stepperInputGroup, function () { return valueInput; }, {
            onStep: function (numberInput) {
                if (typeof numberInput.onChanging === "function") numberInput.onChanging();
            }
        });
        valueInput = stepperInputGroup.add("edittext", undefined, initialValue);
        valueInput.characters = 3;
        valueInput.helpTip = tooltipText;
        bindSteppedArrowKeys(valueInput, valueStepper);
        var unitLabel = inputRow.add("statictext", undefined, getUnitInfo("rulerType").label);
        unitLabel.helpTip = tooltipText;

        return { row: inputRow, input: valueInput };
    }

    /**
     * 「対象」パネル（基準ラジオ＋延長）を構築
     * @param {Group} parentContainer - 追加先コンテナ
     * @returns {object} { canvasRadio, artboardRadio, extensionInput, updateExtensionEnabled }
     */
    function buildTargetPanel(parentContainer) {
        var targetPanel = parentContainer.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);

        /* カンバス→アートボードの順に並べ、デフォルトはアートボード / Canvas → Artboard, default to Artboard */
        var canvasRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.canvas"));
        canvasRadio.helpTip = getLabel("tooltip.canvas");
        var artboardRadio = targetPanel.add("radiobutton", undefined, getLabel("radio.artboard"));
        artboardRadio.helpTip = getLabel("tooltip.artboard");
        artboardRadio.value = true;

        var extensionRow = addUnitInputRow(targetPanel, "fieldLabel.extension", "tooltip.extension", "20");

        /**
         * 「延長」行のディムを更新する（カンバス基準では使わないので行ごとディムする）
         * @returns {void}
         */
        function updateExtensionEnabled() {
            extensionRow.row.enabled = !canvasRadio.value;
            redrawSteppersIn(extensionRow.row); /* 自作描画の∧∨を描き直す / redraw the custom-drawn stepper */
        }
        artboardRadio.onClick = updateExtensionEnabled;
        canvasRadio.onClick = updateExtensionEnabled;
        updateExtensionEnabled();

        return {
            canvasRadio: canvasRadio,
            artboardRadio: artboardRadio,
            extensionInput: extensionRow.input,
            updateExtensionEnabled: updateExtensionEnabled
        };
    }

    /**
     * 「ガイド位置」パネル（十字のチェックボックス）を構築
     * @param {Group} parentContainer - 追加先コンテナ
     * @returns {object} { left, top, right, bottom, center } のチェックボックス
     */
    function buildAxisPanel(parentContainer) {
        var axisPanel = parentContainer.add("panel", undefined, getLabel("panel.axis"));
        setupPanel(axisPanel, 6);

        /**
         * 辺・中心のチェックボックスを追加する（option＋クリックの説明をツールチップに添える）
         * @param {Group|Panel} targetContainer - 追加先コンテナ
         * @param {string} labelKey - ラベルのドット区切りキー
         * @param {string} tooltipKey - ツールチップのドット区切りキー
         * @param {boolean} initialValue - 初期状態
         * @returns {Checkbox} 追加したチェックボックス
         */
        function addEdgeCheckbox(targetContainer, labelKey, tooltipKey, initialValue) {
            var checkbox = addCheckbox(targetContainer, labelKey, tooltipKey, initialValue);
            checkbox.helpTip += "\n" + getLabel("tooltip.edge.soloHint");
            return checkbox;
        }

        var crossGroup = axisPanel.add("group");
        setupRow(crossGroup, undefined, CROSS_SPACING);
        /* パネル内で左右中央に配置 / Center the cross horizontally in the panel */
        crossGroup.alignment = ["center", "top"];

        var leftColumn = crossGroup.add("group");
        setupRow(leftColumn, undefined, 10);
        var leftCheckbox = addEdgeCheckbox(leftColumn, "checkbox.left", "tooltip.edge.left", true);

        var centerColumn = crossGroup.add("group");
        centerColumn.orientation = "column";
        centerColumn.alignChildren = ["left", "center"];
        centerColumn.spacing = 10;
        var topCheckbox = addEdgeCheckbox(centerColumn, "checkbox.top", "tooltip.edge.top", true);
        /* 中心（縦横）の中心線用 / Center lines (vertical & horizontal) */
        var centerCheckbox = addEdgeCheckbox(centerColumn, "checkbox.center", "tooltip.edge.center", false);
        var bottomCheckbox = addEdgeCheckbox(centerColumn, "checkbox.bottom", "tooltip.edge.bottom", true);

        var rightColumn = crossGroup.add("group");
        setupRow(rightColumn, undefined, 10);
        var rightCheckbox = addEdgeCheckbox(rightColumn, "checkbox.right", "tooltip.edge.right", true);

        return {
            left: leftCheckbox,
            top: topCheckbox,
            right: rightCheckbox,
            bottom: bottomCheckbox,
            center: centerCheckbox
        };
    }

    /**
     * プリセットのラジオボタンを生成して十字チェックボックスと連動させる
     * @param {Group} parentContainer - 追加先コンテナ
     * @param {object} crossCheckboxes - 十字のチェックボックス群
     * @param {object} centerLineState - 中心線モードの保持オブジェクト（{ mode: string }）
     * @returns {object} 生成したラジオボタン群
     */
    function buildPresetRadios(parentContainer, crossCheckboxes, centerLineState) {
        /**
         * プリセットのラジオボタンを追加する（show が false のときは作らない）
         * @param {string} labelKey - ラベルのドット区切りキー
         * @param {string} tooltipKey - ツールチップのドット区切りキー
         * @param {boolean} [show] - 表示するかどうか（省略時は表示）
         * @returns {RadioButton|null} 追加したラジオボタン（作らなかった場合は null）
         */
        function addPresetRadio(labelKey, tooltipKey, show) {
            if (typeof show !== "undefined" && !show) return null;
            var presetRadio = parentContainer.add("radiobutton", undefined, getLabel(labelKey));
            presetRadio.helpTip = getLabel(tooltipKey);
            return presetRadio;
        }

        var presetRadios = {
            allOn: addPresetRadio("radio.allOn", "tooltip.preset.allOn"),
            edges: addPresetRadio("radio.edges", "tooltip.preset.edges"),
            topBottom: addPresetRadio("radio.vertical", "tooltip.preset.topBottom", showPresetTopBottom),
            leftRight: addPresetRadio("radio.horizontal", "tooltip.preset.leftRight", showPresetLeftRight),
            topLeft: addPresetRadio("radio.topLeft", "tooltip.preset.topLeft", showPresetTopLeft),
            bottomLeft: addPresetRadio("radio.bottomLeft", "tooltip.preset.bottomLeft", showPresetBottomLeft),
            topRight: addPresetRadio("radio.topRight", "tooltip.preset.topRight", showPresetTopRight),
            bottomRight: addPresetRadio("radio.bottomRight", "tooltip.preset.bottomRight", showPresetBottomRight),
            centerBoth: addPresetRadio("radio.centerBoth", "tooltip.preset.centerBoth"),
            centerVertical: addPresetRadio("radio.centerVertical", "tooltip.preset.centerVertical"),
            centerHorizontal: addPresetRadio("radio.centerHorizontal", "tooltip.preset.centerHorizontal"),
            clear: addPresetRadio("radio.clear", "tooltip.preset.clear")
        };

        /* プリセット定義（crossValues=[左,上,右,下,中心], centerLine=中心線モード）/ Preset table */
        var presetDefinitions = [
            { radio: presetRadios.allOn,            crossValues: [true,  true,  true,  true,  true ] },
            { radio: presetRadios.edges,            crossValues: [true,  true,  true,  true,  false] },
            { radio: presetRadios.topBottom,        crossValues: [false, true,  false, true,  false] },
            { radio: presetRadios.leftRight,        crossValues: [true,  false, true,  false, false] },
            { radio: presetRadios.topLeft,          crossValues: [true,  true,  false, false, false] },
            { radio: presetRadios.bottomLeft,       crossValues: [true,  false, false, true,  false] },
            { radio: presetRadios.topRight,         crossValues: [false, true,  true,  false, false] },
            { radio: presetRadios.bottomRight,      crossValues: [false, false, true,  true,  false] },
            { radio: presetRadios.clear,            crossValues: [false, false, false, false, false] },
            { radio: presetRadios.centerBoth,       crossValues: [false, false, false, false, true ] },
            { radio: presetRadios.centerVertical,   crossValues: [false, false, false, false, false], centerLine: "vertical" },
            { radio: presetRadios.centerHorizontal, crossValues: [false, false, false, false, false], centerLine: "horizontal" }
        ];
        for (var i = 0; i < presetDefinitions.length; i++) {
            (function (presetDefinition) {
                if (!presetDefinition.radio) return;
                presetDefinition.radio.onClick = function () {
                    if (!presetDefinition.radio.value) return;
                    var crossValues = presetDefinition.crossValues;
                    centerLineState.mode = presetDefinition.centerLine || "";
                    crossCheckboxes.left.value = crossValues[0];
                    crossCheckboxes.top.value = crossValues[1];
                    crossCheckboxes.right.value = crossValues[2];
                    crossCheckboxes.bottom.value = crossValues[3];
                    crossCheckboxes.center.value = crossValues[4];
                };
            })(presetDefinitions[i]);
        }

        /* デフォルト選択は四辺を優先 / Default selection (prefer "Edges") */
        if (presetRadios.edges) {
            presetRadios.edges.value = true;
        } else if (presetRadios.allOn) {
            presetRadios.allOn.value = true;
        }

        /* 手動でチェックを変えたら中心線モードを解除し、option＋クリックはクリックした項目だけをオンに
           / Clear center-line mode on manual toggle; Option-click keeps only the clicked item on */
        var crossKeys = ["left", "top", "right", "bottom", "center"];

        /**
         * 十字のチェックボックス1つにクリック時の処理を割り当てる
         * @param {string} crossKey - 対象のキー（"left" / "top" / "right" / "bottom" / "center"）
         * @returns {void}
         */
        function setupCrossCheckbox(crossKey) {
            crossCheckboxes[crossKey].onClick = function () {
                centerLineState.mode = "";
                if (!ScriptUI.environment.keyboardState.altKey) return;
                for (var k = 0; k < crossKeys.length; k++) {
                    crossCheckboxes[crossKeys[k]].value = (crossKeys[k] === crossKey);
                }
            };
        }
        for (var j = 0; j < crossKeys.length; j++) {
            setupCrossCheckbox(crossKeys[j]);
        }

        return presetRadios;
    }

    /**
     * 「描画先」パネルを構築
     * @param {Window} dialog - 追加先ダイアログ
     * @returns {object} { selectionLayerRadio, guideLayerRadio }
     */
    function buildDestinationPanel(dialog) {
        var destinationPanel = dialog.add("panel", undefined, getLabel("panel.destination"));
        setupPanel(destinationPanel, 6);

        var selectionLayerRadio = destinationPanel.add("radiobutton", undefined, getLabel("radio.selectionLayer"));
        selectionLayerRadio.helpTip = getLabel("tooltip.selectionLayer");
        var guideLayerRadio = destinationPanel.add("radiobutton", undefined, getLabel("radio.guideLayer"));
        guideLayerRadio.helpTip = getLabel("tooltip.guideLayer");
        guideLayerRadio.value = true;

        return { selectionLayerRadio: selectionLayerRadio, guideLayerRadio: guideLayerRadio };
    }

    /**
     * 「オプション」パネルを構築
     * @param {Window} dialog - 追加先ダイアログ
     * @param {number} selectionCount - 選択オブジェクト数
     * @returns {object} { usePreviewBoundsCheckbox, deleteGuidesCheckbox, individualCheckbox, groupCheckbox, offsetRow, offsetInput }
     */
    function buildOptionsPanel(dialog, selectionCount) {
        var optionsPanel = dialog.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel, 6);

        var usePreviewBoundsCheckbox = addCheckbox(optionsPanel, "checkbox.usePreviewBounds", "tooltip.usePreviewBounds", true);
        var deleteGuidesCheckbox = addCheckbox(optionsPanel, "checkbox.deleteGuides", "tooltip.deleteGuides", true);
        var individualCheckbox = addCheckbox(optionsPanel, "checkbox.individual", "tooltip.individual", false);
        /* 選択が1つ以下ならまとめて作成と変わらないのでディム / Dim when 0–1 objects are selected */
        if (selectionCount <= 1) individualCheckbox.enabled = false;
        var groupCheckbox = addCheckbox(optionsPanel, "checkbox.group", "tooltip.group", true);

        var offsetRow = addUnitInputRow(optionsPanel, "fieldLabel.offset", "tooltip.offset", "0");
        offsetRow.input.active = true;

        return {
            usePreviewBoundsCheckbox: usePreviewBoundsCheckbox,
            deleteGuidesCheckbox: deleteGuidesCheckbox,
            individualCheckbox: individualCheckbox,
            groupCheckbox: groupCheckbox,
            offsetRow: offsetRow.row,
            offsetInput: offsetRow.input
        };
    }

    /**
     * 選択オブジェクトがすべてアートボードの外にあるか判定
     * @param {Document} doc - 対象ドキュメント
     * @returns {boolean} すべて外にあれば true
     */
    function isSelectionOutsideArtboard(doc) {
        var selectedItems = doc.selection;
        var artboardRect = getActiveArtboardRect(doc);
        if (selectedItems.length === 0 || !artboardRect) return false;

        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = selectedItems[i].geometricBounds;
            var isOutside = itemBounds[0] > artboardRect[2] || itemBounds[2] < artboardRect[0] ||
                itemBounds[1] < artboardRect[3] || itemBounds[3] > artboardRect[1];
            if (!isOutside) return false;
        }
        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メインダイアログを構築して表示
     * @returns {void}
     */
    function buildDialog() {
        var doc = app.activeDocument;
        /* 中心線モードはプリセットと共有するため holder で保持 / Center-line mode holder shared with presets */
        var centerLineState = { mode: "" };

        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var columnsGroup = dialog.add("group");
        setupRow(columnsGroup, "fill", COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "top"];

        var leftColumnGroup = columnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];
        leftColumnGroup.spacing = WINDOW_SPACING;

        var targetControls = buildTargetPanel(leftColumnGroup);
        var crossCheckboxes = buildAxisPanel(leftColumnGroup);

        var presetPanel = columnsGroup.add("panel", undefined, getLabel("panel.preset"));
        setupPanel(presetPanel, 6);
        var presetRadios = buildPresetRadios(presetPanel, crossCheckboxes, centerLineState);

        var destinationControls = buildDestinationPanel(dialog);
        var optionControls = buildOptionsPanel(dialog, doc.selection.length);

        /* ボタン行（左：プレビュー／右：Cancel → 作成）/ Button row (left: preview, right: Cancel → Draw) */
        var buttonRow = addButtonRow(dialog);
        var previewCheckbox = addCheckbox(buttonRow.leftGroup, "checkbox.preview", "tooltip.preview", true);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
        var btnDraw = buttonRow.rightGroup.add("button", undefined, getLabel("button.draw"), { name: "ok" });
        var footerControls = { previewCheckbox: previewCheckbox, cancelButton: btnCancel, drawButton: btnDraw };

        /**
         * 既存ガイドの削除チェックボックスのディムを更新する（「_guide」レイヤーに描くときだけ有効）
         * @returns {void}
         */
        function updateDeleteGuidesEnabled() {
            optionControls.deleteGuidesCheckbox.enabled = destinationControls.guideLayerRadio.value;
        }
        destinationControls.selectionLayerRadio.onClick = updateDeleteGuidesEnabled;
        destinationControls.guideLayerRadio.onClick = updateDeleteGuidesEnabled;
        updateDeleteGuidesEnabled();

        /* 選択がすべてアートボード外なら自動的にカンバス基準へ / Prefer canvas when the selection is off-artboard */
        if (isSelectionOutsideArtboard(doc)) {
            targetControls.canvasRadio.value = true;
            targetControls.artboardRadio.value = false;
            targetControls.updateExtensionEnabled();
        }

        /* ===== プレビュー配線 / Preview wiring ===== */

        /**
         * ダイアログの現在値をガイド作成オプションにまとめる
         * @returns {object} ガイド作成オプション（各方向フラグ・プレビュー境界・個別作成・グループ化・描画先・既存削除）
         */
        function readGuideOptions() {
            return {
                left: crossCheckboxes.left.value,
                right: crossCheckboxes.right.value,
                top: crossCheckboxes.top.value,
                bottom: crossCheckboxes.bottom.value,
                center: crossCheckboxes.center.value,
                centerMode: centerLineState.mode,
                usePreviewBounds: optionControls.usePreviewBoundsCheckbox.value,
                individual: optionControls.individualCheckbox.value,
                groupGuides: optionControls.groupCheckbox.value,
                drawOnSelectionLayer: destinationControls.selectionLayerRadio.value,
                deleteExisting: optionControls.deleteGuidesCheckbox.value
            };
        }

        /**
         * 対象・マージン・延長の現在値を描画設定にまとめる
         * @returns {object} { useCanvas: boolean, offsetPt: number, extensionPt: number }
         */
        function readDrawSettings() {
            return {
                useCanvas: targetControls.canvasRadio.value,
                offsetPt: rulerTextToPoints(optionControls.offsetInput.text),
                extensionPt: rulerTextToPoints(targetControls.extensionInput.text)
            };
        }

        /**
         * マージン行のディムを更新する（マージンは辺のガイドにだけ効くので、中心線だけのときはディムする）
         * @returns {void}
         */
        function updateOffsetEnabled() {
            optionControls.offsetRow.enabled = crossCheckboxes.left.value || crossCheckboxes.right.value ||
                crossCheckboxes.top.value || crossCheckboxes.bottom.value;
            redrawSteppersIn(optionControls.offsetRow); /* 自作描画の∧∨を描き直す / redraw the custom-drawn stepper */
        }

        /* 既存ガイドの表示／非表示（showguide はトグルなので状態を自前で追跡）/ Toggle guide visibility (track state) */
        var guidesHidden = false;

        /**
         * 既存ガイドの表示・非表示を切り替える
         * @param {boolean} hide - 非表示にするなら true
         * @returns {void}
         */
        function setGuidesHidden(hide) {
            if (hide === guidesHidden) return;
            /* メニューコマンドが使えない状況でもダイアログ操作を止めない / Keep the dialog usable if the command is unavailable */
            try {
                app.executeMenuCommand("showguide");
            } catch (e) {}
            guidesHidden = hide;
        }

        /**
         * 現在の設定でライブプレビューを描き直す
         * @returns {void}
         */
        function renderPreview() {
            updateOffsetEnabled();
            removePreviewLayer(doc);
            /* プレビュー中は既存ガイドを隠し、仮ガイド（色付き線）だけ見せる / Hide real guides during preview */
            setGuidesHidden(footerControls.previewCheckbox.value);
            if (footerControls.previewCheckbox.value) {
                try {
                    var segments = collectPreviewSegments(readGuideOptions(), readDrawSettings());
                    if (segments.length > 0) {
                        var previewLayer = createPreviewLayer(doc);
                        drawPreviewSegments(previewLayer, segments, makePreviewColor(doc));
                        previewLayer.locked = true;
                    }
                } catch (e) {
                    removePreviewLayer(doc);
                }
            }
            app.redraw();
        }

        /**
         * 既存の onClick を保持したまま、後ろにプレビュー更新を連結する
         * @param {Object} control - 対象のコントロール（null のときは何もしない）
         * @returns {void}
         */
        function chainPreview(control) {
            if (!control) return;
            var previousOnClick = control.onClick;
            control.onClick = function () {
                if (previousOnClick) previousOnClick();
                renderPreview();
            };
        }
        var previewTriggers = [
            targetControls.canvasRadio, targetControls.artboardRadio,
            presetRadios.allOn, presetRadios.edges, presetRadios.topBottom, presetRadios.leftRight,
            presetRadios.topLeft, presetRadios.bottomLeft, presetRadios.topRight, presetRadios.bottomRight,
            presetRadios.centerBoth, presetRadios.centerVertical, presetRadios.centerHorizontal, presetRadios.clear,
            crossCheckboxes.left, crossCheckboxes.top, crossCheckboxes.right, crossCheckboxes.bottom, crossCheckboxes.center,
            optionControls.usePreviewBoundsCheckbox, optionControls.individualCheckbox, footerControls.previewCheckbox
        ];
        for (var i = 0; i < previewTriggers.length; i++) {
            chainPreview(previewTriggers[i]);
        }
        optionControls.offsetInput.onChanging = renderPreview;
        targetControls.extensionInput.onChanging = renderPreview;

        footerControls.cancelButton.onClick = function () {
            dialog.close();
        };

        footerControls.drawButton.onClick = function () {
            removePreviewLayer(doc); /* プレビューを片付けてから本処理 / clean up the preview before committing */
            try {
                createGuides(readGuideOptions(), readDrawSettings());
                dialog.close();
            } catch (e) {
                alert(getLabel("alert.guideError") + "\n" + (e && e.message ? e.message : e) + "\n" + (e && e.stack ? e.stack : ""));
            }
        };

        /* 表示時に初期プレビュー / Initial preview on show */
        dialog.onShow = function () {
            renderPreview();
        };
        /* 閉じる時は仮ガイドを片付け、隠した既存ガイドを再表示 / On close: clear preview and restore guide visibility */
        dialog.onClose = function () {
            removePreviewLayer(doc);
            setGuidesHidden(false);
            app.redraw();
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        dialog.show();
    }

    /* エントリーポイント / Entry point */
    (function main() {
        if (!app.documents.length) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        buildDialog();
    })();

})();
