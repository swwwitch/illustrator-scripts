#target illustrator
#targetengine "AiCreateArtboardGuidesEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボードを基準に、ルーラーガイドの変換・中心ガイド・エッジガイドをダイアログでまとめて作成します。
作成したガイドは「_guide」レイヤーに集約し、ライブプレビューで結果を確認しながら設定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCreateArtboardGuides.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n56d9c936a364

### Overview

Creates artboard-based guides — converted ruler guides, center guides, and edge guides — from a single dialog.
The generated guides are collected into a "_guide" layer, with a live preview of the result.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCreateArtboardGuides.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiCreateArtboardGuides";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCreateArtboardGuides.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCreateArtboardGuides.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n56d9c936a364"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 既定の延長（現在のルーラー単位）/ Default extend amount (in current ruler unit) */
    var DEFAULT_EXTEND = 0;

    /* エッジ描画の既定の延長（mm）/ Default edge extend amount (mm) */
    var DEFAULT_EDGE_EXTEND_MM = 10;

    /* テキストフィールドの文字数幅 / Width of number fields (in characters) */
    var FIELD_CHARACTERS = 4;

    /* エッジ用チェックボックスの幅（px）/ Width of each edge checkbox (px) */
    var EDGE_CHECKBOX_WIDTH = 48;

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
            title: { ja: "アートボードガイドの作成", en: "Create Artboard Guides" }
        },
        field: {
            extend: { ja: "外側に延長", en: "Extend" },
            edgeExtend: { ja: "外側へ延長", en: "Extend Beyond Edge" }
        },
        panel: {
            convert: { ja: "ルーラーガイドを変換", en: "Convert ruler guides" },
            edge: { ja: "中心とエッジにガイドを描画", en: "Center & Edge Guides" }
        },
        edge: {
            top: { ja: "上", en: "Top" },
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" },
            bottom: { ja: "下", en: "Bottom" }
        },
        checkbox: {
            allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
            convertGuides: { ja: "ルーラーガイドを変換", en: "Convert ruler guides" },
            drawEdges: { ja: "エッジ", en: "Draw edge guides" },
            centerVertical: { ja: "中心（垂直）", en: "Draw vertical center guide" },
            centerHorizontal: { ja: "中心（水平）", en: "Draw horizontal center guide" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        tip: {
            extend: {
                ja: "アートボードの端から外側へ延長する長さ（0で端ぴったり）。↑↓で増減、Shiftで±10",
                en: "How far to extend beyond the artboard edge (0 = flush with the edge). Arrow keys to step; Shift ±10"
            },
            edgeExtend: {
                ja: "エッジガイドをアートボードの角から外側へ延長する長さ",
                en: "How far the edge guides extend beyond the artboard corners"
            },
            edgeTop: { ja: "上辺にガイドを作成", en: "Create a guide on the top edge." },
            edgeBottom: { ja: "下辺にガイドを作成", en: "Create a guide on the bottom edge." },
            edgeLeft: { ja: "左辺にガイドを作成", en: "Create a guide on the left edge." },
            edgeRight: { ja: "右辺にガイドを作成", en: "Create a guide on the right edge." },
            target: {
                ja: "アートボード内にある、変換対象のガイドの数",
                en: "Number of guides inside artboards that will be converted"
            },
            allArtboards: {
                ja: "ON：ルーラーガイドが重なるすべてのアートボードを対象にします（OFF：最初の1つだけ）",
                en: "On: target every artboard the ruler guide overlaps (Off: only the first one)"
            },
            drawAllArtboards: {
                ja: "ON：すべてのアートボードに描画（OFF：アクティブなアートボードのみ）",
                en: "On: draw on every artboard (Off: the active artboard only)"
            },
            centerVertical: {
                ja: "アートボードを左右に分ける垂直の中心ガイドを作成",
                en: "Create a vertical guide at the horizontal center"
            },
            centerHorizontal: {
                ja: "アートボードを上下に分ける水平の中心ガイドを作成",
                en: "Create a horizontal guide at the vertical center"
            },
            drawEdges: {
                ja: "OFFにすると、このパネルのエッジ描画設定を無効化します",
                en: "Turn off to disable all edge settings in this panel"
            },
            convertGuides: {
                ja: "OFFにすると、ルーラーガイドの変換を行いません",
                en: "Turn off to skip converting ruler guides"
            }
        },
        hint: {
            noConvertTargets: {
                ja: "変換できるガイドがありません（エッジ描画のみ実行できます）",
                en: "No guides to convert (you can still draw edge guides)"
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
    function labelWithCount(key, count) {
        if (uiLang === "ja") {
            return getLabel(key) + "（" + count + "）";
        }
        return getLabel(key) + " (" + count + ")";
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

    /* mm を現在のルーラー単位の値へ換算（小数第1位で丸め）/ Convert mm to current ruler unit (rounded to 1 decimal) */
    function mmToCurrentUnit(millimeters) {
        var points = millimeters * UNITS[1].pointsPerUnit;
        var value = points / getUnitInfo().pointsPerUnit;
        return Math.round(value * 10) / 10;
    }

    // =========================================
    // UI ヘルパー / UI helpers
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

    /* テキストを数値へ（不正なら 0）/ Parse text to a number (0 when invalid) */
    function parseNumberOrZero(text) {
        var value = parseFloat(text);
        return isNaN(value) ? 0 : value;
    }

    /* 「ラベル＋数値入力＋単位」の行を追加 / Add a labeled number field with a unit suffix */
    function addUnitField(parent, labelKey, defaultText, unitLabel, tooltipKey) {
        var fieldGroup = parent.add("group");
        fieldGroup.orientation = "row";
        fieldGroup.alignChildren = ["left", "center"];

        var fieldLabel = fieldGroup.add("statictext", undefined, labelText(labelKey));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldGroup.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        /* 下限0。増減後は onChanging を呼んでプレビューを更新（プログラム変更では発火しない）
           minimum 0; call onChanging to refresh the preview (programmatic changes do not fire it) */
        var stepperGroup = addStepper(stepperFieldGroup, function () { return inputField; }, {
            min: 0,
            onStep: function () {
                if (typeof inputField.onChanging === "function") inputField.onChanging();
            }
        });
        var inputField = stepperFieldGroup.add("edittext", undefined, defaultText);
        inputField.characters = FIELD_CHARACTERS;
        fieldGroup.add("statictext", undefined, unitLabel);
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(inputField, stepperGroup);
        inputField.fieldRow = fieldGroup; /* 行ごと有効／無効を切り替えるため / for toggling the whole row */

        if (tooltipKey) {
            var tooltip = getLabel(tooltipKey);
            fieldLabel.helpTip = tooltip;
            inputField.helpTip = tooltip;
        }

        return inputField;
    }

    /* 中心とエッジのガイドパネルを構築 / Build the center & edge guides panel */
    function buildEdgePanel(parent, unitLabel) {
        var edgePanel = parent.add("panel", undefined, getLabel("panel.edge"));
        setupPanel(edgePanel, 6);

        /* 中心ガイド（垂直・水平。マスターとは独立）/ Center guides (vertical/horizontal; independent of the edge master) */
        var verticalCheckbox = edgePanel.add("checkbox", undefined, getLabel("checkbox.centerVertical"));
        verticalCheckbox.helpTip = getLabel("tip.centerVertical");
        verticalCheckbox.value = false;

        var horizontalCheckbox = edgePanel.add("checkbox", undefined, getLabel("checkbox.centerHorizontal"));
        horizontalCheckbox.helpTip = getLabel("tip.centerHorizontal");
        horizontalCheckbox.value = false;

        /* エッジ描画のマスタースイッチ（既定OFF。OFFで十字・延長をディム）/ Edge master toggle (default OFF; dims the cross and extend) */
        var drawEdgesCheckbox = edgePanel.add("checkbox", undefined, getLabel("checkbox.drawEdges"));
        drawEdgesCheckbox.helpTip = getLabel("tip.drawEdges");
        drawEdgesCheckbox.value = false;

        /* 上・左右・下の十字配置（各行を中央寄せ）/ Cross layout (each row centered) */
        var topRow = edgePanel.add("group");
        topRow.alignment = "center";
        var topCheckbox = topRow.add("checkbox", undefined, getLabel("edge.top"));

        var middleRow = edgePanel.add("group");
        middleRow.orientation = "row";
        middleRow.alignment = "center";
        middleRow.spacing = 24;
        var leftCheckbox = middleRow.add("checkbox", undefined, getLabel("edge.left"));
        var rightCheckbox = middleRow.add("checkbox", undefined, getLabel("edge.right"));

        var bottomRow = edgePanel.add("group");
        bottomRow.alignment = "center";
        var bottomCheckbox = bottomRow.add("checkbox", undefined, getLabel("edge.bottom"));

        /* 各チェックボックスの幅を少し広げる / Slightly widen each checkbox */
        topCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
        leftCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
        rightCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
        bottomCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;

        /* 各辺の説明 tooltip / Per-edge tooltips */
        topCheckbox.helpTip = getLabel("tip.edgeTop");
        bottomCheckbox.helpTip = getLabel("tip.edgeBottom");
        leftCheckbox.helpTip = getLabel("tip.edgeLeft");
        rightCheckbox.helpTip = getLabel("tip.edgeRight");

        /* 既定はすべて ON / Default all ON */
        topCheckbox.value = leftCheckbox.value = rightCheckbox.value = bottomCheckbox.value = true;

        /* エッジの延長（メインとは別値、既定 10mm）/ Edge extend length (separate value, default 10mm) */
        var edgeExtendInput = addUnitField(edgePanel, "field.edgeExtend", String(mmToCurrentUnit(DEFAULT_EDGE_EXTEND_MM)), unitLabel, "tip.edgeExtend");

        /* 中心・エッジの描画スコープ（すべて / アクティブのみ。マスターとは独立）/ Center & edge drawing scope (independent of masters) */
        var allArtboardsCheckbox = edgePanel.add("checkbox", undefined, getLabel("checkbox.allArtboards"));
        allArtboardsCheckbox.helpTip = getLabel("tip.drawAllArtboards");
        allArtboardsCheckbox.value = false;

        /* 有効/無効の同期 / Sync enabled state */
        var edgeOnlyControls = [topRow, middleRow, bottomRow];
        function syncEdgeEnabled() {
            // 十字（上下左右）はエッジマスターのみで制御 / Cross is controlled by the edge master only
            for (var i = 0; i < edgeOnlyControls.length; i++) {
                edgeOnlyControls[i].enabled = drawEdgesCheckbox.value;
            }
            // 中心(垂直/水平) か エッジ のいずれかが有効か / Whether center or edge is on
            var anyDraw = drawEdgesCheckbox.value || verticalCheckbox.value || horizontalCheckbox.value;
            // 延長・すべてのアートボードは、何か1つでも描画ONなら使える / Extend and scope are usable when anything draws
            edgeExtendInput.fieldRow.enabled = anyDraw;
            redrawSteppersIn(edgeExtendInput.fieldRow); /* ∧∨のディム表示を切り替える / update the stepper dimming */
            allArtboardsCheckbox.enabled = anyDraw;
        }
        drawEdgesCheckbox.onClick = syncEdgeEnabled;
        verticalCheckbox.onClick = syncEdgeEnabled;
        horizontalCheckbox.onClick = syncEdgeEnabled;
        syncEdgeEnabled(); // 初期状態を反映 / apply initial state

        return {
            verticalCheckbox: verticalCheckbox,
            horizontalCheckbox: horizontalCheckbox,
            drawEdgesCheckbox: drawEdgesCheckbox,
            topCheckbox: topCheckbox,
            leftCheckbox: leftCheckbox,
            rightCheckbox: rightCheckbox,
            bottomCheckbox: bottomCheckbox,
            extendInput: edgeExtendInput,
            allArtboardsCheckbox: allArtboardsCheckbox
        };
    }

    /* 変換パネル（タイトルに件数）を構築 / Build the convert panel (count shown in the title) */
    function buildConvertPanel(parent, unitLabel, convertibleCount) {
        var convertPanel = parent.add("panel", undefined, labelWithCount("panel.convert", convertibleCount));
        setupPanel(convertPanel, 6);
        convertPanel.helpTip = getLabel("tip.target");

        /* 変換のマスタースイッチ（OFFで以下をディム）/ Master toggle (OFF dims the rest) */
        var convertGuidesCheckbox = convertPanel.add("checkbox", undefined, getLabel("checkbox.convertGuides"));
        convertGuidesCheckbox.helpTip = getLabel("tip.convertGuides");
        convertGuidesCheckbox.value = true;

        var extendInput = addUnitField(convertPanel, "field.extend", String(DEFAULT_EXTEND), unitLabel, "tip.extend");

        /* 変換のスコープ（重なるすべてのアートボード。一番下）/ Conversion scope (every overlapping artboard; bottom) */
        var allArtboardsCheckbox = convertPanel.add("checkbox", undefined, getLabel("checkbox.allArtboards"));
        allArtboardsCheckbox.helpTip = getLabel("tip.allArtboards");
        allArtboardsCheckbox.value = false;

        /* マスターOFF時にディムする要素 / Elements dimmed when the master toggle is OFF */
        var dimmableControls = [extendInput.fieldRow, allArtboardsCheckbox];
        convertGuidesCheckbox.onClick = function () {
            for (var i = 0; i < dimmableControls.length; i++) {
                dimmableControls[i].enabled = convertGuidesCheckbox.value;
            }
            redrawSteppersIn(extendInput.fieldRow); /* ∧∨のディム表示を切り替える / update the stepper dimming */
        };

        /* 変換対象が無ければ説明を出してマスターごと無効化 / No targets: show a note and disable the whole section */
        if (convertibleCount === 0) {
            convertGuidesCheckbox.value = false;
            convertGuidesCheckbox.enabled = false;
            convertPanel.add("statictext", undefined, getLabel("hint.noConvertTargets"));
        }
        convertGuidesCheckbox.onClick(); // 初期状態を反映 / apply initial state

        return {
            convertGuidesCheckbox: convertGuidesCheckbox,
            allArtboardsCheckbox: allArtboardsCheckbox,
            extendInput: extendInput
        };
    }

    /* 延長とエッジ描画を入力するダイアログを表示（ライブプレビュー付き）/ Show the dialog (with live preview) */
    function showExtendDialog(doc, convertTargets) {
        var unitLabel = getUnitInfo().label;
        var convertibleCount = convertTargets.length;

        /* ダイアログ本体 / Dialog window */
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 変換パネル（件数はタイトルに表示）/ Convert panel (count in the title) */
        var convertControls = buildConvertPanel(dialog, unitLabel, convertibleCount);
        var extendInput = convertControls.extendInput;

        /* 中心とエッジのガイドパネル / Center & edge guides panel */
        var edgeControls = buildEdgePanel(dialog, unitLabel);

        /* プレビュー切り替え（既定ON）/ Preview toggle (default ON) */
        var previewCheckbox = dialog.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = true;

        /* ボタン（左右中央・Mac 順：Cancel → OK）/ Buttons (centered, Mac order: Cancel → OK) */
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /* 現在のUIから設定を読み取る / Read options from the current UI */
        function readOptions() {
            return {
                convertGuides: convertControls.convertGuidesCheckbox.value,
                extend: parseNumberOrZero(extendInput.text),
                allArtboards: convertControls.allArtboardsCheckbox.value,
                center: {
                    vertical: edgeControls.verticalCheckbox.value,
                    horizontal: edgeControls.horizontalCheckbox.value
                },
                drawAllArtboards: edgeControls.allArtboardsCheckbox.value,
                drawEdges: edgeControls.drawEdgesCheckbox.value,
                edgeExtend: parseNumberOrZero(edgeControls.extendInput.text),
                edges: {
                    top: edgeControls.topCheckbox.value,
                    left: edgeControls.leftCheckbox.value,
                    right: edgeControls.rightCheckbox.value,
                    bottom: edgeControls.bottomCheckbox.value
                }
            };
        }

        /* プレビュー状態 / Preview state */
        var previewColor = makePreviewColor(doc);
        var hiddenGuides = [];   // 一時的に隠した元ガイド / originals temporarily hidden

        /* 隠した元ガイドを元に戻す / Restore originals that were hidden */
        function restoreHiddenGuides() {
            for (var k = 0; k < hiddenGuides.length; k++) {
                try { hiddenGuides[k].hidden = false; } catch (e) {}
            }
            hiddenGuides = [];
        }

        /* プレビューを消去（専用レイヤーごと削除）（app.undo は使わない）/ Clear preview (drop the whole layer; no app.undo) */
        function clearPreview() {
            removePreviewLayer(doc);
            restoreHiddenGuides();
        }

        /* プレビューを再描画 / Re-render the preview */
        function renderPreview() {
            clearPreview();
            if (previewCheckbox.value) {
                try {
                    var layer = createPreviewLayer(doc);
                    var opt = readOptions();
                    var pointsPerUnit = getUnitInfo().pointsPerUnit;
                    var extendPoints = opt.extend * pointsPerUnit;
                    var edgeExtendPoints = opt.edgeExtend * pointsPerUnit;

                    /* 変換：元を隠して新規（色付き線）を描く / Conversion: hide originals, draw colored lines */
                    if (opt.convertGuides) {
                        for (var i = 0; i < convertTargets.length; i++) {
                            var target = convertTargets[i];
                            try { target.guidePath.hidden = true; hiddenGuides.push(target.guidePath); } catch (e) {}
                            addPreviewSegments(layer, convertTargetSegments(target, opt.allArtboards, extendPoints), previewColor);
                        }
                    }

                    /* 中心・エッジ / Center & edge */
                    var rects = getTargetArtboardRects(doc, opt.drawAllArtboards);
                    var anyEdge = opt.drawEdges &&
                        (opt.edges.top || opt.edges.bottom || opt.edges.left || opt.edges.right);
                    for (var r = 0; r < rects.length; r++) {
                        if (anyEdge) {
                            addPreviewSegments(layer, edgeSegments(rects[r], opt.edges, edgeExtendPoints), previewColor);
                        }
                        if (opt.center.vertical || opt.center.horizontal) {
                            addPreviewSegments(layer, centerSegments(rects[r], opt.center, edgeExtendPoints), previewColor);
                        }
                    }

                    layer.locked = true; // 選択不可に / make non-selectable
                } catch (e) {
                    clearPreview();
                }
            }
            app.redraw();
        }

        /* 既存 onClick を保持しつつプレビュー更新を連結（addEventListener併用は不安定なため）/ Chain renderPreview after any existing onClick */
        function chainPreview(control) {
            var previousOnClick = control.onClick;
            control.onClick = function () {
                if (previousOnClick) previousOnClick();
                renderPreview();
            };
        }

        /* 変更を監視してプレビュー更新 / Update preview on any change */
        var previewTriggers = [
            convertControls.convertGuidesCheckbox,
            convertControls.allArtboardsCheckbox,
            edgeControls.verticalCheckbox,
            edgeControls.horizontalCheckbox,
            edgeControls.drawEdgesCheckbox,
            edgeControls.topCheckbox,
            edgeControls.leftCheckbox,
            edgeControls.rightCheckbox,
            edgeControls.bottomCheckbox,
            edgeControls.allArtboardsCheckbox,
            previewCheckbox
        ];
        for (var p = 0; p < previewTriggers.length; p++) {
            chainPreview(previewTriggers[p]);
        }
        extendInput.onChanging = renderPreview;
        edgeControls.extendInput.onChanging = renderPreview;

        extendInput.active = true;

        var dialogResult = null;
        btnOK.onClick = function () {
            dialogResult = readOptions();
            dialog.close();
        };
        btnCancel.onClick = function () {
            dialogResult = null;
            dialog.close();
        };

        /* 閉じる時は必ずプレビューを後始末（本処理はクリーンな状態で実行）/ Always clean up on close */
        dialog.onClose = function () {
            clearPreview();
            app.redraw();
        };
        /* 表示時に初期プレビュー / Initial preview on show */
        dialog.onShow = function () {
            renderPreview();
        };

        prepareDialogWindow(dialog, SCRIPT_NAME);
        dialog.show();
        return dialogResult;
    }

    // =========================================
    // ガイド処理 / Guide processing
    // =========================================

    /* 描画対象のアートボード矩形を取得（ON：全部、OFF：アクティブのみ）/ Get target artboard rects (On: all, Off: active only) */
    function getTargetArtboardRects(doc, allArtboards) {
        var rects = [];
        if (allArtboards) {
            for (var i = 0; i < doc.artboards.length; i++) {
                rects.push(doc.artboards[i].artboardRect);
            }
        } else {
            rects.push(doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect);
        }
        return rects;
    }

    /* 確定ガイドを作成するレイヤー名 / Layer name where committed guides are created */
    var GUIDE_LAYER_NAME = "_guide";

    /* 「_guide」レイヤーを取得（無ければ作成。ロック/非表示は解除）/ Get the "_guide" layer (create if missing; unlock/show) */
    function getGuideLayer(doc) {
        var layer;
        try {
            layer = doc.layers.getByName(GUIDE_LAYER_NAME);
        } catch (e) {
            layer = doc.layers.add();
            layer.name = GUIDE_LAYER_NAME;
        }
        layer.locked = false;
        layer.visible = true;
        return layer;
    }

    /* 直線のガイドを 1 本作成 / Create a single straight guide */
    function addGuideLine(doc, startPoint, endPoint) {
        var guidePath = getGuideLayer(doc).pathItems.add();
        guidePath.setEntirePath([startPoint, endPoint]);
        guidePath.stroked = false;
        guidePath.filled = false;
        guidePath.guides = true;
        return guidePath;
    }

    /* 変換ターゲットの線分配列を生成 / Build line segments for a convert target */
    function convertTargetSegments(target, allArtboards, extendPoints) {
        var segments = [];
        var artboardRects = allArtboards ? target.artboardRects : [target.artboardRects[0]];

        for (var j = 0; j < artboardRects.length; j++) {
            var rect = artboardRects[j]; // [left, top, right, bottom]

            if (target.isVertical) {
                // 上方向は +、下方向は -（Illustrator の Y は上が大きい）/ Y grows upward in Illustrator
                segments.push([[target.position, rect[1] + extendPoints], [target.position, rect[3] - extendPoints]]);
            } else {
                // 左方向は -、右方向は + / Extend left and right
                segments.push([[rect[0] - extendPoints, target.position], [rect[2] + extendPoints, target.position]]);
            }
        }
        return segments;
    }

    /* エッジの線分配列を生成 / Build edge segments for an artboard rect */
    function edgeSegments(rect, edges, edgeExtendPoints) {
        var segments = [];
        var abLeft = rect[0], abTop = rect[1], abRight = rect[2], abBottom = rect[3];

        if (edges.top)    segments.push([[abLeft - edgeExtendPoints, abTop],    [abRight + edgeExtendPoints, abTop]]);
        if (edges.bottom) segments.push([[abLeft - edgeExtendPoints, abBottom], [abRight + edgeExtendPoints, abBottom]]);
        if (edges.left)   segments.push([[abLeft, abTop + edgeExtendPoints],    [abLeft, abBottom - edgeExtendPoints]]);
        if (edges.right)  segments.push([[abRight, abTop + edgeExtendPoints],   [abRight, abBottom - edgeExtendPoints]]);
        return segments;
    }

    /* 中心の線分配列を生成 / Build center segments for an artboard rect */
    function centerSegments(rect, center, edgeExtendPoints) {
        var segments = [];
        var centerX = (rect[0] + rect[2]) / 2;
        var centerY = (rect[1] + rect[3]) / 2;

        if (center.vertical) {
            segments.push([[centerX, rect[1] + edgeExtendPoints], [centerX, rect[3] - edgeExtendPoints]]);
        }
        if (center.horizontal) {
            segments.push([[rect[0] - edgeExtendPoints, centerY], [rect[2] + edgeExtendPoints, centerY]]);
        }
        return segments;
    }

    /* 線分配列からガイドを作成（collector があれば作成物を push）/ Create guides from segments (collect if given) */
    function addGuideSegments(doc, segments, collector) {
        for (var i = 0; i < segments.length; i++) {
            var guidePath = addGuideLine(doc, segments[i][0], segments[i][1]);
            if (collector) collector.push(guidePath);
        }
    }

    /* プレビュー用レイヤー名 / Preview layer name */
    var PREVIEW_LAYER_NAME = "__ArtboardGuidesPreview__";

    /* プレビュー用の色（ドキュメントのカラースペースに合わせる）/ Preview color (matches doc color space) */
    function makePreviewColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0; cmyk.magenta = 90; cmyk.yellow = 0; cmyk.black = 0;
            return cmyk;
        }
        var rgb = new RGBColor();
        rgb.red = 255; rgb.green = 0; rgb.blue = 255;
        return rgb;
    }

    /* プレビュー用レイヤーを削除 / Remove the preview layer if present */
    function removePreviewLayer(doc) {
        try {
            var layer = doc.layers.getByName(PREVIEW_LAYER_NAME);
            layer.locked = false;
            layer.visible = true;
            layer.remove();
        } catch (e) {}
    }

    /* プレビュー用レイヤーを用意（既存は作り直し）/ Create a fresh preview layer */
    function createPreviewLayer(doc) {
        removePreviewLayer(doc);
        var layer = doc.layers.add();
        layer.name = PREVIEW_LAYER_NAME;
        return layer;
    }

    /* 線分配列を色付きプレビュー線としてレイヤーへ描画 / Draw colored preview lines into a layer */
    function addPreviewSegments(layer, segments, color) {
        for (var i = 0; i < segments.length; i++) {
            var path = layer.pathItems.add();
            path.setEntirePath([segments[i][0], segments[i][1]]);
            path.filled = false;
            path.stroked = true;
            path.strokeColor = color;
            path.strokeWidth = 1;
        }
    }

    /* 検出済みガイドをアートボード基準に引き直す / Redraw detected guides to their artboard(s) */
    function convertGuidesToArtboards(doc, convertTargets, allArtboards, extendPoints) {
        for (var i = 0; i < convertTargets.length; i++) {
            var convertTarget = convertTargets[i];
            var segments = convertTargetSegments(convertTarget, allArtboards, extendPoints);

            /* 元ガイドは1回だけ削除（ロック中でも消せるよう先にアンロック）/ Remove the original guide once (unlock first so locked guides can be removed) */
            try { convertTarget.guidePath.locked = false; } catch (e) {}
            convertTarget.guidePath.remove();
            addGuideSegments(doc, segments, null);
        }
    }

    /* アートボードのエッジにガイドを描画 / Draw guides on the artboard edges */
    function drawEdgeGuides(doc, artboardRects, edges, edgeExtendPoints) {
        for (var i = 0; i < artboardRects.length; i++) {
            addGuideSegments(doc, edgeSegments(artboardRects[i], edges, edgeExtendPoints), null);
        }
    }

    /* アートボードの中心にガイドを描画 / Draw guides at the artboard centers */
    function drawCenterGuides(doc, artboardRects, center, edgeExtendPoints) {
        for (var i = 0; i < artboardRects.length; i++) {
            addGuideSegments(doc, centerSegments(artboardRects[i], center, edgeExtendPoints), null);
        }
    }

    /* ガイドが重なるアートボードの矩形をすべて返す / Return the rects of every artboard the guide overlaps */
    function findOverlappingArtboards(doc, isVertical, guideLeft, guideTop, guideRight, guideBottom) {
        var overlappingRects = [];

        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]

            if (isVertical) {
                if (guideLeft >= artboardRect[0] && guideLeft <= artboardRect[2] &&
                    Math.max(guideTop, guideBottom) >= artboardRect[3] &&
                    Math.min(guideTop, guideBottom) <= artboardRect[1]) {
                    overlappingRects.push(artboardRect);
                }
            } else {
                if (guideTop <= artboardRect[1] && guideTop >= artboardRect[3] &&
                    Math.max(guideLeft, guideRight) >= artboardRect[0] &&
                    Math.min(guideLeft, guideRight) <= artboardRect[2]) {
                    overlappingRects.push(artboardRect);
                }
            }
        }

        return overlappingRects;
    }

    /* アートボードに重なる直線ガイドを変換対象として収集 / Collect straight guides that overlap an artboard */
    function collectConvertTargets(doc, guidePaths) {
        var convertTargets = [];

        for (var i = 0; i < guidePaths.length; i++) {
            var guidePath = guidePaths[i];
            // 「ガイドをロック」ON だとルーラーガイドは locked=true になるが、対象から外さない（削除直前に個別アンロック）/ "Lock Guides" sets locked=true on ruler guides; keep them (unlocked right before removal)
            if (guidePath.hidden) continue;

            var guideBounds = guidePath.geometricBounds; // [left, top, right, bottom]
            var guideLeft = guideBounds[0];
            var guideTop = guideBounds[1];
            var guideRight = guideBounds[2];
            var guideBottom = guideBounds[3];

            var isVertical = Math.abs(guideLeft - guideRight) < 0.01;
            var isHorizontal = Math.abs(guideTop - guideBottom) < 0.01;
            if (!isVertical && !isHorizontal) continue;

            var overlappingArtboardRects = findOverlappingArtboards(doc, isVertical, guideLeft, guideTop, guideRight, guideBottom);
            if (overlappingArtboardRects.length === 0) continue;

            convertTargets.push({
                guidePath: guidePath,
                artboardRects: overlappingArtboardRects,
                isVertical: isVertical,
                // 縦ガイドは x（左端）、横ガイドは y（上端）/ x for vertical, y for horizontal
                position: isVertical ? guideLeft : guideTop
            });
        }

        return convertTargets;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    (function () {

        /* ドキュメントの有無を確認 / Check that a document is open */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var guidePaths = [];

        /* 既存のガイドを収集 / Collect existing guides */
        for (var i = 0; i < doc.pathItems.length; i++) {
            if (doc.pathItems[i].guides) {
                guidePaths.push(doc.pathItems[i]);
            }
        }

        /* 変換対象（アートボード内の直線ガイド）を事前に検出 / Pre-detect convertible guides */
        /* ガイドが無くてもエッジ描画だけ実行できるよう、ここでは終了しない / Do not exit when empty: edge-only runs are allowed */
        var convertTargets = collectConvertTargets(doc, guidePaths);

        /* 延長・エッジ描画をダイアログで取得（プレビュー付き）/ Get options via the dialog (with preview) */
        var options = showExtendDialog(doc, convertTargets);
        if (options === null) {
            return; // キャンセル / Cancelled
        }

        var pointsPerUnit = getUnitInfo().pointsPerUnit;
        var extendPoints = options.extend * pointsPerUnit;
        var edgeExtendPoints = options.edgeExtend * pointsPerUnit;

        /* 変換マスターOFFなら対象を空にしてスキップ / Empty the list to skip conversion when the master is OFF */
        if (!options.convertGuides) {
            convertTargets = [];
        }

        /* ルーラーガイドをアートボード基準に引き直す / Redraw ruler guides to their artboard(s) */
        convertGuidesToArtboards(doc, convertTargets, options.allArtboards, extendPoints);

        /* 中心・エッジの描画対象アートボード（OFFはアクティブのみ）/ Target artboards for center/edge drawing (active only when OFF) */
        var targetArtboardRects = getTargetArtboardRects(doc, options.drawAllArtboards);

        /* アートボードのエッジにガイドを描画（マスターON時のみ）/ Draw edge guides (only when the master toggle is ON) */
        if (options.drawEdges &&
            (options.edges.top || options.edges.bottom || options.edges.left || options.edges.right)) {
            drawEdgeGuides(doc, targetArtboardRects, options.edges, edgeExtendPoints);
        }

        /* アートボードの中心にガイドを描画 / Draw center guides */
        if (options.center.vertical || options.center.horizontal) {
            drawCenterGuides(doc, targetArtboardRects, options.center, edgeExtendPoints);
        }
    })();

})();
