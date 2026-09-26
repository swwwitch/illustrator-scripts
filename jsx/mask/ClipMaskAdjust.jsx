#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップグループ（クリッピングマスク）の「マスクパス」と「内容」を調整します。
ダイアログの操作はオートプレビューで即時反映されます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipMaskAdjust.md

### Overview

Adjusts the mask path and the contents of a clipping group.
Every change in the dialog is reflected immediately as an auto-preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipMaskAdjust.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClipMaskAdjust";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v3.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipMaskAdjust.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipMaskAdjust.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_SPACING = 15;               /* ダイアログ内の要素間隔 / Spacing inside the dialog */
    var COLUMN_SPACING = 10;               /* カラムの間隔 / Spacing between columns */
    var PANEL_MARGINS  = [15, 20, 15, 10]; /* パネル余白 [左,上,右,下] / Panel margins */
    var PANEL_SPACING  = 8;                /* パネル内の要素間隔 / Spacing inside panels */
    var ROW_SPACING    = 5;                /* 行内の要素間隔 / Spacing inside rows */
    var ANCHOR_CELL_SIZE = 15;             /* 基準点ラジオの大きさ / Size of the anchor radios */

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} spacing - 要素間隔
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = spacing;
    }

    /**
     * 縦並びのカラムを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加したカラム
     */
    function addColumn(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = COLUMN_SPACING;
        return columnGroup;
    }

    /**
     * 横並びの行を追加する
     * @param {Object} parentContainer - 追加先のパネルやグループ
     * @returns {Group} 追加した行
     */
    function addRow(parentContainer) {
        var rowGroup = parentContainer.add("group");
        rowGroup.orientation = "row";
        rowGroup.spacing = ROW_SPACING;
        return rowGroup;
    }

    /**
     * 左に∧∨を付けた数値入力欄を追加する（∧∨と入力欄は隙間0で突き合わせる）。
     * ↑↓キーも∧∨と同じ処理で増減し、増減後は入力欄の onChanging（プレビュー更新など）を呼ぶ
     * @param {Group} parentRow - 追加先の行
     * @param {string} defaultText - 初期値
     * @param {number} characters - 入力欄の文字数幅
     * @param {Object} stepOptions - min など（addStepper() の stepOptions）
     * @returns {EditText} 追加した入力欄
     */
    function addStepperInput(parentRow, defaultText, characters, stepOptions) {
        var stepperFieldGroup = parentRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        stepOptions.onStep = function (numberInput) {
            /* プログラム変更では onChanging が発火しないため明示的に呼ぶ / programmatic changes do not fire onChanging */
            if (typeof numberInput.onChanging === "function") numberInput.onChanging();
        };
        var numberInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperFieldGroup.add("edittext", undefined, defaultText);
        numberInput.characters = characters;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
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
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkStepperUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var STEPPER_UI_DARK           = isDarkStepperUI();
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
        if (direction > 0) return Math.floor(value / multiple) * multiple + multiple;
        return Math.ceil(value / multiple) * multiple - multiple;
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

    /**
     * 定規の単位の値を pt にする
     * @param {number} unitValue - 定規の単位での値
     * @returns {number} pt の値
     */
    function unitValueToPt(unitValue) {
        return unitValue * getUnitInfo().pointsPerUnit;
    }

    /**
     * pt の値を定規の単位にする
     * @param {number} ptValue - pt の値
     * @returns {number} 定規の単位での値
     */
    function ptToUnitValue(ptValue) {
        return ptValue / getUnitInfo().pointsPerUnit;
    }

    // =========================================
    // 一時アクション（アピアランスを消去）/ Temporary action (Clear Appearance)
    // =========================================

    /* executeMenuCommand("clearAppearance") は環境差が出ることがあるため、アクションを一時的に読み込んで実行する
       Clear Appearance via a temporary action; the menu command behaves differently across environments */
    var CLEAR_APPEARANCE_SET_NAME = "__sttk3_appearance__";
    var CLEAR_APPEARANCE_ACTION_NAME = "clear";
    var CLEAR_APPEARANCE_ACTION_CODE = [
        "/version 3",
        "/name [ 20",
        "    5f5f7374746b335f617070656172616e63655f5f",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        "    /name [ 5",
        "        636c656172",
        "    ]",
        "    /keyIndex 0",
        "    /colorIndex 0",
        "    /isOpen 0",
        "    /eventCount 1",
        "    /event-1 {",
        "        /useRulersIn1stQuadrant 0",
        "        /internalName (ai_plugin_appearance)",
        "        /localizedName [ 18",
        "            e382a2e38394e382a2e383a9e383b3e382b9",
        "        ]",
        "        /isOpen 0",
        "        /isOn 1",
        "        /hasDialog 0",
        "        /parameterCount 1",
        "        /parameter-1 {",
        "            /key 1835363957",
        "            /showInPalette 4294967295",
        "            /type (enumerated)",
        "            /name [ 27",
        "                e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb",
        "            ]",
        "            /value 6",
        "        }",
        "    }",
        "}"
    ].join("\n");

    /**
     * 選択中のオブジェクトに［アピアランスを消去］のアクションを実行する（失敗しても続行）
     * @returns {void}
     */
    function runClearAppearanceAction() {
        var actionFile = new File(Folder.temp + "/tempActionSet.aia");
        try {
            actionFile.open("w");
            actionFile.write(CLEAR_APPEARANCE_ACTION_CODE);
        } catch (e) {
            alert(e);
            return;
        } finally {
            actionFile.close();
        }

        /* 前回の残りがあれば外す（無ければ例外）/ Unload a leftover set; throws when none */
        try { app.unloadAction(CLEAR_APPEARANCE_SET_NAME, ""); } catch (e) { }

        /* 読み込んだ時点でパース済みなので、一時ファイルはすぐ消す / The file is parsed on load, so remove it right away */
        try {
            app.loadAction(actionFile);
        } catch (e) {
            actionFile.remove();
            return;
        }
        actionFile.remove();

        try {
            app.doScript(CLEAR_APPEARANCE_ACTION_NAME, CLEAR_APPEARANCE_SET_NAME, false);
        } catch (e) {
            /* 失敗してもスクリプトは続行 / Keep going even if the action fails */
        } finally {
            try { app.unloadAction(CLEAR_APPEARANCE_SET_NAME, ""); } catch (e) { }
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "クリップグループの調整", en: "Adjust Clip Group" }
        },
        panel: {
            anchor: { ja: "基準点", en: "Anchor" },
            nudge: { ja: "微調整", en: "Nudge" },
            fitScale: { ja: "フィットとスケール", en: "Fit & Scale" },
            maskPath: { ja: "マスクパス", en: "Mask Path" },
            roundCorners: { ja: "角丸", en: "Round Corners" }
        },
        radio: {
            cover: { ja: "縦横比を保持して切り取り", en: "Proportions (Fill)" },
            contain: { ja: "縦横比を保持して縮小", en: "Proportions (Fit)" },
            keepSize: { ja: "サイズ保持", en: "Keep Size" },
            manualScale: { ja: "スケールを指定", en: "Set Scale" },
            maskUnchanged: { ja: "そのまま", en: "Unchanged" },
            fitToContent: { ja: "内容に合わせる", en: "Fit to Content" },
            square: { ja: "正方形に", en: "Square" }
        },
        checkbox: {
            roundCorners: { ja: "角丸", en: "Round Corners" },
            circle: { ja: "正円", en: "Circle" }
        },
        fieldLabel: {
            x: { ja: "X", en: "X" },
            y: { ja: "Y", en: "Y" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
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
            anchor: {
                ja: "サイズを変えるときに動かさない位置です。3×3のマスで選びます。Q/W/E・A/S/D・Z/X/C キーでも選べます。",
                en: "The point that stays fixed when the size changes. Pick one of the 3×3 cells, or press Q/W/E, A/S/D, Z/X/C."
            },
            nudge: {
                ja: "この値だけ位置をずらします。↑↓キーで増減できます。",
                en: "Offsets the position by this amount. Use the Up/Down arrow keys to change the value."
            },
            cover: {
                ja: "縦横比を保ったまま、マスクを埋めるように内容を拡大します。はみ出た部分は切り取られます。",
                en: "Scales the contents proportionally to fill the mask. Anything outside the mask is cropped."
            },
            contain: {
                ja: "縦横比を保ったまま、内容全体がマスクに収まるように縮小します。",
                en: "Scales the contents proportionally so they fit entirely inside the mask."
            },
            keepSize: { ja: "内容の大きさは変えません。", en: "Leaves the size of the contents unchanged." },
            manualScale: { ja: "倍率を数値で指定します。", en: "Sets the scale as a percentage." },
            maskUnchanged: { ja: "マスクパスの形はそのままにします。", en: "Leaves the shape of the mask path unchanged." },
            fitToContent: { ja: "マスクパスを内容の外接範囲に合わせます。", en: "Fits the mask path to the bounds of the contents." },
            square: { ja: "マスクパスを正方形にします。", en: "Makes the mask path a square." },
            roundCorners: {
                ja: "クリップグループに［角を丸くする］効果を適用します。右の欄で半径を指定します。",
                en: "Applies the Round Corners effect to the clip group. Set the radius in the field on the right."
            },
            circle: {
                ja: "［正方形に］のときに使えます。角丸の半径を短辺の半分にして正円にします。",
                en: "Available with Square. Sets the corner radius to half the short side to make a circle."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectClipGroup: { ja: "クリップグループを選択してください。", en: "Select a clip group." }
        }
    };

    /**
     * ラベル（ja/en のリーフ）を現在の言語に解決する
     * @param {Object} labelSet - LABELS のリーフ（{ ja, en }）
     * @returns {string} 現在の言語の文字列（無ければ空文字）
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - LABELS のリーフ（{ ja, en }）
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // クリップグループ / Clip groups
    // =========================================

    /**
     * クリップグループ（clipped な GroupItem）かを判定する
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} クリップグループなら true
     */
    function isClipGroup(item) {
        return !!item && item.typename === "GroupItem" && item.clipped;
    }

    /**
     * 配列からクリップグループだけを抜き出す
     * @param {PageItem[]} items - オブジェクトの配列
     * @returns {GroupItem[]} クリップグループ
     */
    function filterClipGroups(items) {
        var clipGroups = [];
        for (var i = 0; i < items.length; i++) {
            if (isClipGroup(items[i])) clipGroups.push(items[i]);
        }
        return clipGroups;
    }

    /**
     * クリップグループをマスクパスと内容に分ける
     * @param {GroupItem} clipGroup - 対象のクリップグループ
     * @returns {{clipPath: PageItem|null, contents: PageItem[]}} マスクパスと内容
     */
    function splitClipGroup(clipGroup) {
        var clipPath = null;
        var contents = [];
        for (var i = 0; i < clipGroup.pageItems.length; i++) {
            var childItem = clipGroup.pageItems[i];
            if (childItem.clipping) {
                if (!clipPath) clipPath = childItem;
            } else {
                contents.push(childItem);
            }
        }
        return { clipPath: clipPath, contents: contents };
    }

    /**
     * 選択を指定したオブジェクトに置き換える
     * @param {PageItem[]} items - 選択するオブジェクト
     * @returns {void}
     */
    function setSelection(items) {
        app.executeMenuCommand("deselectall");
        for (var i = 0; i < items.length; i++) {
            /* ロック・非表示のオブジェクトは選択できず例外になる / Locked or hidden items throw */
            try { items[i].selected = true; } catch (e) { }
        }
    }

    /**
     * クリップグループのアピアランスを1つずつ消去し、選択を元に戻す
     * @param {GroupItem[]} clipGroups - 対象のクリップグループ
     * @param {PageItem[]} selectionToRestore - 最後に選択し直すオブジェクト
     * @returns {void}
     */
    function clearAppearanceForClipGroups(clipGroups, selectionToRestore) {
        for (var i = 0; i < clipGroups.length; i++) {
            setSelection([clipGroups[i]]);
            runClearAppearanceAction();
        }
        setSelection(selectionToRestore);
    }

    /**
     * 角丸の初期値（定規の単位）を返す：（マスクパスの幅＋高さ）÷25 を切り上げ
     * @param {GroupItem} clipGroup - 対象のクリップグループ
     * @returns {number|null} 初期値（マスクパスが無ければ null）
     */
    function getDefaultRoundRadius(clipGroup) {
        var clipPath = splitClipGroup(clipGroup).clipPath;
        if (!clipPath) return null;
        var maskBounds = clipPath.geometricBounds; /* [L, T, R, B] */
        var radiusPt = ((maskBounds[2] - maskBounds[0]) + (maskBounds[1] - maskBounds[3])) / 25;
        return Math.ceil(ptToUnitValue(radiusPt));
    }

    /**
     * 最初の内容の現在のスケール（％）を返す
     * @param {GroupItem} clipGroup - 対象のクリップグループ
     * @returns {string|number} 小数2桁の文字列（内容が無ければ 100）
     */
    function getContentScalePercent(clipGroup) {
        var contents = splitClipGroup(clipGroup).contents;
        if (contents.length === 0) return 100;
        var contentMatrix = contents[0].matrix;
        var scaleX = Math.sqrt(contentMatrix.mValueA * contentMatrix.mValueA + contentMatrix.mValueB * contentMatrix.mValueB);
        return (scaleX * 100).toFixed(2);
    }

    /**
     * 複数オブジェクトの geometricBounds をまとめた範囲を返す
     * @param {PageItem[]} items - 対象のオブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]（空なら null）
     */
    function getCombinedBounds(items) {
        if (items.length === 0) return null;
        var firstBounds = items[0].geometricBounds;
        var left = firstBounds[0], top = firstBounds[1], right = firstBounds[2], bottom = firstBounds[3];
        for (var i = 1; i < items.length; i++) {
            var itemBounds = items[i].geometricBounds;
            if (itemBounds[0] < left) left = itemBounds[0];
            if (itemBounds[1] > top) top = itemBounds[1];
            if (itemBounds[2] > right) right = itemBounds[2];
            if (itemBounds[3] < bottom) bottom = itemBounds[3];
        }
        return [left, top, right, bottom];
    }

    // =========================================
    // 角丸効果 / Round Corners effect
    // =========================================

    /* 名前に「Round Corners」を含む LiveEffect（名前の揺れに対応）/ Any LiveEffect whose name contains "Round Corners" */
    var ROUND_CORNERS_EFFECT_SOURCE = "<LiveEffect\\b[^>]*\\bname=['\"][^'\"]*Round Corners[^'\"]*['\"][^>]*>[\\s\\S]*?<\\/LiveEffect>";

    /**
     * 角丸の LiveEffect XML を作る
     * @param {number} radiusPt - 半径（pt）
     * @returns {string} LiveEffect の XML
     */
    function createRoundCornersEffectXML(radiusPt) {
        return '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>';
    }

    /**
     * 効果の文字列から角丸だけを取り除く（効果の重複を防ぐ）
     * @param {string} effectText - appliedEffect の文字列
     * @returns {string} 角丸を除いた文字列
     */
    function stripRoundCornersEffect(effectText) {
        if (!effectText) return "";
        return effectText.replace(new RegExp(ROUND_CORNERS_EFFECT_SOURCE, "g"), "");
    }

    /**
     * オブジェクトから角丸の効果を取り除く（他の効果は残す）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {void}
     */
    function removeRoundCornersEffect(targetItem) {
        /* appliedEffect の読み書きは環境により例外になることがある / appliedEffect access may throw */
        try {
            targetItem.appliedEffect = stripRoundCornersEffect(targetItem.appliedEffect || "");
        } catch (e) { }
    }

    /**
     * 既存の角丸を除いてから角丸の効果を適用する
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {number} radiusPt - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(targetItem, radiusPt) {
        try {
            /* appliedEffect への直接連結は無視されることがあるため、除去してから applyEffect で足す
               Appending to appliedEffect is sometimes ignored, so strip first and add via applyEffect */
            targetItem.appliedEffect = stripRoundCornersEffect(targetItem.appliedEffect || "");
            targetItem.applyEffect(createRoundCornersEffectXML(radiusPt));
        } catch (e) {
            /* 除去に失敗しても角丸だけは付ける / Still try to add the effect */
            try { targetItem.applyEffect(createRoundCornersEffectXML(radiusPt)); } catch (e2) { }
        }
    }

    /**
     * 既存の角丸の効果の半径だけを書き換える（効果を新しく足さない）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {number} radiusPt - 半径（pt）
     * @returns {boolean} 書き換えられたら true（角丸が無い・形式が違うときは false）
     */
    function updateRoundCornersRadiusOnly(targetItem, radiusPt) {
        try {
            var effectText = targetItem.appliedEffect || "";
            var effectMatch = effectText.match(new RegExp(ROUND_CORNERS_EFFECT_SOURCE));
            if (!effectMatch) return false;

            /* ブロック内の最初の "R radius <数値>" だけを差し替える / Replace the first radius only */
            var effectBlock = effectMatch[0];
            var radiusPattern = /(R\s+radius\s+)(-?\d+(?:\.\d+)?)/;
            if (!radiusPattern.test(effectBlock)) return false;
            var newBlock = effectBlock.replace(radiusPattern, function (whole, radiusPrefix) {
                return radiusPrefix + String(radiusPt);
            });
            targetItem.appliedEffect = effectText.replace(effectBlock, newBlock);
            return true;
        } catch (e) {
            return false;
        }
    }

    // =========================================
    // 調整の実行 / Applying adjustments
    // =========================================

    /**
     * @typedef {object} AdjustState
     * @property {number} anchorIndex - 基準点（0〜8、左上から右下へ）
     * @property {string} fitMode - "cover" / "contain" / "none" / "manual"
     * @property {string} maskMode - "none" / "fitFrame" / "makeSquare"
     * @property {boolean} applyRound - 角丸を付けるか
     * @property {number} manualScale - 指定スケール（％）
     * @property {number} roundRadiusPt - 角丸の半径（pt）
     * @property {number} nudgeX - 横の微調整（pt）
     * @property {number} nudgeY - 縦の微調整（pt）
     */

    /**
     * マスクパスの形を変える（正方形・内容に合わせる）
     * @param {PageItem} clipPath - マスクパス
     * @param {PageItem[]} contents - 内容
     * @param {string} maskMode - "none" / "fitFrame" / "makeSquare"
     * @returns {void}
     */
    function reshapeMaskPath(clipPath, contents, maskMode) {
        if (maskMode === "makeSquare") {
            var maskBounds = clipPath.geometricBounds;
            var maskWidth = maskBounds[2] - maskBounds[0];
            var maskHeight = maskBounds[1] - maskBounds[3];
            var side = Math.min(maskWidth, maskHeight);
            var centerX = maskBounds[0] + maskWidth / 2;
            var centerY = maskBounds[1] - maskHeight / 2;
            clipPath.width = side;
            clipPath.height = side;
            clipPath.position = [centerX - side / 2, centerY + side / 2];
        } else if (maskMode === "fitFrame") {
            var contentBounds = getCombinedBounds(contents);
            if (contentBounds) {
                clipPath.position = [contentBounds[0], contentBounds[1]];
                clipPath.width = contentBounds[2] - contentBounds[0];
                clipPath.height = contentBounds[1] - contentBounds[3];
            }
        }
    }

    /**
     * 範囲の中で基準点に当たる座標を返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {number} anchorIndex - 基準点（0〜8）
     * @returns {number[]} [x, y]
     */
    function getAnchorPoint(bounds, anchorIndex) {
        var anchorCol = anchorIndex % 3;
        var anchorRow = Math.floor(anchorIndex / 3);
        var anchorX = (anchorCol === 0) ? bounds[0] : (anchorCol === 1) ? (bounds[0] + bounds[2]) / 2 : bounds[2];
        var anchorY = (anchorRow === 0) ? bounds[1] : (anchorRow === 1) ? (bounds[1] + bounds[3]) / 2 : bounds[3];
        return [anchorX, anchorY];
    }

    /**
     * 1つのクリップグループに、マスクパスの形・角丸・内容の拡大縮小と位置合わせを適用する
     * @param {GroupItem} clipGroup - 対象のクリップグループ
     * @param {AdjustState} adjustState - 調整の設定
     * @returns {void}
     */
    function adjustClipGroup(clipGroup, adjustState) {
        var clipParts = splitClipGroup(clipGroup);
        var clipPath = clipParts.clipPath;
        var contents = clipParts.contents;
        if (!clipPath || contents.length === 0) return;

        /* マスクパスの形を変えても、後続の「フィットとスケール」は続けて適用する / Fit & scale still follows */
        reshapeMaskPath(clipPath, contents, adjustState.maskMode);

        if (adjustState.applyRound) {
            applyRoundCornersEffect(clipGroup, adjustState.roundRadiusPt);
        }

        /* 位置合わせは visibleBounds を基準に（上揃えがずれるケース対策）/ Align on visibleBounds */
        var frameBounds = clipPath.visibleBounds;
        var frameWidth = frameBounds[2] - frameBounds[0];
        var frameHeight = frameBounds[1] - frameBounds[3];
        var frameAnchor = getAnchorPoint(frameBounds, adjustState.anchorIndex);

        for (var i = 0; i < contents.length; i++) {
            var content = contents[i];
            var contentBounds = content.visibleBounds;
            var contentWidth = contentBounds[2] - contentBounds[0];
            var contentHeight = contentBounds[1] - contentBounds[3];

            var ratio = 1.0;
            if (adjustState.fitMode === "cover") {
                ratio = Math.max(frameWidth / contentWidth, frameHeight / contentHeight);
            } else if (adjustState.fitMode === "contain") {
                ratio = Math.min(frameWidth / contentWidth, frameHeight / contentHeight);
            } else if (adjustState.fitMode === "manual") {
                var currentScale = parseFloat(getContentScalePercent(clipGroup)) / 100;
                ratio = (currentScale === 0) ? 0 : (adjustState.manualScale / 100) / currentScale;
            }

            if (adjustState.fitMode !== "none") {
                content.resize(ratio * 100, ratio * 100, true, true, true, true, ratio * 100);
                contentBounds = content.visibleBounds;
            }

            var contentAnchor = getAnchorPoint(contentBounds, adjustState.anchorIndex);
            content.translate((frameAnchor[0] - contentAnchor[0]) + adjustState.nudgeX, (frameAnchor[1] - contentAnchor[1]) + adjustState.nudgeY);
        }
    }

    /**
     * すべてのクリップグループに調整を適用する
     * @param {GroupItem[]} clipGroups - 対象のクリップグループ
     * @param {AdjustState} adjustState - 調整の設定
     * @returns {void}
     */
    function adjustClipGroups(clipGroups, adjustState) {
        for (var i = 0; i < clipGroups.length; i++) {
            adjustClipGroup(clipGroups[i], adjustState);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる（イベントはまだ付けない）
     * @param {GroupItem} primaryClipGroup - 初期値を読むクリップグループ
     * @returns {Object} ダイアログと各コントロール
     */
    function buildAdjustDialog(primaryClipGroup) {
        var unitLabel = getUnitInfo().label;

        var adjustDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        adjustDialog.orientation = "column";
        adjustDialog.alignChildren = ["center", "top"];
        adjustDialog.spacing = DIALOG_SPACING;

        var columnsGroup = adjustDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 左カラム：基準点・微調整 / Left column: anchor and nudge */
        var anchorColumn = addColumn(columnsGroup);

        var anchorPanel = anchorColumn.add("panel", undefined, getLabel(LABELS.panel.anchor));
        anchorPanel.margins = PANEL_MARGINS;
        anchorPanel.alignChildren = ["center", "center"];
        var anchorGridGroup = anchorPanel.add("group");
        anchorGridGroup.orientation = "column";
        anchorGridGroup.spacing = ROW_SPACING;
        var anchorRadios = [];
        for (var i = 0; i < 3; i++) {
            var anchorRowGroup = anchorGridGroup.add("group");
            for (var j = 0; j < 3; j++) {
                var anchorRadio = anchorRowGroup.add("radiobutton", undefined, "");
                anchorRadio.helpTip = getLabel(LABELS.tooltip.anchor);
                anchorRadio.size = [ANCHOR_CELL_SIZE, ANCHOR_CELL_SIZE];
                anchorRadios.push(anchorRadio);
            }
        }
        anchorRadios[4].value = true;

        var nudgePanel = anchorColumn.add("panel", undefined, getLabel(LABELS.panel.nudge));
        setupPanel(nudgePanel, 6);
        var nudgeXInput = addNudgeRow(nudgePanel, LABELS.fieldLabel.x, unitLabel);
        var nudgeYInput = addNudgeRow(nudgePanel, LABELS.fieldLabel.y, unitLabel);

        /* 中央カラム：フィットとスケール / Middle column: fit and scale */
        var fitPanel = columnsGroup.add("panel", undefined, getLabel(LABELS.panel.fitScale));
        setupPanel(fitPanel, PANEL_SPACING);
        var radioCover = addRadio(fitPanel, LABELS.radio.cover, LABELS.tooltip.cover);
        var radioContain = addRadio(fitPanel, LABELS.radio.contain, LABELS.tooltip.contain);
        var radioKeepSize = addRadio(fitPanel, LABELS.radio.keepSize, LABELS.tooltip.keepSize);

        var manualScaleGroup = addRow(fitPanel);
        var radioManual = addRadio(manualScaleGroup, LABELS.radio.manualScale, LABELS.tooltip.manualScale);
        var scaleInput = addStepperInput(manualScaleGroup, getContentScalePercent(primaryClipGroup), 6, { min: 0 });
        scaleInput.helpTip = getLabel(LABELS.tooltip.manualScale);
        manualScaleGroup.add("statictext", undefined, "%");
        radioKeepSize.value = true;

        /* 右カラム：マスクパス・角丸 / Right column: mask path and round corners */
        var maskColumn = addColumn(columnsGroup);

        var maskPanel = maskColumn.add("panel", undefined, getLabel(LABELS.panel.maskPath));
        setupPanel(maskPanel, PANEL_SPACING);
        var radioMaskUnchanged = addRadio(maskPanel, LABELS.radio.maskUnchanged, LABELS.tooltip.maskUnchanged);
        var radioFitFrame = addRadio(maskPanel, LABELS.radio.fitToContent, LABELS.tooltip.fitToContent);
        var radioSquare = addRadio(maskPanel, LABELS.radio.square, LABELS.tooltip.square);
        radioMaskUnchanged.value = true;

        var roundPanel = maskColumn.add("panel", undefined, getLabel(LABELS.panel.roundCorners));
        setupPanel(roundPanel, PANEL_SPACING);
        var roundRowGroup = addRow(roundPanel);
        var checkboxRound = roundRowGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.roundCorners));
        checkboxRound.helpTip = getLabel(LABELS.tooltip.roundCorners);
        checkboxRound.value = false;
        var defaultRoundRadius = getDefaultRoundRadius(primaryClipGroup);
        var roundInput = addStepperInput(roundRowGroup, (defaultRoundRadius !== null) ? String(defaultRoundRadius) : "10", 4, { min: 0 });
        roundInput.helpTip = getLabel(LABELS.tooltip.roundCorners);
        roundRowGroup.add("statictext", undefined, unitLabel);

        var checkboxCircle = roundPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.circle));
        checkboxCircle.helpTip = getLabel(LABELS.tooltip.circle);
        checkboxCircle.value = false;

        var btnRowGroup = adjustDialog.add("group");
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return {
            dialog: adjustDialog,
            anchorRadios: anchorRadios,
            nudgeXInput: nudgeXInput,
            nudgeYInput: nudgeYInput,
            radioCover: radioCover,
            radioContain: radioContain,
            radioKeepSize: radioKeepSize,
            radioManual: radioManual,
            scaleRadios: [radioCover, radioContain, radioKeepSize, radioManual],
            scaleInput: scaleInput,
            radioMaskUnchanged: radioMaskUnchanged,
            radioFitFrame: radioFitFrame,
            radioSquare: radioSquare,
            checkboxRound: checkboxRound,
            roundInput: roundInput,
            checkboxCircle: checkboxCircle,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * ラジオボタンを追加する
     * @param {Object} parentContainer - 追加先
     * @param {Object} labelSet - 表示名の LABELS リーフ
     * @param {Object} tooltipSet - tooltip の LABELS リーフ
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parentContainer, labelSet, tooltipSet) {
        var radio = parentContainer.add("radiobutton", undefined, getLabel(labelSet));
        radio.helpTip = getLabel(tooltipSet);
        return radio;
    }

    /**
     * 微調整の「項目名＋入力欄＋単位」の行を追加する
     * @param {Panel} nudgePanel - 追加先
     * @param {Object} labelSet - 項目名の LABELS リーフ
     * @param {string} unitLabel - 単位の表示
     * @returns {EditText} 追加した入力欄
     */
    function addNudgeRow(nudgePanel, labelSet, unitLabel) {
        var nudgeRowGroup = addRow(nudgePanel);
        nudgeRowGroup.add("statictext", undefined, labelText(labelSet));
        var nudgeInput = addStepperInput(nudgeRowGroup, "0", 4, {}); /* 負数も可 / negatives allowed */
        nudgeInput.helpTip = getLabel(LABELS.tooltip.nudge);
        nudgeRowGroup.add("statictext", undefined, unitLabel);
        return nudgeInput;
    }

    /**
     * ダイアログの入力から調整の設定を読み取る
     * @param {Object} dialogControls - buildAdjustDialog() の戻り値
     * @returns {AdjustState} 調整の設定
     */
    function readAdjustState(dialogControls) {
        var anchorIndex = 4;
        for (var i = 0; i < dialogControls.anchorRadios.length; i++) {
            if (dialogControls.anchorRadios[i].value) { anchorIndex = i; break; }
        }

        var fitMode = "cover";
        if (dialogControls.radioContain.value) fitMode = "contain";
        else if (dialogControls.radioKeepSize.value) fitMode = "none";
        else if (dialogControls.radioManual.value) fitMode = "manual";

        var maskMode = "none";
        if (dialogControls.radioFitFrame.value) maskMode = "fitFrame";
        else if (dialogControls.radioSquare.value) maskMode = "makeSquare";

        return {
            anchorIndex: anchorIndex,
            fitMode: fitMode,
            maskMode: maskMode,
            applyRound: !!dialogControls.checkboxRound.value,
            manualScale: parseFloat(dialogControls.scaleInput.text) || 100,
            roundRadiusPt: unitValueToPt(parseFloat(dialogControls.roundInput.text) || 0),
            nudgeX: unitValueToPt(parseFloat(dialogControls.nudgeXInput.text) || 0),
            nudgeY: unitValueToPt(parseFloat(dialogControls.nudgeYInput.text) || 0)
        };
    }

    /**
     * 角丸の半径を、マスクパスの短辺の半分（見た目の寸法）にする
     * @param {Object} dialogControls - buildAdjustDialog() の戻り値
     * @param {GroupItem} clipGroup - 対象のクリップグループ
     * @returns {void}
     */
    function setRoundRadiusToHalfOfMask(dialogControls, clipGroup) {
        var clipPath = splitClipGroup(clipGroup).clipPath;
        if (!clipPath) return;
        /* 線幅なども含めた見た目の寸法 / Visible size including the stroke */
        var maskBounds = clipPath.visibleBounds;
        var radiusPt = Math.min(maskBounds[2] - maskBounds[0], maskBounds[1] - maskBounds[3]) / 2;
        dialogControls.roundInput.text = String(Math.round(ptToUnitValue(radiusPt) * 100) / 100);
    }

    /**
     * ダイアログにイベントを付ける
     * @param {Object} dialogControls - buildAdjustDialog() の戻り値
     * @param {GroupItem[]} clipGroups - 対象のクリップグループ
     * @param {PageItem[]} selectedItems - 実行時に選択していたオブジェクト（選択を戻すのに使う）
     * @returns {Function} プレビューを更新する関数
     */
    function bindAdjustDialogEvents(dialogControls, clipGroups, selectedItems) {
        var adjustDialog = dialogControls.dialog;
        var anchorRadios = dialogControls.anchorRadios;
        var primaryClipGroup = clipGroups[0];

        /* 別グループのラジオは自動で排他にならないので手で切り替える / Radios in different groups are exclusive by hand */
        function setScaleRadio(selectedRadio) {
            for (var i = 0; i < dialogControls.scaleRadios.length; i++) {
                dialogControls.scaleRadios[i].value = (dialogControls.scaleRadios[i] === selectedRadio);
            }
        }

        function setAnchorRadio(index) {
            for (var i = 0; i < anchorRadios.length; i++) {
                anchorRadios[i].value = (i === index);
            }
        }

        function resetNudge() {
            dialogControls.nudgeXInput.text = "0";
            dialogControls.nudgeYInput.text = "0";
        }

        /* ［正円］は［正方形に］のときだけ使える / Circle is available only with Square */
        function updateCircleAvailability() {
            dialogControls.checkboxCircle.enabled = !!dialogControls.radioSquare.value;
        }

        /* 角丸を除いてから今の入力で適用し直す（Undo は使わない）/ Re-apply without undo */
        function updatePreview() {
            for (var i = 0; i < clipGroups.length; i++) {
                removeRoundCornersEffect(clipGroups[i]);
            }
            var adjustState = readAdjustState(dialogControls);
            adjustClipGroups(clipGroups, adjustState);
            if (adjustState.fitMode !== "manual") {
                dialogControls.scaleInput.text = getContentScalePercent(primaryClipGroup);
            }
            app.redraw();
        }

        function selectAnchor(index) {
            setAnchorRadio(index);
            resetNudge();
            updatePreview();
        }

        function onScaleInputChanged() {
            setScaleRadio(dialogControls.radioManual);
            updatePreview();
        }

        function onRoundInputChanged() {
            dialogControls.checkboxRound.value = true;
            clearAppearanceForClipGroups(clipGroups, selectedItems);
            updatePreview();
        }

        /* 基準点のキーボードショートカット / Anchor keyboard shortcuts */
        var ANCHOR_KEY_MAP = { "q": 0, "w": 1, "e": 2, "a": 3, "s": 4, "d": 5, "z": 6, "x": 7, "c": 8 };
        adjustDialog.addEventListener("keydown", function (event) {
            if (adjustDialog.activeControl === dialogControls.scaleInput || adjustDialog.activeControl === dialogControls.roundInput) return;
            var keyName = event.keyName.toLowerCase();
            if (ANCHOR_KEY_MAP.hasOwnProperty(keyName)) selectAnchor(ANCHOR_KEY_MAP[keyName]);
        });

        for (var i = 0; i < anchorRadios.length; i++) {
            (function (anchorIndex) {
                anchorRadios[anchorIndex].onClick = function () { selectAnchor(anchorIndex); };
            })(i);
        }

        dialogControls.nudgeXInput.onChanging = updatePreview;
        dialogControls.nudgeYInput.onChanging = updatePreview;

        for (var k = 0; k < dialogControls.scaleRadios.length; k++) {
            dialogControls.scaleRadios[k].onClick = function () { setScaleRadio(this); updatePreview(); };
        }
        dialogControls.scaleInput.onChanging = onScaleInputChanged;

        dialogControls.radioSquare.onClick = function () {
            /* ［正方形に］は中央基準＋［縦横比を保持して切り取り］に固定。角丸の値は変えない
               Square locks the anchor to center and the fit to Cover; the radius is left as is */
            setScaleRadio(dialogControls.radioCover);
            setAnchorRadio(4);
            updateCircleAvailability();
            updatePreview();
        };
        dialogControls.radioFitFrame.onClick = function () { updateCircleAvailability(); updatePreview(); };
        dialogControls.radioMaskUnchanged.onClick = function () { updateCircleAvailability(); updatePreview(); };

        dialogControls.checkboxRound.onClick = function () {
            /* OFF にしたらアピアランスを消去して重複を防ぐ / Clear the appearance when turned off */
            if (!dialogControls.checkboxRound.value) {
                clearAppearanceForClipGroups(clipGroups, selectedItems);
            }
            updatePreview();
        };
        dialogControls.roundInput.onChanging = onRoundInputChanged;

        dialogControls.checkboxCircle.onClick = function () {
            /* OFF にしただけでは処理しない（［角を丸くする］が新しく付くのを防ぐ）/ Turning off does nothing */
            if (dialogControls.checkboxCircle.value) {
                dialogControls.checkboxRound.value = true;
                setRoundRadiusToHalfOfMask(dialogControls, primaryClipGroup);
                var radiusPt = unitValueToPt(parseFloat(dialogControls.roundInput.text) || 0);
                /* 既存の角丸があれば半径だけ更新し、無ければアピアランスを消去して付け直す
                   Update the radius in place; otherwise clear the appearance and re-apply */
                if (!updateRoundCornersRadiusOnly(primaryClipGroup, radiusPt)) {
                    clearAppearanceForClipGroups(clipGroups, selectedItems);
                    applyRoundCornersEffect(primaryClipGroup, radiusPt);
                }
            }
            app.redraw();
        };

        dialogControls.btnOK.onClick = function () {
            var adjustState = readAdjustState(dialogControls);
            /* 確定時に1回：アピアランスを消去してから入力の状態で適用 / Clear once, then apply the final state */
            clearAppearanceForClipGroups(clipGroups, selectedItems);
            adjustClipGroups(clipGroups, adjustState);
            app.redraw();
            adjustDialog.close();
        };
        dialogControls.btnCancel.onClick = function () {
            /* Undo を使わないため、反映済みの状態のまま閉じる / No undo: close with the preview as is */
            adjustDialog.close();
        };

        updateCircleAvailability();
        return updatePreview;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のクリップグループを調整するダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        /* 選択はあとで戻すので、配列に控えておく / Keep a copy of the selection to restore later */
        var selectedItems = [];
        var docSelection = app.activeDocument.selection;
        for (var i = 0; i < docSelection.length; i++) selectedItems.push(docSelection[i]);
        var clipGroups = filterClipGroups(selectedItems);
        if (clipGroups.length === 0) {
            alert(getLabel(LABELS.alert.selectClipGroup));
            return;
        }

        var dialogControls = buildAdjustDialog(clipGroups[0]);
        var updatePreview = bindAdjustDialogEvents(dialogControls, clipGroups, selectedItems);

        /* 開く前に1回：アピアランスを消去してからプレビュー / Clear the appearance once, then preview */
        clearAppearanceForClipGroups(clipGroups, selectedItems);
        updatePreview();
        dialogControls.dialog.show();
    }

    main();

})();
