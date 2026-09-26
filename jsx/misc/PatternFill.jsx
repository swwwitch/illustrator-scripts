#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

2つのオブジェクトを選択し、大きい方を「容器」、小さい方を「タイル」として、容器のバウンディングボックス内にタイルを等間隔で敷き詰めます。
完全に内側に収まるタイルのみを残し、元の2オブジェクトはそのまま残します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternFill.md

### Overview

With two objects selected, treats the larger as the container and the smaller as the tile, then fills the container's bounding box with an even grid of tiles.
Only the tiles that fit entirely inside are kept, and both originals are left in place.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PatternFill.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PatternFill";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternFill.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PatternFill.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 間隔の初期値を決める比率（タイルの幅に掛ける）/ Initial spacing as a ratio of the tile width */
    var DEFAULT_GAP_RATIO = 0.2;

    /* プレビューの不透明度（%）/ Opacity of the preview (%) */
    var PREVIEW_OPACITY = 60;

    /* 敷き詰めたタイルを入れるグループ名（プレビューは名前で探して消す）/ Group names; previews are found and removed by name */
    var GRID_GROUP_NAME = 'TiledGrid';
    var PREVIEW_GROUP_NAME = 'TiledGrid_preview';

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OFFSET_X = 300;    /* ダイアログを右へずらす量（px）/ Horizontal shift of the dialog */
    var DIALOG_OPACITY = 0.98;    /* ダイアログの不透明度 / Dialog opacity */
    var LABEL_WIDTH = 60;         /* 項目名の幅 / Field label width */
    var FIELD_CHARACTERS = 4;     /* 数値入力欄の桁数 / Width of the number fields in characters */

    /**
     * ダイアログを表示するときに、既定の位置からずらす
     * @param {Window} targetDialog - 対象ダイアログ
     * @param {number} offsetX - 横のずらし量
     * @param {number} offsetY - 縦のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function() {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    /**
     * 右揃えの項目名と入力欄を並べる行を作る
     * @param {Window} parentDialog - 親ダイアログ
     * @param {string} labelPath - 項目名のラベルのパス
     * @returns {Group} 項目名を追加済みの行
     */
    function addFieldRow(parentDialog, labelPath) {
        var fieldRow = parentDialog.add('group');
        fieldRow.alignment = ['fill', 'top'];
        fieldRow.alignChildren = ['left', 'center'];
        var fieldLabel = fieldRow.add('statictext', undefined, labelText(labelPath));
        fieldLabel.justify = 'right';
        fieldLabel.preferredSize.width = LABEL_WIDTH;
        return fieldRow;
    }

    /**
     * 行に数値入力欄を、左に∧∨を付けて追加する（↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} fieldRow - 追加先の行
     * @param {string} initialText - 初期値
     * @param {string} tooltipPath - helpTip のラベルのパス
     * @param {Object} stepOptions - ∧∨の設定（min / integer / onStep など。addStepper() を参照）
     * @returns {EditText} 追加した入力欄
     */
    function addNumberField(fieldRow, initialText, tooltipPath, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldRow.add('group');
        stepperFieldGroup.orientation = 'row';
        stepperFieldGroup.alignChildren = ['left', 'center'];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var numberField;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberField; }, stepOptions);
        numberField = stepperFieldGroup.add('edittext', undefined, initialText);
        numberField.characters = FIELD_CHARACTERS;
        numberField.helpTip = getLabel(tooltipPath);
        bindSteppedArrowKeys(numberField, stepperGroup);
        return numberField;
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
        var upTooltip = stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp";
        var downTooltip = stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown";
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

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
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

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "敷き詰め設定", en: "Tile Fill Settings" }
        },
        fieldLabel: {
            gridCount: { ja: "グリッド数", en: "Grid count" },
            spacing: { ja: "間隔", en: "Spacing" },
            margin: { ja: "マージン", en: "Margin" }
        },
        unit: {
            columns: { ja: "列", en: "Columns" },
            rows: { ja: "行", en: "Rows" }
        },
        checkbox: {
            brick: { ja: "レンガ状", en: "Brick pattern" },
            symbolize: { ja: "シンボル化", en: "Symbolize" }
        },
        tooltip: {
            columns: {
                ja: "横に並べる数です。0 のときは容器に収まるだけ自動で並べます。",
                en: "How many tiles to place horizontally. 0 fills the container automatically."
            },
            rows: {
                ja: "縦に並べる数です。0 のときは容器に収まるだけ自動で並べます。",
                en: "How many tiles to place vertically. 0 fills the container automatically."
            },
            spacing: {
                ja: "隣り合うタイルのアキです。縦横とも同じ値になります。",
                en: "Space between neighbouring tiles, applied both horizontally and vertically."
            },
            margin: {
                ja: "容器の内側に空ける余白です。負の値も入力できます。",
                en: "Inset kept inside the container. Negative values are allowed."
            },
            brick: {
                ja: "1行おきに半個分ずらして、レンガのように並べます。",
                en: "Offsets every other row by half a tile, like brickwork."
            },
            symbolize: {
                ja: "タイルをシンボルとして複製します。あとからまとめて差し替えられます。",
                en: "Duplicates the tile as a symbol, so every copy can be swapped later at once."
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
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "fieldLabel.spacing" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 単位を括弧でくくった表記を返す（日本語は全角括弧）
     * @param {string} unitLabel - 単位の表示名
     * @returns {string} 括弧付きの単位
     */
    function unitSuffix(unitLabel) {
        return (uiLang === "ja") ? ("（" + unitLabel + "）") : (" (" + unitLabel + ")");
    }

    // =========================================
    // 形状と配置の計算 / Geometry
    // =========================================

    /**
     * オブジェクトの表示上の境界と寸法を取得する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number, area: number}} 境界情報
     */
    function getBoundsInfo(pageItem) {
        var visibleBounds = pageItem.visibleBounds; /* [left, top, right, bottom] */
        var itemWidth = visibleBounds[2] - visibleBounds[0];
        var itemHeight = visibleBounds[1] - visibleBounds[3];
        return {
            left: visibleBounds[0],
            top: visibleBounds[1],
            right: visibleBounds[2],
            bottom: visibleBounds[3],
            width: itemWidth,
            height: itemHeight,
            area: Math.abs(itemWidth * itemHeight)
        };
    }

    /**
     * 4点の閉じたパスで、すべてのアンカーが指定の種類か判定する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {PointType} pointType - アンカーの種類
     * @returns {boolean} 当てはまれば true
     */
    function isFourPointPathOfType(pageItem, pointType) {
        if (pageItem.typename !== 'PathItem' || !pageItem.closed || pageItem.pathPoints.length !== 4) return false;
        for (var i = 0; i < 4; i++) {
            if (pageItem.pathPoints[i].pointType !== pointType) return false;
        }
        return true;
    }

    /**
     * 長方形のパス（4点すべてがコーナー）か判定する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {boolean} 長方形なら true
     */
    function isRectanglePath(pageItem) {
        return isFourPointPathOfType(pageItem, PointType.CORNER);
    }

    /**
     * 楕円らしいパス（4点すべてがスムーズ）か判定する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {boolean} 楕円らしければ true
     */
    function isEllipseLikePath(pageItem) {
        return isFourPointPathOfType(pageItem, PointType.SMOOTH);
    }

    /**
     * 並べる列数と行数を決める。指定が 0 のときは容器に収まるだけ並べる
     * @param {Object} containerInfo - 容器の境界情報
     * @param {number} stepX - 横の送り（pt）
     * @param {number} stepY - 縦の送り（pt）
     * @param {boolean} isBrick - レンガ状に並べるか
     * @param {number} fixedColumns - 指定の列数（0 で自動）
     * @param {number} fixedRows - 指定の行数（0 で自動）
     * @returns {{columns: number, rows: number}} 列数と行数
     */
    function computeGridSize(containerInfo, stepX, stepY, isBrick, fixedColumns, fixedRows) {
        return {
            columns: (fixedColumns > 0)
                ? Math.round(fixedColumns)
                : Math.max(1, Math.ceil((containerInfo.width + (isBrick ? stepX * 0.5 : 0)) / stepX)),
            rows: (fixedRows > 0)
                ? Math.round(fixedRows)
                : Math.max(1, Math.ceil(containerInfo.height / stepY))
        };
    }

    /**
     * グリッドの各マスの左上座標を、行ごとに左から順に渡す（レンガ状なら奇数行を半個ずらす）
     * @param {Object} containerInfo - 容器の境界情報
     * @param {number} stepX - 横の送り（pt）
     * @param {number} stepY - 縦の送り（pt）
     * @param {{columns: number, rows: number}} gridSize - 列数と行数
     * @param {boolean} isBrick - レンガ状に並べるか
     * @param {Function} visitCell - (xLeft, yTop) を受け取る関数
     * @returns {void}
     */
    function forEachGridCell(containerInfo, stepX, stepY, gridSize, isBrick, visitCell) {
        for (var rowIndex = 0; rowIndex < gridSize.rows; rowIndex++) {
            var yTop = containerInfo.top - rowIndex * stepY;
            for (var columnIndex = 0; columnIndex < gridSize.columns; columnIndex++) {
                var xLeft = containerInfo.left + columnIndex * stepX + ((isBrick && (rowIndex % 2 === 1)) ? stepX * 0.5 : 0);
                visitCell(xLeft, yTop);
            }
        }
    }

    /**
     * タイルを残すかどうかの判定に使う、容器の形とマージンを差し引いた範囲をまとめる
     * @param {PageItem} container - 容器
     * @param {Object} containerInfo - 容器の境界情報
     * @param {number} marginPt - マージン（pt）
     * @returns {Object} 容器の形（isRect / isEllipse）と判定に使う範囲
     */
    function buildContainerShape(container, containerInfo, marginPt) {
        var isPath = (container.typename === 'PathItem');
        return {
            isRect: isPath && isRectanglePath(container),
            isEllipse: isPath && isEllipseLikePath(container),
            left: containerInfo.left + marginPt,
            right: containerInfo.right - marginPt,
            top: containerInfo.top - marginPt,
            bottom: containerInfo.bottom + marginPt,
            centerX: (containerInfo.left + containerInfo.right) / 2.0,
            centerY: (containerInfo.top + containerInfo.bottom) / 2.0,
            radiusX: Math.max(0, Math.abs(containerInfo.right - containerInfo.left) / 2.0 - marginPt),
            radiusY: Math.max(0, Math.abs(containerInfo.top - containerInfo.bottom) / 2.0 - marginPt)
        };
    }

    /**
     * タイルが容器の内側に収まっているか判定する。
     * 長方形は外接矩形が丸ごと内側、楕円は4隅が楕円の内側、それ以外は中心が内側なら残す
     * @param {Object} tileBounds - タイルの境界情報
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {boolean} 残すなら true
     */
    function isTileInside(tileBounds, containerShape) {
        if (containerShape.isRect) {
            return (tileBounds.left >= containerShape.left && tileBounds.right <= containerShape.right &&
                tileBounds.top <= containerShape.top && tileBounds.bottom >= containerShape.bottom);
        }
        if (containerShape.isEllipse) {
            if (!(containerShape.radiusX > 0 && containerShape.radiusY > 0)) return false;
            var corners = [
                [tileBounds.left, tileBounds.top],
                [tileBounds.right, tileBounds.top],
                [tileBounds.left, tileBounds.bottom],
                [tileBounds.right, tileBounds.bottom]
            ];
            for (var k = 0; k < 4; k++) {
                var dx = (corners[k][0] - containerShape.centerX) / containerShape.radiusX;
                var dy = (corners[k][1] - containerShape.centerY) / containerShape.radiusY;
                if (!((dx * dx + dy * dy) <= 1.000001)) return false;
            }
            return true;
        }
        var tileCenterX = (tileBounds.left + tileBounds.right) / 2.0;
        var tileCenterY = (tileBounds.top + tileBounds.bottom) / 2.0;
        return (tileCenterX >= containerShape.left && tileCenterX <= containerShape.right &&
            tileCenterY <= containerShape.top && tileCenterY >= containerShape.bottom);
    }

    /**
     * グループ内のタイルのうち、容器に収まらないものを削除する（マスクは使わない）
     * @param {GroupItem} tileGroup - タイルを入れたグループ
     * @param {Object} containerShape - buildContainerShape() の結果
     * @returns {void}
     */
    function trimTilesOutside(tileGroup, containerShape) {
        for (var i = tileGroup.pageItems.length - 1; i >= 0; i--) {
            var gridItem = tileGroup.pageItems[i];
            if (!isTileInside(getBoundsInfo(gridItem), containerShape)) gridItem.remove();
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 敷き詰め設定のダイアログを表示する
     * @param {number} tileWidthPt - タイルの幅（pt）。間隔の初期値に使う
     * @param {Function} onPreview - 入力が変わるたびに設定（pt 換算済み）を受け取る関数
     * @returns {{gap: number, margin: number, isBrick: boolean, symbolize: boolean, columns: number, rows: number}|null} 設定（pt）。キャンセル時は null
     */
    function showFillDialog(tileWidthPt, onPreview) {
        /* ダイアログのタイトルはラベル＋バージョン / Dialog title = label + version */
        var fillDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);

        /* ウィンドウ見た目 / Window appearance */
        fillDialog.opacity = DIALOG_OPACITY;
        shiftDialogPosition(fillDialog, DIALOG_OFFSET_X, 0);
        fillDialog.alignChildren = 'fill';

        var rulerUnit = getUnitInfo("rulerType");
        var unitLabel = rulerUnit.label;

        /* グリッド数（0 で容器に合わせて自動）/ Grid count (0 fits the container automatically) */
        var gridCountRow = addFieldRow(fillDialog, 'fieldLabel.gridCount');
        var columnsField = addNumberField(gridCountRow, '0', 'tooltip.columns',
            { min: 0, integer: true, onStep: updatePreviewFromFields });
        gridCountRow.add('statictext', undefined, getLabel('unit.columns'));
        var rowsField = addNumberField(gridCountRow, '0', 'tooltip.rows',
            { min: 0, integer: true, onStep: updatePreviewFromFields });
        gridCountRow.add('statictext', undefined, getLabel('unit.rows'));

        /* 間隔 / Spacing */
        var spacingRow = addFieldRow(fillDialog, 'fieldLabel.spacing');
        var defaultGapValue = Math.round(tileWidthPt / rulerUnit.pointsPerUnit * DEFAULT_GAP_RATIO);
        var gapField = addNumberField(spacingRow, String(defaultGapValue), 'tooltip.spacing',
            { min: 0, onStep: updatePreviewFromFields });
        spacingRow.add('statictext', undefined, unitSuffix(unitLabel));
        gapField.active = true;

        /* マージン / Margin */
        var marginRow = addFieldRow(fillDialog, 'fieldLabel.margin');
        /* マージンは負OK / margin can be negative */
        var marginField = addNumberField(marginRow, '0', 'tooltip.margin', { onStep: updatePreviewFromFields });
        marginRow.add('statictext', undefined, unitSuffix(unitLabel));

        /**
         * 入力欄の値を pt に換算して読む（チェックボックスは作成前なら false として扱う）
         * @returns {{gap: number, margin: number, isBrick: boolean, symbolize: boolean, columns: number, rows: number}} 設定（pt）
         */
        function readFillSettings() {
            var pointsPerUnit = getUnitInfo("rulerType").pointsPerUnit;
            var gapValue = Math.max(0, parseFloat(gapField.text) || 0);
            var marginValue = parseFloat(marginField.text);
            if (isNaN(marginValue)) marginValue = 0;
            return {
                gap: gapValue * pointsPerUnit,
                margin: marginValue * pointsPerUnit,
                isBrick: !!(brickCheckbox && brickCheckbox.value),
                symbolize: !!(symbolizeCheckbox && symbolizeCheckbox.value),
                columns: Math.max(0, parseInt(columnsField.text, 10) || 0),
                rows: Math.max(0, parseInt(rowsField.text, 10) || 0)
            };
        }

        /**
         * 現在の入力内容でプレビューを描き直す
         * @returns {void}
         */
        function updatePreviewFromFields() {
            onPreview(readFillSettings());
        }
        gapField.onChanging = updatePreviewFromFields;
        columnsField.onChanging = updatePreviewFromFields;
        rowsField.onChanging = updatePreviewFromFields;
        updatePreviewFromFields();

        /* レンガ状・シンボル化は中央にまとめる / Brick pattern and Symbolize sit centered */
        var optionRow = fillDialog.add('group');
        optionRow.alignment = ['fill', 'top'];
        optionRow.alignChildren = ['center', 'center'];

        /* レンガ状 / Brick pattern */
        var brickGroup = optionRow.add('group');
        brickGroup.alignChildren = ['left', 'center'];
        var brickCheckbox = brickGroup.add('checkbox', undefined, getLabel('checkbox.brick'));
        brickCheckbox.helpTip = getLabel('tooltip.brick');
        brickCheckbox.value = false;
        brickCheckbox.onClick = updatePreviewFromFields;

        /* シンボル化して複製 / Duplicate as Symbol */
        var symbolizeGroup = optionRow.add('group');
        symbolizeGroup.alignChildren = ['left', 'center'];
        var symbolizeCheckbox = symbolizeGroup.add('checkbox', undefined, getLabel('checkbox.symbolize'));
        symbolizeCheckbox.helpTip = getLabel('tooltip.symbolize');
        symbolizeCheckbox.value = false;

        /* ボタン / Buttons */
        var btnRowGroup = fillDialog.add('group');
        btnRowGroup.alignment = 'right';
        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        if (fillDialog.show() !== 1) return null;

        /* OKで確定値をptに変換 / Convert to pt on OK */
        return readFillSettings();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した2つのオブジェクトを容器とタイルに振り分け、ダイアログで設定して敷き詰める
     * @returns {void}
     */
    function main() {
        var doc = app.documents.length ? app.activeDocument : null;
        if (!doc) {
            alert('ドキュメントが開いていません / No active document.');
            return;
        }

        if (!doc.selection || doc.selection.length !== 2) {
            alert('ちょうど2つを選択してください。/ Please select exactly two objects.');
            return;
        }

        /* 容器とタイルを面積で判定 / Detect container and tile by area */
        var firstItem = doc.selection[0];
        var secondItem = doc.selection[1];
        var firstInfo = getBoundsInfo(firstItem);
        var secondInfo = getBoundsInfo(secondItem);
        var isFirstContainer = (firstInfo.area >= secondInfo.area);
        var fillTarget = {
            container: isFirstContainer ? firstItem : secondItem,
            tile: isFirstContainer ? secondItem : firstItem,
            containerInfo: isFirstContainer ? firstInfo : secondInfo,
            tileInfo: isFirstContainer ? secondInfo : firstInfo,
            previewGroup: null
        };

        if (fillTarget.tileInfo.width <= 0 || fillTarget.tileInfo.height <= 0) {
            alert('タイルのサイズが不正です / Invalid tile size.');
            return;
        }

        var fillSettings = showFillDialog(fillTarget.tileInfo.width, function(previewSettings) {
            renderPreview(fillTarget, previewSettings);
        });
        clearPreview(fillTarget);
        if (!fillSettings) {
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            return;
        }

        placeTileGrid(doc, fillTarget, fillSettings);
    }

    /**
     * プレビューを削除する（参照で消し、さらに名前で残りを掃除する）
     * @param {Object} fillTarget - 容器・タイル・プレビューの参照
     * @returns {void}
     */
    function clearPreview(fillTarget) {
        /* 削除済みのグループは参照が無効になる / The group may already be gone */
        try {
            if (fillTarget.previewGroup && fillTarget.previewGroup.isValid) {
                fillTarget.previewGroup.remove();
            }
        } catch (e) {}
        fillTarget.previewGroup = null;

        /* 名前で余分なプレビューを掃除。ロックされたレイヤーなどで削除できないものは飛ばす
           Sweep stray previews by name; skip any that cannot be removed, e.g. on a locked layer */
        try {
            var containerLayer = fillTarget.container.layer;
            for (var i = containerLayer.groupItems.length - 1; i >= 0; i--) {
                var groupItem = containerLayer.groupItems[i];
                if (groupItem.name.indexOf(PREVIEW_GROUP_NAME) === 0) {
                    try {
                        groupItem.remove();
                    } catch (e) {}
                }
            }
        } catch (e) {}
    }

    /**
     * プレビューを描画する（半透明のグループにタイルを複製し、はみ出すものを削除）
     * @param {Object} fillTarget - 容器・タイル・プレビューの参照
     * @param {Object} previewSettings - showFillDialog() と同じ形の設定（pt）
     * @returns {void}
     */
    function renderPreview(fillTarget, previewSettings) {
        var marginPt = previewSettings.margin || 0;
        var isBrick = !!previewSettings.isBrick;
        clearPreview(fillTarget);

        var stepX = fillTarget.tileInfo.width + previewSettings.gap;
        var stepY = fillTarget.tileInfo.height + previewSettings.gap;
        if (stepX <= 0 || stepY <= 0) return;
        var gridSize = computeGridSize(fillTarget.containerInfo, stepX, stepY, isBrick, previewSettings.columns, previewSettings.rows);

        var previewGroup = fillTarget.container.layer.groupItems.add();
        previewGroup.name = PREVIEW_GROUP_NAME;
        fillTarget.previewGroup = previewGroup;

        var containerShape = buildContainerShape(fillTarget.container, fillTarget.containerInfo, marginPt);

        forEachGridCell(fillTarget.containerInfo, stepX, stepY, gridSize, isBrick, function(xLeft, yTop) {
            var tileCopy = fillTarget.tile.duplicate(previewGroup, ElementPlacement.PLACEATBEGINNING);
            tileCopy.position = [xLeft, yTop];
        });

        /* プレビューもトリム / Trim preview too */
        trimTilesOutside(previewGroup, containerShape);

        try {
            previewGroup.opacity = PREVIEW_OPACITY;
        } catch (e) {}
        app.redraw();
    }

    /**
     * タイルをシンボルとして登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} tile - タイル
     * @returns {Symbol|null} 登録したシンボル。失敗時は null（通常の複製に切り替える）
     */
    function createTileSymbol(doc, tile) {
        /* シンボルにできないオブジェクトがある / Some items cannot become symbols */
        try {
            var symbolSource = tile.duplicate();
            var symbolDefinition = doc.symbols.add(symbolSource);
            try { symbolSource.remove(); } catch (e) {}
            return symbolDefinition;
        } catch (e) {
            return null; /* 失敗時は通常複製 / fallback */
        }
    }

    /**
     * 確定した設定でタイルを敷き詰め、容器からはみ出すものを削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} fillTarget - 容器・タイルの参照
     * @param {Object} fillSettings - showFillDialog() の結果（pt）
     * @returns {void}
     */
    function placeTileGrid(doc, fillTarget, fillSettings) {
        var tile = fillTarget.tile;
        var tileInfo = fillTarget.tileInfo;
        var containerInfo = fillTarget.containerInfo;

        /* 必要ならシンボル作成 / Symbolize if needed */
        var symbolDefinition = fillSettings.symbolize ? createTileSymbol(doc, tile) : null;

        /* 配置数を計算 / Compute counts */
        var stepX = tileInfo.width + fillSettings.gap;
        var stepY = tileInfo.height + fillSettings.gap;
        var gridSize = computeGridSize(containerInfo, stepX, stepY, fillSettings.isBrick, fillSettings.columns, fillSettings.rows);

        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        app.redraw();

        /* 複製グループを作成 / Create group for final duplicates */
        var gridGroup = fillTarget.container.layer.groupItems.add();
        gridGroup.name = GRID_GROUP_NAME;

        /* 敷き詰め実行 / Duplicate and place in grid */
        forEachGridCell(containerInfo, stepX, stepY, gridSize, fillSettings.isBrick, function(xLeft, yTop) {
            /* 同じ場所に元タイルがある場合はスキップ / Skip if same as original tile */
            if (Math.abs(xLeft - tileInfo.left) < 0.01 && Math.abs(yTop - tileInfo.top) < 0.01) return;
            if (fillSettings.symbolize && symbolDefinition) {
                var symbolItem = doc.symbolItems.add(symbolDefinition);
                symbolItem.move(gridGroup, ElementPlacement.PLACEATBEGINNING);
                symbolItem.position = [xLeft, yTop];
            } else {
                var tileCopy = tile.duplicate(gridGroup, ElementPlacement.PLACEATBEGINNING);
                tileCopy.position = [xLeft, yTop];
            }
        });

        /* マスクなしトリム / Trim without mask */
        trimTilesOutside(gridGroup, buildContainerShape(fillTarget.container, containerInfo, fillSettings.margin));

        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        app.redraw();
    }

    /**
     * メイン処理を実行し、例外はアラートで知らせる
     * @returns {void}
     */
    function runMainWithErrorAlert() {
        try {
            main();
        } catch (e) {
            alert('[TileSmallIntoLarge] Error:\n' + e);
        }
    }

    /* メインを1アクションで実行（ScriptLanguage / UndoModes がある環境のみ）/ Run main in single undo where the enums exist */
    try {
        if (typeof app.doScript === 'function' && typeof ScriptLanguage !== 'undefined' && typeof UndoModes !== 'undefined') {
            app.doScript(runMainWithErrorAlert, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'TileSmallIntoLarge');
        } else {
            runMainWithErrorAlert();
        }
    } catch (e) {
        runMainWithErrorAlert();
    }

})();
