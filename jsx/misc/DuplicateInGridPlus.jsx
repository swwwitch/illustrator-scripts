#target illustrator
#targetengine "DuplicateInGridPlusEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、グリッド／行／列／ランダムのいずれかの方式で複製・配置します。
繰り返し数・間隔・方向・アートボードへの敷き詰めを2カラムのダイアログで指定でき、結果はライブプレビューで確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DuplicateInGridPlus.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n228720785a71

### Overview

Duplicates and lays out the selected objects as a grid, a row, a column, or at random.
Repeat count, spacing, direction and filling the artboard are set in a two-column dialog, with a live preview of the result.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DuplicateInGridPlus.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DuplicateInGridPlus";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DuplicateInGridPlus.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DuplicateInGridPlus.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n228720785a71"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビュー用レイヤーと一時オブジェクトの識別タグ / Preview layer and temporary-item tag */
    var PREVIEW_LAYER_NAME = "_preview";
    var PREVIEW_ITEM_TAG = "__grid_preview__";

    /* 繰り返し数の下限・上限（スライダーの範囲）/ Repeat count range (slider bounds) */
    var REPEAT_COUNT_MIN = 1;
    var REPEAT_COUNT_MAX = 20;

    /* プレビュー更新の最小間隔（ミリ秒）/ Minimum interval between preview updates (ms) */
    var PREVIEW_THROTTLE_MS = 120;

    /* 画面ズームの範囲 / Zoom range */
    var VIEW_ZOOM_MIN = 0.1;
    var VIEW_ZOOM_MAX = 16;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS    = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING    = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS     = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING     = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING    = 12;                 /* 2カラムの間隔 / gap between columns */
    var FIELD_ROW_SPACING = 20;                 /* 入力欄と連動アイコンの間隔 / gap between fields and the link icon */
    var FIELD_CHARS       = 4;                  /* 数値入力欄の文字数 / width of numeric fields */
    var ZOOM_SLIDER_WIDTH = 240;                /* ズームスライダーの幅 / zoom slider width */
    var ZOOM_GROUP_MARGINS = [0, 0, 0, 10];     /* ズームの行の余白 / zoom row margins */
    var DIALOG_OFFSET_X   = 300;                /* ダイアログを右へずらす量 / horizontal dialog offset */

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
    // 6. この欄に別の↑↓キー処理を付けない（二重に効く）
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
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
    // 2. アイコンを addLinkToggle(親, 初期値, 切り替え後の関数) で作る。helpTip はコピー先で付ける
    //      var linkToggle = addLinkToggle(fieldsRowGroup, true, function () { syncFields(); });
    //      linkToggle.helpTip = getLabel(LABELS.tooltip.linkToggle);
    // 3. 連動中かは linkToggle.value で読む。コードから変えるときは setLinkToggleValue(linkToggle, true)
    // 4. 有効／無効は setLinkToggleEnabled(linkToggle, isEnabled)（無効の間はクリックが効かず、薄く描く）
    // 5. 2つの入力欄の右に置くときは、行 group の中に「入力欄を縦に積んだ group」とアイコンを並べ、
    //    行 group の alignChildren を ["left", "center"] にすると上下の中央に来る
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkLinkToggleUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var LINK_UI_DARK = isDarkLinkToggleUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle) {
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = LINK_ICON_SIZE;
        linkToggle.minimumSize = LINK_ICON_SIZE;
        linkToggle.maximumSize = LINK_ICON_SIZE;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = LINK_ICON_SIZE[0];
            var iconHeight = LINK_ICON_SIZE[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22;
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.toLowerCase().indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /* ラベル定義 / Label definitions (JA/EN) */
    var LABELS = {
        dialog: {
            title: { ja: "複製配置（グリッド／行／列／ランダム）", en: "Duplicate & Arrange" }
        },
        panel: {
            repeatCount: { ja: "繰り返し数", en: "Count" },
            repeatMethod: { ja: "繰り返し方式", en: "Repeat Method" },
            gap: { ja: "間隔（{unit}）", en: "Gap ({unit})" },
            direction: { ja: "方向", en: "Direction" },
            fill: { ja: "敷き詰め", en: "Fill" }
        },
        fieldLabel: {
            countHorizontal: { ja: "横", en: "Horizontal" },
            countVertical: { ja: "縦", en: "Vertical" },
            gapHorizontal: { ja: "左右", en: "Horizontal" },
            gapVertical: { ja: "上下", en: "Vertical" },
            directionHorizontal: { ja: "横方向", en: "Horizontal" },
            directionVertical: { ja: "縦方向", en: "Vertical" },
            zoom: { ja: "画面ズーム", en: "Zoom" }
        },
        radio: {
            methodGrid: { ja: "グリッド", en: "Grid" },
            methodRow: { ja: "行", en: "Row" },
            methodColumn: { ja: "列", en: "Column" },
            methodRandom: { ja: "ランダム配置", en: "Random" },
            directionRight: { ja: "右", en: "Right" },
            directionLeft: { ja: "左", en: "Left" },
            directionUp: { ja: "上", en: "Up" },
            directionDown: { ja: "下", en: "Down" }
        },
        checkbox: {
            fillToEdge: { ja: "アートボードの端まで", en: "Fill to Artboard Edge" },
            fillFull: { ja: "アートボードいっぱいに", en: "Fill Full Artboard" },
            lightMode: { ja: "軽量モード", en: "Light mode" }
        },
        tooltip: {
            countHorizontal: { ja: "横方向に並べる数です（元のオブジェクトを含む）。", en: "How many to place horizontally, including the original." },
            countVertical: { ja: "縦方向に並べる数です（元のオブジェクトを含む）。", en: "How many to place vertically, including the original." },
            countLink: { ja: "横と縦の数を同じにします。", en: "Keeps the horizontal and vertical counts the same." },
            countSlider: { ja: "繰り返し数をまとめて変更します。", en: "Changes the counts together." },
            methodGrid: { ja: "横と縦の両方に並べます。", en: "Places copies both horizontally and vertically." },
            methodRow: { ja: "横一列に並べます。", en: "Places copies in a single row." },
            methodColumn: { ja: "縦一列に並べます。", en: "Places copies in a single column." },
            methodRandom: { ja: "グリッドの枠内でランダムにずらして配置します。", en: "Scatters the copies randomly within the grid area." },
            gapHorizontal: { ja: "隣り合うオブジェクトの左右のアキです。", en: "Space between neighbouring objects horizontally." },
            gapVertical: { ja: "隣り合うオブジェクトの上下のアキです。", en: "Space between neighbouring objects vertically." },
            gapLink: { ja: "左右と上下の間隔を同じにします。", en: "Keeps the horizontal and vertical gaps the same." },
            directionRight: { ja: "元のオブジェクトの右へ複製します。", en: "Duplicates to the right of the original." },
            directionLeft: { ja: "元のオブジェクトの左へ複製します。", en: "Duplicates to the left of the original." },
            directionUp: { ja: "元のオブジェクトの上へ複製します。", en: "Duplicates above the original." },
            directionDown: { ja: "元のオブジェクトの下へ複製します。", en: "Duplicates below the original." },
            fillToEdge: { ja: "アートボードの端に届くまで数を自動で増やします。", en: "Increases the count automatically until the copies reach the artboard edge." },
            fillFull: { ja: "アートボード全面を埋めるように、元の位置に関係なく敷き詰めます。", en: "Tiles the whole artboard, ignoring the original position." },
            lightMode: { ja: "プレビューを簡易表示にして、重いオブジェクトでも操作を軽くします。", en: "Simplifies the preview so heavy objects stay responsive." },
            zoom: { ja: "作業中の画面表示倍率を変えます。結果には影響しません。", en: "Changes the view zoom while you work. It does not affect the result." },
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
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            invalidCount: { ja: "繰り返し数は1以上の整数を入力してください。", en: "Enter an integer count of 1 or more." },
            invalidGap: { ja: "間隔は数値で入力してください。", en: "Enter a numeric gap value." }
        }
    };

    /**
     * ドット区切りのキーからUI言語のラベルを取得する
     * @param {string} labelPath - "panel.gap" のようなドット区切りのキー
     * @param {object} [params] - {unit: "mm"} のような差し込み値
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath, params) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        var localizedText = labelNode[uiLang] || labelNode.en;
        if (!localizedText) return labelPath;
        if (params) {
            for (var paramKey in params) {
                if (!params.hasOwnProperty(paramKey)) continue;
                localizedText = localizedText.replace(new RegExp("\\{" + paramKey + "\\}", "g"), params[paramKey]);
            }
        }
        return localizedText;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - "fieldLabel.countHorizontal" のようなドット区切りのキー
     * @returns {string} コロンを添えたラベル
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の要素間隔
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
     * 行グループの共通設定を適用する
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [alignment] - グループ自身の配置（既定は "left"）
     * @param {number} [spacing] - グループ内の要素間隔
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.alignment = alignment || "left";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Group|Window} parentContainer - 追加先
     * @param {string} panelTitle - パネルのタイトル
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 左寄せの縦並びグループを生成する（入力欄の列・チェックボックス列など）
     * @param {Group|Panel} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addColumnGroup(parentContainer) {
        var columnGroup = parentContainer.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["left", "center"];
        return columnGroup;
    }

    /**
     * ダイアログウィンドウを生成する
     * @param {string} title - タイトルバーの文字列
     * @returns {Window} 生成したダイアログ
     */
    function createDialogWindow(title) {
        var dialogWindow = new Window("dialog", title);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.spacing = WINDOW_SPACING;
        dialogWindow.margins = WINDOW_MARGINS;
        return dialogWindow;
    }

    /**
     * 数値入力欄の有効／無効を、左の∧∨ごと切り替える
     * @param {EditText} numericInput - addNumericField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setNumericFieldEnabled(numericInput, isEnabled) {
        numericInput.enabled = isEnabled;
        numericInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numericInput.stepperGroup); /* ∧∨は自作描画なので描き直す / redraw the custom-drawn stepper */
    }

    /**
     * ラジオボタンを選択し、クリック時と同じ処理を走らせる
     * @param {RadioButton} radioButton - 選択するラジオボタン
     * @returns {void}
     */
    function selectRadioButton(radioButton) {
        if (!radioButton || radioButton.enabled === false) return;
        radioButton.value = true;
        if (typeof radioButton.onClick === "function") radioButton.onClick();
    }

    /**
     * ラベルと tooltip が同じキーのラジオボタン／チェックボックスを追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} controlType - "radiobutton" または "checkbox"
     * @param {string} labelKey - LABELS.radio（または LABELS.checkbox）と LABELS.tooltip に共通のキー
     * @returns {RadioButton|Checkbox} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var labelCategory = (controlType === "radiobutton") ? "radio" : "checkbox";
        var labeledControl = parentContainer.add(controlType, undefined, getLabel(labelCategory + "." + labelKey));
        labeledControl.helpTip = getLabel("tooltip." + labelKey);
        return labeledControl;
    }

    /**
     * 項目名付きの数値入力欄を追加する（左の∧∨と上下キーで増減できる。増減後は onChanging を呼ぶ）
     * @param {Group} parentGroup - 追加先
     * @param {string} fieldKey - LABELS.fieldLabel と LABELS.tooltip に共通のキー
     * @param {string} initialText - 初期値
     * @param {boolean} isInteger - 整数だけにするなら true
     * @returns {EditText} 追加した入力欄
     */
    function addNumericField(parentGroup, fieldKey, initialText, isInteger) {
        var fieldGroup = parentGroup.add("group");
        fieldGroup.add("statictext", undefined, labelText("fieldLabel." + fieldKey));
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        /* 繰り返し数は1以上の整数、間隔はマイナスも可 / counts are integers of 1 or more; gaps may be negative */
        var stepOptions = {
            onStep: function (steppedInput) {
                if (typeof steppedInput.onChanging === "function") steppedInput.onChanging();
            }
        };
        if (isInteger) {
            stepOptions.integer = true;
            stepOptions.min = REPEAT_COUNT_MIN;
        }
        var numericInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numericInput; }, stepOptions);
        numericInput = stepperInputGroup.add("edittext", undefined, initialText);
        numericInput.helpTip = getLabel("tooltip." + fieldKey);
        numericInput.characters = FIELD_CHARS;
        numericInput.stepperGroup = stepperGroup;
        if (isInteger) numericInput.isInteger = true;
        bindSteppedArrowKeys(numericInput, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */
        return numericInput;
    }

    /**
     * 横／縦の数値入力欄と連動アイコンの組を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} horizontalKey - 横の入力欄のキー（LABELS.fieldLabel / LABELS.tooltip）
     * @param {string} verticalKey - 縦の入力欄のキー（LABELS.fieldLabel / LABELS.tooltip）
     * @param {string} linkTooltipKey - 連動アイコンの tooltip のキー
     * @param {string} initialText - 入力欄の初期値
     * @param {boolean} isInteger - 整数だけにするなら true
     * @returns {{horizontalInput: EditText, verticalInput: EditText, linkToggle: Group}} 追加したコントロール（切り替え時の処理は linkToggle.handleToggle に入れる）
     */
    function addLinkedFieldPair(parentPanel, horizontalKey, verticalKey, linkTooltipKey, initialText, isInteger) {
        var pairRow = parentPanel.add("group");
        setupRow(pairRow, "left", FIELD_ROW_SPACING);
        /* 2つの入力欄の右、上下中央にリンクアイコンを置く / Link icon to the right of the two fields, vertically centred */
        pairRow.alignChildren = ["left", "center"];

        var fieldsColumn = addColumnGroup(pairRow);
        var horizontalInput = addNumericField(fieldsColumn, horizontalKey, initialText, isInteger);
        var verticalInput = addNumericField(fieldsColumn, verticalKey, initialText, isInteger);

        var linkToggle = addLinkToggle(pairRow, true, function () {
            if (linkToggle.handleToggle) linkToggle.handleToggle();
        });
        linkToggle.helpTip = getLabel("tooltip." + linkTooltipKey);

        return { horizontalInput: horizontalInput, verticalInput: verticalInput, linkToggle: linkToggle };
    }

    /**
     * 項目名と2つのラジオボタンの行を追加する（方向の指定用）
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @param {string} firstKey - 1つ目のラジオのキー（LABELS.radio / LABELS.tooltip）
     * @param {string} secondKey - 2つ目のラジオのキー（LABELS.radio / LABELS.tooltip）
     * @returns {RadioButton[]} 追加したラジオボタン [1つ目, 2つ目]
     */
    function addDirectionRow(parentPanel, labelKey, firstKey, secondKey) {
        var directionRow = parentPanel.add("group");
        setupRow(directionRow);
        directionRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        var firstRadio = addLabeledControl(directionRow, "radiobutton", firstKey);
        var secondRadio = addLabeledControl(directionRow, "radiobutton", secondKey);
        return [firstRadio, secondRadio];
    }

    /**
     * 2カラムレイアウトの列を追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加した列
     */
    function addLayoutColumn(parentGroup) {
        var layoutColumn = parentGroup.add("group");
        layoutColumn.orientation = "column";
        layoutColumn.alignChildren = "fill";
        layoutColumn.spacing = WINDOW_SPACING;
        return layoutColumn;
    }

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

    // =========================================
    // 境界と複製 / Bounds and duplication
    // =========================================

    /**
     * クリッピングマスクを優先して境界を取得する（なければ可視境界）
     * @param {PageItem} targetItem - 対象オブジェクト
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getMaskedBounds(targetItem) {
        try {
            if (targetItem.typename === "GroupItem" && targetItem.clipped) {
                for (var i = 0; i < targetItem.pageItems.length; i++) {
                    var childItem = targetItem.pageItems[i];
                    if (childItem.typename === "PathItem" && childItem.clipping) return childItem.geometricBounds;
                    if (childItem.typename === "CompoundPathItem" && childItem.pathItems.length > 0 && childItem.pathItems[0].clipping)
                        return childItem.pathItems[0].geometricBounds;
                }
            }
            if (targetItem.typename === "PathItem" && targetItem.clipping) return targetItem.geometricBounds;
            if (targetItem.typename === "CompoundPathItem" && targetItem.pathItems.length > 0 && targetItem.pathItems[0].clipping)
                return targetItem.pathItems[0].geometricBounds;
        } catch (e) { }
        return targetItem.visibleBounds;
    }

    /**
     * アクティブなアートボードの矩形を取得する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getActiveArtboardRect(doc) {
        return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    }

    /**
     * 複数オブジェクトを囲む境界を求める
     * @param {Array<PageItem>} targetItems - 対象オブジェクトの配列
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getUnionBounds(targetItems) {
        var unionLeft = Infinity, unionTop = -Infinity, unionRight = -Infinity, unionBottom = Infinity;
        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds = getMaskedBounds(targetItems[i]);
            if (itemBounds[0] < unionLeft) unionLeft = itemBounds[0];
            if (itemBounds[1] > unionTop) unionTop = itemBounds[1];
            if (itemBounds[2] > unionRight) unionRight = itemBounds[2];
            if (itemBounds[3] < unionBottom) unionBottom = itemBounds[3];
        }
        return [unionLeft, unionTop, unionRight, unionBottom];
    }

    /**
     * グリッド配置のオフセット一覧を作る（元オブジェクトのぶんは含まない）
     * @param {number} rowCount - 縦の数
     * @param {number} columnCount - 横の数
     * @param {number} sourceWidth - 元オブジェクトの幅（pt）
     * @param {number} sourceHeight - 元オブジェクトの高さ（pt）
     * @param {number} gapX - 左右の間隔（pt）
     * @param {number} gapY - 上下の間隔（pt）
     * @param {string} horizontalDirection - "right" または "left"
     * @param {string} verticalDirection - "up" または "down"
     * @returns {Array<Array<number>>} [dx, dy] の配列（pt）
     */
    function computeGridOffsets(rowCount, columnCount, sourceWidth, sourceHeight, gapX, gapY, horizontalDirection, verticalDirection) {
        var gridOffsets = [];
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                if (rowIndex === 0 && columnIndex === 0) continue;
                var dx = (sourceWidth + gapX) * columnIndex;
                if (horizontalDirection === "left") dx = -dx;
                var dy = (sourceHeight + gapY) * rowIndex;
                if (verticalDirection !== "up") dy = -dy;
                gridOffsets.push([dx, dy]);
            }
        }
        return gridOffsets;
    }

    /**
     * 元オブジェクトを指定オフセットぶん複製する
     * @param {Array<PageItem>} sourceItems - 複製元（複数可。まとめて同じ量だけずらす）
     * @param {Array<Array<number>>} placementOffsets - [dx, dy] の配列（pt）
     * @param {Layer} [targetLayer] - 複製先レイヤー（省略時は元と同じ場所に複製）
     * @returns {Array<PageItem>} 生成した複製の配列
     */
    function duplicateWithOffsets(sourceItems, placementOffsets, targetLayer) {
        var duplicatedItems = [];

        for (var i = 0; i < placementOffsets.length; i++) {
            var dx = placementOffsets[i][0];
            var dy = placementOffsets[i][1];
            for (var j = 0; j < sourceItems.length; j++) {
                /* duplicate() は同じ座標に作られるので、オフセットぶん動かすだけでよい
                   duplicate() keeps the original coordinates, so shifting by the offset is enough */
                var duplicatedItem = targetLayer
                    ? sourceItems[j].duplicate(targetLayer, ElementPlacement.PLACEATBEGINNING)
                    : sourceItems[j].duplicate();
                if (targetLayer) duplicatedItem.note = PREVIEW_ITEM_TAG;

                duplicatedItem.left += dx;
                duplicatedItem.top += dy;
                duplicatedItems.push(duplicatedItem);
            }
        }
        return duplicatedItems;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビュー用レイヤーを取得する（なければ作成し、最前面へ）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} プレビュー用レイヤー
     */
    function getPreviewLayer(doc) {
        var previewLayer;
        try {
            previewLayer = doc.layers.getByName(PREVIEW_LAYER_NAME);
        } catch (e) {
            previewLayer = doc.layers.add();
            previewLayer.name = PREVIEW_LAYER_NAME;
        }
        previewLayer.visible = true;
        previewLayer.locked = false;
        previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        return previewLayer;
    }

    /**
     * プレビューで作った一時オブジェクトだけを削除する（noteタグで判別）
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} [skipRedraw] - true なら再描画を呼ばない（直後に描き直す場合）
     * @returns {void}
     */
    function clearPreview(doc, skipRedraw) {
        var previewLayer;
        try {
            previewLayer = doc.layers.getByName(PREVIEW_LAYER_NAME);
        } catch (e) {
            return;
        }
        /* pageItems はグループの子まで含むため、削除で添字がずれても止まらないよう1件ずつ受ける
           pageItems includes group children, so guard each removal against the shifting index */
        for (var i = previewLayer.pageItems.length - 1; i >= 0; i--) {
            try {
                var previewItem = previewLayer.pageItems[i];
                if (previewItem.note === PREVIEW_ITEM_TAG) previewItem.remove();
            } catch (e) { }
        }
        if (!skipRedraw) app.redraw();
    }

    /**
     * プレビューを描き直す
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} sourceItems - 複製元
     * @param {Array<Array<number>>} placementOffsets - [dx, dy] の配列（pt）
     * @returns {void}
     */
    function renderPreview(doc, sourceItems, placementOffsets) {
        var previewLayer = getPreviewLayer(doc);
        /* 空の状態を挟むとちらつくので、消去時は再描画しない / Skip the intermediate repaint so the canvas never flashes empty */
        clearPreview(doc, true);
        duplicateWithOffsets(sourceItems, placementOffsets, previewLayer);
        app.redraw();
    }

    // =========================================
    // 画面ズーム / View zoom
    // 軽量モードではスライダーを離したときだけ適用 / Light mode applies the zoom only on release
    // =========================================

    /**
     * 現在のビュー状態（ズーム倍率と中心）を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} {view, zoom, center} の状態オブジェクト
     */
    function captureViewState(doc) {
        var viewState = { view: null, zoom: null, center: null };
        /* ビューが取れないことがある / The view may be unavailable */
        try {
            viewState.view = doc.activeView;
            viewState.zoom = viewState.view.zoom;
            viewState.center = viewState.view.centerPoint;
        } catch (e) { }
        return viewState;
    }

    /**
     * 控えておいたビュー状態を復元する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} viewState - captureViewState() の戻り値
     * @returns {void}
     */
    function restoreViewState(doc, viewState) {
        if (!viewState) return;
        /* ビューが閉じられていることがある / The view may be gone */
        try {
            var targetView = viewState.view || doc.activeView;
            if (targetView && viewState.zoom != null) targetView.zoom = viewState.zoom;
            if (targetView && viewState.center != null) targetView.centerPoint = viewState.center;
        } catch (e) { }
    }

    /**
     * 画面ズーム用のスライダーと軽量モードのチェックボックスを追加する
     * @param {Group|Window} parentContainer - 追加先
     * @param {Document} doc - 対象ドキュメント
     * @param {object} initialViewState - captureViewState() の戻り値
     * @returns {{restoreInitial: Function}} 開いたときのビュー状態に戻す関数
     */
    function addZoomControls(parentContainer, doc, initialViewState) {
        var zoomGroup = parentContainer.add("group");
        zoomGroup.orientation = "row";
        zoomGroup.alignChildren = ["center", "center"];
        zoomGroup.alignment = "center";
        zoomGroup.margins = ZOOM_GROUP_MARGINS;

        zoomGroup.add("statictext", undefined, labelText("fieldLabel.zoom"));

        var initialZoom = 1;
        /* ビューが取れないことがある / The view may be unavailable */
        try {
            initialZoom = Number((initialViewState.zoom != null) ? initialViewState.zoom : doc.activeView.zoom);
        } catch (e) { }
        if (!initialZoom || isNaN(initialZoom)) initialZoom = 1;

        var zoomSlider = zoomGroup.add("slider", undefined, initialZoom, VIEW_ZOOM_MIN, VIEW_ZOOM_MAX);
        zoomSlider.helpTip = getLabel("tooltip.zoom");
        zoomSlider.preferredSize.width = ZOOM_SLIDER_WIDTH;

        var lightModeCheck = zoomGroup.add("checkbox", undefined, getLabel("checkbox.lightMode"));
        lightModeCheck.helpTip = getLabel("tooltip.lightMode");
        lightModeCheck.value = false;

        /**
         * 指定倍率をビューへ適用する
         * @param {number} zoomLevel - ズーム倍率
         * @returns {void}
         */
        function applyZoom(zoomLevel) {
            /* ビューが閉じられている・範囲外の倍率のときは何もしない / Ignore a missing view or an out-of-range zoom */
            try {
                var targetView = initialViewState.view ? initialViewState.view : doc.activeView;
                if (!targetView) return;
                targetView.zoom = zoomLevel;
                app.redraw();
            } catch (e) { }
        }

        /* ドラッグ中の追従（軽量モードでは無効）/ Live drag (disabled in light mode) */
        zoomSlider.onChanging = function () {
            if (lightModeCheck.value) return;
            applyZoom(Number(zoomSlider.value));
        };

        /* 離したときは必ず1回適用 / Always apply once on release */
        zoomSlider.onChange = function () {
            applyZoom(Number(zoomSlider.value));
        };

        lightModeCheck.onClick = function () {
            applyZoom(Number(zoomSlider.value));
        };

        return {
            restoreInitial: function () { restoreViewState(doc, initialViewState); }
        };
    }

    // =========================================
    // 繰り返し数とランダム配置 / Repeat counts and random placement
    // =========================================

    /**
     * ［アートボードの端まで］：選択オブジェクトを起点に、アートボードの端まで並ぶ行列数を求める
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {number[]} sourceBounds - 複製元の境界 [左, 上, 右, 下]
     * @param {number} sourceWidth - 複製元の幅（pt）
     * @param {number} sourceHeight - 複製元の高さ（pt）
     * @param {{x: number, y: number}} gapPoints - 左右・上下の間隔（pt）
     * @param {string} repeatMethod - "grid" / "row" / "column" / "random"
     * @param {boolean} isRightward - 右方向に並べるなら true（false なら左）
     * @param {boolean} isUpward - 上方向に並べるなら true（false なら下）
     * @returns {{columns: number, rows: number}} 横と縦の数
     */
    function computeCountsToArtboardEdge(artboardRect, sourceBounds, sourceWidth, sourceHeight, gapPoints, repeatMethod, isRightward, isUpward) {
        var sourceLeft = sourceBounds[0], sourceTop = sourceBounds[1];
        var stepWidth = sourceWidth + gapPoints.x, stepHeight = sourceHeight + gapPoints.y;
        var columnCount = 1, rowCount = 1;

        if (repeatMethod !== "column" && stepWidth > 0) {
            var availableWidth = isRightward ? (artboardRect[2] - sourceLeft) : (sourceLeft - artboardRect[0]);
            columnCount = Math.max(1, Math.floor((availableWidth + gapPoints.x) / stepWidth));
        }
        if (repeatMethod !== "row" && stepHeight > 0) {
            var availableHeight = isUpward ? (artboardRect[1] - sourceTop) : (sourceTop - artboardRect[3]);
            rowCount = Math.max(1, Math.floor((availableHeight + gapPoints.y) / stepHeight));
        }
        return { columns: columnCount, rows: rowCount };
    }

    /**
     * ［アートボードいっぱいに］：アートボードに収まる最大の行列数を求める（方向は無視）
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {number} sourceWidth - 複製元の幅（pt）
     * @param {number} sourceHeight - 複製元の高さ（pt）
     * @param {{x: number, y: number}} gapPoints - 左右・上下の間隔（pt）
     * @param {string} repeatMethod - "grid" / "row" / "column" / "random"
     * @returns {{columns: number, rows: number}} 横と縦の数
     */
    function computeCountsToFullArtboard(artboardRect, sourceWidth, sourceHeight, gapPoints, repeatMethod) {
        var artboardWidth = Math.abs(artboardRect[2] - artboardRect[0]);
        var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);
        var stepWidth = sourceWidth + gapPoints.x, stepHeight = sourceHeight + gapPoints.y;

        return {
            columns: (repeatMethod === "column" || stepWidth <= 0)
                ? 1 : Math.max(1, Math.floor((artboardWidth + gapPoints.x) / stepWidth)),
            rows: (repeatMethod === "row" || stepHeight <= 0)
                ? 1 : Math.max(1, Math.floor((artboardHeight + gapPoints.y) / stepHeight))
        };
    }

    /**
     * 数値を繰り返し数の範囲に収める
     * @param {string|number} countValue - 入力値
     * @returns {number} REPEAT_COUNT_MIN〜REPEAT_COUNT_MAX に収めた整数
     */
    function clampRepeatCount(countValue) {
        /* スライダーは実数を返すので、切り捨てず四捨五入する / Sliders report real numbers, so round instead of truncating */
        var repeatCount = Math.round(Number(countValue));
        if (isNaN(repeatCount) || repeatCount < REPEAT_COUNT_MIN) return REPEAT_COUNT_MIN;
        if (repeatCount > REPEAT_COUNT_MAX) return REPEAT_COUNT_MAX;
        return repeatCount;
    }

    /**
     * 線形合同法で擬似乱数を返す（同じ種から同じ並びを再現するため）
     * @param {object} seedHolder - {v: number} 形式の内部状態
     * @returns {number} 0以上1未満の擬似乱数
     */
    function nextRandomValue(seedHolder) {
        seedHolder.v = (seedHolder.v * 1664525 + 1013904223) % 4294967296;
        return seedHolder.v / 4294967296;
    }

    /**
     * ランダム配置のオフセット一覧を作る
     * @param {number} duplicateCount - 複製する数
     * @param {number} rangeX - 左右の散らばり範囲（pt）
     * @param {number} rangeY - 上下の散らばり範囲（pt）
     * @param {number} randomSeed - 乱数の種
     * @returns {Array<Array<number>>} [dx, dy] の配列（pt）
     */
    function createRandomOffsets(duplicateCount, rangeX, rangeY, randomSeed) {
        var seedHolder = { v: randomSeed >>> 0 };
        var randomOffsets = [];
        for (var i = 0; i < duplicateCount; i++) {
            var dx = (rangeX > 0) ? (-rangeX + 2 * rangeX * nextRandomValue(seedHolder)) : 0;
            var dy = (rangeY > 0) ? (-rangeY + 2 * rangeY * nextRandomValue(seedHolder)) : 0;
            randomOffsets.push([dx, dy]);
        }
        return randomOffsets;
    }

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

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを組み立てる（イベントは showDuplicateDialog() で結線する）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} rulerUnitLabel - 定規の単位の表示名
     * @returns {object} ダイアログ・各コントロール・ズームの操作をまとめたオブジェクト
     */
    function buildDuplicateDialog(doc, rulerUnitLabel) {
        var duplicateDialog = createDialogWindow(getLabel("dialog.title") + " " + SCRIPT_VERSION);

        /* 2カラムレイアウト：左（繰り返し数／方式）、右（間隔／方向／敷き詰め）
           Two-column layout: counts & method on the left, gap / direction / fill on the right */
        var columnsGroup = duplicateDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = addLayoutColumn(columnsGroup);
        var rightColumnGroup = addLayoutColumn(columnsGroup);

        /* 繰り返し数 / Repeat count */
        var repeatCountPanel = addPanel(leftColumnGroup, getLabel("panel.repeatCount"));
        var countFields = addLinkedFieldPair(repeatCountPanel, "countHorizontal", "countVertical", "countLink", "2", true);

        var countSliderGroup = repeatCountPanel.add("group");
        countSliderGroup.orientation = "row";
        countSliderGroup.alignChildren = ["fill", "center"];
        var countSlider = countSliderGroup.add("slider", undefined, 2, REPEAT_COUNT_MIN, REPEAT_COUNT_MAX);
        countSlider.helpTip = getLabel("tooltip.countSlider");
        countSlider.alignment = ["fill", "center"];

        /* 繰り返し方式 / Repeat method */
        var repeatMethodPanel = addPanel(leftColumnGroup, getLabel("panel.repeatMethod"));
        repeatMethodPanel.alignChildren = ["left", "top"];
        var methodGridRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodGrid");
        var methodRowRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodRow");
        var methodColumnRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodColumn");
        var methodRandomRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodRandom");
        methodGridRadio.value = true;

        /* 間隔（現在の定規単位で入力し、内部ではptへ変換）
           Gap (entered in the current ruler unit, converted to points internally) */
        var gapPanel = addPanel(rightColumnGroup, getLabel("panel.gap", { unit: rulerUnitLabel }));
        var gapFields = addLinkedFieldPair(gapPanel, "gapHorizontal", "gapVertical", "gapLink", "10", false);

        /* 方向 / Direction */
        var directionPanel = addPanel(rightColumnGroup, getLabel("panel.direction"));
        directionPanel.alignChildren = ["left", "top"];
        var horizontalDirectionRadios = addDirectionRow(directionPanel, "directionHorizontal", "directionRight", "directionLeft");
        horizontalDirectionRadios[0].value = true;
        var verticalDirectionRadios = addDirectionRow(directionPanel, "directionVertical", "directionUp", "directionDown");
        verticalDirectionRadios[1].value = true;

        /* 敷き詰め / Fill */
        var fillPanel = addPanel(rightColumnGroup, getLabel("panel.fill"));
        fillPanel.alignChildren = ["left", "top"];
        var fillToEdgeCheck = addLabeledControl(fillPanel, "checkbox", "fillToEdge");
        fillToEdgeCheck.value = false;
        var fillFullCheck = addLabeledControl(fillPanel, "checkbox", "fillFull");
        fillFullCheck.value = false;

        /* 画面ズーム / Zoom */
        var zoomControls = addZoomControls(duplicateDialog, doc, captureViewState(doc));

        var btnRowGroup = duplicateDialog.add("group");
        setupRow(btnRowGroup, "center");
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));

        return {
            dialog: duplicateDialog,
            countHorizontalInput: countFields.horizontalInput,
            countVerticalInput: countFields.verticalInput,
            countLinkToggle: countFields.linkToggle,
            countSlider: countSlider,
            methodGridRadio: methodGridRadio,
            methodRowRadio: methodRowRadio,
            methodColumnRadio: methodColumnRadio,
            methodRandomRadio: methodRandomRadio,
            gapHorizontalInput: gapFields.horizontalInput,
            gapVerticalInput: gapFields.verticalInput,
            gapLinkToggle: gapFields.linkToggle,
            directionRightRadio: horizontalDirectionRadios[0],
            directionLeftRadio: horizontalDirectionRadios[1],
            directionUpRadio: verticalDirectionRadios[0],
            directionDownRadio: verticalDirectionRadios[1],
            fillToEdgeCheck: fillToEdgeCheck,
            fillFullCheck: fillFullCheck,
            zoomControls: zoomControls,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 設定ダイアログを表示し、プレビューを結線する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} sourceItems - 複製元オブジェクト（複数選択のまま扱う）
     * @param {number} sourceWidth - 複製元全体の幅（pt）
     * @param {number} sourceHeight - 複製元全体の高さ（pt）
     * @returns {object} OKなら {placementOffsets, fillFullArtboard}、キャンセルなら null
     */
    function showDuplicateDialog(doc, sourceItems, sourceWidth, sourceHeight) {
        var rulerUnit = getUnitInfo("rulerType");
        var dialogControls = buildDuplicateDialog(doc, rulerUnit.label);

        var duplicateDialog = dialogControls.dialog;
        var countHorizontalInput = dialogControls.countHorizontalInput;
        var countVerticalInput = dialogControls.countVerticalInput;
        var countLinkToggle = dialogControls.countLinkToggle;
        var countSlider = dialogControls.countSlider;
        var methodGridRadio = dialogControls.methodGridRadio;
        var methodRowRadio = dialogControls.methodRowRadio;
        var methodColumnRadio = dialogControls.methodColumnRadio;
        var methodRandomRadio = dialogControls.methodRandomRadio;
        var gapHorizontalInput = dialogControls.gapHorizontalInput;
        var gapVerticalInput = dialogControls.gapVerticalInput;
        var gapLinkToggle = dialogControls.gapLinkToggle;
        var directionRightRadio = dialogControls.directionRightRadio;
        var directionLeftRadio = dialogControls.directionLeftRadio;
        var directionUpRadio = dialogControls.directionUpRadio;
        var directionDownRadio = dialogControls.directionDownRadio;
        var fillToEdgeCheck = dialogControls.fillToEdgeCheck;
        var fillFullCheck = dialogControls.fillFullCheck;

        // -----------------------------------------
        // 入力値の読み取り / Reading the input values
        // -----------------------------------------

        /**
         * 選択中の繰り返し方式を返す
         * @returns {string} "grid" / "row" / "column" / "random"
         */
        function getRepeatMethod() {
            if (methodRowRadio.value) return "row";
            if (methodColumnRadio.value) return "column";
            if (methodRandomRadio.value) return "random";
            return "grid";
        }

        /**
         * 繰り返し数を読み取る（方式に応じて不要側を1に固定）
         * @returns {object} {columns, rows}。数値として読めない場合は null
         */
        function readRepeatCounts() {
            var columnCount = parseInt(countHorizontalInput.text, 10);
            var rowCount = parseInt(countVerticalInput.text, 10);
            var repeatMethod = getRepeatMethod();
            if (repeatMethod === "row" || repeatMethod === "random") rowCount = 1;
            if (repeatMethod === "column") columnCount = 1;
            if (isNaN(columnCount) || columnCount < 1 || isNaN(rowCount) || rowCount < 1) return null;
            return { columns: columnCount, rows: rowCount };
        }

        /**
         * 間隔を読み取り、ポイントへ変換する
         * @returns {object} {x, y}（pt）。数値として読めない場合は null
         */
        function readGapPoints() {
            var horizontalGap = parseFloat(gapHorizontalInput.text);
            var verticalGap = parseFloat(gapVerticalInput.text);
            if (isNaN(horizontalGap) || isNaN(verticalGap)) return null;
            return {
                x: horizontalGap * rulerUnit.pointsPerUnit,
                y: verticalGap * rulerUnit.pointsPerUnit
            };
        }

        // -----------------------------------------
        // ランダム配置 / Random placement
        // -----------------------------------------

        /* OK後もプレビューと同じ配置にするためのキャッシュ / Cache so OK keeps the previewed layout */
        var randomOffsetCache = { key: null, offsets: [] };

        /**
         * ランダム配置のオフセットを取得する（同じ条件ならキャッシュを返す）
         * @param {number} repeatCount - 繰り返し数
         * @param {number} gapX - 左右の間隔（pt）
         * @param {number} gapY - 上下の間隔（pt）
         * @returns {Array<Array<number>>} [dx, dy] の配列（pt）
         */
        function getRandomOffsets(repeatCount, gapX, gapY) {
            var GAP_EPSILON = 1e-9;
            var duplicateCount = Math.max(0, repeatCount - 1);

            /* 間隔が0の軸はランダムを完全にOFF / A gap of 0 disables randomness on that axis */
            var rangeX = (Math.abs(gapX) <= GAP_EPSILON) ? 0 : Math.max(0, (sourceWidth + gapX) * (repeatCount - 1) / 2);
            var rangeY = (Math.abs(gapY) <= GAP_EPSILON) ? 0 : Math.max(0, (sourceHeight + gapY) * (repeatCount - 1) / 2);

            var cacheKey = [repeatCount, gapX.toFixed(4), gapY.toFixed(4), sourceWidth.toFixed(4), sourceHeight.toFixed(4)].join("|");
            if (randomOffsetCache.key === cacheKey && randomOffsetCache.offsets.length === duplicateCount) {
                return randomOffsetCache.offsets;
            }

            var randomSeed = (new Date().getTime() & 0xFFFFFFFF) ^ (repeatCount << 16);
            randomOffsetCache.key = cacheKey;
            randomOffsetCache.offsets = createRandomOffsets(duplicateCount, rangeX, rangeY, randomSeed);
            return randomOffsetCache.offsets;
        }

        /**
         * 現在の設定から配置オフセットを組み立てる
         * @returns {Array<Array<number>>} [dx, dy] の配列（pt）。入力値が読めない場合は null
         */
        function buildPlacementOffsets() {
            var repeatCounts = readRepeatCounts();
            var gapPoints = readGapPoints();
            if (!repeatCounts || !gapPoints) return null;

            if (getRepeatMethod() === "random") {
                return getRandomOffsets(repeatCounts.columns, gapPoints.x, gapPoints.y);
            }
            var horizontalDirection = directionRightRadio.value ? "right" : "left";
            var verticalDirection = directionUpRadio.value ? "up" : "down";
            return computeGridOffsets(
                repeatCounts.rows, repeatCounts.columns,
                sourceWidth, sourceHeight,
                gapPoints.x, gapPoints.y,
                horizontalDirection, verticalDirection
            );
        }

        // -----------------------------------------
        // プレビュー更新 / Preview updates
        // -----------------------------------------

        var lastPreviewTime = 0;

        /**
         * 現在の設定でプレビューを描き直す
         * @returns {void}
         */
        function applyPreview() {
            var placementOffsets = buildPlacementOffsets();
            if (!placementOffsets) return;
            renderPreview(doc, sourceItems, placementOffsets);
        }

        /**
         * プレビュー更新を間引く（スライダードラッグ中のチラつき軽減）
         * @param {boolean} force - true なら間引かずに更新する
         * @returns {void}
         */
        function applyPreviewThrottled(force) {
            var currentTime = new Date().getTime();
            if (!force && (currentTime - lastPreviewTime) < PREVIEW_THROTTLE_MS) return;
            lastPreviewTime = currentTime;
            applyPreview();
        }

        // -----------------------------------------
        // UIの同期 / Keeping the UI in sync
        // -----------------------------------------

        /**
         * ［連動］に合わせて縦の繰り返し数を横へそろえる
         * @returns {void}
         */
        function syncCountFields() {
            if (countLinkToggle.value) {
                setNumericFieldEnabled(countVerticalInput, false);
                countVerticalInput.text = countHorizontalInput.text;
            } else {
                setNumericFieldEnabled(countVerticalInput, true);
            }
            updateCountSliderFromFields();
        }

        /**
         * ［連動］に合わせて上下の間隔を左右へそろえる
         * @returns {void}
         */
        function syncGapFields() {
            if (gapLinkToggle.value) {
                setNumericFieldEnabled(gapVerticalInput, false);
                gapVerticalInput.text = gapHorizontalInput.text;
            } else {
                setNumericFieldEnabled(gapVerticalInput, true);
            }
        }

        /**
         * 方向ラジオの有効／無効をまとめて切り替える
         * @param {boolean} horizontalEnabled - 横方向を有効にするか
         * @param {boolean} verticalEnabled - 縦方向を有効にするか
         * @returns {void}
         */
        function setDirectionEnabled(horizontalEnabled, verticalEnabled) {
            directionRightRadio.enabled = horizontalEnabled;
            directionLeftRadio.enabled = horizontalEnabled;
            directionUpRadio.enabled = verticalEnabled;
            directionDownRadio.enabled = verticalEnabled;
        }

        /**
         * 繰り返し方式と敷き詰めの状態から、方向ラジオの有効／無効を決める
         * @returns {void}
         */
        function applyDirectionStateForMethod() {
            /* ［アートボードいっぱいに］のあいだは方向を固定 / Fill Full pins the direction */
            if (fillFullCheck.value) { setDirectionEnabled(false, false); return; }

            var repeatMethod = getRepeatMethod();
            if (repeatMethod === "row") setDirectionEnabled(true, false);
            else if (repeatMethod === "column") setDirectionEnabled(false, true);
            else if (repeatMethod === "random") setDirectionEnabled(false, false);
            else setDirectionEnabled(true, true);
        }

        /**
         * 方向を既定（右・下）に戻す
         * @returns {void}
         */
        function resetDirectionToDefault() {
            directionRightRadio.value = true;
            directionLeftRadio.value = false;
            directionDownRadio.value = true;
            directionUpRadio.value = false;
        }

        /**
         * 入力欄の値をスライダーへ反映し、スライダーの有効／無効を決める
         * @returns {void}
         */
        function updateCountSliderFromFields() {
            var repeatMethod = getRepeatMethod();
            var sourceField = (repeatMethod === "column") ? countVerticalInput : countHorizontalInput;
            countSlider.value = clampRepeatCount(sourceField.text);

            /* 敷き詰め中は行列数を自動計算するので操作させない（上限20で潰れるため）
               While a Fill option drives the counts, block the slider so it cannot clamp them to 20 */
            if (fillToEdgeCheck.value || fillFullCheck.value) {
                countSlider.enabled = false;
                return;
            }
            /* 行／列／ランダムでは常に有効、グリッドは［連動］ONのときだけ有効
               Enabled in Row / Column / Random, and in Grid only while Link is on */
            countSlider.enabled = (repeatMethod !== "grid") || (countLinkToggle.enabled && countLinkToggle.value);
        }

        /**
         * スライダーの値を繰り返し数の入力欄へ反映する
         * @returns {void}
         */
        function applyCountFieldsFromSlider() {
            var repeatCount = clampRepeatCount(countSlider.value);
            countSlider.value = repeatCount;

            var repeatMethod = getRepeatMethod();
            if (repeatMethod === "column") {
                /* 列：横は常に1で、スライダーは縦を操作 / Column: horizontal stays 1, the slider drives vertical */
                countHorizontalInput.text = "1";
                countVerticalInput.text = String(repeatCount);
                return;
            }
            countHorizontalInput.text = String(repeatCount);
            if (repeatMethod === "row" || repeatMethod === "random") {
                countVerticalInput.text = "1";
            } else if (countLinkToggle.value) {
                countVerticalInput.text = String(repeatCount);
            }
        }

        /**
         * 繰り返し数の片方を1に固定し、もう片方だけを入力できるようにする（［連動］は OFF で無効）
         * @param {EditText} fixedInput - 1に固定する入力欄
         * @param {EditText} activeInput - 入力できるようにする入力欄
         * @returns {void}
         */
        function fixCountFieldToOne(fixedInput, activeInput) {
            fixedInput.text = "1";
            setNumericFieldEnabled(fixedInput, false);
            setNumericFieldEnabled(activeInput, true);
            setLinkToggleValue(countLinkToggle, false);
            setLinkToggleEnabled(countLinkToggle, false);
        }

        /**
         * 間隔の［連動］の値と有効／無効をまとめて設定する
         * @param {boolean} isLinked - ON かつ有効にするなら true、OFF かつ無効にするなら false
         * @returns {void}
         */
        function setGapLinkState(isLinked) {
            setLinkToggleValue(gapLinkToggle, isLinked);
            setLinkToggleEnabled(gapLinkToggle, isLinked);
        }

        /**
         * 繰り返し方式に合わせて各コントロールの状態を更新する
         * @returns {void}
         */
        function updateRepeatMethodUI() {
            var repeatMethod = getRepeatMethod();

            if (repeatMethod === "column") {
                /* 列：横は常に1 / Column: horizontal fixed to 1 */
                fixCountFieldToOne(countHorizontalInput, countVerticalInput);
                setGapLinkState(false);
                setNumericFieldEnabled(gapHorizontalInput, false);
                setNumericFieldEnabled(gapVerticalInput, true);

            } else if (repeatMethod === "row") {
                /* 行：縦は常に1 / Row: vertical fixed to 1 */
                fixCountFieldToOne(countVerticalInput, countHorizontalInput);
                setGapLinkState(false);
                setNumericFieldEnabled(gapHorizontalInput, true);
                setNumericFieldEnabled(gapVerticalInput, false);

            } else if (repeatMethod === "random") {
                /* ランダム：繰り返し数は横だけ、間隔は連動ON、方向と敷き詰めは無効
                   Random: a single count, gaps linked, direction & fill turned off */
                fixCountFieldToOne(countVerticalInput, countHorizontalInput);
                setGapLinkState(true);
                setNumericFieldEnabled(gapHorizontalInput, true);
                syncGapFields();
                fillToEdgeCheck.value = false;
                fillFullCheck.value = false;

            } else {
                /* グリッド：繰り返し数・間隔とも［連動］ONに戻す / Grid: restore both link toggles */
                setNumericFieldEnabled(countHorizontalInput, true);
                setLinkToggleEnabled(countLinkToggle, true);
                setLinkToggleValue(countLinkToggle, true);
                syncCountFields();

                setGapLinkState(true);
                setNumericFieldEnabled(gapHorizontalInput, true);
                syncGapFields();
            }

            applyDirectionStateForMethod();
            updateCountSliderFromFields();
        }

        // -----------------------------------------
        // 敷き詰めの自動計算 / Automatic fill counts
        // -----------------------------------------

        /**
         * 求めた行列数を入力欄とスライダーへ反映する
         * @param {{columns: number, rows: number}} fillCounts - 横と縦の数
         * @returns {void}
         */
        function setCountFields(fillCounts) {
            countHorizontalInput.text = String(fillCounts.columns);
            countVerticalInput.text = String(fillCounts.rows);
            updateCountSliderFromFields();
        }

        /**
         * ［アートボードの端まで］：選択オブジェクトを起点に行列数を求めて反映する
         * @returns {void}
         */
        function recalcCountsToArtboardEdge() {
            var gapPoints = readGapPoints();
            if (!gapPoints) return;
            setCountFields(computeCountsToArtboardEdge(
                getActiveArtboardRect(doc), getUnionBounds(sourceItems),
                sourceWidth, sourceHeight, gapPoints,
                getRepeatMethod(), directionRightRadio.value, directionUpRadio.value
            ));
        }

        /**
         * ［アートボードいっぱいに］：アートボードに収まる最大の行列数を求めて反映する（方向は無視）
         * @returns {void}
         */
        function recalcCountsToFullArtboard() {
            var gapPoints = readGapPoints();
            if (!gapPoints) return;
            setCountFields(computeCountsToFullArtboard(
                getActiveArtboardRect(doc), sourceWidth, sourceHeight, gapPoints, getRepeatMethod()
            ));
        }

        /**
         * ONになっている敷き詰めオプションに応じて行列数を再計算する
         * @returns {void}
         */
        function recalcFillCounts() {
            if (fillToEdgeCheck.value) recalcCountsToArtboardEdge();
            else if (fillFullCheck.value) recalcCountsToFullArtboard();
        }

        // -----------------------------------------
        // イベント結線 / Event wiring
        // -----------------------------------------

        /**
         * 横の繰り返し数が変わったときの処理
         * @returns {void}
         */
        function onCountHorizontalChanged() {
            if (countLinkToggle.value) countVerticalInput.text = countHorizontalInput.text;
            updateCountSliderFromFields();
            applyPreview();
        }

        /**
         * 縦の繰り返し数が変わったときの処理
         * @returns {void}
         */
        function onCountVerticalChanged() {
            updateCountSliderFromFields();
            applyPreview();
        }

        /**
         * 左右の間隔が変わったときの処理
         * @returns {void}
         */
        function onGapHorizontalChanged() {
            if (gapLinkToggle.value) gapVerticalInput.text = gapHorizontalInput.text;
            recalcFillCounts();
            applyPreview();
        }

        /**
         * 上下の間隔が変わったときの処理
         * @returns {void}
         */
        function onGapVerticalChanged() {
            recalcFillCounts();
            applyPreview();
        }

        /**
         * 繰り返し数のスライダーが動いたときの処理
         * @param {boolean} isReleased - 離したとき（間引かずにプレビューする）なら true
         * @returns {void}
         */
        function onCountSliderMoved(isReleased) {
            if (!countSlider.enabled) return;
            applyCountFieldsFromSlider();
            applyPreviewThrottled(isReleased);
        }

        /**
         * 左／上の方向が選ばれたときの処理（［アートボードの端まで］は右・下が起点なので自動で OFF）
         * @returns {void}
         */
        function onReverseDirectionChosen() {
            fillToEdgeCheck.value = false;
            updateCountSliderFromFields();
            applyPreview();
        }

        /**
         * 繰り返し方式のラジオが選ばれたときの処理
         * @returns {void}
         */
        function onRepeatMethodChanged() {
            updateRepeatMethodUI();
            recalcFillCounts();
            applyPreview();
        }

        countHorizontalInput.onChanging = onCountHorizontalChanged;
        countHorizontalInput.onChange = onCountHorizontalChanged;
        countVerticalInput.onChanging = onCountVerticalChanged;
        countVerticalInput.onChange = onCountVerticalChanged;
        gapHorizontalInput.onChanging = onGapHorizontalChanged;
        gapHorizontalInput.onChange = onGapHorizontalChanged;
        gapVerticalInput.onChanging = onGapVerticalChanged;
        gapVerticalInput.onChange = onGapVerticalChanged;

        countLinkToggle.handleToggle = function () {
            syncCountFields();
            applyPreview();
        };
        gapLinkToggle.handleToggle = function () {
            syncGapFields();
            /* 上下の間隔が左右にそろうと敷き詰めの行列数も変わる / Linking the gaps changes the fill counts too */
            recalcFillCounts();
            applyPreview();
        };

        countSlider.onChanging = function () {
            onCountSliderMoved(false);
        };
        countSlider.onChange = function () {
            onCountSliderMoved(true);
        };

        methodGridRadio.onClick = onRepeatMethodChanged;
        methodRowRadio.onClick = function () {
            /* 列→行では縦の数を横の数として引き継ぐ / Carry the vertical count over when switching Column -> Row */
            countHorizontalInput.text = String(clampRepeatCount(countVerticalInput.text));
            onRepeatMethodChanged();
        };
        methodColumnRadio.onClick = function () {
            /* 行→列では横の数を縦の数として引き継ぐ / Carry the horizontal count over when switching Row -> Column */
            countVerticalInput.text = String(clampRepeatCount(countHorizontalInput.text));
            onRepeatMethodChanged();
        };
        methodRandomRadio.onClick = function () {
            /* 列→ランダムでは縦の数を横の数として引き継ぐ / Carry the vertical count over when switching Column -> Random */
            if (getRepeatMethod() === "random" && countHorizontalInput.text === "1") {
                countHorizontalInput.text = String(clampRepeatCount(countVerticalInput.text));
            }
            onRepeatMethodChanged();
        };

        directionRightRadio.onClick = applyPreview;
        directionDownRadio.onClick = applyPreview;
        /* 左・上を選んだら［アートボードの端まで］を自動OFF / Choosing Left or Up turns Fill to Edge off */
        directionLeftRadio.onClick = onReverseDirectionChosen;
        directionUpRadio.onClick = onReverseDirectionChosen;

        fillToEdgeCheck.onClick = function () {
            if (fillToEdgeCheck.value) {
                /* 端までは右・下を起点にする（方向は変更可）/ Fill to Edge starts from the right & bottom, direction still editable */
                fillFullCheck.value = false;
                resetDirectionToDefault();
                applyDirectionStateForMethod();
                setLinkToggleValue(countLinkToggle, false);
                syncCountFields();
                recalcCountsToArtboardEdge();
            }
            updateCountSliderFromFields();
            applyPreview();
        };

        fillFullCheck.onClick = function () {
            if (fillFullCheck.value) {
                /* いっぱいには方向を固定してディム / Fill Full fixes the direction and dims the radios */
                fillToEdgeCheck.value = false;
                resetDirectionToDefault();
                recalcCountsToFullArtboard();
            }
            applyDirectionStateForMethod();
            updateCountSliderFromFields();
            applyPreview();
        };

        /* キー操作でラジオを選択 / Keyboard shortcuts for radios */
        duplicateDialog.addEventListener("keydown", function (event) {
            var keyName = event.keyName;
            if (!keyName) return;

            var targetRadio = null;
            if (keyName === "G") targetRadio = methodGridRadio;
            else if (keyName === "C") targetRadio = methodColumnRadio;
            else if (keyName === "A") targetRadio = methodRandomRadio;
            else if (keyName === "R") targetRadio = event.shiftKey ? directionRightRadio : methodRowRadio;
            else if (keyName === "L") targetRadio = directionLeftRadio;
            else if (keyName === "T") targetRadio = directionUpRadio;
            else if (keyName === "B") targetRadio = directionDownRadio;
            if (!targetRadio) return;

            selectRadioButton(targetRadio);
            event.preventDefault();
        });

        var dialogResult = null;

        dialogControls.btnOK.onClick = function () {
            if (!readRepeatCounts()) { alert(getLabel("alert.invalidCount")); return; }
            if (!readGapPoints()) { alert(getLabel("alert.invalidGap")); return; }

            dialogResult = {
                placementOffsets: buildPlacementOffsets(),
                fillFullArtboard: fillFullCheck.value
            };
            clearPreview(doc);
            duplicateDialog.close();
        };

        dialogControls.btnCancel.onClick = function () {
            dialogControls.zoomControls.restoreInitial();
            clearPreview(doc);
            duplicateDialog.close();
        };

        duplicateDialog.onClose = function () {
            /* OK以外で閉じたときはキャンセル扱い / Closing without OK is treated as Cancel */
            if (dialogResult) return;
            dialogControls.zoomControls.restoreInitial();
            clearPreview(doc);
        };

        duplicateDialog.onShow = function () {
            var dialogLocation = duplicateDialog.location;
            duplicateDialog.location = [dialogLocation[0] + DIALOG_OFFSET_X, dialogLocation[1]];
            applyPreview();
            countHorizontalInput.active = true;
        };

        updateRepeatMethodUI();
        prepareDialogWindow(duplicateDialog, SCRIPT_NAME);
        duplicateDialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 複数選択をひとつのグループにまとめる
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} selectedItems - 選択中のオブジェクト
     * @returns {PageItem} まとめたグループ（失敗時は先頭のオブジェクト）
     */
    function groupSelectedItems(doc, selectedItems) {
        try {
            var wrapperGroup = doc.groupItems.add();
            /* 選択配列はライブで変化するのでコピーを回す / The live selection array can change, so iterate a copy */
            var itemsToMove = [];
            for (var i = 0; i < selectedItems.length; i++) itemsToMove.push(selectedItems[i]);
            /* 移動できないオブジェクトが混じっても残りは処理する / Keep going even if one item refuses to move */
            for (var j = 0; j < itemsToMove.length; j++) {
                try { itemsToMove[j].move(wrapperGroup, ElementPlacement.PLACEATEND); } catch (err) { }
            }

            doc.selection = null;
            wrapperGroup.selected = true;
            return wrapperGroup;
        } catch (e) {
            return selectedItems[0];
        }
    }

    /**
     * 元オブジェクトと複製をまとめてアートボード中央へ移動する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} sourceItems - 複製元
     * @param {Array<PageItem>} duplicatedItems - 生成した複製
     * @returns {void}
     */
    function centerOnArtboard(doc, sourceItems, duplicatedItems) {
        var artboardRect = getActiveArtboardRect(doc);
        var itemsToMove = sourceItems.concat(duplicatedItems);
        var unionBounds = getUnionBounds(itemsToMove);

        var dx = (artboardRect[0] + artboardRect[2]) / 2 - (unionBounds[0] + unionBounds[2]) / 2;
        var dy = (artboardRect[1] + artboardRect[3]) / 2 - (unionBounds[1] + unionBounds[3]) / 2;

        for (var i = 0; i < itemsToMove.length; i++) {
            itemsToMove[i].left += dx;
            itemsToMove[i].top += dy;
        }
    }

    /**
     * 検証 → ダイアログ → 複製
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
        var doc = app.activeDocument;
        if (doc.selection.length === 0) { alert(getLabel("alert.noSelection")); return; }

        /* 選択配列はライブで変化するのでコピーを保持 / The live selection array can change, so keep a copy */
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) selectedItems.push(doc.selection[i]);

        var sourceBounds = getUnionBounds(selectedItems);
        var sourceWidth = sourceBounds[2] - sourceBounds[0];
        var sourceHeight = sourceBounds[1] - sourceBounds[3];

        /* グループ化はOK後まで遅らせる（キャンセル時にドキュメントを変更しないため）
           Grouping is deferred until after OK so that cancelling leaves the document untouched */
        var duplicateSettings = showDuplicateDialog(doc, selectedItems, sourceWidth, sourceHeight);
        if (!duplicateSettings) return;

        var sourceItems = (selectedItems.length > 1) ? [groupSelectedItems(doc, selectedItems)] : selectedItems;
        var duplicatedItems = duplicateWithOffsets(sourceItems, duplicateSettings.placementOffsets, null);
        if (duplicateSettings.fillFullArtboard) centerOnArtboard(doc, sourceItems, duplicatedItems);
    }

    main();

})();
