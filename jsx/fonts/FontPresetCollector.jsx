#target illustrator
#targetengine "FontPresetCollectorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなドキュメントで使われている書式（フォント・サイズ・行送り・字間・行揃え・段落前後のアキ）を段落ごとに調べて一覧にし、
一覧で選んだ書式を編集・置換すると、その書式の段落をまとめて更新します（段落スタイルを使わない再定義）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetCollector.md

### Overview

Lists every format (font, size, leading, letter spacing, justification and paragraph spacing) used in the active document,
paragraph by paragraph. Edit or replace a format and every paragraph in that format is updated at once,
much like redefining a paragraph style without using styles.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontPresetCollector.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontPresetCollector";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetCollector.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontPresetCollector.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS = 16;                  /* ウィンドウ外周の余白 / Window margins */
    var WINDOW_SPACING = 12;                  /* ウィンドウ内の要素間隔 / Window spacing */
    var ROW_SPACING = 8;                      /* 行グループ内の間隔 / Spacing inside a row group */
    var PANEL_MARGINS = [12, 16, 12, 12];     /* パネル内の余白 / Margins inside a panel */
    var PANEL_ROW_SPACING = 6;                /* パネル内の行間隔 / Spacing between rows inside a panel */
    var LIST_WIDTH = 340;                     /* 書式一覧の幅 / Width of the format list */
    var LIST_ROW_HEIGHT = 22;                 /* 一覧1行の高さの目安 / Estimated height of one list row */
    var LIST_ROW_COUNT = 14;                  /* 一覧に出す行数 / Rows shown in the list */
    var LABEL_WIDTH = 120;                    /* 項目名の幅 / Width of the row labels */
    var NUMBER_FIELD_CHARS = 6;               /* 数値欄の文字数 / Characters in a number field */
    var DROPDOWN_WIDTH = 220;                 /* ドロップダウンの幅 / Width of the dropdowns */

    /**
     * 行グループの向き・揃え・間隔をまとめて設定する
     * @param {Group} targetGroup - 対象の行グループ
     * @param {string} alignment - 横方向の揃え（"left" / "right" / "fill"）
     * @param {number} spacing - 要素間の間隔（省略時は ROW_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [alignment || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
    }

    /**
     * パネルの向き・揃え・余白・間隔をまとめて設定する
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_ROW_SPACING;
    }

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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function centerButtonRowIfRightOnly(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
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

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "段落書式の一括変更", en: "Batch Paragraph Formatting" }
        },
        panel: {
            edit: { ja: "書式の編集", en: "Edit Format" },
            editPending: { ja: "書式の編集（未反映の変更あり）", en: "Edit Format (unapplied changes)" }
        },
        fieldLabel: {
            font: { ja: "フォント", en: "Font" },
            size: { ja: "フォントサイズ", en: "Font Size" },
            leading: { ja: "行送り", en: "Leading" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            kern: { ja: "カーニング", en: "Kerning" },
            justify: { ja: "行揃え", en: "Justification" },
            spaceBefore: { ja: "段落前のアキ", en: "Space Before" },
            spaceAfter: { ja: "段落後のアキ", en: "Space After" }
        },
        checkbox: {
            autoLeading: { ja: "自動", en: "Auto" }
        },
        dropdown: {
            noReplaceSource: { ja: "—", en: "—" },
            kernMetrics: { ja: "メトリクス", en: "Metrics" },
            kernOptical: { ja: "オプティカル", en: "Optical" },
            kernRomanOnly: { ja: "和文等幅", en: "Metrics - Roman Only" },
            kernNone: { ja: "0", en: "0" },
            justifyLeft: { ja: "左揃え", en: "Align Left" },
            justifyCenter: { ja: "中央揃え", en: "Align Center" },
            justifyRight: { ja: "右揃え", en: "Align Right" },
            justifyLastLeft: { ja: "均等配置（最終行左揃え）", en: "Justify with Last Line Aligned Left" },
            justifyLastCenter: { ja: "均等配置（最終行中央揃え）", en: "Justify with Last Line Aligned Center" },
            justifyLastRight: { ja: "均等配置（最終行右揃え）", en: "Justify with Last Line Aligned Right" },
            justifyAllLines: { ja: "両端揃え", en: "Justify All Lines" }
        },
        button: {
            replace: { ja: "置換", en: "Replace" },
            update: { ja: "更新", en: "Update" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            formatList: {
                ja: "アクティブなドキュメントで使われている書式（かっこ内は段落数）。1つ選ぶと右側で値を編集できます。" +
                    "shift／command（Ctrl）＋クリックで複数選ぶと、まとめて置換できます。",
                en: "Formats used in the active document (paragraph count in brackets). Choose one to edit its values on the right; " +
                    "Shift/Command (Ctrl)-click to pick several and replace them together."
            },
            replace: {
                ja: "一覧で選んでいる書式（複数可）の段落を、ポップアップで選んだ書式にすぐに置き換えます。",
                en: "Replace the paragraphs in the formats chosen in the list (one or more) with the one picked in the popup, right away."
            },
            font: { ja: "ドキュメントで使われているフォントから選びます。", en: "Choose from the fonts used in the document." },
            tracking: { ja: "単位は 1/1000 em です。", en: "In 1/1000 em." },
            autoLeading: {
                ja: "オンにすると、行送りはフォントサイズに合わせて自動で決まります（自動行送りの割合はそのまま）。",
                en: "When on, leading follows the font size automatically (the auto-leading percentage is kept)."
            },
            update: { ja: "欄で編集した値を、その書式の段落へすぐに反映します。", en: "Apply the edited values to the paragraphs in this format right away." },
            ok: {
                ja: "変更を確定して閉じます（［更新］していない編集もここで反映します）。",
                en: "Keep the changes and close (edits not yet applied with Update are applied now)."
            },
            cancel: { ja: "反映した変更をすべて元に戻して閉じます。", en: "Undo every change that was applied and close." },
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
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            noText: { ja: "テキストが見つかりませんでした。", en: "No text was found." }
        }
    };

    /* カーニングの選択肢（並び順＝表示順）/ Kerning choices, in display order */
    var KERN_OPTIONS = [
        { id: "metrics", label: LABELS.dropdown.kernMetrics },
        { id: "optical", label: LABELS.dropdown.kernOptical },
        { id: "mono", label: LABELS.dropdown.kernRomanOnly },
        { id: "zero", label: LABELS.dropdown.kernNone }
    ];

    /* 行揃えの選択肢（並び順＝表示順）/ Justification choices, in display order */
    var JUSTIFY_OPTIONS = [
        { id: "left", label: LABELS.dropdown.justifyLeft },
        { id: "center", label: LABELS.dropdown.justifyCenter },
        { id: "right", label: LABELS.dropdown.justifyRight },
        { id: "justifyLeft", label: LABELS.dropdown.justifyLastLeft },
        { id: "justifyCenter", label: LABELS.dropdown.justifyLastCenter },
        { id: "justifyRight", label: LABELS.dropdown.justifyLastRight },
        { id: "justifyAll", label: LABELS.dropdown.justifyAllLines }
    ];

    // =========================================
    // 書式の読み書き / Reading and writing formats
    // =========================================

    /* 書式が持つ項目。一覧のまとまりと、更新時の比較はこの項目で行う
       The settings a format holds; grouping and the update diff both use these */
    var FORMAT_KEYS = [
        "psName", "size", "autoLeading", "leading", "tracking", "tsume",
        "kern", "justify", "spaceBefore", "spaceAfter"
    ];

    /**
     * 小数第2位までに丸める（10.5pt のような端数は残す）
     * @param {number} value - 丸めたい値
     * @returns {number} 丸めた値
     */
    function roundToHundredths(value) {
        return Math.round(value * 100) / 100;
    }

    /**
     * AutoKernType を方式 ID へ変換する
     * @param {AutoKernType} kernValue - 読み取ったカーニング
     * @returns {string} "metrics" / "optical" / "mono" / "zero"。判定できなければ ""
     */
    function kernMethodToId(kernValue) {
        var kernText = String(kernValue);
        if (kernText === String(AutoKernType.AUTO)) return "metrics";
        if (kernText === String(AutoKernType.OPTICAL)) return "optical";
        if (kernText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (kernText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "";
    }

    /**
     * 方式 ID を AutoKernType へ変換する
     * @param {string} methodId - "metrics" / "optical" / "mono" / "zero"
     * @returns {AutoKernType} カーニング
     */
    function resolveAutoKernType(methodId) {
        if (methodId === "mono") return AutoKernType.METRICSROMANONLY;
        if (methodId === "metrics") return AutoKernType.AUTO;
        if (methodId === "optical") return AutoKernType.OPTICAL;
        return AutoKernType.NOAUTOKERN;
    }

    /**
     * Justification を ID 文字列へ変換する
     * @param {Justification} justifyValue - 読み取った行揃え
     * @returns {string} 行揃えの ID。判定できなければ ""
     */
    function justificationToId(justifyValue) {
        var justifyText = String(justifyValue);
        if (justifyText === String(Justification.LEFT)) return "left";
        if (justifyText === String(Justification.CENTER)) return "center";
        if (justifyText === String(Justification.RIGHT)) return "right";
        if (justifyText === String(Justification.FULLJUSTIFYLASTLINELEFT)) return "justifyLeft";
        if (justifyText === String(Justification.FULLJUSTIFYLASTLINECENTER)) return "justifyCenter";
        if (justifyText === String(Justification.FULLJUSTIFYLASTLINERIGHT)) return "justifyRight";
        if (justifyText === String(Justification.FULLJUSTIFY)) return "justifyAll";
        return "";
    }

    /**
     * 行揃え ID を Justification へ変換する
     * @param {string} justifyId - 行揃えの ID
     * @returns {Justification} 行揃え
     */
    function resolveJustification(justifyId) {
        if (justifyId === "center") return Justification.CENTER;
        if (justifyId === "right") return Justification.RIGHT;
        if (justifyId === "justifyLeft") return Justification.FULLJUSTIFYLASTLINELEFT;
        if (justifyId === "justifyCenter") return Justification.FULLJUSTIFYLASTLINECENTER;
        if (justifyId === "justifyRight") return Justification.FULLJUSTIFYLASTLINERIGHT;
        if (justifyId === "justifyAll") return Justification.FULLJUSTIFY;
        return Justification.LEFT;
    }

    /**
     * PostScript 名からフォントを探す
     * @param {string} psName - フォントの PostScript 名
     * @returns {TextFont|null} フォント。見つからなければ null
     */
    function findFontByName(psName) {
        /* getByName は見つからないと例外を投げる / getByName throws when there is no match */
        try { return app.textFonts.getByName(psName); } catch (e) { return null; }
    }

    /**
     * 段落の書式を読む。文字属性は先頭の文字、段落属性は段落から読む
     * 行送りは、自動なら自動行送り量から算出した値、固定ならその値（pt）
     * @param {TextRange} paragraph - 対象の段落（空でないこと）
     * @returns {Object|null} 書式。フォントが読めなければ null
     */
    function readParagraphFormat(paragraph) {
        var charAttrs = paragraph.characters[0].characterAttributes;
        var paragraphAttrs = paragraph.paragraphAttributes;
        var currentFont = charAttrs.textFont;
        if (!currentFont || !currentFont.name) return null;
        var fontSize = roundToHundredths(charAttrs.size);
        var isAutoLeading = (charAttrs.autoLeading === true);
        var autoLeadingAmount = roundToHundredths(paragraphAttrs.autoLeadingAmount);
        return {
            psName: currentFont.name,
            size: fontSize,
            autoLeading: isAutoLeading,
            autoLeadingAmount: autoLeadingAmount,
            leading: isAutoLeading ? roundToHundredths(fontSize * autoLeadingAmount / 100) : roundToHundredths(charAttrs.leading),
            tracking: Math.round(charAttrs.tracking),
            tsume: Math.round(charAttrs.Tsume),
            kern: kernMethodToId(charAttrs.kerningMethod),
            justify: justificationToId(paragraphAttrs.justification),
            spaceBefore: roundToHundredths(paragraphAttrs.spaceBefore),
            spaceAfter: roundToHundredths(paragraphAttrs.spaceAfter)
        };
    }

    /**
     * 書式を、一覧のまとまりを決めるキー文字列にする
     * 自動行送りは算出値ではなく自動行送り量で比べる（サイズが同じなら結果も同じ）
     * @param {Object} format - 書式
     * @returns {string} キー文字列
     */
    function buildFormatKey(format) {
        var keyParts = [];
        for (var i = 0; i < FORMAT_KEYS.length; i++) {
            var keyName = FORMAT_KEYS[i];
            keyParts.push((keyName === "leading" && format.autoLeading) ? "auto" + format.autoLeadingAmount : String(format[keyName]));
        }
        return keyParts.join("\t");
    }

    /**
     * 書式の複製を返す
     * @param {Object} format - 複製したい書式
     * @returns {Object} 複製
     */
    function copyFormat(format) {
        var copied = {};
        for (var keyName in format) {
            if (format.hasOwnProperty(keyName)) copied[keyName] = format[keyName];
        }
        return copied;
    }

    /**
     * 編集の出発点になる書式を返す。行送りは既定で「自動」にする
     * @param {Object} original - 編集前の書式
     * @returns {Object} 編集用の書式
     */
    function createEditedFormat(original) {
        var edited = copyFormat(original);
        edited.autoLeading = true;
        edited.leading = roundToHundredths(edited.size * edited.autoLeadingAmount / 100);
        return edited;
    }

    /**
     * 2つの書式で値が違う項目の名前を返す
     * 行送り（自動・値）は、ほかの項目が違うときか、includeLeading が true のときだけ数える
     * @param {Object} fromFormat - 比べる元の書式
     * @param {Object} toFormat - 比べる先の書式
     * @param {boolean} includeLeading - 行送りだけの違いも数えるなら true
     * @returns {string[]} 違う項目の名前
     */
    function listDifferingKeys(fromFormat, toFormat, includeLeading) {
        var differingKeys = [], leadingKeys = [];
        for (var i = 0; i < FORMAT_KEYS.length; i++) {
            var keyName = FORMAT_KEYS[i];
            if (fromFormat[keyName] === toFormat[keyName]) continue;
            if (keyName === "autoLeading" || keyName === "leading") leadingKeys.push(keyName);
            else differingKeys.push(keyName);
        }
        /* 自動行送り量はキーの一部なので、自動どうしで量だけ違うときも行送りの違いとして数える
           The auto leading amount is part of the grouping key, so an amount-only difference counts as a leading change */
        if (toFormat.autoLeading && fromFormat.autoLeadingAmount !== toFormat.autoLeadingAmount) leadingKeys.push("autoLeadingAmount");
        if (differingKeys.length || includeLeading) differingKeys = differingKeys.concat(leadingKeys);
        return differingKeys;
    }

    /**
     * 段落へ、指定の項目だけを書き込む
     * 行送りは「自動」ならサイズに追従させ、固定なら値を入れる
     * @param {TextRange} paragraph - 書き込む段落
     * @param {Object} toFormat - 書き込む書式
     * @param {string[]} keyNames - 書き込む項目の名前
     * @param {TextFont} targetFont - フォントを変えるときに当てるフォント（見つからなければ null）
     * @returns {void}
     */
    function writeParagraphFormat(paragraph, toFormat, keyNames, targetFont) {
        var changedFlags = {};
        for (var i = 0; i < keyNames.length; i++) changedFlags[keyNames[i]] = true;
        var charAttrs = paragraph.characterAttributes;
        var paragraphAttrs = paragraph.paragraphAttributes;
        if (changedFlags.psName && targetFont) charAttrs.textFont = targetFont;
        if (changedFlags.size) charAttrs.size = toFormat.size;
        if (changedFlags.autoLeading || changedFlags.leading || changedFlags.autoLeadingAmount) {
            charAttrs.autoLeading = toFormat.autoLeading;
            if (!toFormat.autoLeading) charAttrs.leading = toFormat.leading;
        }
        if (changedFlags.autoLeadingAmount) paragraphAttrs.autoLeadingAmount = toFormat.autoLeadingAmount;
        if (changedFlags.tracking) charAttrs.tracking = toFormat.tracking;
        if (changedFlags.tsume) charAttrs.Tsume = toFormat.tsume;
        if (changedFlags.kern) {
            charAttrs.kerningMethod = resolveAutoKernType(toFormat.kern);
            /* メトリクスのときだけプロポーショナルメトリクスを ON / Proportional metrics ON only for Metrics */
            charAttrs.proportionalMetrics = (toFormat.kern === "metrics");
        }
        if (changedFlags.justify) paragraphAttrs.justification = resolveJustification(toFormat.justify);
        if (changedFlags.spaceBefore) paragraphAttrs.spaceBefore = toFormat.spaceBefore;
        if (changedFlags.spaceAfter) paragraphAttrs.spaceAfter = toFormat.spaceAfter;
    }

    /**
     * 段落へ書式を書き込む。書き込めない段落（ロック中のレイヤーなど）は飛ばす
     * @param {TextRange} paragraph - 書き込む段落
     * @param {Object} toFormat - 書き込む書式
     * @param {string[]} keyNames - 書き込む項目の名前
     * @param {TextFont} targetFont - フォントを変えるときに当てるフォント
     * @returns {void}
     */
    function tryWriteParagraphFormat(paragraph, toFormat, keyNames, targetFont) {
        try { writeParagraphFormat(paragraph, toFormat, keyNames, targetFont); } catch (e) { }
    }

    /**
     * 編集後の書式を、カンバス上の段落へすぐに反映する（反映済みの状態から変わった項目だけ）
     * 再描画は呼び出し側でまとめて行う
     * @param {Object} formatGroup - 書式のまとまり（applied / edited / leadingTouched / paragraphs）
     * @returns {void}
     */
    function applyEditedToCanvas(formatGroup) {
        var keyNames = listDifferingKeys(formatGroup.applied, formatGroup.edited, formatGroup.leadingTouched);
        if (!keyNames.length) return;
        var targetFont = findFontByName(formatGroup.edited.psName);
        for (var i = 0; i < formatGroup.paragraphs.length; i++) {
            tryWriteParagraphFormat(formatGroup.paragraphs[i], formatGroup.edited, keyNames, targetFont);
        }
        formatGroup.applied = copyFormat(formatGroup.edited);
    }

    /**
     * ダイアログを開いたときの書式へ、段落を戻す（今の書式を読み、違う項目だけ書き戻す）
     * 置換のたびに一覧を作り直すので、戻すときは開いたときの控えを使う
     * @param {Object[]} initialGroups - 開いたときに調べた書式のまとまり
     * @returns {void}
     */
    function restoreInitialFormats(initialGroups) {
        for (var i = 0; i < initialGroups.length; i++) {
            var original = initialGroups[i].original;
            var targetFont = findFontByName(original.psName);
            for (var j = 0; j < initialGroups[i].paragraphs.length; j++) {
                var paragraph = initialGroups[i].paragraphs[j];
                var keyNames = listDifferingKeys(readParagraphFormat(paragraph), original, true);
                if (keyNames.length) tryWriteParagraphFormat(paragraph, original, keyNames, targetFont);
            }
        }
    }

    // =========================================
    // ドキュメントの調査 / Surveying the document
    // =========================================

    /**
     * 全テキストフレームを段落ごとに調べ、書式ごとのまとまりにする（使われた文字数の多い順）
     * 空の段落は飛ばす
     * @param {Document} doc - 対象のドキュメント
     * @returns {Object[]} 書式のまとまり（original / applied（カンバス上の状態）/ edited / leadingTouched / paragraphs / chars）
     */
    function collectFormatGroups(doc) {
        var groupsByKey = {}, formatGroups = [];
        var textFrames = doc.textFrames;
        for (var i = 0; i < textFrames.length; i++) {
            var paragraphs = textFrames[i].paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                var paragraph = paragraphs[j];
                var charCount = paragraph.characters.length;
                if (!charCount) continue;
                var paragraphFormat = readParagraphFormat(paragraph);
                if (!paragraphFormat) continue;
                var groupKey = "k" + buildFormatKey(paragraphFormat);
                if (!groupsByKey.hasOwnProperty(groupKey)) {
                    groupsByKey[groupKey] = {
                        original: paragraphFormat,
                        applied: copyFormat(paragraphFormat),
                        edited: createEditedFormat(paragraphFormat),
                        leadingTouched: false,
                        paragraphs: [],
                        chars: 0
                    };
                    formatGroups.push(groupsByKey[groupKey]);
                }
                groupsByKey[groupKey].paragraphs.push(paragraph);
                groupsByKey[groupKey].chars += charCount;
            }
        }
        formatGroups.sort(function (a, b) { return b.chars - a.chars; });
        return formatGroups;
    }

    /**
     * カーソルだけの選択（長さ0）から、そのカーソルがある段落をストーリーの段落から探す
     * @param {TextRange} insertionPoint - カーソル位置の選択
     * @returns {TextRange|null} 段落。見つからなければ null
     */
    function paragraphAtInsertionPoint(insertionPoint) {
        var insertionOffset = insertionPoint.characterOffset;
        var storyParagraphs = insertionPoint.story.paragraphs;
        for (var i = 0; i < storyParagraphs.length; i++) {
            var paragraphStart = storyParagraphs[i].characterOffset;
            if (insertionOffset >= paragraphStart && insertionOffset <= paragraphStart + storyParagraphs[i].length) return storyParagraphs[i];
        }
        return null;
    }

    /**
     * 起動時に選択していた段落を返す
     * 文字の選択ならその最初の段落、オブジェクトの選択は段落が1つのテキストフレーム1つに限る
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextRange|null} 段落。決められなければ null
     */
    function findSelectedParagraph(doc) {
        var currentSelection = doc.selection;
        if (!currentSelection) return null;
        if (currentSelection.typename === "TextRange") {
            var touchedParagraphs = currentSelection.paragraphs;
            return touchedParagraphs.length ? touchedParagraphs[0] : paragraphAtInsertionPoint(currentSelection);
        }
        if (currentSelection.length !== 1 || currentSelection[0].typename !== "TextFrame") return null;
        var frameParagraphs = currentSelection[0].paragraphs;
        return (frameParagraphs.length === 1) ? frameParagraphs[0] : null;
    }

    /**
     * 起動時に選択していた段落の書式が、一覧の何番目かを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {Object[]} formatGroups - 書式のまとまり
     * @returns {number} 一覧の位置。見つからなければ 0（先頭）
     */
    function findInitialIndex(doc, formatGroups) {
        var selectedParagraph = findSelectedParagraph(doc);
        if (!selectedParagraph || !selectedParagraph.characters.length) return 0;
        var selectedFormat = readParagraphFormat(selectedParagraph);
        if (!selectedFormat) return 0;
        var selectedKey = buildFormatKey(selectedFormat);
        for (var i = 0; i < formatGroups.length; i++) {
            if (buildFormatKey(formatGroups[i].original) === selectedKey) return i;
        }
        return 0;
    }

    // =========================================
    // 一覧の表示 / Presenting formats
    // =========================================

    /* 文字サイズ・行送り・段落のアキに使う単位 / Unit for type size, leading and paragraph spacing */
    var textUnit = null;

    /* psName → 表示名（ファミリー＋スタイル）/ psName → display name (family + style) */
    var fontDisplayNames = {};

    /**
     * PostScript 名の表示名を返す（見つからないフォントは PostScript 名のまま）
     * @param {string} psName - フォントの PostScript 名
     * @returns {string} 表示名
     */
    function fontDisplayName(psName) {
        if (!fontDisplayNames.hasOwnProperty(psName)) {
            var font = findFontByName(psName);
            fontDisplayNames[psName] = font ? font.family + (font.style ? " " + font.style : "") : psName;
        }
        return fontDisplayNames[psName];
    }

    /**
     * pt の値を文字の単位で表示する文字列にする
     * @param {number} points - 値（pt）
     * @returns {string} 単位の数値（小数第2位まで）
     */
    function pointsToUnitText(points) {
        return String(roundToHundredths(points / textUnit.pointsPerUnit));
    }

    /**
     * 書式のまとまりの要約を返す（一覧とポップアップの行）
     * @param {Object} formatGroup - 書式のまとまり
     * @returns {string} 「フォント名  サイズ / 行送り  (段落数)」の文字列
     */
    function buildFormatSummary(formatGroup) {
        var format = formatGroup.original;
        return fontDisplayName(format.psName) + "  " +
            pointsToUnitText(format.size) + " / " + pointsToUnitText(format.leading) + " " + textUnit.label +
            "  (" + formatGroup.paragraphs.length + ")";
    }

    /**
     * フォントの選択肢を返す（ドキュメントで使われているフォント、重複なし）
     * @param {Object[]} formatGroups - 書式のまとまり
     * @returns {Object[]} 選択肢（id / label）
     */
    function buildFontChoices(formatGroups) {
        var fontChoices = [], seenFontNames = {};
        for (var i = 0; i < formatGroups.length; i++) {
            var psName = formatGroups[i].original.psName;
            if (seenFontNames.hasOwnProperty(psName)) continue;
            seenFontNames[psName] = true;
            var displayName = fontDisplayName(psName);
            fontChoices.push({ id: psName, label: { ja: displayName, en: displayName } });
        }
        return fontChoices;
    }

    // =========================================
    // ダイアログの組み立て / Building the dialog
    // =========================================

    /**
     * 右揃えの項目名を付けた行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル定義
     * @returns {Group} 追加した行
     */
    function addLabeledRow(parentPanel, labelSet) {
        var fieldRowGroup = parentPanel.add("group");
        setupRow(fieldRowGroup, "left");
        var rowLabel = fieldRowGroup.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return fieldRowGroup;
    }

    /**
     * 数値欄（と単位）の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル定義
     * @param {string} unitLabel - 欄の後ろに出す単位（空なら出さない）
     * @param {Object|null} stepOptions - ∧∨を付けるなら addStepper() に渡す増減の設定（min / max / integer）。付けないなら null
     * @returns {EditText} 追加した数値欄（∧∨は .stepperGroup で参照できる）
     */
    function addNumberRow(parentPanel, labelSet, unitLabel, stepOptions) {
        var fieldRowGroup = addLabeledRow(parentPanel, labelSet);
        var numberField;
        if (stepOptions) {
            /* ∧∨と入力欄は隙間0で突き合わせる。↑↓キーも∧∨と同じ処理で増減する / butt the stepper against the field */
            var stepperInputGroup = fieldRowGroup.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;
            /* 増減したら、手入力の確定と同じく書き戻す / store the value as a typed commit would */
            stepOptions.onStep = function (steppedField) {
                if (typeof steppedField.onChange === "function") steppedField.onChange();
            };
            var stepperGroup = addStepper(stepperInputGroup, function () { return numberField; }, stepOptions);
            numberField = stepperInputGroup.add("edittext", undefined, "");
            numberField.stepperGroup = stepperGroup;
            bindSteppedArrowKeys(numberField, stepperGroup);
        } else {
            numberField = fieldRowGroup.add("edittext", undefined, "");
        }
        numberField.characters = NUMBER_FIELD_CHARS;
        if (unitLabel) fieldRowGroup.add("statictext", undefined, unitLabel);
        return numberField;
    }

    /**
     * 選択肢のドロップダウンの行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル定義
     * @param {Object[]} choiceOptions - 選択肢（id / label）
     * @returns {DropDownList} 追加したドロップダウン
     */
    function addDropdownRow(parentPanel, labelSet, choiceOptions) {
        var fieldRowGroup = addLabeledRow(parentPanel, labelSet);
        var choiceDropdown = fieldRowGroup.add("dropdownlist", undefined, []);
        choiceDropdown.preferredSize.width = DROPDOWN_WIDTH;
        for (var i = 0; i < choiceOptions.length; i++) choiceDropdown.add("item", getLabel(choiceOptions[i].label));
        return choiceDropdown;
    }

    /**
     * 選択肢の ID から位置を返す
     * @param {Object[]} choiceOptions - 選択肢（id / label）
     * @param {string} optionId - 探す ID
     * @returns {number|null} 位置。無ければ null（ドロップダウンは未選択になる）
     */
    function findOptionIndex(choiceOptions, optionId) {
        for (var i = 0; i < choiceOptions.length; i++) {
            if (choiceOptions[i].id === optionId) return i;
        }
        return null;
    }

    /**
     * 左側（書式の一覧と、置換のポップアップ・ボタン）を組み立てる
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} controls - 作ったコントロールの追加先
     * @returns {void}
     */
    function buildListColumn(parentGroup, controls) {
        var listColumnGroup = parentGroup.add("group");
        listColumnGroup.orientation = "column";
        listColumnGroup.alignChildren = ["fill", "top"];
        listColumnGroup.spacing = PANEL_ROW_SPACING;

        controls.formatList = listColumnGroup.add("listbox", undefined, [], { multiselect: true });
        controls.formatList.preferredSize = [LIST_WIDTH, LIST_ROW_HEIGHT * LIST_ROW_COUNT];
        controls.formatList.helpTip = getLabel(LABELS.tooltip.formatList);

        controls.replaceRowGroup = listColumnGroup.add("group");
        setupRow(controls.replaceRowGroup, "fill");
        controls.replaceSourceDropdown = controls.replaceRowGroup.add("dropdownlist", undefined, []);
        controls.replaceSourceDropdown.alignment = ["fill", "center"];
        controls.replaceSourceDropdown.helpTip = getLabel(LABELS.tooltip.replace);
        controls.btnReplace = controls.replaceRowGroup.add("button", undefined, getLabel(LABELS.button.replace));
        controls.btnReplace.helpTip = getLabel(LABELS.tooltip.replace);
    }

    /**
     * 右側（書式の編集欄と［更新］）を組み立てる
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object[]} fontChoices - フォントの選択肢
     * @param {Object} controls - 作ったコントロールの追加先
     * @returns {void}
     */
    function buildEditPanel(parentGroup, fontChoices, controls) {
        var formatEditPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.edit));
        setupPanel(formatEditPanel);
        controls.formatEditPanel = formatEditPanel;

        controls.fontDropdown = addDropdownRow(formatEditPanel, LABELS.fieldLabel.font, fontChoices);
        controls.fontDropdown.helpTip = getLabel(LABELS.tooltip.font);
        controls.sizeField = addNumberRow(formatEditPanel, LABELS.fieldLabel.size, textUnit.label, { step: 1, min: 0 });
        controls.leadingField = addNumberRow(formatEditPanel, LABELS.fieldLabel.leading, textUnit.label, { step: 1, min: 0 });
        controls.autoLeadingCheck = controls.leadingField.parent.add("checkbox", undefined, getLabel(LABELS.checkbox.autoLeading));
        controls.autoLeadingCheck.helpTip = getLabel(LABELS.tooltip.autoLeading);
        controls.trackingField = addNumberRow(formatEditPanel, LABELS.fieldLabel.tracking, "", null);
        controls.trackingField.helpTip = getLabel(LABELS.tooltip.tracking);
        controls.tsumeField = addNumberRow(formatEditPanel, LABELS.fieldLabel.tsume, "%", { step: 1, min: 0, max: 100, integer: true });
        controls.kernDropdown = addDropdownRow(formatEditPanel, LABELS.fieldLabel.kern, KERN_OPTIONS);
        controls.justifyDropdown = addDropdownRow(formatEditPanel, LABELS.fieldLabel.justify, JUSTIFY_OPTIONS);
        controls.spaceBeforeField = addNumberRow(formatEditPanel, LABELS.fieldLabel.spaceBefore, textUnit.label, { step: 1, min: 0 });
        controls.spaceAfterField = addNumberRow(formatEditPanel, LABELS.fieldLabel.spaceAfter, textUnit.label, { step: 1, min: 0 });

        var editButtonRowGroup = formatEditPanel.add("group");
        setupRow(editButtonRowGroup, "right");
        editButtonRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        controls.btnUpdate = editButtonRowGroup.add("button", undefined, getLabel(LABELS.button.update));
        controls.btnUpdate.helpTip = getLabel(LABELS.tooltip.update);
    }

    /**
     * ダイアログを組み立て、操作に使うコントロールをまとめて返す
     * @param {Object[]} fontChoices - フォントの選択肢
     * @returns {Object} コントロール（window / formatList / replaceSourceDropdown / 各欄 / 各ボタン など）
     */
    function buildFormatDialog(fontChoices) {
        var controls = {};
        var formatDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        formatDialog.orientation = "column";
        formatDialog.alignChildren = ["fill", "top"];
        formatDialog.margins = WINDOW_MARGINS;
        formatDialog.spacing = WINDOW_SPACING;
        controls.window = formatDialog;

        var mainRowGroup = formatDialog.add("group");
        mainRowGroup.orientation = "row";
        mainRowGroup.alignChildren = ["fill", "top"];
        mainRowGroup.spacing = WINDOW_SPACING;

        buildListColumn(mainRowGroup, controls);
        buildEditPanel(mainRowGroup, fontChoices, controls);

        /* ボタンエリア（キャンセル／OK） / Button row (Cancel / OK) */
        var buttonRow = addButtonRow(formatDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnCancel.helpTip = getLabel(LABELS.tooltip.cancel);
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        centerButtonRowIfRightOnly(buttonRow);
        btnOK.helpTip = getLabel(LABELS.tooltip.ok);
        controls.btnOK = btnOK;
        return controls;
    }

    // =========================================
    // ダイアログの操作 / Dialog behavior
    // =========================================

    /**
     * 書式の一覧と編集欄のダイアログを表示する。［置換］［更新］はその場でカンバスへ反映する
     * @param {Object} dialogState - doc（対象のドキュメント）と groups（書式のまとまり。置換・更新で差し替える）
     * @param {number} initialIndex - 最初に選んでおく一覧の位置
     * @returns {boolean} OK で閉じたら true
     */
    function showFormatDialog(dialogState, initialIndex) {
        /* 書式のまとまり。置換・更新のたびに調べ直して差し替える / Format groups; replaced by a fresh survey after each Replace/Update */
        var formatGroups = dialogState.groups;
        var fontChoices = buildFontChoices(formatGroups);
        var controls = buildFormatDialog(fontChoices);
        var formatList = controls.formatList;
        var replaceSourceDropdown = controls.replaceSourceDropdown;

        /* 一覧を作り直す間は onChange を止める / Suppress onChange while the list is rebuilt */
        var suppressListEvents = false;

        /* 編集欄で編集中のまとまりの位置。1つだけ選んでいるときだけ 0 以上
           Index of the group being edited in the fields; 0 or more only while exactly one is selected */
        var editingIndex = -1;

        /* 一覧で選んでいるまとまりの位置（昇順）/ Indices of the groups selected in the list, ascending */
        var selectedIndices = [];

        /**
         * 一覧を作り直し、選択位置を戻す（行の文字を1行ずつ書き換えない。Mac では別の行に反映されることがある）
         * @returns {void}
         */
        function fillFormatList() {
            suppressListEvents = true;
            formatList.removeAll();
            for (var i = 0; i < formatGroups.length; i++) formatList.add("item", buildFormatSummary(formatGroups[i]));
            formatList.selection = selectedIndices.length ? selectedIndices : null;
            suppressListEvents = false;
        }

        /**
         * 置換のポップアップへ書式を並べ直す（先頭は「選んでいない」状態の目印）
         * @returns {void}
         */
        function fillReplaceSourceDropdown() {
            replaceSourceDropdown.removeAll();
            replaceSourceDropdown.add("item", getLabel(LABELS.dropdown.noReplaceSource));
            for (var i = 0; i < formatGroups.length; i++) replaceSourceDropdown.add("item", buildFormatSummary(formatGroups[i]));
            replaceSourceDropdown.selection = 0;
        }

        /**
         * 一覧で選んでいる行の位置を昇順で返す
         * @returns {number[]} 選んでいる行の位置
         */
        function readSelectedIndices() {
            var pickedIndices = [];
            var listSelection = formatList.selection;
            if (!listSelection) return pickedIndices;
            /* 複数選択の一覧でも、1行だけのときは配列でなく ListItem が返ることがある / A single pick may come back as a lone ListItem */
            if (listSelection.length === undefined) listSelection = [listSelection];
            for (var i = 0; i < listSelection.length; i++) pickedIndices.push(listSelection[i].index);
            pickedIndices.sort(function (a, b) { return a - b; });
            return pickedIndices;
        }

        /**
         * 位置が選んでいる行に含まれるかを返す
         * @param {number} groupIndex - 調べる位置
         * @returns {boolean} 含まれれば true
         */
        function isSelectedIndex(groupIndex) {
            for (var i = 0; i < selectedIndices.length; i++) {
                if (selectedIndices[i] === groupIndex) return true;
            }
            return false;
        }

        /**
         * まだカンバスに出していない編集を反映してから、ドキュメントを調べ直して一覧を作り直す
         * 同じ書式になったまとまりは1行に統合され、段落数も数え直される
         * @param {string} preferredKey - 作り直したあとに選んでおく書式のキー
         * @returns {void}
         */
        function rescanFormats(preferredKey) {
            for (var i = 0; i < formatGroups.length; i++) applyEditedToCanvas(formatGroups[i]);
            app.redraw();
            formatGroups = collectFormatGroups(dialogState.doc);
            dialogState.groups = formatGroups;
            selectedIndices = [];
            for (var j = 0; j < formatGroups.length; j++) {
                if (buildFormatKey(formatGroups[j].original) === preferredKey) { selectedIndices.push(j); break; }
            }
            fillReplaceSourceDropdown();
            fillFormatList();
            showSelectedFormat();
        }

        /**
         * 編集欄の見出しに、まだカンバスに出していない編集があるかを出す
         * @returns {void}
         */
        function refreshEditPanelTitle() {
            var editingGroup = (editingIndex >= 0) ? formatGroups[editingIndex] : null;
            var hasPending = editingGroup && listDifferingKeys(editingGroup.applied, editingGroup.edited, editingGroup.leadingTouched).length > 0;
            controls.formatEditPanel.text = getLabel(hasPending ? LABELS.panel.editPending : LABELS.panel.edit);
        }

        /**
         * 数値欄の値を pt で読む（読めなければ元の値）
         * @param {EditText} numberField - 対象の数値欄
         * @param {number} fallbackPoints - 読めないときの値（pt）
         * @returns {number} 値（pt、小数第2位まで）
         */
        function readPointsField(numberField, fallbackPoints) {
            var parsedValue = parseFloat(numberField.text);
            if (isNaN(parsedValue)) return fallbackPoints;
            /* 表示を丸めただけで値を変えたことにしない / Unchanged text must not count as an edit */
            if (numberField.text === pointsToUnitText(fallbackPoints)) return fallbackPoints;
            return roundToHundredths(parsedValue * textUnit.pointsPerUnit);
        }

        /**
         * 数値欄の値を整数で読む（読めなければ元の値）
         * @param {EditText} numberField - 対象の数値欄
         * @param {number} fallbackValue - 読めないときの値
         * @returns {number} 整数
         */
        function readIntegerField(numberField, fallbackValue) {
            var parsedValue = parseFloat(numberField.text);
            return isNaN(parsedValue) ? fallbackValue : Math.round(parsedValue);
        }

        /**
         * 行送りの欄と［自動］を、編集後の書式へ書き戻す
         * 自分で触ったら、既定の「自動」ではなく指定として扱う
         * @param {Object} editingGroup - 編集中の書式のまとまり
         * @returns {void}
         */
        function storeLeadingFields(editingGroup) {
            var edited = editingGroup.edited;
            var previousAuto = edited.autoLeading, previousLeading = edited.leading;
            edited.autoLeading = controls.autoLeadingCheck.value;
            if (edited.autoLeading) {
                /* 自動はサイズに追従するので、表示用の値だけ計算し直す / Auto follows the size; recompute the shown value */
                edited.leading = roundToHundredths(edited.size * edited.autoLeadingAmount / 100);
            } else if (previousAuto) {
                /* 自動を外したら、元の固定値（元も自動なら算出値）から始める / Unticking Auto starts from the original fixed value */
                edited.leading = editingGroup.original.autoLeading ? previousLeading : editingGroup.original.leading;
            } else {
                edited.leading = readPointsField(controls.leadingField, edited.leading);
            }
            if (edited.autoLeading !== previousAuto || (!edited.autoLeading && edited.leading !== previousLeading)) {
                editingGroup.leadingTouched = true;
            }
        }

        /**
         * 編集欄の値を、編集中のまとまりの編集後の書式へ書き戻す
         * 表示前には呼ばない（表示前のチェックボックスは値が読み戻せない）
         * @returns {void}
         */
        function storeFieldsToEdited() {
            if (editingIndex < 0) return;
            var editingGroup = formatGroups[editingIndex];
            var edited = editingGroup.edited;
            if (controls.fontDropdown.selection) edited.psName = fontChoices[controls.fontDropdown.selection.index].id;
            edited.size = readPointsField(controls.sizeField, edited.size);
            storeLeadingFields(editingGroup);
            edited.tracking = readIntegerField(controls.trackingField, edited.tracking);
            edited.tsume = readIntegerField(controls.tsumeField, edited.tsume);
            if (controls.kernDropdown.selection) edited.kern = KERN_OPTIONS[controls.kernDropdown.selection.index].id;
            if (controls.justifyDropdown.selection) edited.justify = JUSTIFY_OPTIONS[controls.justifyDropdown.selection.index].id;
            edited.spaceBefore = readPointsField(controls.spaceBeforeField, edited.spaceBefore);
            edited.spaceAfter = readPointsField(controls.spaceAfterField, edited.spaceAfter);
            refreshEditPanelTitle();
        }

        /**
         * 書式を編集欄へ表示する
         * @param {Object} format - 表示する書式
         * @returns {void}
         */
        function showFormatInFields(format) {
            controls.fontDropdown.selection = findOptionIndex(fontChoices, format.psName);
            controls.sizeField.text = pointsToUnitText(format.size);
            controls.leadingField.text = pointsToUnitText(format.leading);
            controls.autoLeadingCheck.value = format.autoLeading;
            controls.leadingField.enabled = !format.autoLeading;
            controls.leadingField.stepperGroup.enabled = !format.autoLeading;
            redrawSteppersIn(controls.leadingField.stepperGroup);
            controls.trackingField.text = String(format.tracking);
            controls.tsumeField.text = String(format.tsume);
            controls.kernDropdown.selection = findOptionIndex(KERN_OPTIONS, format.kern);
            controls.justifyDropdown.selection = findOptionIndex(JUSTIFY_OPTIONS, format.justify);
            controls.spaceBeforeField.text = pointsToUnitText(format.spaceBefore);
            controls.spaceAfterField.text = pointsToUnitText(format.spaceAfter);
        }

        /**
         * 一覧で選んでいる書式を編集欄に出し、置換のポップアップを合わせる（欄の書き戻しはしない）
         * @returns {void}
         */
        function showSelectedFormat() {
            selectedIndices = readSelectedIndices();
            /* 編集欄は1つだけ選んでいるときに使う。複数のときは先頭の値をディム表示 / Fields work with a single pick; with several, the first is shown dimmed */
            editingIndex = (selectedIndices.length === 1) ? selectedIndices[0] : -1;
            controls.formatEditPanel.enabled = (editingIndex >= 0);
            redrawSteppersIn(controls.formatEditPanel);
            controls.replaceRowGroup.enabled = (selectedIndices.length > 0);
            /* 一覧で選んでいる書式は、置換元としてディム表示 / Dim the list's own formats among the replace-with choices */
            for (var i = 0; i < formatGroups.length; i++) replaceSourceDropdown.items[i + 1].enabled = !isSelectedIndex(i);
            if (replaceSourceDropdown.selection && isSelectedIndex(replaceSourceDropdown.selection.index - 1)) replaceSourceDropdown.selection = 0;
            if (selectedIndices.length) showFormatInFields(formatGroups[selectedIndices[0]].edited);
            refreshEditPanelTitle();
        }

        /**
         * 一覧の選択に合わせて編集欄を切り替える（前のまとまりの編集は書き戻してから）
         * @returns {void}
         */
        function onFormatSelected() {
            if (suppressListEvents) return;
            storeFieldsToEdited();
            showSelectedFormat();
        }

        /**
         * 編集欄を変えたら、値を書き戻して表示を整える
         * @returns {void}
         */
        function onFieldChanged() {
            storeFieldsToEdited();
            if (editingIndex >= 0) showFormatInFields(formatGroups[editingIndex].edited);
        }

        /**
         * 更新：欄の値をその書式の段落へ反映し、調べ直す
         * @returns {void}
         */
        function onUpdateClick() {
            if (editingIndex < 0) return;
            storeFieldsToEdited();
            var updatedGroup = formatGroups[editingIndex];
            applyEditedToCanvas(updatedGroup);
            rescanFormats(buildFormatKey(updatedGroup.applied));
        }

        /**
         * 置換：ポップアップで選んだ書式の値を、一覧で選んでいる書式へそのまま写して反映する
         * 置換したものは置換元の書式に統合されるので、調べ直したあとはそれを選んでおく
         * @returns {void}
         */
        function onReplaceClick() {
            var sourceIndex = replaceSourceDropdown.selection ? replaceSourceDropdown.selection.index - 1 : -1;
            if (sourceIndex < 0 || !selectedIndices.length) return;
            var sourceFormat = formatGroups[sourceIndex].original;
            var replacedCount = 0;
            for (var i = 0; i < selectedIndices.length; i++) {
                /* 置換元そのものはディム表示だが、選べてしまう環境に備えて飛ばす / Skip the source itself in case its dimmed entry got picked */
                if (selectedIndices[i] === sourceIndex) continue;
                var targetGroup = formatGroups[selectedIndices[i]];
                targetGroup.edited = copyFormat(sourceFormat);
                /* 置換元の行送り（自動か固定か）をそのまま使う / Keep the source's leading as it is */
                targetGroup.leadingTouched = true;
                replacedCount++;
            }
            if (replacedCount) rescanFormats(buildFormatKey(sourceFormat));
            else replaceSourceDropdown.selection = 0;
        }

        formatList.onChange = onFormatSelected;
        var editControls = [controls.fontDropdown, controls.sizeField, controls.leadingField, controls.trackingField,
            controls.tsumeField, controls.kernDropdown, controls.justifyDropdown, controls.spaceBeforeField, controls.spaceAfterField];
        for (var i = 0; i < editControls.length; i++) editControls[i].onChange = onFieldChanged;
        controls.autoLeadingCheck.onClick = onFieldChanged;
        controls.btnUpdate.onClick = onUpdateClick;
        controls.btnReplace.onClick = onReplaceClick;
        /* OK の前に、表示中の欄を書き戻す / Store the shown fields before closing with OK */
        controls.btnOK.onClick = function () {
            storeFieldsToEdited();
            controls.window.close(1);
        };

        /* 最初の選択では欄を書き戻さない。表示前のチェックボックスは値が読み戻せず、「自動」が外れたことになるため
           The first selection must not store the fields: before show, the checkbox reads back as off and Auto would look unticked */
        selectedIndices = [initialIndex];
        fillReplaceSourceDropdown();
        fillFormatList();
        showSelectedFormat();

        prepareDialogWindow(controls.window, SCRIPT_NAME);
        return controls.window.show() === 1;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを調べてダイアログを表示する。OK なら変更を確定し、キャンセルなら元に戻す
     * 起動時に選択していた段落の書式を最初に選んでおく
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel(LABELS.alert.noDocument)); return; }
        var doc = app.activeDocument;
        textUnit = getUnitInfo("text/units");

        var formatGroups = collectFormatGroups(doc);
        if (!formatGroups.length) { alert(getLabel(LABELS.alert.noText)); return; }

        /* キャンセルで戻すための控え（ダイアログ側の一覧は置換のたびに作り直される）
           Snapshot for Cancel; the dialog rebuilds its own list after every Replace */
        var initialGroups = formatGroups.slice(0);
        var dialogState = { doc: doc, groups: formatGroups };
        if (showFormatDialog(dialogState, findInitialIndex(doc, formatGroups))) {
            /* OK なら、まだカンバスに出していない編集も反映 / OK applies any edits not yet on the canvas */
            for (var i = 0; i < dialogState.groups.length; i++) applyEditedToCanvas(dialogState.groups[i]);
        } else {
            /* キャンセルなら、開いたときの書式へ戻す / Cancel puts everything back as it was when the dialog opened */
            restoreInitialFormats(initialGroups);
        }
        app.redraw();
    }

    main();

})();
