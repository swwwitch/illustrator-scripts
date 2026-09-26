#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

重なって配置されたオブジェクトを、横方向または縦方向へ指定した間隔で並べ直します。
行数・列数を指定すればタイル状に、キーオブジェクトを設定すればその位置を基準に配置できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf426908d8bcd

### Overview

Redistributes stacked objects along the horizontal or vertical axis at the spacing you specify.
Set a row or column count to tile them, or set a key object to anchor the layout to it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignAndTile";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf426908d8bcd"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値 / Initial dialog values */
    var DEFAULT_DIRECTION          = "horizontal"; /* 並べる方向（"horizontal" / "vertical"）/ tiling direction */
    var DEFAULT_LANE_COUNT         = "1";   /* 行数（横）・列数（縦）/ row count (horizontal) or column count (vertical) */
    var DEFAULT_MARGIN             = "0";   /* 横・縦の間隔 / horizontal & vertical spacing */
    var DEFAULT_LINK_MARGINS       = true;  /* 横・縦の間隔を連動 / link both spacings */
    var DEFAULT_USE_PREVIEW_BOUNDS = true;  /* プレビュー境界を使用 / use preview bounds */
    var DEFAULT_USE_GRID           = false; /* グリッド配置 / grid layout */
    var DEFAULT_RANDOMIZE          = false; /* ランダム配置 / random order */

    /* 整列後の位置差をどこまで「動いていない」とみなすか（pt）/ Move tolerance when probing the key object (pt) */
    var KEY_DETECT_TOLERANCE_PT = 0.001;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between columns */
    var FIELD_LABEL_WIDTH  = 40;               /* 項目名の幅 / width of a field label */
    var FIELD_CHAR_WIDTH   = 3;                /* 数値欄の文字数 / character width of a numeric field */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */
    var DIALOG_OFFSET_X    = 300;              /* ダイアログの横位置オフセット / dialog offset X */
    var DIALOG_OFFSET_Y    = 0;                /* ダイアログの縦位置オフセット / dialog offset Y */
    var DIALOG_OPACITY     = 0.97;             /* ダイアログの不透明度 / dialog opacity */

    /**
     * ウィンドウに共通レイアウトを適用する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
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
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 右揃えの項目名を追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} fieldLabelText - 表示する項目名（コロン付き）
     * @returns {StaticText} 生成した項目名
     */
    function addFieldLabel(parentContainer, fieldLabelText) {
        var fieldLabel = parentContainer.add("statictext", undefined, fieldLabelText);
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return fieldLabel;
    }

    /**
     * オプションのチェックボックスを追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} checkboxLabel - 表示するラベル
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 生成したチェックボックス
     */
    function addOptionCheckbox(parentContainer, checkboxLabel, initialValue) {
        var createdCheckbox = parentContainer.add("checkbox", undefined, checkboxLabel);
        createdCheckbox.alignment = "left"; /* パネル幅いっぱいに広げない / Do not stretch to the panel width */
        createdCheckbox.value = initialValue;
        return createdCheckbox;
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
    // 6. この欄には別の↑↓キー増減処理を付けない（↑↓キーが二重に効く）
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

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },
        panel: {
            direction: { ja: "方向", en: "Direction" },
            spacing:   { ja: "間隔", en: "Spacing" },
            alignment: { ja: "揃え", en: "Align" },
            options:   { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            rowCount:    { ja: "行数", en: "Rows" },
            columnCount: { ja: "列数", en: "Cols" },
            hMargin:     { ja: "横", en: "H" },
            vMargin:     { ja: "縦", en: "V" }
        },
        checkbox: {
            linkMargins:      { ja: "連動", en: "Link" },
            useKeyObject:     { ja: "キーオブジェクトを基準", en: "Anchor to key object" },
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            useGrid:          { ja: "グリッド", en: "Grid" },
            randomize:        { ja: "ランダム", en: "Random" }
        },
        radio: {
            directionHorizontal: { ja: "横", en: "Horizontal" },
            directionVertical:   { ja: "縦", en: "Vertical" },
            alignTop:    { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" },
            alignLeft:   { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight:  { ja: "右", en: "Right" },
            alignNone:   { ja: "なし", en: "None" }
        },
        tooltip: {
            directionHorizontal: {
                ja: "左から右へ並べます。行数を指定すると折り返します。",
                en: "Lays the objects out left to right, wrapping at the given number of rows."
            },
            directionVertical: {
                ja: "上から下へ並べます。列数を指定すると折り返します。",
                en: "Stacks the objects top to bottom, wrapping at the given number of columns."
            },
            laneCount: {
                ja: "何行（横並び）／何列（縦並び）で折り返すかを指定します。1 なら折り返しません。",
                en: "How many rows (horizontal) or columns (vertical) to wrap at. 1 means no wrapping."
            },
            useGrid: {
                ja: "各セルの大きさをそろえた格子に配置します。オフのときは各オブジェクトの大きさのまま詰めます。",
                en: "Places the objects on a grid of equal cells. Off packs them at their own sizes."
            },
            hMargin: { ja: "横方向のアキです。↑↓キーで増減できます。", en: "Horizontal gap. The arrow keys step the value." },
            vMargin: { ja: "縦方向のアキです。↑↓キーで増減できます。", en: "Vertical gap. The arrow keys step the value." },
            linkMargins: {
                ja: "横のアキと同じ値を縦にも使います。オフにすると縦を個別に指定できます。",
                en: "Uses the horizontal gap for the vertical one too. Turn it off to set them separately."
            },
            alignVertical: {
                ja: "各行の中でオブジェクトを上下どこにそろえるかです。",
                en: "Where to align the objects vertically within each row."
            },
            alignHorizontal: {
                ja: "各列の中でオブジェクトを左右どこにそろえるかです。",
                en: "Where to align the objects horizontally within each column."
            },
            useKeyObject: {
                ja: "最後にクリックしたキーオブジェクトの位置を動かさずに、他を並べ直します。キーオブジェクトが無いときは選べません。",
                en: "Keeps the key object where it is and arranges the rest around it. Unavailable when there is no key object."
            },
            usePreviewBounds: {
                ja: "線幅や効果を含めた見た目の端を基準にします。オフにするとパスの端が基準になります。",
                en: "Measures by the visible edges including strokes and effects. Off measures the path edges."
            },
            randomize: { ja: "並べる順序をシャッフルします。", en: "Shuffles the order the objects are laid out in." },
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
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:     { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            previewError:    { ja: "プレビューでエラーが発生しました", en: "Preview error" },
            unexpectedError: { ja: "エラーが発生しました", en: "An error has occurred" }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.spacing" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object|string} labelSet - ラベル、またはラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
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

    // =========================================
    // 位置とサイズ / Positions and sizes
    // =========================================

    /**
     * アイテムの境界を取得する（プレビュー境界ONならvisible、OFFならgeometric）
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(pageItem, usePreviewBounds) {
        return usePreviewBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
    }

    /**
     * 控えておいた位置へ戻す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {Array} positions - [[left, top], ...] の配列
     * @returns {void}
     */
    function resetPositions(targetItems, positions) {
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].left = positions[i][0];
            targetItems[i].top = positions[i][1];
        }
    }

    /**
     * 指定量だけまとめて移動する
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {number} dx - 横方向の移動量（pt）
     * @param {number} dy - 縦方向の移動量（pt）
     * @returns {void}
     */
    function shiftItems(targetItems, dx, dy) {
        if (!dx && !dy) {
            return;
        }
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            targetItems[i].left += dx;
            targetItems[i].top += dy;
        }
    }

    /**
     * 左端の座標順に並べ替えた複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortedCopyByLeft(targetItems) {
        var copiedItems = targetItems.slice();
        copiedItems.sort(function(itemA, itemB) {
            return itemA.left - itemB.left;
        });
        return copiedItems;
    }

    /**
     * 上端の座標順（上から下、同じなら左から右）に並べ替えた複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortedCopyByTop(targetItems) {
        var copiedItems = targetItems.slice();
        copiedItems.sort(function(itemA, itemB) {
            if (itemA.top !== itemB.top) return itemB.top - itemA.top;
            return itemA.left - itemB.left;
        });
        return copiedItems;
    }

    /**
     * ランダムに並べ替えた複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function shuffledCopy(targetItems) {
        var copiedItems = targetItems.slice();
        for (var i = copiedItems.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapItem = copiedItems[i];
            copiedItems[i] = copiedItems[j];
            copiedItems[j] = swapItem;
        }
        return copiedItems;
    }

    /**
     * 指定したオブジェクトを先頭へ移した複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} targetItem - 先頭へ移すオブジェクト
     * @returns {PageItem[]} 並べ替えた配列（対象が見つからないときはそのままの複製）
     */
    function movedToFront(targetItems, targetItem) {
        var reordered = targetItems.slice();
        for (var i = 0; i < reordered.length; i++) {
            if (reordered[i] !== targetItem) continue;
            reordered.splice(i, 1);
            reordered.unshift(targetItem);
            break;
        }
        return reordered;
    }

    /**
     * 選択範囲全体の左上を取得する
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {number[]} [左端, 上端]
     */
    function getBlockOrigin(targetItems) {
        var blockLeft = null;
        var blockTop = null;
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            if (blockLeft === null || targetItems[i].left < blockLeft) blockLeft = targetItems[i].left;
            if (blockTop === null || targetItems[i].top > blockTop) blockTop = targetItems[i].top;
        }
        return [blockLeft, blockTop];
    }

    /**
     * もっとも大きいアイテムの幅と高さを取得する（グリッドのセルサイズ）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
     * @returns {object} { width: number, height: number }
     */
    function getMaxItemSize(targetItems, usePreviewBounds) {
        var maxWidth = 0;
        var maxHeight = 0;
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            var itemBounds = getItemBounds(targetItems[i], usePreviewBounds);
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];
            if (itemWidth > maxWidth) maxWidth = itemWidth;
            if (itemHeight > maxHeight) maxHeight = itemHeight;
        }
        return { width: maxWidth, height: maxHeight };
    }

    /**
     * セル内での揃え量を求める（X軸は左→右、Y軸は上→下を start→end とする）
     * @param {string} alignMode - "start" / "center" / "end" / "none"
     * @param {number} cellStart - セルの起点（左端または上端）
     * @param {number} cellEnd - セルの終点（右端または下端）
     * @param {number} itemStart - オブジェクトの起点
     * @param {number} itemEnd - オブジェクトの終点
     * @returns {number} 移動量（pt）
     */
    function getAlignDelta(alignMode, cellStart, cellEnd, itemStart, itemEnd) {
        if (alignMode === "none") {
            return 0;
        }
        if (alignMode === "center") {
            return (cellStart + cellEnd) / 2 - (itemStart + itemEnd) / 2;
        }
        if (alignMode === "end") {
            return cellEnd - itemEnd;
        }
        return cellStart - itemStart;
    }

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================
    // Illustrator の DOM にキーオブジェクトを示すプロパティは無いため、整列コマンドで実測して特定する。
    // キーオブジェクトが設定されていると、どの向きに整列してもそのオブジェクトだけは動かない。
    // 候補が0個または2個以上のときは判定不能として null を返し、UI側でこの基準をディムする。

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * @param {PageItem[]} targetItems - 判定対象のオブジェクト
     * @returns {object} キーオブジェクト。判定できないときは null
     */
    function detectKeyObject(targetItems) {
        if (!targetItems || targetItems.length < 2) {
            return null;
        }
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        var i;
        for (i = 0; i < targetItems.length; i++) {
            stayedPut.push(true);
        }

        for (var c = 0; c < alignCommands.length; c++) {
            var savedPositions = [];
            for (i = 0; i < targetItems.length; i++) {
                savedPositions.push([targetItems[i].left, targetItems[i].top]);
            }
            app.redraw(); /* 直前のDOM変更が反映されていないと executeMenuCommand は空振りする / executeMenuCommand misfires without a redraw */
            app.executeMenuCommand(alignCommands[c]);
            for (i = 0; i < targetItems.length; i++) {
                if (Math.abs(targetItems[i].left - savedPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                    Math.abs(targetItems[i].top - savedPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                    stayedPut[i] = false;
                }
            }
            /* 整列は検出のための試行なので、その場で元の位置へ戻す / Undo the probe right away */
            resetPositions(targetItems, savedPositions);
        }
        app.redraw();

        var foundItem = null;
        for (i = 0; i < targetItems.length; i++) {
            if (!stayedPut[i]) continue;
            if (foundItem !== null) return null; /* 複数残った＝判定不能 / Ambiguous */
            foundItem = targetItems[i];
        }
        return foundItem;
    }

    // =========================================
    // プレビュー管理 / Preview management
    // =========================================

    /**
     * プレビューの適用・巻き戻し・確定をまとめて管理する
     * @constructor
     */
    function PreviewManager() {
        /* プレビュー中に実行したアクションの回数 / Number of preview actions executed */
        this.undoDepth = 0;

        /**
         * 変更操作を実行し、履歴としてカウントする
         * @param {Function} previewAction - 実行したい処理
         * @returns {void}
         */
        this.addStep = function(previewAction) {
            try {
                previewAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(labelText("alert.previewError") + " " + e);
            }
        };

        /**
         * プレビューのために行った変更をすべて取り消す
         * @returns {void}
         */
        this.rollback = function() {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * 現在の状態を確定する（プレビューを巻き戻してから1回だけ本番処理を実行）
         * @param {Function} [finalAction] - 巻き戻したあとに実行する処理
         * @returns {void}
         */
        this.confirm = function(finalAction) {
            if (finalAction) {
                this.rollback();
                finalAction();
            }
            this.undoDepth = 0;
        };
    }

    // =========================================
    // 配置処理 / Arranging
    // =========================================

    /**
     * 1行ずつセルに割り当てて配置する
     * @param {PageItem[]} orderedItems - 配置順に並べたオブジェクト
     * @param {object} arrangeSettings - 配置設定
     * @returns {void}
     */
    function placeItems(orderedItems, arrangeSettings) {
        var usePreviewBounds = arrangeSettings.usePreviewBounds;
        var isHorizontal = (arrangeSettings.direction === "horizontal");
        var cellSize = getMaxItemSize(orderedItems, usePreviewBounds);

        /* 先頭オブジェクトの位置を配置の起点にする / The first item defines the origin of the layout */
        var startBounds = getItemBounds(orderedItems[0], usePreviewBounds);
        /* 主軸＝並べる方向、副軸＝行・列が積み重なる方向 / Main axis follows the tiling direction; lanes stack along the cross axis */
        var mainOrigin = isHorizontal ? startBounds[0] : startBounds[1];
        var crossOrigin = isHorizontal ? startBounds[1] : startBounds[0];
        var mainGap = isHorizontal ? arrangeSettings.hMarginPt : arrangeSettings.vMarginPt;
        var crossGap = isHorizontal ? arrangeSettings.vMarginPt : arrangeSettings.hMarginPt;
        var laneSize = isHorizontal ? cellSize.height : cellSize.width;
        /* 上方向がプラスのY軸に合わせ、横並びは下へ、縦並びは右へレーンを送る / Lanes go down (horizontal) or right (vertical) */
        var laneDirection = isHorizontal ? -1 : 1;

        /* 主軸の揃えはグリッド時のみ意味を持つ（セル＝オブジェクトの大きさでは差が出ない）/ Main-axis align only matters in grid mode */
        var mainAlign = arrangeSettings.useGrid ? (isHorizontal ? arrangeSettings.hAlign : arrangeSettings.vAlign) : "start";
        var crossAlign = isHorizontal ? arrangeSettings.vAlign : arrangeSettings.hAlign;

        var remainingItems = orderedItems.length;
        var itemIndex = 0;
        for (var laneIndex = 0; laneIndex < arrangeSettings.laneCount; laneIndex++) {
            /* 残りを残りのレーン数で割り、指定した行数・列数を使い切る / Split the remainder so every lane is used */
            var itemsPerLane = Math.ceil(remainingItems / (arrangeSettings.laneCount - laneIndex));
            remainingItems -= itemsPerLane;
            var laneOffset = laneIndex * (laneSize + crossGap) * laneDirection;
            var crossStart = crossOrigin + laneOffset;
            var crossEnd = crossStart + laneSize * laneDirection;
            var mainStart = mainOrigin;

            for (var i = 0; i < itemsPerLane && itemIndex < orderedItems.length; i++, itemIndex++) {
                var placedItem = orderedItems[itemIndex];
                if (!placedItem) continue;

                var itemBounds = getItemBounds(placedItem, usePreviewBounds);
                var itemMainSize = isHorizontal ? (itemBounds[2] - itemBounds[0]) : (itemBounds[1] - itemBounds[3]);
                var cellMainSize = arrangeSettings.useGrid ? (isHorizontal ? cellSize.width : cellSize.height) : itemMainSize;
                var mainEnd = mainStart + cellMainSize * (isHorizontal ? 1 : -1);

                /* 主軸：セル内での揃え（なし＝その軸は動かさない）/ Main axis: align inside the cell ("none" leaves it alone) */
                var mainDelta = isHorizontal
                    ? getAlignDelta(mainAlign, mainStart, mainEnd, itemBounds[0], itemBounds[2])
                    : getAlignDelta(mainAlign, mainStart, mainEnd, itemBounds[1], itemBounds[3]);
                /* 副軸：レーンの帯へ揃える（なし＝レーンの送り分だけ平行移動）/ Cross axis: align to the lane band ("none" only applies the lane offset) */
                var crossDelta = (crossAlign === "none")
                    ? laneOffset
                    : (isHorizontal
                        ? getAlignDelta(crossAlign, crossStart, crossEnd, itemBounds[1], itemBounds[3])
                        : getAlignDelta(crossAlign, crossStart, crossEnd, itemBounds[0], itemBounds[2]));

                placedItem.left += isHorizontal ? mainDelta : crossDelta;
                placedItem.top += isHorizontal ? crossDelta : mainDelta;

                /* 次のセルへ / Advance to the next cell */
                mainStart = mainEnd + mainGap * (isHorizontal ? 1 : -1);
            }
        }
    }

    /**
     * 設定に従ってオブジェクトを並べ直す（プレビュー・確定の共通処理）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} arrangeSettings - 配置設定
     * @returns {void}
     */
    function arrangeItems(targetItems, arrangeSettings) {
        if (!targetItems || targetItems.length === 0) {
            return;
        }
        var orderedItems;
        if (arrangeSettings.randomize) {
            orderedItems = shuffledCopy(targetItems);
        } else {
            orderedItems = (arrangeSettings.direction === "horizontal") ? sortedCopyByLeft(targetItems) : sortedCopyByTop(targetItems);
        }
        /* キーオブジェクトは配置の起点にする（そこから右／下へ並べる）/ The key object becomes the origin, so the rest follow to its right or below */
        var useKeyObject = arrangeSettings.useKeyObject && arrangeSettings.keyObject && arrangeSettings.keyOrigin;
        if (useKeyObject) {
            orderedItems = movedToFront(orderedItems, arrangeSettings.keyObject);
        }
        /* 並べ替える前の左上（ランダム時に位置を戻す基準）/ Top-left before the layout, used to keep a random block in place */
        var blockOrigin = getBlockOrigin(targetItems);

        placeItems(orderedItems, arrangeSettings);

        /* 基準の補正：キーオブジェクトを優先し、なければランダム時のみ左上を合わせる / Anchor correction: key object first, random block otherwise */
        if (useKeyObject) {
            shiftItems(orderedItems,
                arrangeSettings.keyOrigin[0] - arrangeSettings.keyObject.left,
                arrangeSettings.keyOrigin[1] - arrangeSettings.keyObject.top);
        } else if (arrangeSettings.randomize) {
            shiftItems(orderedItems,
                blockOrigin[0] - orderedItems[0].left,
                blockOrigin[1] - orderedItems[0].top);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 行に∧∨と数値欄をひと組で追加する（隙間0で突き合わせ、↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - addStepper() に渡す増減の設定。onStep はダイアログ側であとから入れる
     * @returns {EditText} 入力欄（∧∨は .stepperGroup、増減の設定は .stepOptions で参照できる）
     */
    function addStepperInput(parentRow, initialText, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = FIELD_CHAR_WIDTH;
        numberInput.stepperGroup = stepperGroup;
        numberInput.stepOptions = stepOptions;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 入力欄と∧∨の有効／無効をまとめて切り替え、∧∨を描き直す
     * @param {EditText} numberInput - addStepperInput() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * ラジオボタンの一覧に同じツールチップを設定する
     * @param {RadioButton[]} radioList - 対象のラジオボタン
     * @param {string} helpTipText - 設定するツールチップ
     * @returns {void}
     */
    function setRadiosHelpTip(radioList, helpTipText) {
        for (var i = 0; i < radioList.length; i++) {
            radioList[i].helpTip = helpTipText;
        }
    }

    /**
     * ラジオボタンの一覧をまとめて有効・無効にする
     * @param {RadioButton[]} radioList - 対象のラジオボタン
     * @param {boolean} enabled - 有効にするかどうか
     * @returns {void}
     */
    function setRadiosEnabled(radioList, enabled) {
        for (var i = 0; i < radioList.length; i++) {
            radioList[i].enabled = enabled;
        }
    }

    /**
     * 揃えのラジオ（起点・中央・終点・なしの順）から、軸に依らない揃えの値を返す
     * @param {RadioButton[]} alignRadios - [起点, 中央, 終点, なし] の順に並んだラジオボタン
     * @returns {string} "start" / "center" / "end" / "none"
     */
    function getAlignValue(alignRadios) {
        if (alignRadios[1].value) return "center";
        if (alignRadios[2].value) return "end";
        if (alignRadios[3].value) return "none";
        return "start";
    }

    /**
     * 整列と分布ダイアログのパネルとコントロールを組み立てる（振る舞いの結線は呼び出し側で行う）
     * @param {Window} dialogWindow - 組み立て先のダイアログ
     * @param {PageItem} keyObject - キーオブジェクト（無ければ null）
     * @returns {object} 生成したコントロールをまとめたオブジェクト
     */
    function buildArrangeDialogControls(dialogWindow, keyObject) {
        /* 方向パネル（並べる方向と行数／列数）/ Direction panel: tiling direction and lane count */
        var directionPanel = addPanel(dialogWindow, getLabel("panel.direction"));

        var directionRow = directionPanel.add("group");
        setupRow(directionRow);
        var horizontalRadio = directionRow.add("radiobutton", undefined, getLabel("radio.directionHorizontal"));
        horizontalRadio.helpTip = getLabel("tooltip.directionHorizontal");
        var verticalRadio = directionRow.add("radiobutton", undefined, getLabel("radio.directionVertical"));
        verticalRadio.helpTip = getLabel("tooltip.directionVertical");
        horizontalRadio.value = (DEFAULT_DIRECTION === "horizontal");
        verticalRadio.value = !horizontalRadio.value;

        /* 行数（横）／列数（縦）/ Row count (horizontal) or column count (vertical) */
        var laneCountRow = directionPanel.add("group");
        setupRow(laneCountRow);
        var laneCountLabel = addFieldLabel(laneCountRow, "");
        /* 1以上の整数 / integer of 1 or more */
        var laneCountInput = addStepperInput(laneCountRow, DEFAULT_LANE_COUNT, { step: 1, min: 1, integer: true });
        laneCountInput.helpTip = getLabel("tooltip.laneCount");

        var gridCheckbox = addOptionCheckbox(directionPanel, getLabel("checkbox.useGrid"), DEFAULT_USE_GRID);
        gridCheckbox.helpTip = getLabel("tooltip.useGrid");

        /* 間隔パネル（左＝横・縦の入力、右＝連動）/ Spacing panel: fields on the left, link on the right */
        var spacingPanel = addPanel(dialogWindow, getLabel("panel.spacing") + " (" + getUnitInfo().label + ")");
        var spacingRow = spacingPanel.add("group");
        setupRow(spacingRow, "left", COLUMN_SPACING);

        var marginColumn = spacingRow.add("group");
        marginColumn.orientation = "column";
        marginColumn.alignChildren = ["left", "center"];
        marginColumn.spacing = PANEL_SPACING;

        var hMarginRow = marginColumn.add("group");
        setupRow(hMarginRow);
        addFieldLabel(hMarginRow, labelText("fieldLabel.hMargin"));
        /* 負のアキも可 / negative gaps allowed */
        var hMarginInput = addStepperInput(hMarginRow, DEFAULT_MARGIN, { step: 1 });
        hMarginInput.helpTip = getLabel("tooltip.hMargin");

        var vMarginRow = marginColumn.add("group");
        setupRow(vMarginRow);
        addFieldLabel(vMarginRow, labelText("fieldLabel.vMargin"));
        var vMarginInput = addStepperInput(vMarginRow, DEFAULT_MARGIN, { step: 1 });
        vMarginInput.helpTip = getLabel("tooltip.vMargin");

        var linkCheckbox = spacingRow.add("checkbox", undefined, getLabel("checkbox.linkMargins"));
        linkCheckbox.helpTip = getLabel("tooltip.linkMargins");
        linkCheckbox.value = DEFAULT_LINK_MARGINS;
        /* 連動中は縦をディムして横の値に合わせる / While linked, dim V and mirror H */
        setStepperInputEnabled(vMarginInput, !linkCheckbox.value);
        if (linkCheckbox.value) {
            vMarginInput.text = hMarginInput.text;
        }

        /* 揃えパネル（上段＝上下、下段＝左右）/ Align panel: vertical row on top, horizontal row below */
        var alignmentPanel = addPanel(dialogWindow, getLabel("panel.alignment"));

        var vAlignRow = alignmentPanel.add("group");
        setupRow(vAlignRow);
        var vAlignTopRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignTop"));
        var vAlignMiddleRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignMiddle"));
        var vAlignBottomRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignBottom"));
        var vAlignNoneRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignNone"));
        var vAlignRadios = [vAlignTopRadio, vAlignMiddleRadio, vAlignBottomRadio, vAlignNoneRadio];
        setRadiosHelpTip(vAlignRadios, getLabel("tooltip.alignVertical"));
        vAlignTopRadio.value = true;

        var hAlignRow = alignmentPanel.add("group");
        setupRow(hAlignRow);
        var hAlignLeftRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignLeft"));
        var hAlignCenterRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignCenter"));
        var hAlignRightRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignRight"));
        var hAlignNoneRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignNone"));
        var hAlignRadios = [hAlignLeftRadio, hAlignCenterRadio, hAlignRightRadio, hAlignNoneRadio];
        setRadiosHelpTip(hAlignRadios, getLabel("tooltip.alignHorizontal"));
        hAlignLeftRadio.value = true;

        /* オプションパネル / Options panel */
        var optionsPanel = addPanel(dialogWindow, getLabel("panel.options"));

        /* キーオブジェクトが未検出のときはディム / Dimmed when no key object is detected */
        var keyObjectCheckbox = addOptionCheckbox(optionsPanel, getLabel("checkbox.useKeyObject"), !!keyObject);
        keyObjectCheckbox.helpTip = getLabel("tooltip.useKeyObject");
        keyObjectCheckbox.enabled = !!keyObject;
        var previewBoundsCheckbox = addOptionCheckbox(optionsPanel, getLabel("checkbox.usePreviewBounds"), DEFAULT_USE_PREVIEW_BOUNDS);
        previewBoundsCheckbox.helpTip = getLabel("tooltip.usePreviewBounds");
        var randomizeCheckbox = addOptionCheckbox(optionsPanel, getLabel("checkbox.randomize"), DEFAULT_RANDOMIZE);
        randomizeCheckbox.helpTip = getLabel("tooltip.randomize");

        /* ボタンエリア（左右中央）/ Button bar, centered */
        var btnRowGroup = dialogWindow.add("group");
        setupRow(btnRowGroup, "center");
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            horizontalRadio: horizontalRadio,
            verticalRadio: verticalRadio,
            laneCountLabel: laneCountLabel,
            laneCountInput: laneCountInput,
            gridCheckbox: gridCheckbox,
            hMarginInput: hMarginInput,
            vMarginInput: vMarginInput,
            linkCheckbox: linkCheckbox,
            vAlignMiddleRadio: vAlignMiddleRadio,
            vAlignRadios: vAlignRadios,
            hAlignCenterRadio: hAlignCenterRadio,
            hAlignRadios: hAlignRadios,
            keyObjectCheckbox: keyObjectCheckbox,
            previewBoundsCheckbox: previewBoundsCheckbox,
            randomizeCheckbox: randomizeCheckbox
        };
    }

    /**
     * ダイアログの入力内容を配置設定として読み取る
     * @param {object} dialogControls - buildArrangeDialogControls() の戻り値
     * @param {PageItem} keyObject - キーオブジェクト（無ければ null）
     * @param {number[]} keyOrigin - キーオブジェクトのプレビュー前の位置（無ければ null）
     * @returns {object} 配置設定
     */
    function readArrangeSettings(dialogControls, keyObject, keyOrigin) {
        var unitFactor = getUnitInfo().pointsPerUnit;

        var hMarginValue = parseFloat(dialogControls.hMarginInput.text);
        if (isNaN(hMarginValue)) hMarginValue = 0;
        var vMarginValue = parseFloat(dialogControls.vMarginInput.text);
        if (isNaN(vMarginValue)) vMarginValue = 0;
        var laneCount = parseInt(dialogControls.laneCountInput.text, 10);
        if (isNaN(laneCount) || laneCount < 1) laneCount = 1;

        return {
            direction: dialogControls.horizontalRadio.value ? "horizontal" : "vertical",
            laneCount: laneCount,
            hMarginPt: hMarginValue * unitFactor,
            vMarginPt: vMarginValue * unitFactor,
            /* 揃えは軸に依らない形（start / center / end / none）で持つ / Align values are axis-neutral */
            vAlign: getAlignValue(dialogControls.vAlignRadios),
            hAlign: getAlignValue(dialogControls.hAlignRadios),
            usePreviewBounds: dialogControls.previewBoundsCheckbox.value,
            useGrid: dialogControls.gridCheckbox.value,
            randomize: dialogControls.randomizeCheckbox.value,
            useKeyObject: dialogControls.keyObjectCheckbox.value,
            keyObject: keyObject,
            keyOrigin: keyOrigin
        };
    }

    /**
     * 配置ダイアログを表示し、プレビューしながら設定を決める
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} keyObject - キーオブジェクト（未検出のときは null）
     * @returns {object} 確定した配置設定。キャンセル時は null
     */
    function showArrangeDialog(targetItems, keyObject) {
        var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow);
        dialogWindow.opacity = DIALOG_OPACITY;
        dialogWindow.onShow = function() {
            dialogWindow.location = [dialogWindow.location[0] + DIALOG_OFFSET_X, dialogWindow.location[1] + DIALOG_OFFSET_Y];
        };

        var previewManager = new PreviewManager();
        /* キャンセル時に戻せるよう、境界計算の環境設定を控える / Remember the bounds preference so Cancel can restore it */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");
        /* キーオブジェクトのプレビュー前の位置 / Key object position before any preview */
        var keyOrigin = keyObject ? [keyObject.left, keyObject.top] : null;

        var dialogControls = buildArrangeDialogControls(dialogWindow, keyObject);
        var horizontalRadio = dialogControls.horizontalRadio;
        var laneCountInput = dialogControls.laneCountInput;
        var gridCheckbox = dialogControls.gridCheckbox;
        var hMarginInput = dialogControls.hMarginInput;
        var vMarginInput = dialogControls.vMarginInput;
        var linkCheckbox = dialogControls.linkCheckbox;
        var previewBoundsCheckbox = dialogControls.previewBoundsCheckbox;

        /* 直近でプレビューへ反映した値（同じ値での二重更新を避ける）/ Value last pushed to the preview */
        var appliedLaneCountText = laneCountInput.text;
        /**
         * 行数・列数が変わったときだけプレビューを更新する
         * @returns {void}
         */
        function updatePreviewForLaneCount() {
            if (laneCountInput.text === appliedLaneCountText) return;
            appliedLaneCountText = laneCountInput.text;
            updatePreview();
        }
        /* 入力中は数値として読めるときだけ反映する（打っている途中で書き換えない）/ While typing, refresh only when the text parses */
        laneCountInput.onChanging = function() {
            var typedLaneCount = parseInt(laneCountInput.text, 10);
            if (isNaN(typedLaneCount) || typedLaneCount < 1) return;
            updatePreviewForLaneCount();
        };
        /* 確定時に1以上の整数へ丸める（表示と実際に使う値を一致させる）/ Snap to an integer of 1 or more on commit */
        laneCountInput.onChange = function() {
            var laneCountValue = parseInt(laneCountInput.text, 10);
            if (isNaN(laneCountValue) || laneCountValue < 1) laneCountValue = 1;
            laneCountInput.text = laneCountValue;
            updatePreviewForLaneCount();
        };
        /* ∧∨・↑↓キーで増減したあとの更新 / refresh after stepping with the stepper or arrow keys */
        laneCountInput.stepOptions.onStep = updatePreviewForLaneCount;
        hMarginInput.stepOptions.onStep = syncMarginsAndPreview;
        vMarginInput.stepOptions.onStep = updatePreview;

        /**
         * 方向とグリッドの状態に合わせてUIを整える（項目名と揃えの操作可否）
         * @returns {void}
         */
        function syncDirectionUI() {
            var isHorizontal = horizontalRadio.value;
            dialogControls.laneCountLabel.text = labelText(isHorizontal ? "fieldLabel.rowCount" : "fieldLabel.columnCount");
            /* 主軸（並べる方向）の揃えはグリッド時のみ有効 / Main-axis align is available in grid mode only */
            setRadiosEnabled(dialogControls.hAlignRadios, isHorizontal ? gridCheckbox.value : true);
            setRadiosEnabled(dialogControls.vAlignRadios, isHorizontal ? true : gridCheckbox.value);
        }
        syncDirectionUI();

        /**
         * Undo履歴を汚さずにプレビューを更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            /* 境界計算に使う環境設定は変わったときだけ書き換える（切り替えた直後は再描画しないと古い境界のまま計算される）/ Write the bounds preference only when it changes; without a redraw the old bounds are used */
            if (app.preferences.getBooleanPreference("includeStrokeInBounds") !== previewBoundsCheckbox.value) {
                app.preferences.setBooleanPreference("includeStrokeInBounds", previewBoundsCheckbox.value);
                app.redraw();
            }
            previewManager.addStep(function() {
                arrangeItems(targetItems, readArrangeSettings(dialogControls, keyObject, keyOrigin));
            });
        }

        /**
         * 方向・グリッドに合わせてUIを整えてからプレビューを更新する
         * @returns {void}
         */
        function syncDirectionAndPreview() {
            syncDirectionUI();
            updatePreview();
        }

        /**
         * 連動がONなら横の値を縦へ反映してからプレビューを更新する
         * @returns {void}
         */
        function syncMarginsAndPreview() {
            if (linkCheckbox.value) {
                vMarginInput.text = hMarginInput.text;
            }
            updatePreview();
        }

        hMarginInput.onChanging = syncMarginsAndPreview;
        hMarginInput.onChange = syncMarginsAndPreview;
        vMarginInput.onChanging = updatePreview;
        linkCheckbox.onClick = function() {
            setStepperInputEnabled(vMarginInput, !linkCheckbox.value);
            syncMarginsAndPreview();
        };
        horizontalRadio.onClick = syncDirectionAndPreview;
        dialogControls.verticalRadio.onClick = syncDirectionAndPreview;
        for (var i = 0; i < dialogControls.vAlignRadios.length; i++) {
            dialogControls.vAlignRadios[i].onClick = updatePreview;
        }
        for (var j = 0; j < dialogControls.hAlignRadios.length; j++) {
            dialogControls.hAlignRadios[j].onClick = updatePreview;
        }
        dialogControls.keyObjectCheckbox.onClick = updatePreview;
        previewBoundsCheckbox.onClick = updatePreview;
        dialogControls.randomizeCheckbox.onClick = updatePreview;
        gridCheckbox.onClick = function() {
            if (gridCheckbox.value) {
                /* グリッドは天地・左右とも中央を既定にする / Grid defaults to centered on both axes */
                dialogControls.vAlignMiddleRadio.value = true;
                dialogControls.hAlignCenterRadio.value = true;
            }
            syncDirectionAndPreview();
        };

        updatePreview();
        laneCountInput.active = true;

        if (dialogWindow.show() !== 1) {
            /* キャンセル：プレビューを巻き戻し、環境設定も元に戻す / Cancel: roll back the preview and the preference */
            previewManager.rollback();
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return null;
        }

        var arrangeSettings = readArrangeSettings(dialogControls, keyObject, keyOrigin);
        /* 1回のUndoで取り消せるように、巻き戻してから一度だけ実行する / Confirm as a single undoable action */
        previewManager.confirm(function() {
            arrangeItems(targetItems, arrangeSettings);
            /* 環境設定はスクリプトの外へ影響を残さないよう元に戻す / Restore the preference so the script leaves no global side effect */
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
        });
        return arrangeSettings;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを取得し、キーオブジェクトを判定してダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var targetItems = doc.selection;
        if (!targetItems || targetItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* キーオブジェクトを検出（整列コマンドで一時的に動かして元へ戻す）/ Detect the key object; it aligns temporarily, then restores */
        var keyObject = null;
        try {
            keyObject = detectKeyObject(targetItems);
        } catch (err) {
            /* 検出に失敗しても配置は続行する（チェックボックスがディムされるだけ）/ Keep going; only the checkbox is dimmed */
            $.writeln(SCRIPT_NAME + ": キーオブジェクトの検出に失敗 / key object detection failed — " + err);
        }

        showArrangeDialog(targetItems, keyObject);
    }

    try {
        main();
    } catch (e) {
        alert(labelText("alert.unexpectedError") + " " + e.message);
    }

})();
