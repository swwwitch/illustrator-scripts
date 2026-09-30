#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "DialogEngine"

/*

### 概要

アートボードのサイズを「操作」×「対象」×「サイズ」の組み合わせで自動調整します。
選択オブジェクトに合わせるほか、各アートボード内のオブジェクトに合わせて全アートボードを個別に調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardWithMargin.md

note記事も参照してください。
https://note.com/dtp_transit/n/n15d3c6c5a1e5

### Overview

Adjusts artboard size by operation, target and size (width & height).
Fits to the selection, or fits every artboard individually to the objects it contains.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitArtboardWithMargin.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitArtboardWithMargin";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.10.6";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardWithMargin.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitArtboardWithMargin.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_transit/n/n15d3c6c5a1e5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値（保存済みの設定があればそちらが優先） / Dialog defaults (stored settings win) */
    var DIALOG_DEFAULTS = {
        marginByUnit: {
            mm: '5',
            px: '20',
            pt: '10',
            _fallback: '0'
        },
        previewBounds: true,        // プレビュー境界(visibleBounds)を既定に / use visibleBounds by default
        roundMode: 'pixelGrid',     // 既定の丸めモード（pixelGrid / currentUnit / none）/ default rounding mode
        link: true                  // 上下左右の連動を既定ON / link margins by default
    };

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

    /* 行ラベルの幅（言語別） / Row label widths per language */
    var BASIS_LABEL_WIDTHS = { ja: 40, en: 76 };    /* 調整基準パネル / adjustment basis panel */
    var MARGIN_LABEL_WIDTHS = { ja: 32, en: 62 };   /* マージンパネル / margin panel */

    /* マージン入力欄の桁数 / Margin field width in characters */
    var MARGIN_FIELD_CHARACTERS = 4;

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

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

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
    var LINK_UI_DARK = isDarkUI();
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

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

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
     * 現在の定規単位を UnitValue に渡せる文字列で返す（UnitValue が扱えない単位は pt に寄せる）
     * @returns {string} 単位の文字列
     */
    function getRulerUnitString() {
        var unit = getUnitInfo();
        return (unit.code >= 0 && unit.code <= 6) ? unit.label : 'pt';
    }

    /**
     * 単位ごとの初期マージン値を返す
     * @param {string} unit - 単位の文字列
     * @returns {string} 初期マージン値
     */
    function getDefaultMargin(unit) {
        return DIALOG_DEFAULTS.marginByUnit.hasOwnProperty(unit) ?
            DIALOG_DEFAULTS.marginByUnit[unit] :
            DIALOG_DEFAULTS.marginByUnit._fallback;
    }

    /**
     * 数値＋単位を pt に変換する
     * @param {number|string} value - 値
     * @param {string} unit - 単位の文字列
     * @returns {number} pt。変換できなければ NaN
     */
    function toPt(value, unit) {
        var numericValue = Number(value);
        if (isNaN(numericValue)) return NaN;
        // 歯/Q は UnitValue が非対応のため単位表から換算（1H = 1Q = 0.25mm） / H and Q are unsupported by UnitValue
        if (unit === "H" || unit === "Q") {
            return numericValue * UNITS[5].pointsPerUnit;
        }
        /* UnitValue が単位を受け付けないときに備える / guard against units UnitValue rejects */
        try {
            return new UnitValue(numericValue, unit).as('pt');
        } catch (e) {
            return NaN;
        }
    }

    // =========================================
    // ローカライズ / Localize
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
            title: { ja: "アートボードサイズを調整", en: "Adjust Artboard Size" }
        },
        panel: {
            adjustmentBasis: { ja: "調整基準", en: "Adjustment basis" },
            margin: { ja: "マージン", en: "Margin" },
            fineTuning: { ja: "アートボードサイズの微調整", en: "Artboard size fine-tuning" }
        },
        fieldLabel: {
            operation: { ja: "操作", en: "Operation" },
            scope: { ja: "対象", en: "Target" },
            size: { ja: "サイズ", en: "Size" },
            vertical: { ja: "上下", en: "Vertical" },
            horizontal: { ja: "左右", en: "Horizontal" }
        },
        // 操作（fit/expand）と対象（current/all）の2軸、丸めモード / operation (fit/expand), scope (current/all), rounding mode
        radio: {
            fit: { ja: "オブジェクトに合わせる", en: "Fit to objects" },
            expand: { ja: "アートボードを拡張", en: "Expand artboard" },
            currentArtboard: { ja: "現在のアートボード", en: "Current artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
            roundPixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to pixel grid" },
            roundCurrentUnit: { ja: "現在の単位で値を整数値に", en: "Round values in current unit" },
            roundNone: { ja: "何もしない", en: "Do nothing" }
        },
        checkbox: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document first." },
            enterNumber: { ja: "数値を入力してください。", en: "Please enter a number." },
            errorOccurred: { ja: "エラーが発生しました：", en: "An error occurred: " },
            marginTooLarge: {
                ja: "マージンが大きすぎて有効なサイズにできないため、適用をスキップしました。",
                en: "The margin is too large to produce a valid size; skipped."
            }
        },
        tooltip: {
            fit: {
                ja: "オブジェクトの外接＋マージンのサイズにアートボードを合わせます（選択が無いときは各アートボード内のオブジェクトが対象）",
                en: "Resize artboards to the objects' bounds plus margins (with no selection, each artboard uses the objects it contains)"
            },
            expand: {
                ja: "アートボード自身のサイズにマージンを加減します（マイナス値で縮小）",
                en: "Grow/shrink the artboards themselves by the margins (negative shrinks)"
            },
            currentArtboard: {
                ja: "現在のアートボードのみを対象にします",
                en: "Apply to the current artboard only"
            },
            allArtboards: {
                ja: "すべてのアートボードを対象にします（「合わせる」では各アートボード内のオブジェクトに合わせます）",
                en: "Apply to all artboards (Fit uses the objects each artboard contains)"
            },
            marginInput: {
                ja: "↑↓・∧∨で次の整数へ、Shift+↑↓で10の倍数にスナップ、Option+↑↓で±0.1",
                en: "Arrow/stepper: to the next whole number, Shift: snap to 10, Option: ±0.1"
            },
            axisEnable: {
                ja: "OFFにすると実行時のサイズのまま固定します（連動は自動でOFF）。Option+クリックでこちらだけON",
                en: "Off keeps this dimension at its original size (auto-unlinks). Option-click to solo it"
            },
            linked: {
                ja: "上下の値を左右にも自動で適用します",
                en: "Apply the vertical value to horizontal as well"
            },
            previewBounds: {
                ja: "ON：線・効果を含む見た目の境界（プレビュー境界）で計測／OFF：パスの幾何境界で計測",
                en: "On: measure with preview (visible) bounds incl. strokes/effects; Off: geometric path bounds"
            },
            roundPixelGrid: {
                ja: "座標とサイズを整数ピクセルに丸めます",
                en: "Round position and size to integer pixels"
            },
            roundCurrentUnit: {
                ja: "現在の定規単位で座標とサイズを整数に丸めます",
                en: "Round position and size to integers in the current ruler unit"
            },
            roundNone: {
                ja: "丸めずに計測値のまま設定します",
                en: "Apply the measured values without rounding"
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
        }
    };

    // =========================================
    // エラー処理 / Error handling
    // =========================================

    /**
     * Error を行番号・ファイル名付きで読みやすく整形する
     * @param {Error} error - 例外
     * @returns {string} 整形した文字列
     */
    function formatError(error) {
        var messageText = (error && error.message) ? String(error.message) : String(error);
        var lineText = (error && error.line) ? (" line " + error.line) : "";
        var fileText = (error && error.fileName) ? (" (" + error.fileName + ")") : "";
        return messageText + lineText + fileText;
    }

    // =========================================
    // プレビュー管理 / Preview manager
    // =========================================

    /**
     * プレビューの適用と復元を制御するクラス / Preview apply/restore manager
     *
     * - updatePreview() のたびに rollback() で開いた時点の状態へ戻してから addStep() で最新状態を適用
     * - OK/Cancel 時に rollback() してプレビューを開いた時点へ戻す
     *
     * app.undo() の回数に依存すると、複数アートボード書き換え時に undo 粒度とズレて
     * 戻しすぎ/戻し不足が起きる。そのため巻き戻しは restoreFn（スナップショット復元）で行う。
     * @param {function} restoreFn - プレビュー前の状態へ戻す関数
     */
    function PreviewManager(restoreFn) {
        this.restoreFn = restoreFn;
        this.dirty = false; // プレビューによる変更が未復元か / preview changes pending restore

        /**
         * 変更操作を実行し、未復元フラグを立てる
         * @param {function} func - 変更操作
         * @returns {void}
         */
        this.addStep = function (func) {
            try {
                func();
                this.dirty = true;
                app.redraw();
            } catch (e) {
                $.writeln("[PreviewManager] addStep error: " + e);
            }
        };

        /**
         * プレビューを開いた時点の状態へ戻す
         * @returns {void}
         */
        this.rollback = function () {
            if (this.dirty && typeof this.restoreFn === "function") {
                try {
                    this.restoreFn();
                } catch (e) {
                    $.writeln("[PreviewManager] rollback error: " + e);
                }
            }
            this.dirty = false;
            app.redraw();
        };

        /**
         * 現在の状態を確定する（OK時）
         * OK時は一度 rollback() で元に戻してから main() 側で本処理を1回だけ実行するため、ここでは rollback のみ。
         * @returns {void}
         */
        this.confirm = function () {
            this.rollback();
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

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^\uFEFF/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    // =========================================
    // 設定の記憶（セッション内） / Settings persistence (session only)
    // =========================================
    // $.global に設定を保持。#targetengine のためセッション中は保持されるが、再起動でリセット。
    // Kept in $.global; persists during the session but resets when Illustrator restarts.

    /* 旧版の $.global のキーを1度だけ読み継ぐ / The old $.global key is read once */
    var LEGACY_SETTINGS_KEY = "__FitArtboardWithMargin_Settings";
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SETTINGS_KEY] || null; }
    });

    /**
     * 保存済みの設定を取得する。
     * 項目ごとの確認と既定値は resolveInitialSettings() が持つため、既定値は {}（中身を問わず受け取る）
     * @returns {Object} 設定。無ければ空のオブジェクト
     */
    function getStoredSettings() {
        return settingsStore.load({});
    }

    /**
     * 設定をセッションに保存する
     * @param {Object} settings - 設定
     * @returns {void}
     */
    function storeSettings(settings) {
        settingsStore.save(settings);
    }

    /**
     * 保存済み設定と文脈から、ダイアログの初期値を解決する
     * operation: "fit"（オブジェクトに合わせる、要選択）/ "expand"（アートボードを拡張）
     * scope: "current"（現在のアートボード）/ "all"（すべてのアートボード）
     * @param {string} defaultMargin - 保存が無いときのマージン
     * @param {number} artboardCount - アートボードの数
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @returns {Object} ダイアログの初期値
     */
    function resolveInitialSettings(defaultMargin, artboardCount, hasSelection) {
        var saved = getStoredSettings();

        // 操作：選択があれば fit、無ければ expand を既定に / operation default
        var operation = hasSelection ? "fit" : "expand";
        if (saved && (saved.operation === "fit" || saved.operation === "expand")) {
            operation = saved.operation;
        }

        // 対象：選択なし・複数アートボードなら all、それ以外は current を既定に / scope default
        var scope = (!hasSelection && artboardCount > 1) ? "all" : "current";
        if (saved && (saved.scope === "current" || saved.scope === "all")) {
            scope = saved.scope;
        }
        // 「合わせる」で選択が無いときは、各アートボード内のオブジェクトが対象になるため「すべて」固定
        // Fit without a selection works per artboard, so the scope is locked to all.
        if (operation === "fit" && !hasSelection) scope = "all";

        var link = (saved && typeof saved.link === "boolean") ? saved.link : DIALOG_DEFAULTS.link;
        // 連動ONのときは上下・左右とも有効に揃える / when linked, both axes are enabled
        var verticalEnabled = link ? true : ((saved && typeof saved.verticalEnabled === "boolean") ? saved.verticalEnabled : true);
        var horizontalEnabled = link ? true : ((saved && typeof saved.horizontalEnabled === "boolean") ? saved.horizontalEnabled : true);

        return {
            marginV: (saved && saved.marginV != null) ? saved.marginV : defaultMargin,
            marginH: (saved && saved.marginH != null) ? saved.marginH : defaultMargin,
            link: link,
            verticalEnabled: verticalEnabled,
            horizontalEnabled: horizontalEnabled,
            previewBounds: (saved && typeof saved.previewBounds === "boolean") ? saved.previewBounds : DIALOG_DEFAULTS.previewBounds,
            roundMode: (saved && saved.roundMode) ? saved.roundMode : DIALOG_DEFAULTS.roundMode,
            operation: operation,
            scope: scope
        };
    }

    // =========================================
    // 矩形・境界のユーティリティ / Rect & bounds utilities
    // =========================================

    /**
     * 矩形にマージンを加えた新しい矩形を返す
     * Illustrator の artboardRect は [left, top, right, bottom]（上が大・下が小）。
     * @param {number[]} rect - 元の矩形
     * @param {number} verticalMarginPt - 上下のマージン（pt）
     * @param {number} horizontalMarginPt - 左右のマージン（pt）
     * @returns {number[]} マージンを加えた矩形
     */
    function expandRectByMargin(rect, verticalMarginPt, horizontalMarginPt) {
        return [
            rect[0] - horizontalMarginPt,
            rect[1] + verticalMarginPt,
            rect[2] + horizontalMarginPt,
            rect[3] - verticalMarginPt
        ];
    }

    /**
     * アートボード矩形をピクセルグリッドに最適化する（X/Y/W/H を各1回だけ整数化）
     * X(左)・Y(上)・幅・高さの4値をそれぞれ整数に丸め、右下は X+幅 / Y−高さ で再構成する（二重丸めしない）。
     * @param {number[]} rect - 元の矩形
     * @returns {number[]} 丸めた矩形
     */
    function snapRectToPixelGrid(rect) {
        var x = Math.round(rect[0]);
        var y = Math.round(rect[1]);
        var width = Math.round(rect[2] - rect[0]);
        var height = Math.round(rect[1] - rect[3]);
        return [x, y, x + width, y - height];
    }

    /**
     * 矩形の X/Y/W/H を指定単位で各1回だけ整数化する
     * 単位変換に失敗した場合はピクセルグリッドにフォールバック。
     * @param {number[]} rect - 元の矩形
     * @param {string} unit - 単位の文字列
     * @returns {number[]} 丸めた矩形
     */
    function snapRectToUnitGrid(rect, unit) {
        var ptPerUnit = toPt(1, unit);
        if (isNaN(ptPerUnit) || ptPerUnit === 0) return snapRectToPixelGrid(rect);
        var x = Math.round(rect[0] / ptPerUnit) * ptPerUnit;
        var y = Math.round(rect[1] / ptPerUnit) * ptPerUnit;
        var width = Math.round((rect[2] - rect[0]) / ptPerUnit) * ptPerUnit;
        var height = Math.round((rect[1] - rect[3]) / ptPerUnit) * ptPerUnit;
        return [x, y, x + width, y - height];
    }

    /**
     * 矩形の幅・高さが正か（left<right かつ bottom<top）
     * @param {number[]} rect - 矩形
     * @returns {boolean} 有効なら true
     */
    function isValidRect(rect) {
        return rect[2] > rect[0] && rect[1] > rect[3];
    }

    /**
     * 無効化した軸を元アートボードの座標に固定する（＝その軸は動かさない）
     * 横(左右)OFFなら left/right を、縦(上下)OFFなら top/bottom を元アートボード値に戻す。
     * @param {number[]} rect - 計算した矩形
     * @param {number[]} artboardRect - 元のアートボード矩形
     * @param {boolean} verticalEnabled - 上下（高さ）を調整するか
     * @param {boolean} horizontalEnabled - 左右（幅）を調整するか
     * @returns {number[]} 無効軸を固定した矩形
     */
    function lockDisabledAxes(rect, artboardRect, verticalEnabled, horizontalEnabled) {
        var lockedRect = rect.slice();
        if (!horizontalEnabled) { lockedRect[0] = artboardRect[0]; lockedRect[2] = artboardRect[2]; }
        if (!verticalEnabled) { lockedRect[1] = artboardRect[1]; lockedRect[3] = artboardRect[3]; }
        return lockedRect;
    }

    /**
     * 丸めモードに従って矩形を整数化する
     * "pixelGrid" = ピクセル整数、"currentUnit" = 現在の単位で整数、"none" = 丸めなし。
     * @param {number[]} rect - 元の矩形
     * @param {string} roundMode - 丸めモード
     * @param {string} unit - 単位の文字列
     * @returns {number[]} 丸めた矩形
     */
    function applyRounding(rect, roundMode, unit) {
        if (roundMode === "pixelGrid") return snapRectToPixelGrid(rect);
        if (roundMode === "currentUnit") return snapRectToUnitGrid(rect, unit);
        return rect; // "none"
    }

    /**
     * 2つの矩形が実質的に同一か判定する
     * マージン0などで値が変わらない場合は書き換えを省き、取り消し履歴を増やさないために使う。
     * @param {number[]} rectA - 矩形A
     * @param {number[]} rectB - 矩形B
     * @returns {boolean} 同一なら true
     */
    function rectsEqual(rectA, rectB) {
        if (!rectA || !rectB) return false;
        for (var i = 0; i < 4; i++) {
            if (Math.abs(rectA[i] - rectB[i]) > 0.0001) return false;
        }
        return true;
    }

    /**
     * 計測可能なページアイテムか（geometricBounds を持つか）
     * TextRange 等の非ページアイテムを除外し、選択の型不整合を防ぐ。
     * @param {Object} item - 選択の要素
     * @returns {boolean} 計測できれば true
     */
    function isMeasurableItem(item) {
        if (!item) return false;
        /* TextRange などは geometricBounds で例外になる / non-page items throw on geometricBounds */
        try {
            var bounds = item.geometricBounds;
            return (bounds && bounds.length === 4);
        } catch (e) {
            return false;
        }
    }

    /**
     * 選択アイテムを正規化する
     * ・計測できない要素（TextRange 等）は除外
     * ・クリップグループはマスク（テキスト・複合パスの型を含む）のみを採用、それ以外はそのまま
     * @param {*} selection - 選択（doc.selection）
     * @returns {PageItem[]} 計測に使うアイテム
     */
    function collectEffectiveItems(selection) {
        var items = normalizeSelectionItems(selection);
        var effectiveItems = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!isMeasurableItem(item)) continue; // 非ページアイテムをスキップ / skip non-page items
            if (item.typename === "GroupItem" && item.clipped) {
                var clippingPath = getClipMaskItem(item);
                if (clippingPath) effectiveItems.push(clippingPath);
            } else {
                effectiveItems.push(item);
            }
        }
        return effectiveItems;
    }

    /**
     * 計測対象を再帰的に収集する
     * ・TextFrame は複製をアウトライン化（元は不変）。グループ内テキストも対象。
     * ・クリップグループはクリッピングパスのみ。通常グループは中身へ再帰。
     * 一時複製・アウトラインは tempObjects に登録し、呼び出し側で必ず削除する。
     * @param {PageItem[]} items - 対象
     * @param {PageItem[]} tempObjects - 一時オブジェクトの登録先
     * @param {PageItem[]} measureItems - 計測対象の格納先
     * @returns {void}
     */
    function collectMeasureTargets(items, tempObjects, measureItems) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!isMeasurableItem(item)) continue;
            if (item.typename === "TextFrame") {
                // 複製を先に登録 → アウトライン化（失敗時も複製を削除できる） / track clone before outlining
                var clone = item.duplicate();
                tempObjects.push(clone);
                var outlined = clone.createOutline(); // GroupItem を返す / returns a GroupItem
                tempObjects.push(outlined);
                measureItems.push(outlined);
            } else if (item.typename === "GroupItem" && item.clipped) {
                var clippingPath = getClipMaskItem(item);
                if (clippingPath) measureItems.push(clippingPath);
            } else if (item.typename === "GroupItem") {
                // 通常グループは中身を再帰（ネストされたテキストも非破壊計測） / recurse into groups
                collectMeasureTargets(item.pageItems, tempObjects, measureItems);
            } else {
                measureItems.push(item);
            }
        }
    }

    /**
     * 選択オブジェクトの外接境界を取得する
     * テキストは複製をアウトライン化して計測し、計測後に一時オブジェクトを必ず削除する。
     * 元の TextFrame には一切触れないため、ID・重なり順・名前・タグ・ノート・Variable 等が保持される。
     * @param {PageItem[]} items - 対象
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。空なら null
     */
    function measureSelectionBounds(items, usePreviewBounds) {
        var tempObjects = []; // 計測用に作った一時複製・アウトライン（必ず削除） / temp objects to remove
        try {
            var measureItems = [];
            collectMeasureTargets(items, tempObjects, measureItems);
            return getClipAwareUnionBounds(measureItems, usePreviewBounds);
        } finally {
            // 途中で例外が起きても一時オブジェクトは必ず削除 / always remove temp objects, even on error
            for (var k = 0; k < tempObjects.length; k++) {
                /* アウトライン化で消費された複製の remove() は例外になる / a clone consumed by createOutline() throws on remove() */
                try { tempObjects[k].remove(); } catch (e) { }
            }
        }
    }

    /**
     * 計測に使えるアイテムか（ロック・非表示・ガイド、およびそのレイヤーを除外）
     * @param {PageItem} item - 対象
     * @returns {boolean} 使えるなら true
     */
    function isUsableItem(item) {
        try {
            if (!item || item.locked || item.hidden || item.guides) return false;
            var parentLayer = item.parent;
            while (parentLayer && parentLayer.typename === "Layer") {
                if (!parentLayer.visible || parentLayer.locked) return false;
                parentLayer = parentLayer.parent;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * アートボードに重なるページアイテムを収集する
     * レイヤー直下のアイテムだけを見る（グループの中身はグループごと1件として扱う）。
     * @param {number[]} artboardRect - アートボード矩形
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {PageItem[]} 重なるアイテム
     */
    function getItemsInArtboard(artboardRect, usePreviewBounds) {
        var doc = app.activeDocument;
        var overlappingItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            try {
                if (item.parent.typename !== "Layer") continue; // 入れ子はグループと一緒に扱う / nested items travel with their group
                if (!isUsableItem(item)) continue;
                /* クリップグループはマスクの範囲で重なりを見る / Clip groups overlap by their mask */
                var bounds = getClipAwareBounds(item, usePreviewBounds);
                if (!bounds) continue;
                // 一辺でも外れていれば非交差 / no overlap when any edge clears the artboard
                if (bounds[2] <= artboardRect[0] || bounds[0] >= artboardRect[2] || bounds[3] >= artboardRect[1] || bounds[1] <= artboardRect[3]) continue;
                overlappingItems.push(item);
            } catch (e) { /* ignore */ }
        }
        return overlappingItems;
    }

    /**
     * アートボード内オブジェクトの外接境界を取得する
     * @param {number[]} artboardRect - アートボード矩形
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。対象が無ければ null
     */
    function measureArtboardContentBounds(artboardRect, usePreviewBounds) {
        var items = getItemsInArtboard(artboardRect, usePreviewBounds);
        if (items.length === 0) return null;
        return measureSelectionBounds(items, usePreviewBounds);
    }

    // =========================================
    // アートボード矩形の計算 / Artboard rect planning
    // =========================================
    // プレビューと確定で同じ計算（丸め・無効軸固定を含む）を使い、プレビューと結果を一致させる。
    // marginSettings: { verticalPt, horizontalPt, roundMode, unit, verticalEnabled, horizontalEnabled }

    /**
     * マージン適用の共通パイプライン：拡張 → 丸め → 無効軸を元座標に固定
     * @param {number[]} baseRect - 基準の矩形（アートボード自身、またはオブジェクトの外接）
     * @param {number[]} artboardOriginalRect - 実行時のアートボード矩形（無効軸の固定用）
     * @param {Object} marginSettings - マージン適用の設定
     * @returns {number[]} 新しいアートボード矩形
     */
    function computeMarginRect(baseRect, artboardOriginalRect, marginSettings) {
        var rect = expandRectByMargin(baseRect, marginSettings.verticalPt, marginSettings.horizontalPt);
        rect = applyRounding(rect, marginSettings.roundMode, marginSettings.unit);
        return lockDisabledAxes(rect, artboardOriginalRect, marginSettings.verticalEnabled, marginSettings.horizontalEnabled);
    }

    /**
     * 全アートボードの矩形を控える
     * @param {Artboards} artboards - アートボードのコレクション
     * @returns {number[][]} 矩形の配列
     */
    function snapshotArtboardRects(artboards) {
        var rects = [];
        for (var i = 0; i < artboards.length; i++) {
            rects.push(artboards[i].artboardRect.slice());
        }
        return rects;
    }

    /**
     * 操作と対象から、書き換えるアートボードの番号と新しい矩形を求める
     * ・拡張：アートボード自身の矩形にマージンを加減
     * ・合わせる×すべて：各アートボードを、その内側のオブジェクトに合わせる（対象が無いアートボードは据え置き）
     * ・合わせる×現在：現在のアートボードを選択の外接に合わせる
     * @param {string} operation - "fit" / "expand"
     * @param {string} scope - "current" / "all"
     * @param {number} activeIndex - 現在のアートボードの番号
     * @param {number[][]} originalRects - 実行時のアートボード矩形
     * @param {Object} boundsProvider - getContentBounds(index) と getSelectionBounds() を持つ計測元
     * @param {Object} marginSettings - マージン適用の設定
     * @returns {Object[]} { index, rect } の配列
     */
    function planArtboardRects(operation, scope, activeIndex, originalRects, boundsProvider, marginSettings) {
        var targetIndexes = [];
        if (scope === "all") {
            for (var i = 0; i < originalRects.length; i++) targetIndexes.push(i);
        } else {
            targetIndexes.push(activeIndex);
        }

        var rectPlans = [];
        for (var j = 0; j < targetIndexes.length; j++) {
            var artboardIndex = targetIndexes[j];
            var baseRect;
            if (operation === "expand") {
                baseRect = originalRects[artboardIndex];
            } else if (scope === "all") {
                baseRect = boundsProvider.getContentBounds(artboardIndex);
            } else {
                baseRect = boundsProvider.getSelectionBounds();
            }
            if (!baseRect) continue; // 計測対象が無い / nothing to measure
            rectPlans.push({ index: artboardIndex, rect: computeMarginRect(baseRect, originalRects[artboardIndex], marginSettings) });
        }
        return rectPlans;
    }

    /**
     * プレビューとしてアートボード矩形を書き換える
     * 無効な矩形と変化の無い矩形は書き換えない（取り消し履歴のノイズを減らす）。
     * @param {Object[]} rectPlans - planArtboardRects() の結果
     * @returns {void}
     */
    function previewArtboardRects(rectPlans) {
        var artboards = app.activeDocument.artboards;
        for (var i = 0; i < rectPlans.length; i++) {
            var rectPlan = rectPlans[i];
            if (isValidRect(rectPlan.rect) && !rectsEqual(artboards[rectPlan.index].artboardRect, rectPlan.rect)) {
                artboards[rectPlan.index].artboardRect = rectPlan.rect;
            }
        }
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * 調整基準パネルに「項目名＋コントロール群」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} contentOrientation - コントロール群の並び（"column" = 縦 / "row" = 横）
     * @returns {Group} コントロールを入れるグループ
     */
    function addBasisRow(parentPanel, labelSet, contentOrientation) {
        var isColumn = (contentOrientation === "column");
        var basisRow = parentPanel.add("group");
        basisRow.orientation = "row";
        basisRow.alignChildren = ["left", isColumn ? "top" : "center"];
        var rowLabel = basisRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = BASIS_LABEL_WIDTHS[uiLang];
        var contentGroup = basisRow.add("group");
        contentGroup.orientation = contentOrientation;
        contentGroup.alignChildren = isColumn ? "left" : ["left", "center"];
        return contentGroup;
    }

    /**
     * 調整基準パネル（操作＋対象＋サイズ）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildBasisPanel(marginDialog, initialSettings, dialogControls) {
        var basisPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.adjustmentBasis));
        setupPanel(basisPanel, 6);

        /* 操作：オブジェクトに合わせる / アートボードを拡張（ラジオは縦並び） / Operation group (vertical radios) */
        var operationGroup = addBasisRow(basisPanel, LABELS.fieldLabel.operation, "column");
        dialogControls.fitRadio = operationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.fit));
        dialogControls.fitRadio.helpTip = getLabel(LABELS.tooltip.fit);
        dialogControls.expandRadio = operationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.expand));
        dialogControls.expandRadio.helpTip = getLabel(LABELS.tooltip.expand);

        /* 対象：現在のアートボード / すべてのアートボード（ラジオは縦並び） / Scope group (vertical radios) */
        var scopeGroup = addBasisRow(basisPanel, LABELS.fieldLabel.scope, "column");
        dialogControls.currentRadio = scopeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.currentArtboard));
        dialogControls.currentRadio.helpTip = getLabel(LABELS.tooltip.currentArtboard);
        dialogControls.allRadio = scopeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.allArtboards));
        dialogControls.allRadio.helpTip = getLabel(LABELS.tooltip.allArtboards);

        /* サイズ：幅／高さ（OFFにした方は実行時のサイズのまま固定） / Axis targets: width & height (off keeps the original size) */
        var axisGroup = addBasisRow(basisPanel, LABELS.fieldLabel.size, "row");
        axisGroup.spacing = COLUMN_SPACING;
        dialogControls.widthCheckbox = axisGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.width));
        dialogControls.widthCheckbox.value = initialSettings.horizontalEnabled;
        dialogControls.widthCheckbox.helpTip = getLabel(LABELS.tooltip.axisEnable);
        dialogControls.heightCheckbox = axisGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.height));
        dialogControls.heightCheckbox.value = initialSettings.verticalEnabled;
        dialogControls.heightCheckbox.helpTip = getLabel(LABELS.tooltip.axisEnable);

        /* 初期値を適用 / apply initial selection */
        dialogControls.fitRadio.value = (initialSettings.operation === "fit");
        dialogControls.expandRadio.value = (initialSettings.operation !== "fit");
        dialogControls.currentRadio.value = (initialSettings.scope === "current");
        dialogControls.allRadio.value = (initialSettings.scope === "all");
    }

    /**
     * マージンの入力行（項目名＋入力欄）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} initialText - 入力欄の初期値
     * @returns {{label: StaticText, input: EditText}} 項目名と入力欄（∧∨は input.stepperGroup で参照できる）
     */
    function addMarginField(parentGroup, labelSet, initialText) {
        var fieldRow = parentGroup.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        var fieldLabel = fieldRow.add("statictext", undefined, getLabel(labelSet));
        fieldLabel.justify = "right";
        fieldLabel.preferredSize.width = MARGIN_LABEL_WIDTHS[uiLang]; /* 入力欄の位置を揃える / line up the inputs */

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var fieldInput;
        /* 増減後は手入力と同じ処理（連動・プレビュー更新）を通す。onChanging は showMarginDialog() で結線
           after stepping, run the same handler as typing (link and preview) */
        var stepperGroup = addStepper(stepperInputGroup, function () { return fieldInput; }, {
            onStep: function (numberInput) { if (numberInput.onChanging) numberInput.onChanging(); }
        });
        fieldInput = stepperInputGroup.add("edittext", undefined, initialText);
        fieldInput.characters = MARGIN_FIELD_CHARACTERS;
        fieldInput.helpTip = getLabel(LABELS.tooltip.marginInput);
        fieldInput.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(fieldInput, stepperGroup);
        return { label: fieldLabel, input: fieldInput };
    }

    /**
     * マージン入力パネル（入力欄と連動アイコン＋プレビュー境界の2カラム）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {string} rulerUnit - 定規単位
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildMarginPanel(marginDialog, rulerUnit, initialSettings, dialogControls) {
        var marginPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.margin) + " (" + rulerUnit + ")");
        setupPanel(marginPanel);
        marginPanel.orientation = "row";
        marginPanel.alignChildren = ["left", "top"];
        marginPanel.spacing = COLUMN_SPACING;

        /* 上下・左右の2行の右に連動アイコンを置く / Put the link icon to the right of the two margin rows */
        var marginFieldsRow = marginPanel.add("group");
        marginFieldsRow.orientation = "row";
        marginFieldsRow.alignChildren = ["left", "center"];

        var marginFieldsColumn = marginFieldsRow.add("group");
        marginFieldsColumn.orientation = "column";
        marginFieldsColumn.alignChildren = ["left", "center"];

        var optionColumn = marginPanel.add("group");
        optionColumn.orientation = "column";
        optionColumn.alignChildren = ["left", "center"];
        optionColumn.alignment = ["left", "center"];

        /* 上下マージン（「高さ」OFFで無効）・左右マージン（「幅」OFFで無効） / Vertical & horizontal margin inputs */
        var verticalField = addMarginField(marginFieldsColumn, LABELS.fieldLabel.vertical, initialSettings.marginV);
        dialogControls.verticalLabel = verticalField.label;
        dialogControls.verticalInput = verticalField.input;
        var horizontalField = addMarginField(marginFieldsColumn, LABELS.fieldLabel.horizontal,
            initialSettings.link ? initialSettings.marginV : initialSettings.marginH);
        dialogControls.horizontalLabel = horizontalField.label;
        dialogControls.horizontalInput = horizontalField.input;

        /* 連動アイコン。切り替え後の処理は showMarginDialog() で handleLinkToggle に結線
           Link icon; the handler is wired to handleLinkToggle in showMarginDialog() */
        dialogControls.linkToggle = addLinkToggle(marginFieldsRow, initialSettings.link, function () {
            if (dialogControls.handleLinkToggle) dialogControls.handleLinkToggle();
        });
        dialogControls.linkToggle.helpTip = getLabel(LABELS.tooltip.linked);

        /* プレビュー境界（visibleBounds を採用するか。「合わせる」のときのみ有効） / use visibleBounds; only for fit */
        dialogControls.previewBoundsCheckbox = optionColumn.add("checkbox", undefined, getLabel(LABELS.checkbox.previewBounds));
        dialogControls.previewBoundsCheckbox.value = initialSettings.previewBounds;
        dialogControls.previewBoundsCheckbox.helpTip = getLabel(LABELS.tooltip.previewBounds);
        dialogControls.previewBoundsCheckbox.enabled = (initialSettings.operation === "fit");
    }

    /**
     * アートボードサイズの微調整パネル（丸めモード）を作る
     * プレビューにも確定と同じ丸めを反映する。
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildFineTuningPanel(marginDialog, initialSettings, dialogControls) {
        var fineTuningPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.fineTuning));
        setupPanel(fineTuningPanel, 6);

        var roundModeGroup = fineTuningPanel.add("group");
        roundModeGroup.orientation = "column";
        roundModeGroup.alignChildren = "left";
        dialogControls.roundPixelRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundPixelGrid));
        dialogControls.roundUnitRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundCurrentUnit));
        dialogControls.roundNoneRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundNone));
        dialogControls.roundPixelRadio.helpTip = getLabel(LABELS.tooltip.roundPixelGrid);
        dialogControls.roundUnitRadio.helpTip = getLabel(LABELS.tooltip.roundCurrentUnit);
        dialogControls.roundNoneRadio.helpTip = getLabel(LABELS.tooltip.roundNone);
        dialogControls.roundUnitRadio.value = (initialSettings.roundMode === "currentUnit");
        dialogControls.roundNoneRadio.value = (initialSettings.roundMode === "none");
        dialogControls.roundPixelRadio.value = !dialogControls.roundUnitRadio.value && !dialogControls.roundNoneRadio.value; // 既定 / default
    }

    // =========================================
    // ダイアログの状態 / Dialog state
    // =========================================

    /**
     * 選んでいる操作を返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "fit" / "expand"
     */
    function getOperation(dialogControls) {
        return dialogControls.fitRadio.value ? "fit" : "expand";
    }

    /**
     * 選んでいる対象を返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "current" / "all"
     */
    function getScope(dialogControls) {
        return dialogControls.allRadio.value ? "all" : "current";
    }

    /**
     * 選んでいる丸めモードを返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "pixelGrid" / "currentUnit" / "none"
     */
    function getRoundMode(dialogControls) {
        if (dialogControls.roundPixelRadio.value) return "pixelGrid";
        return dialogControls.roundUnitRadio.value ? "currentUnit" : "none";
    }

    /**
     * 「合わせる」で選択が無いときは対象を「すべて」に固定してディムする
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @returns {void}
     */
    function refreshScopeState(dialogControls, hasSelection) {
        var forceAll = (getOperation(dialogControls) === "fit" && !hasSelection);
        if (forceAll) {
            dialogControls.allRadio.value = true;
            dialogControls.currentRadio.value = false;
        }
        dialogControls.currentRadio.enabled = !forceAll;
    }

    /**
     * マージンの入力欄の有効／無効を、∧∨ごとまとめて切り替える
     * @param {EditText} fieldInput - addMarginField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setMarginInputEnabled(fieldInput, isEnabled) {
        fieldInput.enabled = isEnabled;
        fieldInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(fieldInput.stepperGroup);
    }

    /**
     * 入力欄と連動アイコンの有効/無効を現在の状態から更新する
     * 上下=「高さ」ON、左右=「幅」ON かつ 非連動、連動=幅・高さが両方ONのときだけ有効。
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {void}
     */
    function refreshMarginInputStates(dialogControls) {
        var heightOn = dialogControls.heightCheckbox.value;
        var widthOn = dialogControls.widthCheckbox.value;
        dialogControls.verticalLabel.enabled = heightOn;
        setMarginInputEnabled(dialogControls.verticalInput, heightOn);
        dialogControls.horizontalLabel.enabled = widthOn;
        setMarginInputEnabled(dialogControls.horizontalInput, widthOn && !dialogControls.linkToggle.value);
        // 幅・高さのどちらかがOFFなら連動は使えない（自動OFFのうえディム） / disable link when either axis is off
        setLinkToggleEnabled(dialogControls.linkToggle, heightOn && widthOn);
    }

    /**
     * マージン欄の文字列を pt に変換する。無効な軸は 0 とみなす
     * @param {boolean} axisEnabled - その軸を調整するか
     * @param {string} marginText - 入力欄の文字列
     * @param {string} unit - 単位の文字列
     * @returns {number} pt。数値でなければ NaN
     */
    function marginTextToPoints(axisEnabled, marginText, unit) {
        return axisEnabled ? toPt(parseFloat(marginText), unit) : 0;
    }

    /**
     * ダイアログの入力から、マージン適用の設定を作る
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {string} rulerUnit - 定規単位
     * @returns {Object} マージン適用の設定。valid は有効な軸の値がすべて数値のとき true
     */
    function readMarginSettings(dialogControls, rulerUnit) {
        var verticalEnabled = dialogControls.heightCheckbox.value;
        var horizontalEnabled = dialogControls.widthCheckbox.value;
        var verticalPt = marginTextToPoints(verticalEnabled, dialogControls.verticalInput.text, rulerUnit);
        var horizontalPt = marginTextToPoints(horizontalEnabled, dialogControls.horizontalInput.text, rulerUnit);
        return {
            valid: !isNaN(verticalPt) && !isNaN(horizontalPt),
            verticalPt: verticalPt,
            horizontalPt: horizontalPt,
            roundMode: getRoundMode(dialogControls),
            unit: rulerUnit,
            verticalEnabled: verticalEnabled,
            horizontalEnabled: horizontalEnabled
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * マージン入力ダイアログを表示し設定を返す（ライブプレビュー付き）
     * @param {string} defaultMargin - 保存が無いときのマージン
     * @param {string} rulerUnit - 定規単位
     * @param {number} artboardCount - アートボードの数
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @param {PageItem[]} selectionItems - ダイアログ表示時に固定した選択アイテム（「合わせる」で使用）
     * @returns {Object|null} { operation, scope, previewBounds, marginSettings }。キャンセル時は null
     */
    function showMarginDialog(defaultMargin, rulerUnit, artboardCount, hasSelection, selectionItems) {
        var marginDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);

        setupWindow(marginDialog);

        // 保存済み設定（セッション内）から初期値を解決 / resolve initial values from stored settings
        var initialSettings = resolveInitialSettings(defaultMargin, artboardCount, hasSelection);

        var dialogControls = {};
        buildBasisPanel(marginDialog, initialSettings, dialogControls);
        buildMarginPanel(marginDialog, rulerUnit, initialSettings, dialogControls);
        buildFineTuningPanel(marginDialog, initialSettings, dialogControls);
        /* ボタン行（キャンセル → OK の順）/ Button row (Cancel, OK) */
        var buttonRow = addButtonRow(marginDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        dialogControls.btnCancel = btnCancel;
        dialogControls.btnOK = btnOK;
        refreshScopeState(dialogControls, hasSelection);
        refreshMarginInputStates(dialogControls);

        var verticalInput = dialogControls.verticalInput;
        var horizontalInput = dialogControls.horizontalInput;
        var widthCheckbox = dialogControls.widthCheckbox;
        var heightCheckbox = dialogControls.heightCheckbox;
        var linkToggle = dialogControls.linkToggle;
        var previewBoundsCheckbox = dialogControls.previewBoundsCheckbox;

        /* 現在／全アートボードの rect を保存（プレビュー復元用） / Save artboard rects for preview restore */
        var artboards = app.activeDocument.artboards;
        var activeArtboardIndex = artboards.getActiveArtboardIndex();
        var originalArtboardRects = snapshotArtboardRects(artboards);

        /* プレビュー復元：開いた時点の全 artboardRect へ書き戻す（undo回数に依存しない） / Restore all artboard rects to the opening snapshot */
        var previewManager = new PreviewManager(function () {
            var currentArtboards = app.activeDocument.artboards;
            for (var i = 0; i < currentArtboards.length && i < originalArtboardRects.length; i++) {
                if (!rectsEqual(currentArtboards[i].artboardRect, originalArtboardRects[i])) {
                    currentArtboards[i].artboardRect = originalArtboardRects[i];
                }
            }
        });

        /* 計測結果のキャッシュ（元の矩形と固定した選択で計測するのでダイアログ表示中は不変）
           Measured bounds cache; stable while the dialog is open */
        var boundsCache = {};
        var previewBoundsProvider = {
            getContentBounds: function (index) {
                var cacheKey = index + (previewBoundsCheckbox.value ? ":v" : ":g");
                if (!boundsCache.hasOwnProperty(cacheKey)) {
                    boundsCache[cacheKey] = measureArtboardContentBounds(originalArtboardRects[index], previewBoundsCheckbox.value);
                }
                return boundsCache[cacheKey];
            },
            getSelectionBounds: function () {
                var cacheKey = previewBoundsCheckbox.value ? "v" : "g";
                if (!boundsCache.hasOwnProperty(cacheKey)) {
                    boundsCache[cacheKey] = (selectionItems && selectionItems.length > 0) ?
                        measureSelectionBounds(selectionItems, previewBoundsCheckbox.value) : null;
                }
                return boundsCache[cacheKey];
            }
        };

        /* プレビュー更新：直前分を rollback してから最新状態を1回だけ適用 / Refresh preview via PreviewManager */
        function updatePreview() {
            previewManager.rollback();

            var marginSettings = readMarginSettings(dialogControls, rulerUnit);
            if (!marginSettings.valid) return;
            var operation = getOperation(dialogControls);
            var scope = getScope(dialogControls);

            previewManager.addStep(function () {
                previewArtboardRects(planArtboardRects(operation, scope, activeArtboardIndex, originalArtboardRects, previewBoundsProvider, marginSettings));
            });
        }

        /* 操作/対象の変更時：対象の固定・プレビュー境界の有効/無効を切り替えてプレビュー更新 / On change: refresh scope & previewBounds */
        function onBasisChange() {
            refreshScopeState(dialogControls, hasSelection); // 選択なしの「合わせる」は対象を「すべて」に固定 / fit without selection locks scope to all
            previewBoundsCheckbox.enabled = (getOperation(dialogControls) === "fit"); // 合わせる時のみ有効 / only for fit
            updatePreview();
        }

        /* 入力・ラジオ・連動のハンドラ登録（∧∨・↑↓キーも onChanging を通る） / Wire up input handlers (steppers and arrow keys go through onChanging) */
        verticalInput.onChanging = function () {
            if (linkToggle.value) horizontalInput.text = verticalInput.text;
            updatePreview();
        };
        horizontalInput.onChanging = function () {
            if (linkToggle.value) return; // 連動中は水平の直接編集は無効
            updatePreview();
        };

        /* Option(Alt)クリック検出：mousedown で event.altKey を捕捉（onClick では修飾キーを取得できない） / capture Alt on mousedown */
        var axisSoloRequested = { width: false, height: false };
        widthCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.width = (event.altKey === true); });
        heightCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.height = (event.altKey === true); });

        /**
         * 幅・高さチェックの変更処理
         * Option+クリック＝ソロ（クリックした方のみON、もう片方OFF）。どちらかOFFなら連動を自動OFF。
         * @param {boolean} solo - Option+クリックか
         * @param {Checkbox} clickedCheckbox - クリックしたチェックボックス
         * @param {Checkbox} otherCheckbox - もう片方のチェックボックス
         * @returns {void}
         */
        function handleAxisToggle(solo, clickedCheckbox, otherCheckbox) {
            if (solo) {
                clickedCheckbox.value = true;   // クリックした軸をON / keep the clicked axis on
                otherCheckbox.value = false;    // もう片方をOFF / turn the other off
            }
            if (!heightCheckbox.value || !widthCheckbox.value) setLinkToggleValue(linkToggle, false);
            refreshMarginInputStates(dialogControls);
            updatePreview();
        }
        widthCheckbox.onClick = function () {
            var solo = axisSoloRequested.width; axisSoloRequested.width = false;
            handleAxisToggle(solo, widthCheckbox, heightCheckbox);
        };
        heightCheckbox.onClick = function () {
            var solo = axisSoloRequested.height; axisSoloRequested.height = false;
            handleAxisToggle(solo, heightCheckbox, widthCheckbox);
        };

        dialogControls.handleLinkToggle = function () {
            if (linkToggle.value) {
                // 連動ON：幅・高さを有効に揃え、値を上下に統一 / linking re-enables both axes and mirrors V→H
                heightCheckbox.value = true;
                widthCheckbox.value = true;
                horizontalInput.text = verticalInput.text;
            }
            refreshMarginInputStates(dialogControls);
            updatePreview();
        };
        previewBoundsCheckbox.onClick = updatePreview;
        dialogControls.fitRadio.onClick = onBasisChange;
        dialogControls.expandRadio.onClick = onBasisChange;
        dialogControls.currentRadio.onClick = onBasisChange;
        dialogControls.allRadio.onClick = onBasisChange;
        dialogControls.roundPixelRadio.onClick = updatePreview;
        dialogControls.roundUnitRadio.onClick = updatePreview;
        dialogControls.roundNoneRadio.onClick = updatePreview;

        var dialogResult = null;
        dialogControls.btnOK.onClick = function () {
            var marginSettings = readMarginSettings(dialogControls, rulerUnit);
            if (!marginSettings.valid) {
                alert(getLabel(LABELS.alert.enterNumber));
                return;
            }
            dialogResult = {
                operation: getOperation(dialogControls),
                scope: getScope(dialogControls),
                previewBounds: previewBoundsCheckbox.value,
                marginSettings: marginSettings
            };
            // 設定をセッションに保存（次回の初期値に） / store settings for next run
            storeSettings({
                marginV: verticalInput.text,
                marginH: horizontalInput.text,
                link: linkToggle.value,
                verticalEnabled: marginSettings.verticalEnabled,
                horizontalEnabled: marginSettings.horizontalEnabled,
                previewBounds: dialogResult.previewBounds,
                roundMode: marginSettings.roundMode,
                operation: dialogResult.operation,
                scope: dialogResult.scope
            });
            // プレビューを開いた時点へ復元してから閉じる / restore to the opening snapshot, then close
            previewManager.confirm();
            marginDialog.close(1);
        };
        dialogControls.btnCancel.onClick = function () {
            // キャンセル時は必ずロールバックして閉じる / rollback preview and close
            previewManager.rollback();
            marginDialog.close(0);
        };

        updatePreview();

        // 開いたら有効な方のマージン入力欄にフォーカス（上下→無ければ左右） / focus a margin field on open
        if (verticalInput.enabled) {
            verticalInput.active = true;
        } else if (horizontalInput.enabled) {
            horizontalInput.active = true;
        }

        prepareDialogWindow(marginDialog, SCRIPT_NAME);
        marginDialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 対象を決定し、ダイアログの設定に従ってアートボードを調整する
     * @returns {void}
     */
    function main() {
        try {
            // ドキュメント未オープンなら分かりやすく案内して終了 / friendly guard when no document is open
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }
            var doc = app.activeDocument;

            // ダイアログ表示時点の選択を固定（「合わせる」で共通利用） / freeze selection at dialog time
            var selectionItems = collectEffectiveItems(doc.selection);
            var hasSelection = selectionItems.length > 0; // 計測可能な選択があるか / measurable selection?

            var artboards = doc.artboards;
            var rulerUnit = getRulerUnitString();

            /* 単位ごとの初期マージン値。選択なし・複数アートボード時は0（保存済み設定があればそちらが優先される）
               Default margin for the unit; 0 for the expand-all default (stored settings win) */
            var defaultMarginValue = (!hasSelection && artboards.length > 1) ? '0' : getDefaultMargin(rulerUnit);

            var dialogResult = showMarginDialog(defaultMarginValue, rulerUnit, artboards.length, hasSelection, selectionItems);
            if (!dialogResult) return; // キャンセル時は選択ツールにも切り替えない / cancel: leave tool unchanged

            var originalRects = snapshotArtboardRects(artboards);
            var previewBounds = dialogResult.previewBounds;
            var confirmBoundsProvider = {
                getContentBounds: function (index) { return measureArtboardContentBounds(originalRects[index], previewBounds); },
                getSelectionBounds: function () { return measureSelectionBounds(selectionItems, previewBounds); }
            };
            var rectPlans = planArtboardRects(dialogResult.operation, dialogResult.scope, artboards.getActiveArtboardIndex(),
                originalRects, confirmBoundsProvider, dialogResult.marginSettings);

            // 負マージン等で無効サイズになるものは適用しない / skip rects that end up invalid
            var skippedCount = 0;
            for (var i = 0; i < rectPlans.length; i++) {
                if (isValidRect(rectPlans[i].rect)) artboards[rectPlans[i].index].artboardRect = rectPlans[i].rect;
                else skippedCount++;
            }
            if (skippedCount > 0) alert(getLabel(LABELS.alert.marginTooLarge));

            // 適用が完了したときのみ選択ツールへ切り替え（キャンセル・エラー時は切り替えない） / switch tool only after a successful run
            app.selectTool("Adobe Select Tool");

        } catch (e) {
            $.writeln("[FitArtboardWithMargin] ERROR: " + formatError(e));
            alert(getLabel(LABELS.alert.errorOccurred) + formatError(e));
        }
    }

    main();

})();
