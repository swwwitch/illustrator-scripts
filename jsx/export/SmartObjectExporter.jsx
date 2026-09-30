#target illustrator
#targetengine "SmartObjectExporterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを一時アートボードに収め、背景・マージン・枠線・書き出しサイズ・ファイル名を指定してPNG書き出しします。
設定した内容はアートボード上でそのままプレビューでき、よく使う組み合わせはプリセットとして呼び出せます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectExporter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/necf308c39f5d

### Overview

Places the selection on a temporary artboard and exports it as PNG with a chosen background, margin, border, size and filename.
Every setting is previewed on the artboard itself, and frequently used combinations can be recalled as presets.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectExporter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectExporter";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectExporter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectExporter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/necf308c39f5d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 作業用に一時生成するレイヤーの名前 / Name of the temporary working layer */
    var PREVIEW_LAYER_NAME = "__preview";

    /* 単位ごとの枠線幅の初期値 / Initial border width per ruler unit */
    var DEFAULT_BORDER_BY_UNIT = { mm: 0.1, _fallback: 1 };

    /* 透明グリッド1マスの基準サイズ（pt、100%のとき）/ Checker tile size at 100%, in points */
    var CHECKER_TILE_PT = 8;

    /* 透明グリッドの最大マス数（細かすぎる指定で描画が終わらなくなるのを防ぐ）
       / Cap on checker tiles, so a tiny percentage cannot stall the redraw */
    var MAX_CHECKER_TILES = 4000;

    /* PNG書き出しで指定できる最大倍率 / Maximum scale the PNG export accepts */
    var MAX_EXPORT_SCALE = 776.19;

    /* 倍率ラジオボタンに並べる値と初期選択 / Scale choices and the initial selection */
    var SCALE_CHOICES = [100, 200, 300, 400];
    var DEFAULT_SCALE = 400;

    /* プリセット保存ファイルの接頭辞（デスクトップに書き出す）/ Prefix of the preset file saved to the desktop */
    var PRESET_FILE_PREFIX = "export-setting-";

    /* 書き出しファイル名で接尾辞を省いたときの既定語 / Fallback word used when no suffix is given */
    var DEFAULT_SUFFIX_WORD = "selection";

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

    var NUMBER_FIELD_CHARS = 4;                /* 数値入力欄の文字数 / width of a numeric field */
    var MARGIN_FIELD_CHARS = 3;                /* マージン入力欄の文字数 / width of a margin field */
    var COLOR_FIELD_CHARS  = 12;               /* カラーコード入力欄の文字数（C0M100Y100K0 が収まる幅）/ width of a color code field */
    var SUFFIX_FIELD_CHARS = 14;               /* 接尾辞入力欄の文字数 / width of the suffix field */
    var SIZE_RADIO_WIDTH   = { ja: 60, en: 88 }; /* 倍率・横幅ラジオのラベル幅 / label width of the scale rows */
    var MARGIN_CELL_WIDTH  = { ja: 66, en: 82 }; /* マージン3×3グリッドの1マス幅 / cell width of the 3x3 margin grid */
    var FILENAME_ROW_HEIGHT = 22;              /* ファイル名プレビューの行高（ディセンダー切れ防止）/ row height of the filename preview */

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
    // 倍率ラジオの配色 / Scale radio colors
    // =========================================

    /* 選択中／非選択の文字色。暗いUIでは黒が沈むので明暗で切り替える（明暗はステップボタンの判定を流用）
       / Foreground colors of the scale radios; black disappears on a dark UI, so reuse the stepper's brightness check */
    var SCALE_ACTIVE_COLOR   = STEPPER_UI_DARK ? [0.9, 0.9, 0.9] : [0, 0, 0];
    var SCALE_INACTIVE_COLOR = [0.5, 0.5, 0.5];

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
            title: { ja: "選択オブジェクトをPNG書き出し", en: "Export Selected Objects as PNG" }
        },
        panel: {
            background: { ja: "背景", en: "Background" },
            margin: { ja: "マージン", en: "Margin" },
            border: { ja: "枠線", en: "Border" },
            size: { ja: "書き出しサイズ（px）", en: "Export Size (px)" },
            fileName: { ja: "書き出しファイル名", en: "Export Filename" },
            location: { ja: "書き出し先", en: "Export Location" }
        },
        radio: {
            transparent: { ja: "透過", en: "Transparent" },
            black: { ja: "黒", en: "Black" },
            white: { ja: "白", en: "White" },
            checker: { ja: "透明グリッド", en: "Transparency Grid" },
            colorCode: { ja: "カラーコード", en: "Color Code" },
            useDocName: { ja: "含める", en: "Include" },
            ignoreDocName: { ja: "含めない", en: "Exclude" },
            none: { ja: "なし", en: "None" },
            desktop: { ja: "デスクトップ", en: "Desktop" },
            documentFolder: { ja: "ドキュメントと同じ場所", en: "Same Folder as Document" }
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            borderWidth: { ja: "線幅", en: "Width" },
            borderColor: { ja: "枠線カラー", en: "Border Color" },
            customScale: { ja: "倍率", en: "Scale" },
            targetWidth: { ja: "横幅", en: "Target Width" },
            documentName: { ja: "ドキュメント名", en: "Document Name" },
            marginTop: { ja: "上", en: "Top" },
            marginBottom: { ja: "下", en: "Bottom" },
            marginLeft: { ja: "左", en: "Left" },
            marginRight: { ja: "右", en: "Right" },
            roundMode: { ja: "書き出し範囲の丸め", en: "Round Export Area" },
            delimiter: { ja: "区切り文字", en: "Delimiter" },
            suffix: { ja: "接尾辞", en: "Suffix" }
        },
        checkbox: {
            showFolder: { ja: "書き出し後、フォルダーを表示", en: "Show Folder After Export" }
        },
        roundMode: {
            pixelGrid: { ja: "ピクセルグリッドに最適化", en: "Align to pixel grid" },
            currentUnit: { ja: "定規の単位で整数に", en: "Whole ruler units" },
            none: { ja: "丸めない", en: "Don't round" }
        },
        tooltip: {
            numberField: {
                ja: "↑↓キーで増減（shift で10の倍数へ、option で0.1ずつ）",
                en: "Arrow keys step the value (Shift: to 10s, Option: by 0.1)"
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            linkMargin: { ja: "上の値を下・左・右にも適用します。", en: "Apply the top value to bottom, left, and right." },
            checker: {
                ja: "白とグレーの市松模様を背景に描いて書き出します。",
                en: "Exports with a white-and-gray checkerboard drawn as the background."
            },
            checkerScale: {
                ja: "1マスの大きさ。100%で8pt角です。",
                en: "Tile size of the grid; 100% is an 8pt square."
            },
            colorCode: {
                ja: "#RRGGBB / R255G255B255 / C0M100Y100K0 が使えます。",
                en: "Accepts #RRGGBB, R255G255B255 and C0M100Y100K0."
            },
            borderWidth: {
                ja: "書き出し範囲の内側に枠線を描きます。最小1pxまで切り上げます。",
                en: "Draws a border inside the export area, rounded up to at least 1px."
            },
            customScale: {
                ja: "上限は776.19%です。超える場合は上限の倍率で書き出します。",
                en: "Capped at 776.19%; anything higher is exported at the cap."
            },
            targetWidth: {
                ja: "指定した幅（px）になる倍率で書き出します。",
                en: "Exports at the scale that produces this width in pixels."
            },
            documentName: {
                ja: "ファイル名の先頭にドキュメント名を入れるかどうか。",
                en: "Whether the filename starts with the document name."
            },
            delimiter: {
                ja: "ドキュメント名と接尾辞の間に入れる文字。",
                en: "Character placed between the document name and the suffix."
            },
            suffix: {
                ja: "ファイル名の末尾に付ける文字。倍率・横幅を変えると自動で入ります。",
                en: "Text appended to the filename; the scale or width fills it in automatically."
            },
            savePreset: {
                ja: "現在の設定をテキストファイルとしてデスクトップに書き出します。",
                en: "Writes the current settings to a text file on the desktop."
            },
            roundPixel: {
                ja: "書き出し範囲を整数ピクセルまで広げます（倍率100%で1pt＝1px）。",
                en: "Grow the export area to whole pixels (at 100%, one pt equals one px)."
            },
            roundUnit: {
                ja: "書き出し範囲を現在の定規単位で整数になるまで広げます。",
                en: "Grow the export area to whole units in the current ruler unit."
            },
            roundNone: {
                ja: "丸めずに、計測した範囲のまま書き出します。",
                en: "Export the measured area as is, without rounding."
            }
        },
        dropdown: {
            custom: { ja: "カスタム", en: "Custom" }
        },
        button: {
            savePreset: { ja: "プリセットを保存", en: "Save Preset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        preset: {
            transparent200: { ja: "透過・マージンなし・倍率200%", en: "Transparent / No Margin / 200%" },
            whiteHorizontal300: { ja: "白背景・左右3mm・倍率300%", en: "White BG / Horizontal 3mm / 300%" },
            whiteVerticalWidth1000: { ja: "白背景・上下5mm・幅指定1000px", en: "White BG / Vertical 5mm / Width 1000px" },
            blackAll200: { ja: "黒背景・四辺10mm・倍率200%", en: "Black BG / All 10mm / 200%" }
        },
        prompt: {
            presetName: { ja: "プリセット名を入力してください", en: "Enter preset name" },
            defaultPresetName: { ja: "マイプリセット", en: "MyPreset" }
        },
        alert: {
            noSelection: {
                ja: "ドキュメントが開かれていないか、オブジェクトが選択されていません。",
                en: "No document open or no object selected."
            },
            invalidSize: {
                ja: "書き出す範囲を求められませんでした。マージンの値を確認してください。",
                en: "Could not determine the export area. Check the margin values."
            },
            exportFailed: { ja: "書き出しに失敗しました：", en: "Export failed: " },
            scaleLimited: {
                ja: "書き出し倍率が上限を超えたため、次の倍率で書き出しました：",
                en: "The requested scale exceeds the maximum, so the image was exported at: "
            },
            presetSaved: { ja: "プリセットを保存しました：", en: "Preset saved: " },
            presetSaveFailed: { ja: "プリセットの保存に失敗しました：", en: "Failed to save the preset: " }
        }
    };

    /* プリセットのマージン・線幅はmmで持ち、適用時に現在の定規単位へ換算する
       / Preset margins and border widths are stored in mm and converted to the ruler unit on apply */
    var PRESET_UNIT_FACTOR = 72.0 / 25.4;

    /* 初期プリセット（値の書式は background / margin / round / border / location / delimiter / suffix / size）
       / Built-in presets, encoded the same way as a saved preset file */
    var PRESETS = [
        {
            label: LABELS.preset.transparent200,
            background: "transparent",
            margin: "0,0,0,0",
            round: "pixelGrid",
            border: "none",
            location: "desktop",
            delimiter: "",
            suffix: "200",
            size: "scale:200"
        },
        {
            label: LABELS.preset.whiteHorizontal300,
            background: "white",
            margin: "0,0,3,3",
            round: "pixelGrid",
            border: "none",
            location: "desktop",
            delimiter: "_",
            suffix: "300",
            size: "scale:300"
        },
        {
            label: LABELS.preset.whiteVerticalWidth1000,
            background: "white",
            margin: "5,5,0,0",
            round: "pixelGrid",
            border: "1,black",
            location: "desktop",
            delimiter: "_",
            suffix: "1000",
            size: "width:1000"
        },
        {
            label: LABELS.preset.blackAll200,
            background: "black",
            margin: "10,10,10,10",
            round: "pixelGrid",
            border: "none",
            location: "desktop",
            delimiter: "-",
            suffix: "200",
            size: "scale:200"
        }
    ];

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {Object} titleLabelSet - パネル見出しのラベル定義
     * @param {string} [unitLabel] - 見出しに添える単位（省略時は付けない）
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, titleLabelSet, unitLabel) {
        var panelTitle = getLabel(titleLabelSet);
        if (unitLabel) panelTitle += (uiLang === "ja") ? "（" + unitLabel + "）" : " (" + unitLabel + ")";
        var createdPanel = parentContainer.add("panel", undefined, panelTitle);
        /* 行を詰めて並べる / Tighter spacing for the stacked rows */
        setupPanel(createdPanel, 6);
        return createdPanel;
    }

    /**
     * 横並びの行グループを生成する
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @param {string} [horizontalAlign] - 横方向の揃え
     * @returns {Group} 生成したグループ
     */
    function addRow(parentContainer, horizontalAlign) {
        var createdGroup = parentContainer.add("group");
        setupRow(createdGroup, horizontalAlign, 6);
        return createdGroup;
    }

    /**
     * 行の項目名（コロン付き）を追加する
     * @param {Group} parentGroup - 追加先の行グループ
     * @param {Object} labelSet - ラベル定義
     * @returns {StaticText} 追加した項目名
     */
    function addRowLabel(parentGroup, labelSet) {
        return parentGroup.add("statictext", undefined, labelText(labelSet));
    }

    /**
     * 数値入力欄を、左に∧∨を付けて追加する（↑↓キーも∧∨と同じ処理で増減。0 未満にはしない）
     * @param {Group} parentGroup - 追加先の行グループ
     * @param {string} initialText - 初期値
     * @param {number} charWidth - 入力欄の文字数
     * @param {function} onValueChanged - 値が変わったときに呼ぶコールバック（入力欄の文字列を受け取る）
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addNumberField(parentGroup, initialText, charWidth, onValueChanged) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = parentGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var inputField;
        var stepperGroup = addStepper(stepperInputGroup, function () { return inputField; }, {
            min: 0,
            onStep: function (numberInput) { onValueChanged(numberInput.text); }
        });
        inputField = stepperInputGroup.add("edittext", undefined, initialText);
        inputField.characters = charWidth;
        inputField.helpTip = getLabel(LABELS.tooltip.numberField);
        inputField.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(inputField, stepperGroup);
        inputField.onChange = function () { onValueChanged(inputField.text); };
        return inputField;
    }

    /**
     * addNumberField() で作った入力欄と∧∨の有効／無効をまとめて切り替え、∧∨を描き直す
     * @param {EditText} inputField - 対象の入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setNumberFieldEnabled(inputField, isEnabled) {
        inputField.enabled = isEnabled;
        inputField.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(inputField.stepperGroup);
    }

    /**
     * 数値入力欄のツールチップを組み立てる（個別の説明＋キー操作の説明）
     * @param {Object} labelSet - 個別の説明のラベル定義
     * @returns {string} ツールチップ文字列
     */
    function numberFieldTip(labelSet) {
        return getLabel(labelSet) + "\n" + getLabel(LABELS.tooltip.numberField);
    }

    /**
     * 縦に積むカラムを追加する
     * @param {Group} parentGroup - 追加先（横並びのグループ）
     * @returns {Group} 生成したカラム
     */
    function addColumn(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = COLUMN_SPACING;
        return columnGroup;
    }

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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位ごとの既定値を取り出す
     * @param {Object} defaultsByUnit - 単位ラベルをキーにした既定値の表
     * @param {string} unitLabel - 単位ラベル
     * @returns {number} 既定値
     */
    function getDefaultForUnit(defaultsByUnit, unitLabel) {
        return (defaultsByUnit[unitLabel] != null) ? defaultsByUnit[unitLabel] : defaultsByUnit._fallback;
    }

    /**
     * プリセットのmm値を現在の定規単位に換算する
     * @param {number} valueMm - mm値
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {number} 現在の定規単位での値（小数第3位まで）
     */
    function fromPresetUnit(valueMm, unitFactor) {
        return Math.round(valueMm * PRESET_UNIT_FACTOR / unitFactor * 1000) / 1000;
    }

    /**
     * 現在の定規単位の値をプリセット用のmmに換算する
     * @param {number} value - 現在の定規単位での値
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {number} mm値（小数第3位まで）
     */
    function toPresetUnit(value, unitFactor) {
        return Math.round(value * unitFactor / PRESET_UNIT_FACTOR * 1000) / 1000;
    }

    /**
     * マージン指定の4値をまとめて換算する
     * @param {string} marginSpec - マージン指定（上,下,左,右）
     * @param {function} convertValue - 1値ずつの換算関数
     * @returns {string} 換算後のマージン指定
     */
    function convertMarginSpec(marginSpec, convertValue) {
        var marginValues = String(marginSpec).split(",");
        var converted = [];
        for (var i = 0; i < 4; i++) {
            converted.push(convertValue(toNumber(marginValues[i])));
        }
        return converted.join(",");
    }

    /**
     * 枠線指定の線幅を換算する
     * @param {string} borderSpec - 枠線指定（線幅,カラー）
     * @param {function} convertValue - 換算関数
     * @returns {string} 換算後の枠線指定
     */
    function convertBorderSpec(borderSpec, convertValue) {
        if (!borderSpec || borderSpec === "none") return "none";
        var specParts = String(borderSpec).split(",");
        return convertValue(toNumber(specParts[0])) + "," + specParts[1];
    }

    /**
     * 入力文字列を数値として読み取る（不正値は0）
     * @param {string} inputText - 入力文字列
     * @returns {number} 読み取った数値
     */
    function toNumber(inputText) {
        var parsedValue = parseFloat(inputText);
        return isNaN(parsedValue) ? 0 : parsedValue;
    }

    // =========================================
    // カラー / Colors
    // =========================================

    /* 受け付けるカラーコードの書式 / Accepted color code formats */
    var COLOR_CODE_PATTERN = /^(#[0-9A-F]{6}|R\d{1,3}G\d{1,3}B\d{1,3}|C\d{1,3}M\d{1,3}Y\d{1,3}K\d{1,3})$/;

    /* 「カラー指定」を選んだことだけを表す内部キーワード（入力欄の値は書き換えない）
       / Internal keyword meaning "the color code radio is selected", leaving the field untouched */
    var COLOR_CODE_KEYWORD = "colorCode";

    /**
     * ドキュメントのカラースペースに合わせた無彩色を生成する
     * @param {number} rgbLevel - RGBドキュメントでのR・G・Bの値（0〜255）
     * @param {number} cmykBlack - CMYKドキュメントでのKの値（0〜100、C・M・Yは0）
     * @returns {RGBColor|CMYKColor} 無彩色
     */
    function createNeutralColor(rgbLevel, cmykBlack) {
        if (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB) {
            var rgbColor = new RGBColor();
            rgbColor.red = rgbLevel;
            rgbColor.green = rgbLevel;
            rgbColor.blue = rgbLevel;
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = cmykBlack;
        return cmykColor;
    }

    /**
     * ドキュメントのカラースペースに合わせた白を生成する
     * @returns {RGBColor|CMYKColor} 白
     */
    function createWhiteColor() {
        return createNeutralColor(255, 0);
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を生成する
     * @returns {RGBColor|CMYKColor} 黒
     */
    function createBlackColor() {
        return createNeutralColor(0, 100);
    }

    /**
     * カラーコード文字列からカラーを生成する（#RRGGBB / R255G255B255 / C0M100Y100K0）
     * @param {string} colorCode - カラーコード文字列
     * @returns {RGBColor|CMYKColor|null} 生成したカラー。解釈できないときは null
     */
    function createColorFromCode(colorCode) {
        var normalizedCode = normalizeColorCode(colorCode);

        if (/^#[0-9A-F]{6}$/.test(normalizedCode)) {
            var hexColor = new RGBColor();
            hexColor.red = parseInt(normalizedCode.substring(1, 3), 16);
            hexColor.green = parseInt(normalizedCode.substring(3, 5), 16);
            hexColor.blue = parseInt(normalizedCode.substring(5, 7), 16);
            return hexColor;
        }

        var rgbMatch = normalizedCode.match(/^R(\d{1,3})G(\d{1,3})B(\d{1,3})$/);
        if (rgbMatch) {
            var rgbColor = new RGBColor();
            rgbColor.red = Math.min(255, parseInt(rgbMatch[1], 10));
            rgbColor.green = Math.min(255, parseInt(rgbMatch[2], 10));
            rgbColor.blue = Math.min(255, parseInt(rgbMatch[3], 10));
            return rgbColor;
        }

        var cmykMatch = normalizedCode.match(/^C(\d{1,3})M(\d{1,3})Y(\d{1,3})K(\d{1,3})$/);
        if (cmykMatch) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = Math.min(100, parseInt(cmykMatch[1], 10));
            cmykColor.magenta = Math.min(100, parseInt(cmykMatch[2], 10));
            cmykColor.yellow = Math.min(100, parseInt(cmykMatch[3], 10));
            cmykColor.black = Math.min(100, parseInt(cmykMatch[4], 10));
            return cmykColor;
        }

        return null;
    }

    /**
     * カラーコードを正規化する（空白を除き、# の無い6桁16進数には # を補う）
     * @param {string} colorCode - 入力されたカラーコード
     * @returns {string} 正規化したカラーコード
     */
    function normalizeColorCode(colorCode) {
        var normalizedCode = String(colorCode).replace(/\s+/g, "").toUpperCase();
        return /^[0-9A-F]{6}$/.test(normalizedCode) ? "#" + normalizedCode : normalizedCode;
    }

    /**
     * カラーコードとして解釈できる文字列かどうかを判定する
     * @param {string} value - 判定する文字列
     * @returns {boolean} 対応書式なら true
     */
    function isColorCode(value) {
        return (typeof value === "string") && COLOR_CODE_PATTERN.test(normalizeColorCode(value));
    }

    // =========================================
    // 表示状態の保存と復元 / Saving and restoring visibility
    // =========================================

    /**
     * 指定レイヤー以外をすべて非表示にする
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} exceptLayer - 表示したままにするレイヤー
     * @returns {Layer[]} 非表示にしたレイヤー
     */
    function hideOtherLayers(doc, exceptLayer) {
        var hiddenLayers = [];
        for (var i = 0; i < doc.layers.length; i++) {
            var targetLayer = doc.layers[i];
            if (targetLayer === exceptLayer || !targetLayer.visible) continue;
            targetLayer.visible = false;
            hiddenLayers.push(targetLayer);
        }
        return hiddenLayers;
    }

    /**
     * 非表示にしたレイヤーを再表示する
     * @param {Layer[]} hiddenLayers - 非表示にしたレイヤー
     * @returns {void}
     */
    function restoreLayerVisibility(hiddenLayers) {
        for (var i = 0; i < hiddenLayers.length; i++) {
            hiddenLayers[i].visible = true;
        }
    }

    /**
     * 選択オブジェクトを重ね順（前面から背面）で集める
     * app.selection の配列順は重ね順と一致せず、文字ツールでの文字選択（TextRange）も混ざるため使わない
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択オブジェクト（前面から背面の順）
     */
    function collectSelectedItems(doc) {
        var selectedItems = [];
        /* doc.pageItems は前面から背面の順に並ぶ / doc.pageItems runs from front to back */
        for (var i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            /* グループごと選ばれているときは中身を個別に拾わない / Skip children when their group is selected */
            if (item.selected && !hasSelectedAncestor(item)) selectedItems.push(item);
        }
        return selectedItems;
    }

    /**
     * 選択済みの祖先を持つかを調べる
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} 祖先が選択されていれば true
     */
    function hasSelectedAncestor(item) {
        var parentItem = item.parent;
        while (parentItem && parentItem.typename !== "Layer" && parentItem.typename !== "Document") {
            if (parentItem.selected) return true;
            parentItem = parentItem.parent;
        }
        return false;
    }

    /**
     * 選択オブジェクトを指定レイヤーへ複製する（重ね順を維持）
     * @param {PageItem[]} selectedItems - 複製元の選択オブジェクト（前面から背面の順）
     * @param {Layer} targetLayer - 複製先レイヤー
     * @returns {PageItem[]} 複製したオブジェクト
     */
    function duplicateSelectionToLayer(selectedItems, targetLayer) {
        var duplicatedItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            duplicatedItems.push(selectedItems[i].duplicate(targetLayer, ElementPlacement.PLACEATEND));
        }
        /* 前面のものから順に最背面へ送ると、最後には元と同じ重ね順になる
           / Sending each to the back, front-most first, reproduces the original stacking order */
        for (var j = 0; j < duplicatedItems.length; j++) {
            duplicatedItems[j].zOrder(ZOrderMethod.SENDTOBACK);
        }
        return duplicatedItems;
    }

    /**
     * プレビュー用レイヤーを中身ごと削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - 削除するレイヤー名
     * @returns {void}
     */
    function removePreviewLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name !== layerName) continue;
            var targetLayer = doc.layers[i];
            targetLayer.locked = false;
            targetLayer.visible = true;
            targetLayer.remove();
            return;
        }
    }

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
    // 背景・枠線の描画 / Drawing the background and border
    // =========================================

    /* プレビュー・書き出し用に生成した背景・枠線 / The background and border this script created */
    var previewArtwork = [];

    /**
     * 選択オブジェクトの外接範囲を求める
     * テキストは字面の枠ではなく実際の字形で測るため、複製をアウトライン化してから計測する
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]。1つも測れなければ null
     */
    function getSelectionBounds(selectedItems) {
        var temporaryItems = [];
        try {
            var measureTargets = [];
            for (var i = 0; i < selectedItems.length; i++) {
                measureTargets.push(prepareMeasureTarget(selectedItems[i], temporaryItems));
            }
            /* プレビュー境界で測る。クリップグループはマスクで測り、中身が無くなったものは飛ばす
               Measure preview bounds; clip groups by their mask, skipping anything left without geometry */
            return getClipAwareUnionBounds(measureTargets, true);
        } finally {
            /* 途中で失敗しても計測用の一時オブジェクトは必ず片付ける / The temporary artwork goes away whatever fails */
            for (var j = temporaryItems.length - 1; j >= 0; j--) {
                try {
                    temporaryItems[j].remove();
                } catch (e) { /* アウトライン化で無効になった参照は無視 / A reference invalidated by createOutline is fine to skip */ }
            }
        }
    }

    /**
     * 計測に使うオブジェクトを用意する
     * テキストを含むものだけ複製し、複製側のテキストをアウトライン化する（元のオブジェクトは触らない）
     * @param {PageItem} item - 計測したいオブジェクト
     * @param {PageItem[]} temporaryItems - 計測用に作った一時オブジェクトの記録先
     * @returns {PageItem} 計測に使うオブジェクト
     */
    function prepareMeasureTarget(item, temporaryItems) {
        if (!containsText(item)) return item;

        var itemCopy = item.duplicate();

        /* テキスト単体は createOutline() で別のグループに差し替わり、複製側の参照は無効になる。
           無効な参照は比較するだけでもエラーになるので、触らずに戻り値だけを控える
           / createOutline() invalidates the copy's reference, and even comparing it throws, so keep only the result */
        if (itemCopy.typename === "TextFrame") {
            var outlinedItem = outlineText(itemCopy);
            temporaryItems.push(outlinedItem);
            return outlinedItem;
        }

        temporaryItems.push(itemCopy);
        outlineTextsInGroup(itemCopy);
        return itemCopy;
    }

    /**
     * オブジェクトがテキストを含むかを判定する（グループの中も見る）
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} テキストを含むなら true
     */
    function containsText(item) {
        if (item.typename === "TextFrame") return true;
        if (item.typename !== "GroupItem") return false;
        for (var i = 0; i < item.pageItems.length; i++) {
            if (containsText(item.pageItems[i])) return true;
        }
        return false;
    }

    /**
     * テキストをアウトライン化する（複製に対してのみ使う）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {PageItem} アウトライン化したグループ。できなければ元のテキスト
     */
    function outlineText(textFrame) {
        try {
            /* createOutline() は元のテキストを差し替えるので戻り値を使う / createOutline() replaces the frame */
            return textFrame.createOutline();
        } catch (e) {
            /* 空のテキストなどアウトライン化できないものはそのまま測る / Text that cannot be outlined is measured as is */
            return textFrame;
        }
    }

    /**
     * グループの中のテキストをアウトライン化する（複製に対してのみ使う）
     * @param {PageItem} item - 走査するオブジェクト
     * @returns {void}
     */
    function outlineTextsInGroup(item) {
        if (item.typename !== "GroupItem") return;

        /* createOutline() が要素を差し替えて並び順が変わるため、走査前に子を控えておく
           / createOutline() swaps items out and shifts the indexes, so snapshot the children first */
        var children = [];
        for (var i = 0; i < item.pageItems.length; i++) {
            children.push(item.pageItems[i]);
        }
        for (var j = 0; j < children.length; j++) {
            if (children[j].typename === "TextFrame") outlineText(children[j]);
            else outlineTextsInGroup(children[j]);
        }
    }

    /**
     * マージン指定（"上,下,左,右" 形式の "3,3,0,0" など）をpt値に展開する
     * @param {string} marginSpec - マージン指定
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {{top: number, bottom: number, left: number, right: number}} 各辺のマージン（pt）
     */
    function resolveMarginOffsets(marginSpec, unitFactor) {
        var marginValues = String(marginSpec).split(",");
        return {
            top: toNumber(marginValues[0]) * unitFactor,
            bottom: toNumber(marginValues[1]) * unitFactor,
            left: toNumber(marginValues[2]) * unitFactor,
            right: toNumber(marginValues[3]) * unitFactor
        };
    }

    /**
     * 丸めモードから、書き出し範囲を合わせるグリッドの間隔を求める
     * @param {string} roundMode - 丸めモード（pixelGrid / currentUnit / none）
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {number} グリッド間隔（pt）。丸めないときは 0
     */
    function resolveSnapStep(roundMode, unitFactor) {
        if (roundMode === "currentUnit") return unitFactor;
        if (roundMode === "none") return 0;
        /* 既定はピクセルグリッド。倍率100%で 1pt＝1px / Pixel grid by default; at 100% one pt equals one px */
        return 1;
    }

    /**
     * 書き出し範囲を求める（丸めモードに従い、オブジェクトが欠けないよう外側へ広げる）
     * @param {number[]} selectionBounds - 選択オブジェクトの外接範囲 [左, 上, 右, 下]
     * @param {string} marginSpec - マージン指定（上,下,左,右 / 現在の定規単位）
     * @param {string} roundMode - 丸めモード（pixelGrid / currentUnit / none）
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 書き出し範囲
     */
    function buildExportRect(selectionBounds, marginSpec, roundMode, unitFactor) {
        var marginOffsets = resolveMarginOffsets(marginSpec, unitFactor);
        var left = selectionBounds[0] - marginOffsets.left;
        var top = selectionBounds[1] + marginOffsets.top;
        var right = selectionBounds[2] + marginOffsets.right;
        var bottom = selectionBounds[3] - marginOffsets.bottom;

        var snapStep = resolveSnapStep(roundMode, unitFactor);
        if (snapStep > 0) {
            left = Math.floor(left / snapStep) * snapStep;
            top = Math.ceil(top / snapStep) * snapStep;
            right = Math.ceil(right / snapStep) * snapStep;
            bottom = Math.floor(bottom / snapStep) * snapStep;
        }
        return {
            left: left,
            top: top,
            right: right,
            bottom: bottom,
            width: right - left,
            height: top - bottom
        };
    }

    /**
     * 枠線指定（"none" / "0.1,black" など）から線幅（pt、最小1px）を求める
     * @param {string} borderSpec - 枠線指定
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {number} 線幅（pt）。枠線なしなら 0
     */
    function resolveBorderWidth(borderSpec, unitFactor) {
        if (!borderSpec || borderSpec === "none") return 0;
        /* 細くても消えないよう1pxまで切り上げる / Round up so a hairline never disappears */
        var borderWidth = Math.ceil(toNumber(borderSpec.split(",")[0]) * unitFactor);
        return (borderWidth > 0) ? borderWidth : 0;
    }

    /**
     * 枠線指定からカラーを生成する
     * @param {string} borderSpec - 枠線指定
     * @returns {RGBColor|CMYKColor|null} 枠線カラー
     */
    function resolveBorderColor(borderSpec) {
        var colorName = String(borderSpec).split(",")[1];
        if (colorName === "black") return createBlackColor();
        if (colorName === "white") return createWhiteColor();
        return isColorCode(colorName) ? createColorFromCode(colorName) : null;
    }

    /**
     * 背景オブジェクトを生成する（書き出し・プレビュー共用）
     * @param {string} backgroundChoice - 背景指定（transparent / white / black / transparentGrid / #RRGGBB）
     * @param {Object} exportRect - 書き出し範囲
     * @param {number} checkerPercent - 透明グリッドの倍率（%）
     * @returns {PageItem|null} 生成した背景。透過のときは null
     */
    function createExportBackground(backgroundChoice, exportRect, checkerPercent) {
        var doc = app.activeDocument;

        if (backgroundChoice === "transparentGrid") {
            var checkerGroup = registerPreviewArtwork(doc.groupItems.add());
            drawCheckerPattern(checkerGroup, exportRect, CHECKER_TILE_PT * (checkerPercent / 100));
            return checkerGroup;
        }

        var backgroundColor = null;
        if (backgroundChoice === "white") backgroundColor = createWhiteColor();
        else if (backgroundChoice === "black") backgroundColor = createBlackColor();
        else if (isColorCode(backgroundChoice)) backgroundColor = createColorFromCode(backgroundChoice);

        /* 解釈できないカラーコードは描かずに透過のまま見せる / An unreadable color code stays transparent */
        if (!backgroundColor) return null;

        var backgroundRect = registerPreviewArtwork(
            doc.pathItems.rectangle(exportRect.top, exportRect.left, exportRect.width, exportRect.height));
        backgroundRect.filled = true;
        backgroundRect.stroked = false;
        backgroundRect.fillColor = backgroundColor;
        backgroundRect.zOrder(ZOrderMethod.SENDTOBACK);
        return backgroundRect;
    }

    /**
     * 透明グリッド（市松模様）を描画する
     * @param {GroupItem} parentGroup - 追加先グループ
     * @param {Object} exportRect - 書き出し範囲
     * @param {number} tileSize - 1マスの大きさ（pt）
     * @returns {void}
     */
    function drawCheckerPattern(parentGroup, exportRect, tileSize) {
        var grayColor = createNeutralColor(204, 30);
        var whiteColor = createWhiteColor();

        var columnCount = Math.ceil(exportRect.width / tileSize);
        var rowCount = Math.ceil(exportRect.height / tileSize);

        /* マスが細かすぎるときはマスを大きくして描画量を抑える / Enlarge the tiles when there would be too many */
        if (columnCount * rowCount > MAX_CHECKER_TILES) {
            tileSize = tileSize * Math.sqrt(columnCount * rowCount / MAX_CHECKER_TILES);
            columnCount = Math.ceil(exportRect.width / tileSize);
            rowCount = Math.ceil(exportRect.height / tileSize);
        }

        for (var i = 0; i < rowCount; i++) {
            var tileTop = exportRect.top - (i * tileSize);
            /* 端のマスは書き出し範囲の内側で詰める（プレビューで枠からはみ出さないように）
               / Clip the edge tiles to the export area so the preview never spills past its frame */
            var tileHeight = Math.min(tileSize, tileTop - exportRect.bottom);
            for (var j = 0; j < columnCount; j++) {
                var tileLeft = exportRect.left + (j * tileSize);
                var tileRect = parentGroup.pathItems.rectangle(
                    tileTop,
                    tileLeft,
                    Math.min(tileSize, exportRect.right - tileLeft),
                    tileHeight
                );
                tileRect.filled = true;
                tileRect.stroked = false;
                tileRect.fillColor = ((i + j) % 2 === 0) ? grayColor : whiteColor;
            }
        }
        parentGroup.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 書き出し範囲の内側に枠線を描画する
     * @param {Object} exportRect - 書き出し範囲
     * @param {number} borderWidth - 線幅（pt）
     * @param {RGBColor|CMYKColor|null} borderColor - 枠線カラー
     * @returns {PathItem|null} 生成した枠線。描画しないときは null
     */
    function drawBorderRectangle(exportRect, borderWidth, borderColor) {
        if (borderWidth <= 0) return null;
        /* カラーを解釈できないときは、既定の線色で描かずに何も描かない（背景の扱いと揃える）
           / An unreadable color draws nothing rather than inheriting the app default, as the background does */
        if (!borderColor) return null;
        /* 線幅が書き出し範囲より太いと矩形を作れない / A stroke wider than the area leaves no rectangle to draw */
        if (exportRect.width <= borderWidth || exportRect.height <= borderWidth) return null;

        /* 線の中心が範囲の内側に収まるよう半分だけ内側に寄せる / Inset by half the stroke so it stays inside */
        var halfStroke = borderWidth / 2;
        var borderRect = registerPreviewArtwork(app.activeDocument.pathItems.rectangle(
            exportRect.top - halfStroke,
            exportRect.left + halfStroke,
            exportRect.width - borderWidth,
            exportRect.height - borderWidth
        ));
        borderRect.filled = false;
        borderRect.stroked = true;
        borderRect.strokeWidth = borderWidth;
        borderRect.strokeColor = borderColor;
        return borderRect;
    }

    /**
     * 生成した背景・枠線を控えて、あとでまとめて消せるようにする
     * 名前で探すと同名のユーザーオブジェクトまで消してしまうため、参照を持っておく
     * @param {PageItem} item - 生成したオブジェクト
     * @returns {PageItem} 受け取ったオブジェクトをそのまま返す
     */
    function registerPreviewArtwork(item) {
        previewArtwork.push(item);
        return item;
    }

    /**
     * プレビュー・書き出し用に生成した背景と枠線を削除する
     * @returns {void}
     */
    function removePreviewArtwork() {
        for (var i = previewArtwork.length - 1; i >= 0; i--) {
            /* 親ごと削除済みのことがあるため、失敗しても続行 / A parent may already be gone, so keep going */
            try {
                previewArtwork[i].remove();
            } catch (e) {}
        }
        previewArtwork = [];
    }

    // =========================================
    // ファイル名 / Filename
    // =========================================

    /**
     * ファイル名の禁則文字・空白類を区切り文字に置き換える
     * @param {string} fileName - 置換前のファイル名
     * @param {string} delimiter - 区切り文字（"-" または "_"）
     * @returns {string} 置換後のファイル名
     */
    function sanitizeFileName(fileName, delimiter) {
        var replacement = (delimiter === "-") ? "-" : "_";
        /* % と \\ も置換する。File() がパーセントエスケープを復号して別名になるのを防ぐ
           / Replace % and backslash too: File() decodes escapes and would save under a different name */
        return fileName.replace(/[¥\\%\/:*?"<>|\r\n\t　 ]/g, replacement);
    }

    /**
     * 書き出しファイル名を組み立てる（プレビューと書き出しで共用）
     * @param {string} documentBaseName - 拡張子を除いたドキュメント名
     * @param {string} delimiter - 区切り文字
     * @param {string} suffix - 接尾辞
     * @param {boolean} useDocumentName - ドキュメント名を使うかどうか
     * @returns {string} 書き出しファイル名
     */
    function buildExportFileName(documentBaseName, delimiter, suffix, useDocumentName) {
        var baseName = useDocumentName ? documentBaseName : "";
        var suffixWord = suffix ? suffix : DEFAULT_SUFFIX_WORD;
        var suffixDelimiter = suffix ? delimiter : "_";
        /* ドキュメント名を使わないときは区切り文字も出さない（"-400.png" にならないように）
           / Without the document name there is nothing to separate, so the delimiter is dropped */
        var fileName = baseName ? (baseName + suffixDelimiter + suffixWord + ".png") : (suffixWord + ".png");
        return sanitizeFileName(fileName, delimiter);
    }

    // =========================================
    // ダイアログの設定値 / Reading the dialog settings
    // =========================================

    /**
     * 背景の選択状態を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {string} 背景指定（transparent / white / black / transparentGrid / #RRGGBB）
     */
    function getBackgroundChoice(dialogControls) {
        var backgroundControls = dialogControls.background;
        if (backgroundControls.checker.value) return "transparentGrid";
        if (backgroundControls.white.value) return "white";
        if (backgroundControls.black.value) return "black";
        if (backgroundControls.colorCode.value) return normalizeColorCode(backgroundControls.colorCodeInput.text);
        return "transparent";
    }

    /**
     * 透明グリッドの倍率を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {number} 倍率（%）
     */
    function getCheckerPercent(dialogControls) {
        var checkerPercent = toNumber(dialogControls.background.checkerScaleInput.text);
        return (checkerPercent > 0) ? checkerPercent : 100;
    }

    /**
     * マージンの入力値を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {string} マージン指定（上,下,左,右 / 例 "3,3,0,0"）
     */
    function getMarginSpec(dialogControls) {
        var marginControls = dialogControls.margin;
        return [
            toNumber(marginControls.topInput.text),
            toNumber(marginControls.bottomInput.text),
            toNumber(marginControls.leftInput.text),
            toNumber(marginControls.rightInput.text)
        ].join(",");
    }

    /**
     * 丸めモードの選択状態を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {string} 丸めモード（pixelGrid / currentUnit / none）
     */
    function getRoundMode(dialogControls) {
        var marginControls = dialogControls.margin;
        if (marginControls.roundUnit.value) return "currentUnit";
        if (marginControls.roundNone.value) return "none";
        return "pixelGrid";
    }

    /**
     * 枠線の選択状態を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {string} 枠線指定（none / 0.1,black など）
     */
    function getBorderSpec(dialogControls) {
        return dialogControls.border.useBorder.value ? readBorderSpec(dialogControls.border) : "none";
    }

    /**
     * 枠線パネルの入力値から枠線指定を組み立てる
     * @param {Object} borderControls - 枠線のコントロール一式
     * @returns {string} 枠線指定（線幅,カラー）
     */
    function readBorderSpec(borderControls) {
        var borderColorName = "black";
        if (borderControls.white.value) borderColorName = "white";
        else if (borderControls.colorCode.value) borderColorName = normalizeColorCode(borderControls.colorCodeInput.text);
        return borderControls.widthInput.text + "," + borderColorName;
    }

    /**
     * 書き出しサイズの選択状態を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {string} サイズ指定（scale:200 / width:1000 など）
     */
    function getSizeSpec(dialogControls) {
        var sizeControls = dialogControls.size;
        for (var i = 0; i < sizeControls.scaleRadios.length; i++) {
            if (sizeControls.scaleRadios[i].value) return "scale:" + SCALE_CHOICES[i];
        }
        if (sizeControls.customScale.value) return "scale:" + sizeControls.customScaleInput.text;
        if (sizeControls.targetWidth.value) return "width:" + sizeControls.targetWidthInput.text;
        return "scale:100";
    }

    /**
     * 区切り文字の選択状態を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {string} 区切り文字（"-" / "_" / ""）
     */
    function getDelimiter(dialogControls) {
        if (dialogControls.fileName.delimiterDash.value) return "-";
        if (dialogControls.fileName.delimiterUnderscore.value) return "_";
        return "";
    }

    /**
     * 接尾辞の入力値を取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @returns {string} 接尾辞
     */
    function getSuffix(dialogControls) {
        return dialogControls.fileName.useSuffix.value ? dialogControls.fileName.suffixInput.text : "";
    }

    /**
     * 書き出し先フォルダーを取得する
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @param {Document} doc - 対象ドキュメント
     * @returns {Folder} 書き出し先フォルダー
     */
    function getDestinationFolder(dialogControls, doc) {
        if (dialogControls.location.desktop.value) return Folder.desktop;
        /* 未保存でも fullName は返るため、実体があるかで判定してデスクトップへ逃がす
           / Illustrator returns a fullName even when unsaved, so test the file itself */
        try {
            if (doc.fullName.exists) return doc.fullName.parent;
        } catch (e) { /* fullName を取れない書類もデスクトップへ / A document without a usable fullName goes to the desktop */ }
        return Folder.desktop;
    }

    /**
     * ダイアログの入力内容を書き出し設定にまとめる
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} 書き出し設定
     */
    function collectExportSettings(dialogControls, doc) {
        return {
            backgroundChoice: getBackgroundChoice(dialogControls),
            checkerPercent: getCheckerPercent(dialogControls),
            marginSpec: getMarginSpec(dialogControls),
            roundMode: getRoundMode(dialogControls),
            borderSpec: getBorderSpec(dialogControls),
            sizeSpec: getSizeSpec(dialogControls),
            fileName: dialogControls.fileName.previewText.text,
            destinationFolder: getDestinationFolder(dialogControls, doc),
            showFolder: dialogControls.location.showFolder.value
        };
    }

    /**
     * 現在の設定をプリセット定義としてデスクトップに書き出す
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {void}
     */
    function savePresetToFile(dialogControls, unitFactor) {
        /* プリセットはmmで持つ約束なので、現在の定規単位から換算して書き出す
           / Presets are kept in mm, so convert from the current ruler unit on the way out */
        function toMm(value) {
            return toPresetUnit(value, unitFactor);
        }

        var presetName = prompt(getLabel(LABELS.prompt.presetName), getLabel(LABELS.prompt.defaultPresetName));
        if (!presetName) return;

        var today = new Date();
        var dateStamp = today.getFullYear() +
            ("0" + (today.getMonth() + 1)).slice(-2) +
            ("0" + today.getDate()).slice(-2);
        var presetFile = new File(Folder.desktop + "/" + PRESET_FILE_PREFIX + dateStamp + ".txt");
        for (var serialNumber = 2; presetFile.exists; serialNumber++) {
            presetFile = new File(Folder.desktop + "/" + PRESET_FILE_PREFIX + dateStamp + "_" + serialNumber + ".txt");
        }

        /* PRESETS にそのまま貼り込める形で書き出す / Written so it can be pasted straight into PRESETS */
        var presetLines = [
            "{",
            '    label: { ja: "' + presetName + '", en: "' + presetName + '" },',
            '    background: "' + getBackgroundChoice(dialogControls) + '",',
            '    margin: "' + convertMarginSpec(getMarginSpec(dialogControls), toMm) + '",',
            '    round: "' + getRoundMode(dialogControls) + '",',
            '    border: "' + convertBorderSpec(getBorderSpec(dialogControls), toMm) + '",',
            '    location: "' + (dialogControls.location.desktop.value ? "desktop" : "documentFolder") + '",',
            '    delimiter: "' + getDelimiter(dialogControls) + '",',
            '    suffix: "' + getSuffix(dialogControls) + '",',
            '    size: "' + getSizeSpec(dialogControls) + '"',
            "}"
        ];

        /* File の open()・write() は失敗しても例外を投げず false を返す / File.open() and write() return false instead of throwing */
        presetFile.encoding = "UTF-8";
        var isWritten = presetFile.open("w") && presetFile.write(presetLines.join("\n"));
        presetFile.close();
        if (!isWritten) {
            alert(getLabel(LABELS.alert.presetSaveFailed) + presetFile.error);
            return;
        }
        alert(getLabel(LABELS.alert.presetSaved) + presetFile.name);
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* 直前に描いたプレビューの内容。同じなら描き直さない / What the last preview drew; an identical state is skipped */
    var lastPreviewSignature = null;

    /**
     * 現在の設定でプレビュー用の背景・枠線を描き直す
     * 透明グリッドは数千個の矩形を作り直すので、見た目が変わらないときは何もしない
     * @param {Object} dialogControls - ダイアログのコントロール一式
     * @param {number[]} selectionBounds - 選択オブジェクトの外接範囲
     * @param {number} unitFactor - 1単位あたりのpt数
     * @returns {void}
     */
    function renderPreview(dialogControls, selectionBounds, unitFactor) {
        var exportRect = buildExportRect(selectionBounds, getMarginSpec(dialogControls), getRoundMode(dialogControls), unitFactor);
        var backgroundChoice = getBackgroundChoice(dialogControls);
        var checkerPercent = getCheckerPercent(dialogControls);
        var borderSpec = getBorderSpec(dialogControls);

        var signature = [exportRect.left, exportRect.top, exportRect.right, exportRect.bottom,
            backgroundChoice, checkerPercent, borderSpec].join("|");
        if (signature === lastPreviewSignature) return;
        lastPreviewSignature = signature;

        removePreviewArtwork();
        createExportBackground(backgroundChoice, exportRect, checkerPercent);
        drawBorderRectangle(exportRect, resolveBorderWidth(borderSpec, unitFactor), resolveBorderColor(borderSpec));
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 書き出しオプションのダイアログを表示する
     * @param {number[]} selectionBounds - 選択オブジェクトの外接範囲 [左, 上, 右, 下]
     * @param {{code: number, label: string, pointsPerUnit: number}} rulerUnit - 定規の単位情報
     * @param {string} documentBaseName - 拡張子を除いたドキュメント名
     * @returns {Object|null} 書き出し設定。キャンセル時は null
     */
    function showExportOptionsDialog(selectionBounds, rulerUnit, documentBaseName) {
        var doc = app.activeDocument;
        var dialogControls = {};

        /* 現在のマージン設定を含めた書き出し範囲 / Export rect including the current margin */
        function getExportRectFromUI() {
            return buildExportRect(selectionBounds, getMarginSpec(dialogControls), getRoundMode(dialogControls), rulerUnit.pointsPerUnit);
        }

        /* 設定が変わるたびにプレビューと倍率ラベルを描き直す / Redraw the preview and the scale labels on every change */
        function refreshPreview() {
            if (dialogControls.size) updateScaleLabels();
            renderPreview(dialogControls, selectionBounds, rulerUnit.pointsPerUnit);
        }

        /* ファイル名プレビューを更新する / Refresh the filename preview */
        function updateFileNamePreview() {
            dialogControls.fileName.previewText.text = buildExportFileName(
                documentBaseName,
                getDelimiter(dialogControls),
                getSuffix(dialogControls),
                !dialogControls.fileName.ignoreDocName.value
            );
        }

        var exportDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(exportDialog);

        // -----------------------------------------
        // プリセット行 / Preset row
        // -----------------------------------------
        var presetRow = addRow(exportDialog);
        addRowLabel(presetRow, LABELS.fieldLabel.preset);
        var presetNames = [getLabel(LABELS.dropdown.custom)];
        for (var i = 0; i < PRESETS.length; i++) {
            presetNames.push(getLabel(PRESETS[i].label));
        }
        var presetDropdown = presetRow.add("dropdownlist", undefined, presetNames);
        presetDropdown.selection = 0;
        var btnSavePreset = presetRow.add("button", undefined, getLabel(LABELS.button.savePreset));
        btnSavePreset.helpTip = getLabel(LABELS.tooltip.savePreset);

        // -----------------------------------------
        // 2カラム / Two columns
        // -----------------------------------------
        var columnsGroup = exportDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumn = addColumn(columnsGroup);
        var rightColumn = addColumn(columnsGroup);

        dialogControls.background = buildBackgroundPanel(leftColumn);
        dialogControls.margin = buildMarginPanel(leftColumn);
        dialogControls.border = buildBorderPanel(leftColumn);
        dialogControls.size = buildSizePanel(rightColumn);
        dialogControls.fileName = buildFileNamePanel(rightColumn);
        dialogControls.location = buildLocationPanel(rightColumn);

        /**
         * 背景色パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 背景色のコントロール一式
         */
        function buildBackgroundPanel(parentColumn) {
            var backgroundPanel = addPanel(parentColumn, LABELS.panel.background);

            var basicRow = addRow(backgroundPanel);
            var backgroundControls = {
                transparent: basicRow.add("radiobutton", undefined, getLabel(LABELS.radio.transparent)),
                black: basicRow.add("radiobutton", undefined, getLabel(LABELS.radio.black)),
                white: basicRow.add("radiobutton", undefined, getLabel(LABELS.radio.white))
            };

            var checkerRow = addRow(backgroundPanel);
            backgroundControls.checker = checkerRow.add("radiobutton", undefined, labelText(LABELS.radio.checker));
            backgroundControls.checker.helpTip = getLabel(LABELS.tooltip.checker);
            backgroundControls.checkerScaleInput = addNumberField(checkerRow, "100", NUMBER_FIELD_CHARS, refreshPreview);
            backgroundControls.checkerScaleInput.helpTip = numberFieldTip(LABELS.tooltip.checkerScale);
            checkerRow.add("statictext", undefined, "%");

            var colorCodeRow = addRow(backgroundPanel);
            backgroundControls.colorCode = colorCodeRow.add("radiobutton", undefined, labelText(LABELS.radio.colorCode));
            backgroundControls.colorCodeInput = colorCodeRow.add("edittext", undefined, "#ffcc00");
            backgroundControls.colorCodeInput.characters = COLOR_FIELD_CHARS;
            backgroundControls.colorCodeInput.helpTip = getLabel(LABELS.tooltip.colorCode);
            backgroundControls.colorCodeInput.onChange = refreshPreview;

            /* 背景の排他選択と入力欄の有効・無効をまとめて切り替える / Switch the background choice and its fields together */
            backgroundControls.select = function(backgroundChoice) {
                backgroundControls.transparent.value = (backgroundChoice === "transparent");
                backgroundControls.black.value = (backgroundChoice === "black");
                backgroundControls.white.value = (backgroundChoice === "white");
                backgroundControls.checker.value = (backgroundChoice === "transparentGrid");
                /* 決め打ちの4種以外はカラー指定として扱う（読めない値でも選択は外さない）
                   / Anything but the four keywords means the color code, even when it cannot be parsed */
                backgroundControls.colorCode.value = !(backgroundControls.transparent.value || backgroundControls.black.value ||
                    backgroundControls.white.value || backgroundControls.checker.value);
                if (backgroundControls.colorCode.value && backgroundChoice !== COLOR_CODE_KEYWORD) {
                    backgroundControls.colorCodeInput.text = backgroundChoice;
                }
                backgroundControls.colorCodeInput.enabled = backgroundControls.colorCode.value;
                setNumberFieldEnabled(backgroundControls.checkerScaleInput, backgroundControls.checker.value);
            };

            var backgroundChoices = ["transparent", "black", "white", "transparentGrid", COLOR_CODE_KEYWORD];
            var backgroundRadios = [backgroundControls.transparent, backgroundControls.black, backgroundControls.white, backgroundControls.checker, backgroundControls.colorCode];
            for (var i = 0; i < backgroundRadios.length; i++) {
                backgroundRadios[i].onClick = createBackgroundClickHandler(backgroundControls, backgroundChoices[i]);
            }

            backgroundControls.select("transparent");
            return backgroundControls;
        }

        /**
         * 背景ラジオボタンのクリックハンドラーを作る
         * @param {Object} backgroundControls - 背景色のコントロール一式
         * @param {string} backgroundChoice - 選択される背景指定
         * @returns {function} クリックハンドラー
         */
        function createBackgroundClickHandler(backgroundControls, backgroundChoice) {
            return function() {
                backgroundControls.select(backgroundChoice);
                refreshPreview();
            };
        }

        /**
         * マージンパネルを作る（上下左右を3×3に配置し、中央を連動アイコンにする）
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} マージンのコントロール一式
         */
        function buildMarginPanel(parentColumn) {
            var marginPanel = addPanel(parentColumn, LABELS.panel.margin, rulerUnit.label);
            var marginControls = {};

            /* 1行目：［空］［上］［空］/ Row 1: [empty][top][empty] */
            var topRow = addMarginGridRow(marginPanel);
            addMarginGridCell(topRow);
            var topCell = addMarginField(topRow, LABELS.fieldLabel.marginTop, onTopMarginChanged);
            addMarginGridCell(topRow);

            /* 2行目：［左］［連動］［右］/ Row 2: [left][link][right] */
            var middleRow = addMarginGridRow(marginPanel);
            var leftCell = addMarginField(middleRow, LABELS.fieldLabel.marginLeft, refreshPreview);
            var linkCell = addMarginGridCell(middleRow);
            marginControls.linkToggle = addLinkToggle(linkCell, false, onTopMarginChanged);
            marginControls.linkToggle.helpTip = getLabel(LABELS.tooltip.linkMargin);
            var rightCell = addMarginField(middleRow, LABELS.fieldLabel.marginRight, refreshPreview);

            /* 3行目：［空］［下］［空］/ Row 3: [empty][bottom][empty] */
            var bottomRow = addMarginGridRow(marginPanel);
            addMarginGridCell(bottomRow);
            var bottomCell = addMarginField(bottomRow, LABELS.fieldLabel.marginBottom, refreshPreview);
            addMarginGridCell(bottomRow);

            /* サイズの微調整（書き出し範囲の丸め方）/ Size fine-tuning: how the export area is rounded */
            addRowLabel(addRow(marginPanel), LABELS.fieldLabel.roundMode);

            var roundColumn = marginPanel.add("group");
            roundColumn.orientation = "column";
            roundColumn.alignment = ["fill", "top"];
            roundColumn.alignChildren = ["left", "center"];
            roundColumn.spacing = 6;
            marginControls.roundPixel = roundColumn.add("radiobutton", undefined, getLabel(LABELS.roundMode.pixelGrid));
            marginControls.roundUnit = roundColumn.add("radiobutton", undefined, getLabel(LABELS.roundMode.currentUnit));
            marginControls.roundNone = roundColumn.add("radiobutton", undefined, getLabel(LABELS.roundMode.none));
            marginControls.roundPixel.helpTip = getLabel(LABELS.tooltip.roundPixel);
            marginControls.roundUnit.helpTip = getLabel(LABELS.tooltip.roundUnit);
            marginControls.roundNone.helpTip = getLabel(LABELS.tooltip.roundNone);

            /* 丸めモードをUIへ反映する / Apply a round mode to the UI */
            marginControls.selectRoundMode = function(roundMode) {
                marginControls.roundUnit.value = (roundMode === "currentUnit");
                marginControls.roundNone.value = (roundMode === "none");
                /* 未指定・不明な値はピクセルグリッドに寄せる / An unknown or missing value falls back to the pixel grid */
                marginControls.roundPixel.value = !(marginControls.roundUnit.value || marginControls.roundNone.value);
            };

            var roundRadios = [marginControls.roundPixel, marginControls.roundUnit, marginControls.roundNone];
            for (var i = 0; i < roundRadios.length; i++) {
                roundRadios[i].onClick = refreshPreview;
            }
            marginControls.selectRoundMode("pixelGrid");

            marginControls.topInput = topCell.input;
            marginControls.bottomInput = bottomCell.input;
            marginControls.leftInput = leftCell.input;
            marginControls.rightInput = rightCell.input;

            /* 連動中は上の値を残る3辺へ写し、入力できないようにする / While linked, copy the top value and lock the other three */
            function syncLinkedMargins() {
                var isLinked = marginControls.linkToggle.value;
                if (isLinked) {
                    marginControls.bottomInput.text = marginControls.topInput.text;
                    marginControls.leftInput.text = marginControls.topInput.text;
                    marginControls.rightInput.text = marginControls.topInput.text;
                }
                var followerCells = [bottomCell, leftCell, rightCell];
                for (var j = 0; j < followerCells.length; j++) {
                    followerCells[j].group.enabled = !isLinked;
                    /* 自作描画の∧∨を描き直す / redraw the custom-drawn steppers */
                    redrawSteppersIn(followerCells[j].group);
                }
            }

            function onTopMarginChanged() {
                syncLinkedMargins();
                refreshPreview();
            }

            /* マージン指定（"3,3,0,0"）をUIへ反映する / Apply a margin spec to the UI */
            marginControls.select = function(marginSpec) {
                var marginValues = String(marginSpec).split(",");
                marginControls.topInput.text = String(toNumber(marginValues[0]));
                marginControls.bottomInput.text = String(toNumber(marginValues[1]));
                marginControls.leftInput.text = String(toNumber(marginValues[2]));
                marginControls.rightInput.text = String(toNumber(marginValues[3]));
                /* 四辺が同じ値のときだけ連動状態に戻す / Re-link only when all four sides match */
                setLinkToggleValue(marginControls.linkToggle, (marginControls.topInput.text === marginControls.bottomInput.text &&
                    marginControls.topInput.text === marginControls.leftInput.text &&
                    marginControls.topInput.text === marginControls.rightInput.text));
                syncLinkedMargins();
            };

            marginControls.select("0,0,0,0");
            return marginControls;
        }

        /**
         * マージングリッドの1行を作る
         * @param {Panel} parentPanel - 追加先のパネル
         * @returns {Group} 生成した行グループ
         */
        function addMarginGridRow(parentPanel) {
            var rowGroup = parentPanel.add("group");
            rowGroup.orientation = "row";
            /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
            rowGroup.alignment = ["center", "top"];
            rowGroup.alignChildren = ["center", "center"];
            rowGroup.spacing = 0;
            return rowGroup;
        }

        /**
         * マージングリッドの1マス（固定幅の空グループ）を作る
         * @param {Group} parentRow - 追加先の行グループ
         * @returns {Group} 生成したグループ
         */
        function addMarginGridCell(parentRow) {
            var cellGroup = parentRow.add("group");
            cellGroup.orientation = "row";
            cellGroup.alignment = ["center", "center"];
            cellGroup.alignChildren = ["center", "center"];
            cellGroup.spacing = 6;
            /* 入力欄の左に∧∨が入るぶん、どのマスも同じだけ広げてグリッドをそろえる
               Widen every cell by the stepper so the grid stays aligned */
            cellGroup.minimumSize.width = MARGIN_CELL_WIDTH[uiLang] + STEPPER_BUTTON_WIDTH + STEPPER_SIDE_MARGIN;
            return cellGroup;
        }

        /**
         * マージングリッドに項目名＋数値欄のマスを作る
         * @param {Group} parentRow - 追加先の行グループ
         * @param {Object} labelSet - 項目名のラベル定義
         * @param {function} onValueChanged - 値が変わったときに呼ぶコールバック
         * @returns {{group: Group, input: EditText}} マスのグループと入力欄
         */
        function addMarginField(parentRow, labelSet, onValueChanged) {
            var cellGroup = addMarginGridCell(parentRow);
            cellGroup.add("statictext", undefined, labelText(labelSet));
            return { group: cellGroup, input: addNumberField(cellGroup, "0", MARGIN_FIELD_CHARS, onValueChanged) };
        }

        /**
         * 枠線パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 枠線のコントロール一式
         */
        function buildBorderPanel(parentColumn) {
            var borderPanel = addPanel(parentColumn, LABELS.panel.border);

            var widthRow = addRow(borderPanel);
            var borderControls = { useBorder: widthRow.add("checkbox", undefined, labelText(LABELS.fieldLabel.borderWidth)) };
            borderControls.useBorder.helpTip = getLabel(LABELS.tooltip.borderWidth);
            var defaultBorderWidth = getDefaultForUnit(DEFAULT_BORDER_BY_UNIT, rulerUnit.label);
            borderControls.widthInput = addNumberField(widthRow, String(defaultBorderWidth), NUMBER_FIELD_CHARS, refreshPreview);
            borderControls.widthInput.helpTip = numberFieldTip(LABELS.tooltip.borderWidth);
            widthRow.add("statictext", undefined, rulerUnit.label);

            borderControls.colorLabelRow = addRow(borderPanel);
            addRowLabel(borderControls.colorLabelRow, LABELS.fieldLabel.borderColor);

            borderControls.colorRadioRow = addRow(borderPanel);
            borderControls.black = borderControls.colorRadioRow.add("radiobutton", undefined, getLabel(LABELS.radio.black));
            borderControls.white = borderControls.colorRadioRow.add("radiobutton", undefined, getLabel(LABELS.radio.white));

            borderControls.colorCodeRow = addRow(borderPanel);
            borderControls.colorCode = borderControls.colorCodeRow.add("radiobutton", undefined, labelText(LABELS.radio.colorCode));
            borderControls.colorCodeInput = borderControls.colorCodeRow.add("edittext", undefined, "#333333");
            borderControls.colorCodeInput.characters = COLOR_FIELD_CHARS;
            borderControls.colorCodeInput.helpTip = getLabel(LABELS.tooltip.colorCode);
            borderControls.colorCodeInput.onChange = refreshPreview;
            /* 枠線なしで開いてもカラーが未選択にならないようにする / Keep a color selected even when the dialog opens with no border */
            borderControls.black.value = true;

            /* 枠線指定（none / 0.1,black など）をUIへ反映する / Apply a border spec to the UI */
            borderControls.select = function(borderSpec) {
                var hasBorder = (borderSpec !== "none");
                borderControls.useBorder.value = hasBorder;

                if (hasBorder) {
                    var specParts = String(borderSpec).split(",");
                    var borderColorName = specParts[1];
                    borderControls.widthInput.text = specParts[0];
                    borderControls.black.value = (borderColorName === "black");
                    borderControls.white.value = (borderColorName === "white");
                    /* 黒・白以外はカラー指定として扱う / Anything but black or white means the color code */
                    borderControls.colorCode.value = !(borderControls.black.value || borderControls.white.value);
                    if (borderControls.colorCode.value && borderColorName !== COLOR_CODE_KEYWORD) {
                        borderControls.colorCodeInput.text = borderColorName;
                    }
                }

                setNumberFieldEnabled(borderControls.widthInput, hasBorder);
                borderControls.colorLabelRow.enabled = hasBorder;
                borderControls.colorRadioRow.enabled = hasBorder;
                borderControls.colorCodeRow.enabled = hasBorder;
                borderControls.colorCodeInput.enabled = hasBorder && borderControls.colorCode.value;
            };

            borderControls.useBorder.onClick = function() {
                borderControls.select(borderControls.useBorder.value ? readBorderSpec(borderControls) : "none");
                refreshPreview();
            };
            borderControls.black.onClick = createBorderColorHandler(borderControls, "black");
            borderControls.white.onClick = createBorderColorHandler(borderControls, "white");
            borderControls.colorCode.onClick = createBorderColorHandler(borderControls, COLOR_CODE_KEYWORD);

            borderControls.select("none");
            return borderControls;
        }

        /**
         * 枠線カラーのラジオボタンのクリックハンドラーを作る
         * @param {Object} borderControls - 枠線のコントロール一式
         * @param {string} borderColorName - 選択されるカラー名（black / white / COLOR_CODE_KEYWORD）
         * @returns {function} クリックハンドラー
         */
        function createBorderColorHandler(borderControls, borderColorName) {
            return function() {
                borderControls.select(borderControls.widthInput.text + "," + borderColorName);
                refreshPreview();
            };
        }

        /**
         * 書き出しサイズパネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 書き出しサイズのコントロール一式
         */
        function buildSizePanel(parentColumn) {
            var sizePanel = addPanel(parentColumn, LABELS.panel.size);
            var sizeControls = { scaleRadios: [] };

            var scaleColumn = sizePanel.add("group");
            scaleColumn.orientation = "column";
            /* ラベルが伸びても切れないよう、ラジオはパネル幅いっぱいに広げる
               / Stretch the radios to the panel width so a longer label is never clipped */
            scaleColumn.alignment = ["fill", "top"];
            scaleColumn.alignChildren = ["fill", "center"];
            scaleColumn.spacing = 6;

            var exportRect = getExportRectFromUI();
            for (var i = 0; i < SCALE_CHOICES.length; i++) {
                sizeControls.scaleRadios.push(scaleColumn.add("radiobutton", undefined, buildScaleLabel(SCALE_CHOICES[i], exportRect)));
            }

            var customScaleRow = addRow(sizePanel);
            sizeControls.customScale = customScaleRow.add("radiobutton", undefined, labelText(LABELS.fieldLabel.customScale));
            sizeControls.customScale.preferredSize.width = SIZE_RADIO_WIDTH[uiLang];
            sizeControls.customScale.helpTip = getLabel(LABELS.tooltip.customScale);
            sizeControls.customScaleInput = addNumberField(customScaleRow, String(DEFAULT_SCALE), NUMBER_FIELD_CHARS + 1, function(value) {
                sizeControls.select("scale:" + value, false, true);
            });
            sizeControls.customScaleInput.helpTip = numberFieldTip(LABELS.tooltip.customScale);
            customScaleRow.add("statictext", undefined, "%");

            var targetWidthRow = addRow(sizePanel);
            sizeControls.targetWidth = targetWidthRow.add("radiobutton", undefined, labelText(LABELS.fieldLabel.targetWidth));
            sizeControls.targetWidth.preferredSize.width = SIZE_RADIO_WIDTH[uiLang];
            sizeControls.targetWidth.helpTip = getLabel(LABELS.tooltip.targetWidth);
            sizeControls.targetWidthInput = addNumberField(targetWidthRow, "", NUMBER_FIELD_CHARS + 3, function(value) {
                sizeControls.select("width:" + value);
            });
            sizeControls.targetWidthInput.helpTip = numberFieldTip(LABELS.tooltip.targetWidth);
            targetWidthRow.add("statictext", undefined, "px");

            /* サイズ指定（scale:200 / width:1000）をUIへ反映する / Apply a size spec to the UI */
            sizeControls.select = function(sizeSpec, keepSuffix, forceCustomScale) {
                var isWidthMode = (String(sizeSpec).indexOf("width:") === 0);
                var specValue = String(sizeSpec).split(":")[1];
                var matchedScale = false;

                for (var i = 0; i < sizeControls.scaleRadios.length; i++) {
                    /* 倍率指定を選んだときは、値が 1x〜4x と同じでも固定倍率に吸われないようにする
                       / When the custom scale is picked, a matching value must not jump to a fixed radio */
                    var isSelected = !isWidthMode && !forceCustomScale && (specValue === String(SCALE_CHOICES[i]));
                    sizeControls.scaleRadios[i].value = isSelected;
                    setScaleRadioColor(sizeControls.scaleRadios[i], isSelected);
                    if (isSelected) matchedScale = true;
                }

                sizeControls.customScale.value = (!isWidthMode && !matchedScale);
                sizeControls.targetWidth.value = isWidthMode;
                setNumberFieldEnabled(sizeControls.customScaleInput, sizeControls.customScale.value);
                setNumberFieldEnabled(sizeControls.targetWidthInput, isWidthMode);

                if (isWidthMode) {
                    sizeControls.targetWidthInput.text = specValue;
                } else {
                    sizeControls.customScaleInput.text = specValue;
                    sizeControls.targetWidthInput.text = String(Math.ceil(getExportRectFromUI().width * toNumber(specValue) / 100));
                }

                /* 倍率・横幅の値をそのまま接尾辞に流用する / Reuse the scale or width value as the suffix */
                if (!keepSuffix) dialogControls.fileName.setSuffix(specValue);
            };

            for (var j = 0; j < sizeControls.scaleRadios.length; j++) {
                sizeControls.scaleRadios[j].onClick = createScaleClickHandler(sizeControls, SCALE_CHOICES[j]);
            }
            sizeControls.customScale.onClick = function() {
                sizeControls.select("scale:" + sizeControls.customScaleInput.text, false, true);
            };
            sizeControls.targetWidth.onClick = function() {
                sizeControls.select("width:" + sizeControls.targetWidthInput.text);
            };

            return sizeControls;
        }

        /**
         * 倍率ラジオボタンのラベルを組み立てる（書き出しピクセル数を併記）
         * @param {number} scalePercent - 倍率（%）
         * @param {Object} exportRect - 現在のマージンを含めた書き出し範囲
         * @returns {string} ラベル文字列
         */
        function buildScaleLabel(scalePercent, exportRect) {
            var scaleRatio = scalePercent / 100;
            return (scaleRatio + "x") + (uiLang === "ja" ? "：" : ": ") +
                Math.ceil(exportRect.width * scaleRatio) + " × " + Math.ceil(exportRect.height * scaleRatio);
        }

        /**
         * 倍率ラジオボタンのラベルを現在のマージンに合わせて描き直す
         * @returns {void}
         */
        function updateScaleLabels() {
            var exportRect = getExportRectFromUI();
            var labelChanged = false;

            for (var i = 0; i < dialogControls.size.scaleRadios.length; i++) {
                var scaleRadio = dialogControls.size.scaleRadios[i];
                var scaleLabel = buildScaleLabel(SCALE_CHOICES[i], exportRect);
                if (scaleRadio.text !== scaleLabel) {
                    scaleRadio.text = scaleLabel;
                    /* 幅は自動計算に戻す（前のラベル幅のままだと桁が増えたときに切れる）
                       / Reset to auto width, or the control keeps the previous label's width */
                    scaleRadio.preferredSize.width = -1;
                    labelChanged = true;
                }
                setScaleRadioColor(scaleRadio, scaleRadio.value);
            }

            /* 桁が増えてもラベルが切れないよう、変わったときだけ組み直す / Re-layout only when a label changed, so longer text is not clipped */
            if (labelChanged) exportDialog.layout.layout(true);
        }

        /**
         * 倍率ラジオボタンのクリックハンドラーを作る
         * @param {Object} sizeControls - 書き出しサイズのコントロール一式
         * @param {number} scalePercent - 選択される倍率（%）
         * @returns {function} クリックハンドラー
         */
        function createScaleClickHandler(sizeControls, scalePercent) {
            return function() {
                sizeControls.select("scale:" + scalePercent);
            };
        }

        /**
         * 倍率ラジオボタンの文字色を選択状態に合わせる
         * @param {RadioButton} scaleRadio - 対象のラジオボタン
         * @param {boolean} isSelected - 選択中かどうか
         * @returns {void}
         */
        function setScaleRadioColor(scaleRadio, isSelected) {
            var radioGraphics = scaleRadio.graphics;
            radioGraphics.foregroundColor = radioGraphics.newPen(
                radioGraphics.PenType.SOLID_COLOR,
                isSelected ? SCALE_ACTIVE_COLOR : SCALE_INACTIVE_COLOR,
                1
            );
        }

        /**
         * 書き出しファイル名パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} ファイル名のコントロール一式
         */
        function buildFileNamePanel(parentColumn) {
            var fileNamePanel = addPanel(parentColumn, LABELS.panel.fileName);
            var fileNameControls = {};

            var documentNameRow = addRow(fileNamePanel);
            addRowLabel(documentNameRow, LABELS.fieldLabel.documentName).helpTip = getLabel(LABELS.tooltip.documentName);
            fileNameControls.useDocName = documentNameRow.add("radiobutton", undefined, getLabel(LABELS.radio.useDocName));
            fileNameControls.ignoreDocName = documentNameRow.add("radiobutton", undefined, getLabel(LABELS.radio.ignoreDocName));
            fileNameControls.useDocName.value = true;

            var delimiterRow = addRow(fileNamePanel);
            addRowLabel(delimiterRow, LABELS.fieldLabel.delimiter).helpTip = getLabel(LABELS.tooltip.delimiter);
            fileNameControls.delimiterNone = delimiterRow.add("radiobutton", undefined, getLabel(LABELS.radio.none));
            fileNameControls.delimiterDash = delimiterRow.add("radiobutton", undefined, "-");
            fileNameControls.delimiterUnderscore = delimiterRow.add("radiobutton", undefined, "_");

            var suffixRow = addRow(fileNamePanel);
            fileNameControls.useSuffix = suffixRow.add("checkbox", undefined, labelText(LABELS.fieldLabel.suffix));
            fileNameControls.useSuffix.helpTip = getLabel(LABELS.tooltip.suffix);
            fileNameControls.suffixInput = suffixRow.add("edittext", undefined, "");
            fileNameControls.suffixInput.helpTip = getLabel(LABELS.tooltip.suffix);
            fileNameControls.suffixInput.characters = SUFFIX_FIELD_CHARS;
            fileNameControls.suffixInput.enabled = false;

            var previewPanel = fileNamePanel.add("panel");
            previewPanel.alignment = ["fill", "top"];
            previewPanel.margins = [10, 10, 10, 10];
            fileNameControls.previewText = previewPanel.add("statictext", undefined, "");
            fileNameControls.previewText.alignment = ["fill", "center"];
            fileNameControls.previewText.preferredSize.height = FILENAME_ROW_HEIGHT;

            /* 接尾辞を指定してファイル名プレビューを更新する / Set the suffix and refresh the preview */
            fileNameControls.setSuffix = function(suffixValue) {
                fileNameControls.useSuffix.value = true;
                fileNameControls.suffixInput.enabled = true;
                fileNameControls.suffixInput.text = suffixValue;
                updateFileNamePreview();
            };

            /* 接尾辞をなしに戻す / Clear the suffix */
            fileNameControls.clearSuffix = function() {
                fileNameControls.useSuffix.value = false;
                fileNameControls.suffixInput.enabled = false;
                updateFileNamePreview();
            };

            /* 区切り文字を選択する / Select the delimiter */
            fileNameControls.setDelimiter = function(delimiter) {
                fileNameControls.delimiterNone.value = (delimiter === "");
                fileNameControls.delimiterDash.value = (delimiter === "-");
                fileNameControls.delimiterUnderscore.value = (delimiter === "_");
                updateFileNamePreview();
            };

            fileNameControls.useDocName.onClick = updateFileNamePreview;
            fileNameControls.ignoreDocName.onClick = function() {
                /* ドキュメント名を使わないときは接尾辞だけが頼りになる / Without the document name, only the suffix identifies the file */
                fileNameControls.setDelimiter("");
                fileNameControls.setSuffix(fileNameControls.suffixInput.text);
                fileNameControls.suffixInput.active = true;
            };
            fileNameControls.delimiterNone.onClick = updateFileNamePreview;
            fileNameControls.delimiterDash.onClick = updateFileNamePreview;
            fileNameControls.delimiterUnderscore.onClick = updateFileNamePreview;
            fileNameControls.useSuffix.onClick = function() {
                fileNameControls.suffixInput.enabled = fileNameControls.useSuffix.value;
                updateFileNamePreview();
            };
            fileNameControls.suffixInput.onChange = updateFileNamePreview;

            return fileNameControls;
        }

        /**
         * 書き出し先パネルを作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {Object} 書き出し先のコントロール一式
         */
        function buildLocationPanel(parentColumn) {
            var locationPanel = addPanel(parentColumn, LABELS.panel.location);
            var locationControls = {};

            var folderRow = addRow(locationPanel);
            locationControls.desktop = folderRow.add("radiobutton", undefined, getLabel(LABELS.radio.desktop));
            locationControls.documentFolder = folderRow.add("radiobutton", undefined, getLabel(LABELS.radio.documentFolder));
            locationControls.desktop.value = true;

            var showFolderRow = addRow(locationPanel);
            locationControls.showFolder = showFolderRow.add("checkbox", undefined, getLabel(LABELS.checkbox.showFolder));
            locationControls.showFolder.value = true;
            /* フォルダーを開く処理は macOS のみ / Opening the folder is macOS only */
            showFolderRow.visible = (Folder.fs === "Macintosh");

            /* 書き出し先を選択する / Select the destination */
            locationControls.select = function(locationName) {
                locationControls.desktop.value = (locationName === "desktop");
                locationControls.documentFolder.value = (locationName !== "desktop");
            };

            return locationControls;
        }

        /**
         * プリセットの内容をダイアログへ反映する
         * @param {Object} preset - プリセット定義
         * @returns {void}
         */
        function applyPreset(preset) {
            /* プリセットはmmで持っているので、現在の定規単位へ換算して反映する
               / Presets are stored in mm, so convert them to the current ruler unit */
            function toRulerUnit(valueMm) {
                return fromPresetUnit(valueMm, rulerUnit.pointsPerUnit);
            }
            dialogControls.background.select(preset.background);
            dialogControls.margin.select(convertMarginSpec(preset.margin, toRulerUnit));
            dialogControls.margin.selectRoundMode(preset.round);
            dialogControls.border.select(convertBorderSpec(preset.border, toRulerUnit));
            dialogControls.location.select(preset.location);
            dialogControls.size.select(preset.size, true);
            dialogControls.fileName.setDelimiter(preset.delimiter);
            if (preset.suffix) dialogControls.fileName.setSuffix(preset.suffix);
            else dialogControls.fileName.clearSuffix();
            refreshPreview();
        }

        presetDropdown.onChange = function() {
            var selectedIndex = presetDropdown.selection.index;
            /* 先頭は「カスタム」なので何も反映しない / The first entry is "Custom" and applies nothing */
            if (selectedIndex > 0) applyPreset(PRESETS[selectedIndex - 1]);
        };
        btnSavePreset.onClick = function() {
            savePresetToFile(dialogControls, rulerUnit.pointsPerUnit);
        };

        var buttonRow = addButtonRow(exportDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        exportDialog.defaultElement = btnOK;
        exportDialog.cancelElement = btnCancel;

        // -----------------------------------------
        // 初期状態とプレビュー / Initial state and preview
        // -----------------------------------------
        dialogControls.fileName.setDelimiter("-");
        dialogControls.size.select("scale:" + DEFAULT_SCALE);

        refreshPreview();

        exportDialog.onShow = function() {
            updateFileNamePreview();
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(exportDialog, SCRIPT_NAME);
        var dialogResult = exportDialog.show();

        /* プレビュー用の描画は、OK・キャンセルどちらでも必ず消す / Always remove the preview artwork, whichever button was used */
        removePreviewArtwork();

        return (dialogResult === 1) ? collectExportSettings(dialogControls, doc) : null;
    }

    // =========================================
    // 書き出し / Export
    // =========================================

    /**
     * サイズ指定から書き出し倍率を求める
     * @param {string} sizeSpec - サイズ指定（scale:200 / width:1000）
     * @param {number} exportWidthPt - 書き出し範囲の幅（pt）
     * @returns {number} 書き出し倍率（%）
     */
    function resolveExportScale(sizeSpec, exportWidthPt) {
        var specValue = toNumber(String(sizeSpec).split(":")[1]);
        var scalePercent = specValue;

        /* 倍率100%で1pt＝1pxになるため、横幅指定はpt値との比で倍率を求める / At 100% one pt equals one px, so a target width is a ratio against the pt size */
        if (String(sizeSpec).indexOf("width:") === 0) {
            scalePercent = (exportWidthPt > 0) ? (specValue / exportWidthPt * 100) : 100;
        }
        return (scalePercent > 0) ? scalePercent : 100;
    }

    /**
     * 一時アートボードを作ってPNG書き出しする
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} exportSettings - 書き出し設定
     * @param {Object} exportRect - 書き出し範囲
     * @param {{code: number, label: string, pointsPerUnit: number}} rulerUnit - 定規の単位情報
     * @returns {void}
     */
    function exportAsPng(doc, exportSettings, exportRect, rulerUnit) {
        var temporaryArtboardIndex = doc.artboards.length;
        doc.artboards.add([exportRect.left, exportRect.top, exportRect.right, exportRect.bottom]);
        doc.artboards.setActiveArtboardIndex(temporaryArtboardIndex);

        var requestedScale = resolveExportScale(exportSettings.sizeSpec, exportRect.width);
        var exportScale = Math.min(requestedScale, MAX_EXPORT_SCALE);

        /* プレビュー用に描いたものは捨てて、書き出し範囲に合わせて描き直す
           / Drop the preview artwork and redraw it for the final export area */
        removePreviewArtwork();

        /* 背景・枠線・一時アートボードは、書き出しが失敗しても必ず片付ける
           / The background, border and temporary artboard go away even when the export fails */
        try {
            var backgroundItem = createExportBackground(exportSettings.backgroundChoice, exportRect, exportSettings.checkerPercent);
            var borderRect = drawBorderRectangle(
                exportRect,
                resolveBorderWidth(exportSettings.borderSpec, rulerUnit.pointsPerUnit),
                resolveBorderColor(exportSettings.borderSpec)
            );
            if (borderRect) borderRect.zOrder(ZOrderMethod.BRINGTOFRONT);

            var exportOptions = new ExportOptionsPNG24();
            exportOptions.artBoardClipping = true;
            /* 背景を描けなかったときも透過で残す（カラーコードを読めず何も描いていない場合など）
               / Stay transparent whenever no background was actually drawn, e.g. an unreadable color code */
            exportOptions.transparency = !backgroundItem;
            exportOptions.horizontalScale = exportScale;
            exportOptions.verticalScale = exportScale;

            var exportFile = new File(exportSettings.destinationFolder + "/" + exportSettings.fileName);
            try {
                doc.exportFile(exportFile, ExportType.PNG24, exportOptions);
            } catch (e) {
                alert(getLabel(LABELS.alert.exportFailed) + "\n" + e.message);
            }
        } finally {
            removePreviewArtwork();
            doc.artboards.remove(temporaryArtboardIndex);
        }

        if (exportScale < requestedScale) {
            alert(getLabel(LABELS.alert.scaleLimited) + Math.floor(exportScale) + "%");
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = collectSelectedItems(doc);
        if (selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var rulerUnit = getUnitInfo();
        var documentBaseName = doc.name.replace(/\.ai$/i, "");
        var originalArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var hiddenLayers = [];
        var exportSettings = null;

        /* 作業用レイヤーを作った時点から後始末の対象。途中で失敗してもレイヤーを隠したまま放置しない
           / Everything from the working layer on must be undone, whatever fails on the way */
        try {
            /* 選択オブジェクトだけを写した作業用レイヤーで、他のオブジェクトを写り込ませずにプレビューする
               / Work on a copy of the selection so nothing else shows up in the preview */
            var previewLayer = doc.layers.add();
            previewLayer.name = PREVIEW_LAYER_NAME;
            var previewItems = duplicateSelectionToLayer(selectedItems, previewLayer);
            hiddenLayers = hideOtherLayers(doc, previewLayer);

            exportSettings = runExportFlow(doc, getSelectionBounds(previewItems), rulerUnit, documentBaseName);
        } finally {
            /* 作業用レイヤー・レイヤー表示・アクティブアートボードを元に戻す / Undo the temporary layer, visibility and active artboard */
            removePreviewArtwork();
            removePreviewLayerByName(doc, PREVIEW_LAYER_NAME);
            restoreLayerVisibility(hiddenLayers);
            doc.artboards.setActiveArtboardIndex(originalArtboardIndex);
            app.redraw();
        }

        if (exportSettings && exportSettings.showFolder && Folder.fs === "Macintosh") {
            exportSettings.destinationFolder.execute();
        }
    }

    /**
     * ダイアログを開き、確定した設定でPNG書き出しまで行う
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]|null} selectionBounds - 選択オブジェクトの外接範囲
     * @param {{code: number, label: string, pointsPerUnit: number}} rulerUnit - 定規の単位情報
     * @param {string} documentBaseName - 拡張子を除いたドキュメント名
     * @returns {Object|null} 書き出した設定。中止・書き出し不可のときは null
     */
    function runExportFlow(doc, selectionBounds, rulerUnit, documentBaseName) {
        if (!selectionBounds) {
            alert(getLabel(LABELS.alert.invalidSize));
            return null;
        }

        var exportSettings = showExportOptionsDialog(selectionBounds, rulerUnit, documentBaseName);
        if (!exportSettings) return null;

        var exportRect = buildExportRect(selectionBounds, exportSettings.marginSpec, exportSettings.roundMode, rulerUnit.pointsPerUnit);
        if (exportRect.width <= 0 || exportRect.height <= 0) {
            alert(getLabel(LABELS.alert.invalidSize));
            return null;
        }

        exportAsPng(doc, exportSettings, exportRect, rulerUnit);
        return exportSettings;
    }

    main();

})();
