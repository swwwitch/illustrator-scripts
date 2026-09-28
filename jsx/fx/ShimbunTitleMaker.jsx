#target illustrator
#targetengine "ShimbunTitleMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した見出しの左右に二重線を引き、まわりに段組みの新聞ダミーを敷いて、新聞の見出しのような画像を作ります。
新聞ダミーはぼかし、全体を［変形］効果で傾けて、欠けない長方形でマスクします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShimbunTitleMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ndb9bee6b7a2e

### Overview

Draws double rules on both sides of the selected headline and lays newspaper-style dummy columns around it.
The dummy text is blurred, the whole piece is tilted with the Transform effect and masked so no corner is missing.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShimbunTitleMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ShimbunTitleMaker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShimbunTitleMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShimbunTitleMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ndb9bee6b7a2e"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var DEFAULT_VERTICAL_MARGIN_PT = 5;    /* 上下マージンの既定値（pt）/ Default vertical margin */
    var DEFAULT_HORIZONTAL_MARGIN_PT = 8;  /* 左右マージンの既定値（pt）/ Default horizontal margin */
    var DEFAULT_LINE_GAP_PT = 5;         /* 二重線の間隔の既定値（pt）/ Default gap between the two rules */
    var RULE_STROKE_WIDTH_PT = 0.6;      /* 二重線の線幅（pt）/ Stroke width of the rules */
    var DEFAULT_ROTATION_DEG = 3;        /* 回転の既定値（°）/ Default rotation angle */
    var DEFAULT_USE_MASK = true;         /* マスクの既定値 / Default for the mask option */
    var DEFAULT_CHARS_PER_LINE = 11;     /* 新聞ダミーの1行あたりの文字数 / Dummy characters per line */
    var DEFAULT_TIER_COUNT = 3;          /* 新聞ダミーの段数（0で作らない）/ Dummy tiers (0 = none) */
    var DEFAULT_EXTRA_TIER_COUNT = 1;    /* 新聞ダミーの上下に足す段数 / Dummy tiers added above and below */
    var DEFAULT_DUMMY_WIDTH_PT = 85.04;  /* 新聞ダミーの左右それぞれの幅（pt、30mm）/ Dummy width on each side */
    var DUMMY_LEADING_RATIO = 1.4;       /* 新聞ダミーの行送り（文字サイズ比）/ Dummy leading as a ratio of type size */
    var DUMMY_VERTICAL_SCALE = 85;       /* 新聞ダミーの垂直比率（%）/ Dummy vertical scale */
    /* 新聞ダミーのフォント（PostScript 名、前から順に探す）/ Dummy font PostScript names, tried in order */
    var DUMMY_FONT_NAMES = ["HiraMinProN-W3", "HiraMinPro-W3"];
    var DUMMY_RULE_WIDTH_PT = 0.3;       /* 段の区切り罫の線幅（pt）/ Stroke width of the tier rules */
    var DUMMY_BACKGROUND_CMYK = [0, 0, 10, 25]; /* 新聞ダミーの背景色 [C,M,Y,K]（%）/ Dummy background color */
    var RULE_CMYK = [0, 0, 0, 100];      /* 二重線・区切り罫の色 [C,M,Y,K]（%）/ Color of the rules */
    var DEFAULT_BLUR_RADIUS = 8;         /* 新聞ダミーのぼかし（ガウス）の半径（px、0でかけない）/ Dummy blur radius in px (0 = none) */
    var EFFECT_RESOLUTION = 300;         /* 効果のプレビュー解像度（ppi）/ effect preview resolution */
    /* 新聞ダミーの文字列。句点は禁則を「なし」にして行頭にも置く / Dummy text; kinsoku is turned off so periods may start a line */
    var DUMMY_TEXT = "吾輩は猫である。名前はまだ無い。どこで生れたかとんと見当がつかぬ。何でも薄暗い所で泣いていた事だけは記憶している。";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 15;              /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10]; /* パネルの余白 [左,上,右,下] / panel margins */
    var FIELD_LABEL_WIDTH = 104;          /* 項目名の幅 / width of the field labels */
    var FIELD_CHARS = 5;                  /* 数値入力欄の幅（文字数）/ width of the numeric fields */
    var FIELD_ROW_SPACING = 6;            /* 入力行の間隔 / spacing inside a field row */

    /**
     * パネルを縦並び・左揃えにして共通の余白を付ける
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function applyPanelLayout(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "新聞見出しメーカー", en: "Newspaper Headline Maker" }
        },
        panel: {
            sideRules: { ja: "左右の二重線", en: "Side Rules" },
            newspaperDummy: { ja: "新聞ダミー", en: "Newspaper Dummy" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            marginVertical: { ja: "上下の伸び", en: "Extend" },
            marginHorizontal: { ja: "見出しとのアキ", en: "Spacing" },
            lineGap: { ja: "線の間隔", en: "Gap" },
            charsPerLine: { ja: "字詰", en: "Chars per line" },
            tierCount: { ja: "段数", en: "Tiers" },
            extraTierCount: { ja: "上下に追加", en: "Extra tiers" },
            dummyWidth: { ja: "左右の幅", en: "Width" },
            blurRadius: { ja: "ぼかし", en: "Blur" },
            rotation: { ja: "回転", en: "Rotation" }
        },
        checkbox: {
            useMask: { ja: "マスク", en: "Mask" }
        },
        unit: {
            chars: { ja: "文字", en: "chars" },
            tiers: { ja: "段", en: "tiers" }
        },
        tooltip: {
            marginVertical: { ja: "二重線を対象の上端／下端から伸ばす量", en: "How far the rules extend beyond the top and bottom of the object" },
            marginHorizontal: {
                ja: "対象と内側の線との距離。新聞ダミーの行をそろえるため、少し広がることがあります。",
                en: "Distance between the object and the inner rule. It may widen slightly so the dummy lines align."
            },
            lineGap: { ja: "内側の線と外側の線の間隔", en: "Distance between the inner and outer rule" },
            charsPerLine: { ja: "1段の1行に入れる文字数", en: "Characters in one line of a tier" },
            tierCount: {
                ja: "二重線の天地を分ける段数。0で新聞ダミーを作りません。",
                en: "Number of tiers the rule height is split into. 0 adds no dummy text."
            },
            extraTierCount: {
                ja: "上と下にそれぞれ足す、左右の幅いっぱいの段の数。0で足しません。",
                en: "Full-width tiers added above and below each. 0 adds none."
            },
            dummyWidth: { ja: "左右それぞれの新聞ダミーの幅", en: "Width of the dummy text on each side" },
            blurRadius: {
                ja: "新聞ダミーにかけるぼかし（ガウス）の半径（px）。0でかけません。",
                en: "Radius of the Gaussian blur on the dummy text, in pixels. 0 applies none."
            },
            rotation: {
                ja: "全体のグループに［変形］効果で適用する回転角度。0で効果を付けません。",
                en: "Rotation applied to the whole group with the Transform effect. 0 adds no effect."
            },
            useMask: {
                ja: "回転後も角が欠けない長方形で全体をマスクする。［上下に追加］を0にするとOFFになります。",
                en: "Masks the whole piece with a rectangle that leaves no corner missing after the rotation. Turns off when Extra tiers is set to 0."
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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTarget: { ja: "テキストオブジェクトまたはグループを1つ選択してください。", en: "Select one text object or one group." }
        },
        itemName: {
            ruleGroup: { ja: "左右二重線", en: "Side rules" },
            dummyGroup: { ja: "新聞ダミー", en: "Newspaper dummy" }
        }
    };

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
    // 入力の読み取り / Input handling
    // =========================================

    /**
     * 数値の入力文字列を読む。空白を除き、全角のカンマ・小数点を半角にし、
     * 「1,5」のような小数カンマは小数点に、「1,000」のような桁区切りは取り除く
     * @param {string} inputText - 入力された文字列
     * @returns {number} 読み取った数値（読めなければ NaN）
     */
    function parseLocaleNumber(inputText) {
        var normalizedText = String(inputText);
        normalizedText = normalizedText.replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        if (normalizedText.indexOf(",") >= 0 && normalizedText.indexOf(".") < 0) {
            normalizedText = normalizedText.replace(/,/g, ".");
        } else {
            normalizedText = normalizedText.replace(/,/g, "");
        }
        return Number(normalizedText);
    }

    /**
     * 入力欄の値を数値で読む（読めないときは既定値）
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} fallbackValue - 読めなかったときに使う値
     * @param {boolean} allowNegative - 負の値を許すか（許さないときは0に丸める）
     * @returns {number} 読み取った値
     */
    function readFieldAsNumber(inputField, fallbackValue, allowNegative) {
        var enteredValue = parseLocaleNumber(inputField.text);
        if (isNaN(enteredValue)) return fallbackValue;
        if (!allowNegative && enteredValue < 0) return 0;
        return enteredValue;
    }

    /**
     * 入力欄の値（現在の単位）を pt に換算して読む。負の値は0に丸める
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} fallbackPt - 読めなかったときに使う値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {number} 換算した値（pt）
     */
    function readFieldAsPoints(inputField, fallbackPt, pointsPerUnit) {
        var enteredValue = readFieldAsNumber(inputField, NaN, false);
        return isNaN(enteredValue) ? fallbackPt : enteredValue * pointsPerUnit;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確認し、ダイアログでプレビューしながら二重線を描画してグループ化する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var targetItem = getTargetItem(doc.selection);
        if (!targetItem) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        showDialog(doc, targetItem);
    }

    /**
     * 基準にできる選択かを判定して返す（テキスト1つ、またはグループ1つ）
     * @param {Object} currentSelection - ドキュメントの選択
     * @returns {PageItem|null} 基準にするオブジェクト（対象外なら null）
     */
    function getTargetItem(currentSelection) {
        if (!currentSelection || currentSelection.length !== 1) return null;

        var selectedItem = currentSelection[0];
        var isSupported = (
            selectedItem.typename === "TextFrame" ||
            selectedItem.typename === "GroupItem"
        );
        return isSupported ? selectedItem : null;
    }

    /**
     * 二重線の上下端と、内側・外側の線の X 座標を求める
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {{marginVerticalPt: number, marginHorizontalPt: number, lineGapPt: number}} headlineSettings - マージンと間隔（pt）
     * @returns {{topY: number, bottomY: number, innerLeftX: number, innerRightX: number, outerLeftX: number, outerRightX: number}} 二重線の枠
     */
    function getRuleFrame(targetItem, headlineSettings) {
        /* visibleBounds: [左, 上, 右, 下] / visibleBounds: [left, top, right, bottom] */
        var targetBounds = targetItem.visibleBounds;
        var innerLeftX = targetBounds[0] - headlineSettings.marginHorizontalPt;
        var innerRightX = targetBounds[2] + headlineSettings.marginHorizontalPt;
        return {
            topY: targetBounds[1] + headlineSettings.marginVerticalPt,
            bottomY: targetBounds[3] - headlineSettings.marginVerticalPt,
            innerLeftX: innerLeftX,
            innerRightX: innerRightX,
            outerLeftX: innerLeftX - headlineSettings.lineGapPt,
            outerRightX: innerRightX + headlineSettings.lineGapPt
        };
    }

    /**
     * 対象の左右に二重線を描画する
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠
     * @param {{lineGapPt: number}} headlineSettings - 間隔（pt）
     * @returns {GroupItem} 生成した二重線のグループ
     */
    function drawSideRules(doc, targetItem, ruleFrame, headlineSettings) {
        var ruleGroup = targetItem.layer.groupItems.add();
        ruleGroup.name = getLabel("itemName.ruleGroup");

        var strokeColor = createProcessColor(doc, RULE_CMYK);
        var lineXPositions = [ruleFrame.innerLeftX, ruleFrame.innerRightX];
        /* 間隔が0なら内側の線に重なるので、外側の線は描かない / A zero gap would stack the rules, so skip the outer pair */
        if (headlineSettings.lineGapPt > 0) {
            lineXPositions.push(ruleFrame.outerLeftX);
            lineXPositions.push(ruleFrame.outerRightX);
        }
        for (var i = 0; i < lineXPositions.length; i++) {
            createRuleLine(
                ruleGroup,
                [lineXPositions[i], ruleFrame.topY],
                [lineXPositions[i], ruleFrame.bottomY],
                strokeColor,
                RULE_STROKE_WIDTH_PT
            );
        }

        /* 対象のすぐ背面へ / Move the rules just behind the target */
        ruleGroup.move(targetItem, ElementPlacement.PLACEAFTER);
        return ruleGroup;
    }

    /**
     * 新聞ダミーの文字サイズ・行送り・左右の行数を求める。
     * 二重線の天地（A）を段数で分け、段の間は1文字分空ける。
     * 1文字の送り（文字サイズ × 垂直比率）は A ÷（段数 × 字詰 ＋ 段数 − 1）。
     * 左右の幅は行の整数倍に切り詰める
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠
     * @param {{charsPerLine: number, tierCount: number, dummyWidthPt: number}} headlineSettings - 新聞ダミーの設定
     * @returns {Object|null} 新聞ダミーの寸法（作らないときは null）
     */
    function computeDummyLayout(ruleFrame, headlineSettings) {
        var tierCount = headlineSettings.tierCount;
        var charsPerLine = headlineSettings.charsPerLine;
        var ruleHeight = ruleFrame.topY - ruleFrame.bottomY;
        if (tierCount < 1 || charsPerLine < 1 || headlineSettings.dummyWidthPt <= 0 || ruleHeight <= 0) return null;

        /* 縦組みでは垂直比率のぶん字送りが詰まる / in vertical text the vertical scale shortens each character's advance */
        var charAdvance = ruleHeight / (tierCount * charsPerLine + tierCount - 1);
        var fontSize = charAdvance / (DUMMY_VERTICAL_SCALE / 100);
        var dummyLayout = {
            tierCount: tierCount,
            charsPerLine: charsPerLine,
            fontSize: fontSize,
            charAdvance: charAdvance,
            linePitch: fontSize * DUMMY_LEADING_RATIO,
            tierPitch: (charsPerLine + 1) * charAdvance
        };
        dummyLayout.sideLineCount = countFittingLines(dummyLayout, headlineSettings.dummyWidthPt);
        dummyLayout.sideTextWidth = getTextWidth(dummyLayout, dummyLayout.sideLineCount);
        return dummyLayout;
    }

    /**
     * 上下の段の行が左右の段の行と揃うように、二重線を左右へ均等に広げる。
     * 左の本文の右端から右の本文の右端までが行送りの整数倍になればよい
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠（書き換える）
     * @param {Object} dummyLayout - computeDummyLayout() で求めた寸法
     * @returns {void}
     */
    function alignRuleFrameToDummyLines(ruleFrame, dummyLayout) {
        var linePitch = dummyLayout.linePitch;
        /* 外側の線と本文の間は1文字分 / one character of space between the outer rule and the text */
        var lineSpan = (ruleFrame.outerRightX - ruleFrame.outerLeftX) + 2 * dummyLayout.fontSize + dummyLayout.sideTextWidth;
        /* 浮動小数の誤差で1行ぶん余計に広げない / keep float error from adding a whole extra line */
        var alignedSpan = Math.ceil(lineSpan / linePitch - 1e-6) * linePitch;
        var shiftX = (alignedSpan - lineSpan) / 2;

        ruleFrame.innerLeftX -= shiftX;
        ruleFrame.outerLeftX -= shiftX;
        ruleFrame.innerRightX += shiftX;
        ruleFrame.outerRightX += shiftX;
    }

    /**
     * 二重線の外側の左右に段組みの新聞ダミーを描画し、上下には左右の幅いっぱいの段を足す。
     * 段の間は1文字分空けて区切り罫を引く
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @param {Object} ruleFrame - 二重線の枠
     * @param {Object} dummyLayout - computeDummyLayout() で求めた寸法
     * @param {number} extraTierCount - 上下それぞれに足す段数
     * @returns {GroupItem} 生成した新聞ダミーのグループ
     */
    function drawNewspaperDummy(doc, targetLayer, ruleFrame, dummyLayout, extraTierCount) {
        var fontSize = dummyLayout.fontSize;
        var charAdvance = dummyLayout.charAdvance;
        var tierPitch = dummyLayout.tierPitch;
        var sideLineCount = dummyLayout.sideLineCount;
        var sideTextWidth = dummyLayout.sideTextWidth;

        var dummyGroup = targetLayer.groupItems.add();
        dummyGroup.name = getLabel("itemName.dummyGroup");
        var tierContext = {
            targetLayer: targetLayer,
            dummyGroup: dummyGroup,
            charsPerLine: dummyLayout.charsPerLine,
            fontSize: fontSize,
            charAdvance: charAdvance,
            linePitch: dummyLayout.linePitch,
            textFont: findDummyFont(),
            strokeColor: createProcessColor(doc, RULE_CMYK)
        };

        /* 外側の線と本文の間は1文字分 / one character of space between the outer rule and the text */
        var sideTextLefts = [
            ruleFrame.outerLeftX - fontSize - sideTextWidth,
            ruleFrame.outerRightX + fontSize
        ];

        /* 左右の段。2段目以降は上の段との間に罫 / side tiers; a rule above every tier but the first */
        for (var i = 0; i < sideTextLefts.length; i++) {
            for (var k = 0; k < dummyLayout.tierCount; k++) {
                var sideTierTop = ruleFrame.topY - k * tierPitch;
                addDummyTier(tierContext, sideTextLefts[i], sideTierTop, sideLineCount);
                if (k > 0) addTierRule(tierContext, sideTextLefts[i], sideTierTop + charAdvance / 2, sideTextWidth);
            }
        }

        /* 上下の段は左の本文の左端から右の本文の右端まで。alignRuleFrameToDummyLines() で行送りの整数倍になっている
           Tiers above and below span both side regions, already a whole number of lines wide */
        var fullLeft = sideTextLefts[0];
        var fullWidth = sideTextLefts[1] + sideTextWidth - fullLeft;
        var fullLineCount = countFittingLines(tierContext, fullWidth);

        /* どれも内側（ボックス側）の段との間に罫 / each has a rule on the side facing the box */
        for (var j = 1; j <= extraTierCount; j++) {
            var upperTop = ruleFrame.topY + j * tierPitch;
            addDummyTier(tierContext, fullLeft, upperTop, fullLineCount);
            addTierRule(tierContext, fullLeft, upperTop - tierPitch + charAdvance / 2, fullWidth);

            var lowerTop = ruleFrame.bottomY - charAdvance - (j - 1) * tierPitch;
            addDummyTier(tierContext, fullLeft, lowerTop, fullLineCount);
            addTierRule(tierContext, fullLeft, lowerTop + charAdvance / 2, fullWidth);
        }
        return dummyGroup;
    }

    /**
     * 新聞ダミーのフォントを探す（見つからなければ null で既定のフォントのまま）
     * @returns {TextFont|null} フォント
     */
    function findDummyFont() {
        for (var i = 0; i < DUMMY_FONT_NAMES.length; i++) {
            /* getByName は見つからないと例外 / getByName throws when the font is missing */
            try {
                return app.textFonts.getByName(DUMMY_FONT_NAMES[i]);
            } catch (e) {}
        }
        return null;
    }

    /**
     * 指定の幅に収まる行数を返す（最低1行）
     * @param {{fontSize: number, linePitch: number}} tierContext - 段の共通設定（または新聞ダミーの寸法）
     * @param {number} availableWidth - 使える幅（pt）
     * @returns {number} 行数
     */
    function countFittingLines(tierContext, availableWidth) {
        /* ちょうど割り切れる幅が誤差で1行減らないよう、わずかに足す / a tiny epsilon keeps exact fits from losing a line */
        var lineCount = Math.floor((availableWidth - tierContext.fontSize) / tierContext.linePitch + 1e-6) + 1;
        return Math.max(1, lineCount);
    }

    /**
     * 行数ぶんの本文の幅（最初の行の右端から最後の行の左端まで）を返す
     * @param {{fontSize: number, linePitch: number}} tierContext - 段の共通設定
     * @param {number} lineCount - 行数
     * @returns {number} 幅（pt）
     */
    function getTextWidth(tierContext, lineCount) {
        return (lineCount - 1) * tierContext.linePitch + tierContext.fontSize;
    }

    /**
     * 段の区切り罫を1本引く
     * @param {{dummyGroup: GroupItem, strokeColor: Object}} tierContext - 段の共通設定
     * @param {number} ruleLeft - 左端の X 座標
     * @param {number} ruleY - Y 座標
     * @param {number} ruleWidth - 長さ（pt）
     * @returns {void}
     */
    function addTierRule(tierContext, ruleLeft, ruleY, ruleWidth) {
        createRuleLine(
            tierContext.dummyGroup,
            [ruleLeft, ruleY],
            [ruleLeft + ruleWidth, ruleY],
            tierContext.strokeColor,
            DUMMY_RULE_WIDTH_PT
        );
    }

    /**
     * 新聞ダミー1段ぶんの縦組みエリア内文字を作り、新聞ダミーのグループに入れる。
     * 本文がちょうど指定の行数で終わるように、幅と文字数を合わせる
     * @param {{targetLayer: Layer, dummyGroup: GroupItem, charsPerLine: number, fontSize: number, charAdvance: number, linePitch: number, textFont: TextFont}} tierContext - 段の共通設定
     * @param {number} textLeft - 本文の左端の X 座標
     * @param {number} tierTop - 上端の Y 座標
     * @param {number} lineCount - 行数
     * @returns {TextFrame} 生成したテキストフレーム
     */
    function addDummyTier(tierContext, textLeft, tierTop, lineCount) {
        var charsPerLine = tierContext.charsPerLine;
        var charAdvance = tierContext.charAdvance;
        /* 丸めで行や文字がこぼれないよう、幅は左に、高さは下に半端な余裕を足す
           Add slack to the left and bottom so rounding never drops a line or a character */
        var widthSlack = (tierContext.linePitch - tierContext.fontSize) / 2;
        var frameWidth = getTextWidth(tierContext, lineCount) + widthSlack;
        var frameHeight = charsPerLine * charAdvance + charAdvance / 2;

        var targetLayer = tierContext.targetLayer;
        var framePath = targetLayer.pathItems.rectangle(tierTop, textLeft - widthSlack, frameWidth, frameHeight);
        var dummyFrame = targetLayer.textFrames.areaText(framePath, TextOrientation.VERTICAL);
        dummyFrame.contents = buildDummyText(lineCount * charsPerLine);

        /* 句点が行頭に来ても字詰どおりに流す / keep every line at the set length even when a period starts it */
        dummyFrame.textRange.paragraphAttributes.kinsoku = "None";

        var dummyAttributes = dummyFrame.textRange.characterAttributes;
        if (tierContext.textFont) dummyAttributes.textFont = tierContext.textFont;
        dummyAttributes.size = tierContext.fontSize;
        dummyAttributes.verticalScale = DUMMY_VERTICAL_SCALE;
        dummyAttributes.autoLeading = false;
        dummyAttributes.leading = tierContext.linePitch;

        dummyFrame.move(tierContext.dummyGroup, ElementPlacement.PLACEATEND);
        return dummyFrame;
    }

    /**
     * ダミー文字列を指定の文字数まで繰り返す
     * @param {number} charCount - 文字数
     * @returns {string} ダミー文字列
     */
    function buildDummyText(charCount) {
        var dummyText = "";
        while (dummyText.length < charCount) {
            dummyText += DUMMY_TEXT;
        }
        return dummyText.substr(0, charCount);
    }

    /**
     * 対象・二重線・新聞ダミーをグループ化し、新聞ダミーをぼかし、回転があれば［変形］効果で回転する。
     * 最後に、マスクがONなら回転後も欠けない長方形で全体をマスクする
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {Object} headlineSettings - readHeadlineSettings() でまとめた設定
     * @returns {GroupItem} 全体をまとめたグループ（マスクしたときはクリップグループ）
     */
    function buildHeadlineGroup(doc, targetItem, headlineSettings) {
        var ruleFrame = getRuleFrame(targetItem, headlineSettings);
        var dummyLayout = computeDummyLayout(ruleFrame, headlineSettings);
        /* 二重線を描く前に、新聞ダミーの行に合わせて左右の間隔を調整 / adjust the side spacing to the dummy lines before drawing the rules */
        if (dummyLayout) alignRuleFrameToDummyLines(ruleFrame, dummyLayout);

        var ruleGroup = drawSideRules(doc, targetItem, ruleFrame, headlineSettings);
        var dummyGroup = dummyLayout ?
            drawNewspaperDummy(doc, targetItem.layer, ruleFrame, dummyLayout, headlineSettings.extraTierCount) :
            null;
        var headlineGroup = groupHeadlineParts(targetItem, ruleGroup, dummyGroup);

        var blursDummy = dummyGroup && headlineSettings.blurRadius > 0;
        if (dummyGroup) {
            /* ぼかす前の範囲で背景を敷く / lay the background over the unblurred dummy bounds */
            addDummyBackground(doc, headlineGroup, dummyGroup.geometricBounds);
            /* ボックス（対象と二重線）以外の新聞ダミーだけをぼかす / blur only the dummy text, not the title box */
            if (blursDummy) applyGaussianBlur(dummyGroup, headlineSettings.blurRadius);
        }

        /* 効果をかける前の範囲。ぼかしでにじむ縁はマスクの外に出す
           Bounds before the effects; keep the blurred fringe outside the mask */
        var contentBounds = headlineGroup.geometricBounds;
        var blurInset = blursDummy ? headlineSettings.blurRadius : 0;

        if (headlineSettings.rotationDeg !== 0) {
            headlineGroup.applyEffect(buildTransformEffectXml(headlineSettings.rotationDeg));
        }
        if (!headlineSettings.useMask) return headlineGroup;
        return maskWithInscribedRect(headlineGroup, contentBounds, blurInset, headlineSettings.rotationDeg);
    }

    /**
     * 対象の位置にグループを置き、前面から対象・二重線・新聞ダミーの順に入れる
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {GroupItem} ruleGroup - 二重線のグループ
     * @param {GroupItem|null} dummyGroup - 新聞ダミーのグループ（無ければ null）
     * @returns {GroupItem} まとめたグループ
     */
    function groupHeadlineParts(targetItem, ruleGroup, dummyGroup) {
        var headlineGroup = targetItem.layer.groupItems.add();
        headlineGroup.move(targetItem, ElementPlacement.PLACEBEFORE);
        targetItem.move(headlineGroup, ElementPlacement.PLACEATEND);
        ruleGroup.move(headlineGroup, ElementPlacement.PLACEATEND);
        if (dummyGroup) dummyGroup.move(headlineGroup, ElementPlacement.PLACEATEND);
        return headlineGroup;
    }

    /**
     * 新聞ダミー全体を覆う背景色の長方形を、グループの最背面に置く
     * @param {Document} doc - 対象のドキュメント
     * @param {GroupItem} headlineGroup - 追加先のグループ
     * @param {number[]} dummyBounds - 新聞ダミーの範囲 [左, 上, 右, 下]
     * @returns {PathItem} 生成した長方形
     */
    function addDummyBackground(doc, headlineGroup, dummyBounds) {
        var backgroundRect = headlineGroup.pathItems.rectangle(
            dummyBounds[1],
            dummyBounds[0],
            dummyBounds[2] - dummyBounds[0],
            dummyBounds[1] - dummyBounds[3]
        );
        backgroundRect.filled = true;
        backgroundRect.fillColor = createProcessColor(doc, DUMMY_BACKGROUND_CMYK);
        backgroundRect.stroked = false;
        backgroundRect.move(headlineGroup, ElementPlacement.PLACEATEND);
        return backgroundRect;
    }

    /**
     * 回転後の内容に収まる最大の長方形で、グループ全体をマスクする
     * @param {GroupItem} headlineGroup - マスクするグループ
     * @param {number[]} contentBounds - 回転前の範囲 [左, 上, 右, 下]
     * @param {number} edgeInset - 範囲の四辺から内側に詰める量（pt）
     * @param {number} rotationDeg - 回転角度（°）
     * @returns {GroupItem} クリップグループ（範囲が無いときは headlineGroup のまま）
     */
    function maskWithInscribedRect(headlineGroup, contentBounds, edgeInset, rotationDeg) {
        var contentWidth = contentBounds[2] - contentBounds[0] - edgeInset * 2;
        var contentHeight = contentBounds[1] - contentBounds[3] - edgeInset * 2;
        if (contentWidth <= 0 || contentHeight <= 0) return headlineGroup;

        var maskSize = getInscribedRectSize(contentWidth, contentHeight, rotationDeg);
        /* ［変形］効果は中心を基準に回すので、マスクも中心をそろえる / the Transform effect rotates around the center */
        var centerX = (contentBounds[0] + contentBounds[2]) / 2;
        var centerY = (contentBounds[1] + contentBounds[3]) / 2;

        var clipGroup = headlineGroup.layer.groupItems.add();
        clipGroup.move(headlineGroup, ElementPlacement.PLACEBEFORE);
        headlineGroup.move(clipGroup, ElementPlacement.PLACEATEND);

        var maskRect = clipGroup.pathItems.rectangle(
            centerY + maskSize.height / 2,
            centerX - maskSize.width / 2,
            maskSize.width,
            maskSize.height
        );
        /* クリップグループは最前面のパスでマスクする / a clip group masks with its frontmost path */
        maskRect.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        clipGroup.clipped = true;
        return clipGroup;
    }

    /**
     * 幅 × 高さの長方形を回転させたとき、その内側に収まる面積最大の水平な長方形の寸法を返す
     * @param {number} rectWidth - 回転前の幅（pt）
     * @param {number} rectHeight - 回転前の高さ（pt）
     * @param {number} rotationDeg - 回転角度（°）
     * @returns {{width: number, height: number}} 内接する長方形の寸法（pt）
     */
    function getInscribedRectSize(rectWidth, rectHeight, rotationDeg) {
        var angleRad = rotationDeg * Math.PI / 180;
        var sinA = Math.abs(Math.sin(angleRad));
        var cosA = Math.abs(Math.cos(angleRad));
        if (sinA < 1e-10) return { width: rectWidth, height: rectHeight };

        var widthIsLonger = rectWidth >= rectHeight;
        var longSide = widthIsLonger ? rectWidth : rectHeight;
        var shortSide = widthIsLonger ? rectHeight : rectWidth;

        /* 細長い場合や45°付近は、短辺の中点2つで決まる / thin rectangles and angles near 45° are bound by the short side */
        if (shortSide <= 2 * sinA * cosA * longSide || Math.abs(sinA - cosA) < 1e-10) {
            var halfShort = shortSide / 2;
            return widthIsLonger ?
                { width: halfShort / sinA, height: halfShort / cosA } :
                { width: halfShort / cosA, height: halfShort / sinA };
        }
        /* それ以外は4辺すべてが回転した長方形に接する / otherwise all four corners touch the rotated rectangle */
        var cos2A = cosA * cosA - sinA * sinA;
        return {
            width: (rectWidth * cosA - rectHeight * sinA) / cos2A,
            height: (rectHeight * cosA - rectWidth * sinA) / cos2A
        };
    }

    /**
     * ぼかし（ガウス）をライブエフェクトとして適用する
     * @param {PageItem} targetItem - 対象のページアイテム
     * @param {number} blurRadius - ぼかしの半径（px）
     * @returns {void}
     */
    function applyGaussianBlur(targetItem, blurRadius) {
        var effectXml = '<LiveEffect name="Adobe PSL Gaussian Blur">' +
            '<Dict data="R blur ' + blurRadius + ' R PrevDocScale 1 I PrevDres ' + EFFECT_RESOLUTION + ' "/>' +
            '</LiveEffect>';
        targetItem.applyEffect(effectXml);
    }

    /**
     * 中心を基準に回転だけを行う［変形］効果の XML を作る
     * @param {number} rotateDegrees - 回転角度（°、反時計回りが正）
     * @returns {string} LiveEffect の XML
     */
    function buildTransformEffectXml(rotateDegrees) {
        return '<LiveEffect name="Adobe Transform"><Dict data="' +
            'R scaleH_Percent 100' +
            ' R scaleV_Percent 100' +
            ' R scaleH_Factor 1' +
            ' R scaleV_Factor 1' +
            ' R moveH_Pts 0' +
            ' R moveV_Pts 0' +
            ' R rotate_Degrees ' + rotateDegrees +
            ' R rotate_Radians ' + (rotateDegrees * Math.PI / 180) +
            ' I numCopies 0' +
            ' I pinPoint 4' +
            ' B scaleLines 1' +
            ' B transformPatterns 1' +
            ' B transformObjects 1' +
            ' B reflectX 0' +
            ' B reflectY 0' +
            ' B randomize 0' +
            ' "/></LiveEffect>';
    }

    /**
     * 直線を1本描く
     * @param {GroupItem} parentGroup - 追加先のグループ
     * @param {number[]} startPoint - 始点 [X, Y]
     * @param {number[]} endPoint - 終点 [X, Y]
     * @param {Object} strokeColor - 線の色
     * @param {number} strokeWidth - 線幅（pt）
     * @returns {PathItem} 生成したパス
     */
    function createRuleLine(parentGroup, startPoint, endPoint, strokeColor, strokeWidth) {
        var linePath = parentGroup.pathItems.add();
        linePath.setEntirePath([startPoint, endPoint]);
        linePath.stroked = true;
        linePath.filled = false;
        linePath.strokeWidth = strokeWidth;
        linePath.strokeColor = strokeColor;
        linePath.strokeCap = StrokeCap.BUTTENDCAP;
        return linePath;
    }

    /**
     * ドキュメントのカラーモードに合わせた色を CMYK の値から作る。RGB では単純換算した近似色
     * @param {Document} doc - 対象のドキュメント
     * @param {number[]} cmykValues - [C, M, Y, K]（%）
     * @returns {Object} RGBColor または CMYKColor
     */
    function createProcessColor(doc, cmykValues) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var keyRatio = 1 - cmykValues[3] / 100;
            var rgbColor = new RGBColor();
            rgbColor.red = Math.round(255 * (1 - cmykValues[0] / 100) * keyRatio);
            rgbColor.green = Math.round(255 * (1 - cmykValues[1] / 100) * keyRatio);
            rgbColor.blue = Math.round(255 * (1 - cmykValues[2] / 100) * keyRatio);
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cmykValues[0];
        cmykColor.magenta = cmykValues[1];
        cmykColor.yellow = cmykValues[2];
        cmykColor.black = cmykValues[3];
        return cmykColor;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログと各コントロールを作る
     * @param {string} unitLabel - 入力欄に添える単位の表示名
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {{dialog: Window, fieldInputs: Object, maskCheckbox: Checkbox}} ダイアログ、キーごとの入力欄、マスクのチェックボックス
     */
    function buildDialog(unitLabel, pointsPerUnit, onValueChange) {
        var headlineDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        headlineDialog.orientation = "column";
        headlineDialog.alignChildren = ["fill", "top"];
        headlineDialog.margins = DIALOG_MARGINS;

        /* 入力欄はキー名で fieldLabel／tooltip のラベルを引く。pt の既定値は現在の単位に換算して表示
           Each field looks up its fieldLabel/tooltip by key; point defaults are shown in the current unit */
        var fieldInputs = {};
        var maskCheckbox = null;
        var lastExtraTierCount = DEFAULT_EXTRA_TIER_COUNT;

        /**
         * ［上下に追加］が0になったらマスクをOFFにしてから、変更を通知する。0のままでも手動でONに戻せる
         * @returns {void}
         */
        function handleValueChange() {
            /* 空欄は入力途中なので0と見なさない（NaN）/ an empty field is mid-edit, not 0 */
            var extraTierCount = Math.floor(readFieldAsNumber(fieldInputs.extraTierCount, NaN, false));
            if (extraTierCount === 0 && lastExtraTierCount !== 0) maskCheckbox.value = false;
            lastExtraTierCount = extraTierCount;
            onValueChange();
        }

        addNumericFieldPanel(headlineDialog, "panel.sideRules", [
            { key: "marginVertical", initialText: formatUnitValue(DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit), unitLabel: unitLabel },
            { key: "marginHorizontal", initialText: formatUnitValue(DEFAULT_HORIZONTAL_MARGIN_PT, pointsPerUnit), unitLabel: unitLabel },
            { key: "lineGap", initialText: formatUnitValue(DEFAULT_LINE_GAP_PT, pointsPerUnit), unitLabel: unitLabel }
        ], fieldInputs, handleValueChange);

        addNumericFieldPanel(headlineDialog, "panel.newspaperDummy", [
            { key: "charsPerLine", initialText: String(DEFAULT_CHARS_PER_LINE), unitLabel: getLabel("unit.chars"), integer: true, min: 1 },
            { key: "tierCount", initialText: String(DEFAULT_TIER_COUNT), unitLabel: getLabel("unit.tiers"), integer: true },
            { key: "extraTierCount", initialText: String(DEFAULT_EXTRA_TIER_COUNT), unitLabel: getLabel("unit.tiers"), integer: true },
            { key: "dummyWidth", initialText: formatUnitValue(DEFAULT_DUMMY_WIDTH_PT, pointsPerUnit), unitLabel: unitLabel },
            { key: "blurRadius", initialText: String(DEFAULT_BLUR_RADIUS), unitLabel: "px" }
        ], fieldInputs, handleValueChange);

        var optionsPanel = addNumericFieldPanel(headlineDialog, "panel.options", [
            { key: "rotation", initialText: String(DEFAULT_ROTATION_DEG), unitLabel: "°", allowNegative: true }
        ], fieldInputs, handleValueChange);
        maskCheckbox = addCheckboxRow(optionsPanel, "useMask", DEFAULT_USE_MASK && DEFAULT_EXTRA_TIER_COUNT > 0, onValueChange);

        var buttonRow = addButtonRow(headlineDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        return { dialog: headlineDialog, fieldInputs: fieldInputs, maskCheckbox: maskCheckbox };
    }

    /**
     * パネルを作り、数値入力の行を並べる
     * @param {Window} parentDialog - 追加先のダイアログ
     * @param {string} panelLabelPath - パネル見出しのラベルパス
     * @param {Object[]} fieldSpecs - 行の定義（key, initialText, unitLabel, allowNegative, integer, min）
     * @param {Object} fieldInputs - 作った入力欄を key ごとに入れる先
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {Panel} 作ったパネル
     */
    function addNumericFieldPanel(parentDialog, panelLabelPath, fieldSpecs, fieldInputs, onValueChange) {
        var fieldPanel = parentDialog.add("panel", undefined, getLabel(panelLabelPath));
        applyPanelLayout(fieldPanel);
        for (var i = 0; i < fieldSpecs.length; i++) {
            fieldInputs[fieldSpecs[i].key] = addNumericFieldRow(fieldPanel, fieldSpecs[i], onValueChange);
        }
        return fieldPanel;
    }

    /**
     * 「項目名＋数値入力欄＋単位」の行を生成する。項目名と helpTip は fieldLabel.<key> / tooltip.<key> から引く
     * @param {Panel|Group} parentContainer - 追加先
     * @param {{key: string, initialText: string, unitLabel: string, allowNegative: boolean, integer: boolean, min: number}} fieldSpec - 行の定義
     *     （integer は整数の欄、min は下限。負の値を許さない欄の下限は省略時0）
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {EditText} 生成した入力欄
     */
    function addNumericFieldRow(parentContainer, fieldSpec, onValueChange) {
        var fieldRow = parentContainer.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        fieldRow.spacing = FIELD_ROW_SPACING;

        var fieldLabel = fieldRow.add("statictext", undefined, labelText("fieldLabel." + fieldSpec.key));
        fieldLabel.preferredSize = [FIELD_LABEL_WIDTH, -1];
        fieldLabel.justify = "right";

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var inputField;
        var stepOptions = {
            integer: fieldSpec.integer === true,
            /* text の代入では onChanging が発火しないので、増減後に変更を通知する / assigning text does not fire onChanging */
            onStep: function () { onValueChange(); }
        };
        if (fieldSpec.allowNegative !== true) stepOptions.min = (fieldSpec.min !== undefined) ? fieldSpec.min : 0;
        var stepperGroup = addStepper(stepperInputGroup, function () { return inputField; }, stepOptions);

        inputField = stepperInputGroup.add("edittext", undefined, fieldSpec.initialText);
        inputField.characters = FIELD_CHARS;
        inputField.helpTip = getLabel("tooltip." + fieldSpec.key);
        inputField.onChanging = onValueChange;
        bindSteppedArrowKeys(inputField, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */

        fieldRow.add("statictext", undefined, fieldSpec.unitLabel);
        return inputField;
    }

    /**
     * 入力欄の列にそろえてチェックボックスの行を作る。表示名と helpTip は checkbox.<key> / tooltip.<key> から引く
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} checkboxKey - ラベルのキー
     * @param {boolean} initialValue - 初期状態
     * @param {Function} onValueChange - 状態が変わったときに呼ぶ関数
     * @returns {Checkbox} 生成したチェックボックス
     */
    function addCheckboxRow(parentContainer, checkboxKey, initialValue, onValueChange) {
        var checkboxRow = parentContainer.add("group");
        checkboxRow.orientation = "row";
        checkboxRow.alignChildren = ["left", "center"];
        checkboxRow.spacing = FIELD_ROW_SPACING;

        /* 項目名の欄を空けて入力欄の左端にそろえる / leave the label column empty to line up with the fields */
        var labelSpacer = checkboxRow.add("statictext", undefined, "");
        labelSpacer.preferredSize = [FIELD_LABEL_WIDTH, -1];

        var optionCheckbox = checkboxRow.add("checkbox", undefined, getLabel("checkbox." + checkboxKey));
        optionCheckbox.value = initialValue;
        optionCheckbox.helpTip = getLabel("tooltip." + checkboxKey);
        optionCheckbox.onClick = onValueChange;
        return optionCheckbox;
    }

    /**
     * pt の値を現在の単位の表示文字列にする（小数第2位まで）
     * @param {number} valuePt - 値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {string} 入力欄に入れる文字列
     */
    function formatUnitValue(valuePt, pointsPerUnit) {
        var unitValue = valuePt / pointsPerUnit;
        return String(Math.round(unitValue * 100) / 100);
    }

    /**
     * ダイアログを表示し、入力に合わせてプレビューする。
     * プレビューは対象の複製で作り、その間は元の対象を隠しておく。
     * OK で元の対象を使って確定し、キャンセルで元に戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @returns {GroupItem|null} 確定したグループ。キャンセル時は null
     */
    function showDialog(doc, targetItem) {
        var unitInfo = getUnitInfo();
        var pointsPerUnit = unitInfo.pointsPerUnit;
        var previewGroup = null;

        /**
         * 入力欄の値を設定値にまとめる
         * @returns {{marginVerticalPt: number, marginHorizontalPt: number, lineGapPt: number, charsPerLine: number, tierCount: number, extraTierCount: number, dummyWidthPt: number, blurRadius: number, rotationDeg: number, useMask: boolean}} マージン・間隔・幅（pt）、文字数・段数、ぼかし（px）、回転（°）、マスクの有無
         */
        function readHeadlineSettings() {
            var fieldInputs = dialogUI.fieldInputs;
            return {
                marginVerticalPt: readFieldAsPoints(fieldInputs.marginVertical, DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit),
                marginHorizontalPt: readFieldAsPoints(fieldInputs.marginHorizontal, DEFAULT_HORIZONTAL_MARGIN_PT, pointsPerUnit),
                lineGapPt: readFieldAsPoints(fieldInputs.lineGap, DEFAULT_LINE_GAP_PT, pointsPerUnit),
                /* 文字数・段数は整数に切り捨て / character and tier counts are truncated to integers */
                charsPerLine: Math.floor(readFieldAsNumber(fieldInputs.charsPerLine, DEFAULT_CHARS_PER_LINE, false)),
                tierCount: Math.floor(readFieldAsNumber(fieldInputs.tierCount, DEFAULT_TIER_COUNT, false)),
                extraTierCount: Math.floor(readFieldAsNumber(fieldInputs.extraTierCount, DEFAULT_EXTRA_TIER_COUNT, false)),
                dummyWidthPt: readFieldAsPoints(fieldInputs.dummyWidth, DEFAULT_DUMMY_WIDTH_PT, pointsPerUnit),
                blurRadius: readFieldAsNumber(fieldInputs.blurRadius, DEFAULT_BLUR_RADIUS, false),
                rotationDeg: readFieldAsNumber(fieldInputs.rotation, DEFAULT_ROTATION_DEG, true),
                useMask: dialogUI.maskCheckbox.value
            };
        }

        /**
         * 前回のプレビューを削除する
         * @returns {void}
         */
        function removePreview() {
            if (previewGroup) {
                previewGroup.remove();
                previewGroup = null;
            }
        }

        /**
         * 前回のプレビューを消して、対象の複製と現在の入力値で作り直す
         * @returns {void}
         */
        function updatePreview() {
            removePreview();
            /* 複製は隠した元の hidden を引き継ぐので表示に戻す / the duplicate inherits hidden from the original */
            var previewSource = targetItem.duplicate();
            previewSource.hidden = false;
            previewGroup = buildHeadlineGroup(doc, previewSource, readHeadlineSettings());
            app.redraw();
        }

        var dialogUI = buildDialog(unitInfo.label, pointsPerUnit, updatePreview);
        dialogUI.dialog.onShow = updatePreview;

        prepareDialogWindow(dialogUI.dialog, SCRIPT_NAME);
        targetItem.hidden = true;
        var dialogResult = dialogUI.dialog.show();
        removePreview();
        targetItem.hidden = false;

        if (dialogResult !== 1) {
            targetItem.selected = true;
            app.redraw();
            return null;
        }

        var headlineGroup = buildHeadlineGroup(doc, targetItem, readHeadlineSettings());
        doc.selection = [headlineGroup];
        return headlineGroup;
    }

    main();
})();
