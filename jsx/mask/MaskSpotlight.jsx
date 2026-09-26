#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

画像またはベクターオブジェクトのグループと、パスを選択して実行すると、アートワークの複製をパスでマスクして、すりガラス風のぼかし、スポットライト、ルーペ、部分拡大をプレビューしながら作ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MaskSpotlight.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nfc777dda965d

### Overview

With an image or a group of vector objects and a path selected, masks a copy of the artwork with the path and builds a frosted-glass blur, a spotlight, a loupe or a detail callout against a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MaskSpotlight.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MaskSpotlight";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MaskSpotlight.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MaskSpotlight.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nfc777dda965d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // 既定値 / Defaults
    // =========================================
    var DEFAULT_BLUR_RADIUS = 10;   /* ぼかし（ガウス）の半径（px） / Gaussian blur radius in px */
    var MAX_BLUR_RADIUS     = 1000; /* ぼかし（ガウス）の上限（px） / maximum blur radius in px */
    var EFFECT_RESOLUTION   = 300;  /* 効果のプレビュー解像度（ppi） / effect preview resolution */

    /* 効果の種類 / Effect modes */
    var EFFECT_NONE         = "none";      /* 何もしない（マスクだけ） / mask only */
    var EFFECT_BLUR_INSIDE  = "inside";    /* マスク内をぼかす / blur inside the mask */
    var EFFECT_BLUR_OUTSIDE = "outside";   /* マスク外をぼかす / blur outside the mask */
    var EFFECT_SPOTLIGHT    = "spotlight"; /* マスク外を暗くする / darken outside the mask */

    /* スポットライトの設定 / Spotlight settings */
    var DEFAULT_SPOTLIGHT_DARKNESS = 50; /* 暗さ（重ねる黒の不透明度）の初期値（%） / default darkness in percent */

    /* ドロップシャドウの設定 / Drop shadow settings */
    var SHADOW_BLEND_MODE       = 1;    /* 描画モード（1＝乗算）。ダイアログでは変更しない / blend mode (1 = Multiply), not exposed in the dialog */
    var DEFAULT_SHADOW_OPACITY  = 50;   /* 不透明度の初期値（%） / default opacity in percent */
    var DEFAULT_SHADOW_OFFSET_X = 5;    /* X軸オフセットの初期値（pt） / default horizontal offset in points */
    var DEFAULT_SHADOW_OFFSET_Y = 5;    /* Y軸オフセットの初期値（pt） / default vertical offset in points */
    var DEFAULT_SHADOW_BLUR     = 8;    /* ぼかしの初期値（pt） / default blur in points */
    var DEFAULT_SHADOW_DARKNESS = 100;  /* 濃さの初期値（%） / default darkness in percent */
    var MAX_SHADOW_OFFSET       = 1000; /* X/Yオフセットの上限（pt） / offset limit in points */
    var MAX_SHADOW_BLUR         = 1000; /* ぼかしの上限（pt） / blur limit in points */

    /* パスファインダー（効果）の「分割」。グループ全体に適用する
       XMLの書式は AiSmartPathfinderPalette.jsx の buildPathfinderXML() と同じ（Command 5 ＝ 分割）
       The Divide option of the Pathfinder effect, in the same XML form as AiSmartPathfinderPalette.jsx */
    var PATHFINDER_DIVIDE_EFFECT =
        '<LiveEffect name="Adobe Pathfinder" isPre="1"><Dict data="I Command 5' +
        ' B ConvertCustom 1 B ExtractUnpainted 0 R Mix 0.5 R Precision 10 B RemovePoints 0' +
        ' R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0' +
        ' R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">' +
        '<Entry name="DisplayString" value="Divide" valueType="S"/></Dict></LiveEffect>';

    /* 変形（効果）の基準点。3×3グリッドの中央 / The Transform effect's pin point, centre of the 3x3 grid */
    var TRANSFORM_PIN_CENTER = 4;

    /* ルーペの設定 / Loupe settings */
    var DEFAULT_LOUPE_SCALE = 120;  /* 拡大率の初期値（%） / default scale in percent */
    var MIN_LOUPE_SCALE     = 100;  /* 拡大率の下限（%） / minimum scale in percent */
    var MAX_LOUPE_SCALE     = 1000; /* 拡大率の上限（%） / maximum scale in percent */
    var MAX_LOUPE_OFFSET    = 200;  /* X/Yの移動量の上限（マスクサイズ比 %） / offset limit as a share of the mask size */

    /* 「拡大表示のコネクター」の設定。ライブラリーが見つからないときはチェックボックスを出さない
       Zoom connector settings; the checkbox is left out when the library cannot be found */
    var CONNECTOR_LIBRARY_NAME = "graphiclibraryzoomin.ai";   /* グラフィックスタイルのライブラリー（このスクリプトと同じ場所に置く） / graphic style library, kept beside this script */
    var CONNECTOR_STYLE_NAME   = "zoomineffect";              /* 適用するグラフィックスタイルの名前 / name of the graphic style to apply */
    var CONNECTOR_IMPORT_LAYER = "__import_graphic_styles__"; /* 取り込みに使う一時レイヤーの名前 / name of the temporary layer used for the import */

    /* 保存した設定の置き場所 / Where the saved settings live */
    var SETTINGS_FILE = new File(Folder.userData + "/MaskSpotlight/settings.txt");

    /* プリセット。ドロップダウンに並べる順。ダイアログの全項目をここで決める
       Presets, in the order they appear in the dropdown; each one sets every control */
    var PRESETS = [
        {
            labelKey: "preset.foreground",
            effectMode: EFFECT_BLUR_INSIDE, blurRadius: 10, spotlightDarkness: DEFAULT_SPOTLIGHT_DARKNESS,
            addsStroke: false, addsDropShadow: false, addsConnector: false,
            shadowOpacity: DEFAULT_SHADOW_OPACITY, shadowOffsetX: DEFAULT_SHADOW_OFFSET_X, shadowOffsetY: DEFAULT_SHADOW_OFFSET_Y,
            shadowBlur: DEFAULT_SHADOW_BLUR, shadowDarkness: DEFAULT_SHADOW_DARKNESS,
            addsLoupe: false, duplicatesLoupe: false, scalesLoupeByEffect: false,
            loupeScale: 120, loupeOffsetX: 0, loupeOffsetY: 0
        },
        {
            labelKey: "preset.background",
            effectMode: EFFECT_BLUR_OUTSIDE, blurRadius: 15, spotlightDarkness: DEFAULT_SPOTLIGHT_DARKNESS,
            addsStroke: false, addsDropShadow: false, addsConnector: false,
            shadowOpacity: DEFAULT_SHADOW_OPACITY, shadowOffsetX: DEFAULT_SHADOW_OFFSET_X, shadowOffsetY: DEFAULT_SHADOW_OFFSET_Y,
            shadowBlur: DEFAULT_SHADOW_BLUR, shadowDarkness: DEFAULT_SHADOW_DARKNESS,
            addsLoupe: false, duplicatesLoupe: false, scalesLoupeByEffect: false,
            loupeScale: 120, loupeOffsetX: 0, loupeOffsetY: 0
        },
        {
            labelKey: "preset.spotlight",
            effectMode: EFFECT_SPOTLIGHT, blurRadius: 15, spotlightDarkness: DEFAULT_SPOTLIGHT_DARKNESS,
            addsStroke: false, addsDropShadow: false, addsConnector: false,
            shadowOpacity: DEFAULT_SHADOW_OPACITY, shadowOffsetX: DEFAULT_SHADOW_OFFSET_X, shadowOffsetY: DEFAULT_SHADOW_OFFSET_Y,
            shadowBlur: DEFAULT_SHADOW_BLUR, shadowDarkness: DEFAULT_SHADOW_DARKNESS,
            addsLoupe: false, duplicatesLoupe: false, scalesLoupeByEffect: false,
            loupeScale: 120, loupeOffsetX: 0, loupeOffsetY: 0
        },
        {
            labelKey: "preset.maskWithShadow",
            effectMode: EFFECT_NONE, blurRadius: 10, spotlightDarkness: DEFAULT_SPOTLIGHT_DARKNESS,
            addsStroke: false, addsDropShadow: true, addsConnector: false,
            shadowOpacity: DEFAULT_SHADOW_OPACITY, shadowOffsetX: DEFAULT_SHADOW_OFFSET_X, shadowOffsetY: DEFAULT_SHADOW_OFFSET_Y,
            shadowBlur: DEFAULT_SHADOW_BLUR, shadowDarkness: DEFAULT_SHADOW_DARKNESS,
            addsLoupe: false, duplicatesLoupe: false, scalesLoupeByEffect: false,
            loupeScale: 120, loupeOffsetX: 0, loupeOffsetY: 0
        },
        {
            labelKey: "preset.loupe",
            effectMode: EFFECT_BLUR_OUTSIDE, blurRadius: 15, spotlightDarkness: DEFAULT_SPOTLIGHT_DARKNESS,
            addsStroke: false, addsDropShadow: true, addsConnector: false,
            shadowOpacity: DEFAULT_SHADOW_OPACITY, shadowOffsetX: DEFAULT_SHADOW_OFFSET_X, shadowOffsetY: DEFAULT_SHADOW_OFFSET_Y,
            shadowBlur: DEFAULT_SHADOW_BLUR, shadowDarkness: DEFAULT_SHADOW_DARKNESS,
            addsLoupe: true, duplicatesLoupe: false, scalesLoupeByEffect: true,
            loupeScale: 120, loupeOffsetX: 0, loupeOffsetY: 0
        },
        {
            /* 拡大コピーをすぐ右下へ置く（100%＝マスク1つ分。少し外へ逃がして重なりを避ける）
               A magnified copy set just to the lower right, a little past one mask size to clear the original */
            labelKey: "preset.callout",
            effectMode: EFFECT_NONE, blurRadius: 15, spotlightDarkness: DEFAULT_SPOTLIGHT_DARKNESS,
            addsStroke: true, addsDropShadow: true, addsConnector: true,
            shadowOpacity: DEFAULT_SHADOW_OPACITY, shadowOffsetX: DEFAULT_SHADOW_OFFSET_X, shadowOffsetY: DEFAULT_SHADOW_OFFSET_Y,
            shadowBlur: DEFAULT_SHADOW_BLUR, shadowDarkness: DEFAULT_SHADOW_DARKNESS,
            addsLoupe: true, duplicatesLoupe: true, scalesLoupeByEffect: true,
            loupeScale: 160, loupeOffsetX: 110, loupeOffsetY: 110
        },
        {
            /* 上と同じ並べ方で、背景をぼかす / the same layout, with the background blurred */
            labelKey: "preset.calloutBlurred",
            effectMode: EFFECT_BLUR_OUTSIDE, blurRadius: 15, spotlightDarkness: DEFAULT_SPOTLIGHT_DARKNESS,
            addsStroke: true, addsDropShadow: true, addsConnector: true,
            shadowOpacity: DEFAULT_SHADOW_OPACITY, shadowOffsetX: DEFAULT_SHADOW_OFFSET_X, shadowOffsetY: DEFAULT_SHADOW_OFFSET_Y,
            shadowBlur: DEFAULT_SHADOW_BLUR, shadowDarkness: DEFAULT_SHADOW_DARKNESS,
            addsLoupe: true, duplicatesLoupe: true, scalesLoupeByEffect: true,
            loupeScale: 160, loupeOffsetX: 110, loupeOffsetY: 110
        }
    ];

    /* ダイアログを開いたときに選ぶプリセット。保存した設定があればそちらが優先される
       The preset selected when the dialog opens, unless there are saved settings */
    var DEFAULT_PRESET_KEY = "preset.background";

    // =========================================
    // ダイアログの寸法 / Dialog metrics
    // =========================================
    var NUMBER_LABEL_WIDTH = 74;  /* 数値欄のラベル幅（px）。「不透明度：」が切れない幅 / label width of a number row, wide enough for the longest label */
    var SLIDER_LABEL_WIDTH = 30;  /* スライダー行のラベル幅（px）。日本語の全角コロンが切れない幅 / label width of a slider row, wide enough for the Japanese colon */
    var SLIDER_WIDTH       = 130; /* スライダーの幅（px） / slider width */
    var SAVE_BUTTON_WIDTH  = 60;  /* ［保存］ボタンの幅（px） / width of the Save button */

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
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "すりガラス・スポットライト・ルーペ", en: "Frosted Glass, Spotlight and Loupe" }
        },
        panel: {
            effect:     { ja: "効果", en: "Effect" },
            option:     { ja: "オプション", en: "Options" },
            loupe:      { ja: "ルーペ", en: "Loupe" },
            dropShadow: { ja: "ドロップシャドウ", en: "Drop shadow" }
        },
        label: {
            preset:            { ja: "プリセット", en: "Preset" },
            blurRadius:        { ja: "半径", en: "Radius" },
            spotlightDarkness: { ja: "暗さ", en: "Darkness" },
            loupeScale:        { ja: "拡大率", en: "Scale" },
            offsetX:           { ja: "X", en: "X" },
            offsetY:           { ja: "Y", en: "Y" },
            shadowOpacity:     { ja: "不透明度", en: "Opacity" },
            shadowBlur:        { ja: "ぼかし", en: "Blur" },
            shadowDarkness:    { ja: "濃さ", en: "Darkness" }
        },
        preset: {
            foreground:     { ja: "すりガラス", en: "Frosted glass" },
            background:     { ja: "背景をぼかす", en: "Blurred background" },
            spotlight:      { ja: "スポットライト", en: "Spotlight" },
            loupe:          { ja: "ルーペ", en: "Loupe" },
            callout:        { ja: "部分拡大", en: "Detail callout" },
            calloutBlurred: { ja: "部分拡大（背景をぼかす）", en: "Detail callout, blurred background" },
            maskWithShadow: { ja: "マスク＋ドロップシャドウ", en: "Mask with drop shadow" },
            custom:         { ja: "カスタム", en: "Custom" },
            saved:          { ja: "保存した設定", en: "Saved settings" }
        },
        radio: {
            effectNone:  { ja: "なし", en: "None" },
            blurInside:  { ja: "ぼかし（内）", en: "Blur (inside)" },
            blurOutside: { ja: "ぼかし（外）", en: "Blur (outside)" },
            spotlight:   { ja: "スポットライト", en: "Spotlight" }
        },
        checkbox: {
            addStroke:      { ja: "線を追加", en: "Add stroke" },
            addConnector:   { ja: "引き出し線でつなぐ", en: "Link with callout lines" },
            dropShadow:     { ja: "ドロップシャドウを追加", en: "Add drop shadow" },
            addLoupe:       { ja: "ルーペを追加", en: "Add loupe" },
            duplicateLoupe: { ja: "元を残して複製", en: "Duplicate (keep original)" },
            scaleByEffect:  { ja: "変形効果で拡大", en: "Scale with Transform effect" },
            preview:        { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            preset:            { ja: "よく使う設定をまとめて呼び出します。設定を変えると「カスタム」になります。", en: "Loads a set of settings at once. Changing anything switches it to Custom." },
            save:              { ja: "今の設定を保存します。次に開いたときは「保存した設定」として選べます。", en: "Saves the current settings. They appear as Saved settings the next time this dialog opens." },
            blurRadius:        { ja: "ぼかし（ガウス）の半径（px）です。「ぼかし（内）」「ぼかし（外）」のときだけ使えます。", en: "Radius of the Gaussian blur, in pixels. Available only for Blur (inside) and Blur (outside)." },
            effectNone:        { ja: "ぼかしも暗くする処理もしません。マスクを作るだけで、線やドロップシャドウは使えます。", en: "Applies neither blur nor darkening. It only makes the mask; the stroke and drop shadow still apply." },
            blurInside:        { ja: "マスクの内側をぼかします。背景はそのまま残ります。このときルーペは使えません。", en: "Blurs the inside of the mask and leaves the background sharp. The loupe cannot be used with it." },
            blurOutside:       { ja: "背景側をぼかし、マスクの内側をくっきり残します。", en: "Blurs the background and keeps the inside of the mask sharp." },
            spotlight:         { ja: "マスクの外側に黒を重ねて暗くし、マスクの内側だけを明るく残します。ぼかしは使いません。", en: "Lays black over the area outside the mask so only the inside stays bright. No blur is applied." },
            spotlightDarkness: { ja: "マスク外に重ねる黒の不透明度（%）です。", en: "Opacity of the black laid over the area outside the mask, in percent." },
            addStroke:         { ja: "できあがったクリップグループに線を追加します。", en: "Adds a stroke to the clipping group." },
            addConnector:      { ja: "元の位置と拡大したコピーを線でつなぐグラフィックスタイルを適用します。ルーペを「元を残して複製」するときだけ使えます。", en: "Applies a graphic style that links the artwork and the magnified copy with lines. Available only where the loupe keeps the original and duplicates it, as in the Detail callout presets." },
            dropShadow:        { ja: "できあがったクリップグループにドロップシャドウを適用します。", en: "Applies a drop shadow to the clipping group." },
            shadowOpacity:     { ja: "影の不透明度（%）です。", en: "Opacity of the shadow, in percent." },
            shadowOffsetX:     { ja: "影を左右にずらす量（pt）です。正の値で右へ動きます。", en: "How far the shadow shifts sideways, in points. Positive values move it right." },
            shadowOffsetY:     { ja: "影を上下にずらす量（pt）です。正の値で下へ動きます。", en: "How far the shadow shifts vertically, in points. Positive values move it down." },
            shadowBlur:        { ja: "影のぼかし（pt）です。", en: "Blur of the shadow, in points." },
            shadowDarkness:    { ja: "影の濃さ（%）です。カラー指定ではなく「濃さ」で描きます。", en: "Darkness of the shadow, in percent. The shadow uses darkness rather than a color." },
            addLoupe:          { ja: "マスクを拡大して、拡大鏡のように重ねます。効果が「ぼかし（内）」以外のときに使えます。", en: "Scales the mask up and lays it over the artwork like a magnifier. Available unless the blur goes inside the mask." },
            duplicateLoupe:    { ja: "元のマスクを残したままルーペを増やします。「変形効果で拡大」がオンなら、実体は増やさず［変形］効果のコピーで作ります。", en: "Keeps the original mask and adds the loupe beside it. With Scale with Transform effect on, the copy comes from the Transform effect instead of a real duplicate." },
            scaleByEffect:     { ja: "拡大と移動を［変形］効果として適用します。オブジェクトは変形されず、あとから数値を直せます。", en: "Applies the scaling and the move as a Transform effect, leaving the object itself untouched and editable afterwards." },
            loupeScale:        { ja: "ルーペの拡大率（%）です。", en: "How far the loupe is scaled up, in percent." },
            loupeOffsetX:      { ja: "ルーペを左右に動かします。マスクの幅に対する割合（%）です。", en: "Moves the loupe sideways, as a share of the mask width." },
            loupeOffsetY:      { ja: "ルーペを上下に動かします。マスクの高さに対する割合（%）で、下方向が正です。", en: "Moves the loupe up and down, as a share of the mask height. Positive values move it down." },
            resetPosition:     { ja: "XとYを0に戻します。", en: "Sets X and Y back to zero." },
            preview:           { ja: "設定の結果をドキュメント上で確認します。", en: "Shows the result on the document while you adjust the settings." },
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
        alert: {
            noDocument:           { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            needArtworkAndPath:   { ja: "画像またはグループと、パスを1つずつ選択してください。", en: "Select one image or group and one path." },
            connectorUnavailable: { ja: "コネクターのグラフィックスタイルを読み込めませんでした：", en: "The connector's graphic style could not be loaded:" }
        },
        button: {
            save:          { ja: "保存", en: "Save" },
            ok:            { ja: "OK", en: "OK" },
            cancel:        { ja: "キャンセル", en: "Cancel" },
            resetPosition: { ja: "位置をリセット", en: "Reset position" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.effect" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * 単位を渡すと「半径（px）：」のようにコロンの前へ入れる
     * @param {string} labelPath - ドット区切りキー
     * @param {string} [unitText] - 項目名に添える単位
     * @returns {string} コロン付きのラベル
     */
    function labelText(labelPath, unitText) {
        var unitSuffix = "";
        if (unitText) {
            unitSuffix = (uiLang === "ja") ? "（" + unitText + "）" : " (" + unitText + ")";
        }
        return getLabel(labelPath) + unitSuffix + ((uiLang === "ja") ? "：" : ":");
    }

    // =========================================
    // 処理対象 / Targets
    // =========================================
    var doc = null;
    var originalArtwork = null;  /* 選択した画像またはグループ（A） / the selected image or group (A) */
    var originalMaskPath = null; /* 選択したマスク用パス（C） / the selected mask path (C) */

    /* プレビューのためだけに作った／隠したアイテム。確定実行では空のまま
       Items created or hidden only for the preview; the final run leaves these empty */
    var previewOnlyItems = [];
    var previewHiddenItems = [];

    /* マスクされる側に使えるアイテムの種類 / Item types that can be masked */
    var MASKABLE_TYPES = ["PlacedItem", "RasterItem", "GroupItem"];

    /* マスクに使えるパスの種類 / Path types that can serve as a mask */
    var MASK_PATH_TYPES = ["PathItem", "CompoundPathItem"];

    /**
     * アイテムの種類が一覧のどれかに当てはまるか調べる
     * @param {object} targetItem - 調べるページアイテム
     * @param {Array} typeNames - 種類名（typename）の配列
     * @returns {boolean} 当てはまれば true
     */
    function matchesTypename(targetItem, typeNames) {
        for (var i = 0; i < typeNames.length; i++) {
            if (targetItem.typename === typeNames[i]) return true;
        }
        return false;
    }

    /**
     * 選択範囲からマスクされるアイテムとマスク用パスを取り出す
     * 該当するものが何個あったかも返し、1つずつでなければ呼び出し側ではじけるようにする
     * @param {Array} selectedItems - 選択アイテムの配列
     * @returns {object} artwork、maskPath と、それぞれの個数を持つオブジェクト
     */
    function pickArtworkAndMaskPath(selectedItems) {
        var pickedItems = { artwork: null, maskPath: null, artworkCount: 0, maskPathCount: 0 };
        if (!selectedItems) return pickedItems;

        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (matchesTypename(selectedItem, MASKABLE_TYPES)) {
                pickedItems.artworkCount++;
                if (!pickedItems.artwork) pickedItems.artwork = selectedItem;
            } else if (matchesTypename(selectedItem, MASK_PATH_TYPES)) {
                pickedItems.maskPathCount++;
                if (!pickedItems.maskPath) pickedItems.maskPath = selectedItem;
            }
        }
        return pickedItems;
    }

    /* 設定ファイルに書き出す項目。boolean は 1／0 で持つ
       The keys written to the settings file; booleans are stored as 1 and 0 */
    var SETTING_NUMBER_KEYS = ["blurRadius", "spotlightDarkness", "shadowOpacity", "shadowOffsetX", "shadowOffsetY",
        "shadowBlur", "shadowDarkness", "loupeScale", "loupeOffsetX", "loupeOffsetY"];
    var SETTING_FLAG_KEYS = ["addsStroke", "addsDropShadow", "addsConnector", "addsLoupe", "duplicatesLoupe", "scalesLoupeByEffect"];

    /**
     * 今の設定をファイルへ書き出す
     * @param {object} settings - 保存する設定
     * @returns {boolean} 書き出せたら true
     */
    function saveSettingsToFile(settings) {
        try {
            var settingsFolder = SETTINGS_FILE.parent;
            if (!settingsFolder.exists) settingsFolder.create();

            var settingLines = ["effectMode=" + settings.effectMode];
            var i;
            for (i = 0; i < SETTING_NUMBER_KEYS.length; i++) {
                settingLines.push(SETTING_NUMBER_KEYS[i] + "=" + settings[SETTING_NUMBER_KEYS[i]]);
            }
            for (i = 0; i < SETTING_FLAG_KEYS.length; i++) {
                settingLines.push(SETTING_FLAG_KEYS[i] + "=" + (settings[SETTING_FLAG_KEYS[i]] ? 1 : 0));
            }

            if (!SETTINGS_FILE.open("w")) return false;
            SETTINGS_FILE.write(settingLines.join("\n"));
            SETTINGS_FILE.close();
            return true;
        } catch (e) {
            /* 書き込めない環境では保存しないだけ / just skip saving where the file cannot be written */
            return false;
        }
    }

    /**
     * 保存した設定を読み込み、プリセットと同じ形で返す
     * @returns {object} 読み込んだ設定。無ければ null
     */
    function loadSavedPreset() {
        if (!SETTINGS_FILE.exists) return null;

        var savedText = "";
        try {
            if (!SETTINGS_FILE.open("r")) return null;
            savedText = SETTINGS_FILE.read();
            SETTINGS_FILE.close();
        } catch (e) {
            /* 壊れていたり読めないときは無かったことにする / treat an unreadable file as none */
            return null;
        }

        var storedValues = {};
        var settingLines = savedText.split("\n");
        for (var i = 0; i < settingLines.length; i++) {
            var separatorIndex = settingLines[i].indexOf("=");
            if (separatorIndex > 0) {
                storedValues[settingLines[i].substring(0, separatorIndex)] = settingLines[i].substring(separatorIndex + 1);
            }
        }
        /* v1.0.0 では blurArea という名前で書き出していた / v1.0.0 wrote the same value as blurArea */
        var storedEffectMode = storedValues.effectMode || storedValues.blurArea;
        if (!storedEffectMode) return null;

        var savedPreset = { labelKey: "preset.saved", effectMode: storedEffectMode };
        var j;
        for (j = 0; j < SETTING_NUMBER_KEYS.length; j++) {
            /* 項目が欠けている（古い設定ファイル）ときは先頭のプリセットの値で埋める
               A key the file does not carry falls back to the first preset's value */
            var storedNumber = Number(storedValues[SETTING_NUMBER_KEYS[j]]);
            savedPreset[SETTING_NUMBER_KEYS[j]] = isNaN(storedNumber) ? PRESETS[0][SETTING_NUMBER_KEYS[j]] : storedNumber;
        }
        for (j = 0; j < SETTING_FLAG_KEYS.length; j++) {
            savedPreset[SETTING_FLAG_KEYS[j]] = (storedValues[SETTING_FLAG_KEYS[j]] === "1");
        }
        return savedPreset;
    }

    /**
     * ダイアログを開いたときに選ぶプリセットの位置を返す
     * 保存した設定があればそれを優先し、なければ DEFAULT_PRESET_KEY のプリセットを選ぶ
     * @param {Array} presetList - ドロップダウンに並べるプリセット
     * @returns {number} 選択する位置
     */
    function findDefaultPresetIndex(presetList) {
        var i;
        for (i = 0; i < presetList.length; i++) {
            if (presetList[i].labelKey === "preset.saved") return i;
        }
        for (i = 0; i < presetList.length; i++) {
            if (presetList[i].labelKey === DEFAULT_PRESET_KEY) return i;
        }
        return 0;
    }

    /**
     * 値を下限〜上限の範囲に収める
     * @param {number} targetValue - 丸めたい値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @param {number} fallbackValue - 数値として読めないときの値（省略時は下限）
     * @returns {number} 範囲内に収めた値
     */
    function clampRange(targetValue, minValue, maxValue, fallbackValue) {
        if (isNaN(targetValue)) {
            return (typeof fallbackValue === "number") ? fallbackValue : minValue;
        }
        if (targetValue < minValue) return minValue;
        if (targetValue > maxValue) return maxValue;
        return targetValue;
    }

    /**
     * ぼかし（ガウス）をライブエフェクトとして適用する
     * @param {object} targetItem - 対象のページアイテム
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
     * ドロップシャドウをライブエフェクトとして適用する
     * @param {object} targetItem - 対象のページアイテム
     * @param {object} settings - shadowOpacity、shadowOffsetX、shadowOffsetY、shadowBlur、shadowDarkness を持つ設定
     * @returns {void}
     */
    function applyDropShadow(targetItem, settings) {
        /* 「濃さ」指定でも色の項目は残すため、ドキュメントのカラーモードに合わせた黒を入れる
           The color entry stays even in darkness mode, so fill it with black for the document color mode */
        var shadowColor = (doc.documentColorSpace === DocumentColorSpace.CMYK) ? "0 0 0 1" : "0 0 0";
        var effectXml = '<LiveEffect name="Adobe Drop Shadow"><Dict data="' +
            'I blnd ' + SHADOW_BLEND_MODE +
            ' R opac ' + (settings.shadowOpacity / 100) +
            ' R horz ' + settings.shadowOffsetX +
            ' R vert ' + settings.shadowOffsetY +
            ' R blur ' + settings.shadowBlur +
            ' B usePSLBlur 1 I csrc 1 R dark ' + settings.shadowDarkness + ' B pair 1 ">' +
            '<Entry name="sclr" valueType="F"><Fill color="' + shadowColor + '"/></Entry>' +
            '</Dict></LiveEffect>';
        targetItem.applyEffect(effectXml);
    }

    /**
     * パスをマスクにしてページアイテムをクリッピングする
     * @param {object} maskedItem - マスクされるアイテム
     * @param {object} clipPath - マスクに使うパス
     * @returns {object} 作成したクリップグループ
     */
    function createClipGroup(maskedItem, clipPath) {
        var clipGroup = doc.groupItems.add();
        /* マスク用パスがあった重ね順にグループを置く / keep the mask path's stacking position */
        clipGroup.move(clipPath, ElementPlacement.PLACEBEFORE);
        clipPath.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        maskedItem.move(clipGroup, ElementPlacement.PLACEATEND);
        clipGroup.clipped = true;
        return clipGroup;
    }

    /**
     * クリップグループに線を追加する
     * 線には「中マド」が乗り、グループ全体の「分割」は元がグループのときだけ適用する
     * @param {object} clipGroup - 対象のクリップグループ
     * @returns {void}
     */
    function addStrokeToClipGroup(clipGroup) {
        try {
            doc.selection = [clipGroup];
            /* 選択を反映させてからコマンドを流す / let the new selection settle before the command */
            app.redraw();
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            /* 追加した線がアピアランスで選ばれている状態なので、中マドは線に乗る
               The new stroke is the active appearance row, so Exclude lands on the stroke */
            app.executeMenuCommand('Live Pathfinder Exclude');
        } catch (e) {
            /* 環境によっては実行できないため無視 / ignore when the commands are unavailable */
        }

        /* グループ全体への「分割」は、元がベクターのグループのときだけ
           画像（配置画像・ラスター画像）には分割するパスがなく、効かないので適用しない
           Divide only makes sense for a vector group; a placed or raster image has no paths to divide */
        if (originalArtwork.typename === "GroupItem") {
            clipGroup.applyEffect(PATHFINDER_DIVIDE_EFFECT);
        }
    }

    /**
     * 変形（効果）で拡大と移動を適用する（アイテムそのものは変形しない）
     * @param {object} targetItem - 対象のアイテム
     * @param {number} scalePercent - 拡大率（%）
     * @param {number} moveHorizontal - 左右の移動量（pt、右が正）
     * @param {number} moveVertical - 上下の移動量（pt、下が正）
     * @param {number} copyCount - コピーの数（1なら元を残したまま変形したコピーが1つ増える）
     * @returns {void}
     */
    function applyTransformEffect(targetItem, scalePercent, moveHorizontal, moveVertical, copyCount) {
        /* 変形効果は上方向が正なので符号を反転し、呼び出し側は「下が正」で揃える
           The Transform effect counts up as positive, so flip the sign and let callers treat down as positive */
        var effectXml = '<LiveEffect name="Adobe Transform"><Dict data="' +
            'R scaleH_Percent ' + scalePercent +
            ' R scaleV_Percent ' + scalePercent +
            ' R scaleH_Factor ' + (scalePercent / 100) +
            ' R scaleV_Factor ' + (scalePercent / 100) +
            ' R moveH_Pts ' + moveHorizontal +
            ' R moveV_Pts ' + (-moveVertical) +
            ' R rotate_Degrees 0 R rotate_Radians 0 I numCopies ' + copyCount +
            ' I pinPoint ' + TRANSFORM_PIN_CENTER +
            ' B scaleLines 1 B transformPatterns 1 B transformObjects 1' +
            ' B reflectX 0 B reflectY 0 B randomize 0 "/></LiveEffect>';
        targetItem.applyEffect(effectXml);
    }

    /**
     * クリップグループを拡大・移動してルーペにする
     * 「変形効果で拡大」がオンなら［変形］効果だけで済ませ、「元を残して複製」は効果のコピーで作る
     * オフのときは実際に複製してから拡大・移動する（「複製」がオフなら複製もしない）
     * @param {object} clipGroup - もとになるクリップグループ
     * @param {object} settings - duplicatesLoupe、scalesLoupeByEffect、loupeScale、loupeOffsetX、loupeOffsetY を持つ設定
     * @returns {object} ルーペになったグループ（複製、またはクリップグループ自身）
     */
    function applyLoupe(clipGroup, settings) {
        /* 移動量の基準はマスクのパスそのもの。クリップグループの境界だと
           中身（マスクの外まで広がる画像）や効果に引きずられて、縦横で違う値になる
           Measure from the mask path itself: the clip group's bounds follow its contents,
           which reach past the mask, and give a lopsided result */
        var maskBounds = originalMaskPath.geometricBounds; /* [left, top, right, bottom] */
        var maskWidth = maskBounds[2] - maskBounds[0];
        var maskHeight = maskBounds[1] - maskBounds[3];

        var scalePercent = settings.loupeScale;

        /* 移動量はマスクの大きさに対する割合 / the offsets are a share of the mask size */
        var moveHorizontal = maskWidth * settings.loupeOffsetX / 100;
        var moveVertical = maskHeight * settings.loupeOffsetY / 100;

        /* 変形効果で拡大するときは、複製も［変形］効果のコピー（numCopies）で作る
           オブジェクトは増えず、あとからコピーの数も位置も直せる
           With the effect, the copy comes from the Transform effect itself: no extra object
           is created and both the count and the placement stay editable */
        if (settings.scalesLoupeByEffect) {
            applyTransformEffect(clipGroup, scalePercent, moveHorizontal, moveVertical,
                settings.duplicatesLoupe ? 1 : 0);
            return clipGroup;
        }

        var loupeGroup = settings.duplicatesLoupe ?
            clipGroup.duplicate(clipGroup, ElementPlacement.PLACEBEFORE) :
            clipGroup;

        /* マスクも中身もまとめて、中心を動かさずに拡大 / scale mask and contents together about the center */
        loupeGroup.resize(scalePercent, scalePercent, true, true, true, true, scalePercent, Transformation.CENTER);
        /* DOMのY軸は上が正なので符号を反転する / the DOM's Y axis points up, so flip the sign */
        loupeGroup.translate(moveHorizontal, -moveVertical);
        return loupeGroup;
    }

    /**
     * プレビューの間だけアイテムを隠し、あとで戻せるように控える
     * @param {object} targetItem - 隠すページアイテム
     * @returns {void}
     */
    function hideForPreview(targetItem) {
        targetItem.hidden = true;
        previewHiddenItems.push(targetItem);
    }

    /**
     * マスク外をぼかす。確定実行は元のアイテムをそのままぼかし、プレビューはぼかした複製で代用する
     * @param {number} blurRadius - ぼかしの半径（px）
     * @param {boolean} isCommitRun - true で確定実行、false でプレビュー
     * @returns {void}
     */
    function blurOutsideMask(blurRadius, isCommitRun) {
        if (isCommitRun) {
            applyGaussianBlur(originalArtwork, blurRadius);
            return;
        }
        var blurredBackdropCopy = originalArtwork.duplicate(originalArtwork, ElementPlacement.PLACEBEFORE);
        applyGaussianBlur(blurredBackdropCopy, blurRadius);
        previewOnlyItems.push(blurredBackdropCopy);
        hideForPreview(originalArtwork);
    }

    /**
     * ドキュメントのカラーモードに合わせた黒を返す
     * @returns {object} CMYKColor または RGBColor
     */
    function createBlackColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykBlack = new CMYKColor();
            cmykBlack.cyan = 0;
            cmykBlack.magenta = 0;
            cmykBlack.yellow = 0;
            cmykBlack.black = 100;
            return cmykBlack;
        }
        var rgbBlack = new RGBColor();
        rgbBlack.red = 0;
        rgbBlack.green = 0;
        rgbBlack.blue = 0;
        return rgbBlack;
    }

    /**
     * 黒い長方形を元のアートワークのすぐ前面に重ねて、マスク外を暗くする
     * マスクした複製（B）はさらに前面に来るので、マスクの内側は暗くならず、穴は開けなくてよい
     * @param {number} darknessPercent - 暗さ（重ねる黒の不透明度、%）
     * @returns {object} 重ねた長方形
     */
    function dimOutsideMask(darknessPercent) {
        var artworkBounds = originalArtwork.visibleBounds; /* [left, top, right, bottom] */
        var dimOverlay = originalArtwork.parent.pathItems.rectangle(
            artworkBounds[1], artworkBounds[0],
            artworkBounds[2] - artworkBounds[0], artworkBounds[1] - artworkBounds[3]);

        /* 元のアートワークのすぐ前面へ。マスク用パスはこれより前面にあるので、
           クリップグループ（マスクした複製）は暗くならない
           Just in front of the artwork; the mask path sits further forward, so the clip group stays bright */
        dimOverlay.move(originalArtwork, ElementPlacement.PLACEBEFORE);
        dimOverlay.filled = true;
        dimOverlay.stroked = false;
        dimOverlay.fillColor = createBlackColor();
        dimOverlay.opacity = darknessPercent;
        return dimOverlay;
    }

    /**
     * マスクに使うパスを返す。確定実行は元のパス、プレビューはその複製
     * @param {boolean} isCommitRun - true で確定実行、false でプレビュー
     * @returns {object} マスクに使うパス
     */
    function prepareClipPath(isCommitRun) {
        if (isCommitRun) return originalMaskPath;

        var clipPathCopy = originalMaskPath.duplicate(originalMaskPath, ElementPlacement.PLACEBEFORE);
        hideForPreview(originalMaskPath);
        return clipPathCopy;
    }

    /**
     * マスクした複製を作り、ぼかし・スポットライト・線・ドロップシャドウ・ルーペを適用する
     * 確定実行では元のアイテムとパスをそのまま使うため、増えるのは複製1つとクリップグループだけ
     * @param {object} settings - effectMode、blurRadius、spotlightDarkness、addsStroke、addsDropShadow、ルーペ関連を持つ設定
     * @param {boolean} isCommitRun - true で確定実行、false でプレビュー
     * @returns {object} 作成したクリップグループ
     */
    function buildSpotlight(settings, isCommitRun) {
        /* マスクされる複製（B）。元のアイテムを隠す前に作る（隠した状態が複製に写るため）
           The copy to be masked (B), made before the original is hidden so it does not inherit that state */
        var maskedArtworkCopy = originalArtwork.duplicate(originalArtwork, ElementPlacement.PLACEBEFORE);

        /* 半径0ではぼかし効果を作れないので、ぼかさない扱いにする
           A radius of zero cannot make a blur, so treat it as no blur at all */
        if (settings.blurRadius > 0) {
            if (settings.effectMode === EFFECT_BLUR_INSIDE) {
                applyGaussianBlur(maskedArtworkCopy, settings.blurRadius);
            } else if (settings.effectMode === EFFECT_BLUR_OUTSIDE) {
                blurOutsideMask(settings.blurRadius, isCommitRun);
            }
        }

        /* スポットライトはぼかさず、マスク外を黒で暗くする
           暗さ0では見た目が変わらないので、そのときは重ねない
           Spotlight darkens the outside instead of blurring it; a darkness of zero lays nothing over it */
        if (settings.effectMode === EFFECT_SPOTLIGHT && settings.spotlightDarkness > 0) {
            var dimOverlay = dimOutsideMask(settings.spotlightDarkness);
            if (!isCommitRun) previewOnlyItems.push(dimOverlay);
        }

        var clipGroup = createClipGroup(maskedArtworkCopy, prepareClipPath(isCommitRun));
        if (!isCommitRun) {
            previewOnlyItems.push(clipGroup);
        }

        if (settings.addsStroke) {
            addStrokeToClipGroup(clipGroup);
        }
        if (settings.addsDropShadow) {
            applyDropShadow(clipGroup, settings);
        }

        /* マスク内をぼかすと中身がぼけてしまうので、そのときだけルーペは使わない
           Blurring the inside leaves nothing sharp to magnify, so the loupe is skipped there */
        var loupeGroup = null;
        if (settings.addsLoupe && settings.effectMode !== EFFECT_BLUR_INSIDE) {
            loupeGroup = applyLoupe(clipGroup, settings);
        }

        /* 複製していないときのルーペはクリップグループ自身なので、控えも選択も重ねない
           Without a copy the loupe is the clip group itself, so do not track or select it twice */
        var hasSeparateLoupe = (loupeGroup !== null && loupeGroup !== clipGroup);
        if (!isCommitRun && hasSeparateLoupe) {
            previewOnlyItems.push(loupeGroup);
        }

        /* 「拡大表示のコネクター」。グラフィックスタイルは適用先のアピアランスを置き換えるので、
           線・ドロップシャドウ・［変形］効果を残せるよう、グループでくるんでから適用する
           A graphic style replaces the appearance of whatever it lands on, so the result is wrapped
           in a group and the style goes on that group, leaving the stroke, shadow and Transform effect intact */
        if (settings.addsConnector && canAddConnector(settings)) {
            var connectorGroup = groupItemsTogether(hasSeparateLoupe ? [loupeGroup, clipGroup] : [clipGroup]);
            applyConnectorStyle(connectorGroup);
            /* 中身より後に控える。プレビューの後片付けで、空になったグループが最後に消える
               Tracked after its contents, so the cleanup empties the group before removing it */
            if (!isCommitRun) previewOnlyItems.push(connectorGroup);

            doc.selection = [connectorGroup];
            return connectorGroup;
        }

        doc.selection = hasSeparateLoupe ? [clipGroup, loupeGroup] : [clipGroup];
        return clipGroup;
    }

    /**
     * プレビューで作ったアイテムを消し、隠したアイテムを元に戻す
     * @returns {void}
     */
    function discardPreviewItems() {
        /* 削除より先に元へ戻す。削除でつまずいても元のアイテムを隠したままにしないため
           Restore first: a failed remove must not leave the user's artwork hidden */
        for (var i = 0; i < previewHiddenItems.length; i++) {
            previewHiddenItems[i].hidden = false;
        }
        previewHiddenItems = [];

        var leftoverItems = previewOnlyItems;
        previewOnlyItems = [];
        for (var j = 0; j < leftoverItems.length; j++) {
            leftoverItems[j].remove();
        }
    }

    // =========================================
    // コネクターのグラフィックスタイル / The connector's graphic style
    // =========================================

    /* ライブラリーからの取り込みを試したか（実行中だけ覚える）
       Whether the import has been attempted, remembered for this run */
    var hasTriedConnectorImport = false;

    /**
     * コネクターのグラフィックスタイルのライブラリーを返す（このスクリプトと同じ場所）
     * @returns {object} ライブラリーのファイル。無ければ null
     */
    function getConnectorLibraryFile() {
        /* スクリプトの場所が取れない実行方法もある / Some ways of running a script leave no path to work from */
        var scriptFolder = File($.fileName).parent;
        if (!scriptFolder) return null;

        var libraryFile = new File(scriptFolder.fsName + "/" + CONNECTOR_LIBRARY_NAME);
        return libraryFile.exists ? libraryFile : null;
    }

    /**
     * ドキュメントの［グラフィックスタイル］パネルに、その名前のスタイルがあるか
     * @param {string} styleName - 探すスタイル名
     * @returns {boolean} あれば true
     */
    function hasGraphicStyle(styleName) {
        for (var i = 0; i < doc.graphicStyles.length; i++) {
            if (doc.graphicStyles[i].name === styleName) return true;
        }
        return false;
    }

    /**
     * ライブラリーのグラフィックスタイルをドキュメントに取り込む
     * ライブラリーのオブジェクトを一時レイヤーへ貼り付けると、そのレイヤーを捨ててもスタイルはパネルに残る
     * @returns {void}
     */
    function importConnectorStyles() {
        var libraryFile = getConnectorLibraryFile();
        if (!libraryFile) return;

        /* すでに開いているライブラリーは開き直しても手前に来るだけでドキュメント数が変わらない。
           これを目印にすれば、パス文字列の比較（日本語パスで外れる）に頼らず判定できる
           Reopening an already open library only brings it to front, leaving the document count unchanged;
           that is a safer signal than comparing path strings */
        var documentCountBefore = app.documents.length;
        var libraryDoc = app.open(libraryFile);
        var wasLibraryOpen = (app.documents.length === documentCountBefore);

        try {
            /* コピーはメニューコマンドで行う（app.copy() は黙って失敗することがある）
               A menu command does the copy; app.copy() can fail silently */
            app.redraw();
            app.executeMenuCommand("selectallinartboard");
            app.executeMenuCommand("copy");
        } finally {
            /* 開いたのがこのスクリプトなら閉じる。もともと開いていたものは触らない / Close only what this script opened */
            if (!wasLibraryOpen) libraryDoc.close(SaveOptions.DONOTSAVECHANGES);
            app.activeDocument = doc;
        }

        /* 貼り付け前に選択を解除する（解除しないと選択中のオブジェクトが置き換わる）
           Clear the selection first, or the paste replaces what is selected */
        doc.selection = null;

        var importLayer = doc.layers.add();
        importLayer.name = CONNECTOR_IMPORT_LAYER;
        doc.activeLayer = importLayer;
        app.executeMenuCommand("paste");

        /* スタイルはパネルに残るので、貼り付けたものはレイヤーごと捨てる
           The styles stay in the panel, so what was pasted goes with the layer */
        try {
            importLayer.remove();
        } catch (eLayer) {
            /* レイヤーを消せなくても、取り込み自体は済んでいる / The import is done even when the layer cannot be removed */
        }
    }

    /**
     * コネクターのグラフィックスタイルを使える状態にする
     * 取り込みは一度で全スタイルが入るので、ライブラリーは開き直さない
     * @returns {boolean} 使える状態になったら true
     */
    function ensureConnectorStyle() {
        if (hasGraphicStyle(CONNECTOR_STYLE_NAME)) return true;
        if (hasTriedConnectorImport) return false;

        hasTriedConnectorImport = true;
        importConnectorStyles();
        return hasGraphicStyle(CONNECTOR_STYLE_NAME);
    }

    /**
     * 「拡大表示のコネクター」を使える設定か
     * つなぐ相手ができるのは、ルーペを複製するときだけ
     * @param {object} settings - effectMode、addsLoupe、duplicatesLoupe を持つ設定
     * @returns {boolean} 使えるなら true
     */
    function canAddConnector(settings) {
        return settings.addsLoupe && settings.duplicatesLoupe && settings.effectMode !== EFFECT_BLUR_INSIDE;
    }

    /**
     * 渡したアイテムを1つのグループにまとめる
     * @param {Array} targetItems - まとめるアイテム（前面のものから順に）
     * @returns {object} 作ったグループ
     */
    function groupItemsTogether(targetItems) {
        var newGroup = targetItems[0].parent.groupItems.add();
        newGroup.move(targetItems[0], ElementPlacement.PLACEBEFORE);
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].move(newGroup, ElementPlacement.PLACEATEND);
        }
        return newGroup;
    }

    /**
     * コネクターのグラフィックスタイルを適用する
     * @param {object} targetGroup - 適用先のグループ
     * @returns {boolean} 適用できたら true
     */
    function applyConnectorStyle(targetGroup) {
        if (!ensureConnectorStyle()) return false;
        try {
            doc.graphicStyles.getByName(CONNECTOR_STYLE_NAME).applyTo(targetGroup);
            return true;
        } catch (eStyle) {
            /* 適用できなくても、マスクまでの結果は残す / A failure here must not discard the result built so far */
            return false;
        }
    }

    // =========================================
    // UI部品 / UI parts
    // =========================================

    /**
     * ∧∨と数値欄を隙間なく並べて追加する。範囲と増減後の処理は、あとから .stepOptions（min / max / onStep）に入れる
     * @param {object} parentGroup - 追加先
     * @param {number} defaultValue - 初期値
     * @returns {object} 追加した数値欄（∧∨は .stepperGroup、増減の設定は .stepOptions で参照できる）
     */
    function addStepperNumberField(parentGroup, defaultValue) {
        var stepperInputGroup = parentGroup.add('group');
        stepperInputGroup.orientation = 'row';
        stepperInputGroup.alignChildren = ['left', 'center'];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var stepOptions = {};
        var numberField;
        var stepperGroup = addStepper(stepperInputGroup, function() { return numberField; }, stepOptions);
        numberField = stepperInputGroup.add('edittext', undefined, String(defaultValue));
        numberField.characters = 4;
        numberField.stepperGroup = stepperGroup;
        numberField.stepOptions = stepOptions;
        return numberField;
    }

    /**
     * ラベル＋数値欄＋単位の行を作る
     * @param {object} parentPanel - 行を追加する親
     * @param {string} labelPath - 項目名のラベルキー
     * @param {string} tooltipPath - tooltipのラベルキー
     * @param {number} defaultValue - 初期値
     * @param {string} unitText - 単位の表記
     * @returns {object} 追加した数値欄
     */
    function createNumberFieldRow(parentPanel, labelPath, tooltipPath, defaultValue, unitText) {
        var fieldRow = parentPanel.add('group');
        fieldRow.orientation = 'row';
        fieldRow.alignment = ['fill', 'top'];
        fieldRow.alignChildren = ['left', 'center'];

        var rowLabel = fieldRow.add('statictext', undefined, labelText(labelPath));
        rowLabel.preferredSize.width = NUMBER_LABEL_WIDTH;
        rowLabel.justify = 'right';

        var numberField = addStepperNumberField(fieldRow, defaultValue);
        numberField.helpTip = getLabel(tooltipPath);

        fieldRow.add('statictext', undefined, unitText);
        return numberField;
    }

    /**
     * 「半径（px）：」を1行目、数値欄を2行目に置く（幅の狭い列で使う）
     * @param {object} parentPanel - 行を追加する親
     * @param {string} labelPath - 項目名のラベルキー
     * @param {string} tooltipPath - tooltipのラベルキー
     * @param {number} defaultValue - 初期値
     * @param {string} unitText - 項目名に添える単位
     * @returns {object} 入れ物（group）と数値欄（field）を持つオブジェクト
     */
    function createStackedNumberField(parentPanel, labelPath, tooltipPath, defaultValue, unitText) {
        var fieldColumn = parentPanel.add('group');
        fieldColumn.orientation = 'column';
        fieldColumn.alignment = ['left', 'top'];
        fieldColumn.alignChildren = ['left', 'top'];
        fieldColumn.spacing = 4;

        fieldColumn.add('statictext', undefined, labelText(labelPath, unitText));

        var numberField = addStepperNumberField(fieldColumn, defaultValue);
        numberField.helpTip = getLabel(tooltipPath);

        return { group: fieldColumn, field: numberField };
    }

    /**
     * ラベル＋スライダーの行を作る
     * @param {object} parentPanel - 行を追加する親
     * @param {string} labelPath - 項目名のラベルキー
     * @param {string} tooltipPath - tooltipのラベルキー
     * @returns {object} 追加したスライダー
     */
    function createOffsetSliderRow(parentPanel, labelPath, tooltipPath) {
        var sliderRow = parentPanel.add('group');
        sliderRow.orientation = 'row';
        sliderRow.alignChildren = ['left', 'center'];
        sliderRow.alignment = ['fill', 'top'];

        var rowLabel = sliderRow.add('statictext', undefined, labelText(labelPath));
        rowLabel.preferredSize.width = SLIDER_LABEL_WIDTH;
        rowLabel.justify = 'left';

        var offsetSlider = sliderRow.add('slider', undefined, 0, -MAX_LOUPE_OFFSET, MAX_LOUPE_OFFSET);
        offsetSlider.helpTip = getLabel(tooltipPath);
        offsetSlider.preferredSize.width = SLIDER_WIDTH;
        offsetSlider.alignment = ['fill', 'center'];

        return offsetSlider;
    }

    /**
     * X/Yを0に戻すボタンを作る
     * @param {object} parentGroup - ボタンを追加する親
     * @returns {object} 追加したボタン
     */
    function createResetPositionButton(parentGroup) {
        var resetPositionButton = parentGroup.add('button', undefined, getLabel('button.resetPosition'));
        resetPositionButton.helpTip = getLabel('tooltip.resetPosition');
        resetPositionButton.alignment = ['right', 'center'];
        return resetPositionButton;
    }

    /**
     * プリセットの選択行を組み立てる
     * @param {object} parentWindow - 行を追加する親
     * @param {Array} presetList - ドロップダウンに並べるプリセット
     * @returns {object} dropdown と saveButton を持つオブジェクト
     */
    function createPresetRow(parentWindow, presetList) {
        var presetRow = parentWindow.add('group');
        presetRow.orientation = 'row';
        presetRow.alignment = ['fill', 'top'];
        presetRow.alignChildren = ['left', 'center'];

        /* 「プリセット：」は数値欄のラベル幅（NUMBER_LABEL_WIDTH）に収まらないので幅を固定しない
           This label does not fit the number-row label width, so let it size itself */
        var presetLabel = presetRow.add('statictext', undefined, labelText('label.preset'));
        presetLabel.helpTip = getLabel('tooltip.preset');

        var presetDropdown = presetRow.add('dropdownlist', undefined, []);
        for (var i = 0; i < presetList.length; i++) {
            presetDropdown.add('item', getLabel(presetList[i].labelKey));
        }
        /* 末尾は「カスタム」。手で値を変えたときに選ばれる / Custom sits last and is picked on any manual edit */
        presetDropdown.add('item', getLabel('preset.custom'));
        presetDropdown.helpTip = getLabel('tooltip.preset');
        presetDropdown.alignment = ['fill', 'center'];
        presetDropdown.selection = findDefaultPresetIndex(presetList);

        var saveButton = presetRow.add('button', undefined, getLabel('button.save'));
        saveButton.helpTip = getLabel('tooltip.save');
        /* 行が伸びてもボタンは広げない / the row may stretch, the button does not */
        saveButton.alignment = ['right', 'center'];
        saveButton.preferredSize.width = SAVE_BUTTON_WIDTH;

        return { dropdown: presetDropdown, saveButton: saveButton };
    }

    /**
     * 「ドロップシャドウ」パネルを組み立てる
     * @param {object} parentWindow - パネルを追加する親
     * @returns {object} チェックボックスと各数値欄を持つオブジェクト
     */
    function createDropShadowPanel(parentWindow) {
        var dropShadowPanel = parentWindow.add('panel', undefined, getLabel('panel.dropShadow'));
        dropShadowPanel.orientation = 'column';
        dropShadowPanel.alignChildren = ['fill', 'top'];
        dropShadowPanel.margins = [15, 20, 15, 10];

        var addDropShadowCheckbox = dropShadowPanel.add('checkbox', undefined, getLabel('checkbox.dropShadow'));
        addDropShadowCheckbox.helpTip = getLabel('tooltip.dropShadow');
        addDropShadowCheckbox.alignment = ['left', 'center'];
        addDropShadowCheckbox.value = false;

        /* チェックと連動してまとめてディム表示にするための入れ物
           A container so the settings dim together with the checkbox */
        var shadowSettingsGroup = dropShadowPanel.add('group');
        shadowSettingsGroup.orientation = 'column';
        shadowSettingsGroup.alignChildren = ['fill', 'top'];
        shadowSettingsGroup.spacing = 6;

        return {
            addCheckbox: addDropShadowCheckbox,
            settingsGroup: shadowSettingsGroup,
            opacityField:  createNumberFieldRow(shadowSettingsGroup, 'label.shadowOpacity',  'tooltip.shadowOpacity',  DEFAULT_SHADOW_OPACITY,  '%'),
            offsetXField:  createNumberFieldRow(shadowSettingsGroup, 'label.offsetX',        'tooltip.shadowOffsetX',  DEFAULT_SHADOW_OFFSET_X, 'pt'),
            offsetYField:  createNumberFieldRow(shadowSettingsGroup, 'label.offsetY',        'tooltip.shadowOffsetY',  DEFAULT_SHADOW_OFFSET_Y, 'pt'),
            blurField:     createNumberFieldRow(shadowSettingsGroup, 'label.shadowBlur',     'tooltip.shadowBlur',     DEFAULT_SHADOW_BLUR,     'pt'),
            darknessField: createNumberFieldRow(shadowSettingsGroup, 'label.shadowDarkness', 'tooltip.shadowDarkness', DEFAULT_SHADOW_DARKNESS, '%')
        };
    }

    /**
     * 「ルーペ」パネルを組み立てる
     * @param {object} parentWindow - パネルを追加する親
     * @returns {object} パネル本体と各コントロールを持つオブジェクト
     */
    function createLoupePanel(parentWindow) {
        var loupePanel = parentWindow.add('panel', undefined, getLabel('panel.loupe'));
        loupePanel.orientation = 'column';
        loupePanel.alignChildren = ['fill', 'top'];
        loupePanel.margins = [15, 20, 15, 10];

        var addLoupeCheckbox = loupePanel.add('checkbox', undefined, getLabel('checkbox.addLoupe'));
        addLoupeCheckbox.helpTip = getLabel('tooltip.addLoupe');
        addLoupeCheckbox.alignment = ['left', 'center'];
        addLoupeCheckbox.value = false;

        /* チェックと連動してまとめてディム表示にするための入れ物
           A container so the settings dim together with the checkbox */
        var loupeSettingsGroup = loupePanel.add('group');
        loupeSettingsGroup.orientation = 'column';
        loupeSettingsGroup.alignChildren = ['fill', 'top'];
        loupeSettingsGroup.spacing = 6;

        var duplicateLoupeCheckbox = loupeSettingsGroup.add('checkbox', undefined, getLabel('checkbox.duplicateLoupe'));
        duplicateLoupeCheckbox.helpTip = getLabel('tooltip.duplicateLoupe');
        duplicateLoupeCheckbox.alignment = ['left', 'center'];
        duplicateLoupeCheckbox.value = false;

        var scaleByEffectCheckbox = loupeSettingsGroup.add('checkbox', undefined, getLabel('checkbox.scaleByEffect'));
        scaleByEffectCheckbox.helpTip = getLabel('tooltip.scaleByEffect');
        scaleByEffectCheckbox.alignment = ['left', 'center'];
        scaleByEffectCheckbox.value = false;

        return {
            panel: loupePanel,
            addCheckbox: addLoupeCheckbox,
            duplicateCheckbox: duplicateLoupeCheckbox,
            scaleByEffectCheckbox: scaleByEffectCheckbox,
            settingsGroup: loupeSettingsGroup,
            scaleField: createNumberFieldRow(loupeSettingsGroup, 'label.loupeScale', 'tooltip.loupeScale', DEFAULT_LOUPE_SCALE, '%'),
            offsetX: createOffsetSliderRow(loupeSettingsGroup, 'label.offsetX', 'tooltip.loupeOffsetX'),
            offsetY: createOffsetSliderRow(loupeSettingsGroup, 'label.offsetY', 'tooltip.loupeOffsetY'),
            resetPositionButton: createResetPositionButton(loupeSettingsGroup)
        };
    }

    /**
     * 「効果」パネルを組み立てる
     * @param {object} parentWindow - パネルを追加する親
     * @returns {object} 各ラジオと、半径・暗さの欄とその入れ物を持つオブジェクト
     */
    function createEffectPanel(parentWindow) {
        var effectPanel = parentWindow.add('panel', undefined, getLabel('panel.effect'));
        effectPanel.orientation = 'column';
        effectPanel.alignChildren = ['left', 'top'];
        effectPanel.margins = [15, 20, 15, 10];

        /* 左にラジオ、右に数値欄の2カラム / Two columns: the radios on the left, the number fields on the right */
        var effectBodyRow = effectPanel.add('group');
        effectBodyRow.orientation = 'row';
        effectBodyRow.alignment = ['fill', 'top'];
        effectBodyRow.alignChildren = ['left', 'top'];
        effectBodyRow.spacing = 15;

        /* ラジオは同じ親に入れないと排他にならない / radios must share a parent to stay exclusive */
        var effectRadioColumn = effectBodyRow.add('group');
        effectRadioColumn.orientation = 'column';
        effectRadioColumn.alignChildren = ['left', 'center'];
        effectRadioColumn.spacing = 6;

        var noneRadio = effectRadioColumn.add('radiobutton', undefined, getLabel('radio.effectNone'));
        noneRadio.helpTip = getLabel('tooltip.effectNone');

        var blurInsideRadio = effectRadioColumn.add('radiobutton', undefined, getLabel('radio.blurInside'));
        blurInsideRadio.helpTip = getLabel('tooltip.blurInside');
        blurInsideRadio.value = true;

        var blurOutsideRadio = effectRadioColumn.add('radiobutton', undefined, getLabel('radio.blurOutside'));
        blurOutsideRadio.helpTip = getLabel('tooltip.blurOutside');

        var spotlightRadio = effectRadioColumn.add('radiobutton', undefined, getLabel('radio.spotlight'));
        spotlightRadio.helpTip = getLabel('tooltip.spotlight');

        /* 選んだ効果で使う欄だけを有効にする。入れ物ごと切り替えて、項目名も一緒にディム表示にする
           Only the field the chosen effect uses stays enabled; switching the container dims its label too */
        var effectFieldColumn = effectBodyRow.add('group');
        effectFieldColumn.orientation = 'column';
        effectFieldColumn.alignChildren = ['left', 'top'];
        effectFieldColumn.spacing = 8;

        var blurRadius = createStackedNumberField(effectFieldColumn, 'label.blurRadius', 'tooltip.blurRadius', DEFAULT_BLUR_RADIUS, 'px');
        var spotlightDarkness = createStackedNumberField(effectFieldColumn, 'label.spotlightDarkness', 'tooltip.spotlightDarkness', DEFAULT_SPOTLIGHT_DARKNESS, '%');

        return {
            noneRadio: noneRadio,
            blurInsideRadio: blurInsideRadio,
            blurOutsideRadio: blurOutsideRadio,
            spotlightRadio: spotlightRadio,
            radiusGroup: blurRadius.group,
            radiusField: blurRadius.field,
            darknessGroup: spotlightDarkness.group,
            darknessField: spotlightDarkness.field
        };
    }

    /**
     * 「オプション」パネルを組み立てる
     * @param {object} parentWindow - パネルを追加する親
     * @returns {object} addStrokeCheckbox と addConnectorCheckbox を持つオブジェクト
     */
    function createOptionsPanel(parentWindow) {
        var optionsPanel = parentWindow.add('panel', undefined, getLabel('panel.option'));
        optionsPanel.orientation = 'column';
        optionsPanel.alignChildren = ['left', 'top'];
        optionsPanel.margins = [15, 20, 15, 10];

        var addStrokeCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.addStroke'));
        addStrokeCheckbox.helpTip = getLabel('tooltip.addStroke');
        addStrokeCheckbox.value = false;

        /* コネクターはライブラリーのグラフィックスタイルで描くので、ライブラリーが無ければ項目ごと出さない
           The connector is drawn by a graphic style from the library, so without the library the option is left out */
        var addConnectorCheckbox = null;
        if (getConnectorLibraryFile()) {
            addConnectorCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.addConnector'));
            addConnectorCheckbox.helpTip = getLabel('tooltip.addConnector');
            addConnectorCheckbox.value = false;
        }

        return { addStrokeCheckbox: addStrokeCheckbox, addConnectorCheckbox: addConnectorCheckbox };
    }

    /**
     * プレビューチェックとOK／キャンセルボタンの行を組み立てる
     * @param {object} parentWindow - 行を追加する親
     * @returns {object} previewCheckbox、btnCancel、btnOk を持つオブジェクト
     */
    function createButtonRow(parentWindow) {
        var btnRowGroup = parentWindow.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignment = ['fill', 'top'];
        btnRowGroup.alignChildren = ['left', 'center'];

        var btnLeftGroup = btnRowGroup.add('group');
        btnLeftGroup.orientation = 'row';
        btnLeftGroup.alignChildren = ['left', 'center'];

        var previewCheckbox = btnLeftGroup.add('checkbox', undefined, getLabel('checkbox.preview'));
        previewCheckbox.helpTip = getLabel('tooltip.preview');
        previewCheckbox.value = true;

        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.orientation = 'row';
        btnRightGroup.alignChildren = ['right', 'center'];
        btnRightGroup.alignment = ['right', 'center'];

        return {
            previewCheckbox: previewCheckbox,
            btnCancel: btnRightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' }),
            btnOk: btnRightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' })
        };
    }

    // =========================================
    // ダイアログボックス / Dialog
    // =========================================

    /**
     * ダイアログの中身を並べて、各パネルのコントロールを返す
     * @param {object} dialog - 追加先のダイアログ
     * @param {Array} presetList - ドロップダウンに並べるプリセット
     * @returns {object} preset、effect、option、dropShadow、loupe、buttonRow を持つオブジェクト
     */
    function createDialogContents(dialog, presetList) {
        var presetControls = createPresetRow(dialog, presetList);

        /* 上段：効果とオプション / Top row: the effect and the options */
        var topPanelRow = dialog.add('group');
        topPanelRow.orientation = 'row';
        topPanelRow.alignment = ['fill', 'top'];
        topPanelRow.alignChildren = ['fill', 'fill'];

        var effectControls = createEffectPanel(topPanelRow);
        var optionControls = createOptionsPanel(topPanelRow);

        /* 下段：ルーペとドロップシャドウ / Bottom row: loupe and drop shadow */
        var effectPanelRow = dialog.add('group');
        effectPanelRow.orientation = 'row';
        effectPanelRow.alignment = ['fill', 'top'];
        /* 背の低いルーペパネルを引き伸ばさず、上端で揃える
           Keep the shorter loupe panel at its own height, aligned to the top */
        effectPanelRow.alignChildren = ['fill', 'top'];

        return {
            preset: presetControls,
            effect: effectControls,
            option: optionControls,
            loupe: createLoupePanel(effectPanelRow),
            dropShadow: createDropShadowPanel(effectPanelRow),
            buttonRow: createButtonRow(dialog)
        };
    }

    /**
     * 設定ダイアログを表示し、OKなら効果を確定する
     * @returns {void}
     */
    function showSettingsDialog() {
        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = 'fill';

        /* 保存した設定があれば先頭に置く。PRESETS は書き換えないよう複製してから使う
           The saved settings go first; copy PRESETS so the original array is never touched */
        var savedPreset = loadSavedPreset();
        var presetList = savedPreset ? [savedPreset].concat(PRESETS) : PRESETS.slice(0);
        var dialogControls = createDialogContents(dialog, presetList);
        var presetControls = dialogControls.preset;
        var effectControls = dialogControls.effect;
        var optionControls = dialogControls.option;
        var dropShadowControls = dialogControls.dropShadow;
        var loupeControls = dialogControls.loupe;
        var buttonRowControls = dialogControls.buttonRow;

        var isPreviewShown = false;
        var isClosedByButton = false;

        /**
         * 選択中のラジオから効果を読み取る
         * @returns {string} EFFECT_NONE / EFFECT_BLUR_INSIDE / EFFECT_BLUR_OUTSIDE / EFFECT_SPOTLIGHT のいずれか
         */
        function getEffectMode() {
            if (effectControls.noneRadio.value) return EFFECT_NONE;
            if (effectControls.blurOutsideRadio.value) return EFFECT_BLUR_OUTSIDE;
            if (effectControls.spotlightRadio.value) return EFFECT_SPOTLIGHT;
            return EFFECT_BLUR_INSIDE;
        }

        /* 数値欄と、その許容範囲。範囲はここだけに書く
           Every number field with its range, declared in one place */
        var numberFieldSpecs = {
            blurRadius:        { field: effectControls.radiusField,       min: 0,                  max: MAX_BLUR_RADIUS },
            spotlightDarkness: { field: effectControls.darknessField,     min: 0,                  max: 100 },
            shadowOpacity:     { field: dropShadowControls.opacityField,  min: 0,                  max: 100 },
            shadowOffsetX:     { field: dropShadowControls.offsetXField,  min: -MAX_SHADOW_OFFSET, max: MAX_SHADOW_OFFSET, fallback: 0 },
            shadowOffsetY:     { field: dropShadowControls.offsetYField,  min: -MAX_SHADOW_OFFSET, max: MAX_SHADOW_OFFSET, fallback: 0 },
            shadowBlur:        { field: dropShadowControls.blurField,     min: 0,                  max: MAX_SHADOW_BLUR },
            shadowDarkness:    { field: dropShadowControls.darknessField, min: 0,                  max: 100 },
            loupeScale:        { field: loupeControls.scaleField,         min: MIN_LOUPE_SCALE,    max: MAX_LOUPE_SCALE }
        };

        /* チェックボックスも設定キーで引けるようにする。読み取り・流し込み・イベント紐づけで共用する
           The checkboxes keyed by their setting name, shared by reading, applying and wiring */
        var checkboxSpecs = {
            addsStroke:          optionControls.addStrokeCheckbox,
            addsDropShadow:      dropShadowControls.addCheckbox,
            addsLoupe:           loupeControls.addCheckbox,
            duplicatesLoupe:     loupeControls.duplicateCheckbox,
            scalesLoupeByEffect: loupeControls.scaleByEffectCheckbox
        };
        /* コネクターは、ライブラリーがあってチェックボックスを出せたときだけ設定に加わる
           The connector joins the settings only when the library let its checkbox be created */
        if (optionControls.addConnectorCheckbox) {
            checkboxSpecs.addsConnector = optionControls.addConnectorCheckbox;
        }

        /**
         * 数値欄の現在値を、決めた範囲に収めて読み取る
         * @param {string} specKey - numberFieldSpecs のキー
         * @returns {number} 範囲内に収めた値
         */
        function readNumberField(specKey) {
            var fieldSpec = numberFieldSpecs[specKey];
            return clampRange(Number(fieldSpec.field.text), fieldSpec.min, fieldSpec.max, fieldSpec.fallback);
        }

        /**
         * チェックの状態に合わせて各パネルの使用可否を切り替える
         * 半径はぼかすとき、暗さはスポットライトのとき、ルーペはマスク内をぼかすとき以外、
         * コネクターは「部分拡大」のときだけ使える
         * @returns {void}
         */
        function updateControlStates() {
            var effectMode = getEffectMode();
            effectControls.radiusGroup.enabled = (effectMode === EFFECT_BLUR_INSIDE || effectMode === EFFECT_BLUR_OUTSIDE);
            effectControls.darknessGroup.enabled = (effectMode === EFFECT_SPOTLIGHT);
            dropShadowControls.settingsGroup.enabled = dropShadowControls.addCheckbox.value;
            loupeControls.panel.enabled = (effectMode !== EFFECT_BLUR_INSIDE);
            loupeControls.settingsGroup.enabled = loupeControls.addCheckbox.value;
            /* ∧∨は自作描画なので、親の切り替えに合わせて描き直す / redraw the custom-drawn steppers */
            redrawSteppersIn(effectControls.radiusGroup);
            redrawSteppersIn(effectControls.darknessGroup);
            redrawSteppersIn(dropShadowControls.settingsGroup);
            redrawSteppersIn(loupeControls.panel);

            /* コネクターは、拡大したコピーができるとき（ルーペを複製するとき）だけ使える
               The connector needs a magnified copy to link to, so it follows the loupe's Duplicate */
            if (optionControls.addConnectorCheckbox) {
                optionControls.addConnectorCheckbox.enabled = canAddConnector({
                    effectMode: effectMode,
                    addsLoupe: loupeControls.addCheckbox.value,
                    duplicatesLoupe: loupeControls.duplicateCheckbox.value
                });
            }
        }

        /**
         * ダイアログの各コントロールから、今の設定をまとめて読み取る
         * @returns {object} プリセットと同じ形の設定
         */
        function collectSettings() {
            var settings = {
                effectMode: getEffectMode(),
                loupeOffsetX: Math.round(loupeControls.offsetX.value),
                loupeOffsetY: Math.round(loupeControls.offsetY.value)
            };
            var settingKey;
            for (settingKey in numberFieldSpecs) {
                settings[settingKey] = readNumberField(settingKey);
            }
            for (settingKey in checkboxSpecs) {
                settings[settingKey] = checkboxSpecs[settingKey].value;
            }
            return settings;
        }

        /**
         * 現在の設定で効果を組み立てる
         * @param {boolean} isCommitRun - true で確定実行、false でプレビュー
         * @returns {void}
         */
        function buildWithCurrentSettings(isCommitRun) {
            buildSpotlight(collectSettings(), isCommitRun);
        }

        /**
         * 表示中のプレビューを取り消す
         * @returns {void}
         */
        function clearPreview() {
            if (!isPreviewShown) return;
            discardPreviewItems();
            isPreviewShown = false;
        }

        /**
         * 直前のプレビューを取り消して描き直す
         * @returns {void}
         */
        function refreshPreview() {
            clearPreview();
            if (buttonRowControls.previewCheckbox.value) {
                /* 途中で失敗しても消し残さないよう、先にフラグを立てる
                   Flag it first so a failed build still gets cleaned up */
                isPreviewShown = true;
                buildWithCurrentSettings(false);
            }
            app.redraw();
        }

        /**
         * プレビューを取り消して、元の選択状態に戻す
         * @returns {void}
         */
        function restoreOriginalSelection() {
            clearPreview();
            doc.selection = [originalArtwork, originalMaskPath];
            app.redraw();
        }

        /**
         * プリセットの内容をダイアログの各コントロールへ流し込む
         * 値の代入では onClick / onChange が飛ばないので、描き直しは呼び出し側で行う
         * @param {object} preset - PRESETS の1件
         * @returns {void}
         */
        function applyPreset(preset) {
            effectControls.noneRadio.value = (preset.effectMode === EFFECT_NONE);
            effectControls.blurInsideRadio.value = (preset.effectMode === EFFECT_BLUR_INSIDE);
            effectControls.blurOutsideRadio.value = (preset.effectMode === EFFECT_BLUR_OUTSIDE);
            effectControls.spotlightRadio.value = (preset.effectMode === EFFECT_SPOTLIGHT);
            loupeControls.offsetX.value = preset.loupeOffsetX;
            loupeControls.offsetY.value = preset.loupeOffsetY;

            var settingKey;
            for (settingKey in numberFieldSpecs) {
                numberFieldSpecs[settingKey].field.text = String(preset[settingKey]);
            }
            for (settingKey in checkboxSpecs) {
                checkboxSpecs[settingKey].value = preset[settingKey];
            }
        }

        /**
         * 手で値を変えたときに、プリセットの表示を「カスタム」へ移す
         * @returns {void}
         */
        function selectCustomPreset() {
            presetControls.dropdown.selection = presetList.length;
        }

        /**
         * 設定を変えたときの共通処理（カスタム表示にして、使用可否を直してから描き直す）
         * @returns {void}
         */
        function onSettingChanged() {
            selectCustomPreset();
            updateControlStates();
            refreshPreview();
        }

        /**
         * ルーペの位置を初期値（0）に戻して描き直す
         * @returns {void}
         */
        function resetLoupePosition() {
            loupeControls.offsetX.value = 0;
            loupeControls.offsetY.value = 0;
            onSettingChanged();
        }

        /**
         * 数値欄に∧∨・↑↓キーの範囲と入力確定をつなぐ
         * @param {object} fieldSpec - numberFieldSpecs の1件
         * @returns {void}
         */
        function wireNumberField(fieldSpec) {
            var stepOptions = fieldSpec.field.stepOptions;
            stepOptions.min = fieldSpec.min;
            stepOptions.max = fieldSpec.max;
            stepOptions.onStep = onSettingChanged;
            bindSteppedArrowKeys(fieldSpec.field, fieldSpec.field.stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 */
            fieldSpec.field.onChange = onSettingChanged;
        }

        var settingKey;
        for (settingKey in numberFieldSpecs) {
            wireNumberField(numberFieldSpecs[settingKey]);
        }
        for (settingKey in checkboxSpecs) {
            checkboxSpecs[settingKey].onClick = onSettingChanged;
        }

        /* コネクターはオンにした時点で取り込んでおく。プレビューにも出せて、読めないときはその場で分かる
           Importing the moment it is switched on lets the preview show it, and surfaces a failure right away */
        if (optionControls.addConnectorCheckbox) {
            optionControls.addConnectorCheckbox.onClick = function() {
                if (optionControls.addConnectorCheckbox.value && !ensureConnectorStyle()) {
                    optionControls.addConnectorCheckbox.value = false;
                    alert(getLabel('alert.connectorUnavailable') + '\n' + CONNECTOR_STYLE_NAME);
                }
                onSettingChanged();
            };
        }

        /* スライダーはドラッグ中の作り直しが重いので、離したときだけ描き直す
           Rebuilding while dragging is heavy, so only refresh once the slider is released */
        loupeControls.offsetX.onChange = onSettingChanged;
        loupeControls.offsetY.onChange = onSettingChanged;
        loupeControls.resetPositionButton.onClick = resetLoupePosition;

        effectControls.noneRadio.onClick = onSettingChanged;
        effectControls.blurInsideRadio.onClick = onSettingChanged;
        effectControls.blurOutsideRadio.onClick = onSettingChanged;
        effectControls.spotlightRadio.onClick = onSettingChanged;
        buttonRowControls.previewCheckbox.onClick = refreshPreview;

        presetControls.dropdown.onChange = function() {
            /* 末尾の「カスタム」を選んだときは何も流し込まない / picking Custom applies nothing */
            var presetIndex = presetControls.dropdown.selection.index;
            if (presetIndex >= presetList.length) return;

            applyPreset(presetList[presetIndex]);
            updateControlStates();
            refreshPreview();
        };

        presetControls.saveButton.onClick = function() {
            var settingsToSave = collectSettings();
            if (!saveSettingsToFile(settingsToSave)) return;

            /* 保存した設定をドロップダウンの先頭に反映する。既にあれば中身だけ差し替える
               Reflect it at the top of the dropdown, replacing the entry when one is already there */
            settingsToSave.labelKey = "preset.saved";
            if (presetList.length > 0 && presetList[0].labelKey === "preset.saved") {
                presetList[0] = settingsToSave;
            } else {
                presetList.unshift(settingsToSave);
                presetControls.dropdown.add('item', getLabel('preset.saved'), 0);
            }
            presetControls.dropdown.selection = 0;
        };

        buttonRowControls.btnOk.onClick = function() {
            /* プレビューを消してから確定実行 / drop the preview before the real run */
            clearPreview();
            buildWithCurrentSettings(true);
            /* 確定できてから閉じる印を立てる。途中で失敗したらプレビュー扱いのまま後片付けできる
               Flag it only once the build succeeded, so a failure still cleans up as a preview */
            isClosedByButton = true;
            app.redraw();
            dialog.close();
        };

        buttonRowControls.btnCancel.onClick = function() {
            isClosedByButton = true;
            restoreOriginalSelection();
            dialog.close();
        };

        /* 閉じるボタンで閉じられたときもプレビューを後片付けする
           Clean the preview up when the dialog is dismissed by its close box */
        dialog.onClose = function() {
            if (isClosedByButton) return;
            restoreOriginalSelection();
        };

        /* ダイアログを開いた時点の状態を整えてからプレビューを表示
           Settle the control states, then show the preview as the dialog opens */
        applyPreset(presetList[findDefaultPresetIndex(presetList)]);
        updateControlStates();
        refreshPreview();
        dialog.show();
    }

    /**
     * 選択範囲のエッジとバウンディングボックスの表示を切り替える
     * どちらもトグルなので、実行時と終了時に1回ずつ呼べば元の状態に戻る
     * @returns {void}
     */
    function toggleEdgesAndBoundingBox() {
        app.executeMenuCommand('edge');
        app.executeMenuCommand('AI Bounding Box Toggle');
    }

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        doc = app.activeDocument;

        var pickedItems = pickArtworkAndMaskPath(doc.selection);
        if (pickedItems.artworkCount !== 1 || pickedItems.maskPathCount !== 1) {
            alert(getLabel('alert.needArtworkAndPath'));
            return;
        }

        originalArtwork = pickedItems.artwork;
        originalMaskPath = pickedItems.maskPath;

        /* 実行中はエッジとバウンディングボックスを隠して、プレビューを仕上がりのまま確認できるようにする
           Hide the selection edges and the bounding box while the script runs, so the preview reads as the result */
        toggleEdgesAndBoundingBox();
        showSettingsDialog();
        toggleEdgesAndBoundingBox();
    }

    main();
})();
