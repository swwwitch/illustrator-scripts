#target illustrator
#targetengine "SplitForTwoEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

1つのオブジェクトを選択して実行すると、その外接矩形を左右または上下に2分割し、2色の背景に置き換えます。
［バランス］パネルで、左右（または上下）の幅と比率を数値入力とスライダーで調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitForTwo.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1b7b8759e53b

### Overview

With a single object selected, splits its bounding box left/right or top/bottom and replaces the object with a two-color background.
The Balance panel adjusts the width and ratio of each half with a field and a slider.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitForTwo.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitForTwo";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.10.4";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitForTwo.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitForTwo.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1b7b8759e53b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    // 色は "RRGGBB"（RGB の16進数）か "cmyk:C,M,Y,K"（0〜100）で書く
    // Colors are written as "RRGGBB" (RGB hex) or "cmyk:C,M,Y,K" (0-100)

    var DEFAULT_FIRST_COLOR = "DCDCDC";    /* 左（上）側の塗りの初期値 / initial fill of the left (top) side */
    var DEFAULT_SECOND_COLOR = "808080";   /* 右（下）側の塗りの初期値 / initial fill of the right (bottom) side */
    var DEFAULT_STROKE_COLOR = "000000";   /* 外枠と区切り線の色の初期値 / initial color of the frame and divider */
    var DEFAULT_STROKE_PT = 1;             /* 線幅の初期値（pt）/ initial stroke width (pt) */
    var DEFAULT_CORNER_PT = 0;             /* 角丸の初期値（pt）/ initial corner radius (pt) */
    var DEFAULT_GROUP_ITEMS = true;        /* ［グループ化］の初期値 / initial state of Group items */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 18;                     /* ダイアログ外周の余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];        /* パネル・タブの余白 [左,上,右,下] / panel and tab margins */
    var BALANCE_SPACING = 6;                     /* バランスの行の間隔 / spacing between the balance rows */
    var BALANCE_SLIDER_MARGINS = [0, 10, 0, 0];  /* 幅スライダーの上の余白 / space above the width slider */
    var BALANCE_SLIDER_SIZE = [180, 20];         /* 幅スライダーの大きさ / width slider size */
    var WIDTH_INPUT_CHARS = 5;                   /* 幅の入力欄の文字数 / characters for the width fields */
    var PERCENT_INPUT_CHARS = 3;                 /* ％の入力欄の文字数 / characters for the percent fields */
    var STROKE_INPUT_CHARS = 3;                  /* 線幅の入力欄の文字数 / characters for the stroke field */
    var CORNER_INPUT_CHARS = 4;                  /* 角丸の入力欄の文字数 / characters for the corner fields */
    var SWATCH_SIZE = [20, 20];                  /* 色見本の大きさ / swatch size */
    var PRESET_DROPDOWN_WIDTH = 120;             /* プリセットのドロップダウンの幅 / preset dropdown width */
    var PRESET_SAVE_BUTTON_WIDTH = 50;           /* ［保存］ボタンの幅 / Save button width */
    var PRESET_EXPORT_BUTTON_WIDTH = 60;         /* ［書き出し］ボタンの幅 / Export button width */
    var PRESET_NAME_CHARS = 20;                  /* プリセット名の入力欄の文字数 / characters for the preset name field */

    /**
     * 共通レイアウトのパネルを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} titleText - パネルのタイトル
     * @param {string} [childAlignment] - 子の横方向の揃え（既定は "left"）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleText, childAlignment) {
        var newPanel = parent.add("panel", undefined, titleText);
        newPanel.orientation = "column";
        newPanel.alignChildren = [childAlignment || "left", "top"];
        newPanel.alignment = ["fill", "top"];
        newPanel.margins = PANEL_MARGINS;
        return newPanel;
    }

    /**
     * 横並びの行グループを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @returns {Group} 追加した行グループ
     */
    function addRow(parent) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * 行に「∧∨＋入力欄」を追加する。∧∨・↑↓キーで増減したら、入力欄の onStepped（イベント設定時に入れる）を呼ぶ
     * @param {Group} rowGroup - 追加先の行
     * @param {string} initialText - 初期値
     * @param {Object} stepOptions - min / max / integer（addStepper() に渡す）
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSteppedInput(rowGroup, initialText, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = rowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        stepOptions.onStep = function (steppedInput) {
            if (typeof steppedInput.onStepped !== "function") return;
            /* プレビューの描き直しで失敗しても入力は続けられるようにする / keep the field usable if redrawing the preview fails */
            try { steppedInput.onStepped(); } catch (e) { }
        };
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * addSteppedInput() で作った入力欄の有効・無効を、∧∨のディム表示とあわせて切り替える
     * @param {EditText} numberInput - 対象の入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * 縦並びの列グループを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} childAlignment - 子の横方向の揃え
     * @returns {Group} 追加した列グループ
     */
    function addColumn(parent, childAlignment) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = [childAlignment, "top"];
        return columnGroup;
    }

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} labelKey - ラベルの LABELS キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @param {boolean} checked - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parent, labelKey, tooltipKey, checked) {
        var newCheckbox = parent.add("checkbox", undefined, getLabel(labelKey));
        newCheckbox.helpTip = getLabel(tooltipKey);
        newCheckbox.value = !!checked;
        return newCheckbox;
    }

    /**
     * ツールチップ付きのラジオボタンを追加する
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelKey - ラベルの LABELS キー
     * @param {string} tooltipKey - ツールチップの LABELS キー
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parent, labelKey, tooltipKey) {
        var newRadio = parent.add("radiobutton", undefined, getLabel(labelKey));
        newRadio.helpTip = getLabel(tooltipKey);
        return newRadio;
    }

    /**
     * 文字列の中でいちばん長いものを返す（あとで文言を切り替えるラベルの幅を確保するため）
     * @param {string[]} texts - 候補の文字列
     * @returns {string} いちばん長い文字列
     */
    function getLongestText(texts) {
        var longestText = "";
        for (var i = 0; i < texts.length; i++) {
            if (texts[i].length > longestText.length) longestText = texts[i];
        }
        return longestText;
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

    /**
     * 環境設定の単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} pt の値
     */
    function unitToPt(value, prefKey) {
        return value * getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * pt の値を環境設定の単位に換算する
     * @param {number} valuePt - pt の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} 単位の値
     */
    function ptToUnit(valuePt, prefKey) {
        return valuePt / getUnitInfo(prefKey).pointsPerUnit;
    }

    /**
     * 単位の大きさに合わせて丸める（pt・Q は小数1桁、mm は2桁、in・cm・ft など1単位が10pt以上は3桁）
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {number} 丸めた値
     */
    function roundUnitValue(value, prefKey) {
        var pointsPerUnit = getUnitInfo(prefKey).pointsPerUnit;
        var scale = (pointsPerUnit >= 10) ? 1000 : (pointsPerUnit >= 2) ? 100 : 10;
        return Math.round(value * scale) / scale;
    }

    /**
     * 単位の値を入力欄の文字列にする
     * @param {number} value - 単位の値
     * @param {string} prefKey - 環境設定キー
     * @returns {string} 表示用の文字列
     */
    function formatUnitValue(value, prefKey) {
        return String(roundUnitValue(value, prefKey));
    }

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
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（uiLang の定義より後、ダイアログを作る関数より前）に貼る。
    //    識別子はすべて KEY_SHORTCUT_* / *KeyShortcut* の名前。uiLang はコピー先のものをそのまま使う
    // 2. コントロールをすべて作り、onClick を付けたあとで1回だけ呼ぶ（keydown はウィンドウに1つ）
    //      addKeyShortcuts(dialog, {
    //          "L": alignLeftRadio,                    … ラジオ：選んで onClick
    //          "P": previewCheckbox,                   … チェックボックス：反転して onClick
    //          "Shift+R": btnReset,                    … ボタン：onClick（無ければ notify）
    //          "G": function () { toggleGuides(); },   … 関数：呼ぶだけ
    //          "Escape": { target: function () { palette.close(); }, inFields: true }
    //      }, { numericFields: [widthInput, heightInput], afterKey: updatePreview });
    //    キーは keyName と同じ綴り（"A"〜"Z"・"1"・"Semicolon"・"Escape" など。大小文字は区別しない）。
    //    修飾キーは "Shift+" / "Alt+"（option）/ "Cmd+"（⌘、Windows は Ctrl）を前に付ける
    // 3. 修飾キーは完全一致。"R" は Shift・option・⌘ を押しながらでは効かない（⌘C などを横取りしない）。
    //    Shift＋R に別の動作を付けるときは "Shift+R" を並べる
    // 4. 入力欄（edittext）・ドロップダウン・リストにフォーカスがあるときは効かない（文字は普通に入る）。
    //    数値だけの欄で効かせたいときは options.numericFields に並べる（押した文字は欄に入らない）。
    //    入力中でも効かせたいキーは { target: …, inFields: true } にする（Esc で閉じる、option＋数字など）
    // 5. 無効・非表示のコントロールは、親のパネルやグループが無効なときも含めて何もしない
    //    （親を無効にしても子の enabled は true のまま、のため親までたどる）
    // 6. 関数の戻り値：false はこのキーを使わない（文字をそのまま通す）。コントロールを返すと、そのコントロールを
    //    押したことにする（向きによってラジオが変わるときなど）。それ以外は処理済み
    // 7. ツールチップへのキー表記は options.showInTip: true で「…（L）」「… (L)」を末尾に足す。
    //    LABELS の tooltip にすでにキーを書いてあるスクリプトでは付けない（同じキーが書いてあれば二重には足さない）
    // 8. 既存の keydown 処理（bindKeyboardShortcuts・addAlignKeyHandler など）と入力欄の focus／blur による抑止は消して、これに寄せる。
    //    ↑↓キー（StepperButtons の bindSteppedArrowKeys）はそのまま残す
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトの分割背景を作成", en: "Create Split Background" },
            presetSave: { ja: "プリセット保存", en: "Save Preset" }
        },
        panel: {
            splitDirection: { ja: "分割方法", en: "Split Direction" },
            balance: { ja: "バランス", en: "Balance" },
            fill: { ja: "塗り", en: "Fill" },
            stroke: { ja: "線", en: "Stroke" },
            cornerRadius: { ja: "角丸", en: "Corner radius" }
        },
        radio: {
            splitLR: { ja: "左右", en: "Left/Right" },
            splitTB: { ja: "上下", en: "Top/Bottom" }
        },
        checkbox: {
            square: { ja: "正方形", en: "Square" },
            overallFrame: { ja: "外枠", en: "Outer frame" },
            divider: { ja: "区切り線", en: "Divider" },
            pillShape: { ja: "ピル形状", en: "Pill shape" },
            cornerTL: { ja: "左上", en: "TL" },
            cornerBL: { ja: "左下", en: "BL" },
            cornerTR: { ja: "右上", en: "TR" },
            cornerBR: { ja: "右下", en: "BR" },
            groupItems: { ja: "グループ化", en: "Group items" }
        },
        fieldLabel: {
            left: { ja: "左", en: "Left" },
            right: { ja: "右", en: "Right" },
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" },
            strokeWidth: { ja: "線幅", en: "Stroke width" },
            color: { ja: "カラー", en: "Color" },
            preset: { ja: "プリセット", en: "Preset" },
            presetName: { ja: "プリセット名を入力", en: "Enter preset name" }
        },
        dropdown: {
            presetPlaceholder: { ja: "---", en: "---" }
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
            splitLR: { ja: "左右に分割します。", en: "Splits the shape left and right." },
            splitTB: { ja: "上下に分割します。", en: "Splits the shape top and bottom." },
            width: {
                ja: "この側の幅です。0 なら％の指定が使われます。",
                en: "Width of this side. 0 means the percentage below is used instead."
            },
            percent: { ja: "この側が占める割合（％）です。", en: "Share of the whole this side takes, in percent." },
            square: { ja: "この側を正方形にします。", en: "Makes this side a square." },
            widthSlider: { ja: "左右の割合をドラッグで決めます。", en: "Drag to set the split ratio." },
            fillSide: {
                ja: "この側に塗りを付けます。色は右の欄で指定します。",
                en: "Fills this side. The field on the right sets the color."
            },
            overallFrame: {
                ja: "分割した全体を1つの枠線で囲みます。",
                en: "Draws a single frame around the whole split shape."
            },
            divider: { ja: "分割の境目にケイ線を引きます。", en: "Draws a rule along the split." },
            strokeWidth: { ja: "ケイ線の太さです。", en: "Weight of the rules." },
            pillShape: { ja: "左右の端を半円にして、丸いピル型にします。", en: "Rounds both ends into a pill shape." },
            cornerLink: { ja: "4つの角丸の値を連動させます。", en: "Links the four corner radii together." },
            corner: {
                ja: "この角を丸めます。半径は右の欄で指定します。",
                en: "Rounds this corner. The field on the right sets the radius."
            },
            groupItems: {
                ja: "作ったオブジェクトを1つのグループにまとめます。",
                en: "Groups the resulting objects together."
            },
            swatch: { ja: "クリックすると色を選べます。", en: "Click to choose a color." },
            preset: { ja: "保存した設定を読み込みます。", en: "Loads a saved set of settings." },
            presetName: { ja: "保存する設定の名前です。", en: "Name the settings are saved under." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            save: { ja: "保存", en: "Save" },
            exportSettings: { ja: "書き出し", en: "Export" }
        },
        alert: {
            openDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectOneItem: { ja: "1つのオブジェクトを選択してください。", en: "Please select one object." },
            exportedSettings: { ja: "現在の設定を書き出しました：", en: "Exported current settings:" },
            exportFailed: { ja: "ファイルの書き出しに失敗しました。", en: "Failed to export the file." }
        }
    };

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は SETTINGS_STORE_* / createSettingsStore / readSettingsLegacyFile / readSettingsLegacyPreference / settingsStore*
    // 2. 寿命は今のスクリプトに合わせて選ぶ。
    //      "session"    … $.global に置く。Illustrator を終了するまで残る。#targetengine が必須（無いと毎回消える）
    //      "persistent" … Folder.userData/illustrator-scripts/<storeName>.json に書く。再起動しても残る
    //    storeName はふつう SCRIPT_NAME。ダイアログの位置は DialogPosition の部品が持つので、ここには入れない
    // 3. 既定値を1か所にまとめ、load で受け取る。戻り値は毎回新しいオブジェクト（書き換えても保存されない）
    //      var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
    //      var DEFAULT_SETTINGS = { widthPt: 10, addFrame: true, modeKey: "fit", corners: { tl: 0, tr: 0 } };
    //      var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
    //      …OK で閉じたら…
    //      settingsStore.save({ widthPt: …, addFrame: …, modeKey: …, corners: { tl: …, tr: … } });
    //    型は既定値に合わせる（数値の既定値には "12" も 12 として読む。真偽は "1"/"0"/"true"/"false" も読む）。
    //    合わない値・既定値に無い項目は捨てて既定値を使う。{} と null の既定値は中身を問わずそのまま受け取る
    //    （名前をキーにしたプリセット集など）。配列は配列ならそのまま受け取る
    // 4. 保存できるのは文字列・数値・真偽・null と、その配列・入れ子のオブジェクトだけ。
    //    DOM オブジェクト・File・関数は入れない（パスは fsName の文字列で持つ）。長さは pt で持つ
    // 5. 旧形式の設定を読み継ぐときは、3つ目の引数に legacy 関数を渡す。
    //    新しい保存が1度も無いとき（ファイルが無い・$.global に無い）だけ呼ばれ、戻り値を保存値として既定値と突き合わせる。
    //    旧ファイル・旧キーは消さない。キー名が変わったときは legacy の中で詰め替える
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");  … key=value / toSource / JSON を自動判別
    //      } });
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyPreference("SmartTextFindReplace/settings");  … app.preferences の文字列
    //      } });
    // 6. clear() は保存を消す。legacy を渡したストアでは空の保存（{}）を書き、旧設定が戻ってこないようにする
    // 7. 失敗は例外にせず、load は既定値、save は false を返す（$.writeln に理由を出す）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            return textFile.read().replace(/^﻿/, "");
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
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // セッション設定 / Session settings
    // =========================================
    // Illustrator の起動中だけダイアログの値を保持する（長さは単位を変えても崩れないよう pt で持つ）
    // Dialog values are kept while Illustrator is running; lengths are stored in pt so unit changes do not skew them

    /* 旧版の $.global のキー（1度だけ読み継ぐ）/ The old $.global keys (read once) */
    var LEGACY_SESSION_KEY = "SplitForTwo_settings";
    var LEGACY_PRESETS_KEY = "SplitForTwo_presets";

    var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () { return $.global[LEGACY_SESSION_KEY] || null; }
    });
    /* プリセットは名前をキーにした集まり（既定値 {} で中身を問わず受け取る）/ Presets: a map keyed by name (the {} default accepts any content) */
    var presetSettingsStore = createSettingsStore(SCRIPT_NAME + "Presets", "session", {
        legacy: function () { return $.global[LEGACY_PRESETS_KEY] || null; }
    });

    /**
     * セッションに残したダイアログの値を返す（欠けている項目は初期値で補う）
     * @returns {Object} セッション設定（書き換えても保存されない。saveDialogToSession() で保存する）
     */
    function getSessionSettings() {
        var defaults = {
            fillFirst: true,
            fillSecond: true,
            firstColor: DEFAULT_FIRST_COLOR,
            secondColor: DEFAULT_SECOND_COLOR,
            overallFrame: false,
            divider: false,
            strokeWidthPt: DEFAULT_STROKE_PT,
            strokeColor: DEFAULT_STROKE_COLOR,
            pillShape: false,
            cornerLink: false,
            cornerEnabled: { tl: false, bl: false, tr: false, br: false },
            cornerRadiusPt: { tl: DEFAULT_CORNER_PT, bl: DEFAULT_CORNER_PT, tr: DEFAULT_CORNER_PT, br: DEFAULT_CORNER_PT }
        };
        return settingsStore.load(defaults);
    }

    /**
     * 保存したプリセットを返す
     * @returns {Object} プリセット名をキーにした設定の集まり（書き換えても保存されない）
     */
    function getPresetStore() {
        return presetSettingsStore.load({});
    }

    /**
     * プリセットを1つ保存する（同じ名前は上書き）
     * @param {string} presetName - プリセット名
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function savePreset(presetName, presetData) {
        var presetMap = getPresetStore();
        presetMap[presetName] = presetData;
        presetSettingsStore.save(presetMap);
    }

    // =========================================
    // ダイアログ共通 / Dialog helpers
    // =========================================

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

    /**
     * コントロールを描き直させる（onDraw をもう一度呼ばせる）
     * @param {Object} control - 対象のコントロール
     * @returns {void}
     */
    function repaintControl(control) {
        control.hide();
        control.show();
    }

    /**
     * コントロールを単色で塗る（onDraw の中で使う）
     * @param {Object} control - 描画するコントロール
     * @param {{r: number, g: number, b: number}} rgb - 0〜255 の RGB
     * @param {boolean} withBorder - 黒い枠線を付けるか
     * @returns {void}
     */
    function paintSolidColor(control, rgb, withBorder) {
        var graphics = control.graphics;
        var brush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
        /* newPath() を呼ばないとパスが積み重なる / without newPath() the paths accumulate */
        graphics.newPath();
        graphics.rectPath(0, 0, control.size[0], control.size[1]);
        graphics.fillPath(brush);
        if (!withBorder) return;
        var borderPen = graphics.newPen(graphics.PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
        graphics.newPath();
        graphics.rectPath(0, 0, control.size[0], control.size[1]);
        graphics.strokePath(borderPen);
    }

    /**
     * 値を範囲内に収める
     * @param {number} value - 値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 収めた値
     */
    function clampNumber(value, minValue, maxValue) {
        if (value < minValue) return minValue;
        if (value > maxValue) return maxValue;
        return value;
    }

    // =========================================
    // 実行時の状態 / Run state
    // =========================================

    var doc = null;           /* 対象ドキュメント / target document */
    var targetItem = null;    /* 選択したオブジェクト（確定すると削除する）/ the selected object, removed on OK */
    var targetBounds = null;  /* targetItem の外接矩形 [左, 上, 右, 下] / bounds of targetItem [left, top, right, bottom] */

    /* プレビュー専用レイヤー（ダイアログを閉じると消す）/ Dedicated preview layer, removed when the dialog closes */
    var PREVIEW_LAYER_NAME = "__SplitForTwo__PreviewLayer__";

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。使わない関数も消さずに残してよい（互いに呼び合う）。
    //    識別子は SELECTION_ITEMS_TOLERANCE / normalizeSelectionItems / resolveTextRangeFrame /
    //    collectSelectionItems / getTextFrameKindKey / collectSelectionTextFrames / collectSelectionPathItems /
    //    isClipMaskItem / getClipMaskItem / hasClippedDescendant / readUsePreviewBoundsPreference /
    //    getClipAwareBounds / filterMeasurableChildren / getClipAwareUnionBounds / isNearlySameCoordinate / areBoundsNearlyEqual
    // 2. 選択は normalizeSelectionItems(doc.selection) で配列にする。文字カーソルの選択（TextRange）は
    //    配列ではなく1個で返り、しかも .length（文字数）を持つので、length だけで配列と見なさない
    // 3. テキストフレーム:
    //      var frames = collectSelectionTextFrames(doc.selection);                           // 全種類
    //      var frames = collectSelectionTextFrames(doc.selection, { kinds: ["point", "path"] });
    //    パス:
    //      var paths = collectSelectionPathItems(doc.selection);                             // 複合パスは中のパスへ
    //      var paths = collectSelectionPathItems(doc.selection, { compoundPaths: "whole", skipClipMasks: true });
    //    それ以外は collectSelectionItems(source, { accept: function (item) { … } }) で条件を書く
    // 4. 並びは選択と同じ前面→背面（グループの中も pageItems の順）。重なり順を使う処理はこの順を前提にしてよい
    // 5. doc.selection に代入し直す配列は skipLocked / skipHidden を true にする。
    //    ロック・非表示を選択に代入すると例外になり、中の子が選択に残る
    // 6. 境界は getClipAwareBounds(item, usePreviewBounds) / getClipAwareUnionBounds(items, usePreviewBounds)。
    //    usePreviewBounds を省くと環境設定の［プレビュー境界を使用］に従う。返り値は [左, 上, 右, 下] の新しい配列
    //    （書き換えても元のオブジェクトに影響しない）。測れないときは null
    // 7. 座標の一致・前後の判定は isNearlySameCoordinate / areBoundsNearlyEqual で許容値を挟む
    //    （吸着させた辺とガイドは 1e-12 ほどずれる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 計測 / Measuring
    // =========================================

    /**
     * 削除する（createOutline() に消費された複製など、すでに無いものは無視する）
     * @param {PageItem} pageItem - 削除するオブジェクト
     * @returns {void}
     */
    function removeItem(pageItem) {
        if (!pageItem) return;
        try {
            pageItem.remove();
        } catch (e) {
            /* すでに無い / already gone */
        }
    }

    /**
     * テキストを複製してアウトライン化し、その外接矩形を返す（サイドベアリングを含めない）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[]|null} 外接矩形。測れなければ null
     */
    function getOutlineBounds(textFrame) {
        var duplicateText = null;
        var outlineGroup = null;
        var outlineBounds = null;
        try {
            duplicateText = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATBEGINNING);
            outlineGroup = duplicateText.createOutline();
            outlineBounds = outlineGroup.geometricBounds;
        } catch (e) {
            /* 空白だけのテキストは中身の無いグループになり、geometricBounds が例外になる / whitespace-only text outlines to an empty group whose bounds throw */
            outlineBounds = null;
        } finally {
            removeItem(outlineGroup);
            /* createOutline() が成功していれば複製は消費済み / the duplicate is already consumed when createOutline() succeeds */
            removeItem(duplicateText);
        }
        return outlineBounds;
    }

    /**
     * オブジェクトの外接矩形を返す。テキストはアウトライン、クリップグループはマスクの範囲で測る
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number[]} 外接矩形 [左, 上, 右, 下]
     */
    function getItemBounds(pageItem) {
        var itemBounds = null;
        if (pageItem.typename === "TextFrame") {
            itemBounds = getOutlineBounds(pageItem);
        } else {
            /* クリップグループはマスクを測る（テキストのマスクもアウトラインで） / Measure a clip group by its mask (a text mask is outlined too) */
            var maskItem = getClipMaskItem(pageItem);
            if (maskItem) return getItemBounds(maskItem);
            itemBounds = getClipAwareBounds(pageItem, false);
        }
        if (!itemBounds) itemBounds = pageItem.geometricBounds;
        return [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * サブレイヤーをたどって最上位のレイヤーを返す
     * @param {Layer} layer - 対象のレイヤー
     * @returns {Layer} 最上位のレイヤー
     */
    function getTopLevelLayer(layer) {
        while (layer.parent.typename === "Layer") layer = layer.parent;
        return layer;
    }

    /**
     * プレビューレイヤーを探す
     * @returns {Layer|null} プレビューレイヤー。無ければ null
     */
    function findPreviewLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === PREVIEW_LAYER_NAME) return doc.layers[i];
        }
        return null;
    }

    /**
     * プレビューレイヤーを用意し、元のオブジェクトのレイヤーより背面に置く
     * @returns {Layer} プレビューレイヤー
     */
    function ensurePreviewLayer() {
        var previewLayer = findPreviewLayer();
        if (previewLayer) return previewLayer;

        previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;
        /* サブレイヤーなら最上位の親の背面へ / behind the top-level parent when the object sits on a sublayer */
        previewLayer.move(getTopLevelLayer(targetItem.layer), ElementPlacement.PLACEAFTER);
        return previewLayer;
    }

    /**
     * プレビューレイヤーを中身ごと削除する
     * @returns {void}
     */
    function removePreviewLayer() {
        var previewLayer = findPreviewLayer();
        if (previewLayer) previewLayer.remove();
    }

    // =========================================
    // カラー / Colors
    // =========================================

    /**
     * CMYK の色文字列か
     * @param {string} colorText - 色文字列
     * @returns {boolean} "cmyk:" で始まるか
     */
    function isCmykString(colorText) {
        return String(colorText).indexOf("cmyk:") === 0;
    }

    /**
     * "cmyk:C,M,Y,K" を数値に分ける
     * @param {string} colorText - 色文字列
     * @returns {{c: number, m: number, y: number, k: number}} CMYK（0〜100）
     */
    function parseCmykString(colorText) {
        var cmykParts = String(colorText).replace("cmyk:", "").split(",");
        return {
            c: parseFloat(cmykParts[0]) || 0,
            m: parseFloat(cmykParts[1]) || 0,
            y: parseFloat(cmykParts[2]) || 0,
            k: parseFloat(cmykParts[3]) || 0
        };
    }

    /**
     * CMYK から "cmyk:C,M,Y,K" を作る
     * @param {{c: number, m: number, y: number, k: number}} cmyk - CMYK（0〜100）
     * @returns {string} 色文字列
     */
    function cmykStringFromValues(cmyk) {
        return "cmyk:" + Math.round(cmyk.c) + "," + Math.round(cmyk.m) + "," + Math.round(cmyk.y) + "," + Math.round(cmyk.k);
    }

    /**
     * "RRGGBB" を RGB に分ける
     * @param {string} colorText - 16進数の色文字列（先頭の # は無視）
     * @param {string} fallbackHex - 読めないときに使う16進数
     * @returns {{r: number, g: number, b: number}} RGB（0〜255）
     */
    function parseHexRgb(colorText, fallbackHex) {
        var hex = String(colorText).replace(/^#/, "");
        if (!/^[0-9A-Fa-f]{6}$/.test(hex)) hex = fallbackHex;
        return {
            r: parseInt(hex.substring(0, 2), 16),
            g: parseInt(hex.substring(2, 4), 16),
            b: parseInt(hex.substring(4, 6), 16)
        };
    }

    /**
     * RGB から "RRGGBB" を作る
     * @param {{r: number, g: number, b: number}} rgb - RGB（0〜255）
     * @returns {string} 大文字の16進数
     */
    function rgbToHex(rgb) {
        var toHex2 = function (channelValue) {
            var hex = Math.round(channelValue).toString(16).toUpperCase();
            return (hex.length < 2) ? "0" + hex : hex;
        };
        return toHex2(rgb.r) + toHex2(rgb.g) + toHex2(rgb.b);
    }

    /**
     * CMYK を RGB に簡易換算する（プレビュー用の近似値）
     * @param {{c: number, m: number, y: number, k: number}} cmyk - CMYK（0〜100）
     * @returns {{r: number, g: number, b: number}} RGB（0〜255）
     */
    function cmykToRgbApprox(cmyk) {
        var black = 1 - cmyk.k / 100;
        return {
            r: Math.round(255 * (1 - cmyk.c / 100) * black),
            g: Math.round(255 * (1 - cmyk.m / 100) * black),
            b: Math.round(255 * (1 - cmyk.y / 100) * black)
        };
    }

    /**
     * RGB から色を作る
     * @param {{r: number, g: number, b: number}} rgb - RGB（0〜255）
     * @returns {RGBColor} 色
     */
    function makeRGBColor(rgb) {
        var rgbColor = new RGBColor();
        rgbColor.red = rgb.r;
        rgbColor.green = rgb.g;
        rgbColor.blue = rgb.b;
        return rgbColor;
    }

    /**
     * CMYK から色を作る
     * @param {{c: number, m: number, y: number, k: number}} cmyk - CMYK（0〜100）
     * @returns {CMYKColor} 色
     */
    function makeCMYKColor(cmyk) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cmyk.c;
        cmykColor.magenta = cmyk.m;
        cmykColor.yellow = cmyk.y;
        cmykColor.black = cmyk.k;
        return cmykColor;
    }

    /**
     * 色文字列から Illustrator の色を作る
     * @param {string} colorText - "RRGGBB" か "cmyk:C,M,Y,K"
     * @returns {RGBColor|CMYKColor} 色
     */
    function parseColorString(colorText) {
        if (isCmykString(colorText)) return makeCMYKColor(parseCmykString(colorText));
        return makeRGBColor(parseHexRgb(colorText, "CCCCCC"));
    }

    /**
     * 色文字列を画面表示用の RGB にする
     * @param {string} colorText - "RRGGBB" か "cmyk:C,M,Y,K"
     * @returns {{r: number, g: number, b: number}} RGB（0〜255）
     */
    function colorStringToRgb(colorText) {
        if (isCmykString(colorText)) return cmykToRgbApprox(parseCmykString(colorText));
        return parseHexRgb(colorText, "000000");
    }

    // =========================================
    // カラー選択 / Color picker
    // =========================================

    /**
     * Illustrator 標準のカラーピッカーを開いて色を選ぶ。
     * OK なら選んだ色がドキュメントのカラーモードの型で返り、キャンセルなら渡した色がそのまま返る
     * @param {string} initialColor - 開いたときの色（"RRGGBB" か "cmyk:C,M,Y,K"）
     * @returns {string|null} 選んだ色の文字列。キャンセルしたとき・色が変わらないときは null
     */
    function showColorPickerDialog(initialColor) {
        var pickedColor = app.showColorPicker(parseColorString(initialColor));
        var pickedText = null;
        if (pickedColor.typename === "CMYKColor") {
            pickedText = cmykStringFromValues({ c: pickedColor.cyan, m: pickedColor.magenta, y: pickedColor.yellow, k: pickedColor.black });
        } else if (pickedColor.typename === "GrayColor") {
            pickedText = cmykStringFromValues({ c: 0, m: 0, y: 0, k: pickedColor.gray });
        } else if (pickedColor.typename === "RGBColor") {
            pickedText = rgbToHex({ r: pickedColor.red, g: pickedColor.green, b: pickedColor.blue });
        }
        return (pickedText === initialColor) ? null : pickedText;
    }

    /**
     * クリックでカラー選択を開く色見本を追加する
     * @param {Group} parent - 追加先
     * @param {string} initialColor - 初期の色（"RRGGBB" か "cmyk:C,M,Y,K"）
     * @returns {Object} getColor()・setColor() と、色を変えたときに呼ぶ onChange を持つオブジェクト
     */
    function addColorSwatch(parent, initialColor) {
        var colorSwatch = { colorText: initialColor, onChange: null };
        var swatchImage = parent.add("image", undefined, undefined, { name: "swatch" });
        swatchImage.helpTip = getLabel("tooltip.swatch");
        swatchImage.preferredSize = SWATCH_SIZE;
        swatchImage.onDraw = function () { paintSolidColor(this, colorStringToRgb(colorSwatch.colorText), true); };

        colorSwatch.getColor = function () { return colorSwatch.colorText; };
        colorSwatch.setColor = function (colorText) {
            colorSwatch.colorText = colorText;
            repaintControl(swatchImage);
        };

        swatchImage.addEventListener("click", function () {
            var pickedColor = showColorPickerDialog(colorSwatch.colorText);
            if (pickedColor === null) return;
            colorSwatch.setColor(pickedColor);
            if (typeof colorSwatch.onChange === "function") colorSwatch.onChange();
        });
        return colorSwatch;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    var CORNER_KEYS = ["tl", "tr", "bl", "br"];  /* 角の並び（外枠で使う順）/ corner order used for the outer frame */

    /**
     * パスのアンカーの選択をすべて外す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function deselectAnchors(pathItem) {
        var points = pathItem.pathPoints;
        for (var i = 0; i < points.length; i++) {
            points[i].selected = PathPointSelection.NOSELECTION;
        }
    }

    /**
     * 外接矩形の角の座標を返す
     * @param {number[]} bounds - 外接矩形 [左, 上, 右, 下]
     * @param {string} cornerKey - "tl" | "tr" | "bl" | "br"
     * @returns {number[]} 角の座標 [x, y]
     */
    function getCornerPoint(bounds, cornerKey) {
        var x = (cornerKey === "tl" || cornerKey === "bl") ? bounds[0] : bounds[2];
        var y = (cornerKey === "tl" || cornerKey === "tr") ? bounds[1] : bounds[3];
        return [x, y];
    }

    /**
     * 指定した角にいちばん近いアンカーを選択する
     * @param {PathItem} pathItem - 対象のパス
     * @param {string[]} cornerKeys - 選ぶ角（"tl" | "tr" | "bl" | "br"）
     * @returns {boolean} 1つでも選択できたか
     */
    function selectCornerAnchors(pathItem, cornerKeys) {
        var points = pathItem.pathPoints;
        if (points.length < 4) return false;
        var bounds = pathItem.geometricBounds;
        deselectAnchors(pathItem);

        var anySelected = false;
        for (var c = 0; c < cornerKeys.length; c++) {
            var cornerPoint = getCornerPoint(bounds, cornerKeys[c]);
            var nearestIndex = -1;
            var nearestDistance = Infinity;
            for (var j = 0; j < points.length; j++) {
                /* 別の角で選んだアンカーは除く / skip anchors already picked for another corner */
                if (points[j].selected === PathPointSelection.ANCHORPOINT) continue;
                var dx = points[j].anchor[0] - cornerPoint[0];
                var dy = points[j].anchor[1] - cornerPoint[1];
                var squaredDistance = dx * dx + dy * dy;
                if (squaredDistance < nearestDistance) {
                    nearestDistance = squaredDistance;
                    nearestIndex = j;
                }
            }
            if (nearestIndex >= 0) {
                points[nearestIndex].selected = PathPointSelection.ANCHORPOINT;
                anySelected = true;
            }
        }
        return anySelected;
    }

    /**
     * パスの指定した角を、同じ半径でまとめて丸める
     * @param {PathItem} pathItem - 対象のパス
     * @param {string[]} cornerKeys - 丸める角（"tl" | "tr" | "bl" | "br"）
     * @param {number} radiusPt - 角丸の半径（pt）
     * @returns {void}
     */
    function roundPathCorners(pathItem, cornerKeys, radiusPt) {
        if (!(radiusPt > 0)) return;
        if (selectCornerAnchors(pathItem, cornerKeys)) {
            try {
                roundAnyCorner([pathItem], { rr: radiusPt });
            } catch (e) {
                /* 幅0の長方形などで計算に失敗しても、角のまま描く / keep square corners if rounding fails, e.g. on a zero-width rectangle */
            }
        }
        deselectAnchors(pathItem);
    }

    /**
     * 有効な角を1つずつ、それぞれの半径で丸める
     * @param {PathItem|null} pathItem - 対象のパス（null なら何もしない）
     * @param {string[]} cornerKeys - 対象にする角
     * @param {Object} perCorner - 角ごとの設定（buildPerCornerConfig() の戻り値）
     * @returns {void}
     */
    function roundEachEnabledCorner(pathItem, cornerKeys, perCorner) {
        if (!pathItem) return;
        for (var i = 0; i < cornerKeys.length; i++) {
            var cornerSetting = perCorner[cornerKeys[i]];
            if (cornerSetting.enabled) roundPathCorners(pathItem, [cornerKeys[i]], cornerSetting.radius);
        }
    }

    /**
     * 「角を丸くする」効果を適用する
     * @param {PathItem} targetPath - 対象のパス
     * @param {number} cornerRadiusPt - 角丸の半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(targetPath, cornerRadiusPt) {
        if (!(cornerRadiusPt > 0)) return;
        try {
            targetPath.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerRadiusPt + ' "/></LiveEffect>');
        } catch (e) {
            /* 効果を適用できない環境では角のまま描く / keep square corners where the effect cannot be applied */
        }
    }

    /**
     * パスを選択して［ライブパスファインダー］の「合体」をかける（ピル形状の仕上げ）
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function applyLivePathfinderAdd(pathItem) {
        doc.selection = null;
        app.redraw();
        doc.selection = [pathItem];
        app.redraw();
        app.executeMenuCommand("Live Pathfinder Add");
        app.redraw();
    }

    /**
     * 線をケイ線の設定にする
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} strokeWidthPt - 線幅（pt）
     * @param {RGBColor|CMYKColor} strokeColor - 線の色
     * @returns {void}
     */
    function styleLine(pathItem, strokeWidthPt, strokeColor) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeWidth = strokeWidthPt;
        pathItem.strokeColor = strokeColor;
    }

    /**
     * 塗りの長方形を追加する
     * @param {PathItems} pathItems - 追加先
     * @param {number} top - 上
     * @param {number} left - 左
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @param {RGBColor|CMYKColor} fillColor - 塗りの色
     * @returns {PathItem} 追加した長方形
     */
    function addFillRect(pathItems, top, left, width, height, fillColor) {
        var fillRect = pathItems.rectangle(top, left, width, height);
        fillRect.fillColor = fillColor;
        fillRect.stroked = false;
        return fillRect;
    }

    /**
     * 塗りを最背面へ送る
     * @param {PathItem|null} fillRect - 塗りの長方形（null なら何もしない）
     * @returns {void}
     */
    function sendFillToBack(fillRect) {
        if (!fillRect) return;
        try {
            fillRect.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (e) {
            /* ピル形状の塗りはライブパスファインダーの中に入り、重ね順を変えられないことがある / a pill-shaped fill sits inside a live pathfinder, which can refuse a z-order change */
        }
    }

    /**
     * 背景全体の範囲と分割位置を求める
     * @param {Object} drawOptions - 描画の設定（collectDrawOptions() の戻り値）
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number, split: number}} 範囲と分割位置（左右なら X、上下なら Y）
     */
    function computeSplitLayout(drawOptions) {
        var splitLayout = {
            left: targetBounds[0],
            top: targetBounds[1],
            right: targetBounds[2],
            bottom: targetBounds[3]
        };
        splitLayout.width = splitLayout.right - splitLayout.left;
        splitLayout.height = splitLayout.top - splitLayout.bottom;
        if (drawOptions.vertical) {
            /* 上下：オフセットが正なら分割位置が下がる / top/bottom: a positive offset moves the split down */
            splitLayout.split = clampNumber(splitLayout.top - splitLayout.height / 2 - drawOptions.offsetPt,
                splitLayout.bottom, splitLayout.top);
        } else {
            splitLayout.split = clampNumber(splitLayout.left + splitLayout.width / 2 + drawOptions.offsetPt,
                splitLayout.left, splitLayout.right);
        }
        return splitLayout;
    }

    /**
     * 塗りの外側の角を丸める（角ごとの角丸か、ピル形状）
     * @param {PathItem|null} firstFill - 左（上）側の塗り
     * @param {PathItem|null} secondFill - 右（下）側の塗り
     * @param {Object} drawOptions - 描画の設定
     * @param {Object} splitLayout - computeSplitLayout() の戻り値
     * @returns {void}
     */
    function roundFillCorners(firstFill, secondFill, drawOptions, splitLayout) {
        var firstCornerKeys = drawOptions.vertical ? ["tl", "tr"] : ["tl", "bl"];
        var secondCornerKeys = drawOptions.vertical ? ["bl", "br"] : ["tr", "br"];

        if (drawOptions.perCorner) {
            roundEachEnabledCorner(firstFill, firstCornerKeys, drawOptions.perCorner);
            roundEachEnabledCorner(secondFill, secondCornerKeys, drawOptions.perCorner);
        } else if (drawOptions.pillShape) {
            /* 短辺の半分で丸める / round with half the short side */
            var pillRadiusPt = (drawOptions.vertical ? splitLayout.width : splitLayout.height) / 2;
            if (firstFill) roundPathCorners(firstFill, firstCornerKeys, pillRadiusPt);
            if (secondFill) roundPathCorners(secondFill, secondCornerKeys, pillRadiusPt);
            if (firstFill) applyLivePathfinderAdd(firstFill);
            if (secondFill) applyLivePathfinderAdd(secondFill);
        }
    }

    /**
     * 外枠の角を丸める（角ごとの角丸か、ピル形状）
     * @param {PathItem} frameRect - 外枠
     * @param {Object} drawOptions - 描画の設定
     * @param {Object} splitLayout - computeSplitLayout() の戻り値
     * @returns {void}
     */
    function roundFrameCorners(frameRect, drawOptions, splitLayout) {
        var perCorner = drawOptions.perCorner;
        if (perCorner) {
            var enabledKeys = [];
            for (var i = 0; i < CORNER_KEYS.length; i++) {
                if (perCorner[CORNER_KEYS[i]].enabled) enabledKeys.push(CORNER_KEYS[i]);
            }
            if (enabledKeys.length === 0) return;

            var firstRadiusPt = perCorner[enabledKeys[0]].radius;
            var allSameRadius = enabledKeys.length > 1;
            for (var j = 1; j < enabledKeys.length; j++) {
                if (perCorner[enabledKeys[j]].radius !== firstRadiusPt) {
                    allSameRadius = false;
                    break;
                }
            }

            if (allSameRadius && enabledKeys.length === 4) {
                /* 4つとも同じ半径なら効果で丸める / use the effect when all four radii match */
                applyRoundCornersEffect(frameRect, firstRadiusPt);
            } else if (allSameRadius) {
                roundPathCorners(frameRect, enabledKeys, firstRadiusPt);
            } else {
                roundEachEnabledCorner(frameRect, enabledKeys, perCorner);
            }
        } else if (drawOptions.pillShape) {
            roundPathCorners(frameRect, CORNER_KEYS, Math.min(splitLayout.width, splitLayout.height) / 2);
            applyLivePathfinderAdd(frameRect);
        }
    }

    /**
     * 2色の塗り・外枠・区切り線を描く
     * @param {Object} drawOptions - 描画の設定（collectDrawOptions() の戻り値）
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @returns {{firstFill: PathItem, secondFill: PathItem, frame: PathItem, divider: PathItem}} 描いたオブジェクト（描かなかったものは null）
     */
    function drawSplitBackground(drawOptions, targetLayer) {
        var splitLayout = computeSplitLayout(drawOptions);
        var pathItems = targetLayer.pathItems;
        var vertical = drawOptions.vertical;
        var firstFill = null;
        var secondFill = null;

        if (drawOptions.fillFirst) {
            firstFill = vertical
                ? addFillRect(pathItems, splitLayout.top, splitLayout.left,
                    splitLayout.width, splitLayout.top - splitLayout.split, drawOptions.firstColor)
                : addFillRect(pathItems, splitLayout.top, splitLayout.left,
                    splitLayout.split - splitLayout.left, splitLayout.height, drawOptions.firstColor);
        }
        if (drawOptions.fillSecond) {
            secondFill = vertical
                ? addFillRect(pathItems, splitLayout.split, splitLayout.left,
                    splitLayout.width, splitLayout.split - splitLayout.bottom, drawOptions.secondColor)
                : addFillRect(pathItems, splitLayout.top, splitLayout.split,
                    splitLayout.right - splitLayout.split, splitLayout.height, drawOptions.secondColor);
        }
        roundFillCorners(firstFill, secondFill, drawOptions, splitLayout);

        var frameRect = null;
        if (drawOptions.overallFrame) {
            frameRect = pathItems.rectangle(splitLayout.top, splitLayout.left, splitLayout.width, splitLayout.height);
            styleLine(frameRect, drawOptions.strokeWidthPt, drawOptions.strokeColor);
            roundFrameCorners(frameRect, drawOptions, splitLayout);
            frameRect.zOrder(ZOrderMethod.BRINGTOFRONT);
        }

        var dividerLine = null;
        if (drawOptions.divider) {
            dividerLine = pathItems.add();
            styleLine(dividerLine, drawOptions.strokeWidthPt, drawOptions.strokeColor);
            dividerLine.setEntirePath(vertical
                ? [[splitLayout.left, splitLayout.split], [splitLayout.right, splitLayout.split]]
                : [[splitLayout.split, splitLayout.top], [splitLayout.split, splitLayout.bottom]]);
            dividerLine.zOrder(ZOrderMethod.BRINGTOFRONT);
        }

        /* 塗りは最背面へ（左・上が右・下の上に来る順）/ send the fills to the back, keeping the left (top) one above */
        sendFillToBack(firstFill);
        sendFillToBack(secondFill);

        return { firstFill: firstFill, secondFill: secondFill, frame: frameRect, divider: dividerLine };
    }

    /**
     * 描いたオブジェクトを1つのグループにまとめる（線が上、塗りが下）
     * @param {Object} drawnItems - drawSplitBackground() の戻り値
     * @param {Layer} targetLayer - グループを作るレイヤー
     * @returns {void}
     */
    function groupDrawnItems(drawnItems, targetLayer) {
        var resultGroup = targetLayer.groupItems.add();
        var orderedItems = [drawnItems.frame, drawnItems.divider, drawnItems.firstFill, drawnItems.secondFill];
        try {
            for (var i = 0; i < orderedItems.length; i++) {
                if (orderedItems[i]) orderedItems[i].move(resultGroup, ElementPlacement.PLACEATEND);
            }
        } catch (e) {
            /* ライブパスファインダーの中のパスは動かせないことがある。動かせた分だけまとめる / a path inside a live pathfinder may refuse to move; keep what was grouped */
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューレイヤーの中身を消す
     * @returns {void}
     */
    function clearPreview() {
        var previewLayer = findPreviewLayer();
        if (!previewLayer) return;
        for (var i = previewLayer.pageItems.length - 1; i >= 0; i--) {
            previewLayer.pageItems[i].remove();
        }
    }

    /**
     * プレビューを描き直す
     * @param {Object|null} drawOptions - 描画の設定。null ならプレビューを消すだけ
     * @returns {void}
     */
    function drawPreview(drawOptions) {
        clearPreview();
        if (drawOptions) {
            try {
                drawSplitBackground(drawOptions, ensurePreviewLayer());
            } catch (e) {
                /* 描けなくてもダイアログの入力は続けられるようにする / keep the dialog usable even if drawing fails */
                clearPreview();
            }
            /* 角丸とライブパスファインダーが選択を残すので外す / rounding and the live pathfinder leave a selection behind */
            doc.selection = null;
        }
        app.redraw();
    }

    // =========================================
    // プリセット / Presets
    // =========================================
    // ［プリセット］の行は今は隠している（保存・読み込み・書き出しの処理は残す）
    // The preset row is hidden for now; saving, loading and exporting still work behind it

    /**
     * 保存したプリセットの名前を、名前順で返す
     * @returns {string[]} プリセット名
     */
    function getPresetNames() {
        var presetStore = getPresetStore();
        var presetNames = [];
        for (var presetName in presetStore) {
            if (presetStore.hasOwnProperty(presetName)) presetNames.push(presetName);
        }
        presetNames.sort();
        return presetNames;
    }

    /**
     * ダイアログの値をプリセットの形にまとめる
     * @param {Object} controls - コントロールの参照
     * @returns {Object} プリセット
     */
    function collectPresetData(controls) {
        return {
            splitDirection: controls.splitTBRadio.value ? "tb" : "lr",
            fillLeft: controls.fillFirstCheckbox.value,
            fillRight: controls.fillSecondCheckbox.value,
            colorLeft: controls.firstColorSwatch.getColor(),
            colorRight: controls.secondColorSwatch.getColor(),
            overallFrame: controls.overallFrameCheckbox.value,
            divider: controls.dividerCheckbox.value,
            strokeUnit: String(controls.strokeInput.text),
            strokeColor: controls.strokeColorSwatch.getColor(),
            cornerAuto: controls.pillCheckbox.value,
            widthOffset: clampOffset(controls, controls.balanceSlider.value)
        };
    }

    /**
     * プリセットをダイアログに読み込む
     * @param {Object} controls - コントロールの参照
     * @param {Object} presetData - プリセット
     * @returns {void}
     */
    function applyPresetData(controls, presetData) {
        var vertical = (presetData.splitDirection === "tb");
        controls.splitTBRadio.value = vertical;
        controls.splitLRRadio.value = !vertical;
        controls.fillFirstCheckbox.value = !!presetData.fillLeft;
        controls.fillSecondCheckbox.value = !!presetData.fillRight;
        controls.firstColorSwatch.setColor(presetData.colorLeft || DEFAULT_FIRST_COLOR);
        controls.secondColorSwatch.setColor(presetData.colorRight || DEFAULT_SECOND_COLOR);
        controls.overallFrameCheckbox.value = !!presetData.overallFrame;
        controls.dividerCheckbox.value = !!presetData.divider;
        controls.strokeInput.text = presetData.strokeUnit || formatUnitValue(ptToUnit(DEFAULT_STROKE_PT, "strokeUnits"), "strokeUnits");
        controls.strokeColorSwatch.setColor(presetData.strokeColor || DEFAULT_STROKE_COLOR);
        controls.pillCheckbox.value = !!presetData.cornerAuto;

        updateBalanceLabels(controls);
        updateStrokeEnabled(controls);
        updateCornerState(controls);
        updateBalanceRange(controls);
        if (presetData.widthOffset !== undefined) syncBalanceFromOffset(controls, presetData.widthOffset, null);
        refreshPreview(controls);
    }

    /**
     * プリセットのドロップダウンを作り直す
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function fillPresetDropdown(controls) {
        var presetDropdown = controls.presetDropdown;
        var presetNames = getPresetNames();
        presetDropdown.removeAll();
        presetDropdown.add("item", getLabel("dropdown.presetPlaceholder"));
        for (var i = 0; i < presetNames.length; i++) {
            presetDropdown.add("item", presetNames[i]);
        }
        presetDropdown.selection = 0;
    }

    /**
     * プリセット名を尋ねる
     * @returns {string|null} プリセット名。キャンセルか空なら null
     */
    function showPresetNameDialog() {
        var nameDialog = new Window("dialog", getLabel("dialog.presetSave"));
        nameDialog.orientation = "column";
        nameDialog.alignChildren = ["fill", "top"];
        nameDialog.add("statictext", undefined, labelText("fieldLabel.presetName"));
        var nameInput = nameDialog.add("edittext", undefined, "");
        nameInput.helpTip = getLabel("tooltip.presetName");
        nameInput.characters = PRESET_NAME_CHARS;
        nameInput.active = true;

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(nameDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        prepareDialogWindow(nameDialog, SCRIPT_NAME + "_presetName");
        if (nameDialog.show() !== 1) return null;

        var presetName = nameInput.text;
        if (!presetName || presetName === getLabel("dropdown.presetPlaceholder")) return null;
        return presetName;
    }

    /**
     * 今の設定をデスクトップのテキストファイルに書き出す（同名があれば番号を付ける）
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function exportSettingsToDesktop(controls) {
        var presetData = collectPresetData(controls);
        var settingLines = [
            "splitDirection=" + presetData.splitDirection,
            "fillLeft=" + (presetData.fillLeft ? "1" : "0"),
            "fillRight=" + (presetData.fillRight ? "1" : "0"),
            "colorLeft=" + presetData.colorLeft,
            "colorRight=" + presetData.colorRight,
            "overallFrame=" + (presetData.overallFrame ? "1" : "0"),
            "divider=" + (presetData.divider ? "1" : "0"),
            "strokeUnit=" + presetData.strokeUnit,
            "strokeColor=" + presetData.strokeColor,
            "cornerAuto=" + (presetData.cornerAuto ? "1" : "0"),
            "widthOffset=" + presetData.widthOffset
        ];

        var desktopPath = Folder.desktop.fsName;
        var baseName = SCRIPT_NAME + "_settings";
        var exportFile = new File(desktopPath + "/" + baseName + ".txt");
        for (var n = 1; exportFile.exists; n++) {
            exportFile = new File(desktopPath + "/" + baseName + "_" + n + ".txt");
        }

        exportFile.encoding = "UTF-8";
        if (exportFile.open("w")) {
            exportFile.write(settingLines.join("\n"));
            exportFile.close();
            alert(getLabel("alert.exportedSettings") + "\n" + exportFile.fsName);
        } else {
            alert(getLabel("alert.exportFailed"));
        }
    }

    /**
     * プリセットの行（ドロップダウン・保存・書き出し）を組み立てて隠す
     * @param {Window} settingsDialog - 追加先のダイアログ
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @returns {void}
     */
    function buildPresetRow(settingsDialog, controls) {
        var presetRowGroup = addRow(settingsDialog);
        presetRowGroup.alignment = ["center", "top"];
        presetRowGroup.add("statictext", undefined, labelText("fieldLabel.preset"));

        controls.presetDropdown = presetRowGroup.add("dropdownlist", undefined, []);
        controls.presetDropdown.helpTip = getLabel("tooltip.preset");
        controls.presetDropdown.preferredSize = [PRESET_DROPDOWN_WIDTH, -1];
        fillPresetDropdown(controls);

        controls.btnPresetSave = presetRowGroup.add("button", undefined, getLabel("button.save"));
        controls.btnPresetSave.preferredSize = [PRESET_SAVE_BUTTON_WIDTH, -1];
        controls.btnPresetExport = presetRowGroup.add("button", undefined, getLabel("button.exportSettings"));
        controls.btnPresetExport.preferredSize = [PRESET_EXPORT_BUTTON_WIDTH, -1];

        presetRowGroup.preferredSize = [0, 0];
        presetRowGroup.maximumSize = [0, 0];
        presetRowGroup.visible = false;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 左右・上下で切り替わるラベルのうち、いちばん長いもの（幅の確保用）
     * @returns {string} いちばん長いラベル
     */
    function getLongestSideLabel() {
        return getLongestText([
            labelText("fieldLabel.left"), labelText("fieldLabel.right"),
            labelText("fieldLabel.top"), labelText("fieldLabel.bottom")
        ]);
    }

    /**
     * 設定ダイアログを組み立てる（イベントは bindDialogEvents() で付ける）
     * @param {Object} session - セッションに残した前回の値
     * @returns {Object} ダイアログとコントロールの参照
     */
    function buildSettingsDialog(session) {
        var controls = {};
        var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        settingsDialog.orientation = "column";
        settingsDialog.alignChildren = ["fill", "top"];
        settingsDialog.margins = DIALOG_MARGINS;
        controls.dialog = settingsDialog;

        buildPresetRow(settingsDialog, controls);
        buildSplitDirectionPanel(settingsDialog, controls);
        buildBalancePanel(settingsDialog, controls);

        /* 2カラム：左＝塗り、右＝線 / Two columns: fill on the left, stroke on the right */
        var columnsGroup = settingsDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignment = ["fill", "top"];
        columnsGroup.alignChildren = ["fill", "top"];
        buildFillPanel(addColumn(columnsGroup, "fill"), controls, session);
        buildStrokePanel(addColumn(columnsGroup, "fill"), controls, session);

        buildCornerPanel(settingsDialog, controls, session);

        controls.groupItemsCheckbox = addCheckbox(settingsDialog, "checkbox.groupItems", "tooltip.groupItems", DEFAULT_GROUP_ITEMS);
        controls.groupItemsCheckbox.alignment = ["center", "center"];

        /* ボタン行（左右中央） / Button row (centered) */
        var buttonRow = addButtonRow(settingsDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        controls.btnCancel = btnCancel;
        controls.btnOK = btnOK;
        settingsDialog.defaultElement = controls.btnOK;
        settingsDialog.cancelElement = controls.btnCancel;

        updateBalanceRange(controls);
        updateStrokeEnabled(controls);
        updateCornerState(controls);
        return controls;
    }

    /**
     * ［分割方法］パネルを組み立てる（縦長なら上下、それ以外は左右を選んでおく）
     * @param {Window} settingsDialog - 追加先のダイアログ
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @returns {void}
     */
    function buildSplitDirectionPanel(settingsDialog, controls) {
        var splitDirectionPanel = addPanel(settingsDialog, getLabel("panel.splitDirection"));
        splitDirectionPanel.orientation = "row";
        splitDirectionPanel.alignChildren = ["center", "center"];
        controls.splitLRRadio = addRadio(splitDirectionPanel, "radio.splitLR", "tooltip.splitLR");
        controls.splitTBRadio = addRadio(splitDirectionPanel, "radio.splitTB", "tooltip.splitTB");

        var vertical = (targetBounds[1] - targetBounds[3]) > (targetBounds[2] - targetBounds[0]);
        controls.splitTBRadio.value = vertical;
        controls.splitLRRadio.value = !vertical;
    }

    /**
     * バランスの1行（ラベル・幅・％・正方形）を追加する
     * @param {Group} parent - 追加先
     * @param {string} reservedLabel - ラベルの幅を確保する文言（表示は updateBalanceLabels() で入れる）
     * @returns {{label: StaticText, widthInput: EditText, percentInput: EditText, squareCheckbox: Checkbox}} 行のコントロール
     */
    function addBalanceRow(parent, reservedLabel) {
        var rowGroup = addRow(parent);
        var sideLabel = rowGroup.add("statictext", undefined, reservedLabel);
        var widthInput = addSteppedInput(rowGroup, "0", { min: 0 });
        widthInput.helpTip = getLabel("tooltip.width");
        widthInput.characters = WIDTH_INPUT_CHARS;
        rowGroup.add("statictext", undefined, getUnitInfo("rulerType").label);
        var percentInput = addSteppedInput(rowGroup, "50", { min: 0, max: 100 });
        percentInput.helpTip = getLabel("tooltip.percent");
        percentInput.characters = PERCENT_INPUT_CHARS;
        rowGroup.add("statictext", undefined, "%");
        var squareCheckbox = addCheckbox(rowGroup, "checkbox.square", "tooltip.square", false);
        return { label: sideLabel, widthInput: widthInput, percentInput: percentInput, squareCheckbox: squareCheckbox };
    }

    /**
     * ［バランス］パネル（左右の幅・％とスライダー）を組み立てる
     * @param {Window} settingsDialog - 追加先のダイアログ
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @returns {void}
     */
    function buildBalancePanel(settingsDialog, controls) {
        var balancePanel = addPanel(settingsDialog, getLabel("panel.balance"), "fill");
        var balanceGroup = addColumn(balancePanel, "fill");
        balanceGroup.spacing = BALANCE_SPACING;

        var reservedLabel = getLongestSideLabel();
        controls.firstBalance = addBalanceRow(balanceGroup, reservedLabel);
        controls.secondBalance = addBalanceRow(balanceGroup, reservedLabel);

        var sliderRowGroup = addRow(balanceGroup);
        sliderRowGroup.alignChildren = ["fill", "center"];
        sliderRowGroup.margins = BALANCE_SLIDER_MARGINS;
        /* 値は中央からのずれ（定規の単位）。範囲は updateBalanceRange() で決める / the value is the offset from the center in ruler units; updateBalanceRange() sets the range */
        controls.balanceSlider = sliderRowGroup.add("slider", undefined, 0, 0, 0);
        controls.balanceSlider.helpTip = getLabel("tooltip.widthSlider");
        controls.balanceSlider.preferredSize = BALANCE_SLIDER_SIZE;
        controls.balanceMax = 0;
    }

    /**
     * 塗りの1行（チェックボックスと色見本）を追加する
     * @param {Group|Panel} parent - 追加先
     * @param {string} reservedLabel - チェックボックスの幅を確保する文言（表示は updateBalanceLabels() で入れる）
     * @param {boolean} checked - 初期状態
     * @param {string} colorText - 初期の色
     * @returns {{checkbox: Checkbox, swatch: Object}} 行のコントロール
     */
    function addFillRow(parent, reservedLabel, checked, colorText) {
        var rowGroup = addRow(parent);
        var fillCheckbox = rowGroup.add("checkbox", undefined, reservedLabel);
        fillCheckbox.helpTip = getLabel("tooltip.fillSide");
        fillCheckbox.value = !!checked;
        return { checkbox: fillCheckbox, swatch: addColorSwatch(rowGroup, colorText) };
    }

    /**
     * ［塗り］パネルを組み立てる
     * @param {Group} parent - 追加先
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @param {Object} session - セッションに残した前回の値
     * @returns {void}
     */
    function buildFillPanel(parent, controls, session) {
        var fillPanel = addPanel(parent, getLabel("panel.fill"));
        var reservedLabel = getLongestSideLabel();
        var firstFillRow = addFillRow(fillPanel, reservedLabel, session.fillFirst, session.firstColor);
        var secondFillRow = addFillRow(fillPanel, reservedLabel, session.fillSecond, session.secondColor);
        controls.fillFirstCheckbox = firstFillRow.checkbox;
        controls.firstColorSwatch = firstFillRow.swatch;
        controls.fillSecondCheckbox = secondFillRow.checkbox;
        controls.secondColorSwatch = secondFillRow.swatch;
    }

    /**
     * ［線］パネル（外枠・区切り線・線幅・色）を組み立てる
     * @param {Group} parent - 追加先
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @param {Object} session - セッションに残した前回の値
     * @returns {void}
     */
    function buildStrokePanel(parent, controls, session) {
        var strokePanel = addPanel(parent, getLabel("panel.stroke"));
        controls.overallFrameCheckbox = addCheckbox(strokePanel, "checkbox.overallFrame", "tooltip.overallFrame", session.overallFrame);
        controls.dividerCheckbox = addCheckbox(strokePanel, "checkbox.divider", "tooltip.divider", session.divider);

        controls.strokeRowGroup = addRow(strokePanel);
        controls.strokeRowGroup.add("statictext", undefined, labelText("fieldLabel.strokeWidth"));
        controls.strokeInput = addSteppedInput(controls.strokeRowGroup,
            formatUnitValue(ptToUnit(session.strokeWidthPt, "strokeUnits"), "strokeUnits"), { min: 0 });
        controls.strokeInput.helpTip = getLabel("tooltip.strokeWidth");
        controls.strokeInput.characters = STROKE_INPUT_CHARS;
        controls.strokeRowGroup.add("statictext", undefined, getUnitInfo("strokeUnits").label);

        controls.strokeColorRowGroup = addRow(strokePanel);
        controls.strokeColorRowGroup.add("statictext", undefined, labelText("fieldLabel.color"));
        controls.strokeColorSwatch = addColorSwatch(controls.strokeColorRowGroup, session.strokeColor);
    }

    /**
     * 角丸の1行（チェックボックスと半径）を追加する
     * @param {Group} parent - 追加先
     * @param {string} labelKey - チェックボックスの LABELS キー
     * @param {boolean} checked - 初期状態
     * @param {number} radiusPt - 半径の初期値（pt）
     * @returns {{checkbox: Checkbox, input: EditText}} 行のコントロール
     */
    function addCornerRow(parent, labelKey, checked, radiusPt) {
        var rowGroup = addRow(parent);
        var cornerCheckbox = addCheckbox(rowGroup, labelKey, "tooltip.corner", checked);
        var radiusInput = addSteppedInput(rowGroup, formatUnitValue(ptToUnit(radiusPt, "rulerType"), "rulerType"), { min: 0 });
        radiusInput.helpTip = getLabel("tooltip.corner");
        radiusInput.characters = CORNER_INPUT_CHARS;
        return { checkbox: cornerCheckbox, input: radiusInput };
    }

    /**
     * ［角丸］パネル（ピル形状・連動・4つの角）を組み立てる
     * @param {Window} settingsDialog - 追加先のダイアログ
     * @param {Object} controls - コントロールの参照（ここで作ったものを足す）
     * @param {Object} session - セッションに残した前回の値
     * @returns {void}
     */
    function buildCornerPanel(settingsDialog, controls, session) {
        var cornerPanel = addPanel(settingsDialog, getLabel("panel.cornerRadius") + " (" + getUnitInfo("rulerType").label + ")");
        controls.pillCheckbox = addCheckbox(cornerPanel, "checkbox.pillShape", "tooltip.pillShape", session.pillShape);

        /* 左の列に左上・左下、右の列に右上・右下。2列の間に連動アイコンを天地中央で置く
           top-left and bottom-left on the left, top-right and bottom-right on the right; the link icon sits centered between them */
        var cornerColumnsGroup = cornerPanel.add("group");
        cornerColumnsGroup.orientation = "row";
        cornerColumnsGroup.alignment = ["fill", "top"];
        cornerColumnsGroup.alignChildren = ["fill", "top"];
        var leftColumn = addColumn(cornerColumnsGroup, "left");
        /* 切り替えたあとの処理は bindDialogEvents() で入れる / The toggle handler is set in bindDialogEvents() */
        controls.cornerLinkToggle = addLinkToggle(cornerColumnsGroup, session.cornerLink, function () {
            if (controls.onCornerLinkToggle) controls.onCornerLinkToggle();
        });
        controls.cornerLinkToggle.alignment = ["center", "center"];
        controls.cornerLinkToggle.helpTip = getLabel("tooltip.cornerLink");
        var rightColumn = addColumn(cornerColumnsGroup, "left");

        var cornerEnabled = session.cornerEnabled;
        var cornerRadiusPt = session.cornerRadiusPt;
        controls.corners = {};
        controls.corners.tl = addCornerRow(leftColumn, "checkbox.cornerTL", cornerEnabled.tl, cornerRadiusPt.tl);
        controls.corners.bl = addCornerRow(leftColumn, "checkbox.cornerBL", cornerEnabled.bl, cornerRadiusPt.bl);
        controls.corners.tr = addCornerRow(rightColumn, "checkbox.cornerTR", cornerEnabled.tr, cornerRadiusPt.tr);
        controls.corners.br = addCornerRow(rightColumn, "checkbox.cornerBR", cornerEnabled.br, cornerRadiusPt.br);
    }

    /**
     * 分割方法に合わせて、バランスと塗りのラベルを左・右か上・下にする
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function updateBalanceLabels(controls) {
        var vertical = controls.splitTBRadio.value;
        var firstText = labelText(vertical ? "fieldLabel.top" : "fieldLabel.left");
        var secondText = labelText(vertical ? "fieldLabel.bottom" : "fieldLabel.right");
        controls.firstBalance.label.text = firstText;
        controls.secondBalance.label.text = secondText;
        controls.fillFirstCheckbox.text = firstText;
        controls.fillSecondCheckbox.text = secondText;
    }

    /**
     * スライダーの範囲を分割する向きの長さの半分にし、中央に戻す
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function updateBalanceRange(controls) {
        var splitLengthPt = controls.splitTBRadio.value
            ? (targetBounds[1] - targetBounds[3])
            : (targetBounds[2] - targetBounds[0]);
        controls.balanceMax = Math.max(roundUnitValue(ptToUnit(splitLengthPt / 2, "rulerType"), "rulerType"), 0);
        controls.balanceSlider.minvalue = -controls.balanceMax;
        controls.balanceSlider.maxvalue = controls.balanceMax;
        syncBalanceFromOffset(controls, 0, null);
    }

    /**
     * 中央からのずれをスライダーの範囲に収めて丸める
     * @param {Object} controls - コントロールの参照
     * @param {number|string} offsetValue - 中央からのずれ（定規の単位）
     * @returns {number} 収めたずれ
     */
    function clampOffset(controls, offsetValue) {
        var offsetUnit = Number(offsetValue);
        if (isNaN(offsetUnit)) return 0;
        return roundUnitValue(clampNumber(offsetUnit, -controls.balanceMax, controls.balanceMax), "rulerType");
    }

    /**
     * 入力欄に文字を入れる（入力中の欄は書き換えない）
     * @param {EditText} input - 入力欄
     * @param {string} text - 入れる文字
     * @param {EditText|null} skipInput - 書き換えない入力欄
     * @returns {void}
     */
    function setInputText(input, text, skipInput) {
        if (input !== skipInput) input.text = text;
    }

    /**
     * 中央からのずれを、スライダー・幅・％のすべてに反映する
     * @param {Object} controls - コントロールの参照
     * @param {number|string} offsetValue - 中央からのずれ（定規の単位。正なら左・上が広がる）
     * @param {EditText|null} skipInput - 書き換えない入力欄（入力中の欄。小数点を打てるように）
     * @returns {void}
     */
    function syncBalanceFromOffset(controls, offsetValue, skipInput) {
        var offsetUnit = clampOffset(controls, offsetValue);
        var maxUnit = controls.balanceMax;
        controls.balanceSlider.value = offsetUnit;
        setInputText(controls.firstBalance.widthInput, String(roundUnitValue(maxUnit + offsetUnit, "rulerType")), skipInput);
        setInputText(controls.secondBalance.widthInput, String(roundUnitValue(maxUnit - offsetUnit, "rulerType")), skipInput);
        var firstPercent = (maxUnit > 0) ? Math.round((maxUnit + offsetUnit) / (2 * maxUnit) * 100) : 50;
        var secondPercent = (maxUnit > 0) ? Math.round((maxUnit - offsetUnit) / (2 * maxUnit) * 100) : 50;
        setInputText(controls.firstBalance.percentInput, String(firstPercent), skipInput);
        setInputText(controls.secondBalance.percentInput, String(secondPercent), skipInput);
    }

    /**
     * 幅か％の入力欄の値から、中央からのずれを求めて反映する
     * @param {Object} controls - コントロールの参照
     * @param {EditText} input - 値を読む入力欄
     * @param {boolean} whileTyping - 入力中か（true ならその欄は書き換えない）
     * @returns {void}
     */
    function syncBalanceFromInput(controls, input, whileTyping) {
        var value = Number(input.text);
        var maxUnit = controls.balanceMax;
        var offsetUnit;
        if (input === controls.firstBalance.widthInput) {
            offsetUnit = (isNaN(value) ? 0 : value) - maxUnit;
        } else if (input === controls.secondBalance.widthInput) {
            offsetUnit = maxUnit - (isNaN(value) ? 0 : value);
        } else {
            var percent = isNaN(value) ? 50 : clampNumber(value, 0, 100);
            var firstRatio = (input === controls.firstBalance.percentInput) ? percent / 100 : 1 - percent / 100;
            offsetUnit = (firstRatio - 0.5) * 2 * maxUnit;
        }
        syncBalanceFromOffset(controls, offsetUnit, whileTyping ? input : null);
    }

    /**
     * 指定した側が正方形になる位置で分ける
     * @param {Object} controls - コントロールの参照
     * @param {string} side - "first"（左・上）| "second"（右・下）
     * @returns {void}
     */
    function applySquareOffset(controls, side) {
        var widthPt = targetBounds[2] - targetBounds[0];
        var heightPt = targetBounds[1] - targetBounds[3];
        var vertical = controls.splitTBRadio.value;
        var shortSidePt = vertical ? widthPt : heightPt;
        var halfLongSidePt = (vertical ? heightPt : widthPt) / 2;
        var offsetPt = halfLongSidePt - shortSidePt;
        if (side === "first") offsetPt = -offsetPt;
        syncBalanceFromOffset(controls, ptToUnit(offsetPt, "rulerType"), null);
    }

    /**
     * 線幅と線の色の欄を使うか（外枠か区切り線があるとき）
     * @param {Object} controls - コントロールの参照
     * @returns {boolean} 使うか
     */
    function needsStroke(controls) {
        return controls.overallFrameCheckbox.value || controls.dividerCheckbox.value;
    }

    /**
     * 使わない線の欄を無効にする
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function updateStrokeEnabled(controls) {
        var strokeEnabled = needsStroke(controls);
        controls.strokeRowGroup.enabled = strokeEnabled;
        controls.strokeColorRowGroup.enabled = strokeEnabled;
        redrawSteppersIn(controls.strokeRowGroup); /* ∧∨のディム表示を描き直す / redraw the stepper's dimming */
    }

    /**
     * 線幅を pt で読む
     * @param {Object} controls - コントロールの参照
     * @returns {number|null} 線幅（pt）。0以下や数値でなければ null
     */
    function readStrokeWidthPt(controls) {
        var strokeWidthUnit = Number(controls.strokeInput.text);
        if (isNaN(strokeWidthUnit) || strokeWidthUnit <= 0) return null;
        return unitToPt(strokeWidthUnit, "strokeUnits");
    }

    /**
     * 角丸の半径を pt で読む
     * @param {EditText} radiusInput - 半径の入力欄
     * @returns {number} 半径（pt）。負の値や数値でなければ 0
     */
    function readCornerRadiusPt(radiusInput) {
        var radiusUnit = Number(radiusInput.text);
        if (isNaN(radiusUnit) || radiusUnit < 0) return 0;
        return unitToPt(radiusUnit, "rulerType");
    }

    /**
     * 角丸の欄の有効・無効を切り替える（ピル形状なら全部、連動なら左上以外を無効に）
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function updateCornerState(controls) {
        var corners = controls.corners;
        var perCornerEnabled = !controls.pillCheckbox.value;
        var linked = controls.cornerLinkToggle.value;

        setLinkToggleEnabled(controls.cornerLinkToggle, perCornerEnabled);
        corners.tl.checkbox.enabled = perCornerEnabled;
        setSteppedInputEnabled(corners.tl.input, perCornerEnabled && corners.tl.checkbox.value);

        var followerKeys = ["bl", "tr", "br"];
        for (var i = 0; i < followerKeys.length; i++) {
            var corner = corners[followerKeys[i]];
            corner.checkbox.enabled = perCornerEnabled && !linked;
            setSteppedInputEnabled(corner.input, perCornerEnabled && !linked && corner.checkbox.value);
        }
        syncLinkedCornerValues(controls);
    }

    /**
     * 連動しているとき、左上の半径をほかの角に写す
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function syncLinkedCornerValues(controls) {
        if (!controls.cornerLinkToggle.value) return;
        var corners = controls.corners;
        corners.bl.input.text = corners.tl.input.text;
        corners.tr.input.text = corners.tl.input.text;
        corners.br.input.text = corners.tl.input.text;
    }

    /**
     * 連動しているとき、左上のチェックをほかの角に写す
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function syncLinkedCornerChecks(controls) {
        if (!controls.cornerLinkToggle.value) return;
        var corners = controls.corners;
        corners.bl.checkbox.value = corners.tl.checkbox.value;
        corners.tr.checkbox.value = corners.tl.checkbox.value;
        corners.br.checkbox.value = corners.tl.checkbox.value;
    }

    /**
     * 角ごとの角丸の設定をまとめる
     * @param {Object} controls - コントロールの参照
     * @returns {Object|null} 角（tl・bl・tr・br）ごとの { enabled, radius }。ピル形状か、どの角も使わなければ null
     */
    function buildPerCornerConfig(controls) {
        if (controls.pillCheckbox.value) return null;
        var corners = controls.corners;
        var linked = controls.cornerLinkToggle.value;
        var anyEnabled = corners.tl.checkbox.value ||
            (!linked && (corners.bl.checkbox.value || corners.tr.checkbox.value || corners.br.checkbox.value));
        if (!anyEnabled) return null;

        var perCorner = {};
        for (var i = 0; i < CORNER_KEYS.length; i++) {
            /* 連動中は左上のチェックと半径を4つの角に使う / when linked, all four corners follow the top-left one */
            var sourceCorner = linked ? corners.tl : corners[CORNER_KEYS[i]];
            perCorner[CORNER_KEYS[i]] = {
                enabled: !!sourceCorner.checkbox.value,
                radius: readCornerRadiusPt(sourceCorner.input)
            };
        }
        return perCorner;
    }

    /**
     * ダイアログの値を描画の設定にまとめる
     * @param {Object} controls - コントロールの参照
     * @returns {Object|null} 描画の設定。線を使うのに線幅が不正なら null
     */
    function collectDrawOptions(controls) {
        var strokeWidthPt = readStrokeWidthPt(controls);
        if (strokeWidthPt === null) {
            /* 線を使わないなら、不正な値でも既定値で進める / without a frame or divider, fall back to the default */
            if (needsStroke(controls)) return null;
            strokeWidthPt = DEFAULT_STROKE_PT;
        }

        return {
            vertical: controls.splitTBRadio.value,
            offsetPt: unitToPt(clampOffset(controls, controls.balanceSlider.value), "rulerType"),
            fillFirst: controls.fillFirstCheckbox.value,
            fillSecond: controls.fillSecondCheckbox.value,
            firstColor: parseColorString(controls.firstColorSwatch.getColor()),
            secondColor: parseColorString(controls.secondColorSwatch.getColor()),
            overallFrame: controls.overallFrameCheckbox.value,
            divider: controls.dividerCheckbox.value,
            strokeWidthPt: strokeWidthPt,
            strokeColor: parseColorString(controls.strokeColorSwatch.getColor()),
            perCorner: buildPerCornerConfig(controls),
            pillShape: controls.pillCheckbox.value,
            groupItems: controls.groupItemsCheckbox.value
        };
    }

    /**
     * ダイアログの値でプレビューを描き直す（線幅が不正なら消すだけ）
     * @param {Object} controls - コントロールの参照
     * @returns {void}
     */
    function refreshPreview(controls) {
        drawPreview(collectDrawOptions(controls));
    }

    /**
     * 分割方法を切り替え、ラベルとバランスを合わせてプレビューを描き直す
     * @param {Object} controls - コントロールの参照
     * @param {boolean} vertical - 上下に分けるか
     * @returns {void}
     */
    function setSplitDirection(controls, vertical) {
        controls.splitTBRadio.value = vertical;
        controls.splitLRRadio.value = !vertical;
        updateBalanceLabels(controls);
        updateBalanceRange(controls);
        refreshPreview(controls);
    }

    /**
     * バランスの幅・％の入力欄にイベントを付ける
     * @param {Object} controls - コントロールの参照
     * @param {EditText} input - 入力欄
     * @returns {void}
     */
    function bindBalanceInput(controls, input) {
        /* ∧∨・↑↓キーで増減したあと / after stepping with the stepper or arrow keys */
        input.onStepped = function () {
            syncBalanceFromInput(controls, input, false);
            refreshPreview(controls);
        };
        input.addEventListener("changing", function () {
            syncBalanceFromInput(controls, input, true);
            refreshPreview(controls);
        });
        /* 確定したら、入力した欄も範囲内の値にそろえる / on commit, normalize the edited field too */
        input.onChange = function () { syncBalanceFromInput(controls, input, false); };
    }

    /**
     * ［正方形］のクリックを処理する（もう一方の［正方形］は外す）
     * @param {Object} controls - コントロールの参照
     * @param {string} side - "first"（左・上）| "second"（右・下）
     * @returns {void}
     */
    function onSquareClick(controls, side) {
        var clickedCheckbox = (side === "first") ? controls.firstBalance.squareCheckbox : controls.secondBalance.squareCheckbox;
        var otherCheckbox = (side === "first") ? controls.secondBalance.squareCheckbox : controls.firstBalance.squareCheckbox;
        if (clickedCheckbox.value) {
            otherCheckbox.value = false;
            applySquareOffset(controls, side);
        }
        refreshPreview(controls);
    }

    /**
     * ダイアログのイベントを付ける
     * @param {Object} controls - コントロールの参照
     * @param {Object} session - セッション設定（ダイアログの位置を残す）
     * @returns {void}
     */
    function bindDialogEvents(controls, session) {
        var corners = controls.corners;
        var refresh = function () { refreshPreview(controls); };
        var onStrokeOptionClick = function () {
            updateStrokeEnabled(controls);
            refresh();
        };
        var onCornerOptionClick = function () {
            updateCornerState(controls);
            refresh();
        };
        var onTopLeftRadiusChange = function () {
            syncLinkedCornerValues(controls);
            refresh();
        };

        /* プリセット（行は隠している）/ Presets (the row is hidden) */
        controls.presetDropdown.onChange = function () {
            var selectedItem = controls.presetDropdown.selection;
            if (!selectedItem || selectedItem.index === 0) return;
            var presetData = getPresetStore()[selectedItem.text];
            if (presetData) applyPresetData(controls, presetData);
        };
        controls.btnPresetSave.onClick = function () {
            var presetName = showPresetNameDialog();
            if (presetName === null) return;
            savePreset(presetName, collectPresetData(controls));
            fillPresetDropdown(controls);
        };
        controls.btnPresetExport.onClick = function () { exportSettingsToDesktop(controls); };

        /* 分割方法 / Split direction */
        controls.splitLRRadio.onClick = function () { setSplitDirection(controls, false); };
        controls.splitTBRadio.onClick = function () { setSplitDirection(controls, true); };

        /* バランス / Balance */
        bindBalanceInput(controls, controls.firstBalance.widthInput);
        bindBalanceInput(controls, controls.secondBalance.widthInput);
        bindBalanceInput(controls, controls.firstBalance.percentInput);
        bindBalanceInput(controls, controls.secondBalance.percentInput);
        controls.balanceSlider.onChanging = function () {
            controls.firstBalance.squareCheckbox.value = false;
            controls.secondBalance.squareCheckbox.value = false;
            var offsetUnit = Number(controls.balanceSlider.value);
            /* option を押しながらドラッグすると整数に丸める / hold Option while dragging to snap to whole units */
            if (ScriptUI.environment.keyboardState.altKey) offsetUnit = Math.round(offsetUnit);
            syncBalanceFromOffset(controls, offsetUnit, null);
            refresh();
        };
        /* shift+↑↓ でスライダーを範囲の10%ずつ動かす / Shift+Up/Down moves the slider by 10% of its range */
        controls.balanceSlider.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            if (!ScriptUI.environment.keyboardState.shiftKey) return;
            var maxUnit = controls.balanceSlider.maxvalue;
            if (maxUnit <= 0) return;
            var delta = Math.max(Math.round(maxUnit * 0.1), 1);
            var offsetUnit = controls.balanceSlider.value + ((event.keyName === "Up") ? delta : -delta);
            controls.balanceSlider.value = Math.round(clampNumber(offsetUnit, controls.balanceSlider.minvalue, maxUnit));
            controls.balanceSlider.notify("onChanging");
            event.preventDefault();
        });
        controls.firstBalance.squareCheckbox.onClick = function () { onSquareClick(controls, "first"); };
        controls.secondBalance.squareCheckbox.onClick = function () { onSquareClick(controls, "second"); };

        /* 塗り / Fill */
        controls.fillFirstCheckbox.onClick = refresh;
        controls.fillSecondCheckbox.onClick = refresh;
        controls.firstColorSwatch.onChange = refresh;
        controls.secondColorSwatch.onChange = refresh;

        /* 線 / Stroke */
        controls.overallFrameCheckbox.onClick = onStrokeOptionClick;
        controls.dividerCheckbox.onClick = onStrokeOptionClick;
        controls.strokeInput.onStepped = refresh;
        controls.strokeInput.addEventListener("changing", refresh);
        controls.strokeColorSwatch.onChange = refresh;

        /* 角丸 / Corner radius */
        controls.pillCheckbox.onClick = onCornerOptionClick;
        controls.onCornerLinkToggle = function () {
            /* 連動を入れたら左上を有効にして、ほかの角をそろえる / turning Link on enables the top-left corner and copies it to the others */
            if (controls.cornerLinkToggle.value) corners.tl.checkbox.value = true;
            syncLinkedCornerChecks(controls);
            onCornerOptionClick();
        };
        corners.tl.checkbox.onClick = function () {
            syncLinkedCornerChecks(controls);
            onCornerOptionClick();
        };
        corners.bl.checkbox.onClick = onCornerOptionClick;
        corners.tr.checkbox.onClick = onCornerOptionClick;
        corners.br.checkbox.onClick = onCornerOptionClick;
        corners.tl.input.onStepped = onTopLeftRadiusChange;
        corners.tl.input.addEventListener("changing", onTopLeftRadiusChange);
        corners.bl.input.onStepped = refresh;
        corners.tr.input.onStepped = refresh;
        corners.br.input.onStepped = refresh;
        corners.bl.input.addEventListener("changing", refresh);
        corners.tr.input.addEventListener("changing", refresh);
        corners.br.input.addEventListener("changing", refresh);

        /* H・V で分割方法、F で外枠、D で区切り線を切り替える / H and V pick the split direction, F toggles the frame, D the divider */
        addKeyShortcuts(controls.dialog, {
            "H": controls.splitLRRadio,
            "V": controls.splitTBRadio,
            "F": controls.overallFrameCheckbox,
            "D": controls.dividerCheckbox
        }, {
            numericFields: [
                controls.firstBalance.widthInput, controls.firstBalance.percentInput,
                controls.secondBalance.widthInput, controls.secondBalance.percentInput,
                controls.strokeInput,
                controls.corners.tl.input, controls.corners.bl.input, controls.corners.tr.input, controls.corners.br.input
            ]
        });

        controls.btnOK.onClick = function () {
            if (!collectDrawOptions(controls)) {
                /* 線を使うのに線幅が不正なら、その欄に戻す / return to the stroke field when a needed stroke width is invalid */
                controls.strokeInput.active = true;
                return;
            }
            controls.dialog.close(1);
        };
        controls.btnCancel.onClick = function () { controls.dialog.close(0); };

        controls.dialog.onShow = function () {
            /* プレビューの間は元のオブジェクトを隠す / hide the original object while previewing */
            targetItem.hidden = true;
            /* レイアウトが済んでから本来の文言を入れる（幅は長いほうで確保済み）/ set the real labels after layout; the width is already reserved */
            updateBalanceLabels(controls);
            refresh();
        };
    }

    /**
     * ダイアログの値をセッションに残して保存する（不正な値の欄は前回の値のまま）
     * @param {Object} controls - コントロールの参照
     * @param {Object} session - セッション設定
     * @returns {void}
     */
    function saveDialogToSession(controls, session) {
        session.fillFirst = controls.fillFirstCheckbox.value;
        session.fillSecond = controls.fillSecondCheckbox.value;
        session.firstColor = controls.firstColorSwatch.getColor();
        session.secondColor = controls.secondColorSwatch.getColor();
        session.overallFrame = controls.overallFrameCheckbox.value;
        session.divider = controls.dividerCheckbox.value;
        session.strokeColor = controls.strokeColorSwatch.getColor();
        session.pillShape = controls.pillCheckbox.value;
        session.cornerLink = controls.cornerLinkToggle.value;

        var strokeWidthPt = readStrokeWidthPt(controls);
        if (strokeWidthPt !== null) session.strokeWidthPt = strokeWidthPt;

        for (var i = 0; i < CORNER_KEYS.length; i++) {
            var cornerKey = CORNER_KEYS[i];
            var corner = controls.corners[cornerKey];
            var radiusUnit = Number(corner.input.text);
            session.cornerEnabled[cornerKey] = corner.checkbox.value;
            if (!isNaN(radiusUnit) && radiusUnit >= 0) session.cornerRadiusPt[cornerKey] = unitToPt(radiusUnit, "rulerType");
        }
        settingsStore.save(session);
    }

    /**
     * 設定ダイアログを表示する。閉じたらプレビューを片付け、元のオブジェクトを表示に戻す
     * @returns {Object|null} 描画の設定。キャンセルなら null
     */
    function showSettingsDialog() {
        var session = getSessionSettings();
        var controls = buildSettingsDialog(session);
        bindDialogEvents(controls, session);

        /* 中断した実行で残ったプレビューレイヤーを消してから始める / start without a preview layer left over from an interrupted run */
        removePreviewLayer();
        prepareDialogWindow(controls.dialog, SCRIPT_NAME);
        var dialogResult = controls.dialog.show();
        removePreviewLayer();
        targetItem.hidden = false;
        saveDialogToSession(controls, session);

        return (dialogResult === 1) ? collectDrawOptions(controls) : null;
    }

    // =========================================
    // 角丸アルゴリズム / Round Any Corner
    // 選択したアンカーだけを丸める / rounds only the selected anchors
    // Based on: Hiroyuki Sato (MIT) https://github.com/shspage
    // =========================================

    /**
     * 選択したアンカーの角を半径 rr で丸める
     * @param {PathItem[]} targetPaths - 対象のパス
     * @param {Object} roundOptions - 設定（rr: 半径）
     * @returns {void}
     */
    function roundAnyCorner(targetPaths, roundOptions) {
        var rr = roundOptions.rr;

        var p, op, pnts;
        var skipList, adjRdirAtEnd, redrawFlg;
        var i, nxi, pvi, q, d, ds, r, g, t, qb;
        var anc1, ldir1, rdir1, anc2, ldir2, rdir2;

        var hanLen = 4 * (Math.sqrt(2) - 1) / 3;
        var ptyp = PointType.SMOOTH;

        for (var j = 0; j < targetPaths.length; j++) {
            p = targetPaths[j].pathPoints;
            if (readjustAnchors(p) < 2) continue;
            op = !targetPaths[j].closed;
            pnts = op ? [getDat(p[0])] : [];
            redrawFlg = false;
            adjRdirAtEnd = 0;

            skipList = [(op || !isSelected(p[0]) || !isCorner(p, 0))];
            for (i = 1; i < p.length; i++) {
                skipList.push((!isSelected(p[i])
                    || !isCorner(p, i)
                    || (op && i == p.length - 1)));
            }

            for (i = 0; i < p.length; i++) {
                nxi = parseIdx(p, i + 1);
                if (nxi < 0) break;

                pvi = parseIdx(p, i - 1);

                q = [p[i].anchor, p[i].rightDirection,
                p[nxi].leftDirection, p[nxi].anchor];

                ds = dist(q[0], q[3]) / 2;
                if (arrEq(q[0], q[1]) && arrEq(q[2], q[3])) {
                    r = Math.min(ds, rr);
                    g = getRad(q[0], q[3]);
                    anc1 = getPnt(q[0], g, r);
                    ldir1 = getPnt(anc1, g + Math.PI, r * hanLen);

                    if (skipList[nxi]) {
                        if (!skipList[i]) {
                            pnts.push([anc1, anc1, ldir1, ptyp]);
                            redrawFlg = true;
                        }
                        pnts.push(getDat(p[nxi]));
                    } else {
                        if (r < rr) {
                            pnts.push([anc1,
                                getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                ldir1,
                                ptyp]);
                        } else {
                            if (!skipList[i]) pnts.push([anc1, anc1, ldir1, ptyp]);
                            anc2 = getPnt(q[3], g + Math.PI, r);
                            pnts.push([anc2,
                                getPnt(anc2, g, r * hanLen),
                                anc2,
                                ptyp]);
                        }
                        redrawFlg = true;
                    }
                } else {
                    d = getT4Len(q, 0) / 2;
                    r = Math.min(d, rr);
                    t = getT4Len(q, r);
                    anc1 = bezier(q, t);
                    rdir1 = defHan(t, q, 1);
                    ldir1 = getPnt(anc1, getRad(rdir1, anc1), r * hanLen);

                    if (skipList[nxi]) {
                        if (skipList[i]) {
                            pnts.push(getDat(p[nxi]));
                        } else {
                            pnts.push([anc1, rdir1, ldir1, ptyp]);
                            with (p[nxi]) pnts.push([anchor,
                                rightDirection,
                                adjHan(anchor, leftDirection, 1 - t),
                                ptyp]);
                            redrawFlg = true;
                        }
                    } else {
                        if (r < rr) {
                            if (skipList[i]) {
                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    defHan(t, q, 0),
                                    ptyp]);
                            } else {
                                pnts.push([anc1,
                                    getPnt(anc1, getRad(ldir1, anc1), r * hanLen),
                                    ldir1,
                                    ptyp]);
                            }
                        } else {
                            if (skipList[i]) {
                                t = getT4Len(q, -r);
                                anc2 = bezier(q, t);

                                if (!op && i == 0) {
                                    adjRdirAtEnd = t;
                                } else {
                                    pnts[pnts.length - 1][1] = adjHan(q[0], q[1], t);
                                }

                                ldir2 = defHan(t, q, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);

                                pnts.push([anc2, rdir2, ldir2, ptyp]);
                            } else {
                                qb = [anc1, rdir1, adjHan(q[3], q[2], 1 - t), q[3]];
                                t = getT4Len(qb, -r);
                                anc2 = bezier(qb, t);
                                ldir2 = defHan(t, qb, 0);
                                rdir2 = getPnt(anc2, getRad(ldir2, anc2), r * hanLen);
                                rdir1 = adjHan(anc1, rdir1, t);

                                pnts.push([anc1, rdir1, ldir1, ptyp],
                                    [anc2, rdir2, ldir2, ptyp]);
                            }
                        }
                        redrawFlg = true;
                    }
                }
            }
            if (adjRdirAtEnd > 0) {
                pnts[pnts.length - 1][1] = adjHan(p[0].anchor, p[0].rightDirection, adjRdirAtEnd);
            }

            if (redrawFlg) {
                for (i = p.length - 1; i > 0; i--) p[i].remove();

                for (i = 0; i < pnts.length; i++) {
                    var pt = i > 0 ? p.add() : p[0];
                    with (pt) {
                        anchor = pnts[i][0];
                        rightDirection = pnts[i][1];
                        leftDirection = pnts[i][2];
                        pointType = pnts[i][3];
                    }
                }
            }
        }
        app.activeDocument.selection = targetPaths;
    }

    /**
     * 点から角度 rad の方向へ len 進んだ点を返す
     * @param {number[]} pt - 起点
     * @param {number} rad - 角度（ラジアン）
     * @param {number} len - 距離
     * @returns {number[]} 点
     */
    function getPnt(pt, rad, len) {
        return [pt[0] + Math.cos(rad) * len,
        pt[1] + Math.sin(rad) * len];
    }

    /**
     * ベジェ曲線の t における接線方向のハンドルを返す
     * @param {number} t - 曲線上の位置（0〜1）
     * @param {Array} q - 制御点4つ
     * @param {number} n - 0 なら前側、1 なら後側
     * @returns {number[]} ハンドルの点
     */
    function defHan(t, q, n) {
        return [t * (t * (q[n][0] - 2 * q[n + 1][0] + q[n + 2][0]) + 2 * (q[n + 1][0] - q[n][0])) + q[n][0],
        t * (t * (q[n][1] - 2 * q[n + 1][1] + q[n + 2][1]) + 2 * (q[n + 1][1] - q[n][1])) + q[n][1]];
    }

    /**
     * ベジェ曲線の t における点を返す
     * @param {Array} q - 制御点4つ
     * @param {number} t - 曲線上の位置（0〜1）
     * @returns {number[]} 点
     */
    function bezier(q, t) {
        var u = 1 - t;
        return [u * u * u * q[0][0] + 3 * u * t * (u * q[1][0] + t * q[2][0]) + t * t * t * q[3][0],
        u * u * u * q[0][1] + 3 * u * t * (u * q[1][1] + t * q[2][1]) + t * t * t * q[3][1]];
    }

    /**
     * ハンドルの長さを m 倍にする
     * @param {number[]} anc - アンカー
     * @param {number[]} dir - ハンドル
     * @param {number} m - 倍率
     * @returns {number[]} 新しいハンドルの点
     */
    function adjHan(anc, dir, m) {
        return [anc[0] + (dir[0] - anc[0]) * m,
        anc[1] + (dir[1] - anc[1]) * m];
    }

    /**
     * アンカーが角（折れ点）か
     * @param {PathPoints} p - パスポイント
     * @param {number} idx - 添字
     * @returns {boolean} 角か
     */
    function isCorner(p, idx) {
        var pnt0 = getAnglePnt(p, idx, -1);
        var pnt1 = getAnglePnt(p, idx, 1);
        if (!pnt0 || !pnt1) return false;
        if (pnt0.length < 1 || pnt1.length < 1) return false;
        var rad = getRad2(pnt0, p[idx].anchor, pnt1, true);
        if (rad > Math.PI - 0.1) return false;
        return true;
    }

    /**
     * 角度を測るための隣の点を返す
     * @param {PathPoints} p - パスポイント
     * @param {number} idx1 - 添字
     * @param {number} dir - -1 なら前、1 なら後
     * @returns {number[]|null} 点。無ければ null、重なっていれば空配列
     */
    function getAnglePnt(p, idx1, dir) {
        if (!dir) dir = -1;
        var idx2 = parseIdx(p, idx1 + dir);
        if (idx2 < 0) return null;
        var p2 = p[idx2];
        with (p[idx1]) {
            if (dir < 0) {
                if (arrEq(leftDirection, anchor)) {
                    if (arrEq(p2.anchor, anchor)) return [];
                    if (arrEq(p2.anchor, p2.rightDirection)
                        || arrEq(p2.rightDirection, anchor)) return p2.anchor;
                    else return p2.rightDirection;
                } else {
                    return leftDirection;
                }
            } else {
                if (arrEq(anchor, rightDirection)) {
                    if (arrEq(anchor, p2.anchor)) return [];
                    if (arrEq(p2.anchor, p2.leftDirection)
                        || arrEq(anchor, p2.leftDirection)) return p2.anchor;
                    else return p2.leftDirection;
                } else {
                    return rightDirection;
                }
            }
        }
    }

    /**
     * 2つの配列の要素が等しいか
     * @param {number[]} arr1 - 配列1
     * @param {number[]} arr2 - 配列2
     * @returns {boolean} 等しいか
     */
    function arrEq(arr1, arr2) {
        for (var i = 0; i < arr1.length; i++) {
            if (arr1[i] != arr2[i]) return false;
        }
        return true;
    }

    /**
     * 2点間の距離
     * @param {number[]} p1 - 点1
     * @param {number[]} p2 - 点2
     * @returns {number} 距離
     */
    function dist(p1, p2) {
        return Math.sqrt(Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2));
    }

    /**
     * 2点間の距離の2乗
     * @param {number[]} p1 - 点1
     * @param {number[]} p2 - 点2
     * @returns {number} 距離の2乗
     */
    function dist2(p1, p2) {
        return Math.pow(p1[0] - p2[0], 2)
            + Math.pow(p1[1] - p2[1], 2);
    }

    /**
     * p1 から p2 への角度
     * @param {number[]} p1 - 起点
     * @param {number[]} p2 - 終点
     * @returns {number} 角度（ラジアン）
     */
    function getRad(p1, p2) {
        return Math.atan2(p2[1] - p1[1],
            p2[0] - p1[0]);
    }

    /**
     * o を頂点とする p1-o-p2 の角度
     * @param {number[]} p1 - 点1
     * @param {number[]} o - 頂点
     * @param {number[]} p2 - 点2
     * @returns {number} 角度（ラジアン）
     */
    function getRad2(p1, o, p2) {
        var v1 = normalize(p1, o);
        var v2 = normalize(p2, o);
        return Math.acos(v1[0] * v2[0] + v1[1] * v2[1]);
    }

    /**
     * o から p への単位ベクトル
     * @param {number[]} p - 点
     * @param {number[]} o - 原点
     * @returns {number[]} 単位ベクトル
     */
    function normalize(p, o) {
        var d = dist(p, o);
        return d == 0 ? [0, 0] : [(p[0] - o[0]) / d,
        (p[1] - o[1]) / d];
    }

    /**
     * ベジェ曲線上で長さ len の位置の t を返す（len が 0 なら全長）
     * @param {Array} q - 制御点4つ
     * @param {number} len - 長さ（負なら終点側から）
     * @returns {number} t、または全長
     */
    function getT4Len(q, len) {
        var m = [q[3][0] - q[0][0] + 3 * (q[1][0] - q[2][0]),
        q[0][0] - 2 * q[1][0] + q[2][0],
        q[1][0] - q[0][0]];
        var n = [q[3][1] - q[0][1] + 3 * (q[1][1] - q[2][1]),
        q[0][1] - 2 * q[1][1] + q[2][1],
        q[1][1] - q[0][1]];
        var k = [m[0] * m[0] + n[0] * n[0],
        4 * (m[0] * m[1] + n[0] * n[1]),
        2 * ((m[0] * m[2] + n[0] * n[2]) + 2 * (m[1] * m[1] + n[1] * n[1])),
        4 * (m[1] * m[2] + n[1] * n[2]),
        m[2] * m[2] + n[2] * n[2]];

        var fullLen = getLength(k, 1);

        if (len == 0) {
            return fullLen;
        } else if (len < 0) {
            len += fullLen;
            if (len < 0) return 0;
        } else if (len > fullLen) {
            return 1;
        }

        var t, d;
        var t0 = 0;
        var t1 = 1;
        var tolerance = 0.001;

        for (var h = 1; h < 30; h++) {
            t = t0 + (t1 - t0) / 2;
            d = len - getLength(k, t);
            if (Math.abs(d) < tolerance) break;
            else if (d < 0) t1 = t;
            else t0 = t;
        }
        return t;
    }

    /**
     * ベジェ曲線の 0〜t の長さ（シンプソン則）
     * @param {number[]} k - 係数
     * @param {number} t - 終わりの位置
     * @returns {number} 長さ
     */
    function getLength(k, t) {
        var h = t / 128;
        var hh = h * 2;
        var fc = function (t, k) {
            return Math.sqrt(t * (t * (t * (t * k[0] + k[1]) + k[2]) + k[3]) + k[4]) || 0;
        };
        var total = (fc(0, k) - fc(t, k)) / 2;
        for (var i = h; i < t; i += hh) total += 2 * fc(i, k) + fc(i + h, k);
        return total * hh;
    }

    /**
     * 重なったアンカーを1つにまとめる
     * @param {PathPoints} p - パスポイント
     * @returns {number} まとめたあとの点の数
     */
    function readjustAnchors(p) {
        var minDist = 0.0025;
        if (p.length < 2) return 1;
        var i;

        if (p.parent.closed) {
            for (i = p.length - 1; i >= 1; i--) {
                if (dist2(p[0].anchor, p[i].anchor) < minDist) {
                    p[0].leftDirection = p[i].leftDirection;
                    p[i].remove();
                } else {
                    break;
                }
            }
        }

        for (i = p.length - 1; i >= 1; i--) {
            if (dist2(p[i].anchor, p[i - 1].anchor) < minDist) {
                p[i - 1].rightDirection = p[i].rightDirection;
                p[i].remove();
            }
        }

        return p.length;
    }

    /**
     * 添字を点の数の範囲に収める（閉じたパスは巡回、開いたパスは範囲外で -1）
     * @param {PathPoints} p - パスポイント
     * @param {number} n - 添字
     * @returns {number} 収めた添字
     */
    function parseIdx(p, n) {
        var len = p.length;
        if (p.parent.closed) {
            return n >= 0 ? n % len : len - Math.abs(n % len);
        } else {
            return (n < 0 || n > len - 1) ? -1 : n;
        }
    }

    /**
     * パスポイントの座標と種類を配列で返す
     * @param {PathPoint} p - パスポイント
     * @returns {Array} [アンカー, 右ハンドル, 左ハンドル, 種類]
     */
    function getDat(p) {
        with (p) return [anchor, rightDirection, leftDirection, pointType];
    }

    /**
     * アンカーが選択されているか
     * @param {PathPoint} p - パスポイント
     * @returns {boolean} 選択されているか
     */
    function isSelected(p) {
        return p.selected == PathPointSelection.ANCHORPOINT;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した1つのオブジェクトを、外接矩形を2分割した2色の背景に置き換える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.openDocument"));
            return;
        }

        doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字の編集中は TextRange が返り、length は文字数になる / while editing text, the selection is a TextRange whose length counts characters */
        if (!selectedItems || selectedItems.typename === "TextRange" || selectedItems.length !== 1) {
            alert(getLabel("alert.selectOneItem"));
            return;
        }

        targetItem = selectedItems[0];
        targetBounds = getItemBounds(targetItem);

        var originalActiveLayer = doc.activeLayer;
        var drawOptions = showSettingsDialog();
        if (drawOptions) {
            var itemLayer = targetItem.layer;
            var drawnItems = drawSplitBackground(drawOptions, itemLayer);
            /* 元のオブジェクトは背景に置き換える / the background replaces the original object */
            targetItem.remove();
            doc.selection = null;
            if (drawOptions.groupItems) groupDrawnItems(drawnItems, itemLayer);
        } else {
            /* キャンセルしたら選択を戻す / restore the selection on cancel */
            doc.selection = [targetItem];
        }

        try {
            doc.activeLayer = originalActiveLayer;
        } catch (e) {
            /* ロックされたレイヤーなどは戻せないことがある / a locked layer may refuse to become active again */
        }
    }

    main();

})();
