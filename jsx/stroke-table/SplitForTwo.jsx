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
var SCRIPT_VERSION  = "v2.10.1";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

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
    var PICKER_PREVIEW_SIZE = [90, 40];          /* カラー選択の色見本の大きさ / color picker preview size */
    var PICKER_LABEL_WIDTH = 20;                 /* カラー選択のチャンネル名の幅 / color picker channel label width */
    var PICKER_SLIDER_SIZE = [120, 20];          /* カラー選択のスライダーの大きさ / color picker slider size */
    var PICKER_VALUE_CHARS = 3;                  /* カラー選択の数値欄の文字数 / characters for the color picker value fields */
    var PICKER_HEX_CHARS = 6;                    /* 16進数の入力欄の文字数 / characters for the hex field */
    var PICKER_BUTTON_MARGINS = [0, 10, 0, 0];   /* カラー選択のボタン行の余白 / color picker button row margins */

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

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトの分割背景を作成", en: "Create Split Background" },
            colorPicker: { ja: "カラー選択", en: "Color Picker" },
            presetSave: { ja: "プリセット保存", en: "Save Preset" }
        },
        panel: {
            splitDirection: { ja: "分割方法", en: "Split Direction" },
            balance: { ja: "バランス", en: "Balance" },
            fill: { ja: "塗り", en: "Fill" },
            stroke: { ja: "線", en: "Stroke" },
            cornerRadius: { ja: "角丸", en: "Corner radius" }
        },
        tab: {
            rgb: { ja: "RGB", en: "RGB" },
            cmyk: { ja: "CMYK", en: "CMYK" }
        },
        radio: {
            splitLR: { ja: "左右", en: "Left/Right" },
            splitTB: { ja: "上下", en: "Top/Bottom" },
            white: { ja: "白", en: "White" },
            black: { ja: "黒", en: "Black" },
            custom: { ja: "カスタム", en: "Custom" }
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
            groupItems: { ja: "グループ化", en: "Group items" },
            gray: { ja: "グレー", en: "Gray" }
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
            colorType: {
                ja: "塗りに使う色の決め方です。「カスタム」を選ぶと下の欄で指定できます。",
                en: "How the fill color is chosen. Custom lets you set it in the fields below."
            },
            hex: { ja: "塗りの色を16進数で指定します（例: DCDCDC）。", en: "Fill color as a hex value, for example DCDCDC." },
            gray: {
                ja: "CMYKのK版だけで色を作ります（グレースケール）。",
                en: "Builds the color from the K plate only, giving a grayscale."
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

    /**
     * ドット区切りキーで現在の言語のラベルを取得する
     * @param {string} key - LABELS のキー（例: "panel.balance"）
     * @returns {string} ラベル文字列
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            labelNode = labelNode[keyParts[i]];
        }
        return labelNode[uiLang] || labelNode.en;
    }

    /**
     * コロン付きのラベルを取得する（日本語は全角、英語は半角）
     * @param {string} key - LABELS のキー
     * @returns {string} コロンを付けたラベル文字列
     */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // セッション設定 / Session settings
    // =========================================
    // Illustrator の起動中だけダイアログの値を保持する（長さは単位を変えても崩れないよう pt で持つ）
    // Dialog values are kept while Illustrator is running; lengths are stored in pt so unit changes do not skew them

    var SESSION_KEY = "SplitForTwo_settings";
    var PRESETS_KEY = "SplitForTwo_presets";

    /**
     * セッションに残したダイアログの値を返す（欠けている項目は初期値で補う）
     * @returns {Object} セッション設定
     */
    function getSessionSettings() {
        var session = $.global[SESSION_KEY] || {};
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
        for (var key in defaults) {
            if (session[key] === undefined) session[key] = defaults[key];
        }
        $.global[SESSION_KEY] = session;
        return session;
    }

    /**
     * 保存したプリセットの置き場所を返す
     * @returns {Object} プリセット名をキーにした設定の集まり
     */
    function getPresetStore() {
        if (!$.global[PRESETS_KEY]) $.global[PRESETS_KEY] = {};
        return $.global[PRESETS_KEY];
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
        } else if (pageItem.typename === "GroupItem" && pageItem.clipped && pageItem.pageItems.length > 0) {
            /* クリップグループのマスクは pageItems[0] / the mask of a clip group is pageItems[0] */
            itemBounds = getItemBounds(pageItem.pageItems[0]);
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
     * RGB を CMYK に簡易換算する（近似値）
     * @param {{r: number, g: number, b: number}} rgb - RGB（0〜255）
     * @returns {{c: number, m: number, y: number, k: number}} CMYK（0〜100）
     */
    function rgbToCmykApprox(rgb) {
        var red = rgb.r / 255;
        var green = rgb.g / 255;
        var blue = rgb.b / 255;
        var black = 1 - Math.max(red, green, blue);
        if (black >= 1) return { c: 0, m: 0, y: 0, k: 100 };
        return {
            c: Math.round((1 - red - black) / (1 - black) * 100),
            m: Math.round((1 - green - black) / (1 - black) * 100),
            y: Math.round((1 - blue - black) / (1 - black) * 100),
            k: Math.round(black * 100)
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
    // 状態の mode は "rgb" | "cmyk" | "gray"、presetType は "white" | "black" | "custom"、dialogTab は "rgb" | "cmyk"
    // State: mode is "rgb" | "cmyk" | "gray", presetType is "white" | "black" | "custom", dialogTab is "rgb" | "cmyk"

    /**
     * 色文字列からカラー選択の状態を作る
     * @param {string} initialColor - "RRGGBB" か "cmyk:C,M,Y,K"
     * @returns {Object} カラー選択の状態
     */
    function createColorPickerState(initialColor) {
        var pickerState = {
            presetType: "custom",
            mode: "rgb",
            dialogTab: "rgb",
            rgb: { r: 0, g: 0, b: 0 },
            cmyk: { c: 0, m: 0, y: 0, k: 0 }
        };

        if (isCmykString(initialColor)) {
            var cmyk = parseCmykString(initialColor);
            var noColorInk = (cmyk.c === 0 && cmyk.m === 0 && cmyk.y === 0);
            pickerState.cmyk = cmyk;
            pickerState.rgb = cmykToRgbApprox(cmyk);
            pickerState.mode = (noColorInk && cmyk.k > 0) ? "gray" : "cmyk";
            pickerState.dialogTab = "cmyk";
            pickerState.presetType = (noColorInk && cmyk.k === 0) ? "white" : (noColorInk && cmyk.k === 100) ? "black" : "custom";
        } else {
            var rgb = parseHexRgb(initialColor, "000000");
            pickerState.rgb = rgb;
            pickerState.cmyk = rgbToCmykApprox(rgb);
            var isWhite = (rgb.r === 255 && rgb.g === 255 && rgb.b === 255);
            var isBlack = (rgb.r === 0 && rgb.g === 0 && rgb.b === 0);
            pickerState.presetType = isWhite ? "white" : isBlack ? "black" : "custom";
        }
        return pickerState;
    }

    /**
     * RGB から CMYK を求め直す
     * @param {Object} pickerState - カラー選択の状態
     * @returns {void}
     */
    function syncPickerCmykFromRgb(pickerState) {
        pickerState.cmyk = rgbToCmykApprox({
            r: Math.round(pickerState.rgb.r),
            g: Math.round(pickerState.rgb.g),
            b: Math.round(pickerState.rgb.b)
        });
    }

    /**
     * CMYK から RGB を求め直す
     * @param {Object} pickerState - カラー選択の状態
     * @returns {void}
     */
    function syncPickerRgbFromCmyk(pickerState) {
        pickerState.rgb = cmykToRgbApprox(pickerState.cmyk);
    }

    /**
     * 白・黒・カスタムを切り替える
     * @param {Object} pickerState - カラー選択の状態
     * @param {string} presetType - "white" | "black" | "custom"
     * @returns {void}
     */
    function setPickerPresetType(pickerState, presetType) {
        pickerState.presetType = presetType;

        if (presetType === "white" || presetType === "black") {
            var channelValue = (presetType === "white") ? 255 : 0;
            pickerState.rgb = { r: channelValue, g: channelValue, b: channelValue };
            syncPickerCmykFromRgb(pickerState);
            if (pickerState.dialogTab !== "cmyk") pickerState.mode = "rgb";
        } else if (pickerState.dialogTab === "cmyk") {
            if (pickerState.mode !== "gray") pickerState.mode = "cmyk";
            syncPickerRgbFromCmyk(pickerState);
        } else {
            pickerState.mode = "rgb";
            syncPickerCmykFromRgb(pickerState);
        }
    }

    /**
     * RGB の値を入れる（カスタムの RGB に切り替わる）
     * @param {Object} pickerState - カラー選択の状態
     * @param {number} red - R（0〜255）
     * @param {number} green - G（0〜255）
     * @param {number} blue - B（0〜255）
     * @returns {void}
     */
    function setPickerRgb(pickerState, red, green, blue) {
        pickerState.rgb = { r: Math.round(red), g: Math.round(green), b: Math.round(blue) };
        syncPickerCmykFromRgb(pickerState);
        pickerState.mode = "rgb";
        pickerState.dialogTab = "rgb";
        pickerState.presetType = "custom";
    }

    /**
     * CMYK の値を入れる（カスタムの CMYK かグレーに切り替わる）
     * @param {Object} pickerState - カラー選択の状態
     * @param {number} cyan - C（0〜100）
     * @param {number} magenta - M（0〜100）
     * @param {number} yellow - Y（0〜100）
     * @param {number} black - K（0〜100）
     * @param {boolean} keepGrayMode - グレーのままにするか
     * @returns {void}
     */
    function setPickerCmyk(pickerState, cyan, magenta, yellow, black, keepGrayMode) {
        pickerState.cmyk = { c: Math.round(cyan), m: Math.round(magenta), y: Math.round(yellow), k: Math.round(black) };
        syncPickerRgbFromCmyk(pickerState);
        pickerState.mode = keepGrayMode ? "gray" : "cmyk";
        pickerState.dialogTab = "cmyk";
        pickerState.presetType = "custom";
    }

    /**
     * カラー選択の状態を色文字列にする
     * @param {Object} pickerState - カラー選択の状態
     * @returns {string} "RRGGBB" か "cmyk:C,M,Y,K"
     */
    function serializeColorPickerState(pickerState) {
        if (pickerState.presetType === "white") return "FFFFFF";
        if (pickerState.presetType === "black") return "000000";
        if (pickerState.mode === "gray") return cmykStringFromValues({ c: 0, m: 0, y: 0, k: pickerState.cmyk.k });
        if (pickerState.mode === "cmyk") return cmykStringFromValues(pickerState.cmyk);
        return rgbToHex(pickerState.rgb);
    }

    /**
     * カラー選択の状態を画面表示用の RGB にする
     * @param {Object} pickerState - カラー選択の状態
     * @returns {{r: number, g: number, b: number}} RGB（0〜255）
     */
    function getPickerPreviewRgb(pickerState) {
        if (pickerState.presetType === "white") return { r: 255, g: 255, b: 255 };
        if (pickerState.presetType === "black") return { r: 0, g: 0, b: 0 };
        if (pickerState.mode === "gray" || pickerState.mode === "cmyk") return cmykToRgbApprox(pickerState.cmyk);
        return {
            r: Math.round(pickerState.rgb.r),
            g: Math.round(pickerState.rgb.g),
            b: Math.round(pickerState.rgb.b)
        };
    }

    /**
     * カラー選択のスライダー行（チャンネル名・スライダー・数値欄）を追加する
     * @param {Group} parent - 追加先
     * @param {string} channelLabel - チャンネル名（"R" など）
     * @param {number} initialValue - 初期値
     * @param {number} maxValue - 上限（RGB は 255、CMYK は 100）
     * @returns {Object} row・slider・input と、値が変わったときの関数を登録する setOnValueChange
     */
    function addPickerSliderRow(parent, channelLabel, initialValue, maxValue) {
        var onValueChange = null;
        var rowGroup = addRow(parent);
        var channelText = rowGroup.add("statictext", undefined, channelLabel);
        channelText.justify = "center";
        channelText.preferredSize = [PICKER_LABEL_WIDTH, -1];
        var slider = rowGroup.add("slider", undefined, initialValue, 0, maxValue);
        slider.preferredSize = PICKER_SLIDER_SIZE;
        var valueInput = rowGroup.add("edittext", undefined, String(Math.round(initialValue)));
        valueInput.characters = PICKER_VALUE_CHARS;

        /* 値を整数に丸めて範囲内に収め、スライダーと数値欄に反映してから知らせる
           Round and clamp the value, show it on the slider and field, then notify */
        var commitValue = function (value, typedInInput) {
            value = clampNumber(Math.round(value), 0, maxValue);
            slider.value = value;
            /* 入力中の欄は書き換えない / leave the field alone while it is being typed in */
            if (!typedInInput) valueInput.text = String(value);
            if (typeof onValueChange === "function") onValueChange(value, typedInInput ? valueInput : null);
        };

        slider.onChanging = function () {
            var value = slider.value;
            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.shiftKey) {
                /* shift で10%刻み / snap to 10% steps with Shift */
                var shiftStep = maxValue * 0.1;
                value = Math.round(value / shiftStep) * shiftStep;
            } else if (keyboard.altKey) {
                /* option で RGB は16、CMYK は5%刻み / snap to 16 (RGB) or 5% (CMYK) with Option */
                var altStep = (maxValue === 255) ? 16 : maxValue * 0.05;
                value = Math.round(value / altStep) * altStep;
            }
            commitValue(value, false);
        };

        /* shift+↑↓ でスライダーを10%ずつ動かす / Shift+Up/Down moves the slider by 10% */
        slider.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            if (!ScriptUI.environment.keyboardState.shiftKey) return;
            var delta = Math.max(Math.round(maxValue * 0.1), 1);
            commitValue(slider.value + ((event.keyName === "Up") ? delta : -delta), false);
            event.preventDefault();
        });

        /* ↑↓で±1、shift+↑↓で±10 / Up/Down steps by 1, Shift+Up/Down by 10 */
        valueInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var value = parseFloat(valueInput.text);
            if (isNaN(value)) value = 0;
            var delta = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
            commitValue(value + ((event.keyName === "Up") ? delta : -delta), false);
            event.preventDefault();
        });

        valueInput.addEventListener("changing", function () {
            var value = parseFloat(valueInput.text);
            if (!isNaN(value)) commitValue(value, true);
        });

        return {
            row: rowGroup,
            slider: slider,
            input: valueInput,
            setOnValueChange: function (callback) { onValueChange = callback; }
        };
    }

    /**
     * スライダー行に値を入れる
     * @param {Object} sliderRow - addPickerSliderRow() の戻り値
     * @param {number} value - 値
     * @param {EditText|null} skipInput - 書き換えない入力欄（入力中の欄）
     * @returns {void}
     */
    function setPickerSliderRow(sliderRow, value, skipInput) {
        sliderRow.slider.value = Math.round(value);
        if (sliderRow.input !== skipInput) sliderRow.input.text = String(Math.round(value));
    }

    /**
     * カラー選択の状態をダイアログに映す
     * @param {Object} pickerState - カラー選択の状態
     * @param {Object} pickerControls - カラー選択のコントロール
     * @param {Object} [renderOptions] - suppressTabSelection（タブを切り替えない）、skipInput（書き換えない欄）
     * @returns {void}
     */
    function renderColorPicker(pickerState, pickerControls, renderOptions) {
        renderOptions = renderOptions || {};
        var skipInput = renderOptions.skipInput || null;

        pickerControls.whiteRadio.value = (pickerState.presetType === "white");
        pickerControls.blackRadio.value = (pickerState.presetType === "black");
        pickerControls.customRadio.value = (pickerState.presetType === "custom");

        if (!renderOptions.suppressTabSelection) {
            var targetTab = (pickerState.dialogTab === "cmyk") ? pickerControls.cmykTab : pickerControls.rgbTab;
            if (pickerControls.tabPanel.selection !== targetTab) pickerControls.tabPanel.selection = targetTab;
        }

        pickerControls.grayCheckbox.value = (pickerState.mode === "gray");

        setPickerSliderRow(pickerControls.redRow, pickerState.rgb.r, skipInput);
        setPickerSliderRow(pickerControls.greenRow, pickerState.rgb.g, skipInput);
        setPickerSliderRow(pickerControls.blueRow, pickerState.rgb.b, skipInput);
        if (pickerControls.hexInput !== skipInput) pickerControls.hexInput.text = rgbToHex(pickerState.rgb);

        setPickerSliderRow(pickerControls.cyanRow, pickerState.cmyk.c, skipInput);
        setPickerSliderRow(pickerControls.magentaRow, pickerState.cmyk.m, skipInput);
        setPickerSliderRow(pickerControls.yellowRow, pickerState.cmyk.y, skipInput);
        setPickerSliderRow(pickerControls.blackRow, pickerState.cmyk.k, skipInput);

        /* 白・黒のときはタブごと、グレーのときは C・M・Y を使えなくする / disable the tabs for white or black, and C/M/Y for gray */
        var customEnabled = (pickerState.presetType === "custom");
        var colorInkEnabled = customEnabled && pickerState.mode !== "gray";
        pickerControls.tabPanel.enabled = customEnabled;
        pickerControls.cyanRow.row.enabled = colorInkEnabled;
        pickerControls.magentaRow.row.enabled = colorInkEnabled;
        pickerControls.yellowRow.row.enabled = colorInkEnabled;

        pickerControls.refreshPreview();
    }

    /**
     * カラー選択ダイアログを組み立てる（イベントは bindColorPickerEvents() で付ける）
     * @param {Object} pickerState - カラー選択の状態
     * @param {{r: number, g: number, b: number}} initialRgb - 開いたときの色（比較用）
     * @returns {Object} ダイアログとコントロールの参照
     */
    function buildColorPickerDialog(pickerState, initialRgb) {
        var pickerControls = {};
        var pickerDialog = new Window("dialog", getLabel("dialog.colorPicker"));
        pickerDialog.orientation = "column";
        pickerDialog.alignChildren = ["fill", "top"];
        pickerControls.dialog = pickerDialog;

        /* 開いたときの色と今の色を並べる / the original color next to the current one */
        var previewRowGroup = pickerDialog.add("group");
        previewRowGroup.orientation = "row";
        previewRowGroup.alignment = ["center", "top"];
        previewRowGroup.spacing = 1;
        var originalPreview = previewRowGroup.add("group");
        originalPreview.preferredSize = PICKER_PREVIEW_SIZE;
        originalPreview.onDraw = function () { paintSolidColor(this, initialRgb, false); };
        var currentPreview = previewRowGroup.add("group");
        currentPreview.preferredSize = PICKER_PREVIEW_SIZE;
        currentPreview.onDraw = function () { paintSolidColor(this, getPickerPreviewRgb(pickerState), false); };
        pickerControls.refreshPreview = function () { repaintControl(currentPreview); };

        var colorTypeRowGroup = addRow(pickerDialog);
        colorTypeRowGroup.alignment = ["center", "top"];
        pickerControls.whiteRadio = addRadio(colorTypeRowGroup, "radio.white", "tooltip.colorType");
        pickerControls.blackRadio = addRadio(colorTypeRowGroup, "radio.black", "tooltip.colorType");
        pickerControls.customRadio = addRadio(colorTypeRowGroup, "radio.custom", "tooltip.colorType");

        var tabPanel = pickerDialog.add("tabbedpanel");
        tabPanel.alignment = ["fill", "top"];
        tabPanel.alignChildren = ["fill", "top"];
        pickerControls.tabPanel = tabPanel;

        /* RGB タブ / RGB tab */
        var rgbTab = tabPanel.add("tab", undefined, getLabel("tab.rgb"));
        rgbTab.orientation = "column";
        rgbTab.alignChildren = ["fill", "top"];
        rgbTab.margins = PANEL_MARGINS;
        pickerControls.rgbTab = rgbTab;

        var rgbSliderGroup = addColumn(rgbTab, "fill");
        pickerControls.redRow = addPickerSliderRow(rgbSliderGroup, "R", pickerState.rgb.r, 255);
        pickerControls.greenRow = addPickerSliderRow(rgbSliderGroup, "G", pickerState.rgb.g, 255);
        pickerControls.blueRow = addPickerSliderRow(rgbSliderGroup, "B", pickerState.rgb.b, 255);

        var hexRowGroup = addRow(rgbTab);
        hexRowGroup.alignment = ["center", "top"];
        hexRowGroup.add("statictext", undefined, "#");
        pickerControls.hexInput = hexRowGroup.add("edittext", undefined, rgbToHex(pickerState.rgb));
        pickerControls.hexInput.helpTip = getLabel("tooltip.hex");
        pickerControls.hexInput.characters = PICKER_HEX_CHARS;

        /* CMYK タブ / CMYK tab */
        var cmykTab = tabPanel.add("tab", undefined, getLabel("tab.cmyk"));
        cmykTab.orientation = "column";
        cmykTab.alignChildren = ["fill", "top"];
        cmykTab.margins = PANEL_MARGINS;
        pickerControls.cmykTab = cmykTab;

        pickerControls.grayCheckbox = addCheckbox(cmykTab, "checkbox.gray", "tooltip.gray", false);

        var cmykSliderGroup = addColumn(cmykTab, "fill");
        pickerControls.cyanRow = addPickerSliderRow(cmykSliderGroup, "C", pickerState.cmyk.c, 100);
        pickerControls.magentaRow = addPickerSliderRow(cmykSliderGroup, "M", pickerState.cmyk.m, 100);
        pickerControls.yellowRow = addPickerSliderRow(cmykSliderGroup, "Y", pickerState.cmyk.y, 100);
        pickerControls.blackRow = addPickerSliderRow(cmykSliderGroup, "K", pickerState.cmyk.k, 100);

        var btnRowGroup = pickerDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["center", "center"];
        btnRowGroup.margins = PICKER_BUTTON_MARGINS;
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return pickerControls;
    }

    /**
     * カラー選択ダイアログのイベントを付ける
     * @param {Object} pickerState - カラー選択の状態
     * @param {Object} pickerControls - カラー選択のコントロール
     * @returns {void}
     */
    function bindColorPickerEvents(pickerState, pickerControls) {
        /* 描き直し中に届いたイベントでは描き直さない / ignore events raised while the dialog is being redrawn */
        var isRendering = false;
        var render = function (renderOptions) {
            if (isRendering) return;
            isRendering = true;
            try {
                renderColorPicker(pickerState, pickerControls, renderOptions);
            } finally {
                isRendering = false;
            }
        };

        pickerControls.redRow.setOnValueChange(function (value, typedInput) {
            setPickerRgb(pickerState, value, pickerState.rgb.g, pickerState.rgb.b);
            render({ skipInput: typedInput });
        });
        pickerControls.greenRow.setOnValueChange(function (value, typedInput) {
            setPickerRgb(pickerState, pickerState.rgb.r, value, pickerState.rgb.b);
            render({ skipInput: typedInput });
        });
        pickerControls.blueRow.setOnValueChange(function (value, typedInput) {
            setPickerRgb(pickerState, pickerState.rgb.r, pickerState.rgb.g, value);
            render({ skipInput: typedInput });
        });

        pickerControls.cyanRow.setOnValueChange(function (value, typedInput) {
            var cmyk = pickerState.cmyk;
            setPickerCmyk(pickerState, value, cmyk.m, cmyk.y, cmyk.k, pickerControls.grayCheckbox.value);
            render({ skipInput: typedInput });
        });
        pickerControls.magentaRow.setOnValueChange(function (value, typedInput) {
            var cmyk = pickerState.cmyk;
            setPickerCmyk(pickerState, cmyk.c, value, cmyk.y, cmyk.k, pickerControls.grayCheckbox.value);
            render({ skipInput: typedInput });
        });
        pickerControls.yellowRow.setOnValueChange(function (value, typedInput) {
            var cmyk = pickerState.cmyk;
            setPickerCmyk(pickerState, cmyk.c, cmyk.m, value, cmyk.k, pickerControls.grayCheckbox.value);
            render({ skipInput: typedInput });
        });
        pickerControls.blackRow.setOnValueChange(function (value, typedInput) {
            var cmyk = pickerState.cmyk;
            setPickerCmyk(pickerState, cmyk.c, cmyk.m, cmyk.y, value, pickerControls.grayCheckbox.value);
            render({ skipInput: typedInput });
        });

        /* 16進数は6桁そろったときだけ反映する / apply the hex value only once all six digits are in */
        pickerControls.hexInput.addEventListener("changing", function () {
            var hex = pickerControls.hexInput.text.replace(/^#/, "");
            if (!/^[0-9A-Fa-f]{6}$/.test(hex)) return;
            var hexRgb = parseHexRgb(hex, "000000");
            setPickerRgb(pickerState, hexRgb.r, hexRgb.g, hexRgb.b);
            render({ skipInput: pickerControls.hexInput });
        });

        pickerControls.whiteRadio.onClick = function () {
            setPickerPresetType(pickerState, "white");
            render();
        };
        pickerControls.blackRadio.onClick = function () {
            setPickerPresetType(pickerState, "black");
            render();
        };
        pickerControls.customRadio.onClick = function () {
            setPickerPresetType(pickerState, "custom");
            render();
        };

        pickerControls.tabPanel.onChange = function () {
            if (isRendering) return;
            pickerState.dialogTab = (pickerControls.tabPanel.selection === pickerControls.cmykTab) ? "cmyk" : "rgb";
            if (pickerState.dialogTab === "rgb") {
                syncPickerRgbFromCmyk(pickerState);
                pickerState.mode = "rgb";
            } else {
                syncPickerCmykFromRgb(pickerState);
                pickerState.mode = pickerControls.grayCheckbox.value ? "gray" : "cmyk";
            }
            render({ suppressTabSelection: true });
        };

        /* グレーにするときは C・M・Y を0にする / switching to gray clears C, M and Y */
        pickerControls.grayCheckbox.onClick = function () {
            if (pickerControls.grayCheckbox.value) {
                pickerState.cmyk.c = 0;
                pickerState.cmyk.m = 0;
                pickerState.cmyk.y = 0;
                pickerState.mode = "gray";
            } else {
                pickerState.mode = "cmyk";
            }
            pickerState.dialogTab = "cmyk";
            pickerState.presetType = "custom";
            syncPickerRgbFromCmyk(pickerState);
            render();
        };
    }

    /**
     * カラー選択ダイアログを表示する
     * @param {string} initialColor - 開いたときの色（"RRGGBB" か "cmyk:C,M,Y,K"）
     * @returns {string|null} 選んだ色の文字列。キャンセルなら null
     */
    function showColorPickerDialog(initialColor) {
        var pickerState = createColorPickerState(initialColor);
        var initialRgb = { r: pickerState.rgb.r, g: pickerState.rgb.g, b: pickerState.rgb.b };
        var pickerControls = buildColorPickerDialog(pickerState, initialRgb);

        renderColorPicker(pickerState, pickerControls);
        bindColorPickerEvents(pickerState, pickerControls);

        prepareDialogWindow(pickerControls.dialog, SCRIPT_NAME + "_colorPicker");
        if (pickerControls.dialog.show() !== 1) return null;
        return serializeColorPickerState(pickerState);
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

        var btnRowGroup = nameDialog.add("group");
        btnRowGroup.alignment = ["right", "center"];
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
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

        var btnRowGroup = settingsDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["center", "center"];
        controls.btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        controls.btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
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
            getPresetStore()[presetName] = collectPresetData(controls);
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
        controls.dialog.addEventListener("keydown", function (event) {
            if (event.keyName === "H") {
                setSplitDirection(controls, false);
            } else if (event.keyName === "V") {
                setSplitDirection(controls, true);
            } else if (event.keyName === "F") {
                controls.overallFrameCheckbox.value = !controls.overallFrameCheckbox.value;
                onStrokeOptionClick();
            } else if (event.keyName === "D") {
                controls.dividerCheckbox.value = !controls.dividerCheckbox.value;
                onStrokeOptionClick();
            } else {
                return;
            }
            event.preventDefault();
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
     * ダイアログの値をセッションに残す（不正な値の欄は前回の値のまま）
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
