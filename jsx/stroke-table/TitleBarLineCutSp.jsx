#target illustrator
#targetengine "TitleBarLineCutSpEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの周りで、罫線を欠けさせます。
オフセット・線幅・線端・角の形状をダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TitleBarLineCutSp.md

### Overview

Cuts a gap into a rule around the selected objects.
Offset, stroke weight, cap and join are set in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TitleBarLineCutSp.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TitleBarLineCutSp";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TitleBarLineCutSp.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TitleBarLineCutSp.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    // 日英ラベル定義 / Japanese-English label definitions
    var LABELS = {
        dialogTitle: { ja: "罫線の欠け処理", en: "Line Gap Cut" },
        offset: { ja: "オフセット", en: "Offset" },
        strokeWidth: { ja: "線幅", en: "Stroke" },
        strokePanel: { ja: "線", en: "Stroke" }, // Stroke settings panel label
        capPanel: { ja: "線端", en: "Cap" },
        tipOffset: { ja: "元のオブジェクトから線を離す距離です。", en: "How far the rule sits from the original object." },
        tipStroke: { ja: "線の太さです。", en: "Weight of the rule." },
        tipCapButt: { ja: "線の端を切りっぱなしにします。", en: "Leaves the line ends flat." },
        tipCapRound: { ja: "線の端を丸くします。", en: "Rounds the line ends." },
        tipJoinMiter: { ja: "角を尖らせたまま結合します。", en: "Keeps the corners pointed." },
        tipJoinRound: { ja: "角を丸めて結合します。", en: "Rounds off the corners." },
        tipJoinBevel: { ja: "角を面取りして結合します。", en: "Cuts the corners off flat." },
        capButt: { ja: "なし", en: "Butt" },
        capRound: { ja: "丸型線端", en: "Round" },
        joinPanel: { ja: "角の形状", en: "Join" },
        joinMiter: { ja: "マイター", en: "Miter" },
        joinRound: { ja: "ラウンド", en: "Round" },
        joinBevel: { ja: "ベベル", en: "Bevel" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        alertSelectTwo: { ja: "2つのオブジェクト（円と長方形）を選択してください。", en: "Select two objects (circle and rectangle)." },
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        }
    };

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

    /**
     * 行に「∧∨＋入力欄」を追加する（↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} rowGroup - 追加先の行
     * @param {string} initialText - 初期値
     * @param {Object} stepOptions - min / max / integer（addStepper() に渡す）
     * @returns {EditText} 入力欄
     */
    function addSteppedInput(rowGroup, initialText, stepOptions) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = rowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
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

    (function () {
        // ドキュメントが開かれているか確認
        if (app.documents.length === 0) {
            alert(getLabel('alertNoDoc'));
            return;
        }

        var doc = app.activeDocument;

        var currentSelection = doc.selection;

        // 2つのオブジェクトが選択されているか確認
        if (!currentSelection || currentSelection.length !== 2) {
            alert(getLabel('alertSelectTwo'));
            return;
        }

        // selection[0] が最前面（上） = 長方形 (B)
        // selection[1] が最背面（下） = 円 (A)
        var objectB = currentSelection[0];
        var objectA = currentSelection[1];

        // 最後に必ずBを再表示するための保険
        var __restoreB = objectB;

        function safeDo(fn) { try { fn(); } catch (e) { } }
        function safeRemove(it) { safeDo(function () { if (it) it.remove(); }); }
        function safeSelect(it) {
            safeDo(function () { doc.selection = null; });
            safeDo(function () { if (it) it.selected = true; });
        }
        function safeMenu(cmd) { safeDo(function () { app.executeMenuCommand(cmd); }); }
        /**
         * Illustrator 単位ユーティリティ（表示ラベル＋pt換算）
         * - rulerType: 定規
         * - strokeUnits: 線
         * - text/units: 文字
         * - text/asianunits: 東アジア言語（Q/H 表示切替）
         */
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

        function unitToPt(valueInUnit, prefKey) {
            return valueInUnit * getUnitInfo(prefKey).pointsPerUnit;
        }

        function ptToUnit(valuePt, prefKey) {
            return valuePt / getUnitInfo(prefKey).pointsPerUnit;
        }

        function round1(v) {
            return Math.round(v * 10) / 10;
        }
        // デフォルトのオフセット値：長方形化（C）の高さの1/3（pt）
        function getCHeightPtFromB(bItem) {
            if (!bItem) return 0;

            // TextFrame：複製→アウトライン→外接bounds（後始末あり）
            if (bItem.typename === 'TextFrame') {
                var tmpDup = null;
                var tmpOutlined = [];
                var region = null;

                try {
                    tmpDup = bItem.duplicate();
                    safeSelect(tmpDup);
                    safeMenu('outline');

                    var s = doc.selection;
                    if (s && s.length) {
                        for (var i = 0; i < s.length; i++) tmpOutlined.push(s[i]);
                    }

                    if (tmpOutlined.length > 1) {
                        var g = doc.groupItems.add();
                        for (var j = tmpOutlined.length - 1; j >= 0; j--) {
                            safeDo(function () { tmpOutlined[j].move(g, ElementPlacement.PLACEATBEGINNING); });
                        }
                        region = g.geometricBounds;
                        tmpOutlined = [g];
                    } else if (tmpOutlined.length === 1) {
                        region = tmpOutlined[0].geometricBounds;
                    }
                } catch (e) {
                    region = null;
                } finally {
                    for (var k = 0; k < tmpOutlined.length; k++) safeRemove(tmpOutlined[k]);
                    safeRemove(tmpDup);
                    safeDo(function () { doc.selection = null; });
                }

                if (!region) return 0;
                var h = region[1] - region[3];
                return (h > 0) ? h : 0;
            }

            // 通常：Bのbounds高さを採用（CはB複製）
            var b = bItem.geometricBounds; // [L,T,R,B]
            var h2 = b[1] - b[3];
            return (h2 > 0) ? h2 : 0;
        }

        var __cHeightPtForDefault = getCHeightPtFromB(objectB);
        // 高さが取れない場合のフォールバックは従来の 10pt
        var __offsetDefaultPt = (__cHeightPtForDefault > 0) ? (__cHeightPtForDefault / 3) : 10;

        // オフセット値ダイアログ
        var dlg = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        setupWindow(dlg);

        var grp1 = dlg.add('group');
        var lblOffset = grp1.add('statictext', undefined, getLabel('offset'));
        // lblOffset.justification = 'right';
        var __offsetPrefKey = "rulerType";
        // __offsetDefaultPt は「C高さ/3（pt）」で事前計算済み
        var offsetInput = addSteppedInput(grp1, String(round1(ptToUnit(__offsetDefaultPt, __offsetPrefKey))), { min: 0 });
        offsetInput.helpTip = getLabel('tipOffset');
        offsetInput.characters = 3;
        grp1.add('statictext', undefined, getUnitInfo(__offsetPrefKey).label);

        // 線設定パネル
        var panelStroke = dlg.add('panel', undefined, getLabel('strokePanel'));
        setupPanel(panelStroke);

        var grp2 = panelStroke.add('group');
        var lblStroke = grp2.add('statictext', undefined, getLabel('strokeWidth'));

        var __strokePrefKey = "strokeUnits";
        var __strokeDefaultPt = 2;
        var strokeInput = addSteppedInput(grp2, String(round1(ptToUnit(__strokeDefaultPt, __strokePrefKey))), { min: 0 });
        strokeInput.helpTip = getLabel('tipStroke');
        strokeInput.characters = 3;
        grp2.add('statictext', undefined, getUnitInfo(__strokePrefKey).label);

        // 線端設定
        var panelCap = panelStroke.add('panel', undefined, getLabel('capPanel'));
        setupPanel(panelCap, 6);

        var rbCapButt = panelCap.add('radiobutton', undefined, getLabel('capButt'));
        rbCapButt.helpTip = getLabel('tipCapButt');
        var rbCapRound = panelCap.add('radiobutton', undefined, getLabel('capRound'));
        rbCapRound.helpTip = getLabel('tipCapRound');

        rbCapButt.value = true; // デフォルト

        // 丸型線端が選択されたら、角の形状をラウンドに自動切替
        rbCapRound.onClick = function () {
            if (rbCapRound.value) {
                rbJoinRound.value = true;
            }
        };

        // 角の形状（線のジョイン）
        var panelJoin = panelStroke.add('panel', undefined, getLabel('joinPanel'));
        setupPanel(panelJoin, 6);

        var rbJoinMiter = panelJoin.add('radiobutton', undefined, getLabel('joinMiter'));
        rbJoinMiter.helpTip = getLabel('tipJoinMiter');
        var rbJoinRound = panelJoin.add('radiobutton', undefined, getLabel('joinRound'));
        rbJoinRound.helpTip = getLabel('tipJoinRound');
        var rbJoinBevel = panelJoin.add('radiobutton', undefined, getLabel('joinBevel'));
        rbJoinBevel.helpTip = getLabel('tipJoinBevel');

        rbJoinMiter.value = true; // デフォルト

        var buttonRow = addButtonRow(dlg);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('ok'), { name: 'ok' });

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dlg, SCRIPT_NAME);
        if (dlg.show() !== 1) return;

        var offsetValue = parseFloat(offsetInput.text);
        if (isNaN(offsetValue)) offsetValue = ptToUnit(10, __offsetPrefKey);
        var offsetValuePt = unitToPt(offsetValue, __offsetPrefKey);

        var strokeWidthValue = parseFloat(strokeInput.text);
        if (isNaN(strokeWidthValue)) strokeWidthValue = ptToUnit(2, __strokePrefKey);
        var strokeWidthValuePt = unitToPt(strokeWidthValue, __strokePrefKey);

        var strokeCapValue = rbCapRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
        var strokeJoinValue = rbJoinRound.value ? StrokeJoin.ROUNDENDJOIN : (rbJoinBevel.value ? StrokeJoin.BEVELENDJOIN : StrokeJoin.MITERENDJOIN);

        try {

            //念のためロックされていたら解除
            objectA.locked = false;
            objectB.locked = false;

            // 0. selection[1]（=A）を複製し、ガイド化
            //    ガイドはロックしておく（誤操作防止）
            (function makeGuideFromA() {
                var dup = null;
                try {
                    dup = objectA.duplicate();
                    safeDo(function () { dup.name = '__GUIDE_A__'; });

                    // ガイド化（メニュー優先、失敗時はプロパティでフォールバック）
                    safeSelect(dup);
                    safeMenu('makeGuide');
                    // makeGuide が失敗した場合のみ guides=true を試す
                    safeDo(function () {
                        if (dup && dup.guides !== true) dup.guides = true;
                    });

                    safeDo(function () { dup.locked = true; });
                    safeDo(function () { doc.selection = null; });
                } catch (e) {
                    safeRemove(dup);
                    safeDo(function () { doc.selection = null; });
                }
            })();

            // 1. Cを作成
            //  - 通常：Bを複製してC
            //  - Bがテキストの場合：複製→アウトライン→バウンディングから長方形化（C）
            function makeRectFromBounds(bounds) {
                // bounds: [L, T, R, B]
                var L = bounds[0], T = bounds[1], R = bounds[2], B = bounds[3];
                var w = R - L;
                var h = T - B;
                // rectangle(top, left, width, height)
                var rect = doc.pathItems.rectangle(T, L, w, h);
                return rect;
            }

            function createCFromTextFrame(tf) {
                // tf: TextFrame
                var tmpDup = null;
                var tmpOutlined = [];
                try {
                    // 1) 複製（A）
                    tmpDup = tf.duplicate();

                    // 2) アウトライン化（B）
                    safeSelect(tmpDup);
                    safeMenu('outline');

                    // outline後の選択を回収
                    var s = doc.selection;
                    if (!s || s.length === 0) return null;

                    // 選択されたものを一旦配列に退避
                    for (var i = 0; i < s.length; i++) tmpOutlined.push(s[i]);

                    // 3) 長方形化（C）: アウトラインの外接バウンディングから矩形を作る
                    var region = null;

                    // 可能ならグループ化して bounds を安定させる
                    if (tmpOutlined.length > 1) {
                        var g = doc.groupItems.add();
                        for (var j = tmpOutlined.length - 1; j >= 0; j--) {
                            try { tmpOutlined[j].move(g, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
                        }
                        region = g.geometricBounds;
                        // グループを後で消せるように差し替え
                        tmpOutlined = [g];
                    } else {
                        region = tmpOutlined[0].geometricBounds;
                    }

                    if (!region) return null;

                    var rectC = makeRectFromBounds(region);

                    // アウトライン化した一時オブジェクトは削除
                    for (var k = 0; k < tmpOutlined.length; k++) {
                        safeRemove(tmpOutlined[k]);
                    }
                    safeRemove(tmpDup);
                    safeDo(function () { doc.selection = null; });
                    return rectC;
                } catch (e) {
                    // 後始末
                    safeRemove(tmpDup);
                    for (var x = 0; x < tmpOutlined.length; x++) {
                        safeRemove(tmpOutlined[x]);
                    }
                    safeDo(function () { doc.selection = null; });
                    return null;
                }
            }

            var objectC = null;
            if (objectB && objectB.typename === 'TextFrame') {
                objectC = createCFromTextFrame(objectB);
            }
            if (!objectC) {
                // 通常ケース：Bを複製してC
                objectC = objectB.duplicate();
            }

            // Cを上下左右にオフセット（外側へ拡張）
            safeDo(function () {
                objectC.left -= offsetValuePt;
                objectC.top += offsetValuePt;
                objectC.width += offsetValuePt * 2;
                objectC.height += offsetValuePt * 2;
            });

            // Cの領域（後で、この領域内にあるパスを削除する）
            var __cBounds = null; // [L, T, R, B]
            safeDo(function () { __cBounds = objectC.geometricBounds; });

            // 2. Bは隠す
            objectB.hidden = true;

            // 3. AとCに対して設定：塗り=K30、線=なし
            function applyFillK30NoStroke(item) {
                if (!item) return;

                // K30（CMYK）
                var k30 = new CMYKColor();
                k30.cyan = 0;
                k30.magenta = 0;
                k30.yellow = 0;
                k30.black = 30;

                function applyToPath(p) {
                    safeDo(function () {
                        p.filled = true;
                        p.fillColor = k30;
                        p.stroked = false;
                    });
                }

                try {
                    var t = item.typename;
                    if (t === 'PathItem') {
                        applyToPath(item);
                        return;
                    }
                    if (t === 'CompoundPathItem') {
                        for (var i = 0; i < item.pathItems.length; i++) {
                            applyToPath(item.pathItems[i]);
                        }
                        return;
                    }
                    if (t === 'GroupItem') {
                        // 直下のパス
                        for (var j = 0; j < item.pathItems.length; j++) {
                            applyToPath(item.pathItems[j]);
                        }
                        // 直下の複合パス
                        for (var k = 0; k < item.compoundPathItems.length; k++) {
                            applyFillK30NoStroke(item.compoundPathItems[k]);
                        }
                        // ネストしたグループ
                        for (var g = 0; g < item.groupItems.length; g++) {
                            applyFillK30NoStroke(item.groupItems[g]);
                        }
                        return;
                    }
                    if (t === 'TextFrame') {
                        // 念のため（今回の用途では想定外だが）
                        safeDo(function () {
                            item.textRange.characterAttributes.fillColor = k30;
                            item.textRange.characterAttributes.strokeColor = k30;
                            item.textRange.characterAttributes.strokeWeight = 0;
                        });
                        return;
                    }
                } catch (e) { }
            }

            applyFillK30NoStroke(objectA);
            applyFillK30NoStroke(objectC);

            // 3. AとCをグループ化 (D)
            // 新しいグループアイテムを作成
            var groupD = doc.groupItems.add();

            // AとCをグループ内に移動
            // move関数は move(target, elementPlacement)
            objectA.move(groupD, ElementPlacement.PLACEATEND);
            objectC.move(groupD, ElementPlacement.PLACEATBEGINNING);

            // グループDを選択状態にする（executeMenuCommandは選択オブジェクトに対して実行されるため）
            doc.selection = null; // 一旦選択解除
            groupD.selected = true;

            // 4. Dに次を実行し、パスを分割
            // Live Pathfinder Divide (パスファインダー：分割) を適用
            // 注意: 'Live Pathfinder Divide' は効果としてのパスファインダーです。
            // 通常のパネル操作の「分割」挙動に近いのは 'Live Pathfinder Divide' 後に 'expandStyle' するか、
            // あるいは直接 'group' に対してパスファインダーパネルの分割コマンドを送ることです。

            // 方法1: ライブ効果を適用して拡張する（ご要望の手順）
            app.executeMenuCommand('Live Pathfinder Outline');
            app.executeMenuCommand('expandStyle');

            // 5. 生成結果（選択中）をグループ解除し、線幅1px・黒を適用
            function hasSelectedGroup() {
                var s = doc.selection;
                if (!s || s.length === 0) return false;
                for (var i = 0; i < s.length; i++) {
                    if (s[i] && s[i].typename === 'GroupItem') return true;
                }
                return false;
            }

            // 可能な限りグループ解除（ネストも含めて）
            (function ungroupFully() {
                var safety = 0;
                while (hasSelectedGroup() && safety < 50) {
                    app.executeMenuCommand('ungroup');
                    safety++;
                }
            })();

            // 6. Cの領域にあるパスを削除（Cの内側に完全に収まるものを対象）
            function boundsInside(b, region, tol) {
                if (!b || !region) return false;
                tol = tol || 0;
                // b, region: [L, T, R, B]
                return (
                    b[0] >= region[0] - tol &&
                    b[1] <= region[1] + tol &&
                    b[2] <= region[2] + tol &&
                    b[3] >= region[3] - tol
                );
            }

            (function deleteInsideC() {
                if (!__cBounds) return;

                // 実行直後の選択（分割結果）を走査
                var s = doc.selection;
                if (!s || s.length === 0) return;

                // 選択を壊さないようにコピーしてから削除
                var items = [];
                for (var i = 0; i < s.length; i++) items.push(s[i]);

                // 少しだけ許容（境界線上の誤差吸収）
                var tol = 0.01;

                for (var j = 0; j < items.length; j++) {
                    var it = items[j];
                    if (!it) continue;
                    safeDo(function () {
                        // ほとんどの生成物は PathItem / CompoundPathItem だが、念のためtypenameは問わず bounds で判断
                        var b = it.geometricBounds;
                        if (boundsInside(b, __cBounds, tol)) {
                            it.remove();
                        }
                    });
                }

                // selectionを更新（削除されたものを除外）
                safeDo(function () {
                    doc.selection = null;
                    // 残っているものだけ再選択
                    for (var k = 0; k < items.length; k++) {
                        safeDo(function () {
                            if (items[k] && items[k].parent) items[k].selected = true;
                        });
                    }
                });
            })();

            // 黒の作成（ドキュメントのカラースペースに合わせる）
            function makeBlackColor() {
                var isCMYK = false;
                safeDo(function () { isCMYK = (doc.documentColorSpace === DocumentColorSpace.CMYK); });
                if (isCMYK) {
                    var c = new CMYKColor();
                    c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100;
                    return c;
                }
                var r = new RGBColor();
                r.red = 0; r.green = 0; r.blue = 0;
                return r;
            }

            var __black = makeBlackColor();

            function applyStrokeBlack1px(item) {
                if (!item) return;

                function applyToPath(p) {
                    safeDo(function () {
                        p.stroked = true;
                        p.strokeColor = __black;
                        // Illustratorはpt基準。strokeWidthValuePt pt を採用
                        p.strokeWidth = strokeWidthValuePt;
                        p.strokeCap = strokeCapValue;
                        p.strokeJoin = strokeJoinValue;
                    });
                }

                var t = item.typename;
                if (t === 'PathItem') {
                    applyToPath(item);
                    return;
                }
                if (t === 'CompoundPathItem') {
                    for (var i = 0; i < item.pathItems.length; i++) applyToPath(item.pathItems[i]);
                    return;
                }
                if (t === 'GroupItem') {
                    for (var j = 0; j < item.pathItems.length; j++) applyToPath(item.pathItems[j]);
                    for (var k = 0; k < item.compoundPathItems.length; k++) applyStrokeBlack1px(item.compoundPathItems[k]);
                    for (var g = 0; g < item.groupItems.length; g++) applyStrokeBlack1px(item.groupItems[g]);
                    return;
                }
            }

            // 選択中の生成物に適用
            var __outSel = doc.selection;
            if (__outSel && __outSel.length) {
                for (var __i = 0; __i < __outSel.length; __i++) {
                    applyStrokeBlack1px(__outSel[__i]);
                }
            }

        } finally {
            // すべて終了したら、Bを再表示
            safeDo(function () {
                if (__restoreB) __restoreB.hidden = false;
            });
        }

    })();

})();
