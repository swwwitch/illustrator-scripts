#target illustrator
#targetengine "ResetRotationEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの回転を水平（0°）に補正します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetRotation.md

### Overview

Corrects the rotation of the selected objects back to horizontal (0°).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetRotation.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResetRotation";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetRotation.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetRotation.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ダイアログの初期値 / Dialog defaults */
    var DEFAULT_SETTINGS = {
        targetText: true,         /* テキストを対象にする / Level text */
        targetImage: true,        /* 配置画像を対象にする / Level placed images */
        targetRectangle: true,    /* 長方形を対象にする / Level rectangles */
        useClipGroupHost: true,   /* クリップグループごと回転する（常に最上位のクリップ） / Rotate the topmost clipping group */
        resetTextScale: true,     /* 文字の水平・垂直比率を100%に戻す / Reset character scale to 100% */
        epsilonDeg: 0.1           /* 水平とみなす角度（度） / Angle treated as level (degrees) */
    };

    /* 「水平とみなす範囲」の入力範囲（度） / Accepted range of the level tolerance (degrees) */
    var MIN_EPSILON_DEG = 0.01;
    var MAX_EPSILON_DEG = 10;

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

    var EPSILON_FIELD_CHARACTERS = 6;                 /* 「水平とみなす範囲」欄の幅 */

    /**
     * 縦並びのパネルを追加する（チェックボックスが並ぶので間隔を詰める）
     * @param {Window} parent - 追加先
     * @param {string} labelPath - パネル見出しのラベルパス
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelPath) {
        var panel = parent.add("panel", undefined, getLabel(labelPath));
        setupPanel(panel, 6);
        return panel;
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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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
            title: { ja: "水平補正", en: "Level Objects" }
        },
        panel: {
            targets: { ja: "対象", en: "Objects to Level" },
            text: { ja: "テキスト", en: "Text" },
            options: { ja: "補正条件", en: "Correction Options" }
        },
        checkbox: {
            text: { ja: "テキスト", en: "Text" },
            image: { ja: "配置画像（埋め込み・リンク）", en: "Placed Image (Embedded/Linked)" },
            rectangle: { ja: "長方形（パス）", en: "Rectangle (Path)" },
            clipGroup: { ja: "クリップグループ", en: "Clipping Group" },
            resetTextScale: { ja: "縦横比を正す", en: "Reset Character Scale" }
        },
        fieldLabel: {
            epsilon: { ja: "水平とみなす範囲", en: "Level Tolerance" }
        },
        tooltip: {
            text: { ja: "テキストオブジェクトの回転を元に戻します。", en: "Clears the rotation on text objects." },
            image: { ja: "配置画像の回転を元に戻します。", en: "Clears the rotation on placed images." },
            rectangle: { ja: "長方形の回転を元に戻します。", en: "Clears the rotation on rectangles." },
            clipGroup: {
                ja: "クリップグループ内のオブジェクトは、最上位のクリップグループごと回転を戻します。",
                en: "Objects inside clipping groups are leveled by rotating the topmost clipping group."
            },
            resetTextScale: {
                ja: "テキストの文字の水平比率・垂直比率を100%に戻します。",
                en: "Resets the horizontal and vertical character scale of text to 100%."
            },
            epsilon: {
                ja: "これ以下の角度は0とみなします。わずかな傾きを無視するための値です（0.01〜10）。",
                en: "Angles below this count as zero, so tiny tilts are ignored (0.01–10)."
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
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." }
        }
    };

    // =========================================
    // 実行時の設定 / Runtime options
    // =========================================

    /* ダイアログで確定した補正条件（確定までは初期値） / Options confirmed in the dialog (defaults until then) */
    var levelOptions = {
        useClipGroupHost: DEFAULT_SETTINGS.useClipGroupHost,
        resetTextScale: DEFAULT_SETTINGS.resetTextScale,
        epsilonDeg: DEFAULT_SETTINGS.epsilonDeg
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

    /* 入力された単位を欄の単位へ換算するための、1単位あたりのポイント数（値は UNITS 表と同じ。キーは小文字）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Points per unit for converting typed units into the field's unit (same values as the UNITS table; lowercase keys) */
    var STEPPER_POINTS_PER_UNIT = {
        "in": 72, "inch": 72, "mm": 72 / 25.4, "cm": 72 / 2.54, "m": 72 / 25.4 * 1000,
        "pt": 1, "px": 1, "p": 12, "pc": 12, "pica": 12,
        "q": 72 / 25.4 * 0.25, "h": 72 / 25.4 * 0.25, "ft": 72 * 12, "yd": 72 * 36
    };

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
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

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
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
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
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
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) return;
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
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
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldPointsPerUnit = STEPPER_POINTS_PER_UNIT[fieldUnitKey];
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var pointsPerUnit = STEPPER_POINTS_PER_UNIT[unitKey];
            if (pointsPerUnit === undefined || fieldPointsPerUnit === undefined) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = value * pointsPerUnit;
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldPointsPerUnit;
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
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
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラベルと tooltip つきのチェックボックスを追加する
     * @param {Panel} parent - 追加先
     * @param {string} labelKey - checkbox / tooltip 共通のキー
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parent, labelKey, initialValue) {
        var optionCheckbox = parent.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        optionCheckbox.helpTip = getLabel("tooltip." + labelKey);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * 「水平とみなす範囲」の入力値を読み、範囲内に収める
     * @param {string} inputText - 入力された文字列
     * @returns {number} 角度（度）。数値でなければ初期値
     */
    function parseEpsilon(inputText) {
        var value = parseFloat(inputText);
        if (isNaN(value)) return DEFAULT_SETTINGS.epsilonDeg;
        return Math.max(MIN_EPSILON_DEG, Math.min(MAX_EPSILON_DEG, value));
    }

    /**
     * 単位の付いていない数値だけの入力に、欄の単位を書き足す
     * @param {EditText} numberInput - 対象の入力欄
     * @param {string} unitSuffix - 欄の単位（例 "°"）
     * @returns {void}
     */
    function appendUnitToPlainNumber(numberInput, unitSuffix) {
        var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
        if (/^[-+]?(\d+\.?\d*|\.\d+)$/.test(trimmedText)) numberInput.text = trimmedText + unitSuffix;
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

    /**
     * 対象と補正条件を選ぶダイアログを表示する
     * @returns {{targets: {text: boolean, image: boolean, rectangle: boolean}, useClipGroupHost: boolean, resetTextScale: boolean, epsilonDeg: number}|null} 選んだ内容（キャンセル時は null）
     */
    function showLevelDialog() {
        var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dlg);

        var targetsPanel = addPanel(dlg, "panel.targets");
        var textCheckbox = addCheckbox(targetsPanel, "text", DEFAULT_SETTINGS.targetText);
        var imageCheckbox = addCheckbox(targetsPanel, "image", DEFAULT_SETTINGS.targetImage);
        var rectangleCheckbox = addCheckbox(targetsPanel, "rectangle", DEFAULT_SETTINGS.targetRectangle);
        var clipGroupCheckbox = addCheckbox(targetsPanel, "clipGroup", DEFAULT_SETTINGS.useClipGroupHost);

        var textPanel = addPanel(dlg, "panel.text");
        var resetTextScaleCheckbox = addCheckbox(textPanel, "resetTextScale", DEFAULT_SETTINGS.resetTextScale);

        var optionsPanel = addPanel(dlg, "panel.options");
        var epsilonRow = optionsPanel.add("group");
        epsilonRow.orientation = "row";
        epsilonRow.alignChildren = "left";
        epsilonRow.add("statictext", undefined, labelText("fieldLabel.epsilon"));
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var epsilonStepperGroup = epsilonRow.add("group");
        epsilonStepperGroup.orientation = "row";
        epsilonStepperGroup.alignChildren = ["left", "center"];
        epsilonStepperGroup.spacing = 0;
        epsilonStepperGroup.margins = 0;
        var epsilonInput;
        /* 範囲は parseEpsilon() と同じ / same range as parseEpsilon() */
        var epsilonStepper = addStepper(epsilonStepperGroup, function () { return epsilonInput; }, { min: MIN_EPSILON_DEG, max: MAX_EPSILON_DEG, unit: "°" });
        epsilonInput = epsilonStepperGroup.add("edittext", undefined, String(DEFAULT_SETTINGS.epsilonDeg) + "°");
        epsilonInput.helpTip = getLabel("tooltip.epsilon");
        epsilonInput.characters = EPSILON_FIELD_CHARACTERS;
        bindSteppedArrowKeys(epsilonInput, epsilonStepper);
        /* 数値だけで確定したら単位を書き足す / append the unit to a bare number on commit */
        epsilonInput.addEventListener("change", function () { appendUnitToPlainNumber(epsilonInput, "°"); });

        /* ボタン行 / Button row */
        var buttonRow = addButtonRow(dlg);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        dlg.layout.layout(true);
        prepareDialogWindow(dlg, SCRIPT_NAME);
        if (dlg.show() !== 1) return null;

        return {
            targets: {
                text: textCheckbox.value,
                image: imageCheckbox.value,
                rectangle: rectangleCheckbox.value
            },
            useClipGroupHost: clipGroupCheckbox.value,
            resetTextScale: resetTextScaleCheckbox.value,
            epsilonDeg: parseEpsilon(epsilonInput.text)
        };
    }

    // =========================================
    // 対象の判定 / Target detection
    // =========================================

    /**
     * 配置画像（リンク）か
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} PlacedItem なら true
     */
    function isPlacedImage(item) {
        return item.typename === "PlacedItem";
    }

    /**
     * 埋め込み画像か
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} RasterItem なら true
     */
    function isEmbeddedImage(item) {
        return item.typename === "RasterItem";
    }

    /**
     * 単純な長方形（閉じた4点のパス）か
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} 閉じた4点の PathItem なら true
     */
    function isRectangle(item) {
        return item.typename === "PathItem" && item.closed && item.pathPoints.length === 4;
    }

    /**
     * 縦書きのテキストか
     * @param {TextFrame} textFrame - 判定するテキスト
     * @returns {boolean} 縦書きなら true
     */
    function isVertical(textFrame) {
        /* 空のテキストでは textRanges[0] が取れないことがある / textRanges[0] may fail on empty text */
        try {
            return textFrame.textRanges[0].characterAttributes.orientation === TextOrientation.VERTICAL;
        } catch (e) {
            return false;
        }
    }

    /**
     * オブジェクトごとの識別キーを返す（同じホストを二度回さないため）
     * @param {PageItem} item - 対象
     * @returns {string} 識別キー
     */
    function getItemKey(item) {
        /* uuid を持たない環境・オブジェクトに備える / Guard against items without uuid */
        try {
            if (item.uuid) return "u:" + item.uuid;
        } catch (e) {}
        try {
            if (item.index !== undefined) return item.toString() + "|" + item.index;
        } catch (e2) {}
        return "" + item;
    }

    /**
     * 親をたどり、最上位のクリップグループを返す
     * @param {PageItem} item - 起点のオブジェクト
     * @returns {PageItem} 最上位のクリップグループ（無ければ item 自身）
     */
    function getTopmostClippingGroup(item) {
        var host = null;
        /* 親が Layer・Document まで続くので、途中で読めないプロパティに備える / Parents run up to Layer/Document */
        try {
            var parentItem = item.parent;
            while (parentItem) {
                if (parentItem.typename === "GroupItem" && parentItem.clipped) host = parentItem;
                parentItem = parentItem.parent;
            }
        } catch (e) {}
        return host || item;
    }

    /* 回転の対象（ホスト）の解決結果 / Resolved rotation hosts */
    var hostCache = {};

    /**
     * 実際に回転させるオブジェクト（ホスト）を返す
     * クリップグループ設定がONならクリップグループ、OFFならオブジェクト自身
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {PageItem} 回転させるオブジェクト
     */
    function resolveHost(item) {
        if (!levelOptions.useClipGroupHost) return item;
        var itemKey = getItemKey(item);
        if (!hostCache[itemKey]) {
            hostCache[itemKey] = getTopmostClippingGroup(item);
        }
        return hostCache[itemKey];
    }

    /**
     * 選択を再帰的にたどり、種類別に対象を集める
     * @param {Object} items - PageItem のコレクションまたは配列
     * @param {{text: boolean, image: boolean, rectangle: boolean}} targets - 対象の種類
     * @param {{texts: TextFrame[], images: PageItem[], rectangles: PathItem[]}} buckets - 集めた対象の格納先
     * @returns {void}
     */
    function collectTargets(items, targets, buckets) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "GroupItem") {
                collectTargets(item.pageItems, targets, buckets);
            } else if (item.typename === "CompoundPathItem") {
                collectTargets(item.pathItems, targets, buckets);
            } else if (item.pageItems) {
                /* その他の入れ子（クリップグループ等） / Other containers */
                try {
                    collectTargets(item.pageItems, targets, buckets);
                } catch (e) {}
            }

            if (targets.text && item.typename === "TextFrame") buckets.texts.push(item);
            if (targets.image && (isPlacedImage(item) || isEmbeddedImage(item))) buckets.images.push(item);
            /* クリップグループモードではマスクパスを個別に回さない / Mask paths are not leveled on their own in clip-group mode */
            if (targets.rectangle && isRectangle(item) && !(levelOptions.useClipGroupHost && item.clipping)) {
                buckets.rectangles.push(item);
            }
        }
    }

    // =========================================
    // 角度の計算 / Angle computation
    // =========================================

    /**
     * ほぼ水平（0° または 360° 付近）か
     * @param {number} deg - 角度（度）
     * @returns {boolean} しきい値未満なら true
     */
    function isNearlyLevel(deg) {
        var absDeg = Math.abs(deg);
        var eps = levelOptions.epsilonDeg;
        return (absDeg < eps) || (Math.abs(absDeg - 360) < eps);
    }

    /**
     * 角度を (-180, 180] に正規化する
     * @param {number} deg - 角度（度）
     * @returns {number} 正規化した角度
     */
    function normalizeAngle(deg) {
        var angle = deg % 360;
        if (angle <= -180) angle += 360;
        if (angle > 180) angle -= 360;
        return angle;
    }

    /**
     * グループの最初の子で条件に合うものを返す
     * @param {GroupItem} group - グループ
     * @param {function(PageItem): boolean} predicate - 条件
     * @returns {PageItem|null} 見つかった子（無ければ null）
     */
    function findChild(group, predicate) {
        for (var i = 0; i < group.pageItems.length; i++) {
            if (predicate(group.pageItems[i])) return group.pageItems[i];
        }
        return null;
    }

    /**
     * クリップグループの見かけの角度を子から推定する
     * 優先順：クリッピングパス → 配置／埋め込み画像 → 通常のパス → 入れ子のグループ
     * @param {GroupItem} group - グループ
     * @returns {number} 推定した角度（度）。推定できなければ 0
     */
    function getGroupProxyAngle(group) {
        /* 子の読み取りに失敗したら 0° とみなす / Treat unreadable children as 0° */
        try {
            var referenceItem =
                findChild(group, function (child) { return child.typename === "PathItem" && child.clipping; }) ||
                findChild(group, function (child) { return child.typename === "PlacedItem" || child.typename === "RasterItem"; }) ||
                findChild(group, function (child) { return child.typename === "PathItem" && !child.clipping; });
            if (referenceItem) return getRotationDegrees(referenceItem);

            /* 入れ子のグループは 0 以外が出た時点で返す / Return the first non-zero angle from nested groups */
            for (var i = 0; i < group.pageItems.length; i++) {
                if (group.pageItems[i].typename === "GroupItem") {
                    var nestedAngle = getGroupProxyAngle(group.pageItems[i]);
                    if (nestedAngle !== 0) return nestedAngle;
                }
            }
        } catch (e) {}
        return 0;
    }

    /**
     * オブジェクトの変換行列を返す
     * @param {PageItem} item - 対象
     * @returns {Matrix|null} 行列（読めなければ null）
     */
    function getMatrix(item) {
        /* matrix を持たない種類（PathItem・GroupItem 等）がある / Some item types have no matrix */
        try {
            var matrix = item.matrix;
            return (matrix && matrix.mValueA !== undefined) ? matrix : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * オブジェクトの回転角を推定する
     * 行列 → パスの最初の辺 → グループの子、の順に試す
     * @param {PageItem} item - 対象
     * @returns {number} 回転角（度）
     */
    function getRotationDegrees(item) {
        var matrix = getMatrix(item);
        if (matrix && matrix.mValueB !== undefined) {
            var deg = Math.atan2(matrix.mValueB, matrix.mValueA) * 180 / Math.PI;
            /* グループの見かけは子の回転に依存することがある / A group's apparent angle may come from its children */
            if (item.typename === "GroupItem" && isNearlyLevel(deg)) {
                var proxyAngle = getGroupProxyAngle(item);
                if (!isNearlyLevel(proxyAngle)) return proxyAngle;
            }
            return deg;
        }

        /* パスは最初の辺（アンカー0→1）の向きから推定 / Infer a path's angle from its first segment */
        if (item.typename === "PathItem" && item.pathPoints && item.pathPoints.length >= 2) {
            try {
                var p0 = item.pathPoints[0].anchor;
                var p1 = item.pathPoints[1].anchor;
                return Math.atan2(p1[1] - p0[1], p1[0] - p0[0]) * 180 / Math.PI;
            } catch (e) {}
        }

        if (item.typename === "GroupItem") {
            var groupAngle = getGroupProxyAngle(item);
            if (!isNearlyLevel(groupAngle)) return groupAngle;
        }

        /* 推定できない種類は 0° とみなす / Assume 0° for anything else */
        return 0;
    }

    /**
     * 鏡像（行列式が負）か
     * @param {PageItem} item - 対象
     * @returns {boolean} 反転していれば true
     */
    function isMirroredTransform(item) {
        var matrix = getMatrix(item);
        if (!matrix) return false;
        return (matrix.mValueA * matrix.mValueD) - (matrix.mValueB * matrix.mValueC) < 0;
    }

    /**
     * 水平に戻すための回転量を返す
     * @param {PageItem} host - 回転させるオブジェクト
     * @returns {number|null} 回転量（度）。すでに水平なら null
     */
    function getLevelingDelta(host) {
        var normalized = normalizeAngle(getRotationDegrees(host));
        if (isNearlyLevel(normalized)) return null;
        /* 鏡像は回転方向が逆になる / Mirrored items rotate the other way */
        return isMirroredTransform(host) ? normalized : -normalized;
    }

    // =========================================
    // 補正 / Leveling
    // =========================================

    /**
     * オブジェクトを中心で回転し、バウンディングボックスをリセットする
     * @param {PageItem} item - 回転させるオブジェクト
     * @param {number} deg - 回転量（度）
     * @returns {void}
     */
    function rotateBy(item, deg) {
        var doc = app.activeDocument;
        var previousSelection = doc.selection;

        /* ロック・非表示のオブジェクトは回せない / Locked or hidden items cannot be rotated */
        try {
            item.rotate(deg, true, true, true, true, Transformation.CENTER);
        } catch (e) {
            return;
        }

        /* 対象だけを選んでバウンディングボックスをリセット / Select only the item and reset its bounding box */
        try {
            doc.selection = null;
            doc.selection = [item];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (e2) {}

        /* 元の選択に戻す（消えたオブジェクトがあると例外） / Restore the selection (throws if an item is gone) */
        try {
            if (previousSelection) doc.selection = previousSelection;
        } catch (e3) {}
    }

    /**
     * 対象ごとにホストを解決し、同じホストは一度だけ水平に戻す
     * @param {PageItem[]} items - 対象
     * @param {function(PageItem, PageItem): boolean} [prepareItem] - 回転前の処理。false を返すとその対象を飛ばす
     * @returns {number} 回転させたホストの数
     */
    function levelItems(items, prepareItem) {
        var rotatedCount = 0;
        var seenHosts = {};
        for (var i = 0; i < items.length; i++) {
            var host = resolveHost(items[i]);
            if (prepareItem && !prepareItem(items[i], host)) continue;

            var hostKey = getItemKey(host);
            if (seenHosts[hostKey]) continue;
            seenHosts[hostKey] = true;

            var delta = getLevelingDelta(host);
            if (delta === null) continue;
            rotateBy(host, delta);
            rotatedCount++;
        }
        return rotatedCount;
    }

    /**
     * テキストの回転前処理：クリップグループ外の縦書きは飛ばし、必要なら文字比率を戻す
     * @param {TextFrame} textFrame - テキスト
     * @param {PageItem} host - 回転させるオブジェクト
     * @returns {boolean} 処理を続けるなら true
     */
    function prepareTextFrame(textFrame, host) {
        if (host === textFrame && isVertical(textFrame)) return false;
        if (levelOptions.resetTextScale) resetTextScaling(textFrame);
        return true;
    }

    /**
     * テキストの文字の水平・垂直比率を100%に戻す
     * @param {TextFrame} textFrame - テキスト
     * @returns {void}
     */
    function resetTextScaling(textFrame) {
        /* 空のテキストやロック中は書き込めないことがある / May fail on empty or locked text */
        try {
            var charAttributes = textFrame.textRange.characterAttributes;
            charAttributes.horizontalScale = 100;
            charAttributes.verticalScale = 100;
        } catch (e) {}
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで条件を選び、選択オブジェクトの回転を水平に戻す
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var dialogResult = showLevelDialog();
        if (!dialogResult) return;
        levelOptions.useClipGroupHost = dialogResult.useClipGroupHost;
        levelOptions.resetTextScale = dialogResult.resetTextScale;
        levelOptions.epsilonDeg = dialogResult.epsilonDeg;
        hostCache = {};

        var buckets = { texts: [], images: [], rectangles: [] };
        collectTargets(doc.selection, dialogResult.targets, buckets);

        var rotatedCount = 0;
        if (buckets.texts.length) rotatedCount += levelItems(buckets.texts, prepareTextFrame);
        if (buckets.images.length) rotatedCount += levelItems(buckets.images);
        if (buckets.rectangles.length) rotatedCount += levelItems(buckets.rectangles);

        if (rotatedCount > 0) app.redraw();
    }

    main();

})();
