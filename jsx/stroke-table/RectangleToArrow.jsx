#target illustrator
#targetengine "RectangleToArrowEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中の長方形を、その範囲いっぱいの矢印に置き換えます。横長は右向き、縦長は上向きで、種類のボタンを option（Alt）＋クリックすると逆向き、⌘＋option（Ctrl＋Alt）＋クリックすると両矢印になります。
矢印は塗りのほか、軸と矢じりの線の組み合わせ（線端なし／丸型線端）も選べ、太さは常にプレビューで確認しながら調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArrow.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n789072361c12

### Overview

Turns each selected rectangle into an arrow that fills its bounds. Wide rectangles point right, tall ones point up; Option-click (Alt-click) a type button to reverse the direction, Command-Option-click (Ctrl-Alt-click) for a double-headed arrow.
Choose a filled arrow or one built from a stroked shaft and head (butt or round caps), and adjust the thickness with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangleToArrow.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RectangleToArrow";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-10";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RectangleToArrow.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RectangleToArrow.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n789072361c12"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================
    var DEFAULT_FILL_PERCENT = 45;  /* 塗りの矢印の軸の太さの初期値（短辺に対する％） / initial shaft thickness of the filled arrow (% of the short side) */
    var DEFAULT_LINE_PERCENT = 15;  /* 線の矢印の線幅の初期値（短辺に対する％） / initial stroke width of the stroked arrow (% of the short side) */
    var HEAD_RATIO = 1 / 2;         /* 矢じりの長さ（短辺に対する比率。1/2 で直角の矢じり） / head length relative to the short side (1/2 gives a right-angled head) */

    var TOLERANCE = 0.001;          /* 座標比較の許容値（pt） / tolerance for coordinate comparison (pt) */

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
            title: { ja: "長方形を矢印に", en: "Rectangle to Arrow" }
        },
        fieldLabel: {
            arrowType: { ja: "種類", en: "Type" },
            thickness: { ja: "太さ", en: "Thickness" }
        },
        button: {
            arrowType: {
                fill: { ja: "塗り", en: "Fill" },
                line: { ja: "線（線端なし）", en: "Stroke (Butt Cap)" },
                roundLine: { ja: "線（丸型線端）", en: "Stroke (Round Cap)" }
            },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            arrowType: {
                fill: { ja: "塗りの矢印にします。", en: "Makes a filled arrow." },
                line: {
                    ja: "軸の直線とくの字の矢じりを、線端なし・マイター結合の線で組み、［パスのアウトライン］と［パスファインダー（合体）］の効果で1つの形にします。",
                    en: "Builds the arrow from a straight shaft and a chevron head with butt caps and miter joins, merged with the Outline Stroke and Pathfinder (Add) effects."
                },
                roundLine: {
                    ja: "軸の直線とくの字の矢じりを、丸型線端・ラウンド結合の線で組み、［パスのアウトライン］と［パスファインダー（合体）］の効果で1つの形にします。",
                    en: "Builds the arrow from a straight shaft and a chevron head with round caps and round joins, merged with the Outline Stroke and Pathfinder (Add) effects."
                }
            },
            reverse: {
                ja: "option（Alt）＋クリックで矢印の向きを逆にします。",
                en: "Option-click (Alt-click) to reverse the arrow's direction."
            },
            bothEnds: {
                ja: "⌘＋option（Ctrl＋Alt）＋クリックで両矢印にします。",
                en: "Command-Option-click (Ctrl-Alt-click) for a double-headed arrow."
            },
            thickness: {
                ja: "塗りは軸の太さ、線は線幅を、長方形の短辺に対する％で指定します。種類ごとに値を覚えます。",
                en: "Shaft thickness (fill) or stroke width (stroke) as a percentage of the rectangle's short side. Each type keeps its own value."
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
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noRectangle: {
                ja: "長方形が選択されていません（回転した長方形・角丸は対象外です）。",
                en: "No rectangles are selected (rotated or rounded rectangles are not supported)."
            }
        }
    };

    // =========================================
    // レイアウト / Layout
    // =========================================
    var FIELD_CHARACTERS = 5;            /* 太さの入力欄の幅（文字数） / width of the thickness field (characters) */
    var LABEL_WIDTH = 40;                /* 項目名の幅（右揃え） / label width (right-aligned) */
    var TYPE_ICON_SIZE = [96, 32];       /* 種類のボタンの大きさ / arrow type button size */
    var TYPE_ICON_INSET = 5;             /* 種類のボタンの枠と絵の間（px） / gap between a type button's frame and its drawing */
    var TYPE_ICON_SPACING = 6;           /* 種類のボタンどうしの縦の間隔 / vertical spacing between type buttons */

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

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
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
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
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
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
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
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
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
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
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
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
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

    // =========================================
    // 長方形の判定 / Rectangle detection
    // =========================================

    /**
     * 2 つの値が許容値の範囲で等しいかを返す
     * @param {number} firstValue - 値
     * @param {number} secondValue - 値
     * @returns {boolean} 等しければ true
     */
    function isNear(firstValue, secondValue) {
        return Math.abs(firstValue - secondValue) < TOLERANCE;
    }

    /**
     * 2 つの点が許容値の範囲で同じ位置かを返す
     * @param {number[]} firstPoint - [x, y]
     * @param {number[]} secondPoint - [x, y]
     * @returns {boolean} 同じ位置なら true
     */
    function isSamePoint(firstPoint, secondPoint) {
        return isNear(firstPoint[0], secondPoint[0]) && isNear(firstPoint[1], secondPoint[1]);
    }

    /**
     * 回転していない長方形のパスかを返す（閉じた 4 点・ハンドルなし・各辺が水平か垂直）
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 長方形なら true
     */
    function isAxisAlignedRectangle(pageItem) {
        if (pageItem.typename !== "PathItem" || !pageItem.closed) return false;
        var pathPoints = pageItem.pathPoints;
        if (pathPoints.length !== 4) return false;
        for (var i = 0; i < 4; i++) {
            var anchor = pathPoints[i].anchor;
            var nextAnchor = pathPoints[(i + 1) % 4].anchor;
            if (!isSamePoint(pathPoints[i].leftDirection, anchor) || !isSamePoint(pathPoints[i].rightDirection, anchor)) return false;
            if (!isNear(anchor[0], nextAnchor[0]) && !isNear(anchor[1], nextAnchor[1])) return false;
        }
        return true;
    }

    /**
     * 選択からグループの中まで含めて長方形を集める
     * @param {Array|Object} pageItems - 選択または GroupItem.pageItems
     * @param {PathItem[]} foundRectangles - 見つかった長方形を追加する配列
     * @returns {PathItem[]} foundRectangles
     */
    function collectRectangles(pageItems, foundRectangles) {
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.locked || pageItem.hidden) continue;
            if (pageItem.typename === "GroupItem") {
                collectRectangles(pageItem.pageItems, foundRectangles);
            } else if (isAxisAlignedRectangle(pageItem)) {
                foundRectangles.push(pageItem);
            }
        }
        return foundRectangles;
    }

    // =========================================
    // 矢印の形 / Arrow shape
    // =========================================

    /**
     * 長方形の範囲に、矢印の向きに沿った座標系を作る（横長は右向き、縦長は上向き。逆向きなら左向き・下向き）。
     * u は根元から先端へ向かう距離、v は軸からの横のずれ
     * @param {Array} bounds - geometricBounds [左, 上, 右, 下]
     * @param {boolean} isReversed - 逆向きにするなら true
     * @returns {{length: number, thickness: number, toPoint: function(number, number): Array}} 長さ・短辺・座標の変換
     */
    function getArrowFrame(bounds, isReversed) {
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        var width = right - left;
        var height = top - bottom;

        if (width >= height) {
            var centerY = (top + bottom) / 2;
            return {
                length: width,
                thickness: height,
                toPoint: function (u, v) { return [isReversed ? right - u : left + u, centerY + v]; }
            };
        }

        var centerX = (left + right) / 2;
        return {
            length: height,
            thickness: width,
            toPoint: function (u, v) { return [centerX - v, isReversed ? top - u : bottom + u]; }
        };
    }

    /**
     * 矢じりの長さを返す（長方形が短すぎるときは、片矢印は全長、両矢印は半分で止める）
     * @param {Object} arrowFrame - getArrowFrame() の戻り値
     * @param {boolean} isBothEnds - 両矢印なら true
     * @returns {number} 矢じりの長さ
     */
    function getHeadLength(arrowFrame, isBothEnds) {
        return Math.min(arrowFrame.thickness * HEAD_RATIO, isBothEnds ? arrowFrame.length / 2 : arrowFrame.length);
    }

    /**
     * 塗りの矢印の頂点を求める
     * @param {Object} arrowFrame - getArrowFrame() の戻り値
     * @param {number} shaftRatio - 軸の太さ（短辺に対する比率）
     * @param {boolean} isBothEnds - 両矢印なら true
     * @returns {Array} 頂点 [[x, y], …] 片矢印は 7 点、両矢印は 10 点
     */
    function buildFilledArrowPoints(arrowFrame, shaftRatio, isBothEnds) {
        var halfSide = arrowFrame.thickness / 2;
        var halfShaft = arrowFrame.thickness * shaftRatio / 2;
        var headLength = getHeadLength(arrowFrame, isBothEnds);
        var neck = arrowFrame.length - headLength;
        var tailPoints = isBothEnds
            ? [arrowFrame.toPoint(headLength, -halfShaft), arrowFrame.toPoint(headLength, -halfSide), arrowFrame.toPoint(0, 0),
               arrowFrame.toPoint(headLength, halfSide), arrowFrame.toPoint(headLength, halfShaft)]
            : [arrowFrame.toPoint(0, -halfShaft), arrowFrame.toPoint(0, halfShaft)];
        return [
            arrowFrame.toPoint(neck, halfShaft),
            arrowFrame.toPoint(neck, halfSide),
            arrowFrame.toPoint(arrowFrame.length, 0),
            arrowFrame.toPoint(neck, -halfSide),
            arrowFrame.toPoint(neck, -halfShaft)
        ].concat(tailPoints);
    }

    /**
     * 線の矢印の折れ線を求める（軸・先端の矢じり・両矢印なら根元の矢じり）
     * @param {Object} arrowFrame - getArrowFrame() の戻り値
     * @param {boolean} isBothEnds - 両矢印なら true
     * @returns {Array} 折れ線の配列 [[[x, y], …], …]
     */
    function buildStrokedArrowLines(arrowFrame, isBothEnds) {
        var halfSide = arrowFrame.thickness / 2;
        var headLength = getHeadLength(arrowFrame, isBothEnds);
        var neck = arrowFrame.length - headLength;
        var arrowLines = [
            /* 軸は先端まで通す。合体するので継ぎ目は残らない / the shaft runs to the tip; uniting hides the joint */
            [arrowFrame.toPoint(0, 0), arrowFrame.toPoint(arrowFrame.length, 0)],
            [arrowFrame.toPoint(neck, halfSide), arrowFrame.toPoint(arrowFrame.length, 0), arrowFrame.toPoint(neck, -halfSide)]
        ];
        if (isBothEnds) {
            arrowLines.push([arrowFrame.toPoint(headLength, halfSide), arrowFrame.toPoint(0, 0), arrowFrame.toPoint(headLength, -halfSide)]);
        }
        return arrowLines;
    }

    /**
     * 塗りまたは線に使う色を返す（塗り → 線 → 黒の順で探す）
     * @param {PathItem} rectanglePath - 元の長方形
     * @returns {Color} 色
     */
    function getArrowColor(rectanglePath) {
        if (rectanglePath.filled) return rectanglePath.fillColor;
        if (rectanglePath.stroked) return rectanglePath.strokeColor;
        var blackColor = new GrayColor();
        blackColor.gray = 100;
        return blackColor;
    }

    /**
     * パスを開いた折れ線に書き換え、線の矢印の見た目にする（塗りなし・実線・線端と角の形）
     * @param {PathItem} linePath - 書き換えるパス
     * @param {Array} linePoints - 頂点 [[x, y], …]
     * @param {Color} lineColor - 線の色
     * @param {number} strokeWidth - 線幅（pt）
     * @param {boolean} isRound - 丸型線端・ラウンド結合にするなら true
     * @returns {void}
     */
    function setArrowLine(linePath, linePoints, lineColor, strokeWidth, isRound) {
        linePath.setEntirePath(linePoints);
        linePath.closed = false;
        linePath.filled = false;
        linePath.stroked = true;
        linePath.strokeColor = lineColor;
        linePath.strokeWidth = strokeWidth;
        linePath.strokeDashes = [];
        linePath.strokeCap = isRound ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
        linePath.strokeJoin = isRound ? StrokeJoin.ROUNDENDJOIN : StrokeJoin.MITERENDJOIN;
        /* 直角の矢じりが面取りされないように / keep the right-angled head from beveling */
        linePath.strokeMiterLimit = 4;
    }

    /**
     * 長方形を塗りの矢印に書き換える。元のパスを書き換えるので、効果・重ね順はそのまま残る
     * @param {PathItem} rectanglePath - 長方形のパス
     * @param {Object} arrowOptions - shaftRatio（軸の太さ。短辺に対する比率）、isReversed、isBothEnds
     * @returns {PageItem} できた矢印
     */
    function convertToFilledArrow(rectanglePath, arrowOptions) {
        var arrowColor = getArrowColor(rectanglePath);
        var arrowFrame = getArrowFrame(rectanglePath.geometricBounds, arrowOptions.isReversed);
        rectanglePath.setEntirePath(buildFilledArrowPoints(arrowFrame, arrowOptions.shaftRatio, arrowOptions.isBothEnds));
        rectanglePath.fillColor = arrowColor;
        rectanglePath.filled = true;
        rectanglePath.stroked = false;
        return rectanglePath;
    }

    /**
     * 長方形を、軸の直線とくの字の矢じりを線で組んだ矢印に置き換える。
     * 線ごとに［パスのアウトライン］効果を掛けてグループにし、グループに［パスファインダー（合体）］効果を掛ける。
     * パスの形は長方形の範囲に合わせ、線幅のぶんは外へはみ出す
     * @param {PathItem} rectanglePath - 長方形のパス（グループに置き換えて削除する）
     * @param {Object} arrowOptions - shaftRatio（線幅。短辺に対する比率）、isReversed、isBothEnds
     * @param {boolean} isRound - 丸型線端・ラウンド結合にするなら true
     * @returns {PageItem} できた矢印（グループ）
     */
    function convertToStrokedArrow(rectanglePath, arrowOptions, isRound) {
        var arrowFrame = getArrowFrame(rectanglePath.geometricBounds, arrowOptions.isReversed);
        var arrowColor = getArrowColor(rectanglePath);
        var strokeWidth = arrowFrame.thickness * arrowOptions.shaftRatio;
        var arrowLines = buildStrokedArrowLines(arrowFrame, arrowOptions.isBothEnds);

        /* 入れ子のグループで効果のメニューコマンドを実行すると一番外側のグループに付くので、
           部品はレイヤー直下で作って効果を掛け、最後に元の位置へ移す
           Effect menu commands inside nested groups land on the outermost group, so build at the layer's top level and move in at the end */
        var targetLayer = rectanglePath.layer;
        var linePaths = [];
        for (var i = 0; i < arrowLines.length; i++) {
            var linePath = rectanglePath.duplicate(targetLayer, ElementPlacement.PLACEATBEGINNING);
            setArrowLine(linePath, arrowLines[i], arrowColor, strokeWidth, isRound);
            linePaths.push(linePath);
        }

        /* ［パスのアウトライン］は XML（applyEffect）では付かないので、メニューコマンドで線ごとに掛ける
           Outline Stroke does not apply through XML, so use the menu command on each line */
        selectOnly(linePaths);
        app.executeMenuCommand("Live Outline Stroke");

        var arrowGroup = targetLayer.groupItems.add();
        for (var j = 0; j < linePaths.length; j++) linePaths[j].move(arrowGroup, ElementPlacement.PLACEATEND);
        selectOnly([arrowGroup]);
        app.executeMenuCommand("Live Pathfinder Add");

        arrowGroup.move(rectanglePath, ElementPlacement.PLACEBEFORE);
        rectanglePath.remove();
        return arrowGroup;
    }

    /**
     * 指定したオブジェクトだけを選択し、画面を更新する（効果のメニューコマンドの前に呼ぶ）
     * @param {PageItem[]} pageItems - 選択するオブジェクト
     * @returns {void}
     */
    function selectOnly(pageItems) {
        app.activeDocument.selection = null;
        for (var i = 0; i < pageItems.length; i++) pageItems[i].selected = true;
        app.redraw();
    }

    /**
     * 種類に応じて長方形を矢印にする
     * @param {PathItem} rectanglePath - 長方形のパス
     * @param {Object} arrowOptions - arrowType（"fill" / "line" / "roundLine"）、shaftRatio、isReversed、isBothEnds
     * @returns {PageItem} できた矢印
     */
    function convertToArrow(rectanglePath, arrowOptions) {
        if (arrowOptions.arrowType === "fill") return convertToFilledArrow(rectanglePath, arrowOptions);
        return convertToStrokedArrow(rectanglePath, arrowOptions, arrowOptions.arrowType === "roundLine");
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * 複製を矢印にして見せ、元の長方形は一時的に隠すプレビューを作る
     * @param {Array} rectangles - 対象の長方形
     * @returns {{show: function(Object): void, clear: function(): void}} プレビューの操作
     */
    function createArrowPreview(rectangles) {
        var previewItems = [];

        function clear() {
            for (var i = 0; i < previewItems.length; i++) previewItems[i].remove();
            /* 対象は表示中のものだけを集めているので、すべて表示に戻してよい / targets were visible when collected */
            for (var j = 0; j < rectangles.length; j++) rectangles[j].hidden = false;
            previewItems = [];
        }

        function show(arrowOptions) {
            clear();
            for (var i = 0; i < rectangles.length; i++) {
                /* 複製は hidden を引き継ぐので、元を隠す前に作る / duplicates inherit hidden, so duplicate before hiding */
                previewItems.push(convertToArrow(rectangles[i].duplicate(), arrowOptions));
                rectangles[i].hidden = true;
            }
        }

        return { show: show, clear: clear };
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

    // =========================================
    // 種類のボタン / Arrow type buttons
    // =========================================

    /* 種類のボタンの絵（80×20 の座標）。rects は塗る長方形、heads は三角形 [付け根x, 先端x, 中心y, 高さの半分]、
       lines は折れ線（最後の要素が線幅）、discs は丸型線端の円 [x, y]
       Drawings on an 80×20 grid: filled rects, triangle heads, polylines (last element is the width), round-cap discs */
    var ARROW_TYPE_ICON_DESIGN_SIZE = [80, 20];
    var ARROW_TYPE_ICON_LINE_WIDTH = 3;
    var ARROW_TYPE_ICON_SHAPES = {
        fill: { rects: [[1, 5.5, 69, 14.5]], heads: [[69, 79, 10, 10]], lines: [], discs: [] },
        line: { rects: [], heads: [], lines: [[[1, 10], [77, 10], 3], [[69.5, 2.5], [77, 10], [69.5, 17.5], 3]], discs: [] },
        roundLine: {
            rects: [], heads: [],
            lines: [[[2.5, 10], [77, 10], 3], [[69.5, 2.5], [77, 10], [69.5, 17.5], 3]],
            discs: [[2.5, 10], [69.5, 2.5], [69.5, 17.5], [77, 10]]
        },
        /* 両矢印 / double-headed */
        fillBothEnds: { rects: [[11, 5.5, 69, 14.5]], heads: [[69, 79, 10, 10], [11, 1, 10, 10]], lines: [], discs: [] },
        lineBothEnds: {
            rects: [], heads: [],
            lines: [[[3, 10], [77, 10], 3], [[69.5, 2.5], [77, 10], [69.5, 17.5], 3], [[10.5, 2.5], [3, 10], [10.5, 17.5], 3]],
            discs: []
        },
        roundLineBothEnds: {
            rects: [], heads: [],
            lines: [[[3, 10], [77, 10], 3], [[69.5, 2.5], [77, 10], [69.5, 17.5], 3], [[10.5, 2.5], [3, 10], [10.5, 17.5], 3]],
            discs: [[69.5, 2.5], [69.5, 17.5], [77, 10], [10.5, 2.5], [10.5, 17.5], [3, 10]]
        }
    };
    var ARROW_TYPE_KEYS = ["fill", "line", "roundLine"];

    /**
     * ラジオボタンの代わりに、矢印の種類を絵で描いたボタンを縦に並べる（SmartStrokeSettings のよく使う矢印と同じ見た目）。
     * 向き（逆向き・両矢印かどうか）は組で1つの typeState に持ち、絵も向きに合わせて描く
     * @param {Group} parent - 追加先
     * @param {string} selectedKey - 初期選択の種類
     * @param {function(): void} onSelect - 選び直したとき・向きを変えたときに呼ぶ
     * @returns {Group[]} 追加したボタン（.value・.arrowType・.typeState を持つ）
     */
    function addArrowTypeIcons(parent, selectedKey, onSelect) {
        var iconColumn = parent.add("group");
        iconColumn.orientation = "column";
        iconColumn.alignChildren = ["left", "top"];
        iconColumn.spacing = TYPE_ICON_SPACING;
        var typeIcons = [];
        var typeState = { isReversed: false, isBothEnds: false };
        for (var i = 0; i < ARROW_TYPE_KEYS.length; i++) {
            typeIcons.push(addArrowTypeIcon(iconColumn, ARROW_TYPE_KEYS[i], typeIcons, onSelect));
            typeIcons[i].typeState = typeState;
            typeIcons[i].value = (ARROW_TYPE_KEYS[i] === selectedKey);
        }
        return typeIcons;
    }

    /**
     * 種類のボタンを1つ作る。押すと同じ組のほかのボタンを外し、onSelect を呼ぶ。
     * option（Alt）＋クリックでは、そのボタンを選んだうえで矢印の向きを逆にし、
     * ⌘＋option（Ctrl＋Alt）＋クリックでは両矢印のオン／オフを切り替える
     * @param {Group} iconColumn - 追加先の列
     * @param {string} arrowType - 種類（"fill" / "line" / "roundLine"）
     * @param {Group[]} typeIcons - 同じ組のボタン（排他にする）
     * @param {function(): void} onSelect - 選び直したときに呼ぶ
     * @returns {Group} ボタン
     */
    function addArrowTypeIcon(iconColumn, arrowType, typeIcons, onSelect) {
        var typeIcon = iconColumn.add("group");
        typeIcon.preferredSize = TYPE_ICON_SIZE;
        typeIcon.minimumSize = TYPE_ICON_SIZE;
        typeIcon.maximumSize = TYPE_ICON_SIZE;
        typeIcon.helpTip = [
            getLabel("button.arrowType." + arrowType),
            getLabel("tooltip.arrowType." + arrowType),
            getLabel("tooltip.reverse"),
            getLabel("tooltip.bothEnds")
        ].join("\n");
        typeIcon.arrowType = arrowType;
        typeIcon.value = false;
        typeIcon.onDraw = function () { drawArrowTypeIcon(typeIcon); };
        typeIcon.addEventListener("mousedown", function (mouseEvent) {
            var isAltClick = !!(mouseEvent && mouseEvent.altKey);
            var isCommandAltClick = isAltClick && !!(mouseEvent.metaKey || mouseEvent.ctrlKey);
            if (typeIcon.value && !isAltClick) return;
            if (isCommandAltClick) typeIcon.typeState.isBothEnds = !typeIcon.typeState.isBothEnds;
            else if (isAltClick) typeIcon.typeState.isReversed = !typeIcon.typeState.isReversed;
            for (var i = 0; i < typeIcons.length; i++) {
                typeIcons[i].value = (typeIcons[i] === typeIcon);
                redrawStepperGroup(typeIcons[i]);
            }
            onSelect();
        });
        return typeIcon;
    }

    /**
     * 選ばれている種類を返す
     * @param {Group[]} typeIcons - addArrowTypeIcons() で作ったボタン
     * @returns {string} 種類。どれも選ばれていなければ先頭
     */
    function getCheckedArrowType(typeIcons) {
        for (var i = 0; i < typeIcons.length; i++) {
            if (typeIcons[i].value) return typeIcons[i].arrowType;
        }
        return typeIcons[0].arrowType;
    }

    /**
     * 種類のボタンを描く（選択中は灰色の地、そうでなければ白の地。ダークUIでは明暗を反転）
     * @param {Group} typeIcon - addArrowTypeIcon() で作ったボタン
     * @returns {void}
     */
    function drawArrowTypeIcon(typeIcon) {
        var iconGraphics = typeIcon.graphics;
        var isDark = isDarkUI();
        var isDimmed = !isStepperEnabledInTree(typeIcon);
        var inkLevel = isDark ? 0.65 : 0.45; /* 絵と枠は真っ黒・真っ白より抑える / keep the ink softer than pure black or white */
        var groundLevel = typeIcon.value ? (isDark ? 0.45 : 0.7) : (isDark ? 0.2 : 1);
        var inkColor = [inkLevel, inkLevel, inkLevel, isDimmed ? 0.4 : 1];
        var groundColor = [groundLevel, groundLevel, groundLevel, isDimmed ? 0.4 : 1];
        var iconWidth = TYPE_ICON_SIZE[0];
        var iconHeight = TYPE_ICON_SIZE[1];

        /* 地と枠 / ground and frame */
        iconGraphics.newPath();
        iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, groundColor));
        iconGraphics.newPath();
        iconGraphics.moveTo(0.5, 0.5);
        iconGraphics.lineTo(iconWidth - 0.5, 0.5);
        iconGraphics.lineTo(iconWidth - 0.5, iconHeight - 0.5);
        iconGraphics.lineTo(0.5, iconHeight - 0.5);
        iconGraphics.lineTo(0.5, 0.5);
        iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, 1));

        var typeState = typeIcon.typeState;
        drawArrowTypeShape(iconGraphics, ARROW_TYPE_ICON_SHAPES[typeIcon.arrowType + (typeState.isBothEnds ? "BothEnds" : "")],
            [TYPE_ICON_INSET, TYPE_ICON_INSET, iconWidth - TYPE_ICON_INSET * 2, iconHeight - TYPE_ICON_INSET * 2], inkColor, typeState.isReversed && !typeState.isBothEnds);
    }

    /**
     * 矢印の絵を、縦横比を保って範囲の中央に描く。
     * onDraw の多角形は塗れないので、三角形は細い長方形の並び、丸型線端は円で組む
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {Object} iconShape - ARROW_TYPE_ICON_SHAPES の要素
     * @param {number[]} drawArea - 描く範囲 [左, 上, 幅, 高さ]
     * @param {number[]} inkColor - [r, g, b, a]
     * @param {boolean} isReversed - 左右反転して描くなら true
     * @returns {void}
     */
    function drawArrowTypeShape(iconGraphics, iconShape, drawArea, inkColor, isReversed) {
        var designSize = ARROW_TYPE_ICON_DESIGN_SIZE;
        var iconScale = Math.min(drawArea[2] / designSize[0], drawArea[3] / designSize[1]);
        var originX = drawArea[0] + (drawArea[2] - designSize[0] * iconScale) / 2;
        var originY = drawArea[1] + (drawArea[3] - designSize[1] * iconScale) / 2;
        var i, k;

        /**
         * 絵の x 座標を画面の x 座標にする（逆向きなら左右反転）
         * @param {number} designX - 絵の座標系の x
         * @returns {number} 画面の x
         */
        function toScreenX(designX) {
            return originX + (isReversed ? designSize[0] - designX : designX) * iconScale;
        }

        /**
         * 絵の座標系の長方形をパスに足す（反転や左向きの矢じりでも左上・幅・高さが正になるようにする）
         * @param {number} designLeft - 左（右との大小は問わない）
         * @param {number} designTop - 上
         * @param {number} designRight - 右
         * @param {number} designBottom - 下
         * @returns {void}
         */
        function addDesignRect(designLeft, designTop, designRight, designBottom) {
            var screenLeft = Math.min(toScreenX(designLeft), toScreenX(designRight));
            iconGraphics.rectPath(screenLeft, originY + designTop * iconScale,
                Math.abs(designRight - designLeft) * iconScale, (designBottom - designTop) * iconScale);
        }

        iconGraphics.newPath();
        for (i = 0; i < iconShape.rects.length; i++) {
            var shapeRect = iconShape.rects[i];
            addDesignRect(shapeRect[0], shapeRect[1], shapeRect[2], shapeRect[3]);
        }
        for (i = 0; i < iconShape.heads.length; i++) {
            var headShape = iconShape.heads[i];
            var sliceCount = 12;
            var sliceWidth = (headShape[1] - headShape[0]) / sliceCount;
            for (k = 0; k < sliceCount; k++) {
                var sliceHalf = headShape[3] * (1 - (k + 0.5) / sliceCount);
                var sliceLeft = headShape[0] + sliceWidth * k;
                addDesignRect(sliceLeft, headShape[2] - sliceHalf, sliceLeft + sliceWidth, headShape[2] + sliceHalf);
            }
        }
        var discRadius = ARROW_TYPE_ICON_LINE_WIDTH * iconScale / 2;
        for (i = 0; i < iconShape.discs.length; i++) {
            var discCenter = iconShape.discs[i];
            iconGraphics.ellipsePath(toScreenX(discCenter[0]) - discRadius, originY + discCenter[1] * iconScale - discRadius,
                discRadius * 2, discRadius * 2);
        }
        iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, inkColor));

        for (i = 0; i < iconShape.lines.length; i++) {
            var shapeLine = iconShape.lines[i];
            /* 折れ線は1本ずつ newPath から描く（続けると前の線に積み重なる） / one path per polyline */
            iconGraphics.newPath();
            for (k = 0; k < shapeLine.length - 1; k++) {
                var pointX = toScreenX(shapeLine[k][0]);
                var pointY = originY + shapeLine[k][1] * iconScale;
                if (k === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, inkColor, shapeLine[shapeLine.length - 1] * iconScale));
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 矢印の種類と太さを指定するダイアログボックスを表示する
     * @param {Array} rectangles - 対象の長方形（プレビューに使う）
     * @returns {Object|null} arrowType・shaftRatio・isReversed・isBothEnds。キャンセルなら null
     */
    function showArrowDialog(rectangles) {
        var arrowPreview = createArrowPreview(rectangles);
        /* 種類ごとに太さを覚え、切り替えたら戻す / remember the thickness per type */
        var percentByType = { fill: DEFAULT_FILL_PERCENT, line: DEFAULT_LINE_PERCENT, roundLine: DEFAULT_LINE_PERCENT };
        var currentType = "fill";

        var arrowDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(arrowDialog);

        var typeRow = arrowDialog.add("group");
        setupRow(typeRow);
        var typeLabel = typeRow.add("statictext", undefined, labelText("fieldLabel.arrowType"));
        typeLabel.preferredSize.width = LABEL_WIDTH;
        typeLabel.justify = "right";
        typeLabel.alignment = ["left", "top"]; /* 縦に並んだボタンの先頭にそろえる / align with the first button */
        var typeIcons = addArrowTypeIcons(typeRow, currentType, switchType);
        var typeState = typeIcons[0].typeState;

        var thicknessInput = addSteppedField(arrowDialog, {
            label: labelText("fieldLabel.thickness"), labelWidth: LABEL_WIDTH,
            text: percentByType.fill + "%", characters: FIELD_CHARACTERS,
            step: 1, min: 1, max: 100, unit: "%",
            onStep: function () { refreshPreview(); }
        });
        thicknessInput.helpTip = getLabel("tooltip.thickness");

        /* 直接入力でも描き直す。部品の onChange（値の正規化）のあとに呼ぶ / redraw on typed input too, after the part's own normalizing onChange */
        var normalizeThicknessInput = thicknessInput.onChange;
        thicknessInput.onChange = function () {
            normalizeThicknessInput();
            refreshPreview();
        };

        /* ボタン行は左右中央に置く / center the button row */
        var buttonRow = addButtonRow(arrowDialog, { centered: true });
        buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 選ばれた種類に切り替え、その種類で覚えている太さを入力欄に戻す
         * @returns {void}
         */
        function switchType() {
            percentByType[currentType] = parseFloat(thicknessInput.text);
            currentType = getCheckedArrowType(typeIcons);
            thicknessInput.text = percentByType[currentType] + "%";
            thicknessInput.lastValidText = thicknessInput.text;
            refreshPreview();
        }

        /**
         * 今の指定を返す
         * @returns {Object} arrowType・shaftRatio（短辺に対する比率）・isReversed・isBothEnds
         */
        function getArrowOptions() {
            return {
                arrowType: currentType,
                shaftRatio: parseFloat(thicknessInput.text) / 100,
                isReversed: typeState.isReversed,
                isBothEnds: typeState.isBothEnds
            };
        }

        /**
         * プレビューを今の値で描き直す（プレビューは常にオン）
         * @returns {void}
         */
        function refreshPreview() {
            arrowPreview.show(getArrowOptions());
            app.redraw();
        }

        refreshPreview();
        thicknessInput.active = true;

        prepareDialogWindow(arrowDialog, SCRIPT_NAME);
        var dialogResult = arrowDialog.show();

        var arrowOptions = getArrowOptions();
        arrowPreview.clear();
        app.redraw();
        return (dialogResult === 1) ? arrowOptions : null;
    }

    // =========================================
    // メイン / Main
    // =========================================

    (function () {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var rectangles = collectRectangles(app.activeDocument.selection, []);
        if (rectangles.length === 0) {
            alert(getLabel("alert.noRectangle"));
            return;
        }

        var arrowOptions = showArrowDialog(rectangles);
        if (arrowOptions === null) return;

        /* 確定は元の長方形をそのまま使う / commit on the original rectangles */
        var arrows = [];
        for (var i = 0; i < rectangles.length; i++) {
            arrows.push(convertToArrow(rectangles[i], arrowOptions));
        }
        app.activeDocument.selection = arrows;
    })();

})();
