#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード名を解析し、行列グリッドとして再配置します。
名前全体が「行番号＋区切り文字（-, _, x）＋列番号」に一致するアートボードを、行列指定として扱います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridArrangeArtboards.md

### Overview

Parses the artboard names and re-lays the artboards out as a row-column grid.
An artboard counts as positioned when its whole name matches "row number + separator (-, _, x) + column number".

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridArrangeArtboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GridArrangeArtboards";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridArrangeArtboards.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridArrangeArtboards.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値（間隔はアクティブなアートボード幅の1/8） / Dialog defaults (the gap is 1/8 of the active artboard width) */
    var DIALOG_DEFAULTS = {
        exceptionMode: 'rowEnd',            // 'rowEnd'（直前の行の末尾）| 'lastRow'（最後の行にまとめる）
        linkSpacing: true,
        excludeLockedHiddenLayers: false,
        excludeLockedHiddenItems: false,
        changeArtboardOrder: false
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;                    /* ダイアログ外周の余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネルの余白 / panel margins */
    var PANEL_SPACING = 6;                      /* パネル内の間隔 / spacing inside panels */
    var SPACING_COLUMN_GAP = 20;                /* 間隔欄と「連動」の列の間 / gap between the inputs and the link column */
    var SPACING_FIELD_CHARACTERS = 5;           /* 間隔欄の桁数 / width of the gap fields */

    /**
     * パネルを共通スタイルで初期化する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は変更しない）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードを再配置（行列）", en: "Rearrange Artboards (Grid)" }
        },
        panel: {
            spacing: { ja: "間隔", en: "Spacing" },
            exclusion: { ja: "ロック／非表示を除外", en: "Exclude Locked / Hidden" },
            exception: { ja: "未指定／重複", en: "Unmatched / Duplicate" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            spacingX: { ja: "列間", en: "Column gap" },
            spacingY: { ja: "行間", en: "Row gap" }
        },
        checkbox: {
            linkSpacing: { ja: "連動", en: "Link" },
            excludeLockedHiddenLayers: { ja: "レイヤー", en: "Layers" },
            excludeLockedHiddenItems: { ja: "オブジェクト", en: "Objects" },
            changeArtboardOrder: { ja: "パネル上の並び順を変更", en: "Reorder in Artboards panel" }
        },
        radio: {
            exceptionRowEnd: { ja: "直前の行の末尾に配置", en: "Append to current row" },
            exceptionLastRow: { ja: "最後の行にまとめる", en: "Collect in final row" }
        },
        tooltip: {
            spacingX: { ja: "左右に隣り合うアートボードのあいだにあける間隔です。", en: "Gap left between horizontally adjacent artboards." },
            spacingY: { ja: "上下に隣り合うアートボードのあいだにあける間隔です。", en: "Gap left between vertically adjacent artboards." },
            linkSpacing: { ja: "列間と同じ値を行間にも使います。", en: "Uses the column gap for the row gap too." },
            excludeLayers: {
                ja: "ロックまたは非表示のレイヤーに載っているオブジェクトは動かしません。",
                en: "Leaves objects on locked or hidden layers where they are."
            },
            excludeItems: {
                ja: "ロックまたは非表示のオブジェクトは動かしません。",
                en: "Leaves locked or hidden objects where they are."
            },
            exceptionRowEnd: {
                ja: "アートボード名から行列を読み取れなかったものを、それぞれの行の末尾に置きます。",
                en: "Puts artboards with no readable row-column at the end of each row."
            },
            exceptionLastRow: {
                ja: "アートボード名から行列を読み取れなかったものを、最終行の次の行にまとめます。",
                en: "Collects artboards with no readable row-column into a row after the last one."
            },
            changeArtboardOrder: {
                ja: "アートボードパネルの並び順も、配置後の順序に合わせて入れ替えます。",
                en: "Also reorders the Artboards panel to match the new layout."
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
            noMatch: {
                ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）の形式のアートボード名が見つかりませんでした。",
                en: "No artboard names in '<row><sep><column>' format (e.g., 1-2) or '<prefix><sep><number>' format (e.g., banner-1) were found."
            }
        }
    };

    /**
     * ドット区切りのパスから現在言語のラベルを取得する
     * @param {string} labelPath - "panel.spacing" のようなパス
     * @returns {string} 表示言語の文字列（見つからなければ labelPath）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位と数値入力 / Units and numeric input
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
     * pt値を表示単位に変換する
     * @param {number} points - pt値
     * @param {Object} unitInfo - getUnitInfo() の結果
     * @returns {number} 表示単位の値
     */
    function pointsToDisplayUnit(points, unitInfo) {
        return points / unitInfo.pointsPerUnit;
    }

    /**
     * 表示単位をpt値に変換する
     * @param {number} value - 表示単位の値
     * @param {Object} unitInfo - getUnitInfo() の結果
     * @returns {number} pt値
     */
    function displayUnitToPoints(value, unitInfo) {
        return value * unitInfo.pointsPerUnit;
    }

    /**
     * 表示用の数値を整える（小数点以下3桁まで）
     * @param {number} value - 値
     * @returns {string} 表示用の文字列
     */
    function formatDisplayNumber(value) {
        var rounded = Math.round(value * 1000) / 1000;
        return String(rounded);
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
    // 6. この欄に↑↓キーの増減処理を別に付けない（↑↓キーが二重に効く）
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 間隔の入力行（項目名＋入力欄）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルのパス
     * @param {number} initialPoints - 初期値（pt）
     * @param {Object} unitInfo - 定規単位の情報
     * @param {Function} [onStep] - ∧∨・↑↓キーで増減したあとに呼ぶ処理
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSpacingField(parentGroup, labelPath, tooltipPath, initialPoints, unitInfo, onStep) {
        var fieldGroup = parentGroup.add('group');
        fieldGroup.add('statictext', undefined, labelText(labelPath));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldGroup.add('group');
        stepperFieldGroup.orientation = 'row';
        stepperFieldGroup.alignChildren = ['left', 'center'];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var fieldInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return fieldInput; }, { min: 0, onStep: onStep });
        fieldInput = stepperFieldGroup.add('edittext', undefined, formatDisplayNumber(pointsToDisplayUnit(initialPoints, unitInfo)));
        fieldInput.helpTip = getLabel(tooltipPath);
        fieldInput.characters = SPACING_FIELD_CHARACTERS;
        fieldInput.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(fieldInput, stepperGroup);
        return fieldInput;
    }

    /**
     * 間隔欄の値を pt で読み取る。数値でないか負の値なら初期値を使う
     * @param {EditText} spacingInput - 間隔欄
     * @param {number} fallbackPoints - 読み取れないときの値（pt）
     * @param {Object} unitInfo - 定規単位の情報
     * @returns {number} 間隔（pt）
     */
    function readSpacingPoints(spacingInput, fallbackPoints, unitInfo) {
        var parsedValue = parseFloat(spacingInput.text);
        if (isNaN(parsedValue) || parsedValue < 0) return fallbackPoints;
        return displayUnitToPoints(parsedValue, unitInfo);
    }

    /**
     * チェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルのパス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentPanel.add('checkbox', undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = !!initialValue;
        return optionCheckbox;
    }

    /**
     * 設定ダイアログを表示し、ユーザーが入力した値を返す
     * @param {Object} defaultSettings - 初期値（間隔は pt）
     * @returns {Object|null} 設定。キャンセル時は null
     */
    function showSettingsDialog(defaultSettings) {
        var settingsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        settingsDialog.orientation = 'column';
        settingsDialog.alignChildren = 'fill';
        settingsDialog.margins = DIALOG_MARGINS;

        var rulerUnitInfo = getUnitInfo();

        /* 間隔パネル（列間・行間と連動チェックボックス）/ Spacing panel (column/row gaps with link option) */
        var spacingPanel = settingsDialog.add('panel', undefined, getLabel('panel.spacing') + '（' + rulerUnitInfo.label + '）');
        setupPanel(spacingPanel, PANEL_SPACING);

        var spacingContentGroup = spacingPanel.add('group');
        spacingContentGroup.orientation = 'row';
        spacingContentGroup.alignChildren = ['left', 'center'];
        spacingContentGroup.spacing = SPACING_COLUMN_GAP;

        var spacingInputColumn = spacingContentGroup.add('group');
        spacingInputColumn.orientation = 'column';
        spacingInputColumn.alignChildren = 'left';
        spacingInputColumn.spacing = PANEL_SPACING;

        /* 列間を増減したら、連動時は行間にもミラーする / mirror the column gap to the row gap when linked */
        var spacingXInput = addSpacingField(spacingInputColumn, 'fieldLabel.spacingX', 'tooltip.spacingX', defaultSettings.spacingX, rulerUnitInfo,
            function () { syncLinkedSpacingValue(); });
        spacingXInput.active = true;
        var spacingYInput = addSpacingField(spacingInputColumn, 'fieldLabel.spacingY', 'tooltip.spacingY', defaultSettings.spacingY, rulerUnitInfo);

        var linkColumn = spacingContentGroup.add('group');
        linkColumn.orientation = 'column';
        linkColumn.alignChildren = ['left', 'center'];
        linkColumn.alignment = ['left', 'center'];
        var linkCheckbox = addOptionCheckbox(linkColumn, 'checkbox.linkSpacing', 'tooltip.linkSpacing', defaultSettings.linkSpacing);

        /* 連動時は列間を行間にミラーし、行間欄をディムする / When linked, mirror the column gap to the row gap and dim the row field */
        function syncLinkedSpacingValue() {
            if (linkCheckbox.value) spacingYInput.text = spacingXInput.text;
        }
        function syncLinkedSpacingInputState() {
            spacingYInput.enabled = !linkCheckbox.value;
            spacingYInput.stepperGroup.enabled = spacingYInput.enabled;
            redrawSteppersIn(spacingYInput.stepperGroup);
        }
        syncLinkedSpacingInputState();

        spacingXInput.onChanging = syncLinkedSpacingValue;
        linkCheckbox.onClick = function () {
            syncLinkedSpacingValue();
            syncLinkedSpacingInputState();
        };

        /* ロック／非表示の除外パネル / Locked / hidden exclusion panel */
        var exclusionPanel = settingsDialog.add('panel', undefined, getLabel('panel.exclusion'));
        setupPanel(exclusionPanel, PANEL_SPACING);
        var excludeLayersCheckbox = addOptionCheckbox(exclusionPanel, 'checkbox.excludeLockedHiddenLayers', 'tooltip.excludeLayers', defaultSettings.excludeLockedHiddenLayers);
        var excludeItemsCheckbox = addOptionCheckbox(exclusionPanel, 'checkbox.excludeLockedHiddenItems', 'tooltip.excludeItems', defaultSettings.excludeLockedHiddenItems);

        /* 未指定／重複の扱いパネル / Unmatched / duplicate handling panel */
        var exceptionPanel = settingsDialog.add('panel', undefined, getLabel('panel.exception'));
        setupPanel(exceptionPanel, PANEL_SPACING);
        var rowEndRadio = exceptionPanel.add('radiobutton', undefined, getLabel('radio.exceptionRowEnd'));
        rowEndRadio.helpTip = getLabel('tooltip.exceptionRowEnd');
        var lastRowRadio = exceptionPanel.add('radiobutton', undefined, getLabel('radio.exceptionLastRow'));
        lastRowRadio.helpTip = getLabel('tooltip.exceptionLastRow');
        if (defaultSettings.exceptionMode === 'lastRow') {
            lastRowRadio.value = true;
        } else {
            rowEndRadio.value = true;
        }

        /* オプションパネル / Options panel */
        var optionsPanel = settingsDialog.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel, PANEL_SPACING);
        var changeArtboardOrderCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.changeArtboardOrder', 'tooltip.changeArtboardOrder', defaultSettings.changeArtboardOrder);

        /* OK / キャンセルボタン / OK and Cancel buttons */
        var btnRowGroup = settingsDialog.add('group');
        btnRowGroup.alignment = 'right';
        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        if (settingsDialog.show() !== 1) return null;

        return {
            spacingX: readSpacingPoints(spacingXInput, defaultSettings.spacingX, rulerUnitInfo),
            spacingY: readSpacingPoints(spacingYInput, defaultSettings.spacingY, rulerUnitInfo),
            exceptionMode: lastRowRadio.value ? 'lastRow' : 'rowEnd',
            linkSpacing: linkCheckbox.value,
            excludeLockedHiddenLayers: excludeLayersCheckbox.value,
            excludeLockedHiddenItems: excludeItemsCheckbox.value,
            changeArtboardOrder: changeArtboardOrderCheckbox.value
        };
    }

    // =========================================
    // 配置処理 / Placement
    // =========================================

    /**
     * アクティブアートボード幅の1/8を初期間隔として取得する
     * @param {Document} documentRef - 対象のドキュメント
     * @returns {number} 間隔（pt）
     */
    function getDefaultSpacingFromActiveArtboard(documentRef) {
        var activeArtboardIndex = documentRef.artboards.getActiveArtboardIndex();
        var activeArtboardRect = documentRef.artboards[activeArtboardIndex].artboardRect;
        var activeArtboardWidth = activeArtboardRect[2] - activeArtboardRect[0];
        return Math.round(activeArtboardWidth / 8);
    }

    /**
     * アートボード名を解析する
     * 「数字 + 区切り文字 + 数字」（区切り文字 -, _, x。大文字小文字無視）を行列指定とし、
     * 行列指定の形でない場合のみ「接頭辞 + 区切り文字 + 数字」を接頭辞指定として判定する。
     * 0行・0列・0番、数字だけの接頭辞（行列指定と紛らわしい）は無効
     * @param {string} artboardName - アートボード名
     * @returns {Object|null} { matchType: 'rowColumn' | 'prefixNumber', rowNumber, columnNumber, prefixName }。どちらでもなければ null
     */
    function parseArtboardName(artboardName) {
        var rowColumnMatch = artboardName.match(/^(\d+)[-_x](\d+)$/i);
        if (rowColumnMatch) {
            var rowNumber = parseInt(rowColumnMatch[1], 10);
            var columnNumber = parseInt(rowColumnMatch[2], 10);
            if (rowNumber < 1 || columnNumber < 1) return null;
            return { matchType: 'rowColumn', rowNumber: rowNumber, columnNumber: columnNumber, prefixName: '' };
        }

        var prefixNumberMatch = artboardName.match(/^(.+?)[-_x](\d+)$/i);
        if (!prefixNumberMatch || /^\d+$/.test(prefixNumberMatch[1])) return null;
        var prefixColumnNumber = parseInt(prefixNumberMatch[2], 10);
        if (prefixColumnNumber < 1) return null;
        return { matchType: 'prefixNumber', rowNumber: 0, columnNumber: prefixColumnNumber, prefixName: prefixNumberMatch[1] };
    }

    /**
     * 第1パス：アートボードの情報を収集し、行列指定と接頭辞指定を判定する
     * @param {Artboards} artboards - アートボードのコレクション
     * @returns {Object} placementItems と、最大列番号・最大行番号・接頭辞の出現順などをまとめたグリッド情報
     */
    function collectArtboardPlacements(artboards) {
        var gridInfo = {
            placementItems: [],
            maxColumnNumber: 0,
            maxRowNumber: 0,
            prefixOrder: [],
            prefixRowOffsetByName: {},
            hasMatchedArtboard: false
        };

        for (var artboardIndex = 0; artboardIndex < artboards.length; artboardIndex++) {
            var artboard = artboards[artboardIndex];
            var artboardRect = artboard.artboardRect; // [left, top, right, bottom]
            var placementItem = {
                artboard: artboard,
                width: artboardRect[2] - artboardRect[0],
                height: artboardRect[1] - artboardRect[3],
                matched: false,
                matchType: '',
                prefixName: '',
                rowNumber: 0,
                columnNumber: 0,
                assignedRow: 0,
                assignedColumn: 0
            };

            var nameMatch = parseArtboardName(artboard.name);
            if (nameMatch) {
                placementItem.matched = true;
                placementItem.matchType = nameMatch.matchType;
                placementItem.prefixName = nameMatch.prefixName;
                placementItem.rowNumber = nameMatch.rowNumber;
                placementItem.columnNumber = nameMatch.columnNumber;

                if (nameMatch.matchType === 'rowColumn') {
                    if (nameMatch.rowNumber > gridInfo.maxRowNumber) gridInfo.maxRowNumber = nameMatch.rowNumber;
                } else if (gridInfo.prefixRowOffsetByName[nameMatch.prefixName] === undefined) {
                    /* 接頭辞は出現順に1行ずつ割り当てる / one row per prefix, in order of appearance */
                    gridInfo.prefixRowOffsetByName[nameMatch.prefixName] = gridInfo.prefixOrder.length + 1;
                    gridInfo.prefixOrder.push(nameMatch.prefixName);
                }
                if (nameMatch.columnNumber > gridInfo.maxColumnNumber) gridInfo.maxColumnNumber = nameMatch.columnNumber;
                gridInfo.hasMatchedArtboard = true;
            }
            gridInfo.placementItems.push(placementItem);
        }

        /* 接頭辞の行は通常の行列指定の下に続く / prefix rows follow the numeric rows */
        gridInfo.maxMatchedRowNumber = gridInfo.maxRowNumber + gridInfo.prefixOrder.length;
        return gridInfo;
    }

    /**
     * 第2パス：順番に走査して行・列番号を割り当てる
     * 例外処理モード / Exception handling modes:
     *   'rowEnd'  : 非一致名は直前にマッチした行に追従し、列は (maxColumnNumber + 1) から右へ並べる
     *               先頭の非一致は行1に配置 / Leading unmatched goes to row 1
     *   'lastRow' : 非一致名は (maxMatchedRowNumber + 1) 行目に、出現順に列 1, 2, 3, ... で並べる
     * 接頭辞指定は、通常の行列指定の下に接頭辞ごとの行として配置する。
     * 重複名は元の (row, col) と衝突した時点で、その行の末尾へ押し出す。
     * @param {Object} gridInfo - collectArtboardPlacements() の結果（assignedRow / assignedColumn を書き込む）
     * @param {string} exceptionMode - 'rowEnd' | 'lastRow'
     * @returns {void}
     */
    function assignGridSlots(gridInfo, exceptionMode) {
        var occupiedSlots = {};                                     // "row,col" → true
        var nextFreeColumnByRow = {};                               // row → next column to try
        var rowEndFirstColumn = gridInfo.maxColumnNumber + 1;       // 行末に押し出すときの開始列 / first column at a row end
        var exceptionRowNumber = gridInfo.maxMatchedRowNumber + 1;  // 'lastRow' 用の行番号 / row for 'lastRow'
        var currentRowNumber = 1;

        for (var assignIndex = 0; assignIndex < gridInfo.placementItems.length; assignIndex++) {
            var assignTarget = gridInfo.placementItems[assignIndex];

            if (assignTarget.matched) {
                var targetRowNumber = assignTarget.rowNumber;
                if (assignTarget.matchType === 'prefixNumber') {
                    targetRowNumber = gridInfo.maxRowNumber + gridInfo.prefixRowOffsetByName[assignTarget.prefixName];
                }

                currentRowNumber = targetRowNumber;
                assignTarget.assignedRow = targetRowNumber;
                var requestedKey = targetRowNumber + ',' + assignTarget.columnNumber;
                if (!occupiedSlots[requestedKey]) {
                    /* 通常配置 / Normal placement */
                    occupiedSlots[requestedKey] = true;
                    assignTarget.assignedColumn = assignTarget.columnNumber;
                } else {
                    /* 重複 → その行の末尾へ / Duplicate → push to row end */
                    assignTarget.assignedColumn = reserveNextFreeColumn(targetRowNumber, rowEndFirstColumn, occupiedSlots, nextFreeColumnByRow);
                }
            } else if (exceptionMode === 'lastRow') {
                assignTarget.assignedRow = exceptionRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(exceptionRowNumber, 1, occupiedSlots, nextFreeColumnByRow);
            } else {
                /* 'rowEnd' */
                assignTarget.assignedRow = currentRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(currentRowNumber, rowEndFirstColumn, occupiedSlots, nextFreeColumnByRow);
            }
        }
    }

    /**
     * 第3〜5パス：列ごとの最大幅・行ごとの最大高さから、各アートボードの新しい位置と移動量を求める
     * 例外領域に入る境界（未指定・重複の列、未指定の行）では間隔を3倍にする。
     * IllustratorのY座標は上がプラス、下がマイナス。
     * @param {Object[]} placementItems - 行列を割り当てたアートボード情報（oldRect / newLeft / newTop / deltaX / deltaY を書き込む）
     * @param {number} spacingX - 列間（pt）
     * @param {number} spacingY - 行間（pt）
     * @param {number} columnBoundary - 間隔を3倍にする列番号
     * @param {number} rowBoundary - 間隔を3倍にする行番号
     * @returns {void}
     */
    function computeArtboardDestinations(placementItems, spacingX, spacingY, columnBoundary, rowBoundary) {
        var originX = 0;    // 配置の開始X座標 / X origin
        var originY = 0;    // 配置の開始Y座標 / Y origin

        /* 列ごとの最大幅・行ごとの最大高さを集計する / Per-column max width and per-row max height */
        var columnWidthByNumber = {};
        var rowHeightByNumber = {};
        for (var aggregateIndex = 0; aggregateIndex < placementItems.length; aggregateIndex++) {
            var aggregateItem = placementItems[aggregateIndex];
            var columnKey = aggregateItem.assignedColumn;
            var rowKey = aggregateItem.assignedRow;
            if (columnWidthByNumber[columnKey] === undefined || aggregateItem.width > columnWidthByNumber[columnKey]) {
                columnWidthByNumber[columnKey] = aggregateItem.width;
            }
            if (rowHeightByNumber[rowKey] === undefined || aggregateItem.height > rowHeightByNumber[rowKey]) {
                rowHeightByNumber[rowKey] = aggregateItem.height;
            }
        }

        /* 使われている列番号・行番号の昇順で累積オフセットを計算する / Cumulative offsets over the used numbers */
        var columnOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(columnWidthByNumber), columnWidthByNumber, spacingX, columnBoundary
        );
        var rowOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(rowHeightByNumber), rowHeightByNumber, spacingY, rowBoundary
        );

        for (var placeIndex = 0; placeIndex < placementItems.length; placeIndex++) {
            var placeItem = placementItems[placeIndex];
            var oldRect = placeItem.artboard.artboardRect;
            placeItem.oldRect = oldRect;
            placeItem.newLeft = originX + columnOffsetByNumber[placeItem.assignedColumn];
            placeItem.newTop = originY - rowOffsetByNumber[placeItem.assignedRow];
            placeItem.deltaX = placeItem.newLeft - oldRect[0];
            placeItem.deltaY = placeItem.newTop - oldRect[1];
        }
    }

    /**
     * 第6パス：アートボード上のオブジェクトを、アートボードと同じだけ移動する
     * オブジェクトの中心が元のアートボード矩形内にあるものを対象にする。
     * @param {Layers} layerCollection - ドキュメントのレイヤー
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {boolean} excludeLockedHiddenLayers - ロック／非表示レイヤーのオブジェクトを動かさないか
     * @param {boolean} excludeLockedHiddenItems - ロック／非表示のオブジェクトを動かさないか
     * @returns {void}
     */
    function translateItemsWithArtboards(layerCollection, placementItems, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        var itemTranslations = [];
        appendItemTranslations(layerCollection, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems);
        for (var translateIndex = 0; translateIndex < itemTranslations.length; translateIndex++) {
            var pendingMove = itemTranslations[translateIndex];
            if (pendingMove.deltaX === 0 && pendingMove.deltaY === 0) continue;
            try {
                pendingMove.item.translate(pendingMove.deltaX, pendingMove.deltaY);
            } catch (translateError) {
                /* ロック・非表示などで失敗したオブジェクトはスキップ
                   Skip items that can't be translated (locked, hidden layer, etc.) */
            }
        }
    }

    /**
     * 第7パス：アートボードの矩形を更新する
     * @param {Object[]} placementItems - 新しい位置を求めたアートボード情報
     * @returns {void}
     */
    function applyArtboardRects(placementItems) {
        for (var applyIndex = 0; applyIndex < placementItems.length; applyIndex++) {
            var applyItem = placementItems[applyIndex];
            var newRight = applyItem.newLeft + applyItem.width;
            var newBottom = applyItem.newTop - applyItem.height;
            applyItem.artboard.artboardRect = [applyItem.newLeft, applyItem.newTop, newRight, newBottom];
        }
    }

    /**
     * アートボード名を解析してグリッド状に再配置する
     * 「行-列」は行列指定、「接頭辞-番号」は接頭辞ごとの行として扱い、
     * 列幅・行高は「列ごと・行ごとの最大サイズ」で計算する。
     * @returns {void}
     */
    function arrangeArtboardsByRowColumnName() {
        if (app.documents.length === 0) {
            return;
        }

        var documentRef = app.activeDocument;
        var defaultSpacing = getDefaultSpacingFromActiveArtboard(documentRef);

        var arrangeSettings = showSettingsDialog({
            spacingX: defaultSpacing,
            spacingY: defaultSpacing,
            exceptionMode: DIALOG_DEFAULTS.exceptionMode,
            linkSpacing: DIALOG_DEFAULTS.linkSpacing,
            excludeLockedHiddenLayers: DIALOG_DEFAULTS.excludeLockedHiddenLayers,
            excludeLockedHiddenItems: DIALOG_DEFAULTS.excludeLockedHiddenItems,
            changeArtboardOrder: DIALOG_DEFAULTS.changeArtboardOrder
        });
        if (!arrangeSettings) return;

        var gridInfo = collectArtboardPlacements(documentRef.artboards);
        if (!gridInfo.hasMatchedArtboard) {
            alert(getLabel('alert.noMatch'));
            return;
        }

        assignGridSlots(gridInfo, arrangeSettings.exceptionMode);
        computeArtboardDestinations(gridInfo.placementItems, arrangeSettings.spacingX, arrangeSettings.spacingY,
            gridInfo.maxColumnNumber + 1, gridInfo.maxMatchedRowNumber + 1);

        /* オブジェクトを先に動かしてからアートボードを動かす / Move the objects first, then the artboards */
        translateItemsWithArtboards(documentRef.layers, gridInfo.placementItems,
            arrangeSettings.excludeLockedHiddenLayers, arrangeSettings.excludeLockedHiddenItems);
        applyArtboardRects(gridInfo.placementItems);

        /* 第8パス: アートボードパネル上の順番を、行→列の順に並べ替える / Reorder the Artboards panel row-then-column */
        if (arrangeSettings.changeArtboardOrder) {
            reorderArtboardsByGridOrder(documentRef, gridInfo.placementItems);
        }

        /* 画面を全体表示に更新 / Fit all in view */
        app.executeMenuCommand('fitall');
    }

    // =========================================
    // 配置ヘルパー / Placement helpers
    // =========================================

    /**
     * 指定行の firstColumn 以降で空きスロットを予約する
     * @param {number} rowNumber - 行番号
     * @param {number} firstColumn - その行で最初に試す列番号
     * @param {Object} occupiedSlots - "row,col" → true の表
     * @param {Object} nextFreeColumnByRow - 行番号 → 次に試す列番号の表
     * @returns {number} 予約した列番号
     */
    function reserveNextFreeColumn(rowNumber, firstColumn, occupiedSlots, nextFreeColumnByRow) {
        if (nextFreeColumnByRow[rowNumber] === undefined) {
            nextFreeColumnByRow[rowNumber] = firstColumn;
        }
        while (true) {
            var candidateColumn = nextFreeColumnByRow[rowNumber]++;
            var candidateSlotKey = rowNumber + ',' + candidateColumn;
            if (!occupiedSlots[candidateSlotKey]) {
                occupiedSlots[candidateSlotKey] = true;
                return candidateColumn;
            }
        }
    }

    /**
     * オブジェクトの数値キーを昇順で取り出す
     * @param {Object} numericKeyedObject - 数値キーの表
     * @returns {number[]} 昇順のキー
     */
    function collectSortedNumericKeys(numericKeyedObject) {
        var sortedKeys = [];
        for (var numericKey in numericKeyedObject) {
            if (numericKeyedObject.hasOwnProperty(numericKey)) {
                sortedKeys.push(parseInt(numericKey, 10));
            }
        }
        sortedKeys.sort(function (a, b) { return a - b; });
        return sortedKeys;
    }

    /**
     * レイヤーを再帰的に走査し、直下の page item の移動量を集める
     * 必要に応じてロック／非表示レイヤー・オブジェクトを除外する。
     * @param {Layers} layerCollection - 走査するレイヤー
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {Object[]} itemTranslations - { item, deltaX, deltaY } の格納先
     * @param {boolean} excludeLockedHiddenLayers - ロック／非表示レイヤーを除外するか
     * @param {boolean} excludeLockedHiddenItems - ロック／非表示オブジェクトを除外するか
     * @returns {void}
     */
    function appendItemTranslations(layerCollection, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            if (excludeLockedHiddenLayers && shouldSkipLayerForItemMove(layer)) continue;

            for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
                var pageItem = layer.pageItems[itemIndex];
                if (pageItem.parent !== layer) continue;
                if (excludeLockedHiddenItems && shouldSkipPageItemForMove(pageItem)) continue;
                appendTranslationForItem(pageItem, placementItems, itemTranslations);
            }

            if (layer.layers && layer.layers.length > 0) {
                appendItemTranslations(layer.layers, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems);
            }
        }
    }

    /**
     * 移動対象から除外するレイヤーか判定する
     * @param {Layer} layer - 対象のレイヤー
     * @returns {boolean} ロックまたは非表示なら true
     */
    function shouldSkipLayerForItemMove(layer) {
        try {
            return layer.locked || !layer.visible;
        } catch (layerStateError) {
            return false;
        }
    }

    /**
     * 移動対象から除外するオブジェクトか判定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} ロックまたは非表示なら true
     */
    function shouldSkipPageItemForMove(pageItem) {
        try {
            return pageItem.locked || pageItem.hidden;
        } catch (pageItemStateError) {
            return false;
        }
    }

    /**
     * 単一オブジェクトの移動量を追加する（中心を含む最初の元アートボードの移動量）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {Object[]} itemTranslations - { item, deltaX, deltaY } の格納先
     * @returns {void}
     */
    function appendTranslationForItem(pageItem, placementItems, itemTranslations) {
        var center = getItemGeometricCenter(pageItem);
        if (!center) return;

        for (var placementIndex = 0; placementIndex < placementItems.length; placementIndex++) {
            var placement = placementItems[placementIndex];
            if (isCenterInsideRect(center, placement.oldRect)) {
                itemTranslations.push({ item: pageItem, deltaX: placement.deltaX, deltaY: placement.deltaY });
                break;
            }
        }
    }

    /**
     * オブジェクトの幾何中心を取得する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number[]|null} [x, y]。取得できない場合は null
     */
    function getItemGeometricCenter(pageItem) {
        try {
            var bounds = pageItem.geometricBounds; // [left, top, right, bottom]
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (geometricBoundsError) {
            return null;
        }
    }

    /**
     * 中心座標が矩形内にあるか（artboardRect の Y は上がプラス、下がマイナス）
     * @param {number[]} center - [x, y]
     * @param {number[]} rect - [left, top, right, bottom]
     * @returns {boolean} 矩形内なら true
     */
    function isCenterInsideRect(center, rect) {
        return center[0] >= rect[0] && center[0] <= rect[2] &&
            center[1] <= rect[1] && center[1] >= rect[3];
    }

    /**
     * アートボードパネル上の順序を、割り当てた (assignedRow, assignedColumn) の昇順に再構築する
     * 一時名にリネームしてから新規作成→旧アートボードを削除することで名称重複を回避する。
     * @param {Document} documentRef - 対象のドキュメント
     * @param {Object[]} placementItems - 行列を割り当てたアートボード情報
     * @returns {void}
     */
    function reorderArtboardsByGridOrder(documentRef, placementItems) {
        var sortedPlacements = placementItems.slice();
        sortedPlacements.sort(function (placementA, placementB) {
            if (placementA.assignedRow !== placementB.assignedRow) return placementA.assignedRow - placementB.assignedRow;
            return placementA.assignedColumn - placementB.assignedColumn;
        });

        var artboards = documentRef.artboards;
        var originalCount = artboards.length;

        var orderedArtboardSpecs = [];
        for (var collectIndex = 0; collectIndex < sortedPlacements.length; collectIndex++) {
            var sourceArtboard = sortedPlacements[collectIndex].artboard;
            var sourceRect = sourceArtboard.artboardRect;
            orderedArtboardSpecs.push({
                name: sourceArtboard.name,
                rect: [sourceRect[0], sourceRect[1], sourceRect[2], sourceRect[3]]
            });
        }

        for (var renameIndex = 0; renameIndex < originalCount; renameIndex++) {
            artboards[renameIndex].name = '__reorder_tmp__' + renameIndex;
        }

        for (var addIndex = 0; addIndex < orderedArtboardSpecs.length; addIndex++) {
            var newArtboard = artboards.add(orderedArtboardSpecs[addIndex].rect);
            newArtboard.name = orderedArtboardSpecs[addIndex].name;
        }

        for (var removeIndex = originalCount - 1; removeIndex >= 0; removeIndex--) {
            artboards[removeIndex].remove();
        }

        documentRef.artboards.setActiveArtboardIndex(0);
    }

    /**
     * 各番号（列または行）の累積オフセットを計算する
     * boundaryNumber と一致するキーの直前のギャップは3倍に広げる。
     * @param {number[]} sortedNumbers - 昇順の番号
     * @param {Object} sizeByNumber - 番号 → 幅または高さ
     * @param {number} gap - 間隔（pt）
     * @param {number} boundaryNumber - 間隔を3倍にする番号
     * @returns {Object} 番号 → オフセット
     */
    function computeCumulativeOffsets(sortedNumbers, sizeByNumber, gap, boundaryNumber) {
        var offsetByNumber = {};
        var cumulative = 0;
        for (var orderIndex = 0; orderIndex < sortedNumbers.length; orderIndex++) {
            var currentNumber = sortedNumbers[orderIndex];
            if (orderIndex > 0) {
                var currentGap = gap;
                if (currentNumber === boundaryNumber) currentGap *= 3;
                cumulative += currentGap;
            }
            offsetByNumber[currentNumber] = cumulative;
            cumulative += sizeByNumber[currentNumber];
        }
        return offsetByNumber;
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    arrangeArtboardsByRowColumnName();

})();
