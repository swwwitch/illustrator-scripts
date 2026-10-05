#target illustrator
#targetengine "MakeRectangleFromGuidesEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ガイドの交点で区切られた区画に、塗りつぶした長方形を一括生成します。対象にするガイドと作成先はダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeRectangleFromGuides.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4907511336ad

### Overview

Fills every area bounded by guide intersections with a generated rectangle. A dialog selects which guides to use and where the rectangles go.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeRectangleFromGuides.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MakeRectangleFromGuides";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeRectangleFromGuides.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeRectangleFromGuides.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4907511336ad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 作成先に専用レイヤーを選んだときのレイヤー名 / Layer name used when the dedicated destination is selected */
    var OUTPUT_LAYER_NAME = "Generated Rectangles";

    /* 生成した長方形の不透明度（%）/ Opacity of the generated rectangles (%) */
    var RECT_OPACITY = 50;

    /* 2つの座標を同一とみなす許容誤差（pt）/ Tolerance for treating two coordinates as identical (pt) */
    var COORD_TOLERANCE = 0.01;

    /* ルーラーガイド判定：アートボードをこの長さ以上はみ出すガイドを全面ガイドとみなす（pt）/ A guide overhanging the artboard by at least this much counts as a ruler guide (pt) */
    var RULER_GUIDE_MARGIN_PT = 100;

    /* プレビューで描く長方形の上限（多すぎるとダイアログ操作が重くなる）/ Cap on previewed rectangles, beyond which the dialog would crawl */
    var PREVIEW_MAX_RECTS = 500;

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

    var FIELD_ROW_SPACING  = 6;                /* ラジオと入力欄の間隔 / gap between a radio and its input */
    var LAYER_NAME_CHARS   = 15;               /* レイヤー名入力欄の最小幅（文字数）/ minimum width of the layer name field (characters) */
    var INDENT_MARGINS     = [20, 0, 0, 0];    /* ラジオの下に続く行の字下げ / indent for a row that follows its radio */
    var MESSAGE_MARGINS    = [0, 0, 0, 10];    /* パネル最上部のメッセージの下余白 / space under the message at the top of a panel */
    var OFFSET_CHARS       = 5;                /* オフセット入力欄の幅（文字数）/ width of the offset field (characters) */
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
        "q": { unit: "mm", amount: 0.25 }, /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 }, /* 歯 / H */
        "p": { unit: "pc", amount: 1 }     /* パイカ / pica */
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
            var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title:   { ja: "ガイドから長方形を作成", en: "Create Rectangles from Guides" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        guideSource: {
            panelTitle: { ja: "対象となるガイド", en: "Source Guides" }
        },
        guideType: {
            panelTitle: { ja: "ガイドの種類", en: "Guide Type" },
            all:        { ja: "すべてのガイド", en: "All Guides" },
            rulerOnly:  { ja: "ルーラーガイドのみ", en: "Ruler Guides Only" }
        },
        targetLayer: {
            panelTitle: { ja: "レイヤー", en: "Layers" },
            allLayers:  { ja: "すべてのレイヤー", en: "All Layers" },
            activeOnly: { ja: "現在のレイヤーのみ", en: "Current Layer Only" }
        },
        guideOption: {
            panelTitle:         { ja: "絞り込み", en: "Filters" },
            activeArtboardOnly: { ja: "現在のアートボードのみ", en: "Current Artboard Only" },
            includeLocked:      { ja: "ロックされたレイヤーを含む", en: "Include Locked Layers" }
        },
        rectangle: {
            panelTitle: { ja: "作成する長方形", en: "Rectangles to Create" }
        },
        destination: {
            panelTitle:  { ja: "作成先", en: "Destination" },
            activeLayer: { ja: "現在のレイヤー", en: "Current Layer" },
            outputLayer: { ja: "指定レイヤー", en: "Specific Layer" }
        },
        rectangleOption: {
            panelTitle:     { ja: "後処理", en: "After Creation" },
            offset:         { ja: "オフセット", en: "Offset" },
            mergeAdjacent:  { ja: "長方形を1つに結合", en: "Merge Into a Single Path" },
            convertToShape: { ja: "シェイプに変換", en: "Convert to Shape" }
        },
        summary: {
            guides: {
                ja: "縦{vertical}本・横{horizontal}本",
                en: "{vertical} vertical / {horizontal} horizontal"
            },
            count: {
                ja: "{count}個",
                en: "{count}"
            },
            merged: {
                ja: "{count}個 → 結合して1つ",
                en: "{count} → merged into 1"
            }
        },
        tooltip: {
            guideType: {
                ja: "「ルーラーガイドのみ」は、アートボードをまたぐ長さのガイドだけを拾います",
                en: "\"Ruler Guides Only\" keeps just the guides that run past the artboard edges"
            },
            targetLayer: {
                ja: "ガイドを探す範囲です。長方形の作成先とは別の設定です",
                en: "Where to look for guides. This is separate from where the rectangles are created"
            },
            activeArtboardOnly: {
                ja: "現在のアートボードの範囲内にあるガイドだけを使います",
                en: "Uses only the guides that fall inside the current artboard"
            },
            includeLocked: {
                ja: "ロックされたレイヤーを一時的に解除してガイドを集め、集め終わったらロックを戻します",
                en: "Temporarily unlocks locked layers to collect their guides, then restores the lock"
            },
            destination: {
                ja: "長方形を作成するレイヤーです",
                en: "Layer that receives the rectangles"
            },
            outputLayerName: {
                ja: "作成先のレイヤー名。同名のレイヤーがすでにあれば、そのレイヤーを再利用します",
                en: "Name of the destination layer. An existing layer with the same name is reused"
            },
            offset: {
                ja: "各長方形を四辺とも広げます（パスのオフセットと同じで、マイナス値なら縮みます）。単位はルーラーに従います",
                en: "Grows every rectangle on all four sides, like Offset Path; a negative value shrinks it. The unit follows the ruler"
            },
            mergeAdjacent: {
                ja: "作成した長方形をすべて合体して1つのパスにします",
                en: "Unites every generated rectangle into a single path"
            },
            convertToShape: {
                ja: "作成した長方形をライブシェイプに変換します。結合する場合は結合後に変換します",
                en: "Converts the result into live shapes. When merging, the conversion runs after the merge"
            },
            guideCount: {
                ja: "下の条件で見つかったガイドの本数です。同じ位置に重なったガイドは1本として数えます",
                en: "Guides found under the settings below. Guides at the same position count as one"
            },
            preview: {
                ja: "作成される長方形をカンバス上に仮表示します。結合とシェイプ変換はOKを押したときに適用されます",
                en: "Draws the rectangles on the canvas. Merging and shape conversion are applied when you press OK"
            },
            summaryCount: {
                ja: "OKを押したときに作成される長方形の数です",
                en: "How many rectangles pressing OK will create"
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
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:        { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            notEnoughGuides:   { ja: "条件に合うガイドが足りません。縦・横それぞれ2本以上必要です。", en: "Not enough matching guides. At least two vertical and two horizontal guides are required." },
            lockedActiveLayer: { ja: "現在のレイヤーがロックまたは非表示のため作成できません。ロックを解除するか、作成先を「指定レイヤー」にしてください。", en: "Cannot draw because the current layer is locked or hidden. Unlock it, or set the destination to the specific layer." }
        }
    };

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
     * 入力値を pt に変換する
     * @param {string} inputText - 入力欄の値
     * @param {number} factor - 単位の pt 換算係数
     * @returns {number} pt に変換した値（数値として読めなければ0）
     */
    function convertToPt(inputText, factor) {
        var numericValue = parseFloat(inputText);
        return isNaN(numericValue) ? 0 : numericValue * factor;
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel, 6);
        return createdPanel;
    }

    /**
     * 左寄せの縦並びグループを生成する（ラジオ列・チェックボックス列など）
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addLeftAlignedColumn(parentContainer) {
        var createdGroup = parentContainer.add("group");
        createdGroup.orientation = "column";
        createdGroup.alignment = ["left", "top"];
        createdGroup.alignChildren = ["left", "center"];
        return createdGroup;
    }

    /**
     * 2択のラジオボタン列を生成する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} firstLabel - 1つ目のラベル
     * @param {string} secondLabel - 2つ目のラベル
     * @param {number} selectedIndex - 初期選択（0=1つ目、1=2つ目）
     * @param {string} [helpTipText] - ツールチップ（親コンテナと両ボタンに付与）
     * @returns {RadioButton[]} [1つ目, 2つ目] のラジオボタン
     */
    function addRadioPair(parentContainer, firstLabel, secondLabel, selectedIndex, helpTipText) {
        var radioColumn = addLeftAlignedColumn(parentContainer);
        var firstRadio = radioColumn.add("radiobutton", undefined, firstLabel);
        var secondRadio = radioColumn.add("radiobutton", undefined, secondLabel);
        ((selectedIndex === 1) ? secondRadio : firstRadio).value = true;
        if (helpTipText) {
            /* パネルの余白でもボタン上でも出るように両方へ付ける / Attach to both so the tip shows over the panel and the buttons */
            parentContainer.helpTip = helpTipText;
            firstRadio.helpTip = helpTipText;
            secondRadio.helpTip = helpTipText;
        }
        return [firstRadio, secondRadio];
    }

    /**
     * パネル内にチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelText - ラベル
     * @param {boolean} defaultValue - 初期状態
     * @param {string} [helpTipText] - ツールチップ
     * @returns {Checkbox} 生成したチェックボックス
     */
    function addPanelCheckbox(parentPanel, labelText, defaultValue, helpTipText) {
        var checkbox = parentPanel.add("checkbox", undefined, labelText);
        /* パネルの fill を打ち消して幅いっぱいに広げない / Cancel the panel's fill so the control keeps its natural width */
        checkbox.alignment = ["left", "center"];
        checkbox.value = defaultValue;
        if (helpTipText) {
            checkbox.helpTip = helpTipText;
        }
        return checkbox;
    }

    /**
     * パネル最上部に置く1行メッセージを生成する（左右中央、下に余白）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} [helpTipText] - ツールチップ
     * @returns {StaticText} 生成したメッセージ欄
     */
    function addPanelMessage(parentPanel, helpTipText) {
        var messageRow = parentPanel.add("group");
        setupRow(messageRow, "fill", 0);
        messageRow.margins = MESSAGE_MARGINS;
        /* 幅いっぱいの箱にして中央揃えで描く。中身の幅で中央寄せすると、文字数が増えたとき収まらない
           Give it the full width and center the text inside; sizing to the content would clip a longer message */
        var messageText = messageRow.add('statictext {justify: "center"}');
        messageText.alignment = ["fill", "center"];
        if (helpTipText) {
            messageText.helpTip = helpTipText;
        }
        return messageText;
    }

    /**
     * ダイアログウィンドウを生成する（縦積みの構成）
     * @param {string} windowTitle - ウィンドウタイトル
     * @returns {Window} 生成したダイアログ
     */
    function createDialogWindow(windowTitle) {
        var dialogWindow = new Window("dialog", windowTitle);
        setupWindow(dialogWindow);
        return dialogWindow;
    }

    // =========================================
    // レイヤー操作 / Layer handling
    // =========================================

    /**
     * 名前でレイヤーを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - 探すレイヤー名
     * @returns {Layer|null} 見つかったレイヤー（なければ null）
     */
    function findLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
        return null;
    }

    /**
     * レイヤー名の入力値を整える（前後の空白を落とし、空なら既定名に戻す）
     * @param {string} inputText - 入力欄の値
     * @returns {string} 使用するレイヤー名
     */
    function normalizeLayerName(inputText) {
        var trimmedName = (inputText || "").replace(/^\s+|\s+$/g, "");
        return trimmedName || OUTPUT_LAYER_NAME;
    }

    /**
     * 長方形の作成先レイヤーを用意してアクティブにする（メッセージは出さない）
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {{layer: Layer, created: boolean}|null} 用意できた作成先（用意できなければ null）
     */
    function activateOutputLayer(doc, options) {
        if (!options.useOutputLayer) {
            /* 現在のレイヤーに描くので、ロック・非表示なら描けない / Drawing into the current layer, so a locked or hidden one is unusable */
            var activeLayer = doc.activeLayer;
            if (activeLayer.locked || !activeLayer.visible) {
                return null;
            }
            return { layer: activeLayer, created: false };
        }
        /* 同名レイヤーがあれば再利用し、なければ作る / Reuse a layer with the same name, or create one */
        var outputLayer = findLayerByName(doc, options.outputLayerName);
        var wasCreated = false;
        if (!outputLayer) {
            outputLayer = doc.layers.add();
            outputLayer.name = options.outputLayerName;
            wasCreated = true;
        }
        /* 使い回すレイヤーは描ける状態に戻す / Make a reused layer drawable again */
        outputLayer.locked = false;
        outputLayer.visible = true;
        doc.activeLayer = outputLayer;
        return { layer: outputLayer, created: wasCreated };
    }

    /**
     * 作成先レイヤーを用意し、用意できなければ理由を知らせる
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {boolean} 作成先を確保できたら true
     */
    function prepareOutputLayer(doc, options) {
        if (activateOutputLayer(doc, options)) {
            return true;
        }
        alert(getLabel('alert.lockedActiveLayer'));
        return false;
    }

    /**
     * ロックされているレイヤーを一時的に解除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer[]} 解除したレイヤー（復元用）
     */
    function unlockLockedLayers(doc) {
        var unlockedLayers = [];
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].locked) {
                doc.layers[i].locked = false;
                unlockedLayers.push(doc.layers[i]);
            }
        }
        return unlockedLayers;
    }

    /**
     * 一時解除したレイヤーのロックを元に戻す
     * @param {Layer[]} layers - 復元するレイヤー
     * @returns {void}
     */
    function relockLayers(layers) {
        for (var i = 0; i < layers.length; i++) {
            layers[i].locked = true;
        }
    }

    // =========================================
    // ガイドの収集と分類 / Collecting and classifying guides
    // =========================================

    /**
     * ガイドパスの向き・位置・長さを求める（縦横どちらでもないものは対象外）
     * @param {PathItem} guidePath - 対象のガイドパス
     * @returns {{isVertical: boolean, position: number, length: number}|null} 幾何情報（対象外なら null）
     */
    function getGuideGeometry(guidePath) {
        if (guidePath.pathPoints.length < 2) {
            return null;
        }
        var startAnchor = guidePath.pathPoints[0].anchor;
        var endAnchor = guidePath.pathPoints[1].anchor;
        var deltaX = endAnchor[0] - startAnchor[0];
        var deltaY = endAnchor[1] - startAnchor[1];
        var isVertical = Math.abs(deltaX) < COORD_TOLERANCE;
        if (!isVertical && Math.abs(deltaY) >= COORD_TOLERANCE) {
            return null;
        }
        return {
            isVertical: isVertical,
            position: isVertical ? startAnchor[0] : startAnchor[1],
            length: Math.sqrt(deltaX * deltaX + deltaY * deltaY)
        };
    }

    /**
     * アートボードをまたぐ長さのガイドか（＝ルーラーガイドとみなせるか）を判定する
     * @param {PathItem} guidePath - 判定するガイドパス
     * @param {number} artboardWidth - アートボードの幅（pt）
     * @param {number} artboardHeight - アートボードの高さ（pt）
     * @returns {boolean} ルーラーガイドとみなせるなら true
     */
    function isRulerGuide(guidePath, artboardWidth, artboardHeight) {
        /* 面積を持たない2点の直線だけが対象 / Only a two-point line with no area qualifies */
        if (guidePath.pathPoints.length !== 2 || Math.abs(guidePath.area) >= COORD_TOLERANCE) {
            return false;
        }
        var geometry = getGuideGeometry(guidePath);
        if (!geometry) {
            return false;
        }
        var artboardSpan = geometry.isVertical ? artboardHeight : artboardWidth;
        return geometry.length > artboardSpan + RULER_GUIDE_MARGIN_PT;
    }

    /**
     * ガイドが現在のアートボードの範囲内にあるかを判定する
     * @param {PathItem} guidePath - 判定するガイドパス
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @returns {boolean} 範囲内なら true
     */
    function isWithinArtboard(guidePath, artboardRect) {
        var geometry = getGuideGeometry(guidePath);
        if (!geometry) {
            return false;
        }
        /* 縦ガイドは左右、横ガイドは上下の範囲で判定する（Yは上が正）/ Verticals are bounded left-right, horizontals top-bottom (Y grows upward) */
        return geometry.isVertical
            ? (geometry.position >= artboardRect[0] && geometry.position <= artboardRect[2])
            : (geometry.position <= artboardRect[1] && geometry.position >= artboardRect[3]);
    }

    /**
     * 条件に合うガイドパスを集める
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {PathItem[]} 集めたガイドパス
     */
    function collectGuidePaths(doc, options) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var artboardWidth = artboardRect[2] - artboardRect[0];
        var artboardHeight = artboardRect[1] - artboardRect[3];

        /* 「現在のレイヤーのみ」はレイヤー配下だけを列挙する（レイヤー同士の比較を避けられる）/ Enumerate just the active layer, which avoids comparing layer objects */
        /* pathItems は都度の DOM アクセスが重いので参照と件数を控える / Cache the collection and its length; each DOM access is costly */
        var pathItems = options.activeLayerOnly ? doc.activeLayer.pathItems : doc.pathItems;
        var guidePaths = [];
        for (var i = 0, itemCount = pathItems.length; i < itemCount; i++) {
            var pathItem = pathItems[i];
            if (!pathItem.guides || pathItem.locked || pathItem.hidden) continue;
            if (pathItem.layer.locked && !options.includeLocked) continue;
            if (options.useRulerOnly && !isRulerGuide(pathItem, artboardWidth, artboardHeight)) continue;
            if (options.activeArtboardOnly && !isWithinArtboard(pathItem, artboardRect)) continue;
            guidePaths.push(pathItem);
        }
        return guidePaths;
    }

    /**
     * ソート済みの座標配列から、重複とみなせる隣接値を取り除く
     * @param {number[]} sortedPositions - ソート済みの座標配列
     * @returns {number[]} 重複を除いた配列
     */
    function dedupeSortedPositions(sortedPositions) {
        var uniquePositions = [];
        for (var i = 0; i < sortedPositions.length; i++) {
            var previous = uniquePositions[uniquePositions.length - 1];
            if (i === 0 || Math.abs(sortedPositions[i] - previous) >= COORD_TOLERANCE) {
                uniquePositions.push(sortedPositions[i]);
            }
        }
        return uniquePositions;
    }

    /**
     * ガイドを縦・横に分類して座標を返す
     * @param {PathItem[]} guidePaths - 分類するガイドパス
     * @returns {{verticals: number[], horizontals: number[]}} 縦ガイドのX座標（昇順）と横ガイドのY座標（降順）
     */
    function classifyGuidePositions(guidePaths) {
        var verticalXs = [];
        var horizontalYs = [];
        for (var i = 0; i < guidePaths.length; i++) {
            var geometry = getGuideGeometry(guidePaths[i]);
            if (!geometry) continue;
            (geometry.isVertical ? verticalXs : horizontalYs).push(geometry.position);
        }
        /* 縦は左→右、横は上→下の順に並べる / Order verticals left to right and horizontals top to bottom */
        verticalXs.sort(function(a, b) {
            return a - b;
        });
        horizontalYs.sort(function(a, b) {
            return b - a;
        });
        /* 同座標に重なったガイドは1本として扱う（つぶれた長方形を作らない）/ Collapse overlapping guides so no zero-size rectangle is produced */
        return {
            verticals: dedupeSortedPositions(verticalXs),
            horizontals: dedupeSortedPositions(horizontalYs)
        };
    }

    /**
     * オプションに従ってガイドを集め、縦横の座標に分類する
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {{verticals: number[], horizontals: number[]}} 縦ガイドのX座標（昇順）と横ガイドのY座標（降順）
     */
    function getGuidePositions(doc, options) {
        /* ロックされたレイヤー上のガイドも拾えるよう、収集の間だけロックを外す / Unlock during collection so guides on locked layers are visible to the loop */
        var relockTargets = options.includeLocked ? unlockLockedLayers(doc) : [];
        var guidePaths = collectGuidePaths(doc, options);
        relockLayers(relockTargets);
        return classifyGuidePositions(guidePaths);
    }

    // =========================================
    // 長方形の生成 / Creating the rectangles
    // =========================================

    /**
     * 長方形の塗り色を生成する（RGBは黒、それ以外はK100）
     * @param {DocumentColorSpace} docColorSpace - ドキュメントのカラースペース
     * @returns {RGBColor|CMYKColor} 塗り色
     */
    function createFillColor(docColorSpace) {
        if (docColorSpace === DocumentColorSpace.RGB) {
            var rgbColor = new RGBColor();
            rgbColor.red = 0;
            rgbColor.green = 0;
            rgbColor.blue = 0;
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    /**
     * ガイドの座標から、作成される長方形の数を求める
     * @param {{verticals: number[], horizontals: number[]}} guidePositions - 縦横ガイドの座標
     * @returns {number} 作成される長方形の数（足りなければ0）
     */
    function countRectangles(guidePositions) {
        var verticalCount = guidePositions.verticals.length;
        var horizontalCount = guidePositions.horizontals.length;
        if (verticalCount < 2 || horizontalCount < 2) {
            return 0;
        }
        return (verticalCount - 1) * (horizontalCount - 1);
    }

    /**
     * 縦横ガイドで区切られた区画それぞれに長方形を作成する
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} verticalXs - 縦ガイドのX座標（昇順）
     * @param {number[]} horizontalYs - 横ガイドのY座標（降順）
     * @param {number} offsetPt - 四辺を広げる量（pt、マイナスで縮む）
     * @returns {PathItem[]} 作成した長方形
     */
    function createRectangles(doc, verticalXs, horizontalYs, offsetPt) {
        /* 色は使い回す（アイテムごとに新規生成する必要はない）/ Reuse one color object for every rectangle */
        var fillColor = createFillColor(doc.documentColorSpace);
        var createdRects = [];
        for (var xi = 0; xi < verticalXs.length - 1; xi++) {
            for (var yi = 0; yi < horizontalYs.length - 1; yi++) {
                var rectLeft = verticalXs[xi] - offsetPt;
                var rectTop = horizontalYs[yi] + offsetPt;
                var rectWidth = verticalXs[xi + 1] - verticalXs[xi] + offsetPt * 2;
                var rectHeight = horizontalYs[yi] - horizontalYs[yi + 1] + offsetPt * 2;
                /* マイナスのオフセットで潰れる区画は作らない / Skip a cell that a negative offset collapses */
                if (rectWidth <= 0 || rectHeight <= 0) continue;
                var rectangle = doc.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
                rectangle.stroked = false;
                rectangle.filled = true;
                rectangle.fillColor = fillColor;
                rectangle.opacity = RECT_OPACITY;
                createdRects.push(rectangle);
            }
        }
        return createdRects;
    }

    /**
     * 指定したアイテムだけを選択状態にする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 選択するアイテム
     * @returns {void}
     */
    function selectOnly(doc, items) {
        doc.selection = null;
        for (var i = 0; i < items.length; i++) {
            items[i].selected = true;
        }
    }

    /**
     * 選択中の長方形を1つのパスに結合する
     * @returns {void}
     */
    function mergeSelectedRectangles() {
        app.executeMenuCommand('group');
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');
        app.executeMenuCommand('ungroup');
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {object} GuideRectOptions
     * @property {boolean} useRulerOnly - ルーラーガイドのみを対象にするなら true
     * @property {boolean} activeLayerOnly - 現在のレイヤーのガイドだけを対象にするなら true
     * @property {boolean} activeArtboardOnly - 現在のアートボード内のガイドだけを対象にするなら true
     * @property {boolean} includeLocked - ロックされたレイヤー上のガイドも含めるなら true
     * @property {boolean} useOutputLayer - 指定レイヤーに長方形を作成するなら true（false なら現在のレイヤー）
     * @property {string} outputLayerName - 指定レイヤーの名前（同名があれば再利用）
     * @property {number} offsetPt - 各長方形の四辺を広げる量（pt、マイナスで縮む。オフセット未使用なら0）
     * @property {boolean} mergeAdjacent - 隣り合う長方形を結合するなら true
     * @property {boolean} convertToShape - 作成した長方形をライブシェイプに変換するなら true
     */

    /**
     * オプションダイアログを表示する
     * @param {Document} doc - 作成予定数の集計に使うドキュメント
     * @returns {GuideRectOptions|null} 設定値（キャンセル時は null）
     */
    function showOptionsDialog(doc) {
        var dialog = createDialogWindow(getLabel('dialog.title') + ' ' + SCRIPT_VERSION);

        /* 上段：左＝対象となるガイド、右＝作成する長方形 / Top area: target guides on the left, rectangles on the right */
        var columnsRow = dialog.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["fill", "top"];
        columnsRow.spacing = COLUMN_SPACING;

        var guideSourcePanel = addPanel(columnsRow, getLabel('guideSource.panelTitle'));
        /* 各カラムは中身の高さのまま、上端をそろえて並べる / Each column keeps its own height and lines up at the top */
        guideSourcePanel.alignment = ["fill", "top"];

        /* 見つかったガイドの本数（絞り込みの結果がすぐ分かるよう最上部に置く）/ How many guides matched, kept at the top so the filter result is visible at once */
        var guideCountText = addPanelMessage(guideSourcePanel, getLabel('tooltip.guideCount'));

        var guideTypePanel = addPanel(guideSourcePanel, getLabel('guideType.panelTitle'));
        var guideTypeRadios = addRadioPair(guideTypePanel,
            getLabel('guideType.all'), getLabel('guideType.rulerOnly'), 0,
            getLabel('tooltip.guideType'));
        var rulerGuidesRadio = guideTypeRadios[1];

        var targetLayerPanel = addPanel(guideSourcePanel, getLabel('targetLayer.panelTitle'));
        var targetLayerRadios = addRadioPair(targetLayerPanel,
            getLabel('targetLayer.allLayers'), getLabel('targetLayer.activeOnly'), 0,
            getLabel('tooltip.targetLayer'));
        var activeLayerRadio = targetLayerRadios[1];

        var guideOptionPanel = addPanel(guideSourcePanel, getLabel('guideOption.panelTitle'));
        var activeArtboardCheckbox = addPanelCheckbox(guideOptionPanel,
            getLabel('guideOption.activeArtboardOnly'), false, getLabel('tooltip.activeArtboardOnly'));
        var includeLockedCheckbox = addPanelCheckbox(guideOptionPanel,
            getLabel('guideOption.includeLocked'), true, getLabel('tooltip.includeLocked'));

        var rectanglePanel = addPanel(columnsRow, getLabel('rectangle.panelTitle'));
        rectanglePanel.alignment = ["fill", "top"];

        /* 作成予定数（左カラムのガイド本数と同じく、パネルの最上部に置く）/ Expected count, kept at the top like the guide count in the left column */
        var summaryText = addPanelMessage(rectanglePanel, getLabel('tooltip.summaryCount'));

        var destinationPanel = addPanel(rectanglePanel, getLabel('destination.panelTitle'));
        var destinationRadios = addRadioPair(destinationPanel,
            getLabel('destination.activeLayer'), getLabel('destination.outputLayer'), 0,
            getLabel('tooltip.destination'));
        var outputLayerRadio = destinationRadios[1];

        /* 「指定レイヤー」の名前はラジオの次の行に字下げして置く / The layer name sits on the line below its radio, indented */
        var outputLayerNameRow = destinationPanel.add("group");
        setupRow(outputLayerNameRow, "fill", FIELD_ROW_SPACING);
        outputLayerNameRow.margins = INDENT_MARGINS;
        var outputLayerNameInput = outputLayerNameRow.add('edittext {characters: ' + LAYER_NAME_CHARS + '}');
        outputLayerNameInput.text = OUTPUT_LAYER_NAME;
        /* 余った幅で伸ばす（characters を増やすとダイアログごと広がる）/ Stretch into the leftover width; raising characters would widen the dialog */
        outputLayerNameInput.alignment = ["fill", "center"];
        outputLayerNameInput.helpTip = getLabel('tooltip.outputLayerName');

        var rectangleOptionPanel = addPanel(rectanglePanel, getLabel('rectangleOption.panelTitle'));

        /* オフセット：チェックボックス＋数値欄＋ルーラー単位 / Offset: checkbox, value field, and the ruler unit */
        var rulerUnit = getUnitInfo("rulerType");
        var offsetRow = rectangleOptionPanel.add("group");
        setupRow(offsetRow, "left", FIELD_ROW_SPACING);
        var offsetCheckbox = offsetRow.add("checkbox", undefined, labelText('rectangleOption.offset'));
        offsetCheckbox.helpTip = getLabel('tooltip.offset');
        /* ∧∨と入力欄は隙間0で突き合わせる。マイナス値（縮小）も使えるので下限は付けない
           Butt the stepper against the field; no minimum, since negative values shrink the rectangles */
        var offsetStepperGroup = offsetRow.add("group");
        setupRow(offsetStepperGroup, "left", 0);
        offsetStepperGroup.margins = 0;
        var offsetInput;
        var offsetStepper = addStepper(offsetStepperGroup, function () { return offsetInput; }, {
            onStep: function () { refresh(); }
        });
        offsetInput = offsetStepperGroup.add('edittext {characters: ' + OFFSET_CHARS + '}');
        offsetInput.text = "0";
        offsetInput.helpTip = getLabel('tooltip.offset');
        bindSteppedArrowKeys(offsetInput, offsetStepper);
        offsetRow.add("statictext", undefined, rulerUnit.label);

        var mergeAdjacentCheckbox = addPanelCheckbox(rectangleOptionPanel,
            getLabel('rectangleOption.mergeAdjacent'), false, getLabel('tooltip.mergeAdjacent'));
        var convertToShapeCheckbox = addPanelCheckbox(rectangleOptionPanel,
            getLabel('rectangleOption.convertToShape'), true, getLabel('tooltip.convertToShape'));

        /**
         * 指定レイヤーを選んでいるときだけ名前欄を使えるようにする
         * @returns {void}
         */
        function syncOutputLayerField() {
            outputLayerNameInput.enabled = outputLayerRadio.value;
        }

        /**
         * オフセットにチェックが入っているときだけ数値欄を使えるようにする
         * @returns {void}
         */
        function syncOffsetField() {
            var isEnabled = offsetCheckbox.value;
            offsetInput.enabled = isEnabled;
            /* ∧∨は切り替わったときだけ描き直す（refresh のたびに呼ばれるため） / redraw the stepper only when its state flips */
            if (offsetStepper.enabled !== isEnabled) {
                offsetStepper.enabled = isEnabled;
                redrawSteppersIn(offsetStepper);
            }
        }

        /**
         * 現在の入力内容を設定値として読み取る
         * @returns {GuideRectOptions} 設定値
         */
        function readOptions() {
            return {
                useRulerOnly: rulerGuidesRadio.value,
                activeLayerOnly: activeLayerRadio.value,
                activeArtboardOnly: activeArtboardCheckbox.value,
                includeLocked: includeLockedCheckbox.value,
                useOutputLayer: outputLayerRadio.value,
                outputLayerName: normalizeLayerName(outputLayerNameInput.text),
                offsetPt: offsetCheckbox.value ? convertToPt(offsetInput.text, rulerUnit.pointsPerUnit) : 0,
                mergeAdjacent: mergeAdjacentCheckbox.value,
                convertToShape: convertToShapeCheckbox.value
            };
        }

        /**
         * ガイドの本数と作成予定数を各パネルの見出しに反映する
         * @param {GuideRectOptions} options - ダイアログの設定値
         * @param {{verticals: number[], horizontals: number[]}} guidePositions - 縦横ガイドの座標
         * @param {number} rectangleCount - 作成される長方形の数
         * @returns {void}
         */
        function updateSummary(options, guidePositions, rectangleCount) {
            guideCountText.text = getLabel('summary.guides', {
                vertical: guidePositions.verticals.length,
                horizontal: guidePositions.horizontals.length
            });
            /* 判定は main() の結合条件と同じにそろえる / Mirror the merge condition used in main() */
            var willMerge = options.mergeAdjacent && rectangleCount > 1;
            summaryText.text = getLabel(willMerge ? 'summary.merged' : 'summary.count', {
                count: rectangleCount
            });
        }

        var buttonRow = addButtonRow(dialog);
        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel('dialog.preview'));
        previewCheckbox.value = true;
        previewCheckbox.helpTip = getLabel('tooltip.preview');

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });

        /* プレビューで描いた長方形と、そのために新規作成したレイヤー / Previewed rectangles and any layer created just for them */
        var previewRects = null;
        var previewLayer = null;

        /**
         * プレビューの長方形を取り除く
         * @returns {void}
         */
        function clearPreview() {
            if (!previewRects) return;
            for (var i = 0; i < previewRects.length; i++) {
                var previewRect = previewRects[i];
                if (previewRect && !previewRect.locked && previewRect.layer && !previewRect.layer.locked) {
                    previewRect.remove();
                }
            }
            previewRects = null;
        }

        /**
         * プレビューのために作ったレイヤーを、空のままなら片付ける
         * @returns {void}
         */
        function removePreviewLayer() {
            if (!previewLayer) return;
            /* 最後の1枚は削除できないので残す / The last layer cannot be removed, so leave it */
            if (previewLayer.pageItems.length === 0 && doc.layers.length > 1) {
                previewLayer.remove();
            }
            previewLayer = null;
        }

        /**
         * 現在の設定でプレビューを描き直す
         * @param {GuideRectOptions} options - ダイアログの設定値
         * @param {{verticals: number[], horizontals: number[]}} guidePositions - 縦横ガイドの座標
         * @param {number} rectangleCount - 作成される長方形の数
         * @returns {void}
         */
        function updatePreview(options, guidePositions, rectangleCount) {
            clearPreview();
            /* 結合・シェイプ変換はメニューコマンドで重いため、プレビューでは適用しない
               Merging and shape conversion run as menu commands, so the preview leaves them out */
            if (previewCheckbox.value && rectangleCount > 0 && rectangleCount <= PREVIEW_MAX_RECTS) {
                var outputTarget = activateOutputLayer(doc, options);
                if (outputTarget) {
                    if (outputTarget.created) {
                        previewLayer = outputTarget.layer;
                    }
                    previewRects = createRectangles(doc, guidePositions.verticals, guidePositions.horizontals, options.offsetPt);
                }
            }
            app.redraw();
        }

        /**
         * 設定を読み直し、件数表示とプレビューをまとめて更新する
         * @returns {void}
         */
        function refresh() {
            /* 入力欄の有効・無効もここでまとめて反映する（ハンドラーを1本にして取りこぼしを防ぐ）
               Enable/disable the fields here too, so a single handler covers everything */
            syncOutputLayerField();
            syncOffsetField();
            var options = readOptions();
            var guidePositions = getGuidePositions(doc, options);
            var rectangleCount = countRectangles(guidePositions);
            updateSummary(options, guidePositions, rectangleCount);
            updatePreview(options, guidePositions, rectangleCount);
        }

        /* 結果の見込みが変わる操作はすべて反映し直す / Refresh after anything that changes the expected result */
        var refreshControls = guideTypeRadios.concat(targetLayerRadios, destinationRadios);
        refreshControls.push(activeArtboardCheckbox, includeLockedCheckbox, mergeAdjacentCheckbox,
            offsetCheckbox, previewCheckbox);
        for (var i = 0; i < refreshControls.length; i++) {
            /* onClick と onChange の両方に付ける（チェックボックスはどちらで届くかが環境で異なる）
               Bind both events; which one a checkbox delivers varies by platform */
            refreshControls[i].onClick = refresh;
            refreshControls[i].onChange = refresh;
        }
        offsetInput.addEventListener("changing", refresh);
        outputLayerNameInput.addEventListener("changing", refresh);
        refresh();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();
        /* 本番の長方形は main() が作り直すので、プレビューは必ず片付ける / main() draws the real rectangles, so the preview always goes away */
        clearPreview();
        if (dialogResult !== 1) {
            removePreviewLayer();
            app.redraw();
            return null;
        }
        return readOptions();
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * ガイドの交点で区切られた区画に長方形を作成する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;
        var options = showOptionsDialog(doc);
        if (!options) return;

        var guidePositions = getGuidePositions(doc, options);
        if (countRectangles(guidePositions) === 0) {
            alert(getLabel('alert.notEnoughGuides'));
            return;
        }

        if (!prepareOutputLayer(doc, options)) return;

        /* selectOnly が元の選択を解除するので、後続のメニューコマンドは生成分にだけ効く / selectOnly clears the previous selection, so the menu commands below act only on the new rectangles */
        var rectangles = createRectangles(doc, guidePositions.verticals, guidePositions.horizontals, options.offsetPt);
        if (rectangles.length === 0) return;
        selectOnly(doc, rectangles);

        if (options.mergeAdjacent && rectangles.length > 1) {
            mergeSelectedRectangles();
        }
        if (options.convertToShape) {
            app.executeMenuCommand('Convert to Shape');
        }
    }

    main();

})();
