#target illustrator
#targetengine "GenerateGuidesGridEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボードまたは選択オブジェクトの外接矩形を、指定した行数・列数に分割してグリッド用のガイドを生成します。
アートボードのエッジ、セルの長方形化（角丸・中心点の表示）、現在の設定のプリセット書き出しにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenerateGuidesGrid.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n7adc7290b607

### Overview

Divides the artboard, or the bounding box of the selection, into the specified rows and columns and generates grid guides.
It can also draw the artboard edges, draw the cells as rectangles (with round corners and center points), and export the current settings as a preset.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GenerateGuidesGrid.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GenerateGuidesGrid";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenerateGuidesGrid.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GenerateGuidesGrid.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n7adc7290b607"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var GUIDE_LAYER_NAME = "grid_guides";        /* ガイドを格納するレイヤー名 / Layer name for guides */
    var CELL_LAYER_NAME = "cell-rectangle";      /* セル長方形を格納するレイヤー名 / Layer name for cell rectangles */
    var PREVIEW_LAYER_NAME = "_Preview_Guides";  /* プレビュー用レイヤー名 / Layer name for live preview */
    var CELL_OPACITY = 15;                        /* セル長方形の不透明度（%）/ Cell rectangle opacity (%) */
    var RECT_TOLERANCE = 0.5;                    /* アートボード内外判定の許容値（pt）/ Tolerance for the inside-artboard test (pt) */

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

    var ROW_SPACING    = 8;                  /* 行内の要素間隔 / gap inside a row */

    /**
     * 左に∧∨を付けた数値入力欄を追加する（∧∨と入力欄は隙間0で突き合わせる）。
     * ↑↓キーも∧∨と同じ処理で増減し、増減後は入力欄の onChanging を呼んで連動・プレビューを反映する
     * @param {Group} parentRow - 追加先の行
     * @param {string} defaultText - 初期値
     * @param {number} characters - 入力欄の文字数幅
     * @param {Object} stepOptions - min / integer など（addStepper() の stepOptions）
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addStepperInput(parentRow, defaultText, characters, stepOptions) {
        var stepperFieldGroup = parentRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        stepOptions.onStep = function (numberInput) {
            /* プログラム変更では onChanging が発火しないため明示的に呼ぶ / programmatic changes do not fire onChanging */
            if (typeof numberInput.onChanging === "function") numberInput.onChanging();
        };
        var numberInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperFieldGroup.add("edittext", undefined, defaultText);
        numberInput.characters = characters;
        numberInput.stepperGroup = stepperGroup;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * addStepperInput() で作った入力欄の有効／無効を、∧∨ごと切り替える
     * @param {EditText} numberInput - 入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup); /* ∧∨は自作描画なので描き直す / redraw the custom-drawn buttons */
    }

    /**
     * グループの共通設定（row は縦中央、column は左揃え）
     * @param {Group} group - 対象グループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - 要素間隔（省略時は ROW_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
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

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () {
            numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
            numberInput.active = true;
        });

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

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English label definitions (by category) */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "グリッドに分割 Pro", en: "Split into Grid Pro" }
        },
        /* 対象 / Target */
        target: {
            selection: { ja: "選択オブジェクト", en: "Selected Object(s)" },
            artboard: { ja: "アートボード", en: "Artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All Artboards" }
        },
        /* パネル見出し / Panel titles */
        panel: {
            target: { ja: "対象", en: "Target" },
            row: { ja: "行", en: "Row" },
            column: { ja: "列", en: "Column" },
            margin: { ja: "マージン設定", en: "Margin Settings" },
            options: { ja: "セル", en: "Cell" },
            guides: { ja: "ガイド", en: "Guides" },
            originalObject: { ja: "元のオブジェクト", en: "Original Object(s)" }
        },
        /* ラジオボタン / Radio buttons */
        radio: {
            remove: { ja: "削除する", en: "Delete" },
            keep: { ja: "そのまま", en: "Keep" },
            toGuide: { ja: "ガイド化", en: "Make Guides" }
        },
        /* チェックボックス / Checkboxes */
        checkbox: {
            linkGutter: { ja: "行間に連動", en: "Link to Row Gutter" },
            cellRect: { ja: "長方形化", en: "Rectangles" },
            showCenter: { ja: "中心点を表示", en: "Show Center Point" },
            roundCorner: { ja: "角丸", en: "Round Corners" },
            splitCell: { ja: "各セルを左右分割", en: "Split Each Cell Horizontally" },
            drawGuides: { ja: "ガイドを引く", en: "Draw Guides" },
            artboardEdge: { ja: "アートボードのエッジ", en: "Artboard Edges" },
            clearGuides: { ja: "既存ガイドを削除", en: "Clear Existing Guides" }
        },
        /* フィールド見出し（コロンは labelText で付与）/ Field labels (colon added by labelText) */
        field: {
            preset: { ja: "プリセット", en: "Preset" },
            rowCount: { ja: "行数", en: "Number" },
            rowGutter: { ja: "行間", en: "Gutter" },
            columnCount: { ja: "列数", en: "Number" },
            columnGutter: { ja: "列間", en: "Gutter" },
            top: { ja: "上", en: "Top" },
            left: { ja: "左", en: "Left" },
            bottom: { ja: "下", en: "Bottom" },
            right: { ja: "右", en: "Right" },
            guideExtension: { ja: "伸張", en: "Extension" },
            opacity: { ja: "不透明度", en: "Opacity" }
        },
        /* ツールチップ / Tooltips (helpTip) */
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
            linkGutter: { ja: "列間を行間と同じ値に保ちます。", en: "Keep the column gutter equal to the row gutter." },
            linkMargin: { ja: "上の値を下・左・右にも適用します。", en: "Apply the top value to bottom, left, and right." },
            cellRect: { ja: "各セルを長方形として作成します。", en: "Create a rectangle for each cell." },
            showCenter: { ja: "作成した長方形の中心点を属性パネルで表示します（OK時に適用）。", en: "Show the center point of created rectangles in the Attributes panel (applied on OK)." },
            roundCorner: { ja: "各長方形に角丸（ライブエフェクト）を適用します。", en: "Apply the Round Corners live effect to each rectangle." },
            opacity: { ja: "長方形の不透明度（0〜100%）。", en: "Opacity of the rectangles (0–100%)." },
            splitCell: { ja: "各セルの左右中央に縦ガイドを作成します。", en: "Create a vertical guide at each cell's horizontal center." },
            guideExtension: { ja: "アートボード／オブジェクトの外側へガイドを伸ばす距離。", en: "Distance to extend guides beyond the artboard/object." },
            artboardEdge: { ja: "アートボードの上下左右4辺にガイドを引きます。対象が選択オブジェクトのときは使えません。", en: "Draw guides on the four edges of the artboard. Unavailable when the target is the selected objects." },
            clearGuides: { ja: "描画前に専用レイヤー（grid_guides）内の既存ガイドを削除します。「すべてのアートボード」以外は、アクティブなアートボード上のガイドだけが対象です。", en: "Remove existing guides in the dedicated layer (grid_guides) before drawing. Unless all artboards are targeted, only guides on the active artboard are removed." },
            toGuide: { ja: "元の選択オブジェクトをガイドに変換します。", en: "Convert the original selected objects into guides." },
            outline: { ja: "アウトライン表示とプレビュー表示を切り替えます。", en: "Toggle between Outline and Preview view." },
            exportPreset: { ja: "現在の設定をプリセットファイルに書き出します。", en: "Export the current settings to a preset file." }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" },
            ok: { ja: "OK", en: "OK" },
            outline: { ja: "アウトライン表示", en: "Outline" },
            preview: { ja: "プレビュー表示", en: "Preview" },
            exportPreset: { ja: "書き出し", en: "Export" }
        }
    };

    /* 見出しに単位を付与（日本語は全角括弧、英語は半角括弧）/ Append unit to a title (full-width parens JA, half-width EN) */
    function titleWithUnit(key) {
        return getLabel(key) + (uiLang === "ja" ? "（" + unitLabel + "）" : " (" + unitLabel + ")");
    }

    // =========================================
    // 単位 / Units
    // =========================================
    /* rulerType から単位ラベルと pt 換算係数を求める / Resolve unit label and pt factor from rulerType */
    function getUnitInfo(rulerType) {
        switch (rulerType) {
            case 0: return { label: "inch", factor: 72.0 };
            case 1: return { label: "mm", factor: 72.0 / 25.4 };
            case 3: return { label: "pica", factor: 12.0 };
            case 4: return { label: "cm", factor: 72.0 / 2.54 };
            case 5: return { label: "Q", factor: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", factor: 1.0 };
            default: return { label: "pt", factor: 1.0 }; // 2 = pt
        }
    }
    var unitInfo = getUnitInfo(app.preferences.getIntegerPreference("rulerType"));
    var unitLabel = unitInfo.label;
    var unitFactor = unitInfo.factor;

    /* プリセット値は pt 基準で保持。表示は現在単位へ換算する / Preset values are stored in points; convert to the current ruler unit for display */
    var UNIT_DECIMALS = 2; /* 換算時に残す小数桁数 / decimals kept when converting */

    /**
     * 指定桁数で丸める
     * @param {number} value - 対象の値
     * @param {number} decimals - 残す小数桁数
     * @returns {number} 丸めた値
     */
    function roundTo(value, decimals) {
        var digitScale = Math.pow(10, decimals);
        return Math.round(Number(value) * digitScale) / digitScale;
    }

    /**
     * pt → 現在単位
     * @param {number} ptValue - pt 値
     * @returns {number} 現在単位の値
     */
    function ptToUnit(ptValue) {
        return roundTo(Number(ptValue) / unitFactor, UNIT_DECIMALS);
    }

    /**
     * 現在単位 → pt
     * @param {number|string} unitValue - 現在単位の値
     * @returns {number} pt 値
     */
    function unitToPt(unitValue) {
        return roundTo(Number(unitValue) * unitFactor, UNIT_DECIMALS);
    }

    /**
     * 入力欄のテキストを数値として読む（空欄・不正値は fallback）
     * @param {string} inputText - 入力文字列
     * @param {number} [fallbackValue] - 数値にならないときの値（省略時は 0）
     * @returns {number} 数値
     */
    function toNumber(inputText, fallbackValue) {
        var parsedValue = parseFloat(inputText);
        if (isNaN(parsedValue)) return (typeof fallbackValue === "number") ? fallbackValue : 0;
        return parsedValue;
    }

    /**
     * 入力欄のテキストを整数として読む（空欄・不正値は fallback）
     * @param {string} inputText - 入力文字列
     * @param {number} [fallbackValue] - 整数にならないときの値（省略時は 0）
     * @returns {number} 整数
     */
    function toInteger(inputText, fallbackValue) {
        var parsedValue = parseInt(inputText, 10);
        if (isNaN(parsedValue)) return (typeof fallbackValue === "number") ? fallbackValue : 0;
        return parsedValue;
    }

    // =========================================
    // プリセット / Presets
    // =========================================
    // （drawGuides を含む）/ (includes drawGuides)
    var presets = [
        {
            label: "1行2列",
            columns: 2,
            rows: 1,
            guideExtension: 50,
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 0,
            marginRight: 0,
            rowGutter: 0,
            columnGutter: 0,
            drawCells: true,
            drawGuides: true
        }, {
            label: "1つの図形",
            columns: 1,
            rows: 1,
            guideExtension: 10,
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 0,
            marginRight: 0,
            rowGutter: 0,
            columnGutter: 0,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "十字 / Cross",
            columns: 2,
            rows: 2,
            guideExtension: 0,
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 0,
            marginRight: 0,
            rowGutter: 0,
            columnGutter: 0,
            drawCells: false,
            drawGuides: true
        },
        {
            label: "シングル / Single",
            columns: 1,
            rows: 1,
            guideExtension: 50,
            marginTop: 100,
            marginBottom: 100,
            marginLeft: 100,
            marginRight: 100,
            rowGutter: 0,
            columnGutter: 0,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "2行×2列 / 2 Rows × 2 Columns",
            columns: 2,
            rows: 2,
            guideExtension: 20,
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 0,
            marginRight: 0,
            rowGutter: 50,
            columnGutter: 50,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "1行×3列 / 1 Row × 3 Columns",
            columns: 3,
            rows: 1,
            guideExtension: 0,
            marginTop: 30,
            marginBottom: 30,
            marginLeft: 30,
            marginRight: 30,
            rowGutter: 0,
            columnGutter: 30,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "4行×4列 / 4 Rows × 4 Columns",
            columns: 4,
            rows: 4,
            guideExtension: 0,
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 0,
            marginRight: 0,
            rowGutter: 20,
            columnGutter: 20,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "2行×3列 / 2 Rows × 3 Columns",
            columns: 3,
            rows: 2,
            guideExtension: 0,
            marginTop: 100,
            marginBottom: 100,
            marginLeft: 100,
            marginRight: 100,
            rowGutter: 20,
            columnGutter: 20,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "3行×3列 / 3 Rows × 3 Columns",
            columns: 3,
            rows: 3,
            guideExtension: 0,
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 200,
            marginRight: 0,
            rowGutter: 0,
            columnGutter: 0,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "sp / sp",
            columns: 1,
            rows: 1,
            guideExtension: 0,
            marginTop: 220,
            marginBottom: 220,
            marginLeft: 0,
            marginRight: 0,
            rowGutter: 0,
            columnGutter: 0,
            drawCells: true,
            drawGuides: true
        },
        {
            label: "長方形のみ / just rectangle",
            columns: 1,
            rows: 1,
            guideExtension: 10,
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 0,
            marginRight: 0,
            rowGutter: 0,
            columnGutter: 0,
            drawCells: true,
            drawGuides: false
        }
    ];

    // =========================================
    // プレビュー管理 / Preview manager
    // =========================================
    // プレビュー時にapp.undo()で巻き戻して履歴を汚さない / Manage preview with rollback using app.undo()
    function PreviewManager() {
        this.undoDepth = 0; // number of preview actions applied
        this.errorReported = false; // 同じ失敗を何度も知らせない / report a failure only once
    }

    // プレビュー手順を1つ実行してカウント / Run an action as a preview step and count it
    // step は何も変更しなければ false を返し、変更するなら最初の変更の直前に markChanged() を呼ぶ。
    // 例外で抜けた手順は markChanged() 済みのときだけ数える（未変更の手順を数えると、余分な undo でユーザー自身の作業が巻き戻る）
    // A step returns false when it changes nothing, or calls markChanged() just before its first change.
    // A step that threw is counted only if it had marked: counting an unchanged step would roll back the user's own work
    PreviewManager.prototype.addStep = function (step) {
        var changed = false;
        function markChanged() {
            changed = true;
        }
        try {
            changed = (step(markChanged) !== false);
        } catch (e) {
            // changed には markChanged() が呼ばれたかどうかが残る / changed still holds whether markChanged() was called
            $.writeln("[GenerateGuidesGrid] preview step error: " + e);
            if (!this.errorReported) {
                this.errorReported = true;
                alert("プレビューの処理に失敗しました。\nPreview step failed.\n" + e);
            }
        }
        if (changed) this.undoDepth++;
        app.redraw();
    };

    // すべてのプレビュー手順を巻き戻す / Roll back all preview actions
    PreviewManager.prototype.rollback = function () {
        while (this.undoDepth > 0) {
            try {
                app.undo();
            } catch (e) {
                break;
            }
            this.undoDepth--;
        }
        app.redraw();
    };

    // 現在の状態を確定。finalAction があれば巻き戻してから1回だけ実行 / Confirm; if finalAction is given, rollback then run it once
    PreviewManager.prototype.confirm = function (finalAction) {
        if (finalAction) {
            this.rollback();
            finalAction();
        } else {
            this.undoDepth = 0;
        }
    };

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    // =========================================
    // 描画ヘルパー（doc を受け取り、UI/クロージャに依存しない）/ Drawing helpers (take doc; no UI/closure deps)
    // =========================================

    /**
     * 外接矩形（geometricBounds）を持つオブジェクトか
     * @param {object} item - 判定するオブジェクト
     * @returns {boolean} 座標を取得できるなら true
     */
    function hasGeometricBounds(item) {
        try {
            return !!(item && item.geometricBounds && item.geometricBounds.length === 4);
        } catch (e) {
            return false;
        }
    }

    /**
     * オブジェクトの中心が矩形の中にあるか（ガイドは矩形の外へ伸びるので中心で判定）
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Array} boundsRect - 矩形 [左, 上, 右, 下]
     * @returns {boolean} 中心が矩形内なら true
     */
    function isCenterInsideRect(item, boundsRect) {
        if (!hasGeometricBounds(item)) return false;
        var itemBounds = item.geometricBounds;
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        return (centerX >= boundsRect[0] - RECT_TOLERANCE && centerX <= boundsRect[2] + RECT_TOLERANCE &&
            centerY <= boundsRect[1] + RECT_TOLERANCE && centerY >= boundsRect[3] - RECT_TOLERANCE);
    }

    /**
     * 関数を実行し、例外を握りつぶす（try/catch の重複を集約）
     * @param {function} action - 実行する処理
     * @returns {void}
     */
    function safeExecute(action) {
        try {
            action();
        } catch (e) {
            $.writeln("[GenerateGuidesGrid] safeExecute error: " + e);
        }
    }

    /**
     * source の自前プロパティを target にコピーする
     * @param {object} target - コピー先
     * @param {object} source - コピー元
     * @returns {object} コピー先（target）
     */
    function mergeInto(target, source) {
        for (var key in source) {
            if (source.hasOwnProperty(key)) target[key] = source[key];
        }
        return target;
    }

    /**
     * レイヤーのロックを安全に解除する
     * @param {Layer} layer - 対象レイヤー
     * @returns {void}
     */
    function safeUnlockLayer(layer) {
        safeExecute(function () {
            if (layer && layer.locked) layer.locked = false;
        });
    }

    /**
     * 指定名のレイヤーを安全に削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function safeRemoveLayerByName(doc, layerName) {
        safeExecute(function () {
            var layer = doc.layers.getByName(layerName);
            if (layer) layer.remove();
        });
    }

    /**
     * 指定名のレイヤーを取得する。なければ作成する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @param {object} [outInfo] - 新規作成したときに created = true が入るオブジェクト
     * @returns {Layer} 取得または作成したレイヤー
     */
    function getOrCreateLayer(doc, layerName, outInfo) {
        var layer;
        try {
            layer = doc.layers.getByName(layerName);
        } catch (e) {
            layer = doc.layers.add();
            layer.name = layerName;
            if (outInfo) outInfo.created = true;
        }
        return layer;
    }

    /**
     * ガイド線を1本追加する（塗り・線なしのパスをガイド化してレイヤー先頭へ）
     * アクティブレイヤーがロック・非表示・テンプレートでも失敗しないよう、追加先レイヤーに直接作る
     * @param {Layer} layer - 追加先レイヤー
     * @param {Array} startPoint - 始点 [x, y]
     * @param {Array} endPoint - 終点 [x, y]
     * @returns {PathItem} 追加したガイド
     */
    function addGuideLine(layer, startPoint, endPoint) {
        var guideLine = layer.pathItems.add();
        guideLine.setEntirePath([startPoint, endPoint]);
        guideLine.stroked = false;
        guideLine.filled = false;
        guideLine.guides = true;
        return guideLine;
    }

    /**
     * 角丸ライブエフェクトのXMLを作る
     * @param {number} radius - 角丸半径（pt）
     * @returns {string} ライブエフェクトのXML
     */
    function roundCornersEffectXML(radius) {
        var effectXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return effectXml.replace('#value#', radius);
    }

    /**
     * 黒色を作る（CMYK／RGB対応）
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラーオブジェクト
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 100;
            return cmyk;
        }
        var rgb = new RGBColor();
        rgb.red = 0;
        rgb.green = 0;
        rgb.blue = 0;
        return rgb;
    }

    /**
     * 描画できるコンテキストか（行数・列数が正か）
     * @param {object} drawContext - 描画コンテキスト
     * @returns {boolean} 描画できるなら true
     */
    function canDrawGrid(drawContext) {
        return !(isNaN(drawContext.columnCount) || drawContext.columnCount <= 0 || isNaN(drawContext.rowCount) || drawContext.rowCount <= 0);
    }

    // グリッド（ガイド＋セル長方形）を描画し、{ cells: 作成したセル長方形, changed: ドキュメントを変更したか } を返す
    // Draw the grid; return { cells: created cell rects, changed: whether the document was changed }
    function drawGrid(drawContext) {
        var doc = drawContext.doc;
        var isPreview = drawContext.isPreview;
        var columnCount = drawContext.columnCount;
        var rowCount = drawContext.rowCount;
        if (!canDrawGrid(drawContext)) return { cells: [], changed: false };

        var guideExtension = drawContext.guideExtension;
        var marginTop = drawContext.marginTop, marginBottom = drawContext.marginBottom, marginLeft = drawContext.marginLeft, marginRight = drawContext.marginRight;
        var rowGutter = drawContext.rowGutter, columnGutter = drawContext.columnGutter;
        var drawCells = drawContext.drawCells, drawGridGuides = drawContext.drawGridGuides, cornerRadius = drawContext.cornerRadius;
        var splitCells = drawContext.splitCells;
        // 選択オブジェクトが対象のときはアートボードの辺を引かない / No artboard edges when the target is the selection
        var drawArtboardEdge = drawContext.drawArtboardEdge && !drawContext.selBounds;
        var cellOpacity = (typeof drawContext.cellOpacity === "number") ? drawContext.cellOpacity : CELL_OPACITY;
        var createdCells = [];

        // プレビューは1枚に全部描く。確定時はガイド用レイヤーはガイド系を描くときだけ作る
        // Preview draws everything on one layer; on commit, create the guide layer only when guides are drawn
        var gridLayerName = isPreview ? PREVIEW_LAYER_NAME : GUIDE_LAYER_NAME;
        var needGridLayer = isPreview ? (drawGridGuides || splitCells || drawCells || drawArtboardEdge) : (drawGridGuides || splitCells || drawArtboardEdge);
        // ドキュメントを変更する直前に呼び出し元へ知らせる（プレビューの undo 回数を正しく数えるため）
        // Tell the caller just before the document is changed, so the preview counts undo steps correctly
        var reportChange = drawContext.onChange || function () {};

        var layerInfo = {};
        var gridLayer = needGridLayer ? getOrCreateLayer(doc, gridLayerName, layerInfo) : null;
        safeUnlockLayer(gridLayer);

        var cellLayer = gridLayer; // プレビューはセルも gridLayer に描く / preview: cells live on gridLayer
        if (!isPreview && drawCells) {
            cellLayer = getOrCreateLayer(doc, CELL_LAYER_NAME, layerInfo);
            safeUnlockLayer(cellLayer);
        }
        if (layerInfo.created) reportChange();

        var guideCount = 0;
        // ガイドを1本引いて本数を数える / Draw one guide and count it
        function drawGuide(startPoint, endPoint) {
            reportChange();
            addGuideLine(gridLayer, startPoint, endPoint);
            guideCount++;
        }

        var targetRects = [];
        if (drawContext.selBounds) {
            targetRects.push(drawContext.selBounds);
        } else if (drawContext.allBoards) {
            for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
                targetRects.push(doc.artboards[artboardIndex].artboardRect);
            }
        } else {
            targetRects.push(doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect);
        }

        // セルの塗り色はセルごとに作り直さず1回だけ用意 / Build the cell fill color once, not per cell
        var cellFillColor = drawCells ? createBlackColor(doc) : null;

        for (var targetIndex = 0; targetIndex < targetRects.length; targetIndex++) {
            var targetRect = targetRects[targetIndex];
            var targetLeft = targetRect[0],
                targetTop = targetRect[1],
                targetRight = targetRect[2],
                targetBottom = targetRect[3];
            var contentLeft = targetLeft + marginLeft;
            var contentRight = targetRight - marginRight;
            var contentTop = targetTop - marginTop;
            var contentBottom = targetBottom + marginBottom;

            var extendedLeft = targetLeft - guideExtension;
            var extendedRight = targetRight + guideExtension;
            var extendedTop = targetTop + guideExtension;
            var extendedBottom = targetBottom - guideExtension;

            var usableWidth = contentRight - contentLeft;
            var usableHeight = contentTop - contentBottom;
            var totalColumnGutter = (columnCount - 1) * columnGutter;
            var totalRowGutter = (rowCount - 1) * rowGutter;
            // マージン・ガターが過大だとセル幅/高さが0以下になるので、この対象にはグリッドを描かない
            // Margins/gutters too large would make cell width/height <= 0, so this target gets no grid
            var canFitCells = (usableWidth - totalColumnGutter > 0 && usableHeight - totalRowGutter > 0);

            // アートボードの上下左右4辺（マージンに関係なく引く）/ The artboard's four edges (independent of the margins)
            if (drawArtboardEdge && gridLayer) {
                // マージン0の辺はグリッド側の外周ガイドと同じ位置になるので引かない。
                // ただしグリッドを描かない対象では重ならないので引く
                // A zero-margin edge coincides with the grid's outer guide, so skip it —
                // but nothing coincides when the grid is skipped, so draw it then
                var gridCoversEdges = drawGridGuides && canFitCells;
                if (!gridCoversEdges || marginTop > 0) {
                    drawGuide([extendedLeft, targetTop], [extendedRight, targetTop]);
                }
                if (!gridCoversEdges || marginBottom > 0) {
                    drawGuide([extendedLeft, targetBottom], [extendedRight, targetBottom]);
                }
                if (!gridCoversEdges || marginLeft > 0) {
                    drawGuide([targetLeft, extendedTop], [targetLeft, extendedBottom]);
                }
                if (!gridCoversEdges || marginRight > 0) {
                    drawGuide([targetRight, extendedTop], [targetRight, extendedBottom]);
                }
            }

            if (!canFitCells) continue;
            var cellWidth = (usableWidth - totalColumnGutter) / columnCount;
            var cellHeight = (usableHeight - totalRowGutter) / rowCount;

            if (drawGridGuides) {
                // ガイド描画（行・列）/ Grid guides (rows and columns)
                // 最終行は contentBottom にちょうど着地するので、末尾に足すと重なる
                // The last row lands exactly on contentBottom, so a trailing guide would overlap
                var lineY = contentTop;
                drawGuide([extendedLeft, lineY], [extendedRight, lineY]);
                for (var j = 0; j < rowCount; j++) {
                    lineY -= cellHeight;
                    drawGuide([extendedLeft, lineY], [extendedRight, lineY]);
                    // ガター0のときは同じ位置に重なるので引かない / Gutter 0 would draw on the same line
                    if (j < rowCount - 1 && rowGutter > 0) {
                        lineY -= rowGutter;
                        drawGuide([extendedLeft, lineY], [extendedRight, lineY]);
                    }
                }

                var lineX = contentLeft;
                drawGuide([lineX, extendedTop], [lineX, extendedBottom]);
                for (var k = 0; k < columnCount; k++) {
                    lineX += cellWidth;
                    drawGuide([lineX, extendedTop], [lineX, extendedBottom]);
                    // ガター0のときは同じ位置に重なるので足さない / Gutter 0 would draw on the same line
                    if (k < columnCount - 1 && columnGutter > 0) {
                        lineX += columnGutter;
                        drawGuide([lineX, extendedTop], [lineX, extendedBottom]);
                    }
                }
            }

            if (drawCells && cellLayer) {
                var cellOriginX = contentLeft;
                var cellOriginY = contentTop;
                for (var row = 0; row < rowCount; row++) {
                    var cellY = cellOriginY - (cellHeight + rowGutter) * row;
                    for (var column = 0; column < columnCount; column++) {
                        var cellX = cellOriginX + (cellWidth + columnGutter) * column;
                        reportChange();
                        var cellRect = cellLayer.pathItems.rectangle(cellY, cellX, cellWidth, cellHeight);
                        cellRect.stroked = false;
                        cellRect.filled = true;
                        cellRect.fillColor = cellFillColor;
                        cellRect.opacity = cellOpacity;
                        if (cornerRadius > 0) cellRect.applyEffect(roundCornersEffectXML(cornerRadius));
                        createdCells.push(cellRect);
                    }
                }
            }

            // 各セルを分割：セルの左右中央に縦ガイド / Split each cell: vertical guide at each cell's horizontal center
            if (splitCells && gridLayer) {
                var splitOriginX = contentLeft;
                var splitOriginY = contentTop;
                for (var splitRow = 0; splitRow < rowCount; splitRow++) {
                    var splitCellTop = splitOriginY - (cellHeight + rowGutter) * splitRow;
                    var splitCellBottom = splitCellTop - cellHeight;
                    for (var splitColumn = 0; splitColumn < columnCount; splitColumn++) {
                        var splitCenterX = splitOriginX + (cellWidth + columnGutter) * splitColumn + cellWidth / 2;
                        drawGuide([splitCenterX, splitCellTop], [splitCenterX, splitCellBottom]);
                    }
                }
            }
        }

        if (!isPreview && gridLayer) {
            gridLayer.locked = true;
        }
        if (isPreview) {
            app.redraw();
        }
        return {
            cells: createdCells,
            changed: !!(layerInfo.created || guideCount > 0 || createdCells.length > 0)
        };
    }

    // =========================================
    // メイン / Main
    // =========================================
    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてください。\nPlease open a document.");
            return;
        }

        var doc = app.activeDocument;

        // Preview manager (Undo-safe live preview)
        var previewManager = new PreviewManager();

        // 確定描画したセル長方形（中心点表示の対象）/ Cell rects from final draw (targets for center-point display)
        var drawnCellItems = [];

        // 選択オブジェクトの参照と外接矩形をキャッシュ / Cache selection refs and bounds before dialog
        var cachedSelectionItems = [];
        var cachedSelectionBounds = (function () {
            var selection = doc.selection;
            // 文字選択中（TextRange）は座標を持たないので対象外 / A TextRange has no bounds, so it is not a target
            if (!selection || selection.typename === "TextRange" || selection.length === 0) return null;
            for (var i = 0; i < selection.length; i++) {
                if (hasGeometricBounds(selection[i])) cachedSelectionItems.push(selection[i]);
            }
            if (cachedSelectionItems.length === 0) return null;
            var firstItemBounds = cachedSelectionItems[0].geometricBounds;
            var selectionLeft = firstItemBounds[0], selectionTop = firstItemBounds[1], selectionRight = firstItemBounds[2], selectionBottom = firstItemBounds[3];
            for (var j = 1; j < cachedSelectionItems.length; j++) {
                var itemBounds = cachedSelectionItems[j].geometricBounds;
                if (itemBounds[0] < selectionLeft) selectionLeft = itemBounds[0];
                if (itemBounds[1] > selectionTop) selectionTop = itemBounds[1];
                if (itemBounds[2] > selectionRight) selectionRight = itemBounds[2];
                if (itemBounds[3] < selectionBottom) selectionBottom = itemBounds[3];
            }
            return [selectionLeft, selectionTop, selectionRight, selectionBottom];
        })();
        // 初期対象モード / Initial target mode
        // 選択あり→選択オブジェクト、なし→（複数AB→すべて／単一→アートボード）
        var hasMultipleArtboards = doc.artboards.length > 1;
        var targetMode = cachedSelectionBounds
            ? "selection"
            : (hasMultipleArtboards ? "allArtboards" : "artboard");

        function isSelectionMode() {
            return targetMode === "selection";
        }
        function isAllArtboards() {
            return targetMode === "allArtboards";
        }

        // ダイアログ作成 / Create dialog
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        // プリセット選択＋書き出し / Preset selection and export
        var presetGroup = dialog.add("group");
        setupGroup(presetGroup, "row");
        presetGroup.alignment = ["center", "top"]; // 左右中央 / horizontally centered
        presetGroup.margins = [0, 0, 0, 10]; // 下に少し余白 / add some bottom margin

        presetGroup.add("statictext", undefined, labelText("field.preset"));
        var presetDropdown = presetGroup.add("dropdownlist", undefined, []);
        presetDropdown.selection = 0;
        var btnExportPreset = presetGroup.add("button", undefined, getLabel("button.exportPreset"));
        btnExportPreset.alignment = ["left", "center"]; // 横に伸ばさず天地中央 / natural width, vertically centered
        btnExportPreset.helpTip = getLabel("tooltip.exportPreset");

        // プリセット書き出し / Export the current settings as a preset
        btnExportPreset.onClick = function () {
            var saveFile = File.saveDialog("プリセットを書き出す場所と名前を指定してください / Choose where to save the preset", "*.txt");
            if (!saveFile) {
                return;
            }

            // 拡張子がない場合は.txtをつける / Add .txt extension if missing
            if (saveFile.name.indexOf(".") === -1) {
                saveFile = new File(saveFile.fsName + ".txt");
            }

            // ★ファイル名から.txtを正しく除去！ / Remove .txt extension from file name
            var fileName = saveFile.name.replace(/\.txt$/i, "");

            // 長さ系は pt に換算して保存（プリセットは pt 基準）/ Store length values in pt (presets are pt-based)
            var currentPreset = {
                columns: toInteger(columnCountInput.text, 1),
                rows: toInteger(rowCountInput.text, 1),
                guideExtension: unitToPt(toNumber(extensionInput.text, 0)),
                marginTop: unitToPt(toNumber(marginTopInput.text, 0)),
                marginBottom: unitToPt(toNumber(marginBottomInput.text, 0)),
                marginLeft: unitToPt(toNumber(marginLeftInput.text, 0)),
                marginRight: unitToPt(toNumber(marginRightInput.text, 0)),
                rowGutter: unitToPt(toNumber(rowGutterInput.text, 0)),
                columnGutter: unitToPt(toNumber(columnGutterInput.text, 0)),
                drawCells: cellRectCheckbox.value,
                drawGuides: drawGuidesCheckbox.value
            };

            var presetString = '{ label: "' + fileName + '", ' +
                'columns: ' + currentPreset.columns + ', ' +
                'rows: ' + currentPreset.rows + ', ' +
                'guideExtension: ' + currentPreset.guideExtension + ', ' +
                'marginTop: ' + currentPreset.marginTop + ', ' +
                'marginBottom: ' + currentPreset.marginBottom + ', ' +
                'marginLeft: ' + currentPreset.marginLeft + ', ' +
                'marginRight: ' + currentPreset.marginRight + ', ' +
                'rowGutter: ' + currentPreset.rowGutter + ', ' +
                'columnGutter: ' + currentPreset.columnGutter + ', ' +
                'drawCells: ' + currentPreset.drawCells + ', ' +
                'drawGuides: ' + currentPreset.drawGuides +
                ' }';

            if (saveFile.open("w")) {
                saveFile.write(presetString);
                saveFile.close();
                alert("プリセットを書き出しました！ / Preset exported!");
            } else {
                alert("ファイルを書き込めませんでした。 / Failed to write the file.");
            }
        };

        // =========================================
        // 対象パネル / Target panel
        // =========================================
        // 対象パネルと元のオブジェクトパネルを横並び（別パネル）/ Target panel and Original-object panel side by side (separate panels)
        var targetRow = dialog.add("group");
        setupGroup(targetRow, "row", COLUMN_SPACING);
        targetRow.alignChildren = ["left", "fill"];

        // 対象パネル：対象の種類（縦並び）/ Target panel: target type (vertical)
        var targetPanel = targetRow.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel);
        var targetTypeGroup = targetPanel.add("group");
        setupGroup(targetTypeGroup, "column");
        var targetSelectionRadio = targetTypeGroup.add("radiobutton", undefined, getLabel("target.selection"));
        var targetArtboardRadio = targetTypeGroup.add("radiobutton", undefined, getLabel("target.artboard"));
        var targetAllArtboardsRadio = targetTypeGroup.add("radiobutton", undefined, getLabel("target.allArtboards"));

        // 元のオブジェクトパネル（対象パネルの外・横並び、選択オブジェクト時のみ有効）/ Original-object panel (outside target panel, selection mode only)
        var originalObjectPanel = targetRow.add("panel", undefined, getLabel("panel.originalObject"));
        setupPanel(originalObjectPanel);
        var originalObjectGroup = originalObjectPanel.add("group");
        setupGroup(originalObjectGroup, "column");
        var removeOriginalRadio = originalObjectGroup.add("radiobutton", undefined, getLabel("radio.remove"));
        var keepOriginalRadio = originalObjectGroup.add("radiobutton", undefined, getLabel("radio.keep"));
        var makeOriginalGuidesRadio = originalObjectGroup.add("radiobutton", undefined, getLabel("radio.toGuide"));
        makeOriginalGuidesRadio.helpTip = getLabel("tooltip.toGuide");
        keepOriginalRadio.value = true; // デフォルト / default

        // 選択がなければ「選択オブジェクト」を無効化 / Disable selection option when nothing is selected
        if (!cachedSelectionBounds) {
            targetSelectionRadio.enabled = false;
        }
        // アートボードが1つだけなら「すべてのアートボード」を無効化 / Disable "All Artboards" when only one artboard
        if (!hasMultipleArtboards) {
            targetAllArtboardsRadio.enabled = false;
        }
        // 初期選択 / Initial target radio
        targetSelectionRadio.value = (targetMode === "selection");
        targetArtboardRadio.value = (targetMode === "artboard");
        targetAllArtboardsRadio.value = (targetMode === "allArtboards");

        // 対象モード変更 / Target mode change
        function onTargetChange() {
            if (targetSelectionRadio.value) {
                targetMode = "selection";
            } else if (targetAllArtboardsRadio.value) {
                targetMode = "allArtboards";
            } else {
                targetMode = "artboard";
            }
            updateTargetMode();
            safeUpdatePreview();
        }
        targetSelectionRadio.onClick = targetArtboardRadio.onClick = targetAllArtboardsRadio.onClick = onTargetChange;

        // 元オブジェクト処理の切替でプレビュー更新 / Update preview on original-object radio change
        removeOriginalRadio.onClick = keepOriginalRadio.onClick = makeOriginalGuidesRadio.onClick = function () {
            safeUpdatePreview();
        };

        // 対象モードに応じた表示制御 / Enable/disable controls by target mode
        function updateTargetMode() {
            var isSelectionTarget = isSelectionMode();
            originalObjectPanel.enabled = isSelectionTarget;
            // 選択オブジェクトにはアートボードの辺がないのでディム / No artboard edges when the target is the selection
            artboardEdgeCheckbox.enabled = !isSelectionTarget;
        }

        // グリッド設定グループ / Grid settings group
        var gridSettingRow = dialog.add("group");
        setupGroup(gridSettingRow, "row", COLUMN_SPACING);
        gridSettingRow.alignChildren = ["left", "top"];
        var gridLabelWidth = (uiLang === "ja") ? 40 : 50; // unify Number/Gutter label width and right-align

        // 行設定パネル / Row settings panel
        var rowSettingPanel = gridSettingRow.add("panel", undefined, getLabel("panel.row"));
        setupPanel(rowSettingPanel);

        var rowCountGroup = rowSettingPanel.add("group");
        setupGroup(rowCountGroup, "row");
        var rowCountLabel = rowCountGroup.add("statictext", undefined, labelText("field.rowCount"));
        rowCountLabel.preferredSize.width = gridLabelWidth;
        rowCountLabel.justify = "right";
        var rowCountInput = addStepperInput(rowCountGroup, "2", 3, { min: 1, integer: true });

        var rowGutterGroup = rowSettingPanel.add("group");
        setupGroup(rowGutterGroup, "row");
        var rowGutterLabel = rowGutterGroup.add("statictext", undefined, labelText("field.rowGutter"));
        rowGutterLabel.preferredSize.width = gridLabelWidth;
        rowGutterLabel.justify = "right";
        /* 単位換算後は小数2桁になるため桁数を確保 / Room for the 2 decimals unit conversion produces */
        var rowGutterInput = addStepperInput(rowGutterGroup, "0", 5, { min: 0 });
        rowGutterGroup.add("statictext", undefined, unitLabel);

        // 列設定パネル / Column settings panel
        var columnSettingPanel = gridSettingRow.add("panel", undefined, getLabel("panel.column"));
        setupPanel(columnSettingPanel);

        var columnCountGroup = columnSettingPanel.add("group");
        setupGroup(columnCountGroup, "row");
        var columnCountLabel = columnCountGroup.add("statictext", undefined, labelText("field.columnCount"));
        columnCountLabel.preferredSize.width = gridLabelWidth;
        columnCountLabel.justify = "right";
        var columnCountInput = addStepperInput(columnCountGroup, "2", 3, { min: 1, integer: true });

        var columnGutterGroup = columnSettingPanel.add("group");
        setupGroup(columnGutterGroup, "row");
        var columnGutterLabel = columnGutterGroup.add("statictext", undefined, labelText("field.columnGutter"));
        columnGutterLabel.preferredSize.width = gridLabelWidth;
        columnGutterLabel.justify = "right";
        var columnGutterInput = addStepperInput(columnGutterGroup, "0", 5, { min: 0 });
        columnGutterGroup.add("statictext", undefined, unitLabel);

        // 行間に連動（列パネル下部）/ Link to row gutter (under column panel)
        var linkGutterCheckbox = columnSettingPanel.add("checkbox", undefined, getLabel("checkbox.linkGutter"));
        linkGutterCheckbox.helpTip = getLabel("tooltip.linkGutter");
        linkGutterCheckbox.value = true;

        // 連動の切り替えはガター全体の有効/無効と同じ判定なので updateGutterEnabled に任せる
        // Toggling the link runs the same rules as the gutter enable state, so reuse updateGutterEnabled
        linkGutterCheckbox.onClick = function () {
            updateGutterEnabled();
            safeUpdatePreview();
        };

        // マージン全体パネル / Margin panel
        var marginPanel = dialog.add("panel", undefined, titleWithUnit("panel.margin"));
        setupPanel(marginPanel);
        // 3×3 グリッド配置（中央=連動）/ 3×3 grid layout (center = link)
        var MARGIN_CELL_WIDTH = (uiLang === "ja") ? 78 : 92;

        // ラベル＋数値のセル / A label+field cell
        function addMarginCell(parentRow, labelKey) {
            var cellGroup = parentRow.add("group");
            cellGroup.orientation = "row";
            cellGroup.alignment = ["center", "center"];
            cellGroup.minimumSize.width = MARGIN_CELL_WIDTH;
            cellGroup.add("statictext", undefined, labelText(labelKey));
            var marginInput = addStepperInput(cellGroup, "0", 5, { min: 0 });
            return { group: cellGroup, input: marginInput };
        }
        // 位置合わせ用の空セル / Empty cell for alignment
        function addMarginSpacer(parentRow) {
            var spacerGroup = parentRow.add("group");
            spacerGroup.minimumSize.width = MARGIN_CELL_WIDTH;
        }

        // 1行目：［空］［上］［空］/ Row 1: [empty][top][empty]
        var marginTopRow = marginPanel.add("group");
        marginTopRow.orientation = "row";
        marginTopRow.alignment = ["center", "top"];
        addMarginSpacer(marginTopRow);
        var marginTopCell = addMarginCell(marginTopRow, "field.top");
        var marginTopInput = marginTopCell.input;
        addMarginSpacer(marginTopRow);

        // 2行目：［左］［連動］［右］/ Row 2: [left][link][right]
        var marginMiddleRow = marginPanel.add("group");
        marginMiddleRow.orientation = "row";
        marginMiddleRow.alignment = ["center", "top"];
        var marginLeftCell = addMarginCell(marginMiddleRow, "field.left");
        var marginLeftGroup = marginLeftCell.group;
        var marginLeftInput = marginLeftCell.input;
        /* 中央のセルにリンクアイコンを置く（幅は上下の空セルとそろえる）/ Link icon in the centre cell (same width as the empty cells above and below) */
        var linkMarginGroup = marginMiddleRow.add("group");
        linkMarginGroup.orientation = "row";
        linkMarginGroup.alignment = ["center", "center"];
        linkMarginGroup.alignChildren = ["center", "center"];
        linkMarginGroup.minimumSize.width = MARGIN_CELL_WIDTH;
        // デフォルトでON / on by default
        var linkMarginToggle = addLinkToggle(linkMarginGroup, true, function () {
            syncLinkedMargins();
            safeUpdatePreview();
        });
        linkMarginToggle.helpTip = getLabel("tooltip.linkMargin");
        var marginRightCell = addMarginCell(marginMiddleRow, "field.right");
        var marginRightGroup = marginRightCell.group;
        var marginRightInput = marginRightCell.input;

        // 3行目：［空］［下］［空］/ Row 3: [empty][bottom][empty]
        var marginBottomRow = marginPanel.add("group");
        marginBottomRow.orientation = "row";
        marginBottomRow.alignment = ["center", "top"];
        addMarginSpacer(marginBottomRow);
        var marginBottomCell = addMarginCell(marginBottomRow, "field.bottom");
        var marginBottomGroup = marginBottomCell.group;
        var marginBottomInput = marginBottomCell.input;
        addMarginSpacer(marginBottomRow);

        // セルパネル＋ガイドパネルを横並び（左：セル、右：ガイド）/ Cell panel + Guides panel side by side (left: cell, right: guides)
        var cellGuideRow = dialog.add("group");
        setupGroup(cellGuideRow, "row", COLUMN_SPACING);
        cellGuideRow.alignChildren = ["left", "fill"];

        // セルパネル（ガイドパネルの左）/ Cell panel (left of Guides panel)
        var cellPanel = cellGuideRow.add("panel", undefined, getLabel("panel.options"));
        setupPanel(cellPanel);
        var cellRectCheckbox = cellPanel.add("checkbox", undefined, getLabel("checkbox.cellRect"));
        cellRectCheckbox.helpTip = getLabel("tooltip.cellRect");
        var centerPointCheckbox = cellPanel.add("checkbox", undefined, getLabel("checkbox.showCenter"));
        centerPointCheckbox.helpTip = getLabel("tooltip.showCenter");
        centerPointCheckbox.value = true;

        // 角丸（ライブエフェクト）/ Round corners (live effect)
        var roundCornerGroup = cellPanel.add("group");
        setupGroup(roundCornerGroup, "row");
        var roundCornerCheckbox = roundCornerGroup.add("checkbox", undefined, getLabel("checkbox.roundCorner"));
        roundCornerCheckbox.helpTip = getLabel("tooltip.roundCorner");
        roundCornerCheckbox.value = false;
        var roundCornerInput = addStepperInput(roundCornerGroup, "3", 2, { min: 0 });
        roundCornerGroup.add("statictext", undefined, unitLabel);

        // 不透明度スライダー（0-100、ラベルの次の行にスライダー）/ Opacity slider (0-100, slider on the line below the label)
        var opacityGroup = cellPanel.add("group");
        setupGroup(opacityGroup, "column", ROW_SPACING);
        opacityGroup.add("statictext", undefined, labelText("field.opacity"));
        var opacitySlider = opacityGroup.add("slider", undefined, CELL_OPACITY, 0, 100);
        opacitySlider.helpTip = getLabel("tooltip.opacity");
        opacitySlider.alignment = ["fill", "center"];
        opacitySlider.onChanging = function () {
            safeUpdatePreview();
        };

        // ガイドパネル / Guides panel
        var guidesPanel = cellGuideRow.add("panel", undefined, getLabel("panel.guides"));
        setupPanel(guidesPanel);

        // ガイドを引く / Draw guides
        var drawGuidesCheckbox = guidesPanel.add("checkbox", undefined, getLabel("checkbox.drawGuides"));
        drawGuidesCheckbox.value = true;

        // ガイドの伸張（チェックボックスで有効/無効）/ Guide extension (checkbox toggles on/off)
        var extensionGroup = guidesPanel.add("group");
        setupGroup(extensionGroup, "row");
        var extensionCheckbox = extensionGroup.add("checkbox", undefined, getLabel("field.guideExtension"));
        extensionCheckbox.helpTip = getLabel("tooltip.guideExtension");
        extensionCheckbox.value = true;
        var extensionInput = addStepperInput(extensionGroup, "10", 5, { min: 0 });
        extensionGroup.add("statictext", undefined, unitLabel);
        extensionCheckbox.onClick = function () {
            setStepperInputEnabled(extensionInput, extensionCheckbox.value);
            safeUpdatePreview();
        };

        // アートボードのエッジ（上下左右4辺）/ Artboard edges (all four sides)
        var artboardEdgeCheckbox = guidesPanel.add("checkbox", undefined, getLabel("checkbox.artboardEdge"));
        artboardEdgeCheckbox.helpTip = getLabel("tooltip.artboardEdge");
        artboardEdgeCheckbox.value = false;
        artboardEdgeCheckbox.onClick = function () {
            safeUpdatePreview();
        };

        // 各セルを分割（各セルの左右中央に縦ガイド）/ Split each cell (vertical guide at each cell's horizontal center)
        var splitCellCheckbox = guidesPanel.add("checkbox", undefined, getLabel("checkbox.splitCell"));
        splitCellCheckbox.helpTip = getLabel("tooltip.splitCell");
        splitCellCheckbox.value = false;
        splitCellCheckbox.onClick = function () {
            safeUpdatePreview();
        };

        // レイヤークリア / Clear layer
        var clearGuidesCheckbox = guidesPanel.add("checkbox", undefined, getLabel("checkbox.clearGuides"));
        clearGuidesCheckbox.helpTip = getLabel("tooltip.clearGuides");
        clearGuidesCheckbox.value = false;
        clearGuidesCheckbox.onClick = function () {
            safeUpdatePreview();
        };

        // 実際に削除するか（ディム中は値が残っていても削除しない）/ Whether to clear (a dimmed checkbox never clears)
        function shouldClearGuides() {
            return clearGuidesCheckbox.enabled && clearGuidesCheckbox.value;
        }

        // 入力値変更で即時プレビュー / Live preview on any input change
        function attachLivePreview(editText) {
            editText.onChanging = function () {
                safeUpdatePreview();
            };
        }

        // プレビュー更新（Undoで巻き戻してから1回だけ描画）/ Update preview (rollback then draw once)
        function updatePreview() {
            try {
                // If there was a previous preview step, rollback first
                previewManager.rollback();
            } catch (e) {
                $.writeln("[GenerateGuidesGrid] preview rollback error: " + e);
            }

            // 「既存ガイドを削除」ONなら、この時点で削除して見た目に反映する
            // キャンセル時は rollback の app.undo() で元に戻る / Cancel restores them via rollback
            if (shouldClearGuides()) {
                previewManager.addStep(function (markChanged) {
                    return clearExistingGuides(markChanged); // 何も消さなければ false / false when nothing was removed
                });
            }

            // Draw preview as one undoable step
            previewManager.addStep(function (markChanged) {
                var drawContext = buildDrawContext(true);
                if (!canDrawGrid(drawContext)) return false; // 何も描かない＝undo対象なし / nothing drawn, nothing to undo
                drawContext.onChange = markChanged; // 描き始めたことを伝える / report the moment drawing starts
                return drawGrid(drawContext).changed; // 何も作られなければ false / false when nothing was created
            });

            // 選択オブジェクトの表示/非表示をプレビュー / Preview hide/show of selected objects
            if (isSelectionMode() && cachedSelectionItems.length > 0) {
                var shouldHide = (removeOriginalRadio && removeOriginalRadio.value);
                if (shouldHide) {
                    previewManager.addStep(function (markChanged) {
                        markChanged();
                        for (var i = 0; i < cachedSelectionItems.length; i++) {
                            cachedSelectionItems[i].hidden = true;
                        }
                        return true;
                    });
                }
            }
        }

        // updatePreview を安全に呼ぶ（イベントハンドラが落ちないように。エラーはログのみ）/ Call updatePreview safely (keep handlers alive; log the error)
        function safeUpdatePreview() {
            try {
                updatePreview();
            } catch (e) {
                $.writeln("[GenerateGuidesGrid] updatePreview error: " + e);
            }
        }

        // 入力中の変更もリアルタイム反映 / Attach onChanging for live preview
        // 行数・列数はガターの有効/無効も更新するため、別途 onChanging を割り当てる / Row/column counts also refresh the gutter enable state, so they get their own onChanging
        attachLivePreview(extensionInput);
        // 上マージン変更時に連動ONなら左右下も同期 / Sync margins when top changes (if linked)
        marginTopInput.onChanging = function () {
            if (linkMarginToggle.value) {
                marginBottomInput.text = marginTopInput.text;
                marginLeftInput.text = marginTopInput.text;
                marginRightInput.text = marginTopInput.text;
            }
            safeUpdatePreview();
        };
        attachLivePreview(marginBottomInput);
        attachLivePreview(marginLeftInput);
        attachLivePreview(marginRightInput);
        // 行間変更時に連動チェックONなら列間も同期 / Sync col gutter when row gutter changes (if linked)
        rowGutterInput.onChanging = function () {
            if (linkGutterCheckbox.value) {
                columnGutterInput.text = rowGutterInput.text;
            }
            safeUpdatePreview();
        };
        attachLivePreview(columnGutterInput);

        // === ボタンエリア（3カラム：左アウトライン／中央スペーサー／右キャンセル・OK）/ Button area (3 columns: left outline / center spacer / right cancel+ok)
        var buttonRow = addButtonRow(dialog);

        // 左グループ（アウトラインボタン）/ Left group (Outline button)
        var btnOutline = buttonRow.leftGroup.add("button", undefined, getLabel("button.outline"));
        btnOutline.helpTip = getLabel("tooltip.outline");
        // アウトライン⇔プレビュー表示を切り替え、ラベルもトグル / Toggle Outline/Preview view and the button label
        btnOutline.onClick = function () {
            app.executeMenuCommand('preview');
            btnOutline.text = (btnOutline.text === getLabel("button.outline"))
                ? getLabel("button.preview")
                : getLabel("button.outline");
        };

        // 右グループ（キャンセル・OKボタン）/ Right group (Cancel/OK buttons)
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), {
            name: "cancel"
        });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), {
            name: "ok"
        });

        // 表示用ラベルをローカライズ / Localize display label for dropdown
        function presetDisplayLabel(rawLabel) {
            // 日本語UIのときは「 / 」以降を隠す / In Japanese UI, hide text after " / "
            if (uiLang === "ja") return String(rawLabel).replace(/\s*\/.*$/, "");
            return rawLabel;
        }

        // プリセットをドロップダウンに追加 / Add presets to dropdown
        for (var i = 0; i < presets.length; i++) {
            var presetLabel = presetDisplayLabel(presets[i].label);
            presetDropdown.add("item", presetLabel);
        }
        presetDropdown.selection = 0;

        // プリセットの値を入力欄に反映する共通関数 / Common function to apply preset values
        function applyPreset(preset) {
            /* 値を取得（未指定なら既定値）/ Pick a value: the preset's, or the default */
            function pickPresetValue(value, fallbackValue) {
                return (value !== undefined) ? value : fallbackValue;
            }
            // 行数・列数（換算なし）/ Counts (no unit conversion)
            columnCountInput.text = pickPresetValue(preset.columns, 1);
            rowCountInput.text = pickPresetValue(preset.rows, 1);
            // 長さ系は pt 基準なので現在単位へ換算 / Length values are stored in pt — convert to the current unit
            extensionInput.text = ptToUnit(pickPresetValue(preset.guideExtension, 0));
            var presetMarginTop = pickPresetValue(preset.marginTop, 0);
            var presetMarginBottom = pickPresetValue(preset.marginBottom, 0);
            var presetMarginLeft = pickPresetValue(preset.marginLeft, 0);
            var presetMarginRight = pickPresetValue(preset.marginRight, 0);
            marginTopInput.text = ptToUnit(presetMarginTop);
            marginBottomInput.text = ptToUnit(presetMarginBottom);
            marginLeftInput.text = ptToUnit(presetMarginLeft);
            marginRightInput.text = ptToUnit(presetMarginRight);
            rowGutterInput.text = ptToUnit(pickPresetValue(preset.rowGutter, 0));
            columnGutterInput.text = ptToUnit(pickPresetValue(preset.columnGutter, 0));
            // 上下左右が異なるプリセットは連動をOFF（連動が値を上書きして壊すのを防ぐ）
            // If margins differ, turn the link off so it won't overwrite the distinct values
            setLinkToggleValue(linkMarginToggle, (presetMarginTop === presetMarginBottom && presetMarginTop === presetMarginLeft && presetMarginTop === presetMarginRight));
            cellRectCheckbox.value = (typeof preset.drawCells !== "undefined") ? preset.drawCells : false;
            drawGuidesCheckbox.value = (typeof preset.drawGuides !== "undefined") ? preset.drawGuides : true;
            extensionGroup.enabled = drawGuidesCheckbox.value;
            setStepperInputEnabled(extensionInput, drawGuidesCheckbox.value && extensionCheckbox.value);
            clearGuidesCheckbox.enabled = drawGuidesCheckbox.value;
        }

        // プリセット選択時に入力値へ反映 / Apply preset values to inputs on selection
        presetDropdown.onChange = function () {
            applyPreset(presets[presetDropdown.selection.index]);
            updateGutterEnabled();
            updateTargetMode();
            updateCellOptionEnabled();
            syncLinkedMargins();
            safeUpdatePreview();
        };

        // 初期プリセットの値を入力欄に反映 / Apply initial preset values to input fields
        applyPreset(presets[0]);

        // 「連動」同期処理 / Sync for "Link" margin
        function syncLinkedMargins() {
            if (linkMarginToggle.value) {
                var topValue = marginTopInput.text;
                marginBottomInput.text = topValue;
                marginLeftInput.text = topValue;
                marginRightInput.text = topValue;
                marginBottomGroup.enabled = false;
                marginLeftGroup.enabled = false;
                marginRightGroup.enabled = false;
            } else {
                marginBottomGroup.enabled = true;
                marginLeftGroup.enabled = true;
                marginRightGroup.enabled = true;
            }
            /* ∧∨のディム表示を切り替える / update the stepper dimming */
            redrawSteppersIn(marginBottomGroup);
            redrawSteppersIn(marginLeftGroup);
            redrawSteppersIn(marginRightGroup);
        }

        // ガター有効無効切り替え / Enable/disable gutter fields
        function updateGutterEnabled() {
            var columnCount = parseInt(columnCountInput.text, 10);
            var rowCount = parseInt(rowCountInput.text, 10);
            rowGutterGroup.enabled = (rowCount > 1);
            // 行数・列数が両方2以上のときのみ連動を有効 / Enable link only when both >= 2
            var hasMultipleRowsAndColumns = (columnCount > 1 && rowCount > 1);
            linkGutterCheckbox.enabled = hasMultipleRowsAndColumns;
            if (linkGutterCheckbox.value && hasMultipleRowsAndColumns) {
                columnGutterInput.text = rowGutterInput.text;
                columnGutterGroup.enabled = false;
            } else {
                columnGutterGroup.enabled = (columnCount > 1);
            }
            /* ∧∨のディム表示を切り替える / update the stepper dimming */
            redrawSteppersIn(rowGutterGroup);
            redrawSteppersIn(columnGutterGroup);
        }

        // 行数・列数変更時：ガター有効/無効を更新して即時プレビュー / On row/col change: refresh gutter enable + preview
        columnCountInput.onChanging = rowCountInput.onChanging = function () {
            updateGutterEnabled();
            safeUpdatePreview();
        };

        // 「ガイドを引く」切替：伸張・クリアをディム制御してプレビュー / Toggle draw-guides: dim extension/clear, then preview
        drawGuidesCheckbox.onClick = function () {
            var guidesEnabled = drawGuidesCheckbox.value;
            extensionGroup.enabled = guidesEnabled;
            setStepperInputEnabled(extensionInput, guidesEnabled && extensionCheckbox.value);
            clearGuidesCheckbox.enabled = guidesEnabled; // ガイドOFFならクリアをディム / dim clear when guides off
            safeUpdatePreview();
        };

        // 長方形化の有無でセル系オプション（中心点・角丸・不透明度）をディム制御 / Cell options dim unless rectangles are drawn
        // 角丸の数値欄は「長方形化ON かつ 角丸ON」のときだけ有効 / The round-corner field is enabled only when both rectangles and round corners are on
        function updateCellOptionEnabled() {
            var cellsEnabled = cellRectCheckbox.value;
            centerPointCheckbox.enabled = cellsEnabled;
            roundCornerCheckbox.enabled = cellsEnabled;
            setStepperInputEnabled(roundCornerInput, cellsEnabled && roundCornerCheckbox.value);
            opacitySlider.enabled = cellsEnabled;
        }

        cellRectCheckbox.onClick = function () {
            updateCellOptionEnabled();
            safeUpdatePreview();
        };

        // 角丸トグルで数値欄の有効/無効を更新してプレビュー / Toggle round corners: refresh the field enable, then preview
        roundCornerCheckbox.onClick = function () {
            updateCellOptionEnabled();
            safeUpdatePreview();
        };
        attachLivePreview(roundCornerInput);

        // OKボタン押下時（ドキュメントは変更せず閉じるだけ。クリア/描画は確定処理 finalAction で実行）
        // OK button: just close (no document mutation here — clear/draw happens in finalAction to keep undo bookkeeping correct)
        btnOK.onClick = function () {
            updateGutterEnabled();
            dialog.close(1);
        };

        // キャンセル：ダイアログを閉じるだけ（後処理は dialog.show() 後の分岐で rollback）/ Cancel: just close; cleanup/rollback happens after dialog.show()
        btnCancel.onClick = function () {
            dialog.close(0);
        };

        // 行・列・ガター・伸張・ガイド系の設定を収集 / Collect grid (rows/cols/gutters/extension/guide) settings
        function collectGridSettings() {
            return {
                /* 空欄・不正値は書き出し側と同じく 1 として扱う（0 だと削除だけ走って何も描かれない）
                   Treat blank/invalid as 1 like the exporter (0 would clear without drawing) */
                columnCount: toInteger(columnCountInput.text, 1),
                rowCount: toInteger(rowCountInput.text, 1),
                /* ディム中（「ガイドを引く」OFF）の伸張は効かせない / A dimmed extension (draw-guides off) must not affect the output */
                guideExtension: ((extensionGroup.enabled && extensionCheckbox.value) ? toNumber(extensionInput.text, 0) : 0) * unitFactor,
                rowGutter: toNumber(rowGutterInput.text, 0) * unitFactor,
                columnGutter: toNumber(columnGutterInput.text, 0) * unitFactor,
                drawGridGuides: drawGuidesCheckbox.value,
                drawArtboardEdge: artboardEdgeCheckbox.value,
                splitCells: splitCellCheckbox.value
            };
        }

        // 上下左右マージンを収集 / Collect top/bottom/left/right margins
        function collectMarginSettings() {
            return {
                marginTop: toNumber(marginTopInput.text, 0) * unitFactor,
                marginBottom: toNumber(marginBottomInput.text, 0) * unitFactor,
                marginLeft: toNumber(marginLeftInput.text, 0) * unitFactor,
                marginRight: toNumber(marginRightInput.text, 0) * unitFactor
            };
        }

        // セル（長方形化・角丸・不透明度）の設定を収集 / Collect cell (rect/round/opacity) settings
        function collectCellSettings() {
            return {
                drawCells: cellRectCheckbox.value,
                cornerRadius: roundCornerCheckbox.value ? toNumber(roundCornerInput.text, 0) * unitFactor : 0,
                cellOpacity: Math.round(opacitySlider.value)
            };
        }

        // 各 collector をまとめて描画コンテキストを構築 / Combine the collectors into a draw context
        function buildDrawContext(isPreview) {
            var drawContext = {
                doc: doc,
                isPreview: isPreview,
                allBoards: isAllArtboards(),
                selBounds: isSelectionMode() ? cachedSelectionBounds : null
            };
            mergeInto(drawContext, collectGridSettings());
            mergeInto(drawContext, collectMarginSettings());
            mergeInto(drawContext, collectCellSettings());
            return drawContext;
        }

        // 属性パネルの「中心を表示」を選択オブジェクトに適用 / Apply Attributes-panel "Show Center" to the current selection
        // API で直接設定できないため、一時アクション（.aia）を読み込んで再生 / No direct API, so load and play a temporary action
        function applyShowCenterAction() {
            var actionSource = '/version 3' + '/name [ 9' + ' 417474726962757465' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 10' + ' 53686f7743656e746572' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_attributePalette)' + ' /localizedName [ 12' + ' e5b19ee680a7e8a8ade5ae9a' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1668183154' + ' /showInPalette 4294967295' + ' /type (boolean)' + ' /value 1' + ' }' + ' }' + '}';

            // 読み込み・実行・解除は共通の処理で（失敗してもセットと一時ファイルを残さない）/ Shared load/play/unload; never leaves the set or temp file behind
            if (!runTemporaryAction(actionSource, "Attribute", "ShowCenter")) { // set name, action name
                $.writeln("[GenerateGuidesGrid] show center action failed");
            }
        }

        /**
         * grid_guidesレイヤーのガイドを削除する。
         * 対象が「すべてのアートボード」なら全部、それ以外はアクティブなアートボード上のガイドのみ。
         * Remove guides from the grid_guides layer: all of them for "all artboards", otherwise only those on the active artboard.
         * @param {function} [markChanged] - 最初にドキュメントを変更する直前に呼ぶコールバック
         * @returns {boolean} 1つ以上削除したら true
         */
        function clearExistingGuides(markChanged) {
            var guidesLayer = null;
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === GUIDE_LAYER_NAME) {
                    guidesLayer = doc.layers[i];
                    break;
                }
            }
            if (!guidesLayer) return false;

            // すべてのアートボードが対象なら範囲を絞らない / No limit when every artboard is a target
            var limitRect = isAllArtboards()
                ? null
                : doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;

            var removedCount = 0;
            for (var j = guidesLayer.pageItems.length - 1; j >= 0; j--) {
                var item = guidesLayer.pageItems[j];
                if (!item.guides) continue;
                if (limitRect && !isCenterInsideRect(item, limitRect)) continue;
                // 削除対象が見つかってからロックを解除する（何も消さないのに解除したままにしない）
                // Unlock only once there is something to remove (never leave it unlocked for nothing)
                if (removedCount === 0) {
                    if (markChanged) markChanged();
                    safeUnlockLayer(guidesLayer);
                }
                item.remove();
                removedCount++;
            }
            return removedCount > 0;
        }

        // 元の選択を復元（プレビューのundo繰り返しで選択が外れるため）/ Restore original selection (preview undo cycles clear it)
        function safeRestoreSelection() {
            if (cachedSelectionItems.length === 0) return;
            try {
                doc.selection = null;
                for (var rs = 0; rs < cachedSelectionItems.length; rs++) {
                    cachedSelectionItems[rs].selected = true;
                }
            } catch (e) {
                $.writeln("[GenerateGuidesGrid] restore selection error: " + e);
            }
        }

        // ダイアログ初期プレビュー＆終了時処理 / Initial dialog preview & post-process
        updateGutterEnabled();
        updateCellOptionEnabled();
        syncLinkedMargins();
        updateTargetMode();
        safeUpdatePreview();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        if (dialog.show() === 1) {
            // OK: rollback preview and execute final drawing once so user can undo in one step
            previewManager.confirm(function () {
                // プレビューレイヤーを先に削除（最後に消すと選択が解除されるため）/ Remove preview layer first (removing it later clears the selection)
                safeRemoveLayerByName(doc, PREVIEW_LAYER_NAME);
                var finalContext = buildDrawContext(false);
                // 描画できない設定のときは削除もしない（消すだけで終わる事故を防ぐ）
                // Do not clear when nothing can be drawn (avoids a delete-only outcome)
                if (shouldClearGuides() && canDrawGrid(finalContext)) {
                    clearExistingGuides();
                }
                drawnCellItems = drawGrid(finalContext).cells;
                // 選択オブジェクトの処理 / Handle original selected objects
                if (cachedSelectionItems.length > 0 && isSelectionMode()) {
                    if (removeOriginalRadio && removeOriginalRadio.value) {
                        for (var i = cachedSelectionItems.length - 1; i >= 0; i--) {
                            var itemToRemove = cachedSelectionItems[i];
                            // ロック・非表示などで失敗しても他を続行 / Keep going even if one remove fails (locked/hidden, etc.)
                            safeExecute(function () { itemToRemove.remove(); });
                        }
                    } else if (makeOriginalGuidesRadio && makeOriginalGuidesRadio.value) {
                        // 元オブジェクトをガイド化 / Convert originals to guides
                        try {
                            doc.selection = null;
                            for (var j = 0; j < cachedSelectionItems.length; j++) {
                                cachedSelectionItems[j].selected = true;
                            }
                            app.executeMenuCommand("Make Guides");
                        } catch (e) {
                            $.writeln("[GenerateGuidesGrid] make originals guides error: " + e);
                        }
                    }
                    // keepOriginalRadio: 何もしない / do nothing
                }
                // 最終的に選択は解除する（元の図形もセルも選択しない）/ Clear selection at the end (neither originals nor cells)
                // 中心点表示が必要なときだけ、一時的にセルを選択してアクション適用後に解除
                // Only when Show Center is on, select cells transiently to apply the action, then clear
                try {
                    doc.selection = null;
                    if (centerPointCheckbox.value && drawnCellItems.length > 0) {
                        for (var k = 0; k < drawnCellItems.length; k++) {
                            drawnCellItems[k].selected = true;
                        }
                        applyShowCenterAction();
                        doc.selection = null;
                    }
                } catch (e) {
                    $.writeln("[GenerateGuidesGrid] show center error: " + e);
                }
            });
        } else {
            // Cancel: rollback preview changes
            previewManager.rollback();
            // Cleanup preview layer just in case (fallback)
            safeRemoveLayerByName(doc, PREVIEW_LAYER_NAME);
            // プレビューで外れた選択を元に戻す / Restore selection lost during preview
            safeRestoreSelection();
        }
    }

    main();

})();
