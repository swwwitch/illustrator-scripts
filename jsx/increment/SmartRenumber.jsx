#target illustrator
#targetengine "SmartRenumberEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した数字・英字・漢数字のテキストを、位置・重ね順・現在の値などの並び順でソートして連番を振り直します。
書式と開始値はプレビューを見ながら指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenumber.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf3b6601cd165

### Overview

Sorts the selected text — digits, letters, or Japanese numerals — by position, stacking, or current value and renumbers it as a sequence.
The format and start value can be set while watching the preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartRenumber.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartRenumber";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenumber.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartRenumber.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf3b6601cd165"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 同じ行・同じ列と見なす座標の許容差（pt）。Z方向・N方向の並べ替えで使う
       Tolerance in points for treating items as one row or column (Z and N patterns) */
    var ROW_COLUMN_TOLERANCE_PT = 10;

    /* ［重ね順調整］をONで開くか / Whether Adjust Stacking starts checked */
    var REORDER_STACK_BY_DEFAULT = true;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var START_VALUE_FIELD_CHARS = 5;   /* ［開始値］の入力幅（文字数） / width of the start value field */
    var AFFIX_FIELD_CHARS       = 8;   /* 接頭辞・接尾辞の入力幅（文字数） / width of the affix fields */
    var SECTION_TOP_MARGIN      = 10;  /* パネル内で段を分けるときの上マージン / top margin that separates a section inside a panel */
    var LABEL_FIELD_SPACING     = 4;   /* 上に置いた項目名と入力欄の間隔 / gap between a label and the field below it */

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
     * 横並びの行グループの共通設定（ボタン列など）
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - alignment（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = rowAlignment || "left";
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
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

    // =========================================
    // 並び順 / Sort orders
    // =========================================

    /* 並び順の並べる順番。ラジオボタンの並び、LABELS.radio / LABELS.tooltip のキー、
       SORT_COMPARATORS のキーをこの名前で揃える
       Sort order sequence; the same keys are used for the radios, LABELS, and SORT_COMPARATORS */
    var SORT_MODES = ["currentValue", "horizontal", "vertical", "zPattern", "nPattern", "stackOrder"];

    /* 並び順ごとの比較関数。top は上にあるほど値が大きいので、降順で「上から下」になる
       Comparator per sort order; top grows upward, so descending means top-to-bottom */
    var SORT_COMPARATORS = {
        currentValue: function (entryA, entryB) {
            return entryA.sortValue - entryB.sortValue;
        },
        horizontal: function (entryA, entryB) {
            return entryA.left - entryB.left;
        },
        vertical: function (entryA, entryB) {
            return entryB.top - entryA.top;
        },
        zPattern: function (entryA, entryB) {
            /* 行が違うときだけ上下で比べ、同じ行なら左から右 / Compare rows first, then left to right */
            if (Math.abs(entryA.top - entryB.top) > ROW_COLUMN_TOLERANCE_PT) return entryB.top - entryA.top;
            return entryA.left - entryB.left;
        },
        nPattern: function (entryA, entryB) {
            /* 列が違うときだけ左右で比べ、同じ列なら上から下 / Compare columns first, then top to bottom */
            if (Math.abs(entryA.left - entryB.left) > ROW_COLUMN_TOLERANCE_PT) return entryA.left - entryB.left;
            return entryB.top - entryA.top;
        },
        stackOrder: function (entryA, entryB) {
            /* 前面が先。レイヤーをまたぐときはレイヤーの重ね順を先に見る
               Frontmost first; compare the layer order first when the layers differ */
            if (entryA.layerOrder !== entryB.layerOrder) return entryB.layerOrder - entryA.layerOrder;
            if (entryA.itemOrder !== entryB.itemOrder) return entryB.itemOrder - entryA.itemOrder;
            /* 重ね順が読めなかったときは選択順で安定させる / Fall back to the selection order */
            return entryA.selectionIndex - entryB.selectionIndex;
        }
    };

    // =========================================
    // 連番の書式 / Sequence format
    // =========================================

    /* ［基準となる値］の書式ラジオの並び。LABELS.radio のキーと揃える
       Format radio sequence; the keys match LABELS.radio */
    var FORMAT_MODES = ["digit", "upperCase", "lowerCase", "kanji", "roman", "daiji"];

    /* 書式ラジオを選んだときに［開始値］に入れる値 / Value dropped into the start value when a format radio is chosen */
    var FORMAT_FIRST_VALUES = { digit: "1", upperCase: "A", lowerCase: "a", kanji: "一", roman: "I", daiji: "壱" };

    /* 英字だけのテキスト（A、Z、AA など） / Text made of letters only */
    var LETTERS_ONLY_PATTERN = /^[A-Za-z]+$/;

    /* 漢数字だけのテキスト（一、十二、弐、参拾 など）。通常と大字のどちらも受ける
       Text made of Japanese numerals only, plain or formal */
    var KANJI_ONLY_PATTERN = /^[〇零一二三四五六七八九十拾百千万萬壱弐参]+$/;

    /* 大字を含むか。含めば書き出しも大字にする / Formal numerals; their presence picks the formal output */
    var DAIJI_PATTERN = /[零壱弐参拾萬]/;

    /* 書き出しに使う漢数字。大字は一二三十を壱弐参拾に置き換える（法令で定められている4文字）
       Numerals used for output; the formal style replaces 一二三十 with 壱弐参拾 */
    var KANJI_STYLES = {
        kanji: {
            digits: ["〇", "一", "二", "三", "四", "五", "六", "七", "八", "九"],
            units: [
                { value: 1000, character: "千" },
                { value: 100, character: "百" },
                { value: 10, character: "十" }
            ]
        },
        daiji: {
            digits: ["零", "壱", "弐", "参", "四", "五", "六", "七", "八", "九"],
            units: [
                { value: 1000, character: "千" },
                { value: 100, character: "百" },
                { value: 10, character: "拾" }
            ]
        }
    };

    /* 読み取り用。通常・大字のどちらの表記も受ける / Reading maps covering both styles */
    var KANJI_VALUES = {
        "〇": 0, "零": 0, "一": 1, "壱": 1, "二": 2, "弐": 2, "三": 3, "参": 3,
        "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9
    };
    var KANJI_UNIT_VALUES = { "十": 10, "拾": 10, "百": 100, "千": 1000 };

    /* ローマ数字（I、IV、XII など）。入力は小文字でも受け付け、書き出しは大文字に揃える
       Roman numerals such as I, IV and XII; typed in either case, always written in capitals */
    var ROMAN_PATTERN = /^m*(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$/i;

    /* ローマ数字の表記。大きい順に並べる / Roman numeral spellings, largest first */
    var ROMAN_UNITS = [
        { value: 1000, characters: "M" }, { value: 900, characters: "CM" },
        { value: 500, characters: "D" }, { value: 400, characters: "CD" },
        { value: 100, characters: "C" }, { value: 90, characters: "XC" },
        { value: 50, characters: "L" }, { value: 40, characters: "XL" },
        { value: 10, characters: "X" }, { value: 9, characters: "IX" },
        { value: 5, characters: "V" }, { value: 4, characters: "IV" },
        { value: 1, characters: "I" }
    ];

    /**
     * 前後の空白を落とします。
     *
     * @param {string} sourceText - 対象の文字列。
     * @returns {string} 前後の空白を除いた文字列。
     */
    function trimText(sourceText) {
        return String(sourceText).replace(/^\s+|\s+$/g, "");
    }

    /**
     * 数字だけのテキスト（"12"、"-3"、"1.5" など）を数に換算します。
     *
     * @param {string} trimmedText - 前後の空白を除いた文字列。
     * @returns {number} 換算した数。数字として読めないときは null。
     */
    function parsePlainNumber(trimmedText) {
        var numberValue = parseFloat(trimmedText);
        if (isNaN(numberValue) || !isFinite(trimmedText)) return null;
        return numberValue;
    }

    /**
     * 英字を1から始まる位置に換算します（A=1、Z=26、AA=27）。
     *
     * @param {string} letters - 英字だけの文字列。
     * @returns {number} 文字順の位置。
     */
    function lettersToIndex(letters) {
        var upperCaseLetters = letters.toUpperCase();
        var letterIndex = 0;
        for (var i = 0; i < upperCaseLetters.length; i++) {
            letterIndex = letterIndex * 26 + (upperCaseLetters.charCodeAt(i) - 64);
        }
        return letterIndex;
    }

    /**
     * 位置を英字に戻します（1=A、26=Z、27=AA）。
     *
     * @param {number} letterIndex - 1から始まる文字順の位置。
     * @returns {string} 大文字の英字。1以下は "A"。
     */
    function indexToLetters(letterIndex) {
        var letters = "";
        var remaining = Math.floor(letterIndex);
        while (remaining > 0) {
            var remainder = (remaining - 1) % 26;
            letters = String.fromCharCode(65 + remainder) + letters;
            remaining = Math.floor((remaining - 1) / 26);
        }
        return letters || "A";
    }

    /**
     * 漢数字を数に換算します（十二=12、二百三=203、一万=10000）。大字（壱弐参拾萬）も読めます。
     *
     * @param {string} kanjiText - 漢数字だけの文字列。
     * @returns {number} 換算した数。解釈できないときは null。
     */
    function kanjiToNumber(kanjiText) {
        var settledTotal = 0;  /* 万より上の確定分 / settled part above 10000 */
        var sectionTotal = 0;  /* 千・百・十の合計 / sum of the thousand, hundred and ten places */
        var pendingDigit = 0;  /* 直前の1桁 / the digit just read */
        var hasDigit = false;

        for (var i = 0; i < kanjiText.length; i++) {
            var character = kanjiText.charAt(i);

            if (KANJI_VALUES[character] !== undefined) {
                pendingDigit = KANJI_VALUES[character];
                hasDigit = true;
                continue;
            }

            if (character === "万" || character === "萬") {
                settledTotal += (sectionTotal + pendingDigit) * 10000;
                sectionTotal = 0;
                pendingDigit = 0;
                continue;
            }

            var unitValue = KANJI_UNIT_VALUES[character];
            if (unitValue === undefined) return null;

            /* 「十」のように数字を伴わない位は1と見なす / A bare unit such as 十 counts as one */
            sectionTotal += (pendingDigit || 1) * unitValue;
            pendingDigit = 0;
            hasDigit = true;
        }

        return hasDigit ? (settledTotal + sectionTotal + pendingDigit) : null;
    }

    /**
     * 数を漢数字に直します（12=十二、203=二百三、10000=一万）。大字では 12=拾弐 になります。
     *
     * @param {number} numberValue - 0以上の整数。
     * @param {string} styleKey - "kanji"（通常）または "daiji"（大字）。
     * @returns {string} 漢数字。0以下は "〇"（大字は "零"）。
     */
    function numberToKanji(numberValue, styleKey) {
        var kanjiStyle = KANJI_STYLES[styleKey] || KANJI_STYLES.kanji;
        var remaining = Math.floor(numberValue);
        if (remaining <= 0) return kanjiStyle.digits[0];

        var kanjiText = "";

        /* 万の位から先に切り出す / Take the ten-thousands apart first */
        if (remaining >= 10000) {
            kanjiText += numberToKanji(Math.floor(remaining / 10000), styleKey) + "万";
            remaining = remaining % 10000;
            if (remaining === 0) return kanjiText;
        }

        for (var i = 0; i < kanjiStyle.units.length; i++) {
            var digit = Math.floor(remaining / kanjiStyle.units[i].value);
            if (digit > 0) {
                /* 十・百・千の「一」は書かない / The leading one is left out for 十, 百 and 千 */
                if (digit > 1) kanjiText += kanjiStyle.digits[digit];
                kanjiText += kanjiStyle.units[i].character;
                remaining -= digit * kanjiStyle.units[i].value;
            }
        }

        return kanjiText + ((remaining > 0) ? kanjiStyle.digits[remaining] : "");
    }

    /**
     * ローマ数字を数に換算します（IV=4、XII=12）。
     *
     * @param {string} romanText - ローマ数字の文字列。
     * @returns {number} 換算した数。
     */
    function romanToNumber(romanText) {
        var upperCaseText = romanText.toUpperCase();
        var total = 0;
        var position = 0;

        while (position < upperCaseText.length) {
            for (var i = 0; i < ROMAN_UNITS.length; i++) {
                var unitCharacters = ROMAN_UNITS[i].characters;
                if (upperCaseText.substr(position, unitCharacters.length) === unitCharacters) {
                    total += ROMAN_UNITS[i].value;
                    position += unitCharacters.length;
                    break;
                }
            }
        }

        return total;
    }

    /**
     * 数をローマ数字に直します（4=IV、12=XII）。
     *
     * @param {number} numberValue - 1以上の整数。
     * @returns {string} 大文字のローマ数字。1未満は "I"。
     */
    function numberToRoman(numberValue) {
        var remaining = Math.floor(numberValue);
        if (remaining < 1) return "I";

        var romanText = "";
        for (var i = 0; i < ROMAN_UNITS.length; i++) {
            while (remaining >= ROMAN_UNITS[i].value) {
                romanText += ROMAN_UNITS[i].characters;
                remaining -= ROMAN_UNITS[i].value;
            }
        }
        return romanText;
    }

    /**
     * 並べ替えに使う値を返します。数字はその値、英字は文字順の位置、漢数字は換算した数です。
     *
     * @param {string} textContents - テキストの中身。
     * @returns {number} 並べ替え用の値。対象外のときは null。
     */
    function getSortValue(textContents) {
        var trimmedText = trimText(textContents);
        if (trimmedText === "") return null;
        if (LETTERS_ONLY_PATTERN.test(trimmedText)) return lettersToIndex(trimmedText);
        if (KANJI_ONLY_PATTERN.test(trimmedText)) return kanjiToNumber(trimmedText);
        return parsePlainNumber(trimmedText);
    }

    /**
     * ［開始値］を解釈します。数字（0や負の値も可）、英字、漢数字、ローマ数字、大字を受け付けます。
     *
     * @param {string} startValueText - ［開始値］の入力。
     * @param {string} formatModeKey - 選ばれている書式ラジオのキー（紛らわしい表記の判断に使う）。
     * @returns {Object} format（"number" / "letter" / "kanji" / "daiji" / "roman"）と値を持つオブジェクト。
     *                   数字のときは入力した整数部の桁数を digits に入れます。解釈できないときは null。
     */
    function parseStartValue(startValueText, formatModeKey) {
        var trimmedText = trimText(startValueText);
        if (trimmedText === "") return null;

        /* I や X は英字とも読めるため、ローマ数字かどうかはラジオの選択で決める
           Letters like I and X are ambiguous, so the radio decides whether it is Roman */
        if (formatModeKey === "roman" && ROMAN_PATTERN.test(trimmedText)) {
            return { format: "roman", number: romanToNumber(trimmedText) };
        }

        if (LETTERS_ONLY_PATTERN.test(trimmedText)) {
            return {
                format: "letter",
                index: lettersToIndex(trimmedText),
                isUpperCase: (trimmedText === trimmedText.toUpperCase())
            };
        }

        if (KANJI_ONLY_PATTERN.test(trimmedText)) {
            var kanjiNumber = kanjiToNumber(trimmedText);
            if (kanjiNumber === null) return null;
            /* 四〜九は通常と大字で同じ字なので、判別できないときはラジオの選択に従う
               四 to 九 are shared by both styles, so the radio decides when the text cannot tell */
            var isDaiji = DAIJI_PATTERN.test(trimmedText) ||
                (formatModeKey === "daiji" && !/[〇一二三十]/.test(trimmedText));
            return { format: isDaiji ? "daiji" : "kanji", number: kanjiNumber };
        }

        var startNumber = parsePlainNumber(trimmedText);
        if (startNumber === null) return null;

        /* 入力した整数部の桁数を控える。"01" なら 2 で、以降も2桁で書き出す
           Remember how many integer digits were typed: "01" means two, and stays two */
        var integerDigits = /^[-+]?(\d+)/.exec(trimmedText);
        return {
            format: "number",
            number: startNumber,
            digits: integerDigits ? integerDigits[1].length : 1
        };
    }

    /**
     * ［開始値］に対応する書式ラジオのキーを返します。
     *
     * @param {Object} startValue - parseStartValue() の戻り値。
     * @returns {string} FORMAT_MODES のキー。判定できないときは null。
     */
    function getFormatModeKey(startValue) {
        if (!startValue) return null;
        if (startValue.format === "roman") return "roman";
        if (startValue.format === "daiji") return "daiji";
        if (startValue.format === "kanji") return "kanji";
        if (startValue.format === "letter") return startValue.isUpperCase ? "upperCase" : "lowerCase";
        return "digit";
    }

    /**
     * ［開始値］から数えて sequenceOffset 番目の値を文字列で返します。
     *
     * @param {Object} startValue - parseStartValue() の戻り値。
     * @param {number} sequenceOffset - ［開始値］からの位置（0が最初）。
     * @param {number} zeroPadDigits - ゼロ埋めの桁数。0で埋めません。
     * @returns {string} 書き込む文字列。
     */
    function formatSequenceValue(startValue, sequenceOffset, zeroPadDigits) {
        if (startValue.format === "letter") {
            var letters = indexToLetters(startValue.index + sequenceOffset);
            return startValue.isUpperCase ? letters : letters.toLowerCase();
        }

        if (startValue.format === "kanji" || startValue.format === "daiji") {
            return numberToKanji(startValue.number + sequenceOffset, startValue.format);
        }

        if (startValue.format === "roman") {
            return numberToRoman(startValue.number + sequenceOffset);
        }

        var assignedNumber = startValue.number + sequenceOffset;
        return (zeroPadDigits > 1) ? padWithZeros(assignedNumber, zeroPadDigits) : String(assignedNumber);
    }

    /**
     * 書き出す整数部の桁数を返します。入力した桁数（"01" なら2桁）は常に保ち、
     * ［ゼロ埋め］がONのときは最後の番号の桁数まで広げます。
     *
     * @param {Object} startValue - parseStartValue() の戻り値。
     * @param {number} targetCount - 振り直す個数。
     * @param {boolean} isZeroPadded - ［ゼロ埋め］がONか。
     * @returns {number} ゼロ埋めする桁数。埋めないときは0。
     */
    function getSequenceDigits(startValue, targetCount, isZeroPadded) {
        if (startValue.format !== "number") return 0;
        if (!isZeroPadded) return startValue.digits;
        return Math.max(startValue.digits, getZeroPadDigits(startValue.number, targetCount));
    }

    /**
     * ［ゼロ埋め］をONにすると桁数が変わるかを返します。ディム判定に使います。
     *
     * @param {Object} startValue - parseStartValue() の戻り値。
     * @param {number} targetCount - 振り直す個数。
     * @returns {boolean} 桁数が変わるなら true。
     */
    function willZeroPadApply(startValue, targetCount) {
        if (!startValue || startValue.format !== "number") return false;
        return getZeroPadDigits(startValue.number, targetCount) > startValue.digits;
    }

    /**
     * ゼロ埋めの桁数を返します。最初と最後の番号のうち、桁数の多いほうに合わせます。
     *
     * @param {number} startNumber - 開始番号。
     * @param {number} targetCount - 振り直す個数。
     * @returns {number} 整数部の桁数。
     */
    function getZeroPadDigits(startNumber, targetCount) {
        var firstDigits = countIntegerDigits(startNumber);
        var lastDigits = countIntegerDigits(startNumber + targetCount - 1);
        return Math.max(firstDigits, lastDigits);
    }

    /**
     * 数値の整数部の桁数を返します（符号は数えません）。
     *
     * @param {number} numberValue - 対象の数値。
     * @returns {number} 桁数。
     */
    function countIntegerDigits(numberValue) {
        return String(Math.floor(Math.abs(numberValue))).length;
    }

    /**
     * 整数部を指定の桁数までゼロ埋めした文字列を返します。
     *
     * @param {number} numberValue - 対象の数値。
     * @param {number} minDigits - 整数部の桁数。
     * @returns {string} ゼロ埋めした文字列。
     */
    function padWithZeros(numberValue, minDigits) {
        var sign = (numberValue < 0) ? "-" : "";
        var absoluteText = String(Math.abs(numberValue));
        var dotIndex = absoluteText.indexOf(".");
        var integerText = (dotIndex === -1) ? absoluteText : absoluteText.substring(0, dotIndex);
        var decimalText = (dotIndex === -1) ? "" : absoluteText.substring(dotIndex);

        while (integerText.length < minDigits) {
            integerText = "0" + integerText;
        }
        return sign + integerText + decimalText;
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

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "連番振り直し", en: "Renumber Sequence" }
        },
        panel: {
            baseValue: { ja: "基準となる値", en: "Base Value" },
            sortOrder: { ja: "並び順", en: "Sort Order" },
            options: { ja: "オプション", en: "Options" },
            addText: { ja: "テキスト追加", en: "Add Text" }
        },
        fieldLabel: {
            startValue: { ja: "開始値", en: "Start Value" },
            prefix: { ja: "接頭辞", en: "Prefix" },
            suffix: { ja: "接尾辞", en: "Suffix" }
        },
        radio: {
            digit: { ja: "123", en: "123" },
            upperCase: { ja: "ABC", en: "ABC" },
            lowerCase: { ja: "abc", en: "abc" },
            kanji: { ja: "一二三", en: "一二三" },
            roman: { ja: "I II III", en: "I II III" },
            daiji: { ja: "壱弐参", en: "壱弐参" },
            currentValue: { ja: "現在の値順", en: "Current Value Order" },
            horizontal: { ja: "水平方向（左から右）", en: "Horizontal (Left to Right)" },
            vertical: { ja: "垂直方向（上から下）", en: "Vertical (Top to Bottom)" },
            zPattern: { ja: "Z方向（左→右、上→下）", en: "Z-Pattern (Left to Right, Top to Bottom)" },
            nPattern: { ja: "N方向（上→下、左→右）", en: "N-Pattern (Top to Bottom, Left to Right)" },
            stackOrder: { ja: "重ね順（前面から）", en: "Stacking Order (Front to Back)" }
        },
        checkbox: {
            reverse: { ja: "逆順", en: "Reverse" },
            zeroPad: { ja: "ゼロ埋め", en: "Zero Padding" },
            reorderStack: { ja: "重ね順調整（OK時）", en: "Adjust Stacking (on OK)" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Please select text objects." },
            noTargetText: {
                ja: "数字・英字・漢数字だけのテキストオブジェクトが見つかりませんでした。",
                en: "No text objects containing only a number, letters, or Japanese numerals were found."
            }
        },
        tooltip: {
            formatMode: {
                ja: "書き出す書式を選びます。選ぶと［開始値］に 1／A／a／一／I／壱 が入ります。",
                en: "Picks the format; choosing one drops 1, A, a, 一, I, or 壱 into the start value."
            },
            startValue: {
                ja: "振り直しの最初の値です。数字（1、0、-3）、英字（A、a）、漢数字（一）、ローマ数字（I）、大字（壱）を指定できます。",
                en: "First value: a number (1, 0, -3), a letter (A, a), a Japanese numeral (一), a Roman numeral (I), or a formal numeral (壱)."
            },
            startValueStep: {
                ja: "∧∨または↑↓キーで増減します。数字は次の整数へ（shift＋で10の倍数へ、option＋で0.1ずつ）、英字・漢数字・ローマ数字・大字は1つずつ進みます。",
                en: "The arrows or the Up/Down keys step the value: numbers go to the next whole number (Shift snaps to 10s, Option steps by 0.1); letters and numerals step one at a time."
            },
            currentValue: {
                ja: "いま入っている値の小さい順に振り直します（英字はアルファベット順）。",
                en: "Renumbers in ascending order of the current values; letters go in alphabetical order."
            },
            horizontal: {
                ja: "左にあるものから順に振り直します。",
                en: "Renumbers from the leftmost object to the right."
            },
            vertical: { ja: "上にあるものから順に振り直します。", en: "Renumbers from the topmost object down." },
            zPattern: {
                ja: "行ごとに左から右へ、行を上から下へ進みます。",
                en: "Moves left to right within a row, then down to the next row."
            },
            nPattern: {
                ja: "列ごとに上から下へ、列を左から右へ進みます。",
                en: "Moves top to bottom within a column, then right to the next column."
            },
            stackOrder: {
                ja: "重ね順の前面にあるものから順に振り直します。",
                en: "Renumbers from the frontmost object backward."
            },
            reverse: { ja: "並び順を逆さにして振り直します。", en: "Reverses the order before renumbering." },
            zeroPad: {
                ja: "いちばん桁数の多い番号に合わせて、頭に0を足します。入力した桁数（01なら2桁）は、このチェックに関わらず保ちます。",
                en: "Pads with leading zeros to match the widest number. The typed width (two digits for 01) is kept either way."
            },
            reorderStack: {
                ja: "振り直した番号の順に重ね順を並べ替えます（番号の小さいものが前面）。［OK］のときだけ実行します。",
                en: "Reorders the stacking to follow the new numbers, lowest number frontmost. Applied on OK only."
            },
            prefix: { ja: "番号の前に付ける文字列です。", en: "Text placed before the number." },
            suffix: { ja: "番号の後ろに付ける文字列です。", en: "Text placed after the number." },
            /* ステップボタン用（StepperButtons 部品から写す） / For the stepper (copied from the StepperButtons part) */
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

    // =========================================
    // メイン処理 / Main
    // =========================================

    main();

    /**
     * ドキュメントと選択を確認し、振り直しのダイアログを開きます。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var activeDoc = app.activeDocument;
        var selectedItems = activeDoc.selection;
        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var renumberTargets = collectSequenceTargets(selectedItems);
        if (renumberTargets.length === 0) {
            alert(getLabel(LABELS.alert.noTargetText));
            return;
        }

        showRenumberDialog(renumberTargets);
    }

    /**
     * 選択の中から、数字・英字・漢数字だけが入ったテキストオブジェクトを集めます。
     *
     * 並べ替えに使う座標と重ね順は、プレビューで中身が変わる前に控えます。
     *
     * @param {Object[]} selectedItems - ドキュメントの選択。
     * @returns {Object[]} 振り直し対象（frame / originalContents / text / sortValue / left / top /
     *                   layerOrder / itemOrder / selectionIndex）。
     */
    function collectSequenceTargets(selectedItems) {
        var renumberTargets = [];

        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem.typename !== "TextFrame") continue;

            /* 中身が数字・英字・漢数字だけのものに限る（"3行目" などは対象外）
               Only text that is nothing but digits, letters, or Japanese numerals */
            var sortValue = getSortValue(selectedItem.contents);
            if (sortValue === null) continue;

            var stackOrderKey = getStackOrderKey(selectedItem);
            renumberTargets.push({
                frame: selectedItem,
                originalContents: selectedItem.contents,
                text: trimText(selectedItem.contents),
                sortValue: sortValue,
                left: selectedItem.left,
                top: selectedItem.top,
                layerOrder: stackOrderKey.layerOrder,
                itemOrder: stackOrderKey.itemOrder,
                selectionIndex: i
            });
        }

        return renumberTargets;
    }

    /**
     * 重ね順の比較に使うキーを返します。値が大きいほど前面です。
     *
     * @param {PageItem} pageItem - 対象のオブジェクト。
     * @returns {{layerOrder: number, itemOrder: number}} レイヤーとレイヤー内の重ね順。
     */
    function getStackOrderKey(pageItem) {
        /* zOrderPosition を持たないアイテムもあるため、まとめて保護する
           Some items expose no zOrderPosition, so guard the whole lookup */
        try {
            return { layerOrder: pageItem.layer.zOrderPosition, itemOrder: pageItem.zOrderPosition };
        } catch (e) {
            return { layerOrder: 0, itemOrder: 0 };
        }
    }

    /**
     * ［開始値］の初期値を返します。いちばん小さい値のテキストをそのまま使うので、
     * 選択が英字なら英字、漢数字なら漢数字で始まります。
     *
     * @param {Object[]} renumberTargets - 振り直し対象。
     * @returns {string} 初期値のテキスト。
     */
    function getInitialStartValue(renumberTargets) {
        var smallestTarget = renumberTargets[0];
        for (var i = 1; i < renumberTargets.length; i++) {
            if (renumberTargets[i].sortValue < smallestTarget.sortValue) smallestTarget = renumberTargets[i];
        }
        return smallestTarget.text;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 連番振り直しのダイアログを表示します。
     *
     * @param {Object[]} renumberTargets - 振り直し対象。
     * @returns {void}
     */
    function showRenumberDialog(renumberTargets) {
        var dialogControls = buildRenumberDialog(getInitialStartValue(renumberTargets));
        bindRenumberDialogEvents(dialogControls, renumberTargets);
        prepareDialogWindow(dialogControls.renumberDialog, SCRIPT_NAME);
        dialogControls.renumberDialog.show();
    }

    /**
     * 連番振り直しのダイアログを組み立てます（イベントの配線は bindRenumberDialogEvents() で行います）。
     *
     * @param {string} initialStartValue - ［開始値］の初期値。
     * @returns {{renumberDialog: Window, startValueInput: EditText, formatRadios: Object,
     *            optionCheckboxes: {reverse: Checkbox, zeroPad: Checkbox},
     *            sortOrderRadios: Object, reorderStackCheckbox: Checkbox,
     *            affixInputs: {prefix: EditText, suffix: EditText}, btnCancel: Button, btnOK: Button}} 作成したコントロール。
     */
    function buildRenumberDialog(initialStartValue) {
        var renumberDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(renumberDialog);

        /* 上段：2カラム / Upper area: two columns */
        var columnsGroup = renumberDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 左カラム：基準となる値とオプション / Left column: base value and options */
        var leftColumn = addDialogColumn(columnsGroup);
        var baseValueControls = addBaseValuePanel(leftColumn, initialStartValue);
        var optionsPanel = addDialogPanel(leftColumn, LABELS.panel.options, 6);
        var optionCheckboxes = {
            reverse: addLabeledCheckbox(optionsPanel, "reverse"),
            zeroPad: addLabeledCheckbox(optionsPanel, "zeroPad")
        };

        /* 右カラム：並び順とテキスト追加 / Right column: sort order and affixes */
        var rightColumn = addDialogColumn(columnsGroup);
        var sortOrderControls = addSortOrderPanel(rightColumn);
        var affixInputs = addAffixPanel(rightColumn);

        /* 下段：ボタン行 / Lower area: button row */
        var buttonRow = addButtonRow(renumberDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            renumberDialog: renumberDialog,
            startValueInput: baseValueControls.startValueInput,
            formatRadios: baseValueControls.formatRadios,
            optionCheckboxes: optionCheckboxes,
            sortOrderRadios: sortOrderControls.sortOrderRadios,
            reorderStackCheckbox: sortOrderControls.reorderStackCheckbox,
            affixInputs: affixInputs,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 上段のカラム（パネルを縦に積むグループ）を追加します。
     *
     * @param {Group} columnsGroup - 2カラムを並べるグループ。
     * @returns {Group} 追加したカラム。
     */
    function addDialogColumn(columnsGroup) {
        var dialogColumn = columnsGroup.add("group");
        dialogColumn.orientation = "column";
        dialogColumn.alignChildren = ["fill", "top"];
        dialogColumn.spacing = WINDOW_SPACING;
        return dialogColumn;
    }

    /**
     * 見出し付きのパネルを追加して、共通のレイアウトを当てます。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @param {Object} titleLabelSet - パネル見出しの文言オブジェクト。
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）。
     * @returns {Panel} 追加したパネル。
     */
    function addDialogPanel(parentGroup, titleLabelSet, spacing) {
        var dialogPanel = parentGroup.add("panel", undefined, getLabel(titleLabelSet));
        setupPanel(dialogPanel, spacing);
        return dialogPanel;
    }

    /**
     * 上マージンで前の段と分けた、縦並びのグループを追加します。
     *
     * @param {Panel} parentPanel - 追加先のパネル。
     * @returns {Group} 追加したグループ。
     */
    function addSeparatedSection(parentPanel) {
        var sectionGroup = parentPanel.add("group");
        sectionGroup.orientation = "column";
        sectionGroup.alignChildren = ["left", "top"];
        sectionGroup.margins = [0, SECTION_TOP_MARGIN, 0, 0];
        return sectionGroup;
    }

    /**
     * ラジオボタンを縦に並べて追加します（同じ親に入れるので排他になります）。
     *
     * @param {Panel} parentPanel - 追加先のパネル。
     * @param {string[]} radioKeys - LABELS.radio のキーの並び。
     * @param {Object} [sharedTooltip] - 全ラジオ共通の tooltip。省略時は LABELS.tooltip の同じキーを使います。
     * @returns {Object} キーごとのラジオボタン。
     */
    function addRadioButtons(parentPanel, radioKeys, sharedTooltip) {
        var radiosByKey = {};
        for (var i = 0; i < radioKeys.length; i++) {
            var radioKey = radioKeys[i];
            var radioButton = parentPanel.add("radiobutton", undefined, getLabel(LABELS.radio[radioKey]));
            radioButton.helpTip = getLabel(sharedTooltip || LABELS.tooltip[radioKey]);
            radiosByKey[radioKey] = radioButton;
        }
        return radiosByKey;
    }

    /**
     * チェックボックスを、LABELS のキーで tooltip 付きで追加します。
     *
     * @param {Panel|Group} parentContainer - 追加先のパネルかグループ。
     * @param {string} labelKey - LABELS.checkbox と LABELS.tooltip のキー。
     * @returns {Checkbox} 追加したチェックボックス。
     */
    function addLabeledCheckbox(parentContainer, labelKey) {
        var labeledCheckbox = parentContainer.add("checkbox", undefined, getLabel(LABELS.checkbox[labelKey]));
        labeledCheckbox.helpTip = getLabel(LABELS.tooltip[labelKey]);
        return labeledCheckbox;
    }

    /**
     * 選択されているラジオのキーを返します。
     *
     * @param {Object} radiosByKey - キーごとのラジオボタン。
     * @param {string[]} radioKeys - キーの並び（先頭が既定値）。
     * @returns {string} 選ばれているキー。どれも選ばれていなければ先頭のキー。
     */
    function getSelectedRadioKey(radiosByKey, radioKeys) {
        for (var i = 0; i < radioKeys.length; i++) {
            if (radiosByKey[radioKeys[i]].value) return radioKeys[i];
        }
        return radioKeys[0];
    }

    /**
     * ［基準となる値］パネル（書式のラジオと開始値）を作ります。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @param {string} initialStartValue - 初期値。
     * @returns {{startValueInput: EditText, formatRadios: Object}} 開始値の入力欄と書式のラジオ。
     */
    function addBaseValuePanel(parentGroup, initialStartValue) {
        var baseValuePanel = addDialogPanel(parentGroup, LABELS.panel.baseValue, 6);
        var formatRadios = addRadioButtons(baseValuePanel, FORMAT_MODES, LABELS.tooltip.formatMode);

        /* 項目名を上、∧∨と入力欄を下に置く。ラジオとのアキは上マージンで取る
           The label sits above the stepper and field, with a top margin separating it from the radios */
        var startValueGroup = addSeparatedSection(baseValuePanel);
        startValueGroup.spacing = LABEL_FIELD_SPACING;
        startValueGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.startValue));

        return { startValueInput: addStartValueField(startValueGroup, initialStartValue), formatRadios: formatRadios };
    }

    /**
     * ［開始値］の入力欄を、左に∧∨を突き合わせて追加します。
     *
     * ∧∨と↑↓キーはどちらも stepperGroup.stepBy を呼びます（中身は bindStartValueStepper() で差し替えます）。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @param {string} initialStartValue - 初期値。
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）。
     */
    function addStartValueField(parentGroup, initialStartValue) {
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = parentGroup.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var startValueInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return startValueInput; }, {});
        startValueInput = stepperFieldGroup.add("edittext", undefined, initialStartValue);
        startValueInput.characters = START_VALUE_FIELD_CHARS;
        startValueInput.helpTip = getLabel(LABELS.tooltip.startValue) + "\n" + getLabel(LABELS.tooltip.startValueStep);
        startValueInput.active = true;
        startValueInput.stepperGroup = stepperGroup;

        /* 部品の∧∨は内部の増減を直接呼ぶため、差し替えた stepperGroup.stepBy を呼ぶボタンに付け替える
           The part's buttons call its own step function, so rebuild them to call the replaced stepperGroup.stepBy */
        while (stepperGroup.children.length > 0) {
            stepperGroup.remove(stepperGroup.children[0]);
        }
        makeStepperChevronButton(stepperGroup, "up", function () { stepperGroup.stepBy(1); }).helpTip = getLabel(LABELS.tooltip.stepUp);
        makeStepperChevronButton(stepperGroup, "down", function () { stepperGroup.stepBy(-1); }).helpTip = getLabel(LABELS.tooltip.stepDown);

        /* ↑↓キーも∧∨と同じ stepperGroup.stepBy で増減する / arrow keys share the stepper's stepBy */
        bindSteppedArrowKeys(startValueInput, stepperGroup);
        return startValueInput;
    }

    /**
     * ［並び順］パネル（並び順のラジオと重ね順調整）を作ります。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @returns {{sortOrderRadios: Object, reorderStackCheckbox: Checkbox}} 並び順のラジオと重ね順調整のチェックボックス。
     */
    function addSortOrderPanel(parentGroup) {
        var sortOrderPanel = addDialogPanel(parentGroup, LABELS.panel.sortOrder, 6);
        var sortOrderRadios = addRadioButtons(sortOrderPanel, SORT_MODES);
        sortOrderRadios[SORT_MODES[0]].value = true;

        /* 並び順に合わせて重ね順も揃えるので、ラジオの下に置く
           Sits under the radios: it lines the stacking up with the chosen order */
        var reorderStackCheckbox = addLabeledCheckbox(addSeparatedSection(sortOrderPanel), "reorderStack");
        reorderStackCheckbox.value = REORDER_STACK_BY_DEFAULT;

        return { sortOrderRadios: sortOrderRadios, reorderStackCheckbox: reorderStackCheckbox };
    }

    /**
     * ［テキスト追加］パネル（接頭辞・接尾辞）を作ります。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @returns {{prefix: EditText, suffix: EditText}} 作成した入力欄。
     */
    function addAffixPanel(parentGroup) {
        var affixPanel = addDialogPanel(parentGroup, LABELS.panel.addText);
        /* 入力欄の右端をそろえるため、行を右寄せにする / Right-align the rows so the fields share a right edge */
        affixPanel.alignChildren = ["right", "top"];

        return {
            prefix: addAffixRow(affixPanel, LABELS.fieldLabel.prefix, LABELS.tooltip.prefix),
            suffix: addAffixRow(affixPanel, LABELS.fieldLabel.suffix, LABELS.tooltip.suffix)
        };
    }

    /**
     * 接頭辞・接尾辞の1行を作ります。
     *
     * @param {Panel} parentPanel - 追加先のパネル。
     * @param {Object} fieldLabelSet - 項目名の文言オブジェクト。
     * @param {Object} tooltipSet - tooltipの文言オブジェクト。
     * @returns {EditText} 作成した入力欄。
     */
    function addAffixRow(parentPanel, fieldLabelSet, tooltipSet) {
        var affixRow = parentPanel.add("group");
        affixRow.add("statictext", undefined, labelText(fieldLabelSet));

        var affixInput = affixRow.add("edittext", undefined, "");
        affixInput.characters = AFFIX_FIELD_CHARS;
        affixInput.helpTip = getLabel(tooltipSet);

        return affixInput;
    }

    // =========================================
    // プレビューとイベント / Preview and events
    // =========================================

    /**
     * ダイアログのイベントを配線し、初期値でプレビューします。
     *
     * @param {Object} dialogControls - buildRenumberDialog() の戻り値。
     * @param {Object[]} renumberTargets - 振り直し対象。
     * @returns {void}
     */
    function bindRenumberDialogEvents(dialogControls, renumberTargets) {
        var renumberDialog = dialogControls.renumberDialog;
        var startValueInput = dialogControls.startValueInput;
        var formatRadios = dialogControls.formatRadios;
        var optionCheckboxes = dialogControls.optionCheckboxes;
        var sortOrderRadios = dialogControls.sortOrderRadios;
        var affixInputs = dialogControls.affixInputs;

        /* ［OK］で閉じたかどうか。×やキャンセルのときだけ元に戻す
           Whether OK was used; only a cancel or a close restores the originals */
        var isConfirmed = false;

        /**
         * ダイアログの入力値をまとめて読み取ります。
         *
         * @returns {Object} applyRenumber() に渡す設定。
         */
        function getRenumberSettings() {
            return {
                startValueText: startValueInput.text,
                formatModeKey: getSelectedRadioKey(formatRadios, FORMAT_MODES),
                prefix: affixInputs.prefix.text,
                suffix: affixInputs.suffix.text,
                sortMode: getSelectedRadioKey(sortOrderRadios, SORT_MODES),
                isReversed: optionCheckboxes.reverse.value,
                isZeroPadded: optionCheckboxes.zeroPad.value,
                isStackReordered: dialogControls.reorderStackCheckbox.value
            };
        }

        /**
         * いまの入力でプレビューを更新します（書式のラジオとゼロ埋めのディムも合わせます）。
         *
         * @returns {void}
         */
        function updatePreview() {
            var renumberSettings = getRenumberSettings();
            var startValue = parseStartValue(renumberSettings.startValueText, renumberSettings.formatModeKey);

            /* 入力に合わせて書式のラジオを合わせる / Keep the format radios in step with the input */
            var formatModeKey = getFormatModeKey(startValue);
            if (formatModeKey) {
                formatRadios[formatModeKey].value = true;
                renumberSettings.formatModeKey = formatModeKey;
            }

            /* ゼロ埋めが効かない書式・桁数のときはディム / Dim zero padding when it would change nothing */
            optionCheckboxes.zeroPad.enabled = willZeroPadApply(startValue, renumberTargets.length);

            /* ［開始値］を読めないあいだは元の中身に戻しておく
               While the start value cannot be read, put the original contents back */
            if (!applyRenumber(renumberTargets, renumberSettings)) {
                restoreOriginalContents(renumberTargets);
            }
            app.redraw();
        }

        dialogControls.btnCancel.onClick = function () {
            renumberDialog.close(0);
        };

        dialogControls.btnOK.onClick = function () {
            /* 元に戻してから一度だけ適用する。最後の書き込みがまとまるので、Undoで元の状態まで戻せる
               Restore first, then apply once, so undo lands back on the original text */
            restoreOriginalContents(renumberTargets);

            var renumberSettings = getRenumberSettings();
            var orderedTargets = applyRenumber(renumberTargets, renumberSettings);
            /* 重ね順の並べ替えはプレビューでは行わない（控えた重ね順が狂うため）
               Restacking runs here only: doing it in the preview would stale the cached order */
            if (orderedTargets && renumberSettings.isStackReordered) reorderStackToMatch(orderedTargets);

            isConfirmed = true;
            app.redraw();
            renumberDialog.close(1);
        };

        /* キャンセル・×で閉じたときは元の中身に戻す / A cancel or a close puts the originals back */
        renumberDialog.onClose = function () {
            if (!isConfirmed) {
                restoreOriginalContents(renumberTargets);
                app.redraw();
            }
            return true;
        };

        /**
         * 書式のラジオを選んだら、その書式のいちばん若い値を［開始値］に入れるハンドラを作ります。
         *
         * @param {string} formatModeKey - 書式ラジオのキー。
         * @returns {Function} onClick に設定するハンドラ。
         */
        function makeFormatClickHandler(formatModeKey) {
            return function () {
                startValueInput.text = FORMAT_FIRST_VALUES[formatModeKey];
                updatePreview();
            };
        }
        for (var i = 0; i < FORMAT_MODES.length; i++) {
            formatRadios[FORMAT_MODES[i]].onClick = makeFormatClickHandler(FORMAT_MODES[i]);
        }

        bindStartValueStepper(startValueInput, formatRadios, updatePreview);

        startValueInput.onChanging = updatePreview;
        affixInputs.prefix.onChanging = updatePreview;
        affixInputs.suffix.onChanging = updatePreview;
        optionCheckboxes.reverse.onClick = updatePreview;
        optionCheckboxes.zeroPad.onClick = updatePreview;
        for (var j = 0; j < SORT_MODES.length; j++) {
            sortOrderRadios[SORT_MODES[j]].onClick = updatePreview;
        }

        updatePreview();
    }

    /**
     * ［開始値］の∧∨と↑↓キーの増減を決めます。
     *
     * 数字は部品の既定の増減（次の整数へ、shift＋で10の倍数へ、option＋で0.1ずつ）で、入力した桁数（"01" なら2桁）は保ちます。
     * 英字・漢数字・ローマ数字・大字は1つずつ進めます／戻します（A→B、Z→AA、十→十一、III→IV）。
     *
     * @param {EditText} startValueInput - addStartValueField() で作った入力欄。
     * @param {Object} formatRadios - 書式のラジオ（紛らわしい表記の判断に使う）。
     * @param {Function} onStepped - 増減したあとに呼ぶ処理（プレビューの更新）。
     * @returns {void}
     */
    function bindStartValueStepper(startValueInput, formatRadios, onStepped) {
        var stepperGroup = startValueInput.stepperGroup;
        var stepNumberBy = stepperGroup.stepBy; /* 部品の既定の増減 / the part's default step */

        stepperGroup.stepBy = function (direction) {
            if (!isStepperEnabledInTree(startValueInput)) return;

            var startValue = parseStartValue(startValueInput.text, getSelectedRadioKey(formatRadios, FORMAT_MODES));
            /* 空欄や読めない値のときは何もしない / Leave an empty or unreadable value alone */
            if (!startValue) return;

            if (startValue.format === "number") {
                stepNumberBy(direction);
                if (startValue.digits > 1) {
                    startValueInput.text = padWithZeros(parseFloat(startValueInput.text), startValue.digits);
                }
            } else {
                startValueInput.text = formatSequenceValue(startValue, direction, 0);
            }

            /* .text への代入では onChanging が発火しないため、明示的に呼ぶ
               Setting .text does not fire onChanging, so refresh by hand */
            onStepped();
        };
    }

    // =========================================
    // 振り直し / Renumbering
    // =========================================

    /**
     * 並べ替えた順に連番を書き込みます。プレビューと確定で共用します。
     *
     * @param {Object[]} renumberTargets - 振り直し対象。
     * @param {Object} renumberSettings - getRenumberSettings() が返す設定。
     * @returns {Object[]} 番号を振った順に並べた配列。［開始値］が不正なときは null。
     */
    function applyRenumber(renumberTargets, renumberSettings) {
        var startValue = parseStartValue(renumberSettings.startValueText, renumberSettings.formatModeKey);
        if (!startValue) return null;

        var orderedTargets = sortTargets(renumberTargets, renumberSettings.sortMode, renumberSettings.isReversed);
        var zeroPadDigits = getSequenceDigits(startValue, orderedTargets.length, renumberSettings.isZeroPadded);

        for (var i = 0; i < orderedTargets.length; i++) {
            var sequenceText = formatSequenceValue(startValue, i, zeroPadDigits);
            orderedTargets[i].frame.contents = renumberSettings.prefix + sequenceText + renumberSettings.suffix;
        }

        return orderedTargets;
    }

    /**
     * 指定した並び順に並べ替えた配列を返します（元の配列は変更しません）。
     *
     * @param {Object[]} renumberTargets - 振り直し対象。
     * @param {string} sortMode - SORT_MODES のいずれか。
     * @param {boolean} isReversed - 並びを逆さにするか。
     * @returns {Object[]} 並べ替えた配列。
     */
    function sortTargets(renumberTargets, sortMode, isReversed) {
        var orderedTargets = renumberTargets.slice();
        orderedTargets.sort(SORT_COMPARATORS[sortMode] || SORT_COMPARATORS.currentValue);
        if (isReversed) orderedTargets.reverse();
        return orderedTargets;
    }

    /**
     * 振り直した番号の順に重ね順を並べ替えます（番号の小さいものが前面）。
     *
     * @param {Object[]} orderedTargets - 番号を振った順に並んだ対象。
     * @returns {void}
     */
    function reorderStackToMatch(orderedTargets) {
        /* 後ろから順に最前面へ送ると、先頭＝いちばん若い番号が最前面になる
           Bring each one to the front starting from the last, so the lowest number ends up frontmost */
        for (var i = orderedTargets.length - 1; i >= 0; i--) {
            orderedTargets[i].frame.zOrder(ZOrderMethod.BRINGTOFRONT);
        }
    }

    /**
     * 控えておいた元の中身をすべて書き戻します。
     *
     * app.undo() は「1回の addStep ＝ Undo 1段」が前提で、書き込みが起きなかったときに
     * 段数がずれてユーザーの操作まで取り消してしまうため、控えた文字列で戻します。
     *
     * @param {Object[]} renumberTargets - 振り直し対象。
     * @returns {void}
     */
    function restoreOriginalContents(renumberTargets) {
        for (var i = 0; i < renumberTargets.length; i++) {
            renumberTargets[i].frame.contents = renumberTargets[i].originalContents;
        }
    }

})();
