#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

重なって配置されたオブジェクトを、横方向または縦方向へ指定した間隔で並べ直します。
行数・列数を指定すればタイル状に、キーオブジェクトを設定すればその位置を基準に配置できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf426908d8bcd

### Overview

Redistributes stacked objects along the horizontal or vertical axis at the spacing you specify.
Set a row or column count to tile them, or set a key object to anchor the layout to it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignAndTile";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignAndTile.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignAndTile.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf426908d8bcd"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値 / Initial dialog values */
    var DEFAULT_DIRECTION          = "horizontal"; /* 並べる方向（"horizontal" / "vertical"）/ tiling direction */
    var DEFAULT_LANE_COUNT         = "1";   /* 行数（横）・列数（縦）/ row count (horizontal) or column count (vertical) */
    var DEFAULT_MARGIN             = "0";   /* 横・縦の間隔 / horizontal & vertical spacing */
    var DEFAULT_LINK_MARGINS       = true;  /* 横・縦の間隔を連動 / link both spacings */
    var DEFAULT_USE_PREVIEW_BOUNDS = true;  /* プレビュー境界を使用 / use preview bounds */
    var DEFAULT_USE_GRID           = false; /* グリッド配置 / grid layout */
    var DEFAULT_RANDOMIZE          = false; /* ランダム配置 / random order */

    /* 整列後の位置差をどこまで「動いていない」とみなすか（pt）/ Move tolerance when probing the key object (pt) */
    var KEY_DETECT_TOLERANCE_PT = 0.001;

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

    /* 項目名と数値欄の寸法 / Field label and numeric field sizes */
    var FIELD_LABEL_WIDTH  = 40;               /* 項目名の幅 / width of a field label */
    var FIELD_CHAR_WIDTH   = 3;                /* 数値欄の文字数 / character width of a numeric field */

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle, spacing) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel, spacing);
        return createdPanel;
    }

    /**
     * 右揃えの項目名を追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} fieldLabelText - 表示する項目名（コロン付き）
     * @returns {StaticText} 生成した項目名
     */
    function addFieldLabel(parentContainer, fieldLabelText) {
        var fieldLabel = parentContainer.add("statictext", undefined, fieldLabelText);
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return fieldLabel;
    }

    /**
     * オプションのチェックボックスを追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} checkboxLabel - 表示するラベル
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 生成したチェックボックス
     */
    function addOptionCheckbox(parentContainer, checkboxLabel, initialValue) {
        var createdCheckbox = parentContainer.add("checkbox", undefined, checkboxLabel);
        createdCheckbox.alignment = "left"; /* パネル幅いっぱいに広げない / Do not stretch to the panel width */
        createdCheckbox.value = initialValue;
        return createdCheckbox;
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },
        panel: {
            direction: { ja: "方向", en: "Direction" },
            spacing:   { ja: "間隔", en: "Spacing" },
            alignment: { ja: "揃え", en: "Align" },
            options:   { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            rowCount:    { ja: "行数", en: "Rows" },
            columnCount: { ja: "列数", en: "Cols" },
            hMargin:     { ja: "横", en: "H" },
            vMargin:     { ja: "縦", en: "V" }
        },
        checkbox: {
            useKeyObject:     { ja: "キーオブジェクトを基準", en: "Anchor to key object" },
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            useGrid:          { ja: "グリッド", en: "Grid" },
            randomize:        { ja: "ランダム", en: "Random" }
        },
        radio: {
            directionHorizontal: { ja: "横", en: "Horizontal" },
            directionVertical:   { ja: "縦", en: "Vertical" },
            alignTop:    { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" },
            alignLeft:   { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight:  { ja: "右", en: "Right" },
            alignNone:   { ja: "なし", en: "None" }
        },
        tooltip: {
            directionHorizontal: {
                ja: "左から右へ並べます。行数を指定すると折り返します。",
                en: "Lays the objects out left to right, wrapping at the given number of rows."
            },
            directionVertical: {
                ja: "上から下へ並べます。列数を指定すると折り返します。",
                en: "Stacks the objects top to bottom, wrapping at the given number of columns."
            },
            laneCount: {
                ja: "何行（横並び）／何列（縦並び）で折り返すかを指定します。1 なら折り返しません。",
                en: "How many rows (horizontal) or columns (vertical) to wrap at. 1 means no wrapping."
            },
            useGrid: {
                ja: "各セルの大きさをそろえた格子に配置します。オフのときは各オブジェクトの大きさのまま詰めます。",
                en: "Places the objects on a grid of equal cells. Off packs them at their own sizes."
            },
            hMargin: { ja: "横方向のアキです。↑↓キーで増減できます。", en: "Horizontal gap. The arrow keys step the value." },
            vMargin: { ja: "縦方向のアキです。↑↓キーで増減できます。", en: "Vertical gap. The arrow keys step the value." },
            linkMargins: {
                ja: "横のアキと同じ値を縦にも使います。オフにすると縦を個別に指定できます。",
                en: "Uses the horizontal gap for the vertical one too. Turn it off to set them separately."
            },
            alignVertical: {
                ja: "各行の中でオブジェクトを上下どこにそろえるかです。",
                en: "Where to align the objects vertically within each row."
            },
            alignHorizontal: {
                ja: "各列の中でオブジェクトを左右どこにそろえるかです。",
                en: "Where to align the objects horizontally within each column."
            },
            useKeyObject: {
                ja: "最後にクリックしたキーオブジェクトの位置を動かさずに、他を並べ直します。キーオブジェクトが無いときは選べません。",
                en: "Keeps the key object where it is and arranges the rest around it. Unavailable when there is no key object."
            },
            usePreviewBounds: {
                ja: "線幅や効果を含めた見た目の端を基準にします。オフにするとパスの端が基準になります。",
                en: "Measures by the visible edges including strokes and effects. Off measures the path edges."
            },
            randomize: { ja: "並べる順序をシャッフルします。", en: "Shuffles the order the objects are laid out in." },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:     { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            previewError:    { ja: "プレビューでエラーが発生しました", en: "Preview error" },
            unexpectedError: { ja: "エラーが発生しました", en: "An error has occurred" }
        }
    };

    // =========================================
    // 単位 / Units
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

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

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // =========================================
    // 位置とサイズ / Positions and sizes
    // =========================================

    /**
     * 控えておいた位置へ戻す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {Array} positions - [[left, top], ...] の配列
     * @returns {void}
     */
    function resetPositions(targetItems, positions) {
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].left = positions[i][0];
            targetItems[i].top = positions[i][1];
        }
    }

    /**
     * 指定量だけまとめて移動する
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {number} dx - 横方向の移動量（pt）
     * @param {number} dy - 縦方向の移動量（pt）
     * @returns {void}
     */
    function shiftItems(targetItems, dx, dy) {
        if (!dx && !dy) {
            return;
        }
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            targetItems[i].left += dx;
            targetItems[i].top += dy;
        }
    }

    /**
     * 左端の座標順に並べ替えた複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortedCopyByLeft(targetItems) {
        var copiedItems = targetItems.slice();
        copiedItems.sort(function(itemA, itemB) {
            return itemA.left - itemB.left;
        });
        return copiedItems;
    }

    /**
     * 上端の座標順（上から下、同じなら左から右）に並べ替えた複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortedCopyByTop(targetItems) {
        var copiedItems = targetItems.slice();
        copiedItems.sort(function(itemA, itemB) {
            if (itemA.top !== itemB.top) return itemB.top - itemA.top;
            return itemA.left - itemB.left;
        });
        return copiedItems;
    }

    /**
     * ランダムに並べ替えた複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function shuffledCopy(targetItems) {
        var copiedItems = targetItems.slice();
        for (var i = copiedItems.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapItem = copiedItems[i];
            copiedItems[i] = copiedItems[j];
            copiedItems[j] = swapItem;
        }
        return copiedItems;
    }

    /**
     * 指定したオブジェクトを先頭へ移した複製を返す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} targetItem - 先頭へ移すオブジェクト
     * @returns {PageItem[]} 並べ替えた配列（対象が見つからないときはそのままの複製）
     */
    function movedToFront(targetItems, targetItem) {
        var reordered = targetItems.slice();
        for (var i = 0; i < reordered.length; i++) {
            if (reordered[i] !== targetItem) continue;
            reordered.splice(i, 1);
            reordered.unshift(targetItem);
            break;
        }
        return reordered;
    }

    /**
     * 選択範囲全体の左上を取得する
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @returns {number[]} [左端, 上端]
     */
    function getBlockOrigin(targetItems) {
        var blockLeft = null;
        var blockTop = null;
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            if (blockLeft === null || targetItems[i].left < blockLeft) blockLeft = targetItems[i].left;
            if (blockTop === null || targetItems[i].top > blockTop) blockTop = targetItems[i].top;
        }
        return [blockLeft, blockTop];
    }

    /**
     * もっとも大きいアイテムの幅と高さを取得する（グリッドのセルサイズ）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うかどうか
     * @returns {object} { width: number, height: number }
     */
    function getMaxItemSize(targetItems, usePreviewBounds) {
        var maxWidth = 0;
        var maxHeight = 0;
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            var itemBounds = getClipAwareBounds(targetItems[i], usePreviewBounds);
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];
            if (itemWidth > maxWidth) maxWidth = itemWidth;
            if (itemHeight > maxHeight) maxHeight = itemHeight;
        }
        return { width: maxWidth, height: maxHeight };
    }

    /**
     * セル内での揃え量を求める（X軸は左→右、Y軸は上→下を start→end とする）
     * @param {string} alignMode - "start" / "center" / "end" / "none"
     * @param {number} cellStart - セルの起点（左端または上端）
     * @param {number} cellEnd - セルの終点（右端または下端）
     * @param {number} itemStart - オブジェクトの起点
     * @param {number} itemEnd - オブジェクトの終点
     * @returns {number} 移動量（pt）
     */
    function getAlignDelta(alignMode, cellStart, cellEnd, itemStart, itemEnd) {
        if (alignMode === "none") {
            return 0;
        }
        if (alignMode === "center") {
            return (cellStart + cellEnd) / 2 - (itemStart + itemEnd) / 2;
        }
        if (alignMode === "end") {
            return cellEnd - itemEnd;
        }
        return cellStart - itemStart;
    }

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================
    // Illustrator の DOM にキーオブジェクトを示すプロパティは無いため、整列コマンドで実測して特定する。
    // キーオブジェクトが設定されていると、どの向きに整列してもそのオブジェクトだけは動かない。
    // 候補が0個または2個以上のときは判定不能として null を返し、UI側でこの基準をディムする。

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * @param {PageItem[]} targetItems - 判定対象のオブジェクト
     * @returns {object} キーオブジェクト。判定できないときは null
     */
    function detectKeyObject(targetItems) {
        if (!targetItems || targetItems.length < 2) {
            return null;
        }
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        var i;
        for (i = 0; i < targetItems.length; i++) {
            stayedPut.push(true);
        }

        for (var c = 0; c < alignCommands.length; c++) {
            var savedPositions = [];
            for (i = 0; i < targetItems.length; i++) {
                savedPositions.push([targetItems[i].left, targetItems[i].top]);
            }
            app.redraw(); /* 直前のDOM変更が反映されていないと executeMenuCommand は空振りする / executeMenuCommand misfires without a redraw */
            app.executeMenuCommand(alignCommands[c]);
            for (i = 0; i < targetItems.length; i++) {
                if (Math.abs(targetItems[i].left - savedPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                    Math.abs(targetItems[i].top - savedPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                    stayedPut[i] = false;
                }
            }
            /* 整列は検出のための試行なので、その場で元の位置へ戻す / Undo the probe right away */
            resetPositions(targetItems, savedPositions);
        }
        app.redraw();

        var foundItem = null;
        for (i = 0; i < targetItems.length; i++) {
            if (!stayedPut[i]) continue;
            if (foundItem !== null) return null; /* 複数残った＝判定不能 / Ambiguous */
            foundItem = targetItems[i];
        }
        return foundItem;
    }

    // =========================================
    // プレビュー管理 / Preview management
    // =========================================

    /**
     * プレビューの適用・巻き戻し・確定をまとめて管理する
     * @constructor
     */
    function PreviewManager() {
        /* プレビュー中に実行したアクションの回数 / Number of preview actions executed */
        this.undoDepth = 0;

        /**
         * 変更操作を実行し、履歴としてカウントする
         * @param {Function} previewAction - 実行したい処理
         * @returns {void}
         */
        this.addStep = function(previewAction) {
            try {
                previewAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(labelText("alert.previewError") + " " + e);
            }
        };

        /**
         * プレビューのために行った変更をすべて取り消す
         * @returns {void}
         */
        this.rollback = function() {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * 現在の状態を確定する（プレビューを巻き戻してから1回だけ本番処理を実行）
         * @param {Function} [finalAction] - 巻き戻したあとに実行する処理
         * @returns {void}
         */
        this.confirm = function(finalAction) {
            if (finalAction) {
                this.rollback();
                finalAction();
            }
            this.undoDepth = 0;
        };
    }

    // =========================================
    // 配置処理 / Arranging
    // =========================================

    /**
     * 1行ずつセルに割り当てて配置する
     * @param {PageItem[]} orderedItems - 配置順に並べたオブジェクト
     * @param {object} arrangeSettings - 配置設定
     * @returns {void}
     */
    function placeItems(orderedItems, arrangeSettings) {
        var usePreviewBounds = arrangeSettings.usePreviewBounds;
        var isHorizontal = (arrangeSettings.direction === "horizontal");
        var cellSize = getMaxItemSize(orderedItems, usePreviewBounds);

        /* 先頭オブジェクトの位置を配置の起点にする / The first item defines the origin of the layout */
        var startBounds = getClipAwareBounds(orderedItems[0], usePreviewBounds);
        /* 主軸＝並べる方向、副軸＝行・列が積み重なる方向 / Main axis follows the tiling direction; lanes stack along the cross axis */
        var mainOrigin = isHorizontal ? startBounds[0] : startBounds[1];
        var crossOrigin = isHorizontal ? startBounds[1] : startBounds[0];
        var mainGap = isHorizontal ? arrangeSettings.hMarginPt : arrangeSettings.vMarginPt;
        var crossGap = isHorizontal ? arrangeSettings.vMarginPt : arrangeSettings.hMarginPt;
        var laneSize = isHorizontal ? cellSize.height : cellSize.width;
        /* 上方向がプラスのY軸に合わせ、横並びは下へ、縦並びは右へレーンを送る / Lanes go down (horizontal) or right (vertical) */
        var laneDirection = isHorizontal ? -1 : 1;

        /* 主軸の揃えはグリッド時のみ意味を持つ（セル＝オブジェクトの大きさでは差が出ない）/ Main-axis align only matters in grid mode */
        var mainAlign = arrangeSettings.useGrid ? (isHorizontal ? arrangeSettings.hAlign : arrangeSettings.vAlign) : "start";
        var crossAlign = isHorizontal ? arrangeSettings.vAlign : arrangeSettings.hAlign;

        var remainingItems = orderedItems.length;
        var itemIndex = 0;
        for (var laneIndex = 0; laneIndex < arrangeSettings.laneCount; laneIndex++) {
            /* 残りを残りのレーン数で割り、指定した行数・列数を使い切る / Split the remainder so every lane is used */
            var itemsPerLane = Math.ceil(remainingItems / (arrangeSettings.laneCount - laneIndex));
            remainingItems -= itemsPerLane;
            var laneOffset = laneIndex * (laneSize + crossGap) * laneDirection;
            var crossStart = crossOrigin + laneOffset;
            var crossEnd = crossStart + laneSize * laneDirection;
            var mainStart = mainOrigin;

            for (var i = 0; i < itemsPerLane && itemIndex < orderedItems.length; i++, itemIndex++) {
                var placedItem = orderedItems[itemIndex];
                if (!placedItem) continue;

                var itemBounds = getClipAwareBounds(placedItem, usePreviewBounds);
                var itemMainSize = isHorizontal ? (itemBounds[2] - itemBounds[0]) : (itemBounds[1] - itemBounds[3]);
                var cellMainSize = arrangeSettings.useGrid ? (isHorizontal ? cellSize.width : cellSize.height) : itemMainSize;
                var mainEnd = mainStart + cellMainSize * (isHorizontal ? 1 : -1);

                /* 主軸：セル内での揃え（なし＝その軸は動かさない）/ Main axis: align inside the cell ("none" leaves it alone) */
                var mainDelta = isHorizontal
                    ? getAlignDelta(mainAlign, mainStart, mainEnd, itemBounds[0], itemBounds[2])
                    : getAlignDelta(mainAlign, mainStart, mainEnd, itemBounds[1], itemBounds[3]);
                /* 副軸：レーンの帯へ揃える（なし＝レーンの送り分だけ平行移動）/ Cross axis: align to the lane band ("none" only applies the lane offset) */
                var crossDelta = (crossAlign === "none")
                    ? laneOffset
                    : (isHorizontal
                        ? getAlignDelta(crossAlign, crossStart, crossEnd, itemBounds[1], itemBounds[3])
                        : getAlignDelta(crossAlign, crossStart, crossEnd, itemBounds[0], itemBounds[2]));

                placedItem.left += isHorizontal ? mainDelta : crossDelta;
                placedItem.top += isHorizontal ? crossDelta : mainDelta;

                /* 次のセルへ / Advance to the next cell */
                mainStart = mainEnd + mainGap * (isHorizontal ? 1 : -1);
            }
        }
    }

    /**
     * 設定に従ってオブジェクトを並べ直す（プレビュー・確定の共通処理）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} arrangeSettings - 配置設定
     * @returns {void}
     */
    function arrangeItems(targetItems, arrangeSettings) {
        if (!targetItems || targetItems.length === 0) {
            return;
        }
        var orderedItems;
        if (arrangeSettings.randomize) {
            orderedItems = shuffledCopy(targetItems);
        } else {
            orderedItems = (arrangeSettings.direction === "horizontal") ? sortedCopyByLeft(targetItems) : sortedCopyByTop(targetItems);
        }
        /* キーオブジェクトは配置の起点にする（そこから右／下へ並べる）/ The key object becomes the origin, so the rest follow to its right or below */
        var useKeyObject = arrangeSettings.useKeyObject && arrangeSettings.keyObject && arrangeSettings.keyOrigin;
        if (useKeyObject) {
            orderedItems = movedToFront(orderedItems, arrangeSettings.keyObject);
        }
        /* 並べ替える前の左上（ランダム時に位置を戻す基準）/ Top-left before the layout, used to keep a random block in place */
        var blockOrigin = getBlockOrigin(targetItems);

        placeItems(orderedItems, arrangeSettings);

        /* 基準の補正：キーオブジェクトを優先し、なければランダム時のみ左上を合わせる / Anchor correction: key object first, random block otherwise */
        if (useKeyObject) {
            shiftItems(orderedItems,
                arrangeSettings.keyOrigin[0] - arrangeSettings.keyObject.left,
                arrangeSettings.keyOrigin[1] - arrangeSettings.keyObject.top);
        } else if (arrangeSettings.randomize) {
            shiftItems(orderedItems,
                blockOrigin[0] - orderedItems[0].left,
                blockOrigin[1] - orderedItems[0].top);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 行に∧∨と数値欄をひと組で追加する（隙間0で突き合わせ、↑↓キーも∧∨と同じ処理で増減する）
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - addStepper() に渡す増減の設定。onStep はダイアログ側であとから入れる
     * @returns {EditText} 入力欄（∧∨は .stepperGroup、増減の設定は .stepOptions で参照できる）
     */
    function addStepperInput(parentRow, initialText, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        numberInput.characters = FIELD_CHAR_WIDTH;
        numberInput.stepperGroup = stepperGroup;
        numberInput.stepOptions = stepOptions;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 入力欄と∧∨の有効／無効をまとめて切り替え、∧∨を描き直す
     * @param {EditText} numberInput - addStepperInput() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setStepperInputEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        redrawSteppersIn(numberInput.stepperGroup);
    }

    /**
     * ラジオボタンの一覧に同じツールチップを設定する
     * @param {RadioButton[]} radioList - 対象のラジオボタン
     * @param {string} helpTipText - 設定するツールチップ
     * @returns {void}
     */
    function setRadiosHelpTip(radioList, helpTipText) {
        for (var i = 0; i < radioList.length; i++) {
            radioList[i].helpTip = helpTipText;
        }
    }

    /**
     * ラジオボタンの一覧をまとめて有効・無効にする
     * @param {RadioButton[]} radioList - 対象のラジオボタン
     * @param {boolean} enabled - 有効にするかどうか
     * @returns {void}
     */
    function setRadiosEnabled(radioList, enabled) {
        for (var i = 0; i < radioList.length; i++) {
            radioList[i].enabled = enabled;
        }
    }

    /**
     * 揃えのラジオ（起点・中央・終点・なしの順）から、軸に依らない揃えの値を返す
     * @param {RadioButton[]} alignRadios - [起点, 中央, 終点, なし] の順に並んだラジオボタン
     * @returns {string} "start" / "center" / "end" / "none"
     */
    function getAlignValue(alignRadios) {
        if (alignRadios[1].value) return "center";
        if (alignRadios[2].value) return "end";
        if (alignRadios[3].value) return "none";
        return "start";
    }

    /**
     * 整列と分布ダイアログのパネルとコントロールを組み立てる（振る舞いの結線は呼び出し側で行う）
     * @param {Window} dialogWindow - 組み立て先のダイアログ
     * @param {PageItem} keyObject - キーオブジェクト（無ければ null）
     * @returns {object} 生成したコントロールをまとめたオブジェクト
     */
    function buildArrangeDialogControls(dialogWindow, keyObject) {
        /* 方向パネル（並べる方向と行数／列数）/ Direction panel: tiling direction and lane count */
        var directionPanel = addPanel(dialogWindow, getLabel("panel.direction"), 6);

        var directionRow = directionPanel.add("group");
        setupRow(directionRow, "left", 6);
        var horizontalRadio = directionRow.add("radiobutton", undefined, getLabel("radio.directionHorizontal"));
        horizontalRadio.helpTip = getLabel("tooltip.directionHorizontal");
        var verticalRadio = directionRow.add("radiobutton", undefined, getLabel("radio.directionVertical"));
        verticalRadio.helpTip = getLabel("tooltip.directionVertical");
        horizontalRadio.value = (DEFAULT_DIRECTION === "horizontal");
        verticalRadio.value = !horizontalRadio.value;

        /* 行数（横）／列数（縦）/ Row count (horizontal) or column count (vertical) */
        var laneCountRow = directionPanel.add("group");
        setupRow(laneCountRow, "left", 6);
        var laneCountLabel = addFieldLabel(laneCountRow, "");
        /* 1以上の整数 / integer of 1 or more */
        var laneCountInput = addStepperInput(laneCountRow, DEFAULT_LANE_COUNT, { step: 1, min: 1, integer: true });
        laneCountInput.helpTip = getLabel("tooltip.laneCount");

        var gridCheckbox = addOptionCheckbox(directionPanel, getLabel("checkbox.useGrid"), DEFAULT_USE_GRID);
        gridCheckbox.helpTip = getLabel("tooltip.useGrid");

        /* 間隔パネル（左＝横・縦の入力、右＝連動アイコンを上下中央）/ Spacing panel: fields on the left, the link icon centered on the right */
        var spacingPanel = addPanel(dialogWindow, getLabel("panel.spacing") + " (" + getUnitInfo().label + ")");
        var spacingRow = spacingPanel.add("group");
        setupRow(spacingRow, "left", COLUMN_SPACING);

        var marginColumn = spacingRow.add("group");
        marginColumn.orientation = "column";
        marginColumn.alignChildren = ["left", "center"];
        marginColumn.spacing = 6;

        var hMarginRow = marginColumn.add("group");
        setupRow(hMarginRow, "left", 6);
        addFieldLabel(hMarginRow, labelText("fieldLabel.hMargin"));
        /* 負のアキも可 / negative gaps allowed */
        var hMarginInput = addStepperInput(hMarginRow, DEFAULT_MARGIN, { step: 1 });
        hMarginInput.helpTip = getLabel("tooltip.hMargin");

        var vMarginRow = marginColumn.add("group");
        setupRow(vMarginRow, "left", 6);
        addFieldLabel(vMarginRow, labelText("fieldLabel.vMargin"));
        var vMarginInput = addStepperInput(vMarginRow, DEFAULT_MARGIN, { step: 1 });
        vMarginInput.helpTip = getLabel("tooltip.vMargin");

        /* 切り替え後の処理は結線時に onLinkToggle に入れる / The handler is set later when wiring the dialog */
        var linkToggle = addLinkToggle(spacingRow, DEFAULT_LINK_MARGINS, function () {
            if (linkToggle.onLinkToggle) linkToggle.onLinkToggle();
        });
        linkToggle.helpTip = getLabel("tooltip.linkMargins");
        /* 連動中は縦をディムして横の値に合わせる / While linked, dim V and mirror H */
        setStepperInputEnabled(vMarginInput, !linkToggle.value);
        if (linkToggle.value) {
            vMarginInput.text = hMarginInput.text;
        }

        /* 揃えパネル（上段＝上下、下段＝左右）/ Align panel: vertical row on top, horizontal row below */
        var alignmentPanel = addPanel(dialogWindow, getLabel("panel.alignment"), 6);

        var vAlignRow = alignmentPanel.add("group");
        setupRow(vAlignRow, "left", 6);
        var vAlignTopRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignTop"));
        var vAlignMiddleRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignMiddle"));
        var vAlignBottomRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignBottom"));
        var vAlignNoneRadio = vAlignRow.add("radiobutton", undefined, getLabel("radio.alignNone"));
        var vAlignRadios = [vAlignTopRadio, vAlignMiddleRadio, vAlignBottomRadio, vAlignNoneRadio];
        setRadiosHelpTip(vAlignRadios, getLabel("tooltip.alignVertical"));
        vAlignTopRadio.value = true;

        var hAlignRow = alignmentPanel.add("group");
        setupRow(hAlignRow, "left", 6);
        var hAlignLeftRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignLeft"));
        var hAlignCenterRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignCenter"));
        var hAlignRightRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignRight"));
        var hAlignNoneRadio = hAlignRow.add("radiobutton", undefined, getLabel("radio.alignNone"));
        var hAlignRadios = [hAlignLeftRadio, hAlignCenterRadio, hAlignRightRadio, hAlignNoneRadio];
        setRadiosHelpTip(hAlignRadios, getLabel("tooltip.alignHorizontal"));
        hAlignLeftRadio.value = true;

        /* オプションパネル / Options panel */
        var optionsPanel = addPanel(dialogWindow, getLabel("panel.options"), 6);

        /* キーオブジェクトが未検出のときはディム / Dimmed when no key object is detected */
        var keyObjectCheckbox = addOptionCheckbox(optionsPanel, getLabel("checkbox.useKeyObject"), !!keyObject);
        keyObjectCheckbox.helpTip = getLabel("tooltip.useKeyObject");
        keyObjectCheckbox.enabled = !!keyObject;
        var previewBoundsCheckbox = addOptionCheckbox(optionsPanel, getLabel("checkbox.usePreviewBounds"), DEFAULT_USE_PREVIEW_BOUNDS);
        previewBoundsCheckbox.helpTip = getLabel("tooltip.usePreviewBounds");
        var randomizeCheckbox = addOptionCheckbox(optionsPanel, getLabel("checkbox.randomize"), DEFAULT_RANDOMIZE);
        randomizeCheckbox.helpTip = getLabel("tooltip.randomize");

        /* ボタンエリア（左右中央）/ Button bar, centered */
        var buttonRow = addButtonRow(dialogWindow);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            horizontalRadio: horizontalRadio,
            verticalRadio: verticalRadio,
            laneCountLabel: laneCountLabel,
            laneCountInput: laneCountInput,
            gridCheckbox: gridCheckbox,
            hMarginInput: hMarginInput,
            vMarginInput: vMarginInput,
            linkToggle: linkToggle,
            vAlignMiddleRadio: vAlignMiddleRadio,
            vAlignRadios: vAlignRadios,
            hAlignCenterRadio: hAlignCenterRadio,
            hAlignRadios: hAlignRadios,
            keyObjectCheckbox: keyObjectCheckbox,
            previewBoundsCheckbox: previewBoundsCheckbox,
            randomizeCheckbox: randomizeCheckbox
        };
    }

    /**
     * ダイアログの入力内容を配置設定として読み取る
     * @param {object} dialogControls - buildArrangeDialogControls() の戻り値
     * @param {PageItem} keyObject - キーオブジェクト（無ければ null）
     * @param {number[]} keyOrigin - キーオブジェクトのプレビュー前の位置（無ければ null）
     * @returns {object} 配置設定
     */
    function readArrangeSettings(dialogControls, keyObject, keyOrigin) {
        var unitFactor = getUnitInfo().pointsPerUnit;

        var hMarginValue = parseFloat(dialogControls.hMarginInput.text);
        if (isNaN(hMarginValue)) hMarginValue = 0;
        var vMarginValue = parseFloat(dialogControls.vMarginInput.text);
        if (isNaN(vMarginValue)) vMarginValue = 0;
        var laneCount = parseInt(dialogControls.laneCountInput.text, 10);
        if (isNaN(laneCount) || laneCount < 1) laneCount = 1;

        return {
            direction: dialogControls.horizontalRadio.value ? "horizontal" : "vertical",
            laneCount: laneCount,
            hMarginPt: hMarginValue * unitFactor,
            vMarginPt: vMarginValue * unitFactor,
            /* 揃えは軸に依らない形（start / center / end / none）で持つ / Align values are axis-neutral */
            vAlign: getAlignValue(dialogControls.vAlignRadios),
            hAlign: getAlignValue(dialogControls.hAlignRadios),
            usePreviewBounds: dialogControls.previewBoundsCheckbox.value,
            useGrid: dialogControls.gridCheckbox.value,
            randomize: dialogControls.randomizeCheckbox.value,
            useKeyObject: dialogControls.keyObjectCheckbox.value,
            keyObject: keyObject,
            keyOrigin: keyOrigin
        };
    }

    /**
     * 配置ダイアログを表示し、プレビューしながら設定を決める
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {object} keyObject - キーオブジェクト（未検出のときは null）
     * @returns {object} 確定した配置設定。キャンセル時は null
     */
    function showArrangeDialog(targetItems, keyObject) {
        var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialogWindow);

        var previewManager = new PreviewManager();
        /* キャンセル時に戻せるよう、境界計算の環境設定を控える / Remember the bounds preference so Cancel can restore it */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");
        /* キーオブジェクトのプレビュー前の位置 / Key object position before any preview */
        var keyOrigin = keyObject ? [keyObject.left, keyObject.top] : null;

        var dialogControls = buildArrangeDialogControls(dialogWindow, keyObject);
        var horizontalRadio = dialogControls.horizontalRadio;
        var laneCountInput = dialogControls.laneCountInput;
        var gridCheckbox = dialogControls.gridCheckbox;
        var hMarginInput = dialogControls.hMarginInput;
        var vMarginInput = dialogControls.vMarginInput;
        var linkToggle = dialogControls.linkToggle;
        var previewBoundsCheckbox = dialogControls.previewBoundsCheckbox;

        /* 直近でプレビューへ反映した値（同じ値での二重更新を避ける）/ Value last pushed to the preview */
        var appliedLaneCountText = laneCountInput.text;
        /**
         * 行数・列数が変わったときだけプレビューを更新する
         * @returns {void}
         */
        function updatePreviewForLaneCount() {
            if (laneCountInput.text === appliedLaneCountText) return;
            appliedLaneCountText = laneCountInput.text;
            updatePreview();
        }
        /* 入力中は数値として読めるときだけ反映する（打っている途中で書き換えない）/ While typing, refresh only when the text parses */
        laneCountInput.onChanging = function() {
            var typedLaneCount = parseInt(laneCountInput.text, 10);
            if (isNaN(typedLaneCount) || typedLaneCount < 1) return;
            updatePreviewForLaneCount();
        };
        /* 確定時に1以上の整数へ丸める（表示と実際に使う値を一致させる）/ Snap to an integer of 1 or more on commit */
        laneCountInput.onChange = function() {
            var laneCountValue = parseInt(laneCountInput.text, 10);
            if (isNaN(laneCountValue) || laneCountValue < 1) laneCountValue = 1;
            laneCountInput.text = laneCountValue;
            updatePreviewForLaneCount();
        };
        /* ∧∨・↑↓キーで増減したあとの更新 / refresh after stepping with the stepper or arrow keys */
        laneCountInput.stepOptions.onStep = updatePreviewForLaneCount;
        hMarginInput.stepOptions.onStep = syncMarginsAndPreview;
        vMarginInput.stepOptions.onStep = updatePreview;

        /**
         * 方向とグリッドの状態に合わせてUIを整える（項目名と揃えの操作可否）
         * @returns {void}
         */
        function syncDirectionUI() {
            var isHorizontal = horizontalRadio.value;
            dialogControls.laneCountLabel.text = labelText(isHorizontal ? "fieldLabel.rowCount" : "fieldLabel.columnCount");
            /* 主軸（並べる方向）の揃えはグリッド時のみ有効 / Main-axis align is available in grid mode only */
            setRadiosEnabled(dialogControls.hAlignRadios, isHorizontal ? gridCheckbox.value : true);
            setRadiosEnabled(dialogControls.vAlignRadios, isHorizontal ? true : gridCheckbox.value);
        }
        syncDirectionUI();

        /**
         * Undo履歴を汚さずにプレビューを更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            /* 境界計算に使う環境設定は変わったときだけ書き換える（切り替えた直後は再描画しないと古い境界のまま計算される）/ Write the bounds preference only when it changes; without a redraw the old bounds are used */
            if (app.preferences.getBooleanPreference("includeStrokeInBounds") !== previewBoundsCheckbox.value) {
                app.preferences.setBooleanPreference("includeStrokeInBounds", previewBoundsCheckbox.value);
                app.redraw();
            }
            previewManager.addStep(function() {
                arrangeItems(targetItems, readArrangeSettings(dialogControls, keyObject, keyOrigin));
            });
        }

        /**
         * 方向・グリッドに合わせてUIを整えてからプレビューを更新する
         * @returns {void}
         */
        function syncDirectionAndPreview() {
            syncDirectionUI();
            updatePreview();
        }

        /**
         * 連動がONなら横の値を縦へ反映してからプレビューを更新する
         * @returns {void}
         */
        function syncMarginsAndPreview() {
            if (linkToggle.value) {
                vMarginInput.text = hMarginInput.text;
            }
            updatePreview();
        }

        hMarginInput.onChanging = syncMarginsAndPreview;
        hMarginInput.onChange = syncMarginsAndPreview;
        vMarginInput.onChanging = updatePreview;
        linkToggle.onLinkToggle = function() {
            setStepperInputEnabled(vMarginInput, !linkToggle.value);
            syncMarginsAndPreview();
        };
        horizontalRadio.onClick = syncDirectionAndPreview;
        dialogControls.verticalRadio.onClick = syncDirectionAndPreview;
        for (var i = 0; i < dialogControls.vAlignRadios.length; i++) {
            dialogControls.vAlignRadios[i].onClick = updatePreview;
        }
        for (var j = 0; j < dialogControls.hAlignRadios.length; j++) {
            dialogControls.hAlignRadios[j].onClick = updatePreview;
        }
        dialogControls.keyObjectCheckbox.onClick = updatePreview;
        previewBoundsCheckbox.onClick = updatePreview;
        dialogControls.randomizeCheckbox.onClick = updatePreview;
        gridCheckbox.onClick = function() {
            if (gridCheckbox.value) {
                /* グリッドは天地・左右とも中央を既定にする / Grid defaults to centered on both axes */
                dialogControls.vAlignMiddleRadio.value = true;
                dialogControls.hAlignCenterRadio.value = true;
            }
            syncDirectionAndPreview();
        };

        updatePreview();
        laneCountInput.active = true;

        prepareDialogWindow(dialogWindow, SCRIPT_NAME);
        if (dialogWindow.show() !== 1) {
            /* キャンセル：プレビューを巻き戻し、環境設定も元に戻す / Cancel: roll back the preview and the preference */
            previewManager.rollback();
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return null;
        }

        var arrangeSettings = readArrangeSettings(dialogControls, keyObject, keyOrigin);
        /* 1回のUndoで取り消せるように、巻き戻してから一度だけ実行する / Confirm as a single undoable action */
        previewManager.confirm(function() {
            arrangeItems(targetItems, arrangeSettings);
            /* 環境設定はスクリプトの外へ影響を残さないよう元に戻す / Restore the preference so the script leaves no global side effect */
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
        });
        return arrangeSettings;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを取得し、キーオブジェクトを判定してダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var targetItems = doc.selection;
        if (!targetItems || targetItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* キーオブジェクトを検出（整列コマンドで一時的に動かして元へ戻す）/ Detect the key object; it aligns temporarily, then restores */
        var keyObject = null;
        try {
            keyObject = detectKeyObject(targetItems);
        } catch (err) {
            /* 検出に失敗しても配置は続行する（チェックボックスがディムされるだけ）/ Keep going; only the checkbox is dimmed */
            $.writeln(SCRIPT_NAME + ": キーオブジェクトの検出に失敗 / key object detection failed — " + err);
        }

        showArrangeDialog(targetItems, keyObject);
    }

    try {
        main();
    } catch (e) {
        alert(labelText("alert.unexpectedError") + " " + e.message);
    }

})();
